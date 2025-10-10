' Brave Team / IPD No. 4480 / Inder 1985 / 4 Players
' VPX8 version 6.0.0 by jpsalas 2024
' Solenoids and light numbers taken from Destruk's script because of error numbers in the manual (again :) )

Option Explicit
Randomize

Const BallSize = 50
Const BallMass = 1

On Error Resume Next
ExecuteGlobal GetTextFile("controller.vbs")
If Err Then MsgBox "You need the controller.vbs in order to run this table, available in the vp10 package"
On Error Goto 0

Dim bsTrough, dtBankL, dtBankR, bsL, bsR, bcL, bcR, x

Const cGameName = "brvteam"

Dim VarHidden
If Table1.ShowDT = true then
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

LoadVPM "01550000", "inder.vbs", 3.26

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
        .SplashInfoLine = "Brave Team - Inder 1985" & vbNewLine & "VPX table by JPSalas 6.0.0"
        .HandleKeyboard = 0
        .ShowTitle = 0
        .ShowDMDOnly = 1
        .ShowFrame = 0
        .HandleMechanics = 0
        .Hidden = VarHidden
        .Games(cGameName).Settings.Value("rol") = 0   '1= rotated display, 0= normal
        .Games(cGameName).Settings.Value("sound") = 1 '1 enabled rom sound
        On Error Resume Next
        Controller.SolMask(0) = 0
        vpmTimer.AddTimer 2000, "Controller.SolMask(0)=&Hffffffff'" 'ignore all solenoids - then add the Timer to renable all the solenoids after 2 seconds
        Controller.Run GetPlayerHWnd
        On Error Goto 0
    End With

    ' Nudging
    vpmNudge.TiltSwitch = 53
    vpmNudge.Sensitivity = 3
    vpmNudge.TiltObj = Array(Bumper001, Bumper002, Bumper003)

    ' Trough
    Set bsTrough = New cvpmBallStack
    With bsTrough
        .InitSw 0, 91, 0, 0, 0, 0, 0, 0
        .InitKick BallRelease, 80, 6
        .InitEntrySnd "Solenoid", "Solenoid"
        .InitExitSnd SoundFX("fx_ballrel", DOFContactors), SoundFX("fx_Solenoid", DOFContactors)
        .Balls = 1
    End With

    ' Drop targets Left
    set dtBankL = new cvpmdroptarget
    With dtBankL
        .initdrop array(sw64, sw74, sw84, sw94), array(64, 74, 84, 94)
        .initsnd SoundFX("fx_droptarget", DOFDropTargets), SoundFX("fx_resetdrop", DOFContactors)
        .CreateEvents "dtBankL"
    End With

    ' Drop targets Right
    set dtBankR = new cvpmdroptarget
    With dtBankR
        .initdrop array(sw65, sw75, sw85, sw95), array(65, 75, 85, 95)
        .initsnd SoundFX("fx_droptarget", DOFDropTargets), SoundFX("fx_resetdrop", DOFContactors)
        .CreateEvents "dtBankR"
    End With

    ' Left Canon
    Set bsL = New cvpmBallStack
    With bsL
        .InitSaucer sw70, 70, 0, 21
        .InitExitSnd SoundFX("fx_kicker", DOFContactors), SoundFX("fx_Solenoid", DOFContactors)
        .KickForceVar = 3
    End With

    ' Right Canon
    Set bsR = New cvpmBallStack
    With bsR
        .InitSaucer sw80, 80, 0, 18
        .InitExitSnd SoundFX("fx_kicker", DOFContactors), SoundFX("fx_Solenoid", DOFContactors)
        .KickForceVar = 2
    End With

    ' Captive left ball
    Set bcL = New cvpmCaptiveBall
    With bcL
        .InitCaptive bcLtrigger, bcLwall, bcLkicker, 0
        .Start
        .ForceTrans = .99
        .MinForce = 3.8
        .CreateEvents "bcL"
    End With

    ' Captive right ball
    Set bcR = New cvpmCaptiveBall
    With bcR
        .InitCaptive bcRtrigger, bcRwall, bcRkicker, 0
        .Start
        .ForceTrans = .99
        .MinForce = 3.8
        .CreateEvents "bcR"
    End With

    ' Main Timer init
    PinMAMETimer.Interval = PinMAMEInterval
    PinMAMETimer.Enabled = 1
    RealTime.Enabled = 1

    ' Turn on Gi
    vpmtimer.addtimer 1500, "GiOn '"
End Sub

'******************
' RealTime Updates
'******************

Sub RealTime_Timer
    RollingUpdate
End Sub

'***********
' GI lights
'***********

Sub GiOn 'enciende las luces GI
    Dim bulb
    For each bulb in aGiLights
        bulb.State = 1
    Next
End Sub

Sub GiOff 'apaga las luces GI
    Dim bulb
    For each bulb in aGiLights
        bulb.State = 0
    Next
End Sub

'**********
' Keys
'**********

Sub table1_KeyDown(ByVal Keycode)
    If keycode = LeftTiltKey Then Nudge 90, 6:PlaySound SoundFX("fx_nudge", 0), 0, 1, -0.1, 0.25
    If keycode = RightTiltKey Then Nudge 270, 6:PlaySound SoundFX("fx_nudge", 0), 0, 1, 0.1, 0.25
    If keycode = CenterTiltKey Then Nudge 0, 8:PlaySound SoundFX("fx_nudge", 0), 0, 1, 0, 0.25
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
Sub rsband005_Hit:vpmTimer.PulseSw 63:End Sub
Sub rsband006_Hit:vpmTimer.PulseSw 63:End Sub
Sub rsband011_Hit:vpmTimer.PulseSw 71:End Sub
Sub rsband016_Hit:vpmTimer.PulseSw 71:End Sub
Sub rsband007_Hit:vpmTimer.PulseSw 71:End Sub
Sub rsband010_Hit:vpmTimer.PulseSw 71:End Sub

' Bumpers
Sub Bumper001_Hit:vpmTimer.PulseSw 96:PlaySoundAt SoundFX("fx_bumper", DOFContactors), Bumper001:End Sub
Sub Bumper002_Hit:vpmTimer.PulseSw 86:PlaySoundAt SoundFX("fx_bumper", DOFContactors), Bumper002:End Sub
Sub Bumper003_Hit:vpmTimer.PulseSw 76:PlaySoundAt SoundFX("fx_bumper", DOFContactors), Bumper003:End Sub

' Drain & holes
Sub Drain_Hit:PlaysoundAt "fx_drain", Drain:bsTrough.AddBall Me:End Sub
Sub sw70_Hit:PlaysoundAt "fx_kicker_enter", sw70:bsl.AddBall 0:End Sub
Sub sw80_Hit:PlaysoundAt "fx_kicker_enter", sw80:bsr.AddBall 0:End Sub

' Rollovers
Sub sw77_Hit:Controller.Switch(77) = 1:PlaySoundAt "fx_sensor", sw77:End Sub
Sub sw77_UnHit:Controller.Switch(77) = 0:End Sub

Sub sw62a_Hit:Controller.Switch(62) = 1:PlaySoundAt "fx_sensor", sw62a:End Sub
Sub sw62a_UnHit:Controller.Switch(62) = 0:End Sub

Sub sw62b_Hit:Controller.Switch(62) = 1:PlaySoundAt "fx_sensor", sw62b:End Sub
Sub sw62b_UnHit:Controller.Switch(62) = 0:End Sub

Sub sw62c_Hit:Controller.Switch(62) = 1:PlaySoundAt "fx_sensor", sw62c:End Sub
Sub sw62c_UnHit:Controller.Switch(62) = 0:End Sub

Sub sw62d_Hit:Controller.Switch(62) = 1:PlaySoundAt "fx_sensor", sw62d:End Sub
Sub sw62d_UnHit:Controller.Switch(62) = 0:End Sub

Sub sw67_Hit:Controller.Switch(67) = 1:PlaySoundAt "fx_sensor", sw67:End Sub
Sub sw67_UnHit:Controller.Switch(67) = 0:End Sub

Sub sw83b_Hit:Controller.Switch(83) = 1:PlaySoundAt "fx_sensor", sw83b:End Sub
Sub sw83b_UnHit:Controller.Switch(83) = 0:End Sub

Sub sw81_Hit:Controller.Switch(81) = 1:PlaySoundAt "fx_sensor", sw81:End Sub
Sub sw81_UnHit:Controller.Switch(81) = 0:End Sub

Sub sw82_Hit:Controller.Switch(82) = 1:PlaySoundAt "fx_sensor", sw82:End Sub
Sub sw82_UnHit:Controller.Switch(82) = 0:End Sub

Sub sw66_Hit:Controller.Switch(66) = 1:PlaySoundAt "fx_sensor", sw66:End Sub
Sub sw66_UnHit:Controller.Switch(66) = 0:End Sub

'Targets
Sub sw97_Hit:vpmTimer.PulseSw 97:PlaySoundAtBall SoundFX("fx_target", DOFTargets):End Sub

Sub sw87_Hit:vpmTimer.PulseSw 87:PlaySoundAtBall SoundFX("fx_target", DOFTargets):End Sub

Sub sw83_Hit:vpmTimer.PulseSw 83:PlaySoundAtBall SoundFX("fx_target", DOFTargets):End Sub

Sub sw83a_Hit:vpmTimer.PulseSw 83:PlaySoundAtBall SoundFX("fx_target", DOFTargets):End Sub

Sub sw72_Hit:vpmTimer.PulseSw 72:PlaySoundAtBall SoundFX("fx_target", DOFTargets):End Sub

Sub sw73_Hit:vpmTimer.PulseSw 73:PlaySoundAtBall SoundFX("fx_target", DOFTargets):End Sub

'*********
'Solenoids
'*********
SolCallback(2) = "bsL.SolOut"
SolCallback(3) = "bsR.SolOut"
SolCallback(4) = "bsTrough.SolOut"
SolCallback(5) = "vpmNudge.SolGameOn"

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

Sub LeftFlipper_Animate:LeftFlipperTop.RotZ = LeftFlipper.CurrentAngle: End Sub
Sub RightFlipper_Animate: RightFlipperTop.RotZ = RightFlipper.CurrentAngle: End Sub

' flippers hit Sound

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
            If chgLamp(ii, 0) = 37 AND chgLamp(ii, 1) = 0 Then dtBankL.DropSol_On                     'light 37
            If chgLamp(ii, 0) = 38 AND chgLamp(ii, 1) = 0 Then dtBankR.DropSol_On                     'light 38
            If chgLamp(ii, 0) = 40 AND chgLamp(ii, 1) = 1 Then PlaySound SoundFX("fx_Knocker", DOFKnocker) 'light 40
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
    Lamp 5, l5
    Lamp 6, l6
    Lamp 7, l7
    Lamp 8, l8
    Lamp 9, l9
    Lamp 10, l10
    Lampm 11, l11b
    Lamp 11, l11a
    Lamp 12, l12
    Lamp 13, l13    '1
    Lamp 14, l14    '2
    Lamp 15, l15    '3
    'Lamp 16, l16   'possibly a solenoid, bumpers?
    Lamp 17, l17
    Lamp 18, l18
    Lampm 19, l19b
    Lamp 19, l19a
    Lampm 20, l20b
    Lamp 20, l20a
    Lampm 21, l21b
    Lamp 21, l21a
    Lamp 24, l24
    Lamp 26, l26
    Lamp 27, l27
    Lamp 28, l28
    Lamp 29, l29
    Lamp 30, l30
    Lamp 32, l32    'game over
    'Lamp 33, l33   'press start
    'Lamp 34, l34   'turns on from time to time, no clue what it is.
    Lamp 35, l35    'Ball in Play
    Lamp 36, l36    'Match
    ' Lamp 39, l39  'Handicap
    Lamp 41, l41
    Lamp 42, l42
    Lamp 43, l43
    Lamp 44, l44
    Lamp 45, l45
    Lamp 46, l46
    Lamp 47, l47
    Lamp 48, l48
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
Dim Digits(27)
Digits(0) = Array(a00, a02, a05, a06, a04, a01, a03)
Digits(1) = Array(a10, a12, a15, a16, a14, a11, a13)
Digits(2) = Array(a20, a22, a25, a26, a24, a21, a23)
Digits(3) = Array(a30, a32, a35, a36, a34, a31, a33)
Digits(4) = Array(a40, a42, a45, a46, a44, a41, a43)
Digits(5) = Array(a50, a52, a55, a56, a54, a51, a53)
Digits(6) = Array(e00, e02, e05, e06, e04, e01, e03)
Digits(7) = Array(e10, e12, e15, e16, e14, e11, e13)
Digits(8) = Array(e20, e22, e25, e26, e24, e21, e23)
Digits(9) = Array(e30, e32, e35, e36, e34, e31, e33)
Digits(10) = Array(e40, e42, e45, e46, e44, e41, e43)
Digits(11) = Array(e50, e52, e55, e56, e54, e51, e53)
Digits(12) = Array(b00, b02, b05, b06, b04, b01, b03)
Digits(13) = Array(b10, b12, b15, b16, b14, b11, b13)
Digits(14) = Array(b20, b22, b25, b26, b24, b21, b23)
Digits(15) = Array(b30, b32, b35, b36, b34, b31, b33)
Digits(16) = Array(b40, b42, b45, b46, b44, b41, b43)
Digits(17) = Array(b50, b52, b55, b56, b54, b51, b53)
Digits(18) = Array(f00, f02, f05, f06, f04, f01, f03)
Digits(19) = Array(f10, f12, f15, f16, f14, f11, f13)
Digits(20) = Array(f20, f22, f25, f26, f24, f21, f23)
Digits(21) = Array(f30, f32, f35, f36, f34, f31, f33)
Digits(22) = Array(f40, f42, f45, f46, f44, f41, f43)
Digits(23) = Array(f50, f52, f55, f56, f54, f51, f53)
Digits(24) = Array(c00, c02, c05, c06, c04, c01, c03)
Digits(25) = Array(c10, c12, c15, c16, c14, c11, c13)
Digits(26) = Array(d00, d02, d05, d06, d04, d01, d03)
Digits(27) = Array(d10, d12, d15, d16, d14, d11, d13)

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
                ballpitch = Pitch(BOT(b)) + 50000 'increase the pitch on a ramp
                ballvol = Vol(BOT(b)) * 5
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
    Msg(0) = "Brave Team - Inder 1985" &Chr(10)&Chr(10)
    Msg(1) = ""
    Msg(2) = "Droptargets: "
    Msg(3) = "   Blues: Triple Bonus Left Saucer "
    Msg(4) = "   Yellows: Extra Ball Right Saucer"
    Msg(5) = "   Greens: Extra Ball Upper Lanes"
    Msg(6) = "   Reds: lights alt. 50.000 on top lane and side targets"
    Msg(7) = ""
    Msg(8) = ""
    Msg(9) = "Knocking down Green and Red Drop Targets"
    Msg(10) = "   50.000 points or Special alternative in Captive Targets"
    Msg(11) = ""
    Msg(12) = "Knocking down all Left Drop Targets:"
    Msg(13) = "  Extra Ball at Left Outlane"
    Msg(14) = ""
    Msg(15) = "Knocking down all Right Drop Targets:"
    Msg(16) = "  Extra Ball at Right Outlane"
    Msg(17) = ""
    Msg(18) = "Knocking down all Drop Targets:"
    Msg(19) = "   50.000 points or Special alternative in side Targets"
    Msg(20) = ""
    For X = 1 To 20
        Msg(0) = Msg(0) + Msg(X)&Chr(13)
    Next
    MsgBox Msg(0), , "         Instructions and Rule Card"
End Sub

'Inder Brave Team
'added by Inkochnito
Sub editDips
    Dim vpmDips:Set vpmDips = New cvpmDips
    With vpmDips
        .AddForm 700, 400, "Brave Team - DIP switches"
        .AddFrame 2, 0, 190, "Maximum credits", &H00040000, Array("10 credits", 0, "20 credits", &H00040000)                                                                                                                                                                 'dip SL2-n (22) this setting doesn't seem to work. Always on.
        .AddFrame 2, 46, 190, "Handicap value", &H00030000, Array("850,000 points", 0, "900,000 points", &H00010000, "950,000 points", &H00020000, "990,000 points", &H00030000)                                                                                             'dip SL2-p&o(17&18)
        .AddFrame 2, 121, 392, "Credits per coin", &H06000000, Array("1 coin - 1 credit && 1 coin - 4 credits", 0, "1 coin - 1 credit (4 coins - 5 credits) && 1 coin - 5 credits", &H01000000, "1 coin - 1 credit (4 coins - 6 credits) && 1 coin - 6 credits", &H06000000) 'dip SL2-h&g (25&26)
        .AddFrame 205, 0, 190, "Balls per game", &H08000000, Array("3 balls", 0, "5 balls", &H08000000)                                                                                                                                                                      'dip SL1-e (28)
        .AddFrame 205, 46, 190, "Replay threshold", &H30000000, Array("700,000 points", 0, "750,000 points", &H10000000, "800,000 points", &H20000000, "850,000 points", &H30000000)                                                                                         'dip SL1-d&c (29&30)
        .AddLabel 50, 200, 300, 20, "After hitting OK, press F3 to reset game with new settings."
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

' Thalamus : Exit in a clean and proper way
Sub Table1_exit
  Controller.Pause = False
  Controller.Stop
End Sub


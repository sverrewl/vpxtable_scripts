'Disco Dancing / IPD No. 5892 / 2 Players
'LTD do Brasil Diversões Eletrônicas Ltda, of Campinas, São Paulo, Brazil (1977-1984)
'table by jpsalas(VPX), mfuegeman(VPinMAME), Halen(graphics)

Option Explicit
Randomize

Const BallSize = 50
Const BallMass = 1

On Error Resume Next
ExecuteGlobal GetTextFile("controller.vbs")
If Err Then MsgBox "You need the controller.vbs in order to run this table, available in the vp10 package"
On Error Goto 0

Dim bsTrough, dtBank1, dtBank2, FlipperActive, x

Const cGameName = "discodan"

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

LoadVPM "01560000", "ltd3.vbs", 3.2

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
        .SplashInfoLine = "Disco Dancing - LTD do Brasil 1979" & vbNewLine & "VPX table by JPSalas, mfuegemann and Halen 6.0.0"
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
    vpmNudge.TiltSwitch = 1
    vpmNudge.Sensitivity = 3
    vpmNudge.TiltObj = Array(Bumper1, Bumper2, LeftSlingshot, RightSlingshot)

    ' Trough
    Set bsTrough = New cvpmBallStack
    With bsTrough
        .InitSw 0, 57, 0, 0, 0, 0, 0, 0
        .InitKick BallRelease, 90, 4
        .InitExitSnd SoundFX("fx_ballrel", DOFContactors), SoundFX("fx_Solenoid", DOFContactors)
        .Balls = 1
    End With

    ' Drop targets
    Set dtBank1 = New cvpmDropTarget
    With dtBank1
        .InitDrop Array(DT1, DT2, DT3, DT4, DT5), Array(33, 34, 35, 36, 37)
        .InitSnd SoundFX("fx_droptarget", DOFDropTargets), SoundFX("fx_resetdrop", DOFContactors)
        .CreateEvents "dtBank1"
    End With

    Set dtBank2 = New cvpmDropTarget
    With dtBank2
        .InitDrop Array(DT6, DT7, DT8, DT9, DT10), Array(41, 42, 43, 44, 45)
        .InitSnd SoundFX("fx_droptarget", DOFDropTargets), SoundFX("fx_resetdrop", DOFContactors)
        .CreateEvents "dtBank2"
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

'**********
' Keys
'**********

Sub table1_KeyDown(ByVal Keycode)
    If keycode = LeftFlipperKey AND FlipperActive Then SolLFlipper 1
    If keycode = RightFlipperKey AND FlipperActive Then SolRFlipper 1
    If keycode = LeftTiltKey Then Nudge 90, 6:PlaySound SoundFX("fx_nudge", 0), 0, 1, -0.1, 0.25
    If keycode = RightTiltKey Then Nudge 270, 6:PlaySound SoundFX("fx_nudge", 0), 0, 1, 0.1, 0.25
    If keycode = CenterTiltKey Then Nudge 0, 8:PlaySound SoundFX("fx_nudge", 0), 0, 1, 0, 0.25
    If vpmKeyDown(keycode)Then Exit Sub
    If keycode = PlungerKey Then PlaySoundat "fx_PlungerPull", Plunger:Plunger.Pullback
End Sub

Sub table1_KeyUp(ByVal Keycode)
    If keycode = LeftFlipperKey Then SolLFlipper 0
    If keycode = RightFlipperKey Then SolRFlipper 0
    If vpmKeyUp(keycode)Then Exit Sub
    If keycode = PlungerKey Then PlaySoundAt "fx_plunger", Plunger:Plunger.Fire
End Sub

'*************************
' GI - needs new vpinmame
'*************************

'Set GICallback = GetRef("GIUpdate")

Sub GIUpdate(no, Enabled)
    If Enabled Then
        GiOn
    Else
        GiOff
    End If
End Sub

Sub GiOn
    Dim bulb
    PlaySound "fx_gion"
    For each bulb in aGiLights
        bulb.State = 1
    Next
End Sub

Sub GiOff
    Dim bulb
    PlaySound "fx_gioff"
    For each bulb in aGiLights
        bulb.State = 0
    Next
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
    vpmTimer.PulseSw 18
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
    vpmTimer.PulseSw 17
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
Sub rlband008_Hit:vpmTimer.PulseSw 3:End Sub
Sub rlband009_Hit:vpmTimer.PulseSw 3:End Sub
Sub rlband007_Hit:vpmTimer.PulseSw 3:End Sub
Sub rlband005_Hit:vpmTimer.PulseSw 3:End Sub
Sub rsband003_Hit:vpmTimer.PulseSw 3:End Sub
Sub rsband002_Hit:vpmTimer.PulseSw 3:End Sub

' Bumpers
Sub Bumper1_Hit:vpmTimer.PulseSw 9:PlaySoundAt SoundFX("fx_bumper", DOFContactors), Bumper1:End Sub
Sub Bumper2_Hit:vpmTimer.PulseSw 10:PlaySoundAt SoundFX("fx_bumper", DOFContactors), Bumper2:End Sub
Sub Bumper3_Hit:vpmTimer.PulseSw 11:PlaySoundAt SoundFX("fx_bumper", DOFContactors), Bumper3:End Sub

' Drain & holes
Sub Drain_Hit:PlaysoundAt "fx_drain", Drain:bsTrough.AddBall Me:End Sub

' Rollovers
Sub TopLeftLane_Hit:Controller.Switch(5) = 1:PlaySoundAt "fx_sensor", TopLeftLane:End Sub
Sub TopLeftLane_UnHit:Controller.Switch(5) = 0:End Sub

Sub TopRightLane_Hit:Controller.Switch(6) = 1:PlaySoundAt "fx_sensor", TopRightLane:End Sub
Sub TopRightLane_UnHit:Controller.Switch(6) = 0:End Sub

Sub Trigger_Star1_Hit:Controller.Switch(4) = 1:PlaySoundAt "fx_sensor", Trigger_Star1:End Sub
Sub Trigger_Star1_UnHit:Controller.Switch(4) = 0:End Sub

Sub Trigger_Star2_Hit:Controller.Switch(4) = 1:PlaySoundAt "fx_sensor", Trigger_Star2:End Sub
Sub Trigger_Star2_UnHit:Controller.Switch(4) = 0:End Sub

Sub Trigger_Star3_Hit:Controller.Switch(4) = 1:PlaySoundAt "fx_sensor", Trigger_Star3:End Sub
Sub Trigger_Star3_UnHit:Controller.Switch(4) = 0:End Sub

Sub LeftOutlane_Hit:Controller.Switch(4) = 1:PlaySoundAt "fx_sensor", LeftOutlane:End Sub
Sub LeftOutlane_UnHit:Controller.Switch(4) = 0:End Sub

Sub RightOutlane_Hit:Controller.Switch(4) = 1:PlaySoundAt "fx_sensor", RightOutlane:End Sub
Sub RightOutlane_UnHit:Controller.Switch(4) = 0:End Sub

'Targets
Sub Target1_Hit:vpmTimer.PulseSw 13:PlaySoundAtBall SoundFX("fx_target", DOFTargets):End Sub
Sub Target2_Hit:vpmTimer.PulseSw 14:PlaySoundAtBall SoundFX("fx_target", DOFTargets):End Sub

'*********
'Solenoids
'*********

SolCallback(2) = "bsTrough.SolOut"
SolCallback(1) = "dtBank1.SolDropUp"
SolCallback(6) = "dtBank2.SolDropUp"
SolCallback(17) = "Sol_GameOn"
'SolCallback(3) = "vpmSolSound SoundFX(""fx_Knocker"",DOFKnocker),"

Sub Sol_GameOn(enabled)
    vpmNudge.SolGameOn enabled
    Flipperactive = Enabled
End Sub

'*******************
' Flipper Subs Rev3
'*******************

'SolCallback(sLRFlipper) = "SolRFlipper"
'SolCallback(sLLFlipper) = "SolLFlipper"

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
            LampState(chgLamp(ii, 0)) = chgLamp(ii, 1) 'keep the real state in an array
            FadingStep(chgLamp(ii, 0)) = chgLamp(ii, 1)
        Next
    End If
    UpdateLeds
    UpdateLamps
End Sub

Sub UpdateLamps
    ' playfield lights
    NfReel 5, l005 'Ball in Play
    NfReel 8, l008 'Game over
    Lamp 17, l17
    Lamp 18, l18
    Lamp 19, l19
    Lamp 20, l20
    Lamp 22, l22
    Lamp 23, l23
    Lamp 26, l26
    Lamp 27, l27
    Lamp 28, l28
    Lamp 29, l29
    Lamp 30, l30
    Lamp 31, l31
    Lamp 32, l32
    Lamp 33, l33
    Lamp 39, l39
    Lamp 40, l40
    Lamp 41, l41
    Lamp 45, l45
    Lamp 46, l46
    Lamp 47, l47
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

Sub Gi(nr, object)
    Select Case FadingStep(nr)
        Case 1:GiOn:FadingStep(nr) = -1
        Case 0:GiOff:FadingStep(nr) = -1
    End Select
End Sub

'************************************
'          LEDs Display
'     Based on Scapino's LEDs
'************************************

Dim Digits(32)

'Assign 5-digit output to reels
Set Digits(0) = a0
Set Digits(1) = a1
Set Digits(2) = a2
Set Digits(3) = a3
Set Digits(4) = a4
Set Digits(5) = b0
Set Digits(6) = b1
Set Digits(7) = b2
Set Digits(8) = b3
Set Digits(9) = b4
Set Digits(10) = e0
Set Digits(11) = e1
Set Digits(12) = f0

Sub UPdateLEDs
    On Error Resume Next
    Dim ChgLED, ii, jj, chg, stat
    ChgLED = Controller.ChangedLEDs(&H0000003f, &Hffffffff)
    If Not IsEmpty(ChgLED)Then
        For ii = 0 To UBound(ChgLED)
            chg = chgLED(ii, 1):stat = chgLED(ii, 2)
            If chgLED(ii, 0) < 12 Then
                Select Case stat
                    Case 0:Digits(chgLED(ii, 0)).SetValue 0    'empty
                    Case 63:Digits(chgLED(ii, 0)).SetValue 1   '0
                    Case 6:Digits(chgLED(ii, 0)).SetValue 2    '1
                    Case 91:Digits(chgLED(ii, 0)).SetValue 3   '2
                    Case 79:Digits(chgLED(ii, 0)).SetValue 4   '3
                    Case 102:Digits(chgLED(ii, 0)).SetValue 5  '4
                    Case 109:Digits(chgLED(ii, 0)).SetValue 6  '5
                    Case 124:Digits(chgLED(ii, 0)).SetValue 7  '6
                    Case 125:Digits(chgLED(ii, 0)).SetValue 7  '6
                    Case 252:Digits(chgLED(ii, 0)).SetValue 7  '6
                    Case 7:Digits(chgLED(ii, 0)).SetValue 8    '7
                    Case 127:Digits(chgLED(ii, 0)).SetValue 9  '8
                    Case 103:Digits(chgLED(ii, 0)).SetValue 10 '9
                    Case 111:Digits(chgLED(ii, 0)).SetValue 10 '9
                    Case 231:Digits(chgLED(ii, 0)).SetValue 10 '9
                    Case 128:Digits(chgLED(ii, 0)).SetValue 0  'empty
                    Case 191:Digits(chgLED(ii, 0)).SetValue 1  '0
                    Case 832:Digits(chgLED(ii, 0)).SetValue 2  '1
                    Case 896:Digits(chgLED(ii, 0)).SetValue 2  '1
                    Case 768:Digits(chgLED(ii, 0)).SetValue 2  '1
                    Case 134:Digits(chgLED(ii, 0)).SetValue 2  '1
                    Case 219:Digits(chgLED(ii, 0)).SetValue 3  '2
                    Case 207:Digits(chgLED(ii, 0)).SetValue 4  '3
                    Case 230:Digits(chgLED(ii, 0)).SetValue 5  '4
                    Case 237:Digits(chgLED(ii, 0)).SetValue 6  '5
                    Case 253:Digits(chgLED(ii, 0)).SetValue 7  '6
                    Case 135:Digits(chgLED(ii, 0)).SetValue 8  '7
                    Case 255:Digits(chgLED(ii, 0)).SetValue 9  '8
                    Case 239:Digits(chgLED(ii, 0)).SetValue 10 '9
                End Select
            Else
                Select Case stat
                    Case 0:Digits(chgLED(ii, 0)).ImageA = "d_off"   'empty
                    Case 63:Digits(chgLED(ii, 0)).ImageA = "d_0"    '0
                    Case 6:Digits(chgLED(ii, 0)).ImageA = "d_1"     '1
                    Case 91:Digits(chgLED(ii, 0)).ImageA = "d_2"    '2
                    Case 79:Digits(chgLED(ii, 0)).ImageA = "d_3"    '3
                    Case 102:Digits(chgLED(ii, 0)).ImageA = "d_4"   '4
                    Case 109:Digits(chgLED(ii, 0)).ImageA = "d_5"   '5
                    Case 124:Digits(chgLED(ii, 0)).ImageA = "d_6"   '6
                    Case 125:Digits(chgLED(ii, 0)).ImageA = "d_6"   '6
                    Case 252:Digits(chgLED(ii, 0)).ImageA = "d_6"   '6
                    Case 7:Digits(chgLED(ii, 0)).ImageA = "d_7"     '7
                    Case 127:Digits(chgLED(ii, 0)).ImageA = "d_8"   '8
                    Case 103:Digits(chgLED(ii, 0)).ImageA = "d_9"   '9
                    Case 111:Digits(chgLED(ii, 0)).ImageA = "d_9"   '9
                    Case 231:Digits(chgLED(ii, 0)).ImageA = "d_9"   '9
                    Case 128:Digits(chgLED(ii, 0)).ImageA = "d_off" 'empty
                    Case 191:Digits(chgLED(ii, 0)).ImageA = "d_0"   '0
                    Case 832:Digits(chgLED(ii, 0)).ImageA = "d_1"   '1
                    Case 896:Digits(chgLED(ii, 0)).ImageA = "d_1"   '1
                    Case 768:Digits(chgLED(ii, 0)).ImageA = "d_1"   '1
                    Case 134:Digits(chgLED(ii, 0)).ImageA = "d_1"   '1
                    Case 219:Digits(chgLED(ii, 0)).ImageA = "d_2"   '2
                    Case 207:Digits(chgLED(ii, 0)).ImageA = "d_3"   '3
                    Case 230:Digits(chgLED(ii, 0)).ImageA = "d_4"   '4
                    Case 237:Digits(chgLED(ii, 0)).ImageA = "d_5"   '5
                    Case 253:Digits(chgLED(ii, 0)).ImageA = "d_6"   '6
                    Case 135:Digits(chgLED(ii, 0)).ImageA = "d_7"   '7
                    Case 255:Digits(chgLED(ii, 0)).ImageA = "d_8"   '8
                    Case 239:Digits(chgLED(ii, 0)).ImageA = "d_9"   '9
                End Select
            End If
        Next
    End If
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
Sub aTargets_Hit(idx):PlaySound SoundFX("fx_target", DOFTargets), 0, Vol(ActiveBall), pan(ActiveBall), 0.2, Pitch(ActiveBall) * 10, 0, 0, AudioFade(ActiveBall):End Sub
Sub aRollovers_Hit(idx):PlaySoundAt "fx_sensor", aRollovers(idx):End Sub

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
Const maxvel = 36 'max ball velocity
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
        aBallShadow(b).Height = BOT(b).Z - Ballsize / 2 + 1

        If BallVel(BOT(b)) > 1 Then
            If BOT(b).z < 30 Then
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
        If BOT(b).VelZ < -1 and BOT(b).z < 55 and BOT(b).z > 27 Then 'height adjust for ball drop sounds
            PlaySound "fx_balldrop", 0, ABS(BOT(b).velz) / 17, Pan(BOT(b)), 0, Pitch(BOT(b)), 1, 0, AudioFade(BOT(b))
        End If

        ' jps ball speed & spin control
            BOT(b).AngMomZ = BOT(b).AngMomZ * 0.95
        If BOT(b).VelX AND BOT(b).VelY <> 0 Then
            speedfactorx = ABS(maxvel / BOT(b).VelX)
            speedfactory = ABS(maxvel / BOT(b).VelY)
            If speedfactorx < 1 Then
                BOT(b).VelX = BOT(b).VelX * speedfactorx
                BOT(b).VelY = BOT(b).VelY * speedfactorx
            End If
            If speedfactory < 1 Then
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

'***********************************
'Lamp Simulator from my table TOTAN
'***********************************

Dim lampPosition, lampSpinSpeed, lampLastPos
Const cLampSpeedMult = 180             ' 180 - Affects speed transfer to object (deg/sec)
Const cLampFriction = 1.5              ' 2.0 - Friction coefficient (deg/sec/sec)
Const cLampMinSpeed = 16               ' 20 - Object stops at this speed (deg/sec)
Const cLampRadius = 74
Const cBallSpeedDampeningEffect = 0.45 ' 45 - The ball retains this fraction of its speed due to energy absorption by hitting the lamp.
Const pi = 3.14159265358979323846

For Each x In colLampPoles:x.IsDropped = 1:Next
lampSpinSpeed = 0:lampLastPos = -1:SpinTimer.Enabled = True ' force update

' Draw lamp
Sub SpinTimer_Timer
    Dim curPos
    Dim oldLampSpeed:oldLampSpeed = lampSpinSpeed
    lampPosition = lampPosition + lampSpinSpeed * Me.Interval / 1000
    lampSpinSpeed = lampSpinSpeed * (1 - cLampFriction * Me.Interval / 1000)
    'added by jim - makes this sub a bit safer (even though this shouldnt happen in normal use)
    Do While lampPosition < 0
        lampPosition = lampPosition + 360
    Loop
    Do While lampPosition > 360
        lampPosition = lampPosition - 360
    Loop
    curPos = Int((lampPosition * colLampPoles.Count) / 360)
    If curPos <> lampLastPos Then
        Spinner1.RotZ = 360 - lampPosition
        Spinner2.RotZ = 360 - lampPosition
        Spinner3.RotZ = 360 - lampPosition
        Spinner4.RotZ = 360 - lampPosition
        If lampLastPos >= 0 Then ' not first time
            colLampPoles(lampLastPos).IsDropped = True
            colLampPoles2(lampLastPos).IsDropped = True
        End If
        On Error Resume Next
        colLampPoles(curPos).IsDropped = False
        If Err Then msgbox curPoles
        colLampPoles2(curPos).IsDropped = False
        If oldLampSpeed > 0 And lampLastPos > curPos Then
            PlaySoundAt SoundFX("fx_Spinner", DOFGear), LampPr
            'rev anticlockwise
            vpmTimer.PulseSw 12
        ElseIf oldLampSpeed < 0 And lampLastPos < curPos Then
            'rev clockwise
            PlaySoundAt SoundFX("fx_Spinner", DOFGear), LampPr
            vpmTimer.PulseSw 12
        End If
        lampLastPos = curPos
    End If
    If Abs(lampSpinSpeed) < cLampMinSpeed Then
        lampSpinSpeed = 0:Me.Enabled = False
    End If
End Sub

Sub colLampPoles_Hit(idx)
    PlaySoundAtBall "fx_rubber_pin"

    dim lampangle
    lampangle = NormAngle((idx) / Me.Count * 2 * pi + (pi / 17.6) + pi) ' added (pi / 17.6) because the first lamp post (18) is not quite at 0 angle, added another pi to make it standard angle

    dim mball, mlamp, rlamp, ilamp

    dim collisionangle
    dim ballspeedin, ballspeedout, lampspeedin, lampspeedout

    dim fudge:fudge = cLampSpeedMult / 2

    With ActiveBall
        collisionangle = GetCollisionAngle(idx, Me.Count, .X, -.Y) ' this is the angle from the center of ball to the center of post

        Set ballspeedout = new jVector
        ballspeedout.SetXY .VelX, -.VelY
        ballspeedout.ShiftAxes - collisionangle

        lampSpinSpeed = lampSpinSpeed + sqr(.VelX ^2 + .VelY ^2) * sin(collisionangle - lampangle) * fudge
        ballspeedout.SetXY ballspeedout.x, ballspeedout.y * cos(collisionangle - lampangle)

        ballspeedout.ShiftAxes collisionangle
        ' we can give a more accurate ball return speed or let the normal VP physics give the ball speed
        .VelX = ballspeedout.x * cBallSpeedDampeningEffect
        .VelY = - ballspeedout.y * cBallSpeedDampeningEffect
    End With
    SpinTimer.Enabled = True
End Sub

Function GetCollisionAngle(idx, count, X, Y)
    dim angle, postx, posty, dX, dY
    Dim ang
    angle = (idx) / count * 2 * pi + (pi / 17.6) ' added (pi / 17.6) because the first lamp post (18) is not quite at 0 angle
    postx = 443 - 75 * Cos(angle)                ' the actual coordinates of the center of the lamp
    posty = 918 + 75 * Sin(angle)                ' 60.25 is the radius of the circle with center at the center of the lamp and edge at the centers of all the lamp posts
    posty = -1 * posty
    Dim collisionV:Set collisionV = new jVector
    collisionV.SetXY postx - X, posty - Y
    GetCollisionAngle = collisionV.ang
End Function

Function NormAngle(angle)
    NormAngle = angle
    Do While NormAngle > 2 * pi
        NormAngle = NormAngle - 2 * pi
    Loop
    Do While NormAngle < 0
        NormAngle = NormAngle + 2 * pi
    Loop
End Function

Class jVector
    Private m_mag, m_ang, pi

    Sub Class_Initialize
        m_mag = CDbl(0)
        m_ang = CDbl(0)
        pi = CDbl(3.14159265358979323846)
    End Sub

    Public Function add(anothervector)
        Dim tx, ty, theta
        If TypeName(anothervector) = "jVector" then
            Set add = new jVector
            add.SetXY x + anothervector.x, y + anothervector.y
        End If
    End Function

    Public Function multiply(scalar)
        Set multiply = new jVector
        multiply.SetXY x * scalar, y * scalar
    End Function

    Sub ShiftAxes(theta)
        ang = ang - theta
    end Sub

    Sub SetXY(tx, ty)

        if tx = 0 And ty = 0 Then
            ang = 0
        elseif tx = 0 And ty < 0 then
            ang = - pi / 180 ' -90 degrees
        elseif tx = 0 And ty > 0 then
            ang = pi / 180   ' 90 degrees
        else
            ang = atn(ty / tx)
            if tx < 0 then ang = ang + pi ' Add 180 deg if in quadrant 2 or 3
        End if

        mag = sqr(tx ^2 + ty ^2)
    End Sub

    Property Let mag(nmag)
        m_mag = nmag
    End Property

    Property Get mag
        mag = m_mag
    End Property

    Property Let ang(nang)
        m_ang = nang
        Do While m_ang > 2 * pi
            m_ang = m_ang - 2 * pi
        Loop
        Do While m_ang < 0
            m_ang = m_ang + 2 * pi
        Loop
    End Property

    Property Get ang
        Do While m_ang > 2 * pi
            m_ang = m_ang - 2 * pi
        Loop
        Do While m_ang < 0
            m_ang = m_ang + 2 * pi
        Loop
        ang = m_ang
    End Property

    Property Get x
        x = m_mag * cos(ang)
    End Property

    Property Get y
        y = m_mag * sin(ang)
    End Property

    Property Get dump
        dump = "vector "
        Select Case CInt(ang + pi / 8)
            case 0, 8:dump = dump & "->"
            case 1:dump = dump & "/'"
            case 2:dump = dump & "/\"
            case 3:dump = dump & "'\"
            case 4:dump = dump & "<-"
            case 5:dump = dump & ":/"
            case 6:dump = dump & "\/"
            case 7:dump = dump & "\:"
        End Select
        dump = dump & " mag:" & CLng(mag * 10) / 10 & ", ang:" & CLng(ang * 180 / pi) & ", x:" & CLng(x * 10) / 10 & ", y:" & CLng(y * 10) / 10
    End Property
End Class

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


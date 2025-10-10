' Faeton - Juegos Populares (Spain) 1985 / IPD No. 3087 / 4 Players
' VPX8 version by JPSalas, 2024, version 6.0.0
' known issues:
' the hole doesn't always award the 30.000 points as indicated by its light

Option Explicit
Randomize

Const BallSize = 50
Const BallMass = 1

On Error Resume Next
ExecuteGlobal GetTextFile("controller.vbs")
If Err Then MsgBox "You need the controller.vbs in order to run this table, available in the vp10 package"
On Error Goto 0

Dim bsTrough, bsSaucer, cbCaptive, dtL, dtR, x

Const cGameName = "faeton"

Const UseSolenoids = 2 'Fastflips enabled
Const UseLamps = 0
Const UseGI = 0
Const UseSync = 0 'set it to 1 if the table runs too fast
Const HandleMech = 0

Dim VarHidden, UseVPMDMD
If Table1.ShowDT = true then
    UseVPMDMD = True
    VarHidden = 1
    For each x in aReels
        x.Visible = 1
    Next
else
    UseVPMDMD = False
    VarHidden = 0
    For each x in aReels
        x.Visible = 0
    Next
end if

if B2SOn = true then VarHidden = 1

LoadVPM "01550000", "juegos.vbs", 3.26

' Standard Sounds
Const SSolenoidOn = "fx_solenoidon"
Const SSolenoidOff = "fx_solenoidoff"
Const SCoin = "fx_Coin"

'************
' Table init.
'************

Sub table1_Init
    vpmInit me
    With Controller
        .GameName = cGameName
        If Err Then MsgBox "Can't start Game" & cGameName & vbNewLine & Err.Description:Exit Sub
        .SplashInfoLine = "Based on Faeton by Juegos Populares from 1985" & vbNewLine & "VPX table by jpsalas v6.0.0"
        .HandleKeyboard = 0
        .ShowTitle = 0
        .ShowDMDOnly = 1
        .ShowFrame = 0
        .HandleMechanics = 0
        .Hidden = VarHidden
        .Games(cGameName).Settings.Value("rol") = 0 '1= rotated display, 0= normal
        .Games(cGameName).Settings.Value("sound") = 1 '1 enabled rom sound
        '.SetDisplayPosition 0,0, GetPlayerHWnd 'restore dmd window position
        On Error Resume Next
        Controller.SolMask(0) = 0
        vpmTimer.AddTimer 1000, "Controller.SolMask(0)=&Hffffffff'" 'ignore all solenoids - then add the Timer to renable all the solenoids after 2 seconds
        Controller.Run GetPlayerHWnd
        On Error Goto 0
    End With

    ' Nudging
    vpmNudge.TiltSwitch = 30
    vpmNudge.Sensitivity = 3
    vpmNudge.TiltObj = Array(Bumper1, Bumper2, RightSlingshot)

    ' Trough
    Set bsTrough = New cvpmBallStack
    With bsTrough
        .InitNoTrough BallRelease, 25, 90, 7
        .InitExitSnd SoundFX("fx_ballrel", DOFContactors), SoundFX("fx_Solenoid", DOFContactors)
    End With

    ' Saucer
    Set bsSaucer = New cvpmBallStack
    bsSaucer.InitSaucer sw9, 9, 60, 20
    bsSaucer.InitExitSnd SoundFX("fx_popper", DOFContactors), SoundFX("fx_Solenoid", DOFContactors)
    bsSaucer.KickForceVar = 3
    bsSaucer.KickAngleVar = 3

    ' Captive Ball
    Set cbCaptive = New cvpmCaptiveBall
    cbCaptive.InitCaptive CaptiveTrigger, CaptiveWall, CaptiveKicker, 40
    cbCaptive.Start
    cbCaptive.ForceTrans = 1.9
    cbCaptive.MinForce = 3.5
    cbCaptive.CreateEvents "cbCaptive"

    ' Left Drop targets
    Set dtL = New cvpmDropTarget
    dtL.InitDrop Array(sw4, sw12, sw20), Array(4, 12, 20)
    dtL.initsnd SoundFX("fx_droptarget", DOFDropTargets), SoundFX("fx_resetdrop", DOFContactors)
    dtL.CreateEvents "dtL"

    ' Right Drop targets
    Set dtR = New cvpmDropTarget
    dtR.InitDrop Array(sw5, sw13, sw21), Array(5, 13, 21)
    dtR.initsnd SoundFX("fx_droptarget", DOFDropTargets), SoundFX("fx_resetdrop", DOFContactors)
    dtR.CreateEvents "dtR"

    ' Main Timer init
    PinMAMETimer.Interval = PinMAMEInterval
    PinMAMETimer.Enabled = 1
    Realtime.Enabled = 1

    ' Init walls
    sw7.IsDropped = 1:sw7a.IsDropped = 0:sw7c.IsDropped = 1
End Sub

Sub table1_Paused:Controller.Pause = 1:End Sub
Sub table1_unPaused:Controller.Pause = 0:End Sub
Sub table1_exit:Controller.stop:End Sub

'**********
' Keys
'**********

Sub table1_KeyDown(ByVal Keycode)
    If KeyCode = LeftFlipperKey Then Controller.Switch(84) = 1
    If KeyCode = RightFlipperKey Then Controller.Switch(82) = 1
    If keycode = LeftTiltKey Then Nudge 90, 5:PlaySound SoundFX("fx_nudge", 0), 0, 1, -0.1, 0.25
    If keycode = RightTiltKey Then Nudge 270, 5:PlaySound SoundFX("fx_nudge", 0), 0, 1, 0.1, 0.25
    If keycode = CenterTiltKey Then Nudge 0, 6:PlaySound SoundFX("fx_nudge", 0), 0, 1, 0, 0.25
    If vpmKeyDown(keycode)Then Exit Sub
    If keycode = PlungerKey Then PlaySound "fx_PlungerPull", 0, 1, 0.1, 0.25:Plunger.Pullback
End Sub

Sub table1_KeyUp(ByVal Keycode)
    If KeyCode = LeftFlipperKey Then Controller.Switch(84) = 0
    If KeyCode = RightFlipperKey Then Controller.Switch(82) = 0
    If vpmKeyUp(keycode)Then Exit Sub
    If keycode = PlungerKey Then PlaySound "fx_plunger", 0, 1, 0.1, 0.25:Plunger.Fire
End Sub

'*********
' Switches
'*********

' Slings
Dim RStep

Sub RightSlingShot_Slingshot
    PlaySoundAt SoundFX("fx_slingshot", DOFContactors), Remk
    RightSling4.Visible = 1
    Remk.RotX = 26
    RStep = 0
    vpmTimer.PulseSw 16
    RightSlingShot.TimerEnabled = 1
End Sub

Sub RightSlingShot_Timer
    Select Case RStep
        Case 1:RightSLing4.Visible = 0:RightSLing3.Visible = 1:Remk.RotX = 14
        Case 2:RightSLing3.Visible = 0:RightSLing2.Visible = 1:Remk.RotX = 2
        Case 3:RightSLing2.Visible = 0:Remk.RotX = -20:RightSlingShot.TimerEnabled = 0
    End Select
    RStep = RStep + 1
End Sub

' Bumpers
Sub Bumper1_Hit:vpmTimer.PulseSw 24:PlaySoundAt SoundFX("fx_bumper", DOFContactors), Bumper1:End Sub
Sub Bumper2_Hit:vpmTimer.PulseSw 32:PlaySoundAt SoundFX("fx_bumper", DOFContactors), Bumper2:End Sub

' Drain & Saucers
Sub Drain_Hit:PlaySoundAt "fx_drain", Drain:bsTrough.AddBall Me:End Sub
Sub sw9_Hit:PlaySoundAt "fx_hole_enter", sw9:bsSaucer.AddBall 0:End Sub

' Rollovers
Sub sw14_Hit:Controller.Switch(14) = 1:PlaySoundAt "fx_sensor", sw14:End Sub
Sub sw14_UnHit:Controller.Switch(14) = 0:End Sub

Sub sw10_Hit:Controller.Switch(10) = 1:PlaySoundAt "fx_sensor", sw10:End Sub
Sub sw10_UnHit:Controller.Switch(10) = 0:End Sub

Sub sw2_Hit:Controller.Switch(2) = 1:PlaySoundAt "fx_sensor", sw2:End Sub
Sub sw2_UnHit:Controller.Switch(2) = 0:End Sub

Sub sw6_Hit:Controller.Switch(6) = 1:PlaySoundAt "fx_sensor", sw6:DOF 111, DOFOn:End Sub
Sub sw6_UnHit:Controller.Switch(6) = 0:DOF 111, DOFOff:End Sub

Sub sw22b_Hit:Controller.Switch(22) = 1:PlaySoundAt "fx_sensor", sw22:DOF 108, DOFOn:End Sub
Sub sw22b_UnHit:Controller.Switch(22) = 0:DOF 108, DOFOff:End Sub

Sub sw22_Hit:Controller.Switch(22) = 1:PlaySoundAt "fx_sensor", sw22:DOF 109, DOFOn:End Sub
Sub sw22_UnHit:Controller.Switch(22) = 0::DOF 109, DOFOff:End Sub

Sub sw31_Hit:Controller.Switch(31) = 1:PlaySoundAt "fx_sensor", sw31:End Sub
Sub sw31_UnHit:Controller.Switch(31) = 0:End Sub

Sub sw23a_Hit:Controller.Switch(23) = 1:PlaySoundAt "fx_sensor", sw23a:End Sub
Sub sw23a_UnHit:Controller.Switch(23) = 0:End Sub

Sub sw23b_Hit:Controller.Switch(23) = 1:PlaySoundAt "fx_sensor", sw23b:End Sub
Sub sw23b_UnHit:Controller.Switch(23) = 0:End Sub

Sub sw15_Hit:Controller.Switch(15) = 1:PlaySoundAt "fx_sensor", sw15:End Sub
Sub sw15_UnHit:Controller.Switch(15) = 0:End Sub

' Spinners
Sub Spinner1_Spin:vpmTimer.PulseSw 6:PlaySoundAt "fx_spinner", Spinner1:DOF 112, DOFPulse:End Sub

'Targets
Sub sw7_Hit:vpmTimer.PulseSw 7:sw7.IsDropped = 1:sw7a.IsDropped = 0:sw7c.IsDropped = 1:sw7b.IsDropped = 0:PlaySoundAtBall SoundFX("fx_target", DOFDropTargets):End Sub
Sub sw7b_Hit:vpmTimer.PulseSw 7:sw7b.IsDropped = 1:sw7c.IsDropped = 0:sw7.IsDropped = 0:sw7a.IsDropped = 1:PlaySoundAtBall SoundFX("fx_target", DOFDropTargets):End Sub
Sub sw6b_Hit:vpmTimer.PulseSw 6:PlaySoundAtBall SoundFX("fx_target", DOFDropTargets):DOF 104, DOFPulse:sw6bp.Roty= -10:Me.TimerEnabled = 1: End Sub
Sub sw6b_Timer: sw6bp.Roty= 0:Me.TimerEnabled = 0: End Sub
Sub sw6d_Hit:vpmTimer.PulseSw 6:PlaySoundAtBall SoundFX("fx_target", DOFDropTargets):DOF 105, DOFPulse:sw6dp.Roty= -10:Me.TimerEnabled = 1:End Sub
Sub sw6d_Timer: sw6dp.Roty= 0:Me.TimerEnabled = 0: End Sub
Sub sw8_Hit:vpmTimer.PulseSw 8:PlaySoundAtBall SoundFX("fx_target", DOFDropTargets):DOF 106, DOFPulse:sw8p.Roty= -10:Me.TimerEnabled = 1:End Sub
Sub sw8_Timer: sw8p.Roty= 0:Me.TimerEnabled = 0: End Sub
Sub sw8b_Hit:vpmTimer.PulseSw 8:PlaySoundAtBall SoundFX("fx_target", DOFDropTargets):DOF 107, DOFPulse:sw8bp.Roty= -10:Me.TimerEnabled = 1:End Sub
Sub sw8b_Timer: sw8bp.Roty= 0:Me.TimerEnabled = 0: End Sub
Sub sw18_Hit:vpmTimer.PulseSw 18:PlaySoundAtBall SoundFX("fx_target", DOFDropTargets):sw18p.Roty= -10:Me.TimerEnabled = 1:End Sub
Sub sw18_Timer: sw18p.Roty= 0:Me.TimerEnabled = 0: End Sub

'Rubbers
Sub sw17a_Hit:vpmTimer.PulseSw 17:DOF 101, DOFPulse:End Sub '17 '10 point rubbers
Sub sw17b_Hit:vpmTimer.PulseSw 17:DOF 102, DOFPulse:End Sub '17
Sub sw17c_Hit:vpmTimer.PulseSw 17:DOF 103, DOFPulse:End Sub '17

' Gi Subs

Sub Gi_Hit:GiON:End Sub

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
        Next
    End If
    UpdateLeds
    UpdateLamps
End Sub

Dim l18, l19, l33
l18 = 0
l19 = 0
l33 = 0

Sub UpdateLamps
    ' playfield lights
    'Lamp 1, li1
    'Lamp 2, li2
    Lamp 3, li3
    Lamp 4, li4
    Lamp 5, li5
    Lamp 6, li6
    Lamp 7, li7
    'Lamp 8, li8
    Lampm 9, Light1
    Lamp 9, li9
    Lampm 10, Light2
    Lamp 10, li10
    Lampm 11, Light3
    Lamp 11, li11
    Lampm 12, Light4
    Lamp 12, li12
    Lampm 13, Light5
    Lamp 13, li13
    Lampm 14, Light6
    Lamp 14, li14
    Lampm 15, Light7
    Lamp 15, li15
    Lampm 16, Light8
    Lamp 16, li16

    Lamp 17, li17
    'Lamp 18, li18
    'Lamp 19, li19
    Lampm 20, Light9
    Lamp 20, li20
    Lampm 21, Light10
    Lamp 21, li21
    Lamp 22, li22
    Lampm 23, li23a
    Lamp 23, li23
    Lamp 24, li24
    Lamp 25, li25
    Lamp 26, li26
    Lamp 27, li27
    Lampm 28, li28a
    Lamp 28, li28
    Lamp 29, li29
    Lamp 30, li30
    Lamp 31, li31
    Lamp 32, li32 'game over
    'Lamp 33, li33
    'Lamp 34, li34

    'Lampm 35, li35a 'light 35 to 45 Gi lights
    Lampm 37, li37
    Lampm 37, li37b
    Lampm 37, li37d
    Lampm 38, li38
    Lampm 38, li38b
    Lampm 39, li39b
    Lampm 40, li40f
    Lampm 41, li41b
    Lampm 42, li42b
    Lamp 42, li42
    Lampm 43, li43b
    Lampm 44, li44b
    Lampm 45, li45b
    Lamp 50, li50
    Lamp 66, li66 'tilt
    Lamp 81, li81
    'Lamp 82, li82 'lights 82 to 88 maybe backglass decoration lights
    'Lamp 83, li83
    'Lamp 84, li84
    'Lamp 85, li85
    'Lamp 86, li86
    'Lamp 87, li87
    'Lamp 88, li88

    ' Lights turning on Solenoids
    If LampState(2)Then 'Imp.Agujero/Hole
        If bsSaucer.Balls Then bsSaucer.ExitSol_On
    End If

    If LampState(18)Then 'Bancada Dianas Derecho
        If l18 = 0 Then
            l18 = 1
            dtR.DropSol_On         'reset right target bank
            vpmtimer.AddTimer 2000,"l18=0 '"
        end If
    End If

    If LampState(19)Then 'Bancada Dianas Izquierda
        If l19 = 0 Then
            l19 = 1
            dtL.DropSol_On         'reset left target bank
            vpmtimer.AddTimer 2000,"l19=0 '"
        end If
    End If

    If LampState(33)Then 'TACA
        If l33 = 0 Then
            l33 = 1
            PlaySound SoundFX("fx_knocker",DOFKnocker)
            vpmtimer.AddTimer 2000,"l33=0 '"
        end If
    End If

    If LampState(34)Then 'Salida Bolas
        If bsTrough.Balls Then bsTrough.ExitSol_On
    End If
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

Digits(7) = Array(b00, b02, b05, b06, b04, b01, b03, b07)
Digits(8) = Array(b10, b12, b15, b16, b14, b11, b13)
Digits(9) = Array(b20, b22, b25, b26, b24, b21, b23)
Digits(10) = Array(b30, b32, b35, b36, b34, b31, b33, b37)
Digits(11) = Array(b40, b42, b45, b46, b44, b41, b43)
Digits(12) = Array(b50, b52, b55, b56, b54, b51, b53)
Digits(13) = Array(b60, b62, b65, b66, b64, b61, b63)

Digits(14) = Array(e00, e02, e05, e06, e04, e01, e03, e07)
Digits(15) = Array(e10, e12, e15, e16, e14, e11, e13)
Digits(16) = Array(e20, e22, e25, e26, e24, e21, e23)
Digits(17) = Array(e30, e32, e35, e36, e34, e31, e33, e37)
Digits(18) = Array(e40, e42, e45, e46, e44, e41, e43)
Digits(19) = Array(e50, e52, e55, e56, e54, e51, e53)
Digits(20) = Array(e60, e62, e65, e66, e64, e61, e63)

Digits(21) = Array(f00, f02, f05, f06, f04, f01, f03, f07)
Digits(22) = Array(f10, f12, f15, f16, f14, f11, f13)
Digits(23) = Array(f20, f22, f25, f26, f24, f21, f23)
Digits(24) = Array(f30, f32, f35, f36, f34, f31, f33, f37)
Digits(25) = Array(f40, f42, f45, f46, f44, f41, f43)
Digits(26) = Array(f50, f52, f55, f56, f54, f51, f53)
Digits(27) = Array(f60, f62, f65, f66, f64, f61, f63)

Digits(28) = Array(c00, c02, c05, c06, c04, c01, c03)
Digits(29) = Array(c10, c12, c15, c16, c14, c11, c13)

Digits(30) = Array(d00, d02, d05, d06, d04, d01, d03)
Digits(31) = Array(d10, d12, d15, d16, d14, d11, d13)

Digits(32) = Array(c20, c22, c25, c26, c24, c21, c23)

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

'*****************************
' Ball 2 Ball Collision Sound
'*****************************

Sub OnBallBallCollision(ball1, ball2, velocity)
    PlaySound("fx_collide"), 0, Csng(velocity) ^2 / 2000, Pan(ball1), 0, Pitch(ball1), 0, 0, AudioFade(ball1)
End Sub

'******************
' RealTime Updates
'******************

Sub RealTime_Timer
    RollingUpdate
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

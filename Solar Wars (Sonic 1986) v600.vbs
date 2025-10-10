'Solar Wars / IPD No. 3273 / 1986 / 4 Players
'Manufacturer:  Segasa d.b.a. Sonic, of Spain [Trade Name: Sonic]
'VPX8 table by jpsalas version 6.0.0
'vpinmame parts of the script by destruk

Option Explicit
Randomize

Const BallSize = 50
Const BallMass = 1

On Error Resume Next
ExecuteGlobal GetTextFile("controller.vbs")
If Err Then MsgBox "You need the controller.vbs in order to run this table, available in the vp10 package"
On Error Goto 0

Dim bsTrough, dtDrop, x

Const cGameName = "solarwar"

Dim VarHidden
If Table1.ShowDT = true then
    For each x in aReels
        x.visible = True
    Next
    VarHidden = 1
Else
    For each x in aReels
        x.visible = False
    Next
    VarHidden = 0
End If

if B2SOn = true then VarHidden = 1

LoadVPM "01520000", "peyper.vbs", 3.1

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
        .SplashInfoLine = "Solar Wars - Sonic 1986" & vbNewLine & "VPX table by JPSalas v6.0.0"
        .Games(cGameName).Settings.Value("sound") = 1
        '.SetDisplayPosition 0,0,GetPlayerHWnd 'uncomment if you can't see the dmd
        .HandleKeyboard = 0
        .ShowTitle = 0
        .ShowDMDOnly = 1
        .ShowFrame = 0
        .HandleMechanics = 0
        .Hidden = VarHidden
        .Games(cGameName).Settings.Value("rol") = 0
        .Run GetPlayerHWnd
        If Err Then MsgBox Err.Description
    End With

'actual dip switches are backwards and inverted in vpm - so physical dipswitch #1 is seen by vpm as #8, #2 is seen as #7, #3 as #6, etc etc
'-------------------------------------------
'REPLAY1 at 2.200.000, 2nd at 2.800.000= both off
'REPLAY1 at 1.900.000, 2nd at 2.500.000= Controller.Dip(0) Switch 1
'REPLAY1 at 2.500.000, 2nd at 3.100.000= Controller.Dip(0) Switch 2
'REPLAY1 at 2.800.000, 2nd at 3.400.000= Controller.Dip(0) Switch 1 and Controller.Dip(0) Switch 2

'BALLS/GAME= Controller.Dip(0) Switch 3 = 1=5 Balls/0=3 Balls

'Coins/Credit = Dip Bank 0, Switch 4/5
'2-5-1= Controller.Dip(0) 0/0
'1-3-2:1=Controller.Dip(0) 1/1
'3-6-1=Controller.Dip (0) 1/0
'4-8-2=Controller.Dip(0) 0/1

'Controller.Dip(0) Switch 6 is NOT USED

'MATCH = Controller.Dip(0) Switch 7 - Enabled=0, Disabled=1

'SPECIAL RAMP LAMP = Controller.Dip(0) Switch 8= 1=Solid on when lit, 0=Randomized Flashing

'Enter game audits mode = Controller.Dip(1) Switch 4 -- Enabled=0, Disabled=1
'exclusive - only select one of these at one time
'Diplay coin audits -Controller.Dip(1) Switch 1 and Controller.Dip(1) Switch 3
'Display time played audits -Controller.Dip(1) Switch 2 and Controller.Dip(1) Switch 3
'Display Free Game/Extra Ball Awards -Controller.Dip(1) Switch 1 and Controller.Dip(1) Switch 2

'BACKGROUND MUSIC= Controller.Dip(1) Switch 5 = 1=allow switch effects to interrupt music, 0=constant background music

'FACTORY RESET = Controller.Dip(1) Switch 6 (Normal Operation=0)

'LEFT LANE SCORING =Controller.Dip(1) Switch 7 = 1=Resets to 10,000 when complete, 0=Remains at 100,000 when complete

'EXTRA BALL LANE LAMP =Controller.Dip(1) Switch 8= 1-Lit Solid, 0-Flashing
' ------------------------------------------------------------------------------------------------------------------
'                         *     *     *     *     *             *      *
'   Controller.Dip(0) = (1*1 + 1*2 + 0*4 + 0*8 + 0*16 + 0*32 + 0*64 + 0*128) '01-08
'             *     *     *     *     *      *      *      *
'   Controller.Dip(1) = (0*1 + 0*2 + 0*4 + 1*8 + 0*16 + 0*32 + 1*64 + 0*128) '09-16

' Nudging
    vpmNudge.TiltSwitch = -5
    vpmNudge.Sensitivity = 3
    vpmNudge.TiltObj = Array(Bumper1, Bumper2, Bumper3, LeftSlingshot, RightSlingshot)

    ' Trough
    Set bsTrough = New cvpmBallStack
    With bsTrough
        .InitNoTrough BallRelease, 0.1, 90, 6
        .InitExitSnd SoundFX("fx_ballrel", DOFContactors), SoundFX("fx_Solenoid", DOFContactors)
    End With

    ' Drop targets
    Set dtDrop = New cvpmDropTarget
    With dtDrop
        .InitDrop Array(sw34, sw33, sw32, sw29, sw28), Array(34, 33, 32, 29, 28)
        .InitSnd SoundFX("fx_droptarget", DOFDropTargets), SoundFX("fx_resetdrop", DOFContactors)
        .CreateEvents "dtDrop"
    End With

    ' Main Timer init
    PinMAMETimer.Interval = PinMAMEInterval
    PinMAMETimer.Enabled = 1
    RealTime.Enabled = 1
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
    If KeyCode = LeftFlipperKey Then Controller.Switch(103) = 1
    If KeyCode = RightFlipperKey Then Controller.Switch(101) = 1
    If keycode = LeftTiltKey Then Nudge 90, 6:PlaySound SoundFX("fx_nudge", 0), 0, 1, -0.1, 0.25
    If keycode = RightTiltKey Then Nudge 270, 6:PlaySound SoundFX("fx_nudge", 0), 0, 1, 0.1, 0.25
    If keycode = CenterTiltKey Then Nudge 0, 8:PlaySound SoundFX("fx_nudge", 0), 0, 1, 0, 0.25
    If vpmKeyDown(keycode)Then Exit Sub
    If keycode = PlungerKey Then PlaySoundat "fx_PlungerPull", Plunger:Plunger.Pullback
    If keycode = KeyRules Then Rules
End Sub

Sub table1_KeyUp(ByVal Keycode)
    If KeyCode = LeftFlipperKey Then Controller.Switch(103) = 0
    If KeyCode = RightFlipperKey Then Controller.Switch(101) = 0
    If vpmKeyUp(keycode)Then Exit Sub
    If keycode = PlungerKey Then PlaySoundAt "fx_plunger", Plunger:Plunger.Fire
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
    vpmTimer.PulseSw 1
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
    vpmTimer.PulseSw 2
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
Sub rlband007_Hit():vpmTimer.PulseSw 18:End Sub
Sub rlband008_Hit():vpmTimer.PulseSw 16:End Sub
Sub rlband005_Hit():vpmTimer.PulseSw 23:End Sub
Sub rlband009_Hit():vpmTimer.PulseSw 24:End Sub
Sub rlband004_Hit():vpmTimer.PulseSw 35:End Sub
Sub rlband006_Hit():vpmTimer.PulseSw 35:End Sub

' Bumpers
Sub Bumper1_Hit:vpmTimer.PulseSw 3:PlaySoundAt SoundFX("fx_bumper", DOFContactors), Bumper1:End Sub
Sub Bumper2_Hit:vpmTimer.PulseSw 4:PlaySoundAt SoundFX("fx_bumper", DOFContactors), Bumper2:End Sub
Sub Bumper3_Hit:vpmTimer.PulseSw 5:PlaySoundAt SoundFX("fx_bumper", DOFContactors), Bumper3:End Sub

' Drain & holes
Sub Drain_Hit:PlaysoundAt "fx_drain", Drain:bsTrough.AddBall Me:End Sub

' Rollovers
Sub sw8_Hit:Controller.Switch(8) = 1:PlaySoundAt "fx_sensor", sw9:End Sub
Sub sw8_UnHit:Controller.Switch(8) = 0:End Sub

Sub sw9_Hit:Controller.Switch(9) = 1:PlaySoundAt "fx_sensor", Sw9:End Sub
Sub sw9_UnHit:Controller.Switch(9) = 0:End Sub

Sub sw12_Hit:Controller.Switch(12) = 1:PlaySoundAt "fx_sensor", sw12:End Sub
Sub sw12_UnHit:Controller.Switch(12) = 0:End Sub

Sub sw13_Hit:Controller.Switch(13) = 1:PlaySoundAt "fx_sensor", sw13:End Sub
Sub sw13_UnHit:Controller.Switch(13) = 0:End Sub

Sub sw15_Hit:Controller.Switch(15) = 1:PlaySoundAt "fx_sensor", Sw15:End Sub
Sub sw15_UnHit:Controller.Switch(15) = 0:End Sub

Sub sw17_Hit:Controller.Switch(17) = 1:PlaySoundAt "fx_sensor", sw17:End Sub
Sub sw17_UnHit:Controller.Switch(17) = 0:End Sub

Sub sw22_Hit:Controller.Switch(22) = 1:PlaySoundAt "fx_sensor", Sw22:End Sub
Sub sw22_UnHit:Controller.Switch(22) = 0:End Sub

Sub sw25_Hit:Controller.Switch(25) = 1:PlaySoundAt "fx_sensor", Sw25:End Sub
Sub sw25_UnHit:Controller.Switch(25) = 0:End Sub

Sub sw26_Hit:Controller.Switch(26) = 1:PlaySoundAt "fx_sensor", Sw26:End Sub
Sub sw26_UnHit:Controller.Switch(26) = 0:End Sub

Sub sw36_Hit:Controller.Switch(36) = 1:PlaySoundAt "fx_sensor", Sw36:End Sub
Sub sw36_UnHit:Controller.Switch(36) = 0:End Sub

Sub sw37_Hit:Controller.Switch(37) = 1:PlaySoundAt "fx_sensor", Sw37:End Sub
Sub sw37_UnHit:Controller.Switch(37) = 0:End Sub

'Spinners & gates

Sub sw6_Spin():vpmTimer.PulseSw 6:PlaySoundAt "fx_spinner", sw6:End Sub
Sub sw19_Hit():vpmTimer.PulseSw 19:PlaySoundAt "fx_gate", sw19:End Sub

'Targets
Sub sw7_Hit:vpmTimer.PulseSw 7:PlaySoundAtBall SoundFX("fx_target", DOFTargets):End Sub
Sub sw14_Hit:vpmTimer.PulseSw 14:PlaySoundAtBall SoundFX("fx_target", DOFTargets):End Sub
Sub sw27_Hit:vpmTimer.PulseSw 27:PlaySoundAtBall SoundFX("fx_target", DOFTargets):End Sub

'*********
'Solenoids
'*********
' solenoids 1, 2, 3 bumpers
' solenoids 4, 5 slingshots

SolCallback(6) = "dtDrop.SolDropUp"
SolCallback(7) = "vpmSolSound SoundFX(""fx_Knocker"",DOFKnocker),"
SolCallback(8) = "bsTrough.SolOut"
SolCallback(30) = "vpmNudge.SolGameOn"
SolCallback(31) = "SetLamp 131, "

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

Sub LeftFlipper_Animate: LeftFlipperTop.RotZ = LeftFlipper.CurrentAngle: End Sub
Sub RightFlipper_Animate: RightFlipperTop.RotZ = RightFlipper.CurrentAngle: End Sub

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
    Lamp 1, A01
    Lamp 2, A02
    Lamp 3, A03
    Lamp 4, A04
    Lamp 5, A05
    Lamp 6, A06
    Lamp 7, A07
    Lamp 8, A08
    Lamp 9, A09
    Lamp 10, A10
    Lamp 11, A11
    Lamp 12, A12
    Lamp 13, A13
    Lamp 14, A14
    Lamp 15, A15
    Lamp 16, A16
    Gi 26, li026 'gi
    Lamp 33, A33
    Lamp 34, A34
    Lamp 35, A35
    Lamp 36, A36
    Lamp 37, A37
    Lamp 38, A38
    Lamp 39, A39
    Lampm 41, Lbumper1a
    Lamp 41, Lbumper1b
    Lampm 42, Lbumper2a
    Lamp 42, Lbumper2b
    Lampm 43, Lbumper3a
    Lamp 43, Lbumper3b
    Lamp 44, A44
    Lamp 45, A45
    Lamp 46, A46
    Lamp 47, A47
    Lampm 48, li048 'extra ball
    Lamp 48, A48
    Lamp 49, A49
    Lamp 50, A50
    Lamp 51, A51
    Lamp 52, A52
    Lamp 53, A53
    Lamp 54, li054 'ball in play
    Lamp 55, A55
    Lamp 73, li073
    Lamp 74, li074
    Lamp 75, li075
    Lamp 76, li076
    Lamp 78, li078 'gameover
    Lamp 80, li080 'Tilt
    Lampm 131, F30
    Lamp 131, F30a
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

Sub Gi(nr, object)
    Select Case FadingStep(nr)
        Case 1:GiOn:FadingStep(nr) = -1
        Case 0:GiOff:FadingStep(nr) = -1
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
'************************************

Dim Digits(32)
Digits(0) = Array(Light1, Light2, Light3, Light4, Light13, Light14, Light19, Light20)              'ok
Digits(1) = Array(Light31, Light25, Light26, Light27, Light28, Light29, Light30)                   'ok
Digits(2) = Array(Light49, Light33, Light34, Light35, Light41, Light42, Light46)                   'ok
Digits(3) = Array(Light65, Light50, Light51, Light52, Light53, Light55, Light56, Light66)          'ok
Digits(4) = Array(Light90, Light67, Light68, Light86, Light87, Light88, Light89)                   'ok
Digits(5) = Array(Light97, Light91, Light92, Light93, Light94, Light95, Light96)                   'ok
Digits(6) = Array(Light104, Light98, Light99, Light100, Light101, Light102, Light103)              'ok

Digits(7) = Array(Light105, Light106, Light107, Light108, Light109, Light110, Light111, Light112)  'ok
Digits(8) = Array(Light119, Light113, Light114, Light115, Light116, Light117, Light118)            'ok
Digits(9) = Array(Light126, Light120, Light121, Light122, Light123, Light124, Light125)            'ok
Digits(10) = Array(Light133, Light127, Light128, Light129, Light130, Light131, Light132, Light134) 'ok
Digits(11) = Array(Light141, Light135, Light136, Light137, Light138, Light139, Light140)           'ok
Digits(12) = Array(Light148, Light142, Light143, Light144, Light145, Light146, Light147)           'ok
Digits(13) = Array(Light155, Light149, Light150, Light151, Light152, Light153, Light154)           'ok

Digits(14) = Array(Light156, Light157, Light158, Light159, Light160, Light161, Light162, Light163) 'ok
Digits(15) = Array(Light170, Light164, Light165, Light166, Light167, Light168, Light169)           'ok
Digits(16) = Array(Light177, Light171, Light172, Light173, Light174, Light175, Light176)           'ok
Digits(17) = Array(Light184, Light178, Light179, Light180, Light181, Light182, Light183, Light185) 'ok
Digits(18) = Array(Light192, Light186, Light187, Light188, Light189, Light190, Light191)           'ok
Digits(19) = Array(Light199, Light193, Light194, Light195, Light196, Light197, Light198)           'ok
Digits(20) = Array(Light206, Light200, Light201, Light202, Light203, Light204, Light205)           'ok

Digits(21) = Array(Light207, Light208, Light209, Light210, Light211, Light212, Light213, Light214) 'ok
Digits(22) = Array(Light221, Light215, Light216, Light217, Light218, Light219, Light220)           'ok
Digits(23) = Array(Light228, Light222, Light223, Light224, Light225, Light226, Light227)           'ok
Digits(24) = Array(Light235, Light229, Light230, Light231, Light232, Light233, Light234, Light236) 'ok
Digits(25) = Array(Light243, Light237, Light238, Light239, Light240, Light241, Light242)           'ok
Digits(26) = Array(Light250, Light244, Light245, Light246, Light247, Light248, Light249)           'ok
Digits(27) = Array(Light257, Light251, Light252, Light253, Light254, Light255, Light256)           'ok

Digits(28) = Array(Light264, Light258, Light259, Light260, Light261, Light262, Light263)           'ok
Digits(29) = Array(Light271, Light265, Light266, Light267, Light268, Light269, Light270)           'ok

Digits(30) = Array(Light278, Light272, Light273, Light274, Light275, Light276, Light277)           'ok
Digits(31) = Array(Light285, Light279, Light280, Light281, Light282, Light283, Light284)           'ok

Digits(32) = Array(Light292, Light286, Light287, Light288, Light289, Light290, Light291)           'ok

Sub UpdateLeds
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
Sub aTargets_Hit(idx):ActiveBall.VelZ = BallVel(Activeball) * (RND / 2):End Sub

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


'Sonic Solar Wars
'by Inkochnito
Sub editDips
    Dim vpmDips:Set vpmDips = New cvpmDips
    With vpmDips
        .AddForm 700, 600, "Solar Wars - DIP switches"
        .AddFrame 0, 0, 190, "Coin - credits", 24, Array("1 - 3 - 2:1", 24, "2 - 5 - 1", 0, "3 - 6 - 1", 8, "4 - 8 - 2", 16)                                                                              'vpm dip 4 and 5                                    '80
        .AddFrame 0, 80, 190, "Score threshold", 3, Array("1,900,000 && 2,500,000 points", 1, "2,200,000 && 2,800,000 points", 0, "2,500,000 && 3.100,000 points", 2, "2,800,000 && 3,400,000 points", 3) 'vpm dip 1 and 2  '76
        .AddFrame 0, 156, 150, "Match Feature", 64, Array("ON", 0, "OFF", 64)                                                                                                                             'vpm dip 7                                                                  '46
        .AddFrame 0, 202, 190, "Balls per game", 4, Array("3 balls", 0, "5 balls", 4)                                                                                                                     'vpm dip 3                                                              '46
        .AddFrame 0, 248, 190, "Background Music", 4096, Array("constant music", 0, "switches interrupt music", 4096)                                                                                     'vpm dip 13                                           '46
        .AddFrame 0, 294, 190, "Extra Ball Lane", 16384, Array("retains highest value (EASY)", 0, "reset to 10,000 (HARD)", 16384)                                                                        'vpm dip 15                                     '46
        .AddFrame 0, 340, 190, "Extra Ball Lane Award", 32768, Array("Randomized Extra Ball Award", 0, "Always Award Extra Ball When lit", 32768)                                                         'vpm dip 16                             '46
        .AddFrame 0, 386, 190, "Special Ramp Award", 128, Array("Randomized Special Award", 0, "Always Award Special When lit", 128)                                                                      'vpm dip 8                                      '46
        .AddFrame 0, 432, 190, "Bookkeeping", 3584, Array("bookkeeping off", 2048, "coin audits", 3328, "play audits", 3584, "replay/extra ball audits", 2816)                                            'vpm dips 9/10/11/12                        '75
        .AddChk 0, 507, 150, Array("Erase memory", 8192)                                                                                                                                                  'vpm dip 14                                                                         '15
        .AddChk 0, 522, 150, Array("not used", 32)                                                                                                                                                        'vpm dip 6                                                                            '15
        .AddLabel 30, 537, 280, 20, "After hitting OK, press F3 to reset game with new settings."
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


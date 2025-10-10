' Flash 1979 - Williams / 4 Players
' VPX8 by JPSalas v6.0.0 2024
' vpinmame script based on the script by desktruk & gaston

Option Explicit
Randomize

Const BallSize = 50
Const BallMass = 1

On Error Resume Next
ExecuteGlobal GetTextFile("controller.vbs")
If Err Then MsgBox "You need the controller.vbs in order to run this table, available in the vp10 package"
On Error Goto 0

LoadVPM "01120100", "s4.vbs", 3.02

Dim bsTrough, dtRBank, dtLBank, bsLHole, x, plungerIM

Const cGameName = "flash_l1"

Const UseSolenoids = 2
Const UseLamps = 0
Const UseGI = 0
Const UseSync = 0
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
'Const SFlipperOn = "fx_FlipperUp"
'Const SFlipperOff = "fx_FlipperDown"
Const SCoin = "fx_coin"

'************
' Table init.
'************

Sub Table1_Init
    vpmInit me

    With Controller
        .GameName = cGameName
        .SplashInfoLine = "Flash, Williams 1979" & vbNewLine & "VPX table by JPSalas v6.0.0"
        .Games(cGameName).Settings.Value("rol") = 0 '1= rotated display, 0= normal
        .Games(cGameName).Settings.Value("sound") = 1 '1 enabled rom sound
        .HandleMechanics = 0
        .HandleKeyboard = 0
        .ShowDMDOnly = 1
        .ShowFrame = 0
        .ShowTitle = 0
        .Hidden = VarHidden
        If Err Then MsgBox Err.Description
    End With
    On Error Goto 0
    Controller.SolMask(0) = 0
    vpmTimer.AddTimer 2000, "Controller.SolMask(0)=&Hffffffff'" 'ignore all solenoids - then add the timer to renable all the solenoids after 2 seconds
    Controller.Run

    ' Nudging
    vpmNudge.TiltSwitch = 47
    vpmNudge.Sensitivity = 1
    vpmNudge.TiltObj = Array(Bumper1, Bumper2, Bumper3, leftSlingshot, RightSlingshot)

    ' Trough
    Set bsTrough = New cvpmBallStack
    With bsTrough
        .InitSw 0, 48, 0, 0, 0, 0, 0, 0
        .InitKick BallRelease, 80, 6
        .InitEntrySnd "fx_Solenoid", "fx_Solenoid"
        .InitExitSnd SoundFX("fx_ballrel", DOFContactors), SoundFX("fx_Solenoid", DOFContactors)
        .Balls = 1
    End With

    ' Left 5 Drop targets
    set dtLBank = new cvpmdroptarget
    With dtLBank
        .InitDrop Array(sw36, sw35, sw34, sw33, sw32), Array(36, 35, 34, 33, 32)
        .initsnd SoundFX("fx_droptarget", DOFContactors), SoundFX("fx_resetdrop", DOFContactors)
        .AllDownSw = 37
        .CreateEvents "dtLBank"
    End With

    ' Right 3 Drop targets
    set dtRBank = new cvpmdroptarget
    With dtRBank
        .InitDrop Array(sw30, sw29, sw28), Array(30, 29, 28)
        .initsnd SoundFX("fx_droptarget", DOFContactors), SoundFX("fx_resetdrop", DOFContactors)
        .AllDownSw = 31
        .CreateEvents "dtRBank"
    End With

    ' Left Eject Hole
    Set bsLHole = New cvpmBallStack
    With bsLHole
        .InitSaucer sw27, 27, 190, 14
        .InitExitSnd SoundFX("fx_kicker", DOFContactors), SoundFX("fx_kicker", DOFContactors)
        .KickForceVar = 3
        .KickAngleVar = 1
    End With

    ' Main Timer init
    PinMAMETimer.Interval = PinMAMEInterval
    PinMAMETimer.Enabled = 1
    RealTime.Enabled = 1

    ' Turn on Gi
    vpmtimer.addtimer 1500, "GiOn '"
End Sub

'******************
' Realtime Updates
'******************

Sub RealTime_Timer
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
    If keycode = PlungerKey Then PlaySoundat "fx_PlungerPull", Plunger:Plunger.Pullback
    If keycode = KeyRules Then Rules
End Sub

Sub table1_KeyUp(ByVal Keycode)
    If vpmKeyUp(keycode) Then Exit Sub
    If keycode = PlungerKey Then PlaySoundat "fx_plunger",Plunger:Plunger.Fire
End Sub

'*********
' Switches
'*********

' Slings & Rubbers
' Slings
Dim LStep, RStep

Sub LeftSlingShot_Slingshot
    PlaySoundAt SoundFX("fx_slingshot", DOFContactors), Lemk
    LeftSling4.Visible = 1
    Lemk.RotX = 26
    LStep = 0
    vpmTimer.PulseSw 41
    LeftSlingShot.TimerEnabled = 1
End Sub

Sub LeftSlingShot_Timer
    Select Case LStep
        Case 1:LeftSLing4.Visible = 0:LeftSLing3.Visible = 1:Lemk.RotX = 14
        Case 2:LeftSLing3.Visible = 0:LeftSLing2.Visible = 1:Lemk.RotX = 2
        Case 3:LeftSLing2.Visible = 0:Lemk.RotX = -20:LeftSlingShot.TimerEnabled = 0
    End Select
    LStep = LStep + 1
End Sub

Sub RightSlingShot_Slingshot
    PlaySoundAt SoundFX("fx_slingshot", DOFContactors), Remk
    RightSling4.Visible = 1
    Remk.RotX = 26
    RStep = 0
    vpmTimer.PulseSw 42
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

' scoring rubbers
Sub sw24_Hit:vpmTimer.PulseSw 24:End Sub
Sub sw25_Hit:vpmTimer.PulseSw 25:End Sub
Sub sw26_Hit:vpmTimer.PulseSw 26:End Sub
Sub sw16_Hit:vpmTimer.PulseSw 16:End Sub
Sub sw11_Hit:vpmTimer.PulseSw 11:End Sub
Sub sw38_Hit:vpmTimer.PulseSw 38:End Sub
Sub sw38a_Hit:vpmTimer.PulseSw 38:End Sub

' Bumpers
Sub Bumper1_Hit:vpmTimer.PulseSw 18:PlaySoundAt SoundFX("fx_bumper", DOFContactors), Bumper1:End Sub
Sub Bumper2_Hit:vpmTimer.PulseSw 19:PlaySoundAt SoundFX("fx_bumper", DOFContactors), Bumper2:End Sub
Sub Bumper3_Hit:vpmTimer.PulseSw 20:PlaySoundAt SoundFX("fx_bumper", DOFContactors), Bumper3:End Sub

' Spinners

Sub sw17_Spin:vpmTimer.PulseSw 17:PlaySoundAt "spinner",sw17:End Sub

' Drain & holes
Sub Drain_Hit:PlaySoundAt "fx_drain", drain:bsTrough.AddBall Me:End Sub
Sub sw27_Hit:PlaySoundAt "fx_kicker_enter",sw27:bsLHole.AddBall Me:End Sub

' Rollovers
Sub sw43_Hit:Controller.Switch(43) = 1:PlaySoundAt "fx_spinner", sw43:End Sub
Sub sw43_UnHit:Controller.Switch(43) = 0:End Sub

Sub sw44_Hit:Controller.Switch(44) = 1:PlaySoundAt "fx_spinner", sw44:End Sub
Sub sw44_UnHit:Controller.Switch(44) = 0:End Sub

Sub sw45_Hit:Controller.Switch(45) = 1:PlaySoundAt "fx_spinner", sw45:End Sub
Sub sw45_UnHit:Controller.Switch(45) = 0:End Sub

Sub sw46_Hit:Controller.Switch(46) = 1:PlaySoundAt "fx_spinner", sw46:End Sub
Sub sw46_UnHit:Controller.Switch(46) = 0:End Sub

Sub sw12_Hit:Controller.Switch(12) = 1:PlaySoundAt "fx_spinner", sw12:End Sub
Sub sw12_UnHit:Controller.Switch(12) = 0:End Sub

Sub sw13_Hit:Controller.Switch(13) = 1:PlaySoundAt "fx_spinner", sw13:End Sub
Sub sw13_UnHit:Controller.Switch(13) = 0:End Sub

Sub sw14_Hit:Controller.Switch(14) = 1:PlaySoundAt "fx_spinner", sw14:End Sub
Sub sw14_UnHit:Controller.Switch(14) = 0:End Sub

Sub sw15_Hit:Controller.Switch(15) = 1:PlaySoundAt "fx_spinner", sw15:End Sub
Sub sw15_UnHit:Controller.Switch(15) = 0:End Sub

Sub sw9_Hit:Controller.Switch(9) = 1:PlaySoundAt "fx_spinner", sw9:End Sub
Sub sw9_UnHit:Controller.Switch(9) = 0:End Sub

Sub sw21_Hit:Controller.Switch(21) = 1:PlaySoundAt "fx_spinner", sw21:End Sub
Sub sw21_UnHit:Controller.Switch(21) = 0:End Sub

Sub sw22_Hit:Controller.Switch(22) = 1:PlaySoundAt "fx_spinner", sw22:End Sub
Sub sw22_UnHit:Controller.Switch(22) = 0:End Sub

Sub sw23_Hit:Controller.Switch(23) = 1:PlaySoundAt "fx_spinner", sw23:End Sub
Sub sw23_UnHit:Controller.Switch(23) = 0:End Sub

Sub sw10_Hit:Controller.Switch(10) = 1:PlaySoundAt "fx_spinner", sw10:End Sub
Sub sw10_UnHit:Controller.Switch(10) = 0:End Sub

' Targets
Sub sw39_Hit:vpmTimer.PulseSw 39:PlaySoundAtBall SoundFX("fx_target", DOFTargets):End Sub
Sub sw40_Hit:vpmTimer.PulseSw 40:PlaySoundAtBall SoundFX("fx_target", DOFTargets):End Sub

'*********
'Solenoids
'*********
SolCallback(1) = "bsTrough.SolOut"
SolCallback(2) = "SolReset 1,"
SolCallback(3) = "SolReset 2,"
SolCallback(4) = "dtRBank.SolDropUp"
SolCallback(5) = "bsLHole.SolOut"
solcallback(6) = "SolFlash"
solcallback(14) = "vpmsolsound SoundFX(""fx_knocker"",DOFKnocker),"
solcallback(23) = "SolRun"

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

Sub LeftFlipper_Animate:LeftFlipperTop.RotZ = LeftFlipper.CurrentAngle: End Sub
Sub RightFlipper_Animate: RightFlipperTop.RotZ = RightFlipper.CurrentAngle: End Sub
Sub RightFlipper1_Animate: RightFlipperTop1.RotZ = RightFlipper1.CurrentAngle: End Sub

' flippers hit Sound

Sub LeftFlipper_Collide(parm)
    PlaySound "fx_rubber_flipper", 0, parm / 60, pan(ActiveBall), 0.1, 0, 0, 0, AudioFade(ActiveBall)
End Sub

Sub RightFlipper_Collide(parm)
    PlaySound "fx_rubber_flipper", 0, parm / 60, pan(ActiveBall), 0.1, 0, 0, 0, AudioFade(ActiveBall)
End Sub

Sub RightFlipper1_Collide(parm)
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

'Solenoid subs

Sub SolReset(No, Enabled)
    If Enabled Then
        Controller.Switch(37) = False
        If no < 2 Then
            dtLbank.SolUnHit 1, True
            dtLbank.SolUnHit 2, True
            dtLbank.SolUnHit 3, True
        Else
            dtLbank.SolUnHit 4, True
            dtLbank.SolUnHit 5, True
        End If
    End If
End Sub

Sub SolRun(Enabled)
    vpmNudge.SolGameOn Enabled
End Sub

Sub SolFlash(enabled)
    If enabled Then
        f1.state = 1
        f1c.state = 1
    Else
        f1.State = 0
        f1c.State = 0
    End if
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
    Lamp 2, l2
    Lamp 3, l3
    Lamp 4, l4
    Lamp 5, l5
    Lamp 6, l6
    Lamp 7, l7
    Lamp 8, l8
    Lamp 9, l9
    Lamp 10, l10
    Lamp 12, l12
    Lamp 13, l13
    Lamp 14, l14
    Lamp 15, l15
    Lamp 16, l16
    Lamp 17, l17
    Lamp 18, l18
    Lamp 19, l19
    Lamp 20, l20
    Lamp 21, l21
    Lamp 22, l22
    Lamp 23, l23
    Lamp 24, l24
    Lamp 25, l25
    Lamp 26, l26
    Lamp 27, l27
    Lamp 28, l28
    Lamp 29, l29
    Lamp 30, l30
    Lamp 31, l31
    Lamp 32, l32
    Lamp 33, l33
    Lamp 34, l34
    Lamp 35, l35
    Lamp 36, l36
    Lamp 37, l37
    Lamp 38, l38
    Lamp 39, l39
    Lampm 43, l43b
    Lampm 43, l43e
    Lampm 43, l43g
    Lamp 43, l43
    Lampm 44, l44b
    Lampm 44, l44c
    Lamp 44, l44
    Lampm 45, l45b
    Lamp 45, l45


  Lampm 40, BumperL003
  Lamp 40, BumperL003a
  Lampm 41, BumperL001
  Lamp 41, BumperL001a
  Lampm 42, BumperL002
  Lamp 42, BumperL002a

   'Backdrop lights
  Lamp 50, l50     '50 1 can play
  Lamp 51, l51     '51 2 can play
  Lamp 52, l52     '53 3 can play
  Lamp 53, l53     '54 4 can play
  Lamp 57, l57     '57 1 player
  Lamp 58, l58     '58 2 player
  Lamp 59, l59     '59 3 player
  Lamp 60, l60     '60 4 player

  Text 54, l54, "MATCH"
  Text 55, l55, "BALL" & vbNewLine & "In PLAY"
  Text 61, l61, "TILT"
  Text 62, l62, "GAME OVER"
  Text 63, l63, "SAME PLAYER" & vbNewLine & "SHOOTS AGAIN"
  Text 64, l64, "HIGHEST SCORE"
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

'************************************
'          LEDs Display
'************************************

Dim Digits(28)

Set Digits(0) = a1
Set Digits(1) = a2
Set Digits(2) = a3
Set Digits(3) = a4
Set Digits(4) = a5
Set Digits(5) = a6

Set Digits(6) = b1
Set Digits(7) = b2
Set Digits(8) = b3
Set Digits(9) = b4
Set Digits(10) = b5
Set Digits(11) = b6

Set Digits(12) = c1
Set Digits(13) = c2
Set Digits(14) = c3
Set Digits(15) = c4
Set Digits(16) = c5
Set Digits(17) = c6

Set Digits(18) = d1
Set Digits(19) = d2
Set Digits(20) = d3
Set Digits(21) = d4
Set Digits(22) = d5
Set Digits(23) = d6

Set Digits(24) = e1
Set Digits(25) = e2
Set Digits(26) = e3
Set Digits(27) = e4

Sub UPdateLEDs
  On Error Resume Next
  Dim ChgLED, ii, jj, chg, stat
  ChgLED = Controller.ChangedLEDs(&H0000003f, &Hffffffff)
  If Not IsEmpty(ChgLED) Then
    For ii = 0 To UBound(ChgLED)
      chg = chgLED(ii, 1):stat = chgLED(ii, 2)

      Select Case stat
        Case 0:Digits(chgLED(ii, 0) ).SetValue 0    'empty
        Case 63:Digits(chgLED(ii, 0) ).SetValue 1   '0
        Case 6:Digits(chgLED(ii, 0) ).SetValue 2    '1
        Case 91:Digits(chgLED(ii, 0) ).SetValue 3   '2
        Case 79:Digits(chgLED(ii, 0) ).SetValue 4   '3
        Case 102:Digits(chgLED(ii, 0) ).SetValue 5  '4
        Case 109:Digits(chgLED(ii, 0) ).SetValue 6  '5
        Case 124:Digits(chgLED(ii, 0) ).SetValue 7  '6
        Case 125:Digits(chgLED(ii, 0) ).SetValue 7  '6
        Case 252:Digits(chgLED(ii, 0) ).SetValue 7  '6
        Case 7:Digits(chgLED(ii, 0) ).SetValue 8    '7
        Case 127:Digits(chgLED(ii, 0) ).SetValue 9  '8
        Case 103:Digits(chgLED(ii, 0) ).SetValue 10 '9
        Case 111:Digits(chgLED(ii, 0) ).SetValue 10 '9
        Case 231:Digits(chgLED(ii, 0) ).SetValue 10 '9
        Case 128:Digits(chgLED(ii, 0) ).SetValue 0  'empty
        Case 191:Digits(chgLED(ii, 0) ).SetValue 1  '0
        Case 832:Digits(chgLED(ii, 0) ).SetValue 2  '1
        Case 896:Digits(chgLED(ii, 0) ).SetValue 2  '1
        Case 768:Digits(chgLED(ii, 0) ).SetValue 2  '1
        Case 134:Digits(chgLED(ii, 0) ).SetValue 2  '1
        Case 219:Digits(chgLED(ii, 0) ).SetValue 3  '2
        Case 207:Digits(chgLED(ii, 0) ).SetValue 4  '3
        Case 230:Digits(chgLED(ii, 0) ).SetValue 5  '4
        Case 237:Digits(chgLED(ii, 0) ).SetValue 6  '5
        Case 253:Digits(chgLED(ii, 0) ).SetValue 7  '6
        Case 135:Digits(chgLED(ii, 0) ).SetValue 8  '7
        Case 255:Digits(chgLED(ii, 0) ).SetValue 9  '8
        Case 239:Digits(chgLED(ii, 0) ).SetValue 10 '9
      End Select
    Next
  End IF
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

'******************
' GI routines
'******************

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

Dim OldGiState
OldGiState = -1 'start witht he Gi off

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

'******
' Rules
'******
Dim Msg(17)
Sub Rules()
    Msg(0) = "Flash - by Williams" &Chr(10) &Chr(10)
    Msg(1) = ""
    Msg(2) = "Insert coin and wait for machine to reset before inserting coin for"
    Msg(3) = " next player."
    Msg(4) = ""
    Msg(5) = "Making 1 - 2 - 3  lights 2x. Making 1 - 2 - 3 - 4  lights 3x."
    Msg(6) = "Making 3 bank drop targets advances thru Thunder, Lightning, Tempest"
    Msg(7) = "and Super Flash."
    Msg(8) = ""
    Msg(9) = "Making 5 bank targets 1st time advances Hole Kicker value, 2nd"
    Msg(10) = "time lights Extra Ball, 3rd time lights out lane specials."
    Msg(11) = ""
    Msg(12) = "Tilt penalty - Ball in Play - does not disqualify player."
    Msg(13) = "Special scores 1 Credit."
    Msg(14) = "Beating highest score scores 3 credits."
    Msg(15) = "Matching last two numbers on score with numbers in match window"
    Msg(16) = "on backglass scores 1 credit"
    Msg(17) = ""

    For X = 1 To 17
        Msg(0) = Msg(0) + Msg(X) &Chr(13)
    Next
    MsgBox Msg(0), , "         Instructions and Rule Card"
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


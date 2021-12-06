
 ' Harlem Globettroters on Tour - Bally 1979
 ' VP9 - VPM version by JPSalas 2009, version 2.2 FS
 ' Uses 7 digits ROM bootleg
 ' You need both roms: hglbtrtr.zip and hglbtrtb.zip
 ' Script based on Gaston's script

'VPX Conversion - -=DOZER=- 2018

 Option Explicit
 Randomize

Dim Ballsize,BallMass
BallSize = 50
BallMass = (Ballsize^3)/125000

On Error Resume Next
ExecuteGlobal GetTextFile("controller.vbs")
If Err Then MsgBox "You need the controller.vbs in order to run this table, available in the vp10 package"
On Error Goto 0

Dim object, VarHidden

    If ShowDT = True Then
    For each object in DT_Stuff
    Object.visible = 1
    Next
        VarHidden = 1
        RL.visible = 1:RR.visible = 1
        DT4P6.material = "ombra1"
  End If

  If ShowDt = False Then
    For each object in DT_Stuff
    Object.visible = 0
    Next
        VarHidden = 1
        RL.visible = 0:RR.visible = 0

  End If

 LoadVPM "01550000", "Bally.vbs", 3.26

 Dim bsTrough, bsMSaucer, bsRSaucer, dtDrop, x, plungerIM, uHole
 Dim bump1, bump2, bump3

 'Const cGameName = "hglbtrtr" ' normal 6 digits rom
 Const cGameName = "hglbtrtb" ' bootleg 7 digits rom

 Const UseSolenoids = 2
 Const UseLamps = 0
 Const UseGI = 0
 Const UseSync = 0
 Const HandleMech = 0

 ' Standard Sounds
 Const SSolenoidOn = "Solenoid"
 Const SSolenoidOff = ""
 Const SCoin = "Coin"

 '************
 ' Table init.
 '************

 Sub table1_Init
     vpmInit me
     With Controller
         .GameName = cGameName
         If Err Then MsgBox "Can't start Game" & cGameName & vbNewLine & Err.Description:Exit Sub
         .SplashInfoLine = "Harlem Globettroters on Tour - Bally 1979" & vbNewLine & "VP9 table by JPSalas v.2.2"
         .HandleKeyboard = 0
         .ShowTitle = 0
         .ShowDMDOnly = 1
         .ShowFrame = 0
         .HandleMechanics = 0
         .Hidden = VarHidden
         '.Games(cGameName).Settings.Value("rol") = 1 '1= rotated display, 0= normal
         '.SetDisplayPosition 0,0, GetPlayerHWnd 'restore dmd window position
         On Error Resume Next
         .Run GetPlayerHWnd
         If Err Then MsgBox Err.Description
         On Error Goto 0
     End With

     ' Nudging
     vpmNudge.TiltSwitch = swTilt
     vpmNudge.Sensitivity = 1
     vpmNudge.TiltObj = Array(Bumper1, Bumper2, Bumper3, LeftSlingshot, RightSlingshot)

     ' Trough
     Set bsTrough = New cvpmBallStack
     With bsTrough
         .InitSw 0, 8, 0, 0, 0, 0, 0, 0
         .InitKick BallRelease, 40, 6
         '.InitEntrySnd "Solenoid", "Solenoid"
         '.InitExitSnd "ballrel", "Solenoid"
         .Balls = 1
     End With

     ' Drop targets
     set dtDrop = new cvpmdroptarget
     With dtDrop
         .initdrop array(Array(d1),Array(d2),Array(d3),Array(d4)), array(1, 2, 3, 4)
         '.initsnd "droptarget", "resetdrop"
     End With

     ' Middle Saucer
     Set bsMSaucer = New cvpmBallStack
     With bsMSaucer
         .InitSaucer sw24, 24, 192, 12
         .KickAngleVar = 2
         .KickForceVar = 1
          'InitExitSnd "popper", "popper"
     End With

     ' Right Saucer
     Set bsRSaucer = New cvpmBallStack
     With bsRSaucer
         .InitSaucer sw32, 32, 302, 20
         .KickAngleVar = 2
         .KickForceVar = 1
         .InitExitSnd "popper", "popper"
     End With

     ' Saucer Magnet (Using low powered Magnet to simulate drop in playfield surface around the saucer)
     Set uHole = New cvpmMagnet
     With uHole
         .InitMagnet SaucerMagnet, 3
         .GrabCenter = 0
         .MagnetOn = 1
         .CreateEvents "uHole"
     End With


     ' Main Timer init
     PinMAMETimer.Interval = PinMAMEInterval
     PinMAMETimer.Enabled = 1
 End Sub

 Sub table1_Paused:Controller.Pause = 1:End Sub
 Sub table1_unPaused:Controller.Pause = 0:End Sub

 '**********
 ' Keys
 '**********

 Sub table1_KeyDown(ByVal Keycode)
     If keycode = PlungerKey Then Plunger.Pullback:PlaySoundat "Plungerpull", Plunger
     'If keycode = 3 Then uHole.GrabCenter = 1
    If keycode = LeftTiltKey Then Nudge 90, 2:PlaySound SoundFX("fx_nudge", 0), 0, 1, -0.1, 0.25
    If keycode = RightTiltKey Then Nudge 270, 2:PlaySound SoundFX("fx_nudge", 0), 0, 1, 0.1, 0.25
    If keycode = CenterTiltKey Then Nudge 0, 2:PlaySound SoundFX("fx_nudge", 0), 0, 1, 0, 0.25
     If vpmKeyDown(keycode) Then Exit Sub
 End Sub

 Sub table1_KeyUp(ByVal Keycode)
     If vpmKeyUp(keycode) Then Exit Sub
     'If Keycode = 3 Then uHole.GrabCenter = 0
     If keycode = PlungerKey Then Plunger.Fire:PlaySoundAt "Plungerrelease", Plunger
 End Sub

 '*********
 ' Switches
 '*********

 ' Slings

Dim RStep, Lstep, Tstep, Blstep, TLstep, BRstep, TRstep
Dim dtest
Sub LeftSlingShot_Slingshot
    vpmTimer.PulseSw 37
    PlaySoundat SoundFX("Slingshot",DOFContactors),SLING3
    'PlaySoundat SoundFXDOF("slingshot",106,DOFPulse,DOFContactors), SLING3
    LSling.Visible = 0
    LSling1.Visible = 1
    sling3.TransZ = -20
    LStep = 0
    LeftSlingShot.TimerEnabled = 1
  LeftSlingShot.TimerInterval  = 25
End Sub

Sub LeftSlingShot_Timer
    Select Case LStep
        Case 3:LSLing1.Visible = 0:LSLing2.Visible = 1:sling3.TransZ = -10
        Case 4:LSLing2.Visible = 0:LSLing.Visible = 1:sling3.TransZ = 0:LeftSlingShot.TimerEnabled = 0
    End Select
    LStep = LStep + 1
End Sub

Sub RightSlingShot_Slingshot
    vpmTimer.PulseSw 36
    PlaySoundat SoundFX("Slingshot",DOFContactors),SLING4
    'PlaySoundat SoundFXDOF("slingshot",107,DOFPulse,DOFContactors), SLING4
    RSling.Visible = 0
    RSling1.Visible = 1
    SLING4.TransZ = -20
    RStep = 0
    RightSlingShot.TimerEnabled = 1
  RightSlingShot.TimerInterval  = 25
End Sub

Sub RightSlingShot_Timer
    Select Case RStep
        Case 3:RSLing1.Visible = 0:RSLing2.Visible = 1:sling4.TransZ = -10
        Case 4:RSLing2.Visible = 0:RSLing.Visible = 1:sling4.TransZ = 0:RightSlingShot.TimerEnabled = 0
    End Select
    RStep = RStep + 1
End Sub

Sub TopSlingShot_Hit
  'If Tilt>2 Then Exit Sub
    'PlaySoundat SoundFXDOF("slingshot",107,DOFPulse,DOFContactors), SLING4
    TSling.Visible = 0
    TSling1.Visible = 1
    'SLING4.TransZ = -20
    TStep = 0
    TopSlingShot.TimerEnabled = 1
  TopSlingShot.TimerInterval  = 10
    'AddScoresToPlayer 10,ActivePlayer
    'AltLights
End Sub

Sub TopSlingShot_Timer
    Select Case TStep
        Case 3:TSLing1.Visible = 0:TSLing2.Visible = 1
        Case 4:TSLing2.Visible = 0:TSLing.Visible = 1:TopSlingShot.TimerEnabled = 0
    End Select
    TStep = TStep + 1
End Sub

 Sub Sling1_Hit:vpmTimer.PulseSw 34:End Sub
 Sub Sling2_Hit:vpmTimer.PulseSw 34:End Sub

 ' Bumpers
 Sub Bumper1_Hit:vpmTimer.PulseSw 40:PlaySoundat SoundFX("bumper1",DOFContactors),bumper1:End Sub


 Sub Bumper2_Hit:vpmTimer.PulseSw 38:PlaySoundat SoundFX("bumper1",DOFContactors),bumper2:End Sub


 Sub Bumper3_Hit:vpmTimer.PulseSw 39:PlaySoundat SoundFX("bumper1",DOFContactors),bumper3:End Sub

 ' Drain & holes
 Sub Drain_Hit:Playsoundat "drain", Drain:bsTrough.AddBall Me:End Sub
 Sub sw24_Hit:PlaySoundat "kicker_enter",sw24:bsMSaucer.AddBall 0:End Sub
 Sub sw32_Hit:PlaySoundat "kicker_enter",sw32:bsRSaucer.AddBall 0:End Sub

 ' Rollovers
 Sub sw31_Hit:Controller.Switch(31) = 1:PlaySoundAtBall "Sensor":End Sub
 Sub sw31_UnHit:Controller.Switch(31) = 0:End Sub

 Sub sw22_Hit:Controller.Switch(22) = 1:PlaySoundAtBall "Sensor":End Sub
 Sub sw22_UnHit:Controller.Switch(22) = 0:End Sub

 Sub sw23_Hit:Controller.Switch(23) = 1:PlaySoundAtBall "Sensor":End Sub
 Sub sw23_UnHit:Controller.Switch(23) = 0:End Sub

 'Spinners
 Sub sw17_Spin():vpmTimer.PulseSw 17:PlaySoundat "spinner", sw17:End Sub
 Sub sw33_Spin():vpmTimer.PulseSw 33:PlaySoundat "spinner", sw33:End Sub
 Sub sw25_Spin():vpmTimer.PulseSw 25:PlaySoundat "spinner", sw25:End Sub

 ' Droptargets
 Sub d1_Hit:dtDrop.Hit 1:DT1PS.visible = 0:PlaySoundat SoundFX("droptarget",DOFTargets),D1:End Sub
 Sub d2_Hit:dtDrop.Hit 2:DT2PS.visible = 0:PlaySoundat SoundFX("droptarget",DOFTargets),D2:End Sub
 Sub d3_Hit:dtDrop.Hit 3:DT3PS.visible = 0:PlaySoundat SoundFX("droptarget",DOFTargets),D3:End Sub
 Sub d4_Hit:dtDrop.Hit 4:DT4PS.visible = 0:PlaySoundat SoundFX("droptarget",DOFTargets),D4:End Sub

 ' Targets
 Sub sw35_Hit:vpmTimer.PulseSw 35::PlaySoundat SoundFX("target",DOFTargets),sw35:End Sub
 Sub sw26_Hit:vpmTimer.PulseSw 26::PlaySoundat SoundFX("target",DOFTargets),sw26:End Sub
 Sub sw27_Hit:vpmTimer.PulseSw 27::PlaySoundat SoundFX("target",DOFTargets),sw27:End Sub
 Sub sw28_Hit:vpmTimer.PulseSw 28::PlaySoundat SoundFX("target",DOFTargets),sw28:End Sub
 Sub sw29_Hit:vpmTimer.PulseSw 29::PlaySoundat SoundFX("target",DOFTargets),sw29:End Sub
 Sub sw30_Hit:vpmTimer.PulseSw 30::PlaySoundat SoundFX("target",DOFTargets),sw30:End Sub

  ' Gates
 Sub Gate1_Hit():PlaySoundAt "gate", Gate1:End Sub
 Sub Gate2_Hit():PlaySoundAt "gate", Gate2:End Sub
 Sub Gate3_Hit():PlaySoundAt "gate", Gate3:End Sub


 '*********
 'Solenoids
 '*********

 SolCallback(7) = "Ball_Out"
 SolCallback(6) = "SolKnocker"
 SolCallBack(13) = "Mid_Saucer"
 SolCallBack(14) = "Right_Saucer"
 SolCallback(15) = "SolRaiseDrop"
 SolCallback(19) = "vpmNudge.SolGameOn"
 SolCallback(17) = "Soldiv"

Sub Ball_Out(enabled)
If enabled Then
bsTrough.ExitSol_On
PlaySoundat SoundFX("BallRel",DOFContactors),BallRelease
PlaySoundat SoundFX("Solenoid",DOFContactors),BallRelease
End If
End Sub

Sub Soldiv(Enabled)
If enabled Then
'vpmSolDiverter RightLaneGate,True,Not Enabled
'vpmSolDiverter RightLaneGate1,True,Not Enabled
RightLaneGate1.RotatetoStart
PlaySoundat SoundFX("Solenoid",DOFContactors),RightLaneGate1
Else
RightLaneGate1.RotateToEnd
PlaySoundat SoundFX("Solenoid",DOFContactors),RightLaneGate1
End If
End Sub

Sub SolKnocker(Enabled)
If enabled Then
PlaySound SoundFX("Knocker",DOFContactors)
End If
End Sub

Sub SolRaiseDrop(enabled)
If enabled Then
 dtDrop.DropSol_On
DT1PS.visible = 1
DT2PS.visible = 1
DT3PS.visible = 1
DT4PS.visible = 1
End If
End Sub

Sub mid_saucer(enabled)
If enabled Then
bsMSaucer.ExitSol_On
PlaySoundat SoundFX("Popper",DOFContactors),sw24
End If
End Sub

Sub right_saucer(enabled)
If enabled Then
bsRSaucer.ExitSol_On
PlaySoundat SoundFX("Popper",DOFContactors),sw32
End If
End Sub

 '**************
 ' Flipper Subs
 '**************

 SolCallback(sLRFlipper) = "SolRFlipper"
 SolCallback(sLLFlipper) = "SolLFlipper"

 Sub SolLFlipper(Enabled)
     If Enabled Then
         PlaySoundAt SoundFX("flipperup", DOFFlippers),leftflipper:LeftFlipper.RotateToEnd:LeftFlipper2.RotateToEnd
     Else
         PlaySoundAt SoundFX("flipperdown", DOFFlippers),leftflipper:LeftFlipper.RotateToStart:LeftFlipper2.RotateToStart
     End If
 End Sub

 Sub SolRFlipper(Enabled)
     If Enabled Then
         PlaySoundAt SoundFX("flipperup", DOFFlippers),rightflipper:RightFlipper.RotateToEnd
     Else
         PlaySoundAt SoundFX("flipperdown", DOFFlippers),rightflipper:RightFlipper.RotateToStart
     End If
 End Sub


 '****************************************
  '  JP's Fading Lamps 3.5 VP9 Fading only
  '      Based on PD's Fading Lights
  ' SetLamp 0 is Off
  ' SetLamp 1 is On
  ' LampState(x) current state
  '****************************************

 Dim LampState(200)

 AllLampsOff()
 LampTimer.Interval = 35
 LampTimer.Enabled = 1

 Sub LampTimer_Timer()
     Dim chgLamp, num, chg, ii
     chgLamp = Controller.ChangedLamps
     If Not IsEmpty(chgLamp) Then
         For ii = 0 To UBound(chgLamp)
             LampState(chgLamp(ii, 0) ) = chgLamp(ii, 1) + 4
         Next
     End If

     UpdateLamps
     'DisplayLEDs
 End Sub

 Sub UpdateLamps()
     nFadeL 1, l1
     nFadeL 2, l2
     nFadeL 3, l3
     nFadeL 4, l4
     nFadeLm 5, l5
     nFadeL 5, l500
     NFadeL 6, l6
     NFadeL 7, l7
     NFadeL 8, l8
     NFadeL 9, l9
     NFadeL 10, l10
     NFadeL 11, l11
     NFadeL 12, l12
     'NFadeT 13, l13, "Ball In Play"
     nFadeL 14, l14
     NFadeL 15, l15
     nFadeL 17, l17
     nFadeL 18, l18
     NFadeL 19, l19
     nFadeLm 20, l200
     nFadeL 20, l20
     NFadeL 21, l21
     NFadeL 22, l22
     NFadeL 23, l23
     NFadeL 24, l24
     NFadeL 25, l25
     NFadeLm 26, BUM2
     NfadeL 26, BUM2A
     'NFadeT 27, l27, "Match"
     NFadeL 28, l28
     'NFadeT 29, l29, "High Score"
     NFadeL 30, l30
     NFadeL 31, l31
     nFadeL 33, l33
     nFadeL 34, l34
     nFadeL 35, l35
     nFadeLm 36, l360
     nFadeL 36, l36
     NFadeL 37, l37
     NFadeL 38, l38
     NFadeL 39, l39
     NFadeL 40, l40
     NFadeL 41, l41
     NFadeLm 42, BUM1
     NFadeLm 42, BUM1a
     NFadeLm 42, BUM3
     NFadeL 42, BUM3a
     NFadeL 43, l43
     NFadeL 44, l44
     'NFadeT 45, l45, "Game Over"
     NFadeL 46, l46
     NFadeL 47, l47
     nFadeL 49, l49
     'nFadeLm 50, l500
     nFadeL 50, l50
     nFadeL 51, l51
     nFadeLm 52, l520
     nFadeL 52, l52
     nFadeL 53, l53
     NFadeL 54, l54
     NFadeL 55, l55
     nFadeL 56, l56
     NFadeL 57, l57
     NFadeL 58, l58
     NFadeLm 59, l59
     NFadeL 59, l59a
     NFadeL 60, l60
     'NFadeL 61, l61 'TILT
     NFadeL 62, l62
     NFadeL 63, l63

If LampState(45) = 5 Then
GO.text = "GAME OVER"
Else
GO.text = ""
End If

If LampState(61) = 5 Then
TILT.text = "TILT"
Else
TILT.text = ""
End If

If LampState(29) = 5 Then
HS.text = "HIGH SCORE"
Else
HS.text = ""
End If

If LampState(13) = 5 Then
BIP.text = "BALL IN PLAY"
Else
BIP.text = ""
End If

If LampState(27) = 5 Then
MATCH.text = "MATCH"
Else
MATCH.text = ""
End If

 End Sub

 Sub AllLampsOff():For x = 1 to 200:LampState(x) = 4:Next:UpdateLamps:UpdateLamps:Updatelamps:End Sub

  Sub SetLamp(nr, value):LampState(nr) = abs(value) + 4:End Sub

  Sub FadeW(nr, a, b, c)
      Select Case LampState(nr)
          Case 2:c.IsDropped = 1:LampState(nr) = 0                 'Off
          Case 3:b.IsDropped = 1:c.IsDropped = 0:LampState(nr) = 2 'fading...
          Case 4:a.IsDropped = 1:b.IsDropped = 0:LampState(nr) = 3 'fading...
          Case 5:c.IsDropped = 1:a.IsDropped = 0:LampState(nr) = 1 'ON
      End Select
  End Sub

  Sub FadeWm(nr, a, b, c)
      Select Case LampState(nr)
          Case 2:c.IsDropped = 1
          Case 3:b.IsDropped = 1:c.IsDropped = 0
          Case 4:a.IsDropped = 1:b.IsDropped = 0
          Case 5:c.IsDropped = 1:a.IsDropped = 0
      End Select
  End Sub

  Sub NFadeW(nr, a)
      Select Case LampState(nr)
          Case 4:a.IsDropped = 1:LampState(nr) = 0
          Case 5:a.IsDropped = 0:LampState(nr) = 1
      End Select
  End Sub

  Sub NFadeWm(nr, a)
      Select Case LampState(nr)
          Case 4:a.IsDropped = 1
          Case 5:a.IsDropped = 0
      End Select
  End Sub

  Sub NFadeWi(nr, a)
      Select Case LampState(nr)
          Case 5:a.IsDropped = 1:LampState(nr) = 0
          Case 4:a.IsDropped = 0:LampState(nr) = 1
      End Select
  End Sub

  Sub FadeL(nr, a, b)
      Select Case LampState(nr)
          Case 2:b.state = 0:LampState(nr) = 0
          Case 3:b.state = 1:LampState(nr) = 2
          Case 4:a.state = 0:LampState(nr) = 3
          Case 5:a.state = 1:LampState(nr) = 1
      End Select
  End Sub

  Sub FadeLm(nr, a, b)
      Select Case LampState(nr)
          Case 2:b.state = 0
          Case 3:b.state = 1
          Case 4:a.state = 0
          Case 5:a.state = 1
      End Select
  End Sub

  Sub NFadeL(nr, a)
      Select Case LampState(nr)
          Case 4:a.state = 0:LampState(nr) = 0
          Case 5:a.State = 1:LampState(nr) = 1
      End Select
  End Sub

  Sub NFadeLm(nr, a)
      Select Case LampState(nr)
          Case 4:a.state = 0
          Case 5:a.State = 1
      End Select
  End Sub

  Sub FadeR(nr, a)
      Select Case LampState(nr)
          Case 2:a.SetValue 3:LampState(nr) = 0
          Case 3:a.SetValue 2:LampState(nr) = 2
          Case 4:a.SetValue 1:LampState(nr) = 3
          Case 5:a.SetValue 0:LampState(nr) = 1
      End Select
  End Sub

  Sub FadeRm(nr, a)
      Select Case LampState(nr)
          Case 2:a.SetValue 3
          Case 3:a.SetValue 2
          Case 4:a.SetValue 1
          Case 5:a.SetValue 0
      End Select
  End Sub

  Sub NFadeT(nr, a, b)
      Select Case LampState(nr)
          Case 4:a.Text = "":LampState(nr) = 0
          Case 5:a.Text = b:LampState(nr) = 1
      End Select
  End Sub

  Sub NFadeTm(nr, a, b)
      Select Case LampState(nr)
          Case 4:a.Text = ""
          Case 5:a.Text = b
      End Select
  End Sub

  Sub NFadeWi(nr, a)
      Select Case LampState(nr)
          Case 4:a.IsDropped = 0:LampState(nr) = 0
          Case 5:a.IsDropped = 1:LampState(nr) = 1
      End Select
  End Sub

  Sub NFadeWim(nr, a)
      Select Case LampState(nr)
          Case 4:a.IsDropped = 0
          Case 5:a.IsDropped = 1
      End Select
  End Sub

  Sub FadeLCo(nr, a, b) 'fading collection of lights
      Dim obj
      Select Case LampState(nr)
          Case 2:vpmSolToggleObj b, Nothing, 0, 0:LampState(nr) = 0
          Case 3:vpmSolToggleObj b, Nothing, 0, 1:LampState(nr) = 2
          Case 4:vpmSolToggleObj a, Nothing, 0, 0:LampState(nr) = 3
          Case 5:vpmSolToggleObj a, Nothing, 0, 1:LampState(nr) = 1
      End Select
  End Sub

  Sub FlashL(nr, a, b) ' simple light flash, not controlled by the rom
      Select Case LampState(nr)
          Case 2:b.state = 0:LampState(nr) = 0
          Case 3:b.state = 1:LampState(nr) = 2
          Case 4:a.state = 0:LampState(nr) = 3
          Case 5:a.state = 1:LampState(nr) = 4
      End Select
  End Sub

  Sub MFadeL(nr, a, b, c) 'Light acting as a flash. C is the light number to be restored
      Select Case LampState(nr)
          Case 2:b.state = 0:LampState(nr) = 0: If LampState(c) = 1 Then SetLamp c, 1
          Case 3:b.state = 1:LampState(nr) = 2
          Case 4:a.state = 0:LampState(nr) = 3
          Case 5:a.state = 1:LampState(nr) = 1
      End Select
  End Sub

  Sub NFadeB(nr, a, b, c, d) 'New Bally Bumpers: a and b are the off state, c and d and on state, no fading
      Select Case LampState(nr)
          Case 4:a.IsDropped = 0:b.IsDropped = 0:c.IsDropped = 1:d.IsDropped = 1:LampState(nr) = 0
          Case 5:a.IsDropped = 1:b.IsDropped = 1:c.IsDropped = 0:d.IsDropped = 0:LampState(nr) = 1
      End Select
  End Sub

  Sub NFadeBm(nr, a, b, c, d)
      Select Case LampState(nr)
          Case 4:a.IsDropped = 0:b.IsDropped = 0:c.IsDropped = 1:d.IsDropped = 1
          Case 5:a.IsDropped = 1:b.IsDropped = 1:c.IsDropped = 0:d.IsDropped = 0
      End Select
  End Sub


 '**********************************************************************************************************
'Digital Display
'**********************************************************************************************************
Dim Digits(32)

' 1st Player
Digits(0) = Array(LED10,LED11,LED12,LED13,LED14,LED15,LED16)
Digits(1) = Array(LED20,LED21,LED22,LED23,LED24,LED25,LED26)
Digits(2) = Array(LED30,LED31,LED32,LED33,LED34,LED35,LED36)
Digits(3) = Array(LED40,LED41,LED42,LED43,LED44,LED45,LED46)
Digits(4) = Array(LED50,LED51,LED52,LED53,LED54,LED55,LED56)
Digits(5) = Array(LED60,LED61,LED62,LED63,LED64,LED65,LED66)
Digits(6) = Array(LED70,LED71,LED72,LED73,LED74,LED75,LED76)

' 2nd Player
Digits(7) = Array(LED80,LED81,LED82,LED83,LED84,LED85,LED86)
Digits(8) = Array(LED90,LED91,LED92,LED93,LED94,LED95,LED96)
Digits(9) = Array(LED100,LED101,LED102,LED103,LED104,LED105,LED106)
Digits(10) = Array(LED110,LED111,LED112,LED113,LED114,LED115,LED116)
Digits(11) = Array(LED120,LED121,LED122,LED123,LED124,LED125,LED126)
Digits(12) = Array(LED130,LED131,LED132,LED133,LED134,LED135,LED136)
Digits(13) = Array(LED140,LED141,LED142,LED143,LED144,LED145,LED146)

' 3rd Player
Digits(14) = Array(LED150,LED151,LED152,LED153,LED154,LED155,LED156)
Digits(15) = Array(LED160,LED161,LED162,LED163,LED164,LED165,LED166)
Digits(16) = Array(LED170,LED171,LED172,LED173,LED174,LED175,LED176)
Digits(17) = Array(LED180,LED181,LED182,LED183,LED184,LED185,LED186)
Digits(18) = Array(LED190,LED191,LED192,LED193,LED194,LED195,LED196)
Digits(19) = Array(LED200,LED201,LED202,LED203,LED204,LED205,LED206)
Digits(20) = Array(LED210,LED211,LED212,LED213,LED214,LED215,LED216)

' 4th Player
Digits(21) = Array(LED220,LED221,LED222,LED223,LED224,LED225,LED226)
Digits(22) = Array(LED230,LED231,LED232,LED233,LED234,LED235,LED236)
Digits(23) = Array(LED240,LED241,LED242,LED243,LED244,LED245,LED246)
Digits(24) = Array(LED250,LED251,LED252,LED253,LED254,LED255,LED256)
Digits(25) = Array(LED260,LED261,LED262,LED263,LED264,LED265,LED266)
Digits(26) = Array(LED270,LED271,LED272,LED273,LED274,LED275,LED276)
Digits(27) = Array(LED280,LED281,LED282,LED283,LED284,LED285,LED286)

' Credits
Digits(28) = Array(LED4,LED2,LED6,LED7,LED5,LED1,LED3)
Digits(29) = Array(LED18,LED9,LED27,LED28,LED19,LED8,LED17)
' Balls
Digits(30) = Array(LED39,LED37,LED48,LED49,LED47,LED29,LED38)
Digits(31) = Array(LED67,LED58,LED69,LED77,LED68,LED57,LED59)

Sub DisplayTimer_Timer
  Dim ChgLED,ii,num,chg,stat,obj
  ChgLed = Controller.ChangedLEDs(&Hffffffff, &Hffffffff)
If Not IsEmpty(ChgLED) Then
    If not B2SOn then
    For ii = 0 To UBound(chgLED)
      num = chgLED(ii, 0) : chg = chgLED(ii, 1) : stat = chgLED(ii, 2)
      if (num < 32) then
        For Each obj In Digits(num)
          If chg And 1 Then obj.State = stat And 1
          chg = chg\2 : stat = stat\2
        Next
      else

      end if
    next
    end if
end if
End Sub

 'Bally Harlem Globetrotters 7 digits
 'added by Inkochnito
 Sub editDips
     Dim vpmDips:Set vpmDips = New cvpmDips
     With vpmDips
         .AddForm 700, 400, "Harlem GlobeTrotters 7 digits - DIP switches"
         .AddFrame 2, 0, 190, "Maximum credits", &H03000000, Array("10 credits", 0, "15 credits", &H01000000, "25 credits", &H02000000, "free play (40 credits)", &H03000000)                    'dip 25&26
         .AddFrame 2, 87, 190, "Sound features", &H30000000, Array("chime effects", 0, "noises and no background", &H10000000, "noise effects", &H20000000, "noises and background", &H30000000) 'dip 29&30
         .AddFrame 2, 170, 190, "Special adjustment", &H00000060, Array("points", 0, "extra ball", &H00000040, "replay/extra ball", &H00000060)                                                  'dip 6&7
         .AddFrame 2, 230, 190, "5 side target lite adjustment", &H00800000, Array("targets will reset", 0, "targets are held in memory", &H00800000)                                            'dip 24
         .AddFrame 2, 276, 190, "G-L-O-B-E saucer scanning adjustment", &H00002000, Array("Globe lites do not scan", 0, "Globe lites keep scanning", &H00002000)                                 'dip 14
         .AddFrame 2, 322, 190, "Dunk shot target special", &H00400000, Array("is reset after collecting", 0, "stays lit", &H00400000)                                                           'dip 23
         .AddFrame 205, 0, 190, "High game to date", &H00200000, Array("no award", 0, "3 credits", &H00200000)                                                                                   'dip 22
         .AddFrame 205, 46, 190, "Score version", &H00100000, Array("6 digit scoring", 0, "7 digit scoring", &H00100000)                                                                         'dip 21
         .AddFrame 205, 92, 190, "Balls per game", &H40000000, Array("3 balls", 0, "5 balls", &H40000000)                                                                                        'dip 31
         .AddFrame 205, 138, 190, "Saucer targets reset", &H00000080, Array("when ball enters target saucer", 0, "on next ball in play", &H00000080)                                             'dip 8
         .AddFrame 205, 184, 190, "Super bonus", &H00004000, Array("is reset every ball", 0, "is held in memory", &H00004000)                                                                    'dip 15
         .AddFrame 205, 230, 190, "Globe special lite adjustment", &H80000000, Array("left outlane 25K only lit", 0, "left outlane 25K and Globe special lit", &H80000000)                       'dip 32
         .AddFrame 205, 276, 190, "Left and right spinner adjust", 32768, Array("left spinner only which alternates", 0, "left and right spinner lite on", 32768)                                'dip 16
         .AddChk 205, 330, 180, Array("Match feature", &H08000000)                                                                                                                               'dip 28
         .AddChk 205, 350, 115, Array("Credits displayed", &H04000000)                                                                                                                           'dip 27
         .AddLabel 50, 370, 350, 20, "After hitting OK, press F3 to reset game with new settings."
         .ViewDips
     End With
 End Sub
 Set vpmShowDips = GetRef("editDips")

' *********************************************************************
'                      Supporting Ball & Sound Functions
' *********************************************************************
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

Function Vol(ball) ' Calculates the Volume of the sound based on the ball speed
    Vol = Csng(BallVel(ball) ^2 / 2000)
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

Function Pitch(ball) ' Calculates the pitch of the sound based on the ball speed
    Pitch = BallVel(ball) * 20
End Function

Function BallVel(ball) 'Calculates the ball speed
    BallVel = INT(SQR((ball.VelX ^2) + (ball.VelY ^2) ) )
End Function

'*********************************************************************
'                 Positional Sound Playback Functions
'*********************************************************************

' Play a sound, depending on the X,Y position of the table element (especially cool for surround speaker setups, otherwise stereo panning only)
' parameters (defaults): loopcount (1), volume (1), randompitch (0), pitch (0), useexisting (0), restart (1))
' Note that this will not work (currently) for walls/slingshots as these do not feature a simple, single X,Y position
Sub PlayXYSound(soundname, tableobj, loopcount, volume, randompitch, pitch, useexisting, restart)
  PlaySound soundname, loopcount, volume, AudioPan(tableobj), randompitch, pitch, useexisting, restart, AudioFade(tableobj)
End Sub

' Similar subroutines that are less complicated to use (e.g. simply use standard parameters for the PlaySound call)
Sub PlaySoundAt(soundname, tableobj)
    PlaySound soundname, 1, 1, AudioPan(tableobj), 0,0,0, 1, AudioFade(tableobj)
End Sub

Sub PlaySoundAtVol(soundname, tableobj, vol)
    PlaySound soundname, 1, (vol), AudioPan(tableobj), 0,0,0, 1, AudioFade(tableobj)
End Sub

Sub PlaySoundAtBall(soundname)
    PlaySoundAt soundname, ActiveBall
End Sub

'*****************************************
'      JP's VP10 Rolling Sounds
'*****************************************

Const tnob = 5 ' total number of balls
ReDim rolling(tnob)
InitRolling

Sub InitRolling
    Dim i
    For i = 0 to tnob
        rolling(i) = False
    Next
End Sub

Sub RollingSoundUpdate()
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
        If BallVel(BOT(b) ) > 1 AND BOT(b).z < 30 Then
            rolling(b) = True
            PlaySound("fx_ballrolling" & b), -1, Vol(BOT(b) ), Pan(BOT(b) ), 0, Pitch(BOT(b) ), 1, 0
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
    PlaySound("fx_collide"), 0, Csng(velocity) ^2 / 2000, Pan(ball1), 0, Pitch(ball1), 0, 0
End Sub

'*********Sound Effects**************************************************************************************************
                                       'Use these for your sound effects like ball rolling, etc.

Sub Rubbers_Hit(idx)
  dim finalspeed
    finalspeed=SQR(activeball.velx * activeball.velx + activeball.vely * activeball.vely)
  If finalspeed > 20 then
    PlaySound "fx_rubber", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0
  End if
  If finalspeed >= 6 AND finalspeed <= 20 then
    RandomSoundRubber()
  End If
End Sub

Sub Posts_Hit(idx)
  dim finalspeed
    finalspeed=SQR(activeball.velx * activeball.velx + activeball.vely * activeball.vely)
  If finalspeed > 16 then
    PlaySound "fx_rubber", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0
  End if
  If finalspeed >= 6 AND finalspeed <= 16 then
    RandomSoundRubber()
  End If
End Sub

Sub RandomSoundRubber()
  Select Case Int(Rnd*3)+1
    Case 1 : PlaySound "rubber_hit_1", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0
    Case 2 : PlaySound "rubber_hit_2", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0
    Case 3 : PlaySound "rubber_hit_3", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0
  End Select
End Sub

Sub LeftFlipper_Collide(parm)
  RandomSoundFlipper()
End Sub

Sub LeftFlipper2_Collide(parm)
  RandomSoundFlipper()
End Sub

Sub RightFlipper_Collide(parm)
  RandomSoundFlipper()
End Sub

Sub RandomSoundFlipper()
  Select Case Int(Rnd*3)+1
    Case 1 : PlaySound "flip_hit_1", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0
    Case 2 : PlaySound "flip_hit_2", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0
    Case 3 : PlaySound "flip_hit_3", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0
  End Select
End Sub

Sub RandomSoundHole()
  Select Case Int(Rnd*4)+1
    Case 1 : PlaySound "fx_Hole1", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0
    Case 2 : PlaySound "fx_Hole2", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0
    Case 3 : PlaySound "fx_Hole3", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0
    Case 4 : PlaySound "fx_Hole4", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0
  End Select
End Sub

Sub Metals_Hit(idx)
RandomSoundMetal
End Sub

Sub RandomSoundMetal()
  Select Case Int(Rnd*3)+1
    Case 1 : PlaySound "fx_metal_hit_1", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0
    Case 2 : PlaySound "fx_metal_hit_2", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0
    Case 3 : PlaySound "fx_metal_hit_3", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0
  End Select
End Sub

Sub Woods_Hit(idx)
RandomSoundWood
End Sub

Sub RandomSoundWood()
  Select Case Int(Rnd*2)+1
    Case 1 : PlaySound "woodhit", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0
    Case 2 : PlaySound "woodhit2", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0
    Case 3 : PlaySound "woodhit3", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0
  End Select
End Sub

Sub Update_Stuff_Timer()
BallShadowUpdate
RollingSoundUpdate
DiverterP.ObjRotZ = RightLaneGate1.CurrentAngle + 90
FlipperLSh.RotZ = LeftFlipper2.currentangle
FlipperLSh1.RotZ = LeftFlipper.currentangle
FlipperRSh.RotZ = RightFlipper.currentangle
End Sub

'*****************************************
' Ball Shadow
'*****************************************

Dim BallShadow
BallShadow = Array (BallShadow1, BallShadow2, BallShadow3, BallShadow4, BallShadow5)


Sub BallShadowUpdate()
    Dim BOT, b
    BOT = GetBalls

  ' render the shadow for each ball
    For b = 0 to UBound(BOT)
    If BOT(b).X < table1.Width/2 Then
      BallShadow(b).X = ((BOT(b).X) - (Ballsize/6) + ((BOT(b).X - (table1.Width/2))/7)) + 10
    Else
      BallShadow(b).X = ((BOT(b).X) + (Ballsize/6) + ((BOT(b).X - (table1.Width/2))/7)) - 10
    End If

      BallShadow(b).Y = BOT(b).Y + 10
      BallShadow(b).Z = 1
    If BOT(b).Z > 20 Then
      BallShadow(b).visible = 1
    Else
      BallShadow(b).visible = 0
    End If
  Next
End Sub

Sub table1_Exit:Controller.Stop:End Sub

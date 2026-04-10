'Attack / IPD No. 2884 / October, 1980 / 4 Players
'Playmatic, of Barcelona, Spain (1968-1987)
'a VPX10 table by pedator & jpsalas & akiles
'graphics by pedator & akiles

Option Explicit
Randomize

Const BallSize = 50
Const BallMass = 1

'*****************************************************************************************************
' CREDITS
' Initial table created by Destruk, modifications by dboyrecords, and adapted VPX version by pedator
' Inspired by JPSalas previous tables and previous versions of the table by Destruk & dboyrecords.
' Ball rolling sound, leds, flasher, LUT and scripts revisited by jpsalas
' Backglass images ans some plastics images by akiles50000.
' Adapted images and redraws (apron, plastics, playfield, backglass, backdrop, cabinet, cards, etc.) for VP by pedator.
' Plus a lot of input from the whole community (sorry if we forgot you :/)
'*****************************************************************************************************

'First, try to load the Controller.vbs (DOF), which helps controlling additional hardware like lights, gears, knockers, bells and chimes (to increase realism)
'This table uses DOF via the 'SoundFX' calls that are inserted in some of the PlaySound commands, which will then fire an additional event, instead of just playing a sample/sound effect
On Error Resume Next
ExecuteGlobal GetTextFile("controller.vbs")
If Err Then MsgBox "You need the Controller.vbs file in order to run this table (installed with the VPX package in the scripts folder)"
On Error Goto 0


LoadVPM "01500000", "Play2.vbs", 3.32


'=====================      PLAYER OPTION      ===========================

Const FlippersFix = True   'Default: False. Cambiar a True para cambiar el funcionamiento de los flippers si se producen fallos en pulsaciones. / Change to True to enable an alternative flipper handling to avoid missed key presses.

'=====================   END OF PLAYER OPTION  ===========================


Dim bsTrough, bsSaucer48, bsSaucer44, dtbank, cbDerecha, Flipperactive, x

Const cGameName = "attack"

Dim VarHidden:VarHidden = Table1.ShowDT
if B2SOn = true then VarHidden = 1

' Remove desktop items in FS mode
If Table1.ShowDT then
    For each x in aReels
        x.Visible = 1
    Next
Else
    For each x in aReels
        x.Visible = 0
    Next
End If


'********************
'Standard definitions
'********************

Const UseSolenoids = 1
Const UseLamps = 0
Const UseGI = 0
Const UseSync = 0
Const HandleMech = 0

' Standard Sounds
Const SSolenoidOn = ""
Const SSolenoidOff = ""
Const SCoin = "fx_Coin"

'*********
'Solenoids
'*********

SolCallBack(1) = "SolSaucers"
SolCallback(2) = "SolDROP1"                          'Reset dropTarget 1
SolCallback(3) = "SolRandomMusic"              'possible 8-track play solenoid?
SolCallback(4) = "SolDROP2"                          'Reset dropTarget 2
SolCallback(5) = "SolTrough"               'Ball to Plunger And Reset dropTarget 3
SolCallback(6) = "vpmSolSound SoundFX(""fx_Knocker"",DOFKnocker),"
SolCallback(7) = "vpmSolDiverter SaveGate,true,"     'Unstable
SolCallback(8) = "SolFlipperEnabled"           'Unstable


Sub SolTrough(Enabled)
    If Enabled Then
        bsTrough.ExitSol_On
    dtbank.SolunHit 3,true
    End If
End Sub

Sub SolDROP1(Enabled)
    If Enabled Then
        dtbank.SolunHit 1,true
    End If
End Sub

Sub SolDROP2(Enabled)
    If Enabled Then
        dtbank.SolunHit 2,true
    End If
End Sub

Sub SolSaucers(Enabled)
  If Enabled Then
    If bsSaucer48.Balls Then bsSaucer48.ExitSol_On
    If bsSaucer44.Balls Then bsSaucer44.ExitSol_On
  End If
End Sub

Sub SolRandomMusic(Enabled)
  If Enabled Then
      Random_Music
  End If
End Sub

'inicio A

'*******************
' Flipper Subs
'*******************

  If FlippersFix = True Then
    SolCallback(sLRFlipper) = "SolRFlipper"
    SolCallback(sLLFlipper) = "SolLFlipper"
  End If

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

'fin A

Dim GameOn
GameOn = FALSE

Sub SolFlipperEnabled(Enabled)
  vpmNudge.SolGameOn Enabled
  GameOn = Enabled
End Sub

Sub Random_Music
    dim b
    b = RndNbr(21)
    Select Case b
        Case 1
      vpmtimer.addtimer 1550, "PlaySound ""attack-campeon-campeon"" '"
        case 2
      vpmtimer.addtimer 1550, "PlaySound ""attack-chica-jugamos"" '"
        case 3
      vpmtimer.addtimer 1550, "PlaySound ""attack-chica-no-te-vayas"" '"
        case 4
      vpmtimer.addtimer 1550, "PlaySound ""attack-chica-y-es-mi-favorita"" '"
    case 5
      vpmtimer.addtimer 1550, "PlaySound ""attack-eh-te-desafio"" '"
    case 6
      vpmtimer.addtimer 850, "PlaySound ""attack-grito01"" '"
    case 7
      vpmtimer.addtimer 1550, "PlaySound ""attack-grito-fantasma"" '"
    case 8
      vpmtimer.addtimer 1550, "PlaySound ""attack-te-vencere"" '"
    case 9
      vpmtimer.addtimer 1550, "PlaySound ""attack-un-rival-en-la-partida"" '"
    case 10
      vpmtimer.addtimer 850, "PlaySound ""attack-grito-fantasma"" '"
    case 11
      vpmtimer.addtimer 850, "PlaySound ""attack-aplausos"" '"
    case 12
      vpmtimer.addtimer 850, "PlaySound ""attack-aplausos2"" '"
    case 13
      vpmtimer.addtimer 1550, "PlaySound ""attack-atrevete-otra-vez"" '"
    case 14
      vpmtimer.addtimer 1550, "PlaySound ""attack-bien-jugado"" '"
    case 15
      vpmtimer.addtimer 1550, "PlaySound ""attack-ni-puta-idea"" '"
    case 16
      vpmtimer.addtimer 1550, "PlaySound ""attack-quieres-mas"" '"
    case 17
      vpmtimer.addtimer 1550, "PlaySound ""attack-repitelo-de-nuevo"" '"
    case 18
      vpmtimer.addtimer 1550, "PlaySound ""attack-sigue-practicando"" '"
    case 19
      vpmtimer.addtimer 1650, "PlaySound ""attack-tan-poderoso"" '"
    case 20
      vpmtimer.addtimer 1550, "PlaySound ""attack-tu-no-me-venceras-con-tu-poder-maligno"" '"
    case 21
      vpmtimer.addtimer 1550, "PlaySound ""attack-vamos-a-celebrarlo"" '"
    End Select
End Sub

Sub No_Music
      vpmtimer.addtimer 10, "PlaySound ""kicker_enter_center"" '"
End Sub


'************
' Table init.
'************

Sub Table1_Init
    vpmInit me
    With Controller
        .GameName = cGameName
        If Err Then MsgBox "Can't start Game" & cGameName & vbNewLine & Err.Description:Exit Sub
        .SplashInfoLine = "Attack - Playmatic 1980" & vbNewLine & "VPX table by pedator 2.0.0"
        .HandleKeyboard = 0
        .ShowTitle = 0
        .ShowDMDOnly = 1
        .ShowFrame = 0
        .HandleMechanics = 0
        .Hidden = VarHidden
        .Games(cGameName).Settings.Value("rol") = 0
        .Games(cGameName).Settings.Value("sound") = 1 '1 enabled rom sound
        '.SetDisplayPosition 0,0,GetPlayerHWnd 'uncomment if you can't see the dmd
        On Error Resume Next
        Controller.SolMask(0) = 0
        vpmTimer.AddTimer 2000, "Controller.SolMask(0)=&Hffffffff'" 'ignore all solenoids - then add the Timer to renable all the solenoids after 2 seconds
        Controller.Run GetPlayerHWnd
        On Error Goto 0
    End With

    ' Nudging
    vpmNudge.TiltSwitch = swTilt
    vpmNudge.Sensitivity = 5
    vpmNudge.TiltObj = Array(Bumper25, Bumper26)

    ' Trough
    Set bsTrough = New cvpmBallStack
    With bsTrough
    .InitNoTrough BallRelease,5,90,18
        .InitEntrySnd "fx_Solenoid", "fx_Solenoid"
        .InitExitSnd SoundFX("fx_ballrelease", DOFContactors), SoundFX("fx_Solenoid", DOFContactors)
        .Balls = 1
    End With


    'Kicker Holes
    Set bsSaucer48 = New cvpmBallStack
    bsSaucer48.InitSaucer S48, 48, 47, 13
    bsSaucer48.InitExitSnd SoundFX("fx_kicker", DOFContactors), SoundFX("fx_Solenoid", DOFContactors)
    bsSaucer48.KickForceVar = 3

    Set bsSaucer44 = New cvpmBallStack
    bsSaucer44.InitSaucer S44, 44, 195, 14
    bsSaucer44.InitExitSnd SoundFX("fx_kicker", DOFContactors), SoundFX("fx_Solenoid", DOFContactors)
    bsSaucer44.KickForceVar = 3

    Set cbDerecha = New cvpmCaptiveBall
    cbDerecha.InitCaptive CaptiveTriggerDer, CaptiveWallDer, CaptiveKickerDer, 16.5
    cbDerecha.Start
    cbDerecha.ForceTrans = 1.5
    cbDerecha.MinForce = 5
    cbDerecha.CreateEvents "cbDerecha"

    ' Drop targets
    set dtbank = new cvpmdroptarget
    dtbank.InitDrop Array(S21, S22, S23), Array(21, 22, 23)
    dtbank.initsnd SoundFX("fx_droptarget", DOFDropTargets), SoundFX("fx_resetdrop", DOFContactors)
    dtbank.CreateEvents "dtbank"

    ' Main Timer init
    PinMAMETimer.Interval = PinMAMEInterval
    PinMAMETimer.Enabled = 1
    RealTime.Enabled = 1

'    vpmMapLights aLights

End Sub

Sub table1_Paused:Controller.Pause = 1:End Sub
Sub table1_unPaused:Controller.Pause = 0:End Sub
Sub table1_exit:Controller.stop:End Sub


'******************
' RealTime Updates
'******************

Sub RealTime_Timer
    RollingUpdate
End Sub

Sub GateF_Animate: GateP.rotz = 20 - GateF.CurrentAngle: End Sub
Sub SaveGate_Animate: GateP2.rotz = 4 + SaveGate.CurrentAngle: End Sub
Sub GateR_Animate: GateTopPestana.rotx = 90 + GateR.CurrentAngle: End Sub

'**********
' Keys
'**********

Sub Table1_KeyDown(ByVal Keycode)
'inicio A
  If FlippersFix = True Then
    If keycode = LeftFlipperKey AND Flipperactive Then SolLFlipper 1
    If keycode = RightFlipperKey AND Flipperactive Then SolRFlipper 1
'fin A
    Else
'inicio B
    If KeyCode = LeftFlipperKey Then
      If GameOn Then
        Controller.Switch(57) = 1
        Controller.Switch(104) = 1
        PlaySoundAt SoundFX("fx_flipperup", DOFFlippers), LeftFlipper
        LeftFlipper.RotateToEnd
        LeftFlipperOn = 1
      End If
      Exit Sub
    End If
    If KeyCode = RightFlipperKey Then
      If GameOn Then
        Controller.Switch(56) = 1
        Controller.Switch(102) = 1
        PlaySoundAt SoundFX("fx_flipperup", DOFFlippers), RightFlipper
        RightFlipper.RotateToEnd
        RightFlipperOn = 1
      End If
      Exit Sub
    End If
    End If
'fin B
    If keycode = LeftTiltKey Then Nudge 90, 5:PlaySound SoundFX("fx_nudge", 0), 0, 1, -0.1, 0.25
    If keycode = RightTiltKey Then Nudge 270, 5:PlaySound SoundFX("fx_nudge", 0), 0, 1, 0.1, 0.25
    If keycode = CenterTiltKey Then Nudge 0, 6:PlaySound SoundFX("fx_nudge", 0), 0, 1, 0, 0.25
    If KeyCode = KeyInsertCoin1 Then vpmTimer.PulseSw 1:Exit Sub
    If KeyCode = KeyInsertCoin2 Then vpmTimer.PulseSw 2:Exit Sub
    If KeyCode = KeyInsertCoin3 Then vpmTimer.PulseSw 3:Exit Sub
    If KeyCode = keyFront Then vpmTimer.PulseSw 59
  If KeyCode = StartGameKey Then vpmtimer.addtimer 2850, "PlaySound ""attack-un-rival-en-la-partida"" '"
    If vpmKeyDown(keycode) Then Exit Sub
    If keycode = PlungerKey Then PlaySoundAt "fx_PlungerPull", Plunger:Plunger.Pullback
End Sub

Sub Table1_KeyUp(ByVal Keycode)
  If FlippersFix = True Then
'inicio A
    If keycode = LeftFlipperKey AND Flipperactive Then SolLFlipper 0
    If keycode = RightFlipperKey AND Flipperactive Then SolRFlipper 0
'fin A
    Else
'inicio B
    If KeyCode = LeftFlipperKey Then
      If Controller.Switch(104) Then
        LeftFlipper.RotateToStart
        LeftFlipperOn = 0
        PlaySoundAt SoundFX("fx_flipperdown", DOFFlippers), LeftFlipper
        Controller.Switch(57) = 0
        Controller.Switch(104) = 0
      End If
      Exit Sub
    End If
    If KeyCode = RightFlipperKey Then
      If Controller.Switch(102) Then
        RightFlipper.RotateToStart
        RightFlipperOn = 0
        PlaySoundAt SoundFX("fx_flipperdown", DOFFlippers), RightFlipper
        Controller.Switch(56) = 0
        Controller.Switch(102) = 0
      End If
      Exit Sub
    End If
    End If
'fin B
    If vpmKeyUp(keycode) Then Exit Sub
    If keycode = PlungerKey Then PlaySoundAt "fx_plunger", Plunger:Plunger.Fire
End Sub


' flippers top animations

Sub LeftFlipper_Animate: LeftFlipperTop.RotZ = LeftFlipper.CurrentAngle: End Sub
Sub RightFlipper_Animate: RightFlipperTop.RotZ = RightFlipper.CurrentAngle: End Sub


' flippers hit Sound

Sub LeftFlipper_Collide(parm)
    PlaySound "fx_rubber_flipper", 0, Vol(ActiveBall), AudioPan(ActiveBall), 0.2, 0, 0, 0, AudioFade(ActiveBall)
End Sub

Sub RightFlipper_Collide(parm)
    PlaySound "fx_rubber_flipper", 0, Vol(ActiveBall), AudioPan(ActiveBall), 0.2, 0, 0, 0, AudioFade(ActiveBall)
End Sub

'inicio A
'Lights Callback

Set LampCallback = GetRef("UpdateMultipleLamps")

Dim OldState1
Dim OldState8

OldState1 = 0
OldState8 = 0

Sub UpdateMultipleLamps

' Suena voz con falta
    If Controller.Switch(7) = True Then
    vpmtimer.addtimer 6000, "PlaySound ""attack-atrevete-otra-vez"" '"
    End If

   If Controller.Switch(8) <> OldState8 Then 'tilted deactivate the table objects
         If Controller.Switch(8)=0 Then
            VpmNudge.SolGameOn False
            Flipperactive = False
            LeftFlipper.RotateToStart
            RightFlipper.RotateToStart
      Controller.Switch(56)=0
      Controller.Switch(57)=0
      Controller.Switch(102)=0
      Controller.Switch(104)=0
            Bumper25.Threshold = 100
      Bumper26.Threshold = 100
            OldState8 = 1
        Else
            VpmNudge.SolGameOn True
            Flipperactive = True
'           ' activate bumpers in case of tilt
            Bumper25.Threshold = 1
            Bumper26.Threshold = 1
            OldState8 = 0
        End If
    End If

    If Controller.Lamp(11) <> OldState1 Then 'Game Over light
        If Controller.Lamp(11) Then
            VpmNudge.SolGameOn False
            Flipperactive = False
            LeftFlipper.RotateToStart
            RightFlipper.RotateToStart
      Controller.Switch(56)=0
      Controller.Switch(57)=0
      Controller.Switch(102)=0
      Controller.Switch(104)=0
            OldState1 = 1
        Else
            VpmNudge.SolGameOn True
            Flipperactive = True
          ' activate bumpers in case of tilt
            Bumper25.Threshold = 1
            Bumper26.Threshold = 1
            OldState1 = 0
        End If
    End If
End Sub
'fin A

'**********************************************************
'     JP's Lamp Fading for VPX and Vpinmame v4.0
' FadingStep used for all kind of lamps
' FlashLevel used for modulated flashers
' LampState keep the real lamp state in a array
'**********************************************************

Dim LampState(200), FadingStep(200)', FlashLevel(200)

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

Sub UpdateLamps()
    ' 1 is on, 0 is off
  ' Desktop lights
    NfReelm 11, GameOverR
    NfReel 11, Match
    NfReel 15, rhigh
    NfReel 27, modea 'Continuous balls
    NfReel 31, modeb 'Alternate balls

  ' Playfield and b2s lights
    Lamp 1, Light1  '10000 when lit inlane right
    Lamp 2, Light2  '10000 when lit inlane right
    Lamp 3, Light3  'Special bottom left outlane
'    Lamp 4, Light4  'Not used?
    Lampm 5, Light5a  'top flecha roja 1
    Lamp 5, light5b  'top flecha roja 2
    Lampm 6, light6a  'top flecha roja 3
    Lamp 6, light6b  'top flecha roja 4
    Lamp 7, Light7  'star green 15000 middle
'    Lamp 8, Light8  'Not used?
    Lamp 9, Light9    'Bumper Right 2
    Lamp 10, Light10  'Bumper Left 2
'    Lamp 11, Light11  'Backglass Game Over
'    Lamp 12, Light12  'Not used?
    Lamp 13, Light13    'Bumper Right
    Lamp 14, Light14    'Bumper Left
'    Lamp 15, Light15    'Backglass Highscore To Date
'    Lamp 16, Light16  'Not used?
'    Lamp 17, Light17  'backglass?
'    Lamp 18, Light18  'backglass?
'    Lamp 19, Light19  'backglass?
'    Lamp 20, Light20  'Not used?
    Lamp 21, Light21 '8000 yellow
    Lamp 22, Light22 '15000 yellow
    Lamp 23, Light23 '40000 yellow
'    Lamp 24, Light24  'Not used?
'    Lamp 25, Light25  'backglass?
'    Lamp 26, Light26  'backglass?
'    Lamp 27, Light27  'backglass Continuous balls
'    Lamp 28, Light28  'Not used?
'    Lamp 29, Light29  'backglass?
'    Lamp 30, Light30  'backglass?
'    Lamp 31, Light31  'backglass Alternate balls
'    Lamp 32, Light32  'Not used?
    Lampm 33, Light51  'Star green left down inlane
    Lamp 33, Light33   'F
    Lampm 34, Light53  'Star green left down inlane
    Lamp 34, Light34   'E
    Lamp 35, Light35   'D
'    Lamp 36, Light36  'Not used?
    Lampm 37, Light54  'Star green left down inlane
    Lamp 37, Light37   'C
    Lampm 38, Light55  'Star green left down inlane
    Lamp 38, Light38   'B
    Lamp 39, Light39   'A
'    Lamp 40, Light40  'Not used?
    Lamp 41, Light41  'X2 bonus
    Lamp 42, Light42  'X3 bonus
    Lamp 43, Light43  'X5 bonus
'    Lamp 44, Light44  'Not used?
'    Lamp 45, Light45  'backglass?
    Lamp 46, Light46  'Extra ball right outlane
'    Lampm 47, Light47b  'Extra ball center / Solenoid Left Flipper
    Lamp 47, Light47  'Solenoid Flipper
'    Lamp 48, Light48  'Not used?
    Lamp 49, Light49  'Special top captive ball
    Lamp 50, Light50  '30000 captive ball
'    Lamp 51, Light51  'Star green left down inlane
'    Lamp 53, Light53  'Star green left down inlane ?
'    Lamp 54, Light54  'Star green left down inlane ?
'    Lamp 55, Light55  'Star green left down inlane
'    Lamp 56, Light56  'Not used?
    Lamp 57, Light57  '10000
    Lamp 58, Light58  '10000
    Lamp 59, Light59  '10000
'    Lamp 60, Light60  '10000 ?
    Lamp 61, Light61  '10000
    Lamp 62, Light62  '10000
    Lamp 63, Light63  '10000

    Lampi 70, gi001  'Luces fijas
    Lampi 71, gi002
    Lampi 72, gi003
    Lampi 73, gi004
    Lampi 74, gi005
    Lampi 75, gi006
    Lampi 76, gi007
    Lampi 77, gi008
    Lampi 78, gi009
    Lampi 79, gi010
    Lampi 80, gi011
    Lampi 81, gi012
    Lampi 82, gi013
    Lampi 83, gi014
    Lampi 84, gi015
    Lampi 85, gi016
    Lampi 86, gi017
    Lampi 120, gi030p
    Lampi 121, gi032p
    Lampi 122, gi047p
    Lampi 123, gi048p
    Lampi 124, gi049p
    Lampi 125, gi050p
  Lampi 126, gi051p
    Lampi 127, gi052p
    Lampi 128, gi053p
    Lampi 129, gi054p
    Lampi 130, gi055p
    Lampi 131, gi056p
    Lampi 132, gi057p
    Lampi 133, gi058p

End Sub

' div lamp subs

' Normal Lamp & Flasher subs

Sub InitLamps()
    Dim x
    LampTimer.Interval = 20
    LampTimer.Enabled = 1
    For x = 0 to 200
        FadingStep(x) = 0
'        FlashLevel(x) = 0
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

Sub Lampm(nr, object) ' used for multiple lights, it doesn't change the fading step
    Select Case FadingStep(nr)
        Case 1:object.state = 1
        Case 0:object.state = 0
    End Select
End Sub

' Inverted lights

Sub Lampi(nr, object)
    Select Case FadingStep(nr)
        Case 0:object.state = 1:FadingStep(nr) = -1
        Case 1:object.state = 0:FadingStep(nr) = -1
    End Select
End Sub

Sub Lampim(nr, object) ' used for multiple lights, it doesn't change the fading step
    Select Case FadingStep(nr)
        Case 0:object.state = 1
        Case 1:object.state = 0
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

FlipperPower = 5000
FlipperElasticity = 0.8
FullStrokeEOS_Torque = 0.3 ' EOS Torque when flipper hold up ( EOS Coil is fully charged. Ampere increase due to flipper can't move or when it pushed back when "On". EOS Coil have more power )
LiveStrokeEOS_Torque = 0.2 ' EOS Torque when flipper rotate to end ( When flipper move, EOS coil have less Ampere due to flipper can freely move. EOS Coil have less power )

LeftFlipper.EOSTorqueAngle = 10
RightFlipper.EOSTorqueAngle = 10

SOSTorque = 0.1
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
    If RightFlipper.CurrentAngle <= RightFlipper.StartAngle + SOSAngle Then RightFlipper.Strength = FlipperPower * SOSTorque else RightFlipper.Strength = FlipperPower:End If

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

'************************************
' Diverse Collection Hit Sounds v3.0
'************************************
Sub aMetals_Hit(idx):PlaySoundAtBall "fx_metalhit":End Sub
Sub aMetalWires_Hit(idx):PlaySoundAtBall "fx_metalwire":End Sub
Sub aRubber_ShortBands_Hit(idx):PlaySoundAtBall "fx_rubber_band":End Sub
Sub aRubber_LongBands_Hit(idx):PlaySoundAtBall "fx_rubber_longband":End Sub
Sub aRubber_Posts_Hit(idx):PlaySoundAtBall "fx_rubber_post":End Sub
Sub aRubber_Pins_Hit(idx):PlaySoundAtBall "fx_rubber_pin":End Sub
Sub aRubber_Pegs_Hit(idx):PlaySoundAtBall "fx_rubber_peg":End Sub
Sub aPlastics_Hit(idx):PlaySoundAtBall "fx_plastichit":End Sub
Sub aGates_Hit(idx):PlaySoundAtBall "gate4":End Sub
Sub aWoods_Hit(idx):PlaySoundAtBall "fx_woodhit":End Sub
Sub aDropTargets_Hit(idx):PlaySoundAt SoundFX("fx_droptarget", DOFDropTargets), aDropTargets(idx):End Sub
Sub aTargets_Hit(idx):PlaySoundAt SoundFX("fx_target", DOFTargets), aTargets(idx):End Sub
Sub aRollovers_Hit(idx):PlaySoundAt "fx_sensor", aRollovers(idx):End Sub

Sub CaptiveKickerDer_UnHit:PlaySound "fx_collide":Me.TimerEnabled = 1:CaptiveKickerDer.Enabled = 0:End Sub
Sub CaptiveKickerDer_Timer:CaptiveKickerDer.Enabled = 1:Me.TimerEnabled = 0:End Sub

'*********
' Bumpers
'*********
Sub Bumper25_Hit:vpmTimer.PulseSw 25:PlaySoundAt SoundFX("fx_bumper4", DOFContactors), bumper25:DOF 105, DOFPulse:End Sub '25 - Bumper
Sub Bumper26_Hit:vpmTimer.PulseSw 26:PlaySoundAt SoundFX("fx_bumper4", DOFContactors), bumper26:DOF 105, DOFPulse:End Sub '25 - Bumper

' Scoring rubbers
Sub sw11_Hit:vpmtimer.pulsesw 11:End Sub '11 - 100 Points Switch ok
Sub sw12_Hit:vpmtimer.pulsesw 12:End Sub '12 - 10 Points Switch ok
Sub sw13_Hit:vpmtimer.pulsesw 13:End Sub '13 - 20 Points Switch ok
Sub sw14_Hit:vpmtimer.pulsesw 14:End Sub '14 - 10 Points Switch ok
Sub sw15_Hit:vpmtimer.pulsesw 15:End Sub '15 - 200 Points Switch ok
Sub sw16_Hit:vpmtimer.pulsesw 16:End Sub '16 - 50 Points Switch
Sub sw17_Hit:vpmtimer.pulsesw 17:End Sub '17 - 50 Points Switch ok
Sub sw18_Hit:vpmtimer.pulsesw 18:End Sub '18 - 50 Points Switch


'*********
' Rollovers
'*********
Sub sw24_Hit:Controller.Switch(24) = 1:PlaySoundAt "fx_sensor", sw24:End Sub '24 - Right Outlane
Sub sw24_UnHit:Controller.Switch(24) = 0:End Sub
Sub sw31_Hit:Controller.Switch(31) = 1:PlaySoundAt "fx_sensor", sw31:End Sub '31 - Right Inlane
Sub sw31_UnHit:Controller.Switch(31) = 0:End Sub
Sub sw32_Hit:Controller.Switch(32) = 1:PlaySoundAt "fx_sensor", sw32:End Sub '32 - Right Inlane
Sub sw32_UnHit:Controller.Switch(32) = 0:End Sub
Sub sw33_Hit:Controller.Switch(33) = 1:PlaySoundAt "fx_sensor", sw33:End Sub '33 - Top Right Red Rollover / D / 5000
Sub sw33_UnHit:Controller.Switch(33) = 0:End Sub
Sub sw34_Hit:Controller.Switch(34) = 1:PlaySoundAt "fx_sensor", sw34:End Sub '34 - Middle Center Captive Ball Rollover / Open Gate
Sub sw34_UnHit:Controller.Switch(34) = 0:End Sub
Sub sw35_Hit:Controller.Switch(35) = 1:PlaySoundAt "fx_sensor", sw35:End Sub '35 - Top Right Rollover / C
Sub sw35_UnHit:Controller.Switch(35) = 0:End Sub
Sub sw36_Hit:Controller.Switch(36) = 1:PlaySoundAt "fx_sensor", sw36:End Sub '36 - Top Left Rollover / B
Sub sw36_UnHit:Controller.Switch(36) = 0:End Sub
Sub sw37_Hit:Controller.Switch(37) = 1:PlaySoundAt "fx_sensor", sw37:End Sub '37 - Top Left Red Rollover / A / 5000
Sub sw37_UnHit:Controller.Switch(37) = 0:End Sub
Sub sw41_Hit:Controller.Switch(41) = 1:PlaySoundAt "fx_sensor", sw41:End Sub '41 - Left Outlane Special / 10000
Sub sw41_UnHit:Controller.Switch(41) = 0:End Sub
Sub sw42_Hit:Controller.Switch(42) = 1:PlaySoundAt "fx_sensor", sw42:End Sub '42 - Middle Right Green Rollover / 15000 when lite
Sub sw42_UnHit:Controller.Switch(42) = 0:End Sub
Sub sw43_Hit:Controller.Switch(43) = 1:PlaySoundAt "fx_sensor", sw43:End Sub '43 - Left Bottom Inlane Green Rollover
Sub sw43_UnHit:Controller.Switch(43) = 0:End Sub
Sub sw45_Hit:Controller.Switch(45) = 1:PlaySoundAt "fx_sensor", sw45:End Sub '45 - Left Bottom Inlane Green Rollover
Sub sw45_UnHit:Controller.Switch(45) = 0:End Sub
Sub sw46_Hit:Controller.Switch(46) = 1:PlaySoundAt "fx_sensor", sw46:End Sub '46 - Left Bottom Inlane Green Rollover
Sub sw46_UnHit:Controller.Switch(46) = 0:End Sub
Sub sw47_Hit:Controller.Switch(47) = 1:PlaySoundAt "fx_sensor", sw47:End Sub '47 - Left Bottom Inlane Green Rollover
Sub sw47_UnHit:Controller.Switch(47) = 0:End Sub


'*********
' Targets
'*********
Sub sw28_Hit:vpmTimer.PulseSw 28:PlaySoundAt "fx_target", sw28:End Sub '28 - Back Right Standup


'************************
'      Drop Targets
'************************
Sub s21_Animate: dt1p.z= s21.CurrentAnimOffset -10:End Sub
Sub s22_Animate: dt2p.z= s22.CurrentAnimOffset -10:End Sub
Sub s23_Animate: dt3p.z= s23.CurrentAnimOffset -10:End Sub

Sub S21_Hit:dtbank.Hit 1:End Sub                      '21 - Left Drop Target
Sub S22_Hit:dtbank.Hit 2:End Sub                      '22 - Right Drop Target
Sub S23_Hit:dtbank.Hit 3:End Sub                      '23 - Top Drop Target


'*********
' Spinner
'*********
Sub sw38_Spin:vpmTimer.PulseSw 38:PlaySoundAt "fx_spinner", sw38:End Sub '38 - Spinner center
Sub sw38_Animate:PrimitiveTuboSpin.RotZ = sw38.CurrentAngle:End Sub


'*********
' Drain & holes
'*********
Sub S5_Hit:PlaysoundAt "fx_drain", S5:bsTrough.AddBall Me:End Sub
Sub S48_Hit:PlaysoundAt "fx_kicker_enter", S48:bsSaucer48.AddBall 0:End Sub '48 Left Kicker
Sub S44_Hit:PlaysoundAt "fx_kicker_enter", S44:bsSaucer44.AddBall 0:UpdateVoces:End Sub '44 Top Center Kicker y Voces al expulsar bola
Sub Gate1_Hit:UpdateVoces:End Sub 'Voces al sacar bola al carril de saque


'**********************
' Ball Collision Sound
'**********************

Sub OnBallBallCollision(ball1, ball2, velocity)
    PlaySound("fx_collide"), 0, Csng(velocity) ^2 / 2000, AudioPan(ball1), 0, Pitch(ball1), 0, 0, AudioFade(ball1)
End Sub

'***************************************************************
'             Supporting Ball & Sound Functions v4.0
'***************************************************************

Dim TableWidth, TableHeight

TableWidth = Table1.width
TableHeight = Table1.height

Function Vol(ball) ' Calculates the Volume of the sound based on the ball speed
    Vol = Csng(BallVel(ball) ^2 / 2000)
End Function

Function AudioPan(ball) ' Calculates the pan for a ball based on the X position on the table. "table1" is the name of the table
    Dim tmp
    tmp = ball.x * 2 / TableWidth-1
    If tmp> 0 Then
        AudioPan = Csng(tmp ^10)
    Else
        AudioPan = Csng(-((- tmp) ^10) )
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
    PlaySound soundname, 0, 1, AudioPan(tableobj), 0.2, 0, 0, 0, AudioFade(tableobj)
End Sub

Sub PlaySoundAtBall(soundname) ' play a sound at the ball position, like rubbers, targets, metals, plastics
    PlaySound soundname, 0, Vol(ActiveBall), AudioPan(ActiveBall), 0.2, Pitch(ActiveBall) * 10, 0, 0, AudioFade(ActiveBall)
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
            Else
                ballpitch = Pitch(BOT(b) ) + 20000 'increase the pitch on the upper playfield
            End If
            rolling(b) = True
            PlaySound("fx_ballrolling" & b), -1, Vol(BOT(b) ), AudioPan(BOT(b) ), 0, ballpitch, 1, 0, AudioFade(BOT(b) )
        Else
            If rolling(b) = True Then
                StopSound("fx_ballrolling" & b)
                rolling(b) = False
            End If
        End If

        ' rothbauerw's Dropping Sounds
        If BOT(b).VelZ <-1 and BOT(b).z <55 and BOT(b).z> 27 Then 'height adjust for ball drop sounds
            PlaySound "fx_balldrop", 0, ABS(BOT(b).velz) / 17, AudioPan(BOT(b) ), 0, Pitch(BOT(b) ), 1, 0, AudioFade(BOT(b) )
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

'*********************
'LED's based on Eala's
'*********************
Dim Digits(32)
Digits(0) = Array(a00, a02, a05, a06, a04, a01, a03)
Digits(1) = Array(a10, a12, a15, a16, a14, a11, a13)
Digits(2) = Array(a20, a22, a25, a26, a24, a21, a23)
Digits(3) = Array(a30, a32, a35, a36, a34, a31, a33, a07)
Digits(4) = Array(a40, a42, a45, a46, a44, a41, a43)
Digits(5) = Array(a50, a52, a55, a56, a54, a51, a53)

Digits(6) = Array(e00, e02, e05, e06, e04, e01, e03)
Digits(7) = Array(e10, e12, e15, e16, e14, e11, e13)
Digits(8) = Array(e20, e22, e25, e26, e24, e21, e23)
Digits(9) = Array(e30, e32, e35, e36, e34, e31, e33, e07)
Digits(10) = Array(e40, e42, e45, e46, e44, e41, e43)
Digits(11) = Array(e50, e52, e55, e56, e54, e51, e53)

Digits(12) = Array(b00, b02, b05, b06, b04, b01, b03)
Digits(13) = Array(b10, b12, b15, b16, b14, b11, b13)
Digits(14) = Array(b20, b22, b25, b26, b24, b21, b23)
Digits(15) = Array(b30, b32, b35, b36, b34, b31, b33, b07)
Digits(16) = Array(b40, b42, b45, b46, b44, b41, b43)
Digits(17) = Array(b50, b52, b55, b56, b54, b51, b53)

Digits(18) = Array(f00, f02, f05, f06, f04, f01, f03)
Digits(19) = Array(f10, f12, f15, f16, f14, f11, f13)
Digits(20) = Array(f20, f22, f25, f26, f24, f21, f23)
Digits(21) = Array(f30, f32, f35, f36, f34, f31, f33, f07)
Digits(22) = Array(f40, f42, f45, f46, f44, f41, f43)
Digits(23) = Array(f50, f52, f55, f56, f54, f51, f53)

Digits(24) = Array(d00, d02, d05, d06, d04, d01, d03)
Digits(25) = Array(d10, d12, d15, d16, d14, d11, d13)
Digits(26) = Array(d20, d22, d25, d26, d24, d21, d23)

Digits(27) = Array(c00, c02, c05, c06, c04, c01, c03)
Digits(28) = Array(c10, c12, c15, c16, c14, c11, c13)

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
' Table Options F12 User Options
'*********************************
' Table1.Option arguments are:
' - option name, minimum value, maximum value, step between valid values, default value, unit (0=None, 1=Percent), an optional array of literal strings

Dim LUTImage
Dim Card2Image
Dim Kickvoices
Dim BumperCaps

Sub Table1_OptionEvent(ByVal eventId)
    Dim x, y

    'Cards
    Card2Image = Table1.Option("Select Cards Language", 0, 1, 1, 0, 0, Array("English", "Spanish") )
    UpdateCard2

    'LUT
    LutImage = Table1.Option("Select LUT", 0, 21, 1, 0, 0, Array("Normal 0", "Normal 1", "Normal 2", "Normal 3", "Normal 4", "Normal 5", "Normal 6", "Normal 7", "Normal 8", "Normal 9", "Normal 10", _
        "Warm 0", "Warm 1", "Warm 2", "Warm 3", "Warm 4", "Warm 5", "Warm 6", "Warm 7", "Warm 8", "Warm 9", "Warm 10") )
    UpdateLUT

    ' Cabinet rails
    x = Table1.Option("Cabinet Rails", 0, 1, 1, 1, 0, Array("Hide", "Show") )
    For each y in aRails:y.visible = x:next

    ' LockBar
    x = Table1.Option("Lockbar", 0, 1, 1, 1, 0, Array("Hide", "Show") )
    For each y in aLockBar:y.visible = x:next

    ' Desktop Objects
  If Table1.ShowDT then
    x = Table1.Option("Desktop objects", 0, 1, 1, 1, 0, Array("Hide", "Show") )
    For each y in aReels:y.visible = x:next
  End If

  'Voces
  Kickvoices = Table1.Option("Sound Voices", 0, 1, 1, 0, 0, Array("Yes", "No") )
  UpdateVoces

  'Bumper Cap Points
  BumperCaps = Table1.Option("Select BumperCap Points", 0, 1, 1, 0, 0, Array("1K/100", "10K/1000") )
    UpdateBumpers

End Sub

Sub UpdateCard2
    Select Case Card2Image
        Case 0:ApronCardSx.Image = "Cards-instructions-en":ApronCardDx.Image = "Cards-balls-en"
        Case 1:ApronCardSx.Image = "Cards-instructions-es":ApronCardDx.Image = "Cards-balls-es"
    End Select
End Sub

Sub UpdateVoces
    Select Case Kickvoices
        Case 0:Random_Music
        Case 1:No_Music
    End Select
End Sub

Sub UpdateBumpers
    Select Case BumperCaps
        Case 0:Primitive010.Image = "bumper-playmatic-1000":Primitive011.Image = "bumper-playmatic-1000"
        Case 1:Primitive010.Image = "bumper-playmatic-10000":Primitive011.Image = "bumper-playmatic-10000"
    End Select
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


'DIPSWITCHES

' Press Test key (7), pressing test advance the options (01 to 26). Press Start key (1) for change values for selected option.
' Presiona tecla de Test (7), presionando la tecla de test cambias el tipo de opción (01 a 26). Presiona tecla Start (1) para cambiar valores de la opción.

' 01 - Número maximo de creditos (00 to 99). Default: 15
' 02 - Puntuación HIGH SCORE (000 to 990). Default: 700.000
' 03 - Premio tanteo 1º (000 to 990). Default: 500.000
' 04 - Premio tanteo 2º (000 to 990). Default: 650.000
' 05 - Premio tanteo 3º (000 to 990). Default: 000.000
' 06 - Partidas por moneda 1º selector (0,3-0,5-1-1,5 to 39). Default: 0,33
' 07 - Partidas por moneda 2º selector (0,3-0,5-1-1,5 to 39). Default: 2
' 08 - Partidas por moneda 1º selector (0,3-0,5-1-1,5 to 39). Default: 5
' 09 - Número máximo de EXTRA BALLS (0-1-2-3 bolas extras). Default: 3
' 10 - Partidas a dar por HIGH SCORE (0-1-2-3 partidas). Default: 1
' 11 - Música solo en premios, inicio y fin (0=NO, 1=SI). Default: 0
' 12 - Música cada 10 minutos (0=SI, 1=NO). Default: 0
' 13 - Tipo de sonido (0=Efectos, 1=Campana). Default: 0
' 14 - Premio por SPECIAL (0=Partida, 1=Bola). Default: 0
' 15 - Premio por tanteo (0=Partida, 1=Bola). Default: 0
' 16 - Bajada automática del HIGH SCORE (0=SI, 1=NO). Default: 0
' 17 - Lotería (0=SI, 1=NO). Default: 0
' 18 - Número de bolas por partida (0=3 bolas, 1=5 bolas). Default: 0
' 19 - Partidas para un jugador (0=Varias, 1=1 sola). Default: 0
' 20 - Orden de los jugadores (0=Alterno, 1=Seguido). Default: 0
' 21 - Puntuación del BUMPER (0=1K/100, 1=10K/1000). Default: 0
' 22 - Puntuación dianas móviles (0=7000, 1=10000). Default: 0
' 23 - Pasillo bola cautiva (0=1000, 1=10000). Default: 0
' 24 - Banderola (0=100, 1=1000). Default: 0
' 25 - Luces pasillo banderola (0=memorizadas, 1=apagadas). Default: 0
' 26 - No utilizada

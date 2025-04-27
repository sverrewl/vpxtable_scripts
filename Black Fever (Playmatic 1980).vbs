'Black Fever / IPD No. 3645 / December, 1980 / 4 Players
'Playmatic, of Barcelona, Spain (1968-1987)
'a VPX10 table by pedator & jpsalas & akiles50000
'graphics by pedator
' Thalamus - seems ok

Option Explicit
Randomize

Const BallSize = 50
Const BallMass = 1

'*****************************************************************************************************
' CREDITS
' Initial 1.0 table created by pedator, December 24, 2024
' Ball rolling sound, leds, flasher, Lamp fading, LUT and scripts revisited by jpsalas
' Real Time Flipper adjustments - by JLouLouLou & JPSalas
' Plus a lot of input from the whole community (sorry if we forgot you :/)
'*****************************************************************************************************


'First, try to load the Controller.vbs (DOF), which helps controlling additional hardware like lights, gears, knockers, bells and chimes (to increase realism)
'This table uses DOF via the 'SoundFX' calls that are inserted in some of the PlaySound commands, which will then fire an additional event, instead of just playing a sample/sound effect
On Error Resume Next
ExecuteGlobal GetTextFile("controller.vbs")
If Err Then MsgBox "You need the Controller.vbs file in order to run this table (installed with the VPX package in the scripts folder)"
On Error Goto 0

LoadVPM "01500000", "play2.vbs", 3.32

'=====================      PLAYER OPTION      ===========================

Const FlippersFix = True   'Cambiar a True para cambiar el funcionamiento de los flippers si se producen fallos en pulsaciones. / Change to True to enable an alternative flipper handling to avoid missed key presses.
Const ApronCardES = False     'Cambiar a True para ver tarjetas de instrucciones en español. / Change to True for show instructions card in spanish.
Const ApronReplayES = False  'Cambiar a True para ver tarjeta de Partidas en español. / Change to True for show Replay card in spanish.
Const BumperCap1K = True    'Cambiar a True para cambiar la imagen del bumper a 1000 points. / Change to True for show bumper image of 1000 points.
Const PlayfieldRed = False    'Cambiar a True para cambiar la imagen alternativa del tablero de chica pelirroja. / Change to True for show alternate playfield image of red headed woman.
Const MusicAttract = True    'Cambiar a True para reproducir musica aleatoria durante Attract mode (cada minuto aprox) / Change to True for play random music in attract mode.

If ApronCardES Then
    ApronCardSx.Image = "card-instructions-es"
End If

If ApronReplayES Then
    ApronCardDx.Image = "card-replays-es"
End If

If BumperCap1K Then
    Primitive002.Image = "bumper-playmatic-1000"
End If

If PlayfieldRed Then
    Table1.Image = "pf-alt"
End If


'=====================   END OF PLAYER OPTION  ===========================

'If using Visual PinMAME (VPM), place the ROM/game name in the constant below,
'both for VPM, and DOF to load the right DOF config from the Configtool, whether it's a VPM or an Original table

Dim bsTrough, bsTopSaucer, dtbank, dttbank, Flipperactive, x

Const cGameName = "blkfever"

Dim VarHidden:VarHidden = Table1.ShowDT
if B2SOn = true then VarHidden = 1

'Dim x

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
Const SSolenoidOn = "fx_SolenoidOn"
Const SSolenoidOff = "fx_SolenoidOff"
Const SCoin = "fx_Coin"

'*********
'Solenoids
'*********

SolCallBack(1) = "SolSaucers"
SolCallback(2) = "dtbank.SolDropUp"                    'Reset Top Target Bank
SolCallback(4) = "dttbank.SolDropUp"                    'Reset Target Bank
SolCallback(5) = "SolTrough"
SolCallback(6) = "vpmSolSound SoundFX(""fx_Knocker-music"",DOFKnocker),"
SolCallback(8) = "SolFlipperEnabled"              'Unstable
'SolCallback(9) = ""              'Unknown
SolCallback(10) = "SolRandomMusic"              'Unknown
'SolCallback(11) = ""             'Unknown
'SolCallback(12) = ""             'Unknown
'SolCallback(13) = ""             'Unknown Unstable
'SolCallback(14) = ""             'Unknown Unstable
'SolCallback(15) = ""             'Unknown Unstable
'SolCallback(16) = ""             'Unknown Unstable


Sub SolTrough(Enabled)
    If Enabled Then
        bsTrough.ExitSol_On
    End If
End Sub


 Sub SolSaucers(Enabled)
  If Enabled Then
    If bsTopSaucer.Balls Then bsTopSaucer.ExitSol_On
  End If
 End Sub


Dim GameOn
GameOn = FALSE

Sub SolFlipperEnabled(Enabled)
  vpmNudge.SolGameOn Enabled
  GameOn = Enabled
End Sub

Sub SolRandomMusic(Enabled)
  If Enabled Then
'   If GameOn Then
      If Controller.Solenoid(8)=False Then
        If MusicAttract = True Then Random_Music:End If
      End If
'   End If
  End If
End Sub

Sub Random_Music
    dim b
    b = RndNbr(7)
'   If B2SOn then
' Controller.B2SSetData 17, 0
' Controller.B2SSetData 18, 0
' Controller.B2SSetData 19, 0
' End If
    Select Case b
        Case 1
      vpmtimer.addtimer 50, "PlaySound ""bf_m_rock"" '"
      If B2SOn then Controller.B2SSetData 17, 1 End If
        case 2
      vpmtimer.addtimer 50, "PlaySound ""bf_m_bass"" '"
      If B2SOn then Controller.B2SSetData 18, 1 End If
        case 3
      vpmtimer.addtimer 50, "PlaySound ""bf_m_disco"" '"
      If B2SOn then Controller.B2SSetData 19, 1 End If
        case 4
      vpmtimer.addtimer 50, "PlaySound ""bf_m_play-with-me"" '"
      If B2SOn then Controller.B2SSetData 20, 1 End If
    case 5
      vpmtimer.addtimer 50, "PlaySound ""bf_m-come-on-and-dance-with-me"" '"
      If B2SOn then Controller.B2SSetData 19, 1 End If
    case 6
      vpmtimer.addtimer 50, "PlaySound ""bf_m_another-time-baby"" '"
      If B2SOn then Controller.B2SSetData 19, 1 End If
    case 7
      vpmtimer.addtimer 50, "PlaySound ""bf_m_i-miss-you-baby"" '"
      If B2SOn then Controller.B2SSetData 19, 1 End If
    End Select
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
        TRightFlipper.RotateToEnd
        RightFlipperOn = 1
    Else
        PlaySoundAt SoundFX("fx_flipperdown", DOFFlippers), RightFlipper
        RightFlipper.RotateToStart
        TRightFlipper.RotateToStart
        RightFlipperOn = 0
    End If
End Sub

'fin A

'************
' Table init.
'************

Sub Table1_Init
    vpmInit me
    With Controller
        .GameName = cGameName
        If Err Then MsgBox "Can't start Game" & cGameName & vbNewLine & Err.Description:Exit Sub
        .SplashInfoLine = "Black Fever - Playmatic 1980" & vbNewLine & "VPX table by pedator 1.0.0"
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
    vpmNudge.TiltObj = Array(Bumper5)

    ' Trough
    Set bsTrough = New cvpmBallStack
    With bsTrough
    .InitNoTrough BallRelease,5,90,18
'        .InitSw 0, 5, 0, 0, 0, 0, 0, 0
'        .InitKick BallRelease, 90, 18
        .InitEntrySnd "fx_Solenoid", "fx_Solenoid"
        .InitExitSnd SoundFX("fx_ballrelease", DOFContactors), SoundFX("fx_Solenoid", DOFContactors)
        .Balls = 1
    End With

    'Kicker Holes
    Set bsTopSaucer = New cvpmBallStack
    bsTopSaucer.InitSaucer S44, 44, 227, 15
    bsTopSaucer.InitExitSnd SoundFX("fx_kicker", DOFContactors), SoundFX("fx_Solenoid", DOFContactors)
    bsTopSaucer.KickForceVar = 3


    ' Drop targets
    set dtbank = new cvpmdroptarget
    dtbank.InitDrop Array(S24, S25, S26, S28), Array(24, 25, 26, 28)
    dtbank.initsnd SoundFX("fx_droptarget", DOFDropTargets), SoundFX("fx_resetdrop", DOFContactors)
    dtbank.CreateEvents "dtbank"

    set dttbank = new cvpmdroptarget
    dttbank.InitDrop Array(S13, S15, S16, S17, S18), Array(13, 15, 16, 17, 18)
    dttbank.initsnd SoundFX("fx_droptarget", DOFDropTargets), SoundFX("fx_resetdrop", DOFContactors)
    dttbank.CreateEvents "dttbank"

    ' Main Timer init
    PinMAMETimer.Interval = PinMAMEInterval
    PinMAMETimer.Enabled = 1
    RealTime.Enabled = 1

    LoadLUT

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

Sub GateF_Animate: GateP.rotz = 22 - GateF.CurrentAngle: End Sub

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
        TRightFlipper.RotateToEnd
        RightFlipperOn = 1
      End If
      Exit Sub
    End If
    End If
'fin B
    If keycode = LeftTiltKey Then Nudge 90, 5:PlaySound "fx_nudge", 0, 1, -0.1, 0.25
    If keycode = RightTiltKey Then Nudge 270, 5:PlaySound "fx_nudge", 0, 1, -0.1, 0.25
    If keycode = CenterTiltKey Then Nudge 0, 6:PlaySound "bf_v-damn-damn-damn", 0, 1, -0.1, 0.25
    If keycode = LeftMagnaSave Then bLutActive = True:SetLUTLine "Color LUT image " & table1.ColorGradeImage
    If keycode = RightMagnaSave AND bLutActive Then NextLUT:End If
    If KeyCode = KeyInsertCoin1 Then vpmTimer.PulseSw 1
    If KeyCode = KeyInsertCoin2 Then vpmTimer.PulseSw 2
    If KeyCode = KeyInsertCoin3 Then vpmTimer.PulseSw 3
    If KeyCode = keyFront Then vpmTimer.PulseSw 59
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
        TRightFlipper.RotateToStart
        RightFlipperOn = 0
        PlaySoundAt SoundFX("fx_flipperdown", DOFFlippers), RightFlipper
        Controller.Switch(56) = 0
        Controller.Switch(102) = 0
      End If
      Exit Sub
    End If
    End If
'fin B
    If keycode = LeftMagnaSave Then bLutActive = False:HideLUT
    If vpmKeyUp(keycode) Then Exit Sub
    If keycode = PlungerKey Then PlaySoundAt "fx_plunger", Plunger:Plunger.Fire
End Sub


' flippers top animations

Sub LeftFlipper_Animate: LeftFlipperTop.RotZ = LeftFlipper.CurrentAngle: End Sub
Sub RightFlipper_Animate: RightFlipperTop.RotZ = RightFlipper.CurrentAngle: End Sub
Sub TRightFlipper_Animate: TRightFlipperTop.RotZ = TRightFlipper.CurrentAngle: End Sub


' flippers hit Sound

Sub LeftFlipper_Collide(parm)
    PlaySound "fx_rubber_flipper", 0, Vol(ActiveBall), AudioPan(ActiveBall), 0.2, 0, 0, 0, AudioFade(ActiveBall)
End Sub

Sub RightFlipper_Collide(parm)
    PlaySound "fx_rubber_flipper", 0, Vol(ActiveBall), AudioPan(ActiveBall), 0.2, 0, 0, 0, AudioFade(ActiveBall)
End Sub

Sub TRightFlipper_Collide(parm)
    PlaySound "fx_rubber_flipper", 0, Vol(ActiveBall), AudioPan(ActiveBall), 0.2, 0, 0, 0, AudioFade(ActiveBall)
End Sub

'inicio A
'Lights Callback

Set LampCallback = GetRef("UpdateMultipleLamps")

Dim OldState1
'Dim OldState7

OldState1 = 0
'OldState7 = 0

Sub UpdateMultipleLamps

' Enciende luz con falta
    If Controller.Switch(7) = True Then
    LightTilt.state = 1
    End If
' Desactiva objetos con luz de falta
         If LightTilt.state = 1 Then
            VpmNudge.SolGameOn False
            Flipperactive = False
            LeftFlipper.RotateToStart
            RightFlipper.RotateToStart
            TRightFlipper.RotateToStart
      Controller.Switch(56)=0
      Controller.Switch(57)=0
      Controller.Switch(102)=0
      Controller.Switch(104)=0
            Bumper5.Threshold = 100
        End If
'   Activa objetos sin luz de falta
         If LightTilt.state = 0 Then
            VpmNudge.SolGameOn True
            Flipperactive = True
           ' activate bumpers and slings in case of tilt
            Bumper5.Threshold = 1
        End If

    If Controller.Lamp(11) <> OldState1 Then 'Game Over light
        If Controller.Lamp(11) Then
            VpmNudge.SolGameOn False
            Flipperactive = False
            LeftFlipper.RotateToStart
            RightFlipper.RotateToStart
            TRightFlipper.RotateToStart
      Controller.Switch(56)=0
      Controller.Switch(57)=0
      Controller.Switch(102)=0
      Controller.Switch(104)=0
            OldState1 = 1
        Else
            VpmNudge.SolGameOn True
            Flipperactive = True
          ' activate bumpers and slings in case of tilt
            Bumper5.Threshold = 1
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
    NfReel 31, modea
    NfReel 27, modeb
  ' Playfield and b2s lights
    Lamp 1, Light2a  '40000 top OK
    Lamp 2, Light2b  ' 30000 OK
    Lamp 3, Light2c  ' 20000 OK
    Lamp 5, Light1a  ' 10000 top OK
'    Lamp 6, Light1b  ' ? used for b2s
'    Lamp 7, Light7 ' ? used for b2s
'    Lamp 8, Light8 ' Not used?
'    Lamp 9, Light9 ' Not used?
    Lamp 10, LB1a ' Bumper second light OK
'    Lamp 11, Light11 'Match / GameOver
'    Lamp 15, Light15 'rhigh
    Lamp 14, LB1  ' Bumper first light OK
    Lamp 17, Light034 ' 30000 when lit OK
    Lamp 18, Light6b  ' 25000 when lit -  2 bonus OK
    Lamp 19, Light6c ' Star left 10000 when lit OK
    Lamp 21, Light5a   ' Star left 10000 when lit OK
    Lamp 22, Light5b   ' Star left 10000 when lit OK
    Lamp 23, Light5c ' Star left 10000 when lit OK
'    Lamp 25, Light25  '? used for b2s
'    Lamp 26, Light26  '? used for b2s
'    Lamp 27, Light8c ' B2S Continuous Play
    Lamp 29, Light7a ' 10000 left outlane OK
    Lamp 30, Light7b  ' 10000 left inlane OK
'    Lamp 31, Light7c  ' B2S Alternate Play ----------- ON
    Lamp 32, Light33  ' 15000 when lit (not ID > manual managing)
    Lamp 33, LightX2 ' X2 OK
    Lamp 34, LightX3 ' X3 OK
    Lamp 35, LightX5 ' X5 OK
  Lamp 36, Light33 ' 15000 when lit
    Lamp 37, Light11b ' Prepare special OK
    Lamp 38, Light38 ' Special Left OK
    Lamp 39, Light39 ' 50000 when lit
    Lamp 41, Light12a  '5000 kicker OK
    Lamp 42, Light12b  '3 Bonus kicker OK
    Lamp 43, Light12c  '30000 kicker OK
    Lamp 45, Light11a  'Prepare special kicker OK
    Lamp 46, Light1c  'Extra Ball OK
    Lamp 49, Light14a  '1 OK
    Lamp 50, Light14b  '2 OK
    Lamp 51, Light14c  '3 OK
    Lamp 53, Light13a  '4 OK
    Lamp 54, Light13b  '5 OK
    Lamp 55, Light13c  '6 OK
    Lamp 57, Light16a  '7 OK
    Lamp 58, Light16b  '8 OK
    Lamp 59, Light16c  '9 OK
    Lamp 61, Light15a  '10000 OK
    Lamp 62, Light15b  '20000 OK
'    Lamp 63, Light63  '? used for b2s

' Luz 15.000 encendida con 2 dianas verdes superiores bajadas y apaga con cualquier X2,X3 o X5 apagado

  If S24.IsDropped And S28.IsDropped Then
    Light33.state = 1
  End If
  If LightX2.state = 0 And LightX3.state = 0 And LightX5.state = 0 Then
    Light33.state = 0
  End If

' Luces animacion b2s apagadas con luz Game over
  If B2SOn Then
    If Controller.Lamp(11) = True Then
      Controller.B2SSetData 17, 0 'animacion rock b2s
      Controller.B2SSetData 18, 0 'animacion bass b2s
      Controller.B2SSetData 19, 0 'animacion disco b2s
      Controller.B2SSetData 20, 0 'animacion dance b2s
    End If
  End If

' Si tiras 4 dianas verdes suena audio
  If S24.IsDropped And S28.IsDropped And S26.IsDropped And S25.IsDropped Then
    vpmtimer.addtimer 50, "PlaySound ""bf_v_pay-me"" '"
  End If

' Si tiras 5 dianas rojas y blanca suena audio al levantarse
  If S13.IsDropped And S15.IsDropped And S16.IsDropped And S17.IsDropped And S18.IsDropped Then
    vpmtimer.addtimer 50, "PlaySound ""bf_v-thats-the-way-i-like"" '"
  End If

' Luz X5 encendida reproduce audio
  If GameOn Then
    If LightX3.state = 1 Then
      vpmtimer.addtimer 50, "PlaySound ""bf_m-playing-black"" '"
    End If
  End If

' Obtiene bola extra -> reproduce audio
  If GameOn Then
    If Light1c.state = 1 Then
      vpmtimer.addtimer 50, "PlaySound ""bf_v-give-it-to-me"" '"
    End If
  End If

    Lampi 70, gi015  'Luces fijas
    Lampi 71, gi014
    Lampi 72, gi013
    Lampi 73, gi012
    Lampi 74, gi011
    Lampi 75, gi001
    Lampi 76, gi002
    Lampi 77, gi003
    Lampi 78, gi004
    Lampi 79, gi005
    Lampi 80, gi006
    Lampi 81, gi007
    Lampi 82, gi008
    Lampi 83, gi009
    Lampi 84, gi010
    Lampi 85, gi016
    Lampi 86, gi017
    Lampi 87, gi018
    Lampi 88, gi019
    Lampi 89, gi020
    Lampi 90, gi021
    Lampi 91, gi022
    Lampi 92, gi001t
    Lampi 93, gi002t
    Lampi 94, gi003t
    Lampi 95, gi004t
    Lampi 96, gi005t
    Lampi 97, gi006t
    Lampi 98, gi007t
    Lampi 99, gi008t
    Lampi 120, gi009t
    Lampi 121, gi010t
    Lampi 122, gi011t
    Lampi 123, gi012t
    Lampi 124, gi013t
    Lampi 125, gi014t
    Lampi 126, gi015t
    Lampi 127, gi016t
    Lampi 128, gi017t
    Lampi 129, gi018t
    Lampi 130, gi019t

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

Sub aMetals_Hit(idx):PlaySoundAtBall "fx_MetalHit":End Sub
Sub aMetalWires_Hit(idx):PlaySoundAtBall "fx_MetalWire":End Sub
Sub aRubber_ShortBands_Hit(idx):PlaySoundAtBall "fx_rubber_band":End Sub
Sub aRubber_LongBands_Hit(idx):PlaySoundAtBall "fx_rubber_longband":End Sub
Sub aRubber_Posts_Hit(idx):PlaySoundAtBall "fx_rubber_post":End Sub
Sub aRubber_Pins_Hit(idx):PlaySoundAtBall "fx_rubber_pin":End Sub
Sub aRubber_Pegs_Hit(idx):PlaySoundAtBall "fx_rubber_peg":End Sub
Sub aPlastics_Hit(idx):PlaySoundAtBall "fx_PlasticHit":End Sub
Sub aGates_Hit(idx):PlaySoundAtBall "gate4":End Sub
Sub aWoods_Hit(idx):PlaySoundAtBall "fx_Woodhit":End Sub
Sub aDropTargets_Hit(idx):PlaySoundAt SoundFX("fx_droptarget", DOFDropTargets), aDropTargets(idx):End Sub
Sub aTargets_Hit(idx):PlaySoundAt SoundFX("fx_target", DOFTargets), aTargets(idx):End Sub
Sub aRollovers_Hit(idx):PlaySoundAt "fx_sensor", aRollovers(idx):End Sub

'*********
' Bumpers
'*********
Sub Bumper5_Hit:vpmTimer.PulseSw 48:PlaySoundAt SoundFXDOF("fx_bumper4", 105, DOFPulse, DOFContactors), Bumper5:End Sub '48 - Bumper

' Scoring rubbers
Sub sw12_Hit:vpmtimer.pulsesw 12:End Sub '12 - 110 Points Switch
Sub sw31_Hit:vpmtimer.pulsesw 31:End Sub '31 - 50 Points Switch
Sub sw32_Hit:vpmtimer.pulsesw 32:End Sub '32 - 10 Points Switch
Sub sw001_Hit:vpmtimer.addtimer 100, "PlaySound ""bf_v-you-give-me-s"" '":End Sub ' Reproduce audio al tocar goma
Sub sw002_Hit:vpmtimer.addtimer 100, "PlaySound ""bf_v-you-give-me-s"" '":End Sub ' Reproduce audio al tocar goma

'*********
' Rollovers
'*********
Sub sw33_Hit:Controller.Switch(33) = 1:PlaySoundAt "fx_sensor", sw33:End Sub '33 - Right Inlane
Sub sw33_UnHit:Controller.Switch(33) = 0:End Sub
Sub sw34_Hit:Controller.Switch(34) = 1:PlaySoundAt "fx_sensor", sw34:vpmtimer.addtimer 200, "PlaySound ""bf_v_get-out-now"" '":End Sub '34 - Right Outlane
Sub sw34_UnHit:Controller.Switch(34) = 0:End Sub
Sub sw35_Hit:Controller.Switch(35) = 1:PlaySoundAt "fx_sensor", sw35:End Sub '35 - Left Inlane / 10000 when lit
Sub sw35_UnHit:Controller.Switch(35) = 0:End Sub
Sub sw36_Hit:Controller.Switch(36) = 1:PlaySoundAt "fx_sensor", sw36:End Sub '36 - Left Inlane / 10000 when lit
Sub sw36_UnHit:Controller.Switch(36) = 0:End Sub
Sub sw38_Hit:Controller.Switch(38) = 1:PlaySoundAt "fx_sensor", sw38:vpmtimer.addtimer 200, "PlaySound ""bf_v_get-out-now"" '":End Sub '38 - Left Outlane
Sub sw38_UnHit:Controller.Switch(38) = 0:End Sub

Sub sw11_Hit:Controller.Switch(11) = 1:PlaySoundAt "fx_sensor", sw11:End Sub '11 - Top Right Line
Sub sw11_UnHit:Controller.Switch(11) = 0:End Sub
Sub sw41_Hit:Controller.Switch(41) = 1:PlaySoundAt "fx_sensor", sw41:End Sub '41 - Right Green Rollover / 30000 when lit 2 bonus
Sub sw41_UnHit:Controller.Switch(41) = 0:End Sub
Sub sw42_Hit:Controller.Switch(42) = 1:PlaySoundAt "fx_sensor", sw42:End Sub '42 - Left Green Rollover / 25000 when lit
Sub sw42_UnHit:Controller.Switch(42) = 0:End Sub
Sub sw43_Hit:Controller.Switch(43) = 1:PlaySoundAt "fx_sensor", sw43:End Sub '43 - Top Left Red Rollover /10000 add bonus
Sub sw43_UnHit:Controller.Switch(43) = 0:End Sub
Sub sw45_Hit:Controller.Switch(45) = 1:PlaySoundAt "fx_sensor", sw45:End Sub '45 - Top Left Red Rollover /10000 add bonus
Sub sw45_UnHit:Controller.Switch(45) = 0:End Sub
Sub sw46_Hit:Controller.Switch(46) = 1:PlaySoundAt "fx_sensor", sw46:End Sub '46 - Top Left Red Rollover /10000 add bonus
Sub sw46_UnHit:Controller.Switch(46) = 0:End Sub
Sub sw47_Hit:Controller.Switch(47) = 1:PlaySoundAt "fx_sensor", sw47:End Sub '47 - Top Left Red Rollover /10000 add bonus
Sub sw47_UnHit:Controller.Switch(47) = 0:End Sub

Sub sw009_Hit:Controller.Switch(7) = 1:PlaySoundAt "fx_sensor", sw009:End Sub ' Falta Tilt
Sub sw009_UnHit:Controller.Switch(7) = 0:End Sub

'*********
' Targets
'*********

Sub sw21_Hit:vpmTimer.PulseSw 21:PlaySoundAt "fx_target", sw21:End Sub '21 - Target Red Top right
Sub sw22_Hit:vpmTimer.PulseSw 22:PlaySoundAt "fx_target", sw22:End Sub '22 - Target Red Top Middle
Sub sw23_Hit:vpmTimer.PulseSw 23:PlaySoundAt "fx_target", sw23:End Sub '23 - Target Red Top left
Sub sw14_Hit:vpmTimer.PulseSw 14:PlaySoundAt "fx_target", sw14:End Sub '14 - Back target bank top


'*********
' Drain & holes
'*********
Sub S5_Hit:PlaysoundAt "fx_drain", S5:bsTrough.AddBall Me:End Sub
'Sub S31_Hit:PlaysoundAt "fx_kicker_enter", S31:bsSaucer31.AddBall 0:End Sub '31 Left Kicker
Sub S44_Hit:PlaySoundAt SoundFX("fx_kicker_enter", DOFContactors), S44:bsTopSaucer.AddBall 0:End Sub '44 Right Kicker
Sub Gate1_Hit:Random_Music:LightTilt.state = 0:End Sub 'Musica al sacar bola y apaga luz de falta manual para metodo flipperfix

'**********************
' Ball Collision Sound
'**********************

Sub OnBallBallCollision(ball1, ball2, velocity)
    PlaySound("fx_collide"), 0, Csng(velocity) ^2 / 2000, Pan(ball1), 0, Pitch(ball1), 0, 0, AudioFade(ball1)
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
Digits(26) = Array(d17, d19, d22, d23, d21, d18, d20)

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

'************************************
'       LUT - Darkness control
' 10 normal level & 10 warmer levels
'************************************

Dim bLutActive, LUTImage

Sub LoadLUT
    bLutActive = False
    x = LoadValue(cGameName, "LUTImage")
    If(x <> "")Then LUTImage = x Else LUTImage = 0
    UpdateLUT
End Sub

Sub SaveLUT
    SaveValue cGameName, "LUTImage", LUTImage
End Sub

Sub NextLUT:LUTImage = (LUTImage + 1)MOD 22:UpdateLUT:SaveLUT:SetLUTLine "Color LUT image " & table1.ColorGradeImage:End Sub

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

Dim GiIntensity
GiIntensity = 1   'can be used by the LUT changing to increase the GI lights when the table is darker

Sub ChangeGiIntensity(factor) 'changes the intensity scale
    Dim bulb
    For each bulb in aGiLights
        bulb.IntensityScale = GiIntensity * factor
    Next
End Sub

' New LUT postit
Function GetHSChar(String, Index)
    Dim ThisChar
    Dim FileName
    ThisChar = Mid(String, Index, 1)
    FileName = "PostIt"
    If ThisChar = " " or ThisChar = "" then
        FileName = FileName & "BL"
    ElseIf ThisChar = "<" then
        FileName = FileName & "LT"
    ElseIf ThisChar = "_" then
        FileName = FileName & "SP"
    Else
        FileName = FileName & ThisChar
    End If
    GetHSChar = FileName
End Function

Sub SetLUTLine(String)
    Dim Index
    Dim xFor
    Index = 1
    LUBack.imagea="PostItNote"
    For xFor = 1 to 40
        Eval("LU" &xFor).imageA = GetHSChar(String, Index)
        Index = Index + 1
    Next
End Sub

Sub HideLUT
    SetLUTLine ""
    LUBack.imagea="PostitBL"
End Sub


'DIPSWITCHES

' Press Test key (7), pressing test advance the options (01 to 26). Press Start key (1) for change values for selected option.
' Presiona tecla de Test (7), presionando la tecla de test cambias el tipo de opción (01 a 26). Presiona tecla Start (1) para cambiar valores de la opción.

' 01 - Número maximo de creditos (00 to 99). Default: 15
' 02 - Puntuación HIGH SCORE (000 to 990). Default: 700.000
' 03 - Premio tanteo 1º (000 to 990). Default: 500.000
' 04 - Premio tanteo 2º (000 to 990). Default: 650.000
' 05 - Premio tanteo 3º (000 to 990). Default: 000.000
' 06 - Partidas por moneda 1º selector (0,3-0,5-1-1,5 to 39). Default: 0,3
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
' 22 - Puntuación dianas verdes (0=1000, 1=10000). Default: 0
' 23 - Puntuación dianas rojas redondas (0=1000, 1=5000). Default: 0
' 24 - Pasillo P1 superior (0=500, 1=5000). Default: 0
' 25 - Pasillo P3, P4 - inlanes izquierdos (0=2000, 1=500). Default: 0
' 26 - No utilizada

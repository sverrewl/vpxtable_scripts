'Petaco / IPD No. 4883 / 1984 / 4 Players
'Juegos Populares, S.A., of Madrid, Spain (1986)
'
'VPX conversion by Wiesshund, VP9 by MRCMRC
'VPX8, version 6.0.0, by pedator & jpsalas

Option Explicit
Randomize

Const BallSize = 50
Const BallMass = 1

'*************** USER OPTIONS ************************

'If playing with 5 balls you may want to enable the next line to change tha apron to 5 Balls
'Apron.Image = "apron-hd5"

'************ END OF USER OPTIONS ********************

Const cGameName = "petaco"
Const cCredits = "Petaco - Juegos Populares 1984"

Const UseSolenoids = 2
Const UseLamps = 1
Const UseGI = 0

' Standard Sounds
Const SSolenoidOn = "fx_SolenoidOn"
Const SSolenoidOff = "fx_SolenoidOff"
Const SCoin = "fx_Coin"

Dim VarHidden, x
If Petaco.ShowDT = true then
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

On Error Resume Next
ExecuteGlobal GetTextFile("controller.vbs")
If Err Then MsgBox "You need the Controller.vbs file in order to run this table (installed with the VPX package in the scripts folder)"
On Error Goto 0

if B2SOn = true then VarHidden = 1

LoadVPM "01520000", "juegos2.vbs", 3.1

Dim bsTrough, dtIzquierda, dtDerecha, cbIzquierda, cbDerecha
Dim PlayOn, obj

' ****************
' Inicializa tabla
' ****************

Sub Petaco_Init

    vpmMapLights AllLights

    With Controller
        .GameName = cGameName
        If Err Then MsgBox "Can't start Game" & cGameName & vbNewLine & Err.Description:Exit Sub
        .SplashInfoLine = cCredits
        .HandleMechanics = 0
        .HandleKeyboard = 0
        .ShowDMDOnly = 1
        .ShowFrame = 0
        .ShowTitle = 0
        .Hidden = VarHidden
        .Run GetPlayerHwnd
        If Err Then MsgBox Err.Description
    End With

    PinMAMETimer.Interval = PinMAMEInterval
    PinMAMETimer.Enabled = 1

    vpmNudge.TiltSwitch = 39 ' Cuando se produce falta la ROM parece colgarse, en lugar de desconectar
    vpmNudge.Sensitivity = 4 ' el solenoide de Game On (12), con lo cual no se puede seguir jugando.
    vpmNudge.TiltObj = Array(sw7, sw15, Leftslingshot, Rightslingshot)

    ' ===== Switches de opciones =====

    ' El manual que circula por Internet en inglés tiene más opciones que el manual que circula en español,
    ' lo que hace pensar en varias versiones de la placa. No todas las opciones funcionan con la ROM "pública".

    ' DIP0.SW1  ¿Reset High Score (OFF = automático , ON = manual)?
    ' DIP0.SW2  Visualiza partidas y bolas extras (OFF = no , ON = sí)
    ' DIP0.SW3  Visualiza monederos (OFF = no , ON = sí)
    ' DIP0.SW4  Visualiza horas funcionamiento (OFF = no , ON = sí)
    ' DIP0.SW5  Lotería (OFF = sí , ON = no)
    ' DIP0.SW6  Reclamo (OFF = sí , ON = no)
    ' DIP0.SW7  ¿Falta?
    ' DIP0.SW8  ¿Free play mode (OFF = free play , ON = ticket)?

    ' DIP1.SW1  Premios por puntos 1 (OFF/OFF = bajo , OFF/ON = alto , ON/OFF = medio , ON/ON = máximo)
    ' DIP1.SW2  Premios por puntos 2
    ' DIP1.SW3  Bolas por partida (OFF = 5 , ON = 3)
    ' DIP1.SW4  Partidas por moneda 1 (OFF/OFF = S1:2-4-6 , OFF/ON = S3:1-1-3 , ON/OFF = S2:1-3-5 , ON/ON = S4:1-3-6)
    ' DIP1.SW5  Partidas por moneda 2
    ' DIP1.SW6  Bola extra por puntos (OFF = sí , ON = no)
    ' DIP1.SW7  ¿Power-fail?
    ' DIP1.SW8  ¿Pulsar partida?

    ' Configuración manual. Descomentar las líneas, y activar switches con 1 o 0: SW1*1 + SW2*2 + SW3*4 + ...

    'Controller.Dip(0)=(0*1 + 0*2 + 0*4 + 0*8 + 0*16 + 0*32 + 0*64 + 0*128)
    'Controller.Dip(1)=(0*1 + 0*2 + 0*4 + 0*8 + 0*16 + 0*32 + 0*64 + 1*128)

    Set vpmShowDips = GetRef("EditaSwitches")

    ' ===== Físicas iniciales de bumpers y flippers =====

    ' ===== Declara stacks, bancadas y bolas cautivas =====

    Set bsTrough = New cvpmBallStack
    bsTrough.InitNoTrough SalidaBola, 1, 107, 15
    bsTrough.InitExitSnd "salida", "solenoid"

    Set dtIzquierda = New cvpmDropTarget
    dtIzquierda.InitDrop Array(sw13, sw21, sw29), Array(13, 21, 29)
    dtIzquierda.InitSnd " ", "FlapOpen"

    Set dtDerecha = New cvpmDropTarget
    dtDerecha.InitDrop Array(sw14, sw22, sw30), Array(14, 22, 30)
    dtDerecha.InitSnd " ", "FlapOpen"

    Set cbIzquierda = New cvpmCaptiveBall
    cbIzquierda.InitCaptive CaptiveTriggerIzq, CaptiveWallIzq, CaptiveKickerIzq, 331
    cbIzquierda.Start
    cbIzquierda.ForceTrans = 1.5
    cbIzquierda.MinForce = 5
    cbIzquierda.CreateEvents "cbIzquierda"

    Set cbDerecha = New cvpmCaptiveBall
    cbDerecha.InitCaptive CaptiveTriggerDer, CaptiveWallDer, CaptiveKickerDer, 16.5
    cbDerecha.Start
    cbDerecha.ForceTrans = 1.5
    cbDerecha.MinForce = 5
    cbDerecha.CreateEvents "cbDerecha"

    GameOn false
End Sub

Sub Petaco_Paused:Controller.Pause = True:End Sub
Sub Petaco_UnPaused:Controller.Pause = False:End Sub

' *********
' Controles
' *********

Sub Petaco_KeyDown(ByVal KeyCode)
    If keycode = LeftTiltKey Then Nudge 90, 5:PlaySound SoundFX("fx_nudge", 0), 0, 1, -0.1, 0.25
    If keycode = RightTiltKey Then Nudge 270, 5:PlaySound SoundFX("fx_nudge", 0), 0, 1, 0.1, 0.25
    If keycode = CenterTiltKey Then Nudge 0, 6:PlaySound SoundFX("fx_nudge", 0), 0, 1, 0, 0.25
    If vpmKeyDown(KeyCode) Then Exit Sub
    If Keycode = PlungerKey Then:PlaySoundAt "fx_PlungerPull", Plunger:Plunger.Pullback:End If
End Sub

Dim loaded
Sub Chamber_Hit:loaded = 1:End Sub

Sub Chamber_UnHit:Loaded = 0:End Sub

Sub Petaco_KeyUp(ByVal KeyCode)
    If Keycode = PlungerKey Then
        If loaded then
            PlaySoundAt "fx_plunger", Plunger:Plunger.Fire
        Else
            PlaySoundAt "fx_plunger_empty", Plunger:Plunger.Fire
        End If
    End If

    If vpmKeyUp(KeyCode) Then Exit Sub
End Sub

'*******************
'  Flipper Subs
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


Sub LeftFlipper_Collide(parm)
    PlaySound "fx_rubber_flipper", 0, parm / 60, pan(ActiveBall), 0.2, 0, 0, 0, AudioFade(ActiveBall)
End Sub

Sub RightFlipper_Collide(parm)
    PlaySound "fx_rubber_flipper", 0, parm / 60, pan(ActiveBall), 0.2, 0, 0, 0, AudioFade(ActiveBall)
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

LiveCatchSensivity = 15

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

' **********
' Solenoides
' **********

'SolCallback(1)="vpmSolSound ""bumper1"","
'SolCallback(2)="vpmSolSound ""bumper2"","
'SolCallback(3)="vpmSolSound ""slingshot"","
'SolCallback(4)="vpmSolSound ""slingshot"","
SolCallback(5) = "dtIzquierda.SolDropUp"
SolCallback(6) = "dtDerecha.SolDropUp"
SolCallback(7) = "vpmSolSound ""knocker"","
SolCallback(8) = "bsTrough.SolOut"
SolCallback(11) = "s91s"
SolCallback(12) = "GameOn"
SolCallback(13) = "S13s"
SolCallback(14) = "s14s"
SolCallback(15) = "s15s"
SolCallback(16) = "s16s"

Sub s16s(Enabled)
    If Enabled Then
        s16.state = 1
    Else
        s16.state = 0
    End if
End Sub

Sub S14s(Enabled)
    If Enabled Then
        s14a.state = 1
        s14b.state = 1
    Else
        s14a.state = 0
        s14b.state = 0
    End If
End Sub

Sub S15s(Enabled)
    If Enabled Then
        s15a.state = 1
        s15b.state = 1
        s15c.state = 1
        s15d.state = 1
    Else
        s15a.state = 0
        s15b.state = 0
        s15c.state = 0
        s15d.state = 0
    End If
End Sub

Sub S13s(Enabled)
    If Enabled Then
        s13.state = 1
    Else
        s13.state = 0
    End If
End Sub

Sub S91s(Enabled)
    If Enabled Then
        s91.state = 2
    Else
        s91.state = 0
    End If
End Sub

Set LampCallback = GetRef("UpdateMultipleLamps")
Dim NewVal, OldVal
NewVal = 0:OldVal = 0

' What is this?
' I dunno, some kind of analog multiplexor?
' Based on pulses from 4 lamp outputs
' who cares, it works
Sub UpdateMultipleLamps
    NewVal = ABS(Controller.Lamp(1) ) + ABS(Controller.Lamp(2) ) * 2 + ABS(Controller.Lamp(3) ) * 4 + ABS(Controller.Lamp(4) ) * 8
    If NewVal <> OldVal Then
        Select Case OldVal
            Case 0:
            Case 1:MP1.State = 0:MP7.State = 0
            Case 2:MP2.State = 0:MP8.State = 0
            Case 3:MP3.State = 0:MP9.State = 0
            Case 4:MP4.State = 0:MP10.State = 0
            Case 5:MP5.State = 0:MP11.State = 0
            Case 6:MP6.State = 0:MP12.State = 0
            Case 7:MP1a.State = 0
            Case 8:MP2a.State = 0
            Case 9:MP3a.State = 0
            Case 10:MP4a.State = 0
            Case 11:MP5a.State = 0
            Case 12:MP6a.State = 0
            Case 13:MP13.State = 0
            Case 14:MP14.State = 0
            Case 15:
            Case 16:
        End Select
        Select Case NewVal
            Case 0:
            Case 1:MP1.State = 1:MP7.State = 1
            Case 2:MP2.State = 1:mp8.State = 1
            Case 3:MP3.State = 1:mp9.State = 1
            Case 4:MP4.State = 1:mp10.State = 1
            Case 5:MP5.State = 1:MP11.State = 1
            Case 6:MP6.State = 1:MP12.State = 1
            Case 7:MP1a.State = 1
            Case 8:MP2a.State = 1
            Case 9:MP3a.State = 1
            Case 10:MP4a.State = 1
            Case 11:MP5a.State = 1
            Case 12:MP6a.State = 1
            Case 13:MP13.State = 1
            Case 14:MP14.State = 1
            Case 15:
            Case 16:
        End Select
        OldVal = NewVal
    End If
End Sub

' ********
' Switches
' ********

' Drain

Sub sw1_Hit:PlaySoundAt "fx_drain", sw1:bsTrough.AddBall Me:End Sub

' Rollovers

Sub sw11a_Hit:VPMTimer.PulseSW 11:PlaySoundAt "sensor", sw11a:End Sub
Sub sw11b_Hit:VPMTimer.PulseSW 11:PlaySoundAt "sensor", sw11b:End Sub
Sub sw12_Hit:VPMTimer.PulseSW 12:PlaySoundAt "sensor", sw12:End Sub ' Hay una errata en el manual. Los rollovers 12 y 24 están intercambiados
Sub sw18_Hit:VPMTimer.PulseSW 18:PlaySoundAt "sensor", sw18:End Sub
Sub sw19_Hit:VPMTimer.PulseSW 19:PlaySoundAt "sensor", sw19:End Sub
Sub sw24_Hit:VPMTimer.PulseSW 24:PlaySoundAt "sensor", sw24:End Sub
Sub sw25_Hit:VPMTimer.PulseSW 25:PlaySoundAt "sensor", sw25:End Sub
Sub sw26a_Hit:VPMTimer.PulseSW 26:PlaySoundAt "sensor", sw26a:End Sub
Sub sw26b_Hit:VPMTimer.PulseSW 26:PlaySoundAt "sensor", sw26b:End Sub
Sub sw28_Hit:VPMTimer.PulseSW 28:PlaySoundAt "sensor", sw28:End Sub
Sub sw32_Hit:VPMTimer.PulseSW 32:PlaySoundAt "sensor", sw32:End Sub

' Bancada dianas izquierda

Sub sw13_Hit:PlaySoundAt SoundFX("flapclos", DOFDropTargets), sw13:dtIzquierda.Hit 1:End Sub
Sub sw21_Hit:PlaySoundAt SoundFX("flapclos", DOFDropTargets), sw21::dtIzquierda.Hit 2:End Sub
Sub sw29_Hit:PlaySoundAt SoundFX("flapclos", DOFDropTargets), sw29::dtIzquierda.Hit 3:End Sub

' Bancada dianas derecha

Sub sw14_Hit:PlaySoundAt SoundFX("flapclos", DOFDropTargets), sw14::dtDerecha.Hit 1:End Sub
Sub sw22_Hit:PlaySoundAt SoundFX("flapclos", DOFDropTargets), sw22::dtDerecha.Hit 2:End Sub
Sub sw30_Hit:PlaySoundAt SoundFX("flapclos", DOFDropTargets), sw30::dtDerecha.Hit 3:End Sub

' Veleta

Sub sw10_Spin:vpmTimer.PulseSw 10:PlaySoundAt "spinner", sw10:End Sub

' Gomas

Sub sw27a_Hit:vpmTimer.PulseSw 27:End Sub
Sub sw27b_Hit:vpmTimer.PulseSw 27:End Sub
Sub sw27c_Hit:vpmTimer.PulseSw 27:End Sub ' En estas dos gomas hay dos switches, pero en el manual no se
Sub sw27d_Hit:vpmTimer.PulseSw 27:End Sub ' indica a qué switch corresponden. Asumimos que al 27 también

' Dianas

Sub sw8a_Hit:vpmTimer.PulseSw 8:PlaySoundAt SoundFX("Target", DOFTargets), sw8a:End Sub

Sub sw16a_Hit:vpmTimer.PulseSw 16:PlaySoundAt SoundFX("Target", DOFTargets), sw16a:End Sub

' Bumpers

Sub sw7_Hit
    If PlayOn Then
        PlaySoundAt SoundFX("Bumper1", DOFContactors), sw7
        vpmTimer.PulseSw 7
    End If
End Sub

Sub sw15_Hit
    PlaySoundAt SoundFX("Bumper2", DOFContactors), sw15
    If PlayOn Then
        vpmTimer.PulseSw 15
    End If
End Sub

' Slings
Dim LStep, RStep

Sub LeftSlingShot_Slingshot
    PlaySoundAt SoundFX("fx_slingshot", DOFContactors), Lemk
    'DOF 101, DOFPulse
    LeftSling004.Visible = 1
    Lemk.RotX = 26
    LStep = 0
    vpmTimer.PulseSw 23
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
    vpmTimer.PulseSw 31
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

' **************************************
' Iluminación general y bobina de GameOn
' **************************************

Sub GameOn(Enabled)
    If Enabled Then
        PlayOn = true
        For Each obj In GILights
            obj.state = 1
        Next
    Else
        PlayOn = false
        For Each obj In GILights
            obj.state = 0
        Next
    End If
End Sub

' ********************
' Switches de opciones
' ********************

Sub EditaSwitches
    Dim vpmDips
    Set vpmDips = New cvpmDips
    With vpmDips
        .AddForm 700, 400, "Petaco - Switches de opciones"
        .AddFrame 2, 2, 190, "Audit", 14, Array("Off", 0, "Games", 2, "Purses", 4, "Hours", 8)
        .AddFrame 2, 83, 190, "Match Feature", 16, Array("Yes", 0, "No", 16)
        .AddFrame 2, 137, 190, "Claim", 32, Array("Yes", 0, "No", 32)
        .AddFrame 2, 191, 190, "Match probability", 768, Array("Low", 0, "Medium", 256, "High", 512, "Maximum", 768)
        .AddFrame 205, 2, 190, "Coins per Credit", 6144, Array("S1: 2 - 4 - 6", 0, "S2: 1 - 3 - 5", 2048, "S3: 1 - 1 - 3", 4096, "S4: 1 - 3 - 6", 6144)
        .AddFrame 205, 83, 190, "Balls per play", 1024, Array("3 Ball", 1024, "5 Ball", 0)
        .AddFrame 205, 137, 190, "Extra ball for points", 8192, Array("Yes", 0, "No", 8192)
        .AddLabel 2, 272, 380, 20, "For the changes to take effect, after pressing OK press F3 to restart the ROM"
        .ViewDips
    End With
End Sub

' *******
' Sonidos
' *******

Sub CaptiveKickerIzq_UnHit:PlaySound "collide3":Me.TimerEnabled = 1:CaptiveKickerIzq.Enabled = 0:End Sub
Sub CaptiveKickerIzq_Timer:CaptiveKickerIzq.Enabled = 1:Me.TimerEnabled = 0:End Sub
Sub CaptiveKickerDer_UnHit:PlaySound "collide3":Me.TimerEnabled = 1:CaptiveKickerDer.Enabled = 0:End Sub
Sub CaptiveKickerDer_Timer:CaptiveKickerDer.Enabled = 1:Me.TimerEnabled = 0:End Sub

'*********************
'LED's based on Eala's
'*********************
Dim Digits(32)

Digits(0) = Array(a10, a12, a15, a16, a14, a11, a13)
Digits(1) = Array(a20, a22, a25, a26, a24, a21, a23)
Digits(2) = Array(a30, a32, a35, a36, a34, a31, a33, a37)
Digits(3) = Array(a40, a42, a45, a46, a44, a41, a43)
Digits(4) = Array(a50, a52, a55, a56, a54, a51, a53)
Digits(5) = Array(a60, a62, a65, a66, a64, a61, a63)

Digits(6) = Array(b10, b12, b15, b16, b14, b11, b13)
Digits(7) = Array(b20, b22, b25, b26, b24, b21, b23)
Digits(8) = Array(b30, b32, b35, b36, b34, b31, b33, b37)
Digits(9) = Array(b40, b42, b45, b46, b44, b41, b43)
Digits(10) = Array(b50, b52, b55, b56, b54, b51, b53)
Digits(11) = Array(b60, b62, b65, b66, b64, b61, b63)

Digits(12) = Array(e10, e12, e15, e16, e14, e11, e13)
Digits(13) = Array(e20, e22, e25, e26, e24, e21, e23)
Digits(14) = Array(e30, e32, e35, e36, e34, e31, e33, e37)
Digits(15) = Array(e40, e42, e45, e46, e44, e41, e43)
Digits(16) = Array(e50, e52, e55, e56, e54, e51, e53)
Digits(17) = Array(e60, e62, e65, e66, e64, e61, e63)

Digits(18) = Array(f10, f12, f15, f16, f14, f11, f13)
Digits(19) = Array(f20, f22, f25, f26, f24, f21, f23)
Digits(20) = Array(f30, f32, f35, f36, f34, f31, f33, f37)
Digits(21) = Array(f40, f42, f45, f46, f44, f41, f43)
Digits(22) = Array(f50, f52, f55, f56, f54, f51, f53)
Digits(23) = Array(f60, f62, f65, f66, f64, f61, f63)

Digits(24) = Array(c00, c02, c05, c06, c04, c01, c03)
Digits(25) = Array(c10, c12, c15, c16, c14, c11, c13)

Digits(26) = Array(d00, d02, d05, d06, d04, d01, d03)
Digits(27) = Array(d10, d12, d15, d16, d14, d11, d13)

'********************
'Update LED's display
'********************

Sub Leds_Timer()
    Dim ChgLED, ii, num, chg, stat, obj
    ChgLED = Controller.ChangedLEDs(&Hffffffff, &Hffffffff)
    If Not IsEmpty(ChgLED) Then
        For ii = 0 To UBound(chgLED)
            num = chgLED(ii, 0):chg = chgLED(ii, 1):stat = chgLED(ii, 2)
            For Each obj In Digits(num)
                If chg And 1 Then obj.State = stat And 1
                chg = chg \ 2:stat = stat \ 2
            Next
        Next
    End If
End Sub

' SOUND ENGINE Update
'************************************
' Hit Sounds
'************************************

Sub aMetals_Hit(idx):PlaySoundAtBall "fx_MetalHit":End Sub
Sub aWires_Hit(idx):PlaySoundAtBall "fx_Metalwire":End Sub
Sub aRubber_Bands_Hit(idx):PlaySoundAtBall "fx_rubber_band":End Sub
Sub aRubber_LongBands_Hit(idx):PlaySoundAtBall "fx_rubber_longband":End Sub
Sub aRubberPost_Hit(idx):PlaySoundAtBall "fx_rubber_post":End Sub
Sub aRubberPin_Hit(idx):PlaySoundAtBall "fx_rubber_pin":End Sub
Sub aRubberPeg_Hit(idx):PlaySoundAtBall "fx_rubber_peg":End Sub
Sub aPlastics_Hit(idx):PlaySoundAtBall "fx_PlasticHit":End Sub
Sub aGates_Hit(idx):PlaySoundAtBall "fx_Gate":End Sub
Sub aWood_Hit(idx):PlaySoundAtBall "fx_Woodhit":End Sub

' VPX8 animations
Sub LeftFlipper_Animate:LeftFlipperTop.RotZ = LeftFlipper.CurrentAngle:End Sub
Sub RightFlipper_Animate:RightFlipperTop.RotZ = RightFlipper.CurrentAngle:End Sub
Sub Gate_Animate:Gate3Flap.RotZ = Gate.currentangle:End Sub
Sub sw10_Animate:Primitive022.RotY = sw10.currentangle:End Sub

'***************************************************************
'             Supporting Ball & Sound Functions v3.0
'  includes random pitch in PlaySoundAt and PlaySoundAtBall
'***************************************************************

Dim TableWidth, TableHeight

TableWidth = Petaco.width
TableHeight = Petaco.height

Function Vol(ball) ' Calculates the Volume of the sound based on the ball speed
    Vol = Csng(BallVel(ball) ^2 / 2000)
End Function

Function Pan(ball) ' Calculates the pan for a ball based on the X position on the table. "Petaco" is the name of the table
    Dim tmp
    tmp = ball.x * 2 / TableWidth-1
    If tmp> 0 Then
        Pan = Csng(tmp ^10)
    Else
        Pan = Csng(-((- tmp) ^10) )
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
    PlaySound soundname, 0, 1, Pan(tableobj), 0.1, 0, 0, 0, AudioFade(tableobj)
End Sub

Sub PlaySoundAtBall(soundname) ' play a sound at the ball position, like rubbers, targets, metals, plastics
    PlaySound soundname, 0, Vol(ActiveBall), pan(ActiveBall), 0.2, Pitch(ActiveBall) * 10, 0, 0, AudioFade(ActiveBall)
End Sub

Function RndNbr(n) 'returns a random number between 1 and n
    Randomize timer
    RndNbr = Int((n * Rnd) + 1)
End Function

'******************************
'   JP's VP10 Rolling Sounds
'******************************

'EM version

Const tnob = 9    'total number of balls
Const lob = 2     'number of locked balls
Const maxvel = 26 'max ball velocity
ReDim rolling(tnob)
InitRolling

Sub InitRolling
    Dim i
    For i = 0 to tnob
        rolling(i) = False
    Next
    RollingUpdate.Enabled = 1
End Sub

Sub RollingUpdate_timer() 'call this routine from any realtime timer you may have, running at an interval of 10 is good.

    Dim BOT, b, ballpitch, ballvol, speedfactorx, speedfactory
    BOT = GetBalls

    ' stop the sound of deleted balls and hide the shadow
    For b = UBound(BOT) + 1 to tnob
        rolling(b) = False
        StopSound("fx_ballrolling" & b)
    Next

    ' exit the sub if no balls on the table
    If UBound(BOT) = lob - 1 Then Exit Sub 'there no extra balls on this table

    ' draw the ball shadow
    For b = lob to UBound(BOT)

        'play the rolling sound for each ball
        If BallVel(BOT(b) )> 1 Then
            ballpitch = Pitch(BOT(b) )
            ballvol = Vol(BOT(b) )
            rolling(b) = True
            PlaySound("fx_ballrolling" & b), -1, ballvol, Pan(BOT(b) ), 0, ballpitch, 1, 0, AudioFade(BOT(b) )
        Else
            If rolling(b) = True Then
                StopSound("fx_ballrolling" & b)
                rolling(b) = False
            End If
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

'**********************
' Ball Collision Sound
'**********************

Sub OnBallBallCollision(ball1, ball2, velocity)
    PlaySound("collide3"), 0, Csng(velocity) ^2 / 2000, Pan(ball1), 0, Pitch(ball1), 0, 0, AudioFade(ball1)
End Sub

'*********************************
' Table Options F12 User Options
'*********************************
' Table1.Option arguments are:
' - option name, minimum value, maximum value, step between valid values, default value, unit (0=None, 1=Percent), an optional array of literal strings

Dim LUTImage

Sub Petaco_OptionEvent(ByVal eventId)
    Dim x, y

    'LUT
    LutImage = Petaco.Option("Select LUT", 0, 21, 1, 0, 0, Array("Normal 0", "Normal 1", "Normal 2", "Normal 3", "Normal 4", "Normal 5", "Normal 6", "Normal 7", "Normal 8", "Normal 9", "Normal 10", _
        "Warm 0", "Warm 1", "Warm 2", "Warm 3", "Warm 4", "Warm 5", "Warm 6", "Warm 7", "Warm 8", "Warm 9", "Warm 10") )
    UpdateLUT

    ' Cabinet rails
    x = Petaco.Option("Cabinet Rails", 0, 1, 1, 1, 0, Array("Hide", "Show") )
    For each y in aRails:y.visible = x:next
End Sub

Sub UpdateLUT
    Select Case LutImage
        Case 0:Petaco.ColorGradeImage = "LUT0"
        Case 1:Petaco.ColorGradeImage = "LUT1"
        Case 2:Petaco.ColorGradeImage = "LUT2"
        Case 3:Petaco.ColorGradeImage = "LUT3"
        Case 4:Petaco.ColorGradeImage = "LUT4"
        Case 5:Petaco.ColorGradeImage = "LUT5"
        Case 6:Petaco.ColorGradeImage = "LUT6"
        Case 7:Petaco.ColorGradeImage = "LUT7"
        Case 8:Petaco.ColorGradeImage = "LUT8"
        Case 9:Petaco.ColorGradeImage = "LUT9"
        Case 10:Petaco.ColorGradeImage = "LUT10"
        Case 11:Petaco.ColorGradeImage = "LUT Warm 0"
        Case 12:Petaco.ColorGradeImage = "LUT Warm 1"
        Case 13:Petaco.ColorGradeImage = "LUT Warm 2"
        Case 14:Petaco.ColorGradeImage = "LUT Warm 3"
        Case 15:Petaco.ColorGradeImage = "LUT Warm 4"
        Case 16:Petaco.ColorGradeImage = "LUT Warm 5"
        Case 17:Petaco.ColorGradeImage = "LUT Warm 6"
        Case 18:Petaco.ColorGradeImage = "LUT Warm 7"
        Case 19:Petaco.ColorGradeImage = "LUT Warm 8"
        Case 20:Petaco.ColorGradeImage = "LUT Warm 9"
        Case 21:Petaco.ColorGradeImage = "LUT Warm 10"
    End Select
End Sub

' Thalamus : Exit in a clean and proper way
Sub Petaco_exit
  Controller.Pause = False
  Controller.Stop
End Sub


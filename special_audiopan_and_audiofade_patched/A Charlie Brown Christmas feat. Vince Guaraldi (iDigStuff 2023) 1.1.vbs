Option Explicit
Randomize
SetLocale 1033
' ****************************************************************
'               VISUAL PINBALL X EM Script por JPSalas
'         Script Básico para juegos EM Script hasta 4 players
'        usa el core.vbs para funciones extras
'                        Version 4.0.1
' Bumper    IPDB: https://www.ipdb.org/machine.cgi?id=6194
' ****************************************************************

'  Solenoid DOF Added by Outhere
'101 Left Flipper
'102 Right Flipper
'103 Left Slingshot
'104
'105 Right Slingshot
'106
'107 Bumper Left
'108 Bumper Center
'109 Bumper Right
'110 Ball Release
'111 Target3  (Bell)
'112 knocker
'113
'114 Chime
'115 Chime
'116 Chime

'******************************************************************
'******************************************************************

'           Charlie Brown Christmas
'           version 1.0.0
'           by iDigStuff
'
'Scripting, physics, flipper shadows and many other tweaks by Apophis
'PF shadows and b2s by HauntFreaks
'VR team - DarthVito, Leojreimroc, and TastyWasps
'Based on "Bumper" by JPSalas
'Thank you VPW Helpshop testers Passion4pins, Studlygoorite, bietekwiet

'UPDATE 1,1
'apophis - reintegrated DOF code
'hauntfreaks - upscaled PF
'passion4pins - sally hair fix
'idigstuff - increased bumper bell sfx , added standalone code fix from somatik

'******************************************************************

'MUSIC LINK (Place entire Charlie Brown folder into  /visualpinball/music)
'https://mega.nz/folder/1VIghDZK#HE57FDs0_Bb2Hwb3odgUlg

'****************************
'PLAYER OPTIONS
'****************************

Const JukeboxMode = 0   'Jukebox mode songs do not change when ball drains  0 = OFF  1 = ON
Const LMagnaStop = 0        ' Left Magnasave Functionality  0 = Previous Track  1 =  Stop Music
Const BallsPerGame = 5      ' Set to 3 or 5 Balls per game

'----- General Sound Options -----
Const VolumeDial = 0.8        'Overall Mechanical sound effect volume. Recommended values should be no greater than 1.
Const BallRollVolume = 0.5      'Level of ball rolling volume. Value between 0 and 1

'****************************
'END PLAYER OPTIONS
'****************************


' Valores Constantes de las físicas de los flippers - se usan en la creación de las bolas, tienen que cargarse antes del core.vbs
Const BallSize = 50 ' el tamaño normal es 50 unidades de VP.
Const BallMass = 1  ' la pesadez de la bola, este valor va de acuerdo a la fuerza de los flippers y el plunger

' Carga el core.vbs para poder usas sus funciones, sobre todo el vpintimer.addtimer
LoadCoreFiles

Sub LoadCoreFiles
    On Error Resume Next
    ExecuteGlobal GetTextFile("core.vbs")
    If Err Then MsgBox "Can't open core.vbs"
    On Error Resume Next
    ExecuteGlobal GetTextFile("controller.vbs")
    If Err Then MsgBox "Can't open controller.vbs"
End Sub

' Valores que se pueden cambiar
Const Tiempo = 30 '30 seconds, 20 seconds o 10 seconds

' Valores Constants
Const TableName = "bumper" ' se usa para cargar y grabar los highscore y creditos
Const cGameName = "bumper" ' para el B2S
Const MaxPlayers = 1       ' de 1 a 4
Const MaxMultiplier = 3    ' limita el bonus multiplicador a 3
Const Special1 = 520000    ' puntuación a obtener para partida extra
Const Special2 = 780000    ' puntuación a obtener para partida extra
Const Special3 = 990000    ' puntuación a obtener para partida extra

' Variables Globales
Dim PlayersPlayingGame
Dim CurrentPlayer
Dim Credits
Dim Bonus
Dim BonusMultiplier
Dim BallsRemaining(4)
Dim ExtraBallsAwards(4)
Dim Special1Awarded(4)
Dim Special2Awarded(4)
Dim Special3Awarded(4)
Dim Score(4)
Dim HighScore
Dim Match
Dim Tilt
Dim TiltSensitivity
Dim Tilted
Dim Add10
Dim Add100
Dim Add1000
Dim Add10000
Dim x

' Variables de control
Dim LastSwitchHit
Dim BallsOnPlayfield

' Variables de tipo Boolean (verdadero ó falso, True ó False)
Dim bAttractMode
Dim bFreePlay
Dim bGameInPlay
Dim bOnTheFirstBall
Dim bExtraBallWonThisBall
Dim bJustStarted
Dim bBallInPlungerLane
Dim bBallSaverActive

' core.vbs variables, como imanes, impulse plunger
Dim cbball 'captive ball

' *********************************************************************
'                Rutinas comunes para todas las mesas
' *********************************************************************

Sub Table1_Init()

    ' Inicializar diversos objetos de la mesa, como droptargets, animations...
    VPObjects_Init
    LoadEM

    ' bola captiva
    Set cbball = New cvpmCaptiveBall
    With cbball
        .InitCaptive CapTrigger1, CapWall1, CapKicker1, 0
        .ForceTrans = .9
        .MinForce = 3.5
        .CreateEvents "cbball"
        .Start
    End With

    ' Carga los valores grabados highscore y créditos
    Credits = 1
    Loadhs

    UpdateCredits

    ' Juego libre o con monedas: si es True entonces no se usarán monedas
    bFreePlay = False 'queremos monedas

    ' Inicialiar las variables globales de la mesa
    bAttractMode = False
    bOnTheFirstBall = False
    bGameInPlay = False
    bBallInPlungerLane = False
    LastSwitchHit = ""
    BallsOnPlayfield = 0
    Tilt = 0
    TiltSensitivity = 6
    Tilted = False
    Match = 0
    bJustStarted = True
    Add10 = 0
    Add100 = 0
    Add1000 = 0

    ' pone la mesa en modo de espera
    EndOfGame

    'Enciende las luces GI despues de un segundo
    vpmtimer.addtimer 1000, "GiOn '"

    ' Quita los laterales y las puntuaciones cuando la mesa se juega en modo FS
    If Table1.ShowDT then
        lrail.Visible = True
        rrail.Visible = True
        For each x in aReels
            x.Visible = 1
        Next
    Else
        lrail.Visible = False
        rrail.Visible = False
        For each x in aReels
            x.Visible = 0
        Next
    End If

  If VRMode = True Then
    dim vrbgobj
    VRBGGameOverFlash.enabled = True
    FlasherMatch
    For Each vrbgobj in VRBGBalls:vrbgobj.visible = 0 : Next
  End If
    LoadLUT
End Sub

'*************************************************************
' VR Room Auto-Detect
'*************************************************************
Dim VR_Obj, VRMode

If RenderingMode = 2 Then
    VRMode = True
  lrail.Visible = 0
  lrail1.Visible = 0
  rrail.Visible = 0
  rrail1.Visible = 0
  flasher1.Visible = 0
  flasher2.Visible = 0
    For Each VR_Obj in VRCabinet : VR_Obj.Visible = 1 : Next
  For Each VR_Obj in VRBackglass : VR_Obj.Visible = 1 : Next
    For Each VR_Obj in VRCBRoom : VR_Obj.Visible = 1 : Next
  SetBackglass
  VRBGSnoopyFlash.enabled = true
Else
    VRMode = False
    For Each VR_Obj in VRCabinet : VR_Obj.Visible = 0 : Next
  For Each VR_Obj in VRBackglass : VR_Obj.Visible = 0 : Next
    For Each VR_Obj in VRCBRoom : VR_Obj.Visible = 0 : Next
End If

'Dim song
'Sub NextTrack

'If song = 0 Then PlayMusic "Charlie Brown\1.mp3"  End If
'If song = 1 Then PlayMusic "Charlie Brown\4.mp3"  End If
'If song = 2 Then PlayMusic "Charlie Brown\3.mp3"  End If
'If song = 3 Then PlayMusic "Charlie Brown\5.mp3"  End If
'If song = 4 Then PlayMusic "Charlie Brown\7.mp3"  End If
'If song = 5 Then PlayMusic "Charlie Brown\6.mp3"  End If
'If song = 6 Then PlayMusic "Charlie Brown\2.mp3"  End If
'If song = 7 Then PlayMusic "Charlie Brown\8.mp3"  End If

'song = (song + 1) mod 8
'End Sub

'Sub Table1_MusicDone()
' NextTrack
'End Sub




'**************************************
'*         HiRez00: Music Mod         *
'*          RetroG33K added       *
'*        Ablity to stop music        *
'*  Stop song Selection at game Start'*
'*  Reselect song ball launch or Drain*
'**************************************

Dim musicNum
Dim i
Dim musicPlaying : musicPlaying = False
Sub NextTrack



If musicNum = 0 Then PlayMusic "Charlie Brown\1.mp3" End If ' 1
If musicNum = 1 Then PlayMusic "Charlie Brown\4.mp3" End If ' 2
If musicNum = 2 Then PlayMusic "Charlie Brown\3.mp3" End If ' 3
If musicNum = 3 Then PlayMusic "Charlie Brown\5.mp3" End If ' 4
If musicNum = 4 Then PlayMusic "Charlie Brown\7.mp3" End If ' 5
If musicNum = 5 Then PlayMusic "Charlie Brown\6.mp3" End If ' 6
If musicNum = 6 Then PlayMusic "Charlie Brown\2.mp3" End If ' 7
If musicNum = 7 Then PlayMusic "Charlie Brown\8.mp3" End If ' 8


musicNum = (musicNum + 1) mod 9

  i = MusicNum
  if i=1 then
    Primitive_InstructionCardL001.Image = "Track1"
  End if
  if i=2 then
    Primitive_InstructionCardL001.Image = "Track2"
  End if
  if i=3 then
    Primitive_InstructionCardL001.Image = "Track3"
  End if
  if i=4 then
    Primitive_InstructionCardL001.Image = "Track4"
  End if
  if i=5 then
    Primitive_InstructionCardL001.Image = "Track5"
  End if
  if i=6 then
    Primitive_InstructionCardL001.Image = "Track6"
  End if
  if i=7 then
    Primitive_InstructionCardL001.Image = "Track7"
  End if
  if i=8 then
    Primitive_InstructionCardL001.Image = "Track8"
  End if

  If i<9 Then
  Exit Sub
  End If

End Sub

Sub PreviousTrack



musicNum = (musicNum - 1)

If musicNum = 1 Then PlayMusic "Charlie Brown\1.mp3" End If ' 1
If musicNum = 2 Then PlayMusic "Charlie Brown\4.mp3" End If ' 2
If musicNum = 3 Then PlayMusic "Charlie Brown\3.mp3" End If ' 3
If musicNum = 4 Then PlayMusic "Charlie Brown\5.mp3" End If ' 4
If musicNum = 5 Then PlayMusic "Charlie Brown\7.mp3" End If ' 5
If musicNum = 6 Then PlayMusic "Charlie Brown\6.mp3" End If ' 6
If musicNum = 7 Then PlayMusic "Charlie Brown\2.mp3" End If ' 7
If musicNum = 8 Then PlayMusic "Charlie Brown\8.mp3" End If ' 8


  i = MusicNum
  if i=1 then
    Primitive_InstructionCardL001.Image = "Track1"
  End if
  if i=2 then
    Primitive_InstructionCardL001.Image = "Track2"
  End if
  if i=3 then
    Primitive_InstructionCardL001.Image = "Track3"
  End if
  if i=4 then
    Primitive_InstructionCardL001.Image = "Track4"
  End if
  if i=5 then
    Primitive_InstructionCardL001.Image = "Track5"
  End if
  if i=6 then
    Primitive_InstructionCardL001.Image = "Track6"
  End if
  if i=7 then
    Primitive_InstructionCardL001.Image = "Track7"
  End if
  if i=8 then
    Primitive_InstructionCardL001.Image = "Track8"
  End if
  If i= 0 Then
  NextTrack
  End If
End Sub

Sub Table1_MusicDone
        NextTrack
End Sub


Sub SelectSong(keycode)
    If keycode = PlungerKey OR keycode = StartGameKey Then
        bsongSelect = False

    End If
' If KeyCode = RightMagnaSave Then NextTrack
'    If KeyCode = LeftMagnaSave Then PreviousTrack
End Sub


'*****************************
'*       End Music Mod       *
'*****************************




'******
' Keys
'******

Sub Table1_KeyDown(ByVal Keycode)
    If EnteringInitials then
        CollectInitials(keycode)
        Exit Sub
    End If

 '   If keycode = CenterTiltKey Then Nudge 0, 6:SoundNudgeCenter

    If keycode = LeftMagnaSave and LMagnaStop = 0 Then PreviousTrack
  If keycode = LeftMagnaSave and LMagnaStop = 1 Then EndMusic
    If keycode = RightMagnaSave and bGameInPlay = True Then NextTrack
  '      If bLutActive Then NextLUT:End If



    ' añade monedas
    If Keycode = AddCreditKey Then
        If(Tilted = False)Then
      Select Case Int(Rnd * 3)
      Case 0
        PlaySound ("Coin_In_1"), 0, CoinSoundLevel, 0, 0.25
      Case 1
        PlaySound ("Coin_In_2"), 0, CoinSoundLevel, 0, 0.25
      Case 2
        PlaySound ("Coin_In_3"), 0, CoinSoundLevel, 0, 0.25
    End Select
            AddCredits 1
        End If
    End If

    ' el plunger
    If keycode = PlungerKey Then
        Plunger.Pullback
    SoundPlungerPull
    If VRMode=True Then
      TimerPlunger.Enabled = True
      TimerPlunger2.Enabled = False
    End If
    End If

    ' Funcionamiento normal de los flipers y otras teclas durante el juego

    If bGameInPlay AND NOT Tilted Then
        ' teclas de la falta
        If keycode = LeftTiltKey Then Nudge 90,2:SoundNudgeLeft:CheckTilt
        If keycode = RightTiltKey Then Nudge 270, 2:SoundNudgeRight:CheckTilt
        If keycode = CenterTiltKey Then Nudge 0, 2:SoundNudgeCenter:CheckTilt


    If keycode = LeftFlipperKey Then
      FlipperActivate LeftFlipper, LFPress
      SolLFlipper True  'This would be called by the solenoid callbacks if using a ROM
      If VRMode=True Then PinCab_Flipper_Left.X = PinCab_Flipper_Left.X + 10
    End If

    If keycode = RightFlipperKey Then
      FlipperActivate RightFlipper, RFPress
      SolRFlipper True  'This would be called by the solenoid callbacks if using a ROM
      If VRMode=True Then PinCab_Flipper_Right.X = PinCab_Flipper_Right.X - 10
    End If

    If keycode = StartGameKey Then
      SoundStartButton
  End If

        ' tecla de empezar el juego
        If keycode = StartGameKey Then
            If((PlayersPlayingGame < MaxPlayers)AND(bOnTheFirstBall = True))Then

                If(bFreePlay = True)Then
                    PlayersPlayingGame = PlayersPlayingGame + 1
                    PlayersReel.SetValue, PlayersPlayingGame
                'PlaySound "so_fanfare1"
                Else
                    If(Credits > 0)then
                        PlayersPlayingGame = PlayersPlayingGame + 1
                        Credits = Credits - 1
                        UpdateCredits
                        UpdateBallInPlay
                    Else
                    ' no hay suficientes créditos para empezar el juego.
                    'PlaySound "so_nocredits"
                    End If
                End If
            End If
        End If
        Else ' If (GameInPlay)

            If keycode = StartGameKey Then
                If(bFreePlay = True)Then
                    If(BallsOnPlayfield = 0)Then
                        ResetScores
                        ResetForNewGame()
            PlayMusic "Charlie Brown\8.mp3"
            Primitive_InstructionCardL001.Image = "Track8"
                    End If
                Else
                    If(Credits > 0)Then
                        If(BallsOnPlayfield = 0)Then
                            Credits = Credits - 1
                            UpdateCredits
                            ResetScores
                            ResetForNewGame()
              PlayMusic"Charlie Brown\8.mp3"
              Primitive_InstructionCardL001.Image = "Track8"
                        End If
                    Else
                    ' Not Enough Credits to start a game.
                    'PlaySound "so_nocredits"
                    End If
                End If
            End If
    End If ' If (GameInPlay)

End Sub

Sub Table1_KeyUp(ByVal keycode)

    If EnteringInitials then
        Exit Sub
    End If

    If keycode = LeftMagnaSave Then bLutActive = False: LutBox.text = ""

    If bGameInPlay AND NOT Tilted Then
        ' teclas de los flipers
        If keycode = LeftFlipperKey Then
      FlipperDeActivate LeftFlipper, LFPress
      SolLFlipper 0
      If VRMode=True Then PinCab_Flipper_Left.X = PinCab_Flipper_Left.X - 10
    End If
        If keycode = RightFlipperKey Then
      FlipperDeActivate RightFlipper, RFPress
      SolRFlipper 0
      If VRMode=True Then PinCab_Flipper_Right.X = PinCab_Flipper_Right.X + 10
    End If
    End If

    If keycode = PlungerKey Then
        Plunger.Fire
    If VRMode=True Then TimerPlunger.Enabled = False
    If VRMode=True Then TimerPlunger2.Enabled = True
    If VRMode=True Then PinCab_Plunger.Y = 2107.542
        If bBallInPlungerLane Then
            SoundPlungerReleaseBall()
        Else
            SoundPlungerReleaseNoBall()
        End If
    End If
End Sub

'*************
' Para la mesa
'*************

Sub table1_Paused
End Sub

Sub table1_unPaused
End Sub

Sub table1_Exit
    Savehs
'Controller.Stop
End Sub



'*******************
' Luces GI
'*******************

Sub GiOn 'enciende las luces GI
    Dim bulb
    For each bulb in aGiLights
        bulb.State = 1
  shadow1.visible = 1
    Next
End Sub

Sub GiOff 'apaga las luces GI
    Dim bulb
    For each bulb in aGiLights
        bulb.State = 0
  shadow1.visible = 0
    Next
End Sub

'**************
' TILT - Falta
'**************

'el "timer" TiltDecreaseTimer resta .01 de la variable "Tilt" cada ronda

Sub CheckTilt                     'esta rutina se llama cada vez que das un golpe a la mesa
    Tilt = Tilt + TiltSensitivity 'añade un valor al contador "Tilt"
    TiltDecreaseTimer.Enabled = True
    If Tilt > 15 Then             'Si la variable "Tilt" es más de 15 entonces haz falta
        Tilted = True
        TiltReel.SetValue 1       'muestra Tilt en la pantalla
        If B2SOn then
            Controller.B2SSetTilt 1
        end if
    If VRMode = True Then
      VRBGTilt.visible = 1:VRBGTiltBulb.visible = 1
    End If
        DisableTable True
        TiltRecoveryTimer.Enabled = True 'empieza una pausa a fin de que todas las bolas se cuelen
    End If
End Sub

Sub TiltDecreaseTimer_Timer
    ' DecreaseTilt
    If Tilt > 0 Then
        Tilt = Tilt - 0.1
    Else
        TiltDecreaseTimer.Enabled = False
    End If
End Sub

Sub DisableTable(Enabled)
    If Enabled Then
        'Apaga todas las luces Gi de la mesa
        GiOff
        'Disable slings, bumpers etc
        LeftFlipper.RotateToStart
        RightFlipper.RotateToStart
        Bumper1.Threshold = 100
        Bumper2.Threshold = 100
        Bumper3.Threshold = 100
        LeftSlingshot.Disabled = 1
        RightSlingshot.Disabled = 1
    Else
        'enciende de nuevo todas las luces GI, bumpers y slingshots
        GiOn
        Bumper1.Threshold = 1
        Bumper2.Threshold = 1
        Bumper3.Threshold = 1
        LeftSlingshot.Disabled = 0
        RightSlingshot.Disabled = 0
    End If
End Sub

Sub TiltRecoveryTimer_Timer()
    ' si todas las bolas se han colado, entonces ..
    If(BallsOnPlayfield = 0)Then
        '... haz el fin de bola normal
        EndOfBall()
        TiltRecoveryTimer.Enabled = False
    End If
' de lo contrario esta rutina continúa hasta que todas las bolas se han colado
End Sub

'***************************
'   LUT - Darkness control
'***************************

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

Sub NextLUT:LUTImage = (LUTImage + 1)MOD 15:UpdateLUT:SaveLUT:Lutbox.text = "level of darkness " & LUTImage + 1:End Sub

Sub UpdateLUT
    Select Case LutImage
        Case 0:table1.ColorGradeImage = "LUT0":GiIntensity = 1:ChangeGIIntensity 1
        Case 1:table1.ColorGradeImage = "LUT1":GiIntensity = 1.05:ChangeGIIntensity 1
        Case 2:table1.ColorGradeImage = "LUT2":GiIntensity = 1.1:ChangeGIIntensity 1
        Case 3:table1.ColorGradeImage = "LUT3":GiIntensity = 1.15:ChangeGIIntensity 1
        Case 4:table1.ColorGradeImage = "LUT4":GiIntensity = 1.2:ChangeGIIntensity 1
        Case 5:table1.ColorGradeImage = "LUT5":GiIntensity = 1.25:ChangeGIIntensity 1
        Case 6:table1.ColorGradeImage = "LUT6":GiIntensity = 1.3:ChangeGIIntensity 1
        Case 7:table1.ColorGradeImage = "LUT7":GiIntensity = 1.35:ChangeGIIntensity 1
        Case 8:table1.ColorGradeImage = "LUT8":GiIntensity = 1.4:ChangeGIIntensity 1
        Case 9:table1.ColorGradeImage = "LUT9":GiIntensity = 1.45:ChangeGIIntensity 1
        Case 10:table1.ColorGradeImage = "LUT10":GiIntensity = 1.5:ChangeGIIntensity 1
        Case 11:table1.ColorGradeImage = "LUT11":GiIntensity = 1.55:ChangeGIIntensity 1
        Case 12:table1.ColorGradeImage = "LUT12":GiIntensity = 1.6:ChangeGIIntensity 1
        Case 13:table1.ColorGradeImage = "LUT13":GiIntensity = 1.65:ChangeGIIntensity 1
        Case 14:table1.ColorGradeImage = "LUT14":GiIntensity = 1.7:ChangeGIIntensity 1
    End Select
End Sub

Dim GiIntensity
GiIntensity = 1   'used by the LUT changing to increase the GI lights when the table is darker

Sub ChangeGiIntensity(factor) 'changes the intensity scale
    Dim bulb
    For each bulb in aGiLights
        bulb.IntensityScale = GiIntensity * factor
    Next
End Sub


Dim TableWidth, TableHeight

TableWidth = Table1.width
TableHeight = Table1.height

'*******************************************
' ZTIM:Timers
'*******************************************

Sub GameTimer_Timer() 'The game timer interval; should be 10 ms
  Cor.Update    'update ball tracking (this sometimes goes in the RDampen_Timer sub)
  RollingUpdate   'update rolling sounds
' DoDTAnim    'handle drop target animations
' DoSTAnim    'handle stand up target animations
' queue.Tick    'handle the queue system
End Sub

Dim FrameTime, InitFrameTime
InitFrameTime = 0
Sub FrameTimer_Timer() 'The frame timer interval should be -1, so executes at the display frame rate
  FrameTime = GameTime - InitFrameTime
  InitFrameTime = GameTime  'Count frametime
  FlipperLSh.RotZ = LeftFlipper.CurrentAngle
  FlipperRSh.RotZ = RightFlipper.CurrentAngle
  'FlipperVisualUpdate     'update flipper shadows and primitives
  'If DynamicBallShadowsOn Or AmbientBallShadowOn Then DynamicBSUpdate 'update ball shadows
End Sub


'******************************************************
' ZBRL:  BALL ROLLING AND DROP SOUNDS
'******************************************************

' Be sure to call RollingUpdate in a timer with a 10ms interval see the GameTimer_Timer() sub

Const tnob = 19   'total number of balls
ReDim rolling(tnob)
InitRolling

Dim DropCount
ReDim DropCount(tnob)

Sub InitRolling
  Dim i
  For i = 0 To tnob
    rolling(i) = False
  Next
End Sub


Sub RollingUpdate()
  Dim b
     Dim BOT
     BOT = GetBalls

  ' stop the sound of deleted balls
  For b = UBound(BOT) + 1 To tnob - 1
    ' Comment the next line if you are not implementing Dyanmic Ball Shadows
  ' If AmbientBallShadowOn = 0 Then BallShadowA(b).visible = 0
    rolling(b) = False
    StopSound("BallRoll_" & b)
  Next

  ' exit the sub if no balls on the table
  If UBound(BOT) =  - 1 Then Exit Sub

  ' play the rolling sound for each ball
  For b = 0 To UBound(BOT)
    If BallVel(BOT(b)) > 1 And BOT(b).z < 30 Then
      rolling(b) = True
      PlaySound ("BallRoll_" & b), - 1, VolPlayfieldRoll(BOT(b)) * BallRollVolume * VolumeDial, AudioPan(BOT(b)), 0, PitchPlayfieldRoll(BOT(b)), 1, 0, AudioFade(BOT(b))
    Else
      If rolling(b) = True Then
        StopSound("BallRoll_" & b)
        rolling(b) = False
      End If
    End If

    ' Ball Drop Sounds
    If BOT(b).VelZ <  - 1 And BOT(b).z < 55 And BOT(b).z > 27 Then 'height adjust for ball drop sounds
      If DropCount(b) >= 5 Then
        DropCount(b) = 0
        If BOT(b).velz >  - 7 Then
          RandomSoundBallBouncePlayfieldSoft BOT(b)
        Else
          RandomSoundBallBouncePlayfieldHard BOT(b)
        End If
      End If
    End If

    If DropCount(b) < 5 Then
      DropCount(b) = DropCount(b) + 1
    End If

    ' "Static" Ball Shadows
    ' Comment the next If block, if you are not implementing the Dynamic Ball Shadows
'   If AmbientBallShadowOn = 0 Then
'     If BOT(b).Z > 30 Then
'       BallShadowA(b).height = BOT(b).z - BallSize / 4   'This is technically 1/4 of the ball "above" the ramp, but it keeps it from clipping the ramp
'     Else
'       BallShadowA(b).height = 0.1
'     End If
'     BallShadowA(b).Y = BOT(b).Y + offsetY
'     BallShadowA(b).X = BOT(b).X + offsetX
'     BallShadowA(b).visible = 1
'   End If
  Next
End Sub

'******************************************************
'****  END BALL ROLLING AND DROP SOUNDS
'******************************************************


'************************************************************************************************************************
' Solo para VPX 10.2 y posteriores.
' FlashForMs hará parpadear una luz o un flash por unos milisegundos "TotalPeriod" cada tantos milisegundos "BlinkPeriod"
' Cuando el "TotalPeriod" haya terminado, la luz o el flasher se pondrá en el estado especificado por el valor "FinalState"
' El valor de "FinalState" puede ser: 0=apagado, 1=encendido, 2=regreso al estado anterior
'************************************************************************************************************************

Sub FlashForMs(MyLight, TotalPeriod, BlinkPeriod, FinalState)

    If TypeName(MyLight) = "Light" Then ' la luz es del tipo "light"

        If FinalState = 2 Then
            FinalState = MyLight.State  'guarda el estado actual de la luz
        End If
        MyLight.BlinkInterval = BlinkPeriod
        MyLight.Duration 2, TotalPeriod, FinalState
    ElseIf TypeName(MyLight) = "Flasher" Then ' la luz es del tipo "flash"
        Dim steps
        ' Store all blink information
        steps = Int(TotalPeriod / BlinkPeriod + .5) 'número de encendidos y apagados que hay que ejecutar
        If FinalState = 2 Then                      'guarda el estado actual del flash
            FinalState = ABS(MyLight.Visible)
        End If
        MyLight.UserValue = steps * 10 + FinalState 'guarda el número de parpadeos

        ' empieza los parpadeos y crea la rutina que se va a ejecutar como un timer que se va a ejecutar los parpadeos
        MyLight.TimerInterval = BlinkPeriod
        MyLight.TimerEnabled = 0
        MyLight.TimerEnabled = 1
        ExecuteGlobal "Sub " & MyLight.Name & "_Timer:" & "Dim tmp, steps, fstate:tmp=me.UserValue:fstate = tmp MOD 10:steps= tmp\10 -1:Me.Visible = steps MOD 2:me.UserValue = steps *10 + fstate:If Steps = 0 then Me.Visible = fstate:Me.TimerEnabled=0:End if:End Sub"
    End If
End Sub

'****************************************
' Inicializa la mesa para un juego nuevo
'****************************************

Sub ResetForNewGame()
    'debug.print "ResetForNewGame"
    Dim i

    bGameInPLay = True
    bBallSaverActive = False

    'pone a cero los marcadores y apaga las luces de espera.
    StopAttractMode
    If B2SOn then
        Controller.B2SSetGameOver 0
    end if
  If VRMode = True Then
    VRBGGameOverFlash.enabled = False
    VRBGGameOver.visible = 0:VRBGGOBulb1.visible = 0:VRBGGOBulb2.visible = 0
  End If
    ' enciende las luces GI si estuvieran apagadas
    GiOn


    CurrentPlayer = 1
    PlayersPlayingGame = 1
    bOnTheFirstBall = True
    For i = 1 To MaxPlayers
        Score(i) = 0
        ExtraBallsAwards(i) = 0
        Special1Awarded(i) = False
        Special2Awarded(i) = False
        Special3Awarded(i) = False
        BallsRemaining(i) = BallsPerGame
    Next
  If VRMode = True Then
    UpdateVRReels  0 ,0 ,Score(0), 0, 0,0,0,0,0,0
    UpdateVRReels  1 ,1 ,Score(0), 0, 0,0,0,0,0,0
    UpdateVRReels  2 ,2 ,Score(0), 0, 0,0,0,0,0,0
    UpdateVRReels  3 ,3 ,Score(0), 0, 0,0,0,0,0,0
  End If
    BonusMultiplier = 1
    Bonus = 0
    UpdateBallInPlay

    Clear_Match

    ' inicializa otras variables
    Tilt = 0

    ' inicializa las variables del juego
    Game_Init()

    ' ahora puedes empezar una música si quieres
    ' empieza la rutina "Firstball" despues de una pequeña pausa
    vpmtimer.addtimer 2000, "FirstBall '"
End Sub

' esta pausa es para que la mesa tenga tiempo de poner los marcadores a cero y actualizar las luces

Sub FirstBall
    'debug.print "FirstBall"
    ' ajusta la mesa para una bola nueva, sube las dianas abatibles, etc
    ResetForNewPlayerBall()
    ' crea una bola nueva en la zona del plunger
    CreateNewBall()

End Sub

' (Re-)inicializa la mesa para una bola nueva, tanto si has perdido la bola, oe le toca el turno al otro jugador

Sub ResetForNewPlayerBall()
    'debug.print "ResetForNewPlayerBall"
    ' Se asegura que los marcadores están activados para el jugador de turno
    AddScore 0

    ' ajusta el multiplicador del bonus multiplier a 1X (si hubiese multiplicador en la mesa)

    ' enciende las luces, reinicializa las variables del juego, etc
    bExtraBallWonThisBall = False
    ResetNewBallLights
    ResetNewBallVariables
End Sub

' Crea una bola nueva en la mesa

Sub CreateNewBall()
    ' crea una bola nueva basada en el tamaño y la masa de la bola especificados al principio del script
    BallRelease.CreateSizedBallWithMass BallSize / 2, BallMass

    ' incrementa el número de bolas en el tablero, ya que hay que contarlas
    BallsOnPlayfield = BallsOnPlayfield + 1

    ' actualiza las luces del backdrop
    UpdateBallInPlay

    ' y expulsa la bola
    'PlaySoundAt SoundFXDOF("fx_Ballrel",110,DOFPulse,DOFContactors), BallRelease
  RandomSoundBallRelease BallRelease
    BallRelease.Kick 90, 4
  DOF 110,DOFPulse
End Sub

' El jugador ha perdido su bola, y ya no hay más bolas en juego
' Empieza a contar los bonos

Sub EndOfBall()
    'debug.print "EndOfBall"
    Dim AwardPoints, TotalBonus, ii
    AwardPoints = 0
    TotalBonus = 0
    ' La primera se ha perdido. Desde aquí ya no se puede aceptar más jugadores
    bOnTheFirstBall = False

    ' solo recoge los bonos si no hay falta
    ' el sistema del la falta se encargará de nuevas bolas o del fin de la partida
    ' en este juego solo se cuentan los bonos si la luz1 está encendida

    If NOT Tilted AND li1.State Then
        Select Case BonusMultiplier
            Case 1:BonusCountTimer.Interval = 250
            Case 2:BonusCountTimer.Interval = 400
            Case 3:BonusCountTimer.Interval = 550
        End Select
        BonusCountTimer.Enabled = 1
    Else 'Si hay falta simplemente espera un momento y va directo a la segunta parte después de perder la bola
        vpmtimer.addtimer 400, "EndOfBall2 '"
    End If
End Sub

Sub BonusCountTimer_Timer 'añade los bonos y actualiza las luces
    'debug.print "BonusCount_Timer"
    If Bonus > 0 Then
        Bonus = Bonus -1
        AddScore 10000 * BonusMultiplier
        UpdateBonusLights
    Else
        ' termina la cuenta de los bonos y continúa con el fin de bola
        BonusCountTimer.Enabled = 0
        vpmtimer.addtimer 1000, "EndOfBall2 '"
    End If
End Sub

Sub UpdateBonusLights 'enciende o apaga las luces de los bonos según la variable "Bonus"
    Select Case Bonus
        Case 0:li24.State = 0:li23.State = 0:li22.State = 0:li21.State = 0:li20.State = 0:li19.State = 0:li18.State = 0:li17.State = 0:li16.State = 0:li15.State = 0:li7.State = 0
        Case 1:li24.State = 1:li23.State = 0:li22.State = 0:li21.State = 0:li20.State = 0:li19.State = 0:li18.State = 0:li17.State = 0:li16.State = 0:li15.State = 0:li7.State = 0
        Case 2:li24.State = 0:li23.State = 1:li22.State = 0:li21.State = 0:li20.State = 0:li19.State = 0:li18.State = 0:li17.State = 0:li16.State = 0:li15.State = 0:li7.State = 0
        Case 3:li24.State = 0:li23.State = 0:li22.State = 1:li21.State = 0:li20.State = 0:li19.State = 0:li18.State = 0:li17.State = 0:li16.State = 0:li15.State = 0:li7.State = 0
        Case 4:li24.State = 0:li23.State = 0:li22.State = 0:li21.State = 1:li20.State = 0:li19.State = 0:li18.State = 0:li17.State = 0:li16.State = 0:li15.State = 0:li7.State = 0
        Case 5:li24.State = 0:li23.State = 0:li22.State = 0:li21.State = 0:li20.State = 1:li19.State = 0:li18.State = 0:li17.State = 0:li16.State = 0:li15.State = 0:li7.State = 0
        Case 6:li24.State = 0:li23.State = 0:li22.State = 0:li21.State = 0:li20.State = 0:li19.State = 1:li18.State = 0:li17.State = 0:li16.State = 0:li15.State = 0:li7.State = 0
        Case 7:li24.State = 0:li23.State = 0:li22.State = 0:li21.State = 0:li20.State = 0:li19.State = 0:li18.State = 1:li17.State = 0:li16.State = 0:li15.State = 0:li7.State = 0
        Case 8:li24.State = 0:li23.State = 0:li22.State = 0:li21.State = 0:li20.State = 0:li19.State = 0:li18.State = 0:li17.State = 1:li16.State = 0:li15.State = 0:li7.State = 1
        Case 9:li24.State = 0:li23.State = 0:li22.State = 0:li21.State = 0:li20.State = 0:li19.State = 0:li18.State = 0:li17.State = 0:li16.State = 1:li15.State = 0:li7.State = 1
        Case 10:li24.State = 0:li23.State = 0:li22.State = 0:li21.State = 0:li20.State = 0:li19.State = 0:li18.State = 0:li17.State = 0:li16.State = 0:li15.State = 1:li7.State = 0
    End Select
End Sub

' La cuenta de los bonos ha terminado. Mira si el jugador ha ganado bolas extras
' y si no mira si es el último jugador o la última bola
'
Sub EndOfBall2()
    'debug.print "EndOfBall2"
    ' si hubiese falta, quítala, y pon la cuenta a cero de la falta para el próximo jugador, ó bola

    Tilted = False
    Tilt = 0
    TiltReel.SetValue 0
    If B2SOn then
        Controller.B2SSetTilt 0
    end if
  If VRMode = True Then
    VRBGTilt.visible = 0:VRBGTiltBulb.visible = 0
  End If

    DisableTable False 'activa de nuevo los bumpers y los slingshots

    ' ¿ha ganado el jugador una bola extra?
    If(ExtraBallsAwards(CurrentPlayer) > 0)Then
        'debug.print "Extra Ball"

        ' sí? entonces se la das al jugador
        ExtraBallsAwards(CurrentPlayer) = ExtraBallsAwards(CurrentPlayer)- 1

        ' si no hay más bolas apaga la luz de jugar de nuevo
        If(ExtraBallsAwards(CurrentPlayer) = 0)Then
        'LightShootAgain.State = 0
        'If B2SOn then
        '    Controller.B2SSetShootAgain 0
        'end if
        End If

' aquí se podría poner algún sonido de bola extra o alguna luz que parpadee

' En esta mesa hacemos la bola extra igual como si fuese la siguente bola, haciendo un reset de las variables y dianas
        ResetForNewPlayerBall()

        ' creamos una bola nueva en el pasillo de disparo
        CreateNewBall()
    Else ' no hay bolas extras

        BallsRemaining(CurrentPlayer) = BallsRemaining(CurrentPlayer)- 1

        ' ¿Es ésta la última bola?
        If(BallsRemaining(CurrentPlayer) <= 0)Then

            ' miramos si la puntuación clasifica como el Highscore
            CheckHighScore()
        End If

        ' ésta no es la última bola para éste jugador
        ' y si hay más de un jugador continúa con el siguente
        EndOfBallComplete()
    End If
End Sub

' Esta rutina se llama al final de la cuenta del bonus
' y pasa a la siguente bola o al siguente jugador
'
Sub EndOfBallComplete()
    'debug.print "EndOfBallComplete"
    Dim NextPlayer

    'debug.print "EndOfBall - Complete"

    ' ¿hay otros jugadores?
    If(PlayersPlayingGame > 1)Then
        ' entonces pasa al siguente jugador
        NextPlayer = CurrentPlayer + 1
        ' ¿vamos a pasar del último jugador al primero?
        ' (por ejemplo del jugador 4 al no. 1)
        If(NextPlayer > PlayersPlayingGame)Then
            NextPlayer = 1
        End If
    Else
        NextPlayer = CurrentPlayer
    End If

    'debug.print "Next Player = " & NextPlayer

    ' ¿Hemos llegado al final del juego? (todas las bolas se han jugado de todos los jugadores)
    If((BallsRemaining(CurrentPlayer) <= 0)AND(BallsRemaining(NextPlayer) <= 0))Then

        ' aquí se empieza la lotería, normalmente cuando se juega con monedas
        If bFreePlay = False Then
            Verification_Match
        End If

        ' ahora se pone la mesa en el modo de final de juego
        EndOfGame()
    Else
        ' pasamos al siguente jugador
        CurrentPlayer = NextPlayer

        ' nos aseguramos de que el backdrop muestra el jugador actual
        AddScore 0

        ' hacemos un reset del la mesa para el siguente jugador (ó bola)
        ResetForNewPlayerBall()

        ' y sacamos una bola
        CreateNewBall()
    End If
End Sub

' Esta función se llama al final del juego

Sub EndOfGame()
    'debug.print "EndOfGame"
    bGameInPLay = False
    bJustStarted = False
    If B2SOn then
        Controller.B2SSetGameOver 1
        Controller.B2SSetBallInPlay 0
        Controller.B2SSetPlayerUp 0
        Controller.B2SSetCanPlay 0
    end if
  If VRMode = True Then
    dim vrbgobj
    VRBGGameOverFlash.enabled = true
    For Each vrbgobj in VRBGBalls:vrbgobj.visible = 0 : Next
  End If
    ' asegúrate de que los flippers están en modo de reposo
  FlipperDeActivate LeftFlipper, LFPress
    SolLFlipper 0
  FlipperDeActivate RightFlipper, RFPress
    SolRFlipper 0

    ' pon las luces en el modo de fin de juego
    StartAttractMode
  EndMusic
  PlaySound"fx_match"
  Primitive_InstructionCardL001.Image = "Default"
End Sub

' Esta función calcula el no de bolas que quedan
Function Balls
    Dim tmp
    tmp = BallsPerGame - BallsRemaining(CurrentPlayer) + 1
    If tmp > BallsPerGame Then
        Balls = BallsPerGame
    Else
        Balls = tmp
    End If
End Function

' Esta función calcula el Highscore y te da una partida gratis si has conseguido el Highscore
Sub CheckHighscore
    Dim playertops, si, sj, i, stemp, stempplayers
    For i = 1 to 4
        sortscores(i) = 0
        sortplayers(i) = 0
    Next
    playertops = 0
    For i = 1 to PlayersPlayingGame
        sortscores(i) = Score(i)
        sortplayers(i) = i
    Next
    For si = 1 to PlayersPlayingGame
        For sj = 1 to PlayersPlayingGame-1
            If sortscores(sj) > sortscores(sj + 1)then
                stemp = sortscores(sj + 1)
                stempplayers = sortplayers(sj + 1)
                sortscores(sj + 1) = sortscores(sj)
                sortplayers(sj + 1) = sortplayers(sj)
                sortscores(sj) = stemp
                sortplayers(sj) = stempplayers
            End If
        Next
    Next
    HighScoreTimer.interval = 100
    HighScoreTimer.enabled = True
    ScoreChecker = 4
    CheckAllScores = 1
    NewHighScore sortscores(ScoreChecker), sortplayers(ScoreChecker)
End Sub

'******************
'  Match - Loteria
'******************

Sub IncreaseMatch
    Match = (Match + 10)MOD 100
End Sub

Sub Verification_Match()
   'PlaySound "fx_match"
    Display_Match
    If(Score(CurrentPlayer)MOD 100) = Match Then
       ' PlaySound SoundFXDOF("fx_knocker", 112, DOFPulse, DOFContactors)
    PlaySound "fx_knocker"
    DOF 112,DOFPulse
    SolKnocker True
        AddCredits 1
    End If
End Sub

Sub Clear_Match()
  dim vrobj
    MatchReel.SetValue 0
    If B2SOn then
        Controller.B2SSetMatch 0
    end if
  If VRMode = True Then
    For each vrobj in VRBGMatch : vrobj.visible = 0 : Next
  End If
End Sub

Sub Display_Match()
    MatchReel.SetValue 1 + (Match \ 10)
    If B2SOn then
        If Match = 0 then
            Controller.B2SSetMatch 100
        else
            Controller.B2SSetMatch Match
        end if
    end if
  If VRMode = True Then
    FlasherMatch
  End If
End Sub

' *********************************************************************
'                      Drain / Plunger Functions
' *********************************************************************

' has perdido la bola ;-( mira cuantas bolas hay en el tablero.
' si solamente hay una entonces reduce el número de bola y mira si es la última para finalizar el juego
' si hay más de una, significa que hay multiball, entonces continua con la partida
'
Sub Drain_Hit()
    ' destruye la bola
    Drain.DestroyBall
  If JukeboxMode = 0 then NextTrack

    BallsOnPlayfield = BallsOnPlayfield - 1

    ' haz sonar el ruido de la bola
    RandomSoundDrain Drain

    'si hay falta el sistema de tilt se encargará de continuar con la siguente bola/jugador
    If Tilted Then
        StopEndOfBallMode
    End If

    ' si estás jugando y no hay falta
    If(bGameInPLay = True)AND(Tilted = False)Then

        ' ¿está el salva bolas activado?
        If(bBallSaverActive = True)Then

            ' ¿sí?, pues creamos una bola
            CreateNewBall()
        Else
            ' ¿es ésta la última bola en juego?
            If(BallsOnPlayfield = 0)Then
                StopEndOfBallMode
                vpmtimer.addtimer 500, "EndOfBall '" 'hacemos una pequeña pausa anter de continuar con el fin de bola
                Exit Sub
            End If
        End If
    End If
End Sub

Sub swPlungerRest_Hit()
    bBallInPlungerLane = True
End Sub

' La bola ha sido disparada, así que cambiamos la variable, que en esta mesa se usa solo para que el sonido del disparador cambie según hay allí una bola o no
' En otras mesas podrá usarse para poner en marcha un contador para salvar la bola

Sub swPlungerRest_UnHit()
    bBallInPlungerLane = False
End Sub

' *********************************************************************
'               Funciones para la cuenta de los puntos
' *********************************************************************

' Añade puntos al jugador, hace sonar las campanas y actualiza el backdrop

Sub AddScore(Points)
    If Tilted Then Exit Sub
  IncreaseMatch
    Select Case Points
        Case 10, 100, 1000, 10000
            ' añade los puntos a la variable del actual jugador
            Score(CurrentPlayer) = Score(CurrentPlayer) + points
            ' actualiza los contadores
            UpdateScore points
            ' hace sonar las campanillas de acuerdo a los puntos obtenidos
            If Points = 1000 AND(Score(CurrentPlayer)MOD 10000) \ 1000 = 0 Then  'nuevo reel de 10000
                PlaySound SoundFXDOF("fx_bell10000", 114, DOFPulse, DOFContactors)
            ElseIf Points = 100 AND(Score(CurrentPlayer)MOD 1000) \ 100 = 0 Then 'nuevo reel de 1000
                PlaySound SoundFXDOF("fx_bell1000", 115, DOFPulse, DOFContactors)
            ElseIf Points = 10 AND(Score(CurrentPlayer)MOD 100) \ 10 = 0 Then    'nuevo reel de 100
                PlaySound SoundFXDOF("fx_bell100", 116, DOFPulse, DOFContactors)
            Else
                PlaySound  SoundFXDOF("fx_bell", 111, DOFPulse, DOFContactors) &Points
            End If
        Case 50
            Add10 = Add10 + 5
            AddScore10Timer.Enabled = TRUE
        Case 300
            Add100 = Add100 + 3
            AddScore100Timer.Enabled = TRUE
        Case 500
            Add100 = Add100 + 5
            AddScore100Timer.Enabled = TRUE
        Case 2000, 3000, 4000, 5000
            Add1000 = Add1000 + Points \ 1000
            AddScore1000Timer.Enabled = TRUE
        Case 20000, 30000, 40000, 50000
            Add10000 = Add10000 + Points \ 10000
            AddScore10000Timer.Enabled = TRUE
    End Select

    ' ' aquí se puede hacer un chequeo si el jugador ha ganado alguna puntuación alta y darle un crédito ó bola extra
    If Score(CurrentPlayer) >= Special1 AND Special1Awarded(CurrentPlayer) = False Then
        AwardSpecial
        Special1Awarded(CurrentPlayer) = True
    End If
    If Score(CurrentPlayer) >= Special2 AND Special2Awarded(CurrentPlayer) = False Then
        AwardSpecial
        Special2Awarded(CurrentPlayer) = True
    End If
    If Score(CurrentPlayer) >= Special3 AND Special3Awarded(CurrentPlayer) = False Then
        AwardSpecial
        Special3Awarded(CurrentPlayer) = True
    End If
  If VRMode = True Then
    EMMODE = 1
'   UpdateVRReels 0,0 ,score(0), 0, -1,-1,-1,-1,-1
    UpdateVRReels 1,1 ,score(CurrentPlayer), 0, -1,-1,-1,-1,-1,-1
'   UpdateVRReels 2,2 ,score(2), 0, -1,-1,-1,-1,-1
'   UpdateVRReels 3,3 ,score(3), 0, -1,-1,-1,-1,-1
'   UpdateVRReels 4,4 ,score(4), 0, -1,-1,-1,-1,-1
    EMMODE = 0 ' restore EM mode
  End If
End Sub

'************************************
'TIMER DE 10, 100, 1000 y 1000 PUNTOS
'************************************

' hace sonar las campanillas según los puntos

Sub AddScore10Timer_Timer()
    if Add10 > 0 then
        AddScore 10
        Add10 = Add10 - 1
    Else
        Me.Enabled = FALSE
    End If
End Sub

Sub AddScore100Timer_Timer()
    if Add100 > 0 then
        AddScore 100
        Add100 = Add100 - 1
    Else
        Me.Enabled = FALSE
    End If
End Sub

Sub AddScore1000Timer_Timer()
    if Add1000 > 0 then
        AddScore 1000
        Add1000 = Add1000 - 1
    Else
        Me.Enabled = FALSE
    End If
End Sub

Sub AddScore10000Timer_Timer()
    if Add10000 > 0 then
        AddScore 10000
        Add10000 = Add10000 - 1
    Else
        Me.Enabled = FALSE
    End If
End Sub

'*******************
'     BONOS
'*******************

' avanza el bono y actualiza las luces
' Los bonos están limitados a 1500 puntos

Sub AddBonus(bonuspoints)
    If(Tilted = False)Then
        ' añade los bonos al jugador actual
        Bonus = Bonus + bonuspoints
        If Bonus > 10 Then
            Bonus = 10
        End If
        ' actualiza las luces
        UpdateBonusLights
    End if
End Sub

Sub AddBonusMultiplier(multi)
    If(Tilted = False)Then
        ' añade los bonos al jugador actual
        BonusMultiplier = BonusMultiplier + multi
        If BonusMultiplier > MaxMultiplier Then
            BonusMultiplier = MaxMultiplier
        End If
        ' actualiza las luces
        UpdateMultiplierLights
    End if
End Sub

Sub UpdateMultiplierLights
    Select Case BonusMultiplier
        Case 1:li2.State = 1:li3.State = 0:li4.State = 0
        Case 2:li2.State = 0:li3.State = 1:li4.State = 0
        Case 3:li2.State = 0:li3.State = 0:li4.State = 1
    End Select
End Sub

'***********************************************************************************
'        Score reels - puntuaciones - y actualiza otras luces del backdrop
'***********************************************************************************
'esta es al rutina que actualiza la puntuación del jugador

Sub UpdateScore(playerpoints)
    Select Case CurrentPlayer
        Case 1:ScoreReel1.Addvalue playerpoints
    ' Case 2:ScoreReel2.Addvalue playerpoints
    ' Case 3:ScoreReel3.Addvalue playerpoints
    ' Case 4:ScoreReel4.Addvalue playerpoints
    End Select
    If B2SOn then
        Controller.B2SSetScorePlayer CurrentPlayer, Score(CurrentPlayer)
        If Score(CurrentPlayer) >= 100000 then
            Controller.B2SSetScoreRollover 24 + CurrentPlayer, 1
        end if
    end if
End Sub

' pone todos los marcadores a 0
Sub ResetScores
    ScoreReel1.ResetToZero

    If B2SOn then
        Controller.B2SSetScorePlayer1 0
        Controller.B2SSetScoreRolloverPlayer1 0
    end if
End Sub

Sub AddCredits(value)
    If Credits < 9 Then
        Credits = Credits + value
        CreditsReel.SetValue Credits
        UpdateCredits
    end if
End Sub

Sub UpdateCredits
    If Credits > 0 Then 'in las mesas de Bally
    'CreditLight.State = 1
    Else
    'CreditLight.State = 0
    End If
    PlaySound "fx_relay"
    CreditsReel.SetValue credits
    If B2SOn then
        Controller.B2SSetCredits Credits
    end if
  If VRMode = true Then
    cred =reels(4, 0)
    reels(4, 0) = 0
    SetDrum -1,0,  0
    SetReel 0,-1,  Credits
    reels(4, 0) = Credits

  End If
End Sub

Sub UpdateBallInPlay 'actualiza los marcadores de las bolas, el número de jugador y el número total de jugadores
    BallDisplay.SetValue Balls
    If B2SOn then
        Controller.B2SSetBallInPlay Balls
    end if
  If VRMode = True Then
    FlasherBalls
  End If
End Sub


'*************************
' Partidas y bolas extras
'*************************

Sub AwardExtraBall() 'en este juego el número de bola retrocede
    If(BallsRemaining(CurrentPlayer) <= BallsPerGame)Then
        PlaySound "fx_knocker"
    DOF 112,DOFPulse
    SolKnocker True
        BallsRemaining(CurrentPlayer) = BallsRemaining(CurrentPlayer) + 1
        UpdateBallInPlay
    End If
End Sub

Sub AwardSpecial()
    'PlaySound SoundFXDOF("fx_knocker", 112, DOFPulse, DOFContactors)
  PlaySound "fx_knocker"
  DOF 112,DOFPulse
  SolKnocker True
    AddCredits 1
End Sub


'*******************************************
' ZSOL: Other Solenoids
'*******************************************

' Knocker (this sub mimics how you would handle kicker in ROM based tables)
' For this to work, you must create a primitive on the table named KnockerPosition
' SolCallback(XX) = "SolKnocker"  'In ROM based tables, change the solenoid number XX to the correct number for your table.
Sub SolKnocker(Enabled) 'Knocker solenoid
  If enabled Then
    KnockerSolenoid 'Add knocker position object
  End If
End Sub


' ********************************
'        Attract Mode
' ********************************
' las luces simplemente parpadean de acuerdo a los valores que hemos puesto en el "Blink Pattern" de cada luz

Sub StartAttractMode()
    Dim x
    bAttractMode = True
    For each x in aLights
        x.State = 2
    Next
    ' enciente la luz de fin de partida, bola en juego, ++
    GameOverR.SetValue 1
    BallDisplay.SetValue 0
End Sub

Sub StopAttractMode()
    Dim x
    bAttractMode = False
    TurnOffPlayfieldLights
    ResetScores
    ' apaga la luz de fin de partida
    GameOverR.SetValue 0
End Sub

'************************************************
'    Load (cargar) / Save (guardar)/ Highscore
'************************************************

Sub Loadhs
    ' Based on Black's Highscore routines
    Dim FileObj
    Dim ScoreFile, TextStr
    Dim temp1
    Dim temp2
    Dim temp3
    Dim temp4
    Dim temp5
    Dim temp6
    Dim temp8
    Dim temp9
    Dim temp10
    Dim temp11
    Dim temp12
    Dim temp13
    Dim temp14
    Dim temp15
    Dim temp16
    Dim temp17

    Set FileObj = CreateObject("Scripting.FileSystemObject")
    If Not FileObj.FolderExists(UserDirectory)then
        Credits = 0
        Exit Sub
    End If
    If Not FileObj.FileExists(UserDirectory & TableName& ".txt")then
        Credits = 0
        Exit Sub
    End If
    Set ScoreFile = FileObj.GetFile(UserDirectory & TableName& ".txt")
    Set TextStr = ScoreFile.OpenAsTextStream(1, 0)
    If(TextStr.AtEndOfStream = True)then
        Exit Sub
    End If
    temp1 = TextStr.ReadLine
    temp2 = textstr.readline

    HighScore = cdbl(temp1)
    If HighScore < 1 then
        temp8 = textstr.readline
        temp9 = textstr.readline
        temp10 = textstr.readline
        temp11 = textstr.readline
        temp12 = textstr.readline
        temp13 = textstr.readline
        temp14 = textstr.readline
        temp15 = textstr.readline
        temp16 = textstr.readline
        temp17 = textstr.readline
    End If
    TextStr.Close
    If Credits Then
    Credits = cdbl(temp2)
  Else
    Credits = 0
  End If

    If HighScore < 1 then
        HSScore(1) = int(temp8)
        HSScore(2) = int(temp9)
        HSScore(3) = int(temp10)
        HSScore(4) = int(temp11)
        HSScore(5) = int(temp12)

        HSName(1) = temp13
        HSName(2) = temp14
        HSName(3) = temp15
        HSName(4) = temp16
        HSName(5) = temp17
    End If
    Set ScoreFile = Nothing
    Set FileObj = Nothing
  SortHighscore 'added to fix a previous error
End Sub

Sub Savehs
    ' Based on Black's Highscore routines
    Dim FileObj
    Dim ScoreFile
    Dim xx
    Set FileObj = CreateObject("Scripting.FileSystemObject")
    If Not FileObj.FolderExists(UserDirectory)then
        Exit Sub
    End If
    Set ScoreFile = FileObj.CreateTextFile(UserDirectory & TableName& ".txt", True)
    ScoreFile.WriteLine 0
    ScoreFile.WriteLine Credits
    For xx = 1 to 5
        scorefile.writeline HSScore(xx)
    Next
    For xx = 1 to 5
        scorefile.writeline HSName(xx)
    Next
    ScoreFile.Close
    Set ScoreFile = Nothing
    Set FileObj = Nothing
End Sub

Sub SortHighscore
    Dim tmp, tmp2, i, j
    For i = 1 to 5
        For j = 1 to 4
            If HSScore(j) <HSScore(j + 1) Then
                tmp = HSScore(j + 1)
                tmp2 = HSName(j + 1)
                HSScore(j + 1) = HSScore(j)
                HSName(j + 1) = HSName(j)
                HSScore(j) = tmp
                HSName(j) = tmp2
            End If
        Next
    Next
End Sub

'****************************************
' Actualizaciones en tiempo real
'****************************************
' se usa sobre todo para hacer animaciones o sonidos que cambian en tiempo real
' como por ejemplo para sincronizar los flipers, puertas ó molinillos con primitivas

'Sub GameTimer_Timer
 '   RollingUpdate 'actualiza el sonido de la bola rodando
' y también algunas animaciones, sobre todo de primitivas

'End Sub

'***********************************************************************
' *********************************************************************
'            Aquí empieza el código particular a la mesa
' (hasta ahora todas las rutinas han sido muy generales para todas las mesas)
' (y hay muy pocas rutinas que necesitan cambiar de mesa a mesa)
' *********************************************************************
'***********************************************************************

' se inicia las dianas abatibles, primitivas, etc.
' aunque en el VPX no hay muchos objetos que necesitan ser iniciados

Sub VPObjects_Init 'inicia objetos al empezar

    TurnOffPlayfieldLights()
End Sub

' variables de la mesa

Sub Game_Init() 'esta rutina se llama al principio de un nuevo juego

    'Empezar alguna música, si hubiera música en esta mesa

    'iniciar variables, en esta mesa hay muy pocas variables ya que usamos las luces, y el UserValue de las dianas

    'iniciar algún timer

    'Iniciar algunas lucesli11.State = 1
    TurnOffPlayfieldLights()
End Sub

Sub StopEndOfBallMode()     'this sub is called after the last ball is drained

End Sub

Sub ResetNewBallVariables() 'inicia las variable para una bola ó jugador nuevo
    BonusMultiplier = 0:AddBonusMultiplier 1
End Sub

Sub ResetNewBallLights() 'enciende ó apaga las luces para una bola nueva
    LightBumper1.State = 0
    LightBumper2.State = 0
    LightBumper3.State = 0
    ' luces superiores
    li11.State = 1
    li12.State = 1
    li13.State = 1
    li14.State = 1
    ' y las luces de las dianas
    li10.State = 1
    li8.State = 1
    li6.State = 1
    li5.State = 1
    'apaga luces Special
    li25.State = 0
    li26.State = 0
    li9.State = 0
    li7.State = 0
    'y la de contar los bonos

    ' ¿Es ésta la última bola?
    If(BallsRemaining(CurrentPlayer) = 1)Then
        li1.State = 1
    Else
        li1.State = 0
    End If
End Sub

Sub TurnOffPlayfieldLights()
    Dim a
    For each a in aLights
        a.State = 0
    Next
End Sub

' *********************************************************************
'       Eventos de la mesa - choque de la bola contra dianas
'
' En cada diana u objeto con el que la bola choque habrá que hacer:
' - sonar un sonido físico
' - hacer algún movimiento, si es necesario
' - añadir alguna puntuación
' - encender/apagar una luz
' - hacer algún chequeo para ver si el jugador ha completado algo
' *********************************************************************

' la bola choca contra los Slingshots
' hacemos una animación manual de los slingshots usando gomas
Dim LStep, RStep

Sub LeftSlingShot_Slingshot
    If Tilted Then Exit Sub
    'PlaySoundAtBall SoundFXDOF("fx_slingshot", 103, DOFPulse, DOFContactors)
  DOF 103, DOFPulse
  RandomSoundSlingshotLeft Lemk
    LeftSling4.Visible = 1
    Lemk.RotX = 26
    LStep = 0
    LeftSlingShot.TimerEnabled = True
    ' añade algunos puntos
    AddScore 100
' añade algún efecto a la mesa

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
    If Tilted Then Exit Sub
    'PlaySoundAtBall SoundFXDOF("fx_slingshot", 105, DOFPulse, DOFContactors)
  DOF 105, DOFPulse
    RandomSoundSlingshotRight Remk
  RightSling4.Visible = 1
    Remk.RotX = 26
    RStep = 0
    RightSlingShot.TimerEnabled = True
    ' añade algunos puntos
    AddScore 100
' añade algún efecto a la mesa

End Sub

Sub RightSlingShot_Timer
    Select Case RStep
        Case 1:RightSLing4.Visible = 0:RightSLing3.Visible = 1:Remk.RotX = 14
        Case 2:RightSLing3.Visible = 0:RightSLing2.Visible = 1:Remk.RotX = 2
        Case 3:RightSLing2.Visible = 0:Remk.RotX = -20:RightSlingShot.TimerEnabled = 0
    End Select
    RStep = RStep + 1
End Sub

'***********************
'  Gomas de 10 puntos
'***********************

Sub RubberBand4_Hit
    If Tilted Then Exit Sub
    LastSwitchHit = "RubberBand4"
    ' añade 10 puntos
    AddScore 10
End Sub

Sub RubberBand6_Hit
    If Tilted Then Exit Sub
    LastSwitchHit = "RubberBand6"
    ' añade 10 puntos
    AddScore 10
End Sub

Sub RubberBand7_Hit
    If Tilted Then Exit Sub
    LastSwitchHit = "RubberBand7"
    ' añade 10 puntos
    AddScore 10
End Sub

Sub RubberBand8_Hit
    If Tilted Then Exit Sub
    LastSwitchHit = "RubberBand8"
    ' añade 10 puntos
    AddScore 10
End Sub

Sub RubberBand9_Hit
    If Tilted Then Exit Sub
    LastSwitchHit = "RubberBand9"
    ' añade 10 puntos
    AddScore 10
End Sub

'*********
' Bumpers
'*********

Sub Bumper1_Hit
    If Tilted Then Exit Sub
    'PlaySoundAt SoundFXDOF("fx_bumper",107,DOFPulse,DOFContactors), bumper1
  DOF 107,DOFPulse
  RandomSoundBumperTop Bumper1
    ' añade algunos puntos
    AddScore 100 + 900 * LightBumper1.State
    LastSwitchHit = "bumper1"
End Sub

Sub Bumper2_Hit
    If Tilted Then Exit Sub
    'PlaySoundAt SoundFXDOF("fx_bumper",109,DOFPulse,DOFContactors), bumper2
  DOF 109,DOFPulse
  RandomSoundBumperMiddle Bumper2
    ' añade algunos puntos
    AddScore 100 + 900 * LightBumper2.State
    LastSwitchHit = "bumper2"
End Sub

Sub Bumper3_Hit
    If Tilted Then Exit Sub
   'PlaySoundAt SoundFXDOF("fx_bumper",108,DOFPulse,DOFContactors), bumper3
  DOF 108,DOFPulse
    RandomSoundBumperBottom Bumper3

    ' añade algunos puntos
    AddScore 100 + 900 * LightBumper3.State
    LastSwitchHit = "bumper3"
End Sub

'*****************
'     Pasillos
'*****************

' superiores

Sub Trigger9_Hit
    If Tilted Then Exit Sub
    AddScore 10000
    AddBonus 1
    LastSwitchHit = "trigger9"
    'check?
    LightBumper1.State = 1
    LightBumper2.State = 1
    li11.State = 0
    li5.State = 0
    CheckSpecial
End Sub

Sub Trigger10_Hit
    If Tilted Then Exit Sub
    AddScore 10000
    AddBonus 1
    LastSwitchHit = "trigger10"
    'check?
    LightBumper3.State = 1
    li12.State = 0
    li8.State = 0
    CheckSpecial
End Sub

Sub Trigger11_Hit
    If Tilted Then Exit Sub
    AddScore 10000
    AddBonus 1
    LastSwitchHit = "trigger11"
    'check?
    LightBumper3.State = 1
    li13.State = 0
    li10.State = 0
    CheckSpecial
End Sub

Sub Trigger12_Hit
    If Tilted Then Exit Sub
    AddScore 10000
    AddBonus 1
    LastSwitchHit = "trigger12"
    'check?
    LightBumper1.State = 1
    LightBumper2.State = 1
    li14.State = 0
    li6.State = 0
    CheckSpecial
End Sub

' lateral derecha

Sub Trigger8_Hit
    If Tilted Then Exit Sub
    AddScore 10000
    AddBonus 1
    LastSwitchHit = "trigger8"
    'check?
    AddBonusMultiplier 1
End Sub

' outlanes

Sub Trigger1_Hit
    If Tilted Then Exit Sub
    AddScore 10000
    LastSwitchHit = "trigger1"
    'check?
    if li25.state Then AwardSpecial 'las luces se apagaran al colarse la bola
End Sub

Sub Trigger7_Hit
    If Tilted Then Exit Sub
    AddScore 10000
    LastSwitchHit = "trigger7"
    'check?
    if li26.state Then AwardSpecial 'las luces se apagaran al colarse la bola
End Sub

' inlanes

Sub Trigger2_Hit
    If Tilted Then Exit Sub
    AddScore 5000
    LastSwitchHit = "trigger2"
    'check?
    LightBumper1.State = 1
    LightBumper2.State = 1
    li14.State = 0
    li6.State = 0
    CheckSpecial
End Sub

Sub Trigger6_Hit
    If Tilted Then Exit Sub
    AddScore 5000
    LastSwitchHit = "trigger6"
    'check?
    LightBumper1.State = 1
    LightBumper2.State = 1
    li11.State = 0
    li5.State = 0
    CheckSpecial
End Sub

' dentro de la bola cautiva

Sub Trigger3_Hit
    If Tilted Then Exit Sub
    LastSwitchHit = "trigger3"
    AddScore 100
'check?
End Sub

Sub Trigger4_Hit
    If Tilted Then Exit Sub
    LastSwitchHit = "trigger4"
    AddScore 100
'check?
End Sub

Sub Trigger5_Hit
    If Tilted Then Exit Sub
    LastSwitchHit = "trigger5"
    AddScore 100
'check?
End Sub

'************************
'       Dianas
'************************

Sub Target1_hit
    If Tilted Then Exit Sub
    LastSwitchHit = "target1"
    AddScore 5000
'check
End Sub

Sub Target2_hit
    If Tilted Then Exit Sub
    LastSwitchHit = "target2"
    AddScore 5000
'check
End Sub

Sub Target3_hit 'captive ball
    If Tilted Then Exit Sub
    LastSwitchHit = "target3"
    AddScore 5000
    'check
    li1.State = 1 'recount bonus light
    If li9.State Then
        AwardSpecial
        li9.State = 0
        li25.State = 0
        li26.State = 0
    End If
End Sub

Sub Target4_hit
    If Tilted Then Exit Sub
    LastSwitchHit = "target4"
    AddScore 5000
    'check
    If li7.State AND li17.State Then AwardExtraBall
    If li7.State AND li16.State Then AwardExtraBall
    LightBumper3.State = 1
    li12.State = 0
    li8.State = 0
    CheckSpecial
End Sub

'******
' Setas
'******
Dim mushL

Sub Mushroom1_hit 'amarillo
    If Tilted Then Exit Sub
    LastSwitchHit = "ch1"
    AddScore 5000
    AlternaSpecial
  PmushL.transz=7
  mushL=0
  Mushroom1.timerenabled=1
'check
End Sub

sub Mushroom1_timer
  mushL=mushL+1
  Select case mushL
    case 1: PmushL.transz=0
  Mushroom1.timerenabled=0
  end Select
end sub

Sub Mushroom2_hit 'rojo
    If Tilted Then Exit Sub
    LastSwitchHit = "ch2"
    AddScore 5000
    'check
    LightBumper3.State = 1
    li13.State = 0
    li10.State = 0
    CheckSpecial
    AlternaSpecial
  PmushR.transz=7
  mushL=0
  Mushroom2.timerenabled=1
End Sub

sub Mushroom2_timer
  mushL=mushL+1
  Select case mushL
    case 1: PmushR.transz=0
  Mushroom2.timerenabled=0
  end Select
end sub

Sub CheckSpecial
    If(li11.State + li12.State + li13.State + li14.State = 0)Then
        li9.State = 1
        li25.State = 1
    End If
End Sub

Sub AlternaSpecial
    Dim tmp
    tmp = li25.State
    li25.State = li26.State
    li26.State = tmp
End Sub

' ============================================================================================
' GNMOD - Multiple High Score Display and Collection
' jpsalas: changed ramps by flashers
' ============================================================================================

Dim EnteringInitials ' Normally zero, set to non-zero to enter initials
EnteringInitials = False
Dim ScoreChecker
ScoreChecker = 0
Dim CheckAllScores
CheckAllScores = 0
Dim sortscores(4)
Dim sortplayers(4)

Dim PlungerPulled
PlungerPulled = 0

Dim SelectedChar   ' character under the "cursor" when entering initials

Dim HSTimerCount   ' Pass counter For HS timer, scores are cycled by the timer
HSTimerCount = 5   ' Timer is initially enabled, it'll wrap from 5 to 1 when it's displayed

Dim InitialString  ' the string holding the player's initials as they're entered

Dim AlphaString    ' A-Z, 0-9, space (_) and backspace (<)
Dim AlphaStringPos ' pointer to AlphaString, move Forward and backward with flipper keys
AlphaString = "ABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789_<"

Dim HSNewHigh      ' The new score to be recorded

Dim HSScore(5)     ' High Scores read in from config file
Dim HSName(5)      ' High Score Initials read in from config file

' default high scores, remove this when the scores are available from the config file
HSScore(1) = 900000
HSScore(2) = 800000
HSScore(3) = 700000
HSScore(4) = 600000
HSScore(5) = 500000

HSName(1) = "AAA"
HSName(2) = "ZZZ"
HSName(3) = "XXX"
HSName(4) = "ABC"
HSName(5) = "BBB"

Sub HighScoreTimer_Timer
    If EnteringInitials then
        If HSTimerCount = 1 then
            SetHSLine 3, InitialString & MID(AlphaString, AlphaStringPos, 1)
            HSTimerCount = 2
        Else
            SetHSLine 3, InitialString
            HSTimerCount = 1
        End If
    ElseIf bGameInPlay then
        SetHSLine 1, "HIGH SCORE1"
        SetHSLine 2, HSScore(1)
        SetHSLine 3, HSName(1)
        HSTimerCount = 5 ' set so the highest score will show after the game is over
        HighScoreTimer.enabled = false
    ElseIf CheckAllScores then
        NewHighScore sortscores(ScoreChecker), sortplayers(ScoreChecker)
    Else
        ' cycle through high scores
        HighScoreTimer.interval = 2000
        HSTimerCount = HSTimerCount + 1
        If HsTimerCount > 5 then
            HSTimerCount = 1
        End If
        SetHSLine 1, "HIGH SCORE" + FormatNumber(HSTimerCount, 0)
        SetHSLine 2, HSScore(HSTimerCount)
        SetHSLine 3, HSName(HSTimerCount)
    End If
End Sub

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

Sub SetHsLine(LineNo, String)
    Dim Letter
    Dim ThisDigit
    Dim ThisChar
    Dim StrLen
    Dim LetterLine
    Dim Index
    Dim StartHSArray
    Dim EndHSArray
    Dim LetterName
    Dim xFor
    StartHSArray = array(0, 1, 12, 22)
    EndHSArray = array(0, 11, 21, 31)
    StrLen = len(string)
    Index = 1

    For xFor = StartHSArray(LineNo)to EndHSArray(LineNo)
        Eval("HS" &xFor).imageA = GetHSChar(String, Index)
        Index = Index + 1
    Next
End Sub

Sub NewHighScore(NewScore, PlayNum)
    If NewScore > HSScore(5)then
        HighScoreTimer.interval = 500
        HSTimerCount = 1
        AlphaStringPos = 1      ' start with first character "A"
        EnteringInitials = true ' intercept the control keys while entering initials
        InitialString = ""      ' initials entered so far, initialize to empty
        SetHSLine 1, "PLAYER " + FormatNumber(PlayNum, 0)
        SetHSLine 2, "ENTER NAME"
        SetHSLine 3, MID(AlphaString, AlphaStringPos, 1)
        HSNewHigh = NewScore
        AwardSpecial
    End If
    ScoreChecker = ScoreChecker-1
    If ScoreChecker = 0 then
        CheckAllScores = 0
    End If
End Sub

Sub CollectInitials(keycode)
    Dim i
    If keycode = LeftFlipperKey Then
        ' back up to previous character
        AlphaStringPos = AlphaStringPos - 1
        If AlphaStringPos < 1 then
            AlphaStringPos = len(AlphaString) ' handle wrap from beginning to End
            If InitialString = "" then
                ' Skip the backspace If there are no characters to backspace over
                AlphaStringPos = AlphaStringPos - 1
            End If
        End If
        SetHSLine 3, InitialString & MID(AlphaString, AlphaStringPos, 1)
        PlaySound "fx_Previous"
    ElseIf keycode = RightFlipperKey Then
        ' advance to Next character
        AlphaStringPos = AlphaStringPos + 1
        If AlphaStringPos > len(AlphaString)or(AlphaStringPos = len(AlphaString)and InitialString = "")then
            ' Skip the backspace If there are no characters to backspace over
            AlphaStringPos = 1
        End If
        SetHSLine 3, InitialString & MID(AlphaString, AlphaStringPos, 1)
        PlaySound "fx_Next"
    ElseIf keycode = StartGameKey or keycode = PlungerKey Then
        SelectedChar = MID(AlphaString, AlphaStringPos, 1)
        If SelectedChar = "_" then
            InitialString = InitialString & " "
            PlaySound("fx_Esc")
        ElseIf SelectedChar = "<" then
            InitialString = MID(InitialString, 1, len(InitialString)- 1)
            If len(InitialString) = 0 then
                ' If there are no more characters to back over, don't leave the < displayed
                AlphaStringPos = 1
            End If
            PlaySound("fx_Esc")
        Else
            InitialString = InitialString & SelectedChar
            PlaySound("fx_Enter")
        End If
        If len(InitialString) < 3 then
            SetHSLine 3, InitialString & SelectedChar
        End If
    End If
    If len(InitialString) = 3 then
        ' save the score
        For i = 5 to 1 step -1
            If i = 1 or(HSNewHigh > HSScore(i)and HSNewHigh <= HSScore(i - 1))then
                ' Replace the score at this location
                If i < 5 then
                    HSScore(i + 1) = HSScore(i)
                    HSName(i + 1) = HSName(i)
                End If
                EnteringInitials = False
                HSScore(i) = HSNewHigh
                HSName(i) = InitialString
                HSTimerCount = 5
                HighScoreTimer_Timer
                HighScoreTimer.interval = 2000
                PlaySound("fx_Bong")
                Exit Sub
            ElseIf i < 5 then
                ' move the score in this slot down by 1, it's been exceeded by the new score
                HSScore(i + 1) = HSScore(i)
                HSName(i + 1) = HSName(i)
            End If
        Next
    End If
End Sub
' End GNMOD


'******************************************************
'   ZDMP:  RUBBER  DAMPENERS
'******************************************************
' These are data mined bounce curves,
' dialed in with the in-game elasticity as much as possible to prevent angle / spin issues.
' Requires tracking ballspeed to calculate COR

Sub dPosts_Hit(idx)
  RubbersD.dampen ActiveBall
  'TargetBouncer ActiveBall, 1
End Sub

Sub dSleeves_Hit(idx)
  SleevesD.Dampen ActiveBall
  'TargetBouncer ActiveBall, 0.7
End Sub

Dim RubbersD        'frubber
Set RubbersD = New Dampener
RubbersD.name = "Rubbers"
RubbersD.debugOn = False  'shows info in textbox "TBPout"
RubbersD.Print = False    'debug, reports In debugger (In vel, out cor); cor bounce curve (linear)

'for best results, try to match in-game velocity as closely as possible to the desired curve
'   RubbersD.addpoint 0, 0, 0.935   'point# (keep sequential), ballspeed, CoR (elasticity)
RubbersD.addpoint 0, 0, 1.1    'point# (keep sequential), ballspeed, CoR (elasticity)
RubbersD.addpoint 1, 3.77, 0.97
RubbersD.addpoint 2, 5.76, 0.967  'dont take this as gospel. if you can data mine rubber elasticitiy, please help!
RubbersD.addpoint 3, 15.84, 0.874
RubbersD.addpoint 4, 56, 0.64    'there's clamping so interpolate up to 56 at least

Dim SleevesD  'this is just rubber but cut down to 85%...
Set SleevesD = New Dampener
SleevesD.name = "Sleeves"
SleevesD.debugOn = False  'shows info in textbox "TBPout"
SleevesD.Print = False    'debug, reports In debugger (In vel, out cor)
SleevesD.CopyCoef RubbersD, 0.85

'######################### Add new FlippersD Profile
'######################### Adjust these values to increase or lessen the elasticity

Dim FlippersD
Set FlippersD = New Dampener
FlippersD.name = "Flippers"
FlippersD.debugOn = False
FlippersD.Print = False
FlippersD.addpoint 0, 0, 1.1
FlippersD.addpoint 1, 3.77, 0.99
FlippersD.addpoint 2, 6, 0.99

Class Dampener
  Public Print, debugOn   'tbpOut.text
  Public name, Threshold  'Minimum threshold. Useful for Flippers, which don't have a hit threshold.
  Public ModIn, ModOut
  Private Sub Class_Initialize
    ReDim ModIn(0)
    ReDim Modout(0)
  End Sub

  Public Sub AddPoint(aIdx, aX, aY)
    ShuffleArrays ModIn, ModOut, 1
    ModIn(aIDX) = aX
    ModOut(aIDX) = aY
    ShuffleArrays ModIn, ModOut, 0
    If GameTime > 100 Then Report
  End Sub

  Public Sub Dampen(aBall)
    If threshold Then
      If BallSpeed(aBall) < threshold Then Exit Sub
    End If
    Dim RealCOR, DesiredCOR, str, coef
    DesiredCor = LinearEnvelope(cor.ballvel(aBall.id), ModIn, ModOut )
    RealCOR = BallSpeed(aBall) / (cor.ballvel(aBall.id) + 0.0001)
    coef = desiredcor / realcor
    If debugOn Then str = name & " In vel:" & Round(cor.ballvel(aBall.id),2 ) & vbNewLine & "desired cor: " & Round(desiredcor,4) & vbNewLine & _
    "actual cor: " & Round(realCOR,4) & vbNewLine & "ballspeed coef: " & Round(coef, 3) & vbNewLine
    If Print Then Debug.print Round(cor.ballvel(aBall.id),2) & ", " & Round(desiredcor,3)

    aBall.velx = aBall.velx * coef
    aBall.vely = aBall.vely * coef
    If debugOn Then TBPout.text = str
  End Sub

  Public Sub Dampenf(aBall, parm) 'Rubberizer is handle here
    Dim RealCOR, DesiredCOR, str, coef
    DesiredCor = LinearEnvelope(cor.ballvel(aBall.id), ModIn, ModOut )
    RealCOR = BallSpeed(aBall) / (cor.ballvel(aBall.id) + 0.0001)
    coef = desiredcor / realcor
    If Abs(aball.velx) < 2 And aball.vely < 0 And aball.vely >  - 3.75 Then
      aBall.velx = aBall.velx * coef
      aBall.vely = aBall.vely * coef
    End If
  End Sub

  Public Sub CopyCoef(aObj, aCoef) 'alternative addpoints, copy with coef
    Dim x
    For x = 0 To UBound(aObj.ModIn)
      addpoint x, aObj.ModIn(x), aObj.ModOut(x) * aCoef
    Next
  End Sub

  Public Sub Report() 'debug, reports all coords in tbPL.text
    If Not debugOn Then Exit Sub
    Dim a1, a2
    a1 = ModIn
    a2 = ModOut
    Dim str, x
    For x = 0 To UBound(a1)
      str = str & x & ": " & Round(a1(x),4) & ", " & Round(a2(x),4) & vbNewLine
    Next
    TBPout.text = str
  End Sub
End Class



'nfozzy idigstuff


'******************************************************
'  TRACK ALL BALL VELOCITIES
'  FOR RUBBER DAMPENER AND DROP TARGETS
'******************************************************

Dim cor
Set cor = New CoRTracker

Class CoRTracker
  Public ballvel, ballvelx, ballvely

  Private Sub Class_Initialize
    ReDim ballvel(0)
    ReDim ballvelx(0)
    ReDim ballvely(0)
  End Sub

  Public Sub Update() 'tracks in-ball-velocity
    Dim str, b, AllBalls, highestID
    allBalls = GetBalls

    For Each b In allballs
      If b.id >= HighestID Then highestID = b.id
    Next

    If UBound(ballvel) < highestID Then ReDim ballvel(highestID)  'set bounds
    If UBound(ballvelx) < highestID Then ReDim ballvelx(highestID)  'set bounds
    If UBound(ballvely) < highestID Then ReDim ballvely(highestID)  'set bounds

    For Each b In allballs
      ballvel(b.id) = BallSpeed(b)
      ballvelx(b.id) = b.velx
      ballvely(b.id) = b.vely
    Next
  End Sub
End Class

' Note, cor.update must be called in a 10 ms timer. The example table uses the GameTimer for this purpose, but sometimes a dedicated timer call RDampen is used.
'
'Sub RDampen_Timer
' Cor.Update
'End Sub

'******************************************************
'****  END PHYSICS DAMPENERS
'******************************************************


'*******************************************
' ZFLP: Flippers
'*******************************************

Const ReflipAngle = 20

' Flipper Solenoid Callbacks (these subs mimics how you would handle flippers in ROM based tables)
Sub SolLFlipper(Enabled) 'Left flipper solenoid callback
  If Enabled Then
    LF.Fire  'leftflipper.rotatetoend
    DOF 101,DOFOn
    If leftflipper.currentangle < leftflipper.endangle + ReflipAngle Then
      RandomSoundReflipUpLeft LeftFlipper
    Else
      SoundFlipperUpAttackLeft LeftFlipper
      RandomSoundFlipperUpLeft LeftFlipper
    End If
  Else
    LeftFlipper.RotateToStart
    DOF 101,DOFOff
    If LeftFlipper.currentangle < LeftFlipper.startAngle - 5 Then
      RandomSoundFlipperDownLeft LeftFlipper
    End If
    FlipperLeftHitParm = FlipperUpSoundLevel
  End If
End Sub

Sub SolRFlipper(Enabled) 'Right flipper solenoid callback
  If Enabled Then
    RF.Fire 'rightflipper.rotatetoend
    DOF 102,DOFOn
    If rightflipper.currentangle > rightflipper.endangle - ReflipAngle Then
      RandomSoundReflipUpRight RightFlipper
    Else
      SoundFlipperUpAttackRight RightFlipper
      RandomSoundFlipperUpRight RightFlipper
    End If
  Else
    RightFlipper.RotateToStart
    DOF 102,DOFOff
    If RightFlipper.currentangle > RightFlipper.startAngle + 5 Then
      RandomSoundFlipperDownRight RightFlipper
    End If
    FlipperRightHitParm = FlipperUpSoundLevel
  End If
End Sub

' Flipper collide subs
Sub LeftFlipper_Collide(parm)
  CheckLiveCatch ActiveBall, LeftFlipper, LFCount, parm
  LeftFlipperCollide parm
End Sub

Sub RightFlipper_Collide(parm)
  CheckLiveCatch ActiveBall, RightFlipper, RFCount, parm
  RightFlipperCollide parm
End Sub

'Sub FlipperVisualUpdate 'This subroutine updates the flipper shadows and visual primitives
' FlipperLSh.RotZ = LeftFlipper.CurrentAngle
' FlipperRSh.RotZ = RightFlipper.CurrentAngle
' LFLogo.RotZ = LeftFlipper.CurrentAngle
' RFlogo.RotZ = RightFlipper.CurrentAngle
'End Sub


'******************************************************
' ZNFF:  FLIPPER CORRECTIONS by nFozzy
'******************************************************
'
' There are several steps for taking advantage of nFozzy's flipper solution.  At a high level we'll need the following:
' 1. flippers with specific physics settings
' 2. custom triggers for each flipper (TriggerLF, TriggerRF)
' 3. an object or point to tell the script where the tip of the flipper is at rest (EndPointLp, EndPointRp)
' 4. and, special scripting
'
' A common mistake is incorrect flipper length.  A 3-inch flipper with rubbers will be about 3.125 inches long.
' This translates to about 147 vp units.  Therefore, the flipper start radius + the flipper length + the flipper end
' radius should  equal approximately 147 vp units. Another common mistake is is that sometimes the right flipper
' angle was set with a large postive value (like 238 or something). It should be using negative value (like -122).
'
' The following settings are a solid starting point for various eras of pinballs.
' |                    | EM's           | late 70's to mid 80's | mid 80's to early 90's | mid 90's and later |
' | ------------------ | -------------- | --------------------- | ---------------------- | ------------------ |
' | Mass               | 1              | 1                     | 1                      | 1                  |
' | Strength           | 500-1000 (750) | 1400-1600 (1500)      | 2000-2600              | 3200-3300 (3250)   |
' | Elasticity         | 0.88           | 0.88                  | 0.88                   | 0.88               |
' | Elasticity Falloff | 0.15           | 0.15                  | 0.15                   | 0.15               |
' | Fricition          | 0.8-0.9        | 0.9                   | 0.9                    | 0.9                |
' | Return Strength    | 0.11           | 0.09                  | 0.07                   | 0.055              |
' | Coil Ramp Up       | 2.5            | 2.5                   | 2.5                    | 2.5                |
' | Scatter Angle      | 0              | 0                     | 0                      | 0                  |
' | EOS Torque         | 0.3            | 0.3                   | 0.275                  | 0.275              |
' | EOS Torque Angle   | 4              | 4                     | 6                      | 6                  |
'

'******************************************************
' Flippers Polarity (Select appropriate sub based on era)
'******************************************************

Dim LF
Set LF = New FlipperPolarity
Dim RF
Set RF = New FlipperPolarity

InitPolarity

'
''*******************************************
'' Late 70's to early 80's

Sub InitPolarity()
   dim x, a : a = Array(LF, RF)
  for each x in a
    x.AddPt "Ycoef", 0, RightFlipper.Y-65, 1 'disabled
    x.AddPt "Ycoef", 1, RightFlipper.Y-11, 1
    x.enabled = True
    x.TimeDelay = 80
    x.DebugOn=False ' prints some info in debugger

    x.AddPt "Polarity", 0, 0, 0
    x.AddPt "Polarity", 1, 0.05, - 2.7
    x.AddPt "Polarity", 2, 0.33, - 2.7
    x.AddPt "Polarity", 3, 0.37, - 2.7
    x.AddPt "Polarity", 4, 0.41, - 2.7
    x.AddPt "Polarity", 5, 0.45, - 2.7
    x.AddPt "Polarity", 6, 0.576, - 2.7
    x.AddPt "Polarity", 7, 0.66, - 1.8
    x.AddPt "Polarity", 8, 0.743, - 0.5
    x.AddPt "Polarity", 9, 0.81, - 0.5
    x.AddPt "Polarity", 10, 0.88, 0

    x.AddPt "Velocity", 0, 0, 1
    x.AddPt "Velocity", 1, 0.16, 1.06
    x.AddPt "Velocity", 2, 0.41, 1.05
    x.AddPt "Velocity", 3, 0.53, 1 '0.982
    x.AddPt "Velocity", 4, 0.702, 0.968
    x.AddPt "Velocity", 5, 0.95,  0.968
    x.AddPt "Velocity", 6, 1.03, 0.945
  Next

  ' SetObjects arguments: 1: name of object 2: flipper object: 3: Trigger object around flipper
    LF.SetObjects "LF", LeftFlipper, TriggerLF
    RF.SetObjects "RF", RightFlipper, TriggerRF
End Sub

'' Flipper trigger hit subs
'Sub TriggerLF_Hit()
' LF.Addball activeball
'End Sub
'Sub TriggerLF_UnHit()
' LF.PolarityCorrect activeball
'End Sub
'Sub TriggerRF_Hit()
' RF.Addball activeball
'End Sub
'Sub TriggerRF_UnHit()
' RF.PolarityCorrect activeball
'End Sub

'******************************************************
'  FLIPPER CORRECTION FUNCTIONS
'******************************************************

' modified 2023 by nFozzy
' Removed need for 'endpoint' objects
' Added 'createvents' type thing for TriggerLF / TriggerRF triggers.
' Removed AddPt function which complicated setup imo
' made DebugOn do something (prints some stuff in debugger)
'   Otherwise it should function exactly the same as before

Class FlipperPolarity
  Public DebugOn, Enabled
  Private FlipAt    'Timer variable (IE 'flip at 723,530ms...)
  Public TimeDelay    'delay before trigger turns off and polarity is disabled
  Private Flipper, FlipperStart, FlipperEnd, FlipperEndY, LR, PartialFlipCoef
  Private Balls(20), balldata(20)
  Private Name

  Dim PolarityIn, PolarityOut
  Dim VelocityIn, VelocityOut
  Dim YcoefIn, YcoefOut
  Public Sub Class_Initialize
    ReDim PolarityIn(0)
    ReDim PolarityOut(0)
    ReDim VelocityIn(0)
    ReDim VelocityOut(0)
    ReDim YcoefIn(0)
    ReDim YcoefOut(0)
    Enabled = True
    TimeDelay = 50
    LR = 1
    Dim x
    For x = 0 To UBound(balls)
      balls(x) = Empty
      Set Balldata(x) = new SpoofBall
    Next
  End Sub

  Public Sub SetObjects(aName, aFlipper, aTrigger)

    If TypeName(aName) <> "String" Then MsgBox "FlipperPolarity: .SetObjects error: first argument must be a String (And name of Object). Found:" & TypeName(aName) End If
    If TypeName(aFlipper) <> "Flipper" Then MsgBox "FlipperPolarity: .SetObjects error: Second argument must be a flipper. Found:" & TypeName(aFlipper) End If
    If TypeName(aTrigger) <> "Trigger" Then MsgBox "FlipperPolarity: .SetObjects error: third argument must be a trigger. Found:" & TypeName(aTrigger) End If
    If aFlipper.EndAngle > aFlipper.StartAngle Then LR = -1 Else LR = 1 End If
    Name = aName
    Set Flipper = aFlipper
    FlipperStart = aFlipper.x
    FlipperEnd = Flipper.Length * Sin((Flipper.StartAngle / 57.295779513082320876798154814105)) + Flipper.X ' big floats for degree to rad conversion
    FlipperEndY = Flipper.Length * Cos(Flipper.StartAngle / 57.295779513082320876798154814105)*-1 + Flipper.Y

    Dim str
    str = "Sub " & aTrigger.name & "_Hit() : " & aName & ".AddBall ActiveBall : End Sub'"
    ExecuteGlobal(str)
    str = "Sub " & aTrigger.name & "_UnHit() : " & aName & ".PolarityCorrect ActiveBall : End Sub'"
    ExecuteGlobal(str)

  End Sub

  ' Legacy: just no op
  Public Property Let EndPoint(aInput)

  End Property

  Public Sub AddPt(aChooseArray, aIDX, aX, aY) 'Index #, X position, (in) y Position (out)
    Select Case aChooseArray
      Case "Polarity"
        ShuffleArrays PolarityIn, PolarityOut, 1
        PolarityIn(aIDX) = aX
        PolarityOut(aIDX) = aY
        ShuffleArrays PolarityIn, PolarityOut, 0
      Case "Velocity"
        ShuffleArrays VelocityIn, VelocityOut, 1
        VelocityIn(aIDX) = aX
        VelocityOut(aIDX) = aY
        ShuffleArrays VelocityIn, VelocityOut, 0
      Case "Ycoef"
        ShuffleArrays YcoefIn, YcoefOut, 1
        YcoefIn(aIDX) = aX
        YcoefOut(aIDX) = aY
        ShuffleArrays YcoefIn, YcoefOut, 0
    End Select
  End Sub

  Public Sub AddBall(aBall)
    Dim x
    For x = 0 To UBound(balls)
      If IsEmpty(balls(x)) Then
        Set balls(x) = aBall
        Exit Sub
      End If
    Next
  End Sub

  Private Sub RemoveBall(aBall)
    Dim x
    For x = 0 To UBound(balls)
      If TypeName(balls(x) ) = "IBall" Then
        If aBall.ID = Balls(x).ID Then
          balls(x) = Empty
          Balldata(x).Reset
        End If
      End If
    Next
  End Sub

  Public Sub Fire()
    Flipper.RotateToEnd
    processballs
  End Sub

  Public Property Get Pos 'returns % position a ball. For debug stuff.
    Dim x
    For x = 0 To UBound(balls)
      If Not IsEmpty(balls(x) ) Then
        pos = pSlope(Balls(x).x, FlipperStart, 0, FlipperEnd, 1)
      End If
    Next
  End Property

  Public Sub ProcessBalls() 'save data of balls in flipper range
    FlipAt = GameTime
    Dim x
    For x = 0 To UBound(balls)
      If Not IsEmpty(balls(x) ) Then
        balldata(x).Data = balls(x)
      End If
    Next
    PartialFlipCoef = ((Flipper.StartAngle - Flipper.CurrentAngle) / (Flipper.StartAngle - Flipper.EndAngle))
    PartialFlipCoef = abs(PartialFlipCoef-1)
  End Sub
  'Timer shutoff for polaritycorrect
  Private Function FlipperOn()
    If GameTime < FlipAt+TimeDelay Then
      FlipperOn = True
    End If
  End Function

  Public Sub PolarityCorrect(aBall)
    If FlipperOn() Then
      Dim tmp, BallPos, x, IDX, Ycoef
      Ycoef = 1

      'y safety Exit
      If aBall.VelY > -8 Then 'ball going down
        RemoveBall aBall
        Exit Sub
      End If

      'Find balldata. BallPos = % on Flipper
      For x = 0 To UBound(Balls)
        If aBall.id = BallData(x).id And Not IsEmpty(BallData(x).id) Then
          idx = x
          BallPos = PSlope(BallData(x).x, FlipperStart, 0, FlipperEnd, 1)
          If ballpos > 0.65 Then  Ycoef = LinearEnvelope(BallData(x).Y, YcoefIn, YcoefOut)                'find safety coefficient 'ycoef' data
        End If
      Next

      If BallPos = 0 Then 'no ball data meaning the ball is entering and exiting pretty close to the same position, use current values.
        BallPos = PSlope(aBall.x, FlipperStart, 0, FlipperEnd, 1)
        If ballpos > 0.65 Then  Ycoef = LinearEnvelope(aBall.Y, YcoefIn, YcoefOut)                        'find safety coefficient 'ycoef' data
      End If

      'Velocity correction
      If Not IsEmpty(VelocityIn(0) ) Then
        Dim VelCoef
        VelCoef = LinearEnvelope(BallPos, VelocityIn, VelocityOut)

        If partialflipcoef < 1 Then VelCoef = PSlope(partialflipcoef, 0, 1, 1, VelCoef)

        If Enabled Then aBall.Velx = aBall.Velx*VelCoef
        If Enabled Then aBall.Vely = aBall.Vely*VelCoef
      End If

      'Polarity Correction (optional now)
      If Not IsEmpty(PolarityIn(0) ) Then
        Dim AddX
        AddX = LinearEnvelope(BallPos, PolarityIn, PolarityOut) * LR

        If Enabled Then aBall.VelX = aBall.VelX + 1 * (AddX*ycoef*PartialFlipcoef)
      End If
      If DebugOn Then debug.print "PolarityCorrect" & " " & Name & " @ " & GameTime & " " & Round(BallPos*100) & "%" & " AddX:" & Round(AddX,2) & " Vel%:" & Round(VelCoef*100)
    End If
    RemoveBall aBall
  End Sub
End Class

'******************************************************
'  FLIPPER POLARITY AND RUBBER DAMPENER SUPPORTING FUNCTIONS
'******************************************************

' Used for flipper correction and rubber dampeners
Sub ShuffleArray(ByRef aArray, byVal offset) 'shuffle 1d array
  Dim x, aCount
  aCount = 0
  ReDim a(UBound(aArray) )
  For x = 0 To UBound(aArray)   'Shuffle objects in a temp array
    If Not IsEmpty(aArray(x) ) Then
      If IsObject(aArray(x)) Then
        Set a(aCount) = aArray(x)
      Else
        a(aCount) = aArray(x)
      End If
      aCount = aCount + 1
    End If
  Next
  If offset < 0 Then offset = 0
  ReDim aArray(aCount-1+offset)   'Resize original array
  For x = 0 To aCount-1       'set objects back into original array
    If IsObject(a(x)) Then
      Set aArray(x) = a(x)
    Else
      aArray(x) = a(x)
    End If
  Next
End Sub

' Used for flipper correction and rubber dampeners
Sub ShuffleArrays(aArray1, aArray2, offset)
  ShuffleArray aArray1, offset
  ShuffleArray aArray2, offset
End Sub

' Used for flipper correction, rubber dampeners, and drop targets
Function BallSpeed(ball) 'Calculates the ball speed
  BallSpeed = Sqr(ball.VelX^2 + ball.VelY^2 + ball.VelZ^2)
End Function

' Used for flipper correction and rubber dampeners
Function PSlope(Input, X1, Y1, X2, Y2)    'Set up line via two points, no clamping. Input X, output Y
  Dim x, y, b, m
  x = input
  m = (Y2 - Y1) / (X2 - X1)
  b = Y2 - m*X2
  Y = M*x+b
  PSlope = Y
End Function

' Used for flipper correction
Class spoofball
  Public X, Y, Z, VelX, VelY, VelZ, ID, Mass, Radius
  Public Property Let Data(aBall)
    With aBall
      x = .x
      y = .y
      z = .z
      velx = .velx
      vely = .vely
      velz = .velz
      id = .ID
      mass = .mass
      radius = .radius
    End With
  End Property
  Public Sub Reset()
    x = Empty
    y = Empty
    z = Empty
    velx = Empty
    vely = Empty
    velz = Empty
    id = Empty
    mass = Empty
    radius = Empty
  End Sub
End Class

' Used for flipper correction and rubber dampeners
Function LinearEnvelope(xInput, xKeyFrame, yLvl)
  Dim y 'Y output
  Dim L 'Line
  'find active line
  Dim ii
  For ii = 1 To UBound(xKeyFrame)
    If xInput <= xKeyFrame(ii) Then
      L = ii
      Exit For
    End If
  Next
  If xInput > xKeyFrame(UBound(xKeyFrame) ) Then L = UBound(xKeyFrame)    'catch line overrun
  Y = pSlope(xInput, xKeyFrame(L-1), yLvl(L-1), xKeyFrame(L), yLvl(L) )

  If xInput <= xKeyFrame(LBound(xKeyFrame) ) Then Y = yLvl(LBound(xKeyFrame) )     'Clamp lower
  If xInput >= xKeyFrame(UBound(xKeyFrame) ) Then Y = yLvl(UBound(xKeyFrame) )    'Clamp upper

  LinearEnvelope = Y
End Function

'******************************************************
'  FLIPPER TRICKS
'******************************************************
' To add the flipper tricks you must
'  - Include a call to FlipperCradleCollision from within OnBallBallCollision subroutine
'  - Include a call the CheckLiveCatch from the LeftFlipper_Collide and RightFlipper_Collide subroutines

RightFlipper.timerinterval = 1
Rightflipper.timerenabled = True

Sub RightFlipper_timer()
  FlipperTricks LeftFlipper, LFPress, LFCount, LFEndAngle, LFState
  FlipperTricks RightFlipper, RFPress, RFCount, RFEndAngle, RFState
  FlipperNudge RightFlipper, RFEndAngle, RFEOSNudge, LeftFlipper, LFEndAngle
  FlipperNudge LeftFlipper, LFEndAngle, LFEOSNudge,  RightFlipper, RFEndAngle
End Sub

Dim LFEOSNudge, RFEOSNudge

Sub FlipperNudge(Flipper1, Endangle1, EOSNudge1, Flipper2, EndAngle2)
  Dim b
  Dim BOT
  BOT = GetBalls

  If Flipper1.currentangle = Endangle1 And EOSNudge1 <> 1 Then
    EOSNudge1 = 1
    '   debug.print Flipper1.currentangle &" = "& Endangle1 &"--"& Flipper2.currentangle &" = "& EndAngle2
    If Flipper2.currentangle = EndAngle2 Then
      For b = 0 To UBound(BOT)
        If FlipperTrigger(BOT(b).x, BOT(b).y, Flipper1) Then
          'Debug.Print "ball in flip1. exit"
          Exit Sub
        End If
      Next
      For b = 0 To UBound(BOT)
        If FlipperTrigger(BOT(b).x, BOT(b).y, Flipper2) Then
          BOT(b).velx = BOT(b).velx / 1.3
          BOT(b).vely = BOT(b).vely - 0.5
        End If
      Next
    End If
  Else
    If Abs(Flipper1.currentangle) > Abs(EndAngle1) + 30 Then EOSNudge1 = 0
  End If
End Sub


Dim FCCDamping: FCCDamping = 0.4

Sub FlipperCradleCollision(ball1, ball2, velocity)
  if velocity < 0.7 then exit sub   'filter out gentle collisions
    Dim DoDamping, coef
    DoDamping = false
    'Check left flipper
    If LeftFlipper.currentangle = LFEndAngle Then
    If FlipperTrigger(ball1.x, ball1.y, LeftFlipper) OR FlipperTrigger(ball2.x, ball2.y, LeftFlipper) Then DoDamping = true
    End If
    'Check right flipper
    If RightFlipper.currentangle = RFEndAngle Then
    If FlipperTrigger(ball1.x, ball1.y, RightFlipper) OR FlipperTrigger(ball2.x, ball2.y, RightFlipper) Then DoDamping = true
    End If
    If DoDamping Then
    coef = FCCDamping
        ball1.velx = ball1.velx * coef: ball1.vely = ball1.vely * coef: ball1.velz = ball1.velz * coef
        ball2.velx = ball2.velx * coef: ball2.vely = ball2.vely * coef: ball2.velz = ball2.velz * coef
    End If
End Sub


'*****************
' Maths
'*****************

Dim PI
PI = 4 * Atn(1)

Function dSin(degrees)
  dsin = Sin(degrees * Pi / 180)
End Function

Function dCos(degrees)
  dcos = Cos(degrees * Pi / 180)
End Function

Function Atn2(dy, dx)
  If dx > 0 Then
    Atn2 = Atn(dy / dx)
  ElseIf dx < 0 Then
    If dy = 0 Then
      Atn2 = pi
    Else
      Atn2 = Sgn(dy) * (pi - Atn(Abs(dy / dx)))
    End If
  ElseIf dx = 0 Then
    If dy = 0 Then
      Atn2 = 0
    Else
      Atn2 = Sgn(dy) * pi / 2
    End If
  End If
End Function

Function max(a,b)
  If a > b Then
    max = a
  Else
    max = b
  End If
End Function

Function min(a,b)
  If a > b Then
    min = b
  Else
    min = a
  End If
End Function


'*************************************************
'  Check ball distance from Flipper for Rem
'*************************************************

Function Distance(ax,ay,bx,by)
  Distance = Sqr((ax - bx) ^ 2 + (ay - by) ^ 2)
End Function

Function DistancePL(px,py,ax,ay,bx,by) 'Distance between a point and a line where point Is px,py
  DistancePL = Abs((by - ay) * px - (bx - ax) * py + bx * ay - by * ax) / Distance(ax,ay,bx,by)
End Function

Function Radians(Degrees)
  Radians = Degrees * PI / 180
End Function

Function AnglePP(ax,ay,bx,by)
  AnglePP = Atn2((by - ay),(bx - ax)) * 180 / PI
End Function

Function DistanceFromFlipper(ballx, bally, Flipper)
  DistanceFromFlipper = DistancePL(ballx, bally, Flipper.x, Flipper.y, Cos(Radians(Flipper.currentangle + 90)) + Flipper.x, Sin(Radians(Flipper.currentangle + 90)) + Flipper.y)
End Function

Function FlipperTrigger(ballx, bally, Flipper)
  Dim DiffAngle
  DiffAngle = Abs(Flipper.currentangle - AnglePP(Flipper.x, Flipper.y, ballx, bally) - 90)
  If DiffAngle > 180 Then DiffAngle = DiffAngle - 360

  If DistanceFromFlipper(ballx,bally,Flipper) < 48 And DiffAngle <= 90 And Distance(ballx,bally,Flipper.x,Flipper.y) < Flipper.Length Then
    FlipperTrigger = True
  Else
    FlipperTrigger = False
  End If
End Function

'*************************************************
'  End - Check ball distance from Flipper for Rem
'*************************************************

Dim LFPress, RFPress, LFCount, RFCount
Dim LFState, RFState
Dim EOST, EOSA,Frampup, FElasticity,FReturn
Dim RFEndAngle, LFEndAngle

Const FlipperCoilRampupMode = 0 '0 = fast, 1 = medium, 2 = slow (tap passes should work)

LFState = 1
RFState = 1
EOST = leftflipper.eostorque
EOSA = leftflipper.eostorqueangle
Frampup = LeftFlipper.rampup
FElasticity = LeftFlipper.elasticity
FReturn = LeftFlipper.return
Const EOSTnew = 1 'EM's to late 80's
'Const EOSTnew = 0.8 '90's and later
Const EOSAnew = 1
Const EOSRampup = 0
Dim SOSRampup
Select Case FlipperCoilRampupMode
  Case 0
    SOSRampup = 2.5
  Case 1
    SOSRampup = 6
  Case 2
    SOSRampup = 8.5
End Select

Const LiveCatch = 16
Const LiveElasticity = 0.45
Const SOSEM = 0.815
Const EOSReturn = 0.055  'EM's
'   Const EOSReturn = 0.045  'late 70's to mid 80's
' Const EOSReturn = 0.035  'mid 80's to early 90's
'   Const EOSReturn = 0.025  'mid 90's and later

LFEndAngle = Leftflipper.endangle
RFEndAngle = RightFlipper.endangle

Sub FlipperActivate(Flipper, FlipperPress)
  FlipperPress = 1
  Flipper.Elasticity = FElasticity

  Flipper.eostorque = EOST
  Flipper.eostorqueangle = EOSA
End Sub

Sub FlipperDeactivate(Flipper, FlipperPress)
  FlipperPress = 0
  Flipper.eostorqueangle = EOSA
  Flipper.eostorque = EOST * EOSReturn / FReturn

  If Abs(Flipper.currentangle) <= Abs(Flipper.endangle) + 0.1 Then
    Dim b, BOT
    BOT = GetBalls

    For b = 0 To UBound(BOT)
      If Distance(BOT(b).x, BOT(b).y, Flipper.x, Flipper.y) < 55 Then 'check for cradle
        If BOT(b).vely >= - 0.4 Then BOT(b).vely =  - 0.4
      End If
    Next
  End If
End Sub

Sub FlipperTricks (Flipper, FlipperPress, FCount, FEndAngle, FState)
  Dim Dir
  Dir = Flipper.startangle / Abs(Flipper.startangle) '-1 for Right Flipper

  If Abs(Flipper.currentangle) > Abs(Flipper.startangle) - 0.05 Then
    If FState <> 1 Then
      Flipper.rampup = SOSRampup
      Flipper.endangle = FEndAngle - 3 * Dir
      Flipper.Elasticity = FElasticity * SOSEM
      FCount = 0
      FState = 1
    End If
  ElseIf Abs(Flipper.currentangle) <= Abs(Flipper.endangle) And FlipperPress = 1 Then
    If FCount = 0 Then FCount = GameTime

    If FState <> 2 Then
      Flipper.eostorqueangle = EOSAnew
      Flipper.eostorque = EOSTnew
      Flipper.rampup = EOSRampup
      Flipper.endangle = FEndAngle
      FState = 2
    End If
  ElseIf Abs(Flipper.currentangle) > Abs(Flipper.endangle) + 0.01 And FlipperPress = 1 Then
    If FState <> 3 Then
      Flipper.eostorque = EOST
      Flipper.eostorqueangle = EOSA
      Flipper.rampup = Frampup
      Flipper.Elasticity = FElasticity
      FState = 3
    End If
  End If
End Sub

Const LiveDistanceMin = 30  'minimum distance In vp units from flipper base live catch dampening will occur
Const LiveDistanceMax = 114 'maximum distance in vp units from flipper base live catch dampening will occur (tip protection)

Sub CheckLiveCatch(ball, Flipper, FCount, parm) 'Experimental new live catch
  Dim Dir
  Dir = Flipper.startangle / Abs(Flipper.startangle)  '-1 for Right Flipper
  Dim LiveCatchBounce                                                           'If live catch is not perfect, it won't freeze ball totally
  Dim CatchTime
  CatchTime = GameTime - FCount

  If CatchTime <= LiveCatch And parm > 6 And Abs(Flipper.x - ball.x) > LiveDistanceMin And Abs(Flipper.x - ball.x) < LiveDistanceMax Then
    If CatchTime <= LiveCatch * 0.5 Then                        'Perfect catch only when catch time happens in the beginning of the window
      LiveCatchBounce = 0
    Else
      LiveCatchBounce = Abs((LiveCatch / 2) - CatchTime)    'Partial catch when catch happens a bit late
    End If

    If LiveCatchBounce = 0 And ball.velx * Dir > 0 Then ball.velx = 0
    ball.vely = LiveCatchBounce * (32 / LiveCatch) ' Multiplier for inaccuracy bounce
    ball.angmomx = 0
    ball.angmomy = 0
    ball.angmomz = 0
  Else
    If Abs(Flipper.currentangle) <= Abs(Flipper.endangle) + 1 Then FlippersD.Dampenf ActiveBall, parm
  End If
End Sub

'******************************************************
'****  END FLIPPER CORRECTIONS
'******************************************************



'******************************************************
'   ZFLE:  FLEEP MECHANICAL SOUNDS
'******************************************************

' This part in the script is an entire block that is dedicated to the physics sound system.
' Various scripts and sounds that may be pretty generic and could suit other WPC systems, but the most are tailored specifically for the TOM table

' Many of the sounds in this package can be added by creating collections and adding the appropriate objects to those collections.
' Create the following new collections:
'  Metals (all metal objects, metal walls, metal posts, metal wire guides)
'  Apron (the apron walls and plunger wall)
'  Walls (all wood or plastic walls)
'  Rollovers (wire rollover triggers, star triggers, or button triggers)
'  Targets (standup or drop targets, these are hit sounds only ... you will want to add separate dropping sounds for drop targets)
'  Gates (plate gates)
'  GatesWire (wire gates)
'  Rubbers (all rubbers including posts, sleeves, pegs, and bands)
' When creating the collections, make sure "Fire events for this collection" is checked.
' You'll also need to make sure "Has Hit Event" is checked for each object placed in these collections (not necessary for gates and triggers).
' Once the collections and objects are added, the save, close, and restart VPX.
'
' Many places in the script need to be modified to include the correct sound effect subroutine calls. The tutorial videos linked below demonstrate
' how to make these updates. But in summary the following needs to be updated:
' - Nudging, plunger, coin-in, start button sounds will be added to the keydown and keyup subs.
' - Flipper sounds in the flipper solenoid subs. Flipper collision sounds in the flipper collide subs.
' - Bumpers, slingshots, drain, ball release, knocker, spinner, and saucers in their respective subs
' - Ball rolling sounds sub
'
' Tutorial vides by Apophis
' Audio : Adding Fleep Part 1       https://youtu.be/rG35JVHxtx4
' Audio : Adding Fleep Part 2       https://youtu.be/dk110pWMxGo
' Audio : Adding Fleep Part 3       https://youtu.be/ESXWGJZY_EI


'///////////////////////////////  SOUNDS PARAMETERS  //////////////////////////////
Dim GlobalSoundLevel, CoinSoundLevel, PlungerReleaseSoundLevel, PlungerPullSoundLevel, NudgeLeftSoundLevel
Dim NudgeRightSoundLevel, NudgeCenterSoundLevel, StartButtonSoundLevel, RollingSoundFactor

CoinSoundLevel = 1            'volume level; range [0, 1]
NudgeLeftSoundLevel = 1        'volume level; range [0, 1]
NudgeRightSoundLevel = 1        'volume level; range [0, 1]
NudgeCenterSoundLevel = 1        'volume level; range [0, 1]
StartButtonSoundLevel = 0.1      'volume level; range [0, 1]
PlungerReleaseSoundLevel = 0.8 '1 wjr   'volume level; range [0, 1]
PlungerPullSoundLevel = 1        'volume level; range [0, 1]
RollingSoundFactor = 1.1 / 5

'///////////////////////-----Solenoids, Kickers and Flash Relays-----///////////////////////
Dim FlipperUpAttackMinimumSoundLevel, FlipperUpAttackMaximumSoundLevel, FlipperUpAttackLeftSoundLevel, FlipperUpAttackRightSoundLevel
Dim FlipperUpSoundLevel, FlipperDownSoundLevel, FlipperLeftHitParm, FlipperRightHitParm
Dim SlingshotSoundLevel, BumperSoundFactor, KnockerSoundLevel

FlipperUpAttackMinimumSoundLevel = 0.010    'volume level; range [0, 1]
FlipperUpAttackMaximumSoundLevel = 0.635    'volume level; range [0, 1]
FlipperUpSoundLevel = 1.0            'volume level; range [0, 1]
FlipperDownSoundLevel = 0.45          'volume level; range [0, 1]
FlipperLeftHitParm = FlipperUpSoundLevel    'sound helper; not configurable
FlipperRightHitParm = FlipperUpSoundLevel    'sound helper; not configurable
SlingshotSoundLevel = 0.95            'volume level; range [0, 1]
BumperSoundFactor = 4.25            'volume multiplier; must not be zero
KnockerSoundLevel = 1              'volume level; range [0, 1]

'///////////////////////-----Ball Drops, Bumps and Collisions-----///////////////////////
Dim RubberStrongSoundFactor, RubberWeakSoundFactor, RubberFlipperSoundFactor,BallWithBallCollisionSoundFactor
Dim BallBouncePlayfieldSoftFactor, BallBouncePlayfieldHardFactor, PlasticRampDropToPlayfieldSoundLevel, WireRampDropToPlayfieldSoundLevel, DelayedBallDropOnPlayfieldSoundLevel
Dim WallImpactSoundFactor, MetalImpactSoundFactor, SubwaySoundLevel, SubwayEntrySoundLevel, ScoopEntrySoundLevel
Dim SaucerLockSoundLevel, SaucerKickSoundLevel

BallWithBallCollisionSoundFactor = 3.2      'volume multiplier; must not be zero
RubberStrongSoundFactor = 0.055 / 5      'volume multiplier; must not be zero
RubberWeakSoundFactor = 0.075 / 5        'volume multiplier; must not be zero
RubberFlipperSoundFactor = 0.075 / 5      'volume multiplier; must not be zero
BallBouncePlayfieldSoftFactor = 0.025      'volume multiplier; must not be zero
BallBouncePlayfieldHardFactor = 0.025      'volume multiplier; must not be zero
DelayedBallDropOnPlayfieldSoundLevel = 0.8    'volume level; range [0, 1]
WallImpactSoundFactor = 0.075          'volume multiplier; must not be zero
MetalImpactSoundFactor = 0.075 / 3
SaucerLockSoundLevel = 0.8
SaucerKickSoundLevel = 0.8

'///////////////////////-----Gates, Spinners, Rollovers and Targets-----///////////////////////

Dim GateSoundLevel, TargetSoundFactor, SpinnerSoundLevel, RolloverSoundLevel, DTSoundLevel

GateSoundLevel = 0.5 / 5      'volume level; range [0, 1]
TargetSoundFactor = 0.0025 * 10  'volume multiplier; must not be zero
DTSoundLevel = 0.25        'volume multiplier; must not be zero
RolloverSoundLevel = 0.25      'volume level; range [0, 1]
SpinnerSoundLevel = 0.5      'volume level; range [0, 1]

'///////////////////////-----Ball Release, Guides and Drain-----///////////////////////
Dim DrainSoundLevel, BallReleaseSoundLevel, BottomArchBallGuideSoundFactor, FlipperBallGuideSoundFactor

DrainSoundLevel = 0.8          'volume level; range [0, 1]
BallReleaseSoundLevel = 1        'volume level; range [0, 1]
BottomArchBallGuideSoundFactor = 0.2  'volume multiplier; must not be zero
FlipperBallGuideSoundFactor = 0.015  'volume multiplier; must not be zero

'///////////////////////-----Loops and Lanes-----///////////////////////
Dim ArchSoundFactor
ArchSoundFactor = 0.025 / 5      'volume multiplier; must not be zero

'/////////////////////////////  SOUND PLAYBACK FUNCTIONS  ////////////////////////////
'/////////////////////////////  POSITIONAL SOUND PLAYBACK METHODS  ////////////////////////////
' Positional sound playback methods will play a sound, depending on the X,Y position of the table element or depending on ActiveBall object position
' These are similar subroutines that are less complicated to use (e.g. simply use standard parameters for the PlaySound call)
' For surround setup - positional sound playback functions will fade between front and rear surround channels and pan between left and right channels
' For stereo setup - positional sound playback functions will only pan between left and right channels
' For mono setup - positional sound playback functions will not pan between left and right channels and will not fade between front and rear channels

' PlaySound full syntax - PlaySound(string, int loopcount, float volume, float pan, float randompitch, int pitch, bool useexisting, bool restart, float front_rear_fade)
' Note - These functions will not work (currently) for walls/slingshots as these do not feature a simple, single X,Y position
Sub PlaySoundAtLevelStatic(playsoundparams, aVol, tableobj)
  PlaySound playsoundparams, 0, min(aVol,1) * VolumeDial, AudioPan(tableobj), 0, 0, 0, 0, AudioFade(tableobj)
End Sub

Sub PlaySoundAtLevelExistingStatic(playsoundparams, aVol, tableobj)
  PlaySound playsoundparams, 0, min(aVol,1) * VolumeDial, AudioPan(tableobj), 0, 0, 1, 0, AudioFade(tableobj)
End Sub

Sub PlaySoundAtLevelStaticLoop(playsoundparams, aVol, tableobj)
  PlaySound playsoundparams, - 1, min(aVol,1) * VolumeDial, AudioPan(tableobj), 0, 0, 0, 0, AudioFade(tableobj)
End Sub

Sub PlaySoundAtLevelStaticRandomPitch(playsoundparams, aVol, randomPitch, tableobj)
  PlaySound playsoundparams, 0, min(aVol,1) * VolumeDial, AudioPan(tableobj), randomPitch, 0, 0, 0, AudioFade(tableobj)
End Sub

Sub PlaySoundAtLevelActiveBall(playsoundparams, aVol)
  PlaySound playsoundparams, 0, min(aVol,1) * VolumeDial, AudioPan(ActiveBall), 0, 0, 0, 0, AudioFade(ActiveBall)
End Sub

Sub PlaySoundAtLevelExistingActiveBall(playsoundparams, aVol)
  PlaySound playsoundparams, 0, min(aVol,1) * VolumeDial, AudioPan(ActiveBall), 0, 0, 1, 0, AudioFade(ActiveBall)
End Sub

Sub PlaySoundAtLeveTimerActiveBall(playsoundparams, aVol, ballvariable)
  PlaySound playsoundparams, 0, min(aVol,1) * VolumeDial, AudioPan(ballvariable), 0, 0, 0, 0, AudioFade(ballvariable)
End Sub

Sub PlaySoundAtLevelTimerExistingActiveBall(playsoundparams, aVol, ballvariable)
  PlaySound playsoundparams, 0, min(aVol,1) * VolumeDial, AudioPan(ballvariable), 0, 0, 1, 0, AudioFade(ballvariable)
End Sub

Sub PlaySoundAtLevelRoll(playsoundparams, aVol, pitch)
  PlaySound playsoundparams, - 1, min(aVol,1) * VolumeDial, AudioPan(tableobj), randomPitch, 0, 0, 0, AudioFade(tableobj)
End Sub

' Previous Positional Sound Subs

Sub PlaySoundAt(soundname, tableobj)
  PlaySound soundname, 1, 1 * VolumeDial, AudioPan(tableobj), 0,0,0, 1, AudioFade(tableobj)
End Sub

Sub PlaySoundAtVol(soundname, tableobj, aVol)
  PlaySound soundname, 1, min(aVol,1) * VolumeDial, AudioPan(tableobj), 0,0,0, 1, AudioFade(tableobj)
End Sub

Sub PlaySoundAtBall(soundname)
  PlaySoundAt soundname, ActiveBall
End Sub

Sub PlaySoundAtBallVol (Soundname, aVol)
  PlaySound soundname, 1,min(aVol,1) * VolumeDial, AudioPan(ActiveBall), 0,0,0, 1, AudioFade(ActiveBall)
End Sub

Sub PlaySoundAtBallVolM (Soundname, aVol)
  PlaySound soundname, 1,min(aVol,1) * VolumeDial, AudioPan(ActiveBall), 0,0,0, 0, AudioFade(ActiveBall)
End Sub

Sub PlaySoundAtVolLoops(sound, tableobj, Vol, Loops)
  PlaySound sound, Loops, Vol * VolumeDial, AudioPan(tableobj), 0,0,0, 1, AudioFade(tableobj)
End Sub

'******************************************************
'  Fleep  Supporting Ball & Sound Functions
'******************************************************

Function AudioFade(tableobj) ' Fades between front and back of the table (for surround systems or 2x2 speakers, etc), depending on the Y position on the table. "table1" is the name of the table
  Dim tmp
  tmp = tableobj.y * 2 / tableheight - 1

  If tmp > 7000 Then
    tmp = 7000
  ElseIf tmp <  - 7000 Then
    tmp =  - 7000
  End If

' Thalamus, AudioFade - Patched
	If tmp > 0 Then
		AudioFade = CSng(tmp ^ 5) 'was 10
	Else
		AudioFade = CSng( - (( - tmp) ^ 5) ) 'was 10
	End If

End Function

Function AudioPan(tableobj) ' Calculates the pan for a tableobj based on the X position on the table. "table1" is the name of the table
  Dim tmp
  tmp = tableobj.x * 2 / tablewidth - 1

  If tmp > 7000 Then
    tmp = 7000
  ElseIf tmp <  - 7000 Then
    tmp =  - 7000
  End If

' Thalamus, AudioPan - Patched
	If tmp > 0 Then
		AudioPan = CSng(tmp ^ 5) 'was 10
	Else
		AudioPan = CSng( - (( - tmp) ^ 5) ) 'was 10
	End If

End Function

Function Vol(ball) ' Calculates the volume of the sound based on the ball speed
  Vol = CSng(BallVel(ball) ^ 2)
End Function

Function Volz(ball) ' Calculates the volume of the sound based on the ball speed
  Volz = CSng((ball.velz) ^ 2)
End Function

Function Pitch(ball) ' Calculates the pitch of the sound based on the ball speed
  Pitch = BallVel(ball) * 20
End Function

Function BallVel(ball) 'Calculates the ball speed
  BallVel = Int(Sqr((ball.VelX ^ 2) + (ball.VelY ^ 2) ) )
End Function

Function VolPlayfieldRoll(ball) ' Calculates the roll volume of the sound based on the ball speed
  VolPlayfieldRoll = RollingSoundFactor * 0.0005 * CSng(BallVel(ball) ^ 3)
End Function

Function PitchPlayfieldRoll(ball) ' Calculates the roll pitch of the sound based on the ball speed
  PitchPlayfieldRoll = BallVel(ball) ^ 2 * 15
End Function

Function RndInt(min, max) ' Sets a random number integer between min and max
  RndInt = Int(Rnd() * (max - min + 1) + min)
End Function

Function RndNum(min, max) ' Sets a random number between min and max
  RndNum = Rnd() * (max - min) + min
End Function

'/////////////////////////////  GENERAL SOUND SUBROUTINES  ////////////////////////////

Sub SoundStartButton()
  PlaySound ("Start_Button"), 0, StartButtonSoundLevel, 0, 0.25
End Sub

Sub SoundNudgeLeft()
  PlaySound ("Nudge_" & Int(Rnd * 2) + 1), 0, NudgeLeftSoundLevel * VolumeDial, - 0.1, 0.25
End Sub

Sub SoundNudgeRight()
  PlaySound ("Nudge_" & Int(Rnd * 2) + 1), 0, NudgeRightSoundLevel * VolumeDial, 0.1, 0.25
End Sub

Sub SoundNudgeCenter()
  PlaySound ("Nudge_" & Int(Rnd * 2) + 1), 0, NudgeCenterSoundLevel * VolumeDial, 0, 0.25
End Sub

Sub SoundPlungerPull()
  PlaySoundAtLevelStatic ("Plunger_Pull_1"), PlungerPullSoundLevel, Plunger
End Sub

Sub SoundPlungerReleaseBall()
  PlaySoundAtLevelStatic ("Plunger_Release_Ball"), PlungerReleaseSoundLevel, Plunger
End Sub

Sub SoundPlungerReleaseNoBall()
  PlaySoundAtLevelStatic ("Plunger_Release_No_Ball"), PlungerReleaseSoundLevel, Plunger
End Sub

'/////////////////////////////  KNOCKER SOLENOID  ////////////////////////////

Sub KnockerSolenoid()
  PlaySoundAtLevelStatic SoundFX("Knocker_1",DOFKnocker), KnockerSoundLevel, KnockerPosition
End Sub

'/////////////////////////////  DRAIN SOUNDS  ////////////////////////////

Sub RandomSoundDrain(drainswitch)
  PlaySoundAtLevelStatic ("Drain_" & Int(Rnd * 11) + 1), DrainSoundLevel, drainswitch
End Sub

'/////////////////////////////  TROUGH BALL RELEASE SOLENOID SOUNDS  ////////////////////////////

Sub RandomSoundBallRelease(drainswitch)
  PlaySoundAtLevelStatic SoundFX("BallRelease" & Int(Rnd * 7) + 1,DOFContactors), BallReleaseSoundLevel, drainswitch
End Sub

'/////////////////////////////  SLINGSHOT SOLENOID SOUNDS  ////////////////////////////

Sub RandomSoundSlingshotLeft(sling)
  PlaySoundAtLevelStatic SoundFX("Sling_L" & Int(Rnd * 10) + 1,DOFContactors), SlingshotSoundLevel, Sling
End Sub

Sub RandomSoundSlingshotRight(sling)
  PlaySoundAtLevelStatic SoundFX("Sling_R" & Int(Rnd * 8) + 1,DOFContactors), SlingshotSoundLevel, Sling
End Sub

'/////////////////////////////  BUMPER SOLENOID SOUNDS  ////////////////////////////

Sub RandomSoundBumperTop(Bump)
  PlaySoundAtLevelStatic SoundFX("Bumpers_Top_" & Int(Rnd * 5) + 1,DOFContactors), Vol(ActiveBall) * BumperSoundFactor, Bump
End Sub

Sub RandomSoundBumperMiddle(Bump)
  PlaySoundAtLevelStatic SoundFX("Bumpers_Middle_" & Int(Rnd * 5) + 1,DOFContactors), Vol(ActiveBall) * BumperSoundFactor, Bump
End Sub

Sub RandomSoundBumperBottom(Bump)
  PlaySoundAtLevelStatic SoundFX("Bumpers_Bottom_" & Int(Rnd * 5) + 1,DOFContactors), Vol(ActiveBall) * BumperSoundFactor, Bump
End Sub

'/////////////////////////////  SPINNER SOUNDS  ////////////////////////////

Sub SoundSpinner(spinnerswitch)
  PlaySoundAtLevelStatic ("Spinner"), SpinnerSoundLevel, spinnerswitch
End Sub

'/////////////////////////////  FLIPPER BATS SOUND SUBROUTINES  ////////////////////////////
'/////////////////////////////  FLIPPER BATS SOLENOID ATTACK SOUND  ////////////////////////////

Sub SoundFlipperUpAttackLeft(flipper)
  FlipperUpAttackLeftSoundLevel = RndNum(FlipperUpAttackMinimumSoundLevel, FlipperUpAttackMaximumSoundLevel)
  PlaySoundAtLevelStatic SoundFX("Flipper_Attack-L01",DOFFlippers), FlipperUpAttackLeftSoundLevel, flipper
End Sub

Sub SoundFlipperUpAttackRight(flipper)
  FlipperUpAttackRightSoundLevel = RndNum(FlipperUpAttackMinimumSoundLevel, FlipperUpAttackMaximumSoundLevel)
  PlaySoundAtLevelStatic SoundFX("Flipper_Attack-R01",DOFFlippers), FlipperUpAttackLeftSoundLevel, flipper
End Sub

'/////////////////////////////  FLIPPER BATS SOLENOID CORE SOUND  ////////////////////////////

Sub RandomSoundFlipperUpLeft(flipper)
  PlaySoundAtLevelStatic SoundFX("Flipper_L0" & Int(Rnd * 9) + 1,DOFFlippers), FlipperLeftHitParm, Flipper
End Sub

Sub RandomSoundFlipperUpRight(flipper)
  PlaySoundAtLevelStatic SoundFX("Flipper_R0" & Int(Rnd * 9) + 1,DOFFlippers), FlipperRightHitParm, Flipper
End Sub

Sub RandomSoundReflipUpLeft(flipper)
  PlaySoundAtLevelStatic SoundFX("Flipper_ReFlip_L0" & Int(Rnd * 3) + 1,DOFFlippers), (RndNum(0.8, 1)) * FlipperUpSoundLevel, Flipper
End Sub

Sub RandomSoundReflipUpRight(flipper)
  PlaySoundAtLevelStatic SoundFX("Flipper_ReFlip_R0" & Int(Rnd * 3) + 1,DOFFlippers), (RndNum(0.8, 1)) * FlipperUpSoundLevel, Flipper
End Sub

Sub RandomSoundFlipperDownLeft(flipper)
  PlaySoundAtLevelStatic SoundFX("Flipper_Left_Down_" & Int(Rnd * 7) + 1,DOFFlippers), FlipperDownSoundLevel, Flipper
End Sub

Sub RandomSoundFlipperDownRight(flipper)
  PlaySoundAtLevelStatic SoundFX("Flipper_Right_Down_" & Int(Rnd * 8) + 1,DOFFlippers), FlipperDownSoundLevel, Flipper
End Sub

'/////////////////////////////  FLIPPER BATS BALL COLLIDE SOUND  ////////////////////////////

Sub LeftFlipperCollide(parm)
  FlipperLeftHitParm = parm / 10
  If FlipperLeftHitParm > 1 Then
    FlipperLeftHitParm = 1
  End If
  FlipperLeftHitParm = FlipperUpSoundLevel * FlipperLeftHitParm
  RandomSoundRubberFlipper(parm)
End Sub

Sub RightFlipperCollide(parm)
  FlipperRightHitParm = parm / 10
  If FlipperRightHitParm > 1 Then
    FlipperRightHitParm = 1
  End If
  FlipperRightHitParm = FlipperUpSoundLevel * FlipperRightHitParm
  RandomSoundRubberFlipper(parm)
End Sub

Sub RandomSoundRubberFlipper(parm)
  PlaySoundAtLevelActiveBall ("Flipper_Rubber_" & Int(Rnd * 7) + 1), parm * RubberFlipperSoundFactor
End Sub

'/////////////////////////////  ROLLOVER SOUNDS  ////////////////////////////

Sub RandomSoundRollover()
  PlaySoundAtLevelActiveBall ("Rollover_" & Int(Rnd * 4) + 1), RolloverSoundLevel
End Sub

Sub Rollovers_Hit(idx)
  RandomSoundRollover
End Sub

'/////////////////////////////  VARIOUS PLAYFIELD SOUND SUBROUTINES  ////////////////////////////
'/////////////////////////////  RUBBERS AND POSTS  ////////////////////////////
'/////////////////////////////  RUBBERS - EVENTS  ////////////////////////////

Sub Rubbers_Hit(idx)
  Dim finalspeed
  finalspeed = Sqr(ActiveBall.velx * ActiveBall.velx + ActiveBall.vely * ActiveBall.vely)
  If finalspeed > 5 Then
    RandomSoundRubberStrong 1
  End If
  If finalspeed <= 5 Then
    RandomSoundRubberWeak()
  End If
End Sub

'/////////////////////////////  RUBBERS AND POSTS - STRONG IMPACTS  ////////////////////////////

Sub RandomSoundRubberStrong(voladj)
  Select Case Int(Rnd * 10) + 1
    Case 1
      PlaySoundAtLevelActiveBall ("Rubber_Strong_1"), Vol(ActiveBall) * RubberStrongSoundFactor * voladj
    Case 2
      PlaySoundAtLevelActiveBall ("Rubber_Strong_2"), Vol(ActiveBall) * RubberStrongSoundFactor * voladj
    Case 3
      PlaySoundAtLevelActiveBall ("Rubber_Strong_3"), Vol(ActiveBall) * RubberStrongSoundFactor * voladj
    Case 4
      PlaySoundAtLevelActiveBall ("Rubber_Strong_4"), Vol(ActiveBall) * RubberStrongSoundFactor * voladj
    Case 5
      PlaySoundAtLevelActiveBall ("Rubber_Strong_5"), Vol(ActiveBall) * RubberStrongSoundFactor * voladj
    Case 6
      PlaySoundAtLevelActiveBall ("Rubber_Strong_6"), Vol(ActiveBall) * RubberStrongSoundFactor * voladj
    Case 7
      PlaySoundAtLevelActiveBall ("Rubber_Strong_7"), Vol(ActiveBall) * RubberStrongSoundFactor * voladj
    Case 8
      PlaySoundAtLevelActiveBall ("Rubber_Strong_8"), Vol(ActiveBall) * RubberStrongSoundFactor * voladj
    Case 9
      PlaySoundAtLevelActiveBall ("Rubber_Strong_9"), Vol(ActiveBall) * RubberStrongSoundFactor * voladj
    Case 10
      PlaySoundAtLevelActiveBall ("Rubber_1_Hard"), Vol(ActiveBall) * RubberStrongSoundFactor * 0.6 * voladj
  End Select
End Sub

'/////////////////////////////  RUBBERS AND POSTS - WEAK IMPACTS  ////////////////////////////

Sub RandomSoundRubberWeak()
  PlaySoundAtLevelActiveBall ("Rubber_" & Int(Rnd * 9) + 1), Vol(ActiveBall) * RubberWeakSoundFactor
End Sub

'/////////////////////////////  WALL IMPACTS  ////////////////////////////

Sub Walls_Hit(idx)
  RandomSoundWall()
End Sub

Sub RandomSoundWall()
  Dim finalspeed
  finalspeed = Sqr(ActiveBall.velx * ActiveBall.velx + ActiveBall.vely * ActiveBall.vely)
  If finalspeed > 16 Then
    Select Case Int(Rnd * 5) + 1
      Case 1
        PlaySoundAtLevelExistingActiveBall ("Wall_Hit_1"), Vol(ActiveBall) * WallImpactSoundFactor
      Case 2
        PlaySoundAtLevelExistingActiveBall ("Wall_Hit_2"), Vol(ActiveBall) * WallImpactSoundFactor
      Case 3
        PlaySoundAtLevelExistingActiveBall ("Wall_Hit_5"), Vol(ActiveBall) * WallImpactSoundFactor
      Case 4
        PlaySoundAtLevelExistingActiveBall ("Wall_Hit_7"), Vol(ActiveBall) * WallImpactSoundFactor
      Case 5
        PlaySoundAtLevelExistingActiveBall ("Wall_Hit_9"), Vol(ActiveBall) * WallImpactSoundFactor
    End Select
  End If
  If finalspeed >= 6 And finalspeed <= 16 Then
    Select Case Int(Rnd * 4) + 1
      Case 1
        PlaySoundAtLevelExistingActiveBall ("Wall_Hit_3"), Vol(ActiveBall) * WallImpactSoundFactor
      Case 2
        PlaySoundAtLevelExistingActiveBall ("Wall_Hit_4"), Vol(ActiveBall) * WallImpactSoundFactor
      Case 3
        PlaySoundAtLevelExistingActiveBall ("Wall_Hit_6"), Vol(ActiveBall) * WallImpactSoundFactor
      Case 4
        PlaySoundAtLevelExistingActiveBall ("Wall_Hit_8"), Vol(ActiveBall) * WallImpactSoundFactor
    End Select
  End If
  If finalspeed < 6 Then
    Select Case Int(Rnd * 3) + 1
      Case 1
        PlaySoundAtLevelExistingActiveBall ("Wall_Hit_4"), Vol(ActiveBall) * WallImpactSoundFactor
      Case 2
        PlaySoundAtLevelExistingActiveBall ("Wall_Hit_6"), Vol(ActiveBall) * WallImpactSoundFactor
      Case 3
        PlaySoundAtLevelExistingActiveBall ("Wall_Hit_8"), Vol(ActiveBall) * WallImpactSoundFactor
    End Select
  End If
End Sub

'/////////////////////////////  METAL TOUCH SOUNDS  ////////////////////////////

Sub RandomSoundMetal()
  PlaySoundAtLevelActiveBall ("Metal_Touch_" & Int(Rnd * 13) + 1), Vol(ActiveBall) * MetalImpactSoundFactor
End Sub

'/////////////////////////////  METAL - EVENTS  ////////////////////////////

Sub Metals_Hit (idx)
  RandomSoundMetal
End Sub

Sub ShooterDiverter_collide(idx)
  RandomSoundMetal
End Sub

'/////////////////////////////  BOTTOM ARCH BALL GUIDE  ////////////////////////////
'/////////////////////////////  BOTTOM ARCH BALL GUIDE - SOFT BOUNCES  ////////////////////////////

Sub RandomSoundBottomArchBallGuide()
  Dim finalspeed
  finalspeed = Sqr(ActiveBall.velx * ActiveBall.velx + ActiveBall.vely * ActiveBall.vely)
  If finalspeed > 16 Then
    PlaySoundAtLevelActiveBall ("Apron_Bounce_" & Int(Rnd * 2) + 1), Vol(ActiveBall) * BottomArchBallGuideSoundFactor
  End If
  If finalspeed >= 6 And finalspeed <= 16 Then
    Select Case Int(Rnd * 2) + 1
      Case 1
        PlaySoundAtLevelActiveBall ("Apron_Bounce_1"), Vol(ActiveBall) * BottomArchBallGuideSoundFactor
      Case 2
        PlaySoundAtLevelActiveBall ("Apron_Bounce_Soft_1"), Vol(ActiveBall) * BottomArchBallGuideSoundFactor
    End Select
  End If
  If finalspeed < 6 Then
    Select Case Int(Rnd * 2) + 1
      Case 1
        PlaySoundAtLevelActiveBall ("Apron_Bounce_Soft_1"), Vol(ActiveBall) * BottomArchBallGuideSoundFactor
      Case 2
        PlaySoundAtLevelActiveBall ("Apron_Medium_3"), Vol(ActiveBall) * BottomArchBallGuideSoundFactor
    End Select
  End If
End Sub

'/////////////////////////////  BOTTOM ARCH BALL GUIDE - HARD HITS  ////////////////////////////

Sub RandomSoundBottomArchBallGuideHardHit()
  PlaySoundAtLevelActiveBall ("Apron_Hard_Hit_" & Int(Rnd * 3) + 1), BottomArchBallGuideSoundFactor * 0.25
End Sub

Sub Apron_Hit (idx)
  If Abs(cor.ballvelx(ActiveBall.id) < 4) And cor.ballvely(ActiveBall.id) > 7 Then
    RandomSoundBottomArchBallGuideHardHit()
  Else
    RandomSoundBottomArchBallGuide
  End If
End Sub

'/////////////////////////////  FLIPPER BALL GUIDE  ////////////////////////////

Sub RandomSoundFlipperBallGuide()
  Dim finalspeed
  finalspeed = Sqr(ActiveBall.velx * ActiveBall.velx + ActiveBall.vely * ActiveBall.vely)
  If finalspeed > 16 Then
    Select Case Int(Rnd * 2) + 1
      Case 1
        PlaySoundAtLevelActiveBall ("Apron_Hard_1"),  Vol(ActiveBall) * FlipperBallGuideSoundFactor
      Case 2
        PlaySoundAtLevelActiveBall ("Apron_Hard_2"),  Vol(ActiveBall) * 0.8 * FlipperBallGuideSoundFactor
    End Select
  End If
  If finalspeed >= 6 And finalspeed <= 16 Then
    PlaySoundAtLevelActiveBall ("Apron_Medium_" & Int(Rnd * 3) + 1),  Vol(ActiveBall) * FlipperBallGuideSoundFactor
  End If
  If finalspeed < 6 Then
    PlaySoundAtLevelActiveBall ("Apron_Soft_" & Int(Rnd * 7) + 1),  Vol(ActiveBall) * FlipperBallGuideSoundFactor
  End If
End Sub

'/////////////////////////////  TARGET HIT SOUNDS  ////////////////////////////

Sub RandomSoundTargetHitStrong()
  PlaySoundAtLevelActiveBall SoundFX("Target_Hit_" & Int(Rnd * 4) + 5,DOFTargets), Vol(ActiveBall) * 0.45 * TargetSoundFactor
End Sub

Sub RandomSoundTargetHitWeak()
  PlaySoundAtLevelActiveBall SoundFX("Target_Hit_" & Int(Rnd * 4) + 1,DOFTargets), Vol(ActiveBall) * TargetSoundFactor
End Sub

Sub PlayTargetSound()
  Dim finalspeed
  finalspeed = Sqr(ActiveBall.velx * ActiveBall.velx + ActiveBall.vely * ActiveBall.vely)
  If finalspeed > 10 Then
    RandomSoundTargetHitStrong()
    RandomSoundBallBouncePlayfieldSoft ActiveBall
  Else
    RandomSoundTargetHitWeak()
  End If
End Sub

Sub Targets_Hit (idx)
  PlayTargetSound
End Sub

'/////////////////////////////  BALL BOUNCE SOUNDS  ////////////////////////////

Sub RandomSoundBallBouncePlayfieldSoft(aBall)
  Select Case Int(Rnd * 9) + 1
    Case 1
      PlaySoundAtLevelStatic ("Ball_Bounce_Playfield_Soft_1"), volz(aBall) * BallBouncePlayfieldSoftFactor, aBall
    Case 2
      PlaySoundAtLevelStatic ("Ball_Bounce_Playfield_Soft_2"), volz(aBall) * BallBouncePlayfieldSoftFactor * 0.5, aBall
    Case 3
      PlaySoundAtLevelStatic ("Ball_Bounce_Playfield_Soft_3"), volz(aBall) * BallBouncePlayfieldSoftFactor * 0.8, aBall
    Case 4
      PlaySoundAtLevelStatic ("Ball_Bounce_Playfield_Soft_4"), volz(aBall) * BallBouncePlayfieldSoftFactor * 0.5, aBall
    Case 5
      PlaySoundAtLevelStatic ("Ball_Bounce_Playfield_Soft_5"), volz(aBall) * BallBouncePlayfieldSoftFactor, aBall
    Case 6
      PlaySoundAtLevelStatic ("Ball_Bounce_Playfield_Hard_1"), volz(aBall) * BallBouncePlayfieldSoftFactor * 0.2, aBall
    Case 7
      PlaySoundAtLevelStatic ("Ball_Bounce_Playfield_Hard_2"), volz(aBall) * BallBouncePlayfieldSoftFactor * 0.2, aBall
    Case 8
      PlaySoundAtLevelStatic ("Ball_Bounce_Playfield_Hard_5"), volz(aBall) * BallBouncePlayfieldSoftFactor * 0.2, aBall
    Case 9
      PlaySoundAtLevelStatic ("Ball_Bounce_Playfield_Hard_7"), volz(aBall) * BallBouncePlayfieldSoftFactor * 0.3, aBall
  End Select
End Sub

Sub RandomSoundBallBouncePlayfieldHard(aBall)
  PlaySoundAtLevelStatic ("Ball_Bounce_Playfield_Hard_" & Int(Rnd * 7) + 1), volz(aBall) * BallBouncePlayfieldHardFactor, aBall
End Sub

'/////////////////////////////  DELAYED DROP - TO PLAYFIELD - SOUND  ////////////////////////////

Sub RandomSoundDelayedBallDropOnPlayfield(aBall)
  Select Case Int(Rnd * 5) + 1
    Case 1
      PlaySoundAtLevelStatic ("Ball_Drop_Playfield_1_Delayed"), DelayedBallDropOnPlayfieldSoundLevel, aBall
    Case 2
      PlaySoundAtLevelStatic ("Ball_Drop_Playfield_2_Delayed"), DelayedBallDropOnPlayfieldSoundLevel, aBall
    Case 3
      PlaySoundAtLevelStatic ("Ball_Drop_Playfield_3_Delayed"), DelayedBallDropOnPlayfieldSoundLevel, aBall
    Case 4
      PlaySoundAtLevelStatic ("Ball_Drop_Playfield_4_Delayed"), DelayedBallDropOnPlayfieldSoundLevel, aBall
    Case 5
      PlaySoundAtLevelStatic ("Ball_Drop_Playfield_5_Delayed"), DelayedBallDropOnPlayfieldSoundLevel, aBall
  End Select
End Sub

'/////////////////////////////  BALL GATES AND BRACKET GATES SOUNDS  ////////////////////////////

Sub SoundPlayfieldGate()
  PlaySoundAtLevelStatic ("Gate_FastTrigger_" & Int(Rnd * 2) + 1), GateSoundLevel, ActiveBall
End Sub

Sub SoundHeavyGate()
  PlaySoundAtLevelStatic ("Gate_2"), GateSoundLevel, ActiveBall
End Sub

Sub Gates_hit(idx)
  SoundHeavyGate
End Sub

Sub GatesWire_hit(idx)
  SoundPlayfieldGate
End Sub

'/////////////////////////////  LEFT LANE ENTRANCE - SOUNDS  ////////////////////////////

Sub RandomSoundLeftArch()
  PlaySoundAtLevelActiveBall ("Arch_L" & Int(Rnd * 4) + 1), Vol(ActiveBall) * ArchSoundFactor
End Sub

Sub RandomSoundRightArch()
  PlaySoundAtLevelActiveBall ("Arch_R" & Int(Rnd * 4) + 1), Vol(ActiveBall) * ArchSoundFactor
End Sub

Sub Arch1_hit()
  If ActiveBall.velx > 1 Then SoundPlayfieldGate
  StopSound "Arch_L1"
  StopSound "Arch_L2"
  StopSound "Arch_L3"
  StopSound "Arch_L4"
End Sub

Sub Arch1_unhit()
  If ActiveBall.velx <  - 8 Then
    RandomSoundRightArch
  End If
End Sub

Sub Arch2_hit()
  If ActiveBall.velx < 1 Then SoundPlayfieldGate
  StopSound "Arch_R1"
  StopSound "Arch_R2"
  StopSound "Arch_R3"
  StopSound "Arch_R4"
End Sub

Sub Arch2_unhit()
  If ActiveBall.velx > 10 Then
    RandomSoundLeftArch
  End If
End Sub

'/////////////////////////////  SAUCERS (KICKER HOLES)  ////////////////////////////

Sub SoundSaucerLock()
  PlaySoundAtLevelStatic ("Saucer_Enter_" & Int(Rnd * 2) + 1), SaucerLockSoundLevel, ActiveBall
End Sub

Sub SoundSaucerKick(scenario, saucer)
  Select Case scenario
    Case 0
      PlaySoundAtLevelStatic SoundFX("Saucer_Empty", DOFContactors), SaucerKickSoundLevel, saucer
    Case 1
      PlaySoundAtLevelStatic SoundFX("Saucer_Kick", DOFContactors), SaucerKickSoundLevel, saucer
  End Select
End Sub

'/////////////////////////////  BALL COLLISION SOUND  ////////////////////////////

Sub OnBallBallCollision(ball1, ball2, velocity)

  FlipperCradleCollision ball1, ball2, velocity

  Dim snd
  Select Case Int(Rnd * 7) + 1
    Case 1
      snd = "Ball_Collide_1"
    Case 2
      snd = "Ball_Collide_2"
    Case 3
      snd = "Ball_Collide_3"
    Case 4
      snd = "Ball_Collide_4"
    Case 5
      snd = "Ball_Collide_5"
    Case 6
      snd = "Ball_Collide_6"
    Case 7
      snd = "Ball_Collide_7"
  End Select

  PlaySound (snd), 0, CSng(velocity) ^ 2 / 200 * BallWithBallCollisionSoundFactor * VolumeDial, AudioPan(ball1), 0, Pitch(ball1), 0, 0, AudioFade(ball1)
End Sub


'///////////////////////////  DROP TARGET HIT SOUNDS  ///////////////////////////

Sub RandomSoundDropTargetReset(obj)
  PlaySoundAtLevelStatic SoundFX("Drop_Target_Reset_" & Int(Rnd * 6) + 1,DOFContactors), 1, obj
End Sub

Sub SoundDropTargetDrop(obj)
  PlaySoundAtLevelStatic ("Drop_Target_Down_" & Int(Rnd * 6) + 1), 200, obj
End Sub

'/////////////////////////////  GI AND FLASHER RELAYS  ////////////////////////////

Const RelayFlashSoundLevel = 0.315  'volume level; range [0, 1];
Const RelayGISoundLevel = 1.05    'volume level; range [0, 1];

Sub Sound_GI_Relay(toggle, obj)
  Select Case toggle
    Case 1
      PlaySoundAtLevelStatic ("Relay_GI_On"), 0.025 * RelayGISoundLevel, obj
    Case 0
      PlaySoundAtLevelStatic ("Relay_GI_Off"), 0.025 * RelayGISoundLevel, obj
  End Select
End Sub

Sub Sound_Flash_Relay(toggle, obj)
  Select Case toggle
    Case 1
      PlaySoundAtLevelStatic ("Relay_Flash_On"), 0.025 * RelayFlashSoundLevel, obj
    Case 0
      PlaySoundAtLevelStatic ("Relay_Flash_Off"), 0.025 * RelayFlashSoundLevel, obj
  End Select
End Sub

'/////////////////////////////////////////////////////////////////
'         End Mechanical Sounds
'/////////////////////////////////////////////////////////////////

'******************************************************
'****  FLEEP MECHANICAL SOUNDS
'******************************************************


'***************************************************************************
' VR Plunger Code
'***************************************************************************
Sub TimerPlunger_Timer
  If PinCab_Plunger.Y < 2207.542 then
       PinCab_Plunger.Y = PinCab_Plunger.Y + 5
  End If
End Sub

Sub TimerPlunger2_Timer
  PinCab_Plunger.Y = 2107.542 + (5* Plunger.Position) -20
End Sub

' ***************************************************************************
'          VR Backglass
' ****************************************************************************


' ***************************************************************************
'          (EM) 1-4 player 5x drums, 1 credit drum CORE CODE
' ****************************************************************************


' ********************* POSITION EM REEL DRUMS ON BACKGLASS *************************

Dim xoff,yoff,zoff,xrot,zscale, xcen,ycen
Dim inx
xoff =472
yoff = 0
zoff =835
xrot = -90

Const USEEMS = 1 ' 1-4 set between 1 to 4 based on number of players

const idx_emp1r1 =0 'player 1
const idx_emp2r1 =6 'player 2
const idx_emp3r1 =12 'player 3
const idx_emp4r1 =18 'player 4
const idx_emp4r6 =24 'credits


Dim BGObjEM(1)
if USEEMS = 1 then
  BGObjEM(0) = Array(emp1r1, emp1r2, emp1r3, emp1r4, emp1r5, emp1r6, _
  Empty,Empty,Empty,Empty,Empty,Empty,_
  Empty,Empty,Empty,Empty,Empty,Empty,_
  Empty,Empty,Empty,Empty,Empty,Empty,_
  emp4r6) ' credits
elseif USEEMS = 2 then
  BGObjEM(0) = Array(emp1r1, emp1r2, emp1r3, emp1r4, emp1r5, _
  emp2r1, emp2r2, emp2r3, emp2r4, emp2r5, _
  Empty,Empty,Empty,Empty,Empty,_
  Empty,Empty,Empty,Empty,Empty,_
  emp4r6) ' credits
elseif USEEMS = 3 then
  BGObjEM(0) = Array(emp1r1, emp1r2, emp1r3, emp1r4, emp1r5, _
  emp2r1, emp2r2, emp2r3, emp2r4, emp2r5, _
  emp3r1, emp3r2, emp3r3, emp3r4, emp3r5, _
  Empty,Empty,Empty,Empty,Empty,_
  emp4r6) ' credits
elseif USEEMS = 4 then
  BGObjEM(0) = Array(emp1r1, emp1r2, emp1r3, emp1r4, emp1r5, _
  emp2r1, emp2r2, emp2r3, emp2r4, emp2r5, _
  emp3r1, emp3r2, emp3r3, emp3r4, emp3r5, _
  emp4r1, emp4r2, emp4r3, emp4r4, emp4r5, _
  emp4r6) ' credits
end If

Sub center_objects_em()
Dim cnt,ii, xx, yy, yfact, xfact, objs
'exit sub
yoff = -150
zscale = 0.0000001
xcen =(960 /2) - (17 / 2)
ycen = (1065 /2 ) + (313 /2)

yfact = -25
xfact = 0

cnt =0
  For Each objs In BGObjEM(0)
  If Not IsEmpty(objs) then
    if objs.name = emp4r6.name then
    yoff = 45 ' credit drum is 60% smaller
    Else
    yoff = 2
    end if

  xx =objs.x

  objs.x = (xoff - xcen) + xx + xfact
  yy = objs.y
  objs.y =yoff

    If yy < 0 then
    yy = yy * -1
    end if

  objs.z = (zoff - ycen) + yy - (yy * zscale) + yfact

  'objs.rotx = xrot
  end if
  cnt = cnt + 1
  Next

end sub



' ********************* UPDATE EM REEL DRUMS CORE LIB *************************

Dim cred,ix, np,npp, reels(6, 7), scores(6,2)

'reset scores to defaults
for np =0 to 6
scores(np,0 ) = 0
scores(np,1 ) = 0
Next

'reset EM drums to defaults
For np =0 to 3
  For  npp =0 to 6
  reels(np, npp) =0 ' default to zero
  Next
Next


Sub SetScore(player, ndx , val)

Dim ncnt

  if player = 5 or player = 6 then
    if val > 0 then
      If(ndx = 0)Then ncnt = val * 10
      If(ndx = 1)Then ncnt = val

      scores(player, 0) = scores(player, 0) + ncnt
    end if
  else
    if val > 0 then

    If(ndx = 0)then ncnt = val * 100000
    If(ndx = 1)then ncnt = val * 10000
    If(ndx = 2)then ncnt = val * 1000
    If(ndx = 3)Then ncnt = val * 100
    If(ndx = 4)Then ncnt = val * 10
    If(ndx = 5)Then ncnt = val

    scores(player, 0) = scores(player, 0) + ncnt
    'scores(player, 0) + ncnt

    end if
  end if
End Sub


Sub SetDrum(player, drum , val)
Dim cnt
Dim objs : objs =BGObjEM(0)

  If val = 0 then
    Select case player
    case -1: ' the credit drum
    If Not IsEmpty(objs(idx_emp4r6)) then
    objs(idx_emp4r6).ObjrotX = 0 ' 285
    'cnt =objs(idx_emp4r6).ObjrotX
    end if
    Case 0:
    Select Case drum
        Case 1: If Not IsEmpty(objs(idx_emp1r1)) then: objs(idx_emp1r1).ObjrotX = 0: end if
        Case 2: If Not IsEmpty(objs(idx_emp1r1+1)) then: objs(idx_emp1r1+1).ObjrotX=0: end if
        Case 3: If Not IsEmpty(objs(idx_emp1r1+2)) then: objs(idx_emp1r1+2).ObjrotX=0: end if
        Case 4: If Not IsEmpty(objs(idx_emp1r1+3)) then: objs(idx_emp1r1+3).ObjrotX=0: end if
        Case 5: If Not IsEmpty(objs(idx_emp1r1+4)) then: objs(idx_emp1r1+4).ObjrotX=0: end if
        Case 6: If Not IsEmpty(objs(idx_emp1r1+5)) then: objs(idx_emp1r1+5).ObjrotX=0: end if
    End Select
    Case 1:
    Select Case drum
        Case 1: If Not IsEmpty(objs(idx_emp2r1)) then: objs(idx_emp2r1).ObjrotX = 0: end if
        Case 2: If Not IsEmpty(objs(idx_emp2r1+1)) then: objs(idx_emp2r1+1).ObjrotX=0: end if
        Case 3: If Not IsEmpty(objs(idx_emp2r1+2)) then: objs(idx_emp2r1+2).ObjrotX=0: end if
        Case 4: If Not IsEmpty(objs(idx_emp2r1+3)) then: objs(idx_emp2r1+3).ObjrotX=0: end if
        Case 5: If Not IsEmpty(objs(idx_emp2r1+4)) then: objs(idx_emp2r1+4).ObjrotX=0: end if
    End Select
    Case 2:
    Select Case drum
        Case 1: If Not IsEmpty(objs(idx_emp3r1)) then: objs(idx_emp3r1).ObjrotX = 0: end if
        Case 2: If Not IsEmpty(objs(idx_emp3r1+1)) then: objs(idx_emp3r1+1).ObjrotX=0: end if
        Case 3: If Not IsEmpty(objs(idx_emp3r1+2)) then: objs(idx_emp3r1+2).ObjrotX=0: end if
        Case 4: If Not IsEmpty(objs(idx_emp3r1+3)) then: objs(idx_emp3r1+3).ObjrotX=0: end if
        Case 5: If Not IsEmpty(objs(idx_emp3r1+4)) then: objs(idx_emp3r1+4).ObjrotX=0: end if
    End Select
    Case 3:
    Select Case drum
        Case 1: If Not IsEmpty(objs(idx_emp4r1)) then: objs(idx_emp4r1).ObjrotX = 0: end if' 283
        Case 2: If Not IsEmpty(objs(idx_emp4r1+1)) then: objs(idx_emp4r1+1).ObjrotX=0: end if
        Case 3: If Not IsEmpty(objs(idx_emp4r1+2)) then: objs(idx_emp4r1+2).ObjrotX=0: end if
        Case 4: If Not IsEmpty(objs(idx_emp4r1+3)) then: objs(idx_emp4r1+3).ObjrotX=0: end if
        Case 5: If Not IsEmpty(objs(idx_emp4r1+4)) then: objs(idx_emp4r1+4).ObjrotX=0: end if
    End Select
  End Select

  else
  Select case player

    Case -1: ' the credit drum
    'emp4r6.ObjrotX = emp4r6.ObjrotX + val
    If Not IsEmpty(objs(idx_emp4r6)) then
    objs(idx_emp4r6).ObjrotX = objs(idx_emp4r6).ObjrotX + val
    end if

    Case 0:
    Select Case drum
        Case 1: If Not IsEmpty(objs(idx_emp1r1)) then: objs(idx_emp1r1).ObjrotX= objs(idx_emp1r1).ObjrotX + val: end if
        Case 2: If Not IsEmpty(objs(idx_emp1r1+1)) then: objs(idx_emp1r1+1).ObjrotX= objs(idx_emp1r1+1).ObjrotX + val: end if
        Case 3: If Not IsEmpty(objs(idx_emp1r1+2)) then: objs(idx_emp1r1+2).ObjrotX= objs(idx_emp1r1+2).ObjrotX + val: end if
        Case 4: If Not IsEmpty(objs(idx_emp1r1+3)) then: objs(idx_emp1r1+3).ObjrotX= objs(idx_emp1r1+3).ObjrotX + val: end if
        Case 5: If Not IsEmpty(objs(idx_emp1r1+4)) then: objs(idx_emp1r1+4).ObjrotX= objs(idx_emp1r1+4).ObjrotX + val: end if
        Case 6: If Not IsEmpty(objs(idx_emp1r1+5)) then: objs(idx_emp1r1+5).ObjrotX= objs(idx_emp1r1+5).ObjrotX + val: end if

    End Select
    Case 1:
    Select Case drum
        Case 1: If Not IsEmpty(objs(idx_emp2r1)) then: objs(idx_emp2r1).ObjrotX= objs(idx_emp2r1).ObjrotX + val: end if
        Case 2: If Not IsEmpty(objs(idx_emp2r1+1)) then: objs(idx_emp2r1+1).ObjrotX= objs(idx_emp2r1+1).ObjrotX + val: end if
        Case 3: If Not IsEmpty(objs(idx_emp2r1+2)) then: objs(idx_emp2r1+2).ObjrotX= objs(idx_emp2r1+2).ObjrotX + val: end if
        Case 4: If Not IsEmpty(objs(idx_emp2r1+3)) then: objs(idx_emp2r1+3).ObjrotX= objs(idx_emp2r1+3).ObjrotX + val: end if
        Case 5: If Not IsEmpty(objs(idx_emp2r1+4)) then: objs(idx_emp2r1+4).ObjrotX= objs(idx_emp2r1+4).ObjrotX + val: end if
    End Select
    Case 2:
    Select Case drum
        Case 1: If Not IsEmpty(objs(idx_emp3r1)) then: objs(idx_emp3r1).ObjrotX= objs(idx_emp3r1).ObjrotX + val: end if
        Case 2: If Not IsEmpty(objs(idx_emp3r1+1)) then: objs(idx_emp3r1+1).ObjrotX= objs(idx_emp3r1+1).ObjrotX + val: end if
        Case 3: If Not IsEmpty(objs(idx_emp3r1+2)) then: objs(idx_emp3r1+2).ObjrotX= objs(idx_emp3r1+2).ObjrotX + val: end if
        Case 4: If Not IsEmpty(objs(idx_emp3r1+3)) then: objs(idx_emp3r1+3).ObjrotX= objs(idx_emp3r1+3).ObjrotX + val: end if
        Case 5: If Not IsEmpty(objs(idx_emp3r1+4)) then: objs(idx_emp3r1+4).ObjrotX= objs(idx_emp3r1+4).ObjrotX + val: end if
    End Select
    Case 3:
    Select Case drum
        Case 1: If Not IsEmpty(objs(idx_emp4r1)) then: objs(idx_emp4r1).ObjrotX= objs(idx_emp4r1).ObjrotX + val: end if
        Case 2: If Not IsEmpty(objs(idx_emp4r1+1)) then: objs(idx_emp4r1+1).ObjrotX= objs(idx_emp4r1+1).ObjrotX + val: end if
        Case 3: If Not IsEmpty(objs(idx_emp4r1+2)) then: objs(idx_emp4r1+2).ObjrotX= objs(idx_emp4r1+2).ObjrotX + val: end if
        Case 4: If Not IsEmpty(objs(idx_emp4r1+3)) then: objs(idx_emp4r1+3).ObjrotX= objs(idx_emp4r1+3).ObjrotX + val: end if
        Case 5: If Not IsEmpty(objs(idx_emp4r1+4)) then: objs(idx_emp4r1+4).ObjrotX= objs(idx_emp4r1+4).ObjrotX + val: end if
    End Select

  End Select
  end if
End Sub


Sub SetReel(player, drum, val)

'TextBox1.text = "playr:" & player +1 & " drum:" & drum & "val:" & val

Dim  inc , cur, dif, fix, fval

inc = 33.5
fval = -5 ' graphic seam between 5 & 6 fix value, easier to fix here than photoshop

If  (player <= 3) or (drum = -1) then

  If drum = -1 then drum = 0

  cur =reels(player, drum)

  If val <> cur then ' something has changed
  Select Case drum

    Case 0: ' credits drum

      if val > cur then
        dif =val - cur
        fix =0
          If cur < 5 and cur+dif > 5 then
          fix = fix- fval
          end if
        dif = dif * inc

        dif = dif-fix

        SetDrum -1,0,  -dif
      Else
        if val = 0 Then
        SetDrum -1,0,  0' reset the drum to abs. zero
        Else
        dif = 11 - cur
        dif = dif + val

        dif = dif * inc
        dif = dif-fval

        SetDrum -1,0,   -dif
        end if
      end if
    Case 1:
    'TB1.text = val
    if val > cur then
      dif =val - cur
      fix =0
        If cur < 6 and cur+dif > 6 then
        fix = fix- fval
        end if
      dif = dif * inc

      dif = dif-fix

      SetDrum player,drum,  -dif
    Else
      if val = 0 Then
      SetDrum player,drum,  0' reset the drum to abs. zero
      Else
      dif = 11 - cur
      dif = dif + val

      dif = dif * inc
      dif = dif-fval

      SetDrum player,drum,   -dif
      end if

    end if
    reels(player, drum) = val

    Case 2:
    'TB2.text = val

    if val > cur then
      dif =val - cur
      fix =0
        If cur < 6 and cur+dif > 6 then
        fix = fix- fval
        end if
      dif = dif * inc
      dif = dif-fix
      SetDrum player,drum,  -dif
    Else
      if val = 0 Then
      SetDrum player,drum,  0 ' reset the drum to abs. zero
      Else
      dif = 11 - cur
      dif = dif + val
      dif = dif * inc
      dif = dif-fval
      SetDrum player,drum,  -dif
      end if
    end if
    reels(player, drum) = val

    Case 3:
    'TB3.text = val

    if val > cur then
      dif =val - cur
      fix =0
        If cur < 6 and cur+dif > 6 then
        fix = fix- fval
        end if
      dif = dif * inc
      dif = dif-fix

      SetDrum player,drum,  -dif
    Else
      if val = 0 Then
      SetDrum player,drum,  0 ' reset the drum to abs. zero
      Else
      dif = 11 - cur
      dif = dif + val
      dif = dif * inc
      dif = dif-fval
      SetDrum player,drum,  -dif
      end if

    end if
    reels(player, drum) = val

    Case 4:
    'TB4.text = val

    if val > cur then
      dif =val - cur
      fix =0
        If cur < 6 and cur+dif > 6 then
        fix = fix- fval
        end if
      dif = dif * inc
      dif = dif-fix
      SetDrum player,drum,  -dif
    Else
      if val = 0 Then
      SetDrum player,drum,  0 ' reset the drum to abs. zero
      Else
      dif = 11 - cur
      dif = dif + val
      dif = dif * inc
      dif = dif-fval
      SetDrum player,drum,  -dif
      end if

    end if
    reels(player, drum) = val

    Case 5:
    'TB5.text = val

    if val > cur then
      dif =val - cur
      fix =0
        If cur < 5 and cur+dif > 5 then
        fix = fix- fval
        end if
      dif = dif * inc
      dif = dif-fix
      SetDrum player,drum,  -dif
    Else
      if val = 0 Then
      SetDrum player,drum,  0 ' reset the drum to abs. zero
      Else
      dif = 11 - cur
      dif = dif + val
      dif = dif * inc
      dif = dif-fval
      SetDrum player,drum,  -dif
      end if

    end if
    Case 6:
    'TB5.text = val

    if val > cur then
      dif =val - cur
      fix =0
        If cur < 5 and cur+dif > 5 then
        fix = fix- fval
        end if
      dif = dif * inc
      dif = dif-fix
      SetDrum player,drum,  -dif
    Else
      if val = 0 Then
      SetDrum player,drum,  0 ' reset the drum to abs. zero
      Else
      dif = 11 - cur
      dif = dif + val
      dif = dif * inc
      dif = dif-fval
      SetDrum player,drum,  -dif
      end if

    end if
    reels(player, drum) = val
   End Select

  end if
end if
End Sub

Dim EMMODE: EMMODE = 0
'Dim Score10,Score100,Score1000,Score10000,Score100000
Dim Score1000,Score10000,Score100000, ActivePLayer
Dim nplayer,playr,value,curscr,curplayr

' EMMODE = 1
'ex: UpdateVRReels 0-3, 0-3, 0-199999, n/a, n/a, n/a, n/a, n/a, n/a
' EMMODE = 0
'ex: UpdateVRReels 0-3, 0-3, n/a, 0-1,0-99999 ,0-9999, 0-999, 0-99, 0-9

Sub UpdateVRReels (Player,nReels ,nScore, n100K, Score100000, Score10000 ,Score1000,Score100,Score10,Score1)

' UpdateVRReels 1,1 ,score(CurrentPlayer), 0, -1,-1,-1,-1,-1,-1

' to-do find out if player is one or zero based, if 1 based subtract 1.
value =nScore'Score(Player)
  nplayer = Player -1

  curscr = value
  curplayr = nplayer


scores(0,1) = scores(0,0)
  scores(0,0) = 0
  scores(1,1) = scores(1,0)
  scores(1,0) = 0
  scores(2,1) = scores(2,0)
  scores(2,0) = 0
  scores(3,1) = scores(3,0)
  scores(3,0) = 0

  For  ix =0 to 6
    reels(0, ix) =0
    reels(1, ix) =0
    reels(2, ix) =0
    reels(3, ix) =0
  Next

  For  ix =0 to 4

  SetDrum ix, 1 , 0
  SetDrum ix, 2 , 0
  SetDrum ix, 3 , 0
  SetDrum ix, 4 , 0
  SetDrum ix, 5 , 0

  Next

  For playr =0 to nReels

    if EMMODE = 0 then
    If (ActivePLayer) = playr Then
    nplayer = playr

    SetReel nplayer, 1 , Score100000 : SetScore nplayer,0,Score100000
    SetReel nplayer, 2 , Score10000 : SetScore nplayer,1,Score10000
    SetReel nplayer, 3 , Score1000 : SetScore nplayer,2,Score1000
    SetReel nplayer, 4 , Score100 : SetScore nplayer,3,Score100
    SetReel nplayer, 5 , Score10 : SetScore nplayer,4,Score10 '
    SetReel nplayer, 6 , 0 : SetScore nplayer,5,0 ' assumes ones position is always zero

    else
    nplayer = playr
    value =scores(nplayer, 1)

  ' do hundred thousands
    if(value >= 900000)  then:  SetReel nplayer, 1 , 9 : SetScore nplayer,0,9 : value = value - 900000: end if
    if(value >= 800000)  then:  SetReel nplayer, 1 , 8 : SetScore nplayer,0,8 : value = value - 800000: end if
    if(value >= 700000)  then:  SetReel nplayer, 1 , 7 : SetScore nplayer,0,7 : value = value - 700000: end if
    if(value >= 600000)  then:  SetReel nplayer, 1 , 6 : SetScore nplayer,0,6 : value = value - 600000: end if
    if(value >= 500000)  then:  SetReel nplayer, 1 , 5 : SetScore nplayer,0,5 : value = value - 500000: end if
    if(value >= 400000)  then:  SetReel nplayer, 1 , 4 : SetScore nplayer,0,4 : value = value - 400000: end if
    if(value >= 300000)  then:  SetReel nplayer, 1 , 3 : SetScore nplayer,0,3 : value = value - 300000: end if
    if(value >= 200000)  then:  SetReel nplayer, 1 , 2 : SetScore nplayer,0,2 : value = value - 200000: end if
    if(value >= 100000)  then:  SetReel nplayer, 1 , 1 : SetScore nplayer,0,1 : value = value - 100000 :end if


  ' do ten thousands
    if(value >= 90000)  then:  SetReel nplayer, 2 , 9 : SetScore nplayer,1,9 : value = value - 90000: end if
    if(value >= 80000)  then:  SetReel nplayer, 2 , 8 : SetScore nplayer,1,8 : value = value - 80000: end if
    if(value >= 70000)  then:  SetReel nplayer, 2 , 7 : SetScore nplayer,1,7 : value = value - 70000: end if
    if(value >= 60000)  then:  SetReel nplayer, 2 , 6 : SetScore nplayer,1,6 : value = value - 60000: end if
    if(value >= 50000)  then:  SetReel nplayer, 2 , 5 : SetScore nplayer,1,5 : value = value - 50000: end if
    if(value >= 40000)  then:  SetReel nplayer, 2 , 4 : SetScore nplayer,1,4 : value = value - 40000: end if
    if(value >= 30000)  then:  SetReel nplayer, 2 , 3 : SetScore nplayer,1,3 : value = value - 30000: end if
    if(value >= 20000)  then:  SetReel nplayer, 2 , 2 : SetScore nplayer,1,2 : value = value - 20000: end if
    if(value >= 10000)  then:  SetReel nplayer, 2 , 1 : SetScore nplayer,1,1 : value = value - 10000: end if


  ' do thousands
    if(value >= 9000)  then:  SetReel nplayer, 3 , 9 : SetScore nplayer,2,9 : value = value - 9000: end if
    if(value >= 8000)  then:  SetReel nplayer, 3 , 8 : SetScore nplayer,2,8 : value = value - 8000: end if
    if(value >= 7000)  then:  SetReel nplayer, 3 , 7 : SetScore nplayer,2,7 : value = value - 7000: end if
    if(value >= 6000)  then:  SetReel nplayer, 3 , 6 : SetScore nplayer,2,6 : value = value - 6000: end if
    if(value >= 5000)  then:  SetReel nplayer, 3 , 5 : SetScore nplayer,2,5 : value = value - 5000: end if
    if(value >= 4000)  then:  SetReel nplayer, 3 , 4 : SetScore nplayer,2,4 : value = value - 4000: end if
    if(value >= 3000)  then:  SetReel nplayer, 3 , 3 : SetScore nplayer,2,3 : value = value - 3000: end if
    if(value >= 2000)  then:  SetReel nplayer, 3 , 2 : SetScore nplayer,2,2 : value = value - 2000: end if
    if(value >= 1000)  then:  SetReel nplayer, 3 , 1 : SetScore nplayer,2,1 : value = value - 1000: end if

    'do hundreds

    if(value >= 900) then: SetReel nplayer, 4 , 9 : SetScore nplayer,3,9 : value = value - 900: end if
    if(value >= 800) then: SetReel nplayer, 4 , 8 : SetScore nplayer,3,8 : value = value - 800: end if
    if(value >= 700) then: SetReel nplayer, 4 , 7 : SetScore nplayer,3,7 : value = value - 700: end if
    if(value >= 600) then: SetReel nplayer, 4 , 6 : SetScore nplayer,3,6 : value = value - 600: end if
    if(value >= 500) then: SetReel nplayer, 4 , 5 : SetScore nplayer,3,5 : value = value - 500: end if
    if(value >= 400) then: SetReel nplayer, 4 , 4 : SetScore nplayer,3,4 : value = value - 400: end if
    if(value >= 300) then: SetReel nplayer, 4 , 3 : SetScore nplayer,3,3 : value = value - 300: end if
    if(value >= 200) then: SetReel nplayer, 4 , 2 : SetScore nplayer,3,2 : value = value - 200: end if
    if(value >= 100) then: SetReel nplayer, 4 , 1 : SetScore nplayer,3,1 : value = value - 100: end if

    'do tens
    if(value >= 90) then: SetReel nplayer, 5 , 9 : SetScore nplayer,4,9 : value = value - 90: end if
    if(value >= 80) then: SetReel nplayer, 5 , 8 : SetScore nplayer,4,8 : value = value - 80: end if
    if(value >= 70) then: SetReel nplayer, 5 , 7 : SetScore nplayer,4,7 : value = value - 70: end if
    if(value >= 60) then: SetReel nplayer, 5 , 6 : SetScore nplayer,4,6 : value = value - 60: end if
    if(value >= 50) then: SetReel nplayer, 5 , 5 : SetScore nplayer,4,5 : value = value - 50: end if
    if(value >= 40) then: SetReel nplayer, 5 , 4 : SetScore nplayer,4,4 : value = value - 40: end if
    if(value >= 30) then: SetReel nplayer, 5 , 3 : SetScore nplayer,4,3 : value = value - 30: end if
    if(value >= 20) then: SetReel nplayer, 5 , 2 : SetScore nplayer,4,2 : value = value - 20: end if
    if(value >= 10) then: SetReel nplayer, 5 , 1 : SetScore nplayer,4,1 : value = value - 10: end if

    'do ones
    if(value >= 9) then: SetReel nplayer, 6 , 9 : SetScore nplayer,5,9 : value = value - 9: end if
    if(value >= 8) then: SetReel nplayer, 6 , 8 : SetScore nplayer,5,8 : value = value - 8: end if
    if(value >= 7) then: SetReel nplayer, 6 , 7 : SetScore nplayer,5,7 : value = value - 7: end if
    if(value >= 6) then: SetReel nplayer, 6 , 6 : SetScore nplayer,5,6 : value = value - 6: end if
    if(value >= 5) then: SetReel nplayer, 6 , 5 : SetScore nplayer,5,5 : value = value - 5: end if
    if(value >= 4) then: SetReel nplayer, 6 , 4 : SetScore nplayer,5,4 : value = value - 4: end if
    if(value >= 3) then: SetReel nplayer, 6 , 3 : SetScore nplayer,5,3 : value = value - 3: end if
    if(value >= 2) then: SetReel nplayer, 6 , 2 : SetScore nplayer,5,2 : value = value - 2: end if
    if(value >= 1) then: SetReel nplayer, 6 , 1 : SetScore nplayer,5,1 : value = value - 1: end if

    end if
    Else
      If curplayr = playr Then
      nplayer = curplayr
      value = curscr
      else
      value =scores(playr, 1) ' store score
      nplayer = playr
      end if

    scores(playr, 0)  = 0 ' reset score

    if(value >= 1000000) then

    value = value - 1000000

    end if

  ' do hundred thousands
    if(value >= 900000)  then:  SetReel nplayer, 1 , 9 : SetScore nplayer,0,9 : value = value - 900000: end if
    if(value >= 800000)  then:  SetReel nplayer, 1 , 8 : SetScore nplayer,0,8 : value = value - 800000: end if
    if(value >= 700000)  then:  SetReel nplayer, 1 , 7 : SetScore nplayer,0,7 : value = value - 700000: end if
    if(value >= 600000)  then:  SetReel nplayer, 1 , 6 : SetScore nplayer,0,6 : value = value - 600000: end if
    if(value >= 500000)  then:  SetReel nplayer, 1 , 5 : SetScore nplayer,0,5 : value = value - 500000: end if
    if(value >= 400000)  then:  SetReel nplayer, 1 , 4 : SetScore nplayer,0,4 : value = value - 400000: end if
    if(value >= 300000)  then:  SetReel nplayer, 1 , 3 : SetScore nplayer,0,3 : value = value - 300000: end if
    if(value >= 200000)  then:  SetReel nplayer, 1 , 2 : SetScore nplayer,0,2 : value = value - 200000: end if
    if(value >= 100000)  then:  SetReel nplayer, 1 , 1 : SetScore nplayer,0,1 : value = value - 100000: end if


  ' do ten thousands
    if(value >= 90000)  then:  SetReel nplayer, 2 , 9 : SetScore nplayer,1,9 : value = value - 90000: end if
    if(value >= 80000)  then:  SetReel nplayer, 2 , 8 : SetScore nplayer,1,8 : value = value - 80000: end if
    if(value >= 70000)  then:  SetReel nplayer, 2 , 7 : SetScore nplayer,1,7 : value = value - 70000: end if
    if(value >= 60000)  then:  SetReel nplayer, 2 , 6 : SetScore nplayer,1,6 : value = value - 60000: end if
    if(value >= 50000)  then:  SetReel nplayer, 2 , 5 : SetScore nplayer,1,5 : value = value - 50000: end if
    if(value >= 40000)  then:  SetReel nplayer, 2 , 4 : SetScore nplayer,1,4 : value = value - 40000: end if
    if(value >= 30000)  then:  SetReel nplayer, 2 , 3 : SetScore nplayer,1,3 : value = value - 30000: end if
    if(value >= 20000)  then:  SetReel nplayer, 2 , 2 : SetScore nplayer,1,2 : value = value - 20000: end if
    if(value >= 10000)  then:  SetReel nplayer, 2 , 1 : SetScore nplayer,1,1 : value = value - 10000: end if


  ' do thousands
    if(value >= 9000)  then:  SetReel nplayer, 3 , 9 : SetScore nplayer,2,9 : value = value - 9000: end if
    if(value >= 8000)  then:  SetReel nplayer, 3 , 8 : SetScore nplayer,2,8 : value = value - 8000: end if
    if(value >= 7000)  then:  SetReel nplayer, 3 , 7 : SetScore nplayer,2,7 : value = value - 7000: end if
    if(value >= 6000)  then:  SetReel nplayer, 3 , 6 : SetScore nplayer,2,6 : value = value - 6000: end if
    if(value >= 5000)  then:  SetReel nplayer, 3 , 5 : SetScore nplayer,2,5 : value = value - 5000: end if
    if(value >= 4000)  then:  SetReel nplayer, 3 , 4 : SetScore nplayer,2,4 : value = value - 4000: end if
    if(value >= 3000)  then:  SetReel nplayer, 3 , 3 : SetScore nplayer,2,3 : value = value - 3000: end if
    if(value >= 2000)  then:  SetReel nplayer, 3 , 2 : SetScore nplayer,2,2 : value = value - 2000: end if
    if(value >= 1000)  then:  SetReel nplayer, 3 , 1 : SetScore nplayer,2,1 : value = value - 1000: end if

    'do hundreds

    if(value >= 900) then: SetReel nplayer, 4 , 9 : SetScore nplayer,3,9 : value = value - 900: end if
    if(value >= 800) then: SetReel nplayer, 4 , 8 : SetScore nplayer,3,8 : value = value - 800: end if
    if(value >= 700) then: SetReel nplayer, 4 , 7 : SetScore nplayer,3,7 : value = value - 700: end if
    if(value >= 600) then: SetReel nplayer, 4 , 6 : SetScore nplayer,3,6 : value = value - 600: end if
    if(value >= 500) then: SetReel nplayer, 4 , 5 : SetScore nplayer,3,5 : value = value - 500: end if
    if(value >= 400) then: SetReel nplayer, 4 , 4 : SetScore nplayer,3,4 : value = value - 400: end if
    if(value >= 300) then: SetReel nplayer, 4 , 3 : SetScore nplayer,3,3 : value = value - 300: end if
    if(value >= 200) then: SetReel nplayer, 4 , 2 : SetScore nplayer,3,2 : value = value - 200: end if
    if(value >= 100) then: SetReel nplayer, 4 , 1 : SetScore nplayer,3,1 : value = value - 100: end if

    'do tens
    if(value >= 90) then: SetReel nplayer, 5 , 9 : SetScore nplayer,4,9 : value = value - 90: end if
    if(value >= 80) then: SetReel nplayer, 5 , 8 : SetScore nplayer,4,8 : value = value - 80: end if
    if(value >= 70) then: SetReel nplayer, 5 , 7 : SetScore nplayer,4,7 : value = value - 70: end if
    if(value >= 60) then: SetReel nplayer, 5 , 6 : SetScore nplayer,4,6 : value = value - 60: end if
    if(value >= 50) then: SetReel nplayer, 5 , 5 : SetScore nplayer,4,5 : value = value - 50: end if
    if(value >= 40) then: SetReel nplayer, 5 , 4 : SetScore nplayer,4,4 : value = value - 40: end if
    if(value >= 30) then: SetReel nplayer, 5 , 3 : SetScore nplayer,4,3 : value = value - 30: end if
    if(value >= 20) then: SetReel nplayer, 5 , 2 : SetScore nplayer,4,2 : value = value - 20: end if
    if(value >= 10) then: SetReel nplayer, 5 , 1 : SetScore nplayer,4,1 : value = value - 10: end if

    'do ones
    if(value >= 9) then: SetReel nplayer, 6 , 9 : SetScore nplayer,5,9 : value = value - 9: end if
    if(value >= 8) then: SetReel nplayer, 6 , 8 : SetScore nplayer,5,8 : value = value - 8: end if
    if(value >= 7) then: SetReel nplayer, 6 , 7 : SetScore nplayer,5,7 : value = value - 7: end if
    if(value >= 6) then: SetReel nplayer, 6 , 6 : SetScore nplayer,5,6 : value = value - 6: end if
    if(value >= 5) then: SetReel nplayer, 6 , 5 : SetScore nplayer,5,5 : value = value - 5: end if
    if(value >= 4) then: SetReel nplayer, 6 , 4 : SetScore nplayer,5,4 : value = value - 4: end if
    if(value >= 3) then: SetReel nplayer, 6 , 3 : SetScore nplayer,5,3 : value = value - 3: end if
    if(value >= 2) then: SetReel nplayer, 6 , 2 : SetScore nplayer,5,2 : value = value - 2: end if
    if(value >= 1) then: SetReel nplayer, 6 , 1 : SetScore nplayer,5,1 : value = value - 1: end if


    end if
  Next
End Sub

Dim BGObj

Sub SetBackglass()

  For Each BGObj In VRBackglass
    BGObj.x = BGobj.x
    BGObj.height = - BGObj.y + 120
    BGObj.y =170 'adjusts the distance from the backglass towards the user
  Next

End Sub

Sub FlasherBalls
  If Balls = 1 Then FlBIP1.visible = 1 : FlBIP1A.visible = 1 Else FlBIP1.visible = 0 : FlBIP1A.visible = 0 End If
  If Balls = 2 Then FlBIP2.visible = 1 : FlBIP2A.visible = 1 Else FlBIP2.visible = 0 : FlBIP2A.visible = 0 End If
  If Balls = 3 Then FlBIP3.visible = 1 : FlBIP3A.visible = 1 Else FlBIP3.visible = 0 : FlBIP3A.visible = 0 End If
  If Balls = 4 Then FlBIP4.visible = 1 : FlBIP4A.visible = 1 Else FlBIP4.visible = 0 : FlBIP4A.visible = 0 End If
  If Balls = 5 Then FlBIP5.visible = 1 : FlBIP5A.visible = 1 Else FlBIP5.visible = 0 : FlBIP5A.visible = 0 End If
End Sub

Sub FlasherMatch
  If Match = 100 Then FlM00.visible = 1 : FlM00A.visible = 1 : FlM00B.visible = 1 Else FlM00.visible = 0 : FlM00A.visible = 0 : FlM00B.visible = 0  End If
  If Match = 10 Then FlM10.visible = 1 : FlM10A.visible = 1 : FlM10B.visible = 1 Else FlM10.visible = 0 : FlM10A.visible = 0 : FlM10B.visible = 0 End If
  If Match = 20 Then FlM20.visible = 1 : FlM20A.visible = 1 : FlM20B.visible = 1 Else FlM20.visible = 0 : FlM20A.visible = 0 : FlM20B.visible = 0 End If
  If Match = 30 Then FlM30.visible = 1 : FlM30A.visible = 1 : FlM30B.visible = 1 Else FlM30.visible = 0 : FlM30A.visible = 0 : FlM30B.visible = 0 End If
  If Match = 40 Then FlM40.visible = 1 : FlM40A.visible = 1 : FlM40B.visible = 1 Else FlM40.visible = 0 : FlM40A.visible = 0 : FlM40B.visible = 0 End If
  If Match = 50 Then FlM50.visible = 1 : FlM50A.visible = 1 : FlM50B.visible = 1 Else FlM50.visible = 0 : FlM50A.visible = 0 : FlM50B.visible = 0 End If
  If Match = 60 Then FlM60.visible = 1 : FlM60A.visible = 1 : FlM60B.visible = 1 Else FlM60.visible = 0 : FlM60A.visible = 0 : FlM60B.visible = 0 End If
  If Match = 70 Then FlM70.visible = 1 : FlM70A.visible = 1 : FlM70B.visible = 1 Else FlM70.visible = 0 : FlM70A.visible = 0 : FlM70B.visible = 0 End If
  If Match = 80 Then FlM80.visible = 1 : FlM80A.visible = 1 : FlM80B.visible = 1 Else FlM80.visible = 0 : FlM80A.visible = 0 : FlM80B.visible = 0 End If
  If Match = 90 Then FlM90.visible = 1 : FlM90A.visible = 1 : FlM90B.visible = 1 Else FlM90.visible = 0 : FlM90A.visible = 0 : FlM90B.visible = 0 End If
End Sub

Dim Flasherseq1, Flasherseq2

Sub VRBGGameOverFlash_timer
  Select Case Flasherseq1
    case 1: VRBGGameOver.visible = 1:VRBGGOBulb1.visible = 1:VRBGGOBulb2.visible = 1
    case 5: VRBGGameOver.visible = 0:VRBGGOBulb1.visible = 0:VRBGGOBulb2.visible = 0
  End Select
  Flasherseq1 = Flasherseq1 + 1
  If Flasherseq1 > 6 Then Flasherseq1 = 1
End Sub

Sub VRBGSnoopyFlash_Timer
  dim vrbgobj
  Select Case Flasherseq2
    case 1
      For each vrbgobj in VRSnoopy1: vrbgobj.visible = 1 : Next
      For each vrbgobj in VRSnoopy2: vrbgobj.visible = 0 : Next
      For each vrbgobj in VRSnoopy3: vrbgobj.visible = 0 : Next
    case 2
      For each vrbgobj in VRSnoopy1: vrbgobj.visible = 0 : Next
      For each vrbgobj in VRSnoopy2: vrbgobj.visible = 1 : Next
      For each vrbgobj in VRSnoopy3: vrbgobj.visible = 0 : Next
    case 3
      For each vrbgobj in VRSnoopy1: vrbgobj.visible = 0 : Next
      For each vrbgobj in VRSnoopy2: vrbgobj.visible = 0 : Next
      For each vrbgobj in VRSnoopy3: vrbgobj.visible = 1 : Next
  End Select
  Flasherseq2 = Flasherseq2 + 1
  If Flasherseq2 > 3 Then Flasherseq2 = 1
End Sub






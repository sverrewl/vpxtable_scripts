' ****************************************************************
'               VISUAL PINBALL X EM Script por JPSalas
'         Script BÃ¡sico para juegos EM Script hasta 4 players
'         usa el core.vbs para funciones extras
'                        Version 1.3.0
' ****************************************************************

' Thalamus - short info
' V1.1 :
' I had made notes on how the lights work, but, I obviosly forgot to implement them all.
' Nestorgian provided DOF improvements - thanks mate.

' Dragoon by Recreativos Franco (Spanish, domestic version) - believed to be a 5 ball game.
' Dragon by Interflip - belived to be a 3 ball game by default, but, possible to change.

Option Explicit
Randomize

' Table element hit volumes - for other samples, use the SoundManager.
Const MetalsVol     = 0.6
Const RubbersVol    = 0.5
Const RubberPostVol = 0.3
Const RubberPinVol  = 0.4
Const PlasticsVol   = 1
Const GatesVol      = 1
Const WoodsVol      = 0.5
Const DragonsVol    = 0.9

' Thalamus
' In a PAPA video of this games (Dragon), it seems they have disabled special lit when all 5 dragons has been shot.
' Personally, I think that is a good idea, so I've picked that as default.

Const DisableDragonSpecial = 1
Const EnableBIPOnApron     = 0

' Valores Constantes de las fÃ­sicas de los flippers - se usan en la creaciÃ³n de las bolas, tienen que cargarse antes del core.vbs
Const BallSize = 25 ' el radio de la bola. el tamaÃ±o normal es 50 unidades de VP
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

Dim DesktopMode: DesktopMode = Table1.ShowDT

If DesktopMode = True Then
  rrail1.visible=1
  lrail1.visible=1
Else
  rrail1.visible=0
  lrail1.visible=0
End if


' Valores Constants
Const TableName = "DragoonRF" ' se usa para cargar y grabar los highscore y creditos
Const cGameName = "DragoonRF" ' para el B2S
Const MaxPlayers = 1           ' de 1 a 4
Const MaxMultiplier = 5        ' limita el bonus multiplicador a 5
Const BallsPerGame = 5         ' normalmente 3 Ã³ 5
Const Special1 = 820000         ' puntuaciÃ³n a obtener para partida extra
Const Special2 = 990000         ' puntuaciÃ³n a obtener para partida extra
Const Special3 = 100000        ' puntuaciÃ³n a obtener para partida extra

' Variables Globales
Dim PlayersPlayingGame
Dim CurrentPlayer
Dim Credits
Dim Bonus
Dim DoubleBonus
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
Dim RandPitch

' Variables de control
Dim LastSwitchHit
Dim BallsOnPlayfield

' Variables de tipo Boolean (verdadero Ã³ falso, True Ã³ False)
Dim bAttractMode
Dim bFreePlay
Dim bGameInPlay
Dim bOnTheFirstBall
Dim bExtraBallWonThisBall
Dim bJustStarted
Dim bBallInPlungerLane
Dim bBallSaverActive

' core.vbs variables, como imanes, impulse plunger

' *********************************************************************
'                Rutinas comunes para todas las mesas
' *********************************************************************

Sub Table1_Init()
  Dim x

  ' Inicializar diversos objetos de la mesa, como droptargets, animations...
  VPObjects_Init
  LoadEM

  '    ' Carga los valores grabados highscore y crÃ©ditos
  Loadhs
  DisplayHighscore
  '
  '    UpdateCredits
  '
  '    ' Juego libre o con monedas: si es True entonces no se usarÃ¡n monedas
  '    bFreePlay = False 'queremos monedas
  '
  '    ' Inicialiar las variables globales de la mesa
  '    bAttractMode = False
  '    bOnTheFirstBall = False
  '    bGameInPlay = False
  '    bBallInPlungerLane = False
  '    LastSwitchHit = ""
  '    BallsOnPlayfield = 0
  Tilt = 0
  TiltSensitivity = 6
  Tilted = False
  '    bJustStarted = True
  '    Add10 = 0
  '    Add100 = 0
  '    Add1000 = 0

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

  ' Thalamus inits

  D01.visible=false
  D02.visible=false
  D03.visible=false
  D04.visible=false
  D05.visible=false
  D01.IsDropped=false
  D02.IsDropped=false
  D03.IsDropped=false
  D04.IsDropped=false
  D05.IsDropped=false
  ResetDropTargets()
  ResetBumperLights()
  ResetTriggers()
  ResetScoringLights()
  Kicker001.TimerEnabled = 0
  If EnableBIPOnApron = 0 Then
    BallDisplay.Visible = 0
  End If
  If Credits > 0 Then
    DOF 125, DOFOn
  End If
End Sub

'******
' Keys
'******

Sub Table1_KeyDown(ByVal Keycode)
  ' aÃ±ade monedas

  if EnteringInitials then
    CollectInitials(keycode)
    exit sub
  end if

  If Keycode = AddCreditKey Then
    If(Tilted = False)Then
      AddCredits 1
      PlaySound "fx_coin"
      DOF 125, DOFOn
    End If
  End If

  ' el plunger
  If keycode = PlungerKey Then
    Plunger.Pullback
    PlaySoundAt "fx_plungerpull", plunger
  End If

  ' Funcionamiento normal de los flipers y otras teclas durante el juego

  If bGameInPlay AND NOT Tilted Then
    ' teclas de la falta
    If keycode = LeftTiltKey Then Nudge 90, 8:PlaySound "fx_nudge", 0, 1, -0.1, 0.25:CheckTilt
    If keycode = RightTiltKey Then Nudge 270, 8:PlaySound "fx_nudge", 0, 1, 0.1, 0.25:CheckTilt
    If keycode = CenterTiltKey Then Nudge 0, 9:PlaySound "fx_nudge", 0, 1, 1, 0.25:CheckTilt
    If keycode = MechanicalTilt Then CheckTilt : End If

      ' teclas de los flipers
    If keycode = LeftFlipperKey Then SolLFlipper 1
    If keycode = RightFlipperKey Then SolRFlipper 1

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
            If Credits < 1 Then
              DOF 125, DOFOff
            End If
            UpdateCredits
            UpdateBallInPlay
          Else
            ' no hay suficientes crÃ©ditos para empezar el juego.
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
        End If
      Else
        If(Credits > 0)Then
          If(BallsOnPlayfield = 0)Then
            Credits = Credits - 1
            If Credits < 1 Then
              DOF 125, DOFOff
            End If
            UpdateCredits
            ResetScores
            ResetForNewGame()
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
  ' if EnteringInitials then
  '   exit sub
  ' end if
  If bGameInPlay AND NOT Tilted Then
    ' teclas de los flipers
    If keycode = LeftFlipperKey Then SolLFlipper 0
    If keycode = RightFlipperKey Then SolRFlipper 0
  End If

  If keycode = PlungerKey Then
    Plunger.Fire
    If bBallInPlungerLane Then
      PlaySoundAt "fx_plunger", plunger
    Else
      PlaySoundAt "fx_plunger_empty", plunger
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
  Controller.Stop
End Sub

'********************
'     Flippers
'********************

Sub SolLFlipper(Enabled)
  If Enabled Then
    PlaySoundAt SoundFXDOF("fx_flipperup", 101, DOFOn, DOFFlippers), LeftFlipper
    LeftFlipper.RotateToEnd
  Else
    PlaySoundAt SoundFXDOF("fx_flipperdown", 101, DOFOff, DOFFlippers), LeftFlipper
    LeftFlipper.RotateToStart
  End If
End Sub

Sub SolRFlipper(Enabled)
  If Enabled Then
    PlaySoundAt SoundFXDOF("fx_flipperup", 102, DOFOn, DOFFlippers), RightFlipper
    RightFlipper.RotateToEnd
  Else
    PlaySoundAt SoundFXDOF("fx_flipperdown", 102, DOFOff, DOFFlippers), RightFlipper
    RightFlipper.RotateToStart
  End If
End Sub

' el sonido de la bola golpeando los flipers

Sub LeftFlipper_Collide(parm)
  PlaySoundAtBall "fx_rubber_flipper"
End Sub

Sub RightFlipper_Collide(parm)
  PlaySoundAtBall "fx_rubber_flipper"
End Sub

'*******************
' Luces GI
'*******************

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

'**************
' TILT - Falta
'**************

'el "timer" TiltDecreaseTimer resta .01 de la variable "Tilt" cada ronda

Sub CheckTilt                     'esta rutina se llama cada vez que das un golpe a la mesa
  Tilt = Tilt + TiltSensitivity 'aÃ±ade un valor al contador "Tilt"
  TiltDecreaseTimer.Enabled = True
  If Tilt > 15 Then             'Si la variable "Tilt" es mÃ¡s de 15 entonces haz falta
    Tilted = True
    TiltReel.SetValue 1       'muestra Tilt en la pantalla
    If B2SOn then
      Controller.B2SSetTilt 1
    end if
    DisableTable True
    'Esta mesa penaliza la partida, asÃ­ que quÃ­tale las bolas al jugador
    BallsRemaining(CurrentPlayer) = 0
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
    Bumper1.Force = 0
    Bumper3.Force = 0
    Bumper5.Force = 0
    '        LeftSlingshot.Disabled = 1
    '        RightSlingshot.Disabled = 1
  Else
    'enciende de nuevo todas las luces GI, bumpers y slingshots
    GiOn
    Bumper1.Force = 7
    Bumper3.Force = 7
    Bumper5.Force = 7
    '        LeftSlingshot.Disabled = 0
    '        RightSlingshot.Disabled = 0
  End If
End Sub

Sub TiltRecoveryTimer_Timer()
  ' si todas las bolas se han colado, entonces ..
  If(BallsOnPlayfield = 0)Then
    '... haz el fin de bola normal
    EndOfBall()
    TiltRecoveryTimer.Enabled = False
  End If
  ' de lo contrario esta rutina continÃºa hasta que todas las bolas se han colado
End Sub

' *******************************************************************************************************
' Positional Sound Playback Functions by DJRobX, Rothbauerw, Thalamus and Herweh
' PlaySound sound, 0, Vol(ActiveBall), AudioPan(ActiveBall), 0, Pitch(ActiveBall), 0, 1, AudioFade(ActiveBall)
' *******************************************************************************************************

' Play a sound, depending on the X,Y position of the table element (especially cool for surround speaker setups, otherwise stereo panning only)
' parameters (defaults): loopcount (1), volume (1), randompitch (0), pitch (0), useexisting (0), restart (1))
' Note that this will not work (currently) for walls/slingshots as these do not feature a simple, single X,Y position

Sub PlayXYSound(soundname, tableobj, loopcount, volume, randompitch, pitch, useexisting, restart)
  PlaySound soundname, loopcount, volume, AudioPan(tableobj), randompitch, pitch, useexisting, restart, AudioFade(tableobj)
End Sub

' Set position as table object (Use object or light but NOT wall) and Vol to 1

Sub PlaySoundAt(soundname, tableobj)
  PlaySound soundname, 1, 1, AudioPan(tableobj), 0,0,0, 1, AudioFade(tableobj)
End Sub

' set position as table object and Vol + RndPitch manually

Sub PlaySoundAtVolPitch(sound, tableobj, Vol, RndPitch)
  PlaySound sound, 1, Vol, AudioPan(tableobj), RndPitch, 0, 0, 1, AudioFade(tableobj)
End Sub

'Set all as per ball position & speed.

Sub PlaySoundAtBall(soundname)
  PlaySoundAt soundname, ActiveBall
End Sub

'Set position as table object and Vol manually.

Sub PlaySoundAtVol(sound, tableobj, Volume)
  PlaySound sound, 1, Volume, AudioPan(tableobj), 0,0,0, 1, AudioFade(tableobj)
End Sub

'Set all as per ball position & speed, but Vol Multiplier may be used eg; PlaySoundAtBallVol "sound",3

Sub PlaySoundAtBallVol(sound, VolMult)
  PlaySound sound, 0, Vol(ActiveBall) * VolMult, AudioPan(ActiveBall), 0, Pitch(ActiveBall), 0, 1, AudioFade(ActiveBall)
End Sub

Sub PlaySoundAtBallAbsVol(sound, VolMult)
  PlaySound sound, 0, VolMult, AudioPan(ActiveBall), 0, Pitch(ActiveBall), 0, 1, AudioFade(ActiveBall)
End Sub

' requires rampbump1 to 7 in Sound Manager

Sub RandomBump(voladj, freq)
  Dim BumpSnd:BumpSnd= "rampbump" & CStr(Int(Rnd*7)+1)
  PlaySound BumpSnd, 0, Vol(ActiveBall)*voladj, AudioPan(ActiveBall), 0, freq, 0, 1, AudioFade(ActiveBall)
End Sub

' set position as bumperX and Vol manually. Allows rapid repetition/overlaying sound

Sub PlaySoundAtBumperVol(sound, tableobj, Vol)
  PlaySound sound, 1, Vol, AudioPan(tableobj), 0,0,1, 1, AudioFade(tableobj)
End Sub

Sub PlaySoundAtBOTBallZ(sound, BOT)
  PlaySound sound, 0, ABS(BOT.velz)/17, Pan(BOT), 0, Pitch(BOT), 1, 0, AudioFade(BOT)
End Sub

' play a looping sound at a location with volume
Sub PlayLoopSoundAtVol(sound, tableobj, Vol)
  PlaySound sound, -1, Vol, AudioPan(tableobj), 0, 0, 1, 0, AudioFade(tableobj)
End Sub

'*********************************************************************
'                     Supporting Ball & Sound Functions
'*********************************************************************

Function RndNum(min, max)
  RndNum = Int(Rnd() * (max-min + 1) ) + min ' Sets a random number between min and max
End Function

Function AudioFade(tableobj) ' Fades between front and back of the table (for surround systems or 2x2 speakers, etc), depending on the Y position on the table. "table1" is the name of the table
  Dim tmp
  On Error Resume Next
  tmp = tableobj.y * 2 / table1.height-1
  If tmp > 0 Then
    AudioFade = Csng(tmp ^10)
  Else
    AudioFade = Csng(-((- tmp) ^10) )
  End If
End Function

Function AudioPan(tableobj) ' Calculates the pan for a tableobj based on the X position on the table. "table1" is the name of the table
  Dim tmp
  On Error Resume Next
  tmp = tableobj.x * 2 / table1.width-1
  If tmp > 0 Then
    AudioPan = Csng(tmp ^10)
  Else
    AudioPan = Csng(-((- tmp) ^10) )
  End If
End Function

Function Pan(ball) ' Calculates the pan for a ball based on the X position on the table. "table1" is the name of the table
  Dim tmp
  On Error Resume Next
  tmp = ball.x * 2 / table1.width-1
  If tmp > 0 Then
    Pan = Csng(tmp ^10)
  Else
    Pan = Csng(-((- tmp) ^10) )
  End If
End Function

Function Vol(ball) ' Calculates the Volume of the sound based on the ball speed
  Vol = Csng(BallVel(ball) ^2 / 2000)
End Function

Function VolMulti(ball,Multiplier) ' Calculates the Volume of the sound based on the ball speed
  VolMulti = Csng(BallVel(ball) ^2 / 150 ) * Multiplier
End Function

Function DVolMulti(ball,Multiplier) ' Calculates the Volume of the sound based on the ball speed
  DVolMulti = Csng(BallVel(ball) ^2 / 150 ) * Multiplier
  debug.print DVolMulti
End Function

Function BallRollVol(ball) ' Calculates the Volume of the sound based on the ball speed
  BallRollVol = Csng(BallVel(ball) ^2 / (80000 - (79900 * Log(RollVol) / Log(100))))
End Function

Function Pitch(ball) ' Calculates the pitch of the sound based on the ball speed
  Pitch = BallVel(ball) * 20
End Function

Function BallVel(ball) 'Calculates the ball speed
  BallVel = INT(SQR((ball.VelX ^2) + (ball.VelY ^2) ) )
End Function

Function BallVelZ(ball) 'Calculates the ball speed in the -Z
  BallVelZ = INT((ball.VelZ) * -1 )
End Function

Function VolZ(ball) ' Calculates the Volume of the sound based on the ball speed in the Z
  VolZ = Csng(BallVelZ(ball) ^2 / 200)*1.2
End Function

'*** Determines if a Points (px,py) is inside a 4 point polygon A-D in Clockwise/CCW order

Function InRect(px,py,ax,ay,bx,by,cx,cy,dx,dy)
  Dim AB, BC, CD, DA
  AB = (bx*py) - (by*px) - (ax*py) + (ay*px) + (ax*by) - (ay*bx)
  BC = (cx*py) - (cy*px) - (bx*py) + (by*px) + (bx*cy) - (by*cx)
  CD = (dx*py) - (dy*px) - (cx*py) + (cy*px) + (cx*dy) - (cy*dx)
  DA = (ax*py) - (ay*px) - (dx*py) + (dy*px) + (dx*ay) - (dy*ax)

  If (AB <= 0 AND BC <=0 AND CD <= 0 AND DA <= 0) Or (AB >= 0 AND BC >=0 AND CD >= 0 AND DA >= 0) Then
    InRect = True
  Else
    InRect = False
  End If
End Function

'********************************************
'   JP's VP10 Rolling Sounds + Ballshadow
' uses a collection of shadows, aBallShadow
'********************************************

Const tnob = 20 ' total number of balls
Const lob = 0   'number of locked balls
ReDim rolling(tnob)
InitRolling

Sub InitRolling
    Dim i
    For i = 0 to tnob
        rolling(i) = False
    Next
End Sub

Sub RollingUpdate()
  Dim BOT, b, ballpitch, ballvol
  BOT = GetBalls

  ' stop the sound of deleted balls
  For b = UBound(BOT) + 1 to tnob
    rolling(b) = False
    StopSound("fx_ballrolling" & b)
  Next

  ' exit the sub if no balls on the table
  If UBound(BOT) = lob - 1 Then Exit Sub 'there no extra balls on this table

  ' play the rolling sound for each ball and draw the shadow
  For b = lob to UBound(BOT)
    aBallShadow(b).X = BOT(b).X
    aBallShadow(b).Y = BOT(b).Y

    If BallVel(BOT(b)) > 1 Then
      If BOT(b).z < 30 Then
        ballpitch = Pitch(BOT(b)) - 9000 ' Thalamus, this is plastic playfield
        ballvol = Vol(BOT(b))
      Else
        ballpitch = Pitch(BOT(b)) + 25000 'increase the pitch on a ramp
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
    If BOT(b).VelZ < -1 and BOT(b).z < 55 and BOT(b).z > 27 Then 'height adjust for ball drop sounds
      PlaySound "fx_balldrop", 0, ABS(BOT(b).velz) / 57, Pan(BOT(b)), 0, Pitch(BOT(b))-9000, 1, 0, AudioFade(BOT(b))
    End If
  Next
End Sub

'*****************************
' Sonido de las bolas chocando
'*****************************

Sub OnBallBallCollision(ball1, ball2, velocity)
  PlaySound("fx_collide"), 0, Csng(velocity) ^2 / 2000, Pan(ball1), 0, Pitch(ball1), 0, 0, AudioFade(ball1)
End Sub

'***************************************
' Sonidos de las colecciones de objetos
' como metales, gomas, plÃ¡sticos, etc
'***************************************

Sub aMetals_Hit(idx):PlaySoundAtBallVol "fx_MetalHit",MetalsVol:End Sub
Sub aRubber_Bands_Hit(idx):PlaySoundAtBallVol "fx_rubber_band",RubbersVol:End Sub
Sub aRubber_Hit(idx):PlaySoundAtBallVol "fx_rubber_band",RubbersVol:End Sub
Sub aRubber_Posts_Hit(idx):PlaySoundAtBallVol "fx_rubber_post",RubberPostVol:End Sub
Sub aRubber_Pins_Hit(idx):PlaySoundAtBallVol "fx_rubber_pin",RubberPinVol:End Sub
Sub aPlastics_Hit(idx):PlaySoundAtBallVol "fx_PlasticHit",PlasticsVol:End Sub
Sub aGates_Hit(idx):PlaySoundAtBallVol "fx_Gate", GatesVol:End Sub
Sub aWoods_Hit(idx):PlaySoundAtBallVol "fx_Woodhit",WoodsVol:End Sub
Sub aDragons_Hit(idx):PlaySoundAtBallVol "fx_dragonhit",DragonsVol:End Sub

Sub ApronWall_Hit(idx)
  RandomSoundApron
End Sub

Sub RandomSoundApron()
  Select Case Int(Rnd*3)+1
    Case 1 : PlaySoundAtBall ("WD_Drain_On_Metal_Under_Apron_1")
    Case 2 : PlaySoundAtBall ("WD_Drain_On_Metal_Under_Apron_2")
    Case 3 : PlaySoundAtBall ("WD_Drain_On_Metal_Under_Apron_3")
  End Select
End Sub

'************************************************************************************************************************
' Solo para VPX 10.2 y posteriores.
' FlashForMs harÃ¡ parpadear una luz o un flash por unos milisegundos "TotalPeriod" cada tantos milisegundos "BlinkPeriod"
' Cuando el "TotalPeriod" haya terminado, la luz o el flasher se pondrÃ¡ en el estado especificado por el valor "FinalState"
' El valor de "FinalState" puede ser: 0=apagado, 1=encendido, 2=regreso al estado anterior
'************************************************************************************************************************

'Sub FlashForMs(MyLight, TotalPeriod, BlinkPeriod, FinalState)
'
'    If TypeName(MyLight) = "Light" Then ' la luz es del tipo "light"
'
'        If FinalState = 2 Then
'            FinalState = MyLight.State  'guarda el estado actual de la luz
'        End If
'        MyLight.BlinkInterval = BlinkPeriod
'        MyLight.Duration 2, TotalPeriod, FinalState
'    ElseIf TypeName(MyLight) = "Flasher" Then ' la luz es del tipo "flash"
'        Dim steps
'        ' Store all blink information
'        steps = Int(TotalPeriod / BlinkPeriod + .5) 'nÃºmero de encendidos y apagados que hay que ejecutar
'        If FinalState = 2 Then                      'guarda el estado actual del flash
'            FinalState = ABS(MyLight.Visible)
'        End If
'        MyLight.UserValue = steps * 10 + FinalState 'guarda el nÃºmero de parpadeos
'
'        ' empieza los parpadeos y crea la rutina que se va a ejecutar como un timer que se va a ejecutar los parpadeos
'        MyLight.TimerInterval = BlinkPeriod
'        MyLight.TimerEnabled = 0
'        MyLight.TimerEnabled = 1
'        ExecuteGlobal "Sub " & MyLight.Name & "_Timer:" & "Dim tmp, steps, fstate:tmp=me.UserValue:fstate = tmp MOD 10:steps= tmp\10 -1:Me.Visible = steps MOD 2:me.UserValue = steps *10 + fstate:If Steps = 0 then Me.Visible = fstate:Me.TimerEnabled=0:End if:End Sub"
'    End If
'End Sub
'
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
    for i = 82 to 94
      Controller.B2SSetData i,0
    next
  end if
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
  DoubleBonus = 1
  Bonus = 0
  UpdateBallInPlay

  Clear_Match

  ' inicializa otras variables
  Tilt = 0

  ' inicializa las variables del juego
  Game_Init()

  ' ahora puedes empezar una mÃºsica si quieres
  ' empieza la rutina "Firstball" despues de una pequeÃ±a pausa
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
  ' Se asegura que los marcadores estÃ¡n activados para el jugador de turno
  AddScore 0

  ' ajusta el multiplicador del bonus multiplier a 1X (si hubiese multiplicador en la mesa)

  ' enciende las luces, reinicializa las variables del juego, etc
  bExtraBallWonThisBall = 0
  ResetNewBallLights
  ResetNewBallVariables
  ResetBumperLights
  ResetDropTargets
  ResetScoringLights
End Sub

' Crea una bola nueva en la mesa

Sub CreateNewBall()
  ' crea una bola nueva basada en el tamaÃ±o y la masa de la bola especificados al principio del script
  BallRelease.CreateSizedBallWithMass BallSize, BallMass

  ' incrementa el nÃºmero de bolas en el tablero, ya que hay que contarlas
  BallsOnPlayfield = BallsOnPlayfield + 1

  ' actualiza las luces del backdrop
  UpdateBallInPlay

  ' y expulsa la bola
  PlaySoundAt "fx_Ballrel", BallRelease
  BallRelease.Kick 90, 4
  DOF 123, DOFPulse
End Sub

' El jugador ha perdido su bola, y ya no hay mÃ¡s bolas en juego
' Empieza a contar los bonos

Sub EndOfBall()
  'debug.print "EndOfBall"
  Dim AwardPoints, TotalBonus, ii
  AwardPoints = 0
  TotalBonus = 0
  ' La primera se ha perdido. Desde aquÃ­ ya no se puede aceptar mÃ¡s jugadores
  bOnTheFirstBall = False

  ' solo recoge los bonos si no hay falta
  ' el sistema del la falta se encargarÃ¡ de nuevas bolas o del fin de la partida

  If NOT Tilted Then
    If DoubleBonus = 2 Then
      BonusCountTimer.Interval = 400
    Else
      BonusCountTimer.Interval = 250
    End If
    BonusCountTimer.Enabled = 1
  Else 'Si hay falta simplemente espera un momento y va directo a la segunta parte despuÃ©s de perder la bola
    vpmtimer.addtimer 400, "EndOfBall2 '"
  End If
End Sub

Sub BonusCountTimer_Timer 'aÃ±ade los bonos y actualiza las luces
  ' debug.print "BonusCount_Timer"
  If Bonus > 0 Then
    Bonus = Bonus -1
    AddScore 10000
    UpdateBonusLights
  Else
    ' termina la cuenta de los bonos y continÃºa con el fin de bola
    BonusCountTimer.Enabled = 0
    vpmtimer.addtimer 1000, "EndOfBall2 '"
  End If
End Sub

Sub UpdateBonusLights 'enciende o apaga las luces de los bonos segÃºn la variable "Bonus"
    Select Case Bonus
        Case 0:li2.State = 0:li3.State = 0:li4.State = 0:li5.State = 0:li6.State = 0:li7.State = 0:li8.State = 0:li9.State = 0:li10.State = 0:li11.State = 0
        Case 1:li2.State = 1:li3.State = 0:li4.State = 0:li5.State = 0:li6.State = 0:li7.State = 0:li8.State = 0:li9.State = 0:li10.State = 0:li11.State = 0
        Case 2:li2.State = 0:li3.State = 1:li4.State = 0:li5.State = 0:li6.State = 0:li7.State = 0:li8.State = 0:li9.State = 0:li10.State = 0:li11.State = 0
        Case 3:li2.State = 0:li3.State = 0:li4.State = 1:li5.State = 0:li6.State = 0:li7.State = 0:li8.State = 0:li9.State = 0:li10.State = 0:li11.State = 0
        Case 4:li2.State = 0:li3.State = 0:li4.State = 0:li5.State = 1:li6.State = 0:li7.State = 0:li8.State = 0:li9.State = 0:li10.State = 0:li11.State = 0
        Case 5:li2.State = 0:li3.State = 0:li4.State = 0:li5.State = 0:li6.State = 1:li7.State = 0:li8.State = 0:li9.State = 0:li10.State = 0:li11.State = 0
        Case 6:li2.State = 0:li3.State = 0:li4.State = 0:li5.State = 0:li6.State = 0:li7.State = 1:li8.State = 0:li9.State = 0:li10.State = 0:li11.State = 0
        Case 7:li2.State = 0:li3.State = 0:li4.State = 0:li5.State = 0:li6.State = 0:li7.State = 0:li8.State = 1:li9.State = 0:li10.State = 0:li11.State = 0
        Case 8:li2.State = 0:li3.State = 0:li4.State = 0:li5.State = 0:li6.State = 0:li7.State = 0:li8.State = 0:li9.State = 1:li10.State = 0:li11.State = 0
        Case 9:li2.State = 0:li3.State = 0:li4.State = 0:li5.State = 0:li6.State = 0:li7.State = 0:li8.State = 0:li9.State = 0:li10.State = 1:li11.State = 0
        Case 10:li2.State = 0:li3.State = 0:li4.State = 0:li5.State = 0:li6.State = 0:li7.State = 0:li8.State = 0:li9.State = 0:li10.State = 0:li11.State = 1:LightShootAgain.state=1:AwardExtraBall
'        Case 11:li2.State = 1:li3.State = 0:li4.State = 0:li5.State = 0:li6.State = 0:li7.State = 0:li8.State = 0:li9.State = 0:li10.State = 0:li11.State = 1
'        Case 12:li2.State = 0:li3.State = 1:li4.State = 0:li5.State = 0:li6.State = 0:li7.State = 0:li8.State = 0:li9.State = 0:li10.State = 0:li11.State = 1
'        Case 13:li2.State = 0:li3.State = 0:li4.State = 1:li5.State = 0:li6.State = 0:li7.State = 0:li8.State = 0:li9.State = 0:li10.State = 0:li11.State = 1
'        Case 14:li2.State = 0:li3.State = 0:li4.State = 0:li5.State = 1:li6.State = 0:li7.State = 0:li8.State = 0:li9.State = 0:li10.State = 0:li11.State = 1
'        Case 15:li2.State = 0:li3.State = 0:li4.State = 0:li5.State = 0:li6.State = 1:li7.State = 0:li8.State = 0:li9.State = 0:li10.State = 0:li11.State = 1
    End Select
  'AddScore 10000
End Sub

' La cuenta de los bonos ha terminado. Mira si el jugador ha ganado bolas extras
' y si no mira si es el Ãºltimo jugador o la Ãºltima bola
'
Sub EndOfBall2()
  'debug.print "EndOfBall2"
  ' si hubiese falta, quÃ­tala, y pon la cuenta a cero de la falta para el prÃ³ximo jugador, Ã³ bola

  Tilted = False
  Tilt = 0
  TiltReel.SetValue 0
  If B2SOn then
    Controller.B2SSetTilt 0
  end if
  DisableTable False 'activa de nuevo los bumpers y los slingshots

  ' Â¿ha ganado el jugador una bola extra?
  If(ExtraBallsAwards(CurrentPlayer) > 0)Then
    'debug.print "Extra Ball"

    ' sÃ­? entonces se la das al jugador
    ExtraBallsAwards(CurrentPlayer) = ExtraBallsAwards(CurrentPlayer)- 1

    ' si no hay mÃ¡s bolas apaga la luz de jugar de nuevo
    If(ExtraBallsAwards(CurrentPlayer) = 0)Then
      If B2SOn then
        Controller.B2SSetShootAgain 0
      end if
    End If

    ' aquÃ­ se podrÃ­a poner algÃºn sonido de bola extra o alguna luz que parpadee

    ' En esta mesa hacemos la bola extra igual como si fuese la siguente bola, haciendo un reset de las variables y dianas
    ResetForNewPlayerBall()

    ' creamos una bola nueva en el pasillo de disparo
    CreateNewBall()
  Else ' no hay bolas extras

    BallsRemaining(CurrentPlayer) = BallsRemaining(CurrentPlayer)- 1

    ' Â¿Es Ã©sta la Ãºltima bola?
    If(BallsRemaining(CurrentPlayer) <= 0)Then

      ' miramos si la puntuaciÃ³n clasifica como el Highscore
      CheckHighScore()
    End If

    ' Ã©sta no es la Ãºltima bola para Ã©ste jugador
    ' y si hay mÃ¡s de un jugador continÃºa con el siguente
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

  ' Â¿hay otros jugadores?
  If(PlayersPlayingGame > 1)Then
    ' entonces pasa al siguente jugador
    NextPlayer = CurrentPlayer + 1
    ' Â¿vamos a pasar del Ãºltimo jugador al primero?
    ' (por ejemplo del jugador 4 al no. 1)
    If(NextPlayer > PlayersPlayingGame)Then
      NextPlayer = 1
    End If
  Else
    NextPlayer = CurrentPlayer
  End If

  'debug.print "Next Player = " & NextPlayer

  ' Â¿Hemos llegado al final del juego? (todas las bolas se han jugado de todos los jugadores)
  If((BallsRemaining(CurrentPlayer) <= 0)AND(BallsRemaining(NextPlayer) <= 0))Then

    ' aquÃ­ se empieza la loterÃ­a, normalmente cuando se juega con monedas
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

    ' hacemos un reset del la mesa para el siguente jugador (Ã³ bola)
    ResetForNewPlayerBall()

    ' y sacamos una bola
    CreateNewBall()
  End If
End Sub

' Esta funciÃ³n se llama al final del juego

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
  ' asegÃºrate de que los flippers estÃ¡n en modo de reposo
  SolLFlipper 0
  SolRFlipper 0

  ' pon las luces en el modo de fin de juego
  StartAttractMode
End Sub

' Esta funciÃ³n calcula el no de bolas que quedan
Function Balls
  Dim tmp
  tmp = BallsPerGame - BallsRemaining(CurrentPlayer) + 1
  If tmp > BallsPerGame Then
    Balls = BallsPerGame
  Else
    Balls = tmp
  End If
End Function

' Esta funciÃ³n calcula el Highscore y te da una partida gratis si has conseguido el Highscore
Sub CheckHighscore

  Dim playertops
  dim si
  dim sj,i
  dim stemp
  dim stempplayers
  for i=1 to 4
    sortscores(i)=0
    sortplayers(i)=0
  next
  playertops=0
  for i = 1 to PlayersPlayingGame
    sortscores(i)=Score(i)
    sortplayers(i)=i
  next

  for si = 1 to PlayersPlayingGame
    for sj = 1 to PlayersPlayingGame-1
      if sortscores(sj)>sortscores(sj+1) then
        stemp=sortscores(sj+1)
        stempplayers=sortplayers(sj+1)
        sortscores(sj+1)=sortscores(sj)
        sortplayers(sj+1)=sortplayers(sj)
        sortscores(sj)=stemp
        sortplayers(sj)=stempplayers
      end if
    next
  next
  HighScoreTimer.interval=100
  HighScoreTimer.enabled=True
  ScoreChecker=4
  CheckAllScores=1
  NewHighScore sortscores(ScoreChecker),sortplayers(ScoreChecker)


End Sub

' Muestra el highscore usando flashers
Sub DisplayHighscore
  exit sub
  dim tmp, digit
  tmp = HighScore
  '0 digit
  If tmp > 0 Then
    digit = tmp MOD 10 'it should be always 0 in this table
    hs0.ImageA = "h"&Digit
  Else
    hs0.ImageA = "h0"
  End If
  ' 1 digit
  tmp = tmp \10
  If tmp > 0 Then
    digit = tmp MOD 10
    hs1.ImageA = "h"&Digit
  Else
    hs1.ImageA = "h0"
  End If
  ' 2 digit
  tmp = tmp \10
  If tmp > 0 Then
    digit = tmp MOD 10
    hs2.ImageA = "h"&Digit
  Else
    hs2.ImageA = "h0"
  End If
  ' 3 digit
  tmp = tmp \10
  If tmp > 0 Then
    digit = tmp MOD 10
    hs3.ImageA = "h"&Digit
  Else
    hs3.ImageA = "h0"
  End If
  ' 4 digit
  tmp = tmp \10
  If tmp > 0 Then
    digit = tmp MOD 10
    hs4.ImageA = "h"&Digit
  Else
    hs4.ImageA = "h0"
  End If
  ' 5 digit
  tmp = tmp \10
  If tmp > 0 Then
    digit = tmp MOD 10
    hs5.ImageA = "h"&Digit
  Else
    hs5.ImageA = "h0"
  End If
End Sub

'******************
'  Match - Loteria
'******************

Sub Verification_Match()
  PlaySound "fx_match"
  Match = INT(RND(1) * 10) * 10 ' nÃºmero aleatorio entre 0 y 90
  Display_Match
  If(Score(CurrentPlayer)MOD 100) = Match Then
'    PlaySound "fx_knocker"
'    AddCredits 1
  End If
End Sub

Sub Clear_Match()
  Mtext0.SetValue 0
  Mtext1.SetValue 0
  Mtext2.SetValue 0
  Mtext3.SetValue 0
  Mtext4.SetValue 0
  Mtext5.SetValue 0
  Mtext6.SetValue 0
  Mtext7.SetValue 0
  Mtext8.SetValue 0
  Mtext9.SetValue 0
  If B2SOn then
    Controller.B2SSetMatch 0
  end if
End Sub

Sub Display_Match()
  Select Case Match
    Case 0:Mtext0.SetValue 1
    Case 10:Mtext1.SetValue 1
    Case 20:Mtext2.SetValue 1
    Case 30:Mtext3.SetValue 1
    Case 40:Mtext4.SetValue 1
    Case 50:Mtext5.SetValue 1
    Case 60:Mtext6.SetValue 1
    Case 70:Mtext7.SetValue 1
    Case 80:Mtext8.SetValue 1
    Case 90:Mtext9.SetValue 1
  End Select
  If B2SOn then
    If Match = 0 then
      Controller.B2SSetMatch 100
    else
      Controller.B2SSetMatch Match
    end if
  end if
End Sub

' *********************************************************************
'                      Drain / Plunger Functions
' *********************************************************************

' has perdido la bola ;-( mira cuantas bolas hay en el tablero.
' si solamente hay una entonces reduce el nÃºmero de bola y mira si es la Ãºltima para finalizar el juego
' si hay mÃ¡s de una, significa que hay multiball, entonces continua con la partida

Sub Drain_Hit()
  ' destruye la bola
  Drain.DestroyBall

  BallsOnPlayfield = BallsOnPlayfield - 1

  ' haz sonar el ruido de la bola
  PlaySoundAt "fx_drain", Drain
  DOF 122, DOFPulse

  'si hay falta el sistema de tilt se encargarÃ¡ de continuar con la siguente bola/jugador
  If Tilted Then
    Exit Sub
  End If

  ' si estÃ¡s jugando y no hay falta
  If(bGameInPLay = True)AND(Tilted = False)Then

    ' Â¿estÃ¡ el salva bolas activado?
    If(bBallSaverActive = True)Then

      ' Â¿sÃ­?, pues creamos una bola
      CreateNewBall()
    Else
      ' Â¿es Ã©sta la Ãºltima bola en juego?
      If(BallsOnPlayfield = 0)Then
    If LightShootAgain.state=1 then ExtraBallsAwards(CurrentPlayer) = ExtraBallsAwards(CurrentPlayer)+ 1
        vpmtimer.addtimer 500, "EndOfBall '" 'hacemos una pequeÃ±a pausa anter de continuar con el fin de bola
        Exit Sub
      End If
    End If
  End If
End Sub


Sub Trigger003_Hit()
  DOF 124, DOFOn
End Sub

Sub Trigger003_UnHit()
  DOF 124, DOFOff
  DOF 127, DOFPulse
End Sub

'Sub swPlungerRest_Hit()
'    bBallInPlungerLane = True
'End Sub

' La bola ha sido disparada, asÃ­ que cambiamos la variable, que en esta mesa se usa solo para que el sonido del disparador cambie segÃºn hay allÃ­ una bola o no
' En otras mesas podrÃ¡ usarse para poner en marcha un contador para salvar la bola

'Sub swPlungerRest_UnHit()
'    bBallInPlungerLane = False
'End Sub
'
' *********************************************************************
'               Funciones para la cuenta de los puntos
' *********************************************************************

' AÃ±ade puntos al jugador, hace sonar las campanas y actualiza el backdrop

Sub AddScore(Points)
  If Tilted Then Exit Sub
  Select Case Points
    Case 10, 100, 1000, 5000, 10000
      ' aÃ±ade los puntos a la variable del actual jugador
      Score(CurrentPlayer) = Score(CurrentPlayer) + points
      ' actualiza los contadores
      UpdateScore points
      ' hace sonar las campanillas de acuerdo a los puntos obtenidos
      If Points = 100 AND(Score(CurrentPlayer)MOD 1000) \ 100 = 0 Then  'nuevo reel de 1000
        PlaySound SoundFXDOF("fx_score1000",143,DOFPulse,DOFChimes)
      ElseIf Points = 10 AND(Score(CurrentPlayer)MOD 1000) \ 10 = 0 Then 'nuevo reel de 100
        PlaySound SoundFXDOF("fx_score100",142,DOFPulse,DOFChimes)
      Else
        PlaySound SoundFXDOF("fx_score",141,DOFPulse,DOFChimes) &Points
      End If
    Case 1000
      Add10 = Add10 + 5
      AddScore10Timer.Enabled = TRUE
    Case 10000
      Add100 = Add100 + 3
      AddScore100Timer.Enabled = TRUE
    Case 500
      Add100 = Add100 + 5
      AddScore100Timer.Enabled = TRUE
    Case 2000, 3000, 4000, 5000, 10000
      Add1000 = Add1000 + Points \ 1000
      AddScore1000Timer.Enabled = TRUE
  End Select

  ' ' aquÃ­ se puede hacer un chequeo si el jugador ha ganado alguna puntuaciÃ³n alta y darle un crÃ©dito Ã³ bola extra
  If Score(CurrentPlayer) >= Special1 AND Special1Awarded(CurrentPlayer) = False Then
    AwardSpecial
    Special1Awarded(CurrentPlayer) = True
  End If
  If Score(CurrentPlayer) >= Special2 AND Special2Awarded(CurrentPlayer) = False Then
    AwardSpecial
    Special2Awarded(CurrentPlayer) = True
  End If
  If Score(CurrentPlayer) >= Special3 AND Special3Awarded(CurrentPlayer) = False Then
    'AwardSpecial
    'Special3Awarded(CurrentPlayer) = True
  End If
End Sub

'******************************
'TIMER DE 10, 100 y 1000 PUNTOS
'******************************

' hace sonar las campanillas segÃºn los puntos

'Sub AddScore10Timer_Timer()
'    if Add10 > 0 then
'        AddScore 10
'        Add10 = Add10 - 1
'    Else
'        Me.Enabled = FALSE
'    End If
'End Sub
'
'Sub AddScore100Timer_Timer()
'    if Add100 > 0 then
'        AddScore 100
'        Add100 = Add100 - 1
'    Else
'        Me.Enabled = FALSE
'    End If
'End Sub
'
'Sub AddScore1000Timer_Timer()
'    if Add1000 > 0 then
'        AddScore 1000
'        Add1000 = Add1000 - 1
'    Else
'        Me.Enabled = FALSE
'    End If
'End Sub
'
'*******************
'     BONOS
'*******************

' ESTA MESA NO USA BONOS

' avanza el bono y actualiza las luces
' Los bonos estÃ¡n limitados a 1500 puntos

Sub AddBonus(bonuspoints)
    If(Tilted = False)Then
        ' aÃ±ade los bonos al jugador actual
        Bonus = Bonus + bonuspoints
        If Bonus > 10 Then
            Bonus = 10
        End If
        ' actualiza las luces
        UpdateBonusLights
    End if
End Sub
'
'***********************************************************************************
'        Score reels - puntuaciones - y actualiza otras luces del backdrop
'***********************************************************************************
'esta es al rutina que actualiza la puntuaciÃ³n del jugador

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
  If Credits > 0 Then
    'CreditLight.State = 1
  Else
    'CreditLight.State = 0
  End If
  CreditsReel.SetValue credits
  If B2SOn then
    Controller.B2SSetCredits Credits
  end if
End Sub

' Thalamus - I don't see a ball in play display for this machine.

Sub UpdateBallInPlay 'actualiza los marcadores de las bolas, el nÃºmero de jugador y el nÃºmero total de jugadores
  If EnableBIPOnApron = 1 Then
    Select Case Balls
      Case 0:BallDisplay.ImageA = "Ballnr0"
      Case 1:BallDisplay.ImageA = "Ballnr1"
      Case 2:BallDisplay.ImageA = "Ballnr2"
      Case 3:BallDisplay.ImageA = "Ballnr3"
      Case 4:BallDisplay.ImageA = "Ballnr4"
      Case 5:BallDisplay.ImageA = "Ballnr5"
    End Select
  End If
  If B2SOn then
    Controller.B2SSetBallInPlay Balls
  End if
End Sub

'*************************
' Partidas y bolas extras
'*************************

Sub AwardExtraBall()
  If NOT bExtraBallWonThisBall Then
    '        PlaySound "fx_knocker"
    ExtraBallsAwards(CurrentPlayer) = ExtraBallsAwards(CurrentPlayer) + 1
    bExtraBallWonThisBall = 1
    LightShootAgain.State = 1
    If B2SOn then
      Controller.B2SSetShootAgain 1
    End If
  END If
End Sub

Sub AwardSpecial()
  PlaySound SoundFXDOF("fx_knocker", 126, DOFPulse, DOFKnocker)
  AddCredits 1
  DOF 125, DOFOn
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

  ' enciente la luz de fin de partida
  GameOverR.SetValue 1
  BallDisplay.ImageA = "ballnr0"
End Sub

Sub StopAttractMode()
  Dim x
  bAttractMode = False
  ResetScores
  For each x in aLights
    x.State = 0
  Next

  ' apaga la luz de fin de partida
  GameOverR.SetValue 0
End Sub

'************************************************
'    Load (cargar) / Save (guardar)/ Highscore
'************************************************

' solamente guardamos el nÃºmero de crÃ©ditos y la puntuaciÃ³n mÃ¡s alta
'
Sub Loadhs
    ' Based on Black's Highscore routines
 Dim FileObj
 Dim ScoreFile,TextStr
    dim temp1
    dim temp2
 dim temp3
 dim temp4
 dim temp5
 dim temp6
 dim temp8
 dim temp9
 dim temp10
 dim temp11
 dim temp12
 dim temp13
 dim temp14
 dim temp15
 dim temp16
 dim temp17


    Set FileObj=CreateObject("Scripting.FileSystemObject")
 If Not FileObj.FolderExists(UserDirectory) then
   Credits=0
   Exit Sub
 End if
 If Not FileObj.FileExists(UserDirectory & TableName&".txt") then
   Credits=0
   Exit Sub
 End if
 Set ScoreFile=FileObj.GetFile(UserDirectory & TableName&".txt")
 Set TextStr=ScoreFile.OpenAsTextStream(1,0)
   If (TextStr.AtEndOfStream=True) then
     Exit Sub
   End if
   temp1=TextStr.ReadLine
   temp2=textstr.readline

   temp2=0 ' Thalamus, start every game with 0 credits

   HighScore=cdbl(temp1)
   if HighScore<1 then

     temp8=textstr.readline
     temp9=textstr.readline
     temp10=textstr.readline
     temp11=textstr.readline
     temp12=textstr.readline
     temp13=textstr.readline
     temp14=textstr.readline
     temp15=textstr.readline
     temp16=textstr.readline
     temp17=textstr.readline
   end if
   TextStr.Close
     Credits=cdbl(temp2)

   if HighScore<1 then
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
   end if
   Set ScoreFile=Nothing
     Set FileObj=Nothing

End Sub

Sub Savehs
 ' Based on Black's Highscore routines
 Dim FileObj
 Dim ScoreFile
 Dim xx
 Set FileObj=CreateObject("Scripting.FileSystemObject")
 If Not FileObj.FolderExists(UserDirectory) then
   Exit Sub
 End if
 Set ScoreFile=FileObj.CreateTextFile(UserDirectory & TableName&".txt",True)
   ScoreFile.WriteLine 0
   ScoreFile.WriteLine Credits
   for xx=1 to 5
     scorefile.writeline HSScore(xx)
   next
   for xx=1 to 5
     scorefile.writeline HSName(xx)
   next
   ScoreFile.Close
 Set ScoreFile=Nothing
 Set FileObj=Nothing
End Sub

' por si se necesitara quitar la actual puntuaciÃ³n mÃ¡s alta, se le puede poner a una tecla,
' o simplemente abres la ventana de debug y escribes Reseths y le das al enter

Sub Reseths
    HighScore = 0
    Savehs
End Sub

'****************************************
' Actualizaciones en tiempo real
'****************************************
' se usa sobre todo para hacer animaciones o sonidos que cambian en tiempo real
' como por ejemplo para sincronizar los flipers, puertas Ã³ molinillos con primitivas

Sub GameTimer_Timer
    RollingUpdate 'actualiza el sonido de la bola rodando
                  ' y tambiÃ©n algunas animaciones, sobre todo de primitivas

End Sub

'***********************************************************************
' *********************************************************************
'            AquÃ­ empieza el cÃ³digo particular a la mesa
' (hasta ahora todas las rutinas han sido muy generales para todas las mesas)
' (y hay muy pocas rutinas que necesitan cambiar de mesa a mesa)
' *********************************************************************
'***********************************************************************

' se inicia las dianas abatibles, primitivas, etc.
' aunque en el VPX no hay muchos objetos que necesitan ser iniciados

Sub VPObjects_Init 'en esta mesa no hay nada que necesite iniciarse, pero dejamos la rutina para prÃ³ximas mesas
End Sub

' variables de la mesa

'Dim Target6V, Target7V, Target8V, Target9V, Target10V 'valores de los targets centrales
'Dim BumperCAct                                        ' indica si los bumpers centrales estÃ¡n activos
'Dim BumperLAct                                        ' indica si los bumpers centrales estÃ¡n activos
'Dim Lights5Completed
'Dim Lights8Completed
'Dim IsSpin
'
Sub Game_Init() 'esta rutina se llama al principio de un nuevo juego
'  BumperCAct = False
'  BumperLAct = False
'  Lights5Completed = False
'  Lights8Completed = False
'  IsSpin = True
'  'Empezar alguna mÃºsica, si hubiera mÃºsica en esta mesa
'
'  'iniciar variables, en esta mesa hay muy pocas variables ya que usamos las luces, y el UserValue de las dianas
'  StartSpin 'para iniciar las variables Target6V
'  'iniciar algÃºn timer
'
'  'Iniciar algunas luces, en esta mesa son las mismas luces que las de una bola nueva
'  TurnOffPlayfieldLights()
'  ' enciende las luces de las dianas
'  li21.State = 1
'  li22.State = 1
'  li23.State = 1
'  li24.State = 1
'  li25.State = 1
'  'enciende las luces de los pasillos
'  li9.State = 1
'  li6.State = 1
'  li10.State = 1
'  li11.State = 1
'  li15.State = 1
'  li14.State = 1
'  li13.State = 1
'  li12.State = 1
End Sub
'
Sub ResetNewBallVariables() 'inicia las variable para una bola Ã³ jugador nuevo
End Sub

Sub ResetNewBallLights()    'enciende Ã³ apaga las luces para una bola nueva
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
' En cada diana u objeto con el que la bola choque habrÃ¡ que hacer:
' - sonar un sonido fÃ­sico
' - hacer algÃºn movimiento, si es necesario
' - aÃ±adir alguna puntuaciÃ³n
' - encender/apagar una luz
' - hacer algÃºn chequeo para ver si el jugador ha completado algo
' *********************************************************************

' la bola choca contra los Slingshots
' hacemos una animaciÃ³n manual de los slingshots usando gomas

'Dim LStep, RStep
'
'Sub LeftSlingShot_Slingshot
'    If Tilted Then Exit Sub
'    PlaySoundAtBall "fx_slingshot"
'    LeftSling4.Visible = 1
'    Lemk.RotX = 26
'    LStep = 0
'    LeftSlingShot.TimerEnabled = True
'    ' aÃ±ade algunos puntos
'    AddScore 10
'    ' aÃ±ade algÃºn efecto a la mesa
'    AlternateSpecials
'    RotateKJAQ10
'End Sub
'
'Sub LeftSlingShot_Timer
'    Select Case LStep
'        Case 1:LeftSLing4.Visible = 0:LeftSLing3.Visible = 1:Lemk.RotX = 14
'        Case 2:LeftSLing3.Visible = 0:LeftSLing2.Visible = 1:Lemk.RotX = 2
'        Case 3:LeftSLing2.Visible = 0:Lemk.RotX = -20:LeftSlingShot.TimerEnabled = 0
'    End Select
'    LStep = LStep + 1
'End Sub
'
'Sub RightSlingShot_Slingshot
'    If Tilted Then Exit Sub
'    PlaySoundAtBall "fx_slingshot"
'    RightSling4.Visible = 1
'    Remk.RotX = 26
'    RStep = 0
'    RightSlingShot.TimerEnabled = True
'    ' aÃ±ade algunos puntos
'    AddScore 10
'    ' aÃ±ade algÃºn efecto a la mesa
'    AlternateSpecials
'    RotateKJAQ10
'End Sub
'
'Sub RightSlingShot_Timer
'    Select Case RStep
'        Case 1:RightSLing4.Visible = 0:RightSLing3.Visible = 1:Remk.RotX = 14
'        Case 2:RightSLing3.Visible = 0:RightSLing2.Visible = 1:Remk.RotX = 2
'        Case 3:RightSLing2.Visible = 0:Remk.RotX = -20:RightSlingShot.TimerEnabled = 0
'    End Select
'    RStep = RStep + 1
'End Sub
'

' Thalamus - what I've learned from watching youtube.
' 1 lights left Bumper1, 3 lights middle and 5 lights right.
' lit bumper = 10.000 pts, unlit = 1.000
' dragons and kicker = 5.000 - add to bonus - all down - red special
' rollover 100, 1.000 if lit.
' inlane - always 5000 ? - weird, would expect there to be different scores.
' kicker 5000 always ?
' dt 5000 and a different sound ?
' Starts with upper dividers lit, turns off, and lights the inlane lights for the corresponding number.
' Hit inlane lower numer turns corresponding on top.
' All numbers lights both side specials
' All dragons lights special for the kicker.
' Loserman said EB can be adjusted between 5 and 10. Let's put it on a high 10 - be aware, if using a lower amount, extra code for EB is needed..

Dim Bumper1Score,Bumper3Score,Bumper5Score
Dim LTriggerScore1, LTriggerScore2, LTriggerScore3, LTriggerScore4, LTriggerScore5
Dim RTriggerScore1, RTriggerScore2, RTriggerScore3, RTriggerScore4, RTriggerScore5
Dim Dragons

Sub Trigger001_Hit
  LightShootAgain.state=0
  If B2SOn then
    Controller.B2SSetShootAgain 0
  End if
End Sub

Sub Trigger002_Hit
  If ExtraBallsAwards(CurrentPlayer) = 0 Then Exit Sub
  ExtraBallsAwards(CurrentPlayer) = ExtraBallsAwards(CurrentPlayer) - 1
  bExtraBallWonThisBall=0
End Sub

Sub ResetTriggers
  LTriggerScore1 = 100
  LTriggerScore2 = 100
  LTriggerScore3 = 100
  LTriggerScore4 = 100
  LTriggerScore5 = 100
  RTriggerScore1 = 100
  RTriggerScore2 = 100
  RTriggerScore3 = 100
  RTriggerScore4 = 100
  RTriggerScore5 = 100
End Sub

Sub ResetScoringLights
  Dim x
  lli1.state = 0
  lli2.state = 0
  lli3.state = 0
  rli4.state = 0
  rli5.state = 0
  rli3.state = 0
  uli1.state = 0
  uli2.state = 0
  uli3.state = 0
  uli4.state = 0
  uli5.state = 0
  sli1.state = 0
  sli2.state = 0
  sli3.state = 0
  For each x in aButtonLights
    x.State = 0
  Next
  For each x in aUButtonLights
    x.State = 1
  Next
End Sub

Sub ResetBumperLights
  LightBumper1.state = 0
  LightBumper3.state = 0
  LightBumper5.state = 0
  Bumper1Score = 1000
  Bumper3Score = 1000
  Bumper5Score = 1000
End Sub

Sub ResetDropTargets
  If D01.collidable=0 Then
    DOF 131, DOFPulse
  End If
  If D02.collidable=0 Then
    DOF 128, DOFPulse
  End If
  If D03.collidable=0 Then
    DOF 129, DOFPulse
  End If
  If D04.collidable=0 Then
    DOF 132, DOFPulse
  End If
  If D05.collidable=0 Then
    DOF 130, DOFPulse
  End If
  If ( ( D01.collidable = 0 ) or ( D02.collidable = 0 ) or ( D03.collidable = 0 ) or ( D04.collidable = 0) or ( D05.collidable = 0 ) ) Then
    RandPitch = RndNum(4,9)*50
    PlaySoundAtVolPitch "reset_dragons", D03, 1, RandPitch
  End If
  D01.IsDropped=0
  D02.IsDropped=0
  D03.IsDropped=0
  D04.IsDropped=0
  D05.IsDropped=0
  PrimD01.RotX=0:PrimD01.collidable=1
  PrimD02.RotX=0:PrimD02.collidable=1
  PrimD03.RotX=0:PrimD03.collidable=1
  PrimD04.RotX=0:PrimD04.collidable=1
  PrimD05.RotX=0:PrimD05.collidable=1
  D01.collidable=1
  D02.collidable=1
  D03.collidable=1
  D04.collidable=1
  D05.collidable=1
  D01.visible=0
  D02.visible=0
  D03.visible=0
  D04.visible=0
  D05.visible=0
  Dragons = 0
End Sub

Sub TButtons_Hit(Index)
  Stopsound"buttons" : PlaysoundAt "buttons", ActiveBall
End Sub

Sub Kicker001_hit 'right kicker
    PlaySoundAt "fx_kicker_enter", Kicker001
    If Tilted Then Ejectkicker:Exit Sub
  AddScore 5000
  Kicker001.TimerInterval=1500
  Kicker001.TimerEnabled = 1
  If sli2.state = 1 Then
' AddCredits 1
  AwardSpecial
  End If
End Sub

Sub Kicker001_Timer
  PlaySoundAt "fx_kicker", Kicker001
  DOF 121, DOFPulse
  DOF 133, DOFPulse
  Kicker001.kick 168+RndNum(1,6), 10
  Kicker001.TimerEnabled = 0
End Sub

Dim Steps,ii

Sub CheckSpecial ' All Dragons down - lights Special for kicker
  If lli1.state = 1 AND lli2.state = 1 AND lli3.state = 1 AND rli4.state = 1 AND rli5.state = 1 Then
    sli1.state = 1
  sli3.state = 1
  End If
  if Dragons >= 5 and DisableDragonSpecial = 0 Then
  sli2.state = 1
  End If
End Sub

Sub D01_Hit() ' L11 - L21
  If Tilted = 1 Then Exit Sub
  Steps = 10
  li001.State = 1
  eli001.State = 1
  li006.State = 1
  eli006.State = 1
  For ii = 0 to Steps
    PrimD01.RotX=PrimD01.RotX+(90/Steps)
  Next
  PrimD01.collidable=0
  D01.collidable=0
  AddScore 5000
  AddBonus 1
  LTriggerScore1 = 1000
  RTriggerScore1 = 1000
  Dragons = Dragons + 1
  PlaySound "FX_Baow"
  DOF 136, DOFPulse
  CheckSpecial
End Sub

Sub D02_Hit()
  If Tilted = 1 Then Exit Sub
  li002.State = 1
  eli002.State = 1
  li007.State = 1
  eli007.State = 1
  Steps = 89
  For ii = 0 to Steps
    PrimD02.RotX=PrimD02.RotX+(90/Steps)
  Next
  PrimD02.collidable=0
  D02.collidable=0
  AddScore 5000
  AddBonus 1
  LTriggerScore2 = 1000
  RTriggerScore2 = 1000
  Dragons = Dragons + 1
  PlaySound "FX_Baow"
  DOF 136, DOFPulse
  CheckSpecial
End Sub

Sub D03_Hit()
  If Tilted = 1 Then Exit Sub
  Steps = 89
  li003.State = 1
  eli003.State = 1
  li008.State = 1
  eli008.State = 1
  For ii = 0 to Steps
    PrimD03.RotX=PrimD03.RotX+(90/Steps)
  Next
  PrimD03.collidable=0
  D03.collidable=0
  AddScore 5000
  AddBonus 1
  LTriggerScore3 = 1000
  RTriggerScore3 = 1000
  Dragons = Dragons + 1
  PlaySound "FX_Baow"
  DOF 136, DOFPulse
  CheckSpecial
End Sub

Sub D04_Hit()
  If Tilted = 1 Then Exit Sub
  Steps = 89
  li004.State = 1
  eli004.State = 1
  li009.State = 1
  eli009.State = 1
  For ii = 0 to Steps
    PrimD04.RotX=PrimD04.RotX+(90/Steps)
  Next
  PrimD04.collidable=0
  D04.collidable=0
  AddScore 5000
  AddBonus 1
  LTriggerScore4 = 1000
  RTriggerScore4 = 1000
  Dragons = Dragons + 1
  PlaySound "FX_Baow"
  DOF 136, DOFPulse
  CheckSpecial
End Sub

Sub D05_Hit()
  If Tilted = 1 Then Exit Sub
  Steps = 89
  li005.State = 1
  eli005.State = 1
  li010.State = 1
  eli010.State = 1
  For ii = 0 to Steps
    PrimD05.RotX=PrimD05.RotX+(90/Steps)
  Next
  PrimD05.collidable=0
  D05.collidable=0
  AddScore 5000
  AddBonus 1
  LTriggerScore5 = 1000
  RTriggerScore5 = 1000
  Dragons = Dragons + 1
  PlaySound "FX_Baow"
  DOF 136, DOFPulse
  CheckSpecial
End Sub

Sub Target001_Hit()
  If Tilted Then Exit Sub
  DOF 117, DOFPulse
  AddScore 5000
End Sub

Sub Target002_Hit()
  If Tilted Then Exit Sub
  DOF 118, DOFPulse
  AddScore 5000
End Sub

Sub Target003_Hit()
  If Tilted Then Exit Sub
  DOF 119, DOFPulse
  AddScore 5000
  AddBonus 1
  If sli1.state = 1 Then
'    AddCredits 1
    AwardSpecial
  End If
End Sub

Sub Target004_Hit()
  If Tilted Then Exit Sub
  DOF 120, DOFPulse
  AddScore 5000
  AddBonus 1
  If sli1.state = 1 Then
'    AddCredits 1
    AwardSpecial
  End If
End Sub

Sub RTrigger1_Hit
  If Tilted Then Exit Sub
  AddScore RTriggerScore1
End Sub

Sub RTrigger2_Hit
  If Tilted Then Exit Sub
  AddScore RTriggerScore2
End Sub

Sub RTrigger3_Hit
  If Tilted Then Exit Sub
  AddScore RTriggerScore3
End Sub

Sub RTrigger4_Hit
  If Tilted Then Exit Sub
  AddScore RTriggerScore4
End Sub

Sub RTrigger5_Hit
  If Tilted Then Exit Sub
  AddScore RTriggerScore5
End Sub

Sub DTrigger1_Hit
  If Tilted Then Exit Sub
  DOF 111, DOFPulse
  AddScore 5000
  uli1.state = 0
  li011.state = 0
  eli011.state = 0
  lli1.state = 1
  li017.state = 1
  eli017.state = 1
  Bumper1Score = 10000
  LightBumper1.state = 1
End Sub

Sub DTrigger2_Hit
  If Tilted Then Exit Sub
  AddScore 5000
  uli2.state = 0
  li012.state = 0
  eli012.state = 0
  lli2.state = 1
  li018.state = 1
  eli018.state = 1
End Sub

Sub DTrigger3_Hit
  If Tilted Then Exit Sub
  AddScore 5000
  uli3.state = 0
  li013.state = 0
  eli013.state = 0
  lli3.state = 1
  li016.state = 1
  eli016.state = 1
  rli3.state = 1
  li021.state = 1
  eli021.state = 1
  Bumper1Score = 10000
  LightBumper3.state = 1
End Sub

Sub DTrigger4_Hit
  If Tilted Then Exit Sub
  AddScore 5000
  uli4.state = 0
  li014.state = 0
  eli014.state = 0
  rli4.state = 1
  li019.state = 1
  eli019.state = 1
End Sub

Sub DTrigger5_Hit
  If Tilted Then Exit Sub
  AddScore 5000
  uli5.state = 0
  li015.state = 0
  eli015.state = 0
  rli5.state = 1
  li020.state = 1
  eli020.state = 1
  Bumper1Score = 10000
  LightBumper5.state = 1
End Sub

Sub DTrigger6_Hit
  If Tilted Then Exit Sub
  AddScore 5000
  uli3.state = 0
  li013.state = 0
  eli013.state = 0
  rli3.state = 1
  li021.state = 1
  eli021.state = 1
  lli3.state = 1
  li016.state = 1
  eli016.state = 1
  LightBumper3.state = 1
End Sub

Sub LTrigger1_Hit
  If Tilted Then Exit Sub
  AddScore LTriggerScore1
End Sub

Sub LTrigger2_Hit
  If Tilted Then Exit Sub
  AddScore LTriggerScore2
End Sub

Sub LTrigger3_Hit
  If Tilted Then Exit Sub
  AddScore LTriggerScore3
End Sub

Sub LTrigger4_Hit
  If Tilted Then Exit Sub
  AddScore LTriggerScore4
End Sub

Sub LTrigger5_Hit
  If Tilted Then Exit Sub
  AddScore LTriggerScore5
End Sub

Sub TriggerT1_Hit
  If Tilted Then Exit Sub
  DOF 106, DOFPulse
  AddScore 5000
  Bumper1Score = 10000
  LightBumper1.state = 1
  uli1.state = 0
  li011.state = 0
  eli011.state = 0
  lli1.state = 1
  li017.state = 1
  eli017.state = 1
  CheckSpecial
End Sub

Sub TriggerT2_Hit
  If Tilted Then Exit Sub
  DOF 107, DOFPulse
  AddScore 5000
  uli2.state = 0
  li012.state = 0
  eli012.state = 0
  lli2.state = 1
  li018.state = 1
  eli018.state = 1
  CheckSpecial
End Sub

Sub TriggerT3_Hit
  If Tilted Then Exit Sub
  DOF 108, DOFPulse
  AddScore 5000
  Bumper3Score = 10000
  LightBumper3.state = 1
  ' ligth up 2 x 3 inlane
  uli3.state = 0
  li013.state = 0
  eli013.state = 0
  lli3.state = 1
  rli3.state = 1
  li016.state = 1
  eli016.state = 1
  li021.state = 1
  eli021.state = 1
  CheckSpecial
End Sub

Sub TriggerT4_Hit
  If Tilted Then Exit Sub
  DOF 109, DOFPulse
  AddScore 5000
  uli4.state = 0
  li014.state = 0
  eli014.state = 0
  rli4.state = 1
  li019.state = 1
  eli019.state = 1
  CheckSpecial
End Sub

Sub TriggerT5_Hit
  If Tilted Then Exit Sub
  DOF 110, DOFPulse
  AddScore 5000
  Bumper5Score = 10000
  LightBumper5.state = 1
  uli5.state = 0
  li015.state = 0
  eli015.state = 0
  rli5.state = 1
  li020.state = 1
  eli020.state = 1
  CheckSpecial
End Sub

Sub Bumper1_Hit
  If Tilted Then Exit Sub
  ' LastSwitchHit = "bumper1"
    RandPitch = -RndNum(6,9)*50
    PlaySoundAtVolPitch SoundFXDOF("Bumper", 103, DOFPulse, DOFContactors), bumper1, 1, RandPitch
  '  PlaySoundAt "fx_Bumper2", bumper1
  ' aÃ±ade algunos puntos
  AddScore Bumper1Score
End Sub

Sub Bumper3_Hit
  If Tilted Then Exit Sub
  ' LastSwitchHit = "bumper1"
    RandPitch = -RndNum(6,9)*50
    PlaySoundAtVolPitch SoundFXDOF("Bumper", 104, DOFPulse, DOFContactors), bumper3, 1, RandPitch
'  PlaySoundAt "Bumper2", bumper3
  ' aÃ±ade algunos puntos
  AddScore Bumper3Score
End Sub

Sub Bumper5_Hit
  If Tilted Then Exit Sub
    RandPitch = -RndNum(6,9)*50
    PlaySoundAtVolPitch SoundFXDOF("Bumper", 105, DOFPulse, DOFContactors), bumper5, 1, RandPitch
  ' LastSwitchHit = "bumper1"
'  PlaySoundAt "fx_Bumper1", bumper5
  ' aÃ±ade algunos puntos
  AddScore Bumper5Score
End Sub

' ============================================================================================
' GNMOD - Multiple High Score Display and Collection
' changed ramps by flashers, jpsalas
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
HSScore(1) = 500000
HSScore(2) = 400000
HSScore(3) = 300000
HSScore(4) = 250000
HSScore(5) = 150000

HSName(1) = "THA"
HSName(2) = "JPS"
HSName(3) = "LMN"
HSName(4) = "INK"
HSName(5) = "NES"

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
    'AwardSpecial
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

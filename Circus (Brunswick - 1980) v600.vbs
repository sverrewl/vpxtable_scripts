' ****************************************************************
'               VISUAL PINBALL X EM Script por JPSalas
'         Script Básico para juegos EM Script hasta 4 players
'        usa el core.vbs para funciones extras
' ****************************************************************

Option Explicit
Randomize

Const bMusicON = False ' set this to True if you want some circus music while playing
Const bMusicVOL = 0.3   'set the Volume of the music, from 0 to 1.

'
' DOF config - leeoneil
'
' Option for more lights effects with DOF (Undercab effects on bumpers and slingshots)
' "True" to activate (False by default)
Const Epileptikdof = False
'
' Flippers L/R - 101/102
' Slingshot right - 103/104
' Bumper - 105
' Targets L/R - 106/107/108
'
' LED backboard
' Flasher Outside Left - 203/207/210/240/250
' Flasher left - 213/217/218/240/250
' Flasher center - 205/214/218/240/250
' Flasher right - 211/215/218/240/250
' Flasher Outside - 208/209/212/216/219/240/250
' MX Flashers L - 204
' Start Button - 200
' Undercab - 201/202
' Strobe - 230
' Knocker - 300

' Valores Constantes de las físicas de los flippers - se usan en la creación de las bolas, tienen que cargarse antes del core.vbs
Const BallSize = 50 ' el radio de la bola. el tamaño normal es 50 unidades de VP
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

' Valores Constants
Const TableName = "Circus-Brunswick"
Const cGameName = "CircusB"
Const MaxPlayers = 4    ' de 1 a 4
Const MaxMultiplier = 5 ' limita el bonus multiplicador a 5
Const Special1 = 560000 ' puntuación a obtener para partida extra
Const Special2 = 720000 ' puntuación a obtener para partida extra
Const Special3 = 999000 ' puntuación a obtener para partida extra

' Variables Globales
Dim PlayersPlayingGame
Dim CurrentPlayer
Dim Credits
Dim Bonus
Dim BonusMultiplier
Dim ScoreMultiplier
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
Dim ii,jj ' STAT

' Variables de control
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

Dim x

' core.vbs variables, como imanes, impulse plunger

' *********************************************************************
'                Rutinas comunes para todas las mesas
' *********************************************************************

Sub Table1_Init()

    ' Inicializar diversos objetos de la mesa, como droptargets, animations...
    VPObjects_Init
    LoadEM

    ' Juego libre o con monedas: si es True entonces no se usarán monedas
    bFreePlay = True 'no usa monedas

    ' Carga los valores grabados highscore y créditos
    Credits = 1
    Loadhs
    'HighscoreReel.SetValue Highscore
    ' también pon el highscore en el reel del primer jugador
    ' ScoreReel1.SetValue Highscore
    ' en esta mesa el attractmode se encarga de mostrar el highscore
    If  bFreePlay = False Then
        UpdateCredits
    End If

    ' Inicialiar las variables globales de la mesa
    bAttractMode = False
    bOnTheFirstBall = False
    bGameInPlay = False
    bBallInPlungerLane = False
    BallsOnPlayfield = 0
    Tilt = 0
    TiltSensitivity = 6
    Tilted = False
    bJustStarted = True
    Add10 = 0
    Add100 = 0
    Add1000 = 0

    ' pone la mesa en modo de espera
    EndOfGame

    'Enciende las luces GI despues de un segundo
    vpmtimer.addtimer 1000, "GiOn '"

    ' Quita los laterales y las puntuaciones cuando la mesa se juega en modo FS
    If Table1.ShowDT = False then
        lrail.Visible = False
        rrail.Visible = False
        For each x in aReels
            x.Visible = 0
        Next
    End If
    ' Arranca el Game Timer de las animaciones, sonido de las bolas rodando, etc
    GameTimer.Enabled = 1
End Sub

Sub NumPlay_Timer ' STAT
  if B2SOn then
    if ii = 1 and jj < 10 then
      Controller.B2ssetplayerup 30, PlayersPlayingGame: ii = 0
    else
      Controller.B2ssetplayerup 30, 0: ii = 1
    end if
  jj = jj + 1:if jj = 10 then Controller.B2ssetplayerup 30, 1:NumPlay.enabled=false
  end if
End Sub

'******
' Keys
'******

Sub Table1_KeyDown(ByVal Keycode)

    If EnteringInitials then
        CollectInitials(keycode)
        Exit Sub
    End If

    ' añade monedas
    If Keycode = AddCreditKey OR Keycode = AddCreditKey2 Then
        If(Tilted = False)Then
            AddCredits 1
            PlaySound "fx_coin"
            DOF 200, DOFOn
        End If
    End If

    ' el plunger
    If keycode = PlungerKey Then
        Plunger.Pullback
        PlaySoundAt "fx_plungerpull", plunger
    End If

    ' Funcionamiento normal de los flipers y otras teclas durante el juego

    If bGameInPlay AND NOT Tilted Then
        If Credits = 0 Then DOF 200, DOFOff
        ' teclas de la falta
        If keycode = LeftTiltKey Then Nudge 90, 8:PlaySound "fx_nudge", 0, 1, -0.1, 0.25:CheckTilt
        If keycode = RightTiltKey Then Nudge 270, 8:PlaySound "fx_nudge", 0, 1, 0.1, 0.25:CheckTilt
        If keycode = CenterTiltKey Then Nudge 0, 9:PlaySound "fx_nudge", 0, 1, 1, 0.25:CheckTilt
        If keycode = MechanicalTilt Then CheckTilt

        ' teclas de los flipers
        If keycode = LeftFlipperKey Then SolLFlipper 1
        If keycode = RightFlipperKey Then SolRFlipper 1

        ' tecla de empezar el juego
        If keycode = StartGameKey Then
            If((PlayersPlayingGame < MaxPlayers)AND(bOnTheFirstBall = True))Then

                If(bFreePlay = True)Then
                    PlayersPlayingGame = PlayersPlayingGame + 1
          UpdatePlayersCanPlay
                    PlaySound "newplayer"
                Else
                    If(Credits > 0)then
                        PlayersPlayingGame = PlayersPlayingGame + 1
                        Credits = Credits - 1
                        UpdateCredits
                        UpdateBallInPlay
            UpdatePlayersCanPlay
                        PlaySound "newplayer"
                    Else
                        ' no hay suficientes créditos para empezar el juego.
                        PlaySound "10000"
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
                        PlaySound "startup"
                    End If
                Else
                    If(Credits > 0)Then
                        If(BallsOnPlayfield = 0)Then
                            Credits = Credits - 1
                            UpdateCredits
                            ResetScores
                            ResetForNewGame()
                            PlaySound "startup"
                        End If
                    Else
                        DOF 200, DOFOff
                        ' Not Enough Credits to start a game.
                        PlaySound "10000"
                    End If
                End If
            End If
    End If ' If (GameInPlay)
End Sub

Sub Table1_KeyUp(ByVal keycode)

    If EnteringInitials then
        Exit Sub
    End If

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
'Controller.Stop
End Sub

'******************************
'     DOF lights ball entrance
'******************************
'
Sub Trigger001_Hit
    DOF 240, DOFPulse
End sub

'********************
'     Flippers
'********************

Sub SolLFlipper(Enabled)
    If Enabled Then
        PlaySoundAt SoundFXDOF ("fx_flipperup", 101, DOFOn, DOFFlippers), LeftFlipper
        LeftFlipper.RotateToEnd
        LeftFlipperOn = 1
    Else
        PlaySoundAt SoundFXDOF ("fx_flipperdown", 101, DOFOff, DOFFlippers), LeftFlipper
        LeftFlipper.RotateToStart
        LeftFlipperOn = 0
    End If
End Sub

Sub SolRFlipper(Enabled)
    If Enabled Then
        PlaySoundAt SoundFXDOF ("fx_flipperup", 102, DOFOn, DOFFlippers), RightFlipper
        RightFlipper.RotateToEnd
        RightFlipperOn = 1
    Else
        PlaySoundAt SoundFXDOF ("fx_flipperdown", 102, DOFOff, DOFFlippers), RightFlipper
        RightFlipper.RotateToStart
        RightFlipperOn = 0
    End If
End Sub

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

'*******************
' Luces GI
'*******************

Sub GiOn 'enciende las luces GI
    DOF 201, DOFOn
    Dim bulb
    For each bulb in aGiLights
        bulb.State = 1
    Next
End Sub

Sub GiOff 'apaga las luces GI
    DOF 201, DOFOff
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
    Tilt = Tilt + TiltSensitivity 'añade un valor al contador "Tilt"
    TiltDecreaseTimer.Enabled = True
    If Tilt > 15 Then             'Si la variable "Tilt" es más de 15 entonces haz falta
        Tilted = True
        TiltReel.SetValue 1       'muestra Tilt en la pantalla
        If B2SOn then
            Controller.B2SSetTilt 1
        end if
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
        LeftSlingshot.Disabled = 1
        DOF 101, DOFOff
        DOF 102, DOFOff
    Else
        'enciende de nuevo todas las luces GI
        GiOn

        Bumper1.Threshold = 1
        LeftSlingshot.Disabled = 0
    End If
End Sub

Sub TiltRecoveryTimer_Timer()
    ' si todas las bolas se han colado, entonces ..
    If(BallsOnPlayfield = 0)Then
        '... haz el fin de bola normal
        EndOfBall()
        TiltRecoveryTimer.Enabled = False
    End If
' de lo contrario la rutina mirará si todavía hay bolas en la mesa
End Sub

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

'***************************************
' Sonidos de las colecciones de objetos
' como metales, gomas, plásticos, etc
'***************************************

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

    ' enciende las luces GI si estuvieran apagadas
    GiOn

    ' Start the music if enabled
    If bMusicON Then PlaySound"Musica", -1, bMusicVOL

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
    BonusMultiplier = 1
  ScoreMultiplier = 1
    Bonus = 0
    UpdateBallInPlay

If bFreePlay = False Then
    Match = INT(RND(1) * 8) * 10 + 10 'empieza con un número aleatorio de 10 a 90 para la loteria
    Clear_Match
End If

    ' inicializa otras variables
    Tilt = 0

    ' inicializa las variables del juego
    Game_Init()

    ' ahora puedes empezar una música si quieres
    ' empieza la rutina "Firstball" despues de una pequeña pausa
    vpmtimer.addtimer 2000, "FirstBall '"
    If Credits = 0 Then DOF 200, DOFOff
End Sub

' esta pausa es para que la mesa tenga tiempo de poner los marcadores a cero y actualizar las luces

Sub FirstBall
    'debug.print "FirstBall"
    ' ajusta la mesa para una bola nueva, sube las dianas abatibles, etc
    ResetForNewPlayerBall()
    ' crea una bola nueva en la zona del plunger
    CreateNewBall()
    DOF 250, DOFPulse
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
    BallRelease.CreateSizedBallWithMass BallSize/2, BallMass

    ' incrementa el número de bolas en el tablero, ya que hay que contarlas
    BallsOnPlayfield = BallsOnPlayfield + 1

    ' actualiza las luces del backdrop
    UpdateBallInPlay

    ' y expulsa la bola
    PlaySoundAt SoundFXDOF ("fx_Ballrel", 104, DOFPulse, DOFContactors), BallRelease
    If Epileptikdof = True Then DOF 219, DOFPulse End If
    BallRelease.Kick 90, 4
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

    If NOT Tilted Then
        BonusCountTimer.Interval = 250
        BonusCountTimer.Enabled = 1
    Else 'Si hay falta simplemente espera un momento y va directo a la segunta parte después de perder la bola
        vpmtimer.addtimer 400, "EndOfBall2 '"
    End If
End Sub

Sub BonusCountTimer_Timer 'añade los bonos y actualiza las luces
    'debug.print "BonusCount_Timer"
    If Bonus >= 1 Then
        Bonus = Bonus -1
        AddScore 1000 * BonusMultiplier
        UpdateBonusLights
    Else
    If Bonus = 0.5 Then
        Bonus = 0
        AddScore 500 * BonusMultiplier
        UpdateBonusLights
  Else
        ' termina la cuenta de los bonos y continúa con el fin de bola
        BonusCountTimer.Enabled = 0
        vpmtimer.addtimer 1000, "EndOfBall2 '"
    End If
  End If
End Sub

Sub UpdateBonusLights
'    For each x in aBonusLights
'        If x.State = 1 Then
'            x.State = 0
'            Exit Sub
'        End If
'    Next
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
    DisableTable False 'activa de nuevo los bumpers y los slingshots

    ' ¿ha ganado el jugador una bola extra?
    If(ExtraBallsAwards(CurrentPlayer) > 0)Then
        'debug.print "Extra Ball"

        ' sí? entonces se la das al jugador
        ExtraBallsAwards(CurrentPlayer) = ExtraBallsAwards(CurrentPlayer)- 1

        ' si no hay más bolas apaga la luz de jugar de nuevo
        If(ExtraBallsAwards(CurrentPlayer) = 0)Then
            LightShootAgain.State = 0
            If B2SOn then
                Controller.B2SSetShootAgain 0
            end if
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
        ' Controller.B2SSetCanPlay 0
    end if
    ' asegúrate de que los flippers están en modo de reposo
    SolLFlipper 0
    SolRFlipper 0
    DOF 250, DOFPulse
    DOF 201, DOFOff

    ' pon las luces en el modo de fin de juego
    StartAttractMode
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

' Esta función calcula el Highscore
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
            If sortscores(sj)> sortscores(sj + 1) then
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
    PlaySound "fx_match"
    Display_Match
    If(Score(CurrentPlayer)MOD 100) = Match Then
        PlaySound SoundFXDOF  ("fx_knocker", 300, DOFPulse, DOFKnocker)
        DOF 230, DOFPulse
        AddCredits 1
        DOF 200, DOFOn
    End If
End Sub

Sub Clear_Match()
    For each x in aMatchLights
        x.State = 2
    Next
    If B2SOn then
        ' Controller.B2SSetMatch 0
    end if
End Sub

Sub Display_Match()
    For each x in aMatchLights
        x.State = 0
    Next
    aMatchLights(Match \ 10).state = 1
    If B2SOn then
        If Match = 0 then
            ' Controller.B2SSetMatch 100
        else
            ' Controller.B2SSetMatch Match
        end if
    end if
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

    BallsOnPlayfield = BallsOnPlayfield - 1

    ' haz sonar el ruido de la bola
    PlaySoundAt "fx_drain", Drain
    DOF 218, DOFPulse

    'si hay falta el systema de tilt se encargará de continuar con la siguente bola/jugador
    If Tilted Then
        Exit Sub
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
            ' añade los puntos a la variable del actual jugador
            Score(CurrentPlayer) = Score(CurrentPlayer) + points
            ' actualiza los contadores
            UpdateScore
           ' hace sonar las campanillas de acuerdo a los puntos obtenidos
    Select Case Points
        Case 50 : PlaySound "50"
        Case 100 : PlaySound "100"
        Case 150 : PlaySound "50"
        Case 200 : PlaySound "100"
        Case 250 : PlaySound "50"
     Case 500 : PlaySound "500"
        Case 1000 : PlaySound "1000"
        Case 1500 : PlaySound "500"
        Case 2000 : PlaySound "1000"
        Case 3000 : PlaySound "1000"
        Case 4000 : PlaySound "1000"
    End Select

    ' ' aquí se puede hacer un chequeo si el jugador ha ganado alguna puntuación alta y darle un crédito ó bola extra
'    If Score(CurrentPlayer) >= Special1 AND Special1Awarded(CurrentPlayer) = False Then
'        AwardSpecial
'        Special1Awarded(CurrentPlayer) = True
'    End If
'    If Score(CurrentPlayer) >= Special2 AND Special2Awarded(CurrentPlayer) = False Then
'        AwardSpecial
'        Special2Awarded(CurrentPlayer) = True
'    End If
'    If Score(CurrentPlayer) >= Special3 AND Special3Awarded(CurrentPlayer) = False Then
'        AwardSpecial
'        Special3Awarded(CurrentPlayer) = True
'    End If
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
    ' actualiza las luces
    'UpdateBonusLights
    End if
End Sub

'***********************************************************************************
'        Score reels - puntuaciones - y actualiza otras luces del backdrop
'***********************************************************************************
'esta es al rutina que actualiza la puntuación del jugador

Sub UpdateScore  'en este juego solo hay un score reel
  ScoreReel1.Setvalue Score(CurrentPlayer)

    If B2SOn then
        Controller.B2SSetScorePlayer 1, Score(CurrentPlayer)
        If Score(CurrentPlayer) >= 100000 then
            ' Controller.B2SSetScoreRollover 24 + CurrentPlayer, 1
        end if
    end if
End Sub

' pone todos los marcadores a 0
Sub ResetScores
    ScoreReel1.ResetToZero

    If B2SOn then
        Controller.B2SSetScorePlayer 1, 0
        ' Controller.B2SSetScoreRolloverPlayer1 0
    end if
End Sub

Sub AddCredits(value)
    If Credits < 9 Then
        Credits = Credits + value
        ' CreditReel.SetValue Credits
        UpdateCredits
    end if
End Sub

Sub UpdateCredits
    If Credits > 0 Then 'this is for Bally tables
    'CreditLight.State = 1
    Else
    'CreditLight.State = 0
    End If
    ' CreditReel.SetValue credits
    If B2SOn then
        ' Controller.B2SSetCredits Credits
    ' no Credits Reels on the B2S ?
    end if
End Sub

Sub UpdateBallInPlay 'actualiza los marcadores de las bolas, el número de jugador y el número total de jugadores
    Select Case Balls
        Case 1:ball1.State = 1:ball2.State = 0:ball3.State = 0:ball4.State = 0:ball5.State = 0
        Case 2:ball1.State = 0:ball2.State = 1:ball3.State = 0:ball4.State = 0:ball5.State = 0
        Case 3:ball1.State = 0:ball2.State = 0:ball3.State = 1:ball4.State = 0:ball5.State = 0
        Case 4:ball1.State = 0:ball2.State = 0:ball3.State = 0:ball4.State = 1:ball5.State = 0
        Case 5:ball1.State = 0:ball2.State = 0:ball3.State = 0:ball4.State = 0:ball5.State = 1
    End Select
    If B2SOn then
        Controller.B2SSetBallInPlay Balls
    end if
    Select Case CurrentPlayer
        Case 1:pl1.State = 1:pl2.State = 0:pl3.State = 0:pl4.State = 0
        Case 2:pl1.State = 0:pl2.State = 1:pl3.State = 0:pl4.State = 0
        Case 3:pl1.State = 0:pl2.State = 0:pl3.State = 1:pl4.State = 0
        Case 4:pl1.State = 0:pl2.State = 0:pl3.State = 0:pl4.State = 1
    End Select
    If B2SOn then
        Controller.B2SSetPlayerUp CurrentPlayer
    end if
End Sub

Sub UpdatePlayersCanPlay
    Select Case PlayersPlayingGame
        Case 1:Flashforms pl1, 500, 10, 0
        Case 2:Flashforms pl2, 500, 10, 0
        Case 3:Flashforms pl3, 500, 10, 0
        Case 4:Flashforms pl4, 500, 10, 0
    End Select
    If B2SOn then
        ii=1:jj=0:NumPlay.enabled=true
    end if
End Sub

'*************************
' Partidas y bolas extras
'*************************

Sub AwardExtraBall()
    If NOT bExtraBallWonThisBall Then
        PlaySound SoundFXDOF ("fx_knocker", 300, DOFPulse, DOFKnocker)
        DOF 230, DOFPulse
        ExtraBallsAwards(CurrentPlayer) = ExtraBallsAwards(CurrentPlayer) + 1
        bExtraBallWonThisBall = True
        LightShootAgain.State = 1
        If B2SOn then
            Controller.B2SSetShootAgain 1
        end if
    END If
End Sub

Sub AwardSpecial()
    PlaySound SoundFXDOF ("fx_knocker", 300, DOFPulse, DOFKnocker)
    DOF 230, DOFPulse
    AddCredits 1
End Sub

' ********************************
'        Attract Mode
' ********************************
' las luces simplemente parpadean de acuerdo a los valores que hemos puesto en el "Blink Pattern" de cada luz

Sub StartAttractMode()
    bAttractMode = True
    For each x in aLights
        x.State = 2
    Next
    pl1.State = 0:pl2.State = 0:pl3.State = 0:pl4.State = 0
    If B2SOn then
        Controller.B2SSetPlayerUp 0
    end if
    ' enciente la luz de fin de partida
    GameOverR.SetValue 1
    ' empieza un timer
    AttractTimer.Enabled = 1
    If Credits > 0 Then DOF 200, DOFOn
End Sub

Sub StopAttractMode()
    bAttractMode = False
    ResetScores
    For each x in aLights
        x.State = 0
    Next

    ' apaga la luz de fin de partida
    GameOverR.SetValue 0
  highscoreR.SetValue 0
    If B2SOn then
        Controller.B2SSetGameOver 0
    Controller.B2SSetData 99,0
    end if
    ' apaga timer
    AttractTimer.Enabled = 0
End Sub

Dim AttractStep
AttractStep = 0

Sub AttractTimer_Timer
    Select Case AttractStep
        Case 0
            AttractStep = 1
            ScoreReel1.SetValue Score(1)
      pl1.State = 1
      highscoreR.SetValue 0
            If B2SOn then
                Controller.B2SSetScorePlayer1 Score(1)
        Controller.B2SSetPlayerUp 1
        Controller.B2SSetData 99,0
            end if
        Case 1
            AttractStep = 2
            ScoreReel1.SetValue Score(2)
      pl1.State = 0: pl2.State = 1
            If B2SOn then
                Controller.B2SSetScorePlayer1 Score(2)
        Controller.B2SSetPlayerUp 2
            end if
        Case 2
            AttractStep = 3
            ScoreReel1.SetValue Score(3)
      pl2.State = 0: pl3.State = 1
            If B2SOn then
                Controller.B2SSetScorePlayer1 Score(3)
        Controller.B2SSetPlayerUp 3
            end if
        Case 3
            AttractStep = 4
            ScoreReel1.SetValue Score(4)
      pl3.State = 0: pl4.State = 1
            If B2SOn then
                Controller.B2SSetScorePlayer1 Score(4)
        Controller.B2SSetPlayerUp 4
            end if
        Case 4
            AttractStep =0
            ScoreReel1.SetValue Highscore
      pl4.State = 0
      highscoreR.SetValue 1
            If B2SOn then
                Controller.B2SSetScorePlayer1 Highscore
        Controller.B2SSetPlayerUp 0
        Controller.B2SSetData 99,1
            end if
    End Select
End Sub

'************************************************
'    Load (cargar) / Save (guardar)/ Highscore
'************************************************

' solamente guardamos el número de créditos y la puntuación más alta

Sub Loadhs
    x = LoadValue(TableName, "HighScore1")
    If(x <> "")Then HSScore(1) = CDbl(x) End If
    x = LoadValue(TableName, "HighScore2")
    If(x <> "")Then HSScore(2) = CDbl(x) End If
    x = LoadValue(TableName, "HighScore3")
    If(x <> "")Then HSScore(3) = CDbl(x) End If
    x = LoadValue(TableName, "HighScore4")
    If(x <> "")Then HSScore(4) = CDbl(x) End If
    x = LoadValue(TableName, "HighScore5")
    If(x <> "")Then HSScore(5) = CDbl(x) End If

    x = LoadValue(TableName, "HighScore1Name")
    If(x <> "")Then HSName(1) = x End If
    x = LoadValue(TableName, "HighScore2Name")
    If(x <> "")Then HSName(2) = x End If
    x = LoadValue(TableName, "HighScore3Name")
    If(x <> "")Then HSName(3) = x End If
    x = LoadValue(TableName, "HighScore4Name")
    If(x <> "")Then HSName(4) = x End If
    x = LoadValue(TableName, "HighScore5Name")
    If(x <> "")Then HSName(5) = x End If
    x = LoadValue(TableName, "Credits")
    If(x <> "")then Credits = CInt(x)Else Credits = 0 End If
End Sub

Sub Savehs
    SaveValue TableName, "HighScore1", HSScore(1)
    SaveValue TableName, "HighScore2", HSScore(2)
    SaveValue TableName, "HighScore3", HSScore(3)
    SaveValue TableName, "HighScore4", HSScore(4)
    SaveValue TableName, "HighScore5", HSScore(5)
    SaveValue TableName, "HighScore1Name", HSName(1)
    SaveValue TableName, "HighScore2Name", HSName(2)
    SaveValue TableName, "HighScore3Name", HSName(3)
    SaveValue TableName, "HighScore4Name", HSName(4)
    SaveValue TableName, "HighScore5Name", HSName(5)
    SaveValue TableName, "Credits", Credits
End Sub

' por si se necesitara quitar la actual puntuación más alta, se le puede poner a una tecla,
' o simplemente abres la ventana de debug y escribes Reseths y le das al enter
Sub Reseths
    HSScore(1) = 45510
    HSScore(2) = 44000
    HSScore(3) = 43000
    HSScore(4) = 42000
    HSScore(5) = 41000

    HSName(1) = "AAA"
    HSName(2) = "ZZZ"
    HSName(3) = "XXX"
    HSName(4) = "ABC"
    HSName(5) = "BBB"
    Savehs
End Sub

'****************************************
' Actualizaciones en tiempo real
'****************************************
' se usa sobre todo para hacer animaciones o sonidos que cambian en tiempo real
' como por ejemplo para sincronizar los flipers, puertas ó molinillos con primitivas

Sub GameTimer_Timer
    RollingUpdate 'actualiza el sonido de la bola rodando
    ' y también algunas animaciones, sobre todo de primitivas
End Sub

'***********************************************************************
' *********************************************************************
'            Aquí empieza el código particular a la mesa
' (hasta ahora todas las rutinas han sido muy generales para todas las mesas)
' (y hay muy pocas rutinas que necesitan cambiar de mesa a mesa)
' *********************************************************************
'***********************************************************************

' se inicia las dianas abatibles, primitivas, etc.
' aunque en el VPX no hay muchos objetos que necesitan ser iniciados

Sub VPObjects_Init 'en esta mesa no hay nada que necesite iniciarse, pero dejamos la rutina para próximas mesas
End Sub

' variables de la mesa
Dim SpecialReady

Sub Game_Init() 'esta rutina se llama al principio de un nuevo juego

    'Empezar alguna música, si hubiera música en esta mesa

    'iniciar variables, en esta mesa hay muy pocas variables ya que usamos las luces, y el UserValue de las dianas

    'iniciar algún timer

    'Iniciar algunas luces, en esta mesa son las mismas luces que las de una bola nueva
    TurnOffPlayfieldLights()
End Sub

Sub ResetNewBallVariables()                'inicia las variable para una bola ó jugador nuevo
    Bonus = 0
    BonusMultiplier = 1                        'multiplicador de los bonos
    SpecialReady = 0
  ScoreMultiplier = 1
End Sub

Sub ResetNewBallLights()                   'enciende ó apaga las luces para una bola nueva
    TurnOffPlayfieldLights()
End Sub

Sub TurnOffPlayfieldLights()
    For each x in aLights
        x.State = 0
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
Dim LStep

Sub LeftSlingShot_Slingshot
    If Tilted Then Exit Sub
    PlaySoundAtBall SoundFXDOF ("fx_slingshot",103,DOFPulse,DOFcontactors)
    DOF 203, DOFPulse
    If Epileptikdof = True Then DOF 204, DOFPulse End If
    If Epileptikdof = True Then DOF 202, DOFPulse End If
    LeftSling4.Visible = 1
    Lemk.RotX = 26
    LStep = 0
    LeftSlingShot.TimerEnabled = True
    ' añade algunos puntos
    AddScore 100 * ScoreMultiplier
  AddBonus 0.5
' añade algún efecto a la mesa
End Sub

Sub LeftSlingShot_Timer
    Select Case LStep
        Case 1:LeftSLing4.Visible = 0:LeftSLing3.Visible = 1:Lemk.RotX = 14
        Case 2:LeftSLing3.Visible = 0:LeftSLing2.Visible = 1:Lemk.RotX = 2
        Case 3:LeftSLing2.Visible = 0:Lemk.RotX = -10:LeftSlingShot.TimerEnabled = 0
    End Select
    LStep = LStep + 1
End Sub


'*********
' Bumpers
'*********

Sub Bumper1_Hit
    If NOT Tilted Then
    PlaySoundAt SoundFXDOF ("fx_Bumper",105,DOFPulse,DOFContactors), bumper1
    DOF 205, DOFPulse
    If Epileptikdof = True Then DOF 202, DOFPulse End If
        ' añade algunos puntos
        AddScore 1000
    End If
End Sub

'**************************
'  Gomas de 10 y 100 puntos
'**************************

Sub RubberBand3_Hit
    If NOT Tilted Then
        ' añade puntos
        AddScore 50 * ScoreMultiplier
    End If
End Sub

Sub RubberBand8_Hit
    If NOT Tilted Then
        ' añade puntos
        AddScore 50 * ScoreMultiplier
    AddBonus 0.5
    End If
End Sub

Sub RubberBand10_Hit
    If NOT Tilted Then
        ' añade puntos
        AddScore 50 * ScoreMultiplier
    End If
End Sub

Sub RubberBand16_Hit
    If NOT Tilted Then
        ' añade puntos
        AddScore 100 * ScoreMultiplier
    AddBonus 0.5
    End If
End Sub

Sub RubberBand15_Hit
    If NOT Tilted Then
        ' añade puntos
        AddScore 50 * ScoreMultiplier
    End If
End Sub

'************************
'     Dianas fijas
'************************

Sub Target1_hit
    PlaySoundAtBall SoundFXDOF ("fx_target", 106, DOFPulse, DOFTargets)
    If Tilted Then Exit Sub
    DOF 207, DOFPulse
    AddScore 100 * ScoreMultiplier
AddBonus 1
' check?
li2.State = 1
SetScoreMultiplier
End Sub

Sub Target2_hit
    PlaySoundAtBall SoundFXDOF ("fx_target", 107, DOFPulse, DOFTargets)
    If Tilted Then Exit Sub
    DOF 208, DOFPulse
    AddScore 100 * ScoreMultiplier
AddBonus 1
' check?
li3.State = 1
SetScoreMultiplier
End Sub

Sub Target3_hit
    PlaySoundAtBall SoundFXDOF ("fx_target", 108, DOFPulse, DOFTargets)
    If Tilted Then Exit Sub
    DOF 209, DOFPulse
    AddScore 100 * ScoreMultiplier
AddBonus 1
' check?
li7.State = 1
SetScoreMultiplier
End Sub

Sub Target4_hit
    PlaySoundAtBall SoundFXDOF ("fx_target", 108, DOFPulse, DOFTargets)
    If Tilted Then Exit Sub
    DOF 209, DOFPulse
    AddScore 100 * ScoreMultiplier
AddBonus 1
' check?
li8.State = 1
SetScoreMultiplier
End Sub

Sub Target5_hit
    PlaySoundAtBall SoundFXDOF ("fx_target", 108, DOFPulse, DOFTargets)
    If Tilted Then Exit Sub
    DOF 209, DOFPulse
    AddScore 100 * ScoreMultiplier
AddBonus 1
' check?
li9.State = 1
SetScoreMultiplier
End Sub

Sub SetScoreMultiplier
Dim tmp
tmp = li2.State + li3.State +li7.State +li8.State +li9.State
Select Case tmp
Case 1: ScoreMultiplier = 2
Case 2: ScoreMultiplier = 5
Case 3: ScoreMultiplier = 10
Case 4: ScoreMultiplier = 20
Case 5: ScoreMultiplier = 40
End Select

End Sub

'******************************
'         Pasillos
'******************************

Sub Trigger1_Hit
    PlaySoundAt "fx_sensor", Trigger1
    If Tilted Then Exit Sub
    DOF 210, DOFPulse
    AddScore 100 * ScoreMultiplier
    AddBonus 1
End Sub

Sub Trigger2_Hit
    PlaySoundAt "fx_sensor", Trigger2
    If Tilted Then Exit Sub
    DOF 211, DOFPulse
    AddScore 100 * ScoreMultiplier
    AddBonus 1
End Sub

Sub Trigger3_Hit
    PlaySoundAt "fx_sensor", Trigger3
    If Tilted Then Exit Sub
    DOF 212, DOFPulse
    AddScore 100 * ScoreMultiplier
    AddBonus 1
End Sub

Sub Trigger4_Hit
    PlaySoundAt "fx_sensor", Trigger4
    If Tilted Then Exit Sub
    DOF 213, DOFPulse
    AddScore 50 * ScoreMultiplier
  If li10.State = 0 Then
    li10.State = 1
  End If
End Sub

Sub Trigger5_Hit
    PlaySoundAt "fx_sensor", Trigger5
    If Tilted Then Exit Sub
    DOF 214, DOFPulse
    AddScore 100 * ScoreMultiplier
  If li10.State = 1 Then
    li10.State = 2
    li4.State = 1
    BonusMultiplier = 2
    'PlayTune
  End If
End Sub

Sub Trigger6_Hit
    PlaySoundAt "fx_sensor", Trigger6
    If Tilted Then Exit Sub
    DOF 215, DOFPulse
    AddScore 50 * ScoreMultiplier
  If li10.State = 0 Then
    li10.State = 1
  End If
End Sub

Sub Trigger7_Hit
    PlaySoundAt "fx_sensor", Trigger7
    If Tilted Then Exit Sub
    DOF 216, DOfPulse
    AddScore 100 * ScoreMultiplier
  If li4.State = 1 Then
    li4.State = 2
    BonusMultiplier = 3
    'PlayTune
  End If
End Sub


'****************
'    spinner
'****************

Sub Spinner1_Spin
    PlaySoundAt "fx_spinner", Spinner1
    If Tilted Then Exit Sub
    DOF 217, DOFPulse
    AddScore 50 * ScoreMultiplier
  AddBonus 0.5 'solo 500 puntos de bonus
End Sub

' ***************************************************
' GNMOD - Multiple High Score Display and Collection
' jpsalas: changed ramps by flashers to remove extra shadow
' ***************************************************

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
HSScore(1) = 45510
HSScore(2) = 44000
HSScore(3) = 43000
HSScore(4) = 42000
HSScore(5) = 41000

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
        If HsTimerCount> 5 then
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

    For xFor = StartHSArray(LineNo) to EndHSArray(LineNo)
        Eval("HS" &xFor).imageA = GetHSChar(String, Index)
        Index = Index + 1
    Next
End Sub

Sub NewHighScore(NewScore, PlayNum)
    If NewScore> HSScore(5) then
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
        If AlphaStringPos <1 then
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
        If AlphaStringPos> len(AlphaString) or(AlphaStringPos = len(AlphaString) and InitialString = "") then
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
            InitialString = MID(InitialString, 1, len(InitialString) - 1)
            If len(InitialString) = 0 then
                ' If there are no more characters to back over, don't leave the < displayed
                AlphaStringPos = 1
            End If
            PlaySound("fx_Esc")
        Else
            InitialString = InitialString & SelectedChar
            PlaySound("fx_Enter")
        End If
        If len(InitialString) <3 then
            SetHSLine 3, InitialString & SelectedChar
        End If
    End If
    If len(InitialString) = 3 then
        ' save the score
        For i = 5 to 1 step -1
            If i = 1 or(HSNewHigh> HSScore(i) and HSNewHigh <= HSScore(i - 1) ) then
                ' Replace the score at this location
                If i <5 then
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
            ElseIf i <5 then
                ' move the score in this slot down by 1, it's been exceeded by the new score
                HSScore(i + 1) = HSScore(i)
                HSName(i + 1) = HSName(i)
            End If
        Next
    End If
End Sub
' End GNMOD

'*********************************
' Table Options F12 User Options
'*********************************
' Table1.Option arguments are:
' - option name, minimum value, maximum value, step between valid values, default value, unit (0=None, 1=Percent), an optional array of literal strings

Dim LUTImage, BallsPerGame

Sub Table1_OptionEvent(ByVal eventId)
    Dim x, y

    'LUT
    LutImage = Table1.Option("Select LUT", 0, 21, 1, 0, 0, Array("Normal 0", "Normal 1", "Normal 2", "Normal 3", "Normal 4", "Normal 5", "Normal 6", "Normal 7", "Normal 8", "Normal 9", "Normal 10", _
        "Warm 0", "Warm 1", "Warm 2", "Warm 3", "Warm 4", "Warm 5", "Warm 6", "Warm 7", "Warm 8", "Warm 9", "Warm 10") )
    UpdateLUT

    ' Cabinet rails
    x = Table1.Option("Cabinet Rails", 0, 1, 1, 1, 0, Array("Hide", "Show") )
    For each y in aRails:y.visible = x:next

    ' Balls per Game
    x = Table1.Option("Balls per Game", 0, 1, 1, 1, 0, Array("3 Balls", "5 Balls") )
    If x = 1 Then BallsPerGame = 5 Else BallsPerGame = 3
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

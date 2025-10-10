' ****************************************************************
'     Summer Time / IPD No. 2415 / October 16, 1972 / 1 Player
'         Williams Electronics, Incorporated (1967-1985)
' Summer Time video: https://www.youtube.com/watch?v=fyQHfYAjgUk
'           VPX8 table by jpsalas, 2025, version 6.0.1
' ****************************************************************
'
Option Explicit
Randomize

' Valores Constantes de las físicas de los flippers - se usan en la creación de las bolas, tienen que cargarse antes del core.vbs
Const BallSize = 25 ' el radio de la bola. el tamaño normal es 50 unidades de VP
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
Const TableName = "SummerTime" 'saving the highscores and credits
Const cGameName = "SummerTime" ' mostly for DOF
Const MaxPlayers = 1           ' only 1 in this table
Const Special1 = 200000        ' puntuación a obtener para partida extra
Const Special2 = 400000
Const Special3 = 700000

' Variables Globales
Dim PlayersPlayingGame
Dim CurrentPlayer
Dim Credits
Dim Special1Awarded(4)
Dim Special2Awarded(4)
Dim Special3Awarded(4)
Dim Score(4)
Dim HighScore
Dim Tilt
Dim TiltSensitivity
Dim Tilted
Dim Add10
Dim Add100
Dim Add1000
Dim Add10000

' Variables de control
Dim BallsOnPlayfield

' Variables de tipo Boolean (verdadero ó falso, True ó False)
Dim bAttractMode
Dim bFreePlay
Dim bGameInPlay
Dim bOnTheFirstBall
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

    ' Carga los valores grabados highscore y créditos
    Loadhs

    UpdateCredits

    ' Juego libre o con monedas: si es True entonces no se usarán monedas
    bFreePlay = False 'queremos monedas

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
    Add10000 = 0

    ' pone la mesa en modo de espera
    EndOfGame

    'Enciende las luces GI despues de un segundo
    vpmtimer.addtimer 2000, "GiOn '"

    ' Quita las puntuaciones cuando la mesa se juega en modo FS
    If Table1.ShowDT = true then
        For each x in aReels
            x.Visible = 1
        Next
    else
        For each x in aReels
            x.Visible = 0
        Next
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
        If(Tilted = False) Then
            If Not bFreePlay Then DOF 200, DOFOn
            AddCredits 1
            If Credits < 9 then
                PlaySound "fx_coin"
            Else
                PlaySound "fx_coin2"
            End If
        End If
    End If

    ' el plunger
    If keycode = PlungerKey Then
        Plunger.Pullback
        PlaySoundAt "fx_plungerpull", plunger
    End If

    If keycode = LeftTiltKey Then Nudge 90, 8:PlaySound "fx_nudge", 0, 1, -0.1, 0.25
    If keycode = RightTiltKey Then Nudge 270, 8:PlaySound "fx_nudge", 0, 1, 0.1, 0.25
    If keycode = CenterTiltKey Then Nudge 0, 9:PlaySound "fx_nudge", 0, 1, 1, 0.25

    ' Funcionamiento normal de los flipers y otras teclas durante el juego

    If bGameInPlay AND NOT Tilted Then
        ' teclas de la falta
        If keycode = LeftTiltKey Then CheckTilt
        If keycode = RightTiltKey Then CheckTilt
        If keycode = CenterTiltKey Then CheckTilt
        If keycode = MechanicalTilt Then CheckTilt

        ' teclas de los flipers
        If keycode = LeftFlipperKey Then SolLFlipper 1
        If keycode = RightFlipperKey Then SolRFlipper 1

        ' tecla de empezar el juego
        If keycode = StartGameKey Then
            If((PlayersPlayingGame < MaxPlayers) AND(bOnTheFirstBall = True) ) Then

                If(bFreePlay = True) Then
                    PlayersPlayingGame = PlayersPlayingGame + 1
                    PlaySound "GameStart"
                Else
                    If(Credits> 0) then
                        DOF 200, DOFOn
                        PlayersPlayingGame = PlayersPlayingGame + 1
                        Credits = Credits - 1
                        UpdateCredits
                    Else
                        ' no hay suficientes créditos para empezar el juego.
                        DOF 200, DOFOff
                    End If
                End If
            End If
        End If
        Else ' If (GameInPlay)

            If keycode = StartGameKey Then
                If(bFreePlay = True) Then
                    If(BallsOnPlayfield = 0) Then
                        ResetScores
                        ResetForNewGame()
                        PlaySound "GameStart"
                    End If
                Else
                    If(Credits> 0) Then
                        DOF 200, DOFOn
                        If(BallsOnPlayfield = 0) Then
                            Credits = Credits - 1
                            UpdateCredits
                            ResetScores
                            ResetForNewGame()
                            PlaySound "GameStart"
                        End If
                    Else
                        ' Not Enough Credits to start a game.
                        DOF 200, DOFOff
                    End If
                End If
            End If
    End If ' If (GameInPlay)
End Sub

Sub Table1_KeyUp(ByVal keycode)

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
    If B2SOn Then Controller.Stop
End Sub

'********************
'     Flippers
'********************

Sub SolLFlipper(Enabled)
    If Enabled Then
        PlaySoundAt SoundFXDOF("fx_flipperup", 101, DOFOn, DOFFlippers), LeftFlipper
        LeftFlipper.RotateToEnd
        LeftFlipperOn = 1
    Else
        PlaySoundAt SoundFXDOF("fx_flipperdown", 101, DOFOff, DOFFlippers), LeftFlipper
        LeftFlipper.RotateToStart
        LeftFlipperOn = 0
    End If
End Sub

Sub SolRFlipper(Enabled)
    If Enabled Then
        PlaySoundAt SoundFXDOF("fx_flipperup", 102, DOFOn, DOFFlippers), RightFlipper
        RightFlipper.RotateToEnd
        RightFlipperOn = 1
    Else
        PlaySoundAt SoundFXDOF("fx_flipperdown", 102, DOFOff, DOFFlippers), RightFlipper
        RightFlipper.RotateToStart
        RightFlipperOn = 0
    End If
End Sub

' el sonido de la bola golpeando los flipers

Sub LeftFlipper_Collide(parm)
    PlaySound "fx_rubber_flipper", 0, Vol(ActiveBall), pan(ActiveBall), 0.2, 0, 0, 0, AudioFade(ActiveBall)
End Sub

Sub RightFlipper_Collide(parm)
    PlaySound "fx_rubber_flipper", 0, Vol(ActiveBall), pan(ActiveBall), 0.2, 0, 0, 0, AudioFade(ActiveBall)
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

'*******************
' Luces GI
'*******************

Sub GiOn 'enciende las luces GI
    Dim bulb
    For each bulb in aGiLights
        bulb.State = 1
        ' Undercab ON
        DOF 201, DOFOn
    Next
    If B2SOn Then Controller.B2SSetData 60, 1
End Sub

Sub GiOff 'apaga las luces GI
    Dim bulb
    For each bulb in aGiLights
        bulb.State = 0
        ' Undercab OFF
        DOF 201, DOFOff
    Next
    If B2SOn Then Controller.B2SSetData 60, 0
End Sub

Sub TurnON(Col) 'turn on all the lights in a collection
    Dim i
    For each i in Col
        i.State = 1
    Next
End Sub

Sub TurnOFF(Col) 'turn off all the lights in a collection
    Dim i
    For each i in Col
        i.State = 0
    Next
End Sub

'**************
' TILT - Falta
'**************

'el "timer" TiltDecreaseTimer resta .01 de la variable "Tilt" cada ronda

Sub CheckTilt                     'esta rutina se llama cada vez que das un golpe a la mesa
    Tilt = Tilt + TiltSensitivity 'añade un valor al contador "Tilt"
    TiltDecreaseTimer.Enabled = True
    If Tilt> 15 Then              'Si la variable "Tilt" es más de 15 entonces haz falta
        Tilted = True
        TurnOn aTilt              'muestra Tilt en la pantalla
        If B2SOn then
            Controller.B2SSetTilt 1
        end if
        DisableTable True
        TiltRecoveryTimer.Enabled = True 'empieza una pausa a fin de que todas las bolas se cuelen
        ClockValue = ClockValue - 3
        UpdateClock
    End If
End Sub

Sub TiltDecreaseTimer_Timer
    ' DecreaseTilt
    If Tilt> 0 Then
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
        Bumper001.Threshold = 100
        Bumper002.Threshold = 100
        Bumper003.Threshold = 100
    Else
        'enciende de nuevo todas las luces GI
        GiOn
        Bumper001.Threshold = 1
        Bumper002.Threshold = 1
        Bumper003.Threshold = 1
    End If
End Sub

Sub TiltRecoveryTimer_Timer()
    ' si todas las bolas se han colado, entonces ..
    If(BallsOnPlayfield = 0) Then
        '... haz el fin de bola normal
        EndOfBall2()
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
    RollingTimer.Enabled = 1
End Sub

Sub RollingTimer_Timer()
    Dim BOT, b, ballpitch, ballvol, speedfactorx, speedfactory
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
        If BallVel(BOT(b) )> 1 Then
            If BOT(b).z < 30 Then
                ballpitch = Pitch(BOT(b) )
                ballvol = Vol(BOT(b) )
            Else
                ballpitch = Pitch(BOT(b) ) + 50000 'increase the pitch on a ramp
                ballvol = Vol(BOT(b) ) * 5
            End If
            rolling(b) = True
            PlaySound("fx_ballrolling" & b), -1, ballvol, Pan(BOT(b) ), 0, ballpitch, 1, 0, AudioFade(BOT(b) )
        Else
            If rolling(b) = True Then
                StopSound("fx_ballrolling" & b)
                rolling(b) = False
            End If
        End If

        ' rothbauerw's Dropping Sounds
        If BOT(b).VelZ < -1 and BOT(b).z < 55 and BOT(b).z> 27 Then 'height adjust for ball drop sounds
            PlaySound "fx_balldrop", 0, ABS(BOT(b).velz) / 17, Pan(BOT(b) ), 0, Pitch(BOT(b) ), 1, 0, AudioFade(BOT(b) )
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

'*****************************
' Sonido de las bolas chocando
'*****************************

Sub OnBallBallCollision(ball1, ball2, velocity)
    PlaySound "fx_collide", 0, Csng(velocity) ^2 / 2000, Pan(ball1), 0, Pitch(ball1), 0, 0, AudioFade(ball1)
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

'************************************************************************************************************************
' Solo para VPX 10.8 y posteriores.
' FlashForMs hará parpadear una luz por unos milisegundos "TotalPeriod" cada tantos milisegundos "BlinkPeriod"
' Cuando el "TotalPeriod" haya terminado, la luz o el flasher se pondrá en el estado especificado por el valor "FinalState"
' El valor de "FinalState" puede ser: 0=apagado, 1=encendido, 2=regreso al estado anterior
'************************************************************************************************************************

Sub FlashForMs(MyLight, TotalPeriod, BlinkPeriod, FinalState)
    If FinalState = 2 Then
        FinalState = MyLight.State 'guarda el estado actual de la luz
    End If
    MyLight.BlinkInterval = BlinkPeriod
    MyLight.Duration 2, TotalPeriod, FinalState
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
    ' enciende las luces GI si estuvieran apagadas
    GiOn

    CurrentPlayer = 1
    PlayersPlayingGame = 1
    bOnTheFirstBall = True
    For i = 1 To MaxPlayers
        Score(i) = 0
        Special1Awarded(i) = False
        Special2Awarded(i) = False
        Special3Awarded(i) = False
    Next

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
    ' enciende las luces, reinicializa las variables del juego, etc
    ResetNewBallLights
    ResetNewBallVariables
End Sub

' Crea una bola nueva en la mesa

Sub CreateNewBall()
    ' crea una bola nueva basada en el tamaño y la masa de la bola especificados al principio del script
    BallRelease.CreateSizedBallWithMass BallSize, BallMass

    ' incrementa el número de bolas en el tablero, ya que hay que contarlas
    BallsOnPlayfield = BallsOnPlayfield + 1

    ' y expulsa la bola
    PlaySoundAt SoundFXDOF("fx_Ballrel", 104, DOFPulse, DOFContactors), BallRelease
    BallRelease.Kick 90, 4
End Sub

' El jugador ha perdido su bola, y ya no hay más bolas en juego
' Empieza a contar los bonos

Sub EndOfBall()
    'debug.print "EndOfBall"

    ' La primera se ha perdido. Desde aquí ya no se puede aceptar más jugadores
    bOnTheFirstBall = False

    ' no hay cuenta de bonos al final de la bola, así que pasa a la siguente EndOfBall2
    ' si hay falta (TILT) el sistema del la falta se encargará de nuevas bolas o del fin de la partida

    If NOT Tilted Then
        vpmtimer.addtimer 400, "EndOfBall2 '"
    End If
End Sub

' Mira si el jugador ha ganado bolas extras
' y si no mira si es el último jugador o la última bola
'
Sub EndOfBall2()
    'debug.print "EndOfBall2"
    ' si hubiese falta, quítala, y pon la cuenta a cero de la falta para el próximo jugador, ó bola

    Tilted = False
    Tilt = 0
    TurnOff aTilt
    If B2SOn then
        Controller.B2SSetTilt 0
    end if
    DisableTable False 'activa de nuevo los bumpers y los slingshots

    ' ¿más tiempo?
    If ClockValue = 0 Then

        ' miramos si la puntuación clasifica como el Highscore
        CheckHighScore()
    End If

    ' ésta no es la última bola para éste jugador
    ' y si hay más de un jugador continúa con el siguente
    EndOfBallComplete()
End Sub

' pasa a la siguente bola o al siguente jugador

Sub EndOfBallComplete()
    'debug.print "EndOfBallComplete"

    ' ¿Hemos llegado al final del juego? (todas las bolas se han jugado de todos los jugadores)
    If ClockValue = 0 Then

        ' ahora se pone la mesa en el modo de final de juego
        EndOfGame()
    Else

        PlaySound "ReelInitt"

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
    end if
    ' asegúrate de que los flippers están en modo de reposo
    SolLFlipper 0
    SolRFlipper 0
    DOF 260, DOFPulse
    ' pon las luces en el modo de fin de juego
    StartAttractMode
End Sub

' check the highscore
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
    If bGameInPLay = False Then Exit Sub 'don't do anything, just delete the ball

    BallsOnPlayfield = BallsOnPlayfield - 1

    ' haz sonar el ruido de la bola
    PlaySoundAt "fx_drain", Drain
    DOF 201, DOFOff

    'si hay falta el systema de tilt se encargará de continuar con la siguente bola/jugador
    If Tilted Then
        Exit Sub
    End If

    ' si estás jugando y no hay falta
    If(bGameInPLay = True) AND(Tilted = False) Then

        ' ¿está el salva bolas activado?
        If(bBallSaverActive = True) Then

            ' ¿sí?, pues creamos una bola
            CreateNewBall()
        Else
            ' ¿es ésta la última bola en juego?
            If(BallsOnPlayfield = 0) Then
                StopClock
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
    Select Case Points
        Case 10, 100, 1000, 10000
            Score(CurrentPlayer) = Score(CurrentPlayer) + points
            UpdateScore points
            ' sounds
            If Points = 100 AND(Score(CurrentPlayer) MOD 1000) \ 100 = 0 Then  'new reel 100
                PlaySound SoundFXDOF("1000", 132, DOFPulse, DOFChimes)
            ElseIf Points = 10 AND(Score(CurrentPlayer) MOD 100) \ 10 = 0 Then 'new reel 10
                PlaySound SoundFXDOF("100", 132, DOFPulse, DOFChimes)
            Else
                PlaySound SoundFXDOF(Points, 131, DOFPulse, DOFChimes)
            End If
        Case 20, 30, 40, 50
            Add10 = Add10 + Points \ 10
            AddScore10Timer.Enabled = TRUE
        Case 200, 300, 400, 500
            Add100 = Add100 + Points \ 100
            AddScore100Timer.Enabled = TRUE
        Case 2000, 3000, 4000, 5000
            Add1000 = Add1000 + Points \ 1000
            AddScore1000Timer.Enabled = TRUE
        Case 20000, 30000, 40000, 50000
            Add10000 = Add10000 + Points \ 10000
            AddScore10000Timer.Enabled = TRUE
    End Select
    ' aquí se puede hacer un chequeo si el jugador ha ganado alguna puntuación alta y darle un crédito ó bola extra
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
End Sub

'******************************
'TIMER DE 10, 100 y 1000 PUNTOS
'******************************

' hace sonar las campanillas según los puntos

Sub AddScore10Timer_Timer()
    if Add10> 0 then
        AddScore 10
        Add10 = Add10 - 1
    Else
        Me.Enabled = FALSE
    End If
End Sub

Sub AddScore100Timer_Timer()
    if Add100> 0 then
        AddScore 100
        Add100 = Add100 - 1
    Else
        Me.Enabled = FALSE
    End If
End Sub

Sub AddScore1000Timer_Timer()
    if Add1000> 0 then
        AddScore 1000
        Add1000 = Add1000 - 1
    Else
        Me.Enabled = FALSE
    End If
End Sub

Sub AddScore10000Timer_Timer()
    if Add10000> 0 then
        AddScore 10000
        Add10000 = Add10000 - 1
    Else
        Me.Enabled = FALSE
    End If
End Sub

'***********************************************************************************
'        Score reels - puntuaciones - y actualiza otras luces del backdrop
'***********************************************************************************
'esta es al rutina que actualiza la puntuación del jugador

Sub UpdateScore(playerpoints)
    Select Case CurrentPlayer
        Case 1:ScoreReel1.Addvalue playerpoints
    End Select
    If Score(CurrentPlayer) >= 1000000 then
        TurnON amillion:TurnOFF a900
    ElseIf Score(CurrentPlayer) >= 900000 then
        TurnON a900:TurnOFF a800
    ElseIf Score(CurrentPlayer) >= 800000 then
        TurnON a800:TurnOFF a700
    ElseIf Score(CurrentPlayer) >= 700000 then
        TurnON a700:TurnOFF a600
    ElseIf Score(CurrentPlayer) >= 600000 then
        TurnON a600:TurnOFF a500
    ElseIf Score(CurrentPlayer) >= 500000 then
        TurnON a500:TurnOFF a400
    ElseIf Score(CurrentPlayer) >= 400000 then
        TurnON a400:TurnOFF a300
    ElseIf Score(CurrentPlayer) >= 300000 then
        TurnON a300:TurnOFF a200
    ElseIf Score(CurrentPlayer) >= 200000 then
        TurnON a200:TurnOFF a100
    ElseIf Score(CurrentPlayer) >= 100000 then
        TurnON a100
    End If
    If B2SOn then
        Controller.B2SSetScorePlayer CurrentPlayer, Score(CurrentPlayer)
        If Score(CurrentPlayer) >= 1000000 then
            Controller.B2SSetData 89, 1:Controller.B2SSetData 88, 0
        ElseIf Score(CurrentPlayer) >= 900000 then
            Controller.B2SSetData 88, 1:Controller.B2SSetData 87, 0
        ElseIf Score(CurrentPlayer) >= 800000 then
            Controller.B2SSetData 87, 1:Controller.B2SSetData 86, 0
        ElseIf Score(CurrentPlayer) >= 700000 then
            Controller.B2SSetData 86, 1:Controller.B2SSetData 85, 0
        ElseIf Score(CurrentPlayer) >= 600000 then
            Controller.B2SSetData 85, 1:Controller.B2SSetData 84, 0
        ElseIf Score(CurrentPlayer) >= 500000 then
            Controller.B2SSetData 84, 1:Controller.B2SSetData 83, 0
        ElseIf Score(CurrentPlayer) >= 400000 then
            Controller.B2SSetData 83, 1:Controller.B2SSetData 82, 0
        ElseIf Score(CurrentPlayer) >= 300000 then
            Controller.B2SSetData 82, 1:Controller.B2SSetData 81, 0
        ElseIf Score(CurrentPlayer) >= 200000 then
            Controller.B2SSetData 81, 1:Controller.B2SSetData 80, 0
        ElseIf Score(CurrentPlayer) >= 100000 then
            Controller.B2SSetData 80, 1
        End If
    end if
End Sub

' pone todos los marcadores a 0
Sub ResetScores
    ScoreReel1.ResetToZero
    TurnOFF a100
    TurnOFF a200
    TurnOFF a300
    TurnOFF a400
    TurnOFF a500
    TurnOFF a600
    TurnOFF a700
    TurnOFF a800
    TurnOFF a900
    TurnOFF amillion
    If B2SOn then
        Controller.B2SSetScorePlayer1 0
        Controller.B2SSetData 80, 0
        Controller.B2SSetData 81, 0
        Controller.B2SSetData 82, 0
        Controller.B2SSetData 83, 0
        Controller.B2SSetData 84, 0
        Controller.B2SSetData 85, 0
        Controller.B2SSetData 86, 0
        Controller.B2SSetData 87, 0
        Controller.B2SSetData 88, 0
        Controller.B2SSetData 89, 0
    end if
End Sub

Sub AddCredits(value)
    If Credits < 9 Then
        Credits = Credits + value
        UpdateCredits
    end if
End Sub

Sub UpdateCredits
    If Credits> 0 Then 'this is for Bally tables
        CreditsLight.State = 1
    Else
        DOF 200, DOFOff
        CreditsLight.State = 0
    End If
    CreditReel.SetValue credits
    If B2SOn then
        Controller.B2SSetCredits Credits
    end if
End Sub

'*************************
' Partidas y bolas extras
'*************************

Sub AwardSpecial()
    PlaySound SoundFXDOF("fx_knocker", 300, DOFPulse, DOFKnocker)
    DOF 230, DOFPulse
    DOF 200, DOFOn
    If ReplayMode Then
        AddCredits 1
    Else
        AddClock 5
    End If
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

    ' enciente la luz de fin de partida
    TurnOn aGameOver
End Sub

Sub StopAttractMode()
    bAttractMode = False
    ResetScores
    For each x in aLights
        x.State = 0
    Next

    ' apaga la luz de fin de partida
    TurnOff aGameOver
End Sub

'*********************************
'    Load / Save / Highscore
'*********************************

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
    If Not FileObj.FolderExists(UserDirectory) then
        Credits = 0
        Exit Sub
    End If
    If Not FileObj.FileExists(UserDirectory & TableName& ".txt") then
        Credits = 0
        Exit Sub
    End If
    Set ScoreFile = FileObj.GetFile(UserDirectory & TableName& ".txt")
    Set TextStr = ScoreFile.OpenAsTextStream(1, 0)
    If(TextStr.AtEndOfStream = True) then
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
    Credits = cdbl(temp2)

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
    If Not FileObj.FolderExists(UserDirectory) then
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
            If HSScore(j) < HSScore(j + 1) Then
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

' default high scores
HSScore(1) = 169000
HSScore(2) = 159000
HSScore(3) = 149000
HSScore(4) = 139000
HSScore(5) = 129000

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
        If len(InitialString) < 3 then
            SetHSLine 3, InitialString & SelectedChar
        End If
    End If
    If len(InitialString) = 3 then
        ' save the score
        For i = 5 to 1 step -1
            If i = 1 or(HSNewHigh> HSScore(i) and HSNewHigh <= HSScore(i - 1) ) then
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

'****************************************
' Actualizaciones en tiempo real
'****************************************

Sub LeftFlipper_Animate:LeftFlipperTop.RotZ = LeftFlipper.CurrentAngle:End Sub
Sub RightFlipper_Animate:RightFlipperTop.RotZ = RightFlipper.CurrentAngle:End Sub
Sub diverter_animate:diverterp.rotz = diverter.CurrentAngle:End Sub

'***********************************************************************
' *********************************************************************
'            Aquí empieza el código particular a la mesa
' (hasta ahora todas las rutinas han sido muy generales para todas las mesas)
' (y hay muy pocas rutinas que necesitan cambiar de mesa a mesa)
' *********************************************************************
'***********************************************************************

' se inicia las dianas abatibles, primitivas, etc.
' aunque en el VPX no hay muchos objetos que necesitan ser iniciados

Sub VPObjects_Init
End Sub

' variables de la mesa
Dim ClockValue  'each number is 5 seconds, value from 0 to 48
Dim AddC

Sub Game_Init() 'esta rutina se llama al principio de un nuevo juego

    'Empezar alguna música, si hubiera música en esta mesa

    'Iniciar algunas luces, en esta mesa son las mismas luces que las de una bola nueva
    TurnOffPlayfieldLights()

    'iniciar variables, en esta mesa hay muy pocas variables ya que usamos las luces, y el UserValue de las dianas
    ClockValue = 0
    AddC = 0
    AddClock StartupTime \ 5

    'iniciar algún timer
    StopClock
End Sub

Sub ResetNewBallVariables() 'inicia las variable para una bola ó jugador nuevo
End Sub

Sub ResetNewBallLights()    'enciende ó apaga las luces para una bola nueva
'TurnOffPlayfieldLights()
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

'**************************
'  Gomas de 10 y 100 puntos
'**************************

Sub rsband012_Hit
    If NOT Tilted Then
        AddScore 10
    End If
End Sub

Sub rlband001_Hit
    If NOT Tilted Then
        AddScore 10
    End If
End Sub

Sub rlband007_Hit
    If NOT Tilted Then
        AddScore 10
    End If
End Sub

Sub rlband003_Hit
    If NOT Tilted Then
        AddScore 10
    End If
End Sub

Sub rsband010_Hit
    If NOT Tilted Then
        AddScore 10
    End If
End Sub

Sub rsband009_Hit
    If NOT Tilted Then
        AddScore 10
    End If
End Sub

Sub rlband002_Hit
    If NOT Tilted Then
        AddScore 10
    End If
End Sub

'*********
' Bumpers
'*********

Sub Bumper001_Hit 'left
    If NOT Tilted Then
        PlaySoundAt SoundFXDOF("fx_Bumper", 105, DOFPulse, DOFContactors), bumper001
        DOF 212, DOFPulse
        ' añade algunos puntos
        AddScore 100
    End If
End Sub

Sub Bumper002_Hit 'center
    If NOT Tilted Then
        PlaySoundAt SoundFXDOF("fx_Bumper", 107, DOFPulse, DOFContactors), bumper002
        DOF 213, DOFPulse
        ' añade algunos puntos
        AddScore 1000
    End If
End Sub

Sub Bumper003_Hit 'right
    If NOT Tilted Then
        PlaySoundAt SoundFXDOF("fx_Bumper", 108, DOFPulse, DOFContactors), bumper003
        DOF 213, DOFPulse
        AddScore 100
    End If
End Sub

'************************
'     Dianas fijas
'************************

'Travel - Summer
Sub Target005_hit
    PlaySoundAtBall SoundFXDOF("fx_target", 109, DOFPulse, DOFTargets)
    If Tilted Then Exit Sub
    Addscore 100
    Light027.State = 1
    CheckTravel
End Sub

Sub Target006_hit
    PlaySoundAtBall SoundFXDOF("fx_target", 109, DOFPulse, DOFTargets)
    If Tilted Then Exit Sub
    Addscore 100
    Light028.State = 1
    CheckTravel
End Sub

Sub Target007_hit
    PlaySoundAtBall SoundFXDOF("fx_target", 109, DOFPulse, DOFTargets)
    If Tilted Then Exit Sub
    Addscore 100
    Light029.State = 1
    CheckTravel
End Sub

Sub Target008_hit
    PlaySoundAtBall SoundFXDOF("fx_target", 109, DOFPulse, DOFTargets)
    If Tilted Then Exit Sub
    Addscore 100
    Light031.State = 1
    CheckTravel
End Sub

Sub Target009_hit
    PlaySoundAtBall SoundFXDOF("fx_target", 109, DOFPulse, DOFTargets)
    If Tilted Then Exit Sub
    Addscore 100
    Light030.State = 1
    CheckTravel
End Sub

Sub Target010_hit
    PlaySoundAtBall SoundFXDOF("fx_target", 109, DOFPulse, DOFTargets)
    If Tilted Then Exit Sub
    Addscore 100
    Light032.State = 1
    CheckTravel
End Sub

Sub CheckTravel
    If Light027.State + Light028.State + Light029.State + Light030.State + Light031.State + Light032.State = 6 Then
        Light004.State = 1
    End If
End Sub

' Time
Sub Target001_hit
    PlaySoundAtBall SoundFXDOF("fx_target", 109, DOFPulse, DOFTargets)
    If Tilted Then Exit Sub
    Light012.State = 1
    Addscore 500
    CheckTime
End Sub

Sub Target002_hit
    PlaySoundAtBall SoundFXDOF("fx_target", 109, DOFPulse, DOFTargets)
    If Tilted Then Exit Sub
    Light013.State = 1
    Addscore 500
    CheckTime
End Sub

Sub Target003_hit
    PlaySoundAtBall SoundFXDOF("fx_target", 109, DOFPulse, DOFTargets)
    If Tilted Then Exit Sub
    Light014.State = 1
    Addscore 500
    CheckTime
End Sub

Sub Target004_hit
    PlaySoundAtBall SoundFXDOF("fx_target", 109, DOFPulse, DOFTargets)
    If Tilted Then Exit Sub
    Light015.State = 1
    Addscore 500
    CheckTime
End Sub

Sub CheckTime
    If Light012.State + Light013.State + Light014.State + Light015.State = 4 Then
        Light003.State = 1
        Light007.State = 1
    End If
End Sub

'******************************
'     Pasillos superiores
'******************************

Sub Trigger008_Hit
    PlaySoundAt "fx_sensor", Trigger008
    If Tilted Then Exit Sub
    AddScore 1000
    StartClock
End Sub

Sub Trigger009_Hit 'center
    PlaySoundAt "fx_sensor", Trigger009
    If Tilted Then Exit Sub
    AddScore 20000
    AddClock 2
End Sub

Sub Trigger010_Hit
    PlaySoundAt "fx_sensor", Trigger010
    If Tilted Then Exit Sub
    AddScore 1000
    StartClock
End Sub

'******************************
'     Pasillos inferiores
'******************************

Sub Trigger003_Hit 'left outlane
    PlaySoundAt "fx_sensor", Trigger001
    If Tilted Then Exit Sub
    AddScore 10000
    StopClock
End Sub

Sub Trigger002_Hit 'right outlane
    PlaySoundAt "fx_sensor", Trigger001
    If Tilted Then Exit Sub
    AddScore 10000
    StopClock
End Sub

'******************************
'          Buttons
'******************************

Sub Trigger004_Hit
    PlaySoundAt "fx_sensor", Trigger004
    If Tilted Then Exit Sub
    AddScore 50
End Sub

Sub Trigger005_Hit
    PlaySoundAt "fx_sensor", Trigger005
    If Tilted Then Exit Sub
    AddScore 50
End Sub

Sub Trigger006_Hit
    PlaySoundAt "fx_sensor", Trigger006
    If Tilted Then Exit Sub
    StartClock
End Sub

Sub Trigger007_Hit
    PlaySoundAt "fx_sensor", Trigger007
    If Tilted Then Exit Sub
    StartClock
End Sub

Sub Trigger012_Hit
    If Tilted Then Exit Sub
    StartClock
End Sub

Sub Trigger013_Hit
    If Tilted Then Exit Sub
    StartClock
End Sub

'********************************
'         Kickers
'********************************

Sub Kicker001_Hit 'lower
    PlaySoundAt "fx_kicker_enter", Kicker001
    If Tilted Then vpmTimer.AddTimer 500, "PlaySoundAt""fx_kicker"",Kicker001:Kicker001.kick 0, 25 '"
    StartClock
    Dim tmp:tmp = 500
    Addscore 500
    If Light004.State Then 'travel Lights
        AddClock 5
        tmp = tmp + 500
        Light004.State = 0
        Light027.State = 0
        Light028.State = 0
        Light029.State = 0
        Light030.State = 0
        Light031.State = 0
        Light032.State = 0
    End If
    If Light003.State Then 'time Lights
        AddClock 5
        tmp = tmp + 500
        Light003.State = 0
        Light012.State = 0
        Light013.State = 0
        Light014.State = 0
        Light015.State = 0
    End If
    vpmTimer.AddTimer tmp, "PlaySoundAt""fx_kicker"",Kicker001:Kicker001.kick 0, 21 +RndNbr(8) '"
End Sub

Sub Kicker002_Hit 'center
    PlaySoundAt "fx_kicker_enter", Kicker002
    If Tilted Then vpmTimer.AddTimer 500, "PlaySoundAt""fx_kicker"",Kicker002:Kicker002.kick 166, 20 '"
    StopClock
    If Light007.State Then
        Light007.State = 0
        AddScore 50000
    Else
        Addscore 500
    End If
    vpmTimer.AddTimer 500, "PlaySoundAt""fx_kicker"",Kicker002:Kicker002.kick 166, 20 '"
End Sub

'********************************
'        Clock - timer
'********************************
Sub AddClock(value)
    AddC = AddC + value
    AddClockTimer.Enabled = 1
End Sub

Sub UpdateClock
    If ClockValue < 0 Then ClockValue = 0
    Select case ClockValue
        Case 0, 12, 24:l0.State = 1:l5.State = 0:l10.State = 0:l15.State = 0:l20.State = 0:l25.State = 0:l30.State = 0:l35.State = 0:l40.State = 0:l45.State = 0:l50.State = 0:l55.State = 0
        Case 1, 13, 25:l0.State = 0:l5.State = 1:l10.State = 0:l15.State = 0:l20.State = 0:l25.State = 0:l30.State = 0:l35.State = 0:l40.State = 0:l45.State = 0:l50.State = 0:l55.State = 0
        Case 2, 14, 26:l0.State = 0:l5.State = 0:l10.State = 1:l15.State = 0:l20.State = 0:l25.State = 0:l30.State = 0:l35.State = 0:l40.State = 0:l45.State = 0:l50.State = 0:l55.State = 0
        Case 3, 15, 27:l0.State = 0:l5.State = 0:l10.State = 0:l15.State = 1:l20.State = 0:l25.State = 0:l30.State = 0:l35.State = 0:l40.State = 0:l45.State = 0:l50.State = 0:l55.State = 0
        Case 4, 16, 28:l0.State = 0:l5.State = 0:l10.State = 0:l15.State = 0:l20.State = 1:l25.State = 0:l30.State = 0:l35.State = 0:l40.State = 0:l45.State = 0:l50.State = 0:l55.State = 0
        Case 5, 17, 29:l0.State = 0:l5.State = 0:l10.State = 0:l15.State = 0:l20.State = 0:l25.State = 1:l30.State = 0:l35.State = 0:l40.State = 0:l45.State = 0:l50.State = 0:l55.State = 0
        Case 6, 18, 30:l0.State = 0:l5.State = 0:l10.State = 0:l15.State = 0:l20.State = 0:l25.State = 0:l30.State = 1:l35.State = 0:l40.State = 0:l45.State = 0:l50.State = 0:l55.State = 0
        Case 7, 19, 31:l0.State = 0:l5.State = 0:l10.State = 0:l15.State = 0:l20.State = 0:l25.State = 0:l30.State = 0:l35.State = 1:l40.State = 0:l45.State = 0:l50.State = 0:l55.State = 0
        Case 8, 20, 32:l0.State = 0:l5.State = 0:l10.State = 0:l15.State = 0:l20.State = 0:l25.State = 0:l30.State = 0:l35.State = 0:l40.State = 1:l45.State = 0:l50.State = 0:l55.State = 0
        Case 9, 21, 33:l0.State = 0:l5.State = 0:l10.State = 0:l15.State = 0:l20.State = 0:l25.State = 0:l30.State = 0:l35.State = 0:l40.State = 0:l45.State = 1:l50.State = 0:l55.State = 0
        Case 10, 22, 34:l0.State = 0:l5.State = 0:l10.State = 0:l15.State = 0:l20.State = 0:l25.State = 0:l30.State = 0:l35.State = 0:l40.State = 0:l45.State = 0:l50.State = 1:l55.State = 0
        Case 11, 23, 35:l0.State = 0:l5.State = 0:l10.State = 0:l15.State = 0:l20.State = 0:l25.State = 0:l30.State = 0:l35.State = 0:l40.State = 0:l45.State = 0:l50.State = 0:l55.State = 1
    End Select
    If ClockValue < 12 Then
        l60.State = 0
        l120.State = 0
    ElseIf ClockValue < 24 Then
        l60.State = 1
        l120.State = 0
    Else 'it is 24 or more
        l60.State = 0
        l120.State = 1
    End If
End Sub

Sub StartClock
    ReduceClockTimer.Enabled = 1
    Light005.State = 0
    Light006.State = 1
End Sub

Sub StopClock
    ReduceClockTimer.Enabled = 0
    Light005.State = 1
    Light006.State = 0
End Sub

Sub AddClockTimer_Timer()
    if AddC> 0 AND ClockValue < 35 then
        PlaySound "fx_relay"
        ClockValue = ClockValue + 1
        UpdateClock
        AddC = AddC - 1
    Else
        AddC = 0
        Me.Enabled = FALSE
    End If
End Sub

Sub ReduceClockTimer_Timer
    If ClockValue> 0 Then
        PlaySound "fx_relay"
        ClockValue = ClockValue - 1
        UpdateClock
    End If
End Sub


'*********************************
' Table Options F12 User Options
'*********************************
' Table1.Option arguments are:
' - option name, minimum value, maximum value, step between valid values, default value, unit (0=None, 1=Percent), an optional array of literal strings

Dim LUTImage, StartupTime, ReplayMode, Difficulty

Sub Table1_OptionEvent(ByVal eventId)
    Dim x, y

    'LUT
    LutImage = Table1.Option("Select LUT", 0, 21, 1, 0, 0, Array("Normal 0", "Normal 1", "Normal 2", "Normal 3", "Normal 4", "Normal 5", "Normal 6", "Normal 7", "Normal 8", "Normal 9", "Normal 10", _
        "Warm 0", "Warm 1", "Warm 2", "Warm 3", "Warm 4", "Warm 5", "Warm 6", "Warm 7", "Warm 8", "Warm 9", "Warm 10") )
    UpdateLUT

    ' Cabinet rails
    x = Table1.Option("Cabinet Rails", 0, 1, 1, 1, 0, Array("Hide", "Show") )
    For each y in aRails:y.visible = x:next

    ' Startup Time
    x = Table1.Option("Startup time", 0, 1, 1, 0, 0, Array("45", "75") )
    If x = 1 Then StartupTime = 45 Else StartupTime = 75

    ' Game mode: replay or time
    x = Table1.Option("Game Mode", 0, 1, 1, 0, 0, Array("Replay", "Time") )
    If x = 1 Then ReplayMode = True Else ReplayMode = False

    ' Difficulty
    Difficulty = Table1.Option("Difficulty", 1, 3, 1, 2, 0, Array("Hard", "Normal", "Easy") )
    ReduceClockTimer.Interval = Difficulty * 1000
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

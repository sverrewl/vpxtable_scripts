' ****************************************************************
'               VISUAL PINBALL X EM Script por JPSalas
'         Script Básico para juegos EM Script hasta 4 players
'        usa el core.vbs para funciones extras
' ****************************************************************

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
Const TableName = "CheckMate"
Const cGameName = "CheckMate"
Const MaxPlayers = 4    ' de 1 a 4
Const MaxMultiplier = 5 ' limita el bonus multiplicador a 5
Const BallsPerGame = 5  ' normalmente 3 ó 5
Const Special1 = 560000 ' puntuación a obtener para partida extra
Const Special2 = 720000 ' puntuación a obtener para partida extra
Const Special3 = 999000 ' puntuación a obtener para partida extra

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

    ' Carga los valores grabados highscore y créditos
    Loadhs
    'HighscoreReel.SetValue Highscore
    ' también pon el highscore en el reel del primer jugador
    ' ScoreReel1.SetValue Highscore
    ' en esta mesa el attractmode se encarga de mostrar el highscore

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

'******
' Keys
'******

Sub Table1_KeyDown(ByVal Keycode)
    ' añade monedas
    If Keycode = AddCreditKey Then
        If(Tilted = False)Then
            AddCredits 1
            PlaySound "fx_coin"
            PlaySound "coin"
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

        If keycode = MechanicalTilt Then Nudge 0, 4:PlaySound "fx_nudge",0,1,1,0,25:CheckTilt
        ' teclas de los flipers
        If keycode = LeftFlipperKey Then SolLFlipper 1
        If keycode = RightFlipperKey Then SolRFlipper 1

        ' tecla de empezar el juego
        If keycode = StartGameKey Then
            If((PlayersPlayingGame < MaxPlayers)AND(bOnTheFirstBall = True))Then

                If(bFreePlay = True)Then
                    PlayersPlayingGame = PlayersPlayingGame + 1
                    PlaySound "start"
                Else
                    If(Credits > 0)then
                        PlayersPlayingGame = PlayersPlayingGame + 1
                        Credits = Credits - 1
                        UpdateCredits
                        UpdateBallInPlay
                        PlaySound "start"
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
                        PlaySound "start"
                    End If
                Else
                    If(Credits > 0)Then
                        If(BallsOnPlayfield = 0)Then
                            Credits = Credits - 1
                            UpdateCredits
                            ResetScores
                            ResetForNewGame()
                            PlaySound "start"
                        End If
                    Else
                        ' Not Enough Credits to start a game.
                        PlaySound "10000"
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
        PlaySoundAt "fx_flipperup", LeftFlipper
        LeftFlipper.RotateToEnd
    Else
        PlaySoundAt "fx_flipperdown", LeftFlipper
        LeftFlipper.RotateToStart
    End If
End Sub

Sub SolRFlipper(Enabled)
    If Enabled Then
        PlaySoundAt "fx_flipperup", RightFlipper
        RightFlipper.RotateToEnd
    Else
        PlaySoundAt "fx_flipperdown", RightFlipper
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
  If B2SOn Then Controller.B2SSetData 60,1
End Sub

Sub GiOff 'apaga las luces GI
    Dim bulb
    For each bulb in aGiLights
        bulb.State = 0
    Next
  If B2SOn Then Controller.B2SSetData 60,0
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
        Bumper1.Force = 0
        Bumper2.Force = 0
        LeftSlingshot.Disabled = 1
        RightSlingshot.Disabled = 1
    Else
        'enciende de nuevo todas las luces GI
        GiOn

        Bumper1.Force = 10
        Bumper2.Force = 10
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
' de lo contrario la rutina mirará si todavía hay bolas en la mesa
End Sub

' *********************************************************************
'               Funciones para los sonidos de la mesa
' *********************************************************************

Function Vol(ball) ' Calcula el volumen del sonido basado en la velocidad de la bola
    Vol = Csng(BallVel(ball) ^2 / 2000)
End Function

Function Pan(ball) ' Calculala posición estéreo de la bola (izquierda a derecha)
    Dim tmp
    tmp = ball.x * 2 / table1.width-1
    If tmp > 0 Then
        Pan = Csng(tmp ^10)
    Else
        Pan = Csng(-((- tmp) ^10))
    End If
End Function

Function Pitch(ball) ' Calcula el tono según la velocidad de la bola
    Pitch = BallVel(ball) * 20
End Function

Function BallVel(ball) 'Calcula la velocidad de la bola
    BallVel = INT(SQR((ball.VelX ^2) + (ball.VelY ^2)))
End Function

Function AudioFade(ball) 'Solo para VPX 10.4 y siguentes: calcula la posición de arriba/abajo de la bola, para mesas con el sonido dolby
    Dim tmp
    tmp = ball.y * 2 / Table1.height-1
    If tmp > 0 Then
        AudioFade = Csng(tmp ^10)
    Else
        AudioFade = Csng(-((- tmp) ^10))
    End If
End Function

Sub PlaySoundAt(soundname, tableobj) ' Hace sonar un sonido en la posición de un objeto, como bumpers y flippers
    PlaySound soundname, 0, 1, Pan(tableobj), 0, 0, 0, 0, AudioFade(tableobj)
End Sub

Sub PlaySoundAtBall(soundname) ' hace sonar un sonido en la posición de la bola
    PlaySound soundname, 0, Vol(ActiveBall), pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0, AudioFade(ActiveBall)
End Sub

'*****************************************
'      Los sonidos de la bola/s rodando
'*****************************************

Const tnob = 6 ' número total de bola
Const lob = 0  'número de bola encerradas
ReDim rolling(tnob)
InitRolling

Sub InitRolling
    Dim i
    For i = 0 to tnob
        rolling(i) = False
    Next
End Sub

Sub RollingUpdate()
    Dim BOT, b, ballpitch
    BOT = GetBalls

    ' para el sonido de bolas perdidas
    For b = UBound(BOT) + 1 to tnob
        rolling(b) = False
        StopSound("fx_ballrolling" & b)
    Next

    ' sale de la rutina si no hay más bolas en la mesa
    If UBound(BOT) = lob - 1 Then Exit Sub

    ' hace sonar el sonido de la bola rodando para cada bola
    For b = lob to UBound(BOT)
        If BallVel(BOT(b)) > 1 Then
            If BOT(b).z < 30 Then
                ballpitch = Pitch(BOT(b))
            Else
                ballpitch = Pitch(BOT(b)) + 20000 'aumenta el tono del sonido si la bola está sobre una rampa
            End If
            rolling(b) = True
            PlaySound("fx_ballrolling" & b), -1, Vol(BOT(b)), Pan(BOT(b)), 0, ballpitch, 1, 0, AudioFade(BOT(b))
        Else
            If rolling(b) = True Then
                StopSound("fx_ballrolling" & b)
                rolling(b) = False
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

'***************************************
' Sonidos de las colecciones de objetos
' como metales, gomas, plásticos, etc
'***************************************

Sub aMetals_Hit(idx):PlaySoundAtBall "fx_MetalHit":End Sub
Sub aRubber_Bands_Hit(idx):PlaySoundAtBall "fx_rubber_band":End Sub
Sub aRubber_Posts_Hit(idx):PlaySoundAtBall "fx_rubber_post":End Sub
Sub aRubber_Pins_Hit(idx):PlaySoundAtBall "fx_rubber_pin":End Sub
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
    If B2SOn then
        Controller.B2SSetGameOver 0
    end if
    ' enciende las luces GI si estuvieran apagadas
    GiOn

  MatchWheel.enabled=true

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

    Match = INT(RND(1) * 8) * 10 + 10 'empieza con un número aleatorio de 10 a 90 para la loteria
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
    BallRelease.CreateSizedBallWithMass BallSize, BallMass

    ' incrementa el número de bolas en el tablero, ya que hay que contarlas
    BallsOnPlayfield = BallsOnPlayfield + 1

    ' actualiza las luces del backdrop
    UpdateBallInPlay

    ' y expulsa la bola
    PlaySoundAt "fx_Ballrel", BallRelease
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
        If DoubleBonus = 2 Then
            BonusCountTimer.Interval = 400
            CreditLight.BlinkInterval = 200
        Else
            BonusCountTimer.Interval = 250
            CreditLight.BlinkInterval = 125
        End If
        CreditLight.State = 2
        BonusCountTimer.Enabled = 1
    Else 'Si hay falta simplemente espera un momento y va directo a la segunta parte después de perder la bola
        vpmtimer.addtimer 400, "EndOfBall2 '"
    End If
End Sub

Sub BonusCountTimer_Timer 'añade los bonos y actualiza las luces
    'debug.print "BonusCount_Timer"
    If Bonus > 0 Then
        Bonus = Bonus -1
        AddScore 1000 * DoubleBonus
        UpdateBonusLights
    Else
        ' termina la cuenta de los bonos y continúa con el fin de bola
        BonusCountTimer.Enabled = 0
        CreditLight.State = 0
        vpmtimer.addtimer 1000, "EndOfBall2 '"
    End If
End Sub

Sub UpdateBonusLights
    For each x in aBonusLights
        If x.State = 1 Then
            x.State = 0
            Exit Sub
        End If
    Next
    If BankLDown Then
        BankLDown = BankLDown - 1
        For each x in aLeftDropLights
            x.State = 1
        Next
        aLeftDropLights(0).State = 0
        Exit Sub
    End If
    If BankRDown Then
        BankRDown = BankRDown -1
        For each x in aRightDropLights
            x.State = 1
        Next
        aRightDropLights(0).State = 0
    End If
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
  MatchWheel.enabled=false
    bGameInPLay = False
    bJustStarted = False
    If B2SOn then
        Controller.B2SSetGameOver 1
        Controller.B2SSetBallInPlay 0
        Controller.B2SSetPlayerUp 0
        Controller.B2SSetCanPlay 0
    end if
    ' asegúrate de que los flippers están en modo de reposo
    SolLFlipper 0
    SolRFlipper 0

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

' Esta función calcula el Highscore y te da una partida gratis si has conseguido el Highscore
Sub CheckHighscore
    For x = 1 to PlayersPlayingGame
        If Score(x) > Highscore Then
            Highscore = Score(x)
            PlaySound "fx_knocker"
            'HighscoreReel.SetValue Highscore
            AddCredits 1
        End If
    Next
End Sub

'******************
'  Match - Loteria
'******************

sub MatchWheel_timer
  IncreaseMatch
  If B2SOn then
        If Match = 0 then
            Controller.B2SSetMatch 100
        else
            Controller.B2SSetMatch Match
    end if
  end if
end sub

Sub IncreaseMatch
    Match = (Match + 10)MOD 100
End Sub

Sub Verification_Match()
    PlaySound "fx_match"
    Display_Match
    If(Score(CurrentPlayer)MOD 100) = Match Then
        PlaySound "fx_knocker"
        AddCredits 1
    End If
End Sub

Sub Clear_Match()
    For each x in aMatchLights
        x.State = 2
    Next
    If B2SOn then
        Controller.B2SSetMatch 0
    end if
End Sub

Sub Display_Match()
    For each x in aMatchLights
        x.State = 0
    Next
    aMatchLights(Match \ 10).state = 1
'    If B2SOn then
'        If Match = 0 then
'            Controller.B2SSetMatch 100
'        else
'            Controller.B2SSetMatch Match
'        end if
'    end if
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
    'si hay falta ve directo al final de la bola
    If Tilted Then
        vpmtimer.addtimer 500, "EndOfBall '" 'hacemos una pequeña pausa anter de continuar con el fin de bola
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
    Select Case Points
        Case 10, 100, 1000
            ' añade los puntos a la variable del actual jugador
            Score(CurrentPlayer) = Score(CurrentPlayer) + points
            ' actualiza los contadores
            UpdateScore points
            ' hace sonar las campanillas de acuerdo a los puntos obtenidos
            If Points = 100 AND(Score(CurrentPlayer)MOD 1000) \ 100 = 0 Then  'nuevo reel de 1000
                PlaySound "1000"
            ElseIf Points = 10 AND(Score(CurrentPlayer)MOD 100) \ 10 = 0 Then 'nuevo reel de 100
                PlaySound "100"
            Else
                PlaySound "10"
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
End Sub

'******************************
'TIMER DE 10, 100 y 1000 PUNTOS
'******************************

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

Sub UpdateScore(playerpoints)
    Select Case CurrentPlayer
        Case 1:ScoreReel1.Addvalue playerpoints
        Case 2:ScoreReel2.Addvalue playerpoints
        Case 3:ScoreReel3.Addvalue playerpoints
        Case 4:ScoreReel4.Addvalue playerpoints
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
    ScoreReel2.ResetToZero
    ScoreReel3.ResetToZero
    ScoreReel4.ResetToZero
    If B2SOn then
        Controller.B2SSetScorePlayer1 0
        Controller.B2SSetScorePlayer2 0
        Controller.B2SSetScorePlayer3 0
        Controller.B2SSetScorePlayer4 0
        Controller.B2SSetScoreRolloverPlayer1 0
        Controller.B2SSetScoreRolloverPlayer2 0
        Controller.B2SSetScoreRolloverPlayer3 0
        Controller.B2SSetScoreRolloverPlayer4 0
    end if
End Sub

Sub AddCredits(value)
    If Credits < 9 Then
        Credits = Credits + value
        CreditReel.SetValue Credits
        UpdateCredits
    end if
End Sub

Sub UpdateCredits
    If Credits > 0 Then 'this is for Bally tables
    'CreditLight.State = 1
    Else
    'CreditLight.State = 0
    End If
    CreditReel.SetValue credits
    If B2SOn then
        Controller.B2SSetData 29, Credits
    end if
End Sub

Sub UpdateBallInPlay 'actualiza los marcadores de las bolas, el número de jugador y el número total de jugadores
    BallinPlayReel.SetValue Balls
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
    Select Case PlayersPlayingGame
        Case 1:cp1.State = 1:cp2.State = 0:cp3.State = 0:cp4.State = 0
        Case 2:cp1.State = 0:cp2.State = 1:cp3.State = 0:cp4.State = 0
        Case 3:cp1.State = 0:cp2.State = 0:cp3.State = 1:cp4.State = 0
        Case 4:cp1.State = 0:cp2.State = 0:cp3.State = 0:cp4.State = 1
    End Select
    If B2SOn then
        Controller.B2SSetCanPlay PlayersPlayingGame
        Controller.B2SSetCanPlay PlayersPlayingGame
    end if
End Sub

'*************************
' Partidas y bolas extras
'*************************

Sub AwardExtraBall()
    If NOT bExtraBallWonThisBall Then
        PlaySound "fx_knocker"
        ExtraBallsAwards(CurrentPlayer) = ExtraBallsAwards(CurrentPlayer) + 1
        bExtraBallWonThisBall = True
        LightShootAgain.State = 1
        If B2SOn then
            Controller.B2SSetShootAgain 1
        end if
    END If
End Sub

Sub AwardSpecial()
    PlaySound "fx_knocker"
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

    ' enciente la luz de fin de partida
    GameOverR.SetValue 1
    ' empieza un timer
    AttractTimer.Enabled = 1
End Sub

Sub StopAttractMode()
    bAttractMode = False
    ResetScores
    For each x in aLights
        x.State = 0
    Next

    ' apaga la luz de fin de partida
    GameOverR.SetValue 0
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
            ScoreReel2.SetValue Score(2)
            ScoreReel3.SetValue Score(3)
            ScoreReel4.SetValue Score(4)
            If B2SOn then
                Controller.B2SSetScorePlayer1 Score(1)
                Controller.B2SSetScorePlayer2 Score(2)
                Controller.B2SSetScorePlayer3 Score(3)
                Controller.B2SSetScorePlayer4 Score(4)
            end if
        Case 1
            AttractStep = 0
            ScoreReel1.SetValue Highscore
            ScoreReel2.SetValue Highscore
            ScoreReel3.SetValue Highscore
            ScoreReel4.SetValue Highscore
            If B2SOn then
                Controller.B2SSetScorePlayer1 Highscore
                Controller.B2SSetScorePlayer2 Highscore
                Controller.B2SSetScorePlayer3 Highscore
                Controller.B2SSetScorePlayer4 Highscore
            end if
    End Select
End Sub

'************************************************
'    Load (cargar) / Save (guardar)/ Highscore
'************************************************

' solamente guardamos el número de créditos y la puntuación más alta

Sub Loadhs
    x = LoadValue(TableName, "HighScore")
    If(x <> "")Then HighScore = CDbl(x)Else HighScore = 0 End If
    x = LoadValue(TableName, "Credits")
    If(x <> "")then Credits = CInt(x)Else Credits = 0 End If
End Sub

Sub Savehs
    SaveValue TableName, "HighScore", HighScore
    SaveValue TableName, "Credits", Credits
End Sub

' por si se necesitara quitar la actual puntuación más alta, se le puede poner a una tecla,
' o simplemente abres la ventana de debug y escribes Reseths y le das al enter
Sub Reseths
    HighScore = 0
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
    LeftFlipperTop.RotZ = LeftFlipper.CurrentAngle
    RightFlipperTop.RotZ = RightFlipper.CurrentAngle
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
Dim BankLDown   'esta variable contendrá el número de veces que el jugador ha abatido las dianas
Dim BankRDown   'esta variable contendrá el número de veces que el jugador ha abatido las dianas
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
    DoubleBonus = 1                        'multiplicador de los bonos
    BankLDown = 0                          'a cada nueva bola pone a cero el número de veces que las dianas han sido abatidas
    BankRDown = 0                          'a cada nueva bola pone a cero el número de veces que las dianas han sido abatidas
    SpecialReady = 0
    vpmtimer.addtimer 100, "BankLDropup '" 'sube las dianas
    vpmtimer.addtimer 300, "BankRDropup '" 'sube las dianas
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
Dim LStep, RStep

Sub LeftSlingShot_Slingshot
    If Tilted Then Exit Sub
    PlaySoundAtBall "fx_slingshot"
    IncreaseMatch 'incrementa el número de la loteria para simular un número aleatorio
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
        Case 3:LeftSLing2.Visible = 0:Lemk.RotX = -10:LeftSlingShot.TimerEnabled = 0
    End Select
    LStep = LStep + 1
End Sub

Sub RightSlingShot_Slingshot
    If Tilted Then Exit Sub
    PlaySoundAtBall "fx_slingshot"
    IncreaseMatch 'incrementa el número de la loteria para simular un número aleatorio
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
        Case 3:RightSLing2.Visible = 0:Remk.RotX = -10:RightSlingShot.TimerEnabled = 0
    End Select
    RStep = RStep + 1
End Sub

'**************************
'  Gomas de 10 y 100 puntos
'**************************

Sub RubberBand8_Hit
    If NOT Tilted Then
        IncreaseMatch
        ' añade 100 puntos
        AddScore 100
    End If
End Sub

Sub RubberBand15_Hit
    If NOT Tilted Then
        IncreaseMatch
        ' añade 10 puntos
        AddScore 10
    End If
End Sub

Sub RubberBand6_Hit
    If NOT Tilted Then
        IncreaseMatch
        ' añade 10 puntos
        AddScore 10
    End If
End Sub

Sub RubberBand13_Hit
    If NOT Tilted Then
        IncreaseMatch
        ' añade 10 puntos
        AddScore 10
    End If
End Sub

Sub RubberBand19_Hit
    If NOT Tilted Then
        IncreaseMatch
        ' añade 10 puntos
        AddScore 10
    End If
End Sub

Sub RubberBand16_Hit
    If NOT Tilted Then
        IncreaseMatch
        ' añade 10 puntos
        AddScore 10
    End If
End Sub

'*********
' Bumpers
'*********

Sub Bumper1_Hit
    If NOT Tilted Then
        PlaySoundAt "fx_Bumper", bumper1
        ' añade algunos puntos
        If LightBumper1.State = 1 Then
            AddScore 100
        Else
            AddsCore 10
        End If
    End If
End Sub

Sub Bumper2_Hit
    If NOT Tilted Then
        PlaySoundAt "fx_Bumper", bumper2
        ' añade algunos puntos
        If LightBumper2.State = 1 Then
            AddScore 100
        Else
            AddsCore 10
        End If
    End If
End Sub

'***********************************
'    Dianas abatibles izquierda
'***********************************

Sub Drop1_Hit
    PlaySoundAtBall "fx_droptarget"
    If Tilted Then Exit Sub
    AddScore 500
    AddBonus 1
    li3.State = 1
    ' haz algún chequeo
    CheckLBankTargets
End Sub

Sub Drop2_Hit
    PlaySoundAtBall "fx_droptarget"
    If Tilted Then Exit Sub
    AddScore 500
    AddBonus 1
    li4.State = 1
    ' haz algún chequeo
    CheckLBankTargets
End Sub

Sub Drop3_Hit
    PlaySoundAtBall "fx_droptarget"
    If Tilted Then Exit Sub
    AddScore 500
    AddBonus 1
    li5.State = 1
    ' haz algún chequeo
    CheckLBankTargets
End Sub

Sub Drop4_Hit
    PlaySoundAtBall "fx_droptarget"
    If Tilted Then Exit Sub
    AddScore 500
    AddBonus 1
    li6.State = 1
    ' haz algún chequeo
    CheckLBankTargets
End Sub

Sub Drop5_Hit
    PlaySoundAtBall "fx_droptarget"
    If Tilted Then Exit Sub
    AddScore 500
    AddBonus 1
    li7.State = 1
    ' haz algún chequeo
    CheckLBankTargets
End Sub

Sub Drop6_Hit
    PlaySoundAtBall "fx_droptarget"
    If Tilted Then Exit Sub
    AddScore 500
    AddBonus 1
    li8.State = 1
    ' haz algún chequeo
    CheckLBankTargets
End Sub

Sub CheckLBankTargets
    Dim tmp
    tmp = 0
    For each x in aLeftDropLights
        tmp = tmp + x.State
    Next                                        'mira si todas las dianas están abatidas
    If tmp = 8 Then                             'todas las dianas han sido tumbadas
        BankLDown = BankLDown + 1               'significa que han sido derribadas todas
        li1.State = 1                           'enciende la luz doble bonos a la izquierda
        SpecialReady = SpecialReady + 1
        If SpecialReady = 2 Then li19.State = 1 ' enciende especial
        vpmtimer.addtimer 1000, "BankLDropup '"
    End If
End Sub

Sub BankLDropup
    PlaySoundAt "fx_resetdrop", Drop3
    For each x in aLeftDropLights
        x.State = 0
    Next
    Drop1.IsDropped = 0
    Drop2.IsDropped = 0
    Drop3.IsDropped = 0
    Drop4.IsDropped = 0
    Drop5.IsDropped = 0
    Drop6.IsDropped = 0
End Sub

'************************
'     Dianas fijas
'************************

Sub Target1_hit
    PlaySoundAtBall "fx_target"
    If Tilted Then Exit Sub
    IncreaseMatch
    AddScore 1000
    AddBonus 1
    If li9.State = 0 Then
        LightBumper1.State = 1
        li9.State = 1
    Else
        LightBumper2.State = 1
        li10.State = 1
        CheckLBankTargets
    End If
End Sub

'***********************************
'    Dianas abatibles derecha
'***********************************

Sub Drop14_Hit
    PlaySoundAtBall "fx_droptarget"
    If Tilted Then Exit Sub
    AddScore 500
    AddBonus 1
    li11.State = 1
    ' haz algún chequeo
    CheckRBankTargets
End Sub

Sub Drop13_Hit
    PlaySoundAtBall "fx_droptarget"
    If Tilted Then Exit Sub
    AddScore 500
    AddBonus 1
    li12.State = 1
    ' haz algún chequeo
    CheckRBankTargets
End Sub

Sub Drop12_Hit
    PlaySoundAtBall "fx_droptarget"
    If Tilted Then Exit Sub
    AddScore 500
    AddBonus 1
    li13.State = 1
    ' haz algún chequeo
    CheckRBankTargets
End Sub

Sub Drop11_Hit
    PlaySoundAtBall "fx_droptarget"
    If Tilted Then Exit Sub
    AddScore 500
    AddBonus 1
    li14.State = 1
    ' haz algún chequeo
    CheckRBankTargets
End Sub

Sub Drop10_Hit
    PlaySoundAtBall "fx_droptarget"
    If Tilted Then Exit Sub
    AddScore 500
    AddBonus 1
    li15.State = 1
    ' haz algún chequeo
    CheckRBankTargets
End Sub

Sub Drop9_Hit
    PlaySoundAtBall "fx_droptarget"
    If Tilted Then Exit Sub
    AddScore 500
    AddBonus 1
    li16.State = 1
    ' haz algún chequeo
    CheckRBankTargets
End Sub

Sub Drop8_Hit
    PlaySoundAtBall "fx_droptarget"
    If Tilted Then Exit Sub
    AddScore 500
    AddBonus 1
    li17.State = 1
    ' haz algún chequeo
    CheckRBankTargets
End Sub

Sub Drop7_Hit
    PlaySoundAtBall "fx_droptarget"
    If Tilted Then Exit Sub
    AddScore 500
    AddBonus 1
    li18.State = 1
    ' haz algún chequeo
    CheckRBankTargets
End Sub

Sub CheckRBankTargets
    Dim tmp
    tmp = 0
    For each x in aRightDropLights
        tmp = tmp + x.State
    Next                                        'mira si todas las dianas están abatidas
    If tmp = 8 Then                             'todas las dianas han sido tumbadas
        BankRDown = BankRDown + 1               'significa el número de veces que las dianas han sido derribadas
        li2.State = 1                           'enciende la luz doble bonos a la derecha
        SpecialReady = SpecialReady + 1
        If SpecialReady = 2 Then li19.State = 1 ' enciende especial
        vpmtimer.addtimer 1000, "BankRDropup '"
    End If
End Sub

Sub BankRDropup
    PlaySoundAt "fx_resetdrop", Drop11
    For each x in aRightDropLights
        x.State = 0
    Next
    Drop7.IsDropped = 0
    Drop8.IsDropped = 0
    Drop9.IsDropped = 0
    Drop10.IsDropped = 0
    Drop11.IsDropped = 0
    Drop12.IsDropped = 0
    Drop13.IsDropped = 0
    Drop14.IsDropped = 0
End Sub

'******************************
'     Pasillos inferiores
'******************************

Sub Trigger1_Hit
    PlaySoundAt "fx_sensor", Trigger1
    If Tilted Then Exit Sub
    AddScore 1000
    If li1.State = 1 Then ' la luz del Especial está encendida
        DoubleBonus = 2
    End If
End Sub

Sub Trigger4_Hit
    PlaySoundAt "fx_sensor", Trigger4
    If Tilted Then Exit Sub
    AddScore 1000
    If li2.State = 1 Then ' la luz del Especial está encendida
        DoubleBonus = 2
    End If
End Sub

Sub Trigger2_Hit
    PlaySoundAt "fx_sensor", Trigger2
    If Tilted Then Exit Sub
    AddScore 500
End Sub

Sub Trigger3_Hit
    PlaySoundAt "fx_sensor", Trigger3
    If Tilted Then Exit Sub
    AddScore 500
End Sub

'******************************
'     Pasillos superiores
'******************************

Sub Trigger5_Hit ' pasillo superior central
    PlaySoundAt "fx_sensor", Trigger5
    If Tilted Then Exit Sub
    AddScore 1000
    If li19.State = 1 Then
        AwardSpecial
        li19.State = 0
        SpecialReady = 0
    End If
End Sub

' ****************************************************************
' Ace High / IPD No. 7 / February, 1957 / 1 Player
' D. Gottlieb & Company (1931-1977) [Trade Name: Gottlieb]
' VPX8 table by jpsalas and Halen on the graphics, version 6.0.0
' ****************************************************************

Option Explicit
Randomize

Const BallSize = 50
Const BallMass = 1

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
Const TableName = "AceHigh" ' se usa para cargar y grabar los highscore y creditos
Const cGameName = "AceHigh" ' para el B2S y DOF
Const MaxPlayers = 1        ' de 1 a 4
Const MaxMultiplier = 1     ' limita el bonus multiplicador a 5
Const Special1 = 4200000    ' puntuación a obtener para partida extra
Const Special2 = 4500000
Const Special3 = 5000000
Const Special4 = 5000000
Const Special5 = 6000000
Const Special6 = 6500000
Const Special7 = 7000000

' Variables Globales
Dim PlayersPlayingGame
Dim CurrentPlayer
Dim Credits
Dim Bonus
Dim DoubleBonus
Dim BallsRemaining(4)
Dim Special1Awarded(4)
Dim ExtraBallsAwards(4)
Dim Special2Awarded(4)
Dim Special3Awarded(4)
Dim Special4Awarded(4)
Dim Special5Awarded(4)
Dim Special6Awarded(4)
Dim Special7Awarded(4)
Dim Score(4)
Dim HighScore
Dim Match
Dim Tilt
Dim TiltSensitivity
Dim Tilted
Dim Add1
Dim Add10
Dim Add100
Dim x

' Variables de control
Dim LastSwitchHit
Dim BallsOnPlayfield
Dim BallsEjected

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

' *********************************************************************
'                Rutinas comunes para todas las mesas
' *********************************************************************

Sub Table1_Init()
    Dim x

    ' Inicializar diversos objetos de la mesa, como droptargets, animations...
    VPObjects_Init
    LoadEM

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
    Add1 = 0
    Add10 = 0
    Add100 = 0

    ' pone la mesa en modo de espera
    EndOfGame

    'Enciende las luces GI despues de un segundo
    vpmtimer.addtimer 1000, "GiOn '"

    ' Quita las puntuaciones cuando la mesa se juega en modo FS
    If Table1.ShowDT then
        For each x in aReels
            x.Visible = 1
        Next
    Else
        For each x in aReels
            x.Visible = 0
        Next
    End If
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

    ' tilt keys
    If keycode = LeftTiltKey Then Nudge 90, 5:PlaySound "fx_nudge", 0, 1, -0.1, 0.25:CheckTilt
    If keycode = RightTiltKey Then Nudge 270, 5:PlaySound "fx_nudge", 0, 1, 0.1, 0.25:CheckTilt
    If keycode = CenterTiltKey Then Nudge 0, 6:PlaySound "fx_nudge", 0, 1, 1, 0.25:CheckTilt
    If keycode = MechanicalTilt Then Tilt = 16:CheckTilt

        If Credits = 0 Then DOF 200, DOFOff
        ' teclas de los flipers
        If keycode = LeftFlipperKey Then SolLFlipper 1
        If keycode = RightFlipperKey Then SolRFlipper 1

        ' tecla de empezar el juego
        ' si el juego ha empezado saca una bola nueva
        If keycode = StartGameKey AND AutoBallLoading Then
            EjectBall
        End If
        If Keycode = RightMagnaSave And AutoBallLoading = False Then
            If bJustStarted Then
                FirstBall
            Else
                EjectBall
            End If
        End If
    Else ' If (GameInPlay)

        If keycode = StartGameKey Then
            If(bFreePlay = True) Then
                If(BallsOnPlayfield = 0) Then
                    ResetScores
                    ResetForNewGame()
                End If
            Else
                If(Credits> 0) Then
                    If(BallsOnPlayfield = 0) Then
                        Credits = Credits - 1
                        UpdateCredits
                        ResetScores
                        ResetForNewGame()
                    End If
                Else
                    DOF 200, DOFOff
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
    If B2SOn then
        Controller.Stop
    End If
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
    PlaySound "fx_rubber_flipper", 0, Vol(ActiveBall), pan(ActiveBall), 0.1, 0, 0, 0, AudioFade(ActiveBall)
End Sub

Sub RightFlipper_Collide(parm)
    PlaySound "fx_rubber_flipper", 0, Vol(ActiveBall), pan(ActiveBall), 0.1, 0, 0, 0, AudioFade(ActiveBall)
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
    DOF 201, DOFOn
    PlaySound "fx_gion"
    Dim bulb
    For each bulb in aGiLights
        bulb.State = 1
    Next
End Sub

Sub GiOff 'apaga las luces GI
    DOF 201, DOFOff
    PlaySound "fx_gioff"
    Dim bulb
    For each bulb in aGiLights
        bulb.State = 0
    Next
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
    If Tilt> 30 Then              'Si la variable "Tilt" es más de 30 entonces haz falta. Normalmente es 15 el valor
        Tilted = True
        TurnON aTilt              'muestra Tilt en la pantalla
        If B2SOn then
            Controller.B2SSetTilt 1
        end if
        DisableTable True
        TiltRecoveryTimer.Enabled = True 'empieza una pausa a fin de que todas las bolas se cuelen
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
        Bumper004.Threshold = 100
        LeftSlingshot.Disabled = 1
        RightSlingshot.Disabled = 1
    Else
        'enciende de nuevo todas las luces GI, bumpers y slingshots
        GiOn
        Bumper001.Threshold = 1
        Bumper002.Threshold = 1
        Bumper003.Threshold = 1
        Bumper004.Threshold = 1
        LeftSlingshot.Disabled = 0
        RightSlingshot.Disabled = 0
    End If
End Sub

Sub TiltRecoveryTimer_Timer()
    ' si todas las bolas se han colado, entonces ..
    If(BallsOnPlayfield = 0) Then
        '... haz el fin de bola normal
        'finaliza la partida con la siguente línea
        BallsRemaining(CurrentPlayer) = 0
        EndOfBall()
        TiltRecoveryTimer.Enabled = False
    End If
' de lo contrario esta rutina continúa hasta que todas las bolas se han colado
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

'*******************************
'   JP's VP10 Rolling Sounds
'*******************************

Const tnob = 19   'total number of balls
Const lob = 0     'number of locked balls
Const maxvel = 25 'max ball velocity
ReDim rolling(tnob)
InitRolling

Sub InitRolling
    Dim i
    For i = 0 to tnob
        rolling(i) = False
    Next
    RollingTimer.Enabled = 1
End Sub

Sub RollingTimer_Timer
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
    Dim i, itemp

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
        ExtraBallsAwards(i) = 0
        Special1Awarded(i) = False
        Special2Awarded(i) = False
        Special3Awarded(i) = False
        Special4Awarded(i) = False
        Special5Awarded(i) = False
        Special6Awarded(i) = False
        Special7Awarded(i) = False
        BallsRemaining(i) = BallsPerGame
    Next
    DoubleBonus = 1
    Bonus = 0
    UpdateBallInPlay

    Clear_Match
    If B2SOn then
        for itemp = 50 to 57
            Controller.B2SSetData itemp, 0
        next
    end if
    ' inicializa otras variables
    Tilt = 0

    ' inicializa las variables del juego
    Game_Init()
    BallsRelease

    ' ahora puedes empezar una música si quieres
    ' empieza la rutina "Firstball" despues de una pequeña pausa
    bJustStarted = True
    If AutoBallLoading Then
        vpmtimer.addtimer 2000, "FirstBall '"
    End If
    If Credits = 0 Then DOF 200, DOFOff
End Sub

' esta pausa es para que la mesa tenga tiempo de poner los marcadores a cero y actualizar las luces

Sub FirstBall
    'debug.print "FirstBall"
    ' ajusta la mesa para una bola nueva, sube las dianas abatibles, etc
    ResetForNewPlayerBall()
    ' crea una bola nueva en la zona del plunger
    CreateNewBall()
    bJustStarted = False
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
    BallRelease.CreateSizedBallWithMass BallSize / 2, BallMass

    ' incrementa el número de bolas en el tablero, ya que hay que contarlas
    BallsOnPlayfield = BallsOnPlayfield + 1

    ' actualiza las luces del backdrop
    UpdateBallInPlay

    ' y expulsa la bola
    PlaySoundAt SoundFXDOF("fx_ballout", 104, DOFPulse, DOFContactors), BallRelease
    BallRelease.Kick 270, 2
End Sub

' El jugador ha perdido su bola, y ya no hay más bolas en juego
' Empieza a contar los bonos

Sub EndOfBall()
    'debug.print "EndOfBall"
    BallsRemaining(CurrentPlayer) = BallsRemaining(Currentplayer) - 1
    Dim AwardPoints, TotalBonus, ii
    AwardPoints = 0
    TotalBonus = 0
    ' La primera se ha perdido. Desde aquí ya no se puede aceptar más jugadores
    bOnTheFirstBall = False
    vpmtimer.addtimer 400, "EndOfBall2 '"
End Sub

Sub EndOfBall2()
    'debug.print "EndOfBall2"
    ' si hubiese falta, quítala, y pon la cuenta a cero de la falta para el próximo jugador, ó bola

    Tilted = False
    Tilt = 0
    TurnOFF aTilt
    If B2SOn then
        Controller.B2SSetTilt 0
    end if
    DisableTable False 'activa de nuevo los bumpers y los slingshots

    ' ¿Es ésta la última bola?
    If(BallsRemaining(CurrentPlayer) <= 0) Then

        ' miramos si la puntuación clasifica como el Highscore
        CheckHighScore()
    End If

    ' ésta no es la última bola para éste jugador
    ' y si hay más de un jugador continúa con el siguente
    EndOfBallComplete()
End Sub

' Esta rutina se llama al final de la cuenta del bonus
' y pasa a la siguente bola o al siguente jugador
'
Sub EndOfBallComplete()
    'debug.print "EndOfBallComplete"
    Dim NextPlayer

    'debug.print "EndOfBall - Complete"

    ' ¿hay otros jugadores?
    If(PlayersPlayingGame> 1) Then
        ' entonces pasa al siguente jugador
        NextPlayer = CurrentPlayer + 1
        ' ¿vamos a pasar del último jugador al primero?
        ' (por ejemplo del jugador 4 al no. 1)
        If(NextPlayer> PlayersPlayingGame) Then
            NextPlayer = 1
        End If
    Else
        NextPlayer = CurrentPlayer
    End If

    'debug.print "Next Player = " & NextPlayer

    ' ¿Hemos llegado al final del juego? (todas las bolas se han jugado de todos los jugadores)
    If((BallsRemaining(CurrentPlayer) <= 0) AND(BallsRemaining(NextPlayer) <= 0) ) Then

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
        'AddScore 0

        ' hacemos un reset del la mesa para el siguente jugador (ó bola)
        ResetForNewPlayerBall()

        ' y sacamos una bola
        If AutoBallLoading Then
            CreateNewBall()
        End If
    End If
End Sub

Sub EjectBall 'only for this table
    BallsEjected = BallsEjected + 1
    If BallsEjected < BallsPergame Then
        CreateNewBall()
    End If
End Sub

' Esta función se llama al final del juego

Sub EndOfGame()
    'debug.print "EndOfGame"
    bGameInPLay = False
    If B2SOn then
        Controller.B2SSetGameOver 1
        Controller.B2SSetBallInPlay 0
        Controller.B2SSetPlayerUp 0
        Controller.B2SSetCanPlay 0
    end if
    ' asegúrate de que los flippers están en modo de reposo
    SolLFlipper 0
    SolRFlipper 0
    DOF 250, DOFPulse
    ' pon las luces en el modo de fin de juego
    StartAttractMode
End Sub

' Esta función calcula el no de bolas que quedan
Function Balls
    Dim tmp
    tmp = BallsPerGame - BallsRemaining(CurrentPlayer) + 1
    If tmp> BallsPerGame Then
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

Sub Verification_Match()
'    PlaySound "fx_match"
'    Match = INT(RND * 10)
'    Display_Match
'    If(Score(CurrentPlayer) MOD 10) = Match Then
'        PlaySound SoundFXDOF("fx_knocker", 300, DOFPulse, DOFKnocker)
'        DOF 230, DOFPulse
'        AddCredits 1
'        DOF 200, DOFOn
'    End If
End Sub

Sub Clear_Match()
'    TurnOFF aMatch
'    If B2SOn then
'        Controller.B2SSetMatch 0
'    end if
End Sub

Sub Display_Match()
'    aMatch(Match).State = 1
'    If B2SOn then
'        If Match = 0 then
'            Controller.B2SSetMatch 10
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
    DOF 229, DOFPulse

    ' haz la bola visible en el apron,
    kicker007.Enabled = 1 'asegúrate de que este está activado
    drain2.CreateBall
    drain2.kick 76, 4

    'si hay falta el sistema de tilt se encargará de continuar con la siguente bola/jugador
    If Tilted Then
        StopEndOfBallMode
    Else
        EndOfBall
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
    Dim tmp
    Select Case Points
        Case 10000, 100000, 1000000 '10.000 points in display
            Score(CurrentPlayer) = Score(CurrentPlayer) + points
            UpdateScore
            If Points = 10000 Then RotateLights
            ' sounds
            tmp = Points \ 10000
            PlaySound SoundFXDOF("bell" &tmp, 131, DOFPulse, DOFChimes)
        Case 50000
            Add1 = Add1 + 5
            AddScore1Timer.Enabled = TRUE
        Case 500000
            Add10 = Add10 + 5
            AddScore10Timer.Enabled = TRUE
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
    If Score(CurrentPlayer) >= Special4 AND Special4Awarded(CurrentPlayer) = False Then
        AwardSpecial
        Special4Awarded(CurrentPlayer) = True
    End If
    If Score(CurrentPlayer) >= Special5 AND Special5Awarded(CurrentPlayer) = False Then
        AwardSpecial
        Special5Awarded(CurrentPlayer) = True
    End If
    If Score(CurrentPlayer) >= Special6 AND Special6Awarded(CurrentPlayer) = False Then
        AwardSpecial
        Special6Awarded(CurrentPlayer) = True
    End If
    If Score(CurrentPlayer) >= Special7 AND Special7Awarded(CurrentPlayer) = False Then
        AwardSpecial
        Special7Awarded(CurrentPlayer) = True
    End If
End Sub

'******************************
'TIMER DE 1, 10 y 100 PUNTOS
'******************************

' hace sonar las campanillas según los puntos

Sub AddScore10Timer_Timer()
    if Add10> 0 then
        AddScore 100000
        Add10 = Add10 - 1
    Else
        Me.Enabled = FALSE
    End If
End Sub

Sub AddScore100Timer_Timer() 'not used
    if Add100> 0 then
        AddScore 1000000
        Add100 = Add100 - 1
    Else
        Me.Enabled = FALSE
    End If
End Sub

Sub AddScore1Timer_Timer()
    if Add1> 0 then
        AddScore 10000
        Add1 = Add1 - 1
    Else
        Me.Enabled = FALSE
    End If
End Sub

'***********************************************************************************
'        Score reels - puntuaciones - y actualiza otras luces del backdrop
'***********************************************************************************
'esta es al rutina que actualiza la puntuación del jugador
' solo 1 jugador

Sub UpdateScore
    ResetScores
    dim tmp, B2Sid
    tmp = Score(CurrentPlayer) \ 10000

    '10.000 - decenas
    Select Case tmp MOD 10
        Case 1:TurnON d1
        Case 2:TurnON d2
        Case 3:TurnON d3
        Case 4:TurnON d4
        Case 5:TurnON d5
        Case 6:TurnON d6
        Case 7:TurnON d7
        Case 8:TurnON d8
        Case 9:TurnON d9
    End Select
    If B2SOn Then
        B2sid = 50 + tmp MOD 10
        Controller.B2SSetData B2sid, 1
    End If

    '100.000 - centenas
    tmp = tmp \ 10
    Select Case tmp MOD 10
        Case 1:TurnON c1
        Case 2:TurnON c2
        Case 3:TurnON c3
        Case 4:TurnON c4
        Case 5:TurnON c5
        Case 6:TurnON c6
        Case 7:TurnON c7
        Case 8:TurnON c8
        Case 9:TurnON c9
    End Select
    If B2SOn Then
        B2sid = 60 + tmp MOD 10
        Controller.B2SSetData B2sid, 1
    End If

    '1.000.000 - millones
    tmp = tmp \ 10
    Select Case tmp MOD 10
        Case 1:mi1.State = 1
        Case 2:mi2.State = 1
        Case 3:mi3.State = 1
        Case 4:mi4.State = 1
        Case 5:mi5.State = 1
        Case 6:mi6.State = 1
        Case 7:mi7.State = 1
    End Select
    If B2SOn Then
        B2sid = 70 + tmp MOD 10
        Controller.B2SSetData B2sid, 1
    End If
End Sub

' pone todos los marcadores a 0
Sub ResetScores
    Dim i
    For each i in AllScores:i.State = 0:Next

    If B2SOn Then
        For i = 51 to 59:Controller.B2SSetData i, 0:Next
        For i = 61 to 69:Controller.B2SSetData i, 0:Next
        For i = 71 to 77:Controller.B2SSetData i, 0:Next
    End If
End Sub

Sub AddCredits(value)
    If Credits < 9 Then
        Credits = Credits + value
        CreditsReel.SetValue Credits
        UpdateCredits
    end if
End Sub

Sub UpdateCredits
    PlaySound "fx_relay"
    CreditsReel.SetValue credits
    If B2SOn then
        Controller.B2SSetCredits Credits
    end if
End Sub

Sub UpdateBallInPlay 'actualiza los marcadores de las bolas, el número de jugador y el número total de jugadores
End Sub

'*************************
' Partidas y bolas extras
'*************************

Sub AwardExtraBall()
    If NOT bExtraBallWonThisBall Then
        PlaySound SoundFXDOF("fx_knocker", 300, DOFPulse, DOFKnocker)
        DOF 230, DOFPulse
        ExtraBallsAwards(CurrentPlayer) = ExtraBallsAwards(CurrentPlayer) + 1
        bExtraBallWonThisBall = True
    'LightShootAgain.State = 1
    'If B2SOn then
    '    Controller.B2SSetShootAgain 1
    'end if
    END If
End Sub

Sub AwardSpecial()
    PlaySound SoundFXDOF("fx_knocker", 300, DOFPulse, DOFKnocker)
    DOF 230, DOFPulse
    AddCredits 1
End Sub

' ********************************
'        Attract Mode
' ********************************
' las luces simplemente parpadean de acuerdo a los valores que hemos puesto en el "Blink Pattern" de cada luz

Sub StartAttractMode()
    If bGameInPlay Then Exit Sub
    Dim x
    bAttractMode = True
    StartLightSeq
    ' enciente la luz de fin de partida, bola en juego, ++
    TurnON aGameOver
End Sub

Sub StopAttractMode()
    Dim x
    bAttractMode = False
    LightSeqAttract.StopPlay
    ResetScores
    ' apaga la luz de fin de partida
    TurnOFF aGameOver
End Sub

Sub StartLightSeq()
    'lights sequences
    LightSeqAttract.UpdateInterval = 250
    LightSeqAttract.Play SeqRandom, 10, , 50000
End Sub

Sub LightSeqAttract_PlayDone()
    StartLightSeq()
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

'****************************************
' VPX8 actualizaciones en tiempo real
'****************************************

Sub LeftFlipper_Animate:LeftFlipperP.Rotz = LeftFlipper.CurrentAngle:LeftFlipperRubber.Rotz = LeftFlipper.CurrentAngle:End Sub
Sub RightFlipper_Animate:RightFlipperP.Rotz = RightFlipper.CurrentAngle:RightFlipperRubber.Rotz = RightFlipper.CurrentAngle:End Sub

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
    BallsInit
    TurnOffPlayfieldLights()
End Sub

' variables de la mesa
Dim LightPos
Dim SpecialAwarded

Sub Game_Init() 'esta rutina se llama al principio de un nuevo juego

    'Empezar alguna música, si hubiera música en esta mesa
    PlaySound"GameStart"
    'Variables del juego
    BallsEjected = 0
    LightPos = 0:RotateLights
    SpecialAwarded = False

    'iniciar algún timer

    'Iniciar algunas luces
    TurnOffPlayfieldLights
    Light004.State = 1
    Light005.State = 1
    Light006.State = 1
    Light007.State = 1
    Light017.State = 1
    BumperL001.State = 0
    BumperL002.State = 0
    BumperL003.State = 0
    BumperL004.State = 0
End Sub

Sub StopEndOfBallMode()     'this sub is called after the last ball is drained
End Sub

Sub ResetNewBallVariables() 'inicia las variable para una bola ó jugador nuevo
End Sub

Sub ResetNewBallLights()    'enciende ó apaga las luces para una bola nueva
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
    PlaySoundAtBall SoundFXDOF("fx_slingshot", 103, DOFPulse, DOFcontactors)
    LeftSling004.Visible = 1
    Lemk.RotX = 26
    LStep = 0
    LeftSlingShot.TimerEnabled = True
    ' añade algunos puntos
    AddScore 10000
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
    If Tilted Then Exit Sub
    PlaySoundAtBall SoundFXDOF("fx_slingshot", 104, DOFPulse, DOFcontactors)
    RightSling004.Visible = 1
    Remk.RotX = 26
    RStep = 0
    RightSlingShot.TimerEnabled = True
    ' añade algunos puntos
    AddScore 10000
End Sub

Sub RightSlingShot_Timer
    Select Case RStep
        Case 1:RightSLing004.Visible = 0:RightSLing003.Visible = 1:Remk.RotX = 14
        Case 2:RightSLing003.Visible = 0:RightSLing002.Visible = 1:Remk.RotX = 2
        Case 3:RightSLing002.Visible = 0:Remk.RotX = -20:RightSlingShot.TimerEnabled = 0
    End Select
    RStep = RStep + 1
End Sub

'*********
' Bumpers
'*********

Sub Bumper001_Hit
    If Tilted Then Exit Sub
    PlaySoundAt "fx_Bumper", bumper001
    ' añade algunos puntos
    AddScore 10000 + 90000 * BumperL001.State
End Sub

Sub Bumper002_Hit
    If Tilted Then Exit Sub
    PlaySoundAt "fx_Bumper", bumper002
    ' añade algunos puntos
    AddScore 10000 + 90000 * BumperL002.State
End Sub

Sub Bumper003_Hit
    If Tilted Then Exit Sub
    PlaySoundAt "fx_Bumper", bumper003
    ' añade algunos puntos
    AddScore 10000 + 90000 * BumperL003.State
End Sub

Sub Bumper004_Hit
    If Tilted Then Exit Sub
    PlaySoundAt "fx_Bumper", bumper004
    ' añade algunos puntos
    AddScore 10000 + 90000 * BumperL004.State
End Sub

' passive bumpers

Sub Bumper005_Hit
    If Tilted Then Exit Sub
    PlaySoundAtBall "fx_passive_bumper"
    ' añade algunos puntos
    AddScore 10000
End Sub

Sub Bumper006_Hit
    If Tilted Then Exit Sub
    PlaySoundAtBall "fx_passive_bumper"
    ' añade algunos puntos
    AddScore 10000
End Sub

Sub Bumper007_Hit
    If Tilted Then Exit Sub
    PlaySoundAtBall "fx_passive_bumper"
    ' añade algunos puntos
    AddScore 10000
End Sub

'*****************
'     Pasillos
'*****************

' Drain

Sub Trigger001_Hit
    If Tilted Then Exit Sub
    If Light001.State Then AwardSpecial
End Sub

' outlanes

Sub Trigger002_Hit
    PlaySoundAt "fx_sensor", Trigger002
    If Tilted Then Exit Sub
    AddScore 100000
    Light004.State = 0
    BumperL003.State = 1
    CheckDrainSpecial
End Sub

Sub Trigger003_Hit
    PlaySoundAt "fx_sensor", Trigger003
    If Tilted Then Exit Sub
    AddScore 100000
    Light005.State = 0
    BumperL001.State = 1
    CheckDrainSpecial
End Sub

Sub Trigger004_Hit
    PlaySoundAt "fx_sensor", Trigger004
    If Tilted Then Exit Sub
    AddScore 100000
    Light006.State = 0
    BumperL002.State = 1
    CheckDrainSpecial
End Sub

Sub Trigger005_Hit
    PlaySoundAt "fx_sensor", Trigger005
    If Tilted Then Exit Sub
    AddScore 100000
    Light007.State = 0
    BumperL004.State = 1
    CheckDrainSpecial
End Sub

Sub CheckDrainSpecial
    If Light004.State + Light005.State + Light006.State + Light007.State = 4 Then Light001-State = 1
End Sub

'************
'   holes
'************

Sub UpperHole_Hit
    PlaySoundAt "fx_hole_enter", UpperHole
    UpperHoleW.IsDropped = 1
    vpmTimer.AddTimer 2000, "UpperHoleW.IsDropped = 0 '"
    If Tilted Then Exit Sub
    Addscore 500000
    If Light017.State Then vpmTimer.AddTimer 1000, "AddScore 500000 '"
    CheckSpecials
End Sub

Sub LowerHole_Hit
    PlaySoundAt "fx_hole_enter", LowerHole
    LowerHoleW.IsDropped = 1
    vpmTimer.AddTimer 2000, "LowerHoleW.IsDropped = 0 '"
    If Tilted Then Exit Sub
    Addscore 500000
    If Light018.State Then vpmTimer.AddTimer 1000, "AddScore 500000 '"
    CheckSpecials
End Sub

Sub CheckSpecials 'holes
    If Light002.State Then
        'Queens
        If Light008.State + Light011.State + Light014.State = 3 Then vpmTimer.AddTimer 200, "AwardSpecial '"
        'Kings
        If Light009.State + Light012.State + Light015.State = 3 Then
            vpmTimer.AddTimer 550, "AwardSpecial '"
            vpmTimer.AddTimer 700, "AwardSpecial '"
        End If
        'Aces
        If Light010.State Then vpmTimer.AddTimer 1050, "AwardSpecial '"
        If Light013.State Then vpmTimer.AddTimer 1200, "AwardSpecial '"
        If Light016.State Then vpmTimer.AddTimer 1450, "AwardSpecial '"
        If Light010.State + Light013.State + Light016.State = 3 Then
            vpmTimer.AddTimer 1850, "AwardSpecial '"
            vpmTimer.AddTimer 2100, "AwardSpecial '"
        End If
    End If
End Sub

'*************************
'     Targets
'*************************

Sub Target001_Hit
    PlaySoundAtBall "fx_target"
    If Tilted Then Exit Sub
    Addscore 50000
End Sub

Sub Target002_Hit
    PlaySoundAtBall "fx_target"
    If Tilted Then Exit Sub
    Addscore 50000
End Sub

Sub Target003_Hit
    PlaySoundAtBall "fx_target"
    If Tilted Then Exit Sub
    Addscore 50000
End Sub

' Targets Center

Sub Target1C_Hit
    PlaySoundAtBall "fx_target"
    If Tilted Then Exit Sub
    PlaySound "Bell100"
    If Light008.State = 0 Then
        Light008.State = 1
    ElseIf Light009.State = 0 Then
        Light009.State = 1
    ElseIf Light010.State = 0 Then
        Light010.State = 1
    End If
    If Light008.State + Light011.State + Light014.State = 3 Then Light002.State = 1:Light003.State = 3
    'turn on the left sling Light001
    'check first if the right one is on
    If gi012.State = 0 Then TurnON LS
End Sub

Sub Target2C_Hit
    PlaySoundAtBall "fx_target"
    If Tilted Then Exit Sub
    PlaySound "Bell100"
    If Light011.State = 0 Then
        Light011.State = 1
    ElseIf Light012.State = 0 Then
        Light012.State = 1
    Elseif Light013.State = 0 Then
        Light013.State = 1
    End If
    If Light008.State + Light011.State + Light014.State = 3 Then Light002.State = 1:Light003.State = 3
End Sub

Sub Target3C_Hit
    PlaySoundAtBall "fx_target"
    If Tilted Then Exit Sub
    PlaySound "Bell100"
    If Light014.State = 0 Then
        Light014.State = 1
    Elseif Light015.State = 0 Then
        Light015.State = 1
    Elseif Light016.State = 0 Then
        Light016.State = 1
    End If
    If Light008.State + Light011.State + Light014.State = 3 Then Light002.State = 1:Light003.State = 3
    If gi013.State = 0 Then TurnON RS
End Sub

'***********************
' Rotation of the lights
'***********************

Sub RotateLights
    PlaySound "Gion"
    Light017.State = LightPos
    LightPos = ABS(LightPos - 1)
    Light018.State = LightPos
End Sub

'******************
' Gomas - 1 Points
'******************

Sub rlband001_Hit:Addscore 10000:ToggelSlings:End Sub 'left lower sling
Sub rlband003_Hit:Addscore 10000:ToggelSlings:End Sub 'right lower sling

Sub rsband009_Hit:Addscore 10000:End Sub
Sub rsband008_Hit:Addscore 10000:End Sub
Sub rsband007_Hit:Addscore 10000:End Sub
Sub rsband001_Hit:Addscore 10000:End Sub

Sub ToggelSlings
    If Light008.State + Light011.State + Light014.State> 1 Then 'if any of the Queen are on then one of them sling lights are on then can they be swapped
        If gi013.State Then
            TurnOFF LS
            TurnON RS
        Else
            TurnOFF RS
            TurnON LS
        End If
    End If
End Sub

'************************
'    Bolas del apron
'************************
' Las bolas del apron
Sub BallsRelease
    kicker003.enabled = 0
    kicker004.enabled = 0
    kicker005.enabled = 0
    kicker006.enabled = 0
    kicker007.enabled = 0
    kicker003.kick 90, 3
    kicker004.kick 90, 3
    kicker005.kick 90, 3
    kicker006.kick 90, 3
    kicker007.kick 90, 3
End Sub

Sub BallsInit
    kicker003.createball
    kicker004.createball
    kicker005.createball
    kicker006.createball
    kicker007.createball
End Sub

Sub kicker008_Hit
    PlaySoundAt "fx_kicker_enter", kicker008
    Me.destroyball
End sub

Sub kicker007_hit:kicker006.Enabled = 1:End Sub
Sub kicker006_hit:kicker005.Enabled = 1:End Sub
Sub kicker005_hit:kicker004.Enabled = 1:End Sub
Sub kicker004_hit:kicker003.Enabled = 1:End Sub

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
HSScore(1) = 1500000
HSScore(2) = 1400000
HSScore(3) = 1300000
HSScore(4) = 1200000
HSScore(5) = 1100000

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
                'PlaySound("fx_Bong")
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

'*********************************
' Table Options F12 User Options
'*********************************
' Table1.Option arguments are:
' - option name, minimum value, maximum value, step between valid values, default value, unit (0=None, 1=Percent), an optional array of literal strings

Dim LUTImage, BallsPerGame, AutoBallLoading

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

    ' Auto Ball Loading
    x = Table1.Option("Auto Ball Loading", 0, 1, 1, 1, 0, Array("Yes", "No") )
    If x = 1 Then AutoBallLoading = False Else AutoBallLoading = True
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

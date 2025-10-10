' ****************************************************************
'               VISUAL PINBALL X EM Script por JPSalas
'         Script Básico para juegos EM Script hasta 4 players
'        usa el core.vbs para funciones extras
' Hi-Diver / IPD No. 1165 / April, 1959 / 1 Player
'                       VPX8 Version 6.0.0
' ****************************************************************

Option Explicit
Randomize

'
' DOF config - leeoneil
'
' Option for more lights effects with DOF (Undercab effects on bumpers and slingshots)
' Replace "False" by "True" to activate (False by default)

Const Epileptikdof = False
'
' Flippers L/R - 701/702
' Slingshots L/R - 703/704
' Bumpers L/R - 705/706
' Targets L/R - 709/710
' Kickers - 711/712/713
'
' LED backboard
' Flasher Outside Left - 534/560/519/511/520/504/506
' Flasher left - 534/532/530/560/521/515/512/520/502/509
' Flasher center - 560/520/516/508
' Flasher right - 534/533/531/560/522/517/520/503/510
' Flasher Outside Right - 534/560/535/518/514/520/505/507
'
' Start Button - 200
' Undercab - 501/280
' Strobe - 303
' Knocker - 300
' Chimes - 301/302

' Valores Constantes de las físicas de los flippers - se usan en la creación de las bolas, tienen que cargarse antes del core.vbs
Const BallSize = 25 ' el radio de la bola. el tamaño normal es 50 unidades de VP
Const BallMass = 1  ' la pesadez de la bola

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
Const TableName = "hidiver" ' se usa para cargar y grabar los highscore y creditos
Const cGameName = "hidiver" ' para el B2S
Const MaxPlayers = 1        ' de 1 a 4
Const MaxMultiplier = 1     ' limita el bonus multiplicador a 5
Const Special1 = 5700000    ' puntuación a obtener para partida extra
Const Special2 = 6500000    ' puntuación a obtener para partida extra
Const Special3 = 7000000    ' puntuación a obtener para partida extra
Const Special4 = 7500000    ' puntuación a obtener para partida extra
Const Special5 = 7900000    ' puntuación a obtener para partida extra

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
Dim Score(4)
Dim HighScore
Dim Match
Dim Tilt
Dim TiltSensitivity
Dim Tilted
Dim Add10
Dim Add100
Dim B2sid 'STAT
Dim i, j, x

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
    Credits = 0
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
    Add10 = 0
    Add100 = 0

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
    If B2SOn Then Controller.B2SSetData 201, 1 'STAT
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
            If Not bFreePlay Then DOF 200, DOFOn
            AddCredits 1
            PlaySound "fx_coin"
        End If
    End If

    ' el plunger
    If keycode = PlungerKey Then
        Plunger.Pullback
        PlaySoundAt "fx_plungerpull", plunger
    End If

    ' teclas de la falta
    If keycode = LeftTiltKey Then Nudge 90, 8:PlaySound "fx_nudge", 0, 1, -0.1, 0.25:CheckTilt
    If keycode = RightTiltKey Then Nudge 270, 8:PlaySound "fx_nudge", 0, 1, 0.1, 0.25:CheckTilt
    If keycode = CenterTiltKey Then Nudge 0, 9:PlaySound "fx_nudge", 0, 1, 1, 0.25:CheckTilt
    If keycode = MechanicalTilt Then CheckTilt

    ' Funcionamiento normal de los flipers y otras teclas durante el juego

    If bGameInPlay AND NOT Tilted Then

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
            If(bFreePlay = True)Then
                If(BallsOnPlayfield = 0)Then
                    ResetScores
                    ResetForNewGame()
                End If
            Else
                If(Credits > 0)Then
                    DOF 200, DOFOn
                    If(BallsOnPlayfield = 0)Then
                        Credits = Credits - 1
                        UpdateCredits
                        ResetScores
                        ResetForNewGame()
                    End If
                Else
                    ' Not Enough Credits to start a game.
                    'PlaySound "so_nocredits"
                    DOF 200, DOFOff
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
    If B2SOn Then Controller.Stop 'STAT
End Sub

'********************
'     Flippers
'********************

Sub SolLFlipper(Enabled)
    If Enabled Then
        PlaySoundAt SoundFXDOF("fx_flipperup", 701, DOFOn, DOFFlippers), LeftFlipper
        LeftFlipper.RotateToEnd
        LeftFlipper001.RotateToEnd
        LeftFlipperOn = 1
    Else
        PlaySoundAt SoundFXDOF("fx_flipperdown", 701, DOFOff, DOFFlippers), LeftFlipper
        LeftFlipper.RotateToStart
        LeftFlipper001.RotateToStart
        LeftFlipperOn = 0
    End If
End Sub

Sub SolRFlipper(Enabled)
    If Enabled Then
        PlaySoundAt SoundFXDOF("fx_flipperup", 702, DOFOn, DOFFlippers), RightFlipper
        RightFlipper.RotateToEnd
        RightFlipper001.RotateToEnd
        RightFlipperOn = 1
    Else
        PlaySoundAt SoundFXDOF("fx_flipperdown", 702, DOFOff, DOFFlippers), RightFlipper
        RightFlipper.RotateToStart
        RightFlipper001.RotateToStart
        RightFlipperOn = 0
    End If
End Sub

' flippers top animations

Sub LeftFlipper_Animate:LeftFlipperP.RotZ = LeftFlipper.CurrentAngle: End Sub
Sub RightFlipper_Animate: RightFlipperP.RotZ = RightFlipper.CurrentAngle: End Sub
Sub LeftFlipper001_Animate:LeftFlipperP001.RotZ = LeftFlipper001.CurrentAngle: End Sub
Sub RightFlipper001_Animate: RightFlipperP001.RotZ = RightFlipper001.CurrentAngle: End Sub

' flippers hit Sound

Sub LeftFlipper_Collide(parm)
    PlaySound "fx_rubber_flipper", 0, parm / 60, pan(ActiveBall), 0.1, 0, 0, 0, AudioFade(ActiveBall)
End Sub

Sub RightFlipper_Collide(parm)
    PlaySound "fx_rubber_flipper", 0, parm / 60, pan(ActiveBall), 0.1, 0, 0, 0, AudioFade(ActiveBall)
End Sub

Sub LeftFlipper001_Collide(parm)
    PlaySound "fx_rubber_flipper", 0, parm / 60, pan(ActiveBall), 0.1, 0, 0, 0, AudioFade(ActiveBall)
End Sub

Sub RightFlipper001_Collide(parm)
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
    DOF 501, DOFOn
    Dim bulb
    For each bulb in aGiLights
        bulb.State = 1
    Next
End Sub

Sub GiOff 'apaga las luces GI
    DOF 501, DOFOff
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
    If Tilt > 30 Then             'Si la variable "Tilt" es más de 30 entonces haz falta. Normalmente es 15 el valor
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
        LeftFlipper001.RotateToStart
        RightFlipper001.RotateToStart
        Bumper001.Threshold = 100
        Bumper002.Threshold = 100
        LeftSlingshot.Disabled = 1
        RightSlingshot.Disabled = 1
        DOF 701, DOFOff
        DOF 702, DOFOff
    Else
        'enciende de nuevo todas las luces GI, bumpers y slingshots
        GiOn
        Bumper001.Threshold = 1
        Bumper002.Threshold = 1
        LeftSlingshot.Disabled = 0
        RightSlingshot.Disabled = 0
    End If
End Sub

Sub TiltRecoveryTimer_Timer()
    ' si todas las bolas se han colado, entonces ..
    If(BallsOnPlayfield = 0)Then
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
Const maxvel = 28 'max ball velocity
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

        If BallVel(BOT(b)) > 1 Then
            If BOT(b).z < 30 Then
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
        If BOT(b).VelZ < -1 and BOT(b).z < 55 and BOT(b).z > 27 Then 'height adjust for ball drop sounds
            PlaySound "fx_balldrop", 0, ABS(BOT(b).velz) / 17, Pan(BOT(b)), 0, Pitch(BOT(b)), 1, 0, AudioFade(BOT(b))
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
    BallsRelease

    ' ahora puedes empezar una música si quieres
    ' empieza la rutina "Firstball" despues de una pequeña pausa
    bJustStarted = True
    If AutoBallLoading Then
        vpmtimer.addtimer 2000, "FirstBall '"
    End If
End Sub

' esta pausa es para que la mesa tenga tiempo de poner los marcadores a cero y actualizar las luces

Sub FirstBall
    'debug.print "FirstBall"
    ' ajusta la mesa para una bola nueva, sube las dianas abatibles, etc
    ResetForNewPlayerBall()
    ' crea una bola nueva en la zona del plunger
    CreateNewBall()
    bJustStarted = False
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
    PlaySoundAt SoundFXDOF("fx_Ballrel", 704, DOFPulse, DOFcontactors), BallRelease
    If Epileptikdof = True Then DOF 518, DOFPulse End If
    BallRelease.Kick 270, 2
End Sub

' El jugador ha perdido su bola, y ya no hay más bolas en juego
' Empieza a contar los bonos

Sub EndOfBall()
    'debug.print "EndOfBall"
  BallsRemaining(CurrentPlayer) = BallsRemaining(Currentplayer) - 1
  EndOfBall2
    'End If
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

        ' ¿Es ésta la última bola?
        If(BallsRemaining(CurrentPlayer) <= 0)Then

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
    If B2SOn Then
        For i = 202 to 230:Controller.B2SSetData i, 0:Next 'STAT
        Controller.B2SSetData 201, 1
    End If
    'debug.print "EndOfGame"
    bGameInPLay = False
    If B2SOn then
        Controller.B2SSetGameOver 1
    'Controller.B2SSetBallInPlay 0
    'Controller.B2SSetPlayerUp 0
    'Controller.B2SSetCanPlay 0
    end if
    ' asegúrate de que los flippers están en modo de reposo
    SolLFlipper 0
    SolRFlipper 0
    DOF 560, DOFPulse

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
    Match = (Match + 1)MOD 10
End Sub

Sub Verification_Match()
    PlaySound "fx_match"
    Display_Match
    If((Score(CurrentPlayer) \ 10000)MOD 10) = Match Then
        PlaySound SoundFXDOF("fx_knocker", 300, DOFPulse, DOFKnocker)
        DOF 303, DOFPulse
        AddCredits 1
        DOF 200, DOFOn
    End If
End Sub

Sub Clear_Match()
    MatchReel.SetValue 0
    If B2SOn then
        Controller.B2SSetMatch 0
    end if
End Sub

Sub Display_Match()
    MatchReel.SetValue Match + 1
    If B2SOn then
        If Match = 0 then
            Controller.B2SSetMatch 100
        else
            Controller.B2SSetMatch Match * 10
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
    DOF 501, DOFOff

    ' haz la bola visible en el apron,
    kicker003.CreateBall
    kicker003.kick 76, 4
    kicker007.Enabled = 1 'asegúrate de que este está activado

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
    IncreaseMatch
    Select Case Points
        Case 0
            UpdateScore points
        Case 10000
            Score(CurrentPlayer) = Score(CurrentPlayer) + points
            UpdateScore points
            PlaySound SoundFXDOF("Bell-Nairobi", 301, DOFPulse, DOFChimes)
        Case 100000
            Score(CurrentPlayer) = Score(CurrentPlayer) + points
            UpdateScore points
            PlaySound SoundFXDOF("Bell", 302, DOFPulse, DOFChimes)
        Case 50000
            Add10 = Add10 + 5
            AddScore1Timer.Enabled = TRUE
        Case 500000
            Add100 = Add100 + 5
            AddScore100Timer.Enabled = TRUE
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
    If Score(CurrentPlayer) >= Special4 AND Special4Awarded(CurrentPlayer) = False Then
        AwardSpecial
        Special4Awarded(CurrentPlayer) = True
    End If
    If Score(CurrentPlayer) >= Special5 AND Special5Awarded(CurrentPlayer) = False Then
        AwardSpecial
        Special5Awarded(CurrentPlayer) = True
    End If

    ' lit up the backdrop lights according to the score, fair, good, excellent, super and genius

    If B2SOn Then
        For i = 21 to 25:Controller.B2SSetData i, 0:Next 'STAT
    End If

    If Score(CurrentPlayer) >= 3100000 AND Score(CurrentPlayer) < 4100000 Then
        t030.State = 1:If B2SOn Then Controller.B2SSetData 21, 1 'STAT
        Else
            t030.State = 0
    End If

    If Score(CurrentPlayer) >= 4100000 AND Score(CurrentPlayer) < 5100000 Then
        t029.State = 1:If B2SOn Then Controller.B2SSetData 22, 1 'STAT
        Else
            t029.State = 0
    End If

    If Score(CurrentPlayer) >= 5100000 AND Score(CurrentPlayer) < 6100000 Then
        t028.State = 1:If B2SOn Then Controller.B2SSetData 23, 1 'STAT
        Else
            t028.State = 0
    End If

    If Score(CurrentPlayer) >= 6100000 AND Score(CurrentPlayer) < 7100000 Then
        t027.State = 1:If B2SOn Then Controller.B2SSetData 24, 1 'STAT
        Else
            t027.State = 0
    End If

    If Score(CurrentPlayer) >= 7100000 Then
        t026.State = 1:If B2SOn Then Controller.B2SSetData 25, 1 'STAT
        Else
            t026.State = 0
    End If
End Sub

'******************************
'TIMER DE 10, 100 y 1000 PUNTOS
'******************************

' hace sonar las campanillas según los puntos

Sub AddScore10Timer_Timer()
    if Add10 > 0 then
        AddScore 10000
        Add10 = Add10 - 1
    Else
        Me.Enabled = FALSE
    End If
End Sub

Sub AddScore100Timer_Timer()
    if Add100 > 0 then
        AddScore 100000
        Add100 = Add100 - 1
    Else
        Me.Enabled = FALSE
    End If
End Sub

'*******************
'     BONOS
'*******************

' Bonus are called POINTS in this table, and they are only used to get extra games.

' avanza el bono y actualiza las luces
' Los bonos están limitados a 1500 puntos

Sub AddBonus(bonuspoints)
    If(Tilted = False)Then
        ' añade los bonos al jugador actual
        Bonus = Bonus + bonuspoints
        If Bonus = 15 Then AwardSpecial
        If Bonus = 17 Then AwardSpecial
        If Bonus = 19 Then AwardSpecial
        If Bonus > 19 Then
            Bonus = 19
        End If
        ' actualiza las luces
        UpdateBonusLights
    End if
End Sub

Sub UpdateBonusLights 'enciende o apaga las luces de los bonos según la variable "Bonus"
    Select Case Bonus
        Case 0:li001.State = 0:li002.State = 0:li003.State = 0:li004.State = 0:li005.State = 0:li006.State = 0:li007.State = 0:li008.State = 0:li009.State = 0:li010.State = 0
        Case 1:li001.State = 1:li002.State = 0:li003.State = 0:li004.State = 0:li005.State = 0:li006.State = 0:li007.State = 0:li008.State = 0:li009.State = 0:li010.State = 0
        Case 2:li001.State = 0:li002.State = 1:li003.State = 0:li004.State = 0:li005.State = 0:li006.State = 0:li007.State = 0:li008.State = 0:li009.State = 0:li010.State = 0
        Case 3:li001.State = 0:li002.State = 0:li003.State = 1:li004.State = 0:li005.State = 0:li006.State = 0:li007.State = 0:li008.State = 0:li009.State = 0:li010.State = 0
        Case 4:li001.State = 0:li002.State = 0:li003.State = 0:li004.State = 1:li005.State = 0:li006.State = 0:li007.State = 0:li008.State = 0:li009.State = 0:li010.State = 0
        Case 5:li001.State = 0:li002.State = 0:li003.State = 0:li004.State = 0:li005.State = 1:li006.State = 0:li007.State = 0:li008.State = 0:li009.State = 0:li010.State = 0
        Case 6:li001.State = 0:li002.State = 0:li003.State = 0:li004.State = 0:li005.State = 0:li006.State = 1:li007.State = 0:li008.State = 0:li009.State = 0:li010.State = 0
        Case 7:li001.State = 0:li002.State = 0:li003.State = 0:li004.State = 0:li005.State = 0:li006.State = 0:li007.State = 1:li008.State = 0:li009.State = 0:li010.State = 0
        Case 8:li001.State = 0:li002.State = 0:li003.State = 0:li004.State = 0:li005.State = 0:li006.State = 0:li007.State = 0:li008.State = 1:li009.State = 0:li010.State = 0
        Case 9:li001.State = 0:li002.State = 0:li003.State = 0:li004.State = 0:li005.State = 0:li006.State = 0:li007.State = 0:li008.State = 0:li009.State = 1:li010.State = 0
        Case 10:li001.State = 0:li002.State = 0:li003.State = 0:li004.State = 0:li005.State = 0:li006.State = 0:li007.State = 0:li008.State = 0:li009.State = 0:li010.State = 1
        Case 11:li001.State = 1:li002.State = 0:li003.State = 0:li004.State = 0:li005.State = 0:li006.State = 0:li007.State = 0:li008.State = 0:li009.State = 0:li010.State = 1
        Case 12:li001.State = 0:li002.State = 1:li003.State = 0:li004.State = 0:li005.State = 0:li006.State = 0:li007.State = 0:li008.State = 0:li009.State = 0:li010.State = 1
        Case 13:li001.State = 0:li002.State = 0:li003.State = 1:li004.State = 0:li005.State = 0:li006.State = 0:li007.State = 0:li008.State = 0:li009.State = 0:li010.State = 1
        Case 14:li001.State = 0:li002.State = 0:li003.State = 0:li004.State = 1:li005.State = 0:li006.State = 0:li007.State = 0:li008.State = 0:li009.State = 0:li010.State = 1
        Case 15:li001.State = 0:li002.State = 0:li003.State = 0:li004.State = 0:li005.State = 1:li006.State = 0:li007.State = 0:li008.State = 0:li009.State = 0:li010.State = 1
        Case 16:li001.State = 0:li002.State = 0:li003.State = 0:li004.State = 0:li005.State = 0:li006.State = 1:li007.State = 0:li008.State = 0:li009.State = 0:li010.State = 1
        Case 17:li001.State = 0:li002.State = 0:li003.State = 0:li004.State = 0:li005.State = 0:li006.State = 0:li007.State = 1:li008.State = 0:li009.State = 0:li010.State = 1
        Case 18:li001.State = 0:li002.State = 0:li003.State = 0:li004.State = 0:li005.State = 0:li006.State = 0:li007.State = 0:li008.State = 1:li009.State = 0:li010.State = 1
        Case 19:li001.State = 0:li002.State = 0:li003.State = 0:li004.State = 0:li005.State = 0:li006.State = 0:li007.State = 0:li008.State = 0:li009.State = 1:li010.State = 1

    End Select
End Sub

'***********************************************************************************
'        Score reels - puntuaciones - y actualiza otras luces del backdrop
'***********************************************************************************
'esta es al rutina que actualiza la puntuación del jugador
' solo 1 jugador

Sub ClearB2SScores 'STAT
    For i = 2 to 11:Controller.B2SSetData i, 0:Next
    For i = 10 to 90 Step 10:Controller.B2SSetData i, 0:Next
    For i = 100 to 190 Step 10:Controller.B2SSetData i, 0:Next
End Sub

Sub UpdateScore(playerpoints)
    dim tmp
    tmp = Score(CurrentPlayer) \ 1000

    '10
    tmp = tmp \ 10
    Select Case tmp MOD 10
        Case 0:t000.State = 1:t001.State = 0:t002.State = 0:t003.State = 0:t004.State = 0:t005.State = 0:t006.State = 0:t007.State = 0:t008.State = 0:t009.State = 0
        Case 1:t000.State = 0:t001.State = 1:t002.State = 0:t003.State = 0:t004.State = 0:t005.State = 0:t006.State = 0:t007.State = 0:t008.State = 0:t009.State = 0
        Case 2:t000.State = 0:t001.State = 0:t002.State = 1:t003.State = 0:t004.State = 0:t005.State = 0:t006.State = 0:t007.State = 0:t008.State = 0:t009.State = 0
        Case 3:t000.State = 0:t001.State = 0:t002.State = 0:t003.State = 1:t004.State = 0:t005.State = 0:t006.State = 0:t007.State = 0:t008.State = 0:t009.State = 0
        Case 4:t000.State = 0:t001.State = 0:t002.State = 0:t003.State = 0:t004.State = 1:t005.State = 0:t006.State = 0:t007.State = 0:t008.State = 0:t009.State = 0
        Case 5:t000.State = 0:t001.State = 0:t002.State = 0:t003.State = 0:t004.State = 0:t005.State = 1:t006.State = 0:t007.State = 0:t008.State = 0:t009.State = 0
        Case 6:t000.State = 0:t001.State = 0:t002.State = 0:t003.State = 0:t004.State = 0:t005.State = 0:t006.State = 1:t007.State = 0:t008.State = 0:t009.State = 0
        Case 7:t000.State = 0:t001.State = 0:t002.State = 0:t003.State = 0:t004.State = 0:t005.State = 0:t006.State = 0:t007.State = 1:t008.State = 0:t009.State = 0
        Case 8:t000.State = 0:t001.State = 0:t002.State = 0:t003.State = 0:t004.State = 0:t005.State = 0:t006.State = 0:t007.State = 0:t008.State = 1:t009.State = 0
        Case 9:t000.State = 0:t001.State = 0:t002.State = 0:t003.State = 0:t004.State = 0:t005.State = 0:t006.State = 0:t007.State = 0:t008.State = 0:t009.State = 1
    End Select

    'STAT
    If B2SOn Then
        ClearB2SScores
        B2sid = 100 + (tmp MOD 10) * 10
        Controller.B2SSetData B2sid, 1
    End If

    '100
    tmp = tmp \ 10
    Select Case tmp MOD 10
        Case 1:t010.State = 1:t011.State = 0:t012.State = 0:t013.State = 0:t014.State = 0:t015.State = 0:t016.State = 0:t017.State = 0:t018.State = 0
        Case 2:t010.State = 0:t011.State = 1:t012.State = 0:t013.State = 0:t014.State = 0:t015.State = 0:t016.State = 0:t017.State = 0:t018.State = 0
        Case 3:t010.State = 0:t011.State = 0:t012.State = 1:t013.State = 0:t014.State = 0:t015.State = 0:t016.State = 0:t017.State = 0:t018.State = 0
        Case 4:t010.State = 0:t011.State = 0:t012.State = 0:t013.State = 1:t014.State = 0:t015.State = 0:t016.State = 0:t017.State = 0:t018.State = 0
        Case 5:t010.State = 0:t011.State = 0:t012.State = 0:t013.State = 0:t014.State = 1:t015.State = 0:t016.State = 0:t017.State = 0:t018.State = 0
        Case 6:t010.State = 0:t011.State = 0:t012.State = 0:t013.State = 0:t014.State = 0:t015.State = 1:t016.State = 0:t017.State = 0:t018.State = 0
        Case 7:t010.State = 0:t011.State = 0:t012.State = 0:t013.State = 0:t014.State = 0:t015.State = 0:t016.State = 1:t017.State = 0:t018.State = 0
        Case 8:t010.State = 0:t011.State = 0:t012.State = 0:t013.State = 0:t014.State = 0:t015.State = 0:t016.State = 0:t017.State = 1:t018.State = 0
        Case 9:t010.State = 0:t011.State = 0:t012.State = 0:t013.State = 0:t014.State = 0:t015.State = 0:t016.State = 0:t017.State = 0:t018.State = 1
    End Select

    'STAT
    If B2SOn Then
        B2sid = (tmp MOD 10) * 10
        Controller.B2SSetData B2sid, 1
    End If

    '1000
    tmp = tmp \ 10
    Select Case tmp MOD 10
        Case 1:t019.State = 1:t020.State = 0:t021.State = 0:t022.State = 0:t023.State = 0:t024.State = 0:t025.State = 0
        Case 2:t019.State = 0:t020.State = 1:t021.State = 0:t022.State = 0:t023.State = 0:t024.State = 0:t025.State = 0
        Case 3:t019.State = 0:t020.State = 0:t021.State = 1:t022.State = 0:t023.State = 0:t024.State = 0:t025.State = 0
        Case 4:t019.State = 0:t020.State = 0:t021.State = 0:t022.State = 1:t023.State = 0:t024.State = 0:t025.State = 0
        Case 5:t019.State = 0:t020.State = 0:t021.State = 0:t022.State = 0:t023.State = 1:t024.State = 0:t025.State = 0
        Case 6:t019.State = 0:t020.State = 0:t021.State = 0:t022.State = 0:t023.State = 0:t024.State = 1:t025.State = 0
        Case 7:t019.State = 0:t020.State = 0:t021.State = 0:t022.State = 0:t023.State = 0:t024.State = 0:t025.State = 1
    End Select

    'STAT
    If B2SOn Then
        B2sid = (tmp MOD 10)
        If B2sid = 1 Then B2sid = 11
        Controller.B2SSetData B2sid, 1
    End If
End Sub

' pone todos los marcadores a 0
Sub ResetScores
    'ScoreReel1.ResetToZero
    AddScore 0
End Sub

Sub AddCredits(value)
    If Credits < 10 Then
        Credits = Credits + value
        CreditsReel.SetValue Credits
        UpdateCredits
    end if
End Sub

Sub UpdateCredits
    'If Credits > 0 Then 'in las mesas de Bally
    'CreditLight.State = 1
    'Else
    'CreditLight.State = 0
    'End If
    PlaySound "fx_relay"
    CreditsReel.SetValue credits
    If B2SOn then Controller.B2SSetCredits Credits 'STAT
    If Credits = 0 Then DOF 200, DOFOff
    If Credits > 0 Then DOF 200, DOFOn
End Sub

Sub UpdateBallInPlay 'actualiza los marcadores de las bolas, el número de jugador y el número total de jugadores
'Select Case Balls
'    Case 0:BallDisplay.ImageA = "Ballnr0"
'    Case 1:BallDisplay.ImageA = "Ballnr1"
'    Case 2:BallDisplay.ImageA = "Ballnr2"
'    Case 3:BallDisplay.ImageA = "Ballnr3"
'    Case 4:BallDisplay.ImageA = "Ballnr4"
'    Case 5:BallDisplay.ImageA = "Ballnr5"
'End Select
'If B2SOn then
'    Controller.B2SSetBallInPlay Balls
'end if
End Sub

'*************************
' Partidas y bolas extras
'*************************

Sub AwardExtraBall()
    If NOT bExtraBallWonThisBall Then
        PlaySound SoundFXDOF("fx_knocker", 300, DOFPulse, DOFKnocker)
        DOF 303, DOFPulse
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
    DOF 303, DOFPulse
    AddCredits 1
    DOF 200, DOFOn
End Sub

' ********************************
'        Attract Mode
' ********************************
' las luces simplemente parpadean de acuerdo a los valores que hemos puesto en el "Blink Pattern" de cada luz

Sub StartAttractMode()
    Dim x
    bAttractMode = True
    'StartLightSeq
    ' enciente la luz de fin de partida, bola en juego, ++
    GameOverR.SetValue 1
'BallDisplay.ImageA = "ballnr0"
End Sub

Sub StopAttractMode()
    Dim x
    bAttractMode = False
    'LightSeqAttract.StopPlay
    ResetScores
    ' apaga la luz de fin de partida
    GameOverR.SetValue 0
End Sub

Sub StartLightSeq()
'lights sequences
'LightSeqAttract.UpdateInterval = 250
'LightSeqAttract.Play SeqRandom, 10, , 50000
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
    If Not FileObj.FolderExists(UserDirectory)then
        Credits = 0
        DOF 200, DOFOff
        Exit Sub
    End If
    If Not FileObj.FileExists(UserDirectory & TableName& ".txt")then
        Credits = 0
        DOF 200, DOFOff
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

'****************************************
' Actualizaciones en tiempo real
'****************************************
' se usa sobre todo para hacer animaciones o sonidos que cambian en tiempo real
' como por ejemplo para sincronizar los flipers, puertas ó molinillos con primitivas

Sub GameTimer_Timer
    RollingUpdate 'actualiza el sonido de la bola rodando
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

Sub VPObjects_Init 'inicia objetos al empezar
    BallsInit
    TurnOffPlayfieldLights()
End Sub

' variables de la mesa
Dim AlternLights
Dim WheelPos
Dim WheelStep

Sub Game_Init() 'esta rutina se llama al principio de un nuevo juego

    'Empezar alguna música, si hubiera música en esta mesa

    'iniciar variables, en esta mesa hay muy pocas variables ya que usamos las luces
    AlternLights = False
    WheelPos = 0:WheelStep = 0
  BallsEjected = 0
    'iniciar algún timer

    'Iniciar algunas luces
    TurnOffPlayfieldLights
    RotateLights
End Sub

Sub StopEndOfBallMode()     'this sub is called after the last ball is drained

End Sub

Sub ResetNewBallVariables() 'inicia las variable para una bola ó jugador nuevo
    AlternLights = False
End Sub

Sub ResetNewBallLights() 'enciende ó apaga las luces para una bola nueva
    LightBumper001.State = 0
    LightBumper002.State = 0
    ls1.State = 0
    ls2.State = 0
    ls3.State = 0
    ls4.State = 0
    rs1.State = 0
    rs2.State = 0
    rs3.State = 0
    rs4.State = 0
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
    PlaySoundAtBall SoundFXDOF("fx_slingshot", 703, DOFPulse, DOFcontactors)
    DOF 519, DOFPulse
    If Epileptikdof = True Then DOF 521, DOFPulse End If
    If Epileptikdof = True Then DOF 630, DOFPulse End If
    If Epileptikdof = True Then DOF 280, DOFPulse End If
    LeftSling4.Visible = 1
    Lemk.RotX = 26
    LStep = 0
    LeftSlingShot.TimerEnabled = True
    ' añade algunos puntos
    AddScore 10000 + 90000 * ls1.State
    RotateLights
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
    PlaySoundAtBall SoundFXDOF("fx_slingshot", 704, DOFPulse, DOFcontactors)
    DOF 535, DOFPulse
    If Epileptikdof = True Then DOF 522, DOFPulse End If
    If Epileptikdof = True Then DOF 631, DOFPulse End If
    If Epileptikdof = True Then DOF 280, DOFPulse End If
    RightSling4.Visible = 1
    Remk.RotX = 26
    RStep = 0
    RightSlingShot.TimerEnabled = True
    ' añade algunos puntos
    AddScore 10000 + 90000 * rs1.State
    RotateLights
End Sub

Sub RightSlingShot_Timer
    Select Case RStep
        Case 1:RightSLing4.Visible = 0:RightSLing3.Visible = 1:Remk.RotX = 14
        Case 2:RightSLing3.Visible = 0:RightSLing2.Visible = 1:Remk.RotX = 2
        Case 3:RightSLing2.Visible = 0:Remk.RotX = -20:RightSlingShot.TimerEnabled = 0
    End Select
    RStep = RStep + 1
End Sub

'*************
' Rubberbands
'*************

Sub RubberBand004_Hit
    If Tilted Then Exit Sub
    PlaySoundAtBall "fx_passive_bumper"
    AddScore 10000
    RotateLights
End Sub

Sub RubberBand005_Hit
    If Tilted Then Exit Sub
    PlaySoundAtBall "fx_passive_bumper"
    AddScore 10000
    RotateLights
End Sub

Sub RubberBand006_Hit
    If Tilted Then Exit Sub
    PlaySoundAtBall "fx_passive_bumper"
    AddScore 10000
    RotateLights
End Sub

Sub RubberBand007_Hit
    If Tilted Then Exit Sub
    PlaySoundAtBall "fx_passive_bumper"
    AddScore 10000
    RotateLights
End Sub

'*********
' Bumpers
'*********

Sub Bumper001_Hit
    If Tilted Then Exit Sub
    PlaySoundAt SoundFXDOF("fx_Bumper", 705, DOFPulse, DOFContactors), bumper001
    DOF 502, DOFPulse
    If Epileptikdof = True Then DOF 280, DOFPulse End If
    ' añade algunos puntos
    AddScore 10000 + 90000 * LightBumper001.State
    RotateLights
End Sub

Sub Bumper002_Hit
    If Tilted Then Exit Sub
    PlaySoundAt SoundFXDOF("fx_Bumper", 706, DOFPulse, DOFContactors), bumper002
    DOF 503, DOFPulse
    If Epileptikdof = True Then DOF 280, DOFPulse End If
    ' añade algunos puntos
    AddScore 10000 + 90000 * LightBumper002.State
    RotateLights
End Sub

'****************
' Pasive Bumpers
'****************

Sub TopLBumper_Hit
    If Tilted Then Exit Sub
    PlaySoundAt "fx_passive_bumper", LightTopLBumper
    DOF 511, DOFPulse
    ' añade algunos puntos
    AddScore 10000
    RotateLights
End Sub

Sub TopLBumper002_Hit
    If Tilted Then Exit Sub
    PlaySoundAt "fx_passive_bumper", LightTopLBumper001
    DOF 512, DOFPulse
    ' añade algunos puntos
    AddScore 10000
    RotateLights
End Sub

Sub TopLBumper004_Hit
    If Tilted Then Exit Sub
    PlaySoundAt "fx_passive_bumper", LightTopLBumper002
    DOF 513, DOFPulse
    ' añade algunos puntos
    AddScore 10000
    RotateLights
End Sub

Sub TopLBumper006_Hit
    If Tilted Then Exit Sub
    PlaySoundAt "fx_passive_bumper", LightTopLBumper003
    DOF 514, DOFPulse
    ' añade algunos puntos
    AddScore 10000
    RotateLights
End Sub

'**************
'   Dianas
'**************

Sub Target001_Hit
    If Tilted Then Exit Sub
    PlaySoundAt SoundFXDOF("fx_target", 709, DOFPulse, DOFTargets), target001
    DOF 504, DOFPulse
    If Epileptikdof = True Then DOF 530, DOFPulse End If
    ' añade algunos puntos
    AddScore 10000
    RotateWheel 5
    If li013.State Then TurnOnBumpers
    If li014.State Then TurOffBumpers
End Sub

Sub Target002_Hit
    If Tilted Then Exit Sub
    PlaySoundAt SoundFXDOF("fx_target", 710, DOFPulse, DOFTargets), target002
    DOF 505, DOFPulse
    If Epileptikdof = True Then DOF 531, DOFPulse End If
    ' añade algunos puntos
    AddScore 10000
    RotateWheel 5
    If li011.State Then TurnOnBumpers
    If li012.State Then TurOffBumpers
End Sub

Sub TurnOnBumpers
    LightBumper001.State = 1
    LightBumper002.State = 1
    ls1.State = 1
    ls2.State = 1
    ls3.State = 1
    ls4.State = 1
    rs1.State = 1
    rs2.State = 1
    rs3.State = 1
    rs4.State = 1
End Sub

Sub TurOffBumpers
    LightBumper001.State = 0
    LightBumper002.State = 0
    ls1.State = 0
    ls2.State = 0
    ls3.State = 0
    ls4.State = 0
    rs1.State = 0
    rs2.State = 0
    rs3.State = 0
    rs4.State = 0
End Sub

'*****************
'     Pasillos
'*****************

Sub Trigger003_Hit
    If Tilted Then Exit Sub
    PlaySoundAt "fx_sensor", Trigger003
    DOF 507, DOFPulse
    If Epileptikdof = True Then DOF 533, DOFPulse End If
    AddScore 100000
End Sub

Sub Trigger004_Hit
    If Tilted Then Exit Sub
    PlaySoundAt "fx_sensor", Trigger004
    DOF 506, DOFPulse
    If Epileptikdof = True Then DOF 532, DOFPulse End If
    AddScore 100000
End Sub

Sub Trigger006_Hit
    If Tilted Then Exit Sub
    PlaySoundAt "fx_sensor", Trigger006
    DOF 507, DOFPulse
    If Epileptikdof = True Then DOF 533, DOFPulse End If
    AddScore 100000
End Sub

Sub Trigger007_Hit
    If Tilted Then Exit Sub
    PlaySoundAt "fx_sensor", Trigger007
    DOF 506, DOFPulse
    If Epileptikdof = True Then DOF 532, DOFPulse End If
    AddScore 100000
End Sub

'************************
'       Botones
'************************

Sub Trigger001_Hit
    If Tilted Then Exit Sub
    PlaySoundAt "fx_sensor", Trigger001
    DOF 508, DOFPulse
    If Epileptikdof = True Then DOF 534, DOFPulse End If
    AddScore 10000
    RotateLights
    RotateWheel 1
End Sub

Sub Trigger002_Hit
    If Tilted Then Exit Sub
    PlaySoundAt "fx_sensor", Trigger002
    DOF 509, DOFPulse
    AddScore 10000
    RotateLights
    RotateWheel 1
End Sub

Sub Trigger005_Hit
    If Tilted Then Exit Sub
    PlaySoundAt "fx_sensor", Trigger005
    DOF 510, DOFPulse
    AddScore 10000
    RotateLights
    RotateWheel 1
End Sub

'************************
'       Agujeros
'************************

Sub Kicker001_Hit 'centro
    PlaySoundAt "fx_kicker_enter", Kicker001
    If NOT Tilted Then
        AddScore 10000
        RotateWheel 10
    End If
    vpmtimer.addtimer 1000, "DOF 516, DOFPulse: DOF 303, DOFPulse: PlaySoundAt SoundFXDOF(""fx_kicker"",712, DOFPulse, DOFContactors),kicker001: kicker001.kick 189.5,8 + RND(1)*5 '"
End Sub

Sub Kicker002_Hit 'izquierda
    PlaySoundAt "fx_kicker_enter", Kicker002
    If NOT Tilted Then
        AddScore 10000
        RotateWheel 5 + 5 * Light001.State
    End If
    vpmtimer.addtimer 1000, "DOF 515, DOFPulse: DOF 303, DOFPulse: PlaySoundAt SoundFXDOF(""fx_kicker"",711, DOFPulse, DOFContactors),kicker002: kicker002.kick 206,8 + RND(1)*5 '"
End Sub

Sub Kicker009_Hit 'derecha
    PlaySoundAt "fx_kicker_enter", Kicker009
    If NOT Tilted Then
        AddScore 10000
        RotateWheel 5 + 5 * Light002.State
    End If
    vpmtimer.addtimer 1000, "DOF 517, DOFPulse: DOF 303, DOFPulse: PlaySoundAt SoundFXDOF(""fx_kicker"",713, DOFPulse, DOFContactors),kicker009: kicker009.kick 154,8 + RND(1)*5 '"
End Sub

'**********************
'  Diversos
'**********************

Sub RotateLights 'alterna las luces
    AlternLights = NOT AlternLights
    If AlternLights Then
        light001.State = 1
        light002.State = 0
        li011.State = 1
        li012.State = 0
        li013.State = 0
        li014.State = 1
    Else
        light001.State = 0
        light002.State = 1
        li011.State = 0
        li012.State = 1
        li013.State = 1
        li014.State = 0
    End If
End Sub

Sub RotateWheel(steps)
    WheelStep = WheelStep + steps
    WheelTimer.Enabled = 1
End Sub

Sub WheelTimer_Timer()
    if WheelStep > 0 then
        j = 201 + WheelPos / 12:If B2SOn Then Controller.B2SSetData j, 0 'STAT
        WheelPos = WheelPos + 12
        PlaySound "fx_wheel"
        'add a point after each completed swimmer.
        If WheelPos = 180 Then AddBonus 1
        If WheelPos = 360 Then
            WheelPos = 0
            AddBonus 1
        End If
        j = 201 + WheelPos / 12:If B2SOn Then Controller.B2SSetData j, 1 'STAT
        Wheel.Rotz = WheelPos
        WheelStep = WheelStep - 1
    Else
        Me.Enabled = FALSE
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
HSScore(1) = 5000000
HSScore(2) = 4500000
HSScore(3) = 4000000
HSScore(4) = 3500000
HSScore(5) = 3000000

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
'
'******************************
'     DOF lights ball entrance
'******************************
'
Sub Trigger008_Hit
    DOF 520, DOFPulse
End sub

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

'Strike / IPD No. 3363 / September, 1978 / 1 Player
'Zaccaria, of Bologna, Italy (1974-1987)
'a VPX8 table by jpsalas - 2025, version 6.0.0

Option Explicit
Randomize

Const BallSize = 50
Const BallMass = 1

On Error Resume Next
ExecuteGlobal GetTextFile("controller.vbs")
If Err Then MsgBox "You need the controller.vbs in order to run this table, available in the vp10 package"
On Error Goto 0

Dim bsTrough,dtL,dtR, x, FlipperActive

Const cGameName = "strike"

Dim VarHidden
If Table1.ShowDT = true then
    VarHidden = 1
    For each x in aReels
        x.Visible = 1
    Next
Else
    VarHidden = 0
    For each x in aReels
        x.Visible = 0
    Next
End If

if B2SOn = true then VarHidden = 1

LoadVPM "01320000","zacproto.vbs",2.0

'********************
'Standard definitions
'********************

Const UseSolenoids = 1
Const UseLamps = 1
Const UseGI = 0
Const UseSync = 0
Const HandleMech = 0

' Standard Sounds
Const SFlipperOn=""
Const SFlipperOff=""
Const SSolenoidOn = "fx_SolenoidOn"
Const SSolenoidOff = "fx_SolenoidOff"
Const SCoin = "fx_Coin"

'************
' Table init.
'************

Sub table1_Init
    vpmInit me
    With Controller
        .GameName = cGameName
        If Err Then MsgBox "Can't start Game" & cGameName & vbNewLine & Err.Description:Exit Sub
        .SplashInfoLine = "Strike - Zaccaria 1978" & vbNewLine & "VPX8 table by JPSalas 6.0.0"
        .Games(cGameName).Settings.Value("sound") = 1
        .HandleKeyboard = 0
        .ShowTitle = 0
        .ShowDMDOnly = 1
        .ShowFrame = 0
        .HandleMechanics = 0
        .Hidden = VarHidden
        .Games(cGameName).Settings.Value("rol") = 0
        .Games(cGameName).Settings.Value("synclevel")=0
        '.SetDisplayPosition 0,0,GetPlayerHWnd 'uncomment if you can't see the dmd
        On Error Resume Next
        Controller.SolMask(0) = 0
        vpmTimer.AddTimer 2000, "Controller.SolMask(0)=&Hffffffff'" 'ignore all solenoids - then add the Timer to renable all the solenoids after 2 seconds
        Controller.Run GetPlayerHWnd
        On Error Goto 0
    End With

    ' Nudging
    vpmNudge.TiltSwitch = 4
    vpmNudge.Sensitivity = 5
    vpmNudge.TiltObj = Array(Bumper001, Bumper002, Bumper003, LeftSlingshot, RightSlingshot)

    ' Trough
    Set bsTrough = New cvpmBallStack
    With bsTrough
        .InitSw 0, 6, 0, 0, 0, 0, 0, 0
        .InitKick BallRelease, 90, 4
        .InitExitSnd SoundFX("fx_ballrel", DOFContactors), SoundFX("fx_Solenoid", DOFContactors)
        .Balls = 1
    End With

    ' Drop targets
    Set dtL = New cvpmDropTarget
    With dtL
        .InitDrop Array(sw25,sw26,sw27,sw28), Array(25,26,27,28)
        .InitSnd SoundFX("fx_droptarget", DOFDropTargets), SoundFX("fx_resetdrop", DOFContactors)
        .CreateEvents "dtL"
    End With

    Set dtR = New cvpmDropTarget
    With dtR
        .InitDrop Array(sw29,sw30,sw31,sw32), Array(29,30,31,32)
        .InitSnd SoundFX("fx_droptarget", DOFDropTargets), SoundFX("fx_resetdrop", DOFContactors)
        .CreateEvents "dtR"
    End With

    vpmMapLights aLights

    ' Main Timer init
    PinMAMETimer.Interval = PinMAMEInterval
    PinMAMETimer.Enabled = 1
    LampTimer.Enabled = 1

    ' Turn on Gi
    vpmtimer.addtimer 1500, "GiOn '"
End Sub

'*********
'Solenoids
'*********
SolCallback(1)="bsTrough.SolOut"
'SolCallback(2)="vpmSolSound ""sling"","
'SolCallback(3)="vpmSolSound ""sling"","
'SolCallback(4)="vpmSolSound ""bumper"","
'SolCallback(5)="vpmSolSound ""bumper"","
'SolCallback(6)="vpmSolSound ""bumper"","
'SolCallback(7)="vpmSolSound ""knocker"","
SolCallback(8)="vpmSolSound SoundFX(""fx_Knocker"",DOFKnocker),"
SolCallback(9)="dtL.SolDropUp"
SolCallback(10)="dtR.SolDropUp"
SolCallback(17)="vpmNudge.SolGameOn"
'SolCallback(sLLFlipper)="vpmSolFlipper LeftFlipper,Nothing,"
'SolCallback(sLRFlipper)="vpmSolFlipper RightFlipper,Nothing,"

'**********
' Keys
'**********

Sub table1_KeyDown(ByVal Keycode)
    If keycode = LeftFlipperKey AND Flipperactive Then SolLFlipper 1
    If keycode = RightFlipperKey AND Flipperactive Then SolRFlipper 1
    If Keycode = AddCreditKey Then vpmTimer.PulseSw swCoin1 : If Not IsEmpty(Eval("SCoin")) Then Playsound SCoin
    If keycode = LeftTiltKey Then Nudge 90, 5:PlaySound SoundFX("fx_nudge", 0), 0, 1, -0.1, 0.25
    If keycode = RightTiltKey Then Nudge 270, 5:PlaySound SoundFX("fx_nudge", 0), 0, 1, 0.1, 0.25
    If keycode = CenterTiltKey Then Nudge 0, 6:PlaySound SoundFX("fx_nudge", 0), 0, 1, 0, 0.25
    If vpmKeyDown(keycode) Then Exit Sub
    If keycode = PlungerKey Then PlaySoundat "fx_PlungerPull", Plunger:Plunger.Pullback
End Sub

Sub table1_KeyUp(ByVal Keycode)
    If keycode = LeftFlipperKey AND Flipperactive Then SolLFlipper 0
    If keycode = RightFlipperKey AND Flipperactive Then SolRFlipper 0
    If vpmKeyUp(keycode) Then Exit Sub
    If keycode = PlungerKey Then PlaySoundAt "fx_plunger", Plunger:Plunger.Fire
End Sub

'*************************
' GI - needs new vpinmame
'*************************

Set GICallback = GetRef("GIUpdate")

Sub GIUpdate(no, Enabled)
    If Enabled Then
        GiOn
    Else
        GiOff
    End If
End Sub

Sub GiOn
    Dim bulb
    PlaySound "fx_gion"
    For each bulb in aGiLights
        bulb.State = 1
    Next
End Sub

Sub GiOff
    Dim bulb
    PlaySound "fx_gioff"
    For each bulb in aGiLights
        bulb.State = 0
    Next
End Sub

Dim GiIntensity
GiIntensity = 1               'can be used by the LUT changing to increase the GI lights when the table is darker

Sub ChangeGiIntensity(factor) 'changes the intensity scale
    Dim bulb
    For each bulb in aGiLights
        bulb.IntensityScale = GiIntensity * factor
    Next
End Sub

'*********
' Switches
'*********

' Slings
Dim LStep, RStep, RStep2

Sub LeftSlingShot_Slingshot
    PlaySoundAt SoundFX("fx_slingshot", DOFContactors), Lemk
    LeftSling004.Visible = 1
    Lemk.RotX = 26
    LStep = 0
    vpmTimer.PulseSw 15
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
    RightSling004.Visible = 1
    Remk.RotX = 26
    RStep = 0
    vpmTimer.PulseSw 16
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

' Scoring rubbers

Sub rlband005_Hit:vpmTimer.PulseSw 7:End Sub
Sub rlband007_Hit:vpmTimer.PulseSw 7:End Sub
Sub rlband009_Hit:vpmTimer.PulseSw 17:End Sub
Sub rlband010_Hit:vpmTimer.PulseSw 17:End Sub

' Bumpers
Sub Bumper001_Hit:vpmTimer.PulseSw 22:PlaySoundAt SoundFX("fx_bumper", DOFContactors), Bumper001:End Sub
Sub Bumper002_Hit:vpmTimer.PulseSw 23:PlaySoundAt SoundFX("fx_bumper", DOFContactors), Bumper002:End Sub
Sub Bumper003_Hit:vpmTimer.PulseSw 24:PlaySoundAt SoundFX("fx_bumper", DOFContactors), Bumper003:End Sub

' Drain & holes
Sub Drain_Hit:PlaysoundAt "fx_drain", Drain:bsTrough.AddBall Me:End Sub

' Rollovers
Sub sw5_Hit:Controller.Switch(5) = 1:PlaySoundAt "fx_sensor", sw5:End Sub
Sub sw5_UnHit:Controller.Switch(5) = 0:End Sub

Sub sw18_Hit:Controller.Switch(18) = 1:PlaySoundAt "fx_sensor", sw18:End Sub
Sub sw18_UnHit:Controller.Switch(18) = 0:End Sub

Sub sw10_Hit:Controller.Switch(10) = 1:PlaySoundAt "fx_sensor", sw10:End Sub
Sub sw10_UnHit:Controller.Switch(10) = 0:End Sub

Sub sw10a_Hit:Controller.Switch(10) = 1:PlaySoundAt "fx_sensor", sw10a:End Sub
Sub sw10a_UnHit:Controller.Switch(10) = 0:End Sub

Sub sw18a_Hit:Controller.Switch(18) = 1:PlaySoundAt "fx_sensor", sw18a:End Sub
Sub sw18a_UnHit:Controller.Switch(18) = 0:End Sub

Sub sw10x_Hit:Controller.Switch(10) = 1:PlaySoundAt "fx_sensor", sw10x:End Sub
Sub sw10x_UnHit:Controller.Switch(10) = 0:End Sub

Sub sw9_Hit:Controller.Switch(9) = 1:PlaySoundAt "fx_sensor", sw9:End Sub
Sub sw9_UnHit:Controller.Switch(9) = 0:End Sub


' Droptargets

'Targets
Sub sw9a_Hit:vpmTimer.PulseSw 9:PlaySoundAtBall SoundFX("fx_target", DOFTargets):End Sub
Sub sw14a_Hit:vpmTimer.PulseSw 14:PlaySoundAtBall SoundFX("fx_target", DOFTargets):End Sub
Sub sw9b_Hit:vpmTimer.PulseSw 9:PlaySoundAtBall SoundFX("fx_target", DOFTargets):End Sub
Sub sw13a_Hit:vpmTimer.PulseSw 13:PlaySoundAtBall SoundFX("fx_target", DOFTargets):End Sub

Sub sw11_Hit:vpmTimer.PulseSw 11:PlaySoundAtBall SoundFX("fx_target", DOFTargets):End Sub
Sub sw12_Hit:vpmTimer.PulseSw 12:PlaySoundAtBall SoundFX("fx_target", DOFTargets):End Sub
Sub sw20_Hit:vpmTimer.PulseSw 20:PlaySoundAtBall SoundFX("fx_target", DOFTargets):End Sub
Sub sw20a_Hit:vpmTimer.PulseSw 20:PlaySoundAtBall SoundFX("fx_target", DOFTargets):End Sub

Sub sw21_Hit:vpmTimer.PulseSw 21:PlaySoundAtBall SoundFX("fx_target", DOFTargets):End Sub

Sub sw8_Hit:vpmTimer.PulseSw 8:PlaySoundAtBall SoundFX("fx_target", DOFTargets):End Sub

'*******************
'  Flipper Subs
'*******************

'SolCallback(sLRFlipper) = "SolRFlipper"
'SolCallback(sLLFlipper) = "SolLFlipper"

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

Sub LeftFlipper_Animate:LeftFlipperTop.RotZ = LeftFlipper.CurrentAngle:End Sub
Sub RightFlipper_Animate:RightFlipperTop.RotZ = RightFlipper.CurrentAngle:End Sub

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

'************************************
'          LEDs Display
'     Based on Scapino's LEDs
'************************************

Dim Digits(11)

'Assign 7-digit output to reels
Set Digits(0) = a0
Set Digits(1) = a1
Set Digits(2) = a2
Set Digits(3) = a3
Set Digits(4) = a4
Set Digits(5) = a5
Set Digits(6) = a6
Set Digits(7) = e0
Set Digits(8) = e1
Set Digits(9) = f0
Set Digits(10) = f1


Sub LampTimer_Timer
    On Error Resume Next
    Dim ChgLED, ii, jj, chg, stat
    ChgLED = Controller.ChangedLEDs(&H0000003f, &Hffffffff)
    If Not IsEmpty(ChgLED) Then
        For ii = 0 To UBound(ChgLED)
            chg = chgLED(ii, 1):stat = chgLED(ii, 2)
            Select Case stat
                Case 0:Digits(chgLED(ii, 0) ).SetValue 0    'empty
                Case 63:Digits(chgLED(ii, 0) ).SetValue 1   '0
                Case 6:Digits(chgLED(ii, 0) ).SetValue 2    '1
                Case 91:Digits(chgLED(ii, 0) ).SetValue 3   '2
                Case 79:Digits(chgLED(ii, 0) ).SetValue 4   '3
                Case 102:Digits(chgLED(ii, 0) ).SetValue 5  '4
                Case 109:Digits(chgLED(ii, 0) ).SetValue 6  '5
                Case 124:Digits(chgLED(ii, 0) ).SetValue 7  '6
                Case 125:Digits(chgLED(ii, 0) ).SetValue 7  '6
                Case 252:Digits(chgLED(ii, 0) ).SetValue 7  '6
                Case 7:Digits(chgLED(ii, 0) ).SetValue 8    '7
                Case 127:Digits(chgLED(ii, 0) ).SetValue 9  '8
                Case 103:Digits(chgLED(ii, 0) ).SetValue 10 '9
                Case 111:Digits(chgLED(ii, 0) ).SetValue 10 '9
                Case 231:Digits(chgLED(ii, 0) ).SetValue 10 '9
                Case 128:Digits(chgLED(ii, 0) ).SetValue 0  'empty
                Case 191:Digits(chgLED(ii, 0) ).SetValue 1  '0
                Case 832:Digits(chgLED(ii, 0) ).SetValue 2  '1
                Case 896:Digits(chgLED(ii, 0) ).SetValue 2  '1
                Case 768:Digits(chgLED(ii, 0) ).SetValue 2  '1
                Case 134:Digits(chgLED(ii, 0) ).SetValue 2  '1
                Case 219:Digits(chgLED(ii, 0) ).SetValue 3  '2
                Case 207:Digits(chgLED(ii, 0) ).SetValue 4  '3
                Case 230:Digits(chgLED(ii, 0) ).SetValue 5  '4
                Case 237:Digits(chgLED(ii, 0) ).SetValue 6  '5
                Case 253:Digits(chgLED(ii, 0) ).SetValue 7  '6
                Case 135:Digits(chgLED(ii, 0) ).SetValue 8  '7
                Case 255:Digits(chgLED(ii, 0) ).SetValue 9  '8
                Case 239:Digits(chgLED(ii, 0) ).SetValue 10 '9
            End Select
        Next
    End If
End Sub

'*********************************
' Diverse Collection Hit Sounds
'*********************************

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
    Pitch = BallVel(ball) * 200
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
    PlaySound soundname, 0, Vol(ActiveBall), pan(ActiveBall), 0.2, Pitch(ActiveBall), 0, 0, AudioFade(ActiveBall)
End Sub

Function RndNbr(n) 'returns a random number between 1 and n
    Randomize timer
    RndNbr = Int((n * Rnd) + 1)
End Function

'*****************************
'   JP's VPX8 Rolling Sounds
'     early SS version
'*****************************

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
    RollingTimer.Enabled = 1
End Sub

Dim oldstate3, oldstate4
oldstate3 = 0
oldstate4 = 0

Sub RollingTimer_timer()
    'only for this table
    'check for game over lamp to activate the flippers, cheap FastFlips
    If Controller.Lamp(3) AND oldstate3 = 0 Then 'gameover
        FlipperActive = False
        oldstate3 = 1
    ElseIf Controller.Lamp(3)= 0 AND oldstate3 = 1 Then
        FlipperActive = True
        oldstate3 = 0
    End If

    If Controller.Lamp(4) AND oldstate4 = 0 Then 'tilt
        FlipperActive = False
        SolLFlipper 0
        SolRFlipper 0
        vpmNudge.SolGameOn 0
        GiOff
        oldstate4 = 1
    ElseIf Controller.Lamp(4)= 0 AND oldstate4 = 1 Then
        FlipperActive = True
        vpmNudge.SolGameOn 1
        GiOn
        oldstate4 = 0
    End If

    Dim BOT, b, speedfactorx, speedfactory
    BOT = GetBalls

    ' stop the sound of deleted balls
    For b = UBound(BOT) + 1 to tnob
        rolling(b) = False
        StopSound("fx_ballrolling" & b)
    Next

    ' exit the sub if no balls on the table
    If UBound(BOT) = lob - 1 Then Exit Sub 'there are no balls on the table

    ' play the rolling sound for each ball
    For b = lob to UBound(BOT)
        If BallVel(BOT(b) )> 1 Then
            rolling(b) = True
            PlaySound("fx_ballrolling" & b), -1, Vol(BOT(b) ), Pan(BOT(b) ), 0, Pitch(BOT(b) ), 1, 0, AudioFade(BOT(b) )
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

'*********************************
' Table Options F12 User Options
'*********************************
' Table1.Option arguments are:
' - option name, minimum value, maximum value, step between valid values, default value, unit (0=None, 1=Percent), an optional array of literal strings

Dim LUTImage

Sub Table1_OptionEvent(ByVal eventId)
    Dim x, y

    'LUT
    LutImage = Table1.Option("Select LUT", 0, 21, 1, 0, 0, Array("Normal 0", "Normal 1", "Normal 2", "Normal 3", "Normal 4", "Normal 5", "Normal 6", "Normal 7", "Normal 8", "Normal 9", "Normal 10", _
        "Warm 0", "Warm 1", "Warm 2", "Warm 3", "Warm 4", "Warm 5", "Warm 6", "Warm 7", "Warm 8", "Warm 9", "Warm 10") )
    UpdateLUT

    ' Cabinet rails
    x = Table1.Option("Cabinet Rails", 0, 1, 1, 1, 0, Array("Hide", "Show") )
    For each y in aRails:y.visible = x:next
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

' Thalamus : Exit in a clean and proper way
Sub Table1_exit
  Controller.Pause = False
  Controller.Stop
End Sub


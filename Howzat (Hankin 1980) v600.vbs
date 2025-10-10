'Dennis Lillee's Howzat! / IPD No. 3909 / 1980 / 4 Players
'A. Hankin & Company, of Australia (1978-1981) [Trade Name: Hankin]
'VPX8 by jpsalas v6.0.0. vpm script based on Tom & desktruk's script from Assie34's VP9 table

Option Explicit
Randomize

Const BallSize = 50
Const BallMass = 1

On Error Resume Next
ExecuteGlobal GetTextFile("controller.vbs")
If Err Then MsgBox "You need the controller.vbs in order to run this table, available in the vp10 package"
On Error Goto 0

Dim bsTrough, dtBank4, dtBank5, bsSaucer, x

Const cGameName = "howzat"

Dim VarHidden
If Table1.ShowDT = true then
    VarHidden = 1
    For each X in aReels
        X.Visible = 1
    Next
Else
    VarHidden = 0
    For each X in aReels
        X.Visible = 0
    Next
End If

if B2SOn = true then VarHidden = 1

LoadVPM "01210000", "Hankin.VBS", 3.1

'********************
'Standard definitions
'********************

Const UseSolenoids = 2
Const UseLamps = 1
Const UseGI = 1
Const UseSync = 0
Const HandleMech = 0

' Standard Sounds
Const SSolenoidOn = "fx_SolenoidOn"
Const SSolenoidOff = "fx_SolenoidOff"
Const SCoin = "fx_Coin"

'******************
' RealTime Updates
'******************

Sub RealTime_Timer
    RollingUpdate
    UpdateLEDs
End Sub

'**********
' Keys
'**********

Sub table1_KeyDown(ByVal Keycode)
    If keycode = LeftTiltKey Then Nudge 90, 6:PlaySound SoundFX("fx_nudge", 0), 0, 1, -0.1, 0.25
    If keycode = RightTiltKey Then Nudge 270, 6:PlaySound SoundFX("fx_nudge", 0), 0, 1, 0.1, 0.25
    If keycode = CenterTiltKey Then Nudge 0, 8:PlaySound SoundFX("fx_nudge", 0), 0, 1, 0, 0.25
    If vpmKeyDown(keycode)Then Exit Sub
    If keycode = PlungerKey Then PlaySoundat "fx_PlungerPull", Plunger:Plunger.Pullback
    If keycode = KeyRules Then Rules
End Sub

Sub table1_KeyUp(ByVal Keycode)
    If vpmKeyUp(keycode)Then Exit Sub
    If keycode = PlungerKey Then PlaySoundAt "fx_plunger", Plunger:Plunger.Fire
End Sub

'*********
'Solenoids
'*********

SolCallback(1)="vpmSolSound SoundFX(""fx_Knocker"",DOFKnocker),"
SolCallback(2)="bsSaucer.SolOut"
SolCallback(3)="dtBank4.SolDropUp"
SolCallback(4)="dtBank5.SolDropUp"
SolCallback(7)="bsTrough.SolOut"
SolCallback(19)="vpmNudge.SolGameOn"

'*******************
' Flipper Subs Rev3
'*******************

SolCallback(sLRFlipper) = "SolRFlipper"
SolCallback(sLLFlipper) = "SolLFlipper"

Sub SolLFlipper(Enabled)
    If Enabled Then
        PlaySoundAt SoundFX("fx_flipperup", DOFFlippers), LeftFlipper
        LeftFlipper.RotateToEnd
        LeftFlipper1.RotateToEnd
        LeftFlipperOn = 1
    Else
        PlaySoundAt SoundFX("fx_flipperdown", DOFFlippers), LeftFlipper
        LeftFlipper.RotateToStart
        LeftFlipper1.RotateToStart
        LeftFlipperOn = 0
    End If
End Sub

Sub SolRFlipper(Enabled)
    If Enabled Then
        PlaySoundAt SoundFX("fx_flipperup", DOFFlippers), RightFlipper
        RightFlipper.RotateToEnd
        RightFlipper1.RotateToEnd
        RightFlipperOn = 1
    Else
        PlaySoundAt SoundFX("fx_flipperdown", DOFFlippers), RightFlipper
        RightFlipper.RotateToStart
        RightFlipper1.RotateToStart
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

Sub LeftFlipper1_Collide(parm)
    PlaySound "fx_rubber_flipper", 0, parm / 60, pan(ActiveBall), 0.1, 0, 0, 0, AudioFade(ActiveBall)
End Sub

Sub RightFlipper1_Collide(parm)
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

'************************************
'          LEDs Display
'     Based on Scapino's LEDs
'************************************

Dim Digits(32)

    'Assign 6-digit output to reels
    Set Digits(0) = a0
    Set Digits(1) = a1
    Set Digits(2) = a2
    Set Digits(3) = a3
    Set Digits(4) = a4
    Set Digits(5) = a5

    Set Digits(6) = b0
    Set Digits(7) = b1
    Set Digits(8) = b2
    Set Digits(9) = b3
    Set Digits(10) = b4
    Set Digits(11) = b5

    Set Digits(12) = c0
    Set Digits(13) = c1
    Set Digits(14) = c2
    Set Digits(15) = c3
    Set Digits(16) = c4
    Set Digits(17) = c5

    Set Digits(18) = d0
    Set Digits(19) = d1
    Set Digits(20) = d2
    Set Digits(21) = d3
    Set Digits(22) = d4
    Set Digits(23) = d5

    Set Digits(24) = e0
    Set Digits(25) = e1
    Set Digits(26) = e2
    Set Digits(27) = e3

Sub UpdateLEDs
    On Error Resume Next
    Dim ChgLED, ii, jj, chg, stat
    ChgLED = Controller.ChangedLEDs(&H0000003f, &Hffffffff)
    If Not IsEmpty(ChgLED)Then
        For ii = 0 To UBound(ChgLED)
            chg = chgLED(ii, 1):stat = chgLED(ii, 2)
            Select Case stat
                Case 0:Digits(chgLED(ii, 0)).SetValue 0    'empty
                Case 63:Digits(chgLED(ii, 0)).SetValue 1   '0
                Case 6:Digits(chgLED(ii, 0)).SetValue 2    '1
                Case 91:Digits(chgLED(ii, 0)).SetValue 3   '2
                Case 79:Digits(chgLED(ii, 0)).SetValue 4   '3
                Case 102:Digits(chgLED(ii, 0)).SetValue 5  '4
                Case 109:Digits(chgLED(ii, 0)).SetValue 6  '5
                Case 124:Digits(chgLED(ii, 0)).SetValue 7  '6
                Case 125:Digits(chgLED(ii, 0)).SetValue 7  '6
                Case 252:Digits(chgLED(ii, 0)).SetValue 7  '6
                Case 7:Digits(chgLED(ii, 0)).SetValue 8    '7
                Case 127:Digits(chgLED(ii, 0)).SetValue 9  '8
                Case 103:Digits(chgLED(ii, 0)).SetValue 10 '9
                Case 111:Digits(chgLED(ii, 0)).SetValue 10 '9
                Case 231:Digits(chgLED(ii, 0)).SetValue 10 '9
                Case 128:Digits(chgLED(ii, 0)).SetValue 0  'empty
                Case 191:Digits(chgLED(ii, 0)).SetValue 1  '0
                Case 832:Digits(chgLED(ii, 0)).SetValue 2  '1
                Case 896:Digits(chgLED(ii, 0)).SetValue 2  '1
                Case 768:Digits(chgLED(ii, 0)).SetValue 2  '1
                Case 134:Digits(chgLED(ii, 0)).SetValue 2  '1
                Case 219:Digits(chgLED(ii, 0)).SetValue 3  '2
                Case 207:Digits(chgLED(ii, 0)).SetValue 4  '3
                Case 230:Digits(chgLED(ii, 0)).SetValue 5  '4
                Case 237:Digits(chgLED(ii, 0)).SetValue 6  '5
                Case 253:Digits(chgLED(ii, 0)).SetValue 7  '6
                Case 135:Digits(chgLED(ii, 0)).SetValue 8  '7
                Case 255:Digits(chgLED(ii, 0)).SetValue 9  '8
                Case 239:Digits(chgLED(ii, 0)).SetValue 10 '9
            End Select
        Next
    End If
End Sub

'************
' Table init.
'************

Sub table1_Init
    vpmInit me
    With Controller
        .GameName = cGameName
        If Err Then MsgBox "Can't start Game" & cGameName & vbNewLine & Err.Description:Exit Sub
        .SplashInfoLine = "Howzat! - Hankin 1980" & vbNewLine & "VPX table by JPSalas 5.5.0"
     .Games(cGameName).Settings.Value("sound")=1
        .HandleKeyboard = 0
        .ShowTitle = 0
        .ShowDMDOnly = 1
        .ShowFrame = 0
        .HandleMechanics = 0
        .Hidden = VarHidden
        .Games(cGameName).Settings.Value("rol") = 0
        '.SetDisplayPosition 0,0,GetPlayerHWnd 'uncomment if you can't see the dmd
        On Error Resume Next
        Controller.SolMask(0) = 0
        vpmTimer.AddTimer 2000, "Controller.SolMask(0)=&Hffffffff'" 'ignore all solenoids - then add the Timer to renable all the solenoids after 2 seconds
        Controller.Run GetPlayerHWnd
        On Error Goto 0
    End With

    ' Nudging
    vpmNudge.TiltSwitch =2
    vpmNudge.Sensitivity = 3
    vpmNudge.TiltObj = Array(Bumper001, Bumper002, Bumper003, Bumper004, LeftSlingshot, RightSlingshot)

    ' Trough
    Set bsTrough = New cvpmBallStack
    With bsTrough
        .InitSw 0,40,0,0,0,0,0,0
        .InitKick BallRelease, 90, 4
        .InitExitSnd SoundFX("fx_ballrel", DOFContactors), SoundFX("fx_Solenoid", DOFContactors)
        .Balls = 1
        .IsTrough = 1
    End With

    ' Saucer
    Set bsSaucer = New cvpmBallStack
    With bsSaucer
        .InitSaucer sw32, 32, 225, 16
        .InitExitSnd SoundFX("fx_kicker", DOFContactors), SoundFX("fx_Solenoid", DOFContactors)
        .KickForceVar = 2
        .KickAngleVar = 2
    End With

    ' Drop targets
    Set dtBank4 = New cvpmDropTarget
    With dtBank4
        .InitDrop Array(sw8,sw7,sw6,sw5), Array(8,7,6,5)
        .InitSnd SoundFX("fx_droptarget", DOFDropTargets), SoundFX("fx_resetdrop", DOFContactors)
        .CreateEvents "dtBank4"
    End With

    Set dtBank5 = New cvpmDropTarget
    With dtBank5
        .InitDrop Array(sw15,sw14,sw13,sw12,sw11), Array(15,14,13,12,11)
        .InitSnd SoundFX("fx_droptarget", DOFDropTargets), SoundFX("fx_resetdrop", DOFContactors)
        .CreateEvents "dtBank5"
    End With

    vpmMapLights aLights

    ' Main Timer init
    PinMAMETimer.Interval = PinMAMEInterval
    PinMAMETimer.Enabled = 1
    RealTime.Enabled = 1

    ' Turn on Gi
    vpmtimer.addtimer 1500, "GiOn '"
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
  PlaySound"fx_gion"
    For each bulb in aGiLights
        bulb.State = 1
    Next
End Sub

Sub GiOff
    Dim bulb
  PlaySound"fx_gioff"
    For each bulb in aGiLights
        bulb.State = 0
    Next
End Sub

'*********
' Switches
'*********

' Slings
Dim LStep, RStep

Sub LeftSlingShot_Slingshot
    PlaySoundAt SoundFX("fx_slingshot", DOFContactors), Lemk
    'DOF 101, DOFPulse
    LeftSling004.Visible = 1
    Lemk.RotX = 26
    LStep = 0
    vpmTimer.PulseSw 37
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
    vpmTimer.PulseSw 36
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
Sub rlband008_Hit: vpmTimer.PulseSw 1: End Sub

' Bumpers
Sub Bumper001_Hit:vpmTimer.PulseSw 33:PlaySoundAt SoundFX("fx_bumper", DOFContactors), Bumper001:End Sub
Sub Bumper002_Hit:vpmTimer.PulseSw 35:PlaySoundAt SoundFX("fx_bumper", DOFContactors), Bumper002:End Sub
Sub Bumper003_Hit:vpmTimer.PulseSw 34:PlaySoundAt SoundFX("fx_bumper", DOFContactors), Bumper003:End Sub
Sub Bumper004_Hit:vpmTimer.PulseSw 38:PlaySoundAt SoundFX("fx_bumper", DOFContactors), Bumper004:End Sub

' Drain & holes
Sub Drain_Hit:PlaysoundAt "fx_drain", Drain:bsTrough.AddBall Me:End Sub
Sub sw32_Hit:PlaysoundAt "fx_kicker_enter", sw32:bsSaucer.AddBall 0:End Sub

' Rollovers
Sub sw21_Hit:Controller.Switch(21) = 1:PlaySoundAt "fx_sensor", sw21:End Sub
Sub sw21_UnHit:Controller.Switch(21) = 0:End Sub

Sub sw28_Hit:Controller.Switch(28) = 1:PlaySoundAt "fx_sensor", sw28:End Sub
Sub sw28_UnHit:Controller.Switch(28) = 0:End Sub

Sub sw20_Hit:Controller.Switch(20) = 1:PlaySoundAt "fx_sensor", sw20:End Sub
Sub sw20_UnHit:Controller.Switch(20) = 0:End Sub

Sub sw4_Hit:Controller.Switch(4) = 1:PlaySoundAt "fx_sensor", sw4:End Sub
Sub sw4_UnHit:Controller.Switch(4) = 0:End Sub

Sub sw17_Hit:Controller.Switch(17) = 1:PlaySoundAt "fx_sensor", sw17:End Sub
Sub sw17_UnHit:Controller.Switch(17) = 0:End Sub

Sub sw18_Hit:Controller.Switch(18) = 1:PlaySoundAt "fx_sensor", sw18:End Sub
Sub sw18_UnHit:Controller.Switch(18) = 0:End Sub

Sub sw19_Hit:Controller.Switch(19) = 1:PlaySoundAt "fx_sensor", sw19:End Sub
Sub sw19_UnHit:Controller.Switch(19) = 0:End Sub

Sub sw22_Hit:Controller.Switch(22) = 1:PlaySoundAt "fx_sensor", sw22:End Sub
Sub sw22_UnHit:Controller.Switch(22) = 0:End Sub

Sub sw29_Hit:Controller.Switch(29) = 1:PlaySoundAt "fx_sensor", sw29:End Sub
Sub sw29_UnHit:Controller.Switch(29) = 0:End Sub

Sub sw23_Hit:Controller.Switch(23) = 1:PlaySoundAt "fx_sensor", sw23:End Sub
Sub sw23_UnHit:Controller.Switch(23) = 0:End Sub

Sub sw39_Hit:Controller.Switch(39) = 1:PlaySoundAt "fx_sensor", sw39:End Sub
Sub sw39_UnHit:Controller.Switch(39) = 0:End Sub

Sub sw31_Hit:Controller.Switch(31) = 1:PlaySoundAt "fx_sensor", sw31:End Sub
Sub sw31_UnHit:Controller.Switch(31) = 0:End Sub

'Spinners

Sub sw25_Spin():vpmTimer.PulseSw 25:PlaySoundAt "fx_spinner", sw25:End Sub
Sub sw26_Spin():vpmTimer.PulseSw 26:PlaySoundAt "fx_spinner", sw26:End Sub

' Droptargets

'Targets
Sub sw30_Hit:vpmTimer.PulseSw 20:PlaySoundAtBall SoundFX("fx_target", DOFTargets):End Sub

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
Sub aTargets_Hit(idx):PlaySound SoundFX("fx_target",DOFTargets), 0, Vol(ActiveBall), pan(ActiveBall), 0.2, Pitch(ActiveBall) * 10, 0, 0, AudioFade(ActiveBall):End Sub
Sub aRollovers_Hit(idx):PlaySoundAt "fx_sensor", aRollovers(idx):End Sub

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
        aBallShadow(b).Height = BOT(b).Z -Ballsize/2 + 1

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

'*********************************
' Table Options F12 User Options
'*********************************
' Table1.Option arguments are:
' - option name, minimum value, maximum value, step between valid values, default value, unit (0=None, 1=Percent), an optional array of literal strings

Dim LUTImage, Balls

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
    If x = 1 Then Balls = 5 Else Balls = 3
    If Balls=3 Then
    Controller.Dip(0)=(0*1 + 0*2 + 0*4 + 0*8 + 1*16 + 1*32 + 1*64 + 1*128) '01-08
    Controller.Dip(1)=(0*1 + 1*2 + 0*4 + 0*8 + 0*16 + 0*32 + 0*64 + 0*128) '09-16
    Controller.Dip(2)=(0*1 + 0*2 + 0*4 + 0*8 + 0*16 + 1*32 + 0*64 + 0*128) '17-24
    Controller.Dip(3)=(0*1 + 0*2 + 0*4 + 0*8 + 0*16 + 0*32 + 0*64 + 0*128) '25-32
    End If
    If Balls=5 Then
    Controller.Dip(0)=(0*1 + 0*2 + 0*4 + 0*8 + 1*16 + 1*32 + 1*64 + 1*128) '01-08
    Controller.Dip(1)=(0*1 + 1*2 + 0*4 + 0*8 + 0*16 + 0*32 + 0*64 + 1*128) '09-16
    Controller.Dip(2)=(0*1 + 0*2 + 0*4 + 0*8 + 0*16 + 1*32 + 0*64 + 0*128) '17-24
    Controller.Dip(3)=(0*1 + 0*2 + 0*4 + 0*8 + 0*16 + 0*32 + 0*64 + 0*128) '25-32
    End If
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


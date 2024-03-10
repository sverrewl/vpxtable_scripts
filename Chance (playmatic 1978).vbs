' Chance / Playmatic / IPD No. 491 / September, 1978 / 4 Players
' VPX - version by JPSalas 2020, version 1.0.0. Graphics by Pedator
' Script from on destruks & TAB table:
' Playmatic System 1 - each system is different - over 9 years Playmatic created 4 known pinball machine types.
' Playmatic System 1 is known to be used in solid state playmatic games Last Lap, Chance, Big Town, and Space Gambler
' Playmatic System 1 switches are -10 from manual number

Option Explicit
Randomize

Const BallSize = 50
Const BallMass = 1

On Error Resume Next
ExecuteGlobal GetTextFile("controller.vbs")
If Err Then MsgBox "You need the Controller.vbs file in order to run this table (installed with the VPX package in the scripts folder)"
On Error Goto 0

LoadVPM "01120100", "play1.vbs", 3.1

Const UseSolenoids = 1
Const UseLamps = 1
Const UseSync = 0
Const HandleMech = 0
Const SSolenoidOn = "fx_solenoidon"
Const SSolenoidOff = "fx_solenoidoff"
Const SCoin = "fx_Coin"

Dim VarHidden:VarHidden = Table1.ShowDT
if B2SOn = true then VarHidden = 1

Dim x

' Remove desktop items in FS mode
If Table1.ShowDT then
    For each x in aReels
        x.Visible = 1
    Next
Else
    For each x in aReels
        x.Visible = 0
    Next
End If

'************
' Table init.
'************

Dim bsTrough, bsTopSaucer, Flipperactive
Const cGameName = "chance"

Sub Table1_Init

' Thalamus : Was missing 'vpminit me'
  vpminit me

    With Controller
        .GameName = cGameName
        If Err Then MsgBox "Can't start Game" & cGameName & vbNewLine & Err.Description:Exit Sub
        .SplashInfoLine = "Chance, Playmatic 1978" & vbnewline & "Table by jpsalas"
        .Games(cGameName).Settings.Value("rol") = 0   '1= rotated display, 0= normal
        .Games(cGameName).Settings.Value("sound") = 1 '1 enabled rom sound
        .HandleMechanics = 0
        .ShowDMDOnly = 1
        .ShowFrame = 0
        .ShowTitle = 0
        .Hidden = VarHidden
        On Error Resume Next
        .SolMask(0) = 0
        vpmTimer.AddTimer 2000, "Controller.SolMask(0)=&Hffffffff'" 'ignore all solenoids - then add the Timer to renable all the solenoids after 2 seconds
        .Run GetPlayerHWnd
        If Err Then MsgBox Err.Description
        On Error Goto 0
    End With

    vpmNudge.TiltSwitch = 5
    vpmNudge.Sensitivity = 5
    vpmNudge.TiltObj = Array(bumper1, bumper2, bumper3, LeftSlingshot, RightSlingshot)

    Set bsTrough = New cvpmBallStack
    bsTrough.InitNoTrough BallRelease, 3, 90, 8
    bsTrough.InitExitSnd "fx_ballrelease", "fx_solenoidon"

    Set bsTopSaucer = New cvpmBallStack
    bsTopSaucer.InitSaucer Kicker1, 25, 197, 14
    bsTopSaucer.InitExitSnd "fx_kicker", "fx_solenoion"

    vpmMapLights aLights

    ' turn the RealTimeUpdates timer
    RealTimeUpdates.Enabled = 1

    SolCenterPost 1

    Flipperactive = False
End Sub

'******************
' RealTime Updates
'******************

'Set MotorCallback = GetRef("RealTimeUpdates") 'use this for a faster update

Dim OldState1

OldState1 = 0

Sub RealTimeUpdates_Timer
    RollingUpdate
    UpdateLeds
    LeftFlipperTop.RotZ = LeftFlipper.CurrentAngle
    RightFlipperTop.RotZ = RightFlipper.CurrentAngle

    ' Fastflippers (a sort of :) )
    If Controller.Lamp(14)Then 'tilted deactivate the table objects
        Flipperactive = False
        LeftFlipper.RotateToStart
        RightFlipper.RotateToStart
        GiOff
        Bumper1.Force = 0
        Bumper2.Force = 0
        Bumper3.Force = 0
        LeftSlingshot.Disabled = 1
        RightSlingshot.Disabled = 1
        If bsTopSaucer.Balls Then
            vpmtimer.addtimer 200, "kicker1.kick 197, 14 '"
        End If
        SolCenterPost 0
        OldState1 = 0
    Else
        OldState1 = 1
    End If

    If Controller.Lamp(1) <> OldState1 Then 'Game Over light
        If Controller.Lamp(1)Then
            VpmNudge.SolGameOn False
            OldState1 = True
            GiOff
            Flipperactive = False
            LeftFlipper.RotateToStart
            RightFlipper.RotateToStart
        Else
            VpmNudge.SolGameOn True
            OldState1 = False
            GiOn
            Flipperactive = True
            ' activate bumpers and slings in case of tilt
            Bumper1.Force = 10
            Bumper2.Force = 10
            Bumper3.Force = 10
            LeftSlingshot.Disabled = 0
            RightSlingshot.Disabled = 0
        End If
    End If

    ' bumpers
    LightBumper002.State = ABS(Controller.Solenoid(4))
    LightBumper003.State = ABS(Controller.Solenoid(4))
End Sub

'**********
' Keys
'**********

Sub table1_KeyDown(ByVal Keycode)
    If keycode = LeftTiltKey Then Nudge 90, 5:PlaySound SoundFX("fx_nudge", 0), 0, 1, -0.1, 0.25
    If keycode = RightTiltKey Then Nudge 270, 5:PlaySound SoundFX("fx_nudge", 0), 0, 1, 0.1, 0.25
    If keycode = CenterTiltKey Then Nudge 0, 6:PlaySound SoundFX("fx_nudge", 0), 0, 1, 0, 0.25
    If keycode = LeftFlipperKey AND flipperactive Then SolLFlipper 1
    If keycode = RightFlipperKey AND flipperactive Then SolRFlipper 1
    If KeyDownHandler(KeyCode)Then Exit Sub
    If keycode = PlungerKey Then PlaySoundAt "fx_PlungerPull", Plunger:Plunger.Pullback
End Sub

Sub table1_KeyUp(ByVal Keycode)
    If keycode = LeftFlipperKey AND flipperactive Then SolLFlipper 0
    If keycode = RightFlipperKey AND flipperactive Then SolRFlipper 0
    If KeyUpHandler(KeyCode)Then Exit Sub
    If keycode = PlungerKey Then PlaySoundAt "fx_plunger", plunger:Plunger.Fire
End Sub

'********************
'     Flippers
'********************

'SolCallback(sLRFlipper) = "SolRFlipper"
'SolCallback(sLLFlipper) = "SolLFlipper"

Sub SolLFlipper(Enabled)
    If Enabled Then
        PlaySoundAt SoundFX("fx_flipperup", DOFFlippers), LeftFlipper
        LeftFlipper.EOSTorque = 0.85
        LeftFlipper.RotateToEnd
    Else
        If flipperactive Then PlaySoundAt SoundFX("fx_flipperdown", DOFContactors), LeftFlipper
        LeftFlipper.EOSTorque = .25
        LeftFlipper.RotateToStart
    End If
End Sub

Sub SolRFlipper(Enabled)
    If Enabled Then
        PlaySoundAt SoundFX("fx_flipperup", DOFFlippers), RightFlipper
        RightFlipper.EOSTorque = 0.85
        RightFlipper.RotateToEnd
    Else
        If flipperactive Then PlaySoundAt SoundFX("fx_flipperdown", DOFContactors), RightFlipper
        RightFlipper.EOSTorque = 0.25
        RightFlipper.RotateToStart
    End If
End Sub

Sub LeftFlipper_Collide(parm)
    PlaySound "fx_rubber_flipper", 0, parm / 60, pan(ActiveBall), 0.2, 0, 0, 0, AudioFade(ActiveBall)
End Sub

Sub RightFlipper_Collide(parm)
    PlaySound "fx_rubber_flipper", 0, parm / 60, pan(ActiveBall), 0.2, 0, 0, 0, AudioFade(ActiveBall)
End Sub

'*********
'Solenoids
'*********

SolCallback(1) = "RaiseRight"
SolCallback(2) = "RaiseLeft"
SolCallback(7) = "bsTrough.SolOut"
SolCallBack(4) = "SolCenterPost"
SolCallback(8) = "vpmSolSound SoundFX(""fx_Knocker"",DOFKnocker),"
SolCallBack(17) = "SolTopSaucer"

Sub SolCenterPost(Enabled)
    If Not Controller.Lamp(14)Then
        Post.IsDropped = Enabled
        postlight.State = ABS(Enabled)-1
        postlight2.State = ABS(Enabled)-1
    End If
End Sub

Sub RaiseLeft(Enabled)
    If Enabled Then
        PlaySoundat "fx_resetdrop", sw26
        LNum = 0
        If RNum = 0 Then Controller.Switch(45) = 0
        Controller.Switch(16) = 0
        Controller.Switch(17) = 0
        Controller.Switch(18) = 0
        sw26.IsDropped = 0
        sw27a.IsDropped = 0
        sw27b.IsDropped = 0
        sw28a.IsDropped = 0
        sw28b.IsDropped = 0
    End If
End Sub

Sub RaiseRight(Enabled)
    If Enabled Then
        PlaySoundat "fx_resetdrop", sw23
        RNum = 0
        If LNum = 0 Then Controller.Switch(45) = 0
        Controller.Switch(11) = 0
        Controller.Switch(12) = 0
        Controller.Switch(13) = 0
        sw21a.IsDropped = 0
        sw21b.IsDropped = 0
        sw22a.IsDropped = 0
        sw22b.IsDropped = 0
        sw23.IsDropped = 0
    End If
End Sub

Sub SolTopSaucer(Enabled)
    If Enabled Then
        bsTopSaucer.ExitSol_On
        iron.rotx = -20
        vpmtimer.addtimer 300, "iron.rotx = 0 '"
    End If
End Sub

'*********
' Switches
'*********

' Slings
Dim LStep, RStep

Sub LeftSlingShot_Slingshot
    PlaySoundAt SoundFX("fx_slingshot", DOFContactors), Lemk
    DOF 101, DOFPulse
    LeftSling4.Visible = 1
    Lemk.RotX = 26
    LStep = 0
    vpmTimer.PulseSw 33
    LeftSlingShot.TimerEnabled = 1
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
    PlaySoundAt SoundFX("fx_slingshot", DOFContactors), Remk
    DOF 102, DOFPulse
    RightSling4.Visible = 1
    Remk.RotX = 26
    RStep = 0
    vpmTimer.PulseSw 33
    RightSlingShot.TimerEnabled = 1
End Sub

Sub RightSlingShot_Timer
    Select Case RStep
        Case 1:RightSLing4.Visible = 0:RightSLing3.Visible = 1:Remk.RotX = 14
        Case 2:RightSLing3.Visible = 0:RightSLing2.Visible = 1:Remk.RotX = 2
        Case 3:RightSLing2.Visible = 0:Remk.RotX = -10:RightSlingShot.TimerEnabled = 0
    End Select
    RStep = RStep + 1
End Sub

'Rubbers
Sub RubberBand006_Hit:vpmTimer.PulseSw 33:End Sub
Sub RubberBand005_Hit:vpmTimer.PulseSw 33:End Sub
Sub RubberBand011_Hit:vpmTimer.PulseSw 33:End Sub
Sub RubberBand014_Hit:vpmTimer.PulseSw 33:End Sub

Sub RubberBand016_Hit:vpmTimer.PulseSw 44:End Sub
Sub RubberBand015_Hit:vpmTimer.PulseSw 44:End Sub

Sub RubberBand018_Hit:vpmTimer.PulseSw 45:End Sub
Sub RubberBand017_Hit:vpmTimer.PulseSw 45:End Sub

'Drain & Kickers
Sub Drain_Hit:PlaySoundAt "fx_drain", Drain:bsTrough.AddBall Me:End Sub
Sub Kicker1_Hit:PlaySoundAt "fx_kicker_enter", kicker1:bsTopSaucer.AddBall 0:End Sub

'Lanes

Sub SW46a_Hit:Controller.Switch(36) = 1:PlaySoundAt "fx_sensor", sw46a:End Sub '46
Sub SW46a_unHit:Controller.Switch(36) = 0:End Sub
Sub SW46b_Hit:Controller.Switch(36) = 1:PlaySoundAt "fx_sensor", sw46b:End Sub
Sub SW46b_unHit:Controller.Switch(36) = 0:End Sub
Sub SW47a_Hit:Controller.Switch(37) = 1:PlaySoundAt "fx_sensor", sw47a:End Sub
Sub SW47a_unHit:Controller.Switch(37) = 0:End Sub
Sub SW47b_Hit:Controller.Switch(37) = 1:PlaySoundAt "fx_sensor", sw47b:End Sub
Sub SW47b_unHit:Controller.Switch(37) = 0:End Sub
Sub SW48_Hit:Controller.Switch(38) = 1:PlaySoundAt "fx_sensor", sw48:End Sub   '48
Sub SW48_unHit:Controller.Switch(38) = 0:End Sub
Sub SW51a_Hit:Controller.Switch(41) = 1:PlaySoundAt "fx_sensor", sw51a:End Sub '51
Sub SW51a_unHit:Controller.Switch(41) = 0:End Sub
Sub SW51b_Hit:Controller.Switch(41) = 1:PlaySoundAt "fx_sensor", sw51b:End Sub
Sub SW51b_unHit:Controller.Switch(41) = 0:End Sub
Sub SW52a_Hit:Controller.Switch(42) = 1:PlaySoundAt "fx_sensor", sw52a:End Sub '52
Sub SW52a_unHit:Controller.Switch(42) = 0:End Sub
Sub SW52b_Hit:Controller.Switch(42) = 1:PlaySoundAt "fx_sensor", sw52b:End Sub
Sub SW52b_unHit:Controller.Switch(42) = 0:End Sub
Sub SW53_Hit:Controller.Switch(43) = 1:PlaySoundAt "fx_sensor", sw53:End Sub   '53
Sub SW53_unHit:Controller.Switch(43) = 0:End Sub
Sub SW54a_Hit:Controller.Switch(44) = 1:PlaySoundAt "fx_sensor", sw54a:End Sub '54
Sub SW54a_unHit:Controller.Switch(44) = 0:End Sub
Sub SW54b_Hit:Controller.Switch(44) = 1:PlaySoundAt "fx_sensor", sw54b:End Sub
Sub SW54b_unHit:Controller.Switch(44) = 0:End Sub
Sub SW57_Hit:Controller.Switch(47) = 1:PlaySoundAt "fx_sensor", sw57:End Sub '57
Sub SW57_unHit:Controller.Switch(47) = 0:End Sub
Sub SW58_Hit:Controller.Switch(48) = 1:PlaySoundAt "fx_sensor", sw58:End Sub '58
Sub SW58_unHit:Controller.Switch(48) = 0:End Sub

'Droptargets

Dim LNum, RNum
LNum = 0:RNum = 0

Sub sw21a_Dropped '21
    PlaySoundat "fx_droptarget", sw21a
    If sw21b.IsDropped Then
        Controller.Switch(11) = 1
    End If
    CheckRNum
End Sub
Sub sw21b_Dropped
    PlaySoundat "fx_droptarget", sw21b
    If sw21a.IsDropped Then
        Controller.Switch(11) = 1
    End If
    CheckRNum
End Sub
Sub sw22a_Dropped '22
    PlaySoundat "fx_droptarget", sw22a
    If sw22b.IsDropped Then
        Controller.Switch(12) = 1
    End If
    CheckRNum
End Sub
Sub sw22b_Dropped
    PlaySoundat "fx_droptarget", sw22b
    If sw22a.IsDropped Then
        Controller.Switch(12) = 1
    End If
    CheckRNum
End Sub
Sub sw23_Dropped '23
    PlaySoundat "fx_droptarget", sw23
    Controller.Switch(13) = 1
    CheckRNum
End Sub
Sub sw26_Dropped '26
    PlaySoundat "fx_droptarget", sw26
    Controller.Switch(16) = 1
    CheckLNum
End Sub
Sub sw27a_Dropped '27
    PlaySoundat "fx_droptarget", sw27a
    If sw27b.IsDropped Then
        Controller.Switch(17) = 1
    End If
    CheckLNum
End Sub
Sub sw27b_Dropped
    PlaySoundat "fx_droptarget", sw27b
    If sw27a.IsDropped Then
        Controller.Switch(17) = 1
    End If
    CheckLNum
End Sub
Sub sw28a_Dropped '28
    PlaySoundat "fx_droptarget", sw28a
    If sw28b.IsDropped Then
        Controller.Switch(18) = 1
    End If
    CheckLNum
End Sub
Sub sw28b_Dropped
    PlaySoundat "fx_droptarget", sw28b
    If sw28a.IsDropped Then
        Controller.Switch(18) = 1
    End If
    CheckLNum
End Sub

Sub CheckLNum
    LNum = LNum + 1
    If LNum = 5 Then Controller.Switch(45) = 1
End Sub

Sub CheckRNum
    RNum = RNum + 1
    If RNum = 5 Then Controller.Switch(45) = 1
End Sub

'Targets
Sub sw34a_Hit:vpmTimer.PulseSw 23:vpmTimer.PulseSw(24):PlaySoundAtBall SoundFX("fx_target", DOFTargets):End Sub
Sub sw34b_Hit:vpmTimer.PulseSw 23:vpmTimer.PulseSw(24):PlaySoundAtBall SoundFX("fx_target", DOFTargets):End Sub
Sub sw37_Hit:vpmTimer.PulseSw 27:PlaySoundAtBall SoundFX("fx_target", DOFTargets):End Sub
Sub sw38a_Hit:vpmTimer.PulseSw 28:PlaySoundAtBall SoundFX("fx_target", DOFTargets):End Sub
Sub sw38b_Hit:vpmTimer.PulseSw 28:PlaySoundAtBall SoundFX("fx_target", DOFTargets):End Sub

'Bumpers
Sub bumper1_Hit:vpmTimer.PulseSw 37:PlaySoundAt SoundFX("fx_bumper", DOFContactors), bumper1:End Sub
Sub bumper2_Hit:vpmTimer.PulseSw 26:PlaySoundAt SoundFX("fx_bumper", DOFContactors), bumper2:End Sub
Sub bumper3_Hit:vpmTimer.PulseSw 26:PlaySoundAt SoundFX("fx_bumper", DOFContactors), bumper3:End Sub

'Extra Lights

Set LampCallback = GetRef("UpdateMultipleLamps")

Sub UpdateMultipleLamps
    bip1.Setvalue ABS(Controller.Lamp(2))
    bip2.Setvalue ABS(Controller.Lamp(3))
    bip3.Setvalue ABS(Controller.Lamp(4))
    bip4.Setvalue ABS(Controller.Lamp(5))
    bip5.Setvalue ABS(Controller.Lamp(6))

    r6.SetValue ABS(Controller.Lamp(8))
    r7.SetValue ABS(Controller.Lamp(40))
    r8.SetValue ABS(Controller.Lamp(39))
    r9.SetValue ABS(Controller.Lamp(38))
    r1.SetValue ABS(Controller.Lamp(37))
    r2.SetValue ABS(Controller.Lamp(36))
    r3.SetValue ABS(Controller.Lamp(35))
    r4.SetValue ABS(Controller.Lamp(34))
    r5.SetValue ABS(Controller.Lamp(33))

    rhigh.SetValue ABS(Controller.Lamp(33))
    TiltReel.SetValue ABS(Controller.Lamp(14))
    GameOverR.SetValue ABS(Controller.Lamp(1))

    r2x.SetValue ABS(Controller.Lamp(44))
    r3x.SetValue ABS(Controller.Lamp(45))
End Sub

'******************************
' Diverse Collection Hit Sounds
'******************************

Sub aMetals_Hit(idx):PlaySoundAtBall "fx_MetalHit":End Sub
Sub aRubber_Bands_Hit(idx):PlaySoundAtBall "fx_rubber_band":End Sub
Sub aRubber_Posts_Hit(idx):PlaySoundAtBall "fx_rubber_post":End Sub
Sub aRubber_Pins_Hit(idx):PlaySoundAtBall "fx_rubber_pin":End Sub
Sub aPlastics_Hit(idx):PlaySoundAtBall "fx_PlasticHit":End Sub
Sub aGates_Hit(idx):PlaySoundAtBall "fx_Gate":End Sub
Sub aWoods_Hit(idx):PlaySoundAtBall "fx_Woodhit":End Sub

' *********************************************************************
'                      Supporting Ball & Sound Functions
' *********************************************************************

Function Vol(ball) ' Calculates the Volume of the sound based on the ball speed
    Vol = Csng(BallVel(ball) ^2 / 2000)
End Function

Function Pan(ball) ' Calculates the pan for a ball based on the X position on the table. "table1" is the name of the table
    Dim tmp
    tmp = ball.x * 2 / table1.width-1
    If tmp> 0 Then
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
    tmp = ball.y * 2 / Table1.height-1
    If tmp> 0 Then
        AudioFade = Csng(tmp ^10)
    Else
        AudioFade = Csng(-((- tmp) ^10))
    End If
End Function

Sub PlaySoundAt(soundname, tableobj) 'play sound at X and Y position of an object, mostly bumpers, flippers and other fast objects
    PlaySound soundname, 0, 1, Pan(tableobj), 0.06, 0, 0, 0, AudioFade(tableobj)
End Sub

Sub PlaySoundAtBall(soundname) ' play a sound at the ball position, like rubbers, targets, metals, plastics
    PlaySound soundname, 0, Vol(ActiveBall), pan(ActiveBall), 0.2, 0, 0, 0, AudioFade(ActiveBall)
End Sub

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

        If BallVel(BOT(b))> 1 Then
            If BOT(b).z <30 Then
                ballpitch = Pitch(BOT(b))
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
        If BOT(b).VelZ <-1 and BOT(b).z <55 and BOT(b).z> 27 Then 'height adjust for ball drop sounds
            PlaySound "fx_balldrop", 0, ABS(BOT(b).velz) / 17, Pan(BOT(b)), 0, Pitch(BOT(b)), 1, 0, AudioFade(BOT(b))
        End If
    Next
End Sub

'**********************
' Ball Collision Sound
'**********************

Sub OnBallBallCollision(ball1, ball2, velocity)
    PlaySound("fx_collide"), 0, Csng(velocity) ^2 / 2000, Pan(ball1), 0, Pitch(ball1), 0, 0, AudioFade(ball1)
End Sub

'************************************
'          LEDs Display
'     Based on Scapino's LEDs
'************************************

Dim Digits(32)
Dim Patterns(11)
Dim Patterns2(11)

Patterns(0) = 0     'empty
Patterns(1) = 63    '0
Patterns(2) = 6     '1
Patterns(3) = 91    '2
Patterns(4) = 79    '3
Patterns(5) = 102   '4
Patterns(6) = 109   '5
Patterns(7) = 125   '6
Patterns(8) = 7     '7
Patterns(9) = 127   '8
Patterns(10) = 111  '9

Patterns2(0) = 128  'empty
Patterns2(1) = 191  '0
Patterns2(2) = 134  '1
Patterns2(3) = 219  '2
Patterns2(4) = 207  '3
Patterns2(5) = 230  '4
Patterns2(6) = 237  '5
Patterns2(7) = 253  '6
Patterns2(8) = 135  '7
Patterns2(9) = 255  '8
Patterns2(10) = 239 '9

'Assign 7-digit output to reels
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
Set Digits(28) = e4

Sub UpdateLeds
    On Error Resume Next
    Dim ChgLED, ii, jj, chg, stat
    ChgLED = Controller.ChangedLEDs(&HFF, &HFFFF)
    If Not IsEmpty(ChgLED)Then
        For ii = 0 To UBound(ChgLED)
            chg = chgLED(ii, 1):stat = chgLED(ii, 2)
            For jj = 0 to 10
                If stat = Patterns(jj)OR stat = Patterns2(jj)then Digits(chgLED(ii, 0)).SetValue jj
            Next
        Next
    End IF
End Sub

'*******************
' GI Lights
'*******************

Sub GiOn
    Dim bulb
    For each bulb in aGiLights
        bulb.State = 1
    Next
End Sub

Sub GiOff
    Dim bulb
    For each bulb in aGiLights
        bulb.State = 0
    Next
End Sub

'Dipswitch 1 Controls High Score to Date Award
'OFF=3 Credits
'ON=1 Credit
'Dipswitch 2 Controls number of balls
'OFF=3
'ON=5
'Dipswitch 3 Controls Special Award
'OFF=Replay
'ON=Extra Ball

' Thalamus - 2021-04-30 : added proper exit

' Thalamus : Exit in a clean and proper way
Sub Table1_exit()
  Controller.Pause = False
  Controller.Stop
End Sub


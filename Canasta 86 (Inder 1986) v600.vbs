' Inder's Canasta 86 / IPD No. 4097 / 1986 / 4 Players
' VPX8 - version by JPSalas 2024, version 6.0.0
Option Explicit
Randomize

Const BallSize = 50
Const BallMass = 1

On Error Resume Next
ExecuteGlobal GetTextFile("controller.vbs")
If Err Then MsgBox "You need the controller.vbs in order to run this table, available in the vp10 package"
On Error Goto 0

LoadVPM "01550000", "inder.vbs", 3.26

Dim bsTrough, dtBankL, dtBankR, x

Const cGameName = "canasta"

Const UseSolenoids = 2
Const UseLamps = 0
Const UseGI = 0
Const UseSync = 0 'set it to 1 if the table runs too fast
Const HandleMech = 0

Dim VarHidden
If Table1.ShowDT = true then
    VarHidden = 1
    For each x in aReels
        x.Visible = 1
    Next
else
    VarHidden = 0
    For each x in aReels
        x.Visible = 0
    Next
end if

if B2SOn = true then VarHidden = 1

' Standard Sounds
Const SSolenoidOn = "fx_Solenoidon"
Const SSolenoidOff = "fx_Solenoidoff"
Const SCoin = "fx_Coin"

'************
' Table init.
'************

Sub table1_Init
    vpmInit me
    With Controller
        .GameName = cGameName
        If Err Then MsgBox "Can't start Game" & cGameName & vbNewLine & Err.Description:Exit Sub
        .SplashInfoLine = "Canasta 86 - Inder 1986" & vbNewLine & "VPX table by JPSalas v6.0.0"
        .HandleKeyboard = 0
        .ShowTitle = 0
        .ShowDMDOnly = 1
        .ShowFrame = 0
        .HandleMechanics = 0
        .Hidden = VarHidden
        .Games(cGameName).Settings.Value("rol") = 0   '1= rotated display, 0= normal
        .Games(cGameName).Settings.Value("sound") = 1 '1 enabled rom sound
        '.SetDisplayPosition 0,0, GetPlayerHWnd 'restore dmd window position
        On Error Resume Next
        Controller.SolMask(0) = 0
        vpmTimer.AddTimer 2000, "Controller.SolMask(0)=&Hffffffff'" 'ignore all solenoids - then add the Timer to renable all the solenoids after 2 seconds
        Controller.Run GetPlayerHWnd
        On Error Goto 0
    End With

    ' Nudging
    vpmNudge.TiltSwitch = 53
    vpmNudge.Sensitivity = 3
    vpmNudge.TiltObj = Array(Bumper1, Bumper2, Bumper3)

    ' Trough
    Set bsTrough = New cvpmBallStack
    With bsTrough
        .InitSw 0, 91, 0, 0, 0, 0, 0, 0
        .InitKick BallRelease, 90, 4
        .InitExitSnd SoundFX("fx_ballrel", DOFContactors), SoundFX("fx_Solenoid", DOFContactors)
        .Balls = 1
    End With

    ' Drop targets Left
    set dtBankL = new cvpmdroptarget
    With dtBankL
        .initdrop array(sw84, sw64, sw74), array(84, 64, 74)
        .initsnd SoundFX("fx_droptarget", DOFDropTargets), SoundFX("fx_resetdrop", DOFContactors)
        .CreateEvents("dtBankL")
    End With

    ' Drop targets Right
    set dtBankR = new cvpmdroptarget
    With dtBankR
        .initdrop array(sw85, sw65, sw75), array(85, 65, 75)
        .initsnd SoundFX("fx_droptarget", DOFDropTargets), SoundFX("fx_resetdrop", DOFContactors)
        .CreateEvents("dtBankR")
    End With

    ' Main Timer init
    PinMAMETimer.Interval = PinMAMEInterval
    PinMAMETimer.Enabled = 1

    ' Turn on Gi
    vpmtimer.addtimer 2000, "GiOn '"
    ' RealTimer timer
    RealTime.Enabled = 1
End Sub

'*******************************
' RealTime Animations and checks
'*******************************

Sub RealTime_Timer 'mostly object animations
    g1.rotx = 20 - Gate5.Currentangle / 4.5
    g2.rotx = 20 - Gate6.Currentangle / 4.5
    g3.rotx = 20 - Gate7.Currentangle / 4.5
    WireGAte.RotZ = 38 + Gate4.Currentangle / 4.5
    RollingUpdate
End Sub

Sub table1_Paused:Controller.Pause = 1:End Sub
Sub table1_unPaused:Controller.Pause = 0:End Sub
Sub table1_exit:Controller.stop:End Sub

'**********
' Keys
'**********

Sub table1_KeyDown(ByVal Keycode)
    If keycode = LeftTiltKey Then Nudge 90, 5:PlaySound SoundFX("fx_nudge", 0), 0, 1, -0.1, 0.25
    If keycode = RightTiltKey Then Nudge 270, 5:PlaySound SoundFX("fx_nudge", 0), 0, 1, 0.1, 0.25
    If keycode = CenterTiltKey Then Nudge 0, 6:PlaySound SoundFX("fx_nudge", 0), 0, 1, 0, 0.25
    If vpmKeyDown(KeyCode)Then Exit Sub
    If keycode = PlungerKey Then PlaySound "fx_PlungerPull", 0, 1, 0.1, 0.25:Plunger.Pullback
End Sub

Sub table1_KeyUp(ByVal Keycode)
    If vpmKeyUp(KeyCode)Then Exit Sub
    If keycode = PlungerKey Then PlaySound "fx_plunger", 0, 1, 0.1, 0.25:Plunger.Fire
End Sub

'*********
' Switches
'*********

' Bumpers
Sub Bumper1_Hit:vpmTimer.PulseSw 76:PlaySoundAt SoundFX("fx_bumper", DOFContactors), Bumper1:End Sub
Sub Bumper2_Hit:vpmTimer.PulseSw 86:PlaySoundAt SoundFX("fx_bumper", DOFContactors),Bumper2:End Sub
Sub Bumper3_Hit:vpmTimer.PulseSw 96:PlaySoundAt SoundFX("fx_bumper", DOFContactors), Bumper3:End Sub

' Drain & Saucers
Sub Drain_Hit:PlaySoundAt "fx_drain", Drain:bsTrough.AddBall Me:End Sub

' Rollovers
Sub sw77_Hit:Controller.Switch(77) = 1:PlaySoundAt "fx_sensor", sw77:End Sub
Sub sw77_UnHit:Controller.Switch(77) = 0:End Sub

Sub sw61a_Hit:Controller.Switch(61) = 1:PlaySoundAt "fx_sensor", sw61a:DOF 102, DOFOn:End Sub
Sub sw61a_UnHit:Controller.Switch(61) = 0:DOF 102, DOFOff:End Sub

Sub sw61b_Hit:Controller.Switch(61) = 1:PlaySoundAt "fx_sensor", sw61b:DOF 103, DOFOn:End Sub
Sub sw61b_UnHit:Controller.Switch(61) = 0:DOF 103, DOFOff:End Sub

Sub sw61c_Hit:Controller.Switch(61) = 1:PlaySoundAt "fx_sensor", sw61c:DOF 104, DOFOn:End Sub
Sub sw61c_UnHit:Controller.Switch(61) = 0:DOF 104, DOFOff:End Sub

Sub sw61d_Hit:Controller.Switch(61) = 1:PlaySoundAt "fx_sensor", sw61d:DOF 105, DOFOn:End Sub
Sub sw61d_UnHit:Controller.Switch(61) = 0:DOF 105, DOFOff:End Sub

Sub sw67_Hit:Controller.Switch(67) = 1:PlaySoundAt "fx_sensor", sw67:End Sub
Sub sw67_UnHit:Controller.Switch(67) = 0:End Sub

Sub sw81_Hit:Controller.Switch(81) = 1:PlaySoundAt "fx_sensor", sw81:End Sub
Sub sw81_UnHit:Controller.Switch(81) = 0:End Sub

Sub sw82_Hit:Controller.Switch(82) = 1:PlaySoundAt "fx_sensor", sw82:End Sub
Sub sw82_UnHit:Controller.Switch(82) = 0:End Sub

Sub sw80_Hit:Controller.Switch(80) = 1:PlaySoundAt "fx_sensor", sw80:End Sub
Sub sw80_UnHit:Controller.Switch(80) = 0:End Sub

Sub sw80a_Hit:Controller.Switch(80) = 1:PlaySoundAt "fx_sensor", sw80a:DOF 101, DOFOn:End Sub
Sub sw80a_UnHit:Controller.Switch(80) = 0:DOF 101, DOFOff:End Sub

Sub sw83_Hit:Controller.Switch(83) = 1:PlaySoundAt "fx_sensor", sw83:End Sub
Sub sw83_UnHit:Controller.Switch(83) = 0:End Sub

Sub sw66_Hit:Controller.Switch(66) = 1:PlaySoundAt "fx_sensor", sw66:End Sub
Sub sw66_UnHit:Controller.Switch(66) = 0:End Sub

Sub sw72_Hit:Controller.Switch(72) = 1:PlaySoundAt "fx_sensor", sw72:End Sub
Sub sw72_UnHit:Controller.Switch(72) = 0:End Sub

Sub sw71_Hit:Controller.Switch(71) = 1:PlaySoundAt "fx_sensor", sw71:End Sub
Sub sw71_UnHit:Controller.Switch(71) = 0:End Sub

'Spinner
Sub Spinner1_Spin():vpmTimer.PulseSw 73:PlaySoundAt "fx_spinner", Spinner1:End Sub

'Targets
Sub sw97_Hit:vpmTimer.PulseSw 97:PlaySoundAtBall SoundFX("fx_target", DOFDropTargets):End Sub
Sub sw87_Hit:vpmTimer.PulseSw 87:PlaySoundAtBall SoundFX("fx_target", DOFDropTargets):End Sub
Sub sw62_Hit:vpmTimer.PulseSw 62:PlaySoundAtBall SoundFX("fx_target", DOFDropTargets):End Sub
Sub sw63_Hit:vpmTimer.PulseSw 63:PlaySoundAtBall SoundFX("fx_target", DOFDropTargets):End Sub
Sub sw94_Hit:vpmTimer.PulseSw 94:PlaySoundAtBall SoundFX("fx_target", DOFDropTargets):End Sub

'Rubbers 'the sound is played from the rubber collection hit
'Rubber animations
Dim Rub1, Rub2

Sub sw2a_Hit:vpmTimer.PulseSw 2:Rub1 = 1:sw2a_Timer:End Sub

Sub sw2a_Timer
    Select Case Rub1
        Case 1:r12.Visible = 1:sw2a.TimerEnabled = 1
        Case 2:r12.Visible = 0:r13.Visible = 1
        Case 3:r13.Visible = 0:sw2a.TimerEnabled = 0
    End Select
    Rub1 = Rub1 + 1
End Sub

Sub sw2c_Hit:vpmTimer.PulseSw 2:Rub2 = 1:sw2c_Timer:End Sub
Sub sw2c_Timer
    Select Case Rub2
        Case 1:r14.Visible = 1:sw2c.TimerEnabled = 1
        Case 2:r14.Visible = 0:r15.Visible = 1
        Case 3:r15.Visible = 0:sw2c.TimerEnabled = 0
    End Select
    Rub2 = Rub2 + 1
End Sub

Sub sw90a_Hit:vpmTimer.PulseSw 90:End Sub
Sub sw90b_Hit:vpmTimer.PulseSw 90:End Sub
Sub sw92a_Hit:vpmTimer.PulseSw 92:End Sub
Sub sw92b_Hit:vpmTimer.PulseSw 92:End Sub
Sub sw92c_Hit:vpmTimer.PulseSw 92:End Sub
Sub sw92d_Hit:vpmTimer.PulseSw 92:End Sub
Sub sw92e_Hit:vpmTimer.PulseSw 92:End Sub
Sub sw93_Hit:vpmTimer.PulseSw 93:End Sub

'*********
'Solenoids
'*********

SolCallback(2) = "dtBankL.SolDropUp"
SolCallback(3) = "dtBankR.SolDropUp"
SolCallback(4) = "bsTrough.SolOut"
SolCallback(5) = "vpmNudge.SolGameOn"

'*******************
' Flipper Subs Rev3
'*******************

SolCallback(sLRFlipper) = "SolRFlipper"
SolCallback(sLLFlipper) = "SolLFlipper"

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

'*****************
'   Gi Effects
'*****************

Dim OldGiState
OldGiState = -1 'start witht he Gi off

Sub GiON
    For each x in aGiLights
        x.State = 1
    Next
End Sub

Sub GiOFF
    For each x in aGiLights
        x.State = 0
    Next
End Sub

Sub GiEffect(enabled)
    If enabled Then
        For each x in aGiLights
            x.Duration 2, 1000, 1
        Next
    End If
End Sub

Sub GIUpdate
    Dim tmp, obj
    tmp = Getballs
    If UBound(tmp) <> OldGiState Then
        OldGiState = Ubound(tmp)
        If UBound(tmp) = -1 Then
            GiOff
        Else
            GiOn
        End If
    End If
End Sub

'*************
' Update Lamps
'*************

LampTimer.Interval = 40
LampTimer.Enabled = 1

Sub LampTimer_Timer
    GIUpdate
    UpdateLeds
    UpdateLamps
End Sub

Sub UpdateLamps()
    Dim ChgLamp, ii
    ChgLamp = Controller.ChangedLamps
    If Not IsEmpty(ChgLamp)Then
        For ii = 0 To UBound(ChgLamp)
            DisplayLamps chgLamp(ii, 0), chgLamp(ii, 1)
        Next
    End If
End Sub

Sub DisplayLamps(idx, Stat)
    Select Case idx
        Case 1:li1.State = Stat
        Case 2:li2.State = Stat
        Case 3:li3.State = Stat
        Case 4:li4.State = Stat
        Case 5:li5.State = Stat
        Case 6:li6.State = Stat
        Case 7:li7.State = Stat
        Case 8:li8.State = Stat
        Case 9:li9.State = Stat
        Case 10:li10.State = Stat
        Case 11:li11a.State = Stat:li11b.State = Stat:li11c.state = stat:li11d.state = stat:li11e.state = stat:li11f.state = stat
        Case 13:li12.State = Stat 'text="1"
        Case 14:li13.State = Stat 'text="2"
        Case 15:li14.State = Stat 'text="3"
        Case 17:li15.State = Stat
        Case 18:li16.State = Stat
        Case 19:li17.State = Stat
        Case 22:li18.State = Stat
        Case 23:li19.State = Stat
        Case 24:li20.State = Stat
        Case 25:li21.State = Stat
        Case 26:li22.State = Stat
        Case 27:li23.State = Stat
        Case 28:li24.State = Stat:li24b.State = Stat
        Case 29:li25.State = Stat
        Case 30:li26.State = Stat
        Case 31:li27a.State = Stat:li27.State = Stat
        Case 32:li28.State = Stat 'text="Game Over"
        Case 33:li29.State = Stat 'text="Press Start"
        Case 35:li31.State = Stat 'text="Ball in Play"
        Case 36:li32.State = Stat 'text="Match"
        Case 37:li33.State = Stat
        Case 38:li34.State = Stat
        Case 39:li35.State = Stat 'text="Handicap"
        Case 40:If stat Then PlaySound SoundFX("fx_knocker", DOFKnocker)
        Case 41:li36.State = Stat
        Case 42:li37.State = Stat
        Case 43:li38.State = Stat
        Case 44:li39.State = Stat
        Case 45:li40.State = Stat
        Case 46:li41.State = Stat
    End Select
End Sub

'*********************
'LED's based on Eala's
'*********************

Dim Digits(32)
Digits(0) = Array(a00, a02, a05, a06, a04, a01, a03, a07)
Digits(1) = Array(a10, a12, a15, a16, a14, a11, a13)
Digits(2) = Array(a20, a22, a25, a26, a24, a21, a23)
Digits(3) = Array(a30, a32, a35, a36, a34, a31, a33, a37)
Digits(4) = Array(a40, a42, a45, a46, a44, a41, a43)
Digits(5) = Array(a50, a52, a55, a56, a54, a51, a53)
Digits(6) = Array(a60, a62, a65, a66, a64, a61, a63)

Digits(7) = Array(e00, e02, e05, e06, e04, e01, e03, e07)
Digits(8) = Array(e10, e12, e15, e16, e14, e11, e13)
Digits(9) = Array(e20, e22, e25, e26, e24, e21, e23)
Digits(10) = Array(e30, e32, e35, e36, e34, e31, e33, e37)
Digits(11) = Array(e40, e42, e45, e46, e44, e41, e43)
Digits(12) = Array(e50, e52, e55, e56, e54, e51, e53)
Digits(13) = Array(e60, e62, e65, e66, e64, e61, e63)

Digits(14) = Array(b00, b02, b05, b06, b04, b01, b03, b07)
Digits(15) = Array(b10, b12, b15, b16, b14, b11, b13)
Digits(16) = Array(b20, b22, b25, b26, b24, b21, b23)
Digits(17) = Array(b30, b32, b35, b36, b34, b31, b33, b37)
Digits(18) = Array(b40, b42, b45, b46, b44, b41, b43)
Digits(19) = Array(b50, b52, b55, b56, b54, b51, b53)
Digits(20) = Array(b60, b62, b65, b66, b64, b61, b63)

Digits(21) = Array(f00, f02, f05, f06, f04, f01, f03, f07)
Digits(22) = Array(f10, f12, f15, f16, f14, f11, f13)
Digits(23) = Array(f20, f22, f25, f26, f24, f21, f23)
Digits(24) = Array(f30, f32, f35, f36, f34, f31, f33, f37)
Digits(25) = Array(f40, f42, f45, f46, f44, f41, f43)
Digits(26) = Array(f50, f52, f55, f56, f54, f51, f53)
Digits(27) = Array(f60, f62, f65, f66, f64, f61, f63)

Digits(28) = Array(c00, c02, c05, c06, c04, c01, c03)
Digits(29) = Array(c10, c12, c15, c16, c14, c11, c13)

Digits(30) = Array(d00, d02, d05, d06, d04, d01, d03)
Digits(31) = Array(d10, d12, d15, d16, d14, d11, d13)

'********************
'Update LED's display
'********************

Sub UpdateLeds()
    Dim ChgLED, ii, num, chg, stat, obj
    ChgLED = Controller.ChangedLEDs(&Hffffffff, &Hffffffff)
    If Not IsEmpty(ChgLED)Then
        For ii = 0 To UBound(chgLED)
            num = chgLED(ii, 0):chg = chgLED(ii, 1):stat = chgLED(ii, 2)
            For Each obj In Digits(num)
                If chg And 1 Then obj.State = stat And 1
                chg = chg \ 2:stat = stat \ 2
            Next
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
Const lob = 0      'number of locked balls
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

'***************
' Rules by Maior
'***************
Dim Msg(21)
Sub Rules()
    Msg(0) = "Canasta 86 - Inder 1986" &Chr(10)&Chr(10)
    Msg(1) = ""
    Msg(2) = "3 Lanes Top Left: Left: Take Bonus x2 "
    Msg(3) = "   Center: Extra Ball when lit "
    Msg(4) = "   Right: Take bonus (If lit, take bonus x5)."
    Msg(5) = ""
    Msg(6) = "4 Lanes Top Right: Left: Bonus Up (If lit, 100.000 points)"
    Msg(7) = "   Center Left: Lights x5 lane. When lit, 50.000 points"
    Msg(8) = "   Center Right: When Lit, Extra ball"
    Msg(9) = "   Right: When lit Special, free game"
    Msg(10) = ""
    Msg(11) = "3 Left Bank: Two of them lights Extra Ball or Special"
    Msg(12) = "   Three of them lights Extra Ball and Special"
    Msg(13) = ""
    Msg(14) = "3 Right Bank: Two of them lights Extra Ball or Special (see backglass)"
    Msg(15) = "   Three of them lights Extra Ball and Special"
    Msg(16) = ""
    Msg(17) = "Center target: With Orange light, up bonus"
    Msg(18) = "   With Green light, lights x5 lane"
    Msg(19) = "   With Special light, free game"
    Msg(20) = ""
    Msg(21) = "Freegames at 1.500.000 points and 2.500.000 points. "
    For X = 1 To 21
        Msg(0) = Msg(0) + Msg(X)&Chr(13)
    Next
    MsgBox Msg(0), , "         Instructions and Rule Card"
End Sub

' DipVal - translate dip switch letter from Canasta 86 manual to bitmask
' by tynirodent
Function DipVal(Letter)
    If Letter = "A" Then
        DipVal = &H80000000 ' avoid overflow error
    Else
        DipVal = 2 ^(31 -(Asc(Letter)- Asc("A")))
    End If
End Function
Sub myShowDips()
    Dim vpmDips:Set vpmDips = New cvpmDips
    With vpmDips
        .AddForm 500, 400, "Canasta 86 - DIP switches"
        .AddFrame 2, 10, 190, "Balls Per Game", DipVal("E"), Array("3 balls", 0, "5 balls", DipVal("E"))
        .AddFrame 2, 60, 190, "Top Score", DipVal("O") + DipVal("P"), Array("2,000,000", 0, "2,400,000", DipVal("P"), "3,000,000", DipVal("O"), "3,500,000", DipVal("O") + DipVal("P"))
        .AddFrame 205, 10, 190, "Replay 1 Score", DipVal("C") + DipVal("D"), Array("1,500,000", 0, "1,700,000", DipVal("D"), "1,800,000", DipVal("C"), "1,900,000", DipVal("C") + DipVal("D"))
        .AddFrame 205, 90, 190, "Replay 2 Score", DipVal("A") + DipVal("B"), Array("2,500,000", 0, "2,700,000", DipVal("B"), "2,900,000", DipVal("A"), "none", DipVal("A") + DipVal("B"))
        .AddLabel 50, 185, 300, 20, "After hitting OK, press F3 to reset game with new settings."
        .ViewDips
    End With
End Sub
Set vpmShowDips = GetRef("myShowDips")

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

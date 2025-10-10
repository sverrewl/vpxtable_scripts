' El Bueno, EL Feo y el Malo
' Based on the Gotlieb's tables: El Dorado, Sheriff and Dakota
' VPM first version by JPSalas 2009
' VPX8 version 6.0.0 2025
' Uses the ROM from El Dorado City of Gold, "eldorado"

 Option Explicit
 Randomize

Const BallSize = 50
Const BallMass = 1

Const bMusicOn = True 'change it to False if you don't want the music
Const SongVolume = 1 'the volume of the song, from 0 to 1

On Error Resume Next
ExecuteGlobal GetTextFile("controller.vbs")
If Err Then MsgBox "You need the controller.vbs in order to run this table, available in the vp10 package"
On Error Goto 0

 LoadVPM "01120100", "sys80.vbs", 3.02

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
    lrail.Visible = 0
    rrail.Visible = 0
end if

if B2SOn = true then VarHidden = 1

  Dim bsTrough, dtRBank, dtTBank, x

 Const cGameName = "eldorado"

 Const UseSolenoids = 2
 Const UseLamps = 0
 Const UseGI = 0
 Const UseSync = 0
 Const HandleMech = 0

 ' Standard Sounds
 Const SSolenoidOn = "fx_SolenoidOn"
 Const SSolenoidOff = "fx_SolenoidOff"
 'Const SFlipperOn = "FlipperUp"
 'Const SFlipperOff = "FlipperDown"
 Const SCoin = "fx_coin"

 '************
 ' Table init.
 '************

 Sub Table1_Init
    vpmInit me
  Dim ii
    With Controller
       .GameName = cGameName
       .Games(cGameName).Settings.Value("sound") = 0 'turn off rom sound
       .SplashInfoLine = "El Bueno, el Feo y el Malo" & vbNewLine & "basado en varias tablas de Gottlieb"
       .HandleMechanics = 0
       .HandleKeyboard = 0
       .ShowDMDOnly = 1
       .ShowFrame = 0
       .ShowTitle = 0
       .Hidden = VarHidden
       If Err Then MsgBox Err.Description
    End With
    On Error Goto 0
    Controller.SolMask(0) = 0
    vpmTimer.AddTimer 2000, "Controller.SolMask(0)=&Hffffffff'" 'ignore all solenoids - then add the timer to renable all the solenoids after 2 seconds
    Controller.Run

    ' Nudging
    vpmNudge.TiltSwitch = 57
    vpmNudge.Sensitivity = 5
    vpmNudge.TiltObj = Array(Bumper1, Bumper2, LeftSlingshot, RightSlingshot)

  Set bsTrough = New cvpmTrough
  With bsTrough
  .size = 1
  '.entrySw = 18
  .initSwitches Array(67)
  .Initexit BallRelease, 80, 6
        .InitExitSounds  SoundFX("fx_ballrel", DOFContactors), SoundFX("fx_Solenoid", DOFContactors)
  .Balls = 1
  End With

    ' Top Drop targets
    set dtTBank = new cvpmdroptarget
    With dtTBank
       .InitDrop Array(sw0, sw10, sw20, sw30, sw40, sw1, sw11, sw21, sw31, sw41), Array(0, 10, 20, 30, 40, 1, 11, 21, 31, 41)
       .initsnd SoundFX("fx_droptarget", DOFDropTargets), SoundFX("fx_resetdrop", DOFContactors)
       '.CreateEvents "dtTBank"
    End With

    ' Right Drop targets
    set dtRBank = new cvpmdroptarget
    With dtRBank
       .InitDrop Array(sw2, sw12, sw22, sw32, sw42), Array(2, 12, 22, 32, 42)
       .initsnd SoundFX("fx_droptarget", DOFDropTargets), SoundFX("fx_resetdrop", DOFContactors)
       '.CreateEvents "dtRBank"
    End With

    ' Main Timer init
    PinMAMETimer.Interval = PinMAMEInterval
    PinMAMETimer.Enabled = 1
    RealTime.Enabled = 1
 End Sub

 Sub Table1_Paused:Controller.Pause = 1:End Sub
 Sub Table1_unPaused:Controller.Pause = 0:End Sub
Sub table1_exit:Controller.Games(cGameName).Settings.Value("sound") = 1:Controller.stop:End Sub

 '**********
 ' Keys
 '**********

 Sub table1_KeyDown(ByVal Keycode)
    If keyCode = LeftFlipperKey Then Controller.Switch(6) = 1
    If keyCode = RightFlipperKey Then Controller.Switch(16) = 1
    If keycode = LeftTiltKey Then Nudge 90, 4:PlaySound "fx_nudge", 0, 1, -0.1, 0.25
    If keycode = RightTiltKey Then Nudge 270, 4:PlaySound "fx_nudge", 0, 1, 0.1, 0.25
    If keycode = CenterTiltKey Then Nudge 0, 5:PlaySound "fx_nudge", 0, 1, 0, 0.25
    If vpmKeyDown(keycode) Then Exit Sub
    If keycode = PlungerKey Then PlaySoundAt "fx_PlungerPull",Plunger:Plunger.Pullback
 End Sub

 Sub table1_KeyUp(ByVal Keycode)
    If keyCode = LeftFlipperKey Then Controller.Switch(6) = 0
    If keyCode = RightFlipperKey Then Controller.Switch(16) = 0
    If vpmKeyUp(keycode) Then Exit Sub
    If keycode = PlungerKey Then PlaySoundAt "fx_plunger", Plunger:Plunger.Fire
 End Sub

 '*********
 ' Switches
 '*********

 ' Slingsshots
 Dim LStep, RStep

Sub LeftSlingShot_Slingshot
    PlaySoundAt SoundFX("fx_slingshot", DOFContactors),Lemk
    LeftSling4.Visible = 1
    Lemk.RotX = 26
    LStep = 0
    vpmTimer.PulseSw 66
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
    PlaySoundAt SoundFX("fx_slingshot", DOFContactors),Remk
    RightSling4.Visible = 1
    Remk.RotX = 26
    RStep = 0
    vpmTimer.PulseSw 66
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

 ' Bumpers
Sub Bumper1_Hit:vpmTimer.PulseSw 55:PlaySoundAt SoundFX("fx_bumper", DOFContactors),Bumper1:End Sub

Sub Bumper2_Hit:vpmTimer.PulseSw 55:PlaySoundAt SoundFX("fx_bumper", DOFContactors),Bumper2:End Sub

 ' Drain & holes
 Sub Drain_Hit:PlaysoundAt "fx_drain",Drain:bsTrough.AddBall Me:Gate.RotateToStart:PlayBonus.Enabled = 1:End Sub

 ' Rollovers
 Sub sw53_Hit:Controller.Switch(53) = 1:PlaySoundAt "fx_sensor", sw53
If l33.State Then
  PlaySound "fx_5000"
Else
  PlaySound "fx_500"
End If
End Sub

 Sub sw53_UnHit:Controller.Switch(53) = 0:End Sub

 Sub sw63_Hit:Controller.Switch(63) = 1:PlaySoundAt "fx_sensor", sw63
If l34.State Then
  PlaySound "fx_5000"
Else
  PlaySound "fx_500"
End If
End Sub
 Sub sw63_UnHit:Controller.Switch(63) = 0:End Sub

 Sub sw54_Hit:Controller.Switch(54) = 1:PlaySoundAt "fx_sensor", sw54
If l35.State Then
  PlaySound "fx_5000"
Else
  PlaySound "fx_500"
End If
End Sub
 Sub sw54_UnHit:Controller.Switch(54) = 0:End Sub

 Sub sw64_Hit:Controller.Switch(64) = 1:PlaySoundAt "fx_sensor", sw64
If l36.State Then
  PlaySound "fx_5000"
Else
  PlaySound "fx_500"
End If
End Sub
 Sub sw64_UnHit:Controller.Switch(64) = 0:End Sub

 Sub sw52_Hit:Controller.Switch(52) = 1:PlaySoundAt "fx_sensor", sw52
If l31.State Then
  PlaySound "fx_5000"
Else
  PlaySound "fx_500"
End If
End Sub
 Sub sw52_UnHit:Controller.Switch(52) = 0:End Sub

 Sub sw62_Hit:Controller.Switch(62) = 1:PlaySoundAt "fx_sensor", sw62
If l32.State Then
  PlaySound "fx_5000"
Else
  PlaySound "fx_500"
End If
End Sub
 Sub sw62_UnHit:Controller.Switch(62) = 0:End Sub

 Sub sw61_Hit:Controller.Switch(61) = 1:PlaySoundAt "fx_sensor", sw61
If l30.State Then
  PlaySound "fx_5000"
Else
  PlaySound "fx_500"
End If
End Sub
 Sub sw61_UnHit:Controller.Switch(61) = 0:End Sub

  Sub sw61a_Hit:Controller.Switch(61) = 1:PlaySoundAt "fx_sensor", sw61a
If l30b.State Then
  PlaySound "fx_5000"
Else
  PlaySound "fx_500"
End If
End Sub
 Sub sw61a_UnHit:Controller.Switch(61) = 0:End Sub

 Sub sw50_Hit:Controller.Switch(50) = 1:PlaySoundAt "fx_sensor", sw50
If l27.State Then
  PlaySound "fx_5000"
Else
  PlaySound "fx_500"
End If
End Sub
 Sub sw50_UnHit:Controller.Switch(50) = 0:End Sub

 Sub sw50a_Hit:Controller.Switch(50) = 1:PlaySoundAt "fx_sensor", sw50a
If l27b.State Then
  PlaySound "fx_5000"
Else
  PlaySound "fx_500"
End If
End Sub
 Sub sw50a_UnHit:Controller.Switch(50) = 0:End Sub

 Sub sw60_Hit:Controller.Switch(60) = 1:PlaySoundAt "fx_sensor", sw60
If l28.State Then
  PlaySound "fx_5000"
Else
  PlaySound "fx_500"
End If
End Sub
 Sub sw60_UnHit:Controller.Switch(60) = 0:End Sub

  Sub sw60a_Hit:Controller.Switch(60) = 1:PlaySoundAt "fx_sensor", sw60a
If l28b.State Then
  PlaySound "fx_5000"
Else
  PlaySound "fx_500"
End If
End Sub
 Sub sw60a_UnHit:Controller.Switch(60) = 0:End Sub

 Sub sw51_Hit:Controller.Switch(51) = 1:PlaySoundAt "fx_sensor", sw51
If l29.State Then
  PlaySound "fx_5000"
Else
  PlaySound "fx_500"
End If
End Sub
 Sub sw51_UnHit:Controller.Switch(51) = 0:End Sub

  Sub sw51a_Hit:Controller.Switch(51) = 1:PlaySoundAt "fx_sensor", sw51a
If l29b.State Then
  PlaySound "fx_5000"
Else
  PlaySound "fx_500"
End If
End Sub
 Sub sw51a_UnHit:Controller.Switch(51) = 0:End Sub

 ' Targets
 Sub sw65_Hit:vpmTimer.PulseSw 65:PlaySound "Gun2":PlaySoundAtBall SoundFX("fx_target", DOFTargets):Gate.RotateToStart:End Sub
 Sub sw65b_Hit:vpmTimer.PulseSw 65:PlaySound "Gun1":PlaySoundAtBall SoundFX("fx_target", DOFTargets):Gate.RotateToEnd:End Sub
  Sub sw65d_Hit:vpmTimer.PulseSw 65:PlaySound "Gun2":PlaySoundAtBall SoundFX("fx_target", DOFTargets):Gate.RotateToStart:End Sub

 ' Droptargets
 Sub sw0_Hit:dtTBank.hit 1
If l12.State Then
  PlaySound "fx_5000"
Else
  PlaySound "fx_500"
End If
End Sub
 Sub sw10_Hit:dtTBank.hit 2
If l13.State Then
  PlaySound "fx_5000"
Else
  PlaySound "fx_500"
End If
End Sub
 Sub sw20_Hit:dtTBank.hit 3
If l14.State Then
  PlaySound "fx_5000"
Else
  PlaySound "fx_500"
End If
End Sub
 Sub sw30_Hit:dtTBank.hit 4
If l15.State Then
  PlaySound "fx_5000"
Else
  PlaySound "fx_500"
End If
End Sub
 Sub sw40_Hit:dtTBank.hit 5
If l16.State Then
  PlaySound "fx_5000"
Else
  PlaySound "fx_500"
End If
End Sub
 Sub sw1_Hit:dtTBank.hit 6
If l17.State Then
  PlaySound "fx_5000"
Else
  PlaySound "fx_500"
End If
End Sub
 Sub sw11_Hit:dtTBank.hit 7
If l18.State Then
  PlaySound "fx_5000"
Else
  PlaySound "fx_500"
End If
End Sub
 Sub sw21_Hit:dtTBank.hit 8
If l19.State Then
  PlaySound "fx_5000"
Else
  PlaySound "fx_500"
End If
End Sub
 Sub sw31_Hit:dtTBank.hit 9
If l20.State Then
  PlaySound "fx_5000"
Else
  PlaySound "fx_500"
End If
End Sub
 Sub sw41_Hit:dtTBank.hit 10
If l21.State Then
  PlaySound "fx_5000"
Else
  PlaySound "fx_500"
End If
End Sub
 Sub sw2_Hit:dtRBank.hit 1
If l22.State Then
  PlaySound "fx_5000"
Else
  PlaySound "fx_500"
End If
End Sub
 Sub sw12_Hit:dtRBank.hit 2
If l23.State Then
  PlaySound "fx_5000"
Else
  PlaySound "fx_500"
End If
End Sub
 Sub sw22_Hit:dtRBank.hit 3
If l24.State Then
  PlaySound "fx_5000"
Else
  PlaySound "fx_500"
End If
End Sub
 Sub sw32_Hit:dtRBank.hit 4
If l25.State Then
  PlaySound "fx_5000"
Else
  PlaySound "fx_500"
End If
End Sub
 Sub sw42_Hit:dtRBank.hit 5
If l26.State Then
  PlaySound "fx_5000"
Else
  PlaySound "fx_500"
End If
End Sub


 '*********
 'Solenoids
 '*********
 SolCallback(5) = "dtRBank.SolDropUp"
 SolCallback(6) = "dtTbank.SolDropUp"
 solcallback(8) = "vpmsolsound ""fx_knocker"","
 solcallback(9) = "bsTrough.SolOut"

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

' flippers top animations

Sub LeftFlipper_Animate:LeftFlipperTop.RotZ = LeftFlipper.CurrentAngle: End Sub
Sub RightFlipper_Animate: RightFlipperTop.RotZ = RightFlipper.CurrentAngle: End Sub

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
        aBallShadow(b).Height = BOT(b).Z -Ballsize/2

        If BallVel(BOT(b))> 1 Then
            If BOT(b).z <30 Then
                ballpitch = Pitch(BOT(b))
                ballvol = Vol(BOT(b))
            Else
                ballpitch = Pitch(BOT(b)) + 50000 'increase the pitch on a ramp
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

  '************************
 '      Check Bonus
 '************************

 Dim BonusLamps, OldState
 BonusLamps = Array(37, 38, 39, 40, 41, 42, 43, 44, 45, 46) 'bonus lights
 PlayBonus.Enabled = 0

 Sub PlayBonus_Timer()
    Dim lamp, ii, state
    ii = 0
    state = 0
    For each lamp in BonusLamps
       If LampState(lamp) Then state = state + 2 ^ii

       ii = ii + 1
    Next
    If(OldState <> state) Then PlaySound "fx_Bonus"
    OldState = state
    If(state = 1 and LampState(47) = 0) Or LampState(11) = 1 Then PlayBonus.Enabled = 0
 End Sub

 '************************************
 '          LEDs Display
 '************************************

 Dim Digits(32)

 Set Digits(0) = a0
 Set Digits(1) = a1
 Set Digits(2) = a2
 Set Digits(3) = a3
 Set Digits(4) = a4
 Set Digits(5) = a5
 Set Digits(6) = a6

 Set Digits(7) = b0
 Set Digits(8) = b1
 Set Digits(9) = b2
 Set Digits(10) = b3
 Set Digits(11) = b4
 Set Digits(12) = b5
 Set Digits(13) = b6

 Set Digits(14) = c0
 Set Digits(15) = c1
 Set Digits(16) = c2
 Set Digits(17) = c3
 Set Digits(18) = c4
 Set Digits(19) = c5
 Set Digits(20) = c6

 Set Digits(21) = d0
 Set Digits(22) = d1
 Set Digits(23) = d2
 Set Digits(24) = d3
 Set Digits(25) = d4
 Set Digits(26) = d5
 Set Digits(27) = d6

 Set Digits(28) = e0
 Set Digits(29) = e1
 Set Digits(30) = e2
 Set Digits(31) = e3

 Sub UpdateLeds
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
    End IF
 End Sub

'**********************************************************
'     JP's Lamp Fading for VPX and Vpinmame v4.0
' FadingStep used for all kind of lamps
' FlashLevel used for modulated flashers
' LampState keep the real lamp state in a array
'**********************************************************

Dim LampState(200), FadingStep(200), FlashLevel(200)

InitLamps() ' turn off the lights and flashers and reset them to the default parameters

' vpinmame Lamp & Flasher Timers

Sub LampTimer_Timer()
    Dim chgLamp, num, chg, ii
    chgLamp = Controller.ChangedLamps
    If Not IsEmpty(chgLamp)Then
        For ii = 0 To UBound(chgLamp)
            LampState(chgLamp(ii, 0)) = chgLamp(ii, 1)       'keep the real state in an array
            FadingStep(chgLamp(ii, 0)) = chgLamp(ii, 1)
        Next
    End If
    UpdateLeds
    UpdateLamps
End Sub

 Sub UpdateLamps()

     Lamp 3, l3
     Lamp 4, l4
     Lamp 5, l5
     Lamp 6, l6
     Lamp 7, l7
     Reel 8, l8
     Reel 10, l10
     Reel 11, l11
     Lamp 12, l12
      Lamp 13, l13
     Lamp 14, l14
     Lamp 15, l15
      Lamp 16, l16
     Lamp 17, l17
     Lamp 18, l18
     Lamp 19, l19
     Lamp 20, l20
     Lamp 21, l21
     Lamp 22, l22
     Lamp 23, l23
     Lamp 24, l24
     Lamp 25, l25
      Lamp 26, l26
     Lampm 27, l27b
      Lamp 27, l27
     Lampm 28, l28b
      Lamp 28, l28
      Lampm 29, l29b
      Lamp 29, l29
      Lampm 30, l30b
     Lamp 30, l30
     Lamp 31, l31
      Lamp 32, l32
     Lamp 33, l33
     Lamp 34, l34
     Lamp 35, l35
     Lamp 36, l36
     Lamp 37, l37
     Lamp 38, l38
     Lamp 39, l39
     Lamp 40, l40
     Lamp 41, l41
     Lamp 42, l42
      Lamp 43, l43
     Lamp 44, l44
     Lamp 45, l45
     Lamp 46, l46
     Lamp 47, l47
    Lampm 90,bumper1light 'bumper 1
    Lamp 90,bumper1light1 'bumper 1
     Lampm 91, bumper2light 'bumper 22
     Lamp 91, bumper2light1 'bumper 22
  End Sub

' div lamp subs

' Normal Lamp & Flasher subs

Sub InitLamps()
    Dim x
    LampTimer.Interval = 10
    LampTimer.Enabled = 1
    For x = 0 to 200
        FadingStep(x) = 0
        FlashLevel(x) = 0
    Next
End Sub

Sub SetLamp(nr, value) ' 0 is off, 1 is on
    FadingStep(nr) = abs(value)
End Sub

' Lights: used for VPX standard lights, the fading is handled by VPX itself, they are here to be able to make them work together with the flashers

Sub Lamp(nr, object)
    Select Case FadingStep(nr)
        Case 1:object.state = 1:FadingStep(nr) = -1
        Case 0:object.state = 0:FadingStep(nr) = -1
    End Select
End Sub

Sub Lampm(nr, object) ' used for multiple lights, it doesn't change the fading state
    Select Case FadingStep(nr)
        Case 1:object.state = 1
        Case 0:object.state = 0
    End Select
End Sub

' Flashers:  0 starts the fading until it is off

Sub Flash(nr, object)
    Dim tmp
    Select Case FadingStep(nr)
        Case 1:Object.IntensityScale = 1:FadingStep(nr) = -1
        Case 0
            tmp = Object.IntensityScale * 0.85 - 0.01
            If tmp > 0 Then
                Object.IntensityScale = tmp
            Else
                Object.IntensityScale = 0
                FadingStep(nr) = -1
            End If
    End Select
End Sub

Sub Flashm(nr, object) 'multiple flashers, it doesn't change the fading state
    Dim tmp
    Select Case FadingStep(nr)
        Case 1:Object.IntensityScale = 1
        Case 0
            tmp = Object.IntensityScale * 0.85 - 0.01
            If tmp > 0 Then
                Object.IntensityScale = tmp
            Else
                Object.IntensityScale = 0
            End If
    End Select
End Sub

' Desktop Objects: Reels & texts

' Reels - 4 steps fading
Sub Reel(nr, object)
    Select Case FadingStep(nr)
        Case 1:object.SetValue 1:FadingStep(nr) = -1
        Case 0:object.SetValue 2:FadingStep(nr) = 2
        Case 2:object.SetValue 3:FadingStep(nr) = 3
        Case 3:object.SetValue 0:FadingStep(nr) = -1
    End Select
End Sub

Sub Reelm(nr, object)
    Select Case FadingStep(nr)
        Case 1:object.SetValue 1
        Case 0:object.SetValue 2
        Case 2:object.SetValue 3
        Case 3:object.SetValue 0
    End Select
End Sub

' Reels non fading
Sub NfReel(nr, object)
    Select Case FadingStep(nr)
        Case 1:object.SetValue 1:FadingStep(nr) = -1
        Case 0:object.SetValue 0:FadingStep(nr) = -1
    End Select
End Sub

Sub NfReelm(nr, object)
    Select Case FadingStep(nr)
        Case 1:object.SetValue 1
        Case 0:object.SetValue 0
    End Select
End Sub

'Texts

Sub Text(nr, object, message)
    Select Case FadingStep(nr)
        Case 1:object.Text = message:FadingStep(nr) = -1
        Case 0:object.Text = "":FadingStep(nr) = -1
    End Select
End Sub

Sub Textm(nr, object, message)
    Select Case FadingStep(nr)
        Case 1:object.Text = message
        Case 0:object.Text = ""
    End Select
End Sub

' Modulated Subs for the WPC tables

Sub SetModLamp(nr, level)
    FlashLevel(nr) = level / 150 'lights & flashers
End Sub

Sub LampMod(nr, object)          ' modulated lights used as flashers
    Object.IntensityScale = FlashLevel(nr)
    Object.State = 1             'in case it was off
End Sub

Sub FlashMod(nr, object)         'sets the flashlevel from the SolModCallback
    Object.IntensityScale = FlashLevel(nr)
End Sub

'Walls, flashers, ramps and Primitives used as 4 step fading images
'a,b,c,d are the images used from on to off

Sub FadeObj(nr, object, a, b, c, d)
    Select Case FadingStep(nr)
        Case 1:object.image = a:FadingStep(nr) = -1
        Case 0:object.image = b:FadingStep(nr) = 2
        Case 2:object.image = c:FadingStep(nr) = 3
        Case 3:object.image = d:FadingStep(nr) = -1
    End Select
End Sub

Sub FadeObjm(nr, object, a, b, c, d)
    Select Case FadingStep(nr)
        Case 1:object.image = a
        Case 0:object.image = b
        Case 2:object.image = c
        Case 3:object.image = d
    End Select
End Sub

Sub NFadeObj(nr, object, a, b)
    Select Case FadingStep(nr)
        Case 1:object.image = a:FadingStep(nr) = -1
        Case 0:object.image = b:FadingStep(nr) = -1
    End Select
End Sub

Sub NFadeObjm(nr, object, a, b)
    Select Case FadingStep(nr)
        Case 1:object.image = a
        Case 0:object.image = b
    End Select
End Sub

'******************
'   GI effects
' independent routine
' it turns on the gi
' when there is a ball
' in play
'******************

Dim OldGiState, GiIsOn
OldGiState = -1 'start witht the Gi off
GiIsOn = False

Sub GIUpdate
    Dim tmp, obj
    tmp = Getballs
    If UBound(tmp) <> OldGiState Then
        OldGiState = Ubound(tmp)
        If UBound(tmp) = -1 Then '-1 is the value returned by Getballs when no balls are on the table, so turn off the gi
            GiOff
        Else
            Gion
        End If
    End If
End Sub

Sub GiOn
    Dim bulb
    PlaySound"fx_GiOn"
    GiIsOn = True
    For each bulb in aGiLights
        bulb.State = 1
    Next
End Sub

Sub GiOff
    Dim bulb
    PlaySound"fx_GiOff"
    GiIsOn = False
    For each bulb in aGiLights
        bulb.State = 0
    Next
End Sub

' Play a song

Dim Song
Song = ""

Sub PlaySong(name)
    If bMusicOn Then
        If Song <> name Then
            StopSound Song
            Song = name
            PlaySound Song, -1, SongVolume
        End If
    End If
End Sub

Sub swPlunger_Hit: PlaySong "mu_"&RndNbr(5): End Sub

'******************
' RealTime Updates
'******************

Sub RealTime_Timer
    RollingUpdate
  GIUpdate
End Sub

Sub Gate_Animate: Gatep.RotZ = Gate.CurrentAngle: End Sub

 'Gottlieb El Dorado City of Gold
 'by Inkochnito
 Sub editDips
    Dim vpmDips:Set vpmDips = New cvpmDips
    With vpmDips
       .AddForm 700, 400, "El Bueno el Feo y el Malo - DIP switches"
       .AddFrame 2, 2, 190, "Maximum credits", 49152, Array("8 credits", 0, "10 credits", 32768, "15 credits", &H00004000, "20 credits", 49152)                                                                                   'dip 15&16
       .AddFrame 2, 80, 190, "Coin chute 1 and 2 control", &H00002000, Array("seperate", 0, "same", &H00002000)                                                                                                                   'dip 14
       .AddFrame 2, 126, 190, "Playfield special", &H00200000, Array("replay", 0, "extra ball", &H00200000)                                                                                                                       'dip 22
       .AddFrame 2, 172, 190, "Game mode", &H10000000, Array("replay", 0, "extra ball", &H10000000)                                                                                                                               'dip 29
       .AddChk 2, 227, 190, Array("Match feature", &H02000000)                                                                                                                                                                    'dip 26
       .AddChk 2, 248, 190, Array("Background sound", &H40000000)                                                                                                                                                                 'dip 31
       .AddFrame 205, 2, 190, "High score to date awards", &H00C00000, Array("not displayed and no award", 0, "displayed and no award", &H00800000, "displayed and 2 replays", &H00400000, "displayed and 3 replays", &H00C00000) 'dip 23&24
       .AddFrame 205, 80, 190, "Balls per game", &H01000000, Array("5 balls", 0, "3 balls", &H01000000)                                                                                                                           'dip 25
       .AddFrame 205, 126, 190, "Replay limit", &H04000000, Array("no limit", 0, "one per ball", &H04000000)                                                                                                                      'dip 27
       .AddFrame 205, 172, 190, "Novelty", &H08000000, Array("normal", 0, "points", &H0800000)                                                                                                                                    'dip 28
       .AddFrame 205, 218, 190, "3rd coin chute credits control", &H20000000, Array("no effect", 0, "add 9", &H20000000)                                                                                                          'dip 30
       .AddLabel 50, 270, 300, 20, "After hitting OK, press F3 to reset game with new settings."
       .ViewDips
    End With
 End Sub
 Set vpmShowDips = GetRef("editDips")

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

    ' Playfield image
    x = Table1.Option("Playfield image", 0, 1, 1, 1, 0, Array("Alternate", "Default") )
    If x Then table1.Image="pf" Else table1.Image="pf2"
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

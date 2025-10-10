' The Party Zone / IPD No. 1764 / August, 1991 / 4 Players
' VPX8 - version by JPSalas 2024, version 6.0.0

Option Explicit
Randomize

Const BallSize = 50
Const BallMass = 1

On Error Resume Next
ExecuteGlobal GetTextFile("controller.vbs")
If Err Then MsgBox "You need the controller.vbs in order to run this table, available in the vp10 package"
On Error Goto 0

Dim bsTrough, bsDJ, bsRP, x

Const UseSolenoids = 2
Const UseLamps = 0
Const UseSync = 0 'set it to 1 if the table runs too fast

Dim VarHidden, UseVPMColoredDMD
If Table1.ShowDT = true then
    UseVPMColoredDMD = true
    VarHidden = 1
else
    UseVPMColoredDMD = False
    VarHidden = 0
end if

if B2SOn = true then VarHidden = 1

Const cGameName = "pz_f4"

LoadVPM "01560000", "WPC.VBS", 3.26

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
        .SplashInfoLine = "The Party Zone - Bally 1991" & vbNewLine & "VPX table by JPSalas v6.0.0"
        .HandleKeyboard = 0
        .ShowTitle = 0
        .ShowDMDOnly = 1
        .ShowFrame = 0
        .HandleMechanics = 1
        .Hidden = VarHidden
        .Games(cGameName).Settings.Value("rol") = 0   '1= rotated display, 0= normal
        .Games(cGameName).Settings.Value("sound") = 1 '1 enabled rom sound
        '.SetDisplayPosition 0,0, GetPlayerHWnd 'restore dmd window position
        On Error Resume Next
        Controller.SolMask(0) = 0
        vpmTimer.AddTimer 2000, "Controller.SolMask(0)=&Hffffffff'" 'ignore all solenoids - then add the Timer to renable all the solenoids after 2 seconds
        Controller.Run GetPlayerHWnd
        On Error Goto 0
        .Switch(22) = 1 'close coin door
        .Switch(24) = 1 'and keep it close
    End With

    ' Nudging
    vpmNudge.TiltSwitch = 14
    vpmNudge.Sensitivity = 3
    vpmNudge.TiltObj = Array(Bumper1, Bumper2, Bumper3, LeftSlingshot, RightSlingshot)

    ' Trough
    Set bsTrough = New cvpmBallStack
    With bsTrough
        .InitSw 75, 76, 77, 78, 0, 0, 0, 0
        .InitKick BallRelease, 90, 4
        .InitExitSnd SoundFX("fx_ballrel", DOFContactors), SoundFX("fx_Solenoid", DOFContactors)
        .Balls = 3
        .IsTrough = 1
    End With

    ' DJ Hole
    Set bsDJ = New cvpmBallStack
    With bsDJ
        .InitSaucer sw67, 67, 233, 16
        .InitExitSnd SoundFX("fx_kicker", DOFContactors), SoundFX("fx_Solenoid", DOFContactors)
        .KickForceVar = 2
        .KickAngleVar = 2
    End With

    ' Right Popper
    Set bsRP = New cvpmBallStack
    With bsRP
        .InitSaucer sw42, 42, 0, 55
        .KickZ = 1.56
        .InitExitSnd SoundFX("fx_kicker", DOFContactors), SoundFX("fx_Solenoid", DOFContactors)
    End With

    ' Main Timer init
    PinMAMETimer.Interval = PinMAMEInterval
    PinMAMETimer.Enabled = 1
    GameTimer.Enabled = 1
End Sub

Sub table1_Paused:Controller.Pause = 1:End Sub
Sub table1_unPaused:Controller.Pause = 0:End Sub
Sub table1_exit:Controller.stop:End Sub

'**********
' Keys
'**********

Sub table1_KeyDown(ByVal Keycode)
    If keycode = LeftTiltKey Then Nudge 90, 5:PlaySound SoundFX("fx_nudge", 0), 0, 1, -0.1, 0.25:DummyShake2: RocketShake
    If keycode = RightTiltKey Then Nudge 270, 5:PlaySound SoundFX("fx_nudge", 0), 0, 1, 0.1, 0.25:DummyShake2: RocketShake
    If keycode = CenterTiltKey Then Nudge 0, 6:PlaySound SoundFX("fx_nudge", 0), 0, 1, 0, 0.25:DummyShake2: RocketShake
    If vpmKeyDown(keycode) Then Exit Sub
    If keycode = PlungerKey Then PlaySoundAt "fx_PlungerPull", Plunger:Plunger.Pullback
End Sub

Sub table1_KeyUp(ByVal Keycode)
    If vpmKeyUp(keycode) Then Exit Sub
    If keycode = PlungerKey Then PlaySoundAt "fx_plunger",Plunger:Plunger.Fire
End Sub

'******************
' RealTime Updates
'******************
' rolling sounds, flipper & animations

Sub GameTimer_Timer
    RollingUpdate
    DummyUpdate
End Sub

Sub bumper1_Animate: b1p.TransZ = bumper1.CurrentRingOffset:End Sub
Sub bumper2_Animate: b2p.TransZ = bumper2.CurrentRingOffset:End Sub
Sub bumper3_Animate: b3p.TransZ = bumper3.CurrentRingOffset:End Sub

'*********
' Switches
'*********

' Slings
Dim LStep, RStep

Sub LeftSlingShot_Slingshot
    PlaySoundAt SoundFX("fx_slingshot", DOFContactors), Lemk
    DOF 104, DOFPulse
    LeftSling4.Visible = 1
    Lemk.RotX = 26
    LStep = 0
    vpmTimer.PulseSw 73
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
    DOF 105, DOFPulse
    RightSling4.Visible = 1
    Remk.RotX = 26
    RStep = 0
    vpmTimer.PulseSw 74
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
Sub Bumper1_Hit:vpmTimer.PulseSw 43:PlaySoundAt SoundFX("fx_bumper", DOFContactors), Bumper1:End Sub
Sub Bumper2_Hit:vpmTimer.PulseSw 44:PlaySoundAt SoundFX("fx_bumper", DOFContactors), Bumper2:End Sub
Sub Bumper3_Hit:vpmTimer.PulseSw 45:PlaySoundAt SoundFX("fx_bumper", DOFContactors), Bumper3:End Sub

' Drain & Saucers
Sub Drain_Hit:PlaySoundAt "fx_drain", Drain:bsTrough.AddBall Me:End Sub
Sub sw67_Hit::PlaySoundAt "fx_kicker_enter",sw67:bsDJ.AddBall 0:End Sub
Sub sw42_Hit::PlaySoundAt "fx_kicker_enter",sw42:bsRP.AddBall 0:End Sub

' Rollovers
Sub sw55_Hit:Controller.Switch(55) = 1:PlaySoundAt "fx_sensor", sw55:End Sub
Sub sw55_UnHit:Controller.Switch(55) = 0:End Sub

Sub sw54_Hit:Controller.Switch(54) = 1:PlaySoundAt "fx_sensor", sw54:End Sub
Sub sw54_UnHit:Controller.Switch(54) = 0:End Sub

Sub sw57_Hit:Controller.Switch(57) = 1:PlaySoundAt "fx_sensor", sw57:End Sub
Sub sw57_UnHit:Controller.Switch(57) = 0:End Sub

Sub sw58_Hit:Controller.Switch(58) = 1:PlaySoundAt "fx_sensor", sw58:End Sub
Sub sw58_UnHit:Controller.Switch(58) = 0:End Sub

Sub sw63_Hit:Controller.Switch(63) = 1:PlaySoundAt "fx_sensor", sw63:End Sub
Sub sw63_UnHit:Controller.Switch(63) = 0:End Sub

Sub sw26_Hit:Controller.Switch(26) = 1:PlaySoundAt "fx_sensor", sw26:End Sub
Sub sw26_UnHit:Controller.Switch(26) = 0:End Sub

Sub sw27_Hit:Controller.Switch(27) = 1:PlaySoundAt "fx_sensor", sw27:End Sub
Sub sw27_UnHit:Controller.Switch(27) = 0:End Sub

Sub sw28_Hit:Controller.Switch(28) = 1:PlaySoundAt "fx_sensor", sw28:End Sub
Sub sw28_UnHit:Controller.Switch(28) = 0:End Sub

Sub sw65_Hit:Controller.Switch(65) = 1:PlaySoundAt "fx_sensor", sw65:End Sub
Sub sw65_UnHit:Controller.Switch(65) = 0:End Sub

Sub sw62_Hit:Controller.Switch(62) = 1:PlaySoundAt "fx_sensor", sw62:End Sub
Sub sw62_UnHit:Controller.Switch(62) = 0:End Sub

Sub sw85_Hit:Controller.Switch(85) = 1:PlaySoundAt "fx_sensor", sw85:End Sub
Sub sw85_UnHit:Controller.Switch(85) = 0:End Sub

Sub sw71_Hit:Controller.Switch(71) = 1:PlaySoundAt "fx_sensor", sw71:End Sub
Sub sw71_UnHit:Controller.Switch(71) = 0:End Sub

Sub sw72_Hit:Controller.Switch(72) = 1:PlaySoundAt "fx_sensor", sw72:End Sub
Sub sw72_UnHit:Controller.Switch(72) = 0:End Sub

Sub sw61_Hit:Controller.Switch(61) = 1:PlaySoundAt "fx_sensor", sw61:End Sub
Sub sw61_UnHit:Controller.Switch(61) = 0:End Sub

' Rubbers

Sub sw64_Hit:vpmTimer.PulseSw 64:End Sub

'Targets
Sub sw16_Hit:vpmTimer.PulseSw 16:PlaySoundAtBall SoundFX("fx_target", DOFDropTargets):End Sub
Sub sw17_Hit:vpmTimer.PulseSw 17:PlaySoundAtBall SoundFX("fx_target", DOFDropTargets):End Sub
Sub sw18_Hit:vpmTimer.PulseSw 18:PlaySoundAtBall SoundFX("fx_target", DOFDropTargets):End Sub

Sub sw34_Hit:vpmTimer.PulseSw 34:PlaySoundAtBall SoundFX("fx_target", DOFDropTargets):End Sub
Sub sw35_Hit:vpmTimer.PulseSw 35:PlaySoundAtBall SoundFX("fx_target", DOFDropTargets):End Sub
Sub sw36_Hit:vpmTimer.PulseSw 36:PlaySoundAtBall SoundFX("fx_target", DOFDropTargets):End Sub
Sub sw37_Hit:vpmTimer.PulseSw 37:PlaySoundAtBall SoundFX("fx_target", DOFDropTargets):End Sub
Sub sw38_Hit:vpmTimer.PulseSw 38:PlaySoundAtBall SoundFX("fx_target", DOFDropTargets):End Sub

Sub sw56_Hit:vpmTimer.PulseSw 56:PlaySoundAtBall SoundFX("fx_target", DOFDropTargets):End Sub
Sub sw66_Hit:vpmTimer.PulseSw 66:PlaySoundAtBall SoundFX("fx_target", DOFDropTargets):End Sub
Sub sw68_Hit:vpmTimer.PulseSw 68:PlaySoundAtBall SoundFX("fx_target", DOFDropTargets):End Sub

Sub sw81_Hit:vpmTimer.PulseSw 81:PlaySoundAtBall SoundFX("fx_target", DOFDropTargets):End Sub
Sub sw82_Hit:vpmTimer.PulseSw 82:PlaySoundAtBall SoundFX("fx_target", DOFDropTargets):End Sub
Sub sw83_Hit:vpmTimer.PulseSw 83:PlaySoundAtBall SoundFX("fx_target", DOFDropTargets):End Sub
Sub sw84_Hit:vpmTimer.PulseSw 84:PlaySoundAtBall SoundFX("fx_target", DOFDropTargets):End Sub

'*********
'Solenoids
'*********

SolCallback(1) = "SolBackPopper"
SolCallback(2) = "bsRP.SolOut"
SolCallback(3) = "SolDJMouth"
SolCallback(4) = "SolDJPopper"
SolCallback(5) = "SolDancinDummy"
SolCallback(6) = "SolComicMouth"
SolCallback(7) = "vpmSolSound SoundFX(""fx_Knocker"",DOFKnocker),"
SolCallback(9) = "bsTrough.SolIn"
SolCallback(10) = "bsTrough.SolOut"
SolCallback(23) = "SolHeadOnOff"
'SolCallback(24) = "SolHeadDir"

SolCallback(17) = "SetLamp 117,"
SolCallback(18) = "SetLamp 118,"
SolCallback(19) = "SetLamp 119,"
SolCallback(20) = "SetLamp 120,"
SolCallback(21) = "SetLamp 121,"
SolCallback(22) = "SetLamp 122,"
SolCallback(25) = "SetLamp 125,"
SolCallback(26) = "SetLamp 126,"
SolCallback(27) = "SetLamp 127,"
SolCallback(28) = "SetLamp 128,"

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

Sub LeftFlipper_Animate: LeftFlipperp.RotZ = LeftFlipper.CurrentAngle: End Sub
Sub RightFlipper_Animate: RightFlipperp.RotZ = RightFlipper.CurrentAngle: End Sub

' flippers hit Sound

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

'**************
' Solenoid Subs
'**************

Sub SolDJPopper(Enabled)
    If Enabled Then
        bsDJ.SolOut 1
        sw67Metal.Rotx = -10
        sw67.TimerEnabled = 1
    End If
End Sub

Sub sw67_Timer:sw67metal.Rotx = 0:Me.TimerEnabled = 0:End Sub

' Dummy dancing

Dim cBall
DummyInit

Sub SolDancinDummy(Enabled) 'moves dummy up and down
    PlaySoundAt "fx_dancing", dBody
    If Enabled Then
        DummyShake
        dBody.TransZ = -20
        dRightArm.TransZ = -20
        dLeftArm.TransZ = -20
        dRightLeg.TransZ = -20
        dLeftLeg.TransZ = -20
    Else
        dBody.TransZ = 0
        dRightArm.TransZ = 0
        dLeftArm.TransZ = 0
        dRightLeg.TransZ = 0
        dLeftLeg.TransZ = 0
    End If
End Sub

Sub DummyInit
    Set cBall = ckicker.createball
    ckicker.Kick 0, 0
End Sub

Sub DummyShake
    cball.velx = -2 + 2 * RND(1)
    cball.vely = -10 + 2 * RND(1)
End Sub

Sub DummyShake2 'when nudging
    cball.velx = 2 * RND(1)
    cball.vely = -5 + 2 * RND(1)
End Sub

Sub DummyUpdate 'moves arms and legs according to the captive ball, from the GameTimer
    Dim a, b
    a = ckicker.y - cball.y
    b = cball.x - ckicker.x
    dRightArm.rotx = - a
    dLeftArm.rotx = - b
    dRightLeg.rotx = a
    dLeftLeg.rotx = b
End Sub

' Comic Mouth animation

Sub SolComicMouth(Enabled)
    If Enabled Then
        ComicMouth.Z = 65
    Else
        PlaySoundAt "fx_mouth", ComicMouth
        ComicMouth.Z = 75
    End If
End Sub

' Back Popper

Sub sw31_Hit()
    sw31.DestroyBall
    PlaySoundAt "fx_kicker_enter", sw31
    if Controller.Switch(41) Then
        Controller.Switch(31) = 1
    else
        vpmTimer.PulseSwitch 31, 0, ""
        sw31.TimerEnabled = 1
    end if
End Sub

Sub sw31_Timer()
    sw31.TimerEnabled = 0
    Controller.Switch(41) = 1
End Sub

Sub SolBackPopper(Enabled)
    if(Enabled and Controller.Switch(41) ) Then
        PlaySoundAt "fx_kicker", sw41
        Controller.Switch(41) = 0
        sw41.CreateBall
        sw41.Kick 190, 26
    RocketShake
        if Controller.Switch(31) then
            Controller.Switch(31) = 0
            Controller.Switch(41) = 1
        end if
    else
        PlaySoundAt "fx_solenoid", sw41
    end if
End Sub

' Rocket shaking
Dim ShakeStep:ShakeStep = 0

Sub RocketShake
    ShakeStep = 6
    RocketTimer.Enabled = True
End Sub

Sub RocketTimer_Timer
    Rocket.Height = 260 + ShakeStep
    If ShakeStep = 0 Then RocketTimer.Enabled = False:Exit Sub
    If ShakeStep < 0 Then
        ShakeStep = ABS(ShakeStep) - 0.1
    Else
        ShakeStep = - ShakeStep + 0.1
    End If
End Sub

'******************
' DJ Head Movement
'******************

Dim DJPos, DJDest
dim zRot:zRot = 10
dim targetz:targetz = 10
dim motorSound:motorSound = 0
DJPos = 3:DJDest = 3

Sub SolDJMouth(Enabled)
    If Enabled Then
        PlaySoundAt "DJMouth", pDJHead
        pdjhead.visible = 0:pdjheadopen.visible = 1
    Else
        pdjhead.visible = 1:pdjheadopen.visible = 0
    End If
End Sub

' Mech update sub, the actual movement is done is the next subs
Sub UpdateDJ(aNewPos, aSpeed, aLastPos)
End Sub

Sub SolHeadOnOff(Enabled)
    DJTimer.Enabled = 1
End Sub

Sub SoundTimer_Timer
    if motorSound = 1 then
        PlaySoundAt "fx_DJMotor", pDJHead
    else
        SoundTimer.enabled = 0
    end if
End Sub

Sub DJTimer_Timer
    If DJpos = DJDest Then
        Me.Enabled = 0
        motorSound = 0
        Exit Sub
    End If

    if djdest = 1 then targetz = -140 end if
    if djdest = 2 then targetz = -90 end if
    if djdest = 3 then targetz = -35 end if
    if djdest = 4 then targetz = 5 end if
    if djdest = 5 then targetz = 35 end if
    if djdest = 6 then targetz = 90 end if
    if djdest = 7 then targetz = 140 end if
    if zrot> targetz then zrot = zrot-5:end if
    if zrot <targetz then zrot = zrot + 5:end if
    if zrot = targetz then
        djpos = djdest
        motorSound = 0
    end if
    pdjhead.rotz = zrot
    pdjheadopen.rotz = zrot
    if motorSound = 0 then
        motorSound = 1
        SoundTimer.enabled = 1
    end if
End Sub

Sub DJTrigger0_Hit:DJDest = 1:End Sub

Sub DJTrigger1_Hit:DJDest = 2:End Sub

Sub DJTrigger2_Hit:DJDest = 3:End Sub

Sub DJTrigger3_Hit:DJDest = 4:End Sub

Sub DJTrigger4_Hit:DJDest = 5:End Sub

Sub DJTrigger5_Hit:DJDest = 6:End Sub

Sub DJTrigger6_Hit:DJDest = 7:End Sub

'****************************
'        Update GI
'light.State = ABS(status)
'wall.IsDropped = NOT status
'****************************

UpdateGI 0, 0:UpdateGI 1, 0:UpdateGI 2, 0

Set GiCallback2 = GetRef("UpdateGI")

Sub UpdateGI(no, step)
    Dim gistep, ii
    gistep = step / 8
    Select Case no
        Case 0 ' red
            For each ii in aGIRed
                ii.IntensityScale = gistep
            Next
        Case 1 'left
            For each ii in aGILeft
                ii.IntensityScale = gistep
            Next
        Case 2 'right
            For each ii in aGIRight
                ii.IntensityScale = gistep
            Next
    End Select
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
    UpdateLamps
End Sub

Sub UpdateLamps()
    Lamp 11, li11
    Lamp 12, li12
    Lamp 13, li13
    Lamp 14, li14
    Lamp 15, li15
    Lamp 16, li16
    Lamp 17, li17
    Lamp 18, li18
    Lamp 21, li21
    Lamp 22, li22
    Lamp 23, li23
    Lamp 24, li24
    Lamp 25, li25
    Lamp 26, li26
    Lamp 27, li27
    Lamp 28, li28
    Lamp 31, li31
    Lamp 32, li32
    Lamp 33, li33
    Lamp 34, li34
    Lamp 35, li35
    Lamp 36, li36
    Lamp 37, li37
    Lamp 38, li38
    Lamp 41, li41
    Lamp 42, li42
    Lamp 43, li43
    Lamp 44, li44
    Lamp 45, li45
    Lamp 46, li46
    Lamp 47, li47
    Lamp 48, li48
    Lamp 51, li51
    Lamp 52, li52
    Lamp 53, li53
    Lamp 54, li54
    Lamp 55, li55
    Lamp 56, li56
    Lamp 57, li57
    Lamp 58, li58
    Lamp 61, li61
    Lamp 62, li62
    Lamp 63, li63
    Lamp 64, li64
    Lamp 65, li65
    Lamp 66, li66
    Lamp 67, li67
    Lamp 68, li68
    Lamp 71, li71
    Lamp 72, li72
    Lamp 73, li73
    Lamp 74, li74
    FadeObj 75, b1p, "bumper_on", "bumper_a", "bumper_b", "bumper_off"
    FadeObj 76, b2p, "bumper_on", "bumper_a", "bumper_b", "bumper_off"
    FadeObj 77, b3p, "bumper_on", "bumper_a", "bumper_b", "bumper_off"
    'Lamp 78, l78 'credit button
    Lamp 81, li81
    Lamp 82, li82
    Flash 83, li83
    Flash 84, li84
    Flash 85, li85
    Lamp 86, li86

    'flashers

    Lampm 117, f17a
    Lamp 117, f17
    Flash 118, f18
    Lamp 119, f19
    Lamp 120, f20
    Flash 121, f21
    Lampm 122, f22a
    Lamp 122, f22
    Lamp 125, f25
    Lampm 126, f26a
    Lampm 126, f26b
    Lamp 126, f26
    Lampm 127, f27a
    Lampm 127, f27b
    Lamp 127, f27
    Lamp 128, f28
End Sub

' div lamp subs

' Normal Lamp & Flasher subs

Sub InitLamps()
    Dim x
    LampTimer.Interval = 20
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
Const maxvel = 40 'max ball velocity
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
                ballpitch = Pitch(BOT(b)) + 30000 'increase the pitch on a ramp
                ballvol = Vol(BOT(b)) * 3
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

Dim LUTImage

Sub Table1_OptionEvent(ByVal eventId)
    Dim x, y

    'LUT
    LutImage = Table1.Option("Select LUT", 0, 21, 1, 0, 0, Array("Normal 0", "Normal 1", "Normal 2", "Normal 3", "Normal 4", "Normal 5", "Normal 6", "Normal 7", "Normal 8", "Normal 9", "Normal 10", _
        "Warm 0", "Warm 1", "Warm 2", "Warm 3", "Warm 4", "Warm 5", "Warm 6", "Warm 7", "Warm 8", "Warm 9", "Warm 10") )
    UpdateLUT

    ' Desktop DMD
    x = Table1.Option("Desktop DMD", 0, 1, 1, 1, 0, Array("Hide", "Show") )
    DesktopDMD.visible = x

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

' The Party Zone / IPD No. 1764 / August, 1991 / 4 Players
' VPX - version by JPSalas 2018, version 1.0.0

' Thalamus 2018-07-24
' Added/Updated "Positional Sound Playback Functions" and "Supporting Ball & Sound Functions"
' Changed UseSolenoids=1 to 2
' No special SSF tweaks yet.
' This is a JP table. He often uses walls as switches so I need to be careful of using PlaySoundAt

Option Explicit
Randomize

Const BallSize = 50
Const BallMass = 1.1

On Error Resume Next
ExecuteGlobal GetTextFile("controller.vbs")
If Err Then MsgBox "You need the controller.vbs in order to run this table, available in the vp10 package"
On Error Goto 0

Dim bsTrough, bsDJ, bsRP, x

Const UseSolenoids = 2
Const UseLamps = 0
Const UseSync = 0 'set it to 1 if the table runs too fast

Dim VarHidden, UseVPMDMD
If Table1.ShowDT = true then
    UseVPMDMD = true
    VarHidden = 1
else
    UseVPMDMD = False
    VarHidden = 0
    dmd1.Visible = 0
    dmd2.Visible = 0
    dmd3.Visible = 0
    lrail.Visible = 0
    rrail.Visible = 0
end if

if B2SOn = true then VarHidden = 1

Const cGameName = "pz_f4"

LoadVPM "01560000", "WPC.VBS", 3.26

' Standard Sounds
Const SSolenoidOn = "fx_Solenoid"
Const SSolenoidOff = ""
Const SCoin = "fx_Coin"

'************
' Table init.
'************

Sub table1_Init
    vpmInit me
    With Controller
        .GameName = cGameName
        If Err Then MsgBox "Can't start Game" & cGameName & vbNewLine & Err.Description:Exit Sub
        .SplashInfoLine = "The Party Zone - Bally 1991" & vbNewLine & "VPX table by JPSalas v.1.0.0"
        .HandleKeyboard = 0
        .ShowTitle = 0
        .ShowDMDOnly = 1
        .ShowFrame = 0
        .HandleMechanics = 1
        .Hidden = VarHidden
        .Games(cGameName).Settings.Value("rol") = 0 '1= rotated display, 0= normal
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
        .InitSaucer sw67, 67, 250, 9
        .InitExitSnd SoundFX("fx_kicker", DOFContactors), SoundFX("fx_Solenoid", DOFContactors)
        .KickForceVar = 2
        .KickAngleVar = 2
    End With

    ' Right Popper
    Set bsRP = New cvpmBallStack
    With bsRP
        .InitSaucer sw42, 42, 0, 50
        .KickZ = 1.56
        .InitExitSnd SoundFX("fx_kicker", DOFContactors), SoundFX("fx_Solenoid", DOFContactors)
    End With

    ' Main Timer init
    PinMAMETimer.Interval = PinMAMEInterval
    PinMAMETimer.Enabled = 1
    GameTimer.Enabled = 1

    ' Init some objects
    sw67metal.IsDropped = 1
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
    If keycode = PlungerKey Then PlaySound "fx_PlungerPull", 0, 1, 0.1, 0.25:Plunger.Pullback
    If vpmKeyDown(keycode) Then Exit Sub
End Sub

Sub table1_KeyUp(ByVal Keycode)
    If vpmKeyUp(keycode) Then Exit Sub
    If keycode = PlungerKey Then PlaySound "fx_plunger", 0, 1, 0.1, 0.25:Plunger.Fire
End Sub

'******************
' RealTime Updates
'******************
' rolling sounds, flipper & animations

Sub GameTimer_Timer
    RollingUpdate
    b1p.TransZ = - bumper1f.CurrentAngle
    b2p.TransZ = - bumper2f.CurrentAngle
    b3p.TransZ = - bumper3f.CurrentAngle
    LeftFlipperP.Rotz = LeftFlipper.CurrentAngle
    RightFlipperP.Rotz = RightFlipper.CurrentAngle
	FlipperLSh.RotZ = LeftFlipper.currentangle
	FlipperRSh.RotZ = RightFlipper.currentangle
    DummyUpdate
	RocketUpdate
End Sub

'*********
' Switches
'*********

' Slings
Dim LStep, RStep

Sub LeftSlingShot_Slingshot
    PlaySound SoundFX("fx_slingshot", DOFContactors), 0, 1, -0.05, 0.05
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
    PlaySound SoundFX("fx_slingshot", DOFContactors), 0, 1, 0.05, 0.05
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
Sub Bumper1_Hit:vpmTimer.PulseSw 43:PlaySound SoundFX("fx_bumper", DOFContactors), 0, 1, -0.05:bumper1f.RotateToEnd:VpmTimer.AddTimer 100, "bumper1f.RotateToStart '":End Sub
Sub Bumper2_Hit:vpmTimer.PulseSw 44:PlaySound SoundFX("fx_bumper", DOFContactors), 0, 1, 0.05:bumper2f.RotateToEnd:VpmTimer.AddTimer 100, "bumper2f.RotateToStart '":End Sub
Sub Bumper3_Hit:vpmTimer.PulseSw 45:PlaySound SoundFX("fx_bumper", DOFContactors), 0, 1, 0:bumper3f.RotateToEnd:VpmTimer.AddTimer 100, "bumper3f.RotateToStart '":End Sub

' Drain & Saucers
Sub Drain_Hit:Playsound "fx_drain":bsTrough.AddBall Me:End Sub
Sub sw67_Hit::PlaySound "fx_kicker_enter":bsDJ.AddBall 0:End Sub
Sub sw42_Hit::PlaySound "fx_kicker_enter", 0, 1, 0.05:bsRP.AddBall 0:End Sub

' Rollovers
Sub sw55_Hit:Controller.Switch(55) = 1:PlaySound "fx_sensor", 0, 1, pan(ActiveBall):End Sub
Sub sw55_UnHit:Controller.Switch(55) = 0:End Sub

Sub sw54_Hit:Controller.Switch(54) = 1:PlaySound "fx_sensor", 0, 1, pan(ActiveBall):End Sub
Sub sw54_UnHit:Controller.Switch(54) = 0:End Sub

Sub sw57_Hit:Controller.Switch(57) = 1:PlaySound "fx_sensor", 0, 1, pan(ActiveBall):End Sub
Sub sw57_UnHit:Controller.Switch(57) = 0:End Sub

Sub sw58_Hit:Controller.Switch(58) = 1:PlaySound "fx_sensor", 0, 1, pan(ActiveBall):End Sub
Sub sw58_UnHit:Controller.Switch(58) = 0:End Sub

Sub sw63_Hit:Controller.Switch(63) = 1:PlaySound "fx_sensor", 0, 1, pan(ActiveBall):End Sub
Sub sw63_UnHit:Controller.Switch(63) = 0:End Sub

Sub sw26_Hit:Controller.Switch(26) = 1:PlaySound "fx_sensor", 0, 1, pan(ActiveBall):End Sub
Sub sw26_UnHit:Controller.Switch(26) = 0:End Sub

Sub sw27_Hit:Controller.Switch(27) = 1:PlaySound "fx_sensor", 0, 1, pan(ActiveBall):End Sub
Sub sw27_UnHit:Controller.Switch(27) = 0:End Sub

Sub sw28_Hit:Controller.Switch(28) = 1:PlaySound "fx_sensor", 0, 1, pan(ActiveBall):End Sub
Sub sw28_UnHit:Controller.Switch(28) = 0:End Sub

Sub sw65_Hit:Controller.Switch(65) = 1:PlaySound "fx_sensor", 0, 1, pan(ActiveBall):End Sub
Sub sw65_UnHit:Controller.Switch(65) = 0:End Sub

Sub sw62_Hit:Controller.Switch(62) = 1:PlaySound "fx_sensor", 0, 1, pan(ActiveBall):End Sub
Sub sw62_UnHit:Controller.Switch(62) = 0:End Sub

Sub sw85_Hit:Controller.Switch(85) = 1:PlaySound "fx_sensor", 0, 1, pan(ActiveBall):End Sub
Sub sw85_UnHit:Controller.Switch(85) = 0:End Sub

Sub sw71_Hit:Controller.Switch(71) = 1:PlaySound "fx_sensor", 0, 1, pan(ActiveBall):End Sub
Sub sw71_UnHit:Controller.Switch(71) = 0:End Sub

Sub sw72_Hit:Controller.Switch(72) = 1:PlaySound "fx_sensor", 0, 1, pan(ActiveBall):End Sub
Sub sw72_UnHit:Controller.Switch(72) = 0:End Sub

Sub sw61_Hit:Controller.Switch(61) = 1:PlaySound "fx_sensor", 0, 1, pan(ActiveBall):End Sub
Sub sw61_UnHit:Controller.Switch(61) = 0:End Sub

' Rubbers

Sub sw64_Hit:vpmTimer.PulseSw 64:PlaySound SoundFX("fx_rubber", DOFDropTargets), 0, 1, pan(ActiveBall):End Sub

'Targets
Sub sw16_Hit:vpmTimer.PulseSw 16:PlaySound SoundFX("fx_target", DOFDropTargets), 0, 1, pan(ActiveBall):End Sub
Sub sw17_Hit:vpmTimer.PulseSw 17:PlaySound SoundFX("fx_target", DOFDropTargets), 0, 1, pan(ActiveBall):End Sub
Sub sw18_Hit:vpmTimer.PulseSw 18:PlaySound SoundFX("fx_target", DOFDropTargets), 0, 1, pan(ActiveBall):End Sub

Sub sw34_Hit:vpmTimer.PulseSw 34:PlaySound SoundFX("fx_target", DOFDropTargets), 0, 1, pan(ActiveBall):End Sub
Sub sw35_Hit:vpmTimer.PulseSw 35:PlaySound SoundFX("fx_target", DOFDropTargets), 0, 1, pan(ActiveBall):End Sub
Sub sw36_Hit:vpmTimer.PulseSw 36:PlaySound SoundFX("fx_target", DOFDropTargets), 0, 1, pan(ActiveBall):End Sub
Sub sw37_Hit:vpmTimer.PulseSw 37:PlaySound SoundFX("fx_target", DOFDropTargets), 0, 1, pan(ActiveBall):End Sub
Sub sw38_Hit:vpmTimer.PulseSw 38:PlaySound SoundFX("fx_target", DOFDropTargets), 0, 1, pan(ActiveBall):End Sub

Sub sw56_Hit:vpmTimer.PulseSw 56:PlaySound SoundFX("fx_target", DOFDropTargets), 0, 1, pan(ActiveBall):End Sub
Sub sw66_Hit:vpmTimer.PulseSw 66:PlaySound SoundFX("fx_target", DOFDropTargets), 0, 1, pan(ActiveBall):End Sub
Sub sw68_Hit:vpmTimer.PulseSw 68:PlaySound SoundFX("fx_target", DOFDropTargets), 0, 1, pan(ActiveBall):End Sub

Sub sw81_Hit:vpmTimer.PulseSw 81:PlaySound SoundFX("fx_target", DOFDropTargets), 0, 1, pan(ActiveBall):End Sub
Sub sw82_Hit:vpmTimer.PulseSw 82:PlaySound SoundFX("fx_target", DOFDropTargets), 0, 1, pan(ActiveBall):End Sub
Sub sw83_Hit:vpmTimer.PulseSw 83:PlaySound SoundFX("fx_target", DOFDropTargets), 0, 1, pan(ActiveBall):End Sub
Sub sw84_Hit:vpmTimer.PulseSw 84:PlaySound SoundFX("fx_target", DOFDropTargets), 0, 1, pan(ActiveBall):End Sub

' Ramps helpers
Sub RHelp1_Hit()
    PlaySound "fx_balldrop", 0, 1, -0.05
End Sub

Sub RHelp2_Hit()
    PlaySound "fx_balldrop", 0, 1, 0.05
End Sub

'*********
'Solenoids
'*********

SolCallback(1) = "SolBackPopper"
SolCallback(2) = "bsRP.SolOut"
SolCallback(3) = "SolDJMouth"
SolCallback(4) = "SolDJPopper"
SolCallback(5) = "SolDancinDummy"
SolCallback(6) = "SolComicMouth"
SolCallback(7) = "vpmSolSound ""Knocker"","
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

'**************
' Flipper Subs
'**************

SolCallback(sLRFlipper) = "SolRFlipper"
SolCallback(sLLFlipper) = "SolLFlipper"

Sub SolLFlipper(Enabled)
    If Enabled Then
        PlaySound SoundFX("fx_flipperup", DOFFlippers), 0, 1, -0.1, 0.05
        LeftFlipper.RotateToEnd
    Else
        PlaySound SoundFX("fx_flipperdown", DOFFlippers), 0, 1, -0.1, 0.05
        LeftFlipper.RotateToStart
    End If
End Sub

Sub SolRFlipper(Enabled)
    If Enabled Then
        PlaySound SoundFX("fx_flipperup", DOFFlippers), 0, 1, 0.1, 0.05
        RightFlipper.RotateToEnd
    Else
        PlaySound SoundFX("fx_flipperdown", DOFFlippers), 0, 1, 0.1, 0.05
        RightFlipper.RotateToStart
    End If
End Sub

Sub LeftFlipper_Collide(parm)
    PlaySound "fx_rubber_flipper", 0, parm / 10, -0.1, 0.25
End Sub

Sub RightFlipper_Collide(parm)
    PlaySound "fx_rubber_flipper", 0, parm / 10, 0.1, 0.25
End Sub

'**************
' Solenoid Subs
'**************

Sub SolDJPopper(Enabled)
    If Enabled Then
        bsDJ.SolOut 1
        sw67Metal.IsDropped = 0
        sw67metal.TimerEnabled = 1
    End If
End Sub

Sub sw67metal_Timer:sw67metal.IsDropped = 1:Me.TimerEnabled = 0:End Sub

' Dummy dancing

Dim cBall
DummyInit

Sub SolDancinDummy(Enabled) 'moves dummy up and down
    PlaySound "fx_dancing"
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
        PlaySound "fx_mouth", 0, 1, 0.05
        ComicMouth.Z = 75
    End If
End Sub

' Back Popper

Sub sw31_Hit()
    sw31.DestroyBall
    PlaySound "fx_kicker_enter"
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
        PlaySound "fx_kicker"
        Controller.Switch(41) = 0
        sw41.CreateBall
        sw41.Kick 190, 26
		RocketShake
        if Controller.Switch(31) then
            Controller.Switch(31) = 0
            Controller.Switch(41) = 1
        end if
    else
        PlaySound "fx_solenoid"
    end if
End Sub

' Rocket shaking

Dim cBall1
RocketInit

Sub RocketInit
    Set cBall1 = ckicker1.createball
    ckicker1.Kick 0, 0
End Sub

Sub RocketShake
    cball1.velx = 2 * RND(1)
    cball1.vely = -5 + 2 * RND(1)
End Sub

Sub RocketUpdate 'from the GameTimer
    Dim a, b
    a = ckicker1.y - cball1.y
    b = cball1.x - ckicker1.x
    Rocket.Rotz = 10 + a
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
        PlaySound "DJMouth"
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
        PlaySound "fx_DJMotor"
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
    if djdest = 4 then targetz = 10 end if
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

'***************************************************
'       JP's VP10 Fading Lamps & Flashers
'       Based on PD's Fading Light System
' SetLamp 0 is Off
' SetLamp 1 is On
' fading for non opacity objects is 4 steps
'***************************************************

Dim LampState(200), FadingLevel(200)
Dim FlashSpeedUp(200), FlashSpeedDown(200), FlashMin(200), FlashMax(200), FlashLevel(200), FlashRepeat(200)

InitLamps()             ' turn off the lights and flashers and reset them to the default parameters
LampTimer.Interval = 10 ' lamp fading speed
LampTimer.Enabled = 1

' Lamp & Flasher Timers

Sub LampTimer_Timer()
    Dim chgLamp, num, chg, ii
    chgLamp = Controller.ChangedLamps
    If Not IsEmpty(chgLamp) Then
        For ii = 0 To UBound(chgLamp)
            LampState(chgLamp(ii, 0) ) = chgLamp(ii, 1)       'keep the real state in an array
            FadingLevel(chgLamp(ii, 0) ) = chgLamp(ii, 1) + 4 'actual fading step
        Next
    End If
    UpdateLamps
End Sub

Sub UpdateLamps()
    NFadeL 11, li11
    NFadeL 12, li12
    NFadeL 13, li13
    NFadeL 14, li14
    NFadeL 15, li15
    NFadeL 16, li16
    NFadeL 17, li17
    NFadeL 18, li18
    NFadeL 21, li21
    NFadeL 22, li22
    NFadeL 23, li23
    NFadeL 24, li24
    NFadeL 25, li25
    NFadeL 26, li26
    NFadeL 27, li27
    NFadeL 28, li28
    NFadeL 31, li31
    NFadeL 32, li32
    NFadeL 33, li33
    NFadeL 34, li34
    NFadeL 35, li35
    NFadeL 36, li36
    NFadeL 37, li37
    NFadeL 38, li38
    NFadeL 41, li41
    NFadeL 42, li42
    NFadeL 43, li43
    NFadeL 44, li44
    NFadeL 45, li45
    NFadeL 46, li46
    NFadeL 47, li47
    NFadeL 48, li48
    NFadeL 51, li51
    NFadeL 52, li52
    NFadeL 53, li53
    NFadeL 54, li54
    NFadeL 55, li55
    NFadeL 56, li56
    NFadeL 57, li57
    NFadeL 58, li58
    NFadeL 61, li61
    NFadeL 62, li62
    NFadeL 63, li63
    NFadeL 64, li64
    NFadeL 65, li65
    NFadeL 66, li66
    NFadeL 67, li67
    NFadeL 68, li68
    NFadeL 71, li71
    NFadeL 72, li72
    NFadeL 73, li73
    NFadeL 74, li74
    FadeObj 75, b1p, "bumper_on", "bumper_a", "bumper_b", "bumper_off"
    FadeObj 76, b2p, "bumper_on", "bumper_a", "bumper_b", "bumper_off"
    FadeObj 77, b3p, "bumper_on", "bumper_a", "bumper_b", "bumper_off"
    'NFadeL 78, l78 'credit button
    NFadeL 81, li81
    NFadeL 82, li82
    Flash 83, li83
    Flash 84, li84
    Flash 85, li85
    NFadeL 86, li86

    'flashers

    NFadeLm 117, f17a
    NFadeL 117, f17
    Flash 118, f18
    NFadeL 119, f19
    NFadeL 120, f20
    Flash 121, f21
    NFadeLm 122, f22a
    NFadeL 122, f22
    NFadeL 125, f25
    NFadeLm 126, f26a
    NFadeLm 126, f26b
    NFadeL 126, f26
    NFadeLm 127, f27a
    NFadeLm 127, f27b
    NFadeL 127, f27
    NFadeL 128, f28
End Sub

' div lamp subs

Sub InitLamps()
    Dim x
    For x = 0 to 200
        LampState(x) = 0        ' current light state, independent of the fading level. 0 is off and 1 is on
        FadingLevel(x) = 4      ' used to track the fading state
        FlashSpeedUp(x) = 0.2   ' faster speed when turning on the flasher
        FlashSpeedDown(x) = 0.1 ' slower speed when turning off the flasher
        FlashMax(x) = 1         ' the maximum value when on, usually 1
        FlashMin(x) = 0         ' the minimum value when off, usually 0
        FlashLevel(x) = 0       ' the intensity of the flashers, usually from 0 to 1
        FlashRepeat(x) = 20     ' how many times the flash repeats
    Next
End Sub

Sub AllLampsOff
    Dim x
    For x = 0 to 200
        SetLamp x, 0
    Next
End Sub

Sub SetLamp(nr, value)
    If value <> LampState(nr) Then
        LampState(nr) = abs(value)
        FadingLevel(nr) = abs(value) + 4
    End If
End Sub

' Lights: used for VP10 standard lights, the fading is handled by VP itself

Sub NFadeL(nr, object)
    Select Case FadingLevel(nr)
        Case 4:object.state = 0:FadingLevel(nr) = 0
        Case 5:object.state = 1:FadingLevel(nr) = 1
    End Select
End Sub

Sub NFadeLm(nr, object) ' used for multiple lights
    Select Case FadingLevel(nr)
        Case 4:object.state = 0
        Case 5:object.state = 1
    End Select
End Sub

'Lights, Ramps & Primitives used as 4 step fading lights
'a,b,c,d are the images used from on to off

Sub FadeObj(nr, object, a, b, c, d)
    Select Case FadingLevel(nr)
        Case 4:object.image = b:FadingLevel(nr) = 6                   'fading to off...
        Case 5:object.image = a:FadingLevel(nr) = 1                   'ON
        Case 6, 7, 8:FadingLevel(nr) = FadingLevel(nr) + 1            'wait
        Case 9:object.image = c:FadingLevel(nr) = FadingLevel(nr) + 1 'fading...
        Case 10, 11, 12:FadingLevel(nr) = FadingLevel(nr) + 1         'wait
        Case 13:object.image = d:FadingLevel(nr) = 0                  'Off
    End Select
End Sub

Sub FadeObjm(nr, object, a, b, c, d)
    Select Case FadingLevel(nr)
        Case 4:object.image = b
        Case 5:object.image = a
        Case 9:object.image = c
        Case 13:object.image = d
    End Select
End Sub

Sub NFadeObj(nr, object, a, b)
    Select Case FadingLevel(nr)
        Case 4:object.image = b:FadingLevel(nr) = 0 'off
        Case 5:object.image = a:FadingLevel(nr) = 1 'on
    End Select
End Sub

Sub NFadeObjm(nr, object, a, b)
    Select Case FadingLevel(nr)
        Case 4:object.image = b
        Case 5:object.image = a
    End Select
End Sub

' Flasher objects

Sub Flash(nr, object)
    Select Case FadingLevel(nr)
        Case 4 'off
            FlashLevel(nr) = FlashLevel(nr) - FlashSpeedDown(nr)
            If FlashLevel(nr) <FlashMin(nr) Then
                FlashLevel(nr) = FlashMin(nr)
                FadingLevel(nr) = 0 'completely off
            End if
            Object.IntensityScale = FlashLevel(nr)
        Case 5 ' on
            FlashLevel(nr) = FlashLevel(nr) + FlashSpeedUp(nr)
            If FlashLevel(nr)> FlashMax(nr) Then
                FlashLevel(nr) = FlashMax(nr)
                FadingLevel(nr) = 1 'completely on
            End if
            Object.IntensityScale = FlashLevel(nr)
    End Select
End Sub

Sub Flashm(nr, object) 'multiple flashers, it doesn't change anything, it just follows the main flasher
    Select Case FadingLevel(nr)
        Case 4, 5
            Object.IntensityScale = FlashLevel(nr)
    End Select
End Sub

Sub FlashBlink(nr, object)
    Select Case FadingLevel(nr)
        Case 4 'off
            FlashLevel(nr) = FlashLevel(nr) - FlashSpeedDown(nr)
            If FlashLevel(nr) <FlashMin(nr) Then
                FlashLevel(nr) = FlashMin(nr)
                FadingLevel(nr) = 0 'completely off
            End if
            Object.IntensityScale = FlashLevel(nr)
            If FadingLevel(nr) = 0 AND FlashRepeat(nr) Then 'repeat the flash
                FlashRepeat(nr) = FlashRepeat(nr) -1
                If FlashRepeat(nr) Then FadingLevel(nr) = 5
            End If
        Case 5 ' on
            FlashLevel(nr) = FlashLevel(nr) + FlashSpeedUp(nr)
            If FlashLevel(nr)> FlashMax(nr) Then
                FlashLevel(nr) = FlashMax(nr)
                FadingLevel(nr) = 1 'completely on
            End if
            Object.IntensityScale = FlashLevel(nr)
            If FadingLevel(nr) = 1 AND FlashRepeat(nr) Then FadingLevel(nr) = 4
    End Select
End Sub

' Desktop Objects: Reels & texts (you may also use lights on the desktop)

' Reels

Sub FadeR(nr, object)
    Select Case FadingLevel(nr)
        Case 4:object.SetValue 1:FadingLevel(nr) = 6                   'fading to off...
        Case 5:object.SetValue 0:FadingLevel(nr) = 1                   'ON
        Case 6, 7, 8:FadingLevel(nr) = FadingLevel(nr) + 1             'wait
        Case 9:object.SetValue 2:FadingLevel(nr) = FadingLevel(nr) + 1 'fading...
        Case 10, 11, 12:FadingLevel(nr) = FadingLevel(nr) + 1          'wait
        Case 13:object.SetValue 3:FadingLevel(nr) = 0                  'Off
    End Select
End Sub

Sub FadeRm(nr, object)
    Select Case FadingLevel(nr)
        Case 4:object.SetValue 1
        Case 5:object.SetValue 0
        Case 9:object.SetValue 2
        Case 3:object.SetValue 3
    End Select
End Sub

'Texts, AudioFade(ActiveBall)

Sub NFadeT(nr, object, message)
    Select Case FadingLevel(nr)
        Case 4:object.Text = "":FadingLevel(nr) = 0
        Case 5:object.Text = message:FadingLevel(nr) = 1
    End Select
End Sub

Sub NFadeTm(nr, object, message)
    Select Case FadingLevel(nr)
        Case 4:object.Text = ""
        Case 5:object.Text = message
    End Select
End Sub

'******************************
' Diverse Collection Hit Sounds
'******************************

Sub aMetals_Hit(idx):PlaySound "fx_MetalHit2", 0, Vol(ActiveBall), pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0, AudioFade(ActiveBall):End Sub
Sub aRubber_Bands_Hit(idx):PlaySound "fx_rubber_band", 0, Vol(ActiveBall), pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0, AudioFade(ActiveBall):End Sub
Sub aRubber_Posts_Hit(idx):PlaySound "fx_rubber_post", 0, Vol(ActiveBall), pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0, AudioFade(ActiveBall):End Sub
Sub aRubber_Pins_Hit(idx):PlaySound "fx_rubber_pin", 0, Vol(ActiveBall), pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0, AudioFade(ActiveBall):End Sub
Sub aPlastics_Hit(idx):PlaySound "fx_PlasticHit", 0, Vol(ActiveBall), pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0, AudioFade(ActiveBall):End Sub
Sub aGates_Hit(idx):PlaySound "fx_Gate", 0, Vol(ActiveBall), pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0, AudioFade(ActiveBall):End Sub
Sub aWoods_Hit(idx):PlaySound "fx_Woodhit", 0, Vol(ActiveBall), pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0, AudioFade(ActiveBall):End Sub

'*****************************************
'			FLIPPER SHADOWS
'*****************************************
'
'sub FlipperTimer_Timer()
'	FlipperLSh.RotZ = LeftFlipper.currentangle
'	FlipperRSh.RotZ = RightFlipper.currentangle
'
'End Sub

'*****************************************
'			BALL SHADOW
'*****************************************
Dim BallShadow
BallShadow = Array (BallShadow1,BallShadow2,BallShadow3,BallShadow4,BallShadow5,BallShadow6,BallShadow7,BallShadow8)

Sub BallShadowUpdate_timer()
    Dim BOT, b
    BOT = GetBalls
    ' hide shadow of deleted balls
    If UBound(BOT)<(tnob-1) Then
        For b = (UBound(BOT) + 1) to (tnob-1)
            BallShadow(b).visible = 0
        Next
    End If
    ' exit the Sub if no balls on the table
    If UBound(BOT) = -1 Then Exit Sub
    ' render the shadow for each ball
    For b = 0 to UBound(BOT)
        If BOT(b).X < Table1.Width/2 Then
            BallShadow(b).X = ((BOT(b).X) - (Ballsize/6) + ((BOT(b).X - (Table1.Width/2))/7)) '+ 13
        Else
            BallShadow(b).X = ((BOT(b).X) + (Ballsize/6) + ((BOT(b).X - (Table1.Width/2))/7)) '- 13
        End If
        ballShadow(b).Y = BOT(b).Y + 10
        If BOT(b).Z > 20 Then
            BallShadow(b).visible = 1
        Else
            BallShadow(b).visible = 0
        End If
    Next
End Sub

' *******************************************************************************************************
' Positional Sound Playback Functions by DJRobX
' PlaySound sound, 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 0, 1, AudioFade(ActiveBall)
' *******************************************************************************************************

' Play a sound, depending on the X,Y position of the table element (especially cool for surround speaker setups, otherwise stereo panning only)
' parameters (defaults): loopcount (1), volume (1), randompitch (0), pitch (0), useexisting (0), restart (1))
' Note that this will not work (currently) for walls/slingshots as these do not feature a simple, single X,Y position

Sub PlayXYSound(soundname, tableobj, loopcount, volume, randompitch, pitch, useexisting, restart)
  PlaySound soundname, loopcount, volume, AudioPan(tableobj), randompitch, pitch, useexisting, restart, AudioFade(tableobj)
End Sub

' Set position as table object (Use object or light but NOT wall) and Vol to 1

Sub PlaySoundAt(soundname, tableobj)
  PlaySound soundname, 1, 1, AudioPan(tableobj), 0,0,0, 1, AudioFade(tableobj)
End Sub

'Set all as per ball position & speed.

Sub PlaySoundAtBall(soundname)
  PlaySoundAt soundname, ActiveBall
End Sub

'Set position as table object and Vol manually.

Sub PlaySoundAtVol(sound, tableobj, Vol)
  PlaySound sound, 1, Vol, Pan(tableobj), 0,0,0, 1, AudioFade(tableobj)
End Sub

'Set all as per ball position & speed, but Vol Multiplier may be used eg; PlaySoundAtBallVol "sound",3

Sub PlaySoundAtBallVol(sound, VolMult)
  PlaySound sound, 0, Vol(ActiveBall) * VolMult, Pan(ActiveBall), 0, Pitch(ActiveBall), 0, 1, AudioFade(ActiveBall)
End Sub

'Set position as bumperX and Vol manually.

Sub PlaySoundAtBumperVol(sound, tableobj, Vol)
  PlaySound sound, 1, Vol, Pan(tableobj), 0,0,1, 1, AudioFade(tableobj)
End Sub

'*********************************************************************
'                     Supporting Ball & Sound Functions
'*********************************************************************

Function AudioFade(tableobj) ' Fades between front and back of the table (for surround systems or 2x2 speakers, etc), depending on the Y position on the table. "table1" is the name of the table
  Dim tmp
  tmp = tableobj.y * 2 / table1.height-1
  If tmp > 0 Then
    AudioFade = Csng(tmp ^10)
  Else
    AudioFade = Csng(-((- tmp) ^10) )
  End If
End Function

Function AudioPan(tableobj) ' Calculates the pan for a tableobj based on the X position on the table. "table1" is the name of the table
  Dim tmp
  tmp = tableobj.x * 2 / table1.width-1
  If tmp > 0 Then
    AudioPan = Csng(tmp ^10)
  Else
    AudioPan = Csng(-((- tmp) ^10) )
  End If
End Function

Function Pan(ball) ' Calculates the pan for a ball based on the X position on the table. "table1" is the name of the table
    Dim tmp
    tmp = ball.x * 2 / table1.width-1
    If tmp > 0 Then
        Pan = Csng(tmp ^10)
    Else
        Pan = Csng(-((- tmp) ^10) )
    End If
End Function

Function AudioFade(ball) ' Can this be together with the above function ?
  Dim tmp
  tmp = ball.y * 2 / Table1.height-1
  If tmp > 0 Then
    AudioFade = Csng(tmp ^10)
  Else
    AudioFade = Csng(-((- tmp) ^10) )
  End If
End Function

Function Vol(ball) ' Calculates the Volume of the sound based on the ball speed
  Vol = Csng(BallVel(ball) ^2 / 2000)
End Function

Function Pitch(ball) ' Calculates the pitch of the sound based on the ball speed
  Pitch = BallVel(ball) * 20
End Function

Function BallVel(ball) 'Calculates the ball speed
  BallVel = INT(SQR((ball.VelX ^2) + (ball.VelY ^2) ) )
End Function


'*****************************************
'      JP's VP10 Rolling Sounds
'*****************************************

Const tnob = 8 ' total number of balls
Const lob = 2   'number of locked balls
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

    ' stop the sound of deleted balls
    For b = UBound(BOT) + 1 to tnob
        rolling(b) = False
        StopSound("fx_ballrolling" & b)
    Next

    ' exit the sub if no balls on the table
    If UBound(BOT) = lob - 1 Then Exit Sub 'no rolling sound for the captive balls

    ' play the rolling sound for each ball

    For b = 0 to UBound(BOT)
      If BallVel(BOT(b) ) > 1 Then
        rolling(b) = True
        if BOT(b).z < 30 Then ' Ball on playfield
          PlaySound("fx_ballrolling" & b), -1, Vol(BOT(b) ), Pan(BOT(b) ), 0, Pitch(BOT(b) ), 1, 0, AudioFade(BOT(b) )
        Else ' Ball on raised ramp
          PlaySound("fx_ballrolling" & b), -1, Vol(BOT(b) )*.5, Pan(BOT(b) ), 0, Pitch(BOT(b) )+50000, 1, 0, AudioFade(BOT(b) )
        End If
      Else
        If rolling(b) = True Then
          StopSound("fx_ballrolling" & b)
          rolling(b) = False
        End If
      End If
    Next
End Sub

'**********************
' Ball Collision Sound
'**********************

Sub OnBallBallCollision(ball1, ball2, velocity)
  If Table1.VersionMinor > 3 OR Table1.VersionMajor > 10 Then
    PlaySound("fx_collide"), 0, Csng(velocity) ^2 / 200, Pan(ball1), 0, Pitch(ball1), 0, 0, AudioFade(ball1)
  Else
    PlaySound("fx_collide"), 0, Csng(velocity) ^2 / 200, Pan(ball1), 0, Pitch(ball1), 0, 0
  End if
End Sub


' Tales of the Arabian Nights / IPD No. 3824 / May, 1996 / 4 Players
' Parts of the script taken from the older tables by PacDude, Pinball Ken, Kristian, Freylis, Fuseball, Guittar and Wpcmame
' VPX by JPSalas 2015
' DOF by arngrim

Option Explicit
Randomize

' Thalamus 2018-07-24
' Added/Updated "Positional Sound Playback Functions" and "Supporting Ball & Sound Functions"
' Changed UseSolenoids=1 to 2
' Thalamus 2018-08-18 : Improved directional sounds

Const VolDiv = 2000

Const VolBump   = 2    ' Bumpers multiplier.
Const VolGates  = 1    ' Gates volume multiplier.
Const VolMetals = 1    ' Metals volume multiplier.
Const VolRH     = 1    ' Rubber hits multiplier.
Const VolRPo    = 1    ' Rubber posts multiplier.
Const VolPlast  = 1    ' Plastics multiplier.
Const VolWood   = 1    ' Woods multiplier.

Const VolSpin   = 1.5  ' Spinners volume.
Const VolFlip   = 1    ' Flipper volume.


On Error Resume Next
ExecuteGlobal GetTextFile("controller.vbs")
If Err Then MsgBox "You need the controller.vbs in order to run this table, available in the vp10 package"
On Error Goto 0

Const BallSize = 50
Const UseVPMModSol = 1

Dim DesktopMode:DesktopMode = Table1.ShowDT
Dim UseVPMColoredDMD

UseVPMColoredDMD = DesktopMode

LoadVPM "01560000", "WPC.VBS", 3.46

'********************
'Standard definitions
'********************

Const UseSolenoids = 2
Const cSingleLFlip = 0
Const cSingleRFlip = 0
Const UseLamps = 0
Const UseSync = 0
Const HandleMech = 0

' Standard Sounds
Const SSolenoidOn = "fx_solenoid"
Const SSolenoidOff = ""
Const SCoin = "fx_Coin"

Const cGameName = "totan_14"

Dim bsTrough, bsBazaar, bsL, vlLock, cbLeft, cbRight, mVanishMagnet, mLockMagnet, mRampMagnet

Dim x

Set GICallback2 = GetRef("UpdateGI")
Set MotorCallback = GetRef("RollingUpdate") 'realtime updates - rolling sound

'************
' Table init.
'************

Sub Table1_Init
    vpmInit Me
    With Controller
        .GameName = cGameName
        If Err Then MsgBox "Can't start Game " & cGameName & vbNewLine & Err.Description:Exit Sub
        .SplashInfoLine = "Tales of the Arabian Nights - Williams 1996" & vbNewLine & "VPX by JPSalas"
        .Games(cGameName).Settings.Value("rol") = 0 'set it to 1 to rotate the DMD to the left
        .HandleKeyboard = 0
        .ShowTitle = 0
        .ShowDMDOnly = 1
        .ShowFrame = 0
        .HandleMechanics = 0
        .Hidden = DesktopMode
        On Error Resume Next
        .Run GetPlayerHWnd
        If Err Then MsgBox Err.Description
        On Error Goto 0
        .Switch(22) = 1 'close coin door
        .Switch(24) = 1 'and keep it close
    End With

    ' Nudging
    vpmNudge.TiltSwitch = 14
    vpmNudge.Sensitivity = 1
    vpmNudge.TiltObj = Array(bumper1, bumper2, bumper3, LeftSlingshot, RightSlingshot)

    ' Trough
    Set bsTrough = New cvpmTrough
    With bsTrough
        .size = 4
        .initSwitches Array(32, 33, 34, 35)
        .Initexit BallRelease, 90, 4
        .InitEntrySounds "fx_drain", "fx_Solenoid", "fx_Solenoid"
        .InitExitSounds SoundFX("fx_Solenoid",DOFContactors), SoundFX("fx_ballrel",DOFContactors)
        .Balls = 4
    End With

    'Bazaar hole
    Set bsBazaar = New cvpmTrough
    With bsBazaar
        .size = 4
        .initSwitches Array(25)
        .Initexit sw25, 192, 18
        .InitExitSounds SoundFX("fx_Popper",DOFContactors), SoundFx("fx_popper",DOFContactors)
        .InitExitVariance 2, 2
        .MaxBallsPerKick = 2
    End With

    ' Left hole
    Set bsL = New cvpmSaucer
    With bsL
        .InitKicker sw38, 38, 168, 18, 0
        .InitExitVariance 2, 2
        .InitSounds "fx_kicker-enter", "fx_Solenoid", SoundFX("fx_kicker",DOFContactors)
        .CreateEvents "bsL", sw38
    End With

    ' Ball Lock
    Set vlLock = New cvpmVLock
    With vlLock
        .InitVLock Array(TLock1, TLock2, TLock3), Array(Lock1, Lock2, Lock3), Array(66, 67, 68)
        .InitSnd "fx_kicker2", "fx_Solenoid"
        .CreateEvents "vlLock"
    End With

    Set cbLeft = New cvpmCaptiveBall
    With cbLeft
        .InitCaptive CapTrigger1, CapWall1, Array(CapKicker1, CapKicker1a), 352
        .NailedBalls = 1
        .ForceTrans = .9
        .MinForce = 3.5
        .CreateEvents "cbLeft"
        .Start
    End With
    CapKicker1.CreateBall

    Set cbRight = New cvpmCaptiveBall
    With cbRight
        .InitCaptive CapTrigger2, CapWall2, Array(CapKicker2, CapKicker2a), 10
        .NailedBalls = 1
        .ForceTrans = .9
        .MinForce = 3.5
        .CreateEvents "cbRight"
        .Start
    End With
    CapKicker2.CreateBall

    Set mVanishMagnet = New cvpmMagnet
    With mVanishMagnet
        .InitMagnet VanishMagnet, 30
        .GrabCenter = 1
        .CreateEvents "mVanishMagnet"
    End With

    Set mLockMagnet = New cvpmMagnet
    With mLockMagnet
        .InitMagnet LockMagnet, 66
        .Solenoid = 6
        .GrabCenter = 1
        .CreateEvents "mLockMagnet"
    End With

    Set mRampMagnet = New cvpmMagnet
    With mRampMagnet
        .InitMagnet RMagnet, 30
        .Solenoid = 8
        .GrabCenter = 0
        .CreateEvents "mRampMagnet"
    End With

    ' Main Timer init
    PinMAMETimer.Interval = PinMAMEInterval
    PinMAMETimer.Enabled = 1

    ' Init other dropwalls - animations
    SolSpikerLeft 0:SolSpikerRight 0
    HideVanish.isDropped = 0

    ' Init Lamp
    Dim obj
    For Each obj In colLampPoles:obj.IsDropped = 1:Next
    For Each obj In colLampPoles2:obj.IsDropped = 1:Next
    lampSpinSpeed = 0:lampLastPos = -1:SpinTimer.Enabled = True ' force update
    UpdateGI 0, 1:UpdateGI 1, 1:UpdateGI 2, 1

  ' Remove the cabinet rails if in FS mode
  If Table1.ShowDT = False then
    lrail.Visible = False
    rrail.Visible = False
    'ramp4.Visible = False
    'ramp5.Visible = False
  End If
End Sub

Sub table1_Paused:Controller.Pause = 1:End Sub
Sub table1_unPaused:Controller.Pause = 0:End Sub

'***********
' Update GI
'***********
Dim gistep

Sub UpdateGI(no, step)
    Dim ii
    If step = 0 then exit sub 'only values from 1 to 8 are visible and reliable. 0 is not reliable and 7 & 8 are the same so...
    gistep = (step-1) / 7
  DOF 200, gistep
    Select Case no
        Case 0
            For each ii in aGiLLights
                ii.IntensityScale = gistep
            Next
        Case 1
            For each ii in aGiMLights
                ii.IntensityScale = gistep
            'back.Image="backwall"&step
            Next
        Case 2 ' also the bumpers er GI
            For each ii in aGiTLights
                ii.IntensityScale = gistep
            Next
    End Select
    ' change the intensity of the flasher depending on the gi to compensate for the gi lights being off
    For ii = 0 to 200
        FlashMax(ii) = 6 - gistep * 3 ' the maximum value of the flashers
    Next
End Sub

'**********
' Keys
'**********
Sub table1_KeyDown(ByVal Keycode)
    If keycode = LeftTiltKey Then Nudge 90, 5:PlaySound SoundFX("fx_nudge",0), 0, 1, -0.1, 0.25
    If keycode = RightTiltKey Then Nudge 270, 5:PlaySound SoundFX("fx_nudge",0), 0, 1, 0.1, 0.25
    If keycode = CenterTiltKey Then Nudge 0, 6:PlaySound SoundFX("fx_nudge",0), 0, 1, 0, 0.25
    If vpmKeyDown(keycode) Then Exit Sub
    If keycode = PlungerKey Then PlaySoundAt "fx_PlungerPull",Plunger:Plunger.Pullback
End Sub

Sub table1_KeyUp(ByVal Keycode)
    If vpmKeyUp(keycode) Then Exit Sub
    If keycode = PlungerKey Then PLaySoundAt "fx_plunger",Plunger:Plunger.Fire
End Sub

'*********
' Switches
'*********

' Drain hole
Sub Drain_Hit:bsTrough.AddBall Me:End Sub

' Slings & div switches

Dim LStep, RStep

Sub LeftSlingShot_Slingshot
    PlaySoundAt SoundFx("fx_slingshot",DOFContactors), Lemk
    LeftSling3.Visible = 1
    Lemk.RotX = 26
    LStep = 0
    vpmTimer.PulseSw 51
    LeftSlingShot.TimerEnabled = 1
End Sub

Sub LeftSlingShot_Timer
    Select Case LStep
        Case 1:LeftSLing3.Visible = 0:LeftSLing2.Visible = 1:Lemk.RotX = 14
        Case 2:LeftSLing2.Visible = 0:LeftSLing1.Visible = 1:Lemk.RotX = 2
        Case 3:LeftSLing1.Visible = 0:Lemk.RotX = -10:LeftSlingShot.TimerEnabled = 0
    End Select
    LStep = LStep + 1
End Sub

Sub RightSlingShot_Slingshot
    PlaySoundAt SoundFx("fx_slingshot",DOFContactors), Remk
    RightSling3.Visible = 1
    Remk.RotX = 26
    RStep = 0
    vpmTimer.PulseSw 52
    RightSlingShot.TimerEnabled = 1
End Sub

Sub RightSlingShot_Timer
    Select Case RStep
        Case 1:RightSLing3.Visible = 0:RightSLing2.Visible = 1:Remk.RotX = 14
        Case 2:RightSLing2.Visible = 0:RightSLing1.Visible = 1:Remk.RotX = 2
        Case 3:RightSLing1.Visible = 0:Remk.RotX = -10:RightSlingShot.TimerEnabled = 0
    End Select
    RStep = RStep + 1
End Sub

' Bumpers
Sub Bumper1_Hit:vpmTimer.PulseSw 53:PlaySoundAtVol SoundFX("fx_bumper",DOFContactors), Bumper1, VolBump:End Sub
Sub Bumper2_Hit:vpmTimer.PulseSw 55:PlaySoundAtVol SoundFX("fx_bumper",DOFContactors), Bumper1, VolBump:End Sub
Sub Bumper3_Hit:vpmTimer.PulseSw 54:PlaySoundAtVol SoundFX("fx_bumper",DOFContactors), Bumper1, VolBump:End Sub

' Skillshot
Sub sw63_Hit():vpmTimer.PulseSw 63:PlaySoundAt "fx_SkillShot-Falling", sw63:End Sub
Sub sw64_Hit():vpmTimer.PulseSw 64:PlaySoundAt "fx_SkillShot-Falling", sw64:End Sub
Sub sw65_Hit():vpmTimer.PulseSw 65:PlaySoundAt "fx_SkillShot-Falling", sw65:End Sub

' Hole with animation
Dim aBall

Sub sw25_Hit
    PlaySoundAt "fx_hole-enter", sw25
    Set aBall = ActiveBall:Me.TimerEnabled = 1

    bsBazaar.AddBall 0
End Sub

Sub sw25_Timer
    Do While aBall.Z > 0
        aBall.Z = aBall.Z -5
        Exit Sub
    Loop
    Me.DestroyBall
    Me.TimerEnabled = 0
End Sub

Sub HaremSneak_Hit:PlaySoundAt "fx_hole-enter", HaremSneak:Me.DestroyBall:SubwayHandler 11:End Sub

' Rollovers & Ramp Switches
Sub sw26_Hit:Controller.Switch(26) = 1:PlaySoundAt "fx_sensor", sw26:End Sub
Sub sw26_UnHit:Controller.Switch(26) = 0:End Sub

Sub sw17_Hit:Controller.Switch(17) = 1:PlaySoundAt "fx_sensor", sw17:End Sub
Sub sw17_UnHit:Controller.Switch(17) = 0:End Sub

Sub sw27_Hit:Controller.Switch(27) = 1:SetLamp 130, 1:PlaySoundAt "fx_sensor", sw27:End Sub
Sub sw27_Unhit:Controller.Switch(27) = 0:SetLamp 130, 0:End Sub

Sub sw16_Hit:Controller.Switch(16) = 1:SetLamp 129, 1:PlaySoundAt "fx_sensor", sw16:End Sub
Sub sw16_Unhit:Controller.Switch(16) = 0:SetLamp 129, 0:End Sub

Sub sw15_Hit:Controller.Switch(15) = 1
    If ActiveBall.VelY < 0 Then 'on the way up
        PlaySoundAt "fx_rrenter", sw15
    End If
End Sub
Sub sw15_Unhit:Controller.Switch(15) = 0:End Sub

Sub sw28_Hit:Controller.Switch(28) = 1:PlaySoundAt "fx_sensor", sw28:End Sub
Sub sw28_Unhit:Controller.Switch(28) = 0:PlaySoundAt "fx_metalrolling",sw28:End Sub

Sub sw36_Hit:Controller.Switch(36) = 1:PlaySoundAt "fx_sensor", sw36:End Sub
Sub sw36_Unhit:Controller.Switch(36) = 0:End Sub

Sub sw37_Hit:Controller.Switch(37) = 1:PlaySoundAt "fx_sensor", sw37:End Sub
Sub sw37_Unhit:Controller.Switch(37) = 0:End Sub

Sub sw41_Hit:Controller.Switch(41) = 1:PlaySoundAt "fx_sensor", sw41:End Sub
Sub sw41_Unhit:Controller.Switch(41) = 0:End Sub

Sub sw43_Hit:Controller.Switch(43) = 1:PlaySoundAt "fx_sensor", sw43:End Sub
Sub sw43_Unhit:Controller.Switch(43) = 0:End Sub

Sub sw44_Hit:Controller.Switch(44) = 1:PlaySoundAt "fx_sensor", sw44:End Sub
Sub sw44_Unhit:Controller.Switch(44) = 0:End Sub

Sub sw45_Hit:Controller.Switch(45) = 1:PlaySoundAt "fx_sensor", sw45:End Sub
Sub sw45_Unhit:Controller.Switch(45) = 0:End Sub

Sub sw47_Hit:Controller.Switch(47) = 1:PlaySoundAt "fx_sensor", sw47:End Sub
Sub sw47_Unhit:Controller.Switch(47) = 0:End Sub

Sub sw18_Hit:Controller.Switch(18) = 1:PlaySoundAt "fx_sensor", sw18:End Sub
Sub sw18_Unhit:Controller.Switch(18) = 0:End Sub

' Targets
Sub sw23_Hit:vpmTimer.PulseSw 23:PlaySound SoundFX("fx_target",DOFContactors), 0, Vol(ActiveBall) * 10, pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0, AudioFade(ActiveBall):End Sub
Sub sw46_Hit:vpmTimer.PulseSw 46:PlaySound SoundFX("fx_target",DOFContactors), 0, Vol(ActiveBall) * 10, pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0, AudioFade(ActiveBall):DOF 101, DOFPulse:End Sub
Sub sw46b_Hit:vpmTimer.PulseSw 46:PlaySound SoundFX("fx_target",DOFContactors), 0, Vol(ActiveBall) * 10, pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0, AudioFade(ActiveBall):DOF 102, DOFPulse:End Sub
Sub sw48_Hit:vpmTimer.PulseSw 48:PlaySound SoundFX("fx_target",DOFContactors), 0, Vol(ActiveBall) * 10, pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0, AudioFade(ActiveBall):End Sub
Sub sw58_Hit:vpmTimer.PulseSw 58:PlaySound SoundFX("fx_target",DOFContactors), 0, Vol(ActiveBall) * 10, pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0, AudioFade(ActiveBall):End Sub
Sub sw61a_Hit:vpmTimer.PulseSw 61:PlaySound SoundFX("fx_target",DOFContactors), 0, Vol(ActiveBall) * 10, pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0, AudioFade(ActiveBall):End Sub
Sub sw61b_Hit:vpmTimer.PulseSw 61:PlaySound SoundFX("fx_target",DOFContactors), 0, Vol(ActiveBall) * 10, pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0, AudioFade(ActiveBall):End Sub
Sub sw61c_Hit:vpmTimer.PulseSw 61:PlaySound SoundFX("fx_target",DOFContactors), 0, Vol(ActiveBall) * 10, pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0, AudioFade(ActiveBall):End Sub
Sub sw62a_Hit:vpmTimer.PulseSw 62:PlaySound SoundFX("fx_target",DOFContactors), 0, Vol(ActiveBall) * 10, pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0, AudioFade(ActiveBall):End Sub
Sub sw62b_Hit:vpmTimer.PulseSw 62:PlaySound SoundFX("fx_target",DOFContactors), 0, Vol(ActiveBall) * 10, pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0, AudioFade(ActiveBall):End Sub
Sub sw62c_Hit:vpmTimer.PulseSw 62:PlaySound SoundFX("fx_target",DOFContactors), 0, Vol(ActiveBall) * 10, pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0, AudioFade(ActiveBall):End Sub

'********************
'  Genie animation
'********************
Dim GenioPos, GenioCount

Sub sw42_Hit:vpmTimer.PulseSw 42:GenioPos = -5:GenioCount = 15:sw42.TimerEnabled = 1:PlaySound SoundFX("fx_geniehit",DOFShaker), 0, Vol(ActiveBall) * 10, pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0, AudioFade(ActiveBall):End Sub

Sub sw42_timer
    Genio.RotX = GenioPos
    GenioCount = GenioCount -1
    If GenioCount <= 0 Then sw42.TimerEnabled = False:Exit Sub
    If GenioPos < 0 Then
        GenioPos = ABS(GenioPos) - 0.5
    Else
        GenioPos = - GenioPos + 0.5
    End If
End Sub

'*********
'Solenoids
'*********
SolCallback(1) = "SolSpikerLeft"
SolCallback(2) = "SolSpikerRight"
SolCallback(3) = "SolVanishDrop"
SolCallback(4) = "SolLockRelease"
SolCallback(5) = "bsBazaar.SolOut"
SolCallback(7) = "vpmSolSound SoundFX(""fx_Knocker"",DOFContactors),"
SolCallback(8) = "SolRampMagnet"
SolCallback(9) = "SolRelease"
SolCallback(15) = "bsL.SolOut"
SolCallback(21) = "RampDiverter"
SolCallback(34) = "vpmSolDiverter PlayFDiv, SoundFX(""fx_diverter"",DOFContactors),"
SolCallback(35) = "SolVanishMagnet"
SolCallback(36) = "vpmSolWall LoopPostDiverter,True,Not "

' Flashers
SolModCallback(16) = "SetModLamp 116,"
SolModCallback(17) = "SetModLamp 117,"
SolModCallback(18) = "SetModLamp 118,"
SolModCallback(19) = "SetModLamp 119,"
SolModCallback(20) = "SetModLamp 120,"
SolModCallback(22) = "SetModLamp 122,"
SolModCallback(23) = "SetModLamp 123,"
SolModCallback(24) = "SetModLamp 124,"
SolModCallback(25) = "SetModLamp 125,"
SolModCallback(26) = "SetModLamp 126,"
SolModCallback(27) = "SetModLamp 127,"
SolModCallback(28) = "SetModLamp 128,"

Sub SolRelease(Enabled)
    If Enabled And bsTrough.Balls > 0 Then
        vpmTimer.PulseSw 31
        bsTrough.ExitSol_On
    End If
End Sub

Sub SolLockRelease(enabled)
    vlLock.SolExit enabled
End Sub

Sub SolSpikerLeft(Enabled)
    vpmSolWall Array(s11, s12, s13, s14, s15, s16, s17, s18), False, NOT Enabled
    SetLamp 90, Enabled
    If Enabled Then
        PlaySound SoundFX("fx_cage-on",DOFContactors), 0, 1, -0.05
        PlaySound SoundFX("fx_cage-hold",DOFShaker), -1, 1, -0.05
    Else
        StopSound "fx_cage-hold"
        PlaySound SoundFX("fx_cage-off",DOFContactors), 0, 1, -0.05
    End If
End Sub

Sub SolSpikerRight(Enabled)
    vpmSolWall Array(s21, s22, s23, s24, s25, s26, s27, s28), False, NOT Enabled
    SetLamp 91, Enabled
    If Enabled Then
        PlaySound SoundFX("fx_cage-on",DOFContactors), 0, 1, 0.05
        PlaySound "fx_cage-hold", -1, 1, 0.05
    Else
        StopSound "fx_cage-hold"
        PlaySound SoundFX("fx_cage-off",DOFContactors), 0, 1, 0.05
    End If
End Sub

'********************
'    JP Flippers
'********************

SolCallback(sLRFlipper) = "SolRFlipper"
SolCallback(sLLFlipper) = "SolLFlipper"

Sub SolLFlipper(Enabled)
    If Enabled Then
        PlaySoundAtVol SoundFX("fx_flipperup",DOFContactors), LeftFlipper, VolFlip
        LeftFlipper.RotateToEnd
    Else
        PlaySoundAtVol SoundFX("fx_flipperdown",DOFContactors), LeftFlipper, VolFlip
        LeftFlipper.RotateToStart
    End If
End Sub

Sub SolRFlipper(Enabled)
    If Enabled Then
        PlaySoundAtVol SoundFX("fx_flipperup",DOFContactors), RightFlipper, VolFlip
        RightFlipper.RotateToEnd
    Else
        PlaySoundAtVol SoundFX("fx_flipperdown",DOFContactors), RightFlipper, VolFlip
        RightFlipper.RotateToStart
    End If
End Sub

Sub LeftFlipper_Collide(parm)
    PlaySound "fx_rubber_flipper", 0, parm / 10, -0.05, 0.15
End Sub

Sub RightFlipper_Collide(parm)
    PlaySound "fx_rubber_flipper", 0, parm / 10, 0.05, 0.15
End Sub

'***************
' Vanish Magnet
'***************

' Magnet power is pulsed so wait before turning power off
Sub SolVanishMagnet(enabled)
    HideVanish.TimerEnabled = Not enabled
    If enabled Then mVanishMagnet.MagnetOn = 1
End Sub

' Magnet is turned off
' Sends balls in random direction
Sub HideVanish_Timer
    Dim dir, speed, ball
    For Each ball In mVanishMagnet.Balls
        With ball
            If(.X - VanishHole.X) ^2 + (.Y - VanishHole.Y) ^2 < 15 * 15 Then
                dir = Rnd * 6.28:speed = 15 + Rnd * 5
                .VelX = speed * Sin(dir): .VelY = speed * Cos(dir)
            End If
        End With
    Next
    Me.TimerEnabled = False:mVanishMagnet.MagnetOn = False
End Sub

' When vanish hole opens move ball into kicker
Sub SolVanishDrop(enabled)
    VanishHole.Enabled = enabled
    HideVanish.Isdropped = enabled
    If enabled Then
        Dim ball
        For Each ball In mVanishMagnet.Balls
            With ball
                If(.X - vanishHole.X) ^2 + (.Y - vanishHole.Y) ^2 < 26 * 26 Then
                    .X = VanishHole.X: .Y = VanishHole.Y - 26: .VelY = 10: .VelX = 0
                    mVanishMagnet.RemoveBall ball
                    Exit For ' only one ball can be in the center
                End If
            End With
        Next
    End If
End Sub

' Ball has hit kicker, start animation
' and then drop it into the subway tunnel

Dim vBall
Sub VanishHole_Hit
    PlaySound "fx_balldrop", 0, 1, pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0, AudioFade(ActiveBall)
    Set vBall = ActiveBall:Me.TimerEnabled = 1
    vpmTimer.AddTimer 250, Me.Name & ".DestroyBall : : SubwayHandler 12 '"
End Sub

Sub VanishHole_Timer
    Do While vBall.Z > 0
        vBall.Z = vBall.Z -5
        Exit Sub
    Loop
    Me.TimerEnabled = 0
End Sub

'***************
' Handle Subway
'***************
Sub SubwayHandler(swNo)
    Select case swNo
        case 11
            vpmTimer.PulseSwitch 11, 2500, "bsBazaar.AddBall"
        case 12
            vpmTimer.PulseSwitch 12, 1500, "bsBazaar.AddBall"
    End Select
End Sub

'************************
'     Ramp Magnet
'************************

Sub SolRampMagnet(Enabled)
    RampMagnet.Enabled = Enabled:RampMagnet1.Enabled = Enabled
'If Not Enabled Then PlaySound "fx_motor"
End Sub

Sub RampMagnet_Hit:RampMagnet.Destroyball:vpmCreateBall RampMagnet2:RampMagnet2.Kick 90, 1:End Sub

Sub RampMagnet1_Hit:RampMagnet1.Destroyball:vpmCreateBall RampMagnet2:RampMagnet2.Kick 90, 1:End Sub

'*************************************************
' Lock (from below)
' If lock is empty, trap ball and add it to lock.
' Otherwise just let it bounce on the locked ball
'*************************************************

Sub THaremBounce_Hit:HaremBounceOut.Enabled = (ActiveBall.VelY < 0 And vlLock.Balls = 0):End Sub
Sub THaremBounce_UnHit:HaremBounceOut.Enabled = False:End Sub
Sub HaremBounceOut_Hit:PlaySound "fx_kicker-enter", 0, 1, pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0, AudioFade(ActiveBall):Me.DestroyBall:vlLock.Balls = vlLock.Balls + 1:End Sub

'***************
'Lamp Simulator
'***************

Dim lampPosition, lampSpinSpeed, lampLastPos
Const cLampSpeedMult = 180             ' 180 - Affects speed transfer to object (deg/sec)
Const cLampFriction = 1.5              ' 2.0 - Friction coefficient (deg/sec/sec)
Const cLampMinSpeed = 16               ' 20 - Object stops at this speed (deg/sec)
Const cLampRadius = 72
Const cBallSpeedDampeningEffect = 0.45 ' 45 - The ball retains this fraction of its speed due to energy absorption by hitting the lamp.

' Draw lamp
Sub SpinTimer_Timer
  DOF 103,DOFOn
    Dim curPos
    Dim oldLampSpeed:oldLampSpeed = lampSpinSpeed
    lampPosition = lampPosition + lampSpinSpeed * Me.Interval / 1000
    lampSpinSpeed = lampSpinSpeed * (1 - cLampFriction * Me.Interval / 1000)
    'added by jim - makes this sub a bit safer (even though this shouldnt happen in normal use)
    Do While lampPosition < 0
        lampPosition = lampPosition + 360
    Loop
    Do While lampPosition > 360
        lampPosition = lampPosition - 360
    Loop
    curPos = Int((lampPosition * colLampPoles.Count) / 360)
    If curPos <> lampLastPos Then
        LampPr.RotZ = 360 - lampPosition
        If lampLastPos >= 0 Then ' not first time
            colLampPoles(lampLastPos).IsDropped = True
            colLampPoles2(lampLastPos).IsDropped = True
        End If
        On Error Resume Next
        colLampPoles(curPos).IsDropped = False
        If Err Then msgbox curPoles
        colLampPoles2(curPos).IsDropped = False
        If oldLampSpeed > 0 And lampLastPos > curPos Then
            PlaySound SoundFX("fx_lamp",DOFGear)
            'rev anticlockwise
            vpmTimer.PulseSw 56 ' ? or 57 and the other is 56
        ElseIf oldLampSpeed < 0 And lampLastPos < curPos Then
            'rev clockwise
            PlaySound SoundFX("fx_lamp",DOFGear)
            vpmTimer.PulseSw 57
        End If
        lampLastPos = curPos
    End If
    If Abs(lampSpinSpeed) < cLampMinSpeed Then
        lampSpinSpeed = 0:Me.Enabled = False
    End If
  DOF 103, DOFOff
End Sub

Sub colLampPoles_Hit(idx)
    PlaySound "fx_lamphit", 0, Vol(ActiveBall) * 20, pan(ActiveBall), 0.25, Pitch(ActiveBall), 0, 0, AudioFade(ActiveBall)
    Dim pi:pi = 3.14159265358979323846
    dim lampangle
    lampangle = NormAngle((idx) / Me.Count * 2 * pi + (pi / 17.6) + pi) ' added (pi / 17.6) because the first lamp post (18) is not quite at 0 angle, added another pi to make it standard angle

    dim mball, mlamp, rlamp, ilamp

    dim collisionangle
    dim ballspeedin, ballspeedout, lampspeedin, lampspeedout

    dim fudge:fudge = cLampSpeedMult / 2

    With ActiveBall
        collisionangle = GetCollisionAngle(idx, Me.Count, .X, -.Y) ' this is the angle from the center of ball to the center of post

        Set ballspeedout = new jVector
        ballspeedout.SetXY .VelX, -.VelY
        ballspeedout.ShiftAxes - collisionangle

        lampSpinSpeed = lampSpinSpeed + sqr(.VelX ^2 + .VelY ^2) * sin(collisionangle - lampangle) * fudge
        ballspeedout.SetXY ballspeedout.x, ballspeedout.y * cos(collisionangle - lampangle)

        ballspeedout.ShiftAxes collisionangle
    End With
    SpinTimer.Enabled = True
End Sub

Function GetCollisionAngle(idx, count, X, Y)
    Dim pi:pi = 3.14159265358979323846
    dim angle, postx, posty, dX, dY
    Dim ang
    angle = (idx) / count * 2 * pi + (pi / 17.6) ' added (pi / 17.6) because the first lamp post (18) is not quite at 0 angle
    postx = 511.5 - 60.25 * Cos(angle)           ' he actual coordinates of the center of the lamp
    posty = 847 + 60.25 * Sin(angle)             ' 60.25 is the radius of the circle with center at the center of the lamp and edge at the centers of all the lamp posts
    posty = -1 * posty
    Dim collisionV:Set collisionV = new jVector
    collisionV.SetXY postx - X, posty - Y
    GetCollisionAngle = collisionV.ang
End Function

Function NormAngle(angle)
    NormAngle = angle
    Dim pi:pi = 3.14159265358979323846
    Do While NormAngle > 2 * pi
        NormAngle = NormAngle - 2 * pi
    Loop
    Do While NormAngle < 0
        NormAngle = NormAngle + 2 * pi
    Loop
End Function

Class jVector
    Private m_mag, m_ang, pi

    Sub Class_Initialize
        m_mag = CDbl(0)
        m_ang = CDbl(0)
        pi = CDbl(3.14159265358979323846)
    End Sub

    Public Function add(anothervector)
        Dim tx, ty, theta
        If TypeName(anothervector) = "jVector" then
            Set add = new jVector
            add.SetXY x + anothervector.x, y + anothervector.y
        End If
    End Function

    Public Function multiply(scalar)
        Set multiply = new jVector
        multiply.SetXY x * scalar, y * scalar
    End Function

    Sub ShiftAxes(theta)
        ang = ang - theta
    end Sub

    Sub SetXY(tx, ty)

        if tx = 0 And ty = 0 Then
            ang = 0
        elseif tx = 0 And ty < 0 then
            ang = - pi / 180 ' -90 degrees
        elseif tx = 0 And ty > 0 then
            ang = pi / 180   ' 90 degrees
        else
            ang = atn(ty / tx)
            if tx < 0 then ang = ang + pi ' Add 180 deg if in quadrant 2 or 3
        End if

        mag = sqr(tx ^2 + ty ^2)
    End Sub

    Property Let mag(nmag)
        m_mag = nmag
    End Property

    Property Get mag
        mag = m_mag
    End Property

    Property Let ang(nang)
        m_ang = nang
        Do While m_ang > 2 * pi
            m_ang = m_ang - 2 * pi
        Loop
        Do While m_ang < 0
            m_ang = m_ang + 2 * pi
        Loop
    End Property

    Property Get ang
        Do While m_ang > 2 * pi
            m_ang = m_ang - 2 * pi
        Loop
        Do While m_ang < 0
            m_ang = m_ang + 2 * pi
        Loop
        ang = m_ang
    End Property

    Property Get x
        x = m_mag * cos(ang)
    End Property

    Property Get y
        y = m_mag * sin(ang)
    End Property

    Property Get dump
        dump = "vector "
        Select Case CInt(ang + pi / 8)
            case 0, 8:dump = dump & "->"
            case 1:dump = dump & "/'"
            case 2:dump = dump & "/\"
            case 3:dump = dump & "'\"
            case 4:dump = dump & "<-"
            case 5:dump = dump & ":/"
            case 6:dump = dump & "\/"
            case 7:dump = dump & "\:"
        End Select
        dump = dump & " mag:" & CLng(mag * 10) / 10 & ", ang:" & CLng(ang * 180 / pi) & ", x:" & CLng(x * 10) / 10 & ", y:" & CLng(y * 10) / 10
    End Property
End Class

'***************************************************
'       JP's VP10 Fading Lamps & Flashers
'       Based on PD's Fading Light System
' SetLamp 0 is Off
' SetLamp 1 is On
' fading for non opacity objects is 4 steps
'***************************************************

Dim LampState(200), FadingLevel(200)
Dim FlashSpeedUp(200), FlashSpeedDown(200), FlashMin(200), FlashMax(200), FlashLevel(200)

InitLamps()             ' turn off the lights and flashers and reset them to the default parameters
LampTimer.Interval = 10 'lamp fading speed
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

Sub UpdateLamps
    NFadeL 11, l11
    NFadeL 12, l12
    NFadeL 13, l13
    NFadeL 14, l14
    NFadeL 15, l15
    NFadeL 15, l15
    NFadeL 16, l16
    NFadeL 17, l17
    NFadeL 18, l18
    NFadeL 21, l21
    NFadeL 22, l22
    NFadeL 23, l23
    NFadeL 24, l24
    NFadeL 25, l25
    NFadeL 26, l26
    NFadeL 27, l27
    NFadeLm 28, l28a
    NFadeL 28, l28
    NFadeL 31, l31
    NFadeL 32, l32
    NFadeL 33, l33
    NFadeL 34, l34
    NFadeL 35, l35
    NFadeL 36, l36
    NFadeL 37, l37
    NFadeL 38, l38
    NFadeL 41, l41
    NFadeL 42, l42
    NFadeL 43, l43
    NFadeL 44, l44
    NFadeL 45, l45
    NFadeL 46, l46
    NFadeL 47, l47
    NFadeL 48, l48
    NFadeL 51, l51
    NFadeL 52, l52
    NFadeL 53, l53
    NFadeL 54, l54
    NFadeL 55, l55
    NFadeL 56, l56
    NFadeL 57, l57
    NFadeL 58, l58
    NFadeL 61, l61
    NFadeL 62, l62
    NFadeL 63, l63
    NFadeL 64, l64
    NFadeL 65, l65
    NFadeL 66, l66
    NFadeL 67, l67
    NFadeL 68, l68
    NFadeL 71, l71
    NFadeL 72, l72
    NFadeL 73, l73
    NFadeL 74, l74
    NFadeL 75, l75
    NFadeL 76, l76
    NFadeL 77, l77
    NFadeL 78, l78
    NFadeL 81, l81
    NFadeL 82, l82
    NFadeL 83, l83
    NFadeL 84, l84
    NFadeL 85, l85
    NFadeL 86, l86
    NFadeL 87, l87
    NFadeL 88, l88

    'Spikers
    NFadeL 90, f36
    NFadeL 91, f37
    LightMod 116, f16b
    LightMod 116, f16
    FlashMod 116, f16a
    FlashMod 116, f16c
    LightMod 117, f17a
    LightMod 117, f17b
    FlashMod 117, f17c
    FlashMod 117, f17e
    FlashMod 117, f17f
    FlashMod 117, f17d
    LightMod 118, f18
    LightMod 119, f19
    FlashMod 119, f19a
    FlashMod 119, f19b
    FlashMod 120, f20
    LightMod 122, f22
    LightMod 123, f23
    LightMod 124, f24
    LightMod 125, f25
    LightMod 126, f26
    FlashMod 126, f26a
    FlashMod 126, f26b
    LightMod 127, f27
    FlashMod 127, f27b
    FlashMod 127, f27d
    FlashMod 127, f27a
    LightMod 128, f28
    NFadeL 130, sw27l
    NFadeL 129, sw16l
End Sub

Sub InitLamps()
    Dim x
    For x = 0 to 200
        LampState(x) = 0         ' current light state, independent of the fading level. 0 is off and 1 is on
        FadingLevel(x) = 4       ' used to track the fading state
        FlashSpeedUp(x) = 0.5    ' faster speed when turning on the flasher
        FlashSpeedDown(x) = 0.35 ' slower speed when turning off the flasher
        FlashMax(x) = 1          ' the maximum value when on, usually 1
        FlashMin(x) = 0          ' the minimum value when off, usually 0
        FlashLevel(x) = 0        ' the intensity of the flashers, usually from 0 to 1
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

Sub SetModLamp(nr, level)
  FlashLevel(nr) = level /150 'lights & flashers
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

Sub LightMod(nr, object) ' modulated lights used as flashers
    Object.IntensityScale = FlashLevel(nr)
  Object.State = 1
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
            If FlashLevel(nr) < FlashMin(nr) Then
                FlashLevel(nr) = FlashMin(nr)
                FadingLevel(nr) = 0 'completely off
            End if
            Object.IntensityScale = FlashLevel(nr)
        Case 5 ' on
            FlashLevel(nr) = FlashLevel(nr) + FlashSpeedUp(nr)
            If FlashLevel(nr) > FlashMax(nr) Then
                FlashLevel(nr) = FlashMax(nr)
                FadingLevel(nr) = 1 'completely on
            End if
            Object.IntensityScale = FlashLevel(nr)
    End Select
End Sub

Sub Flashm(nr, object) 'multiple flashers, it just sets the flashlevel
    Object.IntensityScale = FlashLevel(nr)
End Sub

Sub FlashMod(nr, object) 'sets the flashlevel from the SolModCallback
    Object.IntensityScale = FlashLevel(nr)
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

'Texts

Sub NFadeT(nr, object, message)
    Select Case FadingLevel(nr)
        Case 4:object.Text = "":FadingLevel(nr) = 0
        Case 5:object.Text = message:FadingLevel(nr) = 1
    End Select
End Sub

Sub NFadeTm(nr, object, b)
    Select Case FadingLevel(nr)
        Case 4:object.Text = ""
        Case 5:object.Text = message
    End Select
End Sub

'******************************
' Diverse Collection Hit Sounds
'******************************

Sub aRubbers_Hit(idx):PlaySound "fx_rubber", 0, Vol(ActiveBall)*VolRH, pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0, AudioFade(ActiveBall):End Sub
Sub aPostRubbers_Hit(idx):PlaySound "fx_postrubber", 0, Vol(ActiveBall)*VolRPo, pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0, AudioFade(ActiveBall):End Sub
Sub aMetals_Hit(idx):PlaySound "fx_MetalHit", 0, Vol(ActiveBall)*VolMetals, pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0, AudioFade(ActiveBall):End Sub
Sub aPlastics_Hit(idx):PlaySound "fx_PlasticHit", 0, Vol(ActiveBall)*VolPlast, pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0, AudioFade(ActiveBall):End Sub
Sub aGates_Hit(idx):PlaySound "fx_Gate", 0, Vol(ActiveBall)*VolGates, pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0, AudioFade(ActiveBall):End Sub
Sub aWoods_Hit(idx):PlaySound "fx_Woodhit", 0, Vol(ActiveBall)*VolWood, pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0, AudioFade(ActiveBall):End Sub

' Ramp Soundss
Sub REnd1_Hit()
    PlaySound "fx_ExitRampToPlayfield", 0, 1, pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0, AudioFade(ActiveBall)
End Sub

Sub REnd2_Hit()
    StopSound "fx_metalrolling"
    PlaySound "fx_ExitRampToPlayfield", 0, 1, pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0, AudioFade(ActiveBall)
End Sub

Sub REnd3_Hit()
    PlaySound "fx_ExitRampToPlayfield", 0, 1, pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0, AudioFade(ActiveBall)
End Sub

Sub REnd4_Hit()
    PlaySound "fx_balldrop", 0, 1, pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0, AudioFade(ActiveBall)
End Sub

' skillshot floor sounds
Sub skillsot_pf_metal_Hit:PlaySound "fx_SkillShot-HitMetal", 0, 1, pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0, AudioFade(ActiveBall):End Sub
Sub skillsot_pf_metal1_Hit:PlaySound "fx_SkillShot-HitMetal", 0, 1, pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0, AudioFade(ActiveBall):End Sub
Sub skillsot_pf_metal2_Hit:PlaySound "fx_SkillShot-HitMetal", 0, 1, pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0, AudioFade(ActiveBall):End Sub

' left ramp sounds
Sub LRHit0_Hit:PlaySound "fx_lr1", 0, 1, pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0, AudioFade(ActiveBall):End Sub
Sub LRHit1_Hit:PlaySound "fx_lr2", 0, 1, pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0, AudioFade(ActiveBall):End Sub
Sub LRHit2_Hit:PlaySound "fx_lr3", 0, 1, pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0, AudioFade(ActiveBall):End Sub
Sub LRHit3_Hit:PlaySound "fx_lr4", 0, 1, pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0, AudioFade(ActiveBall):End Sub
Sub LRHit4_Hit:PlaySound "fx_lr5", 0, 1, pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0, AudioFade(ActiveBall):End Sub
Sub LRHit5_Hit:PlaySound "fx_lr6", 0, 1, pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0, AudioFade(ActiveBall):End Sub
Sub LRHit6_Hit:PlaySound "fx_lr7", 0, 1, pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0, AudioFade(ActiveBall):End Sub

'right ramp sounds
Sub RRHit0_Hit:PlaySound "fx_rr1", 0, 1, pan(ActiveBall), 0, 0, AudioFade(ActiveBall):End Sub
Sub RRHit1_Hit:PlaySound "fx_rr2", 0, 1, pan(ActiveBall), 0, 0, AudioFade(ActiveBall):End Sub
Sub RRHit2_Hit:PlaySound "fx_rr3", 0, 1, pan(ActiveBall), 0, 0, AudioFade(ActiveBall):End Sub
Sub RRHit3_Hit:PlaySound "fx_rr4", 0, 1, pan(ActiveBall), 0, 0, AudioFade(ActiveBall):End Sub
Sub RRHit4_Hit:PlaySound "fx_rr5", 0, 1, pan(ActiveBall), 0, 0, AudioFade(ActiveBall):End Sub
Sub RRHit5_Hit:PlaySound "fx_rr6", 0, 1, pan(ActiveBall), 0, 0, AudioFade(ActiveBall):End Sub
Sub RRHit6_Hit:PlaySound "fx_lr1", 0, 1, pan(ActiveBall), 0, 0, AudioFade(ActiveBall):End Sub
Sub RRHit7_Hit:PlaySound "fx_lr1", 0, 1, pan(ActiveBall), 0, 0, AudioFade(ActiveBall):End Sub
Sub RRHit8_Hit:PlaySound "fx_rr7", 0, 1, pan(ActiveBall), 0, 0, AudioFade(ActiveBall):End Sub

'**********************
' Ramp Diverter
'**********************

Dim RampDivPos, RampDivDir
RampDivPos = 307
RampDivDir = 2 'this is the direction and the speed
RampDiv.Isdropped = 0

Sub RampDiverter(Enabled)
    PlaySound SoundFX("fx_diverter",DOFContactors), 0, 1, 0.02
    If Enabled Then
        RampDivDir = -2
        RampDiv.TimerEnabled = 1
    Else
        RampDivDir = 2
        RampDiv.TimerEnabled = 1
    End If
End Sub

Sub RampDiv_Timer
    RampDivP.RotZ = RampDivPos
    RampDivPos = RampDivPos + RampDivDir
    If RampDivPos < 287 Then
        RampDivPos = 287
        RampDiv.IsDropped = 1
        Me.TimerEnabled = 0
        Exit Sub
    End If
    If RampDivPos > 307 Then
        RampDivPos = 307
        RampDiv.IsDropped = 0
        Me.TimerEnabled = 0
        Exit Sub
    End If
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
  Vol = Csng(BallVel(ball) ^2 / VolDiv)
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

Const tnob = 10 ' total number of balls
ReDim rolling(tnob)
InitRolling

Sub InitRolling
    Dim i
    For i = 0 to tnob
        rolling(i) = False
    Next
End Sub

Sub RollingUpdate()
    Dim BOT, b
    BOT = GetBalls

    ' stop the sound of deleted balls
    For b = UBound(BOT) + 1 to tnob
        rolling(b) = False
        StopSound("fx_ballrolling" & b)
    Next

    ' exit the sub if no balls on the table
    If UBound(BOT) = -1 Then Exit Sub

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


' Thalamus : Exit in a clean and proper way
Sub Table1_exit()
  Controller.Pause = False
  Controller.Stop
End Sub


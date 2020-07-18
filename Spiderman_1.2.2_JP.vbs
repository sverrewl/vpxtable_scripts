' JP's Spider-Man
' Based on Stern's Spider-Man
' Playfield & plastics redrawn by me.
' VPX version by JPSalas 2018, version 1.2.2 (actually this table is from 2014, made while testing VPX beta)

Option Explicit
Randomize

' Thalamus 2018-11-01 : Improved directional sounds
' !! NOTE : Table not verified yet !!

' Options
' Volume devided by - lower gets higher sound

Const VolDiv = 2000    ' Lower number, louder ballrolling/collition sound
Const VolCol = 10      ' Ball collition divider ( voldiv/volcol )

' The rest of the values are multipliers
'
'  .5 = lower volume
' 1.5 = higher volume

Const VolBump   = 2    ' Bumpers volume.
Const VolRol    = 1    ' Rollovers volume.
Const VolGates  = 1    ' Gates volume.
Const VolMetal  = 1    ' Metals volume.
Const VolRB     = 1    ' Rubber bands volume.
Const VolRH     = 1    ' Rubber hits volume.
Const VolPo     = 1    ' Rubber posts volume.
Const VolPi     = 1    ' Rubber pins volume.
Const VolPlast  = 1    ' Plastics volume.
Const VolTarg   = 1    ' Targets volume.
Const VolWood   = 1    ' Woods volume.
Const VolKick   = 1    ' Kicker volume.
Const VolSpin   = 1.5  ' Spinners volume.
Const VolFlip   = 1    ' Flipper volume.


'********************
'Standard definitions
'********************

Dim DesktopMode:DesktopMode = Table1.ShowDT
Dim UseVPMDMD

UseVPMDMD = DesktopMode

Const BallSize = 50
Const BallMass = 1.1

On Error Resume Next
ExecuteGlobal GetTextFile("controller.vbs")
If Err Then MsgBox "You need the controller.vbs in order to run this table, available in the vp10 package"
On Error Goto 0

LoadVPM "01550000", "sam.vbs", 3.26

Const UseSolenoids = 1
Const UseLamps = 0
Const UseGI = 0
Const UseSync = 0 'set it to 1 if the table runs too fast
Const HandleMech = 0

Dim VarHidden
If Table1.ShowDT = true then
    VarHidden = 1
    lrail.Visible = 1
    rrail.Visible = 1
else
    VarHidden = 0
    lrail.Visible = 0
    rrail.Visible = 0
end if

if B2SOn = true then VarHidden = 1

'Const cGameName = "sman_261" 'Spiderman rom
Const cGameName = "smanve_101" 'Spiderman VE rom

'Standard Sounds
Const SSolenoidOn = "fx_solenoidon"
Const SSolenoidOff = "fx_solenoidoff"
Const SFlipperOn = ""
Const SFlipperOff = ""
Const SCoin = "fx_coin"

'Variables
Dim bsTrough, bsSandman, bsDocOck, DocMagnet, PlungerIM, x

'************
' Table init.
'************

Sub Table1_Init
    vpminit Me
    With Controller
        .GameName = cGameName
        If Err Then MsgBox "Can't start Game " & cGameName & vbNewLine & Err.Description:Exit Sub
        .SplashInfoLine = "JP's Spider-Man (Stern 2007)"
        .HandleKeyboard = 0
        .ShowTitle = 0
        .ShowDMDOnly = 1
        .ShowFrame = 0
        .HandleMechanics = 1
        .Hidden = DesktopMode
        On Error Resume Next
        .Run GetPlayerHWnd
        If Err Then MsgBox Err.Description
    End With

    On Error Goto 0

    Controller.Switch(53) = 1 'sandman down

    'Trough
    Set bsTrough = New cvpmBallStack
    bsTrough.InitSw 0, 21, 20, 19, 18, 0, 0, 0
    bsTrough.InitKick BallRelease, 90, 8
    bsTrough.InitExitSnd SoundFX("fx_ballrel", DOFContactors), SoundFX("fx_Solenoid", DOFContactors)
    bsTrough.Balls = 4

    'Sandman VUK
    Set bsSandman = New cvpmBallStack
    bsSandman.InitSw 0, 59, 0, 0, 0, 0, 0, 0
    bsSandman.InitKick sw59a, 90, 35
    bsSandman.KickZ = 1.56
    bsSandman.InitExitSnd SoundFX("fx_kicker", DOFContactors), SoundFX("fx_Solenoid", DOFContactors)
    bsSandman.InitAddSnd "fx_hole_enter"

    'Doc Ock VUK
    Set bsDocOck = New cvpmBallStack
    bsDocOck.InitSw 0, 36, 0, 0, 0, 0, 0, 0
    bsDocOck.InitKick sw36a, 90, 32
    bsDocOck.KickZ = 1.56
    bsDocOck.InitExitSnd SoundFX("fx_kicker", DOFContactors), SoundFX("fx_Solenoid", DOFContactors)
    bsDocOck.InitAddSnd "fx_hole_enter"

    'Doc Ock Magmet
    Set DocMagnet = New cvpmMagnet
    DocMagnet.InitMagnet DocOckMagnet, 5
    DocMagnet.Solenoid = 3
    DocMagnet.GrabCenter = True
    DocMagnet.CreateEvents "DocMagnet"

    'Loop Diverter
    diverter.IsDropped = 1

    'Nudging
    vpmNudge.TiltSwitch = swTilt
    vpmNudge.Sensitivity = 5
    vpmNudge.TiltObj = Array(Bumper1, Bumper2, Bumper3, LeftSlingshot, RightSlingshot)

    'Main Timer init
    PinMAMETimer.Interval = PinMAMEInterval
    PinMAMETimer.Enabled = 1

    'Impulse Plunger
    Const IMPowerSetting = 55 ' Plunger Power
    Const IMTime = 1.1        ' Time in seconds for Full Plunge
    Set plungerIM = New cvpmImpulseP
    With plungerIM
        .InitImpulseP swPlunger, IMPowerSetting, IMTime
        .Switch 23
        .Random 1.5
        .InitExitSnd "fx_plunger2", "fx_plunger"
        .CreateEvents "plungerIM"
    End With

'Fast Flips
	On Error Resume Next 
	InitVpmFFlipsSAM
	If Err Then MsgBox "You need the latest sam.vbs in order to run this table, available with vp10.5"
	On Error Goto 0
End Sub

'**********
' Keys
'**********

Sub Table1_KeyDown(ByVal Keycode)
    If Keycode = RightFlipperKey then Controller.Switch(90) = 1
    If Keycode = StartGameKey Then Controller.Switch(16) = 1
    If keycode = LeftTiltKey Then Nudge 90, 5:PlaySound "fx_nudge", 0, 1, -0.1, 0.25
    If keycode = RightTiltKey Then Nudge 270, 5:PlaySound "fx_nudge", 0, 1, 0.1, 0.25
    If keycode = CenterTiltKey Then Nudge 0, 6:PlaySound "fx_nudge", 0, 1, 0, 0.25
    If keycode = PlungerKey Then PlaySoundAtVol "fx_PlungerPull", Plunger, 1 :Plunger.Pullback
    If vpmKeyDown(Keycode) Then Exit Sub
End Sub

Sub Table1_KeyUp(ByVal Keycode)
    If Keycode = RightFlipperKey then Controller.Switch(90) = 0
    If vpmKeyUp(Keycode) Then Exit Sub
    If Keycode = StartGameKey Then Controller.Switch(16) = 0
    If keycode = PlungerKey Then PlaySoundAtVol "fx_plunger", Plunger, 1:Plunger.Fire
End Sub

'Solenoids
SolCallback(1) = "solTrough"
SolCallback(2) = "solAutofire"
SolCallback(3) = "solDocMagnet"
SolCallback(4) = "solDocOckVUK"
SolCallback(5) = "solDocOck"

SolCallback(7) = "Gate3.open ="
SolCallback(8) = "Gate6.open ="

SolCallback(12) = "solSandmanVUK"
SolCallback(13) = "solSandman"

SolCallback(15) = "SolLFlipper"
SolCallback(16) = "SolRFlipper"

SolCallback(19) = "SolGoblin" 'shake
SolCallback(20) = "sol3Bank"

SolCallback(22) = "solDivert"

'Flashers
SolCallback(21) = "SetLamp 121," 'doc ock
SolCallback(23) = "SetLamp 123," 'sandman x2
SolCallback(25) = "SetLamp 125," 'venom x2
SolCallback(26) = "SetLamp 126," 'sandman arrow
SolCallback(27) = "SetLamp 127," 'sandman dome
SolCallback(28) = "SetLamp 128," 'green goblin x2
SolCallback(29) = "SetLamp 129," 'back panel left
SolCallback(30) = "SetLamp 130," 'back panel right
SolCallback(31) = "SetLamp 131," 'pop bumper x3

'*************
' ShakeGoblin
'*************

Dim GoblinPos

Sub SolGoblin(enabled)
    If enabled Then ShakeGoblin
End Sub

Sub ShakeGoblin
    GoblinPos = 8
    GoblinShakeTimer.Enabled = 1
End Sub

Sub GoblinShakeTimer_Timer
    Goblin.TransY = GoblinPos
    Glider.TransY = GoblinPos
    If GoblinPos = 0 Then GoblinShakeTimer.Enabled = 0:Exit Sub
    If GoblinPos < 0 Then
        GoblinPos = ABS(GoblinPos) - 1
    Else
        GoblinPos = - GoblinPos + 1
    End If
End Sub

'***************
'  Doc Magnet
'***************
' Magnet power is pulsed so wait before turning power off
Sub solDocMagnet(enabled)
    MagnetOffTimer.Enabled = Not enabled
    If enabled Then DocMagnet.MagnetOn = True
End Sub

' Magnet is turned off
' Sends ball/balls to hit Doc
Sub MagnetOffTimer_Timer
    Dim ball
    For Each ball In DocMagnet.Balls 'in case there are more than one ball in the magnet
        With ball
            .VelX = 10: .VelY = -20
        End With
    Next
    Me.Enabled = False:DocMagnet.MagnetOn = False
End Sub

'Solenoid Functions
Sub solTrough(Enabled)
    If Enabled Then
        bsTrough.ExitSol_On
        vpmTimer.PulseSw 22
    End If
End Sub

Sub solAutofire(Enabled)
    If Enabled Then
        PlungerIM.AutoFire
    End If
End Sub

Dim DocDown
DocDown = True
Sub solDocOck(Enabled)
    If Enabled Then
        Controller.Switch(57) = 0
        Controller.Switch(58) = 0
        sw63.TimerInterval = 1000
        sw63.TimerEnabled = 1
    End If
End Sub

Sub sw63_Timer
    If DocDown Then
        DocDown = False
        Controller.Switch(58) = 1
        sw63.IsDropped = 1
    Else
        DocDown = True
        Controller.Switch(57) = 1
        sw63.IsDropped = 0
    End If
    sw63.TimerEnabled = 0
End Sub

Dim SandmanDown
SandmanDown = True
Sub solSandman(Enabled)
    If Enabled Then
        Controller.Switch(53) = 0
        Controller.Switch(54) = 0
        sw42.TimerInterval = 1000
        sw42.TimerEnabled = 1
    End If
End Sub

Sub sw42_Timer
    If SandmanDown Then
        SandmanDown = False
        Controller.Switch(54) = 1
        sw42.IsDropped = 1
    Else
        SandmanDown = True
        Controller.Switch(53) = 1
        sw42.IsDropped = 0
    End If
    sw42.TimerEnabled = 0
End Sub

Sub solSandmanVUK(Enabled)
    If Enabled Then
        bsSandman.ExitSol_On
    End If
End Sub

Sub solDocOckVUK(Enabled)
    If Enabled Then
        bsDocOck.ExitSol_On
    End If
End Sub

Sub solDivert(Enabled)
    If Enabled Then
        Diverter.IsDropped = 0
    Else
        Diverter.IsDropped = 1
    End If
End Sub

'Drains and Kickers
Sub drain_Hit():PlaySoundAtVol "fx_drain", Drain, 1:bsTrough.AddBall Me:End Sub

' Slings
Dim LStep, RStep

Sub LeftSlingShot_Slingshot
    PlaySoundAtVol SoundFX("fx_slingshot", DOFContactors), Lemk, 1
    LeftSling4.Visible = 1
    Lemk.RotX = 26
    LStep = 0
    vpmTimer.PulseSw 26
    LeftSlingShot.TimerEnabled = 1
    ShakeLeftSpider
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
    PlaySoundAtVol SoundFX("fx_slingshot", DOFContactors), Remk, 1
    RightSling4.Visible = 1
    Remk.RotX = 26
    RStep = 0
    vpmTimer.PulseSw 27
    RightSlingShot.TimerEnabled = 1
    ShakeRightSpider
End Sub

Sub RightSlingShot_Timer
    Select Case RStep
        Case 1:RightSLing4.Visible = 0:RightSLing3.Visible = 1:Remk.RotX = 14
        Case 2:RightSLing3.Visible = 0:RightSLing2.Visible = 1:Remk.RotX = 2
        Case 3:RightSLing2.Visible = 0:Remk.RotX = -20:RightSlingShot.TimerEnabled = 0
    End Select
    RStep = RStep + 1
End Sub

' Shake Spiders
Dim SpiderLPos, SpiderRPos

Sub ShakeLeftSpider
    SpiderLPos = 8
    SpiderLTimer.Enabled = 1
End Sub

Sub SpiderLTimer_Timer
    SpiderL.TransY = SpiderLPos
    If SpiderLPos = 0 Then Me.Enabled = 0:Exit Sub
    If SpiderLPos < 0 Then
        SpiderLPos = ABS(SpiderLPos) - 1
    Else
        SpiderLPos = - SpiderLPos + 1
    End If
End Sub

Sub ShakeRightSpider
    SpiderRPos = 8
    SpiderRTimer.Enabled = 1
End Sub

Sub SpiderRTimer_Timer
    SpiderR.TransY = SpiderRPos
    If SpiderRPos = 0 Then Me.Enabled = 0:Exit Sub
    If SpiderRPos < 0 Then
        SpiderRPos = ABS(SpiderRPos) - 1
    Else
        SpiderRPos = - SpiderRPos + 1
    End If
End Sub

'Bumpers
Sub Bumper1_Hit:vpmTimer.PulseSw 30:PlaySoundAtVol SoundFX("fx_bumper", DOFContactors), ActiveBall, VolBump:End Sub
Sub Bumper2_Hit:vpmTimer.PulseSw 31:PlaySoundAtVol SoundFX("fx_bumper", DOFContactors), ActiveBall, VolBump:End Sub
Sub Bumper3_Hit:vpmTimer.PulseSw 32:PlaySoundAtVol SoundFX("fx_bumper", DOFContactors), ActiveBall, VolBump:End Sub

'Rollovers

'Lower Lanes
Sub sw24_Hit:Controller.Switch(24) = 1:PlaySoundAtVol "fx_sensor", ActiveBall, VolRol :End Sub
Sub sw24_UnHit:Controller.Switch(24) = 0:End Sub
Sub sw25_Hit:Controller.Switch(25) = 1:PlaySoundAtVol "fx_sensor", ActiveBall, VolRol :End Sub
Sub sw25_UnHit:Controller.Switch(25) = 0:End Sub
Sub sw28_Hit:Controller.Switch(28) = 1:PlaySoundAtVol "fx_sensor", ActiveBall, VolRol :End Sub
Sub sw28_UnHit:Controller.Switch(28) = 0:End Sub
Sub sw29_Hit:Controller.Switch(29) = 1:PlaySoundAtVol "fx_sensor", ActiveBall, VolRol :End Sub
Sub sw29_UnHit:Controller.Switch(29) = 0:End Sub

'Upper Lanes
Sub sw8_Hit:Controller.Switch(8) = 1:PlaySoundAtVol "fx_sensor", ActiveBall, VolRol :End Sub
Sub sw8_UnHit:Controller.Switch(8) = 0:End Sub
Sub sw33_Hit:Controller.Switch(33) = 1:PlaySoundAtVol "fx_sensor", ActiveBall, VolRol :End Sub
Sub sw33_UnHit:Controller.Switch(33) = 0:End Sub
Sub sw34_Hit:Controller.Switch(34) = 1:PlaySoundAtVol "fx_sensor", ActiveBall, VolRol :End Sub
Sub sw34_UnHit:Controller.Switch(34) = 0:End Sub
Sub sw35_Hit:Controller.Switch(35) = 1:PlaySoundAtVol "fx_sensor", ActiveBall, VolRol :End Sub
Sub sw35_UnHit:Controller.Switch(35) = 0:End Sub

'Right
Sub sw37_Hit:Controller.Switch(37) = 1:PlaySoundAtVol "fx_sensor", ActiveBall, VolRol :End Sub
Sub sw37_UnHit:Controller.Switch(37) = 0:End Sub
Sub sw38_Hit:Controller.Switch(38) = 1:PlaySoundAtVol "fx_sensor", ActiveBall, VolRol :End Sub
Sub sw38_UnHit:Controller.Switch(38) = 0:End Sub

'Right Under Flipper
Sub sw46_Hit:Controller.Switch(46) = 1:PlaySoundAtVol "fx_sensor", ActiveBall, VolRol :End Sub
Sub sw46_UnHit:Controller.Switch(46) = 0:End Sub

'Spinner
Sub sw7_Spin:vpmTimer.PulseSw 7:PlaySoundAtVol "fx_spinner", sw7, VolSpin:End Sub

'Right Ramp
Sub sw45_Hit:Controller.Switch(45) = 1:End Sub
Sub sw45_UnHit:Controller.Switch(45) = 0:End Sub

'Left Ramp
Sub sw47_Hit:Controller.Switch(47) = 1:End Sub
Sub sw47_UnHit:Controller.Switch(47) = 0:End Sub
Sub sw48_Hit:Controller.Switch(48) = 1:End Sub
Sub sw48_UnHit:Controller.Switch(48) = 0:End Sub

'Venom
Sub sw43_Hit:vpmTimer.PulseSw 43:PlaySoundAtVol SoundFX("fx_target", DOFDropTargets), ActiveBall, VolTarg:End Sub
Sub sw43b_Hit:Controller.Switch(43) = 1:End Sub
Sub sw43b_UnHit:Controller.Switch(43) = 0:End Sub

'Doc Ock
Sub sw63_Hit:vpmTimer.PulseSw 63:PlaySoundAtVol SoundFX("fx_target", DOFDropTargets), ActiveBall, VolTarg:End Sub

'Sandman
Sub sw42_Hit:vpmTimer.PulseSw 42:PlaySoundAtVol SoundFX("fx_target", DOFDropTargets), ActiveBall, VolTarg:End Sub

'Lock
Sub sw6_Hit:vpmTimer.PulseSw 6:PlaySoundAtVol SoundFX("fx_target", DOFDropTargets), ActiveBall, VolTarg:End Sub

'Sandman
Sub sw9_Hit:vpmTimer.PulseSw 9:a3BankShake:PlaySoundAtVol SoundFX("fx_target", DOFDropTargets), ActiveBall, VolTarg:End Sub
Sub sw10_Hit:vpmTimer.PulseSw 10:a3BankShake:PlaySoundAtVol SoundFX("fx_target", DOFDropTargets), ActiveBall, VolTarg:End Sub
Sub sw11_Hit:vpmTimer.PulseSw 11:a3BankShake:PlaySoundAtVol SoundFX("fx_target", DOFDropTargets), ActiveBall, VolTarg:End Sub
Sub sw12_Hit:vpmTimer.PulseSw 12:PlaySoundAtVol SoundFX("fx_target", DOFDropTargets), ActiveBall, VolTarg:End Sub
Sub sw13_Hit:vpmTimer.PulseSw 13:PlaySoundAtVol SoundFX("fx_target", DOFDropTargets), ActiveBall, VolTarg:End Sub

'Green Goblin
Sub sw1_Hit:vpmTimer.PulseSw 1:PlaySoundAtVol SoundFX("fx_target", DOFDropTargets), ActiveBall, VolTarg:End Sub
Sub sw2_Hit:vpmTimer.PulseSw 2:PlaySoundAtVol SoundFX("fx_target", DOFDropTargets), ActiveBall, VolTarg:End Sub
Sub sw3_Hit:vpmTimer.PulseSw 3:PlaySoundAtVol SoundFX("fx_target", DOFDropTargets), ActiveBall, VolTarg:End Sub
Sub sw4_Hit:vpmTimer.PulseSw 4:PlaySoundAtVol SoundFX("fx_target", DOFDropTargets), ActiveBall, VolTarg:End Sub
Sub sw5_Hit:vpmTimer.PulseSw 5:PlaySoundAtVol SoundFX("fx_target", DOFDropTargets), ActiveBall, VolTarg:End Sub

'Right 3Bank
Sub sw39_Hit:vpmTimer.PulseSw 39:PlaySoundAtVol SoundFX("fx_target", DOFDropTargets), ActiveBall, VolTarg:End Sub
Sub sw40_Hit:vpmTimer.PulseSw 40:PlaySoundAtVol SoundFX("fx_target", DOFDropTargets), ActiveBall, VolTarg:End Sub
Sub sw41_Hit:vpmTimer.PulseSw 41:PlaySoundAtVol SoundFX("fx_target", DOFDropTargets), ActiveBall, VolTarg:End Sub

'Switch 14
Sub sw14_Hit():vpmTimer.PulseSw 14:PlaySoundAtVol "fx_rubber_band", ActiveBall, 1 :End Sub

'Sandman VUK

Sub sw59_Hit():bsSandman.AddBall Me:End Sub

'DocOck VUK

Sub sw36_Hit():bsDocOck.AddBall Me:End Sub

'*************
' 3Bank Shake
'*************

Dim ccBall
Const cMod = .65 'percentage of hit power transfered to the 3 Bank of targets

a3BankInit

Sub a3BankShake
    ccball.velx = activeball.velx * cMod
    ccball.vely = activeball.vely * cMod
    a3BankTimer.enabled = True
    b3BankTimer.enabled = True
End Sub

Sub a3BankShake2 'when nudging
    a3BankTimer.enabled = True
    b3BankTimer.enabled = True
End Sub

Sub a3BankInit
    Set ccBall = hball.CreateSizedBallWithMass(25, 1.6)
    hball.Kick 0, 0
End Sub

Sub a3BankTimer_Timer            'start animation
    Dim x, y
    x = (hball.x - ccball.x) / 4 'reduce the X axis movement
    y = (hball.y - ccball.y) / 2
    backbank.transy = x
    backbank.transx = - y
    swp9.transy = x
    swp9.transx = - y
    swp10.transy = x
    swp10.transx = - y
    swp11.transy = x
    swp11.transx = - y
End Sub

Sub b3BankTimer_Timer 'stop animation
    backbank.transx = 0
    backbank.transy = 0
    swp9.transz = 0
    swp9.transx = 0
    swp10.transz = 0
    swp10.transx = 0
    swp11.transz = 0
    swp11.transx = 0
    a3BankTimer.enabled = False
    b3BankTimer.enabled = False
End Sub

'******************
'Motor Bank Up Down
'******************
Dim BankDir, BankPos
RiseBank

Sub Sol3Bank(Enabled)
    If Enabled Then
        If BankDir = 1 Then
            RiseBank
        Else
            DropBank
        End If
    End If
End Sub

Sub RiseBank()
    PlaySound "fx_motor"
    'BankPos = 52
    BankDir = -1
    Controller.Switch(49) = 0
    BankTimer.Enabled = 1
End Sub

Sub DropBank()
    PlaySound "fx_motor"
    'BankPos = 0
    BankDir = 1
    Controller.Switch(50) = 0
    BankTimer.Enabled = 1
End Sub

Sub BankTimer_Timer
    BankPos = BankPos + BankDir
    If BankPos > 52 Then
        BankPos = 52
        Me.Enabled = 0
        Controller.Switch(49) = 1
    Else
        If BankPos < 0 Then
            BankPos = 0
            Me.Enabled = 0
            Controller.Switch(50) = 1
        Else
            Update3Bank
        End If
    End If
End Sub

Sub Update3Bank
    backbank.TransZ = - BankPos
    swp9.TransZ = - BankPos
    swp10.TransZ = - BankPos
    swp11.TransZ = - BankPos
    If BankPos > 40 Then
        sw9.Isdropped = 1
        sw10.Isdropped = 1
        sw11.Isdropped = 1
    End If
    If BankPos < 10 Then
        sw9.Isdropped = 0
        sw10.Isdropped = 0
        sw11.Isdropped = 0
    End If
End Sub

'********************
'    JP Flippers
'********************

Sub SolLFlipper(Enabled)
    If Enabled Then
        PlaySoundAtVol SoundFX("fx_flipperup", DOFFlippers), LeftFlipper, VolFlip
        LeftFlipper.RotateToEnd
    Else
        PlaySoundAtVol SoundFX("fx_flipperdown", DOFFlippers), LeftFlipper, VolFlip
        LeftFlipper.RotateToStart
    End If
End Sub

Sub SolRFlipper(Enabled)
    If Enabled Then
        PlaySoundAtVol SoundFX("fx_flipperup", DOFFlippers), RightFlipper, VolFlip
        PlaySoundAtVol "fx_flipperup", RightFlipper1, VolFlip
        RightFlipper.RotateToEnd
        RightFlipper1.RotateToEnd
    Else
        PlaySoundAtVol SoundFX("fx_flipperdown", DOFFlippers), RightFlipper, VolFlip
        PlaySoundAtVol "fx_flipperdown", RightFlipper1, VolFlip
        RightFlipper.RotateToStart
        RightFlipper1.RotateToStart
    End If
End Sub

Sub LeftFlipper_Collide(parm)
    PlaySoundAtBallVol "fx_rubber_flipper", parm / 10
End Sub

Sub RightFlipper_Collide(parm)
    PlaySoundAtBallVol "fx_rubber_flipper", parm / 10
End Sub

Sub RightFlipper1_Collide(parm)
    PlaySoundAtBallVol "fx_rubber_flipper", parm / 10
End Sub

'**********************************************************
'     JP's Flasher Fading for VPX and Vpinmame v 1.0
'       (Based on Pacdude's Fading Light System)
' This is a fast fading for the Flashers in vpinmame tables
'**********************************************************

Dim FadingState(200)

InitLamps() ' turn off the lights and flashers and reset them to the default parameters

' vpinmame Lamp & Flasher Timers

Sub LampTimer_Timer()
    Dim chgLamp, num, chg, ii
    chgLamp = Controller.ChangedLamps
    If Not IsEmpty(chgLamp) Then
        For ii = 0 To UBound(chgLamp)
            FadingState(chgLamp(ii, 0) ) = chgLamp(ii, 1) + 3 'fading step
        Next
    End If
    Lamp 3, l3
    Lamp 4, l4
    Lamp 5, l5
    Lamp 6, l6
    Lamp 7, l7
    Lamp 8, l8
    Lamp 9, l9
    Lamp 10, l10
    Lamp 11, l11
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
    Lamp 27, l27
    Lamp 28, l28
    Lamp 29, l29
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
    Lamp 48, l48
    Lamp 49, l49
    Lamp 50, l50
    Lamp 51, l51
    Lamp 52, l52
    Lamp 53, l53
    Lamp 54, l54
    Lamp 57, l57
    Lamp 58, l58
    Lamp 59, l59
    Lamp 60, Bumper1L
    Lamp 61, Bumper2L
    Lamp 62, Bumper3L
    Lamp 63, l63
    Lamp 64, l64
    Lamp 65, l65
    Flash 66, l66
    Flash 67, l67
    Flash 68, l68
    Flash 69, l69
    Flash 70, l70
    Flash 71, l71
    Lamp 72, l72
    Lampm 74, l74 '74 Sandman
    Flash 74, l74b
    Lampm 75, l75 '75 Venom
    Flash 75, l75b
    Lampm 76, l76 '76 Goblin
    Flash 76, l76b
    Lampm 77, l77 '77 Dock
    Flash 77, l77b
    Lamp 78, l78

    'Flashers
    Flash 121, f21

    Lampm 123, f23a
    Lamp 123, f23
    Lampm 125, f25a
    Lamp 125, f25
    Lamp 126, f26
    Flash 127, f27
    Lampm 128, f28a
    Lamp 128, f28

    Flash 129, f29
    Flash 130, f30
    Lampm 131, f31a
    Lampm 131, f31b
    Lamp 131, f31
End Sub

' div lamp subs

Sub InitLamps()
    Dim x
    LampTimer.Interval = 25 ' flasher fading speed
    LampTimer.Enabled = 1
    For x = 0 to 200
        FadingState(x) = 3 ' used to track the fading state
    Next
End Sub

Sub SetLamp(nr, value) ' 0 is off, 1 is on
    FadingState(nr) = abs(value) + 3
End Sub

' Lights: used for VPX standard lights, the fading is handled by VPX itself, they are here to be able to make them work together with the flashers

Sub Lamp(nr, object)
    Select Case FadingState(nr)
        Case 4:object.state = 1:FadingState(nr) = 0
        Case 3:object.state = 0:FadingState(nr) = 0
    End Select
End Sub

Sub Lampm(nr, object) ' used for multiple lights, it doesn't change the fading state
    Select Case FadingState(nr)
        Case 4:object.state = 1
        Case 3:object.state = 0
    End Select
End Sub

' Flashers: 5 is on, 4,3,2,1 fade steps. 0 is off

Sub Flash(nr, object)
    Select Case FadingState(nr)
        Case 4:Object.IntensityScale = 1:FadingState(nr) = 0
        Case 3:Object.IntensityScale = 0.66:FadingState(nr) = 2
        Case 2:Object.IntensityScale = 0.33:FadingState(nr) = 1
        Case 1:Object.IntensityScale = 0:FadingState(nr) = 0
    End Select
End Sub

Sub Flashm(nr, object) 'multiple flashers, it doesn't change the fading state
    Select Case FadingState(nr)
        Case 4:Object.IntensityScale = 1
        Case 3:Object.IntensityScale = 0.66
        Case 2:Object.IntensityScale = 0.33
        Case 1:Object.IntensityScale = 0
    End Select
End Sub

'******************************
' Diverse Collection Hit Sounds
'******************************

Sub aMetal_Wires_Hit(idx):PlaySound "fx_MetalHit", 0, Vol(ActiveBall)*VolMetal, pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0, AudioFade(ActiveBall):End Sub
Sub aMetal_Lanes_Hit(idx):PlaySound "fx_MetalHit", 0, Vol(ActiveBall)*VolMetal, pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0, AudioFade(ActiveBall):End Sub
Sub aMetal_Guides_Hit(idx):PlaySound "fx_MetalHit", 0, Vol(ActiveBall)*VolMetal, pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0, AudioFade(ActiveBall):End Sub
Sub aRubber_Bands_Hit(idx):PlaySound "fx_rubber_band", 0, Vol(ActiveBall)*VolRB, pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0, AudioFade(ActiveBall):End Sub
Sub aRubber_Posts_Hit(idx):PlaySound "fx_rubber_post", 0, Vol(ActiveBall)*VolPo, pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0, AudioFade(ActiveBall):End Sub
Sub aRubber_Pins_Hit(idx):PlaySound "fx_rubber_pin", 0, Vol(ActiveBall)*VolPi, pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0, AudioFade(ActiveBall):End Sub
Sub aPlastics_Hit(idx):PlaySound "fx_PlasticHit", 0, Vol(ActiveBall)*VolPlast, pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0, AudioFade(ActiveBall):End Sub
Sub aGates_Hit(idx):PlaySound "fx_Gate", 0, Vol(ActiveBall)*VolGates, pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0, AudioFade(ActiveBall):End Sub
Sub aWoods_Hit(idx):PlaySound "fx_Woodhit", 0, Vol(ActiveBall)*VolWood, pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0, AudioFade(ActiveBall):End Sub

' Ramp Soundss
Sub RHelp1_Hit()
    StopSound "fx_metalrolling"
    PlaySoundAtVol "fx_ballrampdrop",  ActiveBall, 1
End Sub

Sub RHelp2_Hit()
    StopSound "fx_metalrolling"
    PlaySoundAtVol "fx_ballrampdrop",  ActiveBall, 1
End Sub

Sub RSound1_Hit:PlaySoundAtVol "fx_metalrolling", ActiveBall, 1 :End Sub

' *******************************************************************************************************
' Positional Sound Playback Functions by DJRobX
' PlaySound sound, 0, Vol(ActiveBall), AudioPan(ActiveBall), 0, Pitch(ActiveBall), 0, 1, AudioFade(ActiveBall)
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

Sub PlaySoundAtVol(sound, tableobj, Volume)
  PlaySound sound, 1, Volume, AudioPan(tableobj), 0,0,0, 1, AudioFade(tableobj)
End Sub

'Set all as per ball position & speed, but Vol Multiplier may be used eg; PlaySoundAtBallVol "sound",3

Sub PlaySoundAtBallVol(sound, VolMult)
  PlaySound sound, 0, Vol(ActiveBall) * VolMult, AudioPan(ActiveBall), 0, Pitch(ActiveBall), 0, 1, AudioFade(ActiveBall)
End Sub

'Set position as bumperX and Vol manually.

Sub PlaySoundAtBumperVol(sound, tableobj, Vol)
  PlaySound sound, 1, Vol, AudioPan(tableobj), 0,0,1, 1, AudioFade(tableobj)
End Sub

'*********************************************************************
'                     Supporting Ball & Sound Functions
'*********************************************************************

Function RndNum(min, max)
    RndNum = Int(Rnd() * (max-min + 1) ) + min ' Sets a random number between min and max
End Function

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
    Dim BOT, b, ballpitch
    BOT = GetBalls

    ' stop the sound of deleted balls
    For b = UBound(BOT) + 1 to tnob
        rolling(b) = False
        StopSound("fx_ballrolling" & b)
    Next

    ' exit the sub if no balls on the table
    If UBound(BOT) = lob - 1 Then Exit Sub 'there no extra balls on this table

    ' play the rolling sound for each ball

    For b = 0 to UBound(BOT)
      If BallVel(BOT(b) ) > 1 Then
        rolling(b) = True
        if BOT(b).z < 30 Then ' Ball on playfield
          PlaySound("fx_ballrolling" & b), -1, Vol(BOT(b) ), AudioPan(BOT(b) ), 0, Pitch(BOT(b) ), 1, 0, AudioFade(BOT(b) )
        Else ' Ball on raised ramp
          PlaySound("fx_ballrolling" & b), -1, Vol(BOT(b) )*.5, AudioPan(BOT(b) ), 0, Pitch(BOT(b) )+50000, 1, 0, AudioFade(BOT(b) )
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
    PlaySound("fx_collide"), 0, Csng(velocity) ^2 / (VolDiv/VolCol), AudioPan(ball1), 0, Pitch(ball1), 0, 0, AudioFade(ball1)
End Sub

'******************
' RealTime Updates
'******************

Set MotorCallback = GetRef("RealTimeUpdates")

Sub RealTimeUpdates
    RollingUpdate
End Sub

'*************************
' GI - needs new vpinmame
'*************************

Set GICallback = GetRef("GIUpdate")

Sub GIUpdate(no, Enabled)
    For each x in aGiLights
        x.State = ABS(Enabled)
    Next
End Sub

' Thalamus : Exit in a clean and proper way
Sub Table1_exit()
  On Error Resume Next
  Controller.Pause = False
  Controller.Stop
End Sub


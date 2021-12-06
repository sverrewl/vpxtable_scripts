' Attack from Mars
' Based on the tables by Bally/Williams
' Use it with VPX.2 and Vpinmame 2.8
' DOF by arngrim
' PMD by WED21

' Thalamus 2018-07-24
' Added/Updated "Positional Sound Playback Functions" and "Supporting Ball & Sound Functions"
' Changed UseSolenoids=1 to 2

Option Explicit
Randomize

On Error Resume Next
ExecuteGlobal GetTextFile("controller.vbs")
If Err Then MsgBox "You need the controller.vbs in order to run this table, available in the vp10 package"
On Error Goto 0

Const BallSize = 50
Const UseVPMModSol = 1

Dim DesktopMode:DesktopMode = Table1.ShowDT
Dim UseVPMColoredDMD

UseVPMColoredDMD = DesktopMode

LoadVPM "01560000", "WPC.VBS", 3.26

'********************
'Standard definitions
'********************

Const UseSolenoids = 2
Const UseLamps = 0
Const UseSync = 0
Const HandleMech = 0

' Standard Sounds
Const SSolenoidOn = "solon"
Const SSolenoidOff = "soloff"
Const SCoin = "fx_Coin"

Set GiCallback2 = GetRef("UpdateGI")

Dim bsTrough, bsL, bsR, dtDrop, x, BallFrame, plungerIM, Mech3bank

'************
' Table init.
'************

Const cGameName = "afm_113b" 'arcade rom - with credits
'Const cGameName = "afm_113"  'home rom - free play

Sub Table1_Init
    vpmInit Me
    With Controller
        .GameName = cGameName
        If Err Then MsgBox "Can't start Game " & cGameName & vbNewLine & Err.Description:Exit Sub
        .SplashInfoLine = "Attack from Mars" & vbNewLine & "VPX table by JPSalas v1.2.0"
        .Games(cGameName).Settings.Value("rol") = 0
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
        '.entrySw = 18
        .initSwitches Array(32, 33, 34, 35)
        .Initexit BallRelease, 80, 6
        .InitExitSounds SoundFX("solon",DOFContactors), SoundFX("fx_ballrel",DOFContactors)
        .Balls = 4
    End With

    ' Droptarget
    Set dtDrop = New cvpmDropTarget
    With dtDrop
        .InitDrop sw77, 77
        .initsnd SoundFX("fx_droptarget",DOFContactors), SoundFX("fx_resetdrop",DOFContactors)
    End With

    ' Left hole
    Set bsL = New cvpmTrough
    With bsL
        .size = 1
        .initSwitches Array(36)
        .Initexit sw36, 0, 2
        .InitExitSounds SoundFX("solon",DOFContactors), SoundFX("Popper_Ball",DOFContactors)
        .InitExitVariance 3, 2
    End With

    ' Right hole
    Set bsR = New cvpmTrough
    With bsR
        .size = 4
        .initSwitches Array(37)
        .Initexit sw37, 200, 24
        .InitExitSounds SoundFX("solon",DOFContactors), SoundFX("Popper_Ball",DOFContactors)
        .InitExitVariance 2, 2
        .MaxBallsPerKick = 1
    End With

    '3 Targets Bank
    Set Mech3Bank = new cvpmMech
    With Mech3Bank
        .Sol1 = 24
        .Mtype = vpmMechLinear + vpmMechReverse + vpmMechOneSol
        .Length = 60
        .Steps = 50
        .AddSw 67, 0, 0
        .AddSw 66, 50, 50
        .Callback = GetRef("Update3Bank")
        .Start
    End With

    ' Impulse Plunger
    Const IMPowerSetting = 42 'Plunger Power
    Const IMTime = 0.6        ' Time in seconds for Full Plunge
    Set plungerIM = New cvpmImpulseP
    With plungerIM
        .InitImpulseP swplunger, IMPowerSetting, IMTime
        .Random 0.3
        .switch 18
        .InitExitSnd SoundFX("plunger",DOFContactors), SoundFX("plunger",DOFContactors)
        .CreateEvents "plungerIM"
    End With

    ' Main Timer init
    PinMAMETimer.Interval = PinMAMEInterval
    PinMAMETimer.Enabled = 1

    ' Init other dropwalls - animations
    UpdateGI 0, 1:UpdateGI 1, 1:UpdateGI 2, 1
    UFORotSpeedSlow
    UfoLed.Enabled = 1

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

'**********
' Keys
'**********

Sub table1_KeyDown(ByVal Keycode)
    If keycode = PlungerKey Then Controller.Switch(11) = 1
    If keycode = LeftTiltKey Then Nudge 90, 4:PlaySound SoundFX("fx_nudge",0), 0, 1, -0.1, 0.25:aSaucerShake:a3BankShake2
    If keycode = RightTiltKey Then Nudge 270, 4:PlaySound SoundFX("fx_nudge",0), 0, 1, 0.1, 0.25:aSaucerShake:a3BankShake2
    If keycode = CenterTiltKey Then Nudge 0, 5:PlaySound SoundFX("fx_nudge",0), 0, 1, 0, 0.25:aSaucerShake:a3BankShake2
    If vpmKeyDown(keycode) Then Exit Sub
End Sub

Sub table1_KeyUp(ByVal Keycode)
    If keycode = PlungerKey Then Controller.Switch(11) = 0
    If vpmKeyUp(keycode) Then Exit Sub
End Sub

'*********
' Switches
'*********

' Slings & div switches

Dim LStep, RStep

Sub LeftSlingShot_Slingshot
    PlaySoundAt SoundFX("SlingshotLeft",DOFContactors), peg2
    LeftSling4.Visible = 1
    Lemk.RotX = 26
    LStep = 0
    vpmTimer.PulseSw 51
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
    PlaySoundAt SoundFX("SlingshotRight",DOFContactors), peg5
    RightSling4.Visible = 1
    Remk.RotX = 26
    RStep = 0
    vpmTimer.PulseSw 52
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
Sub Bumper1_Hit:vpmTimer.PulseSw 53:PlaySoundAt SoundFX("fx_bumper1",DOFContactors), Bumper1:End Sub

Sub Bumper2_Hit:vpmTimer.PulseSw 54:PlaySoundAt SoundFX("fx_bumper2",DOFContactors), Bumper2:End Sub

Sub Bumper3_Hit:vpmTimer.PulseSw 55:PlaySoundAt SoundFX("fx_bumper3",DOFContactors), Bumper3:End Sub

' Drain holes, vuks & saucers
Sub Drain_Hit:PlaySoundAt "fx_drainShort", Drain:bsTrough.AddBall Me:End Sub
Sub sw36a_Hit:PlaySoundAt "KickerEnter", sw36a:bsL.AddBall Me:End Sub
Sub sw37a_Hit:PlaySoundAt "KickerEnter", sw37a:bsR.AddBall Me:End Sub

Dim aBall

Sub sw78_Hit
    PlaySoundAt "fx_balldrop", sw78
    Set aBall = ActiveBall:Me.TimerEnabled = 1
    vpmTimer.PulseSwitch(78), 150, "bsl.addball 0 '"
  Me.Enabled = 0
End Sub

Sub sw78_Timer
    Do While aBall.Z > -25
        aBall.Z = aBall.Z -5
        Exit Sub
    Loop
    Me.DestroyBall
    Me.TimerEnabled = 0
  Me.Enabled = 1
End Sub

Dim bBall

Sub sw37_Hit
    PlaySoundAt "fx_balldrop", sw37
    Set bBall = ActiveBall:Me.TimerEnabled = 1
    bsR.AddBall 0
  Me.Enabled = 0
End Sub

Sub sw37_Timer
    Do While bBall.Z > -25
        bBall.Z = bBall.Z -5
        Exit Sub
    Loop
    Me.DestroyBall
    Me.TimerEnabled = 0
  Me.Enabled = 1
End Sub

' Rollovers & Ramp Switches
Sub sw16_Hit:Controller.Switch(16) = 1:End Sub
Sub sw16_UnHit:Controller.Switch(16) = 0:End Sub

Sub sw26_Hit:Controller.Switch(26) = 1:End Sub
Sub sw26_UnHit:Controller.Switch(26) = 0:End Sub

Sub sw17_Hit:Controller.Switch(17) = 1:End Sub
Sub sw17_UnHit:Controller.Switch(17) = 0:End Sub

Sub sw27_Hit:Controller.Switch(27) = 1:End Sub
Sub sw27_UnHit:Controller.Switch(27) = 0:End Sub

Sub sw38_Hit:Controller.Switch(38) = 1:End Sub
Sub sw38_Unhit:Controller.Switch(38) = 0:End Sub

Sub sw48_Hit:Controller.Switch(48) = 1:End Sub
Sub sw48_Unhit:Controller.Switch(48) = 0:End Sub

Sub sw71_Hit:Controller.Switch(71) = 1:End Sub
Sub sw71_UnHit:Controller.Switch(71) = 0:End Sub

Sub sw72_Hit:Controller.Switch(72) = 1:End Sub
Sub sw72_Unhit:Controller.Switch(72) = 0:End Sub

Sub sw73_Hit:Controller.Switch(73) = 1:End Sub
Sub sw73_Unhit:Controller.Switch(73) = 0:End Sub

Sub sw74_Hit:Controller.Switch(74) = 1:End Sub
Sub sw74_Unhit:Controller.Switch(74) = 0:End Sub

Sub sw61_Hit:Controller.Switch(61) = 1:End Sub
Sub sw61_Unhit:Controller.Switch(61) = 0:End Sub

Sub sw62_Hit:Controller.Switch(62) = 1:End Sub
Sub sw62_Unhit:Controller.Switch(62) = 0:End Sub

Sub sw63_Hit:Controller.Switch(63) = 1:End Sub
Sub sw63_Unhit:Controller.Switch(63) = 0:End Sub

Sub sw64_Hit:Controller.Switch(64) = 1:End Sub
Sub sw64_Unhit:Controller.Switch(64) = 0:End Sub

Sub sw1_Hit:PlaySound "fx_metalrolling", 0, 1, 0.15, 0.25:End Sub

Sub sw65_Hit:Controller.Switch(65) = 1:PlaySoundAt "fx_metalrolling", sw65:End Sub
Sub sw65_Unhit:Controller.Switch(65) = 0:End Sub

' Targets
Sub sw56_Hit:vpmTimer.PulseSw 56:End Sub

Sub sw57_Hit:vpmTimer.PulseSw 57:End Sub

Sub sw58_Hit:vpmTimer.PulseSw 58:End Sub

Sub sw41_Hit:vpmTimer.PulseSw 41:End Sub

Sub sw42_Hit:vpmTimer.PulseSw 42:End Sub

Sub sw43_Hit:vpmTimer.PulseSw 43:End Sub

Sub sw44_Hit:vpmTimer.PulseSw 44:End Sub

Sub sw75_Hit:vpmTimer.PulseSw 75:End Sub

Sub sw76_Hit:vpmTimer.PulseSw 76:End Sub

' 3 bank
Sub sw45_Hit:vpmTimer.PulseSw 45:a3BankShake:PlaySound SoundFX("fx_target",DOFContactors), 0, 1, 0, 0.25:End Sub

Sub sw46_Hit:vpmTimer.PulseSw 46:a3BankShake:PlaySound SoundFX("fx_target",DOFContactors), 0, 1, 0, 0.25:End Sub

Sub sw47_Hit:vpmTimer.PulseSw 47:a3BankShake:PlaySound SoundFX("fx_target",DOFContactors), 0, 1, 0, 0.25:End Sub

' Droptarget

Sub sw77_Hit:PlaySoundAt SoundFX("fx_DTDrop",DOFContactors), sw77:dtDrop.Hit 1:UFORotSpeedFast:End Sub

' Ramps helpers
Sub RHelp1_Hit()
    ActiveBall.VelZ = 0
    ActiveBall.VelY = 0
    ActiveBall.VelX = 0
    StopSound "fx_metalrolling"
    PlaySoundAtBall "fx_ballrampdrop"
End Sub

Sub RHelp2_Hit()
    ActiveBall.VelZ = 0
    ActiveBall.VelY = 0
    ActiveBall.VelX = 0
    StopSound "fx_metalrolling"
    PlaySoundAtBall "fx_ballrampdrop"
End Sub

'*********
'Solenoids
'*********

SolCallback(1) = "Auto_Plunger"
SolCallback(2) = "SolRelease"
SolCallback(3) = "bsL.SolOut"
SolCallback(4) = "bsR.SolOut"
SolCallback(5) = "SolAlien5"
SolCallback(6) = "SolAlien6"
SolCallback(7) = "vpmSolSound SoundFX(""fx_Knocker"",DOFKnocker),"
SolCallback(8) = "SolAlien8"
SolCallback(14) = "SolAlien14"
SolCallBack(15) = "SolUfoShake"
SolCallback(16) = "SolDropTargetUp"
SolModCallback(17) = "SetModLamp 117,"
SolModCallback(18) = "SetModLamp 118,"
SolModCallback(19) = "SetModLamp 119,"
SolModCallback(20) = "SetModLamp 120,"
SolModCallback(21) = "SetModLamp 121,"
SolModCallback(22) = "SetModLamp 122,"
SolModCallback(23) = "SetModLamp 123," ' SolUfoFlash"
'SolCallback(24) = "SolBank" 'used in the Mech
SolModCallback(25) = "SetModLamp 125,"
SolModCallback(26) = "SetModLamp 126,"
SolModCallback(27) = "SetModLamp 127,"
SolModCallback(28) = "SetModLamp 128,"
SolCallback(33) = "vpmSolGate RGate,false,"
SolCallback(34) = "vpmSolGate LGate,false,"
SolCallback(36) = "SolDiv"
SolCallback(43) = "SetLamp 130,"

Sub SolRelease(Enabled)
    If Enabled And bsTrough.Balls > 0 Then
        vpmTimer.PulseSw 31
        bsTrough.ExitSol_On
    End If
End Sub

Sub Auto_Plunger(Enabled)
    If Enabled Then
        PlungerIM.AutoFire
    End If
End Sub

Sub SolDropTargetUp(Enabled)
    If Enabled Then
        dtDrop.DropSol_On
    End If
End Sub

'****************
' Alien solenoids
'****************

Sub SolAlien5(Enabled)
    PlaySound "fx_rubber2", 0, 1, -0.15, 0.25
    If Enabled Then
        Alien5.TransZ = 20
    Else
        Alien5.TransZ = 0
    End If
End Sub

Sub SolAlien6(Enabled)
    PlaySound "fx_rubber2", 0, 1, 0, 0.25
    If Enabled Then
        Alien6.TransZ = 20
    Else
        Alien6.TransZ = 0
    End If
End Sub

Sub SolAlien8(Enabled)
    PlaySound "fx_rubber2", 0, 1, , 0.25
    If Enabled Then
        Alien8.TransZ = 20
    Else
        Alien8.TransZ = 0
    End If
End Sub

Sub SolAlien14(Enabled)
    PlaySound "fx_rubber2", 0, 1, 0.15, 0.25
    If Enabled Then
        Alien14.TransZ = 20
    Else
        Alien14.TransZ = 0
    End If
End Sub

'***********
' Diverter
'***********

Dim DivPos, DivDir
DivPos = 298
DivDir = 2 'this is the direction and the speed
Div.Isdropped = True

Sub SolDiv(Enabled)
    PlaySound SoundFX("fx_diverter",DOFContactors), 0, 1, 0.02
    If Enabled Then
        DivDir = 2
        Div.TimerEnabled = 1
    Else
        DivDir = -2
        Div.TimerEnabled = 1
    End If
End Sub

Sub Div_Timer
    DivP.RotZ = DivPos
    DivPos = DivPos + DivDir
    If DivPos < 298 Then
        DivPos = 298
        Div.IsDropped = True
        Me.TimerEnabled = False
        Exit Sub
    End If
    If DivPos > 338 Then
        DivPos = 338
        Div.IsDropped = False
        Me.TimerEnabled = False
        Exit Sub
    End If
End Sub

'********************
'    JP Flippers
'********************

SolCallback(sLRFlipper) = "SolRFlipper"
SolCallback(sLLFlipper) = "SolLFlipper"

Sub SolLFlipper(Enabled)
    If Enabled Then
        PlaySoundAt SoundFX("fx_flipperupleft",DOFContactors), LeftFlipper
        LeftFlipper.RotateToEnd
    Else
        PlaySoundAt SoundFX("fx_flipperdownleft",DOFContactors), LeftFlipper
        LeftFlipper.RotateToStart
    End If
End Sub

Sub SolRFlipper(Enabled)
    If Enabled Then
        PlaySoundAt SoundFX("fx_flipperupright",DOFContactors), RightFlipper
        RightFlipper.RotateToEnd
    Else
        PlaySoundAt SoundFX("fx_flipperdownright",DOFContactors), RightFlipper
        RightFlipper.RotateToStart
    End If
End Sub

'Sub LeftFlipper_Collide(parm)
'    PlaySound "fx_rubber_flipper", 0, parm / 10, -0.1, 0.25
'End Sub

'Sub RightFlipper_Collide(parm)
'    PlaySound "fx_rubber_flipper", 0, parm / 10, 0.1, 0.25
'End Sub

'******************
'Motor Bank Up Down
'******************

Sub Update3Bank(currpos, currspeed, lastpos)
    If currpos <> lastpos Then
        BackBank.Z = 25 - currpos
        swp45.Z = -(22 + currpos)
        swp46.Z = -(22 + currpos)
        swp47.Z = -(22 + currpos)
    End If
    If currpos > 40 Then
        sw45.Isdropped = 1
        sw46.Isdropped = 1
        sw47.Isdropped = 1
        UFORotSpeedMedium
    End If
    If currpos < 10 Then
        sw45.Isdropped = 0
        sw46.Isdropped = 0
        sw47.Isdropped = 0
        UFORotSpeedSlow
    End If
End Sub

'***********
' Update GI
'***********

Sub UpdateGI(no, step)
    Dim gistep, ii, a
    gistep = step / 8
    Select Case no
        Case 0
            For each ii in aGiLLights
                ii.IntensityScale = gistep
            Next
        Case 1
            For each ii in aGiMLights
                ii.IntensityScale = gistep
            Next
        Case 2 ' also the bumpers er GI
            For each ii in aGiTLights
                ii.IntensityScale = gistep
            Next
    End Select
End Sub

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
    Set ccBall = hball.createball
    hball.Kick 0, 0
    ccball.Mass = 1.6
End Sub

Sub a3BankTimer_Timer            'start animation
    Dim x, y
    x = (hball.x - ccball.x) / 4 'reduce the X axis movement
    y = (hball.y - ccball.y) / 2
    backbank.transy = x
    backbank.transx = - y
    swp45.transy = x
    swp45.transx = - y
    swp46.transy = x
    swp46.transx = - y
    swp47.transy = x
    swp47.transx = - y
End Sub

Sub b3BankTimer_Timer 'stop animation
    backbank.transx = 0
    backbank.transy = 0
    swp45.transz = 0
    swp45.transx = 0
    swp46.transz = 0
    swp46.transx = 0
    swp47.transz = 0
    swp47.transx = 0
    a3BankTimer.enabled = False
    b3BankTimer.enabled = False
End Sub

'***************
' Big UFO Shake
'***************
Dim cBall, UFOLedPos
BigUfoInit

Sub BigUfoInit
    UFOLedPos = 0
    Set cBall = ckicker.createball
    ckicker.Kick 0, 0
End Sub

Sub SolUfoShake(Enabled)
    If Enabled Then
        BigUfoShake
    End If
End Sub

Sub BigUfoShake
    cball.velx = -2 + 2 * RND(1)
    cball.vely = -10 + 2 * RND(1)
End Sub

Sub UFOLed_Timer()
    Select Case UfoLedPos
        Case 0:ufo1.image = "bigufo1":UfoLedPos = 1
        Case 1:ufo1.image = "bigufo2":UfoLedPos = 2
        Case 2:ufo1.image = "bigufo3":UfoLedPos = 0
    End Select
End Sub

Sub BigUfoUpdate
    Dim a, b
    a = ckicker.y - cball.y
    b = cball.x - ckicker.x

    Ufo1.rotx = - a
    Ufo1.transy = a
    Ufo1.roty = b
    Ufo1d.rotx = - a
    Ufo1d.transy = a
    Ufo1d.roty = b
End Sub

'**********************************
'   Small UFOs Rotation Speed
'**********************************

Sub UFORotSpeedSlow()
    UfoLed.Interval = 600
    RotateUFO.Interval = 100
End Sub

Sub UFORotSpeedMedium()
    UfoLed.Interval = 300
    RotateUFO.Interval = 50
End Sub

Sub UFORotSpeedFast()
    UfoLed.Interval = 150
    RotateUFO.Interval = 25
End Sub

Dim UFOSmallPos
UFOSmallPos = 0

Sub RotateUFO_Timer()
    Ufo2.RotZ = (Ufo2.RotZ - 2) MOD 360
    Ufo4.RotZ = (Ufo4.RotZ - 2) MOD 360
    Ufo5.RotZ = (Ufo5.RotZ - 2) MOD 360
    Ufo6.RotZ = (Ufo6.RotZ - 2) MOD 360
    Ufo7.RotZ = (Ufo7.RotZ - 2) MOD 360
    Ufo8.RotZ = (Ufo8.RotZ - 2) MOD 360
    Select Case UFOSmallPos
        Case 0:ufo2.image = "saucer1":ufo4.image = "saucer4":ufo5.image = "saucer5":ufo6.image = "saucer6":ufo7.image = "saucer7":ufo8.image = "saucer8":UFOSmallPos = 1
        Case 1:ufo2.image = "saucer2":ufo4.image = "saucer5":ufo5.image = "saucer6":ufo6.image = "saucer7":ufo7.image = "saucer8":ufo8.image = "saucer1":UFOSmallPos = 2
        Case 2:ufo2.image = "saucer3":ufo4.image = "saucer6":ufo5.image = "saucer7":ufo6.image = "saucer8":ufo7.image = "saucer1":ufo8.image = "saucer2":UFOSmallPos = 3
        Case 3:ufo2.image = "saucer4":ufo4.image = "saucer7":ufo5.image = "saucer8":ufo6.image = "saucer1":ufo7.image = "saucer2":ufo8.image = "saucer3":UFOSmallPos = 4
        Case 4:ufo2.image = "saucer5":ufo4.image = "saucer8":ufo5.image = "saucer1":ufo6.image = "saucer2":ufo7.image = "saucer3":ufo8.image = "saucer4":UFOSmallPos = 5
        Case 5:ufo2.image = "saucer6":ufo4.image = "saucer1":ufo5.image = "saucer2":ufo6.image = "saucer3":ufo7.image = "saucer4":ufo8.image = "saucer5":UFOSmallPos = 6
        Case 6:ufo2.image = "saucer7":ufo4.image = "saucer2":ufo5.image = "saucer3":ufo6.image = "saucer4":ufo7.image = "saucer5":ufo8.image = "saucer6":UFOSmallPos = 7
        Case 7:ufo2.image = "saucer8":ufo4.image = "saucer3":ufo5.image = "saucer4":ufo6.image = "saucer5":ufo7.image = "saucer6":ufo8.image = "saucer7":UFOSmallPos = 0
    End Select
End Sub

'**********************************************************
' Small shake of small Ufos and other objects when nudging
'**********************************************************

Dim SmallShake:SmallShake = 0

Sub aSaucerShake
    SmallShake = 6
    SaucerShake.Enabled = True
End Sub

Sub SaucerShake_Timer
    ufo2.Rotx = SmallShake
    ufo2a.Rotx = - SmallShake
    ufo7.Rotx = SmallShake
    ufo7a.Rotx = - SmallShake
    ufo8.Rotx = SmallShake
    ufo8a.Rotx = - SmallShake
    ufo6.Rotx = SmallShake
    ufo6a.Rotx = - SmallShake
    ufo5.Rotx = SmallShake
    ufo5a.Rotx = - SmallShake
    ufo4.Rotx = SmallShake
    ufo4a.Rotx = - SmallShake
    alien5.Transz = SmallShake / 2
    alien6.Transz = SmallShake / 2
    alien8.Transz = SmallShake / 2
    alien14.Transz = SmallShake / 2
    If SmallShake = 0 Then SaucerShake.Enabled = False:Exit Sub
    If SmallShake < 0 Then
        SmallShake = ABS(SmallShake) - 0.1
    Else
        SmallShake = - SmallShake + 0.1
    End If
End Sub

'***************************************************
'       JP's VP10 Fading Lamps & Flashers v2
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
    FadeL 11, l11, "Lights4", "Lights3", "Lights2", "pf"
    FadeL 12, l12, "Lights4", "Lights3", "Lights2", "pf"
    FadeL 13, l13, "Lights4", "Lights3", "Lights2", "pf"
    FadeL 14, l14, "Lights4", "Lights3", "Lights2", "pf"
    FadeL 15, l15, "Lights4", "Lights3", "Lights2", "pf"
    FadeL 16, l16, "Lights4", "Lights3", "Lights2", "pf"
    FadeL 17, l17, "Lights4", "Lights3", "Lights2", "pf"
    FadeL 18, l18, "Lights4", "Lights3", "Lights2", "pf"
    FadeL 21, l21, "Lights4", "Lights3", "Lights2", "pf"
    FadeL 22, l22, "Lights4", "Lights3", "Lights2", "pf"
    FadeL 23, l23, "Lights4", "Lights3", "Lights2", "pf"
    FadeL 24, l24, "Lights4", "Lights3", "Lights2", "pf"
    FadeL 25, l25, "Lights4", "Lights3", "Lights2", "pf"
    FadeL 26, l26, "Lights4", "Lights3", "Lights2", "pf"
    FadeL 27, l27, "Lights4", "Lights3", "Lights2", "pf"
    FadeL 28, l28, "Lights4", "Lights3", "Lights2", "pf"
    FadeL 31, l31, "Lights4", "Lights3", "Lights2", "pf"
    FadeL 32, l32, "Lights4", "Lights3", "Lights2", "pf"
    FadeL 33, l33, "Lights4", "Lights3", "Lights2", "pf"
    FadeL 34, l34, "Lights4", "Lights3", "Lights2", "pf"
    FadeL 35, l35, "Lights4", "Lights3", "Lights2", "pf"
    FadeL 36, l36, "Lights4", "Lights3", "Lights2", "pf"
    FadeL 37, l37, "Lights4", "Lights3", "Lights2", "pf"
    FadeL 38, l38, "Lights4", "Lights3", "Lights2", "pf"
    FadeL 41, l41, "Lights4", "Lights3", "Lights2", "pf"
    FadeL 42, l42, "Lights4", "Lights3", "Lights2", "pf"
    FadeL 43, l43, "Lights4", "Lights3", "Lights2", "pf"
    FadeL 44, l44, "Lights4", "Lights3", "Lights2", "pf"
    FadeL 45, l45, "Lights4", "Lights3", "Lights2", "pf"
    FadeL 46, l46, "Lights4", "Lights3", "Lights2", "pf"
    FadeL 47, l47, "Lights4", "Lights3", "Lights2", "pf"
    FadeL 48, l48, "Lights4", "Lights3", "Lights2", "pf"
    FadeL 51, l51, "Lights4", "Lights3", "Lights2", "pf"
    FadeL 52, l52, "Lights4", "Lights3", "Lights2", "pf"
    FadeL 53, l53, "Lights4", "Lights3", "Lights2", "pf"
    FadeL 54, l54, "Lights4", "Lights3", "Lights2", "pf"
    FadeL 55, l55, "Lights4", "Lights3", "Lights2", "pf"
    FadeL 56, l56, "Lights4", "Lights3", "Lights2", "pf"
    FadeL 57, l57, "Lights4", "Lights3", "Lights2", "pf"
    FadeL 58, l58, "Lights4", "Lights3", "Lights2", "pf"
    FadeL 61, l61, "Lights4", "Lights3", "Lights2", "pf"
    FadeL 62, l62, "Lights4", "Lights3", "Lights2", "pf"
    FadeL 63, l63, "Lights4", "Lights3", "Lights2", "pf"
    FadeL 64, l64, "Lights4", "Lights3", "Lights2", "pf"
    FadeL 65, l65, "Lights4", "Lights3", "Lights2", "pf"
    FadeL 66, l66, "Lights4", "Lights3", "Lights2", "pf"
    FadeL 67, l67, "Lights4", "Lights3", "Lights2", "pf"
    FadeL 68, l68, "Lights4", "Lights3", "Lights2", "pf"
    FadeL 71, l71, "Lights4", "Lights3", "Lights2", "pf"
    FadeL 72, l72, "Lights4", "Lights3", "Lights2", "pf"
    FadeL 73, l73, "Lights4", "Lights3", "Lights2", "pf"
    FadeL 74, l74, "Lights4", "Lights3", "Lights2", "pf"
    FadeL 75, l75, "Lights4", "Lights3", "Lights2", "pf"
    FadeL 76, l76, "Lights4", "Lights3", "Lights2", "pf"
    FadeL 77, l77, "Lights4", "Lights3", "Lights2", "pf"
    FadeL 78, l78, "Lights4", "Lights3", "Lights2", "pf"
    FadeL 81, l81, "Lights4", "Lights3", "Lights2", "pf"
    FadeL 82, l82, "Lights4", "Lights3", "Lights2", "pf"
    FadeL 83, l83, "Lights4", "Lights3", "Lights2", "pf"
    FadeL 84, l84, "Lights4", "Lights3", "Lights2", "pf"
    FadeL 85, l85, "Lights4", "Lights3", "Lights2", "pf"
    NFadeL 86, l86
    NFadeL 88, l88
    ' ufo red lights
    'NFadeL 91, l91
    'NFadeL 92, l92
    'NFadeL 93, l93
    'NFadeL 94, l94
    'NFadeL 95, l95
    'NFadeL 96, l96
    'NFadeL 97, l97
    'NFadeL 98, l98
    'NFadeL 101, l101
    'NFadeL 102, l102
    'NFadeL 103, l103
    'NFadeL 104, l104
    'NFadeL 105, l105
    'NFadeL 106, l106
    'NFadeL 107, l107
    'NFadeL 108, l108
    'flashers
    FlashMod 117, f17b
    FlashMod 117, f17c
    FlashMod 117, f17a
    FlashMod 118, f18b
    FlashMod 118, f18c
    FlashMod 118, f18a
    FlashMod 119, f19b
    FlashMod 119, f19a
    LightMod 120, f20
    FlashMod 120, f20b
    FlashMod 121, f21
    LightMod 122, f22
    LightMod 123, f23
    FlashMod 125, f25b
    FlashMod 125, f25c
    FlashMod 125, f25a
    FlashMod 126, f26b
    FlashMod 126, f26a
    FlashMod 127, f27b
    FlashMod 127, f27a
    LightMod 128, f28
    FlashMod 128, f28b
    FastFlash 130, Strobe
End Sub

' div lamp subs

Sub InitLamps()
    Dim x
    For x = 0 to 200
        LampState(x) = 0         ' current light state, independent of the fading level. 0 is off and 1 is on
        FadingLevel(x) = 4       ' used to track the fading state
        FlashSpeedUp(x) = 0.2    ' faster speed when turning on the flasher
        FlashSpeedDown(x) = 0.05 ' slower speed when turning off the flasher
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

' Lights: old method, using 4 images

Sub FadeL(nr, light, a, b, c, d)
    Select Case FadingLevel(nr)
        Case 4:light.image = b:FadingLevel(nr) = 6                   'fading to off...
        Case 5:light.image = a:light.State = 1:FadingLevel(nr) = 1   'ON
        Case 6, 7, 8:FadingLevel(nr) = FadingLevel(nr) + 1           'wait
        Case 9:light.image = c:FadingLevel(nr) = FadingLevel(nr) + 1 'fading...
        Case 10, 11, 12:FadingLevel(nr) = FadingLevel(nr) + 1        'wait
        Case 13:light.image = d:Light.State = 0:FadingLevel(nr) = 0  'Off
    End Select
End Sub

Sub FadeLm(nr, light, a, b, c, d)
    Select Case FadingLevel(nr)
        Case 4:light.image = b
        Case 5:light.image = a
        Case 9:light.image = c
        Case 13:light.image = d
    End Select
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

'Ramps & Primitives used as 4 step fading lights
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

Sub FastFlash(nr, object)
    Select Case FadingLevel(nr)
        Case 4:object.Visible = 0:FadingLevel(nr) = 0 'off
        Case 5:object.Visible = 1:FadingLevel(nr) = 1 'on
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

' Ramp Soundss
Sub REnd1_Hit()
    ActiveBall.VelX = 0
    ActiveBall.VelY = 0
    ActiveBall.VelZ = 0
    StopSound "fx_metalrolling"
    PlaySoundAtBall "fx_ballrampdrop"
End Sub

Sub REnd2_Hit()
    ActiveBall.VelX = 0
    ActiveBall.VelY = 0
    ActiveBall.VelZ = 0
    StopSound "fx_metalrolling"
    PlaySoundAtBall "fx_ballrampdrop"
End Sub

Sub RSound1_Hit:PlaySoundAtBall "fx_metalrolling":End Sub

Sub LRampTrig_Hit() : PlaysoundAt "fx_rrenter" ,LRampTrig: End Sub
Sub RRampTrig_Hit() : PlaysoundAt "fx_rrenter" ,RRampTrig: End Sub

'*********************************************************************
'                 Positional Sound Playback Functions
'*********************************************************************

' Play a sound, depending on the X,Y position of the table element (especially cool for surround speaker setups, otherwise stereo panning only)
' parameters (defaults): loopcount (1), volume (1), randompitch (0), pitch (0), useexisting (0), restart (1))
' Note that this will not work (currently) for walls/slingshots as these do not feature a simple, single X,Y position

Sub PlayXYSound(soundname, tableobj, loopcount, volume, randompitch, pitch, useexisting, restart)
  PlaySound soundname, loopcount, volume, AudioPan(tableobj), randompitch, pitch, useexisting, restart, AudioFade(tableobj)
End Sub

' Similar subroutines that are less complicated to use (e.g. simply use standard parameters for the PlaySound call)
Sub PlaySoundAt(soundname, tableobj)
    PlaySound soundname, 1, 1, AudioPan(tableobj), 0,0,0, 1, AudioFade(tableobj)
End Sub

Sub PlaySoundAtBall(soundname)
    PlaySoundAt soundname, ActiveBall
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

Const tnob = 5 ' total number of balls
ReDim rolling(tnob)
InitRolling

Sub InitRolling
    Dim i
    For i = 0 to tnob
        rolling(i) = False
    Next
End Sub

Sub RollingTimer_Timer()
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
        If BallVel(BOT(b) ) > 1 AND BOT(b).z < 30 Then
            rolling(b) = True
            PlaySound("fx_ballrolling" & b), -1, Vol(BOT(b)), AudioPan(BOT(b)), 0, Pitch(BOT(b)), 1, 0, AudioFade(BOT(b))
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
  PlaySound("fx_collide"), 0, Csng(velocity) ^2 / 2000, AudioPan(ball1), 0, Pitch(ball1), 0, 0, AudioFade(ball1)
End Sub


'************************************
' What you need to add to your table
'************************************

' a timer called RollingTimer. With a fast interval, like 10
' one collision sound, in this script is called fx_collide
' as many sound files as max number of balls, with names ending with 0, 1, 2, 3, etc
' for ex. as used in this script: fx_ballrolling0, fx_ballrolling1, fx_ballrolling2, fx_ballrolling3, etc


'******************************************
' Explanation of the rolling sound routine
'******************************************

' sounds are played based on the ball speed and position

' the routine checks first for deleted balls and stops the rolling sound.

' The For loop goes through all the balls on the table and checks for the ball speed and
' if the ball is on the table (height lower than 30) then then it plays the sound
' otherwise the sound is stopped, like when the ball has stopped or is on a ramp or flying.

' The sound is played using the VOL, AUDIOPAN, AUDIOFADE and PITCH functions, so the volume and pitch of the sound
' will change according to the ball speed, and the AUDIOPAN & AUDIOFADE functions will change the stereo position
' according to the position of the ball on the table.


'**************************************
' Explanation of the collision routine
'**************************************

' The collision is built in VP.
' You only need to add a Sub OnBallBallCollision(ball1, ball2, velocity) and when two balls collide they
' will call this routine. What you add in the sub is up to you. As an example is a simple Playsound with volume and paning
' depending of the speed of the collision.


Sub Pins_Hit (idx)
  PlaySound "fx_PinHit", 0, Vol(ActiveBall), AudioPan(ActiveBall), 0, Pitch(ActiveBall), 0, 0, AudioFade(ActiveBall)
End Sub

Sub Targets_Hit (idx)
  PlaySound "target", 0, Vol(ActiveBall), AudioPan(ActiveBall), 0, Pitch(ActiveBall), 0, 0, AudioFade(ActiveBall)
End Sub

Sub Switches_Hit (idx)
  PlaySound "fx_Sensor", 0, 1, AudioPan(ActiveBall), 0, Pitch(ActiveBall), 0, 0, AudioFade(ActiveBall)
End Sub

Sub Metals_Thin_Hit (idx)
  PlaySound "metalhit_thin", 0, Vol(ActiveBall), AudioPan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
End Sub

Sub Metals_Medium_Hit (idx)
  PlaySound "metalhit_medium", 0, Vol(ActiveBall), AudioPan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
End Sub

Sub Metals2_Hit (idx)
  PlaySound "metalhit2", 0, Vol(ActiveBall), AudioPan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
End Sub

Sub Gates_Hit (idx)
  PlaySound "GateWire", 0, Vol(ActiveBall), AudioPan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
End Sub

Sub Spinner_Spin
  PlaySound "fx_spinner", 0, .75, AudioPan(Spinner), 0.25, 0, 0, 1, AudioFade(Spinner)
End Sub

Sub Rubbers_Hit(idx)
  dim finalspeed
    finalspeed=SQR(activeball.velx * activeball.velx + activeball.vely * activeball.vely)
  If finalspeed > 20 then
    PlaySound "fx_rubber2", 0, Vol(ActiveBall), AudioPan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
  End if
  If finalspeed >= 6 AND finalspeed <= 20 then
    RandomSoundRubber()
  End If
End Sub

Sub Posts_Hit(idx)
  dim finalspeed
    finalspeed=SQR(activeball.velx * activeball.velx + activeball.vely * activeball.vely)
  If finalspeed > 16 then
    PlaySound "fx_rubber2", 0, Vol(ActiveBall), AudioPan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
  End if
  If finalspeed >= 6 AND finalspeed <= 16 then
    RandomSoundRubber()
  End If
End Sub

Sub RandomSoundRubber()
  Select Case Int(Rnd*3)+1
    Case 1 : PlaySound "rubber_hit_1", 0, Vol(ActiveBall), AudioPan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
    Case 2 : PlaySound "rubber_hit_2", 0, Vol(ActiveBall), AudioPan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
    Case 3 : PlaySound "rubber_hit_3", 0, Vol(ActiveBall), AudioPan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
  End Select
End Sub

Sub LeftFlipper_Collide(parm)
  RandomSoundFlipper()
End Sub

Sub RightFlipper_Collide(parm)
  RandomSoundFlipper()
End Sub

Sub RandomSoundFlipper()
  Select Case Int(Rnd*3)+1
    Case 1 : PlaySound "flip_hit_1", 0, Vol(ActiveBall), AudioPan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
    Case 2 : PlaySound "flip_hit_2", 0, Vol(ActiveBall), AudioPan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
    Case 3 : PlaySound "flip_hit_3", 0, Vol(ActiveBall), AudioPan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
  End Select
End Sub

'******************
' RealTime Updates
'******************

Set MotorCallback = GetRef("RealTimeUpdates")

Sub RealTimeUpdates
    BigUfoUpdate
    RollingUpdate
End Sub

' Thalamus : Exit in a clean and proper way
Sub Table1_exit()
  Controller.Pause = False
  Controller.Stop
End Sub


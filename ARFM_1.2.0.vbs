' Attack & Revenge from Mars
' Based on the tables by Bally/Williams
' Uses the ROM from Attack from Mars
' Use it with VPX
' DOF by arngrim

Option Explicit
Randomize

' Thalamus 2018-07-19
' Added/Updated "Positional Sound Playback Functions" and "Supporting Ball & Sound Functions"
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

On Error Resume Next
ExecuteGlobal GetTextFile("controller.vbs")
If Err Then MsgBox "You need the controller.vbs in order to run this table, available in the vp10 package"
On Error Goto 0

Const BallSize = 50
Const UseVPMModSol = 1

Dim DesktopMode:DesktopMode = Table1.ShowDT
Dim UseVPMColoredDMD

UseVPMColoredDMD = DesktopMode

Const UseSolenoids = 2
Const UseLamps = 0
Const UseSync = 0
Const HandleMech = 0

'
LoadVPM "01560000", "WPC.VBS", 3.26

' Thalamus - for Fast Flip v2
NoUpperRightFlipper
NoUpperLeftFlipper

'********************
'Standard definitions
'********************

' Standard Sounds
Const SSolenoidOn = "fx_Solenoidon"
Const SSolenoidOff = "fx_solenoidoff"
Const SCoin = "fx_Coin"

Set GiCallback2 = GetRef("UpdateGI")

Dim bsTrough, bsL, bsR, dtDrop, x, BallFrame, plungerIM, Mech3bank

'************
' Table init.
'************

'Const cGameName = "afm_113b" 'arcade rom - with credits
Const cGameName = "afm_113"  'home rom - free play

Sub Table1_Init
    vpmInit Me
    With Controller
        .GameName = cGameName
        If Err Then MsgBox "Can't start Game " & cGameName & vbNewLine & Err.Description:Exit Sub
        .SplashInfoLine = "Attack from Mars" & vbNewLine & "VPX table by JPSalas v1.1.1"
        .Games(cGameName).Settings.Value("rol") = 0
        .HandleKeyboard = 0
        .ShowTitle = 0
        .ShowDMDOnly = 1
        .ShowFrame = 0
        .HandleMechanics = 0
        .Hidden = DesktopMode
		NVOffset (3)
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
	.InitExitSounds SoundFX("fx_Solenoid",DOFContactors), SoundFX("fx_ballrel",DOFContactors)
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
        .InitExitSounds SoundFX("fx_Solenoid",DOFContactors), SoundFX("fx_popper",DOFContactors)
        .InitExitVariance 3, 2
    End With

    ' Right hole
    Set bsR = New cvpmTrough
    With bsR
		.size = 4
        .initSwitches Array(37)
        .Initexit sw37, 200, 24
        .InitExitSounds SoundFX("fx_Solenoid",DOFContactors), SoundFX("fx_popper",DOFContactors)
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
        .InitExitSnd SoundFX("fx_plunger",DOFContactors), SoundFX("fx_plunger",DOFContactors)
        .CreateEvents "plungerIM"
    End With

    ' Main Timer init
    PinMAMETimer.Interval = PinMAMEInterval
    PinMAMETimer.Enabled = 1

    ' Init other dropwalls - animations
    UpdateGI 0, 6:UpdateGI 1, 6:UpdateGI 2, 6
    UFORotSpeedSlow
'UfoLed.Enabled = 1

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
    PlaySoundAtVol SoundFX("fx_slingshot",DOFContactors), lemk, 1
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
    PlaySoundAtVol SoundFX("fx_slingshot",DOFContactors), remk, 1
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
Sub Bumper1_Hit:vpmTimer.PulseSw 53:PlaySoundAtVol SoundFX("fx_bumper",DOFContactors), Bumper1, VolBump:End Sub
Sub Bumper2_Hit:vpmTimer.PulseSw 54:PlaySoundAtVol SoundFX("fx_bumper",DOFContactors), Bumper2, VolBump:End Sub
Sub Bumper3_Hit:vpmTimer.PulseSw 55:PlaySoundAtVol SoundFX("fx_bumper",DOFContactors), Bumper3, VolBump:End Sub

' Drain holes, vuks & saucers
Sub Drain_Hit:PlaySoundAt "fx_drain",Drain:bsTrough.AddBall Me:End Sub
Sub sw36a_Hit:PlaySoundAtVol "fx_kicker_enter", ActiveBall,VolKick:bsL.AddBall Me:End Sub
Sub sw37a_Hit:PlaySoundAtVol "fx_kicker_enter", ActiveBall,VolKick:bsR.AddBall Me:End Sub

Dim aBall

Sub sw78_Hit
    PlaySoundAtVol "fx_balldrop", ActiveBall, 1
    Set aBall = ActiveBall:Me.TimerEnabled = 1
    vpmTimer.PulseSwitch(78), 150, "bsl.addball 0 '"
End Sub

Sub sw78_Timer
    Do While aBall.Z > 0
        aBall.Z = aBall.Z -5
        Exit Sub
    Loop
    Me.DestroyBall
    Me.TimerEnabled = 0
End Sub

Dim bBall

Sub sw37_Hit
    PlaySoundAtVol "fx_balldrop", ActiveBall, 1
    Set bBall = ActiveBall:Me.TimerEnabled = 1
    bsR.AddBall 0
End Sub

Sub sw37_Timer
    Do While bBall.Z > 0
        bBall.Z = bBall.Z -5
        Exit Sub
    Loop
    Me.DestroyBall
    Me.TimerEnabled = 0
End Sub

' Rollovers & Ramp Switches
Sub sw16_Hit:Controller.Switch(16) = 1:PlaySoundAtVol "fx_sensor", ActiveBall, 1:End Sub
Sub sw16_UnHit:Controller.Switch(16) = 0:End Sub

Sub sw26_Hit:Controller.Switch(26) = 1:PlaySoundAtVol "fx_sensor", ActiveBall, 1:End Sub
Sub sw26_UnHit:Controller.Switch(26) = 0:End Sub

Sub sw17_Hit:Controller.Switch(17) = 1:PlaySoundAtVol "fx_sensor", ActiveBall, 1:End Sub
Sub sw17_UnHit:Controller.Switch(17) = 0:End Sub

Sub sw27_Hit:Controller.Switch(27) = 1:PlaySoundAtVol "fx_sensor", ActiveBall, 1:End Sub
Sub sw27_UnHit:Controller.Switch(27) = 0:End Sub

Sub sw38_Hit:Controller.Switch(38) = 1:PlaySoundAtVol "fx_sensor", ActiveBall, 1:End Sub
Sub sw38_Unhit:Controller.Switch(38) = 0:End Sub

Sub sw48_Hit:Controller.Switch(48) = 1:PlaySoundAtVol "fx_sensor", ActiveBall, 1:End Sub
Sub sw48_Unhit:Controller.Switch(48) = 0:End Sub

Sub sw71_Hit:Controller.Switch(71) = 1:PlaySoundAtVol "fx_sensor", ActiveBall, 1:End Sub
Sub sw71_UnHit:Controller.Switch(71) = 0:End Sub

Sub sw72_Hit:Controller.Switch(72) = 1:PlaySoundAtVol "fx_sensor", ActiveBall, 1:End Sub
Sub sw72_Unhit:Controller.Switch(72) = 0:End Sub

Sub sw73_Hit:Controller.Switch(73) = 1:PlaySoundAtVol "fx_sensor", ActiveBall, 1:End Sub
Sub sw73_Unhit:Controller.Switch(73) = 0:End Sub

Sub sw74_Hit:Controller.Switch(74) = 1:PlaySoundAtVol "fx_sensor", ActiveBall, 1:End Sub
Sub sw74_Unhit:Controller.Switch(74) = 0:End Sub

Sub sw61_Hit:Controller.Switch(61) = 1:End Sub
Sub sw61_Unhit:Controller.Switch(61) = 0:End Sub

Sub sw62_Hit:Controller.Switch(62) = 1:Controller.Switch(65) = 1:End Sub
Sub sw62_Unhit:Controller.Switch(62) = 0:Controller.Switch(65) = 0:End Sub

Sub sw63_Hit:Controller.Switch(63) = 1:End Sub
Sub sw63_Unhit:Controller.Switch(63) = 0:End Sub

Sub sw64_Hit:Controller.Switch(64) = 1:End Sub
Sub sw64_Unhit:Controller.Switch(64) = 0:End Sub

Sub sw1_Hit:PlaySoundAtVol "fx_metalrolling", ActiveBall, 1:End Sub

Sub sw65_Hit:Controller.Switch(65) = 1:PlaySoundAtVol "fx_metalrolling", ActiveBall, 1:End Sub
Sub sw65_Unhit:Controller.Switch(65) = 0:End Sub

' Targets
Sub sw56_Hit:vpmTimer.PulseSw 56:PlaySoundAtVol SoundFX("fx_target",DOFContactors),ActiveBall, 1:End Sub

Sub sw57_Hit:vpmTimer.PulseSw 57:PlaySoundAtVol SoundFX("fx_target",DOFContactors),ActiveBall, 1:End Sub

Sub sw58_Hit:vpmTimer.PulseSw 58:PlaySoundAtVol SoundFX("fx_target",DOFContactors),ActiveBall, 1:End Sub

Sub sw41_Hit:vpmTimer.PulseSw 41:PlaySoundAtVol SoundFX("fx_target",DOFContactors),ActiveBall, 1:End Sub

Sub sw42_Hit:vpmTimer.PulseSw 42:PlaySoundAtVol SoundFX("fx_target",DOFContactors),ActiveBall, 1:End Sub

Sub sw43_Hit:vpmTimer.PulseSw 43:PlaySoundAtVol SoundFX("fx_target",DOFContactors),ActiveBall, 1:End Sub

Sub sw44_Hit:vpmTimer.PulseSw 44:PlaySoundAtVol SoundFX("fx_target",DOFContactors),ActiveBall, 1:End Sub

Sub sw75_Hit:vpmTimer.PulseSw 75:PlaySoundAtVol SoundFX("fx_target",DOFContactors),ActiveBall, 1:End Sub

Sub sw76_Hit:vpmTimer.PulseSw 76:PlaySoundAtVol SoundFX("fx_target",DOFContactors),ActiveBall, 1:End Sub

' 3 bank
Sub sw45_Hit:vpmTimer.PulseSw 45:a3BankShake:PlaySoundAtVol SoundFX("fx_target",DOFContactors),ActiveBall, 1:End Sub

Sub sw46_Hit:vpmTimer.PulseSw 46:a3BankShake:PlaySoundAtVol SoundFX("fx_target",DOFContactors),ActiveBall, 1:End Sub

Sub sw47_Hit:vpmTimer.PulseSw 47:a3BankShake:PlaySoundAtVol SoundFX("fx_target",DOFContactors),ActiveBall, 1:End Sub

' Droptarget

Sub sw77_Hit:PlaySoundAtVol SoundFX("fx_droptarget",DOFContactors),ActiveBall, 1:dtDrop.Hit 1:UFORotSpeedFast:End Sub

' Ramps helpers
Sub RHelp1_Hit()
    ActiveBall.VelZ = 0
    ActiveBall.VelY = 0
    ActiveBall.VelX = 0
    StopSound "fx_metalrolling"
    PlaySound "fx_ballrampdrop"
End Sub

Sub RHelp2_Hit()
    ActiveBall.VelZ = 0
    ActiveBall.VelY = 0
    ActiveBall.VelX = 0
    StopSound "fx_metalrolling"
    PlaySound "fx_ballrampdrop"
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
SolCallback(36) = "vpmSolDiverter Diverter, SoundFX(""diverter"",DOFContactor),"
SolCallback(43) = "setlamp 130,"

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

Sub SolAlien5(Enabled)
    PlaySoundAtVol "fx_rubber", Alien5, 1
    If Enabled Then
        Alien5.TransY = 20
    Else
        Alien5.TransY = 0
    End If
End Sub

Sub SolAlien6(Enabled)
    PlaySoundAtVol "fx_rubber", Alien6, 1
    If Enabled Then
        Alien6.TransY = 20
    Else
        Alien6.TransY = 0
    End If
End Sub

Sub SolAlien8(Enabled)
    PlaySoundAtVol "fx_rubber", Alien8, 1
    If Enabled Then
        Alien8.TransY = 20
    Else
        Alien8.TransY = 0
    End If
End Sub

Sub SolAlien14(Enabled)
    PlaySoundAtVol "fx_rubber", Alien14, 1
    If Enabled Then
        Alien14.TransY = 20
    Else
        Alien14.TransY = 0
    End If
End Sub

Sub SolDropTargetUp(Enabled)
    If Enabled Then
        dtDrop.DropSol_On
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
        PlaySoundAtVol "fx_flipperup", LeftFlipper1, VolFlip
        LeftFlipper.RotateToEnd
        LeftFlipper1.RotateToEnd
    Else
        PlaySoundAtVol SoundFX("fx_flipperdown",DOFContactors), LeftFlipper, VolFlip
        PlaySoundAtVol "fx_flipperdown", LeftFlipper1, VolFlip
        LeftFlipper.RotateToStart
        LeftFlipper1.RotateToStart
    End If
End Sub

Sub SolRFlipper(Enabled)
    If Enabled Then
        PlaySoundAtVol SoundFX("fx_flipperup",DOFContactors), RightFlipper, VolFlip
        PlaySoundAtVol "fx_flipperup", RightFlipper, VolFlip
        RightFlipper.RotateToEnd
        RightFlipper1.RotateToEnd
    Else
        PlaySoundAtVol SoundFX("fx_flipperdown",DOFContactors), RightFlipper, VolFlip
        PlaySoundAtVol "fx_flipperdown", RightFlipper1, VolFlip
        RightFlipper.RotateToStart
        RightFlipper1.RotateToStart
    End If
End Sub

Sub LeftFlipper_Collide(parm)
    PlaySound "fx_rubber_flipper", 0, parm / 10, -0.1, 0.25
End Sub

Sub RightFlipper_Collide(parm)
    PlaySound "fx_rubber_flipper", 0, parm / 10, 0.1, 0.25
End Sub

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
    Dim gistep, ii
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
Dim cBall
BigUfoInit

Sub BigUfoInit
    Set cBall = ckicker.createball
    ckicker.Kick 0, 0
End Sub

Sub SolUfoShake(Enabled)
    If Enabled Then
        BigUfoShake
    End If
End Sub

Sub BigUfoShake
    cball.velx = 10 + 2 * RND(1)
    cball.vely = 2 * (RND(1) - RND(1) )
End Sub

Sub BigUfoUpdate
    Dim a, b, c
    a = (ckicker.y - cball.y)
    b = (ckicker.y - cball.y) / 2
    c = cball.x - ckicker.x

    Ufo1.rotx = a
    Ufo1d.rotx = a
    'Ufo1.transx = b
    'Ufo1d.transx = b
    Ufo1.roty = c
    Ufo1d.roty = c
End Sub

'**********************************
' Small and Big UFOs Rotation Speed
'**********************************

Sub UFORotSpeedSlow()
    RotateUFO.Interval = 100
End Sub

Sub UFORotSpeedMedium()
    RotateUFO.Interval = 50
End Sub

Sub UFORotSpeedFast()
    RotateUFO.Interval = 25
End Sub

Sub RotateUFO_Timer()
    Ufo1.RotZ = (Ufo1.RotZ + 5) MOD 360
    Ufo2.RotZ = (Ufo2.RotZ - 5) MOD 360
    Ufo4.RotZ = (Ufo4.RotZ - 5) MOD 360
    Ufo5.RotZ = (Ufo5.RotZ - 5) MOD 360
    Ufo6.RotZ = (Ufo6.RotZ - 5) MOD 360
    Ufo7.RotZ = (Ufo7.RotZ - 5) MOD 360
    Ufo8.RotZ = (Ufo8.RotZ - 5) MOD 360
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
    ufo2.Roty = SmallShake
    ufo2a.Roty = SmallShake
    ufo7.Roty = SmallShake
    ufo7a.Roty = SmallShake
    ufo8.Roty = SmallShake
    ufo8a.Roty = SmallShake
    ufo6.Roty = SmallShake
    ufo6a.Roty = SmallShake
    ufo5.Roty = SmallShake
    ufo5a.Roty = SmallShake
    ufo4.Roty = SmallShake
    ufo4a.Roty = SmallShake
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
    NFadeL 11, l11
    NFadeL 12, l12
    NFadeL 13, l13
    NFadeL 14, l14
    NFadeLm 15, l15a
    NFadeLm 15, l15b
    NFadeL 15, l15
    NFadeL 16, l16
    NFadeL 17, l17
    NFadeLm 18, l18a
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
    NFadeL 88, l88
    ' ufo red lights
    NFadeL 91, l91
    NFadeL 92, l92
    'NFadeL 93, l93
    NFadeL 94, l94
    NFadeL 95, l95
    NFadeL 96, l96
    NFadeL 97, l97
    NFadeL 98, l98
    NFadeL 101, l101
    NFadeL 102, l102
    NFadeL 103, l103
    NFadeL 104, l104
    'NFadeL 105, l105
    'NFadeL 106, l106
    'NFadeL 107, l107
    'NFadeL 108, l108
	'flashers
    LightMod 117, f17
    FlashMod 117, f17a
    LightMod 118, f18
    FlashMod 118, f18a
    LightMod 119, f19
    FlashMod 119, f19a
    LightMod 120, f20
    'NFadeL 121, f21
    FlashMod 121, f21a
    LightMod 122, f22
    LightMod 123, f23
    LightMod 125, f25
    FlashMod 125, f25a
    LightMod 126, f26
    FlashMod 126, f26a
    LightMod 127, f27
    FlashMod 127, f27a
    LightMod 128, f28
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

Sub aRubbers_Hit(idx):PlaySound "fx_rubber", 0, Vol(ActiveBall)*VolRH, pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0, AudioFade(ActiveBall):End Sub
Sub aPostRubbers_Hit(idx):PlaySound "fx_postrubber", 0, Vol(ActiveBall)*VolPo, pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0, AudioFade(ActiveBall):End Sub
Sub aMetals_Hit(idx):PlaySound "fx_MetalHit", 0, Vol(ActiveBall)*VolMetal, pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0, AudioFade(ActiveBall):End Sub
Sub aPlastics_Hit(idx):PlaySound "fx_PlasticHit", 0, Vol(ActiveBall)*VolMetal, pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0, AudioFade(ActiveBall):End Sub
Sub aGates_Hit(idx):PlaySound "fx_Gate", 0, Vol(ActiveBall)*VolGates, pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0, AudioFade(ActiveBall):End Sub
Sub aWoods_Hit(idx):PlaySound "fx_Woodhit", 0, Vol(ActiveBall)*VolWood, pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0, AudioFade(ActiveBall):End Sub

' Ramp Soundss
Sub REnd1_Hit()
    StopSound "fx_metalrolling"
    PlaySound "fx_ballrampdrop", 0, 1, pan(ActiveBall)
End Sub

Sub REnd2_Hit()
    StopSound "fx_metalrolling"
    PlaySound "fx_ballrampdrop", 0, 1, pan(ActiveBall)
End Sub

Sub REnd3_Hit()
    StopSound "fx_metalrolling"
    PlaySound "fx_ballrampdrop", 0, 1, pan(ActiveBall)
End Sub

Sub RSound1_Hit:PlaySoundAtVol "fx_metalrolling", ActiveBall, 1:End Sub
Sub RSound2_Hit:PlaySoundAtVol "fx_metalrolling", ActiveBall, 1:End Sub
Sub RSound3_Hit:PlaySoundAtVol "fx_metalrolling", ActiveBall, 1:End Sub


'******************
' RealTime Updates
'******************

Set MotorCallback = GetRef("RealTimeUpdates")

Sub RealTimeUpdates
    BigUfoUpdate
    RollingUpdate
End Sub

Sub bumperLight1_Init()
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

Sub PlaySoundAtVol(sound, tableobj, Volum)
  PlaySound sound, 1, Volum, Pan(tableobj), 0,0,0, 1, AudioFade(tableobj)
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

Const tnob = 5 ' total number of balls in this table is 4, but always use a higher number here because of the timing
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
    PlaySound("fx_collide"), 0, Csng(velocity) ^2 / (VolDiv/VolCol), Pan(ball1), 0, Pitch(ball1), 0, 0, AudioFade(ball1)
End Sub


' Thalamus : Exit in a clean and proper way
Sub Table1_exit()
  Controller.Pause = False
  Controller.Stop
End Sub


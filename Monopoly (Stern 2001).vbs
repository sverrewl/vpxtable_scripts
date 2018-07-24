' ****************************
' MONOPOLY -  Stern 2001
' ****************************

' Thalamus 2018-07-24
' Added/Updated "Positional Sound Playback Functions" and "Supporting Ball & Sound Functions"
' No special SSF tweaks yet.
' This is a JP table. He often uses walls as switches so I need to be careful of using PlaySoundAt
'
'Release Notes
' 1.0 - 20170403 - VPX 10.2
'       -Used JPSalas' WOW Monopoly version 1.2 as base (playfield, plastics and some lighting changed)
'       -JPSalas' WOW Monopoly and this build used the Monopoly table from Pacdude as reference and for the script
'       -Plastic modifications from photographs found on the internet
'       -Desktop background portions from wildman backglass
'       -DOF by arngrim
'       -Bumpers from Diner (Flupper, Et al.)
'       -Script Options for LED Display, Scoop Lights & Pin Blades

Option Explicit
Randomize

' ***************
'  Table Options
' ***************
' Pin Blades -  Dark Blue (Default) or Money Mod
Const cMoneyBlades=False
' Scoop and Kicker Hole Lights - On (Default) or Off
Const cScoopLights=True
' LED display Color - Red/Grey (Default)
Const cLEDBlue=True  'Blue/Yellow


On Error Resume Next
ExecuteGlobal GetTextFile("controller.vbs")
If Err Then MsgBox "You need the controller.vbs in order to run this table, available in the vp10 package"
On Error Goto 0

Const BallSize = 50

Dim VarHidden, UseVPMColoredDMD
If Table1.ShowDT = true then
    UseVPMColoredDMD = true
    VarHidden = 1
else
    UseVPMColoredDMD = False
    VarHidden = 0
end if

LoadVPM "01120100", "sega.vbs", 3.23

'********************
'Standard definitions
'********************

Const UseSolenoids = 1
Const UseLamps = 0
Const UseSync = 0
Const UseGI = 0
Const HandleMech = 0

' Standard Sounds
Const SSolenoidOn = "fx_Solenoid"
Const SSolenoidOff = "fx_solenoidoff"
Const SCoin = "fx_Coin"

Dim bsTrough, bsChance, bsSaucerWW, bsSaucerEC, bsSaucerDE, WaterMech, BankMech, obj, plungerIM

'************
' Table init.
'************

' choose the ROM
Const cGameName = "monopole"

Sub Table1_Init
    vpmInit Me
    With Controller
        .GameName = cGameName
        If Err Then MsgBox "Can't start Game " & cGameName & vbNewLine & Err.Description:Exit Sub
        .SplashInfoLine = "JP's WOW Monopoly" & vbNewLine & "VP10 table by JPSalas v1.0"
        .Games(cGameName).Settings.Value("sound") = 1: 'ensure the sound is on
        .Games(cGameName).Settings.Value("rol") = 0
        .HandleKeyboard = 0
        .ShowTitle = 0
        .ShowDMDOnly = 1
        .ShowFrame = 0
        .HandleMechanics = True
        .Hidden = VarHidden
        On Error Resume Next
        .Run GetPlayerHWnd
        If Err Then MsgBox Err.Description
        On Error Goto 0
    End With

    ' Nudging
    vpmNudge.TiltSwitch = swTilt     'plumb tilt
    vpmNudge.Sensitivity = 5
    vpmNudge.TiltObj = Array(sw59, sw62, sw41, sw42, sw43, sw49, sw50, sw51) ' Slings and Pop Bumpers

    ' Trough
    Set bsTrough = New cvpmBallStack
    With bsTrough
        .Initsw 0, 13, 12, 11, 14, 0, 0, 0
        .InitKick BallRelease, 90, 4
        .InitAddSnd "fx_Ballrel"
        .InitEntrySnd "fx_Drain", "fx_Solenoid"
        .InitExitSnd SoundFX("fx_ballrel",DOFContactors), SoundFX("fx_Solenoid",DOFContactors)
        .CreateEvents "bsTrough", Drain
        .IsTrough = True
        .Balls = 4
    End With

    ' Chance Scoop
    Set bsChance = New cvpmBallStack
    With bsChance
        .Initsw 0, 9, 0, 0, 0, 0, 0, 0
        .InitKick sw9, 120, 7
        .InitAddSnd "fx_kicker-enter"
        .InitExitSnd SoundFX("fx_kicker",DOFContactors), SoundFX("fx_solenoid",DOFContactors)
        .InitEntrySnd "fx_kicker_enter", "fx_solenoid"
        .CreateEvents "bsChance", sw9
    End With

    ' Saucer (Electric Company)
    Set bsSaucerEC = New cvpmBallStack
    With bsSaucerEC
        .InitSaucer sw26, 26, 315, 8
        .InitExitSnd SoundFX("fx_kicker",DOFContactors), SoundFX("fx_solenoid",DOFContactors)
        .CreateEvents "bsSaucerEC", sw26
    End With

    ' Saucer (Dice Eject)
    Set bsSaucerDE = New cvpmBallStack
    With bsSaucerDE
        .InitSaucer sw52, 52, 90, 7
        .InitExitSnd SoundFX("fx_kicker",DOFContactors), SoundFX("fx_solenoid",DOFContactors)
        .CreateEvents "bsSaucerDE", sw52
    End With

    ' Saucer (Water Works)
    Set bsSaucerWW = New cvpmBallStack
    With bsSaucerWW
        .InitSaucer sw29, 29, 180, 4
        .InitAltKick 340, 11
        .KickForceVar = 3
        .CreateEvents "bsSaucerWW", sw29
    End With

    'Impulse Plunger, used as the autoplunger

    Const IMPowerSetting = 42 ' Plunger Power
    Const IMTime = 0.6        ' Time in seconds for Full Plunge
    Set plungerIM = New cvpmImpulseP
    With plungerIM
        .InitImpulseP sw16, IMPowerSetting, IMTime
        .Random 0.3
        .Switch 16
        .InitExitSnd SoundFX("fx_solenoid",DOFContactors), SoundFX("fx_solenoid",DOFContactors)
        .CreateEvents "plungerIM"
    End With

    ' Main Timer init
    PinMAMETimer.Interval = PinMAMEInterval
    PinMAMETimer.Enabled = 1

    'Misc. Initialize
    controller.switch(30) = True
    controller.switch(33) = True ' sw33-36 are optos and not handled correctly in current VPM (this corrects that)
    controller.switch(34) = True
    controller.switch(35) = True
    controller.switch(36) = True
    UpDnPost.IsDropped = True

	' Remove the cabinet rails if in FS mode
	If Table1.ShowDT = False then
		lrail.Visible = False
		rrail.Visible = False
	End If

	' Pin Blade Slection
    leftPinBlade_1.IsDropped = Not(cMoneyBlades)
    leftPinBlade_2.IsDropped = cMoneyBlades
    rightPinBlade_1.IsDropped = Not(cMoneyBlades)
    rightPinBlade_2.IsDropped = cMoneyBlades

End Sub

'**********
' Keys
'**********
Sub table1_KeyDown(ByVal Keycode)
    If keycode = PlungerKey Then Plunger.Pullback
    If keycode = LeftTiltKey Then Nudge 90, 3:PlaySound SoundFX("fx_nudge",0), 0, 1, -0.1, 0.25
    If keycode = RightTiltKey Then Nudge 270, 3:PlaySound SoundFX("fx_nudge",0), 0, 1, 0.1, 0.25
    If keycode = CenterTiltKey Then Nudge 0, 4:PlaySound SoundFX("fx_nudge",0), 0, 1, 0, 0.25
    If vpmKeyDown(keycode) Then Exit Sub
if keycode = "3" then
setlamp 119,1
setlamp 120,1
setlamp 121,1
setlamp 122,1
setlamp 123,1
setlamp 129,1
End if

End Sub

Sub table1_KeyUp(ByVal Keycode)
    If keycode = PlungerKey Then Plunger.Fire:PlaySound "fx_Plunger"
    If vpmKeyUp(keycode) Then Exit Sub
if keycode = "3" then
setlamp 119,0
setlamp 120,0
setlamp 121,0
setlamp 122,0
setlamp 123,0
setlamp 129,0
End if
End Sub

'***********
' Solenoids
'***********
SolCallback(1) = "bsTrough.SolOut" ' Trough Up-Kicker
SolCallback(2) = "Autofire"        ' AutoLaunch
'SolCallback(3)	= ""							' Lower Left  Pop (handled in the scrip)
'SolCallback(4)	= ""							' Lower Right Pop (handled in the scrip)
'SolCallback(5)	= ""							' Lower Bot   Pop (handled in the scrip)
SolCallback(6) = "BankClose"    ' Bank Close
SolCallback(7) = "DropReset"    ' DropTargetReset
SolCallback(8) = "LockKickBack" ' Lock Kicker
'SolCallback(9)	= "ULJet"						' Upper Left  Pop (handled in the scrip)
'SolCallback(10)	= "URJet"						' Upper Right Pop (handled in the scrip)
'SolCallback(11)	= "UBJet"						' Upper Bot   Pop (handled in the scrip)
SolCallback(12) = "bsChance.SolOut" ' Chance Scoop
SolCallback(13) = "BankOpen"        ' Bank Open
'SolCallback(17)	= ""							' Left  Slingshot (handled in the scrip)
'SolCallback(18)	= ""							' Right Slingshot (handled in the scrip)
SolCallback(19) = "SetLamp 119,"
SolCallback(20) = "SetLamp 120,"
SolCallback(21) = "SetLamp 121,"
SolCallback(22) = "SetLamp 122,"
SolCallback(23) = "SetLamp 123,"
SolCallback(25) = "WWMotor"           ' WaterWorks Motor (Handled by Mech Handler)
SolCallback(26) = "bsSaucerEC.SolOut" ' Electric Company
SolCallback(27) = "MRelay"            ' Motor Relay
SolCallback(28) = "bsSaucerDE.SolOut" ' Dice Eject
SolCallback(29) = "SetLamp 129,"
SolCallback(30) = "DivertLeft"        ' Left  Ramp Diverter
SolCallback(31) = "DivertRight"       ' Right Ramp Diverter
SolCallback(32) = "UpDown"            ' Top Lane Up/Down Post
'SolCallback(33)   = "LeftSave"					' Left   Outlane Post (UK)
'SolCallback(34)   = "CenterSave"					' Center Outlane Post (UK)
'SolCallback(35)   = "RightSave"					' Right  Outlane Post (UK)
' SolCallback()    = "RelayAC"				    ' AC Relay (currently no way to connect to VPM in Stern games)

Sub RelayAC(enabled):vpmNudge.SolGameOn enabled:End Sub ' Tie In Nudge to AC Relay

'*******************
' Solenoid handling
'*******************

Sub Autofire(enabled)
    If enabled Then
        PlungerIM.AutoFire
    End If
End Sub

Sub BankClose(enabled)
    If enabled Then
        BankFlipper.RotateToStart
		DOF 103, DOFPulse
    End If
End Sub

Sub DropReset(enabled)
    If enabled Then
        PlaySound SoundFX("fx_dropreset",DOFContactors)
        sw38.IsDropped = False
        Controller.Switch(38) = False
    End If
End Sub

Sub BankOpen(enabled)
    If enabled Then
        BankFlipper.RotateToEnd
		DOF 103, DOFPulse
    End If
End Sub

Sub DivertLeft(enabled)
    If enabled Then
        DivLeft.RotatetoEnd
    Else
        DivLeft.RotateToStart
    End If
End Sub

Sub DivertRight(enabled)
    If enabled Then
        DivRight.RotatetoEnd
    Else
        DivRight.RotateToStart
    End If
End Sub

Sub UpDown(enabled)
    If enabled Then
        UpDnPost.IsDropped = False
    Else
        UpDnPost.IsDropped = True
    End If
End Sub

'*****************************
' WaterWorks Flipper
' copied from Pacdude's table
'*****************************
'Motor relay - direction
dim Direction:Direction = 0

Sub MRelay(enabled)
    Direction = abs(enabled)
End Sub

' WaterWorks Motor
Sub WWMotor(enabled)
    If enabled then
        UpdateFlipper.enabled = true
    else
        UpdateFlipper.enabled = false
    End If
End Sub

' WaterWorks FlipperController
dim WWPos, WWLastPos, WWSpeed, WWCounter, WWInterval:WWCounter = 0:WWInterval = 24
WWPos = 0:WWLastPos = WWPos

Sub UpdateFlipper_Timer()
    WWSpeed = abs(controller.getmech(1) ) ' Get speed from internal mech since only it has access to that value

    ' Speed select
    If WWSpeed = 0 Then UpdateFlipper.Interval = 28
    If WWSpeed = 1 Then UpdateFlipper.Interval = 26
    If WWSpeed = 2 Then UpdateFlipper.Interval = 14
    If WWSpeed = 3 Then UpdateFlipper.Interval = 12
    If WWSpeed = 4 Then UpdateFlipper.Interval = 7
    '
    If Direction = 1 Then ' Position Counter (direction flag comes from solenoid 27)
        WWPos = WWPos + 1
    Else
        WWPos = WWPos - 1
    End If
    If WWPos > 71 Then WWPos = 0
    If WWPos < 0 Then WWPos = 71

    If WWPos >= 0 and WWPos <= 3 Then controller.switch(30) = 1:else:controller.switch(30) = 0:end if 'home position
    WWFlip(WWLastPos).Visible = False:WWFlip(WWLastPos).Enabled = False
    :WWFlip(WWPos).Visible = True:WWFlip(WWPos).Enabled = True                                        ' visual change
    If WWPos >= 3 and WWPos <= 20 Then:sw29.enabled = False:else:sw29.enabled = true                  ' saucer disabled while flipper moves over it
    If WWPos = 4 or WWPos = 5 Then:bsSaucerWW.SolOut 1:End If                                         ' kick ball out at points near flipper movement
    If WWPos = 18 or WWPos = 17 Then:bsSaucerWW.SolOutAlt 1:End If
    WWLastPos = WWPos
End Sub

'***************
'  Slingshots
'***************
Dim LStep, RStep

Sub sw59_Slingshot
    vpmTimer.PulseSw 59
    PlaySound SoundFX("fx_slingshot",DOFContactors), 0, 1, -0.05, 0.05
    LeftSling4.Visible = 1
    Lemk.RotX = 26
    LStep = 0
    sw59.TimerEnabled = 1
End Sub

Sub sw59_Timer
    Select Case LStep
        Case 1:LeftSLing4.Visible = 0:LeftSLing3.Visible = 1:Lemk.RotX = 14
        Case 2:LeftSLing3.Visible = 0:LeftSLing2.Visible = 1:Lemk.RotX = 2
        Case 3:LeftSLing2.Visible = 0:Lemk.RotX = -10:sw59.TimerEnabled = 0
    End Select
    LStep = LStep + 1
End Sub

Sub sw62_Slingshot
    vpmTimer.PulseSw 62
    PlaySound SoundFX("fx_slingshot",DOFContactors), 0, 1, 0.05, 0.05
    RightSling4.Visible = 1
    Remk.RotX = 26
    RStep = 0
    sw62.TimerEnabled = 1
End Sub

Sub sw62_Timer
    Select Case RStep
        Case 1:RightSLing4.Visible = 0:RightSLing3.Visible = 1:Remk.RotX = 14
        Case 2:RightSLing3.Visible = 0:RightSLing2.Visible = 1:Remk.RotX = 2
        Case 3:RightSLing2.Visible = 0:Remk.RotX = -10:sw62.TimerEnabled = 0
    End Select
    RStep = RStep + 1
End Sub

'***************
'   Bumpers
'***************
Sub sw49_Hit:vpmTimer.PulseSw 49:PlaySound SoundFX("fx_bumper",DOFContactors), 0, 1, -0.04, 0.05:End Sub
Sub sw50_Hit:vpmTimer.PulseSw 50:PlaySound SoundFX("fx_bumper",DOFContactors), 0, 1, -0.01, 0.05:End Sub
Sub sw51_Hit:vpmTimer.PulseSw 51:PlaySound SoundFX("fx_bumper",DOFContactors), 0, 1, -0.02, 0.05:End Sub
Sub sw43_Hit:vpmTimer.PulseSw 43:PlaySound SoundFX("fx_bumper",DOFContactors), 0, 1, -0.06, 0.05:End Sub
Sub sw41_Hit:vpmTimer.PulseSw 41:PlaySound SoundFX("fx_bumper",DOFContactors), 0, 1, -0.05, 0.05:End Sub
Sub sw42_Hit:vpmTimer.PulseSw 42:PlaySound SoundFX("fx_bumper",DOFContactors), 0, 1, -0.03, 0.05:End Sub

'*********************
' Switches & Rollovers
'*********************
Sub sw15_Hit:Controller.Switch(15) = 1:End Sub
Sub sw15_unHit:Controller.Switch(15) = 0:End Sub

Sub sw57_Hit:Controller.Switch(57) = 1:PlaySound "fx_sensor", 0, Vol(ActiveBall), pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0:End Sub
Sub sw57_unHit:Controller.Switch(57) = 0:End Sub

Sub sw58_Hit:Controller.Switch(58) = 1:PlaySound "fx_sensor", 0, Vol(ActiveBall), pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0:End Sub
Sub sw58_unHit:Controller.Switch(58) = 0:End Sub

Sub sw60_Hit:Controller.Switch(60) = 1:PlaySound "fx_sensor", 0, Vol(ActiveBall), pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0:End Sub
Sub sw60_unHit:Controller.Switch(60) = 0:End Sub

Sub sw61_Hit:Controller.Switch(61) = 1:PlaySound "fx_sensor", 0, Vol(ActiveBall), pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0:End Sub
Sub sw61_unHit:Controller.Switch(61) = 0:End Sub

Sub sw25_Hit:Controller.Switch(25) = 1:PlaySound "fx_sensor", 0, Vol(ActiveBall), pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0:End Sub
Sub sw25_unHit:Controller.Switch(25) = 0:End Sub

Sub sw27_Hit:Controller.Switch(27) = 1:PlaySound "fx_sensor", 0, Vol(ActiveBall), pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0:End Sub
Sub sw27_unHit:Controller.Switch(27) = 0:End Sub

Sub sw28_Hit:Controller.Switch(28) = 1:PlaySound "fx_sensor", 0, Vol(ActiveBall), pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0:End Sub
Sub sw28_unHit:Controller.Switch(28) = 0:End Sub

Sub sw17_Hit:Controller.Switch(17) = 1:PlaySound "fx_sensor", 0, Vol(ActiveBall), pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0:End Sub
Sub sw17_unHit:Controller.Switch(17) = 0:End Sub

Sub sw18_Hit:Controller.Switch(18) = 1:PlaySound "fx_sensor", 0, Vol(ActiveBall), pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0:End Sub
Sub sw18_unHit:Controller.Switch(18) = 0:End Sub

Sub sw19_Hit:Controller.Switch(19) = 1:PlaySound "fx_sensor", 0, Vol(ActiveBall), pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0:End Sub
Sub sw19_unHit:Controller.Switch(19) = 0:End Sub

'magnets optos
Sub sw33_Hit:Controller.Switch(33) = False:End Sub
Sub sw33_UnHit:Controller.Switch(33) = True:End Sub
Sub sw34_Hit:Controller.Switch(34) = False:End Sub
Sub sw34_UnHit:Controller.Switch(34) = True:End Sub
Sub sw35_Hit:Controller.Switch(35) = False:End Sub
Sub sw35_UnHit:Controller.Switch(35) = True:End Sub
Sub sw36_Hit:Controller.Switch(36) = False:End Sub
Sub sw36_UnHit:Controller.Switch(36) = True:End Sub

'***************
'   Targets
'***************
Sub sw44_Hit
    vpmTimer.PulseSw 44
    PlaySound SoundFX("fx_target",DOFContactors), 0, Vol(ActiveBall), pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0
End Sub

Sub sw47_Hit
    vpmTimer.PulseSw 47
    PlaySound SoundFX("fx_target",DOFContactors), 0, Vol(ActiveBall), pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0
End Sub

Sub sw46_Hit
    vpmTimer.PulseSw 46
    PlaySound SoundFX("fx_target",DOFContactors), 0, Vol(ActiveBall), pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0
End Sub

'*******************
' Police Drop Target
'*******************
Sub sw38_Hit()
    PlaySound SoundFX("fx_droptarget",DOFContactors), 0, Vol(ActiveBall), pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0
    controller.switch(38) = True
End Sub

'*********
' Spinner
'*********
Sub Spinner1_Spin:vpmTimer.PulseSw 45:PlaySound "fx_spinner", 0, 1, 0.05, 0.05:End Sub

'*********
'  Gates
'*********
Sub sw10_hit:vpmTimer.PulseSw 10:PlaySound "fx_gate", 0, 1, pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0:End Sub

Sub sw39_hit:vpmTimer.PulseSw 39:PlaySound "fx_gate", 0, 1, pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0:End Sub

Sub sw48_hit:vpmTimer.PulseSw 48:PlaySound "fx_gate", 0, 1, pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0:End Sub

Sub sw31_hit:vpmTimer.PulseSw 31:PlaySound "fx_gate", 0, 1, pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0:End Sub

'********************
'    Flippers
'********************
SolCallback(sLRFlipper) = "SolRFlipper"
SolCallback(sLLFlipper) = "SolLFlipper"

Sub SolLFlipper(Enabled)
    If Enabled Then
        PlaySound SoundFX("fx_flipperup",DOFContactors), 0, 1, -0.05, 0.05
        LeftFlipper.RotateToEnd
    Else
        PlaySound SoundFX("fx_flipperdown",DOFContactors), 0, 1, -0.05, 0.05
        LeftFlipper.RotateToStart
    End If
End Sub

Sub SolRFlipper(Enabled)
    If Enabled Then
        PlaySound SoundFX("fx_flipperup",DOFContactors), 0, 1, 0.05, 0.05
        RightFlipper.RotateToEnd
        RightFlipper1.RotateToEnd
    Else
        PlaySound SoundFX("fx_flipperdown",DOFContactors), 0, 1, 0.05, 0.05
        RightFlipper.RotateToStart
        RightFlipper1.RotateToStart
    End If
End Sub

Sub LeftFlipper_Collide(parm)
    PlaySound "fx_rubber_flipper", 0, parm / 10, -0.05, 0.05
End Sub

Sub RightFlipper_Collide(parm)
    PlaySound "fx_rubber_flipper", 0, parm / 10, 0.05, 0.05
End Sub

Sub RightFlipper1_Collide(parm)
    PlaySound "fx_rubber_flipper", 0, parm / 10, 0.05, 0.05
End Sub

Sub BankFlipper_Collide(parm)
    PlaySound SoundFXDOF("fx_woodhit",104,DOFPulse,DOFShaker), 0, parm / 10, 0.05, 0.05
End Sub

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
    NFadeL 1, l1
    NFadeL 2, l2
    NFadeL 3, l3
    NFadeL 4, l4
    NFadeL 5, l5
    NFadeL 6, l6
    NFadeL 7, l7
    NFadeL 8, l8
    NFadeL 9, l9
    NFadeL 10, l10
    NFadeL 11, l11
    NFadeL 12, l12
    NFadeL 13, l13
    NFadeL 14, l14
    NFadeL 15, l15
    NFadeL 16, l16
    NFadeL 17, l17
    NFadeL 18, l18
    NFadeL 19, l19
    NFadeL 20, l20
    NFadeL 21, l21
    NFadeL 22, l22
    NFadeL 23, l23
    NFadeL 24, l24
    NFadeL 25, l25
    NFadeL 26, l26
    NFadeL 27, l27
    NFadeL 28, l28
    NFadeL 29, l29
    NFadeL 30, l30
    NFadeL 31, l31
    NfadeL 32, l32
    NFadeL 33, l33
    NFadeL 34, l34
    NFadeL 35, l35
    NFadeL 36, l36
    NFadeL 37, l37
    NFadeL 38, l38
    NFadeL 39, l39
    NFadeL 40, l40
    NFadeLm 41, l41a
    NFadeL 41, l41
    NFadeLm 42, l42a
    NFadeL 42, l42
    Flash 43, l43
    Flash 44, l44
    NFadeL 45, l45
    NFadeLm 46, l46a
    NFadeL 46, l46
    NFadeLm 47, l47
    Flash 47, l47a
    Flash 48, l48
    NFadeLm 49, l49
    Flash 49, l49a
    NFadeL 50, l50
    NFadeL 51, l51
    NFadeL 52, l52
    NFadeL 53, l53
    NFadeL 54, l54
    NFadeL 55, l55
    NFadeL 56, l56
    NFadeL 57, l57
    NFadeL 58, l58
    NFadeL 59, l59
    NFadeL 60, l60a
    NFadeL 60, l60
    NFadeL 61, l61a
    NFadeL 61, l61
    NFadeL 62, l62a
    NFadeL 62, l62
    Flash 63, l63
    Flash 64, l64
    NFadeL 65, l65
    NFadeL 66, l66
    NFadeL 67, l67
    NFadeL 68, l68
    NFadeL 69, l69
    NFadeL 70, l70
    NFadeL 71, l71
    NFadeL 72, l72
    NFadeL 73, l73
    NFadeL 74, l74
    NFadeL 75, l75
    NFadeL 76, l76
    NFadeL 77, l77
    NFadeL 78, l78
    NFadeL 79, l79
    NFadeL 80, l80

    'Flashers
    Flashm 119, f19a
    Flash 119, f19
    Flashm 120, f20a
    Flashm 120, f20b
    Flash 120, f20
    NFadeLm 121, f21b
    Flashm 121, f21a
    Flashm 121, f21c
    Flashm 121, f21d
    Flash 121, f21
    NFadeLm 122, f22b
    Flashm 122, f22a
    Flashm 122, f22c
    Flash 122, f22
    Flashm 123, f23a
    Flash 123, f23
    Flashm 129, f29a
    Flash 129, f29

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
        Case 6, 7, 8:FadingLevel(nr) = FadingLevel(nr) + 1             'wait
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

' Desktop Objects: Reels & texts (you may also use lights on the desktop)
' Reels
Sub FadeR(nr, object)
    Select Case FadingLevel(nr)
        Case 4:object.SetValue 1:FadingLevel(nr) = 6                   'fading to off...
        Case 5:object.SetValue 0:FadingLevel(nr) = 1                   'ON
        Case 6, 7, 8:FadingLevel(nr) = FadingLevel(nr) + 1              'wait
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
Sub aRubbers_Hit(idx):PlaySound "fx_rubber", 0, Vol(ActiveBall), pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0, AudioFade(ActiveBall):End Sub
Sub aPostRubbers_Hit(idx):PlaySound "fx_postrubber", 0, Vol(ActiveBall), pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0, AudioFade(ActiveBall):End Sub
Sub aMetals_Hit(idx):PlaySound "fx_MetalHit", 0, Vol(ActiveBall), pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0, AudioFade(ActiveBall):End Sub
Sub aPlastics_Hit(idx):PlaySound "fx_PlasticHit", 0, Vol(ActiveBall), pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0, AudioFade(ActiveBall):End Sub
Sub aGates_Hit(idx):PlaySound "fx_Gate", 0, Vol(ActiveBall), pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0, AudioFade(ActiveBall):End Sub
Sub aWoods_Hit(idx):PlaySound "fx_Woodhit", 0, Vol(ActiveBall), pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0, AudioFade(ActiveBall):End Sub

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

Sub RSound1_Hit:PlaySound "fx_metalrolling", 0, 0.2, pan(ActiveBall):End Sub
Sub RSound2_Hit:PlaySound "fx_metalrolling", 0, 0.2, pan(ActiveBall):End Sub
Sub RSound3_Hit:PlaySound "fx_metalrolling", 0, 0.2, pan(ActiveBall):End Sub

'****************************************
' GI - needs new vpinmame (version 2.8+)
'****************************************
Set GICallback = GetRef("GIUpdate")

Sub GIUpdate(no, Enabled)
    For each obj in aGiTopLights
        obj.State = ABS(Enabled)
    Next
    For each obj in aGiBottomLights
        obj.State = ABS(Enabled)
    Next
    For each obj in aFlasherBackWall
        obj.Visible = ABS(Enabled)
    Next
    If cScoopLights=True Then
        For each obj in aScoopLights
            obj.State = ABS(Enabled)
        Next
    End If
End Sub

'****** OLDER VERSION FOR VPINMAME 2.7
'******************
'   GI effects
'******************
'Dim OldGiState
'OldGiState = 2 'start witht the GI off
'
'Sub GIUpdate
'    Dim tmp, obj
'    tmp = Getballs
'    If UBound(tmp) <> OldGiState Then
'        OldGiState = Ubound(tmp)
'        If UBound(tmp) = -1 Then
'            For each obj in aGiTopLights
'                obj.State = 0
'            Next
'            For each obj in aGiBottomLights
'                obj.State = 0
'            Next
'            For each obj in aFlasherBackWall
'                obj.Visible = 0
'            Next
'        Else
'            For each obj in aGiTopLights
'                obj.State = 1
'            Next
'            For each obj in aGiBottomLights
'                obj.State = 1
'            Next
'            For each obj in aFlasherBackWall
'                obj.Visible = 1
'            Next
'        End If
'    End If
'End Sub
'
'Sub GIOn
'    For each obj in aGiTopLights
'        obj.State = 1
'    Next
'    For each obj in aGiBottomLights
'        obj.State = 1
'    Next
'End Sub
'
'Sub GIOff
'    For each obj in aGiTopLights
'        obj.State = 0
'    Next
'    For each obj in aGiBottomLights
'        obj.State = 0
'    Next
'End Sub

'******************
' real time updates
'******************
Set MotorCallback = GetRef("RealTimeUpdates")

Sub RealTimeUpdates
    BankDoor.RotY = - BankFlipper.CurrentAngle + 94
    GIUpdate
    RollingUpdate
End Sub

'***************
'Mini DMD Array
'***************
Dim LED(15)

LED(0) = Array(d1, d2, d3, d4, d5, d6, d7)
LED(1) = Array(d8, d9, d10, d11, d12, d13, d14)
LED(2) = Array(d15, d16, d17, d18, d19, d20, d21)
LED(3) = Array(d22, d23, d24, d25, d26, d27, d28)
LED(4) = Array(d29, d30, d31, d32, d33, d34, d35)
LED(5) = Array(d36, d37, d38, d39, d40, d41, d42)
LED(6) = Array(d43, d44, d45, d46, d47, d48, d49)
LED(7) = Array(d50, d51, d52, d53, d54, d55, d56)
LED(8) = Array(d57, d58, d59, d60, d61, d62, d63)
LED(9) = Array(d64, d65, d66, d67, d68, d69, d70)
LED(10) = Array(d71, d72, d73, d74, d75, d76, d77)
LED(11) = Array(d78, d79, d80, d81, d82, d83, d84)
LED(12) = Array(d85, d86, d87, d88, d89, d90, d91)
LED(13) = Array(d92, d93, d94, d95, d96, d97, d98)
LED(14) = Array(d99, d100, d101, d102, d103, d104, d105)

Sub LedTimer_Timer
    Dim ChgLED, ii, num, chg, stat, obj, objf, x, y
    ChgLED = Controller.ChangedLEDs(&H00000000, &Hffffffff)
    If Not IsEmpty(ChgLED) Then
        For ii = 0 To UBound(chgLED)
            num = chgLED(ii, 0):chg = chgLED(ii, 1):stat = chgLED(ii, 2)
            For y = 0 to 6
                If chg And 1 Then
                    If(stat and 1) = 1 Then
                        If cLEDBlue=True Then
                            Led(num) (y).color = RGB(1, 1, 255)
                        Else
                            Led(num) (y).color = RGB(255, 0, 0)
                        End If
                    Else
                        If cLEDBlue=True Then
                            led(num) (y).color = RGB(255, 244, 1)
                        Else
                            led(num) (y).color = RGB(92, 92, 92)
                        End If
                    End If
                End If
                chg = chg \ 4:stat = stat \ 4
            Next
        Next
    End If
End Sub

'************
'  3 Locks
'************
Sub Lock1_Hit() 'lower lock
    Controller.Switch(22) = 1
    Lock2.Enabled = 1
End Sub

Sub Lock2_Hit() 'middle lock
    Controller.Switch(21) = 1
    Lock3.Enabled = 1
End Sub

Sub Lock3_Hit() 'top lock
    Controller.Switch(20) = 1
End Sub

Sub LockKickBack(enabled)
    If Enabled then
        PlaySound SoundFX("fx_kicker",DOFContactors), 0, 1, -0.2, 0.05
        ' release top lock
        Lock3.enabled = False
        Lock3.Kick 0, 32
        Controller.Switch(20) = 0
        ' release middle lock
        Lock2.enabled = False
        Lock2.Kick 0, 32
        Controller.Switch(21) = 0
        ' release lower lock
        Lock1.Kick 0, 32
        Controller.Switch(22) = 0
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
Const tnob = 5 ' total number of balls in this table is 4, but always use a higher number here bacuse of the timing
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


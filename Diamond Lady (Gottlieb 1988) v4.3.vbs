' Diamond Lady / IPD No. 678 / Premier February, 1988 / 4 Players
' VPX7 v4.3.0 by JPSalas 2023
' Solenoids based on Destruk's script
' DOF commands by Arngrim

Option Explicit
Randomize

On Error Resume Next
ExecuteGlobal GetTextFile("controller.vbs")
If Err Then MsgBox "You need the controller.vbs in order to run this table, available in the vp10 package"
On Error Goto 0

LoadVPM "01210000", "sys80.vbs", 3.1

Dim bsTrough, dtLBank, dtRBank, dtMBank, dtCBank, bsTop, kickbackIM
Dim x, bump1, bump2, bump3

Const cGameName = "diamond"

Const UseSolenoids = 1
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
    lrail.Visible = 0
    rrail.Visible = 0
end if

if B2SOn = true then VarHidden = 1

' Standard Sounds
Const SSolenoidOn = "fx_Solenoid"
Const SSolenoidOff = ""
Const SCoin = "fx_coin"

'************
' Table init.
'************

Sub table1_Init
    With Controller
        .GameName = cGameName
        .Games(cGameName).Settings.Value("sound") = 1   'ensure the sound is on
        .SplashInfoLine = "Diamond Lady, Gottlieb 1988" & vbNewLine & "VPX7 table by JPSalas v.4.3.0"
        .Games(cGameName).Settings.Value("dmd_red") = 0 'set dmd lights to color blue
        .Games(cGameName).Settings.Value("dmd_green") = 128
        .Games(cGameName).Settings.Value("dmd_blue") = 255
        .Games(cGameName).Settings.Value("rol") = 0
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
    vpmTimer.AddTimer 2000, "Controller.SolMask(0)=&Hffffffff'" 'ignore all solenoids - then add the Timer to renable all the solenoids after 2 seconds
    Controller.Run GetPlayerHWnd

' Thalamus : Was missing 'vpminit me'
  vpminit me

    ' Nudging
    vpmNudge.TiltSwitch = 57
    vpmNudge.Sensitivity = 1
    vpmNudge.TiltObj = Array(Bumper1, Bumper2, Bumper3, LeftSlingshot, RightSlingshot)

    ' Trough & Ball Release
    Set bsTrough = New cvpmBallStack
    With bsTrough
        .InitSw 66, 0, 46, 0, 0, 0, 0, 0
        .InitKick BallRelease, 80, 6
        .InitExitSnd SoundFX("fx_ballrel",DOFContactors), SoundFX("fx_Solenoid",DOFContactors)
        .Balls = 2
    End With

    ' Top saucer
    Set bsTop = New cvpmBallStack
    With bsTop
        .InitSaucer sw74, 74, 300, 28
        .InitExitSnd SoundFX("fx_kicker", DOFContactors), SoundFX("fx_Solenoid", DOFContactors)
    End With

    ' Left Drop targets
    set dtLBank = new cvpmdroptarget
    With dtLBank
        .InitDrop Array(sw20, sw30, sw40, sw50, sw60), Array(20, 30, 40, 50, 60)
        .initsnd SoundFX("fx_droptarget", DOFDropTargets), SoundFX("fx_resetdrop", DOFContactors)
        .CreateEvents "dtLBank"
    End With

    ' Middle Drop targets
    set dtMBank = new cvpmdroptarget
    With dtMBank
        .InitDrop Array(sw21, sw31, sw41, sw51), Array(21, 31, 41, 51)
        .initsnd SoundFX("fx_droptarget", DOFDropTargets), SoundFX("fx_resetdrop", DOFContactors)
        .CreateEvents "dtMBank"
    End With

    ' Right Drop targets
    set dtRBank = new cvpmdroptarget
    With dtRBank
        .InitDrop Array(sw22, sw32, sw42, sw52, sw62), Array(22, 32, 42, 52, 62)
        .initsnd SoundFX("fx_droptarget", DOFDropTargets), SoundFX("fx_resetdrop", DOFContactors)
        .CreateEvents "dtRBank"
    End With

    ' Center Drop target
    set dtCBank = new cvpmdroptarget
    With dtCBank
        .InitDrop sw23, 23
        .initsnd SoundFX("fx_droptarget", DOFDropTargets), SoundFX("fx_resetdrop", DOFContactors)
        .CreateEvents "dtCBank"
    End With

    ' Impulse Plunger used as the left kickback
    Set kickbackIM = New cvpmImpulseP
    With kickbackIM
        .InitImpulseP swKickback, 26, 0.4
        .Random 0.6
        '.Switch 43
        .InitExitSnd "fx_popper", "fx_popper"
        .CreateEvents "kickbackIM"
    End With

    ' Main Timer init
    PinMAMETimer.Interval = PinMAMEInterval
    PinMAMETimer.Enabled = 1

    ' Start the RealTime timer
    RealTime.Enabled = 1

    ' Load table color
    LoadLut
End Sub

Sub table1_Paused:Controller.Pause = 1:End Sub
Sub table1_unPaused:Controller.Pause = 0:End Sub

'*******************
' Realtime updates
'*******************

Sub Realtime_Timer
    RollingUpdate
    GIUpdate
    LeftFlipperTop.RotZ = LeftFlipper.CurrentAngle
    RightFlipperTop.RotZ = RightFlipper.CurrentAngle
End Sub

'**********
' Keys
'**********

Sub table1_KeyDown(ByVal Keycode)
    If KeyCode = RightFlipperKey Then Controller.Switch(72) = 1
    If keycode = LeftTiltKey Then Nudge 90, 5:PlaySound SoundFX("fx_nudge", 0), 0, 1, -0.1, 0.25
    If keycode = RightTiltKey Then Nudge 270, 5:PlaySound SoundFX("fx_nudge", 0), 0, 1, 0.1, 0.25
    If keycode = CenterTiltKey Then Nudge 0, 6:PlaySound SoundFX("fx_nudge", 0), 0, 1, 0, 0.25
    If keycode = LeftMagnaSave Then bLutActive = True: SetLUTLine "Color LUT image " & table1.ColorGradeImage
    If keycode = RightMagnaSave AND bLutActive Then NextLUT:End If
    If vpmKeyDown(keycode)Then Exit Sub
    If keycode = KeyRules Then Rules
    If keycode = PlungerKey Then PlaySound "fx_PlungerPull", 0, 1, 0.1, 0.25:Plunger.Pullback
End Sub

Sub table1_KeyUp(ByVal Keycode)
    If keycode = LeftMagnaSave Then bLutActive = False: HideLUT
    If KeyCode = RightFlipperKey Then Controller.Switch(72) = 0
    If vpmKeyUp(keycode)Then Exit Sub
    If keycode = PlungerKey Then PlaySound "fx_plunger", 0, 1, 0.1, 0.25:Plunger.Fire
End Sub

'************************************
'       LUT - Darkness control
' 10 normal level & 10 warmer levels
'************************************

Dim bLutActive, LUTImage

Sub LoadLUT
    bLutActive = False
    x = LoadValue(cGameName, "LUTImage")
    If(x <> "")Then LUTImage = x Else LUTImage = 0
    UpdateLUT
End Sub

Sub SaveLUT
    SaveValue cGameName, "LUTImage", LUTImage
End Sub

Sub NextLUT:LUTImage = (LUTImage + 1)MOD 22:UpdateLUT:SaveLUT:SetLUTLine "Color LUT image " & table1.ColorGradeImage:End Sub

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

Dim GiIntensity
GiIntensity = 1   'can be used by the LUT changing to increase the GI lights when the table is darker

Sub ChangeGiIntensity(factor) 'changes the intensity scale
    Dim bulb
    For each bulb in aGiLights
        bulb.IntensityScale = GiIntensity * factor
    Next
End Sub

' New LUT postit
Function GetHSChar(String, Index)
    Dim ThisChar
    Dim FileName
    ThisChar = Mid(String, Index, 1)
    FileName = "PostIt"
    If ThisChar = " " or ThisChar = "" then
        FileName = FileName & "BL"
    ElseIf ThisChar = "<" then
        FileName = FileName & "LT"
    ElseIf ThisChar = "_" then
        FileName = FileName & "SP"
    Else
        FileName = FileName & ThisChar
    End If
    GetHSChar = FileName
End Function

Sub SetLUTLine(String)
    Dim Index
    Dim xFor
    Index = 1
    LUBack.imagea="PostItNote"
    For xFor = 1 to 40
        Eval("LU" &xFor).imageA = GetHSChar(String, Index)
        Index = Index + 1
    Next
End Sub

Sub HideLUT
SetLUTLine ""
LUBack.imagea="PostitBL"
End Sub

'*********
' Switches
'*********

' Slings & div switches
Dim LStep, RStep

Sub LeftSlingShot_Slingshot
    PlaySoundAt SoundFX("fx_slingshot", DOFContactors), Lemk
    DOF 103, DOFPulse
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
    PlaySoundAt SoundFX("fx_slingshot", DOFContactors),Remk
    DOF 104, DOFPulse
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

' Bumpers
Sub Bumper1_Hit:vpmTimer.PulseSw 61:PlaySoundAt SoundFX("fx_bumper",DOFContactors), Bumper1:DOF 105, DOFPulse:End Sub
Sub Bumper2_Hit:vpmTimer.PulseSw 71:PlaySoundAt SoundFX("fx_bumper",DOFContactors), Bumper2:End Sub
Sub Bumper3_Hit:vpmTimer.PulseSw 61:PlaySoundAt SoundFX("fx_bumper",DOFContactors), Bumper3:DOF 107, DOFPulse:End Sub

' Drain holes
Sub Drain_Hit:PlaySoundAt "fx_drain", Drain:bsTrough.AddBall Me:End Sub

' Rollovers
Sub sw43_Hit:Controller.Switch(43) = 1:PlaySoundAt "fx_sensor",sw43:End Sub
Sub sw43_UnHit:Controller.Switch(43) = 0:End Sub

Sub sw53_Hit:Controller.Switch(53) = 1:PlaySoundAt "fx_sensor",sw53:End Sub
Sub sw53_UnHit:Controller.Switch(53) = 0:End Sub

Sub sw63_Hit:Controller.Switch(63) = 1:PlaySoundAt "fx_sensor",sw63:End Sub
Sub sw63_UnHit:Controller.Switch(63) = 0:End Sub

Sub sw73_Hit:Controller.Switch(73) = 1:PlaySoundAt "fx_sensor",sw73:End Sub
Sub sw73_UnHit:Controller.Switch(73) = 0:End Sub

Sub sw65_Hit:Controller.Switch(65) = 1:PlaySoundAt "fx_sensor",sw65:End Sub
Sub sw65_UnHit:Controller.Switch(65) = 0:End Sub

Sub sw70_Hit:Controller.Switch(70) = 1:PlaySoundAt "fx_sensor",sw70:End Sub
Sub sw70_UnHit:Controller.Switch(70) = 0:End Sub

Sub sw45_Hit:Controller.Switch(45) = 1:PlaySoundAt "fx_sensor",sw45:End Sub
Sub sw45_UnHit:Controller.Switch(45) = 0:End Sub

Sub sw55_Hit:Controller.Switch(55) = 1:PlaySoundAt "fx_sensor",sw55:End Sub
Sub sw55_UnHit:Controller.Switch(55) = 0:End Sub

' Ramp Switches
Sub sw44_Hit():Controller.Switch(44) = 1:End Sub
Sub sw44_UnHit:Controller.Switch(44) = 0:End Sub

'Spinners
Sub sw64_Spin:vpmTimer.PulseSw 64:PlaySoundAt "fx_spinner", sw64:End Sub
Sub sw54_Spin:vpmTimer.PulseSw 54:PlaySoundAt "fx_spinner", sw54:End Sub

' Targets
Sub sw24_Hit:vpmTimer.PulseSw 24:PlaySoundAtBall SoundFX("fx_target", DOFTargets):End Sub

Sub sw34_Hit:vpmTimer.PulseSw 34:PlaySoundAtBall SoundFX("fx_target", DOFTargets):End Sub

Sub sw26_Hit:vpmTimer.PulseSw 26:PlaySoundAtBall SoundFX("fx_target", DOFTargets):End Sub

Sub sw36_Hit:vpmTimer.PulseSw 36:PlaySoundAtBall SoundFX("fx_target", DOFTargets):End Sub

Sub sw25_Hit:vpmTimer.PulseSw 25:PlaySoundAtBall SoundFX("fx_target", DOFTargets):End Sub

Sub sw35_Hit:vpmTimer.PulseSw 35:PlaySoundAtBall SoundFX("fx_target", DOFTargets):End Sub

' VUK & Holes

Sub sw74_Hit:PlaySoundAt "fx_kicker_enter", sw74:bsTop.AddBall 0:End Sub

'********************
'Solenoids & Flashers
'********************
' old solenoid list
'SolCallback(2) = "dtMBank.SolDropUp"
'SolCallback(3) = "SetLamp 103," 'Left Spinner Flasher
'SolCallback(4) = "SetLamp 104," 'Right Orange Flashers
'SolCallback(5) = "dtRBank.SolDropUp"
'SolCallback(6) = "dtLBank.SolDropUp"
'SolCallback(7) = "SetLamp 107," 'Right Red Flashers
'SolCallback(8) = "vpmSolSound SoundFX(""fx_knocker"",DOFKnocker),"
'SolCallback(9) = "bsTrough.SolIn"
'SolCallback(10) = "dtCBank.SolDropUp"

SolCallback(1) = "dtCBank.SolDropUp"
SolCallback(2) = "dtLBank.SolDropUp"
SolCallback(3) = "SetLamp 103," 'Left Spinner Flasher
SolCallback(4) = "SetLamp 104," 'Right Orange Flashers
SolCallback(5) = "dtMBank.SolDropUp"
SolCallback(6) = "dtRBank.SolDropUp"
SolCallback(7) = "SetLamp 107," 'Right Red Flashers
SolCallback(8) = "vpmSolSound SoundFX(""fx_knocker"",DOFKnocker),"
SolCallback(9) = "bsTrough.SolIn"

'*******************
' Flipper Subs Rev3
'*******************

SolCallback(sLRFlipper) = "SolRFlipper"
SolCallback(sLLFlipper) = "SolLFlipper"

Sub SolLFlipper(Enabled)
    If Enabled Then
        PlaySoundAt SoundFX("fx_flipperup", DOFContactors), LeftFlipper
        LeftFlipper.RotateToEnd
        LeftFlipperOn = 1
    Else
        PlaySoundAt SoundFX("fx_flipperdown", DOFContactors), LeftFlipper
        LeftFlipper.RotateToStart
        LeftFlipperOn = 0
    End If
End Sub

Sub SolRFlipper(Enabled)
    If Enabled Then
        PlaySoundAt SoundFX("fx_flipperup", DOFContactors), RightFlipper
        RightFlipper.RotateToEnd
        RightFlipperOn = 1
    Else
        PlaySoundAt SoundFX("fx_flipperdown", DOFContactors), RightFlipper
        RightFlipper.RotateToStart
        RightFlipperOn = 0
    End If
End Sub

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

FlipperPower = 5000
FlipperElasticity = 0.8
FullStrokeEOS_Torque = 0.3  ' EOS Torque when flipper hold up ( EOS Coil is fully charged. Ampere increase due to flipper can't move or when it pushed back when "On". EOS Coil have more power )
LiveStrokeEOS_Torque = 0.2  ' EOS Torque when flipper rotate to end ( When flipper move, EOS coil have less Ampere due to flipper can freely move. EOS Coil have less power )

LeftFlipper.EOSTorqueAngle = 10
RightFlipper.EOSTorqueAngle = 10

SOSTorque = 0.1
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

'*****************************
'Extra Lamps used as solenoids
'Based on destruk's code
'adapted to the fading lights
'*****************************

Dim NewL12, OldL12, NewL2, OldL2, NewL13, OldL13, NewL14, OldL14, NewL15, OldL15, NewL16, OldL16, NewL17, OldL17, NewL18, OldL18
OldL12 = 0:OldL2 = 0:OldL13 = 0:OldL14 = 0:OldL15 = 0:OldL16 = 0:OldL17 = 0:OldL18 = 0

Set LampCallback = GetRef("ExtraLamps")
Sub ExtraLamps
    'Aux Relay
    NewL12 = Controller.Lamp(12)
    If NewL12 <> OldL12 Then
        If NewL12 Then
            Aux.Enabled = 1
        Else
            Aux.Enabled = 0
            SetLamp 111, 0:SetLamp 112, 0:SetLamp 113, 0:SetLamp 114, 0:SetLamp 115, 0
        End If
    End If
    OldL12 = NewL12

    'Ball Release
    NewL2 = Controller.Lamp(2)
    If NewL2 <> OldL2 Then
        If NewL2 Then bsTrough.ExitSol_On
    End If
    OldL2 = NewL2

    'Top Kicker
    NewL13 = Controller.Lamp(13)
    If NewL13 <> OldL13 Then
        If NewL13 Then bsTop.ExitSol_On
    End If
    OldL13 = NewL13

    'Kickback
    NewL14 = Controller.Lamp(14)
    If NewL14 <> OldL14 Then
        If NewL14 Then kickbackIM.AutoFire:DOF 108, DOFPulse
    End If
    OldL14 = NewL14

    '#1 Drop Target Trip Coil
    NewL15 = Controller.Lamp(15)
    If NewL15 <> OldL15 Then
        If NewL15 Then
            dtLBank.Hit 2
            dtLBank.Hit 3
            dtLBank.Hit 4
        End If
    End If
    OldL15 = NewL15

    '#2 Drop Target Trip Coil
    NewL16 = Controller.Lamp(16)
    If NewL16 <> OldL16 Then
        If NewL16 Then
            dtMBank.Hit 2
            dtMBank.Hit 3
        End If
    End If
    OldL16 = NewL16

    '#3 Drop Target Trip Coil
    NewL17 = Controller.Lamp(17)
    If NewL17 <> OldL17 Then
        If NewL17 Then
            dtRBank.Hit 2
            dtRBank.Hit 3
            dtRBank.Hit 4
        End If
    End If
    OldL17 = NewL17

    '#4 Drop Target Trip Coil
    NewL18 = Controller.Lamp(18)
    If NewL18 <> OldL18 Then
        If NewL18 Then dtCBank.Hit 1
    End If
    OldL18 = NewL18
End Sub

Dim AuxCount:AuxCount = 0
Sub Aux_Timer
    Select Case AuxCount
        Case 0:SetLamp 115, 1:SetLamp 111, 0
        Case 1:SetLamp 114, 1:SetLamp 115, 0
        Case 2:SetLamp 113, 1:SetLamp 114, 0
        Case 3:SetLamp 112, 1:SetLamp 113, 0
        Case 4:SetLamp 111, 1:SetLamp 112, 0
    End Select
    AuxCount = AuxCount + 1
    If AuxCount = 5 then AuxCount = 0
End Sub

'*****************
'   Gi Effects
'*****************

Dim OldGiState
OldGiState = -1 'start witht he Gi off

Sub GiON
    For each x in aGiLights
        GiEffect
    Next
End Sub

Sub GiOFF
    For each x in aGiLights
        x.State = 0
    Next
End Sub

Sub GiEffect
    For each x in aGiLights
        x.Duration 2, 1000, 1
    Next
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

Sub UpdateLamps
    Lampm 3, l3a
    Lamp 3, l3
    Lamp 5, l5
    Lamp 6, l6
    Lamp 7, l7
    Lamp 8, l8
    Lamp 9, l9
    Lamp 10, l10
    Lamp 11, l11
    Lampm 19, l19b
    Lamp 19, l19
    Lampm 20, l20a
    Lamp 20, l20
    Flash 21, l21
    Flash 22, l22
    Flash 23, l23
    Flash 24, l24
    Flash 25, l25
    Flash 26, l26
    Flash 27, l27
    Lamp 28, l28
    Lamp 29, l29
    Lamp 30, l30
    Lampm 31, l31
    Lampm 31, l31b
    Lamp 31, l31d
    Lamp 32, l32
    Lamp 33, l33
    Lamp 34, l34
    Lamp 35, l35
    Lamp 36, l36
    Lamp 37, l37
    Lamp 38, l38
    Lamp 39, l39
    Lamp 41, l41
    Lamp 42, l42
    Lamp 43, l43
    Lamp 44, l44
    Lampm 45, l45
    Lamp 45, l45b
    Lamp 46, l46
    Lamp 47, l47
    Lamp 51, l51

    'flashers
    Lampm 103, f31
    Lamp 103, f3
    Flashm 104, f41
    Flash 104, f4
    Flashm 107, f71
    Flash 107, f7

    ' Aux lights
    Lamp 111, Aux1
    Lamp 112, Aux2
    Lamp 113, Aux3
    Lamp 114, Aux4
    Lamp 115, Aux5
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

'*****************
' Leds Display
'*****************

Dim Digits(40)

Digits(0) = Array(a00, a05, a0c, a0d, a08, a01, a06, a0f, a02, a03, a04, a07, a0b, a0a, a09, a0e)
Digits(1) = Array(a10, a15, a1c, a1d, a18, a11, a16, a1f, a12, a13, a14, a17, a1b, a1a, a19, a1e)
Digits(2) = Array(a20, a25, a2c, a2d, a28, a21, a26, a2f, a22, a23, a24, a27, a2b, a2a, a29, a2e)
Digits(3) = Array(a30, a35, a3c, a3d, a38, a31, a36, a3f, a32, a33, a34, a37, a3b, a3a, a39, a3e)
Digits(4) = Array(a40, a45, a4c, a4d, a48, a41, a46, a4f, a42, a43, a44, a47, a4b, a4a, a49, a4e)
Digits(5) = Array(a50, a55, a5c, a5d, a58, a51, a56, a5f, a52, a53, a54, a57, a5b, a5a, a59, a5e)
Digits(6) = Array(a60, a65, a6c, a6d, a68, a61, a66, a6f, a62, a63, a64, a67, a6b, a6a, a69, a6e)
Digits(7) = Array(a70, a75, a7c, a7d, a78, a71, a76, a7f, a72, a73, a74, a77, a7b, a7a, a79, a7e)
Digits(8) = Array(a80, a85, a8c, a8d, a88, a81, a86, a8f, a82, a83, a84, a87, a8b, a8a, a89, a8e)
Digits(9) = Array(a90, a95, a9c, a9d, a98, a91, a96, a9f, a92, a93, a94, a97, a9b, a9a, a99, a9e)
Digits(10) = Array(aa0, aa5, aac, aad, aa8, aa1, aa6, aaf, aa2, aa3, aa4, aa7, aab, aaa, aa9, aae)
Digits(11) = Array(ab0, ab5, abc, abd, ab8, ab1, ab6, abf, ab2, ab3, ab4, ab7, abb, aba, ab9, abe)
Digits(12) = Array(ac0, ac5, acc, acd, ac8, ac1, ac6, acf, ac2, ac3, ac4, ac7, acb, aca, ac9, ace)
Digits(13) = Array(ad0, ad5, adc, add, ad8, ad1, ad6, adf, ad2, ad3, ad4, ad7, adb, ada, ad9, ade)
Digits(14) = Array(ae0, ae5, aec, aed, ae8, ae1, ae6, aef, ae2, ae3, ae4, ae7, aeb, aea, ae9, aee)
Digits(15) = Array(af0, af5, afc, afd, af8, af1, af6, aff, af2, af3, af4, af7, afb, afa, af9, afe)

Digits(16) = Array(b00, b05, b0c, b0d, b08, b01, b06, b0f, b02, b03, b04, b07, b0b, b0a, b09, b0e)
Digits(17) = Array(b10, b15, b1c, b1d, b18, b11, b16, b1f, b12, b13, b14, b17, b1b, b1a, b19, b1e)
Digits(18) = Array(b20, b25, b2c, b2d, b28, b21, b26, b2f, b22, b23, b24, b27, b2b, b2a, b29, b2e)
Digits(19) = Array(b30, b35, b3c, b3d, b38, b31, b36, b3f, b32, b33, b34, b37, b3b, b3a, b39, b3e)
Digits(20) = Array(b40, b45, b4c, b4d, b48, b41, b46, b4f, b42, b43, b44, b47, b4b, b4a, b49, b4e)
Digits(21) = Array(b50, b55, b5c, b5d, b58, b51, b56, b5f, b52, b53, b54, b57, b5b, b5a, b59, b5e)
Digits(22) = Array(b60, b65, b6c, b6d, b68, b61, b66, b6f, b62, b63, b64, b67, b6b, b6a, b69, b6e)
Digits(23) = Array(b70, b75, b7c, b7d, b78, b71, b76, b7f, b72, b73, b74, b77, b7b, b7a, b79, b7e)
Digits(24) = Array(b80, b85, b8c, b8d, b88, b81, b86, b8f, b82, b83, b84, b87, b8b, b8a, b89, b8e)
Digits(25) = Array(b90, b95, b9c, b9d, b98, b91, b96, b9f, b92, b93, b94, b97, b9b, b9a, b99, b9e)
Digits(26) = Array(ba0, ba5, bac, bad, ba8, ba1, ba6, baf, ba2, ba3, ba4, ba7, bab, baa, ba9, bae)
Digits(27) = Array(bb0, bb5, bbc, bbd, bb8, bb1, bb6, bbf, bb2, bb3, bb4, bb7, bbb, bba, bb9, bbe)
Digits(28) = Array(bc0, bc5, bcc, bcd, bc8, bc1, bc6, bcf, bc2, bc3, bc4, bc7, bcb, bca, bc9, bce)
Digits(29) = Array(bd0, bd5, bdc, bdd, bd8, bd1, bd6, bdf, bd2, bd3, bd4, bd7, bdb, bda, bd9, bde)
Digits(30) = Array(be0, be5, bec, bed, be8, be1, be6, bef, be2, be3, be4, be7, beb, bea, be9, bee)
Digits(31) = Array(bf0, bf5, bfc, bfd, bf8, bf1, bf6, bff, bf2, bf3, bf4, bf7, bfb, bfa, bf9, bfe)

Digits(32) = Array(c00, c05, c0c, c0d, c08, c01, c06, c0f, c02, c03, c04, c07, c0b, c0a, c09, c0e)
Digits(33) = Array(c10, c15, c1c, c1d, c18, c11, c16, c1f, c12, c13, c14, c17, c1b, c1a, c19, c1e)
Digits(34) = Array(c20, c25, c2c, c2d, c28, c21, c26, c2f, c22, c23, c24, c27, c2b, c2a, c29, c2e)
Digits(35) = Array(c30, c35, c3c, c3d, c38, c31, c36, c3f, c32, c33, c34, c37, c3b, c3a, c39, c3e)
Digits(36) = Array(c40, c45, c4c, c4d, c48, c41, c46, c4f, c42, c43, c44, c47, c4b, c4a, c49, c4e)
Digits(37) = Array(c50, c55, c5c, c5d, c58, c51, c56, c5f, c52, c53, c54, c57, c5b, c5a, c59, c5e)
Digits(38) = Array(c60, c65, c6c, c6d, c68, c61, c66, c6f, c62, c63, c64, c67, c6b, c6a, c69, c6e)
Digits(39) = Array(c70, c75, c7c, c7d, c78, c71, c76, c7f, c72, c73, c74, c77, c7b, c7a, c79, c7e)

Sub UpdateLeds
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
Const lob = 0     'number of locked balls
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

        ' jps ball speed control
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

'******************************************
' Rules - from Inkochnito instruction cards
'******************************************

Dim Msg(20)
Sub Rules()
    Msg(0) = "HOW TO PLAY DIAMOND LADY" &Chr(10)&Chr(10)
    Msg(1) = ""
    Msg(2) = "SPECIALS: A. JOKER SPECIAL LIT BY COMPLETING ROYAL FLUSH."
    Msg(3) = "B. DIAMOND AND RIGHT OUTLANE LIT BY COMPLETING ALL DIAMONDS. "
    Msg(4) = "C. LIT TOP ROLLOVER AT 10X (MULTIPLIER FEATURE)."
    Msg(5) = ""
    Msg(6) = "EXTRA-BALL: COMPLETING ROYAL FLUSH WHEN LIT FOR 'EXTRA-BALL'."
    Msg(7) = ""
    Msg(8) = "DIAMONDS: COMPLETING ALL DIAMONDS SCORES VALUE, LIGHTS 'CAPTURE',"
    Msg(9) = "COLLECTS DIAMOND BONUS (IF ANY), AND DOUBLES ENTIRE"
    Msg(10) = "SCORE IF ALL SPADE ARE 'UP'."
    Msg(11) = ""
    Msg(12) = "MULTI-BALL: COMPLETING ALL DIAMONDS LIGHTS 'CAPTURE'. PLAYFIELD"
    Msg(13) = "SCORES TIMES 'X'."
    Msg(14) = ""
    Msg(15) = "SPADES: COMPLETING ALL SPADES LIGHTS RAMP TO ADVANCE 'JACKPOT'."
    Msg(16) = ""
    Msg(17) = "JACKPOT: MAKING RAMP WHEN FLASHING ADDS LETTER, LAST LETTER COLLECTS"
    Msg(18) = ""
    Msg(19) = "'SAVE' TARGET: MAKING LOWER LEFT AND RIGHT DIAMONDS RAISE TARGET."
    Msg(20) = ""
    For X = 1 To 20
        Msg(0) = Msg(0) + Msg(X)&Chr(13)
    Next
    MsgBox Msg(0), , "         Instructions and Rule Card"
End Sub

'**********************
'Gottlieb Diamond Lady
'added by Inkochnito
'**********************

Sub editDips
    Dim vpmDips:Set vpmDips = New cvpmDips
    With vpmDips
        .AddForm 700, 400, "Diamond Lady - DIP switches"
        .AddFrame 2, 4, 190, "Maximum credits", 49152, Array("8 credits", 0, "10 credits", 32768, "15 credits", &H00004000, "20 credits", 49152)                                                                                  'dip 15&16
        .AddFrame 2, 80, 190, "Coin chute 1 and 2 control", &H00002000, Array("seperate", 0, "same", &H00002000)                                                                                                                  'dip 14
        .AddFrame 2, 126, 190, "Playfield special", &H00200000, Array("replay", 0, "extra ball", &H00200000)                                                                                                                      'dip 22
        .AddFrame 2, 172, 190, "High games to date control", &H00000020, Array("no effect", 0, "reset high games 2-5 on power off", &H00000020)                                                                                   'dip 6
        .AddFrame 2, 218, 190, "Auto-percentage control", &H00000080, Array("disabled (normal high score mode)", 0, "enabled", &H00000080)                                                                                        'dip 8
        .AddFrame 2, 264, 190, "Royal flush sequence is", &H40000000, Array("reset royal flush value every ball", 0, "memorize royal flush value every ball", &H40000000)                                                         'dip 31
        .AddFrame 2, 310, 190, "Game playing time control", &H80000000, Array("shorter", 0, "longer", &H80000000)                                                                                                                 'dip 32
        .AddFrame 205, 4, 190, "High game to date awards", &H00C00000, Array("not displayed and no award", 0, "displayed and no award", &H00800000, "displayed and 2 replays", &H00400000, "displayed and 3 replays", &H00C00000) 'dip 23&24
        .AddFrame 205, 80, 190, "Balls per game", &H01000000, Array("5 balls", 0, "3 balls", &H01000000)                                                                                                                          'dip 25
        .AddFrame 205, 126, 190, "Replay limit", &H04000000, Array("no limit", 0, "one per game", &H04000000)                                                                                                                     'dip 27
        .AddFrame 205, 172, 190, "Novelty", &H08000000, Array("normal", 0, "extra ball and replay scores 500K", &H08000000)                                                                                                       'dip 28
        .AddFrame 205, 218, 190, "Game mode", &H10000000, Array("replay", 0, "extra ball", &H10000000)                                                                                                                            'dip 29
        .AddFrame 205, 264, 190, "3rd coin chute credits control", &H20000000, Array("no effect", 0, "add 9", &H20000000)                                                                                                         'dip 30
        .AddChk 205, 316, 180, Array("Match feature", &H02000000)                                                                                                                                                                 'dip 26
        .AddChk 205, 331, 190, Array("Attract sound", &H00000040)                                                                                                                                                                 'dip 7
        .AddLabel 50, 360, 300, 20, "After hitting OK, press F3 to reset game with new settings."
        .ViewDips
    End With
End Sub
Set vpmShowDips = GetRef("editDips")


' Thalamus : Exit in a clean and proper way
Sub Table1_exit
  Controller.Pause = False
  Controller.Stop
End Sub

' JP's Nascar Race - Layout and ROM from Stern (uses any rom from Nascar, GP or Dale Jr)
' VP10 version 1.0 by JPSalas
' Used the tables from TAB/destruk as reference for the script (read: copy & paste a lot! :) )
' DOF by arngrim
Option Explicit
Randomize


' Thalamus 2018-07-23
' Added/Updated "Positional Sound Playback Functions" and "Supporting Ball & Sound Functions"
' Thalamus 2018-11-01 : Improved directional sounds
' Thalamus, 2019-02-18 : changed useSolenoids to 2.
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
Const VolGates  = 1    ' Gates volume.
Const VolMetal  = 1    ' Metals volume.
Const VolRH     = 1    ' Rubber hits volume.
Const VolPo     = 1    ' Rubber posts volume.
Const VolPlast  = 1    ' Plastics volume.
Const VolWood   = 1    ' Woods volume.
Const VolKick   = 1    ' Kicker volume.
Const VolSpin   = 1.5  ' Spinners volume.
Const VolFlip   = 1    ' Flipper volume.


On Error Resume Next
ExecuteGlobal GetTextFile("controller.vbs")
If Err Then MsgBox "You need the controller.vbs in order to run this table, available in the vp10 package"
On Error Goto 0

Const BallSize = 50

Dim VarHidden, UseVPMColoredDMD
If Table1.ShowDT = true then
    UseVPMColoredDMD = true
    VarHidden = 1
Ramp16.visible=1
Ramp15.visible=1
Primitive64.visible=1
else
    UseVPMColoredDMD = False
    VarHidden = 0
  TextBox1.Visible = 0
  TextBox5.Visible = 0
Ramp16.visible=0
Ramp15.visible=0
Primitive64.visible=0
end if

LoadVPM "01560000", "SEGA.VBS", 3.26

'********************
'Standard definitions
'********************

Const UseSolenoids = 2
Const UseLamps = 0
Const UseSync = 0
Const UseGI = 0
Const HandleMech = 0

' Standard Sounds
Const SSolenoidOn = "fx_Solenoid"
Const SSolenoidOff = "fx_solenoidoff"
Const SCoin = "fx_Coin"

Dim bsTrough, bsVLock, bsL, bsT, dtR, bsC, bsTE, cbRight

'************
' Table init.
'************

Const cGameName = "gprix" 'Grand Prix


Sub Table1_Init
    vpmInit Me
    With Controller
        .GameName = cGameName
        If Err Then MsgBox "Can't start Game " & cGameName & vbNewLine & Err.Description:Exit Sub
        .SplashInfoLine = "Nascar Race" & vbNewLine & "VP10 table by Zedonius"
        .Games(cGameName).Settings.Value("sound") = 1: 'ensure the sound is on
        .Games(cGameName).Settings.Value("rol") = 0
        .HandleKeyboard = 0
        .ShowTitle = 0
        .ShowDMDOnly = 1
        .ShowFrame = 0
        .HandleMechanics = 0
        .Hidden = VarHidden
        On Error Resume Next
        .Run GetPlayerHWnd
        If Err Then MsgBox Err.Description
        On Error Goto 0
    End With

    ' Nudging
    vpmNudge.TiltSwitch = 56                                                  'plumb tilt
    vpmNudge.Sensitivity = 1
    vpmNudge.TiltObj = Array(sw49, sw50, sw51, LeftSlingShot, RightSlingShot) 'bumpers & slingshots

    ' Trough
    Set bsTrough = New cvpmBallStack
    With bsTrough
        .InitSw 0, 14, 13, 12, 11, 0, 0, 0
        .InitKick BallRelease, 90, 4
        .InitEntrySnd "fx_Solenoid", "fx_Solenoid"
        .InitExitSnd SoundFX("fx_ballrel",DOFContactors), SoundFX("fx_Solenoid",DOFContactors)
        .Balls = 4
    End With

    ' Pit Road Ball Lock
    Set bsVLock = New cvpmVLock
    With bsVLock
      .InitVLock Array(sw32, sw27, sw28), Array(k32, k27, k28), Array(32, 27, 28)
    .ExitDir = 180
    .ExitForce = 0
    .CreateEvents "bsVLock"
      End With

    ' Left-Midway Eject
    Set bsL = New cvpmBallStack
    With bsL
        .InitSaucer sw29, 29, 180, 25
        .InitExitSnd SoundFX("fx_kicker",DOFContactors), SoundFX("fx_Solenoid",DOFContactors)
    End With

    ' Truck
    Set bsT = New cvpmBallStack
    With bsT
        .InitSaucer sw23, 23, 90, 20
        .InitExitSnd SoundFX("fx_kicker",DOFContactors), SoundFX("fx_Solenoid",DOFContactors)
    End With

    ' Right Drop Targets
    Set dtR = New cvpmDropTarget
    With dtR
        .InitDrop Array(sw17, sw18, sw19), Array(17, 18, 19)
        .InitSnd SoundFX("fx_droptarget",DOFContactors), SoundFX("fx_resetdrop",DOFContactors)
    End With

    ' Garage VUK
    Set bsC = New cvpmBallStack
    With bsC
        .InitSw 0, 52, 0, 0, 0, 0, 0, 0
        .InitKick sw52, 185, 26
        .InitExitSnd SoundFX("fx_Popper",DOFContactors), SoundFX("fx_Solenoid",DOFContactors)
    End With

    ' Track Exit Popper - Drain
    Set bsTE = New cvpmBallStack
    With bsTE
        .InitSw 0, 30, 0, 0, 0, 0, 0, 0
    End With

    ' Captive Balls
    Set cbRight = New cvpmCaptiveBall
    With cbRight
        .InitCaptive CapTrigger, CapWall, Array(CapKicker1, CapKicker2), 10
        .NailedBalls = 1
        .ForceTrans = .9
        .MinForce = 3.5
        '.CreateEvents "cbRight" 'the events are done later in the script
        .Start
    End With
    CapKicker1.CreateBall

    ' Main Timer init
    PinMAMETimer.Interval = PinMAMEInterval
    PinMAMETimer.Enabled = 1

    ' Init other dropwalls - animations
    AutoPlunger.PullBack
    OrbitPost.IsDropped = 1
  Controller.Switch(20) = 0
End Sub

'**********
' Keys
'**********

Sub table1_KeyDown(ByVal Keycode)
    If keycode = PlungerKey Then Plunger.Pullback
    If keycode = LeftTiltKey Then vpmNudge.DoNudge 90, 3:PlaySound SoundFX("fx_nudge",0), 0, 1, -0.1, 0.25
    If keycode = RightTiltKey Then vpmNudge.DoNudge 270, 3:PlaySound SoundFX("fx_nudge",0), 0, 1, 0.1, 0.25
    If keycode = CenterTiltKey Then vpmNudge.DoNudge 0, 4:PlaySound SoundFX("fx_nudge",0), 0, 1, 0, 0.25
    If vpmKeyDown(keycode) Then Exit Sub
    if keycode = "3" then
        SetLamp 119, 1
        SetLamp 129, 1
        SetLamp 130, 1
        SetLamp 131, 1
        SetLamp 132, 1
    End if
End Sub

Sub table1_KeyUp(ByVal Keycode)
    If keycode = PlungerKey Then Plunger.Fire:PlaySoundAtVol SoundFX("SolOn",DOFContactors), Plunger, 1
    If vpmKeyUp(keycode) Then Exit Sub
    if keycode = "3" then
        SetLamp 119, 0
        SetLamp 129, 0
        SetLamp 130, 0
        SetLamp 131, 0
        SetLamp 132, 0
    End if
End Sub

'***********
' Solenoids
' the commented solenoids are not in used in this script
'***********

SolCallBack(1) = "SolTrough"
SolCallBack(2) = "vpmSolAutoPlunger AutoPlunger,0,"
SolCallBack(3) = "TruckExit"
SolCallBack(4) = "RotorRotate"      'Truck Motor Drive 'in this table the turn around sign
SolCallBack(5) = "SolGarageDown" 'Garage Release - down animation
SolCallBack(6) = "bsC.SolOut"    ' Garage kicker
SolCallBack(7) = "SolTrackExit"  'Track Exit Popper
SolCallBack(8) = "bsL.SolOut"
'SolCallBack(9)="" 'bumper
'SolCallBack(10)="" 'bumper
'SolCallBack(11)="" 'bumper
SolCallBack(12) = "SolResetDroptargets"
SolCallBack(13) = "SolRightDivert" 'Right Ramp Diverter
SolCallBack(14) = "SolGarageUp"    'Garage Raise - up animation
'SolCallBack(17)="vpmSolSound ""lsling"","      'left slingshot
'SolCallBack(18)="vpmSolSound ""rsling"","      'right slingshot
SolCallBack(19) = "SetLamp 119,"                   'Upper Right Back Panel
SolCallback(20) = "TMag"                           'Upper Accelerator Magnet
SolCallBack(21) = "vpmSolDiverter RDiverter,True," 'Right Track Exit Diverter
SolCallBack(22) = "vpmSolDiverter LDiverter,True," 'Left Track Exit Diverter
SolCallBack(23) = "vpmSolWall OrbitPost, True,"    'Inner Orbit Post
SolCallBack(24) = "vpmSolSound SoundFX(""fx_knocker"",DOFKnocker),"
SolCallback(25) = "LMag"                           'Lower Accelerator Magnet Left
SolCallback(26) = "BMag"                           'Lower Accelerator Magnet Right
SolCallBack(27) = "SolPitRLeft"                    'Pit Lock Release Left
SolCallBack(28) = "SolPitRRight"                   'Pit Lock Release Right
SolCallBack(29) = "FlagRotate"                   'Midway Sign (Hot Dog) In this tabke the lights of the left car
SolCallBack(30) = "SetLamp 130,"                   'Flash Left x3
SolCallBack(31) = "SetLamp 131,"                   'Flash Right x3
SolCallBack(32) = "SetLamp 132,"                   'Flash Test Car x2
'SolCallBack(33)="" 'Left UK Post
'SolCallBack(34)="" 'Center UK Post
'SolCallBack(35)="" 'Right UK Post

Dim Flagpos, FlagDir
FlagPos = 0
FlagDir = 0

Sub FlagRotate(Enabled)
  PlaySoundAtVol SoundFX("fx_Solenoid",DOFContactors), Flag, 1
    If Enabled Then
    Flaganimation.Interval = 8
    FlagDir = -1
    FlagAnimation.Enabled = 1
    End If
End Sub

Sub FlagAnimation_Timer
    Flag.RotZ = FlagPos
    FlagPos = FlagPos + FlagDir
    If FlagPos > 0 Then
        FlagPos = 0
    End If
    If FlagPos < -30 Then
    PlaySoundAtVol SoundFX("fx_Solenoidoff",DOFContactors), Flag, 1
        FlagPos = -35
    End If
  If Flagpos = -30 Then PlaySoundAtVol SoundFX("fx_Solenoidoff",DOFContactors), Flag, 1
  If FlagPos = -30 Then
    FlagDir = +1
    FlagAnimation.Enabled = 1
  End If
End Sub

Dim Rotorpos, RotorDir, Windmile
RotorPos = 0
RotorDir = 0

Sub Rotor_Hit:PlaySoundAtVol SoundFX("fx_target",DOFContactors), Rotor, 1:End Sub

Sub RotorRotate(Enabled)
    If Enabled Then
    Rotoranimation.Interval = 15
    RotorDir = -1
    RotorAnimation.Enabled = 1
  Else
    RotorAnimation.Enabled = 0
    End If
End Sub

Sub RotorAnimation_Timer
    Rotor.RotY = RotorPos
    RotorPos = RotorPos + RotorDir
    If RotorPos > 0 Then
        RotorPos = 0
    End If
    If RotorPos < -360 Then
        RotorPos = 0
    End If
  If RotorPos = -50 Then
    test.IsDropped = 1
    Rotor.collidable = false
  End If
  If RotorPos = -130 Then
    test.IsDropped = 0
    Rotor.collidable = true
  End If
  If RotorPos = -230 Then
    test.IsDropped = 1
    Rotor.collidable = false
  End If
  If RotorPos = -310 Then
    test.IsDropped = 0
    Rotor.collidable = true
  End If
  If RotorPos = -50 Then
    Windmile = 1
  End If
  If RotorPos = -150 Then
    Windmile = 0
  End If
  If RotorPos = -230 Then
    Windmile = 1
  End If
  If RotorPos = -330 Then
    Windmile = 0
  End If
  If Windmile =1 Then
    Controller.Switch(20) = 0
  Else
    Controller.Switch(20) = 1
  End If
End Sub

' Trough
Sub SolTrough(Enabled)
    If Enabled AND bsTrough.Balls > 0 Then
        bsTrough.ExitSol_On
        vpmTimer.PulseSw 15
    End If
End Sub

' Truck ramp kicker
Sub TruckExit(Enabled)
    If Enabled AND bst.Balls > 0 Then
        bsT.ExitSol_On
    End If
End Sub

' Turn around sign
Sub SolTruck(Enabled)
    TurnAround.TimerEnabled = Enabled
End Sub

Sub TurnAround_Timer
    If TurnAround.IsDropped Then
        TurnAround.IsDropped = 0
        TurnAround.TimerInterval = 2000
        Controller.Switch(20) = 1
    Else
        TurnAround.IsDropped = 1
        TurnAround.TimerInterval = 800
        Controller.Switch(20) = 0
    End If
End Sub

Sub TurnAround_Hit:PlaySoundAtVol SoundFX("fx_target",DOFContactors), TurnAround, 1:End Sub

' Garage Release - Down Animation

Dim CarPos, CarDir 'CarDir is both the direction and the speed of the animation
CarPos = 0
CarDir = 0

Sub SolGarageDown(Enabled)
    If Enabled Then
        Controller.Switch(39) = 1
    PlaySoundAtVol "fx_LiftDown", TestCar, 1
        sw36a.IsDropped = 1
        sw36b.IsDropped = 1
        sw40.IsDropped = 0
        CarDir = -1
        GarageAnimation.Enabled = 1
    End If
End Sub

Sub SolGarageUp(Enabled)
    If Enabled Then
        Controller.Switch(39) = 0
    PlaySoundAtVol "fx_LiftUp", TestCar, 1
        sw36a.IsDropped = 0
        sw36b.IsDropped = 0
        sw40.IsDropped = 1
        CarDir = 1
        GarageAnimation.Enabled = 1
    End If
End Sub

Sub GarageAnimation_Timer
    TestCar.TransY = CarPos
  CarSupport.TransZ = CarPos
    Screwcar1.TransZ = CarPos
    Screwcar2.TransZ = CarPos
    t36a.TransY = CarPos
    t36b.TransY = CarPos
    CarPos = CarPos + CarDir
    If CarPos > 0 Then
        CarPos = 0
        Me.Enabled = 0
    End If
    If CarPos < -60 Then
        CarPos = -60
        Me.Enabled = 0
    End If
End Sub

' Track Exit - drain

Sub SolTrackExit(Enabled)
    If Enabled AND bsTE.Balls > 0 Then
        bsTE.ExitSol_On
        bsTrough.AddBall 0
    End If
End Sub

Sub SolResetDropTargets(Enabled)
    If Enabled Then
        dtR.SolDropUp Enabled
        dt17.Z = -16
        dt18.Z = -16
        dt19.Z = -16
    End If
End Sub

' Right Ramp Diverter
Sub SolRightDivert(Enabled)
    If Enabled Then
        diverterRamp.RotateToEnd
    Else
        diverterRamp.RotateToStart
    End If
End Sub

'**********
' Magnets
'**********

Dim M1, M2, M3
M1 = 0:M2 = 0:M3 = 0
Dim MyBall1

Sub TMag(Enabled)
    If Enabled Then
        M1 = 1
        Timer2.Enabled = 0
        Timer2.Enabled = 1
    End If
End Sub

Sub LMag(Enabled)
    If Enabled Then
        M2 = 1
        Timer3.Enabled = 0
        Timer3.Enabled = 1
    End If
End Sub

Sub BMag(Enabled)
    If Enabled Then
        M3 = 1
        Timer4.Enabled = 0
        Timer4.Enabled = 1
    End If
End Sub

Sub Timer2_Timer:M1 = 0:Timer2.Enabled = 0:End Sub
Sub Timer3_Timer:M2 = 0:Timer3.Enabled = 0:End Sub
Sub Timer4_Timer:M3 = 0:Timer4.Enabled = 0:End Sub

Sub TopMagnet_unHit
    If M1 = 1 Then
        Set MyBall1 = ActiveBall
        If MyBall1.VelY < 0 Then
            MyBall1.VelY = MyBall1.VelY * 4.5
            If MyBall1.VelY < -65 Then
                MyBall1.VelY = -65
            Else
                If MyBall1.VelY > -75 Then MyBall1.VelY = -75
            End If
        End If
    End If
End Sub

Sub BottomMagnet_Hit
  PlaySoundAtVol "fx_accelerator", ActiveBall, 1:
End Sub
Sub BottomMagnet_unHit
    If M3 = 1 Then
        Set MyBall1 = ActiveBall
        If MyBall1.VelY > 0 Then
            MyBall1.VelY = MyBall1.VelY * 4.5
            If MyBall1.VelY > 65 Then
                MyBall1.VelY = 65
            Else
                If MyBall1.VelY < 45 Then MyBall1.VelY = 45
                If MyBall1.VelX < 45 Then MyBall1.VelX = 45
            End If
        End If
    End If
End Sub

Sub LeftMagnet_unHit
    If M2 = 1 Then
        Set MyBall1 = ActiveBall
        If MyBall1.VelY > 0 Then
            MyBall1.VelY = MyBall1.VelY * 4.5
            If MyBall1.VelY > 65 Then
                MyBall1.VelY = 65
            Else
                If MyBall1.VelY < 65 Then MyBall1.VelY = 65
            End If
        End If
    End If
End Sub

' Pit Lock Release

Sub SolPitRLeft(Enabled)
    If Enabled Then
        LRL.IsDropped = 1:PlaySoundAtVol SoundFX("fx_Solenoid",DOFContactors), sw27, 1
        LRL.TimerEnabled = 1
       bsVLock.SolExit True
    End If
End Sub

Sub LRL_Timer
    LRL.IsDropped = 0:PlaySoundAtVol SoundFX("fx_Solenoidoff",DOFContactors), sw27, 1
    LRL.TimerEnabled = 0
    bsVLock.SolExit False
End Sub

Sub SolPitRRight(Enabled)
    If Enabled Then
        LRR.IsDropped = 1:PlaySoundAtVol SoundFX("fx_Solenoid",DOFContactors), sw27, 1
        LRR.TimerEnabled = 1
        bsVLock.SolExit True
    End If
End Sub

Sub LRR_Timer
    LRR.IsDropped = 0:PlaySoundAtVol SoundFX("fx_Solenoidoff",DOFContactors), sw27, 1
    LRR.TimerEnabled = 0
    bsVLock.SolExit False
End Sub

' Captive Ball Right - done here to add the ball hit sound.
Sub CapTrigger_Hit:cbRight.TrigHit ActiveBall:End Sub
Sub CapTrigger_UnHit:cbRight.TrigHit 0:End Sub
Sub CapWall_Hit:PlaySoundAtVol "fx_collide", CapKicker2,1:cbRight.BallHit ActiveBall:End Sub
Sub CapKicker2_Hit:cbRight.BallReturn Me:End Sub

'***************
'  Slingshots
'***************
Dim LStep, RStep

Sub LeftSlingShot_Slingshot
    vpmTimer.PulseSw 59
    PlaySoundAtVol SoundFX("fx_slingshot",DOFContactors), lemk, 1
    LeftSling4.Visible = 1
    Lemk.RotX = 26
    LStep = 0
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
    vpmTimer.PulseSw 62
    PlaySoundAtVol SoundFX("fx_slingshot",DOFContactors), remk, 1
    RightSling4.Visible = 1
    Remk.RotX = 26
    RStep = 0
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

'***************
'   Bumpers
'***************

Sub sw49_Hit:vpmTimer.PulseSw 49:PlaySoundAtVol SoundFX("fx_bumper",DOFContactors),sw49, VolBump:End Sub
Sub sw50_Hit:vpmTimer.PulseSw 50:PlaySoundAtVol SoundFX("fx_bumper",DOFContactors),sw50, VolBump:End Sub
Sub sw51_Hit:vpmTimer.PulseSw 51:PlaySoundAtVol SoundFX("fx_bumper",DOFContactors),sw51, VolBump:End Sub

'*********************
' Switches & Rollovers
'*********************

'Sub Sw20_Hit
' Controller.Switch(20) = 0 ' (0 = opto sensor detecting ball)
'End Sub

'Sub Sw20_UnHit
' Controller.Switch(20) = 1 ' (1 = opto sensor not detecting ball)
'End Sub

Sub sw16_Hit:Controller.Switch(16) = 1:PlaySoundAtVol "fx_sensor", ActiveBall, 1:End Sub
Sub sw16_unHit:Controller.Switch(16) = 0:End Sub

Sub sw21_Hit:Controller.Switch(21) = 1:End Sub
Sub sw21_unHit:Controller.Switch(21) = 0:End Sub

Sub sw24_Hit:vpmTimer.PulseSw 24:PlaySoundAtVol "fx_gate", sw24, VolGates:End Sub

Sub sw25_Hit:Controller.Switch(25) = 1:PlaySoundAtVol "fx_sensor", sw25, 1:End Sub
Sub sw25_unHit:Controller.Switch(25) = 0:End Sub

Sub sw26_Hit:Controller.Switch(26) = 1:PlaySoundAtVol "fx_sensor", Activeball, 1:End Sub
Sub sw26_unHit:Controller.Switch(26) = 0:End Sub

Sub sw57_Hit:Controller.Switch(57) = 1:PlaySoundAtVol "fx_sensor", ActiveBall, 1:End Sub
Sub sw57_unHit:Controller.Switch(57) = 0:End Sub

Sub sw58_Hit:Controller.Switch(58) = 1:PlaySoundAtVol "fx_sensor", ActiveBall, 1:End Sub
Sub sw58_unHit:Controller.Switch(58) = 0:End Sub

Sub sw60_Hit:Controller.Switch(60) = 1:PlaySoundAtVol "fx_sensor", ActiveBall, 1:End Sub
Sub sw60_unHit:Controller.Switch(60) = 0:End Sub

Sub sw61_Hit:Controller.Switch(61) = 1:PlaySoundAtVol "fx_sensor", ActiveBall, 1:End Sub
Sub sw61_unHit:Controller.Switch(61) = 0:End Sub

'magnets optos
Sub sw34_Hit:vpmTimer.PulseSw 34:End Sub
Sub sw35_Hit:vpmTimer.PulseSw 35:End Sub
Sub sw37_Hit:vpmTimer.PulseSw 37:End Sub
Sub sw47_Hit:vpmTimer.PulseSw 47:End Sub
Sub sw48_Hit:vpmTimer.PulseSw 48:End Sub

'****************************
' Drain holes, vuks & saucers
'****************************

Sub Drain_Hit:PlaysoundAtVol "fx_drain", drain, 1:bsTE.AddBall Me:End Sub
Sub Drain1_Hit:PlaysoundAtVol "fx_hole-enter", drain1, 1:PlaysoundAtVol "fx_drainmech", drain1, 1:bsTE.AddBall Me:End Sub
Sub Drain2_Hit:PlaysoundAtVol "fx_hole-enter", drain2, 1:PlaysoundAtVol "fx_drainmech", drain2, 1:bsTE.AddBall Me:End Sub
Sub Drain3_Hit:PlaysoundAtVol "fx_hole-enter", drain3, 1:PlaysoundAtVol "fx_drainmech", drain3, 1:bsTE.AddBall Me:End Sub
Sub Drain4_Hit:PlaysoundAtVol "fx_hole-enter", drain4, 1:PlaysoundAtVol "fx_drainmech", drain4, 1:bsTE.AddBall Me:End Sub
Sub Drain5_Hit:PlaysoundAtVol "fx_hole-enter", drain5, 1:PlaysoundAtVol "fx_drainmech", drain5, 1:bsTE.AddBall Me:End Sub
Sub Drain6_Hit:PlaysoundAtVol "fx_hole-enter", drain6, 1:PlaysoundAtVol "fx_drainmech", drain6, 1:bsTE.AddBall Me:End Sub
Sub sw23_Hit:PlaySoundAtVol "fx_kicker_enter", sw23, VolKick:bsT.AddBall 0:End Sub
Sub sw29_Hit:PlaySoundAtVol "fx_kicker_enter", sw29, VolKick:bsL.AddBall 0:End Sub
Sub sw52_Hit:PlaySoundAtVol "fx_hole-enter", sw52,1:bsC.AddBall Me:End Sub
Sub sw52a_Hit:PlaySound "fx_hole-enter":bsC.AddBall Me:End Sub 'TODO
Sub sw52b_Hit:PlaySound "fx_hole-enter":bsC.AddBall Me:End Sub
Sub sw52c_Hit:PlaySound "fx_hole-enter":bsC.AddBall Me:End Sub
Sub sw52d_Hit:PlaySound "fx_hole-enter":bsC.AddBall Me:End Sub
Sub sw52t_Hit:CarShake:End Sub

'***************
'  Targets
'***************

Sub sw36a_Hit
    vpmTimer.PulseSw 36
    PlaySoundAtVol SoundFX("fx_target",DOFContactors), ActiveBall, 1
    PlaySoundAtVol "fx_loosemetalplate", ActiveBall, 1
    CarShake
End Sub

Sub sw36b_Hit
    vpmTimer.PulseSw 36
    PlaySoundAtVol SoundFX("fx_target",DOFContactors), ActiveBall, 1
    PlaySoundAtVol "fx_loosemetalplate", ActiveBall, 1
    CarShake
End Sub

Sub sw38_Hit
    vpmTimer.PulseSw 38
    PlaySoundAtVol SoundFX("fx_target",DOFContactors), ActiveBall, 1
    t38.x = 825
    t38.y = 1119
    Me.TimerEnabled = 1
End Sub

Sub sw38_Timer
    Me.TimerEnabled = 0
    t38.x = 822
    t38.y = 1123
End Sub

Sub sw40_Hit
    vpmTimer.PulseSw 40
    PlaySoundAtVol SoundFX("fx_target",DOFContactors), ActiveBall, 1
    CarShake
End Sub

Sub sw43_Hit
    vpmTimer.PulseSw 43
    PlaySoundAtVol SoundFX("fx_target",DOFContactors), ActiveBall, 1
    t43.x = 365
    t43.y = 711
    Me.TimerEnabled = 1
End Sub

Sub sw43_Timer
    Me.TimerEnabled = 0
    t43.x = 366
    t43.y = 716
End Sub

Sub sw44_Hit
    vpmTimer.PulseSw 44
    PlaySoundAtVol SoundFX("fx_target",DOFContactors), ActiveBall, 1
    t44.x = 231
    t44.y = 748
    Me.TimerEnabled = 1
End Sub

Sub sw44_Timer
    Me.TimerEnabled = 0
    t44.x = 232
    t44.y = 753
End Sub

Sub sw45_Hit
    vpmTimer.PulseSw 45
    PlaySoundAtVol SoundFX("fx_target",DOFContactors), ActiveBall, 1
    t45.x = 194
    t45.y = 1131
    Me.TimerEnabled = 1
End Sub

Sub sw45_Timer
    Me.TimerEnabled = 0
    t45.x = 195
    t45.y = 1136
End Sub

'***************
' Droptargets
'***************

Sub sw17_Hit:dtR.Hit 1:dt17.Z = -70:End Sub
Sub sw18_Hit:dtR.Hit 2:dt18.Z = -70:End Sub
Sub sw19_Hit:dtR.Hit 3:dt19.Z = -70:End Sub

'************
' Spinners
'************

Sub sw33_Spin:vpmTimer.PulseSw 33:PlaySoundAtVol "fx_spinner", sw33, VolSpin:End Sub
Sub sw42_Spin:vpmTimer.PulseSw 42:PlaySoundAtVol "fx_spinner", sw42, VolSpin:End Sub

'************************
' Diverter animation
'************************

Sub SolDiv(Enabled)
    If Enabled Then
        DiverterDir = 1
    Else
        DiverterDir = -1
    End If

    Diverter.Enabled = 0
    If DiverterPos < 1 Then DiverterPos = 1
    If DiverterPos > 5 Then DiverterPos = 5

    Diverter.Enabled = 1
End Sub

Sub Diverter_Timer()
    Select Case DiverterPos
        Case 0:Diverter1.IsDropped = 0:Diverter2.IsDropped = 1:Diverter.Enabled = 0
        Case 1:Diverter2.IsDropped = 0:Diverter1.IsDropped = 1:Diverter3.IsDropped = 1
        Case 2:Diverter3.IsDropped = 0:Diverter2.IsDropped = 1:Diverter4.IsDropped = 1
        Case 3:Diverter4.IsDropped = 0:Diverter3.IsDropped = 1:Diverter5.IsDropped = 1
        Case 4:Diverter5.IsDropped = 0:Diverter4.IsDropped = 1:Diverter6.IsDropped = 1
        Case 5:Diverter6.IsDropped = 0:Diverter5.IsDropped = 1
        Case 6:Diverter.Enabled = 0
    End Select
    DiverterPos = DiverterPos + DiverterDir
End Sub

'********************
'    Flippers
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
    PlaySound "fx_rubber_flipper", 0, parm / 10, -0.1, 0.05
End Sub

Sub RightFlipper_Collide(parm)
    PlaySound "fx_rubber_flipper", 0, parm / 10, 0.1, 0.05
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
  if Gi10.State = 1 then
    gi10a.visible = 1
    gi10b.visible = 1
    gi10c.visible = 1
    gi10d.visible = 1
    gi10e.visible = 1
    gi10f.visible = 1
    gi10g.visible = 1
    gi10h.visible = 1
    gi10i.visible = 1
    gi10j.visible = 1
    gi10k.visible = 1
    else
    gi10a.visible = 0
    gi10b.visible = 0
    gi10c.visible = 0
    gi10d.visible = 0
    gi10e.visible = 0
    gi10f.visible = 0
    gi10g.visible = 0
    gi10h.visible = 0
    gi10i.visible = 0
    gi10j.visible = 0
    gi10k.visible = 0
  end if
  if l40.State = 1 then
    f32.visible = 1
    brightcar1.visible = 1
    brightcar2.visible = 1
    else
    f32.visible = 0
    brightcar1.visible = 0
    brightcar2.visible = 0
  end if
  if l119.State = 1 then
    f19.visible = 1
    f19b.visible = 1
    else
    f19.visible = 0
    f19b.visible = 0
  end if
  if l32.State = 1 then
    BrightTruck.visible = 1
    l32b.visible = 1
    l32c.visible = 1
    else
    BrightTruck.visible = 0
    l32b.visible = 0
    l32c.visible = 0
  end if
  if l60b.State = 1 then
    f30a.state = 1
    f30.visible = 1
    f30h.visible = 1
    f30b.visible = 1
    else
    f30a.state = 0
    f30.visible = 0
    f30h.visible = 0
    f30b.visible = 0
  end if
  if l60a.State = 1 then
    f31a.state = 1
    f31.visible = 1
    f31h.visible = 1
    f31b.visible = 1
    else
    f31a.state = 0
    f31.visible = 0
    f31h.visible = 0
    f31b.visible = 0
  end if
End Sub

Sub UpdateLamps
    NFadeL 1, l1
    NFadeL 2, l2
    NFadeL 3, l3
    NFadeL 4, l4
    NFadeL 5, l5
    NFadeL 6, l6
    NFadeLm 7, l7b
    NFadeL 7, l7
    NFadeLm 8, l8a
    NFadeL 8, l8b
    NFadeL 9, l9
    NFadeL 10, l10
    NFadeL 11, l11
    NFadeL 12, l12
    NFadeL 13, l13
    NFadeL 14, l14
    NFadeLm 15, l15b
    NFadeL 15, l15
    NFadeL 16, l16
    NFadeL 17, l17
    NFadeL 18, l18
    NFadeL 19, l19
    NFadeL 20, l20
    NFadeL 21, l21
    NFadeL 22, l22
    NFadeLm 23, l23b
    NFadeL 23, l23
    NFadeL 24, l24
    NFadeL 25, l25
    NFadeL 26, l26
    NFadeL 27, l27
    NFadeL 28, l28
    NFadeL 29, l29
    NFadeL 30, l30
    NFadeL 31, l31
    NFadeL 32, l32
    NFadeL 33, l33
    NFadeL 34, l34
    NFadeL 35, l35
    NFadeL 36, l36
    NFadeL 37, l37
    NFadeL 38, l38
    NFadeL 39, l39
    NFadeL 40, l40
    NFadeL 41, l41
    NFadeL 42, l42
    NFadeL 43, l43
    NFadeL 44, l44
    NFadeL 45, l45
    NFadeL 46, l46
    NFadeL 47, l47
    NFadeL 48, l48
    NFadeL 49, l49
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
    NFadeLm 60, l60a
    NFadeL 60, l60b
    NFadeL 61, l61
    NFadeL 62, l62
    NFadeL 63, l63
    NFadeL 64, l64
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
    'NFadeL 79, l79 'Tournament Light
    NFadeL 80, l80 'Start Button
    NFadeL 130, l60b
    NFadeL 131, l60a
    NFadeL 132, l40

    'Flashers
    NFadeL 119, l119
End Sub

' div lamp subs

Sub InitLamps()
    Dim x
    For x = 0 to 200
        LampState(x) = 0        ' current light state, independent of the fading level. 0 is off and 1 is on
        FadingLevel(x) = 4      ' used to track the fading state
        FlashSpeedUp(x) = 0.4   ' faster speed when turning on the flasher
        FlashSpeedDown(x) = 0.2 ' slower speed when turning off the flasher
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

Sub aRubbers_Hit(idx):PlaySound "fx_rubber", 0, Vol(ActiveBall)*VolRH, pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0, AudioFade(ActiveBall):End Sub
Sub aPostRubbers_Hit(idx):PlaySound "fx_postrubber", 0, Vol(ActiveBall)*VolPo, pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0, AudioFade(ActiveBall):End Sub
Sub aMetals_Hit(idx):PlaySound "fx_chapa", 0, Vol(ActiveBall)*VolMetal, pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0, AudioFade(ActiveBall):End Sub
Sub aPlastics_Hit(idx):PlaySound "fx_PlasticHit", 0, Vol(ActiveBall)*VolPlast, pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0, AudioFade(ActiveBall):End Sub
Sub Carsupport_Hit(idx):PlaySound "fx_MetalHit", 0, Vol(ActiveBall), pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0, AudioFade(ActiveBall):End Sub
Sub aGates_Hit(idx):PlaySound "fx_Gate", 0, Vol(ActiveBall)*VolGates, pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0, AudioFade(ActiveBall):End Sub
Sub aWoods_Hit(idx):PlaySound "fx_Woodhit", 0, Vol(ActiveBall)*VolWood, pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0, AudioFade(ActiveBall):End Sub

' Ramp Soundss
Sub REnd1_Hit()
    StopSound "fx_metalrolling2"
    PlaySoundAtVol "fx_ballrampdrop", ActiveBall, 1
End Sub

Sub REnd2_Hit()
    StopSound "fx_metalrolling"
    PlaySoundAtVol "fx_ballrampdrop", ActiveBall, 1
End Sub

Sub REnd3_Hit()
    PlaySound "fx_balldrop"
  PlaySound "fx_rollendhit"
End Sub

Sub PlungerEnd_Hit()
    PlaySound "fx_balldrop"
End Sub

Sub LRSound_Hit:PlaySoundAtVol "fx_metalrolling2", ActiveBall, 1:End Sub
Sub RRSound_Hit:PlaySoundAtVol "fx_metalrolling", ActiveBall, 1:End Sub


'******************
'   GI effects
'******************

Dim OldGiState
OldGiState = 2 'start witht he Gi off

Sub GIUpdateTimer_Timer
    Dim tmp, obj
    tmp = Getballs
    If UBound(tmp) <> OldGiState Then
        OldGiState = Ubound(tmp)
        If UBound(tmp) = 2 Then ' since we have 2 captive balls and 1 for the car animation, then Ubound will show 2, so no balls on the table then turn off gi
            For each obj in aGILights
                obj.State = 0
            Next
        Else
            For each obj in aGILights
                obj.State = 1
            Next
        End If
    End If
End Sub

'****************************
'     Test Car Shake
'Inspired by koadic's code
'   from the CV table
' I know it doesn't look
' like Koadic's code but I
' got the idea from his code :)
'****************************

Dim cBall
Const cMod = .65 'percentage of hit power transfered to the car

CarInit

Sub CarShake
    cball.velx = activeball.velx * cMod
    cball.vely = activeball.vely * cMod
    aCarTimer.enabled = True
    bCarTimer.enabled = True
End Sub

Sub CarInit
    Set cBall = hball.createball
    hball.Kick 0, 0
    cball.Mass = 1.6
End Sub

Sub aCarTimer_Timer             'start animation
    Dim x, y
    x = (hball.x - cball.x) / 4 'reduce the X axis movement
    y = (hball.y - cball.y) / 4
    TestCar.transz = x
    TestCar.transx = - y
    t36a.transz = x
    t36a.transx = - y
    t36b.transz = x
    t36b.transx = - y
    CarSupport.transy = - y
    CarSupport.transx = x
  Screwcar1.transy = - y
  Screwcar1.transx = x
  Screwcar2.transy = - y
  Screwcar2.transx = x
End Sub

Sub bCarTimer_Timer 'stop animation
    TestCar.transz = 0
    TestCar.transx = 0
    t36a.transz = 0
    t36a.transx = 0
    t36b.transz = 0
    t36b.transx = 0
    CarSupport.transy = 0
    CarSupport.transx = 0
  Screwcar1.transy = 0
  Screwcar1.transx = 0
  Screwcar2.transy = 0
  Screwcar2.transx = 0
    aCarTimer.enabled = False
    bCarTimer.enabled = False
End Sub

'******************************************************
'           RealTime Updates
'******************************************************

Set MotorCallback = GetRef("GameTimer")

Sub GameTimer
    UpdateMechs
End Sub

Sub UpdateMechs
  LeftBat.RotY=LeftFlipper.currentangle-90
  RightBat.RotY=RightFlipper.currentangle-90
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

Const tnob = 7 ' total number of balls in this table is 4, but always use a higher number here bacuse of the timing
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


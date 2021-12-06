'*****************************************************************************************************
'
'               Gladiators Premier Gattlieb 1993 VPX v1.0.0
'               http://www.ipdb.org/machine.cgi?id=1011
'
'                      Created by Kiwi
'
'*****************************************************************************************************

Option Explicit
Randomize

'*****************************************************************************************************

' Thalamus 2018-11-01 : Improved directional sounds

' !! NOTE :

' Table doesnt work with this script applied
' seems like a timer problem - I haven't got a clue why.

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

'******************************************** OPTIONS **********************************************

Const cGameName   = "gladiatr"

Const VPMorTextDMD = 1  '0 for VPinMAME DMD visible in DT mode, 1 for Internal DMD visible in DT mode

Dim VarHidden, UseVPMColoredDMD
 If Table1.ShowDT = True Then
    UseVPMColoredDMD = VPMorTextDMD
    VarHidden = VPMorTextDMD
Else
    UseVPMColoredDMD = 0
    VarHidden = 0
    TextBoxDMD.Visible = 0
End If

 If B2SOn = True Then VarHidden = 1

'******************************************** Rails Hidden/Visible in FS mode : 0 hidden , 1 visible

Const RailsLights = 0

'******************************************** Ball

'Const BallSize    = 50

'Const BallMass    = 1          'Mass=(50^3)/125000 ,(BallSize^3)/125000

'******************************************** Flashers Level

Const Lumen = 10

'******************************************** Vel Catapult

Const Moltiplicatore = 3

'******************************************** OPTIONS END **********************************************

'******************************************** FSS Init

 If Table1.ShowFSS = True Then
  TextBoxDMD.Visible = False
End If

 If Table1.ShowFSS = False Then
  FlasherDMD.Visible = False
  l112.Visible = 0
  l113.Visible = 0
  l114.Visible = 0
  l115.Visible = 0
  l116.Visible = 0
  l117.Visible = 0

  l184.Visible = 0
  l184a.Visible = 0
  l182.Visible = 0
  l182a.Visible = 0
  l188.Visible = 0
  l188a.Visible = 0
  l189.Visible = 0
  l189a.Visible = 0

  lb1.Visible = 0
  lb2.Visible = 0
  lb3.Visible = 0
  lb4.Visible = 0
  lb5.Visible = 0
  lb6.Visible = 0
  lb7.Visible = 0
  lb8.Visible = 0
  lb9.Visible = 0
  lb10.Visible = 0
  lb11.Visible = 0
  lb12.Visible = 0
  lb13.Visible = 0
  lb14.Visible = 0
  lb15.Visible = 0
  lb16.Visible = 0
  lb17.Visible = 0
  lb18.Visible = 0
  lb19.Visible = 0
  lb20.Visible = 0
  lb21.Visible = 0
  lb22.Visible = 0
  lb23.Visible = 0
  lb24.Visible = 0
  lb25.Visible = 0
  lb26.Visible = 0
  lb27.Visible = 0
  lb28.Visible = 0
  lb29.Visible = 0
  lb30.Visible = 0
  lb31.Visible = 0
  lb32.Visible = 0
  lb33.Visible = 0
  lb34.Visible = 0
  lb35.Visible = 0
  lb36.Visible = 0
  lb37.Visible = 0
  lb38.Visible = 0
  lb39.Visible = 0
  lb40.Visible = 0
  lb41.Visible = 0
  lb42.Visible = 0
  lb43.Visible = 0
  lb44.Visible = 0
End If

'********************************************

LoadVPM "01560000", "gts3.VBS", 3.26

Dim bsTrough, bsTP, bsLP, bsCP, bsRP, dtBank, mRMotor, x, mHole, mHole1, mHole2

Const UseSolenoids = 2
Const UseLamps = 0
Const UseGI = 0
Const UseSync = 0
Const HandleMech = 0

' Standard Sounds
Const SSolenoidOn = "Solenoid"
Const SSolenoidOff = ""
Const SFlipperOn = ""
Const SFlipperOff = ""
Const SCoin = "coin3"


Const swLCoin = 0
Const swRCoin = 1
Const swCCoin = 2
Const swCoinShuttle = 3
Const swStartButton = 4
Const swTournament = 5
Const swFrontDoor = 6

'************
' Table init.
'************

Sub Table1_Init
  vpmInit me
  With Controller
    .GameName = cGameName
    If Err Then MsgBox "Can't start Game" & cGameName & vbNewLine & Err.Description:Exit Sub
    .SplashInfoLine = "Gladiators Premier (Gottlieb 1993)" & vbNewLine & "VPX table by Kiwi 1.0.0"
    .HandleMechanics = 0
    .HandleKeyboard = 0
    .ShowDMDOnly = 1
    .ShowFrame = 0
    .ShowTitle = 0
    .Hidden = VarHidden
'   .Games(cGameName).Settings.Value("dmd_pos_x")=0
'   .Games(cGameName).Settings.Value("dmd_pos_y")=0
'   .Games(cGameName).Settings.Value("dmd_width")=400
'   .Games(cGameName).Settings.Value("dmd_height")=92
'   .Games(cGameName).Settings.Value("rol") = 0
'   .Games(cGameName).Settings.Value("sound") = 1
'   .Games(cGameName).Settings.Value("ddraw") = 1
    .Games(cGameName).Settings.Value("dmd_red")=255
    .Games(cGameName).Settings.Value("dmd_green")=20
    .Games(cGameName).Settings.Value("dmd_blue")=20
         On Error Resume Next
         .Run GetPlayerHWnd
         If Err Then MsgBox Err.Description
         On Error Goto 0
     End With

  ' Nudging
  vpmNudge.TiltSwitch = 151
  vpmNudge.Sensitivity = 2
  vpmNudge.TiltObj = Array(Bumper1, Bumper2, Bumper3, LeftSlingShot, RightSlingShot)

  ' Trough
  Set bsTrough = New cvpmBallStack
  With bsTrough
    .InitSw 23, 0, 33, 0, 0, 0, 0, 0
    .InitKick BallRelease, 80, 4
    .InitEntrySnd "Solenoid", "Solenoid"
    .InitExitSnd SoundFX("popper",DOFContactors), SoundFX("Solenoid",DOFContactors)
    .Balls = 4
  End With

  ' Top Eject Hole
  Set bsTP = New cvpmBallStack
  With bsTP
    .InitSaucer sw21, 21, 150, 20
    .InitExitSnd SoundFX("popper",DOFContactors), SoundFX("Solenoid",DOFContactors)
    .KickAngleVar = 2
    .KickForceVar = 2
    .KickZ = 1.56
  End With

  ' Left Eject Hole
  Set bsLP = New cvpmBallStack
  With bsLP
    .InitSaucer sw31, 31, 105, 14
    .InitExitSnd SoundFX("popper",DOFContactors), SoundFX("Solenoid",DOFContactors)
    .KickAngleVar = 2
    .KickForceVar = 2
    .KickZ = 1
  End With

  ' Center Eject Hole (VUK)
  Set bsCP = New cvpmBallStack
  With bsCP
    .InitSaucer sw22, 22, 0, 38
    .KickZ = 1.56
    .InitExitSnd SoundFX("popper",DOFContactors), SoundFX("Solenoid",DOFContactors)
  End With

  ' Right Eject Hole
  Set bsRP = New cvpmBallStack
  With bsRP
    .InitSaucer sw32, 32, 182, 24
    .InitExitSnd SoundFX("popper",DOFContactors), SoundFX("Solenoid",DOFContactors)
    .KickAngleVar = 0.5
    .KickForceVar = 2
    .KickZ = 1.56
  End With

  ' Drop targets
  set dtBank = new cvpmdroptarget
  With dtBank
    .initdrop array(sw15, sw25, sw35), array(15, 25, 35)
  End With

  ' RMotor mech
    Set mRMotor = New cvpmMech
    With mRMotor
        .MType = vpmMechOneSol + vpmMechReverse + vpmMechNonLinear
        .Sol1 = 23
        .Length = 34 * Moltiplicatore
        .Steps = 136
        .Acc = 10
        .Ret = 1
        .AddSw 30, 0, 0
        .Callback = GetRef("UpdateRMotor")
        .Start
    End With
' Controller.Switch(30) = 1

  ' Init Ramp Motor
  For each x in MotRamp:x.Collidable = 0:Next

  ' Low powered Magnet
    Set mHole = New cvpmMagnet
    With mHole
        .initMagnet MagTrigger, 2
        .X = 321.5
        .Y = 658
        .Size = 35
        .GrabCenter = 0
        .MagnetOn = 1
        .CreateEvents "mHole"
    End With

    Set mHole1 = New cvpmMagnet
    With mHole1
        .initMagnet MagTrigger1, 4
        .X = 747
        .Y = 1166.4
        .Size = 41
        .GrabCenter = 0
        .MagnetOn = 1
        .CreateEvents "mHole1"
    End With

   Set mHole2 = New cvpmMagnet
    With mHole2
        .initMagnet MagTrigger2, 3
        .X = 635.125
        .Y = 914.5
        .Size = 40
        .GrabCenter = 0
        .MagnetOn = 1
        .CreateEvents "mHole2"
    End With

' Main Timer init
  PinMAMETimer.Interval = PinMAMEInterval
  PinMAMETimer.Enabled = 1

' Init Targets Walls
  EjectCP.IsDropped=1:EjectRP.IsDropped=1
  LeftSling.IsDropped=1:LeftSling1.IsDropped=1:LeftSling2.IsDropped=1
  RightSling.IsDropped=1:RightSling1.IsDropped=1:RightSling2.IsDropped=1

' GI Delay Timer
  GIDelay.Enabled = 1

' Init drain
     Drain_Start.CreateBall:Drain_Start.Kick 55, 4

  If ShowDT=True Then
    RailSx.visible=1
    RailDx.visible=1
    Trim.visible=1
    TrimS1.visible=1
    TrimS2.visible=1
    TrimS3.visible=1
    TrimS4.visible=1
    f14.visible=0
    f14a.visible=1
  Else
    RailSx.visible=RailsLights
    RailDx.visible=RailsLights
    Trim.visible=RailsLights
    TrimS1.visible=RailsLights
    TrimS2.visible=RailsLights
    TrimS3.visible=RailsLights
    TrimS4.visible=RailsLights
    f14.visible= ABS(ABS(RailsLights) -1)
    f14a.visible=RailsLights
  End If

' Backbox Flashers Y Position
l112.y = 20
l113.y = 20
l114.y = 20
l115.y = 20
l116.y = 20
l117.y = 20

l184.y = 20
l184a.y = 20
l182.y = 20
l182a.y = 20
l188.y = 20
l188a.y = 20
l189.y = 20
l189a.y = 20

lb1.y = 20
lb2.y = 20
lb3.y = 20
lb4.y = 20
lb5.y = 20
lb6.y = 20
lb7.y = 20
lb8.y = 20
lb9.y = 20
lb10.y = 20
lb11.y = 20
lb12.y = 20
lb13.y = 20
lb14.y = 20
lb15.y = 20
lb16.y = 20
lb17.y = 20
lb18.y = 20
lb19.y = 20
lb20.y = 20
lb21.y = 20
lb22.y = 20
lb23.y = 20
lb24.y = 20
lb25.y = 20
lb26.y = 20
lb27.y = 20
lb28.y = 20
lb29.y = 20
lb30.y = 20
lb31.y = 20
lb32.y = 20
lb33.y = 20
lb34.y = 20
lb35.y = 20
lb36.y = 20
lb37.y = 20
lb38.y = 20
lb39.y = 20
lb40.y = 20
lb41.y = 20
lb42.y = 20
lb43.y = 20
lb44.y = 20

End Sub

Sub Table1_Exit():Controller.Stop:End Sub
Sub Table1_Paused:Controller.Pause = 1:End Sub
Sub Table1_unPaused:Controller.Pause = 0:End Sub

' GI Init
Sub GIDelay_Timer()
  SetLamp 150, 1
  SetLamp 151, 1
  SetLamp 156, 1
  GIDelay.Enabled = 0
  mr.Collidable = True
End Sub

'**********
' Keys
'**********

Dim BumpersOff
Sub Table1_KeyDown(ByVal KeyCode)
  If KeyCode = LeftFlipperKey Then Controller.Switch(82) = 1:BumpersOff=BumpersOff + 1:BumpersEnabled
  If KeyCode = RightFlipperKey Then Controller.Switch(83) = 1:BumpersOff=BumpersOff + 1:BumpersEnabled
  If keycode = PlungerKey Then Plunger.Pullback
  If keycode = LeftTiltKey Then Nudge 90, 4:PlaySound SoundFX("nudge_left",0)
  If keycode = RightTiltKey Then Nudge 270, 4:PlaySound SoundFX("nudge_right",0)
  If keycode = CenterTiltKey Then Nudge 0, 5:PlaySound SoundFX("nudge_forward",0)
  If vpmKeyDown(keycode) Then Exit Sub
    'debug key
    If KeyCode = "3" Then
        SetLamp 190, 1
        SetLamp 191, 1
        SetLamp 194, 1
        SetLamp 195, 1
        SetLamp 196, 1
        SetLamp 197, 1
        SetLamp 199, 1
    End If
End Sub

Sub Table1_KeyUp(ByVal KeyCode)
  If KeyCode = LeftFlipperKey Then Controller.Switch(82) = 0:BumpersOff=BumpersOff - 1:BumpersEnabled
  If KeyCode = RightFlipperKey Then Controller.Switch(83) = 0:BumpersOff=BumpersOff - 1:BumpersEnabled
  If keycode = PlungerKey Then Plunger.Fire:PlaySoundAtVol "plunger2", Plunger, 1
  If vpmKeyUp(keycode) Then Exit Sub
    'debug key
    If KeyCode = "3" Then
        SetLamp 190, 0
        SetLamp 191, 0
        SetLamp 194, 0
        SetLamp 195, 0
        SetLamp 196, 0
        SetLamp 197, 0
        SetLamp 199, 0
    End If
End Sub

Sub BumpersEnabled
 If BumpersOff=2 Then
  Bumper1.HasHitEvent=0
  Bumper2.HasHitEvent=0
  Bumper3.HasHitEvent=0
Else
  Bumper1.HasHitEvent=1
  Bumper2.HasHitEvent=1
  Bumper3.HasHitEvent=1
End If
End Sub

Dim BallinPlunger
Sub sw24_Hit:Controller.Switch(24) = 1:Switch24wire.RotX=-15:BallinPlunger = 1:PlaySoundAtVol "sensor", ActiveBall, 1:End Sub
Sub sw24_UnHit:Controller.Switch(24) = 0:Switch24wire.RotX=0:BallinPlunger = 0:End Sub

' Update RMotor

Dim RampDir

Sub UpdateRMotor(aNewPos, aSpeed, aLastPos)
  RampDir = aNewPos / 4
  RMotor.RotY = RampDir - 4
  MotRamp(aLastPos\4).Collidable = False
  MotRamp(aNewPos\4).Collidable = True
End Sub


'*********
' Switches
'*********

' Slings

Dim LStep, RStep, KickingSxFire, KickingDxFire

Sub LeftSlingshot_Slingshot:vpmTimer.PulseSw 13:TimerKSx.Enabled=0:LeftKicking.RotY = 10:LeftSlingshot.IsDropped=1:LeftSling.IsDropped=0:LStep=0:Me.TimerEnabled=1:PlaySoundAtBallVol SoundFX("bumper5",DOFContactors),1:KickingSxFire = 1:End Sub
Sub LeftSlingshot_Timer
  Select Case LStep
    Case 0:LeftSling.IsDropped = 0:LeftKicking.RotY = 10
    Case 1: 'pause
    Case 2:LeftSling.IsDropped = 1:LeftSling1.IsDropped = 0:LeftKicking.RotY = 7
    Case 3:LeftSling1.IsDropped = 1:LeftSling2.IsDropped = 0:LeftKicking.RotY = 4
    Case 4:LeftSling2.IsDropped = 1:LeftSlingshot.IsDropped=0:Me.TimerEnabled = 0:TimerKSx.Enabled=1:KickingSxFire = 0:LeftKicking.RotY = 0
  End Select
  LStep = LStep + 1
End Sub

Sub TimerKSx_Timer()
 If KickingSxFire = 0 Then
  LeftKicking.RotY = -(-5+(0.16*LeftKickingServo.CurrentAngle))
End If
End Sub

Sub RightSlingshot_Slingshot:vpmTimer.PulseSw 14:TimerKDx.Enabled=0:RightKicking.RotY = 10:RightSlingshot.IsDropped=1:RightSling.IsDropped=0:RStep=0:Me.TimerEnabled=1:PlaySoundAtBallVol SoundFX("bumper5",DOFContactors),1:KickingDxFire = 1:End Sub
Sub RightSlingshot_Timer
  Select Case RStep
    Case 0:RightSling.IsDropped = 0:RightKicking.RotY = 10
    Case 1: 'pause
    Case 2:RightSling.IsDropped = 1:RightSling1.IsDropped = 0:RightKicking.RotY = 7
    Case 3:RightSling1.IsDropped = 1:RightSling2.IsDropped = 0:RightKicking.RotY = 4
    Case 4:RightSling2.IsDropped = 1:RightSlingshot.IsDropped=0:Me.TimerEnabled = 0:TimerKDx.Enabled=1:KickingDxFire = 0:RightKicking.RotY = 0
  End Select
  RStep = RStep + 1
End Sub

Sub TimerKDx_Timer()
 If KickingDxFire = 0 Then
  RightKicking.RotY = -(-5+(0.16*RightKickingServo.CurrentAngle))
End If
End Sub

' Bumpers

Sub Bumper1_Hit:vpmTimer.PulseSw 10:PlaySoundAtVol SoundFX("jet1",DOFContactors), ActiveBall, VolBump:End Sub
Sub Bumper2_Hit:vpmTimer.PulseSw 11:PlaySoundAtVol SoundFX("jet1",DOFContactors), ActiveBall, VolBump:End Sub
Sub Bumper3_Hit:vpmTimer.PulseSw 12:PlaySoundAtVol SoundFX("jet1",DOFContactors), ActiveBall, VolBump:End Sub

' add additional (optional) parameters to PlaySound to increase/decrease the frequency,
' apply all the settings to an already playing sample and choose if to restart this sample from the beginning or not
' PlaySound "name",loopcount,volume,pan,randompitch,pitch,useexisting,restart
' pan ranges from -1.0 (left) over 0.0 (both) to 1.0 (right)
' randompitch ranges from 0.0 (no randomization) to 1.0 (vary between half speed to double speed)
' pitch can be positive or negative and directly adds onto the standard sample frequency
' useexisting is 0 or 1 (if no existing/playing copy of the sound is found, then a new one is created)
' restart is 0 or 1 (only useful if useexisting is 1)

' Spinner
Sub sw20_Spin:vpmTimer.PulseSw 20:PlaySoundAtVol "spinner",sw20,VolSpin:End Sub

' Eject holes
Sub Drain_Hit:bsTrough.AddBall Me:PlaySoundAtVol "drain1a", Drain, 1:End Sub
Sub sw21_Hit:PlaySoundAtVol "fx_kicker_enter1", ActiveBall, VolKick:bsTP.AddBall 0:End Sub
Sub sw31_Hit:PlaySoundAtVol "fx_kicker_enter1", ActiveBall, VolKick:bsLP.AddBall 0:End Sub
Sub sw22_Hit:PlaySoundAtVol "fx_kicker_enter1", ActiveBall, VolKick:bsCP.AddBall 0:End Sub  'VUK
Sub sw32_Hit:PlaySoundAtVol "fx_kicker_enter1", ActiveBall, VolKick:bsRP.AddBall 0:End Sub

' Rollovers
Sub sw80_Hit:Controller.Switch(80) = 1:End Sub
Sub sw80_UnHit:Controller.Switch(80) = 0:End Sub
Sub sw81_Hit:Controller.Switch(81) = 1:End Sub
Sub sw81_UnHit:Controller.Switch(81) = 0:End Sub
Sub sw94_Hit:Controller.Switch(94) = 1:PlaySoundAtVol "sensor", ActiveBall, 1:End Sub
Sub sw94_UnHit:Controller.Switch(94) = 0:End Sub
Sub sw100_Hit:Controller.Switch(100) = 1:End Sub
Sub sw100_UnHit:Controller.Switch(100) = 0:End Sub
Sub sw92_Hit:Controller.Switch(92) = 1:sw92wire.RotX=15:PlaySoundAtVol "sensor", ActiveBall, 1:End Sub
Sub sw92_UnHit:Controller.Switch(92) = 0:sw92wire.RotX=0:End Sub
Sub sw102_Hit:Controller.Switch(102) = 1:PlaySoundAtVol "sensor", ActiveBall, 1:End Sub
Sub sw102_UnHit:Controller.Switch(102) = 0:End Sub
Sub sw93_Hit:Controller.Switch(93) = 1:sw93wire.RotX=15:PlaySoundAtVol "sensor", ActiveBall, 1:End Sub
Sub sw93_UnHit:Controller.Switch(93) = 0:sw93wire.RotX=0:End Sub
Sub sw103_Hit:Controller.Switch(103) = 1:PlaySoundAtVol "sensor", ActiveBall, 1:End Sub
Sub sw103_UnHit:Controller.Switch(103) = 0:End Sub

'Ramp sensors
Sub Gate90_Hit():vpmTimer.PulseSw 90:End Sub
'Sub sw90_Hit():Controller.Switch(90) = 1:End Sub
'Sub sw90_UnHit():Controller.Switch(90) = 0:End Sub
Sub sw91_Hit():Controller.Switch(91) = 1:End Sub
Sub sw91_UnHit():Controller.Switch(91) = 0:End Sub
Sub sw101_Hit():Controller.Switch(101) = 1:PlaySoundAtVol "WireRolling", ActiveBall, 1:End Sub
Sub sw101_UnHit():Controller.Switch(101) = 0:End Sub

' Droptargets
Sub sw15_Hit:PlaySoundAtVol SoundFX("DROPTARG",DOFDropTargets), ActiveBall, 1:End Sub
Sub sw15_Dropped:dtBank.hit 1
 If Bulb8.State=1 Then
  Bulb8T15.State=1
Else
  Bulb8T15.State=0
End If
  Drop15T=1
End Sub

Sub sw25_Hit:PlaySoundAtVol SoundFX("DROPTARG",DOFDropTargets), ActiveBall, 1:End Sub
Sub sw25_Dropped:dtBank.hit 2
 If Bulb8.State=1 Then
  Bulb8T25.State=1
Else
  Bulb8T25.State=0
End If
  Drop25T=1
End Sub

Sub sw35_Hit:PlaySoundAtVol SoundFX("DROPTARG",DOFDropTargets), ActiveBall, 1:End Sub
Sub sw35_Dropped:dtBank.hit 3
 If Bulb8.State=1 Then
  Bulb8T35.State=1
Else
  Bulb8T35.State=0
End If
  Drop35T=1
End Sub

' Targets
Sub sw95_Hit:vpmTimer.PulseSw 95:PlaySoundAtVol SoundFX("target",DOFTargets), ActiveBall, 1:End Sub
Sub sw104_Hit:vpmTimer.PulseSw 104:PlaySoundAtVol SoundFX("target",DOFTargets), ActiveBall, 1:End Sub

' Fx Sounds
Sub CatapultOut_Hit():PlaySoundAtVol "kicker_enter",ActiveBall,1:End Sub
Sub CatapultOut1_Hit():PlaySoundAtVol "kicker_enter",ActiveBall,1:End Sub
Sub Gate_Hit():PlaySoundAtVol "Gate51",ActiveBall, 1:End Sub
Sub Gate2_Hit():PlaySoundAtVol "Gate51",ActiveBall, 1:End Sub
Sub Gate3a_Hit():PlaySoundAtVol "Gate51",ActiveBall, 1:End Sub
Sub GateInSx2_Hit():PlaySoundAtVol "Gate51",ActiveBall, 1:End Sub
Sub GateInDx2_Hit():PlaySoundAtVol "Gate51",ActiveBall, 1:End Sub
Sub Trigger1_Hit():PlaySoundAtVol "kicker_enter",ActiveBall, 1:End Sub
Sub Trigger2_Hit():PlaySoundAtVol "kicker_enter",ActiveBall, 1:End Sub
Sub Perno_Hit():PlaySoundAtVol "WireHit",ActiveBall, 1:End Sub
Sub CatapStop_Hit():PlaySoundAtVol "WireHit",ActiveBall, 1:End Sub
Sub FX1_Hit():PlaySoundAtVol "WireRolling",ActiveBall,1:End Sub
Sub FX2_Hit():PlaySoundAtVol "WireRolling",ActiveBall,1:End Sub
Sub FlapStop_Hit():PlaySoundAtVol "WireHit",ActiveBall,1:End Sub
Sub FlapStop1_Hit():PlaySoundAtVol "WireHit",ActiveBall,1:End Sub
Sub FlapStopTop_Hit():PlaySoundAtVol "WireHit",ActiveBall,1:End Sub

'*********
'Solenoids
'*********

SolCallback(6)  = "SolBallReleaseLP"
SolCallback(7)  = "bsTP.SolOut"
SolCallback(8) = "SetLamp 188,"
SolCallback(9) = "SetLamp 189,"
SolCallback(10) = "dtcbank"
SolCallback(11) = "SolBallReleaseCP"
SolCallback(12) = "SolBallReleaseRP"
SolCallback(13) = "PlungerGate"
SolCallback(14) = "SetLamp 194,"
SolCallback(15) = "SetLamp 195,"
SolCallback(16) = "SetLamp 196,"
SolCallback(17) = "SetLamp 197,"
SolCallback(18) = "SetLamp 198,"
SolCallback(19) = "SetLamp 199,"
SolCallback(20) = "SetLamp 190,"
SolCallback(21) = "SetLamp 191,"
SolCallback(22) = "SetLamp 182,"
SolCallback(23) = "RampMotorRelay"
SolCallback(24) = "SetLamp 184,"
SolCallback(25) = "MShaker"
SolCallback(26) = "GIBackBox"
SolCallback(28) = "bsTrough.SolOut"
SolCallback(29) = "bsTrough.SolIn"
SolCallback(30) = "vpmSolSound SoundFX(""Knocker"",DOFKnocker),"
SolCallback(31) = "GIRelay"
Solcallback(32) = "SolRun"

Sub PlungerGate(Enabled)
 If Enabled Then
  Perno.IsDropped=1
  ServoPerno.Open=1
  PlaySoundAtVol SoundFX("fx_solenoidon",DOFContactors), Primitive98, 1
Else
  Perno.TimerEnabled=1
End If
End Sub

Sub Perno_Timer
  Perno.IsDropped=0
  Perno.TimerEnabled=0
  ServoPerno.Open=0
  PlaySoundAtVol SoundFX("fx_solenoidoff",DOFContactors), Primitive98, 1
End Sub

Sub dtcbank(Enabled)
 If Enabled Then
  dtbank.DropSol_On
  PlaySoundAtVol SoundFX("DTResetB",DOFContactors), Primitive129, VolTarg
End If
  Drop15T=0
  Drop25T=0
  Drop35T=0
  Bulb8T15.TimerEnabled=1
End Sub

Sub Bulb8T15_Timer()
  Bulb8T15.State=0
  Bulb8T25.State=0
  Bulb8T35.State=0
  Bulb8T15.TimerEnabled=0
End Sub

Sub SolBallReleaseRP(Enabled)
 If Enabled Then
  bsRP.ExitSol_On
  EjectRP.IsDropped=0
  sw32Wire.RotX = -15
  EjectRP.TimerEnabled=1
End If
End Sub

Sub EjectRP_Timer
  EjectRP.IsDropped=1
  EjectRP.TimerEnabled=0
End Sub

Sub SolBallReleaseCP(Enabled)
 If Enabled Then
  bsCP.ExitSol_On
  EjectCP.IsDropped=0
  sw22Wire.RotX = -15
  EjectCP.TimerEnabled=1
End If
End Sub

Sub EjectCP_Timer
  EjectCP.IsDropped=1
  EjectCP.TimerEnabled=0
End Sub

Sub SolBallReleaseLP(Enabled)
 If Enabled Then
  bsLP.ExitSol_On
  EjectTimer.Enabled=1
  EjectArm.RotY = -20
End If
End Sub

Sub EjectTimer_Timer
  EjectArm.RotY = EjectArm.RotY +2
 If EjectArm.RotY = 0 Then:EjectTimer.Enabled = 0
End Sub

Sub MShaker(Enabled):PlaySound SoundFX("ShakerSoftPulsing",DOFShaker),0,0.6,0,0:End Sub

Sub RampMotorRelay(Enabled)
 If Enabled Then
  PlaySound SoundFX("MotorNoise",DOFGear),-1,0.4,-0.5,0.1,0,0,1,AudioFade(Primitive99)
Else
  StopSound "MotorNoise"
End If
End Sub

'**************
' GI
'**************

Dim Drop15T, Drop25T, Drop35T
Sub GIRelay(Enabled)
  Dim GIoffon
  GIoffon = ABS(ABS(Enabled) -1)
  SetLamp 150, GIoffon
  SetLamp 151, GIoffon
 If Drop15T=1 Then
  Bulb8T15.State=GIoffon
End If
 If Drop25T=1 Then
  Bulb8T25.State=GIoffon
End If
 If Drop35T=1 Then
  Bulb8T35.State=GIoffon
End If
End Sub

Sub GIBackBox(Enabled)
  Dim GIBBoffon
  GIBBoffon = ABS(ABS(Enabled) -1)
  SetLamp 156, GIBBoffon
End Sub

Sub SolRun(Enabled)
  vpmNudge.SolGameOn Enabled
 If Enabled Then
' PinPlay=1
Else
' PinPlay=0
' Bumper1.HasHitEvent = ABS(Enabled)
' Bumper2.HasHitEvent = ABS(Enabled)
' Bumper3.HasHitEvent = ABS(Enabled)
' LeftSlingshot.Disabled = ABS(Enabled)
' RightSlingShot.Disabled = ABS(Enabled)
End If
End Sub

'**************
' Flipper Subs
'**************

SolCallback(sLRFlipper) = "SolRFlipper"
SolCallback(sLLFlipper) = "SolLFlipper"

Sub SolLFlipper(Enabled)
  If Enabled Then
    PlaySoundAtVol SoundFX("flipperup_left",DOFFlippers), LeftFlipper, VolFlip:LeftFlipper.RotateToEnd:LeftFlipper1.RotateToEnd
    PlaySoundAtVol "flipperup_left", LeftFlipper1, VolFlip
  Else
        PlaySoundAtVol SoundFX("flipperdown_left",DOFFlippers), LeftFlipper, VolFlip:LeftFlipper.RotateToStart:LeftFlipper1.RotateToStart
        PlaySoundAtVol "flipperdown_left", LeftFlipper1, VolFlip
  End If
End Sub

Sub SolRFlipper(Enabled)
  If Enabled Then
    PlaySoundAtVol SoundFX("flipperup_right",DOFFlippers), RightFlipper, VolFlip:RightFlipper.RotateToEnd:RightFlipper1.RotateToEnd
    PlaySoundAtVol "flipperup_right", RightFlipper1, VolFlip
  Else
        PlaySoundAtVol SoundFX("flipperdown_right",DOFFlippers), RightFlipper, VolFlip:RightFlipper.RotateToStart:RightFlipper1.RotateToStart
        PlaySoundAtVol "flipperdown_right", RightFlipper1, VolFlip
  End If
End Sub

Sub LeftFlipper_Collide(parm)
  PlaySoundAtBallVol "rubber_flipper", 1
End Sub

Sub RightFlipper_Collide(parm)
  PlaySoundAtBallVol "rubber_flipper", 1
End Sub

Sub LeftFlipper1_Collide(parm)
  PlaySoundAtBallVol "rubber_flipper", 1
End Sub

Sub RightFlipper1_Collide(parm)
  PlaySoundAtBallVol "rubber_flipper", 1
End Sub

'**********************************
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
  NFadeL 0, l0
  NFadeL 2, l2
' Flash 2, l2a
  NFadeL 3, l3
' Flash 3, l3a
  NFadeL 4, l4
' Flash 4, l4a
  NFadeL 5, l5
' Flash 5, l5a
  NFadeL 6, l6
' Flash 6, l6a
  NFadeL 7, l7
' Flash 7, l7a
  NFadeL 10, l10
  NFadeL 11, l11
  NFadeL 12, l12
' Flash 12, l12a
  NFadeL 13, l13
' Flash 13, l13a
  NFadeL 14, l14
' Flash 14, l14a
  NFadeL 15, l15
' Flash 15, l15a
  NFadeL 16, l16
' Flash 16, l16a
  NFadeL 17, l17
' Flash 17, l17a
  NFadeL 20, l20
  NFadeL 21, l21
  NFadeL 22, l22
  NFadeL 23, l23
  NFadeL 24, l24
  NFadeL 25, l25
  NFadeL 26, l26
  NFadeL 27, l27
  NFadeL 30, l30
  NFadeL 31, l31
  NFadeL 32, l32
  NFadeL 33, l33
  NFadeL 34, l34
  NFadeL 35, l35
  NFadeL 36, l36
  NFadeL 37, l37
  NFadeL 40, l40
  NFadeL 41, l41
  NFadeL 42, l42
  NFadeL 43, l43
  NFadeL 44, l44
  NFadeL 45, l45
  NFadeL 46, l46
  NFadeL 47, l47
  NFadeL 50, l50
  NFadeL 51, l51
  NFadeL 52, l52
  NFadeL 53, l53
  NFadeL 54, l54
  NFadeL 55, l55
  NFadeL 56, l56
  NFadeL 57, l57
  NFadeL 60, l60
  NFadeL 61, l61
  NFadeL 62, l62
  NFadeL 63, l63
  NFadeL 64, l64
  NFadeL 65, l65
  NFadeL 66, l66
  NFadeL 67, l67
  NFadeL 70, l70
  NFadeL 71, l71
  NFadeL 72, l72
  NFadeL 73, l73
  NFadeL 74, l74
  NFadeL 75, l75
  NFadeL 76, l76
  NFadeL 77, l77
  NFadeL 81, l81
  NFadeL 82, l82
  NFadeL 83, l83
  NFadeL 91, l91
  NFadeL 92, l92
  NFadeL 93, l93
  NFadeL 100, l100
  NFadeL 101, l101
  NFadeL 102, l102
  NFadeL 103, l103
  NFadeL 104, l104
  NFadeL 105, l105
  NFadeL 106, l106
  NFadeL 107, l107

'Flashers

  Flash 84, f84
  Flash 85, f85
  Flash 86, f86
  Flash 87, f87
  FastFlash 94, f94
  FastFlash 95, f95
  FastFlash 96, f96
  FastFlash 97, f97

  NFadeLm 150, Bulb1
  NFadeLm 150, Bulb1a
  NFadeLm 150, Bulb2
  NFadeLm 150, Bulb2a
  NFadeLm 150, Bulb3
  NFadeLm 150, Bulb3a
  NFadeLm 150, Bulb4
  NFadeLm 150, Bulb4a
  NFadeLm 150, Bulb5
  NFadeLm 150, Bulb5a
  NFadeLm 150, Bulb6
  NFadeLm 150, Bulb6a
  NFadeLm 150, Bulb7
  NFadeLm 150, Bulb7a
  NFadeLm 150, Bulb8
  NFadeLm 150, Bulb8a
  NFadeLm 150, Bulb9
  NFadeLm 150, Bulb9a
  NFadeLm 150, Bulb10
  NFadeLm 150, Bulb10a
  NFadeLm 150, Bulb11
  NFadeLm 150, Bulb12
  NFadeLm 150, Bulb13
  NFadeLm 150, Bulb13a
  NFadeLm 150, Bulb14
  NFadeLm 150, Bulb15
  NFadeLm 150, Bulb16
  NFadeLm 150, Bulb17
  NFadeLm 150, Bulb17a
  NFadeLm 150, Bulb18
  NFadeLm 150, Bulb18a
  NFadeLm 150, Bulb19
  NFadeLm 150, Bulb19a
  NFadeLm 150, LR1
  NFadeLm 150, LR2
  NFadeLm 150, LR3
  NFadeLm 150, gib1   ' Bumper1
  NFadeLm 150, gib2   ' Bumper2
  NFadeL 150, gib3    ' Bumper3

  Flashm 151, fgiB3   ' Bumper3 Dome
  Flashm 151, fgit1
  Flashm 151, fgit2
  Flashm 151, fgit3
  Flashm 151, fgit4
  Flashm 151, fgit5
  Flashm 151, fgit6
  Flashm 151, fgit7
  Flashm 151, fgit8
  Flashm 151, fgit9
  Flashm 151, fgit10

  Flashm 151, fgit1a
  Flashm 151, fgit2a
  Flashm 151, fgit3a
  Flashm 151, fgit4a
  Flashm 151, fgit8a
  Flashm 151, fgit10a
  Flashm 151, fgit12a
  Flashm 151, fgit13a
  Flashm 151, fgit14a
  Flashm 151, fgit17a
  Flashm 151, fgit18a
  Flash 151, fgit19a

  NFadeLm 194, fl14
  Flashm 194, f14
  Flash 194, f14a
  NFadeLm 195, fl15
  Flash 195, f15
  NFadeLm 196, fl16
  Flash 196, f16
  NFadeLm 197, fl17
  Flash 197, f17
  NFadeLm 198, l197
  NFadeL 198, l197a
  Flash 199, f19
  Flash 190, f20
  Flash 191, f21

' BackBox

  Flash 112, l112
  Flash 113, l113
  Flash 114, l114
  Flash 115, l115
  Flash 116, l116
  Flash 117, l117

  Flashm 182, l182
  Flash 182, l182a
  Flashm 184, l184
  Flash 184, l184a
  Flashm 188, l188
  Flash 188, l188a
  Flashm 189, l189
  Flash 189, l189a

  Flashm 156, lb1
  Flashm 156, lb2
  Flashm 156, lb3
  Flashm 156, lb4
  Flashm 156, lb5
  Flashm 156, lb6
  Flashm 156, lb7
  Flashm 156, lb8
  Flashm 156, lb9
  Flashm 156, lb10
  Flashm 156, lb11
  Flashm 156, lb12
  Flashm 156, lb13
  Flashm 156, lb14
  Flashm 156, lb15
  Flashm 156, lb16
  Flashm 156, lb17
  Flashm 156, lb18
  Flashm 156, lb19
  Flashm 156, lb20
  Flashm 156, lb21
  Flashm 156, lb22
  Flashm 156, lb23
  Flashm 156, lb24
  Flashm 156, lb25
  Flashm 156, lb26
  Flashm 156, lb27
  Flashm 156, lb28
  Flashm 156, lb29
  Flashm 156, lb30
  Flashm 156, lb31
  Flashm 156, lb32
  Flashm 156, lb33
  Flashm 156, lb34
  Flashm 156, lb35
  Flashm 156, lb36
  Flashm 156, lb37
  Flashm 156, lb38
  Flashm 156, lb39
  Flashm 156, lb40
  Flashm 156, lb41
  Flashm 156, lb42
  Flashm 156, lb43
  Flash 156, lb44

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
            Object.IntensityScale = Lumen*FlashLevel(nr)
        Case 5 ' on
            FlashLevel(nr) = FlashLevel(nr) + FlashSpeedUp(nr)
            If FlashLevel(nr) > FlashMax(nr) Then
                FlashLevel(nr) = FlashMax(nr)
                FadingLevel(nr) = 1 'completely on
            End if
            Object.IntensityScale = Lumen*FlashLevel(nr)
    End Select
End Sub

Sub Flashm(nr, object) 'multiple flashers, it just sets the flashlevel
    Object.IntensityScale = Lumen*FlashLevel(nr)
End Sub

Sub FlashMod(nr, object) 'sets the flashlevel from the SolModCallback
    Object.IntensityScale = Lumen*FlashLevel(nr)
End Sub

Sub FastFlash(nr, object)
    Select Case FadingLevel(nr)
    Case 4:object.Visible = 0:FadingLevel(nr) = 0 'off
    Case 5:object.Visible = 1:FadingLevel(nr) = 1 'on
    Object.IntensityScale = Lumen*FlashMax(nr)
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

' *********************************************************************
'           Wall, rubber and metal hit sounds
' *********************************************************************

Sub Rubbers_Hit(idx):PlaySound "rubber1", 0, Vol(ActiveBall), pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0, AudioFade(ActiveBall):End Sub
'Sub aMetals_Hit(idx):PlaySound "metalhit", 0, Vol(ActiveBall), pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0, AudioFade(ActiveBall):End Sub

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

Const tnob = 5 ' total number of balls
Const lob = 0   'number of locked balls
ReDim rolling(tnob)
InitRolling

Sub InitRolling
    Dim i
    For i = 0 to tnob
        rolling(i) = False
    Next
End Sub

Sub RollingUpdate_Timer()
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
    For b = lob to UBound(BOT)
        If BallVel(BOT(b))> 1 Then
            If BOT(b).z <80 Then    'Z Rolling limit
                ballpitch = Pitch(BOT(b))
            Else
                ballpitch = Pitch(BOT(b)) + 15000 'increase the pitch on a ramp or elevated surface
            End If
            rolling(b) = True
            PlaySound("fx_ballrolling" & b), -1, Vol(BOT(b)), Pan(BOT(b)), 0, ballpitch, 1, 0, AudioFade(BOT(b))
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
  PlaySound("fx_collide"), 0, Csng(velocity) ^2 / (VolDiv/VolCol), Pan(ball1), 0, Pitch(ball1), 0, 0, AudioFade(ActiveBall)
End Sub

'******************
' RealTime Updates
'******************
Set MotorCallback = GetRef("GameTimer")

Sub GameTimer
  UpdateGates
' RollingUpdate
  UpdateFlipperLogos
End Sub

Sub UpdateFlipperLogos
  FlipperSx.RotZ = LeftFlipper.CurrentAngle
  FlipperDx.RotZ = RightFlipper.CurrentAngle
  FlipperSx1.RotZ = LeftFlipper1.CurrentAngle
  FlipperDx1.RotZ = RightFlipper1.CurrentAngle
  PernoP1.TransZ = - ServoPerno.CurrentAngle
End Sub

Dim PI
PI = Round(4 * Atn(1), 6) '3.1415926535897932384626433832795

Sub UpdateGates
  GateInSx.RotX = - GateInSx1.CurrentAngle -23
  GateInDx.RotX = - GateInDx1.CurrentAngle -23
  Gate3.RotX = - Gate3a.CurrentAngle -23
  GateDx1.RotX = - (13+(0.9*Gate.CurrentAngle))
' If KickingSxFire = 0 Then
' LeftKicking.RotY = -(-5+(0.15*LeftKickingServo.CurrentAngle))
'End If

' LeftKicking.RotY = -(-8+(0.19*LeftKickingServo.CurrentAngle))
' RightKicking.RotY = -(-8+(0.19*RightKickingServo.CurrentAngle))
  pSpinnerRod.TransX = sin( (sw20.CurrentAngle+180) * (2*PI/360)) * 8
  pSpinnerRod.TransZ = sin( (sw20.CurrentAngle- 90) * (2*PI/360)) * 8
  pSpinnerRod.RotY = sin( (sw20.CurrentAngle-180) * (2*PI/360)) * 6

End Sub

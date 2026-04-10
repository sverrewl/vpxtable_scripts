'*********************************************************************
'1999 South Park by Sega Pinball
'*********************************************************************
'*********************************************************************
'Mechanics by:  Joe Balcer, Rob Hurtado
'Software by: Neil Falconer, Orin Day
'*********************************************************************
'*********************************************************************
'recreated for Visual Pinball by Knorr
'*********************************************************************
'*********************************************************************
'I would like to give my sincere thanks to
'Mfuegemann, Freneticamnesic, Toxie and the VPdevs, Clark Kent,Arngrim
'Gigalula, Kiwi and JP for helping me while building this table
'*********************************************************************

'V1.3
'Added nFozzy Physics to table.  Added wall behind cartman to prevent ball from getting stuck.  Moved left outlane wall to drop ball to try and prevent from going into outlane.
'
'V1.2
'Rebuild KennyKickers
'Rebuild CartmanKickers
'Changed Rolling Sound
'Added RampRollingSounds
'Bug Fixes
'Small Changes With Light

'V1.1
'Fix for "do while" error (thx jp!)

'V1.0
'First Release For VP10.0


' The ultimate South Park Pinball Guide
' https://rileymacdonald.ca/2017/03/17/the-ultimate-south-park-pinball-guide/

Option Explicit
Randomize

Const BallSize = 50
Const BallMass = 1

On Error Resume Next
ExecuteGlobal GetTextFile("controller.vbs")
If Err Then MsgBox "You need the controller.vbs in order to run this table, available in the vp10 package"
On Error Goto 0

LoadVPM "01560000", "sega.VBS", 3.36

'********************
'Standard definitions
'********************

Const cGameName = "sprk_103"
Const UseSolenoids = 2
Const UseLamps = 1
Const SSolenoidOn = "SolOn"
Const SSolenoidOff = "SolOff"
Const SFlipperOn = "FlipperUp"
Const SFlipperOff = "FlipperDown"
Const SCoin = "Coin5"

'///////////////////////-----General Sound Options-----///////////////////////
'// VolumeDial:
'// VolumeDial is the actual global volume multiplier for the mechanical sounds.
'// Values smaller than 1 will decrease mechanical sounds volume.
'// Recommended values should be no greater than 1.
Const VolumeDial = 0.8

'/////////////////////-----BallRoll Sound Amplification -----/////////////////////
Const BallRollAmpFactor = 2 ' 0 = no amplification, 1 = 2.5db amplification, 2 = 5db amplification, 3 = 7.5db amplification, 4 = 9db amplification (aka: Tie Fighter)

'************************************
'Switch SlingShotPlastics to Original
'************************************

SlingPlastics = 1     'Change to 1 for "Yes, i hate Timmy and Butters"

Const RubberizerEnabled = 1     '0 = normal flip rubber, 1 = more lively rubber for flips
Const FlipperCoilRampupMode = 0     '0 = fast, 1 = medium, 2 = slow (tap passes should work)
Const TargetBouncerEnabled = 1      '0 = normal standup targets, 1 = bouncy targets
Const TargetBouncerFactor = 0.7     'Level of bounces. 0.0 - 1.0 are probably usable values.  Never go over 1



Dim bsTrough, PlungerIM, TopVuk, SuperVuk, nBall, aBall, SlingPlastics


Set LampCallback = GetRef("UpdateMultipleLamps")
Set MotorCallback = GetRef("RealTimeUpdates")
Set nBall = ckicker.createball
  ckicker.Kick 0, 0
If SlingPlastics = 1 then
  PlasticSlingshotLinks.Image = "PlasticSlingShotStan"
  PlasticSlingShotRechts.Image = "PlasticSlingShotKenny"
End if

'************
' Table init.
'************

Sub table1_Init
    vpmInit Me
    With Controller
        .GameName = cGameName
        If Err Then MsgBox "Can't start Game " & cGameName & vbNewLine & Err.Description:Exit Sub
        .Games(cGameName).Settings.Value("rol") = 0
        .HandleKeyboard = 0
        .ShowTitle = 0
        .ShowDMDOnly = 1
        .ShowFrame = 0
        .HandleMechanics = 0
        .Hidden = 0
        '.SetDisplayPosition 0, 0, GetPlayerHWnd 'uncomment this line if you don't see the DMD
        On Error Resume Next
        .Run GetPlayerHWnd
        If Err Then MsgBox Err.Description
        On Error Goto 0
'        .Switch(22) = 1 'close coin door
'        .Switch(24) = 1 'and keep it close
        PinMAMETimer.Interval = PinMAMEInterval
        PinMAMETimer.Enabled = true
        vpmNudge.TiltSwitch = swTilt
        vpmNudge.Sensitivity = 2
    End With

    Set bsTrough = New cvpmBallStack
    With bsTrough
        .InitSw 0, 14, 13, 12, 11, 10, 0, 0
        .InitKick BallRelease, 80, 5
        .InitExitSnd SoundFX("BallRelease", DOFContactors), SoundFX("fx_Solenoid", DOFContactors)
        .Balls = 5
        .IsTrough = 1
    End With


  Set TopVuk = New cvpmBallStack
  With TopVuk
    .InitSaucer sw46, 46, -90, 100
    .InitExitSnd SoundFX("SafeHouseKick", DOFContactors), SoundFX("fx_Solenoid", DOFContactors)
    .KickForceVar = 5
    .KickZ = 1.4
        .KickBalls = 1
  End With

    Set SuperVuk = New cvpmBallStack
    With SuperVuk
        .InitSw 0, 45, 0, 0, 0, 0, 0, 0
        .InitKick sw45, -80, 70
        .InitExitSnd SoundFX("KickandWire", DOFContactors), SoundFX("fx_Solenoid", DOFContactors)
        .KickZ = 1.4
    End With

    If table1.ShowDT = False then
        Ramp16.WidthTop = 0
        Ramp16.WidthBottom = 0
        Ramp15.WidthTop = 0
        Ramp15.WidthBottom = 0
        Korpus.Size_Y = 1.7
    Korpus.Z = 60
    End if
End Sub


'******
'Trough
'******
Dim BallCount

Sub SolRelease(Enabled)
    If Enabled Then
'        If bsTrough.Balls = 5 Then vpmTimer.PulseSw 15
        If bsTrough.Balls > 0 Then bsTrough.ExitSol_On
    End If
End Sub

Sub BallRelease_UnHit()
  BallCount = BallCount +1
  End Sub

Sub Drain_Hit
  BallCount = BallCount -1
    PlaySound "Balltruhe", 0, 0.5, 0
    bsTrough.AddBall Me
    If BallCount <1 then GiOff:GiOffState = True
End Sub

'Autoplunger

Sub AutoLaunch(Enabled)
  If Enabled Then
    Plunger.Autoplunger = True
    Plunger.Fire
    PlaySound SoundFX("AutoPlunger", DOFContactors)
    Else
    Plunger.Autoplunger = False
  End If
 End Sub

'********
'ChefVuk
'********

Sub sw46_Hit()
  PlaySound "SafeHouseHit"
  TopVuk.AddBall Me
End Sub

Sub sw46_UnHit()
  GiOn
End Sub

'*********
'CartmanVuk
'*********

Dim cBall, cBall1, cBall2, cBall3, cBall4, cBall5, cBall6, cBall7

Sub CartmanKickerHole_Hit()
    RandomSoundCartman
    Set cBall = ActiveBall:Me.TimerEnabled = 1
    vpmTimer.PulseSw 48
    vpmTimer.AddTimer 1200, "SuperVukAddBall'"
    Me.Enabled = 0
End Sub

Sub CartmanKickerHole1_Hit()
    RandomSoundCartman
    Set cBall1 = ActiveBall:Me.TimerEnabled = 1
    vpmTimer.PulseSw 48
    vpmTimer.AddTimer 1200, "SuperVukAddBall'"
    Me.Enabled = 0
End Sub

Sub CartmanKickerHole2_Hit()
    RandomSoundCartman
    Set cBall2 = ActiveBall:Me.TimerEnabled = 1
    vpmTimer.PulseSw 48
    vpmTimer.AddTimer 1200, "SuperVukAddBall'"
    Me.Enabled = 0
End Sub

Sub CartmanKickerHole3_Hit()
    RandomSoundCartman
    Set cBall3 = ActiveBall:Me.TimerEnabled = 1
    vpmTimer.PulseSw 48
    vpmTimer.AddTimer 1200, "SuperVukAddBall'"
    Me.Enabled = 0
End Sub

Sub CartmanKickerHole4_Hit()
    RandomSoundCartman
    Set cBall4 = ActiveBall:Me.TimerEnabled = 1
    vpmTimer.PulseSw 48
    vpmTimer.AddTimer 1200, "SuperVukAddBall'"
    Me.Enabled = 0
End Sub

Sub CartmanKickerHole5_Hit()
    RandomSoundCartman
    Set cBall5 = ActiveBall:Me.TimerEnabled = 1
    vpmTimer.PulseSw 48
    vpmTimer.AddTimer 1200, "SuperVukAddBall'"
    Me.Enabled = 0
End Sub

Sub CartmanKickerHole6_Hit()
    RandomSoundCartman
    Set cBall6 = ActiveBall:Me.TimerEnabled = 1
    vpmTimer.PulseSw 48
    vpmTimer.AddTimer 1200, "SuperVukAddBall'"
    Me.Enabled = 0
End Sub

Sub CartmanKickerHole7_Hit()
    RandomSoundCartman
    Set cBall7 = ActiveBall:Me.TimerEnabled = 1
    vpmTimer.PulseSw 48
    vpmTimer.AddTimer 1200, "SuperVukAddBall'"
    Me.Enabled = 0
End Sub


Sub SuperVukAddBall()
  SuperVuk.AddBall 1
End Sub


Sub CartmanKickerHole_Timer
    Do While cBall.Z > 0
        cBall.Z = cBall.Z -5
        Exit Sub
    Loop
    Me.DestroyBall
    Me.TimerEnabled = 0
    Me.Enabled = 1
End Sub

Sub CartmanKickerHole1_Timer
    Do While cBall1.Z > 0
        cBall1.Z = cBall1.Z -5
        Exit Sub
    Loop
    Me.DestroyBall
    Me.TimerEnabled = 0
    Me.Enabled = 1
End Sub

Sub CartmanKickerHole2_Timer
    Do While cBall2.Z > 0
        cBall2.Z = cBall2.Z -5
        Exit Sub
    Loop
    Me.DestroyBall
    Me.TimerEnabled = 0
    Me.Enabled = 1
End Sub

Sub CartmanKickerHole3_Timer
    Do While cBall3.Z > 0
        cBall3.Z = cBall3.Z -5
        Exit Sub
    Loop
    Me.DestroyBall
    Me.TimerEnabled = 0
    Me.Enabled = 1
End Sub

Sub CartmanKickerHole4_Timer
    Do While cBall4.Z > 0
        cBall4.Z = cBall4.Z -5
        Exit Sub
    Loop
    Me.DestroyBall
    Me.TimerEnabled = 0
    Me.Enabled = 1
End Sub

Sub CartmanKickerHole5_Timer
    Do While cBall5.Z > 0
        cBall5.Z = cBall5.Z -5
        Exit Sub
    Loop
    Me.DestroyBall
    Me.TimerEnabled = 0
    Me.Enabled = 1
End Sub

Sub CartmanKickerHole6_Timer
    Do While cBall6.Z > 0
        cBall6.Z = cBall6.Z -5
        Exit Sub
    Loop
    Me.DestroyBall
    Me.TimerEnabled = 0
    Me.Enabled = 1
End Sub

Sub CartmanKickerHole7_Timer
    Do While cBall7.Z > 0
        cBall7.Z = cBall7.Z -5
        Exit Sub
    Loop
    Me.DestroyBall
    Me.TimerEnabled = 0
    Me.Enabled = 1
End Sub

Sub sw45_UnHit()
  BallinToilet = 0
End Sub

'*********
'KennyVuk
'*********

Dim KennySolLeft
Dim KennySolRight

Dim kBall, kBall1, kBall2, kBall3, kBall4, kBall5

Sub KennyLeft(Enabled)
  If Enabled then
    UpdateKennyLeft.Enabled = 1
    KennySolLeft = True
    RandomSoundKennyMove
    Else
    UpdateKennyLeft.Enabled = 0
    Kenny.RotZ = 0
    KennyL.RotZ = 0
  End if
End Sub

Sub KennyRight(Enabled)
  If Enabled then
    UpdateKennyRight.Enabled = 1
    KennySolRight = True
    RandomSoundKennyMove
    Else
    UpdateKennyright.Enabled = 0
    Kenny.RotZ = 0
    KennyL.RotZ = 0
  End if
End Sub

Sub KennyDead(Enabled)
  If Enabled then
    KennyFlipper.RotatetoEnd
    PlaySound SoundFX("TrapDoorLow",DOFContactors)
    Else
    KennyFlipper.RotatetoStart
    PlaySound SoundFX("FlipperDown",DOFContactors)
  End if
End Sub

Sub updatekennydead_Timer()
      Kenny.RotX = KennyFlipper.CurrentAngle
End Sub


'Kenny Animation
'RotX -1 = KennyDead
'RotY +1 = rotate clockwise
'RotZ +1 = KennyLeft
'RotZ -1 = KennyRight

Sub UpdateKennyLeft_Timer()
  If KennySolLeft = True And Kenny.RotZ < 8 then Kenny.RotZ = Kenny.RotZ +1
  If KennysolLeft = False And Kenny.RotZ > 0 then Kenny.RotZ = Kenny.RotZ -1
  If Kenny.RotZ >= 8 Then KennySolLeft = False

  If KennySolLeft = True And KennyL.RotZ < 8 then KennyL.RotZ = KennyL.RotZ +1
  If KennysolLeft = False And KennyL.RotZ > 0 then KennyL.RotZ = KennyL.RotZ -1
  If KennyL.RotZ >= 8 Then KennySolLeft = False
End Sub

Sub UpdateKennyRight_Timer()
  If KennySolRight = True And Kenny.RotZ > -8 then Kenny.RotZ = Kenny.RotZ -1
  If KennySolRight = False And Kenny.RotZ < 0 then Kenny.RotZ = Kenny.RotZ +1
  If Kenny.RotZ <= -8 Then KennySolRight = False

  If KennySolRight = True And KennyL.RotZ > -8 then KennyL.RotZ = KennyL.RotZ -1
  If KennySolRight = False And KennyL.RotZ < 0 then KennyL.RotZ = KennyL.RotZ +1
  If KennyL.RotZ <= -8 Then KennySolRight = False
End Sub




Sub KennyKickerHole_Hit()
    RandomSoundKenny
    Set kBall = ActiveBall:Me.TimerEnabled = 1
    vpmTimer.PulseSw 44
    vpmTimer.AddTimer 2000, "SuperVukAddBall'"
    Me.Enabled = 0
End Sub

Sub KennyKickerHole_Timer()
    Do While kball.Z > 0
        kBall.Z = kBall.Z -5
        Exit Sub
    Loop
    Me.DestroyBall
    Me.TimerEnabled = 0
    Me.Enabled = 1
End Sub

Sub KennyKickerHole1_Hit()
    RandomSoundKenny
    Set kBall1 = ActiveBall:Me.TimerEnabled = 1
  vpmTimer.PulseSw 44
    vpmTimer.AddTimer 2000, "SuperVukAddBall'"
    Me.Enabled = 0
End Sub


Sub KennyKickerHole1_Timer()
    Do While kBall1.Z > 0
        kBall1.Z = kBall1.Z -5
        Exit Sub
    Loop
    Me.DestroyBall
    Me.TimerEnabled = 0
    Me.Enabled = 1
End Sub

Sub KennyKickerHole2_Hit()
    RandomSoundKenny
    Set kBall2 = ActiveBall:Me.TimerEnabled = 1
    vpmTimer.PulseSw 44
    vpmTimer.AddTimer 2000, "SuperVukAddBall'"
    Me.Enabled = 0
End Sub

Sub KennyKickerHole2_Timer()
    Do While kBall2.Z > 0
        kBall2.Z = kBall2.Z -5
        Exit Sub
    Loop
    Me.DestroyBall
    Me.TimerEnabled = 0
    Me.Enabled = 1
End Sub

Sub KennyKickerHole3_Hit()
    RandomSoundKenny
    Set kBall3 = ActiveBall:Me.TimerEnabled = 1
    vpmTimer.PulseSw 44
    vpmTimer.AddTimer 2000, "SuperVukAddBall'"
    Me.Enabled = 0
End Sub

Sub KennyKickerHole3_Timer()
    Do While kBall3.Z > 0
        kBall3.Z = kBall3.Z -5
        Exit Sub
    Loop
    Me.DestroyBall
    Me.TimerEnabled = 0
    Me.Enabled = 1
End Sub


Sub KennyKickerHole4_Hit()
    RandomSoundKenny
    Set kBall4 = ActiveBall:Me.TimerEnabled = 1
    vpmTimer.PulseSw 44
    vpmTimer.AddTimer 2000, "SuperVukAddBall'"
    Me.Enabled = 0
End Sub

Sub KennyKickerHole4_Timer()
    Do While kBall4.Z > 0
        kBall4.Z = kBall4.Z -5
        Exit Sub
    Loop
    Me.DestroyBall
    Me.TimerEnabled = 0
    Me.Enabled = 1
End Sub


Sub KennyKickerHole5_Hit()
    RandomSoundKenny
    Set kBall5 = ActiveBall:Me.TimerEnabled = 1
    vpmTimer.PulseSw 44
    vpmTimer.AddTimer 2000, "SuperVukAddBall'"
    Me.Enabled = 0
End Sub

Sub KennyKickerHole5_Timer()
    Do While kBall5.Z > 0
        kBall5.Z = kBall5.Z -5
        Exit Sub
    Loop
    Me.DestroyBall
    Me.TimerEnabled = 0
    Me.Enabled = 1
End Sub


Sub BallTrigger_Hit()
  ActiveBall.VelZ = +15
End Sub


'*********
'Toilet
'*********
Dim FlasherF7Visible


'Seat
Sub ToiletSeatLid(Enabled)
  If Enabled then
    ToiletFlipper.RotatetoEnd
    ToiletKicker.Enabled = True
    PlaySound SoundFX("FlipperDown",DOFContactors)
  Else
    ToiletFlipper.RotatetoStart
    ToiletKicker.Enabled = False
    PlaySound SoundFX("FlipperDown",DOFContactors)
  End if
End Sub


'MrHankeyUp
Sub MrHankeyup(Enabled)
  If Enabled then
  MrHankeyFlipper.RotatetoEnd
  CisternFlipper.RotatetoEnd
  FlasherF7Visible = True
  PlaySound SoundFX("TrapDoorLow",DOFContactors)
  End if
End Sub

'MrHankey
Sub MrHankeyDown(Enabled)
  If Enabled then
  MrHankeyFlipper.RotatetoStart
  CisternFlipper.RotatetoStart
  FlasherF7Visible = False
  End if
End Sub

Dim BallinToilet

Sub ToiletKicker_Hit
  BallinToilet = 1
  sw26spinner.timerenabled = False
    RandomSoundCartman
    Set aBall = ActiveBall:Me.TimerEnabled = 1
    vpmTimer.PulseSw 43
    vpmTimer.AddTimer 2400, "SuperVukAddBall'"
    GiBlinking
  if BallCount < 2 then
    vpmTimer.AddTimer 2400, "GiOff'"
    else
    vpmTimer.AddTimer 2400, "GiOn'"
  End if
End Sub


Sub ToiletKicker_Timer
    Do While aBall.Z > 0
        aBall.Z = aBall.Z -5
        Exit Sub
    Loop
    Me.DestroyBall
    Me.TimerEnabled = 0
End Sub

'''*****************************************************************************************
'''*freneticamnesic level nudge script, based on rascals nudge bobble with help from gtxjoe*
'''*     add timers and "Nudgebobble(keycode)" to left and right tilt keys to activate     *
'''*****************************************************************************************


dim Shaketime, keycode

ShakeTime = 20


Sub LeftHit_Hit()
  nudgebobble2(keycode)
  nudgebobble3(keycode)
  nudgebobble4(keycode)
End Sub

Sub RightHit_Hit()
  nudgebobble2(keycode)
  nudgebobble3(keycode)
  nudgebobble4(keycode)
End Sub



Dim bgcharctr2a:bgcharctr2a = 3
Dim bgcharctr2b:bgcharctr2b = 3
Dim bgcharctr3a:bgcharctr3a = 3
Dim bgcharctr3b:bgcharctr3b = 3
Dim bgcharctr4a:bgcharctr4a = 3
Dim bgcharctr4b:bgcharctr4b = 3

Dim centerlocation2:centerlocation2 = 0
Dim bgdegree2a:bgdegree2a = 5 'move +/- 8 degrees
Dim bgdegree2b:bgdegree2b = 5 'move +/- 8 degrees
Dim bgdurationctr2a:bgdurationctr2a = 0
Dim bgdurationctr2b:bgdurationctr2b = 0
Dim centerlocation3:centerlocation3 = 0
Dim bgdegree3a:bgdegree3a = 5 'move +/- 8 degrees
Dim bgdegree3b:bgdegree3b = 5 'move +/- 8 degrees
Dim bgdurationctr3a:bgdurationctr3a = 0
Dim bgdurationctr3b:bgdurationctr3b = 0
Dim centerlocation4:centerlocation4 = 0
Dim bgdegree4a:bgdegree4a = 5 'move +/- 8 degrees
Dim bgdegree4b:bgdegree4b = 5 'move +/- 8 degrees
Dim bgdurationctr4a:bgdurationctr4a = 0
Dim bgdurationctr4b:bgdurationctr4b = 0

Sub LevelT2_Timer()
  eric.RotAndTra7 = eric.RotAndTra7 + bgcharctr2a  'change rotation value by bgcharctr
  If eric.RotAndTra7 >= bgdegree2a + centerlocation2 then bgcharctr2a = -1:bgdurationctr2a = bgdurationctr2a + 1   'if level moves past max degrees, change direction and increate durationctr
  If eric.RotAndTra7 <= -bgdegree2a + centerlocation2 then bgcharctr2a = 1  'if level moves past min location, change direction
  eric.ObjRotX = eric.ObjRotX + bgcharctr2b  'change rotation value by bgcharctr
  If eric.ObjRotX >= bgdegree2b + centerlocation2 then bgcharctr2b = -1:bgdurationctr2b = bgdurationctr2b + 1   'if level moves past max degrees, change direction and increate durationctr
  If eric.ObjRotX <= -bgdegree2b + centerlocation2 then bgcharctr2b = 1  'if level moves past min location, change direction

  If bgdurationctr2a = ShakeTime then bgdegree2a = bgdegree2a - 2:bgdurationctr2a = 0 'if level has moved back and forth 4 times, decrease amount of movement by -2 and repeat by resetting durationctr
  If bgdurationctr2b = ShakeTime then bgdegree2b = bgdegree2b - 2:bgdurationctr2b = 0 'if level has moved back and forth 4 times, decrease amount of movement by -2 and repeat by resetting durationctr
  If bgdegree2a <= 0 then LevelT2.Enabled = False:bgdegree2a = 5 'if amount of movement is 0, turn off LevelT timer and reset movement back to max 8 degrees
  If bgdegree2b <= 0 then LevelT2.Enabled = False:bgdegree2b = 5 'if amount of movement is 0, turn off LevelT timer and reset movement back to max 8 degrees
End Sub

Sub LevelT3_Timer()
  'Dim loopctr
  Kyle.RotAndTra7 = Kyle.RotAndTra7 + bgcharctr3a  'change rotation value by bgcharctr
  If Kyle.RotAndTra7 >= bgdegree3a + centerlocation3 then bgcharctr3a = -1:bgdurationctr3a = bgdurationctr3a + 1   'if level moves past max degrees, change direction and increate durationctr
  If Kyle.RotAndTra7 <= -bgdegree3a + centerlocation3 then bgcharctr3a = 1  'if level moves past min location, change direction
  Kyle.ObjRotX = Kyle.ObjRotX + bgcharctr3b  'change rotation value by bgcharctr
  If Kyle.ObjRotX >= bgdegree3b + centerlocation3 then bgcharctr3b = -1:bgdurationctr3b = bgdurationctr3b + 1   'if level moves past max degrees, change direction and increate durationctr
  If Kyle.ObjRotX <= -bgdegree3b + centerlocation3 then bgcharctr3b = 1  'if level moves past min location, change direction

  If bgdurationctr3a = ShakeTime then bgdegree3a = bgdegree3a - 2:bgdurationctr3a = 0 'if level has moved back and forth 4 times, decrease amount of movement by -2 and repeat by resetting durationctr
  If bgdurationctr3b = ShakeTime then bgdegree3b = bgdegree3b - 2:bgdurationctr3b = 0 'if level has moved back and forth 4 times, decrease amount of movement by -2 and repeat by resetting durationctr
  If bgdegree3a <= 0 then LevelT3.Enabled = False:bgdegree3a = 5 'if amount of movement is 0, turn off LevelT timer and reset movement back to max 8 degrees
  If bgdegree3b <= 0 then LevelT3.Enabled = False:bgdegree3b = 5 'if amount of movement is 0, turn off LevelT timer and reset movement back to max 8 degrees
End Sub

Sub LevelT4_Timer()
  'Dim loopctr
  Stan.RotAndTra7 = Stan.RotAndTra7 + bgcharctr4a  'change rotation value by bgcharctr
  If Stan.RotAndTra7 >= bgdegree4a + centerlocation4 then bgcharctr4a = -1:bgdurationctr4a = bgdurationctr4a + 1   'if level moves past max degrees, change direction and increate durationctr
  If Stan.RotAndTra7 <= -bgdegree4a + centerlocation4 then bgcharctr4a = 1  'if level moves past min location, change direction
  Stan.ObjRotX = Stan.ObjRotX + bgcharctr4b  'change rotation value by bgcharctr
  If Stan.ObjRotX >= bgdegree4b + centerlocation4 then bgcharctr4b = -1:bgdurationctr4b = bgdurationctr4b + 1   'if level moves past max degrees, change direction and increate durationctr
  If Stan.ObjRotX <= -bgdegree4b + centerlocation4 then bgcharctr4b = 1  'if level moves past min location, change direction

  If bgdurationctr4a = ShakeTime then bgdegree4a = bgdegree4a - 2:bgdurationctr4a = 0 'if level has moved back and forth 4 times, decrease amount of movement by -2 and repeat by resetting durationctr
  If bgdurationctr4b = ShakeTime then bgdegree4b = bgdegree4b - 2:bgdurationctr4b = 0 'if level has moved back and forth 4 times, decrease amount of movement by -2 and repeat by resetting durationctr
  If bgdegree4a <= 0 then LevelT4.Enabled = False:bgdegree4a = 5 'if amount of movement is 0, turn off LevelT timer and reset movement back to max 8 degrees
  If bgdegree4b <= 0 then LevelT4.Enabled = False:bgdegree4b = 5 'if amount of movement is 0, turn off LevelT timer and reset movement back to max 8 degrees
End Sub


Sub Nudgebobble2(keycode)
  LevelT2.Enabled = True:bgdurationctr2a = 0:bgdegree2a = 7:bgdurationctr2b = 0:bgdegree2b = 7
  'LevelT3.Enabled = True:bgdurationctr3 = 0:bgdegree3 = 7
  'LevelT4.Enabled = True:bgdurationctr4 = 0:bgdegree4 = 7
End Sub

Sub Nudgebobble3(keycode)
  'LevelT2.Enabled = True:bgdurationctr2 = 0:bgdegree2 = 7
  LevelT3.Enabled = True:bgdurationctr3a = 0:bgdegree3a = 7:bgdurationctr3b = 0:bgdegree3b = 7
  'LevelT4.Enabled = True:bgdurationctr4 = 0:bgdegree4 = 7
End Sub

Sub Nudgebobble4(keycode)
  'LevelT2.Enabled = True:bgdurationctr2 = 0:bgdegree2 = 7
  'LevelT3.Enabled = True:bgdurationctr3 = 0:bgdegree3 = 7
  LevelT4.Enabled = True:bgdurationctr4a = 0:bgdegree4a = 7:bgdurationctr4b = 0:bgdegree4b = 7
End Sub

'Sub bobblesome_Timer()  'This looks like a free running timer that 1 out of ten times will start movement
' Dim chance
' chance = Int(10*Rnd+1)
' If chance = 5 then Nudgebobble(CenterTiltKey)
'End Sub




'*********
'Solenoids
'*********

SolCallback(1) = "SolRelease"
SolCallback(2) = "AutoLaunch"
SolCallback(3) = "SuperVuk.SolOut"
SolCallback(4) = "TopVuk.SolOut"
SolCallback(5) = "ToiletSeatLid"
SolCallback(6) = "MrHankeyUp"
SolCallback(13) = "MrHankeyDown"
SolCallBack(14) = "KennyDead"
SolCallBack(19) = "KennyLeft"
SolCallback(20) = "KennyRight"
'SolCallback(24) = "vpmSolSound SoundFX(""Knocker"",DOFKnocker),"



'**********
' Keys
'**********

Sub table1_KeyDown(ByVal Keycode)
    If KeyCode = MechanicalTilt Then
        vpmTimer.PulseSw vpmNudge.TiltSwitch
        Exit Sub
    End if

  If keycode = LeftTiltKey Then Nudge 90, 5:SoundNudgeLeft()
  If keycode = RightTiltKey Then Nudge 270, 5:SoundNudgeRight()
  If keycode = CenterTiltKey Then Nudge 0, 3:SoundNudgeCenter()

  If keycode = keyInsertCoin1 or keycode = keyInsertCoin2 or keycode = keyInsertCoin3 or keycode = keyInsertCoin4 Then
    Select Case Int(rnd*3)
      Case 0: PlaySound ("Coin_In_1"), 0, CoinSoundLevel, 0, 0.25
      Case 1: PlaySound ("Coin_In_2"), 0, CoinSoundLevel, 0, 0.25
      Case 2: PlaySound ("Coin_In_3"), 0, CoinSoundLevel, 0, 0.25

    End Select
  End If

  If KeyCode = PlungerKey Then Plunger.Pullback:SoundPlungerPull()

  If keycode = LeftFlipperKey Then FlipperActivate LeftFlipper, LFPress
  If keycode = RightFlipperKey Then FlipperActivate RightFlipper, RFPress

  if keycode=StartGameKey then soundStartButton()

  If vpmKeyDown(keycode) Then Exit Sub
End Sub

Sub table1_KeyUp(ByVal Keycode)
    If keycode = PlungerKey Then Plunger.Fire: SoundPlungerReleaseBall()

  If keycode = LeftFlipperKey Then FlipperDeActivate LeftFlipper, LFPress
  If keycode = RightFlipperKey Then FlipperDeActivate RightFlipper, RFPress

    If vpmKeyUp(keycode) Then Exit Sub
End Sub


'**************
' Flipper Subs
'**************

SolCallback(sLRFlipper) = "SolRFlipper"
SolCallback(sLLFlipper) = "SolLFlipper"

Const ReflipAngle = 20

Sub SolLFlipper(Enabled)
  If Enabled Then
    LF.Fire  'leftflipper.rotatetoend

    If leftflipper.currentangle < leftflipper.endangle + ReflipAngle Then
      RandomSoundReflipUpLeft LeftFlipper
    Else
      SoundFlipperUpAttackLeft LeftFlipper
      RandomSoundFlipperUpLeft LeftFlipper
    End If
  Else
    LeftFlipper.RotateToStart
    If LeftFlipper.currentangle < LeftFlipper.startAngle - 5 Then
      RandomSoundFlipperDownLeft LeftFlipper
    End If
    FlipperLeftHitParm = FlipperUpSoundLevel
    End If
End Sub

Sub SolRFlipper(Enabled)
  If Enabled Then
    RF.Fire 'rightflipper.rotatetoend

    If rightflipper.currentangle > rightflipper.endangle - ReflipAngle Then
      RandomSoundReflipUpRight RightFlipper
    Else
      SoundFlipperUpAttackRight RightFlipper
      RandomSoundFlipperUpRight RightFlipper
    End If
  Else
    RightFlipper.RotateToStart
    If RightFlipper.currentangle > RightFlipper.startAngle + 5 Then
      RandomSoundFlipperDownRight RightFlipper
    End If
    FlipperRightHitParm = FlipperUpSoundLevel
  End If
End Sub

Sub LeftFlipper_Collide(parm)
  CheckLiveCatch Activeball, LeftFlipper, LFCount, parm
    LeftFlipperCollide parm
End Sub

Sub RightFlipper_Collide(parm)
  CheckLiveCatch Activeball, RightFlipper, RFCount, parm
    RightFlipperCollide parm
End Sub

'*********
' Switches
'*********

Sub sw16_Hit:Controller.Switch(16) = 1:sw16wire.RotX = 15:End Sub 'shooterlane'
Sub sw16_UnHit:Controller.Switch(16) = 0:sw16wire.RotX = 0:GiOn:End Sub

Sub sw25_Hit:Controller.Switch(25) = 1:End Sub
Sub sw25_UnHit:Controller.Switch(25) = 0: End Sub

Sub sw26_Hit:Controller.Switch(26) = 1: sw26spinner.timerenabled = True:End Sub
Sub sw26_UnHit:Controller.Switch(26) = 0: End Sub

Sub sw26spinner_Timer()
  If (sw26spinner.currentangle > 9 Or  sw26spinner.currentangle < -9) then
  GiBlinking
  Else
  if Ballcount > 0 then GiOn
' if Ballcount < 1 then GiOff
  End if
End Sub

Sub sw42_Hit: Controller.Switch(42) = 1:ActiveBall.VelZ = 0: End Sub
Sub sw42_UnHit: Controller.Switch(42) = 0: End Sub
'Sub Wall50_Hit: vpmTimer.PulseSw 42: ActiveBall.VelZ = 0:End Sub

Sub sw47_Hit:Controller.Switch(47) = 1:sw47wire.RotX = 15:End Sub 'left orbit'
Sub sw47_UnHit:Controller.Switch(47) = 0:sw47wire.RotX = 0:End Sub

Sub sw57_Hit:Controller.Switch(57) = 1:GiOff:sw57wire.RotX = 15:End Sub 'left outlane'
Sub sw57_UnHit:Controller.Switch(57) = 0:GiOn::sw57wire.RotX = 0:End Sub

Sub sw58_Hit:Controller.Switch(58) = 1:sw58wire.RotX = 15:sw26spinner.timerenabled = False:End Sub 'left inlane'
Sub sw58_UnHit:Controller.Switch(58) = 0:sw58wire.RotX = 0:End Sub

Sub sw60_Hit:Controller.Switch(60) = 1:GiOff:sw60wire.RotX = 15: End Sub 'right outlane'
Sub sw60_UnHit:Controller.Switch(60) = 0:GiOn:sw60wire.RotX = 0:End Sub

Sub sw61_Hit:Controller.Switch(61) = 1:sw61wire.RotX = 15:End Sub 'right inlane'
Sub sw61_UnHit:Controller.Switch(61) = 0:sw61wire.RotX = 0:End Sub

'*********
' Targets
'*********


Sub T17_Hit:vpmTimer.PulseSw 17: DTGrandpa.RotX = DTGrandpa.RotX +5:Me.TimerEnabled = 1:TargetBouncer activeball, 1:End Sub
Sub T17_Timer:DTGrandpa.RotX = DTGrandpa.RotX -5:Me.TimerEnabled = 0:End Sub
Sub T18_Hit:vpmTimer.PulseSw 18: DTMsCartman.RotX = DTMsCartman.RotX +5:Me.TimerEnabled = 1:TargetBouncer activeball, 1:End Sub
Sub T18_Timer:DTMsCartman.RotX = DTMsCartman.RotX -5:Me.TimerEnabled = 0:End Sub
Sub T19_Hit:vpmTimer.PulseSw 19: DTMayor.RotX = DTMayor.RotX +5:Me.TimerEnabled = 1:TargetBouncer activeball, 1:End Sub
Sub T19_Timer:DTMayor.RotX = DTMayor.RotX -5:Me.TimerEnabled = 0:End Sub
Sub T20_Hit:vpmTimer.PulseSw 20: DTBarbrady.RotX = DTBarbrady.RotX +5:Me.TimerEnabled = 1:TargetBouncer activeball, 1:End Sub
Sub T20_Timer:DTBarbrady.RotX = DTBarbrady.RotX -5:Me.TimerEnabled = 0:End Sub
Sub T21_Hit:vpmTimer.PulseSw 21: DTHugo.RotX = DTHugo.RotX +5:Me.TimerEnabled = 1:TargetBouncer activeball, 1:End Sub
Sub T21_Timer:DTHugo.RotX = DTHugo.RotX -5:Me.TimerEnabled = 0:End Sub


Sub T22_Hit:vpmTimer.PulseSw 22: DTJ.RotX = DTJ.RotX +5:Me.TimerEnabled = 1:TargetBouncer activeball, 1:End Sub
Sub T22_Timer:DTJ.RotX = DTJ.RotX -5:Me.TimerEnabled = 0:End Sub
Sub T23_Hit:vpmTimer.PulseSw 23: DTSirJohn.RotX = DTSirJohn.RotX +5:Me.TimerEnabled = 1:TargetBouncer activeball, 1:End Sub
Sub T23_Timer:DTSirJohn.RotX = DTSirJohn.RotX -5:Me.TimerEnabled = 0:End Sub
Sub T24_Hit:vpmTimer.PulseSw 24: DTH.RotX = DTH.RotX +5:Me.TimerEnabled = 1:TargetBouncer activeball, 1:End Sub
Sub T24_Timer:DTH.RotX = DTH.RotX -5:Me.TimerEnabled = 0:End Sub


Sub T34_Hit:vpmTimer.PulseSw 34: DTWendy.RotX = DTWendy.RotX +5:Me.TimerEnabled = 1:TargetBouncer activeball, 1:End Sub
Sub T34_Timer:DTWendy.RotX = DTWendy.RotX -5:Me.TimerEnabled = 0:End Sub
Sub T35_Hit:vpmTimer.PulseSw 35: DTMrGarrison.RotX = DTMrGarrison.RotX +5:Me.TimerEnabled = 1:TargetBouncer activeball, 1:End Sub
Sub T35_Timer:DTMrGarrison.RotX = DTMrGarrison.RotX -5:Me.TimerEnabled = 0:End Sub
Sub T36_Hit:vpmTimer.PulseSw 36: DTMrMackey.RotX = DTMrMackey.RotX +5:Me.TimerEnabled = 1:TargetBouncer activeball, 1:End Sub
Sub T36_Timer:DTMrMackey.RotX = DTMrMackey.RotX -5:Me.TimerEnabled = 0:End Sub


Sub T37_Hit:vpmTimer.PulseSw 37: DTNed.RotX = DTNed.RotX +5:Me.TimerEnabled = 1:TargetBouncer activeball, 1:End Sub
Sub T37_Timer:DTNed.RotX = DTNed.RotX -5:Me.TimerEnabled = 0:End Sub
Sub T38_Hit:vpmTimer.PulseSw 38: DTMephisto.RotX = DTMephisto.RotX +5:Me.TimerEnabled = 1:TargetBouncer activeball, 1:End Sub
Sub T38_Timer:DTMephisto.RotX = DTMephisto.RotX -5:Me.TimerEnabled = 0:End Sub


'*********
' Gates
'*********

Sub GateUpdate()
  WireBumpers.RotX = BumperGate.currentangle +90
  WirePlungerLane.RotX = PlungerLaneGate.currentangle +90
  sw25wire.RotX = sw25spinner.currentangle +90
  sw26wire.RotX = sw26spinner.currentangle +90
End Sub


'*********
' Bumper
'*********

Sub LeftBumper_hit:vpmTimer.pulseSw 49:RandomSoundBumperTop LeftBumper:Me.TimerEnabled = 1:End Sub
Sub LeftBumper_Timer:Me.Timerenabled = 0:End Sub

Sub RightBumper_hit:vpmTimer.pulseSw 50:RandomSoundBumperMiddle RightBumper:Me.TimerEnabled = 1:End Sub
Sub RightBumper_Timer:Me.Timerenabled = 0:End Sub

Sub BottomBumper_hit:vpmTimer.pulseSw 51:RandomSoundBumperBottom BottomBumper:Me.TimerEnabled = 1:End Sub
Sub BottomBumper_Timer:Me.Timerenabled = 0:End Sub



'**********Sling Shot Animations
' Rstep and Lstep  are the variables that increment the animation
'****************


Dim RStep, Lstep

Sub RightSlingShot_Slingshot
    RandomSoundSlingshotRight sling1
    RSling.Visible = 0
    RSling1.Visible = 1
    sling1.TransZ = -20
    RStep = 0
    RightSlingShot.TimerEnabled = 1
    vpmTimer.PulseSw 62
End Sub

Sub RightSlingShot_Timer
    Select Case RStep
        Case 3:RSLing1.Visible = 0:RSLing2.Visible = 1:sling1.TransZ = -10
        Case 4:RSLing2.Visible = 0:RSLing.Visible = 1:sling1.TransZ = 0:RightSlingShot.TimerEnabled = 0
    End Select
    RStep = RStep + 1
End Sub

Sub LeftSlingShot_Slingshot
    RandomSoundSlingshotLeft Sling2
    LSling.Visible = 0
    LSling1.Visible = 1
    sling2.TransZ = -20
    LStep = 0
    LeftSlingShot.TimerEnabled = 1
    vpmTimer.PulseSw 59
End Sub

Sub LeftSlingShot_Timer
    Select Case LStep
        Case 3:LSLing1.Visible = 0:LSLing2.Visible = 1:sling2.TransZ = -10
        Case 4:LSLing2.Visible = 0:LSLing.Visible = 1:sling2.TransZ = 0:LeftSlingShot.TimerEnabled = 0
    End Select
    LStep = LStep + 1
End Sub


'*****************************************
 '  JP's Fading Lamps 3.4 VP9 Fading only
 '      Based on PD's Fading Lights
 ' SetLamp 0 is Off
 ' SetLamp 1 is On
 ' LampState(x) current state
 '****************************************

InitLamps

Sub InitLamps()
    Set Lights(1) = l1
    Set Lights(2) = l2
    Set Lights(3) = l3
    Set Lights(4) = l4
    Set Lights(5) = l5
    Set Lights(6) = l6
    Set Lights(7) = l7
    Set Lights(8) = l8
    Set Lights(9) = l9
    Set Lights(10) = l10
    Set Lights(11) = l11
    Set Lights(12) = l12
    Set Lights(13) = l13
    Set Lights(14) = l14
    Set Lights(15) = l15
    Set Lights(16) = l16
    Set Lights(17) = l17
    Set Lights(18) = l18
    Set Lights(19) = l19
    Set Lights(20) = l20
    Set Lights(21) = l21
    Set Lights(22) = l22
    Set Lights(23) = l23
    Set Lights(24) = l24
    Set Lights(25) = l25
    Set Lights(26) = l26
    Set Lights(27) = l27
    Set Lights(28) = l28
    Set Lights(29) = l29
    Set Lights(30) = l30
    Set Lights(33) = l33
    Set Lights(34) = l34
    Set Lights(35) = l35
    Set Lights(36) = l36
    Set Lights(37) = l37
    Set Lights(38) = l38
    Set Lights(39) = l39
    Set Lights(40) = l40
    Set Lights(41) = l41
    Set Lights(42) = l42
    Set Lights(43) = l43
    Set Lights(44) = l44
    Set Lights(45) = l45
    Set Lights(46) = l46
    Set Lights(48) = l48
    Set Lights(50) = l50
    Set Lights(51) = l51
    Set Lights(52) = l52
    Set Lights(53) = l53
    Set Lights(54) = l54
    Set Lights(55) = l55
    Set Lights(56) = l56
    Set Lights(57) = l57
    Set Lights(58) = l58
    Set Lights(59) = l59
    Set Lights(60) = l60
    Set Lights(61) = l61
    Set Lights(62) = l62
    Set Lights(63) = l63
    Set Lights(64) = l64
End Sub

  Sub UpdateMultipleLamps()
    If l30.state = 1 then Kenny.Visible = False:KennyL.Visible = True: else Kenny.Visible = True: KennyL.Visible = False
    If l22.state = 1 then BulbCoverKennyRed.visible = False:BulbCoverKennyRedL.visible = True:BulbCoverKennyRedF.visible = True: else BulbCoverKennyRed.visible = True: BulbCoverKennyRedL.visible = False:BulbCoverKennyRedF.visible = False
    If l23.state = 1 then BulbCoverKennyYellow.visible = False:BulbCoverKennyYellowL.visible = True:BulbCoverKennyYellowF.visible = True: else BulbCoverKennyYellow.visible = True: BulbCoverKennyYellowF.visible = False: BulbCoverKennyYellowF.visible = False
    If l46.state = 1 then bulbgreenleft.visible = False:bulbgreenleftL.visible = True:bulbgreenleftF.Visible = True: Else bulbgreenleft.visible = True:bulbgreenleftL.visible = False:bulbgreenleftF.visible = False
    If l48.state = 1 then bulbgreenright.visible = False:bulbgreenrightL.visible = True:bulbgreenrightf.visible = True: Else bulbgreenright.visible = True:bulbgreenrightL.visible = False:bulbgreenrightF.visible = False
    If l30.state = 1 then l30a.state = 1:FlasherF6.Visible = True: Else FlasherF6.Visible = False:l30a.state = 0
    If l38.state = 1 then BumperLeftLight:BumperFlasher.Visible = 1: Else BumperLeftLightOff:BumperFlasher.Visible = 0
    If l39.state = 1 then BumperRightLight: Else BumperRightLightOff
    If l40.state = 1 then BumperBottomLight: Else BumperBottomLightOff
End Sub

'RainbowLight

Dim RGBStep, RGBFactor, Red, Green, Blue

RGBStep = 0
RGBFactor = 1
Red = 255
Green = 0
Blue = 0

Sub RGBTimer_timer 'rainbow light color changing
    Select Case RGBStep
        Case 0 'Green
            Green = Green + RGBFactor
            If Green > 255 then
                Green = 255
                RGBStep = 1
            End If
        Case 1 'Red
            Red = Red - RGBFactor
            If Red < 0 then
                Red = 0
                RGBStep = 2
            End If
        Case 2 'Blue
            Blue = Blue + RGBFactor
            If Blue > 255 then
                Blue = 255
                RGBStep = 3
            End If
        Case 3 'Green
            Green = Green - RGBFactor
            If Green < 0 then
                Green = 0
                RGBStep = 4
            End If
        Case 4 'Red
            Red = Red + RGBFactor
            If Red > 255 then
                Red = 255
                RGBStep = 5
            End If
        Case 5 'Blue
            Blue = Blue - RGBFactor
            If Blue < 0 then
                Blue = 0
                RGBStep = 0
            End If
    End Select
    'Light1.color = RGB(Red\10, Green\10, Blue\10)
    l14.colorfull = RGB(Red, Green, Blue)
'    light2.color = RGB(Red, Green, Blue)
'    textbox1.text = Red
'    textbox2.text = Green
'    textbox3.text = Blue
End Sub


'*********************
'Generell Illumination
'*********************

Dim bulb, fbulb, GiOnState, GiOffState

Sub GIOff()
  GiOffState = True
  for each bulb in GI
  bulb.state = 0
  Table1.ColorGradeImage = "-70"
  for each fbulb in GiFlasher
  fbulb.visible = False
  next
  next
End Sub

Sub GIOn()
  GiOnState = True
  for each bulb in GI
  bulb.state = 1
  for each fbulb in GiFlasher
  fbulb.visible = True
  Table1.ColorGradeImage = "ColorGradeLUT256x16_ConSat"
  next
  next
End Sub

Sub GiBlinking()
  GiOnState = False
  GiOffState = False
  FlasherTimer1.Enabled = True
  Table1.ColorGradeImage = "-30"
End Sub


'GiFlasherBlinking

Sub FlasherTimer1_Timer()
  for each fbulb in GiFlasher
  fbulb.visible = 0
  for each bulb in Gi
  bulb.state = 0
  FlasherTimer1.Enabled = 0
  If GiOffState = False then FlasherTimer2.Enabled = 1
  next
  next
End Sub

Sub FlasherTimer2_Timer()
  for each fbulb in GiFlasher
  fbulb.visible = 1
  for each bulb in Gi
  bulb.state = 1
  If GiOnState = False then FlasherTimer1.Enabled = 1: Else
  FlasherTimer2.Enabled = 0
  next
  next
End Sub


'*********************
'BumperRingLights
'*********************

Sub BumperLeftLight()
  for each bulb in RingLightLeftBumper
  bulb.state = 1
  next
End Sub

Sub BumperLeftLightOff()
  for each bulb in RingLightLeftBumper
  bulb.state = 0
  next
End Sub

Sub BumperRightLight()
  for each bulb in RingLightRightBumper
  bulb.state = 1
  next
End Sub

Sub BumperRightLightOff()
  for each bulb in RingLightRightBumper
  bulb.state = 0
  next
End Sub


Sub BumperBottomLight()
  for each bulb in RingLightBottomBumper
  bulb.state = 1
  next
End Sub

Sub BumperBottomLightOff()
  for each bulb in RingLightBottomBumper
  bulb.state = 0
  next
End Sub



'*********
'Flasher
'*********

SolCallback(7) = "FlashPops"
SolCallback(18) = "FlashTopVuk"
SolCallback(25) = "FlashStan"
SolCallback(26) = "FlashChef"
SolCallback(27) = "FlashKenny"
SolCallback(28) = "FlashKyle"
SolCallback(29) = "FlashCartman"
SolCallback(30) = "FlashKennyAndBack"
SolCallback(31) = "FlashMrHankeyToilet"
SolCallback(32) = "FlashSuperVuk"


Sub FlashMrHankeyToilet(Enabled)
  If Enabled Then
  BulbCoverToiletteLeft.Visible = False
  BulbCoverToiletteLeftOn.Visible = True
  BulbCoverToiletteRight.Visible = False
  BulbcoverToiletteRightOn.Visible = True
  F7l.State = 1
  F7r.State = 1
  If FlasherF7Visible = True then FlasherF7.Visible = True
  If FlasherF7Visible = True then FlasherF7a.Visible = True
  Else
  BulbCoverToiletteLeft.Visible = True
  BulbcoverToiletteLeftOn.Visible = False
  BulbCoverToiletteRight.Visible = True
  BulbCoverToiletteRightOn.Visible = False
  F7l.State = 0
  F7r.State = 0
  FlasherF7.Visible = False
  FlasherF7a.Visible = False
  End if
End Sub

Sub FlashStan(Enabled)
  If Enabled Then
  F1.State = 1
  F1a.State = 1
  Else
  F1.State = 0
  F1a.State = 0
  End if
End Sub

Sub FlashChef(Enabled)
  If Enabled Then
  F2.State = 1
  F2a.State = 1
  Else
  F2.State = 0
  F2a.State = 0
  End if
End Sub

Sub FlashKenny(Enabled)
  If Enabled Then
  F3.State = 1
  F3a.State = 1
  Else
  F3.State = 0
  F3a.State = 0
  End if
End Sub

Sub FlashKyle(Enabled)
  If Enabled Then
  F4.State = 1
  F4a.State = 1
  Else
  F4.State = 0
  F4a.State = 0
  End if
End Sub

Sub FlashCartman(Enabled)
  If Enabled Then
  F5.State = 1
  F5a.State = 1
  Else
  F5.State = 0
  F5a.State = 0
  End if
End Sub

Sub FlashSuperVuk(Enabled)
  If Enabled Then
  F8.State = 1
  FlasherCapRed.Image = "dome3_redOn"
  FlasherF8.Visible = True
  if Ballcount > 0 And BallinToilet = 0 then GiBlinking
  Else
  F8.State = 0
  FlasherCapRed.Image = "dome3_red"
  FlasherF8.Visible = False
  If Ballcount > 0 And BallinToilet = 0 then GiOn
  End if
End Sub

Sub FlashPops(Enabled)
  If Enabled Then
  FPa.State = 1
  FPb.State = 1
  FPc.State = 1
  FPd.State = 1
  FlasherCapYellow.Image = "dome3_yellowOn"
  FlasherFPa.Visible = True
  Else
  FPa.State = 0
  FPb.State = 0
  FPc.State = 0
  FPd.State = 0
  FlasherCapYellow.Image = "dome3_yellow"
  FlasherFPa.Visible = False
  End if
End Sub

Sub FlashKennyandBack(Enabled)
  If Enabled Then
  F6b.State = 1
  F6a.Visible = False
  F6a1.Visible = True
  Flasherf6.Visible = True
  FlasherF6a.Visible = True
  Kenny.Visible = False
  KennyL.Visible = True
  Else
  F6b.State = 0
  F6a.Visible = True
  F6a1.Visible = False
  FlasherF6.Visible = False
  FlasherF6a.Visible = False
  Kenny.Visible = True
  KennyL.Visible = False
  End if
End Sub

Sub FlashTopVuk(Enabled)
  If Enabled Then
  F18.State = 1
  FlasherCapChef.Image = "dome3_redOn"
  FlasherF18.Visible = True
  FlasherF18a.Visible = True
  If BallCount > 0 And BallinToilet = 0 then GiBlinking
  Else
  F18.State = 0
  FlasherCapChef.Image = "dome3_red"
  FlasherF18.Visible = False
  FlasherF18a.Visible = False
  If BallCount > 0 And BallinToilet = 0 then GiOn
  End if
End Sub





'******************************************************
'   BALL ROLLING AND DROP SOUNDS
'******************************************************

Const tnob = 9 ' total number of balls
ReDim rolling(tnob)
InitRolling

Dim DropCount
ReDim DropCount(tnob)

Dim ampFactor

Sub InitRolling
  Dim i
  For i = 0 to tnob
    rolling(i) = False
  Next
  Select Case BallRollAmpFactor
    Case 0
      ampFactor = "_amp0"
    Case 1
      ampFactor = "_amp2_5"
    Case 2
      ampFactor = "_amp5"
    Case 3
      ampFactor = "_amp7_5"
    Case 4
      ampFactor = "_amp9"
    Case Else
      ampFactor = "_amp0"
  End Select
End Sub

Sub RollingSoundUpdate()
  Dim BOT, b
  BOT = GetBalls

  ' stop the sound of deleted balls
  For b = UBound(BOT) + 1 to tnob
    rolling(b) = False
    StopSound("BallRoll_" & b & ampFactor)
  Next

  ' exit the sub if no balls on the table
  If UBound(BOT) = -1 Then Exit Sub

  ' play the rolling sound for each ball

  For b = 0 to UBound(BOT)
    If BallVel(BOT(b)) > 1 AND BOT(b).z < 30 Then
      rolling(b) = True
      PlaySound ("BallRoll_" & b & ampFactor), -1, VolPlayfieldRoll(BOT(b)) * 1.1 * VolumeDial, AudioPan(BOT(b)), 0, PitchPlayfieldRoll(BOT(b)), 1, 0, AudioFade(BOT(b))

    Else
      If rolling(b) = True Then
        StopSound("BallRoll_" & b & ampFactor)
        rolling(b) = False
      End If
    End If

    '***Ball Drop Sounds***
    If BOT(b).VelZ < -1 and BOT(b).z < 55 and BOT(b).z > 27 Then 'height adjust for ball drop sounds
      If DropCount(b) >= 5 Then
        DropCount(b) = 0
        If BOT(b).velz > -7 Then
          RandomSoundBallBouncePlayfieldSoft BOT(b)
        Else
          RandomSoundBallBouncePlayfieldHard BOT(b)
        End If
      End If
    End If
    If DropCount(b) < 5 Then
      DropCount(b) = DropCount(b) + 1
    End If
  Next
End Sub


'*******************
' WireRampDropSounds
'*******************

Dim SoundBall

Sub CartmanWireStart_Hit
  GiOn
  WireRampOn False
End Sub

Sub CartmanWireEnd_Hit
  WireRampOff
  PlaySound "BallDrop"
End Sub

Sub Trigger3_Hit()
  'StopSound "plasticrolling"
  WireRampOff
  PlaySound "Balldrop"
End Sub

Sub Trigger4_Hit()
  'StopSound "plasticrolling"
  WireRampOff
  PlaySound "Balldrop"
End Sub

Sub Trigger5_Hit()
  Playsound "Balldrop"
End Sub

Sub Trigger6_Hit()
  'Stopsound "plasticrolling"
  WireRampOff
End Sub

Sub RightRampColl_Hit()
  Set SoundBall = ActiveBall
  WireRampOn True
  'PlaySound "plasticrolling", 0, Vol(ActiveBall) + 0.05, AudioPan(ActiveBall), 0, Pitch(ActiveBall), 1, 0
End Sub

Sub RightRampColl_UnHit()
  WireRampOff
  'StopSound "plasticrolling"
End Sub

Sub LeftRampColl_UnHit()
  WireRampOff
  'StopSound "plasticrolling"
End Sub

Sub LeftRampColl_Hit()
  Set SoundBall = ActiveBall
  'PlaySound "plasticrolling", 0, Vol(ActiveBall) +0.05, AudioPan(ActiveBall), 0, Pitch(ActiveBall), 1, 0
  WireRampOn True
End Sub

Sub StopsoundRamp_Hit()
  'StopSound "plasticrolling"
  WireRampOff
End Sub

Sub StopsoundRamp1_Hit()
  'StopSound "plasticrolling"
  WireRampOff
End Sub


'****************************
'     Realtime Updates
' called by the MotorCallBack
'****************************

Sub RealTimeUpdates
    'flippers
    FlipperLeftP.RotY = LeftFlipper.CurrentAngle
    FlipperRightP.RotY = RightFlipper.CurrentAngle
    ToiletSeat.RotX = ToiletFlipper.CurrentAngle
    Cistern.RotX = CisternFlipper.CurrentAngle
    MrHankey.TransY = MrHankeyFlipper.CurrentAngle
    ' rolling sound
    RollingSoundUpdate
    'Gates
  GateUpdate
End Sub




'******************************
' Diverse Collection Hit Sounds
'******************************


Sub RandomSoundCartman()
    Select Case Int(Rnd * 3) + 1
        Case 1:PlaySound "CartmanHole1", 0, 0.6, AudioPan(ActiveBall), 0, Pitch(ActiveBall), 1, 0
        Case 2:PlaySound "CartmanHole2", 0, 0.6, AudioPan(ActiveBall), 0, Pitch(ActiveBall), 1, 0
        Case 3:PlaySound "CartmanHole4", 0, 0.6, AudioPan(ActiveBall), 0, Pitch(ActiveBall), 1, 0
    End Select
End Sub

Sub RandomSoundKenny()
    Select Case Int(Rnd * 3) + 1
        Case 1:PlaySound "KennyHole1", 0, 0.6, AudioPan(ActiveBall), 0, Pitch(ActiveBall), 1, 0
        Case 2:PlaySound "KennyHole2", 0, 0.6, AudioPan(ActiveBall), 0, Pitch(ActiveBall), 1, 0
        Case 3:PlaySound "KennyHole3", 0, 0.6, AudioPan(ActiveBall), 0, Pitch(ActiveBall), 1, 0
    End Select
End Sub

Sub RandomSoundKennyMove()
    Select Case Int(Rnd * 3) + 1
        Case 1:PlaySound "KennyMove1", 0, 0.2, 0, 0, 0, 1, 0
        Case 2:PlaySound "KennyMove2", 0, 0.2, 0, 0, 0, 1, 0
        Case 3:PlaySound "KennyMove3", 0, 0.2, 0, 0, 0, 1, 0
    End Select
End Sub

'***** Nfozzy Start


dim LF : Set LF = New FlipperPolarity
dim RF : Set RF = New FlipperPolarity

InitPolarity

Sub InitPolarity()
        dim x, a : a = Array(LF, RF)
        for each x in a
                x.AddPoint "Ycoef", 0, RightFlipper.Y-65, 1        'disabled
                x.AddPoint "Ycoef", 1, RightFlipper.Y-11, 1
                x.enabled = True
                x.TimeDelay = 60
        Next

        AddPt "Polarity", 0, 0, 0
        AddPt "Polarity", 1, 0.05, -5.5
        AddPt "Polarity", 2, 0.4, -5.5
        AddPt "Polarity", 3, 0.6, -5.0
        AddPt "Polarity", 4, 0.65, -4.5
        AddPt "Polarity", 5, 0.7, -4.0
        AddPt "Polarity", 6, 0.75, -3.5
        AddPt "Polarity", 7, 0.8, -3.0
        AddPt "Polarity", 8, 0.85, -2.5
        AddPt "Polarity", 9, 0.9,-2.0
        AddPt "Polarity", 10, 0.95, -1.5
        AddPt "Polarity", 11, 1, -1.0
        AddPt "Polarity", 12, 1.05, -0.5
        AddPt "Polarity", 13, 1.1, 0
        AddPt "Polarity", 14, 1.3, 0

        addpt "Velocity", 0, 0,         1
        addpt "Velocity", 1, 0.16, 1.06
        addpt "Velocity", 2, 0.41,         1.05
        addpt "Velocity", 3, 0.53,         1'0.982
        addpt "Velocity", 4, 0.702, 0.968
        addpt "Velocity", 5, 0.95,  0.968
        addpt "Velocity", 6, 1.03,         0.945

        LF.Object = LeftFlipper
        LF.EndPoint = EndPointLp
        RF.Object = RightFlipper
        RF.EndPoint = EndPointRp
End Sub

Sub TriggerLF_Hit() : LF.Addball activeball : End Sub
Sub TriggerLF_UnHit() : LF.PolarityCorrect activeball : End Sub
Sub TriggerRF_Hit() : RF.Addball activeball : End Sub
Sub TriggerRF_UnHit() : RF.PolarityCorrect activeball : End Sub



'******************************************************
'                        FLIPPER CORRECTION FUNCTIONS
'******************************************************

Sub AddPt(aStr, idx, aX, aY)        'debugger wrapper for adjusting flipper script in-game
        dim a : a = Array(LF, RF)
        dim x : for each x in a
                x.addpoint aStr, idx, aX, aY
        Next
End Sub

Class FlipperPolarity
        Public DebugOn, Enabled
        Private FlipAt        'Timer variable (IE 'flip at 723,530ms...)
        Public TimeDelay        'delay before trigger turns off and polarity is disabled TODO set time!
        private Flipper, FlipperStart,FlipperEnd, FlipperEndY, LR, PartialFlipCoef
        Private Balls(20), balldata(20)

        dim PolarityIn, PolarityOut
        dim VelocityIn, VelocityOut
        dim YcoefIn, YcoefOut
        Public Sub Class_Initialize
                redim PolarityIn(0) : redim PolarityOut(0) : redim VelocityIn(0) : redim VelocityOut(0) : redim YcoefIn(0) : redim YcoefOut(0)
                Enabled = True : TimeDelay = 50 : LR = 1:  dim x : for x = 0 to uBound(balls) : balls(x) = Empty : set Balldata(x) = new SpoofBall : next
        End Sub

        Public Property let Object(aInput) : Set Flipper = aInput : StartPoint = Flipper.x : End Property
        Public Property Let StartPoint(aInput) : if IsObject(aInput) then FlipperStart = aInput.x else FlipperStart = aInput : end if : End Property
        Public Property Get StartPoint : StartPoint = FlipperStart : End Property
        Public Property Let EndPoint(aInput) : FlipperEnd = aInput.x: FlipperEndY = aInput.y: End Property
        Public Property Get EndPoint : EndPoint = FlipperEnd : End Property
        Public Property Get EndPointY: EndPointY = FlipperEndY : End Property

        Public Sub AddPoint(aChooseArray, aIDX, aX, aY) 'Index #, X position, (in) y Position (out)
                Select Case aChooseArray
                        case "Polarity" : ShuffleArrays PolarityIn, PolarityOut, 1 : PolarityIn(aIDX) = aX : PolarityOut(aIDX) = aY : ShuffleArrays PolarityIn, PolarityOut, 0
                        Case "Velocity" : ShuffleArrays VelocityIn, VelocityOut, 1 :VelocityIn(aIDX) = aX : VelocityOut(aIDX) = aY : ShuffleArrays VelocityIn, VelocityOut, 0
                        Case "Ycoef" : ShuffleArrays YcoefIn, YcoefOut, 1 :YcoefIn(aIDX) = aX : YcoefOut(aIDX) = aY : ShuffleArrays YcoefIn, YcoefOut, 0
                End Select
                if gametime > 100 then Report aChooseArray
        End Sub

        Public Sub Report(aChooseArray)         'debug, reports all coords in tbPL.text
                if not DebugOn then exit sub
                dim a1, a2 : Select Case aChooseArray
                        case "Polarity" : a1 = PolarityIn : a2 = PolarityOut
                        Case "Velocity" : a1 = VelocityIn : a2 = VelocityOut
                        Case "Ycoef" : a1 = YcoefIn : a2 = YcoefOut
                        case else :tbpl.text = "wrong string" : exit sub
                End Select
                dim str, x : for x = 0 to uBound(a1) : str = str & aChooseArray & " x: " & round(a1(x),4) & ", " & round(a2(x),4) & vbnewline : next
                tbpl.text = str
        End Sub

        Public Sub AddBall(aBall) : dim x : for x = 0 to uBound(balls) : if IsEmpty(balls(x)) then set balls(x) = aBall : exit sub :end if : Next  : End Sub

        Private Sub RemoveBall(aBall)
                dim x : for x = 0 to uBound(balls)
                        if TypeName(balls(x) ) = "IBall" then
                                if aBall.ID = Balls(x).ID Then
                                        balls(x) = Empty
                                        Balldata(x).Reset
                                End If
                        End If
                Next
        End Sub

        Public Sub Fire()
                Flipper.RotateToEnd
                processballs
        End Sub

        Public Property Get Pos 'returns % position a ball. For debug stuff.
                dim x : for x = 0 to uBound(balls)
                        if not IsEmpty(balls(x) ) then
                                pos = pSlope(Balls(x).x, FlipperStart, 0, FlipperEnd, 1)
                        End If
                Next
        End Property

        Public Sub ProcessBalls() 'save data of balls in flipper range
                FlipAt = GameTime
                dim x : for x = 0 to uBound(balls)
                        if not IsEmpty(balls(x) ) then
                                balldata(x).Data = balls(x)
                        End If
                Next
                PartialFlipCoef = ((Flipper.StartAngle - Flipper.CurrentAngle) / (Flipper.StartAngle - Flipper.EndAngle))
                PartialFlipCoef = abs(PartialFlipCoef-1)
        End Sub
        Private Function FlipperOn() : if gameTime < FlipAt+TimeDelay then FlipperOn = True : End If : End Function        'Timer shutoff for polaritycorrect

        Public Sub PolarityCorrect(aBall)
                if FlipperOn() then
                        dim tmp, BallPos, x, IDX, Ycoef : Ycoef = 1

                        'y safety Exit
                        if aBall.VelY > -8 then 'ball going down
                                RemoveBall aBall
                                exit Sub
                        end if

                        'Find balldata. BallPos = % on Flipper
                        for x = 0 to uBound(Balls)
                                if aBall.id = BallData(x).id AND not isempty(BallData(x).id) then
                                        idx = x
                                        BallPos = PSlope(BallData(x).x, FlipperStart, 0, FlipperEnd, 1)
                                        if ballpos > 0.65 then  Ycoef = LinearEnvelope(BallData(x).Y, YcoefIn, YcoefOut)                                'find safety coefficient 'ycoef' data
                                end if
                        Next

                        If BallPos = 0 Then 'no ball data meaning the ball is entering and exiting pretty close to the same position, use current values.
                                BallPos = PSlope(aBall.x, FlipperStart, 0, FlipperEnd, 1)
                                if ballpos > 0.65 then  Ycoef = LinearEnvelope(aBall.Y, YcoefIn, YcoefOut)                                                'find safety coefficient 'ycoef' data
                        End If

                        'Velocity correction
                        if not IsEmpty(VelocityIn(0) ) then
                                Dim VelCoef
         :                         VelCoef = LinearEnvelope(BallPos, VelocityIn, VelocityOut)

                                if partialflipcoef < 1 then VelCoef = PSlope(partialflipcoef, 0, 1, 1, VelCoef)

                                if Enabled then aBall.Velx = aBall.Velx*VelCoef
                                if Enabled then aBall.Vely = aBall.Vely*VelCoef
                        End If

                        'Polarity Correction (optional now)
                        if not IsEmpty(PolarityIn(0) ) then
                                If StartPoint > EndPoint then LR = -1        'Reverse polarity if left flipper
                                dim AddX : AddX = LinearEnvelope(BallPos, PolarityIn, PolarityOut) * LR

                                if Enabled then aBall.VelX = aBall.VelX + 1 * (AddX*ycoef*PartialFlipcoef)
                                'playsound "fx_knocker"
                        End If
                End If
                RemoveBall aBall
        End Sub
End Class

'******************************************************
'                FLIPPER POLARITY AND RUBBER DAMPENER
'                        SUPPORTING FUNCTIONS
'******************************************************

' Used for flipper correction and rubber dampeners
Sub ShuffleArray(ByRef aArray, byVal offset) 'shuffle 1d array
        dim x, aCount : aCount = 0
        redim a(uBound(aArray) )
        for x = 0 to uBound(aArray)        'Shuffle objects in a temp array
                if not IsEmpty(aArray(x) ) Then
                        if IsObject(aArray(x)) then
                                Set a(aCount) = aArray(x)
                        Else
                                a(aCount) = aArray(x)
                        End If
                        aCount = aCount + 1
                End If
        Next
        if offset < 0 then offset = 0
        redim aArray(aCount-1+offset)        'Resize original array
        for x = 0 to aCount-1                'set objects back into original array
                if IsObject(a(x)) then
                        Set aArray(x) = a(x)
                Else
                        aArray(x) = a(x)
                End If
        Next
End Sub

' Used for flipper correction and rubber dampeners
Sub ShuffleArrays(aArray1, aArray2, offset)
        ShuffleArray aArray1, offset
        ShuffleArray aArray2, offset
End Sub

' Used for flipper correction, rubber dampeners, and drop targets
Function BallSpeed(ball) 'Calculates the ball speed
    BallSpeed = SQR(ball.VelX^2 + ball.VelY^2 + ball.VelZ^2)
End Function

' Used for flipper correction and rubber dampeners
Function PSlope(Input, X1, Y1, X2, Y2)        'Set up line via two points, no clamping. Input X, output Y
        dim x, y, b, m : x = input : m = (Y2 - Y1) / (X2 - X1) : b = Y2 - m*X2
        Y = M*x+b
        PSlope = Y
End Function

' Used for flipper correction
Class spoofball
        Public X, Y, Z, VelX, VelY, VelZ, ID, Mass, Radius
        Public Property Let Data(aBall)
                With aBall
                        x = .x : y = .y : z = .z : velx = .velx : vely = .vely : velz = .velz
                        id = .ID : mass = .mass : radius = .radius
                end with
        End Property
        Public Sub Reset()
                x = Empty : y = Empty : z = Empty  : velx = Empty : vely = Empty : velz = Empty
                id = Empty : mass = Empty : radius = Empty
        End Sub
End Class

' Used for flipper correction and rubber dampeners
Function LinearEnvelope(xInput, xKeyFrame, yLvl)
        dim y 'Y output
        dim L 'Line
        dim ii : for ii = 1 to uBound(xKeyFrame)        'find active line
                if xInput <= xKeyFrame(ii) then L = ii : exit for : end if
        Next
        if xInput > xKeyFrame(uBound(xKeyFrame) ) then L = uBound(xKeyFrame)        'catch line overrun
        Y = pSlope(xInput, xKeyFrame(L-1), yLvl(L-1), xKeyFrame(L), yLvl(L) )

        if xInput <= xKeyFrame(lBound(xKeyFrame) ) then Y = yLvl(lBound(xKeyFrame) )         'Clamp lower
        if xInput >= xKeyFrame(uBound(xKeyFrame) ) then Y = yLvl(uBound(xKeyFrame) )        'Clamp upper

        LinearEnvelope = Y
End Function


'******************************************************
'                        FLIPPER TRICKS
'******************************************************

RightFlipper.timerinterval=1
Rightflipper.timerenabled=True

sub RightFlipper_timer()
        FlipperTricks LeftFlipper, LFPress, LFCount, LFEndAngle, LFState
        FlipperTricks RightFlipper, RFPress, RFCount, RFEndAngle, RFState
        FlipperNudge RightFlipper, RFEndAngle, RFEOSNudge, LeftFlipper, LFEndAngle
        FlipperNudge LeftFlipper, LFEndAngle, LFEOSNudge,  RightFlipper, RFEndAngle
end sub

Dim LFEOSNudge, RFEOSNudge

Sub FlipperNudge(Flipper1, Endangle1, EOSNudge1, Flipper2, EndAngle2)
        Dim BOT, b

        If Flipper1.currentangle = Endangle1 and EOSNudge1 <> 1 Then
                EOSNudge1 = 1
                'debug.print Flipper1.currentangle &" = "& Endangle1 &"--"& Flipper2.currentangle &" = "& EndAngle2
                If Flipper2.currentangle = EndAngle2 Then
                        BOT = GetBalls
                        For b = 0 to Ubound(BOT)
                                If FlipperTrigger(BOT(b).x, BOT(b).y, Flipper1) Then
                                        'Debug.Print "ball in flip1. exit"
                                         exit Sub
                                end If
                        Next
                        For b = 0 to Ubound(BOT)
                                If FlipperTrigger(BOT(b).x, BOT(b).y, Flipper2) Then
                                        BOT(b).velx = BOT(b).velx / 1.7
                                        BOT(b).vely = BOT(b).vely - 1
                                end If
                        Next
                End If
        Else
                If Flipper1.currentangle <> EndAngle1 then
                        EOSNudge1 = 0
                end if
        End If
End Sub

'*****************
' Maths
'*****************
Const PI = 3.1415927

Function dSin(degrees)
        dsin = sin(degrees * Pi/180)
End Function

Function dCos(degrees)
        dcos = cos(degrees * Pi/180)
End Function

'*************************************************
' Check ball distance from Flipper for Rem
'*************************************************

Function Distance(ax,ay,bx,by)
        Distance = SQR((ax - bx)^2 + (ay - by)^2)
End Function

Function DistancePL(px,py,ax,ay,bx,by) ' Distance between a point and a line where point is px,py
        DistancePL = ABS((by - ay)*px - (bx - ax) * py + bx*ay - by*ax)/Distance(ax,ay,bx,by)
End Function

Function Radians(Degrees)
        Radians = Degrees * PI /180
End Function

Function AnglePP(ax,ay,bx,by)
        AnglePP = Atn2((by - ay),(bx - ax))*180/PI
End Function

Function DistanceFromFlipper(ballx, bally, Flipper)
        DistanceFromFlipper = DistancePL(ballx, bally, Flipper.x, Flipper.y, Cos(Radians(Flipper.currentangle+90))+Flipper.x, Sin(Radians(Flipper.currentangle+90))+Flipper.y)
End Function

Function FlipperTrigger(ballx, bally, Flipper)
        Dim DiffAngle
        DiffAngle  = ABS(Flipper.currentangle - AnglePP(Flipper.x, Flipper.y, ballx, bally) - 90)
        If DiffAngle > 180 Then DiffAngle = DiffAngle - 360

        If DistanceFromFlipper(ballx,bally,Flipper) < 48 and DiffAngle <= 90 and Distance(ballx,bally,Flipper.x,Flipper.y) < Flipper.Length Then
                FlipperTrigger = True
        Else
                FlipperTrigger = False
        End If
End Function

' Used for drop targets and stand up targets
Function Atn2(dy, dx)
        dim pi
        pi = 4*Atn(1)

        If dx > 0 Then
                Atn2 = Atn(dy / dx)
        ElseIf dx < 0 Then
                If dy = 0 Then
                        Atn2 = pi
                Else
                        Atn2 = Sgn(dy) * (pi - Atn(Abs(dy / dx)))
                end if
        ElseIf dx = 0 Then
                if dy = 0 Then
                        Atn2 = 0
                else
                        Atn2 = Sgn(dy) * pi / 2
                end if
        End If
End Function

'*************************************************
' End - Check ball distance from Flipper for Rem
'*************************************************

dim LFPress, RFPress, LFCount, RFCount
dim LFState, RFState
dim EOST, EOSA,Frampup, FElasticity,FReturn
dim RFEndAngle, LFEndAngle

EOST = leftflipper.eostorque
EOSA = leftflipper.eostorqueangle
Frampup = LeftFlipper.rampup
FElasticity = LeftFlipper.elasticity
FReturn = LeftFlipper.return
Const EOSTnew = 0.8
Const EOSAnew = 1
Const EOSRampup = 0
Dim SOSRampup
Select Case FlipperCoilRampupMode
  Case 0:
    SOSRampup = 2.5
  Case 1:
    SOSRampup = 6
  Case 2:
    SOSRampup = 8.5
End Select
Const LiveCatch = 16
Const LiveElasticity = 0.45
Const SOSEM = 0.815
Const EOSReturn = 0.025

LFEndAngle = Leftflipper.endangle
RFEndAngle = RightFlipper.endangle

Sub FlipperActivate(Flipper, FlipperPress)
        FlipperPress = 1
        Flipper.Elasticity = FElasticity

        Flipper.eostorque = EOST
        Flipper.eostorqueangle = EOSA
End Sub

Sub FlipperDeactivate(Flipper, FlipperPress)
        FlipperPress = 0
        Flipper.eostorqueangle = EOSA
        Flipper.eostorque = EOST*EOSReturn/FReturn


        If Abs(Flipper.currentangle) <= Abs(Flipper.endangle) + 0.1 Then
                Dim BOT, b
                BOT = GetBalls

                For b = 0 to UBound(BOT)
                        If Distance(BOT(b).x, BOT(b).y, Flipper.x, Flipper.y) < 55 Then 'check for cradle
                                If BOT(b).vely >= -0.4 Then BOT(b).vely = -0.4
                        End If
                Next
        End If
End Sub

Sub FlipperTricks (Flipper, FlipperPress, FCount, FEndAngle, FState)
        Dim Dir
        Dir = Flipper.startangle/Abs(Flipper.startangle)        '-1 for Right Flipper

        If Abs(Flipper.currentangle) > Abs(Flipper.startangle) - 0.05 Then
                If FState <> 1 Then
                        Flipper.rampup = SOSRampup
                        Flipper.endangle = FEndAngle - 3*Dir
                        Flipper.Elasticity = FElasticity * SOSEM
                        FCount = 0
                        FState = 1
                End If
        ElseIf Abs(Flipper.currentangle) <= Abs(Flipper.endangle) and FlipperPress = 1 then
                if FCount = 0 Then FCount = GameTime

                If FState <> 2 Then
                        Flipper.eostorqueangle = EOSAnew
                        Flipper.eostorque = EOSTnew
                        Flipper.rampup = EOSRampup
                        Flipper.endangle = FEndAngle
                        FState = 2
                End If
        Elseif Abs(Flipper.currentangle) > Abs(Flipper.endangle) + 0.01 and FlipperPress = 1 Then
                If FState <> 3 Then
                        Flipper.eostorque = EOST
                        Flipper.eostorqueangle = EOSA
                        Flipper.rampup = Frampup
                        Flipper.Elasticity = FElasticity
                        FState = 3
                End If

        End If
End Sub

Const LiveDistanceMin = 30  'minimum distance in vp units from flipper base live catch dampening will occur
Const LiveDistanceMax = 114  'maximum distance in vp units from flipper base live catch dampening will occur (tip protection)

Sub CheckLiveCatch(ball, Flipper, FCount, parm) 'Experimental new live catch
        Dim Dir
        Dir = Flipper.startangle/Abs(Flipper.startangle)    '-1 for Right Flipper
        Dim LiveCatchBounce                                                                                                                        'If live catch is not perfect, it won't freeze ball totally
        Dim CatchTime : CatchTime = GameTime - FCount

        if CatchTime <= LiveCatch and parm > 6 and ABS(Flipper.x - ball.x) > LiveDistanceMin and ABS(Flipper.x - ball.x) < LiveDistanceMax Then
                if CatchTime <= LiveCatch*0.5 Then                                                'Perfect catch only when catch time happens in the beginning of the window
                        LiveCatchBounce = 0
                else
                        LiveCatchBounce = Abs((LiveCatch/2) - CatchTime)        'Partial catch when catch happens a bit late
                end If

                If LiveCatchBounce = 0 and ball.velx * Dir > 0 Then ball.velx = 0
                ball.vely = LiveCatchBounce * (32 / LiveCatch) ' Multiplier for inaccuracy bounce
                ball.angmomx= 0
                ball.angmomy= 0
                ball.angmomz= 0
        End If
End Sub


'****************************************************************************
'PHYSICS DAMPENERS
'****************************************************************************

'These are data mined bounce curves,
'dialed in with the in-game elasticity as much as possible to prevent angle / spin issues.
'Requires tracking ballspeed to calculate COR

'******************************************************
' VPW TargetBouncer for targets and posts by Iaakki, Wrd1972, Apophis
'******************************************************

sub TargetBouncer(aBall,defvalue)
    dim zMultiplier, vel, vratio
    if TargetBouncerEnabled <> 0 and aball.z < 30 then
        'debug.print "velx: " & aball.velx & " vely: " & aball.vely & " velz: " & aball.velz
        vel = BallSpeed(aBall)
        vratio = aBall.vely/aBall.velx
        Select Case Int(Rnd * 6) + 1
            Case 1: zMultiplier = 0.1*defvalue
      Case 2: zMultiplier = 0.2*defvalue
            Case 3: zMultiplier = 0.3*defvalue
      Case 4: zMultiplier = 0.4*defvalue
            Case 5: zMultiplier = 0.5*defvalue
            Case 6: zMultiplier = 0.6*defvalue
        End Select
        aBall.velz = abs(vel * zMultiplier * TargetBouncerFactor)
        aBall.velx = sgn(aBall.velx) * sqr(abs((vel^2 - aBall.velz^2)/(1+vratio^2)))
        aBall.vely = aBall.velx * vratio
        'debug.print "---> velx: " & aball.velx & " vely: " & aball.vely & " velz: " & aball.velz
        'debug.print "conservation check: " & BallSpeed(aBall)/vel
    end if
end sub

Sub dPosts_Hit(idx)
        RubbersD.dampen Activeball
    TargetBouncer activeball, 1
End Sub

Sub dSleeves_Hit(idx)
        SleevesD.Dampen Activeball
    TargetBouncer activeball, 0.7
End Sub


dim RubbersD : Set RubbersD = new Dampener        'frubber
RubbersD.name = "Rubbers"
RubbersD.debugOn = False        'shows info in textbox "TBPout"
RubbersD.Print = False        'debug, reports in debugger (in vel, out cor)
'cor bounce curve (linear)
'for best results, try to match in-game velocity as closely as possible to the desired curve
'RubbersD.addpoint 0, 0, 0.935        'point# (keep sequential), ballspeed, CoR (elasticity)
RubbersD.addpoint 0, 0, 0.96        'point# (keep sequential), ballspeed, CoR (elasticity)
RubbersD.addpoint 1, 3.77, 0.96
RubbersD.addpoint 2, 5.76, 0.967        'dont take this as gospel. if you can data mine rubber elasticitiy, please help!
RubbersD.addpoint 3, 15.84, 0.874
RubbersD.addpoint 4, 56, 0.64        'there's clamping so interpolate up to 56 at least

dim SleevesD : Set SleevesD = new Dampener        'this is just rubber but cut down to 85%...
SleevesD.name = "Sleeves"
SleevesD.debugOn = False        'shows info in textbox "TBPout"
SleevesD.Print = False        'debug, reports in debugger (in vel, out cor)
SleevesD.CopyCoef RubbersD, 0.85

Class Dampener
        Public Print, debugOn 'tbpOut.text
        public name, Threshold         'Minimum threshold. Useful for Flippers, which don't have a hit threshold.
        Public ModIn, ModOut
        Private Sub Class_Initialize : redim ModIn(0) : redim Modout(0): End Sub

        Public Sub AddPoint(aIdx, aX, aY)
                ShuffleArrays ModIn, ModOut, 1 : ModIn(aIDX) = aX : ModOut(aIDX) = aY : ShuffleArrays ModIn, ModOut, 0
                if gametime > 100 then Report
        End Sub

        public sub Dampen(aBall)
                if threshold then if BallSpeed(aBall) < threshold then exit sub end if end if
                dim RealCOR, DesiredCOR, str, coef
                DesiredCor = LinearEnvelope(cor.ballvel(aBall.id), ModIn, ModOut )
                RealCOR = BallSpeed(aBall) / cor.ballvel(aBall.id)
                coef = desiredcor / realcor
                if debugOn then str = name & " in vel:" & round(cor.ballvel(aBall.id),2 ) & vbnewline & "desired cor: " & round(desiredcor,4) & vbnewline & _
                "actual cor: " & round(realCOR,4) & vbnewline & "ballspeed coef: " & round(coef, 3) & vbnewline
                if Print then debug.print Round(cor.ballvel(aBall.id),2) & ", " & round(desiredcor,3)

' Thalamus - patched :                 aBall.velx = aBall.velx * coef : aBall.vely = aBall.vely * coef
    aBall.velx = aBall.velx * coef : aBall.vely = aBall.vely * coef : aBall.velz = aBall.velz * coef
                if debugOn then TBPout.text = str
        End Sub

        Public Sub CopyCoef(aObj, aCoef) 'alternative addpoints, copy with coef
                dim x : for x = 0 to uBound(aObj.ModIn)
                        addpoint x, aObj.ModIn(x), aObj.ModOut(x)*aCoef
                Next
        End Sub


        Public Sub Report()         'debug, reports all coords in tbPL.text
                if not debugOn then exit sub
                dim a1, a2 : a1 = ModIn : a2 = ModOut
                dim str, x : for x = 0 to uBound(a1) : str = str & x & ": " & round(a1(x),4) & ", " & round(a2(x),4) & vbnewline : next
                TBPout.text = str
        End Sub


End Class


'******************************************************
'                TRACK ALL BALL VELOCITIES
'                 FOR RUBBER DAMPENER AND DROP TARGETS
'******************************************************

dim cor : set cor = New CoRTracker

Class CoRTracker
        public ballvel, ballvelx, ballvely

        Private Sub Class_Initialize : redim ballvel(0) : redim ballvelx(0): redim ballvely(0) : End Sub

        Public Sub Update()        'tracks in-ball-velocity
                dim str, b, AllBalls, highestID : allBalls = getballs

                for each b in allballs
                        if b.id >= HighestID then highestID = b.id
                Next

                if uBound(ballvel) < highestID then redim ballvel(highestID)        'set bounds
                if uBound(ballvelx) < highestID then redim ballvelx(highestID)        'set bounds
                if uBound(ballvely) < highestID then redim ballvely(highestID)        'set bounds

                for each b in allballs
                        ballvel(b.id) = BallSpeed(b)
                        ballvelx(b.id) = b.velx
                        ballvely(b.id) = b.vely
                Next
        End Sub
End Class

'*** Determines if a Points (px,py) is inside a 4 point polygon A-D in Clockwise/CCW order
Function InRect(px,py,ax,ay,bx,by,cx,cy,dx,dy)
        Dim AB, BC, CD, DA
        AB = (bx*py) - (by*px) - (ax*py) + (ay*px) + (ax*by) - (ay*bx)
        BC = (cx*py) - (cy*px) - (bx*py) + (by*px) + (bx*cy) - (by*cx)
        CD = (dx*py) - (dy*px) - (cx*py) + (cy*px) + (cx*dy) - (cy*dx)
        DA = (ax*py) - (ay*px) - (dx*py) + (dy*px) + (dx*ay) - (dy*ax)

        If (AB <= 0 AND BC <=0 AND CD <= 0 AND DA <= 0) Or (AB >= 0 AND BC >=0 AND CD >= 0 AND DA >= 0) Then
                InRect = True
        Else
                InRect = False
        End If
End Function

Sub RDampen_Timer()
Cor.Update
End Sub

'////////////////////////////  MECHANICAL SOUNDS  ///////////////////////////
'//  This part in the script is an entire block that is dedicated to the physics sound system.
'//  Various scripts and sounds that may be pretty generic and could suit other WPC systems, but the most are tailored specifically for this table.

'///////////////////////////////  SOUNDS PARAMETERS  //////////////////////////////
Dim GlobalSoundLevel, CoinSoundLevel, PlungerReleaseSoundLevel, PlungerPullSoundLevel, NudgeLeftSoundLevel
Dim NudgeRightSoundLevel, NudgeCenterSoundLevel, StartButtonSoundLevel, RollingSoundFactor

CoinSoundLevel = 1                            'volume level; range [0, 1]
NudgeLeftSoundLevel = 1                         'volume level; range [0, 1]
NudgeRightSoundLevel = 1                        'volume level; range [0, 1]
NudgeCenterSoundLevel = 1                       'volume level; range [0, 1]
StartButtonSoundLevel = 0.1                       'volume level; range [0, 1]
PlungerReleaseSoundLevel = 0.8 '1 wjr                     'volume level; range [0, 1]
PlungerPullSoundLevel = 1                       'volume level; range [0, 1]
RollingSoundFactor = 1.1/5

'///////////////////////-----Solenoids, Kickers and Flash Relays-----///////////////////////
Dim FlipperUpAttackMinimumSoundLevel, FlipperUpAttackMaximumSoundLevel, FlipperUpAttackLeftSoundLevel, FlipperUpAttackRightSoundLevel
Dim FlipperUpSoundLevel, FlipperDownSoundLevel, FlipperLeftHitParm, FlipperRightHitParm
Dim SlingshotSoundLevel, BumperSoundFactor, KnockerSoundLevel

FlipperUpAttackMinimumSoundLevel = 0.010                      'volume level; range [0, 1]
FlipperUpAttackMaximumSoundLevel = 0.635                'volume level; range [0, 1]
FlipperUpSoundLevel = 1.0                                   'volume level; range [0, 1]
FlipperDownSoundLevel = 0.45                                  'volume level; range [0, 1]
FlipperLeftHitParm = FlipperUpSoundLevel                'sound helper; not configurable
FlipperRightHitParm = FlipperUpSoundLevel               'sound helper; not configurable
SlingshotSoundLevel = 0.95                        'volume level; range [0, 1]
BumperSoundFactor = 4.25                        'volume multiplier; must not be zero
KnockerSoundLevel = 1                           'volume level; range [0, 1]

'///////////////////////-----Ball Drops, Bumps and Collisions-----///////////////////////
Dim RubberStrongSoundFactor, RubberWeakSoundFactor, RubberFlipperSoundFactor,BallWithBallCollisionSoundFactor
Dim BallBouncePlayfieldSoftFactor, BallBouncePlayfieldHardFactor, PlasticRampDropToPlayfieldSoundLevel, WireRampDropToPlayfieldSoundLevel, DelayedBallDropOnPlayfieldSoundLevel
Dim WallImpactSoundFactor, MetalImpactSoundFactor, SubwaySoundLevel, SubwayEntrySoundLevel, ScoopEntrySoundLevel
Dim SaucerLockSoundLevel, SaucerKickSoundLevel

BallWithBallCollisionSoundFactor = 3.2                  'volume multiplier; must not be zero
RubberStrongSoundFactor = 0.055/5                     'volume multiplier; must not be zero
RubberWeakSoundFactor = 0.075/5                     'volume multiplier; must not be zero
RubberFlipperSoundFactor = 0.075/5                    'volume multiplier; must not be zero
BallBouncePlayfieldSoftFactor = 0.025                 'volume multiplier; must not be zero
BallBouncePlayfieldHardFactor = 0.025                 'volume multiplier; must not be zero
DelayedBallDropOnPlayfieldSoundLevel = 0.8                  'volume level; range [0, 1]
WallImpactSoundFactor = 0.075                     'volume multiplier; must not be zero
MetalImpactSoundFactor = 0.075/3
SaucerLockSoundLevel = 0.8
SaucerKickSoundLevel = 0.8

'///////////////////////-----Gates, Spinners, Rollovers and Targets-----///////////////////////

Dim GateSoundLevel, TargetSoundFactor, SpinnerSoundLevel, RolloverSoundLevel, DTSoundLevel

GateSoundLevel = 0.5/5                          'volume level; range [0, 1]
TargetSoundFactor = 0.0025 * 10                     'volume multiplier; must not be zero
DTSoundLevel = 0.25                           'volume multiplier; must not be zero
RolloverSoundLevel = 0.25                                       'volume level; range [0, 1]

'///////////////////////-----Ball Release, Guides and Drain-----///////////////////////
Dim DrainSoundLevel, BallReleaseSoundLevel, BottomArchBallGuideSoundFactor, FlipperBallGuideSoundFactor

DrainSoundLevel = 0.8                           'volume level; range [0, 1]
BallReleaseSoundLevel = 1                       'volume level; range [0, 1]
BottomArchBallGuideSoundFactor = 0.2                  'volume multiplier; must not be zero
FlipperBallGuideSoundFactor = 0.015                   'volume multiplier; must not be zero

'///////////////////////-----Loops and Lanes-----///////////////////////
Dim ArchSoundFactor
ArchSoundFactor = 0.025/5                         'volume multiplier; must not be zero


'/////////////////////////////  SOUND PLAYBACK FUNCTIONS  ////////////////////////////
'/////////////////////////////  POSITIONAL SOUND PLAYBACK METHODS  ////////////////////////////
' Positional sound playback methods will play a sound, depending on the X,Y position of the table element or depending on ActiveBall object position
' These are similar subroutines that are less complicated to use (e.g. simply use standard parameters for the PlaySound call)
' For surround setup - positional sound playback functions will fade between front and rear surround channels and pan between left and right channels
' For stereo setup - positional sound playback functions will only pan between left and right channels
' For mono setup - positional sound playback functions will not pan between left and right channels and will not fade between front and rear channels

' PlaySound full syntax - PlaySound(string, int loopcount, float volume, float pan, float randompitch, int pitch, bool useexisting, bool restart, float front_rear_fade)
' Note - These functions will not work (currently) for walls/slingshots as these do not feature a simple, single X,Y position
Sub PlaySoundAtLevelStatic(playsoundparams, aVol, tableobj)
    PlaySound playsoundparams, 0, aVol * VolumeDial, AudioPan(tableobj), 0, 0, 0, 0, AudioFade(tableobj)
End Sub

Sub PlaySoundAtLevelExistingStatic(playsoundparams, aVol, tableobj)
    PlaySound playsoundparams, 0, aVol * VolumeDial, AudioPan(tableobj), 0, 0, 1, 0, AudioFade(tableobj)
End Sub

Sub PlaySoundAtLevelStaticLoop(playsoundparams, aVol, tableobj)
    PlaySound playsoundparams, -1, aVol * VolumeDial, AudioPan(tableobj), 0, 0, 0, 0, AudioFade(tableobj)
End Sub

Sub PlaySoundAtLevelStaticRandomPitch(playsoundparams, aVol, randomPitch, tableobj)
    PlaySound playsoundparams, 0, aVol * VolumeDial, AudioPan(tableobj), randomPitch, 0, 0, 0, AudioFade(tableobj)
End Sub

Sub PlaySoundAtLevelActiveBall(playsoundparams, aVol)
  PlaySound playsoundparams, 0, aVol * VolumeDial, AudioPan(ActiveBall), 0, 0, 0, 0, AudioFade(ActiveBall)
End Sub

Sub PlaySoundAtLevelExistingActiveBall(playsoundparams, aVol)
  PlaySound playsoundparams, 0, aVol * VolumeDial, AudioPan(ActiveBall), 0, 0, 1, 0, AudioFade(ActiveBall)
End Sub

Sub PlaySoundAtLeveTimerActiveBall(playsoundparams, aVol, ballvariable)
  PlaySound playsoundparams, 0, aVol * VolumeDial, AudioPan(ballvariable), 0, 0, 0, 0, AudioFade(ballvariable)
End Sub

Sub PlaySoundAtLevelTimerExistingActiveBall(playsoundparams, aVol, ballvariable)
  PlaySound playsoundparams, 0, aVol * VolumeDial, AudioPan(ballvariable), 0, 0, 1, 0, AudioFade(ballvariable)
End Sub

Sub PlaySoundAtLevelRoll(playsoundparams, aVol, pitch)
    PlaySound playsoundparams, -1, aVol * VolumeDial, AudioPan(tableobj), randomPitch, 0, 0, 0, AudioFade(tableobj)
End Sub

' Previous Positional Sound Subs

Sub PlaySoundAt(soundname, tableobj)
    PlaySound soundname, 1, 1 * VolumeDial, AudioPan(tableobj), 0,0,0, 1, AudioFade(tableobj)
End Sub

Sub PlaySoundAtVol(soundname, tableobj, aVol)
    PlaySound soundname, 1, aVol * VolumeDial, AudioPan(tableobj), 0,0,0, 1, AudioFade(tableobj)
End Sub

Sub PlaySoundAtBall(soundname)
    PlaySoundAt soundname, ActiveBall
End Sub

Sub PlaySoundAtBallVol (Soundname, aVol)
  Playsound soundname, 1,aVol * VolumeDial, AudioPan(ActiveBall), 0,0,0, 1, AudioFade(ActiveBall)
End Sub

Sub PlaySoundAtBallVolM (Soundname, aVol)
  Playsound soundname, 1,aVol * VolumeDial, AudioPan(ActiveBall), 0,0,0, 0, AudioFade(ActiveBall)
End Sub

Sub PlaySoundAtVolLoops(sound, tableobj, Vol, Loops)
  PlaySound sound, Loops, Vol * VolumeDial, AudioPan(tableobj), 0,0,0, 1, AudioFade(tableobj)
End Sub


' *********************************************************************
'                     Fleep  Supporting Ball & Sound Functions
' *********************************************************************

Dim tablewidth, tableheight : tablewidth = table1.width : tableheight = table1.height

Function AudioFade(tableobj) ' Fades between front and back of the table (for surround systems or 2x2 speakers, etc), depending on the Y position on the table. "table1" is the name of the table
  Dim tmp
  tmp = tableobj.y * 2 / tableheight-1

  if tmp > 7000 Then
    tmp = 7000
  elseif tmp < -7000 Then
    tmp = -7000
  end if

      If tmp > 0 Then
    AudioFade = Csng(tmp ^10)
  Else
    AudioFade = Csng(-((- tmp) ^10) )
  End If
End Function

Function AudioPan(tableobj) ' Calculates the pan for a tableobj based on the X position on the table. "table1" is the name of the table
  Dim tmp
  tmp = tableobj.x * 2 / tablewidth-1

  if tmp > 7000 Then
    tmp = 7000
  elseif tmp < -7000 Then
    tmp = -7000
  end if

  If tmp > 0 Then
    AudioPan = Csng(tmp ^10)
  Else
    AudioPan = Csng(-((- tmp) ^10) )
  End If
End Function

Function Vol(ball) ' Calculates the volume of the sound based on the ball speed
  Vol = Csng(BallVel(ball) ^2)
End Function

Function Volz(ball) ' Calculates the volume of the sound based on the ball speed
  Volz = Csng((ball.velz) ^2)
End Function

Function Pitch(ball) ' Calculates the pitch of the sound based on the ball speed
    Pitch = BallVel(ball) * 20
End Function

Function BallVel(ball) 'Calculates the ball speed
    BallVel = INT(SQR((ball.VelX ^2) + (ball.VelY ^2) ) )
End Function

Function VolPlayfieldRoll(ball) ' Calculates the roll volume of the sound based on the ball speed
  VolPlayfieldRoll = RollingSoundFactor * 0.0005 * Csng(BallVel(ball) ^3)
End Function

Function PitchPlayfieldRoll(ball) ' Calculates the roll pitch of the sound based on the ball speed
    PitchPlayfieldRoll = BallVel(ball) ^2 * 15
End Function

Function RndInt(min, max)
    RndInt = Int(Rnd() * (max-min + 1) + min)' Sets a random number integer between min and max
End Function

Function RndNum(min, max)
    RndNum = Rnd() * (max-min) + min' Sets a random number between min and max
End Function

'/////////////////////////////  GENERAL SOUND SUBROUTINES  ////////////////////////////
Sub SoundStartButton()
  PlaySound ("Start_Button"), 0, StartButtonSoundLevel, 0, 0.25
End Sub

Sub SoundNudgeLeft()
  PlaySound ("Nudge_" & Int(Rnd*2)+1), 0, NudgeLeftSoundLevel * VolumeDial, -0.1, 0.25
End Sub

Sub SoundNudgeRight()
  PlaySound ("Nudge_" & Int(Rnd*2)+1), 0, NudgeRightSoundLevel * VolumeDial, 0.1, 0.25
End Sub

Sub SoundNudgeCenter()
  PlaySound ("Nudge_" & Int(Rnd*2)+1), 0, NudgeCenterSoundLevel * VolumeDial, 0, 0.25
End Sub


Sub SoundPlungerPull()
  PlaySoundAtLevelStatic ("Plunger_Pull_1"), PlungerPullSoundLevel, Plunger
End Sub

Sub SoundPlungerReleaseBall()
  PlaySoundAtLevelStatic ("Plunger_Release_Ball"), PlungerReleaseSoundLevel, Plunger
End Sub

Sub SoundPlungerReleaseNoBall()
  PlaySoundAtLevelStatic ("Plunger_Release_No_Ball"), PlungerReleaseSoundLevel, Plunger
End Sub


'/////////////////////////////  KNOCKER SOLENOID  ////////////////////////////
Sub KnockerSolenoid()
  PlaySoundAtLevelStatic SoundFX("Knocker_1",DOFKnocker), KnockerSoundLevel, KnockerPosition
End Sub

'/////////////////////////////  DRAIN SOUNDS  ////////////////////////////
Sub RandomSoundDrain(drainswitch)
  PlaySoundAtLevelStatic ("Drain_" & Int(Rnd*11)+1), DrainSoundLevel, drainswitch
End Sub

'/////////////////////////////  TROUGH BALL RELEASE SOLENOID SOUNDS  ////////////////////////////

Sub RandomSoundBallRelease(drainswitch)
  PlaySoundAtLevelStatic SoundFX("BallRelease" & Int(Rnd*7)+1,DOFContactors), BallReleaseSoundLevel, drainswitch
End Sub

'/////////////////////////////  SLINGSHOT SOLENOID SOUNDS  ////////////////////////////
Sub RandomSoundSlingshotLeft(sling)
  PlaySoundAtLevelStatic SoundFX("Sling_L" & Int(Rnd*10)+1,DOFContactors), SlingshotSoundLevel, Sling
End Sub

Sub RandomSoundSlingshotRight(sling)
  PlaySoundAtLevelStatic SoundFX("Sling_R" & Int(Rnd*8)+1,DOFContactors), SlingshotSoundLevel, Sling
End Sub

'/////////////////////////////  BUMPER SOLENOID SOUNDS  ////////////////////////////
Sub RandomSoundBumperTop(Bump)
  PlaySoundAtLevelStatic SoundFX("Bumpers_Top_" & Int(Rnd*5)+1,DOFContactors), Vol(ActiveBall) * BumperSoundFactor, Bump
End Sub

Sub RandomSoundBumperMiddle(Bump)
  PlaySoundAtLevelStatic SoundFX("Bumpers_Middle_" & Int(Rnd*5)+1,DOFContactors), Vol(ActiveBall) * BumperSoundFactor, Bump
End Sub

Sub RandomSoundBumperBottom(Bump)
  PlaySoundAtLevelStatic SoundFX("Bumpers_Bottom_" & Int(Rnd*5)+1,DOFContactors), Vol(ActiveBall) * BumperSoundFactor, Bump
End Sub

'/////////////////////////////  FLIPPER BATS SOUND SUBROUTINES  ////////////////////////////
'/////////////////////////////  FLIPPER BATS SOLENOID ATTACK SOUND  ////////////////////////////
Sub SoundFlipperUpAttackLeft(flipper)
  FlipperUpAttackLeftSoundLevel = RndNum(FlipperUpAttackMinimumSoundLevel, FlipperUpAttackMaximumSoundLevel)
  PlaySoundAtLevelStatic ("Flipper_Attack-L01"), FlipperUpAttackLeftSoundLevel, flipper
End Sub

Sub SoundFlipperUpAttackRight(flipper)
  FlipperUpAttackRightSoundLevel = RndNum(FlipperUpAttackMinimumSoundLevel, FlipperUpAttackMaximumSoundLevel)
    PlaySoundAtLevelStatic ("Flipper_Attack-R01"), FlipperUpAttackLeftSoundLevel, flipper
End Sub

'/////////////////////////////  FLIPPER BATS SOLENOID CORE SOUND  ////////////////////////////
Sub RandomSoundFlipperUpLeft(flipper)
  PlaySoundAtLevelStatic SoundFX("Flipper_L0" & Int(Rnd*9)+1,DOFFlippers), FlipperLeftHitParm, Flipper
End Sub

Sub RandomSoundFlipperUpRight(flipper)
  PlaySoundAtLevelStatic SoundFX("Flipper_R0" & Int(Rnd*9)+1,DOFFlippers), FlipperRightHitParm, Flipper
End Sub

Sub RandomSoundReflipUpLeft(flipper)
  PlaySoundAtLevelStatic SoundFX("Flipper_ReFlip_L0" & Int(Rnd*3)+1,DOFFlippers), (RndNum(0.8, 1))*FlipperUpSoundLevel, Flipper
End Sub

Sub RandomSoundReflipUpRight(flipper)
  PlaySoundAtLevelStatic SoundFX("Flipper_ReFlip_R0" & Int(Rnd*3)+1,DOFFlippers), (RndNum(0.8, 1))*FlipperUpSoundLevel, Flipper
End Sub

Sub RandomSoundFlipperDownLeft(flipper)
  PlaySoundAtLevelStatic SoundFX("Flipper_Left_Down_" & Int(Rnd*7)+1,DOFFlippers), FlipperDownSoundLevel, Flipper
End Sub

Sub RandomSoundFlipperDownRight(flipper)
  PlaySoundAtLevelStatic SoundFX("Flipper_Right_Down_" & Int(Rnd*8)+1,DOFFlippers), FlipperDownSoundLevel, Flipper
End Sub

'/////////////////////////////  FLIPPER BATS BALL COLLIDE SOUND  ////////////////////////////

Sub LeftFlipperCollide(parm)
  FlipperLeftHitParm = parm/10
  If FlipperLeftHitParm > 1 Then
    FlipperLeftHitParm = 1
  End If
  FlipperLeftHitParm = FlipperUpSoundLevel * FlipperLeftHitParm
  RandomSoundRubberFlipper(parm)
End Sub

Sub RightFlipperCollide(parm)
  FlipperRightHitParm = parm/10
  If FlipperRightHitParm > 1 Then
    FlipperRightHitParm = 1
  End If
  FlipperRightHitParm = FlipperUpSoundLevel * FlipperRightHitParm
  RandomSoundRubberFlipper(parm)
End Sub

Sub RandomSoundRubberFlipper(parm)
  PlaySoundAtLevelActiveBall ("Flipper_Rubber_" & Int(Rnd*7)+1), parm  * RubberFlipperSoundFactor
End Sub

'/////////////////////////////  ROLLOVER SOUNDS  ////////////////////////////
Sub RandomSoundRollover()
  PlaySoundAtLevelActiveBall ("Rollover_" & Int(Rnd*4)+1), RolloverSoundLevel
End Sub

Sub Rollovers_Hit(idx)
  RandomSoundRollover
End Sub

'/////////////////////////////  VARIOUS PLAYFIELD SOUND SUBROUTINES  ////////////////////////////
'/////////////////////////////  RUBBERS AND POSTS  ////////////////////////////
'/////////////////////////////  RUBBERS - EVENTS  ////////////////////////////
Sub Rubbers_Hit(idx)
  dim finalspeed
    finalspeed=SQR(activeball.velx * activeball.velx + activeball.vely * activeball.vely)
  If finalspeed > 5 then
    RandomSoundRubberStrong 1
  End if
  If finalspeed <= 5 then
    RandomSoundRubberWeak()
  End If
End Sub

'/////////////////////////////  RUBBERS AND POSTS - STRONG IMPACTS  ////////////////////////////
Sub RandomSoundRubberStrong(voladj)
  Select Case Int(Rnd*10)+1
    Case 1 : PlaySoundAtLevelActiveBall ("Rubber_Strong_1"), Vol(ActiveBall) * RubberStrongSoundFactor*voladj
    Case 2 : PlaySoundAtLevelActiveBall ("Rubber_Strong_2"), Vol(ActiveBall) * RubberStrongSoundFactor*voladj
    Case 3 : PlaySoundAtLevelActiveBall ("Rubber_Strong_3"), Vol(ActiveBall) * RubberStrongSoundFactor*voladj
    Case 4 : PlaySoundAtLevelActiveBall ("Rubber_Strong_4"), Vol(ActiveBall) * RubberStrongSoundFactor*voladj
    Case 5 : PlaySoundAtLevelActiveBall ("Rubber_Strong_5"), Vol(ActiveBall) * RubberStrongSoundFactor*voladj
    Case 6 : PlaySoundAtLevelActiveBall ("Rubber_Strong_6"), Vol(ActiveBall) * RubberStrongSoundFactor*voladj
    Case 7 : PlaySoundAtLevelActiveBall ("Rubber_Strong_7"), Vol(ActiveBall) * RubberStrongSoundFactor*voladj
    Case 8 : PlaySoundAtLevelActiveBall ("Rubber_Strong_8"), Vol(ActiveBall) * RubberStrongSoundFactor*voladj
    Case 9 : PlaySoundAtLevelActiveBall ("Rubber_Strong_9"), Vol(ActiveBall) * RubberStrongSoundFactor*voladj
    Case 10 : PlaySoundAtLevelActiveBall ("Rubber_1_Hard"), Vol(ActiveBall) * RubberStrongSoundFactor * 0.6*voladj
  End Select
End Sub

'/////////////////////////////  RUBBERS AND POSTS - WEAK IMPACTS  ////////////////////////////
Sub RandomSoundRubberWeak()
  PlaySoundAtLevelActiveBall ("Rubber_" & Int(Rnd*9)+1), Vol(ActiveBall) * RubberWeakSoundFactor
End Sub

'/////////////////////////////  WALL IMPACTS  ////////////////////////////
Sub Walls_Hit(idx)
  dim finalspeed
    finalspeed=SQR(activeball.velx * activeball.velx + activeball.vely * activeball.vely)
  If finalspeed > 5 then
    RandomSoundRubberStrong 1
  End if
  If finalspeed <= 5 then
    RandomSoundRubberWeak()
  End If
End Sub

Sub RandomSoundWall()
  dim finalspeed
    finalspeed=SQR(activeball.velx * activeball.velx + activeball.vely * activeball.vely)
  If finalspeed > 16 then
    Select Case Int(Rnd*5)+1
      Case 1 : PlaySoundAtLevelExistingActiveBall ("Wall_Hit_1"), Vol(ActiveBall) * WallImpactSoundFactor
      Case 2 : PlaySoundAtLevelExistingActiveBall ("Wall_Hit_2"), Vol(ActiveBall) * WallImpactSoundFactor
      Case 3 : PlaySoundAtLevelExistingActiveBall ("Wall_Hit_5"), Vol(ActiveBall) * WallImpactSoundFactor
      Case 4 : PlaySoundAtLevelExistingActiveBall ("Wall_Hit_7"), Vol(ActiveBall) * WallImpactSoundFactor
      Case 5 : PlaySoundAtLevelExistingActiveBall ("Wall_Hit_9"), Vol(ActiveBall) * WallImpactSoundFactor
    End Select
  End if
  If finalspeed >= 6 AND finalspeed <= 16 then
    Select Case Int(Rnd*4)+1
      Case 1 : PlaySoundAtLevelExistingActiveBall ("Wall_Hit_3"), Vol(ActiveBall) * WallImpactSoundFactor
      Case 2 : PlaySoundAtLevelExistingActiveBall ("Wall_Hit_4"), Vol(ActiveBall) * WallImpactSoundFactor
      Case 3 : PlaySoundAtLevelExistingActiveBall ("Wall_Hit_6"), Vol(ActiveBall) * WallImpactSoundFactor
      Case 4 : PlaySoundAtLevelExistingActiveBall ("Wall_Hit_8"), Vol(ActiveBall) * WallImpactSoundFactor
    End Select
  End If
  If finalspeed < 6 Then
    Select Case Int(Rnd*3)+1
      Case 1 : PlaySoundAtLevelExistingActiveBall ("Wall_Hit_4"), Vol(ActiveBall) * WallImpactSoundFactor
      Case 2 : PlaySoundAtLevelExistingActiveBall ("Wall_Hit_6"), Vol(ActiveBall) * WallImpactSoundFactor
      Case 3 : PlaySoundAtLevelExistingActiveBall ("Wall_Hit_8"), Vol(ActiveBall) * WallImpactSoundFactor
    End Select
  End if
End Sub

'/////////////////////////////  METAL TOUCH SOUNDS  ////////////////////////////
Sub RandomSoundMetal()
  PlaySoundAtLevelActiveBall ("Metal_Touch_" & Int(Rnd*13)+1), Vol(ActiveBall) * MetalImpactSoundFactor
End Sub

'/////////////////////////////  METAL - EVENTS  ////////////////////////////

Sub Metals_Hit (idx)
  RandomSoundMetal
End Sub

Sub ShooterDiverter_collide(idx)
  RandomSoundMetal
End Sub

'/////////////////////////////  BOTTOM ARCH BALL GUIDE  ////////////////////////////
'/////////////////////////////  BOTTOM ARCH BALL GUIDE - SOFT BOUNCES  ////////////////////////////
Sub RandomSoundBottomArchBallGuide()
  dim finalspeed
    finalspeed=SQR(activeball.velx * activeball.velx + activeball.vely * activeball.vely)
  If finalspeed > 16 then
    PlaySoundAtLevelActiveBall ("Apron_Bounce_"& Int(Rnd*2)+1), Vol(ActiveBall) * BottomArchBallGuideSoundFactor
  End if
  If finalspeed >= 6 AND finalspeed <= 16 then
    Select Case Int(Rnd*2)+1
      Case 1 : PlaySoundAtLevelActiveBall ("Apron_Bounce_1"), Vol(ActiveBall) * BottomArchBallGuideSoundFactor
      Case 2 : PlaySoundAtLevelActiveBall ("Apron_Bounce_Soft_1"), Vol(ActiveBall) * BottomArchBallGuideSoundFactor
    End Select
  End If
  If finalspeed < 6 Then
    Select Case Int(Rnd*2)+1
      Case 1 : PlaySoundAtLevelActiveBall ("Apron_Bounce_Soft_1"), Vol(ActiveBall) * BottomArchBallGuideSoundFactor
      Case 2 : PlaySoundAtLevelActiveBall ("Apron_Medium_3"), Vol(ActiveBall) * BottomArchBallGuideSoundFactor
    End Select
  End if
End Sub

'/////////////////////////////  BOTTOM ARCH BALL GUIDE - HARD HITS  ////////////////////////////
Sub RandomSoundBottomArchBallGuideHardHit()
  PlaySoundAtLevelActiveBall ("Apron_Hard_Hit_" & Int(Rnd*3)+1), BottomArchBallGuideSoundFactor * 0.25
End Sub

Sub Apron_Hit (idx)
  If Abs(cor.ballvelx(activeball.id) < 4) and cor.ballvely(activeball.id) > 7 then
    RandomSoundBottomArchBallGuideHardHit()
  Else
    RandomSoundBottomArchBallGuide
  End If
End Sub

'/////////////////////////////  FLIPPER BALL GUIDE  ////////////////////////////
Sub RandomSoundFlipperBallGuide()
  dim finalspeed
    finalspeed=SQR(activeball.velx * activeball.velx + activeball.vely * activeball.vely)
  If finalspeed > 16 then
    Select Case Int(Rnd*2)+1
      Case 1 : PlaySoundAtLevelActiveBall ("Apron_Hard_1"),  Vol(ActiveBall) * FlipperBallGuideSoundFactor
      Case 2 : PlaySoundAtLevelActiveBall ("Apron_Hard_2"),  Vol(ActiveBall) * 0.8 * FlipperBallGuideSoundFactor
    End Select
  End if
  If finalspeed >= 6 AND finalspeed <= 16 then
    PlaySoundAtLevelActiveBall ("Apron_Medium_" & Int(Rnd*3)+1),  Vol(ActiveBall) * FlipperBallGuideSoundFactor
  End If
  If finalspeed < 6 Then
    PlaySoundAtLevelActiveBall ("Apron_Soft_" & Int(Rnd*7)+1),  Vol(ActiveBall) * FlipperBallGuideSoundFactor
  End If
End Sub

'/////////////////////////////  TARGET HIT SOUNDS  ////////////////////////////
Sub RandomSoundTargetHitStrong()
  PlaySoundAtLevelActiveBall SoundFX("Target_Hit_" & Int(Rnd*4)+5,DOFTargets), Vol(ActiveBall) * 0.45 * TargetSoundFactor
End Sub

Sub RandomSoundTargetHitWeak()
  PlaySoundAtLevelActiveBall SoundFX("Target_Hit_" & Int(Rnd*4)+1,DOFTargets), Vol(ActiveBall) * TargetSoundFactor
End Sub

Sub PlayTargetSound()
  dim finalspeed
    finalspeed=SQR(activeball.velx * activeball.velx + activeball.vely * activeball.vely)
  If finalspeed > 10 then
    RandomSoundTargetHitStrong()
    RandomSoundBallBouncePlayfieldSoft Activeball
  Else
    RandomSoundTargetHitWeak()
  End If
End Sub

Sub Targets_Hit (idx)
  PlayTargetSound
End Sub

'/////////////////////////////  BALL BOUNCE SOUNDS  ////////////////////////////
Sub RandomSoundBallBouncePlayfieldSoft(aBall)
  Select Case Int(Rnd*9)+1
    Case 1 : PlaySoundAtLevelStatic ("Ball_Bounce_Playfield_Soft_1"), volz(aBall) * BallBouncePlayfieldSoftFactor, aBall
    Case 2 : PlaySoundAtLevelStatic ("Ball_Bounce_Playfield_Soft_2"), volz(aBall) * BallBouncePlayfieldSoftFactor * 0.5, aBall
    Case 3 : PlaySoundAtLevelStatic ("Ball_Bounce_Playfield_Soft_3"), volz(aBall) * BallBouncePlayfieldSoftFactor * 0.8, aBall
    Case 4 : PlaySoundAtLevelStatic ("Ball_Bounce_Playfield_Soft_4"), volz(aBall) * BallBouncePlayfieldSoftFactor * 0.5, aBall
    Case 5 : PlaySoundAtLevelStatic ("Ball_Bounce_Playfield_Soft_5"), volz(aBall) * BallBouncePlayfieldSoftFactor, aBall
    Case 6 : PlaySoundAtLevelStatic ("Ball_Bounce_Playfield_Hard_1"), volz(aBall) * BallBouncePlayfieldSoftFactor * 0.2, aBall
    Case 7 : PlaySoundAtLevelStatic ("Ball_Bounce_Playfield_Hard_2"), volz(aBall) * BallBouncePlayfieldSoftFactor * 0.2, aBall
    Case 8 : PlaySoundAtLevelStatic ("Ball_Bounce_Playfield_Hard_5"), volz(aBall) * BallBouncePlayfieldSoftFactor * 0.2, aBall
    Case 9 : PlaySoundAtLevelStatic ("Ball_Bounce_Playfield_Hard_7"), volz(aBall) * BallBouncePlayfieldSoftFactor * 0.3, aBall
  End Select
End Sub

Sub RandomSoundBallBouncePlayfieldHard(aBall)
  PlaySoundAtLevelStatic ("Ball_Bounce_Playfield_Hard_" & Int(Rnd*7)+1), volz(aBall) * BallBouncePlayfieldHardFactor, aBall
End Sub

'/////////////////////////////  DELAYED DROP - TO PLAYFIELD - SOUND  ////////////////////////////
Sub RandomSoundDelayedBallDropOnPlayfield(aBall)
  Select Case Int(Rnd*5)+1
    Case 1 : PlaySoundAtLevelStatic ("Ball_Drop_Playfield_1_Delayed"), DelayedBallDropOnPlayfieldSoundLevel, aBall
    Case 2 : PlaySoundAtLevelStatic ("Ball_Drop_Playfield_2_Delayed"), DelayedBallDropOnPlayfieldSoundLevel, aBall
    Case 3 : PlaySoundAtLevelStatic ("Ball_Drop_Playfield_3_Delayed"), DelayedBallDropOnPlayfieldSoundLevel, aBall
    Case 4 : PlaySoundAtLevelStatic ("Ball_Drop_Playfield_4_Delayed"), DelayedBallDropOnPlayfieldSoundLevel, aBall
    Case 5 : PlaySoundAtLevelStatic ("Ball_Drop_Playfield_5_Delayed"), DelayedBallDropOnPlayfieldSoundLevel, aBall
  End Select
End Sub

'/////////////////////////////  BALL GATES AND BRACKET GATES SOUNDS  ////////////////////////////

Sub SoundPlayfieldGate()
  PlaySoundAtLevelStatic ("Gate_FastTrigger_" & Int(Rnd*2)+1), GateSoundLevel, Activeball
End Sub

Sub SoundHeavyGate()
  PlaySoundAtLevelStatic ("Gate_2"), GateSoundLevel, Activeball
End Sub

Sub Gates_hit(idx)
  SoundHeavyGate
End Sub

Sub GatesWire_hit(idx)
  SoundPlayfieldGate
End Sub

'/////////////////////////////  LEFT LANE ENTRANCE - SOUNDS  ////////////////////////////

Sub RandomSoundLeftArch()
  PlaySoundAtLevelActiveBall ("Arch_L" & Int(Rnd*4)+1), Vol(ActiveBall) * ArchSoundFactor
End Sub

Sub RandomSoundRightArch()
  PlaySoundAtLevelActiveBall ("Arch_R" & Int(Rnd*4)+1), Vol(ActiveBall) * ArchSoundFactor
End Sub


Sub Arch1_hit()
  If Activeball.velx > 1 Then SoundPlayfieldGate
  StopSound "Arch_L1"
  StopSound "Arch_L2"
  StopSound "Arch_L3"
  StopSound "Arch_L4"
End Sub

Sub Arch1_unhit()
  If activeball.velx < -8 Then
    RandomSoundRightArch
  End If
End Sub

Sub Arch2_hit()
  If Activeball.velx < 1 Then SoundPlayfieldGate
  StopSound "Arch_R1"
  StopSound "Arch_R2"
  StopSound "Arch_R3"
  StopSound "Arch_R4"
End Sub

Sub Arch2_unhit()
  If activeball.velx > 10 Then
    RandomSoundLeftArch
  End If
End Sub

'/////////////////////////////  SAUCERS (KICKER HOLES)  ////////////////////////////

Sub SoundSaucerLock()
  PlaySoundAtLevelStatic ("Saucer_Enter_" & Int(Rnd*2)+1), SaucerLockSoundLevel, Activeball
End Sub

Sub SoundSaucerKick(scenario, saucer)
  Select Case scenario
    Case 0: PlaySoundAtLevelStatic SoundFX("Saucer_Empty", DOFContactors), SaucerKickSoundLevel, saucer
    Case 1: PlaySoundAtLevelStatic SoundFX("Saucer_Kick", DOFContactors), SaucerKickSoundLevel, saucer
  End Select
End Sub

'/////////////////////////////  BALL COLLISION SOUND  ////////////////////////////
Sub OnBallBallCollision(ball1, ball2, velocity)
  Dim snd
  Select Case Int(Rnd*7)+1
    Case 1 : snd = "Ball_Collide_1"
    Case 2 : snd = "Ball_Collide_2"
    Case 3 : snd = "Ball_Collide_3"
    Case 4 : snd = "Ball_Collide_4"
    Case 5 : snd = "Ball_Collide_5"
    Case 6 : snd = "Ball_Collide_6"
    Case 7 : snd = "Ball_Collide_7"
  End Select

  PlaySound (snd), 0, Csng(velocity) ^2 / 200 * BallWithBallCollisionSoundFactor * VolumeDial, AudioPan(ball1), 0, Pitch(ball1), 0, 0, AudioFade(ball1)
End Sub


'/////////////////////////////////////////////////////////////////
'         End Mechanical Sounds
'/////////////////////////////////////////////////////////////////

'=====================================
'   Ramp Rolling SFX updates nf
'=====================================
'Ball tracking ramp SFX 1.0
' Usage:
'- Setup hit events with WireRampOn True or WireRampOn False (True = Plastic ramp, False = Wire Ramp)
'- To stop tracking ball, use WireRampoff
'-- Otherwise, the ball will auto remove if it's below 30 vp units

'Example, from Space Station:
'Sub RampSoundPlunge1_hit() : WireRampOn  False : End Sub           'Enter metal habitrail
'Sub RampSoundPlunge2_hit() : WireRampOff : WireRampOn True : End Sub     'Exit Habitrail, enter onto Mini PF
'Sub RampEntry_Hit() : If activeball.vely < -10 then WireRampOn True : End Sub  'Ramp enterance
dim RampMinLoops : RampMinLoops = 4
dim RampBalls(10,2)
'x,0 = ball x,1 = ID, 2 = Protection against ending early (minimum amount of updates)
'0,0 is boolean on/off, 0,1 unused for now
RampBalls(0,0) = False

dim RampType(10)  'Slapped together support for multiple ramp types... False = Wire Ramp, True = Plastic Ramp

Sub WireRampOn(input)  : Waddball ActiveBall, input : RampRollUpdate: End Sub
Sub WireRampOff() : WRemoveBall ActiveBall.ID : End Sub

Sub Waddball(input, RampInput)  'Add ball
    'Debug.Print "In Waddball() + add ball to loop array"
  dim x : for x = 1 to uBound(RampBalls)  'Check, don't add balls twice
    if RampBalls(x, 1) = input.id then
      if Not IsEmpty(RampBalls(x,1) ) then Exit Sub 'Frustating issue with BallId 0. Empty variable = 0
    End If
  Next

  For x = 1 to uBound(RampBalls)
    if IsEmpty(RampBalls(x, 1)) then
      Set RampBalls(x, 0) = input
      RampBalls(x, 1) = input.ID
      RampType(x) = RampInput
      RampBalls(x, 2) = 0
      'exit For
      RampBalls(0,0) = True
      RampRoll.Enabled = 1   'Turn on timer
      'RampRoll.Interval = RampRoll.Interval 'reset timer
      exit Sub
    End If
    if x = uBound(RampBalls) then   'debug
      Debug.print "WireRampOn error, ball queue is full: " & vbnewline & _
             RampBalls(0, 0) & vbnewline & _
             Typename(RampBalls(1, 0)) & " ID:" & RampBalls(1, 1) & "type:" & RampType(1) & vbnewline & _
             Typename(RampBalls(2, 0)) & " ID:" & RampBalls(2, 1) & "type:" & RampType(2) & vbnewline & _
             Typename(RampBalls(3, 0)) & " ID:" & RampBalls(3, 1) & "type:" & RampType(3) & vbnewline & _
             Typename(RampBalls(4, 0)) & " ID:" & RampBalls(4, 1) & "type:" & RampType(4) & vbnewline & _
             Typename(RampBalls(5, 0)) & " ID:" & RampBalls(5, 1) & "type:" & RampType(5) & vbnewline & _
             " "
    End If
  next
End Sub

Sub WRemoveBall(ID)   'Remove ball
    'Debug.Print "In WRemoveBall() + Remove ball from loop array"
  dim ballcount : ballcount = 0
  dim x : for x = 1 to Ubound(RampBalls)
    if ID = RampBalls(x, 1) then 'remove ball
      Set RampBalls(x, 0) = Nothing
      RampBalls(x, 1) = Empty
      RampType(x) = Empty
      StopSound("RampLoop" & x)
      StopSound("wireloop" & x)
    end If
    'if RampBalls(x,1) = Not IsEmpty(Rampballs(x,1) then ballcount = ballcount + 1
    if not IsEmpty(Rampballs(x,1)) then ballcount = ballcount + 1
  next
  if BallCount = 0 then RampBalls(0,0) = False  'if no balls in queue, disable timer update
End Sub

Sub RampRoll_Timer():RampRollUpdate:End Sub

Sub RampRollUpdate()    'Timer update
  dim x : for x = 1 to uBound(RampBalls)
    if Not IsEmpty(RampBalls(x,1) ) then
      if BallVel(RampBalls(x,0) ) > 1 then ' if ball is moving, play rolling sound
        If RampType(x) then
          PlaySound("RampLoop" & x), -1, VolPlayfieldRoll(RampBalls(x,0)) * 1.1 * VolumeDial, AudioPan(RampBalls(x,0)), 0, BallPitchV(RampBalls(x,0)), 1, 0, AudioFade(RampBalls(x,0))
          StopSound("wireloop" & x)
        Else
          StopSound("RampLoop" & x)
          PlaySound("wireloop" & x), -1, VolPlayfieldRoll(RampBalls(x,0)) * 1.1 * VolumeDial, AudioPan(RampBalls(x,0)), 0, BallPitchV(RampBalls(x,0)), 1, 0, AudioFade(RampBalls(x,0))
        End If
        RampBalls(x, 2) = RampBalls(x, 2) + 1
      Else
        StopSound("RampLoop" & x)
        StopSound("wireloop" & x)
      end if
      if RampBalls(x,0).Z < 30 and RampBalls(x, 2) > RampMinLoops then  'if ball is on the PF, remove  it
        StopSound("RampLoop" & x)
        StopSound("wireloop" & x)
        Wremoveball RampBalls(x,1)
      End If
    Else
      StopSound("RampLoop" & x)
      StopSound("wireloop" & x)
    end if
  next
  if not RampBalls(0,0) then RampRoll.enabled = 0

End Sub


Sub tbWR_Timer()  'debug textbox
  me.text = "on? " & RampBalls(0, 0) & " timer: " & RampRoll.Enabled & vbnewline & _
         "1 " & Typename(RampBalls(1, 0)) & " ID:" & RampBalls(1, 1) & " type:" & RampType(1) & " Loops:" & RampBalls(1, 2) & vbnewline & _
         "2 " & Typename(RampBalls(2, 0)) & " ID:" & RampBalls(2, 1) & " type:" & RampType(2) & " Loops:" & RampBalls(2, 2) & vbnewline & _
         "3 " & Typename(RampBalls(3, 0)) & " ID:" & RampBalls(3, 1) & " type:" & RampType(3) & " Loops:" & RampBalls(3, 2) & vbnewline & _
         "4 " & Typename(RampBalls(4, 0)) & " ID:" & RampBalls(4, 1) & " type:" & RampType(4) & " Loops:" & RampBalls(4, 2) & vbnewline & _
         "5 " & Typename(RampBalls(5, 0)) & " ID:" & RampBalls(5, 1) & " type:" & RampType(5) & " Loops:" & RampBalls(5, 2) & vbnewline & _
         "6 " & Typename(RampBalls(6, 0)) & " ID:" & RampBalls(6, 1) & " type:" & RampType(6) & " Loops:" & RampBalls(6, 2) & vbnewline & _
         " "
End Sub

'-- not needed anymore BallPitch only need BallPitchV
Function BallPitch(ball) ' Calculates the pitch of the sound based on the ball speed
    BallPitch = pSlope(BallVel(ball), 1, -1000, 60, 10000)
End Function

Function BallPitchV(ball) ' Calculates the pitch of the sound based on the ball speed Variation
    BallPitchV = pSlope(BallVel(ball), 1, -4000, 60, 7000)
End Function

' Thalamus : Exit in a clean and proper way
Sub Table1_exit
  Controller.Pause = False
  Controller.Stop
End Sub


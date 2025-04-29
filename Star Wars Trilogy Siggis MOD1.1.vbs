


' Star Wars Trilogy / IPD No. 4054 / March 03, 1997 / 6 Players
' http://www.ipdb.org/machine.cgi?id=4054
' VP915 v1.0 by JPSalas 2013
' Thanks to Magnox, Subzero and wtiger for the old table

Option Explicit
Randomize

On Error Resume Next
ExecuteGlobal GetTextFile("controller.vbs")
If Err Then MsgBox "You need the controller.vbs in order to run this table, available in the vp10 package"
On Error Goto 0

Const BallSize = 50

Dim DesktopMode:DesktopMode = Table1.ShowDT
Dim UseVPMColoredDMD
if DesktopMode = False then ramp1.visible = 0:ramp2.visible = 0

UseVPMColoredDMD = DesktopMode

LoadVPM "01000000", "SEGA.VBS", 3.10

Dim bsTrough, bsTopVUK, bsBottomVUK, dtRight, plungerIM, bsXWingCannon, mXWing, mHole
Dim x, bump1, bump2, bump3, bump4




Const UseSolenoids = 2
Const UseLamps = 0
Const UseGI = 0
Const UseSync = 0 'set it to 1 if the table runs too fast
Const HandleMech = 0

' Standard Sounds
Const SSolenoidOn = "fx_Solenoidon"
Const SSolenoidOff = "fx_SolenoidOff"
Const SCoin = "fx_coin"

'************
' Table init.
'************






Sub table1_Init

' Thalamus : Was missing 'vpminit me'
vpminit me

  Dim cGameName


    With Controller
        cGameName = "swtril43" ' Latest rom version
        .GameName = cGameName
        .SplashInfoLine = "Star Wars Trilogy - Sega 1997" & vbNewLine & "VP9 table by JPSalas / Convert by Hanibal"
        .Games(cGameName).Settings.Value("rol") = 0 'dmd rotated
        .HandleMechanics = 0
        .HandleKeyboard = 0
        .ShowDMDOnly = 1
        .ShowFrame = 0
        .ShowTitle = 0
        .Hidden = 0
        If Err Then MsgBox Err.Description
    End With
    On Error Goto 0
    Controller.SolMask(0) = 0
    vpmTimer.AddTimer 2000, "Controller.SolMask(0)=&Hffffffff'" 'ignore all solenoids - then add the Timer to renable all the solenoids after 2 seconds
    Controller.Run GetPlayerHWnd

    ' Nudging
    vpmNudge.TiltSwitch = 56
    vpmNudge.Sensitivity = 1
    vpmNudge.TiltObj = Array(Bumper1, Bumper2, Bumper3, Bumper4, LeftSlingshot, RightSlingshot)

    ' Trough & Ball Release
    Set bsTrough = New cvpmBallStack
    With bsTrough
        .InitSw 0, 15, 14, 13, 12, 0, 0, 0
        .InitKick BallRelease, 110, 5
        .InitExitSnd SoundFX("fx_Ballrel",DOFContactors), SoundFX("fx_solenoid",DOFContactors)
        .IsTrough = True
        .Balls = 4
    End With

    ' Impulse Plunger
    Const IMPowerSetting = 42 'Plunger Power
    Const IMTime = 0.6        ' Time in seconds for Full Plunge
    Set plungerIM = New cvpmImpulseP
    With plungerIM
        .InitImpulseP swplunger, IMPowerSetting, IMTime
        .Random 0.3
        .switch 16
        .InitExitSnd SoundFX("fx_plunger2",DOFContactors), SoundFX("fx_plunger",DOFContactors)
        .CreateEvents "plungerIM"
    End With

    ' Top Vuk
    Set bsTopVUK = New cvpmBallStack
    With bsTopVUK
        .InitSw 0, 45, 0, 0, 0, 0, 0, 0
        .InitKick sw45a, 30, 10
        .InitExitSnd SoundFX("fx_popper",DOFContactors), SoundFX("fx_solenoid",DOFContactors)
        .Balls = 0
        .KickForceVar = 3
        .KickAngleVar = 3
    End With

    ' Bottom Vuk
    Set bsBottomVUK = New cvpmBallStack
    With bsBottomVUK
        .InitSw 0, 46, 0, 0, 0, 0, 0, 0
        .InitKick BottomVUK, 190, 0
        .InitExitSnd SoundFX("fx_popper",DOFContactors), SoundFX("fx_solenoid",DOFContactors)
        .Balls = 0
        .KickForceVar = 3
        .KickAngleVar = 3
    End With

    ' Right Drop targets
    Set dtRight = New cvpmDropTarget
    With dtRight
        .InitDrop Array(sw17, sw18, sw19, sw20), Array(17, 18, 19, 20)
        .InitSnd "fx_droptarget", SoundFX("fx_resetdrop",DOFContactors)
        .CreateEvents "dtRight"
    End With

    'XWing mech
    Set mXWing = New cvpmMech
    With mXWing
        .MType = vpmMechOneSol + vpmMechReverse + vpmMechNonLinear
        .Sol1 = 18 'motor relay
        .Length = 260
        .Steps = 260
        .Acc = 10
        .Ret = 1
        .AddSw 35, 0, 0
        .AddSw 36, 0, 80
        .Callback = GetRef("UpdateXWing")
        .Start
    End With
    Controller.Switch(35) = 1
    Controller.Switch(36) = 1

    ' Thalamus - more randomness to kickers pls
    ' X-Wing cannon
    Set bsXWingCannon = New cvpmBallStack
    With bsXWingCannon
        .KickForceVar = 3
        .KickAngleVar = 3
        .InitSw 0, 37, 0, 0, 0, 0, 0, 0
        .InitExitSnd SoundFX("fx_popper",DOFContactors), SoundFX("fx_solenoid",DOFContactors)
        .Balls = 0
    End With

    'Low powered Magnet to simulate drop in playfield surface around the Hans Solo hole
    Set mHole = New cvpmMagnet
    With mHole
        .initMagnet MagTrigger, 4.5
        .X = 868
        .Y = 1090
        .Size = 65
        .GrabCenter = 0
        .MagnetOn = 1
        .CreateEvents "mHole"
    End With

    ' Misc. Initialisation
    LeftSLing.IsDropped = 1:LeftSLing2.IsDropped = 1:LeftSLing3.IsDropped = 1
    RightSLing.IsDropped = 1:RightSLing2.IsDropped = 1:RightSLing3.IsDropped = 1
    LeftSLingH.IsDropped = 1:LeftSLingH2.IsDropped = 1:LeftSLingH3.IsDropped = 1
    RightSLingH.IsDropped = 1:RightSLingH2.IsDropped = 1:RightSLingH3.IsDropped = 1


Can1.blenddisablelighting = 30
Can2.blenddisablelighting = 30
Can3.blenddisablelighting = 30
Can4.blenddisablelighting = 30


    ' Main Timer init
    PinMAMETimer.Interval = PinMAMEInterval
    PinMAMETimer.Enabled = 1
    'StartShake
Dim TempSoundTrigger1

End Sub
Sub table1_Paused:Controller.Pause = 1:End Sub
Sub table1_unPaused:Controller.Pause = 0:End Sub



Sub Duese_Timer


Duese1.Opacity = 200 + (RND*1000)
Duese2.Opacity = 200 + (RND*1000)

End Sub




 Sub Augen_Timer
Dim DesktopMode:DesktopMode = Table1.ShowDT
' ******Hanibals Random Lights Script
Lanelight.Intensity = (30+(30*Rnd))
FlashHol1.State =0
 End Sub

 Sub Laser_Timer
can1.y=can1.y+2
can1.x=can1.x+2
can2.y=can2.y+2
can2.x=can2.x+2
If can1.y > 900 Then
can1.y=370
can1.x=418
can2.y=440
can2.x=347
Laser.enabled = False
End If
 End Sub


'**********
' Keys
'**********
DIM PlungerStatus

Sub table1_KeyDown(ByVal Keycode)
    If keycode = PlungerKey Then Controller.Switch(53) = 1: PlungerStatus = 1
    If keycode = LockbarKey Then Controller.Switch(53) = 1: PlungerStatus = 1
    If keycode = LeftTiltKey Then LeftNudge 90, 1.6, 20:PlaySound "fx_nudge_left"
    If keycode = RightTiltKey Then RightNudge 270, 1.6, 20:PlaySound "fx_nudge_right"
    If keycode = CenterTiltKey Then CenterNudge 0, 2.8, 30:PlaySound "fx_nudge_forward"
    'If keycode = 7 Then DMDLocal
    If vpmKeyDown(keycode) Then Exit Sub
End Sub

Sub table1_KeyUp(ByVal Keycode)
    If keycode = PlungerKey Then Controller.Switch(53) = 0
    If keycode = LockbarKey Then Controller.Switch(53) = 0
    If vpmKeyUp(keycode) Then Exit Sub
End Sub


'*********************
' Extra Features
'*********************

'Sub DMDLocal


'If DMDLocalScreen.Visible = 0 THEN
'DMDLocalScreen.Visible = 1
'Else
'DMDLocalScreen.Visible = 0
'End If

'End Sub

'*************************************
'          Nudge System
' based on Noah's nudgetest table
'*************************************

Dim LeftNudgeEffect, RightNudgeEffect, NudgeEffect

Sub LeftNudge(angle, strength, delay)
    vpmNudge.DoNudge angle, (strength * (delay-LeftNudgeEffect) / delay) + RightNudgeEffect / delay
    LeftNudgeEffect = delay
    RightNudgeEffect = 0
    RightNudgeTimer.Enabled = 0
    LeftNudgeTimer.Interval = delay
    LeftNudgeTimer.Enabled = 1
End Sub

Sub RightNudge(angle, strength, delay)
    vpmNudge.DoNudge angle, (strength * (delay-RightNudgeEffect) / delay) + LeftNudgeEffect / delay
    RightNudgeEffect = delay
    LeftNudgeEffect = 0
    LeftNudgeTimer.Enabled = 0
    RightNudgeTimer.Interval = delay
    RightNudgeTimer.Enabled = 1
End Sub

Sub CenterNudge(angle, strength, delay)
    vpmNudge.DoNudge angle, strength * (delay-NudgeEffect) / delay
    NudgeEffect = delay
    NudgeTimer.Interval = delay
    NudgeTimer.Enabled = 1
End Sub

Sub LeftNudgeTimer_Timer()
    LeftNudgeEffect = LeftNudgeEffect-1
    If LeftNudgeEffect = 0 then LeftNudgeTimer.Enabled = 0
End Sub

Sub RightNudgeTimer_Timer()
    RightNudgeEffect = RightNudgeEffect-1
    If RightNudgeEffect = 0 then RightNudgeTimer.Enabled = 0
End Sub

Sub NudgeTimer_Timer()
    NudgeEffect = NudgeEffect-1
    If NudgeEffect = 0 then NudgeTimer.Enabled = False
End Sub

'**********
' Solenoids
'**********

SolCallback(1) = "bsTrough.SolOut"
SolCallback(2) = "AutoPlunger"
SolCallBack(3) = "dtRight.SolDropUp"
SolCallBack(4) = "SolTopVuk"
SolCallBack(5) = "FireXWing"
SolCallBack(6) = "bsBottomVUK.SolOut"
SolCallBack(7) = "RampMagnet"
'SolCallBack(8) =  "Relaissound"
'SolCallBack(9) = "Relaissound"
'SolCallBack(10) = "Relaissound"
'SolCallBack(11) = "Relaissound"
'SolCallBack(12)  'left slingshot
'SolCallBack(13)  'right slingshot
'SolCallBack(14) = "Relaissound"
'SolCallBack(24) = "Relaissound"
SolCallBack(17) = "SolDiv"
SolCallBack(18) = "XWingMotor"
SolCallBack(20) = "dtRight.SolHit 1,"
SolCallBack(21) = "dtRight.SolHit 2,"
SolCallBack(22) = "dtRight.SolHit 3,"
SolCallBack(23) = "dtRight.SolHit 4,"
SolCallBack(25) = "MagnetSlide"
SolCallBack(26) = "TieFighterShaker"
'SolCallBack(27) = "Relaissound"
'SolCallBack(28) = "Relaissound"
'SolCallBack(29) = "Relaissound"
'SolCallBack(31) = "Relaissound"
'SolCallBack(30) = "Relaissound"
'SolCallBack(32) = "Relaissound"

set GICallback = GetRef("UpdateGI")
Sub UpdateGI(no, Enabled)
  If Enabled Then
    dim xx
    For each xx in GI:xx.State = 1: Next
        For each xx in Bulbs:xx.blenddisablelighting = 2: Next
        For each xx in Spots:xx.blenddisablelighting = 30:  Next
        PlaySound "fx_relay"

  Else For each xx in GI:xx.State = 0: Next
    For each xx in Bulbs:xx.blenddisablelighting = 0:   Next
    For each xx in Spots:xx.blenddisablelighting = 0: Next
        PlaySound "fx_relay"
  End If
End Sub

Sub AutoPlunger(Enabled)
    If Enabled Then
        PlungerIM.AutoFire
    End If
End Sub

Sub SolTopVuk(Enabled)
    If Enabled AND bsTopVUK.Balls > 0 Then
        sw45.Destroyball
        bsTopVUK.ExitSol_On
    End If
End Sub

Sub Soldiv(Enabled)
    If Enabled Then
        Diverter.RotateToEnd
        PlaySoundAtVol SoundFX("fx_diverter",DOFContactors), Diverter, 1
    Else
        Diverter.RotateToStart
        PlaySoundAtVol SoundFX("fx_diverter",DOFContactors), Diverter, 1
    End If
End Sub

' Tie Fighter Shaker by Hanibal

Dim tiepos, tiepostemp, tiespeed, tiesound

Sub TieFighterShaker(Enabled)
    If Enabled Then
        tiepostemp = 10
        tiespeed = 10

        Tiefightershake.Enabled = True
'       SetFlash 109, abs(Enabled)
    Else
    tiesound = False
    End If
End Sub

Sub Tiefightershake_timer
  tiefighter1.RotAndTra2 = tiepos
  If tiepostemp <0.1 AND tiepostemp >-0.1  Then : Tiefightershake.Enabled = False: Exit Sub
    If tiesound = False Then PlaySoundAtVol SoundFX("solenoid", DOFContactors), tiefighter1, 1: tiesound = True

If tiepostemp < 0 Then
'    tiepos = ABS(tiepos) - (0.01+(60/Tiefightershake.interval))
  IF tiepos > tiepostemp Then
  tiepos = tiepos - (0.1*(tiespeed/2))
    Else
  tiepostemp = ABS(tiepostemp) - 0.03*tiespeed
    tiespeed = tiepostemp
    End if
Else
  IF tiepos < tiepostemp Then
  tiepos = tiepos + (0.1*(tiespeed/2))
    Else
  tiepostemp = -tiepostemp + 0.03*tiespeed
    End if

'     tiepos = -tiepos + (0.01+(60/Tiefightershake.interval))
End If
End Sub

'X-Wing

Dim cannonpos

Sub MagnetSlide(Enabled)
    If Enabled Then
    cannonpos = 0:
        Cannonanimate.Enabled = True
        PlungerStatus = 0
        PlaySoundAtVol SoundFX("solenoid",DOFContactors), cannonBase, 1
    End If
End Sub

Sub Cannonanimate_Timer



    Select case cannonpos
        Case 0:Laser.enabled = True: cannonMuzzle.Y=cannonMuzzle.Y-15:cannonMuzzle.X=cannonMuzzle.X-15:cannonMuzzle.blenddisablelighting = 2
        Case 1:cannonMuzzle.Y=cannonMuzzle.Y-9:cannonMuzzle.X=cannonMuzzle.X-9:cannonBase.Y=cannonBase.Y-1:cannonBase.X=cannonBase.X-1
        Case 2:cannonMuzzle.Y=cannonMuzzle.Y-5:cannonMuzzle.X=cannonMuzzle.X-5:cannonBase.Y=cannonBase.Y-3:cannonBase.X=cannonBase.X-3
        Case 3:cannonMuzzle.Y=cannonMuzzle.Y-3:cannonMuzzle.X=cannonMuzzle.X-3:cannonBase.Y=cannonBase.Y-1:cannonBase.X=cannonBase.X-1
        Case 4:cannonMuzzle.Y=cannonMuzzle.Y+3:cannonMuzzle.X=cannonMuzzle.X+3:cannonBase.Y=cannonBase.Y+1:cannonBase.X=cannonBase.X+1
        Case 5:cannonMuzzle.Y=cannonMuzzle.Y+5:cannonMuzzle.X=cannonMuzzle.X+5:cannonBase.Y=cannonBase.Y+3:cannonBase.X=cannonBase.X+3
        Case 6:cannonMuzzle.Y=cannonMuzzle.Y+9:cannonMuzzle.X=cannonMuzzle.X+9:cannonBase.Y=cannonBase.Y+1:cannonBase.X=cannonBase.X+1
        Case 7:cannonMuzzle.Y=cannonMuzzle.Y+15:cannonMuzzle.X=cannonMuzzle.X+15
        Case 8:Cannonanimate.Enabled = False:cannonMuzzle.blenddisablelighting = 0.1
    End Select

    cannonpos = cannonpos + 1
  '  canonR.State = ABS(canonR.State -1)
End Sub

' Update X-Wing

Dim KickDir



Sub UpdateXWing(aNewPos, aSpeed, aLastPos)


    KickDir = aNewPos / 4
    XwingCannon.RotAndTra2 = KickDir
    xwingcannon1.RotAndTra2 = KickDir
    xwing1.RotAndTra2 = KickDir

  Duese1.Rotz = KickDir +270
  Duese2.Rotz = KickDir +270



      IF PlungerStatus = 1 AND Xwingshoot = 1 Then '
        bsXWingCannon.InitKick sw37, KickDir -8, 32
        bsXwingCannon.ExitSol_On
        Xwingshoot = 0
      End If

    IF KickDir = 0 Then

        Duese1.visible = False
    Duese2.visible = False

    End If


    IF KickDir > 1 Then
        Duese1.visible = True
    Duese2.visible = True
    PlaySoundAtVol "Gunmotor", xwing1, 1
  END if

    IF KickDir = 1 Then

  END if

End Sub




Sub XWingMotor(Enabled)
End Sub



Sub FireXwing(Enabled)
    If Enabled AND Xwingshoot = 1 Then
        bsXWingCannon.InitKick sw37, KickDir -8, 32
        bsXwingCannon.ExitSol_On
        Xwingshoot = 0
    End if
End Sub

' Cannon Magnet

Sub RampMagnet(Enabled)
    If Enabled Then
        Magnet.Enabled = True
    Else
        Magnet.Enabled = False
    End If
End Sub

Sub Magnet_Hit()
    ActiveBall.X = 182
  ActiveBall.Y = 180
  ActiveBall.Z = 125
  ActiveBall.VelX = 0
  ActiveBall.VelY = 0
End Sub

'*********
' Switches
'*********

' Slings
Dim LStep, RStep

Sub LeftSlingShot_Slingshot:LeftSling.IsDropped = 0:PlaySoundAtVol SoundFX("fx_Sling_L1",DOFContactors), ActiveBall, 1:vpmTimer.PulseSw 59:LStep = 0:Me.TimerEnabled = 1:End Sub

Sub LeftSlingShot_Timer
    Select Case LStep
        Case 0:LeftSLing.IsDropped = 0:LeftSLingH.IsDropped = 0
        Case 1: 'pause
        Case 2:LeftSLing.IsDropped = 1:LeftSLing2.IsDropped = 0:LeftSLingH.IsDropped = 1:LeftSLingH2.IsDropped = 0
        Case 3:LeftSLing2.IsDropped = 1:LeftSLing3.IsDropped = 0:LeftSLingH2.IsDropped = 1:LeftSLingH3.IsDropped = 0
        Case 4:LeftSLing3.IsDropped = 1:LeftSLingH3.IsDropped = 1:Me.TimerEnabled = 0
    End Select

    LStep = LStep + 1
End Sub

Sub RightSlingShot_Slingshot:RightSling.IsDropped = 0:PlaySoundAtVol SoundFX("fx_Sling_R1",DOFContactors), ActiveBall, 1:vpmTimer.PulseSw 62:RStep = 0:Me.TimerEnabled = 1:End Sub
Sub RightSlingShot_Timer
    Select Case RStep
        Case 0:RightSLing.IsDropped = 0:RightSLingH.IsDropped = 0
        Case 1: 'pause
        Case 2:RightSLing.IsDropped = 1:RightSLing2.IsDropped = 0:RightSLingH.IsDropped = 1:RightSLingH2.IsDropped = 0
        Case 3:RightSLing2.IsDropped = 1:RightSLing3.IsDropped = 0:RightSLingH2.IsDropped = 1:RightSLingH3.IsDropped = 0
        Case 4:RightSLing3.IsDropped = 1:RightSLingH3.IsDropped = 1:Me.TimerEnabled = 0
    End Select

    RStep = RStep + 1
End Sub

' Bumpers
Sub Bumper1_Hit:vpmTimer.PulseSw 49:PlaySoundAtVol SoundFX( "fx_Bumpers_1",DOFContactors), ActiveBall, 1:B6.Duration 1, 150, 0: LED1.Duration 1, 150, 0:End Sub
Sub Bumper2_Hit:vpmTimer.PulseSw 50:PlaySoundAtVol SoundFX( "fx_Bumpers_2",DOFContactors), ActiveBall, 1:B5.Duration 1, 150, 0: LED2.Duration 1, 150, 0:End Sub
Sub Bumper3_Hit:vpmTimer.PulseSw 51:PlaySoundAtVol SoundFX( "fx_Bumpers_3",DOFContactors), ActiveBall, 1:B3.Duration 1, 150, 0: LED3.Duration 1, 150, 0:End Sub
Sub Bumper4_Hit:vpmTimer.PulseSw 52:PlaySoundAtVol SoundFX( "fx_Bumpers_4",DOFContactors), ActiveBall, 1:B4.Duration 1, 150, 0: LED4.Duration 1, 150, 0:End Sub


' Eject holes
Sub Drain_Hit:PlaySoundAtVol "fx_drain", Drain, 1:bsTrough.AddBall Me:End Sub

Sub sw21_hit:sw21.destroyball:FlashHolTie:PlaySoundAtVol "fx_hole_enter", sw21, 1:vpmTimer.PulseSwitch 21, 2500, "bsBottomVuk.AddBall":End Sub
Sub sw21a_hit:sw21a.destroyball:FlashHolTie:PlaySoundAtVol "fx_hole_enter", sw21a, 1:vpmTimer.PulseSwitch 21, 2500, "bsBottomVuk.AddBall":End Sub
Sub sw21b_hit:sw21b.destroyball:FlashHolTie:PlaySoundAtVol "fx_hole_enter", sw21b, 1:vpmTimer.PulseSwitch 21, 2500, "bsBottomVuk.AddBall":End Sub
Sub sw21c_hit:sw21c.destroyball:FlashHolTie:PlaySoundAtVol "fx_hole_enter", sw21c, 1:vpmTimer.PulseSwitch 21, 2500, "bsBottomVuk.AddBall":End Sub
Sub sw21d_hit:sw21d.destroyball:FlashHolTie:PlaySoundAtVol "fx_hole_enter", sw21d, 1:vpmTimer.PulseSwitch 21, 2500, "bsBottomVuk.AddBall":End Sub

Sub sw46_hit:PlaySoundAtVol "fx_hole_enter", ActiveBall, 1:bsBottomVuk.AddBall Me:End Sub

Sub sw45_hit:PlaySoundAtVol "fx_kicker_enter", ActiveBall, 1:bsTopVuk.AddBall 0:End Sub

Sub sw37_Hit:PlaySoundAtVol "fx_kicker_enter", ActiveBall, 1:Xwingshoot = 1:bsXWingCannon.AddBall Me:saver.enabled=1:End Sub

Sub TriggerSound1_Hit: PlaySoundAtVol "fx_ballhit", ActiveBall, 1: End Sub

Sub FlashHolTie
FlashHol1.State =1


End Sub


' Rollovers
Sub sw57_Hit:Controller.Switch(57) = 1:PlaySoundAtVol "fx_sensor", ActiveBall, 1:End Sub
Sub sw57_UnHit:Controller.Switch(57) = 0:End Sub

Sub sw58_Hit:Controller.Switch(58) = 1:PlaySoundAtVol "fx_sensor", ActiveBall, 1:End Sub
Sub sw58_UnHit:Controller.Switch(58) = 0:End Sub

Sub sw61_Hit:Controller.Switch(61) = 1:PlaySoundAtVol "fx_sensor", ActiveBall, 1:End Sub
Sub sw61_UnHit:Controller.Switch(61) = 0:End Sub

Sub sw60_Hit:Controller.Switch(60) = 1:PlaySoundAtVol "fx_sensor", ActiveBall, 1:End Sub
Sub sw60_UnHit:Controller.Switch(60) = 0:End Sub

Sub sw43_Hit:Controller.Switch(43) = 1:PlaySoundAtVol "fx_sensor", ActiveBall, 1:End Sub
Sub sw43_UnHit:Controller.Switch(43) = 0:End Sub

Sub sw44_Hit:Controller.Switch(44) = 1:PlaySoundAtVol "fx_sensor", ActiveBall, 1:End Sub
Sub sw44_UnHit:Controller.Switch(44) = 0:End Sub

Sub sw34_Hit:Controller.Switch(34) = 1:PlaySoundAtVol "fx_sensor", ActiveBall, 1:End Sub
Sub sw34_UnHit:Controller.Switch(34) = 0:End Sub

Sub sw47_Hit:Controller.Switch(47) = 1:PlaySoundAtVol "fx_sensor", ActiveBall, 1:End Sub
Sub sw47_UnHit:Controller.Switch(47) = 0:End Sub

' targets

Sub sw38_Hit:vpmTimer.PulseSw 38:PlaySoundAtVol "fx_target", ActiveBall, 1:End Sub
Sub sw39_Hit:vpmTimer.PulseSw 39:PlaySoundAtVol "fx_target", ActiveBall, 1:End Sub
Sub sw40_Hit:vpmTimer.PulseSw 40:PlaySoundAtVol "fx_target", ActiveBall, 1:End Sub

Sub sw28_Hit:vpmTimer.PulseSw 28:PlaySoundAtVol SoundFX("fx_target", DOFTargets), Activeball, 1:End Sub
Sub sw29_Hit:vpmTimer.PulseSw 29:PlaySoundAtVol SoundFX("fx_target", DOFTargets), Activeball, 1:End Sub
Sub sw30_Hit:vpmTimer.PulseSw 30:PlaySoundAtVol SoundFX("fx_target", DOFTargets), ActiveBall, 1:End Sub
Sub sw31_Hit:vpmTimer.PulseSw 31:PlaySoundAtVol SoundFX("fx_target", DOFTargets), ActiveBall, 1:End Sub
Sub sw32_Hit:vpmTimer.PulseSw 32:PlaySoundAtVol SoundFX("fx_target", DOFTargets), ActiveBall, 1:End Sub

' Droptargets

Sub sw17_Hit:dtRight.Hit 1:End Sub
Sub sw18_Hit:dtRight.Hit 2:End Sub
Sub sw19_Hit:dtRight.Hit 3:End Sub
Sub sw20_Hit:dtRight.Hit 4:End Sub

' Ramp switches

Sub sw24_Hit:Controller.Switch(24) = 1:PlaySoundAtVol "fx_sensor", ActiveBall, 1:saver.enabled=0:End Sub
Sub sw24_UnHit:Controller.Switch(24) = 0:End Sub

Sub sw22_Hit:Controller.Switch(22) = 1:PlaySoundAtVol "fx_sensor", ActiveBall, 1:End Sub
Sub sw22_UnHit:Controller.Switch(22) = 0:End Sub

Sub sw23_Hit:Controller.Switch(23) = 1:PlaySoundAtVol "fx_sensor", ActiveBall, 1:End Sub
Sub sw23_UnHit:Controller.Switch(23) = 0:End Sub

Sub LaneTrigger_Hit: Lanelight.State = 1 :End Sub
Sub LaneTrigger_UnHit: Lanelight.State = 0 :End Sub


'********************
' Special JP Flippers
'********************

SolCallback(sLRFlipper) = "SolRFlipper"
SolCallback(sLLFlipper) = "SolLFlipper"

Sub SolLFlipper(Enabled)
    Dim tmp, tmp2
    If Enabled Then
        PlaySoundAtVol SoundFX( "Flipper_L01",DOFFlippers), LeftFlipper, 1:LeftFlipper.RotateToEnd
    Else
        tmp = LeftFlipper.Strength
'        tmp2 = LeftFlipper.Recoil
        LeftFlipper.Strength = 6 'increase return strength to compensate for the slower speed
'        LeftFlipper.Recoil = 0.5 'decrease the recoil to allow drop catches
        PlaySoundAtVol SoundFX( "Flipper_Left",DOFFlippers), LeftFlipper, 1:LeftFlipper.RotateToStart
        LeftFlipper.Strength = tmp
'        LeftFlipper.Recoil = tmp2
    End If
End Sub

Sub SolRFlipper(Enabled)
    Dim tmp, tmp2
    If Enabled Then
        PlaySoundAtVol SoundFX( "Flipper_R01",DOFFlippers), RightFlipper, 1:RightFlipper.RotateToEnd
    Else
        tmp = RightFlipper.Strength
'        tmp2 = LeftFlipper.Recoil
        RightFlipper.Strength = 6 'increase return strength to compensate for the slower speed
'        RightFlipper.Recoil = 0.5 'decrease the recoil to allow drop catches
        PlaySoundAtVol SoundFX( "Flipper_Right",DOFFlippers), RightFlipper, 1:RightFlipper.RotateToStart
        RightFlipper.Strength = tmp
'        RightFlipper.Recoil = tmp2
    End If
End Sub

Sub LeftFlipper_Collide(parm)
    PlaySoundAtVol "fx_rubber_flipper", ActiveBall, 1
End Sub

Sub RightFlipper_Collide(parm)
    PlaySoundAtVol "fx_rubber_flipper", ActiveBall, 1
End Sub

'***********************
'     Flipper Logos
'***********************

Sub UpdateFlipperLogos
    LFLogo.RotAndTra2 = LeftFlipper.CurrentAngle
    RFlogo.RotAndTra2 = RightFlipper.CurrentAngle
End Sub

'******************
' RealTime Updates
'******************
Set MotorCallback = GetRef("GameTimer")

Sub GameTimer
    UpdateFlipperLogos
End Sub

' *******************************************************************************************************
' Positional Sound Playback Functions by DJRobX, Rothbauerw, Thalamus and Herweh
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

' set position as table object and Vol + RndPitch manually

Sub PlaySoundAtVolPitch(sound, tableobj, Vol, RndPitch)
  PlaySound sound, 1, Vol, AudioPan(tableobj), RndPitch, 0, 0, 1, AudioFade(tableobj)
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

Sub PlaySoundAtBallAbsVol(sound, VolMult)
  PlaySound sound, 0, VolMult, AudioPan(ActiveBall), 0, Pitch(ActiveBall), 0, 1, AudioFade(ActiveBall)
End Sub

' requires rampbump1 to 7 in Sound Manager

Sub RandomBump(voladj, freq)
  Dim BumpSnd:BumpSnd= "rampbump" & CStr(Int(Rnd*7)+1)
  PlaySound BumpSnd, 0, Vol(ActiveBall)*voladj, AudioPan(ActiveBall), 0, freq, 0, 1, AudioFade(ActiveBall)
End Sub

' set position as bumperX and Vol manually. Allows rapid repetition/overlaying sound

Sub PlaySoundAtBumperVol(sound, tableobj, Vol)
  PlaySound sound, 1, Vol, AudioPan(tableobj), 0,0,1, 1, AudioFade(tableobj)
End Sub

Sub PlaySoundAtBOTBallZ(sound, BOT)
  PlaySound sound, 0, ABS(BOT.velz)/17, Pan(BOT), 0, Pitch(BOT), 1, 0, AudioFade(BOT)
End Sub

' play a looping sound at a location with volume
Sub PlayLoopSoundAtVol(sound, tableobj, Vol)
  PlaySound sound, -1, Vol, AudioPan(tableobj), 0, 0, 1, 0, AudioFade(tableobj)
End Sub

'*********************************************************************
'            Supporting Ball, Sound Functions and Math
'*********************************************************************

Function RndNum(min, max)
  RndNum = Int(Rnd() * (max-min + 1) ) + min ' Sets a random number between min and max
End Function

Const Pi = 3.1415927

Function dSin(degrees)
  dsin = sin(degrees * Pi/180)
End Function

Function dCos(degrees)
  dcos = cos(degrees * Pi/180)
End Function

Function AudioFade(tableobj) ' Fades between front and back of the table (for surround systems or 2x2 speakers, etc), depending on the Y position on the table. "table1" is the name of the table
  Dim tmp
  On Error Resume Next
  tmp = tableobj.y * 2 / table1.height-1
  If tmp > 0 Then
    AudioFade = Csng(tmp ^10)
  Else
    AudioFade = Csng(-((- tmp) ^10) )
  End If
End Function

Function AudioPan(tableobj) ' Calculates the pan for a tableobj based on the X position on the table. "table1" is the name of the table
  Dim tmp
  On Error Resume Next
  tmp = tableobj.x * 2 / table1.width-1
  If tmp > 0 Then
    AudioPan = Csng(tmp ^10)
  Else
    AudioPan = Csng(-((- tmp) ^10) )
  End If
End Function

Function Pan(ball) ' Calculates the pan for a ball based on the X position on the table. "table1" is the name of the table
  Dim tmp
  On Error Resume Next
  tmp = ball.x * 2 / table1.width-1
  If tmp > 0 Then
    Pan = Csng(tmp ^10)
  Else
    Pan = Csng(-((- tmp) ^10) )
  End If
End Function

Function Vol(ball) ' Calculates the Volume of the sound based on the ball speed
  Vol = Csng(BallVel(ball) ^2 / 2000)
End Function

Function VolMulti(ball,Multiplier) ' Calculates the Volume of the sound based on the ball speed
  VolMulti = Csng(BallVel(ball) ^2 / 150 ) * Multiplier
End Function

Function DVolMulti(ball,Multiplier) ' Calculates the Volume of the sound based on the ball speed
  DVolMulti = Csng(BallVel(ball) ^2 / 150 ) * Multiplier
  debug.print DVolMulti
End Function

Function BallRollVol(ball) ' Calculates the Volume of the sound based on the ball speed
  BallRollVol = Csng(BallVel(ball) ^2 / (80000 - (79900 * Log(RollVol) / Log(100))))
End Function

Function Pitch(ball) ' Calculates the pitch of the sound based on the ball speed
  Pitch = BallVel(ball) * 20
End Function

Function BallVel(ball) 'Calculates the ball speed
  BallVel = INT(SQR((ball.VelX ^2) + (ball.VelY ^2) ) )
End Function

Function BallVelZ(ball) 'Calculates the ball speed in the -Z
  BallVelZ = INT((ball.VelZ) * -1 )
End Function

Function VolZ(ball) ' Calculates the Volume of the sound based on the ball speed in the Z
  VolZ = Csng(BallVelZ(ball) ^2 / 200)*1.2
End Function

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

'*****************************************
'      JP's VP10 Rolling Sounds
'*****************************************

Const tnob = 6 ' total number of balls
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
  PlaySound("fx_collide"), 0, Csng(velocity) ^2 / 2000, AudioPan(ball1), 0, Pitch(ball1), 0, 0, AudioFade(ball1)
End Sub


Sub BallShadowTimer_Timer()
Dim BallShadow
BallShadow = Array (BallShadow1,BallShadow2,BallShadow3,BallShadow4,BallShadow5,BallShadow6)
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
    BallShadow(b).X = BOT(b).X
    ballShadow(b).Y = BOT(b).Y + 10
        If BOT(b).Z > 20 and BOT(b).Z < 140 Then
            BallShadow(b).visible = 1
        Else
            BallShadow(b).visible = 0
        End If
if BOT(b).z > 30 Then
ballShadow(b).height = BOT(b).Z - 20
ballShadow(b).opacity = 110
Else
ballShadow(b).height = BOT(b).Z - 24
ballShadow(b).opacity = 90
End If
    Next
End Sub


'**********************************
'  JP's Fading Lamps v7.0 VP912
'   Based on PD's Fading Lights
' SetLamp 0 is Off
' SetLamp 1 is On
' LampState(x) = current state
'***********************************

Dim LampState(200), FadingLevel(200)
Dim FlashState(200), FlashLevel(200)
Dim FlashSpeedUp, FlashSpeedDown

Dim FlashMin(200), FlashMax(200)


'****** Flash Sound InitAddSnd
IF LampTimer.Enabled = 0 Then DIM FlashSoundA, FlashSoundB, FlashSoundC, Xwingshoot




AllLampsOff()
LampTimer.Interval = 25
LampTimer.Enabled = 1

FlashInit()
FlasherTimer.Interval = 10 'flash fading speed
FlasherTimer.Enabled = 1

Sub LampTimer_Timer()
    Dim chgLamp, num, chg, ii
    chgLamp = Controller.ChangedLamps
    If Not IsEmpty(chgLamp) Then
        For ii = 0 To UBound(chgLamp)
            LampState(chgLamp(ii, 0) ) = chgLamp(ii, 1) + 4
            FadingLevel(chgLamp(ii, 0) ) = chgLamp(ii, 1) + 4
      FlashState(chgLamp(ii, 0) ) = chgLamp(ii, 1)
        Next
    End If

    UpdateLamps
End Sub

Sub UpdateLamps
    LightHan 3, l3
    LightHan 4, l4
    LightHan 5, l5
    LightHan 6, l6
    LightHan 7, l7
    LightHan 8, l8
    LightHan 9, l9
'    LightHan 10, l10
    LightHan 11, l11
    LightHan 12, l12
    LightHan 13, l13
    LightHan 13, l13a
    LightHan 17, l17
    LightHan 18, l18
    LightHan 19, l19
    LightHan 20, l20
    LightHan 21, l21
    LightHan 22, l22
    LightHan 23, l23
    LightHan 24, l24
    LightHan 25, l25
    LightHan 26, l26
    LightHan 27, l27
    LightHan 28, l28
    LightHan 29, l29
    LightHan 30, l30
    LightHan 31, l31
    LightHan 32, l32
    LightHan 33, l33
    LightHan 34, l34
    LightHan 35, l35
    LightHan 36, l36
    LightHan 37, l37
    LightHan 38, l38
    LightHan 39, l39
    LightHan 39, l39a
    LightHan 40, l40
    LightHan 41, l41
    LightHan 42, l42
    LightHan 43, l43
    LightHan 44, l44
    LightHan 45, l45
    LightHan 46, l46
    LightHan 47, l47
    LightHan 48, l48


    FadeDisableLighting 49, HL4, 20
    FadeDisableLighting 50, HL3, 20
    FadeDisableLighting 51, HL2, 20
    FadeDisableLighting 52, HL1, 20
    FadeDisableLighting 53, HR4, 10
    FadeDisableLighting 54, HR3, 10
    FadeDisableLighting 55, HR2, 10
    FadeDisableLighting 56, HR1, 10

    LightHan 61, l61
    LightHan 62, l62
    LightHan 62, l62a

    LightHan 65, l65
    LightHan 66, l66
    LightHan 67, l67
    LightHan 68, l68
    LightHan 69, l69
    LightHan 70, l70
    LightHan 71, l71
    LightHan 71, l71a
    LightHan 72, l72
    LightHan 73, l73
    LightHan 74, l74
    LightHan 75, l75
    LightHan 76, l76
    LightHan 77, l77
    LightHan 77, l77a
    LightHan 78, l78
    LightHan 78, l78a
    LightHan 79, l79
    LightHan 80, l80
    LightHan 80, l80a

End Sub

Sub FlasherTimer_Timer()

'***************Flashers Hanibal Spezial*************

  FlashHan 3, F3a
    FlashHan 3, F3b
'    FlashHan 3, f3c
    FlashHan 3, FlashTie
  FlashHan 6, F6a
    FlashHan 6, F6b
    FlashHan 6, F6c
    FlashHan 6, F6d
'    FlashHan 6, f6e
'    FlashHan 6, f6f
    FlashHan 8, F8a
    FlashHan 8, F8b
    FlashHan 8, F8c
    FlashHan 8, F8d
'    FlashHan 8, f8e
'    FlashHan 8, f8f
 End Sub

Sub AllLampsOff()
    Dim x
    For x = 0 to 200
        LampState(x) = 4
    Next

UpdateLamps:UpdateLamps:Updatelamps
End Sub

Sub SetLamp(nr, value)
    If value <> LampState(nr) Then
    If value = 0 AND LampState(nr) = 0 Then Exit Sub
    If value = 1 AND LampState(nr) = 1 Then Exit Sub
    LampState(nr) = abs(value) + 4
    FadingLevel(nr) = abs(value) + 4
    End If
End Sub



'****************** Hanibal special flasher ****************************




Sub FlashHan(nr, object) ' used for multiple lights and pass playfield
    Select Case FadingLevel(nr)
        Case 4:object.state = 0 ': object.image = "lights"
        Case 5:object.state = 1 ': object.image = "lights_on"
    End Select

End Sub


'****************** Hanibal special lights ****************************



Sub LightHan(nr, object) ' used for multiple lights and pass playfield
    Select Case FadingLevel(nr)
        Case 4:object.state = 0 ': object.image = "lights"
        Case 5:object.state = 1 ': object.image = "lights_on"
    End Select
End Sub


' *****************div flasher subs****************


Sub FlashInit
    Dim i
    For i = 0 to 200
        FlashState(i) = 0
        FlashLevel(i) = 0
    Next

    FlashSpeedUp = 60   ' fast speed when turning on the flasher
    FlashSpeedDown = 25 ' slow speed when turning off the flasher, gives a smooth fading
    ' you could also change the default images for each flasher or leave it as in the editor
    ' for example
    ' Flasher1.Image = "fr"
    AllFlashOff()
End Sub

Sub AllFlashOff
    Dim i
    For i = 0 to 200
        FlashState(i) = 0
    Next
End Sub

Sub Flash(nr, object)
    Select Case FadingLevel(nr)
        Case 4 'off
            'FlashLevel(nr) = FlashLevel(nr)
            'If FlashLevel(nr) < FlashMin(nr) Then
            '    FlashLevel(nr) = FlashMin(nr)
            '    FadingLevel(nr) = 0 'completely off
            'End if
            'Object.IntensityScale = FlashLevel(nr)
      Object.visible = 0



        Case 5 ' on
            'FlashLevel(nr) = FlashLevel(nr)' + FlashSpeedUp(nr)
            'If FlashLevel(nr) > FlashMax(nr) Then
            '    FlashLevel(nr) = FlashMax(nr)
            '    FadingLevel(nr) = 1 'completely on
            'End if
            'Object.IntensityScale = FlashLevel(nr)
            Object.visible = 1
    End Select
End Sub

Sub FadeDisableLighting(nr, a, alvl)
  Select Case FadingLevel(nr)
    Case 4
      a.UserValue = a.UserValue - 0.3
      If a.UserValue < 0 Then
        a.UserValue = 0
        FadingLevel(nr) = 0
      end If
      a.BlendDisableLighting = alvl * a.UserValue 'brightness
    Case 5
      a.UserValue = a.UserValue + 0.50
      If a.UserValue > 1 Then
        a.UserValue = 1
        FadingLevel(nr) = 1
      end If
      a.BlendDisableLighting = alvl * a.UserValue 'brightness
  End Select
End Sub


'********************
' Diverse Help/Sounds
'********************

Sub AllRubbers_Hit(idx):PlaySound "fx_rubber", 0, (20 + Vol(ActiveBall)), (20+ AudioPan(ActiveBall)), Pitch(ActiveBall), 1, 1, AudioFade(ActiveBall):End Sub
Sub AllPostRubbers_Hit(idx):PlaySound "fx_rubber", 0, (20 + Vol(ActiveBall)), (20+ AudioPan(ActiveBall)), Pitch(ActiveBall), 1, 1, AudioFade(ActiveBall):End Sub
Sub AllMetals_Hit(idx):PlaySound "fx_MetalHit", 0, (20 + Vol(ActiveBall)), (20+ AudioPan(ActiveBall)), Pitch(ActiveBall), 1, 1, AudioFade(ActiveBall):End Sub
Sub AllGates_Hit(idx):PlaySound "fx_Gate", 0, (20 + Vol(ActiveBall)), (20+ AudioPan(ActiveBall)), Pitch(ActiveBall), 1, 1, AudioFade(ActiveBall):End Sub
Sub RHelp1_hit
    PlaySoundAtVol "fx_ballhit", ActiveBall, 1
End Sub

Sub RHelp2_hit
    PlaySoundAtVol "fx_ballhit", ActiveBall, 1
End Sub

' Thalamus : Exit in a clean and proper way
Sub Table1_exit
  Controller.Pause = False
  Controller.Stop
End Sub

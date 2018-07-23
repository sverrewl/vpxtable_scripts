


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

UseVPMColoredDMD = DesktopMode

LoadVPM "01000200", "SEGA.VBS", 3.02

Dim bsTrough, bsTopVUK, bsBottomVUK, dtRight, plungerIM, bsXWingCannon, mXWing, mHole
Dim x, bump1, bump2, bump3, bump4




Const UseSolenoids = 1
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
        .InitExitSnd "fx_Ballrel", "fx_solenoid"
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
        .InitExitSnd "fx_plunger2", "fx_plunger"
        .CreateEvents "plungerIM"
    End With

    ' Top Vuk
    Set bsTopVUK = New cvpmBallStack
    With bsTopVUK
        .InitSw 0, 45, 0, 0, 0, 0, 0, 0
        .InitKick sw45a, 30, 20
        .InitExitSnd "fx_popper", "fx_Solenoid"
        .Balls = 0
    End With

    ' Bottom Vuk
    Set bsBottomVUK = New cvpmBallStack
    With bsBottomVUK
        .InitSw 0, 46, 0, 0, 0, 0, 0, 0
        .InitKick BottomVUK, 190, 0
        .InitExitSnd "fx_popper", "fx_Solenoid"
        .Balls = 0
    End With

    ' Right Drop targets
    Set dtRight = New cvpmDropTarget
    With dtRight
        .InitDrop Array(sw17, sw18, sw19, sw20), Array(17, 18, 19, 20)
        .InitSnd "fx_droptarget", "fx_resetdrop"
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

    ' X-Wing cannon
    Set bsXWingCannon = New cvpmBallStack
    With bsXWingCannon
        .InitSw 0, 37, 0, 0, 0, 0, 0, 0
        .InitExitSnd "fx_popper", "fx_Solenoid"
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
    Ring1a.IsDropped = 1:Ring1b.IsDropped = 1:Ring1c.IsDropped = 1
    Ring2a.IsDropped = 1:Ring2b.IsDropped = 1:Ring2c.IsDropped = 1
    Ring3a.IsDropped = 1:Ring3b.IsDropped = 1:Ring3c.IsDropped = 1
    Ring4a.IsDropped = 1:Ring4b.IsDropped = 1:Ring4c.IsDropped = 1
    sw28a.IsDropped = 1:sw29a.IsDropped = 1:sw30a.IsDropped = 1:sw31a.IsDropped = 1
    sw32a.IsDropped = 1


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
z49.height = 200
z50.height = 170
z51.height = 140
z52.height = 110
z53.height = 200
z54.height = 170
z55.height = 140
z56.height = 110
'Han1.height = 200



 End Sub

 Sub Laser_Timer


Laser1.visible = True
Laser2.visible = True
Laser1.y=Laser1.y+2
Laser1.x=Laser1.x+2
Laser2.y=Laser2.y+2
Laser2.x=Laser2.x+2



If DesktopMode = True Then
'Laser1.y=Laser1.y+1.2
'Laser2.y=Laser2.y+1.2
Laser1.rotz =230
Laser2.rotz =230




End If

If Laser1.y > 2000 Then
Laser1.visible = False
Laser2.visible = False
Laser1.y=600
Laser1.x=480
Laser2.y=530
Laser2.x=564
Laser.enabled = False
End If

 End Sub


'**********
' Keys
'**********
DIM PlungerStatus

Sub table1_KeyDown(ByVal Keycode)
    If keycode = PlungerKey Then Controller.Switch(53) = 1: PlungerStatus = 1
    If keycode = LeftTiltKey Then LeftNudge 90, 1.6, 20:PlaySound "fx_nudge_left"
    If keycode = RightTiltKey Then RightNudge 270, 1.6, 20:PlaySound "fx_nudge_right"
    If keycode = CenterTiltKey Then CenterNudge 0, 2.8, 30:PlaySound "fx_nudge_forward"
    If keycode = 7 Then DMDLocal
    If vpmKeyDown(keycode) Then Exit Sub
End Sub

Sub table1_KeyUp(ByVal Keycode)
    If keycode = PlungerKey Then Controller.Switch(53) = 0
    If vpmKeyUp(keycode) Then Exit Sub
End Sub


'*********************
' Extra Features
'*********************

Sub DMDLocal


'If DMDLocalScreen.Visible = 0 THEN 
DMDLocalScreen.Visible = 1
'Else
'DMDLocalScreen.Visible = 0
'End If

End Sub

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
'SolCallBack(9)	= "Relaissound"
'SolCallBack(10) = "Relaissound"
'SolCallBack(11) = "Relaissound"
'SolCallBack(12)	'left slingshot
'SolCallBack(13)	'right slingshot
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
        playsound "fx_diverter"
    Else
        Diverter.RotateToStart
        playsound "fx_diverter"
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
    If tiesound = False Then Playsound "solenoid", 0, 1, 0.20, 1: tiesound = True

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

'	    tiepos = -tiepos + (0.01+(60/Tiefightershake.interval))
End If
End Sub

'X-Wing

Dim cannonpos

Sub MagnetSlide(Enabled)
    If Enabled Then
		cannonpos = 0:
        Cannonanimate.Enabled = True
        PlungerStatus = 0
        Playsound "solenoid", 0, 1, -0.20, 0.25
    End If
End Sub

Sub Cannonanimate_Timer



    Select case cannonpos
        Case 0:Laser.enabled = True: cannonMuzzle.Y=cannonMuzzle.Y-15:cannonMuzzle.X=cannonMuzzle.X-15
        Case 1:cannonMuzzle.Y=cannonMuzzle.Y-9:cannonMuzzle.X=cannonMuzzle.X-9:cannonBase.Y=cannonBase.Y-1:cannonBase.X=cannonBase.X-1 
        Case 2:cannonMuzzle.Y=cannonMuzzle.Y-5:cannonMuzzle.X=cannonMuzzle.X-5:cannonBase.Y=cannonBase.Y-3:cannonBase.X=cannonBase.X-3 
        Case 3:cannonMuzzle.Y=cannonMuzzle.Y-3:cannonMuzzle.X=cannonMuzzle.X-3:cannonBase.Y=cannonBase.Y-1:cannonBase.X=cannonBase.X-1 
        Case 4:cannonMuzzle.Y=cannonMuzzle.Y+3:cannonMuzzle.X=cannonMuzzle.X+3:cannonBase.Y=cannonBase.Y+1:cannonBase.X=cannonBase.X+1 
        Case 5:cannonMuzzle.Y=cannonMuzzle.Y+5:cannonMuzzle.X=cannonMuzzle.X+5:cannonBase.Y=cannonBase.Y+3:cannonBase.X=cannonBase.X+3
        Case 6:cannonMuzzle.Y=cannonMuzzle.Y+9:cannonMuzzle.X=cannonMuzzle.X+9:cannonBase.Y=cannonBase.Y+1:cannonBase.X=cannonBase.X+1
        Case 7:cannonMuzzle.Y=cannonMuzzle.Y+15:cannonMuzzle.X=cannonMuzzle.X+15
        Case 8:Cannonanimate.Enabled = False
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
        bsXWingCannon.InitKick sw37, KickDir, 32
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
		Playsound "Gunmotor", 0, 0.2, -0.20, 0.1
	END if

    IF KickDir = 1 Then

	END if

End Sub




Sub XWingMotor(Enabled)


  If Enabled Then 



  End If

End Sub



Sub FireXwing(Enabled)
    If Enabled AND Xwingshoot = 1 Then
        bsXWingCannon.InitKick sw37, KickDir, 32
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
	ActiveBall.Y = 175
	ActiveBall.Z = 120
	ActiveBall.VelX = -5
	ActiveBall.VelY = 5
End Sub

'*********
' Switches
'*********

' Slings
Dim LStep, RStep

Sub LeftSlingShot_Slingshot:LeftSling.IsDropped = 0:PlaySound "fx_slingshot1":vpmTimer.PulseSw 59:LStep = 0:Me.TimerEnabled = 1:End Sub

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

Sub RightSlingShot_Slingshot:RightSling.IsDropped = 0:PlaySound "fx_slingshot2":vpmTimer.PulseSw 62:RStep = 0:Me.TimerEnabled = 1:End Sub
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
Sub Bumper1_Hit:vpmTimer.PulseSw 49:PlaySound "bumper2", 0, 1, -0.15, 0.25:bump1 = 1:Me.TimerEnabled = 1:B6.State = 1 : LED1.State = 1:If DesktopMode = True Then LED1a.State = 1:END If:End Sub
Sub Bumper1_Timer()
    Select Case bump1
        Case 1:Ring1a.IsDropped = 0:bump1 = 2
        Case 2:Ring1b.IsDropped = 0:Ring1a.IsDropped = 1:bump1 = 3
        Case 3:Ring1c.IsDropped = 0:Ring1b.IsDropped = 1:bump1 = 4
        Case 4:Ring1c.IsDropped = 1:Me.TimerEnabled = 0: B6.State = 0: LED1.State = 0:LED1a.State = 0:
    End Select
End Sub

Sub Bumper2_Hit:vpmTimer.PulseSw 50:PlaySound "bumper2", 0, 1, -0.15, 0.25:bump2 = 1:Me.TimerEnabled = 1:B5.State = 1 : LED2.State = 1:If DesktopMode = True Then LED2a.State = 1:END If:End Sub
Sub Bumper2_Timer()
    Select Case bump2
        Case 1:Ring2a.IsDropped = 0:bump2 = 2
        Case 2:Ring2b.IsDropped = 0:Ring2a.IsDropped = 1:bump2 = 3
        Case 3:Ring2c.IsDropped = 0:Ring2b.IsDropped = 1:bump2 = 4
        Case 4:Ring2c.IsDropped = 1:Me.TimerEnabled = 0: B5.State = 0: LED2.State = 0: LED2a.State = 0:
    End Select
End Sub

Sub Bumper3_Hit:vpmTimer.PulseSw 51:PlaySound "bumper2", 0, 1, -0.15, 0.25:bump3 = 1:Me.TimerEnabled = 1:B3.State = 1 : LED3.State = 1:If DesktopMode = True Then LED3a.State = 1:END If:End Sub
Sub Bumper3_Timer()
    Select Case bump3
        Case 1:Ring3a.IsDropped = 0:bump3 = 2
        Case 2:Ring3b.IsDropped = 0:Ring3a.IsDropped = 1:bump3 = 3
        Case 3:Ring3c.IsDropped = 0:Ring3b.IsDropped = 1:bump3 = 4
        Case 4:Ring3c.IsDropped = 1:Me.TimerEnabled = 0: B3.State = 0: LED3.State = 0:LED3a.State = 0
    End Select
End Sub

Sub Bumper4_Hit:vpmTimer.PulseSw 52:PlaySound "bumper2", 0, 1, -0.15, 0.25:bump4 = 1:Me.TimerEnabled = 1:B4.State = 1 : LED4.State = 1:If DesktopMode = True Then LED4a.State = 1:END If:End Sub
Sub Bumper4_Timer()
    Select Case bump4
        Case 1:Ring4a.IsDropped = 0:bump4 = 2
        Case 2:Ring4b.IsDropped = 0:Ring4a.IsDropped = 1:bump4 = 3
        Case 3:Ring4c.IsDropped = 0:Ring4b.IsDropped = 1:bump4 = 4
        Case 4:Ring4c.IsDropped = 1:Me.TimerEnabled = 0: B4.State = 0: LED4.State = 0: LED4a.State = 0
    End Select
End Sub

' Eject holes
Sub Drain_Hit:Playsound "fx_drain":bsTrough.AddBall Me:End Sub

Sub sw21_hit:sw21.destroyball:FlashHolTie:PlaySound "fx_hole_enter":vpmTimer.PulseSwitch 21, 2500, "bsBottomVuk.AddBall":End Sub
Sub sw21a_hit:sw21a.destroyball:FlashHolTie:PlaySound "fx_hole_enter":vpmTimer.PulseSwitch 21, 2500, "bsBottomVuk.AddBall":End Sub
Sub sw21b_hit:sw21b.destroyball:FlashHolTie:PlaySound "fx_hole_enter":vpmTimer.PulseSwitch 21, 2500, "bsBottomVuk.AddBall":End Sub
Sub sw21c_hit:sw21c.destroyball:FlashHolTie:PlaySound "fx_hole_enter":vpmTimer.PulseSwitch 21, 2500, "bsBottomVuk.AddBall":End Sub
Sub sw21d_hit:sw21d.destroyball:FlashHolTie:PlaySound "fx_hole_enter":vpmTimer.PulseSwitch 21, 2500, "bsBottomVuk.AddBall":End Sub

Sub sw46_hit:PlaySound "fx_hole_enter":bsBottomVuk.AddBall Me:End Sub

Sub sw45_hit:PlaySound "fx_kicker_enter":bsTopVuk.AddBall 0:End Sub

Sub sw37_Hit:PlaySound "fx_kicker_enter":Xwingshoot = 1:bsXWingCannon.AddBall Me:End Sub

Sub TriggerSound1_Hit: Playsound "fx_ballhit": End Sub

Sub FlashHolTie
FlashHol1.State =1


End Sub


' Rollovers
Sub sw57_Hit:Controller.Switch(57) = 1:PlaySound "fx_sensor":End Sub
Sub sw57_UnHit:Controller.Switch(57) = 0:End Sub

Sub sw58_Hit:Controller.Switch(58) = 1:PlaySound "fx_sensor":End Sub
Sub sw58_UnHit:Controller.Switch(58) = 0:End Sub

Sub sw61_Hit:Controller.Switch(61) = 1:PlaySound "fx_sensor":End Sub
Sub sw61_UnHit:Controller.Switch(61) = 0:End Sub

Sub sw60_Hit:Controller.Switch(60) = 1:PlaySound "fx_sensor":End Sub
Sub sw60_UnHit:Controller.Switch(60) = 0:End Sub

Sub sw43_Hit:Controller.Switch(43) = 1:PlaySound "fx_sensor":End Sub
Sub sw43_UnHit:Controller.Switch(43) = 0:End Sub

Sub sw44_Hit:Controller.Switch(44) = 1:PlaySound "fx_sensor":End Sub
Sub sw44_UnHit:Controller.Switch(44) = 0:End Sub

Sub sw34_Hit:Controller.Switch(34) = 1:PlaySound "fx_sensor":End Sub
Sub sw34_UnHit:Controller.Switch(34) = 0:End Sub

Sub sw47_Hit:Controller.Switch(47) = 1:PlaySound "fx_sensor":End Sub
Sub sw47_UnHit:Controller.Switch(47) = 0:End Sub

' targets

Sub sw38_Hit:vpmTimer.PulseSw 38:PlaySound "fx_target":End Sub
Sub sw39_Hit:vpmTimer.PulseSw 39:PlaySound "fx_target":End Sub
Sub sw40_Hit:vpmTimer.PulseSw 40:PlaySound "fx_target":End Sub

Sub sw28_Hit:vpmTimer.PulseSw 28:sw28.IsDropped = 1:sw28a.IsDropped = 0:Me.TimerEnabled = 1:PlaySound "fx_target":End Sub
Sub sw28_Timer:sw28.IsDropped = 0:sw28a.IsDropped = 1:Me.TimerEnabled = 0:End Sub

Sub sw29_Hit:vpmTimer.PulseSw 29:sw29.IsDropped = 1:sw29a.IsDropped = 0:Me.TimerEnabled = 1:PlaySound "fx_target":End Sub
Sub sw29_Timer:sw29.IsDropped = 0:sw29a.IsDropped = 1:Me.TimerEnabled = 0:End Sub

Sub sw30_Hit:vpmTimer.PulseSw 30:sw30.IsDropped = 1:sw30a.IsDropped = 0:Me.TimerEnabled = 1:PlaySound "fx_target":End Sub
Sub sw30_Timer:sw30.IsDropped = 0:sw30a.IsDropped = 1:Me.TimerEnabled = 0:End Sub

Sub sw31_Hit:vpmTimer.PulseSw 31:sw31.IsDropped = 1:sw31a.IsDropped = 0:Me.TimerEnabled = 1:PlaySound "fx_target":End Sub
Sub sw31_Timer:sw31.IsDropped = 0:sw31a.IsDropped = 1:Me.TimerEnabled = 0:End Sub

Sub sw32_Hit:vpmTimer.PulseSw 32:sw32.IsDropped = 1:sw32a.IsDropped = 0:Me.TimerEnabled = 1:PlaySound "fx_target":End Sub
Sub sw32_Timer:sw32.IsDropped = 0:sw32a.IsDropped = 1:Me.TimerEnabled = 0:End Sub

' Droptargets

Sub sw17_Hit:dtRight.Hit 1:End Sub
Sub sw18_Hit:dtRight.Hit 2:End Sub
Sub sw19_Hit:dtRight.Hit 3:End Sub
Sub sw20_Hit:dtRight.Hit 4:End Sub

' Ramp switches

Sub sw24_Hit:Controller.Switch(24) = 1:PlaySound "fx_sensor":End Sub
Sub sw24_UnHit:Controller.Switch(24) = 0:End Sub

Sub sw22_Hit:Controller.Switch(22) = 1:PlaySound "fx_sensor":End Sub
Sub sw22_UnHit:Controller.Switch(22) = 0:End Sub

Sub sw23_Hit:Controller.Switch(23) = 1:PlaySound "fx_sensor":End Sub
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
        PlaySound "fx_flipperup1", 0, 1, -0.15, 0.25:LeftFlipper.RotateToEnd
    Else
        tmp = LeftFlipper.Strength
'        tmp2 = LeftFlipper.Recoil
        LeftFlipper.Strength = 6 'increase return strength to compensate for the slower speed
'        LeftFlipper.Recoil = 0.5 'decrease the recoil to allow drop catches
        PlaySound "fx_flipperdown", 0, 1, -0.15, 0.25:LeftFlipper.RotateToStart
        LeftFlipper.Strength = tmp
'        LeftFlipper.Recoil = tmp2
    End If
End Sub

Sub SolRFlipper(Enabled)
    Dim tmp, tmp2
    If Enabled Then
        PlaySound "fx_flipperup1", 0, 1, 0.15, 0.25:RightFlipper.RotateToEnd
    Else
        tmp = RightFlipper.Strength
'        tmp2 = LeftFlipper.Recoil
        RightFlipper.Strength = 6 'increase return strength to compensate for the slower speed
'        RightFlipper.Recoil = 0.5 'decrease the recoil to allow drop catches
        PlaySound "fx_flipperdown", 0, 1, 0.15, 0.25:RightFlipper.RotateToStart
        RightFlipper.Strength = tmp
'        RightFlipper.Recoil = tmp2
    End If
End Sub

Sub LeftFlipper_Collide(parm)
    PlaySound "fx_rubber_flipper"
End Sub

Sub RightFlipper_Collide(parm)
    PlaySound "fx_rubber_flipper"
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

' *********************************************************************
'                      Supporting Ball & Sound Functions modified by Hanibal
' *********************************************************************

Function Vol(ball) ' Calculates the Volume of the sound based on the ball speed
    Vol = Csng((BallVel(ball)^3/50) )
End Function

Function Pan(ball) ' Calculates the pan for a ball based on the X position on the table. "table1" is the name of the table
    Dim tmp
    tmp = ball.x * 2 / Table1.width-1
    If tmp > 0 Then
        Pan = Csng(tmp ^10)
    Else
        Pan = Csng(-((- tmp) ^10) )
    End If
End Function

Function Pitch(ball) ' Calculates the pitch of the sound based on the ball speed
    Pitch = BallVel(ball)^2.8*2
End Function

Function BallVel(ball) 'Calculates the ball speed
    BallVel = INT(SQR((ball.VelX ^2) + (ball.VelY ^2) ) )
End Function

'*****************************************
'      JP's VP10 Rolling Sounds modified by Hanibal
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

Sub RollingTimer_Timer()

    Dim BOT, b
    Dim Rampy
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
        If BallVel(BOT(b) ) > 1 Then 'AND BOT(b).z < 30 
            If BOT(b).z > 30 Then
            Rampy = 2
            Else
			Rampy = 1
			End if

            rolling(b) = True
            PlaySound("fx_ballrolling" & b), -1, Vol(BOT(b) ), Pan(BOT(b) ), 0, Pitch(BOT(b))*Rampy, 1, 0
			'***********************************************************
			'      Hanibal's VP10 Automated Collision Detection Sounds
			'***********************************************************

			IF Pan(BOT(b) ) > 8 Then 
			PlaySound("fx_rubber2"), 0, Vol(BOT(b)), Pan(BOT(b) ), 0, Pitch(BOT(b) ), 1, 0
			End IF



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
	PlaySound("fx_collide"), 0, Csng(velocity) ^2 / 200, Pan(ball1), 0, Pitch(ball1), 0, 0
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
    LightHan 9, l9a
'    LightHan 10, l10
    LightHan 11, l11
    LightHan 12, l12
    LightHan 13, l13
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
    LightHan 40, l40
    LightHan 41, l41
    LightHan 42, l42
    LightHan 43, l43
    LightHan 44, l44
    LightHan 45, l45
    LightHan 46, l46
    LightHan 47, l47
    LightHan 48, l48

If DesktopMode = False Then
    LightHan 49, Han1
    LightHan 50, Han2
    LightHan 51, Han3
    LightHan 52, Han4
    LightHan 53, Han5
    LightHan 54, Han6
    LightHan 55, Han7
    LightHan 56, Han8
End If

If DesktopMode = True Then
    LightHan 49, Han2
    LightHan 50, Han3
    LightHan 51, Han4
    LightHan 52, Han9
    LightHan 53, Han6
    LightHan 54, Han7
    LightHan 55, Han8
    LightHan 56, Han10
End If

    'FlashAR 57, l57, "rf_on", "rf_a", "rf_b", l57R
    'FlashAR 58, l58, "rf_on", "rf_a", "rf_b", l58R
    'FlashAR 59, l59, "rf_on", "rf_a", "rf_b", l59R
    'FlashAR 60, l60, "rf_on", "rf_a", "rf_b", l60R
    LightHan 61, l61
    LightHan 62, l62
    'FadePri 63, l63, "prigf_on", "prigf_a", "prigf_b", "f_off"
    'FadePri 64, l64, "prigf_on", "prigf_a", "prigf_b", "f_off"
    LightHan 65, l65
    LightHan 66, l66
    LightHan 67, l67
    LightHan 68, l68
    LightHan 69, l69
    LightHan 70, l70
    LightHan 71, l71
    LightHan 72, l72
    LightHan 73, l73
    LightHan 74, l74
    LightHan 75, l75
    LightHan 76, l76
    LightHan 77, l77
    LightHan 78, l78
    LightHan 79, l79
    LightHan 80, l80
    'flashers
    'FadeL 103, f3, f3a
    'FlashAR 104, f4, "gf2_on", "gf2_a", "gf2_b", f4R
    'FlashAR 105, f5, "gf2_on", "gf2_a", "gf2_b", f5R
    'FadeLm 106, f6, f6a
    'FadeL 106, f6b, f6c
    'FlashAR 107, f7, "wf_on", "wf_a", "wf_b", f7R
    'FadeLm 108, f8, f8a
    'FadeL 108, f8b, f8c
End Sub

Sub FlasherTimer_Timer()



    Flash 49, z49
    Flash 50, z50
    Flash 51, z51
    Flash 52, z52
    Flash 53, z53
    Flash 54, z54
    Flash 55, z55
    Flash 56, z56
   


'**************GI LIGHTS*************

    LightHan 13, gi1
    LightHan 13, gi2
    LightHan 13, gi3
    LightHan 13, gi4
    LightHan 13, gi5
    LightHan 13, gi6
    LightHan 13, gi7
    LightHan 13, gi8
    LightHan 13, gi9
    LightHan 13, gi10
    LightHan 13, gi11
    LightHan 13, gi12
    LightHan 13, gi13
    LightHan 13, gi14
    LightHan 13, gi15
    LightHan 13, gi16
    LightHan 13, gi17
    LightHan 13, gi18
    LightHan 13, gi19
    LightHan 13, gi20
    LightHan 13, gi21
    LightHan 13, gi22
    LightHan 13, gi23
    LightHan 13, gi24
    LightHan 13, gi25
    LightHan 13, gi26
    LightHan 13, gi27
    LightHan 13, gi28
    LightHan 13, gi29
    LightHan 13, gi30
    LightHan 13, gi31
    LightHan 13, gi32
    LightHan 13, gi33
    LightHan 13, gi34
    LightHan 13, gi35
    LightHan 13, gi36
   LightHan 13, gi37
   LightHan 13, gi38
   LightHan 13, gi39
   LightHan 13, gi40
   LightHan 13, gi41
   LightHan 13, gi42
   LightHan 13, gi43
   LightHan 13, gi44
   LightHan 13, gi45
   LightHan 13, gi46
   LightHan 13, gi47
   LightHan 13, gi48
   LightHan 13, gi49
   LightHan 13, gi50
   LightHan 13, gi51
   LightHan 13, gi52
   LightHan 13, gi53
   LightHan 13, gi54
   LightHan 13, gi55
   LightHan 13, gi56
   LightHan 13, gi57
   LightHan 13, gi58
   LightHan 13, gi59
   LightHan 13, gi60
   LightHan 13, gi61
   LightHan 13, gi62
   LightHan 13, gi63
   LightHan 13, gi64
   LightHan 13, gi65
   LightHan 13, gi66
   LightHan 13, gi67
   LightHan 13, gi68
   LightHan 13, gi69
   LightHan 13, gi70
   LightHan 13, gi71



'***************Flashers Hanibal Spezial*************
  


 
	FlashHan 3, F3a
    FlashHan 3, F3b
    FlashHan 3, f3c
    FlashHan 3, FlashTie
	FlashHan 6, F6a
    FlashHan 6, F6b
    FlashHan 6, F6c
    FlashHan 6, F6d
    FlashHan 6, f6e
    FlashHan 6, f6f
    FlashHan 8, F8a
    FlashHan 8, F8b
    FlashHan 8, F8c
    FlashHan 8, F8d 
    FlashHan 8, f8e 
    FlashHan 8, f8f

    FlashSoundHan 3, FlashSoundA
    FlashSoundHan 6, FlashSoundB
    FlashSoundHan 8, FlashSoundC


 

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
        Case 4:object.state = 0 : object.image = "lights"
        Case 5:object.state = 1 : object.image = "lights_on"
    End Select

End Sub


'****************** Hanibal special sound flasher ****************************




Sub FlashSoundHan(nr, value) ' used for playing sound to Light

    Select Case FlashState(nr)
        
        Case 0: value = 1
        Case 1: If value = 1 Then Playsound "Relais", 1, 0.005, 0, 1: value = 0
  
    End Select

End Sub


'****************** Hanibal special lights ****************************



Sub LightHan(nr, object) ' used for multiple lights and pass playfield
    Select Case FadingLevel(nr)
        Case 4:object.state = 0 : object.image = "lights"
        Case 5:object.state = 1 : object.image = "lights_on" 
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




'********************
' Diverse Help/Sounds
'********************

Sub AllRubbers_Hit(idx):PlaySound "fx_rubber", 0, (20 + Vol(ActiveBall)), (20+Pan(ActiveBall)), Pitch(ActiveBall), 1, 1:End Sub
Sub AllPostRubbers_Hit(idx):PlaySound "fx_rubber", 0, (20 + Vol(ActiveBall)), (20+Pan(ActiveBall)), Pitch(ActiveBall), 1, 1:End Sub
Sub AllMetals_Hit(idx):PlaySound "fx_MetalHit", 0, (20 + Vol(ActiveBall)), (20+Pan(ActiveBall)), Pitch(ActiveBall), 1, 1:End Sub
Sub AllGates_Hit(idx):PlaySound "fx_Gate", 0, (20 + Vol(ActiveBall)), (20+Pan(ActiveBall)), Pitch(ActiveBall), 1, 1:End Sub
Sub RHelp1_hit
    Playsound "fx_ballhit"
End Sub

Sub RHelp2_hit
    Playsound "fx_ballhit"
End Sub

Sub Relaissound

Playsound "Relais", 0, 1, 0, 1
End Sub
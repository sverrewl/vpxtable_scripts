Option Explicit
Randomize

' Thalamus 2019 February : Improved directional sounds
' !! NOTE : Table not verified yet !!

' Options
' Volume devided by - lower gets higher sound

Const VolDiv = 2000    ' Lower number, louder ballrolling/collition sound
Const VolCol = 10      ' Ball collition divider ( voldiv/volcol )

' The rest of the values are multipliers
'
'  .5 = lower volume
' 1.5 = higher volume

Const VolBump   = 1    ' Bumpers volume.
Const VolGates  = 1    ' Gates volume.
Const VolMetal  = 1    ' Metals volume.
Const VolRH     = 1    ' Rubber hits volume.
Const VolPi     = 1    ' Rubber pins volume.
Const VolTarg   = 1    ' Targets volume.
Const VolSpin   = 1    ' Spinners volume.
Const VolFlip   = 1    ' Flipper volume.

On Error Resume Next
ExecuteGlobal GetTextFile("controller.vbs")
If Err Then MsgBox "You need the controller.vbs in order to run this table, available in the vp10 package"
On Error Goto 0

Const cGameName="swtril43",UseSolenoids=1,UseLamps=0,UseGI=0,SSolenoidOn="SolOn",SSolenoidOff="SolOff", SCoin="coin"

LoadVPM "01000200", "SEGA.VBS", 3.02

Dim DesktopMode: DesktopMode = Table1.ShowDT
If DesktopMode = True Then 'Show Desktop components
Ramp16.visible=1
Ramp15.visible=1
Primitive13.visible=1
Else
Ramp16.visible=0
Ramp15.visible=0
Primitive13.visible=0
End if

'*************************************************************
'Solenoid Call backs
'**********************************************************************************************************
SolCallback(1) = "bsTrough.SolOut"
SolCallback(2) = "AutoPlunger"
SolCallBack(3) = "dtRight.SolDropUp"
SolCallBack(4) = "VukTopPop" ' Wire ramp up kick
SolCallBack(5) = "SolFireXWing" 'X-Wing Cannon
SolCallBack(6) = "bsSaucer.SolOut"
SolCallBack(7) = "SolRampMagnet" 'Cannon Magnet
SolCallBack(17) = "SolDiv"  'X-Wing Diverter
'SolCallBack(18) = "XWingMotor" 'X-Wing Motor Relay
SolCallBack(20) = "dtRight.SolHit 1,"
SolCallBack(21) = "dtRight.SolHit 2,"
SolCallBack(22) = "dtRight.SolHit 3,"
SolCallBack(23) = "dtRight.SolHit 4,"
SolCallBack(25) = "SolMagnetSlide"  'ramp magnet + cannon animation
SolCallBack(26) = "SolTieFighter" 'Tie Fighter Shake
SolCallBack(27) = "SetLamp 103," 'F3
SolCallBack(28) = "SetLamp 104," 'F4
SolCallBack(29) = "SetLamp 105," 'F5
SolCallBack(31) = "SetLamp 106," 'F6
SolCallBack(30) = "SetLamp 107," 'F7
SolCallBack(32) = "SetLamp 108," 'F8

SolCallback(sLRFlipper) = "SolRFlipper"
SolCallback(sLLFlipper) = "SolLFlipper"

Sub SolLFlipper(Enabled)
     If Enabled Then
         PlaySoundAtVol SoundFX("fx_Flipperup",DOFContactors), LeftFlipper, VolFlip:LeftFlipper.RotateToEnd
     Else
         PlaySoundAtVol SoundFX("fx_Flipperdown",DOFContactors), LeftFlipper, VolFlip:LeftFlipper.RotateToStart
     End If
  End Sub

Sub SolRFlipper(Enabled)
     If Enabled Then
         PlaySoundAtVol SoundFX("fx_Flipperup",DOFContactors), RightFlipper, volFlip:RightFlipper.RotateToEnd
     Else
         PlaySoundAtVol SoundFX("fx_Flipperdown",DOFContactors), RightFlipper, volFlip:RightFlipper.RotateToStart
     End If
End Sub
'**********************************************************************************************************

'Solenoid Controlled toys
'**********************************************************************************************************

Sub AutoPlunger(Enabled)
    If Enabled Then
        PlungerIM.AutoFire
    End If
End Sub


Sub SolFireXWing(Enabled)
    If Enabled AND Xwingshoot = 1 Then
        bsSaucer1.InitKick sw37, KickDir, 32
        bsSaucer1.ExitSol_On
        Xwingshoot = 0
        PlaySound SoundFX("Popper",DOFContactors)
    End if
End Sub


Sub SolRampMagnet(Enabled)
    If Enabled Then
        Magnet.Enabled = True
        Magnet1.Enabled = True
        Magnet2.Enabled = True
        Magnet3.Enabled = True
        Magnet4.Enabled = True
    Else
        Magnet.Enabled = False
        Magnet1.Enabled = True
        Magnet2.Enabled = False
        Magnet3.Enabled = True
        Magnet4.Enabled = False
    End If
End Sub


Sub Soldiv(Enabled)
    If Enabled Then
        Diverter.RotateToEnd
		PlaySound SoundFX("fx_Flipperup",DOFContactors)
    Else
        Diverter.RotateToStart
		PlaySound SoundFX("fx_Flipperdown",DOFContactors)
    End If
End Sub


Sub SolMagnetSlide(Enabled)
    If Enabled Then
		cannonpos = 0:
        Cannonanimate.Enabled = True
        PlaySound SoundFX("fx_Flipperup",DOFContactors)
    End If
End Sub


Sub SolTieFighter(Enabled)
    If Enabled Then
        tiepostemp = 10
        tiespeed = 10
        Tiefightershake.Enabled = True
    Else
		tiesound = False
    End If
End Sub


set GICallback = GetRef("UpdateGI")
Sub UpdateGI(no, Enabled)
	If Enabled Then
		dim xx
		For each xx in GI:xx.State = 1:	Next
        PlaySound "fx_relay"
    		DOF 101, DOFOn

	Else For each xx in GI:xx.State = 0: Next
        PlaySound "fx_relay"
	    	DOF 101, DOFOff
	End If
End Sub


'**********************************************************************************************************

'Initiate Table
'**********************************************************************************************************

Dim bsTrough, bsSaucer, dtRight, mHole
Dim bsSaucer1, mXWing

Sub Table1_Init
	vpmInit Me
	On Error Resume Next
		With Controller
		.GameName = cGameName
		If Err Then MsgBox "Can't start Game" & cGameName & vbNewLine & Err.Description : Exit Sub
		.SplashInfoLine = "Star Wars Trilogy - Sega 1997"&chr(13)&"You Suck"
		.HandleMechanics=0
		.HandleKeyboard=0
		.ShowDMDOnly=1
		.ShowFrame=0
		.ShowTitle=0
		.hidden = 0
         On Error Resume Next
         .Run GetPlayerHWnd
         If Err Then MsgBox Err.Description
         On Error Goto 0
     End With
     On Error Goto 0

	PinMAMETimer.Interval=PinMAMEInterval
	PinMAMETimer.Enabled=1
    vpmNudge.TiltSwitch = 56
    vpmNudge.Sensitivity = 2
    vpmNudge.TiltObj = Array(Bumper1, Bumper2, Bumper3, Bumper4, LeftSlingshot, RightSlingshot)

    Set bsTrough = New cvpmBallStack
        bsTrough.InitSw 0, 15, 14, 13, 12, 0, 0, 0
        bsTrough.InitKick BallRelease, 110, 5
        bsTrough.InitExitSnd SoundFX("ballrelease",DOFContactors), SoundFX("Solenoid",DOFContactors)
        bsTrough.IsTrough = True
        bsTrough.Balls = 4

    Set bsSaucer1 = New cvpmBallStack
        bsSaucer1.InitSw 0, 37, 0, 0, 0, 0, 0, 0
        bsSaucer1.InitExitSnd SoundFX("Popper",DOFContactors), SoundFX("Solenoid",DOFContactors)
        bsSaucer1.Balls = 0

    Set bsSaucer = New cvpmBallStack
        bsSaucer.InitSw 0, 46, 0, 0, 0, 0, 0, 0
        bsSaucer.InitKick sw46a, 190, 1
        bsSaucer.InitExitSnd SoundFX("Popper",DOFContactors), SoundFX("Solenoid",DOFContactors)
        bsSaucer.Balls = 0

    Set dtRight = New cvpmDropTarget
        dtRight.InitDrop Array(sw17, sw18, sw19, sw20), Array(17, 18, 19, 20)
        dtRight.InitSnd SoundFX("DTDrop",DOFContactors),SoundFX("DTReset",DOFContactors)

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

End Sub

'**********************************************************************************************************
'Plunger code
'**********************************************************************************************************

Sub Table1_KeyDown(ByVal KeyCode)

    If keycode = PlungerKey Then Controller.Switch(53) = 1
	If KeyDownHandler(keycode) Then Exit Sub
End Sub

Sub Table1_KeyUp(ByVal KeyCode)

    If keycode = PlungerKey Then Controller.Switch(53) = 0
	If KeyUpHandler(keycode) Then Exit Sub
End Sub

	Dim plungerIM
    Const IMPowerSetting = 42 'Plunger Power
    Const IMTime = 0.6        ' Time in seconds for Full Plunge
    Set plungerIM = New cvpmImpulseP
    With plungerIM
        .InitImpulseP swplunger, IMPowerSetting, IMTime
        .Random 0.3
        .switch 16
        .InitExitSnd  SoundFX("Popper",DOFContactors), SoundFX("Solenoid",DOFContactors)
        .CreateEvents "plungerIM"
    End With


'**********************************************************************************************************

 ' Drain hole and kickers
Sub Drain_Hit:bsTrough.addball me : playsoundAtVol"drain" , drain, 1: End Sub

Sub sw21_hit:sw21.destroyball:vpmTimer.PulseSwitch 21, 2500, "bsSaucer.AddBall": playsoundAtVol "popper_ball", sw21, 1: End Sub
Sub sw21a_hit:sw21a.destroyball:vpmTimer.PulseSwitch 21, 2500, "bsSaucer.AddBall": playsoundAtVol "popper_ball", sw21a, 1: End Sub
Sub sw21b_hit:sw21b.destroyball:vpmTimer.PulseSwitch 21, 2500, "bsSaucer.AddBall": playsoundAtVol "popper_ball", sw21b, 1: End Sub
Sub sw21c_hit:sw21c.destroyball:vpmTimer.PulseSwitch 21, 2500, "bsSaucer.AddBall": playsoundAtVol "popper_ball", sw21c, 1: End Sub
Sub sw21d_hit:sw21d.destroyball:vpmTimer.PulseSwitch 21, 2500, "bsSaucer.AddBall": playsoundAtVol "popper_ball", sw21d, 1: End Sub

Sub sw37_Hit:bsSaucer1.AddBall Me : playsoundAtVol "popper_ball", sw37, 1: Xwingshoot = 1 : End Sub

Sub sw46_hit:bsSaucer.AddBall Me : playsoundAtVol "popper_ball", sw46, 1: End Sub

 '***********************************
 'Top Raising VUK
 '***********************************
 'Variables used for VUK
 Dim raiseballsw, raiseball
 Sub sw45_Hit()
 	sw45.Enabled=FALSE
	Controller.switch (45) = True
	playsoundAtVol"popper_ball", ActiveBall, 1
 End Sub

 Sub VukTopPop(enabled)
	if(enabled and Controller.switch (45)) then
		playsound SoundFX("Popper",DOFContactors)
		sw45.DestroyBall
 		Set raiseball = sw45.CreateBall
 		raiseballsw = True
 		sw45raiseballtimer.Enabled = True 'Added by Rascal
		sw45.Enabled=TRUE
 		Controller.switch (45) = False
	else
		PlaySound "Popper"
	end if
End Sub

 Sub sw45raiseballtimer_Timer()
 	If raiseballsw = True then
 		raiseball.z = raiseball.z + 10
 		If raiseball.z > 220 then
 			sw45.Kick 0, 15
 			Set raiseball = Nothing
 			sw45raiseballtimer.Enabled = False
 			raiseballsw = False
 		End If
 	End If
 End Sub

'Drop Targets
 Sub sw17_Dropped:dtRight.Hit 1 :End Sub
 Sub sw18_Dropped:dtRight.Hit 2 :End Sub
 Sub sw19_Dropped:dtRight.Hit 3 :End Sub
 Sub sw20_Dropped:dtRight.Hit 4 :End Sub

'Bumpers
Sub Bumper1_Hit : vpmTimer.PulseSw(49) : playsoundAtVol SoundFX("fx_bumper1",DOFContactors), ActiveBall, VolBump: End Sub
Sub Bumper2_Hit : vpmTimer.PulseSw(50) : playsoundAtVol SoundFX("fx_bumper1",DOFContactors), ActiveBall, VolBump: End Sub
Sub Bumper3_Hit : vpmTimer.PulseSw(51) : playsoundAtVol SoundFX("fx_bumper1",DOFContactors), ActiveBall, VolBump: End Sub
Sub Bumper4_Hit : vpmTimer.PulseSw(52) : playsoundAtVol SoundFX("fx_bumper1",DOFContactors), ActiveBall, VolBump: End Sub

'Wire Triggers
Sub sw16_Hit:: playsoundAtVol"rollover" , ActiveBall, 1: End Sub 'coded to impulse plunger
'Sub sw16_UnHit:Controller.Switch(16) = 0:End Sub
Sub sw34_Hit:Controller.Switch(34) = 1 : playsoundAtVol"rollover" , ActiveBall, 1: End Sub
Sub sw34_UnHit:Controller.Switch(34) = 0:End Sub
Sub sw43_Hit:Controller.Switch(43) = 1 : playsoundAtVol"rollover" , ActiveBall, 1: End Sub
Sub sw43_UnHit:Controller.Switch(43) = 0:End Sub
Sub sw44_Hit:Controller.Switch(44) = 1 : playsoundAtVol"rollover" , ActiveBall, 1: End Sub
Sub sw44_UnHit:Controller.Switch(44) = 0:End Sub
Sub sw47_Hit:Controller.Switch(47) = 1 : playsoundAtVol"rollover" , ActiveBall, 1: End Sub
Sub sw47_UnHit:Controller.Switch(47) = 0:End Sub
Sub sw57_Hit:Controller.Switch(57) = 1 : playsoundAtVol"rollover" , ActiveBall, 1: End Sub
Sub sw57_UnHit:Controller.Switch(57) = 0:End Sub
Sub sw58_Hit:Controller.Switch(58) = 1 : playsoundAtVol"rollover" , ActiveBall, 1: End Sub
Sub sw58_UnHit:Controller.Switch(58) = 0:End Sub
Sub sw60_Hit:Controller.Switch(60) = 1 : playsoundAtVol"rollover" , ActiveBall, 1: End Sub
Sub sw60_UnHit:Controller.Switch(60) = 0:End Sub

' Ramp triggers
Sub sw22_hit:Controller.Switch(22) = 1 : playsoundAtVol"rollover" , ActiveBall, 1: End Sub
Sub sw22_UnHit:Controller.Switch(22) = 0:End Sub
Sub sw23_hit:Controller.Switch(23) = 1 : playsoundAtVol"rollover" , ActiveBall, 1: End Sub
Sub sw23_UnHit:Controller.Switch(23) = 0:End Sub
Sub sw24_hit:Controller.Switch(24) = 1 : playsoundAtVol"rollover" , ActiveBall, 1: End Sub
Sub sw24_UnHit:Controller.Switch(24) = 0:End Sub

'Stand Up Targets
Sub sw28_Hit:vpmTimer.PulseSw 28:End Sub
Sub sw29_Hit:vpmTimer.PulseSw 29:End Sub
Sub sw30_Hit:vpmTimer.PulseSw 30:End Sub
Sub sw31_Hit:vpmTimer.PulseSw 31:End Sub
Sub sw32_Hit:vpmTimer.PulseSw 32:End Sub
Sub sw38_Hit:vpmTimer.PulseSw 38:End Sub
Sub sw39_Hit:vpmTimer.PulseSw 39:End Sub
Sub sw40_Hit:vpmTimer.PulseSw 40:End Sub

'Generic Sounds
Sub Trigger1_Hit:PlaySoundAtVol "fx_ballrampdrop", ActiveBall, 1:End Sub
Sub Trigger2_Hit:PlaySoundAtVol "fx_ballrampdrop", ActiveBall, 1:End Sub
Sub Trigger3_Hit:PlaySoundAtVol "fx_ballrampdrop", ActiveBall, 1:StopSound "Wire Ramp":End Sub

Sub Trigger4_Hit:PlaySoundAtVol "Wire Ramp", ActiveBall, 1:End Sub


Sub Magnet_Hit() 'if magnet enabled move ball to left ramp over Trigger 5
    ActiveBall.X = 197
	ActiveBall.Y = 153
	ActiveBall.Z = 120
	ActiveBall.VelX = 0
	ActiveBall.VelY = 0
End Sub

Sub Magnet1_Hit() 'if magnet enabled move ball to left ramp over Trigger 5
    ActiveBall.X = 197
	ActiveBall.Y = 153
	ActiveBall.Z = 120
	ActiveBall.VelX = 0
	ActiveBall.VelY = 0
End Sub

Sub Magnet2_Hit() 'if magnet enabled move ball to left ramp over Trigger 5
    ActiveBall.X = 197
	ActiveBall.Y = 153
	ActiveBall.Z = 120
	ActiveBall.VelX = 0
	ActiveBall.VelY = 0
End Sub

Sub Magnet3_Hit() 'if magnet enabled move ball to left ramp over Trigger 5
    ActiveBall.X = 197
	ActiveBall.Y = 153
	ActiveBall.Z = 120
	ActiveBall.VelX = 0
	ActiveBall.VelY = 0
End Sub

Sub Magnet4_Hit() 'if magnet enabled move ball to left ramp over Trigger 5
    ActiveBall.X = 197
	ActiveBall.Y = 153
	ActiveBall.Z = 120
	ActiveBall.VelX = 0
	ActiveBall.VelY = 0
End Sub


'**********************************************************************************************************
'Primitive animations
'**********************************************************************************************************

' Tie Fighter
Dim tiepos, tiepostemp, tiespeed, tiesound

Sub Tiefightershake_timer
	TieFighter.RotAndTra2 = tiepos
	If tiepostemp <0.1 AND tiepostemp >-0.1  Then : Tiefightershake.Enabled = False: Exit Sub
    If tiesound = False Then PlaySound SoundFX("Motor",DOFContactors) : tiesound = True

If tiepostemp < 0 Then
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
End If

End Sub

'Cannon
Dim cannonpos

Sub Cannonanimate_Timer
    Select case cannonpos
        Case 0:P64.Y=P63.Y=cannonMuzzle.Y=cannonMuzzle.Y-15:P64.X=P63.X=cannonMuzzle.X=cannonMuzzle.X-15
        Case 1:P64.Y=P63.Y=cannonMuzzle.Y=cannonMuzzle.Y-9:P64.X=P63.X=cannonMuzzle.X=cannonMuzzle.X-9:cannonBase.Y=cannonBase.Y-1:cannonBase.X=cannonBase.X-1
        Case 2:P64.Y=P63.Y=cannonMuzzle.Y=cannonMuzzle.Y-5:P64.X=P63.X=cannonMuzzle.X=cannonMuzzle.X-5:cannonBase.Y=cannonBase.Y-3:cannonBase.X=cannonBase.X-3
        Case 3:P64.Y=P63.Y=cannonMuzzle.Y=cannonMuzzle.Y-3:P64.X=P63.X=cannonMuzzle.X=cannonMuzzle.X-3:cannonBase.Y=cannonBase.Y-1:cannonBase.X=cannonBase.X-1
        Case 4:P64.Y=P63.Y=cannonMuzzle.Y=cannonMuzzle.Y+3:P64.X=P63.X=cannonMuzzle.X=cannonMuzzle.X+3:cannonBase.Y=cannonBase.Y+1:cannonBase.X=cannonBase.X+1
        Case 5:P64.Y=P63.Y=cannonMuzzle.Y=cannonMuzzle.Y+5:P64.X=P63.X=cannonMuzzle.X=cannonMuzzle.X+5:cannonBase.Y=cannonBase.Y+3:cannonBase.X=cannonBase.X+3
        Case 6:P64.Y=P63.Y=cannonMuzzle.Y=cannonMuzzle.Y+9:P64.X=P63.X=cannonMuzzle.X=cannonMuzzle.X+9:cannonBase.Y=cannonBase.Y+1:cannonBase.X=cannonBase.X+1
        Case 7:P64.Y=P63.Y=cannonMuzzle.Y=cannonMuzzle.Y+15:P64.X=P63.X=cannonMuzzle.X=cannonMuzzle.X+15
        Case 8:Cannonanimate.Enabled = False
    End Select

    cannonpos = cannonpos + 1
  '  canonR.State = ABS(canonR.State -1)
End Sub

'X-Wing
Dim KickDir, Xwingshoot

Sub UpdateXWing(aNewPos, aSpeed, aLastPos)
    KickDir = aNewPos / 4
    XwingBase.RotAndTra2 = KickDir
    xwing.RotAndTra2 = KickDir
    PlaySoundAtVol SoundFX("Motor",DOFContactors), xwing, 1
End Sub

'----Xwing test fix------
Sub fixwing_Timer()
     If KickDir <= 5 and Controller.Switch(35) = False Then
     Controller.Switch(35) = True
     End If
End Sub
'---End Fix----

'**********************************************************************************************************
'**********************************************************************************************************

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
LampTimer.Interval = 5 'lamp fading speed
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
    NFadeL 3, l3
    NFadeL 4, l4
    NFadeL 5, l5
    NFadeL 6, l6
    NFadeL 7, l7
    NFadeL 8, l8
    NFadeL 9, l9
    'NFadeL 10, l10 'Launch Button
    NFadeL 11, l11
    NFadeL 12, l12
    NFadeL 13, l13
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
    Flash 49, F49 'han solo toy
    Flash 50, F50 'han solo toy
    Flash 51, F51 'han solo toy
    Flash 52, F52 'han solo toy
    Flash 53, F53 'han solo toy
    Flash 54, F54 'han solo toy
    Flash 55, F55 'han solo toy
    Flash 56, F56 'han solo toy
    NFadeObjm 57, P57, "TopRedOn", "TopRed"
    NFadeLm 57, l57 'Bumer 1
    NFadeL 57, l57a 'Bumer 1
    NFadeObjm 58, P58, "TopRedOn", "TopRed"
    NFadeLm 58, l58 'Bumer 2
    NFadeL 58, l58a 'Bumer 2
    NFadeObjm 59, P59, "TopRedOn", "TopRed"
    NFadeLm 59, l59 'Bumer 3
    NFadeL 59, l59a 'Bumer 3
    NFadeObjm 60, P60, "TopRedOn", "TopRed"
    NFadeLm 60, l60 'Bumer 4
    NFadeL 60, l60a 'Bumer 4
    NFadeL 61, l61
    NFadeL 62, l62
    NFadeObjm 63, P63, "bulbcover1_redOn", "bulbcover1_red"
  '  Flash 59, F63 'Cannon LED
    NFadeObjm 64, P64, "bulbcover1_redOn", "bulbcover1_red"
  '  Flash 64, F64 'Cannon LED
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

'Solenoid Controlled

    NFadeLm 103, F103a
    NFadeL 103, F103b

    NFadeL 104, F104

    NFadeL 105, F105

    NFadeLm 106, F106a
    NFadeLm 106, F106b
    NFadeLm 106, F106c
    NFadeL 106, F106d

    NFadeL 107, F107

    NFadeLm 108, F108a
    NFadeLm 108, F108b
    NFadeLm 108, F108c
    NFadeL 108, F108d


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

'**********Sling Shot Animations
' Rstep and Lstep  are the variables that increment the animation
'****************
Dim RStep, Lstep

Sub RightSlingShot_Slingshot
	vpmTimer.PulseSw 62
    PlaySoundAtVol SoundFX("right_slingshot",DOFContactors), sling1, 1
    RSling.Visible = 0
    RSling1.Visible = 1
    sling1.rotx = 20
    RStep = 0
    RightSlingShot.TimerEnabled = 1
End Sub

Sub RightSlingShot_Timer
    Select Case RStep
        Case 3:RSLing1.Visible = 0:RSLing2.Visible = 1:sling1.rotx = 10
        Case 4:RSLing2.Visible = 0:RSLing.Visible = 1:sling1.rotx = 0:RightSlingShot.TimerEnabled = 0
    End Select
    RStep = RStep + 1
End Sub

Sub LeftSlingShot_Slingshot
	vpmTimer.PulseSw 59
    PlaySoundAtVol SoundFX("left_slingshot",DOFContactors), sling2, 1
    LSling.Visible = 0
    LSling1.Visible = 1
    sling2.rotx = 20
    LStep = 0
    LeftSlingShot.TimerEnabled = 1
End Sub

Sub LeftSlingShot_Timer
    Select Case LStep
        Case 3:LSLing1.Visible = 0:LSLing2.Visible = 1:sling2.rotx = 10
        Case 4:LSLing2.Visible = 0:LSLing.Visible = 1:sling2.rotx = 0:LeftSlingShot.TimerEnabled = 0
    End Select
    LStep = LStep + 1
End Sub

' *******************************************************************************************************
' Positional Sound Playback Functions by DJRobX and Rothbauerw
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

Sub PlaySoundAtBOTBallZ(sound, BOT)
		PlaySound sound, 0, ABS(BOT.velz)/17, Pan(BOT), 0, Pitch(BOT), 1, 0, AudioFade(BOT)
End Sub

'*********************************************************************
'                     Supporting Ball & Sound Functions
'*********************************************************************

Function RndNum(min, max)
    RndNum = Int(Rnd() * (max-min + 1) ) + min ' Sets a random number between min and max
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
  Vol = Csng(BallVel(ball) ^2 / VolDiv)
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

'*****************************************
'	ninuzzu's	FLIPPER SHADOWS
'*****************************************

sub FlipperTimer_Timer()
	FlipperLSh.RotZ = LeftFlipper.currentangle
	FlipperRSh.RotZ = RightFlipper.currentangle
    LFLogo.RotY = LeftFlipper.CurrentAngle
    RFlogo.RotY = RightFlipper.CurrentAngle

End Sub

'*****************************************
'	ninuzzu's	BALL SHADOW
'*****************************************
Dim BallShadow
BallShadow = Array (BallShadow1,BallShadow2,BallShadow3,BallShadow4,BallShadow5)

Sub BallShadowUpdate_timer()
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
        If BOT(b).X < Table1.Width/2 Then
            BallShadow(b).X = ((BOT(b).X) - (Ballsize/6) + ((BOT(b).X - (Table1.Width/2))/7)) + 6
        Else
            BallShadow(b).X = ((BOT(b).X) + (Ballsize/6) + ((BOT(b).X - (Table1.Width/2))/7)) - 6
        End If
        ballShadow(b).Y = BOT(b).Y + 12
        If BOT(b).Z > 20 Then
            BallShadow(b).visible = 1
        Else
            BallShadow(b).visible = 0
        End If
    Next
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
	PlaySound "pinhit_low", 0, Vol(ActiveBall)*VolPi, AudioPan(ActiveBall), 0, Pitch(ActiveBall), 0, 0, AudioFade(ActiveBall)
End Sub

Sub Targets_Hit (idx)
	PlaySound "target", 0, Vol(ActiveBall)*VolTarg, AudioPan(ActiveBall), 0, Pitch(ActiveBall), 0, 0, AudioFade(ActiveBall)
End Sub

Sub Metals_Thin_Hit (idx)
	PlaySound "metalhit_thin", 0, Vol(ActiveBall)*VolMetal, AudioPan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
End Sub

Sub Metals_Medium_Hit (idx)
	PlaySound "metalhit_medium", 0, Vol(ActiveBall)*VolMetal, AudioPan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
End Sub

Sub Metals2_Hit (idx)
	PlaySound "metalhit2", 0, Vol(ActiveBall)*VolMetal, AudioPan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
End Sub

Sub Gates_Hit (idx)
	PlaySound "gate4", 0, Vol(ActiveBall)*VolGates, AudioPan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
End Sub

Sub Spinner_Spin
	PlaySoundAtVol "fx_spinner", Spinner, VolSpin
End Sub

Sub Rubbers_Hit(idx)
 	dim finalspeed
  	finalspeed=SQR(activeball.velx * activeball.velx + activeball.vely * activeball.vely)
 	If finalspeed > 20 then
		PlaySound "fx_rubber2", 0, Vol(ActiveBall)*VolRH, AudioPan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
	End if
	If finalspeed >= 6 AND finalspeed <= 20 then
 		RandomSoundRubber()
 	End If
End Sub

Sub Posts_Hit(idx)
 	dim finalspeed
  	finalspeed=SQR(activeball.velx * activeball.velx + activeball.vely * activeball.vely)
 	If finalspeed > 16 then
		PlaySound "fx_rubber2", 0, Vol(ActiveBall)*VolRH, AudioPan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
	End if
	If finalspeed >= 6 AND finalspeed <= 16 then
 		RandomSoundRubber()
 	End If
End Sub

Sub RandomSoundRubber()
	Select Case Int(Rnd*3)+1
		Case 1 : PlaySound "rubber_hit_1", 0, Vol(ActiveBall)*volRH, AudioPan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
		Case 2 : PlaySound "rubber_hit_2", 0, Vol(ActiveBall)*volRH, AudioPan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
		Case 3 : PlaySound "rubber_hit_3", 0, Vol(ActiveBall)*volRH, AudioPan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
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
		Case 1 : PlaySound "flip_hit_1", 0, Vol(ActiveBall)*VolRH, AudioPan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
		Case 2 : PlaySound "flip_hit_2", 0, Vol(ActiveBall)*VolRH, AudioPan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
		Case 3 : PlaySound "flip_hit_3", 0, Vol(ActiveBall)*VolRH, AudioPan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
	End Select
End Sub

' Thalamus : Exit in a clean and proper way
Sub Table1_exit()
  On Error Resume Next
  Controller.Pause = False
  Controller.Stop
End Sub


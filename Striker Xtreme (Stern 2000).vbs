Option Explicit
Randomize

' Thalamus 2019 March : Improved directional sounds
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

Const cGameName="strikext",UseSolenoids=2,UseLamps=0,UseGI=0,UseSync=1,SSolenoidOn="SolOn",SSolenoidOff="SolOff", SCoin="coin"

LoadVPM "01520000","Sega.VBS",3.02


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
SolCallback(1)= "SolTrough"
SolCallback(2)= "AutoPlunger"
SolCallback(3)=	"VukLEFTPop"
SolCallback(4)="uVUK"

SolCallback(6)= "dtDrop.SolDropUp"
SolCallback(7)= "dtDrop.SolHit 1,"
SolCallback(8)= "StLock"

SolCallback(12)= "Magnet1.MagnetOn="
'SolCallback(13)= "Magnet2.MagnetOn="

SolCallback(19)= "SolBallDeflector"
SolCallback(20)="SetLamp 120," 'Stadium X4
SolCallback(21)= "dtDrop.SolHit 2,"
SolCallback(22)= "dtDrop.SolHit 3,"
SolCallback(23)= "dtDrop.SolHit 4,"

'SolCallback(25)= "" 'Goalie
'SolCallback(26)= "" 'Goalie
SolCallBack(27)="SetLamp 127," 'Upper Flipper X1
SolCallBack(28)="SetLamp 128," 'Spinner X2
SolCallBack(29)="SetLamp 129," 'Rampp X1
SolCallBack(30)="SetLamp 130," 'Back Panel X4
SolCallback(31)="SetLamp 131," 'Pops X4
SolCallback(32)="SetLamp 132," 'Slingshots X4

'Aux board UK only
'SolCallback(33)= "SolLeftPost"
'SolCallback(34)= "SolMidPost"
'SolCallback(35)= "SolRightPost"

SolCallback(sLRFlipper) = "SolRFlipper"
SolCallback(sLLFlipper) = "SolLFlipper"

Sub SolLFlipper(Enabled)
     If Enabled Then
         PlaySoundAtVol SoundFX("fx_Flipperup",DOFContactors), LeftFlipper, VolFlip:LeftFlipper.RotateToEnd:LeftFlipper1.RotateToEnd
     Else
         PlaySoundAtVol SoundFX("fx_Flipperdown",DOFContactors), LeftFlipper, VolFlip:LeftFlipper.RotateToStart:LeftFlipper1.RotateToStart
     End If
  End Sub

Sub SolRFlipper(Enabled)
     If Enabled Then
         PlaySoundAtVol SoundFX("fx_Flipperup",DOFContactors), RightFlipper, VolFlip:RightFlipper.RotateToEnd
     Else
         PlaySoundAtVol SoundFX("fx_Flipperdown",DOFContactors), RightFlipper, VolFlip:RightFlipper.RotateToStart
     End If
End Sub
'**********************************************************************************************************

'Solenoid Controlled toys
'**********************************************************************************************************

Sub SolTrough(Enabled)
	If Enabled Then
		bsTrough.ExitSol_On
		vpmTimer.PulseSw 15
	End If
End Sub

Sub AutoPlunger(Enabled)
    If Enabled Then
        PlungerIM.AutoFire
    End If
End Sub

Sub StLock(enabled)
	if Enabled Then
		Stadiumlock.isDropped = 1
        PlaySound SoundFX("fx_Flipperdown",DOFContactors)
	Else
		Stadiumlock.isDropped = 0
	end If
End Sub

Sub SolBallDeflector(Enabled)
	If Enabled Then
		Deflector.IsDropped=0
        PlaySound SoundFX("fx_Flipperdown",DOFContactors)
	Else
		Deflector.IsDropped=1
	End If
End Sub

set GICallback = GetRef("UpdateGI")
Sub UpdateGI(no, Enabled)
		SetLamp 134, Enabled	'Backwall bulbs
	If Enabled Then
		dim xx
		For each xx in GI:xx.State = 1:	Next
        PlaySound "fx_relay"

	Else For each xx in GI:xx.State = 0: Next
        PlaySound "fx_relay"

	End If
End Sub

'**********************************************************************************************************

'Initiate Table
'**********************************************************************************************************

Dim bsTrough, dtDrop, mGoalie, Magnet1, Magnet2,VLLock

Sub Table1_Init

	vpmInit Me
	On Error Resume Next
		With Controller
		.GameName = cGameName
		If Err Then MsgBox "Can't start Game" & cGameName & vbNewLine & Err.Description : Exit Sub
		.SplashInfoLine = "Striker Xtreme (Stern 2000)"&chr(13)&"Kalavera"
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
	PinMAMETimer.Enabled=1:
	vpmNudge.TiltSwitch=56:
	vpmNudge.Sensitivity=2
	vpmNudge.TiltObj=Array(Bumper1,Bumper2,Bumper3,LeftSlingshot,RightSlingshot)

    Set bsTrough=New cvpmBallStack
		bsTrough.InitSw 0,14,13,12,11,0,0,0
		bsTrough.InitKick BallRelease,95,4
		bsTrough.InitExitSnd SoundFX("ballrelease",DOFContactors), SoundFX("Solenoid",DOFContactors)
		bsTrough.Balls=4

	Set dtDrop=New cvpmDropTarget
		dtDrop.InitDrop Array(sw20,sw19,sw18,sw17),Array(20,19,18,17)
		dtDrop.InitSnd SoundFX("DTDrop",DOFContactors),SoundFX("DTReset",DOFContactors)

    Set Magnet1=New cvpmMagnet
	with Magnet1
		.InitMagnet Trigger8,60
		.CreateEvents "Magnet1"
		.Grabcenter = True
	end With

    Set Magnet2=New cvpmMagnet
	with magnet2
		.InitMagnet RampMag,50
		.Solenoid=13
		.CreateEvents "Magnet2"
		.GrabCenter = 0
	End With

	Set mGoalie=New cvpmMech
		mGoalie.MType=vpmMechOneDirSol+vpmMechReverse+vpmMechLinear
		mGoalie.Sol1=25
		mGoalie.Sol2=26
		mGoalie.Length=20
		mGoalie.Steps=13
		mGoalie.AddSw 41,0,0
		mGoalie.AddSw 47,3,4
		mGoalie.AddSw 42,13,13
		mGoalie.Callback=GetRef("UpdateGoalie")
		mGoalie.Start

	For X=0 To 12:LBPlace(X).IsDropped=1:Next
	wall59.isdropped = 1
	Deflector.IsDropped=1

End Sub

'**********************************************************************************************************
'Plunger code
'**********************************************************************************************************

Sub Table1_KeyDown(ByVal KeyCode)
	'If KeyCode=LeftFlipper Then Controller.Switch(1)=1  'UK only
	'If KeyCode=RightFlipper Then Controller.Switch(8)=1  'UK only
	If keycode = PlungerKey Then Plunger.Pullback:playsoundAtVol"plungerpull", Plunger, 1
	If KeyDownHandler(keycode) Then Exit Sub
End Sub

Sub Table1_KeyUp(ByVal KeyCode)
	'If KeyCode=LeftFlipper Then Controller.Switch(1)=0
	'If KeyCode=RightFlipper Then Controller.Switch(8)=0
	If keycode = PlungerKey Then Plunger.Fire:PlaySoundAtVol"plunger", Plunger, 1
	If KeyUpHandler(keycode) Then Exit Sub
End Sub

	Dim plungerIM
    Const IMPowerSetting = 135 'Plunger Power
    Const IMTime = 1.1       ' Time in seconds for Full Plunge
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
Sub Drain_Hit:bsTrough.addball me : playsoundAtVol"drain" , Drain, 1: End Sub

'GOALIE Subway
Sub sw44a_Hit:vpmTimer.PulseSw 44: playsoundAtVol "subway", ActiveBall, 1: End Sub
Sub sw44b_Hit:vpmTimer.PulseSw 44: playsoundAtVol "subway", ActiveBall, 1: End Sub
Sub sw44c_Hit:vpmTimer.PulseSw 44: playsoundAtVol "subway", ActiveBall, 1: End Sub
Sub sw44d_Hit:vpmTimer.PulseSw 44: playsoundAtVol "subway", ActiveBall, 1: End Sub
Sub sw44e_Hit:vpmTimer.PulseSw 44: playsoundAtVol "subway", ActiveBall, 1: End Sub


 '***********************************
 'Left Raising VUK
 '***********************************
 'Variables used for VUK
 Dim raiseballsw, raiseball
 Sub sw45_Hit()
 	sw45.Enabled=FALSE
	Controller.switch (45) = True
	playsoundAtVol"popper_ball", ActiveBall, 1
 End Sub

 Sub VukLEFTPop(enabled)
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
 		If raiseball.z > 180 then
 			sw45.Kick 180, 3
 			Set raiseball = Nothing
 			sw45raiseballtimer.Enabled = False
 			raiseballsw = False
 		End If
 	End If
 End Sub


 '***********************************
 'Top Raising VUK
 '***********************************
Sub sw46_Hit:Controller.Switch(46)=1: playsoundAtVol "popper_ball", ActiveBall, 1: End Sub

Sub sw46_Timer
	Controller.Switch(46) = 0
	sw46.Timerenabled = 0
End Sub

Sub Kicker5_Timer
	Kicker5.Enabled = 1
	Kicker5.TimerEnabled = 0
End Sub

Sub uVUK(Enabled)
	If Enabled Then
	Kicker5.Enabled = 0
    Kicker5.timerEnabled = 1
	sw46.Kickz 0, 40,89, -20
	'TardisEntrance.KickZ 180, 35, 92, 0
	Playsound SoundFX("Solenoid",DOFContactors)
	sw46.TimerEnabled = 1
	End If
End Sub

 '***********************************
 '***********************************

'Drop Targets
 Sub Sw20_Dropped:dtDrop.Hit 1 :End Sub
 Sub Sw19_Dropped:dtDrop.Hit 2 :End Sub
 Sub Sw18_Dropped:dtDrop.Hit 3 :End Sub
 Sub Sw17_Dropped:dtDrop.Hit 4 :End Sub

'Spinners
Sub sw9_Spin:vpmTimer.PulseSw 9 : playsoundAtVol"fx_spinner" , sw9, 1: End Sub

'Wire Triggers
Sub sw16_Hit:: playsoundAtVol"rollover" , ActiveBall, VolRol: End Sub 'coded to impulse plunger
'Sub sw16_UnHit:Controller.Switch(16) = 0:End Sub
Sub sw31_Hit:Controller.Switch(31)=1 : playsoundAtVol"rollover" , ActiveBall, VolRol: End Sub
Sub sw31_UnHit:Controller.Switch(31)=0:End Sub
Sub sw32_Hit:Controller.Switch(32)=1 : playsoundAtVol"rollover" , ActiveBall, VolRol: End Sub
Sub sw32_UnHit:Controller.Switch(32)=0:End Sub
Sub sw33_Hit:Controller.Switch(33)=1 : playsoundAtVol"rollover" , ActiveBall, VolRol: End Sub
Sub sw33_UnHit:Controller.Switch(33)=0:End Sub
Sub sw34_Hit:Controller.Switch(34)=1 : playsoundAtVol"rollover" , ActiveBall, VolRol: End Sub
Sub sw34_UnHit:Controller.Switch(34)=0:End Sub
Sub sw35_Hit:Controller.Switch(35)=1 : playsoundAtVol"rollover" , ActiveBall, VolRol: End Sub
Sub sw35_UnHit:Controller.Switch(35)=0:End Sub
Sub sw48_Hit:Controller.Switch(48)=1 : playsoundAtVol"rollover" , ActiveBall, VolRol: End Sub
Sub sw48_UnHit:Controller.Switch(48)=0:End Sub
Sub sw57_Hit:Controller.Switch(57)=1 : playsoundAtVol"rollover" , ActiveBall, VolRol: End Sub
Sub sw57_UnHit:Controller.Switch(57)=0:End Sub
Sub sw58_Hit:Controller.Switch(58)=1 : playsoundAtVol"rollover" , ActiveBall, VolRol: End Sub
Sub sw58_UnHit:Controller.Switch(58)=0:End Sub
Sub sw60_Hit:Controller.Switch(60)=1 : playsoundAtVol"rollover" , ActiveBall, VolRol: End Sub
Sub sw60_UnHit:Controller.Switch(60)=0:End Sub
Sub sw61_Hit:Controller.Switch(61)=1 : playsoundAtVol"rollover" , ActiveBall, VolRol: End Sub
Sub sw61_UnHit:Controller.Switch(61)=0:End Sub

'Lock Triggers
Sub SW21_Hit:Controller.Switch(21)=1:End Sub
Sub SW21_unHit:Controller.Switch(21)=0:End Sub
Sub SW22_Hit:Controller.Switch(22)=1:End Sub
Sub SW22_unHit:Controller.Switch(22)=0:End Sub
Sub SW23_Hit:Controller.Switch(23)=1:End Sub
Sub SW23_unHit:Controller.Switch(23)=0:End Sub
Sub SW24_Hit:Controller.Switch(24)=1:End Sub
Sub SW24_unHit:Controller.Switch(24)=0:End Sub

 'Stand Up Targets
Sub sw26_Hit:vpmTimer.PulseSw 26:End Sub
Sub sw27_Hit:vpmTimer.PulseSw 27:End Sub
Sub sw28_Hit:vpmTimer.PulseSw 28:End Sub
Sub sw29_Hit:vpmTimer.PulseSw 29:End Sub
Sub sw52_Hit:vpmTimer.PulseSw 52:End Sub

'Ramp Triggers
Sub sw38_Hit:vpmTimer.pulseSw 38 : End Sub
Sub sw39_Hit:vpmTimer.pulseSw 39 : End Sub

'Goalie Opto Trigger
Sub SW25_Hit:Controller.Switch(25)= 1:End Sub
Sub SW25_unHit:Controller.Switch(25)= 0:End Sub

'Bumpers
Sub Bumper1_Hit:vpmTimer.PulseSw 49 : playsoundAtVol SoundFX("fx_bumper1",DOFContactors), ActiveBall, 1: End Sub
Sub Bumper2_Hit:vpmTimer.PulseSw 50 : playsoundAtVol SoundFX("fx_bumper1",DOFContactors), ActiveBall, 1: End Sub
Sub Bumper3_Hit:vpmTimer.PulseSw 51 : playsoundAtVol SoundFX("fx_bumper1",DOFContactors), ActiveBall, 1: End Sub


 '***********************************
'GOALIE animation
 '***********************************

Dim LBPlace,X
LBPlace=Array(GOALIE0,GOALIE1,GOALIE2,GOALIE3,GOALIE4,GOALIE5,GOALIE6,GOALIE7,GOALIE8,GOALIE9,GOALIE10,GOALIE11,GOALIE12)

Sub UpdateGoalie(aNewPos,aSpeed,aLastPos)
If aNewPos>-1 And aNewPos<13 Then For X=0 To 12:LBPlace(X).IsDropped=1:Next
	Select Case aNewPos
		Case 0:Magnet1.X=153:Magnet1.Y=287:LBPlace(0).IsDropped=0:Goalie.ObjRotZ = 38
		Case 1:Magnet1.X=171:Magnet1.Y=299:LBPlace(1).IsDropped=0:Goalie.ObjRotZ = 26
		Case 2:Magnet1.X=187:Magnet1.Y=303:LBPlace(2).IsDropped=0:Goalie.ObjRotZ = 15
		Case 3:Magnet1.X=213:Magnet1.Y=304:LBPlace(3).IsDropped=0:Goalie.ObjRotZ = 0
		Case 4:Magnet1.X=230:Magnet1.Y=295:LBPlace(4).IsDropped=0:Goalie.ObjRotZ = -15
		Case 5:Magnet1.X=245:Magnet1.Y=295:LBPlace(5).IsDropped=0:Goalie.ObjRotZ = -26
		Case 6:Magnet1.X=261:Magnet1.Y=285:LBPlace(6).IsDropped=0:Goalie.ObjRotZ = -38
		Case 7:Magnet1.X=283:Magnet1.Y=268:LBPlace(7).IsDropped=0:Goalie.ObjRotZ = -48
		Case 8:Magnet1.X=294:Magnet1.Y=254:LBPlace(8).IsDropped=0:Goalie.ObjRotZ = -55
		Case 9:Magnet1.X=302:Magnet1.Y=242:LBPlace(9).IsDropped=0:Goalie.ObjRotZ = -65
		Case 10:Magnet1.X=310:Magnet1.Y=230:LBPlace(10).IsDropped=0:Goalie.ObjRotZ = -75
		Case 11:Magnet1.X=313:Magnet1.Y=213:LBPlace(11).IsDropped=0:Goalie.ObjRotZ = -85
		Case 12:Magnet1.X=312:Magnet1.Y=198:LBPlace(12).IsDropped=0:Goalie.ObjRotZ = -95
	End Select
	Magnet1.Size=75
End Sub


Sub SW43_Hit:Controller.Switch(43)=1:End Sub 'switch 43 is goalie magnet
Sub SW43_unHit:Controller.Switch(43)=0:End Sub
Sub Trigger8_Hit:Magnet1.AddBall ActiveBall:End Sub
Sub Trigger8_UnHit:Magnet1.RemoveBall ActiveBall:End Sub


'Generic Sounds
Sub Trigger1_Hit:PlaySoundAtVol "fx_ballrampdrop", ActiveBall, 1:End Sub
Sub Trigger2_Hit:stopSound "Wire Ramp":PlaySoundAtVol "fx_ballrampdrop", ActiveBall, 1:End Sub
Sub Trigger3_Hit:stopSound "Wire Ramp":PlaySoundAtVol "fx_ballrampdrop", ActiveBall, 1:End Sub
Sub Trigger6_Hit:PlaySoundAtVol "fx_ballrampdrop", ActiveBall, 1:End Sub

Sub Trigger4_Hit:PlaySoundAtVol "Wire Ramp", ActiveBall, 1:End Sub
Sub Trigger5_Hit:PlaySoundAtVol "Wire Ramp", ActiveBall, 1:End Sub

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

	   'Special Handling
	   'If chgLamp(ii,0) = 2 Then solTrough chgLamp(ii,1)
	   'If chgLamp(ii,0) = 4 Then PFGI chgLamp(ii,1)

        Next
    End If
    UpdateLamps
End Sub


Sub UpdateLamps()
NFadeL 1, L1
NFadeL 2, L2
NFadeL 3, L3
NFadeL 4, L4
NFadeL 5, L5
NFadeL 6, L6
NFadeL 7, L7
NFadeL 8, L8
NFadeL 9, L9
NFadeL 10, L10
NFadeL 11, L11
NFadeL 12, L12
NFadeL 13, L13
NFadeL 14, L14
NFadeL 15, L15
NFadeL 16, L16
NFadeL 17, L17
NFadeL 18, L18
NFadeL 19, L19
NFadeL 20, L20
NFadeL 21, L21
NFadeL 22, L22
NFadeL 23, L23
NFadeL 24, L24
Flash 25, F25 'Backwall
Flash 26, F26 'Backwall
Flash 27, F27 'Backwall

NFadeL 29, L29
NFadeL 30, L30
NFadeL 31, L31
NFadeL 32, L32
NFadeL 33, L33
NFadeL 34, L34
NFadeL 35, L35

NFadeLm 38, L38 'Bumper 1
NFadeL 38, L38b
NFadeLm 39, L39 'Bumper 2
NFadeL 39, L39b
NFadeLm 40, L40 'Bumper 3
NFadeL 40, L40b

NFadeL 42, L42
NFadeL 43, L43
NFadeL 44, L44
NFadeL 45, L45
NFadeL 46, L46
NFadeL 47, L47
NFadeL 48, L48
NFadeL 49, L49
NFadeL 50, L50
NFadeL 51, L51
NFadeL 52, L52
NFadeL 53, L53
NFadeL 54, L54
NFadeL 55, L55
NFadeL 56, L56
NFadeL 57, L57
NFadeL 58, L58
NFadeL 59, L59
NFadeL 60, L60
NFadeL 61, L61

NFadeL 64, L64
NFadeL 65, L65
NFadeL 66, L66
NFadeL 67, L67
NFadeL 68, L68
NFadeL 69, L69
NFadeL 70, L70
NFadeL 71, L71
NFadeL 72, L72

'Solenoid Controlled

NFadeLm 120, S20a
NFadeLm 120, S20b
NFadeLm 120, S20c
NFadeL 120, S20d

NFadeObjm 127, p127, "dome2_0_yellowON", "dome2_0_yellow"
NFadeLm 127, S27 'Dome
Flash 127, S27a 'Dome

NFadeObjm 128, p128, "dome2_0_redON", "dome2_0_red"
NFadeLm 128, S28 'Dome
Flash 128, S28a 'Dome

NFadeObjm 129, p129, "dome3_clearON", "dome3_clear"
NFadeLm 129, S29 'Dome
Flash 129, S29a 'Dome

Flash 130, S30 'BackPanel Goal

NFadeLm 131, S31a
NFadeLm 131, S31b
NFadeLm 131, S31c
NFadeL 131, S31d

NFadeLm 132, S32a
NFadeLm 132, S32b
NFadeLm 132, S32c
NFadeL 132, S32d

'Backwall Bulbs controlled by GI string

NFadeLm 134, GI1
NFadeLm 134, GI2
NFadeLm 134, GI3
NFadeLm 134, GI4
NFadeLm 134, GI5
NFadeLm 134, GI6
NFadeLm 134, GI7
NFadeLm 134, GI8
NFadeLm 134, GI9
NFadeLm 134, GI10
NFadeLm 134, GI11
NFadeLm 134, GI12
NFadeLm 134, GI13
NFadeLm 134, GI14
NFadeLm 134, GI15
NFadeLm 134, GI16
NFadeLm 134, GI17
NFadeLm 134, GI18
NFadeLm 134, GI19
NFadeLm 134, GI20
NFadeLm 134, GI21
NFadeLm 134, GI22
NFadeLm 134, GI23
Flashm 134, Flasher1
Flashm 134, Flasher2
Flashm 134, Flasher3
Flashm 134, Flasher4
Flashm 134, Flasher5
Flashm 134, Flasher6
Flashm 134, Flasher7
Flashm 134, Flasher8
Flashm 134, Flasher9
Flash 134, Flasher10

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
	FlipperLSh1.RotZ = LeftFlipper1.currentangle

    LFLogo.RotY = LeftFlipper.CurrentAngle
    RFlogo.RotY = RightFlipper.CurrentAngle
    LFLogo1.RotY = LeftFlipper1.CurrentAngle


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
		PlaySound "fx_rubber2", 0, Vol(ActiveBall)*VolPo, AudioPan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
	End if
	If finalspeed >= 6 AND finalspeed <= 16 then
 		RandomSoundRubber()
 	End If
End Sub

Sub RandomSoundRubber()
	Select Case Int(Rnd*3)+1
		Case 1 : PlaySound "rubber_hit_1", 0, Vol(ActiveBall)*VolRH, AudioPan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
		Case 2 : PlaySound "rubber_hit_2", 0, Vol(ActiveBall)*VolRH, AudioPan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
		Case 3 : PlaySound "rubber_hit_3", 0, Vol(ActiveBall)*VolRH, AudioPan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
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


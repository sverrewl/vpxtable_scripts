'NFL Stern 2001 for VPX by Sliderpoint

' Thalamus 2018-07-24
' Added/Updated "Positional Sound Playback Functions" and "Supporting Ball & Sound Functions"
' Changed UseSolenoids=1 to 2
' No special SSF tweaks yet.

Option Explicit
Randomize

On Error Resume Next
ExecuteGlobal GetTextFile("controller.vbs")
If Err Then MsgBox "You need the controller.vbs in order to run this table, available in the vp10 package"
On Error Goto 0

Sub LoadCoreVBS
 On Error Resume Next
 ExecuteGlobal GetTextFile("core.vbs")
 If Err Then MsgBox "Can't open core.vbs"
 On Error Goto 0
End Sub

'*******table option
toggleModSounds = 0 '1 extra sound clips are on, 0 are off
'*********

Dim DesktopMode: DesktopMode = Table1.ShowDT

DIM UseVPMDMD
UseVPMDMD = DesktopMode

LoadVPM "01520000","Sega.VBS",3.02

Dim toggleModSounds
Function ModSound(sound)
	If toggleModSounds = 0 Then
		ModSound = ""
	Else
		ModSound = sound
	End If
End Function

If toggleModSounds = 1 Then
	PlayMusic "NFL Seahawks 01.mp3"
end If

Dim bsTrough, Magnet1,VLLock,dtDrop,mGoalie,Magnet2

Const UseSolenoids=2,UseLamps=1,UseSync=1, SCoin="coin3"

Sub Table1_Init
	Plunger1.Pullback
    LP.IsDropped=1
    MP.IsDropped=1
    RP.IsDropped=1
	Deflector.IsDropped=1
	For X=0 To 12:LBPlace(X).IsDropped=1:Next
	vpmInit Me
	Controller.GameName="strikext"
	NVOffset (2)
    Controller.Games("strikext").Settings.Value("rol")=0
    Controller.Games("strikext").Settings.Value("ror")=0
	Controller.SplashInfoLine="Striker Xtreme (Stern 2000)"&vbNewLine&"VPX table by Sliderpoint"&vbNewLine&"MOD by Kalavera"
	Controller.HandleKeyboard=0
	Controller.ShowTitle=0
	Controller.ShowDMDOnly=1
	Controller.ShowFrame=0
	Controller.HandleMechanics=0
	On Error Resume Next
		Controller.Run GetPlayerHwnd
		If Err Then MsgBox Err.Description
	On Error Goto 0

	vpmNudge.TiltSwitch=56
	vpmNudge.Sensitivity=5
	vpmNudge.TiltObj=Array(Bumper1,Bumper2,Bumper3,LeftSlingshot,RightSlingshot)

    Set bsTrough=New cvpmBallStack
 	bsTrough.InitSw 0,14,13,12,11,0,0,0
 	bsTrough.InitKick BallRelease,95,4
	bsTrough.InitEntrySnd "Solenoid","Solenoid"
	bsTrough.InitExitSnd "BallRel","Solenoid"
	bsTrough.Balls=4

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

	Set dtDrop=New cvpmDropTarget
	dtDrop.InitDrop Array(Drop1,Drop2,Drop3,Drop4),Array(20,19,18,17)
	dtDrop.InitSnd "flapclos","flapopen"
	dtDrop.CreateEvents "dtDrop"

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

	vpmMapLights AllLights

	DayNight = table1.NightDay
	Intensity 'sets GI brightness depending on day/night slider settings


	Dim iw
		If DesktopMode = True Then
			For each iw in BGstuff: iw.Visible = True:Next
		Else
			For each iw in BGstuff: iw.Visible = False:Next
		End If
	wall59.isdropped = 1

End Sub

Dim FlashLevel1, FlashLevel2, FlashLevel3, FlashLevel4
FlasherLight1.IntensityScale = 0
Flasherlight2.IntensityScale = 0
Flasherlight3.IntensityScale = 0
Flasherlight4.IntensityScale = 0

'KEYS
Sub Table1_KeyDown(ByVal KeyCode)
	If KeyCode=50 Then Controller.Switch(1)=0
	If KeyCode=50 Then Controller.Switch(8)=0
	If Keycode =50  Then
		If toggleModSounds = 1 Then
           ToggleModSounds = 0:EndMusic
		Else
          ToggleModSounds = 1
		End If
	End if
	If KeyCode = 6 Then
		Playsound ModSound("NFL Seahawks 00")
	End If
	If KeyCode = 4 Then
		Playsound ModSound("NFL Seahawks 00")
	End If
	If KeyDownHandler(KeyCode) Then Exit Sub
	If KeyCode=PlungerKey Then
		PlaySound"PlungerPull"
		Plunger.Pullback
	End If

End Sub

Sub Table1_KeyUp(ByVal KeyCode)
	If KeyCode=55 Then Controller.Switch(1)=0
	If KeyCode=55 Then Controller.Switch(8)=0
	If KeyUpHandler(KeyCode) Then Exit Sub
	If KeyCode=PlungerKey Then
		PlaySound"Plunger"
        PlaySound ModSound("football grunt 03")
			if toggleModSounds = 1 Then
			Dim x
			x = INT(3 * RND(1) )
			Select Case x
			Case 0:PlayMusic "NFL Seahawks 01.mp3"
			Case 1:PlayMusic "NFL Seahawks 02.mp3"
			Case 2:PlayMusic "NFL Seahawks 03.mp3"
			End Select
			end if
     End If
            Plunger.Fire
End Sub

'SOLENOIDS
set GICallback = GetRef("UpdateGI")

SolCallback(1)="SolTrough"
SolCallback(2)="SolShooter"
SolCallback(3)="pVUK"
SolCallback(4)="uVUK"
SolCallback(5)="vpmSolSound ""Sling"","
SolCallback(6)="dtDrop.SolDropUp"
SolCallback(7)="dtDrop.SolHit 1,"
SolCallback(8)="StLock"
SolCallback(9)="vpmSolSound ""Jet3"","
SolCallback(10)="vpmSolSound ""Jet3"","
SolCallback(11)="vpmSolSound ""Jet3"","
SolCallback(12) = "Magnet1.MagnetOn="
SolCallback(sLLFlipper)="SolLFlipper"
SolCallback(sLRFlipper)="SolRFlipper"
SolCallback(17)="vpmSolSound ""Sling"","
SolCallback(18)="vpmSolSound ""Sling"","
SolCallback(19)="SolBallDeflector"
SolCallback(20)="SFF" ' flasher under the stadium seats
SolCallback(21)="dtDrop.SolHit 2,"
SolCallback(22)="dtDrop.SolHit 3,"
SolCallback(23)="dtDrop.SolHit 4,"
'24=coin meter optional
SolCallBack(27)="UFF" 'left flasher cap
SolCallBack(28)="Flasher28" '"vpmFlasher Flasher6," 'left ramp flasher cap
SolCallBack(29)="flasher29" ' "vpmFlasher Flasher7," 'upper flasher cap
SolCallBack(30)="flasher30" ' "vpmFlasher Array(FlasherZ1,FlasherZ2,FlasherZ3,FlasherZ4)," 'Back Panel Flashers x4
SolCallback(31)="PopsFlash"
SolCallback(32)="SlingFlash"
SolCallback(33)="SolLeftPost" ' AUX1 not sure these are used in this game
SolCallback(34)="SolMidPost" ' AUX2 not sure these are used in this game
SolCallback(35)="SolRightPost" ' AUX3 not sure these are used in this game

Sub SolLFlipper(Enabled)
    If Enabled Then
        PlaySound SoundFX("FlipperUp",DOFFlippers):LeftFlipper.RotateToEnd:Flipper1.RotateToEnd
    Else
        PlaySound SoundFX("FlipperDown",DOFFlippers):LeftFlipper.RotateToStart:Flipper1.RotateToStart
    End If
End Sub

Sub SolRFlipper(Enabled)
    If Enabled Then
        PlaySound SoundFX("FlipperUp",DOFFlippers):RightFlipper.RotateToEnd
    Else
        PlaySound SoundFX("FlipperDown",DOFFlippers):RightFlipper.RotateToStart
    End If
End Sub

Sub SolLeftPost(Enabled)
	If Enabled Then
		LP.IsDropped=0
	Else
		LP.IsDropped=1
	End If
End Sub
Sub SolRightPost(Enabled)
	If Enabled Then
		RP.IsDropped=0
	Else
		RP.IsDropped=1
	End If
End Sub
Sub SolMidPost(Enabled)
	If Enabled Then
		MP.IsDropped=0
	Else
		MP.IsDropped=1
	End If
End Sub

'GI - add any new GI lights to the collection "GI" to control them together.
Dim ig

Sub UpdateGI(no, Enabled)
	Select Case no
		Case 0 'Top
			If Enabled Then
				For each ig in GI:ig.State = 1:Next
				For each ig in GIbulbs:ig.State = 1:Next
			Else
				For each ig in GI:ig.State = 0:Next
				For each ig in GIbulbs:ig.State = 0:Next
			End If
	End Select
End Sub


Dim GILevel, DayNight, xx

Sub Intensity
	If DayNight <= 20 Then
			GILevel = .5
	ElseIf DayNight <= 40 Then
			GILevel = .4125
	ElseIf DayNight <= 60 Then
			GILevel = .325
	ElseIf DayNight <= 80 Then
			GILevel = .2375
	Elseif DayNight <= 100  Then
			GILevel = .15
	End If

	For each xx in GI: xx.IntensityScale = xx.IntensityScale * (GILevel*2): Next
	For each xx in GIbulbs: xx.IntensityScale = xx.IntensityScale * (GILevel*2): Next
End Sub
' end GI subs

'*** white flasher ***
Sub FlasherFlash1_Timer()
	dim flashx3, matdim
	If not FlasherFlash1.TimerEnabled Then
		FlasherFlash1.TimerEnabled = True
		FlasherFlash1.visible = 1
		FlasherLit1.visible = 1
	End If
	flashx3 = FlashLevel1 * FlashLevel1 * FlashLevel1
	Flasherflash1.opacity = 1000 * flashx3
	FlasherLit1.BlendDisableLighting = 10 * flashx3
	Flasherbase1.BlendDisableLighting =  flashx3
	FlasherLight1.IntensityScale = flashx3
	matdim = Round(10 * FlashLevel1)
	FlasherLit1.material = "domelit" & matdim
	FlashLevel1 = FlashLevel1 * 0.85 - 0.01
	If FlashLevel1 < 0.15 Then
		FlasherLit1.visible = 0
	Else
		FlasherLit1.visible = 1
	end If
	If FlashLevel1 < 0 Then
		FlasherFlash1.TimerEnabled = False
		FlasherFlash1.visible = 0
	End If
End Sub

'*** Red flasher ***
Sub FlasherFlash2_Timer()
	dim flashx3, matdim
	If not Flasherflash2.TimerEnabled Then
		Flasherflash2.TimerEnabled = True
		Flasherflash2.visible = 1
		Flasherlit2.visible = 1
	End If
	flashx3 = FlashLevel2 * FlashLevel2 * FlashLevel2
	Flasherflash2.opacity = 1500 * flashx3
	Flasherlit2.BlendDisableLighting = 10 * flashx3
	Flasherbase2.BlendDisableLighting =  flashx3
	Flasherlight2.IntensityScale = flashx3
	matdim = Round(10 * FlashLevel2)
	Flasherlit2.material = "domelit" & matdim
	FlashLevel2 = FlashLevel2 * 0.9 - 0.01
	If FlashLevel2 < 0.15 Then
		Flasherlit2.visible = 0
	Else
		Flasherlit2.visible = 1
	end If
	If FlashLevel2 < 0 Then
		Flasherflash2.TimerEnabled = False
		Flasherflash2.visible = 0
	End If
End Sub

'*** blue flasher ***
Sub FlasherFlash3_Timer()
	dim flashx3, matdim
	If not Flasherflash3.TimerEnabled Then
		Flasherflash3.TimerEnabled = True
		Flasherflash3.visible = 1
		Flasherlit3.visible = 1
	End If
	flashx3 = FlashLevel3 * FlashLevel3 * FlashLevel3
	Flasherflash3.opacity = 8000 * flashx3
	Flasherlit3.BlendDisableLighting = 10 * flashx3
	Flasherbase3.BlendDisableLighting =  flashx3
	Flasherlight3.IntensityScale = flashx3
	matdim = Round(10 * FlashLevel3)
	Flasherlit3.material = "domelit" & matdim
	FlashLevel3 = FlashLevel3 * 0.85 - 0.01
	If FlashLevel3 < 0.15 Then
		Flasherlit3.visible = 0
	Else
		Flasherlit3.visible = 1
	end If
	If FlashLevel3 < 0 Then
		Flasherflash3.TimerEnabled = False
		Flasherflash3.visible = 0
	End If
End Sub

'*** blue flasher vertical (script is the same as for blue Flasher) ***
Sub FlasherFlash4_Timer()
	dim flashx3, matdim
	If not Flasherflash4.TimerEnabled Then
		Flasherflash4.TimerEnabled = True
		Flasherflash4.visible = 1
		Flasherlit4.visible = 1
	End If
	flashx3 = FlashLevel4 * FlashLevel4 * FlashLevel4
	Flasherflash4.opacity = 8000 * flashx3
	Flasherlit4.BlendDisableLighting = 10 * flashx3
	Flasherbase4.BlendDisableLighting =  flashx3
	Flasherlight4.IntensityScale = flashx3
	matdim = Round(10 * FlashLevel4)
	Flasherlit4.material = "domelit" & matdim
	FlashLevel4 = FlashLevel4 * 0.85 - 0.01
	If FlashLevel4 < 0.15 Then
		Flasherlit4.visible = 0
	Else
		Flasherlit4.visible = 1
	end If
	If FlashLevel4 < 0 Then
		Flasherflash4.TimerEnabled = False
		Flasherflash4.visible = 0
	End If
End Sub



Sub StLock(enabled)
	if Enabled Then
	Stadiumlock.isDropped = 1
	Else
	Stadiumlock.isDropped = 0
	end If
End Sub

Sub pVUK(Enabled)
	If Enabled Then
	Kicker1.Kickz 180, 25, 91, 0
	Wall59.isDropped = 0
	Playsound SoundFX("Solenoid",DOFContactors)
	Kicker1.TimerEnabled = 1
	Controller.Switch(45) = 0
	End If
End Sub

Sub uVUK(Enabled)
	If Enabled Then
	Kicker5.Enabled = 0
    Kicker5.timerEnabled = 1
	Kicker2.Kickz 0, 40,89, -20
'		TardisEntrance.KickZ 180, 35, 92, 0
	Playsound SoundFX("Solenoid",DOFContactors)
	Kicker2.TimerEnabled = 1
	End If
End Sub

Sub Kicker5_Timer
	Kicker5.Enabled = 1
	Kicker5.TimerEnabled = 0
End Sub

Sub Kicker1_Timer
'	Controller.Switch(45) = 0
	Wall59.isdropped = 1
	Kicker1.Timerenabled = 0
End Sub

Sub Kicker2_Timer
	Controller.Switch(46) = 0
	Kicker2.Timerenabled = 0
End Sub

Sub SFF(Enabled)
	If Enabled Then
	Light20a.state=1
	Light20b.state=1
	Light20c.state=1
	Light20d.state=1
	Else
	Light20a.state=0
	Light20b.state=0
	Light20c.state=0
	Light20d.state=0
	End If
End Sub

Sub UFF(Enabled)
	If Enabled Then
	FlashLevel1 = 1: FlasherFlash1_Timer
	Else
	FlashLevel1 = 0: FlasherFlash1_Timer
	End If
End Sub

Sub SlingFlash(Enabled)
	If Enabled Then
	Light28.state = 1
	Light36.State = 1
	Light77.State = 1
	Light78.State = 1
	Else
	Light28.state = 0
	Light36.State = 0
	Light77.State = 0
	Light78.State = 0
	End If
End Sub

Sub PopsFlash(enabled)
	If Enabled Then
		Light75.state = 1
		Light76.state = 1
		Light79.State = 1
		Light80.State = 1
	Else
		Light75.state = 0
		Light76.State = 0
		Light79.State = 0
		Light80.State = 0
	End If
End Sub

Sub Flasher30(enabled)
	If Enabled Then
		Light41.state = 1
	Else
		Light41.state = 0
	End If
End Sub

Sub Flasher29(enabled)
	If Enabled Then
		FlashLevel3 = 1: FlasherFlash3_Timer
	Else
		FlashLevel3 = 0: FlasherFlash3_Timer
	End If
End Sub

Sub Flasher28(enabled)
	If Enabled Then
	FlashLevel2 = 1: FlasherFlash2_Timer
	Else
	FlashLevel2 = 0: FlasherFlash2_Timer
	End If
End Sub

Sub LightTimer_Timer
	F25.visible = Light25.State
	F26.visible = Light26.State
	F27.visible = Light27.State
	F4.visible = Light41.State
End Sub

Sub SolBallDeflector(Enabled)
	If Enabled Then
		Deflector.IsDropped=0
	Else
		Deflector.IsDropped=1
	End If
End Sub

Sub SolTrough(Enabled)
	If Enabled Then
		bsTrough.ExitSol_On
		vpmTimer.PulseSw 15
	End If
End Sub

Sub SolShooter(Enabled)
    If Enabled Then
	PlaySound SoundFX("SolOn",DOFContactors)
	Plunger1.fire
	Plunger1.pullback
	end if
End Sub

'GOALIE
Dim LBPlace,X
LBPlace=Array(L0,L1,L2,L3,L4,L5,L6,L7,L8,L9,L10,L11,L12)

Sub UpdateGoalie(aNewPos,aSpeed,aLastPos)
If aNewPos>-1 And aNewPos<12 Then For X=0 To 12:LBPlace(X).IsDropped=1:Next
	Select Case aNewPos
		Case 0:Magnet1.X=159:Magnet1.Y=282:LBPlace(0).IsDropped=0:Goalie.ObjRotZ = 38
		Case 1:Magnet1.X=175:Magnet1.Y=293:LBPlace(1).IsDropped=0:Goalie.ObjRotZ = 26
		Case 2:Magnet1.X=194:Magnet1.Y=302:LBPlace(2).IsDropped=0:Goalie.ObjRotZ = 15
		Case 3:Magnet1.X=212:Magnet1.Y=304:LBPlace(3).IsDropped=0:Goalie.ObjRotZ = 0
		Case 4:Magnet1.X=234:Magnet1.Y=303:LBPlace(4).IsDropped=0:Goalie.ObjRotZ = -15
		Case 5:Magnet1.X=252:Magnet1.Y=295:LBPlace(5).IsDropped=0:Goalie.ObjRotZ = -26
		Case 6:Magnet1.X=272:Magnet1.Y=287:LBPlace(6).IsDropped=0:Goalie.ObjRotZ = -38
		Case 7:Magnet1.X=286:Magnet1.Y=274:LBPlace(7).IsDropped=0:Goalie.ObjRotZ = -48
		Case 8:Magnet1.X=297:Magnet1.Y=262:LBPlace(8).IsDropped=0:Goalie.ObjRotZ = -55
		Case 9:Magnet1.X=305:Magnet1.Y=246:LBPlace(9).IsDropped=0:Goalie.ObjRotZ = -65
		Case 10:Magnet1.X=310:Magnet1.Y=230:LBPlace(10).IsDropped=0:Goalie.ObjRotZ = -75
		Case 11:Magnet1.X=313:Magnet1.Y=213:LBPlace(11).IsDropped=0:Goalie.ObjRotZ = -85
		Case 12:Magnet1.X=312:Magnet1.Y=198:LBPlace(12).IsDropped=0:Goalie.ObjRotZ = -95
	End Select
	Magnet1.Size=75
End Sub

Sub Spinner1_Spin:vpmTimer.PulseSw 9:End Sub
Sub Drain_Hit:bsTrough.AddBall Me
	If toggleModSounds = 1 Then
		Dim x
		x = INT(16 * RND(1) )
		Select Case x
		Case 0:PlayMusic "NFL Seahawks 04.mp3"
		Case 1:PlayMusic "NFL Seahawks 05.mp3"
		Case 2:PlayMusic "NFL Seahawks 06.mp3"
		Case 3:PlayMusic "NFL Seahawks 07.mp3"
		Case 4:PlayMusic "NFL Seahawks 08.mp3"
		Case 5:PlayMusic "NFL Seahawks 09.mp3"
		Case 6:PlayMusic "NFL Seahawks 10.mp3"
		Case 7:PlayMusic "NFL Seahawks 11.mp3"
		Case 8:PlayMusic "NFL Seahawks 12.mp3"
		Case 9:PlayMusic "NFL Seahawks 13.mp3"
		Case 10:PlayMusic "NFL Seahawks 14.mp3"
		Case 11:PlayMusic "NFL Seahawks 15.mp3"
		Case 12:PlayMusic "NFL Seahawks 16.mp3"
		Case 13:PlayMusic "NFL Seahawks 17.mp3"
		Case 14:PlayMusic "NFL Seahawks 18.mp3"
		Case 15:PlayMusic "NFL Seahawks 19.mp3"
		End Select
	End If
End Sub

Sub Trigger2_Hit:Controller.Switch(16)=1:End Sub
Sub Trigger2_UnHit:Controller.Switch(16)=0:End Sub
Sub SW21_Hit:Controller.Switch(21)=1:End Sub
Sub SW21_unHit:Controller.Switch(21)=0:End Sub
Sub SW22_Hit:Controller.Switch(22)=1:End Sub
Sub SW22_unHit:Controller.Switch(22)=0:End Sub
Sub SW23_Hit:Controller.Switch(23)=1:End Sub
Sub SW23_unHit:Controller.Switch(23)=0:End Sub
Sub SW24_Hit:Controller.Switch(24)=1:End Sub
Sub SW24_unHit:Controller.Switch(24)=0:End Sub
Sub SW25_Hit:Controller.Switch(25)= 1:End Sub
Sub SW25_unHit:Controller.Switch(25)= 0:End Sub
Sub T26_Hit:vpmTimer.PulseSw 26:End Sub
Sub T27_Hit:vpmTimer.PulseSw 27:End Sub
Sub T28_Hit:vpmTimer.PulseSw 28:End Sub
Sub T29_Hit:vpmTimer.PulseSw 29:End Sub
Sub Trigger4_Hit:Controller.Switch(31)=1:End Sub
Sub Trigger4_UnHit:Controller.Switch(31)=0:End Sub
Sub Trigger1_Hit:Controller.Switch(32)=1:End Sub
Sub Trigger1_UnHit:Controller.Switch(32)=0:End Sub
Sub A_Hit:Controller.Switch(33)=1:End Sub
Sub A_UnHit:Controller.Switch(33)=0:End Sub
Sub B_Hit:Controller.Switch(34)=1:End Sub
Sub B_UnHit:Controller.Switch(34)=0:End Sub
Sub C_Hit:Controller.Switch(35)=1:End Sub
Sub C_UnHit:Controller.Switch(35)=0:End Sub
Sub Gate1_Hit:vpmTimer.PulseSw(38):End Sub
Sub Gate39_Hit:vpmTimer.PulseSw (39): End Sub
Sub SW43_Hit:Controller.Switch(43)=1:End Sub 'switch 43 is goalie magnet
Sub SW43_unHit:Controller.Switch(43)=0:End Sub
Sub Trigger8_Hit:Magnet1.AddBall ActiveBall:End Sub
Sub Trigger8_UnHit:Magnet1.RemoveBall ActiveBall:End Sub
Sub Kicker7_Hit:vpmTimer.PulseSw 44:PlaySound ModSound("NFL Seahawks 00a"):End Sub
Sub Kicker8_Hit:vpmTimer.PulseSw 44:PlaySound ModSound("NFL Seahawks 00b"):End Sub
Sub kicker9_Hit:vpmTimer.PulseSw 44:PlaySound ModSound("NFL Seahawks 00a"):End Sub
Sub Kicker10_Hit:vpmTimer.PulseSw 44:PlaySound ModSound("NFL Seahawks 00b"):End Sub
Sub Kicker4_Hit:vpmTimer.PulseSw 44:PlaySound ModSound("NFL Seahawks 00a"):End Sub
Sub Kicker1_Hit:Controller.Switch(45)=1:Playsound ModSound("Defense Chant"):End Sub
Sub Kicker2_Hit:Controller.Switch(46)=1:End Sub
Sub Trigger3_Hit:Controller.Switch(48)=1:End Sub
Sub Trigger3_UnHit:Controller.Switch(48)=0:End Sub
Sub Bumper1_Hit:vpmTimer.PulseSw 49:End Sub
Sub Bumper2_Hit:vpmTimer.PulseSw 50:End Sub
Sub Bumper3_Hit:vpmTimer.PulseSw 51:End Sub
Sub Target1_Hit:vpmTimer.PulseSw 52:End Sub
Sub LeftOutlane_Hit:Controller.Switch(57)=1:PlaySound ModSound("Whistle"):End Sub
Sub LeftOutlane_UnHit:Controller.Switch(57)=0:End Sub
Sub LeftInlane_Hit:Controller.Switch(58)=1:End Sub
Sub LeftInLane_UnHit:Controller.Switch(58)=0:End Sub
Sub LeftSlingShot_SlingShot:vpmTimer.PulseSw 59:PlaySound ModSound("football grunt 01"):End Sub
Sub RightOutlane_Hit:Controller.Switch(60)=1:PlaySound ModSound("Whistle"):End Sub
Sub RightOutlane_UnHit:Controller.Switch(60)=0:End Sub
Sub RightInlane_Hit:Controller.Switch(61)=1:End Sub
Sub RightInlane_UnHit:Controller.Switch(61)=0:End Sub
Sub RightSlingShot_SlingShot:vpmTimer.PulseSw 62:PlaySound ModSound("football grunt 03"):End Sub
Sub RampMag_Hit():Magnet2.addball activeball:End Sub
Sub RampMag_Unhit():Magnet2.removeball activeball:activeball.vely = activeball.vely * 1.5:End Sub

'	Controller.switch(54) = 1' start button
'	Controller.switch(55) = 1' slam tilt
'	Controller.switch(56) = 1' plumbbob tilt

'MISC SOUNDS
Sub Rubbers_Hit(idx)
 	dim finalspeed
  	finalspeed=SQR(activeball.velx * activeball.velx + activeball.vely * activeball.vely)
 	If finalspeed > 20 then
		PlaySound "rubber", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0
	End if
	If finalspeed >= 6 AND finalspeed <= 20 then
 		RandomSoundRubber()
 	End If
End Sub

Sub RandomSoundRubber()
	Select Case Int(Rnd*3)+1
		Case 1 : PlaySound "rubber_hit_1", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
		Case 2 : PlaySound "rubber_hit_2", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
		Case 3 : PlaySound "rubber_hit_3", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
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
		Case 1 : PlaySound "flip_hit_1", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
		Case 2 : PlaySound "flip_hit_2", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
		Case 3 : PlaySound "flip_hit_3", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
	End Select
End Sub

'*****************************************
'	ninuzzu's	FLIPPER SHADOWS
'*****************************************

sub FlipperTimer_Timer()
	Flipperizq1.RotZ = LeftFlipper.currentangle
	Flipperder.RotZ = RightFlipper.currentangle
	Flipperizq2.RotZ = RightFlipper1.currentangle
End Sub


'*****************************************
'	ninuzzu's	BALL SHADOW
'*****************************************
Dim Sombrabola
Sombrabola = Array (Sombrabola1,Sombrabola2,Sombrabola3,Sombrabola4,Sombrabola5)

Sub Sombrabola_timer()
    Dim BOT, b
    BOT = GetBalls
    ' hide shadow of deleted balls
    If UBound(BOT)<(tnob-1) Then
        For b = (UBound(BOT) + 1) to (tnob-1)
            Sombrabola(b).visible = 0
        Next
    End If
    ' exit the Sub if no balls on the table
    If UBound(BOT) = -1 Then Exit Sub
    ' render the shadow for each ball
    For b = 0 to UBound(BOT)
        If BOT(b).X < Table1.Width/2 Then
            Sombrabola(b).X = ((BOT(b).X) - (Ballsize/6) + ((BOT(b).X - (Table1.Width/2))/7)) + 6
        Else
            Sombrabola(b).X = ((BOT(b).X) + (Ballsize/6) + ((BOT(b).X - (Table1.Width/2))/7)) - 6
        End If
        ballShadow(b).Y = BOT(b).Y + 12
        If BOT(b).Z > 20 Then
            Sombrabola(b).visible = 1
        Else
            Sombrabola(b).visible = 0
        End If
    Next
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


'
'           MMMMMMMM
'          MMMMMMMMMM
'        MMMMMMMMMMMM
'      .MMMMMMM   IMM
'     .MMMMMMM   . 8
'    .MMMMMMM.                                          . .
'    .MMMMMMM                      ,MMMM. .D+ .      .MMMMM .
'    .8MMMMMM.  ..M.              MMMMMMM..OMMMMZ   MMMMMMMMM
'      MMMMMM. .$MMM    ... 7$.  MMMMZ.MMM MMMMMM OMMMMMMMMMMM
'      :MMMMMM .MMM . . .MMMMMM.MMMM .,MMO.MM8MMM.MMMMM   .MMM.
'       MMMMMM$ MM .MM MMMMMMMO MMMM. MMM ,MM..MM.M7MMM~    MMM
'   .    MMMMMM$  OMMM  MMMM8M.  MMM DMM  ,MM .MMM   MMM . .MMM
'    MM . MMMMMM. MMM8  MMM8.M, .7MMDMM   .MMMMMMM. .MMM.  .MMM.
'    MMM, ..MMMMM .MMM. .MMO M,   MMMMMMMM.MMMMMMMM..8MMM .,MMM
'   IMMMMM. .MMMM. MMM.. MMZ.M~   MMMM OMMM,MMMMMMMMO MMM..MMM,
'   DMMMMMM  .MMMM.MMM,. MM:7MD.M.IMM+..MMM MM8 MMMMM.MMM MMMM
'   DMMMM ..  .MMM MMM. $MM.NMMMMD.MMM..MMM MMM  MMMM.MMM$MMM..
'   ZMMMM?     MMM.MMM. MMM.MMMMM8 MMM MMMM.MMM  .MM..MMMMMM.
'    MMMMM .   MMM.MMM$.MMM MMMMM: MMMMMMM. MMM      .MMMM. .
'    MMMMMM  .MMMM.MMMM?MMM.MMMMM OMMMMMM  MMM      OMMM$ .
'    .MMMMMMMMMMM.?MM:. ...  .. ..MMMMM.   MM .   MMM. .
'     ?MMMMMMMMMM MM              .        .
'      .MMMMMMMM
'        ZMMMM.
'                      Sinbad by Gottlieb (1978)
'
'               **** Graphics and Sound by Pinuck ****
'                **** Build and Code by Kruge99 ****
'        ** Codebase adapted from VP8 Sinbad by Luvthatapex **
'
'                    ** VPX conversion notes **
'          ** Thanks to fuzzel, Toxie, mukuste for VPX **
'
' In no particular order thanks goes out to the following people for their
' VPX convsion help, tips and suggestions :
' flupper1, BorgDog, ICPjuggla, zany, mfuegeman, cyberpez, allknowing2012, arngrim, JPSalas, hauntfreaks
'
'
'pinDMD3 users who want to enable rainbow mode can change the DMD colours
'starting on line 211.  Change all three values to 0/zero to enable rainbow mode.

' Thalamus 2018-07-24
' Table didn't have standard "Positional Sound Playback Functions" or "Supporting Ball & Sound Functions"
' Changed UseSolenoids=1 to 2
' No special SSF tweaks yet.

Option Explicit
Randomize

' Thalamus 2018-08-26 : Improved directional sounds

' !! NOTE : Table not verified yet !!

Const VolDiv = 2000    ' Lower number, louder ballrolling/collition sound
Const VolCol    = 3    ' Ball collition divider ( voldiv/volcol )

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


If Table1.ShowDT = False then
  BgGO.visible=False
  BgTILT.visible=False
  BgHSTD.visible=False
  BgSPSA.visible=False
  CrD1.visible=False
  CrD2.visible=False
  BaD1.visible=False
  BaD2.visible=False
  P1d1.visible=False
  p1d2.visible=False
  p1d3.visible=False
  p1d4.visible=False
  p1d5.visible=False
  p1d6.visible=False
  p2d1.visible=False
  p2d2.visible=False
  p2d3.visible=False
  p2d4.visible=False
  p2d5.visible=False
  p2d6.visible=False
  p3d1.visible=False
  p3d2.visible=False
  p3d3.visible=False
  p3d4.visible=False
  p3d5.visible=False
  p3d6.visible=False
  p4d1.visible=False
  p4d2.visible=False
  p4d3.visible=False
  p4d4.visible=False
  p4d5.visible=False
  p4d6.visible=False
End If

On Error Resume Next
ExecuteGlobal GetTextFile("controller.vbs")
If Err Then MsgBox "Can't open controller.vbs"
On Error Goto 0

Const cGameName="sinbad",UseSolenoids=2,UseLamps=0,UseGI=0,UseSync=1,SSolenoidOn="SolOn",SSolenoidOff="SolOff",SFlipperOn="FlipperUp",SFlipperOff="FlipperDown"
Const SCoin="coin3",cCredits=""

LoadVPM "01001100", "gts1.vbs", 3.02
Dim DesktopMode: DesktopMode = Table1.ShowDT

Dim counter
Dim bump1, bump2

Const WYTar=6
Const PTar=7
Const RTar=8

SolCallback(1)="bsTrough.SolOut"
SolCallback(2)="VpmSolSound SoundFX(""knock"",DOFKnocker),"
SolCallback(3)="vpmSolSound SoundFX(""10pts"",DOFChimes),"
SolCallback(4)="vpmSolSound SoundFX(""100pts"",DOFChimes),"
SolCallback(5)="vpmSolSound SoundFX(""1000pts"",DOFChimes),"
SolCallback(sLLFlipper)="SolLFlipper"
SolCallback(sLRFlipper)="SolRFlipper"
SolCallback(WYTar)="WYRaised" 'dtDropW.SolDropUp
SolCallBack(PTar)="PRaised"   'dtDropP.SolDropUp
SolCallback(RTar)="RRaised"   'dtDropR.SolDropUp

Sub WYRaised(enabled)
	debug.print "WYRaised"
	If enabled Then
		WYReset.interval=500
		WYReset.enabled=True
	End If
End Sub

Sub WYReset_Timer()
	debug.print "Raise White and Yellow Targets"
	GIPL37.State=0
	WYReset.enabled=False
	dtDropW.DropSol_On
End Sub

Sub PRaised(enabled)
	debug.print "PRaised"
	If enabled Then
		PReset.interval=500
		PReset.enabled=True
	End If
End Sub

Sub PReset_Timer()
	debug.print "Raise Purple Targets"
	GIPL38.State=0:GIPL39.State=0:GIPL40.State=0
	PReset.enabled=False
	dtDropP.DropSol_On
End Sub

Sub RRaised(enabled)
	debug.print "RRaised"
	If enabled Then
		RReset.interval=500
		RReset.enabled=True
	End If
End Sub

Sub RReset_Timer()
	debug.print "Raise Red Targets"
	GIPL30.State=0:GIPL31.State=0:GIPL32.State=0:GIPL33.State=0
	RReset.enabled=False
	dtDropR.DropSol_On
End Sub

Sub SolLFlipper(enabled)
	If Enabled Then
		PlaySoundAtVol SoundFX("PNK_MH_Flip_L_up",DOFContactors),LeftFlipper, VolFlip:LeftFlipper.RotateToEnd:Flipper2.RotateToEnd
	Else
		PlaySoundAtVol SoundFX("PNK_MH_Flip_L_down",DOFContactors),LeftFlipper, VolFlip:LeftFlipper.RotateToStart:Flipper2.RotateToStart
	End If
End Sub

Sub SolRFlipper(enabled)
	If Enabled Then
		PlaySoundAtVol SoundFX("PNK_MH_Flip_R_up",DOFContactors),RightFlipper, VolFlip:RightFlipper.RotateToEnd:Flipper1.RotateToEnd
	Else
		PlaySoundAtVol SoundFX("PNK_MH_Flip_R_down",DOFContactors),RightFlipper, VolFlip:RightFlipper.RotateToStart:Flipper1.RotateToStart
	End If
End Sub

Sub Table1_KeyDown(ByVal keycode)
	If keycode=PlungerKey Then Plunger.Pullback:playsoundAt "PullbackPlunger", Plunger
	If keycode=AddCreditKey then PlaySoundAt "CoinReject", Drain:vpmtimer.pulsesw 1
	If keycode = LeftTiltKey Then Nudge 90, 3:Playsound "fx_nudge_left"			'JPSalas recommended strength = 3
	If keycode = RightTiltKey Then Nudge 270, 3:Playsound "fx_nudge_right"		'JPSalas recommended strength = 3
	If keycode = CenterTiltKey Then Nudge 0, 2.5:Playsound "fx_nudge_forward"	'JPSalas recommended strenght = 2.5
	If keycode = keyReset then
		UpdateGI 0:GIT.Enabled=1
	end if
	If vpmKeyDown(KeyCode) Then Exit Sub
End Sub

Sub Table1_KeyUp(ByVal keycode)
	If vpmKeyUp(KeyCode) Then Exit Sub
	If keycode=PlungerKey Then Plunger.Fire:PlaysoundAt "EmptyPlunger", Plunger
End Sub

sub flippertimer_timer()
	LPFlipper.ObjRotZ = LeftFlipper.CurrentAngle
	LPFlipper1.ObjRotZ = Flipper2.CurrentAngle
	RPFlipper.ObjRotZ = RightFlipper.CurrentAngle
	RPFlipper1.ObjRotZ = Flipper1.CurrentAngle
	SINspinner.RotZ = (Spinner1.currentangle)
end sub

Dim xx
Dim bsTrough,dtDropW,dtDropP,dtDropR,bsSaucer1
Dim GIState
GIState=0

Sub Table1_Init
LS74d.State = 1		'lightstate.state --> 0 = off, 1 = on
LS74c.State = 1
LS74b.State = 1
LS74a.State = 1

Controller.Games(cGameName).Settings.Value("dmd_red")=0 	'** change the dmd colour to dark blue
Controller.Games(cGameName).Settings.Value("dmd_green")=128 '** pinDMD3 users change the colours to black 0/zero for rainbow mode
Controller.Games(cGameName).Settings.Value("dmd_blue")=255
	vpmInit Me
	On Error Resume Next
		With Controller
		.GameName = cGameName
		.GameName = cGameName
		If Err Then MsgBox "Can't start Game" & cGameName & vbNewLine & Err.Description : Exit Sub
		.SplashInfoLine = "Sinbad (Gottlieb 1978)"&chr(13)&"a pinuck/kruge99 joint"
		.HandleMechanics=0
		.HandleKeyboard=0
		.ShowDMDOnly=1
		.ShowFrame=0
		.ShowTitle=0
		.hidden = DesktopMode
		'.Games(cGameName).Settings.Value("dmd_width")=226
		'.Games(cGameName).Settings.Value("dmd_height")=77
		'.Games(cGameName).Settings.Value("rol")=0
		If Err Then MsgBox Err.Description
	End With
	On Error Goto 0
		Controller.SolMask(0)=0
	vpmTimer.AddTimer 2000,"Controller.SolMask(0)=&Hffffffff'" 'ignore all solenoids - then add the timer to renable all the solenoids after 2 seconds
		Controller.Run
	If Err Then MsgBox Err.Description
	On Error Goto 0

 	PinMAMETimer.Interval=PinMAMEInterval
	PinMAMETimer.Enabled=1

 	vpmNudge.TiltSwitch=4
	vpmNudge.Sensitivity=5
	vpmNudge.TiltObj=Array(Bumper1,Bumper2)

 	Set bsTrough = New cvpmBallStack
	bsTrough.initnotrough ballrelease, 66, 55, 6
	bsTrough.InitExitSnd SoundFX("balllaunch",DOFContactors),SoundFX("newball",DOFContactors)

 	Set dtDropW=New cvpmDropTarget'
	dtDropW.InitDrop Array(TW1,TY1,TY2),Array(20,21,24)
	dtDropW.InitSnd SoundFX("PNK_MH_Drop_fall_L",DOFContactors),SoundFX("PNK_MH_Drop_reset_L",DOFContactors)

	Set dtDropP=New cvpmDropTarget'
	dtDropP.InitDrop Array(TP1,TP2,TP3),Array(30,31,34)
	dtDropP.InitSnd SoundFX("PNK_MH_Drop_fall_R",DOFContactors),SoundFX("PNK_MH_Drop_reset_R",DOFContactors)

	Set dtDropR=New cvpmDropTarget'
	dtDropR.InitDrop Array(TR1,TR2,TR3,TR4),Array(50,51,60,61)
	dtDropR.InitSnd SoundFX("PNK_MH_Drop_fall_L",DOFContactors),SoundFX("PNK_MH_Drop_reset_L",DOFContactors)

'**Detect pinwizard
If Plunger.MotionDevice>0 Then PBWPlunger.Enabled=True

'GI Init
UpdateGI 0
GIT.Enabled=1

End Sub

Sub Table1_Paused:Controller.Pause = 1:End Sub
Sub Table1_unPaused:Controller.Pause = 0:End Sub

Sub Bumper1_Hit:vpmTimer.PulseSw 14:PlaySoundAtVol SoundFXDOF("bumper_L",112,DOFPulse,DOFContactors), Bumper1, VolBump:bump1 = 1:End Sub
Sub Bumper2_Hit:vpmTimer.PulseSw 14:PlaySoundAtVol SoundFXDOF("bumper_R",113,DOFPulse,DOFContactors), Bumper2, VolBump:bump2 = 1:End Sub

'**top rollover switches

Sub SW70_Hit():Controller.Switch(70)=1:End Sub
Sub SW70_UnHit():Controller.Switch(70)=0:End Sub
Sub SW71_Hit()::Controller.Switch(71)=1:End Sub
Sub SW71_UnHit():Controller.Switch(71)=0:End Sub
Sub SW40_Hit():Controller.Switch(40)=1:DOF 101, DOFOn:End Sub
Sub SW40_UnHit():Controller.Switch(40)=0:DOF 101, DOFOff:End Sub
Sub SW41_Hit():Controller.Switch(41)=1:DOF 105, DOFOn:End Sub
Sub SW41_UnHit():Controller.Switch(41)=0:DOF 105, DOFOff:End Sub

'**left green star rollover switches
Sub SW74d_Hit():LS74d.State = 0:Controller.Switch(74)=1:DOF 111, DOFOn:End Sub
Sub SW74d_UnHit():LS74d.State = 1:Controller.Switch(74)=0:DOF 111, DOFOff:End Sub
Sub SW74c_Hit():LS74c.State = 0:Controller.Switch(74)=1:DOF 110, DOFOn:End Sub
Sub SW74c_UnHit():LS74c.State = 1:Controller.Switch(74)=0:DOF 110, DOFOff:End Sub
Sub SW74b_Hit():LS74b.State = 0:Controller.Switch(74)=1:DOF 109, DOFOn:End Sub
Sub SW74b_UnHit():LS74b.State = 1:Controller.Switch(74)=0:DOF 109, DOFOff:End Sub
Sub SW74a_Hit():LS74a.State = 0:Controller.Switch(74)=1:DOF 108, DOFOn:End Sub
Sub SW74a_UnHit():LS74a.State = 1:Controller.Switch(74)=0:DOF 108, DOFOff:End Sub

Sub Spinner1_Spin:vpmTimer.PulseSw(10):PlaysoundAtVol "PNK_SB_spinner", Spinner1, VolSpin:End Sub

'**round target switches
Sub TargetSW41a_hit:vpmTimer.PulseSw(41):PlaysoundAtVol SoundFXDOF("spothit",106,DOFPulse,DOFContactors),TargetSW51a,VolTarg:GIPL34.State=0:TargetSW41a.TimerEnabled=True:End Sub
Sub TargetSW41a_Timer:TargetSW41a.TimerEnabled=False:GIPL34.State=1:End Sub

Sub TargetSW41c_hit:vpmTimer.PulseSw(41):PlaysoundAt SoundFXDOF("spothit",107,DOFPulse,DOFContactors),TargetSW41c:End Sub

'**right lane switch
Sub SW44_Hit():Controller.Switch(44)=1:End Sub
Sub SW44_UnHit():Controller.Switch(44)=0:PlaysoundAt "PNK_MH_metal_roll", sw44:End Sub

'**red drop target bank
Sub TR1_Hit():dtDropR.Hit 1:GIPL30.State=1:End Sub
Sub TR2_Hit():dtDropR.Hit 2:GIPL31.State=1:End Sub
Sub TR3_Hit():dtDropR.Hit 3:GIPL32.State=1:End Sub
Sub TR4_Hit():dtDropR.HIt 4:GIPL33.State=1:End Sub

'**white & yellow drop target bank
Sub TW1_Hit():dtDropW.Hit 1:End Sub
Sub TY1_Hit():dtDropW.Hit 2:End Sub
Sub TY2_Hit():dtDropW.Hit 3:GIPL37.State=1:End Sub

'**purple drop target bank
Sub TP1_Hit():dtDropP.Hit 1:GIPL38.State=1:End Sub
Sub TP2_Hit():dtDropP.Hit 2:GIPL39.State=1:End Sub
Sub TP3_Hit():dtDropP.Hit 3:GIPL40.State=1:End Sub

'**triggers for leaf switches linked to switch 54
Sub SW54a_Hit():vpmTimer.PulseSw(54):End Sub
Sub SW54b_Hit():vpmTimer.PulseSw(54):End Sub
Sub SW54c_Hit():vpmTimer.PulseSw(54):End Sub
Sub SW54d_Hit():vpmTimer.PulseSw(54):SW54d.TimerEnabled=True:End Sub
Sub SW54d_Timer:Sw54d.TimerEnabled=False:End Sub
Sub SW54e_Hit():vpmTimer.PulseSw(54):End Sub
Sub SW54f_Hit():vpmTimer.PulseSw(54):End Sub
Sub SW54g_Hit():vpmTimer.PulseSw(54):End Sub
Sub SW54h_Hit():vpmTimer.PulseSw(54):End Sub

'**not sure what these do, there are no matching objects
Sub LeftSlingshot_slingshot():vpmtimer.pulsesw 60:PlaysoundAt SoundFX("slingshot",DOFContactors),ActiveBall:End Sub
Sub RightSlingshot_Slingshot():vpmtimer.pulsesw 60:PlaysoundAt SoundFX("slingshot",DOFContactors),ActiveBall:End Sub
Sub TargetExtraBall_Hit():vpmtimer.pulsesw 64:playsoundAt SoundFX("spothit",DOFContactors),TargetExtraBall:TargetExtraBall.Isdropped=1:TargetExtraBall2.Isdropped=0:TargetExtraBall.Timerenabled=1:End Sub
Sub TargetExtraBall_Timer:TargetExtraBall.TimerEnabled=0:TargetExtraBall.isdropped=0:TargetExtraBall2.isdropped=1:End Sub

'***********************
'  Extra Sounds
'***********************

Sub Gate1_Hit():PlaysoundAt "Gate", Gate1:End Sub
Sub Rubbers_Hit(idx):RandomSoundRubber():End Sub
Sub sw54rubbers_Hit(idx):RandomSoundRubber():End Sub
Sub rollovers_Hit(idx):PlaySound "rollover", 0, Vol(ActiveBall)*VolRol, pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0, AudioFade(ActiveBall):End Sub

Sub RandomSoundRubber()
	Select Case Int(Rnd*3)+1
		Case 1 : PlaySound "PNK_MH_rubber_hit_1", 0, Vol(ActiveBall)*VolRH, pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0, AudioFade(ActiveBall)
		Case 2 : PlaySound "PNK_MH_rubber_hit_2", 0, Vol(ActiveBall)*VolRH, pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0, AudioFade(ActiveBall)
		Case 3 : PlaySound "PNK_MH_rubber_hit_3", 0, Vol(ActiveBall)*VolRH, pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0, AudioFade(ActiveBall)
	End Select
End Sub


'**you hit the drain, bad luck, try again!
Sub Drain_Hit():PlaysoundAt "PNK_MH_drain", Drain:bsTrough.addball me:DOF 103, DOFPulse:End Sub
Sub DrainRoll_Hit():Playsound "PNK_MH_drain_roll", 0, Vol(ActiveBall), pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0, AudioFade(ActiveBall):End Sub

'**you hit the left or right outlane, bad luck, try again!
Sub SW40b_Hit():Controller.Switch(40)=1:DOF 102, DOFOn:End Sub
Sub SW40b_Unhit():Controller.Switch(40)=0:DOF 102, DOFOff:End Sub
Sub SW44b_Hit():Controller.Switch(44)=1:End Sub
Sub SW44b_Unhit():Controller.Switch(44)=0:End Sub

'**if the ball is launched, play a ball rolling sound
Sub SoundOn_Unhit():Playsound "PlungerShootBall", 0, Vol(ActiveBall), pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0, AudioFade(ActiveBall):End Sub
Sub Wall85_Hit():RandomSoundRubber():End Sub

'**********GI Handler
Sub GIT_Timer():GIT.Enabled=0:UpdateGI 1:End Sub
	Sub UpdateGI(enabled)
		If enabled then
		GIState=1
'		GIPF.State=1
		For each xx in GIPL:xx.State=1:Next
		For each xx in SRO:xx.State=1:Next
		For each xx in CapLit:xx.State=0:Next
		For each xx in CapUnlit:xx.State=1:Next
	Else
		GIState=0
'		GIPF.State=0
		For each xx in GIPL:xx.State=0:Next
		For each xx in SRO:xx.State=0:Next
		For each xx in CapLit:xx.State=1:Next
		For each xx in CapUnlit:xx.State=0:Next
	End if
End Sub

'**update the lights
'****************************************
'  JP's Fading Lamps 3.4 VP9 Fading only
'      Based on PD's Fading Lights
' SetLamp 0 is Off
' SetLamp 1 is On
' LampState(x) current state
'****************************************

'Update the BG - GO, HSTD, SPSA, TILT
Dim N1,O1,N2,O2,N3,O3,N4,O4
N1=0:O1=0:N2=0:O2=0:N3=0:O3=0:N4=0:O4=0

Dim LampState(200)
Dim x
AllLampsOff()
LampTimer.Interval = 40
LampTimer.Enabled = 1

Sub LampTimer_Timer()
	Dim chgLamp, num, chg, ii
	chgLamp = Controller.ChangedLamps
	If Not IsEmpty(chgLamp) Then
		For ii = 0 To UBound(chgLamp)
			LampState(chgLamp(ii, 0) ) = chgLamp(ii, 1) + 4
 		Next
	End If

	UpdateLamps
	DisplayLEDs
End Sub

Sub UpdateLamps
'FadeR 1, BgGO   'Game Over
'FadeR 2, BgTilt 'Tilt
'FadeR 3, BgHSTD 'High Score To Date
'FaceR 4, BgSPSA 'Same Player Shoots Again
NFadeL 4, L4
NFadeL 5, L5
NFadeL 6, L6
NFadeL 7, L7
NFadeL 8, L8
NFadeL 9, L9
NFadeL 10, L10
NFadeL 11, L11
'NFadeL 12, L12
'NFadeL 13, L13
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
NFadeL 25, L25
NFadeL 26, L26
NFadeL 27, L27
NFadeL 28, L28
NFadeL 29, L29
NFadeL 30, L30
NFadeL 31, L31
NFadeL 32, L32
NFadeL 33, L33
NFadeL 34, L34

N1=Controller.Lamp(1) 'Game Over (Backglass)
N2=Controller.Lamp(2) 'Tilt (Backglass)
N3=Controller.Lamp(3) 'High Score To Date (Backglass)
N4=Controller.Lamp(4) 'Same Player Shoots Again (Backglass+Playfield)
If N1<>O1 Then
 		If N1 Then
 			'Textbox2.Text=""
			BgGO.setvalue 0
 		Else
 			'Textbox2.Text="GAME OVER"
			BgGO.setvalue 1
 		End If
 	O1=N1
 	End If
 	If N2<>O2 Then
 		If N2 Then
 			'Textbox3.Text="TILT"
			BgTILT.setvalue 1
 		Else
 			'Textbox3.Text=""
			BgTILT.setvalue 0
 		End If
 	O2=N2
 	End If
 	If N3<>O3 Then
 		If N3 Then
 			'HSBox.Text="HIGH SCORE TO DATE"
			BgHSTD.setvalue 1
 		Else
 			'HSBox.Text=""
			BgHSTD.setvalue 0
 		End If
 	O3=N3
 	End If
 	If N4<>O4 Then
 		If N4 Then
 			'Textbox1.Text="SHOOT AGAIN"
			BgSPSA.setvalue 1
 		Else
 			'Textbox1.Text=""
			BgSPSA.setvalue 0
 		End If
 	O4=N4
End If
End Sub

Sub AllLampsOff():For x = 1 to 200:LampState(x) = 4:Next:UpdateLamps:UpdateLamps:Updatelamps:End Sub

Sub SetLamp(nr, value):LampState(nr) = abs(value) + 4:End Sub

'System defines variables and assigns the changed light values to them, which is then turned over to Fader Control
'Dim lightnum(201) : Dim Tset
'Initialize all possible light control variables (0 - 200) including extra spaces for flashers or GI control
'On Error Resume Next
'	For Tset = 0 To 200
'		lightnum(Tset) = 0
'	Next
'
Sub NFadeL(nr, a)
	Select Case LampState(nr)
		Case 4:a.state = 0:LampState(nr) = 0
		Case 5:a.State = 1:LampState(nr) = 1
	End Select
End Sub

'Sub FadeR(lightnumIn, reelname)
'	Select Case lightnum(lightnumIn)
'		Case  5 : reelname.setvalue 1 : lightnum(lightnumIn) = 7
'		Case  4 : reelname.setvalue 3 : lightnum(lightnumIn) = lightnum(lightnumIn) - 1
'		Case  3 : reelname.setvalue 2 : lightnum(lightnumIn) = lightnum(lightnumIn) - 1
'		Case  2 : reelname.setvalue 0 : lightnum(lightnumIn) = 6
'		Case  1 : reelname.setvalue 1 : lightnum(lightnumIn) = 7
'		Case  0 : reelname.setvalue 0 : lightnum(lightnumIn) = 6
'	End Select
'End Sub

'****************************************
'  Ball Rolling
'****************************************
' Thalamus, didn't seem to be active

' dim ballspeed
'
' Sub BallRollSounds_hit(T) 'ball rolling sounds
' 	ballspeed=ABS(ActiveBall.vely)+ABS(ActiveBall.velx)
' '	speedtest.text= ballspeed
' 	if ballspeed >= 6 then
' 		Select Case Int(Rnd*3)+1
' 			Case 1 : PlaySound "PNK_MH_roll_short_1"
' 			Case 2 : PlaySound "PNK_MH_roll_short_2"
' 			Case 3 : PlaySound "PNK_MH_roll_short_3"
' 		End Select
' 	end if
' End Sub

' Thalamus, replaced by JP's standards

' '======================================================================================
' ' Many thanks go to Randy Davis and to all the multitudes of people who have
' ' contributed to VP over the years, keeping it alive!!!  Enjoy, Steely & PK
' '======================================================================================
'
' Dim tnopb, nosf
' '
' tnopb = 4	' <<<<< SET to the "Total Number Of Possible Balls" in play at any one time
' nosf = 9	' <<<<< SET to the "Number Of Sound Files" used / B2B collision volume levels
'
' Dim currentball(3), ballStatus(3)
' Dim iball, cnt, coff, errMessage
'
' XYdata.interval = 1 			' Timer interval starts at 1 for the highest ball data sample rate
' coff = False				' Collision off set to false
'
' For cnt = 0 to ubound(ballStatus)	' Initialize/clear all ball stats, 1 = active, 0 = non-existant
' 	ballStatus(cnt) = 0
' Next
'
' '======================================================
' ' <<<<<<<<<<<<<< Ball Identification >>>>>>>>>>>>>>
' '======================================================
' ' Call this sub from every kicker(or plunger) that creates a ball.
' Sub NewBallID 						' Assign new ball object and give it ID for tracking
' 	For cnt = 1 to ubound(ballStatus)		' Loop through all possible ball IDs
'    	If ballStatus(cnt) = 0 Then			' If ball ID is available...
'    	Set currentball(cnt) = ActiveBall			' Set ball object with the first available ID
'    	currentball(cnt).uservalue = cnt			' Assign the ball's uservalue to it's new ID
'    	ballStatus(cnt) = 1				' Mark this ball status active
'    	ballStatus(0) = ballStatus(0)+1 		' Increment ballStatus(0), the number of active balls
' 	If coff = False Then				' If collision off, overrides auto-turn on collision detection
' 							' If more than one ball active, start collision detection process
' 	If ballStatus(0) > 1 and XYdata.enabled = False Then XYdata.enabled = True
' 	End If
' 	Exit For					' New ball ID assigned, exit loop
'    	End If
'    	Next
' '  	Debugger 					' For demo only, display stats
' End Sub
'
' ' Call this sub from every kicker that destroys a ball, before the ball is destroyed.
' Sub ClearBallID
'   	On Error Resume Next				' Error handling for debugging purposes
'    	iball = ActiveBall.uservalue			' Get the ball ID to be cleared
'    	currentball(iball).UserValue = 0 			' Clear the ball ID
'    	If Err Then Msgbox Err.description & vbCrLf & iball
'     	ballStatus(iBall) = 0 				' Clear the ball status
'    	ballStatus(0) = ballStatus(0)-1 		' Subtract 1 ball from the # of balls in play
'    	On Error Goto 0
' End Sub
'
' '=====================================================
' ' <<<<<<<<<<<<<<<<< XYdata_Timer >>>>>>>>>>>>>>>>>
' '=====================================================
' ' Ball data collection and B2B Collision detection.
' ReDim baX(tnopb,4), baY(tnopb,4), bVx(tnopb,4), bVy(tnopb,4), TotalVel(tnopb,4)
' Dim cForce, bDistance, xyTime, cFactor, id, id2, id3, B1, B2
'
' Sub XYdata_Timer()
' 	' xyTime... Timers will not loop or start over 'til it's code is finished executing. To maximize
' 	' performance, at the end of this timer, if the timer's interval is shorter than the individual
' 	' computer can handle this timer's interval will increment by 1 millisecond.
'     xyTime = Timer+(XYdata.interval*.001)	' xyTime is the system timer plus the current interval time
' 	' Ball Data... When a collision occurs a ball's velocity is often less than it's velocity before the
' 	' collision, if not zero. So the ball data is sampled and saved for four timer cycles.
'    	If id2 >= 4 Then id2 = 0						' Loop four times and start over
'    	id2 = id2+1								' Increment the ball sampler ID
'    	For id = 1 to ubound(ballStatus)					' Loop once for each possible ball
'   	If ballStatus(id) = 1 Then						' If ball is active...
'    		baX(id,id2) = round(currentball(id).x,2)				' Sample x-coord
'    		baY(id,id2) = round(currentball(id).y,2)				' Sample y-coord
'    		bVx(id,id2) = round(currentball(id).velx,2)				' Sample x-velocity
'    		bVy(id,id2) = round(currentball(id).vely,2)				' Sample y-velocity
'    		TotalVel(id,id2) = (bVx(id,id2)^2+bVy(id,id2)^2) 		' Calculate total velocity
'  		If TotalVel(id,id2) > TotalVel(0,0) Then TotalVel(0,0) = int(TotalVel(id,id2))
'    	End If
'    	Next
' 	' Collision Detection Loop - check all possible ball combinations for a collision.
' 	' bDistance automatically sets the distance between two colliding balls. Zero milimeters between
' 	' balls would be perfect, but because of timing issues with ball velocity, fast-traveling balls
' 	' prevent a low setting from always working, so bDistance becomes more of a sensitivity setting,
' 	' which is automated with calculations using the balls' velocities.
' 	' Ball x/y-coords plus the bDistance determines B2B proximity and triggers a collision.
' 	id3 = id2 : B2 = 2 : B1 = 1						' Set up the counters for looping
' 	Do
' 	If ballStatus(B1) = 1 and ballStatus(B2) = 1 Then			' If both balls are active...
' 		bDistance = int((TotalVel(B1,id3)+TotalVel(B2,id3))^1.04)
' 		If ((baX(B1,id3)-baX(B2,id3))^2+(baY(B1,id3)-baY(B2,id3))^2)<2800+bDistance Then collide B1,B2 : Exit Sub
' 		End If
' 		B1 = B1+1							' Increment ball1
' 		If B1 >= ballStatus(0) Then Exit Do				' Exit loop if all ball combinations checked
' 		If B1 >= B2 then B1 = 1:B2 = B2+1				' If ball1 >= reset ball1 and increment ball2
' 	Loop
'
'  	If ballStatus(0) <= 1 Then XYdata.enabled = False 			' Turn off timer if one ball or less
'
' 	If XYdata.interval >= 40 Then coff = True : XYdata.enabled = False	' Auto-shut off
' 	If Timer > xyTime * 3 Then coff = True : XYdata.enabled = False		' Auto-shut off
'    	If Timer > xyTime Then XYdata.interval = XYdata.interval+1		' Increment interval if needed
' End Sub
'
' '=========================================================
' ' <<<<<<<<<<< Collide(ball id1, ball id2) >>>>>>>>>>>
' '=========================================================
' 'Calculate the collision force and play sound accordingly.
' Dim cTime, cb1,cb2, avgBallx, cAngle, bAngle1, bAngle2
'
' Sub Collide(cb1,cb2)
' ' The Collision Factor(cFactor) uses the maximum total ball velocity and automates the cForce calculation, maximizing the
' ' use of all sound files/volume levels. So all the available B2B sound levels are automatically used by adjusting to a
' ' player's style and the table's characteristics.
'  	If TotalVel(0,0)/1.8 > cFactor Then cFactor = int(TotalVel(0,0)/1.8)
' ' The following six lines limit repeated collisions if the balls are close together for any period of time
'   	avgBallx = (bvX(cb2,1)+bvX(cb2,2)+bvX(cb2,3)+bvX(cb2,4))/4
'   	If avgBallx < bvX(cb2,id2)+.1 and avgBallx > bvX(cb2,id2)-.1 Then
'  	If ABS(TotalVel(cb1,id2)-TotalVel(cb2,id2)) < .000005 Then Exit Sub
'  	End If
'   	If Timer < cTime Then Exit Sub
'   	cTime = Timer+.1				' Limits collisions to .1 seconds apart
' ' GetAngle(x-value, y-value, the angle name) calculates any x/y-coords or x/y-velocities and returns named angle in radians
'  	GetAngle baX(cb1,id3)-baX(cb2,id3), baY(cb1,id3)-baY(cb2,id3),cAngle	' Collision angle via x/y-coordinates
' 	id3 = id3 - 1 : If id3 = 0 Then id3 = 4		' Step back one xyData sampling for a good velocity reading
'  	GetAngle bVx(cb1,id3), bVy(cb1,id3), bAngle1	' ball 1 travel direction, via velocity
'  	GetAngle bVx(cb2,id3), bVy(cb2,id3), bAngle2	' ball 2 travel direction, via velocity
' ' The main cForce formula, calculating the strength of a collision
' 	cForce = Cint((abs(TotalVel(cb1,id3)*Cos(cAngle-bAngle1))+abs(TotalVel(cb2,id3)*Cos(cAngle-bAngle2))))
'     	If cForce < 4 Then Exit Sub			' Another collision limiter
'    	cForce = Cint((cForce)/(cFactor/nosf))		' Divides up cForce for the proper sound selection.
'   	If cForce > nosf-1 Then cForce = nosf-1		' First sound file 0(zero) minus one from number of sound files
'    	PlaySound("collide" & cForce)			' Combines "collide" with the calculated sound level and play sound
' End Sub
'
' '=================================================
' ' <<<<<<<< GetAngle(X, Y, Anglename) >>>>>>>>
' '=================================================
' ' A repeated function which takes any set of coordinates or velocities and calculates an angle in radians.
' Dim Xin,Yin,rAngle,Radit,wAngle,Pi
' Pi = Round(4*Atn(1),6)					'3.1415926535897932384626433832795
'
' Sub GetAngle(Xin, Yin, wAngle)
'  	If Sgn(Xin) = 0 Then
'  		If Sgn(Yin) = 1 Then rAngle = 3 * Pi/2 Else rAngle = Pi/2
'  		If Sgn(Yin) = 0 Then rAngle = 0
'  	Else
'  		rAngle = atn(-Yin/Xin)			' Calculates angle in radians before quadrant data
'  	End If
'  	If sgn(Xin) = -1 Then Radit = Pi Else Radit = 0
'  	If sgn(Xin) = 1 and sgn(Yin) = 1 Then Radit = 2 * Pi
'  	wAngle = round((Radit + rAngle),4)		' Calculates angle in radians with quadrant data
' 	'"wAngle = round((180/Pi) * (Radit + rAngle),4)" ' Will convert radian measurements to degrees - to be used in future
' End Sub

Sub Flipper2_Collide(parm)
   RandomSoundFlipper()
End Sub

Sub LeftFlipper_Collide(parm)
   RandomSoundFlipper()
End Sub

Sub Flipper1_Collide(parm)
   RandomSoundFlipper()
End Sub

Sub RightFlipper_Collide(parm)
   RandomSoundFlipper()
End Sub

Sub RandomSoundFlipper()
	Select Case Int(Rnd*3)+1
		Case 1 : PlaySoundAt "PNK_MH_Flip_hit_1", ActiveBall
		Case 2 : PlaySoundAt "PNK_MH_Flip_hit_2", ActiveBall
		Case 3 : PlaySoundAt "PNK_MH_Flip_hit_3", ActiveBall
	End Select
End Sub

'Pacdude's led display, modified from scapinos fathom and borrowed from noah
'
Dim SixDigitOutput(28) 'modified from 7 digit to 6 digit output
'Dim DisplayPatterns(11)

'Binary/Hex Pattern Recognition Array
'DisplayPatterns(0) = 0 '0000000 Blank
'DisplayPatterns(1) = 63 '0111111 zero
'DisplayPatterns(2) = 6 '0000110 one
'DisplayPatterns(3) = 91 '1011011 two
'DisplayPatterns(4) = 79 '1001111 three
'DisplayPatterns(5) = 102 '1100110 four
'DisplayPatterns(6) = 109 '1101101 five
'DisplayPatterns(7) = 125 '1111101 six
'DisplayPatterns(8) = 7 '0000111 seven
'DisplayPatterns(9) = 127 '1111111 eight
'DisplayPatterns(10) = 111 '1101111 nine

'Assign 7-digit output to reels
Set SixDigitOutput(0) = P1D1
Set SixDigitOutput(1) = P1D2
Set SixDigitOutput(2) = P1D3
Set SixDigitOutput(3) = P1D4
Set SixDigitOutput(4) = P1D5
Set SixDigitOutput(5) = P1D6

Set SixDigitOutput(6) = P2D1
Set SixDigitOutput(7) = P2D2
Set SixDigitOutput(8) = P2D3
Set SixDigitOutput(9) = P2D4
Set SixDigitOutput(10) = P2D5
Set SixDigitOutput(11) = P2D6

Set SixDigitOutput(12) = P3D1
Set SixDigitOutput(13) = P3D2
Set SixDigitOutput(14) = P3D3
Set SixDigitOutput(15) = P3D4
Set SixDigitOutput(16) = P3D5
Set SixDigitOutput(17) = P3D6

Set SixDigitOutput(18) = P4D1
Set SixDigitOutput(19) = P4D2
Set SixDigitOutput(20) = P4D3
Set SixDigitOutput(21) = P4D4
Set SixDigitOutput(22) = P4D5
Set SixDigitOutput(23) = P4D6

Set SixDigitOutput(24) = CrD2
Set SixDigitOutput(25) = CrD1
Set SixDigitOutput(26) = BaD2
Set SixDigitOutput(27) = BaD1

Sub DisplayLEDs ' 6-Digit output
On Error Resume Next
Dim ChgLED, ii, stat, TempCount

ChgLED = Controller.ChangedLEDs(&HFF, &HFFFF) 'hex of binary (display 111111, or first 6 digits)

If Not IsEmpty(ChgLED) Then
	If DesktopMode = True Then
For ii = 0 To UBound(ChgLED)
stat = chgLED(ii, 2)
	Select Case stat
		Case 0:SixDigitOutput(chgLED(ii, 0) ).SetValue 0 'empty
		Case 63:SixDigitOutput(chgLED(ii, 0) ).SetValue 1 '0
		Case 6:SixDigitOutput(chgLED(ii, 0) ).SetValue 2 '1
		Case 91:SixDigitOutput(chgLED(ii, 0) ).SetValue 3 '2
		Case 79:SixDigitOutput(chgLED(ii, 0) ).SetValue 4 '3
		Case 102:SixDigitOutput(chgLED(ii, 0) ).SetValue 5 '4
		Case 109:SixDigitOutput(chgLED(ii, 0) ).SetValue 6 '5
		Case 124:SixDigitOutput(chgLED(ii, 0) ).SetValue 7 '6
		Case 125:SixDigitOutput(chgLED(ii, 0) ).SetValue 7 '6
		Case 252:SixDigitOutput(chgLED(ii, 0) ).SetValue 7 '6
		Case 7:SixDigitOutput(chgLED(ii, 0) ).SetValue 8 '7
		Case 127:SixDigitOutput(chgLED(ii, 0) ).SetValue 9 '8
		Case 103:SixDigitOutput(chgLED(ii, 0) ).SetValue 10 '9
		Case 111:SixDigitOutput(chgLED(ii, 0) ).SetValue 10 '9
		Case 231:SixDigitOutput(chgLED(ii, 0) ).SetValue 10 '9
		Case 128:SixDigitOutput(chgLED(ii, 0) ).SetValue 0 'empty
		Case 191:SixDigitOutput(chgLED(ii, 0) ).SetValue 1 '0
		Case 832:SixDigitOutput(chgLED(ii, 0) ).SetValue 2 '1
		Case 896:SixDigitOutput(chgLED(ii, 0) ).SetValue 2 '1
		Case 768:SixDigitOutput(chgLED(ii, 0) ).SetValue 2 '1
		Case 134:SixDigitOutput(chgLED(ii, 0) ).SetValue 2 '1
		Case 219:SixDigitOutput(chgLED(ii, 0) ).SetValue 3 '2
		Case 207:SixDigitOutput(chgLED(ii, 0) ).SetValue 4 '3
		Case 230:SixDigitOutput(chgLED(ii, 0) ).SetValue 5 '4
		Case 237:SixDigitOutput(chgLED(ii, 0) ).SetValue 6 '5
		Case 253:SixDigitOutput(chgLED(ii, 0) ).SetValue 7 '6
		Case 135:SixDigitOutput(chgLED(ii, 0) ).SetValue 8 '7
		Case 255:SixDigitOutput(chgLED(ii, 0) ).SetValue 9 '8
		Case 239:SixDigitOutput(chgLED(ii, 0) ).SetValue 10 '9
	End Select
'For TempCount = 0 to 10
'If stat = DisplayPatterns(TempCount) OR stat = (DisplayPatterns(TempCount) + 128) then
'SevenDigitOutput(chgLED(ii, 0) ).SetValue(TempCount)
'End If
'Next
Next
	End IF
End IF
End Sub

'**dipswitch menu - press F6 to display it

Sub editDips
	Dim vpmDips : Set vpmDips = New cvpmDips
	With vpmDips
		.AddForm 700,400,"System 1 - DIP switches"
		.AddFrame 0,0,190,"Coin chutecontrol",&H00040000,Array("seperate",0,"same",&H00040000)'dip 19
		.AddFrame 0,46,190,"Game mode",&H00000400,Array("extraball",0,"replay",&H00000400)'dip 11
		.AddFrame 0,92,190,"High game to date awards",&H00200000,Array("noaward",0,"3 replays",&H00200000)'dip 22
		.AddFrame 0,138,190,"Balls per game",&H00000100,Array("5 balls",0,"3balls",&H00000100)'dip 9
		.AddFrame 0,184,190,"Tilt effect",&H00000800,Array("game over",0,"ball in play only",&H00000800)'dip 12
		.AddFrame 205,0,190,"Maximum credits",&H00030000,Array("5 credits",0,"8 credits",&H00020000,"10 credits",&H00010000,"15 credits",&H00030000)'dip 17&18
		.AddFrame 205,76,190,"Soundsettings",&H80000000,Array("sounds",0,"tones",&H80000000)'dip 32
		.AddFrame 205,122,190,"Attract tune",&H10000000,Array("no attract tune",0,"attract tune played every 6 minutes",&H10000000)'dip 29
		.AddChk 205,175,190,Array("Match feature",&H00000200)'dip 10
		.AddChk 205,190,190,Array("Credits displayed",&H00001000)'dip 13
		.AddChk 205,205,190,Array("Play credit button tune",&H00002000)'dip 14
		.AddChk 205,220,190,Array("Play tones when scoring",&H00080000)'dip 20
		.AddChk 205,235,190,Array("Play coin switch tune",&H00400000)'dip 23
		.AddChk 205,250,190,Array("High game to date displayed",&H00100000)'dip 21
		.AddLabel 50,280,300,20,"After hitting OK, press F3 to reset game with new settings."
		.ViewDips
	End With
End Sub
Set vpmShowDips = GetRef("editDips")

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


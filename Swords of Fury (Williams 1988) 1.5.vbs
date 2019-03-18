'==============================================================================================='
'																								'
'															           	 		                '
' 					Swords of Fury / IPD No. 2486 / June, 1988 / 4 Players		  				'
'						    http://www.ipdb.org/machine.cgi?id=2486                             '
'																								'
' 	  		 	 	                	Created by takut 										'
'						              modded by HauntFreaks   					 				'
'																								'
'==============================================================================================='

Option Explicit
Randomize

' Thalamus 2018-11-01 : Improved directional sounds
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

Const BallSize = 50

Dim DesktopMode:DesktopMode = Table1.ShowDT

Dim UseVPMColoredDMD

UseVPMColoredDMD = DesktopMode

'LoadVPM "01560000", "DE.VBS", 3.46
LoadVPM"00990300","S11.VBS",3.10

'********************
'Standard definitions
'********************
Const cGameName = "swrds_l2" 'arcade rom - with credits

Const UseSolenoids = 2
Const UseLamps = 1
Const UseGI=0
Const UseSync = 0
Const HandleMech = 0

' Standard Sounds
Const SSolenoidOn = "fx_Solenoidon"
Const SSolenoidOff = "fx_solenoidoff"
Const SFlipperOn="fx_FlipperUp"
Const SFlipperOff="fx_FlipperDown"
Const SCoin = "fx_Coin"


Dim bsTrough, bsL, bsR, dtDrop, x, BallFrame, plungerIM, bsSaucer, bsLock, MagicTunnelLight

Set MotorCallback = GetRef("UpdateMultipleLamps")


'************
' Table init.
'************

If Table1.ShowDT = true then
    B6.Visible = 1
	B7.Visible = 1
    B41.Visible = 1
    B42.Visible = 1
    B51.Visible = 1
    B52.Visible = 1
	B57.Visible = 1
    B58.Visible = 1
    B59.Visible = 1
    B60.Visible = 1
    B61.Visible = 1
    B62.Visible = 1
    B63.Visible = 1
    B64.Visible = 1
	Jackpot.Visible = 1
else
    B6.Visible = 0
	B7.Visible = 0
    B41.Visible = 0
    B42.Visible = 0
    B51.Visible = 0
    B52.Visible = 0
    B57.Visible = 0
    B58.Visible = 0
    B59.Visible = 0
    B60.Visible = 0
    B61.Visible = 0
    B62.Visible = 0
    B63.Visible = 0
    B64.Visible = 0
	Jackpot.Visible = 0
End If

 Sub Table1_Init
    With Controller
       .GameName=cGameName
       If Err Then MsgBox "Can't start Game" & cGameName & vbNewLine & Err.Description : Exit Sub
       .SplashInfoLine="Swords of Fury (Williams 1988)"
       .HandleKeyboard=0
       .ShowTitle=0
       .ShowDMDOnly=1
       .ShowFrame=0
       .HandleMechanics=0
        If DesktopMode then
       .Hidden = 0
		else
        If B2SOn then
			.Hidden = 1
        else
			.Hidden = 0
		End If
        End If
       On Error Resume Next
       .Run GetPlayerHWnd
       If Err Then MsgBox Err.Description
    End With

    On Error Goto 0

    ' Nudging
    vpmNudge.TiltSwitch=swTilt
    vpmNudge.Sensitivity=3
    vpmNudge.TiltObj=Array(LeftSlingshot, RightSlingshot)

    ' Trough
    Set bsTrough=New cvpmBallStack
    bsTrough.InitSw 10,11,12,13,0,0,0,0
    bsTrough.InitKick BallRelease, 80, 6
    'bsTrough.InitEntrySnd "Solenoid", "Solenoid"
    bsTrough.InitExitSnd SoundFX("ballrelease",DOFContactors), SoundFX("fx_Solenoid",DOFContactors)
    bsTrough.Balls=3

 	Set bsSaucer=New cvpmBallStack
	bsSaucer.InitSaucer S54,54,180,2
	bsSaucer.InitExitSnd SoundFX("Popper",DOFContactors),SoundFX("SolOn",DOFContactors)

 	Set bsLock=New cvpmVLock
	bsLock.InitVLock Array(TriggerA,TriggerB),Array(Lock1,Lock2),Array(55,56)
	bsLock.InitSnd SoundFX("SolOn",DOFContactors),SoundFX("SolOn",DOFContactors)
	bsLock.ExitDir=0
	bsLock.ExitForce=25
	bsLock.CreateEvents"bsLock"

    ' Main Timer init
    PinMAMETimer.Interval=PinMAMEInterval
    PinMAMETimer.Enabled=1
'	Leds.Enabled=1

    ' Init Kickback
    KickBack.Pullback

 End Sub

Sub Table1_Paused : Controller.Pause=True : End Sub
Sub Table1_unPaused : Controller.Pause=False : End Sub

'**********
' Keys
'**********

 Sub Table1_KeyDown(ByVal Keycode)
    If keycode=RightFlipperKey Then Controller.Switch(58)=1
    If keycode=LeftFlipperKey Then Controller.Switch(60)=1
    If vpmKeyDown(keycode)Then Exit Sub
    If keycode=PlungerKey Then Plunger.Pullback : PlaySoundAtVol "plungerpull", Plunger, 1
 End Sub

 Sub Table1_KeyUp(ByVal Keycode)
    If keycode=RightFlipperKey Then Controller.Switch(58)=0
    If keycode=LeftFlipperKey Then Controller.Switch(60)=0
    If vpmKeyUp(keycode)Then Exit Sub
    If keycode=PlungerKey Then Plunger.Fire : PlaySoundAtVol "plunger", Plunger, 1
 End Sub


'*********
'Solenoids
'*********

SolCallback(1)="bsTrough.SolIn"
SolCallback(2)="bsTrough.SolOut"
SolCallback(4)="bsSaucer.Solout"
SolCallback(5)="dtDrop.SolunHit 2,"
SolCallback(6)="bsLock.SolExit"
SolCallback(7)="vpmSolSound SoundFX(""Knocker"",DOFKnocker),"
SolCallback(8)="dtDrop.SolunHit 3,"
SolCallback(9)="Flash9"
SolCallback(10)="SolGIUpdate"
'SolCallback(11)="SolFlasher10" 'lion man sword on backglass
'REM SolCallback(12)="SolACSelect"
SolCallback(14)="SolKickback"
SolCallback(15)="dtDrop.SolunHit 1,"
SolCallback(16)="Flash16" 'flasher below kickback insert
SolCallback(17)="vpmSolDiverter BottomDiverter,True,"
SolCallback(18)="vpmSolSound SoundFX(""Sling"",DOFContactors),"
SolCallback(19)="vpmSolDiverter TopDiverter,True,"
SolCallback(20)="vpmSolSound SoundFX(""Sling"",DOFContactors),"
SolCallback(21)="dtDrop.SolunHit 4,"
SolCallback(22)="dtDrop.SolunHit 5,"
SolCallback(25)="Flash25" 'Sword ans Stars Upper Playfield
SolCallback(26)="Flash26" 'SolFlasher 2C
SolCallback(27)="Flash27" 'SolFlasher 4C
SolCallback(28)="Flash28" 'SolFlasher 3C
SolCallback(29)="Flash29" 'SolFlasher 5C
SolCallback(30)="Flash30" 'SolFlasher 6C
SolCallback(31)="Flash31" 'SolFlasher 7C
SolCallback(32)="Flash32" 'SolFlasher 8C

'SolCallback(sLLFlipper)="vpmSolFlipper LeftFlipper,Flipper1,"
'SolCallback(sLRFlipper)="vpmSolFlipper RightFlipper,Flipper2,"

	SolCallback(sLRFlipper) = "SolRFlipper"
	SolCallback(sLLFlipper) = "SolLFlipper"

	Sub SolLFlipper(Enabled)
		 If Enabled Then
			 PlaySoundAtVol SoundFX("fx_FlipperUp",DOFContactors),LeftFlipper,VolFlip:LeftFlipper.RotateToEnd:Flipper1.RotateToEnd
		 Else
			 PlaySoundAtVol SoundFX("fx_FlipperDown",DOFContactors),LeftFlipper,VolFlip:LeftFlipper.RotateToStart:Flipper1.RotateToStart
		 End If
	  End Sub

	Sub SolRFlipper(Enabled)
		 If Enabled Then
			 PlaySoundAtVol SoundFX("fx_FlipperUp",DOFContactors),RightFlipper,VolFlip:RightFlipper.RotateToEnd:Flipper2.RotateToEnd
		 Else
			 PlaySoundAtVol SoundFX("fx_FlipperDown",DOFContactors),RightFlipper,VolFlip:RightFlipper.RotateToStart:Flipper2.RotateToStart
		 End If
	End Sub

Sub Flash9(enabled)
If enabled Then
F9.State = 1
F9a.State = 1
else
F9.State = 0
F9a.State = 0
End If
End Sub

Sub Flash16(enabled)
If enabled Then
F16.State = 1
F16a.State = 1
else
F16.State = 0
F16a.State = 0
End If
End Sub

Sub Flash25(enabled)
If enabled Then
F25.State = 1
F25a.State = 1
F25b.State = 1
F25b1.State = 1
F25b2.State = 1
F25b3.State = 1
F25b4.State = 1
F25b5.State = 1
F25b6.State = 1
F25b7.State = 1
F25b8.State = 1
F25b9.State = 1
F25b10.State = 1
F25b11.State = 1
F25b12.State = 1
F25b13.State = 1
F25b14.State = 1
F25b15.State = 1
F25b16.State = 1
F25b17.State = 1
F25b18.State = 1
F25b19.State = 1
F25b20.State = 1
F25b21.State = 1
F25b22.State = 1
F25b23.State = 1
F25b24.State = 1
F25b25.State = 1
F25b26.State = 1
F25b27.State = 1
F25b28.State = 1
F25b29.State = 1
F25b30.State = 1
F25b31.State = 1
F25b32.State = 1
F25b33.State = 1
F25b34.State = 1
F25b35.State = 1
F25b36.State = 1
F25b37.State = 1
F25b38.State = 1
F25b39.State = 1
else
F25.State = 0
F25a.State = 0
F25b.State = 0
F25b1.State = 0
F25b2.State = 0
F25b3.State = 0
F25b4.State = 0
F25b5.State = 0
F25b6.State = 0
F25b7.State = 0
F25b8.State = 0
F25b9.State = 0
F25b10.State = 0
F25b11.State = 0
F25b12.State = 0
F25b13.State = 0
F25b14.State = 0
F25b15.State = 0
F25b16.State = 0
F25b17.State = 0
F25b18.State = 0
F25b19.State = 0
F25b20.State = 0
F25b21.State = 0
F25b22.State = 0
F25b23.State = 0
F25b24.State = 0
F25b25.State = 0
F25b26.State = 0
F25b27.State = 0
F25b28.State = 0
F25b29.State = 0
F25b30.State = 0
F25b31.State = 0
F25b32.State = 0
F25b33.State = 0
F25b34.State = 0
F25b35.State = 0
F25b36.State = 0
F25b37.State = 0
F25b38.State = 0
F25b39.State = 0
End If
End Sub

Sub Flash26(enabled)
If enabled Then
		L23.State = 1
		L24.State = 1
else
		L23.State = 0
		L24.State = 0
End If
End Sub

Sub Flash27(enabled)
If enabled Then
F27.State = 1
F27a.State = 1
else
F27.State = 0
F27a.State = 0
End If
End Sub

Sub Flash28(enabled)
If enabled Then
F28.State = 1
F28a.State = 1
else
F28.State = 0
F28a.State = 0
End If
End Sub

Sub Flash29(enabled)
If enabled Then
F29.State = 1
F29a.State = 1
else
F29.State = 0
F29a.State = 0
End If
End Sub

Sub Flash30(enabled)
If enabled Then
F30.State = 1
F30a.State = 1
else
F30.State = 0
F30a.State = 0
End If
End Sub

Sub Flash31(enabled)
If enabled Then
F31.State = 1
F31a.State = 1
F31b.State = 1
F31b1.State = 1
else
F31.State = 0
F31a.State = 0
F31b.State = 0
F31b1.State = 0
End If
End Sub

Sub Flash32(enabled)
If enabled Then
F32.State = 1
F32a.State = 1
F32e.State = 1
F32e1.State = 1
F32e2.State = 1
F32f.State = 1
F32f1.State = 1
else
F32.State = 0
F32a.State = 0
F32e.State = 0
F32e1.State = 0
F32e2.State = 0
F32f.State = 0
F32f1.State = 0
End If
End Sub


Sub SolKickBack(enabled)
    If enabled Then
       KickBack.Fire
       PlaySoundAtVol SoundFX("plunger",DOFContactors), KickBack, 1
    Else
       KickBack.PullBack
    End If
End Sub


'*********
'Switches
'*********

Sub Drain_Hit()
	PlaySoundAtVol "drain", Drain, 1
    bsTrough.AddBall Me
End Sub

Sub S14_Hit:Controller.Switch(14)=1:End Sub
Sub S14_unHit:Controller.Switch(14)=0:End Sub
Sub S15_Hit:Controller.Switch(15)=1:End Sub
Sub S15_unHit:Controller.Switch(15)=0:End Sub
Sub S16_Hit:Controller.Switch(16)=1:End Sub
Sub S16_unHit:Controller.Switch(16)=0:End Sub
Sub Spinner1_Spin:vpmTimer.PulseSw 23:End Sub
Sub Spinner2_Spin:vpmTimer.PulseSw 24:End Sub
Sub S25_Hit:Controller.Switch(25)=1:End Sub
Sub S25_unHit:Controller.Switch(25)=0:End Sub
Sub sw26_Hit:vpmTimer.PulseSw 26:PlaySoundAtVol SoundFX("target",DOFContactors), ActiveBall, VolTarg:End Sub
Sub sw27_Hit:vpmTimer.PulseSw 27:PlaySoundAtVol SoundFX("target",DOFContactors), ActiveBall, VolTarg:End Sub
Sub sw28_Hit:vpmTimer.PulseSw 28:PlaySoundAtVol SoundFX("target",DOFContactors), ActiveBall, VolTarg:End Sub
Sub sw29_Hit:vpmTimer.PulseSw 29:PlaySoundAtVol SoundFX("target",DOFContactors), ActiveBall, VolTarg:End Sub
Sub sw30_Hit:vpmTimer.PulseSw 30:PlaySoundAtVol SoundFX("target",DOFContactors), ActiveBall, VolTarg:End Sub
Sub sw31_Hit:vpmTimer.PulseSw 31:PlaySoundAtVol SoundFX("target",DOFContactors), ActiveBall, VolTarg:End Sub
Sub sw32_Hit:vpmTimer.PulseSw 32:PlaySoundAtVol SoundFX("target",DOFContactors), ActiveBall, VolTarg:End Sub
Sub S37_Hit:vpmTimer.PulseSw 37:End Sub
Sub S38_Hit:Controller.Switch(38)=1:End Sub
Sub S38_unHit:Controller.Switch(38)=0:End Sub
Sub Spinner7_Spin:vpmTimer.PulseSw 39:End Sub
Sub S40_Hit:vpmTimer.PulseSw 40:End Sub
Sub S46_Hit:Controller.Switch(46)=1:End Sub
Sub S46_unHit:Controller.Switch(46)=0:End Sub
Sub S47_Hit:Controller.Switch(47)=1:End Sub
Sub S47_unHit:Controller.Switch(47)=0:End Sub
Sub S48_Hit:Controller.Switch(48)=1:End Sub
Sub S48_unHit:Controller.Switch(48)=0:End Sub
Sub S54_Hit:bsSaucer.Addball 0:End Sub
Sub S57_Hit:Controller.Switch(57)=1:End Sub
Sub S57_unHit:Controller.Switch(57)=0:End Sub
Sub S59_Hit:Controller.Switch(59)=1:End Sub
Sub S59_unHit:Controller.Switch(59)=0:End Sub
Sub S61_Hit:Controller.Switch(61)=1:End Sub
Sub S61_unHit:Controller.Switch(61)=0:End Sub
Sub S62_Hit:Controller.Switch(62)=1:End Sub
Sub S62_unHit:Controller.Switch(62)=0:End Sub


'****Targets

Set dtDrop=New cvpmDropTarget
dtDrop.InitDrop Array(sw17,sw18,sw19,sw20,sw21),Array(17,18,19,20,21)
dtDrop.InitSnd SoundFX("DTDrop",DOFContactors),SoundFX("DTReset",DOFContactors)
dtDrop.CreateEvents"dtDrop"


'****Kicker

 Dim BallSpeed
 Sub Kicker4_Hit:BallSpeed=ABS(ActiveBall.VelY):Kicker4.DestroyBall:Kicker3.CreateBall:Kicker3.Kick 180,BallSpeed/20:End Sub


'*****GI Lights On
dim xx
Sub SolGIUpdate(enabled)
If enabled Then
For each xx in GI:xx.State = 0: Next
        Jackpot.SetValue 1
else
For each xx in GI:xx.State = 1: Next
        Jackpot.SetValue 0
End If
End Sub

Set LampCallback = GetRef("Lamps")
	Sub Lamps
		L46.State = Controller.Lamp(46)
		L46a.State = Controller.Lamp(46)
		L46b.State = Controller.Lamp(46)
		L46c.State = Controller.Lamp(46)
		L47.State = Controller.Lamp(47)
		L47a.State = Controller.Lamp(47)
		L47b.State = Controller.Lamp(47)
		L47c.State = Controller.Lamp(47)
		L48.State = Controller.Lamp(48)
		L48a.State = Controller.Lamp(48)
		L48b.State = Controller.Lamp(48)
		L48c.State = Controller.Lamp(48)
End Sub

'**************************************
' Backglass EM Reels
'**************************************

Set motorcallback = GetRef("UpdateMultipleLamps")

 Sub UpdateMultipleLamps
 B6.SetValue ABS(Controller.Lamp(6) )
 B7.SetValue ABS(Controller.Lamp(7) )
 B41.SetValue ABS(Controller.Lamp(41) )
 B42.SetValue ABS(Controller.Lamp(42) )
 B51.SetValue ABS(Controller.Lamp(51) )
 B52.SetValue ABS(Controller.Lamp(52) )
 B57.SetValue ABS(Controller.Lamp(57) )
 B58.SetValue ABS(Controller.Lamp(58) )
 B59.SetValue ABS(Controller.Lamp(59) )
 B60.SetValue ABS(Controller.Lamp(60) )
 B61.SetValue ABS(Controller.Lamp(61) )
 B62.SetValue ABS(Controller.Lamp(62) )
 B63.SetValue ABS(Controller.Lamp(63) )
 B64.SetValue ABS(Controller.Lamp(64) )
 TDP.ObjRotZ = TopDiverter.CurrentAngle - 137
 BDP.ObjRotZ = BottomDiverter.CurrentAngle - 137
End Sub



'VPM Light Subs - VP10 Built in Fading Routines are used.

Set Lights(1) = L1
Set Lights(2) = L2
Set Lights(3) = L3
Set Lights(4) = L4
Set Lights(5) = L5
Set Lights(6) = L6
Set Lights(7) = L7

Set Lights(9) = L9
Set Lights(10) = L10
Set Lights(11) = L11
Set Lights(12) = L12
Set Lights(13) = L13
Set Lights(14) = L14
Set Lights(15) = L15
Set Lights(16) = L16
Set Lights(17) = L17
Set Lights(18) = L18
Set Lights(19) = L19
Set Lights(20) = L20
Set Lights(21) = L21
Set Lights(22) = L22
Set Lights(23) = L23
Set Lights(24) = L24

Set Lights(26) = L26
Set Lights(27) = L27
Set Lights(28) = L28
Set Lights(29) = L29
Set Lights(30) = L30
Set Lights(31) = L31
Set Lights(32) = L32
Set Lights(33) = L33
Set Lights(34) = L34
Set Lights(35) = L35
Set Lights(36) = L36

Set Lights(38) = L38
Set Lights(39) = L39
Set Lights(40) = L40
Set Lights(41) = L41
Set Lights(42) = L42
Set Lights(43) = L43
Set Lights(44) = L44
Set Lights(45) = L45

Set Lights(49) = L49
Set Lights(50) = L50
Set Lights(51) = L51
Set Lights(52) = L52
Set Lights(53) = L53
Set Lights(54) = L54
Set Lights(55) = L55
Set Lights(56) = L56

Sub LampTimer_Timer()
l1f.visible = L1.state
l2f.visible = L2.state
l3f.visible = L3.state
l4f.visible = L4.state
l5f.visible = L5.state
l6f.visible = L6.state
l7f.visible = L7.state
l9f.visible = L9.state
l9f1.visible = L9.state
l9f2.visible = L9.state
l10f.visible = L10.state
l10f1.visible = L10.state
l10f2.visible = L10.state
l11f.visible = L11.state
l11f1.visible = L11.state
l11f2.visible = L11.state
l12f.visible = L12.state
l12f1.visible = L12.state
l12f2.visible = L12.state
l13f.visible = L13.state
l14f.visible = L14.state
l15f.visible = L15.state
l16f.visible = L16.state
l16f1.visible = L16.state
l16f2.visible = L16.state

F17.visible = L17.state
F18.visible = L18.state
F19.visible = L19.state
F20.visible = L20.state
F21.visible = L21.state

l22f.visible = L22.state
l23f.visible = L23.state
l24f.visible = L24.state

l26f.visible = L26.state
l27f.visible = L27.state
l28f.visible = L28.state
l29f.visible = L29.state
l30f.visible = L30.state
l31f.visible = L31.state
l32f.visible = L32.state
l33f.visible = L33.state
l34f.visible = L34.state
l35f.visible = L35.state
l36f.visible = L36.state

l38f.visible = L38.state
l39f.visible = L39.state
l40f.visible = L40.state
l41f.visible = L41.state
l42f.visible = L42.state
l43f.visible = L43.state
l44f.visible = L44.state
l45f.visible = L45.state

l49f.visible = L49.state
l49f1.visible = L49.state
l49f2.visible = L49.state
l50f.visible = L50.state
l50f1.visible = L50.state
l50f2.visible = L50.state
l51f.visible = L51.state
l52f.visible = L52.state
l53f.visible = L53.state
l54f.visible = L54.state
l55f.visible = L55.state
l56f.visible = L56.state

	if MagicTunnelLight = 1 then
		F32b.state = Lightstateon
		F32c.state = Lightstateon
	else
		F32b.state = Lightstateoff
		F32c.state = Lightstateoff

	end if

End Sub


'MagicTunnelLight
Sub MTLOn_Hit()
	MagicTunnelLight = 1
End sub

Sub MTLOff1_Hit()
	MagicTunnelLight = 0
End sub

Sub MTLOff2_Hit()
	MagicTunnelLight = 0
End sub

Sub MTLOff3_Hit()
	MagicTunnelLight = 0
End sub


'**********Sling Shot Animations
' Rstep and Lstep  are the variables that increment the animation
'****************
Dim Lstep, RStep

Sub LeftSlingShot_Slingshot
    PlaySoundAtVol SoundFX("left_slingshot",DOFContactors), sling2, 1
    LSling.Visible = 0
    LSling1.Visible = 1
    sling2.TransZ = -20
    LStep = 0
    vpmTimer.PulseSw 63
    LeftSlingShot.TimerEnabled = 1
	LeftSlingShot.TimerInterval = 10
End Sub

Sub LeftSlingShot_Timer
    Select Case LStep
        Case 3:LSLing1.Visible = 0:LSLing2.Visible = 1:sling2.TransZ = -10
        Case 4:LSLing2.Visible = 0:LSLing.Visible = 1:sling2.TransZ = 0:LeftSlingShot.TimerEnabled = 0
    End Select
    LStep = LStep + 1
End Sub

Sub RightSlingShot_Slingshot
    PlaySoundAtVol SoundFX("right_slingshot",DOFContactors), sling1, 1
    RSling.Visible = 0
    RSling1.Visible = 1
    sling1.TransZ = -20
    RStep = 0
    vpmTimer.PulseSw 64
    RightSlingShot.TimerEnabled = 1
    RightSlingShot.TimerInterval = 10
End Sub

Sub RightSlingShot_Timer
    Select Case RStep
        Case 3:RSLing1.Visible = 0:RSLing2.Visible = 1:sling1.TransZ = -10
        Case 4:RSLing2.Visible = 0:RSLing.Visible = 1:sling1.TransZ = 0:RightSlingShot.TimerEnabled = 0
    End Select
    RStep = RStep + 1
End Sub

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

'**********************
' Ball Bounce Sound
'**********************

Sub Bounce1_Hit()
  PlaySoundAtVol "Ball_Bounce", ActiveBall, 1
End Sub

Sub Bounce2_Hit()
  PlaySoundAtVol "Ball_Bounce", ActiveBall, 1
End Sub

Sub Bounce3_Hit()
  PlaySoundAtVol "Ball_Bounce", ActiveBall, 1
End Sub

Sub Bounce4_Hit()
  PlaySoundAtVol "Ball_Bounce", ActiveBall, 1
End Sub

'*****************************************
'			BALL SHADOW
'*****************************************
Dim BallShadow
BallShadow = Array (BallShadow1,BallShadow2,BallShadow3,BallShadow4,BallShadow5,BallShadow6,BallShadow7,BallShadow8,BallShadow9,BallShadow10)

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
            BallShadow(b).X = ((BOT(b).X) - (Ballsize/6) + ((BOT(b).X - (Table1.Width/2))/7)) '+ 13
        Else
            BallShadow(b).X = ((BOT(b).X) + (Ballsize/6) + ((BOT(b).X - (Table1.Width/2))/7)) '- 13
        End If
        ballShadow(b).Y = BOT(b).Y + 10
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

' The sound is played using the VOL, PAN and PITCH functions, so the volume and pitch of the sound
' will change according to the ball speed, and the PAN function will change the stereo position according
' to the position of the ball on the table.


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

Sub Rubbers_Hit(idx)
 	dim finalspeed
  	finalspeed=SQR(activeball.velx * activeball.velx + activeball.vely * activeball.vely)
 	If finalspeed > 20 then
		PlaySound "fx_rubber2", 0, Vol(ActiveBall)*VolRH, Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0
	End if
	If finalspeed >= 6 AND finalspeed <= 20 then
 		RandomSoundRubber()
 	End If
End Sub

Sub Posts_Hit(idx)
 	dim finalspeed
  	finalspeed=SQR(activeball.velx * activeball.velx + activeball.vely * activeball.vely)
 	If finalspeed > 16 then
		PlaySound "fx_rubber2", 0, Vol(ActiveBall)*VolPo, Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0
	End if
	If finalspeed >= 6 AND finalspeed <= 16 then
 		RandomSoundRubber()
 	End If
End Sub

Sub RandomSoundRubber()
	Select Case Int(Rnd*3)+1
		Case 1 : PlaySound "rubber_hit_1", 0, Vol(ActiveBall)*VolRH, Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0
		Case 2 : PlaySound "rubber_hit_2", 0, Vol(ActiveBall)*VolRH, Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0
		Case 3 : PlaySound "rubber_hit_3", 0, Vol(ActiveBall)*VolRH, Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0
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
		Case 1 : PlaySound "flip_hit_1", 0, Vol(ActiveBall)*VolRH, Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0
		Case 2 : PlaySound "flip_hit_2", 0, Vol(ActiveBall)*VolRH, Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0
		Case 3 : PlaySound "flip_hit_3", 0, Vol(ActiveBall)*VolRH, Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0
	End Select
End Sub

Sub Table1_exit()
	If B2SOn Then Controller.Stop
End Sub

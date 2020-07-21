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
LoadVPM "01520000", "GamePlan.vbs", 3.1
Dim cCredits
Dim FlipperShadows
Dim DesktopMode: DesktopMode = Table1.ShowDT

FlipperShadows = 1  	' 1 turns on, 0 turns off flipper shadows.

cCredits="Sharpshooter 2"
Const cGameName="sshootr2",UseSolenoids=2,UseLamps=1,UseGI=0,UseSync=1,SCoin="coin3"',SFlipperOn="FlipperUp",SFlipperOff="FlipperDown"

'**********************************************************************************************************

'Solenoid Controlled toys
'**********************************************************************************************************

'*****GI Lights On
sub Encendido_timer
dim xx
playsound "encendido"
For each xx in Ambiente:xx.State = 1: Next
	me.enabled=false
end sub

Const sBallRelease=8
Const sKicker=15
Const sDropA=11
Const sEnable=16

SolCallback(sBallRelease)="bsTrough.SolOut"
SolCallback(sKicker)="bsSaucer.SolOut"
SolCallback(sDropA)="dtDrop.SolDropUp"
SolCallback(3)="vpmSolSound ""fx_Knocker"","
SolCallback(sEnable)="vpmNudge.SolGameOn"
SolCallback(sLLFlipper)="SolLFlipper"
SolCallback(sLRFlipper)="SolRFlipper"

'**********************************************************************************************************

'Initiate Table
'**********************************************************************************************************

Dim bsTrough,dtDrop,bsSaucer
Dim bump1,bump2,bump3,bump4

Sub Table1_Init
vpmInit Me
	On Error Resume Next
	With Controller
		.GameName=cGameName
		If Err Then MsgBox "Can't start Game" & cGameName & vbNewLine & Err.Description:Exit Sub
		.SplashInfoLine = "Sharpshooter II (Gameplan 1983)" & vbNewLine & "VPX Table By Kalavera"
		.HandleMechanics=0
		.ShowDMDOnly=1
		.ShowFrame=0
		.ShowTitle=0
		.Run
        .Hidden=1
		If Err Then MsgBox Err.Description
	End With
	On Error Goto 0
'	Controller.Dip(0) = (0*1 + 1*2 + 0*4 + 0*8 + 0*16 + 0*32 + 0*64 + 0*128) '01-08
'	Controller.Dip(1) = (0*1 + 0*2 + 0*4 + 0*8 + 0*16 + 0*32 + 0*64 + 1*128) '09-16
'	Controller.Dip(2) = (0*1 + 1*2 + 0*4 + 0*8 + 0*16 + 0*32 + 0*64 + 0*128) '17-24
'	Controller.Dip(3) = (1*1 + 1*2 + 1*4 + 0*8 + 1*16 + 1*32 + 1*64 + 0*128) '25-32

	PinMAMETimer.Interval=PinMAMEInterval
	PinMAMETimer.Enabled=True

	vpmNudge.TiltSwitch=swTilt
	vpmNudge.Sensitivity=5
	vpmNudge.TiltObj=Array(Bumper1,Bumper2,Bumper3,Bumper4,LeftSlingshot,RightSlingshot,Sling2,Sling5)

	Set dtdrop = New cvpmDropTarget
   	  With dtdrop
   		.InitDrop Array(TargetS,TargetH,TargetO1,TargetO2,TargetT,TargetE,TargetR),Array(31,32,35,36,4,10,17)
         .Initsnd SoundFX("fx_droptarget", DOFDropTargets), SoundFX("fx2_DTReset", DOFContactors)
       End With

	Set bsSaucer = New cvpmBallStack
             With bsSaucer
              .InitSaucer Kicker1,24, 200, 10
      		 .KickForceVar = 1
      		 .KickAngleVar = 1
              .InitExitSnd "fx_kicker", "Solenoid"
    		     .InitAddSnd "fx2_popper_ball"
             End With

	Set bsTrough = New cvpmBallStack
            With bsTrough
             .InitSw 0,11,0,0,0,0,0,0
      	     .InitKick BallRelease,90,8
             bsTrough.InitExitSnd "fx_ballrel", "fx_Solenoid"
             .Balls = 1
            End With
End Sub

	Sub table1_Paused
	Controller.Pause=True
	End Sub

	Sub table1_UnPaused
	Controller.Pause=False
	End Sub

	Sub table1_Exit
	Controller.Stop
	End Sub


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

Sub SolKnocker(Enabled)
	If Enabled Then PlaySound SoundFX("fx_Knocker",DOFKnocker)
End Sub

'**********************************************************************************************************
'Plunger code
'**********************************************************************************************************

Sub Table1_KeyDown(ByVal KeyCode)
	If keycode = LeftTiltKey Then Nudge 90, 2
	If keycode = RightTiltKey Then Nudge 270, 2
	If keycode = CenterTiltKey Then	Nudge 0, 2

	If vpmKeyDown(keycode) Then Exit Sub
	If keycode = PlungerKey Then Plunger.PullBack: PlaySoundAtVol "fx_plungerpull", Plunger, 1: 	End If
	End Sub

Sub Table1_KeyUp(ByVal KeyCode)
	If keycode = PlungerKey Then Plunger.Fire: PlaySoundAtVol "fx_plunger", ActiveBall, 1
	If vpmKeyUp(keycode) Then Exit Sub
End Sub

If FlipperShadows = 1 then
        FlipperLSh.visible=1
        FlipperRSh.visible=1
       else
        FlipperLSh.visible=0
        FlipperRSh.visible=0
    End If
'**********************************************************************************************************
' Drain hole and kickers
'**********************************************************************************************************

Sub Drain_Hit:PlaysoundAtVol "fx2_drain2",Drain,1:bsTrough.AddBall Me:End Sub
Sub Kicker1_Hit:bsSaucer.AddBall 0:End Sub
'Sub TargetS_Hit:dtDrop.Hit 1:End Sub
'Sub TargetH_Hit:dtDrop.Hit 2:End Sub
'Sub TargetO1_Hit:dtDrop.Hit 3:End Sub
'Sub TargetO2_Hit:dtDrop.Hit 4:End Sub
'Sub TargetT_Hit:dtDrop.Hit 5:End Sub
'Sub TargetE_Hit:dtDrop.Hit 6:End Sub
'Sub TargetR_Hit:dtDrop.Hit 7:End Sub
Sub TargetS_Dropped:dtDrop.Hit 1:End Sub
Sub TargetH_Dropped:dtDrop.Hit 2:End Sub
Sub TargetO1_Dropped:dtDrop.Hit 3:End Sub
Sub TargetO2_Dropped:dtDrop.Hit 4:End Sub
Sub TargetT_Dropped:dtDrop.Hit 5:End Sub
Sub TargetE_Dropped:dtDrop.Hit 6:End Sub
Sub TargetR_Dropped:dtDrop.Hit 7:End Sub
Sub RightSlingshot_Hit:vpmTimer.PulseSw 9:End Sub
Sub ShooterSlingshot_Hit:VpmTimer.PulseSw 9:End Sub
Sub Sling2_Hit:vpmTimer.PulseSw 9:End Sub
Sub Wall34_Hit:vpmTimer.PulseSw 9:End Sub
Sub rightwalla_Hit:vpmTimer.PulseSw 9:End Sub
Sub rightwallb_Hit:vpmTimer.PulseSw 9:End Sub
Sub rightangle_Hit:vpmTimer.PulseSw 9:End Sub
Sub Lane3_Hit:Controller.Switch(12)=1:End Sub
Sub Lane3_unHit:Controller.Switch(12)=0:End Sub
Sub Lane5_Hit:Controller.Switch(13)=1:End Sub
Sub Lane5_unHit:Controller.Switch(13)=0:End Sub
Sub Lane6_Hit:Controller.Switch(14)=1:End Sub
Sub Lane6_unHit:Controller.Switch(14)=0:End Sub
Sub LeftInlane_Hit:Controller.Switch(16)=1:End Sub
Sub LeftInlane_unHit:Controller.Switch(16)=0:End Sub
Sub LeftOutlane_Hit:Controller.Switch(18)=1:End Sub
Sub LeftOutlane_unHit:Controller.Switch(18)=0:End Sub
Sub Spot1_Hit:vpmTimer.PulseSw(19):End Sub
Sub Spot2_Hit:vpmTimer.PulseSw(20):End Sub
Sub T1_Hit:vpmTimer.PulseSw 19:PlaySoundAtVol SoundFX("FX2_Target", DOFDropTargets), ActiveBall, 1:End Sub
Sub T2_Hit:vpmTimer.PulseSw 20:PlaySoundAtVol SoundFX("FX2_Target", DOFDropTargets), ActiveBall, 1:End Sub

Sub Bumper1_Hit:RandomSoundBumper:vpmTimer.PulseSw 34:bump1 = 1:Me.TimerEnabled = 1:End Sub
 Sub Bumper1_Timer()
 if Bumper01.state=0 then
	Bumper01a.state=0
	Bumper01b.state=0
	Bumper01c.state=0
	end If
 if Bumper01.state=1 then
	Bumper01a.state=1
	Bumper01b.state=1
	Bumper01c.state=1
	end If
  End Sub
Sub Bumper2_Hit:RandomSoundBumper:vpmTimer.PulseSw 33:bump2 = 1:Me.TimerEnabled = 1:End Sub
 Sub Bumper2_Timer()
 if Bumper02.state=0 then
	Bumper02a.state=0
	Bumper02b.state=0
	Bumper02c.state=0
	end if
 if Bumper02.state=1 then
	Bumper02a.state=1
	Bumper02b.state=1
	Bumper02c.state=1
	end if
End Sub
Sub Bumper3_Hit:RandomSoundBumper:vpmTimer.PulseSw 21:bump3 = 1:Me.TimerEnabled = 1:End Sub
Sub Bumper4_Hit:RandomSoundBumper:vpmTimer.PulseSw 22:bump4 = 1:Me.TimerEnabled = 1:End Sub

Sub RandomSoundBumper()
	Select Case Int(Rnd*3)+1
		Case 1 : PlaySoundAtVol SoundFX("fx2_bumper_1", DOFContactors), ActiveBall, VolTarg
		Case 2 : PlaySoundAtVol SoundFX("fx2_bumper_2", DOFContactors), ActiveBall, VolTarg
		Case 3 : PlaySoundAtVol SoundFX("fx2_bumper_3", DOFContactors), ActiveBall, VolTarg
	End Select
End Sub


Sub Spinner_Spin:PlaySoundAtVol "fx_spinner", Spinner, VolSpin:vpmTimer.PulseSw 23:End Sub

Sub Roll1_Hit:Controller.Switch(25)=1:End Sub
Sub Roll1_unHit:Controller.Switch(25)=0:End Sub
Sub Roll2_Hit:Controller.Switch(27)=1:End Sub
Sub Roll2_unHit:Controller.Switch(27)=0:End Sub
Sub Roll3_Hit:Controller.Switch(28)=1:End Sub
Sub Roll3_unHit:Controller.Switch(28)=0:End Sub
Sub Roll4_Hit:Controller.Switch(29)=1:End Sub
Sub Roll4_unHit:Controller.Switch(29)=0:End Sub
Sub Roll5_Hit:Controller.Switch(30)=1:End Sub
Sub Roll5_unHit:Controller.Switch(30)=0:End Sub

Sub Lane1_Hit:Controller.Switch(37)=1:End Sub
Sub Lane1_unHit:Controller.Switch(37)=0:End Sub
Sub Lane2_Hit:Controller.Switch(38)=1:End Sub
Sub Lane2_unHit:Controller.Switch(38)=0:End Sub
Sub Lane4_Hit:Controller.Switch(39)=1:End Sub
Sub Lane4_unHit:Controller.Switch(39)=0:End Sub
Sub Lane7_Hit:Controller.Switch(40)=1:End Sub
Sub Lane7_unHit:Controller.Switch(40)=0:End Sub



Set Lights(1)=Light1
Set Lights(2)=Light2
Set Lights(3)=Light3
Set Lights(4)=Light4
Set Lights(5)=Light5
Set Lights(6)=Light6
Set Lights(7)=Light7
Set Lights(8)=Light8
Set Lights(9)=Light9
Set Lights(10)=Light10
Set Lights(11)=Light11
Set Lights(12)=Light12
Set Lights(13)=Light13
Set Lights(14)=Light14
Set Lights(15)=Light15
Set Lights(16)=Light16
Set Lights(17)=Light17		'Special 1
Set Lights(18)=Light18		'Special 2
Lights(19)=Array(RollXB,RollXB1,RollXB2,RollXB3)		'Extra Ball 1
Set Lights(20)=Light20		'Extra Ball 2
Set Lights(21)=Light21
Set Lights(22)=Light22
Set Lights(23)=Light23
Set Lights(24)=Light24
Set Lights(25)=Light25
Set Lights(26)=Light26
Set Lights(27)=Light27
Set Lights(28)=Light28
Set Lights(29)=Light29
Set Lights(30)=Light30
Set Lights(31)=Light31
Set Lights(32)=Light32
Set Lights(33)=Light33
Set Lights(34)=Light34
Set Lights(35)=Light35
Set Lights(36)=Light36
Set Lights(37)=Light37
Set Lights(38)=Light38
Lights(39)=Array(Bumper03,Bumper02,Bumper03dt,Bumper02dt)
Set Lights(40)=Light40		'Tilt
Lights(41)=Array(GI101,GI102,GI103,Light41)		'High score
Set Lights(42)=Light42
Lights(43)=Array(Bumper01,Bumper04,Bumper01dt,Bumper04dt)
Set Lights(44)=Light44		'Game over
Set Lights(45)=Light45
Set Lights(46)=Light46
Set Lights(47)=Light47
Set Lights(48)=Light48
Set Lights(49)=Light49
Set Lights(50)=Light50
Set Lights(51)=Light51
Set Lights(52)=Light52		'Shoot again
Lights(53)=Array(Light53,Light53a)		'Ball in play
Set Lights(54)=Light54
Set Lights(55)=Light55
Set Lights(56)=Light56
Set Lights(57)=Light57		'Player1
Set Lights(58)=Light58		'Player2
Set Lights(59)=Light59		'Player3
Set Lights(60)=Light60		'Player4
Set Lights(61)=Light61
Set Lights(62)=Light62
Set Lights(63)=Light63
Set Lights(64)=Light64

'**********************************************************************************************************
'Digital Display
'**********************************************************************************************************

Dim Digits(32)
' 1st Player
Digits(0) = Array(a1,a2,a3,a4,a5,a6,a7)
Digits(1) = Array(a8,a9,a10,a11,a12,a13,a14)
Digits(2) = Array(a15,a16,a17,a18,a19,a20,a21)
Digits(3) = Array(a22,a23,a24,a25,a26,a27,a28)
Digits(4) = Array(a29,a30,a31,a32,a33,a34,a35)
Digits(5) = Array(a36,a37,a38,a39,a40,a41,a42)
Digits(6) = Array(a43,a44,a45,a46,a47,a48,a49)

' 2nd Player
Digits(7) = Array(a50,a51,a52,a53,a54,a55,a56)
Digits(8) = Array(a57,a58,a59,a60,a61,a62,a63)
Digits(9) = Array(a64,a65,a66,a67,a68,a69,a70)
Digits(10) = Array(a71,a72,a73,a74,a75,a76,a77)
Digits(11) = Array(a78,a79,a80,a81,a82,a83,a84)
Digits(12) = Array(a85,a86,a87,a88,a89,a90,a91)
Digits(13) = Array(a92,a93,a94,a95,a96,a97,a98)

' 3rd Player
Digits(14) = Array(a99,a100,a101,a102,a103,a104,a105)
Digits(15) = Array(a106,a107,a108,a109,a110,a111,a112)
Digits(16) = Array(a113,a114,a115,a116,a117,a118,a119)
Digits(17) = Array(a120,a121,a122,a123,a124,a125,a126)
Digits(18) = Array(a127,a128,a129,a130,a131,a132,a133)
Digits(19) = Array(a134,a135,a136,a137,a138,a139,a140)
Digits(20) = Array(a141,a142,a143,a144,a145,a146,a147)

' 4th Player
Digits(21) = Array(a148,a149,a150,a151,a152,a153,a154)
Digits(22) = Array(a155,a156,a157,a158,a159,a160,a161)
Digits(23) = Array(a162,a163,a164,a165,a166,a167,a168)
Digits(24) = Array(a169,a170,a171,a172,a173,a174,a175)
Digits(25) = Array(a176,a177,a178,a179,a180,a181,a182)
Digits(26) = Array(a183,a184,a185,a186,a187,a188,a189)
Digits(27) = Array(a190,a191,a192,a193,a194,a195,a196)

' Credits
Digits(28) = Array(a197,a198,a199,a200,a201,a202,a203)
Digits(29) = Array(a204,a205,a206,a207,a208,a209,a210)
' Balls
Digits(30) = Array(a211,a212,a213,a214,a215,a216,a217)
Digits(31) = Array(a218,a219,a220,a221,a222,a223,a224)

Sub DisplayTimer_Timer
	Dim ChgLED,ii,num,chg,stat,obj
	ChgLed = Controller.ChangedLEDs(&Hffffffff, &Hffffffff)
If Not IsEmpty(ChgLED) Then
		If DesktopMode = True Then
		For ii = 0 To UBound(chgLED)
			num = chgLED(ii, 0) : chg = chgLED(ii, 1) : stat = chgLED(ii, 2)
			if (num < 32) then
				For Each obj In Digits(num)
					If chg And 1 Then obj.State = stat And 1
					chg = chg\2 : stat = stat\2
				Next
			else
			end if
		next
		end if
end if
End Sub


'**********************************************************************************************************
'**********************************************************************************************************

'************Sling Shot Animations Rstep are the variables that increment the animation *******************

Dim RStep, Lstep, Lstep2

Sub RightSlingShot_Slingshot
vpmTimer.PulseSw 15
    PlaySoundAtVol SoundFX("right_slingshot",DOFContactors), pat6, 1
    dera.Visible = 0
    dera1.Visible = 1
    pat6.TransZ = -20
    RStep = 0
    RightSlingShot.TimerEnabled = 1

End Sub

Sub RightSlingShot_Timer
    Select Case RStep
        Case 3:dera1.Visible = 0:dera2.Visible = 1:pat6.TransZ = -10
        Case 4:dera2.Visible = 0:dera.Visible = 1:pat6.TransZ = 0::RightSlingShot.TimerEnabled = 0
    End Select
    RStep = RStep + 1
End Sub

Sub LeftSlingShot_Slingshot
    vpmTimer.PulseSw 15
    PlaySoundAtVol SoundFX("left_slingshot",DOFContactors), pat3, 1
    izq.Visible = 0
    izq1.Visible = 1
    pat3.rotx = 20
    LStep = 0
    LeftSlingShot.TimerEnabled = 1
End Sub

Sub LeftSlingShot_Timer
    Select Case LStep
        Case 3:izq1.Visible = 0:izq2.Visible = 1:pat3.rotx = 10
        Case 4:izq2.Visible = 0:izq.Visible = 1:pat3.rotx = 0:LeftSlingShot.TimerEnabled = 0
    End Select
    LStep = LStep + 1
End Sub

Sub Sling2_Slingshot
    vpmTimer.PulseSw 15
    PlaySoundAtVol SoundFX("left_slingshot",DOFContactors), pat1, 1
    arriba.Visible = 0
    arriba1.Visible = 1
    pat1.rotx = 20
    LStep2 = 0
    Sling2.TimerEnabled = 1
End Sub

Sub Sling2_Timer
    Select Case LStep2
        Case 3:arriba1.Visible = 0:arriba2.Visible = 1:pat1.rotx = 10
        Case 4:arriba2.Visible = 0:arriba.Visible = 1:pat1.rotx = 0:Sling2.TimerEnabled = 0
    End Select
    LStep2 = LStep2 + 1
End Sub

If DesktopMode = True Then 'Show Desktop components
	Chapa1.Visible = 1
	Chapa2.Visible = 1
	Chapa3.Visible = 1
	Bumper01.Visible = 0
	Bumper02.Visible = 0
	Bumper03.Visible = 0
	Bumper04.Visible = 0
	Bumper01dt.Visible = 1
	Bumper02dt.Visible = 1
	Bumper03dt.Visible = 1
	Bumper04dt.Visible = 1
	Light57.Visible = 1
	Light58.Visible = 1
	Light59.Visible = 1
	Light60.Visible = 1
	Lightcredits.Visible = 1
	Light40.Visible = 1
	Light53a.Visible = 1
	Light56.Visible = 1
	Light44.Visible = 1
	Light41.Visible = 1
else
	Chapa1.Visible = 0
	Chapa2.Visible = 0
	Chapa3.Visible = 0
	Bumper01.Visible = 1
	Bumper02.Visible = 1
	Bumper03.Visible = 1
	Bumper04.Visible = 1
	Bumper01dt.Visible = 0
	Bumper02dt.Visible = 0
	Bumper03dt.Visible = 0
	Bumper04dt.Visible = 0
	Light57.Visible = 0
	Light58.Visible = 0
	Light59.Visible = 0
	Light60.Visible = 0
	Lightcredits.Visible = 0
	Light40.Visible = 0
	Light53a.Visible = 0
	Light56.Visible = 0
	Light44.Visible = 0
	Light41.Visible = 0
end if

'******************************
' Diverse Collection Hit Sounds
'******************************
Sub aRollovers_Hit(idx):PlaySound "fx_sensor", 0, Vol(ActiveBall)*VolRol, pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0, AudioFade(ActiveBall):End Sub
Sub aGates_Hit(idx):PlaySound "fx_Gate", 0, Vol(ActiveBall)*VolGates, pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0, AudioFade(ActiveBall):End Sub
Sub aMetals_Hit(idx):PlaySound "fx_MetalHit", 0, Vol(ActiveBall)*VolMetal, pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0, AudioFade(ActiveBall):End Sub
Sub aRubber_Bands_Hit(idx):PlaySound "fx_rubber", 0, Vol(ActiveBall)*VolRB, pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0, AudioFade(ActiveBall):End Sub
Sub aRubbers_Hit(idx):PlaySound "fx_postrubber", 0, Vol(ActiveBall)*VolRH, pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0, AudioFade(ActiveBall):End Sub
Sub aRubber_Posts_Hit(idx):PlaySound "fx_postrubber", 0, Vol(ActiveBall)*VolPo, pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0, AudioFade(ActiveBall):End Sub
Sub aPlastics_Hit(idx):PlaySound "fx_PlasticHit", 0, Vol(ActiveBall)*VolPlast, pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0, AudioFade(ActiveBall):End Sub
Sub aWoods_Hit(idx):PlaySound "fx_Woodhit", 0, Vol(ActiveBall)*VolWood, pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0, AudioFade(ActiveBall):End Sub
' Sub Spiners_Hit(idx):PlaySound "spinner", 0, Vol(ActiveBall), pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0, AudioFade(ActiveBall):End Sub
Sub Kicker_Hit(idx):PlaySound "fx_kicker_enter", 0, Vol(ActiveBall)*VolKick, pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0, AudioFade(ActiveBall):End Sub
Sub Target_Hit (idx):PlaySoundAtVol "target", 0, Vol(ActiveBall)*VolTarg, Pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0, AudioFade(ActiveBall):End Sub
Sub Metals_Thin_Hit (idx):PlaySound "metalhit_thin", 0, Vol(ActiveBall)*VolMetal, Pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0, AudioFade(ActiveBall):End Sub
Sub Banda_Hit (idx):PlaySound "left_slingshot", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0, AudioFade(ActiveBall):End Sub

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
Const tnob = 1 ' total number of balls
ReDim rolling(tnob)
InitRolling

Sub InitRolling
    Dim i
    For i = 0 to tnob
        rolling(i) = False
    Next
End Sub

'Sub RollingUpdate()
Sub RollingTimer_Timer()
   Dim BOT, b, ballpitch
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



'Gameplan Sharpshooter II
 'added by Inkochnito
 Sub EditDips
 	Dim vpmDips:Set vpmDips=New cvpmDips
 	With vpmDips
 		.AddForm  700,400,"Sharpshooter II - DIP switches"
 		.AddFrame   2,5,150,"Maximum credits",&H07000000,Array("5 credits",0,"10 credits",&H01000000,"15 credits",&H02000000,"20 credits",&H03000000,"25 credits",&H04000000,"30 credits",&H05000000,"35 credits",&H06000000,"40 credits",&H07000000)'dip 25&26&27
 		.AddFrame   2,135,150,"High game to date award",&HC0000000,Array("no award",0,"1 credit",&H40000000,"2 credits",&H80000000,"3 credits",&HC0000000)'dip 31&32
 		.AddFrame   170,5,150,"Special award",&H10000000,Array("extra ball",0,"replay",&H10000000)'dip 29
 		.AddFrame   170,51,150,"Balls per game",&H08000000,Array("3 balls",0,"5 balls",&H08000000)'dip 28
 		.AddChk   170,157,150,Array("Free play",&H00000080)'dip 8
 		.AddChk   170,117,150,Array("Play tunes",32768)'dip 16
 		.AddChk   170,137,150,Array("Match feature",&H20000000)'dip 30
 		.AddLabel   30,230,300,20,"After hitting OK, press F3 to reset game with new settings."
 		.ViewDips
 	End With

 End Sub
 Set vpmShowDips=GetRef("EditDips")


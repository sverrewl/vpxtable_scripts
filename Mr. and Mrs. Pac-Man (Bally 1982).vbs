Option Explicit
Randomize

On Error Resume Next
ExecuteGlobal GetTextFile("controller.vbs")
If Err Then MsgBox "You need the controller.vbs in order to run this table, available in the vp10 package"
On Error Goto 0

' Thalamus 2019 February : Improved directional sounds
' !! NOTE : Table not verified yet !!
' Changed useSolenoids from 1 to 2.

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

Const cGameName="m_mpac",UseSolenoids=2,UseLamps=0,UseGI=0,SSolenoidOn="SolOn",SSolenoidOff="SolOff", SCoin="coin"

' Thalamus - added because of useSolenoids=2
Const cSingleRFlip = 0

LoadVPM "01500000","bally.vbs",3.1

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

SolCallback(1)=		"bsSaucer.SolOut"  'Left Kick
SolCallback(2)=		"bsSaucer.SolOutAlt" 'Right Kick
SolCallback(3)=		"bsSaucer2.SolOut"
SolCallback(6)= 	"vpmSolSound SoundFX(""Knocker"",DOFKnocker),"
SolCallback(7)=		"bsTrough.SolOut"
SolCallback(8)=		"dtL.SolUnHit 1,"
SolCallback(13)=	"dtL.SolDropUp"
SolCallback(14)=	"dtR.SolDropUp"
SolCallback(15)=	"dtC.SolDropUp"
SolCallback(17)=	"SolDiv"
SolCallback(19)=	"vpmNudge.SolGameOn"
SolCallback(33)=	"dtL.SolUnHit 2,"
SolCallback(34)=	"dtL.SolUnHit 3,"
SolCallback(35)=	"dtL.SolUnHit 4,"
SolCallback(36)=	"dtR.SolUnHit 1,"
SolCallback(37)=	"dtR.SolUnHit 2,"
SolCallback(38)=	"dtR.SolUnHit 3,"
SolCallback(39)=	"dtR.SolUnHit 4,"

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
         PlaySoundAtVol SoundFX("fx_Flipperup",DOFContactors), RightFlipper, VolFlip:RightFlipper.RotateToEnd:RightFlipper1.RotateToEnd
     Else
         PlaySoundAtVol SoundFX("fx_Flipperdown",DOFContactors), RightFlipper, VolFlip:RightFlipper.RotateToStart:RightFlipper1.RotateToStart
     End If
End Sub
'**********************************************************************************************************

'Solenoid Controlled toys
'**********************************************************************************************************

Sub SolDiv(Enabled)
     If Enabled Then
        PlaySoundAtVol SoundFX("fx_Flipperup",DOFContactors), LeftGatePrim, 1:
		LeftGate.RotateToEnd
		LeftGatePrim.rotY = 43
     Else
        PlaySoundAtVol SoundFX("fx_Flipperdown",DOFContactors), LeftGatePrim, 1:
		LeftGate.RotateToStart
		LeftGatePrim.rotY = 0
     End If
End Sub

'**********************************************************************************************************

'Initiate Table
'**********************************************************************************************************

Dim bsTrough, bsSaucer, bsSaucer2, dtL, dtC, dtR, mHole

Sub Table1_Init
	vpmInit Me
	On Error Resume Next
		With Controller
		.GameName = cGameName
		If Err Then MsgBox "Can't start Game" & cGameName & vbNewLine & Err.Description : Exit Sub
		.SplashInfoLine = "Mr and Ms Pac-Man"&chr(13)&"You Suck"
		.HandleMechanics=0
		.HandleKeyboard=0
		.ShowDMDOnly=1
		.ShowFrame=0
		.ShowTitle=0
        .hidden = 1
         On Error Resume Next
         .Run GetPlayerHWnd
         If Err Then MsgBox Err.Description
         On Error Goto 0
     End With
     On Error Goto 0

	PinMAMETimer.Interval=PinMAMEInterval
	PinMAMETimer.Enabled=1

	vpmNudge.TiltSwitch=15
	vpmNudge.Sensitivity=2
	vpmNudge.TiltObj=Array(Bumper1,Bumper2,LeftSlingshot,RightSlingshot)

	Set bsTrough=new cvpmBallStack
		bsTrough.InitSw 0,5,0,0,0,0,0,0
		bsTrough.InitKick BallRelease,90,5
		bsTrough.InitExitSnd SoundFX("ballrelease",DOFContactors), SoundFX("Solenoid",DOFContactors)
		bsTrough.Balls=1

	Set bsSaucer=New cvpmBallStack
		bsSaucer.InitSaucer sw7,7,45,10
		bsSaucer.InitAltKick 260,10
		bsSaucer.InitExitSnd SoundFX("Popper",DOFContactors), SoundFX("Solenoid",DOFContactors)

	Set bsSaucer2=New cvpmBallStack
		bsSaucer2.InitSaucer sw8,8,178,5
		bsSaucer2.InitExitSnd SoundFX("Popper",DOFContactors), SoundFX("Solenoid",DOFContactors)

	Set dtL=New cvpmDropTarget
		dtL.InitDrop Array(sw1,sw2,sw3,sw4),Array(1,2,3,4)
		dtL.InitSnd SoundFX("DTDrop",DOFContactors),SoundFX("DTReset",DOFContactors)

	Set dtC=New cvpmDropTarget
		dtC.InitDrop Array(sw22,sw23,sw24),Array(22,23,24)
		dtC.InitSnd SoundFX("DTDrop",DOFContactors),SoundFX("DTReset",DOFContactors)

	Set dtR=New cvpmDropTarget
		dtR.InitDrop Array(sw25,sw26,sw27,sw28),Array(25,26,27,28)
		dtR.InitSnd SoundFX("DTDrop",DOFContactors),SoundFX("DTReset",DOFContactors)

     Set mHole = New cvpmMagnet
         mHole.InitMagnet Umagnet, 5
         mHole.GrabCenter = 0
         mHole.MagnetOn = 1
         mHole.CreateEvents "mHole"

End Sub

'**********************************************************************************************************
'Plunger code
'**********************************************************************************************************

Sub Table1_KeyDown(ByVal KeyCode)
	If KeyDownHandler(keycode) Then Exit Sub
	If keycode = PlungerKey Then Plunger.Pullback:playsoundAtVol"plungerpull", Plunger, 1
	If KeyCode=LeftFlipperKey Then Controller.Switch(33)=1
	If KeyCode=RightFlipperKey Then Controller.Switch(37)=1
End Sub

Sub Table1_KeyUp(ByVal KeyCode)
	If KeyUpHandler(keycode) Then Exit Sub
	If keycode = PlungerKey Then Plunger.Fire:PlaySoundAtVol"plunger", Plunger, 1
	If KeyCode=LeftFlipperKey Then Controller.Switch(33)=0
	If KeyCode=RightFlipperKey Then Controller.Switch(37)=0
End Sub

'**********************************************************************************************************

 ' Drain hole and kickers
Sub Drain_Hit:bsTrough.addball me : playsoundAtVol"drain" , Drain, 1: End Sub
Sub sw7_Hit:bsSaucer.AddBall 0 : playsoundAtVol "popper_ball", ActiveBall, 1: End Sub
Sub sw8_Hit:bsSaucer2.AddBall 0 : playsoundAtVol "popper_ball", ActiveBall, 1: End Sub

'Drop Targets
Sub sw1_Dropped:dtL.Hit 1:End Sub
Sub sw2_Dropped:dtL.Hit 2:End Sub
Sub sw3_Dropped:dtL.Hit 3:End Sub
Sub sw4_Dropped:dtL.Hit 4:End Sub

Sub sw22_Dropped:dtC.Hit 1:End Sub
Sub sw23_Dropped:dtC.Hit 2:End Sub
Sub sw24_Dropped:dtC.Hit 3:End Sub

Sub sw25_Dropped:dtR.Hit 1:End Sub
Sub sw26_Dropped:dtR.Hit 2:End Sub
Sub sw27_Dropped:dtR.Hit 3:End Sub
Sub sw28_Dropped:dtR.Hit 4:End Sub

'Wire Triggers
Sub sw12_Hit:Controller.Switch(12)=1 : playsoundAtVol"rollover" , ActiveBall, 1: End Sub
Sub sw12_unHit:Controller.Switch(12)=0:End Sub
Sub sw12a_Hit:Controller.Switch(12)=1 : playsoundAtVol"rollover" , ActiveBall, 1: End Sub
Sub sw12a_unHit:Controller.Switch(12)=0:End Sub
Sub sw13_Hit:Controller.Switch(13)=1 : playsoundAtVol"rollover" , ActiveBall, 1: End Sub
Sub sw13_unHit:Controller.Switch(13)=0:End Sub
Sub sw14_Hit:Controller.Switch(14)=1 : playsoundAtVol"rollover" , ActiveBall, 1: End Sub
Sub sw14_unHit:Controller.Switch(14)=0:End Sub

'Bumpers
Sub Bumper1_Hit : vpmTimer.PulseSw(17) : playsoundAtVol SoundFX("fx_bumper1",DOFContactors), ActiveBall, 1: End Sub
Sub Bumper2_Hit : vpmTimer.PulseSw(18) : playsoundAtVol SoundFX("fx_bumper1",DOFContactors), ActiveBall, 1: End Sub

'Star Trigger
Sub sw21_Hit:Controller.Switch(21)=1 : playsoundAtVol"rollover" , ActiveBall, 1: End Sub
Sub sw21_unHit:Controller.Switch(21)=0:End Sub

'Scoring Rubbers
Sub sw29a_Hit:vpmTimer.PulseSw 29 : playsoundAtVol"flip_hit_3" , ActiveBall, 1: End Sub
Sub sw29b_Hit:vpmTimer.PulseSw 29 : playsoundAtVol"flip_hit_3" , ActiveBall, 1: End Sub
Sub sw29c_Hit:vpmTimer.PulseSw 29 : playsoundAtVol"flip_hit_3" , ActiveBall, 1: End Sub

'Gate switches
Sub sw30a_Hit:vpmTimer.PulseSw 30:End Sub
Sub sw30b_Hit:vpmTimer.PulseSw 30:End Sub

'Stand Up Targets
Sub Sw34_Hit:vpmTimer.PulseSw 34:End Sub
Sub Sw35_Hit:vpmTimer.PulseSw 35:End Sub
Sub Sw36_Hit:vpmTimer.PulseSw 36:End Sub

Sub Sw38_Hit:vpmTimer.PulseSw 38:End Sub
Sub Sw39_Hit:vpmTimer.PulseSw 39:End Sub
Sub Sw40_Hit:vpmTimer.PulseSw 40:End Sub


Set Motorcallback=Getref("UpdateSolenoids")
Sub UpdateSolenoids
	Dim Changed,Count,funcName,ii,sel,solNo
	Changed=Controller.ChangedSolenoids
	If Not IsEmpty(Changed) Then
		sel=Controller.Lamp(87)
		Count=UBound(Changed,1)
		For ii=0 To Count
			solNo=Changed(ii,CHGNO)
			If SolNo>=9 And SolNo<=15 And sel Then solNo=solNo+24 '9->33 etc
			funcName=SolCallback(solNo)
			If funcName<>"" Then Execute funcName&" CBool("&Changed(ii,CHGSTATE)&")"
		Next
	End If
End Sub

Set LampCallback=Getref("UpdateLamps")
Sub UpdateLamps
	Dim ii,Lamp1,Lamp2,Lamp3,stat1,stat2,chgLamp
	ChgLamp=Controller.ChangedLamps
	If IsEmpty(ChgLamp) Then Exit Sub
	On Error Resume Next
	For ii=0 To UBound(ChgLamp)
		Lights(ChgLamp(ii,CHGNO)).State=ChgLamp(ii,CHGSTATE)
	Next
	On Error Goto 0
	For ii=0 To UBound(ChgLamp)
		If IsObject(Matrix(ChgLamp(ii,CHGNO))) Then
			lamp1=chgLamp(ii,CHGNO) And 63:stat1=Controller.Lamp(lamp1)
			lamp2=lamp1+64:stat2=Controller.Lamp(lamp2)
			lamp3=lamp2+8
			If stat1 Then
				If stat2 Then
					Matrix(lamp2).State=0:Matrix(lamp1).State=0:Matrix(lamp3).State=1
				Else
					Matrix(lamp3).State=0:Matrix(lamp2).State=0:Matrix(lamp1).State=1
				End If
			Else
				If stat2 Then
					Matrix(lamp3).State=0:Matrix(lamp1).State=0:Matrix(lamp2).State=1
				Else
					Matrix(lamp3).State=0:Matrix(lamp2).State=0:Matrix(lamp1).State=1:Matrix(lamp1).State=0
				End If
			End If
		End If
	Next
	L8a.State=L8.State
	L9a.State=L9.State
	L10a.State=L10.State
	L25a.State=L25.State
	L41a.State=L41.State
	L42a.State=L42.State
	L57a.State=L57.State
	L71a.State=L71.State

End Sub


'-------------------------------------
' Map lights into array
'-------------------------------------
Dim Matrix(128)

' Matrix Yellow
Set Matrix(1)=L1
Set Matrix(2)=L2
Set Matrix(3)=L3
Set Matrix(4)=L4
Set Matrix(5)=L5
Set Matrix(6)=L6
Set Matrix(17)=L17
Set Matrix(18)=L18
Set Matrix(19)=L19
Set Matrix(20)=L20
Set Matrix(21)=L21
Set Matrix(22)=L22
Set Matrix(33)=L33
Set Matrix(34)=L34
Set Matrix(35)=L35
Set Matrix(36)=L36
Set Matrix(37)=L37
Set Matrix(38)=L38
Set Matrix(49)=L49
Set Matrix(50)=L50
Set Matrix(51)=L51
Set Matrix(52)=L52
Set Matrix(53)=L53
Set Matrix(54)=L54
Set Matrix(55)=L55

' Matrix Red
Set Matrix(65)=L65
Set Matrix(66)=L66
Set Matrix(67)=L67
Set Matrix(68)=L68
Set Matrix(69)=L69
Set Matrix(70)=L70
Set Matrix(81)=L81
Set Matrix(82)=L82
Set Matrix(83)=L83
Set Matrix(84)=L84
Set Matrix(85)=L85
Set Matrix(86)=L86
Set Matrix(97)=L97
Set Matrix(98)=L98
Set Matrix(99)=L99
Set Matrix(100)=L100
Set Matrix(101)=L101
Set Matrix(102)=L102
Set Matrix(113)=L113
Set Matrix(114)=L114
Set Matrix(115)=L115
Set Matrix(116)=L116
Set Matrix(117)=L117
Set Matrix(118)=L118
Set Matrix(119)=L119

' Matrix orange
Set Matrix(73)=L73
Set Matrix(74)=L74
Set Matrix(75)=L75
Set Matrix(76)=L76
Set Matrix(77)=L77
Set Matrix(78)=L78
Set Matrix(89)=L89
Set Matrix(90)=L90
Set Matrix(91)=L91
Set Matrix(92)=L92
Set Matrix(93)=L93
Set Matrix(94)=L94
Set Matrix(105)=L105
Set Matrix(106)=L106
Set Matrix(107)=L107
Set Matrix(108)=L108
Set Matrix(109)=L109
Set Matrix(110)=L110
Set Matrix(121)=L121
Set Matrix(122)=L122
Set Matrix(123)=L123
Set Matrix(124)=L124
Set Matrix(125)=L125
Set Matrix(126)=L126
Set Matrix(127)=L127

'PF lights
Set Lights(7)=L7
Set Lights(8)=L8
Set Lights(9)=L9
Set Lights(10)=L10
'Set Lights(11)=L11'Shoot Again
Set Lights(12)=L12
'Set Lights(13)=L13'Ball In Play
Set Lights(14)=L14
Set Lights(15)=L15
Set Lights(23)=L23
Set Lights(24)=L24
Set Lights(25)=L25
Set Lights(26)=L26
'Set Lights(27)=L27'Match
Set Lights(28)=L28
'Set Lights(29)=L29'High Score
Set Lights(30)=L30
Set Lights(31)=L31
Set Lights(39)=L39
Set Lights(40)=L40
Set Lights(41)=L41
Set Lights(42)=L42
Set Lights(43)=L43
Set Lights(44)=L44
'Set Lights(45)=L45'Game Over
Set Lights(46)=L46
Set Lights(47)=L47
Set Lights(56)=L56
Set Lights(57)=L57
Set Lights(58)=L58
Set Lights(59)=L59' Apron Credit
Set Lights(60)=L60
'Set Lights(61)=L61'Tilt
Set Lights(62)=L62
Set Lights(63)=L63
Set Lights(71)=L71
Set Lights(103)=L103

'**********************************************************************************************************
'Digital Display
'**********************************************************************************************************

Dim Digits(37)

'Player 1
Digits(0)=Array(D1,D2,D3,D4,D5,D6,D7,D267)
Digits(1)=Array(D8,D9,D10,D11,D12,D13,D14)
Digits(2)=Array(D15,D16,D17,D18,D19,D20,D21)
Digits(3)=Array(D22,D23,D24,D25,D26,D27,D28,D268)
Digits(4)=Array(D29,D30,D31,D32,D33,D34,D35)
Digits(5)=Array(D36,D37,D38,D39,D40,D41,D42)
Digits(6)=Array(D43,D44,D45,D46,D47,D48,D49)

'Player 2
Digits(7)=Array(D50,D51,D52,D53,D54,D55,D56,D269)
Digits(8)=Array(D57,D58,D59,D60,D61,D62,D63)
Digits(9)=Array(D64,D65,D66,D67,D68,D69,D70)
Digits(10)=Array(D71,D72,D73,D74,D75,D76,D77,D270)
Digits(11)=Array(D78,D79,D80,D81,D82,D83,D84)
Digits(12)=Array(D85,D86,D87,D88,D89,D90,D91)
Digits(13)=Array(D92,D93,D94,D95,D96,D97,D98)

'PLayer 3
Digits(14)=Array(D99,D100,D101,D102,D103,D104,D105,D271)
Digits(15)=Array(D106,D107,D108,D109,D110,D111,D112)
Digits(16)=Array(D113,D114,D115,D116,D117,D118,D119)
Digits(17)=Array(D120,D121,D122,D123,D124,D125,D126,D272)
Digits(18)=Array(D127,D128,D129,D130,D131,D132,D133)
Digits(19)=Array(D134,D135,D136,D137,D138,D139,D140)
Digits(20)=Array(D141,D142,D143,D144,D145,D146,D147)

'Player 4
Digits(21)=Array(D148,D149,D150,D151,D152,D153,D154,D273)
Digits(22)=Array(D155,D156,D157,D158,D159,D160,D161)
Digits(23)=Array(D162,D163,D164,D165,D166,D167,D168)
Digits(24)=Array(D169,D170,D171,D172,D173,D174,D175,D274)
Digits(25)=Array(D176,D177,D178,D179,D180,D181,D182)
Digits(26)=Array(D183,D184,D185,D186,D187,D188,D189)
Digits(27)=Array(D190,D191,D192,D193,D194,D195,D196)

'Ball in Play
Digits(28)=Array(D197,D198,D199,D200,D201,D202,D203)
Digits(29)=Array(D204,D205,D206,D207,D208,D209,D210)

'Credit
Digits(30)=Array(D211,D212,D213,D214,D215,D216,D217)
Digits(31)=Array(D218,D219,D220,D221,D222,D223,D224)

'PF Lights
Digits(32)=Array(D225,D226,D227,D228,D229,D230,D231)
Digits(33)=Array(D232,D233,D234,D235,D236,D237,D238)

Digits(34)=Array(D239,D240,D241,D242,D243,D244,D245)
Digits(35)=Array(D246,D247,D248,D249,D250,D251,D252)

Digits(36)=Array(D253,D254,D255,D256,D257,D258,D259)
Digits(37)=Array(D260,D261,D262,D263,D264,D265,D266)

Sub DisplayTimer_Timer
	Dim ChgLED,ii,num,chg,stat,obj
	ChgLED=Controller.ChangedLEDs(&Hffffffff,&Hffffffff)
	If Not IsEmpty(ChgLED) Then
		For ii=0 To UBound(chgLED)
			num=chgLED(ii,0):chg=chgLED(ii,1):stat=chgLED(ii,2)
			For Each obj In Digits(num)
				If chg And 1 Then obj.State=stat And 1
				chg=chg\2:stat=stat\2
			Next
		Next
	End If
End Sub

'if Full Screen turn off Backglas Components
If DesktopMode = True Then
dim xxx
For each xxx in BG:xxx.Visible = 1: Next
else
For each xxx in BG:xxx.Visible = 0: Next
End if

'**********************************************************************************************************

'Bally Mr & Mrs Pacman
'added by Inkochnito
Sub editDips
	Dim vpmDips : Set vpmDips = New cvpmDips
	With vpmDips
		.AddForm 700,400,"Mr & Mrs Pacman - DIP switches"
		.AddChk 7,10,180,Array("Match feature",&H08000000)'dip 28
		.AddChk 205,10,115,Array("Credits display",&H04000000)'dip 27
		.AddFrame 2,30,190,"Maximum credits",&H03000000,Array("10 credits",0,"15 credits",&H01000000,"25 credits",&H02000000,"40 credits",&H03000000)'dip 25&26
		.AddFrame 2,106,190,"Ball in Maze-saucer kickout",&H00000020,Array("kicks out before all moves are used",0,"does not kick till all moves are used",&H00000020)'dip 6
		.AddFrame 2,152,190,"Drop target 20,000 yellow arrow",&H00000040,Array("20,000 Arrow starts off",0,"20,000 Arrow starts on",&H00000040)'dip 7
		.AddFrame 2,198,190,"Game over attract",&H20000000,Array("Off",0,"On",&H20000000)'dip 30
		.AddFrame 2,248,190,"Pacman lites at start of the game",&H00002000,Array("3 Pacman lites",0,"4 Pacman lites",&H00002000)'dip 14
		.AddFrame 2,298,190,"Pacman aggressive lite",49152,Array ("goes out after all moves are used",0,"goes out when monster dies",&H00004000,"stays on entire ball",49152)'dip 15&16
		.AddFrame 205,30,190,"Balls per game *",&HC0000000,Array ("2 balls",&HC0000000,"3 balls",0,"4 balls",&H80000000,"5 balls",&H40000000)'dip 31&32
		.AddFrame 205,106,190,"Time to beat",&H00700000,Array("10",0,"15",&H00100000,"20",&H00200000,"25",&H00300000,"30",&H00400000,"35",&H00500000,"40",&H00600000,"45",&H00700000)'dip 21&22&23
		.AddFrame 205,248,190,"Extra ball lite on after",&H00800000,Array("going into saucer 10 times",0,"going into saucer 5 times",&H00800000)'dip 24
		.AddFrame 205,298,190,"Number of replays per game",&H10000000,Array("Only 1 replay per player per game",0,"All replays earned will be collected",&H10000000)'dip 29
		.AddLabel 30,370,350,20,"* Use Insert key to change left apron card to match."
		.AddLabel 30,390,350,20,"After hitting OK, press F3 to reset game with new settings."
		.ViewDips
	End With
End Sub
Set vpmShowDips = GetRef("editDips")


'**********************************************************************************************************


'**********************************************************************************************************

'**********Sling Shot Animations
' Rstep and Lstep  are the variables that increment the animation
'****************
Dim RStep, Lstep

Sub RightSlingShot_Slingshot
	vpmTimer.PulseSw 20
    PlaySoundAtVol SoundFX("right_slingshot",DOFContactors), sling1, 1
    RSling.Visible = 0
    RSling1.Visible = 1
    sling1.TransZ = -20
    RStep = 0
    RightSlingShot.TimerEnabled = 1
End Sub

Sub RightSlingShot_Timer
    Select Case RStep
        Case 3:RSLing1.Visible = 0:RSLing2.Visible = 1:sling1.TransZ = -10
        Case 4:RSLing2.Visible = 0:RSLing.Visible = 1:sling1.TransZ = 0:RightSlingShot.TimerEnabled = 0:
    End Select
    RStep = RStep + 1
End Sub

Sub LeftSlingShot_Slingshot
	vpmTimer.PulseSw 19
    PlaySoundAtVol SoundFX("left_slingshot",DOFContactors), sling2, 1
    LSling.Visible = 0
    LSling1.Visible = 1
    sling2.TransZ = -20
    LStep = 0
    LeftSlingShot.TimerEnabled = 1
End Sub

Sub LeftSlingShot_Timer
    Select Case LStep
        Case 3:LSLing1.Visible = 0:LSLing2.Visible = 1:sling2.TransZ = -10
        Case 4:LSLing2.Visible = 0:LSLing.Visible = 1:sling2.TransZ = 0:LeftSlingShot.TimerEnabled = 0:
    End Select
    LStep = LStep + 1
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
	FlipperRSh1.RotZ = RightFlipper1.currentangle
	Primsw30a.Rotx = sw30a.Currentangle +90
	Primsw30b.Rotx = sw30b.Currentangle +90

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
	PlaySoundAtVol "fx_spinner", Spinnner, VolSpin
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



option explicit
Randomize


Const cGameName="hypbl_l6"
Const UseSolenoids=1,UseLamps=True,UseGI=0,UseSyn=1,SSolenoidOn="SolOn",SSolenoidOff="Soloff"
Const SCoin="coin3",cCredits="Hyperball"



On Error Resume Next
ExecuteGlobal GetTextFile("controller.vbs")
If Err Then MsgBox "Can't open controller.vbs"
On Error Goto 0
LoadVPM "01320000","S7.VBS",3.1 



'***************************************************************
'*   				        Solenoids          	        	   *
'***************************************************************

SolCallback(21)="SolShootBall"'9 Ball Shooter/Ball Feed Motor
'SolCallback(25)="vpmNudge.SolGameOn"

'Flashers
SolCallback(9)="vpmFlasher Array(LightHyperball,LightHyperball1,LightHyperball2),"'1 HYPER Flashers
'SolCallback(11)="vpmFlasher Flasher2,"'3 PLAYER 1 Flashers Backbox
'SolCallback(16)="vpmFlasher Flasher3,"'8 PLAYER 2 Flashers Backbox

'General Illumination
SolCallback(12)= "vpmFlasher array(LightA,LightB,LightC,LightD,LightU,LightV,LightW,LightY,LightApronL,LightApronR,LightE,LightF,LightG,LightR,LightSS,LightT,TopLightCenter,TopLightLeft,TopLightRight,HIJLight,KLMLight,NOPLight),"'GI Playfield
'***************************************************************
'*   				      Table Init          	        	   *
'***************************************************************


Sub Table1_Init
	On Error Resume Next
	With Controller
		.GameName=cGameName
		If Err Then MsgBox "Can't start Game" & cGameName & vbNewLine & Err.Description : Exit Sub
		.SplashInfoLine=cCredits
		.HandleMechanics=0
		.ShowDMDOnly=1
		.ShowFrame=0
		.ShowTitle=0
		.Run
		.Hidden=1
		If Err Then MsgBox Err.Description
	End With
	On Error Goto 0
	PinMAMETimer.Interval=PinMAMEInterval  
	PinMAMETimer.Enabled=1  
	vpmNudge.TiltSwitch=1
	vpmNudge.Sensitivity=5
'	vpmNudge.TiltObj=Array(Wall1)
	vpmMapLights AllLights
End Sub 

'***************************************************************
'*   				         Cannon          	        	   *
'***************************************************************



Dim Angle,Position
Angle=0:Position=38
Sub SolShootBall(Enabled)
	If Enabled Then
     Ball5.transy = 15
     Ball6.transy = 15
     Ball7.transy = 15
     Ball8.transy = 15
     Ball9.transy = 15
     Ball10.transy = 15
     Ball11.transy = 15
     Ball12.transy = 15
     Ball13.transy = 15
     Ball14.transy = 15
     Ball15.transy = 15
     Ball16.transy = 15
     Ball17.transy = 15
     Ball18.transy = 15
     Ball19.transy = 15
     Ball20.transy = 15
     Ball21.transy = 15
     Ball22.transy = 15
     Ball23.transy = 15
     Ball24.transy = 15
     Ball25.transy = 15
     Ball26.transy = 15
     Ball27.transy = 8
     Ball27.transx = 15
     Ball28.transy = 8
     Ball28.transx = 15
     Ball29.transy = 8
     Ball29.transx = 15
     Ball30.transy = 8
     Ball30.transx = 15
     Ball31.transy = 8
     Ball31.transx = 15
     Ball32.transy = 8
     Ball32.transx = 15
     Ball33.transy = 8
     Ball33.transx = 15
     Ball34.transy = 8
     Ball34.transx = 15
     Ball35.transy = 8
     Ball35.transx = 15
     Ball36.transy = 8
     Ball36.transx = 15
     Ball37.transy = 8
     Ball37.transx = 15
     Ball38.transy = 15
     Ball39.transy = 15
     Ball40.transy = 15
     Ball41.transy = 15
     Ball42.transy = 15
     Ball43.transy = 15
     Ball44.transy = 15
     Ball45.transy = 15
Else
     Ball5.transy = 0
     Ball6.transy = 0
     Ball7.transy = 0
     Ball8.transy = 0
     Ball9.transy = 0
     Ball10.transy = 0
     Ball11.transy = 0
     Ball12.transy = 0
     Ball13.transy = 0
     Ball14.transy = 0
     Ball15.transy = 0
     Ball16.transy = 0
     Ball17.transy = 0
     Ball18.transy = 0
     Ball19.transy = 0
     Ball20.transy = 0
     Ball21.transy = 0
     Ball22.transy = 0
     Ball23.transy = 0
     Ball24.transy = 0
     Ball25.transy = 0
     Ball26.transy = 0
     Ball27.transy = 0
     Ball27.transx = 0
     Ball28.transy = 0
     Ball28.transx = 0
     Ball29.transy = 0
     Ball29.transx = 0
     Ball30.transy = 0
     Ball30.transx = 0
     Ball31.transy = 0
     Ball31.transx = 0
     Ball32.transy = 0
     Ball32.transx = 0
     Ball33.transy = 0
     Ball33.transx = 0
     Ball34.transy = 0
     Ball34.transx = 0
     Ball35.transy = 0
     Ball35.transx = 0
     Ball36.transy = 0
     Ball36.transx = 0
     Ball37.transy = 0
     Ball37.transx = 0
     Ball38.transy = 0
     Ball39.transy = 0
     Ball40.transy = 0
     Ball41.transy = 0
     Ball42.transy = 0
     Ball43.transy = 0
     Ball44.transy = 0
     Ball45.transy = 0
		TopRampKicker.CreateSizedBall(12)
       TopRampKicker.Kick 90,1
		Shooter.CreateSizedBall(12)
		Select Case Position
			Case 0:Shooter.Kick 322,50
			Case 1:Shooter.Kick 323,50
			Case 2:Shooter.Kick 324,50
			Case 3:Shooter.Kick 325,50
			Case 4:Shooter.Kick 326,50
			Case 5:Shooter.Kick 327,50
			Case 6:Shooter.Kick 328,50
			Case 7:Shooter.Kick 329,50
			Case 8:Shooter.Kick 330,50
			Case 9:Shooter.Kick 331,50
			Case 10:Shooter.Kick 332,50
			Case 11:Shooter.Kick 333,50
			Case 12:Shooter.Kick 334,50
			Case 13:Shooter.Kick 335,50
			Case 14:Shooter.Kick 336,50
			Case 15:Shooter.Kick 337,50
			Case 16:Shooter.Kick 338,50
			Case 17:Shooter.Kick 339,50
			Case 18:Shooter.Kick 340,50
			Case 19:Shooter.Kick 341,50
			Case 20:Shooter.Kick 342,50
			Case 21:Shooter.Kick 343,50
			Case 22:Shooter.Kick 344,50
			Case 23:Shooter.Kick 345,50
			Case 24:Shooter.Kick 346,50
			Case 25:Shooter.Kick 347,50
			Case 26:Shooter.Kick 348,50
			Case 27:Shooter.Kick 349,50
			Case 28:Shooter.Kick 350,50
			Case 29:Shooter.Kick 351,50
			Case 30:Shooter.Kick 352,50
			Case 31:Shooter.Kick 353,50
			Case 32:Shooter.Kick 354,50
			Case 33:Shooter.Kick 355,50
			Case 34:Shooter.Kick 356,50
			Case 35:Shooter.Kick 357,50
			Case 36:Shooter.Kick 358,50
			Case 37:Shooter.Kick 359,50
			Case 38:Shooter.Kick 0,50
			Case 39:Shooter.Kick 1,50
			Case 40:Shooter.Kick 2,50
			Case 41:Shooter.Kick 3,50
			Case 42:Shooter.Kick 4,50
			Case 43:Shooter.Kick 5,50
			Case 44:Shooter.Kick 6,50
			Case 45:Shooter.Kick 7,50
			Case 46:Shooter.Kick 8,50
			Case 47:Shooter.Kick 9,50
			Case 48:Shooter.Kick 10,50
			Case 49:Shooter.Kick 11,50
			Case 50:Shooter.Kick 12,50
			Case 51:Shooter.Kick 13,50
			Case 52:Shooter.Kick 14,50
			Case 53:Shooter.Kick 15,50
			Case 54:Shooter.Kick 16,50
			Case 55:Shooter.Kick 17,50
			Case 56:Shooter.Kick 18,50
			Case 57:Shooter.Kick 19,50
			Case 58:Shooter.Kick 20,50
			Case 59:Shooter.Kick 21,50
			Case 60:Shooter.Kick 22,50
			Case 61:Shooter.Kick 23,50
			Case 62:Shooter.Kick 24,50
			Case 63:Shooter.Kick 25,50
			Case 64:Shooter.Kick 26,50
			Case 65:Shooter.Kick 27,50
			Case 66:Shooter.Kick 28,50
			Case 67:Shooter.Kick 29,50
			Case 68:Shooter.Kick 30,50
			Case 69:Shooter.Kick 31,50
			Case 70:Shooter.Kick 32,50
			Case 71:Shooter.Kick 33,50
			Case 72:Shooter.Kick 34,50
			Case 73:Shooter.Kick 35,50
			Case 74:Shooter.Kick 36,50
			Case 75:Shooter.Kick 37,50
		End Select
	End If
End Sub

Dim LT,RT
LT=0:RT=0

Sub UpdatePos_Timer
	If LT=1 And RT=0 And Position>0 Then
		Position=Position-1
        Cannon.Rotz = (Position)
	End If
	If LT=0 And RT=1 And Position<75 Then
		Position=Position+1
        Cannon.Rotz = (Position)
	End If
End Sub

'***************************************************************
'*   				         Keys            	        	   *
'***************************************************************

ExtraKeyHelp=KeyName(KeyUpperLeft)&vbTab&"Left Fire"&vbNewLine&KeyName(KeyUpperRight)&vbTab&"Right Fire"
Sub Table1_KeyDown(ByVal KeyCode)
If KeyName(KeyCode)="K" Then
	If AATimer.Enabled Then
		AATimer.Enabled=0
	Else
		AATimer.Enabled=1
	End If
End If
If KeyCode=2 Then	'1-player start
	Controller.Switch(3)=1
	Exit Sub
End If
If KeyCode=3 Then	'2-player start
	Controller.Switch(2)=1
	Exit Sub
End If
	If KeyCode=LeftFlipperKey Then
		LT=1
		RT=0
		Exit Sub
	End If
	If KeyCode=RightFlipperKey Then
		LT=0
		RT=1
		Exit Sub
	End If
	If KeyCode=PlungerKey Then Controller.Switch(32)=1 'Z-Bomb Switch
    If KeyCode=LeftMagnaSave Then Controller.Switch(33)=1 'Left Shooter
    If KeyCode=RightMagnaSave Then Controller.Switch(34)=1 'Right Shooter
    If vpmKeyDown(KeyCode) Then Exit Sub 
    If keycode = 19 Then DisplayHyperBallRules
End Sub  

Sub AATimer_Timer
       vpmTimer.PulseSw 33
End Sub


Sub Table1_KeyUp(ByVal KeyCode) 

If KeyCode=2 Then
	Controller.Switch(3)=0
	Exit Sub
End If
If KeyCode=3 Then
	Controller.Switch(2)=0
	Exit Sub
End If
	If KeyCode=LeftFlipperKey Then
		LT=0
		Exit Sub
	End If
	If KeyCode=RightFlipperKey Then
		RT=0
		Exit Sub
	End If
	If KeyCode=PlungerKey Then Controller.Switch(32)=0 'Z-Bomb Switch
    If KeyCode=LeftMagnaSave Then Controller.Switch(33)=0 'Left Shooter
    If KeyCode=RightMagnaSave Then Controller.Switch(34)=0 'Right Shooter
    If vpmKeyUp(KeyCode) Then Exit Sub
End Sub 

'***************************************************************
'*   		   	      Flap Gate Primitives              	   *
'***************************************************************
Set MotorCallback = GetRef("RealTimeUpdates")

Sub RealTimeUpdates()
	 Flap21.Rotz = sw21gate.currentangle
     Flap22.Rotz = sw22gate.currentangle
     Flap23.Rotz = sw23gate.currentangle
     Flap24.Rotz = sw24gate.currentangle
     Flap25.Rotz = sw25gate.currentangle
     Flap26.Rotz = sw26gate.currentangle
     Flap27.Rotz = sw27gate.currentangle
     Flap28.Rotz = sw28gate.currentangle
     Flap29.Rotz = sw29gate.currentangle
    Cannon.Rotz = (Position)
End Sub

'***************************************************************
'*   				         Kickers           	        	   *
'***************************************************************

Sub Kicker9a_Hit:vpmTimer.PulseSw 9:Me.DestroyBall:End Sub
Sub Kicker10a_Hit:vpmTimer.PulseSw 10:Me.DestroyBall:End Sub
Sub Kicker11a_Hit:vpmTimer.PulseSw 11:Me.DestroyBall:End Sub
Sub Kicker12a_Hit:vpmTimer.PulseSw 12:Me.DestroyBall:End Sub
Sub Kicker17a_Hit:vpmTimer.PulseSw 17:Me.DestroyBall:End Sub
Sub Kicker18a_Hit:vpmTimer.PulseSw 18:Me.DestroyBall:End Sub
Sub Kicker19a_Hit:vpmTimer.PulseSw 19:Me.DestroyBall:End Sub
Sub Kicker13a_Hit:vpmTimer.PulseSw 13:Me.DestroyBall:End Sub
Sub Kicker14a_Hit:vpmTimer.PulseSw 14:Me.DestroyBall:End Sub
Sub Kicker15a_Hit:vpmTimer.PulseSw 15:Me.DestroyBall:End Sub
Sub Kicker16a_Hit:vpmTimer.PulseSw 16:Me.DestroyBall:End Sub
Sub Kicker29a_Hit:vpmTimer.PulseSw 29:Me.DestroyBall:End Sub
Sub Kicker30a_Hit:vpmTimer.PulseSw 30:Me.DestroyBall:End Sub
Sub Kicker31a_Hit:vpmTimer.PulseSw 31:Me.DestroyBall:End Sub
Sub Kicker20_Hit:vpmTimer.PulseSw 20:Me.DestroyBall:End Sub
Sub Kicker21_Hit:vpmTimer.PulseSw 21:Me.DestroyBall:End Sub
Sub Kicker22_Hit:vpmTimer.PulseSw 22:Me.DestroyBall:End Sub
Sub Kicker23_Hit:vpmTimer.PulseSw 23:Me.DestroyBall:End Sub
Sub Kicker24_Hit:vpmTimer.PulseSw 24:Me.DestroyBall:End Sub
Sub Kicker25_Hit:vpmTimer.PulseSw 25:Me.DestroyBall:End Sub
Sub Kicker26_Hit:vpmTimer.PulseSw 26:Me.DestroyBall:End Sub
Sub Kicker27_Hit:vpmTimer.PulseSw 27:Me.DestroyBall:End Sub
Sub Kicker28_Hit:vpmTimer.PulseSw 28:Me.DestroyBall:End Sub
Sub Drain_Hit:Me.DestroyBall:End Sub
Sub Drain2_Hit:Me.DestroyBall:End Sub
Sub Drain3_Hit:Me.DestroyBall:End Sub
Sub BottomRampKicker_Hit:Me.DestroyBall:End Sub
'***************************************************************
'*   				        Displays           	        	   *
'***************************************************************

Sub DisplayTimer_Timer
	Dim ChgLED,ii,jj,num,chg,stat,obj,b,x
	ChgLED = Controller.ChangedLEDs(0, &Hffffffff)
	If Not IsEmpty(ChgLED) Then
		For ii = 0 To UBound(chgLED)
			num = chgLED(ii, 0) : chg = chgLED(ii, 1) : stat = chgLED(ii, 2)
				For Each obj In Digits(num)
					If chg And 1 Then obj.State = stat And 1
					chg = chg\2 : stat = stat\2
				Next
		Next
	End If
End Sub 

Dim Digits(29)
Digits(0)=Array(BLight1,BLight2,BLight3,BLight4,BLight5,BLight6,BLight7,BLight8)
Digits(1)=Array(BLight10,BLight11,BLight12,BLight13,BLight14,BLight15,BLight16)
Digits(2)=Array(BLight17,BLight18,BLight19,BLight20,BLight21,BLight22,BLight23)
Digits(3)=Array(BLight24,BLight25,BLight26,BLight27,BLight28,BLight29,BLight30,BLight31)
Digits(4)=Array(BLight33,BLight34,BLight35,BLight36,BLight37,BLight38,BLight39)
Digits(5)=Array(BLight40,BLight41,BLight42,BLight43,BLight44,BLight45,BLight46)
Digits(6)=Array(BLight47,BLight48,BLight49,BLight50,BLight51,BLight52,BLight53)
Digits(7)=Array(BLight54,BLight55,BLight56,BLight57,BLight58,BLight59,BLight60,BLight61)
Digits(8)=Array(BLight63,BLight64,BLight65,BLight66,BLight67,BLight68,BLight69)
Digits(9)=Array(BLight70,BLight71,BLight72,BLight73,BLight74,BLight75,BLight76)
Digits(10)=Array(BLight77,BLight78,BLight79,BLight80,BLight81,BLight82,BLight83,Blight84)
Digits(11)=Array(BLight86,BLight87,BLight88,BLight89,BLight90,BLight91,BLight92)
Digits(12)=Array(BLight93,BLight94,BLight95,BLight96,BLight97,BLight98,BLight99)
Digits(13)=Array(BLight100,BLight101,BLight102,BLight103,BLight104,BLight105,BLight106)
Digits(14)=Array(BLight107,BLight108,BLight109,BLight110,BLight111,BLight112,BLight113)
Digits(15)=Array(BLight114,BLight115,BLight116,BLight117,BLight118,BLight119,BLight120)
Digits(16)=Array(BLight121,BLight122,BLight123,BLight124,BLight125,BLight126,BLight127)
Digits(17)=Array(BLight128,BLight129,BLight130,BLight131,BLight132,BLight133,BLight134)
Digits(18)=Array(Light10,Light11,Light12,Light13,Light14,Light15,Light16,Light17,Light18,Light19,Light20,Light21,Light22,Light23,Light24,Light25)
Digits(19)=Array(Light26,Light27,Light28,Light29,Light30,Light31,Light32,Light33,Light34,Light35,Light36,Light37,Light38,Light39,Light40,Light41)
Digits(20)=Array(Light42,Light43,Light44,Light45,Light46,Light47,Light48,Light49,Light50,Light51,Light52,Light53,Light54,Light55,Light56,Light57)
Digits(21)=Array(Light58,Light59,Light60,Light61,Light62,Light63,Light64,Light65,Light66,Light67,Light68,Light69,Light70,Light71,Light72,Light73)
Digits(22)=Array(Light74,Light75,Light76,Light77,Light78,Light79,Light80,Light81,Light82,Light83,Light84,Light85,Light86,Light87,Light88,Light89)
Digits(23)=Array(Light90,Light91,Light92,Light93,Light94,Light95,Light96,Light97,Light98,Light99,Light100,Light101,Light102,Light103,Light104,Light105)
Digits(24)=Array(Light106,Light107,Light108,Light109,Light110,Light111,Light112,Light113,Light114,Light115,Light116,Light117,Light118,Light119,Light120,Light121)
Digits(25)=Array(Light122,Light123,Light124,Light125,Light126,Light127,Light128,Light129,Light130,Light131,Light132,Light133,Light134,Light135,Light136,Light137)
Digits(26)=Array(Light138,Light139,Light140,Light141,Light142,Light143,Light144,Light145,Light146,Light147,Light148,Light149,Light150,Light151,Light152,Light153)
Digits(27)=Array(Light154,Light155,Light156,Light157,Light158,Light159,Light160,Light161,Light162,Light163,Light164,Light165,Light166,Light167,Light168,Light169)
Digits(28)=Array(Light170,Light171,Light172,Light173,Light174,Light175,Light176,Light177,Light178,Light179,Light180,Light181,Light182,Light183,Light184,Light185)
Digits(29)=Array(Light186,Light187,Light188,Light189,Light190,Light191,Light192,Light193,Light194,Light195,Light196,Light197,Light198,Light199,Light200,Light201)


'***************************************************************
'*   				       Desktop Mode       	        	   *
'***************************************************************
Dim xx
Dim DesktopMode:DesktopMode = Table1.ShowDT
If DesktopMode = True Then
		For each xx in DTVisible
			xx.visible = 1
		Next
End If
If DesktopMode = False Then
		For each xx in DTVisible
			xx.visible = 0
		Next
End If


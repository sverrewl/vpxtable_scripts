'Eight Ball Deluxe 1.0.1
'based on a script by 32assassin
'by Bord

' Thalamus 2018-07-24
' Table has its own "Positional Sound Playback Functions" and "Supporting Ball & Sound Functions"
' Table doens't need useSolenoid=2, it is included as its own function.

Option Explicit
Randomize

On Error Resume Next
ExecuteGlobal GetTextFile("controller.vbs")
If Err Then MsgBox "You need the controller.vbs in order to run this table, available in the vp10 package"
On Error Goto 0

Const cGameName="eballdlx",UseSolenoids=1,UseLamps=0,UseGI=0,SSolenoidOn="SolOn",SSolenoidOff="SolOff", SCoin="coin"

Dim enableBallControl

enableBallControl = 0 	' 1 to enable, 0 to disable

LoadVPM "01560000", "Bally.vbs", 3.26
Dim DesktopMode: DesktopMode = Table1.ShowDT

Dim BallShadows: Ballshadows=1  		'******************set to 1 to turn on Ball shadows
Dim FlipperShadows: FlipperShadows=1  '***********set to 1 to turn on Flipper shadows

If DesktopMode = True Then 'Show Desktop components
Ramp16.visible=1
Ramp15.visible=1
Else
Ramp16.visible=0
Ramp15.visible=0
End if

'*************************************************************
'Solenoid Call backs
'**********************************************************************************************************

 SolCallback(6) =  "vpmSolSound SoundFX(""Knocker"",DOFKnocker),"
 SolCallback(7) = "Soldtbank3"
 SolCallback(8) = "SolSaucer"
 SolCallback(9) = "SolBallRelease"
 SolCallback(10) = "SolReset"
 SolCallback(11) = "dtbank1.SolHit 3,"
 SolCallback(12) = "dtbank1.SolHit 4,"
 SolCallback(13) = "dtbank1.SolHit 5,"
 SolCallback(14) = "dtbank1.SolHit 6,"
 SolCallback(15) = "dtbank1.SolHit 7,"
 SolCallBack(19) = "FastFlips.TiltSol"
 SolCallback(sLRFlipper) = ""
 SolCallback(sLLFlipper) = ""
' SolCallback(sURFlipper) = ""
' SolCallback(sULFlipper) = ""

'SolCallback(sLRFlipper) = "SolRFlipper"
'SolCallback(sLLFlipper) = "SolLFlipper"

Sub SolLFlipper(Enabled)
     If Enabled Then
         PlaySound SoundFX("fx_Flipperup",DOFContactors):LeftFlipper.RotateToEnd:LeftFlipper1.RotateToEnd
     Else
         PlaySound SoundFX("fx_Flipperdown",DOFContactors):LeftFlipper.RotateToStart:LeftFlipper1.RotateToStart
     End If
  End Sub

Sub SolRFlipper(Enabled)
     If Enabled Then
         PlaySound SoundFX("fx_Flipperup",DOFContactors):RightFlipper.RotateToEnd
     Else
         PlaySound SoundFX("fx_Flipperdown",DOFContactors):RightFlipper.RotateToStart
     End If
End Sub

'**********************************************************************************************************

'Solenoid Controlled toys
'**********************************************************************************************************

 Sub SolSaucer(Enable)
	If Enable then
		If Controller.Lamp(52) then
			bsTP.SolOut Enable
			RKick = 0
			kickarmtop_prim.ObjRotX = -12
			RKickTimer.Enabled = 1
		Else
			dtbank1.SolDropUp True
			For each xx in DTRightLights: xx.state=0:Next
		End If
	End If
 End Sub

 Sub SolBallRelease(Enable)
     If Controller.Lamp(52) then
         bsTrough.SolOut Enable
     Else
         dtbank1.SolHit 1, Enable
	     'L2DG.state=1:L1DG.state=1
     End If
 End Sub

 Sub SolReset(Enable)
     If Controller.Lamp(52) then
         dtbank2.SolDropUp Enable
		 For each xx in DTLeftLights: xx.state=0:Next
     Else
         dtbank1.SolHit 2, Enable
		 'L2DF.state=1:L1DF.state=1
     End If
 End Sub

'*****GI Lights On
dim xx
For each xx in GI:xx.State = 1: Next

''*****Drop Lights Off
For each xx in DTLeftLights: xx.state=0:Next
For each xx in DTRightLights: xx.state=0:Next
For each xx in DT8Lights: xx.state=0:Next

	if ballshadows=1 then
		BallShadowUpdate.enabled=1
	  else
		BallShadowUpdate.enabled=0
	end if

	if flippershadows=1 then
		FlipperLSh.visible=1
		FlipperRSh.visible=1
	  else
		FlipperLSh.visible=0
		FlipperRSh.visible=0
	end if

'Primitive Flipper Code
Sub FlipperTimer_Timer
'	FlipperT1.roty = LeftFlipper.currentangle
'	FlipperT5.roty = RightFlipper.currentangle
	if FlipperShadows=1 then
		FlipperLsh.rotz= LeftFlipper.currentangle
		FlipperLsh1.rotz= LeftFlipper1.currentangle
		FlipperRsh.rotz= RightFlipper.currentangle
	end if
'	Pgate.rotz = Gate3.currentangle+25
End Sub

'*****************************************
'			BALL SHADOW
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
            BallShadow(b).X = ((BOT(b).X) - (Ballsize/16) + ((BOT(b).X - (Table1.Width/2))/17))' + 13
        Else
            BallShadow(b).X = ((BOT(b).X) + (Ballsize/16) + ((BOT(b).X - (Table1.Width/2))/17))' - 13
        End If
        ballShadow(b).Y = BOT(b).Y + 10
        If BOT(b).Z > 20 Then
            BallShadow(b).visible = 1
        Else
            BallShadow(b).visible = 0
        End If
    Next
End Sub


'**********************************************************************************************************

'Initiate Table
'**********************************************************************************************************

 Dim bsTrough, bsTP, dtBank1, dtBank2, dtBank3

 Sub Table1_Init
	vpmInit Me
	On Error Resume Next
		With Controller
		.GameName = cGameName
		If Err Then MsgBox "Can't start Game" & cGameName & vbNewLine & Err.Description : Exit Sub
		.SplashInfoLine = "Eight Ball Deluxe (Bally 1980)"&chr(13)&"v. 1.0"
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

     PinMAMETimer.Interval = PinMAMEInterval
     PinMAMETimer.Enabled = 1

     vpmNudge.TiltSwitch = swTilt
     vpmNudge.Sensitivity = 1
     vpmNudge.TiltObj = Array(Bumper1,Bumper2,Bumper3, LeftSlingshot, RightSlingshot)

     Set bsTrough = New cvpmBallStack
         bsTrough.InitSw 0, 8, 0, 0, 0, 0, 0, 0
         bsTrough.InitKick BallRelease, 80, 6
         bsTrough.InitExitSnd SoundFX("ballrelease",DOFContactors), SoundFX("Solenoid",DOFContactors)
         bsTrough.Balls = 1

     Set bsTP = New cvpmBallStack
         bsTP.InitSaucer sw34, 34, 315, 8
         bsTP.InitExitSnd SoundFX("Popper",DOFContactors), SoundFX("Solenoid",DOFContactors)
         bsTP.KickAngleVar = 3
         bsTP.KickForceVar = 3

     set dtbank1 = new cvpmdroptarget
         dtbank1.initdrop array(sw17, sw18, sw19, sw20, sw21, sw22, sw23), array(17, 18, 19, 20, 21, 22, 23)
         dtbank1.initsnd  SoundFX("DTDrop",DOFContactors),SoundFX("DTReset",DOFContactors)

     set dtbank2 = new cvpmdroptarget
         dtbank2.initdrop array(sw1, sw2, sw3, sw4), array(1, 2, 3, 4)
         dtbank2.initsnd  SoundFX("DTDrop",DOFContactors),SoundFX("DTReset",DOFContactors)

     set dtbank3 = new cvpmdroptarget
         dtbank3.initdrop sw33, 33
         dtbank3.initsnd  SoundFX("DTDrop",DOFContactors),SoundFX("DTReset",DOFContactors)

 End Sub

'**********************************************************************************************************
'Plunger code
'**********************************************************************************************************

Sub Table1_KeyDown(ByVal KeyCode)
	If KeyDownHandler(keycode) Then Exit Sub
	If keycode = PlungerKey Then Plunger.Pullback:playsound"plungerpull"
	If KeyCode = LeftFlipperKey then FastFlips.FlipL True :  'FastFlips.FlipUL True
	If KeyCode = RightFlipperKey then FastFlips.FlipR True :  'FastFlips.FlipUR True

'************************   Start Ball Control 1/3
	if enableBallControl then
		if keycode = 46 then	 			' C Key
			If contball = 1 Then
				contball = 0
			Else
				contball = 1
			End If
		End If
		if keycode = 48 then 				'B Key
			If bcboost = 1 Then
				bcboost = bcboostmulti
			Else
				bcboost = 1
			End If
		End If
	End If
	if keycode = 203 then bcleft = 1		' Left Arrow
	if keycode = 200 then bcup = 1			' Up Arrow
	if keycode = 208 then bcdown = 1		' Down Arrow
	if keycode = 205 then bcright = 1		' Right Arrow

'************************   End Ball Control 1/3

End Sub

Sub Table1_KeyUp(ByVal KeyCode)
	If KeyUpHandler(keycode) Then Exit Sub
	If keycode = PlungerKey Then Plunger.Fire:PlaySound"plunger"
	If KeyCode = LeftFlipperKey then FastFlips.FlipL False :  'FastFlips.FlipUL False
	If KeyCode = RightFlipperKey then FastFlips.FlipR False : ' FastFlips.FlipUR False

'************************   Start Ball Control 2/3
	if keycode = 203 then bcleft = 0		' Left Arrow
	if keycode = 200 then bcup = 0			' Up Arrow
	if keycode = 208 then bcdown = 0		' Down Arrow
	if keycode = 205 then bcright = 0		' Right Arrow
'************************   End Ball Control 2/3

End Sub

'************************   Start Ball Control 3/3
Sub StartControl_Hit()
	Set ControlBall = ActiveBall
	contballinplay = true
End Sub

Sub StopControl_Hit()
	contballinplay = false
End Sub

Dim bcup, bcdown, bcleft, bcright, contball, contballinplay, ControlBall, bcboost
Dim bcvel, bcyveloffset, bcboostmulti

bcboost = 1		'Do Not Change - default setting
bcvel = 4		'Controls the speed of the ball movement
bcyveloffset = -0.01 	'Offsets the force of gravity to keep the ball from drifting vertically on the table, should be negative
bcboostmulti = 3	'Boost multiplier to ball veloctiy (toggled with the B key)

Sub BallControl_Timer()
	If Contball and ContBallInPlay then
		If bcright = 1 Then
			ControlBall.velx = bcvel*bcboost
		ElseIf bcleft = 1 Then
			ControlBall.velx = - bcvel*bcboost
		Else
			ControlBall.velx=0
		End If

		If bcup = 1 Then
			ControlBall.vely = -bcvel*bcboost
		ElseIf bcdown = 1 Then
			ControlBall.vely = bcvel*bcboost
		Else
			ControlBall.vely= bcyveloffset
		End If
	End If
End Sub
'************************   End Ball Control 3/3


'**********************************************************************************************************

 ' Drain hole and kickers
Sub Drain_Hit:bsTrough.addball me : playsound"drain" : End Sub
Sub sw34_Hit:bsTP.AddBall 0 : playsound"popper_ball" : End Sub

 ' Droptargets
 Sub sw1_Dropped:dtbank2.Hit 1
	L1DA.state=1
End Sub

 Sub sw2_Dropped:dtbank2.Hit 2
	L1DB.state=1
     If Controller.Lamp(101) then
         L101DB.state=1
     Else
         L101DB.state=0
     End If
     If Controller.Lamp(117) then
         L117DB.state=1
     Else
         L117DB.state=0
     End If
End Sub

 Sub sw3_Dropped:dtbank2.Hit 3
	L1DC.state=1:L2DC.state=1
     If Controller.Lamp(101) then
         L101DC.state=1
     Else
         L101DC.state=0
     End If
     If Controller.Lamp(117) then
         L117DC.state=1
     Else
         L117DC.state=0
     End If
End Sub

 Sub sw4_Dropped:dtbank2.Hit 4
	L2DD.state=1
     If Controller.Lamp(117) then
         L117DD.state=1
     Else
         L117DD.state=0
     End If
     If Controller.Lamp(118) then
         L118DD.state=1
     Else
         L118DD.state=0
     End If
End Sub

 Sub sw17_Dropped:dtbank1.Hit 1
	L2DG.state=1:L1DG.state=1
End Sub

 Sub sw18_Dropped:dtbank1.Hit 2
	L2DF.state=1:L1DF.state=1
End Sub

 Sub sw19_Dropped:dtbank1.Hit 3
	L2DE.state=1:L1DE.state=1
End Sub

 Sub sw20_Dropped:dtbank1.Hit 4
	L2DD.state=1
End Sub

 Sub sw21_Dropped:dtbank1.Hit 5
	L1DD1.state=1:L1DD2.state=1
End Sub

 Sub sw22_Dropped:dtbank1.Hit 6:End Sub
 Sub sw23_Dropped:dtbank1.Hit 7:End Sub
 Sub sw33_Dropped:dtbank3.Hit 1:L1D8.state=1:L2D8.state=1:End Sub

Sub Soldtbank3(enabled)
	dim xx
	if enabled then
		dtBank3.SolDropUp enabled
		For each xx in DT8Lights: xx.state=0:Next
	end if
End Sub

'Rollovers
 Sub sw12_Hit:Controller.Switch(12) = 1 : playsound"rollover" : End Sub
 Sub sw12_UnHit:Controller.Switch(12) = 0:End Sub
 Sub sw13_Hit:Controller.Switch(13) = 1 : playsound"rollover" : End Sub
 Sub sw13_UnHit:Controller.Switch(13) = 0:End Sub
 Sub sw14_Hit:Controller.Switch(14) = 1 : playsound"rollover" : End Sub
 Sub sw14_UnHit:Controller.Switch(14) = 0:End Sub
 Sub sw15_Hit:Controller.Switch(15) = 1 : playsound"rollover" : End Sub
 Sub sw15_UnHit:Controller.Switch(15) = 0:End Sub
 Sub sw31_Hit:Controller.Switch(31) = 1 : playsound"rollover" : End Sub
 Sub sw31_UnHit:Controller.Switch(31) = 0:End Sub
 Sub sw32_Hit:Controller.Switch(32) = 1 : playsound"rollover" : End Sub
 Sub sw32_UnHit:Controller.Switch(32) = 0:End Sub

'Bumpers
Sub Bumper1_Hit : vpmTimer.PulseSw(38) : playsound SoundFX("fx_bumper1",DOFContactors): End Sub
Sub Bumper2_Hit : vpmTimer.PulseSw(39) : playsound SoundFX("fx_bumper1",DOFContactors): End Sub
Sub Bumper3_Hit : vpmTimer.PulseSw(40) : playsound SoundFX("fx_bumper1",DOFContactors): End Sub

'Stand Up Targets
Sub sw5_Hit:vpmTimer.PulseSw 5:End Sub
Sub sw25_Hit:vpmTimer.PulseSw 25:End Sub
Sub sw26_Hit:vpmTimer.PulseSw 26:End Sub
Sub sw27_Hit:vpmTimer.PulseSw 27:End Sub
Sub sw28_Hit:vpmTimer.PulseSw 28:End Sub
Sub sw29_Hit:vpmTimer.PulseSw 29:End Sub
Sub sw30_Hit:vpmTimer.PulseSw 30:End Sub

'Star Trigger
 Sub sw35_Hit:Controller.Switch(35) = 1 : playsound"rollover" : End Sub
 Sub sw35_UnHit:Controller.Switch(35) = 0:End Sub

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

 Sub UpdateLamps()
     NFadeLm 1, l1
     NFadeL 1, l1b
     NFadeLm 2, l2
     NFadeL 2, l2b
     NFadeLm 3, l3
     NFadeL 3, l3b
     NFadeLm 4, l4
     NFadeL 4, l4b
     NFadeLm 5, l5
     NFadeL 5, l5b
     NFadeLm 6, l6
     NFadeL 6, l6b
     NFadeLm 7, l7
     NFadeL 7, l7b
     NFadeLm 8, l8
     NFadeL 8, l8b
     NFadeLm 9, l9
     NFadeL 9, l9b
     NFadeLm 10, l10
     NFadeL 10, l10b
     NFadeLm 11, l11
     NFadeL 11, l11b
     NFadeLm 12, l12
     NFadeL 12, l12b
     NFadeLm 17, l17
     NFadeL 17, l17b
     NFadeLm 18, l18
     NFadeL 18, l18b
     NFadeLm 19, l19
     NFadeL 19, l19b
     NFadeLm 20, l20
     NFadeL 20, l20b
     NFadeLm 21, l21
     NFadeL 21, l21b
     NFadeLm 22, l22
     NFadeL 22, l22b
     NFadeLm 23, l23
     NFadeL 23, l23b
     NFadeLm 24, l24
     NFadeL 24, l24b
     NFadeLm 25, l25
     NFadeL 25, l25b
     NFadeLm 26, l26
     NFadeL 26, l26b
     NFadeLm 28, l28
     NFadeL 28, l28b
     NFadeLm 33, l33
     NFadeL 33, l33b
     NFadeLm 34, l34
     NFadeL 34, l34b
     NFadeLm 35, l35
     NFadeL 35, l35b
     NFadeLm 36, l36
     NFadeL 36, l36b
     NFadeLm 37, l37
     NFadeL 37, l37b
     NFadeLm 38, l38
     NFadeL 38, l38b
     NFadeLm 39, l39
     NFadeL 39, l39b
     NFadeLm 40, l40
     NFadeL 40, l40b
     NFadeLm 41, l41
     NFadeL 41, l41b
     NFadeLm 42, l42
     NFadeL 42, l42b
     NFadeLm 44, l44
     NFadeL 44, l44b
     NFadeLm 47, l47
     NFadeL 47, l47b
     NFadeLm 49, l49
     NFadeL 49, l49b
     NFadeLm 50, l50
     NFadeL 50, l50b
     NFadeLm 51, l51
     NFadeL 51, l51b
     NFadeLm 53, l53
     NFadeL 53, l53b
     NFadeLm 54, l54
     NFadeL 54, l54b
     NFadeLm 55, l55
     NFadeL 55, l55b
     NFadeLm 57, l57
     NFadeL 57, l57b
     NFadeLm 58, l58
     NFadeL 58, l58b
	 NFadeL 59, l59 'Apron credit
     NFadeLm 60, l60
     NFadeL 60, l60b
     NFadeLm 63, l63
     NFadeL 63, l63b
     NFadeLm 65, l65
     NFadeL 65, l65b
     NFadeLm 66, l66
     NFadeL 66, l66b
     NFadeLm 67, l67
     NFadeL 67, l67b
     NFadeLm 68, l68
     NFadeL 68, l68b
     NFadeLm 69, l69
     NFadeL 69, l69b
	 NFadeLm 70, L70 'Bumper 1
	 NFadeL 70, l70b
     NFadeLm 71, l71
     NFadeL 71, l71b
     NFadeLm 81, l81
     NFadeL 81, l81b
     NFadeLm 82, l82
     NFadeL 82, l82b
     NFadeLm 83, l83
     NFadeL 83, l83b
     NFadeLm 84, l84
     NFadeL 84, l84b
     NFadeLm 85, l85
	 NFadeLm 86, l86 'Bumper 2
	 NFadeL 86, l86b
     NFadeLm 87, l87
     NFadeL 87, l87b
     NFadeLm 97, l97
     NFadeL 97, l97b
     NFadeLm 98, l98
     NFadeL 98, l98b
     NFadeLm 99, l99
     NFadeL 99, l99b
     NFadeLm 100, l100
     NFadeL 100, l100b
     NFadeLm 102, l102 'Bumper3
	 NFadeL 102, l102b
     NFadeLm 103, l103
     NFadeL 103, l103b
     NFadeLm 113, l113
     NFadeL 113, l113b
     NFadeLm 114, l114
     NFadeL 114, l114b
     NFadeLm 115, l115
     NFadeL 115, l115b
     NFadeLm 116, l116
     NFadeL 116, l116b

     NFadeLm 101, L101
     NFadeLm 101, L101a
     NFadeLm 101, L101b
     If L1DA.state=1 then
		NFadeLm 101, L101DA
	 Else
		L101DA.state=0
	End If
     If L1DB.state=1 then
		NFadeLm 101, L101DB
	 Else
		L101DB.state=0
	End If
     If L1DC.state=1 then
		NFadeLm 101, L101DC
	 Else
		L101DC.state=0
	End If
     NFadeL 101, L101c

     NFadeLm 117, L117
     NFadeLm 117, L117a
     NFadeLm 117, L117b
     If L1DB.state=1 then
		NFadeLm 117, L117DB
	 Else
		L117DB.state=0
	End If
    If L1DC.state=1 then
		NFadeLm 117, L117DC
	 Else
		L117DC.state=0
	End If
    If L2DD.state=1 then
		NFadeLm 117, L117DD
	 Else
		L117DD.state=0
	End If
    If L1D8.state=1 then
		NFadeLm 117, L117D8
	 Else
		L117D8.state=0
	End If
     NFadeL 117, L117c

     NFadeLm 118, L118
     NFadeLm 118, L118a
     NFadeLm 118, L118b
     If L2DD.state=1 then
		NFadeLm 118, L118DD
	 Else
		L118DD.state=0
	End If
    If L1D8.state=1 then
		NFadeLm 118, L118D8
	 Else
		L118D8.state=0
	End If
     NFadeL 118, L118c

     NFadeLm 119, L119
     NFadeLm 119, L119a
     NFadeLm 119, L119b
     NFadeL 119, L119c
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

 'Reels
Sub FadeReel(nr, reel)
    Select Case FadingLevel(nr)
        Case 2:FadingLevel(nr) = 0
        Case 3:FadingLevel(nr) = 2
        Case 4:reel.Visible = 0:FadingLevel(nr) = 3
        Case 5:reel.Visible = 1:FadingLevel(nr) = 1
    End Select
End Sub

 'Inverted Reels
Sub FadeIReel(nr, reel)
    Select Case FadingLevel(nr)
        Case 2:FadingLevel(nr) = 0
        Case 3:FadingLevel(nr) = 2
        Case 4:reel.Visible = 1:FadingLevel(nr) = 3
        Case 5:reel.Visible = 0:FadingLevel(nr) = 1
    End Select
End Sub

'**********************************************************************************************************
'Digital Display
'**********************************************************************************************************

Dim Digits(32)
' 1st Player
Digits(0) = Array(LED10,LED11,LED12,LED13,LED14,LED15,LED16)
Digits(1) = Array(LED20,LED21,LED22,LED23,LED24,LED25,LED26)
Digits(2) = Array(LED30,LED31,LED32,LED33,LED34,LED35,LED36)
Digits(3) = Array(LED40,LED41,LED42,LED43,LED44,LED45,LED46)
Digits(4) = Array(LED50,LED51,LED52,LED53,LED54,LED55,LED56)
Digits(5) = Array(LED60,LED61,LED62,LED63,LED64,LED65,LED66)
Digits(6) = Array(LED70,LED71,LED72,LED73,LED74,LED75,LED76)

' 2nd Player
Digits(7) = Array(LED80,LED81,LED82,LED83,LED84,LED85,LED86)
Digits(8) = Array(LED90,LED91,LED92,LED93,LED94,LED95,LED96)
Digits(9) = Array(LED100,LED101,LED102,LED103,LED104,LED105,LED106)
Digits(10) = Array(LED110,LED111,LED112,LED113,LED114,LED115,LED116)
Digits(11) = Array(LED120,LED121,LED122,LED123,LED124,LED125,LED126)
Digits(12) = Array(LED130,LED131,LED132,LED133,LED134,LED135,LED136)
Digits(13) = Array(LED140,LED141,LED142,LED143,LED144,LED145,LED146)

' 3rd Player
Digits(14) = Array(LED150,LED151,LED152,LED153,LED154,LED155,LED156)
Digits(15) = Array(LED160,LED161,LED162,LED163,LED164,LED165,LED166)
Digits(16) = Array(LED170,LED171,LED172,LED173,LED174,LED175,LED176)
Digits(17) = Array(LED180,LED181,LED182,LED183,LED184,LED185,LED186)
Digits(18) = Array(LED190,LED191,LED192,LED193,LED194,LED195,LED196)
Digits(19) = Array(LED200,LED201,LED202,LED203,LED204,LED205,LED206)
Digits(20) = Array(LED210,LED211,LED212,LED213,LED214,LED215,LED216)

' 4th Player
Digits(21) = Array(LED220,LED221,LED222,LED223,LED224,LED225,LED226)
Digits(22) = Array(LED230,LED231,LED232,LED233,LED234,LED235,LED236)
Digits(23) = Array(LED240,LED241,LED242,LED243,LED244,LED245,LED246)
Digits(24) = Array(LED250,LED251,LED252,LED253,LED254,LED255,LED256)
Digits(25) = Array(LED260,LED261,LED262,LED263,LED264,LED265,LED266)
Digits(26) = Array(LED270,LED271,LED272,LED273,LED274,LED275,LED276)
Digits(27) = Array(LED280,LED281,LED282,LED283,LED284,LED285,LED286)

' Credits
Digits(28) = Array(LED4,LED2,LED6,LED7,LED5,LED1,LED3)
Digits(29) = Array(LED18,LED9,LED27,LED28,LED19,LED8,LED17)
' Balls
Digits(30) = Array(LED39,LED37,LED48,LED49,LED47,LED29,LED38)
Digits(31) = Array(LED67,LED58,LED69,LED77,LED68,LED57,LED59)

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

'Bally EightBallDeluxe
'corrected by Inkochnito

Sub editDips
	Dim vpmDips : Set vpmDips = New cvpmDips
	With vpmDips
		.AddForm 315,450,"Eight Ball Deluxe - DIP switches"
		.AddFrame 2,10,140,"Balls per game",&HC0000000,Array("2",&HC0000000,"3",0,"4",&H80000000,"5",&H40000000)'dip 31&32
		.AddFrame 2,90,140,"C && D lanes are",&H00000040,Array("separate",0,"linked",&H00000040)'dip 7
		.AddFrame 2,140,140,"Saucer collects bonus",32768,Array("without multiplier",0,"with multiplier",32768)'dip 16
		.AddFrame 2,190,140,"Replays awarded per game",&H10000000,Array("1 per player per game",0,"unlimited replays",&H10000000)'dip 29
		.AddFrame 2,240,140,"Backbox Deluxe advances",&H00200000,Array("when special is lit",0,"every time",&H00200000)'dip 22
		.AddFrame 2,290,298,"Left lane score sequence",&H00002000,Array("no lite, 10K, 30K, 50K, X-ball, 70K, special, 70K stays on",&H00002000,"1st ball - no special lite, next ball - no extra ball lite",0)'dip 14
		.AddChk 7,345,190,Array("Drop targets not reset after Deluxe",&H00400000)'dip 23
		.AddChk 7,365,160,Array("Playfield Deluxe memory",&H00100000)'dip 21
		.AddChk 7,385,160,Array("Voice on during attact mode",&H20000000)'dip 30
		.AddFrame 160,10,140,"Maximum credits",&H03000000,Array("10",0,"15",&H01000000,"25",&H02000000,"40",&H03000000)'dip 25&26
		.AddFrame 160,90,140,"Making A-B-C-D spots",&H00000080,Array("1 target",0,"2 targets",&H00000080)'dip 8
		.AddFrame 160,140,140,"Target && Deluxe specials",&H00000020,Array("alternates with 50K",0,"scores only once per ball",&H00000020)'dip 6
		.AddFrame 160,190,140,"Eight Ball scores special on",&H00004000,Array("third time",0,"second and third time",&H00004000)'dip 15
		.AddFrame 160,240,140,"Deluxe special scored",&H00800000,Array("before 50K",0,"after 50K",&H00800000)'dip 24
		.AddChk 210,345,90,Array("Credit display",&H04000000)'dip 27
		.AddChk 210,365,90,Array("Match",&H08000000)'dip 28
		.AddLabel 7,410,308,20,"Set selftest position 16,17,18 and 19 to 03 for the best gameplay."
		.AddLabel 7,430,300,20,"After hitting OK, press F3 to reset game with new settings."
		.ViewDips
	End With
End Sub
Set vpmShowDips = GetRef("editDips")



'**********Sling Shot Animations
' Rstep and Lstep  are the variables that increment the animation
'****************
Dim RStep, Lstep, LA, LB, RA, RB, RC, RD, RKick

Sub RightSlingShot_Slingshot
	vpmTimer.PulseSw 36
    PlaySound SoundFX("right_slingshot",DOFContactors), 0, 1, 0.05, 0.05
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
	vpmTimer.PulseSw 37
    PlaySound SoundFX("left_slingshot",DOFContactors),0,1,-0.05,0.05
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

Sub WallLA_Hit
    RubberLA.Visible = 0
    RubberLA1.Visible = 1
    LA = 0
    WallLA.TimerEnabled = 1
End Sub

Sub WallLA_Timer
    Select Case LA
        Case 3:RubberLA1.Visible = 0:RubberLA2.Visible = 1
        Case 4:RubberLA2.Visible = 0:RubberLA.Visible = 1:WallLA.TimerEnabled = 0:
    End Select
    LA = LA + 1
End Sub

Sub WallLB_Hit
    RubberLB.Visible = 0
    RubberLB1.Visible = 1
    LB = 0
    WallLB.TimerEnabled = 1
End Sub

Sub WallLB_Timer
    Select Case LB
        Case 3:RubberLB1.Visible = 0:RubberLB2.Visible = 1
        Case 4:RubberLB2.Visible = 0:RubberLB.Visible = 1:WallLB.TimerEnabled = 0:
    End Select
    LB = LB + 1
End Sub

Sub WallRA_Hit
    RubberRA.Visible = 0
    RubberRA1.Visible = 1
    RA = 0
    WallRA.TimerEnabled = 1
End Sub

Sub WallRA_Timer
    Select Case RA
        Case 3:RubberRA1.Visible = 0:RubberRA2.Visible = 1
        Case 4:RubberRA2.Visible = 0:RubberRA.Visible = 1:WallRA.TimerEnabled = 0:
    End Select
    RA = RA + 1
End Sub

Sub WallRB_Hit
    RubberRB.Visible = 0
    RubberRB1.Visible = 1
    RB = 0
    WallRB.TimerEnabled = 1
End Sub

Sub WallRB_Timer
    Select Case RB
        Case 3:RubberRB1.Visible = 0:RubberRB2.Visible = 1
        Case 4:RubberRB2.Visible = 0:RubberRB.Visible = 1:WallRB.TimerEnabled = 0:
    End Select
    RB = RB + 1
End Sub

Sub sw24a_Slingshot
	vpmTimer.PulseSw 24
    RubberRC.Visible = 0
    RubberRC1.Visible = 1
    RC = 0
    sw24a.TimerEnabled = 1
End Sub

Sub sw24a_Timer
    Select Case RC
        Case 3:RubberRC1.Visible = 0:RubberRC2.Visible = 1
        Case 4:RubberRC2.Visible = 0:RubberRC.Visible = 1:sw24a.TimerEnabled = 0:
    End Select
    RC = RC + 1
End Sub

Sub sw24_Slingshot
	vpmTimer.PulseSw 24
    RubberRD.Visible = 0
    RubberRD1.Visible = 1
    RD = 0
    sw24.TimerEnabled = 1
End Sub

Sub sw24_Timer
    Select Case RD
        Case 3:RubberRD1.Visible = 0:RubberRD2.Visible = 1
        Case 4:RubberRD2.Visible = 0:RubberRD.Visible = 1:sw24.TimerEnabled = 0:
    End Select
    RD = RD + 1
End Sub


Sub RKickTimer_Timer
    Select Case RKick
        Case 1:kickarmtop_prim.ObjRotX = -50
        Case 2:kickarmtop_prim.ObjRotX = -50
        Case 3:kickarmtop_prim.ObjRotX = -50
        Case 4:kickarmtop_prim.ObjRotX = -50
        Case 5:kickarmtop_prim.ObjRotX = -50
        Case 6:kickarmtop_prim.ObjRotX = -50
        Case 7:kickarmtop_prim.ObjRotX = -50
        Case 8:kickarmtop_prim.ObjRotX = -50
        Case 9:kickarmtop_prim.ObjRotX = -50
        Case 10:kickarmtop_prim.ObjRotX = -50
        Case 11:kickarmtop_prim.ObjRotX = -24
        Case 12:kickarmtop_prim.ObjRotX = -12
        Case 13:kickarmtop_prim.ObjRotX = 0:RKickTimer.Enabled = 0
    End Select
    RKick = RKick + 1
End Sub

'*********************************************************************
'                 Positional Sound Playback Functions
'*********************************************************************

' Play a sound, depending on the X,Y position of the table element (especially cool for surround speaker setups, otherwise stereo panning only)
' parameters (defaults): loopcount (1), volume (1), randompitch (0), pitch (0), useexisting (0), restart (1))
' Note that this will not work (currently) for walls/slingshots as these do not feature a simple, single X,Y position
Sub PlayXYSound(soundname, tableobj, loopcount, volume, randompitch, pitch, useexisting, restart)
	PlaySound soundname, loopcount, volume, AudioPan(tableobj), randompitch, pitch, useexisting, restart, AudioFade(tableobj)
End Sub

' Similar subroutines that are less complicated to use (e.g. simply use standard parameters for the PlaySound call)
Sub PlaySoundAt(soundname, tableobj)
    PlaySound soundname, 1, 1, AudioPan(tableobj), 0,0,0, 1, AudioFade(tableobj)
End Sub

Sub PlaySoundAtBall(soundname)
    PlaySoundAt soundname, ActiveBall
End Sub

' *********************************************************************
'                      Supporting Ball & Sound Functions
' *********************************************************************

Function Vol(ball) ' Calculates the Volume of the sound based on the ball speed
    Vol = Csng(BallVel(ball) ^2 / 1000)
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

Function Pitch(ball) ' Calculates the pitch of the sound based on the ball speed
    Pitch = BallVel(ball) * 20
End Function

Function BallVel(ball) 'Calculates the ball speed
    BallVel = INT(SQR((ball.VelX ^2) + (ball.VelY ^2) ) )
End Function

Sub Pins_Hit (idx)
	PlaySound "pinhit_low", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0
End Sub

Sub Targets_Hit (idx)
	PlaySound "target", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0
End Sub

Sub Metals_Thin_Hit (idx)
	PlaySound "metalhit_thin", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0
End Sub

Sub Metals_Medium_Hit (idx)
	PlaySound "metalhit_medium", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0
End Sub

Sub Metals2_Hit (idx)
	PlaySound "metalhit2", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0
End Sub

Sub Gates_Hit (idx)
	PlaySound "gate4", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0
End Sub

Sub Spinner_Spin
	PlaySound "fx_spinner",0,.25,0,0.25
End Sub

Sub Rubbers_Hit(idx)
 	dim finalspeed
  	finalspeed=SQR(activeball.velx * activeball.velx + activeball.vely * activeball.vely)
 	If finalspeed > 20 then
		PlaySound "fx_rubber2", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0
	End if
	If finalspeed >= 6 AND finalspeed <= 20 then
 		RandomSoundRubber()
 	End If
End Sub

Sub Posts_Hit(idx)
 	dim finalspeed
  	finalspeed=SQR(activeball.velx * activeball.velx + activeball.vely * activeball.vely)
 	If finalspeed > 16 then
		PlaySound "fx_rubber2", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0
	End if
	If finalspeed >= 6 AND finalspeed <= 16 then
 		RandomSoundRubber()
 	End If
End Sub

Sub RandomSoundRubber()
	Select Case Int(Rnd*3)+1
		Case 1 : PlaySound "rubber_hit_1", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0
		Case 2 : PlaySound "rubber_hit_2", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0
		Case 3 : PlaySound "rubber_hit_3", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0
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
		Case 1 : PlaySound "flip_hit_1", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0
		Case 2 : PlaySound "flip_hit_2", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0
		Case 3 : PlaySound "flip_hit_3", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0
	End Select
End Sub

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
        If BallVel(BOT(b) ) > 1 AND BOT(b).z < 30 Then
            rolling(b) = True
            PlaySound("fx_ballrolling" & b), -1, Vol(BOT(b) ), Pan(BOT(b) ), 0, Pitch(BOT(b) ), 1, 0
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
	PlaySound("fx_collide"), 0, Csng(velocity) ^2 / 2000, Pan(ball1), 0, Pitch(ball1), 0, 0
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


'cFastFlips by nFozzy
'Bypasses pinmame callback for faster and more responsive flippers
'Version 1.0

'Flipper / game-on Solenoid # reference (incomplete):
'Williams System 11: Sol23 or 24
'Gottlieb System 3: Sol32
'Data East (pre-whitestar): Sol23 or 24
'WPC 90', 92', WPC Security: Sol31

'********************Setup*******************:

'...keydown section... (comment out the upper flippers as needed)
'If KeyCode = LeftFlipperKey then FastFlips.FlipL True :  FastFlips.FlipUL True
'If KeyCode = RightFlipperKey then FastFlips.FlipR True :  FastFlips.FlipUR True
'(Do not use Exit Sub, this script does not handle switch handling at all!)

'...keyUp section...
'If KeyCode = LeftFlipperKey then FastFlips.FlipL False :  FastFlips.FlipUL False
'If KeyCode = RightFlipperKey then FastFlips.FlipR False :  FastFlips.FlipUR False

'...Flipper Callbacks....
'if pinmame flipper callbacks are in use, comment them out. For example 'SolCallback(sLRFlipper)
'But use these subs (they should be the same ones defined in CallBackL / CallBackR) to handle flipper rotation and sounds!

'...Solenoid... set this based on solenoid number
SolCallBack(19) = "FastFlips.TiltSol"
SolCallback(sLRFlipper) = ""
SolCallback(sLLFlipper) = ""
SolCallback(sURFlipper) = ""
SolCallback(sULFlipper) = ""
'//////for a reference of solenoid numbers, see top /////


'One last note - Because this script is super simple it will call flipper return a lot.
'It might be a good idea to add extra conditional logic to your flipper return sounds so they don't play every time the game on solenoid turns off
'Example:
'Instead of
        'LeftFlipper.RotateToStart
        'playsound SoundFX("FlipperDown",DOFFlippers), 0, 1, 0.01   'return
'Add Extra conditional logic:
        'LeftFlipper.RotateToStart
        'if LeftFlipper.CurrentAngle = LeftFlipper.StartAngle then
        '   playsound SoundFX("FlipperDown",DOFFlippers), 0, 1, 0.01    'return
        'end if
'That's it]
'*************************************************
dim FastFlips
Set FastFlips = new cFastFlips
with FastFlips
    .CallBackL = "SolLflipper"  'Point these to flipper subs
    .CallBackR = "SolRflipper"  '...
'    .CallBackUL = "SolULflipper"'...(upper flippers, if needed)
'    .CallBackUR = "SolURflipper"'...
'   .TiltObjects = True 'Optional, if True calls vpmnudge.solgameon automatically. IF YOU GET A LINE 1 ERROR, DISABLE THIS! (or setup vpmNudge.TiltObj!)
'   .InitDelay "FastFlips", 100         'Optional, if > 0 adds some compensation for solenoid jitter (occasional problem on Bram Stoker's Dracula)
'   .DebugOn = False        'Debug, always-on flippers. Call FastFlips.DebugOn True or False in debugger to enable/disable.
end with

Class cFastFlips
    Public TiltObjects, DebugOn
    Private SubL, SubUL, SubR, SubUR, FlippersEnabled, Delay, LagCompensation, Name

    Private Sub Class_Initialize()
        Delay = 0 : FlippersEnabled = False : DebugOn = False : LagCompensation = False
    End Sub

    'set callbacks
    Public Property Let CallBackL(aInput)  : Set SubL  = GetRef(aInput) : End Property
    Public Property Let CallBackUL(aInput) : Set SubUL = GetRef(aInput) : End Property
    Public Property Let CallBackR(aInput)  : Set SubR  = GetRef(aInput) : End Property
    Public Property Let CallBackUR(aInput) : Set SubUR = GetRef(aInput) : End Property
    Public Sub InitDelay(aName, aDelay) : Name = aName : delay = aDelay : End Sub   'Create Delay

    'call callbacks
    Public Sub FlipL(aEnabled)
        if not FlippersEnabled and not DebugOn then Exit Sub
        subL aEnabled
    End Sub

    Public Sub FlipR(aEnabled)
        if not FlippersEnabled and not DebugOn then Exit Sub
        subR aEnabled
    End Sub

    Public Sub FlipUL(aEnabled)
        if not FlippersEnabled and not DebugOn then Exit Sub
        subUL aEnabled
    End Sub

    Public Sub FlipUR(aEnabled)
        if not FlippersEnabled and not DebugOn then Exit Sub
        subUR aEnabled
    End Sub

    Public Sub TiltSol(aEnabled)    'Handle solenoid / Delay (if delayinit)
        if delay > 0 and not aEnabled then  'handle delay
            vpmtimer.addtimer Delay, Name & ".FireDelay" & "'"
            LagCompensation = True
        else
            if Delay > 0 then LagCompensation = False
            EnableFlippers(aEnabled)
        end if
    End Sub

    Sub FireDelay() : if LagCompensation then EnableFlippers False End If : End Sub

    Private Sub EnableFlippers(aEnabled)
        FlippersEnabled = aEnabled
        if TiltObjects then vpmnudge.solgameon aEnabled
        If Not aEnabled then
            subL False
            subR False
            if not IsEmpty(subUL) then subUL False
            if not IsEmpty(subUR) then subUR False
        End If
    End Sub

End Class

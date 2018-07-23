Option Explicit
Randomize

On Error Resume Next
ExecuteGlobal GetTextFile("controller.vbs")
If Err Then MsgBox "You need the controller.vbs in order to run this table, available in the vp10 package"
On Error Goto 0

Const cGameName="sureshot",UseSolenoids=1,UseLamps=0,UseGI=0,SSolenoidOn="SolOn",SSolenoidOff="SolOff", SCoin="coin"

LoadVPM "01300000","Taito.VBS",3.1
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

'Solenoids
'
SolCallback(sLRFlipper) = "SolRFlipper"
SolCallback(sLLFlipper) = "SolLFlipper"

SolCallback(1)="SolBallRelease"
SolCallback(2)="SolSaucer"
SolCallback(3)="SolDropsUpL"
SolCallback(4)="SolDropsUp8"
SolCallback(5)="SolDropsUpR"
SolCallback(6)="dtbank1.SolHit 1,"
SolCallback(7)="dtbank1.SolHit 2,"
SolCallback(8)="dtbank1.SolHit 3,"
SolCallback(9)="dtbank1.SolHit 4,"
SolCallback(10)="dtbank1.SolHit 5,"
SolCallback(11)="dtbank1.SolHit 6,"
SolCallback(12)="dtbank1.SolHit 7,"
SolCallback(18)="vpmNudge.SolGameOn"

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
         bsTP.SolOut Enable
 End Sub
 
 Sub SolBallRelease(Enable)
         bsTrough.SolOut Enable
 End Sub
 
 Sub SolDropsUpL(Enable)
	dim xx
         dtbank2.SolDropUp Enable
		 For each xx in DTLeftLights: xx.state=0:Next
 End Sub

Sub SolDropsUp8 (Enable)
	dim xx
		dtbank3.SolDropUp Enable
		For each xx in dt8Lights: xx.state=0:Next
End Sub

Sub SolDropsUpR(enabled)
	dim xx
		dtBank1.SolDropUp enabled
		For each xx in DTRightLights: xx.state=0:Next
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

 Dim bsTrough, bsTP, dtBank1, dtBank2, dtBank3, dtBank1up

 Sub Table1_Init
	vpmInit Me
	On Error Resume Next
		With Controller
		.GameName = cGameName
		If Err Then MsgBox "Can't start Game" & cGameName & vbNewLine & Err.Description : Exit Sub
		.SplashInfoLine = "Sure Shot (Taito 1981)"&chr(13)&"v. 1.0"
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
 
     PinMAMETimer.Interval = PinMAMEInterval
     PinMAMETimer.Enabled = 1

	vpmNudge.TiltSwitch  = 30
	vpmNudge.Sensitivity = 2
     vpmNudge.TiltObj = Array(Bumper1,Bumper2,Bumper3, LeftSlingshot, RightSlingshot)
 
     Set bsTrough = New cvpmBallStack
         bsTrough.InitSw 0, 1, 0, 0, 0, 0, 0, 0
         bsTrough.InitKick BallRelease, 120, 3
         bsTrough.InitExitSnd SoundFX("ballrelease",DOFContactors), SoundFX("Solenoid",DOFContactors)
         bsTrough.Balls = 1
 
     Set bsTP = New cvpmBallStack
         bsTP.InitSaucer sw2, 2, 315, 18
         bsTP.InitExitSnd SoundFX("Popper",DOFContactors), SoundFX("Solenoid",DOFContactors)
         bsTP.KickAngleVar = 3
         bsTP.KickForceVar = 3
 
     set dtbank1 = new cvpmdroptarget
         dtbank1.initdrop array(sw3, sw13, sw23, sw33, sw43, sw53, sw63), array(3, 13, 23, 33, 43, 53, 63)
         dtbank1.initsnd  SoundFX("DTDrop",DOFContactors),SoundFX("DTReset",DOFContactors)
 
     set dtbank2 = new cvpmdroptarget
         dtbank2.initdrop array(sw12, sw22, sw32, sw42), array(12, 22, 32, 42)
         dtbank2.initsnd  SoundFX("DTDrop",DOFContactors),SoundFX("DTReset",DOFContactors)
 
     set dtbank3 = new cvpmdroptarget
         dtbank3.initdrop sw52, 52
         dtbank3.initsnd  SoundFX("DTDrop",DOFContactors),SoundFX("DTReset",DOFContactors)
 
 End Sub
 
'**********************************************************************************************************
'Plunger code
'**********************************************************************************************************

Sub Table1_KeyDown(ByVal KeyCode)
	If KeyDownHandler(keycode) Then Exit Sub
	If keycode = PlungerKey Then Plunger.Pullback:playsound"plungerpull"
End Sub

Sub Table1_KeyUp(ByVal KeyCode)
	If KeyUpHandler(keycode) Then Exit Sub
	If keycode = PlungerKey Then Plunger.Fire:PlaySound"plunger"
End Sub

'**********************************************************************************************************
 
 ' Drain hole and kickers
Sub Drain_Hit:bsTrough.addball me : playsound"drain" : End Sub
Sub sw2_Hit:bsTP.AddBall 0 : playsound"popper_ball" : End Sub

 ' Droptargets
 Sub sw12_Dropped:dtbank2.Hit 1
	L1DA.state=1
End Sub

 Sub sw22_Dropped:dtbank2.Hit 2
	L1DB.state=1
    L101DB.state=1
    L117DB.state=1
End Sub

 Sub sw32_Dropped:dtbank2.Hit 3
	L1DC.state=1:L2DC.state=1
         L101DC.state=1
         L117DC.state=1
End Sub

 Sub sw42_Dropped:dtbank2.Hit 4
	L2DD.state=1
    L117DD.state=1
    L118DD.state=1
End Sub

 Sub sw3_Dropped:dtbank1.Hit 1
	L2DG.state=1:L1DG.state=1
End Sub

 Sub sw13_Dropped:dtbank1.Hit 2
	L2DF.state=1:L1DF.state=1
End Sub

 Sub sw23_Dropped:dtbank1.Hit 3
	L2DE.state=1:L1DE.state=1
End Sub

 Sub sw33_Dropped:dtbank1.Hit 4
	L2DD.state=1
End Sub

 Sub sw43_Dropped:dtbank1.Hit 5
	L1DD1.state=1:L1DD2.state=1
End Sub

 Sub sw53_Dropped:dtbank1.Hit 6:End Sub
 Sub sw63_Dropped:dtbank1.Hit 7:End Sub
 Sub sw52_Dropped:dtbank3.Hit 1:L1D8.state=1:L2D8.state=1:L117D8.state=1:End Sub

'Rollovers
 Sub sw11_Hit:Controller.Switch(11) = 1 : playsound"rollover" : End Sub 
 Sub sw11_UnHit:Controller.Switch(11) = 0:End Sub
 Sub sw61_Hit:Controller.Switch(61) = 1 : playsound"rollover" : End Sub 
 Sub sw61_UnHit:Controller.Switch(61) = 0:End Sub
 Sub sw51_Hit:Controller.Switch(51) = 1 : playsound"rollover" : End Sub 
 Sub sw51_UnHit:Controller.Switch(51) = 0:End Sub
 Sub sw41_Hit:Controller.Switch(15) = 1 : playsound"rollover" : End Sub 
 Sub sw41_UnHit:Controller.Switch(15) = 0:End Sub
 Sub sw62_Hit:Controller.Switch(62) = 1 : playsound"rollover" : End Sub 
 Sub sw62_UnHit:Controller.Switch(62) = 0:End Sub
 Sub sw31_Hit:Controller.Switch(31) = 1 : playsound"rollover" : End Sub 
 Sub sw31_UnHit:Controller.Switch(31) = 0:End Sub
 Sub sw71_Hit:Controller.Switch(71) = 1 : playsound"rollover" : End Sub 
 Sub sw71_UnHit:Controller.Switch(71) = 0:End Sub
'Bumpers

Sub Bumper1_Hit : vpmTimer.PulseSw(5) : playsound SoundFX("fx_bumper1",DOFContactors): End Sub
Sub Bumper2_Hit : vpmTimer.PulseSw(15) : playsound SoundFX("fx_bumper1",DOFContactors): End Sub
Sub Bumper3_Hit : vpmTimer.PulseSw(25) : playsound SoundFX("fx_bumper1",DOFContactors): End Sub

'Stand Up Targets
Sub sw21_Hit:vpmTimer.PulseSw 21:End Sub
Sub sw4_Hit:vpmTimer.PulseSw 4:End Sub
Sub sw14_Hit:vpmTimer.PulseSw 14:End Sub
Sub sw24_Hit:vpmTimer.PulseSw 24:End Sub
Sub sw34_Hit:vpmTimer.PulseSw 34:End Sub
Sub sw44_Hit:vpmTimer.PulseSw 44:End Sub
Sub sw54_Hit:vpmTimer.PulseSw 54:End Sub

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
    NFadeL 0, li0
    NFadeL 1, li1
    NFadeL 10, li10
    NFadeL 100, li100
    NFadeL 101, li101
    NFadeL 102, li102
    NFadeL 103, li103
    NFadeL 109, li109
    NFadeL 11, li11
    NFadeL 110, li110
    NFadeL 111, li111
    NFadeL 112, li112
    NFadeL 113, li113
    NFadeL 119, li119
    NFadeL 12, li12
    NFadeL 120, li120
    NFadeL 121, li121
    NFadeL 122, li122
    NFadeL 123, li123
    NFadeL 129, li129
    NFadeL 130, li130
    NFadeL 131, li131
    NFadeL 132, li132
    NFadeL 133, li133
    NFadeL 143, li143
    NFadeL 153, li153
    NFadeL 2, li2
    NFadeL 20, li20
    NFadeL 21, li21
    NFadeL 22, li22
    NFadeL 30, li30
    NFadeL 31, li31
    NFadeLm 32, li32a
    NFadeL 32, li32
    NFadeL 40, li40
    NFadeL 41, li41
    NFadeL 42, li42
    NFadeL 50, li50
    NFadeL 51, li51
    NFadeL 52, li52
    NFadeL 60, li60
    NFadeL 61, li61
    NFadeL 62, li62
    NFadeL 70, li70
    NFadeL 71, li71
    NFadeL 72, li72
    NFadeL 79, li79
    NFadeL 80, li80
    NFadeL 81, li81
    NFadeL 82, li82
    NFadeL 83, li83
    NFadeL 89, li89
    NFadeL 90, li90
    NFadeL 91, li91
    NFadeL 92, li92
    NFadeL 93, li93
    NFadeL 99, li99

    'backdrop lights
    'If VarHidden Then
    '    NFadeL 139, li139 ' player 1
    '    NFadeL 140, li140 ' player 2
    '    NFadeL 141, li141 ' player 3
    '    NFadeL 142, li142 ' player 4
     '   NFadeL 149, li149 'game over
     '   NFadeL 150, li150 ' tilt
     '   NFadeL 151, li151 ' novo
     '   NFadeL 152, li152 'recorde
    'End If
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

'**********Sling Shot Animations
' Rstep and Lstep  are the variables that increment the animation
'****************
Dim RStep, Lstep, LA, LB, RA, RB, RC, RD, RKick

Sub RightSlingShot_Slingshot
	vpmTimer.PulseSw 45
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
	vpmTimer.PulseSw 35
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

Sub sw65_Slingshot
	vpmTimer.PulseSw 65
    RubberRC.Visible = 0
    RubberRC1.Visible = 1
    RC = 0
    sw65.TimerEnabled = 1
End Sub

Sub sw65_Timer
    Select Case RC
        Case 3:RubberRC1.Visible = 0:RubberRC2.Visible = 1
        Case 4:RubberRC2.Visible = 0:RubberRC.Visible = 1:sw65.TimerEnabled = 0:
    End Select
    RC = RC + 1
End Sub

Sub sw55_Slingshot
	vpmTimer.PulseSw 55
    RubberRD.Visible = 0
    RubberRD1.Visible = 1
    RD = 0
    sw55.TimerEnabled = 1
End Sub

Sub sw55_Timer
    Select Case RD
        Case 3:RubberRD1.Visible = 0:RubberRD2.Visible = 1
        Case 4:RubberRD2.Visible = 0:RubberRD.Visible = 1:sw55.TimerEnabled = 0:
    End Select
    RD = RD + 1
End Sub

Sub SolSaucer(enabled)
	If enabled Then
		bsTP.ExitSol_On
		RKick = 0
		kickarmtop_prim.ObjRotX = -12
		RKickTimer.Enabled = 1
	End If
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
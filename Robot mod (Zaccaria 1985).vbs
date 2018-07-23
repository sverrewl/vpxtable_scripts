Option Explicit
Randomize

On Error Resume Next
ExecuteGlobal GetTextFile("controller.vbs")
If Err Then MsgBox "You need the controller.vbs in order to run this table, available in the vp10 package"
On Error Goto 0

Const cGameName="robot",UseSolenoids=1,UseLamps=0,UseGI=0,SSolenoidOn="SolOn",SSolenoidOff="SolOff", SCoin="coin"

LoadVPM "01200300","zac2.vbs",3.10
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
SolCallback(4)="vpmSolSound ""knocker"","
SolCallback(5)= "dtR.SolDropUp" 'Drop Targets
SolCallback(6)="SolRob1"
SolCallback(7)="SolRob2"
SolCallback(10)="SolRob3"
SolCallback(14)="SolRob4"
SolCallback(15)="SolRob5"
SolCallback(16)="solrobhit1"
SolCallback(18)="solrobhit2"
SolCallback(21)="solrobhit3"
SolCallback(22)="solrobhit4"
SolCallback(23)="solrobhit5"
SolCallback(24)="SolBallRelease"


SolCallback(sLRFlipper) = "SolRFlipper"
SolCallback(sLLFlipper) = "SolLFlipper"

Sub SolLFlipper(Enabled)
     If Enabled Then
         PlaySound"fx_Flipperup":LeftFlipper.RotateToEnd
     Else
         PlaySound "fx_Flipperdown":LeftFlipper.RotateToStart
     End If
  End Sub
  
Sub SolRFlipper(Enabled)
     If Enabled Then
         PlaySound"fx_Flipperup":RightFlipper.RotateToEnd
     Else
         PlaySound "fx_Flipperdown":RightFlipper.RotateToStart
     End If
End Sub

'Primitive Flippers
Sub FlipperTimer_Timer
	FlipperT12.roty = LeftFlipper.currentangle  '+ 180
	FlipperT10.roty = RightFlipper.currentangle '+ 45
End Sub

'**********************************************************************************************************

'Solenoid Controlled toys
'**********************************************************************************************************
'*****GI Lights On
dim xx
For each xx in GI:xx.State = 1: Next


Sub SolBallRelease(enabled)
	If enabled Then
		vpmTimer.PulseSw 16
		playsound "ballrelease"
		vpmTimer.AddTimer 2260,"bsTrough.ExitSol_On '"
	End if
End Sub

Sub SolRob1(enabled)
	if enabled then
		sw30pos=50:sw30t.enabled=1:playsound"robotup":sw30.isdropped=0:controller.switch(30)=0
	end if
End Sub
Sub SolRob2(enabled)
	if enabled then
		sw29pos=50:sw29t.enabled=1:playsound"robotup":sw29.isdropped=0:controller.switch(29)=0
	end if
End Sub
Sub SolRob3(enabled)
	if enabled then
		sw28pos=50:sw28t.enabled=1:playsound"robotup":sw28.isdropped=0:controller.switch(28)=0
	end if
End Sub
Sub SolRob4(enabled)
	if enabled then
		sw27pos=50:sw27t.enabled=1:playsound"robotup":sw27.isdropped=0:controller.switch(27)=0
	end if
End Sub
Sub SolRob5(enabled)
	if enabled then
		sw26pos=50:sw26t.enabled=1:playsound"robotup":sw26.isdropped=0:controller.switch(26)=0
	end if
End Sub

Sub solRobhit1(enabled)
	if enabled then
		sw30pos=0:sw30.isdropped=1:sw30.timerenabled=1:playsound"robotdrop":controller.switch(30)=1
	end if
End Sub
Sub solRobhit2(enabled)
	if enabled then
		sw29pos=0:sw29.isdropped=1:sw29.timerenabled=1:playsound"robotdrop":controller.switch(29)=1
	end if
End Sub
Sub solRobhit3(enabled)
	if enabled then
		sw28pos=0:sw28.isdropped=1:sw28.timerenabled=1:playsound"robotdrop":controller.switch(28)=1
	end if
End Sub
Sub solRobhit4(enabled)
	if enabled then
		sw27pos=0:sw27.isdropped=1:sw27.timerenabled=1:playsound"robotdrop":controller.switch(27)=1
	end if
End Sub
Sub solRobhit5(enabled)
	if enabled then
		sw26pos=0:sw26.isdropped=1:sw26.timerenabled=1:playsound"robotdrop":controller.switch(26)=1
	end if
End Sub

'**********************************************************************************************************

'Initiate Table
'**********************************************************************************************************

Dim bsTrough, dtR, SubSpeed

Sub Table1_Init
	vpmInit Me
	On Error Resume Next
		With Controller
		.GameName = cGameName
		If Err Then MsgBox "Can't start Game" & cGameName & vbNewLine & Err.Description : Exit Sub
		.SplashInfoLine = "Robot (Zaccaria)"&chr(13)&"You Suck"
		.HandleMechanics=0
		.HandleKeyboard=0
		.ShowDMDOnly=1
		.ShowFrame=0
		.ShowTitle=0
        .hidden = 1
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

	vpmNudge.TiltSwitch = swTilt
	vpmNudge.Sensitivity = 5
	vpmNudge.TiltObj = Array(Bumper1,Bumper2,Bumper3,LeftSlingShot,RightSlingShot,RightSlingShot1)

' Through
	Set bsTrough = New cvpmBallStack : 
	With bsTrough
		.InitSw 0,16,0,0,0,0,0,0
		.InitKick ballrelease,135,7
		.Balls = 1
	End With

	Set dtR = New cvpmDropTarget
		dtR.InitDrop Array(sw47,sw46,sw45,sw44,sw43),Array(47,46,45,44,43)
		dtR.InitSnd "DTDrop","DTReset"

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

 ' Drain hole
Sub Drain_Hit:playsound"drain":bsTrough.addball me:End Sub

'Star Triggers
Sub sw17_hit:Controller.Switch(17)=1 : playsound"rollover" : End Sub
Sub sw17_unhit:Controller.Switch(17)=0:end sub
Sub sw25_hit:Controller.Switch(25)=1 : playsound"rollover" : End Sub
Sub sw25_unhit:Controller.Switch(25)=0:end sub
Sub sw36_hit:Controller.Switch(36)=1 : playsound"rollover" : End Sub
Sub sw36_unhit:Controller.Switch(36)=0:end sub
Sub sw48a_hit:Controller.Switch(48)=1 : playsound"rollover" : End Sub
Sub sw48a_unhit:Controller.Switch(48)=0:end sub
Sub sw48b_hit:Controller.Switch(48)=1 : playsound"rollover" : End Sub
Sub sw48b_unhit:Controller.Switch(48)=0:end sub

'Wire Triggers
Sub sw18_hit:Controller.Switch(18)=1 : playsound"rollover" : End Sub
sub sw18_unhit:Controller.Switch(18)=0:end sub
Sub sw19_hit:Controller.Switch(19)=1 : playsound"rollover" : End Sub
sub sw19_unhit:Controller.Switch(19)=0:end sub
Sub sw22_hit:Controller.Switch(22)=1 : playsound"rollover" : End Sub
sub sw22_unhit:Controller.Switch(22)=0:end sub
Sub sw23_hit:Controller.Switch(23)=1 : playsound"rollover" : End Sub
sub sw23_unhit:Controller.Switch(23)=0:end sub
Sub sw24_hit:Controller.Switch(24)=1 : playsound"rollover" : End Sub
sub sw24_unhit:Controller.Switch(24)=0:end sub
Sub sw50_hit:Controller.Switch(50)=1 : playsound"rollover" : End Sub
sub sw50_unhit:Controller.Switch(50)=0:end sub
Sub sw51_hit:Controller.Switch(51)=1 : playsound"rollover" : End Sub
sub sw51_unhit:Controller.Switch(51)=0:end sub
Sub sw52_hit:Controller.Switch(52)=1 : playsound"rollover" : End Sub
sub sw52_unhit:Controller.Switch(52)=0:end sub

'Drop Targets
Sub Sw43_Hit:dtR.Hit 1 :End Sub  
Sub Sw44_Hit:dtR.Hit 2 :End Sub  
Sub Sw45_Hit:dtR.Hit 3 :End Sub
Sub Sw46_Hit:dtR.Hit 4 :End Sub  
Sub Sw47_Hit:dtR.Hit 5 :End Sub

'Bumpers
Sub Bumper1_Hit : vpmTimer.PulseSw(40) : playsound"fx_bumper1": End Sub
Sub Bumper2_Hit : vpmTimer.PulseSw(41) : playsound"fx_bumper1": End Sub
Sub Bumper3_Hit : vpmTimer.PulseSw(42) : playsound"fx_bumper1": End Sub

'Stand Up Targets
Sub sw31_hit:vpmTimer.PulseSw 31:End Sub
Sub sw32_hit:vpmTimer.PulseSw 32:End Sub
Sub sw33_hit:vpmTimer.PulseSw 33:End Sub
Sub sw34_hit:vpmTimer.PulseSw 34:End Sub
Sub Sw37_Hit:vpmTimer.PulseSw 37:End Sub 
Sub Sw38_Hit:vpmTimer.PulseSw 38:End Sub 
Sub Sw39_Hit:vpmTimer.PulseSw 39:End Sub


'Scoring Rubber
Sub sw49_hit:vpmTimer.PulseSw 49:End Sub

'Wire trigger on ramp
Sub sw53_hit:vpmTimer.PulseSw 53:End Sub

'Ramp Entrance Helper
Sub kicker1_Hit
 	SubSpeed=ABS(ActiveBall.VelY)
 	Kicker1.DestroyBall
 	kicker2.CreateBall
 	kicker2.Kick 180,SQR(SubSpeed)
	playsound"kicker_enter_center"
End Sub

'Robot Drop Target Triggers
Sub sw26_Hit:sw26.isdropped=1:sw26pos=0:sw26.timerenabled=1:playsound"robotdrop":controller.switch(26)=1:End Sub
Sub sw27_Hit:sw27.isdropped=1:sw27pos=0:sw27.timerenabled=1:playsound"robotdrop":controller.switch(27)=1:End Sub
Sub sw28_Hit:sw28.isdropped=1:sw28pos=0:sw28.timerenabled=1:playsound"robotdrop":controller.switch(28)=1:End Sub
Sub sw29_Hit:sw29.isdropped=1:sw29pos=0:sw29.timerenabled=1:playsound"robotdrop":controller.switch(29)=1:End Sub
Sub sw30_Hit:sw30.isdropped=1:sw30pos=0:sw30.timerenabled=1:playsound"robotdrop":controller.switch(30)=1:End Sub

Dim sw26pos,sw27pos,sw28pos,sw29pos,sw30pos
sw26pos=0:sw27pos=0:sw28pos=0:sw29pos=0:sw30pos=0:

Sub sw26_timer()
	sw26pos=sw26pos+5
	sw26b.rotandtra5=0-sw26pos
	if sw26pos=50 then me.timerenabled=0
End Sub
Sub sw26t_timer()
	sw26pos=sw26pos-5
	sw26b.rotandtra5=0-sw26pos
	if sw26pos=0 then me.enabled=0
End Sub
Sub sw27_timer()
	sw27pos=sw27pos+5
	sw27b.rotandtra5=0-sw27pos
	if sw27pos=50 then me.timerenabled=0
End Sub
Sub sw27t_timer()
	sw27pos=sw27pos-5
	sw27b.rotandtra5=0-sw27pos
	if sw27pos=0 then me.enabled=0
End Sub
Sub sw28_timer()
	sw28pos=sw28pos+5
	sw28b.rotandtra5=0-sw28pos
	if sw28pos=50 then me.timerenabled=0
End Sub
Sub sw28t_timer()
	sw28pos=sw28pos-5
	sw28b.rotandtra5=0-sw28pos
	if sw28pos=0 then me.enabled=0
End Sub
Sub sw29_timer()
	sw29pos=sw29pos+5
	sw29b.rotandtra5=0-sw29pos
	if sw29pos=50 then me.timerenabled=0
End Sub
Sub sw29t_timer()
	sw29pos=sw29pos-5
	sw29b.rotandtra5=0-sw29pos
	if sw29pos=0 then me.enabled=0
End Sub
Sub sw30_timer()
	sw30pos=sw30pos+5
	sw30b.rotandtra5=0-sw30pos
	if sw30pos=50 then me.timerenabled=0
End Sub
Sub sw30t_timer()
	sw30pos=sw30pos-5
	sw30b.rotandtra5=0-sw30pos
	if sw30pos=0 then me.enabled=0
End Sub


'Generic Sounds
Sub Trigger1_Hit : playsound"Ball Drop": End Sub
Sub Trigger2_Hit : playsound"Wire Ramp": End Sub
Sub Trigger3_Hit : playsound"kicker_enter_center": End Sub

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

	nfadeL 1,l1
	nfadeL 2,l2
	nfadeL 3,l3
	nfadeL 4,l4
	nfadeL 5,l5
	nfadeL 8,l8
	nfadeL 9,l9
	FadeObj 11,sw30b,"robot_on","robot_med","robot_low","robot"
	nfadeL 12,l12
	nfadeL 14,l14 'bumper
	nfadeL 15,l15 'bumper
	nfadeL 16,l16 'bumper
	nfadeL 18,l18
	nfadeL 19,l19
	FadeObj 21,sw29b,"robot_on","robot_med","robot_low","robot"
	nfadeL 22,l22
	nfadeL 23,l23
	nfadeL 24,l24
	nfadeL 25,l25
	nfadeL 26,l26
	nfadeL 28,l28
	nfadeL 27,l27 'Apron Light
	FadeObj 29,sw28b,"robot_on","robot_med","robot_low","robot"
	nfadeL 30,l30
	nfadeL 32,l32
	nfadeL 34,l34
	FadeObj 35,sw27b,"robot_on","robot_med","robot_low","robot"
	nfadeL 36,l36
	nfadeL 38,l38
	nfadeL 39,l39
	FadeObj 40,sw26b,"robot_on","robot_med","robot_low","robot"
	nfadeL 41,l41
	nfadeL 42,l42
	nfadeL 43,l43
	nfadeL 44,l44
	nfadeL 45,l45
	nfadeL 47,l47
	nfadeL 48,l48
	nfadeL 49,l49
	nfadeL 51,l51
	nfadeL 53,l53
	nfadeL 55,l55
	nfadeL 57,l57
	nfadeL 58,l58
	nfadeL 59,l59
	nfadeL 61,l61
	nfadeL 63,l63
	nfadeL 64,l64
	nfadeL 65,l65
	nfadeL 68,l68
	nfadeL 69,l69
	nfadeL 70,l70
	nfadeL 71,l71
	nfadeL 73,l73
	nfadeL 75,l75
	nfadeL 79,l79
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




'**********************************************************************************************************
'Digital Display
'**********************************************************************************************************
Dim Digits(40)

Digits(0) = Array(LED10,LED11,LED12,LED13,LED14,LED15,LED16)
Digits(1) = Array(LED20,LED21,LED22,LED23,LED24,LED25,LED26)
Digits(2) = Array(LED30,LED31,LED32,LED33,LED34,LED35,LED36)
Digits(3) = Array(LED40,LED41,LED42,LED43,LED44,LED45,LED46)
Digits(4) = Array(LED50,LED51,LED52,LED53,LED54,LED55,LED56)
Digits(5) = Array(LED60,LED61,LED62,LED63,LED64,LED65,LED66)
Digits(6) = Array(LED70,LED71,LED72,LED73,LED74,LED75,LED76)
Digits(7) = Array(LED80,LED81,LED82,LED83,LED84,LED85,LED86)
Digits(8) = Array(LED90,LED91,LED92,LED93,LED94,LED95,LED96)
Digits(9) = Array(LED100,LED101,LED102,LED103,LED104,LED105,LED106)
Digits(10) = Array(LED110,LED111,LED112,LED113,LED114,LED115,LED116)
Digits(11) = Array(LED120,LED121,LED122,LED123,LED124,LED125,LED126)
Digits(12) = Array(LED130,LED131,LED132,LED133,LED134,LED135,LED136)
Digits(13) = Array(LED140,LED141,LED142,LED143,LED144,LED145,LED146)
Digits(14) = Array(LED150,LED151,LED152,LED153,LED154,LED155,LED156)
Digits(15) = Array(LED160,LED161,LED162,LED163,LED164,LED165,LED166)
Digits(16) = Array(LED170,LED171,LED172,LED173,LED174,LED175,LED176)
Digits(17) = Array(LED180,LED181,LED182,LED183,LED184,LED185,LED186)
Digits(18) = Array(LED190,LED191,LED192,LED193,LED194,LED195,LED196)
Digits(19) = Array(LED200,LED201,LED202,LED203,LED204,LED205,LED206)
Digits(20) = Array(LED210,LED211,LED212,LED213,LED214,LED215,LED216)
Digits(21) = Array(LED220,LED221,LED222,LED223,LED224,LED225,LED226)
Digits(22) = Array(LED230,LED231,LED232,LED233,LED234,LED235,LED236)
Digits(23) = Array(LED240,LED241,LED242,LED243,LED244,LED245,LED246)
Digits(24) = Array(LED250,LED251,LED252,LED253,LED254,LED255,LED256)
Digits(25) = Array(LED260,LED261,LED262,LED263,LED264,LED265,LED266)
Digits(26) = Array(LED270,LED271,LED272,LED273,LED274,LED275,LED276)
Digits(27) = Array(LED280,LED281,LED282,LED283,LED284,LED285,LED286)
Digits(28) = Array(LED4,LED2,LED6,LED7,LED5,LED1,LED3)
Digits(29) = Array(LED18,LED9,LED27,LED28,LED19,LED8,LED17)
Digits(30) = Array(LED39,LED37,LED48,LED49,LED47,LED29,LED38)
Digits(31) = Array(LED67,LED58,LED69,LED77,LED68,LED57,LED59)
Digits(32) = Array(LED88,LED79,LED97,LED98,LED89,LED78,LED87)
Digits(33) = Array(LED109,LED107,LED118,LED119,LED117,LED99,LED108)
Digits(34) = Array(LED137,LED128,LED139,LED147,LED138,LED127,LED129)
Digits(35) = Array(LED158,LED149,LED167,LED168,LED159,LED148,LED157)
Digits(36) = Array(LED179,LED177,LED188,LED189,LED187,LED169,LED178)
Digits(37) = Array(LED207,LED198,LED209,LED217,LED208,LED197,LED199)
Digits(38) = Array(LED228,LED219,LED237,LED238,LED229,LED218,LED227)
Digits(39) = Array(LED249,LED247,LED258,LED259,LED257,LED239,LED248)


 Sub DisplayTimer_Timer
    Dim ChgLED, ii, jj, num, chg, stat, obj, b, x
    ChgLED=Controller.ChangedLEDs(&Hffffffff, &Hffffffff)
    If Not IsEmpty(ChgLED)Then
		If DesktopMode = True Then
       For ii=0 To UBound(chgLED)
          num=chgLED(ii, 0) : chg=chgLED(ii, 1) : stat=chgLED(ii, 2)
			if (num < 40) then
              For Each obj In Digits(num)
                   If chg And 1 Then obj.State=stat And 1
                   chg=chg\2 : stat=stat\2
                  Next
			Else
			       end if
        Next
	   end if
    End If
 End Sub

'**********************************************************************************************************
'**********************************************************************************************************



'**********Sling Shot Animations
' Rstep and Lstep  are the variables that increment the animation
'****************
Dim RStep, Lstep, R1Step

Sub RightSlingShot_Slingshot
	vpmTimer.PulseSw 21
    PlaySound "right_slingshot", 0, 1, 0.05, 0.05
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
	vpmTimer.PulseSw 20
    PlaySound "left_slingshot",0,1,-0.05,0.05
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

Sub RightSlingShot1_Slingshot
	vpmTimer.PulseSw 35
    PlaySound "right_slingshot", 0, 1, 0.05, 0.05
    R1Sling.Visible = 0
    R1Sling1.Visible = 1
    sling3.TransZ = -20
    R1Step = 0
    RightSlingShot1.TimerEnabled = 1
End Sub

Sub RightSlingShot1_Timer
    Select Case RStep
        Case 3:R1SLing1.Visible = 0:R1SLing2.Visible = 1:sling3.TransZ = -10
        Case 4:R1SLing2.Visible = 0:R1SLing.Visible = 1:sling3.TransZ = 0:RightSlingShot1.TimerEnabled = 0:
    End Select
    R1Step = R1Step + 1
End Sub

' *********************************************************************
'                      Supporting Ball & Sound Functions
' *********************************************************************

Function Vol(ball) ' Calculates the Volume of the sound based on the ball speed
    Vol = Csng(BallVel(ball) ^2 / 2000)
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

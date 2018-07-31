Option Explicit
Randomize

' Thalamus 2018-07-23
' Added/Updated "Positional Sound Playback Functions" and "Supporting Ball & Sound Functions"
' No special SSF tweaks yet.

On Error Resume Next
ExecuteGlobal GetTextFile("controller.vbs")
If Err Then MsgBox "Can't open controller.vbs"
On Error Goto 0

Const cGameName="galaxy",UseSolenoids=2,UseLamps=0,UseGI=0,SSolenoidOn="SolOn",SSolenoidOff="SolOff",SFlipperOn="fx_Flipperup",SFlipperOff="fx_Flipperdown"
Const SCoin="coin"


LoadVPM "01000100", "Stern.VBS", 3.26

Dim DesktopMode: DesktopMode = Table1.ShowDT

Dim BallShadows: Ballshadows=1  		'******************set to 1 to turn on Ball shadows
Dim FlipperShadows: FlipperShadows=1  '***********set to 1 to turn on Flipper shadows

'Outlane adjustments (seriously, don't do this):
Dim Dullard: Dullard=0
Dim Coward: Coward=0

Dim hiddenvalue

If DesktopMode = True Then 'Show Desktop components
Ramp16.visible=1
Ramp15.visible=1
hiddenvalue=0
Else
Ramp16.visible=0
Ramp15.visible=0
hiddenvalue=1
End if

Dim EnableBallControl
EnableBallControl = false 'Change to true to enable manual ball control (or press C in-game) via the arrow keys and B (boost movement) keys

'Solenoid Call backs
'**********************************************************************************************************
'SolCallback(1)	=  rsling
'SolCallback(2)	= lsling
'SolCallback(3)	= topbump
SolCallback(4)	= "dtdrop"
SolCallback(5)	= "solleftout"
SolCallback(6)  = "vpmSolSound SoundFX(""Knocker"",DOFKnocker),"
'SolCallback(7)	= rbump
SolCallback(8)	=  "solsaucerout"
SolCallback(9)	=  "drop1"
SolCallback(10)	=  "drop2"
SolCallback(11)	=  "bstrough.solout"
'SolCallback(12) = lbump
SolCallback(13)	=  "drop3"
SolCallback(14)	= "drop4"
SolCallback(15)	= "BRelease"
SolCallback(17) = "PFGI"
SolCallback(19)	= "vpmNudge.SolGameOn"
'SolCallback(sLLFlipper)="vpmSolFlipper LeftFlipper,nothing,"
'SolCallback(sLRFlipper)="vpmSolFlipper RightFlipper,nothing,"

	SolCallback(sLRFlipper) = "SolRFlipper"
	SolCallback(sLLFlipper) = "SolLFlipper"

	Sub SolLFlipper(Enabled)
		 If Enabled Then
			 PlaySound SoundFX("fx_Flipperup",DOFContactors):LeftFlipper.RotateToEnd
		 Else
			 PlaySound SoundFX("fx_Flipperdown",DOFContactors):LeftFlipper.RotateToStart
		 End If
	  End Sub

	Sub SolRFlipper(Enabled)
		 If Enabled Then
			 PlaySound SoundFX("fx_Flipperup",DOFContactors):RightFlipper.RotateToEnd
		 Else
			 PlaySound SoundFX("fx_Flipperdown",DOFContactors):RightFlipper.RotateToStart
		 End If
	End Sub

'Solenoid Controlled toys
'**********************************************************************************************************
Sub Drop1(enabled) : If enabled Then: PlaySound SoundFX("DTReset",DOFDropTargets) :sw4.IsDropped = 1 : controller.switch(4) = 1 : End If : End Sub
Sub Drop2(enabled) : If enabled Then: PlaySound SoundFX("DTReset",DOFDropTargets) :sw11.IsDropped = 1 : controller.switch(11) = 1 : End If : End Sub
Sub Drop3(enabled) : If enabled Then: PlaySound SoundFX("DTReset",DOFDropTargets) :sw10.IsDropped = 1 : controller.switch(10) = 1 : End If : End Sub
Sub Drop4(enabled) : If enabled Then: PlaySound SoundFX("DTReset",DOFDropTargets) :sw9.IsDropped = 1 : controller.switch(9) = 1 : End If : End Sub


Sub DTdrop(enabled)
	if enabled then
		dtL.SolDropUp enabled
	end if
End Sub

'Playfield GI
Sub PFGI(Enabled)
	If Enabled Then
'		GiOFF
		dim xx
		For each xx in GI1:xx.State = 0: Next
        PlaySound "fx_relay"
		Table1.ColorGradeImage = "ColorGradeLUT256x16_shadowcrush"
		flasher1.visible=0
		flasher2.visible=1
		gatebracket_prim.blenddisablelighting = 0.05
		upperbracket_prim.blenddisablelighting = 0.05
		spinbracket_prim.blenddisablelighting = 0.05
		plungeguide_prim.blenddisablelighting = 0.05
		bumpercap1_prim.blenddisablelighting = 0.05
		bumpercap2_prim.blenddisablelighting = 0.05
		bumpercap3_prim.blenddisablelighting = 0.05
		bumpercap1_prim.image = "bump1off"
		bumpercap2_prim.image = "bump2off"
		bumpercap3_prim.image = "bump3off"
		plasticedges_prim.blenddisablelighting = 0
		plasticedgeslow_prim.blenddisablelighting = 0
		flipperright_prim.blenddisablelighting = 0
		flipperleft_prim.blenddisablelighting = 0
		sw4.blenddisablelighting = 0
		sw4.image = "drop1off"
		sw11.blenddisablelighting = 0
		sw11.image = "drop2off"
		sw10.blenddisablelighting = 0
		sw10.image = "drop3off"
		sw9.blenddisablelighting = 0
		sw9.image = "drop4off"
		sw26.blenddisablelighting = 0
		sw29.blenddisablelighting = 0
		sw30.blenddisablelighting = 0
		outers_prim.blenddisablelighting = 0
		innerwood_prim.blenddisablelighting = 0
		drop1_S.image="blank"
		drop2_S.image="blank"
		drop3_S.image="blank"
		drop4_S.image="blank"

	Else
'		GiON
		For each xx in GI1:xx.State = 1: Next
        PlaySound "fx_relay"
		Table1.ColorGradeImage = "ColorGradeLUT256x16_1to1"
		flasher1.visible=1
		flasher2.visible=0
		gatebracket_prim.blenddisablelighting = 0.4
		upperbracket_prim.blenddisablelighting = 0.4
		spinbracket_prim.blenddisablelighting = 0.05
		plungeguide_prim.blenddisablelighting = 0.4
		bumpercap1_prim.blenddisablelighting = 0
		bumpercap2_prim.blenddisablelighting = 0
		bumpercap3_prim.blenddisablelighting = 0
		bumpercap1_prim.image = "bump1on"
		bumpercap2_prim.image = "bump2on"
		bumpercap3_prim.image = "bump3on"
		plasticedges_prim.blenddisablelighting = 0.7
		plasticedgeslow_prim.blenddisablelighting = 0.7
		flipperright_prim.blenddisablelighting = 0.25
		flipperleft_prim.blenddisablelighting = 0.25
		sw4.blenddisablelighting = 0.9
		sw4.image = "drop1"
		sw11.blenddisablelighting = 0.9
		sw11.image = "drop2"
		sw10.blenddisablelighting = 0.9
		sw10.image = "drop3"
		sw9.blenddisablelighting = 0.9
		sw9.image = "drop4"
		sw26.blenddisablelighting = 0.3
		sw29.blenddisablelighting = 0.3
		sw30.blenddisablelighting = 0.3
		outers_prim.blenddisablelighting = 0.8
		innerwood_prim.blenddisablelighting = 0.8
		drop1_s.visible = 1
		drop2_s.visible = 1
		drop3_s.visible = 1
		drop4_s.visible = 1
		If sw4.isdropped=1 then drop1_S.image="blank" End If
		If sw4.isdropped=0 then drop1_S.image="drop1shadow" End If
		If sw11.isdropped=1 then drop2_S.image="blank" End If
		If sw11.isdropped=0 then drop2_S.image="drop2shadow" End If
		If sw10.isdropped=1 then drop3_S.image="blank" End If
		If sw10.isdropped=0 then drop3_S.image="drop3shadow" End If
		If sw9.isdropped=1 then drop4_S.image="blank" End If
		If sw9.isdropped=0 then drop4_S.image="drop4shadow" End If
end if
End Sub

'**********************************************************************************************************

'Initiate Table
'**********************************************************************************************************

Dim bsTrough, bskicker, bssaucer, dtL

Sub Table1_Init
'	vpmInit Me
'	On Error Resume Next
	With Controller
		.GameName = cGameName
		If Err Then MsgBox "Can't start Game" & cGameName & vbNewLine & Err.Description : Exit Sub
		.SplashInfoLine = "Galaxy (Stern 1980)"
		.HandleMechanics=0
		.HandleKeyboard=0
		.ShowDMDOnly=1
		.ShowFrame=0
		.ShowTitle=0
        .hidden = hiddenvalue
'		.hidden = 1
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
	vpmNudge.TiltSwitch = 7
	vpmNudge.Sensitivity = 2
	vpmNudge.TiltObj = Array(Bumper1,Bumper2,LeftSlingshot,RightSlingshot)

   	Set bsTrough=New cvpmBallStack
 	with bsTrough
		.InitSw 0,33,0,0,0,0,0,0
		.InitKick BallRelease,90,7
		.InitExitSnd Soundfx("popper",DOFContactors), Soundfx("solenoid",DOFContactors)
		.Balls=1
 	end with

 	Set dtL=New cvpmDropTarget
		dtL.InitDrop Array(sw4,sw11,sw10,sw9),Array(4,11,10,9)
		dtL.InitSnd SoundFX("DTDrop",DOFDropTargets),SoundFX("DTReset",DOFContactors)
End Sub

'**********************************************************************************************************
'Plunger code
'**********************************************************************************************************

Sub Table1_KeyDown(ByVal KeyCode)
	If KeyDownHandler(keycode) Then Exit Sub
	If keycode = PlungerKey Then Plunger.PullBack: PlaySound "plungerpull",0,1,0.25,0.25

   ' Manual Ball Control
	If keycode = 46 Then	 				' C Key
		If EnableBallControl = 1 Then
			EnableBallControl = 0
		Else
			EnableBallControl = 1
		End If
	End If
    If EnableBallControl = 1 Then
		If keycode = 48 Then 				' B Key
			If BCboost = 1 Then
				BCboost = BCboostmulti
			Else
				BCboost = 1
			End If
		End If
		If keycode = 203 Then BCleft = 1	' Left Arrow
		If keycode = 200 Then BCup = 1		' Up Arrow
		If keycode = 208 Then BCdown = 1	' Down Arrow
		If keycode = 205 Then BCright = 1	' Right Arrow
	End If
End Sub

Sub Table1_KeyUp(ByVal KeyCode)
	If KeyUpHandler(keycode) Then Exit Sub
	If keycode = PlungerKey Then Plunger.Fire: PlaySound "plunger",0,1,0.25,0.25

    'Manual Ball Control
	If EnableBallControl = 1 Then
		If keycode = 203 Then BCleft = 0	' Left Arrow
		If keycode = 200 Then BCup = 0		' Up Arrow
		If keycode = 208 Then BCdown = 0	' Down Arrow
		If keycode = 205 Then BCright = 0	' Right Arrow
	End If
End Sub
'**********************************************************************************************************
'**********************************************************************************************************

 'switches
'**********************************************************************************************************
'kicker
 Sub sw40_Hit:playsound"kickerenter": Controller.Switch(40) = 1: End Sub
Sub sw40_Hit
	playsound "kickerenter"
	Controller.Switch(40) = 1
	topsaucersafetytimer.enabled=1
End Sub

'kicker safety in case kickout stalls
Sub topsaucersafetytimer_timer
	sw40.Kick 100, 5 + 5 * Rnd
'testing	playsound "bell"
	topsaucersafetytimer.enabled=0
End Sub

 ' Drain kickers
 Sub Drain_Hit:playsound"drain":bstrough.addball me:End Sub

'Drop Targets
 Sub Sw4_Dropped:dtL.Hit 1 : End Sub
 Sub Sw11_Dropped:dtL.Hit 2 : End Sub
 Sub Sw10_Dropped:dtL.Hit 3 : End Sub
 Sub Sw9_Dropped:dtL.Hit 4 : End Sub

'Stand Up Targets
 Sub sw26_Hit:vpmTimer.PulseSW 26:End Sub
 Sub sw29_Hit:vpmTimer.PulseSW 29:End Sub
 Sub sw30_Hit:vpmTimer.PulseSW 30:End Sub

'Spinners
 Sub sw5_Spin : vpmTimer.PulseSw (5) :PlaySound "fx_spinner": End Sub

'Bumpers
  Sub Bumper1_Hit:vpmTimer.PulseSw 12 : playsound SoundFX("fx_bumper1",DOFContactors): End Sub
  Sub Bumper2_Hit:vpmTimer.PulseSw 13 : playsound SoundFX("fx_bumper1",DOFContactors): End Sub
  Sub Bumper3_Hit:vpmTimer.PulseSw 14 : playsound SoundFX("fx_bumper1",DOFContactors): End Sub

 'Star Rollovers
 Sub sw18_Hit   : Controller.Switch(18) = 1 : playsound"rollover" : End Sub
 Sub sw18_UnHit : Controller.Switch(18) = 0 : End Sub
 Sub sw18a_Hit  : Controller.Switch(18) = 1 : playsound"rollover" : End Sub
 Sub sw18a_UnHit: Controller.Switch(18) = 0 : End Sub
 Sub sw23_Hit   : Controller.Switch(23) = 1 : playsound"rollover" : End Sub
 Sub sw23_UnHit : Controller.Switch(23) = 0 : End Sub
 Sub sw24_Hit   : Controller.Switch(24) = 1 : playsound"rollover"  End Sub
 Sub sw24_UnHit : Controller.Switch(24) = 0 : End Sub
 Sub sw39_Hit   : Controller.Switch(39) = 1 : playsound"rollover" : End Sub
 Sub sw39_UnHit : Controller.Switch(39) = 0 : End Sub

 'Wire Triggers
 Sub sw19_Hit   : Controller.Switch(19) = 1 : playsound"rollover" : End Sub
 Sub sw19_UnHit : Controller.Switch(19) = 0 : End Sub
 Sub sw20_Hit   : Controller.Switch(20) = 1 : playsound"rollover" : End Sub
 Sub sw20_UnHit : Controller.Switch(20) = 0 : End Sub
 Sub sw21_Hit   : Controller.Switch(21) = 1 : playsound"rollover" : End Sub
 Sub sw21_UnHit : Controller.Switch(21) = 0 : End Sub
 Sub sw22_Hit  : Controller.Switch(22) = 1 : playsound"rollover" : End Sub
 Sub sw22_UnHit: Controller.Switch(22) = 0 : End Sub
 Sub sw35_Hit  : Controller.Switch(35) = 1 : playsound"rollover" : leftsaucersafetytimer.enabled = 1:End Sub
 Sub sw35_UnHit: Controller.Switch(35) = 0 : End Sub
 Sub sw36_Hit  : Controller.Switch(36) = 1 : playsound"rollover" : End Sub
 Sub sw36_UnHit: Controller.Switch(36) = 0 : End Sub
 Sub sw37_Hit  : Controller.Switch(37) = 1 : playsound"rollover" : End Sub
 Sub sw37_UnHit: Controller.Switch(37) = 0 : End Sub
 Sub sw38_Hit  : Controller.Switch(38) = 1 : playsound"rollover" : End Sub
 Sub sw38_UnHit: Controller.Switch(38) = 0 : End Sub

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
	NFadeLm 1, l1
   	NFadeLm 1, l1a
	NFadeLm 1, l1z
	NFadeLm 2, l2
   	NFadeLm 2, l2a
	NFadeLm 2, l2z
	NFadeLm 3, l3
   	NFadeLm 3, l3a
	NFadeLm 3, l3z
	NFadeLm 4, l4
   	NFadeLm 4, l4a
	NFadeLm 4, l4z
	NFadeLm 5, l5
   	NFadeLm 5, l5a
	NFadeLm 5, l5z
	NFadeLm 6, l6
   	NFadeLm 6, l6a
	NFadeLm 6, l6z
	NFadeLm 7, l7
   	NFadeLm 7, l7a
	NFadeLm 7, l7z
	NFadeLm 8, l8
   	NFadeLm 8, l8a
	NFadeLm 8, l8z
	NFadeLm 9, l9
   	NFadeLm 9, l9a
	NFadeLm 9, l9z
   	NFadeLm 10, l10
   	NFadeLm 10, l10a
   	NFadeLm 10, l10z
   	NFadeLm 11, l11
   	NFadeLm 11, l11a
   	NFadeLm 11, l11z
   	NFadeLm 12, l12
   	NFadeLm 12, l12a
   	NFadeLm 12, l12z
'   	NFadeLm 13, l13 'High Score To Date
   	NFadeLm 17, l17
   	NFadeLm 17, l17a
   	NFadeLm 17, l17z
   	NFadeLm 18, l18
   	NFadeLm 18, l18a
   	NFadeLm 18, l18z
   	NFadeLm 19, l19
   	NFadeLm 19, l19a
   	NFadeLm 19, l19z
   	NFadeLm 20, l20
   	NFadeLm 20, l20a
   	NFadeLm 20, l20z
   	NFadeLm 21, l21
   	NFadeLm 21, l21a
   	NFadeLm 21, l21z
   	NFadeLm 22, l22
   	NFadeLm 22, l22a
   	NFadeLm 22, l22z
   	NFadeLm 23, l23
   	NFadeLm 23, l23a
   	NFadeLm 23, l23z
   	NFadeLm 24, l24
   	NFadeLm 24, l24a
   	NFadeLm 24, l24z
   	NFadeLm 25, l25
   	NFadeLm 25, l25a
   	NFadeLm 25, l25z
   	NFadeLm 26, l26
   	NFadeLm 26, l26a
   	NFadeLm 26, l26z
   	NFadeLm 27, l27
   	NFadeLm 27, l27a
   	NFadeLm 27, l27z
   	NFadeLm 27, l29
   	NFadeLm 27, l29a
   	NFadeLm 27, l29z
   	NFadeLm 28, l28
   	NFadeLm 28, l28a
   	NFadeLm 28, l28z
   	NFadeLm 33, l33
   	NFadeLm 33, l33a
   	NFadeLm 33, l33z
   	NFadeLm 34, l34
   	NFadeLm 34, l34a
   	NFadeLm 34, l34z
   	NFadeLm 35, l35
   	NFadeLm 35, l35a
   	NFadeLm 35, l35z
	NFadeLm 36, l36
	NFadeLm 36, l36a
	NFadeLm 36, l36z
   	NFadeLm 37, l37
   	NFadeLm 37, l37a
   	NFadeLm 37, l37z
   	NFadeLm 38, l38
   	NFadeLm 38, l38a
   	NFadeLm 38, l38z
   	NFadeLm 39, l39
   	NFadeLm 39, l39a
   	NFadeLm 39, l39z
   	NFadeLm 40, l40
   	NFadeLm 40, l40a
   	NFadeLm 40, l40z
   	NFadeLm 41, l41
   	NFadeLm 41, l41a
   	NFadeLm 41, l41z
   	NFadeLm 42, l42
   	NFadeLm 42, l42a
   	NFadeLm 42, l42z
   	NFadeLm 43, l43
   	NFadeLm 43, l43a
   	NFadeLm 43, l43z
   	NFadeLm 44, l44
   	NFadeLm 44, l44a
   	NFadeLm 44, l44z
'	NFadeLm 45, l45 'Game Over
'	NFadeLm 48, l48
   	NFadeLm 49, l49
   	NFadeLm 49, l49a
   	NFadeLm 49, l49z
   	NFadeLm 50, l50
   	NFadeLm 50, l50a
   	NFadeLm 50, l50z
   	NFadeLm 51, l51
   	NFadeLm 51, l51a
   	NFadeLm 51, l51z
	NFadeLm 52, l52
	NFadeLm 52, l52a
	NFadeLm 52, l52z
   	NFadeLm 53, l53
   	NFadeLm 53, l53a
   	NFadeLm 53, l53z
   	NFadeLm 54, l54
   	NFadeLm 54, l54a
   	NFadeLm 54, l54z
   	NFadeLm 55, l55
   	NFadeLm 55, l55a
   	NFadeLm 55, l55z
   	NFadeLm 56, l56
   	NFadeLm 56, l56a
   	NFadeLm 56, l56z
	NFadeLm 57, l57
	NFadeLm 57, l57a
	NFadeLm 57, l57z
   	NFadeLm 58, l58
   	NFadeLm 58, l58a
   	NFadeLm 58, l58z
   	NFadeLm 59, l59
   	NFadeLm 59, l59a
   	NFadeLm 59, l59z
   	NFadeLm 60, l60
   	NFadeLm 60, l60a
   	NFadeLm 60, l60z
'	NFadeLm 61, l61 TILT
'   NFadeLm 63, l63 MATCH
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

Sub FadePrim(nr, pri, a, b, c, d)
    Select Case FadingLevel(nr)
        Case 2:pri.image = d:FadingLevel(nr) = 0
        Case 3:pri.image = c:FadingLevel(nr) = 2
        Case 4:pri.image = b:FadingLevel(nr) = 3
        Case 5:pri.image = a:FadingLevel(nr) = 1
    End Select
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

 'Stern Galaxy
 'added by Inkochnito
 Sub editDips
 	Dim vpmDips : Set vpmDips = New cvpmDips
 	With vpmDips
 		.AddForm 700,400,"Galaxy - DIP switches"
 		.AddFrame 0,0,190,"Maximum Credits",&H00060000,Array("10 Credits",0,"15 Credits",&H00020000,"25 Credits",&H00040000,"40 Credits",&H00060000)'dip 18&19
 		.AddFrame 0,76,190,"High Game To Date Awards",49152,Array("Nothing",0,"1 Replay",&H00004000,"2 Replays",32768,"3 Replays",49152)'dip 15&16
 		.AddFrame 0,152,190,"Replay Feature",&H00000020,Array("Add-A-Ball (Extra Ball Instead)",0,"Replay Enabled",&H00000020)'dip 6
 		.AddFrame 0,198,190,"Welfare (Inlanes Spot Drop Targets)",&H00000080,Array("Disabled At 4X",0,"Disabled At 3X",&H00000080)'dip 8
 		.AddFrame 0,244,190,"Extra Ball Feature",&H00010000,Array("Novelty (Extra Ball Disabled)",0,"Extra Ball Enabled",&H00010000)'dip 17
 		.AddChk 0,300,190,Array("Match Feature",&H00100000)'dip 21
 		.AddChk 0,315,190,Array("Credits Displayed",&H00080000)'dip 20
 		.AddFrame 205,0,190,"Special Award",&HC0000000,Array("Nothing",0,"100,000 Points",&H40000000,"Extra Ball",&H80000000,"Replay",&HC0000000)'dip 31&32
 		.AddFrame 205,76,190,"Planet Drop Target Special",&H03000000,Array("Lit At Pluto",0,"Lit At Uranus",&H01000000,"Lit At Neptune",&H02000000,"Lit At Saturn",&H03000000)'dip 25&26
 		.AddFrame 205,152,190,"'A' in GALAXY Difficulty",&H00800000,Array("Lite Each 'A' Individually",0,"Lite Both 'A's When One 'A' Lit",&H00800000)'dip 24
 		.AddFrame 205,198,190,"Balls Per Game",&H00000040,Array("3 Balls",0,"5 Balls",&H00000040)'dip 7
 		.AddFrame 205,244,190,"Replay Limit",&H00200000,Array("1 Per Ball",0,"1 Per Game",&H00200000)'dip 22
 		.AddChk 205,300,190,Array("Attract Mode Sound",&H00400000)'dip 23
 		.AddChk 205,315,190,Array("Background Sound",&H00002000)'dip 14
 		.AddLabel 50,340,300,15,"After hitting OK, press F3 to reset game with new settings."
 		.ViewDips
 	End With
 End Sub
 Set vpmShowDips = GetRef("editDips")
'**********************************************************************************************************
'**********************************************************************************************************
'**********************************************************************************************************
'**********************************************************************************************************

'**********Sling Shot Animations
' Rstep and Lstep  are the variables that increment the animation
'****************
Dim RStep, Lstep, LAstep, LBstep, LCstep, lkick, skick

Sub RightSlingShot_Slingshot
    PlaySound SoundFX("right_slingshot",DOFContactors), 0, 1, 0.05, 0.05
    vpmTimer.PulseSw 15
    RSling.Visible = 0
    RSling1.Visible = 1
    sling1.TransZ = -32
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
    PlaySound SoundFX("left_slingshot",DOFContactors),0,1,-0.05,0.05
    vpmTimer.PulseSw 16
    LSling.Visible = 0
    LSling1.Visible = 1
    sling2.TransZ = -28
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
    RubberRA.Visible = 0
    RubberRA1.Visible = 1
    LAstep = 0
    WallLA.TimerEnabled = 1
End Sub

Sub WallLA_Timer
    Select Case LAstep
        Case 3:RubberRA1.Visible = 0:RubberRA2.Visible = 1
        Case 4:RubberRA2.Visible = 0:RubberRA.Visible = 1:WallLA.TimerEnabled = 0:
    End Select
    LAstep = LAstep + 1
End Sub

Sub WallLB_Hit
    RubberRA.Visible = 0
    RubberRA3.Visible = 1
    LBstep = 0
    WallLB.TimerEnabled = 1
End Sub

Sub WallLB_Timer
    Select Case LAstep
        Case 3:RubberRA3.Visible = 0:RubberRA4.Visible = 1
        Case 4:RubberRA4.Visible = 0:RubberRA.Visible = 1:WallLB.TimerEnabled = 0:
    End Select
    LBstep = LBstep + 1
End Sub

Sub WallLC_Hit
    RubberRA.Visible = 0
    RubberRA5.Visible = 1
    LCstep = 0
    WallLC.TimerEnabled = 1
End Sub

Sub WallLC_Timer
    Select Case LCstep
        Case 3:RubberRA5.Visible = 0:RubberRA6.Visible = 1
        Case 4:RubberRA6.Visible = 0:RubberRA.Visible = 1:WallLC.TimerEnabled = 0:
    End Select
    LCstep = LCstep + 1
End Sub

Sub SolLeftOut(enabled)
    if enabled then
        playsound Soundfx("popper_ball",DOFContactors)
'        Controller.Switch(24) = false
        kicker.Kick  0, 18 + 5 * Rnd
		slingkick.rotx = 20
		lkick = 0
		slingkicktimer.enabled = 1
    end if
End Sub

Sub leftsaucersafetytimer_timer
	kicker.Kick 0, 18 + 5 * Rnd
'	playsound "bell"
	leftsaucersafetytimer.enabled=0
End Sub

Sub SlingKickTimer_Timer
    Select Case LKick
        Case 1:slingkick.rotx = 20
        Case 2:slingkick.rotx = 20
        Case 3:slingkick.rotx = 20
        Case 4:slingkick.rotx = 5
        Case 5:slingkick.rotx = 0
		slingkicktimer.enabled = 0
    End Select
    LKick = LKick + 1
End Sub

Sub SolSaucerOut(enabled)
    if enabled then
        playsound Soundfx("popper_ball",DOFContactors)
        playsound Soundfx("solenoid",DOFContactors)
        Controller.Switch(40) = 0
        sw40.Kick  100, 5 + 5 * Rnd
		pkickarm.rotz=2
		skick = 0
		topsaucertimer.enabled=1
    end if
End Sub

Sub topsaucertimer_Timer
    Select Case skick
        Case 1:Pkickarm.Rotz = 15
        Case 2:Pkickarm.Rotz = 15
        Case 3:Pkickarm.Rotz = 15
        Case 4:Pkickarm.Rotz = 8
        Case 5:Pkickarm.Rotz = 4
        Case 6:Pkickarm.Rotz = 2
        Case 7:Pkickarm.Rotz = 0:topsaucertimer.Enabled = 0
    End Select
    skick = skick + 1
End Sub

'*****************************************
'   rothbauerw's Manual Ball Control
'*****************************************

Dim BCup, BCdown, BCleft, BCright
Dim ControlBallInPlay, ControlActiveBall
Dim BCvel, BCyveloffset, BCboostmulti, BCboost

BCboost = 1				'Do Not Change - default setting
BCvel = 4				'Controls the speed of the ball movement
BCyveloffset = -0.01 	'Offsets the force of gravity to keep the ball from drifting vertically on the table, should be negative
BCboostmulti = 3		'Boost multiplier to ball veloctiy (toggled with the B key)

ControlBallInPlay = false

Sub StartBallControl_Hit()
	Set ControlActiveBall = ActiveBall
	ControlBallInPlay = true
End Sub

Sub StopBallControl_Hit()
	ControlBallInPlay = false
End Sub

Sub BallControlTimer_Timer()
	If EnableBallControl and ControlBallInPlay then
		If BCright = 1 Then
			ControlActiveBall.velx =  BCvel*BCboost
		ElseIf BCleft = 1 Then
			ControlActiveBall.velx = -BCvel*BCboost
		Else
			ControlActiveBall.velx = 0
		End If

		If BCup = 1 Then
			ControlActiveBall.vely = -BCvel*BCboost
		ElseIf BCdown = 1 Then
			ControlActiveBall.vely =  BCvel*BCboost
		Else
			ControlActiveBall.vely = bcyveloffset
		End If
	End If
End Sub

if flippershadows=1 then
	FlipperLSh.visible=1
	FlipperRSh.visible=1
else
	FlipperLSh.visible=0
	FlipperRSh.visible=0
end if

'*****************************************
'			FLIPPER SHADOWS
'*****************************************

sub FlipperTimer_Timer()
	FlipperLSh.RotZ = LeftFlipper.currentangle
	FlipperRSh.RotZ = RightFlipper.currentangle
	If sw4.isdropped=1 then drop1_S.image="blank" End If
	If sw4.isdropped=0 then drop1_S.image="drop1shadow" End If
	If sw11.isdropped=1 then drop2_S.image="blank" End If
	If sw11.isdropped=0 then drop2_S.image="drop2shadow" End If
	If sw10.isdropped=1 then drop3_S.image="blank" End If
	If sw10.isdropped=0 then drop3_S.image="drop3shadow" End If
	If sw9.isdropped=1 then drop4_S.image="blank" End If
	If sw9.isdropped=0 then drop4_S.image="drop4shadow" End If
    flipperleft_prim.rotz = leftflipper.currentangle
    flipperright_prim.rotz = rightflipper.currentangle
End Sub

if ballshadows=1 then
	BallShadowUpdate.enabled=1
else
	BallShadowUpdate.enabled=0
end if

if dullard = 1 Then
	dullardL.collidable = 1
	dullardR.collidable = 1
	cowardL.collidable = 0
	cowardR.collidable = 0
else
	dullardL.collidable = 0
	dullardR.collidable = 0
end If

if coward = 1 Then
	dullardL.collidable = 0
	dullardR.collidable = 0
	cowardL.collidable = 1
	cowardR.collidable = 1
Else
	cowardL.collidable = 0
	cowardR.collidable = 0
end if

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
	PlaySound "pinhit_low", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0, AudioFade(ActiveBall)
End Sub

Sub Targets_Hit (idx)
	PlaySound "target", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0, AudioFade(ActiveBall)
End Sub

Sub Metals_Thin_Hit (idx)
	PlaySound "metalhit_thin", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
End Sub

Sub Metals_Medium_Hit (idx)
	PlaySound "metalhit_medium", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
End Sub

Sub Metals2_Hit (idx)
	PlaySound "metalhit2", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
End Sub

Sub Gates_Hit (idx)
	PlaySound "gate4", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
End Sub

Sub aPlastic_Hit (idx)
	PlaySound "fx_rampbump6", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
End Sub

Sub aWood_Hit (idx)
	PlaySound "pinhit_low", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
End Sub

Sub Spinner_Spin
	PlaySound "fx_spinner",0,.25,0,0.25
End Sub

Sub Rubbers_Hit(idx)
 	dim finalspeed
  	finalspeed=SQR(activeball.velx * activeball.velx + activeball.vely * activeball.vely)
 	If finalspeed > 20 then
		PlaySound "fx_rubber2", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
	End if
	If finalspeed >= 6 AND finalspeed <= 20 then
 		RandomSoundRubber()
 	End If
End Sub

Sub Posts_Hit(idx)
 	dim finalspeed
  	finalspeed=SQR(activeball.velx * activeball.velx + activeball.vely * activeball.vely)
 	If finalspeed > 16 then
		PlaySound "fx_rubber2", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
	End if
	If finalspeed >= 6 AND finalspeed <= 16 then
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

Sub table1_Paused:Controller.Pause = 1:End Sub
Sub table1_unPaused:Controller.Pause = 0:End Sub

Sub table1_Exit:Controller.Stop:End Sub

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

' Thalamus : Exit in a clean and proper way
Sub Table1_exit()
  Controller.Pause = False
  Controller.Stop
End Sub


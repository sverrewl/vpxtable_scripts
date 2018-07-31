Option Explicit
Randomize

On Error Resume Next
ExecuteGlobal GetTextFile("controller.vbs")
If Err Then MsgBox "You need the controller.vbs in order to run this table, available in the vp10 package"
On Error Goto 0

' Thalamus 2018-07-24
' Added/Updated "Positional Sound Playback Functions" and "Supporting Ball & Sound Functions"
' Changed UseSolenoids=1 to 2
' No special SSF tweaks yet.

Const cGameName="spectru4",UseSolenoids=2,UseLamps=0,UseGI=0,SSolenoidOn="SolOn",SSolenoidOff="SolOff", SCoin="coin"

LoadVPM "01550000", "BALLY.VBS", 3.26
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
SolCallback(1) = "bsT.SolOut"     'Top Kicker Left Kick
SolCallback(2) = "bsT.SolOutAlt"  'Top Kicker Right kick
SolCallback(3) = "bsL2.SolOut"
SolCallback(4) = "bsR2.SolOut"
SolCallback(5) = "bsD.SolOut"
SolCallback(6) = "vpmSolSound SoundFX(""Knocker"",DOFKnocker),"
SolCallback(8) = "bsL1.SolOut"
SolCallback(9) = "bsR1.SolOut"
SolCallback(10) = "dtB.SolDropUp Not"
SolCallback(11) = "dtG.SolDropUp Not"
SolCallback(12) = "dtY.SolDropUp Not"
SolCallback(13) = "dtR.SolDropUp Not"
SolCallback(19) = "vpmNudge.SolGameOn"

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
'**********************************************************************************************************

'Solenoid Controlled toys
'**********************************************************************************************************


'*****GI Lights On
dim xx
For each xx in GI:xx.State = 1: Next

'**********************************************************************************************************

'Initiate Table
'**********************************************************************************************************

Dim bsD, bsL1, bsL2, bsR1, bsR2, bsT, dtR, dtY, dtG, dtB

Sub Table1_Init
	vpmInit Me
	On Error Resume Next
		With Controller
		.GameName = cGameName
		If Err Then MsgBox "Can't start Game" & cGameName & vbNewLine & Err.Description : Exit Sub
		.SplashInfoLine = "Spectrum - Bally 1981"&chr(13)&"You Suck"
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

    vpmNudge.TiltSwitch = 15
    vpmNudge.Sensitivity = 3
    vpmNudge.Tiltobj = Array()

    ' Drain
    Set bsD = New cvpmBallStack
        bsD.InitSaucer Drain, 8, 356, 14
        bsD.InitExitSnd SoundFX("Popper",DOFContactors), SoundFX("Solenoid",DOFContactors)
        bsD.Balls = 1
        bsD.CreateEvents "bsD", Drain

    ' Left Eject Hole 1
    Set bsL1 = New cvpmBallStack
        bsL1.InitSaucer sw17, 17, 135, 7
        bsL1.InitExitSnd SoundFX("Popper",DOFContactors), SoundFX("Solenoid",DOFContactors)
        bsL1.Balls = 1
        bsL1.CreateEvents "bsL1", sw17

    ' Left Eject Hole 2
    Set bsL2 = New cvpmBallStack
        bsL2.InitSaucer sw41, 41, 259+rnd(1)*3, 6+rnd(1)*3
        bsL2.InitExitSnd SoundFX("Popper",DOFContactors), SoundFX("Solenoid",DOFContactors)
        bsL2.CreateEvents "bsL2", sw41

    ' Right Eject Hole 1
    Set bsR1 = New cvpmBallStack
        bsR1.InitSaucer sw24, 24, 224+rnd(1)*3, 6+rnd(1)*3
        bsR1.InitExitSnd SoundFX("Popper",DOFContactors), SoundFX("Solenoid",DOFContactors)
        bsR1.Balls = 1
        bsR1.CreateEvents "bsR1", sw24

    ' Right Eject Hole 2
    Set bsR2 = New cvpmBallStack
        bsR2.InitSaucer sw48, 48, 99+rnd(1)*3, 6+rnd(1)*3
        bsR2.InitExitSnd SoundFX("Popper",DOFContactors), SoundFX("Solenoid",DOFContactors)
        bsR2.CreateEvents "bsR2", sw48

    ' Top Eject Hole
    Set bsT = New cvpmBallStack
        bsT.InitSaucer sw7, 7, 269+rnd(1)*3, 11+rnd(1)*3 'Left Kick
        bsT.InitAltKick 90, 12         'Right Kick
        bsT.InitExitSnd SoundFX("Popper",DOFContactors), SoundFX("Solenoid",DOFContactors)

    ' Yellow Targets
    Set dtY = New cvpmDropTarget
        dtY.InitDrop Array(sw37, sw38, sw39), Array(37, 38, 39)
        dtY.InitSnd SoundFX("DTDrop",DOFContactors),SoundFX("DTReset",DOFContactors)

    ' Green Targets
    Set dtG = New cvpmDropTarget
        dtG.InitDrop Array(sw34, sw35, sw36), Array(34, 35, 36)
        dtG.InitSnd SoundFX("DTDrop",DOFContactors),SoundFX("DTReset",DOFContactors)

    ' Blue Targets
    Set dtB = New cvpmDropTarget
        dtB.InitDrop Array(sw42, sw43, sw44), Array(42, 43, 44)
        dtB.InitSnd SoundFX("DTDrop",DOFContactors),SoundFX("DTReset",DOFContactors)

    ' Red Targets
    Set dtR = New cvpmDropTarget
        dtR.InitDrop Array(sw45,sw46,sw47), Array(45,46,47)
        dtR.InitSnd SoundFX("DTDrop",DOFContactors),SoundFX("DTReset",DOFContactors)

End Sub

'**********************************************************************************************************
'Plunger code
'**********************************************************************************************************

Sub Table1_KeyDown(ByVal KeyCode)
	If KeyDownHandler(keycode) Then Exit Sub
    If keycode = RightFlipperKey Then Controller.Switch(1) = False
    If keycode = LeftFlipperKey Then Controller.Switch(2) = False
End Sub

Sub Table1_KeyUp(ByVal KeyCode)
	If KeyUpHandler(keycode) Then Exit Sub
    If keycode = RightFlipperKey Then Controller.Switch(1) = True
    If keycode = LeftFlipperKey Then Controller.Switch(2) = True
End Sub

'**********************************************************************************************************

' Holes
Sub sw7_Hit:bsT.AddBall 0:End Sub

'Star Triggers
Sub sw19_Hit:Controller.Switch(19) = 1 : playsound"rollover" : End Sub
Sub sw19_UnHit:Controller.Switch(19) = 0:End Sub
Sub sw19a_Hit:Controller.Switch(19) = 1 : playsound"rollover" : End Sub
Sub sw19a_UnHit:Controller.Switch(19) = 0:End Sub
Sub sw19b_Hit:Controller.Switch(19) = 1 : playsound"rollover" : End Sub
Sub sw19b_UnHit:Controller.Switch(19) = 0:End Sub
Sub sw20_Hit:Controller.Switch(20) = 1 : playsound"rollover" : End Sub
Sub sw20_UnHit:Controller.Switch(20) = 0:End Sub
Sub sw21_Hit:Controller.Switch(21) = 1 : playsound"rollover" : End Sub
Sub sw21_UnHit:Controller.Switch(21) = 0:End Sub
Sub sw21a_Hit:Controller.Switch(21) = 1 : playsound"rollover" : End Sub
Sub sw21a_UnHit:Controller.Switch(21) = 0:End Sub
Sub sw21b_Hit:Controller.Switch(21) = 1 : playsound"rollover" : End Sub
Sub sw21b_UnHit:Controller.Switch(21) = 0:End Sub
Sub sw27_Hit:Controller.Switch(27) = 1 : playsound"rollover" : End Sub
Sub sw27_UnHit:Controller.Switch(27) = 0:End Sub
Sub sw28_Hit:Controller.Switch(28) = 1 : playsound"rollover" : End Sub
Sub sw28_UnHit:Controller.Switch(28) = 0:End Sub
Sub sw29_Hit:Controller.Switch(29) = 1 : playsound"rollover" : End Sub
Sub sw29_UnHit:Controller.Switch(29) = 0:End Sub
Sub sw30_Hit:Controller.Switch(30) = 1 : playsound"rollover" : End Sub
Sub sw30_UnHit:Controller.Switch(30) = 0:End Sub
Sub sw33_Hit:Controller.Switch(33) = 1 : playsound"rollover" : End Sub
Sub sw33_UnHit:Controller.Switch(33) = 0:End Sub
Sub sw33a_Hit:Controller.Switch(33) = 1 : playsound"rollover" : End Sub
Sub sw33a_UnHit:Controller.Switch(33) = 0:End Sub
Sub sw40_Hit:Controller.Switch(40) = 1 : playsound"rollover" : End Sub
Sub sw40_UnHit:Controller.Switch(40) = 0:End Sub
Sub sw40a_Hit:Controller.Switch(40) = 1 : playsound"rollover" : End Sub
Sub sw40a_UnHit:Controller.Switch(40) = 0:End Sub

' Rollovers
Sub sw18_Hit:Controller.Switch(18) = 1 : playsound"rollover" : End Sub
Sub sw18_UnHit:Controller.Switch(18) = 0:End Sub
Sub sw22_Hit:Controller.Switch(22) = 1 : playsound"rollover" : End Sub
Sub sw22_UnHit:Controller.Switch(22) = 0:End Sub
Sub sw23_Hit:Controller.Switch(23) = 1 : playsound"rollover" : End Sub
Sub sw23_UnHit:Controller.Switch(23) = 0:End Sub
Sub sw25a_Hit:Controller.Switch(25) = 1 : playsound"rollover" : End Sub
Sub sw25a_UnHit:Controller.Switch(25) = 0:End Sub
Sub sw25b_Hit:Controller.Switch(25) = 1 : playsound"rollover" : End Sub
Sub sw25b_UnHit:Controller.Switch(25) = 0:End Sub
Sub sw25c_Hit:Controller.Switch(25) = 1 : playsound"rollover" : End Sub
Sub sw25c_UnHit:Controller.Switch(25) = 0:End Sub
Sub sw32a_Hit:Controller.Switch(32) = 1 : playsound"rollover" : End Sub
Sub sw32a_UnHit:Controller.Switch(32) = 0:End Sub
Sub sw32b_Hit:Controller.Switch(32) = 1 : playsound"rollover" : End Sub
Sub sw32b_UnHit:Controller.Switch(32) = 0:End Sub
Sub sw32c_Hit:Controller.Switch(32) = 1 : playsound"rollover" : End Sub
Sub sw32c_UnHit:Controller.Switch(32) = 0:End Sub

' Spinners
Sub sw26_Spin:vpmTimer.PulseSw 26 : playsound"fx_spinner" : End Sub
Sub sw31_Spin:vpmTimer.PulseSw 31 : playsound"fx_spinner" : End Sub

'Drop Targets
 Sub Sw37_Dropped:dtY.Hit 1 :End Sub
 Sub Sw38_Dropped:dtY.Hit 2 :End Sub
 Sub Sw39_Dropped:dtY.Hit 3 :End Sub

 Sub Sw34_Dropped:dtG.Hit 1 :End Sub
 Sub Sw35_Dropped:dtG.Hit 2 :End Sub
 Sub Sw36_Dropped:dtG.Hit 3 :End Sub

 Sub Sw42_Dropped:dtB.Hit 1 :End Sub
 Sub Sw43_Dropped:dtB.Hit 2 :End Sub
 Sub Sw44_Dropped:dtB.Hit 3 :End Sub

 Sub Sw45_Dropped:dtR.Hit 1 :End Sub
 Sub Sw46_Dropped:dtR.Hit 2 :End Sub
 Sub Sw47_Dropped:dtR.Hit 3 :End Sub


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
    NFadeL 1, l1
    NFadeL 2, l2
    NFadeL 3, l3
    NFadeLm 4, l4a
    NFadeL 4, l4
    NFadeL 5, l5
    NFadeL 6, l6
    NFadeL 7, l7
    NFadeL 8, l8
    NFadeL 9, l9
    NFadeL 10, l10
    'FadeReel 13, l13   'Balls to Play
    NFadeL 17, l17
    NFadeL 18, l18
    NFadeL 19, l19
    NFadeLm 20, l20
    NFadeL 20, l20a
    NFadeL 21, l21
    NFadeL 22, l22
    NFadeL 23, l23
    NFadeL 24, l24
    NFadeL 25, l25
    NFadeL 26, l26
    'FadeReel 27, l27 'Match
    'FadeReel 29, l29 'Highscore
    NFadeL 33, l33
    NFadeL 34, l34
    NFadeL 35, l35
    NFadeLm 36, l36a
    NFadeL 36, l36
    NFadeL 37, l37
    NFadeL 38, l38
    NFadeL 39, l39
    NFadeL 40, l40
    NFadeL 41, l41
    NFadeLm 42, l42a
    NFadeL 42, l42
    'FadeReel 45, l45 'GAME OVER
    NFadeL 49, l49
    NFadeL 50, l50
    NFadeL 51, l51
    NFadeLm 52, l52
    NFadeL 52, l52a
    NFadeL 53, l53
    NFadeL 54, l54
    NFadeL 55, l55
    NFadeL 56, l56
    NFadeL 57, l57
    NFadeL 58, l58
    NFadeLm 59, l59a
    NFadeL 59, l59
    'FadeReel 61, l61 'TILT
    NFadeL 65, l65
    NFadeL 66, l66
    NFadeL 67, l67
    NFadeL 68, l68
    NFadeL 69, l69
    NFadeL 70, l70
    NFadeL 71, l71
    NFadeL 72, l72
    NFadeL 73, l73
    NFadeL 74, l74
    NFadeL 75, l75
    NFadeL 76, l76
    NFadeL 81, l81
    NFadeL 82, l82
    NFadeL 83, l83
    NFadeL 84, l84
    NFadeL 85, l85
    NFadeL 86, l86
    NFadeL 87, l87
    NFadeL 88, l88
    NFadeL 89, l89
    NFadeL 90, l90
    NFadeL 91, l91
    NFadeL 92, l92
    NFadeL 97, l97
    NFadeL 98, l98
    NFadeL 99, l99
    NFadeL 100, l100
    NFadeL 101, l101
    NFadeL 102, l102
    NFadeL 103, l103
    NFadeL 104, l104
    NFadeL 105, l105
    NFadeL 106, l106
    NFadeL 107, l107
    NFadeL 108, l108
    NFadeL 113, l113
    NFadeL 114, l114
    NFadeL 115, l115
    NFadeL 116, l116
    NFadeL 117, l117
    NFadeL 118, l118
    NFadeL 119, l119
    NFadeL 120, l120
    NFadeL 121, l121
    NFadeL 122, l122
    NFadeL 123, l123
    NFadeL 124, l124

    'extra lights
    NFadeL 150, l150
    NFadeL 151, l151
    NFadeL 152, l152
    NFadeL 153, l153
    NFadeL 154, l154
    NFadeL 155, l155
    NFadeL 156, l156
    NFadeL 157, l157
    NFadeL 158, l158
    NFadeL 159, l159
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




'Bally Spectrum
'added by Inkochnito
Sub editDips
    Dim vpmDips:Set vpmDips = New cvpmDips
    With vpmDips
        .AddForm 700, 400, "Spectrum - DIP switches"
        .AddFrame 2, 0, 190, "Maximum credits", &H03000000, Array("10 credits", 0, "15 credits", &H01000000, "25 credits", &H02000000, "40 credits", &H03000000)         'dip 25&26
        .AddFrame 2, 74, 190, "Multiplier will step when making", &H00000080, Array("4 green buttons from left to right", 0, "4 green buttons in any order", &H00000080) 'dip 8
        .AddFrame 2, 120, 190, "Outlane special will lite after", &H00200000, Array("making 4 stars", 0, "making 3 stars", &H00200000)                                   'dip 22
        .AddFrame 205, 0, 190, "Balls per game", &HC0000000, Array("2 balls", &HC0000000, "3 balls", 0, "4 balls", &H80000000, "5 balls", &H40000000)                    'dip 31&32
        .AddFrame 205, 74, 190, "4 circling lites will step down after", &H00800000, Array("hitting 3 same targets", 0, "hitting 4 same targets", &H00800000)            'dip 24
        .AddFrame 205, 120, 190, "Replay limit", &H10000000, Array("1 replay per game", 0, "unlimited", &H10000000)                                                      'dip 29
        .AddChk 2, 170, 120, Array("Attract sound", &H20000000)                                                                                                          'dip 30
        .AddChk 2, 185, 240, Array("Any color drop target lites out memory", &H00004000)                                                                                 'dip 15
        .AddChk 2, 200, 240, Array("Any color drop target flashing lites memory", 32768)                                                                                 'dip 16
        .AddChk 2, 215, 190, Array("Left && right spinners memory", &H00000020)                                                                                          'dip 6
        .AddChk 2, 230, 398, Array("Hitting 2 or 3 same targets to step 4 circling computer lites down 1 step", &H00400000)                                              'dip 23
        .AddChk 250, 170, 140, Array("Outlane special memory", &H00002000)                                                                                               'dip 14
        .AddChk 250, 185, 120, Array("Match feature", &H08000000)                                                                                                        'dip 28
        .AddChk 250, 200, 120, Array("Credits displayed", &H04000000)                                                                                                    'dip 27
        .AddChk 250, 215, 150, Array("2X, 3X, 4X lite memory", &H00000040)                                                                                               'dip 7
        .AddLabel 50, 260, 320, 20, "Set selftest position 16,17,18 and 19 to 03 for the best gameplay."
        .AddLabel 50, 280, 300, 20, "After hitting OK, press F3 to reset game with new settings."
        .ViewDips
    End With
End Sub
Set vpmShowDips = GetRef("editDips")

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


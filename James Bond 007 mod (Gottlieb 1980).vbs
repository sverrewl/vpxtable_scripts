Option Explicit
Randomize

On Error Resume Next
ExecuteGlobal GetTextFile("controller.vbs")
If Err Then MsgBox "You need the controller.vbs in order to run this table, available in the vp10 package"
On Error Goto 0

' Thalamus 2018-07-24
' Table has its own "Positional Sound Playback Functions" and "Supporting Ball & Sound Functions"
' Changed UseSolenoids=1 to 2
' No special SSF tweaks yet.


Const cGameName="jamesb" 'Timer based Rom
'Const cGameName="jamesb2" ' 3 Ball based rom

Const UseSolenoids=2,UseLamps=1,UseGI=0,SSolenoidOn="SolOn",SSolenoidOff="SolOff", SCoin="coin"

LoadVPM "01120100","sys80.vbs",3.02

Dim DesktopMode: DesktopMode = Table1.ShowDT
If DesktopMode = True Then 'Show Desktop components
Ramp16.visible=1
Ramp15.visible=1
Primitive13.visible=1
Primitive132.visible=0
Else
Ramp16.visible=0
Ramp15.visible=0
Primitive13.visible=0
Primitive132.visible=1
End if

'*************************************************************
'Solenoid Call backs
'**********************************************************************************************************
SolCallback(2)="dtLeft.SolDropUp"
SolCallback(5)="dtTop.SolDropUp"
SolCallback(6)="dtRight.SolDropUp"
SolCallback(8)="vpmSolSound SoundFX(""Knocker"",DOFKnocker),"
SolCallback(9)="bsTrough.SolOut"

SolCallback(sLRFlipper) = "SolRFlipper"
SolCallback(sLLFlipper) = "SolLFlipper"

Sub SolLFlipper(Enabled)
     If Enabled Then
         PlaySound SoundFX("fx_Flipperup",DOFContactors):LeftFlipper.RotateToEnd:LeftFlipper1.RotateToEnd
     Else
         PlaySound SoundFX("fx_Flipperdown",DOFContactors):LeftFlipper.RotateToStart:LeftFlipper1.RotateToStart
     End If
  End Sub
  
Sub SolRFlipper(Enabled)
     If Enabled Then
         PlaySound SoundFX("fx_Flipperup",DOFContactors):RightFlipper.RotateToEnd:RightFlipper1.RotateToEnd
     Else
         PlaySound SoundFX("fx_Flipperdown",DOFContactors):RightFlipper.RotateToStart:RightFlipper1.RotateToStart
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

Dim dtTop,dtLeft,dtRight,bsTrough

Sub Table1_Init
	vpmInit Me
	On Error Resume Next
		With Controller
		.GameName = cGameName
		If Err Then MsgBox "Can't start Game" & cGameName & vbNewLine & Err.Description : Exit Sub
		.SplashInfoLine = "Memory Lane -- Stern, 1978"&chr(13)&"You Suck"
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
	vpmNudge.TiltSwitch=57
	vpmNudge.Sensitivity=2
	vpmNudge.TiltObj=Array(Bumper1,Bumper2,Bumper3,LeftSlingshot,RightSlingshot)

	Set bsTrough=New cvpmBallstack
		bsTrough.InitNoTrough ballrelease,67,90,3
		bsTrough.InitExitSnd SoundFX("ballrelease",DOFContactors), SoundFX("Solenoid",DOFContactors)

	Set dtTop=New cvpmDropTarget
		dtTop.InitDrop Array(SW03,SW13,SW23,SW33),Array(3,13,23,33)
		dtTop.initsnd SoundFX("DTDrop",DOFContactors),SoundFX("DTReset",DOFContactors)
	
	Set dtLeft=New cvpmDropTarget
		dtLeft.InitDrop Array(SW02,SW12,SW42,SW22,SW32),Array(2,12,42,22,32)
		dtLeft.initsnd SoundFX("DTDrop",DOFContactors),SoundFX("DTReset",DOFContactors)
	
	Set dtRight=New cvpmDropTarget
		dtRight.InitDrop Array(SW04,SW14,SW24,SW34),Array(4,14,24,34)
		dtRight.initsnd SoundFX("DTDrop",DOFContactors),SoundFX("DTReset",DOFContactors)

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

'drop targets
Sub sw03_Dropped:dtTop.Hit 1:End Sub
Sub sw13_Dropped:dtTop.Hit 2:End Sub
Sub sw23_Dropped:dtTop.Hit 3:End Sub
Sub sw33_Dropped:dtTop.Hit 4:End Sub

Sub sw02_Dropped:dtLeft.Hit 1:End Sub
Sub sw12_Dropped:dtLeft.Hit 2:End Sub
Sub sw42_Dropped:dtLeft.Hit 3:End Sub
Sub sw22_Dropped:dtLeft.Hit 4:End Sub
Sub sw32_Dropped:dtLeft.Hit 5:End Sub

Sub sw04_Dropped:dtRight.Hit 1:End Sub
Sub sw14_Dropped:dtRight.Hit 2:End Sub
Sub sw24_Dropped:dtRight.Hit 3:End Sub
Sub sw34_Dropped:dtRight.Hit 4:End Sub

'rollovers
Sub sw43a_Hit:Controller.Switch(43)=1 : playsound"rollover" : End Sub
Sub sw43a_UnHit:Controller.Switch(43)=0:End Sub
Sub sw43b_Hit:Controller.Switch(43)=1 : playsound"rollover" : End Sub
Sub sw43b_unHit:Controller.Switch(43)=0:End Sub
Sub sw53a_Hit:Controller.Switch(53)=1 : playsound"rollover" : End Sub
Sub sw53a_UnHit:Controller.Switch(53)=0:End Sub
Sub sw53b_Hit:Controller.Switch(53)=1 : playsound"rollover" : End Sub
Sub sw53b_UnHit:Controller.Switch(53)=0:End Sub
Sub sw53c_Hit:Controller.Switch(53)=1 : playsound"rollover" : End Sub
Sub sw53c_unHit:Controller.Switch(53)=0:End Sub
Sub sw72_Hit:Controller.Switch(72)=1 : playsound"rollover" : End Sub
Sub sw72_unHit:Controller.Switch(72)=0:End Sub
Sub sw73_Hit:Controller.Switch(73)=1 : playsound"rollover" : End Sub
Sub sw73_unHit:Controller.Switch(73)=0:End Sub
Sub sw74_Hit:Controller.Switch(74)=1 : playsound"rollover" : End Sub
Sub sw74_unHit:Controller.Switch(74)=0:End Sub

'Bumpers
Sub Bumper1_Hit : vpmTimer.PulseSw(44) : playsound SoundFX("fx_bumper1",DOFContactors): End Sub
Sub Bumper2_Hit : vpmTimer.PulseSw(44) : playsound SoundFX("fx_bumper1",DOFContactors): End Sub
Sub Bumper3_Hit : vpmTimer.PulseSw(44) : playsound SoundFX("fx_bumper1",DOFContactors): End Sub

'StandUp Targets
Sub sw52_Hit:VpmTimer.PulseSw 52:End Sub
Sub sw62_Hit:VpmTimer.PulseSw 62:End Sub
Sub sw64_Hit:VpmTimer.PulseSw 64:End Sub

 'Scoring Rubber
Sub sw63a_hit:vpmTimer.pulseSw 63 : playsound"flip_hit_3" : End Sub 
Sub sw63b_hit:vpmTimer.pulseSw 63 : playsound"flip_hit_3" : End Sub 
Sub sw63c_hit:vpmTimer.pulseSw 63 : playsound"flip_hit_3" : End Sub 
Sub sw63d_hit:vpmTimer.pulseSw 63 : playsound"flip_hit_3" : End Sub 
Sub sw63e_hit:vpmTimer.pulseSw 63 : playsound"flip_hit_3" : End Sub 
Sub sw63f_hit:vpmTimer.pulseSw 63 : playsound"flip_hit_3" : End Sub 
Sub sw63g_hit:vpmTimer.pulseSw 63 : playsound"flip_hit_3" : End Sub 
Sub sw63h_hit:vpmTimer.pulseSw 63 : playsound"flip_hit_3" : End Sub 

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

			'Special Handling DT hit whith Light ID pulse

			If (chgLamp(ii,0) = 12 And chgLamp(ii,1) = 1) Then dtLeft.Hit 1
			If (chgLamp(ii,0) = 13 And chgLamp(ii,1) = 1) Then dtLeft.Hit 2
			'Purple DT is dtLeft.Hit 3
			If (chgLamp(ii,0) = 14 And chgLamp(ii,1) = 1) Then dtLeft.Hit 4
			If (chgLamp(ii,0) = 15 And chgLamp(ii,1) = 1) Then dtLeft.Hit 5

			If (chgLamp(ii,0) = 16 And chgLamp(ii,1) = 1) Then dtTop.Hit 1
			If (chgLamp(ii,0) = 17 And chgLamp(ii,1) = 1) Then dtTop.Hit 2
			If (chgLamp(ii,0) = 18 And chgLamp(ii,1) = 1) Then dtTop.Hit 3
			If (chgLamp(ii,0) = 19 And chgLamp(ii,1) = 1) Then dtTop.Hit 4

			If (chgLamp(ii,0) = 20 And chgLamp(ii,1) = 1) Then dtRight.Hit 1
			If (chgLamp(ii,0) = 21 And chgLamp(ii,1) = 1) Then dtRight.Hit 2
			If (chgLamp(ii,0) = 22 And chgLamp(ii,1) = 1) Then dtRight.Hit 3
			If (chgLamp(ii,0) = 23 And chgLamp(ii,1) = 1) Then dtRight.Hit 4

        Next
    End If
    UpdateLamps
End Sub


Sub UpdateLamps()

'NFadeL 0, L0 'Game Over relay
'NFadeL 1, L1 'Tilt
'NFadeL 10, L10 'High Score to date
'NFadeL 11, L11 'Game Over
'NFadeL 12, L12 '1 Yellow DT
'NFadeL 13, L13 '2 Yellow DT
'NFadeL 14, L14 '4 Yellow DT
'NFadeL 15, L15 '5 Yellow DT
'NFadeL 16, L16 '1 Green DT
'NFadeL 17, L17 '2 Green DT
'NFadeL 18, L18 '3 Green DT
'NFadeL 19, L19 '4 Green DT
'NFadeL 20, L20 '1 Red DT
'NFadeL 21, L21 '2 Red DT
'NFadeL 22, L22 '3 Red DT
'NFadeL 23, L23 '4 Red DT
NFadeL 24, L24
NFadeLm 26, L26
NFadeL 26, L26b
NFadeLm 27, L27
NFadeL 27, L27b
NFadeLm 28, L28
NFadeL 28, L28b
NFadeLm 29, L29
NFadeL 29, L29b
NFadeLm 30, L30
NFadeL 30, L30b
NFadeL 31, L31
NFadeL 33, L33
NFadeL 34, L34
NFadeL 35, L35
NFadeL 36, L36
NFadeL 37, L37
NFadeL 38, L38
NFadeL 39, L39
NFadeL 40, L40
NFadeL 41, L41
NFadeL 42, L42
NFadeL 43, L43
NFadeL 44, L44
NFadeL 45, L45
NFadeL 46, L46
NFadeL 47, L47
NFadeLm 48, L48
NFadeL 48, L48a
NFadeLm 49, L49
NFadeLm 49, L49a
NFadeL 49, L49b

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


' Modulated Flasher and Lights objects
Sub SetLampMod(nr, value)
    If value > 0 Then
        LampState(nr) = 1
    Else
        LampState(nr) = 0
    End If
    FadingLevel(nr) = value
End Sub
 
Sub LampMod(nr, object)
    If TypeName(object) = "Light" Then
        Object.IntensityScale = FadingLevel(nr)/128
        Object.State = LampState(nr)
    End If
    If TypeName(object) = "Flasher" Then
        Object.IntensityScale = FadingLevel(nr)/128
        Object.visible = LampState(nr)
    End If
    If TypeName(object) = "Primitive" Then
        Object.DisableLighting = LampState(nr)
    End If
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
' Backglass Light Displays (7 digit 7 segment displays)
'**********************************************************************************************************

Dim Digits(34)
Digits(0)=Array(a00,a01,a02,a03,a04,a05,a06,n,a08)
Digits(1)=Array(a10,a11,a12,a13,a14,a15,a16,n,a18)
Digits(2)=Array(a20,a21,a22,a23,a24,a25,a26,n,a28)
Digits(3)=Array(a30,a31,a32,a33,a34,a35,a36,n,a38)
Digits(4)=Array(a40,a41,a42,a43,a44,a45,a46,n,a48)
Digits(5)=Array(a50,a51,a52,a53,a54,a55,a56,n,a58)

Digits(6)=Array(b00,b01,b02,b03,b04,b05,b06,n,b08)
Digits(7)=Array(b10,b11,b12,b13,b14,b15,b16,n,b18)
Digits(8)=Array(b20,b21,b22,b23,b24,b25,b26,n,b28)
Digits(9)=Array(b30,b31,b32,b33,b34,b35,b36,n,b38)
Digits(10)=Array(b40,b41,b42,b43,b44,b45,b46,n,b48)
Digits(11)=Array(b50,b51,b52,b53,b54,b55,b56,n,b58)

Digits(12)=Array(c00,c01,c02,c03,c04,c05,c06,n,c08)
Digits(13)=Array(c10,c11,c12,c13,c14,c15,c16,n,c18)
Digits(14)=Array(c20,c21,c22,c23,c24,c25,c26,n,c28)
Digits(15)=Array(c30,c31,c32,c33,c34,c35,c36,n,c38)
Digits(16)=Array(c40,c41,c42,c43,c44,c45,c46,n,c48)
Digits(17)=Array(c50,c51,c52,c53,c54,c55,c56,n,c58)

Digits(18)=Array(d00,d01,d02,d03,d04,d05,d06,n,d08)
Digits(19)=Array(d10,d11,d12,d13,d14,d15,d16,n,d18)
Digits(20)=Array(d20,d21,d22,d23,d24,d25,d26,n,d28)
Digits(21)=Array(d30,d31,d32,d33,d34,d35,d36,n,d38)
Digits(22)=Array(d40,d41,d42,d43,d44,d45,d46,n,d48)
Digits(23)=Array(d50,d51,d52,d53,d54,d55,d56,n,d58)

' Apron credit
Digits(24)=Array(D311,D312,D313,D314,D315,D316,D317)		'24
Digits(25)=Array(D321,D322,D323,D324,D325,D326,D327)		'25
Digits(26)=Array(D331,D332,D333,D334,D335,D336,D337)		'26
Digits(27)=Array(D341,D342,D343,D344,D345,D346,D347)		'27

Digits(28)=Array(e00,e01,e02,e03,e04,e05,e06,n,e08)
Digits(29)=Array(e10,e11,e12,e13,e14,e15,e16,n,e18)
Digits(30)=Array(f00,f01,f02,f03,f04,f05,f06,n,f08)
Digits(31)=Array(f10,f11,f12,f13,f14,f15,f16,n,f18)
Digits(32)=Array(e40,e41,e42,e43,e44,e45,e46,n,e48)
Digits(33)=Array(e50,e51,e52,e53,e54,e55,e56,n,e58)

Sub DisplayTimer_Timer
	Dim ChgLED,ii,num,chg,stat,obj
	ChgLed = Controller.ChangedLEDs(&Hffffffff, &Hffffffff)
If Not IsEmpty(ChgLED) Then

		For ii = 0 To UBound(chgLED)
			num = chgLED(ii, 0) : chg = chgLED(ii, 1) : stat = chgLED(ii, 2)
			if (num < 34) then
				For Each obj In Digits(num)
					If chg And 1 Then obj.State = stat And 1 
					chg = chg\2 : stat = stat\2
				Next
			else
				'if char(stat) > "" then msg(num) = char(stat)
			end if
		next
		end if

End Sub

'if Full Screen turn off Backglas Components
If DesktopMode = True Then
dim xxx
For each xxx in BG:xxx.Visible = 1: Next
else
For each xxx in BG:xxx.Visible = 0: Next
End if

'**********************************************************************************************************
'**********************************************************************************************************

'Gottlieb James Bond
'added by Inkochnito
Sub editDips
	Dim vpmDips : Set vpmDips = New cvpmDips
	With vpmDips
		.AddForm 700,400,"James Bond - DIP switches"
		.AddChk 2,10,190,Array("Match feature",&H00020000)'dip 18
		.AddChk 2,25,190,Array("Credits displayed",&H08000000)'dip 28
		.AddChk 2,40,190,Array("Attract features",&H20000000)'dip 30
		.AddFrame 2,60,190,"Maximum credits",49152,Array("8 credits",0,"10 credits",32768,"15 credits",&H00004000,"25 credits",49152)'dip 15&16
		.AddFrame 2,136,190,"Coin chute 1 and 2 control",&H00002000,Array("seperate",0,"same",&H00002000)'dip 14
		.AddFrame 2,182,190,"Playfield special",&H00200000,Array("awards replay",0,"awards extra time (20 units)",&H00200000)'dip 22
		.AddFrame 2,228,190,"Extra time units award",&H80000000,Array("only 1 award per player",0,"multiple awards posible",&H80000000)'dip 32
		.AddFrameExtra 2,274,190,"Attract sound",&H0002,Array("off",0,"play tune every 6 minutes",&H0002)'S-board dip 2
		.AddFrameExtra 2,320,190,"Sounds",&H0001,Array("continuous sound",0,"scoring sounds only",&H0001)'S-board dip 1
		.AddChk 205,10,190,Array("Coin switch tune",&H04000000)'dip 27
		.AddChk 205,25,190,Array("Credit button tune",&H02000000)'dip 26
		.AddChk 205,40,190,Array("Sound when scoring",&H01000000)'dip 25
		.AddFrame 205,60,190,"High game to date awards",&H00C00000,Array("not displayed and no award",0,"displayed and no award",&H00800000,"displayed and 2 replays",&H00400000,"displayed and 3 replays",&H00C00000)'dip 23&24
		.AddFrame 205,136,190,"Replay limit",&H00040000,Array("no replay limit",0,"one per game",&H00040000)'dip 19
		.AddFrame 205,182,190,"Novelty mode",&H00080000,Array("normal mode",0,"special scores 50K",&H00080000)'dip 20
		.AddFrame 205,228,190,"Game mode",&H00100000,Array("replay",0,"extra time",&H00100000)'dip 21
		.AddFrame 205,274,190,"3rd coin chute credits control",&H00001000,Array("no effect",0,"add 9",&H00001000)'dip 13
		.AddFrame 205,320,190,"Tilt penalty",&H10000000,Array("game over",0,"ball in play and loss of 5 time units",&H10000000)'dip 29
		.AddLabel 105,370,190,20,"not used in time based version"
		.AddFrame 105,386,190,"Balls per game",&H00010000,Array("5 balls",0,"3 balls",&H00010000)'dip 17
		.AddLabel 50,445,300,20,"After hitting OK, press F3 to reset game with new settings."
	End With
	Dim extra
	extra = Controller.Dip(4) + Controller.Dip(5)*256
	extra = vpmDips.ViewDipsExtra(extra)
	Controller.Dip(4) = extra And 255
	Controller.Dip(5) = (extra And 65280)\256 And 255
End Sub
Set vpmShowDips = GetRef("editDips")


'**********************************************************************************************************
'**********************************************************************************************************


'**********************************************************************************************************
'**********************************************************************************************************

'**********Sling Shot Animations
' Rstep and Lstep  are the variables that increment the animation
'****************
Dim RStep, Lstep

Sub RightSlingShot_Slingshot
	vpmTimer.PulseSw 63
    PlaySound SoundFX("right_slingshot",DOFContactors), 0,1, 0.05,0.05 '0,1, AudioPan(RightSlingShot), 0.05,0,0,1,AudioFade(RightSlingShot)
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
	vpmTimer.PulseSw 63
    PlaySound SoundFX("left_slingshot",DOFContactors), 0,1, -0.05,0.05 '0,1, AudioPan(LeftSlingShot), 0.05,0,0,1,AudioFade(LeftSlingShot)
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
        If BallVel(BOT(b) ) > 1 AND BOT(b).z < 30 Then
            rolling(b) = True
            PlaySound("fx_ballrolling" & b), -1, Vol(BOT(b)), AudioPan(BOT(b)), 0, Pitch(BOT(b)), 1, 0, AudioFade(BOT(b))
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
	PlaySound("fx_collide"), 0, Csng(velocity) ^2 / 2000, AudioPan(ball1), 0, Pitch(ball1), 0, 0, AudioFade(ball1)
End Sub


'*****************************************
'	ninuzzu's	FLIPPER SHADOWS
'*****************************************

sub FlipperTimer_Timer()

	FlipperLSh.RotZ = LeftFlipper.currentangle
	FlipperRSh.RotZ = RightFlipper.currentangle
	FlipperLSh1.RotZ = LeftFlipper1.currentangle
	FlipperRSh1.RotZ = RightFlipper1.currentangle

	LFPrim.Roty = LeftFlipper.currentangle +240
	RFPrim.Roty = RightFlipper.currentangle +120
	ULFPrim.Roty = LeftFlipper1.currentangle +235
	URFPrim.Roty = RightFlipper1.currentangle +124



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
	PlaySound "pinhit_low", 0, Vol(ActiveBall), AudioPan(ActiveBall), 0, Pitch(ActiveBall), 0, 0, AudioFade(ActiveBall)
End Sub

Sub Targets_Hit (idx)
	PlaySound "target", 0, Vol(ActiveBall), AudioPan(ActiveBall), 0, Pitch(ActiveBall), 0, 0, AudioFade(ActiveBall)
End Sub

Sub Metals_Thin_Hit (idx)
	PlaySound "metalhit_thin", 0, Vol(ActiveBall), AudioPan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
End Sub

Sub Metals_Medium_Hit (idx)
	PlaySound "metalhit_medium", 0, Vol(ActiveBall), AudioPan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
End Sub

Sub Metals2_Hit (idx)
	PlaySound "metalhit2", 0, Vol(ActiveBall), AudioPan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
End Sub

Sub Gates_Hit (idx)
	PlaySound "gate4", 0, Vol(ActiveBall), AudioPan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
End Sub

Sub Spinner_Spin
	PlaySound "fx_spinner", 0, .25, AudioPan(Spinner), 0.25, 0, 0, 1, AudioFade(Spinner)
End Sub

Sub Rubbers_Hit(idx)
 	dim finalspeed
  	finalspeed=SQR(activeball.velx * activeball.velx + activeball.vely * activeball.vely)
 	If finalspeed > 20 then 
		PlaySound "fx_rubber2", 0, Vol(ActiveBall), AudioPan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
	End if
	If finalspeed >= 6 AND finalspeed <= 20 then
 		RandomSoundRubber()
 	End If
End Sub

Sub Posts_Hit(idx)
 	dim finalspeed
  	finalspeed=SQR(activeball.velx * activeball.velx + activeball.vely * activeball.vely)
 	If finalspeed > 16 then 
		PlaySound "fx_rubber2", 0, Vol(ActiveBall), AudioPan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
	End if
	If finalspeed >= 6 AND finalspeed <= 16 then
 		RandomSoundRubber()
 	End If
End Sub

Sub RandomSoundRubber()
	Select Case Int(Rnd*3)+1
		Case 1 : PlaySound "rubber_hit_1", 0, Vol(ActiveBall), AudioPan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
		Case 2 : PlaySound "rubber_hit_2", 0, Vol(ActiveBall), AudioPan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
		Case 3 : PlaySound "rubber_hit_3", 0, Vol(ActiveBall), AudioPan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
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
		Case 1 : PlaySound "flip_hit_1", 0, Vol(ActiveBall), AudioPan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
		Case 2 : PlaySound "flip_hit_2", 0, Vol(ActiveBall), AudioPan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
		Case 3 : PlaySound "flip_hit_3", 0, Vol(ActiveBall), AudioPan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
	End Select
End Sub

' Thalamus : Exit in a clean and proper way
Sub Table1_exit()
  Controller.Pause = False
  Controller.Stop
End Sub


'  Voltan Escapes Cosmic Doom -- Bally, 1978
'  VPX Release December, 2016
'  Thanks to the authors of the VP90 version and for permission to mod:
Option Explicit

'    dboyrecords
'    Design: George Christian
'    Art: Dave Christensen
'    Script: Joe Entropy (joe_entropy@hotmail.com)
'    Table: Eala Dubh Sidhe (EalaDubh@btopenworld.com)

' Thalamus 2018-07-24
' Tables doesn't have "Positional Sound Playback Functions" or "Supporting Ball & Sound Functions"
' Changed UseSolenoids=1 to 2
' Wob 2018-08-08
' Added vpmInit Me to table init and both cSingleLFlip and cSingleRFlip
' Thalamus 2018-11-01 : Improved directional sounds
' Added standard JP ballrolling - you need to enable it on the table, and maybe add some objects into
' the sound enable collections.
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
Const VolGates  = 1    ' Gates volume.
Const VolMetal  = 1    ' Metals volume.
Const VolPi     = 1    ' Rubber pins volume.
Const VolTarg   = 1    ' Targets volume.
Const VolSpin   = 1.5  ' Spinners volume.


LoadVPM "01000100", "BALLY.VBS", 1.2
Dim aX, aY, sX, RSound, SBall 						' ADD FOR STAT BALL ROLL SOUND

Sub LoadVPM(VPMver, VBSfile, VBSver)
	On Error Resume Next
		If ScriptEngineMajorVersion < 5 Then MsgBox "VB Script Engine 5.0 or higher required"
		ExecuteGlobal GetTextFile(VBSfile)
		If Err Then MsgBox "Unable to open " & VBSfile & ". Ensure that it is in the same folder as this table. " & vbNewLine & Err.Description : Err.Clear
		Set Controller = CreateObject("b2s.server")
		'Set Controller = CreateObject("VPinMAME.Controller")
		If Err Then MsgBox "Can't Load VPinMAME." & vbNewLine & Err.Description
		If VPMver>"" Then If Controller.Version < VPMver Or Err Then MsgBox "VPinMAME ver " & VPMver & " required." : Err.Clear
		If VPinMAMEDriverVer < VBSver Or Err Then MsgBox VBSFile & " ver " & VBSver & " or higher required."
	On Error Goto 0
End Sub


Const cGameName="voltan",cCredits="Voltan Escapes Cosmic Doom, Bally 1978",UseSolenoids=2,UseLamps=1,UseGI=0,UseSync=1
' Wob: Added for Fast Flips (No upper Flippers)
Const cSingleLFlip = 0
Const cSingleRFlip = 0
Const SSolenoidOn="solon",SSolenoidOff="soloff",SFlipperOn="FlipperUp",SFlipperOff="FlipperDown",sCoin="coin3"

'TURN STUFF ON OR OFF IF PLAYING IN DESKTOP VS FULLSCREEN
Dim DesktopMode: DesktopMode = Table1.ShowDT


If DesktopMode = True Then 'Show Desktop components
Ramp16.visible=1
Ramp15.visible=1
Primitive13.visible=1
dim xxx
For each xxx in BG:xxx.Visible = 1: Next ' Show DT service lights
Else
'Ramp16.visible=0
'Ramp15.visible=0
'Primitive13.visible=0
For each xxx in BG:xxx.Visible = 0: Next ' Hide DT service lights




End If

Set LampCallback = GetRef("UpdateMultipleLamps")

SolCallback(sBallRelease)="bsTrough.SolOut"
SolCallback(sKnocker)="vpmSolSound ""knocker"","
SolCallback(sLeftSling)="vpmSolSound ""sling"","
SolCallback(sRightSling)="vpmSolSound""sling"","
SolCallback(sBumper1)="vpmSolSound ""jet3"","
SolCallback(sBumper2)="vpmSolSound ""jet3"","
SolCallback(sBumper3)="vpmSolSound ""jet3"","
SolCallback(sEnable)="vpmNudge.SolGameOn"
SolCallback(sLLFlipper)="vpmSolFlipper LeftFlipper,nothing,"
SolCallback(sLRFlipper)="vpmSolFlipper RightFlipper,nothing,"

Const sKnocker=6
Const sBallRelease=7
Const sBumper1=8
Const sBumper2=9
Const sBumper3=10
Const sLeftSling=12
Const sRightSling=11
Const sCLo=18
Const sEnable=19

Dim bsTrough

Sub Table1_Init()
	vpmInit Me
	On Error Resume Next
	With Controller
		.GameName=cGameName
		If Err Then MsgBox "Can't start Game" & cGameName & vbNewLine & Err.Description:Exit Sub
		.SplashInfoLine=cCredits
		.HandleMechanics=0
		.ShowDMDOnly=1
		.ShowFrame=0
		.ShowTitle=0
		.Run
		.Hidden=0

		If Err Then MsgBox Err.Description
	End With
	On Error Goto 0

	PinMAMETimer.Interval=PinMAMEInterval
	PinMAMETimer.Enabled=1
	vpmNudge.TiltSwitch=7
	vpmNudge.Sensitivity=3
	vpmNudge.TiltObj=Array(Bumper1,Bumper2,Bumper3,LeftSlingshot,RightSlingshot)

	Set bsTrough=New cvpmBallStack
    bsTrough.InitSw 0,8,0,0,0,0,0,0
    bsTrough.InitKick BallRelease,80,5.75
    bsTrough.InitExitSnd "ballrel","solon"
    bsTrough.Balls=1
End Sub

' Quick Ball Sound V0.3 by STAT                           ADD BLOCK FOR STAT BALL ROLL SOUND
' -----------------------
Sub TriggerS_Timer()
	aX = int(SBall.VelX): aY = int(SBall.VelY)
	sX = -1.0: If int(Sball.X)>500 Then sX = 1.0
	If (aX>5 OR aY>5) AND Rsound = 0 Then
		RSound=int(RND*4)+1
		PlaySound "Roll "&RSound,0,0.9,sX,0.2
	Elseif (aX<6 AND aY<6) AND Rsound > 0 Then
		StopSound "Roll "&Rsound
		Rsound = 0
	End If
End Sub

Sub TriggerS_Hit()
	Set SBall = Activeball
	TriggerS.TimerInterval = 100
	TriggerS.TimerEnabled = True
	PlaySoundAtVol "Roll 1", ActiveBall, 1
End Sub
' ----------------------- 								   End ADD BLOCK FOR STAT BALL ROLL SOUND

Sub Table1_KeyUp(ByVal KeyCode)
    If vpmKeyUp(KeyCode) Then Exit Sub
    If KeyCode=PlungerKey Then PlaySoundAtVol"Plunger",plunger,1:Plunger.Fire
End Sub

Sub Table1_KeyDown(ByVal KeyCode)
    If vpmKeyDown(KeyCode) Then Exit Sub
    If KeyCode=PlungerKey Then PlaySoundAtVol"PullbackPlunger",plunger,1:Plunger.Pullback
End Sub

Sub Drain_Hit:bsTrough.AddBall Me:TriggerS.TimerEnabled = False:End Sub     ' ADD TriggerS.TimerEnabled = False in a Drain_Hit Sub - FOR STAT BALL ROLL SOUND
																			' and ADD a Trigger in Editor above the Plunger and call it "TriggerS"
Sub RightSpinner_Spin:vpmTimer.PulseSw 17:End Sub
Sub LeftSpinner_Spin:vpmTimer.PulseSw 18:End Sub

Sub Rollover_Hit:Controller.Switch(20)=1:Me.TimerEnabled=1:RollLight.State=0:End Sub
Sub Rollover_unHit:Controller.Switch(20)=0:End Sub
Sub Rollover_Timer:Me.TimerEnabled=0:RollLight.State=1:End Sub

Sub Wall194_Hit:vpmTimer.PulseSw 21:End Sub
Sub Wall193_Hit:vpmTimer.PulseSw 21:End Sub
Sub Wall51_Hit :vpmTimer.PulseSw 21:End Sub
Sub Wall228_Hit:vpmTimer.PulseSw 21:End Sub
Sub Wall192_Hit:vpmTimer.PulseSw 21:End Sub

Sub Target5_Hit:vpmTimer.PulseSw 22:End Sub
Sub TargetA_Hit:vpmTimer.PulseSw 23:End Sub
Sub TargetB_Hit:vpmTimer.PulseSw 24:End Sub
Sub Target8_Hit:vpmTimer.PulseSw 25:End Sub

Sub Lane3_Hit:Controller.Switch(26)=1:End Sub
Sub Lane3_unHit:Controller.Switch(26)=0:End Sub
Sub Lane2_Hit:Controller.Switch(27)=1:End Sub
Sub Lane2_unHit:Controller.Switch(27)=0:End Sub

Sub Target2_Hit:vpmTimer.PulseSw 28:End Sub

Sub Lane4_Hit:Controller.Switch(29)=1:End Sub
Sub Lane4_unHit:Controller.Switch(29)=0:End Sub
Sub LeftInlane_Hit:Controller.Switch(30)=1:End Sub
Sub LeftInlane_unHit:Controller.Switch(30)=0:End Sub
Sub RightInlane_Hit:Controller.Switch(31)=1:End Sub
Sub RightInlane_unHit:Controller.Switch(31)=0:End Sub
Sub Lane1_Hit:Controller.Switch(32)=1:End Sub
Sub Lane1_unHit:Controller.Switch(32)=0:End Sub
Sub RightOutlane_Hit:Controller.Switch(34)=1:End Sub
Sub RightOutlane_unHit:Controller.Switch(34)=0:End Sub
Sub LeftOutlane_Hit:Controller.Switch(35)=1:End Sub
Sub LeftOutlane_unHit:Controller.Switch(35)=0:End Sub

Sub RightSlingshot_Slingshot:vpmTimer.PulseSw 36:End Sub
Sub LeftSlingshot_Slingshot:vpmTimer.PulseSw 37:End Sub

Sub Bumper3_Hit:vpmTimer.PulseSw 38:End Sub
Sub Bumper2_Hit:vpmTimer.PulseSw 39:End Sub
Sub Bumper1_Hit:vpmTimer.PulseSw 40:End Sub

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

If DesktopMode = True Then 'Show Desktop components
FadeReel 13, L13 ' High Score
FadeReel 63, L63 'Match
FadeReel 45 ,L45 'Game Over
FadeReel 61, L61 ' Tilt
End If


End Sub

Sub UpdateMultipleLamps


Set Lights(1)=Light1  'Mercury
Set Lights(2)=Light2   'Jupiter
Set Lights(3)=Light3   'Pluto
Set Lights(4)=Light4   'Number 1
Set Lights(5)=Light5   'Number 7
Set Lights(6)=Light6
Set Lights(7)=Light7   '36000
Set Lights(8)=Light8
Set Lights(9)=Light9
Set Lights(10)=Light10
Set Lights(11)=Light11  'Shoot Again
Set Lights(12)=Light12
Set Lights(13)=Light13 'Ball in Play
Set Lights(14)=Light14  'Player 1
Set Lights(15)=Light15  'Player 1 Number
Set Lights(17)=Light17  'Venus
Set Lights(18)=Light18  'Saturn
Set Lights(19)=Light19
Set Lights(20)=Light20   'Number 4
Set Lights(21)=Light21
Set Lights(22)=Light22
Set Lights(23)=Light23  '18000
Set Lights(24)=Light24   'A Arrow
Set Lights(25)=Light25
Set Lights(26)=Light26  'Bumper3
Set Lights(27)=Light27  'Match
Set Lights(28)=Light28
Set Lights(29)=Light29  'High Score
Set Lights(30)=Light30  'Player 2
Set Lights(31)=Light31  'Player 2 Number
Set Lights(33)=Light33   'Earth
Set Lights(34)=Light34   'Neptune
Set Lights(35)=Light35
Set Lights(36)=Light36   'Number 6
Set Lights(37)=Light37   'Number 2
Set Lights(38)=Light38
Set Lights(39)=Light39   'Special
Set Lights(40)=Light40
Set Lights(41)=Light41 '3X Bonus
Set Lights(42)=Light42
Set Lights(43)=Light43 'Shoot Again
Set Lights(44)=Light44
Set Lights(45)=Light45 'Game Over
Set Lights(46)=Light46  'Player 3
Set Lights(47)=Light47  'Player 3 Number
Set Lights(49)=Light49  'Mars
Set Lights(50)=Light50   'Uranus
Set Lights(51)=Light51
Set Lights(52)=Light52   'Number 9
Set Lights(53)=Light53   'Number 8
Set Lights(54)=Light54
Set Lights(55)=Light55   'Lane Scores 5000
Set Lights(56)=Light56
Set Lights(57)=Light57  'Special
Set Lights(58)=Light58
Set Lights(59)=Light59 'Light on the Apron?  It never turns on.  Shluld the apron light be #19?
Set Lights(60)=Light60
Set Lights(61)=Light61  'Tilt
Set Lights(62)=Light62  'Player 4
Set Lights(63)=Light63  'Player 4 Number



End Sub


'Bally Voltan
'added by Inkochnito
Sub editDips
	Dim vpmDips : Set vpmDips = New cvpmDips
	With vpmDips
		.AddForm 700,400,"Voltan - DIP switches"
		.AddChk 2,10,180,Array("Match feature",&H00100000)'dip 21
		.AddChk 205,10,115,Array("Credits display",&H00080000)'dip 20
		.AddFrame 2,30,190,"Maximum credits",&H00070000,Array("5 credits",0,"10 credits",&H00010000,"15 credits",&H00020000,"20 credits",&H00030000,"25 credits",&H00040000,"30 credits",&H00050000,"35 credits",&H00060000,"40 credits",&H00070000)'dip 17&18&19
		.AddFrame 2,160,190,"Sound features",&H80000080,Array("chime effects",&H80000000,"chime and tunes",0,"noise",&H00000080,"noises and tunes",&H80000080)'dip 8&32
		.AddFrame 2,235,190,"High game to date",&H00000060,Array("no award",0,"1 credit",&H00000020,"2 credits",&H00000040,"3 credits",&H00000060)'dip 6&7
		.AddFrame 2,310,190,"High score feature",&H00006000,Array("no award",0,"extra ball",&H00004000,"replay",&H00006000)'dip 14&15
		.AddFrame 205,30,190,"Balls per game", 32768,Array("3 balls",0,"5 balls", 32768)'dip 16
		.AddFrame 205,76,190,"4 && 6 top lane adjustment",&H00200000,Array("4 && 6 are not tied",0,"4 && 6 tied together",&H00200000)'dip 22
		.AddFrame 205,122,190,"1 && 9 top lane adjustment",&H00400000,Array("1 && 9 are not tied",0,"1 && 9 tied together",&H00400000)'dip 23
		.AddFrame 205,168,190,"Top C arrow lite at start",&H00800000,Array("C arrow lite off",0,"C arrow lite on",&H00800000)'dip 24
		.AddFrame 205,214,190,"Outlane 50K adjustment",&H10000000,Array("alternating",0,"both lanes lit",&H10000000)'dip 29
		.AddFrame 205,260,190,"Extra ball arrow ratio",&H60000000,Array("1 in 10",0,"1 in 15",&H20000000,"1 in 20",&H40000000,"1 in 25",&H60000000)'dip 30&31
		.AddLabel 50,382,300,20,"After hitting OK, press F3 to reset game with new settings."
		.ViewDips
	End With
End Sub
Set vpmShowDips = GetRef("editDips")



Sub LampTimer_Timer()
    Dim chgLamp, num, chg, ii
    chgLamp = Controller.ChangedLamps
    If Not IsEmpty(chgLamp) Then
        For ii = 0 To UBound(chgLamp)
            LampState(chgLamp(ii, 0) ) = chgLamp(ii, 1)       'keep the real state in an array
            FadingLevel(chgLamp(ii, 0) ) = chgLamp(ii, 1) + 4 'actual fading step
        Next
    End If

End Sub

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



' 2nd Player

Digits(6) = Array(LED70,LED71,LED72,LED73,LED74,LED75,LED76)
Digits(7) = Array(LED90,LED91,LED92,LED93,LED94,LED95,LED96)
Digits(8) = Array(LED100,LED101,LED102,LED103,LED104,LED105,LED106)
Digits(9) = Array(LED110,LED111,LED112,LED113,LED114,LED115,LED116)
Digits(10) = Array(LED120,LED121,LED122,LED123,LED124,LED125,LED126)
Digits(11) = Array(LED130,LED131,LED132,LED133,LED134,LED135,LED136)
'

' 3rd Player

Digits(12) = Array(LED150,LED151,LED152,LED153,LED154,LED155,LED156)
Digits(13) = Array(LED160,LED161,LED162,LED163,LED164,LED165,LED166)
Digits(14) = Array(LED170,LED171,LED172,LED173,LED174,LED175,LED176)
Digits(15) = Array(LED180,LED181,LED182,LED183,LED184,LED185,LED186)
Digits(16) = Array(LED190,LED191,LED192,LED193,LED194,LED195,LED196)
Digits(17) = Array(LED200,LED201,LED202,LED203,LED204,LED205,LED206)



' 4th Player
Digits(18) = Array(LED220,LED221,LED222,LED223,LED224,LED225,LED226)
Digits(19) = Array(LED230,LED231,LED232,LED233,LED234,LED235,LED236)
Digits(20) = Array(LED240,LED241,LED242,LED243,LED244,LED245,LED246)
Digits(21) = Array(LED250,LED251,LED252,LED253,LED254,LED255,LED256)
Digits(22) = Array(LED260,LED261,LED262,LED263,LED264,LED265,LED266)
Digits(23) = Array(LED270,LED271,LED272,LED273,LED274,LED275,LED276)


' Credits
Digits(24) = Array(LED4,LED2,LED6,LED7,LED5,LED1,LED3)
Digits(25) = Array(LED18,LED9,LED27,LED28,LED19,LED8,LED17)

' Balls
Digits(26) = Array(LED39,LED37,LED48,LED49,LED47,LED29,LED38)
Digits(27) = Array(LED67,LED58,LED69,LED77,LED68,LED57,LED59)



'Unknowns
'Digits(18) = Array(LED210,LED211,LED212,LED213,LED214,LED215,LED216)
'Digits(13) = Array(LED160,LED161,LED162,LED163,LED164,LED165,LED166)
'Digits(6) = Array(LED80,LED81,LED82,LED83,LED84,LED85,LED86)
'Digits(24) = Array(LED280,LED281,LED282,LED283,LED284,LED285,LED286)


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


' Thalamus : Exit in a clean and proper way
Sub Table1_exit()
  Controller.Pause = False
  Controller.Stop
End Sub

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
	PlaySoundAtVol "fx_spinner", Spinner, VolSpin
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


Option Explicit
Randomize

' Thalamus 2018-07-24
' Table has its own Sound Playback Functions" and "Supporting Ball & Sound Functions"
' Changed UseSolenoids=1 to 2
' No special SSF tweaks yet.

On Error Resume Next
ExecuteGlobal GetTextFile("controller.vbs")
If Err Then MsgBox "You need the controller.vbs in order to run this table, available in the vp10 package"
On Error Goto 0

Const cGameName="stk_sprs",UseSolenoids=2,UseLamps=0,UseGI=0,SSolenoidOn="SolOn",SSolenoidOff="SolOff", SCoin="coin"

LoadVPM "01560000", "bally.vbs", 3.26
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
SolCallback(2)	  =  "vpmSolSound SoundFX(""Bell10"",DOFChimes),"
SolCallback(3)	  =  "vpmSolSound SoundFX(""Bell100"",DOFChimes),"
SolCallback(4)	  =  "vpmSolSound SoundFX(""Bell1000"",DOFChimes),"
SolCallback(5)    =  "vpmSolSound SoundFX(""Bell10000"",DOFChimes),"
SolCallback(6)    =  "vpmSolSound SoundFX(""Knocker"",DOFKnocker),"
SolCallback(7)    =  "bsTrough.SolOut"
SolCallback(8)    =  "bsSaucer.SolOut"
SolCallback(19)   =  "vpmNudge.SolGameOn"

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

Dim bsTrough, bsSaucer, mHole

Sub Table1_Init
	vpmInit Me
	On Error Resume Next
		With Controller
		.GameName = cGameName
		If Err Then MsgBox "Can't start Game" & cGameName & vbNewLine & Err.Description : Exit Sub
		.SplashInfoLine = "Syrikes and Spares Bally"&chr(13)&"You Suck"
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

	vpmNudge.TiltSwitch = 7
	vpmNudge.Sensitivity = 2
	vpmNudge.TiltObj = Array(Bumper1, Bumper2, Bumper3, LeftSlingshot, RightSlingShot)

	Set bsTrough = New cvpmBallStack
 	    bsTrough.InitSw 0,8,0,0,0,0,0,0
      	bsTrough.InitKick BallRelease, 45, 8
  		bsTrough.InitExitSnd SoundFX("ballrelease",DOFContactors), SoundFX("Solenoid",DOFContactors)
 		bsTrough.Balls = 1

	Set bsSaucer = New cvpmBallStack
      	bsSaucer.InitSaucer sw32, 32, 160, 8
 		bsSaucer.KickForceVar = 2
		bsSaucer.KickAngleVar = 2
  		bsSaucer.InitExitSnd SoundFX("Popper",DOFContactors), SoundFX("Solenoid",DOFContactors)

     Set mHole = New cvpmMagnet
         mHole.InitMagnet Umagnet, 3
         mHole.GrabCenter = 0
         mHole.MagnetOn = 1
         mHole.CreateEvents "mHole"
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
Sub sw32_Hit:bsSaucer.AddBall 0 : playsound "popper_ball": End Sub

 'Scoring Rubber
Sub sw13_hit:vpmTimer.pulseSw 13 : playsound"flip_hit_3" : End Sub
Sub sw14_hit:vpmTimer.pulseSw 14 : playsound"flip_hit_3" : End Sub

'Spinners
Sub sw25_Spin:vpmTimer.PulseSw 25 : playsound"fx_spinner" : End Sub

'Wire Triggers
  Sub sw1_Hit:Controller.Switch(1) = 1:PlaySound "rollover":End Sub
  Sub sw1_UnHit:Controller.Switch(1) = 0:End Sub
  Sub sw2_Hit:Controller.Switch(2) = 1:PlaySound "rollover":End Sub
  Sub sw2_UnHit:Controller.Switch(2) = 0:End Sub
  Sub sw3_Hit:Controller.Switch(3) = 1:PlaySound "rollover":End Sub
  Sub sw3_UnHit:Controller.Switch(3) = 0:End Sub
  Sub sw4_Hit:Controller.Switch(4) = 1:PlaySound "rollover":End Sub
  Sub sw4_UnHit:Controller.Switch(4) = 0:End Sub
  Sub sw5_Hit:Controller.Switch(5) = 1:PlaySound "rollover":End Sub
  Sub sw5_UnHit:Controller.Switch(5) = 0:End Sub
  Sub sw34_Hit:Controller.Switch(34) = 1:PlaySound "rollover":End Sub
  Sub sw34_UnHit:Controller.Switch(34) = 0:End Sub
  Sub sw35_Hit:Controller.Switch(35) = 1:PlaySound "rollover":End Sub
  Sub sw35_UnHit:Controller.Switch(35) = 0:End Sub

'Star Triggers
  Sub sw17_Hit:Controller.Switch(17) = 1:PlaySound "rollover":End Sub
  Sub sw17_UnHit:Controller.Switch(17) = 0:End Sub
  Sub sw18_Hit:Controller.Switch(18) = 1:PlaySound "rollover":End Sub
  Sub sw18_UnHit:Controller.Switch(18) = 0:End Sub
  Sub sw19_Hit:Controller.Switch(19) = 1:PlaySound "rollover":End Sub
  Sub sw19_UnHit:Controller.Switch(19) = 0:End Sub
  Sub sw20_Hit:Controller.Switch(20) = 1:PlaySound "rollover":End Sub
  Sub sw20_UnHit:Controller.Switch(20) = 0:End Sub
  Sub sw21_Hit:Controller.Switch(21) = 1:PlaySound "rollover":End Sub
  Sub sw21_UnHit:Controller.Switch(21) = 0:End Sub
  Sub sw22_Hit:Controller.Switch(22) = 1:PlaySound "rollover":End Sub
  Sub sw22_UnHit:Controller.Switch(22) = 0:End Sub
  Sub sw23_Hit:Controller.Switch(23) = 1:PlaySound "rollover":End Sub
  Sub sw23_UnHit:Controller.Switch(23) = 0:End Sub
  Sub sw24_Hit:Controller.Switch(24) = 1:PlaySound "rollover":End Sub
  Sub sw24_UnHit:Controller.Switch(24) = 0:End Sub
  Sub sw26_Hit:Controller.Switch(26) = 1:PlaySound "rollover":End Sub
  Sub sw26_UnHit:Controller.Switch(26) = 0:End Sub
  Sub sw27_Hit:Controller.Switch(27) = 1:PlaySound "rollover":End Sub
  Sub sw27_UnHit:Controller.Switch(27) = 0:End Sub

 'Stand Up Targets
Sub sw12_hit:vpmTimer.pulseSw 12 : End Sub
Sub sw28_hit:vpmTimer.pulseSw 28 : End Sub
Sub sw29_hit:vpmTimer.pulseSw 29 : End Sub
Sub sw30_hit:vpmTimer.pulseSw 30 : End Sub
Sub sw31_hit:vpmTimer.pulseSw 31 : End Sub
Sub sw33_hit:vpmTimer.pulseSw 33 : End Sub

'Bumpers
Sub Bumper1_Hit : vpmTimer.PulseSw(38) : playsound SoundFX("fx_bumper1",DOFContactors): End Sub
Sub Bumper2_Hit : vpmTimer.PulseSw(39) : playsound SoundFX("fx_bumper1",DOFContactors): End Sub
Sub Bumper3_Hit : vpmTimer.PulseSw(40) : playsound SoundFX("fx_bumper1",DOFContactors): End Sub


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
         NFadeL 4, l4
         NFadeL 5, l5
         NFadeL 6, l6
         NFadeL 7, l7
         NFadeL 8, l8
         NFadeL 9, l9
         NFadeL 10, l10
        ' NFadeL 11, l11
         NFadeL 12, l12
         'NFadeL 13, l13  'BIP
         NFadeL 14, l14
         'NFadeL 15, l15  Player1
         NFadeL 17, l17
         NFadeL 18, l18
         NFadeL 19, l19
         NFadeL 20, l20
         NFadeL 21, l21
         NFadeL 22, l22
         NFadeL 23, l23
         NFadeL 24, l24
         NFadeL 25, l25
         NFadeL 26, l26
         'NFadeL 27, l27  'Match
         NFadeL 28, l28
 		'NFadeL 29  'HighScore
         NFadeL 30, l30
         'NFadeL 31, l31 'Player2
         NFadeL 33, l33
         NFadeL 34, l34
         NFadeL 35, l35
         NFadeL 36, l36
         NFadeL 37, l37
         NFadeL 38, l38
         NFadeL 39, l39
         NFadeL 40, l40
         NFadeL 41, l41
         NFadeLm 42, l42   'Center Bumper
         NFadeL 42, l42a
         NFadeL 43, l43
         NFadeL 44, l44
         'NFadeL 45, l45 'Game Over
         NFadeL 46, l46
         'NFadeL 47, l47 'Player3
         NFadeL 49, l49
         NFadeL 50, l50
         NFadeL 51, l51
         NFadeL 52, l52
         NFadeL 53, l53
         NFadeL 54, l54
         NFadeL 55, l55
         NFadeL 56, l56
         NFadeL 57, l57
         NFadeL 58, l58
         NFadeL 59, l59
         NFadeL 60, l60
         'NFadeL 61, l61 'Tilt
         NFadeL 62, l62
         'NFadeL 63, l63 'Player4
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

Dim Digits(28)
' 1st Player
Digits(0) = Array(LED10,LED11,LED12,LED13,LED14,LED15,LED16)
Digits(1) = Array(LED20,LED21,LED22,LED23,LED24,LED25,LED26)
Digits(2) = Array(LED30,LED31,LED32,LED33,LED34,LED35,LED36)
Digits(3) = Array(LED40,LED41,LED42,LED43,LED44,LED45,LED46)
Digits(4) = Array(LED50,LED51,LED52,LED53,LED54,LED55,LED56)
Digits(5) = Array(LED60,LED61,LED62,LED63,LED64,LED65,LED66)

' 2nd Player
Digits(6) = Array(LED80,LED81,LED82,LED83,LED84,LED85,LED86)
Digits(7) = Array(LED90,LED91,LED92,LED93,LED94,LED95,LED96)
Digits(8) = Array(LED100,LED101,LED102,LED103,LED104,LED105,LED106)
Digits(9) = Array(LED110,LED111,LED112,LED113,LED114,LED115,LED116)
Digits(10) = Array(LED120,LED121,LED122,LED123,LED124,LED125,LED126)
Digits(11) = Array(LED130,LED131,LED132,LED133,LED134,LED135,LED136)

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

Sub DisplayTimer_Timer
	Dim ChgLED,ii,num,chg,stat,obj
	ChgLed = Controller.ChangedLEDs(&Hffffffff, &Hffffffff)
If Not IsEmpty(ChgLED) Then
		If DesktopMode = True Then
		For ii = 0 To UBound(chgLED)
			num = chgLED(ii, 0) : chg = chgLED(ii, 1) : stat = chgLED(ii, 2)
			if (num < 28) then
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

'Bally Strikes and spares
'added by Inkochnito
Sub editDips
	Dim vpmDips : Set vpmDips = New cvpmDips
	With vpmDips
		.AddForm 700,400,"Strikes and spares - DIP switches"
		.AddFrame 2,0,190,"Maximum credits",&H00070000,Array("10 credits",&H00010000,"20 credits",&H00030000,"30 credits",&H00050000,"40 credits",&H00070000)'dip 17&18&19
		.AddFrame 2,79,190,"High game to date award",&H00000060,Array("no award",0,"1 credit",&H00000020,"2 credits",&H00000040,"3 credits",&H00000060)'dip 6&7
		.AddFrame 2,153,190,"Multiplier adjust",&H2000000,Array("2X, 3X, 5X, 5K, special",0,"2X, 3X, 5X, 5K && special, special",&H20000000)'dip 30
		.AddFrame 2,199,190,"Automatic strike center lane",&H40000000,Array("lite is alternates",0,"lite is on all the time",&H40000000)'dip 31
		.AddFrame 205,0,190,"Balls per game",32768,Array("3 balls",0,"5 balls",32768)'dip 16
		.AddFrame 205,46,190,"High score feature",&H00006000,Array("no award",0,"extra ball",&H00004000,"replay",&H00006000)'dip 14&15
		.AddFrame 205,107,190,"Extra ball lite return lane",&H00800000,Array("alternating",0,"both on",&H00800000)'dip 24
		.AddFrame 205,153,190,"Outlane special",&H80000000,Array("alternating",0,"both on",&h80000000)'dip 32
		.AddFrame 205,199,190,"Pin adjust",&H10000000,Array("all seperate",0,"1&&3, 2&&5, 8&&9 are tied",&H10000000)'dip 29
		.AddChk 2,250,120,Array("Match",&H00100000)'dip 21
		.AddChk 2,265,120,Array("Credits displayed",&H00080000)'dip 20
		.AddChk 2,280,120,Array("Melody option",&H00000080)'dip 8
		.AddChk 205,250,150,Array("Bonus multiplier memory",&H00200000)'dip 22
		.AddChk 205,265,150,Array("Start strike value at 5K",&H00400000)'dip 23
		.AddLabel 50,300,300,20,"After hitting OK, press F3 to reset game with new settings."
		.ViewDips
	End With
End Sub
Set vpmShowDips = GetRef("editDips")

'**********************************************************************************************************
'**********************************************************************************************************
'	Start of VPX functions
'**********************************************************************************************************
'**********************************************************************************************************


'**********Sling Shot Animations
' Rstep and Lstep  are the variables that increment the animation
'****************
Dim RStep, Lstep

Sub RightSlingShot_Slingshot
	vpmTimer.PulseSw 36
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
	vpmTimer.PulseSw 37
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

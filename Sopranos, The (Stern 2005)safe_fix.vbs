
Randomize

Dim Ballsize,BallMass
Ballsize = 50
BallMass = 1.2

' Thalamus 2018-07-24
' Tables hasits own "Positional Sound Playback Functions" and "Supporting Ball & Sound Functions"
' Changed UseSolenoids=1 to 2
' No special SSF tweaks yet.

On Error Resume Next
ExecuteGlobal GetTextFile("controller.vbs")
If Err Then MsgBox "You need the controller.vbs in order to run this table, available in the vp10 package"
On Error Goto 0

Const cGameName="sopranos",UseSolenoids=2,UseLamps=0,UseGI=0,SSolenoidOn="SolOn",SSolenoidOff="SolOff", SCoin="coin"

LoadVPM "01560000", "sega.VBS", 3.10

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

Const RomSounds	= 1	'1 - use the sounds from the rom, 0 - use sampled sounds (if the rom sounds like shit) (it probably does)

'*************************************************************
'Solenoid Call backs
'**********************************************************************************************************

SolCallback(1)  = "solTrough"
SolCallback(2)  = "solAutofire"
SolCallback(3)  = "bsTEject.SolOut"
SolCallback(4)  = "SolCenterLock" 
SolCallBack(5)  = "SolGateL"
SolCallBack(6)  = "SolGateR"
SolCallback(7)  = "dtSingle.SolHit 1,"
SolCallback(8)  = "SolSafe"
SolCallback(14) = "dtSingle.SolDropUp"
SolCallback(17) = "SolFish"
SolCallback(18) = "Strippers"
SolCallBack(21) = "bsRScoop.SolOut"
SolCallback(22) = "SolBoatLock"
SolCallback(23) = "SolBingLock"
SolCallback(30) = "SolSafeLatch"

SolCallback(19) = "SetLamp 119," 'PF light
SolCallback(20) = "SetLamp 120," 'PF light
SolCallback(25) = "SetLamp 125," 'Fish eyes
SolCallback(26) = "SetLamp 126," 'stage Yellow Dome
SolCallback(27) = "SetLamp 127," 'left sling Yellow Dome
SolCallback(28) = "SetLamp 128," 'right sling Yellow Dome
SolCallback(29) = "SetLamp 129," 'pop bumper Red Dome  X2
SolCallback(31) = "SetLamp 131," 'right spotlight
SolCallback(32) = "SetLamp 132,"

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

Sub solTrough(Enabled)
	If Enabled Then
		bsTrough.ExitSol_On
		vpmTimer.PulseSw 22
	End If
 End Sub

Sub solAutofire(Enabled)
	If Enabled Then
		PlungerIM.AutoFire
	End If
 End Sub

Sub SolCenterLock(Enabled)
    If Enabled Then
        CenterPost.IsDropped = 1
		CenterPin.transz =-40
		playsound SoundFX("Popper",DOFContactors)
	Else
		CenterPost.IsDropped = 0
		CenterPin.transz = 0
    End If
End Sub

Sub SolGateL(Enabled)
    If Enabled then
        sol5.Open=1
		playsound"Diverter" 
    Else
        sol5.Open=0
    End If
End Sub

Sub SolGateR(Enabled)
    If Enabled then
        sol6.Open=1
		playsound"Diverter" 
    Else
        sol6.Open=0
    End If
End Sub

Sub SolFish(Enabled)
    If Enabled then
        fishf.RotateToEnd
		playsound"Diverter" 
    Else
        fishf.RotateToStart
    End If
End Sub

Sub SolBoatLock(Enabled)
    If Enabled Then
        BoatPost.IsDropped = 1 'Drop the post
		BoatPin.transy =-80
		playsound SoundFX("Popper",DOFContactors)
	Else
		BoatPost.IsDropped = 0
		BoatPin.transy = 0
    End If
End Sub

Sub SolBingLock(Enabled)
    If Enabled Then
        BingPost.IsDropped = 1 'Drop the post
		playsound SoundFX("Popper",DOFContactors)
	Else
		BingPost.IsDropped = 0
    End If
End Sub

'Stripper animation
'**********************************************************************************************************
Sub Strippers(Enabled)
	If Enabled Then
		strippert.Enabled = 1
	Else
		strippert.Enabled = 0
	End If
End Sub

Sub strippert_Timer
	stripper2.objrotz = stripper2.objrotz + .5
	stripper1.objrotz = stripper1.objrotz - .5
	stripper1s.objrotz = stripper1.objrotz
	stripper2s.objrotz = stripper2.objrotz
End Sub

'**********************************************************************************************************

'Safe animation
'**********************************************************************************************************
Sub SolSafe(Enabled)
    If Latch = 1 Then
		prisonstate = false
        Exit Sub
    End If
    If Enabled Then
		prisonstate = false
    Else 'this means if the solenoid is NOT enabled
		prisonstate = true
    End If
	PrisonT.Enabled=true
End Sub

Dim Latch

Sub SolSafeLatch(Enabled)
    If Enabled Then
        Latch = 1
    Else
        Latch = 0
    End If
End Sub

Dim SafePos2
SafePos2 = 0

Dim prisonstate
prisonstate = False
Sub PrisonT_Timer()
	If prisonstate = True then 'Opening
		If sw21p.RotX <= 22 then 
			sw21p.RotX = sw21p.RotX + 2
			sw24p.RotX = sw24p.RotX - 2
			sw21p.Z = sw21p.Z + 6
			sw24p.Z = sw24p.Z + 6
		Else
			PrisonT.Enabled = False
			Controller.Switch(10) = 1
		End If
		sw21.isdropped = true
		sw24.isdropped = true
'		Wall25.isdropped = false
	Else	'Closing
		If sw21p.RotX => 2 then 
			sw21p.RotX = sw21p.RotX - 2
			sw24p.RotX = sw24p.RotX + 2
			sw21p.Z = sw21p.Z - 6
			sw24p.Z = sw24p.Z - 6
		Else
			PrisonT.Enabled = False
			Controller.Switch(10) = 0
		End If
		sw21.isdropped = false
		sw24.isdropped = false
'		Wall25.isdropped = true
	End If
End Sub
'**********************************************************************************************************

'Stern-Sega GI 
set GICallback = GetRef("UpdateGI")

Sub UpdateGI(no, Enabled)
	If Enabled Then
		dim xx
		For each xx in GI:xx.State = 1: Next
        PlaySound "fx_relay"
	Else
		For each xx in GI:xx.State = 0: Next
        PlaySound "fx_relay"
	End If
End Sub

'**********************************************************************************************************
'Initiate Table
'**********************************************************************************************************

Dim bsTrough, bsTEject, bsRScoop, dtSingle

Sub Table1_Init
	vpmInit Me
	On Error Resume Next
		With Controller
		.GameName = cGameName
		If Err Then MsgBox "Can't start Game" & cGameName & vbNewLine & Err.Description : Exit Sub
		.SplashInfoLine = "The Sopranos (Stern 2005)"&chr(13)&"You Suck"
		.HandleMechanics=0
		.HandleKeyboard=0
		.ShowDMDOnly=1
		.ShowFrame=0
		.ShowTitle=0
        .hidden = 0
		.Games(cGameName).Settings.Value("sound") = RomSounds
         On Error Resume Next
         .Run GetPlayerHWnd
         If Err Then MsgBox Err.Description
         On Error Goto 0
     End With
     On Error Goto 0

	PinMAMETimer.Interval = PinMAMEInterval
	PinMAMETimer.Enabled = 1
    vpmNudge.TiltSwitch = -7
   	vpmNudge.Sensitivity = 3
	vpmNudge.TiltObj=Array(Bumper1,Bumper2,Bumper3,LeftSlingshot,RightSlingshot)

    Set bsTrough = New cvpmBallStack
		bsTrough.InitSw 0, 14, 13, 12, 11, 0, 0, 0
		bsTrough.InitKick BallRelease, 90, 7.5
		bsTrough.InitExitSnd SoundFX("ballrelease",DOFContactors), SoundFX("Solenoid",DOFContactors)
		bsTrough.Balls = 4

    Set bsTEject = new cvpmBallStack
        bsTEject.InitSaucer sw28, 28, 54, 15
		bsTEject.KickZ = 1
        bsTEject.InitExitSnd SoundFX("Popper",DOFContactors), SoundFX("Solenoid",DOFContactors)

    Set bsRScoop=New cvpmBallStack
        bsRScoop.InitSaucer sw17, 17, 155, 10
		'bsRScoop.KickZ = 1
        bsRScoop.InitExitSnd SoundFX("Popper",DOFContactors), SoundFX("Solenoid",DOFContactors)

	set dtSingle=new cvpmDropTarget
		dtSingle.InitDrop sw26,26
		dtSingle.InitSnd SoundFX("DTDrop",DOFContactors),SoundFX("DTReset",DOFContactors)

  End Sub

'**********************************************************************************************************
'Plunger code
'**********************************************************************************************************

Sub Table1_KeyDown(ByVal KeyCode)
	If KeyDownHandler(keycode) Then Exit Sub
	If keycode = PlungerKey Then Plunger.Pullback:playsound"plungerpull"
	If Keycode = StartGameKey Then Controller.Switch(16) = 1
End Sub

Sub Table1_KeyUp(ByVal KeyCode)
	If KeyUpHandler(keycode) Then Exit Sub
	If keycode = PlungerKey Then Plunger.Fire:PlaySound"plunger"
	If Keycode = StartGameKey Then Controller.Switch(16) = 0
End Sub

	Dim PlungerIM
	Const IMPowerSetting = 50
    Const IMTime = 0.6
    Set plungerIM = New cvpmImpulseP
    With plungerIM
        .InitImpulseP swplunger, IMPowerSetting, IMTime
        .Random 0.3
		.Switch 16
        .InitExitSnd  SoundFX("Popper",DOFContactors), SoundFX("Solenoid",DOFContactors)
        .CreateEvents "plungerIM"
    End With

'**********************************************************************************************************

 ' Drain hole and kickers
Sub Drain_Hit:bsTrough.addball me : playsound"drain" : End Sub
Sub sw17_Hit:bsRScoop.AddBall Me : playsound "popper_ball": End Sub
Sub sw28_Hit:bsTEject.AddBall 0 : playsound "popper_ball": End Sub

'Drop Targets
Sub sw26_Dropped:dtSingle.Hit 1:End Sub

'Bing Ramp Triggers
Sub sw19_Hit  : Controller.Switch(19) = 1 : playsound"rollover" : End Sub
Sub sw19_UnHit: Controller.Switch(19) = 0: End Sub
Sub sw20_Hit  : Controller.Switch(20) = 1 : playsound"rollover" : End Sub
Sub sw20_UnHit: Controller.Switch(20) = 0: End Sub

'Boat Hidden Triggers
Sub sw31_Hit  : Controller.Switch(31) = 1: End Sub
Sub sw31_UnHit: Controller.Switch(31) = 0: End Sub
Sub sw32_Hit  : Controller.Switch(32) = 1: End Sub
Sub sw32_UnHit: Controller.Switch(32) = 0: End Sub

'SafeHouse Primitive
Sub sw21_Hit  : vpmTimer.PulseSw 21:sw21.TimerEnabled = 1:sw21p.TransX = -4: playsound"Target": End Sub
Sub sw21_Timer:Me.TimerEnabled = 0:sw21p.TransX = 0:End Sub
Sub sw24_Hit  : vpmTimer.PulseSw 24:sw24.TimerEnabled = 1:sw24p.TransX = -4: playsound"Target": End Sub
Sub sw24_Timer:Me.TimerEnabled = 0:sw24p.TransX = 0:End Sub

'Center Pin Lock
Sub sw22_Hit:Controller.Switch(22) = 1:End Sub
Sub sw22_UnHit:Controller.Switch(22) = 0:End Sub
Sub sw23_Hit:Controller.Switch(23) = 1:End Sub
Sub sw23_UnHit:Controller.Switch(23) = 0:End Sub

'Gate Triggers
Sub sw9_Hit: vpmTimer.PulseSw 9: End Sub
Sub sw18_Hit: vpmTimer.PulseSw 18: End Sub
Sub sw29_Hit: vpmTimer.PulseSw 29: End Sub
Sub sw33_Hit: vpmTimer.PulseSw 33: End Sub

'Spinners
Sub sw25_Spin:vpmTimer.PulseSw 25 : playsound"fx_spinner" : End Sub
Sub sw27_Spin:vpmTimer.PulseSw 27 : playsound"fx_spinner" : End Sub

'Wire Triggers
Sub sw38_Hit  : Controller.Switch(38) = 1 : playsound"rollover" : End Sub
Sub sw38_UnHit: Controller.Switch(38) = 0: End Sub
Sub sw39_Hit  : Controller.Switch(39) = 1 : playsound"rollover" : End Sub
Sub sw39_UnHit: Controller.Switch(39) = 0: End Sub
Sub sw40_Hit  : Controller.Switch(40) = 1 : playsound"rollover" : End Sub
Sub sw40_UnHit: Controller.Switch(40) = 0: End Sub
Sub sw57_Hit  : Controller.Switch(57) = 1 : playsound"rollover" : End Sub
Sub sw57_UnHit: Controller.Switch(57) = 0: End Sub
Sub sw58_Hit  : Controller.Switch(58) = 1 : playsound"rollover" : End Sub
Sub sw58_UnHit: Controller.Switch(58) = 0: End Sub
Sub sw60_Hit  : Controller.Switch(60) = 1 : playsound"rollover" : End Sub
Sub sw60_UnHit: Controller.Switch(60) = 0: End Sub
Sub sw61_Hit  : Controller.Switch(61) = 1 : playsound"rollover" : End Sub
Sub sw61_UnHit: Controller.Switch(61) = 0: End Sub

'StandUp Targets
Sub sw34_Hit  : vpmTimer.PulseSw 34: End Sub
Sub sw35_Hit  : vpmTimer.PulseSw 35: End Sub
Sub sw36_Hit  : vpmTimer.PulseSw 36: End Sub
Sub sw37_Hit  : vpmTimer.PulseSw 37: End Sub

'Bumpers
Sub Bumper1_Hit : vpmTimer.PulseSw(49) : playsound SoundFX("fx_bumper1",DOFContactors): End Sub
Sub Bumper2_Hit : vpmTimer.PulseSw(50) : playsound SoundFX("fx_bumper1",DOFContactors): End Sub
Sub Bumper3_Hit : vpmTimer.PulseSw(51) : playsound SoundFX("fx_bumper1",DOFContactors): End Sub

'Generic Sounds
Sub Trigger1_Hit : playsound"fx_ballrampdrop" : End Sub
Sub Trigger2_Hit : playsound"fx_ballrampdrop" : End Sub
Sub Trigger3_Hit : playsound"fx_ballrampdrop" : End Sub

Sub Trigger4_Hit : playsound"Wire Ramp" : End Sub
Sub Trigger5_Hit : playsound"Wire Ramp" : End Sub

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
		NFadeL 11, l11
		NFadeL 12, l12
		NFadeL 13, l13
		NFadeL 14, l14
		NFadeL 15, l15
		NFadeL 16, l16
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
		NFadeL 27, l27
		NFadeL 28, l28
		NFadeL 29, l29
		NFadeL 30, l30
		NFadeL 31, l31
		NFadeL 32, l32
		NFadeL 33, l33
		NFadeL 34, l34
		NFadeL 35, l35
		NFadeL 36, l36
		NFadeL 37, l37
		NFadeL 38, l38
		NFadeL 39, l39
		NFadeL 40, l40
		NFadeL 41, l41
		NFadeL 42, l42
		NFadeL 43, l43
		NFadeL 44, l44
		NFadeL 45, l45
		NFadeL 46, l46
		NFadeL 47, l47
		NFadeL 48, l48
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
		NFadeL 61, l61
'		NFadeL 62, l62
'		NFadeL 63, l63
'		NFadeL 64, l64

'BackWall Lights
	Flash 65, l65
	Flash 66, l66
	Flash 67, l67
	Flash 68, l68
	Flash 69, l69
	Flash 70, l70
	Flash 71, l71
	Flash 72, l72

'Stand Up Board
    NFadeObjm 73, l73, "bulbcover1_yellowOn", "bulbcover1_yellow"
	Flash 73, F73
    NFadeObjm 74, l74, "bulbcover1_yellowOn", "bulbcover1_yellow"
	Flash 74, F74
    NFadeObjm 75, l75, "bulbcover1_yellowOn", "bulbcover1_yellow"
	Flash 75, F75
    NFadeObjm 76, l76, "bulbcover1_yellowOn", "bulbcover1_yellow"
	Flash 76, F76
    NFadeObjm 77, l77, "bulbcover1_yellowOn", "bulbcover1_yellow"
	Flash 77, F77


	NFadeL 78, l78
'	NFadeL 79, l79
'	NFadeL 80, l80

'Solenoid Controlled

	NFadeL 119, L119

	NFadeL 120, L120

	Flashm 125, f125a
	Flash 125, f125b

	NFadeObjm 126, P126, "dome2_0_yellowOn", "dome2_0_yellow"
	NFadeL 126, L126

	NFadeObjm 127, P127, "dome2_0_yellowOn", "dome2_0_yellow"
	NFadeL 127, f127

	NFadeObjm 128, P128, "dome2_0_yellowOn", "dome2_0_yellow"
	NFadeL 128, f128

	NFadeObjm 129, P129a, "dome2_0_redOn", "dome2_0_red"
	NFadeObjm 129, P129b, "dome2_0_redOn", "dome2_0_red"
	NFadeLm 129, L29a
	NFadeL 129, L29b

	NFadeL 131, L131

	NFadeLm 132, L32a
	NFadeL 132, L32b

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
'**********************************************************************************************************
'	Start of VPX functions
'**********************************************************************************************************
'**********************************************************************************************************

'**********Sling Shot Animations
' Rstep and Lstep  are the variables that increment the animation
'****************
Dim RStep, Lstep

Sub RightSlingShot_Slingshot
	vpmTimer.PulseSw 62
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
	vpmTimer.PulseSw 59
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
        Case 4:LSLing2.Visible = 0:LSLing.Visible = 1:sling2.TransZ = 0:LeftSlingShot.TimerEnabled = 0
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
    LFLogo.RotY = LeftFlipper.CurrentAngle
    RFlogo.RotY = RightFlipper.CurrentAngle
	sw9p.RotX = -(sw9.currentangle) +90
	sw29p.RotX = -(sw29.currentangle) +90
	fishmouth.ObjRotX = fishf.CurrentAngle
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
            BallShadow(b).X = ((BOT(b).X) - (Ballsize/6) + ((BOT(b).X - (Table1.Width/2))/7)) '+ 6
        Else
            BallShadow(b).X = ((BOT(b).X) + (Ballsize/6) + ((BOT(b).X - (Table1.Width/2))/7)) '- 6
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



'******************************
' ROM SOUNDS played as samples
'   code by destruk/Gaston ?
'******************************

Dim Playing
Playing = 0

If RomSounds = 0 Then
Set MotorCallback=GetRef("TrackSounds")
End If

'Music & Sound Stuff
Sub TrackSounds
    Dim NewSounds, ii, Snd
    NewSounds = Controller.NewSoundCommands
    If Not IsEmpty(NewSounds) Then
        For ii = 0 To UBound(NewSounds)
            Snd = NewSounds(ii, 0)
            If Snd = 252 Then Playing = 1 'FC
            If Snd = 253 Then Playing = 2 'FD
            If Snd = 254 Then Playing = 3 'FE
            If Snd <> 255 And Snd <> 1 And Snd <> 0 And Snd <> 16 Then
                SoundCommand(Snd)
            End If
        Next
    End If
End Sub

Dim HexRed
HexRed = "0"

Sub SoundCommand(Cmd)
    Dim SndName
    If Playing = 1 Then SndName = "FC"
    If Playing = 2 Then SndName = "FD"
    If Playing = 3 Then SndName = "FE"
    If Cmd < 40 And Playing = 2 Then
        If Timer1.Enabled Then Timer1.Enabled = 0
        Select Case Cmd
            Case 3:Timer1.Interval = 20991:Timer1.Enabled = 1:MStore = Cmd:HexRed = "0"
            Case 4:Timer1.Interval = 3816:Timer1.Enabled = 1:MStore = Cmd:HexRed = "0"
            Case 8:Timer1.Interval = 4997:Timer1.Enabled = 1:MStore = Cmd:HexRed = "0"
            Case 10:Timer1.Interval = 613:Timer1.Enabled = 1:MStore = Cmd:HexRed = "0"
            Case 11:Timer1.Interval = 1977:Timer1.Enabled = 1:MStore = Cmd:HexRed = "0"
            Case 12:Timer1.Interval = 6214:Timer1.Enabled = 1:MStore = Cmd:HexRed = "0"
            Case 13:Timer1.Interval = 4260:Timer1.Enabled = 1:MStore = Cmd:HexRed = "0"
            Case 14:Timer1.Interval = 41872:Timer1.Enabled = 1:MStore = Cmd:HexRed = "0"
            Case 18:Timer1.Interval = 3827:Timer1.Enabled = 1:MStore = Cmd
            Case 19:Timer1.Interval = 1278:Timer1.Enabled = 1:MStore = Cmd
            Case 20:Timer1.Interval = 8851:Timer1.Enabled = 1:MStore = Cmd
            Case 25:Timer1.Interval = 3243:Timer1.Enabled = 1:MStore = Cmd
            Case 36:Timer1.Interval = 59780:Timer1.Enabled = 1:MStore = Cmd
            Case 37:Timer1.Interval = 19995:Timer1.Enabled = 1:MStore = Cmd
            Case 38:Timer1.Interval = 59780:Timer1.Enabled = 1:MStore = Cmd
            Case 39:Timer1.Interval = 19995:Timer1.Enabled = 1:MStore = Cmd
        End Select

        MusicCommand(Cmd)
    Else
        Dim FinalSnd
        FinalSnd = HEX(Cmd)
        SndName = SndName & FinalSnd
        PlaySound SndName
    End If
End Sub

Dim LastMus
LastMus = " "
Dim MStore
MStore = 0

Sub MusicCommand(Cmd)
    If Len(LastMus) > 0 Or Cmd = 0 Then
        StopSound LastMus
    End If
    Dim FinalMus
    FinalMus = Hex(Cmd)
    Select Case Cmd
        Case 3:LastMus = "FD" & "0" & FinalMus & "A":PlaySound LastMus
        Case 4:LastMus = "FD" & "0" & FinalMus & "A":PlaySound LastMus
        Case 8:LastMus = "FD" & "0" & FinalMus & "A":PlaySound LastMus
        Case 10:LastMus = "FD" & "0" & FinalMus & "A":PlaySound LastMus
        Case 11:LastMus = "FD" & "0" & FinalMus & "A":PlaySound LastMus
        Case 12:LastMus = "FD" & "0" & FinalMus & "A":PlaySound LastMus
        Case 13:LastMus = "FD" & "0" & FinalMus & "A":PlaySound LastMus
        Case 14:LastMus = "FD" & "0" & FinalMus & "A":PlaySound LastMus
        Case 18:LastMus = "FD" & FinalMus & "A":PlaySound LastMus
        Case 19:LastMus = "FD" & FinalMus & "A":PlaySound LastMus
        Case 20:LastMus = "FD" & FinalMus & "A":PlaySound LastMus
        Case 25:LastMus = "FD" & FinalMus & "A":PlaySound LastMus
        Case 36:LastMus = "FD" & FinalMus & "A":PlaySound LastMus
        Case 37:LastMus = "FD" & FinalMus & "A":PlaySound LastMus
        Case 38:LastMus = "FD" & FinalMus & "A":PlaySound LastMus
        Case 39:LastMus = "FD" & FinalMus & "A":PlaySound LastMus
        Case Else:LastMus = "FD" & FinalMus:PlaySound LastMus, -1
    End Select
End Sub

Sub Timer1_Timer
    Timer1.Enabled = 0
    StopSound LastMus
    Dim FinalMus
    FinalMus = Hex(MStore)
    If CInt(MStore) < 16 Then
        LastMus = "FD" & "0" & FinalMus & "B"
    Else
        LastMus = "FD" & FinalMus & "B"
    End If
    PlaySound LastMus, -1
End Sub

' Thalamus : Exit in a clean and proper way
Sub Table1_exit()
  Controller.Pause = False
  Controller.Stop
End Sub


Option Explicit
Randomize

' Thalamus 2018-07-24
' Added/Updated "Positional Sound Playback Functions" and "Supporting Ball & Sound Functions"
' Changed UseSolenoids=1 to 2
' Moved solenoids above table1_init
' No special SSF tweaks yet.

Const cGameName = "nineball"  'standard rom
'Const cGameName = "ninebafp"  'free play rom
'Const cGameName = "ninebalb"  'alternate rules rom

On Error Resume Next
ExecuteGlobal GetTextFile("controller.vbs")
If Err Then MsgBox "Can't open controller.vbs"
On Error Goto 0

LoadVPM "01110000","stern.vbs",3.1  'Nine Ball

Dim DesktopMode: DesktopMode = table1.ShowDT

Const UseSolenoids = 2
Const UseGI = 0
Const UseLamps = 0
Const UseSync = 0

' Standard Sounds
Const SSolenoidOn = ""
Const SSolenoidOff = ""
Const SCoin = "fx_coin"

Dim bsTrough, dttop, dtleft, dtright, dtloop, bslsaucer, hiddenvalue

'*********** Desktop/Cabinet settings ************************

If DesktopMode = true Then
	CabinetRailLeft.visible=1
	CabinetRailRight.visible=1
Else
	CabinetRailLeft.visible=0
	CabinetRailRight.visible=0
End If

Sub SolLFlipper(Enabled)
    If Enabled Then
        PlaySound SoundFX("fx_flipperup",DOFContactors), 0, 1, -0.1, 0.25
        LeftFlipper.RotateToEnd
    Else
        PlaySound SoundFX("fx_flipperdown",DOFContactors), 0, 1, -0.1, 0.25
        LeftFlipper.RotateToStart
    End If
End Sub
'
Sub SolRFlipper(Enabled)
    If Enabled Then
        PlaySound SoundFX("fx_flipperup",DOFContactors), 0, 1, 0.1, 0.25
        RightFlipper.RotateToEnd
   Else
        PlaySound SoundFX("fx_flipperdown",DOFContactors), 0, 1, 0.1, 0.25
        RightFlipper.RotateToStart
    End If
End Sub

Sub FlipperTimer_Timer
	'Add flipper, gate and spinner rotations here
	LFLogo1.RotY = LeftFlipper.CurrentAngle
	RFLogo1.RotY = RightFlipper.CurrentAngle
	FlipperLSh.RotZ = LeftFlipper.currentangle
	FlipperRSh.RotZ = RightFlipper.currentangle
	spinner_prim.RotX = sw5.CurrentAngle + 90
	BallShadowUpdate
End Sub

'*** SOLENOIDS ***

SolCallback(2)=		"vpmSolSound""fx_slingshot"","				'1 sLSling
SolCallback(1)=		"vpmSolSound""fx_slingshot"","				'2 sRSling
SolCallback(6)=		"SolKnocker"								'3 sKnocker
SolCallback(7)=		"bsTrough.SolOut"							'4 sOutHole
SolCallback(3)=		"SolRightDropUp"							'5 sRReset
SolCallback(4)=		"SolTopDropUp"								'6 sTReset
'SolCallback(5)=		"B1"									'7 sBumper
SolCallback(8)=		"sVLock"									'8 sVLock
SolCallback(11)=		"dtLeft.SolHit 5,"						'9 sTarget5
SolCallback(12)=	"dtLeft.SolHit 4,"							'10 sTarget4
SolCallback(14)=	"dtLeft.SolHit 2,"							'11 sTarget2
SolCallback(13)=	"dtLeft.SolHit 3,"							'12 sTarget3
SolCallback(9)=	"dtLeft.SolHit 7,"								'13 sTarget7
SolCallback(10)=	"dtLeft.SolHit 6,"							'14 sTarget6
SolCallback(19)=	"vpmNudge.SolGameOn"						'15 sEnabled
SolCallback(15)=	"dtLeft.SolHit 1,"							'16 sTarget1
SolCallback(17)=	"SolLeftDropUp"								'17 sLReset
SolCallback(20)=	"SolLoopDropUp"								'18 sLoop
SolCallback(sLRFlipper) = "SolRFlipper"
SolCallback(sLLFlipper) = "SolLFlipper"

Sub Table1_Init
     With Controller
         .GameName = cGameName
         If Err Then MsgBox "Can't start Game " & cGameName & vbNewLine & Err.Description:Exit Sub
         .SplashInfoLine = ""
         .HandleKeyboard = 0
         .ShowTitle = 0
         .ShowDMDOnly = 1
         .ShowFrame = 0
         .HandleMechanics = False
         .Hidden = 1
         On Error Resume Next
         .Run GetPlayerHWnd
         If Err Then MsgBox Err.Description
         On Error Goto 0
     End With

    PinMAMETimer.Interval = PinMAMEInterval
    PinMAMETimer.Enabled = 1


   	Set bsTrough=New cvpmBallStack
 	with bsTrough
	bsTrough.InitSw 0,33,36,37,0,0,0,0
		.InitKick BallRelease,90,7
		bsTrough.InitExitSnd Soundfx("fx_ballrel",DOFContactors), Soundfx("fx_solenoid",DOFContactors)
		.Balls=3
 	end with

    ' Left Saucer
    Set bslSaucer = New cvpmBallStack
    With bslSaucer
        .InitSaucer sw34,34,180,5
        .KickAngleVar = 2
        .KickForceVar = 3
        .InitExitSnd SoundFX("fx_kicker",DOFContactors), SoundFX("fx_kicker",DOFContactors)
    End With

     ' Nudging
     vpmNudge.TiltSwitch = 7
     vpmNudge.Sensitivity = 3
     vpmNudge.TiltObj = Array(leftslingshot, rightslingshot, sw31)

 	Set dtLeft=New cvpmDropTarget
		dtLeft.InitDrop Array(sw9,sw10,sw11,sw12,sw13,sw14,sw15,sw16),Array(9,10,11,12,13,14,15,16)
		dtLeft.InitSnd SoundFX("fx2_droptarget",DOFContactors),SoundFX("fx2_DTReset",DOFContactors)

 	Set dtRight=New cvpmDropTarget
		dtRight.InitDrop Array(sw22,sw23,sw24),Array(22,23,24)
		dtRight.InitSnd SoundFX("fx2_droptarget",DOFContactors),SoundFX("fx2_DTReset",DOFContactors)

 	Set dtTop=New cvpmDropTarget
		dtTop.InitDrop Array(sw18,sw19,sw20),Array(18,19,20)
		dtTop.InitSnd SoundFX("fx2_droptarget",DOFContactors),SoundFX("fx2_DTReset",DOFContactors)

 	Set dtLoop=New cvpmDropTarget
		dtLoop.InitDrop Array(sw21),Array(21)
		dtLoop.InitSnd SoundFX("fx2_droptarget",DOFContactors),SoundFX("fx2_DTReset",DOFContactors)

''*****Drop Lights Off
	dim xx
'
	For each xx in DTLeftLights: xx.state=0:Next
'	For each xx in DTRightLights: xx.state=0:Next

GILights 1

End Sub

Sub GILights (enabled)
	Dim light
	For each light in GI:light.State = Enabled: Next
End Sub

Sub Table1_KeyDown(ByVal keycode)
	If keycode = LeftTiltKey Then Nudge 90, 2
	If keycode = RightTiltKey Then Nudge 270, 2
	If keycode = CenterTiltKey Then	Nudge 0, 2

	If vpmKeyDown(keycode) Then Exit Sub
	If keycode = PlungerKey Then Plunger.PullBack: PlaySound "fx_plungerpull",0,1,0.25,0.25: 	End If
End Sub

Sub Table1_KeyUp(ByVal keycode)
	If keycode = PlungerKey Then Plunger.Fire: PlaySound "fx_plunger",0,1,0.25,0.25
	If vpmKeyUp(keycode) Then Exit Sub

End Sub

Sub Drain_Hit()
	PlaySound "fx2_drain2",0,1,0,0.25 : bstrough.addball me
End Sub

Sub sw34_Hit:PlaySound "fx_kicker_enter":bslSaucer.AddBall 0:End Sub

'Drop Targets
 Sub Sw9_Dropped:dtLeft.Hit 1 :  LDropE2.state=1 : LDropD4.state=1 : End Sub  :
 Sub Sw10_Dropped:dtLeft.Hit 2 : LDropA1.state=1 : LDropB1.state=1 : End Sub
 Sub Sw11_Dropped:dtLeft.Hit 3 : LDropC3.state=1 : LDropD3.state=1 : LDropE1.state=1 : End Sub
 Sub Sw12_Dropped:dtLeft.Hit 4 : LDropA2.state=1 : LDropB2.state=1 : End Sub
 Sub Sw13_Dropped:dtLeft.Hit 5 : LDropC2.state=1 : LDropD2.state=1 : End Sub
 Sub Sw14_Dropped:dtLeft.Hit 6 : LDropA3.state=1 : LDropB3.state=1 : End Sub
 Sub Sw15_Dropped:dtLeft.Hit 7 : LDropC1.state=1 : LDropD1.state=1 : End Sub
 Sub Sw16_Dropped:dtLeft.Hit 8 : LDropB4.state=1 : End Sub
 Sub Sw18_Dropped:dtTop.Hit 1 : 'GI_14c.state=1 : GI_13c.state=1 :
End Sub
 Sub Sw19_Dropped:dtTop.Hit 2 : 'GI_14b.state=1 : GI_13b.state=1 :
End Sub
 Sub Sw20_Dropped:dtTop.Hit 3 : 'GI_14a.state=1 : GI_13a.state=1 :
End Sub

 Sub Sw21_Dropped:dtLoop.Hit 1 : 'GI_14c.state=1 : GI_13c.state=1 :
End Sub

 Sub Sw22_Dropped:dtRight.Hit 1 : 'GI_14c.state=1 : GI_13c.state=1 :
End Sub
 Sub Sw23_Dropped:dtRight.Hit 2 : 'GI_14b.state=1 : GI_13b.state=1 :
End Sub
 Sub Sw24_Dropped:dtRight.Hit 3 : 'GI_14a.state=1 : GI_13a.state=1 :
End Sub

Sub SolRightDropUp(enabled)
	dim xx
	if enabled then
		dtRight.SolDropUp enabled
		'For each xx in DTRightLights: xx.state=0:Next	FIX
	end if
End Sub

Sub SolLeftDropUp(enabled)
	dim xx
	if enabled then
		dtLeft.SolDropUp enabled
		For each xx in DTLeftLights: xx.state=0:Next
	end if
End Sub

Sub SolTopDropUp(enabled)
	dim xx
	if enabled then
		dtTop.SolDropUp enabled
		'For each xx in DTTopLights: xx.state=0:Next	fix
	end if
End Sub

Sub SolLoopDropUp(enabled)
	dim xx
	if enabled then
		dtLoop.SolDropUp enabled
		'For each xx in DTLoopLights: xx.state=0:Next	fix
	end if
End Sub

'Bumpers

Sub sw31_Hit : vpmTimer.PulseSw 31 : playsound SoundFX("fx2_bumper_1",DOFContactors): End Sub
Sub sw17_Hit : vpmTimer.PulseSw 17 : End Sub

'Wire Triggers
Sub sw25_Hit:Controller.Switch(25)=1 : End Sub 	'Right Outlane
Sub sw25_unHit:Controller.Switch(25)=0:End Sub
Sub sw26_Hit:Controller.Switch(26)=1 : End Sub 	'Left Outlane
Sub sw26_unHit:Controller.Switch(26)=0:End Sub
Sub sw27_Hit:Controller.Switch(27)=1 : End Sub 	'Right Inlane
Sub sw27_unHit:Controller.Switch(27)=0:End Sub
Sub sw28_Hit:Controller.Switch(28)=1 : End Sub 	'Left Inlane
Sub sw28_unHit:Controller.Switch(28)=0:End Sub
Sub sw27_Hit:Controller.Switch(27)=1 : End Sub 	'Right Inlane
Sub sw27_unHit:Controller.Switch(27)=0:End Sub
Sub sw32_Hit:Controller.Switch(32)=1 : End Sub 	'Left Midlane
Sub sw32_unHit:Controller.Switch(32)=0:End Sub
Sub sw38_Hit:Controller.Switch(39)=1 : End Sub	'Mid Lock
Sub sw38_unHit:Controller.Switch(39)=0:End Sub
Sub sw39_Hit:Controller.Switch(39)=1 : End Sub	'Top Lock
Sub sw39_unHit:Controller.Switch(39)=0:End Sub
Sub sw40_Hit:Controller.Switch(40)=1 : End Sub	'Plunger Lane
Sub sw40_unHit:Controller.Switch(40)=0:End Sub
Sub sw4_Hit:Controller.Switch(4)=1 	 : End Sub	'Loop Lane
Sub sw4_unHit:Controller.Switch(4)=0 : End Sub

'Targets
Sub sw35_Hit:vpmTimer.PulseSw (35):End Sub

'Spinners
Sub sw5_Spin : vpmTimer.PulseSw (5) :PlaySound "fx_spinner": End Sub

Sub SolKnocker(Enabled)
	If Enabled Then PlaySound SoundFX("Knocker",DOFKnocker)
End Sub

Sub SVlock(Enabled)
If Enabled Then
	bslSaucer.ExitSol_On
End If
End Sub


'**********Sling Shot Animations
' Rstep and Lstep  are the variables that increment the animation
'****************
Dim RStep, Lstep

Sub RightSlingShot_Slingshot
    PlaySound SoundFX("fx_slingshot",DOFContactors), 0, 1, 0.05, 0.05
	vpmTimer.PulseSw(29)
    RSling.Visible = 0
    RSling1.Visible = 1
    sling1.TransZ = -20
    RStep = 0
    RightSlingShot.TimerEnabled = 1
End Sub

Sub RightSlingShot_Timer
    Select Case RStep
        Case 3:RSLing1.Visible = 0:RSLing2.Visible = 1:sling1.TransZ = -10
        Case 4:RSLing2.Visible = 0:RSLing.Visible = 1:sling1.TransZ = 0:RightSlingShot.TimerEnabled = 0
    End Select
    RStep = RStep + 1
End Sub

Sub LeftSlingShot_Slingshot
    PlaySound SoundFX("fx_slingshot",DOFContactors),0,1,-0.05,0.05
	vpmTimer.PulseSw(30)
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

'******************************
' Diverse Collection Hit Sounds
'******************************

'Sub aBumpers_Hit (idx): PlaySound SoundFX("fx_bumper", DOFContactors), 0, 1, pan(ActiveBall): End Sub
Sub aRollovers_Hit(idx):PlaySound "fx_sensor", 0, Vol(ActiveBall), pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0, AudioFade(ActiveBall):End Sub
Sub aGates_Hit(idx):PlaySound "fx_Gate", 0, Vol(ActiveBall), pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0, AudioFade(ActiveBall):End Sub
Sub aMetals_Hit(idx):PlaySound "fx_MetalHit", 0, Vol(ActiveBall), pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0, AudioFade(ActiveBall):End Sub
Sub aRubber_Bands_Hit(idx):PlaySound "fx_rubber", 0, Vol(ActiveBall), pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0, AudioFade(ActiveBall):End Sub
Sub aRubbers_Hit(idx):PlaySound "fx_postrubber", 0, Vol(ActiveBall), pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0, AudioFade(ActiveBall):End Sub
	Sub aRubber_Posts_Hit(idx):PlaySound "fx_postrubber", 0, Vol(ActiveBall), pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0, AudioFade(ActiveBall):End Sub
	Sub aRubber_Pins_Hit(idx):PlaySound "fx_postrubber", 0, Vol(ActiveBall), pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0, AudioFade(ActiveBall):End Sub
	Sub aYellowPins_Hit(idx):PlaySound "fx_postrubber", 0, Vol(ActiveBall), pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0, AudioFade(ActiveBall):End Sub
Sub aPlastics_Hit(idx):PlaySound "fx_PlasticHit", 0, Vol(ActiveBall), pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0, AudioFade(ActiveBall):End Sub
Sub aWoods_Hit(idx):PlaySound "fx_Woodhit", 0, Vol(ActiveBall), pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0, AudioFade(ActiveBall):End Sub

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
   	NFadeLm 1, l1b
	NFadeLm 2, l2
   	NFadeLm 2, l2b
	NFadeLm 3, l3
   	NFadeLm 3, l3b
	NFadeLm 4, l4
   	NFadeLm 4, l4b
	NFadeLm 5, l5
   	NFadeLm 5, l5b
	NFadeLm 6, l6
   	NFadeLm 6, l6b
	NFadeLm 7, l7
   	NFadeLm 7, l7b
	NFadeLm 8, l8
   	NFadeLm 8, l8b
	NFadeLm 9, l9
   	NFadeLm 9, l9a
	NFadeLm 9, l9b
   	NFadeLm 10, l10
   	NFadeLm 10, l10b
   	NFadeLm 11, l11
   	NFadeLm 11, l11a
   	NFadeLm 11, l11b
   	NFadeLm 11, l11c
	Flash 11, f11
   	NFadeLm 12, l12
   	NFadeLm 12, l12b
   	NFadeLm 14, l14
   	NFadeLm 14, l14b
   	NFadeLm 15, l15
   	NFadeLm 15, l15b
   	NFadeLm 17, l17
   	NFadeLm 17, l17b
   	NFadeLm 18, l18
   	NFadeLm 18, l18b
   	NFadeLm 19, l19
   	NFadeLm 19, l19b
   	NFadeLm 20, l20
   	NFadeLm 20, l20b
   	NFadeLm 21, l21
   	NFadeLm 21, l21b
   	NFadeLm 22, l22
   	NFadeLm 22, l22b
   	NFadeLm 23, l23
   	NFadeLm 23, l23b
   	NFadeLm 24, l24
   	NFadeLm 24, l24b
   	NFadeLm 25, l25
   	NFadeLm 25, l25b
   	NFadeLm 26, l26
   	NFadeLm 26, l26b
   	NFadeLm 27, l27
   	NFadeLm 27, l27b
   	NFadeLm 28, l28
   	NFadeLm 28, l28b
   	NFadeLm 29, l29
   	NFadeLm 29, l29a
   	NFadeLm 29, l29b
   	NFadeLm 29, l29c
	Flash 29, f29
   	NFadeLm 30, l30
   	NFadeLm 30, l30a
   	NFadeLm 30, l30b
   	NFadeLm 30, l30c
	Flash	30, f30
	NFadeObjm 30, Bumper2bottom, "deadbumperbottom", "deadbumperbottomoff"
	NFadeObjm 30, Bumper2cap, "deadbumpercap", "deadbumpercapoff"
	NFadeObjm 30, bumper2shaft, "deadbumperbase", "deadbumperbaseoff"
   	NFadeLm 31, l31
   	NFadeLm 31, l31b
   	NFadeLm 33, l33
   	NFadeLm 33, l33b
   	NFadeLm 34, l34
   	NFadeLm 34, l34b
   	NFadeLm 35, l35
   	NFadeLm 35, l35b
   	NFadeLm 36, l36
   	NFadeLm 36, l36b
   	NFadeLm 37, l37
   	NFadeLm 37, l37b
   	NFadeLm 38, l38
   	NFadeLm 38, l38b
   	NFadeLm 39, l39
   	NFadeLm 39, l39b
   	NFadeLm 40, l40
   	NFadeLm 40, l40b
   	NFadeLm 41, l41
   	NFadeLm 41, l41b
   	NFadeLm 42, l42
   	NFadeLm 42, l42b
   	NFadeLm 43, l43
   	NFadeLm 43, l43b
   	NFadeLm 44, l44
   	NFadeLm 44, l44b
   	NFadeLm 46, l46
   	NFadeLm 46, l46b
   	NFadeLm 47, l47
   	NFadeLm 47, l47b
   	NFadeLm 49, l49
   	NFadeLm 49, l49b
   	NFadeLm 50, l50
   	NFadeLm 50, l50b
   	NFadeLm 51, l51
   	NFadeLm 51, l51b
   	NFadeLm 52, l52
   	NFadeLm 52, l52b
   	NFadeLm 53, l53
   	NFadeLm 53, l53b
   	NFadeLm 54, l54
   	NFadeLm 54, l54b
   	NFadeLm 55, l55
   	NFadeLm 55, l55b
   	NFadeLm 56, l56
   	NFadeLm 56, l56b
	NFadeLm 57, L57
	NFadeLm 57, L57a
	NFadeLm 57, L57b
	NFadeLm 57, L57c
	NFadeLm 58, L58
	NFadeLm 58, L58b
	NFadeLm 59, L59
	NFadeLm 59, L59b
   	NFadeLm 60, l60
   	NFadeLm 60, l60b
   	NFadeLm 62, l62
   	NFadeLm 62, l62b
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

'*****************************************
'	Ball Shadow
'*****************************************

Dim BallShadow
BallShadow = Array (BallShadow1, BallShadow1, BallShadow2, BallShadow3, BallShadow4, BallShadow5)


Sub BallShadowUpdate()
    Dim BOT, b, shadowZ
    BOT = GetBalls

	' render the shadow for each ball
    For b = 0 to UBound(BOT)
		If BOT(b).X < Table1.Width/2 Then
			BallShadow(b).X = ((BOT(b).X) - (Ballsize/6) + ((BOT(b).X - (Table1.Width/2))/7)) + 10
		Else
			BallShadow(b).X = ((BOT(b).X) + (Ballsize/6) + ((BOT(b).X - (Table1.Width/2))/7)) - 10
		End If

		If BOT(b).X > 875 AND BOT(b).Y > 935 Then shadowZ = BOT(b).Z : BallShadow(b).X = BOT(b).X Else shadowZ = 1

			BallShadow(b).Y = BOT(b).Y + 20
			BallShadow(b).Z = shadowZ
		If BOT(b).Z > 20 Then
			BallShadow(b).visible = 1
		Else
			BallShadow(b).visible = 0
		End If
	Next
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

'*** DIPS *** Inkochnito ***
Sub editDips
	Dim vpmDips : Set vpmDips = New cvpmDips
	With vpmDips
		.AddForm 700,460,"Nine Ball - DIP switches"
		.AddChk 2,10,105,Array("Match feature",&H00100000)'dip 21
		.AddChk 120,10,105,Array("Credits display",&H00080000)'dip 20
		.AddChk 240,10,110,Array("Background sound",&H00000080)'dip 8
		.AddChk 350,10,110,Array("Add-a-ball memory",&H00000010)'dip 5
		.AddFrame 2,30,100,"Maximum credits",&H00060000,Array("10 credits",0,"15 credits",&H00020000,"25 credits",&H00040000,"40 credits",&H00060000)'dip 18&19
		.AddFrame 2,105,100,"High game to date",49152,Array("points",0,"1 free game",&H00004000,"2 free games",32768,"3 free games",49152)'dip 15&16
		.AddFrame 120,30,100,"Special limit",&HC0000000,Array("1 per ball",0,"more per ball",&H80000000)'dip 32
		.AddFrame 120,81,100,"Balls per game",&H00000040,Array("3 balls",0,"5 balls",&H00000040)'dip 7
		.AddFrame 120,132,100,"High score feature",&H00000020,Array("extra ball",0,"replay",&H00000020)'dip 6
		.AddFrame 2,182,220,"Percentage spot number target lite on",&H00001000,Array("conservative",0,"liberal",&H00001000)'dip 13
		.AddFrame 2,232,220,"Multiplier advances on",&H0002000,Array("both 3 bank targets",0,"one 3 bank targets",&H00002000)'dip 14
		.AddFrame 2,282,220,"Advance spinner value on",&H00010000,Array("3 bank outside targets",0,"3 bank center target",&H00010000)'dip 17
		.AddFrame 240,30,220,"Special - WOW award",&H00E00000,Array("no award - 40,000 points",0,"90,000 points - extra ball",&H00800000,"130,000 points - 70,000 points",&H00400000,"130,000 points - extra ball",&H00C00000,"extra ball (3 if mem is on) - 40,000 points",&H00200000,"extra ball (3 if mem is on) - 70,000 points",&H00A00000,"replay (2 if limit is more) - 70,000 points",&H00600000,"replay (2 if limit is more) - extra ball",&H00E00000)'dip 22&23&24
		.AddFrame 240,182,220,"8 bank WOW lites when",&H10000000,Array("hitting 9 ball target when lit",0,"hitting 8 ball target when lit",&H10000000)'dip 29
		.AddFrame 240,232,220,"8 bank super bonus lites when",&H20000000,Array("hitting 9 ball target when lit",0,"hitting 8 ball target when lit",&H20000000)'dip 30
		.AddFrame 240,282,220,"3 bank WOW lites",&H40000000,Array("on 7X",0,"on 6X or 7X",&H40000000)'dip 31
		.AddLabel 50,335,300,20,"After hitting OK, press F3 to reset game with new settings."
		.ViewDips
	End With
End Sub
Set vpmShowDips = GetRef("editDips")

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


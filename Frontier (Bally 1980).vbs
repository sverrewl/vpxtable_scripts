'Frontier (Bally 1980) v2.0
' by bord

Option Explicit
Randomize

' Thalamus 2018-07-23
' Added/Updated "Positional Sound Playback Functions" and "Supporting Ball & Sound Functions"
' Changed UseSolenoids=1 to 2
' Moved solenoids above table1_init
' Thalamus 2018-08-11 : Improved directional sounds
' Options
' Volume devided by - lower gets higher sound

Const VolDiv = 2000

Const VolBump   = 2    ' Bumpers multiplier.
Const VolRol    = 1    ' Rollovers volume multiplier.
Const VolGates  = 1    ' Gates volume multiplier.
Const VolMetals = 1    ' Metals volume multiplier.
Const VolRB     = 1    ' Rubber bands multiplier.
Const VolRH     = 1    ' Rubber hits multiplier.
Const VolRPo    = 1    ' Rubber posts multiplier.
Const VolRPi    = 1    ' Rubber pins multiplier.
Const VolPlast  = 1    ' Plastics multiplier.
Const VolTarg   = 1    ' Targets multiplier.
Const VolWood   = 1    ' Woods multiplier.

Const VolSpin   = 1.5  ' Spinners volume.
Const VolFlip   = 1    ' Flipper volume.

Const cGameName = "frontier"

On Error Resume Next
ExecuteGlobal GetTextFile("controller.vbs")
If Err Then MsgBox "Can't open controller.vbs"
On Error Goto 0

LoadVPM "01000100", "Bally.VBS", 3.02  'Frontier
Dim DesktopMode: DesktopMode = table1.ShowDT

'************************************************
'************************************************
'************************************************
'************************************************
'************************************************
Const UseSolenoids = 2
Const UseLamps = True
Const UseSync = False

' Standard Sounds
Const SSolenoidOn = "fx_solenoid"
Const SSolenoidOff = "fx_solenoidoff"
Const SCoin = "fx_coin"

Dim bsTrough, dtll, dtrr, hiddenvalue, objekt

If DesktopMode = True Then 'Show Desktop components
	CabinetRailLeft.visible=1
	CabinetRailRight.visible=1
	hiddenvalue=0
  Else
	CabinetRailLeft.visible=0
	CabinetRailRight.visible=0
    For each objekt in aReels
        objekt.Visible = 0
    Next
	hiddenvalue=1
End if

Sub SolLFlipper(Enabled)
    If Enabled Then
        PlaySoundAtVol SoundFX("fx_flipperup",DOFContactors), LeftFlipper, VolFlip
        LeftFlipper.RotateToEnd
    Else
        PlaySoundAtVol SoundFX("fx_flipperdown",DOFContactors), LeftFlipper, VolFlip
        LeftFlipper.RotateToStart
    End If
End Sub

Sub SolRFlipper(Enabled)
    If Enabled Then
        PlaySoundAtVol SoundFX("fx_flipperup",DOFContactors), RightFlipper, VolFlip
        RightFlipper.RotateToEnd
   Else
        PlaySoundAtVol SoundFX("fx_flipperdown",DOFContactors), RightFlipper, VolFlip
        RightFlipper.RotateToStart
    End If
End Sub

Sub FlipperTimer_Timer
	'Add flipper, gate and spinner rotations here
	LFLogo.RotY = LeftFlipper.CurrentAngle
	RFLogo.RotY = RightFlipper.CurrentAngle
End Sub

'************************************************
' Solenoids
'************************************************
SolCallback(1) = 	"vpmSolSound ""fx_slingshot"","		'sLSling
SolCallback(2) = 	"vpmSolSound ""fx_slingshot"","		'sRSling
'SolCallback(3) = 	""
'SolCallback(4) = 	""
'SolCallback(5) = 	""
SolCallback(6) = 	"SolKnocker "						'sKnocker
SolCallback(7) = 	"bsTrough.SolOut"					'sBallRelease
SolCallback(8) = 	"SolInlineTargetReset"				'sInlineTargets
SolCallback(9) = 	"SolRightTargetReset"				'sRightTargets
SolCallback(10) = 	"SolFrontierSaucer"					'sSaucer
'SolCallback(11) = 	"vpmSolSound "						'sLJet
'SolCallback(12) = 	"vpmSolSound "						'sRJet
'SolCallback(13) = 	"vpmSolSound "						'sDJet
'SolCallback(14) = 	""
'SolCallback(15) = 	""
'SolCallback(16) = 	""
SolCallback(17) = 	"vpmSolGate Gate2,true,"			'sGate
SolCallback(18) = 	""
SolCallback(19) = 	"vpmNudge.SolGameOn"				'sEnable
'SolCallback(20) = 	""

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


	'Trough

   	Set bsTrough=New cvpmBallStack
 	with bsTrough
		.InitSw 0,8,0,0,0,0,0,0
		.InitKick BallRelease,90,7
		.InitExitSnd Soundfx("fx_ballrel",DOFContactors), Soundfx("fx_solenoid",DOFContactors)
		.Balls=1
 	end with

     ' Nudging
     vpmNudge.TiltSwitch = 7
     vpmNudge.Sensitivity = 3
     vpmNudge.TiltObj = Array(leftslingshot, rightslingshot, sw38, sw39, sw40)

 	Set dtLL=New cvpmDropTarget
		dtLL.InitDrop Array(sw21,sw22,sw23),Array(21,22,23)
		dtLL.InitSnd SoundFX("fx2_droptarget",DOFContactors),SoundFX("fx2_DTReset",DOFContactors)

 	Set dtRR=New cvpmDropTarget
		dtRR.InitDrop Array(sw33,sw34,sw35),Array(33,34,35)
		dtRR.InitSnd SoundFX("fx2_droptarget2",DOFContactors),SoundFX("fx2_DTReset",DOFContactors)

'*****Drop Lights Off
	dim xx

	For each xx in DTLeftLights: xx.state=0:Next
	For each xx in DTRightLights: xx.state=0:Next

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
	If keycode = PlungerKey Then Plunger.PullBack: PlaySoundAtVol "fx_plungerpull",Plunger,1
End Sub

Sub Table1_KeyUp(ByVal keycode)
	If keycode = PlungerKey Then Plunger.Fire: PlaySoundAtVol "fx_plunger",Plunger,1
	If vpmKeyUp(keycode) Then Exit Sub

End Sub

Sub ShooterLane_Hit
End Sub

Sub BallRelease_UnHit
End Sub

Sub Drain_Hit()
	PlaySoundAtVol "fx2_drain2",Drain,1 : bstrough.addball me
End Sub

Sub sw12_Hit   : Controller.Switch(12) = True : PlaySoundAtVol "fx_hole-enter",sw12,1:End Sub


Sub SolFrontierSaucer(enabled)
	if enabled then
		PlaySound "fx_popper"
		Controller.Switch(12) = false
		sw12.Kick  -10, 15 + 5 * Rnd
	end if
End Sub

'Drop Targets
 Sub Sw21_Dropped:dtLL.Hit 1 : GI_12a.state=1 : End Sub
 Sub Sw22_Dropped:dtLL.Hit 2 : GI_12b.state=1 : End Sub
 Sub Sw23_Dropped:dtLL.Hit 3 : GI_12c.state=1 : End Sub

 Sub Sw33_Dropped:dtRR.Hit 1 : GI_14c.state=1 : GI_13c.state=1 : End Sub
 Sub Sw34_Dropped:dtRR.Hit 2 : GI_14b.state=1 : GI_13b.state=1 : End Sub
 Sub Sw35_Dropped:dtRR.Hit 3 : GI_14a.state=1 : GI_13a.state=1 : End Sub

Sub SolRightTargetReset(enabled)
	dim xx
	if enabled then
		dtRR.SolDropUp enabled
		For each xx in DTRightLights: xx.state=0:Next
	end if
End Sub

Sub SolInlineTargetReset(enabled)
	dim xx
	if enabled then
		dtLL.SolDropUp enabled
		For each xx in DTLeftLights: xx.state=0:Next
	end if
End Sub

'Bumpers

Sub sw38_Hit : vpmTimer.PulseSw 38 : PlaySoundAtVol SoundFX("fx2_bumper_1",DOFContactors),sw38,VolSpin: End Sub
Sub sw39_Hit : vpmTimer.PulseSw 39 : PlaySoundAtVol SoundFX("fx2_bumper_2",DOFContactors),sw39,VolSpin: End Sub
Sub sw40_Hit : vpmTimer.PulseSw 40 : PlaySoundAtVol SoundFX("fx2_bumper_3",DOFContactors),sw40,VolSpin: End Sub

dim ltopstep, lltopstep, rtopstep, RL2step

'leaf switches
Sub sw1a1_hit
	vpmTimer.PulseSw(1)
	RubbertopL.Visible = 0
    RubbertopL1.Visible = 1
    ltopstep = 0
    sw1a1.TimerEnabled = 1
End Sub

Sub sw1a1_Timer
    Select Case ltopstep
        Case 3:RubbertopL1.Visible = 0:RubbertopL2.Visible = 1
        Case 4:RubbertopL2.Visible = 0:RubbertopL.Visible = 1:sw1a1.TimerEnabled = 0
    End Select
    ltopstep = ltopstep + 1
End Sub

Sub sw1a2_hit
	vpmTimer.PulseSw(1)
	RubbertopLL.Visible = 0
    RubbertopLL1.Visible = 1
    lltopstep = 0
    sw1a2.TimerEnabled = 1
End Sub

Sub sw1a2_Timer
    Select Case lltopstep
        Case 3:RubbertopLL1.Visible = 0:RubbertopLL2.Visible = 1
        Case 4:RubbertopLL2.Visible = 0:RubbertopLL.Visible = 1:sw1a2.TimerEnabled = 0
    End Select
    lltopstep = lltopstep + 1
End Sub

Sub sw1a3_hit
	vpmTimer.PulseSw(1)
	RubberL2_.Visible = 0
    RubberL2_1.Visible = 1
    RL2step = 0
    sw1a3.TimerEnabled = 1
End Sub

Sub sw1a3_Timer
    Select Case RL2step
        Case 3:RubberL2_1.Visible = 0:RubberL2_2.Visible = 1
        Case 4:RubberL2_2.Visible = 0:RubberL2_.Visible = 1:sw1a3.TimerEnabled = 0
    End Select
    RL2step = RL2step + 1
End Sub

Sub sw1a4_hit
	vpmTimer.PulseSw(1)
	RubbertopR.Visible = 0
    RubbertopR1.Visible = 1
    rtopstep = 0
    sw1a4.TimerEnabled = 1
End Sub

Sub sw1a4_Timer
    Select Case rtopstep
        Case 3:RubbertopR1.Visible = 0:RubbertopR2.Visible = 1
        Case 4:RubbertopR2.Visible = 0:RubbertopR.Visible = 1:sw1a4.TimerEnabled = 0
    End Select
    rtopstep = rtopstep + 1
End Sub

'Wire Triggers
Sub SW2_Hit:Controller.Switch(2)=1 : End Sub 	'Right Outlane
Sub SW2_unHit:Controller.Switch(2)=0:End Sub
Sub SW3_Hit:Controller.Switch(3)=1 : End Sub 	'Right Inlane
Sub SW3_unHit:Controller.Switch(3)=0:End Sub
Sub SW4_Hit:Controller.Switch(4)=1 : End Sub 	'Left Inlane
Sub SW4_unHit:Controller.Switch(4)=0:End Sub
Sub SW5_Hit:Controller.Switch(5)=1 : End Sub 	'Left Outlane
Sub SW5_unHit:Controller.Switch(5)=0:End Sub
Sub SW13_Hit:Controller.Switch(13)=1 : End Sub 	'A
Sub SW13_unHit:Controller.Switch(13)=0:End Sub
Sub SW14_Hit:Controller.Switch(14)=1 : End Sub 	'B
Sub SW14_unHit:Controller.Switch(14)=0:End Sub
Sub SW15_Hit:Controller.Switch(15)=1 : End Sub 	'C
Sub SW15_unHit:Controller.Switch(15)=0:End Sub
Sub SW25_Hit:Controller.Switch(25)=1 : End Sub 	'Falls
Sub SW25_unHit:Controller.Switch(25)=0:End Sub

'Spinners
Sub sw27_Spin : vpmTimer.PulseSw (27) :PlaySoundAtVol "fx_spinner",sw27,VolSpin: End Sub

'Targets
Sub sw24_Hit:vpmTimer.PulseSw (24):End Sub
Sub sw29_Hit:vpmTimer.PulseSw (29):End Sub
Sub sw30_Hit:vpmTimer.PulseSw (30):End Sub
Sub sw31_Hit:vpmTimer.PulseSw (31):End Sub

Sub SolKnocker(Enabled)
	If Enabled Then PlaySound SoundFX("Knocker",DOFKnocker)
End Sub

'**********Sling Shot Animations
' Rstep and Lstep  are the variables that increment the animation
'****************
Dim RStep, Lstep, RL1step, RL3step, Tdropstep, Bdropstep,  rrtopstep, RR2step, RR3step

Sub RightSlingShot_Slingshot
    PlaySound SoundFX("fx_slingshot",DOFContactors), 0, 1, 0.05, 0.05
	vpmtimer.PulseSw(36)
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
    PlaySoundAtVol SoundFX("fx_slingshot",DOFContactors), sling2, 1
	  vpmtimer.pulsesw(37)
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

Sub RubberL1Sling_hit
	RubberL1_.Visible = 0
    RubberL1_1.Visible = 1
    RL1step = 0
    RubberL1Sling.TimerEnabled = 1
End Sub

Sub RubberL1Sling_Timer
    Select Case RL1step
        Case 3:RubberL1_1.Visible = 0:RubberL1_2.Visible = 1
        Case 4:RubberL1_2.Visible = 0:RubberL1_.Visible = 1:RubberL1Sling.TimerEnabled = 0
    End Select
    RL1step = RL1step + 1
End Sub

Sub RubberL3Sling_hit
	RubberL3_.Visible = 0
    RubberL3_1.Visible = 1
    RL3step = 0
    RubberL3Sling.TimerEnabled = 1
End Sub

Sub RubberL3Sling_Timer
    Select Case RL3step
        Case 3:RubberL3_1.Visible = 0:RubberL3_2.Visible = 1
        Case 4:RubberL3_2.Visible = 0:RubberL3_.Visible = 1:RubberL3Sling.TimerEnabled = 0
    End Select
    RL3step = RL3step + 1
End Sub

Sub RubberRRSling_hit
	RubberRR.Visible = 0
    RubberRR1.Visible = 1
    rrtopstep = 0
    RubberL3Sling.TimerEnabled = 1
End Sub

Sub RubberRRSling_Timer
    Select Case rrtopstep
        Case 3:RubberRR1.Visible = 0:RubberRR2.Visible = 1
        Case 4:RubberRR2.Visible = 0:RubberRR.Visible = 1:RubberRRSling.TimerEnabled = 0
    End Select
    rrtopstep = rrtopstep + 1
End Sub

Sub RubberR2Sling_hit
	RubberR3_.Visible = 0
    RubberR3_1.Visible = 1
    rr2step = 0
    RubberR2Sling.TimerEnabled = 1
End Sub

Sub RubberR2Sling_Timer
    Select Case rr2step
        Case 3:RubberR3_1.Visible = 0:RubberR3_2.Visible = 1
        Case 4:RubberR3_2.Visible = 0:RubberR3_.Visible = 1:RubberR2Sling.TimerEnabled = 0
    End Select
    rr2step = rr2step + 1
End Sub

Sub RubberR3Sling_hit
	RubberR3_.Visible = 0
    RubberR3_3.Visible = 1
    rr3step = 0
    RubberR3Sling.TimerEnabled = 1
End Sub

Sub RubberR3Sling_Timer
    Select Case rr2step
        Case 3:RubberR3_3.Visible = 0:RubberR3_4.Visible = 1
        Case 4:RubberR3_4.Visible = 0:RubberR3_.Visible = 1:RubberR3Sling.TimerEnabled = 0
    End Select
    rr3step = rr3step + 1
End Sub

Sub TDropSling_hit
	RubberDrop.Visible = 0
    RubberDrop1.Visible = 1
    tdropstep = 0
    TDropSling.TimerEnabled = 1
End Sub

Sub TDropSling_Timer
    Select Case tdropstep
        Case 3:RubberDrop1.Visible = 0:RubberDrop2.Visible = 1
        Case 4:RubberDrop2.Visible = 0:RubberDrop.Visible = 1:TDropSling.TimerEnabled = 0
    End Select
    tdropstep = tdropstep + 1
End Sub

Sub BDropSling_hit
	RubberDrop.Visible = 0
    RubberDrop3.Visible = 1
    bdropstep = 0
    BDropSling.TimerEnabled = 1
End Sub

Sub BDropSling_Timer
    Select Case bdropstep
        Case 3:RubberDrop3.Visible = 0:RubberDrop4.Visible = 1
        Case 4:RubberDrop4.Visible = 0:RubberDrop.Visible = 1:BDropSling.TimerEnabled = 0
    End Select
    bdropstep = bdropstep + 1
End Sub

'******************************
' Diverse Collection Hit Sounds
'******************************

' 'Sub aBumpers_Hit (idx): PlaySound SoundFX("fx_bumper", DOFContactors), 0, 1, pan(ActiveBall): End Sub
' Sub aRollovers_Hit(idx):PlaySound "fx_sensor", 0, Vol(ActiveBall), pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0, AudioFade(ActiveBall):End Sub
' Sub aGates_Hit(idx):PlaySound "fx_Gate", 0, Vol(ActiveBall), pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0, AudioFade(ActiveBall):End Sub
' Sub aMetals_Hit(idx):PlaySound "fx_MetalHit", 0, Vol(ActiveBall), pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0, AudioFade(ActiveBall):End Sub
' Sub aRubber_Bands_Hit(idx):PlaySound "fx_rubber", 0, Vol(ActiveBall), pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0, AudioFade(ActiveBall):End Sub
' Sub aRubbers_Hit(idx):PlaySound "fx_postrubber", 0, Vol(ActiveBall), pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0, AudioFade(ActiveBall):End Sub
' Sub aRubber_Posts_Hit(idx):PlaySound "fx_postrubber", 0, Vol(ActiveBall), pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0, AudioFade(ActiveBall):End Sub
' Sub aRubber_Pins_Hit(idx):PlaySound "fx_postrubber", 0, Vol(ActiveBall), pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0, AudioFade(ActiveBall):End Sub
' Sub aPlastics_Hit(idx):PlaySound "fx_PlasticHit", 0, Vol(ActiveBall), pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0, AudioFade(ActiveBall):End Sub
' Sub aTargets_Hit(idx):PlaySound "fx2_target", 0, Vol(ActiveBall), pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0, AudioFade(ActiveBall):End Sub
' Sub aWoods_Hit(idx):PlaySound "fx_Woodhit", 0, Vol(ActiveBall), pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0, AudioFade(ActiveBall):End Sub

Sub aRollovers_Hit(idx):PlaySound "fx_sensor", 0, Vol(ActiveBall)*VolRol, pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0, AudioFade(ActiveBall):End Sub
Sub aGates_Hit(idx):PlaySound "fx_Gate", 0, Vol(ActiveBall)*VolGates, pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0, AudioFade(ActiveBall):End Sub
Sub aMetals_Hit(idx):PlaySound "fx_MetalHit", 0, Vol(ActiveBall)*VolMetals, pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0, AudioFade(ActiveBall):End Sub
Sub aRubber_Bands_Hit(idx):PlaySound "fx_rubber", 0, Vol(ActiveBall)*VolRB, pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0, AudioFade(ActiveBall):End Sub
Sub aRubbers_Hit(idx):PlaySound "fx_postrubber", 0, Vol(ActiveBall)*VolRH, pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0, AudioFade(ActiveBall):End Sub
Sub aRubber_Posts_Hit(idx):PlaySound "fx_postrubber", 0, Vol(ActiveBall)*VolRPo, pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0, AudioFade(ActiveBall):End Sub
Sub aRubber_Pins_Hit(idx):PlaySound "fx_postrubber", 0, Vol(ActiveBall)*VolRPi, pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0, AudioFade(ActiveBall):End Sub
Sub aPlastics_Hit(idx):PlaySound "fx_PlasticHit", 0, Vol(ActiveBall)*VolPlast, pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0, AudioFade(ActiveBall):End Sub
Sub aTargets_Hit(idx):PlaySound "fx2_target", 0, Vol(ActiveBall)*VolTarg, pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0, AudioFade(ActiveBall):End Sub
Sub aWoods_Hit(idx):PlaySound "fx_Woodhit", 0, Vol(ActiveBall)*VolWood, pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0, AudioFade(ActiveBall):End Sub

'-------------------------------------
' Map lights into array
' Set unmapped lamps to Nothing
'-------------------------------------
Set Lights(1)  = l1 'LightSnake1
Set Lights(2)  = l2 'LightHawk1
Set Lights(3)  = l3 'LightCougar1
Set Lights(4)  = l4 'LightWolf1
Set Lights(5)  = l5 'LightBear1
Set Lights(6)  = l6 'LightA
Set Lights(7)  = l7 'LightBonus10K
Set Lights(8)  = l8 'LightBonus50K
Set Lights(9)  = l9 'LightLeftTarget
Set Lights(10) = BumperLeft 'BumperLeft (L10 and BumperLight2)
Set Lights(11) = l11 'LightExtraBall
Set Lights(12) = l12 'LightLeftReturnLane
'Set Lights(13) = LightBallInPlay
Set Lights(14) = l14 'LightDenTarget2X
Set Lights(15) = l15 'LightGateOpen
'Set Lights(16) = unused
Set Lights(17) = l17 'LightSnake2
Set Lights(18) = l18 'LightHawk2
Set Lights(19) = l19 'LightCougar2
Set Lights(20) = l20 'LightWolf2
Set Lights(21) = l21 'LightBear2
Set Lights(22) = l22 'LightB
Set Lights(23) = l23 'LightBonus20K
Set Lights(24) = l24 'LightBonus60K
Set Lights(25) = l25 'LightCenterTarget
Set Lights(26) = BumperCenter 'BumperCenter (L26 and BumperLight3)
'Set Lights(27) = LightMatch
Set Lights(28) = l28 'LightRightReturnLane
'Set Lights(29) = LightHighScore
Set Lights(30) = l30 'LightDenTarget3X
Set Lights(31) = l31 'LightDen2X
'Set Lights(32) = unused
Set Lights(33) = l33 'LightSnake3
Set Lights(34) = l34 'LightHawk3
Set Lights(35) = l35 'LightCougar3
Set Lights(36) = l36 'LightWolf3
Set Lights(37) = l37 'LightBear3
Set Lights(38) = l38 'LightC
Set Lights(39) = l39 'LightBonus30K
Set Lights(40) = l40 'LightSpinner
Set Lights(41) = l41 'LightRightTarget
Set Lights(42) = BumperRight 'BumperRight L42 and BumperLight1)
'Set Lights(43) = LightShootAgain
Set Lights(44) = l44 'LightLeftLaneSpecial
'Set Lights(45) = LightGameOver
Set Lights(46) = l46 'LightDenTarget4X
Set Lights(47) = l47 'LightDen3X
'Set Lights(48) = unused
Set Lights(49) = l49 'LightABCExtraBall
Set Lights(50) = l50 'LightABCSpecial
Set Lights(51) = l51 'LightDen45K
Set Lights(52) = l52 'LightBonus4X
Set Lights(53) = l53 'LightBonusSpecial
Set Lights(54) = l54 'LightGrizzly4X
Set Lights(55) = l55 'LightBonus40K
Set Lights(56) = l56 'LightGrizzlyPoints
Set Lights(57) = l57 'LightBonus2X
Set Lights(58) = l58 'LightBonus3X
Set Lights(59) = l59 'LightCredit
Set Lights(60) = l60 'LightRightLaneSpecial
'Set Lights(61) = LightTilt
Set Lights(62) = l62 'LightDenTargetSpecial
Set Lights(63) = l63'LightDen4X

'Set Lights(10) = GILeft		'PF_Left
'Set Lights(42) = GIRight	'PF_Right
'Set Lights(26) = GICenter	'PF_Center

'---------------------------------------------------------------
' Edit the dip switches
'---------------------------------------------------------------
Sub editDips
	Dim vpmDips : Set vpmDips = New cvpmDips
	with vpmDips
		.AddForm 315, 370, "Frontier DIP Switch Settings"
		.AddFrame 2, 10, 140, "Balls per game", &HC0000000, _
				Array("2", &HC0000000, "3", 0, "4", &H80000000, "5", &H40000000)
		.AddFrame 160, 10, 140, "Max Credits", &H03000000, _
				Array("10", 0, "15", &H01000000, "25", &H02000000, "40", &H03000000)
		.AddFrame 2, 90, 140, "A && C Lanes are", &H00000020, _
				Array("Separate", 0, "Linked", &H00000020)
		.AddFrame 160, 90, 140, "Frontier Bonus Max", &H00400000, _
				Array("60,000", 0, "110,000", &H00400000)
		.AddFrame 2, 140, 140, "Outlane Special lights", &H00000080, _
				Array("Alternate", 0, "Both light", &H00000080)
		.AddFrame 160, 140, 140, "Outlane Special lights when", &H00002000, _
				Array("3X lights", 0, "4X lights", &H00002000)
		.AddFrame 2, 190, 140, "Feeder Lane Special lights", &H20000000, _
				Array("Alternate", 0, "Both light", &H20000000)
		.AddChk 160, 200, 148, Array("Credit display", &H04000000)
		.AddChk 160, 220, 148, Array("Match", &H08000000)
		.AddChk 7, 250, 300, Array("A-B-C Extra Ball light on at start of game", &H00000040)
		.AddChk 7, 270, 300, Array("A-B-C lane memory", &H00004000)
		'Have to use 32768 because vbscript treats &H00008000 as -32768
		.AddChk 7, 290, 300, Array("Outhole collects Frontier Bonus", 32768)
		.AddChk 7, 310, 300, Array("Gate memory", &H00800000)
		.AddLabel 7, 340, 300, 20, "After hitting OK, press F3 to reset game with new settings."
		.ViewDips
	end with
End Sub
Set vpmShowDips = GetRef("editDips")

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

Const tnob = 20 ' total number of balls
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


   Option Explicit
   Randomize

' Thalamus 2018-07-24
' Added/Updated "Positional Sound Playback Functions" and "Supporting Ball & Sound Functions"
' No special SSF tweaks yet.
' Added InitVpmFFlipsSAM

Const Inteceptor = 0 'Replace yellow Showroom Car with Mad Max Inteceptor
Const Flasher_Halos = 1 'Rendber Halos around flashers.
Const Tacho_Mod = 1 'Render spinner activated tacho on right ramp.
Const Erratic_Scoop = 1 'Make the Mustang scoop / saucer behave more organically.
Const Car_Color = 9 'Turntable car color (Select from list below)

'1 Light Blue
'2 Dark Blue
'3 White
'4 Green
'5 Orange
'6 Randr
'7 Red
'8 Purple
'9 Yellow

If Car_Color = 1 Then
Primitive3.image = "blue1"
End If

If Car_Color = 2 Then
Primitive3.image = "blue2"
End If

If Car_Color = 3 Then
Primitive3.image = "diamond_white"
End If

If Car_Color = 4 Then
Primitive3.image = "green"
End If

If Car_Color = 5 Then
Primitive3.image = "orange"
End If

If Car_Color = 6 Then
Primitive3.image = "randrs_car"
End If

If Car_Color = 7 Then
Primitive3.image = "red"
End If

If Car_Color = 8 Then
Primitive3.image = "purple"
End If

If Car_Color = 9 Then
Primitive3.image = "yellow2"
End If


Dim VarHidden, UseVPMDMD

If Table.ShowDT = true then
L104.bulbhaloheight = 260
L104a.bulbhaloheight = 260
L103.bulbhaloheight = 186
L103a.bulbhaloheight = 186
L101.bulbhaloheight = 258
L101a.bulbhaloheight = 258
L102.bulbhaloheight = 280
L102a.bulbhaloheight = 280
L107.bulbhaloheight = 250
L107a.bulbhaloheight = 250
L106.bulbhaloheight = 270
L106a.bulbhaloheight = 270
L105.bulbhaloheight = 250
L105a.bulbhaloheight = 250
xm.bulbhaloheight = 250
xu.bulbhaloheight = 250
xs.bulbhaloheight = 250
xt.bulbhaloheight = 250
xa.bulbhaloheight = 250
xn.bulbhaloheight = 250
xg.bulbhaloheight = 250
Ramp16.visible = 1
Ramp15.visible = 1
UseVPMDMD = True
VarHidden = 1
else
UseVPMDMD = False
VarHidden = 0
Ramp16.visible = 0
Ramp15.visible = 0
End If

Const UseVPMModSol = 1

   LoadVPM "01560000", "sam.VBS", 3.10

	 vpmflips.Delay = 100 ' Thalamus - trying to get rid of dying flippers.

     Sub LoadVPM(VPMver, VBSfile, VBSver)
       On Error Resume Next
       If ScriptEngineMajorVersion < 5 Then MsgBox "VB Script Engine 5.0 or higher required"
       ExecuteGlobal GetTextFile(VBSfile)
       If Err Then MsgBox "Unable to open " & VBSfile & ". Ensure that it is in the same folder as this table. " & vbNewLine & Err.Description
       If Table.ShowDT = true Then
       Set Controller = CreateObject("VPinMAME.Controller")
       B2SOn = 0
       else
       Set Controller = CreateObject("B2S.server")
       'Set Controller = CreateObject("VPinMAME.Controller")
       B2SOn = 1
       End If
       If Err Then MsgBox "Can't Load VPinMAME." & vbNewLine & Err.Description
       If VPMver > "" Then If Controller.Version < VPMver Or Err Then MsgBox "VPinMAME ver " & VPMver & " required."
       If VPMver > "" Then If Controller.Version < VPMver Or Err Then MsgBox "VPinMAME ver " & VPMver & " required."
       If VPinMAMEDriverVer < VBSver Or Err Then MsgBox VBSFile & " ver " & VBSver & " or higher required."
       On Error Goto 0
     End Sub

'********************
'Standard definitions
'********************
     'Const B2SOn = 1

	 Const cGameName = "mt_145hc" 'change the romname here

     Const UseSolenoids = 1
     Const UseLamps = 0
     Const UseSync = 0
     Const HandleMech = 0

     'Standard Sounds
     Const SSolenoidOn = "Solenoid"
     Const SSolenoidOff = ""
     Const SCoin = "CoinIn"

'************
' Table init.
'************
   'Variables
    Dim xx, x
    Dim Bump1, Bump2, Bump3, Bump4, Mech3bank,bsTrough,bsRHole,DTBank5,turntable, cbRight, XTurn, mspinmagnet
	Dim PlungerIM, B2SOn

  Sub Table_Init
	With Controller
        vpmInit Me
		.GameName = cGameName
		If Err Then MsgBox "Can't start Game " & cGameName & vbNewLine & Err.Description:Exit Sub
		.SplashInfoLine = "Mustang (Stern 2014)"
		.HandleKeyboard = 0
		.ShowTitle = 0
		.ShowDMDOnly = 1
		.ShowFrame = 0
		.HandleMechanics = 1
		.Hidden = VarHidden
		On Error Resume Next
		.Run GetPlayerHWnd
		If Err Then MsgBox Err.Description
	End With
  InitVpmFFlipsSAM
    On Error Goto 0


'**Trough
    Set bsTrough = New cvpmBallStack
    bsTrough.InitSw 0, 23, 22, 21, 20, 19, 18, 0
    bsTrough.InitKick BallRelease, 90, 8
    bsTrough.InitExitSnd "ballrelease", "Solenoid"
    bsTrough.Balls = 6

	'***Right Hole bsRHole
     'Set bsRHole = New cvpmBallStack
     'With bsRHole
      '   .InitSw 0, 43, 0, 0, 0, 0, 0, 0
      '   .InitKick sw43, 185, 20
      '   .KickZ = 0.4
      '   .InitExitSnd "scoopexit", "Solenoid"
      '   .KickForceVar = 2
     'End With

   Set bsRHole = New cvpmBallStack
	With bsRHole
		.InitSaucer sw43, 43, 185, 20
		.KickZ = 0.4
		.InitExitSnd "fx_solenoid", "Solenoid"
		.KickForceVar = 2
	End With

      Set cbRight = New cvpmCaptiveBall
     With cbRight
         .InitCaptive RCaptTrigger, RCaptWall, Array(RCaptKicker1, RCaptKicker1a), 10
		 .RestSwitch = 9
         .NailedBalls = 1
         .ForceTrans = 1
         .MinForce = 7
         .Start
     End With
     RCaptKicker1.CreateBall

   Set mSpinMagnet = New cvpmMagnet
    With mSpinMagnet
        .InitMagnet SpinMagnet, 20
        '.Solenoid = 35 'own solenoid sub
        .GrabCenter = 0
        .Size = 100
        .CreateEvents "mSpinMagnet"
    End With


'**Nudging
    	vpmNudge.TiltSwitch=-7
    	vpmNudge.Sensitivity=1
    	vpmNudge.TiltObj=Array(Bumper1,Bumper2,Bumper3,Bumper4,LeftSlingshot,RightSlingshot)

'DropTargets

      '**Main Timer init
	PinMAMETimer.Interval = PinMAMEInterval
	PinMAMETimer.Enabled = 1

  'Slings
    For each xx in RhammerA:xx.IsDropped=1:Next
    For each xx in RhammerB:xx.IsDropped=1:Next
    For each xx in RhammerC:xx.IsDropped=1:Next
    For each xx in LhammerA:xx.IsDropped=1:Next
    For each xx in LhammerB:xx.IsDropped=1:Next
    For each xx in LhammerC:xx.IsDropped=1:Next

'DropTargets
	Set DTBank5 = New cvpmDropTarget
	With DTBank5
	   .InitDrop Array(sw34,sw35,sw36,sw37,sw38), Array(34,35,36,37,38)
	   .InitSnd"DTL", "DTResetL"
	End With

   'Rollovers

  'StandUp Init
	sw1a.IsDropped=1
	sw2a.IsDropped=1
	sw3a.IsDropped=1
	sw4a.IsDropped=1
	sw5a.IsDropped=1
	sw41a.IsDropped=1
	sw42a.IsDropped=1
	sw54a.IsDropped=1
	sw55a.IsDropped=1

    'Gi_On
	Plunger1.Pullback
 'mSpinMagnet.MagnetOn = True

If erratic_scoop = 1 Then
sw43.enabled = 0
Else
sw43.enabled = 1
End If

 End Sub

 Sub Table_Paused:Controller.Pause = 1:End Sub
 Sub Table_unPaused:Controller.Pause = 0:End Sub


'*****Keys
 Sub Table_KeyDown(ByVal keycode)
If KeyCode = MechanicalTilt Then
        vpmTimer.PulseSw vpmNudge.TiltSwitch
        Exit Sub
    End if
If Keycode = StartGameKey Then Controller.Switch(16) = 1
If Keycode = 3 Then Controller.Switch(15) = 1
If Keycode = RightMagnasave Then Controller.Switch(71) = 1
' 	If Keycode = LeftFlipperKey then
'	End If
' 	If Keycode = RightFlipperKey then
'	End If
    If keycode = PlungerKey Then Plunger.Pullback
  	'If keycode = LeftTiltKey Then LeftNudge 80, 1, 20
    'If keycode = RightTiltKey Then RightNudge 280, 1, 20
    'If keycode = CenterTiltKey Then CenterNudge 0, 1, 25
    'If keycode = 3 Then msgbox Disc2.ObjRotZ
    If vpmKeyDown(keycode) Then Exit Sub

End Sub

Sub Table_KeyUp(ByVal keycode)
	If Keycode = RightMagnasave Then Controller.Switch(71) = 0
	If Keycode = 3 Then Controller.Switch(15) = 0
	If Keycode = StartGameKey Then Controller.Switch(16) = 0':GridTargetReset 'Moved before vpmkey
	If vpmKeyUp(keycode) Then Exit Sub
 	If Keycode = LeftFlipperKey then
		SolLFlipper false
	End If
 	If Keycode = RightFlipperKey then
		SolRFlipper False
	End If
    If keycode = PlungerKey Then
        Plunger.Fire
            PlaySound "Plunger2"
End If
End Sub


   'Solenoids
SolCallback(1) = "solTrough"
SolCallback(2) = "solAutofire"
SolCallback(3) = "lower_ramp"
SolCallback(4) = "lower_hold"
SolCallback(5) = "upper_ramp"
SolCallback(6) = "upper_hold"
 SolCallback(7) = "Reset_Drops"
'SolCallback(8) = shaker
'SolCallback(9) = l pop
'SolCallback(10) = r pop
'SolCallback(11) = B pop
'SolCallback(12) = T pop
'SolCallback(13) = l sling
'SolCallback(14) = r sling
SolCallback(15) = "SolLFlipper"
SolCallback(16) = "SolRFlipper"
SolCallback(22) = "Ttable"
SolModCallback(17) = "Flash17"
SolModCallback(18) = "Flash18"
SolModCallback(19) = "Flash19"
SolModCallback(20) = "Flash20"
SolModCallback(21) = "Flash21"
SolModCallback(23) = "Flash23"
SolModCallback(25) = "Flash25"
SolModCallback(26) = "Flash26"
SolModCallback(27) = "Flash27"
SolCallback(28) = "Flash28"
SolModCallback(29) = "Flash29"
SolCallback(30) = "Flash30"
SolCallback(51) = "Left_Orbit_Gate"
SolCallback(52) = "Right_Orbit_Gate"
SolCallback(56) = "LHDTD"
SolCallback(55) = "LHDTU"
SolCallback(54) = "RHDTD"
SolCallback(53) = "RHDTU"
SolCallback(59) = "Kicker_Out"
SolCallback(60) = "Ramp_Div"

If Inteceptor = 1 Then
Primitive3.visible = 0
Primitive53.visible = 1
else
Primitive3.visible = 1
Primitive53.visible = 0
End If

Dim dtff

If Flasher_Halos = 1 Then
For each dtff in Flashers:dtff.visible = 1:next
else
For each dtff in Flashers:dtff.visible = 0:next
End If

If Tacho_Mod = 1 Then
Disc4.visible = 1
Disc2.visible = 1
Disc3.visible = 1
Tacho_Bod.visible = 1
else
Disc4.visible = 0
Disc2.visible = 0
Disc3.visible = 0
Tacho_Bod.visible = 0
End If

Sub Wall261_Hit()
If erratic_scoop = 1 Then
mSpinMagnet.MagnetOn = True
shopstep = 1 : shoph.enabled = 1
End If
End Sub

'Sub Wall19_Hit()
'mSpinMagnet.MagnetOn = True
'shopstep = 1 : shoph.enabled = 1
'End Sub

'Sub Wall18_Hit()
'mSpinMagnet.MagnetOn = True
'shopstep = 1 : shoph.enabled = 1
'End Sub

Dim shopstep

Sub shoph_Timer()
Select Case shopstep
Case 1:
mSpinMagnet.MagnetOn = False
sw43.enabled = 1
shopstep = 2
Case 2:
me.enabled = 0
End Select
End Sub

Sub Reset_Drops(Enabled)
If enabled Then
 DTBank5.DropSol_On
End If
End Sub

Sub LHDTD(Enabled)
If enabled Then
LHDT.isdropped = 1
LHDTW.isdropped = 1
LHDTW1.isdropped = 1
PlaySound "dtresetl"
LHDBU.enabled = 1
Controller.Switch(57) = 1
End If
End Sub

Sub LHDBU_Timer()
If LHDT.isdropped = 0 Then
LHDT.isdropped = 1
End If
me.enabled = 0
End Sub

Sub LHDTU(Enabled)
If enabled Then
LHDT.isdropped = 0
LHDTW.isdropped = 0
LHDTW1.isdropped = 0
PlaySound "dtresetl"
Controller.Switch(57) = 0
End If
End Sub

Sub RHDTD(Enabled)
If enabled Then
XRHDT.isdropped = 1
XRHDTW.isdropped = 1
XRHDTW1.isdropped = 1
PlaySound "dtresetr"
RHDBU.enabled = 1
Controller.Switch(56) = 1
End If
End Sub

Sub RHDBU_Timer()
If XRHDT.isdropped = 0 Then
XRHDT.isdropped = 1
End If
me.enabled = 0
End Sub

Sub RHDTU(Enabled)
If enabled Then
XRHDT.isdropped = 0
XRHDTW.isdropped = 0
XRHDTW1.isdropped = 0
PlaySound "dtresetr"
Controller.Switch(56) = 0
End If
End Sub


Sub Ramp_Div(Enabled)
If enabled Then
DivKick1.enabled = 1
PlaySound "diverter"
else
DivKick1.enabled = 0
PlaySound "diverter"
End If
End Sub

Sub Flash29(Enabled)
F29.Intensity = Enabled / 20
F29a.Intensity = Enabled / 20
F29b.Intensity = Enabled / 20
End Sub

Sub Flash19(Enabled)
F19.Intensity = Enabled / 5
F19a.Intensity = Enabled / 5
F19b.Intensity = Enabled / 60
F19c.Intensity = Enabled / 60
F19F.Opacity = Enabled * 1.5
F19F1.Opacity = Enabled * 1.5
End Sub

Sub Flash20(Enabled)
F20.Intensity = Enabled / 5
F20a.Intensity = Enabled / 5
F20b.Intensity = Enabled / 60
F20c.Intensity = Enabled / 60
F20F.Opacity = Enabled * 1.5
F20F1.Opacity = Enabled * 1.5
End Sub

Sub Flash28(Enabled)
If enabled Then
F28.State = 1
F28a.State = 1
F28b.State = 1
F28c.State = 1
else
F28.State = 0
F28a.State = 0
F28b.State = 0
F28c.State = 0
End If
End Sub

Sub Flash25(Enabled)
F25.Intensity = Enabled / 10
F25a.Intensity = Enabled / 10
End Sub

Sub Flash26(Enabled)
F26.Intensity = Enabled / 10
F26a.Intensity = Enabled / 10
End Sub

Sub Flash23(Enabled)
F23Flash1.Opacity = Enabled * 1.5
F23Flash2.Opacity = Enabled * 1.5
F23.Intensity = Enabled / 30
F23a.Intensity = Enabled / 10
F23b.Intensity = Enabled / 10
F23c.Intensity = Enabled / 50
F23d.Intensity = Enabled / 10
End Sub

Sub Flash21(Enabled)
F21Flash2.Opacity = Enabled * 1.5
F21F2.Opacity = Enabled * 1.5
Fl21.Opacity = Enabled * 1.5
F21.Intensity = Enabled / 30
F21a.Intensity = Enabled / 10
F21b.Intensity = Enabled / 10
F21c.Intensity = Enabled / 50
F21d.Intensity = Enabled / 10
End Sub

Sub Flash17(Enabled)
F17.Intensity = Enabled / 10
F17a.Intensity = Enabled / 10
End Sub

Sub Flash18(Enabled)
F18.Intensity = Enabled / 10
F18a.Intensity = Enabled / 10
End Sub

Sub Flash27(Enabled)
F27.Intensity = Enabled / 10
F27a.Intensity = Enabled / 10
End Sub

Sub Flash30(Enabled)
If Enabled Then
F30.State = 1
F30a.State = 1
else
F30.State = 0
F30a.State = 0
End If
End Sub

Sub Left_Orbit_Gate(Enabled)
If enabled Then
Lorbit.open = 1
PlaySound "dtl"
else
Lorbit.open = 0
End If
End Sub

Sub Right_Orbit_Gate(Enabled)
If enabled Then
Rorbit.open = 1
PlaySound "dtr"
else
Rorbit.open = 0
End If
End Sub

Sub Kicker_Out(Enabled)
If enabled Then
BsRHole.ExitSol_On
If Erratic_Scoop = 1 Then
sw43.enabled = 0
End If
End If
End Sub

Sub upper_ramp(Enabled)
	If Enabled Then
PlaySound "dtl"
	End If
 End Sub

Sub upper_hold(Enabled)
    If enabled Then

		trpos = 2
        top_ramp.enabled = 1
        topramp5.collidable = 1


  else
		trpos = 1
		top_ramp.enabled = 1
        topramp5.collidable = 0
    PlaySound "dtl"
	End If
End Sub

Sub lower_ramp(Enabled)
	If Enabled Then
PlaySound "dtl"
	End If
End Sub

Sub lower_hold(Enabled)
    If Enabled Then

        brpos = 2
        bottom_ramp.enabled = 1
        BotRamp2.collidable = 1
else
	brpos = 1
	bottom_ramp.enabled = 1
    BotRamp2.collidable = 0
    PlaySound "dtl"
	End If
End Sub

Dim brpos
brpos = 0

Sub Bottom_Ramp_Timer()
Select Case brpos
Case 1:
If BotRamp.heightbottom => 60 Then
BotRamp.heightbottom = 60
me.enabled = 0
End If
BotRamp.heightbottom = BotRamp.heightbottom + 1
BotRamp_Cover.heightbottom = BotRamp_Cover.heightbottom  + 1.5
BotRamp_Cover.heighttop = BotRamp_Cover.heighttop  + 1
Case 2:
If BotRamp.heightbottom <= 0 Then
BotRamp.heightbottom = 0
me.enabled = 0
End If
BotRamp.heightbottom = BotRamp.heightbottom - 1
BotRamp_Cover.heightbottom = BotRamp_Cover.heightbottom  - 1.5
BotRamp_Cover.heighttop = BotRamp_Cover.heighttop  - 1
Case 3:
End Select
End Sub

Dim trpos
trpos = 0

Sub Top_Ramp_Timer()
Select Case trpos
Case 1:
If topRamp.heightbottom => 120 Then
topRamp.heightbottom = 120
topRamp2.collidable = 1
Controller.Switch(50) = 1
me.enabled = 0
End If
topRamp.heightbottom = topRamp.heightbottom + 1
topRamp_Cover.heightbottom = topRamp_Cover.heightbottom  + 1
topRamp_Cover.heighttop = topRamp_Cover.heighttop  + 0.5
Case 2:
If topRamp.heightbottom <= 60 Then
topRamp.heightbottom = 60
topRamp2.collidable = 0
Controller.Switch(50) = 0
me.enabled = 0
End If
topRamp.heightbottom = topRamp.heightbottom - 1
topRamp_Cover.heightbottom = topRamp_Cover.heightbottom  - 1
topRamp_Cover.heighttop = topRamp_Cover.heighttop  - 0.5
End Select
End Sub

'Flashers


Sub solTrough(Enabled)
	If Enabled Then
		bsTrough.ExitSol_On
		vpmTimer.PulseSw 22
	End If
 End Sub

Sub solAutofire(Enabled)
	If Enabled Then
		'PlungerIM.AutoFire
         Plunger1.Fire
Else
		 Plunger1.Pullback
	End If
 End Sub


 ' Captive Ball Right
Sub RCaptTrigger_Hit:cbRight.TrigHit ActiveBall:End Sub
Sub RCaptTrigger_UnHit:cbRight.TrigHit 0:End Sub
Sub RCaptWall_Hit:PlaySound "collide0":cbRight.BallHit ActiveBall:End Sub
Sub RCaptKicker1a_Hit:cbRight.BallReturn Me:End Sub

'***********************************************

'******************************************
' Use FlipperTimers to call div subs
'******************************************

'Dim LFTCount:LFTCount=1
'
'Sub LeftFlipperTimer_Timer()
'	If LFTCount < 6 Then
'		LFTCount = LFTCount + 1
'		LeftFlipper.Strength = StartLeftFlipperStrength*(LFTCount/6)
'	Else
'		Me.Enabled=0
'	End If
'End Sub
'
Dim RFTCount:RFTCount=1
'
'Sub RightFlipperTimer_Timer()
'	If RFTCount < 6 Then
'		RFTCount = RFTCount + 1
'		RightFlipper.Strength = StartRightFlipperStrength*(RFTCount/6)
'	Else
'		Me.Enabled=0
'	End If
'End Sub


Sub SolLFlipper(Enabled)
     If Enabled Then
		 PlaySound "flipperup"
		 LeftFlipper.RotateToEnd:
     Else
'		 LFTCount=1
		 PlaySound "flipperdown"
		 LeftFlipper.RotateToStart
     End If
 End Sub

Sub SolRFlipper(Enabled)
     If Enabled Then
		 PlaySound "flipperup"
		 RightFlipper.RotateToEnd
     Else
'		 RFTCount=1
		 PlaySound "flipperdown"
		 RightFlipper.RotateToStart
     End If
 End Sub

 'Drains and Kickers
   Sub Drain_Hit():PlaySound "Drain":bsTrough.AddBall Me
   End Sub

Sub orbitpost(Enabled)
	If Enabled Then
		UpPost.Isdropped=false
		UpPost2.Isdropped=false
	Else
		UpPost.Isdropped=true
		UpPost2.Isdropped=true
	End If
 End Sub


'Switches

Sub sw1_Hit:sw1.IsDropped = 1:sw1a.IsDropped = 0:Me.TimerEnabled = 1:vpmTimer.PulseSw 1:PlaySound "target":End Sub
Sub sw1_Timer:sw1.IsDropped = 0:sw1a.IsDropped = 1:Me.TimerEnabled = 0:End Sub
Sub sw2_Hit:sw2.IsDropped = 1:sw2a.IsDropped = 0:Me.TimerEnabled = 1:vpmTimer.PulseSw 2:PlaySound "target":End Sub
Sub sw2_Timer:sw2.IsDropped = 0:sw2a.IsDropped = 1:Me.TimerEnabled = 0:End Sub
Sub sw3_Hit:sw3.IsDropped = 1:sw3a.IsDropped = 0:Me.TimerEnabled = 1:vpmTimer.PulseSw 3:PlaySound "target":End Sub
Sub sw3_Timer:sw3.IsDropped = 0:sw3a.IsDropped = 1:Me.TimerEnabled = 0:End Sub
Sub sw4_Hit:sw4.IsDropped = 1:sw4a.IsDropped = 0:Me.TimerEnabled = 1:vpmTimer.PulseSw 4:PlaySound "target":End Sub
Sub sw4_Timer:sw4.IsDropped = 0:sw4a.IsDropped = 1:Me.TimerEnabled = 0:End Sub
Sub sw5_Hit:sw5.IsDropped = 1:sw5a.IsDropped = 0:Me.TimerEnabled = 1:vpmTimer.PulseSw 5:PlaySound "target":End Sub
Sub sw5_Timer:sw5.IsDropped = 0:sw5a.IsDropped = 1:Me.TimerEnabled = 0:End Sub
Sub sw8_Hit:Controller.Switch(8) = 1:PlaySound "rollover":End Sub
Sub sw8_UnHit:Controller.Switch(8) = 0:End Sub
'sw9 captive ball front
Sub sw10_Hit:Controller.Switch(10)=1:End Sub
Sub sw10_unHit:Controller.Switch(10)=0:End Sub
Sub sw11_Hit:la11.IsDropped = 1:Controller.Switch(11) = 1:PlaySound "rollover":End Sub
Sub sw11_UnHit:la11.IsDropped = 0:Controller.Switch(11) = 0:End Sub
Sub sw12_Hit:la12.IsDropped = 1:Controller.Switch(12) = 1:PlaySound "rollover":End Sub
Sub sw12_UnHit:la12.IsDropped = 0:Controller.Switch(12) = 0:End Sub
Sub sw13_Hit:la13.IsDropped = 1:Controller.Switch(13) = 1:PlaySound "rollover":End Sub
Sub sw13_UnHit:la13.IsDropped = 0:Controller.Switch(13) = 0:End Sub
Sub sw14_Hit:la14.IsDropped = 1:Controller.Switch(14) = 1:PlaySound "rollover":End Sub
Sub sw14_UnHit:la14.IsDropped = 0:Controller.Switch(14) = 0:End Sub
Sub sw24_Hit:la24.IsDropped = 1:Controller.Switch(24) = 1:PlaySound "rollover":End Sub
Sub sw24_UnHit:la24.IsDropped = 0:Controller.Switch(24) = 0:End Sub
Sub sw25_Hit:la25.IsDropped = 1:Controller.Switch(25) = 1:PlaySound "rollover":End Sub
Sub sw25_UnHit:la25.IsDropped = 0:Controller.Switch(25) = 0:End Sub
Sub sw28_Hit:la28.IsDropped = 1:Controller.Switch(28) = 1:PlaySound "rollover":End Sub
Sub sw28_UnHit:la28.IsDropped = 0:Controller.Switch(28) = 0:End Sub
Sub sw29_Hit:la29.IsDropped = 1:Controller.Switch(29) = 1:PlaySound "rollover":End Sub
Sub sw29_UnHit:la29.IsDropped = 0:Controller.Switch(29) = 0:End Sub
Sub sw34_Hit:DTBank5.Hit 1:End Sub'Light54.State = 1:End Sub
Sub sw35_Hit:DTBank5.Hit 2:End Sub'Light53.State = 1:End Sub
Sub sw36_Hit:DTBank5.Hit 3:End Sub'Light51.State = 1:End Sub
Sub sw37_Hit:DTBank5.Hit 4:End Sub'Light50.State = 1:End Sub
Sub sw38_Hit:DTBank5.Hit 5:End Sub'Light52.State = 1:End Sub
Sub sw39_Hit:Controller.Switch(39) = 1:PlaySound "subway2":End Sub'Gi_Off:Gi_BO.enabled = 1:End Sub
Sub sw39_UnHit:Controller.Switch(39) = 0:End Sub
Sub sw40_Hit:Controller.Switch(40) = 1:PlaySound "subway2":End Sub'Gi_Off:Gi_BO.enabled = 1:End Sub
Sub sw40_UnHit:Controller.Switch(40) = 0:End Sub
Sub sw41_Hit:sw41.IsDropped = 1:sw41a.IsDropped = 0:Me.TimerEnabled = 1:vpmTimer.PulseSw 41:PlaySound "target":End Sub
Sub sw41_Timer:sw41.IsDropped = 0:sw41a.IsDropped = 1:Me.TimerEnabled = 0:End Sub
Sub sw42_Hit:sw42.IsDropped = 1:sw42a.IsDropped = 0:Me.TimerEnabled = 1:vpmTimer.PulseSw 42:PlaySound "target":End Sub
Sub sw42_Timer:sw42.IsDropped = 0:sw42a.IsDropped = 1:Me.TimerEnabled = 0:End Sub
Sub sw44_Hit:Controller.Switch(44)=1:PlaySound "rollover":End Sub
Sub sw44_unHit:Controller.Switch(44)=0:End Sub
Sub sw45_Hit:Controller.Switch(45)=1:PlaySound "rollover":End Sub
Sub sw45_unHit:Controller.Switch(45)=0:End Sub
Sub sw46_Hit:Controller.Switch(46) = 1:PlaySound "bowl":bowl_sw.rotatetoend:End Sub
Sub sw46_UnHit:Controller.Switch(46) = 0:bowl_sw.rotatetostart:End Sub
Sub sw47_Hit:Controller.Switch(47) = 1:PlaySound "rollover":End Sub
Sub sw47_UnHit:Controller.Switch(47) = 0:End Sub'GI_On:End Sub
Sub sw48_Spin:vpmTimer.PulseSw 48:playsound"spinner": tach.enabled = 1:End Sub
Sub sw54_Hit:sw54.IsDropped = 1:sw54a.IsDropped = 0:Me.TimerEnabled = 1:vpmTimer.PulseSw 54:PlaySound "target":End Sub
Sub sw54_Timer:sw54.IsDropped = 0:sw54a.IsDropped = 1:Me.TimerEnabled = 0:End Sub
Sub sw55_Hit:sw55.IsDropped = 1:sw55a.IsDropped = 0:Me.TimerEnabled = 1:vpmTimer.PulseSw 55:PlaySound "target":End Sub
Sub sw55_Timer:sw55.IsDropped = 0:sw55a.IsDropped = 1:Me.TimerEnabled = 0:End Sub
Sub lhdtw_Hit:vpmTimer.PulseSw 57:Playsound "target":End Sub
Sub xrhdtw_Hit:vpmTimer.PulseSw 56:Playsound "target":End Sub

Sub rub1_Hit():playsound"rubber":End Sub
Sub rub2_Hit():playsound"rubber":End Sub
Sub rub3_Hit():playsound"rubber":End Sub
Sub rub4_Hit():playsound"rubber":End Sub
Sub rub5_Hit():playsound"rubber":End Sub
Sub rub6_Hit():playsound"rubber":End Sub
Sub rub7_Hit():playsound"rubber":End Sub
Sub rub8_Hit():playsound"rubber":End Sub
Sub rub9_Hit():playsound"rubber":End Sub
Sub rub10_Hit():playsound"rubber":End Sub
Sub rub11_Hit():playsound"rubber":End Sub
Sub rub12_Hit():playsound"rubber":End Sub
Sub rub13_Hit():playsound"rubber":End Sub
Sub rub14_Hit():playsound"rubber":End Sub
Sub rub15_Hit():playsound"rubber":End Sub
Sub rub16_Hit():playsound"rubber":End Sub
Sub UpPost2_Hit():playsound"rubber":End Sub


'Scoop
 Dim aBall, aZpos
 Dim bBall, bZpos

Sub Wall17_Hit
PlaySound "fx_collide"
End Sub

Sub Wall16_Hit
PlaySound "fx_collide"
End Sub

Sub sw43_Hit
     PlaySound "ballhit"
     Me.TimerInterval = 2
     Me.TimerEnabled = 1
 End Sub

 Sub sw43_Timer
          Me.TimerEnabled = 0
          bsRHole.AddBall Me
 End Sub

'***Slings and rubbers
  ' Slings
 Dim LStep, RStep

 Sub LeftSlingShot_Slingshot
 	For each xx in LHammerA:xx.IsDropped = 0:Next
 	PlaySound "Slingshot":vpmTimer.PulseSw 26:LStep = 0:Me.TimerEnabled = 1
  End Sub

 Sub LeftSlingShot_Timer
     Select Case LStep
         Case 0: 'pause
         Case 1: 'pause
         Case 2:For each xx in LHammerA:xx.IsDropped = 1:Next
 				For each xx in LHammerB:xx.IsDropped = 0:Next
         Case 3:For each xx in LHammerB:xx.IsDropped = 1:Next
 				For each xx in LHammerC:xx.IsDropped = 0:Next
         Case 4:For each xx in LHammerC:xx.IsDropped = 1:Next
 				Me.TimerEnabled = 0
     End Select
     LStep = LStep + 1
 End Sub

 Sub RightSlingShot_Slingshot
 	For each xx in RHammerA:xx.IsDropped = 0:Next
 	PlaySound "Slingshot":vpmTimer.PulseSw 27:RStep = 0:Me.TimerEnabled = 1
  End Sub

 Sub RightSlingShot_Timer
     Select Case RStep
         Case 0: 'pause
         Case 1: 'pause
         Case 2:For each xx in RHammerA:xx.IsDropped = 1:Next
 				For each xx in RHammerB:xx.IsDropped = 0:Next
         Case 3:For each xx in RHammerB:xx.IsDropped = 1:Next
 				For each xx in RHammerC:xx.IsDropped = 0:Next
         Case 4:For each xx in RHammerC:xx.IsDropped = 1:Next
 				Me.TimerEnabled = 0
     End Select
     RStep = RStep + 1
 End Sub

   'Bumpers
      Sub Bumper1_Hit:vpmTimer.PulseSw 30:PlaySound "bumper":End Sub


      Sub Bumper2_Hit:vpmTimer.PulseSw 32:PlaySound "bumper":End Sub


      Sub Bumper3_Hit:vpmTimer.PulseSw 33:PlaySound "bumper":End Sub


      Sub Bumper4_Hit:vpmTimer.PulseSw 31:PlaySound "bumper":End Sub


 'Sounds
 dim speedx
 dim speedy
 dim finalspeed
  Sub Rubbers_Hit(IDX)
 	finalspeed=SQR(ActiveBall.velx * ActiveBall.velx + ActiveBall.vely * ActiveBall.vely)
 	if finalspeed > 11 then PlaySound "rubber" else PlaySound "rubberFlipper":end if
   End Sub
  Sub Gates_Hit(IDX):PlaySound"Gate":End Sub
  Sub LeftFlipper_Collide(parm)
      PlaySound "rubberFlipper"
      'GridCancelGameOverTimer
   End Sub
   Sub RightFlipper_Collide(parm)
      PlaySound "rubberFlipper"
      'GridCancelGameOverTimer
   End Sub

Dim LampState(400)

AllLampsOff()
LampTimer.Interval = 35
LampTimer.Enabled = 1

Sub LampTimer_Timer()
    Dim chgLamp, ii
    chgLamp = Controller.ChangedLamps
    If Not IsEmpty(chgLamp) Then
        For ii = 0 To UBound(chgLamp)

			LampState(chgLamp(ii, 0) ) = chgLamp(ii, 1)

			'Print Light state changes
			debug.print "Lamp " & chgLamp(ii, 0) & ": " & chgLamp(ii, 1)


        Next
    End If
    UpdateLamps
End Sub

Sub AllLampsOff()
    On Error Resume Next

	Dim x
	For x = 0 to 360
		LampState(x) = 0
	Next

	UpdateLamps:UpdateLamps:Updatelamps
End Sub

Sub SetLamp(nr, value)
	If value <> LampState(nr) Then
		LampState(nr) = abs(value)
	End If
End Sub


'Dim gon
Dim xxtt

Sub UpdateLamps
       On Error Resume Next

For each xxtt in GI_ALL:xxtt.Intensity = LampState(98) / 20:next

If LampState(98) < 10 Then
Table.ColorGradeImage = "ColorGrade_off"
else
Table.ColorGradeImage = "ColorGrade_on"
End If

if sw34.isdropped AND LampState(98) > 10 Then
Light54.State = 1
else
Light54.State = 0
End If

if sw35.isdropped AND LampState(98) > 10 Then
Light53.State = 1
else
Light53.State = 0
End If

if sw36.isdropped AND LampState(98) > 10 Then
Light51.State = 1
else
Light51.State = 0
End If

if sw37.isdropped AND LampState(98) > 10 Then
Light50.State = 1
else
Light50.State = 0
End If

if sw38.isdropped AND LampState(98) > 10 Then
Light52.State = 1
else
Light52.State = 0
End If

l3.State = Lampstate(3)
l4.State = Lampstate(4)
l5.State = Lampstate(5)
l6.State = Lampstate(6)
l7.State = Lampstate(7)
l8.State = Lampstate(8)

xg.State = Lampstate(9)
xn.State = Lampstate(10)
xa.State = Lampstate(11)
xt.State = Lampstate(12)
xs.State = Lampstate(13)
xu.State = Lampstate(14)
xm.State = Lampstate(15)

L33.State = Lampstate(33)
L34.State = Lampstate(34)
L35.State = Lampstate(35)

L16.State = Lampstate(16)
L17.State = Lampstate(17)
L18.State = Lampstate(18)
L19.State = Lampstate(19)

L20.State = Lampstate(20)
L21.State = Lampstate(21)
L22.State = Lampstate(22)
L23.State = Lampstate(23)

L20a.State = Lampstate(20)
L21a.State = Lampstate(21)
L22a.State = Lampstate(22)
L23a.State = Lampstate(23)

L24.State = Lampstate(24)
L26.State = Lampstate(26)
L27.State = Lampstate(27)
L28.State = Lampstate(28)
L29.State = Lampstate(29)
L30.State = Lampstate(30)
L31.State = Lampstate(31)
L32.State = Lampstate(32)
L25.State = Lampstate(25)

L36.State = Lampstate(36)
L37.State = Lampstate(37)
L38.State = Lampstate(38)

L39.State = Lampstate(39)
L40.State = Lampstate(40)
L41.State = Lampstate(41)

L44.State = Lampstate(44)
L45.State = Lampstate(45)
L46.State = Lampstate(46)

L47.State = Lampstate(47)
L48.State = Lampstate(48)
L42.State = Lampstate(42)
L43.State = Lampstate(43)
L49.State = Lampstate(49)
L50.State = Lampstate(50)
L51.State = Lampstate(51)

L52.State = Lampstate(52)
L53.State = Lampstate(53)
L54.State = Lampstate(54)

L55.State = Lampstate(55)
L56.State = Lampstate(56)
L57.State = Lampstate(57)

L58.State = Lampstate(58)
L59.State = Lampstate(59)
L60.State = Lampstate(60)
L61.State = Lampstate(61)
L62.State = Lampstate(62)
L63.State = Lampstate(63)
L64.State = Lampstate(64)

L65.State = Lampstate(65)
L66.State = Lampstate(66)
L67.State = Lampstate(67)
L68.State = Lampstate(68)
L69.State = Lampstate(69)
L70.State = Lampstate(70)

L73.State = Lampstate(73)

L77.State = Lampstate(77)
L78.State = Lampstate(78)
L79.State = Lampstate(79)
L79a.State = Lampstate(79)
L80.State = Lampstate(80)

L108.State = Lampstate(112)

'L39.State = Lampstate(39)

L81.State = Lampstate(81)
L82.State = Lampstate(82)
L83.State = Lampstate(83)
L84.State = Lampstate(84)
L85.State = Lampstate(85)

L86.State = Lampstate(86)
L87.State = Lampstate(87)
L88.State = Lampstate(88)
L89.State = Lampstate(89)
L90.State = Lampstate(90)

L91.State = Lampstate(91)
L92.State = Lampstate(92)
L93.State = Lampstate(93)
L94.State = Lampstate(94)
L95.State = Lampstate(95)

L96.State = Lampstate(96)
L97.State = Lampstate(97)

L99.State = Lampstate(102)
L99a.State = Lampstate(102)
L99b.State = Lampstate(102)

L98.State = Lampstate(103)
L98a.State = Lampstate(103)
L98b.State = Lampstate(103)

L100.State = Lampstate(104)
L100a.State = Lampstate(104)
L100b.State = Lampstate(104)

L103.State = Lampstate(107)
L103a.State = Lampstate(107)
L103b.State = Lampstate(107)
L104.State = Lampstate(108)
L104a.State = Lampstate(108)
L104b.State = Lampstate(108)

L101.State = Lampstate(105)
L101a.State = Lampstate(105)
L101b.State = Lampstate(105)
L102.State = Lampstate(106)
L102a.State = Lampstate(106)
L102b.State = Lampstate(106)
L105.State = Lampstate(109)
L105a.State = Lampstate(109)
L105b.State = Lampstate(109)

L106.State = Lampstate(110)
L106a.State = Lampstate(110)
L106b.State = Lampstate(110)

L107.State = Lampstate(111)
L107a.State = Lampstate(111)
L107b.State = Lampstate(111)

rgb1.Color = RGB(0,0,0)
rgb1.ColorFull = RGB(Lampstate(125),Lampstate(123),Lampstate(124))

rgb2.Color = RGB(0,0,0)
rgb2.ColorFull = RGB(Lampstate(128),Lampstate(126),Lampstate(127))

rgb3.Color = RGB(0,0,0)
rgb3.ColorFull = RGB(Lampstate(122),Lampstate(120),Lampstate(121))

rgb8.Color = RGB(0,0,0)
rgb8.ColorFull = RGB(Lampstate(119),Lampstate(117),Lampstate(118))

rgb4.Color = RGB(0,0,0)
rgb4.ColorFull = RGB(Lampstate(141),Lampstate(139),Lampstate(140))

rgb5.Color = RGB(0,0,0)
rgb5.ColorFull = RGB(Lampstate(135),Lampstate(133),Lampstate(134))

rgb6.Color = RGB(0,0,0)
rgb6.ColorFull = RGB(Lampstate(132),Lampstate(130),Lampstate(131))

rgb7.Color = RGB(0,0,0)
rgb7.ColorFull = RGB(Lampstate(138),Lampstate(136),Lampstate(137))


N8.State = Lampstate(98)
N9.State = Lampstate(99)
OH.State = Lampstate(100)
OH1.State = Lampstate(101)

End Sub

 '*************************
 ' Plunger kicker animation
 '*************************


Sub Trigger1_hit
	PlaySound "DROP_LEFT"
 End Sub

 Sub Trigger2_hit
	PlaySound "DROP_RIGHT"
 End Sub

 Sub RHD_hit
	PlaySound "DROP_RIGHT"
 End Sub

Sub Table_exit()
	Controller.Pause = False
	Controller.Stop
End Sub


Sub Ttable(Enabled)
If enabled Then
car_wheel.enabled = 1
else
car_wheel.enabled = 0
End If
End Sub

Sub car_wheel_Timer()

If Disc2.ObjRotz => 360 Then
Disc2.ObjRotz = 0
End If

If Disc2.ObjRotz > 3 AND Disc2.ObjRotz < 53 AND Controller.Switch (52) = False Then
Controller.Switch (52) = 1
else
if Disc2.ObjRotz => 57 AND Disc2.ObjRotz <=60 Then
Controller.Switch (52) = 0
End If
End If


If Disc2.ObjRotz > 60 AND Disc2.ObjRotz < 108 AND Controller.Switch (52) = False Then
Controller.Switch (52) = 1
else
if Disc2.ObjRotz => 112 AND Disc2.ObjRotz <=115 Then
Controller.Switch (52) = 0
End If
End If


If Disc2.ObjRotz > 115 AND Disc2.ObjRotz < 158 AND Controller.Switch (52) = False Then
Controller.Switch (52) = 1
else
if Disc2.ObjRotz => 162 AND Disc2.ObjRotz <=165 Then
Controller.Switch (52) = 0
End If
End If

If Disc2.ObjRotz > 165 AND Disc2.ObjRotz < 208 AND Controller.Switch (52) = False Then
Controller.Switch (52) = 1
else
if Disc2.ObjRotz => 212 AND Disc2.ObjRotz <=215 Then
Controller.Switch (52) = 0
End If
End If

If Disc2.ObjRotz > 215 AND Disc2.ObjRotz < 258 AND Controller.Switch (52) = False Then
Controller.Switch (52) = 1
else
if Disc2.ObjRotz => 264 AND Disc2.ObjRotz <=267 Then
Controller.Switch (52) = 0
End If
End If

If Disc2.ObjRotz > 267 AND Disc2.ObjRotz < 308 AND Controller.Switch (52) = False Then
Controller.Switch (52) = 1
else
if Disc2.ObjRotz => 312 AND Disc2.ObjRotz <=315 Then
Controller.Switch (52) = 0
End If
End If

If Disc2.ObjRotz > 315 AND Disc2.ObjRotz < 353 AND Controller.Switch (52) = False Then
Controller.Switch (52) = 1
else
if Disc2.ObjRotz => 358 Then
Controller.Switch (52) = 0
End If
End If

If Disc2.ObjRotz > 352 Then
Controller.Switch (53) = 1
else
if Disc2.ObjRotz => 3 AND Disc2.ObjRotz <= 5 Then
Controller.Switch (53) = 0
End If
End If

If Controller.Switch(52) = True Then
Light61.State = 1
else
Light61.State = 0
End If

If Controller.Switch(53) = True Then
Light62.State = 1
else
Light62.State = 0
End If

Primitive3.ObjRotZ = Primitive3.ObjRotZ + 0.1
Primitive53.ObjRotZ = Primitive3.ObjRotZ + 0.1
Disc2.ObjRotZ = Disc2.ObjRotZ + 0.1
Disc1.ObjRotZ = Disc1.ObjRotZ - 0.2
End Sub

Dim tachpos

Tachpos = 1

Sub tach_timer()'255
Select Case tachpos
Case 1:
If Disc4.Rotz => 200 Then
redline.state = 2
End If
If Disc4.Rotz => 255 Then
tachpos = 2
Disc4.Rotz = 255
End If
me.enabled = 0
Disc4.rotz = Disc4.rotz + 5

Case 2:
If Disc4.Rotz <= 200 Then
redline.state = 0
End If
If Disc4.Rotz <= 0 Then
tachpos = 1
Disc4.Rotz = 0
me.enabled = 0
End If
Disc4.rotz = Disc4.rotz - 5
End Select
End Sub

Sub DivKick1_Hit()
me.DestroyBall
DivKick2.CreateBall
DivKick2.Kick 110,2
PlaySound "metalrolling2"
End Sub

Sub XGameTimer_Timer
RightFlipperP.Rotz = RightFlipper.CurrentAngle
LeftFlipperP.Rotz = LeftFlipper.CurrentAngle
End Sub

Sub XGameTimer1_Timer
RollingSoundUpdate
End Sub

Sub LRampT_Hit()
If ActiveBall.Vely < -10 Then
PlaySound "metalrolling2"
End If
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
'    JP's VP10 Rolling Sounds
'*****************************************

Const tnob = 9 ' total number of balls
ReDim rolling(tnob)
InitRolling

Sub InitRolling
    Dim i
    For i = 0 to tnob
        rolling(i) = FALSE
    Next
End Sub

Sub RollingSoundUpdate()
    Dim BOT, b
    BOT = GetBalls

    ' stop the sound of deleted balls
    For b = UBound(BOT) + 1 to tnob
        rolling(b) = FALSE
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


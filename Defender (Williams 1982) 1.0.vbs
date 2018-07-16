Option Explicit
Randomize

On Error Resume Next
ExecuteGlobal GetTextFile("controller.vbs")
If Err Then MsgBox "You need the controller.vbs in order to run this table, available in the vp10 package"
On Error Goto 0

'**********************Options********************************

dim FlipperReturnLow, FlipperReturnHigh, FlipperReturnOverride, SoundLevelMult, DesktopMode : DesktopMode = Table1.ShowDT

SoundLevelMult = 1 'Increase table SFX (may cause some normalization)

Const HardFlips = 1
Const FastFlips = 1
const FlipperReturnMod = 1 'Extra flipper return curve
FlipperReturnLow = 81	'% adjust flipper return min/max values
FlipperReturnHigh = 110	'%
FlipperReturnOverride = 0 'This can affect the rigidity of the flipper when it's at rest. 0 = Disabled

const SingleScreenFS = 0 '1 = VPX display 2 = Vpinmame display rotated
const VPXdisplay = 1 'Enable/Disable VPX display. Disable for greater performance. (Default: 1)

'*************************************************************

Ramp15.visible = Table1.ShowDT
Ramp16.visible = Table1.ShowDT


LoadVPM "", "S7.VBS", 2.0

Const cGameName="dfndr_l4"

dim VPMversion
Set MotorCallback = GetRef("UpdateSolenoids")

Sub UpdateSolenoids
	Dim Changed, Count, funcName, ii, sol11, solNo
	Changed = Controller.ChangedSolenoids
	If Not IsEmpty(Changed) Then
		sol11 = Controller.Solenoid(11)
		Count = UBound(Changed, 1)
		For ii = 0 To Count
			solNo = Changed(ii, CHGNO)
			' multiplex solenoid #11 fixed in VPM 1.52beta and newer
			if VPMversion < "01519901" then
				If SolNo < 11 And sol11 Then solNo = solNo + 32
			else
				' no need to evaluate sol 11 anymore, VPM does it
				if SolNo > 50 then solNo = solNo - 18
			end if
			funcName = SolCallback(solNo)
			If funcName <> "" Then Execute funcName & " CBool(" & Changed(ii, CHGSTATE) &")"
		Next
	End If
 BRG.ObjRotZ = BallReturnGate.CurrentAngle -5
End Sub

dim BallMass : BallMass = 1
Dim gameRun

Const SCoin="fx_coin"

Const UseSolenoids = 0
Const UseLamps = 0
Const UseGi = 0
Const UseSync = 0
Const HandleMech = 0

'***********************************************************************************
'****							Solenoid                						****
'***********************************************************************************

Const sDTLeft1	  = 1	' Drop Target 5-bank, Left Lander Drop Targets
Const sDTLeft2	  = 2	' Drop Target 5-bank, Left Lander Drop Targets
Const sDTLeft3	  = 3	' Drop Target 5-bank, Left Lander Drop Targets
Const sDTLeft4	  = 4	' Drop Target 5-bank, Left Lander Drop Targets
Const sDTLeft5	  = 5	' Drop Target 5-bank, Left Lander Drop Targets
Const sDTLeftRel  = 6	' Drop Target 5-bank, Left Lander Drop Targets
Const sDTRight1	  = 33	' Drop Target 5-bank, Right Lander Drop Targets
Const sDTRight2	  = 34	' Drop Target 5-bank, Right Lander Drop Targets
Const sDTRight3	  = 35	' Drop Target 5-bank, Right Lander Drop Targets
Const sDTRight4	  = 36	' Drop Target 5-bank, Right Lander Drop Targets
Const sDTRight5	  = 37	' Drop Target 5-bank, Right Lander Drop Targets
Const sDTRightRel = 38	' Drop Target 5-bank, Right Lander Drop Targets

Const sDTPodBot   = 7	' First Drop Down Left Upper Lane
Const sDTPodTop   = 39	' Second Drop Down Left Upper Lane
Const sLockupRel  = 40	' Lock Right Lane

Const sDTBaitBot  = 9	' Playfield Drop Target Left
Const sDTBaitTop  = 10	' Drop Down Target Top, located at the start of the Right Lane
Const sDTBaitMid  = 41	' Playfield Drop Target Right
Const sDTBaitRel  = 42	' Playfield Drop Target Reset

Const BallRelease = 8	' Ball Release
Const Drain       = 12	' Drain

Const sKickBack   = 13	' Left Auto Plunger
Const sGI 		  = 14  ' Generall Illumination
Const sBell       = 15	' Bell Sound
Const sLJet       = 17	' Left Jet Bumper
Const sRJet       = 18	' Right Jet Bumper

Const sGate	      = 22
Const sEnable     = 25

SolCallback(sDTLeft1)     = "dtLBank.SolUnHit 1,"	' Drop Target 5-bank, Left Lander Drop Targets
SolCallback(sDTLeft2)     = "dtLBank.SolUnHit 2,"	' Drop Target 5-bank, Left Lander Drop Targets
SolCallback(sDTLeft3)     = "dtLBank.SolUnHit 3,"	' Drop Target 5-bank, Left Lander Drop Targets
SolCallback(sDTLeft4)     = "dtLBank.SolUnHit 4,"	' Drop Target 5-bank, Left Lander Drop Targets
SolCallback(sDTLeft5)     = "dtLBank.SolUnHit 5,"	' Drop Target 5-bank, Left Lander Drop Targets
SolCallback(sDTLeftRel)   = "dtLBank.SolDropDown"	' Drop Target 5-bank, Left Lander Drop Targets

SolCallback(sDTRight1)    = "dtRBank.SolUnHit 1,"	' Drop Target 5-bank, Right Lander Drop Targets
SolCallback(sDTRight2)    = "dtRBank.SolUnHit 2,"	' Drop Target 5-bank, Right Lander Drop Targets
SolCallback(sDTRight3)    = "dtRBank.SolUnHit 3,"	' Drop Target 5-bank, Right Lander Drop Targets
SolCallback(sDTRight4)    = "dtRBank.SolUnHit 4,"	' Drop Target 5-bank, Right Lander Drop Targets
SolCallback(sDTRight5)    = "dtRBank.SolUnHit 5,"	' Drop Target 5-bank, Right Lander Drop Targets
SolCallback(sDTRightRel)  = "dtRBank.SolDropDown"	' Drop Target 5-bank, Right Lander Drop Targets

SolCallback(sDTBaitBot)   = "dtBait.SolUnhit 1,"	' Playfield Drop Target Left
SolCallback(sDTBaitMid)   = "dtBait.SolUnhit 2,"
SolCallback(sDTBaitTop)   = "dtBait.SolUnhit 3,"
SolCallback(sDTBaitRel)   = "dtBait.SolDropDown"	' Playfield Drop Target Reset

SolCallback(sDTPodBot)    = "dtBPod.SolDropUp"		' First Drop Down Left Upper Lane
SolCallback(sDTPodTop)    = "dtTPod.SolDropUp"		' Second Drop Down Left Upper Lane

SolCallback(sLockupRel)   = "bsLockup.SolOut"		' Lock Right Lane

SolCallback(sKickBack)    = "SolKickBack"			' Left Auto Plunger
SolCallback(sGI)          = "SolGI"

SolCallback(sBell)		  = "vpmSolSound ""Bell"","
SolCallback(sLJet)        = "vpmSolSound ""RightBumper_Hit"","
SolCallback(sRJet)        = "vpmSolSound ""RightBumper_Hit"","
SolCallback(sGate)        = "vpmSolDiverter BallReturnGate,True,"	' Ball Return Gate to Plunger Lane
SolCallback(sEnable)	  = "SolRun"


Dim bsTrough, bsLeftBallPopper, bsRightBallPopper, bsRightBallPopper42, bsRightBallPopper43, bsRightBallPopper44, bsRightLock, bsLeftLock, dtdrop, dt3bank
Dim bsLockup, dtLBank, dtRBank, dtBait, dtBaitBot, dtBPod, dtTPod 


Sub Table1_Init
	vpmInit Me
	With Controller
		.GameName = cGameName
		If Err Then MsgBox "Can't start Game" & cGameName & vbNewLine & Err.Description:Exit Sub
		.SplashInfoLine = "Space Station (Williams 1987)"
'		.Games(cGameName).Settings.Value("rol") = 0
		.HandleKeyboard = 0
		.ShowTitle = 0
		.ShowDMDOnly = 1
		.ShowFrame = 0
		.HandleMechanics = 0
'		.Hidden = 0
		.SetDisplayPosition 0,0,GetPlayerHWnd 'if you can't see the DMD then uncomment this line
		On Error Resume Next
		.Run GetPlayerHwnd
		If Err Then MsgBox Err.Description
		On Error Goto 0
	End With

	VPMversion=Controller.version

	If SingleScreenFS = 2 and not Table1.ShowDT then 
		Controller.Games(cGameName).Settings.Value("rol") = 1	
	Else
		Controller.Games(cGameName).Settings.Value("rol") = 0
		Controller.Hidden = 1
	End If
	If VPXdisplay = 0 then 
		Controller.Hidden = 0
	End If

  On Error Goto 0

	' Nudging
	vpmNudge.TiltSwitch = 1
	vpmNudge.Sensitivity = 2
	vpmNudge.TiltObj = Array(SlingL, SlingR, bumper1, bumper2)	

	'Drop Target 5-bank, Left Lander Drop Targets
	Set dtLBank = New cvpmDropTarget
	with dtLBank
		.InitDrop Array(sw13,sw14,sw15,sw16,sw17),Array(13,14,15,16,17)
		.InitSnd SoundFX("DTDrop",DOFDropTargets),SoundFX("DTReset",DOFContactors)
'		.CreateEvents"dtLBank"
    End With

	'Drop Target 5-bank, Right Lander Drop Targets
	Set dtRBank = New cvpmDropTarget
	with dtRBank
		.InitDrop Array(sw23,sw24,sw25,sw26,sw27),Array(23,24,25,26,27)
		.InitSnd SoundFX("DTDrop",DOFDropTargets),SoundFX("DTReset",DOFContactors)
'		.CreateEvents"dtRBank"
    End With

	'Drop Target Left Upper Lane
	Set dtBPod = New cvpmDropTarget
	with dtBPod
		.InitDrop sw39,39
		.InitSnd SoundFX("DTDrop",DOFContactors),SoundFX("DTReset",DOFContactors)
    End With

	Set dtTPod = New cvpmDropTarget
	with dtTPod
		.InitDrop sw40,40
		.InitSnd SoundFX("DTDrop",DOFContactors),SoundFX("DTReset",DOFContactors)
    End With

	Set bsLockup = New cvpmBallStack
		bsLockup.InitSw 0, 42,43,44,0,0,0,0
		bsLockup.InitKick sw42,180,1
		bsLockup.InitExitSnd SoundFX("Popper",DOFContactors), SoundFX("Solenoid",DOFContactors)

	Set dtBait = New cvpmDropTarget
	with dtBait
		.InitDrop Array(Array(sw33,sw33a),Array(sw34,sw34a),sw35),Array(33,34,35)
		.InitSnd SoundFX("DTDrop",DOFContactors),SoundFX("DTReset",DOFContactors)
    End With

     ' Main Timer init
     PinMAMETimer.Interval = PinMAMEInterval
     PinMAMETimer.Enabled = 1
	 LampTimer.Enabled = 1

	' Init Kickback
    KickBack.Pullback

	'start trough
	sw49.CreateSizedBallWithMass 25, BallMass
	Sw48.CreateSizedBallWithMass 25, BallMass
	sw47.CreateSizedBallWithMass 25, BallMass

	BallSearch	'init switches
	
	gameRun = False
End Sub

Sub table1_Paused:Controller.Pause = 1:End Sub
Sub table1_unPaused:Controller.Pause = 0:End Sub
Sub Table1_exit():Controller.Pause = False:Controller.Stop:End Sub
Sub Destroyer_Hit:Me.destroyball : End Sub

'***********************************************************************************
'****								Plunger & Flipper      						****
'***********************************************************************************

Sub Table1_KeyDown(ByVal keycode)
	If vpmKeyDown(keycode) Then Exit Sub
    If keycode = PlungerKey Then PlaySound SoundFx("plungerpull",0),0,LVL(1),0.25,0.25:Plunger.Pullback:End If
	If keycode = LeftFlipperKey Then flipnf 0, 1: exit sub : End If
	If keycode = RightFlipperKey Then flipnf 1, 1 : controller.Switch(59) = 1 : exit sub : End If
	if Keycode = KeyRules then 
		If DesktopMode Then
			P_InstructionsDT.PlayAnim 0, 3*CGT/500
		Else
			P_InstructionsFS.PlayAnim 0, 3*CGT/500
		End If
		Exit Sub
	End If
	If keycode = RightMagnaSave Then controller.Switch(60) = 1
	If keycode = LeftMagnaSave Then controller.Switch(61) = 1
End Sub

Sub Table1_KeyUp(ByVal keycode)
    If vpmKeyUp(keycode) Then Exit Sub
    If keycode = PlungerKey Then 
		Plunger.Fire
		if BallInPlunger then 
			PlaySound SoundFX("plunger3",0),0,LVL(1),0.05,0.02
		Else
			PlaySound SoundFX("plunger",0),0,LVL(0.8),0.05,0.02
		end if
	End If
	if Keycode = KeyRules then 
		If DesktopMode Then
			P_InstructionsDT.ShowFrame 0
		Else
			P_InstructionsFs.ShowFrame 0
		End If
		Exit Sub
	End If

    If keycode = LeftFlipperKey Then : flipnf 0, 0: exit sub : End If
	If keycode = RightFlipperKey Then : flipnf 1, 0: controller.Switch(59) = 0 : exit sub:End If

	If keycode = RightMagnaSave Then controller.Switch(60) = 0
	If keycode = LeftMagnaSave Then controller.Switch(61) = 0
End Sub


'***********************************************************************************
'****								Trough Handling      						****
'***********************************************************************************

SolCallback(BallRelease)   = "TroughOut"      ' Ball Release
SolCallback(Drain)	       = "TroughIn"       ' Drain

Sub TroughIn(enabled)
	if Enabled then 
		sw46.Kick 58, 16 : 
		Tgate.Open = True
		If sw46.BallCntOver > 0 Then Playsound SoundFX("Trough1",DOFcontactors), 0, LVL(0.5), 0.01 : End If
	Else 
		Tgate.Open = False
		BallSearch
	End If 
end Sub

Sub TroughOut(enabled)
	if Enabled then 
		sw47.Kick 58, 8 :Playsound SoundFX("BallRelease",DOFcontactors), 0, LVL(0.4), 0.02
	End If
end Sub

'***********************************************************************************
'****						Ball Ramp Trough Switches      						****
'***********************************************************************************

Sub Sw46_hit():controller.Switch(46) = 1 : End Sub	'Drain
Sub TroughSFX_Hit(): Playsound "Trough2", 0, LVL(0.2), 0:: End Sub
Sub Sw49_hit():controller.Switch(49) = 1 : UpdateTrough : End Sub
Sub Sw48_hit():controller.Switch(48) = 1 : UpdateTrough : End Sub
Sub Sw47_hit():controller.Switch(47) = 1 : UpdateTrough : End Sub

Sub Sw46_UnHit():controller.Switch(46) = 0 : UpdateTrough : End Sub
Sub Sw49_UnHit():controller.Switch(49) = 0 : UpdateTrough : End Sub
Sub Sw48_UnHit():controller.Switch(48) = 0 : UpdateTrough : End Sub
Sub Sw47_UnHit():controller.Switch(47) = 0 : UpdateTrough : End Sub

Sub UpdateTrough: TroughTimer.enabled = 1 : TroughTimer.Interval = 200: end sub

Sub TroughTimer_Timer()
	If sw47.BallCntOver = 0 then sw48.kick 58, 12
	If sw48.BallCntOver = 0 then sw49.kick 58, 12
End Sub

Sub BallSearch()	'In case of hard pinmame reset. Called by PF solenoids firing empty.
	if Sw46.BallCntOver > 0 then controller.Switch(46) = 1 else controller.Switch(46) = 0
	if Sw49.BallCntOver > 0 then controller.Switch(49) = 1 else controller.Switch(49) = 0
	if Sw48.BallCntOver > 0 then controller.Switch(48) = 1 else controller.Switch(48) = 0
	if Sw47.BallCntOver > 0 then controller.Switch(47) = 1 else controller.Switch(47) = 0
End Sub

'***********************************************************************************
'****								Rollover Switches      						****
'***********************************************************************************

'Primitives from Victory (Gottlieb 1987)

Sub sw9_Hit():PlaySound "Sensor":Controller.Switch(9) = 1:sw9p.Visible = 0:End Sub
Sub sw9_UnHit():Controller.Switch(9) = 0:sw9p.Visible = 1:End Sub
Sub sw10_Hit():PlaySound "Sensor":Controller.Switch(10) = 1:sw10p.Visible = 0:End Sub
Sub sw10_UnHit():Controller.Switch(10) = 0:sw10p.Visible = 1:End Sub
Sub sw11_Hit():PlaySound "Sensor":Controller.Switch(11) = 1:sw11p.Visible = 0:End Sub
Sub sw11_UnHit():Controller.Switch(11) = 0:sw11p.Visible = 1:End Sub
Sub sw12_Hit():PlaySound "Sensor":Controller.Switch(12) = 1:sw12p.Visible = 0:End Sub
Sub sw12_UnHit():Controller.Switch(12) = 0:sw12p.Visible = 1:End Sub

Sub sw45_Hit():PlaySound "Sensor":Controller.Switch(45) = 1:sw45p.Visible = 0:End Sub
Sub sw45_UnHit():Controller.Switch(45) = 0:sw45p.Visible = 1:End Sub
Sub sw51_Hit():PlaySound "Sensor":Controller.Switch(51) = 1:sw51p.Visible = 0:End Sub
Sub sw51_UnHit():Controller.Switch(51) = 0:sw51p.Visible = 1:End Sub
Sub sw52_Hit():PlaySound "Sensor":Controller.Switch(52) = 1:sw52p.Visible = 0:End Sub
Sub sw52_UnHit():Controller.Switch(52) = 0:sw52p.Visible = 1:End Sub
Sub sw53_Hit():PlaySound "Sensor":Controller.Switch(53) = 1:sw53p.Visible = 0:End Sub
Sub sw53_UnHit():Controller.Switch(53) = 0:sw53p.Visible = 1:End Sub
Sub sw54_Hit():PlaySound "Sensor":Controller.Switch(54) = 1:sw54p.Visible = 0:End Sub
Sub sw54_UnHit():Controller.Switch(54) = 0:sw54p.Visible = 1:End Sub

'***********************************************************************************
'****								Stand Up Targets       						****
'***********************************************************************************

Sub sw18_hit:vpmTimer.pulseSw 18:PlaySound "target", 0, LVL(0.1), 0.3:End Sub 
Sub sw19_hit:vpmTimer.pulseSw 19:PlaySound "target", 0, LVL(0.1), 0.3:End Sub 
Sub sw20_hit:vpmTimer.pulseSw 20:PlaySound "target", 0, LVL(0.1), 0.3:End Sub 
Sub sw21_hit:vpmTimer.pulseSw 21:PlaySound "target", 0, LVL(0.1), 0.3:End Sub 
Sub sw22_hit:vpmTimer.pulseSw 22:PlaySound "target", 0, LVL(0.1), 0.3:End Sub 

Sub sw28_hit:vpmTimer.pulseSw 28:PlaySound "target", 0, LVL(0.1), 0.3:End Sub 
Sub sw29_hit:vpmTimer.pulseSw 29:PlaySound "target", 0, LVL(0.1), 0.3:End Sub 
Sub sw30_hit:vpmTimer.pulseSw 30:PlaySound "target", 0, LVL(0.1), 0.3:End Sub 
Sub sw31_hit:vpmTimer.pulseSw 31:PlaySound "target", 0, LVL(0.1), 0.3:End Sub 
Sub sw32_hit:vpmTimer.pulseSw 32:PlaySound "target", 0, LVL(0.1), 0.3:End Sub 

Sub sw36_hit:vpmTimer.pulseSw 36:PlaySound "target", 0, LVL(0.1), 0.3:End Sub 
Sub sw37_hit:vpmTimer.pulseSw 37:PlaySound "target", 0, LVL(0.1), 0.3:End Sub 

Sub sw41_hit:vpmTimer.pulseSw 41:PlaySound "target", 0, LVL(0.1), 0.3:End Sub 


'***********************************************************************************
'****								Drop Targets         						****
'***********************************************************************************

'Left Lander Targets
Sub Sw13_Dropped:dtLBank.Hit 1 : End Sub
Sub Sw14_Dropped:dtLBank.Hit 2 : End Sub  
Sub Sw15_Dropped:dtLBank.Hit 3 : End Sub
Sub Sw16_Dropped:dtLBank.Hit 4 : End Sub
Sub Sw17_Dropped:dtLBank.Hit 5 : End Sub

'Right Lander Targets
Sub sw23_Dropped:dtRBank.Hit 1 : End Sub
Sub Sw24_Dropped:dtRBank.Hit 2 : End Sub  
Sub Sw25_Dropped:dtRBank.Hit 3 : End Sub
Sub Sw26_Dropped:dtRBank.Hit 4 : End Sub
Sub Sw27_Dropped:dtRBank.Hit 5 : End Sub

'Baiter Targets
Sub sw33_Dropped : dtBait.Hit 1  : End Sub
Sub sw34_Dropped : dtBait.Hit 2  : End Sub
Sub sw35_Dropped : dtBait.Hit 3  : End Sub

'Pod Targets
Sub sw39_Dropped:dtBPod.Hit 1 : End Sub
Sub sw40_Dropped:dtTPod.Hit 1 : End Sub

Sub Sw33a_Hit() : Playsound "soloff", 0, LVL(0.2), 0:: End Sub 'dont drop target but play sound
Sub Sw34a_Hit() : Playsound "soloff", 0, LVL(0.2), 0:: End Sub 'dont drop target but play sound

'***********************************************************************************
'****								Switches		      						****
'***********************************************************************************

Sub sw38_hit:vpmTimer.pulseSw 38 : PlaySound "soloff", 0, LVL(0.1), 0.3:End Sub 

'***********************************************************************************
'****								Bumper              						****
'***********************************************************************************

Const swJet1 = 55 ' Left Jet Bumper
Const swJet2 = 56 ' Right Jet Bumper

Sub Bumper1_Hit():vpmTimer.PulseSwitch (swJet1), 0, "" : PlaySound SoundFX("LeftBumper_Hit",DOFContactors) : End Sub
Sub Bumper2_Hit():vpmTimer.PulseSwitch (swJet2), 0, "" : PlaySound SoundFX("RightBumper_Hit",DOFContactors) : End Sub


'***********************************************************************************
'****								Drain hole & Kicker    						****
'*********************************************************************************** 

Sub sw42_Hit:bsLockup.AddBall me : playsound "popper_ball": End Sub

'***********************************************************************************
'****								Knocker               						****
'*********************************************************************************** 

Sub KnockerSol(enabled)
	If enabled Then PlaySound SoundFX("Knocker",DOFKnocker)
End Sub

'***********************************************************************************
'****								KickBack               						****
'*********************************************************************************** 
	
Sub SolKickBack(enabled)
    If enabled Then
       Kickback.Fire
       PlaySound SoundFX("DiverterLeft",DOFContactors), 0, LVL(1), -0.02
    Else
       KickBack.PullBack
    End If
End Sub

'***********************************************************************************
'****								Shooter Lane           						****
'*********************************************************************************** 

Sub sw50_Hit():PlaySound "Sensor":Controller.Switch(50) = 1:sw50p.Visible = 0:End Sub
Sub sw50_UnHit():Controller.Switch(50) = 0:sw50p.Visible = 1:End Sub

dim BallInPlunger :BallInPlunger = False
sub PlungerLane_hit():ballinplunger = True: End Sub
Sub PlungerLane_unhit():BallInPlunger = False: End Sub

'===================
'	NF Flipper Return Mod 2017
'===================
dim returnspeed, lfstep, rfstep, LFstep1
returnspeed = cInt(leftflipper.return*100)
LFstep = 1
LFstep1 = 1
RFstep = 1

sub LeftFlipper_timer()
	select case LFstep
		Case 1: LeftFlipper.return = ReturnCurve(LFstep) :LFstep = LFstep + 1
		Case 2: LeftFlipper.return = ReturnCurve(LFstep) :LFstep = LFstep + 1
		Case 3: LeftFlipper.return = ReturnCurve(LFstep) :LFstep = LFstep + 1
		Case 4: LeftFlipper.return = ReturnCurve(LFstep) :LFstep = LFstep + 1
		Case 5: LeftFlipper.return = ReturnCurve(LFstep) :LFstep = LFstep + 1
		Case 6: LeftFlipper.return = ReturnCurve(LFstep) :LFstep = LFstep + 1
		Case 7: LeftFlipper.timerenabled = 0 : LFstep = 1
	end select
'	tb2.text = LeftFlipper.Return
end sub

sub LeftFlipper1_timer()
	select case LFstep1
		Case 1: LeftFlipper1.return = ReturnCurve(LFstep1) :LFstep1 = LFstep1 + 1
		Case 2: LeftFlipper1.return = ReturnCurve(LFstep1) :LFstep1 = LFstep1 + 1
		Case 3: LeftFlipper1.return = ReturnCurve(LFstep1) :LFstep1 = LFstep1 + 1
		Case 4: LeftFlipper1.return = ReturnCurve(LFstep1) :LFstep1 = LFstep1 + 1
		Case 5: LeftFlipper1.return = ReturnCurve(LFstep1) :LFstep1 = LFstep1 + 1
		Case 6: LeftFlipper1.return = ReturnCurve(LFstep1) :LFstep1 = LFstep1 + 1
		Case 7: LeftFlipper1.timerenabled = 0 : LFstep1 = 1
	end select
'	tb2.text = LeftFlipper.Return
end sub

sub RightFlipper_timer()
	select case RFstep
		Case 1: RightFlipper.return = ReturnCurve(RFstep) :RFstep = RFstep + 1
		Case 2: RightFlipper.return = ReturnCurve(RFstep) :RFstep = RFstep + 1
		Case 3: RightFlipper.return = ReturnCurve(RFstep) :RFstep = RFstep + 1
		Case 4: RightFlipper.return = ReturnCurve(RFstep) :RFstep = RFstep + 1
		Case 5: RightFlipper.return = ReturnCurve(RFstep) :RFstep = RFstep + 1
		Case 6: RightFlipper.return = ReturnCurve(RFstep) :RFstep = RFstep + 1
		Case 7: RightFlipper.timerenabled = 0 : RFstep = 1
	end select
end sub

FlipperReturnLow = cInt(FlipperReturnLow)	'% adjust flipper return min/max values

Function ReturnCurve(step)	'Adjust flipper curve
	dim x, s
	s = FlipperReturnHigh - FlipperReturnLow
	Select Case step
		case 0 'initial curve
			x = FlipperReturnLow
		case 1 '16ms
			x = FlipperReturnLow + ((s/6) * 1)
		case 2 '32ms
			x = FlipperReturnLow + ((s/6) * 2)
		case 3 'etc
			x = FlipperReturnLow + ((s/6) * 3)
		case 4
			x = FlipperReturnLow + ((s/6) * 4)
		case 5
			x = FlipperReturnLow + ((s/6) * 5)
		case 6
			if FlipperReturnOverride > 0 then 
				ReturnCurve = FlipperReturnOverride
				Exit Function
			else
				x = FlipperReturnHigh
			End If
	End Select
	ReturnCurve = (returnspeed * x) / 10000
'	tb.text = x & vbnewline & ReturnCurve
End Function

'====================
'   HARD
'       FLIPS
'====================
'just switches EOStorque when hit
dim defaultEOS, hardEOS
defaulteos = leftflipper.eostorque
hardeos = 2200 / leftflipper.strength	'eos equivalent to 2200 strength
hardeos = 2200 / leftflipper1.strength	'eos equivalent to 2200 strength
if hardflips = 0 then TriggerLF.enabled = 0 :TriggerLF1.enabled = 0 : TriggerRF.enabled = 0 end if

Sub TriggerLF_Hit()
    dim finalspeed
    finalspeed = SQR(activeball.velx * activeball.velx + activeball.vely * activeball.vely)
    If finalspeed > 0 then
		leftflipper.eostorque = hardeos
    End if
end sub

Sub TriggerLF1_Hit()
    dim finalspeed
    finalspeed = SQR(activeball.velx * activeball.velx + activeball.vely * activeball.vely)
    If finalspeed > 0 then
		LeftFlipper1.eostorque = hardeos
    End if
end sub

Sub TriggerRF_Hit()
    dim finalspeed
    finalspeed = SQR(activeball.velx * activeball.velx + activeball.vely * activeball.vely)
    If finalspeed > 0 then
		rightflipper.eostorque = hardeos
    End if
end sub

'***********************************************************************************
'****		nFozzy's FastFlips NF 'Pre-solid state flippers lag reduction		****
'***********************************************************************************

dim FlippersEnabled

Sub SolRun(enabled)
	FlippersEnabled = Enabled
	if enabled then
		gameRun = True
'		PFGI.State = 1
		SLingL.Disabled = 0
		SLingR.Disabled = 0
	Elseif not Enabled Then
		gameRun = false
'		PFGI.State = 0
		if leftflipper.startangle > leftflipper.endangle Then
			if leftflipper.currentangle < leftflipper.startangle then leftflipper.rotatetostart : leftflippersound 0 : end if
		elseif leftflipper.startangle < leftflipper.endangle Then
			if leftflipper.currentangle > leftflipper.startangle then leftflipper.rotatetostart : leftflippersound 0 : end If
		end If
		if leftflipper1.startangle > leftflipper1.endangle Then
			if leftflipper1.currentangle < leftflipper1.startangle then leftflipper1.rotatetostart : end if
		elseif leftflipper1.startangle < leftflipper1.endangle Then
			if leftflipper1.currentangle > leftflipper1.startangle then leftflipper1.rotatetostart : end If
		end If
		if rightflipper.startangle > rightflipper.endangle Then
			if rightflipper.currentangle < rightflipper.startangle then rightflipper.rotatetostart : rightflippersound 0 : end if
		elseif rightflipper.startangle < rightflipper.endangle Then
			if rightflipper.currentangle > rightflipper.startangle then rightflipper.rotatetostart : rightflippersound 0 : end If
		end If
		SLingL.Disabled = 1
		SLingR.Disabled = 1
	end if
	vpmNudge.SolGameOn enabled
End Sub

sub flipnf(LR, DU)	'object, left-right, downup
	if FastFlips = 0 then 
		if LR = 0 then
			controller.switch(LLFlip) = DU
		elseif LR = 1 then
			controller.switch(LRFlip) = DU	
		End If
		exit Sub
	end if
	if LR = 0 Then		'left flipper
		if DU = 1 then
			If FlippersEnabled = True then
				leftflipper.rotatetoend
				leftflipper1.rotatetoend
				LeftFlipperSound 1
			end if
			controller.Switch(swLLFlip) = True
		Elseif DU = 0 then
			If FlippersEnabled = True then
				leftflipper.rotatetoStart
				leftflipper1.rotatetoStart
				LeftFlipperSound 0
			end if
			controller.Switch(swLLFlip) = False							
		end if
	elseif LR = 1 then		''right flipper
		if DU = 1 then
			If FlippersEnabled = True then
				RightFlipper.rotatetoend
				RightFlipperSound 1
			end if
			controller.Switch(swLRFlip) = True
		Elseif DU = 0 then
			If FlippersEnabled = True then
				RightFlipper.rotatetoStart
				RightFlipperSound 0
			end if
			controller.Switch(swLRFlip) = False						
		end if
	end if
end sub

sub LeftFlipperSound(updown)'called along with the flipper, so feel free to add stuff, EOStorque tweaks, animation updates, upper flippers, whatever.
	if updown = 1 Then
		playsound SoundFX("FlipperUp",DOFFlippers)	'flip
	Else
		playsound SoundFX("FlipperDown",DOFFlippers)	'return
		LeftFlipper.eostorque = defaultEOS
		LeftFlipper1.eostorque = defaultEOS
			if FlipperReturnMod = 1 then
				LeftFlipper.TimerEnabled = 1
				LeftFlipper1.TimerEnabled = 1
				LeftFlipper.TimerInterval = 16	'400 test
				LeftFlipper1.TimerInterval = 16	'400 test
				LeftFlipper.return = FlipperReturnLow / 10000
				LeftFlipper1.return = FlipperReturnLow / 10000
				LFstep = 1
				LFstep1 = 1
			end if
	end if
end sub
sub RightFlipperSound(updown)
	if updown = 1 Then
		playsound SoundFX("FlipperUp",DOFFlippers), 0, LVL(1), 0.01	'flip
	Else
		playsound SoundFX("FlipperDown",DOFFlippers), 0, LVL(1), 0.01	'return
		RightFlipper.eostorque = defaultEOS
		if FlipperReturnMod = 1 then
			rightflipper.TimerEnabled = 1
			rightflipper.TimerInterval = 16
			rightflipper.return = FlipperReturnLow / 10000
			RFstep = 1
		end if
	end if
end sub

'***********************************************************************************
'****				Special JP Flippers 'Legacy Flippers using callbacks		****
'***********************************************************************************

 SolCallback(sLRFlipper) = "SolRFlipper"
 SolCallback(sLLFlipper) = "SolLFlipper"

Sub SolLFlipper(Enabled)
	if FastFlips = 0 then 
		If Enabled Then
			playsound SoundFX("FlipperUp",DOFFlippers), 0, LVL(1), -0.01	'flip
			LeftFlipper.RotateToEnd
			LeftFlipper1.RotateToEnd
		Else
			playsound SoundFX("FlipperDown",DOFFlippers), 0, LVL(1), -0.01	'return
			LeftFlipper.RotateToStart
			LeftFlipper1.RotateToStart
			LeftFlipper.eostorque = defaulteos
			LeftFlipper1.eostorque = defaulteos
			if FlipperReturnMod = 1 then
				LeftFlipper.TimerEnabled = 1
				LeftFlipper1.TimerEnabled = 1
				LeftFlipper.TimerInterval = 16
				LeftFlipper1.TimerInterval = 16
				LeftFlipper.return = FlipperReturnLow / 10000
				LeftFlipper1.return = FlipperReturnLow / 10000
				LFstep = 1
				LFstep1 = 1
			end if
		End If
	End If
End Sub

Sub SolRFlipper(Enabled)
	if FastFlips = 0 then 
		If Enabled Then
			playsound SoundFX("FlipperUp",DOFFlippers), 0, LVL(1), 0.01	'flip
			RightFlipper.RotateToEnd
		Else
			playsound SoundFX("FlipperDown",DOFFlippers), 0, LVL(1), 0.01	'return
			RightFlipper.RotateToStart
			RightFlipper.eostorque = defaulteos 
			if FlipperReturnMod = 1 then
				rightflipper.TimerEnabled = 1
				rightflipper.TimerInterval = 16
				rightflipper.return = FlipperReturnLow / 10000
				RFstep = 1
			end if
		End If
	End If
End Sub

'***********************************************************************************
'****							Slingshot               						****
'*********************************************************************************** 
Dim SLPos,SRPos

Sub SlingL_Slingshot()
'	If FlipperEnabled = False Then Exit Sub
	LSling1.Visible = 0:LSling2.Visible = 0:LSling3.Visible = 0: LSling4.Visible = 1: LSling.TransZ = -27
	vpmTimer.PulseSw 57
	PlaySound SoundFX ("left_slingshot",DOFContactors),1:DOF dLeftSlingshot, 2
	SLPos = 0: Me.TimerEnabled = 1
'	LightshowChangeSide
End Sub 

Sub SlingL_Timer
    Select Case SLPos
        Case 2:	LSling1.Visible = 0:LSling2.Visible = 0:LSling3.Visible = 1: LSling4.Visible = 0: LSling.TransZ = -17
		Case 3:	LSling1.Visible = 0:LSling2.Visible = 1:LSling3.Visible = 0: LSling4.Visible = 0: LSling.TransZ = -8
		Case 4:	LSling1.Visible = 1:LSling2.Visible = 0:LSling3.Visible = 0: LSling4.Visible = 0: LSling.TransZ = 0 :Me.TimerEnabled = 0
    End Select
    SLPos = SLPos + 1
End Sub

Sub SlingR_Slingshot()
'	If FlipperEnabled = False Then Exit Sub
	RSling1.Visible = 0:RSling2.Visible = 0:RSling3.Visible = 0: RSling4.Visible = 1: RSling.TransZ = -27
	vpmTimer.PulseSw 58
	PlaySound SoundFX ("right_slingshot",DOFContactors):DOF dRightSlingshot, 2
	SRPos = 0: Me.TimerEnabled = 1
'	LightshowChangeSide
End Sub 

Sub SlingR_Timer
    Select Case SRPos
        Case 2:	RSling1.Visible = 0:RSling2.Visible = 0:RSling3.Visible = 1: RSling4.Visible = 0: RSling.TransZ = -17
		Case 3:	RSling1.Visible = 0:RSling2.Visible = 1:RSling3.Visible = 0: RSling4.Visible = 0: RSling.TransZ = -8
		Case 4:	RSling1.Visible = 1:RSling2.Visible = 0:RSling3.Visible = 0: RSling4.Visible = 0: RSling.TransZ = 0 :Me.TimerEnabled = 0
    End Select
    SRPos = SRPos + 1
End Sub

Sub Slingshots_Init
	LSling1.Visible = 1:LSling2.Visible = 0:LSling3.Visible = 0: LSling4.Visible = 0: LSling.TransZ = 0
	RSling1.Visible = 1:RSling2.Visible = 0:RSling3.Visible = 0: RSling4.Visible = 0: RSling.TransZ = 0
End Sub

'***************************************************
'       JP's VP10 Fading Lamps & Flashers
'       Based on PD's Fading Light System
' SetLamp 0 is Off
' SetLamp 1 is On
' fading for non opacity objects is 4 steps
'***************************************************

'Dim bulb
Dim LampState(200), FadingLevel(200)
Dim FlashSpeedUp(200), FlashSpeedDown(200), FlashMin(200), FlashMax(200), FlashLevel(200)

InitLamps()             ' turn off the lights and flashers and reset them to the default parameters
LampTimer.Interval = 10 'lamp fading speed
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

Sub InitLamps()
    Dim x
    For x = 0 to 200
        LampState(x) = 0         ' current light state, independent of the fading level. 0 is off and 1 is on
        FadingLevel(x) = 4       ' used to track the fading state
        FlashSpeedUp(x) = 0.5    ' faster speed when turning on the flasher
        FlashSpeedDown(x) = 0.1 ' slower speed when turning off the flasher
        FlashMax(x) = 1          ' the maximum value when on, usually 1
        FlashMin(x) = 0          ' the minimum value when off, usually 0
        FlashLevel(x) = 0        ' the intensity of the flashers, usually from 0 to 1
    Next
End Sub

Sub UpdateLamps
'    On Error Resume Next
	nFadeLm 9, Light9
	nFadeLm 9, Light9a
	nFadeLm 10, Light10
	nFadeLm 10, Light10a
	nFadeLm 11, Light11
	nFadeLm 11, Light11a
	nFadeLm 12, Light12
	nFadeLm 12, Light12a
	nFadeLm 13, Light13
	nFadeLm 13, Light13a
	nFadeLm 14, Light14
	nFadeLm 14, Light14a
	nFadeLm 15, Light15
	nFadeLm 15, Light15a
	nFadeLm 16, Light16
	nFadeLm 16, Light16a
	nFadeLm 17, Light17
	nFadeLm 17, Light17a
	nFadeLm 18, Light18
	nFadeLm 18, Light18a
	nFadeLm 19, Light19
	nFadeLm 19, Light19a
	nFadeLm 20, Light20
	nFadeLm 20, Light20a
	nFadeLm 21, Light21
	nFadeLm 21, Light21a
	nFadeLm 22, Light22
	nFadeLm 22, Light22a
	nFadeLm 23, Light23
	nFadeLm 23, Light23a
	nFadeLm 24, Light24
	nFadeLm 24, Light24a
	nFadeLm 25, Light25
	nFadeLm 25, Light25a
	nFadeLm 26, Light26
	nFadeLm 26, Light26a
	nFadeLm 27, Light27
	nFadeLm 27, Light27a
	nFadeLm 28, Light28
	nFadeLm 28, Light28a
	nFadeLm 29, Light29
	nFadeLm 29, Light29a
	nFadeLm 30, Light30
	nFadeLm 30, Light30a
	nFadeLm 31, Light31
	nFadeLm 31, Light31a
	nFadeLm 32, Light32
	nFadeLm 32, Light32a
	nFadeLm 33, Light33
	nFadeLm 33, Light33a
	nFadeLm 34, Light34
	nFadeLm 34, Light34a
	nFadeLm 35, Light35
	nFadeLm 35, Light35a
	NFadeLm 36, Light36
	NFadeLm 36, Light36a
	NFadeLm 37, Light37
	NFadeLm 37, Light37a
	NFadeLm 38, Light38
	NFadeLm 38, Light38a
	NFadeLm 39, Light39
	NFadeLm 39, Light39a
	FlashC 39, Flash39b
	NFadeLm 40, Light40
	NFadeLm 40, Light40a
	FlashC 40, Flash40b
	NFadeLm 41, Light41
	NFadeLm 41, Light41a
	FlashC 41, Flash41b
	NFadeLm 42, Light42
	NFadeLm 42, Light42a
	FlashC 42, Flash42b
	nFadeLm 43, Light43
	nFadeLm 43, Light43a
	nFadeLm 44, Light44
	nFadeLm 44, Light44a
	nFadeLm 45, Light45
	nFadeLm 45, Light45a
	NFadeLm 46, Light46
	NFadeLm 46, Light46a
	NFadeLm 47, Light47
	NFadeLm 47, Light47a
	NFadeLm 48, Light48
	NFadeLm 48, Light48a
	NFadeLm 49, Light49
	NFadeLm 49, Light49a
	NFadeLm 50, Light50
	NFadeLm 50, Light50a
	NFadeLm 51, Light51
	NFadeLm 51, Light51a
	NFadeLm 52, Light52
	NFadeLm 52, Light52a
	NFadeLm 53, Light53
	NFadeLm 53, Light53a
	NFadeLm 54, Light54
	NFadeLm 54, Light54a
	nFadeLm 55, Light55
	nFadeLm 55, Light55a
	nFadeLm 56, Light56
	nFadeLm 56, Light56a
	nFadeLm 57, Light57
	nFadeLm 57, Light57a
	nFadeLm 58, Light58
	nFadeLm 58, Light58a
	nFadeLm 59, Light59
	nFadeLm 59, Light59a
	nFadeLm 60, Light60
	nFadeLm 60, Light60a
	nFadeLm 61, Light61
	nFadeLm 61, Light61a
	nFadeLm 62, Light62
	nFadeLm 62, Light62a
	Flashc 49, Flash49
	Flashc 50, Flash50
	Flashc 51, Flash51
	Flashc 52, Flash52
	Flashc 53, Flash53
	Flashc 54, Flash54
	Flashc 60, Flash60
	Flashc 61, Flash61
	Flashc 61, Flash61a

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

Sub Flashc(nr, object)
    Select Case FadingLevel(nr)
		Case 3
            FadingLevel(nr) = 0 'completely off
            Object.IntensityScale = FlashLevel(nr)
        Case 4 'off
            FlashLevel(nr) = FlashLevel(nr) - (FlashSpeedDown(nr) * CGT)
            If FlashLevel(nr) < FlashMin(nr) Then
                FlashLevel(nr) = FlashMin(nr)
               FadingLevel(nr) = 3 'completely off
            End if
            Object.IntensityScale = FlashLevel(nr)
        Case 5 ' on
            FlashLevel(nr) = FlashLevel(nr) + (FlashSpeedUp(nr) * CGT)
            If FlashLevel(nr) > FlashMax(nr) Then
                FlashLevel(nr) = FlashMax(nr)
                FadingLevel(nr) = 6 'completely on
            End if
            Object.IntensityScale = FlashLevel(nr)
        Case 6 ' on
			FadingLevel(nr) = 1 'completely on
			Object.IntensityScale = FlashLevel(nr)
    End Select
End Sub

dim DisplayColor, DisplayColorG
InitDisplay
Sub InitDisplay()
	dim x
	for each x in Display
		x.Color = RGB(1,1,1)
	next
	Displaytimer.enabled = 1
	DisplayColor = dx.Color
	DisplayColorG = dxG.Color 
	If DesktopMode then 
		for each x in Display
			x.height = 380
			x.x = x.x - 956
			x.y = x.y -37
			x.rotx = -42
			x.visible = 1
		next
		for each x in Display2
			x.x = x.x +109
			x.y = x.y + 1
		next
		for each x in Display3
			x.height = 325
			x.x = x.x +32
			x.y = x.y +30
		next
		for each x in Display4
			x.height = 325
			x.x = x.x +125
			x.y = x.y +30
		next
	Else
		Select Case SingleScreenFS
			Case 1
				Displaytimer.enabled = 1
				for each x in Display
					x.visible = 1
				next
			case 0, 2
				Displaytimer.enabled = 0
				for each x in Display
					x.visible = 0
				next		
		End Select
	End If

	if VPXdisplay = 0 Then
		Displaytimer.enabled = 0
		for each x in Display
			x.visible = 0
		next
	End If
End Sub

Sub Displaytimer_Timer
	On Error Resume Next
	Dim ii, jj, obj, b, x
	Dim ChgLED,num, chg, stat
	ChgLED=Controller.ChangedLEDs(&Hffffffff, &Hffffffff)
		If Not IsEmpty(ChgLED) Then
			For ii=0 To UBound(chgLED)
				num=chgLED(ii, 0) : chg=chgLED(ii, 1) : stat=chgLED(ii, 2)
				For Each obj In Digits(num)
					If chg And 1 Then FadeDisplay obj, stat And 1	
					chg=chg\2 : stat=stat\2
				Next
			Next
		End If
End Sub

Sub FadeDisplay(object, onoff)
	If OnOff = 1 Then
		object.color = DisplayColor
	Else
		Object.Color = RGB(1,1,1)
	End If
End Sub

Dim Digits(32)
Digits(0)=Array(D1,D2,D3,D4,D5,D6,D7,D8,D9,D10,D11,D12,D13,D14,D15)
Digits(1)=Array(D16,D17,D18,D19,D20,D21,D22,D23,D24,D25,D26,D27,D28,D29,D30)
Digits(2)=Array(D31,D32,D33,D34,D35,D36,D37,D38,D39,D40,D41,D42,D43,D44,D45)
Digits(3)=Array(D46,D47,D48,D49,D50,D51,D52,D53,D54,D55,D56,D57,D58,D59,D60)
Digits(4)=Array(D61,D62,D63,D64,D65,D66,D67,D68,D69,D70,D71,D72,D73,D74,D75)
Digits(5)=Array(D76,D77,D78,D79,D80,D81,D82,D83,D84,D85,D86,D87,D88,D89,D90)
Digits(6)=Array(D91,D92,D93,D94,D95,D96,D97,D98,D99,D100,D101,D102,D103,D104,D105)
Digits(7)=Array(D106,D107,D108,D109,D110,D111,D112,D113,D114,D115,D116,D117,D118,D119,D120)
Digits(8)=Array(D121,D122,D123,D124,D125,D126,D127,D128,D129,D130,D131,D132,D133,D134,D135)
Digits(9)=Array(D136,D137,D138,D139,D140,D141,D142,D143,D144,D145,D146,D147,D148,D149,D150)
Digits(10)=Array(D151,D152,D153,D154,D155,D156,D157,D158,D159,D160,D161,D162,D163,D164,D165)
Digits(11)=Array(D166,D167,D168,D169,D170,D171,D172,D173,D174,D175,D176,D177,D178,D179,D180)
Digits(12)=Array(D181,D182,D183,D184,D185,D186,D187,D188,D189,D190,D191,D192,D193,D194,D195)
Digits(13)=Array(D196,D197,D198,D199,D200,D201,D202,D203,D204,D205,D206,D207,D208,D209,D210)
Digits(14)=Array(D211,D212,D213,D214,D215,D216,D217,D218)
Digits(15)=Array(D219,D220,D221,D222,D223,D224,D225,D226)
Digits(16)=Array(D227,D228,D229,D230,D231,D232,D233,D234)
Digits(17)=Array(D235,D236,D237,D238,D239,D240,D241,D242)
Digits(18)=Array(D243,D244,D245,D246,D247,D248,D249,D250)
Digits(19)=Array(D251,D252,D253,D254,D255,D256,D257,D258)
Digits(20)=Array(D259,D260,D261,D262,D263,D264,D265,D266)
Digits(21)=Array(D267,D268,D269,D270,D271,D272,D273,D274)
Digits(22)=Array(D275,D276,D277,D278,D279,D280,D281,D282)
Digits(23)=Array(D283,D284,D285,D286,D287,D288,D289,D290)
Digits(24)=Array(D291,D292,D293,D294,D295,D296,D297,D298)
Digits(25)=Array(D299,D300,D301,D302,D303,D304,D305,D306)
Digits(26)=Array(D307,D308,D309,D310,D311,D312,D313,D314)
Digits(27)=Array(D315,D316,D317,D318,D319,D320,D321,D322)
Digits(28)=Array(D323,D324,D325,D326,D327,D328,D329,D330)
Digits(29)=Array(D331,D332,D333,D334,D335,D336,D337,D338)
Digits(30)=Array(D339,D340,D341,D342,D343,D344,D345,D346)
Digits(31)=Array(D347,D348,D349,D350,D351,D352,D353,D354)
Digits(32)=Array(D355,D356,D357,D358,D359,D360,D361,D362)

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







' Walls

Sub FadeWS(nr, a, b, c, d)
    Select Case FadingLevel(nr)
        Case 2:a.IsDropped = 1:b.IsDropped = 1:c.IsDropped = 1:d.IsDropped = 0:FadingLevel(nr) = 0 'Off
        Case 3:a.IsDropped = 1:b.IsDropped = 0:c.IsDropped = 1:d.IsDropped = 1:FadingLevel(nr) = 2 'fading...
        Case 4:a.IsDropped = 1:b.IsDropped = 1:c.IsDropped = 0:d.IsDropped = 1:FadingLevel(nr) = 3 'fading...
        Case 5:a.IsDropped = 0:b.IsDropped = 1:c.IsDropped = 1:d.IsDropped = 1:FadingLevel(nr) = 1 'ON
    End Select
End Sub

'Lights, Ramps & Primitives used as 4 step fading lights
'a,b,c,d are the images used from on to off

Sub FadeObj(nr, object, a, b, c, d)
    Select Case FadingLevel(nr)
        Case 4:object.image = b:FadingLevel(nr) = 6                   'fading to off...
        Case 5:object.image = a:FadingLevel(nr) = 1                   'ON
        Case 6, 7, 8:FadingLevel(nr) = FadingLevel(nr) + 1            'wait
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
        Case 4:object.image = b:FadingLevel(nr) = 0
        Case 5:object.image = a:FadingLevel(nr) = 1
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

Sub Flashc(nr, object)
    Select Case FadingLevel(nr)
		Case 3
            FadingLevel(nr) = 0 'completely off
            Object.IntensityScale = FlashLevel(nr)
        Case 4 'off
            FlashLevel(nr) = FlashLevel(nr) - (FlashSpeedDown(nr))
            If FlashLevel(nr) < FlashMin(nr) Then
                FlashLevel(nr) = FlashMin(nr)
               FadingLevel(nr) = 3 'completely off
            End if
            Object.IntensityScale = FlashLevel(nr)
        Case 5 ' on
            FlashLevel(nr) = FlashLevel(nr) + (FlashSpeedUp(nr))
            If FlashLevel(nr) > FlashMax(nr) Then
                FlashLevel(nr) = FlashMax(nr)
                FadingLevel(nr) = 6 'completely on
            End if
            Object.IntensityScale = FlashLevel(nr)
        Case 6 ' on
			FadingLevel(nr) = 1 'completely on
			Object.IntensityScale = FlashLevel(nr)
    End Select
End Sub

Sub Flashm(nr, object) 'multiple flashers, it just sets the flashlevel
    Object.IntensityScale = FlashLevel(nr)
End Sub

' RGB Leds

Sub RGBLED (object,red,green,blue)
object.color = RGB(0,0,0)
object.colorfull = RGB(2.5*red,2.5*green,2.5*blue)
object.state=1
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

Sub FlashMod(nr, object)
	Object.IntensityScale = FadingLevel(nr)/255
End Sub

Sub LampMod(nr, object)
Object.IntensityScale = FadingLevel(nr)/255
Object.State = LampState(nr)
End Sub




Sub SolGi(enabled)
  If enabled Then
		LightCenter1.State = 1 'Red Center Light On
		LightCenter2.State = 1 'Red Center Light On
		LightCenter3.State = 1 'Red Center Light On
		LightCenter4.State = 1 'Red Center Light On
		LightCenter5.State = 1 'Red Center Light On
     GiOFF
	 GiFlasher1.Visible = 0
	 GiFlasher2.Visible = 0
	 GiFlasher3.Visible = 0
	 GiFlasher4.Visible = 0

	Playsound "fx_relay_off"
	Table1.ColorGradeImage = "ColorGrade_1"
   Else
		LightCenter1.State = 0 'Red Center Light Off
		LightCenter2.State = 0 'Red Center Light Off
		LightCenter3.State = 0 'Red Center Light Off
		LightCenter4.State = 0 'Red Center Light Off
		LightCenter5.State = 0 'Red Center Light Off
     GiON
	 GiFlasher1.Visible = 1
	 GiFlasher2.Visible = 1
	 GiFlasher3.Visible = 1
	 GiFlasher4.Visible = 1

	Playsound "fx_relay_on"
	Table1.ColorGradeImage = "ColorGrade_8"
		GITimer.enabled = 1
 End If
End Sub

Sub GITimer_Timer
		LightCenter1.State = 0 'Red Center Light Off
		LightCenter2.State = 0 'Red Center Light Off
		LightCenter3.State = 0 'Red Center Light Off
		LightCenter4.State = 0 'Red Center Light Off
		LightCenter5.State = 0 'Red Center Light Off
'	PFGI.State = 1
'	plasticsOn.WidthTop=1024	
'	plasticsOn.WidthBottom=1024
	GITimer.enabled = 0
End Sub

Sub GiON
	dim xx
    For each xx in GiLights
        xx.State = 1
    Next
End Sub

Sub GiOFF
	dim xx
    For each xx in GiLights
        xx.State = 0
    Next
End Sub



'***********************************************************************************
'****				        		Sound routines      				    	****
'***********************************************************************************
'****                      Supporting Ball & Sound Functions					****
'***********************************************************************************

Function RndNum(min, max)
    RndNum = Int(Rnd() * (max-min + 1) ) + min ' Sets a random number between min and max
End Function

Function LVL(input)
	LVL = Input * SoundLevelMult
End Function

Function Vol(ball) ' Calculates the Volume of the sound based on the ball speed
    Vol = Csng(BallVel(ball) ^2 / 2000)
End Function

Function Vol2(ball1, ball2) ' Calculates the Volume of the sound based on the speed of two balls
    Vol2 = (Vol(ball1) + Vol(ball2) ) / 2
End Function

Function Pan(ball) ' Calculates the pan for a ball based on the X position on the table. "table1" is the name of the table
    Dim tmp
    tmp = ball.x * 2 / Table1.width-1
    If tmp> 0 Then
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

'***********************************************************************************
'****				        	JP's VP10 Rolling Sounds      			    	****
'***********************************************************************************

Const tnob = 9 ' total number of balls
ReDim rolling(tnob)
InitRolling

Sub InitRolling
    Dim i
    For i = 0 to tnob
        rolling(i) = False
    Next
End Sub

Sub RollingTimer_Timer() : RollingUpdate() : End Sub

Sub RollingUpdate()
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
            PlaySound("fx_ballrolling" & b), -1, Vol(BOT(b) )*5, Pan(BOT(b) ), 0, Pitch(BOT(b) ), 1, 0
        Else
            If rolling(b) = True Then
                StopSound("fx_ballrolling" & b)
                rolling(b) = False
            End If
        End If
    Next
End Sub

'***********************************************************************************
'****				       		Ball Collision Sound	   				    	****
'***********************************************************************************

Sub OnBallBallCollision(ball1, ball2, velocity)
	PlaySound("fx_collide"), 0, LVL(Csng(velocity) ^2 / 2000), Pan(ball1), 0, Pitch(ball1), 0, 0
End Sub

Sub zApron_Hit (idx)
	PlaySound "woodhitaluminium", 0, LVL((Vol(ActiveBall)^2.5)*10), Pan(ActiveBall)/4, 0, Pitch(ActiveBall), 1, 0
End Sub

Sub zInlanes_Hit (idx)
	PlaySound "MetalHit2", 0, LVL(Vol(ActiveBall)*5), Pan(ActiveBall)*55, 0, Pitch(ActiveBall), 1, 0
End Sub

Sub RampEntry_Hit()
 	If activeball.vely < -10 then 
		PlaySound "ramp_hit2", 0, LVL(Vol(ActiveBall)/5), Pan(ActiveBall), 0, Pitch(ActiveBall)*10, 1, 0
	Elseif activeball.vely > 3 then 
		PlaySound "PlayfieldHit", 0, LVL(Vol(ActiveBall) ), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0
	End If
End Sub

Sub RampHit2_Hit()	
	PlaySound "ramp_hit3", 0, LVL(Vol(ActiveBall) ), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0
End Sub

Sub RampEntry2_Hit()
	Playsound "WireRamp", 0, LVL(0.3)
End Sub

Sub RampSound0_Hit()
	If Activeball.velx > 0 then
		Playsound "pmaReriW", 0, LVL(1)
	Else
		Playsound "WireRamp1", 0, LVL(1)
	End If
End Sub

Sub Metals_Hit (idx)
	PlaySound "metalhit_medium", 0, LVL(Vol(ActiveBall)*30 ), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0
End Sub

Sub Targets_Hit (idx)
	PlaySound SoundFX("target",DOFTargets), 0, LVL(0.2), Pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0
	PlaySound SoundFX("targethit",0), 0, LVL(Vol(ActiveBall)*18 ), Pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0
End Sub

Sub Gates_Hit (idx)
	PlaySound "gate", 0, LVL(Vol(ActiveBall) ), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0
End Sub

Sub Rubbers_Hit(idx)
 	dim finalspeed
  	finalspeed=SQR(activeball.velx * activeball.velx + activeball.vely * activeball.vely)
 	If finalspeed > 15 then 
		PlaySound "fx_rubber2", 0, LVL(Vol(ActiveBall)*30 ), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0
	End if

	If finalspeed >= 6 AND finalspeed <= 15 then
 		RandomSoundRubber()
	else
		PlaySound "rubber_hit_3", 0, LVL(Vol(ActiveBall)*30 ), Pan(ActiveBall)*50, 0, Pitch(ActiveBall), 1, 0
 	End If
End Sub

Sub RubberPosts_Hit(idx)
 	dim finalspeed
  	finalspeed=SQR(activeball.velx * activeball.velx + activeball.vely * activeball.vely)
 	If finalspeed > 16 then 
		PlaySound "fx_rubber2", 0, LVL(Vol(ActiveBall)*80 ), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0
	End if
	If finalspeed >= 6 AND finalspeed <= 15 then
 		RandomSoundRubber()
	else
		PlaySound "rubber_hit_3", 0, LVL(Vol(ActiveBall)*1 ), Pan(ActiveBall)*50, 0, Pitch(ActiveBall), 1, 0
 	End If
End Sub

Sub RandomSoundRubber()
	Select Case Int(Rnd*3)+1
		Case 1 : PlaySound "rubber_hit_1", 0, LVL(Vol(ActiveBall)*50 ), Pan(ActiveBall)*50, 0, Pitch(ActiveBall), 1, 0
		Case 2 : PlaySound "rubber_hit_2", 0, LVL(Vol(ActiveBall)*50 ), Pan(ActiveBall)*50, 0, Pitch(ActiveBall), 1, 0
		Case 3 : PlaySound "rubber_hit_3", 0, LVL(Vol(ActiveBall)*50 ), Pan(ActiveBall)*50, 0, Pitch(ActiveBall), 1, 0
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
		Case 1 : PlaySound "flip_hit_1", 0, LVL(Vol(ActiveBall) ), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0
		Case 2 : PlaySound "flip_hit_2", 0, LVL(Vol(ActiveBall) ), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0
		Case 3 : PlaySound "flip_hit_3", 0, LVL(Vol(ActiveBall) ), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0
	End Select
End Sub

'Ramp drops using collision events
Sub col_UPF_Fall_hit():if FallSFX1.Enabled = 0 then playsound "drop_mono", 0, LVL(0.5), 0 :FallSFX1.Enabled = 1	:end if :end sub'name,loopcount,volume,pan,randompitch
Sub FallSFX1_Timer():me.Enabled = 0:end sub

Sub Col_Rramp_Fall_Hit():if FallSFX2.Enabled = 0 then playsound "drop_mono", 0, LVL(0.5), 0.05 :FallSFX2.Enabled = 1	:end if :end sub'name,loopcount,volume,pan,randompitch
Sub FallSFX2_Timer():me.Enabled = 0:end sub

Sub col_Lramp_Fall_Hit():if FallSFX3.Enabled = 0 then playsound "drop_mono", 0, LVL(0.5), -0.05 :FallSFX3.Enabled = 1	:end if :end sub'name,loopcount,volume,pan,randompitch
Sub FallSFX3_Timer():me.Enabled = 0:end sub

'***********************************************************************************
'****				       		DOF reference	 	     				    	****
'***********************************************************************************
const dShooterLane	 				= 201	' Shooterlane
const dLeftSlingshot 				= 211	' Left Slingshot
const dRightSlingshot 				= 212	' Right Slingshot

Option Explicit

On Error Resume Next
ExecuteGlobal GetTextFile("controller.vbs")
If Err Then MsgBox "You need the controller.vbs in order to run this table, available in the vp10 package"
On Error Goto 0

' Thalamus 2018-07-23
' Added/Updated "Positional Sound Playback Functions" and "Supporting Ball & Sound Functions"
' No special SSF tweaks yet.
' , AudioFade(ActiveBall)

'E101 Left flipper2
'E102 Right flipper1

LoadVPM "01500000","S7.VBS",3.02

'Const Ballmass = 1.1
Const cGameName="Jst_L2",cCredits="",UseSolenoids=1,UseLamps=0,UseGI=1,UseSync=1
Const SSolenoidOn="solon",SSolenoidOff="soloff",SFlipperOn="FlipperUp",SFlipperOff="FlipperDown",SCoin=""
dim sGameOn

SolCallback(1)="bsTroughP1.SolIn"
SolCallback(2)="bsTroughP2.SolIn"
SolCallback(3)="SolP1Trough"
SolCallback(4)="SolP2Trough"
SolCallback(5)="bsP1.SolOut"
SolCallback(6)="bsP2.SolOut"
SolCallback(7)="bsEjectP1.SolOut"
SolCallback(8)="bsEjectP2.SolOut"
SolCallback(9)="dtA.SolDropUp"
SolCallback(10)="dtB.SolDropUp"
SolCallback(11)="dtC.SolDropUp"
SolCallback(12)="dtD.SolDropUp"
SolCallback(13)="dtR.SolDropUp"
SolCallback(14)="dtL.SolDropUp"
SolCallback(15)="UpDateGI"
Const sGI=15
Const sCLO=16
Const sFlipperRelay=17'does not work - scripted around below
SolCallback(19)="vpmSolSound SoundFX(""sling"",DOFContactors),"
SolCallback(20)="vpmSolSound SoundFX(""sling"",DOFContactors),"
SolCallback(21)="vpmSolSound SoundFX(""sling"",DOFContactors),"
SolCallback(22)="vpmSolSound SoundFX(""sling"",DOFContactors),"


Sub SolP1Trough(Enabled)
	If Enabled Then
		If bsTroughP1.Balls Then bsTroughP1.ExitSol_On
	End If
End Sub

Sub SolP2Trough(Enabled)
	If Enabled Then
		If bsTroughP2.Balls Then bsTroughP2.ExitSol_On
	End If
End Sub
'********GI
Sub UpDateGI(enabled)
	Dim xx
	If Enabled Then
		For Each xx in GI:xx.State = 1:Next
	Else
		For Each xx in GI:xx.State = 0:Next
	End If
End Sub

Sub LoadVPM(VPMver, VBSfile, VBSver)
	On Error Resume Next
		If ScriptEngineMajorVersion < 5 Then MsgBox "VB Script Engine 5.0 or higher required"
		ExecuteGlobal GetTextFile(VBSfile)
		If Err Then MsgBox "Unable to open " & VBSfile & ". Ensure that it is in the same folder as this table. " & vbNewLine & Err.Description : Err.Clear
		Set Controller = CreateObject("VPinMAME.Controller")
		If Err Then MsgBox "Can't Load VPinMAME." & vbNewLine & Err.Description
		If VPMver>"" Then If Controller.Version < VPMver Or Err Then MsgBox "VPinMAME ver " & VPMver & " required." : Err.Clear
		If VPinMAMEDriverVer < VBSver Or Err Then MsgBox VBSFile & " ver " & VBSver & " or higher required."
	On Error Goto 0
End Sub

Dim Flips

Set LampCallback=GetRef("UpdateMultipleLamps")
Dim bsTroughP1,bsTroughP2,bsEjectP1,bsEjectP2,dtA,dtB,dtC,dtD,dtL,dtR,bsP1,bsP2

Sub Table1_Init
	On Error Resume Next
	With Controller
		.GameName=cGameName
		If Err Then MsgBox "Can't start Game" & cGameName & vbNewLine & Err.Description : Exit Sub
		.SplashInfoLine=cCredits
		.ShowDMDOnly=1
		.ShowFrame=1
		.ShowTitle=0
		.Hidden=1
		.Run
		If Err Then MsgBox Err.Description
	End With
	On Error Goto 0
	PinMAMETimer.Interval=PinMAMEInterval
	PinMAMETimer.Enabled=1
	vpmNudge.TiltSwitch=1
	vpmNudge.Sensitivity=5
	vpmNudge.TiltObj=Array(LeftSlingshot,RightSlingshot)

	Set bsTroughP1=New cvpmBallStack
    bsTroughP1.InitSw 9,10,11,12,13,0,0,0
    bsTroughP1.InitKick BallRelease,60,5
    bsTroughP1.InitExitSnd SoundFX("BallRel",DOFContactors),SoundFX("Solon",DOFContactors)
    bsTroughP1.Balls=2

	Set bsTroughP2=New cvpmBallStack
    bsTroughP2.InitSw 33,34,35,36,37,0,0,0
    bsTroughP2.InitKick Kicker3,120,120
    bsTroughP2.InitExitSnd SoundFX("Popper",DOFContactors),SoundFX("Solon",DOFContactors)
    bsTroughP2.Balls=2

	Set bsEjectP1=New cvpmBallStack
	bsEjectP1.InitSaucer Switch15,15,220,8
	bsEjectP1.InitExitSnd SoundFX("Popper",DOFContactors),SoundFX("Solon",DOFContactors)

	Set bsEjectP2=New cvpmBallStack
    bsEjectP2.InitSaucer Switch39,39,20,8
    bsEjectP2.InitExitSnd SoundFX("Popper",DOFContactors),SoundFX("Solon",DOFContactors)

    Set dtA=New cvpmDropTarget
    dtA.InitDrop Sw17,17
    dtA.InitSnd SoundFX("FlapClos",DOFDropTargets),SoundFX("FlapOpen",DOFDropTargets)

	Set dtB=New cvpmDropTarget
    dtB.InitDrop Sw41,41
    dtB.InitSnd SoundFX("FlapClos",DOFDropTargets),SoundFX("FlapOpen",DOFDropTargets)

    Set dtC=New cvpmDropTarget
    dtC.InitDrop Sw18,18
    dtC.InitSnd SoundFX("FlapClos",DOFDropTargets),SoundFX("FlapOpen",DOFDropTargets)

    Set dtD=New cvpmDropTarget
    dtD.InitDrop Sw42,42
    dtD.InitSnd SoundFX("FlapClos",DOFDropTargets),SoundFX("FlapOpen",DOFDropTargets)

    Set dtL=New cvpmDropTarget
    dtL.InitDrop Array(Sw26,Sw27,Sw28,Sw50,Sw51,Sw52),Array(26,27,28,50,51,52)
    dtL.InitSnd SoundFX("FlapClos",DOFDropTargets),SoundFX("FlapOpen",DOFDropTargets)

    Set dtR=New cvpmDropTarget
    dtR.InitDrop Array(Sw23,Sw24,Sw25,Sw47,Sw48,Sw49),Array(23,24,25,47,48,49)
    dtR.InitSnd SoundFX("FlapClos",DOFDropTargets),SoundFX("FlapOpen",DOFDropTargets)

	Set bsP1=New cvpmBallStack
    bsP1.InitSaucer P1K,14,0,25
    bsP1.InitExitSnd "Plunger","Solon"

	Set bsP2=New cvpmBallStack
    bsP2.InitSaucer P2K,38,180,25
    bsP2.InitExitSnd "Plunger","Solon"

'vpmMapLights AllLights
'Lamps 8 and 38 are backwards in manual
	operatormenu=0
	Flips=0
	LoadOptions


UpdateGI 1
p2wall1.isdropped = True
p2wall.isdropped = True
Flipper3.enabled = false
Flipper4.enabled = false
hard
End Sub


Dim sw23up, sw24up, sw25up, sw26up, sw27up, sw28up, sw41up, sw42up, sw47up, sw48up, sw49up, sw50up, sw17up, sw18up, sw51up, sw52up
Dim PrimT


Sub PrimT_Timer
    if sw23.IsDropped = True then sw23up = False:Ramp108.Image = "off" else sw23up = True:Ramp108.Image = "pf2"
    if sw24.IsDropped = True then sw24up = False:Ramp107.Image = "off":Ramp110.Image = "off" else sw24up = True:Ramp107.Image = "pf1":Ramp110.Image = "pf1"
    if sw25.IsDropped = True then sw25up = False:Ramp106.Image = "off" else sw25up = True:Ramp106.Image = "pf2"
    if sw26.IsDropped = True then sw26up = False:Ramp100.Image = "off" else sw26up = True:Ramp100.Image = "pf2"
    if sw27.IsDropped = True then sw27up = False:Ramp102.Image = "off":Ramp101.Image = "off" else sw27up = True:Ramp102.Image = "pf1":Ramp101.Image = "pf1"
    if sw28.IsDropped = True then sw28up = False:Ramp103.Image = "off" else sw28up = True:Ramp103.Image = "pf2"
    if sw41.IsDropped = True then sw41up = False else sw41up = True
    if sw42.IsDropped = True then sw42up = False else sw42up = True
    if sw47.IsDropped = True then sw47up = False:Ramp103p2.Image = "off" else sw47up = True:Ramp103p2.Image = "pf2"
    if sw48.IsDropped = True then sw48up = False:Ramp102p2.Image = "off":Ramp101p2.Image = "off" else sw48up = True:Ramp102p2.Image = "pf1":Ramp101p2.Image = "pf1"
    if sw49.IsDropped = True then sw49up = False:Ramp100p2.Image = "off" else sw49up = True:Ramp100p2.Image = "pf2"
    if sw50.IsDropped = True then sw50up = False:Ramp106p2.Image = "off" else sw50up = True:Ramp106p2.Image = "pf2"
    if sw17.IsDropped = True then sw17up = False else sw17up = True
    if sw18.IsDropped = True then sw18up = False else sw18up = True
    if sw51.IsDropped = True then sw51up = False:Ramp107p2.Image = "off":Ramp110p2.Image = "off" else sw51up = True:Ramp107p2.Image = "pf1":Ramp110p2.Image = "pf1"
    if sw52.IsDropped = True then sw52up = False :Ramp108p2.Image = "off" else sw52up = True:Ramp108p2.Image = "pf2"
    'End Sub


if GIB29.State = 0 then
		Ramp108.Visible = false
		Ramp108p2.Visible = false
		Ramp107.Visible = false
		Ramp107p2.Visible = false
		Ramp110.Visible = false
		Ramp110p2.Visible = false
		Ramp106.Visible = false
		Ramp106p2.Visible = false
		Ramp103.Visible = false
		Ramp103p2.Visible = false
		Ramp102.Visible = false
		Ramp102p2.Visible = false
		Ramp101.Visible = false
		Ramp101p2.Visible = false
		Ramp100.Visible = false
		Ramp100p2.Visible = false

	Else
		Ramp108.Visible = true
		Ramp108p2.Visible = true
		Ramp107.Visible = true
		Ramp107p2.Visible = true
		Ramp110.Visible = true
		Ramp110p2.Visible = true
		Ramp106.Visible = true
		Ramp106p2.Visible = true
		Ramp103.Visible = true
		Ramp103p2.Visible = true
		Ramp102.Visible = true
		Ramp102p2.Visible = true
		Ramp101.Visible = true
		Ramp101p2.Visible = true
		Ramp100.Visible = true
		Ramp100p2.Visible = true

	End If

End Sub


Sub DT_Timer()
    If sw23up = True and sw23p.z < -.5 then sw23p.z = sw23p.z + 3
    If sw24up = True and sw24p.z < -.1 then sw24p.z = sw24p.z + 3
    If sw25up = True and sw25p.z < -1 then sw25p.z = sw25p.z + 3
    If sw26up = True and sw26p.z < -1 then sw26p.z = sw26p.z + 3
    If sw27up = True and sw27p.z < -.1 then sw27p.z = sw27p.z + 3
    If sw28up = True and sw28p.z < -1 then sw28p.z = sw28p.z + 3
    If sw41up = True and sw41p.z < 1 then sw41p.z = sw41p.z + 3
    If sw42up = True and sw42p.z < -.5 then sw42p.z = sw42p.z + 3
    If sw47up = True and sw47p.z < -.5 then sw47p.z = sw47p.z + 3
    If sw48up = True and sw48p.z < .1 then sw48p.z = sw48p.z + 3
    If sw49up = True and sw49p.z < 1 then sw49p.z = sw49p.z + 3
    If sw50up = True and sw50p.z < 1 then sw50p.z = sw50p.z + 3
    If sw17up = True and sw17p.z < -1 then sw17p.z = sw17p.z + 3
    If sw18up = True and sw18p.z < -1 then sw18p.z = sw18p.z + 3
    If sw51up = True and sw51p.z < .1 then sw51p.z = sw51p.z + 3
    If sw52up = True and sw52p.z < -.5 then sw52p.z = sw52p.z + 3
    If sw23up = False and sw23p.z > -41 then sw23p.z = sw23p.z - 3
    If sw24up = False and sw24p.z > -40.5 then sw24p.z = sw24p.z - 3
    If sw25up = False and sw25p.z > -40 then sw25p.z = sw25p.z - 3
    If sw26up = False and sw26p.z > -40 then sw26p.z = sw26p.z - 3
    If sw27up = False and sw27p.z > -40.5 then sw27p.z = sw27p.z - 3
    If sw28up = False and sw28p.z > -42 then sw28p.z = sw28p.z - 3
    If sw41up = False and sw41p.z > -40 then sw41p.z = sw41p.z - 3
    If sw42up = False and sw42p.z > -40 then sw42p.z = sw42p.z - 3
    If sw47up = False and sw47p.z > -20 then sw47p.z = sw47p.z - 3
    If sw48up = False and sw48p.z > -20 then sw48p.z = sw48p.z - 3
    If sw49up = False and sw49p.z > -20 then sw49p.z = sw49p.z - 3
    If sw50up = False and sw50p.z > -20 then sw50p.z = sw50p.z - 3
    If sw17up = False and sw17p.z > -40 then sw17p.z = sw17p.z - 3
    If sw18up = False and sw18p.z > -40 then sw18p.z = sw18p.z - 3
    If sw51up = False and sw51p.z > -20 then sw51p.z = sw51p.z - 3
    If sw52up = False and sw52p.z > -20 then sw52p.z = sw52p.z - 3
    If sw23p.z >= 20 then sw23up = False
    If sw24p.z >= 20 then sw24up = False
    If sw25p.z >= 20 then sw25up = False
    If sw26p.z >= -20 then sw26up = False
    If sw27p.z >= -20 then sw27up = False
    If sw28p.z >= -20 then sw28up = False
    If sw41p.z >= -20 then sw41up = False
    If sw42p.z >= -20 then sw42up = False
    If sw47p.z >= -20 then sw47up = False
    If sw48p.z >= -20 then sw48up = False
    If sw49p.z >= -20 then sw49up = False
    If sw50p.z >= -20 then sw50up = False
    If sw17p.z >= -20 then sw17up = False
    If sw18p.z >= -20 then sw18up = False
    If sw51p.z >= -20 then sw51up = False
    If sw52p.z >= -20 then sw52up = False
End Sub


Dim operatormenu, options

Sub Table1_KeyDown(ByVal KeyCode)
	if operatormenu=0 then
		If KeyCode=StartGameKey Then
			If flips > 0 then
				Controller.Switch(57)=1
			  else
				Controller.Switch(3)=1
			end if
			exit sub
		End If
		If KeyCode=LeftFlipperKey Then
			FlipL
			if RightSlingShot.SlingshotThreshold=100 then OperatorMenuTimer.Enabled = true
		end if
		If KeyCode=RightFlipperKey Then FlipR
		If KeyCode=RIGHTMagnaSave Then Flip2L
		If KeyCode=LEFTMagnaSave Then Flip2R
		If vpmKeyDown(KeyCode) Then Exit Sub
	'	                     if keycode=leftflipperkey Then:ScrewHead2.z = ScrewHead2.z -1: debug.print ScrewHead2.z
	else
		If keycode=LeftFlipperKey then
			Options=Options+1
			If Options=3 then Options=1
			playsound "target"
			Select Case (Options)
				Case 1:
					Option1.visible=true
					Option2.visible=False
				Case 2:
					Option2.visible=true
					Option1.visible=False
			End Select
		end if
		If keycode=RightFlipperKey then
			  PlaySound "metalhit2"
			  Select Case (Options)
				Case 1:
					Flips=Flips+1
					If Flips>3 then Flips=0
					OptionPlayers.image = "OptionPlayers"&Flips
				Case 2:
					OperatorMenu=0
					SaveOptions
					HideOptions
					Hard
			  End Select
		end if
	end if

End Sub

Sub OperatorMenuTimer_Timer
	OperatorMenu=1
	Displayoptions
	Options=1
End Sub

Sub DisplayOptions
	OptionsBack.visible = true
	Option1.visible = True
	OptionPlayers.image = "OptionPlayers"&Flips
	OptionPlayers.visible = True
End Sub

Sub HideOptions
	dim objekt
	for each objekt In OptionMenu
		objekt.visible = false
	next
End Sub

Sub SaveOptions
    savevalue "Joust", "Flips", Flips
end Sub

Sub LoadOptions
	dim temp
	temp = LoadValue("Joust", "Flips")
    If (temp <> "") then Flips = CDbl(temp)
end sub

Sub FlipL
	If Tilted=0 Then
		If Flips>0 Then
			LeftFlipper.RotateToEnd
			PlaySound SoundFX("FlipperUp",DOFFlippers)
			Controller.Switch(58)=1
		Else
			LeftFlipper.RotateToEnd
			Flipper1.RotateToEnd
			PlaySound SoundFX("FlipperUp",DOFFlippers)
			Controller.Switch(58)=1
			Controller.Switch(60)=1
		End If
	End If
End Sub

Sub FlipR
	If Tilted=0 Then
		If Flips>0 Then
			RightFlipper.RotateToEnd
			PlaySound SoundFX("FlipperUp",DOFFlippers)
			Controller.Switch(59)=1
		Else
			RightFlipper.RotateToEnd
			Flipper2.RotateToEnd
			PlaySound SoundFX("FlipperUp",DOFFlippers)
			Controller.Switch(59)=1
			Controller.Switch(61)=1
		End If
	End If
End Sub

Sub Flip2L
	If Tilted=0 Then
		Flipper1.RotateToEnd
		DOF 102,1
		Flipper3.RotateToEnd
		PlaySound SoundFX("FlipperUp",DOFFlippers)
		Controller.Switch(60)=1
	End If
End Sub

Sub Flip2R
	If Tilted=0 Then
		Flipper2.RotateToEnd
		DOF 101,1
		Flipper4.RotateToEnd
		PlaySound SoundFX("FlipperUp",DOFFlippers)
		Controller.Switch(61)=1
	End If
End Sub

Sub FlipLA
	If Tilted=0 Then
		If Flips>0 Then
			LeftFlipper.RotateToStart
			PlaySound SoundFX("FlipperDown",DOFFlippers)
			Controller.Switch(58)=0
		Else
			LeftFlipper.RotateToStart
			Flipper1.RotateToStart
			PlaySound SoundFX("FlipperDown",DOFFlippers)
			DOF 102,0
			Controller.Switch(58)=0
			Controller.Switch(60)=0
		End If
	End If
End Sub

Sub FlipRA
	If Tilted=0 Then
		If Flips>0 Then
			RightFlipper.RotateToStart
			PlaySound SoundFX("FlipperDown",DOFFlippers)
			Controller.Switch(59)=0
		Else
			RightFlipper.RotateToStart
			Flipper2.RotateToStart
			PlaySound SoundFX("FlipperDown",DOFFlippers)
			DOF 101,0
			Controller.Switch(59)=0
			Controller.Switch(61)=0
		End If
	End If
End Sub

Sub Flip2LA
	If Tilted=0 Then
		Flipper1.RotateToStart
		Flipper3.RotateToStart
		DOF 102,0
		PlaySound SoundFX("FlipperDown",DOFFlippers)
		Controller.Switch(60)=0
	End If
End Sub

Sub Flip2RA
	If Tilted=0 Then
		Flipper2.RotateToStart
		Flipper4.RotateToStart
		DOF 101,0
		PlaySound SoundFX("FlipperDown",DOFFlippers)
		Controller.Switch(61)=0
	End If
End Sub

Sub Table1_KeyUp(ByVal KeyCode)
    If KeyCode=startgamekey Then
    	Controller.Switch(57)=0
		Controller.Switch(3)=0
    	Exit Sub
    End If
	If KeyCode=LeftFlipperKey Then FlipLA:	OperatorMenuTimer.Enabled = false
    If KeyCode=RightFlipperKey Then FlipRA
	If KeyCode=RIGHTMagnaSave Then Flip2LA
	If KeyCode=LEFTMagnaSave Then Flip2RA
    If vpmKeyUp(KeyCode) Then Exit Sub
End Sub
'-----------------------------------
' Flipper Primitives
'-----------------------------------
sub FlipperTimer_Timer()
	pleftFlipper.rotz=leftFlipper.CurrentAngle
	prightFlipper.rotz=rightFlipper.CurrentAngle
end sub


'***********************************************
'**********P2 auto flipper control Easy*********
'***********************************************
Sub p2right_Hit
	if flips>1 then flip2r

End Sub

Sub p2left_Hit
	if flips>1 then Flip2L

End Sub

Sub p2center_Hit
	if flips>1 then
		flip2r
		flip2L

	end If
End Sub

Sub p2right_unHit
	if flips>1 then Flip2RA

	End Sub

Sub p2left_unHit
	if flips>1 then Flip2LA
End Sub

Sub p2center_unHit
	if flips>1 then
		Flip2RA
		Flip2LA
		end if
End Sub


'**********************************************************
'************PP2 auto flipper control Hard*****************
'******No logic anyone want to make this smart?*******
'**********************************************************

Sub Hard
	if flips>2 Then
		p2wall1.IsDropped = 1
		p2wall.isdropped = 0
		Flipper3.enabled = true
		Flipper4.enabled = true
		Flipper1.enabled = false
		Flipper2.enabled = false
Else
		p2wall1.IsDropped = 1
		p2wall.isdropped = 1
		Flipper3.enabled = false
		Flipper4.enabled = false
		Flipper1.enabled = true
		Flipper2.enabled = true
end if
End Sub



Sub Outholep1_Hit:PlaySound "drain":bsTroughP1.AddBall Me:End Sub				'9-13
'Sub Switch14_Hit:vpmTimer.PulseSw 14:End Sub 'not working
'Sub switch38_Hit:vpmTimer.PulseSw 38:End Sub 'not working
Sub P1K_Hit:bsP1.AddBall 0:End Sub							'14
Sub Switch15_Hit:bsEjectP1.AddBall Me:End Sub
'Sub Switch15_unHit:Primitive_BallEjectArm1.z=30:End Sub
Sub Switch16_Spin:vpmTimer.PulseSw 16:End Sub				'16
Sub Sw17_Hit:dtA.Hit 1:End Sub							'17
Sub Sw18_Hit:dtC.Hit 1:End Sub							'18
Sub Switch19_Hit:Controller.Switch(19)=1:End Sub			'19
Sub Switch19_unHit:Controller.Switch(19)=0:End Sub
Sub Switch20_Hit:Controller.Switch(20)=1:End Sub			'20
Sub Switch20_unHit:Controller.Switch(20)=0:End Sub
Sub Switch21_Hit:vpmTimer.PulseSw 21:End Sub				'21
Sub Switch22_Hit:vpmTimer.PulseSw 22:End Sub				'22
Sub Sw23_Hit:dtR.Hit 1:End Sub							'23
Sub Sw24_Hit:dtR.Hit 2:End Sub							'24
Sub Sw25_Hit:dtR.Hit 3:End Sub							'25
Sub Sw26_Hit:dtL.Hit 1:End Sub							'26
Sub Sw27_Hit:dtL.Hit 2:End Sub							'27
Sub Sw28_Hit:dtL.Hit 3:End Sub							'28
Sub Switch31_Hit:Controller.Switch(31)=1:End Sub			'31
Sub Switch31_unHit:Controller.Switch(31)=0:End Sub
Sub Outholep2_Hit:PlaySound "drain":bsTroughP2.AddBall Me:End Sub				'33-37
Sub P2K_Hit:bsP2.AddBall 0:End Sub							'38
Sub Switch39_Hit:Primitive_BallEjectArm2.visible = False:bsEjectP2.AddBall Me:End Sub
Sub Switch39_unHit:Primitive_BallEjectArm2.visible = True:End Sub		'39
Sub Switch40_Spin:vpmTimer.PulseSw 40:End Sub				'40
Sub Sw41_Hit:dtB.Hit 1:End Sub							'41
Sub Sw42_Hit:dtD.Hit 1:End Sub							'42
Sub Switch43_Hit:Controller.Switch(43)=1:End Sub			'43
Sub Switch43_unHit:Controller.Switch(43)=0:End Sub
Sub Switch44_Hit:Controller.Switch(44)=1:End Sub			'44
Sub Switch44_unHit:Controller.Switch(44)=0:End Sub
Sub Switch45_Hit:vpmTimer.PulseSw 45:End Sub				'45
Sub Switch46_Hit:vpmTimer.PulseSw 46:End Sub				'46
Sub Sw47_Hit:dtR.Hit 4:End Sub							'47
Sub Sw48_Hit:dtR.Hit 5:End Sub							'48
Sub Sw49_Hit:dtR.Hit 6:End Sub							'49
Sub Sw50_Hit:dtL.Hit 4:End Sub							'50
Sub Sw51_Hit:dtL.Hit 5:End Sub							'51
Sub Swi52_Hit:dtL.Hit 6:End Sub							'52
Sub Switch55_Hit:Controller.Switch(55)=1:End Sub			'55
Sub Switch55_unHit:Controller.Switch(55)=0:End Sub
Sub Trigger1_hit:Playsound "gate" End Sub
Sub Trigger2_hit:Playsound "gate" End Sub
Sub Trigger3_hit:Playsound "gate" End Sub
Sub Trigger4_hit:Playsound "gate" End Sub



'******GATES******

Sub timergate_Timer
	primitive54.rotz=gate3.CurrentAngle
End Sub

Sub Gate3_hit
PlaySound "Gate"
End Sub


Sub timergate1_Timer
	primitive56.rotz=gate4.CurrentAngle
End Sub

Sub Gate4_hit
PlaySound "Gate"
End Sub


Sub timergate2_Timer
	primitive58.rotz=gate5.CurrentAngle
End Sub

Sub Gate5_hit
PlaySound "Gate"
End Sub


Sub timergate3_Timer
	primitive61.rotz=gate6.CurrentAngle
End Sub

Sub Gate6_hit
PlaySound "Gate"
End Sub


'**********Sling Shot Animations
' Rstep and Lstep  are the variables that increment the animation
'****************
Dim RStep, Lstep

Sub RightSlingShot_Slingshot
    PlaySound SoundFX("RightSlingShot",DOFContactors), 0, 1, 0.05, 0.05
	vpmTimer.PulseSw 30
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

Sub RightSlingShot1_Slingshot
    PlaySound SoundFX("RightSlingShot",DOFContactors), 0, 1, 0.05, 0.05
	vpmTimer.PulseSw 54
    RSling3.Visible = 0
    RSling4.Visible = 1
    sling4.TransZ = -20
    RStep = 0
    RightSlingShot1.TimerEnabled = 1

End Sub

Sub RightSlingShot1_Timer
    Select Case RStep
        Case 3:RSLing4.Visible = 0:RSLing5.Visible = 1:sling4.TransZ = -10
        Case 4:RSLing5.Visible = 0:RSLing3.Visible = 1:sling4.TransZ = 0:RightSlingShot1.TimerEnabled = 0
    End Select
    RStep = RStep + 1
End Sub

Sub LeftSlingShot_Slingshot
	PlaySound SoundFX("LeftSlingShot",DOFContactors),0,1,-0.05,0.05
    vpmTimer.PulseSw 29
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

Sub LeftSlingShot1_Slingshot
	PlaySound SoundFX("LeftSlingShot",DOFContactors),0,1,-0.05,0.05
    vpmTimer.PulseSw 53
    LSling5.Visible = 0
    LSling3.Visible = 1
    sling3.TransZ = -20
    LStep = 0
    LeftSlingShot1.TimerEnabled = 1
	End Sub

Sub LeftSlingShot1_Timer
    Select Case LStep
        Case 3:LSLing3.Visible = 0:LSLing4.Visible = 1:sling3.TransZ = -10
        Case 4:LSLing4.Visible = 0:LSLing5.Visible = 1:sling3.TransZ = 0:LeftSlingShot1.TimerEnabled = 0
    End Select
    LStep = LStep + 1
End Sub


Dim Tilted
Tilted=0

Sub UpdateMultipleLamps
If Controller.Lamp(1) Then
	Timer1.Enabled=0
	Timer1.Enabled=1
End If
If Not Timer1.Enabled And Not Controller.Lamp(3) Then
	Tilted=0
Else
	Tilted=1
	Flipper1.RotateToStart
	Flipper2.RotateToStart
	LeftFlipper.RotateToStart
	RightFlipper.RotateToStart
	Controller.Switch(58)=0
	Controller.Switch(59)=0
	Controller.Switch(60)=0
	Controller.Switch(61)=0
	vpmNudge.SolGameOn False
	UpdateGI 1
End If

End Sub

Sub Timer1_Timer:Timer1.Enabled=0:Tilted=0:vpmNudge.SolGameOn True:UpdateGI 1:End Sub

'VpinMame Broken Display
Dim Digits(31)
Digits(8)=Array(L131,L132,L133,L134,L135,L136,L137)
Digits(9)=Array(L138,L139,L140,L141,L142,L143,L144)
Digits(12)=Array(L145,L146,L147,L148,L149,L150,L151)
Digits(13)=Array(L152,L153,L154,L155,L156,L157,L158)
Digits(14)=Array(L1,L2,L3,L4,L5,L6,L7,L8)
Digits(15)=Array(L9,L10,L11,L12,L13,L14,L15)
Digits(16)=Array(L16,L17,L18,L19,L20,L21,L22)
Digits(17)=Array(L23,L24,L25,L26,L27,L28,L29,L30)
Digits(18)=Array(L31,L32,L33,L34,L35,L36,L37)
Digits(19)=Array(L38,L39,L40,L41,L42,L43,L44)
Digits(20)=Array(L45,L46,L47,L48,L49,L50,L51)
Digits(21)=Array(L52,L53,L54,L55,L56,L57,L58,L59)
Digits(22)=Array(L60,L61,L62,L63,L64,L65,L66)
Digits(23)=Array(L67,L68,L69,L70,L71,L72,L73)
Digits(24)=Array(L74,L75,L76,L77,L78,L79,L80,L81)
Digits(25)=Array(L82,L83,L84,L85,L86,L87,L88)
Digits(26)=Array(L89,L90,L91,L92,L93,L94,L95)
Digits(27)=Array(L96,L97,L98,L99,L100,L101,L102)
Digits(28)=Array(L117,L118,L119,L120,L121,L122,L123)
Digits(29)=Array(L124,L125,L126,L127,L128,L129,L130)
Digits(30)=Array(L103,L104,L105,L106,L107,L108,L109)
Digits(31)=Array(L110,L111,L112,L113,L114,L115,L116)


Sub DisplayTimer_Timer
	Dim ChgLED,ii,num,chg,stat,obj
	ChgLED = Controller.ChangedLEDs(&HFFFFFFFF, &HFFFFFFFF) 'hex of binary (display 111111, or first 6 digits)
	If Not IsEmpty(ChgLED) Then
		For ii = 0 To UBound(chgLED)
			num = chgLED(ii, 0) : chg = chgLED(ii, 1) : stat = chgLED(ii, 2)
		If Not IsEmpty(Digits(Num)) Then
			For Each obj In Digits(num)
				If chg And 1 Then obj.State = stat And 1
				chg = chg\2 : stat = stat\2
			Next
		End If
		Next
	End If
End Sub


'***********************************************
'***********************************************
                    ' Lamps
'***********************************************
'***********************************************



 Dim LampState(200), FadingLevel(200), FadingState(200)
Dim FlashState(200), FlashLevel(200)
Dim FlashSpeedUp, FlashSpeedDown
Dim x

AllLampsOff()
LampTimer.Interval = 40 'lamp fading speed
LampTimer.Enabled = 1
'
FlashInit()
FlasherTimer.Interval = 10 'flash fading speed
FlasherTimer.Enabled = 1

Sub LampTimer_Timer()
    Dim chgLamp, num, chg, ii
    chgLamp = Controller.ChangedLamps
    If Not IsEmpty(chgLamp) Then
        For ii = 0 To UBound(chgLamp)
            LampState(chgLamp(ii, 0) ) = chgLamp(ii, 1)
            FadingLevel(chgLamp(ii, 0) ) = chgLamp(ii, 1) + 4
            FlashState(chgLamp(ii, 0) ) = chgLamp(ii, 1)

            'GI_CheckWalkerLights chgLamp(ii, 0), chgLamp(ii, 1)
        Next
    End If

    UpdateLamps
End Sub

Sub FlashInit
    Dim i
    For i = 0 to 200
        FlashState(i) = 0
        FlashLevel(i) = 0
    Next

    FlashSpeedUp = 500   ' fast speed when turning on the flasher
    FlashSpeedDown = 250 ' slow speed when turning off the flasher, gives a smooth fading
    AllFlashOff()
End Sub

Sub AllFlashOff
    Dim i
    For i = 0 to 200
        FlashState(i) = 0
    Next
End Sub

 Sub UpdateLamps()

	'NFadeLm 1, Light1a
	'NFadeL 1, Light1
	'NFadeLm 2, Light2a
	'NFadeL 2, Light2
	'NFadeLm 3, Light3a
	'NFadeL 3, Light3
	'NFadeLm 4, Light4a
	'NFadeL 4, Light4
    NFadePm 5, p5, "JoustPlayfield-ON", "JoustPlayfield-OFF"
	Flash 5, f5
    NFadePm 6, p6, "JoustPlayfield-ON", "JoustPlayfield-OFF"
	Flash 6, f6
    NFadePm 7, p7, "JoustPlayfield-ON", "JoustPlayfield-OFF"
	Flash 7, F7
    NFadePm 8, p8, "JoustPlayfield-ON", "JoustPlayfield-OFF"
	Flash 8, F8
    NFadePm 9, p9, "JoustPlayfield-ON", "JoustPlayfield-OFF"
	Flash 9, F9
    NFadePm 10, p10, "JoustPlayfield-ON", "JoustPlayfield-OFF"
	Flash 10, F10
    NFadePm 11, p11, "JoustPlayfield-ON", "JoustPlayfield-OFF"
	Flash 11, F11
    NFadePm 12, p12, "JoustPlayfield-ON", "JoustPlayfield-OFF"
	Flash 12, F12
    NFadePm 13, p13, "JoustPlayfield-ON", "JoustPlayfield-OFF"
	Flash 13, F13
    NFadePm 14, p14, "JoustPlayfield-ON", "JoustPlayfield-OFF"
	Flash 14, F14
    NFadePm 15, p15, "JoustPlayfield-ON", "JoustPlayfield-OFF"
	Flash 15, F15
    NFadePm 16, p16, "JoustPlayfield-ON", "JoustPlayfield-OFF"
	Flash 16, F16
    NFadePm 17, p17, "JoustPlayfield-ON", "JoustPlayfield-OFF"
	Flash 17, F17
    NFadePm 18, p18, "JoustPlayfield-ON", "JoustPlayfield-OFF"
	Flash 18, F18
    NFadePm 19, p19, "JoustPlayfield-ON", "JoustPlayfield-OFF"
	Flash 19, F19
    NFadePm 20, p20, "JoustPlayfield-ON", "JoustPlayfield-OFF"
	Flash 20, F20
    NFadePm 21, p21, "JoustPlayfield-ON", "JoustPlayfield-OFF"
	Flash 21, F21
    NFadePm 22, p22, "JoustPlayfield-ON", "JoustPlayfield-OFF"
	Flash 22, F22
    NFadePm 23, p23, "JoustPlayfield-ON", "JoustPlayfield-OFF"
	Flash 23, F23
    NFadePm 24, p24, "JoustPlayfield-ON", "JoustPlayfield-OFF"
	Flash 24, F24
    NFadePm 25, p25, "JoustPlayfield-ON", "JoustPlayfield-OFF"
	Flash 25, F25
    NFadePm 26, p26, "JoustPlayfield-ON", "JoustPlayfield-OFF"
	Flash 26, F26
    NFadePm 27, p27, "JoustPlayfield-ON", "JoustPlayfield-OFF"
	Flash 27, F27
    NFadePm 28, p28, "JoustPlayfield-ON", "JoustPlayfield-OFF"
	Flash 28, F28
    NFadePm 29, p29, "JoustPlayfield-ON", "JoustPlayfield-OFF"
	Flash 29, F29
    NFadePm 30, p30, "JoustPlayfield-ON", "JoustPlayfield-OFF"
	Flash 30, F30
    NFadePm 31, p31, "JoustPlayfield-ON", "JoustPlayfield-OFF"
	Flash 31, F31
    NFadePm 32, p32, "JoustPlayfield-ON", "JoustPlayfield-OFF"
	Flash 32, F32
    NFadePm 33, p33, "JoustPlayfield-ON", "JoustPlayfield-OFF"
	Flash 33, F33
    NFadePm 34, p34, "JoustPlayfield-ON", "JoustPlayfield-OFF"
	Flash 34, F34
	NFadePm 35, p35, "JoustPlayfield-ON", "JoustPlayfield-OFF"
    Flash 35, F35
	NFadePm 36, p36, "JoustPlayfield-ON", "JoustPlayfield-OFF"
    Flash 36, F36
	NFadePm 37, p37, "JoustPlayfield-ON", "JoustPlayfield-OFF"
    Flash 37, F37
	NFadePm 38, p38, "JoustPlayfield-ON", "JoustPlayfield-OFF"
    Flash 38, F38
	NFadePm 39, p39, "JoustPlayfield-ON", "JoustPlayfield-OFF"
    Flash 39, F39
	NFadePm 40, p40, "JoustPlayfield-ON", "JoustPlayfield-OFF"
	Flash 40, F40
	NFadePm 41, p41, "JoustPlayfield-ON", "JoustPlayfield-OFF"
    Flash 41, F41
	NFadePm 42, p42, "JoustPlayfield-ON", "JoustPlayfield-OFF"
    Flash 42, F42
	NFadePm 43, p43, "JoustPlayfield-ON", "JoustPlayfield-OFF"
    Flash 43, F43
	NFadePm 44, p44, "JoustPlayfield-ON", "JoustPlayfield-OFF"
    Flash 44, F44
	NFadePm 45, p45, "JoustPlayfield-ON", "JoustPlayfield-OFF"
    Flash 45, F45
	NFadePm 46, p46, "JoustPlayfield-ON", "JoustPlayfield-OFF"
    Flash 46, F46
	NFadePm 47, p47, "JoustPlayfield-ON", "JoustPlayfield-OFF"
    Flash 47, F47
	NFadePm 48, p48, "JoustPlayfield-ON", "JoustPlayfield-OFF"
    Flash 48, F48
	NFadePm 49, p49, "JoustPlayfield-ON", "JoustPlayfield-OFF"
    Flash 49, F49
	NFadePm 50, p50, "JoustPlayfield-ON", "JoustPlayfield-OFF"
    Flash 50, F50
	NFadePm 51, p51, "JoustPlayfield-ON", "JoustPlayfield-OFF"
    Flash 51, F51
	NFadePm 52, p52, "JoustPlayfield-ON", "JoustPlayfield-OFF"
    Flash 52, F52
	NFadePm 53, p53, "JoustPlayfield-ON", "JoustPlayfield-OFF"
    Flash 53, F53
	NFadePm 54, p54, "JoustPlayfield-ON", "JoustPlayfield-OFF"
    Flash 54, F54
	NFadePm 55, p55, "JoustPlayfield-ON", "JoustPlayfield-OFF"
    Flash 55, F55
	NFadePm 56, p56, "JoustPlayfield-ON", "JoustPlayfield-OFF"
    Flash 56, F56
	NFadePm 57, p57, "JoustPlayfield-ON", "JoustPlayfield-OFF"
    Flash 57, F57
	NFadePm 58, p58, "JoustPlayfield-ON", "JoustPlayfield-OFF"
    Flash 58, F58
	NFadePm 59, p59, "JoustPlayfield-ON", "JoustPlayfield-OFF"
    Flash 59, F59
	NFadePm 60, p60, "JoustPlayfield-ON", "JoustPlayfield-OFF"
    Flash 60, F60
	NFadePm 61, p61, "JoustPlayfield-ON", "JoustPlayfield-OFF"
    Flash 61, F61
	NFadePm 62, p62, "JoustPlayfield-ON", "JoustPlayfield-OFF"
    Flash 62, F62
	NFadePm 63, p63, "JoustPlayfield-ON", "JoustPlayfield-OFF"
    Flash 63, F63
	NFadePm 64, p64, "JoustPlayfield-ON", "JoustPlayfield-OFF"
    Flash 64, F64

End Sub

Sub FadePrim(nr, pri, a, b, c, d)
    Select Case FadingLevel(nr)
        Case 2:pri.image = d:FadingLevel(nr) = 0
        Case 3:pri.image = c:FadingLevel(nr) = 1
        Case 4:pri.image = b:FadingLevel(nr) = 2
        Case 5:pri.image = a:FadingLevel(nr) = 3
    End Select
End Sub

''Lights

Sub NFadeL(nr, a)
    Select Case FadingLevel(nr)
        Case 4:a.state = 0:FadingLevel(nr) = 0
        Case 5:a.State = 1:FadingLevel(nr) = 1
    End Select
End Sub

Sub NFadeLm(nr, a)
    Select Case FadingLevel(nr)
        Case 4:a.state = 0
        Case 5:a.State = 1
    End Select
End Sub

Sub NFadeP(nr, a, b, c)
    Select Case FadingLevel(nr)
        Case 4:a.image = c:FadingLevel(nr) = 0
        Case 5:a.image = b:FadingLevel(nr) = 1
    End Select
End Sub

Sub NFadePm(nr, a, b, c)
    Select Case FadingLevel(nr)
        Case 4:a.image = c
        Case 5:a.image = b
    End Select
End Sub

Sub Flash(nr, object)
    Select Case FlashState(nr)
        Case 0 'off
            FlashLevel(nr) = FlashLevel(nr) - FlashSpeedDown
            If FlashLevel(nr) < 0 Then
                FlashLevel(nr) = 0
                FlashState(nr) = -1 'completely off
            End if
            Object.opacity = FlashLevel(nr)
        Case 1 ' on
            FlashLevel(nr) = FlashLevel(nr) + FlashSpeedUp
            If FlashLevel(nr) > 1000 Then
                FlashLevel(nr) = 1000
                FlashState(nr) = -2 'completely on
            End if
            Object.opacity = FlashLevel(nr)
    End Select
End Sub

Sub Flashm(nr, object) 'multiple flashers, it doesn't change the flashstate
    Select Case FlashState(nr)
        Case 0         'off
            Object.opacity = FlashLevel(nr)
        Case 1         ' on
            Object.opacity = FlashLevel(nr)
    End Select
End Sub

 Sub AllLampsOff():For x = 1 to 200:LampState(x) = 4:FadingLevel(x) = 4:Next:UpdateLamps:UpdateLamps:Updatelamps:End Sub


Sub SetLamp(nr, value)
    If value = 0 AND LampState(nr) = 0 Then Exit Sub
    If value = 1 AND LampState(nr) = 1 Then Exit Sub
    LampState(nr) = abs(value) + 4
FadingLevel(nr ) = abs(value) + 4: FadingState(nr ) = abs(value) + 4
End Sub

Sub SetFlash(nr, stat)
    FlashState(nr) = ABS(stat)
End Sub



'********FLASHERS*********


Sub FlasherTimer_Timer()
Flashm 1, f1a
Flash 1, f1
Flashm 2, f2a
Flash 2, f2
Flashm 3, f3a
Flash 3, f3
Flashm 4, f4a
Flash 4, f4


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


Sub Table1_Exit
    Controller.Stop
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

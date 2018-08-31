Option Explicit
Randomize

' Thalamus 2018-07-24
' Added/Updated "Positional Sound Playback Functions" and "Supporting Ball & Sound Functions"
' Changed UseSolenoids=1 to 2
' Table uses non stanard ball rolling routine
' No special SSF tweaks yet.
' Wob 2018-08-08
' Added vpmInit Me to table init

Const cGameName="txsector",UseSolenoids=2,UseLamps=1,UseGI=0,SSolenoidOn="SolOn",SSolenoidOff="SolOff",SFlipperOn="FlipperUp",SFlipperOff="FlipperDown"
Const SCoin="coin3",cCredits="Destruk & TAB & MNPG. Dip settings menu added by Inkochnito"

LoadVPM "01210000","sys80.vbs",3.1

 '**********************TABLE OPTIONS************************************************************************************************
'1=VPinMAME, 2=UVP backglass server, 3=B2S backglass server(and implicitely DOF), 4= B2S, implicitely DOF and disable mech sounds
Const cController = 3
'***********************************************************************************************************************************


Sub LoadVPM(VPMver, VBSfile, VBSver)
  On Error Resume Next
  If ScriptEngineMajorVersion < 5 Then MsgBox "VB Script Engine 5.0 or higher required"
  ExecuteGlobal GetTextFile(VBSfile)
  If Err Then MsgBox "Unable to open " & VBSfile & ". Ensure that it is in the same folder as this table. " & vbNewLine & Err.Description
  Select Case cController
  Case 1
    Set Controller = CreateObject("VPinMAME.Controller")
	If Err Then MsgBox "Can't Load VPinMAME." & vbNewLine & Err.Description
	If VPMver>"" Then If Controller.Version < VPMver Or Err Then MsgBox "VPinMAME ver " & VPMver & " required."
	If VPinMAMEDriverVer < VBSver Or Err Then MsgBox VBSFile & " ver " & VBSver & " or higher required."
  Case 2
	Set Controller = CreateObject("UltraVP.BackglassServ")
  Case else
    Set Controller = CreateObject("B2S.Server")
  End Select
  If Err Then MsgBox "Can't Load VPinMAME." & vbNewLine & Err.Description
  If VPMver > "" Then If Controller.Version < VPMver Or Err Then MsgBox "VPinMAME ver " & VPMver & " required."
  If VPinMAMEDriverVer < VBSver Or Err Then MsgBox VBSFile & " ver " & VBSver & " or higher required."
  On Error Goto 0
End Sub

'Sub LoadVPM(VPMver, VBSfile, VBSver)
'	On Error Resume Next
'		If ScriptEngineMajorVersion < 5 Then MsgBox "VB Script Engine 5.0 or higher required"
'		ExecuteGlobal GetTextFile(VBSFile)
'		If Err Then MsgBox "Unable to open " & VBSfile & ". Ensure that it is in the same folder as this table. " & vbNewLine & Err.Description : Err.Clear
'		Set Controller = CreateObject("VPinMAME.Controller")
'		If Err Then MsgBox "Can't Load VPinMAME." & vbNewLine & Err.Description
'		If VPMver>"" Then If Controller.Version < VPMver Or Err Then MsgBox "VPinMAME ver " & VPMver & " required." : Err.Clear
'		If VPinMAMEDriverVer < VBSver Or Err Then MsgBox VBSFile & " ver " & VBSver & " or higher required."
'	On Error Goto 0
'End Sub

'==============================
Sub Table1_KeyDown(ByVal keycode)

If KeyCode=LeftFlipperKey Then Controller.Switch(6)=1
If KeyCode=RightFlipperKey Then Controller.Switch(16)=1

    If Keycode = 3 Then
    'KickerTRE.kick 180,5
    'kICKERTRE1.kick 240,7
    'KickerPassageBottom.kick 200,5
    'KickerPassageBottom1.kick 120,5
     'MsgBox L2.State
    'BallRelease.CreateBall
	'BallRelease.Kick 90,10
    MsgBox cController
    End If

	If keycode = PlungerKey Then
		Plunger.PullBack
		PlaySound "plungerpull",0,1,0.25,0.25
	End If

	'If keycode = LeftFlipperKey Then
	'	LeftFlipper.RotateToEnd
    '    Flipper1.RotateToEnd
	'	PlaySound "fx_flipperup", 0, .67, -0.05, 0.05
	'End If

	'If keycode = RightFlipperKey Then
	'	RightFlipper.RotateToEnd
	'	PlaySound "fx_flipperup", 0, .67, 0.05, 0.05
	'End If

	If keycode = LeftTiltKey Then
		Nudge 90, 2
	End If

	If keycode = RightTiltKey Then
		Nudge 270, 2
	End If

	If keycode = CenterTiltKey Then
		Nudge 0, 2
	End If
    If vpmKeyDown(KeyCode) Then Exit Sub
End Sub

Sub Table1_KeyUp(ByVal keycode)

If KeyCode=LeftFlipperKey Then Controller.Switch(6)=0
If KeyCode=RightFlipperKey Then Controller.Switch(16)=0

	If keycode = PlungerKey Then
		Plunger.Fire
		PlaySound "plunger",0,1,0.25,0.25
	End If

	'If keycode = LeftFlipperKey Then
    '	LeftFlipper.RotateToStart
    '    Flipper1.RotateToStart
    '    PlaySound "fx_flipperdown", 0, 1, -0.05, 0.05
	'End If

	'If keycode = RightFlipperKey Then
	'	RightFlipper.RotateToStart
	'	PlaySound "fx_flipperdown", 0, 1, 0.05, 0.05
	'End If
    If vpmKeyUp(KeyCode) Then Exit Sub
End Sub

SolCallback(1)="TopLeft"
SolCallback(2)="BotLeft"
SolCallback(3)="dtR.SolDropUp"
SolCallback(4)="dtL.SolDropUp"
SolCallback(5)="TopRight"
SolCallback(6)="BotRight"
SolCallback(7)="dtOne.SolDropUp"
SolCallback(8)="VpmSolSound""knocker"","
SolCallback(9)="SolOuthole"
SolCallback(sLLFlipper)="VpmSolFlipper LeftFlipper,Flipper1,"
SolCallback(sLRFlipper)="VpmSolFlipper RightFlipper,nothing,"

Dim TroughBalls
TroughBalls=3
Dim DrainBalls
DrainBalls=True

Sub SolOutHole(Enabled)
	If DrainBalls Then
		Drain.Destroyball
		TroughBalls=TroughBalls+1
		If TroughBalls=3 Then Controller.Switch(56)=1
		Controller.Switch(66)=0
        'LightS.State = 1
		DrainBalls=False
	End If
End Sub

Dim SStat
SStat=0

'Sub TopLeft(Enabled)
'	If Enabled Then
'		If SStat=0 And bsLTop.Balls Then bsLTop.ExitSol_On
'	End If
'End Sub

Sub TopLeft(Enabled)
If Enabled AND SStat = 0 AND bsLTop.Balls Then
   bsLTop.ExitSol_On
   DOF 110,2
   End If
   If Enabled AND SStat = 1 Then
   BLR.State = 1
   FlasherR.Visible = 1
   DOF 111,1
   else
   BLR.State = 0
   FlasherR.Visible = 0
   DOF 111,0
   End If
End Sub

'Sub BotLeft(Enabled)
'	If Enabled Then
'		If SStat=0 And bsLBot.Balls Then
'			'Kicker4.TimerEnabled=1
'			bsLBot.ExitSol_On
'		End If
'	End If
'End Sub

Sub BotLeft(Enabled)
If Enabled AND SStat = 0 AND bsLbot.Balls Then
   bsLBot.ExitSol_On
   DOF 112,2
   KickerTRE.TimerEnabled = 1
   End If
   If Enabled AND SStat = 1 Then
   FRO.State = 1
   FlasherO.visible = 1
   DOF 113,1
   else
   FRO.State = 0
   FlasherO.visible = 0
   DOF 113,0
   End If
End Sub

Sub KickerTRE_Timer
	KickerTRE.TimerEnabled=0
	If bsLTop.Balls Then bsLTop.ExitSol_On
End Sub

'Sub TopRight(Enabled)
'	If Enabled Then
'		If SStat=0 And bsRTop.Balls Then bsRTop.ExitSol_On
'	End If
'End Sub

Sub TopRight(Enabled)
If Enabled AND SStat = 0 AND bsRTop.Balls Then
   bsRTop.ExitSol_On
   DOF 114,2
   End If
   If Enabled AND SStat = 1 Then
   LR.State = 1:LRA.State = 1:FlasherLR.visible = 1
   DOF 115,1
   else
   LR.State = 0:LRA.State = 0:FlasherLR.visible = 0
   DOF 115,0
   End If
End Sub

'Sub BotRight(Enabled)
'	If Enabled Then
'		If SStat=0 And bsRBot.Balls Then bsRBot.ExitSol_On
'	End If
 '   If Enabled Then
 '       If SStat = 1 Then RR.State = 1:RRA.State = 1
 '   End If
'End Sub

Sub BotRight(Enabled)
   If Enabled AND SStat = 0 AND bsRbot.Balls Then
   bsRBot.ExitSol_On
   DOF 116,2
   End If
   If Enabled AND SStat = 1 Then
   RR.State = 1:RRA.State = 1:FlasherRR.visible = 1
   DOF 117,1
   else
   RR.State = 0:RRA.State = 0:FlasherRR.visible = 0
   DOF 117,0
   End If
End Sub

Sub UpdateLeftFlipperLogo()
	LFLogo.RotY = LeftFlipper.CurrentAngle
End Sub
Sub UpdateLeftMiniFlipperLogo()
	LFLogoMini.RotY = Flipper1.CurrentAngle
End Sub
Sub UpdateRightFlipperLogo()
	RFLogo.RotY = RightFlipper.CurrentAngle
End Sub

'Main INIT
Dim dtL,dtR,dtOne,bsRTop,bsRBot,bsLTop,bsLBot

Sub Table1_Init
vpmInit Me
Controller.Games(cGameName).Settings.Value("dmd_red")=0
Controller.Games(cGameName).Settings.Value("dmd_green")=128
Controller.Games(cGameName).Settings.Value("dmd_blue")=255
	On Error Resume Next
	With Controller
		.GameName=cGameName
		If Err Then MsgBox "Can't start Game" & cGameName & vbNewLine & Err.Description : Exit Sub
		.SplashInfoLine=cCredits
		.HandleMechanics=0
		.HandleKeyboard=0
		.ShowDMDOnly=1
		.ShowFrame=0
		.ShowTitle=0
        '.Hidden = 1
If Table1.ShowDT = false then
    'Scoretext.Visible = false
    .Hidden = 1
End If

If Table1.ShowDT = true then
'Scoretext.Visible = false
    .Hidden = 0
End If
	End With
		Controller.SolMask(0)=0
		vpmTimer.AddTimer 1000,"Controller.SolMask(0)=&Hffffffff'"
		Controller.Run
		If Err Then MsgBox Err.Description
	On Error Goto 0


PinMAMETimer.Interval=PinMAMEInterval
PinMAMETimer.Enabled=1
vpmNudge.TiltSwitch=57
vpmNudge.Sensitivity=5
vpmNudge.TiltObj=Array(LeftSlingshot,RightSlingshot,Bumper1,Bumper2,Bumper3,LminiKick,RminiKick)

	set dtL=New cvpmDropTarget
	dtL.InitDrop Array(SW40,SW50,SW60,SW70),Array(40,50,60,70)
	dtL.initsnd "flapclos","flapopen"

	set dtR=New cvpmDropTarget
	dtR.InitDrop Array(SW41,SW51,SW61,SW71),Array(41,51,61,71)
	dtR.initsnd "flapclos","flapopen"

	set dtOne=New cvpmDropTarget
	dtOne.InitDrop SW42,42
	dtOne.initsnd "flapclos","flapopen"

	Set bsLTop=New cvpmBallStack
	bsLTop.InitSaucer KickerTRE,45,205,1
	bsLTop.InitExitSnd "popper","popper"

	Set bsLBot=New cvpmBallStack
	bsLBot.InitSaucer KickerTRE1,55,200,5
	bsLBot.InitExitSnd "popper","popper"

	Set bsRTop=New cvpmBallStack
	bsRTop.InitSaucer KickerPassageBottom,65,230,1
	bsRTop.InitExitSnd "popper","popper"

	Set bsRBot=New cvpmBallStack
	bsRBot.InitSaucer KickerPassageBottom1,75,120,10
	bsRBot.InitExitSnd "popper","popper"

	Controller.Switch(56)=1
	Controller.Switch(66)=1

LminiKick1.isdropped = 1
LminiKick2.isdropped = 1
RminiKick1.isdropped = 1
RminiKick2.isdropped = 1
SW64a.isdropped = 1
End Sub

'SWITCHES

Sub Drain_Hit()
PlaySound "drain"
Controller.Switch(66)=1
DrainBalls=True
End Sub

Sub SW63_Hit():Controller.Switch(63) = 1:End Sub
Sub SW63_Unhit():Controller.Switch(63) = 0:End Sub
Sub SW46L_Hit():Controller.Switch(46) = 1:End Sub
Sub SW46L_Unhit():Controller.Switch(46) = 0:End Sub
Sub SW46R_Hit():Controller.Switch(46) = 1:End Sub
Sub SW46R_Unhit():Controller.Switch(46) = 0:End Sub
Sub SW35_Hit():Controller.Switch(35) = 1:End Sub
Sub SW35_Unhit():Controller.Switch(35) = 0:End Sub
Sub SW36_Hit():Controller.Switch(36) = 1:End Sub
Sub SW36_Unhit():Controller.Switch(36) = 0:End Sub
Sub SW43_Hit():Controller.Switch(43) = 1:End Sub
Sub SW43_Unhit():Controller.Switch(43) = 0:End Sub
Sub SW41_Hit():dtR.Hit 1:End Sub
Sub SW44_Hit():vpmTimer.PulseSw 44:End Sub
Sub SW51_Hit():dtR.Hit 2:End Sub
Sub SW61_Hit():dtR.Hit 3:End Sub
Sub SW71_Hit():dtR.Hit 4:End Sub
Sub SW40_Hit():dtL.Hit 1:End Sub
Sub SW50_Hit():dtL.Hit 2:End Sub
Sub SW60_Hit():dtL.Hit 3:End Sub

Sub SW64_Hit()
vpmTimer.PulseSw 64
SW64.IsDropped = 1
SW64a.IsDropped = 0
me.timerenabled = 1
End Sub

Sub SW64_Timer()
SW64a.IsDropped = 1
SW64.IsDropped = 0
End Sub

Sub SW70_Hit():dtL.Hit 4:End Sub
Sub SW42_Hit():dtOne.Hit 1:End Sub

Sub RR_Help_Hit()
'MsgBox ActiveBall.VelY
If ActiveBall.VelY < -10 Then
PlaySound "rail"
End If
End Sub

Sub TR_Help_Hit()
'MsgBox ActiveBall.VelY
PlaySound "rail2"
End Sub

Sub TR_Help1_Hit()
'MsgBox ActiveBall.VelY
PlaySound "scoopenter", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0
End Sub

Sub SW53_Hit()
vpmTimer.PulseSw 53
sw53p.transx = -7
Me.TimerEnabled = 1
End Sub

Sub sw53_Timer
sw53p.transx = 0
Me.TimerEnabled = 0
End Sub

Sub SW54_Hit()
vpmTimer.PulseSw 54
sw54p.transx = -7
Me.TimerEnabled = 1
End Sub

Sub sw54_Timer
sw54p.transx = 0
Me.TimerEnabled = 0
End Sub

Sub SW74_Spin:vpmTimer.PulseSw 74:PlaySound "fx_spinner":End Sub
Sub SW73_Spin:vpmTimer.PulseSw 73:PlaySound "fx_spinner":End Sub
Sub LBANKSW_SlingShot():vpmTimer.PulseSw 76:End Sub
Sub TBANKSW_SlingShot():vpmTimer.PulseSw 76:End Sub
Sub TR72SW_SlingShot():vpmTimer.PulseSw 76:End Sub
Sub LminiKick_SlingShot():DOF 123,2:PlaySound "left_slingshot":LminiKick.isdropped = 1:me.timerenabled = 1:me.timerinterval = 20:leftm=1:vpmTimer.PulseSw 76:End Sub
Sub RminiKick_SlingShot():DOF 124,2:PlaySound "right_Slingshot":RminiKick.isdropped = 1:me.timerenabled = 1:me.timerinterval = 20:rightm=1:vpmTimer.PulseSw 76:End Sub

Dim leftm
Sub LminiKick_Timer()
Select Case leftm
Case 1:LMiniKick1.isdropped = 0:leftm=2
Case 2:LMiniKick1.isdropped = 1:LMiniKick2.isdropped = 0:leftm=3
Case 3:LMiniKick2.isdropped = 1:LMiniKick1.isdropped = 0:leftm=4
Case 4:LMiniKick1.isdropped = 1:LMiniKick.isdropped = 0:me.timerenabled = 0
End Select
End Sub

Dim rightm
Sub RminiKick_Timer()
Select Case rightm
Case 1:RMiniKick1.isdropped = 0:rightm=2
Case 2:RMiniKick1.isdropped = 1:RMiniKick2.isdropped = 0:rightm=3
Case 3:RMiniKick2.isdropped = 1:RMiniKick1.isdropped = 0:rightm=4
Case 4:RMiniKick1.isdropped = 1:RMiniKick.isdropped = 0:me.timerenabled = 0
End Select
End Sub

'Right Hole
Sub KickerPassageBottom_Hit:bsRTop.AddBall 0:PlaySound "kicker_enter_center", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0:End Sub								'65 Top Right
Sub KickerPassageBottom1_Hit:bsRBot.AddBall 0:PlaySound "kicker_enter_center", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0:End Sub								'75 Bottom Right

'Left Hole
Sub KickerTRE_Hit:bsLTop.AddBall 0:PlaySound "kicker_enter_center", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0:End Sub								'45 Top Left
Sub KickerTRE1_Hit:bsLBot.AddBall 0:PlaySound "kicker_enter_center", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0:End Sub								'55 Bottom Left
Sub KickerTR_Hit:PlaySound "scoopenter", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0:End Sub

Set MotorCallback=GetRef("UpdateMultipleLamps")

Dim CurrentB,OldB,L13,CurrentA,OldA,L12,GIX
CurrentB=0:OldB=0:CurrentA=0:OldA=0:L12=0

Sub UpdateMultipleLamps
	SStat=Light12.State
	CurrentA=Light13.State
		If CurrentA<>OldA Then
			If CurrentA=1 Then dtOne.Hit 1
		End If
	OldA=CurrentA
	CurrentB=Light2.State
		If CurrentB<>OldB Then
		If CurrentB=1 Then
			If TroughBalls>0 Then
				BallRelease.CreateBall
				BallRelease.Kick 90,3
				PlaySound"ballrel"
                DOF 125,2
				TroughBalls=TroughBalls-1
			End If
			If TroughBalls=0 Then Controller.Switch(56)=0
		End If
		End If
	OldB=CurrentB
If L1.State = 0 Then
For each GIX in GI:GIX.State = 1:next
LFLogoMini.Image = "flipper_remake_leftlit"
LFLogo.image = "flipper_remake_leftlit"
RFLogo.image = "flipper_remake_rightlit"
else
For each GIX in GI:GIX.State = 0:next
LFLogoMini.Image = "flipper_remake_left"
LFLogo.image = "flipper_remake_left"
RFLogo.image = "flipper_remake_right"
End If
UpdateLeftFlipperLogo
UpdateLeftMiniFlipperLogo
UpdateRightFlipperLogo
End Sub

Set Lights(1)=L1
Set Lights(2)=Light2
Lights(3)=Array(L3,L3A)
Set Lights(5)=L5
Set Lights(6)=L6
Set Lights(7)=L7
Set Lights(8)=L8
Set Lights(9)=L9
Set Lights(10)=L10
Lights(11)=Array(Light11,Light11A)
Set Lights(12)=Light12
Set Lights(13)=Light13
Lights(15)=Array(L15,L15A,L15B,L15C,L15C2,L15C3)
Lights(16)=Array(L16,L16A,L16B,L16C,L16C1,L15C2)
Lights(17)=Array(L17,L17A,L17B,L17C,L17C1,L17C2)
Lights(18)=Array(L18,L18A,L18B,L18C,l18C1,L18C2)
Lights(19)=Array(L19,L19A,L19B,L19C,L19C1,L19C2)
Set Lights(20)=L20
Set Lights(21)=L21
Set Lights(22)=L22
Set Lights(23)=L23
Set Lights(24)=L24
Set Lights(25)=L25
Set Lights(26)=L26
Set Lights(27)=L27
Set Lights(28)=L28
Set Lights(29)=L29
Set Lights(30)=L30
Set Lights(31)=L31
Set Lights(32)=L32
Set Lights(33)=L33
Set Lights(34)=L34
Set Lights(35)=L35
Set Lights(36)=L36
Set Lights(37)=L37
Set Lights(38)=L38
Set Lights(39)=L39
Set Lights(40)=L40
Set Lights(41)=L41
Set Lights(42)=L42
Set Lights(43)=L43
Set Lights(44)=L44
Set Lights(45)=L45
Set Lights(46)=L46
Lights(47)=Array(L47,L47A)
'Set Lights(48)=L48
Set Lights(50)=L50
Lights(51)=Array(L51,L51A)

'Sub Plunger_Init()
'	PlaySound "ballrelease",0,0.5,0.5,0.25
'	'Plunger.CreateBall
'	BallRelease.CreateBall
'	BallRelease.Kick 90, 8
'	BIP = BIP +1
'End Sub


Sub Bumper1_Hit
    DOF 118,2
    vpmTimer.PulseSw 72
	PlaySound "fx_bumper4"
	'B1L1.State = 1:B1L2. State = 1
	Me.TimerEnabled = 1
End Sub

Sub Bumper1_Timer
	'B1L1.State = 0:B1L2. State = 0
	Me.Timerenabled = 0
End Sub

Sub Bumper2_Hit
    DOF 119,2
    vpmTimer.PulseSw 52
	PlaySound "fx_bumper4"
	'B2L1.State = 1:B2L2. State = 1
	Me.TimerEnabled = 1
End Sub

Sub Bumper2_Timer
	'B2L1.State = 0:B2L2. State = 0
	Me.Timerenabled = 0
End Sub

Sub Bumper3_Hit
    DOF 120,2
    vpmTimer.PulseSw 62
	PlaySound "fx_bumper4"
	'B3L1.State = 1:B3L2. State = 1
	Me.TimerEnabled = 1
End Sub

Sub Bumper3_Timer
	'B3L1.State = 0:B3L2. State = 0
	Me.Timerenabled = 0
End Sub


'**********Sling Shot Animations
' Rstep and Lstep  are the variables that increment the animation
'****************
Dim RStep, Lstep

Sub RightSlingShot_Slingshot
    DOF 121,2
    vpmTimer.PulseSw 76
    PlaySound "left_slingshot", 0, 1, 0.05, 0.05
    RSling.Visible = 0
    RSling1.Visible = 1
    sling1.TransZ = -20
    RStep = 0
    RightSlingShot.TimerEnabled = 1
	'gi1.State = 0:Gi2.State = 0
End Sub

Sub RightSlingShot_Timer
    Select Case RStep
        Case 3:RSLing1.Visible = 0:RSLing2.Visible = 1:sling1.TransZ = -10
        Case 4:RSLing2.Visible = 0:RSLing.Visible = 1:sling1.TransZ = 0:RightSlingShot.TimerEnabled = 0
    End Select
    RStep = RStep + 1
End Sub

Sub LeftSlingShot_Slingshot
    DOF 122,2
    vpmTimer.PulseSw 76
    PlaySound "right_slingshot",0,1,-0.05,0.05
    LSling.Visible = 0
    LSling1.Visible = 1
    sling2.TransZ = -20
    LStep = 0
    LeftSlingShot.TimerEnabled = 1
	'gi3.State = 0:Gi4.State = 0
End Sub

Sub LeftSlingShot_Timer
    Select Case LStep
        Case 3:LSLing1.Visible = 0:LSLing2.Visible = 1:sling2.TransZ = -10
        Case 4:LSLing2.Visible = 0:LSLing.Visible = 1:sling2.TransZ = 0:LeftSlingShot.TimerEnabled = 0
    End Select
    LStep = LStep + 1
End Sub

'************************************
' What you need to add to your table
'************************************

' a timer called CollisionTimer. With a fast interval, like from 1 to 10
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

' The Double For loop: This is a double cycle used to check the collision between a ball and the other ones.
' If you look at the parameters of both cycles, you’ll notice they are designed to avoid checking
' collision twice. For example, I will never check collision between ball 2 and ball 1,
' because I already checked collision between ball 1 and 2. So, if we have 4 balls,
' the collision checks will be: ball 1 with 2, 1 with 3, 1 with 4, 2 with 3, 2 with 4 and 3 with 4.

' Sum first the radius of both balls, and if the height between them is higher then do not calculate anything more,
' since the balls are on different heights so they can't collide.

' The next 3 lines calculates distance between xth and yth balls with the Pytagorean theorem,

' The first "If": Checking if the distance between the two balls is less than the sum of the radius of both balls,
' and both balls are not already colliding.

' Why are we checking if balls are already in collision?
' Because we do not want the sound repeting when two balls are resting closed to each other.

' Set the collision property of both balls to True, and we assign the number of the ball colliding

' Play the collide sound of your choice using the VOL, PAN and PITCH functions

' Last lines: If the distance between 2 balls is more than the radius of a ball,
' then there is no collision and then set the collision property of the ball to False (-1).


Sub Pins_Hit (idx)
	PlaySound "pinhit_low", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0, AudioFade(ActiveBall)
End Sub

Sub Targets_Hit (idx)
	PlaySound "target"
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

Sub Rollovers_Hit (idx)
    PlaySound "rollover"
End Sub

Sub Start_Ramp_Hit
    PlaySound "Ball_Bounce", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
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

Sub editDips
	Dim vpmDips : Set vpmDips = New cvpmDips
	With vpmDips
		.AddForm  700,400,"TX Sector - DIP switches"
		.AddFrame 2,4,190,"Maximum credits",49152,Array("8 credits",0,"10 credits",32768,"15 credits",&H00004000,"20 credits",49152)'dip 15&16
		.AddFrame 2,80,190,"Coin chute 1 and 2 control",&H00002000,Array("seperate",0,"same",&H00002000)'dip 14
		.AddFrame 2,126,190,"Playfield special",&H00200000,Array("replay",0,"extra ball",&H00200000)'dip 22
		.AddFrame 2,172,190,"High games to date control",&H00000020,Array("no effect",0,"reset high games 2-5 on power off",&H00000020)'dip6
		.AddFrame 2,218,190,"Auto-percentage control",&H00000080,Array("disabled (normal high score mode)",0,"enabled",&H00000080)'dip 8
		.AddFrame 2,264,190,"Playfield special control",&H40000000,Array("special lit after banks compl. 3 times",0,"special lit after banks compl. 2 times",&H40000000)'dip 31
		.AddFrame 2,310,190,"Extra bal control",&H80000000,Array("Reset energy level",0,"Memorize energy level",&H80000000)'dip 32
		.AddFrame 205,4,190,"High score to date awards",&H00C00000,Array("not displayed and no award",0,"displayed and no award",&H00800000,"displayed and 2 replays",&H00400000,"displayed and 3 replays",&H00C00000)'dip 23&24
		.AddFrame 205,80,190,"Balls per game",&H01000000,Array("5 balls",0,"3 balls",&H01000000)'dip 25
		.AddFrame 205,126,190,"Replay limit",&H04000000,Array("no limit",0,"one per ball",&H04000000)'dip 27
		.AddFrame 205,172,190,"Novelty",&H08000000,Array("normal",0,"extra ball and replay scores 500000",&H08000000)'dip 28
		.AddFrame 205,218,190,"Game mode",&H10000000,Array("replay",0,"extra ball",&H10000000)'dip 29
		.AddFrame 205,264,190,"3rd coin chute credits control",&H20000000,Array("no effect",0,"add 9",&H20000000)'dip 30
		.AddChk 205,316,180,Array("Match feature",&H02000000)'dip 26
		.AddChk	205,331,190,Array("Attract sound",&H00000040)'dip 7
		.AddLabel 50,360,300,20,"After hitting OK, press F3 to reset game with new settings."
		.ViewDips
	End With
End Sub
Set vpmShowDips = GetRef("editDips")

Sub DOF(dofevent, dofstate)
	If cController = 3 Then
		If dofstate = 2 Then
			Controller.B2SSetData dofevent, 1:Controller.B2SSetData dofevent, 0
		Else
			Controller.B2SSetData dofevent, dofstate
		End If
	End If
End Sub

'-------------------------------

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
'    JP's VP10 Collision & Rolling Sounds
'*****************************************

Const tnob = 3 ' total number of balls
ReDim rolling(tnob)
ReDim collision(tnob)
Initcollision

Sub Initcollision
    Dim i
    For i = 0 to tnob
        collision(i) = -1
        rolling(i) = False
    Next
End Sub

Sub CollisionTimer_Timer()
    Dim BOT, B, B1, B2, dx, dy, dz, distance, radii
    BOT = GetBalls

    ' rolling

	For B = UBound(BOT) +1 to tnob
        rolling(b) = False
        StopSound("fx_ballrolling" & b)
	Next

    If UBound(BOT) = -1 Then Exit Sub

    For B = 0 to UBound(BOT)
        If BallVel(BOT(b) ) > 1 AND BOT(b).z < 30 Then
            rolling(b) = True
            PlaySound("fx_ballrolling" & b), -1, Vol(BOT(b) ), Pan(BOT(b) ), 0, Pitch(BOT(b) ), 1, 0
        Else
            If rolling(b) = True Then
                StopSound("fx_ballrolling" & b)
                rolling(b) = False
            End If
        End If
    Next

    'collision

    If UBound(BOT) < 1 Then Exit Sub

    For B1 = 0 to UBound(BOT)
        For B2 = B1 + 1 to UBound(BOT)
            dz = INT(ABS((BOT(b1).z - BOT(b2).z) ) )
            radii = BOT(b1).radius + BOT(b2).radius
			If dz <= radii Then

            dx = INT(ABS((BOT(b1).x - BOT(b2).x) ) )
            dy = INT(ABS((BOT(b1).y - BOT(b2).y) ) )
            distance = INT(SQR(dx ^2 + dy ^2) )

            If distance <= radii AND (collision(b1) = -1 OR collision(b2) = -1) Then
                collision(b1) = b2
                collision(b2) = b1
                PlaySound("fx_collide"), 0, Vol2(BOT(b1), BOT(b2)), Pan(BOT(b1)), 0, Pitch(BOT(b1)), 0, 0
            Else
                If distance > (radii + 10)  Then
                    If collision(b1) = b2 Then collision(b1) = -1
                    If collision(b2) = b1 Then collision(b2) = -1
                End If
            End If
			End If
        Next
    Next
End Sub


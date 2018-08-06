'
'
' Bone Busters Inc. / IPD No. 347 / August, 1989 / 4 Players
' http://www.ipdb.org/machine.cgi?id=347
' version 1.0 created in VPX by MaX
'
' \_  _/  /| |\/
'  \\//  / | |/\
'         /
'
'v1.0 - 2015, December, 24th
'
' Thanks to:
' gtxjoe for his resources and the ramp, the scoop and the flasher dome primitives
' jpsalas for allowing me to use his scripts and resources
' Unclewilly for allowing me to use his scripts and resources
' Destruk/ExtremeMame/Gravitar/TAB for the VP8 table
' Inkochnito for the DIP switches

' Thalamus 2018-07-19
' Added/Updated "Positional Sound Playback Functions" and "Supporting Ball & Sound Functions"
' Table has the 4th light on all the time - weird.
' Changed UseSolenoids=1 to 2
' No special SSF tweaks yet.

Option Explicit

'**************************************
'**************************************
'Choose Controller: 1-VPM, 2-UVP, 3-B2S
'**************************************
'**************************************

Const cController = 3

'**************************************
'**************************************
'**************************************
'**************************************

LoadVPM"01520000","sys80.VBS",1.2

Sub LoadVPM(VPMver, VBSfile, VBSver)
	On Error goto 0
		If ScriptEngineMajorVersion < 5 Then MsgBox "VB Script Engine 5.0 or higher required"
		ExecuteGlobal GetTextFile(VBSfile)
		If Err Then MsgBox "Unable to open " & VBSfile & ". Ensure that it is in the same folder as this table. " & vbNewLine & Err.Description : Err.Clear
        Select Case cController
			Case 1:
				Set Controller = CreateObject("VPinMAME.Controller")
				If Err Then MsgBox "Can't Load VPinMAME." & vbNewLine & Err.Description
				If VPMver>"" Then If Controller.Version < VPMver Or Err Then MsgBox "VPinMAME ver " & VPMver & " required."
				If VPinMAMEDriverVer < VBSver Or Err Then MsgBox VBSFile & " ver " & VBSver & " or higher required."
            Case 2:
				Set Controller = CreateObject("UltraVP.BackglassServ")
            Case 3:
                Set Controller = CreateObject("B2S.Server")
         End Select
	On Error Goto 0
End Sub

Const cGameName="bonebstr",cCredits="Bone Busters",UseSolenoids=2,UseLamps=1,UseGI=0
Const SSolenoidOn="SolOn",SSolenoidOff="SolOff",SFlipperOn="fx_flipperup",SFlipperOff="fx_flipperdown",SCoin="Coin3"

Const sBottomHole=1
Const sTopHole=2
Const sSkull=3
Const sUpKicker=5
Const s5BankReset=6
Const sKnock=8

SolCallback(1)="Bot"
SolCallback(2)="Holes"
SolCallback(3)="KickB"
SolCallback(5)="UPK"
SolCallback(6)="TBank"
SolCallback(8)="vpmSolSound""knocker"","
SolCallback(9)="bsTrough.SolIn"
SolCallback(sLLFlipper)="vpmSolFlipper LeftFlipper,LeftFlipper2,"
SolCallback(sLRFlipper)="vpmSolFlipper RightFlipper,Rightflipper2,"

Dim dtLow,bsTopHole,bsBottom,bsKickback,bsUpKick,bsTrough

Sub BoneBusters_Init
	Controller.Games(cGameName).Settings.Value("dmd_red")=0
	Controller.Games(cGameName).Settings.Value("dmd_green")=128
	Controller.Games(cGameName).Settings.Value("dmd_blue")=255
	vpmInit Me
	On Error Resume Next
	With Controller
		.GameName=cGameName
		If Err Then MsgBox"Can't start Game"&cGameName&vbNewLine&Err.Description:Exit Sub
		.SplashInfoLine=cCredits
		.HandleMechanics=0
		.ShowDMDOnly=1
		.ShowFrame=0
		.ShowTitle=0
		.Hidden = showdt
		If Err Then MsgBox Err.Description
	End With
	On Error Goto 0
	Controller.SolMask(0)=0
	vpmTimer.AddTimer 2000,"Controller.SolMask(0)=&Hffffffff'" 'ignore all solenoids - then add the timer to renable all the solenoids after 2 seconds
	Controller.Run GetPlayerHwnd

'	Controller.Dip(0) = (0*1 + 0*2 + 0*4 + 0*8 + 0*16 + 0*32 + 0*64 + 1*128) '01-08
'	Controller.Dip(1) = (0*1 + 0*2 + 0*4 + 0*8 + 0*16 + 0*32 + 1*64 + 1*128) '09-16
'	Controller.Dip(2) = (0*1 + 0*2 + 0*4 + 0*8 + 0*16 + 1*32 + 1*64 + 1*128) '17-24
'	Controller.Dip(3) = (1*1 + 1*2 + 0*4 + 0*8 + 1*16 + 1*32 + 1*64 + 0*128) '25-32

	PinMAMETimer.Interval=PinMAMEInterval:PinMAMETimer.Enabled=1:vpmNudge.TiltSwitch=57:vpmNudge.Sensitivity=5
	vpmNudge.TiltObj=Array(BumperSW52,LeftFlipper,LeftFlipper2,RightFlipper,RightFlipper2,LeftSlingShot,rightSlingShot,sw50)
	vpmNudge.TiltSwitch=57
	vpmNudge.Sensitivity=5

	Set bsTrough=New cvpmBallStack
	bsTrough.InitSw 56,0,0,55,0,0,0,0
	bsTrough.initkick ballrelease,75,5
	bsTrough.InitExitSnd "ballrelease",""
	bsTrough.Balls=4

	Set dtLow=New cvpmDropTarget
	dtLow.InitDrop Array(droptarget1a,droptarget2a,droptarget3a,droptarget4a,droptarget5a),Array(0,10,20,30,40)
	dtLow.InitSnd "droptargetdown","bankresetA 100P"

	Set bsTopHole=New cvpmBallStack
	bsTopHole.InitSaucer KickerTopHole,26,180,5
	bsTopHole.InitExitSnd"popper_ball",""

	Set bsBottom=New cvpmBallStack
	bsBottom.InitSaucer BottomholeSW36,36,160,5
	bsBottom.InitExitSnd"popper_ball",""

	Set bsKickback=New cvpmBallStack
	bsKickback.InitSaucer KickerBack,42,0,20
	bsKickback.InitExitSnd"popper_ball",""

	Set bsUpKick = New cvpmBallStack : With bsUpKick
		.InitSw 0,46,0,0,0,0,0,0
		.InitSaucer Kickerrighthole,46,180,29
		.KickZ = 1.4
		.InitExitSnd "popper_ball", ""
'		.CreateEvents "bsUpKick", bsUpKick
	End With

	vpmTimer.AddTimer 500,"Balltrap.CreateBall'"
	vpmTimer.AddTimer 500,"Balltrap.Kick 0,1'"

 	bsTrough.addball 0

 	vpmMapLights AllLights

End Sub

Dim droptarget1pos,droptarget2pos,droptarget3pos,droptarget4pos,droptarget5pos
droptarget1pos=0:droptarget2pos=0:droptarget3pos=0:droptarget4pos=0:droptarget5pos=0

'Gottlieb Bone Busters
'added by Inkochnito
Sub editDips
 	Dim vpmDips:Set vpmDips=New cvpmDips
 	With vpmDips
 		.AddForm  700,400,"Bone Busters - DIP switches"
 		.AddFrame 2,4,190,"Maximum credits",49152,Array("8 credits",0,"10 credits",32768,"15 credits",&H00004000,"20 credits",49152)'dip 15&16
 		.AddFrame 2,80,190,"Coin chute 1 and 2 control",&H00002000,Array("seperate",0,"same",&H00002000)'dip 14
 		.AddFrame 2,126,190,"Playfield special",&H00200000,Array("replay",0,"extra ball",&H00200000)'dip 22
 		.AddFrame 2,172,190,"High games to date control",&H00000020,Array("no effect",0,"reset high games 2-5 on power off",&H00000020)'dip 6
 		.AddFrame 2,218,190,"Auto-percentage control",&H00000080,Array("disabled (normal high score mode)",0,"enabled",&H00000080)'dip 8
 		.AddFrame 2,264,190,"Extra ball && special timing",&H40000000,Array("shorter",0,"longer",&H40000000)'dip 31
 		.AddChk 2,316,150,Array("Bonus ball feature",&H80000000)'dip 32
 		.AddFrame 205,4,190,"High game to date awards",&H00C00000,Array("not displayed and no award",0,"displayed and no award",&H00800000,"displayed and 2 replays",&H00400000,"displayed and 3 replays",&H00C00000)'dip 23&24
 		.AddFrame 205,80,190,"Balls per game",&H01000000,Array("5 balls",0,"3 balls",&H01000000)'dip 25
 		.AddFrame 205,126,190,"Replay limit",&H04000000,Array("no limit",0,"one per game",&H04000000)'dip 27
 		.AddFrame 205,172,190,"Novelty",&H08000000,Array("normal",0,"extra ball and replay scores 500K",&H08000000)'dip 28
 		.AddFrame 205,218,190,"Game mode",&H10000000,Array("replay",0,"extra ball",&H10000000)'dip 29
 		.AddFrame 205,264,190,"3rd coin chute credits control",&H20000000,Array("no effect",0,"add 9",&H20000000)'dip 30
 		.AddChk 170,316,100,Array("Match feature",&H02000000)'dip 26
 		.AddChk 290,316,100,Array("Attract sound",&H00000040)'dip 7
 		.AddLabel 50,340,300,20,"After hitting OK, press F3 to reset game with new settings."
 		.ViewDips
 	End With
End Sub
Set vpmShowDips=GetRef("EditDips")

Sub BoneBusters_KeyDown(ByVal KeyCode)
	If KeyCode=LeftMagnaSave Then Controller.Switch(6)=1
	If KeyCode=RightMagnaSave Then Controller.Switch(16)=1
	If KeyCode=LeftFlipperKey Then
		Controller.Switch(53)=1
	End If
	If KeyCode=RightFlipperKey Then
		Controller.Switch(35)=1
		Controller.Switch(45)=1
	End If
	If vpmKeyDown(KeyCode) Then Exit Sub
	If KeyCode=PlungerKey Then Plunger.Pullback:PlaySound "plungerpull",0,1,0.25,0.25
	If keycode = LeftTiltKey Then
		Nudge 90, 2
	End If

	If keycode = RightTiltKey Then
		Nudge 270, 2
	End If

	If keycode = CenterTiltKey Then
		Nudge 0, 2
	End If
End Sub

Sub BoneBusters_KeyUp(ByVal KeyCode)
	If KeyCode=LeftMagnaSave Then Controller.Switch(6)=0
	If KeyCode=RightMagnaSave Then Controller.Switch(16)=0
	If KeyCode=LeftFlipperKey Then
		Controller.Switch(53)=0
	End If
	If KeyCode=RightFlipperKey Then
		Controller.Switch(35)=0
		Controller.Switch(45)=0
	End If
	If vpmKeyUp(KeyCode) Then Exit Sub
	If KeyCode=PlungerKey Then Plunger.Fire:PlaySound "plunger",0,1,0.25,0.25
End Sub

Dim SRelay,OSRelay,LG,RG,SkullRelay,OS2,LT2,OLT2
SRelay=0:LG=0:RG=0:SkullRelay=0:OS2=0:OSRelay=0

Dim LT14,OLT14,LT15,OLT15,LT18,OLT18
LT14=0:OLT14=0:LT15=0:OLT15=0:LT18=0:OLT18=0

dim rA,rB,rC
rA=0:rB=0:rC=0

sub relayA_timer
	rA=rA+1
	select case rA
		case 1: z1.state=1:z2.state=0:z3.state=0:z4.state=0
		case 2: z1.state=0:z2.state=1:z3.state=0:z4.state=0
		case 3: z1.state=0:z2.state=0:z3.state=1:z4.state=0
		case 4: z1.state=0:z2.state=0:z3.state=0:z4.state=1:rA=0
	end select
End sub

sub relayB_timer
	rB=rB+1
	select case rB
		case 1: z5.state=1:z6.state=0:z7.state=0:z8.state=0:z9.state=0
		case 2: z5.state=1:z6.state=1:z7.state=0:z8.state=0:z9.state=0
		case 3: z5.state=1:z6.state=0:z7.state=1:z8.state=0:z9.state=0
		case 4: z5.state=1:z6.state=0:z7.state=0:z8.state=1:z9.state=0
		case 5: z5.state=1:z6.state=0:z7.state=0:z8.state=0:z9.state=1:rB=0
	end select
End sub

sub relayC_timer
	rC=rC+1
	select case rC
		case 1: z10.state=1:z11.state=0
		case 2: z10.state=0:z11.state=1:rC=0
	end select
End sub

Sub sw1_Hit:vpmTimer.PulseSw 1:End Sub					    '1
Sub sw2_Hit:Controller.Switch(2)=1:End Sub					'2
Sub sw2_Unhit:Controller.Switch(2)=0:End Sub
Sub sw3_Hit:Controller.Switch(3)=1:End Sub					'3
Sub sw3_Unhit:Controller.Switch(3)=0:End Sub
Sub sw4_Hit:Controller.Switch(4)=1:End Sub					'4
Sub sw4_Unhit:Controller.Switch(4)=0:End Sub
Sub sw5_Hit:Controller.Switch(5)=1:End Sub					'5
Sub sw5_Unhit:Controller.Switch(5)=0:End Sub

Sub sw12_Hit:Controller.Switch(12)=1:End Sub				'12
Sub sw12_Unhit:Controller.Switch(12)=0:End Sub
Sub sw13_Hit:Controller.Switch(13)=1:End Sub				'13
Sub sw13_Unhit:Controller.Switch(13)=0:End Sub
Sub SpinnerSW14_Spin:vpmTimer.PulseSw 14:End Sub			'14
Sub sw15_Hit:Controller.Switch(15)=1:End Sub				'15
Sub sw15_Unhit:Controller.Switch(15)=0:End Sub

Sub sw20_Hit:dtLow.Hit 3:End Sub							'20

Sub SpinnerSW24_Spin:vpmTimer.PulseSw 24:End Sub			'24
Sub sw25_Hit:Controller.Switch(25)=1:End Sub	'25
Sub sw25_Unhit:Controller.Switch(25)=0:End Sub
Sub KickerTopHole_Hit:bsTopHole.AddBall 0:End Sub			'26

Sub sw32_Hit:Controller.Switch(32)=1:End Sub				'32
Sub sw32_Unhit:Controller.Switch(32)=0:End Sub
Sub sw33_Hit:Controller.Switch(33)=1:End Sub				'33
Sub sw33_unHit:Controller.Switch(33)=0:End Sub
Sub sw34_Hit:Controller.Switch(34)=1:End Sub				'33
Sub sw34_unHit:Controller.Switch(34)=0:End Sub


Sub BottomholeSW36_Hit:bsBottom.AddBall 0:End Sub			'36


Sub KickerBack_Hit:bsKickback.AddBall 0:End Sub				'42
Sub sw43_Hit:Controller.Switch(43)=1:End Sub				'43
Sub sw43_Unhit:Controller.Switch(43)=0:End Sub
Sub sw44_Hit:Controller.Switch(44)=1:End Sub				'44
Sub sw44_Unhit:Controller.Switch(44)=0:End Sub

Sub Kickerrighthole_Hit:bsUpKick.AddBall Me:End Sub			'46
Sub sw50_Hit:vpmTimer.PulseSw 50:End Sub					'50

Sub sw11_Hit
	sw11p.transx = -10
	vpmTimer.PulseSw 11
	Me.TimerEnabled = 1
End Sub
Sub sw11_Timer
	sw11p.transx = 0
	Me.TimerEnabled = 0
End Sub

Sub sw21_Hit
	sw21p.transx = -10
	vpmTimer.PulseSw 21
	Me.TimerEnabled = 1
End Sub
Sub sw21_Timer
	sw21p.transx = 0
	Me.TimerEnabled = 0
End Sub

Sub sw31_Hit
	sw31p.transx = -10
	vpmTimer.PulseSw 31
	Me.TimerEnabled = 1
End Sub
Sub sw31_Timer
	sw31p.transx = 0
	Me.TimerEnabled = 0
End Sub

Sub sw41_Hit
	sw41p.transx = -10
	vpmTimer.PulseSw 41
	Me.TimerEnabled = 1
End Sub
Sub sw41_Timer
	sw41p.transx = 0
	Me.TimerEnabled = 0
End Sub

Sub sw51_Hit
	sw51p.transx = -10
	vpmTimer.PulseSw 51
	Me.TimerEnabled = 1
End Sub
Sub sw51_Timer
	sw51p.transx = 0
	Me.TimerEnabled = 0
End Sub

Sub BumperSW52_Hit												'52
	PlaySound "fx_bumper4"
	vpmTimer.PulseSw 52
	B1L1.State = 1:B1L2. State = 1
	Me.TimerEnabled = 1
End Sub

Sub BumperSW52_Timer
	B1L1.State = 0:B1L2. State = 0
	Me.Timerenabled = 0
End Sub

Sub Drain_Hit									'56
	PlaySound "drain",0,1,0,0.25
	bsTrough.AddBall Me
End Sub

Dim obj
for each obj in Array(FlasherSol1a,FlasherSol1b): obj.visible=0:next
for each obj in Array(FlasherSol2a,FlasherSol2b): obj.visible=0:next
for each obj in Array(FlasherSol3a): obj.visible=0:next
for each obj in Array(FlasherSol5a,FlasherSol5b): obj.visible=0:next
for each obj in Array(FlasherSol6a,FlasherSol6b): obj.visible=0:next

Sub Bot(Enabled)
	If Enabled Then
		If SRelay=0 Then
			If bsBottom.Balls Then bsBottom.ExitSol_On
		Else
			for each obj in Array(FlasherSol1a,FlasherSol1b): obj.visible=1:next
		End If
	Else
		If SRelay=1 Then
			for each obj in Array(FlasherSol1a,FlasherSol1b): obj.visible=0:next
		End If
	End If
End Sub

Sub Holes(Enabled)
	If Enabled Then
		If SRelay=0 Then
			If bsTopHole.Balls Then bsTopHole.ExitSol_On
		Else
			for each obj in Array(FlasherSol2a,FlasherSol2b): obj.visible=1:next
			z100.state=1:z13.state=1
		End If
	Else
		If SRelay=1 Then
			for each obj in Array(FlasherSol2a,FlasherSol2b): obj.visible=0:next
			z100.state=0:z13.state=0
		end if
	End If
End Sub

Sub KickB(Enabled)
	If Enabled Then
		If SRelay=1 Then
			If bsKickback.Balls Then
				bsKickback.ExitSol_On
			Else
			for each obj in Array(FlasherSol3a): obj.visible=1:next
			End If
		End If
	Else
		If SRelay=1 Then
			for each obj in Array(FlasherSol3a): obj.visible=0:next
		end if
	End If
End Sub

Sub UPK(Enabled)
	If Enabled Then
		If SRelay=0 Then
			If bsUpKick.Balls Then bsUpKick.ExitSol_On
		Else
			for each obj in Array(FlasherSol5a,FlasherSol5b): obj.visible=1:next
		End If
	Else
		If SRelay=1 Then for each obj in Array(FlasherSol5a,FlasherSol5b): obj.visible=0:next
	End If
End Sub

Sub TBank(Enabled)
	If Enabled Then
		If SRelay=0 Then
			dtLow.DropSol_On
			if droptarget1pos=48 then rdt1.enabled=1
			if droptarget2pos=48 then rdt2.enabled=1
			if droptarget3pos=48 then rdt3.enabled=1
			if droptarget4pos=48 then rdt4.enabled=1
			if droptarget5pos=48 then rdt5.enabled=1
		Else
			for each obj in Array(FlasherSol6a,FlasherSol6b): obj.visible=1:next
			z12.state=1:z14.state=1:z15.state=1
		End If
	Else
		If SRelay=1 Then for each obj in Array(FlasherSol6a,FlasherSol6b): obj.visible=0:next
		z12.state=0:z14.state=0:z15.state=0
	End If
End Sub

Sub droptarget1a_Hit:dtLow.Hit 1:me.timerenabled=1:End Sub
Sub droptarget2a_Hit:dtLow.Hit 2:me.timerenabled=1:End Sub
Sub droptarget3a_Hit:dtLow.Hit 3:me.timerenabled=1:End Sub
Sub droptarget4a_Hit:dtLow.Hit 4:me.timerenabled=1:End Sub
Sub droptarget5a_Hit:dtLow.Hit 5:me.timerenabled=1:End Sub

Sub droptarget1a_timer()
	droptarget1pos=droptarget1pos+4
	droptarget1b.TransZ=0-droptarget1pos
	if droptarget1pos=48 then me.timerenabled=0
End Sub
Sub rdt1_timer()
	droptarget1pos=droptarget1pos-12
	droptarget1b.TransZ=0-droptarget1pos
	if droptarget1pos=0 then me.enabled=0
End Sub

Sub droptarget2a_timer()
	droptarget2pos=droptarget2pos+4
	droptarget2b.TransZ=0-droptarget2pos
	if droptarget2pos=48 then me.timerenabled=0
End Sub
Sub rdt2_timer()
	droptarget2pos=droptarget2pos-12
	droptarget2b.TransZ=0-droptarget2pos
	if droptarget2pos=0 then me.enabled=0
End Sub

Sub droptarget3a_timer()
	droptarget3pos=droptarget3pos+4
	droptarget3b.TransZ=0-droptarget3pos
	if droptarget3pos=48 then me.timerenabled=0
End Sub
Sub rdt3_timer()
	droptarget3pos=droptarget3pos-12
	droptarget3b.TransZ=0-droptarget3pos
	if droptarget3pos=0 then me.enabled=0
End Sub

Sub droptarget4a_timer()
	droptarget4pos=droptarget4pos+4
	droptarget4b.TransZ=0-droptarget4pos
	if droptarget4pos=48 then me.timerenabled=0
End Sub
Sub rdt4_timer()
	droptarget4pos=droptarget4pos-12
	droptarget4b.TransZ=0-droptarget4pos
	if droptarget4pos=0 then me.enabled=0
End Sub
Sub droptarget5a_timer()
	droptarget5pos=droptarget5pos+4
	droptarget5b.TransZ=0-droptarget5pos
	if droptarget5pos=48 then me.timerenabled=0
End Sub
Sub rdt5_timer()
	droptarget5pos=droptarget5pos-12
	droptarget5b.TransZ=0-droptarget5pos
	if droptarget5pos=0 then me.enabled=0
End Sub

'*****GI Lights On
dim xx
For each xx in GI:xx.State = 1: Next

Sub RightSlingShot_Slingshot
    PlaySound "left_slingshot", 0, 1, 0.05, 0.05
	vpmTimer.PulseSw 23
End Sub


Sub LeftSlingShot_Slingshot
    PlaySound "right_slingshot",0,1,-0.05,0.05
	vpmTimer.PulseSw 22
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
Sub LeftFlipper2_Collide(parm)
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

sub BoneBusters_Exit
	Controller.Stop
end sub

Sub FlipperDecals_Timer
	LFD.RotAndTra2=LeftFlipper.CurrentAngle
	RFD.RotAndTra2=RightFlipper.CurrentAngle
	ULFD.RotAndTra2=Leftflipper2.CurrentAngle
End Sub

Sub DisplayTimer_Timer
	OSRelay=L13.State
	If OSRelay<>SRelay Then SRelay=OSRelay
	LT2=L2.State
	If LT2<>OLT2 Then
		If LT2=1 Then
			bsTrough.ExitSol_On
		End If
	End If
	OLT2=LT2
	LG=L16.State
	RG=L17.State
	SkullRelay=L19.State'Skull JAW
	If LG=1 Then
		LeftGate.RotateToEnd
		If bsKickback.Balls Then bsKickback.ExitSol_On
		Else
		LeftGate.RotateToStart
	End If
	If RG=1 Then
		RightGate.RotateToEnd
		Else
		RightGate.RotateToStart
	End If
	LT14=L14.State
	If LT14<>OLT14 Then
		If LT14=1 Then
			relayA.enabled=1
		Else
			relayA.enabled=0:rA=0:z1.state=0:z2.state=0:z3.state=0:z4.state=0
		End If
	End If
	OLT14=LT14
	LT15=L15.State
	If LT15<>OLT15 Then
		If LT15=1 Then
			relayB.enabled=1
		Else
			relayB.enabled=0:rB=0:z5.state=0:z6.state=0:z7.state=0:z8.state=0:z9.state=0
		End If
	End If
	OLT15=LT15
	LT18=L18.State
	If LT18<>OLT18 Then
		If LT18=1 Then
			relayC.enabled=1
		Else
			relayC.enabled=0:rC=0:z10.state=0:z11.state=0
		End If
	End If
	OLT18=LT18

	Dim ChgLED,II,Num,Chg,Stat,Obj
	ChgLED=Controller.ChangedLEDs(&Hffffffff,&Hffffffff)
	If Not IsEmpty(ChgLED) Then
		For II=0 To UBound(ChgLED)
			Num=ChgLED(II,0):Chg=ChgLED(II,1):Stat=ChgLED(II,2)
			For Each Obj In Digits(Num)
				If Chg And 1 Then Obj.State=Stat And 1
				Chg=Chg\2:Stat=Stat\2
			Next
		Next
	End If
End Sub

Dim Digits(39)
Digits(0)=Array(Light1,Light2,Light3,Light4,Light5,Light6,Light7,Light8,Light9,Light10,Light11,Light12,Light13,Light14,Light15)
Digits(1)=Array(Light17,Light18,Light19,Light20,Light21,Light22,Light23,Light24,Light25,Light26,Light27,Light28,Light29,Light30,Light31)
Digits(2)=Array(Light33,Light34,Light35,Light36,Light37,Light38,Light39,Light40,Light41,Light42,Light43,Light44,Light45,Light46,Light47)
Digits(3)=Array(Light49,Light50,Light51,Light52,Light53,Light54,Light55,Light56,Light57,Light58,Light59,Light60,Light61,Light62,Light63)
Digits(4)=Array(Light65,Light66,Light67,Light68,Light69,Light70,Light71,Light72,Light73,Light74,Light75,Light76,Light77,Light78,Light79)
Digits(5)=Array(Light81,Light82,Light83,Light84,Light85,Light86,Light87,Light88,Light89,Light90,Light91,Light92,Light93,Light94,Light95)
Digits(6)=Array(Light97,Light98,Light99,Light100,Light101,Light102,Light103,Light104,Light105,Light106,Light107,Light108,Light109,Light110,Light111)
Digits(7)=Array(Light113,Light114,Light115,Light116,Light117,Light118,Light119,Light120,Light121,Light122,Light123,Light124,Light125,Light126,Light127)
Digits(8)=Array(Light129,Light130,Light131,Light132,Light133,Light134,Light135,Light136,Light137,Light138,Light139,Light140,Light141,Light142,Light143)
Digits(9)=Array(Light145,Light146,Light147,Light148,Light149,Light150,Light151,Light152,Light153,Light154,Light155,Light156,Light157,Light158,Light159)
Digits(10)=Array(Light161,Light162,Light163,Light164,Light165,Light166,Light167,Light168,Light169,Light170,Light171,Light172,Light173,Light174,Light175)
Digits(11)=Array(Light177,Light178,Light179,Light180,Light181,Light182,Light183,Light184,Light185,Light186,Light187,Light188,Light189,Light190,Light191)
Digits(12)=Array(Light193,Light194,Light195,Light196,Light197,Light198,Light199,Light200,Light201,Light202,Light203,Light204,Light205,Light206,Light207)
Digits(13)=Array(Light209,Light210,Light211,Light212,Light213,Light214,Light215,Light216,Light217,Light218,Light219,Light220,Light221,Light222,Light223)
Digits(14)=Array(Light225,Light226,Light227,Light228,Light229,Light230,Light231,Light232,Light233,Light234,Light235,Light236,Light237,Light238,Light239)
Digits(15)=Array(Light241,Light242,Light243,Light244,Light245,Light246,Light247,Light248,Light249,Light250,Light251,Light252,Light253,Light254,Light255)
Digits(16)=Array(Light257,Light258,Light259,Light260,Light261,Light262,Light263,Light264,Light265,Light266,Light267,Light268,Light269,Light270,Light271)
Digits(17)=Array(Light273,Light274,Light275,Light276,Light277,Light278,Light279,Light280,Light281,Light282,Light283,Light284,Light285,Light286,Light287)
Digits(18)=Array(Light289,Light290,Light291,Light292,Light293,Light294,Light295,Light296,Light297,Light298,Light299,Light300,Light301,Light302,Light303)
Digits(19)=Array(Light305,Light306,Light307,Light308,Light309,Light310,Light311,Light312,Light313,Light314,Light315,Light316,Light317,Light318,Light319)
Digits(20)=Array(Light321,Light322,Light323,Light324,Light325,Light326,Light327,Light328,Light329,Light330,Light331,Light332,Light333,Light334,Light335)
Digits(21)=Array(Light337,Light338,Light339,Light340,Light341,Light342,Light343,Light344,Light345,Light346,Light347,Light348,Light349,Light350,Light351)
Digits(22)=Array(Light353,Light354,Light355,Light356,Light357,Light358,Light359,Light360,Light361,Light362,Light363,Light364,Light365,Light366,Light367)
Digits(23)=Array(Light369,Light370,Light371,Light372,Light373,Light374,Light375,Light376,Light377,Light378,Light379,Light380,Light381,Light382,Light383)
Digits(24)=Array(Light385,Light386,Light387,Light388,Light389,Light390,Light391,Light392,Light393,Light394,Light395,Light396,Light397,Light398,Light399)
Digits(25)=Array(Light401,Light402,Light403,Light404,Light405,Light406,Light407,Light408,Light409,Light410,Light411,Light412,Light413,Light414,Light415)
Digits(26)=Array(Light417,Light418,Light419,Light420,Light421,Light422,Light423,Light424,Light425,Light426,Light427,Light428,Light429,Light430,Light431)
Digits(27)=Array(Light433,Light434,Light435,Light436,Light437,Light438,Light439,Light440,Light441,Light442,Light443,Light444,Light445,Light446,Light447)
Digits(28)=Array(Light449,Light450,Light451,Light452,Light453,Light454,Light455,Light456,Light457,Light458,Light459,Light460,Light461,Light462,Light463)
Digits(29)=Array(Light465,Light466,Light467,Light468,Light469,Light470,Light471,Light472,Light473,Light474,Light475,Light476,Light477,Light478,Light479)
Digits(30)=Array(Light481,Light482,Light483,Light484,Light485,Light486,Light487,Light488,Light489,Light490,Light491,Light492,Light493,Light494,Light495)
Digits(31)=Array(Light497,Light498,Light499,Light500,Light501,Light502,Light503,Light504,Light505,Light506,Light507,Light508,Light509,Light510,Light511)
Digits(32)=Array(Light513,Light514,Light515,Light516,Light517,Light518,Light519,Light520,Light521,Light522,Light523,Light524,Light525,Light526,Light527)
Digits(33)=Array(Light529,Light530,Light531,Light532,Light533,Light534,Light535,Light536,Light537,Light538,Light539,Light540,Light541,Light542,Light543)
Digits(34)=Array(Light545,Light546,Light547,Light548,Light549,Light550,Light551,Light552,Light553,Light554,Light555,Light556,Light557,Light558,Light559)
Digits(35)=Array(Light561,Light562,Light563,Light564,Light565,Light566,Light567,Light568,Light569,Light570,Light571,Light572,Light573,Light574,Light575)
Digits(36)=Array(Light577,Light578,Light579,Light580,Light581,Light582,Light583,Light584,Light585,Light586,Light587,Light588,Light589,Light590,Light591)
Digits(37)=Array(Light593,Light594,Light595,Light596,Light597,Light598,Light599,Light600,Light601,Light602,Light603,Light604,Light605,Light606,Light607)
Digits(38)=Array(Light609,Light610,Light611,Light612,Light613,Light614,Light615,Light616,Light617,Light618,Light619,Light620,Light621,Light622,Light623)
Digits(39)=Array(Light625,Light626,Light627,Light628,Light629,Light630,Light631,Light632,Light633,Light634,Light635,Light636,Light637,Light638,Light639)

if showdt=false then
	for each obj in backdroplights
	obj.visible=0
	next
end if

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

Function AudioFade(tableobj) ' Fades between front and back of the table (for surround systems or 2x2 speakers, etc), depending on the Y position on the table. "BoneBusters" is the name of the table
  Dim tmp
  tmp = tableobj.y * 2 / BoneBusters.height-1
  If tmp > 0 Then
    AudioFade = Csng(tmp ^10)
  Else
    AudioFade = Csng(-((- tmp) ^10) )
  End If
End Function

Function AudioPan(tableobj) ' Calculates the pan for a tableobj based on the X position on the table. "BoneBusters" is the name of the table
  Dim tmp
  tmp = tableobj.x * 2 / BoneBusters.width-1
  If tmp > 0 Then
    AudioPan = Csng(tmp ^10)
  Else
    AudioPan = Csng(-((- tmp) ^10) )
  End If
End Function

Function Pan(ball) ' Calculates the pan for a ball based on the X position on the table. "BoneBusters" is the name of the table
    Dim tmp
    tmp = ball.x * 2 / BoneBusters.width-1
    If tmp > 0 Then
        Pan = Csng(tmp ^10)
    Else
        Pan = Csng(-((- tmp) ^10) )
    End If
End Function

Function AudioFade(ball) ' Can this be together with the above function ?
  Dim tmp
  tmp = ball.y * 2 / BoneBusters.height-1
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

Const tnob = 8 ' total number of balls
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
  If BoneBusters.VersionMinor > 3 OR BoneBusters.VersionMajor > 10 Then
    PlaySound("fx_collide"), 0, Csng(velocity) ^2 / 200, Pan(ball1), 0, Pitch(ball1), 0, 0, AudioFade(ball1)
  Else
    PlaySound("fx_collide"), 0, Csng(velocity) ^2 / 200, Pan(ball1), 0, Pitch(ball1), 0, 0
  End if
End Sub

'	---------------------------------------------
'	HERCULES
'	---------------------------------------------

Option Explicit
    Randomize
	'StartShake

' Thalamus 2018-07-23
' Added/Updated "Positional Sound Playback Functions" and "Supporting Ball & Sound Functions"
' Changed UseSolenoids=1 to 2
' No special SSF tweaks yet.
' Replaced the whole Ballrolling routine


	Const BallSize = 100 'Its huge!

Dim VarHidden, UseVPMDMD
If Hercules.ShowDT = true then
	UseVPMDMD = True
	VarHidden = 0
	lockdown.visible = true
	Primitive8.visible = true
	leftrail.visible = true
	rightrail.visible = true
else
	UseVPMDMD = False
	VarHidden = 1
	lockdown.visible = false
	leftrail.visible = false
	rightrail.visible = false
	Primitive8.visible = false

end if



LoadVPM "01120000", "ATARI2.VBS", 3.0

Sub LoadVPM(VPMver, VBSfile, VBSver)
	On Error Resume Next
		If ScriptEngineMajorVersion < 5 Then MsgBox "VB Script Engine 5.0 or higher required"
		If VPBuildVersion < 0 Or Err Then
			ExecuteGlobal CreateObject("Scripting.FileSystemObject").OpenTextFile(VBSfile, 1).ReadAll
		Else
			ExecuteGlobal GetTextFile(VBSfile)
		End If
		If Err Then MsgBox "Unable to open " & VBSfile & ". Ensure that it is in the same folder as this table. " & vbNewLine & Err.Description : Err.Clear
		'Set Controller = CreateObject("VPinMAME.Controller")
		Set Controller = CreateObject("b2s.server")
		If Err Then MsgBox "Can't Load VPinMAME." & vbNewLine & Err.Description
		If VPMver>"" Then If Controller.Version < VPMver Or Err Then MsgBox "VPinMAME ver " & VPMver & " required." : Err.Clear
		If VPinMAMEDriverVer < VBSver Or Err Then MsgBox VBSFile & " ver " & VBSver & " or higher required."
End Sub






'==================================================================
' Game Specific code starts here
'==================================================================
Const cGameName     = "hercules"   ' PinMAME short name
Const cCredits      = "Hercules Total rebuild in vp10 by randr"
Const UseSolenoids  = 2
Const UseLamps      = True
Const UseGI         = False

' Standard Sounds
Const SSolenoidOn   = ""
Const SSolenoidOff  = ""
Const SFlipperOn    = "flipup"
Const SFlipperOff   = "flipdown"
Const SCoin         = "coin3"

' Solenoids
Const sKnocker		=  3
Const sRJet			=  4
Const sLJet			=  5
Const sRSling		=  8
Const sLSling		=  9
Const sKnocker2		= 10
Const sKnocker3		= 11
Const sKnocker4		= 12
Const sKnocker5		= 13
Const sBallRelease	= 15
Const sEnable		= 16
Const sCLo			= 17

'----------------------------------------------------
' Bind events to solenoids
' Last argument will always be "enabled" (True, False)
'----------------------------------------------------
SolCallback(sBallRelease)	= "bsTrough.SolOut"
SolCallback(sKnocker)		= "vpmSolSound ""knock"","
SolCallback(sLSling)		= "vpmSolSound ""hercsling"","
SolCallback(sRSling)		= "vpmSolSound ""hercsling"","
SolCallback(sLJet)			= "Bumper 1,"
SolCallback(sRJet)			= "Bumper 2,"
SolCallback(sEnable)		= "vpmNudge.SolGameOn"
SolCallback(sLLFlipper)		= "vpmSolFlipper LeftFlipper, Nothing,"
SolCallback(sLRFlipper)		= "vpmSolFlipper RightFlipper, Nothing,"

Sub Bumper(number, enabled)
	vpmSolSound "hercjet", enabled
	if enabled then BumperState(number) = 3
End Sub

'--------------------------------
' Init the table, Start VPinMAME
'--------------------------------

Dim AttractCount
Dim Cue
Dim BallActive
'Dim BallVel(4)
Dim BumperState(2)
Dim bsTrough
dim bump1
dim bump2

Sub Hercules_Init
	vpmInit Me
	On Error Resume Next
	With Controller
		.GameName = cGameName
		.SplashInfoLine = cCredits
		.HandleKeyboard = 0
		.ShowTitle = 0
		.ShowDMDOnly = 1
		.ShowFrame = 0
		'.HandleMechanics = 1
		.Hidden = VarHidden
		On Error Resume Next
			.Run
			If Err Then MsgBox Err.Description
		On Error Goto 0
	End With
	' Main Timer init
	PinMAMETimer.Interval = PinMAMEInterval
	PinMAMETimer.Enabled = True
	' Nudging
	vpmNudge.TiltSwitch = 48
	vpmNudge.Sensitivity = 3
	vpmNudge.TiltObj = Array(Bumper1,Bumper2,LeftSlingshot,RightSlingshot)

	Set bsTrough = New cvpmBallStack ' Trough handler
	bsTrough.InitSw 0,1,0,0,0,0,0,0
	bsTrough.InitKick Kicker1, 90, 4
	bsTrough.InitExitSnd "BallOut", "BallOut"
	bsTrough.BallImage = "Cue"
	bsTrough.Balls = 1

	Dim f
	For f=1 to 6

	Next
End Sub




'-------------------
' keyboard handlers
'-------------------
Sub Hercules_KeyUp(ByVal keycode)
	If vpmKeyUp(keycode) Then Exit Sub
	If keycode = PlungerKey Then PlaySound "Plunger" : Plunger.Fire
End Sub

Sub Hercules_KeyDown(ByVal keycode)
	If vpmKeyDown(keycode) Then Exit Sub
	If keycode = PlungerKey Then Plunger.Pullback
End Sub


Sub FlipperDecals_Timer
	LFD.RotAndTra2=LeftFlipper.CurrentAngle
	RFD.RotAndTra2=RightFlipper.CurrentAngle

End Sub

Sub timergate_Timer
	prightgate.rotz=gate3.CurrentAngle
	pleftgate.rotz=gate1.CurrentAngle

End Sub

Sub Bumper1_Hit:bump1 = 1:Me.TimerEnabled = 1:vpmTimer.PulseSw 53::End Sub

Sub Bumper1_Timer()
           Select Case bump1
               Case 1:Pring1.z = 3:bump1 = 2
               Case 2:Pring1.z = -3:bump1 = 3
               Case 3:Pring1.z = -13:bump1 = 4
               Case 4:Pring1.z = -13:bump1 = 5
               Case 5:Pring1.z = -3:bump1 = 6
               Case 6:Pring1.z = 3:bump1 = 7
               case 7:Me.TimerEnabled = 0
           End Select
       End Sub

Sub Bumper2_Hit:vpmTimer.PulseSw 52:bump2 = 1:Me.TimerEnabled = 1:End Sub

Sub Bumper2_Timer()
           Select Case bump2
               Case 1:Pring2.z = 3:bump2 = 2
               Case 2:Pring2.z = -3:bump2 = 3
               Case 3:Pring2.z = -13:bump2 = 4
               Case 4:Pring2.z = -13:bump2 = 5
               Case 5:Pring2.z = -3:bump2 = 6
               Case 6:Pring2.z = 3:bump2 = 7
               case 7:Me.TimerEnabled = 0
           End Select
       End Sub

'---------------------------------------------------------------
' SWITCH HANDLING
'---------------------------------------------------------------
Sub Drain_Hit : PlaySound "Drain" : bsTrough.AddBall Me : BallActive = 0 : End Sub
Sub Roll3_Hit : Controller.Switch(12) = True : End Sub
Sub Roll3_Unhit : Controller.Switch(12) = False : End Sub

Sub Rubber7_Slingshot : vpmTimer.pulseSw 13 : End Sub
Sub Rubber12_Slingshot : vpmTimer.pulseSw 13 : End Sub
Sub Roll1_Hit : Controller.Switch(13) = True : End Sub
Sub Roll1_Unhit : Controller.Switch(13) = False : End Sub
Sub Roll4_Hit : Controller.Switch(13) = True : End Sub
Sub Roll4_Unhit : Controller.Switch(13) = False : End Sub
Sub Roll5_Hit : Controller.Switch(13) = True : End Sub
Sub Roll5_Unhit : Controller.Switch(13) = False : End Sub

Sub Roll6_Hit : Controller.Switch(18) = True : End Sub
Sub Roll6_Unhit : Controller.Switch(18) = False : End Sub
Sub Roll7_Hit : Controller.Switch(19) = True : End Sub
Sub Roll7_Unhit : Controller.Switch(19) = False : End Sub
Sub Roll8_Hit : Controller.Switch(20) = True : End Sub
Sub Roll8_Unhit : Controller.Switch(20) = False : End Sub
Sub Lane1_Hit : Controller.Switch(21) = True : End Sub
Sub Lane1_Unhit : Controller.Switch(21) = False : End Sub
Sub Lane2_Hit : Controller.Switch(24) = True : End Sub
Sub Lane2_Unhit : Controller.Switch(24) = False : End Sub
Sub Lane3_Hit : Controller.Switch(25) = True : End Sub
Sub Lane3_Unhit : Controller.Switch(25) = False : End Sub

Sub Roll2_Hit : Controller.Switch(26) = True : End Sub
Sub Roll2_Unhit : Controller.Switch(26) = False : End Sub
Sub Roll9_Hit : Controller.Switch(26) = True : End Sub
Sub Roll9_Unhit : Controller.Switch(26) = False : End Sub

Sub LeftOutLane_Hit : Controller.Switch(28) = True : End Sub
Sub LeftOutLane_Unhit : Controller.Switch(28) = False : End Sub
Sub LeftInLane_Hit : Controller.Switch(29) = True : End Sub
Sub LeftInLane_Unhit : Controller.Switch(29) = False : End Sub
Sub RightInLane_Hit : Controller.Switch(32) = True : End Sub
Sub RightInLane_Unhit : Controller.Switch(32) = False : End Sub
Sub RightOutLane_Hit : Controller.Switch(33) = True : End Sub
Sub RightOutLane_Unhit : Controller.Switch(33) = False : End Sub
Sub Spinner1_Spin :	vpmTimer.pulseSw 40 : End Sub
'Sub Bumper2_Hit : vpmTimer.PulseSw 52 : End Sub
'Sub Bumper1_Hit : vpmTimer.PulseSw 53 : End Sub
Sub RightSlingShot_Slingshot  : vpmTimer.PulseSw 56 : End Sub
Sub LeftSlingShot_Slingshot : vpmTimer.PulseSw 57 : End Sub

Sub Gate1_Hit : PlaySound "Gate" : End Sub
Sub Gate2_Hit : PlaySound "Gate" : End Sub
Sub Gate3_Hit : PlaySound "Gate" : End Sub
Sub WoodHit1_Hit : PlaySound "Bump3" : End Sub
Sub WoodHit2_Hit : PlaySound "Bump3" : End Sub
Sub WoodHit3_Hit : PlaySound "Bump3" : End Sub
Sub WoodHit4_Hit : PlaySound "metalhit_thin" : End Sub

Sub ActiveOn_Hit()
	Set Cue = ActiveBall
	BallActive = 1
End Sub

'-------------------------------------
' Map lights into array
'-------------------------------------
Set Lights( 1) = Bonus1K
Set Lights( 2) = Bonus2K
Set Lights( 3) = Bonus3K
Set Lights( 4) = Bonus4K
Set Lights( 5) = Bonus5K
Set Lights( 6) = Bonus6K
Set Lights( 7) = Bonus7K
Set Lights( 8) = Bonus8K
Set Lights( 9) = Bonus9K
Set Lights(10) = Bonus10K
Set Lights(11) = Bonus20K
Set Lights(12) = Bonus30K
Set Lights(13) = Bonus40K
Set Lights(14) = Yellow1
Set Lights(15) = Yellow2
Set Lights(16) = Yellow3
Set Lights(17) = Yellow4
Set Lights(18) = TopYellow
Set Lights(19) = Orange1
Set Lights(20) = Orange2
Set Lights(21) = Orange3
Set Lights(22) = Orange4
Set Lights(23) = Orange5
Set Lights(24) = TopOrange
Set Lights(25) = Red1
Set Lights(26) = Red2
Set Lights(27) = Red3
Set Lights(28) = Red4
Set Lights(29) = Red5
Set Lights(30) = Red6
Set Lights(31) = TopRed
Set Lights(32) = LeftSP
Set Lights(33) = LeftXB
Set Lights(34) = RightXB
Set Lights(35) = RightSP
Set Lights(36) = TargetLight1
Set Lights(37) = TargetLight2
Set Lights(38) = TargetLight3
Set Lights(39) = TargetLight4
Set Lights(42) = SpinRed
Set Lights(44) = ShootAgainLight
Set Lights(45) = SideYellow
Set Lights(46) = SideOrange
Set Lights(47) = SideRed
Set Lights(49) = SpinYellow
Set Lights(50) = SpinOrange

Sub editDips
	With vpmDips
		.AddForm  334,215,"Hercules DIP switches"
		.AddChk   194,146,150,Array("Match feature",&H00000080)
		.AddChk   194,162,150,Array("High Score Million Limit",&H00400000)
		.AddChk   194,178,150,Array("High Score Display",&H00800000)
		.AddChk   194,194,150,Array("Feature ladder memory",&H20000000)
		.AddFrame   2,  2, 90,"High score award",&HC0000000,Array("3 replays",&H40000000,"2 replays",&HC0000000,"1 replay",0,"Nothing",&H80000000)
		.AddFrame   2, 77, 90,"Maximum credits",&H00000007,Array("5 credits",7,"10 credits",6,"15 credits",5,"20 credits",4,"25 credits",3,"30 credits",2,"35 credits",1,"40 credits",0)
		.AddFrame 107,  2, 90,"Special award",&H00000050,Array("Replay",&H50,"Add a ball",&H40,"50,000 points",&H10,"60,000 points",0)
		.AddFrame 107, 82, 90,"Extra award",&H00006000,Array("Extra ball",&H6000,"20,000 points",&H4000,"30,000 points",0)
		.AddFrame 107,147, 72,"Return lanes",&H00208000,Array("Always off",0,"Alternating",&H200000,"Both on",&H00208000)
		.AddFrame 212,  2,100,"Upper lanes",&H00001000,Array("Start + advance",&H1000,"Start only",0)
		.AddFrame 212, 49,100,"Game mode",&H00000020,Array("Coin-op",&H20,"Free play",0)
		.AddFrame 212, 96,100,"Balls in play",&H1F1F0008,Array("3 balls",&H1D150008,"5 balls",&H1D150000)
		.ViewDips
	End With
End Sub
Set vpmDips = New cvpmDips
Set vpmShowDips = GetRef("editDips")

Sub GameTimer_Timer()

End Sub

''''''''''''''''''''''''''''''''''''''''''''
'''''''''Targets''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''

Dim TargetTR, TargetTL, TargetML, TargetMR

Sub S_Top_P_Hit:TargetTL = 1:Me.TimerEnabled = 1:vpmTimer.PulseSw(35):Playsound("Bump3"):End Sub

       Sub S_Top_P_Timer()
           Select Case TargetTL
               Case 1:Target2.Transy = -5:TargetTL = 2
               Case 2:TargetTL = 3
               Case 3:Target2.Transy = 0:TargetTL = 4
               Case 4:Me.TimerEnabled = 0
           End Select
       End Sub

Sub S_L_Target_Hit:TargetML = 1:Me.TimerEnabled = 1:vpmTimer.PulseSw(34):Playsound("Bump3"):End Sub

       Sub S_L_Target_Timer()
           Select Case TargetML
               Case 1:Target1.Transy = -5:TargetML = 2
               Case 2:TargetML = 3
               Case 3:Target1.Transy = 0:TargetML = 4
               Case 4:Me.TimerEnabled = 0
           End Select
       End Sub

Sub S_R_Target_Hit:TargetMR = 1:Me.TimerEnabled = 1:vpmTimer.PulseSw(36):Playsound("Bump3"):End Sub

       Sub S_R_Target_Timer()
           Select Case TargetMR
               Case 1:Target3.Transy = -5:TargetMR = 2
               Case 2:TargetMR = 3
               Case 3:Target3.Transy = 0:TargetMR = 4
               Case 4:Me.TimerEnabled = 0
           End Select
       End Sub

Sub S_Top_Y_Hit:TargetTR = 1:Me.TimerEnabled = 1:vpmTimer.PulseSw(37):Playsound("Bump3"):End Sub

       Sub S_Top_Y_Timer()
           Select Case TargetTR
               Case 1:Target4.Transy = -5:TargetTR = 2
               Case 2:TargetTR = 3
               Case 3:Target4.Transy = 0:TargetTR = 4
               Case 4:Me.TimerEnabled = 0
           End Select
       End Sub

'**********Sling Shot Animations
' Rstep and Lstep  are the variables that increment the animation
'****************
Dim RStep, Lstep

Sub RightSlingShot_Slingshot
vpmTimer.PulseSw 56
    PlaySound "hercsling", 0, 1, 0.05, 0.05
    RSling.Visible = 0
    RSling1.Visible = 1
    sling1.TransZ = -20
    RStep = 0
    RightSlingShot.TimerEnabled = 1
End Sub

Sub RightSlingShot_Timer

    Select Case RStep
        Case 3:RSLing1.Visible = 0:RSLing2.Visible = 1:sling1.TransZ = -10
        Case 4:RSLing2.Visible = 0:RSLing.Visible = 1:sling1.TransZ = 0':RightSlingShot.TimerEnabled' = 0:gi1.State = 1:Gi2.State = 1
    End Select
    RStep = RStep + 1
End Sub

Sub LeftSlingShot_Slingshot
vpmTimer.PulseSw 57
    PlaySound "hercsling",0,1,-0.05,0.05
    LSling.Visible = 0
    LSling1.Visible = 1
    sling2.TransZ = -20
    LStep = 0
    LeftSlingShot.TimerEnabled = 1
End Sub

Sub LeftSlingShot_Timer

    Select Case LStep
        Case 3:LSLing1.Visible = 0:LSLing2.Visible = 1:sling2.TransZ = -10
        Case 4:LSLing2.Visible = 0:LSLing.Visible = 1:sling2.TransZ = 0':LeftSlingShot.TimerEnabled' = 0:gi3.State = 1:Gi4.State = 1
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
	vpmTimer.pulseSw 40
End Sub

Sub Rubbers_Hit(idx)
 	dim finalspeed
  	finalspeed=SQR(activeball.velx * activeball.velx + activeball.vely * activeball.vely)
 	If finalspeed > 1 then
		PlaySound "fx_rubber2", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
	End if
	If finalspeed >= 2 AND finalspeed <= 1 then
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
		Case 1 : PlaySound "fx_rubber2", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
		Case 2 : PlaySound "fx_rubber2", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
		Case 3 : PlaySound "fx_rubber2", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
	End Select
End Sub

Sub AttractTimer_Timer()
	AttractCount = AttractCount + 1
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

Function AudioFade(tableobj) ' Fades between front and back of the table (for surround systems or 2x2 speakers, etc), depending on the Y position on the table. "Hercules" is the name of the table
  Dim tmp
  tmp = tableobj.y * 2 / Hercules.height-1
  If tmp > 0 Then
    AudioFade = Csng(tmp ^10)
  Else
    AudioFade = Csng(-((- tmp) ^10) )
  End If
End Function

Function AudioPan(tableobj) ' Calculates the pan for a tableobj based on the X position on the table. "Hercules" is the name of the table
  Dim tmp
  tmp = tableobj.x * 2 / Hercules.width-1
  If tmp > 0 Then
    AudioPan = Csng(tmp ^10)
  Else
    AudioPan = Csng(-((- tmp) ^10) )
  End If
End Function

Function Pan(ball) ' Calculates the pan for a ball based on the X position on the table. "Hercules" is the name of the table
    Dim tmp
    tmp = ball.x * 2 / Hercules.width-1
    If tmp > 0 Then
        Pan = Csng(tmp ^10)
    Else
        Pan = Csng(-((- tmp) ^10) )
    End If
End Function

Function AudioFade(ball) ' Can this be together with the above function ?
  Dim tmp
  tmp = ball.y * 2 / Hercules.height-1
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
  If Hercules.VersionMinor > 3 OR Hercules.VersionMajor > 10 Then
    PlaySound("fx_collide"), 0, Csng(velocity) ^2 / 200, Pan(ball1), 0, Pitch(ball1), 0, 0, AudioFade(ball1)
  Else
    PlaySound("fx_collide"), 0, Csng(velocity) ^2 / 200, Pan(ball1), 0, Pitch(ball1), 0, 0
  End if
End Sub

' Thalamus : Exit in a clean and proper way
Sub Hercules_exit()
  Controller.Pause = False
  Controller.Stop
End Sub


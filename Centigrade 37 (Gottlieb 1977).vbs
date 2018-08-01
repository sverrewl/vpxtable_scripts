'############ Centigrade 37 by Gottlieb (1977) ############

'               Original Design by:	Allen Edwall
'				Original Art by:	Gordon Morison
'
'		   
'		   VPX version by bord
'		   Script by Borgdog
'		   Based on VP9 Code by Pinuck
'          Arngrim for the DOF

' Thalamus 2018-07-20
' Added/Updated "Positional Sound Playback Functions" and "Supporting Ball & Sound Functions"
' No special SSF tweaks yet.

Option Explicit
Randomize

On Error Resume Next
ExecuteGlobal GetTextFile("controller.vbs")
If Err Then MsgBox "You need the controller.vbs in order to run this table."
On Error Goto 0

Const cGameName = "centigrade"

'  DECLARATIONS
Const FreeLevel1=90000
Const FreeLevel2=130000
Const FreeLevel3=180000

Const SMFadeResetValue = 18
Const maxCredits = 9
Const TiltMax = 3 'nudges till tilt
Const BallSize = 50

'**********************************************************
'********   	OPTIONS		*******************************
'**********************************************************

Dim BallShadows: Ballshadows=1  		'******************set to 1 to turn on Ball shadows
Dim FlipperShadows: FlipperShadows=1  '***********set to 1 to turn on Flipper shadows

Dim Score, Credits, hisc
Dim BallsRemaining, GameStarted, Match, CurrentBall, SpecialLit, BonusValue
Dim BallsPerGame
Dim Dropped(4)
Dim DropDone
Dim Drop(4)
Set Drop(1)=target1
Set Drop(2)=target2
Set Drop(3)=target3
Set Drop(4)=target4
Dim x, i, j, k, obj
Dim MazeState
Dim MotorRunning
Dim Temperature
Dim Score100k, Score10k, ScoreK, Score100, Score10, Score1, Scorex	'Define 5 different score values for each reel to use
Dim Free1, Free2, Free3 'free game levels reached
Dim Tilt, tiltsens
Dim Tilted
Dim CurrBall
Dim mhole
Dim BallObj, BallinPlay
Dim LastChime10:Dim LastChime100:Dim LastChime1000:
Dim AllRoRewarded:Dim SpecialRewarded
Dim light, objekt
Dim AttractTick

Sub C37_Init
	LoadEM
	DelayTimer.Enabled = 1:DelayTimer.Interval=45

	BallsPerGame=5 		'SET BALLS PER GAME

	hisc=22420  		'SET DEFAULT VALUES PRIOR TO LOADING FROM SAVED FILE
	HSA1=4
	HSA2=15
	HSA3=7
	match=0
	Temperature=7

	Matchtxt.text=""
	BIPbox.text=""
	Tilttxt.text=""
	GOtxt.text="GAME OVER"

	set BallObj=Drain.CreateBall

	loadhs
	UpdatePostIt
	CreditReel.setvalue Credits
	ScoreReel.setvalue Score MOD 100000
	If Score>99999 then Reel100k.setvalue Int(Score/100000)
	ThermoReel.setvalue Temperature
	if match=0 then
		matchtxt.text="00"
	  else
		matchtxt.text=match*10
	end if
	BallinPlay=0
	MotorRunning=0
	MazeState=0
	MatchPause=0
	CheckLights
	For each light in GIlights:light.state=1:next
	BumpersOff
	ThermRewindTimer.interval = 20	
	ThermRewindTimer.Enabled = 0
	AttractTimer.interval = 75
	AttractTimer.Enabled = 0

	CheckLights

	If B2SOn Then
		Controller.B2sSetScorePlayer 1, Score
		Controller.B2SSetData 150, 1
		Controller.B2SSetGameover 4
		Controller.B2SSetCredits Credits
		Controller.B2ssetMatch 34, Match
	End If

	If ShowDT = True Then
		For each objekt in backdropstuff 
		Objekt.visible = 1 	
		Next
	End If
	
	If ShowDt = False Then
		For each objekt in backdropstuff 
		Objekt.visible = 0 	
		Next
	End If

	if ballshadows=1 then
		BallShadowUpdate.enabled=1
	  else
		BallShadowUpdate.enabled=0
	end if

	if flippershadows=1 then 
		FlipperLSh.visible=1
		FlipperRSh.visible=1
	  else
		FlipperLSh.visible=0
		FlipperRSh.visible=0
	end if

End Sub

'Primitive Flipper Code
Sub FlipperTimer_Timer
	FlipperT1.roty = LeftFlipper.currentangle 
	FlipperT5.roty = RightFlipper.currentangle
	if FlipperShadows=1 then
		FlipperLsh.rotz= LeftFlipper.currentangle
		FlipperRsh.rotz= RightFlipper.currentangle
	end if
	Pgate.rotz = Gate3.currentangle+25
End Sub

'*****************************************
'			BALL SHADOW
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
        If BOT(b).X < C37.Width/2 Then
            BallShadow(b).X = ((BOT(b).X) - (Ballsize/16) + ((BOT(b).X - (C37.Width/2))/17))' + 13
        Else
            BallShadow(b).X = ((BOT(b).X) + (Ballsize/16) + ((BOT(b).X - (C37.Width/2))/17))' - 13
        End If
        ballShadow(b).Y = BOT(b).Y + 10
        If BOT(b).Z > 20 Then
            BallShadow(b).visible = 1
        Else
            BallShadow(b).visible = 0
        End If
    Next
End Sub


Sub StartGame()
	PlaySound "startup_norm"
	AllLampsOff
	GIon
	ResetDrops
	NextBallPause = 60
	ZeroTherm
	BumpersOn
	MazeState = 1:CheckLights
	AllRoRewarded=0:SpecialRewarded=0
	Score = 0 
	SetLamp 1,1:SetLamp 2,1:SetLamp 3,1:SetLamp 4,1 'turn on rollover lights
	BallsRemaining = BallsPerGame ' first ball
	TiltDelay = 0
	Tilted = False
	Tilt = 0
	TurnOn
	GameStarted = 1
	Credits = Credits - 1
	If Credits < 1 Then DOF 138, DOFOff
	CurrBall = 0
	CreditReel.setvalue Credits
	ScoreReel.resettozero
	Reel100k.setvalue 0
	GOtxt.text=""
	Matchtxt.text=""
	If B2SOn Then
		Controller.B2SSetMatch 34,0
		Controller.B2SSetCredits Credits
		Controller.B2SSetScoreRollover 25,1
	End If
End Sub

Sub GameOver()
	MatchPause = 20
	GameStarted = 0
	BIPBox.text=""
	GOtxt.text="GAME OVER"
	BumpersOff
	GIoff
	AllLampsOff
	For each light in DTLeftlights:light.state=0:next
	LeftFlipper.RotateToStart:StopSound "PNK_MH_Flip_L_up":DOF 101, DOFOff
	RightFlipper.RotateToStart:StopSound "PNK_MH_Flip_R_up":DOF 102, DOFOff
'	Bumper1.Disabled = 1
'	Bumper2.Disabled = 1
'	Bumper3.Disabled = 1

	If B2SOn Then
		Controller.B2SSetGameover 4
		Controller.B2SSetScoreRollover 25,0
	End If

	If Score > hisc Then 
		hisc=score
		GetSpecial

		HighScoreEntryInit()
		HStimer.uservalue = 0
		HStimer.enabled=1
		UpdatePostIt
	End If
	savehs
End Sub

Sub ResetMachine
	AttractTimer.Enabled = 0
	GIoff
	AllLampsOff
	For each light in DTLeftlights:light.state=0:next
	RestartDelay=12

	If B2SOn Then
		Controller.B2SSetScorePlayer 1, 0
		Controller.B2SSetData 6, 0
		Controller.B2SSetData 1, 4
		Controller.B2SSetData 2, 0
		Controller.B2SSetData 3, 0
		Controller.B2SSetData 4, 0
		Controller.B2SSetData 5, 0
		Controller.B2SSetGameover 0
	End If
End Sub

'KICKER

Dim rkickstep

Sub kicker1_hit()	'Kicker at the top of the playfield
   	If Tilted = true then
		PkickarmR.rotz=10
		rkickstep = 0
		kicker1.timerenabled=1
		exit sub
    else
		PlaySound "Hole_enter"
		ScoreKOH
		PkickarmR.rotz=10
		rkickstep = 0
		kicker1.timerenabled=1  'sit in kicker?
	end if
End Sub


Sub kicker1_timer
 	kicker1.kick 167+INT(RND*6),6+INT(RND*2)
  	Playsound SoundFXDOF("popper_ball",119,DOFPulse,DOFContactors)
	DOF 117, DOFPulse
	PkickarmR.rotz=0
	kicker1.timerenabled=0
	rkickstep=rkickstep+1
End Sub



Sub ScoreKOH
	dim KOHBonus
	KOHBonus=1000
	If LampState(13)=1 Then KOHBonus=KOHBonus+1000
	If LampState(14)=1 Then KOHBonus=KOHBonus+1000
	If LampState(15)=1 Then KOHBonus=KOHBonus+1000
	If LampState(16)=1 Then KOHBonus=KOHBonus+1000
	SetMotor(KOHBonus)
	If LampState(17)=1 And SpecialRewarded=0 Then 
		GetSpecial
		SpecialRewarded=1 'PF special
		SetLamp 17,0
	End If
End Sub

Sub GetSpecial
		playsound SoundFXDOF("knocker",106,DOFPulse,DOFKnocker)
		DOF 117, DOFPulse
		AddCredits(1)
End Sub

Sub AddCredits(x)
	Credits=Credits + x
	DOF 138, DOFOn
	If Credits>maxCredits then Credits=maxCredits
	CreditReel.setvalue Credits
	If B2SOn Then Controller.B2SSetCredits Credits
End Sub
	
Sub AttractTimer_Timer()
	AttractTick=AttractTick+1
	Select Case AttractTick
		Case 1:
			AllLampsOff:SetLamp 28,1
		Case 2
			SetLamp 25,1:SetLamp 26,1:SetLamp 24,1:SetLamp 27,1
		Case 3
			SetLamp 9,1:SetLamp 10,1
			
		Case 4
			SetLamp 5,1	
			SetLamp 28,0
			
		Case 5
			SetLamp 4,1:SetLamp 6,1	
			SetLamp 25,0:SetLamp 26,0:SetLamp 24,0:SetLamp 27,0
			
		Case 6
			SetLamp 7,1:SetLamp 22,1:SetLamp 23,1	
			SetLamp 9,0:SetLamp 10,0
			
		Case 7
			SetLamp 8,1:SetLamp 11,1:SetLamp 12,1
			SetLamp 5,0
				
		Case 8
			SetLamp 20,1:SetLamp 21,1:SetLamp 31,1
			SetLamp 4,0:SetLamp 6,0
		Case 9
			SetLamp 13,1
			SetLamp 7,0:SetLamp 22,0:SetLamp 23,0
		Case 10
			SetLamp 14,1:SetLamp 29,1:SetLamp 30,1	
			SetLamp 8,0:SetLamp 11,0:SetLamp 12,0
		Case 11
			SetLamp 15,1:SetLamp 18,1:SetLamp 19,1	
			SetLamp 20,0:SetLamp 21,0:SetLamp 31,0
		Case 12
			SetLamp 16,1	
			SetLamp 13,0
		Case 13
			SetLamp 17,1	
			SetLamp 14,0:SetLamp 29,0:SetLamp 30,0
		Case 14
			SetLamp 1,1:SetLamp 2,1:SetLamp 3,1
			SetLamp 15,0:SetLamp 18,0:SetLamp 19,0
		Case 15 
			SetLamp 16,0	
		Case 16
			SetLamp 17,0
		Case 17
		'	SetLamp 1,0:SetLamp 2,0:SetLamp 3,0
		'rev back
		Case 18 
			SetLamp 17,1:GIon
		Case 19 
			SetLamp 16,1	
		Case 20
			SetLamp 15,1:SetLamp 18,1:SetLamp 19,1
		Case 21
			SetLamp 14,1:SetLamp 29,1:SetLamp 30,1	
		Case 22
			SetLamp 13,1
		Case 23
			SetLamp 20,1:SetLamp 21,1:SetLamp 31,1
		Case 24
			SetLamp 8,1:SetLamp 11,1:SetLamp 12,1
		Case 25
			SetLamp 7,1:SetLamp 22,1:SetLamp 23,1
		Case 26
			SetLamp 4,1:SetLamp 6,1	
		Case 27
			SetLamp 5,1	
		Case 28
			SetLamp 9,1:SetLamp 10,1
		Case 29
			SetLamp 25,1:SetLamp 26,1:SetLamp 24,1:SetLamp 27,1
		Case 30
			SetLamp 28,1
		Case 33
			AllLampsOff
		Case 45
			GIOn:For x = 1 to 31:SetLamp x, 1:Next:
		Case 55
			AllLampsOff
		Case 65
			GIOn:For x = 1 to 31:SetLamp x, 1:Next:
		Case 75
			AllLampsOff
		Case 90
			AttractTick=0
	End Select
End Sub



'*******************************************
' KEYS
'******************************************

Sub C37_KeyDown(ByVal keycode)

	If keycode = PlungerKey Then
		Plunger.PullBack
		PlaySound "plungerpull",0,1,0.25,0.25
	End If

  If Gamestarted AND NOT Tilted Then
	If keycode = LeftFlipperKey Then
		LeftFlipper.RotateToEnd
		PlaySound SoundFXDOF("fx_flipperup",101,DOFOn,DOFFlippers), 0, .67, -0.05, 0.05
		PlaySound "Buzz",-1,.05,-0.05, 0.05
	End If
    
	If keycode = RightFlipperKey Then
		RightFlipper.RotateToEnd
		PlaySound SoundFXDOF("fx_flipperup",102,DOFOn,DOFFlippers), 0, .67, 0.05, 0.05
		PlaySound "Buzz1",-1,.05,0.05,0.05
	End If
    
	If keycode = LeftTiltKey Then
		Nudge 90, 2
		checktilt
	End If
    
	If keycode = RightTiltKey Then
		Nudge 270, 2
		checktilt
	End If
    
	If keycode = CenterTiltKey Then
		Nudge 0, 2
		checktilt
	End If

	If keycode = MechanicalTilt Then
		gametilted
	End If

  end if

	If keycode = AddCreditKey Then AddCredits(1):Playsound "Coin"
	If keycode = StartGameKey AND GameStarted = 0 And Credits > 0 And Not HSEnterMode=true Then ResetMachine
	If HSEnterMode Then HighScoreProcessKey(keycode)

End Sub

Sub C37_KeyUp(ByVal keycode)

	If keycode = PlungerKey Then
		Plunger.Fire
		PlaySound "plunger",0,1,0.25,0.25
	End If

   'If Gamestarted AND NOT Tilted Then
	If keycode = LeftFlipperKey Then
		LeftFlipper.RotateToStart
		PlaySound SoundFXDOF("fx_flipperdown",101,DOFOff,DOFFlippers), 0, 1, -0.05, 0.05
		StopSound "Buzz"
	End If
    
	If keycode = RightFlipperKey Then
		RightFlipper.RotateToStart
		PlaySound SoundFXDOF("fx_flipperdown",102,DOFOff,DOFFlippers), 0, 1, 0.05, 0.05
		StopSound "Buzz1"
	End If
	'End If
End Sub


Sub FeedBall()
	PlaySound SoundFXDOF("ballrel",107,DOFPulse,DOFContactors)

	BallinPlay=1
	Drain.kick 70,15

	If B2SOn Then
		Controller.B2SSetData 1,0
		Controller.B2SSetData 2,0
		Controller.B2SSetData 3,0
		Controller.B2SSetData 4,0
		Controller.B2SSetData 5,0
		Controller.B2SSetData CurrBall,4
	End If
End Sub

Sub Drain_Hit()
	BumpersOff
	PlaySound "drain"
	DOF 139, DOFPulse
	If GameStarted = 1 then NextBallPause = 27
	If B2SOn Then Controller.B2SSetTilt 0
	Tilted=False
	Tilttxt.text=""
	turnon
End Sub

Sub Kicker_Hit()
	KickOut
End Sub

'## DROP TARGETS

Sub CheckDrop(f)
	Dropped(f)=1
	DropDone=DropDone+1
	PlaySound SoundFX("DTDrop",DOFContactors)
	DOF 111, DOFPulse
	Drop(f).IsDropped=True
	If Tilted=0 then
		SetMotor(500) 'score 500
		If LampState(f+4)=1 Then AddTemperature(1)
	End If
	CheckLights
End Sub

Sub target1_dropped()
  if not tilted then
	Ldt20.state=1
	CheckDrop(1)
  end if
End Sub

Sub target2_dropped()
  if not tilted then
	Ldt21.state=1
	Ldt1.state=1
	CheckDrop(2)
  end if
End Sub

Sub target3_dropped()
  if not tilted then
	Ldt13.state=1
	Ldt2.state=1
	CheckDrop(3)
  end if
End Sub

Sub target4_dropped()
  if not tilted then
	Ldt14.state=1
	Ldt3.state=1
	CheckDrop(4)
  end if
End Sub

Sub ResetDrops
	playsound SoundFXDOF("DTreset",118,DOFPulse,DOFContactors)
	DropDone=0
	SetLamp 9,0:SetLamp 10,0:SetLamp 11,0:SetLamp 12,0
	Ldt20.state=0
	Ldt21.state=0
	Ldt1.state=0
	Ldt2.state=0
	Ldt13.state=0
	Ldt14.state=0
	Ldt3.state=0
	Drop(1).IsDropped=False
	Drop(2).IsDropped=False
	Drop(3).IsDropped=False
	Drop(4).IsDropped=False
End Sub

'## STANDUPS

Sub standup1a_Hit 'lower standup
  if not tilted then
	PlaySound SoundFXDOF("target",108,DOFPulse,DOFContactors)	
	If LampState(9)=0 And LampState(10)=0 Then SetMotor(500)
	If LampState(9)=1 Then SetMotor(5000):SetLamp 28,1 'double advance - is it 5000pts too?
	If LampState(10)=1 Then SetMotor(5000):ResetDrops 'reset drops
  end if
End Sub

Sub standup2a_Hit 'upper standup
  if not tilted then
	PlaySound SoundFXDOF("target",108,DOFPulse,DOFContactors)	
	If LampState(11)=0 And LampState(12)=0 Then SetMotor(500)
	If LampState(11)=1 Then SetMotor(5000):SetLamp 28,1 'double advance - is it 5000pts too?
	If LampState(12)=1 Then SetMotor(5000):ResetDrops 'reset drops
  end if
End Sub

'## POPS

Dim bump1, bump2, bump3

''bumpers
Sub Bumper1_Hit
  if not tilted then
	PlaySound SoundFXDOF("PNK_MH_Pop_L",103,DOFPulse,DOFContactors)
	AddScore(100)
  end if
End Sub
'
Sub Bumper2_Hit
  if not tilted then
	PlaySound SoundFXDOF("PNK_MH_Pop_R",104,DOFPulse,DOFContactors)
	AddScore (100)
  end if
End Sub

Sub Bumper3_Hit
  if not tilted then
	PlaySound SoundFXDOF("PNK_MH_Pop_C",105,DOFPulse,DOFContactors)
	AddScore (1000)
  end if
End Sub

Sub Wall10_hit
  if not tilted then
	AddScore (10)
  end if
End Sub

Sub Wall26_hit
  if not tilted then
	AddScore (10)
  end if
End Sub

Sub Wall28_hit
  if not tilted then
	AddScore (10)
  end if
End Sub

Dim RStep, RRStep

Sub Wall25_Slingshot
  if not tilted then
	AddScore (10)
  end if
    R72a.Visible = 0
    R72a1.Visible = 1
    RStep = 0
    Wall25.TimerEnabled = 1
End Sub

Sub Wall25_Timer
    Select Case RStep
        Case 3:R72a1.Visible = 0:R72a2.Visible = 1
        Case 4:R72a2.Visible = 0:R72a.Visible = 1:Wall25.TimerEnabled = 0 
    End Select
    RStep = RStep + 1
End Sub

Sub Wall27_Slingshot
  if not tilted then
	AddScore (10)
  end if
    R73a.Visible = 0
    R73a1.Visible = 1
    RRStep = 0
    Wall27.TimerEnabled = 1
End Sub

Sub Wall27_Timer
    Select Case RRStep
        Case 3:R73a1.Visible = 0:R73a2.Visible = 1
        Case 4:R73a2.Visible = 0:R73a.Visible = 1:Wall27.TimerEnabled = 0 
    End Select
    RRStep = RRStep + 1
End Sub

'****************************************
'  DELAY TIMER LOOP
'****************************************
Dim NextBallPause, MatchPause, TiltDelay, RestartDelay

Sub DelayTimer_Timer

	If RestartDelay> 0 Then
		RestartDelay = RestartDelay-1
		If RestartDelay=1 Then StartGame
	End If

	If MatchPause> 0 Then
		MatchPause = MatchPause-1
		If MatchPause=0 Then 
			if match=0 then
				matchtxt.text="00"
				If B2SOn then Controller.B2SSetMatch 100
			  else
				matchtxt.text=match*10
				If B2SOn then Controller.B2SSetMatch 34,Match*10
			end if
			if (match*10)=(score mod 100) then 
				GetSpecial
			end if
		End if
	End If

	If TiltDelay> 0 Then
		TiltDelay = TiltDelay - 1
		If TiltDelay = 60 Then Tilt = Tilt - 1:End If

		If TiltDelay = 30 Then Tilt = Tilt - 1:End If
		If TiltDelay = 1 Then Tilt = Tilt - 1:End If
	End If

	If NextBallPause > 0 Then
		NextBallPause = NextBallPause - 1
		If NextBallPause = 0 Then
			BumpersOn':CheckLights
			CurrBall=CurrBall+1
			BIPBox.text=CurrBall
			If CurrBall > BallsPerGame then
				GameOver
			Else
				'new ball
				SetLamp 28,0
				FeedBall
			End If
		End If
	End If
	
End Sub


'## CUSTOM LAMP ROUTINES

Sub GIon
	For each Light in GILights:light.state=1:next
End Sub

Sub GIoff
	For each Light in GILights:light.state=0:next
	For each Light in DTLeftLights:light.state=0:next
End Sub

Sub ToggleMaze
	Select Case MazeState
		Case 0

		Case 1
			MazeState=2
		Case 2
			MazeState=1
	End Select
	CheckLights
End Sub

Sub CheckLights
    Select Case MazeState
		Case 0:'all off
			SetLamp 18,0:SetLamp 21,0:SetLamp 22,0
			SetLamp 19,0:SetLamp 20,0:SetLamp 23,0
			SetLamp 24,0:SetLamp 27,0:SetLamp 25,0:SetLamp 26,0
			SetLamp 9,0:SetLamp 10,0:SetLamp 11,0:SetLamp 12,0
		Case 1:'maze state1
			SetLamp 18,1:SetLamp 21,1:SetLamp 22,1
			SetLamp 19,0:SetLamp 20,0:SetLamp 23,0
			SetLamp 24,1:SetLamp 27,1:SetLamp 25,0:SetLamp 26,0
			If DropDone=4 then
				SetLamp 9,1:SetLamp 10,0:SetLamp 11,0:SetLamp 12,1
			End If
		Case 2:'maze state2
			SetLamp 18,0:SetLamp 21,0:SetLamp 22,0
			SetLamp 19,1:SetLamp 20,1:SetLamp 23,1
			SetLamp 24,0:SetLamp 27,0:SetLamp 25,1:SetLamp 26,1
			If DropDone=4 then
				SetLamp 9,0:SetLamp 10,1:SetLamp 11,1:SetLamp 12,0
			End If
	End Select

End Sub 

Sub BumpersOn()
	SetLamp 29,1:SetLamp 30,1:SetLamp 31,1
End Sub

Sub BumpersOff()
	SetLamp 29,0:SetLamp 30,0:SetLamp 31,0
End Sub


'****************************************
'  SCORE MOTOR
'****************************************

ScoreMotorTimer.Enabled = 1
ScoreMotorTimer.Interval = 135 '135
Dim MotorMode
Dim MotorPosition

Sub SetMotor(x)
	If MotorRunning<>1 And GameStarted = 1 then
		MotorRunning=1
		BumpersOff
		Select Case x
			Case 500:
				MotorMode=100
				MotorPosition=5
			Case 1000:
				AddScore(1000)
				MotorRunning=0
				BumpersOn
			Case 2000:
				MotorMode=1000
				MotorPosition=2
			Case 3000:
				MotorMode=1000
				MotorPosition=3
			Case 4000:
				MotorMode=1000
				MotorPosition=4		
			Case 5000:
				MotorMode=1000
				MotorPosition=5
		End Select
	End If
End Sub

Sub ScoreMotorTimer_Timer
	If MotorPosition > 0 Then
		Select Case MotorPosition
			Case 5,4,3,2:
				If MotorMode=1000 Then 
					AddScore(1000)
				Else
					AddScore(100):ToggleMaze
				End If
				MotorPosition=MotorPosition-1
			Case 1:
				If MotorMode=1000 Then 
					AddScore(1000)
				Else 
					AddScore(100):ToggleMaze
				End If
				MotorPosition=0:MotorRunning=0:BumpersOn
		End Select
	End If
End Sub

Sub AddScore(x)
  if not Tilted then
'	debugtext.text=score
	Select Case x
		Case 10:
			match=match+1
			if match>9 then match=0
			PlayChime(10)
			Score=Score+10:b2sSetScore: ScoreReel.addvalue 10
		Case 100:
			PlayChime(100)
			Score=Score+100:b2sSetScore: ScoreReel.addvalue 100
		Case 1000:
			PlayChime(1000)
			Score=Score+1000:b2sSetScore: ScoreReel.addvalue 1000
	End Select
	'check for digit rollover and fire second chime
	Scorex = Score
	Score100K=Int (Scorex/100000)'Calculate the value for the 100,000's digit
	Score10K=Int ((Scorex-(Score100k*100000))/10000) 'Calculate the value for the 10,000's digit
	ScoreK=Int((Scorex-(Score100k*100000)-(Score10K*10000))/1000) 'Calculate the value for the 1000's digit
	Score100=Int((Scorex-(Score100k*100000)-(Score10K*10000)-(ScoreK*1000))/100) 'Calculate the value for the 100's digit
	Score10=Int((Scorex-(Score100k*100000)-(Score10K*10000)-(ScoreK*1000)-(Score100*100))/10) 'Calculate the value for the 10's digit
	Select Case x
		Case 10:
			If Score10=0 Then PlayChime(100)
		Case 100:
			If Score100=0 Then PlayChime(1000)
	End Select

	If Score100k>0 then Reel100k.setvalue Score100k

	If B2SOn Then Controller.B2SSetScorePlayer 1, Score
  end if
End Sub

Sub PlayChime(x)
	Select Case x
		Case 10
			If LastChime10=1 Then
				PlaySound SoundFXDOF("SJ_Chime_10a",140,DOFPulse,DOFChimes)
				LastChime10=0
			Else
				PlaySound SoundFXDOF("SJ_Chime_10b",140,DOFPulse,DOFChimes)
				LastChime10=1
			End If
		Case 100
			If LastChime100=1 Then
				PlaySound SoundFXDOF("SJ_Chime_100a",141,DOFPulse,DOFChimes)
				LastChime100=0
			Else
				PlaySound SoundFXDOF("SJ_Chime_100b",141,DOFPulse,DOFChimes)
				LastChime100=1
			End If
		Case 1000
			If LastChime1000=1 Then
				PlaySound SoundFXDOF("SJ_Chime_1000a",142,DOFPulse,DOFChimes)
				LastChime1000=0
			Else
				PlaySound SoundFXDOF("SJ_Chime_1000b",142,DOFPulse,DOFChimes)
				LastChime1000=1
			End If
	End Select
End Sub

'****************************************
'  THERMOMETER
'****************************************

Sub ZeroTherm
	ThermRewindTimer.enabled = 1
End Sub

Sub ThermRewindTimer_Timer
	If Temperature > 0 Then
		Temperature=Temperature-1
		ThermoReel.setvalue Temperature
		If B2SOn Then TemperaturB2S(Temperature)
	Else
		Me.enabled = 0
	End If
End Sub

Sub AddTemperature(x)
	If GameStarted = 0 Then Exit Sub
	Select Case x
		Case 1
			If LampState(28)=1 Then x=2
		Case 5
			x=5
	End Select
	Temperature=Temperature+x
	If Temperature>=24 Then
		Temperature=24
		'light special
		If SpecialRewarded=0 then SetLamp 17, 1
	End If
	ThermoReel.setvalue Temperature
	If B2SOn Then TemperaturB2S(Temperature)
End Sub

Sub TemperaturB2S(ttt)
	for x = 150 to 174
		Controller.B2SSetData x,0
	next
	Controller.B2SSetData 150+ttt, 1
End Sub

'****************************************
'  ROLL-OVER SWITCHES
'****************************************

Sub checkRollovers
  if not tilted then
	If GameStarted = 0 Then Exit Sub
	If LampState(1)=0 Then 'A done
		SetLamp 8,1 ' light drop
		SetLamp 16,1 ' light KOH spot
	End If
	If LampState(2)=0 Then 'B done
		SetLamp 7,1 ' light drop
		SetLamp 15,1 ' light KOH spot
	End If
	If LampState(3)=0 Then 'C done
		SetLamp 6,1 ' light drop
		SetLamp 14,1 ' light KOH spot
	End If
	If LampState(4)=0 Then 'D done
		SetLamp 5,1 ' light drop
		SetLamp 13,1 ' light KOH spot
	End If
	If LampState(1)=0 And LampState(2)=0 And LampState(3)=0 And LampState(4)=0 Then
		'only redeem this once
		If AllRoRewarded=0 Then
			AllRoRewarded=1:AddTemperature(5)
		End If
	End If
  end if
End Sub

Sub SwTr1_Hit() 'lane A
  if not tilted then
	SetLamp 1,0:checkRollovers
	SetMotor(500)
	DOF 121, DOFPulse
  end if
	PlaySound "rollover"
End Sub
Sub SwTr2_Hit() 'lane B
  if not tilted then
	SetLamp 2,0:checkRollovers
	SetMotor(500)
	DOF 122, DOFPulse
  end if
	PlaySound "rollover"
End Sub
Sub SwTr3_Hit() 'lane C
  if not tilted then
	SetLamp 3,0:checkRollovers
	SetMotor(500)
	DOF 123, DOFPulse
  end if
	PlaySound "rollover"
End Sub
Sub SwTr4_Hit() 'lane D
  if not tilted then
	SetLamp 4,0:checkRollovers
	SetMotor(500)
	DOF 124, DOFPulse
  end if
	PlaySound "rollover"
End Sub
Sub SwTr5_Hit() 'maze 1
  if not tilted then
	If LampState(18)=1 then SetMotor(5000):AddTemperature(1):else SetMotor(500):end if
	DOF 125, DOFPulse
  end if
	PlaySound "rollover"
End Sub
Sub SwTr6_Hit() 'maze 2
  if not tilted then
	If LampState(19)=1 then SetMotor(5000):AddTemperature(1):else SetMotor(500):end if
	DOF 126, DOFPulse
  end if
	PlaySound "rollover"
End Sub
Sub SwTr7_Hit() 'maze 3
  if not tilted then
	If LampState(20)=1 then SetMotor(5000):AddTemperature(1):else SetMotor(500):end if
	DOF 127, DOFPulse
  end if
	PlaySound "rollover"
End Sub
Sub SwTr8_Hit() 'maze 4
  if not tilted then
	If LampState(21)=1 then SetMotor(5000):AddTemperature(1):else SetMotor(500):end if
	DOF 128, DOFPulse
  end if
	PlaySound "rollover"
End Sub
Sub SwTr9_Hit() 'maze 5
  if not tilted then
	If LampState(22)=1 then SetMotor(5000):AddTemperature(1):else SetMotor(500):end if
	DOF 129, DOFPulse
  end if
	PlaySound "rollover"
End Sub
Sub SwTr10_Hit() 'maze 6
  if not tilted then
	If LampState(23)=1 then SetMotor(5000):AddTemperature(1):else SetMotor(500):end if
	DOF 130, DOFPulse
  end if
	PlaySound "rollover"
End Sub
Sub SwTr11_Hit() 'left outlane
  if not tilted then
	SetMotor(500)
	AddTemperature(1)
	DOF 132, DOFPulse
  end if
	PlaySound "rollover"
End Sub
Sub SwTr12_Hit() 'left inlane 1
  if not tilted then
	If LampState(24)=1 then SetMotor(5000):AddTemperature(1):else SetMotor(500):end if
	DOF 133, DOFPulse
  end if
	PlaySound "rollover"
End Sub
Sub SwTr13_Hit() 'left inlane 2
  if not tilted then
	If LampState(25)=1 then SetMotor(5000):AddTemperature(1):else SetMotor(500):end if
	DOF 134, DOFPulse
  end if
	PlaySound "rollover"
End Sub
Sub SwTr14_Hit() 'right inlane 1
  if not tilted then
	If LampState(26)=1 then SetMotor(5000):AddTemperature(1):else SetMotor(500):end if
	DOF 135, DOFPulse
  end if
	PlaySound "rollover"
End Sub
Sub SwTr15_Hit() 'right inlane 2
  if not tilted then
	If LampState(27)=1 then SetMotor(5000):AddTemperature(1):else SetMotor(500):end if
	DOF 136, DOFPulse
  end if
	PlaySound "rollover"
End Sub
Sub SwTr16_Hit() 'right outlane
  if not tilted then
	SetMotor(500)
	AddTemperature(1)
	DOF 137, DOFPulse
  end if
	PlaySound "rollover"
End Sub

'*****************************************
'  JP's Fadings Lamps 3.0 VP9 Fading only
'        for originals only
'     (based on PD's fading lights)
' SetLamp 0 is Off
' SetLamp 1 is On
' LampState(x) current state
' FadingState(x) current fading state
'*****************************************


Const NumberOfLamps = 33
Redim LampState(NumberOfLamps), FadingState(NumberOfLamps)
dim FadingLevel(200)
 
InitLamps()
LampTimer.Interval = 52  '45 reco
LampTimer.Enabled = 1
 
Sub LampTimer_Timer()
	UpdateLamps
End Sub
 
Sub UpdateLamps()
	FadeL 1, l1, l1 'A Top Rollovers
	FadeL 2, l2, l2 'B
	FadeL 3, l3, l3 'C
	FadeLm 4, l4, l4 'D
	'NFadeObjm 4, metalwall_prim, "metalDON", "metalDOFF"
	FadeL 5, l5, l5 '1 D Drop Target /note reversed mapped 
	FadeL 6, l6, l6 '2 C
	FadeL 7, l7, l7 '3 B
	FadeL 8, l8, l8 '4 A
	FadeL 9, l9, l9 'Lower Standup Pink (DA)
	FadeL 10, l10, l10 'Lower Standup Green (5k/reset)
	FadeL 11, l11, l11 'Upper Standup Pink (DA)
	FadeL 12, l12, l12 'Upper Standup Green (5k/reset)
	FadeL 13, l13, l13 'KOH 1 D
	FadeL 14, l14, l14 'KOH 2 C
	FadeL 15, l15, l15 'KOH 3 B
	FadeL 16, l16, l16 'KOH 4 A
	FadeL 17, l17, l17 'KOH Special
	FadeL 18, l18, l18 'Maze 1
	FadeL 19, l19, l19 'Maze 2
	FadeL 20, l20, l20 'Maze 3
	FadeL 21, l21, l21 'Maze 4
	FadeL 22, l22, l22 'Maze 5
	FadeL 23, l23, l23 'Maze 6
	FadeL 24, l24, l24 'Inlane 1
	FadeL 25, l25, l25 'Inlane 2
	FadeL 26, l26, l26 'Inlane 3
	FadeL 27, l27, l27 'Inlane 4
	FadeL 28, l28, l28 'Double Advance
	FadeLm 29, bumperlight1, bumperlight1 'Bumper 1
	FadeLm 30, bumperlight2, bumperlight2 'Bumper 2
	FadeLm 31, bumperlight3, bumperlight3 'Bumper 3
	FadeLm 29, bumperlightA1, bumperlightA1 'Bumper 1
	FadeLm 30, bumperlightA2, bumperlightA2 'Bumper 2
	FadeLm 31, bumperlightA3, bumperlightA3 'Bumper 3	
	FadeLm 29, bumperlightB1, bumperlightB1 'Bumper 1
	FadeLm 30, bumperlightB2, bumperlightB2 'Bumper 2
	FadeLm 31, bumperlightB3, bumperlightB3 'Bumper 3	
	'GI
'	FadeLm 32, gip1, gip1a 'lane cover lowerleft	FIX
'	FadeLm 32, gip2, gip2a 'lane cover lower right
'	FadeLm 32, gip3, gip3a 'left triangle
'	FadeLm 32, gip4, gip4a 'DT plastic
'	FadeLM 32, gip5, gip5a 'top left plastic
'	FadeLM 32, gip6, gip6a 'top right plastic
'	FadeLM 32, gip7, gip7a 'upper right maze plastic
'	FadeLM 32, gip8, gip8a 'lower right maze plastic
'	FadeLM 32, gip9, gip9a '2 top rollover cover
'	FadeLM 32, gip10, gip10a '3 top rollover cover
'	FadeLM 32, gip11, gip11a 'maze lane cover
'	FadeLM 32, gip12, gip12a '1 top rollover cover
'	FadeL 32, gipf1, gipf1a ' GI playfield lights
	
End Sub


Sub InitLamps():For x = 1 to NumberOfLamps:LampState(x) = 0:FadingState(x) = 4:Next:End Sub 'turn off all the lamps and init arrays

Sub AllLampsOff():For x = 1 to NumberOfLamps:SetLamp x, 0:Next:End Sub

Sub SetLamp(nr, value)
		LampState(nr) = value
		FadingState(nr) = abs(value) + 4
End Sub


'Sub SetBonus(x)
'	Select Case x
'		Case 1:SetLamp 19,1:SetLamp 20,0:SetLamp 21,0:SetLamp 22,0:SetLamp 23,0:SetLamp 24,0
'		Case 2:SetLamp 19,0:SetLamp 20,1:SetLamp 21,0:SetLamp 22,0:SetLamp 23,0:SetLamp 24,0
'		Case 3:SetLamp 19,0:SetLamp 20,0:SetLamp 21,1:SetLamp 22,0:SetLamp 23,0:SetLamp 24,0
'		Case 4:SetLamp 19,0:SetLamp 20,0:SetLamp 21,0:SetLamp 22,1:SetLamp 23,0:SetLamp 24,0
'		Case 5:SetLamp 19,0:SetLamp 20,0:SetLamp 21,0:SetLamp 22,0:SetLamp 23,1:SetLamp 24,0
'		Case 6:SetLamp 19,0:SetLamp 20,0:SetLamp 21,0:SetLamp 22,0:SetLamp 23,0:SetLamp 24,1
'	End Select
'End Sub

Sub FadeW(nr, a, b, c)
	Select Case FadingState(nr)
		Case 2:a.IsDropped = 1:b.IsDropped = 1:c.IsDropped = 1:FadingState(nr) = 0 'Off
		Case 3:b.IsDropped = 1:c.IsDropped = 1:a.IsDropped = 0:FadingState(nr) = 2 'fading 33
		Case 4:a.IsDropped = 1:c.IsDropped = 1:b.IsDropped = 0:FadingState(nr) = 3 'fading 66
		Case 5:b.IsDropped = 1:c.IsDropped = 1:a.IsDropped = 0:FadingState(nr) = 6 'turning ON 33
		Case 6:a.IsDropped = 1:c.IsDropped = 1:b.IsDropped = 0:FadingState(nr) = 7 'turning ON 66
		Case 7:a.IsDropped = 1:b.IsDropped = 1:c.IsDropped = 0:FadingState(nr) = 1 'ON
	End Select
End Sub

'Sub FadeW(nr, a, b, c)
'	Select Case FadingState(nr)
'		Case 2:c.IsDropped = 1:FadingState(nr) = 0                 'Off
'		Case 3:b.IsDropped = 1:c.IsDropped = 0:FadingState(nr) = 2 'fading...
'		Case 4:a.IsDropped = 1:b.IsDropped = 0:FadingState(nr) = 3 'fading...
'		Case 5:b.IsDropped = 1:c.IsDropped = 0:FadingState(nr) = 6 'turning ON
'		Case 6:c.IsDropped = 1:a.IsDropped = 0:FadingState(nr) = 1 'ON
'	End Select
'End Sub

Sub FadeWm(nr, a, b, c)
	Select Case FadingState(nr)
		Case 2:c.IsDropped = 1
		Case 3:b.IsDropped = 1:c.IsDropped = 0
		Case 4:a.IsDropped = 1:b.IsDropped = 0
		Case 5:b.IsDropped = 1:c.IsDropped = 0
	Case 6:c.IsDropped = 1:a.IsDropped = 0
	End Select
End Sub

Sub NFadeW(nr, a)
	Select Case FadingState(nr)
		Case 4:a.IsDropped = 1:FadingState(nr) = 0
		Case 5:a.IsDropped = 0:FadingState(nr) = 1
	End Select
End Sub
 
Sub NFadeWm(nr, a)
	Select Case FadingState(nr)
		Case 4:a.IsDropped = 1
		Case 5:a.IsDropped = 0
	End Select
End Sub
 
Sub NFadeWi(nr, a)
	Select Case FadingState(nr)
		Case 5:a.IsDropped = 1:FadingState(nr) = 0
		Case 4:a.IsDropped = 0:FadingState(nr) = 1
	End Select
End Sub

Sub FadeL(nr, a, b)
	Select Case FadingState(nr)
		Case 2:b.state = 1:b.state = 0:FadingState(nr) = 0 'OFF
		Case 3:b.state = 1:FadingState(nr) = 2 'fading 33
		Case 4:a.state = 1:a.state = 0:FadingState(nr) = 3 'fading 66
		Case 5:b.state = 1:FadingState(nr) = 6 'turning ON 33
		Case 6:a.state = 1:a.state = 0:FadingState(nr) = 7 'turning ON 66
		Case 7:a.state = 1:FadingState(nr) = 1 'ON
	End Select
End Sub

'Sub FadeL(nr, a, b)
'	Select Case FadingState(nr)
'		Case 2:b.state = 0:FadingState(nr) = 0
'		Case 3:b.state = 1:FadingState(nr) = 2
'		Case 4:a.state = 0:FadingState(nr) = 3
'		Case 5:b.state = 1:FadingState(nr) = 6
'		Case 6:a.state = 1:FadingState(nr) = 1
'	End Select
'End Sub
 
Sub FadeLm(nr, a, b)
	Select Case FadingState(nr)
		Case 2:b.state = 0
		Case 3:b.state = 1
		Case 4:a.state = 0
		Case 5:b.state = 1
		Case 6:a.state = 1
	End Select
End Sub
 
Sub NFadeL(nr, a)
	Select Case FadingState(nr)
		Case 4:a.state = 0:FadingState(nr) = 0
		Case 5:a.State = 1:FadingState(nr) = 1
	End Select
End Sub
 
Sub NFadeLm(nr, a)
	Select Case FadingState(nr)
		Case 4:a.state = 0
		Case 5:a.State = 1
	End Select
End Sub
 
Sub FadeR(nr, a)
	Select Case FadingState(nr)
		Case 2:a.SetValue 3:FadingState(nr) = 0
		Case 3:a.SetValue 2:FadingState(nr) = 2
		Case 4:a.SetValue 1:FadingState(nr) = 3
		Case 5:a.SetValue 1:FadingState(nr) = 6
		Case 6:a.SetValue 0:FadingState(nr) = 1
	End Select
End Sub

Sub FadeRm(nr, a)
	Select Case FadingState(nr)
		Case 2:a.SetValue 3
		Case 3:a.SetValue 2
		Case 4:a.SetValue 1
		Case 5:a.SetValue 1
		Case 6:a.SetValue 0
	End Select
End Sub

Sub NFadeT(nr, a, b)
	Select Case FadingState(nr)
		Case 4:a.Text = "":FadingState(nr) = 0
		Case 5:a.Text = b:FadingState(nr) = 1
	End Select
End Sub

Sub NFadeTm(nr, a, b)
	Select Case FadingState(nr)
		Case 4:a.Text = ""
		Case 5:a.Text = b
	End Select
End Sub

Sub NFadeWi(nr, a)
	Select Case FadingState(nr)
		Case 4:a.IsDropped = 0:FadingState(nr) = 0
		Case 5:a.IsDropped = 1:FadingState(nr) = 1
	End Select
End Sub

Sub NFadeWim(nr, a)
	Select Case FadingState(nr)
		Case 4:a.IsDropped = 0
		Case 5:a.IsDropped = 1
	End Select
End Sub

Sub FadeLCo(nr, a, b) 'fading collection of lights
	Dim obj
	Select Case FadingState(nr)
		Case 2:vpmSolToggleObj b, Nothing, 0, 0:FadingState(nr) = 0
		Case 3:vpmSolToggleObj b, Nothing, 0, 1:FadingState(nr) = 2
		Case 4:vpmSolToggleObj a, Nothing, 0, 0:FadingState(nr) = 3
		Case 5:vpmSolToggleObj b, Nothing, 0, 1:FadingState(nr) = 6
		Case 6:vpmSolToggleObj a, Nothing, 0, 1:FadingState(nr) = 1
	End Select
End Sub

Sub FlashL(nr, a, b) ' simple light flash, not controlled by the rom
	Select Case FadingState(nr)
		Case 2:b.state = 0:FadingState(nr) = 0
		Case 3:b.state = 1:FadingState(nr) = 2
		Case 4:a.state = 0:FadingState(nr) = 3
		Case 5:a.state = 1:FadingState(nr) = 4
	End Select
End Sub

Sub MFadeL(nr, a, b, c) 'Light acting as a flash. C is the light number to be restored
	Select Case FadingState(nr)
		Case 2:b.state = 0:FadingState(nr) = 0
			If FadingState(c) = 1 Then SetLamp c, 1
		Case 3:b.state = 1:FadingState(nr) = 2
		Case 4:a.state = 0:FadingState(nr) = 3
		Case 5:a.state = 1:FadingState(nr) = 1
	End Select
End Sub

Sub NFadeB(nr, a, b, c, d, e) 'New Bally Bumpers: a and b are the off state, c and d and on state, no fading. e only for Kiss
	Select Case FadingState(nr)
		Case 4:a.IsDropped = 0:b.IsDropped = 0:c.IsDropped = 1:d.IsDropped = 1:e.IsDropped = 1:FadingState(nr) = 0
		Case 5:a.IsDropped = 1:b.IsDropped = 1:c.IsDropped = 0:d.IsDropped = 0:e.IsDropped = 0:FadingState(nr) = 1
	End Select
End Sub

Sub NFadeBm(nr, a, b, c, d, e)
	Select Case FadingState(nr)
		Case 4:a.IsDropped = 0:b.IsDropped = 0:c.IsDropped = 1:d.IsDropped = 1:e.IsDropped = 1
		Case 5:a.IsDropped = 1:b.IsDropped = 1:c.IsDropped = 0:d.IsDropped = 0:e.IsDropped = 0
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

'***********************************************
'  END OF JP's Fadings Lamps 3.0 VP9 Fading only
'***********************************************



'*****************
'      Tilt
'*****************

Sub CheckTilt
	If Tilttimer.Enabled = True Then 
	 TiltSens = TiltSens + 1
	 if TiltSens = 3 Then GameTilted
	Else
	 TiltSens = 0
	 Tilttimer.Enabled = True
	End If
End Sub

Sub Tilttimer_Timer()
	Tilttimer.Enabled = False
End Sub

Sub GameTilted
	Tilted = True
	tilttxt.text="TILT"
    If B2SOn Then Controller.B2SSetTilt 33,1
    If B2SOn Then Controller.B2ssetdata 1, 0
	playsound "tilt"
	turnoff
End Sub

sub turnoff
	for i=1 to 3
		EVAL("Bumper"&i).hashitevent = 0
	Next
  	LeftFlipper.RotateToStart
	StopSound "Buzz"
	DOF 101, DOFOff
	RightFlipper.RotateToStart
	StopSound "Buzz1"
	DOF 102, DOFOff
end sub    

sub turnon
	for i=1 to 3
		EVAL("Bumper"&i).hashitevent = 1
	Next
end sub 

'****************************************
'  B2S Updating 
'**************************************** 

Sub b2sSetScore
'	debugtext.text=score
	Scorex = Score
	Score100K=Int (Scorex/100000)'Calculate the value for the 100,000's digit
	Score10K=Int ((Scorex-(Score100k*100000))/10000) 'Calculate the value for the 10,000's digit
	ScoreK=Int((Scorex-(Score100k*100000)-(Score10K*10000))/1000) 'Calculate the value for the 1000's digit
	Score100=Int((Scorex-(Score100k*100000)-(Score10K*10000)-(ScoreK*1000))/100) 'Calculate the value for the 100's digit
	Score10=Int((Scorex-(Score100k*100000)-(Score10K*10000)-(ScoreK*1000)-(Score100*100))/10) 'Calculate the value for the 10's digit
	Score1=Int(Scorex-(Score100k*100000)-(Score10K*10000)-(ScoreK*1000)-(Score100*100)-(Score10*10)) 'Calculate the value for the 1's digit

	If B2SOn and Score100k=1 Then Controller.B2SSetData 6, 4

	If Score >= FreeLevel1 AND Free1 <1 then
		GetSpecial
		Free1 = 1
	End If
	If Score >= FreeLevel2 AND Free2 <1 then
		GetSpecial
		Free2 = 1
	End If
	If Score >= FreeLevel3 AND Free3 <1 then
		GetSpecial
		Free3 = 1
	End If
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

Sub C37_Exit()
	savehs
	If B2SOn Then Controller.stop
End Sub

sub savehs
    savevalue "C37", "credit", Credits
    savevalue "C37", "hiscore", hisc
    savevalue "C37", "match", match
    savevalue "C37", "score", score
    savevalue "C37", "temperature", Temperature
'	savevalue "C37", "balls", BallsPerGame
	savevalue "C37", "hsa1", HSA1
	savevalue "C37", "hsa2", HSA2
	savevalue "C37", "hsa3", HSA3
end sub

sub loadhs
    dim temp
	temp = LoadValue("C37", "credit")
    If (temp <> "") then Credits = CDbl(temp)
    temp = LoadValue("C37", "hiscore")
    If (temp <> "") then hisc = CDbl(temp)
    temp = LoadValue("C37", "match")
    If (temp <> "") then match = CDbl(temp)
    temp = LoadValue("C37", "score")
    If (temp <> "") then score = CDbl(temp)
    temp = LoadValue("C37", "temperature")
    If (temp <> "") then Temperature = CDbl(temp)
'    temp = LoadValue("C37", "balls")
'    If (temp <> "") then BallsPerGame = CDbl(temp)
    temp = LoadValue("C37", "hsa1")
    If (temp <> "") then HSA1 = CDbl(temp)
    temp = LoadValue("C37", "hsa2")
    If (temp <> "") then HSA2 = CDbl(temp)
    temp = LoadValue("C37", "hsa3")
    If (temp <> "") then HSA3 = CDbl(temp)
end sub

'==========================================================================================================================================
'============================================================= START OF HIGH SCORES ROUTINES =============================================================
'==========================================================================================================================================
'
'ADD LINE TO TABLE_KEYDOWN SUB WITH THE FOLLOWING:    If HSEnterMode Then HighScoreProcessKey(keycode) AFTER THE STARTGAME ENTRY
'ADD: And Not HSEnterMode=true TO IF KEYCODE=STARTGAMEKEY
'TO SHOW THE SCORE ON POST-IT ADD LINE AT RELEVENT LOCATION THAT HAS:  UpdatePostIt
'TO INITIATE ADDING INITIALS ADD LINE AT RELEVENT LOCATION THAT HAS:  HighScoreEntryInit()
'ADD THE FOLLOWING LINES TO TABLE_INIT TO SETUP POSTIT
'	if HSA1="" then HSA1=25
'	if HSA2="" then HSA2=25
'	if HSA3="" then HSA3=25
'	UpdatePostIt
'ADD HSA1, HSA2 AND HSA3 TO SAVE AND LOAD VALUES FOR TABLE
'ADD A TIMER NAMED HighScoreFlashTimer WITH INTERVAL 100 TO TABLE
'SET HSSSCOREX BELOW TO WHATEVER VARIABLE YOU USE FOR HIGH SCORE.
'ADD OBJECTS TO PLAYFIELD (EASIEST TO JUST COPY FROM THIS TABLE)
'IMPORT POST-IT IMAGES


Dim HSA1, HSA2, HSA3
Dim HSEnterMode, hsLetterFlash, hsEnteredDigits(3), hsCurrentDigit, hsCurrentLetter
Dim HSArray  
Dim HSScoreM,HSScore100k, HSScore10k, HSScoreK, HSScore100, HSScore10, HSScore1, HSScorex	'Define 6 different score values for each reel to use
HSArray = Array("Postit0","postit1","postit2","postit3","postit4","postit5","postit6","postit7","postit8","postit9","postitBL","postitCM","Tape")
Const hsFlashDelay = 4

' ***********************************************************
'  HiScore DISPLAY 
' ***********************************************************

Sub UpdatePostIt
	dim tempscore
	HSScorex = hisc
	TempScore = HSScorex
	HSScore1 = 0
	HSScore10 = 0
	HSScore100 = 0
	HSScoreK = 0
	HSScore10k = 0
	HSScore100k = 0
	HSScoreM = 0
	if len(TempScore) > 0 Then
		HSScore1 = cint(right(Tempscore,1))
	end If
	if len(TempScore) > 1 Then
		TempScore = Left(TempScore,len(TempScore)-1)
		HSScore10 = cint(right(Tempscore,1))
	end If
	if len(TempScore) > 1 Then
		TempScore = Left(TempScore,len(TempScore)-1)
		HSScore100 = cint(right(Tempscore,1))
	end If
	if len(TempScore) > 1 Then
		TempScore = Left(TempScore,len(TempScore)-1)
		HSScoreK = cint(right(Tempscore,1))
	end If
	if len(TempScore) > 1 Then
		TempScore = Left(TempScore,len(TempScore)-1)
		HSScore10k = cint(right(Tempscore,1))
	end If
	if len(TempScore) > 1 Then
		TempScore = Left(TempScore,len(TempScore)-1)
		HSScore100k = cint(right(Tempscore,1))
	end If
	if len(TempScore) > 1 Then
		TempScore = Left(TempScore,len(TempScore)-1)
		HSScoreM = cint(right(Tempscore,1))
	end If
	Pscore6.image = HSArray(HSScoreM):If HSScorex<1000000 Then PScore6.image = HSArray(10)
	Pscore5.image = HSArray(HSScore100K):If HSScorex<100000 Then PScore5.image = HSArray(10)
	PScore4.image = HSArray(HSScore10K):If HSScorex<10000 Then PScore4.image = HSArray(10)
	PScore3.image = HSArray(HSScoreK):If HSScorex<1000 Then PScore3.image = HSArray(10)
	PScore2.image = HSArray(HSScore100):If HSScorex<100 Then PScore2.image = HSArray(10)
	PScore1.image = HSArray(HSScore10):If HSScorex<10 Then PScore1.image = HSArray(10)
	PScore0.image = HSArray(HSScore1):If HSScorex<1 Then PScore0.image = HSArray(10)
	if HSScorex<1000 then
		PComma.image = HSArray(10)
	else
		PComma.image = HSArray(11)
	end if
	if HSScorex<1000000 then
		PComma2.image = HSArray(10)
	else
		PComma2.image = HSArray(11)
	end if
'	if showhisc=1 and showhiscnames=1 then
'		for each objekt in hiscname:objekt.visible=1:next
		HSName1.image = ImgFromCode(HSA1, 1)
		HSName2.image = ImgFromCode(HSA2, 2)
		HSName3.image = ImgFromCode(HSA3, 3)
'	  else
'		for each objekt in hiscname:objekt.visible=0:next
'	end if
End Sub

Function ImgFromCode(code, digit)
	Dim Image
	if (HighScoreFlashTimer.Enabled = True and hsLetterFlash = 1 and digit = hsCurrentLetter) then
		Image = "postitBL"
	elseif (code + ASC("A") - 1) >= ASC("A") and (code + ASC("A") - 1) <= ASC("Z") then
		Image = "postit" & chr(code + ASC("A") - 1)
	elseif code = 27 Then
		Image = "PostitLT"
    elseif code = 0 Then
		image = "PostitSP"
    Else
      msgbox("Unknown display code: " & code)
	end if
	ImgFromCode = Image
End Function

Sub HighScoreEntryInit()
	HSA1=0:HSA2=0:HSA3=0
	HSEnterMode = True
	hsCurrentDigit = 0
	hsCurrentLetter = 1:HSA1=1
	HighScoreFlashTimer.Interval = 250
	HighScoreFlashTimer.Enabled = True
	hsLetterFlash = hsFlashDelay
End Sub

Sub HighScoreFlashTimer_Timer()
	hsLetterFlash = hsLetterFlash-1
	UpdatePostIt
	If hsLetterFlash=0 then 'switch back
		hsLetterFlash = hsFlashDelay
	end if
End Sub


' ***********************************************************
'  HiScore ENTER INITIALS 
' ***********************************************************

Sub HighScoreProcessKey(keycode)
    If keycode = LeftFlipperKey Then
		hsLetterFlash = hsFlashDelay
		Select Case hsCurrentLetter
			Case 1:
				HSA1=HSA1-1:If HSA1=-1 Then HSA1=26 'no backspace on 1st digit
				UpdatePostIt
			Case 2:
				HSA2=HSA2-1:If HSA2=-1 Then HSA2=27
				UpdatePostIt
			Case 3:
				HSA3=HSA3-1:If HSA3=-1 Then HSA3=27
				UpdatePostIt
		 End Select
    End If

	If keycode = RightFlipperKey Then
		hsLetterFlash = hsFlashDelay
		Select Case hsCurrentLetter
			Case 1:
				HSA1=HSA1+1:If HSA1>26 Then HSA1=0
				UpdatePostIt
			Case 2:
				HSA2=HSA2+1:If HSA2>27 Then HSA2=0
				UpdatePostIt
			Case 3:
				HSA3=HSA3+1:If HSA3>27 Then HSA3=0
				UpdatePostIt
		 End Select
	End If
	
    If keycode = StartGameKey Then
		Select Case hsCurrentLetter
			Case 1:
				hsCurrentLetter=2 'ok to advance
				HSA2=HSA1 'start at same alphabet spot
'				EMReelHSName1.SetValue HSA1:EMReelHSName2.SetValue HSA2
			Case 2:
				If HSA2=27 Then 'bksp
					HSA2=0
					hsCurrentLetter=1
				Else
					hsCurrentLetter=3 'enter it
					HSA3=HSA2 'start at same alphabet spot
				End If
			Case 3:
				If HSA3=27 Then 'bksp
					HSA3=0
					hsCurrentLetter=2
				Else
					savehs 'enter it
					HighScoreFlashTimer.Enabled = False
					HSEnterMode = False
				End If
		End Select
		UpdatePostIt
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

Function AudioFade(tableobj) ' Fades between front and back of the table (for surround systems or 2x2 speakers, etc), depending on the Y position on the table. "C37" is the name of the table
  Dim tmp
  tmp = tableobj.y * 2 / C37.height-1
  If tmp > 0 Then
    AudioFade = Csng(tmp ^10)
  Else
    AudioFade = Csng(-((- tmp) ^10) )
  End If
End Function

Function AudioPan(tableobj) ' Calculates the pan for a tableobj based on the X position on the table. "C37" is the name of the table
  Dim tmp
  tmp = tableobj.x * 2 / C37.width-1
  If tmp > 0 Then
    AudioPan = Csng(tmp ^10)
  Else
    AudioPan = Csng(-((- tmp) ^10) )
  End If
End Function

Function Pan(ball) ' Calculates the pan for a ball based on the X position on the table. "C37" is the name of the table
    Dim tmp
    tmp = ball.x * 2 / C37.width-1
    If tmp > 0 Then
        Pan = Csng(tmp ^10)
    Else
        Pan = Csng(-((- tmp) ^10) )
    End If
End Function

Function AudioFade(ball) ' Can this be together with the above function ?
  Dim tmp
  tmp = ball.y * 2 / C37.height-1
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
  If C37.VersionMinor > 3 OR C37.VersionMajor > 10 Then
    PlaySound("fx_collide"), 0, Csng(velocity) ^2 / 200, Pan(ball1), 0, Pitch(ball1), 0, 0, AudioFade(ball1)
  Else
    PlaySound("fx_collide"), 0, Csng(velocity) ^2 / 200, Pan(ball1), 0, Pitch(ball1), 0, 0
  End if
End Sub

' Thalamus : Exit in a clean and proper way
Sub C37_exit()
  Controller.Pause = False
  Controller.Stop
End Sub


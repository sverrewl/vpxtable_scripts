'############################################################################################
'############################################################################################
'#######                                                                             ########
'#######          Cavalcade                                                          ########
'#######          (Stoner 1935)                                                      ########
'#######                                                                             ########
'############################################################################################
'############################################################################################
' Version 1.0 FS mfuegemann 2018


' Thalamus 2018-07-19
' Added/Updated "Positional Sound Playback Functions" and "Supporting Ball & Sound Functions"
' Thalamus 2018-08-15 : Improved directional sounds
' !! NOTE : Table not verified yet !!

Option Explicit

On Error Resume Next
ExecuteGlobal GetTextFile("controller.vbs")
If Err Then MsgBox "You need the controller.vbs in order to run this table, available in the vp10 package"
On Error Goto 0

Const cGameName = "Cavalcade"

'---------------------------
'-----  Configuration  -----
'---------------------------
Const BallsperGame = 10				'the original table has 10 Balls (this is the Max)
Const ShowHighScore = True			'enable to show HighScore tape on Apron, a new HighScore will be saved anyway
Const ShowPlayfieldReel = True		'set to True to show Playfield EM-Reel and Post-It for Player Scores
Const FreePlay = True
const RollingSoundFactor = 2.5		'set sound level factor here for Ball Rolling Sound, 1=default level

Const ResetHighscore = True		'enable to manually reset the Highscores

Const Special1Score = 12500			'set the 3 replay scores
Const Special2Score = 18500
Const Special3Score = 25000

Const VolumeMultiplier = 3 		'adjusts table sound volume


Const BallSize = 25  'Ball radius

Dim i,HighScore,Credits,GameStarted,BallinPlay,BallReleaseLocked
'--------------------------
'------  Table Init  ------
'--------------------------

Sub Cavalcade_Init
	RampGate2.visible = False
	LoadEM
	LoadHighScore
	if ResetHighScore then
		SetDefaultHSTD
	end if
	if ShowHighScore then
		PTape.image = HSArray(12)
		UpdatePostIt
	else
		PTape.image = HSArray(10)
		PComma.image = HSArray(10)
		Pscore0.image = HSArray(10)
		PScore1.image = HSArray(10)
		PScore2.image = HSArray(10)
		PScore3.image = HSArray(10)
		PScore4.image = HSArray(10)
	end if

	Randomize

	if FreePlay then
		Credits = 5
		DOF 126, DOFOn
	end If

	if not Cavalcade.ShowDT then
		Primitive52.visible = False
		Ramp16.widthbottom = 0
		Ramp16.widthtop = 0
		Ramp15.widthbottom = 0
		Ramp15.widthtop = 0
	end If

	if not ShowPlayfieldReel then
		ReelWall.isdropped = True
		GameOnLight.visible = False
		P_Reel0.Transz = -5
		P_Reel1.Transz = -5
		P_Reel2.Transz = -5
		P_Reel3.Transz = -5
		P_Reel4.Transz = -5
		ReelWall2.isdropped = True
		P_BallinPlay.Transz = -10
		P_BallinPlay1.Transz = -10
		P_BallinPlayText.Transz = -10
	end If

	BallinPlay = 0
	GameStarted = False
	Val10 = 0
	Val100 = 0
	Val1000 = 0
	Val10000 = 0
	Val100000 = 0
	Oldscore = 0
	UpdateScoreReel.enabled = True
	Tilt = 0
End Sub


'-----------------------------
'------  Game Features  ------
'-----------------------------

Sub MainTimer_Timer
	if (Ballinplay = BallsperGame) and (not BallReleaseLocked) Then
		If BallVel(LBall) = 0 Then
			GameEndTimer.enabled = True
		Else
			if GameEndTimer.enabled Then
				GameEndTimer.enabled = False
			end If
		end If
	end If

	if GameStarted Then
		if Ballinplay = BallsperGame then
			GameOnLight.state = Lightstateblinking
		Else
			GameOnLight.state = Lightstateon
		end If
	Else
		GameOnLight.state = Lightstateoff
	End If
End Sub

Sub StartGame
	if Credits > 0 Then
		if not GameStarted then
			GameEndTimer.interval = 2000
			BallReleaseLocked = False
			Playerscore = 0
			Oldscore = 0
			Val10 = 0
			Val100 = 0
			Val1000 = 0
			Val10000 = 0
			Val100000 = 0
			Special1 = False
			Special2 = False
			Special3 = False
			UpdateScoreReel.enabled = True
			GameStarted = True
			BallinPlay = 0
			P_BallinPlay.image = CStr(0)
			P_BallinPlay1.image = CStr(0)
			P_Tiltsign.TransY = 0
			Tilt = 0
			Credits = Credits - 1
			If Credits < 1 Then DOF 126, DOFOff
			if FreePlay and (Credits = 0) then
				Credits = 5
				DOF 126, DOFOn
			end If
		end If
	end If
End Sub

Sub NextBall
	if BallinPlay < BallsperGame Then
		BallReleaseLocked = True
		PlaySoundAt "Plunger_feed", Plunger
		BallinPlay = BallinPlay + 1
		if Ballinplay < 10 then
			P_BallinPlay.image = CStr(BallinPlay)
			P_BallinPlay1.image = CStr(0)
		Else
			P_BallinPlay.image = CStr(0)
			P_BallinPlay1.image = CStr(1)
		end If
		BallRelease.Timerenabled = True
	End If
End Sub

Sub BallRelease_Timer
	BallRelease.Timerenabled = False
	BallRelease.CreateBall
	BallRelease.Kick 270, 7
	DOF 122, DOFPulse
End Sub

Sub GameEndTimer_Timer
	GameEndTimer.enabled = False
	EndGame
End Sub

Sub EndGame
	CheckNewHighScorePostIt(Playerscore)
	GameStarted = False
	BallinPlay = 0
End Sub


'--------------------------------
'------  Holes and Points  ------
'--------------------------------

Sub Enter5K_Hit
	if BallinPlay = BallsperGame then
		GameEndTimer.interval = 3000
	end if
End Sub

Sub Award5K_Hit
	Addpoints(5000)
End Sub

Sub Trigger2500_Hit:Addpoints(2500):End Sub
Sub Trigger2500_unHit:Addpoints(-2500):End Sub
Sub Trigger2Ka_Hit:Addpoints(2000):End Sub
Sub Trigger2Ka_unHit:Addpoints(-2000):End Sub

Sub Trigger1500a_Hit:Addpoints(1500):End Sub
Sub Trigger1500a_unHit:Addpoints(-1500):End Sub
Sub Trigger1500b_Hit:Addpoints(1500):End Sub
Sub Trigger1500b_unHit:Addpoints(-1500):End Sub
Sub Trigger1500c_Hit:Addpoints(1500):End Sub
Sub Trigger1500c_unHit:Addpoints(-1500):End Sub
Sub Trigger1500d_Hit:Addpoints(1500):End Sub
Sub Trigger1500d_unHit:Addpoints(-1500):End Sub

Sub Trigger1Ka_Hit:Addpoints(1000):End Sub
Sub Trigger1Ka_unHit:Addpoints(-1000):End Sub
Sub Trigger1Kb_Hit:Addpoints(1000):End Sub
Sub Trigger1Kb_unHit:Addpoints(-1000):End Sub
Sub Trigger1Kc_Hit:Addpoints(1000):End Sub
Sub Trigger1Kc_unHit:Addpoints(-1000):End Sub
Sub Trigger1Kd_Hit:Addpoints(1000):End Sub
Sub Trigger1Kd_unHit:Addpoints(-1000):End Sub
Sub Trigger1Ke_Hit:Addpoints(1000):End Sub
Sub Trigger1Ke_unHit:Addpoints(-1000):End Sub
Sub Trigger1Kf_Hit:Addpoints(1000):End Sub
Sub Trigger1Kf_unHit:Addpoints(-1000):End Sub

Sub Trigger700_Hit:Addpoints(700):End Sub
Sub Trigger700_unHit:Addpoints(-700):End Sub

Sub Trigger500a_Hit:Addpoints(500):End Sub
Sub Trigger500a_unHit:Addpoints(-500):End Sub
Sub Trigger500b_Hit:Addpoints(500):End Sub
Sub Trigger500b_unHit:Addpoints(-500):End Sub
Sub Trigger500c_Hit:Addpoints(500):End Sub
Sub Trigger500c_unHit:Addpoints(-500):End Sub
Sub Trigger500d_Hit:Addpoints(500):End Sub
Sub Trigger500d_unHit:Addpoints(-500):End Sub
Sub Trigger500e_Hit:Addpoints(500):End Sub
Sub Trigger500e_unHit:Addpoints(-500):End Sub
Sub Trigger500f_Hit:Addpoints(500):End Sub
Sub Trigger500f_unHit:Addpoints(-500):End Sub
Sub Trigger500g_Hit:Addpoints(500):End Sub
Sub Trigger500g_unHit:Addpoints(-500):End Sub

Sub Trigger_300a_Hit:Addpoints(300):End Sub
Sub Trigger_300a_UnHit:Addpoints(-300):End Sub
Sub Trigger_300b_Hit:Addpoints(300):End Sub
Sub Trigger_300b_UnHit:Addpoints(-300):End Sub
Sub Trigger_300c_Hit:Addpoints(300):End Sub
Sub Trigger_300c_UnHit:Addpoints(-300):End Sub

Sub Trigger_300d_Hit:Addpoints(300):End Sub
Sub Trigger_300d_UnHit:Addpoints(-300):End Sub
Sub Trigger_300e_Hit:Addpoints(300):End Sub
Sub Trigger_300e_UnHit:Addpoints(-300):End Sub
Sub Trigger_300f_Hit:Addpoints(300):End Sub
Sub Trigger_300f_UnHit:Addpoints(-300):End Sub

Sub Trigger_100a_Hit:Addpoints(100):End Sub
Sub Trigger_100a_UnHit:Addpoints(-100):End Sub
Sub Trigger_100b_Hit:Addpoints(100):End Sub
Sub Trigger_100b_UnHit:Addpoints(-100):End Sub

Sub Trigger_100c_Hit:Addpoints(100):End Sub
Sub Trigger_100c_UnHit:Addpoints(-100):End Sub
Sub Trigger_100d_Hit:Addpoints(100):End Sub
Sub Trigger_100d_UnHit:Addpoints(-100):End Sub

Sub Trigger_1K_Hit:Addpoints(1000):End Sub
Sub Trigger_1K_UnHit:Addpoints(-1000):End Sub
Sub Trigger_2K_Hit:Addpoints(2000):End Sub
Sub Trigger_2K_UnHit:Addpoints(-2000):End Sub
Sub Trigger_3K_Hit:Addpoints(3000):End Sub
Sub Trigger_3K_UnHit:Addpoints(-3000):End Sub



'--------------------------------
'------  Helper Functions  ------
'--------------------------------
Sub PostWall_Hit
	P_Left1.RotY = 185
	P_Left2.RotY = 185
	P_Left3.RotY = 185
	PostWall.timerenabled = True
End Sub

Sub PostWall_Timer
	PostWall.timerenabled = False
	P_Left1.RotY = 180
	P_Left2.RotY = 180
	P_Left3.RotY = 180
End Sub

Sub RampGateTrigger_Hit
	RampGate1.visible = False
	RampGate2.visible = True
	RampGateTrigger.timerenabled = True
  PlaySoundAt "metalhit2", RampGateTrigger
End Sub

Sub RampGateTrigger_Timer
	RampGate1.visible = True
	RampGate2.visible = False
	RampGateTrigger.timerenabled = False
End Sub

Sub Drain_Hit()
  PlaySoundAt "drain3", Drain
	DOF 121, DOFPulse
	Drain.DestroyBall
End Sub

Sub StarterDrain_Hit()
  PlaySoundAt "drain3", StarterDrain
	DOF 121, DOFPulse
	StarterDrain.DestroyBall
	BallinPlay = BallinPlay - 1
	if Ballinplay < 10 then
		P_BallinPlay.image = CStr(BallinPlay)
		P_BallinPlay1.image = CStr(0)
	Else
		P_BallinPlay.image = CStr(0)
		P_BallinPlay1.image = CStr(1)
	end If
End Sub

Dim BOT2, LBall
Sub EnterPF_Hit
	BallReleaseLocked = False
	Set LBall = ActiveBall
End Sub

Dim BoardDir
Sub MoveBoardTimer_Timer
	P_Plate.RotY = P_Plate.RotY - BoardDir

	P_Board.Transy = P_Board.Transy + BoardDir
	if P_Board.Transy < -40 Then
		LockWall.isdropped = True
		TopBallDoor.isdropped = True
		PostKicker.enabled = False
		PostKicker.kick 0,4
	Else
		LockWall.isdropped = False
		TopBallDoor.isdropped = False
		PostKicker.enabled = True
	End If

	if P_Board.Transy < -57 Then
		BoardDir = -Boarddir
		HoldBoardTimer.enabled = True
		MoveBoardTimer.enabled = False
	end If

	if P_Board.Transy >= 0 Then
		P_Board.Transy = 0
		MoveBoardTimer.enabled = False
		StartGame
	end If
End Sub

Sub HoldBoardTimer_Timer
	HoldBoardTimer.enabled = False
	MoveBoardTimer.enabled = True
	PlaySound "MoveBoard_reset",0,0.75,0,0.25,0,0,1,0
end Sub

Dim LeftGateDir
Sub HitLeftGate_Hit:LeftGateDir = 5:HitLeftGate.Timerenabled = True:End Sub
Sub HitLeftGate_UnHit:LeftGateDir = -3:End Sub
Sub HitLeftGate_Timer
	P_LeftGate.roty = P_LeftGate.roty + LeftGateDir
	if P_LeftGate.roty > 25 Then
		P_LeftGate.roty = 25
		LeftGateDir = 0
	end If
	if P_LeftGate.roty <= 0 Then
		HitLeftGate.Timerenabled = False
	end If
End Sub

Dim RightGateDir
Sub HitRightGate_Hit:RightGateDir = -5:HitRightGate.Timerenabled = True:End Sub
Sub HitRightGate_UnHit:RightGateDir = 3:End Sub
Sub HitRightGate_Timer
	P_RightGate.roty = P_RightGate.roty + RightGateDir
	if P_RightGate.roty < -25 Then
		P_RightGate.roty = -25
		RightGateDir = 0
	end If
	if P_RightGate.roty >= 0 Then
		HitRightGate.Timerenabled = False
	end If
End Sub

Sub PostKicker_Hit
  PlaySoundAt "flip_hit_2", PostKicker
End Sub



'--- Tilt recognition ---
Dim Tilt
Sub CheckNudge
	if NudgeTimer1.enabled then
		if NudgeTimer2.enabled then
			NudgeTimer1.enabled = False
			NudgeTimer2.enabled = False
			if Tilt = 0 and GameStarted then
				GameTilted
			end if
		else
			NudgeTimer2.enabled = True
		end if
	else
		NudgeTimer1.enabled = True
	end if
End Sub

Sub NudgeTimer1_Timer
	NudgeTimer1.enabled = False
End Sub

Sub NudgeTimer2_Timer
	NudgeTimer2.enabled = False
End Sub

Sub GameTilted
	Tilt = 1
	Ballinplay = BallsperGame
	P_Tiltsign.TransY = 42
  PlaySound "Knocker"
	PlayerScore = 0
	UpdateScoreReel.enabled = True
End Sub




'------------------------------
'------  Switch Handler  ------
'------------------------------

Sub EnterPost_Hit
	Addpoints(200)
End Sub

Sub LeavePostLeft_Hit
	Addpoints(-200)
End Sub

Sub LeavePostRight_Hit
	Addpoints(-200)
End Sub

Dim KickDir,MoveDir
Sub Starter_Hit
	KickDir = 270
	PostKicker.Timerenabled = True
	P_Starter.Objroty = -60
	Starter.Timerenabled = True
End Sub

Sub Starter_Timer
	P_Starter.Objroty = P_Starter.Objroty + 1
	if P_Starter.Objroty >= -6 then
		Starter.Timerenabled = False
	end if
End Sub

Sub LeftTarget_Hit
	KickDir = 90
	P_LeftTrigger.Roty = 3
	PostKicker.Timerenabled = True
	LeftTarget.Timerenabled = True
End Sub

Sub LeftTarget_Timer
	LeftTarget.Timerenabled = False
	P_LeftTrigger.Roty = -5
End Sub

Sub RightTarget_Hit
	KickDir = 270
	P_RightTrigger.Roty = 3
	PostKicker.Timerenabled = True
	RightTarget.Timerenabled = True
End Sub

Sub RightTarget_Timer
	RightTarget.Timerenabled = False
	P_RightTrigger.Roty = -5
End Sub

Sub PostKicker_Timer
	PostKicker.Timerenabled = False
	PostKicker.kick KickDir,20
	if KickDir = 90 Then
		FlashLight2
		P_Kickerarm.objrotz = 30
		MoveDir = -1
	Else
		FlashLight1
		P_Kickerarm.objrotz = -30
		MoveDir = 1
	end If
  PlaySound "solenoid"
	KickerTimer.enabled = True
End Sub

Sub KickerTimer_Timer
	P_Kickerarm.objrotz = P_Kickerarm.objrotz + MoveDir
	if abs(P_Kickerarm.objrotz) <= 1 Then
		KickerTimer.enabled = False
		P_Kickerarm.objrotz = 0
	end If
End Sub

Sub Kicker_Hit
	Kicker.Timerenabled = True
End Sub

Sub Kicker_Timer
	FlashLight3
	Kicker.Timerenabled = False
	Kicker.kick 275,10
  PlaySound "Popper"
  PlaySound "1000a"
End Sub

Sub FlashLight1
	Light1.state = 1
	Light1a.state = 1
	Light1.timerenabled = True
End Sub

Sub Light1_timer
	Light1.timerenabled = False
	Light1.state = 0
	Light1a.state = 0
End Sub

Sub FlashLight2
	Light2.state = 1
	Light2a.state = 1
	Light2.timerenabled = True
End Sub

Sub Light2_timer
	Light2.timerenabled = False
	Light2.state = 0
	Light2a.state = 0
End Sub

Sub FlashLight3
	Light3.state = 1
	Light3a.state = 1
	Light3.timerenabled = True
End Sub

Sub Light3_timer
	Light3.timerenabled = False
	Light3.state = 0
	Light3a.state = 0
End Sub


Dim Val10,Val100,Val1000,Val10000,Val100000,Score10,Score100,Score1000,Score10000,Score100000,TempScore,Oldscore
Dim PlayerScore,Special1,Special2,Special3
Sub Addpoints(ScorePar)
	if Tilt = 0 Then
		if GameStarted Then
			Playerscore = Playerscore + ScorePar
		end If

		if (Playerscore >= Special1Score) and not Special1 Then
			Special1 = True
'			AddSpecial
		end If
		if (Playerscore >= Special2Score) and not Special2 Then
			Special2 = True
'			AddSpecial
		end If
		if (Playerscore >= Special3Score) and not Special3 Then
			Special3 = True
'			AddSpecial
		end If
		UpdateScoreReel.enabled = True
	end If
End Sub

Sub AddSpecial
	Credits = Credits + 1
  PlaySound "1000a"
End Sub

Sub ShooterLaneLaunch_Hit
	if ActiveBall.vely < -8 then PlaySoundAt "Launch", ShooterLaneLaunch
	DOF 124, DOFPulse
End Sub


Sub UpdateScoreReel_Timer
	TempScore = Playerscore
	Score10 = 0
	Score100 = 0
	Score1000 = 0
	Score10000 = 0
	if len(TempScore) > 1 Then
		TempScore = Left(TempScore,len(TempScore)-1)
		Score10 = cstr(right(Tempscore,1))
	end If
	if len(TempScore) > 1 Then
		TempScore = Left(TempScore,len(TempScore)-1)
		Score100 = cstr(right(Tempscore,1))
	end If
	if len(TempScore) > 1 Then
		TempScore = Left(TempScore,len(TempScore)-1)
		Score1000 = cstr(right(Tempscore,1))
	end If
	if len(TempScore) > 1 Then
		TempScore = Left(TempScore,len(TempScore)-1)
		Score10000 = cstr(right(Tempscore,1))
	end If

	Val10 = ReelValue(Val10,Score10,1)
	Val100 = ReelValue(Val100,Score100,2)
	Val1000 = ReelValue(Val1000,Score1000,3)
	Val10000 = ReelValue(Val10000,Score10000,0)
	Tempscore = Val10 * 10 + Val100 * 100 + Val1000 * 1000 + Val10000 * 10000
	if Oldscore <> TempScore Then
  PlaySound "solon"
		Oldscore = TempScore
		P_Reel1.image = cstr(Val10)
		P_Reel2.image = cstr(Val100)
		P_Reel3.image = cstr(Val1000)
		P_Reel4.image = cstr(Val10000)
	Else
		UpdateScoreReel.enabled = False
	end If
End Sub

Function ReelValue(ValPar,ScorPar,ChimePar)
	ReelValue = cint(ValPar)
	if ReelValue <> cint(ScorPar) Then
		ReelValue = ReelValue + 1
		if ReelValue > 9 Then
			ReelValue = 0
		end If
	end If
End Function

Sub SetScoreReel
	TempScore = Playerscore
	Score10 = 0
	Score100 = 0
	Score1000 = 0
	Score10000 = 0
	if len(TempScore) > 1 Then
		TempScore = Left(TempScore,len(TempScore)-1)
		Score10 = cstr(right(Tempscore,1))
	end If
	if len(TempScore) > 1 Then
		TempScore = Left(TempScore,len(TempScore)-1)
		Score100 = cstr(right(Tempscore,1))
	end If
	if len(TempScore) > 1 Then
		TempScore = Left(TempScore,len(TempScore)-1)
		Score1000 = cstr(right(Tempscore,1))
	end If
	if len(TempScore) > 1 Then
		TempScore = Left(TempScore,len(TempScore)-1)
		Score10000 = cstr(right(Tempscore,1))
	end If
	P_Reel1.image = cstr(Score10)
	P_Reel2.image = cstr(Score100)
	P_Reel3.image = cstr(Score1000)
	P_Reel4.image = cstr(Score10000)
End Sub



'-------------------------------
'------  Keybord Handler  ------
'-------------------------------
Sub Cavalcade_KeyDown(ByVal keycode)
	If keycode = StartGameKey Then
		if not GameStarted and not MoveBoardTimer.enabled and not HoldBoardTimer.enabled Then
			Ballinplay = 0
			BoardDir = -1
			MoveBoardTimer.enabled = True
			PlaySound "MoveBoard_empty",0,0.75,0,0.25,0,0,1,0
		end If
	End If

	If keycode = PlungerKey Then
		Plunger.PullBack
		PlaySoundAt "plungerpull", Plunger
	End If

	If keycode = LeftFlipperKey Then

	End If

	If keycode = RightFlipperKey Then
		if GameStarted and not BallReleaseLocked Then
			NextBall
		end If
	End If

	If keycode = LeftTiltKey Then
		Nudge 90, 2
		CheckNudge
	End If

	If keycode = RightTiltKey Then
		Nudge 270, 2
		CheckNudge
	End If

	If keycode = CenterTiltKey Then
		Nudge 0, 2
		CheckNudge
	End If

End Sub

Sub Cavalcade_KeyUp(ByVal keycode)
	If keycode = PlungerKey Then
		Plunger.Fire
		PlaySoundAt "plunger", Plunger
	End If

	If keycode = LeftFlipperKey Then

	End If

	If keycode = RightFlipperKey Then

	End If

End Sub


'---------------------------
'----- High Score Code -----
'---------------------------
Const HighScoreFilename = "Cavalcade.txt"

Dim HSArray,HSAHighScore, HSA1, HSA2, HSA3
Dim HSScore10k, HSScoreK, HSScore100, HSScore10, HSScore1, HSScorex	'Define 5 different score values for each reel to use
HSArray = Array("HS_0","HS_1","HS_2","HS_3","HS_4","HS_5","HS_6","HS_7","HS_8","HS_9","HS_Space","HS_Comma","Tape")
Const DefaultHighScore = 0

Sub LoadHighScore
	Dim FileObj
	Dim ScoreFile
	Dim TextStr
    Dim SavedDataTemp3 'HighScore
    Set FileObj=CreateObject("Scripting.FileSystemObject")
	If Not FileObj.FolderExists(UserDirectory) then
		Exit Sub
	End if
	If Not FileObj.FileExists(UserDirectory & HighScoreFilename) then
		SetDefaultHSTD:UpdatePostIt:SaveHighScore
		Exit Sub
	End if
	Set ScoreFile=FileObj.GetFile(UserDirectory & HighScoreFilename)
	Set TextStr=ScoreFile.OpenAsTextStream(1,0)
		If (TextStr.AtEndOfStream=True) then
			SetDefaultHSTD:UpdatePostIt:SaveHighScore
			Exit Sub
		End if
		SavedDataTemp3=Textstr.ReadLine ' HighScore
		TextStr.Close
		HSAHighScore=SavedDataTemp3
		UpdatePostIt
	    Set ScoreFile = Nothing
	    Set FileObj = Nothing
End Sub

Sub SetDefaultHSTD  'bad data or missing file - reset and resave
	HSAHighScore = DefaultHighScore
	SaveHighScore
End Sub

Sub UpdatePostIt
	HSScorex = HSAHighScore
	TempScore = HSScorex
	HSScore1 = 0
	HSScore10 = 0
	HSScore100 = 0
	HSScoreK = 0
	HSScore10k = 0
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
End Sub

Dim FileObj,ScoreFile
Sub SaveHighScore
	Set FileObj=CreateObject("Scripting.FileSystemObject")
	If Not FileObj.FolderExists(UserDirectory) then
		Exit Sub
	End if
	Set ScoreFile=FileObj.CreateTextFile(UserDirectory & HighScoreFilename,True)
		ScoreFile.WriteLine HSAHighScore
		ScoreFile.Close
	Set ScoreFile=Nothing
	Set FileObj=Nothing
End Sub

Sub CheckNewHighScorePostIt(newScore)
		If CLng(newScore) > CLng(HSAHighScore) Then
			AddSpecial
			HSAHighScore=newScore
			SaveHighScore
			UpdatePostIt
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


Sub C_TriggerSound_Hit (idx)
	PlaySound "flip_hit_3", 0, 0.15, Pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0, AudioFade(ActiveBall)
End Sub

Sub Pins_Hit (idx)
	PlaySound "pinhit_low", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0, AudioFade(ActiveBall)
End Sub

Sub Targets_Hit (idx)
	PlaySound SoundFX("target",DOFTargets), 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0, AudioFade(ActiveBall)
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
  PlaySoundAt "fx_spinner", Spinner
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

Sub URightFlipper_Collide(parm)
 	RandomSoundFlipper()
End Sub

Sub RandomSoundFlipper()
	Select Case Int(Rnd*3)+1
		Case 1 : PlaySound "flip_hit_1", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
		Case 2 : PlaySound "flip_hit_2", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
		Case 3 : PlaySound "flip_hit_3", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
	End Select
End Sub

'*****************************************
'	ninuzzu's	BALL SHADOW
'*****************************************
Dim BallShadow
BallShadow = Array (BallShadow1,BallShadow2,BallShadow3,BallShadow4,BallShadow5,BallShadow6,BallShadow7,BallShadow8,BallShadow9,BallShadow10,BallShadow11,BallShadow12,BallShadow13,BallShadow14,BallShadow15,BallShadow16,BallShadow17,BallShadow18,BallShadow19,BallShadow20,BallShadow21)

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
        If BOT(b).X < Cavalcade.Width/2 Then
            BallShadow(b).X = ((BOT(b).X) - (Ballsize/6) + ((BOT(b).X - (Cavalcade.Width/2))/7)) + 6
        Else
            BallShadow(b).X = ((BOT(b).X) + (Ballsize/6) + ((BOT(b).X - (Cavalcade.Width/2))/7)) - 6
        End If
        ballShadow(b).Y = BOT(b).Y + 12
        If BOT(b).Z > 100 Then
            BallShadow(b).visible = 1
			ballShadow(b).TransZ = 5 - (0.2 * b)
        Else
            BallShadow(b).visible = 0
        End If
    Next
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

Function AudioFade(tableobj) ' Fades between front and back of the table (for surround systems or 2x2 speakers, etc), depending on the Y position on the table. "Cavalcade" is the name of the table
  Dim tmp
  tmp = tableobj.y * 2 / Cavalcade.height-1
  If tmp > 0 Then
    AudioFade = Csng(tmp ^10)
  Else
    AudioFade = Csng(-((- tmp) ^10) )
  End If
End Function

Function AudioPan(tableobj) ' Calculates the pan for a tableobj based on the X position on the table. "Cavalcade" is the name of the table
  Dim tmp
  tmp = tableobj.x * 2 / Cavalcade.width-1
  If tmp > 0 Then
    AudioPan = Csng(tmp ^10)
  Else
    AudioPan = Csng(-((- tmp) ^10) )
  End If
End Function

Function Pan(ball) ' Calculates the pan for a ball based on the X position on the table. "Cavalcade" is the name of the table
    Dim tmp
    tmp = ball.x * 2 / Cavalcade.width-1
    If tmp > 0 Then
        Pan = Csng(tmp ^10)
    Else
        Pan = Csng(-((- tmp) ^10) )
    End If
End Function

Function AudioFade(ball) ' Can this be together with the above function ?
  Dim tmp
  tmp = ball.y * 2 / Cavalcade.height-1
  If tmp > 0 Then
    AudioFade = Csng(tmp ^10)
  Else
    AudioFade = Csng(-((- tmp) ^10) )
  End If
End Function

Function Vol(ball) ' Calculates the Volume of the sound based on the ball speed
  Vol = Csng(BallVel(ball) ^2 / 2000)
  Vol = Vol + RollingSoundFactor
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

Const tnob = 10 ' total number of balls
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
  If Cavalcade.VersionMinor > 3 OR Cavalcade.VersionMajor > 10 Then
    PlaySound("fx_collide"), 0, Csng(velocity) ^2 / 200, Pan(ball1), 0, Pitch(ball1), 0, 0, AudioFade(ball1)
  Else
    PlaySound("fx_collide"), 0, Csng(velocity) ^2 / 200, Pan(ball1), 0, Pitch(ball1), 0, 0
  End if
End Sub


' Thalamus : Exit in a clean and proper way
Sub Cavalcade_exit()
  Controller.Pause = False
  Controller.Stop
End Sub


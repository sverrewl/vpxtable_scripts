'############################################################################################
'############################################################################################
'#######                                                                             ########
'#######                               Lady Luck                                     ########
'#######                              (Recel 1976)                                   ########
'#######                                                                             ########
'############################################################################################
'############################################################################################
' Version 1.4 WED21 2017
'
' Thanks To:
' Zany for the Flipper model and texture
' JPSalas for the EM Reel images
' GTXJoe for the Highscore saving code
' Loserman and mfuegemann for script help
' Herweh amd Lucasbuck for some images from their VP9 Lady Luck
' Flupper for the physics starting point
' Ninuzzo for Ball Shadow routine
' Every other author for their amazing work!


'Version 1.1 Changes:
'Added beeps or chimes option in script under Configuration section


'Version 1.2 Changes:
'Top diverter changed from wall to a rail
'Adjusted GI and environment lighting
'Made plastics more 3 dimensional
'Adjusted Ball Shadow Size
'Added backdrop and moved reels in Desktop Mode
'Added in more table sounds

'Version 1.3 Changes:
'Added Flipper shadows
'Adjusted Slingshot animation
'Used a new Metal0.8 Material
'Remade shadows based on new rails.
'Changed Ambient lighting "Thanks Hauntfreaks"
'Removed all DOF calls
'Encoded sounds down to 16bit to play on all computers

'Version 1.4 Changes:
'Lowered switch height
'Used rail and stuff to make captive ball holder more like actual machine instead of 2 posts
'50K Primitive!!!! "Thanks to Evan at VPF for making it!!!
'Redrew Plastic where 50K prim is
'Sped the table up just a touch
'Visual touch ups thanks to HauntFreaks
'Adjusted all sounds to PMD for full surround

' Thalamus 2018-08-11 : Improved directional sounds


option Explicit

On Error Resume Next
ExecuteGlobal GetTextFile("controller.vbs")
If Err Then MsgBox "Can't open controller.vbs"
On Error Goto 0

Const ballsize = 50
Const ballmass = 2

Const cGameName = "Lady_Luck"

'---------------------------
'-----  Configuration  -----
'---------------------------
Const BallsperGame = 5				'Change to 3 or 5 balls
Const ShowHighScore = True			'enable to show HighScore tape on Apron, a new HighScore will be saved anyway
Const ShowPlayfieldReel = False		'set to True to show Playfield EM-Reel and Post-It for Player Scores (instead of a B2S backglass)
Const FreePlay = False
const RollingSoundFactor = .9		'set sound level factor here for Ball Rolling Sound, 1=default level
Const ChimesEnabled = False			'True = Chime Sounds, False = Beeps (Default)
Const ChimeVolume = 2

Const ResetHighscore = False		'enable to manually reset the Highscores

Const Special1Score = 600000			'set the 3 replay scores
Const Special2Score = 720000
Const Special3Score = 840000

Dim GameActive,NoofPlayers,i,HighScore,Credits
Dim AwardExtraBall					' flag for extra ball award


Sub Lady_Luck_Init
	LoadEM
	LoadHighScore
	HideScoreboard
	if ResetHighScore then
		SetDefaultHSTD
	end if
	if ShowHighScore then
		PTape.image = HSArray(12)
		UpdatePostIt
	else
		PTape.image = HSArray(10)
		PComma.image = HSArray(10)
		PComma2.image = HSArray(10)
		Pscore0.image = HSArray(10)
		PScore1.image = HSArray(10)
		PScore2.image = HSArray(10)
		PScore3.image = HSArray(10)
		PScore4.image = HSArray(10)
		PScore5.image = HSArray(10)
		PScore6.image = HSArray(10)
	end if

	Randomize
'code added by Loserman
	AwardExtraBall = 0
'
	if FreePlay then
		Credits = 5

	end If

	if B2SOn then
		Controller.B2SSetMatch MatchValue
		Controller.B2SSetScoreRolloverPlayer1 0
		Controller.B2SSetScoreRolloverPlayer2 0
		Controller.B2SSetScoreRolloverPlayer3 0
		Controller.B2SSetScoreRolloverPlayer4 0
		Controller.B2SSetScore 6,HighScore
		Controller.B2SSetTilt 0

		Controller.B2SSetCredits Credits

		Controller.B2SSetGameOver 1

		for i=1 to 4
			Controller.B2SSetScorePlayer i, 0
		next
	end if

	if not Lady_Luck.ShowDT then
		EMReel1.visible = False
		EMReel2.visible = False
		EMReel3.visible = False
		EMReel4.visible = False
		EMReel_BiP.visible = False
		EMReel_Credits.visible = False
		CaseWall1.isdropped = True
		CaseWall2.isdropped = True
		CaseWall3.isdropped = True
		Ramp1.widthbottom = 0
		Ramp1.widthtop = 0
		Ramp15.widthbottom = 0
		Ramp15.widthtop = 0
	end If

	if not ShowPlayfieldReel or Lady_Luck.ShowDT Then
		ReelWall.isdropped = True
		P_Reel0.Transz = -5
		P_Reel1.Transz = -5
		P_Reel2.Transz = -5
		P_Reel3.Transz = -5
		P_Reel4.Transz = -5
		P_Reel5.Transz = -5
		P_ActivePlayer.Transz = -10
		P_Credits.Transz = -10
		P_CreditsText.Transz = -10
		P_BallinPlay.Transz = -10
		P_BallinPlayText.Transz = -10
	end If

	BallinPlay = 0
	GameActive = False
	NoofPlayers = 0
	GameStarted = False
	for i = 1 to 4
		Val10(i) = 0
		Val100(i) = 0
		Val1000(i) = 0
		Val10000(i) = 0
		Val100000(i) = 0
		Oldscore(i) = 0
	Next
	EMReel1.ResetToZero
	EMReel2.ResetToZero
	EMReel3.ResetToZero
	EMReel4.ResetToZero
	EMReel_Credits.setvalue Credits

	P_Credits.image = cstr(Credits)
	Tilt = 0
End Sub

Dim GameStarted,NoPointsScored
Sub StartGame
	if Credits > 0 Then
		if not GameActive then
			for i = 1 to 4
				Playerscore(i) = 0
				Oldscore(i) = 0
				Val10(i) = 0
				Val100(i) = 0
				Val1000(i) = 0
				Val10000(i) = 0
				Val100000(i) = 0
				Special1(i) = False
				Special2(i) = False
				Special3(i) = False
				If B2SOn Then
					Controller.B2SSetScorePlayer i,0
				end if
			Next
			EMReel1.ResetToZero
			EMReel2.ResetToZero
			EMReel3.ResetToZero
			EMReel4.ResetToZero
			if not GameStarted Then
				GameStarted = True
				Playsound "FD_GameStartwithBallrelease"
				GameStartTimer.enabled = True
				ScoreMotor
				ScoreMotorStartTimer.enabled = True
			end If
			if NoofPlayers < 4 Then
				Credits = Credits - 1
				'If Credits < 1 Then DOF 126, DOFOff
				if FreePlay and (Credits = 0) then
					Credits = 5
				'	DOF 126, DOFOn
				end If
				if NoofPlayers > 0 Then
					Playsound "AddPlayer"
				end If
				NoofPlayers = NoofPlayers + 1
				UpdateScoreboard
				EMReel_Credits.setvalue Credits
			end If

			If B2SOn Then
				Controller.B2SSetTilt 0
				Controller.B2SSetGameOver 0
				Controller.B2SSetMatch 0
				Controller.B2SSetBallInPlay BallInPlay
				if Credits=0 then
					Controller.B2SSetCredits 10
				else
					Controller.B2SSetCredits Credits
				end if
				Controller.B2SSetPlayerUp 1
				Controller.B2SSetBallInPlay BallInPlay
				Controller.B2SSetCanPlay NoofPlayers
				Controller.B2SSetScore 6,HSAHighScore
			End if
			EMReel_BiP.setvalue BallinPlay
		end If
	end If
	P_Credits.image = cstr(Credits)
End Sub

Sub GameStartTimer_Timer
	GameStartTimer.enabled = False
	ScoreMotorStartTimer.enabled = False
	ActivePlayer = 0
	BallinPlay = 1
	NextBall
End Sub

Sub ScoreMotorStartTimer_Timer
	ScoreMotor
End Sub

Sub NextBall
	DB=False
	if (ActivePlayer = NoofPlayers) and (BallinPlay = BallsperGame) Then
		EndGame
	Else
		UpdateScoreboard
' Loserman code mod - added extra layer of IF statement checking AwardExtraBall variable
'
		if AwardExtraBall = 0 then
			if ActivePlayer < NoofPlayers Then
				ActivePlayer = ActivePlayer + 1
			Else
				if ActivePlayer = NoofPlayers Then
					BallinPlay = BallinPlay + 1
					ActivePlayer = 1
				end If
			end If
		Else
			AwardExtraBall = 0
			if B2SOn Then
				Controller.B2SSetShootAgain 0
			end If
		end if
'

		P_ActivePlayer.image = "Player" & CStr(ActivePlayer)
		P_BallinPlay.image = CStr(BallinPlay)
		EMReel_BiP.setvalue BallinPlay

		SetScoreReel

		If B2SOn Then
			Controller.B2SSetBallInPlay BallInPlay
			Controller.B2SSetPlayerUp ActivePlayer
			Controller.B2SSetBallInPlay BallInPlay
			Controller.B2SSetTilt 0
		End if

'********GAME START*******************
		ResetPlayfield
		NoPointsScored = True
		BallRelease.CreateBall
		BallRelease.Kick 90, 7
		'DOF 122, DOFPulse
	End If
End Sub
'********************************************************



Dim MatchValue
Sub EndGame
	UpdateScoreboard
	for i = 1 to NoofPlayers
		CheckNewHighScorePostIt(Playerscore(i))
	next
	GameActive = False
	GameStarted = False
	MatchValue=Int(Rnd*10)
	If B2SOn Then
		If MatchValue = 0 Then
			Controller.B2SSetMatch 100
		Else
			Controller.B2SSetMatch MatchValue
		End If
	End if
	for i = 1 to NoofPlayers
		if MatchValue=cint(right(Playerscore(i),2)) then
			AddSpecial
		end if
	next
	BallinPlay = 0
	NoofPlayers = 0
	if B2Son Then
		Controller.B2SSetGameOver 1
		Controller.B2SSetBallInPlay BallInPlay
	end If
End Sub

Sub Drain_Hit()
	PlaySoundAt "drainLong", Drain
	'DOF 121, DOFPulse
	Drain.DestroyBall
	if NoPointsScored and (Tilt = 0) Then
		PlaysoundAt "FD_DrainwithBallrelease", Drain
		ReleaseSameBallTimer.enabled = True
	Else
		AwardBonus
	End If
End Sub

Sub ReleaseSameBallTimer_Timer
	ReleaseSameBallTimer.enabled = False
	BallRelease.CreateBall
	BallRelease.Kick 90, 7
	'DOF 122, DOFPulse
End Sub

Sub MainTimer_Timer

	if not GameStarted Then
		LeftFlipper.RotateToStart
		RightFlipper.RotateToStart
	end If
	if Lady_Luck.showdt Then
		P1Light.State = (Activeplayer = 1)
		P2Light.State = (Activeplayer = 2)
		P3Light.State = (Activeplayer = 3)
		P4Light.State = (Activeplayer = 4)
	end If
End Sub

Sub GameOverTimer_Timer
	if GameStarted Then
		P_GameOver.visible = False
	Else
		if Lady_Luck.showdt Then
			P_GameOver.visible = not P_GameOver.visible
		End If
	End If
End Sub

Sub TiltTimer_Timer
	if Tilt = 0 Then
		P_Tilt.visible = False
	Else
		if Lady_Luck.showdt Then
			P_Tilt.visible = not P_Tilt.visible
		End If
	End If
End Sub


' DT Score Reels
Dim Val10(4),Val100(4),Val1000(4),Val10000(4),Val100000(4),Score10,Score100,Score1000,Score10000,Score100000,TempScore,Oldscore(5)
Sub UpdateScoreReel_Timer
	TempScore = Playerscore(ActivePLayer)
	Score10 = 0
	Score100 = 0
	Score1000 = 0
	Score10000 = 0
	Score100000 = 0
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
	if len(TempScore) > 1 Then
		TempScore = Left(TempScore,len(TempScore)-1)
		Score100000 = cstr(right(Tempscore,1))
	end If

	Val10(ActivePLayer) = ReelValue(Val10(ActivePLayer),Score10,0)
	Val100(ActivePLayer) = ReelValue(Val100(ActivePLayer),Score100,1)
	Val1000(ActivePLayer) = ReelValue(Val1000(ActivePLayer),Score1000,2)
	Val10000(ActivePLayer) = ReelValue(Val10000(ActivePLayer),Score10000,3)
	Val100000(ActivePLayer) = ReelValue(Val100000(ActivePLayer),Score100000,0)
	Tempscore = Val10(ActivePLayer) * 10 + Val100(ActivePLayer) * 100 + Val1000(ActivePLayer) * 1000 + Val10000(ActivePLayer) * 10000 + Val100000(ActivePLayer) * 100000
	if Oldscore(ActivePLayer) <> TempScore Then
		Oldscore(ActivePLayer) = TempScore
		select Case ActivePlayer
			case 1:	EMReel1.setvalue TempScore
			case 2:	EMReel2.setvalue TempScore
			case 3:	EMReel3.setvalue TempScore
			case 4:	EMReel4.setvalue TempScore
		End Select
		If B2SOn Then
			Controller.B2SSetScorePlayer ActivePlayer, TempScore
		End If
		P_Reel1.image = cstr(Val10(ActivePLayer))
		P_Reel2.image = cstr(Val100(ActivePLayer))
		P_Reel3.image = cstr(Val1000(ActivePLayer))
		P_Reel4.image = cstr(Val10000(ActivePLayer))
		P_Reel5.image = cstr(Val100000(ActivePLayer))
	Else
		UpdateScoreReel.enabled = False
	end If
End Sub

Function ReelValue(ValPar,ScorPar,ChimePar)
	ReelValue = cint(ValPar)
	if ReelValue <> cint(ScorPar) Then
		if ChimesEnabled Then
			If ChimePar = 1 Then
				PlaySound"100a",0,ChimeVolume
			end If
			If ChimePar = 2 Then
				PlaySound "1000a",0,ChimeVolume
			end If
			If ChimePar = 3 Then
				PlaySound "10000a",0,ChimeVolume
			End If
			Else
			If ChimePar = 1 Then
				PlaySound "LadyLuckBeepLo2",0,ChimeVolume
			end If
			If ChimePar = 2 Then
				PlaySound "LadyLuckBeepMid2",0,ChimeVolume
			end If
			If ChimePar = 3 Then
				PlaySound "LadyLuckBeepHi2",0,ChimeVolume
			End If
		end If
		ReelValue = ReelValue + 1
		if ReelValue > 9 Then
			ReelValue = 0
		end If
	end If
End Function

Sub SetScoreReel
	TempScore = Playerscore(ActivePLayer)
	Score10 = 0
	Score100 = 0
	Score1000 = 0
	Score10000 = 0
	Score100000 = 0
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
	if len(TempScore) > 1 Then
		TempScore = Left(TempScore,len(TempScore)-1)
		Score100000 = cstr(right(Tempscore,1))
	end If
	P_Reel1.image = cstr(Score10)
	P_Reel2.image = cstr(Score100)
	P_Reel3.image = cstr(Score1000)
	P_Reel4.image = cstr(Score10000)
	P_Reel5.image = cstr(Score100000)
End Sub



' FS Post-It Playerscores
Sub UpdateScoreboard
	if not ShowPlayfieldReel or (NoofPlayers < 2) or Lady_Luck.ShowDT Then
		HideScoreBoard
	Else
		P_SB_Postit.image = "Postit"
		Select case NoofPlayers
			Case 2:
				P_SB_Postit.size_y = 50
				P_SB_Postit.Transy = -30
				PScore1_P.image = "Player1"
				PScore2_P.image = "Player2"
				SetScoreBoard(1)
				SetScoreBoard(2)
			Case 3:
				P_SB_Postit.size_y = 75
				P_SB_Postit.Transy = -15
				PScore1_P.image = "Player1"
				PScore2_P.image = "Player2"
				PScore3_P.image = "Player3"
				SetScoreBoard(1)
				SetScoreBoard(2)
				SetScoreBoard(3)
			Case 4:
				P_SB_Postit.size_y = 100
				P_SB_Postit.Transy = 0
				PScore1_P.image = "Player1"
				PScore2_P.image = "Player2"
				PScore3_P.image = "Player3"
				PScore4_P.image = "Player4"
				SetScoreBoard(1)
				SetScoreBoard(2)
				SetScoreBoard(3)
				SetScoreBoard(4)
		end Select
	end If
End Sub

Sub HideScoreboard
	for each obj in C_ScoreBoard
		obj.image = HSArray(10)
	Next
End Sub

Dim SBScore100k, SBScore10k, SBScoreK, SBScore100, SBScore10, SBScore1, SBTempScore,obj

Sub SetScoreBoard(PlayerPar)
	SBTempScore = PlayerScore(PlayerPar)
	SBScore1 = 0
	SBScore10 = 0
	SBScore100 = 0
	SBScoreK = 0
	SBScore10k = 0
	SBScore100k = 0
	if len(SBTempScore) > 0 Then
		SBScore1 = cint(right(SBTempscore,1))
	end If
	if len(SBTempScore) > 1 Then
		SBTempScore = Left(SBTempScore,len(SBTempScore)-1)
		SBScore10 = cint(right(SBTempscore,1))
	end If
	if len(SBTempScore) > 1 Then
		SBTempScore = Left(SBTempScore,len(SBTempScore)-1)
		SBScore100 = cint(right(SBTempscore,1))
	end If
	if len(SBTempScore) > 1 Then
		SBTempScore = Left(SBTempScore,len(SBTempScore)-1)
		SBScoreK = cint(right(SBTempscore,1))
	end If
	if len(SBTempScore) > 1 Then
		SBTempScore = Left(SBTempScore,len(SBTempScore)-1)
		SBScore10k = cint(right(SBTempscore,1))
	end If
	if len(SBTempScore) > 1 Then
		SBTempScore = Left(SBTempScore,len(SBTempScore)-1)
		SBScore100k = cint(right(SBTempscore,1))
	end If
	Select case PlayerPar
		Case 1:
			Pscore6_1.image = HSArray(SBScore100K):If PlayerScore(PlayerPar)<100000 Then Pscore6_1.image = HSArray(10)
			Pscore5_1.image = HSArray(SBScore10K):If PlayerScore(PlayerPar)<10000 Then Pscore5_1.image = HSArray(10)
			Pscore4_1.image = HSArray(SBScoreK):If PlayerScore(PlayerPar)<1000 Then Pscore4_1.image = HSArray(10)
			Pscore3_1.image = HSArray(SBScore100):If PlayerScore(PlayerPar)<100 Then Pscore3_1.image = HSArray(10)
			Pscore2_1.image = HSArray(SBScore10):If PlayerScore(PlayerPar)<10 Then Pscore2_1.image = HSArray(10)
			Pscore1_1.image = HSArray(SBScore1):If PlayerScore(PlayerPar)<1 Then Pscore1_1.image = HSArray(10)
			if PlayerScore(PlayerPar)<1000 then
				PscoreComma_1.image = HSArray(10)
			else
				PscoreComma_1.image = HSArray(11)
			end if
		Case 2:
			Pscore6_2.image = HSArray(SBScore100K):If PlayerScore(PlayerPar)<100000 Then Pscore6_2.image = HSArray(10)
			Pscore5_2.image = HSArray(SBScore10K):If PlayerScore(PlayerPar)<10000 Then Pscore5_2.image = HSArray(10)
			Pscore4_2.image = HSArray(SBScoreK):If PlayerScore(PlayerPar)<1000 Then PScore4_2.image = HSArray(10)
			Pscore3_2.image = HSArray(SBScore100):If PlayerScore(PlayerPar)<100 Then Pscore3_2.image = HSArray(10)
			Pscore2_2.image = HSArray(SBScore10):If PlayerScore(PlayerPar)<10 Then Pscore2_2.image = HSArray(10)
			Pscore1_2.image = HSArray(SBScore1):If PlayerScore(PlayerPar)<1 Then Pscore1_2.image = HSArray(10)
			if PlayerScore(PlayerPar)<1000 then
				PscoreComma_2.image = HSArray(10)
			else
				PscoreComma_2.image = HSArray(11)
			end if
		Case 3:
			Pscore6_3.image = HSArray(SBScore100K):If PlayerScore(PlayerPar)<100000 Then Pscore6_3.image = HSArray(10)
			Pscore5_3.image = HSArray(SBScore10K):If PlayerScore(PlayerPar)<10000 Then Pscore5_3.image = HSArray(10)
			Pscore4_3.image = HSArray(SBScoreK):If PlayerScore(PlayerPar)<1000 Then Pscore4_3.image = HSArray(10)
			Pscore3_3.image = HSArray(SBScore100):If PlayerScore(PlayerPar)<100 Then Pscore3_3.image = HSArray(10)
			Pscore2_3.image = HSArray(SBScore10):If PlayerScore(PlayerPar)<10 Then Pscore2_3.image = HSArray(10)
			Pscore1_3.image = HSArray(SBScore1):If PlayerScore(PlayerPar)<1 Then Pscore1_3.image = HSArray(10)
			if PlayerScore(PlayerPar)<1000 then
				PscoreComma_3.image = HSArray(10)
			else
				PscoreComma_3.image = HSArray(11)
			end if
		Case 4:
			Pscore6_4.image = HSArray(SBScore100K):If PlayerScore(PlayerPar)<100000 Then Pscore6_4.image = HSArray(10)
			Pscore5_4.image = HSArray(SBScore10K):If PlayerScore(PlayerPar)<10000 Then Pscore5_4.image = HSArray(10)
			Pscore4_4.image = HSArray(SBScoreK):If PlayerScore(PlayerPar)<1000 Then Pscore4_4.image = HSArray(10)
			Pscore3_4.image = HSArray(SBScore100):If PlayerScore(PlayerPar)<100 Then Pscore3_4.image = HSArray(10)
			Pscore2_4.image = HSArray(SBScore10):If PlayerScore(PlayerPar)<10 Then Pscore2_4.image = HSArray(10)
			Pscore1_4.image = HSArray(SBScore1):If PlayerScore(PlayerPar)<1 Then Pscore1_4.image = HSArray(10)
			if PlayerScore(PlayerPar)<1000 then
				PscoreComma_4.image = HSArray(10)
			else
				PscoreComma_4.image = HSArray(11)
			end if
	End Select
End Sub


' Keyboard Input

Sub Lady_Luck_KeyDown(ByVal keycode)
	If keycode = PlungerKey Then
		Plunger.PullBack
		PlaySoundAt "plungerpull", Plunger
	End If

	if GameStarted and (Tilt = 0) Then
		If keycode = LeftFlipperKey Then
			LeftFlipper.RotateToEnd
			PlaySoundAt"fx_FlipperUp", LeftFlipper
		End If

		If keycode = RightFlipperKey Then
			RightFlipper.RotateToEnd
			PlaySoundAt "fx_FlipperUp", RightFlipper
		End If
	end If

	If keycode = LeftTiltKey Then
		Nudge 90, 2
		checkNudge
	End If

	If keycode = RightTiltKey Then
		Nudge 270, 2
		checkNudge
	End If

	If keycode = CenterTiltKey Then
		Nudge 0, 2
		checkNudge
	End If

	If keycode = MechanicalTilt Then
		mechchecknudge
	End If


	if keycode = StartGameKey Then
		StartGame
	end If

	if keycode = AddCreditKey Then
		Playsound "fx_Coin"
		Credits = Credits + 1
		'DOF 126, DOFOn
		if Credits > 9 then
			Credits = 9
		end If
		if B2Son Then
			Controller.B2SSetCredits Credits
		end If
		P_Credits.image = cstr(Credits)
		EMReel_Credits.setvalue Credits
	end If

	if keycode = AddCreditKey2 Then
		Playsound "fx_Coin"
		Credits = Credits + 2
		'DOF 126, DOFOn
		if Credits > 9 then
			Credits = 9
		end If
		if B2Son Then
			Controller.B2SSetCredits Credits
		End If
		P_Credits.image = cstr(Credits)
		EMReel_Credits.setvalue Credits
	end If
End Sub

Sub Lady_Luck_KeyUp(ByVal keycode)
	If keycode = PlungerKey Then
		Plunger.Fire
		PlaySoundAt "plunger", Plunger
	End If

	If keycode = LeftFlipperKey Then
		LeftFlipper.RotateToStart
		if GameStarted and (Tilt = 0) Then
			PlaySoundAt "fx_flipperdown", LeftFlipper
		end If
	End If

	If keycode = RightFlipperKey Then
		RightFlipper.RotateToStart
		if GameStarted and (Tilt = 0) Then
			PlaySoundAt "fx_flipperdown", RightFlipper
		end If
	End If
End Sub

Sub AddSpecial
	playsound "knocker"
	Credits = Credits + 1
	'DOF 126, DOFOn
	if Credits > 9 then
		Credits = 9
	end If
	if B2Son Then
		Controller.B2SSetCredits Credits
	end If
	EMReel_Credits.setvalue Credits
End Sub

'--- Tilt recognition ---
Dim Tilt
Sub CheckNudge
	if GameActive then
		if NudgeTimer1.enabled then
			if NudgeTimer2.enabled then
				NudgeTimer1.enabled = False
				NudgeTimer2.enabled = False
				if Tilt = 0 then
					GameTilted
				end if
			else
				NudgeTimer2.enabled = True
			end if
		else
			NudgeTimer1.enabled = True
		end if
	end if
End Sub

Sub MechCheckNudge
	if GameActive then
		GameTilted
		LeftFlipper.RotateToStart
		RightFlipper.RotateToStart
	end if
End Sub


Sub NudgeTimer1_Timer
	NudgeTimer1.enabled = False
End Sub

Sub NudgeTimer2_Timer
	NudgeTimer2.enabled = False
End Sub

Sub GameTilted
	AdvanceBonusTimer.enabled = False
	Tilt = 1
	if B2SOn then
		Controller.B2SSetTilt 1
	end If

	TargetSound = "target"
	RolloverSound1 = "fx_sensor"
	OutlaneSound = "fx_sensor"

End Sub





'#############################################################
'#####                Lady Luck Scoring                  #####
'#############################################################


Dim BallinPlay


Dim HoleValue

'*****CAPTIVE BALL*****

	Kicker3.CreateBall
	Kicker3.Kick 70,1
	Kicker3.Enabled=False



'********************



'*****BUMPER****
Sub Bumper1_Hit:Addpoints(1000):PlaySoundAt "fx_Bumper1",Bumper1:End Sub
Sub Bumper2_Hit:Addpoints(1000):PlaySoundAt "fx_Bumper2",Bumper2:End Sub
'***************



'*****SWITCHES*****
Sub TriggerTopA_Hit:PlaySound RolloverSound1, 0, 1, AudioPan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall):Addpoints(500):ScoreMotor:LightTopA.State = 0:LightDTA.State = 1:LightA.State = 1:LightA1.State = 1:End Sub
Sub TriggerTopK_Hit:PlaySound RolloverSound1, 0, 1, AudioPan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall):Addpoints(500):ScoreMotor:LightTopK.State = 0:LightTopRightLane.State = 1:LightDTK.State = 1:End Sub
Sub TriggerTopQ_Hit:PlaySound RolloverSound1, 0, 1, AudioPan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall):Addpoints(500):ScoreMotor:LightTopQ.State = 0:LightMidRightLane.State = 1:LightDTQ.State = 1:End Sub
Sub TriggerTopJ_Hit:Playsound RolloverSound1, 0, 1, AudioPan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall):Addpoints(500):ScoreMotor:LightTopJ.State = 0:LightRightInLane.State = 1:LightDTJ.State = 1:End Sub
Sub Wall38_Hit:Addpoints(100):ScoreMotor:End Sub
Sub TriggerTopRight_Hit:Playsound RolloverSound1, 0, 1, AudioPan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall):Addpoints(5000):Addbonus:ScoreMotor:End Sub
Sub TriggerLeftOutlane_Hit:PlaySound OutlaneSound, 0, 1, AudioPan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall):Addpoints(10000):ScoreMotor:End Sub
Sub TriggerRightOutlane_Hit:PlaySound OutlaneSound, 0, 1, AudioPan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall):Addpoints(10000):ScoreMotor:End Sub

Sub TriggerMidRight_Hit()
	If LightTopRightLane.State=1 then
		Addpoints(1000)
		AddBonus
	Else
		Addpoints(1000)
		End If
		Playsound RolloverSound1, 0, 1, AudioPan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
End Sub

Sub TriggerRightLower_Hit()
    	If LightMidRightLane.State=1 then
			Addpoints(1000)
			AddBonus
		Else
			Addpoints(1000)
		End If
		Playsound RolloverSound1, 0, 1, AudioPan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
End Sub

Sub TriggerRightInlane_Hit()
	If LightRightInLane.State=1 then
		Addpoints(1000)
		AddBonus
	Else
		Addpoints(1000)
		End If
		Playsound RolloverSound1, 0, 1, AudioPan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
End Sub

Sub TriggerLeftInLane_Hit()
    	If LightLeftInlane.State=1 then
			Addpoints(1000)
			AddBonus
		Else
			Addpoints(1000)
		End If
		Playsound RolloverSound1, 0, 1, AudioPan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
End Sub



'**********************



'*****TARGETS*****

Sub Target2_Hit: Playsoundat "Target", Target2: Addpoints(1000):LightLeftInlane.State = 1:LightDT10.State = 1:Light10.State = 0:Light101.State = 0:ScoreMotor:End Sub

Sub TargetCBUpper_Hit()
	If LightTopDB.State=1 then
		Addpoints(50000)
		LightTopDB.State=0
		LightExtraBall.State=1
		LightBottomDB.State=1
	Else
		Addpoints(50000)
' Loserman code add - logic to detect if Extra Ball is lit and if so award it
		If LightExtraBall.state = 1 Then
			AwardExtraBall = 1
			if B2SOn Then
				controller.B2SSetShootAgain 1
			end If
			LightExtraBall.state = 0
		end if
	End If
		Playsound "LadyLuckBell1"
End Sub

Sub Target1_Hit()
	If LightDTA.State=1 then
		Addpoints(5000)
		AddBonus
	Else
		Addpoints(500)
		End If

End Sub
Sub Target3_Hit()
	If LightDTA.State=1 then
		Addpoints(5000)
		AddBonus
	Else
		Addpoints(500)
		End If

End Sub
Sub Target4_Hit()
	If LightDTK.State=1 then
		Addpoints(5000)
		AddBonus
	Else
		Addpoints(500)
		End If

End Sub
Sub Target5_Hit()
	If LightDTQ.State=1 then
		Addpoints(5000)
		AddBonus
	Else
		Addpoints(500)
		End If

End Sub
Sub Target6_Hit()
	If LightDTJ.State=1 then
		Addpoints(5000)
		AddBonus
	Else
		Addpoints(500)
		End If

End Sub
Sub Target7_Hit()
	If LightDT10.State=1 then
		Addpoints(5000)
		AddBonus
	Else
		Addpoints(500)
		End If

End Sub




'**********************



'*****ROLLOVER TARGETS AND LIGHTS*****

Sub Trigger10_Hit()
	If LightRollOver10.State=1 then
		LightRollOverJ.State=1:LightRollOver10.State=0:LightCB10.State=1
		Addpoints(100)
	Else
		LightRollOver10.State=0
		Addpoints(100)
		End If
	Playsound "solon", 0, 0.25, AudioPan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
End Sub

Sub TriggerJ_Hit()
	If LightRollOverJ.State=1 then
		LightRollOverQ.State=1:LightRollOverJ.State=0:LightCBJ.State=1
		Addpoints(100)
	Else
		LightRollOverJ.State=0
		Addpoints(100)
		End If
	Playsound "solon", 0, 0.25, AudioPan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
End Sub

Sub TriggerQ_Hit()
	If LightRollOverQ.State=1 then
		LightRollOverK.State=1:LightRollOverQ.State=0:LightCBQ.State=1
		Addpoints(100)
	Else
		LightRollOverQ.State=0
		Addpoints(100)
		End If
	Playsound "solon", 0, 0.25, AudioPan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
End Sub

Sub TriggerK_Hit()
	If LightRollOverK.State=1 then
		LightRollOverA.State=1:LightRollOverK.State=0:LightCBK.State=1
		Addpoints(100)
	Else
		LightRollOverK.State=0
		Addpoints(100)
		End If
	Playsound "solon", 0, 0.25, AudioPan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
End Sub

Sub TriggerA_Hit()
	If LightRollOverA.State=1 then
		LightRollOverA.State=1:LightRollOverK.State=0:LightCBA.State=1:LightTopDB.State=1
		Addpoints(100)
		End If
	Playsound "solon", 0, 0.25, AudioPan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
End Sub



'***********************************************************************



'***************IN GAME BONUS********************

Dim TotalBonus,DB
'DB = Double Bonus


Sub AddBonus
	if Tilt = 0 Then
		if not AdvanceBonusTimer.enabled Then
			if TotalBonus < 55 then
				TotalBonus = TotalBonus + 1
			end If
			UpdateBonusLights
		end If
	end If
End Sub

Dim TargetSound, RolloverSound1,OutlaneSound
Sub ScoreMotor
	if Tilt = 0 Then
		AdvanceBonusTimer.enabled = True
		TargetSound = "target"
		RolloverSound1 = "fx_sensor"
		OutlaneSound = "fx_sensor"
	end If
End Sub

Sub AdvanceBonusTimer_Timer
	AdvanceBonusTimer.enabled = False

End Sub


'*****AWARD BONUS*****************************************************


Dim BonusX
Sub AwardBonus
	if Tilt = 1 Then
		TotalBonus = 0
	end If
			if LightBottomDB.State = 1 Then
			DB = True
			end If
	BonusX = 1
	if DB then BonusX = 2
'	if TB then BonusX = 3
	AwardBonusTimer.enabled = True
End Sub

Sub AwardBonusTimer_Timer
	if TotalBonus > 0 Then
		Playsound "FD_BonusCount"
		Addpoints(10000 * BonusX)
		TotalBonus = TotalBonus - 1
		UpdateBonusLights
	Else
		AwardBonusTimer.enabled = False
		Playsound "FD_BonusEnd"
		EndBounsTimer1.enabled = True
	end If
End Sub

Sub UpdateBonusLights
		LightB1.state = ((TotalBonus mod 10) = 1) or TotalBonus > 54
		LightB2.state = ((TotalBonus mod 10) = 2) or TotalBonus > 53
		LightB3.state = ((TotalBonus mod 10) = 3) or TotalBonus > 51
		LightB4.state = ((TotalBonus mod 10) = 4) or TotalBonus > 48
		LightB5.state = ((TotalBonus mod 10) = 5) or TotalBonus > 44
		LightB6.state = ((TotalBonus mod 10) = 6) or TotalBonus > 39
		LightB7.state = ((TotalBonus mod 10) = 7) or TotalBonus > 33
		LightB8.state = ((TotalBonus mod 10) = 8) or TotalBonus > 26
		LightB9.state = ((TotalBonus mod 10) = 9) or TotalBonus > 18
		LightB10.state = TotalBonus > 9
End Sub

Sub EndBounsTimer1_Timer
	EndBounsTimer1.enabled = False
	if (ActivePlayer < NoofPlayers) Or (BallinPlay < BallsperGame) Then
		Playsound "FD_DrainwithBallrelease"
	end If
	EndBounsTimer2.enabled = True
End Sub

Sub EndBounsTimer2_Timer
	EndBounsTimer2.enabled = False
	NextBall
End Sub

'**************************************************************************************



'*****RESET PLAYFIELD*****
Sub ResetPlayfield

	'ResetLights
	LightCBA.state = Lightstateoff
	LightCBK.state = Lightstateoff
	LightCBQ.state = Lightstateoff
    LightCBJ.state = Lightstateoff
	LightCB10.state = Lightstateoff
	LightDTA.state = Lightstateoff
	LightDTK.state = Lightstateoff
	LightDTQ.state = Lightstateoff
	LightDTJ.state = Lightstateoff
	LightDT10.state = Lightstateoff
	LightTopA.state = Lightstateon
	LightTopK.state = Lightstateon
	LightTopQ.state = Lightstateon
    LightTopJ.state = Lightstateon
	LightA.state = Lightstateoff
	LightA1.state = Lightstateoff
	Light10.state = Lightstateon
	Light101.state = Lightstateon
	LightRollOverA.state = Lightstateoff
	LightRollOverK.state = Lightstateoff
	LightRollOverQ.state = Lightstateoff
	LightRollOverJ.state = Lightstateoff
	LightRollOver10.state = Lightstateon
	LightTopRightSpecial.state = Lightstateoff
	LightTopRightLane.state = Lightstateoff
	LightMidRightLane.state = Lightstateoff
	LightRightInLane.state = Lightstateoff
	LightBottomDB.state = Lightstateoff
	LightLeftInlane.state = Lightstateoff
	LightB1.state = Lightstateoff
	LightB2.state = Lightstateoff
	LightB3.state = Lightstateoff
	LightB4.state = Lightstateoff
	LightB5.state = Lightstateoff
	LightB6.state = Lightstateoff
	LightB7.state = Lightstateoff
	LightB8.state = Lightstateoff
	LightB9.state = Lightstateoff
	LightB10.state = Lightstateoff
	Light50K.state = Lightstateon
	Light50K2.state = Lightstateon
	LightTopDB.state = Lightstateoff
	LightExtraBall.state = Lightstateoff

	Tilt = 0
End Sub
'********************************




'*************************************************************************************


Dim PlayerScore(4),ActivePlayer,Special1(4),Special2(4),Special3(4)
Sub Addpoints(ScorePar)
	if Tilt = 0 Then
		Nopointsscored = False
		if not GameActive then
			GameActive = True
		end If
		if ScorePar < 50 Then
			Playerscore(ActivePLayer) = Playerscore(ActivePLayer) + ScorePar
		Else
			if not AdvanceBonusTimer.enabled Then
				Playerscore(ActivePLayer) = Playerscore(ActivePLayer) + ScorePar
			end If
		end If

		if (Playerscore(ActivePLayer) >= Special1Score) and not Special1(ActivePlayer) Then
			Special1(ActivePlayer) = True
			AddSpecial
		end If
		if (Playerscore(ActivePLayer) >= Special2Score) and not Special2(ActivePlayer) Then
			Special2(ActivePlayer) = True
			AddSpecial
		end If
		if (Playerscore(ActivePLayer) >= Special3Score) and not Special3(ActivePlayer) Then
			Special3(ActivePlayer) = True
			AddSpecial
		end If
		UpdateScoreReel.enabled = True
	end If
End Sub

Sub Flippertimer_Timer
	RFPrim.RotY = RightFlipper.currentangle
	LFPrim.RotY = LeftFlipper.currentangle
End Sub

Sub ShooterLaneLaunch_Hit
	if ActiveBall.vely < -8 then playsound "Launch",0,0.3,0.1,0.25
	'DOF 124, DOFPulse
End Sub





'---------------------------
'----- High Score Code -----
'---------------------------
Const HighScoreFilename = "Lady_Luck.txt"

Dim HSArray,HSAHighScore, HSA1, HSA2, HSA3
Dim HSScoreM,HSScore100k, HSScore10k, HSScoreK, HSScore100, HSScore10, HSScore1, HSScorex	'Define 6 different score values for each reel to use
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
	if B2SOn Then
		Controller.B2SSetScore 6,HSAHighScore
	End If
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





'**********SLING SHOTS************
'
' Rstep and Lstep  are the variables that increment the animation
'****************
Dim RStep, Lstep

Sub RightSlingShot_Slingshot
    PlaySoundAt "SlingshotRight", sling1
    Addpoints(100)
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
    PlaySound "SlingshotLeft", sling2
    Addpoints(100)
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

'******************************************

Sub Gate1_hit:Playsound "GateFlap", 0, 1, AudioPan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall) End Sub
Sub Gate7_hit:Playsound "GateFlap", 0, 1, AudioPan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall) End Sub



''*****************************************
''			FLIPPER SHADOWS
''*****************************************
'
sub FlipperTimer1_Timer()
	FlipperLSh.RotZ = LeftFlipper.currentangle
	FlipperRSh.RotZ = RightFlipper.currentangle
End Sub



'*********** BALL SHADOW *********************************
Dim BallShadow
BallShadow = Array (BallShadow1, BallShadow2)

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
        If BOT(b).X < Lady_Luck.Width/2 Then
            BallShadow(b).X = ((BOT(b).X) - (Ballsize/6) + ((BOT(b).X - (Lady_Luck.Width/2))/7)) + 10
        Else
            BallShadow(b).X = ((BOT(b).X) + (Ballsize/6) + ((BOT(b).X - (Lady_Luck.Width/2))/7)) - 10
        End If
        ballShadow(b).Y = BOT(b).Y + 10
        If BOT(b).Z > 20 Then
            BallShadow(b).visible = 1
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

Function AudioFade(tableobj) ' Fades between front and back of the table (for surround systems or 2x2 speakers, etc), depending on the Y position on the table. "Lady_luck" is the name of the table
  Dim tmp
  tmp = tableobj.y * 2 / Lady_luck.height-1
  If tmp > 0 Then
    AudioFade = Csng(tmp ^10)
  Else
    AudioFade = Csng(-((- tmp) ^10) )
  End If
End Function

Function AudioPan(tableobj) ' Calculates the pan for a tableobj based on the X position on the table. "Lady_luck" is the name of the table
  Dim tmp
  tmp = tableobj.x * 2 / Lady_luck.width-1
  If tmp > 0 Then
    AudioPan = Csng(tmp ^10)
  Else
    AudioPan = Csng(-((- tmp) ^10) )
  End If
End Function

Function Pan(ball) ' Calculates the pan for a ball based on the X position on the table. "Lady_luck" is the name of the table
    Dim tmp
    tmp = ball.x * 2 / Lady_luck.width-1
    If tmp > 0 Then
        Pan = Csng(tmp ^10)
    Else
        Pan = Csng(-((- tmp) ^10) )
    End If
End Function

Function AudioFade(ball) ' Can this be together with the above function ?
  Dim tmp
  tmp = ball.y * 2 / Lady_luck.height-1
  If tmp > 0 Then
    AudioFade = Csng(tmp ^10)
  Else
    AudioFade = Csng(-((- tmp) ^10) )
  End If
End Function

Function Vol(ball) ' Calculates the Volume of the sound based on the ball speed
  Vol = Csng(BallVel(ball) ^2 / 2000)
  Vol = Vol * RollingSoundFactor
End Function

Function Pitch(ball) ' Calculates the pitch of the sound based on the ball speed
  Pitch = BallVel(ball) * 20
End Function

Function BallVel(ball) 'Calculates the ball speed
  BallVel = INT(SQR((ball.VelX ^2) + (ball.VelY ^2) ) )
End Function


'*****************************************
'   rothbauerw's Manual Ball Control
'*****************************************

Dim BCup, BCdown, BCleft, BCright
Dim ControlBallInPlay, ControlActiveBall
Dim BCvel, BCyveloffset, BCboostmulti, BCboost

BCboost = 1				'Do Not Change - default setting
BCvel = 4				'Controls the speed of the ball movement
BCyveloffset = -0.01 	'Offsets the force of gravity to keep the ball from drifting vertically on the table, should be negative
BCboostmulti = 3		'Boost multiplier to ball veloctiy (toggled with the B key)

ControlBallInPlay = false

Sub StartBallControl_Hit()
	Set ControlActiveBall = ActiveBall
	ControlBallInPlay = true
End Sub

Sub StopBallControl_Hit()
	ControlBallInPlay = false
End Sub

Sub BallControlTimer_Timer()
	If EnableBallControl and ControlBallInPlay then
		If BCright = 1 Then
			ControlActiveBall.velx =  BCvel*BCboost
		ElseIf BCleft = 1 Then
			ControlActiveBall.velx = -BCvel*BCboost
		Else
			ControlActiveBall.velx = 0
		End If

		If BCup = 1 Then
			ControlActiveBall.vely = -BCvel*BCboost
		ElseIf BCdown = 1 Then
			ControlActiveBall.vely =  BCvel*BCboost
		Else
			ControlActiveBall.vely = bcyveloffset
		End If
	End If
End Sub


'*****************************************
'      JP's VP10 Rolling Sounds
'*****************************************

Const tnob = 2 ' total number of balls
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
  If Lady_luck.VersionMinor > 3 OR Lady_luck.VersionMajor > 10 Then
    PlaySound("fx_collide"), 0, Csng(velocity) ^2 / 200, Pan(ball1), 0, Pitch(ball1), 0, 0, AudioFade(ball1)
  Else
    PlaySound("fx_collide"), 0, Csng(velocity) ^2 / 200, Pan(ball1), 0, Pitch(ball1), 0, 0
  End if
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

' The sound is played using the VOL, AUDIOPAN, AUDIOFADE and PITCH functions, so the volume and pitch of the sound
' will change according to the ball speed, and the AUDIOPAN & AUDIOFADE functions will change the stereo position
' according to the position of the ball on the table.


'**************************************
' Explanation of the collision routine
'**************************************

' The collision is built in VP.
' You only need to add a Sub OnBallBallCollision(ball1, ball2, velocity) and when two balls collide they
' will call this routine. What you add in the sub is up to you. As an example is a simple Playsound with volume and paning
' depending of the speed of the collision.


Sub Pins_Hit (idx)
	PlaySound "pinhit_low", 0, Vol(ActiveBall), AudioPan(ActiveBall), 0, Pitch(ActiveBall), 0, 0, AudioFade(ActiveBall)
End Sub

Sub Targets_Hit (idx)
	PlaySound "target", 0, Vol(ActiveBall), AudioPan(ActiveBall), 0, Pitch(ActiveBall), 0, 0, AudioFade(ActiveBall)
End Sub

Sub Metals_Thin_Hit (idx)
	PlaySound "metalhit_thin", 0, Vol(ActiveBall), AudioPan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
End Sub

Sub Metals_Medium_Hit (idx)
	PlaySound "metalhit_medium", 0, Vol(ActiveBall), AudioPan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
End Sub

Sub Metals2_Hit (idx)
	PlaySound "metalhit2", 0, Vol(ActiveBall), AudioPan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
End Sub

Sub Gates_Hit (idx)
	PlaySound "gate4", 0, Vol(ActiveBall), AudioPan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
End Sub

Sub Spinner_Spin
	PlaySound "fx_spinner", 0, .25, AudioPan(Spinner), 0.25, 0, 0, 1, AudioFade(Spinner)
End Sub

Sub Rubbers_Hit(idx)
 	dim finalspeed
  	finalspeed=SQR(activeball.velx * activeball.velx + activeball.vely * activeball.vely)
 	If finalspeed > 20 then
		PlaySound "fx_rubber2", 0, Vol(ActiveBall), AudioPan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
	End if
	If finalspeed >= 6 AND finalspeed <= 20 then
 		RandomSoundRubber()
 	End If
End Sub

Sub Posts_Hit(idx)
 	dim finalspeed
  	finalspeed=SQR(activeball.velx * activeball.velx + activeball.vely * activeball.vely)
 	If finalspeed > 16 then
		PlaySound "fx_rubber2", 0, Vol(ActiveBall), AudioPan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
	End if
	If finalspeed >= 6 AND finalspeed <= 16 then
 		RandomSoundRubber()
 	End If
End Sub

Sub RandomSoundRubber()
	Select Case Int(Rnd*3)+1
		Case 1 : PlaySound "rubber_hit_1", 0, Vol(ActiveBall), AudioPan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
		Case 2 : PlaySound "rubber_hit_2", 0, Vol(ActiveBall), AudioPan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
		Case 3 : PlaySound "rubber_hit_3", 0, Vol(ActiveBall), AudioPan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
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
		Case 1 : PlaySound "flip_hit_1", 0, Vol(ActiveBall), AudioPan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
		Case 2 : PlaySound "flip_hit_2", 0, Vol(ActiveBall), AudioPan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
		Case 3 : PlaySound "flip_hit_3", 0, Vol(ActiveBall), AudioPan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
	End Select
End Sub

' Thalamus : Exit in a clean and proper way
Sub Lady_Luck_exit()
  Controller.Pause = False
  Controller.Stop
End Sub


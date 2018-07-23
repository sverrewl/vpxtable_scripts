'############################################################################################
'############################################################################################
'#######                                                                             ########
'#######                               Hot Shot                                      ########
'#######                           (Gottlieb 1973)                                   ########
'#######                                                                             ########
'############################################################################################
'############################################################################################
' Version 1.5 WED21 2017
'
' Thanks To:
' Zany for the Flipper model and texture
' JPSalas for the EM Reel images
' GTXJoe for the Highscore saving code
' mfuegemann for script help (YOU ARE A GENIUS!) and Fast Draw as template
' loserman78, GNance, and Mike Farmer for some images and other parts from their VP9 Hot Shot
' Flupper for the physics starting point
' Every other author for their amazing work!



'Version 1.2 Changes:
'Pretty much rebuilt everything.....


'Version 1.3 Changes:
'Fixed Tilt crash (Thanks Ganjafarmer)
'Fixed ball getting stuck below bumper and Middle Target area.
'May have fixed DOF Light issue (Thanks Thalamus)
'Fixed Ball Shadow (Thanks Ganjafarmer and rothbauerw)
'Put the correct plastic on Top Right section
'Added some missing shadows
'Added Flipper Shadows

'Version 1.4 Changes:
'Changed GI lights settings
'Changed environment settings
'Removed all DOF references
'Sounds encoded down to 16Bit to play on all computers
'Small Shadow fixes
'Added Drop Target Shadows
'Fixed Backglass credit issue (Thanks STAT)

'Version 1.5 Changes
'New Materials (Thanks again Hauntfreaks!)
'New Environment Lighting
'Table Sounds in full Surround
'All New Plastics from scan of when I owned this table (Forgot I had them!)
'Redrew the playfield
'Added option to use Cue Ball for ball image
'Removed over 200 lines of useless code!!
'Numerous other tweeks




option Explicit

On Error Resume Next
ExecuteGlobal GetTextFile("controller.vbs")
If Err Then MsgBox "Can't open controller.vbs"
On Error Goto 0

Const ballsize = 50
Const ballmass = 2
Const cGameName = "Hot_Shot_1973"

'---------------------------
'-----  Configuration  ----- 
'---------------------------
Const BallsperGame = 3				'Change to 3 or 5 balls

Const CustomBall = 0                'Change ball to Cue Ball. Change to 0 for regular ball 

Const ShowHighScore = True			'enable to show HighScore tape on Apron, a new HighScore will be saved anyway

Const ShowPlayfieldReel = False		'set to True to show Playfield EM-Reel and Post-It for Player Scores (instead of a B2S backglass)

Const FreePlay = False

Const ChimesEnabled = True			'You have been warned!

Const ChimeVolume = 1				'to limit ear bleeding set between 0 and 1

Const ResetHighscore = False		'enable to manually reset the Highscores

Const Special1Score = 62000			'set the 3 replay scores
Const Special2Score = 76000
Const Special3Score = 84000

'*****************************************************************************************************************


Dim GameActive,NoofPlayers,i,HighScore,Credits
Sub Hot_Shot_1973_Init
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

	if not Hot_Shot_1973.ShowDT then
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
		
	if not ShowPlayfieldReel or Hot_Shot_1973.ShowDT Then
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
	BumperWall1.isdropped = True
	SwitchA_2.isdropped = True
	SwitchB_2.isdropped = True
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
				if FreePlay and (Credits = 0) then 
					Credits = 5
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
				Controller.B2SSetMatch MatchValue
				Controller.B2SSetBallInPlay BallInPlay
				Controller.B2SSetCredits Credits
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
	if (ActivePlayer = NoofPlayers) and (BallinPlay = BallsperGame) Then
		EndGame
	Else
		UpdateScoreboard
		if ActivePlayer < NoofPlayers Then
			ActivePlayer = ActivePlayer + 1
		Else
			if ActivePlayer = NoofPlayers Then
				BallinPlay = BallinPlay + 1		
				ActivePlayer = 1
			end If
		end If
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

'********BONUS MULTIPLIER AND LIGHTS*******************
		ResetPlayfield
		BL1.state = lightstateon
		BL2.state = lightstateoff
		BL3.state = lightstateoff
		BonusX = 1
		if BallInPlay >= BallsperGame - 1 then
			BL1.state = lightstateoff
			if BallInPlay = BallsperGame then
				BL3.state = lightstateon
				BonusX = 3
			else
				BL2.state = lightstateon
				BonusX = 2
			end if
		end if
		if BallinPlay = BallsperGame Then
			DB = True
		end If
		NoPointsScored = True
		BallRelease.CreateBall
		BallRelease.Kick 90, 7
		PlaySoundAt "BallRelease", BallRelease
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
	PlaySoundAt "DrainLong", Drain
	Drain.DestroyBall
	if NoPointsScored and (Tilt = 0) Then
		Playsound "FD_DrainwithBallrelease"
		ReleaseSameBallTimer.enabled = True
	Else
		AwardBonus
	End If
End Sub

Sub ReleaseSameBallTimer_Timer
	ReleaseSameBallTimer.enabled = False
	BallRelease.CreateBall
	BallRelease.Kick 90, 7
	PlaySoundAt "BallRelease", BallRelease
End Sub

Sub MainTimer_Timer
	if not GameStarted Then
		LeftFlipper.RotateToStart
		RightFlipper.RotateToStart
	end If
	if Hot_Shot_1973.showdt Then		
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
		if Hot_Shot_1973.showdt Then
			P_GameOver.visible = not P_GameOver.visible
		End If
	End If
End Sub

Sub TiltTimer_Timer
	if Tilt = 0 Then	
		P_Tilt.visible = False
	Else	
		if Hot_Shot_1973.showdt Then
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
	
	Val10(ActivePLayer) = ReelValue(Val10(ActivePLayer),Score10,1)
	Val100(ActivePLayer) = ReelValue(Val100(ActivePLayer),Score100,2)
	Val1000(ActivePLayer) = ReelValue(Val1000(ActivePLayer),Score1000,3)
	Val10000(ActivePLayer) = ReelValue(Val10000(ActivePLayer),Score10000,0)
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
				PlaySound"10aBell",0,ChimeVolume
			end If
			If ChimePar = 2 Then 
				PlaySound"100aBell",0,ChimeVolume
			end If
			If ChimePar = 3 Then 
				PlaySound"1000aBell",0,ChimeVolume
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
	if not ShowPlayfieldReel or (NoofPlayers < 2) or Hot_Shot_1973.ShowDT Then		
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

Sub Hot_Shot_1973_KeyDown(ByVal keycode)
	If keycode = PlungerKey Then
		Plunger.PullBack
		PlaySoundAt "plungerpull", Plunger
	End If

	if GameStarted and (Tilt = 0) Then
		If keycode = LeftFlipperKey Then
			LeftFlipper.RotateToEnd
			PlaySoundAt "EMflipperup", LeftFlipper
			PlaySound "BuzzL", -1, 1, AudioPan(LeftFlipper), .05,0, 0, 1, AudioFade(LeftFlipper)
		End If
		
		If keycode = RightFlipperKey Then
			RightFlipper.RotateToEnd
			PlaySoundAt "EMflipperup", RightFlipper
			PlaySound "Buzz", -1, 1, AudioPan(RightFlipper), .05,0, 0, 1, AudioFade(RightFlipper)
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
    
	if keycode = StartGameKey Then
		StartGame
	end If

	if keycode = AddCreditKey Then
		PlaySoundAt "Coin", Drain
		Credits = Credits + 1
		if Credits > 9 then 
			Credits = 9
		end If
		if B2SOn Then
			Controller.B2SSetCredits Credits
		end If
		P_Credits.image = cstr(Credits)
		EMReel_Credits.setvalue Credits
	end If

	if keycode = AddCreditKey2 Then
		PlaySoundAt "Coin", Drain
		Credits = Credits + 2
		if Credits > 9 then 
			Credits = 9
		end If
		if B2SOn Then
			Controller.B2SSetCredits Credits
		End If
		P_Credits.image = cstr(Credits)
		EMReel_Credits.setvalue Credits
	end If
End Sub

Sub Hot_Shot_1973_KeyUp(ByVal keycode)
	If keycode = PlungerKey Then
		Plunger.Fire
		PlaySoundAt "plunger", Plunger
	End If
    
	If keycode = LeftFlipperKey Then
		LeftFlipper.RotateToStart
		if GameStarted and (Tilt = 0) Then
			PlaySoundAt "EMflipperdown", LeftFlipper
			Stopsound "BuzzL"
		end If
	End If
    
	If keycode = RightFlipperKey Then
		RightFlipper.RotateToStart
		if GameStarted and (Tilt = 0) Then
			PlaySoundAt "EMflipperdown", RightFlipper
			Stopsound "Buzz"
		end If
	End If
End Sub

Sub AddSpecial
	playsound"knocker"
	Credits = Credits + 1
	if Credits > 9 then 
		Credits = 9
	end If
	if B2SOn Then
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

Sub NudgeTimer1_Timer
	NudgeTimer1.enabled = False
End Sub

Sub NudgeTimer2_Timer
	NudgeTimer2.enabled = False
End Sub

Sub GameTilted
	AdvanceBonusTimer.enabled = False
	Tilt = 1	
	b1l.state = Lightstateoff
End Sub

'#############################################################
'#####                Hot Shot Scoring                   #####
'#############################################################

Dim BallinPlay
Dim HoleValue

'*****KICKER STUFF*****
Sub KickerBall8_Hit:Addpoints(500):ScoreMotor:Light8.State = 1
	PlaySoundAtVol "KickerEnter", KickerBall8, 0.5
	KickerBall8.timerenabled = False
	KickerBall8.timerenabled = True
	DiverterFlipper.RotateToEnd
	PlaySoundAt "SolOff", DiverterFlipper
End Sub	

Sub DiverterSwitch_Hit: DiverterFlipper.RotateToEnd:End Sub

Sub KickerBall8_Timer
	KickerBall8.timerenabled = False
	KickerBall8.kick 150,15,0.85
	P_Kicker.rotx = -85
	PlaySoundAt "PopperBall", KickerBall8
	Kickertimer.enabled = False
	Kickertimer.enabled = True
End Sub

Sub KickerTimer_Timer
	P_Kicker.rotx = P_Kicker.rotx + 5

	if P_Kicker.rotx >= -10 then
		P_Kicker.rotx = -10
		Kickertimer.enabled = False
	end if	
End Sub
'********************

'*****BUMPER****
Sub Bumper1_Hit:Addpoints(10):PlaySoundAt "Bumper1", Bumper1:End Sub
'***************

'*****SWITCHES*****
Sub RolloverRightWall_Hit:Addpoints(10):ScoreMotor:End Sub
Sub RolloverLeftWall_Hit:Addpoints(10):ScoreMotor:End Sub
Sub TriggerLeft1000_Hit:Addpoints(1000):ScoreMotor:End Sub
Sub TriggerLeftRollover_Hit: Addpoints(500):ScoreMotor:End Sub
Sub Trigger8Ball_Hit: Addpoints(500):ScoreMotor:Light8.State = 1:End Sub
Sub TriggerRight1000_Hit: Addpoints(1000):ScoreMotor:End Sub
Sub TriggerRightInlane_Hit: Addpoints(500):ScoreMotor:End Sub
Sub TriggerLeftInlane_Hit: Addpoints(500):ScoreMotor:End Sub
Sub Trigger_LeftOutlane_Hit: Addpoints(1000):ScoreMotor:End Sub
Sub Trigger_RightOutlane_Hit: Addpoints(1000):ScoreMotor:End Sub
Sub TargetRight1_Hit:Addpoints(100):ScoreMotor:End Sub
Sub TargetLeft1_Hit:Addpoints(100):ScoreMotor:End Sub
Sub TriggerRightRollover_Hit()
	If LightRightRollover.State=1 then
		Addpoints(1000)
	Else
		Addpoints(500)
		End If
		
End Sub
Sub TriggerLeftRollover_Hit()
    	If LightLeftRollover.State=1 then
			Addpoints(1000)
		Else
			Addpoints(500)
		End If
End Sub
Sub DropWallLeft_Hit:Addpoints(10):ScoreMotor:End Sub
Sub DropWallRight_Hit:Addpoints(10):ScoreMotor:End Sub
'**********************

'*****DROP TARGETS*****
Sub DropBall1_Hit: AddBonus:Addpoints(500):ScoreMotor:Light1.State = 1:dropshadow1.isdropped = 1:End Sub
Sub DropBall2_Hit:AddBonus:Addpoints(500):ScoreMotor:Light2.State = 1:dropshadow2.isdropped = 1:End Sub
Sub DropBall3_Hit:AddBonus:Addpoints(500):ScoreMotor:Light3.State = 1:dropshadow3.isdropped = 1:End Sub
Sub DropBall4_Hit:AddBonus:Addpoints(500):ScoreMotor:Light4.State = 1:dropshadow4.isdropped = 1:End Sub
Sub DropBall5_Hit:AddBonus:Addpoints(500):ScoreMotor:Light5.State = 1:dropshadow5.isdropped = 1:End Sub
Sub DropBall6_Hit:AddBonus:Addpoints(500):ScoreMotor:Light6.State = 1:dropshadow6.isdropped = 1:End Sub
Sub DropBall7_Hit:AddBonus:Addpoints(500):ScoreMotor:Light7.State = 1:dropshadow7.isdropped = 1:End Sub
Sub DropBall9_Hit:AddBonus:Addpoints(500):ScoreMotor:Light9.State = 1:dropshadow9.isdropped = 1:End Sub
Sub DropBall10_Hit:AddBonus:Addpoints(500):ScoreMotor:Light10.State = 1:dropshadow10.isdropped = 1:End Sub
Sub DropBall11_Hit:AddBonus:Addpoints(500):ScoreMotor:Light11.State = 1:dropshadow11.isdropped = 1:End Sub
Sub DropBall12_Hit:AddBonus:Addpoints(500):ScoreMotor:Light12.State = 1:dropshadow12.isdropped = 1:End Sub
Sub DropBall13_Hit:AddBonus:Addpoints(500):ScoreMotor:Light13.State = 1:dropshadow13.isdropped = 1:End Sub
Sub DropBall14_Hit:AddBonus:Addpoints(500):ScoreMotor:Light14.State = 1:dropshadow14.isdropped = 1:End Sub
Sub DropBall15_Hit:AddBonus:Addpoints(500):ScoreMotor:Light15.State = 1:dropshadow15.isdropped = 1:End Sub
'**********************

'*****ACTIVATING SIDE ROLLOVER LIGHTS*****
Sub CheckDTStateTimer_Timer
   if DropBall1.isdropped And DropBall2.isdropped And DropBall3.isdropped And DropBall4.isdropped And DropBall5.isdropped And DropBall6.isdropped And DropBall7.isdropped Then
				SpecialActivated = False
				RightSpecial = False
                LightLeftRollover.state = Lightstateon
		end If
		if DropBall9.isdropped And DropBall10.isdropped And DropBall11.isdropped And DropBall12.isdropped And DropBall13.isdropped And DropBall14. isdropped And DropBall15.isdropped Then
				SpecialActivated = False
				LeftSpecial = False
                LightRightRollover.state = Lightstateon
        end if
				if DropBall1.isdropped And DropBall2.isdropped And DropBall3.isdropped And DropBall4.isdropped And DropBall5.isdropped And DropBall6.isdropped And DropBall7.isdropped And DropBall9.isdropped And DropBall10.isdropped And DropBall11.isdropped And DropBall12.isdropped And DropBall13.isdropped And DropBall14. isdropped And DropBall15.isdropped Then
				SpecialActivated = True
				LeftSpecial = True
				LightRightTarget.state = Lightstateon
                LightLeftTarget.state = Lightstateon
				end if
End Sub
'***********************************************************************

'*****RESETTING SPECIAL AWARD*****
Sub ResetSpecialTimer_Timer
	ResetSpecialTimer.enabled = False	
	DropBall1.isdropped = True
	DropBall2.isdropped = True
	DropBall3.isdropped = True
    DropBall4.isdropped = True
	DropBall5.isdropped = True
	DropBall6.isdropped = True
	DropBall7.isdropped = True
	DropBall9.isdropped = True
	DropBall10.isdropped = True
    DropBall11.isdropped = True
	DropBall12.isdropped = True
	DropBall13.isdropped = True
	DropBall14.isdropped = True
    DropBall15.isdropped = True
End Sub
'********************************

Dim DTSpecial,LeftSpecial,RightSpecial,SpecialActivated

'*****RESET PLAYFIELD*****
Sub ResetPlayfield
	'ResetLights
	Light1.state = Lightstateoff
	Light2.state = Lightstateoff
	Light3.state = Lightstateoff
    Light4.state = Lightstateoff
	Light5.state = Lightstateoff
	Light6.state = Lightstateoff
	Light7.state = Lightstateoff
	Light8.state = Lightstateoff
	Light9.state = Lightstateoff
	Light10.state = Lightstateoff
	Light11.state = Lightstateoff
	Light12.state = Lightstateoff
	Light13.state = Lightstateoff
    Light14.state = Lightstateoff
	Light15.state = Lightstateoff

	'ResetLeftDT
	DropBall1.isdropped = False
	DropBall2.isdropped = False
	DropBall3.isdropped = False
	DropBall4.isdropped = False
	DropBall5.isdropped = False
    DropBall6.isdropped = False
	DropBall7.isdropped = False
	PlaySoundAt "DTReset", DropBall4

	'ResetRightDT
	DropBall6.isdropped = False
	DropBall7.isdropped = False
	DropBall9.isdropped = False
	DropBall10.isdropped = False
    DropBall11.isdropped = False
	DropBall12.isdropped = False
	DropBall13.isdropped = False
	DropBall14.isdropped = False
	DropBall15.isdropped = False
	PlaySoundAt "DTReset", DropBall11

	'Reset Drop Target Shadows
	dropshadow1.isdropped = 0
	dropshadow2.isdropped = 0
	dropshadow3.isdropped = 0
	dropshadow4.isdropped = 0
	dropshadow5.isdropped = 0
	dropshadow6.isdropped = 0
	dropshadow7.isdropped = 0
	dropshadow9.isdropped = 0
	dropshadow10.isdropped = 0
	dropshadow11.isdropped = 0
	dropshadow12.isdropped = 0
	dropshadow13.isdropped = 0
	dropshadow14.isdropped = 0
	dropshadow15.isdropped = 0

	'Reset Rollover Lights
	LightRightRollover.state = Lightstateoff
	LightLeftRollover.state = Lightstateoff
	LightRightTarget.state = Lightstateoff
	LightLeftTarget.state = Lightstateoff

	DB = False
	TB = False
	TotalBonus = 0
	DTSpecial = False
	LeftSpecial = False
	RightSpecial = False
	SpecialActivated = False
	Tilt = 0
	BumperWall1.isdropped = True
End Sub
'********************************

Dim TotalBonus,DB,TB
'DB = Double Bonus
'TB = Triple Bonus (DB on last Ball)

Sub AddBonus
	if Tilt = 0 Then
		if not AdvanceBonusTimer.enabled Then
			if TotalBonus < 15 then
				TotalBonus = TotalBonus + 1
			end If
'			UpdateBonusLights
		end If
	end If
End Sub

Sub ScoreMotor
	if Tilt = 0 Then
		AdvanceBonusTimer.enabled = True
	end If
End Sub

Sub AdvanceBonusTimer_Timer
	AdvanceBonusTimer.enabled = False
End Sub

'*****AWARD BONUS*****************************************************
Dim BonusX
Sub AwardBonus
 TotalBonus = 15
 if Tilt = 1 Then
  TotalBonus = 0
 end If
 AwardBonusTimer.enabled = True
End Sub
 
Sub AwardBonusTimer_Timer
 if TotalBonus > 0 Then
  Playsound "FD_BonusCount"
  select case TotalBonus
   Case 1: if Light1.state = 1 then Addpoints(1000 * BonusX):end if:Light1.state = 0
   Case 2: if Light2.state = 1 then Addpoints(1000 * BonusX):end if:Light2.state = 0
   Case 3: if Light3.state = 1 then Addpoints(1000 * BonusX):end if:Light3.state = 0
   Case 4: if Light4.state = 1 then Addpoints(1000 * BonusX):end if:Light4.state = 0
   Case 5: if Light5.state = 1 then Addpoints(1000 * BonusX):end if:Light5.state = 0
   Case 6: if Light6.state = 1 then Addpoints(1000 * BonusX):end if:Light6.state = 0
   Case 7: if Light7.state = 1 then Addpoints(1000 * BonusX):end if:Light7.state = 0
   Case 8: if Light8.state = 1 then Addpoints(1000 * BonusX):end if:Light8.state = 0
   Case 9: if Light9.state = 1 then Addpoints(1000 * BonusX):end if:Light9.state = 0
   Case 10: if Light10.state = 1 then Addpoints(1000 * BonusX):end if:Light10.state = 0
   Case 11: if Light11.state = 1 then Addpoints(1000 * BonusX):end if:Light11.state = 0
   Case 12: if Light12.state = 1 then Addpoints(1000 * BonusX):end if:Light12.state = 0
   Case 13: if Light13.state = 1 then Addpoints(1000 * BonusX):end if:Light13.state = 0
   Case 14: if Light14.state = 1 then Addpoints(1000 * BonusX):end if:Light14.state = 0
   Case 15: if Light15.state = 1 then Addpoints(1000 * BonusX):end if:Light15.state = 0
  End Select
  TotalBonus = TotalBonus - 1
 Else
  AwardBonusTimer.enabled = False
  Playsound "FD_BonusEnd"
  EndBounsTimer1.enabled = True
 end If
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
	diverter.RotY = DiverterFlipper.CurrentAngle+90
	FlipperLSh.RotZ = LeftFlipper.currentangle
	FlipperRSh.RotZ = RightFlipper.currentangle
End Sub

Sub ShooterLaneLaunch_Hit
	DiverterFlipper.RotateToStart
End Sub

'---------------------------
'----- High Score Code -----
'---------------------------
Const HighScoreFilename = "Hot_Shot_1973.txt"

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
Dim RStep, Lstep

Sub RightSlingShot_Slingshot
    PlaySoundAt "SlingshotRight", SLING1:Addpoints(10)
    RSling.Visible = 0
    RSling1.Visible = 1
    sling1.TransZ = -20
    RStep = 0
    RightSlingShot.TimerEnabled = 1
End Sub

Sub RightSlingShot_Timer
    Select Case RStep
        Case 3:RSLing1.Visible = 0:RSLing2.Visible = 1:sling1.TransZ = -10
        Case 4:RSLing2.Visible = 0:RSLing.Visible = 1:sling1.TransZ = 0:RightSlingShot.TimerEnabled = 0:gi1.State = 1:Gi2.State = 1
    End Select
    RStep = RStep + 1
End Sub

Sub LeftSlingShot_Slingshot
    PlaySoundAt "SlingshotLeft", sling2:Addpoints(10)
    LSling.Visible = 0
    LSling1.Visible = 1
    sling2.TransZ = -20
    LStep = 0
    LeftSlingShot.TimerEnabled = 1
End Sub

Sub LeftSlingShot_Timer
    Select Case LStep
        Case 3:LSLing1.Visible = 0:LSLing2.Visible = 1:sling2.TransZ = -10
        Case 4:LSLing2.Visible = 0:LSLing.Visible = 1:sling2.TransZ = 0:LeftSlingShot.TimerEnabled = 0:gi3.State = 1:Gi4.State = 1
    End Select
    LStep = LStep + 1
End Sub
'******************************************

Sub Gate1_hit:PlaySound "GateWire", 0, Vol(ActiveBall), AudioPan(Gate1), 0, Pitch(ActiveBall), 1, 0, AudioFade(Gate1): End Sub

'*********** BALL SHADOW *********************************
Dim BallShadow
BallShadow = Array (BallShadow1)

Sub BallShadowUpdate_timer()
    Dim BOT, b
    BOT = GetBalls
	'Custom Cue Ball Image
	If CustomBall=1 Then
	For b=0 to 1
	On Error Resume Next
	BOT(b).DecalMode=True
	BOT(b).image="BALL BLANK BLACK":BOT(b).FrontDecal="BALL RED DOT"
	BOT(b).BulbIntensityScale=1:BOT(b).PlayfieldReflectionScale=1
	Next
	End If
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
        If BOT(b).X < Hot_Shot_1973.Width/2 Then
            BallShadow(b).X = ((BOT(b).X) - (Ballsize/6) + ((BOT(b).X - (Hot_Shot_1973.Width/2))/7)) + 10
        Else
            BallShadow(b).X = ((BOT(b).X) + (Ballsize/6) + ((BOT(b).X - (Hot_Shot_1973.Width/2))/7)) - 10
        End If
        ballShadow(b).Y = BOT(b).Y + 10
        If BOT(b).Z > 20 Then
            BallShadow(b).visible = 1
        Else
            BallShadow(b).visible = 0
        End If
    Next
End Sub

'*********************************************************************
'                 Positional Sound Playback Functions
'*********************************************************************
' Play a sound, depending on the X,Y position of the table element (especially cool for surround speaker setups, otherwise stereo panning only)
' parameters (defaults): loopcount (1), volume (1), randompitch (0), pitch (0), useexisting (0), restart (1))
' Note that this will not work (currently) for walls/slingshots as these do not feature a simple, single X,Y position

Sub PlayXYSound(soundname, tableobj, loopcount, volume, randompitch, pitch, useexisting, restart)
	PlaySound soundname, loopcount, volume, AudioPan(tableobj), randompitch, pitch, useexisting, restart, AudioFade(tableobj)
End Sub

' Similar subroutines that are less complicated to use (e.g. simply use standard parameters for the PlaySound call)
Sub PlaySoundAt(soundname, tableobj)
    PlaySound soundname, 1, 1, AudioPan(tableobj), 0,0,0, 1, AudioFade(tableobj)
End Sub

Sub PlaySoundAtBall(soundname)
    PlaySoundAt soundname, ActiveBall
End Sub

'Set position as table object and Vol manually.

Sub PlaySoundAtVol(sound, tableobj, Vol)
		PlaySound sound, 1, Vol, AudioPan(tableobj), 0,0,0, 1, AudioFade(tableobj)
End Sub

'*********************************************************************
'                     Supporting Ball & Sound Functions
'*********************************************************************
Function AudioFade(tableobj) ' Fades between front and back of the table (for surround systems or 2x2 speakers, etc), depending on the Y position on the table. "Hot_Shot_1973" is the name of the table
	Dim tmp
    tmp = tableobj.y * 2 / Hot_Shot_1973.height-1
    If tmp > 0 Then
		AudioFade = Csng(tmp ^10)
    Else
        AudioFade = Csng(-((- tmp) ^10) )
    End If
End Function

Function AudioPan(tableobj) ' Calculates the pan for a tableobj based on the X position on the table. "Hot_Shot_1973" is the name of the table
    Dim tmp
    tmp = tableobj.x * 2 / Hot_Shot_1973.width-1
    If tmp > 0 Then
        AudioPan = Csng(tmp ^10)
    Else
        AudioPan = Csng(-((- tmp) ^10) )
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
Const tnob = 1 ' total number of balls
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
        If BallVel(BOT(b) ) > 1 and not BOT(b).z Then
			rolling(b) = True
			if BOT(b).z < 30 Then ' Ball on playfield
						PlaySound("fx_ballrolling" & b), -1, Vol(BOT(b) ), audioPan(BOT(b) ), 0, Pitch(BOT(b) ), 1, 0, AudioFade(BOT(b) )
			Else ' Ball on raised ramp
						PlaySound("fx_ballrolling" & b), -1, Vol(BOT(b) )*.5, audioPan(BOT(b) ), 0, Pitch(BOT(b) )+30000, 1, 0, AudioFade(BOT(b) )
				End If
		Else
            If rolling(b) = True Then
                StopSound("fx_ballrolling" & b)
                rolling(b) = False
            End If
        End If
    Next

	on error resume next ' In case VP is too old..
		' Kill ball spin
	if mLockMagnet.MagnetOn then
		dim rampballs:rampballs = mLockMagnet.Balls
		dim obj
		for each obj in rampballs
			obj.AngMomZ= 0
			obj.AngVelZ= 0
			obj.AngMomY= 0
			obj.AngVelY= 0
			obj.AngMomX= 0
			obj.AngVelX= 0
			obj.velx = 0
			obj.vely = 0
			obj.velz = 0
		next 
	end if
	on error goto 0
End Sub

'**********************
' Ball Collision Sound
'**********************
Sub OnBallBallCollision(ball1, ball2, velocity)
	PlaySound("fx_collide"), 0, Csng(velocity) ^2 / 2000, AudioPan(ball1), 0, Pitch(ball1), 0, 0, AudioFade(ball1)
End Sub

'**********************
' Misc Sounds
'**********************

Sub Targets_Hit (idx)
	PlaySound "target", 0, Vol(ActiveBall), AudioPan(ActiveBall), 0, Pitch(ActiveBall), 0, 0, AudioFade(ActiveBall)
End Sub

Sub Switches_Hit (idx)
	PlaySound "Sensor", 0, 1, AudioPan(ActiveBall), 0, Pitch(ActiveBall), 0, 0, AudioFade(ActiveBall)
End Sub

Sub DropTargets_Hit (idx)
	PlaySound "DTDrop", 0, 1, AudioPan(ActiveBall), 0, Pitch(ActiveBall), 0, 0, AudioFade(ActiveBall)
End Sub

Sub Metals_Thin_Hit (idx)
	PlaySound "metalhit_thin", 0, Vol(ActiveBall), AudioPan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
End Sub

Sub Metals_Medium_Hit (idx)
	PlaySound "metalhit_medium", 0, Vol(ActiveBall)/2, AudioPan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
End Sub

Sub Metals2_Hit (idx)
	PlaySound "metalhit2", 0, Vol(ActiveBall), AudioPan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
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

'*************END GAME**************
Sub Hot_Shot_1973_Exit()
	If B2SOn Then Controller.Stop
End Sub
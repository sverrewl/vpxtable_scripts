option explicit
'  Rogo -- Bally, 1974
'    VP 8 Release 1.0 			17 June 2002 		Table: Ash
'    VP 9 Release 1.0 			09 June 2011 		Table: humid
'	 Backglass image/DirectB2S 	23 November 2013 	Itchigo
'    VP X Release 1.0 			11 March 2017 		Table HSM		
'	 							 
' Change playfield image to "GreenPlayfield" and  CenterPlastic (layer 8) to "GreenHorseShoe" for Green color scheme.

BallsPerGame = 5

Const GameTitle = "Rogo"
Const MaxPlayers = 4
dim BallsPerGame
dim playerScore(4)
dim highScore
dim gameInProgress
dim credits
dim bonus
dim bonusTotal
dim ballInPlay
dim numberOfBalls
dim currentPlayer
dim numberOfPlayers
dim highScoreToDate
dim highScoreDisplayed
dim samePlayerShootsAgain
dim score					
dim freeGame(3,4)			
dim tilt					
dim ballOut					
dim newGame					
dim extraBallActive    		
dim specialActive   		
dim tiltDisabled       		
dim gameStart				
dim resetCount				
dim demoCount
dim drainHit
dim specialAwarded
dim extraBallAwarded
dim matchNumber
dim state
dim tiltsens
dim objekt

Const DoNewBall = "NewBall"

dim LightBonus(11)
set LightBonus(1) = LightBonus1
set LightBonus(2) = LightBonus2
set LightBonus(3) = LightBonus3
set LightBonus(4) = LightBonus4
set LightBonus(5) = LightBonus5
set LightBonus(6) = LightBonus6
set LightBonus(7) = LightBonus7
set LightBonus(8) = LightBonus8
set LightBonus(9) = LightBonus9
set LightBonus(10) = LightBonus10
set LightBonus(11) = LightBonus20

dim LightMisc(5)
set LightMisc(1) = LightSpecial
set LightMisc(2) = LightExtraBall
set LightMisc(3) = LightShootAgain
set LightMisc(4) = LightTopGate
set LightMisc(5) = LightBottomGate

dim EMReel(4)
set EMReel(1) = ScoreReel1
set EMReel(2) = ScoreReel2
set EMReel(3) = ScoreReel3
set EMReel(4) = ScoreReel4

dim Rollover(4)
set Rollover(1) = RolloverPlayer1
set Rollover(2) = RolloverPlayer2
set Rollover(3) = RolloverPlayer3
set Rollover(4) = RolloverPlayer4

dim LightPlayer(4)
set LightPlayer(1) = LightPlayer1
set LightPlayer(2) = LightPlayer2
set LightPlayer(3) = LightPlayer3
set LightPlayer(4) = LightPlayer4

Dim Ballsize, BallMass
BallSize = 50
BallMass = (Ballsize^3)/125000
credits=0

On Error Resume Next
ExecuteGlobal GetTextFile("controller.vbs")
If Err Then MsgBox "Can't open controller.vbs"
On Error Goto 0

Sub TableRogo_Init()
	LoadEM
	dim i,n
	Randomize
	numberOfBalls = (BallsPerGame)
	numberOfPlayers = 0
	for i = 1 to MaxPlayers
		playerScore(i) = 0
		TextBoxPlayer.text = ""
	next
	gameInProgress = false
	TextBoxMatch.Text = ""
	gameStart = true
	set score = new ScoringThing
	TextBoxGame.Text = "GAME"
	TextBoxOver.Text = "OVER"
	TextBoxBall.Text = ""
	TextBoxMatch.Text = ""
	LoadHighScore
	TextBoxHigh.text = "HIGH SCORE  "&(highScore)
	TextBoxCredits.setvalue(credits)
	currentPlayer = 1
	DisplayScore
	BonusTimer.enabled = false
    if credits > 0 then LightCredits.state = 1:end if
	If TableRogo.ShowDT = false then
		for each objekt in backdrop : objekt.visible = 0 : next
	End If
End Sub


sub setBackglass_timer
	Controller.B2ssetCredits credits
	Controller.B2ssetMatch 34, matchNumber
    Controller.B2SSetGameOver 35,1
    Controller.B2SSetScorePlayer 5, highScore
	me.enabled=0
end Sub

sub tilttxt_timer
	if state=false then
		tilttxt.text=""
		If B2SOn then Controller.B2SSetTilt 33,0	
		ttimer.enabled=true
	end if
		tilttxt.timerenabled=False
end sub

sub ttimer_timer
	if state=True then
		tilttxt.text="TILT"
		If B2SOn then Controller.B2SSetTilt 33,1
		tilttxt.timerenabled=1
		tilttxt.timerinterval= (INT (RND*10)+5)*100
	end if
	me.enabled=0
end sub

Sub TableRogo_KeyUp(ByVal keycode)
	If keycode = PlungerKey Then
		Plunger.Fire
		PlaySound "fx_plunger",0,1,0.25,0.25
	End If

	if tilt=false and state=true and gameInProgress = true then
	If keycode = LeftFlipperKey Then
		LeftFlipper.RotateToStart
		PlaySound SoundFXDOF("flipperdown",101,DOFOff,DOFContactors), 0, 1, 0.2, 0.05
		StopSound "Buzz"
	End If
    
	If keycode = RightFlipperKey Then
		RightFlipper.RotateToStart
		PlaySound SoundFXDOF("flipperdown",102,DOFOff,DOFContactors), 0, 1, -0.2, 0.05
		StopSound "Buzz"
	End If
   End if
    
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
 ' end if  
End Sub

Sub TableRogo_KeyDown(ByVal keycode)

	If keycode = PlungerKey Then
		Plunger.PullBack
		PlaySound "fx_plungerpull",0,1,0.25,0.25
	End If

	if tilt=false and state=true and gameInProgress = true then
		If keycode = LeftFlipperKey Then
			LeftFlipper.RotateToEnd
			PlaySound SoundFXDOF("flipperup",101,DOFOn,DOFContactors), 0, .67, 0.2, 0.05
			PlaySound "Buzz",-1,.05,0.5, 0.05
		End If
    
		If keycode = RightFlipperKey Then
			RightFlipper.RotateToEnd
			PlaySound SoundFXDOF("flipperup",102,DOFOn,DOFContactors), 0, .67, -0.2, 0.05
			PlaySound "Buzz",-1,.05,-0.5,0.05
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
End If
	 if keycode = 19 then 	'R
	 DisplayRuleSheet
	 End If
	if keycode=AddCreditKey then
		playsound "coinin" 
		coindelay.enabled=true 
		end if

	if keycode=StartGameKey and credits>0 then AddPlayer
End Sub

Sub AddPlayer
	if not gameInProgress then numberOfPlayers = 0
	if credits > 0 and numberOfPlayers < 1 and gameStart then
		state=True
		resetCount = 0
		TextBoxBall.text = 1
		PlaySound "StartGame"
		GameStartTimer.enabled = true
	end if

	if credits > 0 and numberOfPlayers < MaxPlayers then
		credits = credits-1
		TextBoxCredits.setvalue (credits)
		if credits > 0 then LightCredits.state = 1:end if
		If gameInProgress and ballInPlay = 1 then
			numberOfPlayers = numberOfPlayers + 1
			TextBoxPlayer.text = (numberOfPlayers)
		else
			newGame = true
			StartGame
			numberOfPlayers = 1
			TextBoxPlayer.text = "1"
			ballInPlay = 1
			NextPlayer
		end if
 	end if
	If B2SOn then
			Controller.B2ssetCredits Credits
			Controller.B2ssetcanplay 31, Numberofplayers
			Controller.B2ssetballinplay 32, Ballinplay
			Controller.B2ssetplayerup 30, currentPlayer
			Controller.B2SSetGameOver 0
			Controller.B2SSetScorePlayer 5, highScoreToDate
	End If

End Sub

sub coindelay_timer
	addcredits
    coindelay.enabled=false
end sub

Sub addcredits
      credits=credits+1
	  DOF 137, DOFOn
      if credits>25 then credits=25
	  if credits > 0 then LightCredits.state = 1:end if	
	  TextBoxCredits.setvalue (credits)
	  If B2SOn Then Controller.B2ssetCredits Credits
End sub

Sub StartGame
	dim i, j, k
	for i = 1 to MaxPlayers
		playerScore(i) = 0
		If B2SOn then Controller.B2SSetScoreRollover 24 + i, 0
	next
	gameInProgress = true
	tiltDisabled = 1
	currentPlayer = 0
	TextBoxGame.Text = ""
	TextBoxOver.Text = ""
	TextBoxMatch.Text = ""
	TextBoxHigh.text = ""
	TextShootAgain.text = ""
If B2SOn Then 
	Controller.B2SSetShootAgain 36,0
	Controller.B2ssetCredits credits
	Controller.B2ssetballinplay 32, Ballinplay
	Controller.B2ssetplayerup 30, 1
	Controller.B2ssetcanplay 31, 1
	Controller.B2SSetGameOver 0
	Controller.B2SSetScorePlayer 5, highScoreToDate
		End If
	specialActive = false
	specialAwarded = false
	extraBallActive = false
	extraBallAwarded = false
	drainHit = False
	samePlayerShootsAgain = false
End Sub

sub NewBall
	dim i
	score.ClearHasScored
	score.StartScoring
	TableDark
	if not newGame then 
	score.DoThisWhenDone DoNewBall, 600, "", 0
	end if
	newGame = false
	specialActive = false
	specialAwarded = false
	extraBallActive = false
	extraBallAwarded = false
	samePlayerShootsAgain = false
	drainHit = false
	Tilt = False
	bonus = 0
	EnableTable
	DoBonusLight bonus,LightStateOn
	LightPlayer(currentPlayer).text = (currentPlayer)
	If B2SOn then Controller.B2ssetplayerup 30, currentplayer
	TextShootAgain.text = ""
end sub

sub NextPlayer
	if currentPlayer <> 0 then
		LightPlayer(currentPlayer).text = ""
	end if

 	if not samePlayerShootsAgain then
		currentPlayer = currentPlayer + 1
		If B2SOn then Controller.B2ssetplayerup 30, currentPlayer
	else
		LightShootAgain.State = LightStateBlinking
	end if
	
	if currentPlayer > numberOfPlayers then
		ballInPlay = ballInPlay + 1
		currentPlayer = 1
	If B2SOn then Controller.B2ssetballinplay 32, Ballinplay
	end if

	if ballInPlay > numberOfBalls then
		GameOver
	else
		TextBoxBall.Text = (ballInPlay)
		Tilt=False
		tilttxt.text=""
		If B2SOn then Controller.B2SSetTilt 33,0
		If B2Son then Controller.B2SSetShootAgain 36,0
		TextShootAgain.text = ""
		NewBall
	end if
end sub

Sub EnableTable()
	BumperLeft.hashitevent = 1
	BumperRight.hashitevent = 1
	RightSlingshot.hashitevent = 1
	Leftslingshot.hashitevent = 1
End Sub

sub GameStartTimer_timer
	dim i
	resetCount = resetCount + 1
	for i = 1 to maxPlayers
		EMReel(i).UpdateInterval = 150
		EMReel(i).ResetToZero
		If B2SOn then Controller.B2SSetScorePlayer i, 0
	next
	if resetCount = 9 then
		gameStart = False
		score.DoThisWhenDone DoNewBall, 1600, "", 0  
		for i = 1 to maxPlayers
			freeGame(1,i) = 125000
			freeGame(2,i) = 174000
			freeGame(3,i) = 210000
			currentPlayer = i
			playerScore(currentPlayer) = 0
			DisplayScore
			EMReel(i).UpdateInterval = 50
		next
		GameStartTimer.enabled = false
		resetCount = 0
		currentPlayer= 1
If B2SOn then
		Controller.B2SSetGameOver 35,0
		Controller.B2SSetTilt 33,0
		Controller.B2SSetMatch 34,0
	End If
	end if
end sub

Sub Drain_Hit()
	PlaySound "Drain"
	Drain.DestroyBall
    tiltDisabled = 1
	drainHit = true
	if not gameInProgress then 
		GameOver
	elseif score.HasScored() then
		CollectBonus
	else
		samePlayerShootsAgain = true
		NextPlayer
	end if
End Sub

sub GameOver()
	dim i 
	If B2SOn then Controller.B2SSetGameOver 35,0
	LeftFlipper.RotateToStart
	RightFlipper.RotateToStart
	gameInProgress = false
	gameStart = True
	Match
	TextBoxGame.Text = "GAME"
	TextBoxOver.Text = "OVER"
	TextBoxBall.Text = ""
	TextShootAgain.text = ""
	TextBoxHigh.text = "HIGH SCORE TO DATE "&(highScore)
	TableDark
	for i = 1 to numberOfPlayers
		TextBoxPlayer.text = ""
		LightPlayer(i).text = ""
		if playerScore(i) > highScore then
			highScore = playerScore(i)
			TextBoxHigh.text = "NEW HIGH SCORE "&(highScore)
		end if
	next
	SaveHighScore
	If B2SOn then 
        Controller.B2SSetGameOver 35,1
        Controller.B2ssetballinplay 32, 0
	    Controller.B2SSetScorePlayer 5, highScore
	    Controller.B2ssetcanplay 31, 0
	    Controller.B2ssetcanplay 30, 0
		Controller.B2SSetShootAgain 36,0
	
	End If
    if credits = 0 then
    	LightCredits.state = LightStateOff
    end if
end sub


'********************** Slingsshots***************************
 Dim LStep, RStep
Sub LeftSlingShot_Slingshot
	If Tilt = True then Exit Sub
	score.AddScore currentPlayer, 10, 10, 10, "10", "LeftSlingshot"
    PlaySound "left_slingshot", 0, 1, -0.05, 0.05
    LeftSling4.Visible = 1
    Lemk.RotX = 26
    LStep = 0
    SwitchLights
    LeftSlingShot.TimerEnabled = 1
End Sub

Sub LeftSlingShot_Timer
    Select Case LStep
        Case 1:LeftSLing4.Visible = 0:LeftSLing3.Visible = 1:Lemk.RotX = 14
        Case 2:LeftSLing3.Visible = 0:LeftSLing2.Visible = 1:Lemk.RotX = 2
        Case 3:LeftSLing2.Visible = 0:Lemk.RotX = -10:LeftSlingShot.TimerEnabled = 0
    End Select

    LStep = LStep + 1
End Sub
 
Sub RightSlingShot_Slingshot
	If Tilt = True then Exit Sub
	score.AddScore currentPlayer, 10, 10, 10, "10", "RightSlingshot"
    PlaySound "right_slingshot", 0, 1, 0.05, 0.05
    RightSling4.Visible = 1
    Remk.RotX = 26
    RStep = 0
    SwitchLights
    RightSlingShot.TimerEnabled = 1
End Sub

Sub RightSlingShot_Timer
    Select Case RStep
        Case 1:RightSLing4.Visible = 0:RightSLing3.Visible = 1:Remk.RotX = 14
        Case 2:RightSLing3.Visible = 0:RightSLing2.Visible = 1:Remk.RotX = 2
        Case 3:RightSLing2.Visible = 0:Remk.RotX = -10:RightSlingShot.TimerEnabled = 0
    End Select
    RStep = RStep + 1
End Sub

Sub Wall8_Hit()
	If Tilt = True then Exit Sub
    score.Addscore currentPlayer, 10, 10, 10, "10", "RightSlingshot"
	SwitchLights
    PlaySound "10"
End Sub

Sub Wall61_Hit()
	If Tilt = True then Exit Sub
    score.Addscore currentPlayer, 10, 10, 10, "10", "LeftSlingshot"
	SwitchLights
    PlaySound "10"
End Sub

Sub Wall56_Hit()
	If Tilt = True then Exit Sub
    score.Addscore currentPlayer, 10, 10, 10, "10", "LeftSlingshot"
	SwitchLights
    PlaySound "10"
End Sub

Sub Wall107_Hit()
	If Tilt = True then Exit Sub
    score.Addscore currentPlayer, 10, 10, 10, "10", "RightSlingshot"
	SwitchLights
    PlaySound "10"
End Sub

Sub Wall115_Hit()
	If Tilt = True then Exit Sub
	score.AddScore currentPlayer, 10, 10, 10, "10", "LeftSlingshot"
	SwitchLights
	PlaySound "10"
End Sub

Sub Wall45_Hit()
	If Tilt = True then Exit Sub
	score.AddScore currentPlayer, 10, 10, 10, "10", "LeftSlingshot"
	SwitchLights
	PlaySound "10"
End Sub

'***************************Bumpers****************************************
Sub BumperLeft_Hit()
	If Tilt = True then Exit Sub
	if BumperLeftLight.State = LightStateOn then
		score.AddScore currentPlayer, 100, 100, 10, "100", "BumperLeft"
	else
		score.AddScore currentPlayer, 10, 10, 10, "10", "BumperLeft"
		SwitchLights
	end if
End Sub


Sub BumperRight_Hit()
	If Tilt = True then Exit Sub
	if BumperRightLight.State = LightStateOn then
		score.AddScore currentPlayer, 100, 100, 10, "100", "BumperRight"
	else
		score.AddScore currentPlayer, 10, 10, 10, "10", "BumperRight"
		SwitchLights
	end if
End Sub

Sub BumperLower_Hit()
	If Tilt = True then Exit Sub
	score.AddScore currentPlayer, 1000, 1000, 10, "1000", "BumperLower"
	AdvanceBonus
End Sub

Sub TriggerCenterLane_Hit()
	If Tilt = True then Exit Sub
	score.AddScore currentPlayer, 1000, 1000, 10, "1000", "TriggerCenterLane"
	BumperLeftLight.State = LightStateOn
	BumperRightLight.State = LightStateOn
	BumperLeftLight1.State = LightStateOn
	BumperRightLight1.State = LightStateOn
	LightBottomGate.state = LightStateOn
	LightTopGate.state = LightStateOn
	Wall3.isdropped=True
	Wall4.isdropped=True
	playsound "gate"
	playsound "gate"
	AdvanceBonus
End Sub

Sub TriggerLeftLoop_Hit()
	If Tilt = True then Exit Sub
	score.AddScore currentPlayer, 1000, 1000, 100, "1000", "TriggerLeftLoop"
	AdvanceBonus
End Sub

Sub TriggerRightLoop_Hit()
	If Tilt = True then Exit Sub
	score.AddScore currentPlayer, 1000, 1000, 100, "1000", "TriggerRightLoop"
	AdvanceBonus
End Sub

Sub TriggerTopLoop_Hit()
	If Tilt = True then Exit Sub
	if LightSpecial.state = LightStateOn then
		AddSpecial
	end if
	score.AddScore currentPlayer, 5000, 1000, 150, "+1000", "TriggerTopLoop"
	AdvanceBonus
	AdvanceBonus
	AdvanceBonus
End Sub

Sub TriggerSide1_Hit()
	if tilt = True then Exit Sub
	score.AddScore currentPlayer, 1000, 1000, 10, "1000", "TriggerSide1"
End Sub

Sub TriggerSide2_Hit()
	if tilt = True then Exit Sub
	score.AddScore currentPlayer, 1000, 1000, 10, "1000", "TriggerSide2"
End Sub

Sub TriggerSide3_Hit()
	if tilt=True then Exit Sub
	score.AddScore currentPlayer, 3000, 1000, 150, "+1000", "TriggerSide3"
	LightTopGate.state = LightStateOff
	LightBottomGate.state = LightStateOff
End Sub

Sub TriggerSide4_Hit()
	If Tilt = True then Exit Sub
	score.AddScore currentPlayer, 3000, 1000, 150, "+1000", "TriggerSide4"
End Sub

Sub Target1_Hit()
	If Tilt = True then Exit Sub
	score.AddScore currentPlayer, 1000, 1000, 10, "1000", "Target1"
	if LightExtraBall.state = LightStateOn then
		extraBallActive = false
		extraBallAwarded = true
		samePlayerShootsAgain = true
		If B2Son then Controller.B2SSetGameOver 36,1
		LightExtraBall.State = LightStateOff
		LightShootAgain.State = LightStateOn
		TextShootAgain.text = "SHOOT AGAIN"
	end if
	AdvanceBonus
End Sub

Sub Target2_Hit()
	If Tilt = True then Exit Sub
	score.AddScore currentPlayer, 1000, 1000, 10, "1000", "Target2"
	AdvanceBonus
End Sub

Sub TriggerRightOutlane_Hit()
	If Tilt = True then Exit Sub
	score.AddScore currentPlayer, 1000, 1000, 10, "1000", "TriggerRightInlane"
End Sub

Sub TriggerLeftOutlane_Hit()
	If Tilt = True then Exit Sub
	score.AddScore currentPlayer, 1000, 1000, 10, "1000", "TriggerLeftInlane"
End Sub

Sub SaucerLeft_Hit()
	SaucerTimer.interval = 1850
	SaucerTimer.enabled = true
	If Tilt = True then Exit Sub
	score.AddScore currentPlayer, 500, 100, 150, "+100", "SaucerLeft"
	BumperRightLight.State = LightStateOn
	BumperRightLight1.State = LightStateOn
	LightTopGate.state = LightStateOn
	Wall3.isdropped=True
	PlaySoundAt "NewBall", SaucerLeft
	AdvanceBonus
	AdvanceBonus
End Sub			

Sub SaucerRight_Hit()
	SaucerTimer.interval = 1850
	SaucerTimer.enabled = true
	If Tilt = True then Exit Sub
	score.AddScore currentPlayer, 500, 100, 150, "+100", "SaucerRight"
	BumperLeftLight.State = LightStateOn
	BumperLeftLight1.State = LightStateOn
	LightBottomGate.state = LightStateOn
	Wall4.isdropped=True
	PlaySoundAt "NewBall", SaucerLeft
	AdvanceBonus
	AdvanceBonus
End Sub			

Sub SaucerTimer_Timer()
	SaucerLeft.kick 155,10
	SaucerRight.kick 215,10
	Playsound "Saucer"
	SaucerTimer.enabled = false
End Sub

Sub SaucerBonus_Hit()
	SaucerBonusTimer.interval = ((bonus / 1000) * 200) + 200  
	SaucerBonusTimer.enabled = true
	CollectBonus
End Sub

Sub SaucerBonusTimer_Timer()
	SaucerBonus.kick 0,50
	SaucerBonusTimer.enabled = false
	Playsound "Saucer"
	bonus = 0
	DoBonusLight bonus, LightStateOn
End Sub
	
Sub SwitchLights
	if specialActive and not specialAwarded then
		if LightSpecial.State = LightStateOn then
			LightSpecial.State = LightStateOff
		else
			LightSpecial.State = LightStateOn
		end if
	end if
End Sub

Sub KickerNewBall_Hit()
	KickerNewBall.kick 60,5
End Sub

sub AdvanceBonus
'Add 1000 to the bonus
	if bonus < 20000 then
		DoBonusLight bonus, LightStateOff
		bonus = bonus + 1000
		DoBonusLight bonus,LightStateOn
		PlaySound "BonusLight"
		PlaySound "BonusLight"
	end if
	if bonus = 20000 and not specialActive and not specialAwarded then
		specialActive = true
		LightSpecial.state = LightStateOn
	end if
	if bonus = 12000 or bonus = 14000 or bonus = 16000 or bonus = 18000 or bonus = 20000 then
		if not extraBallAwarded then
			LightExtraBall.state = LightStateOn
		end if
	else
		LightExtraBall.state = LightStateOff
	end if
end sub

sub DoBonusLight(score, newState)
	dim ones
	ones = (score mod 10000) / 1000
	if ones <> 0 then LightBonus(ones).State = newState
	if score >= 10000 then
		LightBonus10.State = newState
	end if
	if score >= 20000 then
		LightBonus20.State = newState
		LightBonus10.State = LightStateOff
	end if
end sub

sub CollectBonus
	bonusTotal = bonus
	BonusTimer.enabled = true
end sub

sub BonusTimer_Timer()
	If tilt=true then 
		BonusTimer.enabled = false
		if drainHit = true then
		NextPlayer
		playsound "EndGame"
		end If
	Else
		if not score.IsScoring() then
		if bonus = 0 then
		BonusTimer.enabled = false
		if drainHit = true then
		NextPlayer
		playsound "EndGame"
		end if
	Else
		score.AddScore currentPlayer, 1000, 1000, 10, "1000", "Bonus"
		DoBonusLight bonus, LightStateOff
		bonus = bonus - 1000
		if bonus = 20000 then bonuspause.enabled = True
		if bonus = 15000 then bonuspause.enabled = True
		if bonus = 10000 then bonuspause.enabled = True
		if bonus = 5000 then bonuspause.enabled = True
		DoBonusLight bonus, LightStateOn
		PlaySound "BonusLight"
		end if
	end if
end If
end sub

Sub BonusPause_Timer
	BonusTimer.interval = 250
	bonuspause.enabled = False
	BonusTimer.interval = 150
End Sub

sub AddSpecial()
		PlaySound "Special"
		Credits = credits + 1
		specialActive = false
		specialAwarded = true
		LightSpecial.State = LightStateOff
end sub

Sub Match()
	dim matchNumber, i
	matchNumber = int(rnd(1)*10)*10
	TextBoxMatch.Text = ""
	If matchNumber = 0 then
		TextBoxMatch.Text = "00"
	    If B2SOn then Controller.B2SSetMatch 00
	Else
		TextBoxMatch.Text = matchNumber
		If B2SOn then Controller.B2SSetMatch 34,matchNumber'*10
	End If
	For i = 1 to numberOfPlayers
		If matchNumber = (playerScore(i) Mod 100) then AddSpecial
	Next
End Sub

Sub DisplayScore
	dim j
	for j = 1 to 3
		if playerScore(currentPlayer) > freeGame(j,currentPlayer) and freeGame(j,currentPlayer) <> 0 then
			AddSpecial
			freeGame(j,currentPlayer) = 0
		end if
	next
	if playerScore(currentPlayer) < 100000 then
		Rollover(currentPlayer).text = ""
		Exit Sub
	end if
	if playerScore(currentPlayer) > 100000 and Rollover(currentPlayer).text <> "100,000" then
		Playsound "OverTheTop"
		Rollover(currentPlayer).text = "100,000"
		If B2SOn then Controller.B2SSetScoreRollover 24 + currentplayer, 1
		Exit Sub
	end if
End Sub

'***** ScoringThing 
Const MaxQueueSize = 100 
Const ScorePause = -100001
Const ScoreDoSomething = -100002
Class ScoringThing
'ST encapsulates a lot of scoring issues. It is not a very pure object
'because it references things on the table directly (like certain lights
'and the score text). To try and keep things as object oriented as possible,
'I've isolated outside influences at the end of the class.
	dim players()		'Player number of this scoring item
	dim points()		'Point value of this scoring item, say 500
	dim steps()			'# of steps, so we can score 500 in 100 point chunks
	dim intervals()		'Time between each step
	dim obj()			'Object that made the request, for later reference
	dim soundName()		'Noise scoring event will make
	dim playSoundOnce()	'Whether the sound is played once or each time
	dim name()			'Name of playfield element that scored this
	dim queueStart		'Scores are held in a standard fifo queue
	dim queueEnd
	dim scored			'A flag, so we can track 0 point balls
	dim acceptNewScores	'Usually true, set to false when tilted

	Private Sub Class_Initialize
		redim players(MaxQueueSize), points(MaxQueueSize), steps(MaxQueueSize)
		redim intervals(MaxQueueSize), obj(MaxQueueSize), soundName(MaxQueueSize)
		redim playSoundOnce(MaxQueueSize), name(MaxQueueSize)
		queueStart = 0
		queueEnd = 0
		scored = false
		acceptNewScores = true
		StopTimer
   	End Sub

	Public Sub AddScore(player, value, step, time, sound, elementName)
	'Adds a scoring element to the queue.
	'Repeating sounds must come in with "+" as the first character.
		If not acceptNewScores then Exit Sub
		If (queueEnd + 1) Mod MaxQueueSize = queueStart Then
			MsgBox "Score queue filled!"
		Else
			name(queueEnd) = elementName
			players(queueEnd) = player
			points(queueEnd) = value
			steps(queueEnd) = step
			intervals(queueEnd) = time
			playSoundOnce(queueEnd) = (Left(sound, 1) <> "+")
			if playSoundOnce(queueEnd) then
				soundName(queueEnd) = sound
			else
				soundName(queueEnd) = Mid(sound, 2)
			end if
			
			If Not IsScoring() Then StartTimer time
			queueEnd = (queueEnd + 1) Mod MaxQueueSize
			scored = true
		End If
	End Sub
	
	Public Sub AddPause(time)
	    AddScore 0, ScorePause, 0, time, "Pause"
	End Sub
	
	Public Sub DoThisWhenDone(whatToDo, time, sound, value)
		obj(queueEnd) = whatToDo
	    AddScore 0, ScoreDoSomething, value, time, sound, "Action"
	End Sub

	Public Sub TimeToScore()
		if not acceptNewScores then exit sub
		If points(queueStart) > 0 Then
			AddToScore
			points(queueStart) = points(queueStart) - steps(queueStart)
		End If
		if points(queueStart) = ScoreDoSomething then
			DoIt
		End If
		If points(queueStart) <= 0 Then
			queueStart = (queueStart + 1) Mod MaxQueueSize
			If queueStart <> queueEnd Then
				StartTimer intervals(queueStart)
			Else
				StopTimer
			End If
		End If
	End Sub
	
	Public Sub ClearHasScored()
		scored = false
	End Sub
	
	Public Function HasScored()
		HasScored = scored
	End Function

	Public Sub StartScoring
		acceptNewScores = true
	End Sub
	
	Public Sub StopScoring
		acceptNewScores = false
	End Sub

	Public Sub ClearScoreQueue
		StopTimer
		queueStart = 0
		queueEnd = 0
	End Sub
	
	Public Function NameInQueue(elementName)
		dim i, done
		
		NameInQueue = false
		done = false
		i = queueStart
		while not done
			if queueStart = queueEnd then
				done = true
			else
				if name(i) = elementName then
					NameInQueue = true
					done = true
				end if
			end if
			i = (i + 1) mod MaxQueueSize
		wend
	End Function

	Private Sub StartTimer(time)
	    ScoreTimer.Interval = time
	    ScoreTimer.Enabled = True
	End Sub
	
	Private Sub StopTimer()
	    ScoreTimer.Enabled = False
	End Sub
	
	Public Function IsScoring()
	    IsScoring = ScoreTimer.Enabled
	End Function
	
	Private Sub AddToScore()
		playerScore(players(queueStart)) = playerScore(players(queueStart)) + steps(queueStart)
		EMReel(currentPlayer).addvalue(steps(queueStart))
		If B2SOn Then Controller.B2SSetScorePlayer currentPlayer, playerscore(currentplayer)
		if soundName(queueStart) <> "" then 
			PlaySound soundName(queueStart)
			if playSoundOnce(queueStart) then soundName(queueStart) = ""
		end if
		DisplayScore
	End Sub

	Private Sub DoIt()
		if soundName(queueStart) <> "" then PlaySound soundName(queueStart)
		select case obj(queueStart)
		  case DoNewBall
		    KickerNewBall.CreateBall
		    KickerNewBall.kick 60,5
		    tiltDisabled = 0
			PlaySound "PlungeBall"
			ballOut = false
		end select
	End Sub

End Class

Sub ScoreTimer_Timer()
	score.TimeToScore
End Sub

Sub CheckTilt()
	dim acceptNewScores
	If Tilttimer.Enabled = True Then 
	 TiltSens = TiltSens + 1
	 if TiltSens = 3 Then
	   Tilt = True
	   acceptNewScores = false
	   tilttxt.text="TILT"
       	If B2SOn Then Controller.B2SSetTilt 33,1
       	If B2SOn Then Controller.B2ssetdata 1, 0
	    playsound "tilt"
	TableDark
	DisableTable
	 End If
	Else
	 TiltSens = 0
	 Tilttimer.Enabled = True
	End If
End Sub

Sub Tilttimer_Timer()
	Tilttimer.Enabled = False
End Sub

Sub DisableTable()
		BumperLeft.hashitevent = 0
		BumperRight.hashitevent = 0
		RightSlingshot.hashitevent = 0
		Leftslingshot.hashitevent = 0
End Sub

Sub TableDark()
	dim i
	for i = 1 to 10
		LightBonus(i).State = LightStateOff
	next

	for i = 1 to 5
		LightMisc(i).State = LightStateOff
	next
	BumperLeftLight.State =  LightStateOff
	BumperRightLight.State =  LightStateOff
	BumperLeftLight1.State =  LightStateOff
	BumperRightLight1.State =  LightStateOff
	Wall3.isdropped=False
	Wall4.isdropped=False
	If credits = 0 then LightCredits.state = 0
End Sub

Sub SaveHighScore
	Dim FileObj
	Dim ScoreFile
	Set FileObj=CreateObject("Scripting.FileSystemObject")
	If Not FileObj.FolderExists(UserDirectory) then 
		Exit Sub
	End if
	Set ScoreFile=FileObj.CreateTextFile(UserDirectory & "Rogo.txt",True)
		'ScoreFile.WriteLine credits
		ScoreFile.WriteLine playerScore(1)
		ScoreFile.WriteLine playerScore(2)
		ScoreFile.WriteLine playerScore(3)
		ScoreFile.WriteLine playerScore(4)
		ScoreFile.WriteLine highScore
		ScoreFile.Close
	Set ScoreFile=Nothing
	Set FileObj=Nothing
end sub

Sub LoadHighScore
	dim n
	Dim FileObj
	Dim ScoreFile
	Dim TextStr
    dim temp1
    dim temp2
    dim temp3
    dim temp4
    dim temp5
    dim temp6
    Set FileObj=CreateObject("Scripting.FileSystemObject")
	If Not FileObj.FolderExists(UserDirectory) then 
		Exit Sub
	End if
	If Not FileObj.FileExists(UserDirectory & "Rogo.txt") then
		Exit Sub
	End if
	Set ScoreFile=FileObj.GetFile(UserDirectory & "Rogo.txt")
	Set TextStr=ScoreFile.OpenAsTextStream(1,0)
	Set TextStr=ScoreFile.OpenAsTextStream(1,0)
		If (TextStr.AtEndOfStream=True) then
			Exit Sub
		End if
		'temp1=TextStr.ReadLine
		temp2=Textstr.ReadLine 
		temp3=Textstr.ReadLine
		temp4=Textstr.ReadLine
		temp5=Textstr.ReadLine
		temp6=Textstr.ReadLine
		TextStr.Close
	    'credits = CDbl(temp1)
	    playerScore(1) = CDbl(temp2)
	    playerScore(2) = CDbl(temp3)
	    playerScore(3) = CDbl(temp4)
	    playerScore(4) = CDbl(temp5)
	    highScore = CDbl(temp6)
	    Set ScoreFile = Nothing
	    Set FileObj = Nothing
		for n = 1 to maxPlayers
			currentPlayer = n
 	        DisplayScore
			EMReel(currentPlayer).SetValue(playerScore(currentPlayer))
		next
End Sub

sub DisplayRuleSheet()
	dim t
	t = "Pop Bumpers score 10pts, 100 when lit."
	t = t+chr(13)&chr(13)
	t = t&"Slingshots score 10pts."
	t = t+chr(13)&chr(13)
	t = t&"Top Skill Lane scores 1000pts, advances bonus, lite bumpers, & open Gates."
	t = t+chr(13)&chr(13)
	t = t&"Top Left Saucer scores 500pts, advances bonus 2000, lites Right Bumper,"
	t = t+chr(13)&"& opens Valhalla Gate."
	t = t+chr(13)&chr(13)
	t = t&"Top Right Saucer scores 500pts, advances bonus 2000, lites Left Bumper,"
	t = t+chr(13)&"& opens Odin Gate."
	t = t+chr(13)&chr(13)
	t = t&"Upper standing Targets score 1000pts, advance bonus, & award Extra Ball when lit."
	t = t+chr(13)&chr(13)
	t = t&"Lower Bumper scores 1000pts & advances bonus."
	t = t+chr(13)&chr(13)
	t = t&"Left Saucer collects lit bonus value."
	t = t+chr(13)&chr(13)
	t = t&"Outlanes score 1000pts."
	t = t+chr(13)&chr(13)
	t = t&"Ball through Top Gate scores 8000pts, ball through Bottom Gate scores"
	t = t+chr(13)&"6000pts."
	t = t+chr(13)&chr(13)
	t = t&"Center island scores 7000pts, 7000 bonus per loop & awards"
	t = t+chr(13)&"Special when lit."
	t = t+chr(13)&chr(13)
	t = t&"Bonus value of 12000 activates Extra Ball light. Extra ball awards one ball."
	t = t+chr(13)&chr(13)
	t = t&"Bonus value of 20000 activates Special light. Special adds one replay."
	t = t+chr(13)&"Special light switches off/on upon each hit of a 10 pt target."
	t = t+chr(13)&chr(13)
	t = t&"Replay awarded at 125000 & 178000 pts."
	t = t+chr(13)&chr(13)
	t = t&"Tilt disqualfies ball in play."
	t = t+chr(13)&chr(13)
	t = t&"Match awards 1 credit."
	msgBox t, vbOKonly, "Rules and Scoring"
End Sub
' *********************************************************************
'                      Supporting Ball & Sound Functions
' *********************************************************************

Function Vol(ball) ' Calculates the Volume of the sound based on the ball speed
    Vol = Csng(BallVel(ball) ^2 / 1200)
End Function

Function Pan(ball) ' Calculates the pan for a ball based on the X position on the table. "table1" is the name of the table
    Dim tmp
    tmp = ball.x * 2 / TableRogo.width-1
    If tmp > 0 Then
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

'Sub RollingUpdate()
 Sub RollingTimer_Timer()
   Dim BOT, b, ballpitch
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
            If BOT(b).z < 30 Then
                ballpitch = Pitch(BOT(b) )
            Else
                ballpitch = Pitch(BOT(b) ) * 100
            End If
            rolling(b) = True
            PlaySound("fx_ballrolling" & b), -1, Vol(BOT(b) ), Pan(BOT(b) ), 0, ballpitch, 1, 0
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
'	PlaySound("fx_collide"), 0, Csng(velocity) ^2 / 2000, Pan(ball1), 0, Pitch(ball1), 0, 0
	PlaySound("fx_collide"), 0, Csng(velocity) ^2 / 2000, AudioPan(ball1), 0, Pitch(ball1), 0, 0, AudioFade(ball1)
End Sub



'*****************************************
'	ninuzzu's	FLIPPER SHADOWS
'*****************************************

sub FlipperTimer_Timer()
	FlipperLSh.RotZ = LeftFlipper.currentangle
	FlipperRSh.RotZ = RightFlipper.currentangle
End Sub

''*****************************************
''	ninuzzu's	BALL SHADOW
''*****************************************
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
        If BOT(b).X < TableRogo.Width/2 Then
            BallShadow(b).X = ((BOT(b).X) - (Ballsize/6) + ((BOT(b).X - (TableRogo.Width/2))/7)) + 6
        Else
            BallShadow(b).X = ((BOT(b).X) + (Ballsize/6) + ((BOT(b).X - (TableRogo.Width/2))/7)) - 6
        End If
        ballShadow(b).Y = BOT(b).Y + 12
        If BOT(b).Z > 20 Then
            BallShadow(b).visible = 1
        Else
            BallShadow(b).visible = 0
        End If
    Next
End Sub

'******************************
' Diverse Collection Hit Sounds
'******************************

Sub aBumpers_Hit (idx): PlaySound "fx_bumper":End Sub'SoundFX("fx_bumper", DOFContactors), 0, 1, pan(ActiveBall): End Sub 
Sub aRollovers_Hit(idx):PlaySound "fx_sensor", 0, Vol(ActiveBall), pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0:End Sub
Sub aRollovers_Hit(idx):PlaySound "fx_sensor", 0, Vol(ActiveBall), pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0:End Sub
Sub aGates_Hit(idx):PlaySound "fx_Gate", 0, Vol(ActiveBall), pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0:End Sub
Sub aMetals_Hit(idx):PlaySound "fx_MetalHit", 0, Vol(ActiveBall), pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0:End Sub
Sub aRubber_Bands_Hit(idx):PlaySound "fx_rubber_band", 0, Vol(ActiveBall), pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0:End Sub
Sub aRubbers_Hit(idx):PlaySound "fx_rubber_band", 0, Vol(ActiveBall), pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0:End Sub
	Sub aRubber_Posts_Hit(idx):PlaySound "fx_postrubber", 0, Vol(ActiveBall), pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0:End Sub
	Sub aRubber_Pins_Hit(idx):PlaySound "fx_postrubber", 0, Vol(ActiveBall), pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0:End Sub
	Sub aYellowPins_Hit(idx):PlaySound "fx_postrubber", 0, Vol(ActiveBall), pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0:End Sub
Sub aPlastics_Hit(idx):PlaySound "fx_PlasticHit", 0, Vol(ActiveBall), pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0:End Sub
Sub aWoods_Hit(idx):PlaySound "fx_Woodhit", 0, Vol(ActiveBall), pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0:End Sub

Sub LeftFlipper_Collide(parm)
 	RandomSoundFlipper()
End Sub

Sub RightFlipper_Collide(parm)
 	RandomSoundFlipper()
End Sub

Sub RandomSoundFlipper()
	Select Case Int(Rnd*3)+1
		Case 1 : PlaySound "flip_hit_1", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0
		Case 2 : PlaySound "flip_hit_2", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0
		Case 3 : PlaySound "flip_hit_3", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0
	End Select
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


'*********************************************************************
'                     Supporting Ball & Sound Functions
'*********************************************************************

Function AudioFade(tableobj) ' Fades between front and back of the table (for surround systems or 2x2 speakers, etc), depending on the Y position on the table. "table1" is the name of the table
	Dim tmp
    tmp = tableobj.y * 2 / TableRogo.height-1
    If tmp > 0 Then
		AudioFade = Csng(tmp ^10)
    Else
        AudioFade = Csng(-((- tmp) ^10) )
    End If
End Function

Function AudioPan(tableobj) ' Calculates the pan for a tableobj based on the X position on the table. "TableRogo" is the name of the table
    Dim tmp
    tmp = tableobj.x * 2 / TableRogo.width-1
    If tmp > 0 Then
        AudioPan = Csng(tmp ^10)
    Else
        AudioPan = Csng(-((- tmp) ^10) )
    End If
End Function


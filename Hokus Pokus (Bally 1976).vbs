'*
'*        Bally's Hokus Pokus (1976)
'*        Table primary build/scripted by Loserman76
'*        Table images by GNance
'*        Source photos for images by Wrd1972
'*
'*

option explicit
Randomize
ExecuteGlobal GetTextFile("core.vbs")

On Error Resume Next
ExecuteGlobal GetTextFile("Controller.vbs")
If Err Then MsgBox "Unable to open Controller.vbs. Ensure that it is in the scripts folder."
On Error Goto 0

Const cGameName = "hokuspokus_1976"

Const ShadowFlippersOn = true
Const ShadowBallOn = true

Const ShadowConfigFile = false




Dim Controller	' B2S
Dim B2SScore	' B2S Score Displayed
Const HSFileName="HokusPokus_76VPX.txt"
Const B2STableName="HokusPokus_1976"
Const LMEMTableConfig="LMEMTables.txt"
Const LMEMShadowConfig="LMEMShadows.txt"
Dim EnableBallShadow
Dim EnableFlipperShadow

'* this value adjusts score motor behavior - 0 allows you to continue scoring while the score motor is running - 1 sets score motor to behave more like a real EM
Const ScoreMotorAdjustment=1

'* this is a debug setting to use an older scoring routine vs a newer score routine - don't change this value
Const ScoreAdditionAdjustment=1

'* this controls whether you hear Gottlieb-ish chimes (0) or Zaccaria sounds (1) when scoring - personally recommend setting to 0 because the Zaccaria sounds are awful
'Const ChimesOn=0

dim ScoreChecker
dim CheckAllScores
dim sortscores(4)
dim sortplayers(4)

Dim ChimesOn
Dim B2SOn
Dim B2SFrameCounter
Dim BackglassBallFlagColor
Dim TextStr,TextStr2
Dim i
Dim obj
Dim BallsPerGame
Dim InProgress
Dim BallInPlay
Dim CreditsPerCoin
Dim Score100K(4)
Dim Score(4)
Dim ScoreDisplay(4)
Dim HighScorePaid(4)
Dim HighScore
Dim HighScoreReward
Dim BonusMultiplier
Dim Credits
Dim Match
Dim Replay1
Dim Replay2
Dim Replay3
Dim Replay4
Dim Replay1Paid(4)
Dim Replay2Paid(4)
Dim Replay3Paid(4)
Dim Replay4Paid(4)
Dim TableTilted
Dim TiltCount

Dim OperatorMenu

Dim TargetSetting
Dim BonusBooster
Dim BonusBoosterCounter
Dim BonusCounter
Dim HoleCounter

Dim ZeroNineCounter
Dim Ones
Dim Tens
Dim Hundreds
Dim Thousands

Dim Player
Dim Players

Dim rst
Dim bonuscountdown
Dim TempMultiCounter
dim TempPlayerup
dim RotatorTemp

Dim bump1
Dim bump2
Dim bump3
Dim bump4

Dim LastChime10
Dim LastChime100
Dim LastChime1000

Dim Score10
Dim Score100

Dim targettempscore
Dim SpecialLightCounter
Dim SpecialLightOption
Dim HorseshoeCounter
Dim DropTargetCounter

Dim LeftTargetCounter
Dim RightTargetCounter

Dim MotorRunning
Dim Replay1Table(15)
Dim Replay2Table(15)
Dim Replay3Table(15)
Dim Replay4Table(15)
Dim ReplayTableSet
Dim ReplayLevel
Dim ReplayTableMax

Dim BonusSpecialThreshold
Dim TargetLightsOn
Dim AdvanceLightCounter
dim bonustempcounter

dim raceRflag
dim raceAflag
dim raceCflag
dim raceEflag
dim race5flag

Dim LStep, LStep2, RStep, L2Step, R2Step, xx

Dim ReelCounter
Dim BallCounter
Dim BallReelAddStart(12)
Dim BallReelAddFrames(12)
Dim BallReelDropStart(12)
Dim BallReelDropFrames(12)

Dim EightLit

Dim X
Dim Y
Dim Z
Dim AddABall
Dim AddABallFrames
Dim DropABall
Dim DropABallFrames
Dim CurrentFrame
Dim BonusMotorRunning
Dim QueuedBonusAdds
Dim QueuedBonusDrops


Dim TempLightTracker

Dim TargetLeftFlag
Dim TargetCenterFlag
Dim TargetRightFlag

Dim TargetSequenceComplete

Dim SpecialLightsFlag

Dim AlternatingRelay

Dim ZeroToNineUnit

Dim Kicker1Hold,Kicker2Hold,Kicker3Hold

Dim mHole,mHole2, mHole3, mHole4, mHole5

Dim SpinPos,Spin,Count,Count1,Count2,Count3,Reset,VelX,VelY,BallSpeed,LitSpinner

Dim GoldBonusCounter,SilverBonusCounter

Dim ScoreMotorStepper

Dim GameOption,CircusSetting

Dim LeftSpinnerCounter

Dim bgpos, dooralreadyopen

Dim ABCDCounter

Sub Table1_Init()
	
	LoadEM
	LoadLMEMConfig2

	If Table1.ShowDT = false then
		For each obj in DesktopCrap
			obj.visible=False
		next
	End If

	OperatorMenuBackdrop.image = "PostitBL"
	For XOpt = 1 to MaxOption
		Eval("OperatorOption"&XOpt).image = "PostitBL"
	next
		
	For XOpt = 1 to 256
		Eval("Option"&XOpt).image = "PostItBL"
	next


	CircusSetting=0
	LeftSpinnerCounter=0
	Count=0
    Count1=0
    Count2=0
	Count3=0
    Reset=0
	ZeroToNineUnit=Int(Rnd*10)
	Kicker1Hold=0
	Kicker2Hold=0
	Kicker3Hold=0

	EightLit=0

	BallCounter=0
	ReelCounter=0
	AddABall=0
	DropABall=0
	HideOptions
	SetupReplayTables
	PlasticsOff
	BumpersOff
	OperatorMenu=0
	HighScore=0
	MotorRunning=0
	HighScoreReward=1
	Credits=0
	BallsPerGame=5
	ReplayLevel=1
	ChimesOn=0
	TargetSetting=0
	AlternatingRelay=1
	ZeroNineCounter=0
	SpecialLightOption=2
	BackglassBallFlagColor=1
	GameOption=1
	loadhs
	if HighScore=0 then HighScore=500000


	TableTilted=false
	
	Match=int(Rnd*10)
	MatchReel.SetValue((Match)+1)
	Match=Match*10

	CanPlayReel.SetValue(0)
	GameOverReel.SetValue(1)



	For each obj in PlayerHuds
		obj.SetValue(0)
	next
	For each obj in PlayerScores
		obj.ResetToZero
	next

	For each obj in CircusLights
		obj.state=0
	next

	For each obj in GoldBonus
		obj.state=0
	next

	For each obj in CircusTargetLights
		obj.state=0
	next
	for each obj in bottgate
		obj.isdropped=true
    next
    bgpos=6
    bottgate(bgpos).isdropped=false

	primgate.RotY=90
	ChimesOn=0
 	dooralreadyopen=0
	Replay1=Replay1Table(ReplayLevel)
	Replay2=Replay2Table(ReplayLevel)
	Replay3=Replay3Table(ReplayLevel)
	Replay4=Replay4Table(ReplayLevel)

	BonusCounter=0
	HoleCounter=0

	BallCard.image="BC"+FormatNumber(BallsPerGame,0)

	InstructCard.image="IC"+FormatNumber(BallsPerGame,0)
'	CardLight1.State = ABS(CardLight1.State-1)
	RefreshReplayCard

	CurrentFrame=0
	
	Bumper1Light.state=0
	Bumper2Light.state=0

	



	AdvanceLightCounter=0
	
	Players=0
	RotatorTemp=1
	InProgress=false
	TargetLightsOn=false


	ScoreText.text=HighScore


	If B2SOn Then

		if Match=0 then
			Controller.B2SSetMatch 100
		else
			Controller.B2SSetMatch Match
		end if
		Controller.B2SSetScoreRolloverPlayer1 0
		Controller.B2SSetScoreRolloverPlayer2 0
		Controller.B2SSetScoreRolloverPlayer3 0
		Controller.B2SSetScoreRolloverPlayer4 0

		
		Controller.B2SSetTilt 0
		Controller.B2SSetCredits Credits
		Controller.B2SSetGameOver 1
	End If

	for i=1 to 4
    player=i
		If B2SOn Then 
			Controller.B2SSetScorePlayer player, 0
		End If
	next
	bump1=1
	bump2=1
	InitPauser5.enabled=true
	If Credits > 0 Then 
		DOF 127, DOFOn
		CreditLight.state=1
	end if
		
End Sub

Sub Table1_exit()
	savehs
	SaveLMEMConfig
	SaveLMEMConfig2
	If B2SOn Then Controller.Stop
end sub



Sub Table1_KeyDown(ByVal keycode)

	' GNMOD
	if EnteringInitials then
		CollectInitials(keycode)
		exit sub
	end if


	if EnteringOptions then
		CollectOptions(keycode)
		exit sub
	end if



	If keycode = PlungerKey Then
		Plunger.PullBack
		PlungerPulled = 1
	End If

	if keycode = LeftFlipperKey and InProgress = false then
		OperatorMenuTimer.Enabled = true
	end if
	' END GNMOD

	If keycode = LeftFlipperKey and InProgress=true and TableTilted=false Then
		LeftFlipper.RotateToEnd
		PlaySound "buzzL",-1
		PlaySound SoundFXDOF("FlipperUp", 101, DOFOn, DOFFlippers)
	End If
    
	If keycode = RightFlipperKey  and InProgress=true and TableTilted=false Then
		RightFlipper.RotateToEnd

		PlaySound SoundFXDOF("FlipperUp", 102, DOFOn, DOFFlippers)
		PlaySound "buzz",-1
	End If
    
	If keycode = LeftTiltKey Then
		Nudge 90, 2
		TiltIt
	End If
    
	If keycode = RightTiltKey Then
		Nudge 270, 2
		TiltIt
	End If
    
	If keycode = CenterTiltKey Then
		Nudge 0, 2
		TiltIt
	End If

	If keycode = MechanicalTilt Then
		TiltCount=2
		TiltIt
	End If


	If keycode = AddCreditKey or keycode = 4 then
		If B2SOn Then
			'Controller.B2SSetScorePlayer6 HighScore
			
		End If

		playsound "coinin"
		CreditLight.state=1
		AddSpecial2
	end if

   if keycode = 5 then 
		playsound "coinin"
		AddSpecial2
		CreditLight.state=1
		keycode= StartGameKey
	end if


   if keycode = StartGameKey and Credits>0 and InProgress=true and Players>0 and Players<2 and BallInPlay<2 then
		Credits=Credits-1
		If Credits < 1 Then
			DOF 127, DOFOff
			CreditLight.state=0
		end if
		CreditsReel.SetValue(Credits)
		Players=Players+1
		CanPlayReel.SetValue(Players)
		playsound "BallyStartButtonPlayers2-4"
		If B2SOn Then
			Controller.B2SSetCanPlay Players
			If Players=2 Then
				Controller.B2SSetScoreRolloverPlayer2 0
			End If
			If Players=3 Then
				Controller.B2SSetScoreRolloverPlayer3 0
			End If
			If Players=4 Then
				Controller.B2SSetScoreRolloverPlayer4 0
			End If
			Controller.B2SSetCredits Credits
		End If
    end if

	if keycode=StartGameKey and Credits>0 and InProgress=false and Players=0 and EnteringOptions = 0 then
'GNMOD
		OperatorMenuTimer.Enabled = false
'END GNMOD
		Credits=Credits-1
		If Credits < 1 Then
			DOF 127, 0
			CreditLight.state=0
		end if
		CreditsReel.SetValue(Credits)
		Players=1
		CanPlayReel.SetValue(Players)
		MatchReel.SetValue(0)
		Player=1
		playsound "startup_norm"
		TempPlayerUp=Player
'		PlayerUpRotator.enabled=true
		rst=0
		BallInPlay=1
		InProgress=true
		resettimer.enabled=true
		BonusMultiplier=1
		TimerTitle1.enabled=0
		TimerTitle2.enabled=0
		TimerTitle3.enabled=0
		TimerTitle4.enabled=0
		If B2SOn Then
			Controller.B2SSetCanPlay Players
			Controller.B2SSetTilt 0
			Controller.B2SSetGameOver 0
			Controller.B2SSetMatch 0
			Controller.B2SSetCredits Credits
			
			Controller.B2SSetCanPlay 1
			Controller.B2SSetPlayerUp 1
			Controller.B2SSetData 81,1
			Controller.B2SSetBallInPlay BallInPlay
			Controller.B2SSetScoreRolloverPlayer1 0
			Controller.B2SSetShootAgain 0

			
		End If
		If Table1.ShowDT = True then
			For each obj in PlayerScores
'				obj.ResetToZero
				obj.Visible=true
			next
			For each obj in PlayerScoresOn
'				obj.ResetToZero
				obj.Visible=false
			next

			For each obj in PlayerHuds
				obj.SetValue(0)
			next
			For each obj in PlayerHUDScores
				obj.state=0
			next
			PlayerHuds(Player-1).SetValue(1)
			PlayerHUDScores(Player-1).state=1
			PlayerScores(Player-1).Visible=0
			PlayerScoresOn(Player-1).Visible=1
		end If
		GameOverReel.SetValue(0)


	end if


    
End Sub

Sub Table1_KeyUp(ByVal keycode)

	' GNMOD
	if EnteringInitials then
		exit sub
	end if

	If keycode = PlungerKey Then
		
		if PlungerPulled = 0 then
			exit sub
		end if
		
		PlaySound"plungerrelease"
		Plunger.Fire
	End If

	if keycode = LeftFlipperKey then
		OperatorMenuTimer.Enabled = false
	end if

	' END GNMOD  

	If keycode = LeftFlipperKey and InProgress=true and TableTilted=false Then
		LeftFlipper.RotateToStart
		PlaySound SoundFXDOF("FlipperDown", 101, DOFOff, DOFFlippers)
		StopSound "buzzL"
	End If
    
	If keycode = RightFlipperKey and InProgress=true and TableTilted=false Then
		RightFlipper.RotateToStart
		PlaySound SoundFXDOF("FlipperDown", 102, DOFOff, DOFFlippers)
        StopSound "buzz"
	End If

End Sub



Sub Drain_Hit()
	DOF 124, DOFPulse
	Drain.DestroyBall
	PlaySound "fx_drain"
	Pause4Bonustimer.enabled=true
		
End Sub

Sub Trigger0_Unhit()	
	DOF 126, 2
End Sub

Sub Pause4Bonustimer_timer
	Pause4Bonustimer.enabled=0
	ScoreGoldBonus
End Sub

Sub NewBonusHolder_timer
	if NewBonusTimer.enabled=0 then
		NewBonusHolder.enabled=0
		NextBallDelay.enabled=true
	end if

end sub

'***********************
'     Flipper Logos
'***********************

Sub UpdateFlipperLogos_Timer
'    LFLogo.ObjRotZ = LeftFlipper.CurrentAngle+90
'    RFlogo.ObjRotZ = RightFlipper.CurrentAngle+90
	PGate.Rotz = (Gate.CurrentAngle*.75) + 25
	PGate1.Rotz = (Gate1.CurrentAngle*.75) + 25
	FlipperLSh.RotZ = LeftFlipper.currentangle
	FlipperRSh.RotZ = RightFlipper.currentangle
End Sub

'***********************
' slingshots
'

Sub RightSlingShot_Slingshot
    PlaySound SoundFXDOF("right_slingshot", 104, DOFPulse, DOFContactors), 0, 1, 0.05, 0.05
    DOF 131, DOFPulse
	RSling0.Visible = 0
    RSling1.Visible = 1
    sling1.TransZ = -20
    RStep = 0
    RightSlingShot.TimerEnabled = 1
	AddScore 10
End Sub

Sub RightSlingShot_Timer
    Select Case RStep
        Case 3:RSLing1.Visible = 0:RSLing2.Visible = 1:sling1.TransZ = -10
        Case 4:RSLing2.Visible = 0:RSling0.Visible = 1:sling1.TransZ = 0:RightSlingShot.TimerEnabled = 0
    End Select
    RStep = RStep + 1
End Sub

Sub LeftSlingShot_Slingshot
    PlaySound SoundFXDOF("left_slingshot", 103, DOFPulse, DOFContactors),0,1,-0.05,0.05
	DOF 130, DOFPulse
    LSling0.Visible = 0
    LSling1.Visible = 1
    sling2.TransZ = -20
    LStep = 0
    LeftSlingShot.TimerEnabled = 1
	AddScore 10
End Sub

Sub LeftSlingShot_Timer
    Select Case LStep
        Case 3:LSLing1.Visible = 0:LSLing2.Visible = 1:sling2.TransZ = -10
        Case 4:LSLing2.Visible = 0:LSLing0.Visible = 1:sling2.TransZ = 0:LeftSlingShot.TimerEnabled = 0
    End Select
    LStep = LStep + 1
End Sub

'***********************************
' Walls
'***********************************

Sub RubberWallSwitches_Hit(idx)
	if TableTilted=false then
		AddScore(10)
	end if
end Sub


'***********************************
' Bumpers
'***********************************


Sub Bumper1_Hit
	If TableTilted=false then
		PlaySound SoundFXDOF("bumper1", 105, DOFPulse, DOFContactors)
		DOF 132, DOFPulse
		bump1 = 1
		AddScore(10)
    end if
    
End Sub

Sub Bumper2_Hit
	If TableTilted=false then
		PlaySound SoundFXDOF("bumper1", 106, DOFPulse, DOFContactors)
		DOF 133, DOFPulse
		bump2 = 1
		AddScore(10)
    end if
    
End Sub

'************************************
'  Rollover lanes
'************************************

Sub TriggerTop1_Hit
	If TableTilted=false then
		DOF 107, DOFPulse
		SetMotor(500)
		for i=0 to 3
			LetterA(i).state=0
		Next
		LetterA(3).state=1
		IncreaseGoldBonus
		CheckABCDSequence
	end if
end sub


Sub TriggerTop2_Hit
	If TableTilted=false then
		DOF 108, DOFPulse
		SetMotor(500)
		for i=0 to 3
			LetterB(i).state=0
		Next
		LetterB(3).state=1
		IncreaseGoldBonus		
		CheckABCDSequence
	end if
end sub

Sub TriggerTop3_Hit
	If TableTilted=false then
		DOF 109, DOFPulse
		SetMotor(500)
		for i=0 to 3
			LetterC(i).state=0
		Next
		LetterC(3).state=1
		IncreaseGoldBonus		
		CheckABCDSequence
	end if
end sub

Sub TriggerTop4_Hit
	If TableTilted=false then
		DOF 110, DOFPulse
		SetMotor(500)
		for i=0 to 3
			LetterD(i).state=0
		Next
		LetterD(3).state=1
		IncreaseGoldBonus		
		CheckABCDSequence
	end if
end sub

Sub TriggerLeftOutlane_Hit
	if TableTilted=false then
		DOF 111, DOFPulse
		AddScore(1000)
		
		If LeftSpecial.state=1 then
			AddSpecial
		end if
	end if
end sub

Sub TriggerLeftInlane_Hit
	if TableTilted=false then
		DOF 112, DOFPulse
		SetMotor(500)
		for i=0 to 3
			LetterB(i).state=0
		Next
		LetterB(3).state=1
		IncreaseGoldBonus		
		CheckABCDSequence
	end if
end sub

Sub TriggerRightOutlane_Hit
	if TableTilted=false then
		DOF 114, DOFPulse
		AddScore(1000)
		
		If RightSpecial.state=1 then
			AddSpecial
		end if
	end if
end sub

Sub TriggerRightInlane_Hit
	if TableTilted=false then
		DOF 113, DOFPulse
		SetMotor(500)
		for i=0 to 3
			LetterC(i).state=0
		Next
		LetterC(3).state=1
		IncreaseGoldBonus		
		CheckABCDSequence
	end if
end sub


'******************************************
' Spinners


Sub Spinner1_Spin()
	If TableTilted=false and MotorRunning=0 then
			DOF 122, DOFPulse
		if LeftSpinnerLight.state=1 then
			AddScore(100)
		else
			AddScore(10)
		end if
		
	end if

end Sub

Sub Spinner2_Spin()
	If TableTilted=false and MotorRunning=0 then
		DOF 123, DOFPulse
		if RightSpinnerLight.state=1 then
			AddScore(100)
		else
			AddScore(10)
		end if

	end if

end Sub


Sub Spinner3_Spin()
	If TableTilted=false and MotorRunning=0 then
		DOF 129, DOFPulse		
		if CenterSpinnerLight.state=1 then
			AddScore(100)
		else
			AddScore(10)
		end if

	end if

end Sub



'****************************************************
'*   Star Triggers
'****************************************************

Sub Trigger1_Hit
	If TableTilted=false then
		SetMotor(500)
		If TriggerLight1.state=1 then
			IncreaseGoldBonus
		end if
		
	end if
end sub

Sub Trigger2_Hit
	If TableTilted=false then
		SetMotor(500)
		for i=0 to 3
			LetterA(i).state=0
		Next
		LetterA(3).state=1
		IncreaseGoldBonus		
		CheckABCDSequence
	end if
end sub

Sub Trigger3_Hit
	If TableTilted=false then
		SetMotor(500)
		for i=0 to 3
			LetterD(i).state=0
		Next
		LetterD(3).state=1
		IncreaseGoldBonus		
		CheckABCDSequence
	end if
end sub

Sub Trigger4_Hit
	If TableTilted=false then
		SetMotor(500)
		for i=0 to 3
			LetterB(i).state=0
		Next
		LetterB(3).state=1
		IncreaseGoldBonus		
		CheckABCDSequence
	end if
end sub

Sub Trigger5_Hit
	If TableTilted=false then
		SetMotor(500)
		for i=0 to 3
			LetterC(i).state=0
		Next
		LetterC(3).state=1
		IncreaseGoldBonus		
		CheckABCDSequence
	end if
end sub


'****************************************************
'*   Stationary Targets
'****************************************************

Sub Target1_Hit
	If TableTilted=false then
		DOF 119, DOFPulse
		SetMotor(500)
		for i=0 to 3
			LetterA(i).state=0
		Next
		LetterA(3).state=1
		IncreaseGoldBonus		
		CheckABCDSequence
	end if
end sub	

Sub Target2_Hit
	If TableTilted=false then
		DOF 121, DOFPulse
		AddScore(1000)
		If TopCenterTargetLight.state=1 then
			if dooralreadyopen=0 then
				openg.enabled=true
			
			end if
		end if

	end if
end sub	



Sub Target5_Hit
	If TableTilted=false then
		DOF 120, DOFPulse
		SetMotor(500)
		for i=0 to 3
			LetterD(i).state=0
		Next
		LetterD(3).state=1
		IncreaseGoldBonus		
		CheckABCDSequence
	end if
end sub	


'****************************************************

Sub CheckABCDSequence
	
	if LowerCenterA.state=1 and LowerCenterB.state=1 and LowerCenterC.state=1 and LowerCenterD.state=1 then
		ABCDCounter=ABCDCounter+1
		if ABCDCounter=1 then
			DoubleBonus.state=1
			LeftSpinnerLight.state=1
			CenterSpinnerLight.state=1
			RightSpinnerLight.state=1
		end if
		if ABCDCounter=2 then
			ShootAgainLight.state=1
			ShootAgainReel.SetValue(1)
			If B2SOn then
				Controller.B2SSetShootAgain 1
			end if
			ABCDSpecialLight.state=1
			if AlternatingRelay=0 then
				LeftSpecial.state=1
				RightSpecial.state=0
			else
				LeftSpecial.state=0
				RightSpecial.state=1
			end if
		end if
		if ABCDCounter>2 then
			AddSpecial
		end if
		ResetABCD
	end if
end sub

Sub CheckCenterTargetSpecials
	
		
end sub



'**************************************

Sub AddSpecial()
	If ChimesOn=1 then
		PlaySound "Circus-BiriBiri"
	else
		PlaySound SoundFXDOF("knocker", 128, DOFPulse, DOFKnocker)
	end if
	Credits=Credits+1
	DOF 127, DOFOn
	CreditLight.state=1
	if Credits>36 then Credits=36
	If B2SOn Then
		Controller.B2SSetCredits Credits
	End If
	CreditsReel.SetValue(Credits)
End Sub

Sub AddSpecial2()
	PlaySound"click"
	Credits=Credits+1
	DOF 127, DOFOn
	CreditLight.state=1
	if Credits>36 then Credits=36
	If B2SOn Then
		Controller.B2SSetCredits Credits
	End If
	CreditsReel.SetValue(Credits)
End Sub

Sub CloseGateTrigger_Hit()
	if dooralreadyopen=1 then
		closeg.enabled=true

	end if
End Sub


 sub openg_timer
    bottgate(bgpos).isdropped=true
    bgpos=bgpos-1
    bottgate(bgpos).isdropped=false  
	primgate.RotY=30+(bgpos*10)
     if bgpos=0 then 
		playsound "postup"
		GateOpenLight.state=1
		openg.enabled=false
	
 		dooralreadyopen=1
	end if

 end sub 

sub closeg_timer
    closeg.interval=10
    bottgate(bgpos).isdropped=true
    bgpos=bgpos+1
    bottgate(bgpos).isdropped=false   
    primgate.RotY=30+(bgpos*10)
 	if bgpos=6 then
		GateOpenLight.state=0
 		closeg.enabled=false

		dooralreadyopen=0
 	end if
end sub   



Sub ToggleAlternatingRelay
	If GameOption=1 then
		if AlternatingRelay=0 then
			AlternatingRelay=1
			TriggerLight1.state=1
			TopCenterTargetLight.state=0
			If ABCDCounter>2 then
				LeftSpecial.state=1
				RightSpecial.state=0
			end if
		else
			AlternatingRelay=0
			TriggerLight1.state=0
			TopCenterTargetLight.state=1
			If ABCDCounter>2 then
				LeftSpecial.state=0
				RightSpecial.state=1
			end if
		end if
	end if
	

end sub

Sub ToggleRedBumper

	

	BumpersOn
		
end sub

Sub ResetABCD
	For i = 0 to 3
		LetterA(i).state=1
		LetterB(i).state=1
		LetterC(i).state=1
		LetterD(i).state=1
	Next
	LetterA(3).state=0
	LetterB(3).state=0
	LetterC(3).state=0
	LetterD(3).state=0

end sub


Sub ResetBallDrops

	For each obj in GoldBonus
		obj.state=0
	next
	For each obj in CircusLights
		obj.state=0
	next

	For each obj in CircusTargetLights
		obj.state=1
	next
	ResetABCD
	ABCDCounter=0
	DoubleBonus.state=0
	ABCDSpecialLight.state=0
	TargetSequenceComplete=0

	LeftSpecial.state=0
	RightSpecial.state=0
	
	HoleCounter=0
	SilverBonusCounter=0
	GoldBonusCounter=0
	If BallsPerGame=5 then
		LeftSpinnerLight.state=0
		RightSpinnerLight.state=0
		CenterSpinnerLight.state=0
	else
		LeftSpinnerLight.state=1
		RightSpinnerLight.state=1
		CenterSpinnerLight.state=1
	end if

	
End Sub



Sub LightsOut
	
	BonusCounter=0
	HoleCounter=0
	Bumper1Light.state=0
	Bumper2Light.state=0
	


end sub

Sub ResetBalls()
	StopSound"BallyBuzzer"
	RolloverReel.Setvalue(0)

	TempMultiCounter=BallsPerGame-BallInPlay

	ResetBallDrops
	BonusMultiplier=1
	TableTilted=false
	If B2SOn then
		Controller.B2SSetScoreRolloverPlayer1 0
		Controller.B2SSetScoreRolloverPlayer2 0
		Controller.B2SSetTilt 0
		Controller.B2SSetData 80+Player,1
	end if
	TiltReel.SetValue(0)
	GoldBonusCounter=1
	GoldBonus(GoldBonusCounter).state=1
	PlasticsOn
	'CreateBallID BallRelease
	Ballrelease.CreateSizedBall 25
    Ballrelease.Kick 40,7
	DOF 125, DOFPulse
	BallInPlayReel.SetValue(BallInPlay)
'	InstructCard.image="IC"+FormatNumber(BallsPerGame,0)+FormatNumber(TargetSetting,0)+FormatNumber(BallInPlay,0)
'	CardLight1.State = ABS(CardLight1.State-1)
	

End Sub




sub IncreaseGoldBonus
	if GoldBonusCounter<15 then
		If GoldBonusCounter<10 then
			GoldBonus(GoldBonusCounter).state=0
			GoldBonusCounter=GoldBonusCounter+1
			GoldBonus(GoldBonusCounter).state=1
		elseIf GoldBonusCounter>9 then
			GoldBonus(GoldBonusCounter-10).state=0
			GoldBonusCounter=GoldBonusCounter+1
			GoldBonus(10).state=1
			GoldBonus(GoldBonusCounter-10).state=1	
		end if
	end if
end sub

sub ScoreSilverBonus
	
	If DoubleBonus.state=0 then
		ScoreMotorStepper=0
		
	else
		ScoreMotorStepper=0
		
	end if
	CollectSilverBonus.enabled=1
end sub

sub CollectSilverBonus_timer
	if MotorRunning=1 then
		exit sub
	end if
	if GoldBonusCounter<1 then
		CollectSilverBonus.enabled=0
		GoldBonusCounter=1
		GoldBonus(GoldBonusCounter).state=1
	else
		if ScoreMotorStepper<5 then
'			If TableTilted=false then
'				SetMotor(5000)
'			end if
			If DoubleBonus.state=0 then
				if TableTilted=false then
					SetMotor(5000)
				end if
			else
				if TableTilted=false then
					AddScore(10000)
				end if
			end if
			GoldBonus(GoldBonusCounter).state=0
			GoldBonusCounter=GoldBonusCounter-1
			if GoldBonusCounter>=0 then
				GoldBonus(GoldBonusCounter).state=1
			end if
			ScoreMotorStepper=ScoreMotorStepper+1
		else
			If DoubleBonus.state=0 then
				ScoreMotorStepper=0
			else
				ScoreMotorStepper=0
				
			end if
		end if
	end if
end sub



sub ScoreGoldBonus
	ScoreMotorStepper=0
	CollectGoldBonus.interval=135
	CollectGoldBonus.enabled=1
end sub

sub CollectGoldBonus_timer
	if MotorRunning=1 then
		exit sub
	end if
	if GoldBonusCounter<1 then
		CollectGoldBonus.enabled=0
		NextBallDelay.enabled=true
	else
		If DoubleBonus.state=1 then
			Select case ScoreMotorStepper
				case 0,1,3,4:
					AddScore(1000)
					
				case 2,5:
					PlaySound"BallyClunk"
					If GoldBonusCounter>10 then
						GoldBonus(GoldBonusCounter-10).state=0
					else
						GoldBonus(GoldBonusCounter).state=0
						
					end if

					GoldBonusCounter=GoldBonusCounter-1
					if GoldBonusCounter>=0 then
						If GoldBonusCounter<=10 then
							GoldBonus(GoldBonusCounter).state=1
						elseif GoldBonusCounter>10 then
							GoldBonus(10).state=1
							GoldBonus(GoldBonusCounter-10).state=1
						end if
					end if
				case 6:


				case 7:


			end select
			ScoreMotorStepper=ScoreMotorStepper+1
			If ScoreMotorStepper>7 then
				ScoreMotorStepper=0
			end if
		else
			Select Case ScoreMotorStepper
				case 0,1,2,3,4:
					AddScore(1000)
					If GoldBonusCounter>10 then
						GoldBonus(GoldBonusCounter-10).state=0
					else
						GoldBonus(GoldBonusCounter).state=0
						
					end if
					GoldBonusCounter=GoldBonusCounter-1
					if GoldBonusCounter>=0 then
						If GoldBonusCounter<=10 then
							GoldBonus(GoldBonusCounter).state=1
						elseif GoldBonusCounter>10 then
							GoldBonus(10).state=1
							GoldBonus(GoldBonusCounter-10).state=1
						end if
					end if
				case 5:
					PlaySound"BallyClunk"

					
			end select
			ScoreMotorStepper=ScoreMotorStepper+1
			If ScoreMotorStepper>5 then
				ScoreMotorStepper=0
			end if
		end if
	end if
end sub			
						





sub OffRollovers
	

end sub



sub resettimer_timer
    rst=rst+1
	if rst>1 and rst<12 then
		ResetReelsToZero(1)
	end if  
    if rst=13 then
    playsound "StartBall1"
    end if
    if rst=14 then
    newgame
    resettimer.enabled=false
    end if
end sub

Sub ResetReelsToZero(reelzeroflag)
	dim d1(5)
	dim d2(5)
	dim scorestring1, scorestring2

	If reelzeroflag=1 then
		scorestring1=CStr(Score(1))
		scorestring2=CStr(Score(2))
		scorestring1=right("00000" & scorestring1,5)
		scorestring2=right("00000" & scorestring2,5)
		for i=0 to 4
			d1(i)=CInt(mid(scorestring1,i+1,1))
			d2(i)=CInt(mid(scorestring2,i+1,1))
		next
		for i=0 to 4
			if d1(i)>0 then 
				d1(i)=d1(i)+1
				if d1(i)>9 then d1(i)=0
			end if
			if d2(i)>0 then 
				d2(i)=d2(i)+1
				if d2(i)>9 then d2(i)=0
			end if

		next
		Score(1)=(d1(0)*10000) + (d1(1)*1000) + (d1(2)*100) + (d1(3)*10) + d1(4)
		Score(2)=(d2(0)*10000) + (d2(1)*1000) + (d2(2)*100) + (d2(3)*10) + d2(4)
		If B2SOn Then
			Controller.B2SSetScorePlayer 1, Score(1)
			Controller.B2SSetScorePlayer 2, Score(2)
		End If
		PlayerScores(0).SetValue(Score(1))
		PlayerScoresOn(0).SetValue(Score(1))
		PlayerScores(1).SetValue(Score(2))
		PlayerScoresOn(1).SetValue(Score(2))

		scorestring1=CStr(Score(3))
		scorestring2=CStr(Score(4))
		scorestring1=right("00000" & scorestring1,5)
		scorestring2=right("00000" & scorestring2,5)
		for i=0 to 4
			d1(i)=CInt(mid(scorestring1,i+1,1))
			d2(i)=CInt(mid(scorestring2,i+1,1))
		next
		for i=0 to 4
			if d1(i)>0 then 
				d1(i)=d1(i)+1
				if d1(i)>9 then d1(i)=0
			end if
			if d2(i)>0 then 
				d2(i)=d2(i)+1
				if d2(i)>9 then d2(i)=0
			end if

		next
		Score(3)=(d1(0)*10000) + (d1(1)*1000) + (d1(2)*100) + (d1(3)*10) + d1(4)
		Score(4)=(d2(0)*10000) + (d2(1)*1000) + (d2(2)*100) + (d2(3)*10) + d2(4)
		If B2SOn Then
			Controller.B2SSetScorePlayer 3, Score(3)
			Controller.B2SSetScorePlayer 4, Score(4)
		End If
		PlayerScores(2).SetValue(Score(3))
		PlayerScoresOn(2).SetValue(Score(3))
		PlayerScores(3).SetValue(Score(4))
		PlayerScoresOn(3).SetValue(Score(4))

	end if

end sub
	
	

sub NextBallDelay_timer()
	NextBallDelay.enabled=false
	nextball

end sub

sub newgame
	InProgress=true
	queuedscore=0
	for i = 1 to 4
		Score(i)=0
		Score100K(1)=0
		HighScorePaid(i)=false
		Replay1Paid(i)=false
		Replay2Paid(i)=false
		Replay3Paid(i)=false
	next
	If B2SOn Then
		Controller.B2SSetTilt 0
		Controller.B2SSetGameOver 0
		Controller.B2SSetMatch 0
'		Controller.B2SSetScorePlayer1 0
'		Controller.B2SSetScorePlayer2 0
'		Controller.B2SSetScorePlayer3 0
'		Controller.B2SSetScorePlayer4 0
		Controller.B2SSetBallInPlay BallInPlay
		
	End if

	Bumper1Light.state=1
	Bumper2Light.state=1
	AlternatingRelay=0
	ZeroNineCounter=0
	LeftSpinnerCounter=0
	BumpersOn
	BonusCounter=0
	BallCounter=0
	TargetLeftFlag=1
	TargetCenterFlag=1
	TargetRightFlag=1
	TargetSequenceComplete=0
	TriggerLight1.state=1

	AlternatingRelay=1
	ToggleAlternatingRelay
'	IncreaseBonus
'	ToggleBumper
	EightLit=1
	ResetBalls
End sub

sub nextball
	If B2SOn Then
		Controller.B2SSetTilt 0
		Controller.B2SSetData 81,0
		Controller.B2SSetData 82,0
		Controller.B2SSetData 83,0
		Controller.B2SSetData 84,0
	End If
	
	If ShootAgainLight.state=0 then
		Player=Player+1
	else
		ShootAgainLight.state=0
		ShootAgainReel.SetValue(0)
		if B2SOn then
			Controller.B2SSetShootAgain 0
		end if
	end if

	If Player>Players Then
		BallInPlay=BallInPlay+1
		If BallInPlay>BallsPerGame then
			PlaySound("MotorLeer")
			InProgress=false
			
			If B2SOn Then
				Controller.B2SSetGameOver 1
				Controller.B2SSetPlayerUp 0
				Controller.B2SSetBallInPlay 0
				Controller.B2SSetCanPlay 0
			End If
			For each obj in PlayerHuds
				obj.SetValue(0)
			next
			If Table1.ShowDT = True then
				For each obj in PlayerScores
					obj.visible=1
				Next
				For each obj in PlayerScoresOn
					obj.visible=0
				Next
				For each obj in PlayerHUDScores
					obj.state=0
				Next
			end If
			GameOverReel.SetValue(1)
'			InstructCard.image="IC"+FormatNumber(BallsPerGame,0)+FormatNumber(TargetSetting,0)+"0"
'			CardLight1.State = ABS(CardLight1.State-1)
			BallInPlayReel.SetValue(0)
'			CanPlayReel.SetValue(0)
			LeftFlipper.RotateToStart
			RightFlipper.RotateToStart
			LightsOut
			BumpersOff
			PlasticsOff
			checkmatch
			CheckHighScore
			Players=0


			TimerTitle1.enabled=1
			TimerTitle2.enabled=1
			TimerTitle3.enabled=1
			TimerTitle4.enabled=1
			TimerTitle4.enabled=1
			HighScoreTimer.interval=100
			HighScoreTimer.enabled=True
		Else
			Player=1
			If B2SOn Then
				Controller.B2SSetPlayerUp Player
				Controller.B2SSetBallInPlay BallInPlay

			End If
'			PlaySound("RotateThruPlayers")
			TempPlayerUp=Player
'			PlayerUpRotator.enabled=true
			PlayStartBall.enabled=true
			For each obj in PlayerHuds
				obj.SetValue(0)
			next
			If Table1.ShowDT = True then
				For each obj in PlayerScores
					obj.visible=1
				Next
				For each obj in PlayerScoresOn
					obj.visible=0
				Next
				For each obj in PlayerHUDScores
					obj.state=0
				Next
				PlayerHuds(Player-1).SetValue(1)
				PlayerHUDScores(Player-1).state=1
				PlayerScores(Player-1).visible=0
				PlayerScoresOn(Player-1).visible=1
			end If

			ResetBalls
		End If
	Else 
		If B2SOn Then
			Controller.B2SSetPlayerUp Player
			Controller.B2SSetBallInPlay BallInPlay
		End If
'		PlaySound("RotateThruPlayers")
		TempPlayerUp=Player
'		PlayerUpRotator.enabled=true
		PlayStartBall.enabled=true
		For each obj in PlayerHuds
			obj.SetValue(0)
		next
		If Table1.ShowDT = True then
			For each obj in PlayerScores
					obj.visible=1
			Next
			For each obj in PlayerScoresOn
					obj.visible=0
			Next
			For each obj in PlayerHUDScores
					obj.state=0
			Next
			PlayerHuds(Player-1).SetValue(1)
			PlayerHUDScores(Player-1).state=1
			PlayerScores(Player-1).visible=0
			PlayerScoresOn(Player-1).visible=1
		end If
		ResetBalls
	End If

End sub

sub CheckHighScore
	Dim playertops
		dim si
	dim sj
	dim stemp
	dim stempplayers
	for i=1 to 4
		sortscores(i)=0
		sortplayers(i)=0
	next
	playertops=0
	for i = 1 to Players
		sortscores(i)=Score(i)
		sortplayers(i)=i
	next

	for si = 1 to Players
		for sj = 1 to Players-1
			if sortscores(sj)>sortscores(sj+1) then
				stemp=sortscores(sj+1)
				stempplayers=sortplayers(sj+1)
				sortscores(sj+1)=sortscores(sj)
				sortplayers(sj+1)=sortplayers(sj)
				sortscores(sj)=stemp
				sortplayers(sj)=stempplayers
			end if
		next
	next
	ScoreChecker=4
	CheckAllScores=1
	NewHighScore sortscores(ScoreChecker),sortplayers(ScoreChecker)
	savehs
end sub


sub checkmatch
	Dim tempmatch
	tempmatch=Int(Rnd*10)
	Match=tempmatch*10
	MatchReel.SetValue(tempmatch+1)
	
	If B2SOn Then
		If Match = 0 Then
			Controller.B2SSetMatch 100
		Else
			Controller.B2SSetMatch Match
		End If
	End if
	for i = 1 to Players
		if (Match*10)=(Score(i) mod 100) then
			AddSpecial
		end if
	next
end sub

Sub TiltTimer_Timer()
	if TiltCount > 0 then TiltCount = TiltCount - 1
	if TiltCount = 0 then
		TiltTimer.Enabled = False
	end if
end sub

Sub TiltIt()
		TiltCount = TiltCount + 1
		if TiltCount = 3 then
			TableTilted=True
			PlasticsOff
			BumpersOff
			LeftFlipper.RotateToStart
			RightFlipper.RotateToStart
			StopSound"buzz"
			StopSound"buzzL"
			TiltReel.SetValue(1)
			If B2Son then
				Controller.B2SSetTilt 1
			end if
		else
			TiltTimer.Interval = 500
			TiltTimer.Enabled = True
		end if
	
end sub



Sub PlayStartBall_timer()

	PlayStartBall.enabled=false
	PlaySound("StartBall1")
end sub

Sub PlayerUpRotator_timer()
		If RotatorTemp<5 then
			TempPlayerUp=TempPlayerUp+1
			If TempPlayerUp>4 then
				TempPlayerUp=1
			end if
			If B2SOn Then
				Controller.B2SSetPlayerUp TempPlayerUp
			End If

		else
			if B2SOn then
				Controller.B2SSetPlayerUp Player
			end if
			PlayerUpRotator.enabled=false
			RotatorTemp=1
		end if
		RotatorTemp=RotatorTemp+1

	
end sub

sub savehs
	' Based on Black's Highscore routines
	Dim FileObj
	Dim ScoreFile
	Set FileObj=CreateObject("Scripting.FileSystemObject")
	If Not FileObj.FolderExists(UserDirectory) then 
		Exit Sub
	End if
	Set ScoreFile=FileObj.CreateTextFile(UserDirectory & HSFileName,True)
		ScoreFile.WriteLine 0
		ScoreFile.WriteLine Credits
		scorefile.writeline BallsPerGame
		scorefile.writeline ChimesOn
		scorefile.writeline ReplayLevel
		scorefile.writeline GameOption
		scorefile.writeline CircusSetting
		for xx=1 to 5
			scorefile.writeline HSScore(xx)
		next
		for xx=1 to 5
			scorefile.writeline HSName(xx)
		next
		ScoreFile.Close
	Set ScoreFile=Nothing
	Set FileObj=Nothing
end sub

sub SaveLMEMConfig
	Dim FileObj
	Dim LMConfig
	dim temp1
	dim tempb2s
	tempb2s=0
	if B2SOn=true then
		tempb2s=1
	else
		tempb2s=0
	end if
	Set FileObj=CreateObject("Scripting.FileSystemObject")
	If Not FileObj.FolderExists(UserDirectory) then 
		Exit Sub
	End if
	Set LMConfig=FileObj.CreateTextFile(UserDirectory & LMEMTableConfig,True)
	LMConfig.WriteLine tempb2s
	LMConfig.Close
	Set LMConfig=Nothing
	Set FileObj=Nothing

end Sub

sub LoadLMEMConfig
	Dim FileObj
	Dim LMConfig
	dim tempC
	dim tempb2s

    Set FileObj=CreateObject("Scripting.FileSystemObject")
	If Not FileObj.FolderExists(UserDirectory) then 
		Exit Sub
	End if
	If Not FileObj.FileExists(UserDirectory & LMEMTableConfig) then
		Exit Sub
	End if
	Set LMConfig=FileObj.GetFile(UserDirectory & LMEMTableConfig)
	Set TextStr2=LMConfig.OpenAsTextStream(1,0)
	If (TextStr2.AtEndOfStream=True) then
		Exit Sub
	End if
	tempC=TextStr2.ReadLine
	TextStr2.Close
	tempb2s=cdbl(tempC)
	if tempb2s=0 then
		B2SOn=false
	else
		B2SOn=true
	end if
	Set LMConfig=Nothing
	Set FileObj=Nothing
end sub

sub SaveLMEMConfig2
	If ShadowConfigFile=false then exit sub
	Dim FileObj
	Dim LMConfig2
	dim temp1
	dim temp2
	dim tempBS
	dim tempFS

	if EnableBallShadow=true then
		tempBS=1
	else
		tempBS=0
	end if
	if EnableFlipperShadow=true then
		tempFS=1
	else
		tempFS=0
	end if

	Set FileObj=CreateObject("Scripting.FileSystemObject")
	If Not FileObj.FolderExists(UserDirectory) then 
		Exit Sub
	End if
	Set LMConfig2=FileObj.CreateTextFile(UserDirectory & LMEMShadowConfig,True)
	LMConfig2.WriteLine tempBS
	LMConfig2.WriteLine tempFS
	LMConfig2.Close
	Set LMConfig2=Nothing
	Set FileObj=Nothing

end Sub

sub LoadLMEMConfig2
	If ShadowConfigFile=false then
		EnableBallShadow = ShadowBallOn
		BallShadowUpdate.enabled = ShadowBallOn
		EnableFlipperShadow = ShadowFlippersOn
		FlipperLSh.visible = ShadowFlippersOn
		FlipperRSh.visible = ShadowFlippersOn
		exit sub
	end if
	Dim FileObj
	Dim LMConfig2
	dim tempC
	dim tempD
	dim tempFS
	dim tempBS

    Set FileObj=CreateObject("Scripting.FileSystemObject")
	If Not FileObj.FolderExists(UserDirectory) then 
		Exit Sub
	End if
	If Not FileObj.FileExists(UserDirectory & LMEMShadowConfig) then
		Exit Sub
	End if
	Set LMConfig2=FileObj.GetFile(UserDirectory & LMEMShadowConfig)
	Set TextStr2=LMConfig2.OpenAsTextStream(1,0)
	If (TextStr2.AtEndOfStream=True) then
		Exit Sub
	End if
	tempC=TextStr2.ReadLine
	tempD=TextStr2.Readline
	TextStr2.Close
	tempBS=cdbl(tempC)
	tempFS=cdbl(tempD)
	if tempBS=0 then
		EnableBallShadow=false
		BallShadowUpdate.enabled=false
	else
		EnableBallShadow=true
	end if
	if tempFS=0 then
		EnableFlipperShadow=false
		FlipperLSh.visible=false
		FLipperRSh.visible=false
	else
		EnableFlipperShadow=true
	end if
	Set LMConfig2=Nothing
	Set FileObj=Nothing

end sub


sub loadhs
    ' Based on Black's Highscore routines
	Dim FileObj
	Dim ScoreFile
    dim temp1
    dim temp2
	dim temp3
	dim temp4
	dim temp5
	dim temp6
	dim temp7
	
	dim temp8
	dim temp9
	dim temp10
	dim temp11
	dim temp12
	dim temp13
	dim temp14
	dim temp15
	dim temp16
	dim temp17

    Set FileObj=CreateObject("Scripting.FileSystemObject")
	If Not FileObj.FolderExists(UserDirectory) then 
		Exit Sub
	End if
	If Not FileObj.FileExists(UserDirectory & HSFileName) then
		Exit Sub
	End if
	Set ScoreFile=FileObj.GetFile(UserDirectory & HSFileName)
	Set TextStr=ScoreFile.OpenAsTextStream(1,0)
		If (TextStr.AtEndOfStream=True) then
			Exit Sub
		End if
		temp1=TextStr.ReadLine
		temp2=textstr.readline
		temp3=textstr.readline
		temp4=textstr.readline
		temp5=textstr.readline
		temp6=textstr.readline
		temp7=textstr.readline

		HighScore=cdbl(temp1)
		if HighScore<1 then
			temp8=textstr.readline
			temp9=textstr.readline
			temp10=textstr.readline
			temp11=textstr.readline
			temp12=textstr.readline
			temp13=textstr.readline
			temp14=textstr.readline
			temp15=textstr.readline
			temp16=textstr.readline
			temp17=textstr.readline
		end if

		TextStr.Close
	    Credits=cdbl(temp2)
		BallsPerGame=cdbl(temp3)
		ChimesOn=cdbl(temp4)
		ReplayLevel=cdbl(temp5)
		GameOption=cdbl(temp6)
		CircusSetting=cdbl(temp7)
		if HighScore<1 then
			HSScore(1) = int(temp8)
			HSScore(2) = int(temp9)
			HSScore(3) = int(temp10)
			HSScore(4) = int(temp11)
			HSScore(5) = int(temp12)
			
			HSName(1) = temp13
			HSName(2) = temp14
			HSName(3) = temp15
			HSName(4) = temp16
			HSName(5) = temp17
		end if
		Set ScoreFile=Nothing
	    Set FileObj=Nothing
end sub

Sub DisplayHighScore
	

end sub

sub InitPauser5_timer
		
		DisplayHighScore
		CreditsReel.SetValue(Credits)
		InitPauser5.enabled=false
end sub



sub BumpersOff
	
	Bumper1Light.visible=0
	Bumper2Light.visible=0
	
	

end sub

sub BumpersOn


	Bumper1Light.visible=0

	If Bumper1Light.state=1 then
		
		Bumper1Light.visible=1
	end if
	
	Bumper2Light.visible=0

	If Bumper2Light.state=1 then
		
		Bumper2Light.visible=1
	end if


	
end sub

Sub PlasticsOn
	
	For each obj in Flashers
		obj.visible=1
	next


		
end sub

Sub PlasticsOff
	
	For each obj in Flashers
		obj.visible=0
	next



	StopSound "buzz"
	StopSound "buzzL"

end sub

Sub SetupReplayTables

	Replay1Table(1)=61000
	Replay1Table(2)=62000
	Replay1Table(3)=65000
	Replay1Table(4)=68000
	Replay1Table(5)=72000
	Replay1Table(6)=76000
	Replay1Table(7)=80000
	Replay1Table(8)=83000
	Replay1Table(9)=86000
	Replay1Table(10)=9990000
	Replay1Table(11)=9990000
	Replay1Table(12)=9990000
	Replay1Table(13)=9990000
	Replay1Table(14)=9990000
	Replay1Table(15)=9990000

	Replay2Table(1)=99000
	Replay2Table(2)=86000
	Replay2Table(3)=99000
	Replay2Table(4)=99000
	Replay2Table(5)=99000
	Replay2Table(6)=99000
	Replay2Table(7)=99000
	Replay2Table(8)=99000
	Replay2Table(9)=99000
	Replay2Table(10)=9990000
	Replay2Table(11)=9990000
	Replay2Table(12)=9990000
	Replay2Table(13)=9990000
	Replay2Table(14)=9990000
	Replay2Table(15)=9990000

	Replay3Table(1)=9990000
	Replay3Table(2)=99000
	Replay3Table(3)=9990000
	Replay3Table(4)=9990000
	Replay3Table(5)=9990000
	Replay3Table(6)=9990000
	Replay3Table(7)=9990000
	Replay3Table(8)=9990000
	Replay3Table(9)=9990000
	Replay3Table(10)=9990000
	Replay3Table(11)=9990000
	Replay3Table(12)=9990000
	Replay3Table(13)=9990000
	Replay3Table(14)=9990000
	Replay3Table(15)=9990000

	Replay4Table(1)=9990000
	Replay4Table(2)=9990000
	Replay4Table(3)=9990000
	Replay4Table(4)=9990000
	Replay4Table(5)=9990000
	Replay4Table(6)=9990000
	Replay4Table(7)=9990000
	Replay4Table(8)=9990000
	Replay4Table(9)=9990000
	Replay4Table(10)=9990000
	Replay4Table(11)=9990000
	Replay4Table(12)=9990000
	Replay4Table(13)=9990000
	Replay4Table(14)=9990000
	Replay4Table(15)=9990000

	ReplayTableMax=9


end sub

Sub RefreshReplayCard
	Dim tempst1
	Dim tempst2
	
	tempst1=FormatNumber(BallsPerGame,0)
	tempst2=FormatNumber(ReplayLevel,0)


	ReplayCard.image = "SC" + tempst2
	Replay1=Replay1Table(ReplayLevel)
	Replay2=Replay2Table(ReplayLevel)
	Replay3=Replay3Table(ReplayLevel)
	Replay4=Replay4Table(ReplayLevel)
'	CardLight2.State = ABS(CardLight2.State-1)
end sub


'****************************************
'  SCORE MOTOR
'****************************************

ScoreMotorTimer.Enabled = 1
ScoreMotorTimer.Interval = 135 '135
AddScoreTimer.Enabled = 1
AddScoreTimer.Interval = 135

Dim queuedscore
Dim MotorMode
Dim MotorPosition

Sub SetMotor(y)
	Select Case ScoreMotorAdjustment
		Case 0:
			queuedscore=queuedscore+y
		Case 1:
			If MotorRunning<>1 And InProgress=true then
				queuedscore=queuedscore+y
			end if
		end Select
end sub

Sub SetMotor2(x)
	If MotorRunning<>1 And InProgress=true then
		MotorRunning=1
		
		Select Case x
			Case 500:
				MotorMode=100
				MotorPosition=5
				
			Case 1000:
				MotorMode=1000
				MotorPosition=1
				
			Case 5000:
				MotorMode=1000
				MotorPosition=5
							
			Case 10000:
				MotorMode=10000
				MotorPosition=1
			Case 20000:
				MotorMode=10000
				MotorPosition=2
				
			Case 30000:
				MotorMode=10000
				MotorPosition=3
				
			Case 40000:
				MotorMode=10000
				MotorPosition=4
				
			Case 50000:
				MotorMode=10000
				MotorPosition=5
				
		End Select
	End If
End Sub

Sub AddScoreTimer_Timer
	Dim tempscore
	
	
	If MotorRunning<>1 And InProgress=true then
		if queuedscore>=50000 then
			tempscore=50000
			queuedscore=queuedscore-50000
			SetMotor2(50000)
			exit sub
		end if
		if queuedscore>=40000 then
			tempscore=4000
			queuedscore=queuedscore-40000
			SetMotor2(40000)
			exit sub
		end if
				
		if queuedscore>=30000 then
			tempscore=30000
			queuedscore=queuedscore-30000
			SetMotor2(30000)
			exit sub
		end if
			
		if queuedscore>=20000 then
			tempscore=20000
			queuedscore=queuedscore-20000
			SetMotor2(20000)
			exit sub
		end if
			
		if queuedscore>=10000 then
			tempscore=10000
			queuedscore=queuedscore-10000
			SetMotor2(10000)
			exit sub
		end if

		if queuedscore>=5000 then
			tempscore=5000
			queuedscore=queuedscore-5000
			SetMotor2(5000)
			exit sub
		end if

		if queuedscore>=1000 then
			tempscore=1000
			queuedscore=queuedscore-1000
			SetMotor2(1000)
			exit sub
		end if

		if queuedscore>=500 then
			tempscore=500
			queuedscore=queuedscore-500
			SetMotor2(500)
			exit sub
		end if	

	End If


end Sub

Sub ScoreMotorTimer_Timer
	If MotorPosition > 0 Then
		Select Case MotorPosition
			Case 5,4,3,2:
				If MotorMode=10000 Then 
					AddScore(10000)
				end if
				if MotorMode=1000 then
					AddScore(1000)
				End If
				if MotorMode=100 then
					AddScore(100)
				End if
				MotorPosition=MotorPosition-1
			Case 1:
				If MotorMode=10000 Then 
					AddScore(10000)
				end if
				If MotorMode=1000 then 
					AddScore(1000)
				End If
				if MotorMode=100 then
					AddScore(100)
				End if
				MotorPosition=0:MotorRunning=0
		End Select
	End If
End Sub


Sub AddScore(x)
	If TableTilted=true then exit sub
	Select Case ScoreAdditionAdjustment
		Case 0:
			AddScore1(x)
		Case 1:
			AddScore2(x)
	end Select

end sub


Sub AddScore1(x)
'	debugtext.text=score
	Select Case x
		Case 10:
			PlayChime(10)
			Score(Player)=Score(Player)+10
'			If DoubleBonus.state=1 then
'				Score(Player)=Score(Player)+10
'			end if
			ToggleAlternatingRelay
		Case 100:
			PlayChime(100)
			Score(Player)=Score(Player)+100
'			If DoubleBonus.state=1 then
'				Score(Player)=Score(Player)+100
'			end if

'			debugscore=debugscore+10
			
		Case 1000:
			PlayChime(1000)
			Score(Player)=Score(Player)+1000
'			If DoubleBonus.state=1 then
'				Score(Player)=Score(Player)+1000
'			end if

'			debugscore=debugscore+100
			
			
		Case 10000:
			PlayChime(10000)
			Score(Player)=Score(Player)+10000
'			If DoubleBonus.state=1 then
'				Score(Player)=Score(Player)+10000
'			end if

'			debugscore=debugscore+1000

		Case 100000:
			PlayChime(10000)
			Score(Player)=Score(Player)+100000
'			If DoubleBonus.state=1 then
'				Score(Player)=Score(Player)+10000
'			end if

'			debugscore=debugscore+1000
			

	End Select
	PlayerScores(Player-1).AddValue(x)
	PlayerScoresOn(Player-1).AddValue(x)
'	if DoubleBonus.state=1 then
'		PlayerScores(Player-1).AddValue(x)
'	end if
	If ScoreDisplay(Player)<100000 then
		ScoreDisplay(Player)=Score(Player)
	Else
		Score100K(Player)=Int(Score(Player)/100000)
		ScoreDisplay(Player)=Score(Player)-100000
	End If
	if Score(Player)=>100000 then
		
		If B2SOn Then
			If Player=1 Then
				Controller.B2SSetScoreRolloverPlayer1 Score100K(Player)
			End If
			If Player=2 Then
				Controller.B2SSetScoreRolloverPlayer2 Score100K(Player)
			End If

			If Player=3 Then
				Controller.B2SSetScoreRolloverPlayer3 Score100K(Player)
			End If

			If Player=4 Then
				Controller.B2SSetScoreRolloverPlayer4 Score100K(Player)
			End If
		End If
	End If
	If B2SOn Then
		Controller.B2SSetScorePlayer Player, ScoreDisplay(Player)
	End If
	If Score(Player)>Replay1 and Replay1Paid(Player)=false then
		Replay1Paid(Player)=True
		AddSpecial
	End If
	If Score(Player)>Replay2 and Replay2Paid(Player)=false then
		Replay2Paid(Player)=True
		AddSpecial
	End If
	If Score(Player)>Replay3 and Replay3Paid(Player)=false then
		Replay3Paid(Player)=True
		AddSpecial
	End If
	If Score(Player)>Replay4 and Replay4Paid(Player)=false then
		Replay4Paid(Player)=True
		AddSpecial
	End If
'	ScoreText.text=debugscore
End Sub

Sub AddScore2(x)
	Dim OldScore, NewScore, OldTestScore, NewTestScore
    OldScore = Score(Player)

	Select Case x
        Case 10:
            Score(Player)=Score(Player)+10
'			If DoubleBonus.state=1 then
'				Score(Player)=Score(Player)+10
'			end if

		Case 100:
			Score(Player)=Score(Player)+100
'			If DoubleBonus.state=1 then
'				Score(Player)=Score(Player)+100
'			end if

		Case 1000:
			Score(Player)=Score(Player)+1000
'			If DoubleBonus.state=1 then
'				Score(Player)=Score(Player)+1000
'			end if

		Case 10000:
			Score(Player)=Score(Player)+10000
'			If DoubleBonus.state=1 then
'				Score(Player)=Score(Player)+10000
'			end if

		Case 100000:
			Score(Player)=Score(Player)+100000
'			If DoubleBonus.state=1 then
'				Score(Player)=Score(Player)+100000
'			end if

	End Select
	if Score(Player)=>100000 then
		If Score100K(Player)<1 then
			PlaySound"BallyBuzzer"
			RolloverReel.Setvalue(1)
			If B2SOn Then
				If Player=1 Then
					Controller.B2SSetScoreRolloverPlayer1 1
				End If
				If Player=2 Then
					Controller.B2SSetScoreRolloverPlayer2 1
				End If
	
				If Player=3 Then
					Controller.B2SSetScoreRolloverPlayer3 Score100K(Player)
				End If
	
				If Player=4 Then
					Controller.B2SSetScoreRolloverPlayer4 Score100K(Player)
				End If
			End If
		end if
		Score100K(Player)=1
	End If
	NewScore = Score(Player)

	OldTestScore = OldScore
	NewTestScore = NewScore
	Do
		if OldTestScore < Replay1 and NewTestScore >= Replay1 then
			AddSpecial()
			NewTestScore = 0
		Elseif OldTestScore < Replay2 and NewTestScore >= Replay2 then
			AddSpecial()
			NewTestScore = 0
		Elseif OldTestScore < Replay3 and NewTestScore >= Replay3 then
			AddSpecial()
			NewTestScore = 0
		Elseif OldTestScore < Replay4 and NewTestScore >= Replay4 then
			AddSpecial()
			NewTestScore = 0
		End if
		NewTestScore = NewTestScore - 1000000
		OldTestScore = OldTestScore - 1000000
	Loop While NewTestScore > 0

    OldScore = int(OldScore / 10)	' divide by 10 for games with fixed 0 in 1s position, by 1 for games with real 1s digits
    NewScore = int(NewScore / 10)	' divide by 10 for games with fixed 0 in 1s position, by 1 for games with real 1s digits
	' MsgBox("OldScore="&OldScore&", NewScore="&NewScore&", OldScore Mod 10="&OldScore Mod 10 & ", NewScore % 10="&NewScore Mod 10)

    if (OldScore Mod 10 <> NewScore Mod 10) then
		PlayChime(10)
		ToggleAlternatingRelay
    end if

    OldScore = int(OldScore / 10)
    NewScore = int(NewScore / 10)
	' MsgBox("OldScore="&OldScore&", NewScore="&NewScore)
    if (OldScore Mod 10 <> NewScore Mod 10) then
		PlayChime(100)
		
		
    end if

    OldScore = int(OldScore / 10)
    NewScore = int(NewScore / 10)
	' MsgBox("OldScore="&OldScore&", NewScore="&NewScore)
    if (OldScore Mod 10 <> NewScore Mod 10) then
		PlayChime(1000)
		
		

    end if

    OldScore = int(OldScore / 10)
    NewScore = int(NewScore / 10)
	' MsgBox("OldScore="&OldScore&", NewScore="&NewScore)
    if (OldScore Mod 10 <> NewScore Mod 10) then
		PlayChime(10000)
    end if

	If B2SOn Then
		Controller.B2SSetScorePlayer Player, Score(Player)
	End If

	OldScore = int(OldScore / 10)
    NewScore = int(NewScore / 10)
	' MsgBox("OldScore="&OldScore&", NewScore="&NewScore)
    if (OldScore Mod 10 <> NewScore Mod 10) then
		PlayChime(10000)
    end if

	If B2SOn Then
		Controller.B2SSetScorePlayer Player, Score(Player)
	End If
'	EMReel1.SetValue Score(Player)
	PlayerScores(Player-1).AddValue(x)
	PlayerScoresOn(Player-1).AddValue(x)
'	If DoubleBonus.state=1 then
'		PlayerScores(Player-1).AddValue(x)
'	end if
End Sub



Sub PlayChime(x)
	if ChimesOn=0 then
		Select Case x
			Case 10
				
				PlaySound SoundFXDOF("10a",141,DOFPulse,DOFChimes)
				
			Case 100
				
				PlaySound SoundFXDOF("100a",142,DOFPulse,DOFChimes)
			Case 1000
				
				PlaySound SoundFXDOF("1000a",143,DOFPulse,DOFChimes)
		End Select
	else
		Select Case x
			Case 10
				If LastChime10=1 Then
					PlaySound SoundFXDOF("SJ_Chime_10a",141,DOFPulse,DOFChimes)
					LastChime10=0
				Else
					PlaySound SoundFXDOF("SJ_Chime_10b",141,DOFPulse,DOFChimes)
					LastChime10=1
				End If
			Case 100
				If LastChime100=1 Then
					PlaySound SoundFXDOF("SJ_Chime_100a",142,DOFPulse,DOFChimes)
					LastChime100=0
				Else
					PlaySound SoundFXDOF("SJ_Chime_100b",142,DOFPulse,DOFChimes)
					LastChime100=1
				End If
			Case 1000
				If LastChime1000=1 Then
					PlaySound SoundFXDOF("SJ_Chime_1000a",143,DOFPulse,DOFChimes)
					LastChime1000=0
				Else
					PlaySound SoundFXDOF("SJ_Chime_1000b",143,DOFPulse,DOFChimes)
					LastChime1000=1
				End If
		End Select
	end if
End Sub


Sub HideOptions()

end sub

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

Sub RollingSoundTimer_Timer()
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
            PlaySound("fx_ballrolling" & b), -1, Vol(BOT(b)), AudioPan(BOT(b)), 0, Pitch(BOT(b)), 1, 0, AudioFade(BOT(b))
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
	PlaySound("fx_collide"), 0, Csng(velocity) ^2 / 2000, AudioPan(ball1), 0, Pitch(ball1), 0, 0, AudioFade(ball1)
End Sub



'*****************************************
'	ninuzzu's	BALL SHADOW
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
        If BOT(b).X < Table1.Width/2 Then
            BallShadow(b).X = ((BOT(b).X) - (Ballsize/6) + ((BOT(b).X - (Table1.Width/2))/7)) + 6
        Else
            BallShadow(b).X = ((BOT(b).X) + (Ballsize/6) + ((BOT(b).X - (Table1.Width/2))/7)) - 6
        End If
        ballShadow(b).Y = BOT(b).Y + 12
        If BOT(b).Z > 20 Then
            BallShadow(b).visible = 1
        Else
            BallShadow(b).visible = 0
        End If
    Next
End Sub

'*****************************************
'	Object sounds
'*****************************************

Sub Plastics_Hit (idx)
	PlaySound "woodhit_low", 0, Vol(ActiveBall), AudioPan(ActiveBall), 0, Pitch(ActiveBall), 0, 0, AudioFade(ActiveBall)
End Sub

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

' ============================================================================================
' GNMOD - Multiple High Score Display and Collection
' ============================================================================================
Dim EnteringInitials		' Normally zero, set to non-zero to enter initials
EnteringInitials = 0

Dim PlungerPulled
PlungerPulled = 0

Dim SelectedChar			' character under the "cursor" when entering initials

Dim HSTimerCount			' Pass counter for HS timer, scores are cycled by the timer
HSTimerCount = 5			' Timer is initially enabled, it'll wrap from 5 to 1 when it's displayed

Dim InitialString			' the string holding the player's initials as they're entered

Dim AlphaString				' A-Z, 0-9, space (_) and backspace (<)
Dim AlphaStringPos			' pointer to AlphaString, move forward and backward with flipper keys
AlphaString = "ABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789_<"

Dim HSNewHigh				' The new score to be recorded

Dim HSScore(5)				' High Scores read in from config file
Dim HSName(5)				' High Score Initials read in from config file

' default high scores, remove this when the scores are available from the config file
HSScore(1) = 75000
HSScore(2) = 70000
HSScore(3) = 60000
HSScore(4) = 55000
HSScore(5) = 50000

HSName(1) = "AAA"
HSName(2) = "ZZZ"
HSName(3) = "XXX"
HSName(4) = "ABC"
HSName(5) = "BBB"

Sub HighScoreTimer_Timer
	
	if EnteringInitials then
		if HSTimerCount = 1 then
			SetHSLine 3, InitialString & MID(AlphaString, AlphaStringPos, 1)
			HSTimerCount = 2
		else
			SetHSLine 3, InitialString
			HSTimerCount = 1
		end if
	elseif InProgress then
		SetHSLine 1, "HIGH SCORE1"
		SetHSLine 2, HSScore(1)
		SetHSLine 3, HSName(1)
		HSTimerCount = 5	' set so the highest score will show after the game is over
		HighScoreTimer.enabled=false
	elseif CheckAllScores then
		NewHighScore sortscores(ScoreChecker),sortplayers(ScoreChecker)

	else
		' cycle through high scores
		HighScoreTimer.interval=2000
		HSTimerCount = HSTimerCount + 1
		if HsTimerCount > 5 then
			HSTimerCount = 1
		End If
		SetHSLine 1, "HIGH SCORE"+FormatNumber(HSTimerCount,0)
		SetHSLine 2, HSScore(HSTimerCount)
		SetHSLine 3, HSName(HSTimerCount)
	end if
End Sub

Function GetHSChar(String, Index)
	dim ThisChar
	dim FileName
	ThisChar = Mid(String, Index, 1)
	FileName = "PostIt"
	if ThisChar = " " or ThisChar = "" then
		FileName = FileName & "BL"
	elseif ThisChar = "<" then
		FileName = FileName & "LT"
	elseif ThisChar = "_" then
		FileName = FileName & "SP"
	else
		FileName = FileName & ThisChar
	End If
	GetHSChar = FileName
End Function

Sub SetHsLine(LineNo, String)
	dim Letter
	dim ThisDigit
	dim ThisChar
	dim StrLen
	dim LetterLine
	dim Index
	dim StartHSArray
	dim EndHSArray
	dim LetterName
	dim xfor
	StartHSArray=array(0,1,12,22)
	EndHSArray=array(0,11,21,31)
	StrLen = len(string)
	Index = 1

	for xfor = StartHSArray(LineNo) to EndHSArray(LineNo)
		Eval("HS"&xfor).image = GetHSChar(String, Index)
		Index = Index + 1
	next

End Sub

Sub NewHighScore(NewScore, PlayNum)
	if NewScore > HSScore(5) then
		HighScoreTimer.interval = 500
		HSTimerCount = 1
		AlphaStringPos = 1		' start with first character "A"
		EnteringInitials = 1	' intercept the control keys while entering initials
		InitialString = ""		' initials entered so far, initialize to empty
		SetHSLine 1, "PLAYER "+FormatNumber(PlayNum,0)
		SetHSLine 2, "ENTER NAME"
		SetHSLine 3, MID(AlphaString, AlphaStringPos, 1)
		HSNewHigh = NewScore
		For xx=1 to HighScoreReward
			AddSpecial
		next
	End if
	ScoreChecker=ScoreChecker-1
	if ScoreChecker=0 then
		CheckAllScores=0
	end if
End Sub

Sub CollectInitials(keycode)
	If keycode = LeftFlipperKey Then
		' back up to previous character
		AlphaStringPos = AlphaStringPos - 1
		if AlphaStringPos < 1 then
			AlphaStringPos = len(AlphaString)		' handle wrap from beginning to end
			if InitialString = "" then
				' Skip the backspace if there are no characters to backspace over
				AlphaStringPos = AlphaStringPos - 1
			End if
		end if
		SetHSLine 3, InitialString & MID(AlphaString, AlphaStringPos, 1)
		PlaySound "DropTargetDropped"
	elseif keycode = RightFlipperKey Then
		' advance to next character
		AlphaStringPos = AlphaStringPos + 1
		if AlphaStringPos > len(AlphaString) or (AlphaStringPos = len(AlphaString) and InitialString = "") then
			' Skip the backspace if there are no characters to backspace over
			AlphaStringPos = 1
		end if
		SetHSLine 3, InitialString & MID(AlphaString, AlphaStringPos, 1)
		PlaySound "DropTargetDropped"
	elseif keycode = StartGameKey or keycode = PlungerKey Then
		SelectedChar = MID(AlphaString, AlphaStringPos, 1)
		if SelectedChar = "_" then
			InitialString = InitialString & " "
			PlaySound("Ding10")
		elseif SelectedChar = "<" then
			InitialString = MID(InitialString, 1, len(InitialString) - 1)
			if len(InitialString) = 0 then
				' If there are no more characters to back over, don't leave the < displayed
				AlphaStringPos = 1
			end if
			PlaySound("Ding100")
		else
			InitialString = InitialString & SelectedChar
			PlaySound("Ding10")
		end if
		if len(InitialString) < 3 then
			SetHSLine 3, InitialString & SelectedChar
		End If
	End If
	if len(InitialString) = 3 then
		SetHSLine 3, InitialString
		' save the score
		for i = 5 to 1 step -1
			if i = 1 or (HSNewHigh > HSScore(i) and HSNewHigh <= HSScore(i - 1)) then
				' Replace the score at this location
				if i < 5 then
' MsgBox("Moving " & i & " to " & (i + 1))
					HSScore(i + 1) = HSScore(i)
					HSName(i + 1) = HSName(i)
				end if
' MsgBox("Saving initials " & InitialString & " to position " & i)
				EnteringInitials = 0
				HSScore(i) = HSNewHigh
				HSName(i) = InitialString
				HSTimerCount = 5
				HighScoreTimer_Timer
				HighScoreTimer.interval = 2000
				PlaySound("Ding1000")
				exit sub
			elseif i < 5 then
				' move the score in this slot down by 1, it's been exceeded by the new score
' MsgBox("Moving " & i & " to " & (i + 1))
				HSScore(i + 1) = HSScore(i)
				HSName(i + 1) = HSName(i)
			end if
		next
	End If

End Sub
' END GNMOD
' ============================================================================================
' GNMOD - New Options menu
' ============================================================================================
Dim EnteringOptions
Dim CurrentOption
Dim OptionCHS
Dim MaxOption
Dim OptionHighScorePosition
Dim XOpt
Dim StartingArray
Dim EndingArray

StartingArray=Array(0,1,2,30,33,61,89,117,145,173,201,229)
EndingArray=Array(0,1,29,32,60,88,116,144,172,200,228,256)
EnteringOptions = 0
MaxOption = 9
OptionCHS = 0
OptionHighScorePosition = 0
Const OptionLinesToMark="111010011"
Const OptionLine1="" 'do not use this line
Const OptionLine2="" 'do not use this line
Const OptionLine3="" 'do not use this line
Const OptionLine4=""
Const OptionLine5="Game Option"
Const OptionLine6=""
Const OptionLine7=""
Const OptionLine8="" 'do not use this line
Const OptionLine9="" 'do not use this line

Sub OperatorMenuTimer_Timer
	EnteringOptions = 1
	OperatorMenuTimer.enabled=false
	ShowOperatorMenu
end sub

sub ShowOperatorMenu
	OperatorMenuBackdrop.image = "OperatorMenu"

	OptionCHS = 0
	CurrentOption = 1
	DisplayAllOptions
	OperatorOption1.image = "BluePlus"
	SetHighScoreOption

End Sub

Sub DisplayAllOptions
	dim linecounter
	dim tempstring
	For linecounter = 1 to MaxOption
		tempstring=Eval("OptionLine"&linecounter)
		Select Case linecounter
			Case 1:
				tempstring=tempstring + FormatNumber(BallsPerGame,0)
				SetOptLine 1,tempstring
			Case 2:
				if Replay3Table(ReplayLevel)=9990000 then
					tempstring = FormatNumber(Replay1Table(ReplayLevel),0) + "/" + FormatNumber(Replay2Table(ReplayLevel),0)
				elseif Replay4Table(ReplayLevel)=9990000 then
					tempstring = FormatNumber(Replay1Table(ReplayLevel),0) + "/" + FormatNumber(Replay2Table(ReplayLevel),0) + "/" + FormatNumber(Replay3Table(ReplayLevel),0)
				else
					tempstring = FormatNumber(Replay1Table(ReplayLevel),0,0,0,0) + "/" + FormatNumber(Replay2Table(ReplayLevel),0,0,0,0) + "/" + FormatNumber(Replay3Table(ReplayLevel),0,0,0,0) + "/" + FormatNumber(Replay4Table(ReplayLevel),0,0,0,0)
				end if
			SetOptLine 2,tempstring
			Case 3:
				If OptionCHS=0 then
					tempstring = "NO"
				else
					tempstring = "YES"
				end if
				SetOptLine 3,tempstring
			Case 4:
				SetOptLine 4, tempstring


				SetOptLine 5, tempstring
			Case 5:
				SetOptLine 6, tempstring
				if GameOption=1 then
					tempstring = "Conservative"
				else
					tempstring = "Liberal"
				end if
				SetOptLine 7, tempstring

			Case 6:
				SetOptLine 8, tempstring


				SetOptLine 9, tempstring
				
			Case 7:
				SetOptLine 10, tempstring
				SetOptLine 11, tempstring
			
			Case 8:
		
			Case 9:
			
	
		End Select
		
	next
end sub

sub MoveArrow
	do 
		CurrentOption = CurrentOption + 1
		If CurrentOption>Len(OptionLinesToMark) then
			CurrentOption=1
		end if
	loop until Mid(OptionLinesToMark,CurrentOption,1)="1"
end sub

sub CollectOptions(ByVal keycode)
	if Keycode = LeftFlipperKey then
		PlaySound "DropTargetDropped"
		For XOpt = 1 to MaxOption
			Eval("OperatorOption"&XOpt).image = "PostitBL"
		next
		MoveArrow
		if CurrentOption<8 then
			Eval("OperatorOption"&CurrentOption).image = "BluePlus"
		elseif CurrentOption=8 then
			Eval("OperatorOption"&CurrentOption).image = "GreenCheck"
		else
			Eval("OperatorOption"&CurrentOption).image = "RedX"
		end if
			
	elseif Keycode = RightFlipperKey then
		PlaySound "DropTargetDropped"
		if CurrentOption = 1 then
			If BallsPerGame = 3 then
				BallsPerGame = 5
			else
				BallsPerGame = 3
			end if
			DisplayAllOptions
		elseif CurrentOption = 2 then
			ReplayLevel=ReplayLevel+1
			If ReplayLevel>ReplayTableMax then
				ReplayLevel=1
			end if
			DisplayAllOptions
		elseif CurrentOption = 3 then
			if OptionCHS = 0 then
				OptionCHS = 1
				
			else
				OptionCHS = 0
				
			end if
			DisplayAllOptions


		elseif CurrentOption = 5 then
			if GameOption=1 then
				GameOption=2
			
			else
				GameOption=1
			
			end if
			DisplayAllOptions

		elseif CurrentOption = 8 or CurrentOption = 9 then
				if OptionCHS=1 then
					HSScore(1) = 75000	
					HSScore(2) = 70000
					HSScore(3) = 60000
					HSScore(4) = 55000
					HSScore(5) = 50000

					HSName(1) = "AAA"
					HSName(2) = "ZZZ"
					HSName(3) = "XXX"
					HSName(4) = "ABC"
					HSName(5) = "BBB"
				end if
	
				if CurrentOption = 8 then
					savehs
				else
					loadhs
				end if
				OperatorMenuBackdrop.image = "PostitBL"
				For XOpt = 1 to MaxOption
					Eval("OperatorOption"&XOpt).image = "PostitBL"
				next
			
				For XOpt = 1 to 256
					Eval("Option"&XOpt).image = "PostItBL"
				next
				RefreshReplayCard
				InstructCard.image="IC"+FormatNumber(BallsPerGame,0)
				EnteringOptions = 0

		end if
	end if
End Sub

Sub SetHighScoreOption
	
End Sub

Function GetOptChar(String, Index)
	dim ThisChar
	dim FileName
	ThisChar = Mid(String, Index, 1)
	FileName = "PostIt"
	if ThisChar = " " or ThisChar = "" then
		FileName = FileName & "BL"
	elseif ThisChar = "<" then
		FileName = FileName & "LT"
	elseif ThisChar = "_" then
		FileName = FileName & "SP"
	elseif ThisChar = "/" then
		FileName = FileName & "SL"
	elseif ThisChar = "," then
		FileName = FileName & "CM"
	else
		FileName = FileName & ThisChar
	End If
	GetOptChar = FileName
End Function

dim LineLengths(22)	' maximum number of lines
Sub SetOptLine(LineNo, String)
	Dim DispLen
    Dim StrLen
	dim xfor
	dim Letter
	dim ThisDigit
	dim ThisChar
	dim LetterLine
	dim Index
	dim LetterName
	StrLen = len(string)
	Index = 1

	StrLen = len(String)
    DispLen = StrLen
    if (DispLen < LineLengths(LineNo)) Then
        DispLen = LineLengths(LineNo)
    end If

	for xfor = StartingArray(LineNo) to StartingArray(LineNo) + DispLen
		Eval("Option"&xfor).image = GetOptChar(string, Index)
		Index = Index + 1
	next
	LineLengths(LineNo) = StrLen

End Sub
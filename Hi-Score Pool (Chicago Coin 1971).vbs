' ****************************************************************
' Hi-Score Pool from Chicago Coins, 1971
' STAT, stefanaustria and BorgDog
' 14.05.2016
' Version 1.1
'
' Thanks to:
' HauntFreaks graphic upgrades
' PinballShawn (much Info and Video)
' JP, kiwi: Informations
' DOF by Arngrim
'
' Option menu added to change game modes, hold left flipper before starting game
' ****************************************************************

' Thalamus 2018-07-23
' Added/Updated "Positional Sound Playback Functions" and "Supporting Ball & Sound Functions"
' No special SSF tweaks yet.

Option Explicit
Randomize

On Error Resume Next
ExecuteGlobal GetTextFile("controller.vbs")
If Err Then MsgBox "Can't open controller.vbs"
On Error Goto 0

Const cGameName = "hiscorepool"

Dim BIP, Balls, gLight
Dim PlayerScores(2), SubScores(3), ScoresRoll, SubPlayer
Dim Player, Players, Credits
Dim j, GameMode, Objekt, wdown
Dim BallsLit, PlayBall, D, R
Dim operatormenu, options, hiscore
Dim Rack, Racks, Tilt, tiltsens
Dim StarLit
Dim BonusScore
Dim ExtraBalls
Dim motorsound, moveballs, KnockerOn
Dim ballangle(15), direction(15)

Const BallSize = 22

Sub Table1_Init()
    LoadEM
	Players = 0
	D = 0 'for Disc Degree
	R = -1
	hideoptions
	loadhs
	EndScoresTimer.Enabled = false
	If Credits="" or Credits<0 or Credits>15 then Credits=0
	If GameMode="" or GameMode <1 or GameMode >2 then GameMode=2
	If Hiscore="" then hiscore=10000
	If motorsound="" or motorsound<1 or motorsound>2 then motorsound=1
	If moveballs="" or moveballs<1 or moveballs>2 then moveballs=2
	If knockeron="" then knockeron=True
	If knockeron=0 then
		knockeron=False
	  Else
		knockeron=True
	end if
	TextHiScore.text = hiscore
	OptionMode.image = "OptionsMode"&GameMode
	OptionMotor.image = "OptionsMode"&motorsound
	OptionMove.image = "OptionsMode"&moveballs
	If B2SOn Then
		Backglasstimer.enabled=1
		for each objekt in BackDropStuff: Objekt.visible=0:Next
	end if
	EMReel3.SetValue Credits
	GOReel.setvalue 1
	EMReel1.Image = "NumbersDark"
	EMReel2.Image = "NumbersDark"
	Tilt=False

	For each objekt in RotWalls: objekt.isdropped=True: Next

	rotWall0.IsDropped = false

	If Credits > 0 Then DOF 103, DOFOn

End Sub

sub backglasstimer_timer
	Controller.B2SSetCredits Credits
	Controller.B2SSetGameover 1
	Controller.B2SSetScorePlayer 5, hiscore
	me.enabled=0
end sub


Sub Table1_KeyDown(ByVal keycode)

	if keycode=AddCreditKey then
		playsound "coinin"
		addcredit
    end if

	If keycode = LeftFlipperKey and Rack = 0 and BIP = 1 and tilt=False Then
		LeftFlipper.RotateToEnd
		PlaySound SoundFXDOF("fx_flipperup",101,DOFOn,DOFContactors), 0, .67, -0.05, 0.05
	End If

	If keycode = RightFlipperKey and Rack = 0 and BIP = 1 and tilt=False Then
		RightFlipper.RotateToEnd
		PlaySound SoundFXDOF("fx_flipperup",102,DOFOn,DOFContactors), 0, .67, 0.05, 0.05
	End If


	If keycode=LeftFlipperKey and disctimer.enabled = false and OperatorMenu=0 then
		OperatorMenuTimer.Enabled = true
	end if

	If keycode=LeftFlipperKey and disctimer.enabled = false and OperatorMenu=1 then
		Options=Options+1
		If Options=7 then Options=1
		playsound "target"
		Select Case (Options)
			Case 1:
				Option1.visible=true
				Option6.visible=False
			Case 2:
				Option2.visible=true
				Option1.visible=False
			Case 3:
				Option3.visible=true
				Option2.visible=False
			Case 4:
				Option4.visible=true
				Option3.visible=False
			Case 5:
				Option5.visible=true
				Option4.visible=False
			Case 6:
				Option6.visible=true
				Option5.visible=False
		End Select

	end if

	If keycode=RightFlipperKey and disctimer.enabled = false and OperatorMenu=1 then
	  PlaySound "metalhit2"
	  Select Case (Options)
		Case 1:
			if GameMode=1 then
				GameMode=2
			  else
				GameMode=1
			end if
			OptionMode.image = "OptionsMode"&GameMode
		Case 2:
			if motorsound=1 Then
				motorsound=2
			  Else
				motorsound=1
			end If
			OptionMotor.image = "OptionsMode"&motorsound
		Case 3:
			if knockeron=true Then
				KnockerOn=False
				OptionKnocker.image = "OptionsMode2"
			  Else
				KnockerOn=True
				OptionKnocker.image = "OptionsMode1"
			end if
		Case 4:
			if moveballs=1 Then
				moveballs=2
			  Else
				moveballs=1
			end If
			OptionMove.image = "OptionsMode"&moveballs
		Case 5:
			Hiscore=10000
			TextHiScore.text = hiscore
		Case 6:
			OperatorMenu=0
			savehs
			HideOptions
	  End Select
	End If

	If (keycode = PlungerKey or keycode = RightFlipperKey) and ScoresRoll = 0 and (( D > -1 and D < 35) or (D > 325 and D <360)) Then
		'Plunger.Fire
		If Balls > 0 and Player > 0 and BIP = 0 Then
			'PlaySound "plunger",0,1,0.25,0.25
			PlaySound SoundFXDOF("ballrelease",106,DOFPulse,DOFContactors),0,1,0,0.25
			DOF 107, DOFPulse
			BallRelease.CreateSizedBallWithMass BallSize, BallSize^3/15625
			BallRelease.Kick D, 40

			BIP = 1
			Disc.Image = "rotDisc"
			wallsup.enabled=1
			LightTriangle.State = 0
			LightLeft.State = 0
			LightRight.State = 0
			LightLeft1.State = 0
			LightRight1.State = 0
			ShowExtraBalls(ExtraBalls)
		End If
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
		mechchecktilt
	End If

End Sub

Sub CheckTilt
	If Tilttimer.Enabled = True Then
	 TiltSens = TiltSens + 1
	 if TiltSens = 3 Then
	   Tilt = True
		TILTreel.setvalue 1
       	If B2SOn Then Controller.B2SSetTilt 33,1
	    playsound "tilt"
		turnoff
	 End If
	Else
	 TiltSens = 0
	 Tilttimer.Enabled = True
	End If
End Sub

Sub CheckTilt
	   Tilt = True
		TILTreel.setvalue 1
       	If B2SOn Then Controller.B2SSetTilt 33,1
	    playsound "tilt"
		turnoff
End Sub

Sub Tilttimer_Timer()
	Tilttimer.Enabled = False
End Sub

Sub turnoff
	LeftFlipper.RotateToStart:DOF 101, DOFOff
	RightFlipper.RotateToStart:DOF 102, DOFOff
end sub

Sub OperatorMenuTimer_Timer
	OperatorMenu=1
	Displayoptions
	Options=1
End Sub

Sub DisplayOptions
	OptionsBack.visible = true
	OptionMode.visible = true
	OptionMotor.visible = true
	OptionMove.visible = true
	OptionKnocker.visible = True
	Option1.visible = True
End Sub

Sub HideOptions
	for each objekt In OptionMenu
		objekt.visible = false
	next
End Sub

Sub Table1_KeyUp(ByVal keycode)

	If Keycode = AddCreditKey Then
        Credits = Credits + 1: If Credits > 9 Then Credits = 9
		DOF 103, DOFOn
		EMReel3.SetValue Credits
		PlaySound "coinin",0,1,0.25,0.25
		If B2SOn Then Controller.B2SSetCredits Credits
	End If

	If keycode = StartGameKey and Credits > 0 and Players < 2 Then
		PlaySound "BallLoad",0,1,0.25,0.25
'		LightOn.State = 1
		LightOn2.State = 1
		LightOn3.State = 1
		' all BallLights on
		For each gLight in BallLights
			gLight.state = 1
		Next
		BallsLit = 0
		tilt=False
		wdown=1
		WallsDown
		DiscTimer.Enabled = true
		if motorsound=1 then playsound SoundFXDOF("motor",105,DOFOn,DOFGear), 1,.1
		GOReel.setvalue 0
		Players = Players + 1
		Credits = Credits - 1
		If Credits < 1 Then DOF 103, DOFOff
		EMReel3.SetValue Credits
		EMReel1.Image = "Numbers"
		EVAL("Lcanplay"&Players).state=1
		Player = 1
		Balls = 5
		ExtraBalls = 0
		BIP = 0
		Rack = 0
		Racks = 0
		Bonus			'Show Bonus Score
		TextBalls.Text = "1"
		PlayerScores(1) = 0: Scores 1,0
		PlayerScores(2) = 0: Scores 2,0
		SubScores(1) = 0 ' Scores for Balls after Drain
		SubScores(2) = 0 ' Scores for Bonus
		SubScores(3) = 0 ' Scores for Star Lit: 10.000
		If B2SOn Then
			B2SBIP(1)
			Controller.B2SSetCredits Credits
			Controller.B2SSetPlayerUp 1
			Controller.B2SSetData 39+Players, 1
			Controller.B2SSetGameover 0
			Controller.B2SSetCanPlay Players
		End if
	End If

	if keycode = LeftFlipperKey then
		OperatorMenuTimer.Enabled = false
	end if

  if BIP=1 Then
	If keycode = LeftFlipperKey and tilt=False Then
		LeftFlipper.RotateToStart
		PlaySound SoundFXDOF("fx_flipperdown",101,DOFOff,DOFContactors), 0, 1, -0.05, 0.05
	End If

	If keycode = RightFlipperKey and tilt=False Then
		RightFlipper.RotateToStart
		PlaySound SoundFXDOF("fx_flipperdown",102,DOFOff,DOFContactors), 0, 1, 0.05, 0.05
	End If
  end if

End Sub


Sub addcredit
      credits=Credits+1
	  DOF 103, DOFOn
      if credits>15 then credits=15
	  EMReel3.setvalue(credits)
	  If B2SOn Then Controller.B2ssetCredits Credits
End sub

Sub WallsDown
	wdown=1
	For each objekt in RotWalls: objekt.isdropped=True: Next
End Sub

Sub WallsUp_Timer
	wdown=0
	me.enabled=0
End Sub

Sub DiscTimer_Timer()
	Eval("rotWall"&D).IsDropped = true


	If R = -1  Then D = D - 2
	If D < 0   Then D = 358
	If D = 308 Then
		if motorsound=1 then stopsound "motor":DOF 105, DOFOff
		if motorsound=1 then playsound SoundFXDOF("motor",105,DOFOn,DOFGear),1,.1
		R = 1
	end if
	If R =  1  Then D = D + 2
	If D > 359 Then D = 0
	If D = 50  Then
		if motorsound=1 then stopsound "motor":DOF 105, DOFOff
		if motorsound=1 then playsound SoundFXDOF("motor",105,DOFOn,DOFGear),1,.1
		R = -1
	end if
	Disc.RotZ = D

	if wdown=0 then Eval("rotWall"&D).IsDropped = false

End Sub

Sub Bonus
	If B2SOn Then
		Controller.B2SSetData 13,0
		Controller.B2SSetData 14,0
		Controller.B2SSetData 15,0
	End If
	TextBonus.Text = ""
	BonusScore = 0
	' clear, then check:
	If Balls = 5 Then
		If B2SOn Then Controller.B2SSetData 13,1
		TextBonus.Text = "4000"
		BonusScore = 4000
	End if
	If Balls = 4 Then
		If B2SOn Then Controller.B2SSetData 14,1
		TextBonus.Text = "7000"
		BonusScore = 7000
	End if
	If Balls > 0 and Balls < 4 Then
		If B2SOn Then Controller.B2SSetData 15,1
		TextBonus.Text = "10000"
		BonusScore = 10000
	End if
End Sub

Sub B2SBIP(b)
	For j = 1 to 5
	Controller.B2SSetData j, 0
	Next
	If b > 0 Then Controller.B2SSetData b, 3
End Sub

Sub Scores(pl, points)
	PlayerScores(pl) = PlayerScores(pl) + points
	If B2SOn Then Controller.B2SSetScorePlayer pl, PlayerScores(pl)
	EMReel1.SetValue PlayerScores(1)
	EMReel2.SetValue PlayerScores(2)
End Sub

Sub EndScoresTimer_Timer()
	ScoresRoll = 1
		If SubScores(1)+SubScores(2)+SubScores(3) = 0 Then
			If Rack = 0 and ExtraBalls = 0 Then
				If Players = 2 and Player = 1 Then
					Player = 2
					If B2SOn Then
						Controller.B2SSetPlayerUp 2
					End If
					EMReel1.Image = "NumbersDark"
					EMReel2.Image = "Numbers"
				Else
					Player = 1: Balls = Balls - 1
					If Balls = 0 Then GameOver
					Bonus
					If Balls > 0 Then
						TextBalls.Text = (6-Balls)
						EMReel1.Image = "Numbers"
						EMReel2.Image = "NumbersDark"
					End If
					If B2SOn and Balls > 0 Then
						Controller.B2SSetPlayerUp 1
						B2SBIP(6-Balls)
					End If
				End If
			End If 'If Rack = 0
			ScoresRoll = 0

			If ExtraBalls > 0 and Rack = 0 Then ExtraBalls = ExtraBalls - 1
			LightTriangle.State = 0
			LightLeft.State = 0
			LightRight.State = 0
			LightLeft1.State = 0
			LightRight1.State = 0
			If GameMode = 1 Then
				For each gLight in BallLights
					gLight.state = 1
				Next
				BallsLit = 0
			End If
			If BallsLit = 0 Then
				For each gLight in BallLights
					gLight.state = 1
				Next
			End IF
			tilt=False
			TILTreel.setvalue 0
			if b2son then controller.b2ssettilt 0
			EndScoresTimer.Enabled = false
		End If

  if tilt=False Then
	If SubScores(1) > 0 Then
		SubScores(1) = SubScores(1) - 100
		PlayerScores(SubPlayer) = PlayerScores(SubPlayer) + 100
		If B2SOn Then Controller.B2SSetScorePlayer SubPlayer, PlayerScores(SubPlayer)
		EMReel1.SetValue PlayerScores(1)
		EMReel2.SetValue PlayerScores(2)
		PlaySound SoundFXDOF("Chimes",109,DOFPulse,DOFChimes),0,1,0,0.25
	End If
	If SubScores(2) > 0 and SubScores(1) = 0 Then
		SubScores(2) = SubScores(2) - 1000
		PlayerScores(SubPlayer) = PlayerScores(SubPlayer) + 1000
		If B2SOn Then Controller.B2SSetScorePlayer SubPlayer, PlayerScores(SubPlayer)
		EMReel1.SetValue PlayerScores(1)
		EMReel2.SetValue PlayerScores(2)
		PlaySound SoundFXDOF("Chimes",109,DOFPulse,DOFChimes),0,1,0,0.25
	End If
	If SubScores(3) > 0 and SubScores(1) = 0 and SubScores(2) = 0 Then
		SubScores(3) = SubScores(3) - 1000
		PlayerScores(SubPlayer) = PlayerScores(SubPlayer) + 1000
		If B2SOn Then Controller.B2SSetScorePlayer SubPlayer, PlayerScores(SubPlayer)
		EMReel1.SetValue PlayerScores(1)
		EMReel2.SetValue PlayerScores(2)
		PlaySound SoundFXDOF("Chimes",109,DOFPulse,DOFChimes),0,1,0,0.25
	End If
  Else
	for j=1 to 3: SubScores(j)=0: Next
  end if

End Sub

Sub ShowExtraBalls(ExtraBalls)
	TextShootAgain.Text = ""
	If B2SOn Then
		Controller.B2SSetShootAgain 0
		Controller.B2SSetData 42,0:Controller.B2SSetData 43,0:Controller.B2SSetData 44,0
	End If
	If ExtraBalls > 0 Then
		If B2SOn Then
			Controller.B2SSetShootAgain 1
			Controller.B2SSetData 41+ExtraBalls,1
		End if
		TextShootAgain.Text = ExtraBalls&" added Balls"
	End if
End Sub

Sub Drain_Hit()
	PlaySound "Drains",0,1,0,0.25
	Drain.DestroyBall
	WallsDown
	Disc.Image = "rotDiscBall"

	SubPlayer = Player
	EndScoresTimer.Enabled = true

	BIP = 0
	turnoff

	If Rack = 0 Then
	' PF Lights off
		LightTriangle.State = 0
		LightLeft.State = 0
		LightRight.State = 0
		LightLeft1.State = 0
		LightRight1.State = 0
	' ------------------
		If B2SOn Then
			For j = 6 to 10
				Controller.B2SSetdata j, 0
			Next
		End if
		TextRacks.Text = ""
		Racks = 0
	End If


	Rack = 0
End Sub

Sub GameOver
	PlaySound "GameOver",0,1,0.25,0.25
	' all BallLights off
	For each gLight in BallLights
		gLight.state = 0
	Next
	DiscTimer.Enabled = false
	stopsound "motor":DOF 105, DOFOff
	LightOn2.State = 0
	LightOn3.State = 0
	Players = 0
	Player = 0
	Lcanplay1.state=0
	Lcanplay2.state=0
	EMReel1.Image = "NumbersDark"
	EMReel2.Image = "NumbersDark"
	GOReel.setvalue 1
	TextBalls.Text = ""

	' all PF Lights off
		LightStar.State = 0
		LightTriangle.State = 0
		LightLeft.State = 0
		LightRight.State = 0
		LightLeft1.State = 0
		LightRight1.State = 0
	' ------------------
		StarReel.setvalue 0
	' check for highscore
	if PlayerScores(1)>hiscore then hiscore=PlayerScores(1)
	if PlayerScores(2)>hiscore then hiscore=PlayerScores(2)
	TextHiScore.text = hiscore
	savehs
	If B2SOn Then
		Controller.B2SSetData 16, 0
		Controller.B2SSetGameover 1
		Controller.B2SSetPlayerUp 0
		Controller.B2SSetCanPlay 0
		Controller.B2SSetScorePlayer 5, hiscore
		B2SBIP(0)
	End If

End Sub

Sub TriggerLeft_Hit()
  if tilt=false Then
	Scores Player,100
	If LightLeft.State = 1 Then Scores Player,900
	PlaySound SoundFXDOF("Chimes",109,DOFPulse,DOFChimes),0,1,0,0.25
  end if
End Sub

Sub TriggerRight_Hit()
  if tilt=false Then
	Scores Player,100
	If LightLeft.State = 1 Then Scores Player,900
	PlaySound SoundFXDOF("Chimes",109,DOFPulse,DOFChimes),0,1,0,0.25
  end if
End Sub

Sub Rubber1_Hit()
  if tilt=false Then
	Scores Player,100
	If LightLeft.State = 1 Then Scores Player,900
	StarCheck
	PlaySound SoundFXDOF("Chimes",109,DOFPulse,DOFChimes),,0,1,0,0.25
  end if
End Sub

Sub Rubber2_Hit()
  if tilt=false Then
	Scores Player,100
	If LightLeft.State = 1 Then Scores Player,900
	StarCheck
	PlaySound SoundFXDOF("Chimes",109,DOFPulse,DOFChimes),0,1,0,0.25
  end if
End Sub

Sub StarCheck()
	StarLit = StarLit + 1
	If StarLit = 10 Then
		LightStar.State = 1: StarLit = 0
		StarReel.setvalue 1
		If B2SOn Then Controller.B2SSetData 16, 1
	Else
		LightStar.State = 0
		StarReel.setvalue 0
		If B2SOn Then Controller.B2SSetData 16,0
	End If
End Sub

' *** Balls Triggers 1 to 15

Sub CheckBalls()

	If Rack = 0 Then BallsLit = BallsLit + 1

	If BallsLit < 16 and Rack = 0 Then
		If KnockerOn Then PlaySound SoundFXDOF("knocker",104,DOFPulse,DOFKnocker)
		'Scores Player,100
		SubScores(1) = SubScores(1) + 100
	End if

	If BallsLit = 15 and Rack = 0 Then

		' Get 10.000 Extra, when Star Lit
		If LightStar.State = 1 Then
			'Scores Player, 10000
			SubScores(3) = SubScores(3) + 10000
			PlaySound SoundFXDOF("knocker",104,DOFPulse,DOFKnocker),3,1,0.25,0.25
            DOF 108, DOFPulse
		End If

		' Get Bonus
		'Scores Player, BonusScore
		SubScores(2) = SubScores(2) + BonusScore

		LightTriangle.State = 1
		LightLeft.State = 1
		LightRight.State = 1
		LightLeft1.State = 1
		LightRight1.State = 1
		DOF 108, DOFPulse
		Rack = 1: Racks = Racks + 1
		ExtraBalls = 1
		If Racks = 3 Then ExtraBalls = 2 'Extraballs UP
		If Racks = 5 Then ExtraBalls = ExtraBalls + 1
		If Racks = 7 Then Extraballs = 3
		If Racks = 8 Then Racks = 7

		ShowExtraBalls(ExtraBalls)

		BallsLit = 0

'		For each gLight in BallLights
'			gLight.state = 1
'		Next

		TextRacks.Text = Racks
		If B2SOn Then
			For j = 6 to 12
				Controller.B2SSetdata j, 0
			Next
			Controller.B2SSetData 5+Racks, 1
		End If

	End If
End Sub

Sub Triggers_hit(idx)
	dim rotball
	rotball=idx+1
	if moveballs=1 Then
		direction(rotball)=Int(Rnd*2)-1
		if direction(rotball)=0 then direction(rotball)=1
		if ballangle(rotball)=0 then
			ballangle(rotball)=(12*direction(rotball))
			EVAL("Primball"&rotball).roty=ballangle(rotball)
		end if
	end If
	If Rack = 0 and BIP = 1 and tilt=false Then
		If EVAL("Light"&rotball).State = 1 Then
			DOF 200+idx, DOFPulse
			If KnockerOn = False Then DOF 300+idx, DOFPulse
			EVAL("Light"&rotball).State = 0
			CheckBalls
		End If
	End If

End Sub

Sub BallTurnTimer_Timer()
	dim x
	for x=1 to 15
		if ballangle(x)<>0 Then
			ballangle(x)=ballangle(x)+(12*direction(x))
			if (ballangle(x) > 359) or (ballangle(x) < -359) Then
				ballangle(x) = 0
			End If
			EVAL("Primball"&x).roty=ballangle(x)
		end if
	next
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

sub savehs
    savevalue "HiScorePool", "credit", Credits
    savevalue "HiScorePool", "hiscore", hiscore
	savevalue "HiScorePool", "mode", GameMode
	savevalue "HiScorePool", "motor", motorsound
	savevalue "HiScorePool", "moveballs", moveballs
	savevalue "HiScorePool", "knockeron", KnockerOn
end sub

sub loadhs
    dim temp
	temp = LoadValue("HiScorePool", "credit")
    If (temp <> "") then Credits = CDbl(temp)
    temp = LoadValue("HiScorePool", "hiscore")
    If (temp <> "") then hiscore = CDbl(temp)
    temp = LoadValue("HiScorePool", "mode")
    If (temp <> "") then GameMode = CDbl(temp)
    temp = LoadValue("HiScorePool", "motor")
    If (temp <> "") then motorsound = CDbl(temp)
    temp = LoadValue("HiScorePool", "moveballs")
    If (temp <> "") then moveballs = CDbl(temp)
    temp = LoadValue("HiScorePool", "knockeron")
    If (temp <> "") then KnockerOn = CDbl(temp)
end sub

Sub Table1_Exit()
	turnoff
	savehs
	If B2SOn Then Controller.Stop
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
    Dim BOT, b, ll
    BOT = GetBalls
	For ll = 1 to 15
		EVAL("Lighta"&ll).state=EVAL("light"&ll).state
		EVAL("Lightb"&ll).state=EVAL("light"&ll).state
	Next

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

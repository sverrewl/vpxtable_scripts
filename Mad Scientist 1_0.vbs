Option Explicit
Randomize

' Thalamus 2018-07-23
' Added/Updated "Positional Sound Playback Functions" and "Supporting Ball & Sound Functions"
' Thalamus 2018-11-01 : Improved directional sounds
' !! NOTE : Table not verified yet !!
' TODO need to go throug again - don't really know where the objects are located.

' Options
' Volume devided by - lower gets higher sound

Const VolDiv = 2000    ' Lower number, louder ballrolling/collition sound
Const VolCol = 10      ' Ball collition divider ( voldiv/volcol )

' The rest of the values are multipliers
'
'  .5 = lower volume
' 1.5 = higher volume

Const VolGates  = 1    ' Gates volume.
Const VolTarg   = 1    ' Targets volume.
Const VolKick   = 1    ' Kicker volume.
Const VolFlip   = 1    ' Flipper volume.


Const BallSize = 50
Const BallMass = 1.7

Dim ContrastSetting, GlowAmountDay, InsertBrightnessDay
Dim GlowAmountNight, InsertBrightnessNight

' *** Contrast level, possible values are 0 - 7, can be done in game with magnasave keys **
' *** 0: bright, good for desktop view, daytime settings in insert lighting below *********
' *** 1: same as 0, but with nighttime settings in insert lighting below ******************
' *** 2: darker, better contrast, daytime settings in insert lighting below ***************
' *** 3: same as 2, but with nighttime settings in insert lighting below ******************
' *** etc for 4-7; default is 3 ***********************************************************
ContrastSetting = 0

' *** Insert Lighting settings ************************************************************
' *** The settings below together with ContrastSetting determines how the lighting looks **
' *** for all values: 1.0 = default, useful range 0.1 - 5 *********************************
GlowAmountDay = 0.5
InsertBrightnessDay = 0.8
GlowAmountNight = 0.5
InsertBrightnessNight = 0.6


'DOF mapping by Arngrim
'101 - Left Flipper1
'102 - Right Flipper1
'103 - Left Slingshot
'104 - Left Slingshot Flasher
'105 - Right Slingshot
'106 - Right Slingshot Flasher
'107 - Bumper Left
'108 - Bumper Left Flasher
'109 - Bumper Center
'110 - Bumper Center Flasher
'111 - Bumper Right
'112 - Bumper Right Flasher
'113 - Ballrelease
'114 - Plunger
'115 - Drain
'116 - Credits Button
'117 - Knocker
'118 - Strobe 300 ms for knocker and kickers
'119 - Auto Drop Target and Reset
'120 - Drop Targets Hit
'121 - Techpoint Diverter
'122 - MercuryBath and kicker
'123 - kicker1
'124 - IgorSink
'125 - KickerEinsteinium
'126 - vuk
'127 - WallSink
'128 - Ramps done
'129 - Spinner1
'130 - Spinner2
'131 - Spinner3

'DOF MX Leds mapping by TerryRed

'201 - Left Flipper
'202 - Right Flipper
'203 - Left Slingshot
'204 - Right Slingshot
'205 - Bumper Left
'206 - Bumper Center
'207 - Bumper Right
'208 - Ballrelease
'209 - Plunger
'210 - Drained
'216 - Credits Inserted
'217 - Left Outlane - Green
'218 - Left InLane - Yellow
'219 - Right InLane - Red Flask
'220 - Right OutLane - Blue Flask
'221 - Nudge
'222 - TILT
'223 - Launch Button
'224 - Spinner1
'225 - Spinner2
'226 - Spinner3
'227 - MONSTER Drop Target M
'228 - MONSTER Drop Target O
'229 - MONSTER Drop Target N
'230 - MONSTER Drop Target S
'231 - MONSTER Drop Target T
'232 - MONSTER Drop Target E
'233 - MONSTER Drop Target R
'234 - Strobe 300 ms for knocker and kickers
'235 - Beacon  (not used yet)
'236 - Plunger Lane - Green Flask
'238 - Left Ramp - Trigger 1
'239 - Right Ramp - Trigger 2
'240 - Skillshot - SKILL flash
'241 - Jackpot
'242 - Electrical Effects
'250 - MONSTER flash
'251 - Extra Ball



'******************************************************
' 				VARIABLE DECLARATIONS
'******************************************************

Const cGameName = "madscientist"

On Error Resume Next
ExecuteGlobal GetTextFile("core.vbs")
If Err Then MsgBox "Can't open core.vbs"
On Error Goto 0

On Error Resume Next
ExecuteGlobal GetTextFile("controller.vbs")
If Err Then MsgBox "Can't open controller.vbs"
On Error Goto 0

'********* UltraDMD **************
Dim UltraDMD:UltraDMD=0
Const UltraDMD_VideoMode_Stretch = 0
Const UltraDMD_VideoMode_Top = 1
Const UltraDMD_VideoMode_Middle = 2
Const UltraDMD_VideoMode_Bottom = 3
Const UltraDMD_Animation_FadeIn = 0
Const UltraDMD_Animation_FadeOut = 1
Const UltraDMD_Animation_ZoomIn = 2
Const UltraDMD_Animation_ZoomOut = 3
Const UltraDMD_Animation_ScrollOffLeft = 4
Const UltraDMD_Animation_ScrollOffRight = 5
Const UltraDMD_Animation_ScrollOnLeft = 6
Const UltraDMD_Animation_ScrollOnRight = 7
Const UltraDMD_Animation_ScrollOffUp = 8
Const UltraDMD_Animation_ScrollOffDown = 9
Const UltraDMD_Animation_ScrollOnUp = 10
Const UltraDMD_Animation_ScrollOnDown = 11
Const UltraDMD_Animation_None = 14
Const UltraDMD_deOn = 1500

'********* End UltraDMD **************

Const movieSpeed= 76 'msecs
Dim nvcredits: nvcredits = 0
Dim nvBallsPerGame:nvBallsPerGame = 3
' Define any Constants
Const constMaxPlayers = 1									' Maximum number of players per game - only Single player
' Define Game Control Variables
Dim BallsOnPlayfield										' number of balls on playfield
Dim ModeTime
Dim vpTilted:vpTilted=False
Dim vpGameInPlay:vpGameInPlay=False
Dim BallInPlay:BallInPlay=0
Dim Score:Score=0
' Define Game Flags
Dim bFreePlay:bFreePlay = True								' Either in Free Play or Handling Credits
Dim bBallInPlungerLane										' Is there a ball in the plunger lane ?
Dim bEnteringAHighScore:bEnteringAHighScore = FALSE			' player is entering their name into the high score table
Dim CurrentMusicTunePlaying:CurrentMusicTunePlaying = 0
Dim nvjackpot:nvjackpot = 0

'******************************************************
' 					TABLE INIT
'******************************************************


Sub LoadTable()
	'LoadEM
	LoadUltraDMD
	PlaySound "Tabd58"
	EndOfGame 1

	wall24.collidable = 0
	wall7.collidable = 1
	splash1.visible = false
	splash2.visible = false
	splash3.visible = false

	If bFreePlay Then DOF 116, DOFOn
End Sub


Dim AttractVideo
AttractVideo = Array("intro ", 76)

Sub Table1_Init()
	LoadEM
	LoadHighScores
	PlaySound "intro"' INTRO SFX***********************************
	attractdelay.Interval = 10
	attractdelay.Enabled = 1
	Blockwall.visible = 1
	MovieWall.visible = 1
	VideoIntro.interval = 5500
	VideoIntro.enabled = true
	Animation AttractVideo(0), AttractVideo(1), 1
	ColorGrade
End Sub

Sub Table1_Exit()
	If B2SOn Then Controller.Stop
	if VideoIntro.enabled = false then
		If UltraDMD.IsRendering Then UltraDMD.CancelRendering
		UltraDMD = Null
		SaveHighScores
	End If
End Sub

Sub Animation(name, numframes, loops)
	HSteps = numframes
	HPos=0
	Hname = name
	Hloops = loops
	posinc = 1
    holotimer.interval = movieSpeed
	holotimer.enabled=1
End Sub

Dim Hname, Hsteps, Hloops, Hpos, posinc

Sub holotimer_timer()
	Dim imagename
	HPos=(HPos+posinc) mod AttractVideo(1)
	if HPos = 0 then HPos = 1
	if HPos < 10 then
		imagename = Hname & "00" & Hpos
	elseif HPos < 100 then
		imagename = Hname & "0" & Hpos
	else
		imagename = Hname & Hpos
	end if
	Moviewall.image = imagename
end Sub

Sub VideoIntro_timer()
	holotimer.enabled = False
	Blockwall.visible = False
	MovieWall.visible = False
	LoadTable
	me.enabled = False
End Sub

Sub ChangeGlow(day)
	Dim Light
	If day Then
		For Each Light in GlowLights : Light.IntensityScale = GlowAmountDay: Light.FadeSpeedUp = Light.Intensity * GlowAmountDay / 2: Light.FadeSpeedDown = Light.FadeSpeedUp / 5: Next 'fadespeeddown = fadespeedup / 25
		For Each Light in InsertLights : Light.IntensityScale = InsertBrightnessDay : Light.FadeSpeedUp = Light.Intensity * InsertBrightnessDay / 2 : Light.FadeSpeedDown = Light.FadeSpeedUp / 5 : Next 'fadespeeddown = fadespeedup / 25
	Else
		For Each Light in GlowLights : Light.IntensityScale = GlowAmountNight: Light.FadeSpeedUp = Light.Intensity * GlowAmountNight / 2: Light.FadeSpeedDown = Light.FadeSpeedUp / 5: Next 'fadespeeddown = fadespeedup / 25
		For Each Light in InsertLights : Light.IntensityScale = InsertBrightnessNight : Light.FadeSpeedUp = Light.Intensity * InsertBrightnessNight / 2 : Light.FadeSpeedDown = Light.FadeSpeedUp / 5 : Next 'fadespeeddown = fadespeedup / 25
	End If
End Sub

Sub ColorGrade()
	Dim lutlevel, ContrastLut
	Lutlevel = 0
	If ContrastSetting=0 or ContrastSetting=2 or ContrastSetting=4 or ContrastSetting=6 Then
		ChangeGlow(True)
		ContrastLut = ContrastSetting / 2
	Else
		ChangeGlow(False)
		ContrastLut = ContrastSetting / 2 - 0.5
	End If
	table1.ColorGradeImage = "LUT" & ContrastLut & "_" & lutlevel
End Sub

'******************************************************
' 						KEYS
'******************************************************

Dim LeftFlipperDown, RightFlipperDown

Sub Table1_KeyDown(ByVal KeyCode)

	If (KeyCode = AddCreditKey) or (KeyCode = AddCreditKey2) Then
		PlaySoundAtVol "CoinIn", Drain, 1
		CreditsTimer.Enabled=True
		If (vpGameInPlay = False) Then
		DOF 216, DOFOn  'DOF MX - Credits Inserted, Ready To Play
		End If
	End if

	If (KeyCode = PlungerKey) and (vpGameInPlay = TRUE) and (bEnteringAHighScore=False) and (BonusTimer.Enabled=False) then
		PlungerState(1)
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

	If keycode = MechanicalTilt Then
		checktilt
	End If

	If KeyCode = LeftFlipperKey Then
		LeftFlipperDown = 1
	End If

	If KeyCode = RightFlipperKey Then
		RightFlipperDown = 1
	End If


	If keycode = RightMagnaSave Then
		ContrastSetting = ContrastSetting + 1
		If ContrastSetting > 7 Then ContrastSetting = 7 End If
		ColorGrade
	End If
	If keycode = LeftMagnaSave Then
		ContrastSetting = ContrastSetting - 1
		If ContrastSetting < 0 Then ContrastSetting = 0 End If
		ColorGrade
	End If

	If (vpGameInPlay = TRUE) Then
		If (vpTilted = FALSE) Then

			If (KeyCode = LeftFlipperKey) and (bEnteringAHighScore=False) and (BonusTimer.Enabled=False) Then
				PlaySoundAtVol SoundFXDOF("fx_flipperup",101,DOFOn,DOFFlippers), LeftFlipper, VolFlip
				DOF 201, DOFPulse  'DOF MX - Left Flipper
				LeftFlipper.RotateToEnd
				Flipper2.RotateToEnd
				if flaskactive = 1 then
					MixElements()
				end if
			Elseif (KeyCode = LeftFlipperKey) and (bEnteringAHighScore=True) And NOT cursorPos=40 Then
				inChar=inChar-1
				if inChar<65 AND inChar>57 Then inChar=57
				if inChar<46 Then inChar=91
				NameEntry:Playsound "fx_Next"
			End If

			If (KeyCode = RightFlipperKey) and (bEnteringAHighScore=False) and (BonusTimer.Enabled=False) Then
				PlaySoundAtVol SoundFXDOF("fx_flipperup",102,DOFOn,DOFFlippers), RightFlipper, VolFlip
				DOF 202, DOFPulse  'DOF MX - Right Flipper
				RightFlipper.RotateToEnd
				Flipper1.RotateToEnd
				if flaskactive=1 then
					ClearFlask()
				end if
			Elseif (KeyCode = RightFlipperKey) and (bEnteringAHighScore=True) And NOT cursorPos=40 Then
				inChar=inChar+1
				if inChar>91 Then inChar=46
				if inChar>57 AND inChar<65 Then inChar=65
				NameEntry:Playsound "fx_Previous"
			End If

			If keycode = StartGameKey AND NOT cursorPos=40 Then
				if (inChar = 91 and cursorPos > 0) Then				'Set the backspace character
					cursorPos = cursorPos - 1						'Send cursor back
					initials(cursorPos) = 32						'Set that initial back to an empty SPACE
					Playsound "fx_Esc"
				End If
				if (inChar <> 91) Then								'Set a character, as long as it's not a backspace
					initials(cursorPos) = inChar					'Set the character
					If inChar = 47 Then initials(cursorPos) = 32	'Set the space character
					cursorPos = cursorPos + 1
					Playsound "fx_Enter"
				End If
				if (cursorPos < 4) Then								'write the 3 characters
					NameEntry
					If cursorPos=3 Then
						cursorPOS=40
						ExitHS.Interval=500
						ExitHS.Enabled=1
					End If
				End If
			End If
		End if
	Else
		If (KeyCode = (StartGameKey)) Then

			If (bFreePlay = TRUE) Then

				DOF 216, DOFOff  'DOF MX - Start Game - Credits Off

				If (BallsOnPlayfield = 0) or (vpTilted=False) Then
					NewGame()

				End If
			Else

				If (nvCredits > 0)  Then

					If (BallsOnPlayfield = 0) or (vpTilted=False) Then

						nvCredits = nvCredits - 1
						If nvCredits < 1 Then DOF 116, DOFOff
						DOF 216, DOFOff  'DOF MX - Start Game - Credits Off
						NewGame()

					End if
				Else
					DMD_CancelRendering
					DMD_DisplayScene "","Insert Coin", UltraDMD_Animation_None, UltraDMD_deOn, UltraDMD_Animation_None
				End if
			End if
		End if
	End if

	If (vpGameInPlay = False) and (vpTilted = FALSE) Then 'play random sound on flipper press during GameOver
		If (KeyCode = (LeftFlipperKey) OR KeyCode = (RightFlipperKey)) and (bEnteringAHighScore=False) and (BonusTimer.Enabled=False) Then
			GameoverTimer.interval=(Int((1444)*Rnd+333)):GameoverTimer.enabled=True
		End If



	End if

End Sub

Sub Table1_KeyUp(ByVal KeyCode)
	If KeyCode = LeftFlipperKey Then
		LeftFlipperDown = 0
		DOF 101, DOFOff
	End If

	If KeyCode = RightFlipperKey Then
		RightFlipperDown = 0
		DOF 102, DOFOff
	End If


	If (KeyCode = (PlungerKey)) and (vpGameInPlay = TRUE) and (bEnteringAHighScore=False) and (BonusTimer.Enabled=False) then
		PlungerState(0)
		DOF 236, DOFOff  'DOF MX - Plunger Lane - Green flask
	End If

	if (vpGameInPlay = TRUE) and (vpTilted = FALSE) Then
		If (KeyCode = (LeftFlipperKey)) Then
			PlaySoundAtVol "fx_flipperdown", LeftFlipper, VolFlip
			LeftFlipper.RotateToStart
			Flipper2.RotateToStart
		End If

        If (KeyCode = (RightFlipperKey)) Then
			PlaySoundAtVol "fx_flipperdown", RightFlipper, VolFlip
			RightFlipper.RotateToStart
			Flipper1.RotateToStart
		End If
	End If
End Sub

'******************************************************
' 						SCORING
'******************************************************

Sub AddScore(points)
	If (vpGameInPlay=TRUE And vpTilted = FALSE) Then
		If SkillshotEnabled Then
			ScoreSkillShot("Any")
		Else
			Score = Score + (points)*TableMultiplier

		End if
	End If
	ScoreTimer.enabled = True
End Sub

Sub DisplayScore()
	Dim FlaskContents, BallDisplay
	If battongue = 1 or einsteinium = 1 or lecithin = 1 or yttrium = 1 Then
		FlaskContents = "Flask:"
		If battongue = 1 then FlaskContents = FlaskContents & "B"
		If einsteinium = 1 then FlaskContents = FlaskContents & "E"
		If lecithin = 1 then FlaskContents = FlaskContents & "L"
		If yttrium = 1 then FlaskContents = FlaskContents & "Y"
	Else
		FlaskContents = ""
	End If
	If BallInPlay = 0 then
		BallDisplay = "Game Over"
	Else
		BallDisplay = "Ball " & BallInPlay
	End If
	'if UltraDMD.GetMinorVersion < 4 then
		UltraDMD.DisplayScoreboard 2, 1, Score, nvjackpot, 0, 0, BallDisplay, FlaskContents & " IQ:"&IQ
	'Else
	'	UltraDMD.DisplayScoreboard00 2, 1, Score, nvjackpot, 0, 0, BallDisplay, FlaskContents & " IQ:"&IQ
	'End If
End Sub

Sub AddJackpot(score)
	If (vpTilted = False) and (vpGameInPlay=TRUE) Then
		nvJackpot = nvJackpot + score

		DisplayJackpot
	End if
End Sub

Sub DisplayJackpot()
	If Not UltraDMD.IsRendering Then DMD_DisplaySceneExWithId "jackpot","Jackpot",nvjackpot,UltraDMD_Animation_ScrollOnUp,UltraDMD_deOn,UltraDMD_Animation_None
	DMD_ModifyScene "jackpot","Jackpot",nvjackpot
End Sub

'******************************************************
' 						SLINGSHOTS
'******************************************************

Dim RStep, Lstep

Sub LeftSlingshot_Slingshot()
	PlaySoundAtVol SoundFXDOF("fx_slingshot",103,DOFPulse,DOFContactors), sling2, 1
	DOF 104, DOFPulse
	DOF 203, DOFPulse  'DOF MX - Left Slingshot
	LSling.Visible = 0
	LSling1.Visible = 1
	sling2.TransZ = -20
	LStep = 0
	LeftSlingShot.TimerEnabled = 1
	FlashForMs LightLeftsling, 200, 100, LightStateOn
	FlashForMs LeftSlingshotBulb1, 200, 100, LightStateOn
	FlashForMs LeftSlingshotBulb2, 200, 100, LightStateOn
	AddScore 400
End Sub

Sub LeftSlingShot_Timer
	Select Case LStep
		Case 3:LSLing1.Visible = 0:LSLing2.Visible = 1:sling2.TransZ = -10
		Case 4:LSLing2.Visible = 0:LSLing.Visible = 1:sling2.TransZ = 0:LeftSlingShot.TimerEnabled = 0
	End Select
	LStep = LStep + 1
End Sub

Sub RightSlingshot_Slingshot()
	PlaySoundAtVol SoundFXDOF("fx_slingshot",105,DOFPulse,DOFContactors), sling1, 1
	DOF 106, DOFPulse
	DOF 204, DOFPulse  'DOF MX - Right Slingshot
	RSling.Visible = 0
	RSling1.Visible = 1
	sling1.TransZ = -20
	RStep = 0
	RightSlingShot.TimerEnabled = 1
	FlashForMs LightRightsling, 200, 100, LightStateOn
	FlashForMs RightSlingshotBulb1, 200, 100, LightStateOn
	FlashForMs RightSlingshotBulb2, 200, 100, LightStateOn
	AddScore 400
End Sub

Sub RightSlingShot_Timer
	Select Case RStep
		Case 3:RSLing1.Visible = 0:RSLing2.Visible = 1:sling1.TransZ = -10
		Case 4:RSLing2.Visible = 0:RSLing.Visible = 1:sling1.TransZ = 0:RightSlingShot.TimerEnabled = 0
	End Select
	RStep = RStep + 1
End Sub

'******************************************************
' 						BUMPERS
'******************************************************

Dim BumperHits:BumperHits=0

Sub Bumper1_Hit()
	if (vpGameInPlay = TRUE) And (vpTilted = FALSE) then
		DOF 111, DOFPulse
		DOF 112, DOFPulse
		DOF 207, DOFPulse  'DOF MX - Right Bumper
		FlashForMs LightBumper1, 200, 100, LightStateOff
		BumperScore
	end if
End Sub

Sub Bumper2_Hit()
	if (vpGameInPlay = TRUE) And (vpTilted = FALSE) then
		DOF 107, DOFPulse
		DOF 108, DOFPulse
		DOF 205, DOFPulse  'DOF MX - Left Bumper
		FlashForMs LightBumper2, 200, 100, LightStateOff
		BumperScore
	end if
End Sub

Sub Bumper3_Hit()
	if (vpGameInPlay = TRUE) And (vpTilted = FALSE) then
		DOF 109, DOFPulse
		DOF 110, DOFPulse
		DOF 206, DOFPulse  'DOF MX - Center Bumper
		FlashForMs LightBumper3, 200, 100, LightStateOff
		BumperScore
	end if
End Sub

Sub BumperScore()
	BumperHits = BumperHits + 1
	If BumperHits Mod 10 = 0 and BumperHits < 100 then
		playsound "tabd02"
		if Not UltraDMD.isRendering Then DMD_DisplayScene "Advance Bumper","Multiplier " & (BumperHits/10) + 1 &"X",UltraDMD_Animation_ScrollOnLeft, UltraDMD_deOn, UltraDMD_Animation_None
	If BumperHits <= 10 Then AddScore 1000
	Elseif BumperHits <= 20 Then AddScore 2000
	Elseif BumperHits <= 30 Then AddScore 3000
	Elseif BumperHits <= 40 Then AddScore 4000
	Elseif BumperHits <= 50 Then AddScore 5000
	Elseif BumperHits <= 60 Then AddScore 6000
	Elseif BumperHits <= 70 Then AddScore 7000
	Elseif BumperHits <= 80 Then AddScore 8000
	Elseif BumperHits <= 90 Then AddScore 9000
	Else AddScore 10000

	End If
End Sub


'******************************************************
' 						SPINNERS
'******************************************************

Sub Spinner1_spin():SpinnerScore(1):DOF 129, DOFPulse:DOF 224, DOFPulse:End Sub
Sub Spinner2_spin():SpinnerScore(0):DOF 130, DOFPulse:DOF 225, DOFPulse:End Sub
Sub Spinner3_spin():SpinnerScore(1):DOF 131, DOFPulse:DOF 226, DOFPulse:End Sub

Sub SpinnerScore(flash)
	if flash and Countdown.enabled = false and CurExperiment <> "MonsterToLife" then
		FlashForMs Lning2, 1700, 100, LightStateOff
		FlashForMs Lning1, 1700, 100, LightStateOff
	end if
	PlaySound "Tabd34"
	AddScore 2000
	AddJackpot 10000
End Sub

'******************************************************
' 					PLASMA TARGETS
'******************************************************

Sub plasma_Dropped()
	DropTargetSound
	if (vpGameInPlay = TRUE And vpTilted = FALSE) then
		If Countdown.enabled = true and CurExperiment = "AngryMob" Then
			Light25.state = LightStateOff:ScoreRedTarget
		Else
			Light25.state = LightStateOn
			AddScore 25000
			CheckTargetsPlasmafields()
		End If
	end if
End sub

Sub field_Dropped()
	DropTargetSound
	if (vpGameInPlay = TRUE And vpTilted = FALSE) then
		If Countdown.enabled = true and CurExperiment = "AngryMob" Then
			Light26.state = LightStateOff:ScoreRedTarget
		Else
			Light26.state = LightStateOn
			AddScore 25000
			CheckTargetsPlasmafields()
		End If
	end if
End sub

Sub CheckTargetsPlasmafields()
	if field.isdropped And plasma.isdropped Then
		plasmatext.visible=false
		PlaySound "Tabd22"
		AddScore 150000
		If Not PlasmaFieldEnabled Then
			EnablePlasmaField
		End If
	End if
End sub

'******************************************************
' 					PLASMA FIELD
'******************************************************

Dim PlasmaFieldEnabled,PlasmaMiddleEnabled,PlasmaRightEnabled,PlasmaLeftEnabled
PlasmaFieldEnabled=0:PlasmaMiddleEnabled=0:PlasmaRightEnabled=0:PlasmaLeftEnabled=0

dim plasmarightcount

Sub plasma_effect_timer()
	plasmarightcount = plasmarightcount+1
	if plasmarightcount = 6 then plasmarightcount=1

	Select Case (plasmarightcount)
		Case 1: Flasher_R0.visible = true
		Case 2: Flasher_R0.visible=false:Flasher_R2.visible=true
		Case 3:	Flasher_R2.visible=false:Flasher_R3.visible=true
		Case 4:	Flasher_R3.visible=false:Flasher_R2.visible=true
		Case 5: Flasher_R2.visible=false:Flasher_R0.visible=true
	End Select
End Sub

dim plasmamiddlecount

Sub plasma_effectM_timer()
	plasmamiddlecount = plasmamiddlecount+1
	if plasmamiddlecount = 5 then plasmamiddlecount=1

	Select Case (plasmamiddlecount)
		Case 1: Flasher_M1.visible=false:Flasher_M2.visible=true
		Case 2:	Flasher_M2.visible=false:Flasher_M3.visible=true
		Case 3:	Flasher_M3.visible=false:Flasher_M2.visible=true
		Case 4: Flasher_M2.visible=false:Flasher_M1.visible=true

	End Select
End Sub

dim plasmaleftcount

Sub plasma_effectL_timer()
	plasmaleftcount = plasmaleftcount+1
	if plasmaleftcount = 6 then plasmaleftcount=1

	Select Case (plasmaleftcount)
		Case 1: Flasher_L1.visible = true
		Case 2: Flasher_L1.visible=false:Flasher_L2.visible=true
		Case 3:	Flasher_L2.visible=false:Flasher_L3.visible=true
		Case 4:	Flasher_L3.visible=false:Flasher_L2.visible=true
		Case 5: Flasher_L2.visible=false:Flasher_L1.visible=true

	End Select

End Sub

Sub EnablePlasmaField()
	PlaySound "Tabd22"
	PlasmaFieldEnabled = 1
	PlasmaMiddleEnabled = 1
	plasmamiddlecount = 0
	plasma_effectM.enabled = true
	If Not UltraDMD.isRendering then DMD_DisplayScene "Plasmafield","Activated!!!", UltraDMD_Animation_ScrollOnRight , UltraDMD_deOn, UltraDMD_Animation_None

End Sub

Sub PlasmaMiddleTrig_hit()
	If PlasmaMiddleEnabled Then
		PlaySound "Tabd10"
		DOF 242, DOFPulse  'DOF MX - Electrical Effects
		PlasmaMiddleEnabled=0
		PlasmaMiddleEnabled=0
		plasma_effectM.enabled = false
		Flasher_M1.visible=false
		Flasher_M2.visible=false
		Flasher_M3.visible=false
		ScorePlasma

		PlasmaRightEnabled=1
		plasmarightcount = 0
		plasma_effect.enabled = true

	End If
End Sub

Sub PlasmaRightTrig_hit()
	If PlasmaRightEnabled Then
		PlaySound "Tabd10"
		DOF 242, DOFPulse  'DOF MX - Electrical Effects
		PlasmaRightEnabled=0
		plasma_effect.enabled = false
		Flasher_R0.visible=false
		Flasher_R2.visible=false
		Flasher_R3.visible=false
		ScorePlasma

		PlasmaLeftEnabled=1
		plasmaleftcount = 0
		plasma_effectL.enabled = true
	End If
End Sub

Sub PlasmaLeftTrig_hit()
	If PlasmaLeftEnabled Then
		PlaySound "Tabd10"
		DOF 242, DOFPulse  'DOF MX - Electrical Effects
		PlasmaLeftEnabled=0
		plasma_effectL.enabled = false
		Flasher_L1.visible=false
		Flasher_L2.visible=false
		Flasher_L3.visible=false

		ScorePlasma
		if PlasmaLeftEnabled = 0 and PlasmaMiddleEnabled = 0 and PlasmaRightEnabled = 0 then
			ResetPlasmaField
		end if
	End If
End Sub

Sub ScorePlasma()
		AddScore 250000
		If Not UltraDMD.isRendering then DMD_DisplayScene "Plasmafield Points",250000*TableMultiplier, UltraDMD_Animation_FadeIn, UltraDMD_deOn,UltraDMD_Animation_None
End Sub

Sub ResetPlasmaField()
	PlasmaFieldEnabled=0

	PlasmaLeftEnabled=0
	plasma_effectL.enabled = false
	Flasher_L1.visible=false
	Flasher_L2.visible=false
	Flasher_L3.visible=false

	PlasmaMiddleEnabled=0
	plasma_effectM.enabled = false
	Flasher_M1.visible=false
	Flasher_M2.visible=false
	Flasher_M3.visible=false

	PlasmaRightEnabled=0
	plasma_effect.enabled = false
	Flasher_R0.visible=false
	Flasher_R2.visible=false
	Flasher_R3.visible=false

	ResetPlasmaTargets
End Sub

Sub ResetPlasmaTargets()
	If PlasmaLeftEnabled = 0 and PlasmaMiddleEnabled = 0 And PlasmaRightEnabled = 0 Then
		DOF 119, DOFPulse
		If NOT (Countdown.enabled = true and CurExperiment = "AngryMob") Then
			field.IsDropped = 0:Light26.State = LightStateOff
			plasma.IsDropped = 0:Light25.State = LightStateOff
			plasmatext.visible=true
		End If
	Else
		Playsound SoundFXDOF("droptargetreset",119,DOFPulse,DOFContactors) ' TODO
		field.IsDropped = 1:Light26.State = LightStateOn
		plasma.IsDropped = 1:Light25.State = LightStateOn
	End If
End Sub

'******************************************************
' 					MONSTER TARGETS
'******************************************************

Dim TargetResets:TargetResets=0
Dim SkillshotEnabled:SkillshotEnabled=0

Sub TargetM_Dropped()
	DropTargetSound
	If (vpGameInPlay = TRUE) And (vpTilted = FALSE) Then
		DOF 227, DOFPulse   'DOF MX - Monster Target M
		If Countdown.enabled = true and CurExperiment = "AngryMob" Then
			LightM.state = LightStateOff:ScoreRedTarget
		Else
			If SkillshotEnabled Then ScoreSkillShot("M") Else MonsterTargetScore
			Switch1.rotx=180
		End If
	End if
End Sub

Sub TargetO_Dropped()
	DropTargetSound
	If (vpGameInPlay = TRUE) And (vpTilted = FALSE) Then
		DOF 228, DOFPulse   'DOF MX - Monster Target O
		If Countdown.enabled = true and CurExperiment = "AngryMob" Then
			LightO.state = LightStateOff:ScoreRedTarget
		Else
			If SkillshotEnabled Then ScoreSkillShot("O") Else MonsterTargetScore
			Switch2.rotx=180
		End If
	end if
End Sub

Sub TargetN_Dropped()
	DropTargetSound
	If (vpGameInPlay = TRUE) And (vpTilted = FALSE) Then
		DOF 229, DOFPulse   'DOF MX - Monster Target N
		If Countdown.enabled = true and CurExperiment = "AngryMob" Then
			LightN.state = LightStateOff:ScoreRedTarget
		Else
			If SkillshotEnabled Then ScoreSkillShot("N") Else MonsterTargetScore
			Switch3.rotx=180
		End If
	end if
End Sub

Sub TargetS_Dropped()
	DropTargetSound
	If (vpGameInPlay = TRUE) And (vpTilted = FALSE) Then
		DOF 230, DOFPulse   'DOF MX - Monster Target S
		If Countdown.enabled = true and CurExperiment = "AngryMob" Then
			Light_S.state = LightStateOff:ScoreRedTarget
		Else
			If SkillshotEnabled Then ScoreSkillShot("S") Else MonsterTargetScore
			Switch4.rotx=180
		End If
	end if
End Sub

Sub TargetT_Dropped()
	DropTargetSound
	If (vpGameInPlay = TRUE) And (vpTilted = FALSE) Then
		DOF 231, DOFPulse   'DOF MX - Monster Target T
		If Countdown.enabled = true and CurExperiment = "AngryMob" Then
			LightT.state = LightStateOff:ScoreRedTarget
		Else
			If SkillshotEnabled Then ScoreSkillShot("T") Else MonsterTargetScore
			Switch5.rotx=180
		End If
	end if
End Sub

Sub TargetE_Dropped()
	DropTargetSound
	If (vpGameInPlay = TRUE) And (vpTilted = FALSE) Then
		DOF 232, DOFPulse   'DOF MX - Monster Target E
		If Countdown.enabled = true and CurExperiment = "AngryMob" Then
			LightE.state = LightStateOff:ScoreRedTarget
		Else
			If SkillshotEnabled Then ScoreSkillShot("E") Else MonsterTargetScore
			Switch6.rotx=180
		End If
	End if
End Sub

Sub TargetR_Dropped()
	DropTargetSound
	If (vpGameInPlay = TRUE) and (vpTilted = FALSE) Then
		DOF 233, DOFPulse   'DOF MX - Monster Target R
		If Countdown.enabled = true and CurExperiment = "AngryMob" Then
			LightR.state = LightStateOff:ScoreRedTarget
		Else
			If SkillshotEnabled Then ScoreSkillShot("R") Else MonsterTargetScore
			Switch7.rotx=180
		End If
	end if
End Sub

Sub MonsterTargetScore()
	Dim ResetTargets

	If TargetM.isDropped and TargetO.isDropped and TargetN.isDropped and TargetS.isDropped and TargetT.isDropped and TargetE.isDropped and TargetR.IsDropped Then
		ResetTargets=1
		DOF 250, DOFPulse   'DOF MX - Monster Flash
	Else
		ResetTargets=0
	End If

	DropTargetLights

	If CurExperiment = "MonsterToLife" And E.state = LightStateOn  Then
		If ResetTargets = 1 Then
			PlaySound "Tabd61"
			Helper.collidable = true
			OpenTeleporter
			DMD_CancelRendering
			DMD_DisplayScene "Give Monster Life","Hit the Tower", UltraDMD_Animation_ZoomIn, UltraDMD_deOn*4, UltraDMD_Animation_None
			Lning1.state = LightStateBlinking
			Lning2.state = LightStateBlinking
			ResetTargets = 0
			diverter1p.rotx=0
			diverter2p.rotx=0
			diverter3p.rotx=0
			towertext.visible = True
			towertext.imageA = "a_tower"
		End If
		AddScore 100000
	Else
		If TargetResets < 1 Then
			AddScore 15000
			If ResetTargets Then
				AddScore 150000
				if Not UltraDMD.isRendering then DMD_DisplayScene "Teleporter Open",150000*TableMultiplier & " Points", UltraDMD_Animation_ScrollOnLeft, UltraDMD_deOn, UltraDMD_Animation_None
			End If
		ElseIf TargetResets < 2 Then
			AddScore 30000
			If ResetTargets Then
				AddScore 300000:
				if Not UltraDMD.isRendering then DMD_DisplayScene "Teleporter Open",300000*TableMultiplier & " Points", UltraDMD_Animation_ScrollOnLeft, UltraDMD_deOn, UltraDMD_Animation_None
			End If
		ElseIf TargetResets < 3 Then
			AddScore 45000
			If ResetTargets Then
				AddScore 450000
				if Not UltraDMD.isRendering then DMD_DisplayScene "Teleporter Open",450000*TableMultiplier & " Points", UltraDMD_Animation_ScrollOnLeft, UltraDMD_deOn, UltraDMD_Animation_None
			End If
		Else
			AddScore 60000
			If ResetTargets Then
				AddScore 600000
				if Not UltraDMD.isRendering then DMD_DisplayScene "Teleporter Open",600000*TableMultiplier & " Points", UltraDMD_Animation_ScrollOnLeft, UltraDMD_deOn, UltraDMD_Animation_None
			End If
		End If
	End If

	If ResetTargets Then
		TargetResets = TargetResets + 1
		ResetMonsterTargets
		OpenTeleporter
		PlaySound "Tabd05"
	End If

End Sub

Sub ResetMonsterTargets()
	targetstext.visible = true
	targetstext.imageA = "a_targets"
	PlaysoundAtVol SoundFXDOF("droptargetreset",119,DOFPulse,DOFContactors), Switch4, VolTarg
	TargetM.isDropped = 0:Switch1.rotx=90
	TargetO.isDropped = 0:Switch2.rotx=90
	TargetN.isDropped = 0:Switch3.rotx=90
	TargetS.isDropped = 0:Switch4.rotx=90
	TargetT.isDropped = 0:Switch5.rotx=90
	TargetE.isDropped = 0:Switch6.rotx=90
	TargetR.IsDropped = 0:Switch7.rotx=90

	TargetM.timerinterval = 300
	TargetM.timerenabled = true
End Sub

Sub TargetM_Timer()
	If NOT (Countdown.enabled = true and CurExperiment = "AngryMob") AND SkillshotEnabled <> 1 Then
		DropTargetLights()
	End If
	TargetM.timerenabled=false
End Sub

Sub DropMonsterTargets()
	TargetM.isDropped = 0:Switch1.rotx=90
	TargetO.isDropped = 0:Switch2.rotx=90
	TargetN.isDropped = 0:Switch3.rotx=90
	TargetS.isDropped = 0:Switch4.rotx=90
	TargetT.isDropped = 0:Switch5.rotx=90
	TargetE.isDropped = 0:Switch6.rotx=90
	TargetR.IsDropped = 0:Switch7.rotx=90
End Sub

Sub DropTargetLights()
	if CurExperiment = "MonsterToLife" and Lning1.state = LightStateOff Then
		If TargetM.isDropped Then LightM.state = LightStateOff Else LightM.state = LightStateBlinking
		If TargetO.isDropped Then LightO.state = LightStateOff Else LightO.state = LightStateBlinking
		If TargetN.isDropped Then LightN.state = LightStateOff Else LightN.state = LightStateBlinking
		If TargetS.isDropped Then Light_S.state = LightStateOff Else Light_S.state = LightStateBlinking
		If TargetT.isDropped Then LightT.state = LightStateOff Else LightT.state = LightStateBlinking
		If TargetE.isDropped Then LightE.state = LightStateOff Else LightE.state = LightStateBlinking
		If TargetR.isDropped Then LightR.state = LightStateOff Else LightR.state = LightStateBlinking
	Else
		LightM.state=TargetM.isDropped
		LightO.state=TargetO.isDropped
		LightN.state=TargetN.isDropped
		Light_S.state=TargetS.isDropped
		LightT.state=TargetT.isDropped
		LightE.state=TargetE.isDropped
		LightR.state=TargetR.isDropped
	End If
End Sub

'******************************************************
' 						SKILL SHOT
'******************************************************

Sub EnableSkillShot()
	dim xx

	SkillShotEnabled=1
	for each xx in aMonsterLights
		xx.state=LightStateBlinking
	Next
End Sub

Sub ScoreSkillShot(Letter)
	Dim SSPoints
	SkillShotEnabled=0

	DropTargetLights

	Select Case (Letter)
		Case "M":
			AddScore 1000000:SSPoints = "1000000"
			EnablePlasmaField
			DMD_CancelRendering
			DMD_DisplayScene "Activate Plasmafield", "1000000 Points", UltraDMD_Animation_ScrollOnUp, UltraDMD_deOn, UltraDMD_Animation_None
		Case "O":
			SSPoints="150,000"
			StartExperiment
		Case "N": AddScore 500000:SSPoints="500000"
		Case "S": AddScore 100000:SSPoints="100000"
		Case "T","E","R": AddScore 50000:SSPoints="50000"
		Case "Any": AddScore 10000
	End Select

	If NOT (Letter = "Any") Then
		PlaySound "skillshot"
		DOF 239, DOFPulse  'DOF MX - Skillshot
		If NOT UltraDMD.isRendering then DMD_DisplayScene "SKILLSHOT!!!", SSPoints &" Points", UltraDMD_Animation_ScrollOnUp, UltraDMD_deOn, UltraDMD_Animation_None
	End If
End Sub

Sub TeleporterGate1_hit()
	PlaySoundAtVol "fx_Gate", TeleporterGate1, VolGates
	If SkillshotEnabled = 1 Then
		EnableBallSaver
	End If
End Sub

Sub TeleporterGate_hit()
	PlaySoundAtVol "fx_Gate", TeleporterGate, VolGates
End Sub

Sub EnableBallSaver()
	BallSaver.Enabled=1
	BallSaver.interval=10000
	BallSaverLight.state=LightStateOn
End Sub

'******************************************************
' 						TRIGGERS
'******************************************************

'*** Inlanes

Sub LeftInlaneTrigger_Hit()
	ScoreInlane("L")
	DOF 218, DOFPulse  'DOF MX - Left Inlane - Yellow
End Sub

Sub RightInlaneTrigger_Hit()
	ScoreInlane("R")
	DOF 219, DOFPulse  'DOF MX - Right Inlane - Red
End Sub

Sub ScoreInlane(side)
	if (vpGameInPlay = TRUE And vpTilted = FALSE) then
		PlaySound "Tabd09"
		AddScore 10000
		if (Light23.state <> LightStateOff) and (Light24.state <> LightStateOff) then
			if NOT UltraDMD.isRendering then DMD_DisplayScene "Kick Back","Activated", UltraDMD_Animation_ScrollOnDown, UltraDMD_deOn, UltraDMD_Animation_None
			kickbackLight.state = LightStateOn
			PlaySound "Tabd19"
			Light23.state = LightStateOff
			Light24.state = LightStateOff
		Elseif side = "L" Then
			Light24.state = LightStateOn
		Elseif side = "R" Then
			Light23.state = LightStateOn
		End If
	End If
End Sub

Sub ResetReCapture()
	KickbackLight.State=LightStateOff
	Light7blue.State=LightStateOff
	Light20blue.State=LightStateOff
End Sub


Sub LeftOutLaneTrigger_Hit ()
	if (vpGameInPlay = TRUE And vpTilted = FALSE) then
		AddScore 25000
		PlaySoundAtVol "Tabd19", LeftOutLaneTrigger, 1
		DOF 217, DOFPulse  'DOF MX - Left Outlane - Green
		if (KickbackLight.state <> LightStateOff) then
			Light7blue.State = LightStateOn
			Light20blue.State = LightStateOn
			If Not UltraDMD.isRendering then DMD_DisplayScene "","Kick Back!!!", UltraDMD_Animation_ZoomIn, UltraDMD_deOn, UltraDMD_Animation_None
		End if
	end if
End Sub

Sub RightOutlaneTrigger_Hit()
	if (vpGameInPlay = TRUE) And (vpTilted = FALSE) then
		DOF 220, DOFPulse  'DOF MX - Right Outlane - Blue
		Lightblue.state = LightStateOn
		AddScore 25000
		PlaySoundAtVol "Tabd19", RightOutlaneTrigger, 1
	end if
End Sub

Sub RightOutlaneTrigger_Unhit()
	Lightblue.state = LightStateOff
End Sub

'******************************************************
' 						TELEPORTER
'******************************************************

Sub OpenTeleporter()
	TechpointDiverterp.roty = -30
	playsoundAtVol SoundFXDOF("fx_diverter",121,DOFPulse,DOFContactors), TechpointDiverterp, 1
	Wall7.collidable = 0
	wall24.collidable = 1
	poddoor2.visible = False
	BulbBlue1.state = lightstateblinking
	teletext.visible = true

	if Not UltraDMD.isRendering then DMD_DisplayScene "","Teleporter Open",UltraDMD_Animation_ScrollOnLeft, UltraDMD_deOn, UltraDMD_Animation_None

End Sub

Sub CloseTeleporter()
	TechpointDiverterp.roty = 22
	playsoundAtVol SoundFXDOF("fx_diverter",121,DOFPulse,DOFContactors), TechpointDiverterp, 1
	Wall7.collidable = 1
	wall24.collidable = 0
	poddoor2.visible = True
	BulbBlue1.state = lightstateoff
	teletext.visible=false
End Sub

Sub KickerTeleporter_Hit()
	CloseTeleporter
	Kicker1Timer.Enabled = true
	BulbBlue.state = LightStateOn
	LightClon.state = LightStateOn
	poddoor.visible = False
	PlaySoundAtVol "Tabd16", KickerTeleporter, 1
	AddScore 150000
End Sub

Dim BallImageCount
BallImageCount=0

Sub Kicker1Timer_Timer()
	BallImageCount = BallImageCount + 1
	If BallImageCount = 14 then BallImageCount = 1

	Dim Ball
	Set Ball = Kicker1.CreateSizedballWithMass(Ballsize/2,Ballmass)

	Select Case BallImageCount
		Case 1: Ball.image = "powerball"
		Case 2: Ball.image = "ball_eye3"
		Case 3: Ball.image = "ball_lizardeye"

		Case 4: Ball.image = "ball_plasma"
		Case 5: Ball.image = "ball_snakeeye2"
		Case 6: Ball.image = "ball_greeneye"
		Case 7: Ball.image = "ball_purpleeye"
		Case 8: Ball.image = "ball_blackeye"
		Case 9: Ball.image = "ball_eye"
		Case 10: Ball.image = "ball_snakeeye"
		Case 11: Ball.image = "ball_bloodshoteye"
		Case 12: Ball.image = "ball_snakeeye3"
		Case 13: Ball.image = "ball_eye2"
	End Select
	KickerTeleporter.DestroyBall
	Kicker1.kick 120, 10
	DOF 123, DOFPulse
	DOF 118, DOFPulse
	DOF 234, DOFPulse  'DOF MX - Strobe
	Kicker1Timer.Enabled = false
	if Not UltraDMD.isRendering then DMD_DisplayScene "Teleportation","Successful???", UltraDMD_Animation_ScrollOnLeft, UltraDMD_deOn, UltraDMD_Animation_None

	PlaySound "Tabd11"
	poddoor.visible=True
End sub

'******************************************************
' 					YTTRIUM RAMP
'******************************************************

Dim YttriumCount:YttriumCount=0

Sub Trigger1_Hit()
	If (vpGameInPlay=TRUE) And (vpTilted=FALSE) Then
		If activeball.velx > 0 or activeball.vely < 0  Then
			PlaySoundAtVol "Tabd24", Trigger1, 1
			DOF 128, DOFPulse
			DOF 238, DOFPulse   'DOF MX - Left Ramp
			If Countdown.enabled = false Then
				If yttrium = 0 then
					PlaySound "Tabd69"
					yttrium = 1
					if Not UltraDMD.isRendering then DMD_DisplayScene "Yttrium","Collected", UltraDMD_Animation_ScrollOnLeft, UltraDMD_deOn, UltraDMD_Animation_None
				End If
				If Light5.state = LightStateOff Then
					YttriumCount = 0
				End If
				YttriumCount = YttriumCount + 1
				If YttriumCount > 5 Then YttriumCount = 5
				FlashForMS Light5, 10000, 0, LightStateOff
				if YttriumCount > 1 and Not UltraDMD.isRendering then DMD_DisplayScene "Yttrium Combo", 40000*YttriumCount, UltraDMD_Animation_ScrollOnLeft, UltraDMD_deOn, UltraDMD_Animation_None
				AddScore 40000*YttriumCount
			Elseif CurExperiment = "ElixirOfLife" Then
				If yttrium = 0 Then
					DMD_CancelRendering
					DMD_DisplayScene "Yttrium","Collected", UltraDMD_Animation_ScrollOnLeft, UltraDMD_deOn, UltraDMD_Animation_None
					yttrium = 1
					ScoreElixirHit
				End If
			ElseIf CurExperiment = "RaidGraveyard" Then
				AddBodyPartYT
			Elseif 	CurExperiment = "BrainAndIQ" And Light5.state <> LightStateOff Then
				AddIQYttrium
			Else
				AddScore 40000
			end if
		End If
	End If
	UpdateElixir
End Sub

'******************************************************
' 					LECITHIN RAMP
'******************************************************

Dim LecithinCount:LecithinCount=0

Sub Trigger2_Hit()
	If (vpGameInPlay=TRUE) And (vpTilted=FALSE) Then
		If activeball.velx > 0 or activeball.vely < 0  Then
			PlaySoundAtVol "Tabd24", Trigger2, 1
			DOF 128, DOFPulse
			DOF 239, DOFPulse   'DOF MX - Right Ramp
			If Countdown.enabled = false Then
				If Lecithin = 0 then
					PlaySound "Tabd69"
					Lecithin = 1
					if Not UltraDMD.isRendering then DMD_DisplayScene "Lecithin","Collected", UltraDMD_Animation_ScrollOnRight, UltraDMD_deOn, UltraDMD_Animation_None
				End If
				If Light4.state = LightStateOff Then
					LecithinCount = 0
				End If
				LecithinCount = LecithinCount + 1
				If LecithinCount > 5 Then LecithinCount = 5
				FlashForMS Light4, 10000, 0, LightStateOff
				If LecithinCount = 1 Then
					AddScore 50000
				Else
					if Not UltraDMD.isRendering then DMD_DisplayScene "Lecithin Combo", 75000*(LecithinCount-1), UltraDMD_Animation_ScrollOnLeft, UltraDMD_deOn, UltraDMD_Animation_None
					Addscore 75000*(LecithinCount-1)
				End If
			Elseif CurExperiment = "ElixirOfLife" Then
				If Lecithin = 0 Then
					Lecithin = 1
					DMD_CancelRendering
					DMD_DisplayScene "Lecithin","Collected", UltraDMD_Animation_ScrollOnRight, UltraDMD_deOn, UltraDMD_Animation_None
					ScoreElixirHit
				End If
			Elseif 	CurExperiment = "BrainAndIQ" And Light4.state <> LightStateOff Then
				AddIQLecithin
			Else
				AddScore 50000
			end if
		End If
	End If
	UpdateElixir
End Sub

'******************************************************
' 					MERCURY RAMP/BATH
'******************************************************

Dim MercuryCount

Sub Trigger3_Hit()
	If Countdown.enabled = true and CurExperiment = "BrainAndIQ" And Lightmercury.state <> LightStateOff Then
		If activeball.velx < 0 or activeball.vely < 0  Then
			AddIQMercury
		End if
	End If
End Sub

Dim CloneImage, CloneBalls: CloneBalls = 0

dim splashcount

Sub splash_timer()
	splashcount = splashcount+1
	Select Case (splashcount)
		Case 1: splash1.visible=false:splash2.visible=true
		Case 2:	splash2.visible=false:splash3.visible=true
		Case 3:	splash3.visible=false:splash2.visible=true
		Case 4: splash2.visible=false:splash1.visible=true
		Case 5: splash1.visible = false
				me.enabled = False
	End Select
End Sub

Sub MercuryBath_Hit()

	PlaySoundAtVol "tabd30", MercuryBath, 1
	splashcount = 0
	splash1.visible = true
	splash.enabled = true

	if (vpGameInPlay = TRUE) And (vpTilted = FALSE) then
		'FlashForMs Lightquicksilver, 250, 0, LightStateOff
		FlashForMs LightClonLight, 100, 0, LightStateOff  '' Messes up the next line so created a 2nd copy

		CloneImage = ActiveBall.image
		MercuryBath.destroyball


		If LightClon.state <> LightStateOff Then
			If BallsOnPlayfield < 5 Then
				CloneBalls = CloneBalls + 2
				BallsOnPlayfield=BallsOnPlayfield+1
			Else
				CloneBalls = CloneBalls + 1
			End If
			ClonTimer.Interval = 500
			ClonTimer.Enabled = 1
			if Not UltraDMD.isRendering then DMD_DisplayScene "","Clone Ball", UltraDMD_Animation_ScrollOnLeft, UltraDMD_deOn, UltraDMD_Animation_None
		Else
			CreateSaveBall
			If LightMercury.state = LightStateOff Then
				MercuryCount = 0
			End If

			If Countdown.enabled = false then
				FlashForMS LightMercury, 10000, 0, LightStateOff
				MercuryCount = MercuryCount + 1
				If MercuryCount > 5 Then MercuryCount = 5
			Else
				MercuryCount = 1
			End If
			If MercuryCount = 1 Then
				AddScore 25000
			Else
				if Not UltraDMD.isRendering then DMD_DisplayScene "Mercury Combo", 50000*(MercuryCount-1), UltraDMD_Animation_ScrollOnLeft, UltraDMD_deOn, UltraDMD_Animation_None
				Addscore 50000*(MercuryCount-1)
			End If
		End If

	Else
		MercuryBath.kick 170, 7
		DOF 122, DOFPulse
		DOF 118, DOFPulse
		DOF 234, DOFPulse  'DOF MX - Strobe
	End If
End Sub

Sub ClonTimer_Timer()
	Dim KickAngle

	Dim Ball
	Set Ball = MercuryBath1.CreateSizedballWithMass(Ballsize/2,Ballmass)
	Ball.image= CloneImage
	'MercuryBath1.CreateBall.image = CloneImage

	If (CloneBalls Mod 2) = 0 then
		KickAngle = 90 + (4*rnd)
		MercuryBath1.kick KickAngle, 55
	Else
		KickAngle = 35 + (2*rnd)
		MercuryBath1.kick KickAngle, 20
	End If

	DOF 122, DOFPulse
	DOF 118, DOFPulse
	DOF 234, DOFPulse  'DOF MX - Strobe

	CloneBalls = CloneBalls - 1
	If CloneBalls = 0 Then
		ClonTimer.enabled = False
	End If
End Sub

'******************************************************
' 						IGOR SINK
'******************************************************

Sub IgorSink_Hit()
	PlaySoundAtVol "Scoopenter", IgorSink, 1

	IgorSink.TimerInterval=1000:IgorSink.TimerEnabled=True
	FlashForMS Bulb1, 500, 100, LightStateOff
	If Countdown.enabled=True Then
		FlashForMs Lightigor, 1000,100, LightStateOff
		ModeTime = ModeTime + 10
		DisplayModeTime
	Elseif CurExperiment = "MonsterToLife" Then
		FlashForMs Lightigor, 1000,100, LightStateOff
	Else
		If BallsOnPlayfield > 1 Then
			if Not UltraDMD.isRendering Then DMD_DisplayScene "No New Experiments","During Multiball", UltraDMD_Animation_ZoomIn, UltraDMD_deOn, UltraDMD_Animation_None
		Else
			StartExperiment
		End If
	End If
End Sub

Sub IgorSink_Timer()
	IgorSink.TimerEnabled = FALSE
	IgorSink.kick 205, 12
	DOF 124, DOFPulse
	DOF 118, DOFPulse
	DOF 234, DOFPulse  'DOF MX - Strobe
End Sub

'******************************************************
' 					BAT TONGUE KICKOUT
'******************************************************

Sub kicker_Hit()
	PlaySoundAtVol "fx2_kicker_enter_left", Kicker, VolKick
	If (vpGameInPlay=TRUE) And (vpTilted=FALSE) Then
		If Countdown.enabled = false Then
			FlashForMs Light9, 1000,150, LightStateOff
			If battongue = 0 then
				PlaySound "Tabd69"
				battongue = 1
				if Not UltraDMD.isRendering then DMD_DisplayScene "Bat Tongue","Collected", UltraDMD_Animation_ScrollOnUp, UltraDMD_deOn, UltraDMD_Animation_None
			Else
				PlaySound "Tabd23"
			End If
			AddScore 75000
		Elseif CurExperiment = "ElixirOfLife" Then
			If battongue = 0 Then
				battongue = 1
				DMD_CancelRendering
				DMD_DisplayScene "Bat Tongue","Collected", UltraDMD_Animation_ScrollOnUp, UltraDMD_deOn, UltraDMD_Animation_None
				ScoreElixirHit
			Else
				PlaySound "Tabd23"
			End If
		ElseIf CurExperiment = "RaidGraveyard" Then
			AddBodyPartBT
		Else
			PlaySound "Tabd23"
			AddScore 75000
		end if
	End If
	Kicker.TimerInterval = 1000
	Kicker.TimerEnabled = True
	UpdateElixir
End Sub

Sub kicker_Timer()
	kicker.TimerEnabled = FALSE
	kicker.kick 168, 10
	PlaySoundAtVol "fx_kicker", kicker, VolKick
	DOF 122, DOFPulse
	DOF 118, DOFPulse
	DOF 234, DOFPulse  'DOF MX - Strobe
End Sub

'******************************************************
' 					EINSTEINIUM KICKOUT
'******************************************************

Sub KickerEinsteinium_Hit()
	PlaySoundAtVol "fx2_kicker_enter_left", KickerEinsteinium, 1
	If (vpGameInPlay=TRUE) And (vpTilted=FALSE) Then
		If extraballight.State <> LightStateOff Then
			extraballight.state=LightStateOff
			AwardExtraBall
		End If
		If Countdown.enabled = false Then
			FlashForMs Light12, 1000,150, LightStateOff
			If einsteinium = 0 Then
				einsteinium = 1
				PlaySound "Tabd69"
				if Not UltraDMD.isRendering then DMD_DisplayScene "Einsteinium","Collected", UltraDMD_Animation_ScrollOnDown, UltraDMD_deOn, UltraDMD_Animation_None
			Else
				PlaySound "Tabd23"
			End If
			AddScore 40000
		Elseif CurExperiment = "ElixirOfLife" Then
			If einsteinium = 0 Then
				einsteinium = 1
				DMD_CancelRendering
				DMD_DisplayScene "Einsteinium","Collected", UltraDMD_Animation_ScrollOnDown, UltraDMD_deOn, UltraDMD_Animation_None
				ScoreElixirHit
			Else
				PlaySound "Tabd23"
			End If
		Else
			PlaySound "Tabd23"
			AddScore 40000
		End If
	end if
	KickerEinsteinium.TimerInterval = 1000
	KickerEinsteinium.TimerEnabled=True
	UpdateElixir
End Sub

Sub KickerEinsteinium_Timer()
	KickerEinsteinium.TimerEnabled = FALSE
	KickerEinsteinium.Kick 280, 10
	PlaySoundAtVol "fx_kicker", KickerEinsteinium, VolKick
	DOF 125, DOFPulse
	DOF 118, DOFPulse
	DOF 234, DOFPulse  'DOF MX - Strobe
End Sub

'******************************************************
' 					TOWER KICKOUT
'******************************************************

Sub Tower_Hit()
	PlaySoundAtVol "Scoopenter", Tower, 1
	ActivateLightning


	If CurExperiment = "MonsterToLife" and Lning1.state <> LightStateOff then ' and monster ready
		ScoreJackpot
		Tower.TimerInterval = 5000
		Tower.TimerEnabled = True
	Else
		Tower.TimerInterval = 1000
		Tower.TimerEnabled = True
		if Countdown.enabled = true and CurExperiment = "BrainAndIQ" then
			AddIQTower
		Else
			AddScore 50000
		end if
	end if
End Sub

Sub Tower_Timer()
	if IgorSink.TimerEnabled = False Then
		Tower.TimerEnabled = False
		FlashForMS Bulb1, 500, 100, LightStateOff

		Dim Ball
		Set Ball = IgorSink.CreateSizedballWithMass(Ballsize/2,Ballmass)
		Ball.image="Ball_Final"
		'IgorSink.CreateBall.image="Ball_Final"

		IgorSink.TimerEnabled = True
		Tower.DestroyBall
		PlaySound "Tabd38"
	end if
End Sub

Sub ActivateLightning()
	Lightning.State = LightStateBlinking
	GroundLightning.State = LightStateBlinking
	PlaySound "Tabd14" 'Lightning
	DOF 242, DOFPulse  'DOF MX - Electrical Effects
	Lightning.TimerInterval=500
	Lightning.TimerEnabled=True
End Sub

Sub Lightning_Timer()
	Lightning.TimerEnabled=False
	Lightning.State = LightStateOff
	GroundLightning.State = LightStateOff
End Sub

'******************************************************
' 					FLASK KICKOUT
'******************************************************
Dim Flaskactive, SkipAward, battongue, yttrium, einsteinium, lecithin
Flaskactive=0:SkipAward=0:battongue=0:yttrium=0:einsteinium=0:Lecithin=0

Sub VUK_Hit()
	PlaySoundAtVol "fx2_kicker_enter_left", VUK, 1
	vuk.timerinterval=1000
	vuk.timerenabled=True
	If Countdown.enabled = false then FlashForMs vuklight, 1000,150, LightStateOff
End Sub

sub vuk_timer()
	Dim numElements
	numElements = battongue + yttrium + einsteinium + lecithin
	vuk.timerenabled = 0
	if numElements = 0 or (Countdown.enabled=true and CurExperiment <> "ElixirOfLife") or (numElements <> 4 and Countdown.enabled = true and CurExperiment = "ElixirOfLife") then
		vuk.kick 185, 10
		PlaySoundAtVol "fx_kicker", vuk, 1
		DOF 126, DOFPulse
		DOF 118, DOFPulse
		DOF 234, DOFPulse  'DOF MX - Strobe
		PlaySound "Tabd06"
		AddScore 40000
	Elseif numElements = 4 And Countdown.enabled = true And CurExperiment = "ElixirOfLife"  then
		ElixirOfLifeComplete = 1
		EndExperiment
		SkipAward = 1
		MixElements
	Else
		Flaskactive=1
		MixTimer.Interval=10
		MixTimer.UserValue=0
		MixTimer.enabled = true
	end if
end sub

Sub MixTimer_Timer()
	DMD_CancelRendering
	If mixtimer.uservalue = 0 then
		DMD_DisplayScene DisplayAward,"<<< Mix Elements", UltraDMD_Animation_ScrollOnUp, 4000, UltraDMD_Animation_None
		If MixTimer.interval <> 2000 then MixTimer.Interval=2000
		mixtimer.uservalue = 1
	Else
		DMD_DisplayScene (battongue + einsteinium + lecithin + yttrium)*50000*TableMultiplier & " Points","Empty Flask >>>", UltraDMD_Animation_ScrollOnUp, 4000, UltraDMD_Animation_None
		mixtimer.uservalue = 0
	End If
End Sub

Sub flask_Hit()
	PlaySound"Tabd36" 'mix elements

	If SkipAward = 1 Then
		playsound "tabd58"
		AddScore 500000
		SkipAward = 0
	Else
		ElixirAward()
		PotionBonus = PotionBonus + 1
	End If
	flask.destroyball
	ResetElixir()
	CreateSaveBall()
End Sub

Sub MixElements()
	Flaskactive=0
	MixTimer.enabled=false
	PlaySound "Tabd07" 'as you wish master (mix elements)
	VUK.DestroyBall

	Dim Ball
	Set Ball = flask1.CreateSizedballWithMass(Ballsize/2,Ballmass)
	Ball.image="Ball_Final"
	'flask1.CreateBall.image="Ball_Final"

	flask1.kick 226, 5
	DOF 126, DOFPulse
	DOF 118, DOFPulse
	DOF 234, DOFPulse  'DOF MX - Strobe
End Sub

Sub ClearFlask()
	Flaskactive=0
	MixTimer.enabled=false
	if battongue then AddScore 50000
	if yttrium then AddScore 50000
	if einsteinium then AddScore 50000
	if lecithin then AddScore 50000
	ResetElixir()
	PlaySound"Tabd03"
	vuk.kick 185, 10
	PlaySoundAtVol "fx_kicker", vuk, 1
	DOF 126, DOFPulse
	DOF 118, DOFPulse
	DOF 234, DOFPulse  'DOF MX - Strobe
End Sub

Sub ResetElixir()
	battongue=0:yttrium=0:einsteinium=0:Lecithin=0
	UpdateElixir
End Sub

'******************************************************
' 					ELIXIR AWARDS
'******************************************************
Dim Einsteiniumcount,TableMultiplier,AllElements, ExtraBalls
Einsteiniumcount=0:TableMultiplier=1:AllElements=0:ExtraBalls=0

Sub ElixirAward()
	DMD_CancelRendering
	If battongue=0 and Yttrium=0 and einsteinium=1 and lecithin=0 Then
		Select Case (einsteiniumcount)
				Case 0: AddScore 10000
				Case 1: AddScore 25000
				Case 2: AddScore 50000
				Case 3: AddScore 100000
				Case 4: AddScore 250000
				Case 5: AddScore 500000
		End Select
		If einsteiniumcount <> 5 Then einsteiniumcount = einsteiniumcount + 1
	Elseif battongue=0 and Yttrium=0 and einsteinium=0 and lecithin=1 Then
		DMD_DisplayScene "Add to Jackpot","500000 Points",UltraDMD_Animation_ScrollOnLeft, UltraDMD_deOn, UltraDMD_Animation_None
		AddJackpot(500000)
	Elseif battongue=0 and Yttrium=1 and einsteinium=0 and lecithin=0 Then
		OpenTeleporter
	Elseif battongue=1 and Yttrium=0 and einsteinium=0 and lecithin=0 Then
		Bumperhits = (INT(Bumperhits/10) + 1)*10
		DMD_DisplayScene "Advance Bumpers", (BumperHits/10) + 1 &"X",UltraDMD_Animation_ScrollOnLeft, UltraDMD_deOn, UltraDMD_Animation_None
	Elseif battongue=0 and Yttrium=0 and einsteinium=1 and lecithin=1 Then
		BallSaver.Enabled=1
		BallSaver.interval=25000
		BallSaverLight.state=LightStateOn
		DMD_DisplayScene "Ball Saver","Activated",UltraDMD_Animation_ScrollOnLeft, UltraDMD_deOn, UltraDMD_Animation_None
	Elseif battongue=0 and Yttrium=1 and einsteinium=1 and lecithin=0 Then
		LightClon.state = LightStateOn
		DMD_DisplayScene "Ball Cloning","Activated",UltraDMD_Animation_ScrollOnLeft, UltraDMD_deOn, UltraDMD_Animation_None
	Elseif battongue=1 and Yttrium=0 and einsteinium=1 and lecithin=0 Then
		Bumperhits = 100
		DMD_DisplayScene "Maximum Bumpers", (BumperHits/10) + 1 &"X",UltraDMD_Animation_ScrollOnLeft, UltraDMD_deOn, UltraDMD_Animation_None
	Elseif battongue=0 and Yttrium=1 and einsteinium=0 and lecithin=1 Then
		If TableMultiplier <> 5 Then TableMultiplier = TableMultiplier + 1
		SetTableMultiplier
	Elseif battongue=1 and Yttrium=0 and einsteinium=0 and lecithin=1 Then
		EnablePlasmaField
	Elseif battongue=1 and Yttrium=1 and einsteinium=0 and lecithin=0 Then
		DMD_DisplayScene "Big Score!",500000*TableMultiplier & " Points",UltraDMD_Animation_ScrollOnLeft, UltraDMD_deOn, UltraDMD_Animation_None
		AddScore(500000)
	Elseif battongue=0 and Yttrium=1 and einsteinium=1 and lecithin=1 Then
		DMD_DisplayScene "Extra Ball Lit",500000*TableMultiplier & " Points",UltraDMD_Animation_ScrollOnLeft, UltraDMD_deOn, UltraDMD_Animation_None
		LightExtraBall
		AddScore(500000)
		AddJackpot(500000)
	Elseif battongue=1 and Yttrium=0 and einsteinium=1 and lecithin=1 Then
		DMD_DisplayScene "Big Score!",1000000*TableMultiplier & " Points",UltraDMD_Animation_ScrollOnLeft, UltraDMD_deOn, UltraDMD_Animation_None
		AddScore(1000000)
	Elseif battongue=1 and Yttrium=1 and einsteinium=1 and lecithin=0 Then
		DMD_DisplayScene "","Multiball!!!",UltraDMD_Animation_ScrollOnLeft, UltraDMD_deOn*2, UltraDMD_Animation_None
		StartMultiball
	Elseif battongue=1 and Yttrium=1 and einsteinium=0 and lecithin=1 Then
		TableMultiplier = 5
		SetTableMultiplier
	Elseif battongue=1 and Yttrium=1 and einsteinium=1 and lecithin=1 Then
		If Countdown.enabled = True And CurExperiment = "ElixirOfLife" Then
			ElixirOfLifeComplete = 1
		Else
			DMD_DisplayScene "Extra Ball!!!",1000000*TableMultiplier & " Points",UltraDMD_Animation_ScrollOnLeft, UltraDMD_deOn, UltraDMD_Animation_None
			AwardExtraBall
			AddScore(1000000)
			If AllElements = 1 Then
				AddJackpot(1000000)
			Else
				AllElements = 1
			End If
		End If
	End if
End Sub

Dim DisplayAward, DisplayElements

Sub UpdateElixir()
	If battongue=0 and Yttrium=0 and einsteinium=1 and lecithin=0 Then
		Select Case (einsteiniumcount)
				Case 0: DisplayAward = 10000*TableMultiplier & " Points"
				Case 1: DisplayAward = 25000*TableMultiplier & " Points"
				Case 2: DisplayAward = 50000*TableMultiplier & " Points"
				Case 3: DisplayAward = 100000*TableMultiplier & " Points"
				Case 4: DisplayAward = 250000*TableMultiplier & " Points"
				Case 5: DisplayAward = 500000*TableMultiplier & " Points"
		End Select
	Elseif battongue=0 and Yttrium=0 and einsteinium=0 and lecithin=1 Then
		DisplayAward = "Jackpot+ 500000"
	Elseif battongue=0 and Yttrium=1 and einsteinium=0 and lecithin=0 Then
		DisplayAward = "Open Teleporter"
	Elseif battongue=1 and Yttrium=0 and einsteinium=0 and lecithin=0 Then
		DisplayAward = "Increment Bumpers"
	Elseif battongue=0 and Yttrium=0 and einsteinium=1 and lecithin=1 Then
		DisplayAward = "Ball Saver On"
	Elseif battongue=0 and Yttrium=1 and einsteinium=1 and lecithin=0 Then
		DisplayAward = "Enable Cloning"
	Elseif battongue=1 and Yttrium=0 and einsteinium=1 and lecithin=0 Then
		DisplayAward = "Maximum Bumpers"
	Elseif battongue=0 and Yttrium=1 and einsteinium=0 and lecithin=1 Then
		DisplayAward = "Adv Table Multiplier"
	Elseif battongue=1 and Yttrium=0 and einsteinium=0 and lecithin=1 Then
		DisplayAward = "Enable Plasma Field"
	Elseif battongue=1 and Yttrium=1 and einsteinium=0 and lecithin=0 Then
		DisplayAward = "Score " & 500000*TableMultiplier
	Elseif battongue=0 and Yttrium=1 and einsteinium=1 and lecithin=1 Then
		DisplayAward = "Lite Ex Ball + Pts"
	Elseif battongue=1 and Yttrium=0 and einsteinium=1 and lecithin=1 Then
		DisplayAward = 1000000*TableMultiplier & " Points"
	Elseif battongue=1 and Yttrium=1 and einsteinium=1 and lecithin=0 Then
		DisplayAward = "Multiball"
	Elseif battongue=1 and Yttrium=1 and einsteinium=0 and lecithin=1 Then
		DisplayAward = "5x Table Multiplier"
	Elseif battongue=1 and Yttrium=1 and einsteinium=1 and lecithin=1 Then

		DisplayAward = "Ex Ball + " & 1*TableMultiplier &"M Pts"

	Else
		DisplayAward = ""
	End If

	DisplayElements = ""

	If battongue = 1 Then DisplayElements = DisplayElements & "Bat Tongue" &  chr(13)
	If einsteinium = 1 Then DisplayElements = DisplayElements & "Einsteinium" &  chr(13)
	If yttrium = 1 Then DisplayElements = DisplayElements & "Yttrium" &  chr(13)
	If lecithin = 1 Then DisplayElements = DisplayElements & "Lecithin" &  chr(13)

End Sub

Dim ElixirCount

Sub ScoreElixirHit()
	ElixirCount = ElixirCount + 1
	AddScore 50000 + (50000 * ElixirCount)
	PlaySound "Tabd69"
	Playsound "Tabd57"
End Sub

Sub StartMultiball()

	WallSinkTimer.interval=750
	WallSinkTimer.enabled=True
End Sub

Sub WallSinkTimer_Timer()
	If BallsOnPlayfield < 5 Then
		FlashForMs Flasher99, 100, 100, LightStateOff

		Dim Ball
		Set Ball = WallSink.CreateSizedballWithMass(Ballsize/2,Ballmass)
		Ball.image="Ball_Final"
		'WallSink.Createball.image = "Ball_Final"

		WallSink.kick 80, 10
		BallsOnPlayfield = BallsOnPlayfield + 1
		PlaySoundAtVol SoundFXDOF("scoopexit",127,DOFPulse,DOFcontactors), WallSink, 1
		DOF 118, DOFPulse
		DOF 234, DOFPulse  'DOF MX - Strobe
	Else
		WallSinkTimer.Enabled = False
	End If
End Sub

Sub SetTableMultiplier()
	LightTable2x.State = LightStateOff
	LightTable3x.State = LightStateOff
	LightTable4x.State = LightStateOff
	LightTable5x.State = LightStateOff

	Select Case (TableMultiplier)
		'Case 1:
		Case 2:LightTable2x.state = 1
		Case 3:LightTable2x.state = 1:LightTable3x.state = 1
		Case 4:LightTable2x.state = 1:LightTable3x.state = 1:LightTable4x.state = 1
		Case 5:LightTable2x.state = 1:LightTable3x.state = 1:LightTable4x.state = 1:LightTable5x.state = 1
	End Select

	Playsound "Tabd23"

	If TableMultiplier > 1 Then
		DMD_DisplayScene "Adv Table Multiplier", TableMultiplier & "X",UltraDMD_Animation_ScrollOnLeft, UltraDMD_deOn, UltraDMD_Animation_None
	End If
End Sub

Sub LightExtraBall()
	extraballight.state = LightStateOn

End Sub

Sub AwardExtraBall()
	if Not UltraDMD.isrendering then DMD_DisplayScene "","Extra Ball!!!",UltraDMD_Animation_ZoomIn, UltraDMD_deOn, UltraDMD_Animation_None
	ExtraBalls = ExtraBalls + 1
	PlaySound "Tabd15" 'This could actually work
	DOF 251, DOFPulse   'DOF MX - Extra Ball
End Sub


'******************************************************
' 						EXPERIMENTS
'******************************************************
Dim CurExperiment, ShowExpUpdate

Sub StartExperiment()
	If CountDown.Enabled = False Then
		ResetLights
		AddScore(150000)
		DMD_CancelRendering
		If (A.state=LightStateBlinking) then
			FightAngryMobMode()
		ElseIf (B.state=LightStateBlinking) then
			RaidTheGraveyardMode()
		ElseIf (C.State=LightStateBlinking) then
			BrainAndIQMode()
		ElseIf (D.State=LightStateBlinking) then
			ElixirOfLifeMode()
		Elseif E.state=LightStateBlinking Then
			MonsterToLifeMode
		end if

	Else
		ModeTime = ModeTime + 10
		DisplayModeTime
	End If
End Sub

Countdown.interval = 1000

Sub ResetLights()
	LightMercury.BlinkInterval=100:Lightmercury.state=LightStateOff
	vuklight.BlinkInterval=100:vuklight.state=LightStateOff
	Light4.BlinkInterval=100:Light4.state=LightStateOff
	Light5.BlinkInterval=100:Light5.state=LightStateOff
	Light9.BlinkInterval=100:Light9.state=LightStateOff
	Light12.BlinkInterval=100:Light12.state=LightStateOff
	LightIgor.state=LightStateOff
End Sub

Sub Countdown_Timer()
	ModeTime = ModeTime - 1
	If ModeTime < 0 Then ModeTime = 0
	DisplayModeTime
	If ModeTime = 0 then
		ModeTime = 0
		EndExperiment
		Countdown.Enabled=False
	End if
End Sub

Sub DisplayModeTime()
	Dim localExp
	If CurExperiment = "AngryMob" Then localExp = "Angry Mob"
	If CurExperiment = "RaidGraveyard" Then localExp = "Raid Graveyard"
	If CurExperiment = "BrainAndIQ" Then localExp = "Find a Brain"
	If CurExperiment = "ElixirOfLife" Then localExp = "Elixir of Life"

	if Not UltraDMD.isRendering then DMD_DisplaySceneExWithId "timer",localExp, ModeTime, UltraDMD_Animation_ZoomIn, ModeTime*1000, UltraDMD_Animation_None:ShowExpUpdate = 1
	DMD_ModifySceneEx "timer", localExp, ModeTime, ModeTime*1000
End Sub

Sub EndExperiment()
	Countdown.enabled = False
	towertext.visible = False
	mercurytext.visible = False
	LightIgor.state=LightStateBlinking
	DMD_CancelRendering
	If CurExperiment = "AngryMob" Then
		If AngryMobComplete Then
			A.state=LightStateOn:A1.state=LightStateOn
			DMD_DisplayScene "Experiment Completed","Mob Defeated",UltraDMD_Animation_ZoomIn, UltraDMD_deOn, UltraDMD_Animation_None
			ExperimentBonus = ExperimentBonus + 1
		Else
			A.state=LightStateOff:A1.state=LightStateOff
			ExperimentFailed
		end if
		DropRedTargets
		ResetMonsterTargets
		ResetPlasmaTargets
	Elseif CurExperiment = "RaidGraveyard" Then
		If RaidGraveyardComplete Then
			B.state=LightStateOn:B1.state=LightStateOn
			DMD_DisplayScene "Experiment Completed","Body Collected",UltraDMD_Animation_ZoomIn, UltraDMD_deOn, UltraDMD_Animation_None
			ExperimentBonus = ExperimentBonus + 1
		Else
			B.state=LightStateOff:B1.state=LightStateOff
			ExperimentFailed
		End If
		Light5.state = LightStateOff
		Light9.state = LightStateOff
	Elseif CurExperiment = "BrainAndIQ" Then
		If BrainAndIQComplete Then
			C.state=LightStateOn:C1.state=LightStateOn
			DMD_DisplayScene "Experiment Completed","Brain Collected",UltraDMD_Animation_ZoomIn, UltraDMD_deOn, UltraDMD_Animation_None
			ExperimentBonus = ExperimentBonus + 1
		Else
			C.state=LightStateOff:C1.state=LightStateOff
			ExperimentFailed
		End If
		Light4.state=LightStateOff
		Light5.state=LightStateOff
		Lightmercury.state=LightStateOff
		Lning1.state=LightStateOff
		Lning2.state=LightStateOff
		Helper.collidable = False
	Elseif CurExperiment = "ElixirOfLife" Then
		If ElixirOfLifeComplete Then
			D.state=LightStateOn:D1.state = LightStateOn
			DMD_DisplayScene "Experiment Completed","Elixir Charged",UltraDMD_Animation_ZoomIn, UltraDMD_deOn, UltraDMD_Animation_None
			ExperimentBonus = ExperimentBonus + 1
		Else
			D.state=LightStateOff:D1.state = LightStateOff
			ExperimentFailed
		End If
		Light4.state=LightStateOff
		Light5.state=LightStateOff
		Light9.state=LightStateOff
		Light12.state=LightStateOff
		vuklight.state=LightStateOff
	End If

	myPlayMusicForMode(1)
	RandomExperiment
End Sub

Sub ExperimentFailed()
	DMD_DisplayScene "Experiment","Failed",UltraDMD_Animation_ZoomIn, UltraDMD_deOn, UltraDMD_Animation_None
	Lning1.state = LightStateOff
	Lning2.state = LightStateOff
	CurExperiment = ""
	Helper.collidable = False
End Sub

Sub RandomExperiment()


	If A.state = LightStateBlinking then A.state = LightStateOff:A1.state = LightStateOff
	If B.state = LightStateBlinking then B.state = LightStateOff:B1.state = LightStateOff
	If C.state = LightStateBlinking then C.state = LightStateOff:C1.state = LightStateOff
	If D.state = LightStateBlinking then D.state = LightStateOff:D1.state = LightStateOff
	If RaidGraveyardComplete = 1 and BrainAndIQComplete = 1 and ElixirOfLifeComplete = 1 Then
		E.state = LightStateBlinking:E1.state = LightStateBlinking
	Else
		select case (int(4*rnd)+1)
			Case 1: If AngryMobComplete Then RandomExperiment: Else A.state = LightStateBlinking:A1.state = LightStateBlinking
			Case 2: If RaidGraveYardComplete Then RandomExperiment: Else B.state = LightStateBlinking:B1.state = LightStateBlinking
			Case 3: If BrainAndIQComplete Then RandomExperiment: Else C.state = LightStateBlinking:C1.state = LightStateBlinking
			Case 4: If ElixirOfLifeComplete Then RandomExperiment: Else D.state = LightStateBlinking:D1.state = LightStateBlinking
		End Select
	End If
	LightIgor.state = LightStateBlinking
End Sub

Sub ResetExperiments()
	A.state = LightStateOff:A1.state = LightStateOff:AngryMobComplete = 0
	B.state = LightStateOff:B1.state = LightStateOff:RaidGraveyardComplete = 0:ResetMonster
	C.state = LightStateOff:C1.state = LightStateOff:BrainAndIQComplete = 0:IQ = 0
	D.state = LightStateOff:D1.state = LightStateOff:ElixirOfLifeComplete = 0
	E.state = LightStateOff:E1.state = LightStateOff
	CurExperiment = ""
	towertext.visible=false
	Lning1.state = LightStateOff
	Lning2.state = LightStateOff
	RandomExperiment
End Sub

'******************************************************
' 					FIGHT THE ANGRY MOB
'******************************************************

Dim RedTargetsHit,AngryMobComplete
AngryMobComplete = 0

Sub FightAngryMobMode()
	plasmatext.visible = false
	targetstext.imageA = "a_mobtargets"
	CurExperiment = "AngryMob"
	ShowExpUpdate = 0
	ModeTime = 40
	Countdown.enabled=True
	PlaySound "Tabd45"

	DMD_DisplayScene "Fight the Angry","Mob",UltraDMD_Animation_ScrollOnLeft, UltraDMD_deOn, UltraDMD_Animation_None
	DMD_DisplayScene "Hit Five","Red Targets",UltraDMD_Animation_ScrollOnLeft, UltraDMD_deOn, UltraDMD_Animation_None

	A.state=LightStateOn:A1.state = LightStateOn

	RedTargetsHit = 0

	ResetRedTargets

	myPlayMusicForMode(3)
End Sub

Sub ScoreRedTarget()
	Dim MobDefeated:MobDefeated = "Fight Angry Mob"
	AddScore(50000 + (RedTargetsHit * 10000))
	Select Case (RedTargetsHit)
		Case 0,5: Playsound "Tabd46"
		Case 1,6: Playsound "Tabd47"
		Case 2,7: Playsound "Tabd48"
		Case 3,8: Playsound "Tabd49":
		Case 4: Playsound "Tabd50"
	End Select
	RedTargetsHit = RedTargetsHit + 1
	If RedTargetsHit = 5 Then AngryMobComplete = 1
	If RedTargetsHit > 4 Then MobDefeated = "Mob Defeated"
	If ShowExpUpdate = 1 Then
		DMD_CancelRendering
		DMD_DisplayScene MobDefeated,"Hits " & RedTargetsHit,UltraDMD_Animation_ZoomIn, UltraDMD_deOn, UltraDMD_Animation_None
	End If
	If RedTargetsHit = 9 Then EndExperiment
End Sub

Sub ResetRedTargets()
	playsoundAtVol SoundFXDOF("droptargetreset",119,DOFPulse,DOFContactors), TargetN, VolTarg
	TargetM.isDropped = 0:TargetM.image="red":LightM.state = LightStateBlinking:switch1.rotx=180
	TargetO.isDropped = 0:TargetO.image="red":LightO.state = LightStateBlinking:switch2.rotx=180
	TargetN.isDropped = 0:TargetN.image="red":LightN.state = LightStateBlinking:switch3.rotx=180
	TargetS.isDropped = 0:TargetS.image="red":Light_S.state = LightStateBlinking:switch4.rotx=180
	TargetT.isDropped = 0:TargetT.image="red":LightT.state = LightStateBlinking:switch5.rotx=180
	TargetE.isDropped = 0:TargetE.image="red":LightE.state = LightStateBlinking:switch6.rotx=180
	TargetR.IsDropped = 0:TargetR.image="red":LightR.state = LightStateBlinking:switch7.rotx=180
	plasma.IsDropped = 0:plasma.image="red":Light25.State = LightStateBlinking
	field.IsDropped = 0:field.image="red":Light26.State = LightStateBlinking
End Sub

Sub DropRedTargets()
	TargetM.image="m"
	TargetO.image="o"
	TargetN.image="n"
	TargetS.image="s"
	TargetT.image="t"
	TargetE.image="e"
	TargetR.image="r"
	plasma.image="plasma"
	field.image="plasma"
End Sub


'******************************************************
' 					RAID THE GRAVEYARD
'******************************************************
Dim RaidGraveyardComplete:RaidGraveyardComplete=0

Sub RaidTheGraveyardMode()
	CurExperiment = "RaidGraveyard"
	ModeTime = 35
	Countdown.enabled=True
	PlaySound "Tabd42"

	DMD_DisplayScene "Raid The","Graveyard",UltraDMD_Animation_ScrollOnLeft, UltraDMD_deOn, UltraDMD_Animation_None
	DMD_DisplayScene "Shoot Flashing","Lights",UltraDMD_Animation_ScrollOnLeft, UltraDMD_deOn, UltraDMD_Animation_None


	B.state=LightStateOn:B1.state = LightStateOn

	if leg1.visible = false then Light5.state = LightStateBlinking
	if leg2.visible = false then Light9.state = LightStateBlinking

	myPlayMusicForMode(2)
End Sub

Sub AddBodyPartBT()
	If leg2.visible = false Then
		DMD_CancelRendering
		AddScore 125000
		PlaySound "Tabd44"
		PartsBonus = PartsBonus + 1
	Else
		AddScore 75000
	End If

	If head.visible = false Then
		Head.visible = True
		DMD_DisplayScene "Head Collected",125000*TableMultiplier & " Points",UltraDMD_Animation_ZoomIn, UltraDMD_deOn, UltraDMD_Animation_None
	elseif leftarm.visible = false Then
		leftarm.visible = True
		DMD_DisplayScene "Left Arm Collected",125000*TableMultiplier & " Points",UltraDMD_Animation_ZoomIn, UltraDMD_deOn, UltraDMD_Animation_None
	elseif leg2.visible = false Then
		leg2.visible = True
		DMD_DisplayScene "Left Leg Collected",125000*TableMultiplier & " Points",UltraDMD_Animation_ZoomIn, UltraDMD_deOn, UltraDMD_Animation_None
		Light9.state = LightStateOff
		CheckMonsterParts
	End If
End Sub

Sub AddBodyPartYT()
	If leg1.visible = false Then
		DMD_CancelRendering
		AddScore 75000
		PlaySound "Tabd43"
		PartsBonus = PartsBonus + 1
	Else
		AddScore 40000
	End If

	If torso.visible = false Then
		torso.visible = True
		DMD_DisplayScene "Torso Collected",75000*TableMultiplier & " Points",UltraDMD_Animation_ZoomIn, UltraDMD_deOn, UltraDMD_Animation_None
	elseif rightarm.visible = false Then
		rightarm.visible = True
		DMD_DisplayScene "Right Arm Collected",75000*TableMultiplier & " Points",UltraDMD_Animation_ZoomIn, UltraDMD_deOn, UltraDMD_Animation_None
	elseif leg1.visible = false Then
		leg1.visible = True
		DMD_DisplayScene "Right Leg Collected",75000*TableMultiplier & " Points",UltraDMD_Animation_ZoomIn, UltraDMD_deOn, UltraDMD_Animation_None
		Light5.state = LightStateOff
		CheckMonsterParts
	End If
End Sub

Sub CheckMonsterParts()
	If leg1.visible = true and leg2.visible = true Then

		RaidGraveyardComplete = 1
		EndExperiment
	End If
End Sub

'******************************************************
' 					BRAIN AND IQ
'******************************************************

Dim BrainAndIQComplete,IQ:BrainAndIQComplete=0:IQ=0

Sub BrainAndIQMode()
	CurExperiment = "BrainAndIQ"
	ModeTime = 35
	IQ = 0
	Countdown.enabled=True
	PlaySound "Tabd51"

	DMD_DisplayScene "Find a Magnificent","Brain",UltraDMD_Animation_ScrollOnLeft, UltraDMD_deOn, UltraDMD_Animation_None
	DMD_DisplayScene "Shoot Flashing","Lights",UltraDMD_Animation_ScrollOnLeft, UltraDMD_deOn, UltraDMD_Animation_None

	C.state=LightStateOn:C1.state = LightStateOn

	Light4.state=LightStateBlinking
	Light5.state=LightStateBlinking
	Lightmercury.state=LightStateBlinking
	Lning1.state=LightStateBlinking
	Lning2.state=LightStateBlinking
	Helper.collidable = True

	myPlayMusicForMode(5)
End Sub

Sub AddIQMercury()
	brain.visible = true
	BrainAndIQComplete = 1
	PlaySound "tabd52"

	If IQ < 20 then DMD_CancelRendering

	If IQ < 5 Then
		IQ = 5
		DMD_DisplayScene "Chicken Brain","IQ = 5",UltraDMD_Animation_ScrollOnDown, UltraDMD_deOn, UltraDMD_Animation_None
	elseif IQ < 10 then
		IQ = 10
		DMD_DisplayScene "Cow Brain","IQ = 10",UltraDMD_Animation_ScrollOnDown, UltraDMD_deOn, UltraDMD_Animation_None
	elseif IQ < 15 then
		IQ = 15
		DMD_DisplayScene "Sheep Brain","IQ = 15",UltraDMD_Animation_ScrollOnDown, UltraDMD_deOn, UltraDMD_Animation_None
	elseif IQ < 20 then
		IQ = 20
		DMD_DisplayScene "Pig Brain","IQ = 20",UltraDMD_Animation_ScrollOnDown, UltraDMD_deOn, UltraDMD_Animation_None
		Lightmercury.state = LightStateOff
	End If

	if IQ <= 20 then
		AddScore 2000*IQ
	Else
		AddScore 2000*20
	End If
End Sub

Sub AddIQLecithin()
	brain.visible = true
	BrainAndIQComplete = 1
	PlaySound "tabd53"
	Lightmercury.state = LightStateOff

	If IQ < 65 then DMD_CancelRendering

	If IQ < 35 Then
		IQ = 35
		DMD_DisplayScene "Damaged Brain","IQ = 35",UltraDMD_Animation_ScrollOnDown, UltraDMD_deOn, UltraDMD_Animation_None
	elseif IQ < 45 then
		IQ = 45
		DMD_DisplayScene "Slightly Used Brain","IQ = 45",UltraDMD_Animation_ScrollOnDown, UltraDMD_deOn, UltraDMD_Animation_None
	elseif IQ < 55 then
		IQ = 55
		DMD_DisplayScene "Bargain Brain","IQ = 55",UltraDMD_Animation_ScrollOnDown, UltraDMD_deOn, UltraDMD_Animation_None
	elseif IQ < 65 then
		IQ = 65
		DMD_DisplayScene "Designer Brain","IQ = 65",UltraDMD_Animation_ScrollOnDown, UltraDMD_deOn, UltraDMD_Animation_None
		Light4.state = LightStateOff
	End If

	if IQ <= 65 then
		AddScore 2000*IQ
	Else
		AddScore 2000*65
	End If
End Sub

Sub AddIQYttrium()
	brain.visible = true
	BrainAndIQComplete = 1
	PlaySound "tabd54"
	Lightmercury.state = LightStateOff
	Light4.state = LightStateOff

	If IQ < 105 then DMD_CancelRendering

	if IQ < 75 then
		IQ = 75
		DMD_DisplayScene "Programmer Brain","IQ = 75",UltraDMD_Animation_ScrollOnDown, UltraDMD_deOn, UltraDMD_Animation_None
	elseif IQ < 85 then
		IQ = 85
		DMD_DisplayScene "Abnormal Brain","IQ = 85",UltraDMD_Animation_ScrollOnDown, UltraDMD_deOn, UltraDMD_Animation_None
	elseif IQ < 95 then
		IQ = 95
		DMD_DisplayScene "John Doe Brain","IQ = 95",UltraDMD_Animation_ScrollOnDown, UltraDMD_deOn, UltraDMD_Animation_None
	elseif IQ < 105 then
		IQ = 105
		DMD_DisplayScene "Jane Doe Brain","IQ = 105",UltraDMD_Animation_ScrollOnDown, UltraDMD_deOn, UltraDMD_Animation_None
		Light5.state = LightStateOff
	End If

	if IQ <= 105 then
		AddScore 2000*IQ
	Else
		AddScore 2000*105
	End If
End Sub

Sub AddIQTower()
	brain.visible = true
	BrainAndIQComplete = 1
	PlaySound "tabd55"
	Lightmercury.state = LightStateOff
	Light4.state = LightStateOff
	Light5.state = LightStateOff

	If IQ < 200 then DMD_CancelRendering

	if (IQ < 140) then
		IQ = 140
		DMD_DisplayScene "Monkey Brain","IQ = 140",UltraDMD_Animation_ScrollOnDown, UltraDMD_deOn, UltraDMD_Animation_None
	elseif (IQ < 160) then
		IQ = 160
		DMD_DisplayScene "Scientist Brain","IQ = 160",UltraDMD_Animation_ScrollOnDown, UltraDMD_deOn, UltraDMD_Animation_None
	elseif (IQ < 180) then
		IQ = 180
		DMD_DisplayScene "Artificial Brain","IQ = 180",UltraDMD_Animation_ScrollOnDown, UltraDMD_deOn, UltraDMD_Animation_None
	elseif (IQ < 200) then
		IQ = 200
		DMD_DisplayScene "Einstein Brain","IQ = 200",UltraDMD_Animation_ScrollOnDown, UltraDMD_deOn, UltraDMD_Animation_None
		EndExperiment
	End If

	AddScore 2000*IQ
End Sub

'******************************************************
' 					ELIXIR OF LIFE
'******************************************************

Dim ElixirOfLifeComplete,ElixirHits:ElixirOfLifeComplete=0

Sub ElixirOfLifeMode()
	CurExperiment = "ElixirOfLife"
	ModeTime = 35
	ElixirHits=0
	Countdown.enabled=True
	If battongue = 0 or lecithin = 0 or yttrium = 0 or einsteinium = 0 Then PlaySound "Tabd56"

	DMD_DisplayScene "Charge the Elixir","Of Life",UltraDMD_Animation_ScrollOnLeft, UltraDMD_deOn, UltraDMD_Animation_None
	DMD_DisplayScene "Collect Elements","Then Shoot Flask",UltraDMD_Animation_ScrollOnLeft, UltraDMD_deOn, UltraDMD_Animation_None

	D.state=LightStateOn:D1.state = LightStateOn

	CheckElements

	myPlayMusicForMode(4)
End Sub

Sub ScoreElixirHit()
	ElixirHits = ElixirHits + 1
	AddScore 50000 + 50000*ElixirHits
	PlaySound "tabd57"
	CheckElements
End Sub

Sub CheckElements()
	If battongue = 1 then light9.state = LightStateOff Else light9.state = LightStateBlinking
	If lecithin = 1 then light4.state = LightStateOff Else light4.state = LightStateBlinking
	If yttrium = 1 then light5.state = LightStateOff Else light5.state = LightStateBlinking
	If einsteinium = 1 then light12.state = LightStateOff Else light12.state = LightStateBlinking
	If battongue = 1 and lecithin = 1 and yttrium = 1 and einsteinium = 1 Then
		Playsound "tabd66"
		vuklight.state = LightStateBlinking
	End If
End Sub

'******************************************************
' 				BRING THE MONSTER TO LIFE
'******************************************************

Sub MonsterToLifeMode()
	CurExperiment = "MonsterToLife"
	PlaySound "Tabd59"

	DMD_DisplayScene "Bring the Monster","To Life",UltraDMD_Animation_ScrollOnLeft, UltraDMD_deOn*2, UltraDMD_Animation_None
	DMD_DisplayScene "Shoot Monster","Targets",UltraDMD_Animation_ScrollOnLeft, UltraDMD_deOn*2, UltraDMD_Animation_None

	E.state = LightStateOn:E1.state = LightStateOn
	Lightigor.state = LightStateOff

	ResetMonsterTargets
	targetstext.imageA = "a_monstertargets"
	myPlayMusicForMode(6)
End Sub

Sub ResetMonster()
	head.visible = False
	Rightarm.visible = False
	Leftarm.visible = False
	Torso.visible = False
	leg1.visible = False
	leg2.visible = False
	brain.visible = False
	diverter1p.rotx=70
	diverter2p.rotx=70
	diverter3p.rotx=70
End Sub

Sub ScoreJackpot()
	ExperimentBonus = ExperimentBonus + 1
	Helper.collidable=false
	Jackpot.Play SeqMiddleInHorizOn, 90, 10
	DMD_CancelRendering
	DMD_DisplayScene "J a c k p o t", nvjackpot, UltraDMD_Animation_ZoomIn, 6000, UltraDMD_Animation_FadeOut
	PlaySound "Tabd62"
	DOF 241, DOFPulse  'DOF MX - Jackpot
	DOF 242, DOFPulse  'DOF MX - Electrical Effects
	AddScore(nvJackpot)
	nvJackpot=1000000
	JackpotTimer.Enabled=True
End Sub

Sub JackpotTimer_Timer ()
	myPlayMusicForMode(6)
	ResetExperiments
	ResetMonsterTargets
	IQ=0
	JackpotTimer.Enabled = FALSE
End Sub

'******************************************************
' 						NEW GAME
'******************************************************

Sub NewGame()
	DmdIntro.enabled = False
	DMD_CancelRendering
	DMD_DisplayScene "","MAD Scientist", UltraDMD_Animation_None, UltraDMD_deOn, UltraDMD_Animation_None
	DMD_DisplayScene "Bring the Monster","To Life", UltraDMD_Animation_None, UltraDMD_deOn, UltraDMD_Animation_None
	DMD_DisplayScene "","Now Have Fun", UltraDMD_Animation_None, UltraDMD_deOn, UltraDMD_Animation_None
	Score=0
	nvjackpot=1000000
	vpTilted= False
	vpGameInPlay = True

	ResetExperiments
	ResetElixir
	CloseTeleporter

	PlaySound "Tabd63"

	Jackpot.StopPlay()
	Lightwheel.StopPlay():Lightwheel1.StopPlay()
	Random.StopPlay()

	Bulb1.state = LightStateOff
	BulbBlue.state = LightStateOff
	Lightspot2A.state = LightStateOn
	Lightspot1.state = LightStateOn

	AllElements = 0
	ExtraBalls = 0
	Einsteiniumcount = 0
	BallImageCount = 0
	BallInPlay = 1
	BonusHeld = 0

	FirstBallDelayTimer.interval=3500:FirstBallDelayTimer.enabled=True
End Sub

'******************************************************
' 						NEW BALL
'******************************************************

Sub NewBall()
	TableMultiplier = 1
	SetTableMultiplier
	ResetReCapture
	ResetPlasmaField
	ResetMonsterTargets
	EnableSkillShot
	RandomExperiment

	myPlayMusicForMode(1)

	BumperHits = 0
	TargetResets = 0

	PotionBonus = 0
	ExperimentBonus = 0
	PartsBonus = 0

	Light23.state = LightStateOff
	Light24.state = LightStateOff
	Light4.state = LightStateOff
	Light5.state = LightStateOff
	Light9.state = LightStateOff
	LightMercury.state = LightStateOff
	LightClon.state = LightStateOff
	Light12.state = LightStateOff
	extraballight.state = LightStateOff
	vuklight.state = LightStateOff

	ScoreTimer.enabled = True
End Sub

Sub FirstBallDelayTimer_Timer()
	FirstBallDelayTimer.Enabled = FALSE
	NewBall
	BallreleaseTimer.Enabled = True
End Sub

'******************************************************
' 						PLUNGER
'******************************************************

Sub TriggerPlungerLane_Hit()
	bBallInPlungerLane = TRUE
	DOF 223, DOFOn  'DOF MX - Ball Ready to Shoot
	playsoundAtVol "fx_sensor", TriggerPlungerLane, 1
End Sub

Sub TriggerPlungerLane_unhit()
	bBallInPlungerLane = FALSE
	DOF 223, DOFOff  'DOF MX - Ball Ready to Shoot - Off
End Sub

Sub PlungerState(state)
	Dim Blinking, Solid
	If state=1 or state=2 or state=3 then
		Blinking = 2
		Solid = 1
		Plunger.Pullback
		Plunger.timerinterval=300
		Plunger.timerenabled=1
		DOF 236, DOFOn  'DOF MX - Plunger Lane - Green flask
		DOF 223, DOFOff 'DOF MX - Turn off - Ready to Shoot
		if state = 1 Then
			if bBallInPlungerLane = TRUE Then PlaySound "Tabd64"
		else
			PlaySound "Tabd70"
			autoTimer.Enabled=True
			CreateSaveBallTimer.Enabled = True
		end if
	else
		Solid = 0
		Blinking = 0
		Plunger.Fire
		PlaySound SoundFXDOF("Tabd27",114,DOFPulse,DOFContactors)
		DOF 208, DOFPulse  'DOF MX - Ball Release
		DOF 236, DOFOff  'DOF MX - Plunger Lane - Green flask - OFF
		DOF 223, DOFOff 'DOF MX - Turn off - Ready to Shoot
		cork1 = 1
		corktimer.enabled = true
		Plunger.timerenabled=0
	End If

	Lightbottle.state = Solid
	Bubble.State = Blinking
	Bubble1.State = Blinking
	Bubble2.State = Blinking
	Bubble3.State = Blinking
	Bubble4.State = Blinking
	Bubble5.State = Blinking
	Bubble6.State = Blinking
	Bubble7.State = Blinking
	Bubble8.State = Blinking
	Firef.State = Blinking
	flamecount = 0
	flame.enabled = true
End Sub

dim flamecount

Sub flame_timer()
	flamecount = flamecount+1
	if flamecount = 5 then flamecount=1
	flame1.visible=true

	Select Case (flamecount)
		Case 1: Flame1.imageA="flame1"
		Case 2:	Flame1.imageA="flame2"
		Case 3:	Flame1.imageA="flame3"
		Case 4: Flame1.imageA="flame2"
	End Select
End Sub

Sub AutoTimer_Timer()
	PlungerState(0)
	AutoTimer.Enabled = FALSE
End Sub

Sub Plunger_Timer()
	Playsound "Tabd35", 0.5
End Sub

Dim cork1
Sub CorkTimer_timer()
	Select Case cork1
		Case 1:cork.transy = 5:cork1 = 2
		Case 2:cork.transy = 15:cork1 = 3
		Case 3:cork.transy = 50:cork1 = 4
		Case 4:cork.transy = 50:cork1 = 5
		Case 5:cork.transy = 15:cork1 = 6
		Case 6:cork.transy = 5:cork1 = 7
		Case 7:
			cork.transy = 0
			corktimer.enabled = false
			Flame.enabled = false
			flame1.visible = False
	End Select

End Sub

Sub BallreleaseTimer_Timer()
	Dim Ball
	Set Ball = Ballrelease.CreateSizedballWithMass(Ballsize/2,Ballmass)
	Ball.image="Ball_Final"
	BallsOnPlayfield = BallsOnPlayfield + 1
	Ballrelease.kick 180, 10
	DOF 113, DOFPulse
	BallreleaseTimer.Enabled = FALSE
End Sub


'******************************************************
' 						SAVE BALL
'******************************************************

Sub CreateSaveBall()
	SaverBalls = SaverBalls + 1
	SaverTimer.enabled = true
End sub

Sub CreateSaveBallTimer_Timer()
	Dim Ball
	Set Ball = Ballrelease.CreateSizedballWithMass(Ballsize/2,Ballmass)
	Ball.image="Ball_Final"
	Ballrelease.kick 180, 10
	DOF 113, DOFPulse
	CreateSaveBallTimer.Enabled =False
End Sub

Sub SaverTimer_Timer()
	If SaverBalls = 0 Then
		SaverTimer.enabled = False
		SaverTimer.Interval = 500
	Else
		PlungerState(2)
		SaverBalls = SaverBalls - 1
		SaverTimer.Interval = 1500
	End If
End Sub

'******************************************************
' 						DRAIN
'******************************************************

Dim SaverBalls:SaverBalls = 0

Sub Drain_Hit()
	DOF 115, DOFPulse
	Drain.DestroyBall
	PlaySoundAtVol "Drain", Drain, 1
	DOF 210, DOFPulse 'DOF MX - Ball Drained
	if (Light7blue.state = LightStateOff) and (Ballsaverlight.state = LightStateOff) then
		BallsOnPlayfield = BallsOnPlayfield - 1
		if (BallsOnPlayfield=0) then
			EndOfBall()
		end if
	end if
	If (vpGameInPlay = TRUE) And (vpTilted = FALSE) Then
		if (Light7Blue.state <> LightStateOff) then
			ResetReCapture

			Dim Ball
			Set Ball = Drain2.CreateSizedballWithMass(Ballsize/2,Ballmass)
			Ball.image="Ball_Final"
			'Drain2.CreateBall.image="Ball_Final"

			Drain2.kick 0,40
			PlaySound "TABD20"
		elseif (Ballsaverlight.state <> LightStateOff) Then
			PlaySoundAtVol "Drain", Drain, 1
			If Not UltraDMD.isRendering then DMD_DisplayScene "","Ball Saved", UltraDMD_Animation_ZoomIn, UltraDMD_deOn, UltraDMD_Animation_None
			CreateSaveBall()
		End If
	Elseif vpTilted = True then
		BallsOnPlayfield = BallsOnPlayfield - 1
		if (BallsOnPlayfield=0) then
			EndOfBall()
			Ballsaver.enabled = False
			ballsaverlight.state = LightStateOff
		end if
	End If
End Sub

Sub BallSaver_Timer()
	FlashForMs ballsaverlight, 5000, 100, LightStateOff
	Me.enabled = false
End Sub

'******************************************************
' 					END OF BALL
'******************************************************

Sub EndOfBall()
	If Countdown.enabled = True Then
		EndExperiment
	Elseif e.state = LightStateOn Then
		towertext.visible = false
		ExperimentFailed
		E.state = LightStateOff:E1.state = LightStateOff
	End If
	If INT(rnd*2) = 1 then
		PlaySound "Tabd67"
	Else
		Playsound "Tabd65"
	End If
	If vpTilted = False Then
		CalcBonus
	Else
		vpTilted = False
		playmusic song
		EndOfBall2
	End If

End Sub

Sub EndOfBall2()
	If ExtraBalls > 0 Then
		ExtraBalls = ExtraBalls - 1
		If ExtraBalls > 0 then
			DMD_CancelRendering
			DMD_DisplayScene "Extra Balls Left "& ExtraBalls, "Shoot Again",UltraDMD_Animation_ZoomIn, UltraDMD_deOn, UltraDMD_Animation_None
		Else
			DMD_DisplayScene "", "Shoot Again",UltraDMD_Animation_ZoomIn, UltraDMD_deOn, UltraDMD_Animation_None
		End If
        EndOfBallComplete
	Else
		BallInPlay = BallInPlay + 1
  		If BallInPlay > nvBallsPerGame Then
			BallInPlay = 0
			CheckHS
			EndOfGameTimer.enabled = True
		Else
			EndOfBallComplete
		End if
	End If
End Sub

Sub EndOfBallComplete()
	FirstBallDelayTimer.interval=1000:FirstBallDelayTimer.enabled=True
End Sub

'******************************************************
' 						BONUS
'******************************************************
Dim PotionBonus,ExperimentBonus,PartsBonus,BonusHeld

Sub CalcBonus()
	Dim Potions, Experiments, Parts, TotalBonus
	Potions = PotionBonus * 50000
	Experiments = ExperimentBonus * 250000
	Parts = PartsBonus * 100000
	TotalBonus = Potions + Experiments + Parts + BonusHeld

	DMD_DisplayScene "Potions Bonus",PotionBonus & " X 50000",UltraDMD_Animation_ZoomIn, UltraDMD_deOn/2, UltraDMD_Animation_None
	DMD_DisplayScene "Experiment Bonus",ExperimentBonus & " X 250000",UltraDMD_Animation_ZoomIn, UltraDMD_deOn/2, UltraDMD_Animation_None
	DMD_DisplayScene "Body Part Bonus",PartsBonus & " X 100000",UltraDMD_Animation_ZoomIn, UltraDMD_deOn/2, UltraDMD_Animation_None
	DMD_DisplayScene "Bonus Held",BonusHeld,UltraDMD_Animation_ZoomIn, UltraDMD_deOn/2, UltraDMD_Animation_None
	DMD_DisplayScene "Total Bonus",TotalBonus,UltraDMD_Animation_ZoomIn, UltraDMD_deOn/2, UltraDMD_Animation_None

	AddScore TotalBonus
	ScoreTimer.enabled = False

	BonusHeld = Potions + Experiments + Parts
	BonusTimer.enabled = true
End Sub

Sub BonusTimer_Timer()
	If LeftFlipperDown = 1 and RightFlipperDown = 1 Then DMD_CancelRendering
	if Not UltraDMD.isrendering Then
		EndOfBall2
		BonusTimer.enabled = False
	end if
End Sub

'******************************************************
' 					END OF GAME
'******************************************************

Sub EndOfGame(dmd)
	vpGameInPlay = False
	If dmd = 0 then
		DMDIntroState = 4
	Else
		DMDIntroState = 0
	End If
	DMDIntro.interval = 4000
	DMDIntro.Enabled = True

	myPlayMusicForMode(7)

	Jackpot.Play SeqMiddleInHorizOn, 90, 3
	Lightwheel.Play SeqArcBottomLeftUpOn, 90,4:Lightwheel1.Play SeqArcBottomLeftUpOn, 90,4
	Lightwheel.Play SeqArcBottomrightUpOn, 90,4:Lightwheel1.Play SeqArcBottomrightUpOn, 90,4
	Lightwheel.Play SeqRandom, 90,4:Lightwheel1.Play SeqRandom, 90,4
End Sub

Sub EndOfGameTimer_Timer()
	if bEnteringAHighScore = False Then
		EndofGame 0
		me.enabled = False
	End If
End Sub

Sub GameoverTimer_Timer()
	select case ((int(10*rnd+1)))
		Case 0:playsound "Tabd05"
		Case 1:playsound "Tabd07"
		Case 2:playsound "Tabd15"
		Case 3:PlaySound "Tabd42"
		Case 4:PlaySound "Tabd43"
		Case 5:PlaySound "Tabd44"
		Case 6:PlaySound "Tabd14"
		Case 7:PlaySound "Tabd54"
		Case 8:PlaySound "Tabd45"
    Case 9:PlaySound "Tabd68"
	end select
	GameoverTimer.Enabled=False
End Sub


'******************************************************
' 						TILT
'******************************************************

Dim TiltSens

Sub CheckTilt
	If Tilttimer.Enabled = True and vpGameInPlay = TRUE and VPTilted = False Then
		TiltSens = TiltSens + 1
		if TiltSens = 3 Then
			vpTilted = True
			PlaySound "Tabd67"
			DMD_CancelRendering
			DMD_DisplayScene "","T I L T", UltraDMD_Animation_ScrollOnUp, UltraDMD_deOn, UltraDMD_Animation_None
			endmusic
		Else
			PlaySound "Tabd52"
			DOF 221, DOFPulse  'DOF MX - Nudge
			DMD_CancelRendering
			if TiltSens = 1 Then
				DMD_DisplayScene "","Careful ...", UltraDMD_Animation_ScrollOnUp, UltraDMD_deOn, UltraDMD_Animation_None
			ElseIf TiltSens = 2 Then
				DMD_DisplayScene "Do Not Shake","The Chemicals!", UltraDMD_Animation_ScrollOnUp, UltraDMD_deOn, UltraDMD_Animation_None
			Else
				DMD_DisplayScene "","Gently ...", UltraDMD_Animation_ScrollOnUp, UltraDMD_deOn, UltraDMD_Animation_None
			End If
			Tilttimer.enabled = False
			Tilttimer.enabled = True
		End If
	Else
		TiltSens = 0
		Tilttimer.Enabled = True
	End If
End Sub

Sub Tilttimer_Timer()
	Tilttimer.Enabled = False
End Sub

Sub CreditsTimer_Timer()
	nvCredits = nvCredits + 1
	DOF 116, DOFOn
	CreditsTimer.Enabled=False

	if (vpGameInPlay=False) And (vpTilted=False) Then
		DMD_CancelRendering
		DMD_DisplayScene "Credits " & nvCredits , "Press Start", UltraDMD_Animation_None, 2000, UltraDMD_Animation_None
	end if
End Sub



'******************************************************
' 					FLASK LIGHTS
'******************************************************

'stolen from gtxjoe, with credit

Dim light1pos(2):light1pos(0) = Flask22.x: light1pos(1) = Flask22.y
Dim light2pos(2):light2pos(0) = Flask2.x: light2pos(1) = Flask2.y
Dim light3pos(2):light3pos(0) = Flask3.x: light3pos(1) = Flask3.y
Dim light4pos(2):light4pos(0) = Flask4.x: light4pos(1) = Flask4.y
Dim light5pos(2):light5pos(0) = Flask5.x: light5pos(1) = Flask5.y
Dim light6pos(2):light6pos(0) = Flask6.x: light6pos(1) = Flask6.y
Dim light7pos(2):light7pos(0) = Flask7.x: light7pos(1) = Flask7.y
Dim light8pos(2):light8pos(0) = Flask8.x: light8pos(1) = Flask8.y
Dim light9pos(2):light9pos(0) = Flask9.x: light9pos(1) = Flask9.y
Dim light10pos(2):light10pos(0) = Flask10.x: light10pos(1) = Flask10.y
Dim light11pos(2):light11pos(0) = Flask11.x: light11pos(1) = Flask11.y
Dim light12pos(2):light12pos(0) = Flask12.x: light12pos(1) = Flask12.y
Dim light13pos(2):light13pos(0) = Flask13.x: light13pos(1) = Flask13.y
Dim light14pos(2):light14pos(0) = Flask14.x: light14pos(1) = Flask14.y
Dim light15pos(2):light15pos(0) = Flask15.x: light15pos(1) = Flask15.y
Dim light16pos(2):light16pos(0) = Flask16.x: light16pos(1) = Flask16.y
Dim light17pos(2):light17pos(0) = Flask17.x: light17pos(1) = Flask17.y
Dim light18pos(2):light18pos(0) = Flask18.x: light18pos(1) = Flask18.y
Dim light19pos(2):light19pos(0) = Flask19.x: light19pos(1) = Flask19.y
Dim light20pos(2):light20pos(0) = Flask20.x: light20pos(1) = Flask20.y

Const length1offset = 25
Const length2offset = 12
Const length3offset = 19
Const length4offset = 16
Const length5offset = 10
Const length6offset = 25
Const length7offset = 12
Const length8offset = 16
Const length9offset = 18
Const length10offset = 11
Const length11offset = 14
Const length12offset = 21
Const length13offset = 12
Const length14offset = 17
Const length15offset = 21
Const length16offset = 23
Const length17offset = 20
Const length18offset = 16
Const length19offset = 20
Const length20offset = 24

dim angle3, dir3, anglez3, dirz3, flasklevel, flasklightstate
angle3 = 0:dir3 = 0.5:dirz3 = 1:flasklevel=10


FlaskTimer.interval = 30

Sub FlaskTimer_timer()
	dim quartersfull
	quartersfull = battongue + yttrium + lecithin + einsteinium

	if quartersfull = 1 and flasklevel < 5 then flasklevel = flasklevel + 1
	if quartersfull = 2 and flasklevel < 10 then flasklevel = flasklevel + 1
	if quartersfull = 3 and flasklevel < 15 then flasklevel = flasklevel + 1
	if quartersfull = 4 and flasklevel < 20 then flasklevel = flasklevel + 1
	if quartersfull = 0 and flasklevel > 0 then flasklevel = flasklevel - 1

	if quartersfull = 1 then prim25.visible=true else prim25.visible=false
	if quartersfull = 2 then prim50.visible=true else prim50.visible=false
	if quartersfull = 3 then prim75.visible=true else prim75.visible=false
	if quartersfull = 4 then prim100.visible=true else prim100.visible=false

	If flasklevel = 0 then
		flasklightstate=0
		LightSeq1.StopPlay
		Flask21.State = 0
	Else
		if flasklightstate = 0 then
			LightSeq1.UpdateInterval = 5
			LightSeq1.Play SeqRandom, 340, 100
			Flask21.State = 1
			flasklightstate = 1
		end if
	End If

	angle3 = angle3 + 1 * dir3
	dir3 = -1 * dir3

	anglez3 = anglez3 + 1 * dirz3
	if anglez3 => 360 then
		anglez3 = 0
	end if

	calcXYZ Flask22, light1Pos, length1offset, angle3, angleZ3
	calcXYZ Flask2, light2Pos, length2offset, angle3, angleZ3
	calcXYZ Flask3, light3Pos, length3offset, angle3, angleZ3
	calcXYZ Flask4, light4Pos, length4offset, angle3, angleZ3
	calcXYZ Flask5, light5Pos, length5offset, angle3, angleZ3
	calcXYZ Flask6, light6Pos, length6offset, angle3, angleZ3
	calcXYZ Flask7, light7Pos, length7offset, angle3, angleZ3
	calcXYZ Flask8, light8Pos, length8offset, angle3, angleZ3
	calcXYZ Flask9, light9Pos, length9offset, angle3, angleZ3
	calcXYZ Flask10, light10Pos, length10offset, angle3, angleZ3
	calcXYZ Flask11, light11Pos, length11offset, angle3, angleZ3
	calcXYZ Flask12, light12Pos, length12offset, angle3, angleZ3
	calcXYZ Flask13, light13Pos, length13offset, angle3, angleZ3
	calcXYZ Flask14, light14Pos, length14offset, angle3, angleZ3
	calcXYZ Flask15, light15Pos, length15offset, angle3, angleZ3
	calcXYZ Flask16, light16Pos, length16offset, angle3, angleZ3
	calcXYZ Flask17, light17Pos, length17offset, angle3, angleZ3
	calcXYZ Flask18, light18Pos, length18offset, angle3, angleZ3
	calcXYZ Flask19, light19Pos, length19offset, angle3, angleZ3
	calcXYZ Flask20, light20Pos, length20offset, angle3, angleZ3
End sub


' *********************************************************
' Calculate X, Y and Z from latitude, longitude and radius
' *********************************************************
Sub CalcXYZ (ballname, ballpos, length, angXY, angZ)
	Dim lat, lon

	lat = angXY *  3.14159265359/180
	lon = angZ *  3.14159265359/180

	ballname.x = ballpos(0) + length * cos(lat) * cos(lon)
	ballname.y = ballpos(1) + (length * 2.5) * cos(lat) * sin(lon)
	ballname.bulbhaloheight = flasklevel + 60
	Flask21.bulbhaloheight = flasklevel	+ 60
End Sub

'************************************************************************
'***********HIGHSCORES***************************************************
'************************************************************************
Dim highScore(3)									'The highscore values
Dim initials(3)										'What has been entered on the initial screen
Dim HighScoreName(3)								'The player initials
Dim inChar:inChar=65								'Which character the player is entering
Dim cursorPos:cursorPos=40							'Cursor position of character entry (0-2) Hitting START on character 3 finishes entry

Dim HSStrings:HSStrings=Array("High Score 2","High Score 1","Grand Champion")
Dim HSNames:HSNames=Array("HighScore2Name","HighScore1Name","GrandChampionName")
Dim HSPlace:HSPlace=-1
Dim HSCase:HSCase=0
'***********LOAD/SAVE HIGHSCORES*********************************************
Sub LoadHighScores
	Dim x
    x = LoadValue(cGameName, "HighScore2")
    If(x <> "")then HighScore(0) = CDbl(x) Else HighScore(0) = 5000000 End If
    x = LoadValue(cGameName, "HighScore2Name")
    If(x <> "")then HighScoreName(0) = x Else HighScoreName(0) = "FRK" End If
    x = LoadValue(cGameName, "HighScore1")
    If(x <> "")then HighScore(1) = CDbl(x) Else HighScore(1) = 10000000 End If
    x = LoadValue(cGameName, "HighScore1Name")
    If(x <> "")then HighScoreName(1) = x Else HighScoreName(1) = "IGR" End If
    x = LoadValue(cGameName, "GrandChampion")
    If(x <> "")then HighScore(2) = CDbl(x) Else HighScore(2) = 15000000 End If
    x = LoadValue(cGameName, "GrandChampionName")
    If(x <> "")then HighScoreName(2) = x Else HighScoreName(2) = "MS" End If
    x = LoadValue(cGameName, "Credits")
    If(x <> "")then nvcredits = CInt(x) Else nvcredits = 0 End If
    x = LoadValue(cGameName, "Score")
    If(x <> "")then Score = x Else Score = 0 End If
    x = LoadValue(cGameName, "Jackpot")
    If(x <> "")then nvJackpot = x Else Score = 0 End If
    x = LoadValue(cGameName, "IQ")
    If(x <> "")then IQ = x Else Score = 0 End If
End Sub

Sub SaveHighScores
	'SaveValue cGameName, "HighScore2", HighScore(0)
	'SaveValue cGameName, "HighScore2Name", HighScoreName(0)
	'SaveValue cGameName, "HighScore1", HighScore(1)
	'SaveValue cGameName, "HighScore1Name", HighScoreName(1)
	'SaveValue cGameName, "GrandChampion", HighScore(2)
	'SaveValue cGameName, "GrandChampionName", HighScoreName(2)
	SaveValue cGameName, "Score", Score
	SaveValue cGameName, "Credits", nvcredits
	SaveValue cGameName, "Jackpot", nvjackpot
	SaveValue cGameName, "IQ", IQ
End Sub

'***********NAME ENTERING SCREEN *********************************************
Sub nameEntry
	dim tmp, tmp1, tmp2, tmp3

	If UltraDMD.IsRendering Then UltraDMD.CancelRendering

	If inChar = 91 Then		'character [
		tmp = "<"
	Else
		tmp = Chr(inChar)
	End If


	if cursorPos = 0 Then
		tmp1 = tmp
		tmp2 = ""
		tmp3 = ""
	elseif cursorPos = 1 Then
		tmp1 = Chr(Initials(0))
		tmp2 = tmp
		tmp3 = ""
		initials(1) = inChar
		initials(2) = ""
	elseif cursorPos = 2 Then
		tmp1 = Chr(Initials(0))
		tmp2 = Chr(Initials(1))
		tmp3 = tmp
	end If

	UltraDMD.DisplayScene00 "", "ENTER YOUR NAME", 15, tmp1 & tmp2 & tmp3, 15, UltraDMD_Animation_None, 900000, UltraDMD_Animation_None
End Sub

'***********CHECK FOR HIGHSCORES *********************************************
Sub CheckHS
	If score>=highScore(2) Then
		SaveValue cGameName, "HighScore2", HighScore(1)
		SaveValue cGameName, "HighScore2Name", HighScoreName(1)
		SaveValue cGameName, "HighScore1", HighScore(2)
		SaveValue cGameName, "HighScore1Name", HighScoreName(2)
		SaveValue cGameName, "GrandChampion", Score
		SaveValue cGameName, "GrandChampionName", " "
		highScore(2)=Score:HSPLACE=2
	ElseIf score>=highScore(1) Then
		SaveValue cGameName, "HighScore2", HighScore(1)
		SaveValue cGameName, "HighScore2Name", HighScoreName(1)
		SaveValue cGameName, "HighScore1", Score
		SaveValue cGameName, "HighScore1Name", " "
		highScore(1)=Score:HSPLACE=1
	ElseIf score>=highScore(0) Then
		SaveValue cGameName, "HighScore2", Score
		SaveValue cGameName, "HighScore2Name", " "
		highScore(0)=Score:HSPLACE=0
	Else
		HSPLACE=-1
	End If
	If HSPlace <> -1 Then
		bEnteringAHighScore = True
		cursorPos = 0
		nameEntry
	End If
End Sub

'***********EXIT HIGHSCORES *********************************************
Sub ExitHs_timer
	Me.Enabled=0
	bEnteringAHighScore=False
	HighScoreName(HSPlace)=Chr(initials(0)) & Chr(initials(1)) & Chr(initials(2))
	initials(0)=""
	initials(1)=""
	initials(2)=""
	inChar = 65
	SaveValue cGameName, HSNames(HSPlace),HighScoreName(HSPlace)
	UltraDMD.CancelRendering
	DMD_DisplayScene HSStrings(HSPlace),HighScoreName(HSPlace) & " " & formatnumber(highScore(HSPlace),0), UltraDMD_Animation_None, 4000, UltraDMD_Animation_None
End Sub


'******************************************************
' 						MUSIC
'******************************************************

Dim song

Sub myPlayMusicForMode(Mode)
	If (CurrentMusicTunePlaying <> Mode) Then
		Select Case (Mode)
			Case 0: EndMusic
			Case 1:	song = "bgout_TABD1.mp3":endmusic:playmusic song
			Case 2: song = "bgout_TABD2.mp3":endmusic:playmusic song
			Case 3: song = "bgout_TABD3.mp3":endmusic:playmusic song
			Case 4: song = "bgout_TABD4.mp3":endmusic:playmusic song
			Case 5:	song = "bgout_TABD5.mp3":endmusic:playmusic song
			Case 6: song = "bgout_TABD6.mp3":endmusic:playmusic song
			Case 7:	song = "bgout_TABD8.mp3":endmusic:playmusic song
		End Select
		CurrentMusicTunePlaying = Mode
	End If
End Sub

Sub Table1_MusicDone()
	PlayMusic Song
End Sub

'********************************************************************************************
' FlashForMs will blink light or a flasher for TotalPeriod(ms) at rate of BlinkPeriod(ms)
' When TotalPeriod done, light or flasher will be set to FinalState value where
' Final State values are:   0=Off, 1=On, 2=Return to previous State
'********************************************************************************************

Sub FlashForMs(MyLight, TotalPeriod, BlinkPeriod, FinalState) 'thanks gtxjoe for the first version
	If FinalState = 2 Then
		FinalState = MyLight.State 				'Keep the current light state
	End If

	if BlinkPeriod = 0 Then
		MyLight.Duration 1, TotalPeriod, FinalState
	Else
		MyLight.BlinkInterval = BlinkPeriod
		MyLight.Duration 2, TotalPeriod, FinalState
	End If
End Sub


'******************************************************
' 					ULTRADMD
'******************************************************

Sub LoadUltraDMD
    Set UltraDMD = CreateObject("UltraDMD.DMDObject")
    UltraDMD.Init

    Dim fso, curDir
    Set fso = CreateObject("Scripting.FileSystemObject")
    curDir = fso.GetAbsolutePathName(".")
    Set fso = nothing

    ' A Major version change indicates the version is no longer backward compatible
    If Not UltraDMD.GetMajorVersion = 1 Then
        MsgBox "Incompatible Version of UltraDMD found."
        Exit Sub
    End If


    'A Minor version change indicates new features that are all backward compatible
    If UltraDMD.GetMinorVersion < 1 Then
        MsgBox "Incompatible Version of UltraDMD found.  Please update to version 1.1 or newer."
        Exit Sub
    End If


    UltraDMD.SetProjectFolder curDir & "\madscientist.UltraDMD"
    UltraDMD.SetVideoStretchMode UltraDMD_VideoMode_Middle

End Sub

Sub DMD_DisplayScene(toptext,bottomtext,animateIn,pauseTime,animateOut)
	UltraDMD.DisplayScene00 "", toptext, 13, bottomtext, 15, animateIn, pauseTime, animateOut
	ScoreTimer.enabled = true
End Sub

Sub DMD_DisplaySceneExWithId(id,toptext,bottomtext,animateIn,pauseTime,animateOut)
	UltraDMD.DisplayScene00ExWithId id, 0, "", toptext, 13, 9, bottomtext, 15, 9, animateIn, pauseTime, animateOut
	ScoreTimer.enabled = true
End Sub

Sub DMD_ScrollingCredits(text,animateIn,pauseTime,animateOut)
	UltraDMD.ScrollingCredits "", text, 15, animateIn, pauseTime, animateOut
	ScoreTimer.enabled = true
End Sub

Sub DMD_ModifyScene(id,toptext,bottomtext)
	UltraDMD.ModifyScene00 id, toptext, bottomtext
End Sub

Sub DMD_ModifySceneEx(id,toptext,bottomtext,pauseTime)
	UltraDMD.ModifyScene00Ex id, toptext, bottomtext,pauseTime
End Sub

Sub DMD_CancelScene(id)
	UltraDMD.CancelRenderingWithId id
End Sub

Sub DMD_CancelRendering()
	UltraDMD.CancelRendering
End Sub

Dim DMDIntroState:DMDIntroState=0

Sub DMDIntro_Timer()
	DMDIntroState = DMDIntroState + 1
	Select Case (DMDIntroState)
		Case 1: DMD_DisplayScene "","MAD Scientist", UltraDMD_Animation_None, 4000, UltraDMD_Animation_None
		Case 2: DMD_DisplayScene "Bring the Monster","To Life", UltraDMD_Animation_None, 4000, UltraDMD_Animation_None
		Case 3: DMD_DisplayScene "","Now Have Fun", UltraDMD_Animation_None, 4000, UltraDMD_Animation_None
		Case 4: if bFreePlay = True Then
					DMD_DisplayScene "Press Start","Free Play", UltraDMD_Animation_None, 4000, UltraDMD_Animation_None
				elseif (nvcredits = 0) then
					DMD_DisplayScene "Insert Coin","Credits " & nvCredits, UltraDMD_Animation_None, 4000, UltraDMD_Animation_None
				elseif (nvcredits > 0) then
					DMD_DisplayScene "Press Start","Credits " & nvCredits, UltraDMD_Animation_None, 4000, UltraDMD_Animation_None
				end if
		Case 5: DMD_DisplayScene "","Game Over", UltraDMD_Animation_None, 4000, UltraDMD_Animation_None
		Case 6: DMD_DisplayScene "Grand Champion",HighScoreName(2) & " " & formatnumber(highScore(2),0), UltraDMD_Animation_None, 4000, UltraDMD_Animation_None
		Case 7: DMD_DisplayScene "High Score 1",HighScoreName(1) & " " & formatnumber(highScore(1),0), UltraDMD_Animation_None, 4000, UltraDMD_Animation_None
		Case 8: DMD_DisplayScene "High Score 2",HighScoreName(0) & " " & formatnumber(highScore(0),0), UltraDMD_Animation_None, 4000, UltraDMD_Animation_None
		Case 9: DisplayScore
		'Case 10:
		'Case 11:
		Case 12: DMDIntroState = 0
	End Select
End Sub

Sub ScoreTimer_Timer()
	If Countdown.enabled = false Then
		if NOT UltraDMD.isrendering then
			IF vpTilted = True Then
				DMD_DisplayScene "","T I L T", UltraDMD_Animation_ScrollOnUp, UltraDMD_deOn, UltraDMD_Animation_None
			Else
				DisplayScore
				ScoreTimer.enabled = false
			End If
		Else
			DisplayModeTime
		end if

	End If

End Sub

Function NumberFormat(num)
	NumberFormat = Replace(FormatNumber(num,0,False,False,True), ",", ".")
End Function

Dim OldIQ

Sub RollingTimer_Timer()
    LFLogo.Roty = LeftFlipper.CurrentAngle
    RFlogo.Roty = RightFlipper.CurrentAngle

	if oldIQ <> IQ Then
		dim brainsize: brainsize = 50
		oldIQ = IQ

		select case (IQ)
			Case 5: brainsize = 20
			Case 10: brainsize = 22
			Case 15: brainsize = 24
			Case 20: brainsize = 26
			Case 35: brainsize = 28
			Case 45: brainsize = 30
			Case 55: brainsize = 32
			Case 65: brainsize = 34
			Case 75: brainsize = 36
			Case 85: brainsize = 38
			Case 95: brainsize = 40
			Case 105: brainsize = 42
			Case 140: brainsize = 44
			Case 160: brainsize = 46
			Case 180: brainsize = 48
			Case 200: brainsize = 50
		end select


		brain.size_X = brainsize
		brain.size_Y = brainsize
		brain.size_Z = brainsize
	end if

	if vpGameInPlay = TRUE then
		If CountDown.Enabled = False Then
			If E.state=LightStateOn OR BallsOnPlayfield > 1 Then
				igortext.visible=false
			Elseif (A.state=LightStateBlinking) then
				setigortext "a_fightangrymob"
			ElseIf (B.state=LightStateBlinking) then
				setigortext "a_graveyard"
			ElseIf (C.State=LightStateBlinking) then
				setigortext "a_brain"
			ElseIf (D.State=LightStateBlinking) then
				setigortext "a_elixir"
			Elseif E.state=LightStateBlinking Then
				setigortext "a_monstertolife"
			end if
		Else
			setigortext "a_extendexperiment"
		end if

		If LightClon.state <> LightStateOff Then
			clonetext.visible=True
		Else
			clonetext.visible=false
		End If

		If extraballight.state <> LightStateOff Then
			extraballtext.visible=True
		Else
			extraballtext.visible=false
		End If

		if Countdown.Enabled = False Or CurExperiment = "ElixirOfLife" Then
			If battongue = 0 Then
				battonguetext.visible=True
				battonguetext.ImageA="a_battongue"
			Else
				battonguetext.visible=false
			End If

			If einsteinium = 0 Then
				einsteiniumtext.visible=True
				einsteiniumtext.ImageA="a_einsteinium"
			Else
				einsteiniumtext.visible=false
			End If

			If lecithin = 0 Then
				lecithintext.visible=True
				lecithintext.ImageA="a_lecithin"
			Else
				lecithintext.visible=false
			End If

			If yttrium = 0 Then
				yttriumtext.visible=True
				yttriumtext.ImageA="a_yttrium"
			Else
				yttriumtext.visible=false
			End If
		Elseif Countdown.Enabled = True AND CurExperiment = "RaidGraveyard" Then
			lecithintext.visible=false
			einsteiniumtext.visible=false
			If leg2.visible = true Then
				battonguetext.visible=false
			Elseif leftarm.visible = true Then
				battonguetext.visible=True
				battonguetext.imageA = "a_leftleg"
			Elseif head.visible = true Then
				battonguetext.visible=True
				battonguetext.imageA = "a_leftarm"
			Else
				battonguetext.visible=True
				battonguetext.imageA = "a_head"
			End If

			If leg1.visible = true Then
				yttriumtext.visible=false
			Elseif rightarm.visible = true Then
				yttriumtext.visible=True
				yttriumtext.imageA = "a_rightleg"
			Elseif torso.visible = true Then
				yttriumtext.visible=True
				yttriumtext.imageA = "a_rightarm"
			Else
				yttriumtext.visible=True
				yttriumtext.imageA = "a_torso"
			End If
		Elseif Countdown.Enabled = True AND CurExperiment = "BrainAndIQ" Then
			battonguetext.visible = False
			einsteiniumtext.visible = False
			If IQ <20 Then
				mercurytext.visible = true
			Else
				mercurytext.visible = false
			End If
			If IQ <65 Then
				lecithintext.visible = true
				lecithintext.imageA = "a_brainandlube"
			Else
				lecithintext.visible = false
			End If
			If IQ <105 Then
				yttriumtext.visible = true
				yttriumtext.imageA = "a_morgue"
			Else
				yttriumtext.visible = false
			End If
			If IQ <200 Then
				towertext.visible = true
				towertext.imageA = "a_research"
			Else
				towertext.visible = false
			End If
		Else
			battonguetext.visible=false
			lecithintext.visible=false
			yttriumtext.visible=false
			einsteiniumtext.visible=false
		End If

	Else
		igortext.visible=false
		teletext.visible=false
		plasmatext.visible=false
		targetstext.visible = false
		clonetext.visible=false
		battonguetext.visible=false
		lecithintext.visible=false
		yttriumtext.visible=false
		einsteiniumtext.visible=false
		extraballtext.visible=false
		towertext.visible=false
		mercurytext.visible=false
	end if
	RollingUpdate
End Sub

sub setigortext(imagename)
	igortext.visible=true
	igortext.imageA = imagename
end sub

'******************************
' Diverse Collection Hit Sounds
'******************************
Dim VolumeDial
VolumeDial=2

Sub aBumpers_Hit (idx): PlaySound SoundFX("fx_bumper", DOFContactors), 0, 1, pan(ActiveBall), AudioFade(ActiveBall): End Sub
Sub aRollovers_Hit(idx):PlaySound "fx_sensor", 0, Vol(ActiveBall)*VolumeDial, pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0, AudioFade(ActiveBall):End Sub
Sub aGates_Hit(idx):PlaySound "fx_Gate", 0, Vol(ActiveBall)*VolumeDial, pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0, AudioFade(ActiveBall):End Sub
Sub aMetals_Hit(idx):PlaySound "fx_MetalHit", 0, Vol(ActiveBall)*VolumeDial, pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0, AudioFade(ActiveBall):End Sub
Sub aRubber_Bands_Hit(idx):PlaySound "fx_rubber", 0, Vol(ActiveBall)*VolumeDial, pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0, AudioFade(ActiveBall):End Sub
Sub aRubbers_Hit(idx):PlaySound "fx_postrubber", 0, Vol(ActiveBall)*VolumeDial, pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0, AudioFade(ActiveBall):End Sub
Sub aRubber_Posts_Hit(idx):PlaySound "fx_postrubber", 0, Vol(ActiveBall)*VolumeDial, pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0, AudioFade(ActiveBall):End Sub
Sub aRubber_Pins_Hit(idx):PlaySound "fx_postrubber", 0, Vol(ActiveBall)*VolumeDial, pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0, AudioFade(ActiveBall):End Sub
Sub aYellowPins_Hit(idx):PlaySound "fx_postrubber", 0, Vol(ActiveBall)*VolumeDial, pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0, AudioFade(ActiveBall):End Sub
Sub aPlastics_Hit(idx):PlaySound "fx_PlasticHit", 0, Vol(ActiveBall)*VolumeDial, pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0, AudioFade(ActiveBall):End Sub
Sub aWoods_Hit(idx):PlaySound "fx_Woodhit", 0, Vol(ActiveBall)*VolumeDial, pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0, AudioFade(ActiveBall):End Sub

Sub DropTargetSound():PlaySoundAtVol SoundFXDOF("fx_droptarget",120,DOFPulse,DOFDropTargets),TargetN,1:End Sub


' Ramp Soundss

Sub rrail_Hit
	if activeball.vely < 0 then
		PlaySoundAtVol "fx_metalrolling", ActiveBall, 1
	end if
End Sub

Sub rrail_unhit
	if activeball.vely > 0 Then
		StopSound "fx_metalrolling"
	end if
End Sub

Sub lrail_Hit
	if activeball.vely < 0 then
		PlaySoundAtVol "fx_metalrolling", ActiveBall, 1
	end if
End Sub

Sub lrail_unhit
	if activeball.vely > 0 Then
		StopSound "fx_metalrolling"
	end if
End Sub


Sub RRHelp_Hit()
    StopSound "fx_metalrolling"
    PlaySound "fx_ballrampdrop", 0, 1, pan(ActiveBall)
End Sub

Sub LRHelp_Hit()
    StopSound "fx_metalrolling"
    PlaySound "fx_ballrampdrop", 0, 1, pan(ActiveBall)
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

Sub PlaySoundAtVol(sound, tableobj, Volume)
  PlaySound sound, 1, Volume, Pan(tableobj), 0,0,0, 1, AudioFade(tableobj)
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
  Vol = Csng(BallVel(ball) ^2 / VolDiv)
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

Sub RollingUpdate()
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
    PlaySound("fx_collide"), 0, Csng(velocity) ^2 / (VolDiv/VolCol), Pan(ball1), 0, Pitch(ball1), 0, 0, AudioFade(ball1)
End Sub


' Thalamus : Exit in a clean and proper way
Sub Table1_exit()
  Controller.Pause = False
  Controller.Stop
End Sub


'
'
'      _   _    ____ _  ______     ___  ____  _____ _   _
'     | | / \  / ___| |/ / ___|   / _ \|  _ \| ____| \ | |
'  _  | |/ _ \| |   | ' /\___ \  | | | | |_) |  _| |  \| |
' | |_| / ___ \ |___| . \ ___) | | |_| |  __/| |___| |\  |
'  \___/_/   \_\____|_|\_\____/   \___/|_|   |_____|_| \_|
'
'                Jacks Open by Gottlieb (1977)
'                       Original
'        **** Graphics/Sound/Build/Code by Pinuck ****
'
'  	NEW	-  Updated graphics/lighting/primitives/etc. for VPX by Sliderpoint
'   NEW -  New High Score Sticky and additional photo resources from GNance
'   NEW -  Score routine fixes by BorgDog

' Thalamus 2018-07-23
' Added/Updated "Positional Sound Playback Functions" and "Supporting Ball & Sound Functions"
' No special SSF tweaks yet.

Option Explicit
Randomize

Const cGameName = "jacksopen"

'****************************************
'  DECLARATIONS
'****************************************

Dim Score, Credits
Dim TargetValue, TargetMode
Dim Target1Down, Target2Down, Target3Down, Target4Down, Target5Down
Dim Target6Down, Target7Down, Target8Down, Target9Down
Dim TopLanesUnlit
Dim B1Lit, B2Lit, B3Lit
Dim bump1, bump2, bump3 ' for bumper animation
Dim T, obj, x
Dim dropsUp ' flag for resetting red drops
Dim BallRelPause, NextBallPause, TiltPause, GameOverPause, MatchPause
Dim ChimesPause, Chimes2Pause, DropResetPause, RedResetPause, SMFadePause, FiveKPause, FiveHPause
Dim PlungerIm
Dim SpecialLit
Dim LROlit : Dim RROlit : Dim LOLlit : Dim ROLlit : Dim LILlit : Dim RILlit  ' bottom rollovers and outlanes
Dim BallsRemaining, GameStarted, Tilt, Tilted, Match, CurrentBall
Dim HighScore, LastScore, HighScoreName, PlayerTLA, HSA1, HSA2, HSA3, HSDefault  'HSA=high score alpha (letter)
Dim Free100, Free130, Free170
Dim Score1M, Score100k, Score10k, ScoreK, Score100, Score10, Score1, Scorex	'Define 5 different score values for each reel to use
Dim HSEnterMode, hsLetterFlash, hsEnteredDigits(3), hsCurrentDigit, hsCurrentLetter
Dim Roll, CurrBall, TSCount
Dim DesktopMode: DesktopMode = JacksOpen.ShowDT

Const DefaultHighScore = 22420
Const DefaultHSA1 = 25
Const DefaultHSA2 = 25
Const DefaultHSA3 = 26
Const BallsPerGame = 5
Const SMFadeResetValue = 18
Const hsFlashDelay = 4
Const tiltSettle = 80
Const tiltAllow = 3 'how many nudges to tilt

On Error Resume Next
ExecuteGlobal GetTextFile("controller.vbs")
If Err Then MsgBox "Can't open controller.vbs"
On Error Goto 0

Dim BlinkingLightState(5)

'Const B2STableName="Jacks Open"
'Dim B2sOn
'B2SOn = 1 ' Set to 0 to disable B2S backglass, 1 to enable
'Dim Controller
'if B2SOn then
'	On Error Resume Next
'	ExecuteGlobal GetTextFile("controller.vbs")
'	If Err Then
'		MsgBox "Can't open controller.vbs"
'		B2SOn = 0
'	End If
'	On Error Goto 0
'	If B2SOn then
'		Set Controller = CreateObject("B2S.Server")
'		Controller.B2SName = B2STableName
'		Controller.Run()
'		If Err Then
'			MsgBox "Can't Load B2S.Server."
'			B2SOn = 0
'		End If
'	End if
'end if

'****************************************
'  TABLE INIT
'****************************************
 Sub JacksOpen_Init
	LoadEM
	LoadHighScore
	GameStarted = 0
	AllLampsOff
	If B2SOn then
		Controller.B2SSetMatch 34, 0 'match lights off
		Controller.B2SSetData 51, 0 'turn off 100k lights
		Controller.B2SSetCredits credits
	end if
	EMReelMatch.setvalue 0
	EMReel100k.setvalue 0
	LastScoreReels
	B1Lit=0:B2Lit=0:B3Lit=0
	If Credits <1 then Credits = 0 'credits=0
	If credits > 0 Then DOF 142, DOFOn
	EMReelCredits.setvalue(Credits)
	SaveHighScore
	EMReelGO.setvalue 1
	roll = 0
	Set CurrBall = Drain.CreateBall
	Dim ChgLed, iw
	If B2SOn Then
		For each iw in BGstuff: iw.Visible = False:Next
	Else
		For each iw in BGstuff: iw.Visible = True:Next
	End If

End Sub

'****************************************
'  TABLE EXIT
'****************************************
Sub JacksOpen_Exit
	If B2SOn Then Controller.Stop
End Sub

'****************************************
'  KEYSTROKES
'****************************************

Sub JacksOpen_KeyDown(ByVal keycode)
	If Gamestarted AND(NOT Tilted) Then
		If keycode = LeftFlipperKey Then
			PlaySound SoundFXDOF("PNK_MH_Flip_L_up",101,DOFOn,DOFContactors)
			LeftFlipper.RotateToEnd
			PlaySound "Buzz1",-1,1,.1
		end if
		If keycode = RightFlipperKey Then
			PlaySound SoundFXDOF("PNK_MH_Flip_R_up",102,DOFOn,DOFContactors)
			RightFlipper.RotateToEnd
			PlaySound "Buzz",-1,1,.1
		end if
	End If
	If keycode = LeftTiltKey Then
		Nudge 90, 2
		Bump
	End If

	If keycode = RightTiltKey Then
		Nudge 270, 2
		Bump
	End If

	If keycode = CenterTiltKey Then
		Nudge 0, 2
		Bump
	End If

	If keycode = MechanicalTilt Then Tilted = 1:mechchecktilt:End If

	If keycode = PlungerKey Then Plunger.Pullback
	If keycode = StartGameKey AND GameStarted = 0 And Credits > 0 And Not HSEnterMode=true Then StartGame
	If keycode = AddCreditKey Then Addcredit:EMReelCredits.setvalue(credits)
    If HSEnterMode Then HighScoreProcessKey(keycode)
End Sub

Sub JacksOpen_KeyUp(ByVal keycode)
	If keycode = PlungerKey Then Plunger.Fire
	If Gamestarted AND(NOT Tilted) Then
		If keycode = LeftFlipperKey Then
			LeftFlipper.RotateToStart
			PlaySound SoundFXDOF("PNK_MH_Flip_L_down",101,DOFOff,DOFContactors)
			StopSound "Buzz1"
		end if
		If keycode = RightFlipperKey Then
			RightFlipper.RotateToStart
			PlaySound SoundFXDOF("PNK_MH_Flip_R_down",102,DOFOff,DOFContactors)
			StopSound "Buzz"
		end if
	End If
End Sub


'******************************
'  INIT FIRST BALL
'****************************************

Sub Initialize
    GameOverFlasher.enabled = 0
    If B2SOn then Controller.B2SSetGameOver 0
	LightTRO1.state=1 : LightTRO2.state=1 : LightTRO3.state=1 : LightTRO4.state=1  ' light suit rollovers
	TopLanesUnlit = 0
	ResetTargets
	SpecialLit = 0
	Free100=0 : Free130=0 : Free170 = 0
	TargetMode = 1 : LightMode1.state=1 ' reset mode
	TargetValue = 1000 ' Target value - starts at 1k, +1k for each lane completed
	B1Lit=1:B2Lit=0:B3Lit=1
	SMFadePause=6 'LIGHT BUMPERS
	Score = 0
	BallsRemaining = BallsPerGame + 1 ' first ball
	TiltPause = 0
	Tilted = 0
	Tilt = 0
	MatchPause = 0
	GameOverPause = 0
	NextBallPause = 0
	GameStarted = 1
	Credits = Credits -1
	If credits < 1 Then DOF 142, DOFOff
	If B2SOn then
		Controller.B2SSetCredits credits
		Controller.B2SSetMatch 34, 0 'match lights off
		Controller.B2SSetTilt 33, 0
		Controller.B2SSetData 51, 0 'turn off 100k lights

	End if
	EMReelCredits.setvalue(credits)
	SaveHighScore
	EMReelMatch.setvalue 0
	BlinkingLogoOnOff(0)
	EMReel100k.setvalue 0
	EMReelGO.setvalue 0
'    If b2son then controller.B2SSetData 0, 3, 0, 0, 120, "reel"  'reset to zero
	ResetScore
	PlaySound "MekSound"
End Sub

Sub StartGame
	Initialize
	NextBallPause = 30
End Sub

Sub GetSpecial
	PlaySound SoundFXDOF("knocker",110,DOFPulse,DOFKnocker)
	Credits = Credits + 1
	DOF 142, DOFOn
	If Credits > 15 Then Credits = 15
	If B2SOn then Controller.B2SSetCredits credits
	EMReelCredits.setvalue(Credits)
	SaveHighScore
End Sub

Sub AddCredit
	Credits = Credits + 1
	DOF 142, DOFOn
	PlaySound "coin"
	If Credits > 15 Then Credits = 15
	If B2SOn then Controller.B2SSetCredits credits
	EMReelCredits.setvalue(Credits)
	SaveHighScore
End Sub

Sub AllLampsOff
	LightMode2.state=0 : LightMode3.state=0 : LightMode4.state=0 : LightMode5.state=0
	LightLILb.state=1 : LightRILb.state=1 : LightLOLb.state=1 : LightROLb.state=1 : LightLROb.state=1 : LightRROb.state=1
	LightLILb.state=0 : LightRILb.state=0 : LightLOLb.state=0 : LightROLb.state=0 : LightLROb.state=0 : LightRROb.state=0
	LightTRO1.state=0 : LightTRO2.state=0 : LightTRO3.state=0 : LightTRO4.state=0
	Bumper1Light.state = 0 : Bumper2Light.state = 0 : Bumper3Light.state = 0
	LROlit=0:RROlit=0:LOLlit=0:ROLlit=0:LILlit=0:RILlit=0
End Sub

'****************************************
'  DELAY TIMER LOOP
'****************************************

Sub DelayTimer_Timer
	If NextBallPause > 0 Then
		NextBallPause = NextBallPause - 1
		If NextBallPause = 0 Then
			BallsRemaining = BallsRemaining - 1
			CurrentBall = BallsPerGame - BallsRemaining +1
						If BallsRemaining = 0 Then
				MatchPause = 20
			Else
				If B2SOn then Controller.B2ssetballinplay 32, CurrentBall
				EMReelBall.SetValue(CurrentBall)
				Drain.Kick 65, 15
			    PlaySound SoundFXDOF("ballrel",109,DOFPulse,DOFContactors)
			End If
		End If
	End If

	If DropResetPause > 0 Then
		DropResetPause = DropResetPause  - 1
		If DropResetPause = 0 Then
			PlaySound "dropResetCenter"
			For each obj in Drops
				obj.IsDropped=false
			Next
		End If
	End If

	If RedResetPause > 0 Then
		RedResetPause = RedResetPause  - 1
		If RedResetPause = 0 Then
			PlaySound "dropResetCenter"
			Target2.IsDropped = True:Target4.IsDropped = True:Target6.IsDropped = True:Target8.IsDropped = True
		End If
	End If

     If TiltPause > 0 Then
        TiltPause = TiltPause - 1
        If TiltPause = 1 Then
			Tilt = Tilt - 1
			If Tilt > 0 Then
				TiltPause = tiltSettle
			End If
		End If
     End If

	If GameOverPause > 0 Then
		GameOverPause = GameOverPause - 1
		If GameOverPause = 15 Then
			PlaySound "MekSound"
			GameStarted = 0
			AllLampsOff
			SMFadePause=15
		End If
		If GameOverPause = 0 Then GameOver
	End If

	If MatchPause > 0 then
		MatchPause = MatchPause - 1
		If MatchPause = 0 then
			Match = 10 * (INT(10 * Rnd(1) ) ):If Match < 10 Then Match = "00"
			Select Case Match
				Case 0
					If B2SOn then Controller.B2SSetMatch 34,0 'match
					EMReelMatch.setvalue 0
				Case "00":
					If B2SOn then Controller.B2SSetMatch 34,1 'match
					EMReelMatch.setvalue 1
				Case 10:
					If B2SOn then Controller.B2SSetMatch 34,2
					EMReelMatch.setvalue 2
				Case 20:
					If B2SOn then Controller.B2SSetMatch 34,3
					EMReelMatch.setvalue 3
				Case 30:
					If B2SOn then Controller.B2SSetMatch 34,4
					EMReelMatch.setvalue 4
				Case 40:
					If B2SOn then Controller.B2SSetMatch 34,5
					EMReelMatch.setvalue 5
				Case 50:
					If B2SOn then Controller.B2SSetMatch 34,6
					EMReelMatch.setvalue 6
				Case 60:
					If B2SOn then Controller.B2SSetMatch 34,7
					EMReelMatch.setvalue 7
				Case 70:
					If B2SOn then Controller.B2SSetMatch 34,8
					EMReelMatch.setvalue 8
				Case 80:
					If B2SOn then Controller.B2SSetMatch 34,9
					EMReelMatch.setvalue 9
				Case 90:
					If B2SOn then Controller.B2SSetMatch 34,10
					EMReelMatch.setvalue 10
			End Select
			x = Match
	'		MatchTemp.Text= x'debug
			If Match = "00" Then x = 0
			If x = Score MOD 100 Then
				PlaySound SoundFXDOF("knocker",110,DOFPulse,DOFKnocker)
				Credits = Credits + 1
				DOF 142, DOFOn
				If Credits > 15 Then Credits = 15
				If B2SOn then Controller.B2SSetCredits credits
				EMReelCredits.setvalue(Credits)
				SaveHighScore
			End If
			GameOverPause = 20
		End If
	End If

	If FiveKPause > 0 Then
		FiveKPause = FiveKPause - 1
		If FiveKPause = 12 Then AddScore(1000):PlaySound SoundFXDOF("1000",146,DOFPulse,DOFChimes)
		If FiveKPause = 9 Then AddScore(1000):PlaySound SoundFXDOF("1000",146,DOFPulse,DOFChimes)
		If FiveKPause = 6 Then AddScore(1000):PlaySound SoundFXDOF("1000",146,DOFPulse,DOFChimes)
		If FiveKPause = 3 Then AddScore(1000):PlaySound SoundFXDOF("1000",146,DOFPulse,DOFChimes)
		If FiveKPause = 0 Then AddScore(1000):PlaySound SoundFXDOF("1000",146,DOFPulse,DOFChimes)
	End If
	If FiveHPause > 0 Then
		FiveHPause = FiveHPause - 1
		If FiveHPause = 12 Then AddScore(100):PlaySound SoundFXDOF("100",145,DOFPulse,DOFChimes)
		If FiveHPause = 9 Then AddScore(100):PlaySound SoundFXDOF("100",145,DOFPulse,DOFChimes)
		If FiveHPause = 6 Then AddScore(100):PlaySound SoundFXDOF("100",145,DOFPulse,DOFChimes)
		If FiveHPause = 3 Then AddScore(100):PlaySound SoundFXDOF("100",145,DOFPulse,DOFChimes)
		If FiveHPause = 0 Then AddScore(100):PlaySound SoundFXDOF("100",145,DOFPulse,DOFChimes)
	End If
 End Sub

'****************************************
'  LIGHT TIMER LOOP
'****************************************

Sub LightTimer_Timer
	If SMFadePause > 0 Then
		SMFadePause = SMFadePause - 1
		Select Case SMFadePause
			Case 15:
				Bumper1Light.state = 1
				If B2Lit=1 Then:Bumper2Light.state = 1
				Bumper3Light.state = 1
				If LROlit=1 Then LightLROb.state=1
				If RROlit=1 Then LightRROb.state=1
				If LOLlit=1 Then LightLOLb.state=1
				If ROLlit=1 Then LightROLb.state=1
				If LILlit=1 Then LightLILb.state=1
				If RILlit=1 Then LightRILb.state=1
			Case 14:
				Bumper1Light.state = 1
				If B2Lit=1 Then:Bumper2Light.state = 1
				Bumper3Light.state = 1
				If LROlit=1 Then LightLROb.state=0
				If RROlit=1 Then LightRROb.state=0
				If LOLlit=1 Then LightLOLb.state=0
				If ROLlit=1 Then LightROLb.state=0
				If LILlit=1 Then LightLILb.state=0
				If RILlit=1 Then LightRILb.state=0
			Case 13:
				Bumper1Light.state = 0
				If B2Lit=1 Then:Bumper2Light.state = 0
				Bumper3Light.state = 0
				If LROlit=1 Then LightLROb.state=1:LightLROb.state=0
				If RROlit=1 Then LightRROb.state=1:LightRROb.state=0
				If LOLlit=1 Then LightLOLb.state=1:LightLOLb.state=0
				If ROLlit=1 Then LightROLb.state=1:LightROLb.state=0
				If LILlit=1 Then LightLILb.state=1:LightLILb.state=0
				If RILlit=1 Then LightRILb.state=1:LightRILb.state=0
			Case 2:
				If Gamestarted AND(NOT Tilted) Then
				Bumper1Light.state = 1
				If B2Lit=1 Then:Bumper2Light.state = 1
				Bumper3Light.state = 1
					If LROlit=1 Then LightLROb.state=1:LightLROb.state=0
					If RROlit=1 Then LightRROb.state=1:LightRROb.state=0
					If LOLlit=1 Then LightLOLb.state=1:LightLOLb.state=0
					If ROLlit=1 Then LightROLb.state=1:LightROLb.state=0
					If LILlit=1 Then LightLILb.state=1:LightLILb.state=0
					If RILlit=1 Then LightRILb.state=1:LightRILb.state=0
				End If
			Case 1:
				If Gamestarted AND(NOT Tilted) Then
				Bumper1Light.state = 1
				If B2Lit=1 Then:Bumper2Light.state = 1
				Bumper3Light.state = 1
					If LROlit=1 Then LightLROb.state=1
					If RROlit=1 Then LightRROb.state=1
					If LOLlit=1 Then LightLOLb.state=1
					If ROLlit=1 Then LightROLb.state=1
					If LILlit=1 Then LightLILb.state=1
					If RILlit=1 Then LightRILb.state=1
				End If
			Case 0:
				If Gamestarted AND(NOT Tilted) Then
				Bumper1Light.state = 1
				If B2Lit=1 Then:Bumper2Light.state = 1
				Bumper3Light.state = 1
					If LROlit=1 Then LightLROb.state=1
					If RROlit=1 Then LightRROb.state=1
					If LOLlit=1 Then LightLOLb.state=1
					If ROLlit=1 Then LightROLb.state=1
					If LILlit=1 Then LightLILb.state=1
					If RILlit=1 Then LightRILb.state=1
				End If
		End Select
		Bumper1Light1.state = Bumper1Light.state
		Bumper1Light2.state = Bumper1Light.state
		Bumper2Light1.state = Bumper2Light.state
		Bumper2Light2.state = Bumper2Light.state
		Bumper3Light2.state = Bumper3Light.state
		Bumper1Light3.state = Bumper1Light.state
		Bumper2Light3.state = Bumper2Light.state
		Bumper3Light3.state = Bumper3Light.state
		Bumper3Light1.state = Bumper3Light.state
	End If
	If Target2.isdropped Then
		Target2LD.state = 1
	Else
		Target2LD.State = 0
	End If

	If Target3.isdropped Then
		Target3LD.state = 1
	Else
		Target3LD.State = 0
	End If

	If Target4.isdropped Then
		Target4LD.state = 1
	Else
		Target4LD.State = 0
	End If

	If Target5.isdropped Then
		Target5LD.state = 1
	Else
		Target5LD.State = 0
	End If
	If Target6.isdropped Then
		Target6LD.state = 1
	Else
		Target6LD.State = 0
	End If

	If Target7.isdropped Then
		Target7LD.state = 1
	Else
		Target7LD.State = 0
	End If

	If Target8.isdropped Then
		Target8LD.state = 1
	Else
		Target8LD.State = 0
	End If

End Sub

'****************************************
'  OBJECT HANDLERS
'****************************************

Sub Gate_Hit():Playsound "PNK_MH_gate":End Sub
Sub LeftFlipper_Collide(parm):RandomSoundFlipper():End Sub
Sub RightFlipper_Collide(parm):RandomSoundFlipper():End Sub

Sub RandomSoundFlipper()
	Select Case Int(Rnd*3)+1
		Case 1 : PlaySound "PNK_MH_Flip_hit_1"
		Case 2 : PlaySound "PNK_MH_Flip_hit_2"
		Case 3 : PlaySound "PNK_MH_Flip_hit_3"
	End Select
End Sub

'****************************************
'  RULES AND SCORING
'****************************************

Sub ScoreTimer_timer
	EMReelScore1.setvalue(Score10k)
	EMReelScore2.setvalue(ScoreK)
	EMReelScore3.setvalue(Score100)
	EMReelScore4.setvalue(Score10)
	EMReelScore5.setvalue(0)
End Sub

Sub ResetScore
	Scorex = 0
	Score10k = 0
	ScoreK = 0
	Score100 = 0
	Score10 = 0
	If B2SOn then Controller.B2SSetScorePlayer 1, 0
End Sub

Sub AddScore(sumtoadd)
	If Tilted=0 Then
		Select Case sumtoadd
		Case 500:
			FiveHPause = 13 '500
		Case 2000:
			FiveKPause = 4 '5000
		Case 3000:
			FiveKPause = 7 '5000
		Case 4000:
			FiveKPause = 10 '5000
		Case 5000:
			FiveKPause = 13 '5000
		Case 10,100,1000:
			Score = Score + sumtoadd
			' adjust individual reels
			Scorex = Score
			Score1M = Int (Scorex/1000000)'Calculate the value for the 1,000,000's digit
			Score100K=Int (Scorex/100000)'Calculate the value for the 100,000's digit
			Score10K=Int ((Scorex-(Score100k*100000))/10000) 'Calculate the value for the 10,000's digit
			ScoreK=Int((Scorex-(Score100k*100000)-(Score10K*10000))/1000) 'Calculate the value for the 1000's digit
			Score100=Int((Scorex-(Score100k*100000)-(Score10K*10000)-(ScoreK*1000))/100) 'Calculate the value for the 100's digit
			Score10=Int((Scorex-(Score100k*100000)-(Score10K*10000)-(ScoreK*1000)-(Score100*100))/10) 'Calculate the value for the 10's digit
			Score1=Int(Scorex-(Score100k*100000)-(Score10K*10000)-(ScoreK*1000)-(Score100*100)-(Score10*10)) 'Calculate the value for the 1's digit
			If B2SOn then
				Controller.B2SSetScorePlayer 1, scorex
				Controller.B2SSetData 51, Score100k 'set 100k lights
			End if
			EMReel100k.setvalue(Score100k)
			If Score >= 100000 AND Free100 <1 then
				GetSpecial
				Free100 = 1
			End If
			If Score >= 130000 AND Free130 <1 then
				GetSpecial
				Free130 = 1
			End If
			If Score >= 170000 AND Free170 <1 then
				GetSpecial
				Free170 = 1
			End If
			If sumtoAdd = 10 Then PlaySound SoundFXDOF("10",144,DOFPulse,DOFChimes)
			If sumtoAdd = 100 Then PlaySound SoundFXDOF("100",145,DOFPulse,DOFChimes)
			If sumtoAdd = 1000 Then PlaySound SoundFXDOF("1000",146,DOFPulse,DOFChimes)
		End Select
	End If

End Sub

Sub Sw10pt_hit(idx)
	AddScore(10) '10
End Sub

Sub TargetScore_timer()
	Select Case TSCount
		Case 1:
			If LightTRO1.state=0 then AddScore(1000)
		Case 2:
			If LightTRO2.state=0 then AddScore(1000)
		Case 3:
			If LightTRO3.state=0 then AddScore(1000)
		Case 4:
			If LightTRO4.state=0 then AddScore(1000)
		Case 5:
			CheckTargets
			me.enabled=0
	End Select
	tscount=tscount+1
End Sub

Sub CheckTargets
	' mode 1 - going for 2 Jacks (starting mode)
	' mode 2 - going for 3 Queens
	' mode 3 - going for Full House 3 Queens 2 Kings
	' mode 4 - going for Royal Flush Red
	' mode 5 - going for Special (Royal Flush Red)
	Select Case TargetMode
		Case 1
			If Target4Down=True AND Target7Down=True Then
				AdvanceMode
			End If
		Case 2
			If Target2Down=True AND Target5Down=True AND Target8Down=True Then
				AdvanceMode
			End If
		Case 3
			If Target2Down=True AND Target5Down=True AND Target8Down=True AND Target3Down=True AND Target6Down=True Then
				AdvanceMode
			End If
		Case 4
			If Target1Down=True AND Target3Down=True AND Target5Down=True AND Target7Down=True AND Target9Down=True Then
				AdvanceMode
			End If
		Case 5
			If Target1Down=True AND Target3Down=True AND Target5Down=True AND Target7Down=True AND Target9Down=True Then
				AdvanceMode
			End If
	End Select
End Sub

Sub AdvanceMode
	Select Case TargetMode
		Case 1:AddScore(5000):ResetTargets:LightMode1.state=0:LightMode2.state=1:TargetMode=2:B2Lit=1:LightLILb.state=1:LightRILb.state=1:LILlit=1:RILlit=1
		Case 2:AddScore(5000):ResetTargets:LightMode2.state=0:LightMode3.state=1:TargetMode=3
		Case 3:AddScore(5000):ResetTargets:LightMode3.state=0:LightMode4.state=1:TargetMode=4
		Case 4:AddScore(5000):ResetReds:LightMode4.state=0:LightMode5.state=1:TargetMode=5:SpecialLit = 1
		Case 5:AddScore(5000):ResetReds:GetSpecial
	End Select
End Sub

Sub ResetReds
	Target1Down=False:Target2Down=False:Target3Down=False:Target4Down=False
	Target5Down=False:Target6Down=False:Target7Down=False:Target8Down=False:Target9Down=False
	DropResetPause = 15
	RedResetPause = 25
End Sub

Sub ResetTargets
	Target1Down=False:Target2Down=False:Target3Down=False:Target4Down=False
	Target5Down=False:Target6Down=False:Target7Down=False:Target8Down=False:Target9Down=False
	DropResetPause = 15
End Sub

'****************************************
'  TARGET HANDLERS
'****************************************

Sub Target1_Hit()
	If Gamestarted AND(NOT Tilted) Then
		Target1Down=True
		SMFadePause=SMFadeResetValue
		playsound SoundFXDOF("dropLeft",111,DOFPulse,DOFDropTargets)
		AddScore (1000)
		if targetscore.enabled=false then
			TSCount=1
			TargetScore.enabled=1
		end If
	End If
End Sub

Sub Target2_Hit()
	If Gamestarted AND(NOT Tilted) Then
		Target2Down=True
		SMFadePause=SMFadeResetValue
		playsound SoundFXDOF("dropLeft",111,DOFPulse,DOFDropTargets)
		AddScore (1000)
			if targetscore.enabled=false then
				TSCount=1
				TargetScore.enabled=1
			end If
	End If
End Sub

Sub Target3_Hit()
	If Gamestarted AND(NOT Tilted) Then
		Target3Down=True
		SMFadePause=SMFadeResetValue
		playsound SoundFXDOF("dropLeft",111,DOFPulse,DOFDropTargets)
		AddScore (1000)
			if targetscore.enabled=false then
				TSCount=1
				TargetScore.enabled=1
			end If
	End If
End Sub

Sub Target4_Hit()
	If Gamestarted AND(NOT Tilted) Then
		Target4Down=True
		SMFadePause=SMFadeResetValue
		playsound SoundFXDOF("dropCenter",112,DOFPulse,DOFDropTargets)
		AddScore (1000)
			if targetscore.enabled=false then
				TSCount=1
				TargetScore.enabled=1
			end If
	End If
End Sub

Sub Target5_Hit()
	If Gamestarted AND(NOT Tilted) Then
		Target5Down=True
		SMFadePause=SMFadeResetValue
		playsound SoundFXDOF("dropCenter",112,DOFPulse,DOFDropTargets)
		AddScore (1000)
			if targetscore.enabled=false then
				TSCount=1
				TargetScore.enabled=1
			end If
	End If
End Sub

Sub Target6_Hit()
	If Gamestarted AND(NOT Tilted) Then
		Target6Down=True
		SMFadePause=SMFadeResetValue
		playsound SoundFXDOF("dropCenter",112,DOFPulse,DOFDropTargets)
		AddScore (1000)
			if targetscore.enabled=false then
				TSCount=1
				TargetScore.enabled=1
			end If
	End If
End Sub

Sub Target7_Hit()
	If Gamestarted AND(NOT Tilted) Then
		Target7Down=True
		SMFadePause=SMFadeResetValue
		playsound SoundFXDOF("dropRight",113,DOFPulse,DOFDropTargets)
		AddScore (1000)
			if targetscore.enabled=false then
				TSCount=1
				TargetScore.enabled=1
			end If
	End If
End Sub

Sub Target8_Hit()
	If Gamestarted AND(NOT Tilted) Then
		Target8Down=True
		SMFadePause=SMFadeResetValue
		playsound SoundFXDOF("dropRight",113,DOFPulse,DOFDropTargets)
		AddScore (1000)
			if targetscore.enabled=false then
				TSCount=1
				TargetScore.enabled=1
			end If
	End If
End Sub

Sub Target9_Hit()
	If Gamestarted AND(NOT Tilted) Then
		Target9Down=True
		SMFadePause=SMFadeResetValue
		playsound SoundFXDOF("dropRight",113,DOFPulse,DOFDropTargets)
		AddScore (1000)
			if targetscore.enabled=false then
				TSCount=1
				TargetScore.enabled=1
			end If
	End If
End Sub



'****************************************
'  TOP ROLLOVERS
'****************************************

Sub SwTop1_Hit()
	PlaySound "RollOver"
	DOF 120, DOFPulse
	If Gamestarted AND(NOT Tilted) Then
		If LightTRO1.state=1 Then
			SMFadePause=SMFadeResetValue
			AddScore(5000)
			LightTRO1.state=0
			LightLROb.state=1 'light bottom light (H=LRO/C=RRO/D=LOL/S=ROL)
			LROlit=1
			TargetValue = TargetValue+1000
			TopLanesUnlit = TopLanesUnlit + 1
			CheckTopRollovers()
		Else
			AddScore(500)
		End if
	End if
End Sub

Sub SwTop2_Hit()
	PlaySound "RollOver"
	DOF 103, DOFPulse
	If Gamestarted AND(NOT Tilted) Then
		If LightTRO2.state=1 Then
			SMFadePause=SMFadeResetValue
			AddScore(5000)
			LightTRO2.state=0
			LightLOLb.state=1 'light bottom light (H=LRO/C=RRO/D=LOL/S=ROL)
			LOLlit=1
			TargetValue = TargetValue+1000
			TopLanesUnlit = TopLanesUnlit + 1
			CheckTopRollovers()
		Else
			AddScore(500)
		End if
	End if
End Sub

Sub SwTop3_Hit()
	PlaySound "RollOver"
	DOF 104, DOFPulse
	If Gamestarted AND(NOT Tilted) Then
		If LightTRO3.state=1 Then
			SMFadePause=SMFadeResetValue
			AddScore(5000)
			LightTRO3.state=0
			LightROLb.state=1 'light bottom light (H=LRO/C=RRO/D=LOL/S=ROL)
			ROLlit=1
			TargetValue = TargetValue+1000
			TopLanesUnlit = TopLanesUnlit + 1
			CheckTopRollovers()
		Else
			AddScore(500)
		End if
	End if
End Sub

Sub SwTop4_Hit()
	PlaySound "RollOver"
	DOF 105, DOFPulse
	If Gamestarted AND(NOT Tilted) Then
		If LightTRO4.state=1 Then
			SMFadePause=SMFadeResetValue
			AddScore(5000)
			LightTRO4.state=0
			LightRROb.state=1 'light bottom light (H=LRO/C=RRO/D=LOL/S=ROL)
			RROlit=1
			TargetValue = TargetValue+1000
			TopLanesUnlit = TopLanesUnlit + 1
			CheckTopRollovers()
		Else
			AddScore(500)
		End if
	End if
End Sub


Sub CheckTopRollovers()

	If TopLanesUnlit = 4 Then
		LightTRO1.state=1 : LightTRO2.state=1 : LightTRO3.state=1 : LightTRO4.state=1
		LightLROb.state=1:LightLROb.state=0:LightRROb.state=1:LightRROb.state=0:
		LightLOLb.state=1:LightLOLb.state=0:LightROLb.state=1:LightROLb.state=0
		LROLit=0:RROLit=0:LOLlit=0:ROLlit=0
		TargetValue=1000
		TopLanesUnlit = 0
		AdvanceMode()
	End If
End Sub

'****************************************
'  BOTTOM ROLLOVERS
'****************************************

Sub SwLRO_Hit()
	PlaySound "RollOver"
	DOF 131, DOFPulse
	If LightLROb.state=1 Then
		AddScore(1000)
	Else
		AddScore(100)
	End If
End Sub

Sub SwRRO_Hit()
	PlaySound "RollOver"
	DOF 134, DOFPulse
	If LightRROb.state=1 Then
		AddScore(1000)
	Else
		AddScore(100)
	End If
End Sub

Sub SwLIL_Hit()
	SwLILPrim.RotX = 135
	PlaySound "RollOver"
	DOF 132, DOFPulse
	If LightLILb.state=1 Then
		AddScore(1000)
	Else
		AddScore(100)
	End If
End Sub

Sub SwLIL_unHit
	SwLILPrim.RotX = 160
End Sub

Sub SwRIL_Hit()
	SwRILPrim.RotX = 135
	PlaySound "RollOver"
	DOF 133, DOFPulse
	If LightRILb.state=1 Then
		AddScore(1000)
	Else
		AddScore(100)
	End If
End Sub

Sub SwRIL_unHit
	SwRILPrim.RotX = 160
End Sub

Sub SwLOL_Hit()
	SwLOLPrim.RotX = 135
	PlaySound "RollOver"
	DOF 131, DOFPulse
	If LightLOLb.state=1 Then
		AddScore(5000)
	Else
		AddScore(500)
	End If
End Sub

Sub SwLOL_unHit
	SwLOLPrim.RotX = 160
End Sub

Sub SwROL_Hit()
	SwROLPrim.RotX = 135
	PlaySound "RollOver"
	DOF 134, DOFPulse
	If LightROLb.state=1 Then
		AddScore(5000)
	Else
		AddScore(500)
	End If
End Sub

Sub SwRol_unHit
	SwROLPrim.RotX = 160
End Sub

'****************************************
'  BUMPER HANDLERS
'****************************************

Sub Bumper1_Hit()
    PlaySound SoundFXDOF("bumperLeft",103, DOFPulse,DOFContactors)
	If B1Lit=1 Then:AddScore(100):Else:AddScore(10):End If
End Sub

Sub Bumper2_Hit()
    PlaySound SoundFXDOF("bumperCenter",104, DOFPulse,DOFContactors)
	If B2Lit=1 Then:AddScore(1000):Else:AddScore(100):End If
End Sub

Sub Bumper3_Hit()
    PlaySound SoundFXDOF("bumperRight",105, DOFPulse,DOFContactors)
	If B3Lit=1 Then:AddScore(100):Else:AddScore(10):End If
End Sub

Sub Rubbers_Hit(idx):RandomSoundRubber():End Sub

Sub RandomSoundRubber()
	Select Case Int(Rnd*3)+1
		Case 1 : PlaySound "PNK_MH_rubber_hit_1"
		Case 2 : PlaySound "PNK_MH_rubber_hit_2"
		Case 3 : PlaySound "PNK_MH_rubber_hit_3"
	End Select
End Sub

Sub Wires_Hit(idx):RandomSoundWire():End Sub

Sub RandomSoundWire()
	Select Case Int(Rnd*3)+1
		Case 1 : PlaySound "PNK_MH_Wirehit_1"
		Case 2 : PlaySound "PNK_MH_Wirehit_2"
		Case 3 : PlaySound "PNK_MH_Wirehit_3"
	End Select
End Sub

'****************************************
'  DRAIN
'****************************************

Sub Drain_Hit:EndOfBall:End Sub 'RollingSoundsOff:SoundTimer.enabled=0:EndOfBall:End Sub

Sub EndOfBall()
	PlaySound "drain"
	DOF 141, DOFPulse
	SMFadePause=SMFadeResetValue:PlaySound "MekSound"
	'Enable again if Tilted
	If Tilted Then
		EMReelTILT.SetValue 0
		If B2SOn then Controller.B2SSetTilt 33, 0
		Tilted = 0
		Bumper1.Threshold = 0.75
		Bumper2.Threshold = 0.75
		Bumper3.Threshold = 0.75
	End If
	NextBallPause = 30
End Sub

'****************************************
'  TILT
'****************************************

Sub Bump()
	If Gamestarted AND(NOT Tilted) Then
		Tilt = Tilt + 1
		TiltPause = tiltSettle
		If Tilt > tiltAllow Then
			Tilted = 1
			EMReelTILT.SetValue 1
			Tilt = 0
			If B2SOn then Controller.B2SSetTilt 33, 1
			PlaySound "tilt"
			LeftFlipper.RotateToStart:StopSound "PNK_MH_Flip_L_up":StopSound "Buzz1"
			RightFlipper.RotateToStart:StopSound "PNK_MH_Flip_R_up":StopSound "Buzz"
			Bumper1.Threshold = 10
			Bumper2.Threshold = 10
			Bumper3.Threshold = 10
			SMFadePause=SMFadeResetValue
		End If
		TiltPause = 10
	End If
End Sub

Sub Mechchecktilt()
	If Gamestarted Then
			Tilted = 1
			LeftFlipper.RotateToStart
			RightFlipper.RotateToStart
			EMReelTILT.SetValue 1
			Tilt = 0
			If B2SOn then Controller.B2SSetTilt 33, 1
			PlaySound "tilt"
			LeftFlipper.RotateToStart:StopSound "PNK_MH_Flip_L_up":StopSound "Buzz1"
			RightFlipper.RotateToStart:StopSound "PNK_MH_Flip_R_up":StopSound "Buzz"
			Bumper1.Threshold = 10
			Bumper2.Threshold = 10
			Bumper3.Threshold = 10
			SMFadePause=SMFadeResetValue
	End If
End Sub

'****************************************
'  GAME OVER
'****************************************

Sub GameOver
	If B2SOn then
		Controller.B2ssetballinplay 32, 0
		Controller.B2SSetGameOver 1
	end if
	BlinkingLogoOnOff(1)
	GameOverFlasher.enabled = 1
	EMReelGO.setvalue 1
	EMReelBall.setvalue 0
	LeftFlipper.RotateToStart
	RightFlipper.RotateToStart
	If Score > CDbl(HighScore) Then
		GetSpecial
		HighScore=Score:HSA1 = 0:HSA2 = 0:HSA3 = 0:UpdatePostIt
		HighScoreEntryInit()
	End If
	StopSound "buzz"
	StopSound "buzz1"
	LastScore = Score:SaveHighScore
End Sub

'****************************************
'  HSTD + STORED DATA
'****************************************
Sub LoadHighScore
	Dim FileObj
	Dim ScoreFile
	Dim TextStr
    Dim SavedDataTemp1 'Credits
    Dim SavedDataTemp2 'LastScore
    Dim SavedDataTemp3 'HighScore
    Dim SavedDataTemp4 'HSA1
    Dim SavedDataTemp5 'HSA2
    Dim SavedDataTemp6 'HSA3
    Set FileObj=CreateObject("Scripting.FileSystemObject")
	If Not FileObj.FolderExists(UserDirectory) then
		Exit Sub
	End if
	If Not FileObj.FileExists(UserDirectory & "JacksOpen.txt") then
		SetDefaultHSTD:UpdatePostIt:SaveHighScore
		Exit Sub
	End if
	Set ScoreFile=FileObj.GetFile(UserDirectory & "JacksOpen.txt")
	Set TextStr=ScoreFile.OpenAsTextStream(1,0)
		If (TextStr.AtEndOfStream=True) then
			SetDefaultHSTD:UpdatePostIt:SaveHighScore
			Exit Sub
		End if
		SavedDataTemp1=TextStr.ReadLine ' credits
		SavedDataTemp2=Textstr.ReadLine ' LastScore
		SavedDataTemp3=Textstr.ReadLine ' HighScore
		SavedDataTemp4=Textstr.ReadLine ' HSA1
		SavedDataTemp5=Textstr.ReadLine ' HSA2
		SavedDataTemp6=Textstr.ReadLine ' HSA3
		TextStr.Close
	    Credits = CDbl(SavedDataTemp1)
		LastScore=SavedDataTemp2
		HighScore=SavedDataTemp3
		HSA1=SavedDataTemp4
		HSA2=SavedDataTemp5
		HSA3=SavedDataTemp6
		UpdatePostIt
	    Set ScoreFile = Nothing
	    Set FileObj = Nothing
End Sub

Sub SetDefaultHSTD  'bad data or missing file - reset and resave
	HighScore = DefaultHighScore
	HSA1 = DefaultHSA1
	HSA2 = DefaultHSA2
	HSA3 = DefaultHSA3
	Credits = 0
	LastScore = DefaultHighScore
	SaveHighScore
End Sub

Sub LastScoreReels
		Scorex = LastScore
		Score1M=Int (Scorex/1000000)'Calculate the value for the 1,000,000's digit
		Score100K=Int (Scorex/100000)'Calculate the value for the 100,000's digit
		Score10K=Int ((Scorex-(Score100k*100000))/10000) 'Calculate the value for the 10,000's digit
		ScoreK=Int((Scorex-(Score100k*100000)-(Score10K*10000))/1000) 'Calculate the value for the 1000's digit
		Score100=Int((Scorex-(Score100k*100000)-(Score10K*10000)-(ScoreK*1000))/100) 'Calculate the value for the 100's digit
		Score10=Int((Scorex-(Score100k*100000)-(Score10K*10000)-(ScoreK*1000)-(Score100*100))/10) 'Calculate the value for the 10's digit
		Score1=Int(Scorex-(Score100k*100000)-(Score10K*10000)-(ScoreK*1000)-(Score100*100)-(Score10*10)) 'Calculate the value for the 1's digit

		' residue from B2S exe file, if these seemingly unused lines are removed, the program crashes With
		' Line:1
		' Invalid procedure call or agrument: 'Asc"'

		If B2SOn then
			Controller.B2SSetData 0,Score10K  'Display the appropriate digits in the appropriate reel
			Controller.B2SSetData 1,ScoreK
			Controller.B2SSetData 2,Score100
			Controller.B2SSetData 3,Score10
 			Controller.B2SSetData 10, 1
			Controller.B2SSetData 11, 1
			Controller.B2SSetData 12, 1
			Controller.B2SSetData 51, Score100k 'set 100k lights
			Controller.B2SSetScorePlayer 1, scorex
		End if
		EMReel100k.setvalue(Score100k)
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

Sub UpdatePostIt
	Dim scoreStr
	Dim length
	scoreStr = cstr(HighScore)
	HS1.image = "PostitBL"
	HS2.image = "PostitBL"
	HS3.image = "PostitBL"
	HS4.image = "PostitBL"
	HS5.image = "PostitBL"
	HS6.image = "PostitBL"
	HS7.image = "PostitBL"
	if len(scoreStr) > 6 Then
		HS1.image = "postit" & mid(scoreStr, 1, 1)
		scoreStr = mid(scoreStr, 2)
	end if
	if len(scoreStr) > 5 Then
		HS2.image = "postit" & mid(scoreStr, 1, 1)
		scoreStr = mid(scoreStr, 2)
	end if
	if len(scoreStr) > 4 Then
		HS3.image = "postit" & mid(scoreStr, 1, 1)
		scoreStr = mid(scoreStr, 2)
	end if
	if len(scoreStr) > 3 Then
		HS4.image = "postit" & mid(scoreStr, 1, 1)
		scoreStr = mid(scoreStr, 2)
	end if
	if len(scoreStr) > 1 Then
		HS5.image = "postit" & mid(scoreStr, 1, 1)
		scoreStr = mid(scoreStr, 2)
	end if
	if len(scoreStr) > 1 Then
		HS6.image = "postit" & mid(scoreStr, 1, 1)
		scoreStr = mid(scoreStr, 2)
	end if
	HS7.image = "postit" & scoreStr
	HSName1.image = ImgFromCode(HSA1, 1)
	HSName2.image = ImgFromCode(HSA2, 2)
	HSName3.image = ImgFromCode(HSA3, 3)
End Sub



Sub SaveHighScore
	Dim FileObj
	Dim ScoreFile
	Set FileObj=CreateObject("Scripting.FileSystemObject")
	If Not FileObj.FolderExists(UserDirectory) then
		Exit Sub
	End if
	Set ScoreFile=FileObj.CreateTextFile(UserDirectory & "JacksOpen.txt",True)
		ScoreFile.WriteLine Credits
		ScoreFile.WriteLine LastScore
		ScoreFile.WriteLine HighScore
		ScoreFile.WriteLine HSA1
		ScoreFile.WriteLine HSA2
		ScoreFile.WriteLine HSA3
		ScoreFile.Close
	Set ScoreFile=Nothing
	Set FileObj=Nothing
End Sub

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
'  HS ENTER INITIALS
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
					SaveHighScore 'enter it
					HighScoreFlashTimer.Enabled = False
					HSEnterMode = False
				End If
		End Select
		UpdatePostIt
    End If
End Sub

Sub GameOverFlasher_Timer()
	Dim FlasherOnOff
	Dim indx
	FlasherOnOff = Int(Rnd*10)
	if FlasherOnOff < 6 then
        If B2SOn then Controller.B2SSetGameOver 1
    Else
        If B2SOn then Controller.B2SSetGameOver 0
	End If
End Sub

Sub InitTimer_Timer()
	If B2SOn then Controller.B2SSetScorePlayer 1, LastScore
	InitTimer.enabled = 0
	if LastScore >= 100000 then
		If B2SOn then Controller.B2SSetData 51, 1
	Else
		If B2SOn then Controller.B2SSetData 51, 0
	End If
	BlinkingLogoOnOff(1)
End Sub


Sub BlinkingLogoOnOff(BlinkOn)
	Dim BlinkingLight
	BlinkingLogo.enabled = BlinkOn
	for BlinkingLight = 45 to 50
		If B2SOn then
			Controller.B2SSetData BlinkingLight, 1
		End If
		BlinkingLightState(BlinkingLight - 45) = 1
	next
End Sub

Sub BlinkingLogo_Timer()
	Dim BlinkingLight, BlinkingRandom
	for BlinkingLight = 45 to 50
		BlinkingRandom = int(Rnd*25)
		if BlinkingLightState(BlinkingLight - 45) = 1 then
			' if the light is on, turn it off 20% of the time
			if BlinkingRandom < 5 then
				BlinkingLightState(BlinkingLight - 45) = 0
				If B2SOn then
					Controller.B2SSetData BlinkingLight, 0
				End If
			end if
		else
			' if the light is off, turn it on 40% of the time
			if (BlinkingRandom < 10) then
				BlinkingLightState(BlinkingLight - 45) = 1
				If B2SOn then
					Controller.B2SSetData BlinkingLight, 1
				End if
			end if
		end if
	next
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

Function AudioFade(tableobj) ' Fades between front and back of the table (for surround systems or 2x2 speakers, etc), depending on the Y position on the table. "JacksOpen" is the name of the table
  Dim tmp
  tmp = tableobj.y * 2 / JacksOpen.height-1
  If tmp > 0 Then
    AudioFade = Csng(tmp ^10)
  Else
    AudioFade = Csng(-((- tmp) ^10) )
  End If
End Function

Function AudioPan(tableobj) ' Calculates the pan for a tableobj based on the X position on the table. "JacksOpen" is the name of the table
  Dim tmp
  tmp = tableobj.x * 2 / JacksOpen.width-1
  If tmp > 0 Then
    AudioPan = Csng(tmp ^10)
  Else
    AudioPan = Csng(-((- tmp) ^10) )
  End If
End Function

Function Pan(ball) ' Calculates the pan for a ball based on the X position on the table. "JacksOpen" is the name of the table
    Dim tmp
    tmp = ball.x * 2 / JacksOpen.width-1
    If tmp > 0 Then
        Pan = Csng(tmp ^10)
    Else
        Pan = Csng(-((- tmp) ^10) )
    End If
End Function

Function AudioFade(ball) ' Can this be together with the above function ?
  Dim tmp
  tmp = ball.y * 2 / JacksOpen.height-1
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

Const tnob = 0 ' total number of balls
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
        StopSound "fx_ballrolling0"
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
  If JacksOpen.VersionMinor > 3 OR JacksOpen.VersionMajor > 10 Then
    PlaySound("fx_collide"), 0, Csng(velocity) ^2 / 200, Pan(ball1), 0, Pitch(ball1), 0, 0, AudioFade(ball1)
  Else
    PlaySound("fx_collide"), 0, Csng(velocity) ^2 / 200, Pan(ball1), 0, Pitch(ball1), 0, 0
  End if
End Sub

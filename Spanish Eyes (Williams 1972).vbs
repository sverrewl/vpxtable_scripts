Option Explicit

Const cGameName = "spanisheyes_1972"

On Error Resume Next
ExecuteGlobal GetTextFile("controller.vbs")
If Err Then MsgBox "You need the controller.vbs"
On Error Goto 0

Dim Digit(5)
Dim Lane(5)
Dim Target(6)
Dim TLight(12)
Dim LLight(5)
Dim RLight(5)
Dim ALight(4)
Dim BLight(4)
Dim HighScore
Dim Score
Dim EnableBallControl
Dim BallSize
Dim BallMass

BallSize = 25
BallMass = 1

EnableBallControl = 0

Set TLight(1)=Light1a
Set TLight(2)=Light2a
Set TLight(3)=Light3a
Set TLight(4)=Light4a
Set TLight(5)=Light5a
Set TLight(6)=Light6a
Set TLight(7)=Light1b
Set TLight(8)=Light2b
Set TLight(9)=Light3b
Set TLight(10)=Light4b
Set TLight(11)=Light5b
Set TLight(12)=Light6b

Set LLight(1)=LightA
Set LLight(2)=LightB
Set LLight(3)=LightC
Set LLight(4)=LightD
Set LLight(5)=LightE

'Set RLight(1)=RollLight1
'Set RLight(2)=RollLight2
'Set RLight(3)=RollLight3
'Set RLight(4)=RollLight4
'Set RLight(5)=RollLight5

Set BLight(1)=Bumperl1
Set BLight(2)=Bumperl2
Set BLight(3)=Bumperl3
Set BLight(4)=Bumperl4

Set ALight(1)=LoopLight1
Set ALight(2)=LoopLight2
Set ALight(3)=LeftLaneLight
Set ALight(4)=RightLaneLight

Dim Credits
Dim BallsToPlay
Dim GotReplay
Dim GotSpecial
Dim GotAddA
Dim GotAddB
Dim GameType
Dim HolePause
Dim BallOut
Dim BallActive
Dim ThouMove
Dim HunMove
Dim TenMove
Dim DigitClock
Dim DigitDo
Dim TargetDone
Dim LanesDone
Dim BumpOn
Dim CreditPause
Dim TiltSensor
Dim MachineTilt
Dim ResetPause
Dim Got100K

Dim f,g,h

Randomize

Sub GameTimer_Timer

	If TiltSensor>99 and MachineTilt=0 and BallActive=1 then
		PlaySound "Tilt"
		TiltBox.Text="TILT"
       	If B2SOn Then Controller.B2SSetTilt 33,1
		MachineTilt=1
'		TargetDone=0
'		LanesDone=0
		TenMove=0
		HunMove=0
		ThouMove=0
		for f=1 to 4
			EVAL("Bumper"&f).hashitevent = 0
		Next
		LightSeqTilt.Play SeqAllOff
	End If
	If TiltSensor>0 then
		TiltSensor=TiltSensor-1
	End If

	If ResetPause>0 then
		ResetPause=ResetPause-1
		If ResetPause<55 and ResetPause>0 and ResetPause/6=int(ResetPause/6)then
			For f=1 to 5
				If Digit(f)>0 then Digit(f)=Digit(f)-1
			Next
			Digit1.Text=Digit(1)
			Digit2.Text=Digit(2)
			Digit3.Text=Digit(3)
			Digit4.Text=Digit(4)
			Digit5.Text=Digit(5)
		End If
		If ResetPause=1 then
			BallOut=50
			BallsToPlay=5
			If B2SOn then
				Controller.B2SSetGameOver 35,0
				Controller.B2SSetTilt 33,0
				Controller.B2SSetMatch 34,0
				Controller.B2SSetScorePlayer 1, 0
				Controller.B2ssetballinplay 32, BallsToPlay
				Controller.B2SSetScorePlayer 5, HighScore
			End If
			For f=1 to 6
				Target(f)=0
			Next
			For f=1 to 5
				Lane(f)=0
			Next
			TargetDone=0
			LanesDone=0
			NewBall
		End If
	End If

	If CreditPause>0 then
		CreditPause=CreditPause-1
		If CreditPause=0 then
			If Credits<9 then
				Credits=Credits+1
				DOF 121, DOFOn
				CreditLight.state=1
				CreditBox.Text=Credits
				If B2SOn Then Controller.B2ssetCredits Credits
			End If
		End If
	End If

	If HolePause>0 then
		HolePause=HolePause-1
		If HolePause=0 then
			Hole.Kick 160, 10
			PlaySound SoundFXDOF("popper_ball",119,DOFPulse,DOFContactors)
			DOF 120, DOFPulse
		End If
	End If

	DigitDo=0
	If TenMove>0 then DigitDo=4
	If HunMove>0 then DigitDo=3
	If ThouMove>0 then DigitDo=2
	If DigitClock>0 then
		DigitClock=DigitClock-1
		DigitDo=0
	End If

	If DigitDo>0 then
	If DigitDo=4 then
		Digit(4)=Digit(4)+1
		PlaySound SoundFXDOF("Bell10",141,DOFPulse,DOFChimes)
		TenMove=TenMove-1
		DigitClock=5
	End If
	If DigitDo=3 then
		Digit(3)=Digit(3)+1
		PlaySound SoundFXDOF("Bell100",142,DOFPulse,DOFChimes)
		HunMove=HunMove-1
		DigitClock=5
	End If
	If DigitDo=2 then
		Digit(2)=Digit(2)+1
		PlaySound SoundFXDOF("Bell1000",143,DOFPulse,DOFChimes)
		ThouMove=ThouMove-1
		DigitClock=5
	End If
	If Digit(4)>9 then
		Digit(4)=Digit(4)-10
		Digit(3)=Digit(3)+1
	End If
	If Digit(3)>9 then
		Digit(3)=Digit(3)-10
		Digit(2)=Digit(2)+1
	End If
	If Digit(2)>9 then
		Digit(2)=Digit(2)-10
		Digit(1)=Digit(1)+1
	End If
		If Digit(1)>9 then
		Digit(1)=Digit(1)-10
		Got100K=Got100K+1
		OTTBox.Text=Got100K*100000
		if b2son then Controller.B2SSetScoreRollover 25, 1
	End If
	Digit1.Text=Digit(1)
	Digit2.Text=Digit(2)
	Digit3.Text=Digit(3)
	Digit4.Text=Digit(4)
	Digit5.Text=Digit(5)
	If B2SOn Then
		Controller.B2SSetReel 1, Digit(1)
		Controller.B2SSetReel 2, Digit(2)
		Controller.B2SSetReel 3, Digit(3)
		Controller.B2SSetReel 4, Digit(4)
		Controller.B2SSetReel 5, Digit(5)
	End If
	DigitDo=0
		If Got100K=0 and GameType=2 and Digit(2)>4 then
		If Digit(1)>4+(GotReplay*2) then
		GotReplay=GotReplay+1
		PlaySound SoundFXDOF("Knocker",122,DOFPulse,DOFKnocker)
		DOF 120, DOFPulse
		If Credits<9 then Credits=Credits+1:DOF 121, DOFOn
		CreditLight.state=1
		CreditBox.Text=Credits
		If B2SOn Then Controller.B2ssetCredits Credits
		End If
		End If
	End If

	If BallOut>0 then
		BallOut=BallOut-1
'		If TargetDone>0 and BallOut=50+(TargetDone*11) then
'			TargetDone=TargetDone-1
'			ThouMove=ThouMove+1
'			f=0
'			For g=7 to 12
'			If f=0 and TLight(g).State=LightStateOn then
'				TLight(g).State=LightStateOff
'				f=1
'			End If
'			Next
'		End If

		If BallOut=49 then
			If BallsToPlay>0 then
				PlaySound "NewBall"
				DisplayBalls
			Else
				Score=Score+ (Got100k * 100000)
				Score=Score+ (Digit(1)*10000)
				Score=Score+ (Digit(2)*1000)
				Score=Score+ (Digit(3)*100)
				Score=Score+ (Digit(4)*10)
				Score=Score+ Digit(5)
				If score>HighScore Then
					HighScore=score
					SaveValue "SpEyes","HS",HighScore
					Sbest.Text=FormatNumber(HighScore, 0, -1, 0, -1)
				End if
				BallBox1.Text=""
				GameOverBox.Text="GAME OVER"
				  If b2son then 
					Controller.B2SSetGameOver 35,1
					Controller.B2ssetballinplay 32, 0
					Controller.B2ssetPlayerUp 30, 0
					Controller.B2SSetScorePlayer 5, HighScore
				  End If
				AllLightsOff
			End If
		End If
		
		If BallOut=1 then
			If BallsToPlay>0 then
				BallActive=1
				DOF 135, DOFPulse
				BallKicker.CreateBall.image="pinball"
				BallKicker.Kick 100,5
				NewBall
			 Else
				BallActive=0
	'			AllLightsOff
				If GameType=2 then
					f=int(rnd*10)
					MatchBox2.Text="0"
					MatchBox1.Text=f
					If b2son then 
						if f=0 then 
							Controller.B2SSetMatch 34,100
						  Else
							Controller.B2SSetMatch 34,f*10
						End If
					End If
					If f=Digit(4) then
						PlaySound SoundFXDOF("Knocker",122,DOFPulse,DOFKnocker)
						DOF 120, DOFPulse
						If Credits<9 then Credits=Credits+1:DOF 121, DOFOn
						CreditLight.state=1
						CreditBox.Text=Credits
						If B2SOn Then Controller.B2ssetCredits Credits
					End If
				End If
			End If
		End If
		
	End If

End Sub

Sub SpEyes_KeyDown(ByVal keycode)


	If keycode = PlungerKey Then
		Plunger.PullBack
		PlaySound "plungerpull",0,1,0.25,0.25
	End If

	If keycode = LeftFlipperKey and MachineTilt=0 and BallActive=1 Then
		LeftFlipper.RotateToEnd
		PlaySound SoundFXDOF("fx_FlipperUp",101,DOFOn,DOFContactors), 0, .67, -0.05, 0.05
		PlaySound "Buzz",-1,.05,-0.05, 0.05
	End If
    
	If keycode = RightFlipperKey and MachineTilt=0 and BallActive=1 Then
		RightFlipper.RotateToEnd
		PlaySound SoundFXDOF("fx_FlipperUp",102,DOFOn,DOFContactors), 0, .67, 0.05, 0.05
		PlaySound "Buzz1",-1,.05,0.05,0.05
	End If
    
	If keycode = LeftTiltKey and MachineTilt=0 and BallActive=1 Then
		Nudge 90, 2
		TiltSensor=TiltSensor+50
	End If
    
	If keycode = RightTiltKey and MachineTilt=0 and BallActive=1 Then
		Nudge 270, 2
		TiltSensor=TiltSensor+50
	End If
    
	If keycode = CenterTiltKey and MachineTilt=0 and BallActive=1Then
		Nudge 0, 2
		TiltSensor=TiltSensor+50
	End If
    
    If keycode=AddCreditKey and CreditPause=0 and Credits<9 then
		CreditLight.state=1
    	PlaySound "CoinIn"
		CreditPause=60
	End If
	
	If keycode=StartGameKey and BallActive=0 then	
		If Credits>0 then
			BallActive=2
			Credits=Credits-1
			If Credits < 1 Then DOF 121, DOFOff
			If credits=0 then creditlight.state=0
			Score=0
			CreditBox.Text=Credits
			If B2SOn Then 
				Controller.B2ssetCredits Credits
				Controller.B2ssetplayerup 30, 1
				Controller.B2SSetGameOver 0
				Controller.B2SSetScoreRollover 25, 0
				Controller.B2SSetScorePlayer 5, HighScore
			End If
			TargetDone=0
			AllLightsOff
			MatchBox1.Text=""
			MatchBox2.Text=""
			OTTBox.Text=""
			GameOverBox.Text=""
			Got100K=0
			GotReplay=0
			ResetPause=400
			PlaySound "Reset"
		End If
	End If

  ' Manual Ball Control
	If keycode = 46 Then	 				' C Key
		If EnableBallControl = 1 Then
			EnableBallControl = 0
		Else
			EnableBallControl = 1
		End If
	End If
    If EnableBallControl = 1 Then
		If keycode = 48 Then 				' B Key
			If BCboost = 1 Then
				BCboost = BCboostmulti
			Else
				BCboost = 1
			End If
		End If
		If keycode = 203 Then BCleft = 1	' Left Arrow
		If keycode = 200 Then BCup = 1		' Up Arrow
		If keycode = 208 Then BCdown = 1	' Down Arrow
		If keycode = 205 Then BCright = 1	' Right Arrow
	End If


 	'If keycode=10 and BallActive=0 and GameType=2 then	
	'	GameType=1
	'	PlaySound "TargetsUp"
	'	AddWall1.IsDropped=False
	'	AddWall2.IsDropped=False
		'AddWall3.IsDropped=False
	''	AddWall4.IsDropped=True
	'End If
 
  	'If keycode=11 and BallActive=0 and GameType=1 then	
	'	GameType=2
	'	PlaySound "TargetsUp"
	'	AddWall1.IsDropped=True
	'	AddWall2.IsDropped=True
	'	AddWall3.IsDropped=True
	'	AddWall4.IsDropped=False
	'End If
    
End Sub

Sub SpEyes_KeyUp(ByVal keycode)

	If keycode = PlungerKey Then
		Plunger.Fire
		PlaySound "plunger",0,1,0.25,0.25
	End If
    
	If keycode = LeftFlipperKey Then
		LeftFlipper.RotateToStart
		If MachineTilt=0 and BallActive=1 Then PlaySound SoundFXDOF("fx_flipperdown",101,DOFOff,DOFContactors), 0, 1, -0.05, 0.05
		StopSound "Buzz"
	End If
    
	If keycode = RightFlipperKey Then
		RightFlipper.RotateToStart
		If MachineTilt=0 and BallActive=1 Then PlaySound SoundFXDOF("fx_flipperdown",102,DOFOff,DOFContactors), 0, 1, 0.05, 0.05
		StopSound "Buzz1"
	End If

    'Manual Ball Control
	If EnableBallControl = 1 Then
		If keycode = 203 Then BCleft = 0	' Left Arrow
		If keycode = 200 Then BCup = 0		' Up Arrow
		If keycode = 208 Then BCdown = 0	' Down Arrow
		If keycode = 205 Then BCright = 0	' Right Arrow
	End If


End Sub

Sub Drain_Hit()
	PlaySound "Drain"
	DOF 134, DOFPulse
	BallOut=100+(TargetDone*11)+(ThouMove*6)+(HunMove*6)+(TenMove*6)
	Drain.DestroyBall
	BallstoPlay = Ballstoplay - 1
	BallActive=2
	If MachineTilt=1 then 
		For f=1 to 6
			If Target(f)=1 then TLight(f+6).State=LightStateOn
		Next
		for f=1 to 4
			EVAL("Bumper"&f).hashitevent = 1
		Next
		LightSeqTilt.StopPlay
	End If
End Sub

Sub ShooterLane_Unhit()
	DOF 136, DOFPulse
End Sub

dim rastep, rbstep, rcstep, rdstep, restep, rfstep, rgstep, rhstep, lastep, lbstep, lcstep, ldstep, lestep, lfstep, lgstep

Sub Wallra_Hit
    Rubberra.Visible = 0
    Rubberra1.Visible = 1
    rastep = 0
    Wallra.TimerEnabled = 1
End Sub
 
Sub Wallra_Timer
    Select Case rastep
        Case 3:Rubberra1.Visible = 0:Rubberra2.Visible = 1
        Case 3:Rubberra2.Visible = 0:Rubberra3.Visible = 1
        Case 4:Rubberra3.Visible = 0:Rubberra.Visible = 1:Wallra.TimerEnabled = 0:
    End Select
    rastep = rastep + 1
End Sub

Sub Wallrb_Hit
    Rubberrb.Visible = 0
    Rubberrb1.Visible = 1
    rbstep = 0
    Wallrb.TimerEnabled = 1
End Sub
 
Sub Wallrb_Timer
    Select Case rbstep
        Case 3:Rubberrb1.Visible = 0:Rubberrb2.Visible = 1
        Case 3:Rubberrb2.Visible = 0:Rubberrb3.Visible = 1
        Case 4:Rubberrb3.Visible = 0:Rubberrb.Visible = 1:Wallrb.TimerEnabled = 0:
    End Select
    rbstep = rbstep + 1
End Sub

Sub Wallrc_Hit
    Rubberrc.Visible = 0
    Rubberrc1.Visible = 1
    rcstep = 0
    Wallrc.TimerEnabled = 1
End Sub
 
Sub Wallrc_Timer
    Select Case rcstep
        Case 3:Rubberrc1.Visible = 0:Rubberrc2.Visible = 1
        Case 3:Rubberrc2.Visible = 0:Rubberrc3.Visible = 1
        Case 4:Rubberrc3.Visible = 0:Rubberrc.Visible = 1:Wallrc.TimerEnabled = 0:
    End Select
    rcstep = rcstep + 1
End Sub

Sub Wallrd_Hit
	If MachineTilt=0 then TenMove=TenMove+1
    Rubberrd.Visible = 0
    Rubberrd1.Visible = 1
    rdstep = 0
    Wallrd.TimerEnabled = 1
End Sub
 
Sub Wallrd_Timer
    Select Case rdstep
        Case 3:Rubberrd1.Visible = 0:Rubberrd2.Visible = 1
        Case 3:Rubberrd2.Visible = 0:Rubberrd3.Visible = 1
        Case 4:Rubberrd3.Visible = 0:Rubberrd.Visible = 1:Wallrd.TimerEnabled = 0:
    End Select
    rdstep = rdstep + 1
End Sub

Sub Wallre_Hit
	If MachineTilt=0 then TenMove=TenMove+1
    Rubberre.Visible = 0
    Rubberre1.Visible = 1
    restep = 0
    Wallre.TimerEnabled = 1
End Sub
 
Sub Wallre_Timer
    Select Case restep
        Case 3:Rubberre1.Visible = 0:Rubberre2.Visible = 1
        Case 3:Rubberre2.Visible = 0:Rubberre3.Visible = 1
        Case 4:Rubberre3.Visible = 0:Rubberre.Visible = 1:Wallre.TimerEnabled = 0:
    End Select
    restep = restep + 1
End Sub

Sub Wallrf_Hit
	If MachineTilt=0 then TenMove=TenMove+1
    Rubberrf.Visible = 0
    Rubberrf1.Visible = 1
    rfstep = 0
    Wallrf.TimerEnabled = 1
End Sub
 
Sub Wallrf_Timer
    Select Case rfstep
        Case 3:Rubberrf1.Visible = 0:Rubberrf2.Visible = 1
        Case 3:Rubberrf2.Visible = 0:Rubberrf3.Visible = 1
        Case 4:Rubberrf3.Visible = 0:Rubberrf.Visible = 1:Wallrf.TimerEnabled = 0:
    End Select
    rfstep = rfstep + 1
End Sub

Sub Wallrg_Hit
	If MachineTilt=0 then TenMove=TenMove+1
    Rubberrg.Visible = 0
    Rubberrg1.Visible = 1
    rgstep = 0
    Wallrg.TimerEnabled = 1
End Sub
 
Sub Wallrg_Timer
    Select Case rgstep
        Case 3:Rubberrg1.Visible = 0:Rubberrg2.Visible = 1
        Case 3:Rubberrg2.Visible = 0:Rubberrg3.Visible = 1
        Case 4:Rubberrg3.Visible = 0:Rubberrg.Visible = 1:Wallrg.TimerEnabled = 0:
    End Select
    rgstep = rgstep + 1
End Sub

Sub Wallrh_Hit
    Rubberrh.Visible = 0
    Rubberrh1.Visible = 1
    rhstep = 0
    Wallrh.TimerEnabled = 1
End Sub
 
Sub Wallrh_Timer
    Select Case rhstep
        Case 3:Rubberrh1.Visible = 0:Rubberrh2.Visible = 1
        Case 3:Rubberrh2.Visible = 0:Rubberrh3.Visible = 1
        Case 4:Rubberrh3.Visible = 0:Rubberrh.Visible = 1:Wallrh.TimerEnabled = 0:
    End Select
    rhstep = rhstep + 1
End Sub

Sub wallla_Hit
    rubberla.Visible = 0
    rubberla1.Visible = 1
    lastep = 0
    wallla.TimerEnabled = 1
End Sub
 
Sub wallla_Timer
    Select Case lastep
        Case 3:rubberla1.Visible = 0:rubberla2.Visible = 1
        Case 3:rubberla2.Visible = 0:rubberla3.Visible = 1
        Case 4:rubberla3.Visible = 0:rubberla.Visible = 1:wallla.TimerEnabled = 0:
    End Select
    lastep = lastep + 1
End Sub

Sub walllb_Hit
    rubberlb.Visible = 0
    rubberlb1.Visible = 1
    lbstep = 0
    walllb.TimerEnabled = 1
End Sub
 
Sub walllb_Timer
    Select Case lbstep
        Case 3:rubberlb1.Visible = 0:rubberlb2.Visible = 1
        Case 3:rubberlb2.Visible = 0:rubberlb3.Visible = 1
        Case 4:rubberlb3.Visible = 0:rubberlb.Visible = 1:walllb.TimerEnabled = 0:
    End Select
    lbstep = lbstep + 1
End Sub

Sub walllc_Hit
    rubberlc.Visible = 0
    rubberlc1.Visible = 1
    lcstep = 0
    Wallrc.TimerEnabled = 1
End Sub
 
Sub walllc_Timer
    Select Case lcstep
        Case 3:rubberlc1.Visible = 0:rubberlc2.Visible = 1
        Case 3:rubberlc2.Visible = 0:rubberlc3.Visible = 1
        Case 4:rubberlc3.Visible = 0:rubberlc.Visible = 1:walllc.TimerEnabled = 0:
    End Select
    lcstep = lcstep + 1
End Sub

Sub wallld_Hit
	If MachineTilt=0 then TenMove=TenMove+1
    rubberld.Visible = 0
    rubberld1.Visible = 1
    ldstep = 0
    wallld.TimerEnabled = 1
End Sub
 
Sub wallld_Timer
    Select Case ldstep
        Case 3:rubberld1.Visible = 0:rubberld2.Visible = 1
        Case 3:rubberld2.Visible = 0:rubberld3.Visible = 1
        Case 4:rubberld3.Visible = 0:rubberld.Visible = 1:wallld.TimerEnabled = 0:
    End Select
    ldstep = ldstep + 1
End Sub

Sub wallle_Hit
	If MachineTilt=0 then TenMove=TenMove+1
    rubberle.Visible = 0
    rubberle1.Visible = 1
    lestep = 0
    wallle.TimerEnabled = 1
End Sub
 
Sub wallle_Timer
    Select Case lestep
        Case 3:rubberle1.Visible = 0:rubberle2.Visible = 1
        Case 3:rubberle2.Visible = 0:rubberle3.Visible = 1
        Case 4:rubberle3.Visible = 0:rubberle.Visible = 1:wallle.TimerEnabled = 0:
    End Select
    lestep = lestep + 1
End Sub

Sub walllf_Hit
	If MachineTilt=0 then TenMove=TenMove+1
    rubberlf.Visible = 0
    rubberlf1.Visible = 1
    lfstep = 0
    walllf.TimerEnabled = 1
End Sub
 
Sub walllf_Timer
    Select Case lfstep
        Case 3:rubberlf1.Visible = 0:rubberlf2.Visible = 1
        Case 3:rubberlf2.Visible = 0:rubberlf3.Visible = 1
        Case 4:rubberlf3.Visible = 0:rubberlf.Visible = 1:walllf.TimerEnabled = 0:
    End Select
    lfstep = lfstep + 1
End Sub

Sub walllg_Hit
	If MachineTilt=0 then TenMove=TenMove+1
    rubberlg.Visible = 0
    rubberlg1.Visible = 1
    lgstep = 0
    walllg.TimerEnabled = 1
End Sub
 
Sub walllg_Timer
    Select Case lgstep
        Case 3:rubberlg1.Visible = 0:rubberlg2.Visible = 1
        Case 3:rubberlg2.Visible = 0:rubberlg3.Visible = 1
        Case 4:rubberlg3.Visible = 0:rubberlg.Visible = 1:walllg.TimerEnabled = 0:
    End Select
    lgstep = lgstep + 1
End Sub

Sub Sling1_Slingshot()
	PlaySound SoundFXDOF("Sling",103,DOFPulse,DOFContactors)
	DOF 104, DOFPulse
	If MachineTilt=0 then TenMove=TenMove+1
End Sub

Sub Sling2_Slingshot()
	PlaySound SoundFXDOF("Sling",107,DOFPulse,DOFContactors)
	If MachineTilt=0 then TenMove=TenMove+1
End Sub

Sub Sling3_Slingshot()
	PlaySound SoundFXDOF("Sling",109,DOFPulse,DOFContactors)
	If MachineTilt=0 then TenMove=TenMove+1
End Sub

Sub Sling4_Slingshot()
	PlaySound SoundFXDOF("Sling",110,DOFPulse,DOFContactors)
	If MachineTilt=0 then TenMove=TenMove+1
End Sub

Sub Sling5_Slingshot()
	PlaySound SoundFXDOF("Sling",108,DOFPulse,DOFContactors)
	If MachineTilt=0 then TenMove=TenMove+1
End Sub

Sub Sling6_Slingshot()
	PlaySound SoundFXDOF("Sling",105,DOFPulse,DOFContactors)
	DOF 106, DOFPulse
	If MachineTilt=0 then TenMove=TenMove+1
End Sub

Sub Sling7_Slingshot()
	PlaySound "Sling"
	If MachineTilt=0 then TenMove=TenMove+1
End Sub

Sub Sling8_Slingshot()
	PlaySound "Sling"
	If MachineTilt=0 then TenMove=TenMove+1
End Sub

Sub LaneA_Hit
	DOF 127, DOFPulse
	If MachineTilt=0 then
	PlaySound "LaneSwitch"
	'LLight(1).State=LightStateOn
	HunMove=HunMove+5
	CheckLanes(1)
	End If
End Sub

Sub LaneB_Hit
	DOF 128, DOFPulse
	If MachineTilt=0 then
	PlaySound "LaneSwitch"
	'LLight(2).State=LightStateOn
	HunMove=HunMove+5
	CheckLanes(2)
	End If
End Sub

Sub LaneC_Hit
	DOF 129, DOFPulse
	If MachineTilt=0 then
	PlaySound "LaneSwitch"
	'LLight(3).State=LightStateOn
	HunMove=HunMove+5
	CheckLanes(3)
	End If
End Sub

Sub LaneD_Hit
	DOF 130, DOFPulse
	If MachineTilt=0 then
	PlaySound "LaneSwitch"
	'LLight(4).State=LightStateOn
	HunMove=HunMove+5
	CheckLanes(4)
	End If
End Sub

Sub LaneE_Hit
	DOF 131, DOFPulse
	If MachineTilt=0 then
	PlaySound "LaneSwitch"
	'LLight(5).State=LightStateOn
	HunMove=HunMove+5
	CheckLanes(5)
	End If
End Sub

Sub CheckLanes(f)
	If Lane(f)=0 then LanesDone=LanesDone+1
	Lane(f)=1
	LLight(f).State=LightStateOn
	If LanesDone=5 then
'		If GameType=2 and GotSpecial=0 then
'			ALight(1).State=LightStateOn
'			ALight(2).State=LightStateOn
'		End If
'		If GameType=1 and GotAddA=0 then
'			ALight(1).State=LightStateOn
'			ALight(2).State=LightStateOn
'		End If
		BallsToPlay=BallsToPlay+1
		PlaySound SoundFXDOF("Bell10",141,DOFPulse,DOFChimes)
		PlaySound SoundFXDOF("Bell100",142,DOFPulse,DOFChimes)
		PlaySound SoundFXDOF("Bell1000",143,DOFPulse,DOFChimes)
		DisplayBalls
		For f=1 to 5
			LLight(f).State=LightStateOff
			Lane(f)=0
		Next 
		LanesDone=0
		If GotAddA=0 then
			GotAddA=1
		End If
	End If
End Sub

Sub LeftOutlane_Hit
	DOF 132, DOFPulse
	If MachineTilt=0 then
	PlaySound "LaneSwitch"
	If TargetDone=6 then
		If GameType=1 then
		BallsToPlay=BallsToPlay+1
		PlaySound SoundFXDOF("Bell10",141,DOFPulse,DOFChimes)
		PlaySound SoundFXDOF("Bell100",142,DOFPulse,DOFChimes)
		PlaySound SoundFXDOF("Bell1000",143,DOFPulse,DOFChimes)
		DisplayBalls
		Else
		ThouMove=ThouMove+5
		End If
	End If
	ALight(3).State=LightStateOff
	ALight(4).State=LightStateOff
	If TargetDone<6 then ThouMove=ThouMove+1
	End If
End Sub

Sub RightOutlane_Hit
	DOF 133, DOFPulse
	If MachineTilt=0 then
	PlaySound "LaneSwitch"
	If TargetDone=6 then
		If GameType=1 then
		BallsToPlay=BallsToPlay+1
		PlaySound SoundFXDOF("Bell10",141,DOFPulse,DOFChimes)
		PlaySound SoundFXDOF("Bell100",142,DOFPulse,DOFChimes)
		PlaySound SoundFXDOF("Bell1000",143,DOFPulse,DOFChimes)
		DisplayBalls
		Else
		ThouMove=ThouMove+5
		End If
	End If
	ALight(3).State=LightStateOff
	ALight(4).State=LightStateOff
	If TargetDone<6 then ThouMove=ThouMove+1
	End If
End Sub

Sub TopRollover_Hit
	TopRollover_prim.z = -3
	If MachineTilt=0 then
		PlaySound "Rollover"
		BumpOn=1
		For f=1 to 4
			BLight(f).State=LightStateOn
		Next
	End If
End Sub

Sub TopRollover_Unhit
	Toprollover_prim.z = 0
End Sub

Sub LoopRollover_Hit
	If MachineTilt=0 then
		PlaySound "Rollover"
		If TargetDone<6 then 
			ThouMove=ThouMove+5
'		If GotSpecial=0 and TargetDone=6 then
'			GotSpecial=1
		  else
			If GameType=2 then
				BallsToPlay=BallsToPlay+1
				PlaySound SoundFXDOF("Bell10",141,DOFPulse,DOFChimes)
				PlaySound SoundFXDOF("Bell100",142,DOFPulse,DOFChimes)
				PlaySound SoundFXDOF("Bell1000",143,DOFPulse,DOFChimes)
				DisplayBalls
				ALight(1).State=LightStateOff
				ALight(2).State=LightStateOff
				For f= 1 to 6
					TLight(f).state=0
					TLight(f+6).state=0
					Target(f)=0
				Next
				TargetDone=0
			  Else
				PlaySound SoundFXDOF("Knock",122,DOFPulse,DOFKnocker)
				DOF 120, DOFPulse
				If Credits<9 then 
					Credits=Credits+1
					DOF 121, DOFOn
					CreditLight.state=1
					CreditBox.Text=Credits
				End If
			End If
		End If
	End If
End Sub

Sub LeftRollover_Hit
	LeftRollover_Prim.z = -3
	If MachineTilt=0 then
		PlaySound "Rollover"
		TenMove=TenMove+1
	End If
End Sub

Sub LeftRollover_Unhit
	Leftrollover_prim.z = -0
End Sub

Sub RightRollover_Hit
	Rightrollover_prim.z = -3
	If MachineTilt=0 then
		PlaySound "Rollover"
		TenMove=TenMove+1
	End If
End Sub

Sub RightRollover_Unhit
	Rightrollover_prim.z = 0
End Sub

Sub BottomRollover_Hit
	bottomrollover_prim.z = -3
	If MachineTilt=0 then
		PlaySound "Rollover"
		TenMove=TenMove+1
	End If
End Sub

Sub BottomRollover_Unhit
	bottomrollover_prim.z = 0
End Sub

Sub Bumper1_Hit
	PlaySound SoundFXDOF("",111,DOFPulse,DOFContactors)
	PlaySoundAt "fx_bumper4", Bumper1
	DOF 112, DOFPulse
	If MachineTilt=0 then
		If BumpOn=1 then
			HunMove=HunMove+1
		Else
			TenMove=TenMove+1
		End If
	End If
End Sub

Sub Bumper2_Hit
	PlaySound SoundFXDOF("",113,DOFPulse,DOFContactors)
	PlaySoundAt "fx_bumper4", Bumper2
	DOF 114, DOFPulse
	If MachineTilt=0 then
		HunMove=HunMove+1
	End If
End Sub

Sub Bumper3_Hit
	PlaySound SoundFXDOF("",115,DOFPulse,DOFContactors)
	PlaySoundAt "fx_bumper4", Bumper3
	DOF 116, DOFPulse
	If MachineTilt=0 then
		If BumpOn=1 then
			HunMove=HunMove+1
		Else
			TenMove=TenMove+1
		End If
	End If
End Sub

Sub Bumper4_Hit
	PlaySound SoundFXDOF("",117,DOFPulse,DOFContactors)
	PlaySoundAt "fx_bumper4", Bumper4
	DOF 118, DOFPulse
	If MachineTilt=0 then
		If BumpOn=1 then
			HunMove=HunMove+1
		Else
			TenMove=TenMove+1
		End If
	End If
End Sub

Sub Target1_Hit
	DOF 125, DOFPulse
	If MachineTilt=0 then
	HunMove=HunMove+3
	CheckTarget(1)
	End If
End Sub

Sub Target2_Hit
	DOF 125, DOFPulse
	If MachineTilt=0 then
	HunMove=HunMove+3
	CheckTarget(2)
	End If
End Sub

Sub Target3_Hit
	DOF 126, DOFPulse
	If MachineTilt=0 then
	HunMove=HunMove+3
	CheckTarget(3)
	End If
End Sub

Sub Target4_Hit
	DOF 126, DOFPulse
	If MachineTilt=0 then
	HunMove=HunMove+3
	CheckTarget(4)
	End If
End Sub

Sub Target5_Hit
	DOF 123, DOFPulse
	If MachineTilt=0 then
	HunMove=HunMove+3
	CheckTarget(5)
	End If
End Sub

Sub Target6_Hit
	DOF 124, DOFPulse
	If MachineTilt=0 then
	HunMove=HunMove+3
	CheckTarget(6)
	End If
End Sub

Sub CheckTarget(f)
	If Target(f)=0 then TargetDone=TargetDone+1
	Target(f)=1
	TLight(f).State=LightStateOn
	TLight(f+6).State=LightStateOn
	If TargetDone=6 then
		ALight(1).State=LightStateOn
		ALight(2).State=LightStateOn
		ALight(3).State=LightStateOn
		ALight(4).State=LightStateOn
'		If GameType=1 and GotAddB=0 then
'			GotAddB=1
'			BallsToPlay=BallsToPlay+1
'			PlaySound "Bell10"
'			PlaySound "Bell100"
'			PlaySound "Bell1000"
'			DisplayBalls
'		End If
	End If
End Sub

Sub Hole_Hit()
	HolePause=40+(TargetDone+6)
	PlaySound "KickerIn"
	If MachineTilt=0 then
		If TargetDone>0 then
			ThouMove=ThouMove+TargetDone
		Else
			HunMove=HunMove+5
			HolePause=100
		End If
	End If
End Sub


Sub AllLightsOff
	For f=1 to 4
		ALight(f).State=LightStateOff
		BLight(f).State=LightStateOff
	Next 
	For f=1 to 5
		LLight(f).State=LightStateOff
'		RLight(f).State=LightStateOff
	Next 
	For f=1 to 12
		TLight(f).State=LightStateOff
	Next 
End Sub


Sub AllLightsOn
	For f=1 to 4
		ALight(f).State=LightStateOn
		BLight(f).State=LightStateOn
	Next 
	For f=1 to 5
		LLight(f).State=LightStateOn
'		RLight(f).State=LightStateOn
	Next 
	For f=1 to 12
		TLight(f).State=LightStateOn
	Next 
End Sub

Sub AllLightsBlink
	For f=1 to 4
		ALight(f).State=LightStateBlinking
		BLight(f).State=LightStateBlinking
	Next 
	For f=1 to 5
		LLight(f).State=LightStateBlinking
'		RLight(f).State=LightStateBlinking
	Next 
	For f=1 to 12
		TLight(f).State=LightStateBlinking
	Next 
End Sub


Sub NewBall
	GotAddA=0
	GotAddB=0
	GotSpecial=0
	TiltSensor=0
	MachineTilt=0
	TiltBox.Text=""
	If b2son then Controller.B2SSetTilt 33,0
	For f=1 to 4
		BLight(f).state=LightStateOff
	Next
	BumperL2.State=LightStateOn
	BumpOn=0
'	TargetDone=0
'	LanesDone=0
'	For f=1 to 5
'		RLight(f).State=LightStateOn
'		Lane(f)=0
'	Next
	DisplayBalls
End Sub

Sub DisplayBalls
	If BallsToPlay>10 then BallsToPlay=10
	For f=1 to 5
		EVAL("BallBox"&f).text=""
	Next
	If BallsToPlay<6 Then
		EVAL("BallBox"&BallsToPlay).text=BallsToPlay
	  Else
		BallBox5.Text=BallsToPlay
	End If
	If b2son then Controller.B2ssetballinplay 32, BallsToPlay
End Sub

Sub SpEyes_Init()
	LoadEM
    HighScore=LoadValue("SpEyes","HS") ' Load saved Highscore
	If HighScore="" then
		HighScore=100
		SaveValue "SpEyes","HS",HighScore
	Else
		Highscore=CDbl(LoadValue("SpEyes","HS"))
	End If
	Sbest.Text=FormatNumber(HighScore, 0, -1, 0, -1)
	PlaySound "TargetsUp"
	GameType=2
    Credits=LoadValue("SpEyes","Credits") ' Load saved Credits
	If Credits="" then
		Credits=0
		SaveValue "SpEyes","Credits",Credits
	Else
		Credits=CDbl(LoadValue("SpEyes","Credits"))
	End If
	If B2SOn then
		Tiltbox.Visible = false
		OTTbox.visible = False
		Digit1.visible = False
		Digit2.visible = False
		Digit3.visible = False
		Digit4.visible = False
		Digit5.visible = False
		GameOverBox.visible = False
		creditbox.visible = False
		matchbox1.visible = False
		matchbox2.visible = False
		sbest.visible = False
		ballBox1.visible = False
		ballBox2.visible = False
		ballBox3.visible = False
		ballBox4.visible = False
		ballBox5.visible = False
		HSbox.visible = False	
		Controller.B2ssetCredits Credits
	End If
	If Credits>0 then 
		CreditLight.state=1
		DOF 121, DOFOn
	  Else
		CreditLight.state=0
	End If
			
	For f=1 to 5
		Digit(f)=0
		Lane(f)=0
		Target(f)=0
	Next
	Target(6)=0
End Sub

Sub SpEyes_Exit
	SaveValue "SpEyes","Credits",Credits
	If B2SOn Then Controller.stop
End Sub

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

' *********************************************************************
'                      Supporting Ball & Sound Functions
' *********************************************************************

Function Vol(ball) ' Calculates the Volume of the sound based on the ball speed
    Vol = Csng(BallVel(ball) ^2 / 2000)
End Function

Function Pan(ball) ' Calculates the pan for a ball based on the X position on the table. "SPEyes" is the name of the table
    Dim tmp
    tmp = ball.x * 2 / SPEyes.width-1
	Pan = tmp
End Function

Function Fade(ball) ' Calculates the pan for a ball based on the Y position on the table. "SPEyes" is the name of the table
    Dim tmp
    tmp = ball.x * 2 / SPEyes.height-1
	Fade = tmp
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
        If BallVel(BOT(b) ) > 1 AND BOT(b).z < 30 Then
            rolling(b) = True
            PlaySound("fx_ballrolling" & b), -1, Vol(BOT(b)), Pan(BOT(b) ), 0, Pitch(BOT(b) ), 1, 0, Fade(BOT(b))
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
	PlaySound("fx_collide"), 0, Csng(velocity) ^2 / 2000, Pan(ball1), 0, Pitch(ball1), 0, 0, Fade(Ball1)
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

'*************Hit Sound Routines
Sub Pins_Hit (idx)
	PlaySound "pinhit_low", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0, Fade(ActiveBall)
End Sub

Sub Targets_Hit (idx)
	PlaySound "target", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0, Fade(ActiveBall)
End Sub

Sub Metals_Thin_Hit (idx)
	PlaySound "metalhit_thin", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, Fade(ActiveBall)
End Sub

Sub Metals_Medium_Hit (idx)
	PlaySound "metalhit_medium", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, Fade(ActiveBall)
End Sub

Sub Metals2_Hit (idx)
	PlaySound "metalhit2", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, Fade(ActiveBall)
End Sub

Sub Gates_Hit (idx)
	PlaySound "gate4", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, Fade(ActiveBall)
End Sub

Sub Rubbers_Hit(idx)
 	dim finalspeed
  	finalspeed=SQR(activeball.velx * activeball.velx + activeball.vely * activeball.vely)
 	If finalspeed > 20 then 
		PlaySound "fx_rubber2", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, Fade(ActiveBall)
	End if
	If finalspeed >= 6 AND finalspeed <= 20 then
 		RandomSoundRubber()
 	End If
End Sub

Sub RubberWheel_hit
 	dim finalspeed
  	finalspeed=SQR(activeball.velx * activeball.velx + activeball.vely * activeball.vely)
 	If finalspeed > 20 then 
		PlaySound "fx_rubber2", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, Fade(ActiveBall)
	End if
	If finalspeed >= 6 AND finalspeed <= 20 then
 		RandomSoundRubber()
 	End If
End sub

Sub Posts_Hit(idx)
 	dim finalspeed
  	finalspeed=SQR(activeball.velx * activeball.velx + activeball.vely * activeball.vely)
 	If finalspeed > 14 then 
		PlaySound "fx_rubber2", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, Fade(ActiveBall)
	End if
	If finalspeed >= 4 AND finalspeed <= 14 then
 		RandomSoundRubber()
 	End If
End Sub

Sub RandomSoundRubber()
	Select Case Int(Rnd*3)+1
		Case 1 : PlaySound "rubber_hit_1", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, Fade(ActiveBall)
		Case 2 : PlaySound "rubber_hit_2", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, Fade(ActiveBall)
		Case 3 : PlaySound "rubber_hit_3", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, Fade(ActiveBall)
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
		Case 1 : PlaySound "flip_hit_1", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, Fade(ActiveBall)
		Case 2 : PlaySound "flip_hit_2", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, Fade(ActiveBall)
		Case 3 : PlaySound "flip_hit_3", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, Fade(ActiveBall)
	End Select
End Sub

'*****************************************
'		FLIPPER SHADOWS
'*****************************************

sub FlipperTimer_Timer()
	FlipperLSh.RotZ = LeftFlipper.currentangle
	FlipperRSh.RotZ = RightFlipper.currentangle
'    flipperleft_prim.rotz = leftflipper.currentangle
'    flipperright_prim.rotz = rightflipper.currentangle
End Sub

'*****************************************
'           BALL SHADOW
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
        If BOT(b).X < SPEyes.Width/2 Then
            BallShadow(b).X = ((BOT(b).X) - (Ballsize/16) + ((BOT(b).X - (SPEyes.Width/2))/17))' + 13
        Else
            BallShadow(b).X = ((BOT(b).X) + (Ballsize/16) + ((BOT(b).X - (SPEyes.Width/2))/17))' - 13
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

'*********************************************************************
'                     Supporting Ball & Sound Functions
'*********************************************************************

Function AudioFade(tableobj) ' Fades between front and back of the table (for surround systems or 2x2 speakers, etc), depending on the Y position on the table. "SPEyes" is the name of the table
	Dim tmp
    tmp = tableobj.y * 2 / SPEyes.height-1
    If tmp > 0 Then
		AudioFade = Csng(tmp ^10)
    Else
        AudioFade = Csng(-((- tmp) ^10) )
    End If
End Function

Function AudioPan(tableobj) ' Calculates the pan for a tableobj based on the X position on the table. "SPEyes" is the name of the table
    Dim tmp
    tmp = tableobj.x * 2 / SPEyes.width-1
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

' Thalamus : Exit in a clean and proper way
Sub SpEyes_exit()
  Controller.Pause = False
  Controller.Stop
End Sub


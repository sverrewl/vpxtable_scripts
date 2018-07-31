'*** Cabaret ***'

Option Explicit  'Force explicit variable declaration.
Randomize

' Thalamus 2018-07-19
' Added/Updated "Positional Sound Playback Functions" and "Supporting Ball & Sound Functions"
' No special SSF tweaks yet.

On Error Resume Next
ExecuteGlobal GetTextFile("controller.vbs")
If Err Then MsgBox "You need the controller.vbs in order to run this table, available in the vp10 package"
On Error Goto 0

Const cGameName = "cabaret_1968"

dim score(4)		'Reel score
dim truesc(4)		'True reel score if overring the numer og digits
dim reel(5)			'Maximum number of scoring reels
dim state
dim credit			'Credits
dim eg				'End Game
dim currpl			'Current player status
dim playno			'Number of players
dim up(4)			'Player up
dim play(4)			'Player status
dim rst				'Reset Variable
dim bip				'Ball in play
dim ballrelenabled	'Used for shooter sound
dim match(10)
dim tilt
dim tiltsens
dim rep(4)
dim matchnumb
dim bell			'Global scoring sound variable
dim points
dim replay1
dim replay2
dim replay3
dim replay4
dim hisc			'High Score to date
dim wv				'Wheel Value (positiion)
dim wn 				'Wheel Change animation steps
dim sbv				'Super Bonus value
dim i				'Indexing variable
dim scn				'Single points score timer
dim scn1			'Single points end score timer
dim scn2			'Multiple points score timer
dim scn3			'Multiple points end score timer
dim obj
dim bgl1,bgl2,bgl3,bgl4

ExecuteGlobal GetTextFile("core.vbs")

Sub Timer1_timer
	if B2SOn then
		if bgl1=0 Then
			Controller.B2SSetData 10,1
			bgl1=1
		Else
			Controller.B2SSetData 10,0
			bgl1=0
		end If
	end If
end Sub

Sub Timer2_timer
	if B2SOn then
		if bgl2=0 Then
			Controller.B2SSetData 11,1
			bgl2=1
		Else
			Controller.B2SSetData 11,0
			bgl2=0
		end If
	end If
end Sub

Sub Timer3_timer
	if B2SOn then
		if bgl3=0 Then
			Controller.B2SSetData 12,1
			bgl3=1
		Else
			Controller.B2SSetData 12,0
			bgl3=0
		end If
	end If
end Sub

Sub Timer4_timer
	if B2SOn then
		if bgl4=0 Then
			Controller.B2SSetData 13,1
			bgl4=1
		Else
			Controller.B2SSetData 13,0
			bgl4=0
		end If
	end If
end Sub

Sub Table1_Init
	LoadEM
	If ShowDT=false Then
		for each obj in DesktopItems
			obj.visible=False
		Next
	end If
	bgl1=0
	bgl2=0
	bgl3=0
	bgl4=0
	set play(1)=up1
	set play(2)=up2
	set play(3)=up3
	set play(4)=up4
	set match(0)=m0
	set match(1)=m1
	set match(2)=m2
	set match(3)=m3
	set match(4)=m4
	set match(5)=m5
	set match(6)=m6
	set match(7)=m7
	set match(8)=m8
	set match(9)=m9
	set reel(1)=reel1
	set reel(2)=reel2
	set reel(3)=reel3
	set reel(4)=reel4
	set reel(5)=reel5
	replay1=6500
	replay2=9800
	replay3=16500
	replay4=19800
	loadhs
	if hisc="" then hisc=1000
    reel5.setvalue(hisc)
    if wv="" then wv=0
    for i=0 to 19
    wheel(i).isdropped=True
    next
    wheel(wv).isdropped=false
    if credit="" then
		credit=0
	Else
		DOF 130, DOFOn
	End If
    credittxt.text=credit
    if matchnumb="" then matchnumb=int(rnd(1)*9)
    select case(matchnumb)
    case 0:
    m0.text="0"
    case 1:
    m1.text="1"
    case 2:
    m2.text="2"
    case 3:
    m3.text="3"
    case 4:
    m4.text="4"
    case 5:
    m5.text="5"
    case 6:
    m6.text="6"
    case 7:
    m7.text="7"
    case 8:
    m8.text="8"
    case 9:
    m9.text="9"
    end select
    for i=1 to 4
    currpl=i
    reel(i).setvalue(score(i))
    sbreel.setvalue(sbv)
    next
    currpl=0
	If B2SOn Then
		timer1.enabled=True
		timer2.enabled=True
		timer3.enabled=True
		timer4.enabled=True
		if matchnumb=0 then
			Controller.B2SSetMatch 10
		else
			Controller.B2SSetMatch matchnumb
		end if
		Controller.B2SSetScoreRolloverPlayer1 0
		Controller.B2SSetScoreRolloverPlayer2 0
		Controller.B2SSetScoreRolloverPlayer3 0
		Controller.B2SSetScoreRolloverPlayer4 0
		Controller.B2SSetShootAgain 0
		Controller.B2SSetTilt 1
		Controller.B2SSetCredits Credit
		Controller.B2SSetGameOver 1
		Controller.B2SSetData 81,0
		Controller.B2SSetData 82,0
		Controller.B2SSetData 83,0
		Controller.B2SSetData 84,0
	End If
	for i=1 to 4
		If B2SOn Then
			Controller.B2SSetScorePlayer i, 0
		End If
	next
    If B2SOn then Controller.B2SSetScorePlayer5 sbv

End Sub


Sub Table1_KeyDown(ByVal keycode)

	If keycode = PlungerKey Then
	Plunger.PullBack
	End If

    if keycode = 6 then
	playsound "coin3"
	coindelay.enabled=true
	If B2SOn Then
		Controller.B2SSetCredits Credit
	end if
	end if

    if keycode = 5 then
	playsound "coin3"
	coindelay1.enabled=true
	If B2SOn Then
		Controller.B2SSetCredits Credit
	end if
	end if

	if keycode = 2 and credit>0 and state=false and playno=0 then
	credit=credit-1
	If credit < 1 Then DOF 130, DOFOff
	credittxt.text=credit
    eg=0
    playno=1
    playno1.state=1
    currpl=1
    play(currpl).state=1
    playsound "click"
    playsound "initialize"
    rst=0
    bip=1
    resettimer.enabled=true
		If B2SOn Then
			Controller.B2SSetTilt 0
			Controller.B2SSetGameOver 0
			Controller.B2SSetMatch 0
			Controller.B2SSetCredits Credit

			Controller.B2SSetCanPlay 1
			Controller.B2SSetPlayerUp 1
			Controller.B2SSetData 81,1
			Controller.B2SSetData 82,0
			Controller.B2SSetData 83,0
			Controller.B2SSetData 84,0
			Controller.B2SSetScoreRolloverPlayer1 0

		End If
    end if

    if keycode = 2 and credit>0 and state=true and playno>0 and playno<4 and bip<2 then
    credit=credit-1
	If credit < 1 Then DOF 130, DOFOff
    credittxt.text=credit
    playno=playno+1
    if playno=2 then playno2.state=1
    if playno=3 then playno3.state=1
    if playno=4 then playno4.state=1
    playsound "click"
	If B2SOn Then
		Controller.B2SSetCredits Credit
		Controller.B2SSetCanPlay playno
	end if
    end if

    if state=true and tilt=false then

	If keycode = LeftFlipperKey Then
		LeftFlipper.RotateToEnd
		PlaySound SoundFXDOF("FlipperUp",101,DOFOn,DOFFlippers)
	End If

	If keycode = RightFlipperKey Then
		RightFlipper.RotateToEnd
        PlaySound SoundFXDOF("FlipperUp",102,DOFOn,DOFFlippers)
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
		Tilt=1
		mechchecktilt
	End If

	end if

End Sub

Sub Table1_KeyUp(ByVal keycode)

	If keycode = PlungerKey Then
		Plunger.Fire
	End If

	If keycode = LeftFlipperKey Then
		LeftFlipper.RotateToStart
		stopsound "buzz"
		if state=true and tilt=false then PlaySound SoundFXDOF("FlipperDown",101,DOFOff,DOFFlippers)
	End If

	If keycode = RightFlipperKey Then
		RightFlipper.RotateToStart
		stopsound "buzz"
        if state=true and tilt=false then PlaySound SoundFXDOF("FlipperDown",102,DOFOff,DOFFlippers)
	End If

End Sub

sub flippertimer_timer()
	LFlip.RotY = LeftFlipper.CurrentAngle+90
	RFlip.RotY = RightFlipper.CurrentAngle+90
end sub

sub newgame
	state=true
	eg=0
	sbv=0
	for i=0 to 3
	score(i)=0
	truesc(i)=0
	rep(i)=0
	next
	bip5.text=" "
	bip1.text="1"
	for i=0 to 9
	match(i).text=" "
	next
	tilttext.text=" "
	gamov.text=" "
	tilt=false
	tiltsens=0
	wheelcheck
	bip=1
	nb.CreateBall
	nb.kick 90,6
	DOF 126, DOFPulse
	If B2SOn then Controller.B2SSetBallInPlay 1
end sub

sub resettimer_timer
    rst=rst+1
    reel1.resettozero
    reel2.resettozero
    reel3.resettozero
    reel4.resettozero
    sbreel.resettozero
	If B2SOn Then
		Controller.B2SSetScorePlayer1 0
		Controller.B2SSetScorePlayer2 0
		Controller.B2SSetScorePlayer3 0
		Controller.B2SSetScorePlayer4 0
		Controller.B2SSetScorePlayer5 0
	end if
    if rst=14 then
    playsound "kickerkick"
    end if
    if rst=18 then
    newgame
    resettimer.enabled=false
    end if
end sub

sub coindelay_timer
	playsound "click"
	credit=credit+5
	DOF 130, DOFOn
	credittxt.text=credit
    coindelay.enabled=false
	If B2SOn Then
		Controller.B2SSetCredits Credit
	end if
end sub

sub coindelay1_timer
	playsound "click"
	credit=credit+1
	DOF 130, DOFOn
	credittxt.text=credit
    coindelay1.enabled=false
	If B2SOn Then
		Controller.B2SSetCredits Credit
	end if
end sub


'*** General points scoring ***

sub addscore(points)

    if tilt=false then bell=0
    if points=10 then matchnumb=matchnumb+1
    if matchnumb=10 then matchnumb=0

	if points = 1 or points = 10 or points = 100 then scn=1
	if points = 300 then scn2=3

    if points = 100 then
    reel(currpl).addvalue(100)
    bell=100
    scn1=1
    end if

    if points = 10 then
    reel(currpl).addvalue(10)
    bell=10
    scn1=1
    end if

    if points = 1 then
    reel(currpl).addvalue(1)
    bell=1
    scn1=1
    end if

    if points = 300 then
    reel(currpl).addvalue(300)
    playsound "motorshort1s"
    scn3=3
    bell=100
    end if

    if points = 1 or points = 10 or points = 100 then scn1=0 : scntimer.enabled=true : end if
    if points = 300 then scn3=0 : scntimer1.enabled=true : end if

    score(currpl)=score(currpl)+points
    truesc(currpl)=truesc(currpl)+points
	If B2SOn Then
		Controller.B2SSetScore currpl,score(currpl)
	end if
    if score(currpl)>9999 then
    score(currpl)=score(currpl)-10000
    rep(currpl)=0
    end if

    if score(currpl)=>replay1 and rep(currpl)=0 then
    credit=credit+1
	DOF 130, DOFOn
    playsound SoundFXDOF("knocke",129,DOFPulse,DOFKnocker)
	DOF 128, DOFPulse
    credittxt.text=credit
    rep(currpl)=1
    playsound "click"
    end if

    if score(currpl)=>replay2 and rep(currpl)=1 then
    credit=credit+1
	DOF 130, DOFOn
    playsound SoundFXDOF("knocke",129,DOFPulse,DOFKnocker)
	DOF 128, DOFPulse
    credittxt.text=credit
    rep(currpl)=2
    playsound "click"
	If B2SOn Then
		Controller.B2SSetCredits Credit
	end if
    end if

    if score(currpl)=>replay3 and rep(currpl)=2 then
    credit=credit+1
	DOF 130, DOFOn
    playsound SoundFXDOF("knocke",129,DOFPulse,DOFKnocker)
	DOF 128, DOFPulse
    credittxt.text=credit
    rep(currpl)=3
    playsound "click"
	If B2SOn Then
		Controller.B2SSetCredits Credit
	end if
	end If

    if score(currpl)=>replay4 and rep(currpl)=3 then
    credit=credit+1
	DOF 130, DOFOn
    playsound SoundFXDOF("knocke",129,DOFPulse,DOFKnocker)
	DOF 128, DOFPulse
    credittxt.text=credit
    rep(currpl)=4
    playsound "click"
	If B2SOn Then
		Controller.B2SSetCredits Credit
	end if
    end if
end sub

sub scntimer_timer
    scn1=scn1 + 1
    if bell=100 then playsound SoundFXDOF("bell100",143,DOFPulse,DOFChimes),0, 0.45, 0, 0
    if bell=10 then playsound SoundFXDOF("bell10",142,DOFPulse,DOFChimes),0, 0.3, 0, 0
    if bell=1 then playsound SoundFXDOF("bell1",141,DOFPulse,DOFChimes),0, 0.15, 0, 0
    if scn1=scn then scntimer.enabled=false
end sub

sub scntimer1_timer
	scn3=scn3 + 1
    if bell=100 then playsound SoundFXDOF("bell1",143,DOFPulse,DOFChimes),0, 0.45, 0, 0
    if bell=10 then playsound SoundFXDOF("bell10",142,DOFPulse,DOFChimes),0, 0.3, 0, 0
    if bell=1 then playsound SoundFXDOF("bell1",141,DOFPulse,DOFChimes),0, 0.15, 0, 0
    if scn3=scn2 then scntimer1.enabled=false
end sub


'*** Ball handling ***

sub ballhome_hit
    ballrelenabled=1
    sbv=0
    sbreel.resettozero
	If B2SOn Then
		Controller.B2SSetScorePlayer5 0
	end if
    Light3.State=0
end sub

sub ballhome_unhit
	DOF 132, DOFPulse
end sub

sub ballrelease_hit
	if ballrelenabled=1 then playsound "launchball": ballrelenabled=0: end if
end sub

Sub Drain_Hit()
	DOF 131, DOFPulse
	playsound "drainshorter"
	Drain.DestroyBall
	nextball
End Sub

sub nextball
	if tilt=true then
	tilt=false
	tilttext.text=" "
	If B2SOn Then Controller.B2SSetTilt 0:Controller.B2SSetShootAgain 0
	tiltseq.stopplay
	end if
	if (Light3.state)=1 then
	playsound "kickerkick"
	ballreltimer.enabled=true
	else
	currpl=currpl+1
	If B2SOn Then
		Controller.B2SSetData 81,0
		Controller.B2SSetData 82,0
		Controller.B2SSetData 83,0
		Controller.B2SSetData 84,0
		Controller.B2SSetData (80+currpl),1
	end If
	end if
	if currpl>playno then
	bip=bip+1
	if bip>5 then
	playsound "motorleer"
	If B2SOn then Controller.B2SSetBallInPlay 0
	eg=1
	ballreltimer.enabled=true
	else
	if state=true and tilt=false then
	play(currpl-1).state=0
	currpl=1
	play(currpl).state=1
	If B2SOn Then
		Controller.B2SSetData 81,0
		Controller.B2SSetData 82,0
		Controller.B2SSetData 83,0
		Controller.B2SSetData 84,0
		Controller.B2SSetData (80+currpl),1
	end If
	playsound "kickerkick"
	ballreltimer.enabled=true
	end if
	If B2SOn then Controller.B2SSetBallInPlay bip
	select case (bip)
	case 1:
	bip1.text="1"
	case 2:
	bip1.text=" "
	bip2.text="2"
	case 3:
	bip2.text=" "
	bip3.text="3"
	case 4:
	bip3.text=" "
	bip4.text="4"
	case 5:
	bip4.text=" "
	bip5.text="5"
	end select
	end if
	end if
	if currpl>1 and currpl<(playno+1) then
	if state=true and tilt=false then
	play(currpl-1).state=0
	play(currpl).state=1
	If B2SOn Then
		Controller.B2SSetData 81,0
		Controller.B2SSetData 82,0
		Controller.B2SSetData 83,0
		Controller.B2SSetData 84,0
		Controller.B2SSetData (80+currpl),1
	end If
	playsound "kickerkick"
	ballreltimer.enabled=true
	end if
	end if
end Sub

sub ballreltimer_timer
	if eg=1 then
	matchnum
	bip3.text=" "
	bip5.text=" "
	state=false
	for i=1 to 4
	if truesc(i)>hisc then
	hisc=truesc(i)
	reel5.setvalue(hisc)
	end if
	next
	play(currpl-1).state=0
	playno=0
	gamov.text="GAME OVER"
	If B2SOn Then
		Controller.B2SSetGameOver 1
	end if
	savehs
	ballreltimer.enabled=false
	else
	nb.CreateBall
	nb.kick 90,6
    ballreltimer.enabled=false
    end if
end sub

sub matchnum
    select case(matchnumb)
    case 0:
    m0.text="0"
    case 1:
    m1.text="1"
    case 2:
    m2.text="2"
    case 3:
    m3.text="3"
    case 4:
    m4.text="4"
    case 5:
    m5.text="5"
    case 6:
    m6.text="6"
    case 7:
    m7.text="7"
    case 8:
    m8.text="8"
    case 9:
    m9.text="9"
    end select
	If B2SOn Then

		if matchnumb=0 then
			Controller.B2SSetMatch 10
		else
			Controller.B2SSetMatch matchnumb
		end if
	end if
    for i=1 to playno
    if matchnumb=(score(i) mod 10) then
    credit=credit+1
	DOF 130, DOFOn
	If B2SOn Then
		Controller.B2SSetCredits Credit
	end if
    playsound SoundFXDOF("knocke",129,DOFPulse,DOFKnocker)
	DOF 128, DOFPulse
    credittxt.text= credit
    playsound "click"
    end if
    next
end sub

Sub CheckTilt
	If Tilttimer.Enabled = True Then
	TiltSens = TiltSens + 1
	if TiltSens = 3 Then
	Tilt = True
	tilttext.text="TILT"
	If B2SOn Then Controller.B2SSetTilt 1
	post.isdropped=true
	tiltsens = 0
	playsound "tilt"
	turnoff
	End If
	Else
	TiltSens = 0
	Tilttimer.Enabled = True
	End If
End Sub

Sub MechCheckTilt

	Tilt = True
	LeftFlipper.RotateToStart
	RightFlipper.RotateToStart
	tilttext.text="TILT"
	If B2SOn Then Controller.B2SSetTilt 1
	post.isdropped=true
	tiltsens = 0
	playsound "tilt"
	turnoff

End Sub


Sub Tilttimer_Timer()
	Tilttimer.Enabled = False
End Sub

sub turnoff
	if (post.isdropped)=false then playsound "postdown"
	post.isdropped=true
	tiltseq.play seqalloff
end sub

Sub Gate_hit()
	Playsound "gate"
End Sub


'**********Sling Shot Animations
' Rstep and Lstep  are the variables that increment the animation
'****************
Dim RStep, Lstep

Sub RightSlingShot_Slingshot
	if tilt = false then
    PlaySound SoundFXDOF("right_slingshot",105,DOFPulse,DOFContactors), 0, 1, 0.05, 0.05
	DOF 106, DOFPulse
    RSling.Visible = 0
    RSling1.Visible = 1
    sling1.TransZ = -20
    RStep = 0
    RightSlingShot.TimerEnabled = 1
    addscore 1
	end if
End Sub

Sub RightSlingShot_Timer
    Select Case RStep
        Case 3:RSLing1.Visible = 0:RSLing2.Visible = 1:sling1.TransZ = -10
        Case 4:RSLing2.Visible = 0:RSLing.Visible = 1:sling1.TransZ = 0:RightSlingShot.TimerEnabled = 0
    End Select
    RStep = RStep + 1
End Sub

Sub LeftSlingShot_Slingshot
	if tilt=false then
	PlaySound SoundFXDOF("left_slingshot",103,DOFPulse,DOFContactors),0,1,-0.05,0.05
	DOF 104, DOFPulse
	LSling.Visible = 0
	LSling1.Visible = 1
	sling2.TransZ = -20
	LStep = 0
	LeftSlingShot.TimerEnabled = 1
	addscore 1
	end if
End Sub

Sub LeftSlingShot5_Slingshot
	if tilt=false then
	PlaySound "left_slingshot",0,1,-0.05,0.05
	LSling.Visible = 0
	LSling1.Visible = 1
	sling2.TransZ = -20
	LStep = 0
	LeftSlingShot.TimerEnabled = 1
	addscore 1
	end if
End Sub

Sub RightSlingShot5_Slingshot
	if tilt=false then
	PlaySound "left_slingshot",0,1,-0.05,0.05
	LSling.Visible = 0
	LSling1.Visible = 1
	sling2.TransZ = -20
	LStep = 0
	LeftSlingShot.TimerEnabled = 1
	addscore 1
	end if
End Sub

Sub LeftSlingShot_Timer
    Select Case LStep
        Case 3:LSLing1.Visible = 0:LSLing2.Visible = 1:sling2.TransZ = -10
        Case 4:LSLing2.Visible = 0:LSLing.Visible = 1:sling2.TransZ = 0:LeftSlingShot.TimerEnabled = 0
    End Select
    LStep = LStep + 1
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

Sub a_Pins_Hit (idx)
	PlaySound "pinhit_low", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0, AudioFade(ActiveBall)
End Sub

Sub a_Targets_Hit (idx)
	PlaySound "target", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0, AudioFade(ActiveBall)
End Sub

Sub a_Metals_Thin_Hit (idx)
	PlaySound "metalhit_thin", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
End Sub

Sub a_Metals_Medium_Hit (idx)
	PlaySound "metalhit_medium", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
End Sub

Sub a_Metals2_Hit (idx)
	PlaySound "metalhit2", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
End Sub

Sub a_Gates_Hit (idx)
	PlaySound "gate4", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
End Sub

Sub a_Spinner_Spin
	PlaySound "fx_spinner",0,.25,0,0.25
End Sub

Sub a_Rubbers_Hit(idx)
 	dim finalspeed
  	finalspeed=SQR(activeball.velx * activeball.velx + activeball.vely * activeball.vely)
 	If finalspeed > 20 then
		PlaySound "fx_rubber2", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
	End if
	If finalspeed >= 6 AND finalspeed <= 20 then
 		RandomSoundRubber()
 	End If
End Sub

Sub RubberWheel_hit
 	dim finalspeed
  	finalspeed=SQR(activeball.velx * activeball.velx + activeball.vely * activeball.vely)
 	If finalspeed > 20 then
		PlaySound "fx_rubber2", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
	End if
	If finalspeed >= 6 AND finalspeed <= 20 then
 		RandomSoundRubber()
 	End If
End sub

Sub a_Posts_Hit(idx)
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



'*** Table specific scores and rules handling ***

sub bumper1_hit
	if tilt=false then
	playsound SoundFXDOF("jet1",111,DOFPulse,DOFContactors)
	DOF 112, DOFPulse
	if Light14.state=1 then addscore 100 else addscore 10
	end if
end sub

sub bumper2_hit
	if tilt=false then
	playsound SoundFXDOF("jet1",107,DOFPulse,DOFContactors)
	DOF 108, DOFPulse
	if Light11.state=1 then addscore 10 else addscore 1
	end if
end sub

sub bumper3_hit
	if tilt=false then
	playsound SoundFXDOF("jet1",113,DOFPulse,DOFContactors)
	DOF 114, DOFPulse
	if Light15.state=1 then addscore 10 else addscore 1
	end if
end sub

sub bumper4_hit
	if tilt=false then
	playsound SoundFXDOF("jet1",109,DOFPulse,DOFContactors)
	DOF 110, DOFPulse
	if Light12.state=1 then addscore 10 else addscore 1
	end if
end sub

sub bumper5_hit
	if tilt=false then
	playsound SoundFXDOF("jet1",115,DOFPulse,DOFContactors)
	DOF 116, DOFPulse
	if Light13.state=1 then addscore 10 else addscore 1
	end if
end sub

sub lout1_hit
	DOF 122, DOFPulse
	if tilt=false then
	if Light4.state=1 then addscore 300 Else addscore 100
	end if
end sub

sub lout2_hit
	DOF 125, DOFPulse
	if tilt=false then
	if Light2.state=1 then addscore 300 Else addscore 100
	end if
end sub

sub rout1_hit
	DOF 123, DOFPulse
	if tilt=false then
	if Light1.state=1 then addscore 300 Else addscore 100
	end if
end sub

sub rout2_hit
	DOF 124, DOFPulse
	if tilt=false then
	if Light5.state=1 then addscore 300 Else addscore 100
	end if
end sub

sub ml_hit
	DOF 117, DOFPulse
	if tilt=false then
	addscore 100
	if Light6.state=1 then Light3.state=1:If B2SOn Then Controller.B2SSetShootAgain 1
	end if
end sub

sub ul1_hit
	DOF 118, DOFPulse
	if tilt=false then
	if Light10.state=1 then addscore 300 Else addscore 100
	end if
end sub

sub ul2_hit
	DOF 119, DOFPulse
	if tilt=false then
	if Light9.state=1 then addscore 100 Else addscore 10
	end if
end sub

sub ul3_hit
	DOF 120, DOFPulse
	if tilt=false then
	if Light8.state=1 then addscore 300 Else addscore 100
	end if
end sub

sub ul4_hit
	DOF 121, DOFPulse
	if tilt=false then
	if Light7.state=1 then addscore 100 Else addscore 10
	end if
end sub

Sub LeftSlingShot1_Slingshot
	if tilt=false then
	addscore 1
	end if
End Sub

Sub LeftSlingShot2_Slingshot
	if tilt=false then
	if (wheelchange).enabled=false then
	playsound "reset2"
	wn=0
	wheelchange.enabled=true
	addscore 1
	end if
	end if
End Sub

Sub LeftSlingShot3_Slingshot
	if tilt=false then
	addscore 1
	end if
End Sub

Sub LeftSlingShot4_Slingshot
	if tilt=false then
	addscore 1
	if (wheelchange).enabled=false then
	playsound "reset2"
	wn=0
	wheelchange.enabled=true
	end if
	end if
End Sub

Sub RightSlingShot1_Slingshot
	if tilt=false then
	addscore 1
	end if
End Sub

Sub RightSlingShot2_Slingshot
	if tilt=false then
	addscore 1
	if (wheelchange).enabled=false then
	playsound "reset2"
	wn=0
	wheelchange.enabled=true
	end if
	end if
End Sub

Sub RightSlingShot3_Slingshot
	if tilt=false then
	addscore 1
	end if
End Sub

Sub RightSlingShot4_Slingshot
	if tilt=false then
	addscore 1
	if (wheelchange).enabled=false then
	playsound "reset2"
	wn=0
	wheelchange.enabled=true
	end if
	end if
End Sub


sub wheelcheck
    If wv=2 or wv=6 or wv=10 or wv=14 or wv=18 then post.isdropped=False
    If wv=0 or wv=4 or wv=8 or wv=12 or wv=16 then post.isdropped=true
	If wv=2 or wv=6 or wv=10 or wv=14 or wv=18 then postlight.state=1
	If wv=0 or wv=4 or wv=8 or wv=12 or wv=16 then postlight.state=0

'Wheel Green
    If wv=0 or wv=6 or wv=12 Then
    Light1.state=1
    Light2.state=0
    Light4.state=0
    Light5.state=0
    Light6.state=0
    light7.state=1
    light8.state=0
    Light9.state=0
    Light10.state=0
    Light11.state=0
    Light12.state=1
    Light13.state=1
    Light14.state=0
    Light15.state=0
    End If

'Wheel Yellow
    if wv=2 or wv=8 or wv=16 then
    Light1.state=0
    Light2.state=1
    Light4.state=0
    Light5.state=0
    Light6.state=0
    Light7.state=0
    Light8.state=0
    Light9.state=1
    Light10.state=0
    Light11.state=1
    Light12.state=0
    Light13.state=0
    Light14.state=0
    Light15.state=1
    end if

'Wheel White
    if wv=4 or wv=10 or wv=18 then
    Light1.state=0
    Light2.state=0
    Light4.state=0
    Light5.state=0
    Light6.state=0
    light7.state=0
    light8.state=1
    light9.state=0
    Light10.state=0
    Light11.state=0
    Light12.state=0
    Light13.state=0
    Light14.state=1
    Light15.state=0
    end if

'Wheel Red
    if wv=14 then
    light1.state=1
    light2.state=1
    light4.state=1
    light5.state=1
    Light6.state=1
    Light7.state=1
    Light8.state=1
    Light9.state=1
    Light10.state=1
    Light11.state=1
    Light12.state=1
    Light13.state=1
    Light14.state=1
    Light15.state=1
    end if

end sub

sub wheelchange_timer
    wn=wn+1
    playsound SoundFXDOF("clerker",133,DOFPulse,DOFGear)
    if tilt=false then
    wv=wv+1
    if wv>19 then wv=0
    if wv=0 then
    wheel(19).isdropped=true
    wheel(wv).isdropped=false
    end if
    if wv>0 then
    wheel(wv-1).isdropped=true
    wheel(wv).isdropped=false
    end if
    wheelcheck
    end if
    if wn=2 then wheelchange.enabled=false
end sub

sub ck_hit					'Set the Kicker Hit Height between 21 to 25. 21 = difficult, 25 = easy. 23 is a good average.
    if tilt=false then
    sbv=sbv+1
    if sbv<100 then sbreel.addvalue(1)
    if sbv>99 then sbv=99
	If B2SOn then Controller.B2SSetScorePlayer5 sbv
    addscore 300
    'playsound "motorshort1s"
    end if
    playsound "kickerkick"
    kicktimer.enabled=true
    if sbv=8 or sbv=18 or sbv=28 or sbv=38 or sbv=48 then
    credit=credit+1
	DOF 130, DOFOn
    playsound SoundFXDOF("knocke",129,DOFPulse,DOFKnocker)
	DOF 128, DOFPulse
    credittxt.text= credit
	If B2SOn Then
		Controller.B2SSetCredits Credit
	end if
    playsound "click"
    end if
end sub

sub kicktimer_timer
    ck.kick (int(rnd(1)*5)-165),15
	DOF 127, DOFPulse
	DOF 128, DOFPulse
    kicktimer.enabled=false
end sub


'*** High Score Handling ***

sub loadhs
    ' Based on Black's Highscore routines
	dim FileObj
	dim ScoreFile
	dim TextStr
	dim temp1
	dim temp2
	dim temp3
	dim temp4
	dim temp5
	dim temp6
	dim temp7
	dim temp8
	dim temp9
	Set FileObj=CreateObject("Scripting.FileSystemObject")
	If Not FileObj.FolderExists(UserDirectory) then
	Exit Sub
	End if
	If Not FileObj.FileExists(UserDirectory & "Cabaret.txt") then
	Exit Sub
	End if
	Set ScoreFile=FileObj.GetFile(UserDirectory & "Cabaret.txt")
	Set TextStr=ScoreFile.OpenAsTextStream(1,0)
	If (TextStr.AtEndOfStream=True) then
	Exit Sub
	End if
	temp1=TextStr.ReadLine
	temp2=textstr.readline
	temp3=textstr.readline
	temp4=Textstr.ReadLine
	temp5=Textstr.ReadLine
	temp6=Textstr.ReadLine
	temp7=Textstr.ReadLine
	temp8=Textstr.ReadLine
	temp9=Textstr.ReadLine
	TextStr.Close
	credit = CDbl(temp1)
	score(0) = CDbl(temp2)
	score(1) = CDbl(temp3)
	score(2) = CDbl(temp4)
	score(3) = CDbl(temp5)
	hisc = CDbl(temp6)
	matchnumb = CDbl(temp7)
	wv = CDbl(temp8)
	sbv = CDbl(temp9)
	Set ScoreFile=Nothing
	Set FileObj=Nothing
end sub

sub savehs
	' Based on Black's Highscore routines
	Dim FileObj
	Dim ScoreFile
	Set FileObj=CreateObject("Scripting.FileSystemObject")
	If Not FileObj.FolderExists(UserDirectory) then
	Exit Sub
	End if
	Set ScoreFile=FileObj.CreateTextFile(UserDirectory & "Cabaret.txt",True)
	scorefile.writeline credit
	ScoreFile.WriteLine score(0)
	ScoreFile.WriteLine score(1)
	ScoreFile.WriteLine score(2)
	ScoreFile.WriteLine score(3)
	scorefile.writeline hisc
	scorefile.writeline matchnumb
	scorefile.writeline wv
	scorefile.writeline sbv
	ScoreFile.Close
	Set ScoreFile=Nothing
	Set FileObj=Nothing
end sub

Sub Table1_Exit
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
  If Table1.VersionMinor > 3 OR Table1.VersionMajor > 10 Then
    PlaySound("fx_collide"), 0, Csng(velocity) ^2 / 200, Pan(ball1), 0, Pitch(ball1), 0, 0, AudioFade(ball1)
  Else
    PlaySound("fx_collide"), 0, Csng(velocity) ^2 / 200, Pan(ball1), 0, Pitch(ball1), 0, 0
  End if
End Sub


' Thalamus : Exit in a clean and proper way
Sub Table1_exit()
  Controller.Pause = False
  Controller.Stop
End Sub


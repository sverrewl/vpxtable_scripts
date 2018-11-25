Option Explicit
Randomize

' Thalamus 2018-07-20
' Added/Updated "Positional Sound Playback Functions" and "Supporting Ball & Sound Functions"
' Thalamus 2018-08-09 : Improved directional sounds
' !! NOTE : Table not verified yet !!

' Options
' Volume devided by - lower gets higher sound

Const VolDiv = 2000    ' Lower number, louder ballrolling/collition sound
Const VolCol = 10      ' Ball collition divider ( voldiv/volcol )

' The rest of the values are multipliers
'
'  .5 = lower volume
' 1.5 = higher volume

Const VolBump   = 2    ' Bumpers volume.
Const VolGates  = 1    ' Gates volume.
Const VolMetal  = 1    ' Metals volume.
Const VolRH     = 1    ' Rubber hits volume.
Const VolPo     = 1    ' Rubber posts volume.
Const VolPi     = 1    ' Rubber pins volume.
Const VolTarg   = 1    ' Targets volume.
Const VolKick   = 1    ' Kicker volume.
Const VolSpin   = 1.5  ' Spinners volume.
Const VolFlip   = 1    ' Flipper volume.


Const cGameName = "Champ_1974"
Const B2STableName="Champ_1974"

On Error Resume Next
ExecuteGlobal GetTextFile("controller.vbs")
ExecuteGlobal GetTextFile("core.vbs")
If Err Then MsgBox "Can't open controller.vbs"
On Error Goto 0

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
dim rv(5)
dim i				'Indexing variable
dim scn				'Single points score timer
dim scn1			'Single points end score timer
dim scn2			'Multiple points score timer
dim scn3			'Multiple points end score timer
Dim TargetsDown 	'Number of targets hit in the bank. Set this to 0 at start of your game
dim tv
dim ba
dim sa
dim batemp
dim satemp
dim bn(10)
dim mHole

Sub Table1_Init
	LoadEM
	set play(1)=up1
	set play(2)=up2
	set play(3)=up3
	set play(4)=up4
	set bn(1)=b1
	set bn(2)=b2
	set bn(3)=b3
	set bn(4)=b4
	set bn(5)=b5
	set bn(6)=b6
	set bn(7)=b7
	set bn(8)=b8
	set bn(9)=b9
	set bn(10)=b10
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
	replay1=102000
	replay2=126000
	replay3=150000
	replay4=174000
	loadhs
	if hisc="" then hisc=10000
	reel5.setvalue(hisc)
	if credit="" then credit=0
	credittxt.text=credit
	Targetsdown=0
	select case(matchnumb)
	case 0:
	m0.text="00"
	case 1:
	m1.text="10"
	case 2:
	m2.text="20"
	case 3:
	m3.text="30"
	case 4:
	m4.text="40"
	case 5:
	m5.text="50"
	case 6:
	m6.text="60"
	case 7:
	m7.text="70"
	case 8:
	m8.text="80"
	case 9:
	m9.text="90"
	end select
	for i=1 to 4
	currpl=i
	reel(i).setvalue(score(i))
	next
	currpl=0
	If B2SOn Then

		if matchnumb=0 then
			Controller.B2SSetMatch 100
		else
			Controller.B2SSetMatch matchnumb*10
		end if
		Controller.B2SSetScoreRolloverPlayer1 0
		Controller.B2SSetScoreRolloverPlayer2 0
		Controller.B2SSetScoreRolloverPlayer3 0
		Controller.B2SSetScoreRolloverPlayer4 0

		Controller.B2SSetTilt 1
		Controller.B2SSetCredits credit
		Controller.B2SSetGameOver 1
	End If
	for i=1 to 4
		If B2SOn Then
			Controller.B2SSetScorePlayer i, score(i)
		End If
	next
	ba=1
	sa=1
	spinneradvance
	KickerYellow.CreateBall
	KickerYellow.kick 180,3
	KickerGreen.CreateBall
	KickerGreen.kick 180,3

End Sub

Sub Table1_exit()
	If B2SOn Then Controller.Stop
end sub


Sub Table1_KeyDown(ByVal keycode)

	If keycode = PlungerKey Then
	Plunger.PullBack
	End If

    if keycode = 6 then
	playsoundAtVol "coin3", drain, 1
	coindelay.enabled=true
	end if

    if keycode = 5 then
	playsoundAtVol "coin3", drain, 1
	coindelay1.enabled=true
	end if

	if keycode = 2 and credit>0 and state=false and playno=0 then
	credit=credit-1
	credittxt.text=credit
    eg=0
    playno=1
    playno1.state=1
    currpl=1
    play(currpl).state=1
		If B2SOn Then
			Controller.B2SSetTilt 0
			Controller.B2SSetGameOver 0
			Controller.B2SSetMatch 0
			Controller.B2SSetCredits credit

			Controller.B2SSetCanPlay 1
			Controller.B2SSetPlayerUp 1

			Controller.B2SSetScoreRolloverPlayer1 0
			Controller.B2SSetScoreRolloverPlayer2 0
			Controller.B2SSetScoreRolloverPlayer3 0
			Controller.B2SSetScoreRolloverPlayer4 0
		End If
    playsound "click"
    playsound "initialize"
    rst=0
    bip=1
    resettimer.enabled=true
    end if

    if keycode = 2 and credit>0 and state=true and playno>0 and playno<4 and bip<2 then
    credit=credit-1
    credittxt.text=credit
    playno=playno+1
    if playno=2 then playno2.state=1
    if playno=3 then playno3.state=1
    if playno=4 then playno4.state=1
    playsound "click"
		If B2SOn Then
			Controller.B2SSetCanPlay playno

		End If
    end if

    if state=true and tilt=false then

	If keycode = LeftFlipperKey Then
		LeftFlipper.RotateToEnd
		PlaySoundAtVol "FlipperUp", LeftFlipper, VolFlip
	End If

	If keycode = RightFlipperKey Then
		RightFlipper.RotateToEnd
        PlaySoundAtVol "FlipperUp", RightFlipper, VolFlip
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
	end if

End Sub

Sub Table1_KeyUp(ByVal keycode)

	If keycode = PlungerKey Then
		Plunger.Fire
	End If

	If keycode = LeftFlipperKey Then
		LeftFlipper.RotateToStart
		stopsound "buzz"
		if state=true and tilt=false then PlaySoundAtVol "FlipperDown", LeftFlipper, VolFlip
	End If

	If keycode = RightFlipperKey Then
		RightFlipper.RotateToStart
		stopsound "buzz"
        if state=true and tilt=false then PlaySoundAtVol "FlipperDown", RightFlipper, VolFlip
	End If

End Sub


sub newgame
	state=true
	eg=0
	ba=1
	sa=1
	ton1.text=" "
	ton2.text=" "
	ton3.text=" "
	ton4.text=" "
	LightReset
	bumperlight1.state=0
	bumperlight2.state=0
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
	bip=1
	nb.CreateBall
	nb.kick 90,6
     Set mHole = New cvpmMagnet
     With mHole
         .InitMagnet Umagnet, 4
         .GrabCenter = 0
         .MagnetOn = 1
         .CreateEvents "mHole"
     End With
end sub

sub resettimer_timer
	rst=rst+1
	reel1.resettozero
	reel2.resettozero
	reel3.resettozero
	reel4.resettozero
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
	If B2SOn then
		Controller.B2SSetCredits credit
	end if
	credittxt.text=credit
    coindelay.enabled=false
end sub

sub coindelay1_timer
	playsound "click"
	credit=credit+1
	If B2SOn then
		Controller.B2SSetCredits credit
	end if
	credittxt.text=credit
    coindelay1.enabled=false
end sub


'*** General points scoring ***

sub addscore(points)

    if tilt=false then
	bell=0
    if points=100 then matchnumb=matchnumb+1
    if matchnumb=10 then matchnumb=0

	if points = 10 or points = 100 or points = 1000 then scn=1

	if points = 500 then scn2=5
	if points = 5000 then scn2=5
	if points = 4000 then scn2=4
	if points = 3000 then scn2=3
	if points = 2000 then scn2=2

    if points = 500 then
    reel(currpl).addvalue(500)
    playsound "motorshort1s"
    scn3=5
    bell=100
    end if

    if points = 5000 then
    reel(currpl).addvalue(5000)
    playsound "motorshort1s"
    scn3=5
    bell=1000
    end if

    if points = 4000 then
    reel(currpl).addvalue(4000)
    playsound "motorshort1s"
    scn3=4
    bell=1000
    end if

    if points = 3000 then
    reel(currpl).addvalue(3000)
    playsound "motorshort1s"
    scn3=3
    bell=1000
    end if

    if points = 2000 then
    reel(currpl).addvalue(2000)
    playsound "motorshort1s"
    scn3=2
    bell=1000
    end if

    if points = 1000 then
    reel(currpl).addvalue(1000)
    bell=1000
    scn1=1
    end if

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

    if points = 10 or points = 100 or points = 1000 then scn1=0 : scntimer.enabled=true : end if
    if points = 5000 then scn3=0 : scntimer1.enabled=true : end if
    if points = 4000 then scn3=0 : scntimer1.enabled=true : end if
    if points = 3000 then scn3=0 : scntimer1.enabled=true : end if
    if points = 2000 then scn3=0 : scntimer1.enabled=true : end if
    if points = 500 then scn3=0 : scntimer1.enabled=true : end if


	score(currpl)=score(currpl)+points
	truesc(currpl)=truesc(currpl)+points
	end if

    if score(currpl)>199990 then
    score(currpl)=score(currpl)-100000
    rep(currpl)=0
    end if

    if score(currpl)=>replay1 and rep(currpl)=0 then
    credit=credit+1
    playsound "knocke"
    credittxt.text=credit
	If B2SOn then
		Controller.B2SSetCredits credit
	end if
    rep(currpl)=1
    playsound "click"
    end if

    if score(currpl)=>replay2 and rep(currpl)=1 then
    credit=credit+1
    playsound "knocke"
    credittxt.text=credit
	If B2SOn then
		Controller.B2SSetCredits credit
	end if
    rep(currpl)=2
    playsound "click"
    end if

	if score(currpl)=>replay3 and rep(currpl)=2 then
	credit=credit+1
	playsound "knocke"
	credittxt.text=credit
	If B2SOn then
		Controller.B2SSetCredits credit
	end if
	rep(currpl)=3
	playsound "click"
	end if

	if score(currpl)=>replay4 and rep(currpl)=3 then
	credit=credit+1
	playsound "knocke"
	credittxt.text=credit
	If B2SOn then
		Controller.B2SSetCredits credit
	end if
	rep(currpl)=4
	playsound "click"
	end if
	if score(1)>=100000 then ton1.text="100,000"
	if score(2)>=100000 then ton2.text="100,000"
	if score(3)>=100000 then ton3.text="100,000"
	if score(4)>=100000 then ton4.text="100,000"
	If B2SOn Then
		Controller.B2SSetScorePlayer currpl, score(currpl)
		if score(1)>=100000 then Controller.B2SSetScoreRolloverPlayer1 1
		if score(2)>=100000 then Controller.B2SSetScoreRolloverPlayer2 1
		if score(3)>=100000 then Controller.B2SSetScoreRolloverPlayer3 1
		if score(4)>=100000 then Controller.B2SSetScoreRolloverPlayer4 1
	End If
end sub

sub scntimer_timer
    scn1=scn1 + 1
    if bell=1000 then playsound "bell1000",0, 0.45, 0, 0
    if bell=100 then playsound "bell100",0, 0.3, 0, 0
    if bell=10 then playsound "bell10",0, 0.15, 0, 0
    if scn1=scn then scntimer.enabled=false
end sub

sub scntimer1_timer
	scn3=scn3 + 1
    if bell=1000 then playsound "bell1000",0, 0.45, 0, 0
    if bell=100 then playsound "bell100",0, 0.3, 0, 0
    if scn3=scn2 then scntimer1.enabled=false
end sub


'*** Ball handling ***

sub ballhome_hit
    ballrelenabled=1
end sub

sub ballrelease_hit
	if ballrelenabled=1 then playsoundAtVol "launchball", ballrelease, 1: ballrelenabled=0: end if
end sub

Sub Drain_Hit()
	playsoundAtVol "drainshorter", drain, 1
	Drain.DestroyBall
	bonuscount.enabled=true
End Sub

sub nextball
	LightReset
	ba=1
	bumperlight1.state=0
	bumperlight2.state=0
	if tilt=true then
	tilt=false
	tilttext.text=" "
	tiltseq.stopplay
	end if
	if (Light17.state)=1 then
	playsound "kickerkick"
	If B2SOn then
		Controller.B2SSetShootAgain 0
	end if
	ballreltimer.enabled=true
	light17.state=0
	else
	currpl=currpl+1
	end if
	if currpl>playno then
	bip=bip+1
	if bip=5 then db.state=1 else db.state=0
	if bip>5 then
	playsound "motorleer"
	eg=1
	ballreltimer.enabled=true
	else
	if state=true and tilt=false then
	If B2SOn then
		Controller.B2SSetBallInPlay bip
		Controller.B2SSetPlayerUp currpl
	end if
	play(currpl-1).state=0
	currpl=1
	If B2SOn then
		Controller.B2SSetBallInPlay bip
		Controller.B2SSetPlayerUp currpl
	end if
	play(currpl).state=1
	playsound "kickerkick"
	ballreltimer.enabled=true
	end if
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
	If B2SOn then
		Controller.B2SSetBallInPlay bip
		Controller.B2SSetPlayerUp currpl
	end if
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
	If B2SOn then
		Controller.B2SSetGameOver 1
		Controller.B2SSetPlayerUp 0
		Controller.B2SSetBallInPlay 0
		Controller.B2SSetCanPlay 0
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
    m0.text="00"
    case 1:
    m1.text="10"
    case 2:
    m2.text="20"
    case 3:
    m3.text="30"
    case 4:
    m4.text="40"
    case 5:
    m5.text="50"
    case 6:
    m6.text="60"
    case 7:
    m7.text="70"
    case 8:
    m8.text="80"
    case 9:
    m9.text="90"
    end select
	If B2SOn then
		If matchnumb = 0 Then
			Controller.B2SSetMatch 100
		Else
			Controller.B2SSetMatch matchnumb*10
		End If
	End if
    for i=1 to playno
    if matchnumb*10=(score(i) mod 100) then
    credit=credit+1
	If B2SOn then
		Controller.B2SSetCredits credit
	end if
    playsound "knocke"
    credittxt.text= credit
    end if
    next
end sub

Sub CheckTilt
	If Tilttimer.Enabled = True Then
	TiltSens = TiltSens + 1
	if TiltSens = 3 Then
	Tilt = True
	tilttext.text="TILT"
	tiltsens = 0
	playsound "tilt"
	If B2Son then
		Controller.B2SSetTilt 1
	end if
	turnoff
	End If
	Else
	TiltSens = 0
	Tilttimer.Enabled = True
	End If
End Sub

Sub Tilttimer_Timer()
	Tilttimer.Enabled = False
End Sub

sub turnoff
	tiltseq.play seqalloff
end sub

Sub Gate1_hit()
	PlaysoundAtVol "gate", gate1, VolGates
End Sub

'**********Sling Shot Animations
' Rstep and Lstep  are the variables that increment the animation
'****************
Dim RStep, Lstep

Sub RightSlingShot_Slingshot
	If Tilt=False then addscore 10
    PlaySoundAtVol "right_slingshot", sling1, 1
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
	If Tilt=False then addscore 10
    PlaySoundAtVol "left_slingshot", sling2, 1
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
	PlaySound "pinhit_low", 0, Vol(ActiveBall)*VolPi, AudioPan(ActiveBall), 0, Pitch(ActiveBall), 0, 0, AudioFade(ActiveBall)
End Sub

Sub a_Targets_Hit (idx)
	PlaySound "target", 0, Vol(ActiveBall)*VolTarg, AudioPan(ActiveBall), 0, Pitch(ActiveBall), 0, 0, AudioFade(ActiveBall)
End Sub

Sub a_Metals_Thin_Hit (idx)
	PlaySound "metalhit_thin", 0, Vol(ActiveBall)*VolMetal, AudioPan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
End Sub

Sub a_Metals_Medium_Hit (idx)
	PlaySound "metalhit_medium", 0, Vol(ActiveBall)*VolMetal, AudioPan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
End Sub

Sub a_Metals2_Hit (idx)
	PlaySound "metalhit2", 0, Vol(ActiveBall)*VolMetal, AudioPan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
End Sub

Sub a_Gates_Hit (idx)
	PlaySound "gate4", 0, Vol(ActiveBall)*VolGates, AudioPan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
End Sub

Sub a_Spinner_Spin
	PlaySoundAtVol "fx_spinner", a_Spinner, VolSpin
End Sub

Sub a_Rubbers_Hit(idx)
 	dim finalspeed
  	finalspeed=SQR(activeball.velx * activeball.velx + activeball.vely * activeball.vely)
 	If finalspeed > 20 then
		PlaySound "fx_rubber2", 0, Vol(ActiveBall)*VolRH, Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
	End if
	If finalspeed >= 6 AND finalspeed <= 20 then
 		RandomSoundRubber()
 	End If
End Sub

Sub RubberWheel_hit
 	dim finalspeed
  	finalspeed=SQR(activeball.velx * activeball.velx + activeball.vely * activeball.vely)
 	If finalspeed > 20 then
		PlaySound "fx_rubber2", 0, Vol(ActiveBall)*VolRH, Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
	End if
	If finalspeed >= 6 AND finalspeed <= 20 then
 		RandomSoundRubber()
 	End If
End sub

Sub a_Posts_Hit(idx)
 	dim finalspeed
  	finalspeed=SQR(activeball.velx * activeball.velx + activeball.vely * activeball.vely)
 	If finalspeed > 16 then
		PlaySound "fx_rubber2", 0, Vol(ActiveBall)*VolPo, Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
	End if
	If finalspeed >= 6 AND finalspeed <= 16 then
 		RandomSoundRubber()
 	End If
End Sub

Sub RandomSoundRubber()
	Select Case Int(Rnd*3)+1
		Case 1 : PlaySound "rubber_hit_1", 0, Vol(ActiveBall)*VolRH, Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
		Case 2 : PlaySound "rubber_hit_2", 0, Vol(ActiveBall)*VolRH, Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
		Case 3 : PlaySound "rubber_hit_3", 0, Vol(ActiveBall)*VolRH, Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
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
		Case 1 : PlaySound "flip_hit_1", 0, Vol(ActiveBall)*VolRH, Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
		Case 2 : PlaySound "flip_hit_2", 0, Vol(ActiveBall)*VolRH, Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
		Case 3 : PlaySound "flip_hit_3", 0, Vol(ActiveBall)*VolRH, Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
	End Select
End Sub


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
	Set FileObj=CreateObject("Scripting.FileSystemObject")
	If Not FileObj.FolderExists(UserDirectory) then
	Exit Sub
	End if
	If Not FileObj.FileExists(UserDirectory & "Champ.txt") then
	Exit Sub
	End if
	Set ScoreFile=FileObj.GetFile(UserDirectory & "Champ.txt")
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
	TextStr.Close
	credit = CDbl(temp1)
	score(0) = CDbl(temp2)
	score(1) = CDbl(temp3)
	score(2) = CDbl(temp4)
	score(3) = CDbl(temp5)
	hisc = CDbl(temp6)
	matchnumb = CDbl(temp7)
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
	Set ScoreFile=FileObj.CreateTextFile(UserDirectory & "Champ.txt",True)
	scorefile.writeline credit
	ScoreFile.WriteLine score(0)
	ScoreFile.WriteLine score(1)
	ScoreFile.WriteLine score(2)
	ScoreFile.WriteLine score(3)
	scorefile.writeline hisc
	scorefile.writeline matchnumb
	ScoreFile.Close
	Set ScoreFile=Nothing
	Set FileObj=Nothing
end sub


'*** Tabel specific score handling ***

Sub LightReset
	BumperLight1.state=0
	BumperLight2.state=0
	Light11.state=0
	Light12.state=0
	Light13.state=0
	Light14.state=0
	Light15.state=0
	Light16.state=0
End Sub

Sub Bumper1_Hit()
	PlaySoundAtVol "jet1", Bumper1, VolBump
    If bumperLight1.state=1 then addscore 100 else addscore 10
    addscore 100
End Sub

Sub Bumper2_Hit()
	PlaySoundAtVol "jet1", Bumper2, VolBump
    If bumperLight2.state=1 then addscore 100 else addscore 10
End Sub


Sub Kicker_Hit()
 With mHole
         .InitMagnet Umagnet, 0
         .GrabCenter = 0
         .MagnetOn = 1
         .CreateEvents "mHole"
     End With
	Addscore 3000
	Light13.state=1
	Light14.state=1
	BumperLight1.state=1
	BumperLight2.state=1
	KickTimer.enabled=True
End Sub


Sub KickTimer_Timer
    KickTimer.enabled=False
    Kicker.Kick (int(rnd(1)*4)+176),(int(rnd(1)*5)+14)
    PlaysoundAtVol "HoleKick", Kicker, 1
 With mHole
         .InitMagnet Umagnet, 4
         .GrabCenter = 0
         .MagnetOn = 1
         .CreateEvents "mHole"
     End With
End Sub


Sub Spinner1_Spin()
	spinneradvance
	addscore 100
End Sub




' *** Bonus handling ***

sub bonusadvance
	if tilt=false then
	ba=ba+1
	if ba>15 then ba=15		'Set the maximim bonus limit
	if ba>14 then light11.state=1: light12.state=1: end if
	If ba=8 then light15.state=1: light16.state=0: end if
	If ba=9 then light15.state=0: light16.state=1: end if
	If ba=10 then light15.state=1: light16.state=0: end if
	If ba=11 then light15.state=0: light16.state=1: end if
	If ba=12 then light15.state=1: light16.state=0: end if
	If ba=13 then light15.state=0: light16.state=1: end if
	If ba=14 then light15.state=1: light16.state=0: end if
	If ba=15 then light15.state=0: light16.state=0: end if
    batemp=ba
    if batemp>9 then
    b10.state=1
    batemp=batemp-10
    end if
    select case (batemp)
    case 0:
    b9.state=0
    case 1:
    b1.state=1
    case 2:
    b1.state=0
    b2.state=1
    case 3:
    b2.state=0
    b3.state=1
    case 4:
    b3.state=0
    b4.state=1
    case 5:
    b4.state=0
    b5.state=1
    case 6:
    b5.state=0
    b6.state=1
    case 7:
    b6.state=0
    b7.state=1
    case 8:
    b7.state=0
    b8.state=1
    case 9:
    b8.state=0
    b9.state=1
    case 10:
    b9.state=0
    b10.state=1
    case 11:
    b10.state=0
    b1.state=1
    case 12:
    b1.state=0
    b2.state=1
    case 13:
    b2.state=0
    b3.state=1
    case 14:
    b3.state=0
    b4.state=1
    case 15:
    b4.state=0
    b5.state=1
    case 16:
    b5.state=0
    b6.state=1
    case 17:
    b6.state=0
    b7.state=1
    case 18:
    b7.state=0
    b8.state=1
    case 19:
    b8.state=0
    b9.state=1
    end select
    playsound "click"
    end if
end sub

sub bonuscount_timer
	playsound "motorshorter"
	batemp=ba
	if batemp<10 then b10.state=0
	select case (batemp)
	case 0:
	b1.state=0
	case 1:
	b2.state=0
	b1.state=1
	case 2:
	b3.state=0
	b2.state=1
	case 3:
	b4.state=0
	b3.state=1
	case 4:
	b5.state=0
	b4.state=1
	case 5:
	b6.state=0
	b5.state=1
	case 6:
	b7.state=0
	b6.state=1
	case 7:
	b8.state=0
	b7.state=1
	case 8:
	b9.state=0
	b8.state=1
	case 9:
	b9.state=1
	case 10:
	b1.state=0
	case 11:
	b2.state=0
	b1.state=1
	case 12:
	b3.state=0
	b2.state=1
	case 13:
	b4.state=0
	b3.state=1
	case 14:
	b5.state=0
	b4.state=1
	case 15:
	b6.state=0
	b5.state=1
	case 16:
	b7.state=0
	b6.state=1
	case 17:
	b8.state=0
	b7.state=1
	case 18:
	b9.state=0
	b8.state=1
	case 19:
	b9.state=1
	end select
	if (db.state)=0 then addscore 1000
    if (db.state)=1 then addscore 2000
    ba=ba-1
    if ba<=0 then
    nextball
    bonuscount.enabled=false
    end if
end sub


Sub SpinnerAdvance
	sa=sa+1
	if sa>10 then sa=1
	satemp=sa
	select case (satemp)
	case 0:
	s1.state=0
	s10.state=1
	case 1:
	s1.state=0
	s2.state=1
	case 2:
	s2.state=0
	s3.state=1
	case 3:
	s3.state=0
	s4.state=1
	case 4:
	s4.state=0
	s5.state=1
	case 5:
	s5.state=0
	s6.state=1
	case 6:
	s6.state=0
	s7.state=1
	case 7:
	s7.state=0
	s8.state=1
	case 8:
	s8.state=0
	s9.state=1
	case 9:
	s9.state=0
	s10.state=1
	case 10:
	s10.state=0
	s1.state=1
	bonusadvance
	end select
End Sub


Sub BallReleaseGate_Hit()
	b1.state=1
	tv=0
End Sub



Sub Trigger1_Hit()
	addscore 500
	BumperLight1.state=1
	Light13.state=1
End Sub

Sub Trigger2_Hit()
	addscore 500
	BumperLight2.state=1
	Light14.state=1
End Sub

Sub Trigger3_Hit()
	addscore 1000
	If light13.state=1 then bonusadvance
End Sub

Sub Trigger4_Hit()
	addscore 1000
	If light13.state=1 then bonusadvance
End Sub

Sub Trigger5_Hit()
	addscore 1000
	If light14.state=1 then bonusadvance
End Sub

Sub Trigger6_Hit()
	addscore 1000
	If light14.state=1 then bonusadvance
End Sub

Sub Trigger12_Hit()
	addscore 1000
End Sub

Sub Trigger13_Hit()
	bonusadvance
	addscore 1000
End Sub

Sub Trigger14_Hit()
	bonusadvance
	addscore 1000
End Sub

Sub Trigger15_Hit()
	addscore 1000
End Sub


Sub Bumper3_Hit()
	addscore 500
	BumperLight1.state=1
	Light13.state=1
End Sub

Sub Bumper4_Hit()
	addscore 500
	BumperLight2.state=1
	Light14.state=1
End Sub

Sub sw1_Hit()
	Addscore 1000
	If light15.state=1 then
		light17.state=1
		if B2SOn then
			Controller.B2SSetShootAgain 1
		end if
	end if
End Sub

Sub sw2_Hit()
	Addscore 1000
	If light16.state=1 then
		light17.state=1
		if B2SOn then
			Controller.B2SSetShootAgain 1
		end if
	end if
End Sub

Sub sw3_Hit()
	Addscore 1000
	If light11.state=1 then
		credit=credit+1
		playsound "knocke"
		if B2SOn then
			Controller.B2SSetCredits credit
		end if
	end if
	credittxt.text=credit
End Sub

Sub sw4_Hit()
	Addscore 1000
	If light12.state=1 then
		credit=credit+1
		playsound "knocke"
		if B2SOn then
			Controller.B2SSetCredits credit
		end if
	end if
    credittxt.text=credit
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
    PlaySound("fx_collide"), 0, Csng(velocity) ^2 / (VolDiv/VolCol), Pan(ball1), 0, Pitch(ball1), 0, 0, AudioFade(ball1)
End Sub


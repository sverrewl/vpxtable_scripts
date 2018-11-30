Option Explicit
Randomize

' Thalamus 2018-07-23
' Added/Updated "Positional Sound Playback Functions" and "Supporting Ball & Sound Functions"
' Thalamus 2018-11-01 : Improved directional sounds
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
Const VolSpin   = 1.5  ' Spinners volume.
Const VolFlip   = 1    ' Flipper volume.


Const cGameName = "JollyRoger_1967"

dim score(4)
dim truesc(4)
dim reel(4)
dim creel
dim ballrelenabled
dim state
dim credit
dim eg
dim currpl
dim playno
dim plno(4)
dim play(4)
dim rst
dim ballinplay
dim tilt
dim tiltsens
dim rep(4)
dim plm(4)
dim matchnumb
dim bell
dim points
dim tempscore
dim replay1
dim replay2
dim replay3
dim replay4
dim hisc
dim scn
dim scn1
dim i
dim rv
dim objekt

On Error Resume Next
ExecuteGlobal GetTextFile("controller.vbs")
If Err Then MsgBox "Can't open controller.vbs"
On Error Goto 0

sub JollyRoger_init
	LoadEM
    set play(1)=plno1
    set play(2)=plno2
    set play(3)=plno3
    set play(4)=plno4
    set reel(1)=reel1
    set reel(2)=reel2
    set reel(3)=reel3
    set reel(4)=reel4
    Bumperlight1.state=0
    Bumperlight2.state=1
    Bumperlight3.state=1
    Bumperlight4.state=0
    upper1l.state=0
    upper2l.state=1
    upper3l.state=0
    upper4l.state=1
    replay1=3200
    replay2=4300
    replay3=5400
    replay4=6500
    loadhs
    if creel="" then creel=30
	creelP.rotx=((creel/10)-1)*36
    if hisc="" then hisc=1000
    hisctxt.text=hisc
    if credit="" then credit=0
	CreditReel.setvalue credit
    if matchnumb="" then matchnumb=int(rnd(1)*9)
	MatchReel.setvalue matchnumb
    for i=1 to 4
		currpl=i
		reel(i).setvalue(score(i))
    next
	GameOverReel.setvalue 1
    currpl=0

	If B2SOn Then
		for each objekt in backdropstuff: objekt.visible=0: next
		if matchnumb=0 then
			Controller.B2SSetMatch 10
		else
			Controller.B2SSetMatch matchnumb
		end if
		Controller.B2SSetScoreRolloverPlayer1 0
		Controller.B2SSetScoreRolloverPlayer2 0
		Controller.B2SSetScoreRolloverPlayer3 0
		Controller.B2SSetScoreRolloverPlayer4 0

		Controller.B2SSetTilt 1
		Controller.B2SSetCredits Credit
		Controller.B2SSetGameOver 1
		Controller.B2SSetData 81,0
		Controller.B2SSetData 82,0
		Controller.B2SSetData 83,0
		Controller.B2SSetData 84,0
		for i=1 to 4
			Controller.B2SSetScorePlayer i, 0
		next
	End If

	If credit > 0 Then DOF 130, DOFOn

end sub


Sub JollyRoger_KeyDown(ByVal keycode)

	If keycode = PlungerKey Then
	Plunger.PullBack
	End If

	if keycode = AddCreditKey then
    credit=credit+4
	DOF 130, DOFOn
	playsoundAtVol "coin3", drain, 1
	coindelay.enabled=true
	If B2SOn Then
		Controller.B2SSetCredits Credit
	end if
	end if

	if keycode = 5 then
	playsoundAtVol "coin3", drain, 1
	coindelay.enabled=true
	If B2SOn Then
		Controller.B2SSetCredits Credit
	end if
	end if

	if keycode = StartGameKey and credit>0 and state=false and playno=0 then
	credit=credit-1
	If credit < 1 Then DOF 130, DOFOff
	eg=0
	CreditReel.setvalue credit
	playno=1
	CanPlayReel.setvalue 1
	currpl=1
	play(currpl).state=1
    playsound "click"
    playsound "initialize"
    rst=0
    ballinplay=1
    resettimer.enabled=true
		If B2SOn Then
			Controller.B2SSetTilt 0
			Controller.B2SSetGameOver 0
			Controller.B2SSetMatch 0
			Controller.B2SSetCredits Credit
			Controller.B2SSetScore 3,hisc
			Controller.B2SSetCanPlay 1
			Controller.B2SSetPlayerUp 1
			Controller.B2SSetData 81,1
			Controller.B2SSetData 82,0
			Controller.B2SSetData 83,0
			Controller.B2SSetData 84,0
			Controller.B2SSetBallInPlay BallInPlay
			Controller.B2SSetScoreRolloverPlayer1 0

		End If
    end if

    if keycode = StartGameKey and credit>0 and state=true and playno>0 and playno<4 and ballinplay<2 then
		credit=credit-1
		If credit < 1 Then DOF 130, DOFOff
		CreditReel.setvalue credit
		playno=playno+1
		CanPlayReel.setvalue playno
		playsound "click"
		If B2SOn Then
			Controller.B2SSetCredits Credit
			Controller.B2SSetCanPlay playno
		end if
    end if

    if state=true and tilt=false then

		If keycode = LeftFlipperKey Then
		LeftFlipper.RotateToEnd
		PlaySoundAtVol SoundFXDOF("FlipperUp",101,DOFOn,DOFFlippers), LeftFlipper, VolFlip
		End If

		If keycode = RightFlipperKey Then
		RightFlipper.RotateToEnd
		PlaySoundAtVol SoundFXDOF("FlipperUp",102,DOFOn,DOFFlippers), RightFlipper, VolFlip
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
			gametilted
		End If

	End If

End Sub

sub flippertimer_timer()
	LFlip.RotY = LeftFlipper.CurrentAngle-90
	RFlip.RotY = RightFlipper.CurrentAngle+90
end sub

Sub JollyRoger_KeyUp(ByVal keycode)

	If keycode = PlungerKey Then
	Plunger.Fire
	playsoundAtVol "plungerrelease", plunger, 1
	End If

	If keycode = LeftFlipperKey Then
	LeftFlipper.RotateToStart
	stopsound "buzz"
	if state=true and tilt=false then PlaySoundAtVol SoundFXDOF("FlipperDown",101,DOFOff,DOFFlippers), LeftFlipper, VolFlip
	End If

	If keycode = RightFlipperKey Then
	RightFlipper.RotateToStart
	stopsound "buzz"
	if state=true and tilt=false then PlaySoundAtVol SoundFXDOF("FlipperDown",102,DOFOff,DOFFlippers), RightFlipper, VolFlip
	End If

End Sub

sub coindelay_timer
	playsound "click"
	credit=credit+1
	DOF 130, DOFOn
	if credit>15 then credit=15
	CreditReel.setvalue credit
    coindelay.enabled=false
	If B2SOn Then
		Controller.B2SSetCredits Credit
	end if
end sub

sub resettimer_timer
    rst=rst+1
    reel1.resettozero
    reel2.resettozero
    reel3.resettozero
    reel4.resettozero
	If B2SOn Then
		Controller.B2SSetScorePlayer1 0
		Controller.B2SSetScorePlayer2 0
		Controller.B2SSetScorePlayer3 0
		Controller.B2SSetScorePlayer4 0
	end if
    if rst=14 then
    playsound "kickerkick" ' TODO
    end if
    if rst=18 then
    newgame
    resettimer.enabled=false
    end if
end sub


sub addscore(points)
    if tilt=false then
		bell=0
		if points = 1 or points = 10 or points = 100 then scn=1
		if points=10 or points=1 then
			matchnumb=matchnumb+1
		end if

		if matchnumb=10 then matchnumb=0

		if score(currpl)>9999 then
			score(currpl)=score(currpl)-10000
			rep(currpl)=0
		end if

		if score(currpl)=>replay1 and rep(currpl)=0 then
			credit=credit+1
			DOF 130, DOFOn
			if credit>15 then credit=15
			playsound SoundFXDOF("knocke",128, DOFPulse,DOFKnocker):DOF 129, DOFPulse
			CreditReel.setvalue credit
			If B2SOn Then
				Controller.B2SSetCredits Credit
			end if
			rep(currpl)=1
			playsound "click"
		end if

    if score(currpl)=>replay2 and rep(currpl)=1 then
		credit=credit+1
		DOF 130, DOFOn
		if credit>15 then credit=15
		PlaySound SoundFXDOF("knocke",128,DOFPulse,DOFKnocker):DOF 129, DOFPulse
		CreditReel.setvalue credit
		If B2SOn Then
			Controller.B2SSetCredits Credit
		end if
		rep(currpl)=2
		playsound "click"
    end if

    if score(currpl)=>replay3 and rep(currpl)=2 then
		credit=credit+1
		DOF 130, DOFOn
		if credit>15 then credit=15
		PlaySound SoundFXDOF("knocke",128, DOFPulse,DOFKnocker):DOF 129, DOFPulse
		CreditReel.setvalue credit
		If B2SOn Then
			Controller.B2SSetCredits Credit
		end if
		rep(currpl)=3
		playsound "click"
    end if

    if score(currpl)=>replay4 and rep(currpl)=3 then
		credit=credit+1
		DOF 130, DOFOn
		if credit>15 then credit=15
		PlaySound SoundFXDOF("knocke",128, DOFPulse,DOFKnocker):DOF 129, DOFPulse
		CreditReel.setvalue credit
		If B2SOn Then
			Controller.B2SSetCredits Credit
		end if
		rep(currpl)=4
		playsound "click"
    end if

    if points = 100 then
		reel(currpl).addvalue(100)
		bell=100
    end if

    if points = 10 then
		reel(currpl).addvalue(10)
		bell=10
    end if

    if points = 1 then
		reel(currpl).addvalue(1)
		bell=1
    end if

    if points = 500 then
		reel(currpl).addvalue(500)
		scn=5
		bell=100
    end if

    if points = 400 then
		reel(currpl).addvalue(400)
		scn=4
		bell=100
    end if

    if points = 300 then
		reel(currpl).addvalue(300)
		scn=3
		bell=100
    end if

    if points = 200 then
		reel(currpl).addvalue(200)
		scn=2
		bell=100
    end if

    if points = 50 then
		reel(currpl).addvalue(50)
		scn=5
		bell=10
    end if

    if points = 40 then
		bell=10
		reel(currpl).addvalue(40)
		scn=4
		bell=10
    end if

    if points = 30 then
		reel(currpl).addvalue(30)
		scn=3
		bell=10
    end if

    if points = 20 then
		reel(currpl).addvalue(20)
		scn=2
		bell=10
    end if

    scn1=0
    scntimer.enabled=true
    score(currpl)=score(currpl)+points
    truesc(currpl)=truesc(currpl)+points
	If B2SOn Then
		Controller.B2SSetScore currpl,score(currpl)
	end if
  end if
End Sub

sub upperl1_hit
	DOF 115, DOFPulse
    if (upper1l.state)=1 then
    addscore 100
    else
    addscore 10
    end if
end sub

sub upperl2_hit
	DOF 116, DOFPulse
    if (upper2l.state)=1 then
    addscore 100
    else
    addscore 10
    end if
end sub

sub upperl3_hit
	DOF 117, DOFPulse
    if (upper3l.state)=1 then
    addscore 100
    else
    addscore 10
    end if
end sub

sub upperl4_hit
	DOF 118, DOFPulse
    if (upper4l.state)=1 then
    addscore 100
    else
    addscore 10
    end if
end sub

sub upperll1_hit
    addscore 300
end sub

sub upperr2_hit
	DOF 124, DOFPulse
    if (upperr2l.state)=1 then
    addscore 100
    else
    addscore 10
    end if
end sub

sub scntimer_timer
    scn1=scn1 + 1
    if bell=1 then playsound SoundFXDOF("bell1",141,DOFPulse,DOFChimes),0, 0.15, 0, 0
    if bell=10 then playsound SoundFXDOF("bell10",142,DOFPulse,DOFChimes),0, 0.30, 0, 0
    if bell=100 then playsound SoundFXDOF("bell100",143,DOFPulse,DOFChimes),0, 0.45, 0, 0
    if scn1=scn then scntimer.enabled=false
end sub

sub top1_hit
    if (canon5l.state)=1 then canon5_hit
    if (canon4l.state)=1 then canon4_hit
    if (canon3l.state)=1 then canon3_hit
    if (canon2l.state)=1 then canon2_hit
    if (canon1l.state)=1 then canon1_hit
end sub

sub top2_hit
    if (canon5l.state)=1 then canon5_hit
    if (canon4l.state)=1 then canon4_hit
    if (canon3l.state)=1 then canon3_hit
    if (canon2l.state)=1 then canon2_hit
    if (canon1l.state)=1 then canon1_hit
end sub

sub top3_hit
    if (canon5l.state)=1 then canon5_hit
    if (canon4l.state)=1 then canon4_hit
    if (canon3l.state)=1 then canon3_hit
    if (canon2l.state)=1 then canon2_hit
    if (canon1l.state)=1 then canon1_hit
end sub

sub targetl_hit
	DOF 121, DOFPulse
    addscore 10
	advancecreel
end sub

sub targetm_hit
	DOF 122, DOFPulse
	if (targetml.state)=1 then
		addscore (creel*10)
		upperr2l.state=0
		targetml.state=0
		upper1l.state=1
		upper2l.state=0
		upper3l.state=1
		upper4l.state=0
		bumperlight1.state=0
		bumperlight2.state=1
		bumperlight3.state=1
		bumperlight4.state=0
		topl1.state=1
		topl2.state=1
		topl3.state=1
    else
		addscore (creel)
    end if
end sub

sub targetr_hit
	DOF 123, DOFPulse
    addscore 10
	advancecreel
end sub

sub canon1_hit
    if (canon1l.state)=1 then
		addscore 10
		canon1l.state=0
		canon2l.state=1
		Bumperlight1.state=0
		Bumperlight2.state=1
		Bumperlight3.state=1
		Bumperlight4.state=0
		upper1l.state=0
		upper2l.state=1
		upper3l.state=0
		upper4l.state=1
    else
		addscore 1
    end if
end sub

sub canon2_hit
    if (canon2l.state)=1 then
		addscore 10
		canon2l.state=0
		canon3l.state=1
		Bumperlight1.state=1
		Bumperlight2.state=0
		Bumperlight3.state=0
		Bumperlight4.state=1
		upper1l.state=1
		upper2l.state=0
		upper3l.state=1
		upper4l.state=0
    else
		addscore 1
    end if
end sub

sub canon3_hit
    if (canon3l.state)=1 then
		addscore 10
		canon3l.state=0
		canon4l.state=1
		Bumperlight1.state=0
		Bumperlight2.state=1
		Bumperlight3.state=1
		Bumperlight4.state=0
		upper1l.state=0
		upper2l.state=1
		upper3l.state=0
		upper4l.state=1
    else
		addscore 1
    end if
end sub

sub canon4_hit
    if (canon4l.state)=1 then
		addscore 10
		canon4l.state=0
		canon5l.state=1
		Bumperlight1.state=1
		Bumperlight2.state=1
		Bumperlight3.state=1
		Bumperlight4.state=1
		upper1l.state=1
		upper2l.state=1
		upper3l.state=1
		upper4l.state=1
		upperr2l.state=1
		LeftOutlanel1.state=1
		RightOutlanel1.state=1
		ebll.state=1
    else
		addscore 1
    end if
end sub

sub canon5_hit
    if (canon5l.state)=1 then
		addscore 10
		canon5l.state=0
		canon1l.state=1
		LeftOutlanel1.state=0
		RightOutlanel1.state=0
		ebll.state=0
		upperr2l.state=1
		targetml.state=1
    else
		addscore 1
    end if
end sub

sub ebl_hit
	DOF 131, DOFPulse
    if (ebll.state)=1 then
    shootagain.state=1
    ebll.state=0
    end if
end sub


sub LeftOutlane_hit
	DOF 119, DOFPulse
    if (LeftOutlanel1.state)=1 then
    addscore 300
    else
    addscore 100
    end if
end sub

sub RightOutlane_hit
	DOF 120, DOFPulse
    if (RightOutlanel1.state)=1 then
    addscore 300
    else
    addscore 100
    end if
end sub

Sub LeftSlingShot_Slingshot()
	if tilt=false then PlaySoundAtVol SoundFXDOF("slingshot",103,DOFPulse,DOFContactors), ActiveBall, 1:DOF 104, DOFPulse
    addscore 1
End Sub

Sub RightSlingShot_Slingshot()
	if tilt=false then PlaySoundAtVol SoundFXDOF("slingshot",105,DOFPulse,DOFContactors), ActiveBall, 1:DOF 106, DOFPulse
	addscore 1
End Sub

sub bumper1_hit
    if tilt=false then playsoundAtVol SoundFXDOF("jet1",107,DOFPulse,DOFContactors), Bumper1, VolBump:DOF 111, DOFPulse
    if bumperlight1.state=1 then addscore 10 else addscore 1
end sub

sub bumper2_hit
    if tilt=false then playsoundAtVol SoundFXDOF("jet1",108,DOFPulse,DOFContactors), Bumper2, VolBump:DOF 112, DOFPulse
    if bumperlight2.state=1 then addscore 10 else addscore 1
end sub

sub bumper3_hit
    if tilt=false then playsoundAtVol SoundFXDOF("jet1",109,DOFPulse,DOFContactors), Bumper3, VolBump:DOF 113, DOFPulse
    if bumperlight3.state=1 then addscore 10 else addscore 1
end sub

sub bumper4_hit
    if tilt=false then playsoundAtVol SoundFXDOF("jet1",110,DOFPulse,DOFContactors), Bumper4, VolBump:DOF 114, DOFPulse
    if bumperlight4.state=1 then addscore 10 else addscore 1
end sub

sub LeftSlingShot1_Slingshot()
    addscore 1
	advancecreel
end sub

sub LeftSlingShot2_Slingshot()
    addscore 1
end sub

sub LeftSlingShot3_Slingshot()
    addscore 1
end sub

sub LeftSlingShot4_Slingshot()
    addscore 1
end sub

sub RightSlingShot1_Slingshot()
    addscore 1
	advancecreel
end sub

sub RightSlingShot2_Slingshot()
    addscore 1
end sub

sub advancecreel
    creel=creel+10
    if creel>50 then creel=10
	CRtimer.enabled=1
end sub

Sub CRTimer_Timer()
	DOF 132, DOFPulse
	CreelP.RotX = CreelP.rotx+6
	If CreelP.RotX = 360 Then:CRTimer.Enabled=0:CreelP.RotX=0:End If
	If CreelP.RotX =  36 Then:CRTimer.Enabled=0:End If
	If CreelP.RotX =  72 Then:CRTimer.Enabled=0:End If
	If CreelP.RotX = 108 Then:CRTimer.Enabled=0:End If
	If CreelP.RotX = 144 Then:CRTimer.Enabled=0:End If
	If CreelP.RotX = 180 Then:CRTimer.Enabled=0:End If
	If CreelP.RotX = 216 Then:CRTimer.Enabled=0:End If
	If CreelP.RotX = 252 Then:CRTimer.Enabled=0:End If
	If CreelP.RotX = 288 Then:CRTimer.Enabled=0:End If
	If CreelP.RotX = 324 Then:CRTimer.Enabled=0:End If
End Sub

sub newgame
	state=true
	rep(1)=0
	rep(2)=0
	rep(3)=0
	rep(4)=0
	eg=0
	score(1)=0
	score(2)=0
	score(3)=0
	score(4)=0
	truesc(1)=0
	truesc(2)=0
	truesc(3)=0
	truesc(4)=0
	bumper1.hashitevent=1
	bumper2.hashitevent=1
	bumper3.hashitevent=1
	bumper4.hashitevent=1
	bumperlight1.state=1
	bumperlight2.state=0
	bumperlight3.state=0
	bumperlight4.state=1
	canon1l.state=1
	canon2l.state=0
	canon3l.state=0
	canon4l.state=0
	canon5l.state=0
	upper1l.state=1
	upper2l.state=0
	upper3l.state=1
	upper4l.state=0
	topl1.state=1
	topl2.state=1
	topl3.state=1
	upperr2l.state=0
	shootagain.state=0
	ebll.state=0
	LeftOutlanel1.state=0
	RightOutlanel1.state=0
	targetml.state=0
	BIPReel.setvalue 1
	MatchReel.setvalue 10
	TiltReel.setvalue 0
	GameOverReel.setvalue 0
	tilt=false
	tiltsens=0
	ballinplay=1
	nb.CreateBall
	nb.kick 90,6
	If B2SOn then Controller.B2SSetBallInPlay 1
end sub

Sub Drain_Hit()
	DOF 125, DOFPulse
	playsoundAtVol "drainshorter", Drain, 1
	Drain.DestroyBall
	nextball
End Sub

sub nextball
	if tilt=true then
		tilt=false
		If B2SOn Then Controller.B2SSetTilt 0
		TiltReel.setvalue 0
		bumper1.hashitevent=1
		bumper2.hashitevent=1
		bumper3.hashitevent=1
		bumper4.hashitevent=1
		bumperlight1.state=1
		bumperlight2.state=0
		bumperlight3.state=0
		bumperlight4.state=1
'		canon1l.state=1
'		canon2l.state=0
'		canon3l.state=0
'		canon4l.state=0
'		canon5l.state=0
		upperr2l.state=0
		shootagain.state=0
		LeftOutlanel1.state=0
		RightOutlanel1.state=0
		topl1.state=1
		topl2.state=1
		topl3.state=1
		upper1l.state=1
		upper2l.state=0
		upper3l.state=1
		upper4l.state=0
		ebll.state=0
		targetml.state=0
		tiltseq.stopplay
	end if

	if (shootagain.state)=1 then
	playsound SoundFXDOF("kickerkick",126, DOFPulse,DOFContactors)
	newball
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
	ballinplay=ballinplay+1
	if ballinplay>5 then
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
	newball
	playsound "kickerkick"
	ballreltimer.enabled=true
	end if
	If B2SOn then Controller.B2SSetBallInPlay ballinplay
	BIPReel.setvalue ballinplay
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
	newball
	playsound "kickerkick"
	ballreltimer.enabled=true
	end if
	end if
end Sub

sub ballreltimer_timer
	if eg=1 then
	matchnum
	BIPReel.setvalue 0
	state=false
	for i=1 to 4
	if truesc(i)>hisc then
	hisc=truesc(i)
	hisctxt.text=hisc
	end if
	next
	play(currpl-1).state=0
	playno=0
	GameOverReel.setvalue 1
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

sub newball
'	canon1l.state=1
'	canon2l.state=0
'	canon3l.state=0
'	canon4l.state=0
'	canon5l.state=0
	upper1l.state=0
	upper2l.state=0
	upper3l.state=0
	upper4l.state=0
	topl1.state=1
	topl2.state=1
	topl3.state=1
	bumperlight1.state=1
	bumperlight2.state=0
	bumperlight3.state=0
	bumperlight4.state=1
	upper1l.state=1
	upper2l.state=0
	upper3l.state=1
	upper4l.state=0
	upperr2l.state=0
	shootagain.state=0
	LeftOutlanel1.state=0
	RightOutlanel1.state=0
	ebll.state=0
	targetml.state=0
end sub

sub matchnum

	MatchReel.setvalue matchnumb
	If B2SOn Then

		if matchnumb=0 then
			Controller.B2SSetMatch 10
		else
			Controller.B2SSetMatch matchnumb
		end if
	end if
    for i=1 to playno
    if (matchnumb*10)=(score(i) mod 10) then
    credit=credit+1
	DOF 130, DOFOn
	if credit>15 then credit=15
	If B2SOn Then
		Controller.B2SSetCredits Credit
	end if
    PlaySound SoundFXDOF("knocke",128, DOFPulse,DOFKnocker):DOF 129, DOFPulse
	CreditReel.setvalue credit
    playsound "click"
    end if
    next
end sub

Sub CheckTilt
	If Tilttimer.Enabled = True Then
		TiltSens = TiltSens + 1
		if TiltSens = 3 Then GameTilted
	Else
		TiltSens = 0
		Tilttimer.Enabled = True
	End If
End Sub

Sub Tilttimer_Timer()
	Tilttimer.Enabled = False
End Sub

Sub GameTilted
	Tilt = True
	TiltReel.setvalue 1
			tiltsens = 0
	If B2SOn Then Controller.B2SSetTilt 1
	playsound "tilt"
	turnoff
End Sub

sub turnoff
    bumper1.hashitevent=0
    bumper2.hashitevent=0
    bumper3.hashitevent=0
    bumper4.hashitevent=0
    tiltseq.play seqalloff
end sub


sub savehs
    savevalue "JollyRoger", "credit", credit
    savevalue "JollyRoger", "hiscore", hisc
    savevalue "JollyRoger", "match", matchnumb
    savevalue "JollyRoger", "score1", score(1)
    savevalue "JollyRoger", "score2", score(2)
    savevalue "JollyRoger", "score3", score(3)
    savevalue "JollyRoger", "score4", score(4)
	savevalue "JollyRoger", "creel", creel
end sub

sub loadhs
    dim temp
	temp = LoadValue("JollyRoger", "credit")
    If (temp <> "") then credit = CDbl(temp)
    temp = LoadValue("JollyRoger", "hiscore")
    If (temp <> "") then hisc = CDbl(temp)
    temp = LoadValue("JollyRoger", "match")
    If (temp <> "") then matchnumb = CDbl(temp)
    temp = LoadValue("JollyRoger", "score1")
    If (temp <> "") then score(1) = CDbl(temp)
    temp = LoadValue("JollyRoger", "score2")
    If (temp <> "") then score(2) = CDbl(temp)
    temp = LoadValue("JollyRoger", "score3")
    If (temp <> "") then score(3) = CDbl(temp)
    temp = LoadValue("JollyRoger", "score4")
    If (temp <> "") then score(4) = CDbl(temp)
    temp = LoadValue("JollyRoger", "creel")
    If (temp <> "") then creel = CDbl(temp)
end sub

sub ballhome_hit
    ballrelenabled=1
end sub

Sub Shooter_Unhit
	DOF 127, DOFPulse
End Sub

sub ballout_hit
    shootagain.state=0
    'spsa.text=" "
end sub

sub ballrel_hit
    if ballrelenabled=1 then
    playsoundAtVol "launchball", ballrel, 1
    ballrelenabled=0
    end if
end sub

Sub JollyRoger_Exit()
	Savehs
	turnoff
	If B2SOn Then Controller.stop
End Sub


'**********Sling Shot Animations
' Rstep and Lstep  are the variables that increment the animation
'****************
Dim RStep, Lstep

'Sub RightSlingShot_Slingshot
'    PlaySound "right_slingshot", 0, 1, 0.05, 0.05
'    RSling.Visible = 0
'    RSling1.Visible = 1
'    sling1.TransZ = -20
'    RStep = 0
'    RightSlingShot.TimerEnabled = 1
'    addscore 1
'End Sub

'Sub RightSlingShot_Timer
'    Select Case RStep
'        Case 3:RSLing1.Visible = 0:RSLing2.Visible = 1:sling1.TransZ = -10
'        Case 4:RSLing2.Visible = 0:RSLing.Visible = 1:sling1.TransZ = 0:RightSlingShot.TimerEnabled = 0
'    End Select
'    RStep = RStep + 1
'End Sub

'Sub LeftSlingShot_Slingshot
'    PlaySound "left_slingshot",0,1,-0.05,0.05
'    LSling.Visible = 0
'    LSling1.Visible = 1
'    sling2.TransZ = -20
'    LStep = 0
'    LeftSlingShot.TimerEnabled = 1
'    addscore 1
'End Sub

'Sub LeftSlingShot_Timer
'    Select Case LStep
'        Case 3:LSLing1.Visible = 0:LSLing2.Visible = 1:sling2.TransZ = -10
'        Case 4:LSLing2.Visible = 0:LSLing.Visible = 1:sling2.TransZ = 0:LeftSlingShot.TimerEnabled = 0
'    End Select
'    LStep = LStep + 1
'End Sub

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

Sub PlaySoundAtVol(sound, tableobj, Volum)
  PlaySound sound, 1, Volum, Pan(tableobj), 0,0,0, 1, AudioFade(tableobj)
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

Function AudioFade(tableobj) ' Fades between front and back of the table (for surround systems or 2x2 speakers, etc), depending on the Y position on the table. "JollyRoger" is the name of the table
  Dim tmp
  tmp = tableobj.y * 2 / JollyRoger.height-1
  If tmp > 0 Then
    AudioFade = Csng(tmp ^10)
  Else
    AudioFade = Csng(-((- tmp) ^10) )
  End If
End Function

Function AudioPan(tableobj) ' Calculates the pan for a tableobj based on the X position on the table. "JollyRoger" is the name of the table
  Dim tmp
  tmp = tableobj.x * 2 / JollyRoger.width-1
  If tmp > 0 Then
    AudioPan = Csng(tmp ^10)
  Else
    AudioPan = Csng(-((- tmp) ^10) )
  End If
End Function

Function Pan(ball) ' Calculates the pan for a ball based on the X position on the table. "JollyRoger" is the name of the table
    Dim tmp
    tmp = ball.x * 2 / JollyRoger.width-1
    If tmp > 0 Then
        Pan = Csng(tmp ^10)
    Else
        Pan = Csng(-((- tmp) ^10) )
    End If
End Function

Function AudioFade(ball) ' Can this be together with the above function ?
  Dim tmp
  tmp = ball.y * 2 / JollyRoger.height-1
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

' Thalamus : Exit in a clean and proper way
Sub JollyRoger_exit()
  Controller.Pause = False
  Controller.Stop
End Sub


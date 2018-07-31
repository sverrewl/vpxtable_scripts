'	-------------------------------------------------
'Aspen Brunswick 1976 Updon719
'Many thanks to Borgdog for the use of his Vulcan table
'as a template and help on a few of the scripting problems
'	-------------------------------------------------
'
' 4 Player at 1 Display - Edition by STAT, just a bit changes on the Script
'

' Thalamus 2018-07-19
' Added/Updated "Positional Sound Playback Functions" and "Supporting Ball & Sound Functions"
' No special SSF tweaks yet.

Option Explicit
Randomize

Const cGameName = "Aspen"

Dim balls
dim ebcount
Dim replays
Dim hisc
Dim Controller
Dim maxplayers
Dim players
Dim player
Dim credit, dbonus
Dim score(4)
Dim Add50, Add100, Add1000
Dim sreels(4)
Dim Shoot(4), CanPlay(4)
Dim state
Dim tilt
Dim tiltsens
Dim target(9)
Dim Bonus
Dim Bonuslight(19)
Dim ballinplay
Dim ballrenabled
Dim rep
Dim rst
Dim eg
Dim bell
Dim i,j,tnumber, objekt, light, ts
Dim awardcheck
dim rstep, lstep
Dim dw1step, dw2step, dw3step, dw4step, dw5step, dw6step
Dim rw1step, rw2step
On Error Resume Next
ExecuteGlobal GetTextFile("controller.vbs")
If Err Then MsgBox "Can't open controller.vbs"
On Error Goto 0

sub Aspen_init
	LoadEM
	maxplayers=4
	set sreels(1) = ScoreReel1
	set sreels(2) = ScoreReel2
	set sreels(3) = ScoreReel3
	set sreels(4) = ScoreReel4
	set Shoot(1)= Shoot1
	set Shoot(2)= Shoot2
	set Shoot(3)= Shoot3
	set Shoot(4)= Shoot4
	set CanPlay(1) = CanPlay1
	set CanPlay(2) = CanPlay2
	set CanPlay(3) = CANPLAY3
	set CanPlay(4) = CANPLAY4

	player=1
	For each light in BonusLights:light.State = 0: Next
	For each light in bonusGI:light.State = 0: Next
	For each light in bumperlights:light.State = 0: Next
	loadhs
	if hisc="" then hisc=10000
	hstxt.text=hisc
	gamov.text="Game Over"
	gamov.timerenabled=1
	tilttxt.timerenabled=1
	if balls="" then balls=5
	if balls<>3 and balls<>5 then balls=5

 	If B2SOn OR Aspen.ShowDT = False Then
		setBackglass.enabled=True
		For each objekt in Backdropstuff: objekt.visible=false: next
	End If

	for i = 1 to maxplayers
		sreels(i).setvalue(score(i))
	next
	PlaySound "Tone100"
	tilt=false

	If Aspen.ShowDT = False then

	End If
End sub

sub setBackglass_timer
    Controller.B2SSetGameOver 1
    Controller.B2SSetScorePlayer 5, hisc
    Controller.B2SSetScorePlayer 1, Score(1)
	me.enabled=false
end sub

sub gamov_timer
	if state=false then
		If B2SOn then Controller.B2SSetGameOver 0
		gamov.text=""
		gtimer.enabled=true
	end if
	gamov.timerenabled=0
end sub

sub gtimer_timer
	if state=false then
		gamov.text="Game Over"
		If B2SOn then Controller.B2SSetGameOver 1
		gamov.timerenabled=1
		gamov.timerinterval= (INT (RND*10)+5)*100
	end if
	me.enabled=0
end sub

sub tilttxt_timer
	if state=false then
		tilttxt.text=""
		If B2SOn then Controller.B2SSetTilt 33,0
		ttimer.enabled=true
	end if
	tilttxt.timerenabled=0
end sub

sub ttimer_timer
	if state=false then
		tilttxt.text="TILT"
		If B2SOn then Controller.B2SSetTilt 33,1
		tilttxt.timerenabled=1
		tilttxt.timerinterval= (INT (RND*10)+5)*100
	end if
	me.enabled=0
end sub

Sub NumPlay_Timer
	if B2SOn then
		if i = 1 and j < 10 then
			Controller.B2ssetplayerup 30, players: i = 0
		else
			Controller.B2ssetplayerup 30, 0: i = 1
		end if
	j = j + 1:if j = 10 then Controller.B2ssetplayerup 30, 1:NumPlay.enabled=false
	end if
End Sub

Sub Aspen_KeyDown(ByVal keycode)

    if keycode=StartGameKey then
		TurnScores.enabled=false
		if state=false then
		coindelay.enabled=true
		playsound "startup"
		ballinplay=1
		player=1
		If b2son Then
			Controller.B2ssetballinplay 32, Ballinplay
			Controller.B2ssetplayerup 30, 1
			Controller.B2SSetGameOver 0
		End If
		shoot(player).state=1
		tilt=false
		state=true
		CanPlay(player).state=1
		players=1
		rst=0
		resettimer.enabled=true
	    else if state=true and players < maxplayers and Ballinplay=1 then
		credit=credit+1
		players=players+1
		i=1:j=0:NumPlay.enabled=true
		CanPlay(players).state=1
		playsound "tone500"
	   end if
	  end if
	end if
	If keycode = PlungerKey Then
		Plunger.PullBack
		PlaySound "plungerpull",0,1,0.25,0.25
	End If

  if tilt=false and state=true then
	If keycode = LeftFlipperKey Then
		LeftFlipper.RotateToEnd
		PlaySound SoundFXDOF("flipperup",101,DOFOn,DOFContactors), 0, .67, -0.05, 0.05
		PlaySound "Buzz",-1,.05,-0.05, 0.05
	End If

	If keycode = RightFlipperKey Then
		RightFlipper.RotateToEnd
		PlaySound SoundFXDOF("flipperup",102,DOFOn,DOFContactors), 0, .67, 0.05, 0.05
		PlaySound "Buzz1",-1,.05,0.05,0.05
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
  end IF
end sub


Sub Aspen_KeyUp(ByVal keycode)

	If keycode = PlungerKey Then
		Plunger.Fire
		PlaySound "plunger",0,1,0.25,0.25
	End If

   	if keycode = LeftFlipperKey then
		OperatorMenuTimer.Enabled = false
	end if

   If tilt=false and state=true then
	If keycode = LeftFlipperKey Then
		LeftFlipper.RotateToStart
		PlaySound SoundFXDOF("flipperdown",101,DOFOff,DOFContactors), 0, 1, -0.05, 0.05
		StopSound "Buzz"
	End If

	If keycode = RightFlipperKey Then
		RightFlipper.RotateToStart
		PlaySound SoundFXDOF("flipperdown",102,DOFOff,DOFContactors), 0, 1, 0.05, 0.05
		StopSound "Buzz1"
	End If
   End if
End Sub

sub flippertimer_timer()
	LFlip.RotY = LeftFlipper.CurrentAngle
	RFlip.RotY = RightFlipper.CurrentAngle
end sub

sub coindelay_timer
    coindelay.enabled=false
end sub

sub resettimer_timer
   rst=rst+1
	for i = 1 to maxplayers
		sreels(i).resettozero
   next
    If b2son then
		for i = 1 to maxplayers
		  Controller.B2SSetScorePlayer 1, 0
		  Controller.B2SSetScoreRollover 24 + i, 0
		next
	End If
    if rst=18 then
		playsound "flipperup"
    end if
    if rst=22 then
		newgame
		resettimer.enabled=false
    end if
end sub

Sub addcredit
      credit=credit+1
	  DOF 133, DOFOn
End sub

Sub Drain_Hit()
	DOF 129, DOFPulse
	PlaySound "drain",0,1,0,0.25
	Drain.DestroyBall
	me.timerenabled=1
	If DBlight.state=1 or TRlight.state=1 Then
	DBlight.state=0
	TRlight.state=0
	'LightA.state=1
	'LightB.state=1
	'LightC.state=1
	end if
	for each light in targetlights:light.state=1:Next
	for each light in bonusGI:light.state=0:Next
	for each light in BonusLights:light.state=0:next
	LTxball.state=0
End Sub

Sub Drain_timer
	if trlight.state=1 then
		bonus=bonus*3
	 else
		if dblight.state=1 then bonus=bonus*2
	end if
	dbonus=1
	scorebonus.enabled=true
	me.timerenabled=0
End Sub

sub ballhome_hit
	ballrenabled=1
end sub

sub ballhome_unhit
	DOF 134, DOFPulse
end sub

sub ballrel_hit
	if ballrenabled=1 then
		shootagain.state=lightstateoff
		ballrenabled=0
	end if
end sub

sub ScoreBonus_timer
   if tilt=false and bonus>0 then
		score(player)=score(player)+(1000*dbonus)
		sreels(player).addvalue(1000*dbonus)

		If B2SOn Then Controller.B2SSetScorePlayer 1, score(player)  'MOD 100000
		If dbonus=2 Then
			PlaySound SoundFXDOF("tone1000",143,DOFPulse,DOFChimes)
		  Elseif dbonus=3 then
			PlaySound SoundFXDOF("tone1000",143,DOFPulse,DOFChimes)
		  Else
			PlaySound SoundFXDOF("tone1000",143,DOFPulse,DOFChimes)
		End If
		bonus=bonus-1
		else
		bonus=0
		for each light in bonuslights: light.state=0: next
   end if
   if bonus=0 then
     if shootagain.state=lightstateon then
	    newball
 	    ballreltimer.enabled=true
     else
	  if players=1 or player=players then
		player=1
	   Else
		player=player+1
	  end if
	  If B2SOn Then Controller.B2SSetScorePlayer 1, score(player):Controller.B2ssetplayerup 30, player
	  for i = 1 to 4: Shoot(i).state=0: Next
	  shoot(player).state=1
	  nextball
	 end if
	 scorebonus.enabled=false
   end if
End sub


sub newgame
	player=1
	shoot(player).state=1
	ebcount=0
	for i = 1 to 4:
		score(i)=0
	next
	If b2son then
		Controller.B2SSetScorePlayer 1, 0
	End If
    eg=0
    rep=0
	shootagain.state=lightstateoff
	for each light in bonusGI:light.state=0:next
	for each light in bumperlights:light.state=1:next
	for each light in BonusLights:light.state=0:next
    for each light in targetlights:light.state=1:next
	gamov.text=" "
	tilttxt.text=" "
    If b2son then
		Controller.B2SSetGameOver 0
		Controller.B2SSetTilt 33,0
	End If
    biptext.text="1"
    BallRelease.CreateBall
   	BallRelease.kick 60,35,0
	playsound SoundFXDOF("flipperup",128,DOFPulse,DOFContactors)
end sub

sub newball
	ShootAgain.state=0
End Sub

sub nextball
    if tilt=true then
	  bumper1.force=12
      tilt=false
      tilttxt.text=" "
		If b2son then
			Controller.B2SSetTilt 33,0
			Controller.B2ssetdata 1, 1
		End If
    end if
	ebcount=0
	if player=1 then ballinplay=ballinplay+1
	if ballinplay>balls then
		playsound "game end"
		eg=1
		ballreltimer.enabled=true
	  else
		if state=true and tilt=false then
		  newball
		  ballreltimer.enabled=true
		end if
		biptext.text=ballinplay
		If b2son then Controller.B2ssetballinplay 32, Ballinplay
	end if
End Sub

sub ballreltimer_timer
  if eg=1 then
	  turnoff
	  biptext.text=" "
	  state=false
	  gamov.text="GAME OVER"
	  for i=1 to maxplayers
		CanPlay(i).state=0
		Shoot(i).state=0
		if score(i)>hisc then hisc=score(i)
	  next
	  hstxt.text=hisc
	  savehs
	  If b2son then
        Controller.B2SSetGameOver 1
        Controller.B2ssetballinplay 32, 0
	    Controller.B2SSetScorePlayer 5, hisc
	    Controller.B2ssetPlayerUp 30, 0
	  End If
	ballreltimer.enabled=false
	ts = 1:TurnScores.enabled=true
	else
    BallRelease.CreateBall
	BallRelease.kick 60,45,0
	playsound SoundFXDOF("flipperup",128,DOFPulse,DOFContactors)
    ballreltimer.enabled=false
  end if
end sub

Sub TurnScores_Timer
	If B2SOn Then
		Controller.B2SSetScorePlayer 1, score(ts)
		Controller.B2ssetPlayerUp 30, ts
		ts = ts + 1
		if ts > players then ts = 1
	End If
End Sub


'********** Bumpers

Sub Bumper1_Hit
	if tilt=false then
	LightBumper1.state=1
	playsound SoundFXDOF("fx_bumper4",107,DOFPulse,DOFContactors)
	DOF 108,DOFPulse
	addscore 1000
	me.timerenabled=1
	end if
	if Lightdingwall5.state=1 and LightTGtop.state=1 and LightBumper1.state=1 and Lightdingwall6.state=1 and LightSpinner2.state=1 and LightSpinner3.state=1 and LightSpinner1.state=1 and LightLeftSlingShot.state=1 and LightRightSlingShot.state=1 Then
	LTxball.state=1
	end if
End Sub

Sub Bumper1_timer
	Bumper1Ring.Enabled=0
	BumperRing1.transz=BumperRing1.transz-4
	if BumperRing1.transz=-36 then
		Bumper1Ring.enabled=1
		me.timerenabled=0
	end if
End Sub

Sub Bumper1Ring_timer
	BumperRing1.transz=BumperRing1.transz+4
	If BumperRing1.transz=0 then Bumper1Ring.enabled=0
End sub

'************Spinners

Sub Spinner1_Spin
	if tilt=false then
	LightSpinner1.state=1
	if Lbulb7.state=1 Then
	addscore 500
	Else
	addscore 50
	PlaySound "fx_spinner",0,.25,0,0.25
	if Lightdingwall5.state=1 and LightTGtop.state=1 and LightBumper1.state=1 and Lightdingwall6.state=1 and LightSpinner2.state=1 and LightSpinner3.state=1 and LightSpinner1.state=1 and LightLeftSlingShot.state=1 and LightRightSlingShot.state=1 Then
	LTxball.state=1
	end if
	end If
	end if
End Sub

Sub Spinner2_Spin
	if tilt=false then
	LightSpinner2.state=1
	if Lbulb7.state=1 Then
	addscore 500
	Else
	addscore 50
	PlaySound "fx_spinner",0,.25,0,0.25
	if Lightdingwall5.state=1 and LightTGtop.state=1 and LightBumper1.state=1 and Lightdingwall6.state=1 and LightSpinner2.state=1 and LightSpinner3.state=1 and LightSpinner1.state=1 and LightLeftSlingShot.state=1 and LightRightSlingShot.state=1 Then
	LTxball.state=1
	end if
	end If
	end if
End Sub

Sub Spinner3_Spin
	if tilt=false then
	LightSpinner3.state=1
	if Lbulb7.state=1 Then
	addscore 500
	Else
	addscore 50
	PlaySound "fx_spinner",0,.25,0,0.25
	if Lightdingwall5.state=1 and LightTGtop.state=1 and LightBumper1.state=1 and Lightdingwall6.state=1 and LightSpinner2.state=1 and LightSpinner3.state=1 and LightSpinner1.state=1 and LightLeftSlingShot.state=1 and LightRightSlingShot.state=1 Then
	LTxball.state=1
	end if
	end if
	end if
End Sub

'************** Slings

Sub RightSlingShot_Slingshot
	playsound SoundFXDOF("right_slingshot",105,DOFPulse,DOFContactors), 0, 1, 0.05, 0.05
	DOF 106,DOFPulse
	LightRightSlingShot.State=1
	if Lbulb7.state=1 Then
	addscore 500
	Else
	addscore 50
    addbonus
	LightRightSlingShot.state=1
	if Lightdingwall5.state=1 and LightTGtop.state=1 and LightBumper1.state=1 and Lightdingwall6.state=1 and LightSpinner2.state=1 and LightSpinner3.state=1 and LightSpinner1.state=1 and LightLeftSlingShot.state=1 and LightRightSlingShot.state=1 Then
	LTxball.state=1
	RSling.Visible = 0
    RSling1.Visible = 1
	slingR.objroty = -15
    RStep = 1
    RightSlingShot.TimerEnabled = 1
	end if
	end if
End Sub

Sub RightSlingShot_Timer
    Select Case RStep
        Case 3:RSLing1.Visible = 0:RSLing2.Visible = 1:slingR.objroty = -7
        Case 4:	slingR.objroty = 0:RSLing2.Visible = 0:RSLing.Visible = 1:RightSlingShot.TimerEnabled = 0
    End Select
    RStep = RStep + 1
End Sub

Sub LeftSlingShot_Slingshot
	playsound SoundFXDOF("left_slingshot",103,DOFPulse,DOFContactors), 0, 1, -0.05, 0.05
	DOF 104,DOFPulse
	LightLeftSlingShot.state=1
	if Lbulb7.state=1 Then
	addscore 500
	Else
	addscore 50
    addbonus
	LightLeftSlingShot.state=1
	if Lightdingwall5.state=1 and LightTGtop.state=1 and LightBumper1.state=1 and Lightdingwall6.state=1 and LightSpinner2.state=1 and LightSpinner3.state=1 and LightSpinner1.state=1 and LightLeftSlingShot.state=1 and LightRightSlingShot.state=1 Then
	LTxball.state=1
	LSling.Visible = 0
    LSling1.Visible = 1
	slingL.objroty = 15
    LStep = 1
    LeftSlingShot.TimerEnabled = 1
	end if
	end if
End Sub

Sub LeftSlingShot_Timer
    Select Case LStep
        Case 3:LSLing1.Visible = 0:LSLing2.Visible = 1:slingL.objroty = 7
        Case 4:slingL.objroty = 0:LSLing2.Visible = 0:LSLing.Visible = 1:LeftSlingShot.TimerEnabled = 0
    End Select
    LStep = LStep + 1
End Sub

'************** Dingwalls

sub dingwall1_hit
	if tilt=false then
	if Lbulb7.state=1 Then
	addscore 500
	Else
	addscore 50
	end if
	rdw1.visible=0
	RDW1a.visible=1
	dw1step=1
	Me.timerenabled=1
	end if
end sub

sub dingwall1_timer
	select case dw1step
		Case 1: RDW1a.visible=0: rdw1.visible=1
		case 2:	rdw1.visible=0: rdw1b.visible=1
		Case 3: rdw1b.visible=0: rdw1.visible=1: Me.timerenabled=0
	end Select
	dw1step=dw1step+1
end sub

sub dingwall2_hit
	if tilt=false then
	if Lbulb7.state=1 Then
	addscore 500
	Else
	addscore 50
	end if
	rdw2.visible=0
	RDW2a.visible=1
	dw2step=1
	Me.timerenabled=1
	end if
end sub

sub dingwall2_timer
	select case dw2step
		Case 1: RDW2a.visible=0: rdw2.visible=1
		case 2:	rdw2.visible=0: rdw2b.visible=1
		Case 3: rdw2b.visible=0: rdw2.visible=1: me.timerenabled=0
	end Select
	dw2step=dw2step+1
end sub

sub dingwall3_hit
	if tilt=false then
	if Lbulb7.state=1 Then
	addscore 500
	Else
	addscore 50
	Rdw3.visible=0
	RDW3a.visible=1
	dw3step=1
	Me.timerenabled=1
	end if
	end if
end sub

sub dingwall3_timer
	select case dw3step
		Case 1: RDW3a.visible=0: Rdw3.visible=1
		case 2:	Rdw3.visible=0: rdw3b.visible=1
		Case 3: rdw3b.visible=0: Rdw3.visible=1: me.timerenabled=0
	end Select
	dw3step=dw3step+1
end sub

sub dingwall4_hit
	if tilt=false then
	if Lbulb7.state=1 Then
	addscore 500
	Else
	addscore 50
	Rdw4.visible=0
	RDW4a.visible=1
	dw4step=1
	Me.timerenabled=1
	end if
	end if
end sub

sub dingwall4_timer
	select case dw4step
		Case 1: RDW4a.visible=0: Rdw4.visible=1
		case 2:	Rdw4.visible=0: rdw4b.visible=1
		Case 3: rdw4b.visible=0: Rdw4.visible=1: me.timerenabled=0
	end Select
	dw4step=dw4step+1
end sub

sub dingwall5_hit
	if tilt=false then
	Lightdingwall5.state=1
	if Lbulb7.state=1 Then
	addscore 500
	Else
	addscore 50
	addbonus
	end if
	if Lightdingwall5.state=1 and LightTGtop.state=1 and LightBumper1.state=1 and Lightdingwall6.state=1 and LightSpinner2.state=1 and LightSpinner3.state=1 and LightSpinner1.state=1 and LightLeftSlingShot.state=1 and LightRightSlingShot.state=1 Then
	LTxball.state=1
	end if
	Rdw5.visible=0
	RDW5a.visible=1
	dw5step=1
	Me.timerenabled=1
	end if
end sub

sub dingwall5_timer
	select case dw5step
		Case 1: RDW5a.visible=0: Rdw5.visible=1
		case 2:	Rdw5.visible=0: rdw5b.visible=1
		Case 3: rdw5b.visible=0: Rdw5.visible=1: me.timerenabled=0
	end Select
	dw5step=dw5step+1
end sub

sub dingwall6_hit
	If tilt=false then
	Lightdingwall6.state=1
	if Lbulb7.state=1 Then
	addscore 500
	Else
	addscore 50
	addbonus
	end if
	if Lightdingwall5.state=1 and LightTGtop.state=1 and LightBumper1.state=1 and Lightdingwall6.state=1 and LightSpinner2.state=1 and LightSpinner3.state=1 and LightSpinner1.state=1 and LightLeftSlingShot.state=1 and LightRightSlingShot.state=1 Then
	LTxball.state=1
	end if
	Rdw6.visible=0
	RDW6a.visible=1
	dw6step=1
	Me.timerenabled=1
	end if
end sub

sub dingwall6_timer
	select case dw6step
		Case 1: RDW6a.visible=0: Rdw6.visible=1
		case 2:	Rdw6.visible=0: rdw6b.visible=1
		Case 3: rdw6b.visible=0: Rdw6.visible=1: me.timerenabled=0
	end Select
	dw6step=dw6step+1
end sub

'********** Triggers

sub TGtop_hit
	DOF 115, DOFPulse
    if tilt=false then
	addscore 1000
	for each light in bonusGI:light.state=1:Next
	playsound "bonus"
    LightTGtop.state=1
	if Lightdingwall5.state=1 and LightTGtop.state=1 and LightBumper1.state=1 and Lightdingwall6.state=1 and LightSpinner2.state=1 and LightSpinner3.state=1 and LightSpinner1.state=1 and LightLeftSlingShot.state=1 and LightRightSlingShot.state=1 Then
	LTxball.state=1

	end if
	end if
end sub

sub outL_hit
   DOF 120, DOFPulse
   if tilt=False then
	addbonus
	if Lbulb7.state=1 Then
	addscore 1000
	Else
	addscore 100
	end if
	if LTxball.state=1 then Shootagain.state=1
	end if
end Sub

sub outR_hit
   DOF 121, DOFPulse
   if tilt=False then
	addbonus
	if Lbulb7.state=1 Then
	addscore 1000
	Else
	addscore 100
	end if
	if LTxball.state=1 then Shootagain.state=1
	end if
end Sub

'******Targets

Sub TargetA_hit
	DOF 126, DOFPulse
	LightA.state=0
	addbonus
	addbonus
	if Lbulb7.state=1 Then
		addscore 1000
		Else
		addscore 100
	end if
	if LightA.state=0 and LightC.state=0 then DBlight.state=1
	if LightA.state=0 and LightB.state=0 and LightC.state=0 Then TRlight.state=1 and DBlight.state=1

end Sub

Sub TargetB_hit
	DOF 127, DOFPulse
	LightB.state=0
	addbonus
	addbonus
	if Lbulb7.state=1 Then
	addscore 1000
	Else
	addscore 100
	end if
	if LightA.state=0 and LightC.state=0  then DBlight.state=1
	if LightA.state=0 and LightB.state=0 and LightC.state=0  Then TRlight.state=1 and DBlight.state=1

end sub

Sub TargetC_hit
	DOF 127, DOFPulse
	LightC.state=0
	addbonus
	addbonus
	if Lbulb7.state=1 Then
	addscore 1000
	Else
	addscore 100
	end if
	if LightA.state=0 and LightC.state=0 then DBlight.state=1
	if LightA.state=0 and LightB.state=0 and LightC.state=0 Then TRlight.state=1 and DBlight.state=1:
end Sub

sub addscore(points)
  if tilt=false then

    If Points < 100 and AddScore50Timer.enabled = false Then
        Add50 = Points \ 50
        AddScore50Timer.Enabled = TRUE
      ElseIf Points < 1000 and AddScore100Timer.enabled = false Then
       Add100 = Points \ 100
        AddScore100Timer.Enabled = TRUE
    ElseIf AddScore1000Timer.enabled = false Then
        Add1000 = Points \ 1000
        AddScore1000Timer.Enabled = TRUE
   End If
  end if
End Sub

Sub AddScore50Timer_Timer()
    if Add50 > 0 then
       AddPoints 50
        Add50 = Add50 - 1
    Else
        Me.Enabled = FALSE
    End If
End Sub

Sub AddScore100Timer_Timer()
    if Add100 > 0 then
        AddPoints 100
        Add100 = Add100 - 1
   Else
        Me.Enabled = FALSE
    End If
End Sub

Sub AddScore1000Timer_Timer()
   if Add1000 > 0 then
        AddPoints 1000
        Add1000 = Add1000 - 1
    Else
       Me.Enabled = FALSE
    End If
End Sub

Sub AddPoints(Points)
   score(player)=score(player)+points
	sreels(player).addvalue(points)
	If B2SOn Then Controller.B2SSetScorePlayer 1, score(player) 'MOD 100000

     'Sounds: there are 3 sounds: tens, hundreds and thousands
     If Points = 100 AND(Score(player) MOD 1000) \ 100 = 0 Then
		PlaySound SoundFXDOF("tone50",143,DOFPulse,DOFChimes)
      ElseIf Points = 50 AND(Score(player) MOD 100) \ 50 = 0 Then
		PlaySound SoundFXDOF("tone50",142,DOFPulse,DOFChimes)
      ElseIf points = 1000 Then
		PlaySound SoundFXDOF("tone1000",143,DOFPulse,DOFChimes)
	  elseif Points = 100 Then
		PlaySound SoundFXDOF("tone100",142,DOFPulse,DOFChimes)
      Else
		PlaySound SoundFXDOF("tone50",141,DOFPulse,DOFChimes)
    End If
end sub

Sub CheckTilt
	If Tilttimer.Enabled = True Then
	 TiltSens = TiltSens + 1
	 if TiltSens = 3 Then
	   Tilt = True
	   tilttxt.text="TILT"
       	If b2son Then Controller.B2SSetTilt 33,1
       	If b2son Then Controller.B2ssetdata 1, 0
	   playsound "tilt"
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

	LeftFlipper.RotateToStart
	DOF 101, DOFOff
	StopSound "Buzz"
	RightFlipper.RotateToStart
	DOF 102, DOFOff
	StopSound "Buzz1"
end sub

Sub addbonus
	bonus=bonus+1
	if bonus > 19 then bonus = 19
End sub


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

Sub RubberWheel_Hit()
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
	savevalue "Aspen", "hiscore", hisc
    savevalue "Aspen", "score1", score(1)
    savevalue "Aspen", "score2", score(2)
    savevalue "Aspen", "score3", score(3)
    savevalue "Aspen", "score4", score(4)
	savevalue "Aspen", "balls", balls
end sub

sub loadhs
    dim temp
	'temp = LoadValue("Aspen", "credit")
    'If (temp <> "") then credit = CDbl(temp)
    temp = LoadValue("Aspen", "hiscore")
    If (temp <> "") then hisc = CDbl(temp)
    'temp = LoadValue("Aspen", "match")
    'If (temp <> "") then matchnumb = CDbl(temp)
    temp = LoadValue("Aspen", "score1")
    If (temp <> "") then score(1) = CDbl(temp)
    temp = LoadValue("Aspen", "score2")
    If (temp <> "") then score(2) = CDbl(temp)
    temp = LoadValue("Aspen", "score3")
    If (temp <> "") then score(3) = CDbl(temp)
    temp = LoadValue("Aspen", "score4")
    If (temp <> "") then score(4) = CDbl(temp)
    temp = LoadValue("Aspen", "replays")
    If (temp <> "") then replays = CDbl(temp)
    temp = LoadValue("Aspen", "balls")
    If (temp <> "") then balls = CDbl(temp)
end sub

Sub Aspen_Exit()
	turnoff
	Savehs
	If B2SOn Then Controller.stop
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

Function AudioFade(tableobj) ' Fades between front and back of the table (for surround systems or 2x2 speakers, etc), depending on the Y position on the table. "aspen" is the name of the table
  Dim tmp
  tmp = tableobj.y * 2 / aspen.height-1
  If tmp > 0 Then
    AudioFade = Csng(tmp ^10)
  Else
    AudioFade = Csng(-((- tmp) ^10) )
  End If
End Function

Function AudioPan(tableobj) ' Calculates the pan for a tableobj based on the X position on the table. "aspen" is the name of the table
  Dim tmp
  tmp = tableobj.x * 2 / aspen.width-1
  If tmp > 0 Then
    AudioPan = Csng(tmp ^10)
  Else
    AudioPan = Csng(-((- tmp) ^10) )
  End If
End Function

Function Pan(ball) ' Calculates the pan for a ball based on the X position on the table. "aspen" is the name of the table
    Dim tmp
    tmp = ball.x * 2 / aspen.width-1
    If tmp > 0 Then
        Pan = Csng(tmp ^10)
    Else
        Pan = Csng(-((- tmp) ^10) )
    End If
End Function

Function AudioFade(ball) ' Can this be together with the above function ?
  Dim tmp
  tmp = ball.y * 2 / aspen.height-1
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

Const tnob = 2 ' total number of balls
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
  If aspen.VersionMinor > 3 OR aspen.VersionMajor > 10 Then
    PlaySound("fx_collide"), 0, Csng(velocity) ^2 / 200, Pan(ball1), 0, Pitch(ball1), 0, 0, AudioFade(ball1)
  Else
    PlaySound("fx_collide"), 0, Csng(velocity) ^2 / 200, Pan(ball1), 0, Pitch(ball1), 0, 0
  End if
End Sub


' Thalamus : Exit in a clean and proper way
Sub Aspen_exit()
  Controller.Pause = False
  Controller.Stop
End Sub


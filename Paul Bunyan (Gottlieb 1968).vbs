'	-------------------------------------------------
'	PAUL BUNYAN
'	-------------------------------------------------
'
'	Original table by D. Gottlieb & Co., 1968
'	VP8 version by PinballPerson/Leon Spalding
'	VP10 update version by BorgDog, 2015
'	thanks randr for the triangle post and zany for the gottlieb mini-flippers
'	DOF by BorgDog, fixed by arngrim ;)
'	-------------------------------------------------

' Thalamus 2018-07-24
' Added/Updated "Positional Sound Playback Functions" and "Supporting Ball & Sound Functions"
' No special SSF tweaks yet.

Option Explicit
Randomize

Const cGameName = "paulbunyan_1968"

Dim operatormenu, options
Dim bumperlitscore
Dim bumperoffscore
Dim balls
Dim replays
Dim Replay1Table(3)
Dim Replay2Table(3)
Dim Replay3Table(3)
Dim Replay4Table(3)
dim replay1
dim replay2
dim replay3
dim replay4
dim replaytext(3)
dim bip(5)
Dim hisc
Dim Controller
Dim maxplayers
Dim players
Dim player
Dim credit
Dim score(2)
Dim sreels(2)
Dim state
Dim tilt
Dim tiltsens
Dim target(9)
Dim abcscore
Dim ballinplay
Dim matchnumb
dim ballrenabled
dim rstep
dim lstep
Dim rep(2)
Dim rst
Dim eg
Dim scn
Dim scn1
Dim bell
Dim i,j, kl, aabc, objekt, light

On Error Resume Next
ExecuteGlobal GetTextFile("controller.vbs")
If Err Then MsgBox "Can't open controller.vbs"
On Error Goto 0

sub PaulBunyan_init
	LoadEM
	maxplayers=2
	Replay1Table(1)=3400
	Replay2Table(1)=4000
	Replay3Table(1)=4600
	Replay4Table(1)=5200
	Replay1Table(2)=4100
	Replay2Table(2)=4700
	Replay3Table(2)=5300
	Replay4Table(2)=5900
	Replay1Table(3)=5000
	Replay2Table(3)=5600
	Replay3Table(3)=6200
	Replay4Table(3)=6800
	replaytext(1)="3400, 4000, 4600, 5200"
	replaytext(2)="4100, 4700, 5300, 5900"
	replaytext(3)="5000, 5600, 6200, 6800"
	set sreels(1) = ScoreReel1
	set sreels(2) = ScoreReel2
	set bip(1) = bip1:set bip(2) = bip2:set bip(3) = bip3:set bip(4) = bip4:set bip(5) = bip5
	hideoptions
	player=1
	For each light in LaneLights:light.State = 0: Next
	For each light in BumperLights:light.State = 0: Next
	For each light in lights:light.State = 0: Next
	loadhs
	if hisc="" then hisc=1000
	hstxt.text=hisc
	gameover.state=1
	credittxt.setvalue(credit)
	if balls="" then balls=3
	if balls<>3 and balls<>5 then balls=3
	if replays="" then replays=1
	if replays<>1 and replays<>2 and replays<>3 then replays=1
	Replay1=Replay1Table(Replays)
	Replay2=Replay2Table(Replays)
	Replay3=Replay3Table(Replays)
	Replay4=Replay4Table(Replays)
	RepCard.image = "ReplayCard"&replays
	OptionBalls.image="OptionsBalls"&Balls
	OptionReplays.image="OptionsReplays"&replays
	bumperlitscore=10
	bumperoffscore=1
	if balls=3 then
		InstCard.image="InstCard3balls"
		abcscore = 100
	  else
		InstCard.image="InstCard5balls"
		abcscore = 50
	end if
    if kl="" then kl=0
    kicklights(kl).state=lightstateon
    kicklights(kl+3).state=lightstateon
	If B2SOn then
		setBackglass.enabled=true
		dim objekt : for each objekt in backdropstuff : objekt.visible = 0 : next
	End If
	if matchnumb="" then matchnumb=0
	matchtxt.text=matchnumb
	for i = 1 to maxplayers
		sreels(i).setvalue(score(i))
	next
	PlaySound "motor"
	tilt=false

	If credit>0 then DOF 134, DOFOn

End sub

sub setBackglass_timer
		Controller.B2ssetCredits Credit
		Controller.B2SSetData 6,1
		Controller.B2ssetMatch 34, Matchnumb
        Controller.B2SSetGameOver 35,1
        Controller.B2SSetScorePlayer 5, hisc
		me.enabled=False
end Sub

Sub PaulBunyan_KeyDown(ByVal keycode)

	if keycode=AddCreditKey then
		playsound "coinin"
		coindelay.enabled=true
    end if

    if keycode=StartGameKey and credit>0 then
	  if state=false then
		credit=credit-1
		if credit < 1 then DOF 134, DOFOff
		playsound "cluper"
		credittxt.setvalue(credit)
		ballinplay=1
		If B2SOn Then
			Controller.B2ssetCredits Credit
			Controller.B2sStartAnimation "Startup"
			Controller.B2ssetballinplay 32, Ballinplay
			Controller.B2ssetplayerup 30, 1
			Controller.B2ssetcanplay 31, 1
			Controller.B2SSetGameOver 0
			Controller.B2SSetScorePlayer 5, hisc
		End If
	    pup1.setvalue(1)
		tilt=false
		state=true
		playsound "initialize"
		players=1
		canplayreel.setvalue(players)
		rst=0
		resettimer.enabled=true
	  else if state=true and players < maxplayers and Ballinplay=1 then
		credit=credit-1
		if credit < 1 then DOF 134, DOFOff
		players=players+1
		canplayreel.setvalue(players)
		credittxt.setvalue(credit)
		If B2SOn then
			Controller.B2ssetCredits Credit
			Controller.B2ssetcanplay 31, players
		End If
		playsound "cluper"
	   end if
	  end if
	end if

	If keycode = PlungerKey Then
		Plunger.PullBack
		PlaySound "plungerpull",0,1,0.25,0.25
	End If
	If keycode=LeftFlipperKey and State = false and OperatorMenu=0 then
		OperatorMenuTimer.Enabled = true
	end if

	If keycode=LeftFlipperKey and State = false and OperatorMenu=1 then
		Options=Options+1
		If Options=4 then Options=1
		playsound "drop1"
		Select Case (Options)
			Case 1:
				Option1.visible=true
				Option3.visible=False
			Case 2:
				Option2.visible=true
				Option1.visible=False
			Case 3:
				Option3.visible=true
				Option2.visible=False
		End Select
	end if

	If keycode=RightFlipperKey and State = false and OperatorMenu=1 then
	  PlaySound "switch"
	  Select Case (Options)
		Case 1:
			if Balls=3 then
				Balls=5
				InstCard.image="InstCard5balls"
			  else
				Balls=3
				InstCard.image="InstCard3balls"
			end if
			OptionBalls.image = "OptionsBalls"&Balls
		Case 2:
			Replays=Replays+1
			if Replays>3 then
				Replays=1
			end if
			Replay1=Replay1Table(Replays)
			Replay2=Replay2Table(Replays)
			Replay3=Replay3Table(Replays)
			OptionReplays.image = "OptionsReplays"&replays
			repcard.image = "ReplayCard"&replays
		Case 3:
			OperatorMenu=0
			savehs
			HideOptions
	  End Select
	End If


  if tilt=false and state=true then
	If keycode = LeftFlipperKey Then
		LeftFlipper.RotateToEnd
		LeftFlipper1.RotateToEnd
		LeftFlipper2.RotateToEnd
		PlaySound SoundFXDOF("flipperup",101,DOFOn,DOFContactors), 0, .67, -0.05, 0.05
		PlaySound "Buzz",-1,.05,-0.05, 0.05
	End If

	If keycode = RightFlipperKey Then
		RightFlipper.RotateToEnd
		RightFlipper1.RotateToEnd
		RightFlipper2.RotateToEnd
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

	If keycode = MechanicalTilt Then
		mechchecktilt
	End If

  end if
End Sub

Sub OperatorMenuTimer_Timer
	OperatorMenu=1
	Displayoptions
	Options=1
End Sub

Sub DisplayOptions
	OptionsBack.visible = true
	Option1.visible = True
	OptionBalls.visible = True
    OptionReplays.visible = True
End Sub

Sub HideOptions
	for each objekt In OptionMenu
		objekt.visible = false
	next
End Sub


Sub PaulBunyan_KeyUp(ByVal keycode)

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
		LeftFlipper1.RotateToStart
		LeftFlipper2.RotateToStart
		PlaySound SoundFXDOF("flipperdown",101,DOFOff,DOFContactors), 0, 1, -0.05, 0.05
		StopSound "Buzz"
	End If

	If keycode = RightFlipperKey Then
		RightFlipper.RotateToStart
		RightFlipper1.RotateToStart
		RightFlipper2.RotateToStart
		PlaySound SoundFXDOF("flipperdown",102,DOFOff,DOFContactors), 0, 1, 0.05, 0.05
		StopSound "Buzz1"
	End If
   End if
End Sub

sub flippertimer_timer()
	LFlip.RotY = LeftFlipper.CurrentAngle-90
	RFlip.RotY = RightFlipper.CurrentAngle-90
	LFlip1.RotY = LeftFlipper.CurrentAngle-90
	RFlip1.RotY = RightFlipper.CurrentAngle-90
	LFlip2.RotY = LeftFlipper.CurrentAngle-90
	RFlip2.RotY = RightFlipper.CurrentAngle-90
	PrimGate1.Rotz = gate.CurrentAngle
end sub


Sub PairedlampTimer_timer
	bumper1light1.state = bumper1light.state
	bumper2light1.state = bumper2light.state
	bumper3light1.state = bumper3light.state
end sub

sub coindelay_timer
	addcredit
    coindelay.enabled=false
end sub

sub resettimer_timer
    rst=rst+1
	for i = 1 to maxplayers
		sreels(i).resettozero
    next
    If B2SOn then
		for i = 1 to maxplayers
		  Controller.B2SSetScorePlayer i, score(i)
		  Controller.B2SSetScoreRollover 24 + i, 0
		next
	End If
    if rst=18 then
		playsound "kickerkick"
    end if
    if rst=22 then
		newgame
		resettimer.enabled=false
    end if
end sub

Sub addcredit
      credit=credit+1
	  DOF 134, DOFOn
      if credit>15 then credit=15
	  credittxt.setvalue(credit)
	  If B2SOn Then Controller.B2ssetCredits Credit
      playsound "click"
	  savehs
End sub

Sub Drain_Hit()
	DOF 130, DOFPulse
	PlaySound "drain",0,1,0,0.25
	Drain.DestroyBall
	me.timerenabled=1
End Sub

Sub Drain_timer
	scorebonus.enabled=true
	me.timerenabled=0
End Sub

sub ballhome_hit
	ballrenabled=1
end sub

sub ballhome_unhit
	DOF 135, DOFPulse
end sub

sub ballout_hit
'    bb.isdropped=false
end sub

sub ballrel_hit
	if ballrenabled=1 then
		shootagain.state=lightstateoff
		playsound "ballrelease"
		ballrenabled=0
	end if
end sub

sub scorebonus_timer
   	 if shootagain.state=lightstateon and tilt=false then
	    newball
 	    ballreltimer.enabled=true
     else
	  if players=1 or player=players then
		player=1
		If B2SOn then Controller.B2ssetplayerup 30, player
	    pup1.setvalue(1)
	    pup2.setvalue(0)
		nextball
	  else
		player=player+1
		If B2SOn then Controller.B2ssetplayerup 30, player
		pup2.setvalue(1)
		pup1.setvalue(0)
		nextball
	  end if
	 end if
	 scorebonus.enabled=false
End sub


sub newgame
	bumper1.force=10
	bumper2.force=10
	bumper3.force=10
	player=1
    pup1.setvalue(1)
    score(1)=0
	score(2)=0
	If B2SOn then
	  for i = 1 to maxplayers
		Controller.B2SSetScorePlayer i, score(i)
	  next
	End If
    eg=0
    rep(1)=0
	rep(2)=0
    rulBall.state=0
	shootagain.state=lightstateoff
	for each light in bumperlights:light.state=0:next
	for each light in lanelights:light.state=1:next
	for each light in outlights:light.state=0: next
	for each light in lights:light.state=0: next
	for each light in abclights:light.state=1:next
    lklight.state=lightstateoff
    rklight.state=lightstateoff
    lol2.state=lightstateon
    rol2.state=lightstateon
    gameover.state=0
    tilttxt.text=" "
    If B2SOn then
		Controller.B2SSetGameOver 35,0
		Controller.B2SSetTilt 33,0
		Controller.B2SSetMatch 34,0
	End If
	bipoff
    bip(ballinplay).state=1
	matchtxt.text=" "
    BallRelease.CreateBall
   	BallRelease.kick 135,4,0
	playsound SoundFXDOF("kickerkick",129,DOFPulse,DOFContactors)
end sub

sub bipoff
	for i = 1 to 5
		bip(i).state = 0
	next
end sub

sub newball
	for each light in lights: light.state=0: next
	for each light in abclights:light.state=1:next
	for each light in bumperlights:light.state=0:next
	lklight.state=lightstateoff
    rklight.state=lightstateoff
    aabc=0
End Sub


sub nextball
    if tilt=true then
	  bumper1.force=10
	  bumper2.force=10
	  bumper3.force=10
      tilt=false
      tilttxt.text=" "
		If B2SOn then
			Controller.B2SSetTilt 33,0
			Controller.B2ssetdata 1, 1
		End If
    end if
	rulball.state=0
	if player=1 then ballinplay=ballinplay+1
	if ballinplay>balls then
		playsound "GameOver"
		eg=1
		ballreltimer.enabled=true
	  else
		if state=true then
		  newball
		  ballreltimer.enabled=true
		end if
		bipoff
		bip(ballinplay).state=1
	end if
End Sub

sub ballreltimer_timer
  if eg=1 then
	  turnoff
      matchnum
	  bipoff
	  state=false
	  gameover.state=1
	  tilttxt.timerenabled=1
	  canplayreel.setvalue(0)
	  pup1.setvalue(0)
	  pup2.setvalue(0)
	  for i=1 to maxplayers
		if score(i)>hisc then hisc=score(i)
	  next
	  hstxt.text=hisc
	  savehs
	  If B2SOn then
        Controller.B2SSetGameOver 35,1
        Controller.B2ssetballinplay 32, 0
	    Controller.B2sStartAnimation "EOGame"
	    Controller.B2SSetScorePlayer 5, hisc
	    Controller.B2ssetcanplay 31, 0
	    Controller.B2ssetcanplay 30, 0
	  End If
	  For each light in LaneLights:light.State = 0: Next
	  ballreltimer.enabled=false
  else
    BallRelease.CreateBall
	BallRelease.kick 135,4,0
	playsound SoundFXDOF("kickerkick",129,DOFPulse,DOFContactors)
    ballreltimer.enabled=false
  end if
end sub

sub matchnum
	matchnumb=INT (RND*10)
	matchtxt.text=matchnumb
	If B2SOn then Controller.B2SSetMatch 34,Matchnumb
	For i=1 to players
		if (matchnumb)=(score(i) mod 10) then
		  addcredit
		playsound SoundFXDOF("knock",131,DOFPulse,DOFKnocker)
		DOF 132,DOFPulse
	    end if
    next
end sub

'********** Bumpers

Sub Bumper1_Hit
   if tilt=false then
	playsound SoundFXDOF("fx_bumper4",107,DOFPulse,DOFContactors)
	DOF 108,DOFPulse
	if bumper1light.state=1 then
		addscore bumperlitscore
	  else
		addscore bumperoffscore
	end if
	me.timerenabled=1
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


Sub Bumper2_Hit
   if tilt=false then
	playsound SoundFXDOF("fx_bumper4",109,DOFPulse,DOFContactors)
	DOF 110,DOFPulse
	if bumper2light.state=1 then
		addscore bumperlitscore
	  else
		addscore bumperoffscore
	end if
	me.timerenabled=1
   end if
End Sub

Sub Bumper2_timer
	Bumper2Ring.enabled=0
	BumperRing2.transz=BumperRing2.transz-4
	if BumperRing2.transz=-36 then
		Bumper2Ring.enabled=1
		me.timerenabled=0
	end if
End Sub

Sub Bumper2Ring_timer
	BumperRing2.transz=BumperRing2.transz+4
	If BumperRing2.transz=0 then Bumper2Ring.enabled=0
End sub

Sub Bumper3_Hit
   if tilt=false then
	playsound SoundFXDOF("fx_bumper4",111,DOFPulse,DOFContactors)
	DOF 112,DOFPulse
	if bumper3light.state=1 then
		addscore bumperlitscore
	  else
		addscore bumperoffscore
	end if
	me.timerenabled=1
   end if
End Sub

Sub Bumper3_timer
	Bumper3Ring.enabled=0
	BumperRing3.transz=BumperRing3.transz-4
	if BumperRing3.transz=-36 then
		Bumper3Ring.enabled=1
		me.timerenabled=0
	end if
End Sub

Sub Bumper3Ring_timer
	BumperRing3.transz=BumperRing3.transz+4
	If BumperRing3.transz=0 then Bumper3Ring.enabled=0
End sub

sub FlashBumpers
	For each light in BumperLights
	  if light.state = 1 then
		Light.State=0
	  else
		Light.state=1
	  end if
	next
	FlashB.enabled=1
end sub

sub FlashB_timer
	For each light in BumperLights
	  if light.state = 1 then
		Light.State=0
	  else
		Light.state=1
	  end if
	next
	FlashB.enabled=0
end sub

Sub RightSlingShot_Slingshot
	playsound SoundFXDOF("left_slingshot",103,DOFPulse,DOFContactors), 0, 1, 0.05, 0.05
	DOF 104, DOFPulse
	if (rkl.state)=lightstateon then
		addscore 10
	  else
		addscore 1
	end if
    RSling.Visible = 0
    RSling1.Visible = 1
    sling1.TransZ = -8
    RStep = 1
    RightSlingShot.TimerEnabled = 1
End Sub

Sub RightSlingShot_Timer
    Select Case RStep
        Case 3:RSLing1.Visible = 0:RSLing2.Visible = 1:sling1.TransZ = -8
        Case 4:sling1.TransZ = 0:RSLing2.Visible = 0:RSLing.Visible = 1:RightSlingShot.TimerEnabled = 0
    End Select
    RStep = RStep + 1
End Sub

Sub LeftSlingShot_Slingshot
	playsound SoundFXDOF("right_slingshot",105,DOFPulse,DOFContactors), 0, 1, -0.05, 0.05
	DOF 106, DOFPulse
	if (lkl.state)=lightstateon then
		addscore 10
	  else
		addscore 1
	end if
    LSling.Visible = 0
    LSling1.Visible = 1
    sling2.TransZ = -8
    LStep = 1
    LeftSlingShot.TimerEnabled = 1
End Sub

Sub LeftSlingShot_Timer
    Select Case LStep
        Case 3:LSLing1.Visible = 0:LSLing2.Visible = 1:sling2.TransZ = -8
        Case 4:sling2.TransZ = 0:LSLing2.Visible = 0:LSLing.Visible = 1:LeftSlingShot.TimerEnabled = 0
    End Select
    LStep = LStep + 1
End Sub


sub tlr_hit
    addscore 1
end sub

sub trr_hit
    addscore 1
end sub

sub mlr_hit
    addscore 1
end sub

sub mrr_hit
    addscore 1
end sub

sub ru_hit
    if (rulball.state)=lightstateon then shootagain.state=lightstateon:rulball.state=lightstateoff
    if (rul500.state)=lightstateon then
		addscore 500
	  else
		addscore 100
    end if
end sub

sub umlt_hit
	DOF 116, DOFPulse
    if (umll.state)=lightstateon then
    addscore 100
    else
    addscore 10
    end if
end sub

sub umrt_hit
	DOF 118, DOFPulse
    if (umrl.state)=lightstateon then
    addscore 100
    else
    addscore 10
    end if
end sub

sub lmlt_hit
	DOF 117, DOFPulse
    if (lmll.state)=lightstateon then
    addscore 100
    else
    addscore 10
    end if
end sub

sub lmrt_hit
	DOF 119, DOFPulse
    if (lmrl.state)=lightstateon then
    addscore 100
    else
    addscore 10
    end if
end sub

sub ult_hit
	DOF 124, DOFPulse
    if (ultl.state)=lightstateon then
    rulball.state=lightstateon
    addscore 100
    else
    addscore 100
    end if
end sub

sub urt_hit
	DOF 125, DOFPulse
    if (urtl.state)=lightstateon then
    rulball.state=lightstateon
    addscore 100
    else
    addscore 100
    end if
end sub

sub mlt_hit
	DOF 126, DOFPulse
    addscore 50
end sub

sub mrt_hit
	DOF 126, DOFPulse
    addscore 50
end sub

sub lout1_hit
	DOF 120, DOFPulse
    if (lol1.state)=lightstateon then
		addscore 200
    else
		addscore 50
    end if
end sub

sub lout2_hit
	DOF 121, DOFPulse
    if (lol2.state)=lightstateon then
		addscore 200
    else
		addscore 50
    end if
end sub

sub rout1_hit
	DOF 122, DOFPulse
    if (rol1.state)=lightstateon then
		addscore 200
    else
		addscore 50
    end if
end sub

sub rout2_hit
	DOF 123, DOFPulse
    if (rol2.state)=lightstateon then
		addscore 200
    else
		addscore 50
    end if
end sub

sub lk_hit
    if (lklight.state)=lightstateon then
		if (lkl1.state)=lightstateon then addscore 200
		if (lkl2.state)=lightstateon then addscore 400
		if (lkl3.state)=lightstateon then addscore 600
      else
		if (lkl1.state)=lightstateon then addscore 50
		if (lkl2.state)=lightstateon then addscore 100
		if (lkl3.state)=lightstateon then addscore 50:addscore 100
    end if
    kicktimer.enabled=true
end sub

sub rk_hit
    if (rklight.state)=lightstateon then
		if (rkl1.state)=lightstateon then addscore 200
		if (rkl2.state)=lightstateon then addscore 400
		if (rkl3.state)=lightstateon then addscore 600
    else
		if (rkl1.state)=lightstateon then addscore 50
		if (rkl2.state)=lightstateon then addscore 100
		if (rkl3.state)=lightstateon then addscore 50:addscore 100
    end if
    kicktimer.enabled=true
end sub

sub kicktimer_timer
	playsound SoundFXDOF("kickerkick",127,DOFPulse,DOFContactors)
	playsound SoundFXDOF("kickerkick",128,DOFPulse,DOFContactors)
	DOF 132, DOFPulse
    lk.kick 120,20
    rk.kick 240,20
    kicktimer.enabled=false
end sub

sub tll_hit   '***** top A rollover
	DOF 113, DOFPulse
	if (tlll.state)=lightstateon and abcl.state=1 then
		addscore 500
		awards
    else
		addscore abcscore
    end if
  if balls = 5 then
    if (tlll.state)=lightstateon then
		tlll.state=lightstateoff
		lltl.state=lightstateoff
		umrl.state=lightstateon
		lkl.state=lightstateon
		bumper1light.state=lightstateon
    end if
  else
	if (tlll.state)=lightstateon then
		tlll.state=lightstateoff
		lltl.state=lightstateoff
		umrl.state=lightstateon
		lkl.state=lightstateon
		bumper1light.state=lightstateon
    end if
	if tlll.state + tcll.state + trll.state=0 then
		lklight.state=1
		rklight.state=1
		abcl.state=1
		abclights(aabc).state=lightstateon
		abclights(aabc+3).state=lightstateon
		urtl.state=1
		ultl.state=1
	end if
  end if

end sub

sub tcl_hit   '***** top B rollover
	DOF 114, DOFPulse
	if (tcll.state)=lightstateon and abcl.state=1 then
		addscore 500
		awards
      else
		addscore abcscore
    end if
  if balls = 5 then
    if (tcll.state)=lightstateon and (tlll.state)=lightstateoff then
		tcll.state=lightstateoff
		lctl.state=lightstateoff
		lmll.state=lightstateon
		lmrl.state=lightstateon
		bumper2light.state=lightstateon
    end if
  else
	if (tcll.state)=lightstateon then
		tcll.state=lightstateoff
		lctl.state=lightstateoff
		lmll.state=lightstateon
		lmrl.state=lightstateon
		bumper2light.state=lightstateon
    end if
	if tlll.state + tcll.state + trll.state=0 then
		lklight.state=1
		rklight.state=1
		abcl.state=1
		abclights(aabc).state=lightstateon
		abclights(aabc+3).state=lightstateon
		urtl.state=1
		ultl.state=1
	end if
  end if

end sub

sub trl_hit   '***** top C rollover
	DOF 115, DOFPulse
    if (trll.state)=lightstateon and abcl.state=1 then
		addscore 500
		awards
    else
		addscore abcscore
    end if
  if balls = 5 then
    if (trll.state)=lightstateon and (tcll.state)=lightstateoff then
		trll.state=lightstateoff
		lrtl.state=lightstateoff
		umll.state=lightstateon
		rkl.state=lightstateon
		bumper3light.state=lightstateon
		abcl.state=lightstateon
		lklight.state=lightstateon
		urtl.state=lightstateon
		abclights(aabc).state=lightstateon
		abclights(aabc+3).state=lightstateon
    end if
  else
	if (trll.state)=lightstateon then
		trll.state=lightstateoff
		lrtl.state=lightstateoff
		umll.state=lightstateon
		rkl.state=lightstateon
		bumper3light.state=lightstateon
    end if
	if tlll.state + tcll.state + trll.state=0 then
		lklight.state=1
		rklight.state=1
		abcl.state=1
		abclights(aabc).state=lightstateon
		abclights(aabc+3).state=lightstateon
		urtl.state=1
		ultl.state=1
	end if
  end if

end sub

sub llt_hit   '***** lower A target
	DOF 126, DOFPulse
    if (lltl.state)=lightstateon and abcl.state=1 then
		addscore 500
		awards
    else
		addscore abcscore
    end if
  if balls = 5 then
    if (lltl.state)=lightstateon then
		tlll.state=lightstateoff
		lltl.state=lightstateoff
		umrl.state=lightstateon
		lkl.state=lightstateon
		bumper1light.state=lightstateon
    end if
  else
	if (lltl.state)=lightstateon then
		tlll.state=lightstateoff
		lltl.state=lightstateoff
		umrl.state=lightstateon
		lkl.state=lightstateon
		bumper1light.state=lightstateon
    end if
	if tlll.state + tcll.state + trll.state=0 then
		lklight.state=1
		rklight.state=1
		abcl.state=1
		abclights(aabc).state=lightstateon
		abclights(aabc+3).state=lightstateon
		urtl.state=1
		ultl.state=1
	end if
  end if

end sub

sub lct_hit    '***** lower B target
	DOF 126, DOFPulse
    if (lctl.state)=lightstateon and abcl.state=1 then
		addscore 500
		awards
    else
		addscore abcscore
    end if
  if balls = 5 then
    if (lctl.state)=lightstateon and (lltl.state)=lightstateoff then
		tcll.state=lightstateoff
		lctl.state=lightstateoff
		lmll.state=lightstateon
		lmrl.state=lightstateon
		bumper2light.state=lightstateon
    end if
  else
	if (lctl.state)=lightstateon then
		tcll.state=lightstateoff
		lctl.state=lightstateoff
		lmll.state=lightstateon
		lmrl.state=lightstateon
		bumper2light.state=lightstateon
    end if
	if tlll.state + tcll.state + trll.state=0 then
		lklight.state=1
		rklight.state=1
		abcl.state=1
		abclights(aabc).state=lightstateon
		abclights(aabc+3).state=lightstateon
		urtl.state=1
		ultl.state=1
	end if
  end if

end sub

sub lrt_hit   '***** lower C target
	DOF 126, DOFPulse
    if (lrtl.state)=lightstateon and abcl.state=1 then
		addscore 500
		awards
    else
		addscore abcscore
    end if
  if balls = 5 then
    if (lrtl.state)=lightstateon and (lctl.state)=lightstateoff then
		trll.state=lightstateoff
		lrtl.state=lightstateoff
		umll.state=lightstateon
		rkl.state=lightstateon
		bumper3light.state=lightstateon
		abcl.state=lightstateon
		lklight.state=lightstateon
		urtl.state=lightstateon
		abclights(aabc).state=lightstateon
		abclights(aabc+3).state=lightstateon
    end if
  else
	if (lrtl.state)=lightstateon then
		trll.state=lightstateoff
		lrtl.state=lightstateoff
		umll.state=lightstateon
		rkl.state=lightstateon
		bumper3light.state=lightstateon
    end if
	if tlll.state + tcll.state + trll.state=0 then
		lklight.state=1
		rklight.state=1
		abcl.state=1
		abclights(aabc).state=lightstateon
		abclights(aabc+3).state=lightstateon
		urtl.state=1
		ultl.state=1
	end if
  end if

end sub

sub awards
	for each light in awardlights:light.state=0: next
	for each light in bumperlights:light.state=0: next
    for each light in abclights: light.state = lightstateon: next
    rul500.state=lightstateon
end sub

sub klupdate
    for each light in kicklights: light.state = lightstateoff: next
	kicklights(kl).state=lightstateon
	kicklights(kl+3).state=lightstateon
end sub

sub updateabc
    for each light in abclights: light.state = lightstateoff: next
  	abclights(aabc).state=lightstateon
	abclights(aabc+3).state=lightstateon
end sub

sub addscore(points)
  if tilt=false then
	sreels(player).addvalue(points)
	If B2SOn Then Controller.B2SSetScorePlayer player, score(player) + points
    if points = 100 then
		PlaySound SoundFXDOF("gigibump4loud",142,DOFPulse,DOFChimes)
    end if
    if points = 10 then
		PlaySound SoundFXDOF("gigibump3loud",141,DOFPulse,DOFChimes)
		aabc=aabc+1
		if aabc>2 then aabc=0
    end if
    if points = 1 then
		if lol2.state=lightstateon then
			lol2.state=lightstateoff
			rol1.state=lightstateon
			lol1.state=lightstateon
			rol2.state=lightstateoff
		  else
			lol2.state=lightstateon
			rol1.state=lightstateoff
			lol1.state=lightstateoff
			rol2.state=lightstateon
		end if
    end if
    if points = 400 then
		playsound "motorshort1s"
		scn=3
		bell=100
    end if
    if points = 200 then
		playsound "motorshort1s"
		scn=2
		bell=100
    end if
    if points = 500 then
		playsound "motorshort1s"
		scn=5
		bell=100
    end if
    if points = 50 then
		playsound "motorshort1s"
		scn=5
		bell=10
	end if
	if points = 600 then
		playsound "motorshort1s"
		scn=6
		bell=100
	end if
    if points<>1 and points<>10 and points<>100 then
		scn1=0
		scntimer.enabled=true
    end if
    score(player)=score(player)+points
    if score(player)=>replay1 and rep(player)=0 then
		addcredit
		rep(player)=1
		PlaySound SoundFXDOF("knock",131,DOFPulse,DOFKnocker)
		DOF 132, DOFPulse
    end if
    if score(player)=>replay2 and rep(player)=1 then
		addcredit
		rep(player)=2
		PlaySound SoundFXDOF("knock",131,DOFPulse,DOFKnocker)
		DOF 132, DOFPulse
    end if
    if score(player)=>replay3 and rep(player)=2 then
		addcredit
		rep(player)=3
		PlaySound SoundFXDOF("knock",131,DOFPulse,DOFKnocker)
		DOF 132, DOFPulse
    end if
    if score(player)=>replay4 and rep(player)=3 then
		addcredit
		rep(player)=4
		PlaySound SoundFXDOF("knock",131,DOFPulse,DOFKnocker)
		DOF 132, DOFPulse
    end if
  end if
end sub

sub scntimer_timer
    scn1=scn1 + 1
    if bell=10 then
		PlaySound SoundFXDOF("gigibump3loud",141,DOFPulse,DOFChimes)
		kl=kl+1
		if kl>2 then kl=0
		klupdate
    end if
    if bell=100 then PlaySound SoundFXDOF("gigibump4loud",142,DOFPulse,DOFChimes)
    if scn1=scn then scntimer.enabled=false
end sub


Sub CheckTilt
	If Tilttimer.Enabled = True Then
	 TiltSens = TiltSens + 1
	 if TiltSens = 3 Then
	   Tilt = True
	   tilttxt.text="TILT"
       	If B2SOn Then Controller.B2SSetTilt 33,1
       	If B2SOn Then Controller.B2ssetdata 1, 0
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
	   tilttxt.text="TILT"
       	If B2SOn Then Controller.B2SSetTilt 33,1
       	If B2SOn Then Controller.B2ssetdata 1, 0
	   playsound "tilt"
	   turnoff
End Sub

Sub Tilttimer_Timer()
	Tilttimer.Enabled = False
End Sub

sub turnoff
    bumper1.force=0
    bumper2.force=0
	bumper3.force=0
  	LeftFlipper.RotateToStart
	LeftFlipper1.RotateToStart
	LeftFlipper2.RotateToStart
	DOF 101, DOFOff
	StopSound "Buzz"
	RightFlipper.RotateToStart
	RightFlipper1.RotateToStart
	RightFlipper2.RotateToStart
	DOF 102, DOFOff
	StopSound "Buzz1"
end sub

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

Sub Rubberwheel_Hit
 	dim finalspeed
  	finalspeed=SQR(activeball.velx * activeball.velx + activeball.vely * activeball.vely)
 	If finalspeed > 20 then
		PlaySound "fx_rubber2", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
	End if
	If finalspeed >= 6 AND finalspeed <= 20 then
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

Sub LeftFlipper1_Collide(parm)
 	RandomSoundFlipper()
End Sub

Sub RightFlipper1_Collide(parm)
 	RandomSoundFlipper()
End Sub

Sub LeftFlipper2_Collide(parm)
 	RandomSoundFlipper()
End Sub

Sub RightFlipper2_Collide(parm)
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

    savevalue "PaulBunyan", "credit", credit
    savevalue "PaulBunyan", "hiscore", hisc
    savevalue "PaulBunyan", "match", matchnumb
    savevalue "PaulBunyan", "score1", score(1)
    savevalue "PaulBunyan", "score2", score(2)
	savevalue "PaulBunyan", "balls", balls
	savevalue "PaulBunyan", "replays", replays

end sub

sub loadhs
    dim temp
	temp = LoadValue("PaulBunyan", "credit")
    If (temp <> "") then credit = CDbl(temp)
    temp = LoadValue("PaulBunyan", "hiscore")
    If (temp <> "") then hisc = CDbl(temp)
    temp = LoadValue("PaulBunyan", "match")
    If (temp <> "") then matchnumb = CDbl(temp)
    temp = LoadValue("PaulBunyan", "score1")
    If (temp <> "") then score(1) = CDbl(temp)
    temp = LoadValue("PaulBunyan", "score2")
    If (temp <> "") then score(2) = CDbl(temp)
    temp = LoadValue("PaulBunyan", "balls")
    If (temp <> "") then balls = CDbl(temp)
    temp = LoadValue("PaulBunyan", "replays")
    If (temp <> "") then replays = CDbl(temp)
end sub

Sub PaulBunyan_Exit()
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

Function AudioFade(tableobj) ' Fades between front and back of the table (for surround systems or 2x2 speakers, etc), depending on the Y position on the table. "PaulBunyan" is the name of the table
  Dim tmp
  tmp = tableobj.y * 2 / PaulBunyan.height-1
  If tmp > 0 Then
    AudioFade = Csng(tmp ^10)
  Else
    AudioFade = Csng(-((- tmp) ^10) )
  End If
End Function

Function AudioPan(tableobj) ' Calculates the pan for a tableobj based on the X position on the table. "PaulBunyan" is the name of the table
  Dim tmp
  tmp = tableobj.x * 2 / PaulBunyan.width-1
  If tmp > 0 Then
    AudioPan = Csng(tmp ^10)
  Else
    AudioPan = Csng(-((- tmp) ^10) )
  End If
End Function

Function Pan(ball) ' Calculates the pan for a ball based on the X position on the table. "PaulBunyan" is the name of the table
    Dim tmp
    tmp = ball.x * 2 / PaulBunyan.width-1
    If tmp > 0 Then
        Pan = Csng(tmp ^10)
    Else
        Pan = Csng(-((- tmp) ^10) )
    End If
End Function

Function AudioFade(ball) ' Can this be together with the above function ?
  Dim tmp
  tmp = ball.y * 2 / PaulBunyan.height-1
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
  If PaulBunyan.VersionMinor > 3 OR PaulBunyan.VersionMajor > 10 Then
    PlaySound("fx_collide"), 0, Csng(velocity) ^2 / 200, Pan(ball1), 0, Pitch(ball1), 0, 0, AudioFade(ball1)
  Else
    PlaySound("fx_collide"), 0, Csng(velocity) ^2 / 200, Pan(ball1), 0, Pitch(ball1), 0, 0
  End if
End Sub


' Thalamus : Exit in a clean and proper way
Sub PaulBunyan_exit()
  Controller.Pause = False
  Controller.Stop
End Sub


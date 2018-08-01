'	-------------------------------------------------
'	ROLLER COASTER - Gottlieb 1971
'	-------------------------------------------------
'	
'	by HauntFreaks and BorgDog, release 2016
'
'	Layer usage
'		1 - most stuff
'		2 - triggers
'		4 - options menu
'		6 - GI lighting
'		7 - plastics
'		8 - insert and bumper lighting
'		
'	
'	-------------------------------------------------

Option Explicit
Randomize

Const cGameName = "RollerCoaster_1971"
Const Ballsize = 50		

Dim operatormenu, options
Dim bumperlitscore
Dim bumperoffscore
Dim balls
Dim replays
Dim Replay1Table(3), Replay2Table(3), Replay3Table(3)
dim replay1, replay2, replay3
dim replaytext(3)
Dim Add10, Add100, Add1000
Dim hisc
Dim maxplayers
Dim players
Dim player
Dim credit
Dim score(2)
Dim sreels(2)
Dim p100k(2)
Dim cplay(2)
Dim Pups(2)
Dim state
Dim tilt
Dim tiltsens
Dim target(9)
Dim ballinplay
Dim matchnumb
dim rstep, lstep
Dim rep(2)
Dim rst
Dim eg
Dim bell
Dim i,j, ii, objekt, light
Dim awardcheck
Dim SpinArrow, SpinPos, Spin, BallSpeed, Count3, TempVelX, TempVely

Spin=Array(500,50,50,100,100,50,50,200,200,50,50,300,300,50,50,400,400,50,50,500)
SpinArrow=Array(9,27,45,63,81,99,117,135,153,171,189,207,225,243,261,279,297,315,333,351)

On Error Resume Next
ExecuteGlobal GetTextFile("controller.vbs")
If Err Then MsgBox "Can't open controller.vbs"
On Error Goto 0

sub RollerCoaster_init
	LoadEM
	maxplayers=2
	Replay1Table(1)=2100
	Replay2Table(1)=3200
	Replay3Table(1)=3700
	Replay1Table(2)=2400
	Replay2Table(2)=3800
	Replay3Table(2)=4600
	Replay1Table(3)=3900
	Replay2Table(3)=5300
	Replay3Table(3)=6100
	set sreels(1) = ScoreReel1
	set sreels(2) = ScoreReel2
	set Pups(1) = Pup1
	set Pups(2) = Pup2
'	hideoptions
	turnoff
	player=1
	For each light in BumperLights:light.State = 0: Next
	For each light in bonuslights:light.State = 0: Next
	hisc=4000
	balls=5
	replays=3
	spinpos = INT(RND*20)
	HSA1=2
	HSA2=18
	HSA3=9

	loadhs
	UpdatePostIt	'UPDATE HIGH SCORE STICKY
	bip_game.state=1
	gamov.text="Game Over"
	gamov.timerenabled=1
	tilttxt.timerenabled=1
	credittxt.setvalue(credit)
	For i=0 To 19
		Spinner(i).IsDropped=True
		Spinner(i+20).IsDropped=True
	Next
	Spinner(SpinPos).IsDropped=False
	Spinner(SpinPos+20).IsDropped=False
'	Spinner(SpinPos+40).Visible=false
	ArrowPrimitive.ObjRotZ=SpinArrow(SpinPos)
	Replay1=Replay1Table(Replays)
	Replay2=Replay2Table(Replays)
	Replay3=Replay3Table(Replays)
	RepCard.image = "ReplayCard"&replays
	OptionBalls2.image="OptionsBalls"&Balls
	OptionReplays2.image="OptionsReplays"&replays
	bumperoffscore=1
	if balls=3 then	
		InstCard.image="InstCard3balls"
	  else
		InstCard.image="InstCard5balls"
	end if
	If B2SOn then
		setBackglass.enabled=true
		for each objekt in backdropstuff : objekt.visible = 0 : next
	End If
	if matchnumb="" or matchnumb<0 or matchnumb>9 then matchnumb=0
	if matchnumb=0 then
		matchtxt.text="0"
	  else
		matchtxt.text=matchnumb
	end if
	for i = 1 to maxplayers
		sreels(i).setvalue(score(i))
	next
	for i = 1 to 5: EVAL("bip"&i).state=0: Next
	PlaySound "motor"
	tilt=false
    Drain.CreateBall
	If credit > 1 Then DOF 117, DOFOn
End sub

sub setBackglass_timer
	Controller.B2ssetCredits Credit
	if matchnumb=0 then
		Controller.B2ssetMatch 34, 10
	  else
		Controller.B2ssetMatch 34, Matchnumb
	end if
    Controller.B2SSetGameOver 35,1

	me.enabled=0
end Sub


sub gamov_timer
	if state=false then
		If B2SOn then Controller.B2SSetGameOver 35,0			
		gamov.text=""
		gtimer.enabled=true
	end if
	gamov.timerenabled=0
end sub

sub gtimer_timer
	if state=false then
		gamov.text="Game Over"
		If B2SOn then Controller.B2SSetGameOver 35,1
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

Sub RollerCoaster_KeyDown(ByVal keycode)

' start roth's ball control
	if keycode = 46 then' C Key
		If contball = 1 Then
			contball = 0
		  Else
			contball = 1
		End If
	End If

	if keycode = 48 then 'B Key
		If bcboost = 1 Then
			bcboost = bcboostmulti
		  Else
			bcboost = 1
		End If
	End If

	if keycode = 203 then Cleft = 1' Left Arrow

	if keycode = 200 then Cup = 1' Up Arrow

	if keycode = 208 then Cdown = 1' Down Arrow

	if keycode = 205 then Cright = 1' Right Arrow

' end roth's ball control
   
	if keycode=AddCreditKey then
		if credit+1<16 then
			playsound "coinin" 
		 Else
			playsound "coin3"
		end if
		DOF 117, DOFOn
		coindelay.enabled=true 
    end if

    if keycode=StartGameKey and credit>0 And Not HSEnterMode=true then
	  if state=false then
		credit=credit-1
		If credit < 1 Then DOF 117, DOFOff
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

		End If
	    pups(1).state=1
		tilt=false
		state=true
		playsound "initialize" 
		players=1
		canplayreel.setvalue(players)
		rst=0
		newgame.enabled=true
	  else if state=true and players < maxplayers and Ballinplay=1 then
		credit=credit-1
		If credit < 1 Then DOF 117, DOFOff
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

	If HSEnterMode Then HighScoreProcessKey(keycode)

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
		playsound "target"
		Select Case (Options)
			Case 1:
				Option1a.visible=true
				Option3a.visible=False
			Case 2:
				Option2a.visible=true
				Option1a.visible=False
			Case 3:
				Option3a.visible=true
				Option2a.visible=False
		End Select
	end if

	If keycode=RightFlipperKey and State = false and OperatorMenu=1 then
	  PlaySound "metalhit2"
	  Select Case (Options)
		Case 1:
			if Balls=3 then
				Balls=5
				InstCard.image="InstCard5balls"
			  else
				Balls=3
				InstCard.image="InstCard3balls"
			end if
			OptionBalls2.image = "OptionsBalls"&Balls     
		Case 2:
			Replays=Replays+1
			if Replays>3 then
				Replays=1
			end if
			Replay1=Replay1Table(Replays)
			Replay2=Replay2Table(Replays)
			OptionReplays2.image = "OptionsReplays"&replays
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
  end if  
End Sub

Sub OperatorMenuTimer_Timer
	OperatorMenu=1
	Displayoptions
	Options=1
End Sub

Sub DisplayOptions
	OptionsBack2.visible = true
	Option1a.visible = True
	OptionBalls2.visible = True
    OptionReplays2.visible = True
End Sub

Sub HideOptions
	for each objekt In OptionMenu
		objekt.visible = false
	next
End Sub


Sub RollerCoaster_KeyUp(ByVal keycode)

' start roth's ball control
	if keycode = 203 then Cleft = 0' Left Arrow

	if keycode = 200 then Cup = 0' Up Arrow

	if keycode = 208 then Cdown = 0' Down Arrow

	if keycode = 205 then Cright = 0' Right Arrow
' end roth's ball control

	If keycode = PlungerKey Then
		Plunger.Fire
		if ballhome.ballcntover>0 then
			PlaySoundAt "plungerreleaseball", Plunger	'PLAY WHEN BALL IS HIT
		  else
			PlaySoundAt "plungerreleasefree", Plunger		'PLAY WHEN NO BALL TO PLUNGE
		end if
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

' start roth's ball control

Sub StartControl_Hit()
	Set ControlBall = ActiveBall
	contballinplay = true
End Sub

Sub EndControl_Hit()
	contballinplay = false
End Sub

Dim Cup, Cdown, Cleft, Cright, contball, contballinplay, ControlBall, bcboost
Dim bcvel, bcyveloffset, bcboostmulti

bcboost = 1 'Do Not Change - default setting
bcvel = 4 'Controls the speed of the ball movement
bcyveloffset = -0.01 'Offsets the force of gravity to keep the ball from drifting vertically on the table, should be negative
bcboostmulti = 3 'Boost multiplier to ball veloctiy (toggled with the B key)

Sub BallControl_Timer()
	If Contball and ContBallInPlay then
		If Cright = 1 Then
			ControlBall.velx = bcvel*bcboost
		  ElseIf Cleft = 1 Then
			ControlBall.velx = - bcvel*bcboost
		  Else
			ControlBall.velx=0
		End If
		If Cup = 1 Then
			ControlBall.vely = -bcvel*bcboost
		  ElseIf Cdown = 1 Then
			ControlBall.vely = bcvel*bcboost
 		  Else
			ControlBall.vely= bcyveloffset
		End If
	End If
End Sub

' end roth's ball control

sub flippertimer_timer()
	LFlip.RotY = LeftFlipper.CurrentAngle
	RFlip.RotY = RightFlipper.CurrentAngle
	FlipperLSh.RotZ = LeftFlipper.currentangle
	FlipperRSh.RotZ = RightFlipper.currentangle
	ArrowSh.RotZ = 	ArrowPrimitive.ObjRotZ
end sub


Sub PairedlampTimer_timer
	bumperlighta1.state = bumperlight1.state
	bumperlighta2.state = bumperlight2.state
	bumperlighta3.state = bumperlight3.state
	Lrollover1.state=Ltarget1.state
	Lrollover2.state=Ltarget2.state
	Lrollover3.state=Ltarget3.state
	Lrollover4.state=Ltarget4.state
end sub

sub coindelay_timer
	addcredit
    coindelay.enabled=false
end sub

Sub addcredit
      credit=credit+1
	  DOF 117, DOFOn
      if credit>15 then credit=15
	  credittxt.setvalue(credit)
	  If B2SOn Then Controller.B2ssetCredits Credit
End sub

Sub Drain_Hit()
	PlaySound "drain",0,1,0,0.25
	DOF 118, DOFPulse
	me.timerenabled=1
End Sub

Sub Drain_timer
	scorebonus.enabled=true
	me.timerenabled=0
End Sub	


sub ballhome_unhit
	DOF 119, DOFPulse
end sub


sub scorebonus_timer
	  if players=1 or player=players then 
		player=1
		pups(players).state=0
		pup1.state=1
	   Else
		player=player+1
		pups(player).state=1
		pups(player-1).state=0
	  end if
	  If B2SOn then Controller.B2ssetplayerup 30, player
	  nextball
	 scorebonus.enabled=false
End sub


sub newgame_timer
	bumper1.force=8
	bumper2.force=6
	bumper3.force=8
	player=1
	for i = 1 to maxplayers
	    score(i)=0
		rep(i)=0
		pups(i).state=0
		sreels(i).resettozero
	next
    pups(1).state=1
	If B2SOn then
	  for i = 1 to maxplayers
		Controller.B2SSetScorePlayer i, score(i)
	  next 
	End If
    eg=0
	For each light in GIlights:light.state=1:next
	for each light in lanelights:light.state=1:next
	bip_game.state=0
    tilttxt.text=" "
	gamov.text=" "
    If B2SOn then
		Controller.B2SSetGameOver 35,0
		Controller.B2SSetTilt 33,0
		Controller.B2SSetMatch 34,0
	End If
	bip1.state=1
	matchtxt.text=" "
	newball
	ballrelTimer.enabled=true
	newgame.enabled=false
end sub


sub newball
	for each light in bumperlights:light.state=0:next
	for each light in bonuslights: light.state=1:next
	ltargetl.state=0: LtargetR.state=0
	spinnerlight.state=0
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
		for i = 1 to Balls: EVAL("bip"&i).state=0: Next
		EVAL("bip"&ballinplay).state=1
		If B2SOn then Controller.B2ssetballinplay 32, Ballinplay
	end if
End Sub

sub ballreltimer_timer
  dim hiscstate
  if eg=1 then
	  hiscstate=0
	  turnoff
      matchnum
	  state=false
	  for i = 1 to Balls: EVAL("bip"&i).state=0: Next
	  bip_game.state=1
	  gamov.text="GAME OVER"
	  gamov.timerenabled=1
	  tilttxt.timerenabled=1
	  canplayreel.setvalue(0)
	  for i=1 to maxplayers
		if score(i)>hisc then 
			hisc=score(i)
			hiscstate=1
		end if
		pups(i).state=0
	  next
	  if hiscstate=1 then 
			HighScoreEntryInit()
			HStimer.uservalue = 0
			HStimer.enabled=1
	  end if
	  UpdatePostIt 

	  savehs
	  If B2SOn then 
        Controller.B2SSetGameOver 35,1
        Controller.B2ssetballinplay 32, 0

	    Controller.B2ssetcanplay 31, 0
	    Controller.B2ssetcanplay 30, 0
	  End If
	  For each light in LaneLights:light.State = 0: Next
	  For each light in GIlights:light.state=0:next
	  ballreltimer.enabled=false
  else
	Drain.kick 60,15,0
    ballreltimer.enabled=false
	playsound SoundFXDOF("kickerkick",115,DOFPulse,DOFContactors)
  end if
end sub

sub matchnum
	if matchnumb=0 then
		matchtxt.text="0"
		If B2SOn then Controller.B2SSetMatch 10
	  else
		matchtxt.text=matchnumb
		If B2SOn then Controller.B2SSetMatch 34,Matchnumb
	end if
	For i=1 to players
		if (matchnumb)=(score(i) mod 10) then 
		  addcredit
		  playsound SoundFXDOF("knock",116,DOFPulse,DOFKnocker)
		  DOF 125, DOFPulse
	    end if
    next
end sub

'********** Bumpers

Sub Bumper1_Hit
   if tilt=false then
	PlaySound SoundFXDOF("fx_bumper4",105,DOFPulse,DOFContactors)
	DOF 122, DOFPulse
	if bumperlight1.state=1 then
		addscore 10
	  else
		addscore bumperoffscore
	end if
   end if
End Sub



Sub Bumper2_Hit
   if tilt=false then
	PlaySound SoundFXDOF("fx_bumper4",106,DOFPulse,DOFContactors)
	DOF 123, DOFPulse
	if BumperLight2.state=1 then
		addscore 100
	  else
		addscore bumperoffscore
	end if	
   end if
End Sub


Sub Bumper3_Hit
   if tilt=false then
	PlaySound SoundFXDOF("fx_bumper4",107,DOFPulse,DOFContactors)
	DOF 124, DOFPulse
	if BumperLight3.state=1 then
		addscore 10
	  else
		addscore bumperoffscore
	end if	
   end if
End Sub



'************** Slings

Sub RightSlingShot_Slingshot
    PlaySound SoundFXDOF("right_slingshot",104,DOFPulse,DOFcontactors), 0, 1, 0.05, 0.05
	DOF 121, DOFPulse
	addscore 1
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
    PlaySound SoundFXDOF("left_slingshot",103,DOFPulse,DOFcontactors),0,1,-0.05,0.05
	DOF 120, DOFPulse
	addscore 1
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

Sub Dingwalls_hit (idx)
	addscore 1
End Sub

'********** Pointer workings

Sub Spinner_Hit(Idx)
	Dim TempPos
	If Idx>39 Then
		TempVelX=ActiveBall.VelX
		TempVelY=ActiveBall.VelY
		Exit Sub
	End If

	If Idx>19 then
		TempPos=Idx-20
	else
		TempPos=Idx
	end if

	BallSpeed=Int(Abs(TempVelX*5))+(Abs(TempVelY*5))
	If BallSpeed<14 Then Exit Sub
	If Idx<20 Then
		DOF 114, DOFPulse
		SpinPos=SpinPos+1
		If SpinPos=20 Then SpinPos=0
		Me(Idx).IsDropped=True
		Me(Idx+20).IsDropped=True
	Else
		DOF 114, DOFPulse
		SpinPos=SpinPos-1
		If SpinPos=-1 Then SpinPos=19
		Me(Idx-20).IsDropped=True
		Me(Idx).IsDropped=True
	End If
	Me(SpinPos).IsDropped=False
	Me(SpinPos+20).IsDropped=False
	ArrowPrimitive.ObjRotZ=SpinArrow(SpinPos)
	Activeball.velx=TempVelX*.8
	Activeball.vely=TempVely*.8
	Count3=0
	SpinnerLight.TimerEnabled=True
End Sub


Sub SpinnerLight_Timer    '*****Score Spinner
  if tilt=false then
	if SpinnerLight.state=1 then
		if Spin(SpinPos)=50 then
			addscore 500
			for each light in bonuslights: light.state=1: next
			for each light in bumperlights: light.state=0: next
			spinnerlight.state=0
			LtargetL.state=0
			LtargetR.state=0
		else
			addscore Spin(SpinPos)
		end if
	else
		addscore Spin(SpinPos)
	end if
  end if
	SpinnerLight.timerenabled=false
End Sub

'********** Kickers

sub kicker1_hit
	if tilt=false then addscore Spin(SpinPos)
	me.uservalue=1
	me.timerenabled=1
End Sub

sub kicker1_timer
	select case me.uservalue
	  case 5:
		playsound SoundFXDOF("HoleKickL",112,DOFPulse,DOFContactors)
		playsound "WireRampShortL"
		DOF 125, DOFPulse
		Kicker1.kick 180,17,20
		Pkickarm.rotz = 20
		Pkickarm1.rotz = 20
		Lkicker1.state=1
	  case 8:
		Pkickarm.rotz = 0
		Pkickarm1.rotz = 0
	  case 10:
		Lkicker1.state=0
		me.timerenabled=0
	end Select
	me.uservalue=me.uservalue+1
end Sub


sub kicker2_hit
	if tilt=false then addscore Spin(SpinPos)
	me.uservalue=1
	me.timerenabled=1
End Sub  

sub kicker2_timer
	select case me.uservalue
	  case 5:
		playsound SoundFXDOF("HoleKickR",113,DOFPulse,DOFContactors)
		playsound "WireRampShortR"
		DOF 125, DOFPulse
		Kicker2.kick 180,17,20
		Pkickarm2.rotz = 20
		Pkickarm3.rotz = 20
		Lkicker2.state=1
	  case 8:
		Pkickarm2.rotz = 0
		Pkickarm3.rotz = 0
	  case 10:
		Lkicker2.state=0
		me.timerenabled=0
	end Select
	me.uservalue=me.uservalue+1
end Sub

'sub kicker2_timer
'	playsound SoundFXDOF("holekick",113,DOFPulse,DOFContactors)
'	DOF 125, DOFPulse
'	Kicker2.kick 180,17,20
'	Lkicker2.state=1
'	Lkicker2.timerenabled=1
'	me.timerenabled=0
'end Sub   
'
'sub Lkicker2_timer
'	Lkicker2.state=0
'	Lkicker2.timerenabled=0
'end Sub

'********** Trigger Rollover

sub TGoutL_hit
	DOF 205, DOFPulse
	if tilt=false Then addscore Spin(SpinPos)
end sub    

sub TGoutR_hit
	DOF 206, DOFPulse
	if tilt=false Then addscore Spin(SpinPos)
end sub

sub TGrollover1_hit
	DOF 203, DOFPulse
	if tilt=false Then
		addscore 50
		Ltarget1.state=0
		LtargetR.state=1
		if balls=3 Then	
			ltarget4.state=0
			BumperLight1.state=1
			BumperLight3.state=1
		end If
		checkspecial
	end If
end sub

sub TGrollover2_hit
	DOF 200, DOFPulse
	if tilt=false Then
		addscore 50
		Ltarget2.state=0
		LtargetL.state=1
		if balls=3 Then
			Ltarget3.state=0
			BumperLight2.state=1
		end if
		checkspecial
	end If
end sub

sub TGrollover3_hit
	DOF 202, DOFPulse
	if tilt=false Then
		addscore 50
		Ltarget3.state=0
		BumperLight2.state=1
		if balls=3 Then
			Ltarget2.state=0
			LtargetL.state=1
		end if
		checkspecial
	end If
end sub

sub TGrollover4_hit
	DOF 204, DOFPulse
	if tilt=false Then
		addscore 50
		Ltarget4.state=0
		BumperLight1.state=1
		bumperlight3.state=1
		if balls=3 Then	
			ltarget1.state=0
			ltargetr.state=1
		end If
		checkspecial
	end If
end sub
 
sub TGpointer_hit  
	DOF 201, DOFPulse
	if tilt=false then addscore Spin(SpinPos)
end sub



'********** Targets

sub Target1_hit
	DOF 108, DOFPulse
	target1P.transx=-5
	target1.timerenabled=1
	if tilt=false then
		addscore 50
		Ltarget1.state=0
		LtargetR.state=1
		if balls=3 Then
			Ltarget4.state=0
			BumperLight1.state=1
			BumperLight3.state=1
		end if
		checkspecial
     end if
end Sub

Sub target1_timer
	target1P.transx=0
End Sub

sub Target2_hit
	DOF 108, DOFPulse
	target2P.transx=-5
	Target2.timerenabled=1
	if tilt=false then
		addscore 50
		Ltarget2.state=0
		LtargetL.state=1
		if balls=3 then 
			Ltarget3.state=0
			BumperLight2.state=1
		end if
		checkspecial
     end if
end Sub

Sub target2_timer
	target2P.transx=0
End Sub

sub Target3_hit
	DOF 109, DOFPulse
	target3P.transx=-5
	Target3.timerenabled=1
	if tilt=false then
		addscore 50
		Ltarget3.state=0
		if balls=3 then 
			LtargetL.state=1
			Ltarget2.state=0
		end if
		checkspecial
     end if
end Sub

Sub target3_timer
	target3P.transx=0
End Sub

sub Target4_hit
	DOF 109, DOFPulse
	target4P.transx=-5
	Target4.timerenabled=1
	if tilt=false then
		addscore 50
		Ltarget4.state=0
		if balls=3 then 
			LtargetR.state=1
			Ltarget1.state=0
		end if
		checkspecial
     end if
end Sub

Sub target4_timer
	target4P.transx=0
End Sub

sub TargetL_hit
	DOF 110, DOFPulse
	if tilt=false Then
		if LtargetL.state=1 Then
			addscore 100
		  Else
			addscore 10
		end If
	end if
end Sub

sub TargetR_hit
	DOF 111, DOFPulse
	if tilt=false Then
		if LtargetR.state=1 Then
			addscore 100
		  Else
			addscore 10
		end If
	end if
end Sub

sub checkspecial
	if Ltarget1.state+Ltarget2.state+Ltarget3.state+Ltarget4.state=0 then spinnerlight.state=1
end Sub

'********************Scoring

sub addscore(points)
  if tilt=false then
	If points = 1 then 
		matchnumb=matchnumb+1
		if matchnumb>9 then matchnumb=0
		addpoints 1
	end if
	If points=10 or points=100 Then
		AddPoints Points
	  else
		If Points < 100 and AddScore10Timer.enabled = false Then
			Add10 = Points \ 10
			AddScore10Timer.Enabled = TRUE
		  ElseIf Points < 1000 and AddScore100Timer.enabled = false Then
			Add100 = Points \ 100
			AddScore100Timer.Enabled = TRUE
		End If
	 end if
  end if
End Sub

Sub AddScore10Timer_Timer()
    if Add10 > 0 then
        AddPoints 10
        Add10 = Add10 - 1
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

Sub AddPoints(Points)
    score(player)=score(player)+points
	sreels(player).addvalue(points)
	If B2SOn Then Controller.B2SSetScorePlayer player, score(player)

    ' Sounds: there are 3 sounds: ones, tens, and hundreds
    If Points = 10 AND(Score(player) MOD 100) \ 10 = 0 Then  'New 100 reel
        PlaySound SoundFXDOF("bell100",143,DOFPulse,DOFChimes)
      ElseIf Points = 1 AND(Score(player) MOD 10) = 0 Then 'New 10 reel
        PlaySound SoundFXDOF("bell10",142,DOFPulse,DOFChimes)
      ElseIf points = 100 Then
        PlaySound SoundFXDOF("bell100",143,DOFPulse,DOFChimes)
	  elseif Points = 10 Then
        PlaySound SoundFXDOF("bell10",142,DOFPulse,DOFChimes)
      Else
        PlaySound SoundFXDOF("bell1",141,DOFPulse,DOFChimes)
    End If
    ' check replays and rollover
	if score(player)=>99999 then
		If B2SOn then Controller.B2SSetScoreRollover 24 + player, 1
		P100k(player).text="100,000"
	End if
    if score(player)=>replay1 and rep(player)=0 then
		addcredit
		rep(player)=1
		PlaySound SoundFXDOF("knock",116,DOFPulse,DOFKnocker)
		DOF 125, DOFPulse
    end if
    if score(player)=>replay2 and rep(player)=1 then
		addcredit
		rep(player)=2
		PlaySound SoundFXDOF("knock",116,DOFPulse,DOFKnocker)
		DOF 125, DOFPulse
    end if
    if score(player)=>replay3 and rep(player)=2 then
		addcredit
		rep(player)=3
		PlaySound SoundFXDOF("knock",116,DOFPulse,DOFKnocker)
		DOF 125, DOFPulse
    end if
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

Sub Tilttimer_Timer()
	Tilttimer.Enabled = False
End Sub

sub turnoff
    bumper1.force=0
    bumper2.force=0
	bumper3.force=0
  	LeftFlipper.RotateToStart
	StopSound "Buzz"
	RightFlipper.RotateToStart
	StopSound "Buzz1"
end sub    





'*****************************************
'			BALL SHADOW
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
        If BOT(b).X < RollerCoaster.Width/2 Then
            BallShadow(b).X = ((BOT(b).X) - (Ballsize/6) + ((BOT(b).X - (RollerCoaster.Width/2))/7)) '+ 13
        Else
            BallShadow(b).X = ((BOT(b).X) + (Ballsize/6) + ((BOT(b).X - (RollerCoaster.Width/2))/7)) '- 13
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

'**************************************************************************
'                 Additional Positional Sound Playback Functions by DJRobX
'**************************************************************************

Sub PlaySoundAtVol(sound, tableobj, Vol)
		PlaySound sound, 1, Vol, AudioPan(tableobj), 0,0,0, 1, AudioFade(tableobj)
End Sub

'Set position at table object, vol, and loops manually.

Sub PlaySoundAtVolLoops(sound, tableobj, Vol, Loops)
		PlaySound sound, Loops, Vol, AudioPan(tableobj), 0,0,0, 1, AudioFade(tableobj)
End Sub


'*********************************************************************
'                     Supporting Ball & Sound Functions
'*********************************************************************

Function AudioFade(tableobj) ' Fades between front and back of the table (for surround systems or 2x2 speakers, etc), depending on the Y position on the table. "RollerCoaster" is the name of the table
	Dim tmp
    tmp = tableobj.y * 2 / RollerCoaster.height-1
    If tmp > 0 Then
		AudioFade = Csng(tmp ^10)
    Else
        AudioFade = Csng(-((- tmp) ^10) )
    End If
End Function

Function AudioPan(tableobj) ' Calculates the pan for a tableobj based on the X position on the table. "RollerCoaster" is the name of the table
    Dim tmp
    tmp = tableobj.x * 2 / RollerCoaster.width-1
    If tmp > 0 Then
        AudioPan = Csng(tmp ^10)
    Else
        AudioPan = Csng(-((- tmp) ^10) )
    End If
End Function

Function Vol(ball) ' Calculates the Volume of the sound based on the ball speed
    Vol = Csng(BallVel(ball) ^2 / 4000)
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
	PlaySound("fx_collide"), 0, Csng(velocity) ^2 / 1000, AudioPan(ball1), 0, Pitch(ball1), 0, 0, AudioFade(ball1)
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


Sub a_Triggers_Hit (idx)
	playsound "Switch", 0,1,AudioPan(ActiveBall),0,0,0,1,AudioFade(ActiveBall)
End sub

Sub a_Pins_Hit (idx)
	PlaySound "pinhit_low", 0, Vol(ActiveBall), AudioPan(ActiveBall), 0, Pitch(ActiveBall), 0, 0, AudioFade(ActiveBall)
End Sub

Sub a_Targets_Hit (idx)
	PlaySound "target", 0, Vol(ActiveBall), AudioPan(ActiveBall), 0, Pitch(ActiveBall), 0, 0, AudioFade(ActiveBall)
End Sub

Sub a_DropTargets_Hit (idx)
	PlaySound "DTDrop", 0, Vol(ActiveBall), AudioPan(ActiveBall), 0, Pitch(ActiveBall), 0, 0, AudioFade(ActiveBall)
End Sub

Sub a_Metals_Thin_Hit (idx)
	PlaySound "metalhit_thin", 0, Vol(ActiveBall), AudioPan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
End Sub

Sub a_Metals_Medium_Hit (idx)
	PlaySound "metalhit_medium", 1, Vol(ActiveBall), AudioPan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
End Sub

Sub a_Metals2_Hit (idx)
	PlaySound "metalhit2", 0, Vol(ActiveBall), AudioPan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
End Sub

Sub a_Gates_Hit (idx)
	PlaySound "gate4", 0, Vol(ActiveBall), AudioPan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
End Sub

Sub Spinner_Spin
	PlaySound "fx_spinner", 1, .25, AudioPan(Spinner), 0.25, 0, 0, 1, AudioFade(Spinner)
End Sub

Sub a_Rubbers_Hit(idx)
 	dim finalspeed
  	finalspeed=SQR(activeball.velx * activeball.velx + activeball.vely * activeball.vely)
 	If finalspeed > 16 then 
		PlaySound "fx_rubber2", 1, Vol(ActiveBall), AudioPan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
	End if
	If finalspeed >= 6 AND finalspeed <= 16 then
 		RandomSoundRubber()
 	End If
End Sub

Sub a_Posts_Hit(idx)
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
		Case 1 : PlaySound "rubber_hit_1", 1, Vol(ActiveBall), AudioPan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
		Case 2 : PlaySound "rubber_hit_2", 1, Vol(ActiveBall), AudioPan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
		Case 3 : PlaySound "rubber_hit_3", 1, Vol(ActiveBall), AudioPan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
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

sub savehs

    savevalue "RollerCoaster", "credit", credit
    savevalue "RollerCoaster", "hiscore", hisc
    savevalue "RollerCoaster", "match", matchnumb
    savevalue "RollerCoaster", "score1", score(1)
    savevalue "RollerCoaster", "score2", score(2)
	savevalue "RollerCoaster", "balls", balls
	savevalue "RollerCoaster", "replays", replays
	savevalue "RollerCoaster", "spinpos", spinpos
	savevalue "RollerCoaster", "hsa1", HSA1
	savevalue "RollerCoaster", "hsa2", HSA2
	savevalue "RollerCoaster", "hsa3", HSA3

end sub

sub loadhs
    dim temp
	temp = LoadValue("RollerCoaster", "credit")
    If (temp <> "") then credit = CDbl(temp)
    temp = LoadValue("RollerCoaster", "hiscore")
    If (temp <> "") then hisc = CDbl(temp)
    temp = LoadValue("RollerCoaster", "match")
    If (temp <> "") then matchnumb = CDbl(temp)
    temp = LoadValue("RollerCoaster", "score1")
    If (temp <> "") then score(1) = CDbl(temp)
    temp = LoadValue("RollerCoaster", "score2")
    If (temp <> "") then score(2) = CDbl(temp)
    temp = LoadValue("RollerCoaster", "balls")
    If (temp <> "") then balls = CDbl(temp)
    temp = LoadValue("RollerCoaster", "replays")
    If (temp <> "") then replays = CDbl(temp)
    temp = LoadValue("RollerCoaster", "spinpos")
    If (temp <> "") then spinpos = CDbl(temp)
    temp = LoadValue("RollerCoaster", "hsa1")
    If (temp <> "") then HSA1 = CDbl(temp)
    temp = LoadValue("RollerCoaster", "hsa2")
    If (temp <> "") then HSA2 = CDbl(temp)
    temp = LoadValue("RollerCoaster", "hsa3")
    If (temp <> "") then HSA3 = CDbl(temp)
end sub

Sub RollerCoaster_Exit()
	Savehs
	turnoff
	If B2SOn Then Controller.stop
End Sub

'==========================================================================================================================================
'============================================================= START OF HIGH SCORES ROUTINES =============================================================
'==========================================================================================================================================
'
'ADD LINE TO TABLE_KEYDOWN SUB WITH THE FOLLOWING:    If HSEnterMode Then HighScoreProcessKey(keycode) AFTER THE STARTGAME ENTRY
'ADD: And Not HSEnterMode=true TO IF KEYCODE=STARTGAMEKEY
'TO SHOW THE SCORE ON POST-IT ADD LINE AT RELEVENT LOCATION THAT HAS:  UpdatePostIt
'TO INITIATE ADDING INITIALS ADD LINE AT RELEVENT LOCATION THAT HAS:  HighScoreEntryInit()
'ADD THE FOLLOWING LINES TO TABLE_INIT TO SETUP POSTIT
'	if HSA1="" then HSA1=25
'	if HSA2="" then HSA2=25
'	if HSA3="" then HSA3=25
'	UpdatePostIt
'ADD HSA1, HSA2 AND HSA3 TO SAVE AND LOAD VALUES FOR TABLE
'ADD A TIMER NAMED HighScoreFlashTimer WITH INTERVAL 100 TO TABLE
'SET HSSSCOREX BELOW TO WHATEVER VARIABLE YOU USE FOR HIGH SCORE.
'ADD OBJECTS TO PLAYFIELD (EASIEST TO JUST COPY FROM THIS TABLE)
'IMPORT POST-IT IMAGES


Dim HSA1, HSA2, HSA3
Dim HSEnterMode, hsLetterFlash, hsEnteredDigits(3), hsCurrentDigit, hsCurrentLetter
Dim HSArray  
Dim HSScoreM,HSScore100k, HSScore10k, HSScoreK, HSScore100, HSScore10, HSScore1, HSScorex	'Define 6 different score values for each reel to use
HSArray = Array("Postit0","postit1","postit2","postit3","postit4","postit5","postit6","postit7","postit8","postit9","postitBL","postitCM","Tape")
Const hsFlashDelay = 4

' ***********************************************************
'  HiScore DISPLAY 
' ***********************************************************

Sub UpdatePostIt
	dim tempscore
	HSScorex = hisc
	TempScore = HSScorex
	HSScore1 = 0
	HSScore10 = 0
	HSScore100 = 0
	HSScoreK = 0
	HSScore10k = 0
	HSScore100k = 0
	HSScoreM = 0
	if len(TempScore) > 0 Then
		HSScore1 = cint(right(Tempscore,1))
	end If
	if len(TempScore) > 1 Then
		TempScore = Left(TempScore,len(TempScore)-1)
		HSScore10 = cint(right(Tempscore,1))
	end If
	if len(TempScore) > 1 Then
		TempScore = Left(TempScore,len(TempScore)-1)
		HSScore100 = cint(right(Tempscore,1))
	end If
	if len(TempScore) > 1 Then
		TempScore = Left(TempScore,len(TempScore)-1)
		HSScoreK = cint(right(Tempscore,1))
	end If
	if len(TempScore) > 1 Then
		TempScore = Left(TempScore,len(TempScore)-1)
		HSScore10k = cint(right(Tempscore,1))
	end If
	if len(TempScore) > 1 Then
		TempScore = Left(TempScore,len(TempScore)-1)
		HSScore100k = cint(right(Tempscore,1))
	end If
	if len(TempScore) > 1 Then
		TempScore = Left(TempScore,len(TempScore)-1)
		HSScoreM = cint(right(Tempscore,1))
	end If
	Pscore6.image = HSArray(HSScoreM):If HSScorex<1000000 Then PScore6.image = HSArray(10)
	Pscore5.image = HSArray(HSScore100K):If HSScorex<100000 Then PScore5.image = HSArray(10)
	PScore4.image = HSArray(HSScore10K):If HSScorex<10000 Then PScore4.image = HSArray(10)
	PScore3.image = HSArray(HSScoreK):If HSScorex<1000 Then PScore3.image = HSArray(10)
	PScore2.image = HSArray(HSScore100):If HSScorex<100 Then PScore2.image = HSArray(10)
	PScore1.image = HSArray(HSScore10):If HSScorex<10 Then PScore1.image = HSArray(10)
	PScore0.image = HSArray(HSScore1):If HSScorex<1 Then PScore0.image = HSArray(10)
	if HSScorex<1000 then
		PComma.image = HSArray(10)
	else
		PComma.image = HSArray(11)
	end if
	if HSScorex<1000000 then
		PComma2.image = HSArray(10)
	else
		PComma2.image = HSArray(11)
	end if
'	if showhisc=1 and showhiscnames=1 then
'		for each objekt in hiscname:objekt.visible=1:next
		HSName1.image = ImgFromCode(HSA1, 1)
		HSName2.image = ImgFromCode(HSA2, 2)
		HSName3.image = ImgFromCode(HSA3, 3)
'	  else
'		for each objekt in hiscname:objekt.visible=0:next
'	end if
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
'  HiScore ENTER INITIALS 
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
'				EMReelHSName1.SetValue HSA1:EMReelHSName2.SetValue HSA2
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
					savehs 'enter it
					HighScoreFlashTimer.Enabled = False
					HSEnterMode = False
				End If
		End Select
		UpdatePostIt
    End If
End Sub


' Thalamus : Exit in a clean and proper way
Sub RollerCoaster_exit()
  Controller.Pause = False
  Controller.Stop
End Sub


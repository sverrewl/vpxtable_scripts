'	-------------------------------------------------
'	Royal Flush Gottlieb 1976
'	-------------------------------------------------
'	
'	by BorgDog, 2017
'
'	Layer usage (my usual, but not always mathes this)
'		1 - most stuff
'		2 - triggers
'		3 - stuff with shadows
'		4 - options menu
'		5 - hi score sticky
'		6 - GI lighting
'		7 - plastics
'		8 - insert and bumper lighting
'
'	DOF config
'		101 Left Flipper, 102 Right Flipper, 
'		103 Left sling, 104 Left sling flasher, 105 right sling, 106 right sling flasher,
'		107 Bumper1, 108 Bumper1 flasher, 
'		111 Drop Targets left 4, 112 Drop Targets left reset, 113 Drop Targets right 5, 114 Drop Targets right Reset
'		115 top target, 116 upper right target, 117 lower left target
'		118 right kicker
'		119 top left rollover, 120 top green rollover, 121 top white rollover, 122 top purple rollover, 123 top right rollover
'		124 left outlane rollover, 125 left inlane rollover, 126 right outlane rollover, 127 right inlane rollover
'		128 left 500 target, 129 mid 3000 target, 130 right 500 target
'		134 Drain, 135 Ball Release, 136 Shooter Lane/launch ball, 137 credit light, 
'		138 knocker, 139 knocker/kicker Flasher
'		141 Chime1-10s, 142 Chime2-100s, 143 Chime3-1000s
'
'	-------------------------------------------------

Option Explicit
Randomize

Const cGameName = "cardwhiz_1976"

Dim operatormenu, options
Dim balls
Dim replays, freeplay
Dim Replay1Table(3), Replay2Table(3), Replay3Table(3)
dim replay1, replay2, replay3
Dim Add10, Add100, Add1000
Dim hisc
Dim maxplayers, players, player, TempPlayerUp, rotatortemp
Dim credit, bonuscount
Dim score(4)
Dim state, holevalue, specialstate
Dim tilt, tiltsens
Dim ballinplay, PlungeBall
Dim matchnumb
dim rstep, lstep
Dim rep(4), rst, eg
Dim bell
Dim i,j, ii, objekt, light
Dim awardcheck

On Error Resume Next
ExecuteGlobal GetTextFile("controller.vbs")
If Err Then MsgBox "Can't open controller.vbs"
On Error Goto 0

sub RoyalFlush_init
	LoadEM
	maxplayers=4
	Replay1Table(1)=62000
	Replay2Table(1)=76000
	Replay3Table(1)=84000
	Replay1Table(2)=67000
	Replay2Table(2)=81000
	Replay3Table(2)=89000
	Replay1Table(3)=75000
	Replay2Table(3)=89000
	Replay3Table(3)=97000
	hideoptions
	player=1
	rotatortemp=1
	balls=5
	replays=3
	freeplay=0
	matchnumb=0
	hisc=20000
	state=False
	loadhs
	if HSA1="" then HSA1=4
	if HSA2="" then HSA2=15
	if HSA3="" then HSA3=7
	UpdatePostIt
	for i = 1 to maxplayers
		EVAL("ScoreReel"&i).setvalue score(i)
		EVAL("Reel100K"&i).setvalue(int(score(i)/100000))
	next
	bipreel.setvalue 0
	credittxt.setvalue(credit)
	Replay1=Replay1Table(Replays)
	Replay2=Replay2Table(Replays)
	Replay3=Replay3Table(Replays)
	RepCard.image = "ReplayCard"&replays
	OptionBalls.image="OptionsBalls"&Balls
	OptionReplays.image="OptionsReplays"&replays
	OptionFreeplay.image="OptionsFreeplay"&freeplay
	if balls=3 then	
		InstCard.image="InstCard3balls"
	  else
		InstCard.image="InstCard5balls"
	end if
	If B2SOn then
		for each objekt in backdropstuff : objekt.visible = 0 : next
	End If
	startGame.enabled=true

	tilt=false
	turnoff
	If credit>0 or freeplay=1 then DOF 137, DOFOn
    Drain.CreateBall
End sub

sub startGame_timer
	playsound "poweron"
	lightdelay.enabled=true
	me.enabled=false
end sub

sub lightdelay_timer
	if matchnumb=0 then
		matchtxt.text="00"
	  else
		matchtxt.text=matchnumb*10
	end if
	gamov.text="Game Over"
	gamov.timerenabled=1
	tilttxt.timerenabled=1
	For each light in GIlights:light.state=1:Next
	For each light in BumperLights:light.state=1:Next
	if B2SOn then
		Controller.B2SSetData 90,1
		for i = 1 to maxplayers
			Controller.B2SSetScorePlayer i, Score(i) MOD 100000
		next
		Controller.B2ssetCredits Credit
		Controller.B2ssetMatch 34, Matchnumb
		Controller.B2SSetGameOver 35,1
		Controller.B2SSetScorePlayer 5, hisc
	end if
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

Sub RoyalFlush_KeyDown(ByVal keycode)
   
	if keycode=AddCreditKey then
		playsound "coinin" 
		coindelay.enabled=true 
    end if

    if keycode=StartGameKey and (Credit>0 or freeplay=1) And Not HSEnterMode=true then
	  if state=false then
		player=1
		if freeplay=0 then 
			credit=credit-1
			if credit < 1 then DOF 137, DOFOff
			credittxt.setvalue(credit)
		End if
		playsound "cluper"
		ballinplay=1
		for i = 1 to maxplayers
			score(i)=0
			rep(i)=0
			EVAL("Pup"&i).state=0
			EVAL("Pups"&i).state=0
			EVAL("ScoreReel"&i).resettozero
		next
		If B2SOn Then 
			Controller.B2ssetCredits Credit
			Controller.B2ssetballinplay 32, Ballinplay
			Controller.B2ssetplayerup 30, 1
			Controller.B2ssetcanplay 31, 1
			Controller.B2SSetGameOver 0
			for i = 1 to maxplayers
				Controller.B2SSetScorePlayer i, score(i)
			next 
		End If
	    Pup1.state=1
		pups1.state=1
		tilt=false
		state=true
		playsound "initialize"
		PlaySound("RotateThruPlayers"),0,.05,0,0.25
 		TempPlayerUp=Player
		PlayerUpRotator.enabled=true
		players=1
		for i = 1 to maxplayers
			EVAL("CanPlay"&i).state=0
		next
		EVAL("CanPlay"&players).state=1
		rst=0
		newgame.enabled=true
	  else if state=true and players < maxplayers and Ballinplay=1 then
		if freeplay=0 then 
			credit=credit-1
			if credit < 1 then DOF 137, DOFOff
			credittxt.setvalue(credit)
		End if
		players=players+1
'		canplayreel.setvalue(players)
		for i = 1 to maxplayers
			EVAL("CanPlay"&i).state=0
		next
		EVAL("CanPlay"&players).state=1
		If B2SOn then
			if freeplay=0 then Controller.B2ssetCredits Credit
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
		If Options=5 then Options=1
		playsound "target"
		Select Case (Options)
			Case 1:
				Option1.visible=true
				Option4.visible=False
			Case 2:
				Option2.visible=true
				Option1.visible=False
			Case 3:
				Option3.visible=true
				Option2.visible=False
			Case 4:
				Option4.visible=true
				Option3.visible=False
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
			OptionBalls.image = "OptionsBalls"&Balls 
		Case 2:
			if freeplay=0 Then
				freeplay=1
			  Else
				freeplay=0
			end if
			OptionFreeplay.image="OptionsFreeplay"&freeplay    
		Case 3:
			Replays=Replays+1
			if Replays>3 then
				Replays=1
			end if
			Replay1=Replay1Table(Replays)
			Replay2=Replay2Table(Replays)
			OptionReplays.image = "OptionsReplays"&replays
			repcard.image = "ReplayCard"&replays
		Case 4:
			OperatorMenu=0
			savehs
			HideOptions
	  End Select
	End If

    If HSEnterMode Then HighScoreProcessKey(keycode)

  if tilt=false and state=true then
	If keycode = LeftFlipperKey Then
		LeftFlipper.RotateToEnd
		PlaySound SoundFXDOF("flipperup",101,DOFOn,DOFFlippers), 0, .67, -0.05, 0.05
		PlaySound "Buzz",-1,.05,-0.05, 0.05
	End If
    
	If keycode = RightFlipperKey Then
		RightFlipper.RotateToEnd
		PlaySound SoundFXDOF("flipperup",102,DOFOn,DOFFlippers), 0, .67, 0.05, 0.05
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
		gametilted
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
	OptionFreeplay.visible = True
End Sub

Sub HideOptions
	for each objekt In OptionMenu
		objekt.visible = false
	next
End Sub


Sub RoyalFlush_KeyUp(ByVal keycode)

	If keycode = PlungerKey Then
		Plunger.Fire
		if PlungeBall=1 then
			playsound "plunger",0,1,0.1,.25,0.25
		  else
			playsound "plungerreleasefree",0,1,0.1,.25,0.25
		end if
	End If

	if keycode = LeftFlipperKey then
		OperatorMenuTimer.Enabled = false
	end if

   If tilt=false and state=true then
	If keycode = LeftFlipperKey Then
		LeftFlipper.RotateToStart
		PlaySound SoundFXDOF("flipperdown",101,DOFOff,DOFFlippers), 0, 1, -0.05, 0.05
		StopSound "Buzz"
	End If
    
	If keycode = RightFlipperKey Then
		RightFlipper.RotateToStart
		PlaySound SoundFXDOF("flipperdown",102,DOFOff,DOFFlippers), 0, 1, 0.05, 0.05
		StopSound "Buzz1"
	End If
   End if
End Sub

sub flippertimer_timer()
	Lbulb3000whenLit.state=l3000whenLit.state
	LFlip.RotY = LeftFlipper.CurrentAngle
	RFlip.RotY = RightFlipper.CurrentAngle
	Pgate.rotz = (Gate.currentangle*0.75)+25
	diverter.RotY = DiverterFlipper.CurrentAngle+90
testbox.text=PlayerUpRotator.enabled
end sub


Sub PairedlampTimer_timer
	bumperlighta1.state = bumperlight1.state
	bumperlighta2.state = bumperlight2.state
	bumperlighta3.state = bumperlight3.state
end sub

sub coindelay_timer
	addcredit
    coindelay.enabled=false
end sub

Sub addcredit
	if freeplay=0 then 
      credit=credit+1
	  DOF 137, DOFOn
      if credit>25 then credit=25
	  credittxt.setvalue(credit)
	  If B2SOn Then Controller.B2ssetCredits Credit
	end if
End sub

Sub addspecial
	playsound SoundFXDOF("knock",138,DOFPulse,DOFKnocker)
	addcredit
end sub

Sub Drain_Hit()
	DOF 134, DOFPulse
	PlaySound "drain",0,1,0,0.25
	me.timerenabled=1
End Sub

Sub Drain_timer
	scorebonus.enabled=true
	me.timerenabled=0
End Sub	

sub ballhome_hit
	plungeball=1
end sub

sub ballhome_unhit
	DOF 136, DOFPulse
	plungeball=0
end sub


sub scorebonus_timer			'************add bonus scoring if applicable
	if tilt=true then
		nextball
	  else
		if Lbonus1.state+lbonus2.state+lbonus3.state+lbonus4.state+lbonus5.state > 0 then
			bonuscount=-1
			for each light in FlashLights:light.state=0:next
			Lbonus1.timerenabled=true
		  else
			nextball
		end if
	end if
	scorebonus.enabled=false
End sub

sub Lbonus1_timer
	bonuscount=bonuscount+1
	if bonuscount>DoubleBonus.state then
		bonuscount=-1
		Lbonus2.timerenabled=true
		for each light in FlashLights:light.state=1:next
		Lbonus1.timerenabled=false
		Lbonus1.timerinterval=325
	 else
		if AddScore1000Timer.enabled = false then
			playsound "scorebonus"
			Lbonus1A.state=1
			if Lbonus1.state=1 then addscore 1000
		else
			bonuscount=bonuscount-1
			Lbonus1.timerinterval=150		
		end if
	end if
end sub

sub Lbonus2_timer
	bonuscount=bonuscount+1
	if bonuscount>DoubleBonus.state then
		bonuscount=-1
		Lbonus3.timerenabled=true
		for each light in FlashLights:light.state=0:next
		Lbonus2.timerenabled=false
		Lbonus2.timerinterval=325	
	 else
		  if AddScore1000Timer.enabled = false then
			  playsound "scorebonus"
			  Lbonus1A.state=0
			  Lbonus2A.state=1
			  if Lbonus2.state=1 then addscore 2000
		  else
			bonuscount=bonuscount-1
			Lbonus2.timerinterval=150	
		  end if
	end if
end sub

sub Lbonus3_timer
	bonuscount=bonuscount+1
	if bonuscount>DoubleBonus.state then
		bonuscount=-1
		Lbonus4.timerenabled=true
		for each light in FlashLights:light.state=1:next
		Lbonus3.timerenabled=false
		Lbonus3.timerinterval=325	
	 else
		  if AddScore1000Timer.enabled = false then
			  playsound "scorebonus"
			  Lbonus2A.state=0
			  Lbonus3A.state=1
			  if Lbonus3.state=1 then addscore 3000
		  else
			bonuscount=bonuscount-1
			Lbonus3.timerinterval=150	
		  end if
	end if
end sub

sub Lbonus4_timer
	bonuscount=bonuscount+1
	if bonuscount>DoubleBonus.state then
		bonuscount=-1
		Lbonus5.timerenabled=true
		for each light in FlashLights:light.state=0:next
		Lbonus4.timerenabled=false
		Lbonus4.timerinterval=325	
	 else
		  if AddScore1000Timer.enabled = false then
			  playsound "scorebonus"
			  Lbonus3A.state=0
			  Lbonus4A.state=1
			  if Lbonus4.state=1 then addscore 4000
		  else
			bonuscount=bonuscount-1
			Lbonus4.timerinterval=150	
		  end if
	end if
end sub

sub Lbonus5_timer
	bonuscount=bonuscount+1
	if bonuscount>DoubleBonus.state then
		DoubleBonus.timerenabled=true
		for each light in FlashLights:light.state=1:next
		Lbonus5.timerenabled=false
		Lbonus5.timerinterval=325	
	 else
		  if AddScore1000Timer.enabled = false then
			  playsound "scorebonus"
			  Lbonus4A.state=0
			  Lbonus5A.state=1
			  if Lbonus5.state=1 then addscore 5000
		  else
			bonuscount=bonuscount-1
			Lbonus5.timerinterval=150	
		  end if
	end if
end sub

sub DoubleBonus_timer
	Lbonus5A.state=0
	nextball
	doublebonus.timerenabled=False
end sub

sub newgame_timer
	player=1
	tilt=False
    eg=0
	For each light in GIlights:light.state=1:next
    tilttxt.text=" "
	gamov.text=" "
	for i=1 to 1
		EVAL("Bumper"&i).hashitevent = 1
	Next
	LeftSlingShot.isdropped=False
	RightSlingShot.isdropped=False
    If B2SOn then
		Controller.B2SSetGameOver 35,0
		Controller.B2SSetTilt 33,0
		Controller.B2SSetMatch 34,0
	End If
	bipreel.setvalue ballinplay
    matchtxt.text=" "
	newball
	ballrelTimer.enabled=true
	newgame.enabled=false
end sub


sub newball
	holevalue=1000
	for each objekt in holelights: objekt.state=0: next
	for each objekt in jokerlights: objekt.state=1: next
	for each light in bonuslights: light.state=0: next
	resetDT
	specialstate=0
	if (balls=3 and (ballinplay=2 or ballinplay=3)) or (balls=5 and ballinplay=5) then DoubleBonus.state=1
End Sub


sub nextball
    if tilt=true then
	  Bumper1.hashitevent = 1
	  LeftSlingShot.isdropped=False
	  RightSlingShot.isdropped=False
      tilt=false
      tilttxt.text=" "
		If B2SOn then
			Controller.B2SSetTilt 33,0
			Controller.B2ssetdata 1, 1
		End If
    end if
	if players=1 or player=players then 
		player=1
	   Else
		player=player+1
	end if
	if player=1 then ballinplay=ballinplay+1
	if ballinplay>balls then
		playsound "endgame",0,.25
		eg=1
		ballreltimer.enabled=true
	  else
		For each objekt in PlayerHuds
			objekt.State = 0
		next
		For each objekt in PlayerHUDScores
			objekt.state=0
		next
		PlayerHuds(Player-1).State =1
		PlayerHUDScores(Player-1).state=1
		PlaySound("RotateThruPlayers"),0,.05,0,0.25
		TempPlayerUp=Player
		PlayerUpRotator.enabled=true
		if state=true then
		  newball
		  ballreltimer.enabled=true
		end if
		bipreel.setvalue ballinplay
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
	  bipreel.setvalue 0
	  gamov.text="GAME OVER"
	  gamov.timerenabled=1
	  tilttxt.timerenabled=1
	  for i=1 to maxplayers
		EVAL("CanPlay"&i).state=0
		if score(i)>hisc then 
			hisc=score(i)
			hiscstate=1
		end if
		EVAL("Pup"&i).state=0
		EVAL("Pups"&i).state=0
	  next
	  if hiscstate=1 then HighScoreEntryInit()
	  UpdatePostIt 
	  savehs
	  If B2SOn then 
        Controller.B2SSetGameOver 35,1
        Controller.B2ssetballinplay 32, 0
	    Controller.B2SSetScorePlayer 5, hisc
	    Controller.B2ssetcanplay 31, 0
	    Controller.B2ssetcanplay 30, 0
	  End If
	  ballreltimer.enabled=false
  else
	Drain.kick 60,45,0
    ballreltimer.enabled=false
	playsound SoundFXDOF("drainkick",135,DOFPulse,DOFContactors) 
  end if
end sub

Sub PlayerUpRotator_timer()
		If RotatorTemp<maxplayers+1 then
			TempPlayerUp=TempPlayerUp+1
			If TempPlayerUp>maxplayers then
				TempPlayerUp=1
			end if
			For each objekt in PlayerHuds
				objekt.state = 0
			next
			For each objekt in PlayerHUDScores
				objekt.state=0
			next
			PlayerHuds(TempPlayerUp-1).State = 1
			PlayerHUDScores(TempPlayerUp-1).state=1
			If B2SOn Then 
				Controller.B2SSetPlayerUp TempPlayerUp
				Controller.B2SSetData 81,0
				Controller.B2SSetData 82,0
				Controller.B2SSetData 83,0
				Controller.B2SSetData 84,0
				Controller.B2SSetData (80+TempPlayerUp),1
			end if
		else
			if B2SOn then 
				Controller.B2SSetPlayerUp Player
				Controller.B2SSetData 81,0
				Controller.B2SSetData 82,0
				Controller.B2SSetData 83,0
				Controller.B2SSetData 84,0
				Controller.B2SSetData (80+Player),1
			end if
			PlayerUpRotator.enabled=false
			RotatorTemp=1
			For each objekt in PlayerHuds
				objekt.State = 0
			next
			For each objekt in PlayerHUDScores
				objekt.state=0
			next
			PlayerHuds(Player-1).State = 1
			PlayerHUDScores(Player-1).state=1
		end if
		RotatorTemp=RotatorTemp+1
end sub

sub matchnum
	if matchnumb=0 then
		matchtxt.text="00"
		If B2SOn then Controller.B2SSetMatch 100
	  else
		matchtxt.text=matchnumb*10
		If B2SOn then Controller.B2SSetMatch 34,Matchnumb*10
	end if
	For i=1 to players
		if (matchnumb*10)=(score(i) mod 100) then 
		  addcredit
		  playsound SoundFXDOF("knock",138,DOFPulse,DOFKnocker)
		  DOF 139,DOFPulse
	    end if
    next
end sub

'********** Bumpers

Sub Bumper1_Hit
   if tilt=false then
	playsound SoundFXDOF("fx_bumper4",107,DOFPulse,DOFContactors)
	DOF 108,DOFPulse
	FlashBumpers
	if balls=3 then
		addscore 1000
	  else
		addscore 100
	end if
   end if
End Sub

sub FlashBumpers
	if bumperlight1.state = 1 then
		for each light in BumperLights: light.duration 0, 200, 1:Next
	end If
end sub

'************** Slings

Sub RightSlingShot_Slingshot
	playsound SoundFXDOF("right_slingshot",105,DOFPulse,DOFContactors), 0, 1, 0.05, 0.05
	DOF 106,DOFPulse
	addscore 10
    RSling.Visible = 0
    RSling1.Visible = 1
	slingR.objroty = -15	
    RStep = 1
    RightSlingShot.TimerEnabled = 1
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
	addscore 10
    LSling.Visible = 0
    LSling1.Visible = 1
	slingL.objroty = 15	
    LStep = 1
    LeftSlingShot.TimerEnabled = 1
End Sub

Sub LeftSlingShot_Timer
    Select Case LStep
        Case 3:LSLing1.Visible = 0:LSLing2.Visible = 1:slingL.objroty = 7
        Case 4:slingL.objroty = 0:LSLing2.Visible = 0:LSLing.Visible = 1:LeftSlingShot.TimerEnabled = 0
    End Select
    LStep = LStep + 1
End Sub

'********** Dingwalls - animated - timer 50

sub dingwalla_hit
	if state=true and tilt=false then addscore 10
	SlingA.visible=0
	SlingA1.visible=1
	dingwalla.uservalue=1
	Me.timerenabled=1
end sub

sub dingwalla_timer
	select case dingwalla.uservalue
		Case 1: SlingA1.visible=0: SlingA.visible=1
		case 2:	SlingA.visible=0: SlingA2.visible=1
		Case 3: SlingA2.visible=0: SlingA.visible=1: Me.timerenabled=0
	end Select
	dingwalla.uservalue=dingwalla.uservalue+1
end sub

sub dingwallb_hit
	if state=true and tilt=false then addscore 10
	SlingB.visible=0
	Slingb1.visible=1
	DingwallB.uservalue=1
	Me.timerenabled=1
end sub

sub dingwallb_timer
	select case dingwallb.uservalue
		Case 1: Slingb1.visible=0: SlingB.visible=1
		case 2:	Slingb.visible=0: Slingb2.visible=1
		Case 3: Slingb2.visible=0: Slingb.visible=1: Me.timerenabled=0
	end Select
	dingwallb.uservalue=DingwallB.uservalue+1
end sub

sub dingwallc_hit
	if state=true and tilt=false then addscore 10
	Slingc.visible=0
	Slingc1.visible=1
	dingwallc.uservalue=1
	Me.timerenabled=1
end sub

sub dingwallc_timer
	select case dingwallc.uservalue
		Case 1: Slingc1.visible=0: Slingc.visible=1
		case 2:	Slingc.visible=0: Slingc2.visible=1
		Case 3: Slingc2.visible=0: Slingc.visible=1: Me.timerenabled=0
	end Select
	dingwallc.uservalue=dingwallc.uservalue+1
end sub

sub dingwalld_hit
	Slingd.visible=0
	Slingd1.visible=1
	dingwalld.uservalue=1
	Me.timerenabled=1
end sub

sub dingwalld_timer
	select case dingwalld.uservalue
		Case 1: Slingd1.visible=0: Slingd.visible=1
		case 2:	Slingd.visible=0: Slingd2.visible=1
		Case 3: Slingd2.visible=0: Slingd.visible=1: Me.timerenabled=0
	end Select
	DingwallD.uservalue=dingwalld.uservalue+1
end sub

'********** Triggers     

sub TGoutL_hit
	DOF 124, DOFPulse
    if (LoutL.state)=lightstateon then
		addscore 3000
    else
		addscore 500
    end if
end sub    

sub TGinL_hit
	DOF 125, DOFPulse
    if (LinL.state)=lightstateon then
		addscore 3000
    else
		addscore 500
    end if
end sub

sub TGoutR_hit
	DOF 126, DOFPulse
    if (LoutR.state)=lightstateon then
		addscore 3000
    else
		addscore 500
    end if
end sub

sub TGinR_hit
	DOF 127, DOFPulse
    if (LinR.state)=lightstateon then
		addscore 3000
    else
		addscore 500
    end if
end sub

sub TGtopoutL_hit   '***** top L rollover
	DOF 119, DOFPulse
	if (LtopoutL.state)=lightstateon then
		addscore 3000
    else
		addscore 505
    end if
end sub    

sub TGtopoutR_hit  	'****** top R rollover
	DOF 123, DOFPulse
	if LtopoutR.state=1 then 
		addscore 3000
	  else
		addscore 500
	end if
	DiverterFlipper.RotateToEnd
	playsound "fx_spinner"
end sub


sub TGjokerG_hit   '****** green joker rollover
	DOF 120, DOFPulse
	addscore 500
	if (LjokerG1.state)=lightstateon then
		LjokerG1.state=0		
		LjokerG.state=0	
		LholeG.state=1
		LinR.state=1
		LoutL.state=1
		holevalue=holevalue+1000
    end if
	checkspecial
end sub

sub TGjokerW_hit   '****** white joker rollover
	DOF 121, DOFPulse
	addscore 500
	if (LjokerW1.state)=lightstateon then
		LjokerW1.state=0		
		LjokerW.state=0	
		LholeW.state=1
		LtopoutL.state=1
		LtopoutR.state=1
		holevalue=holevalue+1000
    end if
	checkspecial
end sub

sub TGjokerP_hit   '****** purple joker rollover
	DOF 122, DOFPulse
	addscore 500
	if (LjokerP1.state)=lightstateon then
		LjokerP1.state=0		
		LjokerP.state=0	
		LholeP.state=1
		LinL.state=1
		LoutR.state=1
		holevalue=holevalue+1000
    end if
	checkspecial
end sub

sub TGballgate_hit
	DiverterFlipper.RotateToStart
	playsound "fx_spinner"
end sub



'********** Targets

sub Tjacks10_hit
	PlaySound SoundFXDOF("target",130, DOFPulse, DOFTargets), 0, .7, 0, 0.05
	addscore 500
end sub

sub Tacekings_hit
	PlaySound SoundFXDOF("target",128, DOFPulse, DOFTargets), 0, .7, 0, 0.05
	addscore 500
end sub

sub Tqueens_hit
	PlaySound SoundFXDOF("target",129, DOFPulse, DOFTargets), 0, .7, 0, 0.05
	if L3000whenLit.state=1 then 
		addscore 3000
	  else
		addscore 500
	end if
end sub

sub Tpurple_hit
	PlaySound SoundFXDOF("target",116, DOFPulse, DOFTargets), 0, .7, .5, 0.05
	addscore 500
	if LholeP.state=0 then
		LjokerP1.state=0		
		LjokerP.state=0	
		LholeP.state=1
		LinL.state=1
		LoutR.state=1
		holevalue=holevalue+1000
    end if
	checkspecial
end sub

sub Twhite_hit
	PlaySound SoundFXDOF("target",115, DOFPulse, DOFTargets), 0, .7, 0, 0.05
	addscore 500
	if LholeW.state=0 then
		LjokerW1.state=0		
		LjokerW.state=0	
		LholeW.state=1
		LtopoutL.state=1
		LtopoutR.state=1
		holevalue=holevalue+1000
    end if
	checkspecial
end sub

sub Tgreen_hit
	PlaySound SoundFXDOF("target",117, DOFPulse, DOFTargets), 0, .7, -.5, 0.05
	addscore 500
	if Lholeg.state=0 then
		LjokerG1.state=0		
		LjokerG.state=0	
		LholeG.state=1
		LinR.state=1
		LoutL.state=1
		holevalue=holevalue+1000
    end if
	checkspecial
end sub

sub checkspecial
	if Lholeg.state+LholeP.state+LholeW.state=3 then
		Lkicker.state=1
		specialstate=1
	end if
end sub

'********** Drop Targets


sub DTA_dropped
	playsound SoundFXDOF("drop1",111,DOFPulse,DOFDropTargets)
	if tilt=false then addscore holevalue
	checkbonus
end sub

sub DTKS_dropped
	playsound SoundFXDOF("drop1",111,DOFPulse,DOFDropTargets)
	if tilt=false then addscore holevalue
	checkbonus
end sub

sub DTKH_dropped
	playsound SoundFXDOF("drop1",111,DOFPulse,DOFDropTargets)
	if tilt=false then addscore holevalue
	checkbonus
end sub

sub DTQS_dropped
	playsound SoundFXDOF("drop1",111,DOFPulse,DOFDropTargets)
	if tilt=false then addscore holevalue
	checkbonus
end sub

sub DTQH_dropped
	playsound SoundFXDOF("drop1",113,DOFPulse,DOFDropTargets)
	if tilt=false then addscore holevalue
	checkbonus
end sub

sub DTQC_dropped
	playsound SoundFXDOF("drop1",113,DOFPulse,DOFDropTargets)
	if tilt=false then addscore holevalue
	checkbonus
end sub

sub DTJH_dropped
	playsound SoundFXDOF("drop1",113,DOFPulse,DOFDropTargets)
	if tilt=false then addscore holevalue
	checkbonus
end sub

sub DTJC_dropped
	playsound SoundFXDOF("drop1",113,DOFPulse,DOFDropTargets)
	if tilt=false then addscore holevalue
	checkbonus
end sub

sub DT10H_dropped
	playsound SoundFXDOF("drop1",113,DOFPulse,DOFDropTargets)
	if tilt=false then addscore holevalue
	checkbonus
end sub

sub checkbonus
	if (DTJC.isdropped and dtjh.isdropped) or (dtks.isdropped and dtkh.isdropped) then Lbonus1.state=1
	if DTJC.isdropped and dtjh.isdropped and dtks.isdropped and dtkh.isdropped then Lbonus2.state=1
	if dtqs.isdropped and dtqh.isdropped and dtqc.isdropped then Lbonus3.state=1
	if (dtqs.isdropped and dtqh.isdropped and dtqc.isdropped) and ((DTJC.isdropped and dtjh.isdropped) or (dtks.isdropped and dtkh.isdropped)) then Lbonus4.state=1
	if dta.isdropped and dtkh.isdropped and dtqh.isdropped and dtjh.isdropped and dt10h.isdropped then Lbonus5.state=1
	if dta.isdropped and dtks.isdropped and dtkh.isdropped and dtqs.isdropped and dtqh.isdropped and dtqc.isdropped and dtjh.isdropped and dtjc.isdropped and dt10h.isdropped then L3000whenLit.state=1
end sub

sub resetDT
	PlaySound SoundFXDOF("DTreset",112,DOFPulse,DOFContactors)
	PlaySound SoundFXDOF("DTreset",114,DOFPulse,DOFContactors)
	for i = 0 to 8
		DropTargets(i).isdropped=False
	next
	L3000whenLit.state=0
end sub

'********** Kicker

Sub kicker_Hit()
	addscore holevalue
	if Lkicker.state=1 then addspecial
	PlaySound "cluper"
	me.uservalue=1
	me.timerenabled=1

End Sub

Sub kicker_timer
	select case kicker.uservalue
	  case 4:
		playsound SoundFXDOF("holekick",118,DOFPulse,DOFContactors)
		DOF 139, DOFPulse
		Kicker.kick 215,6,0
		Pkickarm.rotz=15
	  case 6:
		Pkickarm.rotz=0
		DiverterFlipper.rotatetostart
		me.timerenabled=0
	end Select
	kicker.uservalue=kicker.uservalue+1
End Sub	

'**********  Scoring

sub addscore(points)
  if tilt=false and state=true then
	If points = 10 then 
		matchnumb=matchnumb+1
		if matchnumb>9 then matchnumb=0
		if specialstate=1 and balls=5 then 
			if lkicker.state=1 then
				lkicker.state=0
			  else
				lkicker.state=1
			end if
		end if
	end if
	if points=10 or points=100 or points=1000 then 
		addpoints Points
	  else
		If Points < 100 and AddScore10Timer.enabled = false Then
			Add10 = Points \ 10
			AddScore10Timer.Enabled = TRUE
		  ElseIf Points < 1000 and AddScore100Timer.enabled = false Then
			Add100 = Points \ 100
			AddScore100Timer.Enabled = TRUE
		  ElseIf AddScore1000Timer.enabled = false Then
			Add1000 = Points \ 1000
			AddScore1000Timer.Enabled = TRUE
		End If
	End If
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
	EVAL("ScoreReel"&player).addvalue(points)
	If B2SOn Then Controller.B2SSetScorePlayer player, score(player) MOD 100000

    ' Sounds: there are 3 sounds: tens, hundreds and thousands
    If Points = 100 AND(Score(player) MOD 1000) \ 100 = 0 Then  'New 1000 reel
        PlaySound SoundFXDOF("bell1000",143,DOFPulse,DOFChimes)
      ElseIf Points = 10 AND(Score(player) MOD 100) \ 10 = 0 Then 'New 100 reel
        PlaySound SoundFXDOF("bell100",142,DOFPulse,DOFChimes)
      ElseIf points = 1000 Then
        PlaySound SoundFXDOF("bell1000",143,DOFPulse,DOFChimes)
	  elseif Points = 100 Then
        PlaySound SoundFXDOF("bell100",142,DOFPulse,DOFChimes)
      Else
        PlaySound SoundFXDOF("bell10",141,DOFPulse,DOFChimes)
    End If
    ' check replays and rollover
	if score(player)=>99999 then
		If B2SOn then Controller.B2SSetScoreRollover 24 + player, 1
		EVAL("Reel100K"&player).setvalue (int(score(player)/100000))
	End if
    if score(player)=>replay1 and rep(player)=0 then
		addcredit
		rep(player)=1
		PlaySound SoundFXDOF("knock",138,DOFPulse,DOFKnocker)
		DOF 139, DOFPulse
    end if
    if score(player)=>replay2 and rep(player)=1 then
		addcredit
		rep(player)=2
		PlaySound SoundFXDOF("knock",138,DOFPulse,DOFKnocker)
		DOF 139, DOFPulse
    end if
    if score(player)=>replay3 and rep(player)=2 then
		addcredit
		rep(player)=3
		PlaySound SoundFXDOF("knock",138,DOFPulse,DOFKnocker)
		DOF 139, DOFPulse
    end if
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

Sub MechCheckTilt
	GameTilted
End Sub

Sub Tilttimer_Timer()
	Tilttimer.Enabled = False
End Sub

Sub GameTilted
	Tilt = True
	tilttxt.text="TILT"
    If B2SOn Then Controller.B2SSetTilt 33,1
    If B2SOn Then Controller.B2ssetdata 1, 0
	playsound "tilt"
	turnoff
End Sub

sub turnoff
	for i=1 to 1
		EVAL("Bumper"&i).hashitevent = 0
	Next
	LeftSlingShot.isdropped=true
	RightSlingShot.isdropped=true
  	LeftFlipper.RotateToStart
	StopSound "Buzz"
	DOF 101, DOFOff
	RightFlipper.RotateToStart
	StopSound "Buzz1"
	DOF 102, DOFOff
end sub    

' *********************************************************************
'                      Supporting Ball & Sound Functions
' *********************************************************************

Function Vol(ball) ' Calculates the Volume of the sound based on the ball speed
    Vol = Csng(BallVel(ball) ^2 / 2000)
End Function

Function Pan(ball) ' Calculates the pan for a ball based on the X position on the table. "table1" is the name of the table
    Dim tmp
    tmp = ball.x * 2 / RoyalFlush.width-1
    If tmp > 0 Then
        Pan = Csng(tmp ^10)
    Else
        Pan = Csng(-((- tmp) ^10) )
    End If
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
            PlaySound("fx_ballrolling" & b), -1, Vol(BOT(b)), Pan(BOT(b) ), 0, Pitch(BOT(b) ), 1, 0
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
	PlaySound("fx_collide"), 0, Csng(velocity) ^2 / 2000, Pan(ball1), 0, Pitch(ball1), 0, 0
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
	PlaySound "pinhit_low", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0
End Sub

Sub a_Targets_Hit (idx)
	PlaySound "target", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0
End Sub

Sub a_Metals_Thin_Hit (idx)
	PlaySound "metalhit_thin", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0
End Sub

Sub a_Metals_Medium_Hit (idx)
	PlaySound "metalhit_medium", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0
End Sub

Sub a_Metals2_Hit (idx)
	PlaySound "metalhit2", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0
End Sub

Sub a_Gates_Hit (idx)
	PlaySound "gate4", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0
End Sub

Sub a_Spinner_Spin
	PlaySound "fx_spinner",0,.25,0,0.25
End Sub

Sub a_Rubbers_Hit(idx)
 	dim finalspeed
  	finalspeed=SQR(activeball.velx * activeball.velx + activeball.vely * activeball.vely)
 	If finalspeed > 20 then 
		PlaySound "fx_rubber2", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0
	End if
	If finalspeed >= 6 AND finalspeed <= 20 then
 		RandomSoundRubber()
 	End If
End Sub

Sub RubberWheel_hit
 	dim finalspeed
  	finalspeed=SQR(activeball.velx * activeball.velx + activeball.vely * activeball.vely)
 	If finalspeed > 20 then 
		PlaySound "fx_rubber2", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0
	End if
	If finalspeed >= 6 AND finalspeed <= 20 then
 		RandomSoundRubber()
 	End If
End sub

Sub a_Posts_Hit(idx)
 	dim finalspeed
  	finalspeed=SQR(activeball.velx * activeball.velx + activeball.vely * activeball.vely)
 	If finalspeed > 16 then 
		PlaySound "fx_rubber2", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0
	End if
	If finalspeed >= 6 AND finalspeed <= 16 then
 		RandomSoundRubber()
 	End If
End Sub

Sub RandomSoundRubber()
	Select Case Int(Rnd*3)+1
		Case 1 : PlaySound "rubber_hit_1", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0
		Case 2 : PlaySound "rubber_hit_2", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0
		Case 3 : PlaySound "rubber_hit_3", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0
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
		Case 1 : PlaySound "flip_hit_1", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0
		Case 2 : PlaySound "flip_hit_2", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0
		Case 3 : PlaySound "flip_hit_3", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0
	End Select
End Sub


sub savehs
    savevalue "RoyalFlush", "credit", credit
    savevalue "RoyalFlush", "hiscore", hisc
    savevalue "RoyalFlush", "match", matchnumb
    savevalue "RoyalFlush", "score1", score(1)
    savevalue "RoyalFlush", "score2", score(2)
    savevalue "RoyalFlush", "score3", score(3)
    savevalue "RoyalFlush", "score4", score(4)
	savevalue "RoyalFlush", "balls", balls
	savevalue "RoyalFlush", "hsa1", HSA1
	savevalue "RoyalFlush", "hsa2", HSA2
	savevalue "RoyalFlush", "hsa3", HSA3
	savevalue "RoyalFlush", "freeplay", freeplay

end sub

sub loadhs
    dim temp
	temp = LoadValue("RoyalFlush", "credit")
    If (temp <> "") then credit = CDbl(temp)
    temp = LoadValue("RoyalFlush", "hiscore")
    If (temp <> "") then hisc = CDbl(temp)
    temp = LoadValue("RoyalFlush", "match")
    If (temp <> "") then matchnumb = CDbl(temp)
    temp = LoadValue("RoyalFlush", "score1")
    If (temp <> "") then score(1) = CDbl(temp)
    temp = LoadValue("RoyalFlush", "score2")
    If (temp <> "") then score(2) = CDbl(temp)
    temp = LoadValue("RoyalFlush", "score3")
    If (temp <> "") then score(3) = CDbl(temp)
    temp = LoadValue("RoyalFlush", "score4")
    If (temp <> "") then score(4) = CDbl(temp)
    temp = LoadValue("RoyalFlush", "balls")
    If (temp <> "") then balls = CDbl(temp)
    temp = LoadValue("RoyalFlush", "hsa1")
    If (temp <> "") then HSA1 = CDbl(temp)
    temp = LoadValue("RoyalFlush", "hsa2")
    If (temp <> "") then HSA2 = CDbl(temp)
    temp = LoadValue("RoyalFlush", "hsa3")
    If (temp <> "") then HSA3 = CDbl(temp)
    temp = LoadValue("RoyalFlush", "freeplay")
    If (temp <> "") then freeplay = CDbl(temp)
end sub

Sub RoyalFlush_Exit()
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

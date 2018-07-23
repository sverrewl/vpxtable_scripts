'///////////////////////////////////////////////////////////
'			SPACE WALK by Gottlieb (1979)
'
'
'
' vp10 assembled and scripted by BorgDog, 2015
'
' thanks to hauntfreaks for graphics updates and the indispensable leafswitch primitive
' thanks to zany for the mini-gottlieb flippers
' menu system borrowed from loserman76 and gnance, hold left flipper to bring up game options menu
' big thanks to Scapino for sharing his countdown resources
' 
' DOF by arngrim
'
'///////////////////////////////////////////////////////////

Option Explicit
Randomize

Const cGameName = "spacewalk_1979"

dim award
Dim balls
Dim ebcount
Dim credit, freeplay
Dim score(2)
Dim sreels(2)
Dim player, players, maxplayers
Dim Add10, Add100, Add1000
Dim state
Dim tilt, tiltsens
Dim Replay1Table(3)
Dim Replay2Table(3)
Dim Replay3Table(3)
dim replay1, replay2, replay3, replays
dim hisc, hiscstate, showhisc
Dim bumperlitscore
Dim bumperoffscore
Dim Bonus, dbonus
Dim Bonuslight(19)
Dim ballinplay
Dim matchnumb
dim plungeball
dim tflash
dim eflash
dim fivek
Dim rep(2)
Dim eg
Dim starstate
Dim abonus(2,4)
Dim OperatorMenu, options, objekt
Dim i,j,e,t,a,light

On Error Resume Next
ExecuteGlobal GetTextFile("controller.vbs")
If Err Then MsgBox "Can't open controller.vbs"
On Error Goto 0


Sub SpaceWalk_init
	LoadEM
	maxplayers=2
	Replay1Table(1)=90000
	Replay1Table(2)=120000
	Replay1Table(3)=140000
	Replay2Table(1)=120000
	Replay2Table(2)=140000
	Replay2Table(3)=180000
	Replay3Table(1)=150000
	Replay3Table(2)=180000
	Replay3Table(3)=200000
	set sreels(1) = ScoreReel1
	set sreels(2) = ScoreReel2
	set abonus(1,1)=ABonus1
	set abonus(1,2)=ABonus2
	set abonus(1,3)=ABonus3
	set abonus(1,4)=ABonus4
	set abonus(2,1)=ABonus4
	set abonus(2,2)=ABonus3
	set abonus(2,3)=ABonus2
	set abonus(2,4)=ABonus1
	set bonuslight(1)=bonus1
	set bonuslight(2)=bonus2
	set bonuslight(3)=bonus3
	set bonuslight(4)=bonus4
	set bonuslight(5)=bonus5
	set bonuslight(6)=bonus6
	set bonuslight(7)=bonus7
	set bonuslight(8)=bonus8
	set bonuslight(9)=bonus9
	set bonuslight(10)=bonus10
	set bonuslight(11)=bonus11
	set bonuslight(12)=bonus12
	set bonuslight(13)=bonus13
	set bonuslight(14)=bonus14
	set bonuslight(15)=bonus15
	set bonuslight(16)=bonus16
	set bonuslight(17)=bonus17
	set bonuslight(18)=bonus18
	set bonuslight(19)=bonus19
	turnoff
	hideoptions
	balls=5
	replays=2
	hisc=50000
	showhisc=1
	freeplay=0
	HSA1=4
	HSA2=15
	HSA3=7
	loadhs
	bipreel.setvalue 0
	creditreel.setvalue credit
	Replay1=Replay1Table(Replays)
	Replay2=Replay2Table(Replays)
	Replay3=Replay3Table(Replays)
	player=1
	plungeball=0
	if balls=3 then
		bumperlitscore=1000
		bumperoffscore=1000
		InstCard.image="InstCard3balls"
	  else
		bumperlitscore=100
		bumperoffscore=100
		InstCard.image="InstCard5balls"
	end if
	OptionBalls.imageA="OptionsBalls"&Balls
	OptionReplays.imageA="OptionsReplays"&replays
	OptionFreeplay.imageA="OptionsFreeplay"&freeplay
	OptionShowhisc.imageA="OptionsYN"&showhisc
	RepCard.image = "ReplayCard"&replays
	If showhisc=0 then 
		for each objekt in hiscStuff: objekt.visible = 0: next
	  else
		UpdatePostIt
	end if
	If B2SOn then
		for each objekt in backdropstuff : objekt.visible = 0 : next
	End If
    scorereel1.setvalue(score(1))
    scorereel2.setvalue(score(2))
	tilt=false
	startGame.enabled=true
    Drain.CreateBall
End Sub

sub startGame_timer
	playsound "poweron"
	lightdelay.enabled=true
	me.enabled=false
end sub

sub lightdelay_timer
	If credit > 0 or freeplay=1 Then DOF 121, DOFOn
	for i = 1 to maxplayers
		if score(i)>99999 then EVAL("p100k"&i).setvalue 1 'int(score(i)/100000)
	next
	gamov.text="Game Over"
	gamov.timerenabled=1
	tilttxt.timerenabled=1
	For each light in GIlights:light.state=1:Next
	BumperLight1.state=1
	if matchnumb=0 then
		matchtxt.text="00"
	  else
		matchtxt.text=matchnumb*10
	end if
	if B2SOn then
		for i = 1 to 2
			Controller.B2SSetScorePlayer i, Score(i) MOD 100000
			if score(i)>99999 then Controller.B2SSetScoreRollover 24+i, 1
		next
		Controller.B2SSetData 6,1
		Controller.B2ssetCredits Credit
		Controller.B2ssetMatch 34, Matchnumb
		Controller.B2SSetGameOver 35,1
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

Sub SpaceWalk_KeyDown(ByVal keycode)

	if keycode=AddCreditKey then
		playsound "coin3" 
		coindelay.enabled=true 
    end if

    if keycode=StartGameKey and (credit>0 or freeplay=1) And Not HSEnterMode=true and startGame.enabled=0 and lightdelay.enabled=0 then
	  if state=false then
		state=true
		if freeplay=0 then credit=credit-1
		If credit < 1 and freeplay=0 Then DOF 121, DOFOff
		if freeplay=0 then creditreel.setvalue credit
		ballinplay=1
		If B2SOn Then 
			Controller.B2ssetCredits Credit
			Controller.B2sStartAnimation "Startup"
			Controller.B2ssetballinplay 32, Ballinplay
			Controller.B2ssetplayerup 30, 1
			Controller.B2ssetcanplay 31, 1
			Controller.B2SSetGameOver 0
		End If
		Pup1.setvalue(1)
		tilt=false
		playsound "initialize" 
		players=1
		canplayreel.setvalue(players)
		newgame.enabled=true
	  else if state=true and players < 2 and Ballinplay=1 then
		if freeplay=0 then credit=credit-1
		If credit < 0 and freeplay=0 Then DOF 121, DOFOff
		players=players+1
		if freeplay=0 then creditreel.setvalue credit
		If B2SOn then
			Controller.B2ssetCredits Credit
			Controller.B2ssetcanplay 31, 2
		End If
		canplayreel.setvalue(players)
		playsound "click" 
	   end if 
	  end if
	end if

	If HSEnterMode Then HighScoreProcessKey(keycode)

	If keycode = PlungerKey Then
		Plunger.PullBack
		PlaySound "plungerpull",0,1,0.25,0.25
	End If

	If keycode=LeftFlipperKey and State = false and OperatorMenu=0 And Not HSEnterMode=true then
		OperatorMenuTimer.Enabled = true
	end if

	If keycode=LeftFlipperKey and State = false and OperatorMenu=1 then
		Options=Options+1
		If Options=6 then Options=1
		playsound "drop1"
		Select Case (Options)
			Case 1:
				Option1.visible=true
				Option5.visible=False
			Case 2:
				Option2.visible=true
				Option1.visible=False
			Case 3:
				Option3.visible=true
				Option2.visible=False
			Case 4:
				Option4.visible=true
				Option3.visible=False
			Case 5:
				Option5.visible=true
				Option4.visible=False
		End Select
	end if

	If keycode=RightFlipperKey and State = false and OperatorMenu=1 then
	  PlaySound "cluper"
	  Select Case (Options)
		Case 1:
			if Balls=3 then
				Balls=5
				InstCard.image="InstCard5balls"
			  else
				Balls=3
				InstCard.image="InstCard3balls"
			end if
			OptionBalls.imageA = "OptionsBalls"&Balls 
		Case 2:
			if Freeplay=0 then
				Freeplay=1
			  else
				Freeplay=0
			end if
			OptionFreeplay.imageA = "OptionsFreeplay"&freeplay
		Case 3:
			if showhisc=0 then
				showhisc=1
				for each objekt in hiscStuff: objekt.visible = 1: next
				UpdatePostIt
			  else
				for each objekt in hiscStuff: objekt.visible = 0: next
				showhisc=0
			end if
			OptionShowhisc.imageA = "OptionsYN"&showhisc
		Case 4:
			Replays=Replays+1
			if Replays>3 then
				Replays=1
			end if
			Replay1=Replay1Table(Replays)
			Replay2=Replay2Table(Replays)
			Replay3=Replay3Table(Replays)
			OptionReplays.imageA = "OptionsReplays"&replays
			repcard.image = "replaycard"&replays
		Case 5:
			OperatorMenu=0
			savehs
			HideOptions
	  End Select
	End If

  if tilt=false and state=true then
	If keycode = LeftFlipperKey Then
		LeftFlip.RotateToEnd
		LeftFlip1.RotateToEnd
		if state=true then TLLFlip.state=1
		PlaySound SoundFXDOF("flipperup",101,DOFOn,DOFFlippers), 0, .67, -0.05, 0.05
		PlaySound "Buzz",-1,.05,-0.05, 0.05
	End If
    
	If keycode = RightFlipperKey Then
		RightFlip.RotateToEnd
		RightFlip1.RotateToEnd
		if state=true then TLRFlip.state=1
		PlaySound SoundFXDOF("flipperup",102, DOFOn, DOFFlippers), 0, .67, 0.05, 0.05
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
	OptionShowhisc.visible = true
End Sub

Sub HideOptions
	for each objekt In OptionMenu
		objekt.visible = false
	next
End Sub


Sub SpaceWalk_KeyUp(ByVal keycode)

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

   if tilt=false and state=true then   
	If keycode = LeftFlipperKey Then
		LeftFlip.RotateToStart
		LeftFlip1.RotateToStart
		TLLFlip.state=0
		PlaySound SoundFXDOF("flipperdown", 101, DOFOff, DOFContactors), 0, 1, -0.05, 0.05
		StopSound "Buzz"
	End If
    
	If keycode = RightFlipperKey Then
		RightFlip.RotateToStart
		RightFlip1.RotateToStart
		TLRFlip.state=0
		PlaySound SoundFXDOF("flipperdown", 102, DOFOff, DOFContactors), 0, 1, -0.05, 0.05
		StopSound "Buzz1"
	End If
  end if
End Sub

sub newgame_timer
    bumper1.hashitevent=1
	fivek=0
	player=1
    pup1.setvalue(1)
	scorereel1.resettozero 
	scorereel2.resettozero 
	for i = 1 to maxplayers
	    score(i)=0
		rep(i)=0
		EVAL("p100k"&i).setvalue 0
	next
	ebcount=0
	If B2SOn then
		Controller.B2SSetScore 1, score(1)
		Controller.B2SSetScore 2, score(2)
		Controller.B2SSetData 3,0
		Controller.B2SSetData 4,0
		Controller.B2SSetScoreRollover 25, 0
		Controller.B2SSetScoreRollover 26, 0
	End If
    eg=0
    gamov.text=" "
    tilttxt.text=" "
    If B2SOn then
		Controller.B2SSetGameOver 35,0
		Controller.B2SSetTilt 33,0
		Controller.B2SSetMatch 34,0
		Controller.B2ssetdata 11, 0
		Controller.B2ssetdata 7,0
		Controller.B2ssetdata 8,0
	End If
	bipreel.setvalue 1
	matchtxt.text=" "
	me.enabled=false
    ballreltimer.enabled=true
end sub

sub nextball
    if tilt=true then
	  bumper1.hashitevent=1
      tilt=false
      tilttxt.text=" "
		If B2SOn then
			Controller.B2SSetTilt 33,0
			Controller.B2ssetdata 1, 1
		End If
    end if
	if player=1 then ballinplay=ballinplay+1
	if ballinplay>balls then
		playsound "endgame"
		eg=1
		ballreltimer.enabled=true
	  else
		if state=true and tilt=false then
		  ebcount=0
		  ballreltimer.enabled=true
		end if
		bipreel.setvalue ballinplay
		If B2SOn then Controller.B2ssetballinplay 32, Ballinplay
	end if
End Sub

sub ballreltimer_timer
  for each light in bonuslights: light.state=0: next
  if eg=1 then
	  turnoff
      matchnum
	  bipreel.setvalue 0
	  state=false
	  gamov.text="GAME OVER"
	  gamov.timerenabled=1
	  tilttxt.timerenabled=1
	  canplayreel.setvalue(0)
	  pup1.setvalue(0)
	  pup2.setvalue(0)
	  hiscstate=0
	  for i=1 to maxplayers
		if score(i)>hisc then 
			hisc=score(i)
			hiscstate=1
		end if
		EVAL("Pup"&i).setvalue 0
	  next
	  if hiscstate=1 and showhisc=1 then 
		HighScoreEntryInit()
		UpdatePostIt 
		savehs
	  end if
	  If B2SOn then 
        Controller.B2SSetGameOver 35,1
        Controller.B2ssetballinplay 32, 0
	    Controller.B2sStartAnimation "EOGame"
	    Controller.B2ssetcanplay 31, 0
	    Controller.B2ssetcanplay 30, 0
	  End If
	  ballreltimer.enabled=false
  else

	L5k.state=0
	L1k.state=0
	LTrigLStar.state=0
	Special.state=0
	shootagain.state=0
	ExtraBall.state=0
	bonus=0	
	starstate=0
	BumperLight1.State=1
	resetDT.enabled=1
	for i = 1 to 4
		abonus(1,i).state=0
	next
	if ballinplay=balls then bonusx2.state=1
	playsound "drainkick"
 	Drain.kick 60,45,0
	DOF 118, 2
    ballreltimer.enabled=false
  end if
end sub

sub matchnum
	matchnumb=INT (RND*10)
	if matchnumb=0 then
		matchtxt.text="00"
	  else
		matchtxt.text=matchnumb*10
	end if
  If B2SOn then Controller.B2SSetMatch 34,Matchnumb*10
  For i=1 to players
    if (matchnumb*10)=(score(i) mod 100) then 
		addcredit
		playsound SoundFXDOF("knocker", 117, DOFPulse, DOFKnocker)
		DOF 116, DOFPulse		
	end if
  next
end sub

sub flippertimer_timer()
	LFlip.RotY = LeftFlip.CurrentAngle
	LFlip1.RotY = LeftFlip.CurrentAngle-90
	RFlip.RotY = RightFlip.CurrentAngle
	RFlip1.RotY = RightFlip.CurrentAngle-90
	PGate.Rotz = (gate1.CurrentAngle*.75) + 25
	BumperLightA1.state=bumperlight1.state
	Special1.state = Special.state
	L5k1.state = L5k.state
	L5k2.state = L5k.state
	L5k3.state = L5k.state
	L5k4.state = L5k.state
	L5k5.state = L5k.state
	L1k1.state = L1k.state
	LTrigRStar.state = LTrigLStar.state
end sub

sub coindelay_timer
	addcredit
    coindelay.enabled=false
end sub

Sub UpperKicker_Hit
	addscore 500
	addbonus
	if extraball.state=1 then
		playsound "ding1",0,.5
		if ebcount=0 then
			shootagain.state=lightstateon
			extraball.state=0
			ebcount=1
		 end if
	end if
	me.uservalue=1
	me.timerenabled=1
End Sub

Sub UpperKicker_timer
	select case UpperKicker.uservalue
	  case 4:
		UpperKicker.Kick 173, 15
		PlaySound SoundFXDOF("holekick",115,DOFPulse,DOFContactors),0,1,0,0.25
		DOF 116, DOFPulse
		Pkickarm.rotz=15
	  case 6:
		Pkickarm.rotz=0
		me.timerenabled=0
	end Select
	Upperkicker.uservalue=Upperkicker.uservalue+1
End Sub

Sub Drain_Hit()
	PlaySound "drain",0,1,0,0.25
	DOF 119, DOFPulse
	if l5k.state=1 then 
		fivek=1
	  else
		fivek=0
	end if
	if bonusx3.state=1 then 
		dbonus=3
	  elseif bonusx2.state=1 then 
			dbonus=2
		  else 
			dbonus=1
    end if
	bonus=bonus*dbonus
	scorebonus.enabled=1
End Sub

sub ballhome_hit
	plungeball=1
	shootagain.state=lightstateoff
	If B2SOn then Controller.B2SSetShootAgain 36,0
end sub

sub ballhome_unhit
	DOF 120, DOFPulse
	plungeball=0
end sub

sub scorebonus_timer
   if tilt=true then
		bonus=0
	 Else
		if bonus>0 then
			if bonus/2 = int(bonus/2) then
				BumperLight1.State=0
				if fivek=1 then 
					l5k.state=0
				end if
			  else
				BumperLight1.State=1
				if fivek=1 then 
					l5k.state=1
				end if
			end if
			if bonus/dbonus = int(bonus/dbonus) then
				if bonus/dbonus<19 then bonuslight((bonus/dbonus)+1).state=0
				if bonus > dbonus-1 then bonuslight((bonus/dbonus)).state=1
				bonus=bonus-1
				addpoints 1000
			  else
				bonus=bonus-1
				addpoints 1000
			  end if
		   else 
			bonus=0
		end if
		if bonus=0 then 
		  fivek=0
		  if shootagain.state=lightstateon then
			  ballreltimer.enabled=true
		   else
			  if players=1 or player=2 then 
				player=1
				If B2SOn then Controller.B2ssetplayerup 30, 1
				pup1.setvalue(1)
				pup2.setvalue(0)
				nextball
			  else
				player=2
				If B2SOn then Controller.B2ssetplayerup 30, 2
				pup2.setvalue(1)
				pup1.setvalue(0)
				nextball
			  end if
		  end if
		  scorebonus.enabled=false
		end if
	end if
End sub

sub awardcheck
	award=0
	for a = 4 to 11	
		if Targets(a).isdropped then award=award+1
	next
	if award = 8 then
		bonusx2.state=1
		extraball.state=1
		l1k.state=1
		l5k.state=1
		if ballinplay=balls then bonusx3.state=1
	end if
	award=0
	for a = 0 to 3
		if targets(a).isdropped then award=award+1
	next
	if award = 4 then
		bonusx2.state=1
		l1k.state=1
		l5k.state=1
		if extraball.state=1 then special.state=1
		if ballinplay=balls then bonusx3.state=1
	end if
	award=0
	for a = 12 to 15
		if Targets(a).isdropped then award=award+1
	next
	if award = 4 then
		bonusx2.state=1
		l1k.state=1
		l5k.state=1
		if extraball.state=1 then special.state=1
		if ballinplay=balls then bonusx3.state=1
	end if
end sub

'**** Bumper

Sub Bumper1_Hit
   if tilt=false then
    PlaySound SoundFXDOF("fx_bumper4",107,DOFPulse,DOFContactors)
	DOF 108, DOFPulse
	if BumperLight1.state=1 then
		addscore bumperlitscore
	  else
		addscore bumperoffscore
	end if
   end if
End Sub

sub FlashBumper
	BumperLight1.duration 0, 500, 1
end sub


'**** Wire Triggers

sub TrigLUpper_hit
	DOF 113, DOFPulse
	addscore 500
	addbonus
end sub

sub TrigRUpper_hit
	DOF 114, DOFPulse
	addscore 500
	addbonus
end sub

sub TrigLeftLane_hit
	DOF 110, DOFPulse
	if L1k.state=1 then
		addscore 1000
	  else
		addscore 100
	end if
end sub

sub TrigRightLane_hit
	DOF 111, DOFPulse
	if L1k.state=1 then
		addscore 1000
	  else
		addscore 100
	end if
end sub

sub TrigLOut_hit
	DOF 109, DOFPulse
	addbonus
	if L5k.state=1 then
		addscore 5000
	  else
		addscore 500
	end if
end sub

sub TrigROut_hit
	DOF 112, DOFPulse
	addbonus
	if L5k.state=1 then
		addscore 5000
	  else
		addscore 500
	end if
end sub

'**** Star Triggers

sub TrigRStar_hit
	if LTrigRStar.state=1 then
		addbonus
	  else
		starstate=starstate+1
		if starstate=5 then 
			LTrigLStar.state=1
			abonus(player,4).state=0
		  else
			abonus(player,starstate).state=1
			if (starstate-1)>0 then abonus(player,(starstate-1)).state=0
		end if
	end if
end sub

sub TrigLStar_hit
	if LTrigLStar.state=1 then
		addbonus
	  else
		starstate=starstate+1
		if starstate=5 then 
			LTrigLStar.state=1
			abonus(player,4).state=0
		  else
			abonus(player,starstate).state=1
			if (starstate-1)>0 then abonus(player,(starstate-1)).state=0
		end if
	end if
end sub


'**** Drop Targets

Sub DTred1_dropped
  if tilt=false then
	playsound SoundFXDOF("target",103,DOFPulse,DOFContactors)
	flashbumper
	addbonus
	if L5K.state=1 then
		addscore 5000
	  else
		addscore 500
	end if
	AwardCheck
	TLred1.state = 1
	TLred1b.state = 1
  end if
End Sub

Sub DTred2_dropped
  if tilt=false then
	playsound SoundFXDOF("target",103,DOFPulse,DOFContactors)
	flashbumper
	addbonus
	if L5K.state=1 then
		addscore 5000
	  else
		addscore 500
	end if
	AwardCheck
	TLred2.state=1
	TLred2b.state=1
  end if
End Sub

Sub DTred3_dropped
  if tilt=false then
	playsound SoundFXDOF("target",103,DOFPulse,DOFContactors)
	flashbumper
	addbonus
	if L5K.state=1 then
		addscore 5000
	  else
		addscore 500
	end if
	AwardCheck
	TLred3.state=1
  end if
End Sub

Sub DTred4_dropped
  if tilt=false then
	playsound SoundFXDOF("target",103,DOFPulse,DOFContactors)
	flashbumper
	addbonus
	if L5K.state=1 then
		addscore 5000
	  else
		addscore 500
	end if
	AwardCheck
		TLred4.state=1
  end if
End Sub

Sub DTblue1_dropped
  if tilt=false then
	playsound SoundFXDOF("target",104,DOFPulse,DOFContactors)
	flashbumper
	addbonus
	if L5K.state=1 then
		addscore 5000
	  else
		addscore 500
	end if
	AwardCheck
	TLblue1.state=1
  end if
End Sub

Sub DTblue2_dropped
  if tilt=false then
	playsound SoundFXDOF("target",104,DOFPulse,DOFContactors)
	flashbumper
	addbonus
	if L5K.state=1 then
		addscore 5000
	  else
		addscore 500
	end if
	AwardCheck
	TLblue2.state=1
  end if
End Sub


Sub DTblue3_dropped
  if tilt=false then
	playsound SoundFXDOF("target",104,DOFPulse,DOFContactors)
	flashbumper
	addbonus
	if L5K.state=1 then
		addscore 5000
	  else
		addscore 500
	end if
	AwardCheck
	TLblue3.state=1
	TLblue3b.state=1
  end if
End Sub

Sub DTblue4_dropped
  if tilt=false then
	playsound SoundFXDOF("target",104,DOFPulse,DOFContactors)
	flashbumper
	addbonus
	if L5K.state=1 then
		addscore 5000
	  else
		addscore 500
	end if
	AwardCheck
	TLblue4.state=1
	TLblue4b.state=1
  end if
End Sub

Sub DTblack1_dropped
  if tilt=false then
	playsound SoundFXDOF("target",105,DOFPulse,DOFContactors)
	flashbumper
	addbonus
	if L5K.state=1 then
		addscore 5000
	  else
		addscore 500
	end if
	AwardCheck
	TLblack1.state=1
  end if
End Sub

Sub DTblack2_dropped
  if tilt=false then
	playsound SoundFXDOF("target",105,DOFPulse,DOFContactors)
	flashbumper
	addbonus
	if L5K.state=1 then
		addscore 5000
	  else
		addscore 500
	end if
	AwardCheck
	TLblack2.state=1
  end if
End Sub

Sub DTblack3_dropped
  if tilt=false then
	playsound SoundFXDOF("target",105,DOFPulse,DOFContactors)
	flashbumper
	addbonus
	if L5K.state=1 then
		addscore 5000
	  else
		addscore 500
	end if
	AwardCheck
	TLblack3.state=1
  end if
End Sub

Sub DTblack4_dropped
  if tilt=false then
	playsound SoundFXDOF("target",105,DOFPulse,DOFContactors)
	flashbumper
	addbonus
	if L5K.state=1 then
		addscore 5000
	  else
		addscore 500
	end if
	AwardCheck
	TLblack4.state=1
  end if
End Sub

Sub DTblack5_dropped
  if tilt=false then
	playsound SoundFXDOF("target",106,DOFPulse,DOFContactors)
	flashbumper
	addbonus
	if L5K.state=1 then
		addscore 5000
	  else
		addscore 500
	end if
	AwardCheck
	TLblack5.state=1
  end if
End Sub

Sub DTblack6_dropped
  if tilt=false then
	playsound SoundFXDOF("target",106,DOFPulse,DOFContactors)
	flashbumper
	addbonus
	if L5K.state=1 then
		addscore 5000
	  else
		addscore 500
	end if
	AwardCheck
	TLblack6.state=1
  end if
End Sub

Sub DTblack7_dropped
  if tilt=false then
	playsound SoundFXDOF("target",106,DOFPulse,DOFContactors)
	flashbumper
	addbonus
	if L5K.state=1 then
		addscore 5000
	  else
		addscore 500
	end if
	AwardCheck
	TLblack7.state=1
  end if
End Sub

Sub DTblack8_dropped
  if tilt=false then
	playsound SoundFXDOF("target",106,DOFPulse,DOFContactors)
	flashbumper
	addbonus
	if L5K.state=1 then
		addscore 5000
	  else
		addscore 500
	end if
	AwardCheck
	TLblack8.state=1
  end if
End Sub

Sub resetDT_timer
	playsound SoundFXDOF("bankreset",126,DOFPulse,DOFContactors)
	for j = 0 to 15
		targets(j).isdropped = false
	next
	for each light in dtlights:light.state=0:next
	me.enabled=0
end sub

'********Rubber walls

sub DingwallI_Hit   'top left rubber by kicker
	if state=true and tilt=false then 
		addscore 10
		if LTrigLStar.state=0 then
			starstate=starstate+1
			if starstate=5 then 
				LTrigLStar.state=1
				abonus(player,4).state=0
			  else
				abonus(player,starstate).state=1
				if (starstate-1)>0 then abonus(player,(starstate-1)).state=0
			end if
		end if
	END IF
	SlingI.visible=0
	SlingG.visible=0
	SlingI1.visible=1
	me.uservalue=1
	Me.timerenabled=1
end sub

sub dingwallI_timer									'default 50 timer
	select case dingwallI.uservalue
		Case 1: SlingI1.visible=0: SlingI.visible=1
		case 2:	SlingI.visible=0: SlingI2.visible=1
		Case 3: SlingI2.visible=0: SlingI.visible=1: SlingG.visible=1: Me.timerenabled=0
	end Select
	me.uservalue=me.uservalue+1
end sub

sub DingwallJ_Hit    ' top right rubber by kicker
	if state=true and tilt=false then 
		addscore 10
		if LTrigLStar.state=0 then
			starstate=starstate+1
			if starstate=5 then 
				LTrigLStar.state=1
				abonus(player,4).state=0
			  else
				abonus(player,starstate).state=1
				if (starstate-1)>0 then abonus(player,(starstate-1)).state=0
			end if
		end if
	END IF
	SlingJ.visible=0
	SlingH.visible=0
	SlingJ1.visible=1
	me.uservalue=1
	Me.timerenabled=1
end sub

sub dingwallJ_timer									'default 50 timer
	select case dingwallJ.uservalue
		Case 1: SlingJ1.visible=0: SlingJ.visible=1
		case 2:	SlingJ.visible=0: SlingJ2.visible=1
		Case 3: SlingJ2.visible=0: SlingJ.visible=1: SlingH.visible=1: Me.timerenabled=0
	end Select
	me.uservalue=me.uservalue+1
end sub

sub dingwalla_hit
	if state=true and tilt=false then addscore 10
	SlingA.visible=0
	SlingA1.visible=1
	me.uservalue=1
	Me.timerenabled=1
end sub

sub dingwalla_timer									'default 50 timer
	select case dingwalla.uservalue
		Case 1: SlingA1.visible=0: SlingA.visible=1
		case 2:	SlingA.visible=0: SlingA2.visible=1
		Case 3: SlingA2.visible=0: SlingA.visible=1: Me.timerenabled=0
	end Select
	me.uservalue=me.uservalue+1
end sub

sub dingwallb_hit
	if state=true and tilt=false then addscore 10
	SlingB.visible=0
	SlingB1.visible=1
	me.uservalue=1
	Me.timerenabled=1
end sub

sub dingwallb_timer									'default 50 timer
	select case DingwallB.uservalue
		Case 1: Slingb1.visible=0: SlingB.visible=1
		case 2:	SlingB.visible=0: Slingb2.visible=1
		Case 3: Slingb2.visible=0: SlingB.visible=1: Me.timerenabled=0
	end Select
	DingwallB.uservalue=DingwallB.uservalue+1
end sub

sub dingwallc_hit
	if state=true and tilt=false then addscore 10
	Slingc.visible=0
	Slingc1.visible=1
	me.uservalue=1
	Me.timerenabled=1
end sub

sub dingwallc_timer									'default 50 timer
	select case DingwallC.uservalue
		Case 1: SlingC1.visible=0: SlingC.visible=1
		case 2:	SlingC.visible=0: SlingC2.visible=1
		Case 3: SlingC2.visible=0: SlingC.visible=1: Me.timerenabled=0
	end Select
	DingwallC.uservalue=DingwallC.uservalue+1
end sub

sub dingwallD_hit
	if state=true and tilt=false then addscore 10
	SlingD.visible=0
	SlingD1.visible=1
	DingwallD.uservalue=1
	Me.timerenabled=1
end sub

sub dingwallD_timer									'default 50 timer
	select case DingwallD.uservalue
		Case 1: SlingD1.visible=0: SlingD.visible=1
		case 2:	SlingD.visible=0: SlingD2.visible=1
		Case 3: SlingD2.visible=0: SlingD.visible=1: Me.timerenabled=0
	end Select
	DingwallD.uservalue=DingwallD.uservalue+1
end sub

sub dingwallE_hit
	if state=true and tilt=false then addscore 10
	SlingE.visible=0
	SlingE1.visible=1
	Me.uservalue=1
	Me.timerenabled=1
end sub

sub dingwallE_timer									'default 50 timer
	select case DingwallE.uservalue
		Case 1: SlingE1.visible=0: SlingE.visible=1
		case 2:	SlingE.visible=0: SlingE2.visible=1
		Case 3: SlingE2.visible=0: SlingE.visible=1: Me.timerenabled=0
	end Select
	DingwallE.uservalue=DingwallE.uservalue+1
end sub

sub dingwallF_hit
	if state=true and tilt=false then addscore 10
	SlingF.visible=0
	SlingF1.visible=1
	Me.uservalue=1
	Me.timerenabled=1
end sub

sub dingwallF_timer									'default 50 timer
	select case DingwallF.uservalue
		Case 1: SlingF1.visible=0: SlingF.visible=1
		case 2:	SlingF.visible=0: SlingF2.visible=1
		Case 3: SlingF2.visible=0: SlingF.visible=1: Me.timerenabled=0
	end Select
	DingwallF.uservalue=DingwallF.uservalue+1
end sub

sub dingwallG_hit
	if state=true and tilt=false then addscore 10
	SlingG.visible=0
	SlingI.visible=0
	SlingG1.visible=1
	Me.uservalue=1
	Me.timerenabled=1
end sub

sub dingwallG_timer									'default 50 timer
	select case DingwallG.uservalue
		Case 1: SlingG1.visible=0: SlingG.visible=1
		case 2:	SlingG.visible=0: SlingG2.visible=1
		Case 3: SlingG2.visible=0: SlingG.visible=1: SlingI.visible=1: Me.timerenabled=0
	end Select
	DingwallG.uservalue=DingwallG.uservalue+1
end sub

sub dingwallH_hit
	if state=true and tilt=false then addscore 10
	SlingH.visible=0
	SlingJ.visible=0
	SlingH1.visible=1
	Me.uservalue=1
	Me.timerenabled=1
end sub

sub dingwallH_timer									'default 50 timer
	select case DingwallH.uservalue
		Case 1: SlingH1.visible=0: SlingH.visible=1
		case 2:	SlingH.visible=0: SlingH2.visible=1
		Case 3: SlingH2.visible=0: SlingH.visible=1: SlingJ.visible=1: Me.timerenabled=0
	end Select
	DingwallH.uservalue=DingwallH.uservalue+1
end sub

'******************Scoring

sub addscore(points)
  if tilt=false and state=true then
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
	sreels(player).addvalue(points)
	If B2SOn Then Controller.B2SSetScorePlayer player, score(player) MOD 100000

    ' Sounds: there are 3 sounds: tens, hundreds and thousands
    If Points = 100 AND(Score(player) MOD 1000) \ 100 = 0 Then  'New 1000 reel
        PlaySound SoundFXDOF("bell1000",125,DOFPulse,DOFChimes)
      ElseIf Points = 10 AND(Score(player) MOD 100) \ 10 = 0 Then 'New 100 reel
        PlaySound SoundFXDOF("bell100",124,DOFPulse,DOFChimes)
      ElseIf points = 1000 Then
        PlaySound SoundFXDOF("bell1000",125,DOFPulse,DOFChimes)
	  elseif Points = 100 Then
        PlaySound SoundFXDOF("bell100",124,DOFPulse,DOFChimes)
      Else
        PlaySound SoundFXDOF("bell10",123,DOFPulse,DOFChimes)
		
    End If
	checkreplay
end sub 

sub checkreplay
	if score(player)>99999 then
		If B2SOn then Controller.B2SSetScoreRollover 24 + player, 1
		EVAL("p100k"&player).setvalue 1   'int(score(player)/100000)
	End if
    if score(player)=>replay1 and rep(player)=0 then
		addreplay
		rep(player)=1
    end if
    if score(player)=>replay2 and rep(player)=1 then
		addreplay
		rep(player)=2
    end if
    if score(player)=>replay3 and rep(player)=2 then
		addreplay
		rep(player)=3
    end if
end sub

Sub addreplay
	addcredit
	PlaySound SoundFXDOF("knocker",117,DOFPulse,DOFKnocker)
	DOF 116, DOFPulse
End sub

Sub CheckTilt
	If Tilttimer.Enabled = True Then 
	 TiltSens = TiltSens + 1
	 if TiltSens = 3 Then GameTilted
	Else
	 TiltSens = 0
	 Tilttimer.Enabled = True
	End If
End Sub

Sub GameTilted
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
    bumper1.hashitevent=0
	LeftFlip.RotateToStart
	LeftFlip1.RotateToStart
	TLLFlip.state=0
	DOF 101, DOFOff
	StopSound "Buzz"
	RightFlip.RotateToStart
	RightFlip1.RotateToStart
	TLRFlip.state=0
	DOF 102, DOFOff
	StopSound "Buzz1"
end sub    

Sub addbonus
	if state=true and tilt=false then 
		bonus=bonus+1
		if bonus>19 then bonus=19
		if  bonus = 1 then 
			bonuslight(bonus).state=1
		  else
			bonuslight(bonus).state=1
			bonuslight(bonus-1).state=0
			if bonus>10 then bonuslight(10).state=1
		End if
	end if
End sub

sub addcredit
	if freeplay=0 then 
	  credit=credit+1
	  DOF 121, DOFOn
      if credit>25 then credit=25
	  creditreel.setvalue credit
	  If B2SOn Then Controller.B2ssetCredits Credit
	end if
end sub

sub savehs

    savevalue "SpaceWalk", "credit", credit
    savevalue "SpaceWalk", "hiscore", hisc
    savevalue "SpaceWalk", "match", matchnumb
    savevalue "SpaceWalk", "score1", score(1)
    savevalue "SpaceWalk", "score2", score(2)
	savevalue "SpaceWalk", "replays", replays
	savevalue "SpaceWalk", "balls", balls
	savevalue "SpaceWalk", "freeplay", freeplay
	savevalue "SpaceWalk", "hsa1", HSA1
	savevalue "SpaceWalk", "hsa2", HSA2
	savevalue "SpaceWalk", "hsa3", HSA3
    savevalue "SpaceWalk", "showhisc", showhisc
end sub

sub loadhs
    dim temp
	temp = LoadValue("SpaceWalk", "credit")
    If (temp <> "") then credit = CDbl(temp)
    temp = LoadValue("SpaceWalk", "hiscore")
    If (temp <> "") then hisc = CDbl(temp)
    temp = LoadValue("SpaceWalk", "match")
    If (temp <> "") then matchnumb = CDbl(temp)
    temp = LoadValue("SpaceWalk", "score1")
    If (temp <> "") then score(1) = CDbl(temp)
    temp = LoadValue("SpaceWalk", "score2")
    If (temp <> "") then score(2) = CDbl(temp)
    temp = LoadValue("SpaceWalk", "replays")
    If (temp <> "") then replays = CDbl(temp)
    temp = LoadValue("SpaceWalk", "balls")
    If (temp <> "") then balls = CDbl(temp)
    temp = LoadValue("SpaceWalk", "freeplay")
    If (temp <> "") then freeplay = CDbl(temp)
    temp = LoadValue("SpaceWalk", "hsa1")
    If (temp <> "") then HSA1 = CDbl(temp)
    temp = LoadValue("SpaceWalk", "hsa2")
    If (temp <> "") then HSA2 = CDbl(temp)
    temp = LoadValue("SpaceWalk", "hsa3")
    If (temp <> "") then HSA3 = CDbl(temp)
    temp = LoadValue("SpaceWalk", "showhisc")
    If (temp <> "") then showhisc = CDbl(temp)
end sub

Sub SpaceWalk_Exit()
	turnoff
	Savehs
	If B2SOn Then Controller.stop
End Sub

' *********************************************************************
'                      Supporting Ball & Sound Functions
' *********************************************************************

Function Vol(ball) ' Calculates the Volume of the sound based on the ball speed
    Vol = Csng(BallVel(ball) ^2 / 2000)
End Function

Function Pan(ball) ' Calculates the pan for a ball based on the X position on the table. "table1" is the name of the table
    Dim tmp
    tmp = ball.x * 2 / SpaceWalk.width-1
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
        If BallVel(BOT(b) ) > 1 AND BOT(b).z < 30 Then
            rolling(b) = True
            PlaySound("fx_ballrolling" & b), -1, Vol(BOT(b) ), Pan(BOT(b) ), 0, Pitch(BOT(b) ), 1, 0
        Else
            If rolling(b) = True Then
                StopSound("fx_ballrolling" & b)
                rolling(b) = False
            End If
        End If
    Next
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

Sub Pins_Hit (idx)
	PlaySound "pinhit_low", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0
End Sub

Sub Targets_Hit (idx)
	PlaySound "target", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0
End Sub

Sub Metals_Thin_Hit (idx)
	PlaySound "metalhit_thin", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0
End Sub

Sub Metals_Medium_Hit (idx)
	PlaySound "metalhit_medium", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0
End Sub

Sub Metals2_Hit (idx)
	PlaySound "metalhit2", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0
End Sub

Sub Gates_Hit (idx)
	PlaySound "gate4", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0
End Sub

Sub Spinner_Spin
	PlaySound "fx_spinner",0,.25,0,0.25
End Sub

Sub Rubbers_Hit(idx)
 	dim finalspeed
  	finalspeed=SQR(activeball.velx * activeball.velx + activeball.vely * activeball.vely)
 	If finalspeed > 20 then 
		PlaySound "fx_rubber2", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0
	End if
	If finalspeed >= 6 AND finalspeed <= 20 then
 		RandomSoundRubber()
 	End If
End Sub

Sub Posts_Hit(idx)
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


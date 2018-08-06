'	-------------------------------------------------
'	DOUBLE-UP   Bally, 1970
'	-------------------------------------------------
'	
'	VPX - 2016
'	Playfield and scripting by BorgDog
'	Backglass, plastics and other misc parts by HauntFreaks
'	Initial inspiration, images, rules and gameplay provided by Pecos on pinside - THANKS!
'
'	my DOF scripting fixed by arngrim :)
'	
'	-------------------------------------------------

Option Explicit
Randomize

Const cGameName = "doubleup_1970"
Const Ballsize = 50			'used by ball shadow routine

Dim operatormenu, options
Dim bumperlitscore
Dim bumperoffscore
Dim balls
Dim replays, liberal, liblight
Dim Replay1Table(3), Replay2Table(3), Replay3Table(3), Replay4Table(3)
dim replay1, replay2, replay3, replay4
Dim Add10, Add100, Add1000
Dim hisc
Dim maxplayers, players, player
Dim credit
Dim holescore, bonus
Dim score(4)
Dim sreels(4)
Dim p100k(4)
Dim cplay(4)
Dim Pups(4)
Dim state
Dim tilt, tiltsens
Dim ballinplay
Dim matchnumb
dim rstep, lstep, mushL, mushR
Dim rep(4)
Dim rst
Dim eg
Dim bell
Dim i,j, ii, objekt, light
Dim awardcheck

On Error Resume Next
ExecuteGlobal GetTextFile("controller.vbs")
If Err Then MsgBox "Can't open controller.vbs"
On Error Goto 0

'**********************************************************
'********   	OPTIONS		*******************************
'**********************************************************

Dim BallShadows: Ballshadows=1  		'******************set to 1 to turn on Ball shadows
Dim FlipperShadows: FlipperShadows=1  	'***********set to 1 to turn on Flipper shadows
Dim freeplay:  freeplay=0				'********set to 1 for freeplay, 0 for credit play

sub DoubleUp_init
	LoadEM
	maxplayers=1
	Replay1Table(1)=20000
	Replay2Table(1)=25000
	Replay3Table(1)=28000
	Replay4Table(1)=31000

	Replay1Table(2)=25000
	Replay2Table(2)=30000
	Replay3Table(2)=33000
	Replay4Table(2)=35000

	Replay1Table(3)=40000
	Replay2Table(3)=49000
	Replay3Table(3)=58000
	Replay4Table(3)=66000
	set sreels(1) = ScoreReel1

	set Pups(1) = Pup1

	hisc=20000		'SET DEFAULT VALUES PRIOR TO LOADING FROM SAVED FILE
	credit=0
	freeplay=0
	balls=5
	matchnumb=0
	liberal=0
	replays=2
	HSA1=4
	HSA2=15
	HSA3=7

	hideoptions
	player=1
	For each light in BumperLights:light.State = 0: Next
	For each light in bonuslights:light.State = 0: Next
	loadhs
	UpdatePostIt	'UPDATE HIGH SCORE STICKY

	if freeplay=0 then credittxt.setvalue(credit)

	Replay1=Replay1Table(Replays)
	Replay2=Replay2Table(Replays)
	Replay3=Replay3Table(Replays)
	Replay4=Replay4Table(Replays)
	RepCard.image = "ReplayCard"&replays
	OptionBalls.image="OptionsBalls"&Balls
	OptionReplays.image="OptionsReplays"&replays
	OptionLiberal.image="OptionsLiberal"&liberal
	bumperlitscore=100
	bumperoffscore=10
	if balls=3 then	
		InstCard.image="InstCard3balls"
	  else
		InstCard.image="InstCard5balls"
	end if
	setBackglass.enabled=true
	If B2SOn then
		for each objekt in backdropstuff : objekt.visible = 0 : next
	End If
	if matchnumb="" then matchnumb=0
	if matchnumb=0 then
		matchtxt.text="00"
	  else
		matchtxt.text=matchnumb*10
	end if
	for i = 1 to maxplayers
		sreels(i).setvalue(score(i))
	next
	tilt=false
	liblight=0

	if ballshadows=1 then
		BallShadowUpdate.enabled=1
	  else
		BallShadowUpdate.enabled=0
	end if

	if flippershadows=1 then 
		FlipperLSh.visible=1
		FlipperRSh.visible=1
	  else
		FlipperLSh.visible=0
		FlipperRSh.visible=0
	end if

	Drain.CreateSizedBallWithMass 25,1.3 	'CreateBall

End sub

sub setBackglass_timer
	playsound "turnon"
	IF Credit > 0 Then DOF 132, DOFOn
	gamov.text="Game Over"
	gamov.timerenabled=1
	tilttxt.timerenabled=1
	for each light in GIlights:light.state=1: next
	if b2son then
		Controller.B2ssetCredits Credit
		if matchnumb=0 then 
			Controller.B2ssetMatch 100
		  Else
			Controller.B2ssetMatch Matchnumb*10
		end if
		Controller.B2SSetGameOver 1
		Controller.B2SSetScorePlayer 1, Score(1) MOD 100000
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
	  Else
		If B2SOn then Controller.B2SSetGameOver 35,0
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
	  Else
		If B2SOn then Controller.B2SSetTilt 33,0
	end if
	me.enabled=0
end sub

Sub DoubleUp_KeyDown(ByVal keycode)
   
	if keycode=AddCreditKey then
		if credit<14 then 
			playsound "coinin" 
		  Else
			playsound "coin3"
		end if
		coindelay.enabled=true 
    end if

    if keycode=StartGameKey and (credit>0 or freeplay=1) And Not HSEnterMode=true and NOT setBackglass.enabled then
	  if state=false then
		state=true
		tilttxt.timerenabled=0
		ttimer.enabled=0
		gamov.timerenabled=0
		gtimer.enabled=0
		if freeplay=0 then
			credit=credit-1
			If credit < 1 Then DOF 132, DOFOff
			playsound "cluper"
			credittxt.setvalue(credit)
		end if
		ballinplay=1
		If B2SOn Then 
			Controller.B2ssetCredits Credit
			Controller.B2sStartAnimation "Startup"
			Controller.B2ssetballinplay 32, Ballinplay
			Controller.B2ssetplayerup 30, 1
			Controller.B2SSetGameOver 0
		End If
	    pups(1).state=1
		tilt=false
		playsound "initialize" 
		players=1

		rst=0
		newgame.enabled=true
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
			if Liberal=1 Then
				Liberal=0
			  Else
				Liberal=1
			end If
			OptionLiberal.image= "OptionsLiberal"&Liberal
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
	If keycode = MechanicalTilt Then
		CheckTilt
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
	OptionLiberal.visible = True
End Sub

Sub HideOptions
	for each objekt In OptionMenu
		objekt.visible = false
	next
End Sub


Sub DoubleUp_KeyUp(ByVal keycode)

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

Sub PairedlampTimer_timer
testbox.text=matchnumb
	bumperlighta1.state = bumperlight1.state
	bumperlighta2.state = bumperlight2.state
	bumperlight3.state = bumperlight2.state
	bumperlighta3.state = bumperlight2.state
	bumperlight4.state = BumperLight1.state
	bumperlighta4.state = BumperLight1.state
	lhole1r.state=Lhole1.state
	lhole2r.state=lhole2.state
	lhole3r.state=lhole3.state
	Lhole4r.state=lhole4.state
	lspecial1.state=Lspecial.state
	Pgate.rotz=gate.currentangle+25
	Pgate1.rotz=gate1.currentangle+25
	pup2.state=pup1.state
	Pup3.state=pup1.state
	Pup4.state=pup1.state
	Pup5.state=pup1.state

	if FlipperShadows=1 then
		FlipperLSh.RotZ = LeftFlipper.currentangle
		FlipperRSh.RotZ = RightFlipper.currentangle
	end if

end sub

sub coindelay_timer
	addcredit
    coindelay.enabled=false
end sub

Sub addcredit
	if freeplay=0 then
      credit=credit+2
	  DOF 132, DOFOn
      if credit>10 then 
		credit=9
		credittxt.setvalue(25)
	   Else
		credittxt.addvalue(2)
	  end if
	  If B2SOn Then Controller.B2ssetCredits Credit
	end if
End sub

Sub Drain_Hit()
	PlaySound "drain",0,1,0,0.25
	DOF 127, DOFPulse
	me.timerenabled=1
End Sub

Sub Drain_timer
	  if players=1 or player=players then 
		player=1
		pups(players).state=0
		pups(player).state=1
	   Else
		player=player+1
		pups(player).state=1
		pups(player-1).state=0
	  end if
	  If B2SOn then Controller.B2ssetplayerup 30, player
	  nextball
	me.timerenabled=0
End Sub	

sub ballhome_unhit
	DOF 133, DOFPulse
end sub

sub newgame_timer
	bumper1.force=10
	bumper2.force=10
	bumper3.force=10
	Bumper4.force=10
	player=1
	for i = 1 to maxplayers
	    score(i)=0
		rep(i)=0
		pups(i).state=0
		sreels(i).resettozero
	next
	holescore=1
	bonus=0
	Lhole1.state=1
    pups(1).state=1
	if liberal=1 Then
		LmushL.state=1
		LmushR.state=1
	  Else
		LmushL.state=1
	end If
	bumperlight1.state=0
	bumperlight2.state=0
	loutl.state=0
	loutr.state=0
	If B2SOn then
	  for i = 1 to maxplayers
		Controller.B2SSetScorePlayer i, score(i)
	  next 
	End If
    eg=0
	for each light in GIlights:light.state=1: next
	for each light in bumperlights:light.state=0:next
	for each light in bonuslights: light.state=0: next
	for i=2 to 4: EVAL("Lhole"&i).state=0: next
    tilttxt.text=" "
	gamov.text=" "
    If B2SOn then
		Controller.B2SSetGameOver 0
		Controller.B2SSetTilt 0
		Controller.B2SSetMatch 0
	End If
    biptext.text="1"
	matchtxt.text=" "
	newball
	ballreltimer.enabled=true
	newgame.enabled=false
end sub


sub newball
	for each light in GIlights:light.state=1:next
End Sub


sub nextball
    if tilt=true then
	  bumper1.hashitevent=true
	  bumper2.hashitevent=true
	  bumper3.hashitevent=true
	  Bumper4.hashitevent=true
	  LeftSlingShot.isdropped=False
	  RightSlingShot.isdropped=False
      tilt=false
      tilttxt.text=" "
		If B2SOn then
			Controller.B2SSetTilt 33,0
'			Controller.B2ssetdata 1, 1
		End If
    end if
	if player=1 then ballinplay=ballinplay+1
	if ballinplay>balls then
		playsound "gameover"
		eg=1
		ballreltimer.enabled=true
	  else
		if state=true then
		  newball
		  ballreltimer.enabled=true
		end if
		biptext.text=ballinplay
		If B2SOn then Controller.B2ssetballinplay 32, Ballinplay
	end if
End Sub

sub ballreltimer_timer
  if eg=1 then
      matchnum
	  state=false
	  biptext.text=" "
	  gamov.text="GAME OVER"
	  gamov.timerenabled=1
	  tilttxt.timerenabled=1
	  for i=1 to maxplayers
		if score(i)>hisc then 
			hisc=score(i)
			HStimer.uservalue = 0
			HStimer.enabled=1
			HighScoreEntryInit()
		end if
		pups(i).state=0
	  next

	  UpdatePostIt
	  savehs
	  If B2SOn then 
        Controller.B2SSetGameOver 35,1
        Controller.B2ssetballinplay 32, 0
	    Controller.B2sStartAnimation "EOGame"
	  End If
	  ballreltimer.enabled=false
  else
	  plunger.timerenabled=true
      ballreltimer.enabled=false
  end if
end sub

Sub HStimer_timer
	PlaySoundAtVol SoundFXDOF("knock",128,DOFPulse,DOFKnocker), Plunger, 1
	DOF 129,DOFPulse
	HStimer.uservalue=HStimer.uservalue+1
	if HStimer.uservalue=3 then me.enabled=0
end sub

Sub plunger_timer
 	Drain.kick 60, 16,0
    PlaySound SoundFXDOF("kickerkick",126,DOFPulse,DOFContactors)
	Plunger.timerenabled=false
end sub

sub matchnum
	if matchnumb=0 then
		matchtxt.text="00"
		If B2SOn then Controller.B2ssetMatch 100
	  else
		matchtxt.text=matchnumb*10
		If B2SOn then Controller.B2ssetMatch Matchnumb*10
	end if
	For i=1 to players
		if (matchnumb*10)=(score(i) mod 100) then 
			addcredit
			PlaySoundAtVol SoundFXDOF("knock",128,DOFPulse,DOFKnocker), Plunger, 1
			DOF 129,DOFPulse
	    end if
    next
end sub

'********** Bumpers

Sub Bumper1_Hit
   if tilt=false then
	playsound SoundFXDOF("fx_bumper4",107,DOFPulse,DOFContactors)
	DOF 108,DOFPulse
	if bumperlight1.state=1 then
		addscore bumperlitscore
	  else
		addscore bumperoffscore
	end if
   end if
End Sub



Sub Bumper2_Hit
   if tilt=false then
	playsound SoundFXDOF("fx_bumper4",109,DOFPulse,DOFContactors)
	DOF 110,DOFPulse
	if BumperLight2.state=1 then
		addscore bumperlitscore
	  else
		addscore bumperoffscore
	end if	
   end if
End Sub

Sub Bumper3_Hit
   if tilt=false then
	playsound SoundFXDOF("fx_bumper4",111,DOFPulse,DOFContactors)
	DOF 112,DOFPulse
	if BumperLight3.state=1 then
		addscore bumperlitscore
	  else
		addscore bumperoffscore
	end if	
   end if
End Sub



Sub Bumper4_Hit
   if tilt=false then
	playsound SoundFXDOF("fx_bumper4",113,DOFPulse,DOFContactors)
	DOF 114,DOFPulse
	if BumperLight4.state=1 then
		addscore bumperlitscore
	  else
		addscore bumperoffscore
	end if	
   end if
End Sub



'************** Slings and animated rubbers


Sub RightSlingShot_Slingshot
    PlaySound SoundFXDOF("right_slingshot",105,DOFPulse,DOFContactors), 0, 1, 0.05, 0.05
	DOF 106, DOFPulse 
	addscore 10
    RSling.Visible = 0
    rsling2.Visible = 1
    sling1.rotx = 20
    RStep = 1
    RightSlingShot.TimerEnabled = 1
End Sub

Sub RightSlingShot_Timer
    Select Case RStep
        Case 3:rsling2.Visible = 0:RSLing1.Visible = 1:sling1.rotx = 10
        Case 4:sling1.rotx = 0:RSLing1.Visible = 0:RSLing.Visible = 1:RightSlingShot.TimerEnabled = 0
    End Select
    RStep = RStep + 1
End Sub

Sub LeftSlingShot_Slingshot
    PlaySound SoundFXDOF("left_slingshot",103,DOFPulse,DOFContactors),0,1,-0.05,0.05
	DOF 104, DOFPulse 
	addscore 10
    LSling.Visible = 0
    LSling2.Visible = 1
    sling2.rotx = 20
    LStep = 1
    LeftSlingShot.TimerEnabled = 1
End Sub

Sub LeftSlingShot_Timer
    Select Case LStep
        Case 3:LSLing2.Visible = 0:LSLing1.Visible = 1:sling2.rotx = 10
        Case 4:sling2.rotx = 0:LSLing1.Visible = 0:LSLing.Visible = 1:LeftSlingShot.TimerEnabled = 0
    End Select
    LStep = LStep + 1
End Sub

Sub Dingwalls_Hit(idx)
	addscore 10
End Sub

sub dingwalla_hit
	if state=true and tilt=false then addscore 10
	SlingA.visible=0
	SlingA1.visible=1
	me.uservalue=1
	Me.timerenabled=1
end sub

sub dingwalla_timer									
	select case me.uservalue
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

sub dingwallb_timer									
	select case me.uservalue
		Case 1: Slingb1.visible=0: SlingB.visible=1
		case 2:	SlingB.visible=0: Slingb2.visible=1
		Case 3: Slingb2.visible=0: SlingB.visible=1: Me.timerenabled=0
	end Select
	me.uservalue=me.uservalue+1
end sub


sub dingwallc_hit
	if state=true and tilt=false then addscore 10
	Slingc.visible=0
	Slingc1.visible=1
	me.uservalue=1
	Me.timerenabled=1
end sub

sub dingwallc_timer									
	select case me.uservalue
		Case 1: Slingc1.visible=0: Slingc.visible=1
		case 2:	Slingc.visible=0: Slingc2.visible=1
		Case 3: Slingc2.visible=0: Slingc.visible=1: Me.timerenabled=0
	end Select
	me.uservalue=me.uservalue+1
end sub

sub dingwalld_hit
	if state=true and tilt=false then addscore 10
	Slingd.visible=0
	Slingd1.visible=1
	me.uservalue=1
	Me.timerenabled=1
end sub

sub dingwalld_timer									
	select case me.uservalue
		Case 1: Slingd1.visible=0: Slingd.visible=1
		case 2:	Slingd.visible=0: Slingd2.visible=1
		Case 3: Slingd2.visible=0: Slingd.visible=1: Me.timerenabled=0
	end Select
	me.uservalue=me.uservalue+1
end sub

'********** Triggers     

sub TGoutL_hit
	addscore 100
	DOF 123,DOFPulse
    if (LoutL.state)=lightstateon then TGoutL.timerenabled=1
end sub    

sub TGoutR_hit
	addscore 100
	DOF 124,DOFPulse
    if (LoutR.state)=lightstateon then TGoutL.timerenabled=1
end sub

sub TGoutL_timer
	select case bonus
		Case 1: addscore 500
		Case 2: addscore 1000
		Case 3: addscore 2000
		Case 4: addscore 4000
		Case 5: addspecial
	end Select
	addhole
	TGoutL.timerenabled=0
end sub

sub TGtop1_hit   '***** top L 1000 rollover
	DOF 119,DOFPulse
	addscore 1000
	if (LtopLD.state)=1 then addhole
end sub    

sub TGtop2_hit   '***** top L 100 rollover
	DOF 120,DOFPulse
	if (LtopL1k.state)=1 then
		addscore 1000
      else
		addscore 100
    end if
end sub

sub TGtop3_hit   '***** top R 100 rollover
	DOF 121,DOFPulse
	if (LtopR1k.state)=1 then
		addscore 1000
      else
		addscore 100
    end if
end sub

sub TGtop4_hit   '***** top R 1000 rollover
	DOF 122,DOFPulse
	addscore 1000
	if (LtopRD.state)=1 then addhole
end sub

sub TmushL_hit		'***** left mushroom bumper
	DOF 115,DOFPulse
	PmushL.transz=7
	mushL=0
	me.timerenabled=1
	if LmushL.state=1 Then	
		addscore 100
		addhole
	  Else
		addscore 10
	end if	
end Sub

sub TmushL_timer
	mushL=mushL+1
	Select case mushL
		case 1: PmushL.transz=12
		case 2: PmushL.transz=7
		case 3: 
			PmushL.transz=0
			me.timerenabled=0
	end Select
end sub
	

sub TmushR_hit		'***** right mushroom bumper
	DOF 116,DOFPulse
	PmushR.transz=7
	mushR=0
	me.timerenabled=1
	if LmushR.state=1 Then	
		addscore 100
		addhole
	  Else
		addscore 10
	end if	
end Sub

sub TmushR_timer
	mushR=mushR+1
	Select case mushR
		case 1: PmushR.transz=12
		case 2: PmushR.transz=7
		case 3: 
			PmushR.transz=0
			me.timerenabled=0
	end Select
end sub

sub TbuttonL_hit		'****** top left rollover button
	DOF 117,DOFPulse
	if holescore <>4 then LtopLD.state=1
	LtopL1k.state=1
end Sub


sub TbuttonR_hit		'****** top right rollover button
	DOF 118,DOFPulse
	if holescore <>4 then LtopRD.state=1
	LtopR1k.state=1
end Sub

sub target_hit
	PlaySound SoundFXDOF("target",125,DOFPulse,DOFContactors),0,.5
	addscore 1000
	if Lspecial1.state=1 then addspecial
end Sub

'******************  Kickers

sub KickerL_hit
	collecthole
	me.timerenabled=1
end Sub

sub KickerL_timer
	addhole
	playsound SoundFXDOF("holekick",131,DOFPulse,DOFContactors)
	DOF 129,DOFPulse
	KickerL.kick 155,15
	KickerL.timerenabled=0
end Sub

sub KickerR_hit
	collecthole
	me.timerenabled=1
end Sub

sub kickerR_timer
	addhole
	playsound SoundFXDOF("holekick",130,DOFPulse,DOFContactors)
	DOF 129,DOFPulse
	KickerR.kick 190,15
	KickerR.timerenabled=0
end Sub

sub collecthole
	select case holescore
		Case 1: addscore 500
		Case 2: addscore 1000
		Case 3: addscore 2000
		Case 4: addscore 4000
	end Select
end sub

sub addspecial
	playsound SoundFXDOF("knock",128,DOFPulse,DOFKnocker)
	DOF 129,DOFPulse
	addcredit
end sub

sub addhole
	holescore=holescore+1
	if holescore>4 then 
		addbonus
		If Lspecial.state=0 then 
			if liberal=1 Then
				LmushL.state=1
				LmushR.state=1
			  Else
				if liblight=1 Then
					LmushL.state=0
					LmushR.state=1
				  Else
					LmushL.state=1
					LmushR.state=0
				end If
			end If
			holescore=1
		  Else
			holescore=4
		End If
	end If
	if holescore=4 Then for each light in addholelights:light.state=0:next
	for i=1 to 4
		EVAL("Lhole"&i).state=0
	Next
	EVAL("Lhole"&holescore).state=1
end Sub

sub addbonus
	bonus=bonus+1
	if bonus=1 Then
		If liberal=1 then 
			LoutL.state=1
			LoutR.state=1
		  Else
			if liblight=1 Then
				LoutL.state=0
				LoutR.state=1
			  Else
				LoutL.state=1
				LoutR.state=0
			end If
		end If
		BumperLight1.state=1
	end If
	if bonus=2 Then	BumperLight2.state=1
	if bonus>4 then 
		Lspecial.state=1
		bonus=5
	end If
	for i=1 to 4
		EVAL("Lbonus"&i).state=0
	Next
	if bonus<5 then EVAL("Lbonus"&bonus).state=1
end Sub

sub addscore(points)
  if tilt=false then
	if points=10 then 
		matchnumb=matchnumb+1
		if matchnumb>9 then matchnumb=0
		if liblight=0 Then
			liblight=1
		  Else
			liblight=0
		end If
		if liberal=0 and bonus>0 Then
			if liblight=1 then
					LoutL.state=0
					LoutR.state=1
			  Else
					LoutL.state=1
					LoutR.state=0
			end If
		end if
		if liberal=0 and holescore<>4 Then
			if liblight=1 Then
				LmushL.state=1
				LmushR.state=0
			  Else
				LmushL.state=0
				LmushR.state=1
			end If
		end if
		for each light in inlanelights: light.state=0: next
	end If
	if points=10 or points=100 or points=1000 then 
		addpoints Points
	  Else
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
  End If
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
	If B2SOn Then Controller.B2SSetScorePlayer player, score(player)

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
		P100k(player).text="100,000"
	End if
    if score(player)=>replay1 and rep(player)=0 then
		addcredit
		rep(player)=1
		PlaySound SoundFXDOF("knock",128,DOFPulse,DOFKnocker)
		DOF 129, DOFPulse
    end if
    if score(player)=>replay2 and rep(player)=1 then
		addcredit
		rep(player)=2
		PlaySound SoundFXDOF("knock",128,DOFPulse,DOFKnocker)
		DOF 129, DOFPulse
    end if
    if score(player)=>replay3 and rep(player)=2 then
		addcredit
		rep(player)=3
		PlaySound SoundFXDOF("knock",128,DOFPulse,DOFKnocker)
		DOF 129, DOFPulse
    end if
    if score(player)=>replay4 and rep(player)=3 then
		addcredit
		rep(player)=4
		PlaySound SoundFXDOF("knock",128,DOFPulse,DOFKnocker)
		DOF 129, DOFPulse
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
    bumper1.hashitevent=0
    bumper2.hashitevent=0
	bumper3.hashitevent=0
	Bumper4.hashitevent=0
	LeftSlingShot.isdropped=true
	RightSlingShot.isdropped=true
  	LeftFlipper.RotateToStart
	StopSound "Buzz"
	DOF 101, DOFOff
	RightFlipper.RotateToStart
	StopSound "Buzz1"
	DOF 102, DOFOff
end sub    

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

Function AudioFade(tableobj) ' Fades between front and back of the table (for surround systems or 2x2 speakers, etc), depending on the Y position on the table. "table1" is the name of the table
	Dim tmp
    tmp = tableobj.y * 2 / DoubleUp.height-1
    If tmp > 0 Then
		AudioFade = Csng(tmp ^10)
    Else
        AudioFade = Csng(-((- tmp) ^10) )
    End If
End Function

Function AudioPan(tableobj) ' Calculates the pan for a tableobj based on the X position on the table. "table1" is the name of the table
    Dim tmp
    tmp = tableobj.x * 2 / DoubleUp.width-1
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
	PlaySound("fx_collide"), 0, Csng(velocity) ^2 / 2000, AudioPan(ball1), 0, Pitch(ball1), 0, 0, AudioFade(ball1)
End Sub


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
        If BOT(b).X < DoubleUp.Width/2 Then
            BallShadow(b).X = ((BOT(b).X) - (Ballsize/16) + ((BOT(b).X - (DoubleUp.Width/2))/17))' + 13
        Else
            BallShadow(b).X = ((BOT(b).X) + (Ballsize/16) + ((BOT(b).X - (DoubleUp.Width/2))/17))' - 13
        End If
        ballShadow(b).Y = BOT(b).Y + 10
        If BOT(b).Z > 20 Then
            BallShadow(b).visible = 1
        Else
            BallShadow(b).visible = 0
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


'**************************************
' Explanation of the collision routine
'**************************************

' The collision is built in VP.
' You only need to add a Sub OnBallBallCollision(ball1, ball2, velocity) and when two balls collide they 
' will call this routine. What you add in the sub is up to you. As an example is a simple Playsound with volume and paning
' depending of the speed of the collision.


Sub a_Triggers_Hit (idx)
	playsound "sensor", 0,1,AudioPan(ActiveBall),0,0,0,1,AudioFade(ActiveBall)
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
	PlaySound "metalhit_medium", 0, Vol(ActiveBall), AudioPan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
End Sub

Sub a_Metals2_Hit (idx)
	PlaySound "metalhit2", 0, Vol(ActiveBall), AudioPan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
End Sub

Sub a_Gates_Hit (idx)
	PlaySound "gate4", 0, Vol(ActiveBall), AudioPan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
End Sub

Sub Spinner_Spin
	PlaySound "fx_spinner", 0, .25, AudioPan(Spinner), 0.25, 0, 0, 1, AudioFade(Spinner)
End Sub

Sub a_Rubbers_Hit(idx)
 	dim finalspeed
  	finalspeed=SQR(activeball.velx * activeball.velx + activeball.vely * activeball.vely)
 	If finalspeed > 20 then 
		PlaySound "fx_rubber2", 0, Vol(ActiveBall), AudioPan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
	End if
	If finalspeed >= 6 AND finalspeed <= 20 then
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
		Case 1 : PlaySound "rubber_hit_1", 0, Vol(ActiveBall), AudioPan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
		Case 2 : PlaySound "rubber_hit_2", 0, Vol(ActiveBall), AudioPan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
		Case 3 : PlaySound "rubber_hit_3", 0, Vol(ActiveBall), AudioPan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
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

    savevalue "DoubleUp", "credit", credit
    savevalue "DoubleUp", "hiscore", hisc
    savevalue "DoubleUp", "match", matchnumb
    savevalue "DoubleUp", "replays", replays
    savevalue "DoubleUp", "score1", score(1)
	savevalue "DoubleUp", "balls", balls
	savevalue "DoubleUp", "liberal", liberal

end sub

sub loadhs
    dim temp
	temp = LoadValue("DoubleUp", "credit")
    If (temp <> "") then credit = CDbl(temp)
    temp = LoadValue("DoubleUp", "hiscore")
    If (temp <> "") then hisc = CDbl(temp)
    temp = LoadValue("DoubleUp", "match")
    If (temp <> "") then matchnumb = CDbl(temp)
    temp = LoadValue("DoubleUp", "replays")
    If (temp <> "") then replays = CDbl(temp)
    temp = LoadValue("DoubleUp", "score1")
    If (temp <> "") then score(1) = CDbl(temp)
    temp = LoadValue("DoubleUp", "balls")
    If (temp <> "") then balls = CDbl(temp)
    temp = LoadValue("DoubleUp", "liberal")
    If (temp <> "") then liberal = CDbl(temp)
end sub

Sub DoubleUp_Exit()
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


'///////////////////////////////////////////////////////////////////////
'			TKO by Gottlieb (1979)
'
'
'
' vp10 assembled and scripted by BorgDog 2015, update 2018
' 
' borrowed some play logic from Itchigo's and pbecker1946's vp9 TKO wip. Thanks guys.
'
' Controller vbs implementation and DOF by Arngrim
'
' menu system idea borrowed from loserman76 and gnance, hold left flipper to bring up game options menu
' 
'////////////////////////////////////////////////////////////////////////

Option Explicit
Randomize

On Error Resume Next
ExecuteGlobal GetTextFile("controller.vbs")
If Err Then MsgBox "Can't open controller.vbs"
On Error Goto 0

Const cGameName = "tko_1979"
Const Ballsize = 50			'used by ball shadow routine

Dim operatormenu, options, optionsX
Dim bumperlitscore
Dim bumperoffscore
Dim balls
dim ebcount
Dim replays
Dim Replay1Table(2)
Dim Replay2Table(2)
dim replay1, replay2, replay3
Dim Add10, Add100, Add1000
Dim hisc, hiscstate
Dim Controller
Dim credit, freeplay, ballshadows, flippershadows
Dim score(2)
Dim sreels(2)
Dim player, Players, maxplayers
Dim state
Dim tilt, tiltsens
Dim Advance
Dim AdvanceL(16)
Dim Whitecount
Dim ballinplay
Dim matchnumb
dim rlight
Dim rep(2)
Dim rst
Dim eg
Dim scn
Dim scn1
Dim bells
Dim i,j, objekt, light
Dim awardcheck
Dim lstep


sub TKO_init
	LoadEM
	maxplayers=1
	Replay1Table(1)=80000
	Replay1Table(2)=90000
	Replay2Table(1)=110000
	Replay2Table(2)=140000
	set sreels(1) = scoreReel1
	set AdvanceL(1) = advance1
	set AdvanceL(2) = advance2
	set AdvanceL(3) = advance3
	set AdvanceL(4) = advance4
	set AdvanceL(5) = advance5
	set AdvanceL(6) = advance6
	set AdvanceL(7) = advance7
	set AdvanceL(8) = advance8
	set AdvanceL(9) = advance9
	set AdvanceL(10) = advance10
	set AdvanceL(11) = advance11
	set AdvanceL(12) = advance12
	set AdvanceL(13) = advance13
	set AdvanceL(14) = advance14
	set AdvanceL(15) = advance15
	set AdvanceL(16) = advance16
	hideoptions
	player=1
	turnoff

	hisc=60000		'SET DEFAULT VALUES PRIOR TO LOADING FROM SAVED FILE
	credit=0
	freeplay=0
	flippershadows=1
	ballshadows=0
	balls=5
	matchnumb=0
	replays=2
	HSA1=4
	HSA2=15
	HSA3=7

	if ShowDT then
		OptionReel.objrotx=10
		OptionsReel.objrotx=0
		OptionBox.objrotx=-14
		OptionBox.z=25
		OptionReel.z=30
		OptionsReel.z=30
	  else
		OptionReel.objrotx=28
		OptionsReel.objrotx=15
		OptionBox.objrotx=0
		OptionBox.z=0
		OptionReel.z=5
		OptionsReel.z=5
	end If

	loadhs			'LOADS SAVED OPTIONS AND HIGH SCORE IF EXISTING
	UpdatePostIt	'UPDATE HIGH SCORE STICKY


	bumperlitscore=1000
	bumperoffscore=100

	Replay1=Replay1Table(Replays)
	Replay2=Replay2Table(Replays)
	RepCard.image = "ReplayCard"&replays
	if balls=3 then
		InstCard.image="instcard3balls"
	end if
	if balls=5 then
		InstCard.image="instcard5balls"
	end if
	If B2SOn then
		for each objekt in backdropstuff : objekt.visible = 0 : next
	End If
	startGame.enabled=true

	for i = 1 to maxplayers
		sreels(i).setvalue(score(i))
		if score(i)>99999 then
			Reel100K1.setvalue(int(score(1)/100000))
		end if
	next

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

	tilt=false
	If credit > 0 or freeplay=1 Then DOF 123, DOFOn

	Drain.CreateBall
End sub

sub startGame_timer
	PlaySoundAt "poweron", Plunger
	setBackglass.enabled=true
	me.enabled=false
end sub

sub setBackglass_timer
	gamov.text="GAME OVER"
	gamov.timerenabled=1
	tilttxt.timerenabled=1
	creditreel.setvalue(credit)
	if matchnumb=0 then
		matchtxt.text="00"
	  else
		matchtxt.text=matchnumb*10
	end if

	For each light in BumperLights:light.State = 1: Next
	For each light in GI:light.State = 1: Next
	For each light in AdvanceLights:light.State = 2: Next
	Spot1.state=2:Spot2.state=1:Spot3.state=2:Spot4.state=1

	if b2son then 
		Controller.B2ssetCredits Credit
		Controller.B2SSetData 6,1
		Controller.B2ssetMatch 34, Matchnumb
		Controller.B2SSetGameOver 35,1
		Controller.B2SSetScorePlayer 5, hisc
		Controller.B2SSetScorePlayer 1, Score(1) MOD 100000
		if score(1)>999999 then Controller.B2SSetScoreRollover 25, 1
	end if
	me.enabled=false
end sub

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
		gamov.text="GAME OVER"
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

Sub TKO_KeyDown(ByVal keycode)
   
	if keycode=AddCreditKey then
		if freeplay=1 or credit>15 Then
			playsoundat "coin3", Drain
		  Else
			playsoundat "coinIn6", Drain
			addcredit
		end If

    end if

    if keycode=StartGameKey and (Credit>0 or freeplay=1) And Not HSEnterMode=true and Not Startgame.enabled then
	  if state=false then
		if freeplay=0 then 
			credit=credit-1
			If credit < 1 Then DOF 123, DOFOff
			creditreel.setvalue(credit)
		end if
		playsound "cluper"
		creditreel.setvalue(credit)
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
		tilt=false
		state=true
		playsound "initialize" 
		players=1
		rst=0
		resettimer.enabled=true
	  else if state=true and players < maxplayers and Ballinplay=1 and (Credit>0 or freeplay=1) then
		credit=credit-1
		If credit < 1 Then DOF 123, DOFOff
		players=players+1
		creditreel.setvalue(credit)
		If B2SOn then
			Controller.B2ssetCredits Credit
			Controller.B2ssetcanplay 31, players
		End If
		CPlay(players).Text=players
		playsound "cluper" 
	   end if 
	  end if
	end if

	If keycode = PlungerKey Then
		Plunger.PullBack
		PlaySoundat "plungerpull", plunger
	End If

	If keycode=LeftFlipperKey and State = false and OperatorMenu=0 then
		OperatorMenuTimer.Enabled = true
	end if

	If keycode=LeftFlipperKey and State = false and OperatorMenu=1 and Not OptionReelTimer.enabled and Not OptionsReelTimer.enabled then
		Options=Options+1
		If Options=7 then Options=1
		PlaySoundAt SoundFXDOF("Solenoid",104,DOFPulse,DOFContactors), OptionReel
		OptionReelTimer.enabled=true
		Select Case (Options)
			Case 1:
				if Balls=3 then
					OptionsX=1
				  else
					OptionsX=2
				end if
			Case 2:
				if ballshadows=0 Then
					OptionsX=4    	'NO
				  Else
					OptionsX=3		'yes
				end If
			Case 3:
				if freeplay=0 Then
					OptionsX=4    	'no
				  Else
					OptionsX=3		'yes
				end if
			Case 4:
				Select Case (replays)
					Case 1:
						OptionsX=5		'Low
					Case 2:
						OptionsX=7		'High
				End Select
			Case 5:
				if flippershadows=0 Then
					OptionsX=4    	'NO
				  Else
					OptionsX=3		'yes
				end If
			Case 6:
				OptionsX=8
		End Select
		OptionsReelTimer.enabled=true
	end if

	If keycode=RightFlipperKey and State = false and OperatorMenu=1 and Not OptionReelTimer.enabled and Not OptionsReelTimer.enabled then
	  PlaySoundAt SoundFXDOF("reelclick",103,DOFPulse,DOFContactors), OptionsReel
	  Select Case (Options)
		Case 1:
			if Balls=3 then
				Balls=5
				InstCard.image="InstCard5balls"
				OptionsX=2
			  else
				Balls=3
				InstCard.image="InstCard3balls"
				OptionsX=1
			end if
			OptionsReelTimer.enabled=true
		Case 2:
			if ballshadows=0 Then
				ballshadows=1
				BallShadowUpdate.enabled=1
				OptionsX=3    	'yes
			  Else
				ballshadows=0
				BallShadowUpdate.enabled=0
				OptionsX=4		'no
			end If
			OptionsReelTimer.enabled=true
		Case 3:
			if freeplay=0 Then
				freeplay=1
				OptionsX=3    	'yes
			  Else
				freeplay=0
				OptionsX=4		'no
			end if
			OptionsReelTimer.enabled=true
		Case 4:
			Replays=Replays+1
			if Replays>2 then
				Replays=1
			end if
			Replay1=Replay1Table(Replays)
			Replay2=Replay2Table(Replays)
			repcard.image = "replaycard"&replays
			Select Case (replays)
				Case 1:
					OptionsX=5		'Low
				Case 2:
					OptionsX=7		'High
			End Select
			OptionsReelTimer.enabled=true
		Case 5:
			if flippershadows=0 Then
				flippershadows=1
				FlipperLSh.visible=1
				FlipperRSh.visible=1
				OptionsX=3    	'yes
			  Else
				flippershadows=0
				FlipperLSh.visible=0
				FlipperRSh.visible=0
				OptionsX=4		'no
			end If
			OptionsReelTimer.enabled=true
		Case 6:
			OperatorMenu=0
			savehs
			HideOptions
	  End Select
	End If

    If HSEnterMode Then HighScoreProcessKey(keycode)


  if tilt=false and state=true then
	If keycode = LeftFlipperKey Then
		LeftFlipper.RotateToEnd
        PlaySoundAtVol SoundFXDOF("flipperup",101,DOFOn,DOFFlippers), LeftFlipper, 1
        PlaySoundAtVolLoops "Buzz", LeftFlipper, 0.01, -1
	End If
    
	If keycode = RightFlipperKey Then
		RightFlipper.RotateToEnd
        PlaySoundAtVol SoundFXDOF("flipperup",102,DOFOn,DOFFlippers), RightFlipper, 1
        PlaySoundAtVolLoops "Buzz1", RightFlipper, 0.01, -1
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

Sub OptionReelTimer_timer
	PlaySoundAt "fx_switch", OptionsReel
	if OptionReel.roty <> ((Options*60)-60) Then
		optionreel.roty=(OptionReel.roty+10) mod 360
	  Else
		me.enabled=False
	end If
end Sub

Sub OptionsReelTimer_timer
	PlaySoundAt "fx_switch", OptionsReel
	if OptionsReel.roty <> ((OptionsX*45)-45) Then
		OptionsReel.roty=(OptionsReel.roty+9) mod 360
	  Else
		me.enabled=False
	end If
end Sub

Sub OperatorMenuTimer_Timer
	OperatorMenu=1
	Displayoptions
	Options=1
End Sub

Sub DisplayOptions
	OptionBox.visible = true
	OptionReel.visible = True
	OptionsReel.visible = True
	if Balls=3 then
		OptionsReel.roty=0
	  else
		OptionsReel.roty=45
	end if
	OptionReel.roty=0
End Sub

Sub HideOptions
	for each objekt In OptionMenu
		objekt.visible = false
	next
End Sub


Sub TKO_KeyUp(ByVal keycode)

	If keycode = PlungerKey Then
		Plunger.Fire
		if ballhome.BallCntOver=1 then 
			PlaySoundAt "plunger", Plunger			'PLAY WHEN BALL IS HIT
		  Else
			PlaySoundAt "plungerfree", Plunger      'PLAY WHEN NO BALL TO PLUNGE
		end if
	End If

	if keycode = LeftFlipperKey then
		OperatorMenuTimer.Enabled = false
	end if

   If tilt=false and state=true then
	If keycode = LeftFlipperKey Then
		LeftFlipper.RotateToStart
        PlaySoundAtVol SoundFXDOF("flipperdown",101,DOFOff,DOFFlippers), LeftFlipper, 1
		StopSound "Buzz"
	End If
    
	If keycode = RightFlipperKey Then
		RightFlipper.RotateToStart
        PlaySoundAtVol SoundFXDOF("flipperdown",102,DOFOff,DOFFlippers), RightFlipper, 1
		StopSound "Buzz1"
	End If
   End if
End Sub

Sub LeftSlingShot_Slingshot
  if tilt=false then
    PlaySoundat SoundFXDOF("left_slingshot", 103, DOFPulse, DOFContactors), slingL
	addscore 10
	switchblue
    LSling.Visible = 0
    LSling1.Visible = 1
	slingL.objroty = 15	
    LStep = 1
    LeftSlingShot.TimerEnabled = 1
  end if
End Sub

Sub LeftSlingShot_Timer
    Select Case LStep
        Case 3:LSLing1.Visible = 0:LSLing2.Visible = 1:slingL.objroty = 7
        Case 4:LSLing2.Visible = 0:LSLing.Visible = 1:slingL.objroty = 0:LeftSlingShot.TimerEnabled = 0
    End Select
    LStep = LStep + 1
End Sub

sub flippertimer_timer()
	LFlip.RotY = LeftFlipper.CurrentAngle
	RFlip.RotY = RightFlipper.CurrentAngle
	PGate.Rotz = gate.CurrentAngle+25
	if flippershadows=1 then
		FlipperLSh.RotZ = LeftFlipper.currentangle
		FlipperRSh.RotZ = RightFlipper.currentangle
	end if
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
		  Controller.B2SSetScorePlayer i, score(i)  MOD 100000
		  Controller.B2SSetScoreRollover 24 + i, 0
		next
	End If
    if rst=22 then
		newgame
		resettimer.enabled=false
    end if
end sub

Sub addcredit
	if freeplay=0 Then
      credit=credit+1
	  DOF 123, DOFOn
      if credit>25 then credit=25
	  creditreel.setvalue(credit)
	  If B2SOn Then Controller.B2ssetCredits Credit
	end If
End sub

Sub Drain_Hit()
	PlaySoundat "drain", Drain
	DOF 120, DOFPulse
	me.timerenabled=1
End Sub

Sub Drain_timer
	scorebonus.enabled=true
	me.timerenabled=0
End Sub	


sub ballhome_unhit
	DOF 124, DOFPulse
end sub

sub ballrel_hit
	shootagain.state=lightstateoff
end sub

sub scorebonus_timer
	 if shootagain.state=lightstateon then
	    newball
 	    ballreltimer.enabled=true
     else
	  if players=1 or player=players then 
		player=1
		If  B2SOn then Controller.B2ssetplayerup 30, 1
		nextball
	   else
		player=player+1
		If B2SOn then Controller.B2ssetplayerup 30, player
		nextball
	  end if
	end if
    scorebonus.enabled=false
End sub


sub newgame
	for i=1 to 3
		EVAL("Bumper"&i).hashitevent = 1
	Next
	LeftSlingShot.isdropped=False
	player=1
    score(1)=0
	ebcount=0
	If B2SOn then
	  for i = 1 to maxplayers
		Controller.B2SSetScorePlayer i, score(i) 
	  next 
	End If
    eg=0
    rep(1)=0
	advance=1
	for each light in advancelights:light.state=0:next
	AdvanceL(1).state = 1
	LkickExtraAdvance.state = 1
    LspecialL.state=0
    LspecialR.state=0
    ExtraBall.state=0
	shootagain.state=lightstateoff
	for each light in bumperlights:light.state=1:next
	for each light in GI:light.state=1:next
	for each light in starlights:light.state=1:next
	LkickAdvanceR.state=1
	LupRT.state=1
	LupR.state=1
    gamov.text=" "
    tilttxt.text=" "
    If B2SOn then
		Controller.B2SSetGameOver 35,0
		Controller.B2SSetTilt 33,0
		Controller.B2SSetMatch 34,0
	End If
    biptext.text="1"
	matchtxt.text=" "
	ballreltimer.enabled=true
end sub

sub newball
  
End Sub

sub nextball
    if tilt=true then
		for each light in bumperlights:light.state=1:next
		for each light in GI:light.state=1:next
		for i=1 to 3
			EVAL("Bumper"&i).hashitevent = 1
		Next
		LeftSlingShot.isdropped=False
		tilt=false
		tilttxt.text=" "
		If B2SOn then
			Controller.B2SSetTilt 33,0
			Controller.B2ssetdata 6, 1
		End If
    end if
	ebcount=0
	if player=1 then ballinplay=ballinplay+1
	if ballinplay>balls then
		playsound "GameOver"
		eg=1
		ballreltimer.enabled=true
	  else
		if state=true and tilt=false then
		  newball
		  ballreltimer.enabled=true
		end if
		biptext.text=ballinplay
		If B2SOn then Controller.B2ssetballinplay 32, Ballinplay
	end if
End Sub

sub ballreltimer_timer
  if eg=1 then
	  hiscstate=0
      matchnum
	  biptext.text=" "
	  state=false
	  gamov.text="GAME OVER"
	  gamov.timerenabled=1
	  tilttxt.timerenabled=1
	  for i=1 to maxplayers
		if score(i)>hisc then 
			hisc=score(i)
			hiscstate=1
		end if
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
	    Controller.B2sStartAnimation "EOGame"
	    Controller.B2SSetScorePlayer 5, hisc
	    Controller.B2ssetcanplay 31, 0
	    Controller.B2ssetcanplay 30, 0
	  End If
	  turnoff
  else
	Drain.kick 60,11,0
	PlaySoundat SoundFXDOF("drainkick",122, DOFPulse, DOFContactors), Drain
  end if
    ballreltimer.enabled=false
end sub

Sub HStimer_timer
	PlaySoundAtVol SoundFXDOF("knocker",121,DOFPulse,DOFKnocker), Plunger, 1
	DOF 139,DOFPulse
	HStimer.uservalue=HStimer.uservalue+1
	if HStimer.uservalue=3 then me.enabled=0
end sub

sub matchnum
	if matchnumb=0 then
		matchtxt.text="00"
	  else
		matchtxt.text=matchnumb*10
	end if
  If B2SOn then Controller.B2SSetMatch 34,Matchnumb*10
  For i=1 to players
    if (matchnumb*10)=(score(i) mod 100) then 
		addcredit
		PlaySoundAtVol SoundFXDOF("knocker", 121, DOFPulse, DOFKnocker), Plunger, 1
	end if
  next
end sub

'********** Bumpers

Sub Bumper1_Hit
   if state=true and tilt=false then
	PlaySoundAt SoundFXDOF("fx_bumper4", 105, DOFPulse, DOFContactors), Bumper1
	if Lbumpers.state=1 then
		addscore bumperlitscore
	  else
		addscore bumperoffscore
	end if
   end if
End Sub

Sub Bumper2_Hit
   if state=true and tilt=false then
	PlaySoundat SoundFXDOF("fx_bumper4", 104, DOFPulse, DOFContactors), Bumper2
	if Lbumpers.state=1 then
		addscore bumperlitscore
	  else
		addscore bumperoffscore
	end if	
   end if
End Sub

Sub Bumper3_Hit
   if state=true and tilt=false then
	PlaySoundat SoundFXDOF("fx_bumper4", 106, 2, DOFContactors), Bumper3
	if Lbumpers.state=1 then
		addscore bumperlitscore
	  else
		addscore bumperoffscore
	end if	
   end if
End Sub


sub FlashBumpers
	for each light in bumperlights: light.duration 0, 350, 1: Next
end sub

Sub Dingwalls_Hit (idx)
	if state=true and tilt=false then 
		addscore 10
		switchblue
	end if
end sub

'********** Dingwalls (scoring rubbers) - and/or animated Rubbers- timer 50

sub dingwalla_hit
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

sub dingwalle_hit
	SlingE.visible=0
	Slinge1.visible=1
	me.uservalue=1
	Me.timerenabled=1
end sub

sub dingwalle_timer									
	select case me.uservalue
		Case 1: Slinge1.visible=0: SlingE.visible=1
		case 2:	SlingE.visible=0: SlingE2.visible=1
		Case 3: Slinge2.visible=0: SlingE.visible=1: Me.timerenabled=0
	end Select
	me.uservalue=me.uservalue+1
end sub

Sub switchblue
	if LkickAdvanceL.state=1 then
		LkickAdvanceL.state=0
		LkickAdvanceR.state=1
	  else
		LkickAdvanceR.state=0
		LkickAdvanceL.state=1
	end if
	if LupLT.state=1 then
		LupLT.state=0
		LupRt.state=1
	  else
		LupRT.state=0
		LupLT.state=1
	end if
	if LupR.state=1 then
		LupR.state=0
	  else
		LupR.state=1
	end if
	whitecount=whitecount+1
	if whitecount>3 then whitecount=0
	if advance = 2 or advance=6 or advance=9 or advance =13 then
		for each light in advancedlights: light.state=0: next
		whites(whitecount).state=1
	end if
end sub


'**********  Targets

sub TGupL_hit
	DOF 112, DOFPulse
	if LupLT.state=1 then 
		addadvance
	end if
	if LupT.state=1 then
		addscore 5000
	  else
		addscore 500
	end if		
end sub

sub TGupR_hit
	DOF 112, DOFPulse
	if LupRT.state=1 then 
		addadvance
	end if
	if LupT.state=1 then
		addscore 5000
	  else
		addscore 500
	end if		
end sub

sub TGmidR_hit
	if ExtraBall.state=1 and ebcount=0 then 
		addscore 5000
		playsound SoundFXDOF("gong", 128, DOFPulse, DOFContactors)
		shootagain.state=1
		ebcount = 1
	  else
		if ebcount=1 then 
			addscore 5000
		  else
			addscore 500
		end if
	end if
end sub

'*************  Kickers

sub KickerL_hit
	if LkickAdvanceL.state=1 then
		addadvance
		addscore 5000
	  else
		addscore 500
	end if
	if LkickExtraAdvance.state=1 then addadvance
	if LspecialL.state=1 then 
		playsound SoundFXDOF("gong", 128, DOFPulse, DOFContactors)
		addcredit
	end if
	me.uservalue=1
	me.timerenabled=1
end sub

Sub KickerL_timer
	Select case me.uservalue
	  case 3:
		KickerL.Kick 141, 8
		PlaySoundAt SoundFXDOF("holekick", 113, DOFPulse, DOFContactors), KickerL
		PkickarmL.rotz=15
	  case 5:
		PkickarmL.rotz=0
		me.timerenabled=0
	end Select
	me.uservalue=me.uservalue+1
End Sub

sub KickerR_hit
	if LkickAdvanceR.state=1 then
		addadvance
		addscore 5000
	  else
		addscore 500
	end if
	if LkickExtraAdvance.state=1 then addadvance
	if LspecialR.state=1 then 
		playsound SoundFXDOF("gong", 128, DOFPulse, DOFContactors)
		addcredit
	end if
	me.uservalue=1
	me.timerenabled=1
end sub

Sub KickerR_timer
	Select case me.uservalue
	  case 3:
		KickerR.Kick 145, 15
		PlaySoundat SoundFXDOF("holekick", 113, DOFPulse, DOFContactors), KickerR
		PkickarmR.rotz=15
	  case 5:
		PkickarmR.rotz=0
		me.timerenabled=0
	end Select
	me.uservalue=me.uservalue+1
End Sub

'****** Wire Triggers

Sub TupR_Hit
   if tilt=false then
	DOF 118, DOFPulse
	flashbumpers
	If LupR.State = 1 then 
		addscore 5000
		addadvance
	  else
		addscore 500
	end if
   end if
End Sub

Sub TmidR_Hit
   if tilt=false then
	DOF 119, DOFPulse
	flashbumpers
	addadvance
	If LmidR.State = 1 then 
		addscore 5000
	  else
		addscore 500
	end if
   end if
End Sub

Sub TOutR_Hit
   if tilt=false then
	DOF 116, DOFPulse
	flashbumpers
	addadvance
	if LOutR.state = 1 then 
		addscore 5000
	  else
		addscore 500
	end if
   end if
End Sub

Sub TOutL_Hit
   if tilt=false then
	DOF 114, DOFPulse
	flashbumpers
	addadvance
	if LOutL.state = 1 then 
		addscore 5000
	  else
		addscore 500
	end if
   end if
End Sub

Sub TInL_Hit
   if tilt=false then
	DOF 115, DOFPulse
	flashbumpers
	if LOutL.state = 1 then 
		addscore 5000
	  else
		addscore 500
	end if
   end if
End Sub

'**** Rollovers

Sub TStar1_Hit
	if tilt=false then
		flashbumpers
		DOF 107, DOFPulse
		if LStar1.state=1 then
			addscore 5000
			Lstar1.state=0
		  else
			addscore 500
		end if
		starcheck
	end if
end sub

Sub TStar2_Hit
	if tilt=false then
		flashbumpers
		DOF 108, DOFPulse
		if Lstar2.state=1 then
			addscore 5000
			Lstar2.state=0
		  else
			addscore 500
		end if
		starcheck
	end if
end sub

Sub TStar3_Hit
	if tilt=false then
		flashbumpers
		DOF 109, DOFPulse
		if Lstar3.state=1 then
			addscore 5000
			Lstar3.state=0
		  else
			addscore 500
		end if
		starcheck
	end if
end sub

Sub TStar4_Hit
	if tilt=false then
		flashbumpers
		DOF 110, DOFPulse
		if Lstar4.state=1 then
			addscore 5000
			Lstar4.state=0
		  else
			addscore 500
		end if
		starcheck
	end if
end sub

Sub TStar5_Hit
	if tilt=false then
		flashbumpers
		DOF 111, DOFPulse
		if Lstar5.state=1 then
			addscore 5000
			Lstar5.state=0
		  else
			addscore 500
		end if
		starcheck
	end if
end sub

sub starcheck
	if Lstar1.state + Lstar2.state + Lstar3.state + Lstar4.state + Lstar5.state = 0 then
		if balls=3 then 
			for i = 1 to 5: addadvance: Next
		  Else
			for i = 1 to 2: addadvance: Next
		end If
		for each light in starlights: light.state=1: next
	end if
end sub		

sub addscore(points)
  if tilt=false then
    If Points < 100 and AddScore10Timer.enabled = false Then
        Add10 = Points \ 10
		matchnumb=matchnumb+1
		if matchnumb>9 then matchnumb=0
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

Sub AddPoints(Points)			  ' Sounds: there are 3 sounds: tens, hundreds and thousands
    score(player)=score(player)+points
	sreels(player).addvalue(points)
	If B2SOn Then Controller.B2SSetScorePlayer player, score(player)

    If Points = 100 AND(Score(player) MOD 1000) \ 100 = 0 Then  'New 1000 reel
        PlaySoundAt SoundFXDOF("bell1000",125,DOFPulse,DOFChimes), chimesound
      ElseIf Points = 10 AND(Score(player) MOD 100) \ 10 = 0 Then 'New 100 reel
        PlaySoundAt SoundFXDOF("bell100",126,DOFPulse,DOFChimes), chimesound
      ElseIf points = 1000 Then
        PlaySoundAt SoundFXDOF("bell1000",127,DOFPulse,DOFChimes), chimesound
	  elseif Points = 100 Then
        PlaySoundAt SoundFXDOF("bell100",126,DOFPulse,DOFChimes), chimesound
	  Else
        PlaySoundAt SoundFXDOF("bell10",125,DOFPulse,DOFChimes), chimesound
    End If
	checkreplays
end sub

Sub checkreplays
    ' check replays and rollover
	if score(player)>99999 then
		If B2SOn then Controller.B2SSetScoreRollover 24 + player, 1
		EVAL("Reel100K"&player).setvalue(int(score(player)/100000))
	End if
    if score(player)=>replay1 and rep(player)=0 then
		addcredit
		rep(player)=1
		PlaySoundAt SoundFXDOF("knocker",121,DOFPulse,DOFKnocker), plunger
    end if
    if score(player)=>replay2 and rep(player)=1 then
		addcredit
		rep(player)=2
		PlaySoundAt SoundFXDOF("knocker",121,DOFPulse,DOFKnocker), plunger
    end if
    if score(player)=>replay3 and rep(player)=2 then
		addcredit
		rep(player)=3
		PlaySoundAt SoundFXDOF("knocker",121,DOFPulse,DOFKnocker), plunger
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

Sub GameTilted
	Tilt = True
	tilttxt.text="TILT"
    If B2SOn Then Controller.B2SSetTilt 33,1
    If B2SOn Then Controller.B2ssetdata 6, 0
	playsoundat "tilt", Plunger
	turnoff
End Sub

sub turnoff
	if tilt=true then
		For each light in BumperLights:light.State = 0: Next
		For each light in GI:light.State = 0: Next
	end If

	for i=1 to 3
		EVAL("Bumper"&i).hashitevent = 0
	Next
	LeftSlingShot.isdropped=true
	LeftFlipper.RotateToStart
	StopSound "Buzz"
	RightFlipper.RotateToStart
	StopSound "Buzz1"
end sub    


Sub addadvance
	advance=advance+1
	if advance>16 then advance=16
	advanceL(advance).state=1
	if advance>1 then advanceL(advance-1).state=0
	for each light in advancedlights: light.state=0: next
	select case(advance)
		case 1, 5, 10, 14:
			LkickExtraAdvance.state=1
		case 2, 6, 9, 13:  'white lights
			whites(whitecount).state=1
		case 3, 7, 11:
			LupT.state=1
		case 4, 8, 12, 15:
			Lbumpers.state=1
		case 16:			
			if LspecialL.state=0 then
				LspecialL.state=1
				LspecialR.state=0
			  else
				LspecialL.state=0
				LspecialR.state=1
			end if
	end select
End sub


sub savehs

    savevalue "TKO", "credit", credit
    savevalue "TKO", "hiscore", hisc
    savevalue "TKO", "match", matchnumb
    savevalue "TKO", "score1", score(1)
	savevalue "TKO", "replays", replays
	savevalue "TKO", "balls", balls
	savevalue "TKO", "freeplay", freeplay
	savevalue "TKO", "ballshadows", ballshadows
	savevalue "TKO", "flippershadows", flippershadows
	savevalue "TKO", "hsa1", HSA1
	savevalue "TKO", "hsa2", HSA2
	savevalue "TKO", "hsa3", HSA3
end sub

sub loadhs
    dim temp
	temp = LoadValue("TKO", "credit")
    If (temp <> "") then credit = CDbl(temp)
    temp = LoadValue("TKO", "hiscore")
    If (temp <> "") then hisc = CDbl(temp)
    temp = LoadValue("TKO", "match")
    If (temp <> "") then matchnumb = CDbl(temp)
    temp = LoadValue("TKO", "score1")
    If (temp <> "") then score(1) = CDbl(temp)
    temp = LoadValue("TKO", "replays")
    If (temp <> "") then replays = CDbl(temp)
    temp = LoadValue("TKO", "balls")
    If (temp <> "") then balls = CDbl(temp)
    temp = LoadValue("TKO", "freeplay")
    If (temp <> "") then freeplay = CDbl(temp)
    temp = LoadValue("TKO", "ballshadows")
    If (temp <> "") then ballshadows = CDbl(temp)
    temp = LoadValue("TKO", "flippershadows")
    If (temp <> "") then flippershadows = CDbl(temp)
    temp = LoadValue("TKO", "hsa1")
    If (temp <> "") then HSA1 = CDbl(temp)
    temp = LoadValue("TKO", "hsa2")
    If (temp <> "") then HSA2 = CDbl(temp)
    temp = LoadValue("TKO", "hsa3")
    If (temp <> "") then HSA3 = CDbl(temp)
end sub

Sub TKO_Exit()
	savehs
	If B2SOn Then Controller.stop
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

Function AudioFade(tableobj) ' Fades between front and back of the table (for surround systems or 2x2 speakers, etc), depending on the Y position on the table. "TKO" is the name of the table
	Dim tmp
    tmp = tableobj.y * 2 / TKO.height-1
    If tmp > 0 Then
		AudioFade = Csng(tmp ^10)
    Else
        AudioFade = Csng(-((- tmp) ^10) )
    End If
End Function

Function AudioPan(tableobj) ' Calculates the pan for a tableobj based on the X position on the table. "TKO" is the name of the table
    Dim tmp
    tmp = tableobj.x * 2 / TKO.width-1
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
        If BOT(b).X < TKO.Width/2 Then
            BallShadow(b).X = ((BOT(b).X) - (Ballsize/16) + ((BOT(b).X - (TKO.Width/2))/17))' + 13
        Else
            BallShadow(b).X = ((BOT(b).X) + (Ballsize/16) + ((BOT(b).X - (TKO.Width/2))/17))' - 13
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
Sub TKO_exit()
  Controller.Pause = False
  Controller.Stop
End Sub


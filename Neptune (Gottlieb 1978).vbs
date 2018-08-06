'	-------------------------------------------------
'	NEPTUNE by Gottlieb 1978
'
'	built on BorgDog's Gottlieb EM 4 player VPX table blank with options menu
'	-------------------------------------------------
'	
'	scripted by BorgDog, 2018
'	artwork by HauntFreaks
'
'	High Score sticky routines from mfuegemann's Fast Draw VPX table
'		- flippers to change letter, Start to Select
'	Ball control script from rothbauerw
'		- press C during play to control ball, use arrows to move the ball
'	Option menu and player up light rotation borrowed from loserman76 and gnance 
'		- hold down left shift during Game Over to bring up menu
'	Ball shadows from ninuzzu 
'		- set option below to enable or disable
'	primitives from Dark, zany, sliderpoint, hauntfreaks, borgdog and I'm sure others
'	Instruction cards by Inkochinito 
'
'	Layer usage generally
'		1 - most playfield parts and stuff
'		2 - triggers
'		3 - hi score sticky
'		4 - options menu, ShadowRamp
'		5 - playfield holes and other primitives
'		6 - GI lighting
'		7 - plastics
'		8 - insert lights
'
'	Basic DOF config,
'		101 Left Flipper, 102 Right Flipper, 
'		103 Left sling, 104 Left sling flasher, 105 right sling, 106 right sling flasher,
'		107 Bumper1, 108 Bumper1 flasher, 109 bumper2, 110 bumper2 Flasher
'		111 Bumper3, 112 Bumper3 flasher
'		115 Left Kicker, 116 Top Kicker, 117 Right Kicker
'		118 Target Queen and Target King
'		134 Drain_Hit flasher, 135 Drain_kick
'		136 Shooter Lane/launch ball, 137 credit light, 138 knocker, 139 knocker/kicker Flasher
'		141 Chime1-10s, 142 Chime2-100s, 143 Chime3-1000s
'
'	-------------------------------------------------

Option Explicit
Randomize

Const cGameName = "neptune"
Const Ballsize = 50				'used by ball shadow routine, doesn't actually change ball size

'*************** needed for bumper skirt animation
'*************** NOTE: set bumper object timer to around 150-175 in order to be able
'***************       to actually see the animation, adjust to your liking

Const PI = 3.1415926
Const SkirtTilt=5		'angle of skirt tilting in degrees

'*************** end of bumper skirt animation constants

Dim operatormenu, options
Dim bumperlitscore
Dim bumperoffscore
Dim balls, TempPlayerUp, rotatortemp
Dim replays, freeplay
Dim Replay1Table(3), Replay2Table(3), Replay3Table(3)
dim replay1, replay2
Dim Add10, Add100, Add1000
Dim hisc
Dim maxplayers, players, player
Dim credit, wowred, wowblack, wows, wowspot, rbspot
Dim score(4)
Dim state, redsequence, blacksequence
Dim tilt, tiltsens
Dim ballinplay
Dim matchnumb, mysterycnt
dim ballrenabled
dim rstep, lstep
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
Dim FlipperShadows: FlipperShadows=1  '***********set to 1 to turn on Flipper shadows

sub Neptune_init
	LoadEM
	wowspot=0
	rbspot=0
	maxplayers=1
	mysterycnt=1
	Replay1Table(1)=70000
	Replay1Table(2)=90000
	Replay1Table(3)=110000

	Replay2Table(1)=120000
	Replay2Table(2)=140000
	Replay2Table(3)=160000

	hideoptions
	player=1
	RotatorTemp=1
	balls=5   		'SET DEFAULT VALUES PRIOR TO LOADING FROM SAVED FILE
	replays=1
	freeplay=1
	hisc=50000
	HSA1=4
	HSA2=15
	HSA3=7
	matchnumb=0
	loadhs		'LOADS SAVED OPTIONS AND HIGH SCORE IF EXISTING
	UpdatePostIt	'UPDATE HIGH SCORE STICKY
	for i = 1 to maxplayers
		EVAL("ScoreReel"&i).setvalue score(i)
		EVAL("Reel100K"&i).setvalue(int(score(i)/100000))
	next
	bipreel.setvalue 0

'	credittxt.setvalue(credit)
	Replay1=Replay1Table(Replays)
	Replay2=Replay2Table(Replays)
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

    Drain.CreateBall
End sub

sub startGame_timer
	PlaySoundAt "poweron", Plunger
	lightdelay.enabled=true
	me.enabled=false
end sub

sub lightdelay_timer
	gamov.text="GAME OVER"
	if matchnumb=0 then
		matchtxt.text="00"
	  else
		matchtxt.text=matchnumb*10
	end if
	if credit>0 or freeplay=1 then 
		creditlight.state=1
		DOF 137, DOFOn
	  else
		creditlight.state=0
	end if
	For each light in GIlights:light.state=1:Next
	For each light in BumperLights:light.state=1:Next
	if B2SOn then
		for i = 1 to maxplayers
			Controller.B2SSetScorePlayer i, Score(i) MOD 100000
			if score(player)>1000000 then 
				Controller.b2ssetscorerollover i+24, 10
			  Else
				If score(i)>100000 Then Controller.b2ssetscorerollover i+24, Int(score(i)/100000)
			end if
		next
		Controller.B2SSetData 100,1
'		Controller.B2ssetCredits Credit
'		Controller.B2ssetMatch 34, Matchnumb
		Controller.B2SSetGameOver 35,1
		Controller.B2SSetScorePlayer 5, hisc
	end if
	me.enabled=0
end Sub


Sub Neptune_KeyDown(ByVal keycode)
   
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

	if keycode=AddCreditKey then
		if freeplay=1 then 
			PlaySoundAt "coinreturn", Drain
		  else
			PlaySoundAt "coinin" , Drain
		end if
		coindelay.enabled=true 
    end if

    if keycode=StartGameKey and (Credit>0 or freeplay=1) And Not HSEnterMode=true and Not startgame.enabled then
	  if state=false then
		player=1
		if freeplay=0 then 
			credit=credit-1
			if credit>0 or freeplay=1 then 
				creditlight.state=1
				DOF 137, DOFOn
			  else
				creditlight.state=0
				DOF 137, DOFOff
			end if
'			credittxt.setvalue(credit)
		End if
		PlaySoundAt "cluper", soundtrigger
		ballinplay=balls
		If B2SOn Then 
			Controller.B2ssetCredits Credit
			Controller.B2ssetballinplay 32, Ballinplay
			Controller.B2Ssetscorerollover 25, 0
			Controller.B2ssetplayerup 30, 1
			Controller.B2SSetGameOver 0
			Controller.B2SSetScorePlayer 5, hisc
		End If
		for i = 1 to maxplayers
			score(i)=0
			rep(i)=0
			EVAL("Pup"&i).state=0
			EVAL("ScoreReel"&i).resettozero
			EVAL("Reel100K"&i).setvalue 0
			If B2SOn then Controller.B2SSetScorePlayer i, score(i)
		next
	    Pup1.state=1
		PlaySound("RotateThruPlayers"),0,.05,0,0.25
		TempPlayerUp=Player
		PlayerUpRotator.enabled=true
		tilt=false
		state=true
		PlaySoundAt "initialize", soundtrigger
		players=1
		rst=0
		newgame.enabled=true
	  else if state=true and players < maxplayers and Ballinplay=balls then
		if freeplay=0 then 
			credit=credit-1
			if credit>0 or freeplay=1 then 
				creditlight.state=1
				DOF 137, DOFOn
			  else
				creditlight.state=0
				DOF 137, DOFOff
			end if
		End if
'		credittxt.setvalue(credit)
		players=players+1

		If B2SOn then
			Controller.B2ssetCredits Credit

		End If
		playsoundat "cluper", soundtrigger
	   end if 
	  end if
	end if

	If keycode = PlungerKey Then
		Plunger.PullBack
		PlaySoundat "plungerpull", Plunger
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
		PlaySoundAtVol SoundFXDOF("fx_flipperup",101,DOFOn,DOFContactors), LeftFlipper, 1
		PlaySoundAtVolLoops "Buzz", LeftFlipper, 0.01, -1

	End If
    
	If keycode = RightFlipperKey Then
		RightFlipper.RotateToEnd
		PlaySoundAtVol SoundFXDOF("fx_flipperup",102,DOFOn,DOFContactors), RightFlipper, 1
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


Sub Neptune_KeyUp(ByVal keycode)

	if keycode = 203 then Cleft = 0' Left Arrow

	if keycode = 200 then Cup = 0' Up Arrow

	if keycode = 208 then Cdown = 0' Down Arrow

	if keycode = 205 then Cright = 0' Right Arrow

	If keycode = PlungerKey Then
		Plunger.Fire
		if Ballhome.BallCntOver<>0 then
			PlaySoundAt "plunger", Plunger				'PLAY WHEN BALL IS HIT
		  else
			PlaySoundAt "plungerreleasefree", Plunger	'PLAY WHEN NO BALL TO PLUNGE
		end if
	End If

	if keycode = LeftFlipperKey then
		OperatorMenuTimer.Enabled = false
	end if

   If tilt=false and state=true then
	If keycode = LeftFlipperKey Then
		LeftFlipper.RotateToStart
		PlaySoundAtVol SoundFXDOF("fx_flipperdown",101,DOFOff,DOFContactors), LeftFlipper, 1
		StopSound "Buzz"
	End If
    
	If keycode = RightFlipperKey Then
		RightFlipper.RotateToStart
		PlaySoundAtVol SoundFXDOF("fx_flipperdown",102,DOFOff,DOFContactors), RightFlipper, 1
		StopSound "Buzz1"
	End If
   End if
End Sub

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

sub flippertimer_timer()
	LFlip.RotY = LeftFlipper.CurrentAngle
	RFlip.RotY = RightFlipper.CurrentAngle
	Pgate.rotx = Gate.currentangle*.6

	if FlipperShadows=1 then
		FlipperLSh.RotZ = LeftFlipper.currentangle
		FlipperRSh.RotZ = RightFlipper.currentangle
	end if

end sub


Sub PairedlampTimer_timer
	bumperlighta1.state = bumperlight1.state
	bumperlighta2.state = bumperlight2.state
	bumperlighta3.state = bumperlight3.state
	LspotsB1.state = LspotsB.state
	LspotsR1.state = LspotsR.state
	LaceD2.state = LaceD.state
	LkingS2.state = LkingS.state
	Pup1a.state = pup1.state
	Pup1a1.state = pup1.state
	Pup1a2.state = pup1.state
	Pup1a3.state = pup1.state
end sub

sub coindelay_timer
	addcredit
    coindelay.enabled=false
end sub

Sub addcredit
      credit=credit+1
	  DOF 137, DOFOn
	  creditlight.state=1
      if credit>25 then credit=25
'	  credittxt.setvalue(credit)
	  If B2SOn Then Controller.B2ssetCredits Credit
End sub

Sub Drain_Hit()
	DOF 134, DOFPulse
	PlaySoundAt "drain", Drain
	me.timerenabled=1
End Sub

Sub Drain_timer
	scorebonus.enabled=true
	me.timerenabled=0
End Sub	


sub ballhome_unhit
	DOF 136, DOFPulse
end sub


sub scorebonus_timer			'************add bonus scoring if applicable
   	  if players=1 or player=players then 
		player=1
	   Else
		player=player+1
	  end if
	  If B2SOn then Controller.B2ssetplayerup 30, player
	  nextball

	 scorebonus.enabled=false
End sub


sub newgame_timer
    eg=0
	for each light in Cardlights:light.state=1:next
	for each light in Starlights:light.state=0:next
	for each light in wowLightsRed:light.state=0:next
	for each light in wowlightsBlack:light.state=0:next
	wowblack=0
	wowred=0
	movespots
	wows=0
    tilttxt.text=" "
	gamov.text=" "
	for i=1 to 2
		EVAL("Bumper"&i).hashitevent = 1
	Next
	LeftSlingShot.isdropped=False
	RightSlingShot.isdropped=False
    If B2SOn then
		Controller.B2SSetGameOver 35,0
		Controller.B2SSetTilt 33,0
		Controller.B2SSetMatch 34,0
	End If
	bipreel.setvalue balls
	wowsreel.setvalue 0
    matchtxt.text=" "
	newball
	ballrelTimer.enabled=true
	newgame.enabled=false
end sub


sub newball
	for each light in GIlights:light.state=1:next
	If B2SOn then Controller.B2SSetData 100,1
	if wowblack=1 then resetblack
	if wowred=1 then resetred
End Sub

sub resetblack
	for each light in wowLightsBlack: light.state=0:Next
	for each light in CardLightsBlack: light.state=1: Next
	blacksequence=0
	wowblack=0
end Sub

sub resetred
	for each light in wowLightsRed: light.state=0:Next
	for each light in CardLightsRed: light.state=1: Next
	redsequence=0
	wowred=0
end Sub

sub nextball
    if tilt=true then
	  for i=1 to 3
		EVAL("Bumper"&i).hashitevent = 1
	  Next
	  LeftSlingShot.isdropped=False
	  RightSlingShot.isdropped=False
      tilt=false
      tilttxt.text=" "
		If B2SOn then
			Controller.B2SSetTilt 33,0
			Controller.B2ssetdata 1, 1
		End If
    end if
	if player=1 then 
		if wows>0 then
			wows=wows-1
		  else
			ballinplay=ballinplay-1
		end if
	end if
	if ballinplay=0 then
		playsound "GameOver"
		eg=1
		ballreltimer.enabled=true
	  else
		For each objekt in PlayerHUDScores
			objekt.state=0
		next
		PlayerHUDScores(Player-1).state=1
		PlaySound("RotateThruPlayers"),0,.05,0,0.25
		TempPlayerUp=Player
		PlayerUpRotator.enabled=true
		if state=true then
		  newball
		  ballreltimer.enabled=true
		end if
		If ballinplay+wows<=balls then
			wowsreel.setvalue 0
			bipreel.setvalue ballinplay
			If B2SOn then 
				Controller.B2ssetballinplay 32, Ballinplay
				Controller.B2Ssetdata 11, 0
			end if
		  else
			bipreel.setvalue Balls
			if B2SOn then Controller.B2ssetballinplay 32, Balls
			wowsreel.setvalue wows
			if B2SOn then 
				for i= 1 to 5
					Controller.B2ssetdata 10+i, 0
				next
				for i= 1 to wows
					Controller.B2ssetdata 10+i, 1
				next
			end if
		end if
	end if
End Sub

sub ballreltimer_timer
  dim hiscstate
  if eg=1 then
	  hiscstate=0
      matchnum
	  state=false
	  bipreel.setvalue 0
	  wowsreel.setvalue 0
	  gamov.text="GAME OVER"
	  gamov.timerenabled=1
	  tilttxt.timerenabled=1

	  for i=1 to maxplayers
		if score(i)>hisc then 
			hisc=score(i)
			hiscstate=1
		end if
		EVAL("Pup"&i).state=0
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
	    Controller.B2SSetScorePlayer 5, hisc
	  End If
	  ballreltimer.enabled=false
  else
	Drain.kick 60,45,0
    ballreltimer.enabled=false
	PlaySoundAtVol SoundFXDOF("kickerkick",135,DOFPulse,DOFContactors), Drain, 1
  end if
end sub

Sub HStimer_timer
	PlaySoundAtVol SoundFXDOF("knock",138,DOFPulse,DOFKnocker), Plunger, 1
	DOF 139,DOFPulse
	HStimer.uservalue=HStimer.uservalue+1
	if HStimer.uservalue=3 then me.enabled=0
end sub

Sub PlayerUpRotator_timer()
		If RotatorTemp<maxplayers+1 then
			TempPlayerUp=TempPlayerUp+1
			If TempPlayerUp>maxplayers then
				TempPlayerUp=1
			end if

			For each objekt in PlayerHUDScores
				objekt.state=0
			next
			PlayerHUDScores(TempPlayerUp-1).state=1
			If B2SOn Then Controller.B2SSetPlayerUp TempPlayerUp
		else
			if B2SOn then Controller.B2SSetPlayerUp Player
			PlayerUpRotator.enabled=false
			RotatorTemp=1

			For each objekt in PlayerHUDScores
				objekt.state=0
			next
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
		  PlaySoundAtVol SoundFXDOF("knock",138,DOFPulse,DOFKnocker), Plunger, 1
		  DOF 139,DOFPulse
	    end if
    next
end sub

'********** Bumpers

Sub Bumper1_Hit
	if tilt=false then
		PlaySoundAtVol SoundFXDOF("fx_bumper4", 107, DOFPulse, DOFContactors), Bumper1, 1
		DOF 108,DOFPulse
		movespots
		movewow
		if balls=3 then
			addscore 100
		  else
			addscore 500
		end if
	end if
End Sub


Sub Bumper2_Hit
	if tilt=false then
		PlaySoundAtVol SoundFXDOF("fx_bumper4", 109, DOFPulse, DOFContactors), Bumper2, 1
		DOF 110,DOFPulse
'		movespots
'		movewow
		if Balls=3 then
			addscore 1000
		  else
			addscore 100
		end if	
	end if
End Sub

Sub Bumper3_Hit
	if tilt=false then
		PlaySoundAtVol SoundFXDOF("fx_bumper4", 111, DOFPulse, DOFContactors), Bumper3, 1
		DOF 112,DOFPulse
		movespots
		movewow
		if balls=3 then
			addscore 100
		  else
			addscore 500
		end if
	end if
End Sub


sub FlashBumpers
	if bumperlight1.state = 1 then
		for each light in BumperLights: light.duration 0, 500, 1:Next
	end If
end sub

Sub movespots
	if LspotsB.state=1 Then
		LspotsB.state=0
		LspotsR.state=1
	  Else
		LspotsB.state=1
		LspotsR.state=0
	end if
end Sub

'************** Slings

Sub RightSlingShot_Slingshot
	PlaySoundAtVol SoundFXDOF("right_slingshot",105,DOFPulse,DOFContactors), slingR, 1
	DOF 106,DOFPulse
	addscore 10
    RSling.Visible = 0
    RSling1.Visible = 1
	slingR.rotx = 20	
    RStep = 1
    RightSlingShot.TimerEnabled = 1
End Sub

Sub RightSlingShot_Timer
    Select Case RStep
        Case 3:RSLing1.Visible = 0:RSLing2.Visible = 1:slingR.rotx = 10
        Case 4:	slingR.rotx = 0:RSLing2.Visible = 0:RSLing.Visible = 1:RightSlingShot.TimerEnabled = 0
    End Select
    RStep = RStep + 1
End Sub

Sub LeftSlingShot_Slingshot
	PlaySoundAtVol SoundFXDOF("left_slingshot",103,DOFPulse,DOFContactors), slingL, 1
	DOF 104,DOFPulse
	addscore 10
    LSling.Visible = 0
    LSling1.Visible = 1
	slingL.rotx = 20	
    LStep = 1
    LeftSlingShot.TimerEnabled = 1
End Sub

Sub LeftSlingShot_Timer
    Select Case LStep
        Case 3:LSLing1.Visible = 0:LSLing2.Visible = 1:slingL.rotx = 10
        Case 4:slingL.rotx = 0:LSLing2.Visible = 0:LSLing.Visible = 1:LeftSlingShot.TimerEnabled = 0
    End Select
    LStep = LStep + 1
End Sub

'********** Dingwalls  (scoring rubbers) - animated - timer 50

sub dingwalla_hit
	if state=true and tilt=false then addscore 10
	SlingA.visible=0
	SlingA1.visible=1
	me.uservalue=1
	Me.timerenabled=1
end sub

sub dingwalla_timer									'default 50 timer
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

sub dingwallb_timer									'default 50 timer
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

sub dingwallc_timer									'default 50 timer
	select case me.uservalue
		Case 1: Slingc1.visible=0: Slingc.visible=1
		case 2:	Slingc.visible=0: Slingc2.visible=1
		Case 3: Slingc2.visible=0: Slingc.visible=1: Me.timerenabled=0
	end Select
	me.uservalue=me.uservalue+1
end sub

'********** Triggers     


sub TGnineD_hit 
  if not tilt then
	addscore 500
	LnineD.state=0
	if LnineS.state=0 then LstarNines.state=1
	if LnineDwow.state=1 then 
		addwow 1
		wowred=1
	end if
	Sequences
  end if
end sub

sub TGtenD_hit 
  if not tilt then
	addscore 500
	LtenD.state=0
	if LtenS.state=0 then LstarTens.state=1
	if LtenDwow.state=1 then 
		addwow 1
		wowred=1
	end if
	Sequences
  end if
end sub

sub TGjackD_hit 
  if not tilt then
	addscore 500
	LjackD.state=0
	if LjackS.state=0 then LstarJacks.state=1
	if LjackDwow.state=1 then 
		addwow 1
		wowred=1
	end if
	Sequences
  end if
end sub

sub TGqueenD_hit 
  if not tilt then
	addscore 500
	LqueenD.state=0
	if LqueenS.state=0 then LstarQueens.state=1
	if LqueenDwow.state=1 then 
		addwow 1
		wowred=1
	end if
	Sequences
  end if
end sub

sub TGkingD_hit 
  if not tilt then
	addscore 500
	LkingD.state=0
	if LkingS.state=0 then LstarKings.state=1
	if LkingDwow.state=1 then 
		addwow 1
		wowred=1
	end if
	Sequences
  end if
end sub

sub TGaceD_hit 
  if not tilt then
	addscore 500
	LaceD.state=0
	if LaceS.state=0 then LstarAces.state=1
	if LaceDwow.state=1 then 
		addwow 1
		wowred=1
	end if
	Sequences
  end if
end sub

sub TGaceD2_hit 
  if not tilt then
	addscore 500
	LaceD.state=0
	if LaceS.state=0 then LstarAces.state=1
	if LaceD2wow.state=1 then 
		addwow 1
		wowred=1
	end if
	Sequences
  end if
end sub

sub TGnineS_hit 
  if not tilt then
	addscore 500
	LnineS.state=0
	if LnineD.state=0 then LstarNines.state=1
	if LnineSwow.state=1 then 
		addwow 1
		wowblack=1
	end if
	Sequences
  end if
end sub

sub TGtenS_hit 
  if not tilt then
	addscore 500
	LtenS.state=0
	if LtenD.state=0 then LstarTens.state=1
	if LtenSwow.state=1 then 
		addwow 1
		wowblack=1
	end if
	Sequences
  end if
end sub

sub TGjackS_hit 
  if not tilt then
	addscore 500
	LjackS.state=0
	if LjackD.state=0 then LstarJacks.state=1
	if LjackSwow.state=1 then 
		addwow 1
		wowblack=1
	end if
	Sequences
  end if
end sub

sub TGkingS2_hit 
  if not tilt then
	addscore 500
	LkingS.state=0
	if LkingD.state=0 then LstarKings.state=1
	if LkingS2wow.state=1 then 
		addwow 1
		wowblack=1
	end if
	Sequences
  end if
end sub


sub TGaceS_hit 
  if not tilt then
	addscore 500
	LaceS.state=0
	if LaceD.state=0 then LstarAces.state=1
	if LaceSwow.state=1 then 
		addwow 1
		wowblack=1
	end if
	Sequences
  end if
end sub

sub TGstarNines_hit
  if not tilt then
	if LstarNines.state=1 Then
		addscore 1000
	  Else
		addscore 100
	end If
  end if
end Sub

sub TGstarTens_hit
  if not tilt then
	if LstarTens.state=1 Then
		addscore 1000
	  Else
		addscore 100
	end If
  end if
end Sub

sub TGstarJacks_hit
  if not tilt then
	if LstarJacks.state=1 Then
		addscore 1000
	  Else
		addscore 100
	end If
  end if
end Sub

sub TGstarQueens_hit
  if not tilt then
	if LstarQueens.state=1 Then
		addscore 1000
	  Else
		addscore 100
	end If
  end if
end Sub

sub TGstarKings_hit
  if not tilt then
	if LstarKings.state=1 Then
		addscore 1000
	  Else
		addscore 100
	end If
  end if
end Sub

sub TGstarAces_hit
  if not tilt then
	if LstarAces.state=1 Then
		addscore 1000
	  Else
		addscore 100
	end If
  end if
end Sub

dim balltokick

sub TGmystery_hit
	set balltokick = activeball
	if Lmystery.state=1 then 
		addwow mysterycnt
	  else
		me.uservalue=25-(mysterycnt*8)
		me.timerenabled=1
	end if
end Sub

sub TGmystery_timer
	select case me.uservalue
	  case 1:
		addscore 5000
	  case 9:
	    addscore 5000
	  case 17:
		addscore 5000
	  case 26:
		PlaySoundAtVol SoundFXDOF("holekick",115,DOFPulse,DOFContactors), TGmystery, 1
		MysteryPlunger.fire
		slingtop.rotx=20
	  case 28:
		slingtop.rotx=0
		me.timerenabled=0
	end Select
	me.uservalue=me.uservalue+1
End Sub



'**********  Targets

sub TkingS_hit 
  if Not tilt then
	DOF 118, DOFPulse
	addscore 500
	LkingS.state=0
	if LkingD.state=0 then LstarKings.state=1
	if LkingSwow.state=1 then 
		addwow 1
		wowblack=1
	end if
	Sequences
  end if
end sub

sub TqueenS_hit 
  if Not tilt then
	DOF 118, DOFPulse
	addscore 500
	LqueenS.state=0
	if LqueenD.state=0 then LstarQueens.state=1
	if LqueenSwow.state=1 then 
		addwow 1
		wowblack=1
	end if
	Sequences
  end if
end sub

'**********  Kickers

Sub KickSpotsRight_Hit()
  if not tilt then
	PlaySoundAt "cluper", KickSpotsRight
	if LspotsR1.state = 1 then 
		spotred
	  elseif LspotsB1.state = 1 then 
		spotblack
	end if
	if LspotsRight.state = 1 Then
		addwow 1
	  Else
		addscore 5000
	end if
  end if
	me.uservalue=1
	me.timerenabled=1

End Sub

Sub KickSpotsRight_timer
	select case me.uservalue
	  case 4:
		PlaySoundAtVol SoundFXDOF("holekick",117,DOFPulse,DOFContactors), KickSpotsRight, 1
		DOF 139, DOFPulse
		KickSpotsRight.kick 230,10,0
		PkickarmRight.rotz=15
	  case 6:
		PkickarmRight.rotz=0
		me.timerenabled=0
	end Select
	Me.uservalue=me.uservalue+1
End Sub	

Sub KickSpotsTop_Hit()
  if not tilt then
	PlaySoundAt "cluper", KickSpotsTop
	if LspotsR.state = 1 then 
		spotred
	  elseif LspotsB.state = 1 then 
		spotblack
	end if
	if LspotsTop.state = 1 Then
		addwow 1
	  Else
		addscore 5000
	end if
  end if
	me.uservalue=1
	me.timerenabled=1

End Sub

Sub KickSpotsTop_timer
	select case me.uservalue
	  case 4:
		PlaySoundAtVol SoundFXDOF("holekick",116,DOFPulse,DOFContactors), KickSpotsTop, 1
		DOF 139, DOFPulse
		KickSpotsTop.kick 220,10,0
		PkickarmTop.rotz=15
	  case 6:
		PkickarmTop.rotz=0
		me.timerenabled=0
	end Select
	Me.uservalue=me.uservalue+1
End Sub	


sub spotred
	cardLightsRed(rbspot).state=0
end sub

sub spotblack
	cardLightsBlack(rbspot).state=0
end sub

'********** Sequences and WOWS

Sub Sequences
	if LnineD.state+LtenD.state+LjackD.state+LqueenD.state+LkingD.state+LaceD.state=0 then
		redsequence=1
		for each light in wowLightsRed: light.state=0:next
		wowLightsRed(wowspot).state=1
	end If
	if LnineS.state+LtenS.state+LjackS.state+LqueenS.state+LkingS.state+LaceS.state=0 then
		blacksequence=1
		for each light in wowLightsBlack: light.state=0:next
		wowLightsBlack(wowspot).state=1
	end If
	if redsequence=1 and blacksequence=1 Then Lmystery.state=1
end sub

sub movewow
	rbspot=rbspot+1
	if rbspot>5 then rbspot=0
	wowspot=wowspot+1
	if wowspot>7 then wowspot=0
	if redsequence=1 then 
		for each light in wowLightsRed: light.state=0:next
		wowLightsRed(wowspot).state=1
	end if
	if blacksequence=1 then 
		for each light in wowLightsBlack: light.state=0:next
		wowLightsBlack(wowspot).state=1
	end if
end sub

sub addwow(numwows)
	if ballinplay<balls then
		if ballinplay+numwows<balls then	'*********add in section to count up wows with solenoid kick
			ballinplay=ballinplay+numwows
		  else
			numwows=numwows-(balls-ballinplay)
			ballinplay=Balls
			wows=wows+numwows
		end if
	  else
		wows=wows+numwows
		if wows>5 then wows=5
	end if

	If ballinplay+wows<=balls then
		wowsreel.setvalue 0
		bipreel.setvalue ballinplay
		If B2SOn then Controller.B2ssetballinplay 32, Ballinplay
	  else
		bipreel.setvalue Balls
		if B2SOn then Controller.B2ssetballinplay 32, Balls
		wowsreel.setvalue wows
		if B2SOn then 
			for i= 1 to wows
				Controller.B2ssetdata 10+i, 1
			next
		end if
	end if

end Sub

'********** Scoring

sub addscore(points)
  if tilt=false and state=true then
	If points = 10 then 
		matchnumb=matchnumb+1
		if matchnumb>9 then matchnumb=0
		mysterycnt=mysterycnt+1
		if mysterycnt>3 then mysterycnt=1
	end if
	if points=10 or points=100 or points=1000 then 
		addpoints Points
	  else
		FlashBumpers
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
        PlaySoundAtVol SoundFXDOF("bell1000",143,DOFPulse,DOFChimes), chimesound, 1
      ElseIf Points = 10 AND(Score(player) MOD 100) \ 10 = 0 Then 'New 100 reel
        PlaySoundAtVol SoundFXDOF("bell100",142,DOFPulse,DOFChimes), chimesound, 1
      ElseIf points = 1000 Then
        PlaySoundAtVol SoundFXDOF("bell1000",143,DOFPulse,DOFChimes), chimesound, 1
	  elseif Points = 100 Then
        PlaySoundAtVol SoundFXDOF("bell100",142,DOFPulse,DOFChimes), chimesound, 1
      Else
        PlaySoundAtVol SoundFXDOF("bell10",141,DOFPulse,DOFChimes), chimesound, 1
    End If
    ' check replays and rollover
	if score(player)=>99999 then
		If B2SOn then Controller.B2SSetScoreRollover 24 + player, int(score(player)/100000)
		EVAL("Reel100K"&player).setvalue(int(score(player)/100000))
	End if
    if score(player)=>replay1 and rep(player)=0 then
		addwow 1
		rep(player)=1
		PlaySoundAtVol SoundFXDOF("knock",138,DOFPulse,DOFKnocker), Plunger, 1
		DOF 139, DOFPulse
    end if
    if score(player)=>replay2 and rep(player)=1 then
		addwow 1
		rep(player)=2
		PlaySoundAtVol SoundFXDOF("knock",138,DOFPulse,DOFKnocker), Plunger, 1
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

Sub GameTilted
	Tilt = True
	tilttxt.text="TILT"
    If B2SOn Then Controller.B2SSetTilt 33,1
    If B2SOn Then Controller.B2ssetdata 1, 0
	playsoundat "tilt", soundtrigger
	turnoff
End Sub

sub turnoff
	For each light in GIlights:light.state=0:Next
	If B2SOn then Controller.B2SSetData 100,0
	for i=1 to 2
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

Function AudioFade(tableobj) ' Fades between front and back of the table (for surround systems or 2x2 speakers, etc), depending on the Y position on the table. "Neptune" is the name of the table
	Dim tmp
    tmp = tableobj.y * 2 / Neptune.height-1
    If tmp > 0 Then
		AudioFade = Csng(tmp ^10)
    Else
        AudioFade = Csng(-((- tmp) ^10) )
    End If
End Function

Function AudioPan(tableobj) ' Calculates the pan for a tableobj based on the X position on the table. "Neptune" is the name of the table
    Dim tmp
    tmp = tableobj.x * 2 / Neptune.width-1
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
        If BOT(b).X < Neptune.Width/2 Then
            BallShadow(b).X = ((BOT(b).X) - (Ballsize/16) + ((BOT(b).X - (Neptune.Width/2))/17))' + 13
        Else
            BallShadow(b).X = ((BOT(b).X) + (Ballsize/16) + ((BOT(b).X - (Neptune.Width/2))/17))' - 13
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
    savevalue "Neptune", "credit", credit
    savevalue "Neptune", "hiscore", hisc
    savevalue "Neptune", "match", matchnumb
    savevalue "Neptune", "score1", score(1)
    savevalue "Neptune", "score2", score(2)
    savevalue "Neptune", "score3", score(3)
    savevalue "Neptune", "score4", score(4)
	savevalue "Neptune", "balls", balls
	savevalue "Neptune", "hsa1", HSA1
	savevalue "Neptune", "hsa2", HSA2
	savevalue "Neptune", "hsa3", HSA3
	savevalue "Neptune", "freeplay", freeplay

end sub

sub loadhs
    dim temp
	temp = LoadValue("Neptune", "credit")
    If (temp <> "") then credit = CDbl(temp)
    temp = LoadValue("Neptune", "hiscore")
    If (temp <> "") then hisc = CDbl(temp)
    temp = LoadValue("Neptune", "match")
    If (temp <> "") then matchnumb = CDbl(temp)
    temp = LoadValue("Neptune", "score1")
    If (temp <> "") then score(1) = CDbl(temp)
    temp = LoadValue("Neptune", "score2")
    If (temp <> "") then score(2) = CDbl(temp)
    temp = LoadValue("Neptune", "score3")
    If (temp <> "") then score(3) = CDbl(temp)
    temp = LoadValue("Neptune", "score4")
    If (temp <> "") then score(4) = CDbl(temp)
    temp = LoadValue("Neptune", "balls")
    If (temp <> "") then balls = CDbl(temp)
    temp = LoadValue("Neptune", "hsa1")
    If (temp <> "") then HSA1 = CDbl(temp)
    temp = LoadValue("Neptune", "hsa2")
    If (temp <> "") then HSA2 = CDbl(temp)
    temp = LoadValue("Neptune", "hsa3")
    If (temp <> "") then HSA3 = CDbl(temp)
    temp = LoadValue("Neptune", "freeplay")
    If (temp <> "") then freeplay = CDbl(temp)
end sub

Sub Neptune_Exit()
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


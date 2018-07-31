'
'         _, __, __,  _,    _,  _,   __,  _, __, __,  _,
'        /_\ |_) |_) /_\   / ` /_\   | \ /_\ |_) |_) /_\
'        | | |_) | \ | |   \ , | |   |_/ | | |_) | \ | |
'
'                Abra_Ca_Dabra by Gottlieb (1975)
'
'	-------------------------------------------------
'	Gottlieb EM 4 player VPX table blank with options menu
'	-------------------------------------------------
'
'	BorgDog, 2017
'
'	High Score sticky routines from mfuegemann's Fast Draw VPX table
'		- flippers to change letter, Start to Select
'	Ball control script from rothbauerw
'		- press C during play to control ball, use arrows to move the ball
'	Option menu and player up light rotation borrowed from loserman76 and gnance
'		- hold down left shift during Game Over to bring up menu
'	Ball shadows from ninuzzu
'		- set option below to enable or disable
'	primitives from Dark, zany, sliderpoint, hauntfreaks and I'm sure others
'
'	Layer usage generally
'		1 - most stuff
'		2 - triggers
'		3 - under apron walls and ramps
'		4 - options menu
'		5 - hi score sticky
'		6 - GI lighting
'		7 - plastics
'		8 - insert and bumper lighting
'
' Thalamus 2018-07-24
' Added/Updated "Positional Sound Playback Functions" and "Supporting Ball & Sound Functions"
' No special SSF tweaks yet.

'
'	-------------------------------------------------

Option Explicit
Randomize

Const cGameName = "Abracadabra"
Const Ballsize = 50
Const BallMass = 1.3

On Error Resume Next
ExecuteGlobal GetTextFile("controller.vbs")
If Err Then MsgBox "You need the Controller.vbs file in order to run this table (installed with the VPX package in the scripts folder)"
On Error Goto 0

Dim operatormenu, options
Dim bumperlitscore
Dim bumperoffscore
Dim balls, TempPlayerUp, rotatortemp
Dim replays, freeplay
Dim Replay1Table(3), Replay2Table(3), Replay3Table(3)
dim replay1, replay2, replay3
Dim Add10, Add100, Add1000
Dim hisc
Dim maxplayers, players, player
Dim credit, DTcount, creditnum, credithole
Dim score(4), bonus
Dim state
Dim tilt, tiltsens
Dim ballinplay, PlungeBall
Dim matchnumb
dim ballrenabled
dim rstep, lstep
Dim rep(4)
Dim rst
Dim eg
Dim bell
Dim i,j, ii, objekt, light


Dim DesktopMode: DesktopMode = Abra_Ca_Dabra.ShowDT

If DesktopMode = True Then 'Show Desktop components (in FS Mode, DesktopMode = False)
		for each objekt in backdropstuff : objekt.visible = 1 : next
	Else
		for each objekt in backdropstuff : objekt.visible = 0 : next
End if



'**********************************************************
'********   	OPTIONS		*******************************
'**********************************************************

Dim BallShadows: Ballshadows=1  		'******************set to 1 to turn on Ball shadows
Dim FlipperShadows: FlipperShadows=1  '***********set to 1 to turn on Flipper shadows

sub Abra_Ca_Dabra_init
	LoadEM
	maxplayers=1
	Replay1Table(1)=55000
	Replay2Table(1)=69000
	Replay3Table(1)=77000
	Replay1Table(2)=65000
	Replay2Table(2)=79000
	Replay3Table(2)=87000
	Replay1Table(3)=75000
	Replay2Table(3)=89000
	Replay3Table(3)=97000

	hideoptions
	player=1
	RotatorTemp=1
	balls=5   		'SET DEFAULT VALUES PRIOR TO LOADING FROM SAVED FILE
	replays=1
	freeplay=0
	credit=0
	credithole=0
	creditnum=0
	hisc=10000
	HSA1=4
	HSA2=15
	HSA3=7
	matchnumb=1
	DTcount=5
	loadhs		'LOADS SAVED OPTIONS AND HIGH SCORE IF EXISTING
	UpdatePostIt	'UPDATE HIGH SCORE STICKY
	for i = 1 to maxplayers
		EVAL("ScoreReel"&i).setvalue score(i)
		EVAL("Reel100K"&i).setvalue(int(score(i)/100000))
	next
	bipreel.setvalue 0
	Creditreel.setvalue(credit)
	CreditHoleReel.setvalue(credithole)
	CreditNumReel.setvalue(creditnum)
	Replay1=Replay1Table(Replays)
	Replay2=Replay2Table(Replays)
	Replay3=Replay3Table(Replays)
	RepCard.image = "ReplayCard"&replays
	OptionBalls.image="OptionsBalls"&Balls
	OptionReplays.image="OptionsReplays"&replays
	OptionFreeplay.image="OptionsFreeplay"&freeplay
	bumperlitscore=1000
	bumperoffscore=100
	if balls=3 then
		InstCard.image="InstCard3balls"
	  else
		InstCard.image="InstCard5balls"
	end if
	If B2SOn then
		showb2scredit
		for each objekt in backdropstuff : objekt.visible = 0 : next
	End If
	startGame.enabled=true
	tilt=false
	turnoff
	If credit>0 then DOF 125, DOFOn

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

    Drain.CreateSizedballWithMass Ballsize/2,Ballmass
End sub

sub startGame_timer
	PlaySoundAt "poweron", Plunger
	lightdelay.enabled=true
	me.enabled=false
end sub

sub lightdelay_timer
	gamov.text="Game Over"
	if matchnumb=10 then
		matchtxt.text="00"
	  else
		matchtxt.text=matchnumb*10
	end if
	For each light in GIlights:light.state=1:Next

	If (DTcount Mod 2) = 1 then
		for each light in BumperLights: light.state=0:Next
	  else
		for each light in BumperLights: light.state=1:Next
	end if

	if B2SOn then
		for i = 1 to maxplayers
			Controller.B2SSetScorePlayer i, Score(i) MOD 100000
		next
		showb2scredit
		Controller.B2ssetMatch 34, matchnumb
		Controller.B2SSetGameOver 35,1
		Controller.B2SSetScorePlayer 5, hisc
	end if
	me.enabled=0
end Sub



Sub Abra_Ca_Dabra_KeyDown(ByVal keycode)

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
		if freeplay = 1 or credit=9 then
			PlaySoundAt "coinreturn", Drain
		  else
			PlaySoundAt "coinin" , Drain
			coindelay.enabled=true
		end if
    end if

    if keycode=StartGameKey  and (Credit>0 or freeplay=1) And Not HSEnterMode=true then
	  if state=false then
		if freeplay=0 then
			credit=credit-1
			creditreel.setvalue(credit)
			creditnum=creditnum+1
			if creditnum>9 then creditnum=0
			creditnumreel.setvalue(creditnum)
		end if
		If credit < 1 Then DOF 125, DOFOff
		playsoundat "click", Drain

		ballinplay=1
		If B2SOn Then
			showb2scredit
			Controller.B2ssetballinplay 32, Ballinplay
			Controller.B2ssetplayerup 30, 1
			Controller.B2ssetcanplay 31, 1
			Controller.B2SSetGameOver 0
			Controller.B2SSetScorePlayer 5, hisc
		End If
		for i = 1 to maxplayers
			score(i)=0
			rep(i)=0
			EVAL("Pup"&i).state=0
			EVAL("PupN"&i).state=0
			EVAL("ScoreReel"&i).resettozero
			EVAL("Reel100K"&i).setvalue 0
			If B2SOn then Controller.B2SSetScorePlayer i, score(i)
		next
	    Pup1.state=1
		PupN1.state=1
		PlaySound("RotateThruPlayers"),0,.05,0,0.25
		TempPlayerUp=Player
		PlayerUpRotator.enabled=true
		tilt=false
		state=true
		PlaySoundAt "initialize", soundtrigger
		players=1
		canplayreel.setvalue(players)
		rst=0
		newgame.enabled=true
	  else if state=true and players < maxplayers and Ballinplay=1 then
		if freeplay=0 then
			credit=credit-1
			if credit < 1 then DOF 125, DOFOff
			players=players+1
			Creditreel.setvalue(credit)
			creditnum=creditnum+1
			if creditnum>9 then creditnum=0
			creditnumreel.setvalue(creditnum)
		End if
		players=players+1
		canplayreel.setvalue(players)
		Creditreel.setvalue(credit)
		If B2SOn then
			showb2scredit
			Controller.B2ssetcanplay 31, players
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
		PlaySoundAt SoundFXDOF("fx_flipperup",101,DOFOn,DOFContactors), LeftFlipper
		PlaySound "Buzz", -1, .01, AudioPan(LeftFlipper), .05,0, 0, 1, AudioFade(LeftFlipper)
	End If

	If keycode = RightFlipperKey Then
		RightFlipper.RotateToEnd
		PlaySoundAt SoundFXDOF("fx_flipperup",102,DOFOn,DOFContactors), RightFlipper
		PlaySound "Buzz1", -1, .01, AudioPan(RightFlipper), .05,0, 0, 1, AudioFade(RightFlipper)
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


Sub Abra_Ca_Dabra_KeyUp(ByVal keycode)

	if keycode = 203 then Cleft = 0' Left Arrow

	if keycode = 200 then Cup = 0' Up Arrow

	if keycode = 208 then Cdown = 0' Down Arrow

	if keycode = 205 then Cright = 0' Right Arrow

	If keycode = PlungerKey Then
		Plunger.Fire
		if PlungeBall=1 then
			PlaySoundAt "plunger", Plunger	'PLAY WHEN BALL IS HIT
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
		PlaySoundAt SoundFXDOF("fx_flipperdown",101,DOFOff,DOFContactors), LeftFlipper
		StopSound "Buzz"
	End If

	If keycode = RightFlipperKey Then
		RightFlipper.RotateToStart
		PlaySoundAt SoundFXDOF("fx_flipperdown",102,DOFOff,DOFContactors), RightFlipper
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
	LFlip.RotY = LeftFlipper.CurrentAngle+90
	RFlip.RotY = RightFlipper.CurrentAngle+90
	Pgate.rotz = (Gate.currentangle*.75)+25

	if FlipperShadows=1 then
		FlipperLSh.RotZ = LeftFlipper.currentangle
		FlipperRSh.RotZ = RightFlipper.currentangle
	end if

end sub


Sub PairedlampTimer_timer
	bumperlighta1.state = bumperlight1.state
	bumperlighta2.state = bumperlight2.state
	l10.state=l1.state
	l9.state=l2.state
	l8.state=l3.state
	l7.state=l4.state
	l6.state=l5.state
    LoutL.state=LTop1.state
	LinL.state=LTop2.state
	LinR.state=LTop3.state
	LoutR.state=LTop4.state
	l25a.state=l25.state
end sub

sub coindelay_timer
	addcredit
    coindelay.enabled=false
end sub

Sub addcredit
  if freeplay=0 then
      credit=credit+1
      if credit>9 then
			credit=9
		Else
			credithole=credithole+1
			if credithole>9 then credithole=0
			creditholereel.setvalue(credithole)
	  end if
	  DOF 125, DOFOn

	  Creditreel.setvalue(credit)

	  If B2SOn Then showb2scredit
  end if
End sub

sub showb2scredit
	for i = 1 to 10
		Controller.B2ssetData i, 0
	next
	if Credit=0 then
		Controller.B2ssetData 10, 1
	  else
		Controller.B2ssetData Credit, 1
	end if
end sub

Sub Drain_Hit()
'	DOF 134, DOFPulse
	PlaySoundAt "drain", Drain
	me.timerenabled=1
End Sub

Sub Drain_timer
'	scorebonus.enabled=true
	nextball
	me.timerenabled=0
End Sub

sub ballhome_hit
	ballrenabled=1
	plungeball=1
end sub

sub ballhome_unhit
'	DOF 136, DOFPulse
	plungeball=0
end sub

sub ballrel_hit
	if ballrenabled=1 then
		ballrenabled=0
	end if
end sub

sub scorebonus			'************add bonus scoring if applicable

	addscore (bonus*1000)

End sub


sub newgame_timer
    eg=0
	For each light in GIlights:light.state=1:next
	for each light in Lanes_lights:light.state=1: next
	for each light in Bonus_lights:light.state=0: next
	l25.state=0
	dt1.timerenabled=1
	l1000.state=1
	bonus=1
	DTlights(DTcount-1).state=1
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
	bipreel.setvalue 1
    matchtxt.text=" "
	newball
	ballrelTimer.enabled=true
	newgame.enabled=false
end sub


sub newball
'	for each light in bumperlights:light.state=1:next
'	for each light in Lanes_lights:light.state=1:next
End Sub


sub nextball
    if tilt=true then
	  for i=1 to 2
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
	if player=1 then ballinplay=ballinplay+1
	if ballinplay>balls then
		playsound "GameOver"
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
	  canplayreel.setvalue(0)
	  for i=1 to maxplayers
		if score(i)>hisc then
			hisc=score(i)
			hiscstate=1
		end if
		EVAL("Pup"&i).state=0
		EVAL("PupN"&i).state=0
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
	    Controller.B2ssetcanplay 31, 0
	    Controller.B2ssetcanplay 30, 0
	  End If
	  For each light in GIlights:light.state=0:next
	  ballreltimer.enabled=false
  else
	Drain.kick 60,45,0
    ballreltimer.enabled=false
	playsoundat SoundFXDOF("kickerkick",141,DOFPulse,DOFContactors), Drain
  end if
end sub

Sub HStimer_timer
	playsoundat SoundFXDOF("knock",140,DOFPulse,DOFKnocker), Plunger

	HStimer.uservalue=HStimer.uservalue+1
	if HStimer.uservalue=3 then me.enabled=0
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
			If B2SOn Then Controller.B2SSetPlayerUp TempPlayerUp
		else
			if B2SOn then Controller.B2SSetPlayerUp Player
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
	if matchnumb=10 then
		matchtxt.text="00"
		If B2SOn then Controller.B2SSetMatch 100
	  else
		matchtxt.text=matchnumb*10
		If B2SOn then Controller.B2SSetMatch 34,Matchnumb
	end if
	For i=1 to players
		if (matchnumb*10)=(score(i) mod 100) then
		  addcredit
		  PlaySoundAt SoundFXDOF("knock",140,DOFPulse,DOFKnocker), Plunger
	    end if
    next
end sub

'********** Bumpers

Sub Bumper1_Hit
   if tilt=false then
	PlaySoundAt SoundFXDOF("fx_bumper4",104,DOFPulse,DOFContactors), Bumper1

	if bumperlight1.state=1 then
		addscore bumperlitscore
	  else
		addscore bumperoffscore
	end if
   end if
End Sub


Sub Bumper2_Hit
   if tilt=false then
	PlaySoundAt SoundFXDOF("fx_bumper4",105,DOFPulse,DOFContactors), Bumper2

	if BumperLight2.state=1 then
		addscore bumperlitscore
	  else
		addscore bumperoffscore
	end if
   end if
End Sub

sub ChangeBumperLights
	if BumperLight1.state=1 then
		for each light in BumperLights: light.state=0:Next
	  else
		for each light in BumperLights: light.state=1:Next
	end if
end sub

sub FlashBumpers
	if bumperlight1.state = 1 then
		for each light in BumperLights: light.duration 0, 300, 1:Next
	end If
end sub

'************** Slings

Sub RightSlingShot_Slingshot
	PlaySoundAt SoundFXDOF("right_slingshot",127,DOFPulse,DOFContactors), slingR

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
	PlaySoundAt SoundFXDOF("left_slingshot",126,DOFPulse,DOFContactors), slingL

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
	lSling.visible=0
	SlingC1.visible=1
	me.uservalue=1
	Me.timerenabled=1
end sub

sub dingwallc_timer									'default 50 timer
	select case me.uservalue
		Case 1: Slingc1.visible=0: lSling.visible=1
		case 2:	lsling.visible=0: Slingc2.visible=1
		Case 3: Slingc2.visible=0: lsling.visible=1: Me.timerenabled=0
	end Select
	me.uservalue=me.uservalue+1
end sub

sub dingwalld_hit
	if state=true and tilt=false then addscore 10
	rSling.visible=0
	SlingD1.visible=1
	me.uservalue=1
	Me.timerenabled=1
end sub

sub dingwalld_timer									'default 50 timer
	select case me.uservalue
		Case 1: Slingd1.visible=0: rSling.visible=1
		case 2:	rsling.visible=0: Slingd2.visible=1
		Case 3: Slingd2.visible=0: rsling.visible=1: Me.timerenabled=0
	end Select
	me.uservalue=me.uservalue+1
end sub

'********** Triggers

sub TGBonusL_hit
	DOF 145, DOFPulse
	scorebonus
'  addscore 500
end sub

sub TGoutL_hit	'****** bot 1 rollover
	DOF 146, DOFPulse
	addscore 500
	LTop1.state=0
	awardcheck
end sub

sub TGinL_hit	'****** bot 2 rollover
	DOF 146, DOFPulse
	addscore 500
    Ltop2.state=0
	If balls=3 then LTop3.state =0
	awardcheck
end sub

sub TGinR_hit	'****** bot 3 rollover
	DOF 147, DOFPulse
	addscore 500
    Ltop3.state=0
	If balls=3 then Ltop2.state =0
	awardcheck
end sub

sub TGoutR_hit	'****** bot 4 rollover
	DOF 147, DOFPulse
	addscore 500
    Ltop4.state=0
	awardcheck
end sub

sub TGBonusR_hit
	DOF 148, DOFPulse
	scorebonus
'	addscore 500
end sub

sub TGtop1_hit   '***** top 1 rollover
	DOF 135, DOFPulse
	addscore 500
    LTop1.state=0
	awardcheck
end sub

sub TGtop2_hit   '***** top 2 rollover
	DOF 136, DOFPulse
	addscore 500
    LTop2.state=0
	If balls=3 then LTop3.state =0
	awardcheck
end sub

sub TGtop3_hit   '***** top 3 rollover
	DOF 137, DOFPulse
	addscore 500
    LTop3.state=0
	If balls=3 then Ltop2.state =0
	awardcheck
end sub

sub TGtop4_hit   '***** top 4 rollover
	DOF 138, DOFPulse
	addscore 500
    LTop4.state=0
	awardcheck
end sub

sub awardcheck
	if Ltop1.state+ltop2.state+ltop3.state+ltop4.state=0 then
		if l5000.state=1 then addspecial
		ScoreBonus
		addbonus
		for each light in lanes_lights: light.state=1: next
	end if
end sub

'**********  Targets

Sub Target_hit
	PlaySoundAt SoundFXDOF("target",103,DOFPulse,DOFTargets), Target
	if l25.state=1 then
		addbonus
		if balls=5 then
			dt1.timerenabled=1
		  else
			if DT1.isdropped and DT2.isdropped and DT3.isdropped and DT4.isdropped and DT5.isdropped then dt2.timerenabled=1
			if DT6.isdropped and DT7.isdropped and DT8.isdropped and DT9.isdropped and DT10.isdropped then dt6.timerenabled=1
		end if
	  else
		addscore 500
	end if
	if lspecial.state=1 Then
		addspecial
		Lspecial.state=0
	end If
end sub

Sub DT1_timer
	PlaySoundAt SoundFXDOF("DTreset",142,DOFPulse,DOFContactors), DT1
	for i = 1 to 5
		EVAL("DT"&i).isdropped=False
	next
	DT6.timerenabled=1
	DT1.timerenabled=0
End Sub

Sub DT2_timer
	PlaySoundAt SoundFXDOF("DTreset",142,DOFPulse,DOFContactors), DT1
	for i = 1 to 5
		EVAL("DT"&i).isdropped=False
	next
	DT2.timerenabled=0
	L25.state=0
End Sub

Sub DT6_timer
	PlaySoundAt SoundFXDOF("DTreset",143,DOFPulse,DOFContactors), DT6
	for i = 6 to 10
		EVAL("DT"&i).isdropped=False
	next
	DT6.timerenabled=0
	L25.state=0
End sub

Sub addspecial
	if freeplay=0 then addcredit
	PlaySoundAt SoundFXDOF("knock",140,DOFPulse,DOFKnocker), Plunger
end sub

'**********  Drop Targets

Sub DToff
	for each objekt in DTlights: objekt.state=0: next
end sub

Sub MoveDTlights
	DToff
	DTcount=DTcount-1
	if DTcount<1 then DTcount=5
	DTlights(DTcount-1).state=1
end sub

Sub DT1_dropped
	PlaySoundAt SoundFXDOF("drop1",111,DOFPulse,DOFDropTargets), DT1
'	PlaySoundAt "drop1", DT1
	if l1.state=1 then
		scorebonus
	  else
		addscore 500
	end if
	DTcheck
end sub

Sub DT2_dropped
	PlaySoundAt SoundFXDOF("drop1",111,DOFPulse,DOFDropTargets), DT2
'	PlaySoundAt "drop1", DT2
	if l2.state=1 then
		scorebonus
	  else
		addscore 500
	end if
	DTcheck
end sub

Sub DT3_dropped
	PlaySoundAt SoundFXDOF("drop1",111,DOFPulse,DOFDropTargets), DT3
'	PlaySoundAt "drop1", DT3
	if l3.state=1 then
		scorebonus
	  else
		addscore 500
	end if
	DTcheck
end sub

Sub DT4_dropped
	PlaySoundAt SoundFXDOF("drop1",111,DOFPulse,DOFDropTargets), DT4
'	PlaySoundAt "drop1", DT4
	if l4.state=1 then
		scorebonus
	  else
		addscore 500
	end if
	DTcheck
end sub

Sub DT5_dropped
	PlaySoundAt SoundFXDOF("drop1",111,DOFPulse,DOFDropTargets), DT5
'	PlaySoundAt "drop1", DT5
	if l5.state=1 then
		scorebonus
	  else
		addscore 500
	end if
	DTcheck
end sub

Sub DT6_dropped
	PlaySoundAt SoundFXDOF("drop1",116,DOFPulse,DOFDropTargets), DT6
'	PlaySoundAt "drop1", DT6
	if l6.state=1 then
		scorebonus
	  else
		addscore 500
	end if
	DTcheck
end sub

Sub DT7_dropped
	PlaySoundAt SoundFXDOF("drop1",116,DOFPulse,DOFDropTargets), DT7
'	PlaySoundAt "drop1", DT7
	if l7.state=1 then
		scorebonus
	  else
		addscore 500
	end if
	DTcheck
end sub

Sub DT8_dropped
	PlaySoundAt SoundFXDOF("drop1",116,DOFPulse,DOFDropTargets), DT8
'	PlaySoundAt "drop1", DT8
	if l8.state=1 then
		scorebonus
	  else
		addscore 500
	end if
	DTcheck
end sub

Sub DT9_dropped
	PlaySoundAt SoundFXDOF("drop1",116,DOFPulse,DOFDropTargets), DT9
'	PlaySoundAt "drop1", DT9
	if l9.state=1 then
		scorebonus
	  else
		addscore 500
	end if
	DTcheck
end sub

Sub DT10_dropped
	PlaySoundAt SoundFXDOF("drop1",116,DOFPulse,DOFDropTargets), DT10
'	PlaySoundAt "drop1", DT10
	if l10.state=1 then
		scorebonus
	  else
		addscore 500
	end if
	DTcheck
end sub

Sub DTcheck
	if balls=5 then
		if DT1.isdropped and DT2.isdropped and DT3.isdropped and DT4.isdropped and DT5.isdropped and DT6.isdropped and DT7.isdropped and DT8.isdropped and DT9.isdropped and DT10.isdropped then
			L25.state=1
			if l5000.state=1 then lspecial.state=1
		end if
	  else
		if (DT1.isdropped and DT2.isdropped and DT3.isdropped and DT4.isdropped and DT5.isdropped) or (DT6.isdropped and DT7.isdropped and DT8.isdropped and DT9.isdropped and DT10.isdropped) then
			L25.state=1
			if l5000.state=1 then lspecial.state=1
		end if
	end if
end sub



'********** Scoring

sub addbonus
	bonus=bonus+1
	if bonus>5 then bonus=5
	Bonus_lights(bonus-1).state=1
	if bonus>1 then
		Bonus_lights(bonus-2).state=0
	  else
		l1000.state=0
	end if
end sub

sub addscore(points)
  if tilt=false and state=true then

	If points = 10 then
		matchnumb=matchnumb+1
		if matchnumb>10 then matchnumb=1
		if balls=5 then
			MoveDTlights
			ChangeBumperLights
		end if
	end if
	If points = 100 then
		if balls=3 then
			MoveDTlights
			ChangeBumperLights
		end if
	end if

	If points = 500 then flashbumpers

	If points = 1000 then
		MoveDTlights
		ChangeBumperLights
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
	EVAL("ScoreReel1").addvalue(points)
	If B2SOn Then Controller.B2SSetScorePlayer player, score(player) MOD 100000

    ' Sounds: there are 3 sounds: tens, hundreds and thousands
    If Points = 100 AND(Score(player) MOD 1000) \ 100 = 0 Then  'New 1000 reel
        PlaySoundAt SoundFXDOF("bell1000",153,DOFPulse,DOFChimes), chimesound
      ElseIf Points = 10 AND(Score(player) MOD 100) \ 10 = 0 Then 'New 100 reel
        PlaySoundAt SoundFXDOF("bell100",152,DOFPulse,DOFChimes), chimesound
      ElseIf points = 1000 Then
        PlaySoundAt SoundFXDOF("bell1000",153,DOFPulse,DOFChimes), chimesound
	  elseif Points = 100 Then
        PlaySoundAt SoundFXDOF("bell100",152,DOFPulse,DOFChimes), chimesound
      Else
        PlaySoundAt SoundFXDOF("bell10",151,DOFPulse,DOFChimes), chimesound
    End If
    ' check replays and rollover
	if score(player)=>99999 then
		If B2SOn then Controller.B2SSetScoreRollover 24 + player, 1
        EVAL("Reel100K"&player).setvalue(int(score(player)/100000))
	End if
    if score(player)=>replay1 and rep(player)=0 then
		addcredit
		rep(player)=1
		PlaySoundAt SoundFXDOF("knock",140,DOFPulse,DOFKnocker), Plunger

    end if
    if score(player)=>replay2 and rep(player)=1 then
		addcredit
		rep(player)=2
		PlaySoundAt SoundFXDOF("knock",140,DOFPulse,DOFKnocker), Plunger

    end if
    if score(player)=>replay3 and rep(player)=2 then
		addcredit
		rep(player)=3
		PlaySoundAt SoundFXDOF("knock",140,DOFPulse,DOFKnocker), Plunger

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


'*********************************************************************
'                     Supporting Ball & Sound Functions
'*********************************************************************

Function AudioFade(tableobj) ' Fades between front and back of the table (for surround systems or 2x2 speakers, etc), depending on the Y position on the table. "Abra_Ca_Dabra" is the name of the table
	Dim tmp
    tmp = tableobj.y * 2 / Abra_Ca_Dabra.height-1
    If tmp > 0 Then
		AudioFade = Csng(tmp ^10)
    Else
        AudioFade = Csng(-((- tmp) ^10) )
    End If
End Function

Function AudioPan(tableobj) ' Calculates the pan for a tableobj based on the X position on the table. "Abra_Ca_Dabra" is the name of the table
    Dim tmp
    tmp = tableobj.x * 2 / Abra_Ca_Dabra.width-1
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
        If BOT(b).X < Abra_Ca_Dabra.Width/2 Then
            BallShadow(b).X = ((BOT(b).X) - (Ballsize/16) + ((BOT(b).X - (Abra_Ca_Dabra.Width/2))/17))' + 13
        Else
            BallShadow(b).X = ((BOT(b).X) + (Ballsize/16) + ((BOT(b).X - (Abra_Ca_Dabra.Width/2))/17))' - 13
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

'Sub a_Pins_Hit (idx)
'	PlaySound "pinhit_low", 0, Vol(ActiveBall), AudioPan(ActiveBall), 0, Pitch(ActiveBall), 0, 0, AudioFade(ActiveBall)
'End Sub

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

'Sub a_Metals2_Hit (idx)
'	PlaySound "metalhit2", 0, Vol(ActiveBall), AudioPan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
'End Sub

Sub a_Gates_Hit (idx)
	PlaySound "gate4", 0, Vol(ActiveBall), AudioPan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
End Sub

'Sub Spinner_Spin
'	PlaySound "fx_spinner", 0, .25, AudioPan(Spinner), 0.25, 0, 0, 1, AudioFade(Spinner)
'End Sub


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
    savevalue "Abra_Ca_Dabra", "credit", credit
    savevalue "Abra_Ca_Dabra", "credithole", credithole
    savevalue "Abra_Ca_Dabra", "creditnum", creditnum
    savevalue "Abra_Ca_Dabra", "hiscore", hisc
    savevalue "Abra_Ca_Dabra", "match", matchnumb
    savevalue "Abra_Ca_Dabra", "score1", score(1)
    savevalue "Abra_Ca_Dabra", "score2", score(2)
    savevalue "Abra_Ca_Dabra", "score3", score(3)
    savevalue "Abra_Ca_Dabra", "score4", score(4)
	savevalue "Abra_Ca_Dabra", "balls", balls
	savevalue "Abra_Ca_Dabra", "DTcount", DTcount
	savevalue "Abra_Ca_Dabra", "hsa1", HSA1
	savevalue "Abra_Ca_Dabra", "hsa2", HSA2
	savevalue "Abra_Ca_Dabra", "hsa3", HSA3
	savevalue "Abra_Ca_Dabra", "freeplay", freeplay

end sub

sub loadhs
    dim temp
	temp = LoadValue("Abra_Ca_Dabra", "credit")
    If (temp <> "") then credit = CDbl(temp)

	temp = LoadValue("Abra_Ca_Dabra", "credithole")
    If (temp <> "") then credithole = CDbl(temp)

	temp = LoadValue("Abra_Ca_Dabra", "creditnum")
    If (temp <> "") then creditnum = CDbl(temp)

    temp = LoadValue("Abra_Ca_Dabra", "hiscore")
    If (temp <> "") then hisc = CDbl(temp)
    temp = LoadValue("Abra_Ca_Dabra", "match")
    If (temp <> "") then matchnumb = CDbl(temp)
    temp = LoadValue("Abra_Ca_Dabra", "score1")
    If (temp <> "") then score(1) = CDbl(temp)
    temp = LoadValue("Abra_Ca_Dabra", "score2")
    If (temp <> "") then score(2) = CDbl(temp)
    temp = LoadValue("Abra_Ca_Dabra", "score3")
    If (temp <> "") then score(3) = CDbl(temp)
    temp = LoadValue("Abra_Ca_Dabra", "score4")
    If (temp <> "") then score(4) = CDbl(temp)
    temp = LoadValue("Abra_Ca_Dabra", "balls")
    If (temp <> "") then balls = CDbl(temp)

    temp = LoadValue("Abra_Ca_Dabra", "DTcount")
    If (temp <> "") then DTcount = CDbl(temp)

    temp = LoadValue("Abra_Ca_Dabra", "hsa1")
    If (temp <> "") then HSA1 = CDbl(temp)
    temp = LoadValue("Abra_Ca_Dabra", "hsa2")
    If (temp <> "") then HSA2 = CDbl(temp)
    temp = LoadValue("Abra_Ca_Dabra", "hsa3")
    If (temp <> "") then HSA3 = CDbl(temp)
    temp = LoadValue("Abra_Ca_Dabra", "freeplay")
    If (temp <> "") then freeplay = CDbl(temp)
end sub

Sub Abra_Ca_Dabra_Exit()
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
Sub Abra_Ca_Dabra_exit()
  Controller.Pause = False
  Controller.Stop
End Sub


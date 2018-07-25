'	-------------------------------------------------
'	JUNGLE PRINCESS by Gottlieb (1977)
'
'	rebuilt from resources and logic in Jungle Queen vp9 table by Starman
'
'	VPX by BorgDog, 2016
'	thanks to hauntfreaks for graphics and lighting tweaks and making the new kicker primitives
'
'	General layer usage
'		1 - most stuff
'		2 - triggers
'		4 - options menu
'		5 - shadow ramp (full cover)
'		6 - GI lighting
'		7 - plastics
'		8 - insert and bumper lighting
'
'	DOF config - checked and approved by arngrim :)
'		101 Left Flipper, 102 Right Flipper,
'		107 Bumper1, 108 Bumper1 flasher, 109 bumper2, 110 bumper2 Flasher, 111 Bumper3, 112 Bumper3 flasher
'		113 Drop targets left, 114 drop target reset left, 115 drop targets right, 116 drop target reset right
'		117 Rollover top A (left), 118 Rollover top B (mid), 119 Rollover top C (right)
'		120 Rollover left side, 121 Rollover right side
'		122 Lout Rollover, 123 Lin Rollover, 125 Rout rollover, 125 Rin rollover
'		132 Upper left kicker, 133 Upper right kicker, 134 Drain, 135 Ball Release
'		136 Shooter Lane/launch ball, 137 credit light, 138 knocker, 139 knocker/kicker Flasher
'		141 Chime1-10s, 142 Chime2-100s, 143 Chime3-1000s
'
'	-------------------------------------------------

' Thalamus 2018-07-23
' Added/Updated "Positional Sound Playback Functions" and "Supporting Ball & Sound Functions"
' No special SSF tweaks yet.

Option Explicit
Randomize

Const cGameName = "jungleprincess_1977"

Dim operatormenu, options
Dim bumperlitscore, bumperoffscore
Dim balls, extraball, special, ebstate, spstate
Dim replays, Replay1Table(3), Replay2Table(3), Replay3Table(3)
dim replay1, replay2, replay3
Dim Add10, Add100, Add1000, bonus, hisc
Dim maxplayers, players, player
Dim credit
Dim score(4)
Dim sreels(4)
Dim p100k(4)
Dim cplay(4)
Dim Pups(4)
Dim state, freeplay
Dim tilt, tiltsens
Dim dw1step, dw2step, dw3step, dw4step, dw5step, dw6step
Dim target(10)
Dim bonuslight(19)
Dim ballinplay
Dim matchnumb
dim ballrenabled
dim rstep, lstep, lkickstep, rkickstep, dtreset
Dim rep(4)
Dim eg
Dim bell
Dim i,j, ii, objekt, light


On Error Resume Next
ExecuteGlobal GetTextFile("controller.vbs")
If Err Then MsgBox "Can't open controller.vbs"
On Error Goto 0

sub JunglePrincess_init
	LoadEM
	maxplayers=2
	Replay1Table(1)=90000
	Replay2Table(1)=120000
	Replay3Table(1)=160000
	Replay1Table(2)=100000
	Replay2Table(2)=130000
	Replay3Table(2)=170000
	Replay1Table(3)=120000
	Replay2Table(3)=150000
	Replay3Table(3)=190000
	set sreels(1) = ScoreReel1
	set sreels(2) = ScoreReel2
	set Pups(1) = Pup1
	set Pups(2) = Pup2
	set cplay(1) = CanPlay1
	set cplay(2) = CanPlay2
	set p100k(1) = P1100k
	set p100k(2) = P2100k
	set target(1)=DTmonkey1
	set target(2)=DTmonkey2
	set target(3)=DTmonkey3
	set target(4)=DTmonkey4
	set target(5)=DTMonkey5
	set target(6)=DTMonkey6
	set target(7)=DTMonkey7
	set target(8)=DTMonkey8
	set target(9)=DTMonkey9
	set target(10)=DTMonkey10
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
	hideoptions
	player=1
	For each light in BumperLights:light.State = 0: Next
	For each light in lights:light.State = 0: Next
	loadhs
	if hisc="" then hisc=90000
	hstxt.text=hisc
	gamov.text="Game Over"
	gamov.timerenabled=1
	tilttxt.timerenabled=1
	credittxt.setvalue(credit)
	if balls="" then balls=5
	if balls<>3 and balls<>5 then balls=5
	if freeplay="" or freeplay<0 or freeplay>1 then freeplay=0
	if replays="" then replays=1
	if replays<>1 and replays<>2 and replays<>3 then replays=1
	Replay1=Replay1Table(Replays)
	Replay2=Replay2Table(Replays)
	Replay3=Replay3Table(Replays)
	RepCard.image = "ReplayCard"&replays
	OptionBalls.image="OptionsBalls"&Balls
	OptionReplays.image="OptionsReplays"&replays
	OptionFreeplay.image="OptionsFreeplay"&freeplay
	if balls=3 then
		InstCard.image="InstCard3balls"
		bumperlitscore=1000
		bumperoffscore=100
	  else
		InstCard.image="InstCard5balls"
		bumperlitscore=100
		bumperoffscore=10
	end if
	If B2SOn then
		setBackglass.enabled=true
		for each objekt in backdropstuff : objekt.visible = 0 : next
	End If
	if matchnumb="" then matchnumb=100
	if matchnumb=0 then
		matchtxt.text="00"
	  else
		matchtxt.text=matchnumb
	end if
	for i = 1 to maxplayers
		sreels(i).setvalue(score(i))
	next
	PlaySound "motor"
	tilt=false
	If credit>0 then DOF 137, DOFOn
    Drain.CreateBall
End sub

sub setBackglass_timer
	Controller.B2ssetCredits Credit
	Controller.B2ssetMatch 34, Matchnumb
    Controller.B2SSetGameOver 35,1
    Controller.B2SSetScorePlayer 5, hisc
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
		If not B2SOn then gamov.text="Game Over"
		If B2SOn then Controller.B2SSetGameOver 35,1
		gamov.timerenabled=1
		gamov.timerinterval= (INT (RND*10)+5)*100
	end if
	me.enabled=0
end sub

sub tilttxt_timer
	if state=false then
		tilttxt.visible=0
		If B2SOn then Controller.B2SSetTilt 33,0
		ttimer.enabled=true
	end if
	tilttxt.timerenabled=0
end sub

sub ttimer_timer
	if state=false then
		If not B2SOn then tilttxt.visible=1
		If B2SOn then Controller.B2SSetTilt 33,1
		tilttxt.timerinterval= (INT (RND*10)+5)*100
		tilttxt.timerenabled=1
	end if
	me.enabled=0
end sub

Sub JunglePrincess_KeyDown(ByVal keycode)

	if keycode=AddCreditKey then
		playsound "coinin"
		coindelay.enabled=true
    end if

    if keycode=StartGameKey and (credit>0 or freeplay=1) and OperatorMenu=0 then
	  if state=false then
		if freeplay=0 then credit=credit-1
		if credit < 1 then DOF 137, DOFOff
		playsound "cluper"
		credittxt.setvalue(credit)
		ballinplay=1
		For each light in BumperLights:light.State = 0: Next
		For each light in RedBumperLights:light.State = 1: Next
		for each light in GIlights:light.state=1:next
		for each light in lights:light.state=1:next
		If B2SOn Then
			Controller.B2ssetCredits Credit
			Controller.B2sStartAnimation "Startup"
			Controller.B2ssetballinplay 32, Ballinplay
			Controller.B2ssetplayerup 30, 1
			Controller.B2ssetcanplay 31, 1
			Controller.B2SSetGameOver 0
			Controller.B2SSetScorePlayer 5, hisc
		End If
	    pups(1).state=1
		tilt=false
		state=true
		playsound "initialize"
		players=1
		cplay(players).state=1
		newgame.enabled=true
	  else if state=true and players < maxplayers and Ballinplay=1 then
		if freeplay=0 then credit=credit-1
		players=players+1
		cplay(players).state=1
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

  if tilt=false and state=true then
	If keycode = LeftFlipperKey Then
		LeftFlipper.RotateToEnd
		LeftFlip1.RotateToEnd
		PlaySound SoundFXDOF("flipperup",101,DOFOn,DOFContactors), 0, .67, -0.05, 0.05
		PlaySound "Buzz",-1,.05,-0.05, 0.05
	End If

	If keycode = RightFlipperKey Then
		RightFlipper.RotateToEnd
		RightFlip1.RotateToEnd
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
	OptionFreeplay.visible = True
End Sub

Sub HideOptions
	for each objekt In OptionMenu
		objekt.visible = false
	next
End Sub


Sub JunglePrincess_KeyUp(ByVal keycode)

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
		LeftFlip1.RotateToStart
		PlaySound SoundFXDOF("flipperdown",101,DOFOff,DOFContactors), 0, 1, -0.05, 0.05
		StopSound "Buzz"
	End If

	If keycode = RightFlipperKey Then
		RightFlipper.RotateToStart
		RightFlip1.RotateToStart
		PlaySound SoundFXDOF("flipperdown",102,DOFOff,DOFContactors), 0, 1, 0.05, 0.05
		StopSound "Buzz1"
	End If
   End if
End Sub

sub flippertimer_timer()
	LFlip.RotY = LeftFlipper.CurrentAngle
	LFlip1.RotY = LeftFlipper.CurrentAngle-90
	RFlip.RotY = RightFlipper.CurrentAngle
	RFlip1.RotY = RightFlipper.CurrentAngle-90
	Pgate.rotz = Gate.currentangle+25
end sub

Sub PairedlampTimer_timer
testbox.text=extraball
	bumperlighta1.state = bumperlight1.state
	bumperlighta2.state = bumperlight2.state
	bumperlighta3.state = bumperlight3.state
	PupA1.state = Pup1.state
	PupA2.state = pup2.state
	LbotA.state = LtopA.state
	LbotC.state = LtopC.state
end sub

sub coindelay_timer
	addcredit
    coindelay.enabled=false
end sub

Sub addcredit
      credit=credit+1
	  DOF 137, DOFOn
      if credit>25 then credit=25
	  credittxt.setvalue(credit)
	  If B2SOn Then Controller.B2ssetCredits Credit
End sub

Sub Drain_Hit()
	DOF 134, DOFPulse
	PlaySound "drain",0,1,0,0.25
	for each light in GIlights:light.state=0:next
	for each light in bumperlights:light.state=0:next
	for each light in lights:light.state=0:next
	for each light in DTlights:light.state=0:next
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
	DOF 136, DOFPulse
end sub

sub ballrel_hit
	if ballrenabled=1 then
		shootagain.state=lightstateoff
		ballrenabled=0
	end if
end sub

sub scorebonus_timer
   if tilt=false then
		me.interval=110*(doublebonus.state+1)
		if bonus>0 then
			score(player)=score(player)+(1000*(doublebonus.state+1))
			sreels(player).addvalue(1000*(doublebonus.state+1))
			If B2SOn Then Controller.B2SSetScorePlayer player, score(player)
			PlaySound SoundFXDOF("bell" & (doublebonus.state+1)*1000,143,DOFPulse,DOFChimes)
			checkreplays
		  BonusLight(bonus).state=0
		  bonus=bonus-1
		  if bonus >0 then
			BonusLight(bonus).state=1
		  end if
	     else
		  bonus=0
	   end if
	Else
		bonus=0
		for each light in bonuslights: light.state=0: next
   end if
	if bonus=0 then
	 for each light in bonuslights: light.state=0: next
   	 if shootagain.state=lightstateon and tilt=false then
	      newball
 	      ballreltimer.enabled=true
        else
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
	 end if
	 scorebonus.enabled=false
	end if
End sub

sub newgame_timer
	extraball=1
	special=3
	player=1
	DoubleBonus.state=0
	for i = 1 to maxplayers
		sreels(i).resettozero
	    score(i)=0
		rep(i)=0
		pups(i).state=0
	next
    pups(1).state=1
	If B2SOn then
	  for i = 1 to maxplayers
		Controller.B2SSetScorePlayer i, score(i)
		Controller.B2SSetScoreRollover 24 + i, 0
	  next
	End If
    eg=0
	shootagain.state=0
    tilttxt.visible=0
	gamov.text=" "
    If B2SOn then
		Controller.B2SSetGameOver 35,0
		Controller.B2SSetTilt 33,0
		Controller.B2SSetMatch 34,0
	End If
    biptext.text="1"
	matchtxt.text=" "
	dtreset=1
	resetDT.enabled=1
	newball
	ballrelTimer.enabled=true
	newgame.enabled=false
end sub


sub newball
	bonus=0
	ebstate=0
	spstate=0
	bumper1.hashitevent=True
    bumper2.hashitevent=True
    bumper3.hashitevent=True
	Lleft5k.state=0
	Lright5k.state=0
	For each light in BumperLights:light.State = 0: Next
	For each light in RedBumperLights:light.State = 1: Next
	for each light in GIlights:light.state=1:next
	for each light in lights:light.state=1:next
	for each light in lextraball:light.state=0:next
	dtreset=1
	resetDT.enabled=1
End Sub


sub nextball
    if tilt=true then
      tilt=false
      tilttxt.visible=0
	  doublebonus.state=0
		If B2SOn then
			Controller.B2SSetTilt 33,0
			Controller.B2ssetdata 1, 1
		End If
    end if
	if player=1 then ballinplay=ballinplay+1
	DoubleBonus.state=0
	if ballinplay=balls then doublebonus.state=1
	if ballinplay>balls then
		playsound "GameOver"
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
	  turnoff
      matchnum
	  state=false
	  biptext.text=" "
	  gamov.text="GAME OVER"
	  gamov.timerenabled=1
	  tilttxt.timerenabled=1
	  for i=1 to maxplayers
		cplay(i).state=0
		if score(i)>hisc then hisc=score(i)
		pups(i).state=0
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
	  For each light in GIlights:light.state=0:next
  else
	Drain.kick 60,28,0
    ballreltimer.enabled=false
	playsound SoundFXDOF("kickerkick",135,DOFPulse,DOFContactors)
  end if
  ballreltimer.enabled=false
end sub

sub matchnum
	matchnumb=(INT (RND*10))*10
	if matchnumb=0 then
		matchtxt.text="00"
	  else
		matchtxt.text=matchnumb
	end if
	If B2SOn then Controller.B2SSetMatch 34,Matchnumb
	For i=1 to players
		if (matchnumb)=(score(i) mod 100) then
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
	if balls = 3 then
		addscore 1000
	  else
		addscore 100
	end if
	me.timerenabled=1
   end if
End Sub

Sub Bumper1_timer
	BumperTimerRing1.Enabled=0
	if bumperring1.transz>-36 then 	BumperRing1.transz=BumperRing1.transz-4
	if BumperRing1.transz=-36 then
		BumperTimerRing1.enabled=1
		me.timerenabled=0
	end if
End Sub

Sub BumperTimerRing1_timer
	if bumperring1.transz<0 then BumperRing1.transz=BumperRing1.transz+4
	If BumperRing1.transz=0 then BumperTimerRing1.enabled=0
End sub


Sub Bumper2_Hit
   if tilt=false then
	playsound SoundFXDOF("fx_bumper4",109,DOFPulse,DOFContactors)
	DOF 110,DOFPulse
	if BumperLight2.state=1 then
		addscore bumperlitscore
	  else
		addscore bumperoffscore
	end if
	me.timerenabled=1
   end if
End Sub

Sub Bumper2_timer
	BumperTimerRing2.enabled=0
	if BumperRing2.transz>-36 then 	BumperRing2.transz=BumperRing2.transz-4
	if BumperRing2.transz=-36 then
		BumperTimerRing2.enabled=1
		me.timerenabled=0
	end if
End Sub

Sub BumperTimerRing2_timer
	if BumperRing2.transz<0 then BumperRing2.transz=BumperRing2.transz+4
	If BumperRing2.transz=0 then BumperTimerRing2.enabled=0
End sub

Sub Bumper3_Hit
   if tilt=false then
	playsound SoundFXDOF("fx_bumper4",111,DOFPulse,DOFContactors)
	DOF 112,DOFPulse
	if balls = 3 then
		addscore 1000
	  else
		addscore 100
	end if
	me.timerenabled=1
   end if
End Sub

Sub Bumper3_timer
	BumperTimerRing3.enabled=0
	if bumperring3.transz>-36 then BumperRing3.transz=BumperRing3.transz-4
	if BumperRing3.transz=-36 then
		BumperTimerRing3.enabled=1
		me.timerenabled=0
	end if
End Sub

Sub BumperTimerRing3_timer
	if bumperring3.transz<0 then BumperRing3.transz=BumperRing3.transz+4
	If BumperRing3.transz=0 then BumperTimerRing3.enabled=0
End sub


'************** Dingwalls

sub Dingwalls_hit(idx)
	addscore 10
	extraball=extraball+1
	if extraball>3 then extraball=0
	if ebstate=1 Then
		for each light in Lextraball:light.state=0:Next
		Lextraball(extraball).state=1
	end if
	special=special+1
	if special>3 then special=0
	if spstate=1 Then
		for each light in Lspecial:light.state=0:Next
		Lspecial(special).state=1
	end if
end sub

sub dingwall1_hit
	rdw1.visible=0
	RDW1a.visible=1
	dw1step=1
	Me.timerenabled=1
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
	rdw2.visible=0
	RDW2a.visible=1
	dw2step=1
	Me.timerenabled=1
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
	Rdw3.visible=0
	RDW3a.visible=1
	dw3step=1
	Me.timerenabled=1
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
	Rdw4.visible=0
	RDW4a.visible=1
	dw4step=1
	Me.timerenabled=1
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
	Rdw5.visible=0
	RDW5a.visible=1
	dw5step=1
	Me.timerenabled=1
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
	Rdw6.visible=0
	RDW6a.visible=1
	dw6step=1
	Me.timerenabled=1
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

sub TGtopA_hit
	DOF 117, DOFPulse
    LtopA.state=0
	addscore 500
	addbonus
	checkaward
end sub

sub TGtopB_hit
	DOF 118, DOFPulse
	BumperLight2.state=1
    LtopB.state=0
	addscore 500
	addbonus
	checkaward
end sub

sub TGtopC_hit
	DOF 119, DOFPulse
    LtopC.state=0
	addscore 500
	addbonus
	checkaward
end sub

sub TGleft5k_hit
	DOF 120, DOFPulse
	if Lleft5k.state=1 Then
		addscore 5000
	  Else
		addscore 500
	end If
end sub

sub TGright5k_hit
	DOF 121, DOFPulse
	if Lright5k.state=1 Then
		addscore 5000
	  Else
		addscore 500
	end If
end sub

sub TGextraballL_hit
	DOF 122, DOFPulse
	addscore 5000
	if Lextraball1.state = 1 then shootAgain.state = 1
	if Lspecial1.state = 1 then ShootAgain.state = 1
end sub

sub TGbotA_hit
	DOF 123, DOFPulse
	LtopA.state=0
	addscore 500
	addbonus
	checkaward
end sub

sub TGbotC_hit
	DOF 125, DOFPulse
	LtopC.state=0
	addscore 500
	addbonus
	checkaward
end sub

sub TGextraballR_hit
	DOF 124, DOFPulse
	addscore 5000
	if Lextraball4.state = 1 then shootAgain.state = 1
	if Lspecial4.state = 1 then ShootAgain.state = 1
end sub

Sub addbonus
	bonus=bonus+1
	if bonus>15 then bonus=15
	if  bonus = 1 then
		BonusLight(bonus).state=1
	  else
		BonusLight(bonus).state=1
		BonusLight(bonus-1).state=0
		if bonus>10 then bonuslight(10).state=1
	End if
End sub


'********** DropTargets

sub DropTargets_hit (idx)
	addscore 500
	addbonus
end Sub

sub DTMonkey1_dropped
	playsound SoundFXDOF("drop1",113,DOFPulse,DOFContactors)
	Ldt1.state=1
	checkaward
end Sub

sub DTMonkey2_dropped
	playsound SoundFXDOF("drop1",113,DOFPulse,DOFContactors)
	Ldt2.state=1
	checkaward
end Sub

sub DTMonkey3_dropped
	playsound SoundFXDOF("drop1",113,DOFPulse,DOFContactors)
	Ldt3.state=1
	checkaward
end Sub

sub DTMonkey4_dropped
	playsound SoundFXDOF("drop1",113,DOFPulse,DOFContactors)
	Ldt4.state=1
	checkaward
end Sub

sub DTMonkey5_dropped
	playsound SoundFXDOF("drop1",113,DOFPulse,DOFContactors)
	Ldt5.state=1
	checkaward
end Sub

sub DTMonkey6_dropped
	playsound SoundFXDOF("drop1",115,DOFPulse,DOFContactors)
	Ldt6.state=1
	checkaward
end Sub

sub DTMonkey7_dropped
	playsound SoundFXDOF("drop1",115,DOFPulse,DOFContactors)
	Ldt7.state=1
	checkaward
end Sub

sub DTMonkey8_dropped
	playsound SoundFXDOF("drop1",115,DOFPulse,DOFContactors)
	Ldt8.state=1
	checkaward
end Sub

sub DTMonkey9_dropped
	playsound SoundFXDOF("drop1",115,DOFPulse,DOFContactors)
	Ldt9.state=1
	checkaward
end Sub

sub DTMonkey10_dropped
	playsound SoundFXDOF("drop1",115,DOFPulse,DOFContactors)
	Ldt10.state=1
	checkaward
end Sub

'********** Kickers

sub KtopL_hit
	addscore 5000
	if Lextraball2.state = 1 then
		shootAgain.state = 1
		playsound "bell10"
	end If
	if Lspecial2.state = 1 then
		ShootAgain.state = 1
		playsound "bell10"
	end If
	lkickstep=1
	me.timerenabled=1
end Sub

sub KtopL_timer
	Select Case lkickstep
	  Case 2:
		playsound SoundFXDOF("holekick",132,DOFPulse,DOFContactors),0,1,0,0.25
		DOF 139, DOFPulse
		PkickarmL.rotz=10
		KtopL.kick 150,15
	  Case 3:
		PkickarmL.rotz=0
		me.timerenabled=0
	End Select
	lkickstep=lkickstep+1
end Sub

sub KtopR_hit
	addscore 5000
	if Lextraball3.state = 1 then
		shootAgain.state = 1
		playsound "bell10"
	end If
	if Lspecial3.state = 1 then
		ShootAgain.state = 1
		playsound "bell10"
	end If
	rkickstep=1
	me.timerenabled=1
end Sub

sub KtopR_timer
	Select Case rkickstep
	  Case 2:
		playsound SoundFXDOF("holekick",133,DOFPulse,DOFContactors),0,1,0,0.25
		DOF 139, DOFPulse
		PkickarmR.rotz=10
		KtopR.kick 210,15
	  Case 3:
		PkickarmR.rotz=0
		me.timerenabled=0
	End Select
	rkickstep=rkickstep+1
end Sub

sub checkaward
	dim check, check1
	check=DTMonkey1.isdropped+DTMonkey2.isdropped+DTMonkey3.isdropped+DTMonkey4.isdropped+DTMonkey5.isdropped
	if check=5 then
		Lleft5k.state=1
	  Else
		Lleft5k.state=0
	end If
	check1=DTMonkey6.isdropped+DTMonkey7.isdropped+DTMonkey8.isdropped+DTMonkey9.isdropped+DTMonkey10.isdropped
	if check1=5 then
		Lright5k.state=1
	  Else
		Lright5k.state=0
	end If
	if check+check1 = 10 then
		for each light in Lextraball:light.state=0:Next
		Lextraball(Extraball).state=1
		ebstate=1
		if LtopA.state+LtopB.state+LtopC.state=0 then
			for each light in Lspecial:light.state=0:Next
			Lspecial(special).state=1
			spstate=1
		end if
	end if
	if LtopA.state+LtopB.state+LtopC.state=0 then
		DoubleBonus.state=1
	end if
end Sub

sub addscore(points)
  if tilt=false then
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
	checkreplays
end Sub

sub checkreplays
    ' check replays and rollover
	if score(player)=>99999 then
		If B2SOn then Controller.B2SSetScoreRollover 24 + player, 1
		P100k(player).text="100,000"
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

sub resetDT_timer
	select case dtreset
	  Case 1:
		playsound SoundFXDOF("BankReset",114,DOFPulse,DOFContactors)
		for i = 1 to 5: target(i).isdropped=False: DTLights(i-1).state=0: Next
  	  Case 2:
		playsound SoundFXDOF("BankReset",116,DOFPulse,DOFContactors)
		for i = 6 to 10: target(i).isdropped=False: DTLights(i-1).state=0: Next
	  Case 3:
		me.enabled=0
	end select
	dtreset=dtreset+1
end sub


Sub CheckTilt
	If Tilttimer.Enabled = True Then
	 TiltSens = TiltSens + 1
	 if TiltSens = 3 Then
	   Tilt = True
	   tilttxt.visible=1
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
	   tilttxt.visible=1
       	If B2SOn Then Controller.B2SSetTilt 33,1
       	If B2SOn Then Controller.B2ssetdata 1, 0
	   playsound "tilt"
	   turnoff
End Sub

Sub Tilttimer_Timer()
	Tilttimer.Enabled = False
End Sub

sub turnoff
    bumper1.hashitevent=False
    bumper2.hashitevent=False
	bumper3.hashitevent=False
  	LeftFlipper.RotateToStart
	LeftFlip1.RotateToStart
	StopSound "Buzz"
	DOF 101, DOFOff
	RightFlipper.RotateToStart
	RightFlip1.RotateToStart
	StopSound "Buzz1"
	DOF 102, DOFOff
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

Sub a_Pins_Hit (idx)
	PlaySound "pinhit_low", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0, AudioFade(ActiveBall)
End Sub

Sub a_Targets_Hit (idx)
	PlaySound "target", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0, AudioFade(ActiveBall)
End Sub

Sub a_Metals_Thin_Hit (idx)
	PlaySound "metalhit_thin", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
End Sub

Sub a_Metals_Medium_Hit (idx)
	PlaySound "metalhit_medium", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
End Sub

Sub a_Metals2_Hit (idx)
	PlaySound "metalhit2", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
End Sub

Sub a_Gates_Hit (idx)
	PlaySound "gate4", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
End Sub

Sub a_Spinner_Spin
	PlaySound "fx_spinner",0,.25,0,0.25
End Sub

Sub a_Rubbers_Hit(idx)
 	dim finalspeed
  	finalspeed=SQR(activeball.velx * activeball.velx + activeball.vely * activeball.vely)
 	If finalspeed > 20 then
		PlaySound "fx_rubber2", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
	End if
	If finalspeed >= 6 AND finalspeed <= 20 then
 		RandomSoundRubber()
 	End If
End Sub

Sub RubberWheel_hit
 	dim finalspeed
  	finalspeed=SQR(activeball.velx * activeball.velx + activeball.vely * activeball.vely)
 	If finalspeed > 20 then
		PlaySound "fx_rubber2", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
	End if
	If finalspeed >= 6 AND finalspeed <= 20 then
 		RandomSoundRubber()
 	End If
End sub

Sub a_Posts_Hit(idx)
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
    savevalue "JunglePrincess", "credit", credit
    savevalue "JunglePrincess", "hiscore", hisc
    savevalue "JunglePrincess", "match", matchnumb
    savevalue "JunglePrincess", "score1", score(1)
    savevalue "JunglePrincess", "score2", score(2)
	savevalue "JunglePrincess", "balls", balls
	savevalue "JunglePrincess", "freeplay", freeplay
end sub

sub loadhs
    dim temp
	temp = LoadValue("JunglePrincess", "credit")
    If (temp <> "") then credit = CDbl(temp)
    temp = LoadValue("JunglePrincess", "hiscore")
    If (temp <> "") then hisc = CDbl(temp)
    temp = LoadValue("JunglePrincess", "match")
    If (temp <> "") then matchnumb = CDbl(temp)
    temp = LoadValue("JunglePrincess", "score1")
    If (temp <> "") then score(1) = CDbl(temp)
    temp = LoadValue("JunglePrincess", "score2")
    If (temp <> "") then score(2) = CDbl(temp)
    temp = LoadValue("JunglePrincess", "balls")
    If (temp <> "") then balls = CDbl(temp)
    temp = LoadValue("JunglePrincess", "freeplay")
    If (temp <> "") then freeplay = CDbl(temp)
end sub

Sub JunglePrincess_Exit()
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

Function AudioFade(tableobj) ' Fades between front and back of the table (for surround systems or 2x2 speakers, etc), depending on the Y position on the table. "JunglePrincess" is the name of the table
  Dim tmp
  tmp = tableobj.y * 2 / JunglePrincess.height-1
  If tmp > 0 Then
    AudioFade = Csng(tmp ^10)
  Else
    AudioFade = Csng(-((- tmp) ^10) )
  End If
End Function

Function AudioPan(tableobj) ' Calculates the pan for a tableobj based on the X position on the table. "JunglePrincess" is the name of the table
  Dim tmp
  tmp = tableobj.x * 2 / JunglePrincess.width-1
  If tmp > 0 Then
    AudioPan = Csng(tmp ^10)
  Else
    AudioPan = Csng(-((- tmp) ^10) )
  End If
End Function

Function Pan(ball) ' Calculates the pan for a ball based on the X position on the table. "JunglePrincess" is the name of the table
    Dim tmp
    tmp = ball.x * 2 / JunglePrincess.width-1
    If tmp > 0 Then
        Pan = Csng(tmp ^10)
    Else
        Pan = Csng(-((- tmp) ^10) )
    End If
End Function

Function AudioFade(ball) ' Can this be together with the above function ?
  Dim tmp
  tmp = ball.y * 2 / JunglePrincess.height-1
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
  If JunglePrincess.VersionMinor > 3 OR JunglePrincess.VersionMajor > 10 Then
    PlaySound("fx_collide"), 0, Csng(velocity) ^2 / 200, Pan(ball1), 0, Pitch(ball1), 0, 0, AudioFade(ball1)
  Else
    PlaySound("fx_collide"), 0, Csng(velocity) ^2 / 200, Pan(ball1), 0, Pitch(ball1), 0, 0
  End if
End Sub

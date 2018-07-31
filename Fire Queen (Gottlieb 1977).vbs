'	-------------------------------------------------
'	FIRE QUEEN  Gottlieb 1977
'	-------------------------------------------------
'
'	Original table by D. Gottlieb & Co., October 1977
'	VP   Vulcan adaptation by Dave Sanders, December 2002
'	VP10 Vulcan adaptation started by MaX and hauntfreaks, July 2015
'   VP10 Fire Queen and Vulcan by hauntfreaks and BorgDog, Sept 2015
'	DOF coded by BorgDog with instruction(and correction) by arngrim - thanks!
'	-------------------------------------------------

' Thalamus 2018-07-23
' Added/Updated "Positional Sound Playback Functions" and "Supporting Ball & Sound Functions"
' No special SSF tweaks yet.

Option Explicit
Randomize

Const cGameName = "firequeen_1977"

Dim operatormenu, options
Dim bumperlitscore
Dim bumperoffscore
Dim balls
dim ebcount
Dim replays
Dim Replay1Table(3)
Dim Replay2Table(3)
Dim Replay3Table(3)
dim replay1
dim replay2
dim replay3
dim replaytext(3)
Dim hisc
Dim Controller
Dim maxplayers
Dim players
Dim player
Dim credit, dbonus
Dim score(4)
Dim Add10, Add100, Add1000
Dim sreels(4)
Dim p100k(4)
Dim state
Dim tilt
Dim tiltsens
Dim target(9)
Dim holebonus
Dim holeb(4)
Dim StarState
Dim Bonus
Dim Bonuslight(19)
Dim ballinplay
Dim matchnumb
dim ballrenabled
Dim rep
Dim rst
Dim eg
Dim bell
Dim i,j,tnumber, objekt, light
Dim awardcheck

On Error Resume Next
ExecuteGlobal GetTextFile("controller.vbs")
If Err Then MsgBox "Can't open controller.vbs"
On Error Goto 0

sub FireQueen_init
	LoadEM
	maxplayers=2
	Replay1Table(1)=70000
	Replay1Table(2)=90000
	Replay1Table(3)=100000
	Replay2Table(1)=120000
	Replay2Table(2)=110000
	Replay2Table(3)=120000
	Replay3Table(1)=150000
	Replay3Table(2)=160000
	Replay3Table(3)=170000
	replaytext(1)="70000,120000,150000"
	replaytext(2)="90000,110000,160000"
	replaytext(3)="100000,120000,170000"
	set sreels(1) = ScoreReel1
	set sreels(2) = ScoreReel2
	set sreels(3) = ScoreReel3
	set sreels(4) = ScoreReel4
	set p100k(1) = p1100k
	set p100k(2) = p2100k
	set p100k(3) =p3100k
	set p100k(4)=p4100k
	set target(1)=W1a
	set target(2)=W2a
	set target(3)=W3a
	set target(4)=W4a
	set target(5)=W5a
	set target(6)=G1a
	set target(7)=G2a
	set target(8)=G3a
	set target(9)=G4a
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
	set holeb(1)=Hole1
	set holeb(2)=Hole2
	set holeb(3)=Hole3
	set holeb(4)=Hole4
	hideoptions
	player=1
	For each light in LaneLights:light.State = 0: Next
	For each light in BumperLights:light.State = 1: Next
	For each light in Tlights:light.State = 0: Next
	For each light in DTlights:light.State = 0: Next
	loadhs
	if hisc="" then hisc=50000
	hstxt.text=hisc
	gamov.text="Game Over"
	gamov.timerenabled=1
	tilttxt.timerenabled=1
	credittxt.setvalue(credit)
	if balls="" then balls=5
	if balls<>3 and balls<>5 then balls=5
	if replays="" then replays=2
	if replays<>1 and replays<>2 and replays<>3 then replays=2
	Replay1=Replay1Table(Replays)
	Replay2=Replay2Table(Replays)
	Replay3=Replay3Table(Replays)
	OptionBalls.image="OptionsBalls"&Balls
	OptionReplays.image="OptionsReplays"&replays
	RepCard.image = "ReplayCard"&replays
		bumperlitscore=100
		bumperoffscore=100
	if balls=3 then
		InstCard.image="InstCard3balls"
	  Else
		InstCard.image="InstCard5balls"
	end if
 	If B2SOn Then
		setBackglass.enabled=true
		For each objekt in Backdropstuff: objekt.visible=false: next
	End If
	if matchnumb=0 then
		matchtxt.text="00"
	  else
		matchtxt.text=matchnumb*10
	end if
	for i = 1 to maxplayers
		sreels(i).setvalue(score(i))
		if score(i)>99999 then
			p100k(i).text="100,000"
		  else
			p100k(i).text=" "
		end if
	next
	PlaySound "motor"
	tilt=false

	If credit>0 then DOF 133, DOFOn

End sub

sub setBackglass_timer
	Controller.B2ssetCredits Credit
	Controller.B2ssetMatch 34, Matchnumb*10
    Controller.B2SSetGameOver 35,1
    Controller.B2SSetScorePlayer 5, hisc
    Controller.B2SSetScorePlayer 1, Score(1) MOD 100000
    Controller.B2SSetScorePlayer 2, Score(2) MOD 100000
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

Sub FireQueen_KeyDown(ByVal keycode)

	if keycode=AddCreditKey then
		playsound "coinin"
		coindelay.enabled=true
    end if

    if keycode=StartGameKey and credit>0 then
	  if state=false then
		credit=credit-1
		if credit < 1 then DOF 133, DOFOff
		playsound "cluper"
		credittxt.setvalue(credit)
		ballinplay=1
		If b2son Then
			Controller.B2ssetCredits Credit
			Controller.B2sStartAnimation "Startup"
			Controller.B2ssetballinplay 32, Ballinplay
			Controller.B2ssetplayerup 30, 1
			Controller.B2ssetcanplay 31, 1
			Controller.B2SSetGameOver 0
		End If
		shoot1.state=1
		tilt=false
		state=true
		CanPlay1.state=1
		playsound "initialize"
		players=1
		rst=0
		resettimer.enabled=true
	  else if state=true and players < maxplayers and Ballinplay=1 then
		credit=credit-1
		if credit < 1 then DOF 133, DOFOff
		players=players+1
		credittxt.setvalue(credit)
		If b2son then
			Controller.B2ssetCredits Credit
			Controller.B2ssetcanplay 31, 2
		End If
		CanPlay2.state=1
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
			OptionBalls.image = "OptionsBalls"&Balls
		Case 2:
			Replays=Replays+1
			if Replays>3 then
				Replays=1
			end if
			Replay1=Replay1Table(Replays)
			Replay2=Replay2Table(Replays)
			replay3=Replay3Table(Replays)
			OptionReplays.image = "OptionsReplays"&replays
			repcard.image = "replaycard"&replays
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

	If keycode = MechanicalTilt Then
		mechchecktilt
	End If

  end if
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


Sub OperatorMenuTimer_Timer
	OperatorMenu=1
	Displayoptions
	Options=1
End Sub

Sub FireQueen_KeyUp(ByVal keycode)

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

Sub SPtimer_timer
	starstate=0
	if top1.state=0 then starstate=starstate+1
	if top2.state=0 then starstate=starstate+1
	if top3.state=0 then starstate=starstate+1
	if top4.state=0 then starstate=starstate+1
	if spot5.state=0 then starstate=starstate+1
	if starstate=5 then LeftSP.state=1
end sub

Sub PairedlampTimer_timer
	bottom1.state = top1.state
	bottom2.state = top2.state
	bottom3.state = top3.state
	bottom4.state = top4.state
	roll1.state = drop1.state
	roll2.state = drop2.state
	roll3.state = drop3.state
	roll4.state = drop4.state
	roll5.state = drop5.state
	rightSP.state = leftSP.state
	ShootA1.state = Shoot1.state
	ShootA2.state = Shoot2.state
end sub

sub coindelay_timer
	addcredit
    coindelay.enabled=false
end sub

sub resettimer_timer
    rst=rst+1
	for i = 1 to maxplayers
		sreels(i).resettozero
		p100k(i).text = " "
    next
    If b2son then
		for i = 1 to maxplayers
		  Controller.B2SSetScorePlayer i, score(i)  MOD 100000
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
	  DOF 133, DOFOn
      if credit>15 then credit=15
	  credittxt.setvalue(credit)
	  If b2son Then Controller.B2ssetCredits Credit
End sub

Sub Drain_Hit()
	DOF 129, DOFPulse
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
		if bonusx2.state=1 then
			dbonus=2
			me.interval=200
		  else
			dbonus=1
			me.interval=125
		End if
		score(player)=score(player)+(1000*dbonus)
		sreels(player).addvalue(1000*dbonus)
		If B2SOn Then Controller.B2SSetScorePlayer player, score(player)  MOD 100000
		If dbonus=2 Then
			PlaySound SoundFXDOF("bell2000",143,DOFPulse,DOFChimes)
		  Else
			PlaySound SoundFXDOF("bell1000",143,DOFPulse,DOFChimes)
		End If
		checkreplay
		bonuslight(bonus).state=0
		bonus=bonus-1
		if bonus>0 then bonuslight(bonus).state=1
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
		If b2son then Controller.B2ssetplayerup 30, 1
		shoot1.state=1
		shoot2.state=0
		nextball
	  else
		player=player+1
		If b2son then Controller.B2ssetplayerup 30, player
		shoot2.state=1
		shoot1.state=0
		nextball
	  end if
	 end if
	 scorebonus.enabled=false
   end if
End sub


sub newgame
	bumper1.force=12
	bumper2.force=12
	resetDT
	player=1
	shoot1.state=1
	ebcount=0
	for i = 1 to 4:
		holeb(i).state=0
		score(i)=0
	next
	If b2son then
	  for i = 1 to maxplayers
		Controller.B2SSetScorePlayer i, score(i)
	  next
	End If
    eg=0
    rep=0
	for each light in bonuslights:light.state=0:next
    leftsp.state=0
    ExtraBall.state=0
	shootagain.state=lightstateoff
	for each light in bumperlights:light.state=1:next
	for each light in DTlights:light.state=0:next
	for each light in tlights:light.state=1:next
	for each light in lanelights:light.state=1:next
	for each light in numberlights:light.state=1:next
    gamov.text=" "
    tilttxt.text=" "
    If b2son then
		Controller.B2SSetGameOver 35,0
		Controller.B2SSetTilt 33,0
		Controller.B2SSetMatch 34,0
	End If
    biptext.text="1"
	matchtxt.text=" "
    BallRelease.CreateBall
   	BallRelease.kick 135,4,0
	playsound SoundFXDOF("kickerkick",128,DOFPulse,DOFContactors)
end sub

sub newball
	for each light in numberlights:light.state=1:next
	for each light in starlights:light.state=0:next
	if ebcount=0 then
			resetDT
	  else
			resetWT
	end if
	extraball.state=0
	bonusx2.state=0
	leftsp.state=0
	if ballinplay=balls then bonusx2.state=1
End Sub

Sub resetWT
	playsound SoundFXDOF("BankReset",111,DOFPulse,DOFContactors)
	for each objekt in droptargetsWhite: objekt.isdropped=false: Next
	for each light in WDTlights:light.state=0:next
end sub

Sub resetDT
	playsound SoundFXDOF("BankReset",114,DOFPulse,DOFContactors)
	DOF 114, DOFPulse
	for each objekt in droptargets: objekt.isdropped=false: Next
	for each light in DTlights:light.state=0:next
end sub


sub nextball
    if tilt=true then
	  bumper1.force=12
	  bumper2.force=12
      tilt=false
      tilttxt.text=" "
		If b2son then
			Controller.B2SSetTilt 33,0
			Controller.B2ssetdata 1, 1
		End If
    end if
	ebcount=0
	for i = 1 to 4
		holeb(i).state=0
	next
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
		If b2son then Controller.B2ssetballinplay 32, Ballinplay
	end if
End Sub

sub ballreltimer_timer
  if eg=1 then
	  turnoff
      matchnum
	  biptext.text=" "
	  state=false
	  gamov.text="GAME OVER"
	  CanPlay1.state=0
	  CanPlay2.state=0
'	  CanPlay3.text=" "
'	  CanPlay4.text=" "
	  shoot1.state=0
	  shoot2.state=0
'	  shoot3.state=" "
'	  shoot4.state=" "
	  for i=1 to maxplayers
		if score(i)>hisc then hisc=score(i)
	  next
	  hstxt.text=hisc
	  savehs
	  If b2son then
        Controller.B2SSetGameOver 35,1
        Controller.B2ssetballinplay 32, 0
	    Controller.B2sStartAnimation "EOGame"
	    Controller.B2SSetScorePlayer 5, hisc
	    Controller.B2ssetcanplay 31, 0
	    Controller.B2ssetcanplay 30, 0
	  End If
	  For each light in LaneLights:light.State = 0: Next
	  For each light in Tlights:light.State = 0: Next
	  For each light in DTlights:light.State = 0: Next
	  ballreltimer.enabled=false
  else
    BallRelease.CreateBall
	BallRelease.kick 135,4,0
	playsound SoundFXDOF("kickerkick",128,DOFPulse,DOFContactors)
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
  If b2son then Controller.B2SSetMatch 34,Matchnumb*10
  For i=1 to players
    if (matchnumb*10)=(score(i) mod 100) then
		addcredit
		playsound SoundFXDOF("knock",130,DOFPulse,DOFKnocker)
		DOF 131,DOFPulse
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

sub FlashBumpers
	For each light in BumperLights
	  Light.State=0
	next
	FlashB.enabled=1
end sub

sub FlashB_timer
	For each light in BumperLights
	  Light.State=1
	next
	FlashB.enabled=0
end sub

Sub DingWalls_Hit (idx)
	addscore 10
end sub


sub SpotTarget_hit
	spot5.state=0
	PlaySound SoundFXDOF("target",135,DOFPulse,DOFContactors),0,.5
	if balls = 3 then
		addscore 3000
	  else
		addscore 500
	end if
	if extraball.state=1 then
		PlaySound SoundFXDOF("bell10",141,DOFPulse,DOFChimes),0,.5
		if ebcount=0 then
			shootagain.state=lightstateon
			extraball.state=0
			ebcount=1
		 end if
	end if

end sub

sub TopHole_hit
	addscore 1000
	holebonus=0
	for i = 1 to 4
		if holeb(i).state=1 then holebonus=holebonus+1
	next
	addscore 1000*holebonus
	me.timerenabled=1
end sub

Sub TopHole_timer
	TopHole.Kick 205, 12
	playsound SoundFXDOF("ballrelease",132,DOFPulse,DOFContactors),0,1,0,0.25
	DOF 131, DOFPulse
	me.timerenabled=0
End Sub

'****** Wire Triggers

Sub Toplane1_Hit
   if tilt=false then
	DOF 115,DOFPulse
	flashbumpers
	Drop1.State  = 1
	Top1.state = 0
	if balls = 3 then
		addscore 3000
	  else
		addscore 500
	end if
   end if
End Sub

Sub Toplane2_Hit
   if tilt=false then
	DOF 116,DOFPulse
	flashbumpers
	Drop2.State  = 1
	if balls = 3 then
		addscore 3000
	  else
		addscore 500
	end if
	Top2.state = 0
   end if
End Sub

Sub Toplane3_Hit
   if tilt=false then
	DOF 117,DOFPulse
	flashbumpers
	Drop3.State  = 1
	if balls = 3 then
		addscore 3000
	  else
		addscore 500
	end if
	Top3.state = 0
   end if
End Sub

Sub Toplane4_Hit
   if tilt=false then
	DOF 118,DOFPulse
	flashbumpers
	Drop4.State  = 1
	if balls = 3 then
		addscore 3000
	  else
		addscore 500
	end if
	Top4.state = 0
   end if
End Sub

Sub LeftOutlane_Hit
   if tilt=false then
	DOF 124,DOFPulse
	flashbumpers
	  if LeftSP.state=1 then
		playsound SoundFXDOF("knock",130,DOFPulse,DOFKnocker)
		DOF 131,DOFPulse
		addcredit
	  end if
	Drop1.State  = 1
	if balls = 3 then
		addscore 3000
	  else
		addscore 500
	end if
	Top1.state = 0
   end if
End Sub

Sub LeftInlane_Hit
   if tilt=false then
	DOF 125,DOFPulse
	flashbumpers
	Drop2.State  = 1
	if balls = 3 then
		addscore 3000
	  else
		addscore 500
	end if
	Top2.state = 0
   end if
End Sub

Sub RightInlane_Hit
   if tilt=false then
	DOF 127,DOFPulse
	flashbumpers
	Drop3.State  = 1
	if balls = 3 then
		addscore 3000
	  else
		addscore 500
	end if
	Top3.state = 0
   end if
End Sub

Sub RightOutlane_Hit
   if tilt=false then
	DOF 126,DOFPulse
	flashbumpers
	  if LeftSP.state=1 then
		playsound SoundFXDOF("knock",130,DOFPulse,DOFKnocker)
		DOF 131,DOFPulse
		addcredit
	  end if
	Drop4.State  = 1
	if balls = 3 then
		addscore 3000
	  else
		addscore 500
	end if
	Top4.state = 0
   end if
End Sub

'**** Rollovers

Sub Rollover1_Hit
	if tilt=false then
		DOF 119, DOFPulse
		if roll1.state=1 then
			addscore 1000
		  else
			addscore 100
		end if
	end if
end sub

Sub Rollover2_Hit
	if tilt=false then
		DOF 120, DOFPulse
		if roll2.state=1 then
			addscore 1000
		  else
			addscore 100
		end if
	end if
end sub

Sub Rollover3_Hit
	if tilt=false then
		DOF 121, DOFPulse
		if roll3.state=1 then
			addscore 1000
		  else
			addscore 100
		end if
	end if
end sub

Sub Rollover4_Hit
	if tilt=false then
		DOF 122, DOFPulse
		if roll4.state=1 then
			addscore 1000
		  else
			addscore 100
		end if
	end if
end sub

Sub Rollover5_Hit
	if tilt=false then
		DOF 123, DOFPulse
		if roll5.state=1 then
			addscore 1000
		  else
			addscore 100
		end if
	end if
end sub


'**** Drop Targets

Sub AwardCheckTimer_timer
	awardcheck=0
	For j = 1 to 5
		if target(j).isdropped=true then awardcheck=awardcheck+1
	next
	if awardcheck=5 then
		bonusx2.state=1
		resetWT
	end if
	awardcheck=0
	For j = 6 to 9
		if target(j).isdropped=true then awardcheck=awardcheck+1
	next
	if awardcheck=4 and ebcount=0 then extraball.state=1
End Sub

Sub W1a_Hit
	playsound SoundFXDOF("drop1",112,DOFPulse,DOFContactors)
	addscore 500
	addbonus
	If Drop1.State = 1 then
		addbonus
		if balls =3 then addbonus
	end if
	me.isdropped=true
	giw1.state=1
End Sub

Sub W2a_Hit
	playsound SoundFXDOF("drop1",112,DOFPulse,DOFContactors)
	addscore 500
	addbonus
	If Drop2.State = 1 then
		addbonus
		if balls =3 then addbonus
	end if
	me.isdropped=true
	giw2.state=1
End Sub

Sub W3a_Hit
	playsound SoundFXDOF("drop1",112,DOFPulse,DOFContactors)
	addscore 500
	addbonus
	If Drop3.State = 1 then
		addbonus
		if balls =3 then addbonus
	end if
	me.isdropped=true
	giw3.state=1
	giw3b.state=1
End Sub

Sub W4a_Hit
	playsound SoundFXDOF("drop1",112,DOFPulse,DOFContactors)
	addscore 500
	addbonus
	If Drop4.State = 1 then
		addbonus
		if balls =3 then addbonus
	end if
	me.isdropped=true
	giw4.state=1
End Sub

Sub W5a_Hit
	playsound SoundFXDOF("drop1",112,DOFPulse,DOFContactors)
	addscore 500
	addbonus
	If Drop5.State = 1 then
		addbonus
		if balls =3 then addbonus
	end if
	me.isdropped=true
	giw5.state=1
End Sub

Sub G1a_Hit
	playsound SoundFXDOF("drop1",113,DOFPulse,DOFContactors)
	addscore 500
	addbonus
	Hole1.state =1
	me.isdropped=true
End Sub

Sub G2a_Hit
	playsound SoundFXDOF("drop1",113,DOFPulse,DOFContactors)
	addscore 500
	addbonus
	Hole2.state =1
	me.isdropped=true
	gig2.state=1
End Sub

Sub G3a_Hit
	playsound SoundFXDOF("drop1",113,DOFPulse,DOFContactors)
	addscore 500
	addbonus
	Hole3.state =1
	me.isdropped=true
	gig3.state=1
End Sub

Sub G4a_Hit
	playsound SoundFXDOF("drop1",113,DOFPulse,DOFContactors)
	addscore 500
	addbonus
	Hole4.state =1
	me.isdropped=true
	gig4.state=1
End Sub

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
	checkreplay
end sub

sub checkreplay
	if score(player)=>99999 then
		if b2son then Controller.B2SSetScoreRollover 24 + player, 1
		P100k(player).text="100,000"
	End if
    if score(player)=>replay1 and rep=0 then
		playsound SoundFXDOF("knock",130,DOFPulse,DOFKnocker)
		DOF 131,DOFPulse
	  addcredit
      rep=1
    end if
    if score(player)=>replay2 and rep=1 then
		playsound SoundFXDOF("knock",130,DOFPulse,DOFKnocker)
		DOF 131,DOFPulse
      addcredit
      rep=2
    end if
    if score(player)=>replay3 and rep=2 then
		playsound SoundFXDOF("knock",130,DOFPulse,DOFKnocker)
		DOF 131,DOFPulse
      addcredit
      rep=3
    end if
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

Sub MechCheckTilt
	Tilt = True
	tilttxt.text="TILT"
	If b2son Then Controller.B2SSetTilt 33,1
	If b2son Then Controller.B2ssetdata 1, 0
	playsound "tilt"
	turnoff
	LeftFlipper.RotateToStart
	RightFlipper.RotateToStart
End Sub

Sub Tilttimer_Timer()
	Tilttimer.Enabled = False
End Sub

sub turnoff
    bumper1.force=0
    bumper2.force=0
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
	Bonuslight(bonus).state=1
	if bonus>1 then Bonuslight(bonus-1).state=0
	if bonus>10 then Bonuslight(10).state=1
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

    savevalue "FireQueen", "credit", credit
    savevalue "FireQueen", "hiscore", hisc
    savevalue "FireQueen", "match", matchnumb
    savevalue "FireQueen", "score1", score(1)
    savevalue "FireQueen", "score2", score(2)
    savevalue "FireQueen", "score3", score(3)
    savevalue "FireQueen", "score4", score(4)
	savevalue "FireQueen", "replays", replays
	savevalue "FireQueen", "balls", balls

end sub

sub loadhs
    dim temp
	temp = LoadValue("FireQueen", "credit")
    If (temp <> "") then credit = CDbl(temp)
    temp = LoadValue("FireQueen", "hiscore")
    If (temp <> "") then hisc = CDbl(temp)
    temp = LoadValue("FireQueen", "match")
    If (temp <> "") then matchnumb = CDbl(temp)
    temp = LoadValue("FireQueen", "score1")
    If (temp <> "") then score(1) = CDbl(temp)
    temp = LoadValue("FireQueen", "score2")
    If (temp <> "") then score(2) = CDbl(temp)
    temp = LoadValue("FireQueen", "score3")
    If (temp <> "") then score(3) = CDbl(temp)
    temp = LoadValue("FireQueen", "score4")
    If (temp <> "") then score(4) = CDbl(temp)
    temp = LoadValue("FireQueen", "replays")
    If (temp <> "") then replays = CDbl(temp)
    temp = LoadValue("FireQueen", "balls")
    If (temp <> "") then balls = CDbl(temp)
end sub

Sub FireQueen_Exit()
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

Function AudioFade(tableobj) ' Fades between front and back of the table (for surround systems or 2x2 speakers, etc), depending on the Y position on the table. "table1" is the name of the table
  Dim tmp
  tmp = tableobj.y * 2 / table1.height-1
  If tmp > 0 Then
    AudioFade = Csng(tmp ^10)
  Else
    AudioFade = Csng(-((- tmp) ^10) )
  End If
End Function

Function AudioPan(tableobj) ' Calculates the pan for a tableobj based on the X position on the table. "table1" is the name of the table
  Dim tmp
  tmp = tableobj.x * 2 / table1.width-1
  If tmp > 0 Then
    AudioPan = Csng(tmp ^10)
  Else
    AudioPan = Csng(-((- tmp) ^10) )
  End If
End Function

Function Pan(ball) ' Calculates the pan for a ball based on the X position on the table. "table1" is the name of the table
    Dim tmp
    tmp = ball.x * 2 / table1.width-1
    If tmp > 0 Then
        Pan = Csng(tmp ^10)
    Else
        Pan = Csng(-((- tmp) ^10) )
    End If
End Function

Function AudioFade(ball) ' Can this be together with the above function ?
  Dim tmp
  tmp = ball.y * 2 / FireQueen.height-1
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
  If FireQueen.VersionMinor > 3 OR FireQueen.VersionMajor > 10 Then
    PlaySound("fx_collide"), 0, Csng(velocity) ^2 / 200, Pan(ball1), 0, Pitch(ball1), 0, 0, AudioFade(ball1)
  Else
    PlaySound("fx_collide"), 0, Csng(velocity) ^2 / 200, Pan(ball1), 0, Pitch(ball1), 0, 0
  End if
End Sub

' Thalamus : Exit in a clean and proper way
Sub FireQueen_exit()
  Controller.Pause = False
  Controller.Stop
End Sub


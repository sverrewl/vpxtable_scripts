'	-------------------------------------------------
'	JOKER POKER Gottlieb 1978
'	vp10 by BorgDog, 2015
' 	thanks to scottamus on flickr for the backglass photo
'	-------------------------------------------------
Option Explicit
Randomize


' Thalamus 2018-07-23
' Added/Updated "Positional Sound Playback Functions" and "Supporting Ball & Sound Functions"
' Thalamus 2018-11-01 : Improved directional sounds
' !! NOTE : Table not verified yet !!

' Options
' Volume devided by - lower gets higher sound

Const VolDiv = 2000    ' Lower number, louder ballrolling/collition sound
Const VolCol = 10      ' Ball collition divider ( voldiv/volcol )

' The rest of the values are multipliers
'
'  .5 = lower volume
' 1.5 = higher volume

Const VolBump   = 2    ' Bumpers volume.
Const VolGates  = 1    ' Gates volume.
Const VolMetal  = 1    ' Metals volume.
Const VolRH     = 1    ' Rubber hits volume.
Const VolPo     = 1    ' Rubber posts volume.
Const VolPi     = 1    ' Rubber pins volume.
Const VolTarg   = 1    ' Targets volume.
Const VolFlip   = 1    ' Flipper volume.

Const cGameName = "jokerpoker_1978"

Dim operatormenu, Options
Dim bonuscount
Dim balls, PlungeBall, rotatortemp, tempplayerup
Dim replays
Dim Replay1Table(3), Replay2Table(3), Replay3Table(3)
dim replay1, replay2, replay3
Dim Add10, Add100, Add1000
Dim hisc
Dim maxplayers, players, player
Dim credit, freeplay
Dim score(4)
Dim state
Dim tilt
Dim tiltsens
Dim ballinplay
Dim matchnumb
dim ballrenabled
dim rstep, lstep, dtstep
Dim rep(4)
Dim rst
Dim eg
Dim scn
Dim scn1
Dim bell
Dim i,j, ii, objekt, light

On Error Resume Next
ExecuteGlobal GetTextFile("controller.vbs")
If Err Then MsgBox "Can't open controller.vbs"
On Error Goto 0

sub JokerPoker_init
	LoadEM
	maxplayers=4
	Replay1Table(1)=90000
	Replay2Table(1)=130000
	Replay3Table(1)=160000
	Replay1Table(2)=110000
	Replay2Table(2)=130000
	Replay3Table(2)=190000
	Replay1Table(3)=120000
	Replay2Table(3)=150000
	Replay3Table(3)=190000
	hideoptions
	player=1
	freeplay=0
	balls=5
	replays=1
	hisc=20000
	loadhs
	if HSA1="" then HSA1=0
	if HSA2="" then HSA2=0
	if HSA3="" then HSA3=0
	UpdatePostIt
	gamov.text=""
	credittxt.setvalue(credit)
	Replay1=Replay1Table(Replays)
	Replay2=Replay2Table(Replays)
	Replay3=Replay3Table(Replays)
	OptionBalls.image="OptionsBalls"&Balls
	OptionReplays.image="OptionsReplays"&replays
	OptionFreeplay.image="OptionsFreeplay"&freeplay
	RepCard.image = "ReplayCard"&replays
	if balls=3 then
		InstCard.image="InstCard3balls"
	  else
		InstCard.image="InstCard5balls"
	end if
	If B2SOn then
		for each objekt in backdropstuff : objekt.visible = 0 : next
	End If

	startGame.enabled=true
	if matchnumb="" then matchnumb=100
	if matchnumb=100 then
		matchtxt.text="00"
	  else
		matchtxt.text=matchnumb
	end if
	for i = 1 to maxplayers
		EVAL("ScoreReel"&i).setvalue(score(i))
		if score(i)>99999 and score(i)<>"" then
			EVAL("p100k"&i).text="100,000"
		  else
			EVAL("p100k"&i).text=" "
		end if
	next
	tilt=false
	If credit > 0 or freeplay = 1 Then DOF 113, DOFOn
	Drain.CreateBall
End sub

sub startGame_timer
	playsound "poweron"
	lightdelay.enabled=true
	gamov.timerenabled=1
	tilttxt.timerenabled=1
	If B2SOn Then
		if freeplay=0 then Controller.B2ssetCredits Credit
		Controller.B2ssetMatch 34, Matchnumb
		Controller.B2SSetGameOver 35,1
		Controller.B2SSetScorePlayer 5, hisc
		for i = 1 to maxplayers
			Controller.B2SSetScorePlayer i, Score(i) MOD 100000
		next
	End If
	me.enabled=false
end sub

sub lightdelay_timer
	If (Credit > 0 or freeplay=1) Then DOF 113, 1
	For each light in GIlights:light.state=1:Next
	For each light in BumperLights:light.state=1:Next
	If B2SOn then Controller.B2SSetData 99,1
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

Sub JokerPoker_KeyDown(ByVal keycode)

	if keycode=AddCreditKey then
		playsoundAtVol "coinin", drain, 1
		coindelay.enabled=true
    end if

    if keycode=StartGameKey and (Credit>0 or freeplay=1) And Not HSEnterMode=true then
	  if state=false then
		player=1
'		playsound "cluper"
		if freeplay=0 then
			credit=credit-1
			If credit < 1 Then DOF 113, DOFOff
			credittxt.setvalue(credit)
		end if
		ballinplay=1
		If B2SOn Then
			if freeplay=0 then Controller.B2ssetCredits Credit
			Controller.B2ssetballinplay 32, Ballinplay
			Controller.B2ssetplayerup 1
			Controller.B2ssetcanplay 31, 1
			Controller.B2SSetGameOver 0
			Controller.B2SSetScorePlayer 5, hisc
		End If
	    pup1.state=1
		tilt=false
		state=true
		playsound "initialize"
		PlaySound("RotateThruPlayers"),0,.05,0,0.25
 		TempPlayerUp=Player
		PlayerUpRotator.enabled=true
		players=1
		CANPLAY1.state=1
		rst=0
		resettimer.enabled=true
	  else if state=true and (Credit>0 or freeplay=1) and players < maxplayers and Ballinplay=1 then
		if freeplay=0 then
			credit=credit-1
			If credit < 1 Then DOF 113, DOFOff
			credittxt.setvalue(credit)
		end if
		EVAL("Canplay"&players).state=0
		players=players+1
		EVAL("Canplay"&players).state=1
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
		PlaySoundAtVol "plungerpull", Plunger, 1
	End If

	If keycode=LeftFlipperKey and State = false and OperatorMenu=0 then
		OperatorMenuTimer.Enabled = true
	end if

	If keycode=LeftFlipperKey and State = false and OperatorMenu=1 then
		Options=Options+1
		If Options=5 then Options=1
		playsound "drop1"
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
			Replay3=Replay3Table(Replays)
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
		PlaySoundAtVol SoundFXDOF("flipperup",101,DOFOn,DOFContactors), LeftFlipper, VolFlip
		PlaySoundAtVol "Buzz", LeftFlipper, VolFlip
	End If

	If keycode = RightFlipperKey Then
		RightFlipper.RotateToEnd
		RightFlip1.RotateToEnd
		Lupflip.state=1
		PlaySoundAtVol SoundFXDOF("flipperup",102,DOFOn,DOFContactors), RightFlipper, VolFlip
		PlaySoundAtVol SoundFXDOF("flipperup",102,DOFOn,DOFContactors), RightFlip1, VolFlip
		PlaySoundAtVol "Buzz1", RightFlipper, VolFlip
		PlaySoundAtVol "Buzz1", RightFlip1, VolFlip
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


Sub JokerPoker_KeyUp(ByVal keycode)

	If keycode = PlungerKey Then
		Plunger.Fire
		if PlungeBall=1 then
			playsoundAtVol "plunger", Plunger, 1
		  else
			playsoundAtVol "plungerreleasefree", Plunger, 1
		end if
	End If

	if keycode = LeftFlipperKey then
		OperatorMenuTimer.Enabled = false
	end if

   If tilt=false and state=true then
	If keycode = LeftFlipperKey Then
		LeftFlipper.RotateToStart
		PlaySoundAtVol SoundFXDOF("flipperdown",101,DOFOff,DOFContactors), LeftFlipper, VolFlip
		StopSound "Buzz"
	End If

	If keycode = RightFlipperKey Then
		RightFlipper.RotateToStart
		Rightflip1.RotateToStart
		Lupflip.state=0
		PlaySoundAtVol SoundFXDOF("flipperdown",102,DOFOff,DOFContactors), RightFlipper, VolFlip
		PlaySoundAtVol SoundFXDOF("flipperdown",102,DOFOff,DOFContactors), RightFlip1, VolFlip
		StopSound "Buzz1"
	End If
   End if
End Sub

sub coindelay_timer
	addcredit
    coindelay.enabled=false
end sub

sub resettimer_timer
    rst=rst+1
	for i = 1 to maxplayers
		EVAL("ScoreReel"&i).resettozero
    next
    If B2SOn then
		for i = 1 to maxplayers
		  Controller.B2SSetScorePlayer i, score(i)
		  Controller.B2SSetScoreRollover 24 + i, 0
		next
	End If
    if rst=18 then
		newgame
		resettimer.enabled=false
    end if
end sub

Sub addcredit
	if freeplay=0 then
      credit=credit+1
	  DOF 113, DOFOn
      if credit>15 then credit=15
	  credittxt.setvalue(credit)
	  If B2SOn Then Controller.B2ssetCredits Credit
	end if
End sub

Sub Drain_Hit()
	PlaySoundAtVol "drain", drain, 1
	DOF 128, DOFPulse
	me.timerenabled=1
End Sub

Sub Drain_timer
'	for each light in BonusLights:light.state=1: next    '***** for testing bonus scoring routine
	scorebonus.enabled=true
	me.timerenabled=0
End Sub

sub ballhome_hit
	ballrenabled=1
	plungeball=1
end sub

sub ballhome_unhit
	DOF 129, DOFPulse
	plungeball=0
end sub

sub ballrel_hit
	if ballrenabled=1 then
		shootagain.state=lightstateoff
		If B2SOn then Controller.B2SSetShootAgain 36,0
		ballrenabled=0
	end if
end sub

sub scorebonus_timer
	if tilt=true Then
		LDoubleBonus.timerenabled=true
	  else
		bonuscount=-1
		Lbonus1.timerenabled=true
	end if
	scorebonus.enabled=false
End sub

sub Lbonus1_timer
	bonuscount=bonuscount+1
	if bonuscount>LDoubleBonus.state then
		bonuscount=-1
		Lbonus2.timerenabled=true
		Lbonus1.timerenabled=false
		Lbonus1.timerinterval=300
	 else
		if AddScore1000Timer.enabled = false then
			playsound "scorebonus"
			Lbonus1A.state=1
			if Lbonus1.state=1 then addscore 1000
		else
			bonuscount=bonuscount-1
			Lbonus1.timerinterval=100
		end if
	end if
end sub

sub Lbonus2_timer
	bonuscount=bonuscount+1
	if bonuscount>LDoubleBonus.state then
		bonuscount=-1
		Lbonus3.timerenabled=true
		Lbonus2.timerenabled=false
		Lbonus2.timerinterval=300
	 else
		  if AddScore1000Timer.enabled = false then
			  playsound "scorebonus"
			  Lbonus1A.state=0
			  Lbonus2A.state=1
			  if Lbonus2.state=1 then addscore 2000
		  else
			bonuscount=bonuscount-1
			Lbonus2.timerinterval=100
		  end if
	end if
end sub

sub Lbonus3_timer
	bonuscount=bonuscount+1
	if bonuscount>LDoubleBonus.state then
		bonuscount=-1
		Lbonus4.timerenabled=true
		Lbonus3.timerenabled=false
		Lbonus3.timerinterval=300
	 else
		  if AddScore1000Timer.enabled = false then
			  playsound "scorebonus"
			  Lbonus2A.state=0
			  Lbonus3A.state=1
			  if Lbonus3.state=1 then addscore 3000
		  else
			bonuscount=bonuscount-1
			Lbonus3.timerinterval=100
		  end if
	end if
end sub

sub Lbonus4_timer
	bonuscount=bonuscount+1
	if bonuscount>LDoubleBonus.state then
		bonuscount=-1
		Lbonus5.timerenabled=true
		Lbonus4.timerenabled=false
		Lbonus4.timerinterval=300
	 else
		  if AddScore1000Timer.enabled = false then
			  playsound "scorebonus"
			  Lbonus3A.state=0
			  Lbonus4A.state=1
			  if Lbonus4.state=1 then addscore 4000
		  else
			bonuscount=bonuscount-1
			Lbonus4.timerinterval=100
		  end if
	end if
end sub

sub Lbonus5_timer
	bonuscount=bonuscount+1
	if bonuscount>LDoubleBonus.state then
		LDoubleBonus.timerenabled=true
		Lbonus5.timerenabled=false
		Lbonus5.timerinterval=300
	 else
		  if AddScore1000Timer.enabled = false then
			  playsound "scorebonus"
			  Lbonus4A.state=0
			  Lbonus5A.state=1
			  if Lbonus5.state=1 then addscore 5000
		  else
			bonuscount=bonuscount-1
			Lbonus5.timerinterval=100
		  end if
	end if
end sub

sub LDoubleBonus_timer
		  if AddScore1000Timer.enabled = false then
			  Lbonus5a.state=0
			  if shootagain.state=lightstateon and tilt=false then
				newball
				ballreltimer.enabled=true
			  else
			   if players=1 or player=players then
				 player=1
			   else
				 player=player+1
			   end if
				If B2SOn then Controller.B2ssetplayerup player
'				for i = 1 to maxplayers
'					EVAL("pup"&i).state=0
'				Next
'				EVAL("pup"&player).state=1
				nextball
			 end if
			 LDoubleBonus.timerenabled=false
		  end if
end sub


sub newgame
	for i=1 to 2
		EVAL("Bumper"&i).hashitevent = 1
	Next
	player=1
	for i = 1 to maxplayers
		EVAL("p100k"&i).text=""
		EVAL("pup"&i).state=0
	    score(i)=0
		rep(i)=0
	next
    pup1.state=1
	If B2SOn then
	  for i = 1 to maxplayers
		Controller.B2SSetScorePlayer i, score(i)
	  next
	End If
    eg=0
	shootagain.state=lightstateoff
	shootagain.state=0
    tilttxt.text=" "
	gamov.text=" "
    If B2SOn then
		Controller.B2SSetGameOver 35,0
		Controller.B2SSetTilt 33,0
		Controller.B2SSetMatch 34,0
	End If
	bipreel.setvalue 1
	matchtxt.text=" "
	newball
	Drain.kick 60,45,0
    PlaySoundAtVol SoundFXDOF("holekick",127,DOFPulse,DOFContactors), drain, 1
end sub


sub newball
	DTstep=1
	resetDT.enabled=true
	bumperlight1.state=1
	for each light in GIlights:light.state=1:next
	For each light in ABClights:light.State = 1: Next
	for each light in BonusLights:light.state=0: next
	LExtraBall.state=0
	Lspecial.state=0
End Sub


sub nextball
    if tilt=true then
	  for i=1 to 2
		EVAL("Bumper"&i).hashitevent = 1
	  Next
      tilt=false
      tilttxt.text=" "
		If B2SOn then
			Controller.B2SSetTilt 33,0
'			Controller.B2ssetdata 1, 1
		End If
    end if
	if player=1 then ballinplay=ballinplay+1
	if ballinplay>balls then
		playsoundAtVol "GameOver"
		eg=1
		ballreltimer.enabled=true
	  else
			PlaySound("RotateThruPlayers"),0,.05,0,0.25
			TempPlayerUp=Player
			PlayerUpRotator.enabled=true
'		biptext.text=ballinplay
		bipreel.setvalue ballinplay
		If B2SOn then
			Controller.B2ssetballinplay 32, Ballinplay
			Controller.B2ssetplayerUp Player
		end if
		if state=true then
		  newball
		  ballreltimer.enabled=true
		end if

	end if
End Sub

sub ballreltimer_timer
  if eg=1 then
	  turnoff
      matchnum
	  state=false
	  bipreel.setvalue 0
	  gamov.text="GAME OVER"
	  gamov.timerenabled=1
	  tilttxt.timerenabled=1
	  for i = 1 to maxplayers
		EVAL("Canplay"&i).state=0
		EVAL("pup"&i).state = 0
	  next
	  checkhighscore
	  If B2SOn then
        Controller.B2SSetGameOver 35,1
        Controller.B2ssetballinplay 32, 0
	    Controller.B2sStartAnimation "EOGame"
	    Controller.B2SSetScorePlayer 5, hisc
	    Controller.B2ssetcanplay 31, 0
	    Controller.B2ssetplayerup 0
	  End If
	  ballreltimer.enabled=false
  else
	Drain.kick 60,45,0
    PlaySoundAtVol SoundFXDOF("holekick",127,DOFPulse,DOFContactors), drain, 1
    ballreltimer.enabled=false
  end if
end sub

Sub PlayerUpRotator_timer()
		If RotatorTemp<5 then
			TempPlayerUp=TempPlayerUp+1
			If TempPlayerUp>4 then
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
				Controller.B2SSetData 80+TempPlayerUp,1
			End If

		else
			if B2SOn then
				Controller.B2SSetPlayerUp Player
				Controller.B2SSetData 81,0
				Controller.B2SSetData 82,0
				Controller.B2SSetData 83,0
				Controller.B2SSetData 84,0
				Controller.B2SSetData 80+Player,1
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

sub CheckHighScore
	dim hiscstate
	hiscstate=0
	for i=1 to players
		if score(i)>hisc then
			hisc=score(i)
			hiscstate=1
		end if
	next
	if hiscstate=1 then HighScoreEntryInit()
	UpdatePostIt
	savehs
end sub

sub matchnum
	matchnumb=(INT (RND*10))*10
	if matchnumb=0 then
		matchtxt.text="00"
		matchnumb=100
	  else
		matchtxt.text=matchnumb
	end if
	If B2SOn then Controller.B2SSetMatch 34,Matchnumb
	For i=1 to players
		if (matchnumb)=(score(i) mod 100) then
		  addcredit
		  PlaySound SoundFXDOF("knock",111,DOFPulse,DOFKnocker)
		  DOF 112, DOFPulse
	    end if
    next
end sub

'********** Bumpers

Sub Bumper1_Hit     'left bumper
   if tilt=false then
    PlaySoundAtVol SoundFXDOF("fx_bumper4",107, DOFPulse,DOFContactors), Bumper1, VolBump
	DOF 108, DOFPulse
	FlashBumpers
	if balls = 3 then
		addscore 1000
	  else
		addscore 100
	end if
   end if
End Sub


Sub Bumper2_Hit    'right bumper
   if tilt=false then
    PlaySoundAtVol SoundFXDOF("fx_bumper4",109, DOFPulse,DOFContactors), Bumper2, VolBump
	DOF 110, DOFPulse
	FlashBumpers
	if balls = 3 then
		addscore 1000
	  else
		addscore 100
	end if
   end if
End Sub



Sub PairedlampTimer_timer
	bumper1light1.state=BumperLight1.state
	bumper2light1.state=bumperlight2.state
	LleftA.state=LtopA.state
	LleftB.state=LtopB.state
	LrightB.state=LtopB.state
	LrightC.state=LtopC.state
	PupA1.state = Pup1.state
	PupA2.state = pup2.state
	PupA3.state = PUP3.state
	PupA4.state = PUP4.state
	LFlip.RotY = LeftFlipper.CurrentAngle
	RFlip.RotY = RightFlipper.CurrentAngle
	RFlip1.RotY = RightFlip1.CurrentAngle-90
	Pgate.rotz = (Gate.currentangle*.75)+25
end sub

sub FlashBumpers
	if bumperlight1.state = 1 then
		for each light in BumperLights: light.duration 0, 200, 1:Next
	end If
end sub



'************** Slings

Sub RightSlingShot_Slingshot
    PlaySoundAtVol SoundFXDOF("slingshot",103,DOFPulse,DOFContactors), sling1, 1
	DOF 105, DOFPulse
	addscore 10
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
    PlaySoundAtVol SoundFXDOF("slingshot",103,DOFPulse,DOFContactors), slingL, 1
	DOF 105, DOFPulse
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

Sub Dingwalls_Hit(idx)
	addscore 10
End Sub

Sub DingwallA_Hit()
	SlingA.visible=0
	SlingA1.visible=1
	dingwalla.uservalue=1
	Me.timerenabled=1
End Sub

sub dingwalla_timer
	select case dingwalla.uservalue
		Case 1: SlingA1.visible=0: SlingA.visible=1
		case 2:	SlingA.visible=0: SlingA2.visible=1
		Case 3: SlingA2.visible=0: SlingA.visible=1: Me.timerenabled=0
	end Select
	dingwalla.uservalue=dingwalla.uservalue+1
end sub

Sub DingwallB_Hit()
	Slingb.visible=0
	Slingb1.visible=1
	dingwallb.uservalue=1
	Me.timerenabled=1
End Sub

sub dingwallb_timer
	select case DingwallB.uservalue
		Case 1: Slingb1.visible=0: SlingB.visible=1
		case 2:	SlingB.visible=0: Slingb2.visible=1
		Case 3: Slingb2.visible=0: SlingB.visible=1: Me.timerenabled=0
	end Select
	dingwallb.uservalue=DingwallB.uservalue+1
end sub

Sub DingwallC_Hit()
	Slingc.visible=0
	Slingc1.visible=1
	dingwallc.uservalue=1
	Me.timerenabled=1
End Sub

sub dingwallc_timer
	select case Dingwallc.uservalue
		Case 1: Slingc1.visible=0: Slingc.visible=1
		case 2:	Slingc.visible=0: Slingc2.visible=1
		Case 3: Slingc2.visible=0: Slingc.visible=1: Me.timerenabled=0
	end Select
	dingwallc.uservalue=Dingwallc.uservalue+1
end sub

Sub DingwallD_Hit()
	Slingd.visible=0
	Slingd1.visible=1
	dingwalld.uservalue=1
	Me.timerenabled=1
End Sub

sub dingwalld_timer
	select case Dingwalld.uservalue
		Case 1: Slingd1.visible=0: Slingd.visible=1
		case 2:	Slingd.visible=0: Slingd2.visible=1
		Case 3: Slingd2.visible=0: Slingd.visible=1: Me.timerenabled=0
	end Select
	dingwalld.uservalue=Dingwalld.uservalue+1
end sub

Sub DingwallE_Hit()
	Slinge.visible=0
	Slinge1.visible=1
	dingwalle.uservalue=1
	Me.timerenabled=1
End Sub

sub dingwalle_timer
	select case Dingwalle.uservalue
		Case 1: Slinge1.visible=0: Slinge.visible=1
		case 2:	Slinge.visible=0: Slinge2.visible=1
		Case 3: Slinge2.visible=0: Slinge.visible=1: Me.timerenabled=0
	end Select
	dingwalle.uservalue=Dingwalle.uservalue+1
end sub

Sub DingwallF_Hit()
	Slingf.visible=0
	Slingf1.visible=1
	dingwallf.uservalue=1
	Me.timerenabled=1
End Sub

sub dingwallf_timer
	select case Dingwallf.uservalue
		Case 1: Slingf1.visible=0: Slingf.visible=1
		case 2:	Slingf.visible=0: Slingf2.visible=1
		Case 3: Slingf2.visible=0: Slingf.visible=1: Me.timerenabled=0
	end Select
	dingwallf.uservalue=Dingwallf.uservalue+1
end sub

Sub DingwallG_Hit()
	Slingg.visible=0
	Slingg1.visible=1
	dingwallg.uservalue=1
	Me.timerenabled=1
End Sub

sub dingwallg_timer
	select case Dingwallg.uservalue
		Case 1: Slingg1.visible=0: Slingg.visible=1
		case 2:	Slingg.visible=0: Slingg2.visible=1
		Case 3: Slingg2.visible=0: Slingg.visible=1: Me.timerenabled=0
	end Select
	dingwallg.uservalue=Dingwallg.uservalue+1
end sub

Sub DingwallH_Hit()
	SlingH.visible=0
	Slingh1.visible=1
	dingwallh.uservalue=1
	Me.timerenabled=1
End Sub

sub dingwallh_timer
	select case Dingwallh.uservalue
		Case 1: Slingh1.visible=0: Slingh.visible=1
		case 2:	Slingh.visible=0: Slingh2.visible=1
		Case 3: Slingh2.visible=0: SlingH.visible=1: Me.timerenabled=0
	end Select
	DingwallH.uservalue=Dingwallh.uservalue+1
end sub

Sub DingwallI_Hit()
	SlingI.visible=0
	Slingi1.visible=1
	DingwallI.uservalue=1
	Me.timerenabled=1
End Sub

sub dingwalli_timer
	select case DingwallI.uservalue
		Case 1: Slingi1.visible=0: SlingI.visible=1
		case 2:	SlingI.visible=0: Slingi2.visible=1
		Case 3: Slingi2.visible=0: SlingI.visible=1: Me.timerenabled=0
	end Select
	DingwallI.uservalue=DingwallI.uservalue+1
end sub

Sub DingwallJ_Hit()
	SlingJ.visible=0
	Slingj1.visible=1
	DingwallJ.uservalue=1
	Me.timerenabled=1
End Sub

sub dingwallj_timer
	select case DingwallJ.uservalue
		Case 1: Slingj1.visible=0: Slingj.visible=1
		case 2:	SlingJ.visible=0: Slingj2.visible=1
		Case 3: Slingj2.visible=0: SlingJ.visible=1: Me.timerenabled=0
	end Select
	DingwallJ.uservalue=DingwallJ.uservalue+1
end sub

'********** Triggers

sub TGspecial_hit
	DOF 130, DOFPulse
	addscore 500
	if Lspecial.state = 1 then
		addcredit
		PlaySoundAtVol SoundFXDOF("knock",111,DOFPulse,DOFKnocker), ActiveBall, 1
		DOF 112, DOFPulse
	end if
end sub

sub TGleftA_hit
	DOF 123, DOFPulse
	if LtopA.state=1 then
		addscore 5000
	  else
		addscore 500
	end if
	LtopA.state=0
	CheckAwards
end sub

sub TGleftB_hit
	DOF 124, DOFPulse
	if LtopB.state=1 then
		addscore 5000
	  else
		addscore 500
	end if
	LtopB.state=0
	CheckAwards
end sub

sub TGrightB_hit
	DOF 125, DOFPulse
	if LtopB.state=1 then
		addscore 5000
	  else
		addscore 500
	end if
	LtopB.state=0
	CheckAwards
end sub

sub TGrightC_hit
	DOF 126, DOFPulse
	if LtopC.state=1 then
		addscore 5000
	  else
		addscore 500
	end if
	LtopC.state=0
	CheckAwards
end sub

sub TGtopA_hit
	DOF 120, DOFPulse
	if LtopA.state=1 then
		addscore 5000
	  else
		addscore 500
	end if
	LtopA.state=0
	CheckAwards
end sub

sub TGtopB_hit
	DOF 121, DOFPulse
	if LtopB.state=1 then
		addscore 5000
	  else
		addscore 500
	end if
	LtopB.state=0
	CheckAwards
end sub

sub TGtopC_hit
	DOF 122, DOFPulse
	if LtopC.state=1 then
		addscore 5000
	  else
		addscore 500
	end if
	LtopC.state=0
	CheckAwards
end sub


'********** Drop targets

Sub DT10_dropped
	playsoundAtVol SoundFXDOF("drop1",118,DOFPulse,DOFContactors), ActiveBall, 1
	addscore 500
	lbonus1.state=1
End Sub


Sub DTJ1_dropped
	playsoundAtVol SoundFXDOF("drop1",116,DOFPulse,DOFContactors), ActiveBall, 1
	addscore 500
	CheckAwards
End Sub


Sub DTJ2_dropped
	playsoundAtVol SoundFXDOF("drop1",116,DOFPulse,DOFContactors), ActiveBall, 1
	addscore 500
	CheckAwards
End Sub

Sub DTQ1_dropped
	playsoundAtVol SoundFXDOF("drop1",116,DOFPulse,DOFContactors), ActiveBall, 1
	addscore 500
	CheckAwards
End Sub

Sub DTQ2_dropped
	playsoundAtVol SoundFXDOF("drop1",116,DOFPulse,DOFContactors), ActiveBall, 1
	Ldtq2.state=1
	addscore 500
	CheckAwards
End Sub

Sub DTQ3_dropped
	playsoundAtVol SoundFXDOF("drop1",116,DOFPulse,DOFContactors), ActiveBall, 1
	Ldtq3.state=1
	addscore 500
	CheckAwards
End Sub

Sub DTK1_dropped
	playsoundAtVol SoundFXDOF("drop1",117,DOFPulse,DOFContactors), ActiveBall, 1
	LdtK1.state=1
	addscore 1000
	CheckAwards
End Sub

Sub DTK2_dropped
	playsoundAtVol SoundFXDOF("drop1",117,DOFPulse,DOFContactors), ActiveBall, 1
	LdtK2.state=1
	addscore 1000
	CheckAwards
End Sub

sub DTK3_dropped
	playsoundAtVol SoundFXDOF("drop1",117,DOFPulse,DOFContactors), ActiveBall, 1
	LdtK3.state=1
	addscore 1000
	CheckAwards
End Sub

Sub DTK4_dropped
	playsoundAtVol SoundFXDOF("drop1",117,DOFPulse,DOFContactors), ActiveBall, 1
	LdtK4.state=1
	addscore 1000
	CheckAwards
End Sub

Sub DTA1_dropped
	playsoundAtVol SoundFXDOF("drop1",114,DOFPulse,DOFContactors), ActiveBall, 1
	Ldta1.state=1
	addscore 1000
	CheckAwards
End Sub

Sub DTA2_dropped
	playsoundAtVol SoundFXDOF("drop1",114,DOFPulse,DOFContactors), ActiveBall, 1
	LdtA2.state=1
	Ldta2b.state=1
	addscore 1000
	CheckAwards
End Sub

Sub DTA3_dropped
	playsoundAtVol SoundFXDOF("drop1",114,DOFPulse,DOFContactors), ActiveBall, 1
	LdtA3.state=1
	LdtA3b.state=1
	addscore 1000
	CheckAwards
End Sub

Sub DTA4_dropped
	playsoundAtVol SoundFXDOF("drop1",114,DOFPulse,DOFContactors), ActiveBall, 1
	LdtA4.state=1
	addscore 1000
	CheckAwards
End Sub

Sub DTJK_dropped
	playsoundAtVol SoundFXDOF("drop1",114,DOFPulse,DOFContactors), ActiveBall, 1
	Ldtjk.state=1
	Ldtjkb.state=1
	addscore 1000
	CheckAwards
End Sub

Sub resetDT_timer

	Select Case DTStep
		Case 1:							'TEN AND JACKS
			DOF 134,2
			PlaySoundAtVol SoundFX("BankReset",DOFContactors), DTJK, 1
			for i = 0 to 2
				DropTargets(i).isdropped=False
			Next
		Case 2:							'QUEENS
			DOF 109,2
			PlaySoundAtVol SoundFX("BankReset",DOFContactors), DTQ3, 1
			for i = 3 to 5
				DropTargets(i).isdropped=False
			next
			for i = 0 to 1
				DTlights(i).state=0
			Next
		Case 3:							'KINGS
			DOF 135,2
			PlaySoundAtVol SoundFX("BankReset",DOFContactors), DTK3, 1
			for i = 6 to 9
				DropTargets(i).isdropped=False
			next
			for i = 2 to 7
				DTlights(i).state=0
			Next
		Case 4:							'ACES AND JOKER
			DOF 107,2
			PlaySoundAtVol SoundFX("BankReset", DOFContactors), DTA4, 1
			for i = 10 to 14
				DropTargets(i).isdropped=False
			next
			for i = 8 to 13
				DTlights(i).state=0
			Next
		Case 5:
			resetDT.enabled=False
	end Select
	DTStep=DTStep+1
end sub



'************ Target


Sub Textraball_Hit()
	DOF 115, DOFPulse
	addscore 500
	if LExtraBall.state=1 then
		ShootAgain.state=1
		If B2SOn then Controller.B2SSetShootAgain 36,1
	end if
End Sub

Sub CheckAwards
	if DropTargets(1).isdropped+DropTargets(2).isdropped=2 then Lbonus2.state=1		'JACKS
	if DropTargets(3).isdropped+DropTargets(4).isdropped+DropTargets(5).isdropped=3 then Lbonus3.state=1	'QUEENS
	if DropTargets(6).isdropped+DropTargets(7).isdropped+DropTargets(8).isdropped+DropTargets(9).isdropped=4 then Lbonus4.state=1	'KINGS
	if DropTargets(10).isdropped+DropTargets(11).isdropped+DropTargets(12).isdropped+DropTargets(13).isdropped+DropTargets(14).isdropped=5 then 	'ACES AND JOKER
		LExtraBall.state=1
		if balls=3 then Lspecial.state=1
		Lbonus5.state=1
	end if
	If LtopA.state + LtopB.state + LtopC.state= 0 Then
		LExtraBall.state= 1
		Ldoublebonus.state= 1
		if balls=3 then Lspecial.state=1
	End If
End sub


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
	EVAL("ScoreReel"&player).addvalue(points)
	If B2SOn Then Controller.B2SSetScorePlayer player, score(player) MOD 100000

    ' Sounds: there are 3 sounds: tens, hundreds and thousands
    If Points = 100 AND(Score(player) MOD 1000) \ 100 = 0 Then  'New 1000 reel
        PlaySound SoundFXDOF("bell1000",133,DOFPulse,DOFChimes)
      ElseIf Points = 10 AND(Score(player) MOD 100) \ 10 = 0 Then 'New 100 reel
        PlaySound SoundFXDOF("bell100",132,DOFPulse,DOFChimes)
      ElseIf points = 1000 Then
        PlaySound SoundFXDOF("bell1000",133,DOFPulse,DOFChimes)
	  elseif Points = 100 Then
        PlaySound SoundFXDOF("bell100",132,DOFPulse,DOFChimes)
      Else
        PlaySound SoundFXDOF("bell10",131,DOFPulse,DOFChimes)
    End If
    ' check replays and rollover
	if score(player)=>99999 then
		If B2SOn then Controller.B2SSetScoreRollover 24 + player, 1
		EVAL("p100k"&player).text="100,000"
	End if
    if score(player)=>replay1 and rep(player)=0 then
		addcredit
		rep(player)=1
		PlaySound SoundFXDOF("knock",111,DOFPulse,DOFKnocker)
		DOF 112, DOFPulse
    end if
    if score(player)=>replay2 and rep(player)=1 then
		addcredit
		rep(player)=2
		PlaySound SoundFXDOF("knock",111,DOFPulse,DOFKnocker)
		DOF 112, DOFPulse
    end if
    if score(player)=>replay3 and rep(player)=2 then
		addcredit
		rep(player)=3
		PlaySound SoundFXDOF("knock",111,DOFPulse,DOFKnocker)
		DOF 112, DOFPulse
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
	playsound "tilt"
	turnoff
End Sub

sub turnoff
    bumper1.hashitevent=0
    bumper2.hashitevent=0
	LeftSlingShot.isdropped=true
	LeftFlipper.RotateToStart
	DOF 101,DOFOff
	StopSound "Buzz"
	RightFlipper.RotateToStart
	Rightflip1.RotateToStart
	Lupflip.state=0
	DOF 102,DOFOff
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
	PlaySound "pinhit_low", 0, Vol(ActiveBall)*VolPi, AudioPan(ActiveBall), 0, Pitch(ActiveBall), 0, 0, AudioFade(ActiveBall)
End Sub

Sub Targets_Hit (idx)
	PlaySound "target", 0, Vol(ActiveBall)*VolTarg, AudioPan(ActiveBall), 0, Pitch(ActiveBall), 0, 0, AudioFade(ActiveBall)
End Sub

Sub Metals_Thin_Hit (idx)
	PlaySound "metalhit_thin", 0, Vol(ActiveBall)*VolMetal, AudioPan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
End Sub

Sub Metals_Medium_Hit (idx)
	PlaySound "metalhit_medium", 0, Vol(ActiveBall)*VolMetal, AudioPan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
End Sub

Sub Metals2_Hit (idx)
	PlaySound "metalhit2", 0, Vol(ActiveBall)*VolMetal, AudioPan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
End Sub

Sub Gates_Hit (idx)
	PlaySound "gate4", 0, Vol(ActiveBall)*VolGates, AudioPan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
End Sub

Sub Rubbers_Hit(idx)
 	dim finalspeed
  	finalspeed=SQR(activeball.velx * activeball.velx + activeball.vely * activeball.vely)
 	If finalspeed > 20 then
		PlaySound "fx_rubber2", 0, Vol(ActiveBall)*VolRH, Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
	End if
	If finalspeed >= 6 AND finalspeed <= 20 then
 		RandomSoundRubber()
 	End If
End Sub

Sub RubberWheel_Hit()
 	dim finalspeed
  	finalspeed=SQR(activeball.velx * activeball.velx + activeball.vely * activeball.vely)
 	If finalspeed > 20 then
		PlaySound "fx_rubber2", 0, Vol(ActiveBall)*VolPH, Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
	End if
	If finalspeed >= 6 AND finalspeed <= 20 then
 		RandomSoundRubber()
 	End If
End Sub

Sub Posts_Hit(idx)
 	dim finalspeed
  	finalspeed=SQR(activeball.velx * activeball.velx + activeball.vely * activeball.vely)
 	If finalspeed > 16 then
		PlaySound "fx_rubber2", 0, Vol(ActiveBall)*VolPo, Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
	End if
	If finalspeed >= 6 AND finalspeed <= 16 then
 		RandomSoundRubber()
 	End If
End Sub

Sub RandomSoundRubber()
	Select Case Int(Rnd*3)+1
		Case 1 : PlaySound "rubber_hit_1", 0, Vol(ActiveBall)*VolRH, Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
		Case 2 : PlaySound "rubber_hit_2", 0, Vol(ActiveBall)*VolRH, Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
		Case 3 : PlaySound "rubber_hit_3", 0, Vol(ActiveBall)*VolRH, Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
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
		Case 1 : PlaySound "flip_hit_1", 0, Vol(ActiveBall)*VolRh, Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
		Case 2 : PlaySound "flip_hit_2", 0, Vol(ActiveBall)*VolRh, Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
		Case 3 : PlaySound "flip_hit_3", 0, Vol(ActiveBall)*VolRh, Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
	End Select
End Sub

sub savehs
    savevalue "JokerPoker", "credit", credit
    savevalue "JokerPoker", "hiscore", hisc
    savevalue "JokerPoker", "match", matchnumb
    savevalue "JokerPoker", "score1", score(1)
    savevalue "JokerPoker", "score2", score(2)
    savevalue "JokerPoker", "score3", score(3)
    savevalue "JokerPoker", "score4", score(4)
	savevalue "JokerPoker", "balls", balls
	savevalue "JokerPoker", "hsa1", HSA1
	savevalue "JokerPoker", "hsa2", HSA2
	savevalue "JokerPoker", "hsa3", HSA3
	savevalue "JokerPoker", "freeplay", freeplay
end sub

sub loadhs
    dim temp
	temp = LoadValue("JokerPoker", "credit")
    If (temp <> "") then credit = CDbl(temp)
    temp = LoadValue("JokerPoker", "hiscore")
    If (temp <> "") then hisc = CDbl(temp)
    temp = LoadValue("JokerPoker", "match")
    If (temp <> "") then matchnumb = CDbl(temp)
    temp = LoadValue("JokerPoker", "score1")
    If (temp <> "") then score(1) = CDbl(temp)
    temp = LoadValue("JokerPoker", "score2")
    If (temp <> "") then score(2) = CDbl(temp)
    temp = LoadValue("JokerPoker", "score3")
    If (temp <> "") then score(3) = CDbl(temp)
    temp = LoadValue("JokerPoker", "score4")
    If (temp <> "") then score(4) = CDbl(temp)
    temp = LoadValue("JokerPoker", "balls")
    If (temp <> "") then balls = CDbl(temp)
    temp = LoadValue("JokerPoker", "hsa1")
    If (temp <> "") then HSA1 = CDbl(temp)
    temp = LoadValue("JokerPoker", "hsa2")
    If (temp <> "") then HSA2 = CDbl(temp)
    temp = LoadValue("JokerPoker", "hsa3")
    If (temp <> "") then HSA3 = CDbl(temp)
    temp = LoadValue("JokerPoker", "freeplay")
    If (temp <> "") then freeplay = CDbl(temp)
end sub



Sub JokerPoker_Exit()
	Savehs
	turnoff
	If B2SOn Then Controller.stop
End Sub

'==========================================================================================================================================
'============================================================= START OF HIGH SCORES ROUTINES =============================================================
'==========================================================================================================================================
'
'ADD LINE TO TABLE_KEYDOWN SUB WITH THE FOLLOWING:    If HSEnterMode Then HighScoreProcessKey(keycode) AFTER THE STARTGAME ENTRY
'ADD And Not HSEnterMode=true TO IF KEYCODE=STARTGAMEKEY
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

Sub PlaySoundAtVol(sound, tableobj, Volum)
  PlaySound sound, 1, Volum, Pan(tableobj), 0,0,0, 1, AudioFade(tableobj)
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

Function AudioFade(tableobj) ' Fades between front and back of the table (for surround systems or 2x2 speakers, etc), depending on the Y position on the table. "JokerPoker" is the name of the table
  Dim tmp
  tmp = tableobj.y * 2 / JokerPoker.height-1
  If tmp > 0 Then
    AudioFade = Csng(tmp ^10)
  Else
    AudioFade = Csng(-((- tmp) ^10) )
  End If
End Function

Function AudioPan(tableobj) ' Calculates the pan for a tableobj based on the X position on the table. "JokerPoker" is the name of the table
  Dim tmp
  tmp = tableobj.x * 2 / JokerPoker.width-1
  If tmp > 0 Then
    AudioPan = Csng(tmp ^10)
  Else
    AudioPan = Csng(-((- tmp) ^10) )
  End If
End Function

Function Pan(ball) ' Calculates the pan for a ball based on the X position on the table. "JokerPoker" is the name of the table
    Dim tmp
    tmp = ball.x * 2 / JokerPoker.width-1
    If tmp > 0 Then
        Pan = Csng(tmp ^10)
    Else
        Pan = Csng(-((- tmp) ^10) )
    End If
End Function

Function AudioFade(ball) ' Can this be together with the above function ?
  Dim tmp
  tmp = ball.y * 2 / JokerPoker.height-1
  If tmp > 0 Then
    AudioFade = Csng(tmp ^10)
  Else
    AudioFade = Csng(-((- tmp) ^10) )
  End If
End Function

Function Vol(ball) ' Calculates the Volume of the sound based on the ball speed
  Vol = Csng(BallVel(ball) ^2 / VolDiv)
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
    PlaySound("fx_collide"), 0, Csng(velocity) ^2 / (VolDiv/VolCol), Pan(ball1), 0, Pitch(ball1), 0, 0, AudioFade(ball1)
End Sub

' Thalamus : Exit in a clean and proper way
Sub JokerPoker_exit()
  Controller.Pause = False
  Controller.Stop
End Sub


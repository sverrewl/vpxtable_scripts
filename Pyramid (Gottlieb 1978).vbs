'	-------------------------------------------------
'	PYRAMID Gottlieb 1978
'
'	vp8 by jpsalas, 2006
'	vp10 by BorgDog, 2015
'
'	-------------------------------------------------
Option Explicit
Randomize


' Thalamus 2018-07-24
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
Const VolSpin   = 1.5  ' Spinners volume.
Const VolFlip   = 1    ' Flipper volume.

Const cGameName = "pyramid_1978"

Dim CleanOrDirty
Dim operatormenu, Options
Dim bumperlitscore
Dim bumperoffscore
Dim bonuscount
Dim DBonus
Dim balls
Dim replays
Dim Replay1Table(3), Replay2Table(3), Replay3Table(3)
dim replay1, replay2, replay3
Dim Add10, Add100, Add1000
dim replaytext(3)
Dim hisc
Dim maxplayers, players, player
Dim credit
Dim score(2)
Dim sreels(2)
Dim target(5)
Dim DTprim(5)
Dim state
Dim tilt
Dim tiltsens
Dim ballinplay
Dim matchnumb
dim ballrenabled
dim rstep, lstep
Dim rep(2)
Dim rst
Dim eg
Dim scn
Dim scn1
Dim bell
Dim ExtraBallOK, SpecialOK, CenterOK
Dim rollovers, targets
Dim i,j, ii, objekt, light

On Error Resume Next
ExecuteGlobal GetTextFile("controller.vbs")
If Err Then MsgBox "Can't open controller.vbs"
On Error Goto 0

sub Pyramid_init
	P100kreel1.setvalue 2
	P100kreel2.setvalue 3
	LoadEM
	for each objekt in DropTargets:objekt.isdropped=true:next
	maxplayers=2
	Replay1Table(1)=90000
	Replay2Table(1)=120000
	Replay3Table(1)=180000
	Replay1Table(2)=110000
	Replay2Table(2)=130000
	Replay3Table(2)=190000
	Replay1Table(3)=120000
	Replay2Table(3)=150000
	Replay3Table(3)=190000
	set sreels(1) = ScoreReel1
	set sreels(2) = ScoreReel2
	set target(1)=DTyellow
	set target(2)=DTwhite
	set target(3)=DTred
	set target(4)=DTblue
	set target(5)=DTgreen
	hideoptions
	player=1
	For each light in BumperLights:light.State = 1: Next
	For each light in LaneLights:light.State = 1: Next
	For each light in lights:light.State = 0: Next
	loadhs
	if CleanOrDirty = "" then CleanOrDirty=1
	if CleanOrDirty = 0 then SetClean
	if CleanOrDirty = 1 then SetDirty
	if hisc="" then hisc=90000
	hstxt.text=hisc
	gamov.text="Game Over"
	gamov.timerenabled=1
	tilttxt.timerenabled=1
	credittxt.setvalue(credit)
	if balls="" then balls=3
	if balls<>3 and balls<>5 then balls=3
	if replays="" then replays=1
	if replays<>1 and replays<>2 and replays<>3 then replays=1
	Replay1=Replay1Table(Replays)
	Replay2=Replay2Table(Replays)
	Replay3=Replay3Table(Replays)
	RepCard.image = "ReplayCard"&replays
	OptionBalls.image="OptionsBalls"&Balls
	OptionReplays.image="OptionsReplays"&replays
	bumperlitscore=100
	bumperoffscore=10
	if balls=3 then
		InstCard.image="InstCard3balls"
	  else
		InstCard.image="InstCard5balls"
	end if
	If B2SOn then
		setBackglass.enabled=true
		for each objekt in backdropstuff : objekt.visible = 0 : next
	End If
	if matchnumb="" or matchnumb<10 or matchnumb>100 then matchnumb=100
	if matchnumb=100 then
		matchtxt.text="00"
	  else
		matchtxt.text=matchnumb
	end if
	for i = 1 to maxplayers
		sreels(i).setvalue(score(i))
		EVAL("p100kreel"&i).setvalue (int(score(i)/100000))
	next
	PlaySound "motor"
    Drain.CreateBall
	tilt=false
	If credit > 0 Then DOF 126, DOFOn
End sub

sub setBackglass_timer
	Controller.B2ssetCredits Credit
	Controller.B2SSetData 91,1
	Controller.B2ssetMatch 34, Matchnumb
    Controller.B2SSetGameOver 35,1
    Controller.B2SSetScorePlayer 5, hisc
    Controller.B2SSetScorePlayer 1, Score(1) MOD 100000
    Controller.B2SSetScorePlayer 2, Score(2) MOD 100000
	me.enabled=false
end sub

sub setDirty
	CleanOrDirty=1
	for each objekt In Plastics: objekt.image = "PyramidPlasticsWORN": next
	DirtyPlayfield.visible = true
	Pcover.image = "oldPlungercover"
	Lflip.image = "flipper_dirty_left"
	Rflip.image = "flipper_dirty_right"
	Tbonus.image = "dirty5000wl"
	for each objekt In DirtyLights: objekt.visible = true: next
	for each objekt In CleanLights: objekt.visible = false: next
	OptionCorD.image = "OptionsCorD1"
	Pyramid.ColorGradeImage="ColorGradeLUT256x16_Sepia"
	ScoreReel1.image = "ReelTape"
	ScoreReel2.image = "ReelTape"
End sub

sub setClean
	CleanOrDirty=0
	for each objekt In Plastics: objekt.image = "PyramidPlasticsCLEAN": next
	DirtyPlayfield.visible = false
	Pcover.image = "Plungercover"
	Lflip.image = "flipper_gottlieb_left"
	Rflip.image = "flipper_gottlieb_right"
	Tbonus.image = "clean5000wl"
	for each objekt In DirtyLights: objekt.visible = false: next
	for each objekt In CleanLights: objekt.visible = true: next
	OptionCorD.image = "OptionsCorD0"
	Pyramid.ColorGradeImage="ColorGradeLUT256x16_1to1"
	ScoreReel1.image = "ReelTapeC"
	ScoreReel2.image = "ReelTapeC"
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

Sub Pyramid_KeyDown(ByVal keycode)

	if keycode=AddCreditKey then
		playsoundAtVol "coinin", drain, 1
		coindelay.enabled=true
    end if

    if keycode=StartGameKey and credit>0 then
	  if state=false then
		credit=credit-1
		If credit < 1 Then DOF 126, DOFOff
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
			Controller.B2SSetScorePlayer 5, hisc
			Controller.B2SSetData 91,0
		End If
	    pup1.setvalue(1)
		tilt=false
		state=true
		playsound "initialize"
		players=1
		canplayreel.setvalue(players)
		rst=0
		resetDT
		resettimer.enabled=true
	  else if state=true and players < maxplayers and Ballinplay=1 then
		credit=credit-1
		DOF 126, DOFOff
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

	If keycode = PlungerKey Then
		Plunger.PullBack
		PlaySoundAtVol "plungerpull", plunger, 1
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
			Replays=Replays+1
			if Replays>3 then
				Replays=1
			end if
			Replay1=Replay1Table(Replays)
			Replay2=Replay2Table(Replays)
			Replay3=Replay3Table(Replays)
			OptionReplays.image = "OptionsReplays"&replays
			repcard.image = "replaycard"&replays
		Case 3:
			if CleanOrDirty=0 then
				setDirty
			else
				setClean
			end if
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
		PlaySoundAtVol SoundFXDOF("flipperup",102,DOFOn,DOFContactors), RightFlipper, VolFlip
		PlaySoundAtVol "Buzz1", RightFlipper, VolFlip
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
	OptionCorD.visible = True
    OptionReplays.visible = True
End Sub

Sub HideOptions
	for each objekt In OptionMenu
		objekt.visible = false
	next
End Sub


Sub Pyramid_KeyUp(ByVal keycode)

	If keycode = PlungerKey Then
		Plunger.Fire
		PlaySoundAtVol "plunger", Plunger, 1
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
		PlaySoundAtVol SoundFXDOF("flipperdown",102,DOFOff,DOFContactors), RightFlipper, VolFlip
		StopSound "Buzz1"
	End If
   End if
End Sub

sub flippertimer_timer()
	LFlip.RotY = LeftFlipper.CurrentAngle
	RFlip.RotY = RightFlipper.CurrentAngle
	Pgate2.RotX = Gate2.CurrentAngle+110
	Pgate1.RotX = Gate1.CurrentAngle+110
	Pgate.rotz = Gate.currentangle+25
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
	P100kreel1.setvalue 0
	p100kreel2.setvalue 0
    If B2SOn then
		for i = 1 to maxplayers
		  Controller.B2SSetScorePlayer i, score(i) MOD 100000
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
	  DOF 126, DOFOn
      if credit>15 then credit=15
	  credittxt.setvalue(credit)
	  If B2SOn Then Controller.B2ssetCredits Credit
      playsound "click"
End sub

Sub Drain_Hit()
	if state=true Then
		PlaySoundAtVol "drain", drain, 1
		me.timerenabled=1
		DOF 125, DOFPulse
	end if
End Sub

Sub Drain_timer
	bonuscount=0
	dbonus=1	'***** normally 1, use 2 to test bonus scoring routine
'	for each light in BonusLights:light.state=1: next    '***** for testing bonus scoring routine
	if DbonusL.state + DbonusR.state = 2 then dbonus=2
	scorebonus.enabled=true
	me.timerenabled=0
End Sub

sub ballhome_hit
	ballrenabled=1
end sub

sub ballhome_unhit
	DOF 113, DOFPulse
end sub

sub ballrel_hit
	if ballrenabled=1 then
		shootagain.state=lightstateoff
		If B2SOn then Controller.B2SSetShootAgain 36,0
		ballrenabled=0
	end if
end sub

sub scorebonus_timer
	if tilt=false then
		bonuscount=bonuscount+1
	 Else
		bonuscount=6
	end if
	select case(bonuscount)
		case 1:
		  if AddScore1000Timer.enabled = false then
			playsound "motorshorter"
			LbonusGA.state=1
			if LbonusGL.state + LbonusGR.state=2 then addscore 1000*dbonus
			me.interval=350
		  else
			bonuscount=bonuscount-1
			me.timerinterval=100
		  end if
		case 2:
		  if AddScore1000Timer.enabled = false then
			  playsound "motorshorter"
			  LbonusGA.state=0
			  LbonusYA.state=1
			  if LbonusYL.state + LbonusYR.state=2 then addscore 2000*dbonus
			  me.interval=350
		  else
			bonuscount=bonuscount-1
			me.interval=100
		  end if
		case 3:
		  if AddScore1000Timer.enabled = false then
			  playsound "motorshorter"
			  LbonusYA.state=0
			  LbonusBA.state=1
			  if LbonusBL.state + LbonusBR.state=2 then addscore 3000*dbonus
			  me.interval=350
		  else
			bonuscount=bonuscount-1
			me.interval=100
		  end if
		case 4:
		  if AddScore1000Timer.enabled = false then
			  playsound "motorshorter"
			  LbonusBA.state=0
			  LbonusWA.state=1
			  if LbonusWL.state + LbonusWR.state=2 then addscore 4000*dbonus
			  me.interval=350
		  else
			bonuscount=bonuscount-1
			me.interval=100
		  end if
		case 5:
		  if AddScore1000Timer.enabled = false then
			  playsound "motorshorter"
			  LbonusWA.state=0
			  LbonusRA.state=1
			  if LbonusRL.state + LbonusRR.state=2 then addscore 5000*dbonus
			  me.interval=350
		  else
			bonuscount=bonuscount-1
			me.interval=100
		  end if
		case 6:
		  if AddScore1000Timer.enabled = false then
			  if shootagain.state=lightstateon and tilt=false then
				newball
				ballreltimer.enabled=true
			  else
			   if players=1 or player=players then
				 player=1
				 If B2SOn then Controller.B2ssetplayerup 30, player
				 pup1.setvalue(1)
				 pup2.setvalue(0)
				 nextball
			   else
				player=player+1
				If B2SOn then Controller.B2ssetplayerup 30, player
				pup2.setvalue(1)
				pup1.setvalue(0)
				nextball
			   end if
			 end if
			 me.interval=350
			 scorebonus.enabled=false
		  else
			bonuscount=bonuscount-1
			me.interval=100
		  end if
	end select
End sub


sub newgame
	bumper1.hashitevent=true
	bumper2.hashitevent=true
	bumper3.hashitevent=true
	ExtraBallOK=0
	SpecialOK=0
	CenterOK=0
	resetDT
	player=1
    pup1.setvalue(1)
	for i = 1 to maxplayers
	    score(i)=0
		rep(i)=0
	next
	If B2SOn then
	  for i = 1 to maxplayers
		Controller.B2SSetScorePlayer i, score(i) MOD 100000
	  next
	End If
    eg=0
	shootagain.state=lightstateoff
	for each light in lanelights:light.state=1:next
	for each light in lights:light.state=1: next
	for each light in BonusLights:light.state=0: next
	bumper1light.state=0: Bumper2Light.state=1: Bumper3Light.state=0
	shootagain.state=0
    tilttxt.text=" "
	gamov.text=" "
    If B2SOn then
		Controller.B2SSetGameOver 35,0
		Controller.B2SSetTilt 33,0
		Controller.B2SSetMatch 34,0
	End If
    biptext.text="1"
	matchtxt.text=" "
   	Drain.kick 60,25,0
    PlaySoundAtVol SoundFXDOF("ballrelease",128,DOFPulse,DOFContactors), drain, 1
end sub


sub newball
	bumper1light.state=0: Bumper3Light.state=0
	for each light in lights:light.state=1:next
	for each light in BonusLights:light.state=0: next
	if ballinplay = balls then
		DBonusL.state = 1
		DBonusR.state = 1
	end if
	ExtraBallOK=0
	SpecialOK=0
	CenterOK=0
	resetDT
End Sub


sub nextball
    if tilt=true then
	  bumper1.hashitevent=true
	  bumper2.hashitevent=true
	  bumper3.hashitevent=true
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
	  canplayreel.setvalue(0)
	  pup1.setvalue(0)
	  pup2.setvalue(0)
	  for i=1 to maxplayers
		if score(i)>hisc then hisc=score(i)
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
		LeftFlipper.RotateToStart
		StopSound "Buzz"
		RightFlipper.RotateToStart
		StopSound "Buzz1"
	  For each light in LaneLights:light.State = 0: Next
	  ballreltimer.enabled=false
  else
	Drain.kick 60,25,0
    PlaySoundAtVol SoundFXDOF("ballrelease",128,DOFPulse,DOFContactors), drain, 1
    ballreltimer.enabled=false
  end if
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
		  PlaySound SoundFXDOF("knock",134,DOFPulse,DOFKnocker)
		  DOF 135, DOFPulse
	    end if
    next
end sub

'********** Bumpers

Sub Bumper1_Hit     'left bumper
   if tilt=false then
    PlaySoundAtVol SoundFXDOF("fx_bumper4",107,DOFPulse,DOFContactors), bumper1, volbump
	DOF 108, DOFPulse
	FlashBumpers
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


Sub Bumper2_Hit    'center bumper
   if tilt=false then
    PlaySoundAtVol SoundFXDOF("fx_bumper4",109,DOFPulse,DOFContactors), bumper2, volbump
	DOF 110, DOFPulse
	FlashBumpers
	if balls = 3 then
		addscore 1000
	  else
		addscore 100
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

Sub Bumper3_Hit      'right bumper
   if tilt=false then
    PlaySoundAtVol SoundFXDOF("fx_bumper4",111,DOFPulse,DOFContactors), bumper3, volbump
	DOF 112, DOFPulse
	FlashBumpers
	if bumper3light.state=1 then
		addscore bumperlitscore
	  else
		addscore bumperoffscore
	end if
	me.timerenabled=1
   end if
End Sub

Sub Bumper3_timer
	Bumper3Ring.enabled=0
	BumperRing3.transz=BumperRing3.transz-4
	if BumperRing3.transz=-36 then
		Bumper3Ring.enabled=1
		me.timerenabled=0
	end if
End Sub

Sub Bumper3Ring_timer
	BumperRing3.transz=BumperRing3.transz+4
	If BumperRing3.transz=0 then Bumper3Ring.enabled=0
End sub

Sub PairedlampTimer_timer
	bumper1light1.state=bumper1light.state
	bumper1lightC.state=bumper1light.state
	bumper2light1.state=bumper2light.state
	bumper2lightC.state=bumper2light.state
	bumper3light1.state=bumper3light.state
	bumper3lightC.state=bumper3light.state
end sub

sub FlashBumpers
	if FlashLightSeq.timerenabled=1 then Exit sub
	FlashLightSeq.Play SeqAlloff
	FlashLightSeq.timerenabled=1
end sub

sub FlashLightSeq_timer
	FlashLightSeq.StopPlay
	me.timerenabled=0
end sub

'************** Slings

Sub RightSlingShot_Slingshot
    PlaySoundAtVol SoundFXDOF("slingshot",104,DOFPulse,DOFContactors), slingr, 1
	DOF 106, DOFPulse
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
        Case 4:slingR.objroty = 0:RSLing2.Visible = 0:RSLing.Visible = 1:RightSlingShot.TimerEnabled = 0
    End Select
    RStep = RStep + 1
End Sub

Sub LeftSlingShot_Slingshot
    PlaySoundAtVol SoundFXDOF("slingshot",103,DOFPulse,DOFContactors), slingl, 1
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
	RotateSpecials
End Sub


'********** Triggers

sub TGoutL_hit
	DOF 118, DOFPulse
	addscore 500
	LbonusYL.state=1
	LoutL.state=0
	LtopYellow.state=0
	CheckAwards
end sub

sub TGinL_hit
	DOF 119, DOFPulse
	addscore 500
	LbonusWL.state=1
	LinL.state=0
	LtopWhite.state=0
	CheckAwards
end sub

sub TGoutR_hit
	DOF 122, DOFPulse
	addscore 500
	LbonusGL.state=1
	LoutR.state=0
	LtopGreen.state=0
	CheckAwards
end sub

sub TGinR_hit
	DOF 121, DOFPulse
	addscore 500
	LbonusBL.state=1
	LinR.state=0
	LtopBlue.state=0
	CheckAwards
end sub

sub TGtopYellow_hit
	DOF 118, DOFPulse
	addscore 500
	LtopYellow.state=0
	LoutL.state=0
	LbonusYL.state=1
	CheckAwards
end sub

sub TGtopWhite_hit
	DOF 119, DOFPulse
	addscore 500
	LtopWhite.state=0
	LinL.state=0
	LbonusWL.state=1
	CheckAwards
end sub

sub TGtopRed_hit
	DOF 120, DOFPulse
	addscore 500
	LtopRed.state=0
	LbonusRL.state=1
	CheckAwards
end sub

sub TGtopBlue_hit()
	DOF 121, DOFPulse
	addscore 500
	LtopBlue.state=0
	LinR.state=0
	LbonusBL.state=1
	CheckAwards
end sub

sub TGtopGreen_hit()
	DOF 122, DOFPulse
	addscore 500
	LtopGreen.state=0
	LoutR.state=0
	LbonusGL.state=1
	CheckAwards
end sub

sub TGstarleft_hit()
	DOF 123, DOFPulse
	if LTGstarleft.state=1 then
		addscore 5000
	  else
		addscore 500
	end if
end sub

sub TGstarright_hit()
	DOF 124, DOFPulse
	if LTGstarright.state=1 then
		addscore 5000
	  else
		addscore 500
	end if
end sub

'********** Drop targets

Sub DTyellow_dropped()
	PlaySoundAtVol SoundFXDOF("drop1",116,DOFPulse,DOFContactors), DTyellow, 1
	LbonusYR.state=1
	ScoreDT
End Sub


Sub DTwhite_dropped()
	PlaySoundAtVol SoundFXDOF("drop1",116,DOFPulse,DOFContactors), DTwhite, 1
	LbonusWR.state=1
	ScoreDT
End Sub


Sub DTred_dropped()
	PlaySoundAtVol SoundFXDOF("drop1",116,DOFPulse,DOFContactors), DTred, 1
	LbonusRR.state=1
	ScoreDT
End Sub

Sub DTblue_dropped()
	PlaySoundAtVol SoundFXDOF("drop1",116,DOFPulse,DOFContactors), DTblue, 1
	LbonusBR.state=1
	ScoreDT
End Sub

Sub DTgreen_dropped()
	PlaySoundAtVol SoundFXDOF("drop1",116,DOFPulse,DOFContactors), DTgreen, 1
	LbonusGR.state=1
	ScoreDT
End Sub

sub ScoreDT
	CheckAwards
	if LbonusGL.state=1 then addscore 1000
	if LbonusYL.state=1 then addscore 1000
	if LbonusBL.state=1 then addscore 1000
	if LbonusWL.state=1 then addscore 1000
	if LbonusRL.state=1 then addscore 1000
end sub

Sub resetDT
    PlaySoundAtVol SoundFXDOF("BankReset",117,DOFPulse,DOFContactors), DTred, 1
	for each objekt in DropTargets:objekt.isdropped=false:next
end sub



'************ Targets

Sub TBonus_Hit()
	DOF 117, DOFPulse
	if LTBonus.state=1 then
		addscore 5000
		DOF 136, DOFPulse
	  else
		addscore 500
	end if
End Sub

Sub Tleft_Hit()
	addscore 500
	if SpecialL.state=1 then
		addcredit
		PlaySound SoundFXDOF("knock",134,DOFPulse,DOFKnocker)
		DOF 135, DOFPulse
	end if
	if ExtraBallL.state=1 then
		ShootAgain.state=1
		If B2SOn then Controller.B2SSetShootAgain 36,1
	end if
End Sub

Sub Tright_Hit()
	addscore 500
	if SpecialR.state=1 then
		addcredit
		PlaySound SoundFXDOF("knock",134,DOFPulse,DOFKnocker)
		DOF 135, DOFPulse
	end if
	if ExtraBallR.state=1 then
		ShootAgain.state=1
		If B2SOn then Controller.B2SSetShootAgain 36,1
	end if
End Sub

'************ Kickers

Sub KickerL_Hit()
	if balls = 3 then
		addscore 5000
	  else
		addscore 3000
	end if
	DBonusL.state=1
	AddBonusLight
	me.timerenabled=1
End Sub

Sub KickerL_Timer():
	KickerL.kick 70, 20
    PlaySoundAtVol SoundFXDOF("holekick",114,DOFPulse,DOFContactors), kickerL, 1
	DOF 135, DOFPulse
	Me.TimerEnabled= 0
End Sub

Sub KickerR_Hit()
	if balls = 3 then
		addscore 5000
	  else
		addscore 3000
	end if
	DBonusR.state=1
	AddBonusLight
	me.timerenabled=1
End Sub

Sub KickerR_Timer():
	KickerR.kick -70, 20
    PlaySoundAtVol SoundFXDOF("holekick",115,DOFPulse,DOFContactors), kickerR, 1
	DOF 135, DOFPulse
	Me.TimerEnabled= 0
End Sub

Sub AddBonusLight
     ii= INT(RND*5)
     Select Case ii
         Case 0: LBonusRL.state= 2: LBonusRL.TimerEnabled= 1
         Case 1: LBonusWL.state= 2: LBonusWL.TimerEnabled= 1
         Case 2: LBonusBL.state= 2: LBonusBL.TimerEnabled= 1
         Case 3: LBonusYL.state= 2: LBonusYL.TimerEnabled= 1
         Case 4: LBonusGL.state= 2: LBonusGL.TimerEnabled= 1
     End Select
End Sub

Sub LBonusRL_timer(): Me.TimerEnabled= 0: LBonusRL.state= 1: LtopRed.state= 0: CheckAwards: End Sub
Sub LBonusWL_timer(): Me.TimerEnabled= 0: LBonusWL.state= 1: LtopWhite.state= 0: LinL.state= 0: CheckAwards: End Sub
Sub LBonusBL_timer(): Me.TimerEnabled= 0: LBonusBL.state= 1: LtopBlue.state= 0: LinR.state= 0: CheckAwards: End Sub
Sub LBonusYL_timer(): Me.TimerEnabled= 0: LBonusYL.state= 1: LtopYellow.state= 0: LoutL.state= 0: CheckAwards: End Sub
Sub LBonusGL_timer(): Me.TimerEnabled= 0: LBonusGL.state= 1: LtopGreen.state= 0: LoutR.state= 0: CheckAwards: End Sub


Sub CheckAwards
    Rollovers= LbonusGL.state + LbonusBL.state + LbonusYL.state + LbonusRL.state + LbonusWL.state
    Targets= LbonusGR.state + LbonusBR.state + LbonusYR.state + LbonusRR.state + LbonusWR.state
	If Rollovers= 5 Then  Bumper3Light.state= 1
	If Targets= 5 then Bumper1Light.state=1: CenterOK=1: Centertarget
	If Rollovers= 5 Or Targets= 5 Then
		LTGstarleft.state= 1
		LTGstarright.state= 1
		ExtraBallOK=1
		ExtraBallLights
	End If
	If Rollovers+Targets=10 then
		SpecialOK=1
		SpecialLights
	End if
End Sub

Sub RotateSpecials
     CenterTarget
     ExtraBallLights
     SpecialLights
End Sub

Sub Centertarget
     If CenterOk <> 1 then Exit Sub
     If LTBonus.state = 1 then
         LTBonus.state = 0
		 LTBonus1.state = 0
     Else
         LTBonus.state = 1
		 LTbonus1.state = 1
     End If
End Sub

Sub ExtraBallLights
    If ExtraBallOk <> 1 Then Exit Sub
	If balls=3 then
		If ExtraBallL.state= 1 then
			ExtraBallL.state= 0: ExtraBallR.state= 1
		  Else
			ExtraBallL.state= 1: ExtraBallR.state= 0
		End If
	  else
		If INT((score(player) mod 100)/10) = 0 or INT((score(player) mod 100)/10) = 5 then
			If ExtraBallL.state= 1 then
				ExtraBallL.state= 0: ExtraBallR.state= 1
			  Else
				ExtraBallL.state= 1: ExtraBallR.state= 0
			End If
		End if
	End if
End Sub

Sub SpecialLights
     If SpecialOk <> 1 Then Exit Sub
		If ExtraBallL.state=1 then
			SpecialL.state=0: SpecialR.state=1
		  else
			SpecialL.state=1: SpecialR.state=0
		End if
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
		EVAL("p100kreel"&player).setvalue (int(score(player)/100000))
	End if
    if score(player)=>replay1 and rep(player)=0 then
		addcredit
		rep(player)=1
		PlaySound SoundFXDOF("fx_knocker",134,DOFPulse,DOFKnocker)
		DOF 135, DOFPulse
    end if
    if score(player)=>replay2 and rep(player)=1 then
		addcredit
		rep(player)=2
		PlaySound SoundFXDOF("fx_knocker",134,DOFPulse,DOFKnocker)
		DOF 135, DOFPulse
    end if
    if score(player)=>replay3 and rep(player)=2 then
		addcredit
		rep(player)=3
		PlaySound SoundFXDOF("fx_knocker",134,DOFPulse,DOFKnocker)
		DOF 135, DOFPulse
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

Sub MechCheckTilt
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
	DOF 101, DOFOff
	DOF 102, DOFOff
	LeftFlipper.RotateToStart
	StopSound "Buzz"
	RightFlipper.RotateToStart
	StopSound "Buzz1"
    bumper1.hashitevent=False
    bumper2.hashitevent=False
	bumper3.hashitevent=False
end sub

'************************************
' What you need to add to your table
'************************************

' a timer called CollisionTimer. With a fast interval, like from 1 to 10
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

Sub Spinner_Spin
	PlaySoundAtVol "fx_spinner", Spinner, VolSpin
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

Sub RubberWheel_hit
 	dim finalspeed
  	finalspeed=SQR(activeball.velx * activeball.velx + activeball.vely * activeball.vely)
 	If finalspeed > 20 then
		PlaySound "fx_rubber2", 0, Vol(ActiveBall)*VolRH, Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
	End if
	If finalspeed >= 6 AND finalspeed <= 20 then
 		RandomSoundRubber()
 	End If
End sub
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
		Case 1 : PlaySound "flip_hit_1", 0, Vol(ActiveBall)*VolRH, Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
		Case 2 : PlaySound "flip_hit_2", 0, Vol(ActiveBall)*VolRH, Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
		Case 3 : PlaySound "flip_hit_3", 0, Vol(ActiveBall)*VolRH, Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
	End Select
End Sub

sub savehs

    savevalue "Pyramid", "credit", credit
    savevalue "Pyramid", "hiscore", hisc
    savevalue "Pyramid", "match", matchnumb
    savevalue "Pyramid", "score1", score(1)
    savevalue "Pyramid", "score2", score(2)
	savevalue "Pyramid", "balls", balls
	savevalue "Pyramid", "CleanorDirty", CleanOrDirty

end sub

sub loadhs
    dim temp
	temp = LoadValue("Pyramid", "credit")
    If (temp <> "") then credit = CDbl(temp)
    temp = LoadValue("Pyramid", "hiscore")
    If (temp <> "") then hisc = CDbl(temp)
    temp = LoadValue("Pyramid", "match")
    If (temp <> "") then matchnumb = CDbl(temp)
    temp = LoadValue("Pyramid", "score1")
    If (temp <> "") then score(1) = CDbl(temp)
    temp = LoadValue("Pyramid", "score2")
    If (temp <> "") then score(2) = CDbl(temp)
    temp = LoadValue("Pyramid", "balls")
    If (temp <> "") then balls = CDbl(temp)
    temp = LoadValue("Pyramid", "CleanOrDirty")
    If (temp <> "") then CleanOrDirty = CDbl(temp)
end sub



Sub Pyramid_Exit()
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

Function AudioFade(tableobj) ' Fades between front and back of the table (for surround systems or 2x2 speakers, etc), depending on the Y position on the table. "Pyramid" is the name of the table
  Dim tmp
  tmp = tableobj.y * 2 / Pyramid.height-1
  If tmp > 0 Then
    AudioFade = Csng(tmp ^10)
  Else
    AudioFade = Csng(-((- tmp) ^10) )
  End If
End Function

Function AudioPan(tableobj) ' Calculates the pan for a tableobj based on the X position on the table. "Pyramid" is the name of the table
  Dim tmp
  tmp = tableobj.x * 2 / Pyramid.width-1
  If tmp > 0 Then
    AudioPan = Csng(tmp ^10)
  Else
    AudioPan = Csng(-((- tmp) ^10) )
  End If
End Function

Function Pan(ball) ' Calculates the pan for a ball based on the X position on the table. "Pyramid" is the name of the table
    Dim tmp
    tmp = ball.x * 2 / Pyramid.width-1
    If tmp > 0 Then
        Pan = Csng(tmp ^10)
    Else
        Pan = Csng(-((- tmp) ^10) )
    End If
End Function

Function AudioFade(ball) ' Can this be together with the above function ?
  Dim tmp
  tmp = ball.y * 2 / Pyramid.height-1
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
'    JP's VP10 Rolling Sounds
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

Sub CollisionTimer_Timer()
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
Sub Pyramid_exit()
  Controller.Pause = False
  Controller.Stop
End Sub


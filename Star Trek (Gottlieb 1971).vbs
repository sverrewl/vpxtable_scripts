'///////////////////////////////////////////////////////////
'			STAR TREK by Gottlieb (1971)
'
'
'
' vp10 assembled and scripted by BorgDog, 2015
'
' thanks to Plumb for the playfield scan on VPF
'
' DOF by BorgDog, reviewed by Arngrim :)
'
'///////////////////////////////////////////////////////////

' Thalamus 2018-07-24
' Added/Updated "Positional Sound Playback Functions" and "Supporting Ball & Sound Functions"
' No special SSF tweaks yet.

Option Explicit
Randomize

Const cGameName = "startrek_1971"

Dim operatormenu, options
Dim bumperlitscore
Dim bumperoffscore
Dim balls
dim award
Dim replays
Dim Add10, Add100, Add1000
Dim Replay1Table(2)
Dim Replay2Table(2)
dim replay1
dim replay2
Dim hisc
Dim Controller
Dim maxplayers
Dim players
Dim player
Dim credit
Dim score(4)
Dim sreels(4)
Dim cplay(4)
Dim state
Dim tilt
Dim tiltsens
Dim target(9)
Dim DTprim(9)
Dim holebonus
Dim holeb(4)
Dim StarState
Dim Bonus
Dim ballstoplay
dim ballrenabled
dim rlight
Dim rep(2)
Dim rst
Dim eg
Dim bell
Dim i,j, objekt, light
Dim awardcheck
Dim rstep, lstep

On Error Resume Next
ExecuteGlobal GetTextFile("controller.vbs")
If Err Then MsgBox "Can't open controller.vbs"
On Error Goto 0


sub StarTrek_init
	LoadEM
	maxplayers=1
	Replay1Table(1)=20000
	Replay1Table(2)=30000
	Replay2Table(1)=40000
	Replay2Table(2)=50000
	set sreels(1) = ScoreReel1
	hideoptions
	player=1
	For each light in BumperLights:light.State = 1: Next
	For each light in Tlights:light.State = 0: Next
	For each light in starlights:light.State = 0: Next
	btpoff
	loadhs
	if hisc="" then hisc=20000
	hstxt.text=hisc
	btp_go.state=1
	tiltreel.timerenabled=1
	if balls="" then balls=5
	if balls<>3 and balls<>5 and balls<>8 then balls=5
	if replays="" then replays=2
	if replays<>1 and replays<>2 then replays=2
	Replay1=Replay1Table(Replays)
	Replay2=Replay2Table(Replays)
	RepCard.image = "ReplayCard1"
	OptionBalls.image="OptionsBalls"&Balls
	OptionReplays.image="OptionsReplays"&replays
	if balls=3 then
		bumperlitscore=1000
		bumperoffscore=100
		InstCard.image="InstCard3balls"
	end if
	if balls=5 then
		bumperlitscore=1000
		bumperoffscore=100
		InstCard.image="InstCard5balls"
	end if
	if balls=8 then
		bumperlitscore=1000
		bumperoffscore=100
		InstCard.image="InstCard8balls"
	end if
	If B2SOn then
		setBackglass.enabled=true
		for each objekt in backdropstuff : objekt.visible = 0 : next
	End If
	for i = 1 to maxplayers
		sreels(i).setvalue(score(i))
	next
	for i = 1 to 10: EVAL("rollover"&i).visible=0: next
	if not b2son then
		if score(player)>1000000 then
			rollover10.state=1
		  Else
			If score(player)>100000 Then EVAL("rollover"&(Int(score(player)/100000))).visible=1
		end if
	end if
	PlaySound "motor"
	tilt=false
	If credit>0 then DOF 136, DOFOn
    Drain.CreateBall
	If nightday<=5 then
		For each light in tlights:light.intensityscale = 2:Next
		For each light in PFLights:light.intensityscale = 3:Next
	End If
	If nightday>5 then
		For each light in tlights:light.intensityscale = 1.5:Next
	End If
	If nightday>10 then
		For each light in tlights:light.intensityscale = .8:Next
		For each light in PFLights:light.intensityscale = 1:Next
	End If
	If nightday>30 then
		For each light in tlights:light.intensityscale = .7:Next
		For each light in bumperlights:light.intensityscale = .6:Next
	End If
	If nightday>80 then
		For each light in tlights:light.intensityscale = .5:Next
		For each light in PFLights:light.intensityscale = .5:Next
	End If
End sub

sub setBackglass_timer
    Controller.B2SSetGameOver 35,1
    Controller.B2SSetScorePlayer 5, hisc
    Controller.B2SSetScorePlayer 1, Score(1) MOD 100000
	if score(player)>1000000 then
		Controller.b2ssetscorerollover player+24, 10
	  Else
		If score(player)>100000 Then Controller.b2ssetscorerollover player+24, Int(score(player)/100000)
	end if
	dim objekt : for each objekt in backdropstuff : objekt.visible = 0 : next
	me.enabled=false
end sub


sub tiltreel_timer
	if state=false then
		tiltreel.visible=0
		If B2SOn then Controller.B2SSetTilt 33,0
		ttimer.enabled=true
	end if
	tiltreel.timerenabled=0
end sub

sub ttimer_timer
	if state=false then
		if not B2Son then tiltreel.visible=1
		If B2SOn then Controller.B2SSetTilt 33,1
		tiltreel.timerinterval= (INT (RND*10)+5)*100
		tiltreel.timerenabled=1
	end if
	me.enabled=0
end sub

Sub StarTrek_KeyDown(ByVal keycode)

	if keycode=AddCreditKey then
		playsound "coin3"
		coindelay.enabled=true
    end if

    if keycode=StartGameKey and credit>0 and state=false then
		credit=credit-1
		if credit < 1 then DOF 136, DOFOff
		ballstoplay=balls
		If B2SOn Then
			Controller.B2ssetplayerup 30, 1
			Controller.B2SSetGameOver 0
			Controller.B2SSetScoreRollover 25, 0
		End If
		tilt=false
		state=true
		playsound "initialize"
		players=1
	    For each light in Tlights:light.State = 1: Next
		rst=0
		resettimer.enabled=true
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
		playsound "target"
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
	  PlaySound "metalhit2"
	  Select Case (Options)
		Case 1:
			if Balls=3 then
				Balls=5
				InstCard.image="InstCard5balls"
			  elseif balls=5 Then
				Balls=8
				InstCard.image="InstCard8balls"
			  else
				Balls=3
				InstCard.image="InstCard3balls"
			end if
			OptionBalls.image = "OptionsBalls"&Balls
		Case 2:
			Replays=Replays+1
			if Replays>2 then
				Replays=1
			end if
			Replay1=Replay1Table(Replays)
			Replay2=Replay2Table(Replays)
			OptionReplays.image = "OptionsReplays"&replays
			repcard.image = "ReplayCard1"
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
		PlaySound "tilt"
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
End Sub

Sub HideOptions
	for each objekt In OptionMenu
		objekt.visible = false
	next
End Sub


Sub StarTrek_KeyUp(ByVal keycode)

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
	LFlip.RotY = LeftFlipper.CurrentAngle-90
	RFlip.RotY = RightFlipper.CurrentAngle-90
	PGate.Rotz = gate.CurrentAngle+25
end sub

Sub PairedlampTimer_timer
	if credit>0 then
		creditlight.state=1
	  else
		creditlight.state=0
	end if
	LStarS2.state = LStarS1.state
	LStarT2.state = LStarT1.state
	LStarA2.state = LStarA1.state
	LStarR2.state = LStarR1.state
	LTrekT2.state = LTrekT1.state
	LTrekR2.state = LTrekR1.state
	LTrekE2.state = LTrekE1.state
	LTrekK2.state = LTrekK1.state
	LGreenStar1.state = LGreenStar.state
	LPurpleStar1.state = LPurpleStar.state
	LRwow.state = LLwow.state
	Bumper1Light1.state = bumper1light.state
	Bumper2Light1.state = bumper2light.state
	Bumper3Light.state = bumper2light.state
	Bumper3Light1.state = bumper3light.state
end sub

sub coindelay_timer
	addcredit
    coindelay.enabled=false
end sub

sub resettimer_timer
    rst=rst+1
	if rst=2 then
		for i = 1 to maxplayers
			sreels(i).resettozero
		next
		for i = 1 to 10: EVAL("rollover"&i).visible=0: next
		If B2SOn then
			for i = 1 to maxplayers
			  Controller.B2SSetScorePlayer i, score(i) MOD 100000
			next
		End If
	end if
    if rst=22 then
		newgame
		resettimer.enabled=false
    end if
end sub

Sub addcredit
      credit=credit+1
	  DOF 136, DOFOn
      if credit>15 then credit=15
End sub

Sub Drain_Hit()
	DOF 132, DOFPulse
	PlaySound "drain",0,1,0,0.25
	me.timerenabled=1
End Sub

Sub Drain_timer
	nextball
	me.timerenabled=0
End Sub

sub ballhome_hit
	ballrenabled=1
end sub

sub ballhome_unhit
	DOF 137, DOFPulse
end sub

sub ballrel_hit
	if ballrenabled=1 then
		ballrenabled=0
	end if
end sub

sub newgame
	bumper1.force=8
	bumper2.force=8
	bumper3.force=8
	player=1
    score(1)=0
	award=0
	If B2SOn then
	  for i = 1 to maxplayers
		Controller.B2SSetScorePlayer i, score(i) MOD 100000
	  next
	End If
    eg=0
    rep(player)=0
	for each light in bumperlights:light.state=0:next
	for each light in tlights:light.state=1:next
	for each light in treklights:light.state=1:next
	for each light in starlights:light.state=0:next
    btp_go.state=0
    tiltreel.visible=0
    If B2SOn then
		Controller.B2SSetGameOver 35,0
		Controller.B2SSetTilt 33,0
	End If
	btpoff
	if ballstoplay>5 Then
		btp5.state=1
	  Else
		EVAL("btp"&ballstoplay).state=1
	end If
   	Drain.kick 60,35,0
	playsound SoundFXDOF("kickerkick",131,DOFPulse,DOFContactors)
end sub

sub btpoff
	for i = 1 to 5
		Eval("btp"&i).state=0
	next
end sub


sub nextball
    if tilt=true then
	  bumper1.force=8
	  bumper2.force=8
	  bumper3.force=8
      tilt=false
      tiltreel.visible=0
		If B2SOn then
			Controller.B2SSetTilt 33,0
			Controller.B2ssetdata 1, 1
		End If
    end if
	if player=1 then ballstoplay=ballstoplay-1
	if ballstoplay=0 then
		playsound "motorleer"
		eg=1
		ballreltimer.enabled=true
	  else
		if state=true and tilt=false then
		  ballreltimer.enabled=true
		end if
	    For each light in Tlights:light.State = 1: Next
		btpoff
		if ballstoplay>5 Then
			btp5.state=1
		  Else
			EVAL("btp"&ballstoplay).state=1
		end If
	end if
End Sub

sub ballreltimer_timer
  if eg=1 then
	  turnoff
	  btpoff
	  For each light in Tlights:light.State = 0: Next
	  state=false
	  btp_go.state=1
	  tiltreel.timerenabled=1
  	  for i = 1 to maxplayers
		if score(i)>hisc then hisc=score(i)
	  next
	  hstxt.text=hisc
	  savehs
	  If B2SOn then
        Controller.B2SSetGameOver 35,1
	    Controller.B2SSetScorePlayer 5, hisc
	    Controller.B2ssetcanplay 31, 0
	    Controller.B2ssetcanplay 30, 0
	  End If
  else
	if award=1 then
		for each light in spotlights:light.state = 1: Next
		for each light in spotoff:light.state = 0: Next
		award = 0
	end if
	Drain.kick 60,35,0
	playsound SoundFXDOF("kickerkick",131,DOFPulse,DOFContactors)
  end if
  ballreltimer.enabled=false
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

Sub Bumper3_Hit
   if tilt=false then
	playsound SoundFXDOF("fx_bumper4",111,DOFPulse,DOFContactors)
	DOF 112,DOFPulse
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


Sub RightSlingShot_Slingshot
	playsound SoundFXDOF("left_slingshot",105,DOFPulse,DOFContactors), 0, 1, 0.05, 0.05
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
        Case 4:slingR.objroty = 0:RSLing2.Visible = 0:RSLing.Visible = 1:RightSlingShot.TimerEnabled = 0
    End Select
    RStep = RStep + 1
End Sub

Sub LeftSlingShot_Slingshot
	playsound SoundFXDOF("right_slingshot",103,DOFPulse,DOFContactors), 0, 1, -0.05, 0.05
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


'****** Triggers & Targets

Sub TStarS1_Hit
   if tilt=false then
	DOF 113,DOFPulse
	addscore 500
	LStarS1.state = 0
	LStarS3.state = 1
   end if
 End Sub

Sub TStarT1_Hit
   if tilt=false then
	DOF 114,DOFPulse
	addscore 500
	LStarT1.state = 0
	LStarT3.state = 1
   end if
End Sub

Sub TStarA1_Hit
   if tilt=false then
	DOF 115,DOFPulse
	addscore 500
	LStarA1.state = 0
	LStarA3.state = 1
   end if
End Sub

Sub TStarR1_Hit
   if tilt=false then
	DOF 116,DOFPulse
	addscore 500
	LStarR1.state = 0
	LStarR3.state = 1
   end if
End Sub

Sub TStarS2_Hit
   if tilt=false then
	DOF 121,DOFPulse
	addscore 500
	LStarS1.state = 0
	LStarS3.state = 1
   end if
End Sub

Sub TStarT2_Hit
   if tilt=false then
	DOF 122,DOFPulse
	addscore 500
	LStarT1.state = 0
	LStarT3.state = 1
   end if
End Sub

Sub TStarA2_Hit
   if tilt=false then
	DOF 124,DOFPulse
	addscore 500
	LStarA1.state = 0
	LStarA3.state = 1
   end if
End Sub

Sub TStarR2_Hit
   if tilt=false then
	DOF 123,DOFPulse
	addscore 500
	LStarR1.state = 0
	LStarR3.state = 1
   end if
End Sub

Sub TTrekT1_Hit
   if tilt=false then
	DOF 125,DOFPulse
	addscore 500
	LTrekT1.state = 0
	LTrekT3.state = 1
   end if
End Sub

Sub TTrekR1_Hit
   if tilt=false then
	DOF 125,DOFPulse
	addscore 500
	LTrekR1.state = 0
	LTrekR3.state = 1
   end if
End Sub

Sub TTrekE1_Hit
   if tilt=false then
	DOF 126,DOFPulse
	addscore 500
	LTrekE1.state = 0
	LTrekE3.state = 1
   end if
End Sub

Sub TTrekK1_Hit
   if tilt=false then
	DOF 126,DOFPulse
	addscore 500
	LTrekK1.state = 0
	LTrekK3.state = 1
   end if
End Sub

Sub TTrekT2_Hit
   if tilt=false then
	DOF 118,DOFPulse
	addscore 500
	LTrekT1.state = 0
	LTrekT3.state = 1
   end if
End Sub

Sub TTrekR2_Hit
   if tilt=false then
	DOF 129,DOFPulse
	addscore 500
	LTrekR1.state = 0
	LTrekR3.state = 1
   end if
End Sub

Sub TTrekE2_Hit
   if tilt=false then
	DOF 130,DOFPulse
	addscore 500
	LTrekE1.state = 0
	LTrekE3.state = 1
   end if
End Sub

Sub TTrekK2_Hit
   if tilt=false then
	DOF 120,DOFPulse
	addscore 500
	LTrekK1.state = 0
	LTrekK3.state = 1
   end if
End Sub

Sub TLwow_Hit
   if tilt=false then
	DOF 127,DOFPulse
	addscore 10
	LGreenStar.state = 0
	LGreenStar3.state = 1
	if LLwow.state=1 then addballs
   end if
End Sub

Sub TLstar_Hit
   if tilt=false then
	DOF 117,DOFPulse
	addscore 10
	LGreenStar.state = 0
	LGreenStar3.state = 1
   end if
End Sub

Sub TRwow_Hit
   if tilt=false then
	DOF 128,DOFPulse
	addscore 10
	LPurpleStar.state = 0
	LPurpleStar3.state = 1
	if LLwow.state=1 then addballs
   end if
End Sub

Sub TRstar_Hit
   if tilt=false then
	DOF 119,DOFPulse
	addscore 10
	LPurpleStar.state = 0
	LPurpleStar3.state = 1
   end if
End Sub

Sub AwardCheckTimer_timer
	Dim starcheck, trekcheck
	starcheck = LStarS3.state + LStarT3.state + LStarA3.state + LStarR3.state
	trekcheck = LTrekT3.state + LTrekR3.state + LTrekE3.state + LTrekK3.state
	if starcheck = 4 then Bumper1Light.state = 1
	if trekcheck = 4 then Bumper2Light.state = 1
	if starcheck + trekcheck + LGreenStar3.state + LPurpleStar3.state = 10 then
		if award=0 then
			addballs
			LLwow.state = 1
			award=1
		end if
	end if
End Sub

sub addballs
	ballstoplay=ballstoplay+1
	if ballstoplay>10 then ballstoplay=10
	btpoff
	if ballstoplay>5 Then
		btp5.state=1
	  Else
		EVAL("btp"&ballstoplay).state=1
	end If
	playsound SoundFXDOF("knocker",133,DOFPulse,DOFKnocker)
	DOF 134,DOFPulse
end sub

sub addscore(points)
  if tilt=false then
	If points=10 or points=100 or points=1000 Then
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
	end If
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
	If B2SOn Then
		Controller.B2SSetScorePlayer player, score(player) MOD 100000
		if score(player)>1000000 then
			Controller.b2ssetscorerollover player+24, 10
		  Else
			If score(player)>100000 Then Controller.b2ssetscorerollover player+24, Int(score(player)/100000)
		end if
	  Else
		for i = 1 to 10: EVAL("rollover"&i).visible=0: next
		if score(player)>1000000 then
			rollover10.state=1
		  Else
			If score(player)>100000 Then EVAL("rollover"&(Int(score(player)/100000))).visible=1
		end if
	End If

    ' Sounds: there are 3 sounds: tens, hundreds and thousands
    If Points = 100 AND(Score(player) MOD 1000) \ 100 = 0 Then  'New 1000 reel
        PlaySound SoundFXDOF("ding3",143,DOFPulse,DOFChimes)
      ElseIf Points = 10 AND(Score(player) MOD 100) \ 10 = 0 Then 'New 100 reel
        PlaySound SoundFXDOF("ding2",142,DOFPulse,DOFChimes)
      ElseIf points = 1000 Then
        PlaySound SoundFXDOF("ding3",143,DOFPulse,DOFChimes)
	  elseif Points = 100 Then
        PlaySound SoundFXDOF("ding2",142,DOFPulse,DOFChimes)
      Else
        PlaySound SoundFXDOF("ding1",141,DOFPulse,DOFChimes)
	End If
	checkreplay
end sub

sub checkreplay
    if score(player)=>replay1 and rep(player)=0 then
		addballs
		rep(player)=1
    end if
    if score(player)=>replay2 and rep(player)=1 then
		addballs
		rep(player)=2
    end if
end sub

Sub CheckTilt
	If Tilttimer.Enabled = True Then
	 TiltSens = TiltSens + 1
	 if TiltSens = 3 Then
	   Tilt = True
		tiltreel.visible=1
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
		tiltreel.visible=1
       	If B2SOn Then Controller.B2SSetTilt 33,1
       	If B2SOn Then Controller.B2ssetdata 1, 0
	   playsound "tilt"
	   turnoff
End Sub

Sub Tilttimer_Timer()
	Tilttimer.Enabled = False
End Sub

sub turnoff
    bumper1.force=0
    bumper2.force=0
	bumper3.force=0
  	LeftFlipper.RotateToStart
	DOF 101, DOFOff
	StopSound "Buzz"
	RightFlipper.RotateToStart
	DOF 102, DOFOff
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
	PlaySound "pinhit_low", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0, AudioFade(ActiveBall)
End Sub

Sub Targets_Hit (idx)
	PlaySound "target", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0, AudioFade(ActiveBall)
End Sub

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

Sub Spinner_Spin
	PlaySound "fx_spinner",0,.25,0,0.25
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
    savevalue "GStarTrek", "credit", credit
    savevalue "GStarTrek", "hiscore", hisc
    savevalue "GStarTrek", "score1", score(1)
	savevalue "GStarTrek", "replays", replays
	savevalue "GStarTrek", "balls", balls
end sub

sub loadhs
    dim temp
	temp = LoadValue("GStarTrek", "credit")
    If (temp <> "") then credit = CDbl(temp)
    temp = LoadValue("GStarTrek", "hiscore")
    If (temp <> "") then hisc = CDbl(temp)
    temp = LoadValue("GStarTrek", "score1")
    If (temp <> "") then score(1) = CDbl(temp)
    temp = LoadValue("GStarTrek", "replays")
    If (temp <> "") then replays = CDbl(temp)
    temp = LoadValue("GStarTrek", "balls")
    If (temp <> "") then balls = CDbl(temp)
end sub

Sub StarTrek_Exit()
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

Function AudioFade(tableobj) ' Fades between front and back of the table (for surround systems or 2x2 speakers, etc), depending on the Y position on the table. "StarTrek" is the name of the table
  Dim tmp
  tmp = tableobj.y * 2 / StarTrek.height-1
  If tmp > 0 Then
    AudioFade = Csng(tmp ^10)
  Else
    AudioFade = Csng(-((- tmp) ^10) )
  End If
End Function

Function AudioPan(tableobj) ' Calculates the pan for a tableobj based on the X position on the table. "StarTrek" is the name of the table
  Dim tmp
  tmp = tableobj.x * 2 / StarTrek.width-1
  If tmp > 0 Then
    AudioPan = Csng(tmp ^10)
  Else
    AudioPan = Csng(-((- tmp) ^10) )
  End If
End Function

Function Pan(ball) ' Calculates the pan for a ball based on the X position on the table. "StarTrek" is the name of the table
    Dim tmp
    tmp = ball.x * 2 / StarTrek.width-1
    If tmp > 0 Then
        Pan = Csng(tmp ^10)
    Else
        Pan = Csng(-((- tmp) ^10) )
    End If
End Function

Function AudioFade(ball) ' Can this be together with the above function ?
  Dim tmp
  tmp = ball.y * 2 / StarTrek.height-1
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
  If StarTrek.VersionMinor > 3 OR StarTrek.VersionMajor > 10 Then
    PlaySound("fx_collide"), 0, Csng(velocity) ^2 / 200, Pan(ball1), 0, Pitch(ball1), 0, 0, AudioFade(ball1)
  Else
    PlaySound("fx_collide"), 0, Csng(velocity) ^2 / 200, Pan(ball1), 0, Pitch(ball1), 0, 0
  End if
End Sub


' Thalamus : Exit in a clean and proper way
Sub StarTrek_exit()
  Controller.Pause = False
  Controller.Stop
End Sub


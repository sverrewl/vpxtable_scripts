'	FLYING CARPET
'	Gottlieb, 1972
'
'   Layout and graphics by HauntFreaks
'	scripting by BorgDog
'
'	-------------------------------------------------
'	built on BorgDog's Gottlieb EM 4 player VPX table blank with options menu
'	hold left flipper before starting game for options (idea from loserman and gnance)
'	-------------------------------------------------
'
'
'	Layer usage
'		1 - most stuff
'		2 - triggers
'		4 - options menu
'		6 - GI lighting
'		7 - plastics
'		8 - insert and bumper lighting
'
'	-------------------------------------------------

' Thalamus 2018-07-23
' Added/Updated "Positional Sound Playback Functions" and "Supporting Ball & Sound Functions"
' No special SSF tweaks yet.

Option Explicit
Randomize

Const cGameName = "flying_carpet"

Dim operatormenu, options
Dim bumperlitscore
Dim bumperoffscore
Dim balls
Dim replays
Dim Replay1Table(3), Replay2Table(3), Replay3Table(3)
dim replay1, replay2, replay3
dim replaytext(3)
Dim Add10, Add100, Add1000
Dim hisc
Dim maxplayers
Dim players
Dim player
Dim credit
Dim score(4)
Dim sreels(4)
Dim Pups(4)
Dim state
Dim tilt
Dim tiltsens
Dim target(9)
Dim ballinplay
Dim matchnumb
dim ballrenabled
dim rdwstep(8)
Dim rep(4)
Dim rst
Dim eg
Dim bell
Dim i,j, f, ii, objekt, light
Dim specialon, specials

On Error Resume Next
ExecuteGlobal GetTextFile("controller.vbs")
If Err Then MsgBox "Can't open controller.vbs"
On Error Goto 0

sub Flying_Carpet_init
	LoadEM
	maxplayers=1
	Replay1Table(1)=4700
	Replay2Table(1)=6100
	Replay3Table(1)=6900
	Replay1Table(2)=5900
	Replay2Table(2)=7300
	Replay3Table(2)=8100
	set sreels(1) = ScoreReel1
	set Pups(1) = Pup1
	hideoptions
	player=1
	loadhs
	if hisc="" then hisc=1000
	hstxt.text=hisc
	bip_game.state=1
	tilttxt.timerenabled=1
	credittxt.setvalue(credit)
	if credit>0 then DOF 123, DOFOn
	if balls="" then balls=5
	if balls<>3 and balls<>5 then balls=5
	if replays="" then replays=1
	if replays<>1 and replays<>2 then replays=1
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
	if matchnumb="" or matchnumb<0 or matchnumb>9 then matchnumb=0
	matchtxt.text=matchnumb
	for i = 1 to maxplayers
		sreels(i).setvalue(score(i))
	next
	PlaySound "motor"
	tilt=false
    Drain.CreateBall
End sub

sub setBackglass_timer
	Controller.B2ssetCredits Credit
	Controller.B2ssetMatch 34, Matchnumb
    Controller.B2SSetGameOver 35,1
    Controller.B2SSetScorePlayer 2, hisc
	me.enabled=0
end Sub

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

Sub Flying_Carpet_KeyDown(ByVal keycode)

	if keycode=AddCreditKey then
		playsound "coinin"
		coindelay.enabled=true
    end if

    if keycode=StartGameKey and credit>0 then
	  if state=false then
		credit=credit-1
		playsound "cluper"
		credittxt.setvalue(credit)
		if credit<1 then DOF 123, DOFOff
		ballinplay=1
		If B2SOn Then
			Controller.B2ssetCredits Credit
			Controller.B2sStartAnimation "Startup"
			Controller.B2ssetballinplay 32, Ballinplay
			Controller.B2ssetplayerup 30, 1
			Controller.B2ssetcanplay 31, 1
			Controller.B2SSetGameOver 0
		End If
	    pups(1).state=1
		tilt=false
		state=true
		playsound "initialize"
		players=1
		rst=0
		newgame.enabled=true
	  else if state=true and players < maxplayers and Ballinplay=1 then
		credit=credit-1
		players=players+1
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
			repcard.image = "ReplayCard"&replays
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


Sub Flying_Carpet_KeyUp(ByVal keycode)

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
	Pgate.rotz=Gate1.currentangle+25
end sub


Sub PairedlampTimer_timer
	bumperlighta1.state = bumperlight1.state
	bumperlighta2.state = bumperlight2.state
	bumperlighta3.state = bumperlight3.state
	LF2.state=LF.state: LL2.state=LL.state: LY2.state=LY.state: LI2.state=LI.state: LN2.state=LN.state: LG2.state=LG.state
	LC2.state=LC.state: LA2.state=LA.state: LR2.state=LR.state: LP2.state=LP.state: LE2.state=LE.state: LT2.state=LT.state
	LF3.state=LF.state: LL3.state=LL.state: LY3.state=LY.state: LI3.state=LI.state: LN3.state=LN.state: LG3.state=LG.state
	LC3.state=LC.state: LA3.state=LA.state: LR3.state=LR.state: LP3.state=LP.state: LE3.state=LE.state: LT3.state=LT.state
	Lspecial3.state=Lspecial1.state: Lspecial4.state=Lspecial2.state
	Llane1t.state=llane1.state: Llane2t.state=llane2.state: Llane3t.state=Llane3.state: Llane4t.state=Llane4.state: Llane5t.state=Llane5.state
end sub

sub coindelay_timer
	addcredit
    coindelay.enabled=false
end sub

Sub addcredit
      credit=credit+1
      if credit>25 then credit=25
	  credittxt.setvalue(credit)
	  If B2SOn Then Controller.B2ssetCredits Credit
	  DOF 123, DOFOn
End sub

Sub Drain_Hit()
	PlaySound "drain",0,1,0,0.25
	DOF 132, DOFPulse
	me.timerenabled=1
End Sub

Sub Drain_timer
	nextball
	me.timerenabled=0
End Sub

sub ballhome_hit
	DOF 124, DOFOn
	ballrenabled=1
end sub

sub ballhome_unhit
	DOF 124, DOFOff
	DOF 125, DOFPulse
end sub


sub ballrel_hit
	if ballrenabled=1 then
		ballrenabled=0
	end if
end sub

sub newgame_timer
	for f=1 to 3
		EVAL("Bumper"&f).hashitevent = 1
	Next
	player=1
	for i = 1 to maxplayers
	    score(i)=0
		rep(i)=0
		pups(i).state=0
		sreels(i).setvalue(score(i))
	next
	BumperLight1.state=0
	BumperLight2.state=1
	BumperLight3.state=0
    pups(1).state=1
	specialon=0
    eg=0
	For each light in GIlights:light.state=1:next
	for each light in bonuslights:light.state=0:next
	for each light in letterlights:light.state=1: next
	for each light in speciallights:light.state=0: next
    tilttxt.text=" "
	bip_game.state=0
    If B2SOn then
		  for i = 1 to maxplayers
			Controller.B2SSetScorePlayer i, score(i)
		  next
		Controller.B2SSetGameOver 35,0
		Controller.B2SSetTilt 33,0
		Controller.B2SSetMatch 34,0
	End If
    bip1.state=1
	matchtxt.text=" "
	ballrelTimer.enabled=true
	newgame.enabled=false
end sub

sub nextball
    if tilt=true then
	  for f=1 to 3
		EVAL("Bumper"&f).hashitevent = 1
	  Next
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
		ballreltimer.enabled=true
		end if
		Eval("bip"&ballinplay).state=1
		Eval("bip"&ballinplay-1).state=0
		If B2SOn then Controller.B2ssetballinplay 32, Ballinplay
	end if
End Sub

sub ballreltimer_timer
  if eg=1 then
	  turnoff
      matchnum
	  state=false
	  bip5.state=0
	  bip_game.state=1
	  tilttxt.timerenabled=1
	  for i=1 to maxplayers
		if score(i)>hisc then hisc=score(i)
		pups(i).state=0
	  next
	  hstxt.text=hisc
	  savehs
	  If B2SOn then
        Controller.B2SSetGameOver 35,1
        Controller.B2ssetballinplay 32, 0
	    Controller.B2SSetScorePlayer 2, hisc
	    Controller.B2ssetcanplay 31, 0
	    Controller.B2ssetcanplay 30, 0
	  End If
	  For each light in LaneLights:light.State = 0: Next
	  For each light in GIlights:light.state=0:next
	  ballreltimer.enabled=false
  else
	Drain.kick 60,12,0
    ballreltimer.enabled=false
	playsound SoundFXDOF("kickerkick", 121, DOFPulse,DOFContactors)
  end if
end sub

sub matchnum
	matchtxt.text=matchnumb
	If B2SOn then Controller.B2SSetMatch matchnumb
	For i=1 to players
		if matchnumb=(score(i) mod 10) then
		  addcredit
		  PlaySound SoundFXDOF("knock",122,DOFPulse,DOFKnocker)
	    end if
    next
end sub

'********** Bumpers

Sub Bumper1_Hit
   if tilt=false then
    PlaySound SoundFXDOF("fx_bumper4",104, DOFPulse,DOFContactors)
	if bumperlight1.state=1 then
		addscore bumperlitscore
	  else
		addscore bumperoffscore
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
    PlaySound SoundFXDOF("fx_bumper4",103, DOFPulse,DOFContactors)
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
    PlaySound SoundFXDOF("fx_bumper4",105, DOFPulse,DOFContactors)
	if BumperLight3.state=1 then
		addscore bumperlitscore
	  else
		addscore bumperoffscore
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

Sub Dingwalls_slingshot(idx)
	Dim tempid
	tempid=idx+1
	addscore 1
	if specials=1 then
		specials=0
	  Else
		specials=1
	end If
	if SpecialOn=1 Then
		if specials=1 Then
			LSpecial1.state=1
			Lspecial2.state=0
		  Else
			Lspecial2.state=1
			Lspecial1.state=0
		end if
	end If
	Eval("RDW"&tempid).visible = 0
	Eval("RDWa"&tempid).visible = 1
	rdwstep(tempid) = 1
	Eval("Dingwall"&tempid).timerenabled = 1
End Sub

Sub Dingwall1_Timer
	Select Case rdwstep(1)
		Case 1: RDWa1.visible = 0:RDW1.visible = 1
		Case 2: RDW1.visible = 0:RDWb1.visible = 1
		Case 3: RDWb1.visible = 0:RDW1.visible = 1: Dingwall1.timerenabled = 0
	End Select
	rdwstep(1)=rdwstep(1)+1
End Sub

Sub Dingwall2_Timer
	Select Case rdwstep(2)
		Case 1: RDWa2.visible = 0:RDW2.visible = 1
		Case 2: RDW2.visible = 0:RDWb2.visible = 1
		Case 3: RDWb2.visible = 0:RDW2.visible = 1: Dingwall2.timerenabled = 0
	End Select
	rdwstep(2)=rdwstep(2)+1
End Sub

Sub Dingwall3_Timer
	Select Case rdwstep(3)
		Case 1: RDWa3.visible = 0:RDW3.visible = 1
		Case 2: RDW3.visible = 0:RDWb3.visible = 1
		Case 3: RDWb3.visible = 0:RDW3.visible = 1: Dingwall3.timerenabled = 0
	End Select
	rdwstep(3)=rdwstep(3)+1
End Sub

Sub Dingwall4_Timer
	Select Case rdwstep(4)
		Case 1: RDWa4.visible = 0:RDW4.visible = 1
		Case 2: RDW4.visible = 0:RDWb4.visible = 1
		Case 3: RDWb4.visible = 0:RDW4.visible = 1: Dingwall4.timerenabled = 0
	End Select
	rdwstep(4)=rdwstep(4)+1
End Sub

Sub Dingwall5_Timer
	Select Case rdwstep(5)
		Case 1: RDWa5.visible = 0:RDW5.visible = 1
		Case 2: RDW5.visible = 0:RDWb5.visible = 1
		Case 3: RDWb5.visible = 0:RDW5.visible = 1: Dingwall5.timerenabled = 0
	End Select
	rdwstep(5)=rdwstep(5)+1
End Sub

Sub Dingwall6_Timer
	Select Case rdwstep(6)
		Case 1: RDWa6.visible = 0:RDW6.visible = 1
		Case 2: RDW6.visible = 0:RDWb6.visible = 1
		Case 3: RDWb6.visible = 0:RDW6.visible = 1: Dingwall6.timerenabled = 0
	End Select
	rdwstep(6)=rdwstep(6)+1
End Sub

Sub Dingwall7_Timer
	Select Case rdwstep(7)
		Case 1: RDWa7.visible = 0:RDW7.visible = 1
		Case 2: RDW7.visible = 0:RDWb7.visible = 1
		Case 3: RDWb7.visible = 0:RDW7.visible = 1: Dingwall7.timerenabled = 0
	End Select
	rdwstep(7)=rdwstep(7)+1
End Sub

Sub Dingwall8_Timer
	Select Case rdwstep(8)
		Case 1: RDWa8.visible = 0:RDW8.visible = 1
		Case 2: RDW8.visible = 0:RDWb8.visible = 1
		Case 3: RDWb8.visible = 0:RDW8.visible = 1: Dingwall8.timerenabled = 0
	End Select
	rdwstep(8)=rdwstep(8)+1
End Sub

'********** Triggers

Sub triggers_hit(idx)
	Llane1.uservalue=1
	Llane1.timerenabled=1
End Sub

Sub Llane1_timer
	select case LLane1.uservalue
		Case 1: Llane1.state=0
		Case 2: Llane2.state=0: Llane1.state=1
		Case 3: Llane3.state=0: Llane2.state=1
		Case 4: Llane4.state=0: Llane3.state=1
		Case 5: Llane5.state=0: Llane4.state=1
		Case 6: Llane5.state=1: Llane1.timerenabled=0
	end Select
	LLane1.uservalue=Llane1.uservalue+1
End Sub

Sub TGF_Hit
	addscore 100
	DOF 112, DOFPulse
	LF.state=0
	LF1.state=1
	checkspecials
end Sub

Sub TGL_Hit
	addscore 50
	DOF 106, DOFPulse
	LL.state=0
	LL1.state=1
	checkspecials
end Sub

Sub TGY_Hit
	addscore 100
	DOF 111, DOFPulse
	LY.state=0
	LY1.state=1
	if Lspecial4.state=1 then addspecial
	checkspecials
end Sub

Sub TGI_Hit
	addscore 50
	DOF 109, DOFPulse
	LI.state=0
	LI1.state=1
	checkspecials
end Sub

Sub TGN_Hit
	addscore 50
	DOF 115, DOFPulse
	LN.state=0
	LN1.state=1
	checkspecials
end Sub

Sub TGG_Hit
	addscore 50
	DOF 116, DOFPulse
	LG.state=0
	LG1.state=1
	checkspecials
end Sub

Sub TGC_Hit
	addscore 100
	DOF 113, DOFPulse
	LC.state=0
	LC1.state=1
	checkspecials
end Sub

Sub TGA_Hit
	addscore 50
	DOF 108, DOFPulse
	LA.state=0
	LA1.state=1
	if balls=3 or (balls=5 and LP.state=0) Then
		BumperLight1.state=1
		Bumperlight3.state=1
	end If
	checkspecials
end Sub

Sub TGR_Hit
	addscore 100
	DOF 110, DOFPulse
	LR.state=0
	LR1.state=1
	if Lspecial1.state=1 then addspecial
	checkspecials
end Sub

Sub TGP_Hit
	addscore 100
	DOF 107, DOFPulse
	LP.state=0
	LP1.state=1
	if balls=3 or (balls=5 and LA.state=0) Then
		BumperLight1.state=1
		Bumperlight3.state=1
	end If
	checkspecials
end Sub

Sub TGE_Hit
	addscore 100
	DOF 117, DOFPulse
	LE.state=0
	LE1.state=1
	checkspecials
end Sub

Sub TGT_Hit
	addscore 100
	DOF 114, DOFPulse
	LT.state=0
	LT1.state=1
	checkspecials
end Sub

'********** Targets

Sub TargetF_hit
	addscore 100
	DOF 120, DOFPulse
	LF.state=0
	LF1.state=1
	checkspecials
End Sub

Sub TargetL_Hit
	addscore 50
	DOF 126, DOFPulse
	LL.state=0
	LL1.state=1
	checkspecials
end Sub

Sub TargetY_Hit
	addscore 100
	DOF 128, DOFPulse
	LY.state=0
	LY1.state=1
	If Lspecial2.state=1 then addspecial
	checkspecials
end Sub

Sub TargetI_Hit
	addscore 50
	DOF 127, DOFPulse
	LI.state=0
	LI1.state=1
	checkspecials
end Sub

Sub TargetN_Hit
	addscore 50
	DOF 119, DOFPulse
	LN.state=0
	LN1.state=1
	checkspecials
end Sub

Sub TargetG_Hit
	addscore 50
	DOF 119, DOFPulse
	LG.state=0
	LG1.state=1
	checkspecials
end Sub

Sub TargetC_Hit
	addscore 100
	DOF 118, DOFPulse
	LC.state=0
	LC1.state=1
	checkspecials
end Sub

Sub TargetA_Hit
	addscore 100
	DOF 127, DOFPulse
	LA.state=0
	LA1.state=1
	if balls=3 or (balls=5 and LP.state=0) Then
		BumperLight1.state=1
		Bumperlight3.state=1
	end if
	checkspecials
end Sub

Sub TargetR_Hit
	addscore 100
	DOF 128, DOFPulse
	LR.state=0
	LR1.state=1
	If Lspecial3.state=1 then addspecial
	checkspecials
end Sub

Sub TargetP_Hit
	addscore 100
	DOF 126, DOFPulse
	LP.state=0
	LP1.state=1
	if balls=3 or (balls=5 and LA.state=0) Then
		BumperLight1.state=1
		Bumperlight3.state=1
	end if
	checkspecials
end Sub

Sub TargetE_Hit
	addscore 100
	DOF 119, DOFPulse
	LE.state=0
	LE1.state=1
	checkspecials
end Sub

Sub TargetT_Hit
	addscore 100
	DOF 119, DOFPulse
	LT.state=0
	LT1.state=1
	checkspecials
end Sub


sub checkspecials
	if balls=3 Then
		if (LF.state+LL.state+LY.state+LI.state+LN.state+LG.state=0) or (LC.state+LA.state+LR.state+LP.state+LE.state+LT.state=0) Then
			SpecialOn=1
			if specials=1 Then
				LSpecial1.state=1
				Lspecial2.state=0
			  Else
				Lspecial2.state=1
				Lspecial1.state=0
			end if
		end If
	  Else
		if (LF.state+LL.state+LY.state+LI.state+LN.state+LG.state+LC.state+LA.state+LR.state+LP.state+LE.state+LT.state=0) Then
			SpecialOn=1
			if specials=1 Then
				LSpecial1.state=1
				Lspecial2.state=0
			  Else
				Lspecial2.state=1
				Lspecial1.state=0
			end if
		end If
	end If
end Sub

sub addspecial
	addcredit
	PlaySound SoundFXDOF("knock",123,DOFPulse,DOFKnocker)
	DOF 122, DOFPulse
end Sub

sub addscore(points)
  if tilt=false then
	If points = 1 then
		matchnumb=matchnumb+1
		if matchnumb>9 then matchnumb=0
	end if
	If points=1 or points=10 or points=100 Then
		addpoints Points
	  Else
		If Points < 10 and AddScore10Timer.enabled = false Then
			Add10 = Points
			AddScore10Timer.Enabled = TRUE
		  ElseIf Points < 100 and AddScore100Timer.enabled = false Then
			Add100 = Points \ 10
			AddScore100Timer.Enabled = TRUE
		  ElseIf AddScore1000Timer.enabled = false Then
			Add1000 = Points \ 100
			AddScore1000Timer.Enabled = TRUE
		End If
	End if
  end if
End Sub

Sub AddScore10Timer_Timer()
    if Add10 > 0 then
        AddPoints 1
        Add10 = Add10 - 1
    Else
        Me.Enabled = FALSE
    End If
End Sub

Sub AddScore100Timer_Timer()
    if Add100 > 0 then
        AddPoints 10
        Add100 = Add100 - 1
    Else
        Me.Enabled = FALSE
    End If
End Sub

Sub AddScore1000Timer_Timer()
    if Add1000 > 0 then
        AddPoints 100
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
        PlaySound SoundFXDOF("bell1000",131,DOFPulse,DOFChimes)
      ElseIf Points = 10 AND(Score(player) MOD 100) \ 10 = 0 Then 'New 100 reel
        PlaySound SoundFXDOF("bell100",130,DOFPulse,DOFChimes)
      ElseIf points = 1000 Then
        PlaySound SoundFXDOF("bell1000",131,DOFPulse,DOFChimes)
	  elseif Points = 100 Then
        PlaySound SoundFXDOF("bell100",130,DOFPulse,DOFChimes)
      Else
        PlaySound SoundFXDOF("bell10",129,DOFPulse,DOFChimes)
    End If
    ' check replays
    if score(player)=>replay1 and rep(player)=0 then
		addcredit
		rep(player)=1
		PlaySound SoundFXDOF("knock",123,DOFPulse,DOFKnocker)
		DOF 122, DOFPulse
    end if
    if score(player)=>replay2 and rep(player)=1 then
		addcredit
		rep(player)=2
		PlaySound SoundFXDOF("knock",123,DOFPulse,DOFKnocker)
		DOF 122, DOFPulse
    end if
    if score(player)=>replay3 and rep(player)=2 then
		addcredit
		rep(player)=3
		PlaySound SoundFXDOF("knock",123,DOFPulse,DOFKnocker)
		DOF 122, DOFPulse
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
	 LeftFlipper.RotateToStart
	 RightFlipper.RotateToStart
End Sub

Sub Tilttimer_Timer()
	Tilttimer.Enabled = False
End Sub

sub turnoff
	for f=1 to 3
		EVAL("Bumper"&f).hashitevent = 0
	Next
  	LeftFlipper.RotateToStart
	StopSound "Buzz"
	DOF 101, DOFOff
	RightFlipper.RotateToStart
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

    savevalue "Flying_Carpet", "credit", credit
    savevalue "Flying_Carpet", "hiscore", hisc
    savevalue "Flying_Carpet", "match", matchnumb
    savevalue "Flying_Carpet", "score1", score(1)
	savevalue "Flying_Carpet", "balls", balls

end sub

sub loadhs
    dim temp
	temp = LoadValue("Flying_Carpet", "credit")
    If (temp <> "") then credit = CDbl(temp)
    temp = LoadValue("Flying_Carpet", "hiscore")
    If (temp <> "") then hisc = CDbl(temp)
    temp = LoadValue("Flying_Carpet", "match")
    If (temp <> "") then matchnumb = CDbl(temp)
    temp = LoadValue("Flying_Carpet", "score1")
    If (temp <> "") then score(1) = CDbl(temp)
    temp = LoadValue("Flying_Carpet", "balls")
    If (temp <> "") then balls = CDbl(temp)
end sub

Sub Flying_Carpet_Exit()
	Savehs
	turnoff
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

Function AudioFade(tableobj) ' Fades between front and back of the table (for surround systems or 2x2 speakers, etc), depending on the Y position on the table. "Flying_Carpet" is the name of the table
  Dim tmp
  tmp = tableobj.y * 2 / Flying_Carpet.height-1
  If tmp > 0 Then
    AudioFade = Csng(tmp ^10)
  Else
    AudioFade = Csng(-((- tmp) ^10) )
  End If
End Function

Function AudioPan(tableobj) ' Calculates the pan for a tableobj based on the X position on the table. "Flying_Carpet" is the name of the table
  Dim tmp
  tmp = tableobj.x * 2 / Flying_Carpet.width-1
  If tmp > 0 Then
    AudioPan = Csng(tmp ^10)
  Else
    AudioPan = Csng(-((- tmp) ^10) )
  End If
End Function

Function Pan(ball) ' Calculates the pan for a ball based on the X position on the table. "Flying_Carpet" is the name of the table
    Dim tmp
    tmp = ball.x * 2 / Flying_Carpet.width-1
    If tmp > 0 Then
        Pan = Csng(tmp ^10)
    Else
        Pan = Csng(-((- tmp) ^10) )
    End If
End Function

Function AudioFade(ball) ' Can this be together with the above function ?
  Dim tmp
  tmp = ball.y * 2 / Flying_Carpet.height-1
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
  If Flying_Carpet.VersionMinor > 3 OR Flying_Carpet.VersionMajor > 10 Then
    PlaySound("fx_collide"), 0, Csng(velocity) ^2 / 200, Pan(ball1), 0, Pitch(ball1), 0, 0, AudioFade(ball1)
  Else
    PlaySound("fx_collide"), 0, Csng(velocity) ^2 / 200, Pan(ball1), 0, Pitch(ball1), 0, 0
  End if
End Sub


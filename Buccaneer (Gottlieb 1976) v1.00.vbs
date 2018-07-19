'	-------------------------------------------------
'	Buccaneer by Gottlieb (1976)
'
'
'	Retheme of Jungle Princess by BorgDog, 2016
'   Allknowing2012, Hauntfreaks
'   DOF from Argrin

'
' Lane 1 Dividers set back to the Liberal settings
' SpinNSpot coded with the conservative settings
' Outlane posts set to the liberal settings

' 127 Credit Light
' 128 knocker
' 101,102 flippers
' 126 ballrelease
'	-------------------------------------------------

' Thalamus 2018-07-19
' Added/Updated "Positional Sound Playback Functions" and "Supporting Ball & Sound Functions"
' No special SSF tweaks yet.

Option Explicit
Randomize

Const cGameName = "buccaneer"
Dim PI: PI = Round(4 * Atn(1), 6)

Dim operatormenu, options
Dim bumperlitscore, bumperoffscore
Dim balls, special, spstate
Dim replays, Replay1Table(3), Replay2Table(3), Replay3Table(3)
dim replay1, replay2, replay3
Dim Add10, Add100, Add1000, hisc
Dim maxplayers, players, player
Dim credit
Dim score(4)
Dim sreels(4)
Dim p100k(4)
Dim cplay(4)
Dim Pups(4)
Dim state, freeplay, pfLightState(13),LightState(13)
Dim tilt, tiltsens
Dim dw1step, dw2step, dw3step, dw4step, dw5step, dw6step, dw7step, dw8step, dw9step, dw10step
Dim ballinplay
Dim matchnumb
dim ballrenabled
Dim rep(4)
Dim eg
Dim bell
Dim a,i,j, ii, objekt, light, wheelno, spinnspot

On Error Resume Next
ExecuteGlobal GetTextFile("controller.vbs")
If Err Then MsgBox "Can't open controller.vbs"
On Error Goto 0

sub AllLampsOff
DIM x
	For each light in lights:light.State = 0: Next
    For each Light in YLights
      Light.State=LightStateOff
    Next
    x=0
    for each Light in WLights
      LightState(x) = Light.State
      Light.State=LightStateOff
      x=x+1
    Next
    x=0
    for each Light in pfLights
      pfLightState(x) = Light.State
      Light.State=LightStateOff
      x=x+1
    Next
End Sub

sub Buccaneer_init
	LoadEM
	maxplayers=1
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
	set Pups(1) = Pup1
	set cplay(1) = CanPlay1
	set p100k(1) = P1100k
	hideoptions
	player=1
    spinnspot=0
	For each light in BumperLights:light.State = 0: Next
    AllLampsOff
	loadhs
	if hisc="" then hisc=90000
	hstxt.text=hisc
	gamov.text="Game Over"
	gamov.timerenabled=1
	tilttxt.timerenabled=1
	credittxt.setvalue(credit)
    if credit=0 then
      DOF 127,0
    else
      DOF 127,1
    end if
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
	if matchnumb="" then matchnumb=0
	if matchnumb=0 then
		matchtxt.text="00"
	  else
		matchtxt.text=matchnumb
	end if
	for i = 1 to maxplayers
		sreels(i).setvalue(score(i))
	next
    biptext.text=" "
	PlaySound "motor"
	tilt=false
	If credit>0 then DOF 127, DOFOn
    Drain.CreateBall

    'Init Lights
    a=int(rnd*10)
    if a=10 then a=0
    YLights(a).state=LightStateOn:wheelno=a
End sub

sub setBackglass_timer
	Controller.B2ssetCredits Credit
	Controller.B2ssetMatch 34, Matchnumb
    Controller.B2SSetGameOver 35,1
    Controller.B2SSetScorePlayer 5, hisc
    Controller.B2SSetScorePlayer 1, score(1)
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
		If Buccaneer.ShowDT then tilttxt.visible=1
		If B2SOn then Controller.B2SSetTilt 33,1
		tilttxt.timerinterval= (INT (RND*10)+5)*100
		tilttxt.timerenabled=1
	end if
	me.enabled=0
end sub

Sub FlashBumpers
  For each light in BumperLights:light.State = 0: Next
  BumperLightT.enabled=True
End Sub

Sub BumperLightT_timer
  BumperLightT.enabled=False
  BumperLight2.state = Light1W.state
  BumperLight1.state = Light4W.state
  BumperLight3.state = Light5W.state
  bumperlight4.state = LightStateON
  bumperlight5.state = LightStateON
  Bumperlighta4.state = bumperlight4.state
  Bumperlighta5.state = bumperlight5.state
End Sub

Sub Buccaneer_KeyDown(ByVal keycode)

	if keycode=AddCreditKey then
		playsound "coinin"
		coindelay.enabled=true
    end if

    if keycode=StartGameKey and (credit>0 or freeplay=1) and OperatorMenu=0 then
	  if state=false then
		if freeplay=0 then credit=credit-1
		if credit < 1 then DOF 127, DOFOff
		playsound "cluper"
		credittxt.setvalue(credit)
		ballinplay=1
		For each light in BumperLights:light.State = 0: Next
		For each light in YellowBumperLights:light.State = 1: Next
		for each light in GIlights:light.state=1:next
		for each light in lights:light.state=1:next
		for each light in wlights:light.state=0:next
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
        if credt < 1 then DOF 127,DOFOff
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
		LeftFlipper.RotateToStart
		RightFlipper.RotateToStart
		StopSound "Buzz"
		StopSound "Buzz1"
		Tilt=true
		tilttxt.visible=1
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

Sub Buccaneer_KeyUp(ByVal keycode)

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
        LeftFlipper.TimerEnabled = 1    ' nFozzy Flipper Code
        LeftFlipper.TimerInterval = 16
        LeftFlipper.return = returnspeed * 0.5
		PlaySound SoundFXDOF("flipperdown",101,DOFOff,DOFContactors), 0, 1, -0.05, 0.05
		StopSound "Buzz"
	End If

	If keycode = RightFlipperKey Then
		RightFlipper.RotateToStart
        Rightflipper.TimerEnabled = 1  ' nFozzy Flipper Code
        Rightflipper.TimerInterval = 16
        Rightflipper.return = returnspeed * 0.5
		PlaySound SoundFXDOF("flipperdown",102,DOFOff,DOFContactors), 0, 1, 0.05, 0.05
		StopSound "Buzz1"
	End If
   End if
End Sub

sub flippertimer_timer()
	LFlip.RotY = LeftFlipper.CurrentAngle
	RFlip.RotY = RightFlipper.CurrentAngle
	Pgate.rotz = Gate.currentangle+25
end sub

Sub PairedlampTimer_timer
	bumperlighta1.state = bumperlight1.state
	bumperlighta2.state = bumperlight2.state
	bumperlighta3.state = bumperlight3.state
	bumperlighta4.state = bumperlight4.state
	bumperlighta5.state = bumperlight5.state
end sub

sub coindelay_timer
	addcredit
    coindelay.enabled=false
end sub

Sub addcredit
      credit=credit+1
	  DOF 127, DOFOn
      if credit>15 then credit=15
	  credittxt.setvalue(credit)
	  If B2SOn Then Controller.B2ssetCredits Credit
End sub

Sub Drain_Hit()
	DOF 130, DOFPulse
	PlaySound "drain",0,1,0,0.25
	for each light in GIlights:light.state=0:next
	'for each light in lights:light.state=0:next
	me.timerenabled=1
    FlashBumpers()
	tilt=false
	tilttxt.visible=0
End Sub

Sub Drain_timer
	me.timerenabled=0
    nextball
End Sub

sub ballhome_hit
    DOF 131, DOFOn
	ballrenabled=1
end sub

sub ballhome_unhit
    DOF 131, DOFOff
	DOF 132, DOFPulse
end sub

sub ballrel_hit
	if ballrenabled=1 then
		ballrenabled=0
	end if
end sub

sub newgame_timer
	special=0
	player=1
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
    tilttxt.visible=0
	gamov.text=" "
    If B2SOn then
		Controller.B2SSetGameOver 35,0
		Controller.B2SSetTilt 33,0
		Controller.B2SSetMatch 34,0
	End If
    biptext.text="1"
	matchtxt.text=" "
	newball
	ballrelTimer.enabled=true
	newgame.enabled=false
end sub


sub newball
	spstate=0
	bumper1.hashitevent=True
    bumper2.hashitevent=True
    bumper3.hashitevent=True
    bumper4.hashitevent=True
    bumper5.hashitevent=True
	'For each light in BumperLights:light.State = 0: Next
	'For each light in YellowBumperLights:light.State = 1: Next
	for each light in GIlights:light.state=1:next
	'for each light in lights:light.state=1:next
End Sub


sub nextball
    if tilt=true then
      tilt=false
      tilttxt.visible=0
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
	    'Controller.B2ssetcanplay 31, 0
	    'Controller.B2ssetcanplay 30, 0
	  End If
	  For each light in GIlights:light.state=0:next
  else
	Drain.kick 60,28,0
	biptext.text=ballinplay
    If B2SOn then Controller.B2ssetballinplay 32, Ballinplay
    ballreltimer.enabled=false
	playsound SoundFXDOF("kickerkick",126,DOFPulse,DOFContactors)
    playsound "kickerkick"
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
		  playsound SoundFXDOF("knock",128,DOFPulse,DOFKnocker)
	    end if
    next
end sub

'********** Bumpers

Sub Bumper1_Hit
   if tilt=false then
	playsound SoundFXDOF("fx_bumper4",105,DOFPulse,DOFContactors)
	if BumperLight1.state=1 then
		addscore (bumperlitscore)
	  else
		addscore (bumperoffscore)
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
	playsound SoundFXDOF("fx_bumper4",106,DOFPulse,DOFContactors)
	if BumperLight2.state=1 then
		addscore 1000  'Center Red Bumper
	  else
		addscore 10 ' only 10 pts for 5 ball
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
	playsound SoundFXDOF("fx_bumper4",107,DOFPulse,DOFContactors)
	if BumperLight3.state=1 then
		addscore (bumperlitscore)
	  else
		addscore (bumperoffscore)
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

Sub Bumper4_Hit
   if tilt=false then
	playsound SoundFXDOF("fx_bumper4",108,DOFPulse,DOFContactors)
	if balls = 3 then
		addscore 1000
	  else
		addscore 100
	end if
	me.timerenabled=1
   end if
End Sub

Sub Bumper4_timer
	BumperTimerRing4.enabled=0
	if bumperring4.transz>-36 then BumperRing4.transz=BumperRing4.transz-4
	if BumperRing4.transz=-36 then
		BumperTimerRing4.enabled=1
		me.timerenabled=0
	end if
End Sub

Sub BumperTimerRing4_timer
	if bumperring4.transz<0 then BumperRing4.transz=BumperRing4.transz+4
	If BumperRing4.transz=0 then BumperTimerRing4.enabled=0
End sub

Sub Bumper5_timer
	BumperTimerRing5.enabled=0
	if bumperring5.transz>-36 then BumperRing5.transz=BumperRing5.transz-4
	if BumperRing5.transz=-36 then
		BumperTimerRing5.enabled=1
		me.timerenabled=0
	end if
End Sub

Sub Bumper5_Hit
   if tilt=false then
	playsound SoundFXDOF("fx_bumper4",109,DOFPulse,DOFContactors)
	if balls = 3 then
		addscore 1000
	  else
		addscore 100
	end if
	me.timerenabled=1
   end if
End Sub

Sub BumperTimerRing5_timer
	if bumperring5.transz<0 then BumperRing5.transz=BumperRing5.transz+4
	If BumperRing5.transz=0 then BumperTimerRing5.enabled=0
End sub

'****************************************
'  CIRCLE TARGETS HANDLERS
'****************************************

Sub ctarget1_Hit()
	 DOF 125, DOFPulse
     If Not Tilt Then
       addscore(500)
     End if
     ChkSpinNSpot()
End Sub

Sub ctarget2_hit()
	 DOF 124, DOFPulse
     If Not Tilt Then
       addscore(500)
     End if
     ChkSpinNSpot()
End Sub


' SpinnerRod code from Cyperpez and http://www.vpforums.org/index.php?showtopic=35497
Sub CheckSpinnerRod_timer()
	SpinnerRod.TransZ = sin( (spinner1.CurrentAngle+180) * (2*PI/360)) * 5
	SpinnerRod.TransX = -1*(sin( (spinner1.CurrentAngle- 90) * (2*PI/360)) * 5)
End Sub

Sub ChkSpinNSpot  ' light spin-n-spot every 3rd rollover or circletarget
 if not Tilt then
   spinnspot=spinnspot+1
   if spinnspot > 2 then spinnspot=0
    ' if spinnspot = 0 or spinnspot=1 or spinnspot=2 then   '  Liberal toughness
    ' if spinnspot = 0 or spinnspot=1 then                  '  Medium toughness
    if spinnspot=0 then                                     ' Conservative toughness
        if spstate=0 then   ' if Special is lit we dont need spinNspot
          lightSpinner.state=LightStateOn
        end if
    else
        lightSpinner.state=LightStateOff
    end if
    if spstate=0 then   ' No Special Lit yet
      if light1w.state=1 and light2w.state=1 and light3w.state=1 and light4w.state=1 and light5w.state=1 then
        if light6w.state=1 and light7w.state=1 and light8w.state=1 and light9w.state=1 and light10w.state=1 And light11w.state=1 Then
          spstate=1
          lightSpinner.state=LightStateOff
          CheckSpecial()
        end If
      end if
    Else
      CheckSpecial() 'Move Special around
    end if
  end if
End Sub

'****************************************
'  SPINNERS HANDLERS
'****************************************

Sub Spinner1_Spin()
     If Not Tilt Then
       addscore(100)
       Spin
       Spinner1.TimerInterval = 1000
       Spinner1.Timerenabled=false:Spinner1.Timerenabled=true   ' reset timer and start the clock again
     End if
End Sub

Sub Spinner1_Timer
   Spinner1.Timerenabled=false
   if lightSpinner.state=1 and light2y.state=1 and light2.state=1 then light2.state=0:light2w.state=1:ChkSpinNSpot():addscore(5000)
   if lightSpinner.state=1 and light3y.state=1 and light3.state=1 then light3.state=0:light3w.state=1:ChkSpinNSpot():addscore(5000)
   if lightSpinner.state=1 and light4y.state=1 and light4.state=1 then light4.state=0:light4w.state=1:light4b.state=0:ChkSpinNSpot():addscore(5000)
   if lightSpinner.state=1 and light5y.state=1 and light5.state=1 then light5.state=0:light5w.state=1:light5b.state=0:ChkSpinNSpot():addscore(5000)
   if lightSpinner.state=1 and light6y.state=1 and light6.state=1 then light6.state=0:light6w.state=1:ChkSpinNSpot():addscore(5000)
   if lightSpinner.state=1 and light7y.state=1 and light7.state=1 then light7.state=0:light7w.state=1:ChkSpinNSpot():addscore(5000)
   if lightSpinner.state=1 and light8y.state=1 and light8.state=1 then light8.state=0:light8w.state=1:ChkSpinNSpot():addscore(5000)
   if lightSpinner.state=1 and light9y.state=1 and light9.state=1 then light9.state=0:light9w.state=1:ChkSpinNSpot():addscore(5000)
   if lightSpinner.state=1 and light10y.state=1 and light10.state=1 then light10.state=0:light10w.state=1:ChkSpinNSpot():addscore(5000)
   if lightSpinner.state=1 and light11y.state=1 and light11.state=1 then light11.state=0:light11w.state=1:ChkSpinNSpot():addscore(5000)

   if lightSpinner.state=LightStateOn then lightSpinner.state=LightStateOff
End Sub

Sub Spin
    If Not Tilt Then
       YLights(wheelno).state=LightStateOff
       wheelno = wheelno +1
       if wheelno=10 then wheelno=0  ' 0-10 for ChkSpinNSpot 2-11
       YLights(wheelno).state=LightStateOn
    End if
End Sub

'************** Dingwalls

sub Dingwalls_hit(idx)
	addscore 10
end Sub

Sub CheckSpecial()
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

'**********Sling Shot Animations
' Rstep and Lstep  are the variables that increment the animation
'****************
Dim RStep, Lstep

Sub RightSlingShot_Slingshot
	addscore 10
	playsound SoundFXDOF("right_slingshot",104,DOFPulse,DOFContactors)
    RSling.Visible = 0
    RSling1.Visible = 1
    sling1.TransZ = -10
    RStep = 0
    RightSlingShot.TimerEnabled = 1
    lgi17.state=0:lgi18.state=0
End Sub

Sub RightSlingShot_Timer
    Select Case RStep
        Case 3:RSLing1.Visible = 0:RSLing2.Visible = 1:sling1.TransZ = -20
        Case 4:RSLing2.Visible = 0:RSLing.Visible = 1:sling1.TransZ = 0:RightSlingShot.TimerEnabled = 0:lgi17.state=1:lgi18.state=1
    End Select
    RStep = RStep + 1
End Sub

Sub LeftSlingShot_Slingshot
	addscore 10
	playsound SoundFXDOF("left_slingshot",103,DOFPulse,DOFContactors)
    LSling.Visible = 0
    LSling1.Visible = 1
    sling2.TransZ = -10
    LStep = 0
    LeftSlingShot.TimerEnabled = 1
    lgi15.state=0:lgi16.state=0
End Sub

Sub LeftSlingShot_Timer
    Select Case LStep
        Case 3:LSLing1.Visible = 0:LSLing2.Visible = 1:sling2.TransZ = -20
        Case 4:LSLing2.Visible = 0:LSLing.Visible = 1:sling2.TransZ = 0:LeftSlingShot.TimerEnabled = 0:lgi16.state=1:lgi16.state=1
    End Select
    LStep = LStep + 1
End Sub

Sub LeftSlingshot_Init
	LSling1.Visible = 0:LSling2.Visible = 0: Sling2.TransZ = 0
End Sub

Sub RightSlingshot_Init
	RSling1.Visible = 0:RSling2.Visible = 0: Sling1.TransZ = 0
End Sub

'********** Triggers

sub trigger1_hit
	DOF 112, DOFPulse
    if not tilt Then
      if light1.state=0 then
        addscore 5000 ' always 5000
      else
        addscore 5000
        light1.state=0:light1w.state=1
        BumperLight2.state=1
      end if
      FlashBumpers()
      ChkSpinNSpot()
    end if
end sub

sub trigger2_hit
	DOF 111, DOFPulse
    if not tilt Then
      if light2.state=0 then
        addscore 500
      else
        addscore 5000
        light2.state=0:light2w.state=1
        if balls=3 then light3.state=0:light3w.state=1
      end if
      FlashBumpers()
      ChkSpinNSpot()
    end if
end sub

sub trigger3_hit
	DOF 113, DOFPulse
    if not tilt Then
      if light3.state=0 then
        addscore 500
      else
        addscore 5000
        light3.state=0:light3w.state=1
        if balls=3 then light2.state=0:light2w.state=1
      end if
      FlashBumpers()
      ChkSpinNSpot()
    end if
end sub

sub trigger4_hit
	DOF 110, DOFPulse
    if not tilt Then
      if light4.state=0 then
        addscore 500
      else
        addscore 5000
        light4.state=0:light4w.state=1:light4b.state=0
        BumperLight1.state=1
      end if
      FlashBumpers()
      ChkSpinNSpot()
    end if
end sub

sub trigger4b_hit
	DOF 110, DOFPulse
    if not tilt Then
      if light4.state=0 then
        addscore 500
      else
        addscore 5000
        light4.state=0:light4w.state=1:light4b.state=0
        BumperLight1.state=1
      end if
      FlashBumpers()
      ChkSpinNSpot()
    end if
end sub

sub trigger5_hit
	DOF 114, DOFPulse
    if not tilt Then
      if light5.state=0 then
        addscore 500
      else
        addscore 5000
        light5.state=0:light5w.state=1:light5b.state=0
        BumperLight3.state=1
      end if
      FlashBumpers()
      ChkSpinNSpot()
    end if
end sub

sub trigger5b_hit
	DOF 114, DOFPulse
    if not tilt Then
      if light5.state=0 then
        addscore 500
      else
        addscore 5000
        light5.state=0:light5w.state=1:light5b.state=0
        BumperLight3.state=1
      end if
      FlashBumpers()
      ChkSpinNSpot()
    end if
end sub

sub trigger6_hit
	DOF 123, DOFPulse
    if not tilt Then
      if light6.state=0 then
        addscore 5000 ' outlane
      else
        addscore 5000
        light6.state=0:light6w.state=1
        if balls=3 then light7.state=0:light7w.state=1
      end if
      FlashBumpers()
      ChkSpinNSpot()
    end if
end sub

sub trigger7_hit
	DOF 120, DOFPulse
    if not tilt Then
      if light7.state=0 then
        addscore 5000 ' outlane
      else
        addscore 5000
        light7.state=0:light7w.state=1
        if balls=3 then light6.state=0:light6w.state=1
      end if
      FlashBumpers()
      ChkSpinNSpot()
    end if
end sub

sub trigger8_hit
	DOF 116, DOFPulse
    if not tilt Then
      if light8.state=0 then
        addscore 500
      else
        addscore 5000
        light8.state=0:light8w.state=1
        if balls=3 then light9.state=0:light9w.state=1
      end if
      FlashBumpers()
      ChkSpinNSpot()
    end if
end sub

sub trigger9_hit
	DOF 119, DOFPulse
    if not tilt Then
      if light9.state=0 then
        addscore 500
      else
        addscore 5000
        light9.state=0:light9w.state=1
        if balls=3 then light8.state=0:light8w.state=1
      end if
      FlashBumpers()
      ChkSpinNSpot()
    end if
end sub

sub trigger10_hit
	DOF 117, DOFPulse
    if not tilt Then
      if light10.state=0 then
        addscore 500
      else
        addscore 5000
        light10.state=0:light10w.state=1
      end if
      if lightSpec1.state=1 then Getspecial()
      FlashBumpers()
      ChkSpinNSpot()
    end if
end sub

sub trigger11_hit
	DOF 118, DOFPulse
    if not tilt Then
      if light11.state=0 then
        addscore 500
      else
        addscore 5000
        light11.state=0:light11w.state=1
      end if
      if lightSpec2.state=1 then Getspecial()
      FlashBumpers()
      ChkSpinNSpot()
    end if
end sub

Sub GetSpecial
    PlaySound SoundFXDOF("knock",128,DOFPulse,DOFKnocker)
	addcredit
End Sub

' =================

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
	    special=special+1:if special>1 then special=0   ' Special light moves every 100
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
		PlaySound SoundFXDOF("knock",128,DOFPulse,DOFKnocker)
		DOF 127, DOFPulse
    end if
    if score(player)=>replay2 and rep(player)=1 then
		addcredit
		rep(player)=2
		PlaySound SoundFXDOF("knock",128,DOFPulse,DOFKnocker)
		DOF 127, DOFPulse
    end if
    if score(player)=>replay3 and rep(player)=2 then
		addcredit
		rep(player)=3
		PlaySound SoundFXDOF("knock",128,DOFPulse,DOFKnocker)
		DOF 127, DOFPulse
    end if
end sub

Sub CheckTilt
	If Tilttimer.Enabled = True Then
	 TiltSens = TiltSens + 1
	 if TiltSens = 3 Then
	   Tilt = True
	    If Buccaneer.ShowDT then tilttxt.visible=1
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
    bumper1.hashitevent=False
    bumper2.hashitevent=False
	bumper3.hashitevent=False
    bumper4.hashitevent=False
    bumper5.hashitevent=False
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

Sub a_Spinner_Spin(idx)
    PlaySound SoundFXDOF("fx_spinner",129,DOFPulse,DOFKnocker),0,.25,0,0.25
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


' nFozzy Flipper Code
dim returnspeed, lfstep, rfstep
returnspeed = leftflipper.return
lfstep = 1
rfstep = 1

sub leftflipper_timer()
	select case lfstep
		Case 1: Leftflipper.return = returnspeed * 0.6 :lfstep = lfstep + 1
		Case 2: Leftflipper.return = returnspeed * 0.7 :lfstep = lfstep + 1
		Case 3: Leftflipper.return = returnspeed * 0.8 :lfstep = lfstep + 1
		Case 4: Leftflipper.return = returnspeed * 0.9 :lfstep = lfstep + 1
		Case 5: Leftflipper.return = returnspeed * 1 :lfstep = lfstep + 1
		Case 6: Leftflipper.timerenabled = 0 : lfstep = 1
	end select
end sub

sub rightflipper_timer()
	select case rfstep
		Case 1: Rightflipper.return = returnspeed * 0.6 :rfstep = rfstep + 1
		Case 2: Rightflipper.return = returnspeed * 0.7 :rfstep = rfstep + 1
		Case 3: Rightflipper.return = returnspeed * 0.8 :rfstep = rfstep + 1
		Case 4: Rightflipper.return = returnspeed * 0.9 :rfstep = rfstep + 1
		Case 5: Rightflipper.return = returnspeed * 1 :rfstep = rfstep + 1
		Case 6: Rightflipper.timerenabled = 0 : rfstep = 1
	end select
end sub


sub savehs
    savevalue "Buccaneer", "credit", credit
    savevalue "Buccaneer", "hiscore", hisc
    savevalue "Buccaneer", "match", matchnumb
    savevalue "Buccaneer", "score1", score(1)
	savevalue "Buccaneer", "balls", balls
	savevalue "Buccaneer", "freeplay", freeplay
end sub

sub loadhs
    dim temp
	temp = LoadValue("Buccaneer", "credit")
    If (temp <> "") then credit = CDbl(temp)
    temp = LoadValue("Buccaneer", "hiscore")
    If (temp <> "") then hisc = CDbl(temp)
    temp = LoadValue("Buccaneer", "match")
    If (temp <> "") then matchnumb = CDbl(temp)
    'temp = LoadValue("Buccaneer", "score1") mod 100000
    If (temp <> "") then score(1) = CDbl(temp)
    temp = LoadValue("Buccaneer", "balls")
    If (temp <> "") then balls = CDbl(temp)
    temp = LoadValue("Buccaneer", "freeplay")
    If (temp <> "") then freeplay = CDbl(temp)
end sub

Sub Buccaneer_Exit()
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
  tmp = ball.y * 2 / Table1.height-1
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
  If Table1.VersionMinor > 3 OR Table1.VersionMajor > 10 Then
    PlaySound("fx_collide"), 0, Csng(velocity) ^2 / 200, Pan(ball1), 0, Pitch(ball1), 0, 0, AudioFade(ball1)
  Else
    PlaySound("fx_collide"), 0, Csng(velocity) ^2 / 200, Pan(ball1), 0, Pitch(ball1), 0, 0
  End if
End Sub

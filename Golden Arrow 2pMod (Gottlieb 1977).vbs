'**************************************************************
'          GOLDEN ARROW
'
'	Layer information
'		1	general - most stuff
'		2	triggers
'		4	bumper primitives
'		6	playfield ambient lights
'		7	plastics
'		8	table lights
'
' vp10 assembled and scripted by BorgDog, 2015
' thanks to hauntfreaks and sliderpoint for graphics updates
' thansk to sliderpoint for answering all my newb questions in this the first table I made
'
' thanks to PinballPerson for the original coding of golden Arrow that I build upon.
'**************************************************************

' Thalamus 2018-07-23
' Added/Updated "Positional Sound Playback Functions" and "Supporting Ball & Sound Functions"
' No special SSF tweaks yet.

Dim DesktopMode: DesktopMode = Table1.ShowDT
dim ballinplay
dim ballrenabled
dim rst
dim eg
dim credit
dim score(2)
dim player
dim players
dim match(9)
dim rep
dim tilt
dim tiltsens
dim state
dim scn
dim scn1
dim bell
dim points
dim matchnumb
dim replay1
dim replay2
dim replay3
dim hisc
dim arrow(10)
dim numb(10)
dim numbstate(2,10)
dim apos(2)
dim ac(2)
dim sa(2)
dim spstate(2)
Dim RStep, Lstep

 Dim Controller
 LoadController
 Sub LoadController()
	If DesktopMode = False Then
		Set Controller = CreateObject("B2S.Server")
		Controller.B2SName = "Golden Arrow VPX"  ' *****!!!!This must match the name of your directb2s file!!!!
		Controller.Run()
	End If
 End Sub

sub table1_init
    set match(0)=m0
    set match(1)=m1
    set match(2)=m2
    set match(3)=m3
    set match(4)=m4
    set match(5)=m5
    set match(6)=m6
    set match(7)=m7
    set match(8)=m8
    set match(9)=m9
	set arrow(1)=a1a
	set arrow(2)=a2
	set arrow(3)=a3
	set arrow(4)=a4a
	set arrow(5)=a5a
	set arrow(6)=a6a
	set arrow(7)=a7
	set arrow(8)=a8
	set arrow(9)=a9a
	set arrow(10)=a10a
	set numb(1)=n1a
	set numb(2)=n2
	set numb(3)=n3
	set numb(4)=n4a
	set numb(5)=n5a
	set numb(6)=n6a
	set numb(7)=n7
	set numb(8)=n8
	set numb(9)=n9a
	set numb(10)=n10a
	player=1
    replay1=120000
    replay2=150000
    replay3=190000
    bumper1light.state=lightstateoff
    bumper2light.state=lightstateoff
    bumper3light.state=lightstateoff
    bumper1light1.state=lightstateoff
    bumper2light1.state=lightstateoff
    bumper3light1.state=lightstateoff
    loadhs
    if hisc="" then hisc=50000
    hstxt.text=hisc
    if credit="" then credit=0
    credittxt.text=credit
 	If DesktopMode = False Then
		Controller.B2ssetCredits Credit
		Controller.B2SSetScore 4, hisc
		Controller.B2ssetMatch 34, Matchnumb*10
		scoreReel1.visible = False
        ScoreReel2.visible = False
        gamov.visible = False
		M0.Visible = false
		M1.visible = False
		M2.visible = False
		M3.visible = False
		M4.visible = False
		M5.visible = False
		M6.visible = False
		M7.visible = False
		M8.visible = False
		M9.visible = False
		shoot1.visible = False
		shoot2.visible = False
   		tilttxt.visible = False
		hstxt.visible = False
		credittxt.visible = False
		bip1.visible = False
		bip2.visible = False
		bip3.visible = False
		bip4.visible = False
		bip5.visible = False
		canplay.visible = False
		HighScore.visible = False
	End If

   select case(matchnumb)
		case 0:
		m0.text="00"
		case 1:
		m1.text="10"
		case 2:
		m2.text="20"
		case 3:
		m3.text="30"
		case 4:
		m4.text="40"
		case 5:
		m5.text="50"
		case 6:
		m6.text="60"
		case 7:
		m7.text="70"
		case 8:
		m8.text="80"
		case 9:
		m9.text="90"
    end select
    scorereel1.setvalue(score(1))
    scorereel2.setvalue(score(2))
    if apos(1)<1 or apos(1)>16 then apos(1)=1
    if apos(2)<1 or apos(2)>16 then apos(2)=1
end sub


Sub Table1_KeyDown(ByVal keycode)

    if keycode=AddCreditKey then
		playsound "coin3"
		if state=false then
			credittxt.text=credit
		end if
		If DesktopMode = False Then Controller.B2SSetScore 4, hisc
		coindelay.enabled=true
    end if


    if keycode=StartGameKey and credit>0 then
	  If DesktopMode = False Then Controller.B2SSetScore 4, hisc
	  if state=false then
		credit=credit-1
		playsound "click"
		credittxt.text=credit
		ballinplay=1
		If DesktopMode = False Then
			Controller.B2ssetCredits Credit
			Controller.B2sStartAnimation "Startup"
			Controller.B2ssetballinplay 32, Ballinplay
			Controller.B2ssetplayerup 30, 1
			Controller.B2ssetcanplay 31, 1
			Controller.B2ssetdata 1, 1
		End If
		tilt=false
		state=true
		rst=0
		CanPlay.Text="One Player"
		playsound "initialize"
		resettimer.enabled=true
		players=1
	  else if state=true and players < 2 and Ballinplay=1 then
		credit=credit-1
		players=players+1
		credittxt.text=credit
		If DesktopMode = False then
			Controller.B2ssetCredits Credit
			Controller.B2sStartAnimation "Startup"
			Controller.B2ssetcanplay 31, 2
		End If
		CanPlay.Text="Two Players"
		playsound "click"
	   end if
	  end if
	end if

	If keycode = PlungerKey Then
		Plunger.PullBack
		PlaySound "plungerpull"
	End If

    if tilt=false and state=true then
	  If keycode = LeftFlipperKey Then
		LeftFlipper.RotateToEnd
		PlaySound "FlipperUp"
		PlaySound "Buzz",-1,.1
	  End If

	  If keycode = RightFlipperKey Then
		RightFlipper.RotateToEnd
        PlaySound "FlipperUp"
		PlaySound "Buzz",-1,.1
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
	  end If

	If keycode = MechanicalTilt Then
		mechchecktilt
	End If

    end if

End Sub

Sub Table1_KeyUp(ByVal keycode)

	If keycode = PlungerKey Then
		Plunger.Fire
		playsound "plunger"
	End If

	If keycode = LeftFlipperKey Then
		LeftFlipper.RotateToStart
		if tilt= false and state=true then
			PlaySound "FlipperDown"
			StopSound "Buzz"
		end if
	End If

	If keycode = RightFlipperKey Then
		RightFlipper.RotateToStart
        if tilt= false and state=true then
			PlaySound "FlipperDown"
			StopSound "Buzz"
		end if
	End If

End Sub

sub flippertimer_timer()
	LFlip.RotY = LeftFlipper.CurrentAngle
	RFlip.RotY = RightFlipper.CurrentAngle
	PrimGate1.Rotz = gate.CurrentAngle
end sub

Sub PairedlampTimer_timer
	a1b.state = a1a.state
	a4b.state = a4a.state
	a5b.state = a5a.state
	a6b.state = a6a.state
	a9b.state = a9a.state
	a10b.state = a10a.state
	n1b.state = n1a.state
	n4b.state = n4a.state
	n5b.state = n5a.state
	n6b.state = n6a.state
	n9b.state = n9a.state
	n10b.state = n10a.state
	rtl.state = ltl.state
end sub


sub coindelay_timer
	addcredit
    coindelay.enabled=false
end sub

sub resettimer_timer
    rst=rst+1
    scorereel1.resettozero
    scorereel2.resettozero
    If DesktopMode = False then
		Controller.B2SSetScore 1, score(1)
		Controller.B2SSetScore 2, score(2)
	End If
    if rst=18 then
    playsound "kickerkick"
    end if
    if rst=22 then
    newgame
    resettimer.enabled=false
    end if
end sub

sub newgame
    bumper1.force=7
    bumper2.force=7
    bumper3.force=7
	player=1
	shoot1.text="Player 1"
    score(1)=0
	score(2)=0
	If DesktopMode = False then
		Controller.B2SSetScore 1, score(1)
		Controller.B2SSetScore 2, score(2)
	End If
    eg=0
    rep=0
    sa(1)=0
	sa(2)=0
	spstate(1)=0
	spstate(2)=0
    sp.state=lightstateoff
    ltl.state=lightstateoff
    fivek.state=lightstateon
    bumper1light.state=lightstateon
    bumper2light.state=lightstateon
    bumper3light.state=lightstateon
    bumper1light1.state=lightstateon
    bumper2light1.state=lightstateon
    bumper3light1.state=lightstateon
	PlightRM.state=lightstateon
	TlightRM.State=lightstateon
	PlightLM.state=lightstateon
	TlightLM.state=lightstateon
	PlightLL.state=lightstateon
	PlightRL.state=lightstateon
	TlightLL.state=lightstateon
	TlightRL.state=lightstateon
	PlightMR.state=lightstateon
	TlightMR.state=lightstateon
	PlightML.state=lightstateon
	TlightML.state=lightstateon
	PlightLU.state=lightstateon
	TlightLU.state=lightstateon
	PlightRU.state=lightstateon
	TlightRU.state=lightstateon
	For Each Light in LaneLights
		Light.state=lightstateon
	Next
	for i=1 to 10
		arrow(i).state=lightstateoff
		numb(i).state=lightstateon
		numbstate(1,i)=1
		numbstate(2,i)=1
	next
    movearrow
    gamov.text=" "
    tilttxt.text=" "
    If DesktopMode = False then
		Controller.B2SSetGameOver 35,0
		Controller.B2SSetTilt 33,0
		Controller.B2SSetMatch 34,0
		Controller.B2ssetdata 11, 0
		Controller.B2ssetdata 7,0
		Controller.B2ssetdata 8,0
	End If
    bip5.text=" "
    bip1.text="1"
    for i=0 to 9
      match(i).text=" "
    next
    nb.CreateBall
   	nb.kick 135,4,0
end sub

sub newball
    if sa(player)=1 then
      for i=1 to 10
		numb(i).state=lightstateon
		numbstate(player,i)=1
      next
      sp.state=lightstateoff
	  spstate(player)=0
      rtl.state=lightstateoff
      ltl.state=lightstateoff
      sa(player)=0
	 else
		for i=1 to 10
			if numbstate(player,i)=1 then
			    numb(i).state=lightstateon
			  else
				numb(i).state=lightstateoff
			end if
			arrow(i).state=lightstateoff
		next
		arrow(apos(player)).state=lightstateon
	end if
end sub

Sub Drain_Hit()
	playsound "drainshorter"
	Drain.DestroyBall
		if players=1 or player=2 then
			player=1
			If DesktopMode = False then Controller.B2ssetplayerup 30, 1
			shoot1.text="Player 1"
			shoot2.text=" "
			nextball
		  else
			player=2
			If DesktopMode = False then Controller.B2ssetplayerup 30, 2
			shoot2.text="Player 2"
			shoot1.text=" "
			nextball
		end if
'	end if
End Sub

sub nextball
    if tilt=true then
    bumper1.force=7
    bumper2.force=7
    bumper3.force=7
    tilt=false
    tilttxt.text=" "
		If DesktopMode = False then
			Controller.B2SSetTilt 33,0
			Controller.B2ssetdata 1, 1
		End If
    end if
	if player=1 then ballinplay=ballinplay+1
	If DesktopMode = False then Controller.B2ssetballinplay 32, Ballinplay
	if ballinplay>5 then
		playsound "motorleer"
		eg=1
		ballreltimer.enabled=true
	  else
		if state=true and tilt=false then
		  newball
		  playsound "kickerkick"
		  ballreltimer.enabled=true
		end if
	  select case (ballinplay)
		case 1:
		bip1.text="1"
		case 2:
		bip1.text=" "
		bip2.text="2"
		case 3:
		bip2.text=" "
		bip3.text="3"
		case 4:
		bip3.text=" "
		bip4.text="4"
		case 5:
		bip4.text=" "
		bip5.text="5"
	  end select
	end if
End Sub

sub ballreltimer_timer
    if eg=1 then
      matchnum
	  bip5.text=" "
	  state=false
	  gamov.text="GAME OVER"
	  CanPlay.text=" "
	  shoot1.text=" "
	  shoot2.text=" "
	  if score(1)>hisc then hisc=score(1)
	  if score(2)>hisc then hisc=score(2)
	  hstxt.text=hisc
	If DesktopMode = False then
      Controller.B2SSetGameOver 35,1
      Controller.B2ssetballinplay 32, 0
	  Controller.B2sStartAnimation "EOGame"
	  Controller.B2SSetScore 4, hisc
	  Controller.B2ssetcanplay 31, 0
	  Controller.B2ssetcanplay 30, 0
	End If
      bumper1light.state=lightstateoff
      bumper2light.state=lightstateoff
      bumper3light.state=lightstateoff
      bumper1light1.state=lightstateoff
      bumper2light1.state=lightstateoff
      bumper3light1.state=lightstateoff
	  PlightRM.state=lightstateoff
	  TlightRM.State=lightstateoff
	  PlightLM.state=lightstateoff
	  TlightLM.state=lightstateoff
	  PlightLL.state=lightstateoff
	  PlightRL.state=lightstateoff
	  TlightLL.state=lightstateoff
	  TlightRL.state=lightstateoff
	  PlightMR.state=lightstateoff
	  TlightMR.state=lightstateoff
	  PlightML.state=lightstateoff
	  TlightML.state=lightstateoff
	  PlightLU.state=lightstateoff
	  TlightLU.state=lightstateoff
	  PlightRU.state=lightstateoff
	  TlightRU.state=lightstateoff
	For each Light in LaneLights
	  Light.state=lightstateoff
	next
	  savehs
	  ballreltimer.enabled=false
    else
      nb.CreateBall
	  nb.kick 135,4,0
      ballreltimer.enabled=false
    end if
end sub

sub matchnum
    select case(matchnumb)
      case 0:
      m0.text="00"
      case 1:
      m1.text="10"
      case 2:
      m2.text="20"
      case 3:
      m3.text="30"
      case 4:
      m4.text="40"
      case 5:
      m5.text="50"
      case 6:
      m6.text="60"
      case 7:
      m7.text="70"
      case 8:
      m8.text="80"
      case 9:
      m9.text="90"
    end select
  If DesktopMode = False then Controller.B2SSetMatch 34,Matchnumb*10
  For i=1 to players
    if (matchnumb*10)=(score(i) mod 100) then
		addcredit
		playsound "knocke"
	end if
  next
end sub

sub addscore(points)
  if tilt=false then
    bell=0
    if points = 10 or points = 100 or points=1000 or points=10000 then scn=1

    if player=1 then scorereel1.addvalue(points)
    if player=2 then scorereel2.addvalue(points)
	If DesktopMode = False Then Controller.B2SSetScore player, score(player) + Points
    if points = 10 then bell=1
    if points = 1000 then bell=3
    if points = 100 then bell=2
    if points = 10000 then bell=3
    if points = 5000 then
      bell=3
      scn=5
    end if
    if points = 500 then
      bell=2
      scn=5
    end if
    scn1=0
    scntimer.enabled=true
    score(player)=score(player)+points
  end if
    if score(player)=>replay1 and rep=0 then
      playsound "knocke"
	  addcredit
      rep=1
    end if
    if score(player)=>replay2 and rep=1 then
      playsound "knocke"
      addcredit
      rep=2
    end if
    if score(player)=>replay3 and rep=2 then
      playsound "knocke"
      addcredit
      rep=3
    end if
end sub

Sub addcredit
      credit=credit+1
      if credit>25 then credit=25
	  credittxt.text=credit
	  If DesktopMode = False Then Controller.B2ssetCredits Credit
      playsound "click"
End sub

sub scntimer_timer
    scn1=scn1 + 1
    if bell=2 then
      playsound "ding2"
      matchnumb=matchnumb+1
    if matchnumb=10 then matchnumb=0
    end if
    if bell=3 then playsound "ding3"
    if bell=1 then playsound "ding1"
    if scn1=scn then
      scntimer.enabled=false
    end if
end sub

Sub CheckTilt
	If Tilttimer.Enabled = True Then
	 TiltSens = TiltSens + 1
	 if TiltSens = 2 Then
	   Tilt = True
	   tilttxt.text="TILT"
       	If DesktopMode = False Then Controller.B2SSetTilt 33,1
       	If DesktopMode = False Then Controller.B2ssetdata 1, 0
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
       	If DesktopMode = False Then Controller.B2SSetTilt 33,1
       	If DesktopMode = False Then Controller.B2ssetdata 1, 0
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
end sub

Sub RightSlingShot_Slingshot
    PlaySound "left_slingshot", 0, 1, 0.05, 0.05
	addscore 10
    RSling.Visible = 0
    RSling1.Visible = 1
    sling1.TransZ = -8
    RStep = 1
    RightSlingShot.TimerEnabled = 1
'	PLightRL.State = 0
'	TLightRL.State = 0
End Sub

Sub RightSlingShot_Timer
    Select Case RStep
        Case 3:RSLing1.Visible = 0:RSLing2.Visible = 1:sling1.TransZ = -8
        Case 4:sling1.TransZ = 0:RSLing2.Visible = 0:RSLing.Visible = 1:RightSlingShot.TimerEnabled = 0:PLightRL.State = 1:TLightRL.State = 1
    End Select
    RStep = RStep + 1
End Sub

Sub LeftSlingShot_Slingshot
    PlaySound "right_slingshot",0,1,-0.05,0.05
	addscore 10
    LSling.Visible = 0
    LSling1.Visible = 1
    sling2.TransZ = -8
    LStep = 1
    LeftSlingShot.TimerEnabled = 1
'	PLightLL.State = 0
'	TLightLL.State = 0
End Sub

Sub LeftSlingShot_Timer
    Select Case LStep
        Case 3:LSLing1.Visible = 0:LSLing2.Visible = 1:sling2.TransZ = -8
        Case 4:sling2.TransZ = 0:LSLing2.Visible = 0:LSLing.Visible = 1:LeftSlingShot.TimerEnabled = 0:PLightLL.State = 1:TLightLL.State = 1
    End Select
    LStep = LStep + 1
End Sub

'************************************
' What you need to add to your table
'************************************

' a timer called CollisionTimer. With a fast interval, like from 1 to 10
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


' *********************************************************************

sub bumper1_hit
    if tilt=false then
	  playsound "jet2"
	  addscore 100
	  If FlashB.Enabled = False then
		Bumper1Light.State = 0
		Bumper1Light1.State = 0
		Me.TimerEnabled = 1
	  End if
	  If TenKTimer.enabled = false then
		if ltl.state=lightstateon then
			ltl.state = lightstateoff
			tenktimer.enabled = 1
		  else
			ltl.state = lightstateon
			tenktimer.enabled = 1
		end if
	  end if
	end if
end sub

sub Bumper1_Timer
	Bumper1Light.State = 1
	Bumper1Light1.State = 1
	Me.Timerenabled = 0
End Sub

sub bumper2_hit
    if tilt=false then playsound "jet2"
	if (bumper2light.state)=lightstateon then
		addscore 1000
	else
		addscore 100
    end if
	If FlashB.Enabled = false then
		Bumper2Light.State = 0
		Bumper2Light1.State = 0
		Me.Timerenabled = 1
	end if
end sub

sub Bumper2_Timer
	Bumper2Light.State = 1
	Bumper2Light1.State = 1
	Me.Timerenabled = 0
End Sub

sub bumper3_hit
    if tilt=false then
	  playsound "jet2"
      addscore 100
	  If FlashB.Enabled = false then
		Bumper3Light.State = 0
		Bumper3Light1.State = 0
		Me.TimerEnabled = 1
	  End if
	  If TenKTimer.enabled = false then
		if ltl.state=lightstateon then
			ltl.state = lightstateoff
			tenktimer.enabled = 1
		  else
			ltl.state = lightstateon
			tenktimer.enabled = 1
		end if
	  end if
	end if
end sub

sub Bumper3_Timer
	Bumper3Light.State = 1
	Bumper3Light1.State = 1
	Me.Timerenabled = 0
End Sub

sub TenKTimer_Timer
	TenKTimer.enabled = false
end Sub

sub ctrig_hit
    movearrow
    addscore 100
end sub

sub spinner1_spin
    movearrow
    addscore 100
end sub

sub spinner2_spin
    movearrow
    addscore 100
end sub

sub tlt_hit
    if tilt=false then
	FlashBumpers
	 PlightLU.state=lightstateoff
	 TlightLU.state=lightstateoff
	 Me.TimerEnabled = 1
	 if (ltl.state)=lightstateon then
	    addscore 10000
		playsound "dabell"
	  else
	   addscore 1000
	 end if
	end if
end sub

sub tlt_Timer
	PlightLU.State=lightstateon
	TlightLU.state=lightstateon
	Me.Timerenabled = 0
End Sub

sub trt_hit
    if tilt=false then
	  FlashBumpers
	  PlightRU.state=lightstateoff
	  TlightRU.state=lightstateoff
	  Me.TimerEnabled = 1
      if (rtl.state)=lightstateon then
		playsound "dabell"
        addscore 10000
      else
        addscore 1000
      end if
	end if
end sub

sub trt_Timer
	PlightRU.State=lightstateon
	TlightRU.state=lightstateon
	Me.Timerenabled = 0
End Sub

sub t1a_hit
	flashbumpers
    if tilt=false then
	playsound "ding3"
	PlightMR.state=lightstateoff
	TlightMR.state=lightstateoff
	Me.TimerEnabled = 1
    if (n1a.state)=lightstateon then
		n1a.state=lightstateoff
		numbstate(player,1)=0
		addscore 5000
		checkaward
    else
		addscore 500
    end if
    if apos(player)=1 then
    if spstate(player)=1 then
    playsound "knocke"
	addcredit
    fivekdelay.enabled=true
    else
	playsound "dabell"
    fivekdelay.enabled=true
    end if
    end if
    end if
end sub

sub t1a_Timer
	PlightMR.State=lightstateon
	TlightMR.state=lightstateon
	Me.Timerenabled = 0
End Sub

sub t1b_hit
    if tilt=false then
		FlashBumpers
		if (n1b.state)=lightstateon then
			n1a.state=lightstateoff
			numbstate(player,1)=0
			addscore 5000
			checkaward
		  else
			addscore 500
		end if
		if apos(player)=1 then
			if spstate(player)=1 then
				playsound "knocke"
				addcredit
				fivekdelay.enabled=true
			else
				fivekdelay.enabled=true
				playsound "dabell"
			end if
		end if
    end if
end sub

sub t2_hit
    if tilt=false then
		FlashBumpers
		if (n2.state)=lightstateon then
			n2.state=lightstateoff
			numbstate(player,2)=0
			addscore 5000
			checkaward
		else
			addscore 500
		end if
		if apos(player)=2 then
			if spstate(player)=1 then
				playsound "knocke"
				addcredit
				fivekdelay.enabled=true
			  else
				fivekdelay.enabled=true
				playsound "dabell"
			end if
		end if
    end if
end sub

sub t3_hit
    if tilt=false then
	Flashbumpers
    if (n3.state)=lightstateon then
		n3.state=lightstateoff
		numbstate(player,3)=0
		addscore 5000
		checkaward
    else
		addscore 500
    end if
    if apos(player)=3 then
    if spstate(player)=1 then
      playsound "knocke"
	  addcredit
      fivekdelay.enabled=true
    else
      fivekdelay.enabled=true
	  playsound "dabell"
    end if
    end if
    end if
end sub

sub t4a_hit
    if tilt=false then
	flashbumpers
	playsound "ding3"
	PlightML.state=lightstateoff
	TlightML.state=lightstateoff
	Me.TimerEnabled = 1
    if (n4a.state)=lightstateon then
		n4a.state=lightstateoff
		numbstate(player,4)=0
		addscore 5000
		checkaward
    else
		addscore 500
    end if
    if apos(player)=4 then
		if spstate(player)=1 then
		playsound "knocke"
		addcredit
		fivekdelay.enabled=true
    else
		fivekdelay.enabled=true
		playsound "dabell"
    end if
    end if
    end if
end sub

sub t4a_Timer
	PlightML.State=lightstateon
	TlightML.state=lightstateon
	Me.Timerenabled = 0
End Sub

sub t4b_hit
    if tilt=false then
	FlashBumpers
	    if (n4b.state)=lightstateon then
			n4a.state=lightstateoff
			numbstate(player,4)=0
			addscore 5000
			checkaward
		else
			addscore 500
		end if
    if apos(player)=4 then
		if spstate(player)=1 then
			playsound "knocke"
			addcredit
			fivekdelay.enabled=true
		  else
			fivekdelay.enabled=true
			playsound "dabell"
		end if
    end if
    end if
end sub

sub t5a_hit
    if tilt=false then
	FlashBumpers
    if (n5a.state)=lightstateon then
		n5a.state=lightstateoff
		numbstate(player,5)=0
		addscore 5000
		checkaward
		if apos(player)=5 then
			addscore 5000
			playsound "dabell"
		end if
    else
		addscore 500
    end if
    if apos(player)=5 then
		if spstate(player)=1 then
			playsound "knocke"
			addcredit
			fivekdelay.enabled=true
		else
			playsound "dabell"
			fivekdelay.enabled=true
		end if
    end if
    end if
end sub

sub t5b_hit
    if tilt=false then
	flashbumpers
    if (n5b.state)=lightstateon then
		n5a.state=lightstateoff
		numbstate(player,5)=0
		addscore 5000
		checkaward
      else
		addscore 500
    end if
    if apos(player)=5 then
    if spstate(player)=1 then
    playsound "knocke"
	addcredit
    fivekdelay.enabled=true
    else
    fivekdelay.enabled=true
	playsound "dabell"
    end if
    end if
    end if
end sub

sub t6a_hit
    if tilt=false then
	FlashBumpers
    if (n6a.state)=lightstateon then
		n6a.state=lightstateoff
		numbstate(player,6)=0
		addscore 5000
		checkaward
		else
		addscore 500
    end if
    if apos(player)=6 then
		if spstate(player)=1 then
			playsound "knocke"
			addcredit
			fivekdelay.enabled=true
		else
			playsound "dabell"
			fivekdelay.enabled=true
		end if
    end if
    end if
end sub

sub t6b_hit
 if tilt=false then
	flashbumpers
   if (n6b.state)=lightstateon then
		n6a.state=lightstateoff
		numbstate(player,6)=0
		addscore 5000
		checkaward
    else
		addscore 500
   end if
    if apos(player)=6 then
      if spstate(player)=1 then
        playsound "knocke"
		addcredit
        fivekdelay.enabled=true
      else
        fivekdelay.enabled=true
		playsound "dabell"
      end if
    end if
 end if
end sub

sub t7_hit
    if tilt=false then
	flashbumpers
	playsound "ding3"
	PlightRM.state=lightstateoff
	TlightRM.state=lightstateoff
	Me.TimerEnabled = 1
    if (n7.state)=lightstateon then
		n7.state=lightstateoff
	numbstate(player,7)=0
    addscore 5000
    checkaward
    else
    addscore 500
    end if
    if apos(player)=7 then
    if spstate(player)=1 then
    playsound "knocke"
	addcredit
    fivekdelay.enabled=true
    else
	playsound "dabell"
    fivekdelay.enabled=true
    end if
    end if
    end if
end sub

sub t7_Timer
	PlightRM.State=lightstateon
	TlightRM.State=lightstateon
	Me.Timerenabled = 0
End Sub

sub t8_hit
    if tilt=false then
	flashbumpers
	playsound "ding3"
	PlightLM.state=lightstateoff
	TlightLM.state=lightstateoff
	Me.TimerEnabled = 1
    if (n8.state)=lightstateon then
		n8.state=lightstateoff
	numbstate(player,8)=0
    addscore 5000
    checkaward
    else
    addscore 500
    end if
    if apos(player)=8 then
    if spstate(player)=1 then
    playsound "knocke"
	addcredit
    fivekdelay.enabled=true
    else
	playsound "dabell"
    fivekdelay.enabled=true
    end if
    end if
    end if
end sub

sub t8_Timer
	PlightLM.State=lightstateon
	TlightLM.State=lightstateon
	Me.Timerenabled = 0
End Sub

sub t9a_hit
  if tilt=false then
	FlashBumpers
    if (n9a.state)=lightstateon then
		n9a.state=lightstateoff
		numbstate(player,9)=0
		addscore 5000
		checkaward
    else
		addscore 500
    end if
    if apos(player)=9 then
		if spstate(player)=1 then
			playsound "knocke"
			addcredit
			fivekdelay.enabled=true
		else
			playsound "dabell"
			fivekdelay.enabled=true
		end if
    end if
  end if
end sub

sub t9b_hit
    if tilt=false then
	flashbumpers
    if (n9b.state)=lightstateon then
		n9a.state=lightstateoff
 	numbstate(player,9)=0
    addscore 5000
    checkaward
    else
    addscore 500
    end if
    if apos(player)=9 then
    if spstate(player)=1 then
    playsound "knocke"
	addcredit
    fivekdelay.enabled=true
    else
	playsound "dabell"
    fivekdelay.enabled=true
    end if
    end if
    end if
end sub

sub t10a_hit
    if tilt=false then
	FlashBumpers
    if (n10a.state)=lightstateon then
		n10a.state=lightstateoff
		numbstate(player,10)=0
		addscore 5000
		checkaward
    else
		addscore 500
    end if
    if apos(player)=10 then
		if spstate(player)=1 then
			playsound "knocke"
			addcredit
			fivekdelay.enabled=true
		else
			playsound "dabell"
			fivekdelay.enabled=true
		end if
    end if
    end if
end sub

sub t10b_hit
    if tilt=false then
	flashbumpers
    if (n10b.state)=lightstateon then
		n10a.state=lightstateoff
	numbstate(player,10)=0
    addscore 5000
    checkaward
    else
    addscore 500
    end if
    if apos(player)=10 then
		if spstate(player)=1 then
			playsound "knocke"
			addcredit
			fivekdelay.enabled=true
		  else
			playsound "dabell"
			fivekdelay.enabled=true
		  end if
		end if
    end if
end sub

sub FlashBumpers
	For each light in BumperFlashLights
	  Light.State=0
	next
	FlashB.enabled=1
end sub

sub FlashB_timer
	For each light in BumperFlashLights
	  Light.State=1
	next
	FlashB.enabled=0
end sub

sub fivekdelay_timer
    addscore 5000
    fivekdelay.enabled=false
end sub

sub movearrow
    if tilt=false then
	  for i = 1 to 10
		arrow(apos(player)).state=lightstateoff
	  next
      apos(player)=apos(player)+1
      if apos(player)>10 then apos(player)=1
	  arrow(apos(player)).state=lightstateon
	end if
end sub

sub checkaward
    ac(player)=0
    for i=1 to 10
      if numbstate(player,i)=0 then ac(player)=ac(player)+1
    next
    if ac(player)=10 then
      sp.state=lightstateon
	  spstate(player)=1
      sa(player)=1
    end if
end sub

sub savehs

    savevalue "GoldenArrow", "credit", credit
    savevalue "GoldenArrow", "hiscore", hisc
    savevalue "GoldenArrow", "match", matchnumb

    savevalue "GoldenArrow", "score1", score(1)
    savevalue "GoldenArrow", "arrowpos1", apos(1)
    savevalue "GoldenArrow", "score2", score(2)
    savevalue "GoldenArrow", "arrowpos2", apos(2)

end sub

sub loadhs
    dim temp

	temp = LoadValue("GoldenArrow", "credit")
    If (temp <> "") then credit = CDbl(temp)
    temp = LoadValue("GoldenArrow", "hiscore")
    If (temp <> "") then hisc = CDbl(temp)
    temp = LoadValue("GoldenArrow", "match")
    If (temp <> "") then matchnumb = CDbl(temp)

    temp = LoadValue("GoldenArrow", "score1")
    If (temp <> "") then score(1) = CDbl(temp)
    temp = LoadValue("GoldenArrow", "arrowpos1")
    If (temp <> "") then apos(1) = CDbl(temp)
    temp = LoadValue("GoldenArrow", "score2")
    If (temp <> "") then score(2) = CDbl(temp)
    temp = LoadValue("GoldenArrow", "arrowpos2")
    If (temp <> "") then apos(2) = CDbl(temp)
end sub

sub ballhome_hit
	ballrenabled=1
end sub

sub ballrel_hit
	if ballrenabled=1 then
		playsound "ballrelease"
		ballrenabled=0
	end if
end sub

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
'    JP's VP10 Collision & Rolling Sounds
'*****************************************

Const tnob = 1 ' total number of balls
ReDim rolling(tnob)

Initcollision

Sub Initcollision
    Dim i
    For i = 0 to tnob
         rolling(i) = False
    Next
End Sub

Sub CollisionTimer_Timer()
    Dim BOT, B
    BOT = GetBalls

    ' rolling

	For B = UBound(BOT) +1 to tnob
        rolling(b) = False
        StopSound("fx_ballrolling" & b)
	Next

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

' Thalamus : Exit in a clean and proper way
Sub table1_exit()
  Controller.Pause = False
  Controller.Stop
End Sub


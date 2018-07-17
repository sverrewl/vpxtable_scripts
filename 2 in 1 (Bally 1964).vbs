'*** 2 in 1 ***

' Thalamus 2018-07-17
' Added "Positional Sound Playback Functions" and "Supporting Ball & Sound Functions"
'

Option Explicit
Randomize

dim score(4)
dim truesc(4)
dim reel(4)
dim ballrelenabled
dim state
dim credit
dim eg
dim currpl
dim playno
dim plno(2)
dim play(2)
dim rst
dim ballinplay
dim match(10)
dim tilt
dim tiltsens
dim rep(4)
dim matchnumb
dim bell
dim points
dim tempscore
dim replay1
dim replay2
dim replay3
dim replay4
dim hisc
dim scn
dim scn1
dim scn2
dim scn3
dim wv
dim mv
dim qs
dim wn
dim sbv1
dim sbv2
dim i

Dim B2SOn
Dim controller
ExecuteGlobal GetTextFile("core.vbs")

Sub Table1_Init
  scntimer.enabled=False
  scntimer1.enabled=False
  scntimer2.enabled=False
  set play(1)=plno1
  set play(2)=plno2
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
  set reel(1)=reel1
  set reel(2)=reel2
  set reel(3)=reel3
  set reel(4)=reel4
  replay1=800
  replay2=1000
  replay3=1200
  replay4=1400
  loadhs
  if hisc="" then hisc=1000
    hisctxt.text=hisc
  if credit="" then credit=0
   credittxt.text=credit
  if matchnumb="" then matchnumb=int(rnd(1)*9)
  select case(matchnumb)
    case 0:
      m0.text="0"
    case 1:
      m1.text="1"
    case 2:
      m2.text="2"
    case 3:
      m3.text="3"
    case 4:
      m4.text="4"
    case 5:
      m5.text="5"
    case 6:
      m6.text="6"
    case 7:
      m7.text="7"
    case 8:
      m8.text="8"
    case 9:
      m9.text="9"
  end select
  for i=1 to 2
    currpl=i
    reel(i).setvalue(score(i))
  next
  currpl=0
  B2SOn=True
  if B2SOn then
    Set Controller = CreateObject("B2S.Server")
    Controller.B2SName = "Bally2In1_1964"
    Controller.Run()
    If Err Then MsgBox "Can't Load B2S.Server."
    End if
    If B2SOn Then
      if matchnumb=0 then
        Controller.B2SSetMatch 10
      else
        Controller.B2SSetMatch matchnumb
    End if
    Controller.B2SSetScoreRolloverPlayer1 0
    Controller.B2SSetScoreRolloverPlayer2 0
    Controller.B2SSetScoreRolloverPlayer3 0
    Controller.B2SSetScoreRolloverPlayer4 0

    Controller.B2SSetTilt 1
    Controller.B2SSetCredits Credit
    Controller.B2SSetGameOver 1
    Controller.B2SSetData 81,0
    Controller.B2SSetData 82,0
    Controller.B2SSetData 83,0
    Controller.B2SSetData 84,0
  End if
  for i=1 to 4
    If B2SOn Then
      Controller.B2SSetScorePlayer i, 0
    End if
  next
  bell=0
  pno1.state=0
  pno2.state=0
  mv=0
  reel3.setvalue(sbv1)
  reel4.setvalue(sbv2)
  pno1.state=1
End Sub

Sub Table1_KeyDown(ByVal keycode)

  If keycode = PlungerKey Then Plunger.PullBack: PlaySound "plungerpull",0,1,0.25,0.25: End if

    if keycode = 6 then
      playsound "coin3"
      coindelay.enabled=true
      If B2SOn Then
        Controller.B2SSetCredits Credit
      End if
    End if

    if keycode = 5 then
      playsound "coin3"
      coindelay1.enabled=true
      If B2SOn Then
        Controller.B2SSetCredits Credit
      End if
    End if

    if keycode = 2 and credit>0 and state=false and playno=0 then
      credit=credit-1
      qs=qs-1
      if qs<0 then qs=0
        credittxt.text=credit
        eg=0
        playno=1
        currpl=1
        pno.setvalue(playno)
        play(currpl).state=1
        playsound "click"
        playsound "initialize"
        rst=0
        ballinplay=1
        resettimer.enabled=true
        If B2SOn Then
          Controller.B2SSetTilt 0
          Controller.B2SSetGameOver 0
          Controller.B2SSetMatch 0
          Controller.B2SSetCredits Credit
          Controller.B2SSetCanPlay 1
          Controller.B2SSetPlayerUp 1
          Controller.B2SSetData 81,0
          Controller.B2SSetData 82,0
          Controller.B2SSetData 83,0
          Controller.B2SSetData 84,0
          'Controller.B2SSetBallInPlay BallInPlay
          Controller.B2SSetScoreRolloverPlayer1 0
        End if
      End if

      if keycode = 2 and credit>0 and state=true and playno>0 and playno<2 and ballinplay<2 then
        credit=credit-1
        qs=qs-1
        if qs<0 then qs=0
          credittxt.text=credit
          playno=playno+1
          select case(playno)
            case 1:
              pno1.state=1
            case 2:
              pno2.state=1
          end select
          pno.setvalue(playno)
          playsound "click"
          If B2SOn Then
            Controller.B2SSetCredits Credit
            Controller.B2SSetCanPlay playno
          End if

        End if

        if state=true and tilt=false then

          If keycode = LeftFlipperKey Then
            LeftFlipper.RotateToEnd
            PlaySound "FlipperUp",0,1,-0.05,0
            playsound "buzz",0 ,1,-0.05,0
          End if

          If keycode = RightFlipperKey Then
            RightFlipper.RotateToEnd
            PlaySound "FlipperUp",0, 1,0.05,0
            playsound "buzz",0, 1,-0.05,0
          End if

          If keycode = LeftTiltKey Then
            Nudge 90, 1
            checktilt
          End if

          If keycode = RightTiltKey Then
            Nudge 270, 1
            checktilt
          End if

          If keycode = CenterTiltKey Then
            Nudge 0, 1
            checktilt
          End if

          If keycode = MechanicalTilt Then
            mechchecktilt
          End if

        End if

        If keycode = RightMagnaSave then Light50.state=1: Light51.state=1: Playsound "click": End if
        If keycode = RightMagnaSave then
          Light1.state=0
          Light2.state=0
          Light3.state=0
          Light4.state=0
          Mushroomlight1.state=0
          Mushroomlight2.state=0
          Mushroomlight3.state=0
          Mushroomlight4.state=0
          Mushroomlight5.state=0
          if B2SOn Then
            Controller.B2SSetData 81,1
          End if
        End if

End Sub

Sub Table1_KeyUp(ByVal keycode)

  If keycode = PlungerKey Then
    Plunger.Fire
    playsound "plungerrelease"
  End if

  If keycode = LeftFlipperKey Then
    LeftFlipper.RotateToStart
    stopsound "buzz"
    if state=true and tilt=false then PlaySound "FlipperDown",0,1,-0.05, 0
  End if

  If keycode = RightFlipperKey Then
    RightFlipper.RotateToStart
    stopsound "buzz"
    if state=true and tilt=false then PlaySound "FlipperDown",0,1,0.05, 0
  End if

End Sub

sub flippertimer_timer()
	LFlip.RotY = LeftFlipper.CurrentAngle+90
	RFlip.RotY = RightFlipper.CurrentAngle+90
End Sub


sub newgame
	Light17.state=0
	Light18.state=0
	Light19.state=0
	Light20.state=0
	Light21.state=0
	Light22.state=0
	Light23.state=0
	Light24.state=0
	Light25.state=0
	Light26.state=0
	Light27.state=0
	Light28.state=0
	Light50.state=0
	Light51.state=0
	sbv1=0
	sbv2=0
	pno2.state=0
	scntimer.enabled=False
	scntimer1.enabled=False
	scntimer2.enabled=False
	state=true
	eg=0
	for i=0 to 3
	score(i)=0
	truesc(i)=0
	rep(i)=0
	next
	bip5.text=" "
	bip1.text="1"
	for i=0 to 9
	match(i).text=" "
	next
	tilttext.text=" "
	gamov.text=" "
	tilt=false
	tiltsens=0
	ballinplay=1
	wheelcheck
	motorcheck
	nb.CreateBall
	nb.kick 90,6
	If B2SOn then Controller.B2SSetBallInPlay 1
End Sub

sub resettimer_timer
    rst=rst+1
    reel1.resettozero
    reel2.resettozero
    reel3.resettozero
    reel4.resettozero
	If B2SOn Then
		Controller.B2SSetScorePlayer1 0
		Controller.B2SSetScorePlayer2 0
		Controller.B2SSetScorePlayer3 0
		Controller.B2SSetScorePlayer4 0
	End if
    if rst=14 then
    playsound "kickerkick"
    End if
    if rst=18 then
    newgame
    resettimer.enabled=false
    End if
End Sub

sub coindelay_timer
	playsound "click"
	credit=credit+5
	credittxt.text=credit
    coindelay.enabled=false
	If B2SOn Then
		Controller.B2SSetCredits Credit
	End if
End Sub

sub coindelay1_timer
	playsound "click"
	credit=credit+1
	credittxt.text=credit
    coindelay1.enabled=false
	If B2SOn Then
		Controller.B2SSetCredits Credit
	End if
End Sub

'*********************
'*** Points Scoring***
'*********************

sub addscore(points)
	if tilt=false then
	bell=0
	scn=0
	scn1=0
	if points = 1 or points = 10 then scn=1
	if points = 1 or points = 10 then matchnumb=matchnumb+1
	if matchnumb=10 then matchnumb=0

	if points=1 or points=10 then
	Light13.state=0
	Light14.state=0
	Light15.state=0
	Light16.state=0
	End if

	if points = 50 then
	reel(currpl).addvalue(50)
	bell=2
	scn=5
	mv=mv+1
	motorcheck
    End if

	if points = 20 then
	reel(currpl).addvalue(20)
	bell=2
	scn=2
	mv=mv+1
	motorcheck
    End if

	if points = 10 then
	reel(currpl).addvalue(10)
	bell=2
    End if

	if points = 1 then
	reel(currpl).addvalue(1)
	bell=1
    End if

	scntimer.enabled=true
	score(currpl)=score(currpl)+points
	truesc(currpl)=truesc(currpl)+points
	If B2SOn Then
		Controller.B2SSetScore currpl,score(currpl)
	End if
	if score(currpl)>9999 then
	score(currpl)=score(currpl)-10000
	rep(currpl)=0
	End if

	if score(currpl)=>replay1 and rep(currpl)=0 then
	credit=credit+1
	qs=qs+2
	playsound "knocke"
	credittxt.text=credit
	If B2SOn Then
		Controller.B2SSetCredits Credit
	End if
	rep(currpl)=1
	playsound "click"
	End if

	if score(currpl)=>replay2 and rep(currpl)=1 then
	qs=qs+2
	credit=credit+1
	playsound "knocke"
	credittxt.text=credit
	If B2SOn Then
		Controller.B2SSetCredits Credit
	End if
	rep(currpl)=2
	playsound "click"
	End if

	if score(currpl)=>replay3 and rep(currpl)=2 then
	credit=credit+1
	qs=qs+2
	playsound "knocke"
	credittxt.text=credit
	If B2SOn Then
		Controller.B2SSetCredits Credit
	End if
	rep(currpl)=3
	playsound "click"
	End if

	if score(currpl)=>replay4 and rep(currpl)=3 then
	credit=credit+1
	qs=qs+2
	playsound "knocke"
	credittxt.text=credit
	If B2SOn Then
		Controller.B2SSetCredits Credit
	End if
	rep(currpl)=4
	playsound "click"
	End if
    End if
End Sub


sub scntimer_timer
    scn1=scn1 + 1
    if bell=1 then playsound "bell1",0, 0.3, 0, 0
    if bell=2 then playsound "bell10",0, 0.6, 0, 0
    if bell=3 then playsound "bell100",0, 0.9, 0, 0
    wv=wv+1
    if wv>9 then wv=0
    wheelcheck
    if scn1=scn then scntimer.enabled=false
End Sub

sub scntimer1_timer
    scn2=scn2 + 1
    if bell=3 then playsound "bell100",0, 0.9, 0, 0
    if scn2=scn3 then scntimer1.enabled=false
    wheelcheck
End Sub


sub wheelcheck
	If wv=0 or wv=1 then
	BumperLight1.state=1
	LeftSlingLight.state=1
	BumperLight2.state=0
	BumperLight3.state=0
	RightSlingLight.state=0
	If Light50.state=1 then Mushroomlight1.state=0 else MushroomLight1.state=1
	MushroomLight2.state=0
	MushroomLight3.state=0
	If Light50.state=1 then Mushroomlight4.state=0 else	MushroomLight4.state=1
	If Light50.state=1 then Mushroomlight5.state=0 else	MushroomLight5.state=1
	If Light50.state=1 then Light1.state=0 else Light1.state=1
	Light2.state=0
	Light3.state=0
	If Light50.state=1 then Light4.state=0 else	Light4.state=1
    Light8.state=0
    Light9.state=0
    Light10.state=0
    Light11.state=1
    Light12.state=1
    quickstep
    End if

    If wv=2 or wv=3 then
	BumperLight1.state=1
	LeftSlingLight.state=1
	BumperLight2.state=1
	bumperLight3.state=0
	RightSlingLight.state=0
	MushroomLight1.state=0
	If Light50.state=1 then Mushroomlight2.state=0 else	MushroomLight2.state=1
	If Light50.state=1 then Mushroomlight3.state=0 else MushroomLight3.state=1
	MushroomLight4.state=0
	MushroomLight5.state=0
	Light1.state=0
	If Light50.state=1 then Light2.state=0 else Light2.state=1
	If Light50.state=1 then Light2.state=0 else Light3.state=1
	Light4.state=0
    Light8.state=1
    Light9.state=1
    Light10.state=1
    Light11.state=0
    Light12.state=0
    quickstep
    End if

    If wv=4 or wv=5 then
	BumperLight1.state=1
	LeftSlingLight.state=1
	BumperLight2.state=1
	BumperLight3.state=1
	RightSlingLight.state=1
	If Light50.state=1 then Mushroomlight1.state=0 else Mushroomlight1.state=1
	MushroomLight2.state=0
	MushroomLight3.state=0
	If Light50.state=1 then Mushroomlight4.state=0 else MushroomLight4.state=1
	If Light50.state=1 then Mushroomlight5.state=0 else MushroomLight5.state=1
	If Light50.state=1 then Light1.state=0 else Light1.state=1
	Light2.state=0
	Light3.state=0
	If Light50.state=1 then Light4.state=0 else Light4.state=1
    Light8.state=0
    Light9.state=0
    Light10.state=0
    Light11.state=1
    Light12.state=1
    quickstep
    End if

    If wv=6 or wv=7 then
	BumperLight1.state=0
	LeftSlingLight.state=0
	BumperLight2.state=1
	BumperLight3.state=1
	RightSlingLight.state=1
	MushroomLight1.state=0
	If Light50.state=1 then Mushroomlight2.state=0 else MushroomLight2.state=1
	If Light50.state=1 then Mushroomlight3.state=0 else MushroomLight3.state=1
	MushroomLight4.state=0
	MushroomLight5.state=0
	Light1.state=0
	If Light50.state=1 then Light2.state=0 else Light2.state=1
	If Light50.state=1 then Light3.state=0 else Light3.state=1
	Light4.state=0
    Light8.state=1
    Light9.state=1
    Light10.state=1
    Light11.state=0
    Light12.state=0
    quickstep
    End if

    If wv=8 or wv=9 then
	BumperLight1.state=0
	LeftSlingLight.state=0
	BumperLight2.state=0
	BumperLight3.state=1
	RightSlingLight.state=1
	If Light50.state=1 then Mushroomlight1.state=0 else MushroomLight1.state=1
	MushroomLight2.state=0
	MushroomLight3.state=0
	If Light50.state=1 then Mushroomlight4.state=0 else MushroomLight4.state=1
	If Light50.state=1 then Mushroomlight5.state=0 else MushroomLight5.state=1
	If Light50.state=1 then Light1.state=0 else Light1.state=1
	Light2.state=0
	Light3.state=0
	If Light50.state=1 then Light4.state=0 else Light4.state=1
    Light8.state=0
    Light9.state=0
    Light10.state=0
    Light11.state=1
    Light12.state=1
    quickstep
    End if
End Sub


Sub quickstep

    If wv=1 and qs>=8 then
    BumperLight1.state=0
    LeftSlingLight.state=0
    BumperLight2.state=0
    BumperLight3.state=0
    RightSlingLight.state=0
    End if

    If wv=3 and qs>=12 then
    BumperLight1.state=0
    LeftSlingLight.state=0
    BumperLight2.state=0
    BumperLight3.state=0
    RightSlingLight.state=0
    End if

    If (wv=4 or wv=5) and qs>=16 then
    BumperLight1.state=0
    LeftSlingLight.state=0
    BumperLight2.state=0
    BumperLight3.state=0
    RightSlingLight.state=0
    End if

    If wv=7 and qs>=20 then
    BumperLight1.state=0
    LeftSlingLight.state=0
    BumperLight2.state=0
    BumperLight3.state=0
    RightSlingLight.state=0
    End if

    If wv=9 and qs>=24 then
    BumperLight1.state=0
    LeftSlingLight.state=0
    BumperLight2.state=0
    BumperLight3.state=0
    RightSlingLight.state=0
    End if

End Sub


Sub motorcheck

	If mv>2 then mv=0

	If mv=0 then
	Light5.state=1
	Light6.state=1
	Light7.state=0
    End if

	If mv=1 then
	Light5.state=1
	Light6.state=0
	Light7.state=1
    End if

	If mv=2 then
	Light5.state=0
	Light6.state=1
	Light7.state=1
	End if

End Sub


Sub Lightcheck
	If sbv1>21 then Light22.State=0: End if
	If sbv1=21 then Light22.state=1: Light21.state=0: Light20.state=0: Light19.state=0: End if
	If sbv1=20 then Light21.state=1: Light20.state=0: Light19.state=0: Light18.state=0: End if
	If sbv1=19 then Light20.state=1: Light19.state=0: Light18.state=0: Light17.state=0: End if
	If sbv1=18 then Light19.state=1: Light18.state=0: Light17.state=0: End if
	If sbv1=17 then Light18.state=1: Light17.state=0: End if
	If sbv1=16 then Light17.state=1: End if

	If sbv2>21 then Light28.State=0: End if
	If sbv2=21 then Light28.state=1: Light27.state=0: Light26.state=0: Light25.state=0: End if
	If sbv2=20 then Light27.state=1: Light26.state=0: Light25.state=0: Light24.state=0: End if
	If sbv2=19 then Light26.state=1: Light25.state=0: Light24.state=0: Light23.state=0: End if
	If sbv2=18 then Light25.state=1: Light24.state=0: Light23.state=0: End if
	If sbv2=17 then Light24.state=1: Light23.state=0: End if
	If sbv2=16 then Light23.state=1: End if

End Sub


sub wheelchange_timer
    wn=wn+1
    if tilt=false then
    If points=1 or points=10 then wv=wv+1
    if wv>9 then wv=0
    wheelcheck
    End if
End Sub

Sub CheckTilt
	If Tilttimer.Enabled = True Then
	TiltSens = TiltSens + 1
	if TiltSens = 2 Then
	Tilt = True
	tilttext.text="TILT"
	If B2SOn Then Controller.B2SSetTilt 1
	tiltsens = 0
	playsound "tilt"
	turnoff
	End if
	Else
	TiltSens = 0
	Tilttimer.Enabled = True
	End if
End Sub

Sub MechCheckTilt
	Tilt = True
	tilttext.text="TILT"
	If B2SOn Then Controller.B2SSetTilt 1
	tiltsens = 0
	playsound "tilt"
	LeftFlipper.RotateToStart
	RightFlipper.RotateToStart
	turnoff
End Sub

Sub Tilttimer_Timer()
	Tilttimer.Enabled = False
End Sub

sub turnoff
    tiltseq.play seqalloff
    for i=0 to 10
    next
End Sub

Sub Drain_Hit()
	playsound "drainshorter"
	Light50.state=0
	Light51.state=0
	wheelcheck
	Drain.DestroyBall
	nextball
End Sub

sub nextball
	if tilt=true then
	tilt=false
	tilttext.text=" "
	If B2SOn Then Controller.B2SSetTilt 0:Controller.B2SSetData 81,0
	tiltseq.stopplay
	pno.setvalue(playno)
	End if
	If B2SOn Then Controller.B2SSetData 81,0
	ballreltimer.enabled=true
	scntimer.enabled=False
	scntimer1.enabled=False
	scntimer2.enabled=False
	currpl=currpl+1
	if B2SOn then Controller.B2SSetPlayerUp currpl
	if currpl>playno then
	ballinplay=ballinplay+1
	if ballinplay>5 then
	playsound "motorleer"
	If B2SOn then Controller.B2SSetBallInPlay 0
	if B2SOn then Controller.B2SSetPlayerUp 0
	eg=1
	if sbv1=21 then credit=credit+1: playsound "knocke": qs=qs+2: End if
	if sbv2=21 then credit=credit+1: playsound "knocke": qs=qs+2: End if
	credittxt.text=credit
	If B2SOn Then
		Controller.B2SSetCredits Credit
	End if
	ballreltimer.enabled=true
	else
	if state=true and tilt=false then
	play(currpl-1).state=0
	currpl=1
	play(currpl).state=1
	if B2SOn then Controller.B2SSetPlayerUp currpl
	playsound "kickerkick"
	ballreltimer.enabled=true
	End if
	If B2SOn then Controller.B2SSetBallInPlay ballinplay
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
	End if
	End if
	if currpl>1 and currpl<(playno+1) then
	if state=true and tilt=false then
	play(currpl-1).state=0
	play(currpl).state=1
	if B2SOn then Controller.B2SSetPlayerUp currpl
	playsound "kickerkick"
	ballreltimer.enabled=true
	wheelcheck
	lightcheck
	End if
	End if
End Sub

sub ballreltimer_timer
	if eg=1 then
	matchnum
	bip3.text=" "
	bip5.text=" "
	state=false
	for i=1 to 2
	if truesc(i)>hisc then
	hisc=truesc(i)
	hisctxt.text=hisc
	End if
	next
	pno.setvalue(0)
	play(currpl-1).state=0
	playno=0
	gamov.text="GAME OVER"
	If B2SOn Then
		Controller.B2SSetGameOver 1
	End if

	savehs
	ballreltimer.enabled=false
	else
	nb.CreateBall
	nb.kick 90,6
    ballreltimer.enabled=false
    End if
End Sub

sub matchnum
    select case(matchnumb)
    case 0:
    m0.text="0"
    case 1:
    m1.text="1"
    case 2:
    m2.text="2"
    case 3:
    m3.text="3"
    case 4:
    m4.text="4"
    case 5:
    m5.text="5"
    case 6:
    m6.text="6"
    case 7:
    m7.text="7"
    case 8:
    m8.text="8"
    case 9:
    m9.text="9"
    end select
	If B2SOn Then

		if matchnumb=0 then
			Controller.B2SSetMatch 10
		else
			Controller.B2SSetMatch matchnumb
		End if
	End if
    for i=1 to playno
    if matchnumb=(score(i) mod 10) then
    credit=credit+1
	If B2SOn Then
		Controller.B2SSetCredits Credit
	End if
    qs=qs+2
    playsound "knocke"
    credittxt.text= credit
    End if
    next
End Sub



'**********Sling Shot Animations
' Rstep and Lstep  are the variables that increment the animation
'****************
Dim RStep, Lstep, RStep1, LStep1

Sub RightSlingShot_Slingshot
	PlaySound "right_slingshot", 0, 1, 0.05, 0.05
	RSling.Visible = 0
	RSling1.Visible = 1
	sling1.TransZ = -20
	RStep = 0
	RightSlingShot.TimerEnabled = 1
	If RightSlingLight.state=1 then Addscore 10 else Addscore 1
End Sub

Sub RightSlingShot_Timer
    Select Case RStep
    Case 3:RSLing1.Visible = 0:RSLing2.Visible = 1:sling1.TransZ = -10
    Case 4:RSLing2.Visible = 0:RSLing.Visible = 1:sling1.TransZ = 0:RightSlingShot.TimerEnabled = 0
    End Select
    RStep = RStep + 1
End Sub


Sub LeftSlingShot_Slingshot
	PlaySound "left_slingshot",0,1,-0.05,0.05
	LSling.Visible = 0
	LSling1.Visible = 1
	sling2.TransZ = -20
	LStep = 0
	LeftSlingShot.TimerEnabled = 1
	If LeftSlingLight.state=1 then Addscore 10 else Addscore 1
End Sub

Sub LeftSlingShot_Timer
    Select Case LStep
    Case 3:LSLing1.Visible = 0:LSLing2.Visible = 1:sling2.TransZ = -10
    Case 4:LSLing2.Visible = 0:LSLing.Visible = 1:sling2.TransZ = 0:LeftSlingShot.TimerEnabled = 0
    End Select
    LStep = LStep + 1
End Sub

Sub LeftSlingShot1_Slingshot
    addscore 1
    PlaySound "left_slingshot",0,1,-0.05,0.05
    LSling4.Visible = 0
    LSling5.Visible = 1
    sling4.TransZ = -20
    LStep1 = 0
    LeftSlingShot1.TimerEnabled = 1
End Sub

Sub LeftSlingShot1_Timer
    Select Case LStep1
    Case 3:LSLing4.Visible = 0:LSLing5.Visible = 1:sling4.TransZ = -10
    Case 4:LSLing5.Visible = 0:LSLing3.Visible = 1:sling4.TransZ = 0:LeftSlingShot1.TimerEnabled = 0
    End Select
    LStep1 = LStep1 + 1
End Sub


' *********************************************************************
'                      Supporting Ball & Sound Functions
' *********************************************************************

Function Vol(ball) ' Calculates the Volume of the sound based on the ball speed
    Vol = Csng(BallVel(ball) ^2 / 2000)
End Function

Function Pan(ball) ' Calculates the pan for a ball based on the X position on the table. "table1" is the name of the table
    Dim tmp
    tmp = ball.x * 2 / table1.width-1
    If tmp > 0 Then
        Pan = Csng(tmp ^10)
    Else
        Pan = Csng(-((- tmp) ^10) )
    End if
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

Const tnob = 10 ' total number of balls
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
            PlaySound("fx_ballrolling" & b), -1, Vol(BOT(b)), Pan(BOT(b) ), 0, Pitch(BOT(b) ), 1, 0, AudioFade(BOT(b))
        Else
            If rolling(b) = True Then
                StopSound("fx_ballrolling" & b)
                rolling(b) = False
            End if
        End if
    Next
End Sub

'**********************
' Ball Collision Sound
'**********************

Sub OnBallBallCollision(ball1, ball2, velocity)
	PlaySound("fx_collide"), 0, Csng(velocity) ^2 / 2000, Pan(ball1), 0, Pitch(ball1), 0, 0
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

Sub a_Pins_Hit (idx)
	PlaySound "pinhit_low", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0
End Sub

Sub a_Targets_Hit (idx)
	PlaySound "target", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0
End Sub

Sub a_Metals_Thin_Hit (idx)
	PlaySound "metalhit_thin", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0
End Sub

Sub a_Metals_Medium_Hit (idx)
	PlaySound "metalhit_medium", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0
End Sub

Sub a_Metals2_Hit (idx)
	PlaySound "metalhit2", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0
End Sub

Sub a_Gates_Hit (idx)
	PlaySound "gate4", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0
End Sub

Sub a_Spinner_Spin
	PlaySound "fx_spinner",0,.25,0,0.25
End Sub

Sub a_Rubbers_Hit(idx)
 	dim finalspeed
  	finalspeed=SQR(activeball.velx * activeball.velx + activeball.vely * activeball.vely)
 	If finalspeed > 20 then
		PlaySound "fx_rubber2", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0
	End if
	If finalspeed >= 6 AND finalspeed <= 20 then
 		RandomSoundRubber()
 	End if
End Sub

Sub RubberWheel_hit
 	dim finalspeed
  	finalspeed=SQR(activeball.velx * activeball.velx + activeball.vely * activeball.vely)
 	If finalspeed > 20 then
		PlaySound "fx_rubber2", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0
	End if
	If finalspeed >= 6 AND finalspeed <= 20 then
 		RandomSoundRubber()
 	End if
End Sub

Sub a_Posts_Hit(idx)
 	dim finalspeed
  	finalspeed=SQR(activeball.velx * activeball.velx + activeball.vely * activeball.vely)
 	If finalspeed > 16 then
		PlaySound "fx_rubber2", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0
	End if
	If finalspeed >= 6 AND finalspeed <= 16 then
 		RandomSoundRubber()
 	End if
End Sub

Sub RandomSoundRubber()
	Select Case Int(Rnd*3)+1
		Case 1 : PlaySound "rubber_hit_1", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0
		Case 2 : PlaySound "rubber_hit_2", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0
		Case 3 : PlaySound "rubber_hit_3", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0
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
		Case 1 : PlaySound "flip_hit_1", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0
		Case 2 : PlaySound "flip_hit_2", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0
		Case 3 : PlaySound "flip_hit_3", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0
	End Select
End Sub

sub savehs
	' Based on Black's Highscore routines
	Dim FileObj
	Dim ScoreFile
	Set FileObj=CreateObject("Scripting.FileSystemObject")
	If Not FileObj.FolderExists(UserDirectory) then
	Exit Sub
	End if
	Set ScoreFile=FileObj.CreateTextFile(UserDirectory & "2in1.txt",True)
	scorefile.writeline credit
	ScoreFile.WriteLine score(0)
	ScoreFile.WriteLine score(1)
	scorefile.writeline hisc
	Scorefile.writeline matchnumb
	scorefile.writeline wv
	scorefile.writeline mv
	scorefile.writeline qs
	ScoreFile.Close
	Set ScoreFile=Nothing
	Set FileObj=Nothing
End Sub

sub loadhs
    ' Based on Black's Highscore routines
	Dim FileObj
	Dim ScoreFile
	dim TextStr
	dim temp1
	dim temp2
	dim temp3
	dim temp4
	dim temp5
	dim temp6
	dim temp7
	dim temp8
	Set FileObj=CreateObject("Scripting.FileSystemObject")
	If Not FileObj.FolderExists(UserDirectory) then
	Exit Sub
	End if
	If Not FileObj.FileExists(UserDirectory & "2in1.txt") then
	Exit Sub
	End if
	Set ScoreFile=FileObj.GetFile(UserDirectory & "2in1.txt")
	Set TextStr=ScoreFile.OpenAsTextStream(1,0)
	If (TextStr.AtEndOfStream=True) then
	Exit Sub
	End if
	temp1=TextStr.ReadLine
	temp2=textstr.readline
	temp3=textstr.readline
	temp4=Textstr.ReadLine
	temp5=Textstr.ReadLine
	temp6=Textstr.ReadLine
	temp7=Textstr.ReadLine
	temp8=Textstr.ReadLine
	TextStr.Close
	Credit = CDbl(temp1)
	score(0) = CDbl(temp2)
	score(1) = CDbl(temp3)
	hisc = CDbl(temp4)
	matchnumb = CDbl(temp5)
    wv = CDbl(temp6)
    mv = CDbl(temp7)
    qs = CDbl(temp8)
    Set ScoreFile=Nothing
    Set FileObj=Nothing
End Sub

sub bumper1_hit
    if tilt=false then playsound "jet1"
    if BumperLight1.state=1 then addscore 10 else addscore 1
End Sub

sub bumper2_hit
    if tilt=false then playsound "jet1"
    if BumperLight2.state=1 then addscore 10 else addscore 1
End Sub

sub bumper3_hit
    if tilt=false then playsound "jet1"
    if BumperLight3.state=1 then addscore 10 else addscore 1
End Sub


Sub Trigger1_Hit()
	If Light1.state=1 and plno1.state=1 then sbv1=sbv1+1: reel3.addvalue(1): End if
	If Light1.state=1 and plno2.state=1 then sbv2=sbv2+1: reel4.addvalue(1): End if
	If B2SOn Then
		if currpl=1 then
			Controller.B2SSetScore 3,sbv1
		Else
			Controller.B2SSetScore 4,sbv2
		End if
	End if

	Lightcheck
	If Light13.State=1 then addscore 20 else addscore 10
End Sub

Sub Trigger2_Hit()
	If Light2.state=1 and plno1.state=1 then sbv1=sbv1+2: reel3.addvalue(2): End if
	If Light2.state=1 and plno2.state=1 then sbv2=sbv2+2: reel4.addvalue(2): End if
	If B2SOn Then
		if currpl=1 then
			Controller.B2SSetScore 3,sbv1
		Else
			Controller.B2SSetScore 4,sbv2
		End if
	End if
	Lightcheck
	If Light14.State=1 then addscore 50 else addscore 10
End Sub

Sub Trigger3_Hit()
	If Light3.state=1 and plno1.state=1 then sbv1=sbv1+3: reel3.addvalue(3): End if
	If Light3.state=1 and plno2.state=1 then sbv2=sbv2+3: reel4.addvalue(3): End if
	If B2SOn Then
		if currpl=1 then
			Controller.B2SSetScore 3,sbv1
		Else
			Controller.B2SSetScore 4,sbv2
		End if
	End if
	Lightcheck
	If Light15.State=1 then addscore 50 else addscore 10
End Sub

Sub Trigger4_Hit()
	If Light4.state=1 and plno1.state=1 then sbv1=sbv1+1: reel3.addvalue(1): End if
	If Light4.state=1 and plno2.state=1 then sbv2=sbv2+1: reel4.addvalue(1): End if
	If B2SOn Then
		if currpl=1 then
			Controller.B2SSetScore 3,sbv1
		Else
			Controller.B2SSetScore 4,sbv2
		End if
	End if
	Lightcheck
	If Light16.State=1 then addscore 20 else addscore 10
End Sub

Sub Trigger5_Hit()
	if tilt = false then
		If Light5.State=1 then addscore 50 else addscore 10
	End if
End Sub

Sub Trigger6_Hit()
	If tilt = false then
		If Light50.state=1 then Addscore 10 else addscore 1
	End if
End Sub

Sub Trigger7_Hit()
	If tilt = false then
		If Light51.state=1 then Addscore 10 else addscore 1
	End if
End Sub

Sub Trigger8_Hit()
	If Light6.State=1 then addscore 50 else addscore 10
End Sub

Sub Trigger9_Hit()
	If Light7.State=1 then addscore 50 else addscore 10
End Sub

Sub Bumper4_Hit()
	if tilt = false then
	If MushroomLight1.state=1 and plno1.state=1 then sbv1=sbv1+1: reel3.addvalue(1): End if
	If MushroomLight1.state=1 and plno2.state=1 then sbv2=sbv2+1: reel4.addvalue(1): End if
	If B2SOn Then
		if currpl=1 then
			Controller.B2SSetScore 3,sbv1
		Else
			Controller.B2SSetScore 4,sbv2
		End if
	End if
	Lightcheck
	If Light8.State=1 then addscore 10 else addscore 1
	End if
End Sub

Sub Bumper5_Hit()
	if tilt = false then
	If MushroomLight5.state=1 and plno1.state=1 then sbv1=sbv1+1: reel3.addvalue(1): End if
	If MushroomLight5.state=1 and plno2.state=1 then sbv2=sbv2+1: reel4.addvalue(1): End if
	If B2SOn Then
		if currpl=1 then
			Controller.B2SSetScore 3,sbv1
		Else
			Controller.B2SSetScore 4,sbv2
		End if
	End if
	Lightcheck
	If Light9.State=1 then addscore 10 else addscore 1
	End if
End Sub

Sub Bumper6_Hit()
	If tilt = false then
	If MushroomLight4.state=1 and plno1.state=1 then sbv1=sbv1+1: reel3.addvalue(1): End if
	If MushroomLight4.state=1 and plno2.state=1 then sbv2=sbv2+1: reel4.addvalue(1): End if
	If B2SOn Then
		if currpl=1 then
			Controller.B2SSetScore 3,sbv1
		Else
			Controller.B2SSetScore 4,sbv2
		End if
	End if
	Lightcheck
	If Light10.State=1 then addscore 10 else addscore 1
	End if
End Sub

Sub Bumper7_Hit()
	if tilt = false then
	If MushroomLight2.state=1 and plno1.state=1 then sbv1=sbv1+2: reel3.addvalue(2): End if
	If MushroomLight2.state=1 and plno2.state=1 then sbv2=sbv2+2: reel4.addvalue(2): End if
	If B2SOn Then
		if currpl=1 then
			Controller.B2SSetScore 3,sbv1
		Else
			Controller.B2SSetScore 4,sbv2
		End if
	End if
	Lightcheck
	If Light11.State=1 then addscore 20 else addscore 1
	End if
End Sub

Sub Bumper8_Hit()
	if tilt = false then
	If MushroomLight3.state=1 and plno1.state=1 then sbv1=sbv1+2: reel3.addvalue(2): End if
	If MushroomLight3.state=1 and plno2.state=1 then sbv2=sbv2+2: reel4.addvalue(2): End if
	If B2SOn Then
		if currpl=1 then
			Controller.B2SSetScore 3,sbv1
		Else
			Controller.B2SSetScore 4,sbv2
		End if
	End if
	Lightcheck
	If Light12.State=1 then addscore 20 else addscore 1
	End if
End Sub


Sub LeftSlingShot2_Slingshot()
	Addscore 1
End Sub

Sub LeftSlingShot3_Slingshot()
	Addscore 1
End Sub

Sub LeftSlingShot4_Slingshot()

End Sub

Sub RightSlingShot1_Slingshot()
    addscore 1
End Sub

Sub RightSlingShot2_Slingshot()
	Addscore 1
End Sub

Sub RightSlingShot3_Slingshot()
	Addscore 1
End Sub

Sub RightSlingShot4_Slingshot()
	Light13.state=1
	Light14.state=1
	Light15.state=1
	Light16.state=1
	Light1.state=0
	Light2.state=0
	Light3.state=0
	Light4.state=0
	Playsound "click"
End Sub

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

Function AudioFade(tableobj) ' Fades between front and back of the table (for surround systems or 2x2 speakers, etc), depending on the Y position on the table. "table1" is the name of the table
  Dim tmp
  tmp = tableobj.y * 2 / table1.height-1
  If tmp > 0 Then
    AudioFade = Csng(tmp ^10)
  Else
    AudioFade = Csng(-((- tmp) ^10) )
  End if
End Function

Function AudioPan(tableobj) ' Calculates the pan for a tableobj based on the X position on the table. "table1" is the name of the table
  Dim tmp
  tmp = tableobj.x * 2 / table1.width-1
  If tmp > 0 Then
    AudioPan = Csng(tmp ^10)
  Else
    AudioPan = Csng(-((- tmp) ^10) )
  End if
End Function

' Function Vol(ball) ' Calculates the Volume of the sound based on the ball speed
'   Vol = Csng(BallVel(ball) ^2 / 2000)
' End Function
'
' Function Pitch(ball) ' Calculates the pitch of the sound based on the ball speed
'   Pitch = BallVel(ball) * 20
' End Function
'
' Function BallVel(ball) 'Calculates the ball speed
'   BallVel = INT(SQR((ball.VelX ^2) + (ball.VelY ^2) ) )
' End Function


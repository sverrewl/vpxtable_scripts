'Option Explicit
'Randomize

' Thalamus 2018-07-18
' Added/Updated "Positional Sound Playback Functions" and "Supporting Ball & Sound Functions"
' Added , AudioFade(ActiveBall) to most of the routines to test if it is done right.

dim ballinplay
dim ballrelenabled
dim rst
dim eg
dim credit
dim score
dim truesc
dim match(9)
dim rep
dim tilt
dim tiltsens
dim state
dim cred
dim plm
dim update
dim digit
dim tempscore
dim scn
dim scn1
dim bell
dim points
dim matchnumb
dim replay1
dim replay2
dim replay3
dim replay4
dim hisc
dim rs(3)
dim rv(3)
dim sv
dim sv1
dim i
dim TextStr
dim reel(2)
dim hv
dim wv
dim qs

Dim B2SOn
Dim controller
ExecuteGlobal GetTextFile("core.vbs")

sub Table1_init
	If ShowDT=false Then
		for each obj in DTItems
			obj.visible=False
		Next
	end If

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
    replay1=43000
    replay2=54000
    replay3=65000
    replay4=76000
    loadhs
    if hisc="" then hisc=10000
    reel1.setvalue(score)
    reel2.setvalue(hisc)
    if credit="" then credit=0
    credittxt.text=credit
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
	B2SOn=True
	if B2SOn then
		Set Controller = CreateObject("B2S.Server")
		Controller.B2SName = "Bally4Queens_1970"
		Controller.Run()
		If Err Then MsgBox "Can't Load B2S.Server."
	end if
	If B2SOn Then

		if matchnumb=0 then
			Controller.B2SSetMatch 10
		else
			Controller.B2SSetMatch matchnumb
		end if
		Controller.B2SSetScoreRolloverPlayer1 0


		Controller.B2SSetTilt 1
		Controller.B2SSetCredits Credit
		Controller.B2SSetGameOver 1
		Controller.B2SSetData 81,0
		Controller.B2SSetData 82,0
		Controller.B2SSetData 83,0
		Controller.B2SSetData 84,0
	End If
	for i=1 to 4
		If B2SOn Then
			Controller.B2SSetScorePlayer i, 0
		End If
	next
End Sub


Sub Table1_KeyDown(ByVal keycode)
	if keycode=6 then
	playsound "coin3"
	coindelay.enabled=true
	end if
	if keycode=2 and credit>0 and state=false then
	credit=credit-1
	If B2SOn Then
		Controller.B2SSetCredits Credit
	end if
	qs=qs-1
	if qs<0 then qs=0
	playsound "click"
	credittxt.text=credit
	tilt=false
	state=true
	rst=0
	ballinplay=1
	playsound "initialize"
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
			Controller.B2SSetBallInPlay BallInPlay
			Controller.B2SSetScoreRolloverPlayer1 0

		End If
	end if
	If keycode = PlungerKey Then
	Plunger.PullBack
	End If
	if tilt=false and state=true then
	If keycode = LeftFlipperKey Then
	LeftFlipper.RotateToEnd
	LeftFlipper1.RotateToEnd
	playsound "buzz"
	PlaySound "FlipperUp"
	End If
	If keycode = RightFlipperKey Then
	RightFlipper.RotateToEnd
	RightFlipper1.RotateToEnd
	playsound "buzz"
	PlaySound "FlipperUp"
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

Sub Table1_KeyUp(ByVal keycode)
	If keycode = PlungerKey Then
	playsound "plunger"
	Plunger.Fire
	End If
	If keycode = LeftFlipperKey Then
	LeftFlipper.RotateToStart
	LeftFlipper1.RotateToStart
	stopsound "buzz"
	if tilt=false and state=true then PlaySound "FlipperDown"
	End If
	If keycode = RightFlipperKey Then
	RightFlipper.RotateToStart
	RightFlipper1.RotateToStart
	stopsound "buzz"
	if tilt=false and state=true then PlaySound "FlipperDown"
	End If
End Sub

sub coindelay_timer
	playsound "click"
	credit=credit+5
	if credit>25 then credit=25
	credittxt.text=credit
	coindelay.enabled=false
	If B2SOn Then
		Controller.B2SSetCredits Credit
	end if
End Sub

sub resettimer_timer
    rst=rst+1
    reel1.resettozero
	If B2SOn Then
		Controller.B2SSetScorePlayer1 0
	end if
    if rst=14 then
    playsound "kickerkick"
    end if
    if rst=18 then
    newgame
    resettimer.enabled=false
    end if
End Sub

sub newgame
	bumper1.force=12
	bumper2.force=12
	bumper3.force=12
	score=0
	truesc=0
	eg=0
	rep=0
	sv=0
	bumperlight1.state=lightstateoff
	bumperlight2.state=0
	bumperlight3.state=0
	Light7.State=0
	Light8.State=0
	Light9.State=0
	Light10.State=0
	Light11.State=0
	LightA.State=2
	LightB.State=0
	LightC.State=0
	LightD.State=0
	gamov.text=" "
	tilttxt.text=" "
	bip5.text=" "
	bip1.text="1"
	for i=0 to 9
	match(i).text=" "
	next
	resettimer.enabled=Trus
	nb.CreateBall
	nb.kick 135,6
	If B2SOn then Controller.B2SSetBallInPlay 1
End Sub

sub newball
    bumperlight1.state=0
    bumperlight2.state=0
    bumperlight3.state=0
End Sub

Sub Drain_Hit()
	PlaySoundAt "drainshorter", Drain
	Drain.DestroyBall
	flipopen
	nextball
End Sub

sub nextball
    if tilt=true then
    bumper1.force=12
    bumper2.force=12
    bumper3.force=12
    tilt=false
	If B2SOn Then Controller.B2SSetTilt 0
    tilttxt.text=" "
    end if

    LightA.State=2
    LightB.State=0
    LightC.State=0
    LightD.State=0

    hv=0
    Light1.State=0
    Light2.State=0
    Light3.State=0
    Light4.State=0
    Light5.State=0

	ballinplay=ballinplay+1
	if ballinplay>5 then
	playsound "motorleer"
	If B2SOn then Controller.B2SSetBallInPlay 0
	eg=1
	ballreltimer.enabled=true
	else
	if state=true and tilt=false then
	newball
	playsound "kickerkick"
	ballreltimer.enabled=true
	end if
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
	end if
End Sub

sub ballreltimer_timer
	if eg=1 then
	matchnum
	bip5.text=" "
	state=false
	gamov.text="GAME OVER"
	If B2SOn Then
		Controller.B2SSetGameOver 1
	end if
	if truesc>hisc then hisc=truesc
	reel2.setvalue(hisc)
	savehs
	ballreltimer.enabled=false
	else
	nb.CreateBall
	nb.kick 90,4
    ballreltimer.enabled=false
    end if
End Sub

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
	If B2SOn Then

		if matchnumb=0 then
			Controller.B2SSetMatch 100
		else
			Controller.B2SSetMatch (matchnumb*10)
		end if
	end if
	if (matchnumb*10)=(score mod 100) then
	credit=credit+1
    qs=qs+2
	playsound "knocke"
	if credit>25 then credit=25
	If B2SOn Then
		Controller.B2SSetCredits Credit
	end if
	credittxt.text=credit
    playsound "click"
    end if
End Sub

sub addscore(points)

    if tilt=false then
    bell=0
    scn=0
    if points = 10 or points = 100 or points=1000 then scn=1

    if points = 5000 then
    reel(1).addvalue(5000)
    bell=1
    scn=5
    end if

    if points = 4000 then
    reel(1).addvalue(4000)
    bell=1
    scn=4
    end if

    if points = 3000 then
    reel(1).addvalue(3000)
    bell=1
    scn=3
    end if

    if points = 2000 then
    reel(1).addvalue(2000)
    bell=1
    scn=2
    end if

    if points = 1000 then
    reel(1).addvalue(1000)
    bell=1
    end if

    if points = 500 then
    reel(1).addvalue(500)
    bell=2
    scn=5
    end if

    if points = 300 then
    reel(1).addvalue(300)
    bell=2
    scn=3
    end if

    if points = 100 then
    reel(1).addvalue(100)
    bell=2
    end if

    if points = 10 then
    reel(1).addvalue(10)
    bell=3
    end if

    scn1=0
    scntimer.enabled=true
    score=score+points
    truesc=truesc+points
	If B2SOn Then
		Controller.B2SSetScore 1,score
	end if
    if score=>100000 then
    score=score-100000
    rep=0
    end if

    if score=>replay1 and rep=0 then
    credit=credit+1
    qs=qs+2
    playsound "knocke"

	if credit>25 then credit=25
	credittxt.text=credit
	If B2SOn Then
		Controller.B2SSetCredits Credit
	end if
    rep=1
    playsound "click"
    end if

    if score=>replay2 and rep=1 then
    credit=credit+1
    qs=qs+2
    playsound "knocke"

	if credit>25 then credit=25
	credittxt.text=credit
	If B2SOn Then
		Controller.B2SSetCredits Credit
	end if
    rep=2
    playsound "click"
    end if

    if score=>replay3 and rep=2 then
    credit=credit+1
    qs=qs+2
    playsound "knocke"

	if credit>25 then credit=25
	credittxt.text=credit
	If B2SOn Then
		Controller.B2SSetCredits Credit
	end if
    rep=3
    playsound "click"
    end if

    if score=>replay4 and rep=3 then
    credit=credit+1

    qs=qs+2
    playsound "knocke"

	if credit>25 then credit=25
	credittxt.text=credit
	If B2SOn Then
		Controller.B2SSetCredits Credit
	end if
    rep=4
    playsound "click"
    end if

    end if
End Sub


sub wheelcheck
    If wv=>10 then wv=0
    If wv=0 or wv=1 then
    LightT1.state=1
    LightT2.state=0
    Light6.state=0
    quickstep
    End If

    If wv=2 or wv=3 then
    LightT1.state=1
    LightT2.state=1
    Light6.state=0
    quickstep
    End If

    If wv=4 or wv=5 then
    LightT1.state=1
    LightT2.state=1
    Light6.state=1
    quickstep
    End If

    If wv=6 or wv=7 then
    LightT1.state=0
    LightT2.state=1
    Light6.state=1
    quickstep
    End If

    If wv=8 or wv=9 then
    LightT1.state=0
    LightT2.state=0
    Light6.state=1
    quickstep
    End If
End Sub

Sub quickstep

    If wv=1 and qs>=8 then
    LightT1.state=0
    LightT2.state=0
    Light6.state=0
    End If

    If wv=3 and qs>=12 then
    LightT1.state=0
    LightT2.state=0
    Light6.state=0
    End If

    If (wv=4 or wv=5) and qs>=16 then
    LightT1.state=0
    LightT2.state=0
    Light6.state=0
    End If

    If wv=7 and qs>=20 then
    LightT1.state=0
    LightT2.state=0
    Light6.state=0
    End If

    If wv=9 and qs>=24 then
    LightT1.state=0
    LightT2.state=0
    Light6.state=0
    End If

	If qs>32 then qs=32

End Sub

sub scntimer_timer
	scn1=scn1 + 1
	matchnumb=matchnumb+1
	if matchnumb=10 then matchnumb=0
	if bell=3 then playsound "bell10",0, 0.30, 0, 0
    if bell=2 then playsound "bell100",0, 0.30, 0, 0
    if bell=1 then playsound "bell1000",0, 0.45, 0, 0
    if scn1=scn then
    wv=wv+1
    wheelcheck
    if wv=>10 then wv=0
    scntimer.enabled=false
    end if
End Sub

Sub CheckTilt
	If Tilttimer.Enabled = True Then
	TiltSens = TiltSens + 1
	if TiltSens = 2 Then
	Tilt = True
	tilttxt.text="TILT"
	If B2SOn Then Controller.B2SSetTilt 1
	playsound "tilt"
	if truesc>hisc then hisc=truesc
	reel2.setvalue(hisc)
	savehs
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
	If B2SOn Then Controller.B2SSetTilt 1
	playsound "tilt"
	if truesc>hisc then hisc=truesc
	reel2.setvalue(hisc)
	savehs
	turnoff
	flipopen
	LeftFlipper.RotateToStart
	RightFlipper.RotateToStart
End Sub

Sub Tilttimer_Timer()
	Tilttimer.Enabled = False
End Sub

sub turnoff
    bumper1.force=0
    bumper2.force=0
    bumper3.force=0
End Sub


sub bumper1_hit
  if tilt=false then PlaySoundAt "jet2", bumper1
    if (bumperlight1.state)=1 then
    addscore 100
    else
    addscore 10
    end if
End Sub

sub bumper2_hit
    if tilt=false then PlaySoundAt "jet2", bumper2
    if (bumperlight2.state)=1 then
    addscore 100
    else
    addscore 10
    end if
End Sub

sub bumper3_hit
    if tilt=false then PlaySoundAt "jet2", bumper3
    if (bumperlight3.state)=1 then
    addscore 100
    else
    addscore 10
    end if
End Sub

Sub BumperA_Hit()
	PlaySoundAt "target", BumperA
    Addscore 1000
    If LightA.State=2 then LightCheck
End Sub

Sub BumperB_Hit()
	PlaySoundAt "target", BumperB
    Addscore 1000
    If LightB.State=2 then LightCheck
End Sub

Sub BumperC_Hit()
	PlaySoundAt "target", BumperC
    Addscore 1000
    If LightC.State=2 then LightCheck
End Sub

Sub BumperD_Hit()
	PlaySoundAt "target", BumperD
    Addscore 1000
    If LightD.State=2 then LightCheck
End Sub


Sub Kicker1_Hit()
    If Light6.State=1 then LightCheck
    If Light6.State=1 then Addscore 3000 Else Addscore 300
    PlaySoundAt "kickerkick", Kicker1
    KickTimer.enabled=True
End Sub

Sub Kicker2_Hit()
    If hv=0 then Addscore 500
    if tilt=false then
    hv=hv+1
    If hv>=6 and (Light5.State=1 or Light5.State=2) then Addscore 5000: Light5.State=1
    If hv=5 then Light5.State=2: Addscore 4000: Light4.State=1
    If hv=4 then Light4.State=2: Addscore 3000: Light3.State=1
    If hv=3 then Light3.State=2: Addscore 2000: Light2.State=1
    If hv=2 then Light2.State=2: Addscore 1000: Light1.State=1
    If hv=1 then Light1.State=2
    If hv>6 then hv=0
    If hv=0 then Light1.State=0: Light2.State=0: Light3.State=0: Light4.State=0: Light5.State=0
    End If
    flipclose
    PlaySoundAt "kickerkick", Kicker2
    KickTimer.enabled=True
End Sub


sub kicktimer_timer
    Kicker1.kick (int(rnd(1)*10)+230),12
    Kicker2.kick (int(rnd(1)*10)+170),12
    kicktimer.enabled=false
End Sub

Sub Trigger1_Hit()
    Addscore 10
End Sub

Sub Trigger2_Hit()
    Addscore 10
End Sub

Sub Trigger3_Hit()
	Addscore 100
	BumperLight1.State=1
	BumperLight3.State=1
End Sub

Sub Trigger4_Hit()
	Addscore 1000
    If Light7.State=1 then
    credit= credit+1
	If B2SOn Then
		Controller.B2SSetCredits Credit
	end if
    qs=qs+2
    Playsound "Knocke"
    Light7.State=0
    End If
End Sub

Sub Trigger5_Hit()
	Addscore 100
	BumperLight2.State=1
End Sub

Sub Trigger6_Hit()
	Addscore 1000
End Sub

Sub Trigger7_Hit()
	Addscore 1000
End Sub

Sub Target1_Hit()
    If LightT1.State=1 then LightCheck
    If LightT1.State=1 then Addscore 1000 Else Addscore 100
End Sub

Sub Target2_Hit()
    If LightT2.State=1 then LightCheck
    If LightT2.State=1 then Addscore 1000 Else Addscore 100
End Sub

Sub LightCheck
    If LightD.State=2 then LightD.State=1
    If LightC.State=2 and LightD.State=0 then LightC.State=1: LightD.State=2
    If LightB.State=2 and LightC.State=0 then LightB.State=1: LightC.State=2
    If LightA.State=2 and LightB.State=0 then LightA.State=1: LightB.State=2

    If LightA.State=1 and LightB.State=1 and LightC.State=1 and LightD.State=1 and Light10.State=1 then Light11.State=1:If B2SOn Then Controller.B2SSetData 84,1
    If LightA.State=1 and LightB.State=1 and LightC.State=1 and LightD.State=1 and Light9.State=1 then Light10.State=1:If B2SOn Then Controller.B2SSetData 83,1
    If Light10.State=1 then Light7.State=1
    If Light11.State=1 then Light7.State=1
    If LightA.State=1 and LightB.State=1 and LightC.State=1 and LightD.State=1 and Light8.State=1 then Light9.State=1:If B2SOn Then Controller.B2SSetData 82,1
    If LightA.State=1 and LightB.State=1 and LightC.State=1 and LightD.State=1 then Light8.State=1:If B2SOn Then Controller.B2SSetData 81,1

	If LightA.State=1 and LightB.State=1 and LightC.State=1 and LightD.State=1 and Light8.State=1 and Light9.State=1 and Light10.State=1 and LightD.State=1 then
	credit=credit+1
	qs=qs+2
	playsound "knocke"
	if credit>25 then credit=25
	If B2SOn Then
		Controller.B2SSetCredits Credit
	end if
	End If

    If LightA.State=1 and LightB.State=1 and LightC.State=1 and LightD.State=1 then
    LightA.State=2
    LightB.State=0
    LightC.State=0
    LightD.State=0
    End If
End Sub

Sub LeftSlingShot1_Slingshot()
	Addscore 10
	flipopen
End Sub

Sub LeftSlingShot2_Slingshot()
	Addscore 10
End Sub

Sub LeftSlingShot3_Slingshot()
	Addscore 10
End Sub

Sub LeftSlingShot4_Slingshot()
	Addscore 10
End Sub

Sub RughtSlingShot1_Slingshot()
	Addscore 10
	flipopen
End Sub

Sub RughtSlingShot2_Slingshot()
	Addscore 10
End Sub


sub ballrel_hit
    if ballrelenabled=1 then
    playsound "launchball"
    ballrelenabled=0
    end if
End Sub

sub flipclose
    if tilt=false then
    if (rightflipper.visible)=true then playsound "zclose"
    rightflipper.visible=false
    rightflipper1.visible=true
    leftflipper.visible=false
    leftflipper1.visible=true
    leftflipper.enabled=true
    leftflipper1.enabled=true
    rightflipper.enabled=true
    rightflipper1.enabled=true
    end if
End Sub

sub flipopen
    if (rightflipper.visible)=false then playsound "zopen"
    rightflipper.visible=true
    rightflipper1.visible=false
    leftflipper.visible=true
    leftflipper1.visible=false
    leftflipper.enabled=true
    leftflipper1.enabled=false
    rightflipper.enabled=true
    rightflipper1.enabled=false
End Sub


sub savehs
	' Based on Black's Highscore routines
	Dim FileObj
	Dim ScoreFile
	Set FileObj=CreateObject("Scripting.FileSystemObject")
	If Not FileObj.FolderExists(UserDirectory) then
	Exit Sub
	End if
	Set ScoreFile=FileObj.CreateTextFile(UserDirectory & "4-Queens.txt",True)
	ScoreFile.WriteLine score
   	scorefile.writeline credit
	scorefile.writeline matchnumb
	scorefile.writeline hisc
	scorefile.writeline wv
	scorefile.writeline qs
	ScoreFile.Close
	Set ScoreFile=Nothing
	Set FileObj=Nothing
End Sub

sub loadhs
    ' Based on Black's Highscore routines
	Dim FileObj
	Dim ScoreFile
	dim temp1
	dim temp2
	dim temp3
	dim temp4
	dim temp5
	dim temp6
	Set FileObj=CreateObject("Scripting.FileSystemObject")
	If Not FileObj.FolderExists(UserDirectory) then
	Exit Sub
	End if
	If Not FileObj.FileExists(UserDirectory & "4-Queens.txt") then
	Exit Sub
	End if
	Set ScoreFile=FileObj.GetFile(UserDirectory & "4-Queens.txt")
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
	TextStr.Close
	score = CDbl(temp1)
	credit= CDbl(temp2)
	matchnumb= CDbl(temp3)
	hisc=cdbl(temp4)
    wv=cdbl(temp5)
	qs=cdbl(temp6)
	Set ScoreFile=Nothing
	Set FileObj=Nothing
End Sub

'************************************************
'************************************************
'************************************************
'************************************************
'************************************************


'**********Sling Shot Animations
' Rstep and Lstep  are the variables that increment the animation
'****************
Dim RStep, Lstep

Sub RightSlingShot_Slingshot
    PlaySound "right_slingshot", 0, 1, 0.05, 0.05
    RSling.Visible = 0
    RSling1.Visible = 1
    sling1.TransZ = -20
    RStep = 0
    RightSlingShot.TimerEnabled = 1
    Addscore 10
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
    Addscore 10
End Sub

Sub LeftSlingShot_Timer
    Select Case LStep
    Case 3:LSLing1.Visible = 0:LSLing2.Visible = 1:sling2.TransZ = -10
    Case 4:LSLing2.Visible = 0:LSLing.Visible = 1:sling2.TransZ = 0:LeftSlingShot.TimerEnabled = 0
    End Select
    LStep = LStep + 1
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
    End If
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
End Sub

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

Function AudioFade(ball) ' Can this be together with the above function ?
  Dim tmp
  tmp = ball.y * 2 / Table1.height-1
  If tmp > 0 Then
    AudioFade = Csng(tmp ^10)
  Else
    AudioFade = Csng(-((- tmp) ^10) )
  End If
End Function


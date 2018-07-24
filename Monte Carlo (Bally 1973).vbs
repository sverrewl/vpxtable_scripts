Option Explicit
Randomize

' Thalamus 2018-07-24
' Added/Updated "Positional Sound Playback Functions" and "Supporting Ball & Sound Functions"
' No special SSF tweaks yet.


ExecuteGlobal GetTextFile("core.vbs")
On Error Resume Next
ExecuteGlobal GetTextFile("Controller.vbs")
If Err Then MsgBox "Unable to open Controller.vbs. Ensure that it is in the scripts folder."
On Error Goto 0

Const cGameName = "MonteCarlo_1973"

Const B2STableName="MonteCarlo_1973"

dim score(4)
dim truesc(4)
dim reel(4)
dim ballrelenabled
dim state
dim credit
dim eg
dim currpl
dim playno
dim plno(4)
dim play(4)
dim up(4)
dim rst
dim ballinplay
dim match(9)
dim tilt
dim tiltsens
dim rep(4)
dim plm(4)
dim matchnumb
dim cred
dim bell
dim points
dim tempscore
dim replay1
dim replay2
dim replay3
dim replay4
dim hisc
dim B2SOn
dim scn
dim scn1
dim scn2
dim scn3
dim scn4
dim scn5
dim wv
dim mv
dim wn
dim sbv
dim tt
dim i



Sub Table1_Init
	LoadEM
    scntimer.enabled=false
    set play(1)=plno1
    set play(2)=plno2
    set play(3)=plno3
    set play(4)=plno4
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
    replay1=77000
    replay2=88000
    replay3=99000
    replay4=110000
    loadhs
    if hisc="" then hisc=10000
    hisctxt.text=hisc
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
    for i=1 to 4
    currpl=i
    reel(i).setvalue(score(i))
    next
    currpl=0
	If B2SOn Then

		if matchnumb=0 then
			Controller.B2SSetMatch 100
		else
			Controller.B2SSetMatch matchnumb*10
		end if
		Controller.B2SSetScoreRolloverPlayer1 0
		Controller.B2SSetScoreRolloverPlayer2 0
		Controller.B2SSetScoreRolloverPlayer3 0
		Controller.B2SSetScoreRolloverPlayer4 0

		Controller.B2SSetTilt 1
		Controller.B2SSetCredits credit
		Controller.B2SSetGameOver 1
	End If
	for i=1 to 4
		If B2SOn Then
			Controller.B2SSetScorePlayer i, score(i)
		End If
	next
End Sub

Sub Table1_exit()
	If B2SOn Then Controller.Stop
end sub

Sub Table1_KeyDown(ByVal keycode)

	If keycode = PlungerKey Then
	Plunger.PullBack
	End If

    if keycode = 6 then
	playsound "coin3"
	coindelay.enabled=true
	end if

    if keycode = 5 then
	playsound "coin3"
	coindelay1.enabled=true
	end if

	if keycode = 2 and credit>0 and state=false and playno=0 then
	credit=credit-1
	credittxt.text=credit
    eg=0
    playno=1
    playno1.state=1
    currpl=1
    pno.setvalue(playno)
		If B2SOn Then
			Controller.B2SSetTilt 0
			Controller.B2SSetGameOver 0
			Controller.B2SSetMatch 0
			Controller.B2SSetCredits credit

			Controller.B2SSetCanPlay 1
			Controller.B2SSetPlayerUp 1

			Controller.B2SSetScoreRolloverPlayer1 0
			Controller.B2SSetScoreRolloverPlayer2 0
			Controller.B2SSetScoreRolloverPlayer3 0
			Controller.B2SSetScoreRolloverPlayer4 0
		End If
    play(currpl).state=1
    playsound "click"
    playsound "initialize"
    rst=0
    ballinplay=1
    resettimer.enabled=true
    end if

    if keycode = 2 and credit>0 and state=true and playno>0 and playno<4 and ballinplay<2 then
    credit=credit-1
    credittxt.text=credit
    playno=playno+1
    pno.setvalue(playno)
    if playno=2 then playno2.state=1
    if playno=3 then playno3.state=1
    if playno=4 then playno4.state=1
    playsound "click"
		If B2SOn Then
			Controller.B2SSetCanPlay playno

		End If
    end if

    if state=true and tilt=false then

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
    End if
    End If

	if state=true and tilt=false then
	If keycode = PlungerKey Then Plunger.PullBack: PlaySound "plungerpull",0,1,0.25,0.25: 	End If
	If keycode = LeftFlipperKey Then LeftFlipper.RotateToEnd: PlaySound "flipperup"
	If keycode = RightFlipperKey Then RightFlipper.RotateToEnd: PlaySound "flipperup"
	End If
End Sub

Sub Table1_KeyUp(ByVal keycode)
	If keycode = PlungerKey Then Plunger.Fire: PlaySound "plunger",0,1,0.25,0.25
	If keycode = LeftFlipperKey Then LeftFlipper.RotateToStart: PlaySound "flipperdown"
	If keycode = RightFlipperKey Then RightFlipper.RotateToStart: PlaySound "flipperdown"
End Sub


Sub newgame
	LightReset
	ton1.text=" "
	ton2.text=" "
	ton3.text=" "
	ton4.text=" "
	credtimer.enabled=false
	scntimer.enabled=false
	state=true
	eg=0
	sbv=0
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
	Stopper.Isdropped=True
	Light51.state=0
	nb.CreateBall
	nb.kick 90,6
end sub

sub resettimer_timer
    rst=rst+1
    reel1.resettozero
    reel2.resettozero
    reel3.resettozero
    reel4.resettozero
    if rst=14 then
    playsound "kickerkick"
    end if
    if rst=18 then
    newgame
    resettimer.enabled=false
    end if
end sub

sub coindelay_timer
    playsound "click"
    credit=credit+5
	If B2SOn then
		Controller.B2SSetCredits credit
	end if
	credittxt.text=credit
    coindelay.enabled=false
end sub

sub coindelay1_timer
    playsound "click"
    credit=credit+1
	If B2SOn then
		Controller.B2SSetCredits credit
	end if
	credittxt.text=credit
    coindelay1.enabled=false
end sub

Sub LightReset
    Light1.State=0
    Light2.State=0
    Light3.State=0
    Light4.State=0
    Light5.State=0
    Light6.State=0
    Light7.State=0
    Light8.State=0
    Light9.State=0
    Light10.State=0
    Light11.State=1
    Light12.State=1
    Light13.State=1
    Light14.State=1
    Light15.State=1
    Light16.State=1
    Light17.State=1
    Light18.State=1
    Light19.State=1
    Light20.State=1
    Light21.State=1
    Light22.State=0
    Light23.State=0
    Light24.State=0
    Light25.State=0
    Light26.State=0
    Light27.State=0
    Light28.State=1
    Light29.State=0
    Light30.State=0
    Light31.State=0
    Light32.State=0
    Light33.State=0
    Light34.State=0
    Light35.State=0
    Light36.State=0
    Light37.State=0
    Wall7.IsDropped=False
    Wall8.IsDropped=False
End Sub



'*********************
'*** Points Scoring***
'*********************

sub addscore(points)

    if state=true and tilt=false then

    if tilt=false then bell=0
    if points=10 or points=100 then matchnumb=matchnumb+1
    if matchnumb>=10 then matchnumb=0

    if points = 10 then
    reel(currpl).addvalue(10)
    bell=10
    scn=1
    end if

    if points = 100 then
    reel(currpl).addvalue(100)
    bell=100
    scn=1
    end if

    if points = 1000 then
    reel(currpl).addvalue(1000)
    bell=1000
    scn=1
    end if

    if points = 500 then
    reel(currpl).addvalue(500)
    playsound "motorshort1s"
    bell=100
    scn=5
    end if

    if points = 5000 then
    reel(currpl).addvalue(5000)
    playsound "motorshort1s"
    bell=1000
    scn=5
    end if

    scn1=0
    scntimer.enabled=true

    score(currpl)=score(currpl)+points
    truesc(currpl)=truesc(currpl)+points

    if score(currpl)=>replay1 and rep(currpl)=0 then
    credit=credit+1
    playsound "knocker"
	If B2SOn then
		Controller.B2SSetCredits credit
	end if
    credittxt.text=credit
    rep(currpl)=1
    playsound "click"
    end if

    if score(currpl)=>replay2 and rep(currpl)=1 then
    credit=credit+1
    playsound "knocker"
	If B2SOn then
		Controller.B2SSetCredits credit
	end if
    credittxt.text=credit
    rep(currpl)=2
    playsound "click"
    end if

    if score(currpl)=>replay3 and rep(currpl)=2 then
    credit=credit+1
    playsound "knocker"
	If B2SOn then
		Controller.B2SSetCredits credit
	end if
    credittxt.text=credit
    rep(currpl)=3
    playsound "click"
    end if

    if score(currpl)=>replay4 and rep(currpl)=3 then
    credit=credit+1
    playsound "knocker"
	If B2SOn then
		Controller.B2SSetCredits credit
	end if
    credittxt.text=credit
    rep(currpl)=4
    playsound "click"
    end if
    if score(1)>=100000 then ton1.text="100,000"
    if score(2)>=100000 then ton2.text="100,000"
    if score(3)>=100000 then ton3.text="100,000"
    if score(4)>=100000 then ton4.text="100,000"
	If B2SOn Then
		Controller.B2SSetScorePlayer currpl, score(currpl)
		if score(1)>=100000 then Controller.B2SSetScoreRolloverPlayer1 1
		if score(2)>=100000 then Controller.B2SSetScoreRolloverPlayer2 1
		if score(3)>=100000 then Controller.B2SSetScoreRolloverPlayer3 1
		if score(4)>=100000 then Controller.B2SSetScoreRolloverPlayer4 1
	End If
end If
end sub


sub scntimer_timer
    scn1=scn1 + 1
    if bell=10 then playsound "bell10",0, 0.15, 0, 0
    if bell=100 then playsound "bell100",0, 0.3, 0, 0
    if bell=1000 then playsound "bell1000",0, 0.45, 0, 0
    if scn1>=scn then scntimer.enabled=false
end sub



Sub Trigger1_Hit()
    Light11.State=0
    Light16.State=0
    Light1.State=1
    Light6.State=1
    If Light1.State=1 and Light3.State=1 and Light5.State=1 and Light7.State=1 and Light9.State=1 then Light33.State=1
    If Light2.State=1 and Light4.State=1 and Light6.State=1 and Light8.State=1 and Light10.State=1 then Light34.State=1: Light35.State=1
    If Light33.State=1 and Light34.State=1 then Light36.State=1: Light32.State=1: Wall8.IsDropped=True
    PlaySound "click"
End Sub

Sub Trigger2_Hit()
    Light12.State=0
    Light17.State=0
    Light2.State=1
    Light7.State=1
    If Light1.State=1 and Light3.State=1 and Light5.State=1 and Light7.State=1 and Light9.State=1 then Light33.State=1
    If Light2.State=1 and Light4.State=1 and Light6.State=1 and Light8.State=1 and Light10.State=1 then Light34.State=1: Light35.State=1
    If Light33.State=1 and Light34.State=1 then Light36.State=1: Light32.State=1: Wall8.IsDropped=True
    PlaySound "click"
End Sub

Sub Trigger3_Hit()
    Light13.State=0
    Light3.State=1
    If Light1.State=1 and Light3.State=1 and Light5.State=1 and Light7.State=1 and Light9.State=1 then Light33.State=1
    If Light2.State=1 and Light4.State=1 and Light6.State=1 and Light8.State=1 and Light10.State=1 then Light34.State=1: Light35.State=1
    If Light33.State=1 and Light34.State=1 then Light36.State=1: Light32.State=1: Wall8.IsDropped=True
    PlaySound "click"
End Sub

Sub Trigger4_Hit()
    Light14.State=0
    Light19.State=0
    Light4.State=1
    If Light1.State=1 and Light3.State=1 and Light5.State=1 and Light7.State=1 and Light9.State=1 then Light33.State=1
    If Light2.State=1 and Light4.State=1 and Light6.State=1 and Light8.State=1 and Light10.State=1 then Light34.State=1: Light35.State=1
    Light9.State=1
    If Light33.State=1 and Light34.State=1 then Light36.State=1: Light32.State=1: Wall8.IsDropped=True
    PlaySound "click"
End Sub

Sub Trigger5_Hit()
    Light15.State=0
    Light5.State=1
    If Light1.State=1 and Light3.State=1 and Light5.State=1 and Light7.State=1 and Light9.State=1 then Light33.State=1
    If Light2.State=1 and Light4.State=1 and Light6.State=1 and Light8.State=1 and Light10.State=1 then Light34.State=1: Light35.State=1
    If Light33.State=1 and Light34.State=1 then Light36.State=1: Light32.State=1: Wall8.IsDropped=True
    PlaySound "click"
End Sub

Sub Trigger6_Hit()
    Light11.State=0
    Light16.State=0
    Light1.State=1
    Light6.State=1
    If Light1.State=1 and Light3.State=1 and Light5.State=1 and Light7.State=1 and Light9.State=1 then Light33.State=1
    If Light2.State=1 and Light4.State=1 and Light6.State=1 and Light8.State=1 and Light10.State=1 then Light34.State=1: Light35.State=1
    If Light33.State=1 and Light34.State=1 then Light36.State=1: Light32.State=1: Wall8.IsDropped=True
    PlaySound "click"
End Sub

Sub Trigger7_Hit()
    Light17.State=0
    Light12.State=0
    Light2.State=1
    Light7.State=1
    If Light1.State=1 and Light3.State=1 and Light5.State=1 and Light7.State=1 and Light9.State=1 then Light33.State=1
    If Light2.State=1 and Light4.State=1 and Light6.State=1 and Light8.State=1 and Light10.State=1 then Light34.State=1: Light35.State=1
    If Light33.State=1 and Light34.State=1 then Light36.State=1: Light32.State=1: Wall8.IsDropped=True
    PlaySound "click"
End Sub

Sub Trigger8_Hit()
    Light18.State=0
    Light8.State=1
    If Light1.State=1 and Light3.State=1 and Light5.State=1 and Light7.State=1 and Light9.State=1 then Light33.State=1
    If Light2.State=1 and Light4.State=1 and Light6.State=1 and Light8.State=1 and Light10.State=1 then Light34.State=1: Light35.State=1
    If Light33.State=1 and Light34.State=1 then Light36.State=1: Light32.State=1: Wall8.IsDropped=True
    PlaySound "click"
End Sub

Sub Trigger9_Hit()
    Light14.State=0
    Light19.State=0
    Light4.State=1
    Light9.State=1
    If Light1.State=1 and Light3.State=1 and Light5.State=1 and Light7.State=1 and Light9.State=1 then Light33.State=1
    If Light2.State=1 and Light4.State=1 and Light6.State=1 and Light8.State=1 and Light10.State=1 then Light34.State=1: Light35.State=1
    If Light33.State=1 and Light34.State=1 then Light36.State=1: Light32.State=1: Wall8.IsDropped=True
    PlaySound "click"
End Sub

Sub Trigger10_Hit()
    Light20.State=0
    Light10.State=1
    If Light1.State=1 and Light3.State=1 and Light5.State=1 and Light7.State=1 and Light9.State=1 then Light33.State=1
    If Light2.State=1 and Light4.State=1 and Light6.State=1 and Light8.State=1 and Light10.State=1 then Light34.State=1: Light35.State=1
    If Light33.State=1 and Light34.State=1 then Light36.State=1: Light32.State=1: Wall8.IsDropped=True
    PlaySound "click"
End Sub


Sub MushRoom1_Hit()
	PlaySound "click"
    Light31.State=1
    Wall7.IsDropped=True
    addscore 100
End Sub

Sub MushRoom2_Hit()
	PlaySound "click"
    Light32.State=1
    Wall8.Isdropped=True
    addscore 100
End Sub

Sub Bumper3_Hit()
	If Tilt=False then
	PlaySound "jet1"
	addscore 100
	End If
End Sub

Sub Bumper4_Hit()
	If Tilt=False then
	PlaySound "jet1"
	addscore 100
	End If
End Sub

Sub Bumper5_Hit()
	If Tilt=False then
	PlaySound "jet1"
	addscore 100
	End If
End Sub


Sub Kicker1_Hit()
    If Light34.State=1 then
		Light37.State=1
		if B2SOn then
			Controller.B2SSetShootAgain 1
		end if
	end if
    If Light35.State=1 then addscore 5000 else addscore 500
    If Light22.state=1 then Trigger2_Hit: Light22.State=0
    If Light26.state=1 then Trigger6_Hit: Light26.State=0: Light22.State=1
    If Light30.state=1 then Trigger10_Hit: Light30.State=0: Light26.State=1
    If Light24.state=1 then Trigger4_Hit: Light24.State=0: Light30.State=1
    If Light28.state=1 then Trigger8_Hit: Light28.State=0: Light24.State=1
    If Light1.State=1 and Light3.State=1 and Light5.State=1 and Light7.State=1 and Light9.State=1 then Light33.State=1
    If Light2.State=1 and Light4.State=1 and Light6.State=1 and Light8.State=1 and Light10.State=1 then Light34.State=1: Light35.State=1
    Stopper.IsDropped=False
    Light51.state=1
    KickTimer.enabled=True
End Sub

Sub Kicker2_Hit()
    If Light33.State=1 then addscore 5000 else addscore 500
    If Light27.state=1 then Trigger7_Hit: Light27.State=0
    If Light23.state=1 then Trigger3_Hit: Light23.State=0: Light27.State=1
    If Light29.state=1 then Trigger9_Hit: Light29.State=0: Light23.State=1
    If Light25.state=1 then Trigger5_Hit: Light25.State=0: Light29.State=1
    If Light21.state=1 then Trigger1_Hit: Light21.State=0: Light25.State=1
    If Light1.State=1 and Light3.State=1 and Light5.State=1 and Light7.State=1 and Light9.State=1 then Light33.State=1
    If Light2.State=1 and Light4.State=1 and Light6.State=1 and Light8.State=1 and Light10.State=1 then Light34.State=1: Light35.State=1
    Stopper.IsDropped=False
    Light51.state=1
    KickTimer.enabled=True
End Sub


Sub KickTimer_Timer
    KickTimer.enabled=False
    Kicker1.Kick (int(rnd(1)*20)+90),15
    Kicker2.Kick (int(rnd(1)*20)+165),15
    If Light33.State=1 and Light34.State=1 then Light36.State=1
    Playsound "Eject"
End Sub


Sub Trigger11_Hit()
	Stopper.IsDropped=True
	Light51.state=0
	Playsound "Tilt"
End Sub

Sub Trigger12_Hit()
	Stopper.IsDropped=True
	Light51.state=0
	Playsound "Tilt"
End Sub

Sub Trigger13_Hit()
    If Light36.State=1 then credit=credit+1: Playsound "Knocker": credittxt.text=credit
    Addscore 1000
End Sub

Sub Trigger14_Hit()
    Addscore 1000
End Sub

Sub Trigger15_Hit()
    Addscore 1000
End Sub

Sub Trigger16_Hit()
    Addscore 1000
End Sub

Sub Trigger17_Hit()
    Addscore 1000
End Sub

Sub Trigger18_Hit()
    Addscore 1000
End Sub

Sub Bonus
    scn2=0
    scn3=0

     If Light1.state=1 then
    scn2=scn2+1
    BonusTimer.enabled=True
    End If

    If Light3.state=1 then
    scn2=scn2+1
    BonusTimer.enabled=True
    End If

    If Light5.state=1 then
    scn2=scn2+1
    BonusTimer.enabled=True
    End If

    If Light7.state=1 then
    scn2=scn2+1
    BonusTimer.enabled=True
    End If

    If Light9.state=1 then
    scn2=scn2+1
    BonusTimer.enabled=True
    End If

    If Light2.state=1 then
    scn2=scn2+1
    BonusTimer.enabled=True
    End If

    If Light4.state=1 then
    scn2=scn2+1
    BonusTimer.enabled=True
    End If

    If Light6.state=1 then
    scn2=scn2+1
    BonusTimer.enabled=True
    End If

    If Light8.state=1 then
    scn2=scn2+1
    BonusTimer.enabled=True
    End If

    If Light10.state=1 then
    scn2=scn2+1
    BonusTimer.enabled=True
    End If

    If Light1.State=1 and Light3.State=1 and Light5.State=1 and Light7.State=1 and Light9.State=1 then
    addscore 1000
    BonusTimer.enabled=True
    scn2=scn2+4
    End If

    If Light2.State=1 and Light4.State=1 and Light6.State=1 and Light8.State=1 and Light10.State=1 then
    addscore 1000
    BonusTimer.enabled=True
    scn2=scn2+4
    End If

End Sub


Sub BonusTimer_Timer()
    scn3=scn3+1
    addscore 1000
    if scn3=scn2 then BonusTimer.enabled=False
    if scn3=scn2 then nextball
End Sub


Sub CheckTilt
	If Tilttimer.Enabled = True Then
	TiltSens = TiltSens + 1
	if TiltSens = 2 Then
	Tilt = True
	tilttext.text="TILT"
	tiltsens = 0
	playsound "tilt"
	If B2Son then
		Controller.B2SSetTilt 1
	end if
	Stopper.isdropped=True
	Light51.state=0
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
    tiltseq.play seqalloff
end sub


Sub Drain_Hit()
	playsound "drainshorter"
	Drain.DestroyBall
    If Light1.State=0 and Light3.State=0 and Light5.State=0 and Light7.State=0 and Light9.State=0 and Light2.State=0 and Light4.State=0 and Light6.State=0 and Light8.State=0 and Light10.State=0 then nextball
    Bonus
End Sub

sub nextball
	Stopper.Isdropped=True
	Light51.state=0
	if tilt=true then
	tilt=false
	If B2Son then
		Controller.B2SSetTilt 0
	end if
	tilttext.text=" "
	tiltseq.stopplay
	bipreel.setvalue(1)
	pno.setvalue(playno)
	end if
	If Light37.State=1 then
	If B2SOn then
		Controller.B2SSetShootAgain 0
	end if
	playsound "kickerkick"
	nb.CreateBall
	nb.kick 90,6
	LightReset
	else LightReset
	currpl=currpl+1
	end if
	if currpl>playno then
	ballinplay=ballinplay+1
	if ballinplay>5 then
	playsound "motorleer"
	eg=1
	ballreltimer.enabled=true
	else
	If B2SOn then
		Controller.B2SSetBallInPlay ballinplay
		Controller.B2SSetPlayerUp currpl
	end if
	if state=true and tilt=false then
	play(currpl-1).state=0
	currpl=1
	If B2SOn then
		Controller.B2SSetBallInPlay ballinplay
		Controller.B2SSetPlayerUp currpl
	end if
	play(currpl).state=1
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
	end if
	if currpl>1 and currpl<(playno+1) then
	if state=true and tilt=false then
	play(currpl-1).state=0
	play(currpl).state=1
	If B2SOn then
		Controller.B2SSetBallInPlay ballinplay
		Controller.B2SSetPlayerUp currpl
	end if
	playsound "kickerkick"
	ballreltimer.enabled=true
	end if
	end if
end Sub

sub ballreltimer_timer
	if eg=1 then
	matchnum
	bip3.text=" "
	bip5.text=" "
	bipreel.setvalue(0)
	state=false
	for i=1 to 4
	if truesc(i)>hisc then
	hisc=truesc(i)
	hisctxt.text=hisc
	end if
	next
	pno.setvalue(0)
	play(currpl-1).state=0
	playno=0
	gamov.text="GAME OVER"
	If B2SOn then
		Controller.B2SSetGameOver 1
		Controller.B2SSetPlayerUp 0
		Controller.B2SSetBallInPlay 0
		Controller.B2SSetCanPlay 0
	end if
	savehs
	credtimer.enabled=true
	ballreltimer.enabled=false
	else
	nb.CreateBall
	nb.kick 90,6
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
	If B2SOn then
		If matchnumb = 0 Then
			Controller.B2SSetMatch 100
		Else
			Controller.B2SSetMatch matchnumb*10
		End If
	End if
    for i=0 to playno
    if matchnumb*10=(score(i) mod 100) then
    credit=credit+1
	If B2SOn then
		Controller.B2SSetCredits credit
	end if
    playsound "knocker"
    credittxt.text= credit
    playsound "click"
    end if
    next
end sub

'**********Sling Shot Animations
' Rstep and Lstep  are the variables that increment the animation
'****************
Dim RStep, Lstep

Sub RightSlingShot_Slingshot
    addscore 10
    PlaySound "right_slingshot", 0, 1, 0.05, 0.05
    RSling.Visible = 0
    RSling1.Visible = 1
    sling1.TransZ = -20
    RStep = 0
    RightSlingShot.TimerEnabled = 1
End Sub

Sub RightSlingShot_Timer
    Select Case RStep
    Case 3:RSLing1.Visible = 0:RSLing2.Visible = 1:sling1.TransZ = -10
    Case 4:RSLing2.Visible = 0:RSLing.Visible = 1:sling1.TransZ = 0:RightSlingShot.TimerEnabled = 0
    End Select
    RStep = RStep + 1
End Sub

Sub LeftSlingShot_Slingshot
    addscore 10
    PlaySound "left_slingshot",0,1,-0.05,0.05
    LSling.Visible = 0
    LSling1.Visible = 1
    sling2.TransZ = -20
    LStep = 0
    LeftSlingShot.TimerEnabled = 1
End Sub

Sub LeftSlingShot_Timer
    Select Case LStep
    Case 3:LSLing1.Visible = 0:LSLing2.Visible = 1:sling2.TransZ = -10
    Case 4:LSLing2.Visible = 0:LSLing.Visible = 1:sling2.TransZ = 0:LeftSlingShot.TimerEnabled = 0
    End Select
    LStep = LStep + 1
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
	' Based on Black's Highscore routines
	Dim FileObj
	Dim ScoreFile
	Set FileObj=CreateObject("Scripting.FileSystemObject")
	If Not FileObj.FolderExists(UserDirectory) then
	Exit Sub
	End if
	Set ScoreFile=FileObj.CreateTextFile(UserDirectory & "Monte_Carlo.txt",True)
	scorefile.writeline credit
	ScoreFile.WriteLine score(0)
	ScoreFile.WriteLine score(1)
	ScoreFile.WriteLine score(2)
	ScoreFile.WriteLine score(3)
	scorefile.writeline hisc
	scorefile.writeline matchnumb
	ScoreFile.Close
	Set ScoreFile=Nothing
	Set FileObj=Nothing
end sub

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
	Set FileObj=CreateObject("Scripting.FileSystemObject")
	If Not FileObj.FolderExists(UserDirectory) then
	Exit Sub
	End if
	If Not FileObj.FileExists(UserDirectory & "Monte_Carlo.txt") then
	Exit Sub
	End if
	Set ScoreFile=FileObj.GetFile(UserDirectory & "Monte_Carlo.txt")
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
	TextStr.Close
	credit = CDbl(temp1)
	score(0) = CDbl(temp2)
	score(1) = CDbl(temp3)
	score(2) = CDbl(temp4)
	score(3) = CDbl(temp5)
	hisc = CDbl(temp6)
	matchnumb = CDbl(temp7)
    Set ScoreFile=Nothing
    Set FileObj=Nothing
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


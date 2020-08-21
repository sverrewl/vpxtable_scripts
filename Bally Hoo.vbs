Option Explicit

'*****************************************************************************************************
' CREDITS
' Initial table created by fuzzel, jimmyfingers, jpsalas, toxie & unclewilly (in alphabetical order)
' Flipper primitives by zany
' Ball rolling sound script by jpsalas
' Ball shadow by ninuzzu
' Ball control by rothbauerw
' DOF by arngrim
' Positional sound helper functions by djrobx
' Plus a lot of input from the whole community (sorry if we forgot you :/)
'*****************************************************************************************************


On Error Resume Next
	ExecuteGlobal GetTextFile("controller.vbs")
	If Err Then MsgBox "You need the Controller.vbs file"
	On Error Goto 0


Const cGameName = "Bally_Hoo"
Const BallSize = 25 'Ball radius
B2SOn=True			'Set this to false if you don't use direct2BS


' Thalamus 2020 March : Improved directional sounds
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
Const VolRol    = 1    ' Rollovers volume.
Const VolGates  = 1    ' Gates volume.
Const VolMetal  = 1    ' Metals volume.
Const VolRB     = 1    ' Rubber bands volume.
Const VolRH     = 1    ' Rubber hits volume.
Const VolPo     = 1    ' Rubber posts volume.
Const VolPi     = 1    ' Rubber pins volume.
Const VolPlast  = 1    ' Plastics volume.
Const VolTarg   = 1    ' Targets volume.
Const VolWood   = 1    ' Woods volume.
Const VolKick   = 1    ' Kicker volume.
Const VolSpin   = 1.5  ' Spinners volume.
Const VolFlip   = 1    ' Flipper volume.


Dim Controller
Dim B2SScore
Dim B2SOn
dim score(5)		'Reel score
dim truesc(4)		'True reel score if overring the number of digits
dim reel(5)			'Maximum number of scoring reels
dim state
dim credit			'Credits
dim eg				'End Game
dim currpl			'Current player status
dim playno			'Number of players
dim up(4)			'Player up
dim play(4)			'Player status
dim rst				'Reset Variable
dim bip				'Ball in play
dim match(10)
dim tilt
dim tiltsens
dim rep(4)
dim plm(4)
dim matchnumb
dim obj
dim update
dim cred
dim bell			'Global scoring sound variable
dim points
dim tempscore
dim replay1
dim replay2
dim replay3
dim replay4
dim hisc			'High Score to date
dim ba				'Bonus Advance index
dim batemp			'Bonus advance temporary variable
dim wv				'Wheel Value (positiion)
dim qs				'Quickstep (automatic difficulty level)
dim i				'Indexing variable
dim scn				'Single points score timer
dim scn1			'Single points end score timer
dim scn2			'Multiple points score timer
dim scn3			'Multiple points end score timer
dim scn4			'Multiple points end score timer
dim scn5			'Multiple points end score timer
Dim einsstellung

If Table1.ShowDT=false then
	reel1.Visible=false
	reel2.Visible=false
	reel3.Visible=false
	reel4.Visible=false
	reel5.Visible=false
	playno1.Visible=false
	playno2.Visible=false
	playno3.Visible=false
	playno4.Visible=false
	up1.Visible=false
	up2.Visible=false
	up3.Visible=false
	up4.Visible=false
	bip1.Visible=false
	bip2.Visible=false
	bip3.Visible=false
	bip4.Visible=false
	bip5.Visible=false
	m0.Visible=false
	m1.Visible=false
	m2.Visible=false
	m3.Visible=false
	m4.Visible=false
	m5.Visible=false
	m6.Visible=false
	m7.Visible=false
	m8.Visible=false
	m9.Visible=false
	credittxt.Visible=false
	gamov.Visible=false
tilttext.Visible=false
End If


Sub Table1_Init
	scntimer.enabled=false
	set play(1)=up1
	set play(2)=up2
	set play(3)=up3
	set play(4)=up4
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
	set reel(5)=reel5
	replay1=4300
	replay2=5400
	replay3=6500
	replay4=7600
	loadhs
	if hisc="" then hisc=1000
	if credit="" then credit=0
	reel5.setvalue(hisc)
	credittxt.text=credit
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
	for i=1 to 4
	currpl=i
	reel(i).setvalue(score(i))
	next
	currpl=0
'	B2SOn=True
	if B2SOn then
	Set Controller = CreateObject("B2S.Server")
	Controller.B2SName = "Bally_Hoo"
	Controller.Run()
	If Err Then MsgBox "Can't Load B2S.Server."
	end if
	If B2SOn Then
	if matchnumb=0 then
	Controller.B2SSetMatch 10
	else
	Controller.B2SSetMatch matchnumb
	end if
	If B2SOn Then
	Controller.B2SSetShootAgain 0
	Controller.B2SSetTilt 0
	Controller.B2SSetCredits Credit
	Controller.B2SSetGameOver 1
	Controller.B2SSetData 81,0
	Controller.B2SSetData 82,0
	Controller.B2SSetData 83,0
	Controller.B2SSetData 84,0
	Controller.B2SSetScorePlayer1 0
	Controller.B2SSetScorePlayer2 0
	Controller.B2SSetScorePlayer3 0
	Controller.B2SSetScorePlayer4 0
'	Controller.B2SSetScorePlayer5 hisc
	End If
	End If
	for i=1 to 4
	If B2SOn Then
	Controller.B2SSetScorePlayer(i), 0
	End If
	next
	centerpost.visible=true
	centerpost1.visible=false
	stopper.isdropped=true
	postlight1.state=0
End Sub

Sub Table1_KeyDown(ByVal keycode)

	If keycode = PlungerKey Then
	Plunger.PullBack
	End If

	if keycode = 6 then
	playsoundAtVol "coin3", Drain, 1
	coindelay.enabled=true
	end if

	if keycode = 5 then
	playsoundAtVol "coin3", Drain, 1
	coindelay1.enabled=true
	end if

	if keycode = 2 and credit>0 and state=false and playno=0 then
	credit=credit-1
	qs=qs-1
	if qs<0 then qs=0
	credittxt.text=credit
	eg=0
	playno=1
	playno1.state=1
	currpl=1
	play(currpl).state=1
	playsound "click"
	playsound "initialize"
	rst=0
	bip=1
	resettimer.enabled=true
	If B2SOn Then
	Controller.B2SSetTilt 0
	Controller.B2SSetGameOver 0
	Controller.B2SSetMatch 0
	Controller.B2SSetCredits Credit
	Controller.B2SSetCanPlay 1
	Controller.B2SSetPlayerUp 1
	Controller.B2SSetShootAgain 0
	Controller.B2SSetData 81,0
	Controller.B2SSetData 82,0
	Controller.B2SSetData 83,0
	Controller.B2SSetData 84,0
	Controller.B2SSetBallInPlay bip
	Controller.B2SSetScoreRolloverPlayer1 0
	Controller.B2SSetScorePlayer1 0
	Controller.B2SSetScorePlayer2 0
	Controller.B2SSetScorePlayer3 0
	Controller.B2SSetScorePlayer4 0
'	Controller.B2SSetScorePlayer5 hisc
	End If
    end if

	if keycode = 2 and credit>0 and state=true and playno>0 and playno<4 and bip<2 then
	credit=credit-1
	credittxt.text=credit
	playno=playno+1
	if playno=2 then playno2.state=1: playno1.state=0
	if playno=3 then playno3.state=1: playno2.state=0
	if playno=4 then playno4.state=1: playno3.state=0
	playsound "click"
	If B2SOn Then
	Controller.B2SSetCredits Credit
	Controller.B2SSetCanPlay playno
	end if
    end if

    if state=true and tilt=false then

	If keycode = LeftTiltKey Then
	Nudge 90, 6
	checktilt
	End If

	If keycode = RightTiltKey Then
	Nudge 270, 5
	checktilt
	End If

	If keycode = CenterTiltKey Then
	Nudge 0, 2
	checktilt
    End if
    End If

	if state=true and tilt=false then
	If keycode = PlungerKey Then Plunger.PullBack: PlaySoundAtVol "plungerpull",Plunger, 1: 	End If
	If keycode = LeftFlipperKey Then LeftFlipper.RotateToEnd: PlaySoundAtVol "flipperup", LeftFlipper, 1 : PlayLoopSoundAtVol "buzzl", LeftFlipper, 1
	If keycode = RightFlipperKey Then RightFlipper.RotateToEnd: PlaySoundAtVol "flipperup", RightFlipper, 1 : PlayLoopSoundAtVol "buzzr", RightFlipper, 1
	End If
End Sub

Sub Table1_KeyUp(ByVal keycode)
	If keycode = PlungerKey Then Plunger.Fire: PlaySoundAtVol "plunger", Plunger, 1
	If keycode = LeftFlipperKey and state=true and tilt=false then LeftFlipper.RotateToStart: PlaySoundAtVol "flipperdown", LeftFlipper, 1 : Stopsound "buzzl"
	If keycode = RightFlipperKey and state=true and tilt=false Then RightFlipper.RotateToStart: PlaySoundAtVol "flipperdown", RightFlipper, 1 : Stopsound "buzzr"
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
	bip=1
	centerpost.visible=true
	centerpost1.visible=false
	stopper.isdropped=true
	postlight1.state=0
'	Bumper1.force=11
'	Bumper2.force=11
'	Bumper3.force=11
	wheelcheck
	nb.CreateBall
	nb.kick 90,6
   	bumper1.HasHitEvent=1
	bumper2.HasHitEvent=1
	bumper3.HasHitEvent=1
LeftSlingShot.Disabled =0
RightSlingShot.Disabled =0
ba =1
b1.state =1
	tilt=false
end sub

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
'	Controller.B2SSetScorePlayer5 hisc
	End If
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
	credit=credit+1
If credit >16 then credit =16
	credittxt.text=credit
	coindelay.enabled=false
	If B2SOn Then
	Controller.B2SSetCredits Credit
	End If
end sub

sub coindelay1_timer
	playsound "click"
	credit=credit+1
If credit >16 then credit =16
	credittxt.text=credit
	coindelay1.enabled=false
	If B2SOn Then
	Controller.B2SSetCredits Credit
	End If
end sub

Sub LightReset
	b1.State=0
	b2.State=0
	b3.State=0
	b4.State=0
	b5.State=0
	b6.State=0
	b7.State=0
	b8.State=0
	b9.State=0
	b10.State=0
'	Light14.state=0
'	If B2SOn Then Controller.B2SSetShootAgain 0
	BumperLight1.State=0
	BumperLight2.State=0
	BumperLight3.State=0
	ba=0
End Sub



'*********************
'*** Points Scoring***
'*********************

sub addscore(points)

    if state=true and tilt=false then

    if tilt=false then bell=0
    if points=1 or points=10 then matchnumb=matchnumb+1
    if matchnumb>=10 then matchnumb=0

	if points = 1 then
	reel(currpl).addvalue(1)
	bell=1
	scn=1
	wheelcheck
    end if

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

	if points = 200 then
    reel(currpl).addvalue(200)
    playsound "motorshort1s"
    bell=100
    scn=2
    end if

	if points = 300 then
    reel(currpl).addvalue(300)
    playsound "motorshort1s"
    bell=100
    scn=3
    end if

	if points = 400 then
    reel(currpl).addvalue(400)
    playsound "motorshort1s"
    bell=100
    scn=4
    end if

    if points = 50 then
    reel(currpl).addvalue(50)
    playsound "motorshort1s"
    bell=10
    scn=5
    end if

    if points = 500 then
    reel(currpl).addvalue(500)
    playsound "motorshort1s"
    bell=100
    scn=5
    end if

    scn1=0
    scntimer.enabled=true

	score(currpl)=score(currpl)+points
	truesc(currpl)=truesc(currpl)+points
	If B2SOn Then
	Controller.B2SSetScore currpl,score(currpl)
	end if
	if score(currpl)>9999 then
	score(currpl)=score(currpl)-10000
	rep(currpl)=0
	end if

	if score(currpl)=>replay1 and rep(currpl)=0 then
	credit=credit+1
If credit >16 then credit =16
	qs=qs+2
	playsound "knocker"
	credittxt.text=credit
	If B2SOn Then
	Controller.B2SSetCredits Credit
	end if
    rep(currpl)=1
    playsound "click"
    end if

	if score(currpl)=>replay2 and rep(currpl)=1 then
	credit=credit+1
If credit >16 then credit =16
	playsound "knocker"
	credittxt.text=credit
	If B2SOn Then
	Controller.B2SSetCredits Credit
	end if
	qs=qs+2
    rep(currpl)=2
    playsound "click"
    end if

	if score(currpl)=>replay3 and rep(currpl)=2 then
	credit=credit+1
If credit >16 then credit =16
	qs=qs+2
	playsound "knocker"
	credittxt.text=credit
	If B2SOn Then
	Controller.B2SSetCredits Credit
	end if
    rep(currpl)=3
    playsound "click"
    end if

	if score(currpl)=>replay4 and rep(currpl)=3 then
	credit=credit+1
If credit >16 then credit =16
	qs=qs+2
	playsound "knocker"
	credittxt.text=credit
	If B2SOn Then
	Controller.B2SSetCredits Credit
	end if
    rep(currpl)=4
    playsound "click"
    end if
    if score(1)>=10000 then ton1.text="10,000"
    if score(2)>=10000 then ton2.text="10,000"
    if score(3)>=10000 then ton3.text="10,000"
    if score(4)>=10000 then ton4.text="10,000"
end if
end sub


sub scntimer_timer
	scn1=scn1 + 1
	if bell=1 then playsound "bell1",0, 0.15, 0, 0
	if bell=10 then playsound "bell10",0, 0.3, 0, 0
	if bell=100 then playsound "bell100",0, 0.45, 0, 0
	if scn1>=scn then scntimer.enabled=false
end sub


sub wheelcheck
	If wv>9 or wv<0 then wv=0

	If wv=0 or wv=5 then
	Light2.state=0
	Light3.state=0
	Light4.state=0
	Light5.state=0
	b11.state=0
	Light12.state=1
	Light13.state=0
	quickstep
	End If

	If wv=1 or wv=6 then
	Light2.state=1
	Light3.state=0
	Light4.state=0
	Light5.state=0
	b11.state=1
	Light12.state=0
	Light13.state=0
	quickstep
	End If

	If wv=2 or wv=7 then
	Light2.state=0
	Light3.state=1
	Light4.state=0
	Light5.state=0
	b11.state=0
	Light12.state=1
	Light13.state=0
	quickstep
	End If

	If wv=3 or wv=8 then
	Light2.state=0
	Light3.state=0
	Light4.state=1
	Light5.state=0
	b11.state=0
	Light12.state=1
	Light13.state=0
	quickstep
	End If

	If wv=4 or wv=9 then
	Light2.state=0
	Light3.state=0
	Light4.state=0
	Light5.state=1
	b11.state=1
	Light12.state=0
	Light13.state=1
	quickstep
	End If

	wv=wv+1
end sub


Sub quickstep

	If wv=1 and qs>=8 then
	b11.state=0
	Light12.state=0
	Light13.state=0
    End If

	If wv=3 and qs>=12 then
	b11.state=0
	Light12.state=0
	Light13.state=0
    End If

	If (wv=4 or wv=5) and qs>=16 then
	b11.state=0
	Light12.state=0
	Light13.state=0
    End If

	If wv=7 and qs>=20 then
	b11.state=0
	Light12.state=0
	Light13.state=0
    End If

	If wv=9 and qs>=24 then
	b11.state=0
	Light12.state=0
	Light13.state=1
	End If

End Sub


sub bonusadvance
	if tilt=false then
	ba=ba+1
	if ba>10 then ba=10		'Set the maximim bonus limit
	batemp=ba
    if batemp>9 then
    b10.state=1
    batemp=batemp-10
    end if
    select case (batemp)
    case 1:
    b1.state=1
    case 2:
    b2.state=1
    case 3:
    b3.state=1
    case 4:
    b4.state=1
    case 5:
    b5.state=1
    case 6:
    b6.state=1
    case 7:
    b7.state=1
    case 8:
    b8.state=1
    case 9:
    b9.state=1
    case 10:
    b10.state=1
    end select
    end if
end sub

sub bonuscount_timer
	playsound "motorshorter"
	batemp=ba
	if batemp<11 then b10.state=0
	select case (batemp)
	case 1:
	b1.state=0
	case 2:
	b2.state=0
	case 3:
	b3.state=0
	case 4:
	b4.state=0
	case 5:
	b5.state=0
	case 6:
	b6.state=0
	case 7:
	b7.state=0
	case 8:
	b8.state=0
	case 9:
	b9.state=0
	case 10:
	b10.state=0
	end select
	if (b11.state)=0 then addscore 10
	if (b11.state)=1 then addscore 100
	ba=ba-1
	if ba<=0 then
	bonuscount.enabled=false
	b1.state=0
einsstellung =1
	Kicktimer.enabled=True
    end if
end sub


Sub CheckTilt
	If Tilttimer.Enabled = True Then
	TiltSens = TiltSens + 1
	if TiltSens = 2 Then
	Tilt = True
	tilttext.text="TILT"
	If B2SOn then Controller.B2SSetTilt 1
	tiltsens = 0
	playsound "tilt"
LeftFlipper.RotateToStart: PlaySoundAtVol "flipperdown", LeftFlipper, 1
RightFlipper.RotateToStart: PlaySoundAtVol "flipperdown", RightFlipper, 1
Stopsound "buzzr"
Stopsound "buzzl"
	centerpost.visible=true
	centerpost1.visible=false
If stopper.isdropped=false then PlaysoundAtVol "postdown", centerpost, 1
	stopper.isdropped=true
	postlight1.state=0
   	bumper1.HasHitEvent=0
	bumper2.HasHitEvent=0
	bumper3.HasHitEvent=0
LeftSlingShot.Disabled =1
RightSlingShot.Disabled =1
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
	playsoundAtVol "drainshorter", Drain, 1
	Drain.DestroyBall
	nextball
End Sub

sub nextball
	centerpost.visible=true
	centerpost1.visible=false
	stopper.isdropped=true
	postlight1.state=0
'	Bumper1.force=11
'	Bumper2.force=11
'	Bumper3.force=11
	if tilt=true then
	tilt=false
	tilttext.text=" "
	tiltseq.stopplay
	If B2SOn then Controller.B2SSetTilt 0
	end if
	If Light14.State=1 then
	playsound "kickerkick"
	ballreltimer.enabled=true
	nb.kick 90,6
   	bumper1.HasHitEvent=1
	bumper2.HasHitEvent=1
	bumper3.HasHitEvent=1
LeftSlingShot.Disabled =0
RightSlingShot.Disabled =0
	LightReset
	else LightReset
	currpl=currpl+1
	end if
	if currpl>playno then
	bip=bip+1
	if bip>3 then
	playsound "motorleer"
	eg=1
	ballreltimer.enabled=true
	else
	if state=true and tilt=false then
	play(currpl-1).state=0
	currpl=1
	play(currpl).state=1
	playsound "kickerkick"
	ballreltimer.enabled=true
	end if
	If B2SOn then Controller.B2SSetBallInPlay bip
	If B2SOn then Controller.B2SSetPlayerUp currpl
	select case (bip)
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
	If B2SOn Then
		Controller.B2SSetData 81,0
		Controller.B2SSetData 82,0
		Controller.B2SSetData 83,0
		Controller.B2SSetData 84,0
		Controller.B2SSetData (80+currpl),1
		Controller.B2SSetPlayerUp currpl
		Controller.B2SSetBallInPlay bip

	end If
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
	state=false
	for i=1 to 4
	if truesc(i)>hisc then
	hisc=truesc(i)
	reel(i)=hisc
	end if
	next
	play(currpl-1).state=0
	playno=0
	gamov.text="GAME OVER"
	If B2SOn Then
	Controller.B2SSetGameOver 1
	Controller.B2SSetBallInPlay 0
	Controller.B2SSetPlayerUp 0
'	Controller.B2SSetScorePlayer5 hisc
	end if
	savehs
	reel5.setvalue(hisc)
	credtimer.enabled=true
	ballreltimer.enabled=false
	else
	nb.CreateBall
	nb.kick 90,6
   	bumper1.HasHitEvent=1
	bumper2.HasHitEvent=1
	bumper3.HasHitEvent=1
LeftSlingShot.Disabled =0
RightSlingShot.Disabled =0
ba =1
b1.state =1
    ballreltimer.enabled=false
    end if
end sub


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
	end if
	End If
    for i=0 to playno
    if matchnumb=(score(i) mod 10) then
    credit=credit+1
If credit >16 then credit =16
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
	bonusadvance
    addscore 10
    PlaySoundAtVol "right_slingshot", sling1, 1
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
	bonusadvance
    addscore 10
    PlaySoundAtVol "left_slingshot", sling2, 1
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

' *******************************************************************************************************
' Positional Sound Playback Functions by DJRobX, Rothbauerw and Herweh
' PlaySound sound, 0, Vol(ActiveBall), AudioPan(ActiveBall), 0, Pitch(ActiveBall), 0, 1, AudioFade(ActiveBall)
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

' set position as table object and Vol + RndPitch manually

Sub PlaySoundAtVolPitch(sound, tableobj, Vol, RndPitch)
  PlaySound sound, 1, Vol, AudioPan(tableobj), RndPitch, 0, 0, 1, AudioFade(tableobj)
End Sub

'Set all as per ball position & speed.

Sub PlaySoundAtBall(soundname)
  PlaySoundAt soundname, ActiveBall
End Sub

'Set position as table object and Vol manually.

Sub PlaySoundAtVol(sound, tableobj, Volume)
  PlaySound sound, 1, Volume, AudioPan(tableobj), 0,0,0, 1, AudioFade(tableobj)
End Sub

'Set all as per ball position & speed, but Vol Multiplier may be used eg; PlaySoundAtBallVol "sound",3

Sub PlaySoundAtBallVol(sound, VolMult)
  PlaySound sound, 0, Vol(ActiveBall) * VolMult, AudioPan(ActiveBall), 0, Pitch(ActiveBall), 0, 1, AudioFade(ActiveBall)
End Sub

Sub PlaySoundAtBallAbsVol(sound, VolMult)
  PlaySound sound, 0, VolMult, AudioPan(ActiveBall), 0, Pitch(ActiveBall), 0, 1, AudioFade(ActiveBall)
End Sub

' requires rampbump1 to 7 in Sound Manager

Sub RandomBump(voladj, freq)
  Dim BumpSnd:BumpSnd= "rampbump" & CStr(Int(Rnd*7)+1)
  PlaySound BumpSnd, 0, Vol(ActiveBall)*voladj, AudioPan(ActiveBall), 0, freq, 0, 1, AudioFade(ActiveBall)
End Sub

' set position as bumperX and Vol manually. Allows rapid repetition/overlaying sound

Sub PlaySoundAtBumperVol(sound, tableobj, Vol)
  PlaySound sound, 1, Vol, AudioPan(tableobj), 0,0,1, 1, AudioFade(tableobj)
End Sub

Sub PlaySoundAtBOTBallZ(sound, BOT)
  PlaySound sound, 0, ABS(BOT.velz)/17, Pan(BOT), 0, Pitch(BOT), 1, 0, AudioFade(BOT)
End Sub

' play a looping sound at a location with volume
Sub PlayLoopSoundAtVol(sound, tableobj, Vol)
  PlaySound sound, -1, Vol, AudioPan(tableobj), 0, 0, 1, 0, AudioFade(tableobj)
End Sub

'*********************************************************************
'                     Supporting Ball & Sound Functions
'*********************************************************************

Function RndNum(min, max)
  RndNum = Int(Rnd() * (max-min + 1) ) + min ' Sets a random number between min and max
End Function

Function AudioFade(tableobj) ' Fades between front and back of the table (for surround systems or 2x2 speakers, etc), depending on the Y position on the table. "table1" is the name of the table
  Dim tmp
  On Error Resume Next
  tmp = tableobj.y * 2 / table1.height-1
  If tmp > 0 Then
    AudioFade = Csng(tmp ^10)
  Else
    AudioFade = Csng(-((- tmp) ^10) )
  End If
End Function

Function AudioPan(tableobj) ' Calculates the pan for a tableobj based on the X position on the table. "table1" is the name of the table
  Dim tmp
  On Error Resume Next
  tmp = tableobj.x * 2 / table1.width-1
  If tmp > 0 Then
    AudioPan = Csng(tmp ^10)
  Else
    AudioPan = Csng(-((- tmp) ^10) )
  End If
End Function

Function Pan(ball) ' Calculates the pan for a ball based on the X position on the table. "table1" is the name of the table
  Dim tmp
  On Error Resume Next
  tmp = ball.x * 2 / table1.width-1
  If tmp > 0 Then
    Pan = Csng(tmp ^10)
  Else
    Pan = Csng(-((- tmp) ^10) )
  End If
End Function

Function Vol(ball) ' Calculates the Volume of the sound based on the ball speed
  Vol = Csng(BallVel(ball) ^2 / VolDiv)
End Function

Function BallRollVol(ball) ' Calculates the Volume of the sound based on the ball speed
  BallRollVol = Csng(BallVel(ball) ^2 / (80000 - (79900 * Log(RollVol) / Log(100))))
End Function

Function Pitch(ball) ' Calculates the pitch of the sound based on the ball speed
  Pitch = BallVel(ball) * 20
End Function

Function BallVel(ball) 'Calculates the ball speed
  BallVel = INT(SQR((ball.VelX ^2) + (ball.VelY ^2) ) )
End Function

Function BallVelZ(ball) 'Calculates the ball speed in the -Z
  BallVelZ = INT((ball.VelZ) * -1 )
End Function

Function VolZ(ball) ' Calculates the Volume of the sound based on the ball speed in the Z
  VolZ = Csng(BallVelZ(ball) ^2 / 200)*1.2
End Function

'*** Determines if a Points (px,py) is inside a 4 point polygon A-D in Clockwise/CCW order

Function InRect(px,py,ax,ay,bx,by,cx,cy,dx,dy)
  Dim AB, BC, CD, DA
  AB = (bx*py) - (by*px) - (ax*py) + (ay*px) + (ax*by) - (ay*bx)
  BC = (cx*py) - (cy*px) - (bx*py) + (by*px) + (bx*cy) - (by*cx)
  CD = (dx*py) - (dy*px) - (cx*py) + (cy*px) + (cx*dy) - (cy*dx)
  DA = (ax*py) - (ay*px) - (dx*py) + (dy*px) + (dx*ay) - (dy*ax)

  If (AB <= 0 AND BC <=0 AND CD <= 0 AND DA <= 0) Or (AB >= 0 AND BC >=0 AND CD >= 0 AND DA >= 0) Then
    InRect = True
  Else
    InRect = False
  End If
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
          PlaySound("fx_ballrolling" & b), -1, Vol(BOT(b) ), AudioPan(BOT(b) ), 0, Pitch(BOT(b) ), 1, 0, AudioFade(BOT(b) )
        Else ' Ball on raised ramp
          PlaySound("fx_ballrolling" & b), -1, Vol(BOT(b) )*.5, AudioPan(BOT(b) ), 0, Pitch(BOT(b) )+50000, 1, 0, AudioFade(BOT(b) )
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
  PlaySound("fx_collide"), 0, Csng(velocity) ^2 / (VolDiv/VolCol), AudioPan(ball1), 0, Pitch(ball1), 0, 0, AudioFade(ball1)
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
	PlaySound "gate4"', 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
End Sub

Sub a_Spinner_Spin
	PlaySoundAtVol "fx_spinner", a_spinner, 1
End Sub

Sub a_Rubbers_Hit(idx)
' 	dim finalspeed
'  	finalspeed=SQR(activeball.velx * activeball.velx + activeball.vely * activeball.vely)
' 	If finalspeed > 5 then
	PlaySound "fx_rubber2"', 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
'	End if
'	If finalspeed >= 1 AND finalspeed <= 5 then
' 	RandomSoundRubber()
' 	End If
End Sub

Sub RubberWheel_hit
 	dim finalspeed
  	finalspeed=SQR(activeball.velx * activeball.velx + activeball.vely * activeball.vely)
 	If finalspeed > 5 then
	PlaySound "fx_rubber2", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
	End if
	If finalspeed >= 1 AND finalspeed <= 5 then
 	RandomSoundRubber()
 	End If
End sub

Sub a_Posts_Hit(idx)
 	dim finalspeed
  	finalspeed=SQR(activeball.velx * activeball.velx + activeball.vely * activeball.vely)
 	If finalspeed > 5 then
	PlaySound "fx_rubber2", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
	End if
	If finalspeed >= 1 AND finalspeed <= 5 then
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
	Case 1 : PlaySound "flip_hit_1"', 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
	Case 2 : PlaySound "flip_hit_2"', 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
	Case 3 : PlaySound "flip_hit_3"', 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
	End Select
End Sub

Sub Stopper_Hit
PlaysoundAtVol "pfosten", ActiveBall, 1
End Sub


sub savehs
	' Based on Black's Highscore routines
	Dim FileObj
	Dim ScoreFile
	Set FileObj=CreateObject("Scripting.FileSystemObject")
	If Not FileObj.FolderExists(UserDirectory) then
	Exit Sub
	End if
	Set ScoreFile=FileObj.CreateTextFile(UserDirectory & "Bally_Hoo.txt",True)
	scorefile.writeline credit
	ScoreFile.WriteLine score(0)
	ScoreFile.WriteLine score(1)
	ScoreFile.WriteLine score(2)
	ScoreFile.WriteLine score(3)
	scorefile.writeline hisc
	scorefile.writeline matchnumb
	scorefile.writeline wv
	scorefile.writeline qs
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
	dim temp8
	dim temp9
	Set FileObj=CreateObject("Scripting.FileSystemObject")
	If Not FileObj.FolderExists(UserDirectory) then
	Exit Sub
	End if
	If Not FileObj.FileExists(UserDirectory & "Bally_Hoo.txt") then
	Exit Sub
	End if
	Set ScoreFile=FileObj.GetFile(UserDirectory & "Bally_Hoo.txt")
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
	temp9=Textstr.ReadLine
	TextStr.Close
	credit = CDbl(temp1)
	score(0) = CDbl(temp2)
	score(1) = CDbl(temp3)
	score(2) = CDbl(temp4)
	score(3) = CDbl(temp5)
	hisc = CDbl(temp6)
	matchnumb = CDbl(temp7)
	wv = CDbl(temp8)
	qs = CDbl(temp9)
    Set ScoreFile=Nothing
    Set FileObj=Nothing
end sub


' Table specific scoring

Sub Trigger1_Hit()
PlaysoundAtVol "rollover", ActiveBall, 1
	If tilt=false then Addscore 1
End Sub

Sub Trigger2_Hit()
PlaysoundAtVol "rollover", ActiveBall, 1
	If tilt=false then Addscore 1
End Sub

Sub Trigger3_Hit()
PlaysoundAtVol "sensor", ActiveBall, 1
	If tilt=false then
	Addscore 10
bonusadvance
	BumperLight1.State=1
	End If
End Sub

Sub Trigger4_Hit()
PlaysoundAtVol "sensor", ActiveBall, 1
	If tilt=false then Addscore 100
End Sub

Sub Trigger5_Hit()
PlaysoundAtVol "sensor", ActiveBall, 1
	If tilt=false then
	Addscore 10
bonusadvance
	BumperLight2.State=1
	BumperLight3.State=1
	End If
End Sub

Sub Trigger6_Hit()
PlaysoundAtVol "rollover", ActiveBall, 1
	If tilt=false then
If stopper.isdropped =false then
PlaysoundAtVol "postdown", centerpost, 1
End If
	centerpost.visible=true
	centerpost1.visible=false
	stopper.isdropped=true
	postlight1.state=0
	End If
End Sub

Sub Trigger7_Hit()
PlaysoundAtVol "rollover", ActiveBall, 1
	If tilt=false then
If stopper.isdropped =false then
PlaysoundAtVol "postdown", centerpost, 1
End If
	centerpost.visible=true
	centerpost1.visible=false
	stopper.isdropped=true
	postlight1.state=0
	End If
End Sub

Sub Trigger8_Hit()
PlaysoundAtVol "rollover", ActiveBall, 1
	If tilt=false then
If stopper.isdropped =true then
PlaysoundAtVol "postup", centerpost, 1
End If
	Addscore 10
	centerpost.visible=false
	centerpost1.visible=true
	stopper.isdropped=false
	postlight1.state=1
	End If
End Sub

Sub Trigger9_Hit()
PlaysoundAtVol "sensor", ActiveBall, 1
	If tilt=false then Addscore 100
End Sub

Sub Trigger10_Hit()
PlaysoundAtVol "sensor", ActiveBall, 1
    If tilt=false then Addscore 100
End Sub

'Sub Trigger11_Hit()
'    If tilt=false then Addscore 100
'End Sub

Sub Trigger17_Hit()
PlaysoundAtVol "sensor", ActiveBall, 1
    If tilt=false then Addscore 100
End Sub

Sub LeftSlingShot1_Slingshot()
	If tilt=false then Addscore 1
End Sub

Sub LeftSlingShot2_Slingshot()
	If tilt=false then Addscore 1
End Sub

Sub LeftSlingShot3_Slingshot()
	If tilt=false then Addscore 1
End Sub

Sub RightSlingShot1_Slingshot()
	If tilt=false then Addscore 1
End Sub

Sub RightSlingShot2_Slingshot()
	If tilt=false then
	Bonusadvance
	Addscore 10
	End If
End Sub

Sub RightSlingShot3_Slingshot()
	If tilt=false then Addscore 1
End Sub

Sub MushRoom1_Hit()
PmushL3.transy =7
Timer2.enabled =1
PlaySoundAtVol "click", ActiveBall, 1
If tilt = True then Exit Sub
	addscore 10
If stopper.isdropped =true then
PlaysoundAtVol "postup", centerpost, 1
End If
	centerpost.visible=false
	centerpost1.visible=true
	stopper.isdropped=false
	postlight1.state=1
End Sub

Sub Timer2_Timer
PmushL3.transy =0
Me.enabled =0
End Sub

Sub Bumper1_Hit()
	If Tilt=False then
	PlaySoundAtVol "jet1", ActiveBall, 1
	If BumperLight1.State=1 Then Addscore 10 Else addscore 1
	End If
End Sub

Sub Bumper2_Hit()
	If Tilt=False then
	PlaySoundAtVol "jet1", ActiveBall, 1
	If BumperLight2.State=1 Then Addscore 10 Else addscore 1
	End If
End Sub

Sub Bumper3_Hit()
	If Tilt=False then
	PlaySoundAtVol "jet1", ActiveBall, 1
	If BumperLight3.State=1 Then Addscore 10 Else addscore 1
	End If
End Sub


Sub Kicker1_Hit()
PlaysoundAtVol "ballrein", Kicker1, 1
Playsound "motorleise"
If Tilt =True then KickTimer.enabled=True : Exit Sub
    if tilt=false then
    If Light5.State=1 then Addscore 500
    If Light4.State=1 then Addscore 400
    If Light3.State=1 then Addscore 300
    If Light2.State=1 then Addscore 200
    If Light2.State=0 and Light3.State=0 and Light4.State=0 and Light5.State=0 then addscore 100
    End If
    KickTimer.enabled=True
End Sub

Sub Kicker2_Hit()
PlaysoundAtVol "ballrein", Kicker2, 1
Playsound "motorleise",-1
If Tilt =True then KickTimer.enabled=True : Exit Sub
	if tilt=false then
	bonuscount.enabled=true
	end if
End Sub

Sub Kicker3_Hit()
PlaysoundAtVol "ballrein", Kicker3, 1
Playsound "motorleise"
If Tilt =True then Timer3.enabled=True : Exit Sub
	if tilt=false then
	If Light13.State=1 then
	Light14.State=1
	If B2SOn Then Controller.B2SSetShootAgain 1
	End If
'	Addscore 100
	end if
	Timer3.enabled=True
End Sub

Sub Kicker4_Hit()
If Tilt =True then Kicker4.Kick 120,3 : Exit Sub
	If Light12.State=1 then Kicker4.Kick 0,25 : PlaysoundAtVol "direktabschuss", Kicker4, 1: Else Kicker4.Kick 120,3
	Light12.State=0
End Sub

Sub KickTimer_Timer
    KickTimer.enabled=False
    Kicker1.Kick (int(rnd(1)*10)+170),15
    Kicker2.Kick (int(rnd(1)*5)+180),15
'    Kicker3.Kick 0,(int(rnd(1)*10)+30)
hebel1.rotz =15
hebel2.rotz =15
Timer1.enabled =1
    Playsound "Eject"
Stopsound "motorleise"
If einsstellung =1 then
ba=1
b1.state =1
einsstellung =0
End If
End Sub

Sub Timer3_Timer
Kicker3.Kick 0,(int(rnd(1)*10)+30)
Playsound "direktabschuss"
Stopsound "motorleise"
Me.enabled =0
End Sub


Sub Timer1_Timer
hebel1.rotz =0
hebel2.rotz =0
Me.enabled =0
End Sub

Sub Gate_Hit
Light14.state=0
If B2SOn Then Controller.B2SSetShootAgain 0
PlaysoundAtVol "gate", ActiveBall, 1
End Sub

Dim zeit
Dim lampe1
Dim lampe2
Dim lampe3
Dim lampe4
Dim lampe5
Dim min
Dim max

Sub licht_Timer
min =0
max =4
lampe1=(Int((max-min+1)*Rnd+min))
min =0
max =4
lampe2=(Int((max-min+1)*Rnd+min))
min =0
max =4
lampe3=(Int((max-min+1)*Rnd+min))
min =0
max =4
lampe4=(Int((max-min+1)*Rnd+min))
min =0
max =4
lampe5=(Int((max-min+1)*Rnd+min))
If B2SOn then
If lampe1 >0 then Controller.B2SSetData 91,1
If lampe1 =0 then Controller.B2SSetData 91,0
If lampe2 >0 then Controller.B2SSetData 92,1
If lampe2 =0 then Controller.B2SSetData 92,0
If lampe3 >0 then Controller.B2SSetData 93,1
If lampe3 =0 then Controller.B2SSetData 93,0
If lampe4 >0 then Controller.B2SSetData 94,1
If lampe4 =0 then Controller.B2SSetData 94,0
If lampe5 >0 then Controller.B2SSetData 95,1
If lampe5 =0 then Controller.B2SSetData 95,0
End If
min =250
max =500
zeit=(Int((max-min+1)*Rnd+min))
licht.Interval =zeit
End Sub

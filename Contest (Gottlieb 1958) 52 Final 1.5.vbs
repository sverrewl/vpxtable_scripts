Const cGameName = "contest_1958"
Const ReflectionMod=True			'enable JPJ reflection mod

' Thalamus 2018-07-20
' Added/Updated "Positional Sound Playback Functions" and "Supporting Ball & Sound Functions"
' No special SSF tweaks yet.

Dim Score(4)
dim truesc(4)
dim ballrelenabled
dim activeballs
dim state
dim playno
dim credit
dim eg
dim currpl
dim play(4)
dim rst
dim ballinplay
dim tilt
dim tiltsens
dim rep(4)
dim plm(4)
dim matchnumb
dim cred
dim scn
dim scn1
dim scn2
dim scn3
dim points
dim tempscore
dim replay1
dim replay2
dim replay3
dim replay4
dim hisc
dim pltilt(4)
dim motion
dim rv
dim rt
dim rt1
dim tt
dim rsv 'Rotor initial score value without multiply
dim ylv 
dim rstep, rstepb, rstepc, rstepd, rstepe, lstep, lstepb, lstepc, lstepd, lstepe
dim BallRadius, BallMass, cball1
dim LaunchBall
dim automatic
dim rotorR, rotation
dim wheelprice
dim c, a
dim EMR(4)
dim GL, bumpT, bumpB 'state of gi left and right and high bumper states
DIM HZ 'Hit Zone for the gates to provide double hits
dim b2svarlightA, b2svarlightB, b2sSunL, b2sSunR
Dim GlobalSoundLevel 'Diner's Method for amplify mechanical sounds, thanks to them :

GlobalSoundLevel = 2

GL="L"
bumpT=0
bumpB=0

credit = 0

On Error Resume Next
ExecuteGlobal GetTextFile("controller.vbs")
If Err Then MsgBox "You need the controller.vbs in order to run this table, available in the vp10 package"
On Error Goto 0


sub table1_init
	LoadEM
	If ShowDT=false Then
		for each obj in DTItems
			obj.visible=False
		Next
	Primitive78.visible = 0
	Primitive77.visible = 0
	Primitive81.visible = 0
	Primitive82.visible = 0
	Primitive83.visible = 0
	Primitive84.visible = 0
	RampLBlack.visible = 1
	RampRBlack.visible = 1
	RampLBlack1.visible = 1
	RampRBlack1.visible = 1
		Else

	RampLBlack.visible = 0
	RampRBlack.visible = 0
	RampLBlack1.visible = 0
	RampRBlack1.visible = 0

	end If

'***********************************************************************
'*** Automatic launch ball or not, if not use the RightMagnaSave Key ***
'*** 0 = RightMagnaSave Key - 1 = automatic                          ***
'***********************************************************************
	automatic = 1                                                   '***
'***********************************************************************
'***********************************************************************
	Scoreb2s.enabled = 1 '*** Activation of the Score sync on the b2s **
'***********************************************************************

	set EMR(1)=score1
	set EMR(2)=score2
	set EMR(3)=score3
	set EMR(4)=score4
	if credit="" then 
		credit=0
	end If
    credittxt.text=credit
If credit > 0 Then DOF 119, DOFOn
	GL="L"
	GILeft.enabled = 1


	BallRadius = 25
	BallMass = 2.5
	bouton01.z=2.5	
	bouton01f.z=2.5
	bouton02.z=2.5
	bouton02f.z=2.5
	bouton03.z=2.5
	bouton03f.z=2.5
	rotorR=Wheel.objroty
	Wheelprice = 10
	c=0

    randomize
    replay1=995
    replay2=1300
    replay3=1700
    replay4=1850
    r1.text=replay1
    r2.text=replay2
    r3.text=replay3
    r4.text=replay4

    for each obj in droppers
    obj.isdropped=true
    next

    loadhs 

    if hisc="" then hisc=1000
    hstxt.text=hisc
    if credit="" then credit=0 
    credittxt.text=credit
	if B2SOn then 
	Scoreb2s.enabled =1 
	end If
    if matchnumb="" then matchnumb=int(rnd(1)*9)
    match(matchnumb).text=matchnumb 
    for i=0 to 3
    currpl=i
    reel(i).setvalue(score(i)) 
    if score(i)>9999 then threel(i).setvalue(1)
    next
    currpl=0
    cred=0

    credtimer.enabled=true

    if rv="" then rv=0
    if ylv="" then ylv=0
    ylupdate

	AutomaticBallLaunch.enabled=True


'************************************************************
'***             B2S                                      ***
'************************************************************
	If B2SOn Then 
	for i=0 to 3
			Controller.B2SSetScore i, Score(i)
	next
	End If

HZ = 0
Ballinplay = 0
b2svarlightA=1
b2svarlightB=2

b2svar1.enabled = 1
b2svar2.enabled = 1
SunL.enabled = 1
b2sSunL=1
b2sSunR=1

 if motion=0 and tilt=false and c=0 then
    rt=(int(rnd(1)*8)+4)
    motion=1	
	c=1
	a=int((rnd*6)+1)
    rotort.enabled=true
 end if
end sub
'******** end of table ini *********


Sub Table1_KeyDown(ByVal keycode)
    If keycode = PlungerKey Then
		Plunger.PullBack
	End If

	
    if keycode = AddCreditKey then
	playsound "coin3"
	coindelay.enabled=true 
	end if
    
    if keycode = StartGameKey and credit>0 and state=false and playno=0 then
	credit=credit-1
If credit < 1 Then DOF 119, DOFOff
	credittxt.text=credit
	
	eg=0
    playno=1
    currpl=0
    activeballs=0
    tilt=false
	tilttxt.text=" "
	bumper1.force=5
    bumper2.force=5
    tiltseq.stopplay


    for each obj in indtilt
    obj.text=" "
    next

  
    gamov.text="GAME OVER"
    playup(currpl).setvalue(1)  
    player.setvalue(1)
    playsound "click"
    rst=0
    playsound "initialize"
    resettimer.enabled=true
	if B2SOn then Scoreb2s.enabled =1 
    end if 
    
    if keycode = StartGameKey and credit>0 and state=true and playno>0 and playno<4 and ballinplay<2 then
    credit=credit-1
If credit < 1 Then DOF 119, DOFOff
    credittxt.text=credit
    player.addvalue(1) 
    playno=playno+1
    playsound "click"
	if B2SOn then Scoreb2s.enabled =1 
	End If


    if automatic = 0 and keycode = RightMagnaSave and activeballs=0 and state=true and ballinplay<6 then
    activeballs=activeballs+1
	nb.CreateSizedballWithMass BallRadius,BallMass 'old nb.createball 
    nb.kick 270,10
playsound SoundFXDOF("ballout",115,DOFPulse,DOFContactors)
    end if

    if tilt=false and state=true then  
        
    If keycode = LeftFlipperKey Then
		LeftFlipper.RotateToEnd
		Flipper1.RotateToEnd
PlaySound SoundFXDOF("FlipperUp",101,DOFOn,DOFFlippers)
		playsound "buzz"
	End If
    
	If keycode = RightFlipperKey Then
		RightFlipper.RotateToEnd
        Flipper2.RotateToEnd
PlaySound SoundFXDOF("FlipperUp",102,DOFOn,DOFFlippers)
		playsound "buzz1"
	End If
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

End Sub



Sub Table1_KeyUp(ByVal keycode)

	If keycode = PlungerKey Then
		Plunger.Fire
		playsound "plunger"
	End If
    
	If keycode = LeftFlipperKey Then
		LeftFlipper.RotateToStart
		Flipper1.RotateToStart
if state=true and tilt=false then PlaySound SoundFXDOF("FlipperDown",101,DOFOff,DOFFlippers)
		stopsound "buzz"
	End If
    
	If keycode = RightFlipperKey Then
		RightFlipper.RotateToStart
        Flipper2.RotateToStart
if state=true and tilt=false then PlaySound SoundFXDOF("FlipperDown",102,DOFOff,DOFFlippers)
		stopsound "buzz1"
	End If

End Sub

sub AutomaticBallLaunch_Timer
    if automatic = 1 and activeballs=0 and state=true and ballinplay<6 then
    activeballs=activeballs+1
	nb.CreateSizedballWithMass BallRadius,BallMass
    nb.kick 270,10
DOF 115, DOFPulse
    end if
End Sub

sub coindelay_timer
    playsound "click" 
    credit=credit+1
DOF 119, DOFOn
	credittxt.text=credit 
    coindelay.enabled=false
	if B2SOn then Scoreb2s.enabled =1 
end sub

sub resettimer_timer
    rst=rst+1
    if rst=6 or rst=12 or rst=1 then 
    tiltseq.play seqalloff
    playsound "clerker"
    for each obj in plast
    obj.isdropped=true
    next
    end if
    if rst=7 or rst=13 or rst=2 then 
    tiltseq.stopplay
    playsound "relay"
    ylv=ylv+1
    if ylv>4 then ylv=0
    ylupdate
    for each obj in plast
    obj.isdropped=false
    next
    end if
    for i=0 to 3
    reel(i).resettozero
    next
    if rst=14 then
    playsound "kickerkick"
    end if
    ylv=ylv+1
    if ylv>4 then ylv=0
    ylupdate
    if rst=18 then
    newgame
    resettimer.enabled=false
    end if
end sub

sub addscore(points)
    if tilt=false then
    bell=0
    if points=10 or points=1 then 
    matchnumb=matchnumb+1
    if matchnumb=10 then matchnumb=0
    end if
    if points = 100 then 
    reel(currpl).addvalue(100)
playsound SoundFXDOF("gigibump4loud",142,DOFPulse,DOFChimes)
    end if
    if points = 10 then 
    reel(currpl).addvalue(10)
playsound SoundFXDOF("gigibump3loud",141,DOFPulse,DOFChimes)
    end if
    if points = 1 then 
    reel(currpl).addvalue(1)
    ylv=ylv+1
    if ylv>4 then ylv=0
    ylupdate
    end if
    if points = 5 then 
    playsound "motorshort1s"
    reel(currpl).addvalue(5)
    end if
    if points = 20 then 
    reel(currpl).addvalue(20)
    scn=2
    scn1=0
    playsound "motorshort1s"
    scntimer.enabled=true
    end if
    if points = 30 then 
    reel(currpl).addvalue(30)
    scn=3
    scn1=0
    playsound "motorshort1s"
    scntimer.enabled=true
    end if
    if points = 40 then 
    reel(currpl).addvalue(40)
    scn=4
    scn1=0
    playsound "motorshort1s"
    scntimer.enabled=true
    end if
    if points = 50 then 
    reel(currpl).addvalue(50)
    scn=5
    scn1=0
    playsound "motorshort1s"
    scntimer.enabled=true
    end if
    if points = 200 then 
    reel(currpl).addvalue(200)
    scn2=2
    scn3=0
    playsound "motorshort1s"
    scntimer1.enabled=true
    end if
    if points = 300 then 
    reel(currpl).addvalue(300)
    scn2=3
    scn3=0
    playsound "motorshort1s"
    scntimer1.enabled=true 
    end if
    if points = 400 then 
    reel(currpl).addvalue(400)
    scn2=4
    scn3=0
    playsound "motorshort1s"
    scntimer1.enabled=true
    end if
    if points = 500 then 
    reel(currpl).addvalue(500)
    scn2=5
    scn3=0
    playsound "motorshort1s"
    scntimer1.enabled=true
    end if 
    score(currpl)=score(currpl)+points
    truesc(currpl)=truesc(currpl)+points
    'if truesc(currpl)>1999 then exsc(currpl).text=truesc(currpl)
    if score(currpl)>9999 then threel(currpl).setvalue(1)
'    if score(currpl)>1999 then
'    score(currpl)=score(currpl)-1000
'    rep(currpl)=0
'    end if
    if score(currpl)=>replay1 and rep(currpl)=0 then 
    credit=credit+1
DOF 119, DOFOn
playsound SoundFXDOF("knocke",117,DOFPulse,DOFKnocker)
DOF 118, DOFPulse
    credittxt.text=credit
    rep(currpl)=1
    playsound "click"
	if B2SOn then Scoreb2s.enabled =1 
    end if
    if score(currpl)=>replay2 and rep(currpl)=1 then
    credit=credit+1
DOF 119, DOFOn
playsound SoundFXDOF("knocke",117,DOFPulse,DOFKnocker)
DOF 118, DOFPulse
    credittxt.text=credit
    rep(currpl)=2
    playsound "click"
	if B2SOn then Scoreb2s.enabled =1 
	End if
    if score(currpl)=>replay3 and rep(currpl)=2 then
    credit=credit+1
DOF 119, DOFOn
playsound SoundFXDOF("knocke",117,DOFPulse,DOFKnocker)
DOF 118, DOFPulse
    credittxt.text=credit
    rep(currpl)=3
    playsound "click"
	if B2SOn then Scoreb2s.enabled =1 
    end if
    if score(currpl)=>replay4 and rep(currpl)=3 then
    credit=credit+1
DOF 119, DOFOn
playsound SoundFXDOF("knocke",117,DOFPulse,DOFKnocker)
DOF 118, DOFPulse
    credittxt.text=credit
    rep(currpl)=4
    playsound "click"
	if B2SOn then Scoreb2s.enabled =1 
    end if
    end if
end sub  

sub scntimer_timer
    scn1=scn1 + 1
playsound SoundFXDOF("gigibump3loud",141,DOFPulse,DOFChimes)
    if scn1=scn then scntimer.enabled=false
end sub

sub scntimer1_timer
    scn3=scn3 + 1
playsound SoundFXDOF("gigibump4loud",142,DOFPulse,DOFChimes)
    if scn2=scn3 then scntimer1.enabled=false
end sub

sub matchnum
    match(matchnumb).text=matchnumb
    for i=1 to playno
    if matchnumb=(score(i-1) mod 10) then
    credit=credit+1
DOF 119, DOFOn
playsound SoundFXDOF("knocke",117,DOFPulse,DOFKnocker)
DOF 118, DOFPulse
    credittxt.text= credit
    playsound "click"
    end if
    next 
    playno=0
end sub

Sub CheckTilt
	If Tilttimer.Enabled = True Then
	TiltSens = TiltSens + 1
	if TiltSens = 2 Then
	Tilt = True
	tilttxt.text="TILT"
	tiltsens = 0
	pltilt(currpl)=1
	indtilt(currpl).text="TILT"
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
	tiltsens = 0
	pltilt(currpl)=1
	indtilt(currpl).text="TILT"
	playsound "tilt"
	turnoff 
	LeftFlipper.RotateToStart
	RightFlipper.RotateToStart	
End Sub

Sub Tilttimer_Timer()
	Tilttimer.Enabled = False
End Sub

sub turnoff
    bumper1.force=0
    bumper2.force=0
    tiltseq.play seqalloff 
    for each obj in match
    obj.text=" "
    next
    playup(currpl).setvalue(0)
    player.setvalue(0)
    for each obj in ballplay
    obj.text=" "
    next
for each obj in GIGEN
	obj.state=0
Next
for each obj in GIR
	obj.state=0
Next
for each obj in GIL
	obj.state=0
Next
	currpl=0
	ballinplay=0
end sub

sub newgame
    credtimer.enabled=false
    credtxt.text="Ported To VP By Leon Spalding, Ported To VPX by Team PP, JPJ, Chucky"
    state=true
    eg=0
    for i=0 to 3
    score(i)=0
    truesc(i)=0
    rep(i)=0
    pltilt(i)=0
    threel(i).setvalue(0)
    next
    bumper1.force=5
    bumper2.force=5 
for each obj in GIGEN
	obj.state=1
Next

'    for each obj in lights
'    obj.state=0
'    next

    on1.state=1
    off2.state=1
    light3.state=1:light3z.state=1
    light4.state=1:light4z.state=1
    light5.state=1:light5z.state=1

    if motion=0 then
    rt=(int(rnd(1)*8)+4)
    motion=1
    rotort.enabled=true
    end if
    bip5.text=" " 
    bip1.text="1"
    for i=0 to 9
    match(i).text=" "
    next
    tilttxt.text=" "  
    gamov.text=" " 
    tilt=false
    tiltsens=0
    ballinplay=1
    playsound "reset1"
	if B2SOn then Scoreb2s.enabled =1 
end sub

Sub Drain_Hit()
    playsound "drainshort"
DOF 116, DOFPulse
	Drain.DestroyBall
	if ballinplay<5 then playsound "motorshort"
    activeballs=activeballs-1
	if activeballs=0 then tt=0:nextball 
	if B2SOn then Scoreb2s.enabled =1 
End Sub

sub nextball
    if ballinplay<5 and tilt=false then playsound "motorshorter1s"

    for i=0 to 4
    if pltilt(i)=1 then tt=tt+1
    next

for each obj in GIGEN
	obj.state=1
Next

    if tt=playno then
    for i=0 to 4
    ballplay(i).text=" "
    next
	currpl=currpl+1
	state=false
	savehs
	playno=0
	'ballreltimer.enabled=true
	exit sub
	end if	
	if tilt=true then 
	tilt=false
	tilttxt.text=" "
	bumper1.force=6
    bumper2.force=6
    tiltseq.stopplay
    ballplay(ballinplay).text=ballinplay
    player.setvalue(playno)
	end if
	currpl=currpl+1
	if currpl>playno-1 then
	ballinplay=ballinplay+1
	playsound "relay"
	if ballinplay>5 then
	playsound "motorleer"
	ballreltimer.enabled=true
	else
	if state=true and tilt=false then
	playup(currpl-1).setvalue(0)
	currpl=0
	playup(currpl).setvalue(1)
	playsound "relay"
	tilttxt.text=" " 
	end if
	select case ballinplay
	case 0: 
	bip1.text=" "
	bip2.text=" "
	bip3.text=" "
	bip4.text=" "
	bip5.text=" "
	case 1:
	bip1.text="1"
	bip2.text=" "
	bip3.text=" "
	bip4.text=" "
	bip5.text=" "
	case 2: 
	bip1.text=" "
	bip2.text="2"
	bip3.text=" "
	bip4.text=" "
	bip5.text=" "
	case 3:
	bip1.text=" "
	bip2.text=" "
	bip3.text="3"
	bip4.text=" "
	bip5.text=" "
	case 4:
	bip1.text=" "
	bip2.text=" "
	bip3.text=" "
	bip4.text="4"
	bip5.text=" "
	case 5:
	bip1.text=" "
	bip2.text=" "
	bip3.text=" "
	bip4.text=" "
	bip5.text="5"
	end select
	playsound "relay"
	end if
	end if
	if currpl>0 and currpl<(playno) then
	if state=true and tilt=false then
	playup(currpl-1).setvalue(0)
	playup(currpl).setvalue(1)
	playsound "relay"
	tilttxt.text=" "  
	end if
	end if
	if pltilt(currpl)=1 then tilttxt.text="TILT":tt=0:nextball
end Sub

sub ballreltimer_timer
    if tilt=false then matchnum
	bip5.text=" "
	state=false
	for i=0 to 3
	if truesc(i)>hisc then 
	hisc=truesc(i)
	hstxt.text=hisc
	end if
	next
	playup(currpl-1).setvalue(0)
	player.setvalue(0)
	gamov.text="GAME OVER"
	savehs
	cred=0
	credtimer.enabled=true
	if ballinplay > 5 then ballinplay = 0
	ballreltimer.enabled=false
end sub

'**********************************************
'**            New wheel part jpj            **
'**********************************************

sub rotort_timer()

	rotorR = Wheel.objroty
playsound SoundFXDOF("lightmotor",120,DOFOn,DOFGear)
	rotationW.enabled=True

   rt1=rt1+1
'   rotoradv mis à la fin
   if rt1=rt*2 then 
   rt1=0
   motion=0
   ylupdate
	end If
   if rsv<40 then
   l400.state=0:l400z.state=0
   l500.state=0:l500z.state=0
   end if  

end sub




sub rotationW_timer
		select case c
			case 1 : 
				Wheel.objroty = Wheel.objroty + 9
				if wheel.objroty = 360 then wheel.objroty = 0:end If
				c=2
			case 2 : 
				Wheel.objroty = Wheel.objroty + 9
				if wheel.objroty = 360 then wheel.objroty = 0:end If
				c=3
			case 3 : 
				Wheel.objroty = Wheel.objroty + 9
				if wheel.objroty = 360 then wheel.objroty = 0:end If
				c=4
			case 4 : 
				Wheel.objroty = Wheel.objroty + 9
				if wheel.objroty = 360 then wheel.objroty = 0:end If
				a=a-1
				if a >0 then 
					c=1 
				Else 
				if a=0 then c=5
				rotoradv
				end If

			case 5 : 
				Wheel.objroty = Wheel.objroty + 25
				c=6
			case 6 : 
				Wheel.objroty = Wheel.objroty + 10
				c=7
			case 7 : 
				Wheel.objroty = Wheel.objroty + 5
				c=8
			case 8 : 
				Wheel.objroty = Wheel.objroty +2
				c=9
			case 9 : 
				Wheel.objroty = Wheel.objroty - 7
				c=10
			case 10 : 
				Wheel.objroty = Wheel.objroty -46
				c=11
			case 11 : 
				Wheel.objroty = Wheel.objroty +11
c=0:rotationW.enabled=false:DOF 120, DOFOff
		end Select


	motion=0 and tilt=false


	rotort.enabled=false
if c=0 then rotationW.enabled=false:DOF 120, DOFOff:end if
end Sub
'*************************************************

'************** Slingshots ***********************

Sub RightSlingShot_Slingshot
	playsound SoundFXDOF("rightslingshot",104,DOFPulse,DOFContactors)
	DOF 106, DOFPulse
	if GI1.state=1 then
	addscore 10
	else 
	addscore 1
	end if
	if GL = "L" then 
		GL = "R" 
	Else 
		GL = "L"	
	end if
    RSling.Visible = 0
    RSling1.Visible = 1
	slingR.objroty = -15	
	RStep = 2
    RightSlingShot.TimerEnabled = 1
End Sub

Sub RightSlingShot_Timer
    Select Case RStep
        Case 3:RSLing1.Visible = 0:RSLing2.Visible = 1:slingR.objroty = -7
        Case 4:	slingR.objroty = 0:RSLing2.Visible = 0:RSLing.Visible = 1:RightSlingShot.TimerEnabled = 0:b2sSunR=b2sSunR+1
    End Select
    RStep = RStep + 1
End Sub

Sub LeftSlingShot_Slingshot
	playsound SoundFXDOF("leftslingshot",103,DOFPulse,DOFContactors)
	DOF 105, DOFPulse
	if GI01.state=1 then
	addscore 10
	else 
	addscore 1
	end if
	if GL = "L" then 
		GL = "R" 
	Else 
		GL = "L"	
	end if
    LSling.Visible = 0
    LSling1.Visible = 1
	slingL.objroty = 15	
    LStep = 2
    LeftSlingShot.TimerEnabled = 1
End Sub

Sub LeftSlingShot_Timer()
    Select Case LStep
        Case 3:LSLing1.Visible = 0:LSLing2.Visible = 1:slingL.objroty = 7
        Case 4:slingL.objroty = 0:LSLing2.Visible = 0:LSLing.Visible = 1:LeftSlingShot.TimerEnabled = 0:b2sSunL=b2sSunL+1
    End Select
    LStep = LStep + 1
End Sub

Sub LeftSlingShotb_Slingshot
	playsound SoundFX("leftslingshot",DOFcontactors)
	DOF 124, DOFPulse
	addscore 1
	if GL = "L" then 
		GL = "R" 
	Else 
		GL = "L"	
	end if
    LSlingb.Visible = 0
    LSlingb1.Visible = 1
	slingLb.objroty = 15	
    LStepb = 2
    LeftSlingShotb.TimerEnabled = 1
End Sub

Sub LeftSlingShotb_Timer()
    Select Case LStepb
        Case 3:LSLingb1.Visible = 0:LSLingb2.Visible = 1:slingLb.objroty = 7
        Case 4:slingLb.objroty = 0:LSLingb2.Visible = 0:LSLingb.Visible = 1:LeftSlingShotb.TimerEnabled = 0:b2sSunL=b2sSunL+1
    End Select
    LStepb = LStepb + 1
End Sub

Sub LeftSlingShotc_Slingshot
	playsound SoundFX("leftslingshot",DOFcontactors)
	DOF 122, DOFPulse
    addscore 5
    if motion=0 and tilt=false and c=0 then
		rt=(int(rnd(1)*8)+4)
		motion=1	
		c=1
		a=int((rnd*6)+1)
		rotort.enabled=true
	end if
	if GL = "L" then 
		GL = "R" 
	Else 
		GL = "L"	
	end if
    LSlingc.Visible = 0
    LSlingc1.Visible = 1
	slingLc.objroty = 15	
    LStepc = 2
    LeftSlingShotc.TimerEnabled = 1
End Sub

Sub LeftSlingShotc_Timer()
    Select Case LStepc
        Case 3:LSLingc1.Visible = 0:LSLingc2.Visible = 1:slingLc.objroty = 7
        Case 4:slingLc.objroty = 0:LSLingc2.Visible = 0:LSLingc.Visible = 1:LeftSlingShotc.TimerEnabled = 0:b2sSunL=b2sSunL+1
    End Select
    LStepc = LStepc + 1
End Sub

Sub LeftSlingShotd_Slingshot
	playsound SoundFX("leftslingshot",DOFContactors)
	DOF 120, DOFPulse
	addscore 1
	if GL = "L" then 
		GL = "R" 
	Else 
		GL = "L"	
	end if
    LSlingc.Visible = 0
    LSlingd1.Visible = 1
	slingLd.objroty = 15	
    LStepd = 2
    LeftSlingShotd.TimerEnabled = 1
End Sub

Sub LeftSlingShotd_Timer()
    Select Case LStepd
        Case 3:LSLingd1.Visible = 0:LSLingd2.Visible = 1:slingLd.objroty = 7
        Case 4:slingLd.objroty = 0:LSLingd2.Visible = 0:LSLingc.Visible = 1:LeftSlingShotd.TimerEnabled = 0:b2sSunL=b2sSunL+1
    End Select
    LStepd = LStepd + 1
End Sub

Sub LeftSlingShote_Slingshot
	playsound SoundFX("leftslingshot",DOFcontactors)
	DOF 126, DOFPulse
	addscore 1
	if GL = "L" then 
		GL = "R" 
	Else 
		GL = "L"	
	end if
    CentralSling.Visible = 0
    CentralSlingLe1.Visible = 1
	slingLe.objroty = 15	
    LStepe = 2
    LeftSlingShote.TimerEnabled = 1
End Sub

Sub LeftSlingShote_Timer()
    Select Case LStepe
        Case 3:CentralSlingLe1.Visible = 0:CentralSlingLe2.Visible = 1:slingLe.objroty = 7
        Case 4:slingLe.objroty = 0:CentralSlingLe2.Visible = 0:CentralSling.Visible = 1:LeftSlingShote.TimerEnabled = 0:b2sSunL=b2sSunL+1
    End Select
    LStepe = LStepe + 1
End Sub

Sub RightSlingShotb_Slingshot
	playsound SoundFX("leftslingshot",DOFContactors)
	DOF 125, DOFPulse
	addscore 1
	if GL = "L" then 
		GL = "R" 
	Else 
		GL = "L"	
	end if
    RSlingb.Visible = 0
    RSlingb1.Visible = 1
	slingRb.objroty = -15	
    RStepb = 2
    RightSlingShotb.TimerEnabled = 1
End Sub

Sub RightSlingShotb_Timer()
    Select Case RStepb
        Case 3:RSLingb1.Visible = 0:RSLingb2.Visible = 1:slingRb.objroty = 7
        Case 4:slingRb.objroty = 0:RSLingb2.Visible = 0:RSLingb.Visible = 1:RightSlingShotb.TimerEnabled = 0:b2sSunR=b2sSunR+1
    End Select
    RStepb = RStepb + 1
End Sub

Sub RightSlingShotc_Slingshot
	playsound SoundFX("leftslingshot",DOFcontactors)
	DOF 123, DOFPulse
    addscore 5
    if motion=0 and tilt=false and c=0 then
		rt=(int(rnd(1)*8)+4)
		motion=1	
		c=1
		a=int((rnd*6)+1)
		rotort.enabled=true
	end if
	if GL = "L" then 
		GL = "R" 
	Else 
		GL = "L"	
	end if
    RSlingc.Visible = 0
    RSlingc1.Visible = 1
	slingRc.objroty = -15	
    RStepc = 2
    RightSlingShotc.TimerEnabled = 1
End Sub

Sub RightSlingShotc_Timer()
    Select Case RStepc
        Case 3:RSLingc1.Visible = 0:RSLingc2.Visible = 1:slingRc.objroty = 7
        Case 4:slingRc.objroty = 0:RSLingc2.Visible = 0:RSLingc.Visible = 1:RightSlingShotc.TimerEnabled = 0:b2sSunR=b2sSunR+1
    End Select
    RStepc = RStepc + 1
End Sub

Sub RightSlingShotd_Slingshot
	playsound SoundFX("leftslingshot",DOFcontactors)
	DOF 121, DOFPulse
	addscore 1
	if GL = "L" then 
		GL = "R" 
	Else 
		GL = "L"	
	end if
    RSlingc.Visible = 0
    RSlingd1.Visible = 1
	slingRd.objroty = -15	
    RStepd = 2
    RightSlingShotd.TimerEnabled = 1
End Sub

Sub RightSlingShotd_Timer()
    Select Case RStepd
        Case 3:RSLingd1.Visible = 0:RSLingd2.Visible = 1:slingRd.objroty = 7
        Case 4:slingRd.objroty = 0:RSLingd2.Visible = 0:RSLingc.Visible = 1:RightSlingShotd.TimerEnabled = 0:b2sSunR=b2sSunR+1
    End Select
    RStepd = RStepd + 1
End Sub

Sub RightSlingShote_Slingshot
	playsound SoundFX("leftslingshot",DOFcontactors)
	DOF 126, DOFPulse
	addscore 1
	if GL = "L" then 
		GL = "R" 
	Else 
		GL = "L"	
	end if
    CentralSling.Visible = 0
    CentralSlingRe1.Visible = 1
	slingRe.objroty = 15	
    RStepe = 2
    RightSlingShote.TimerEnabled = 1
End Sub

Sub RightSlingShote_Timer()
    Select Case RStepe
        Case 3:CentralSlingRe1.Visible = 0:CentralSlingRe2.Visible = 1:slingRe.objroty = 7
        Case 4:slingRe.objroty = 0:CentralSlingRe2.Visible = 0:CentralSling.Visible = 1:RightSlingShote.TimerEnabled = 0:b2sSunR=b2sSunR+1
    End Select
    RStepe = RStepe + 1
End Sub

'*********** End Slingshots *************

'*********** Bumpers + points **********
sub bumper1_hit
	if GI10.state=1 then
		addscore 10
	else
		addscore 1
	end if
    if tilt=false then
playsound SoundFXDOF("jet1",107,DOFPulse,DOFContactors)
DOF 109, DOFPulse
    ring1.isdropped=false
    end if
b2sSunL=b2sSunL+1
end sub    

sub bumper2_hit
	if GI11.state=1 then
		addscore 10
	else
		addscore 1
	end if
    if tilt=false then
playsound SoundFXDOF("jet1",108,DOFPulse,DOFContactors)
DOF 110, DOFPulse
    ring2.isdropped=false
    end if
b2sSunR=b2sSunR+1
end sub

sub bumper3_hit
    addscore 5
    if motion=0 and tilt=false and c=0 then
    rt=(int(rnd(1)*8)+4)
    motion=1	
	c=1
	a=int((rnd*6)+1)
    rotort.enabled=true
    end if
b2sSunL=b2sSunL+1
end sub  

sub bumper4_hit
    addscore 5
    if motion=0 and tilt=false and c=0 then
    rt=(int(rnd(1)*8)+4)
    motion=1
	c=1
	a=int((rnd*6)+1)
    rotort.enabled=true
    end if
b2sSunR=b2sSunR+1
end sub
'**********************************

'****** Buttons points *****
sub ulbt_hit
    addscore 1
	bouton03.z=-1
	bouton03f.z=-1
	b2sSunL=b2sSunL+1
end sub
sub ulbt_unhit
	bouton03.z=2.5
	bouton03f.z=2.5
end sub

sub urbt_hit
    addscore 1
	bouton02.z=-1
	bouton02f.z=-1
	b2sSunR=b2sSunR+1
end sub
sub urbt_unhit
	bouton02.z=2.5
	bouton02f.z=2.5
end sub

sub lbt_hit
    addscore 1
	bouton01.z=-1
	bouton01f.z=-1
	b2sSunL=b2sSunL+1
	b2sSunR=b2sSunR+1
end sub
sub lbt_unhit
	bouton01.z=2.5
	bouton01f.z=2.5
end sub
'**************************** 

'****** Outlanes points *****
sub lout_hit
DOF 121, DOFPulse
    if loutl.state=1 then
		addscore 100
    else
		addscore 10
    end if
end sub

sub rout_hit
DOF 122, DOFPulse
    if routl.state=1 then
		addscore 100
    else
		addscore 10
    end if
end sub       
'**************************** 


'********* Triggers between targets *************
sub lru_hit

    if on1.state=1 and tilt=false and HZ=0 then
		BumpT = 1
		BumpB = 0
    end if

 if on1.state=0 and tilt=false and HZ=0 then
		BumpT = 0
		BumpB = 0
 end if

	if HZ=0 Then
	if GL = "L"  then 
		GL = "R" 
	Else 
		GL = "L"	
	end if
	end If

	HZ=1
End Sub

sub HitZoneL_unhit
HZ = 0
end Sub

sub rru_hit
    if on2.state=1 and tilt=false and HZ=0 then
		BumpT = 0
		BumpB = 1
	End If

    if on2.state=0 and tilt=false and HZ=0 then
		BumpT = 0
		BumpB = 0
    end if

	if HZ=0 Then
	if GL = "L"  then 
		GL = "R" 
	Else 
		GL = "L"	
	end if
	end If

	HZ=1
end Sub

sub HitZoneR_unhit
HZ = 0
end Sub

sub gate_hit
playsound "gate"
end Sub

sub gate1_hit
playsound "gate"
end Sub

sub gate2_hit
playsound "gate"
end Sub

'******************************************************


'******************************************************
'******                 Targets                 *******
'******************************************************

sub ctarga
    addscore 100
end sub

Sub SW24_hit():sw24p.transx = -10:Me.TimerEnabled = 1:playsound "metal4":ltarga:DOF 111, DOFPulse:End Sub
Sub SW24_Timer():sw24p.transx = 0:Me.TimerEnabled = 0:End Sub

Sub SW25_hit():sw25p.transx = -10:Me.TimerEnabled = 1:playsound "metal4":ctarga:DOF 112, DOFPulse:End Sub
Sub SW25_Timer():sw25p.transx = 0:Me.TimerEnabled = 0:End Sub

Sub SW26_hit():sw26p.transx = -10:Me.TimerEnabled = 1:playsound "metal4":rtarga:DOF 113, DOFPulse:End Sub
Sub SW26_Timer():sw26p.transx = 0:Me.TimerEnabled = 0:End Sub

sub rtarga
    if urtlb.state=0 and urtlf.state=0 then addscore 1
    if urtlb.state=1 and urtlf.state=0 then addscore 10
    if urtlf.state=1 then addscore (rsv*10)
end sub 

sub ltarga
    if ultlb.state=0 and ultlf.state=0 then addscore 1
    if ultlb.state=1 and ultlf.state=0 then addscore 10
    if ultlf.state=1 then addscore (rsv*10)
end sub                        

'*******************************************************

'******************************************************
'******         Wheel price et score            *******
'******************************************************

Sub SWheelPrice_hit():Wheel.transx = 10:Me.TimerEnabled = 1:playsound "metal4":swheelscore:DOF 114, DOFPulse:End Sub
Sub SWheelPrice_Timer():Wheel.transx = 0:Me.TimerEnabled = 0:End Sub

sub swheelscore
	If l400.state = 1 then 
		addscore 400
	end if

	If l500.state = 1 then
		addscore 500
	end if
	
    if rl.state=1 and l400.state=0 and l500.state = 0 then
		addscore (rsv*10)
    end if

    if rl.state=0 and l400.state=0 and l500.state = 0 then
		addscore (rsv) 
	end if
end Sub

'************************************************
'****               	Holes		         ****
'************************************************
sub lgob_hit
    playsound "drain"
    if lgl.state=1 then addscore (rsv*10) else addscore (rsv) end if
end sub    

sub rgob_hit
    playsound "drain"
    if rgl.state=1 then addscore (rsv*10) else addscore (rsv) end if
end sub

Sub TriggerLgob_hit:Me.DestroyBall:gobtimer.enabled=1:End Sub
Sub TriggerRgob_hit:Me.DestroyBall:gobtimer.enabled=1:End Sub

sub gobtimer_timer
    Me.enabled=0
    playsound "motorshort"
    activeballs=activeballs-1
	if activeballs=0 then tt=0:nextball
	if B2SOn then Scoreb2s.enabled =1:end If
end sub

'************************************************
'****               wheel Price              ****
'************************************************
sub rotoradv
	If wheel.objroty = 0 then rsv =10
	If wheel.objroty = 36 then rsv =20
	If wheel.objroty = 72 then rsv =30
	If wheel.objroty = 108 then rsv =40
	If wheel.objroty = 144 then rsv =50
	If wheel.objroty = 180 then rsv =10
	If wheel.objroty = 216 then rsv =20
	If wheel.objroty = 252 then rsv =30
	If wheel.objroty = 288 then rsv =40
	If wheel.objroty = 324 then rsv =50
    rv=rv+1
    if rv>9 then rv=0:end If
    playsound "lightmotor"
end sub    

'*************************************************
'*********          Scoring lights        ********
sub ylupdate
    for each obj in rtvalue
    obj.state=0
    next
    rtvalue(ylv).state=1
'    if ylv=0 or ylv=3 and rv>5 then
'		if rsv=40 then l400.state=1:l500.state=0:l400z.state=1:l500z.state=0:end if
'		if rsv=50 then l500.state=1:l400.state=0:l500z.state=1:l400z.state=0:end if
'    end if
'    if rsv<40 then
'    l400.state=0:l400z.state=0
'    l500.state=0:l500z.state=0
'    end if
end sub    
'*************************************************

sub credtimer_timer 
    cred=cred+1
    if cred=1 then credtxt.text="Ported To VPX By JPJ, from a Leon Spalding's Table" 
    if cred=2 then credtxt.text="'5'Coin:'S'Start:'Right magnet or automatic 'Load Ball"
    if cred=3 then credtxt.text="Thanx Go To:"
    if cred=4 then credtxt.text="Randy Davis & Black (VP GODS)"
    if cred=5 then credtxt.text="Most At VPF, VPFF & MFF"
    if cred=6 then credtxt.text="'5'Coin:'S'Start:'Right magnet or automatic 'Load Ball"
    if cred=6 then cred=0
end sub

'****************** Save and load values, score, credit *********************
sub savehs
    savevalue "Contest", "credit", credit
    savevalue "Contest", "plsc1", score(0)
    savevalue "Contest", "plsc2", score(1)
    savevalue "Contest", "plsc3", score(2)
    savevalue "Contest", "plsc4", score(3)
    savevalue "Contest", "hiscore", hisc
    savevalue "Contest", "match", matchnumb
    savevalue "Contest", "rotpos", rv
    savevalue "Contest", "yellowlpos", ylv
end sub

sub loadhs
    dim temp
    temp = LoadValue("Contest", "credit")
    If (temp <> "") then credit = CDbl(temp)
    temp = LoadValue("Contest", "plsc1")
    If (temp <> "") then score(0) = CDbl(temp)
    temp = LoadValue("Contest", "plsc2")
    If (temp <> "") then score(1) = CDbl(temp)
    temp = LoadValue("Contest", "plsc3")
    If (temp <> "") then score(2) = CDbl(temp)
    temp = LoadValue("Contest", "plsc4")
    If (temp <> "") then score(3) = CDbl(temp)
    temp = LoadValue("Contest", "hiscore")
    If (temp <> "") then hisc = CDbl(temp)
    temp = LoadValue("Contest", "match")
    If (temp <> "") then matchnumb = CDbl(temp)
    temp = LoadValue("Contest", "rotpos")
    If (temp <> "") then rv = CDbl(temp)
    temp = LoadValue("Contest", "yellowlpos")
    If (temp <> "") then ylv = CDbl(temp)
If ylv>4 then ylv=0
end sub

'********* launching ball triger **********
sub ballhome_hit
    ballrelenabled=1
end sub

sub ballrel_hit
    if ballrelenabled=1 then
    playsound "launchball"
    ballrelenabled=0
    end if
end sub

'plus utile
'sub ballout_hit
'    bb.isdropped=false
'end sub

'******* Light to the B2S *******
Sub BGLight(id, value)
	If B2SOn Then
        Controller.B2SSetData id, value
    End If
End Sub

'******* Score, credit lights to B2s *******

sub Scoreb2s_timer()

'********** Scores ************
if B2SOn Then
Controller.B2SSetScore 1, Score(0)
Controller.B2SSetScore 2, Score(1)
Controller.B2SSetScore 3, Score(2)
Controller.B2SSetScore 4, Score(3)
Controller.B2SSetCredits credit
Controller.B2SSetPlayerUp playno

'********** active PLayer's light in game ************
if currpl=0 Then
	if ballinplay>0 and ballinplay<6 then 
		Controller.B2SSetData 25, 1 
	end if
	else 
		Controller.B2SSetData 25, 0 
end If
if currpl=1 Then
	if ballinplay>0 and ballinplay<6 then 
		Controller.B2SSetData 26, 1 
	end if
	else 
		Controller.B2SSetData 26, 0 
end If
if currpl=2 Then
	if ballinplay>0 and ballinplay<6 then 
		Controller.B2SSetData 27, 1 
	end if
	else 
		Controller.B2SSetData 27, 0 
end If
if currpl=3 Then
	if ballinplay>0 and ballinplay<6 then 
		Controller.B2SSetData 28, 1 
	end if
	else 
		Controller.B2SSetData 28, 0 
end If

'********** Tilt ************
Controller.B2SSetBallInPlay ballinplay
Controller.B2SSetTilt 33, pltilt(0)
Controller.B2SSetTilt 33, pltilt(1)
Controller.B2SSetTilt 33, pltilt(2)
Controller.B2SSetTilt 33, pltilt(3)
End If
'********** Game Over ************
If ballinplay<1 or ballinplay>5 and b2svarlightA=0 then 
	If B2SOn Then Controller.B2SSetData 35, 1: Controller.B2SSetPlayerUp 0
	b2svar1.enabled=1
	b2svar2.enabled=0
	sunzero
	sunL.enabled=0
	sunR.enabled=0
Else
	If B2SOn Then
		Controller.B2SSetData 35, 0
		Controller.B2SSetData 10, 0
		Controller.B2SSetData 11, 0	
	End If
	b2svarlightA=0
	b2svar1.enabled=0
	b2svar2.enabled=1
	sunL.enabled=1
	sunR.enabled=1
end If
If ballinplay<1 or ballinplay>5 and b2svarlightA<>0 then 
	If B2SOn Then Controller.B2SSetData 35, 1
	b2svar1.enabled=1
	b2svar2.enabled=0
	sunzero
	sunL.enabled=0
	sunR.enabled=0
Else
	If B2SOn Then 
		Controller.B2SSetData 35, 0
		Controller.B2SSetData 10, 0
		Controller.B2SSetData 11, 0	
	End If
	b2svarlightA=0
	b2svar1.enabled=0
	b2svar2.enabled=1
	sunL.enabled=1
	sunR.enabled=1
end If
end Sub


sub B2svar1_timer
if b2svarlightA=0 then b2svar1.enabled = 0:end If
	If B2SOn Then
	select case b2svarlightA
		case 1:Controller.B2SSetData 10, 0
		Controller.B2SSetData 11, 0	
		case 2:Controller.B2SSetData 10, 1
		Controller.B2SSetData 11, 1	
		case 3:
		Controller.B2SSetData 10, 1
		Controller.B2SSetData 11, 0	
		case 4:
		Controller.B2SSetData 10, 0
		Controller.B2SSetData 11, 1	
		case 5:
		Controller.B2SSetData 10, 0
		Controller.B2SSetData 11, 0	
		case 6:
		Controller.B2SSetData 10, 1
		Controller.B2SSetData 11, 1	
	end Select
	End If

	b2svarlightA=b2svarlightA+1
	if b2svarlightA=7 then b2svarlightA=1
end sub

sub b2svar2_timer
if b2svarlightB=0 then b2svar2.enabled = 0:end If
	If B2SOn Then
	select case b2svarlightB
		case 1:
		Controller.B2SSetData 8, 1
		Controller.B2SSetData 9, 0	
		case 2:
		Controller.B2SSetData 8, 1
		Controller.B2SSetData 9, 0	
		case 3:
		Controller.B2SSetData 8, 0
		Controller.B2SSetData 9, 1	
		case 4:
		Controller.B2SSetData 8, 0
		Controller.B2SSetData 9, 1	
		case 5:
		Controller.B2SSetData 8, 0
		Controller.B2SSetData 9, 0	
		case 6:
		Controller.B2SSetData 8, 1
		Controller.B2SSetData 9, 1
		case 7:
		Controller.B2SSetData 8, 0
		Controller.B2SSetData 9, 0	
		case 8:
		Controller.B2SSetData 8, 1
		Controller.B2SSetData 9, 1	
		case 9:
		Controller.B2SSetData 8, 0
		Controller.B2SSetData 9, 0	
		case 10:
		Controller.B2SSetData 8, 0
		Controller.B2SSetData 9, 1	
		case 11:
		Controller.B2SSetData 8, 1
		Controller.B2SSetData 9, 0	
		case 12:
		Controller.B2SSetData 8, 0
		Controller.B2SSetData 9, 0	
	end Select
	End If
	b2svarlightB=b2svarlightB+1
	if b2svarlightB=10 then b2svarlightB=1
End Sub

sub sunL_Timer
	If B2SOn Then
	select case b2sSunL
		case 1:
		Controller.B2SSetData 40, 1
		Controller.B2SSetData 41, 0	
		Controller.B2SSetData 42, 0	
		Controller.B2SSetData 43, 0	
		case 2:
		Controller.B2SSetData 40, 0
		Controller.B2SSetData 41, 1	
		Controller.B2SSetData 42, 0	
		Controller.B2SSetData 43, 0	
		case 3:
		Controller.B2SSetData 40, 0
		Controller.B2SSetData 41, 0	
		Controller.B2SSetData 42, 1	
		Controller.B2SSetData 43, 0	
		case 4:
		Controller.B2SSetData 40, 0
		Controller.B2SSetData 41, 0	
		Controller.B2SSetData 42, 0	
		Controller.B2SSetData 43, 1	
		case 5:
		Controller.B2SSetData 40, 0
		Controller.B2SSetData 41, 0	
		Controller.B2SSetData 42, 1	
		Controller.B2SSetData 43, 0	
		case 6:
		Controller.B2SSetData 40, 0
		Controller.B2SSetData 41, 1	
		Controller.B2SSetData 42, 0	
		Controller.B2SSetData 43, 0	
	end Select
	End If
'	b2sSunL=b2sSunL+1
	if b2sSunL=7 then b2sSunL=1:end If

end Sub

Sub sunR_Timer
	If B2SOn Then
	select case b2sSunR
		case 1:
		Controller.B2SSetData 44, 0
		Controller.B2SSetData 45, 0	
		Controller.B2SSetData 46, 0	
		Controller.B2SSetData 47, 1	
		case 2:
		Controller.B2SSetData 44, 0
		Controller.B2SSetData 45, 0	
		Controller.B2SSetData 46, 1	
		Controller.B2SSetData 47, 0	
		case 3:
		Controller.B2SSetData 44, 0
		Controller.B2SSetData 45, 1	
		Controller.B2SSetData 46, 0	
		Controller.B2SSetData 47, 0	
		case 4:
		Controller.B2SSetData 44, 1
		Controller.B2SSetData 45, 0	
		Controller.B2SSetData 46, 0	
		Controller.B2SSetData 47, 0	
		case 5:
		Controller.B2SSetData 44, 0
		Controller.B2SSetData 45, 1	
		Controller.B2SSetData 46, 0	
		Controller.B2SSetData 47, 0	
		case 6:
		Controller.B2SSetData 44, 0
		Controller.B2SSetData 45, 0	
		Controller.B2SSetData 46, 1	
		Controller.B2SSetData 47, 0	
	end Select
	End If
'	b2sSunR=b2sSunR+1
	if b2sSunR=7 then b2sSunR=1:end If
end Sub

Sub sunzero
	If B2SOn Then
		Controller.B2SSetData 40, 0
		Controller.B2SSetData 41, 0
		Controller.B2SSetData 42, 0	
		Controller.B2SSetData 43, 0	
		Controller.B2SSetData 44, 0
		Controller.B2SSetData 45, 0	
		Controller.B2SSetData 46, 0	
		Controller.B2SSetData 47, 0	
	End If
end Sub


'******* Lights Timers, Lights mechanic *******

sub GILeft_Timer
if GL="R" then GIRight:end If
if GL="L" then GILeftt:end If

if ylv=0 or ylv=3 and rv>5 then
		if rsv=40 then l400.state=1:l500.state=0:l400z.state=1:l500z.state=0:end if
		if rsv=50 then l500.state=1:l400.state=0:l500z.state=1:l400z.state=0:end if
    end if
    if rsv<40 then
    l400.state=0:l400z.state=0
    l500.state=0:l500z.state=0
    end if

End Sub

sub GIGen_Timer
for each obj in GIR
	obj.state=0
Next
for each obj in GIL
	obj.state=0
Next
'condition hasard
GL="L"
end Sub

sub GILeftt
for each obj in GIR
	obj.state=0
Next
for each obj in GIL
	obj.state=1
Next
'********** bottom clouds in B2S ************
If B2SOn Then
if ballinplay>0 and ballinplay<6 then
	Controller.B2SSetData 12, 1
	Controller.B2SSetData 14, 0
	Else
	Controller.B2SSetData 12, 0
	Controller.B2SSetData 14, 0
End If
End If
'********************************************

'Targets Top
GI01.state = 1
GI02.state = 1
GI1.state = 0
GI2.state = 0

'on/off
on1.state=1
off2.state=1
on2.state=0
off1.state=0

'Right Outlane
routl.state = 1
loutl.state = 0

'Bumpers
if bumpB = 1 or bumpT = 1 Then
	GI11.state = 1
	Light7.state = 1
	GI10.state = 0
	Light6.state = 0
end If
if bumpB = 0 and bumpT = 0 Then
	GI11.state = 0
	Light7.state = 0
	GI10.state =0
	Light6.state = 0
end if
end Sub

sub GIRight
for each obj in GIR
	obj.state=1
Next
for each obj in GIL
	obj.state=0
Next
'********** bottom clouds in B2S ************
If B2SOn Then
if ballinplay>0 and ballinplay<6 then 
	Controller.B2SSetData 12, 0
	Controller.B2SSetData 14, 1
	Else
	Controller.B2SSetData 12, 0
	Controller.B2SSetData 14, 0
End If
End If
'********************************************

'condition
'Targets Top
GI01.state = 0
GI02.state = 0
GI1.state = 1
GI2.state = 1

'on/off
on1.state=0
off2.state=0
on2.state=1
off1.state=1

'Left Outlane
loutl.state = 1
routl.state = 0

'Bumpers
if bumpB = 1 or bumpT = 1 Then
	GI11.state = 0
	Light7.state = 0
	GI10.state =1
	Light6.state = 1
end If
if bumpB = 0 and bumpT = 0 Then
	GI11.state = 0
	Light7.state = 0
	GI10.state =0
	Light6.state = 0
end If
end Sub

' *********************************************************************
' 						Other Sound FX
' *********************************************************************

'Rubbers Collision
Sub Rubbers_Hit(idx)
 	dim finalspeed
  	finalspeed=SQR(activeball.velx * activeball.velx + activeball.vely * activeball.vely)
 	If finalspeed > 20 then 
		PlaySound "rubber_hit_1", 0, Vol(ActiveBall)*10*GlobalSoundLevel , pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0, AudioFade(ActiveBall)
	End if
	If finalspeed >= 1 AND finalspeed <= 20 then
 		RandomSoundRubber()
 	End If
End Sub

Sub RandomSoundRubber()
	Select Case Int(Rnd*3)+1
		Case 1 : PlaySound "rubber_hit_1", 0, Vol(ActiveBall)*10*GlobalSoundLevel , pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0, AudioFade(ActiveBall)
		Case 2 : PlaySound "rubber_hit_2", 0, Vol(ActiveBall)*10*GlobalSoundLevel , pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0, AudioFade(ActiveBall)
		Case 3 : PlaySound "rubber_hit_3", 0, Vol(ActiveBall)*10*GlobalSoundLevel , pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0, AudioFade(ActiveBall)
	End Select
End Sub

'Metals Collision
Sub MetalsThin_Hit (idx)
	PlaySound "metalhit_thin", 0, Vol(ActiveBall)*GlobalSoundLevel, Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
End Sub

'Post Collision
Sub wood_hit(idx)
PlaySound "Post1", 0, Vol(ActiveBall)*10*GlobalSoundLevel , pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0, AudioFade(ActiveBall)
End Sub

'Flippers Collision
Sub LeftFlipper_Collide(parm) : RandomSoundFlipper() : End Sub
Sub RightFlipper_Collide(parm) : RandomSoundFlipper() : End Sub

Sub RandomSoundFlipper()
	Select Case Int(Rnd*3)+1
		Case 1 :PlaySound "flip_hit_1", 0, GlobalSoundLevel * ballvel(ActiveBall) / 50, -0.1, 0.25
		Case 2 :PlaySound "flip_hit_2", 0, GlobalSoundLevel * ballvel(ActiveBall) / 50, -0.1, 0.25
		Case 3 :PlaySound "flip_hit_3", 0, GlobalSoundLevel * ballvel(ActiveBall) / 50, -0.1, 0.25
	End Select
End Sub

'******************************************************
'       	RealTime Updates
'******************************************************
Const tnob = 1 ' total number of balls for this table (at the same time)

Sub GameTimer_timer()
	RollingSoundUpdate
	BallShadowUpdate
	FlippersUpdate
	If ReflectionMod Then RMUpdate
End Sub


'*********** BALL SHADOW *********************************
ReDim BallShadow(tnob-1)
InitBallShadow

Sub InitBallShadow
	Dim i:For i = 0 to tnob-1
		ExecuteGlobal "Set BallShadow(" & i & ") = BallShadow" & (i+1) & " :"
	Next
End Sub

Sub BallShadowUpdate
    Dim BOT, b
    BOT = GetBalls
	' hide shadow of deleted balls
	If UBound(BOT)<(tnob-1) Then
		For b = (UBound(BOT) + 1) to (tnob-1)
			BallShadow(b).visible = 0
		Next
	End If
	' exit the Sub if no balls on the table
    If UBound(BOT) = -1 Then Exit Sub
	' render the shadow for each ball
    For b = 0 to UBound(BOT)
		If BOT(b).X < Table1.Width/2 Then
			BallShadow(b).X = ((BOT(b).X) - (50/6) + ((BOT(b).X - (Table1.Width/2))/7)) + 10
		Else
			BallShadow(b).X = ((BOT(b).X) + (50/6) + ((BOT(b).X - (Table1.Width/2))/7)) - 10
		End If
	    ballShadow(b).Y = BOT(b).Y + 20
		If BOT(b).Z > 20 Then
			BallShadow(b).visible = 1
		Else
			BallShadow(b).visible = 0
		End If
	Next
End Sub

'*********** FLIPPERS MODEL AND SHADOW SYNC *********************************
Sub FlippersUpdate
'**     Include Flippers Primitives    **
LeftFlipperP.ObjRotZ = LeftFlipper.CurrentAngle-130
Primitive85.Roty = LeftFlipper.CurrentAngle-130
Primitive87.Roty = LeftFlipper.CurrentAngle-130
LeftFlipperP1.ObjRotZ = Flipper1.CurrentAngle-125
RightFlipperP.ObjRotZ = RightFlipper.CurrentAngle-230
Primitive86.Roty = RightFlipper.CurrentAngle-230
Primitive88.Roty = RightFlipper.CurrentAngle-230
RightFlipperP1.ObjRotZ = Flipper2.CurrentAngle-235

'**     Include Flipper's shadows from Ninuzzu    **
LeftFlipperSh.RotZ = LeftFlipper.currentangle
RightFlipperSh.RotZ = RightFlipper.currentangle
LeftFlipperSh1.RotZ = Flipper1.currentangle
RightFlipperSh1.RotZ = Flipper2.currentangle
End Sub

'******* JPJ reflection mod  *******
Sub RMUpdate
if loutl.state = 1 then loutlz.state = 1 else loutlz.state = 0:end If
if routl.state = 1 then routlz.state = 1 else routlz.state = 0:end If
if lgl.state = 1 then
		lgl2.state = 1
		lglz.state = 1
	else 
		lglz.state = 0
		lgl2.state = 0
end If
if rl.state = 1 then rlz.state = 1 else rlz.state = 0:end If
if rgl.state = 1 then 
		rgl2.state = 1
		rglz.state = 1
	else 
		rglz.state = 0
		rgl2.state = 0
end If
if ultlf.state = 1 then ultlfz.state = 1 else ultlfz.state = 0:end If
if ultlb.state = 1 then ultlbz.state = 1 else ultlbz.state = 0:end If
if on1.state = 1 then on1z.state = 1 else on1z.state = 0:end If
if off1.state = 1 then off1z.state = 1 else off1z.state = 0:end If
if on2.state = 1 then on2z.state = 1 else on2z.state = 0:end If
if off2.state = 1 then off2z.state = 1 else off2z.state = 0:end If
if urtlb.state = 1 then urtlbz.state = 1 else urtlbz.state = 0:end If
if urtlf.state = 1 then urtlfz.state = 1 else urtlfz.state = 0:end If
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

'*********** ROLLING SOUND *********************************

ReDim rolling(tnob)
InitRolling
Sub InitRolling:Dim i:For i=0 to (tnob-1):rolling(i) = False:Next:End Sub

Sub RollingSoundUpdate
    Dim BOT, b
    BOT = GetBalls
	' stop the sound of deleted balls
	If UBound(BOT)<(tnob-1) Then
		For b = (UBound(BOT) + 1) to (tnob-1)
			rolling(b) = False
			StopSound("fx_ballrolling" & (b+1))
		Next
	End If
	' exit the Sub if no balls on the table
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

' Thalamus : Exit in a clean and proper way
Sub table1_exit()
  Controller.Pause = False
  Controller.Stop
End Sub


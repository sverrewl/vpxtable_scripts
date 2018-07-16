'**************************************************
'  ____    _    ____  _                        
' / ___|  / \  |  _ \| |_ ___   ___  _ __  ___ 
'| |     / _ \ | |_) | __/ _ \ / _ \| '_ \/ __|
'| |___ / ___ \|  _ <| || (_) | (_) | | | \__ \
' \____/_/   \_\_| \_\\__\___/ \___/|_| |_|___/
'                                              
' ___             _       ___             _    
''| _ \___ __ _ __| |_    / __|___ __ _ __| |_  
''|   / _ / _` / _| ' \  | (__/ _ / _` / _| ' \ 
''|_|_\___\__,_\__|_||_|  \___\___\__,_\__|_||_|
'
' >>>>>>>>>>>>started 2016 - 2017 released<<<<<<<<<<
'***************************************************
' Created by: HauntFreaks - BorgDog - SliderPoint
' ------------------------------------------------
' Roach Effect Credits: Dark / Randr / Toxie / Fuzzel / HauntFreaks / BorgDog
' ----------------------------------------------------------------------------
' Graphics/Lighting: HauntFreaks - SliderPoint - BorgDog
' -------------------------------------------------------
' Code/Animation: BorgDog - SliderPoint
'*******************************************************

Option Explicit
Randomize

'**********************************************************
'********   	OPTIONS		*******************************
'**********************************************************

Dim BallShadows: Ballshadows=1  		'******************	set to 1 to turn on Ball shadows
Dim FlipperShadows: FlipperShadows=1  	'***********		set to 1 to turn on Flipper shadows


Const cGameName = "Cartoons"

Dim operatormenu, options
Dim bumperlitscore
Dim bumperoffscore
Dim balls
Dim replays
Dim Replay1Table(3), Replay2Table(3), Replay3Table(3)
Dim Add10, Add100, Add1000
dim replay1, replay2, replay3
dim replaytext(3)
Dim hisc, hiscstate, ebmode, extraballs, Rhisc, hircstate
Dim players
Dim player
Dim credit, freeplay
Dim score(2), RoachScore(2)
Dim sreels(2), Rreels(2)
Dim p100k(2)
Dim state
Dim tilt
Dim tiltsens
Dim target(8)
Dim StarState
Dim Bonus
Dim Bonuslight(19)
Dim ballinplay
Dim matchnumb
dim ballrenabled, PlungeBall
Dim rep(2)
Dim rst
Dim eg, goroaches
Dim scn
Dim scn1
Dim bell, saahc
Dim i,j,tnumber, light, objekt
Dim awardcount, dbonus
dim roachstep, roachxy1, roachxy2, roachxy3, roachxy4
dim RoachBall, MirrorBall, RealBall, roachran, RoachLoc
dim frameRate:frameRate=0.8
dim frame:frame=0	
Dim PI:PI=3.1415926
Dim Magnet0, Magnet1, magnet2, Magnet3, Magnet4


On Error Resume Next
ExecuteGlobal GetTextFile("controller.vbs")
If Err Then MsgBox "Can't open controller.vbs"
On Error Goto 0

On Error Resume Next
ExecuteGlobal GetTextFile("core.vbs")
If Err Then MsgBox "You need the core.vbs in order to run this table, available in the vp10 package"
On Error Goto 0

sub CARtoons_init
	LoadEM
	Replay1Table(1)=90000
	Replay1Table(2)=120000
	Replay1Table(3)=140000
	Replay2Table(1)=120000
	Replay2Table(2)=150000
	Replay2Table(3)=170000
	Replay3Table(1)=150000
	Replay3Table(2)=180000
	Replay3Table(3)=200000
	replaytext(1)="90000,120000,150000"
	replaytext(2)="120000,150000,180000"
	replaytext(3)="140000,170000,200000"
	set sreels(1) = ScoreReel1
	set sreels(2) = ScoreReel2
	set Rreels(1) = roachReel1
	set Rreels(2) = roachReel2
	set p100k(1) = P1100k
	set p100k(2) = P2100k
	set target(1)=DTredL
	set target(2)=DTredR
	set target(3)=DTwhiteL
	set target(4)=DTwhiteR
	set target(5)=DTgreenL
	set target(6)=DTgreenR
	set target(7)=DTyellowL
	set target(8)=DTyellowR
	set bonuslight(1)=bonus1
	set bonuslight(2)=bonus2
	set bonuslight(3)=bonus3
	set bonuslight(4)=bonus4
	set bonuslight(5)=bonus5
	set bonuslight(6)=bonus6
	set bonuslight(7)=bonus7
	set bonuslight(8)=bonus8
	set bonuslight(9)=bonus9
	set bonuslight(10)=bonus10
	set bonuslight(11)=bonus11
	set bonuslight(12)=bonus12
	set bonuslight(13)=bonus13
	set bonuslight(14)=bonus14
	set bonuslight(15)=bonus15
	set bonuslight(16)=bonus16
	set bonuslight(17)=bonus17
	set bonuslight(18)=bonus18
	set bonuslight(19)=bonus19
	roachxy1=Array(Array(195,200,215,245,285,335,390,450,500,550,600,640,660,640,610,555,500,460,435,425,430,445,460,475),Array(1430,1485,1545,1590,1625,1650,1665,1660,1655,1650,1620,1575,1520,1465,1425,1400,1410,1440,1495,1550,1605,1660,1715,1775))
	roachxy2=Array(Array(640,590,540,485,425,380,335,315,310,320,340,370,400,420,440,450,450,450,450,450,450,450,450,450),Array(1310,1275,1250,1240,1250,1265,1305,1355,1415,1470,1525,1570,1615,1665,1715,1765,1765,1765,1765,1765,1765,1765,1765,1765))
	roachxy3=Array(Array(110,165,215,270,330,380,430,480,515,540,555,560,555,540,525,505,495,485,475,475,480,480,480,480),Array(1020,995,980,970,975,990,1015,1045,1090,1135,1190,1250,1300,1355,1405,1455,1510,1565,1620,1670,1725,1775,1775,1775))
	roachxy4=Array(Array(705,650,600,540,490,440,400,370,340,330,320,330,335,350,375,390,415,445,470,490,495,495,500,500),Array(900,880,875,880,895,925,965,1010,1060,1115,1170,1225,1285,1335,1385,1440,1490,1535,1585,1640,1695,1750,1800,1800))
	hideoptions
	player=1
	For each light in LaneLights:light.State = 0: Next
	For each light in BumperLights:light.State = 0: Next
	For each light in Tlights:light.State = 0: Next
	For each light in PFlights:light.State = 0: Next
	TlightRL1.State = 0:light4.state = 0 'mike add
	RoachLoc=0
	
	ebmode=0
	goroaches=0
	hisc=50000
	Rhisc=5

	HSA1=4
	HSA2=15
	HSA3=7
	RSA1=18
	RSA2=1
	RSA3=4
	loadhs

	UpdatePostIt
	UpdatePostIt2
	gamov.text="Game Over"
	gamov.timerenabled=1
	tilttxt.timerenabled=1
    if credit="" then credit=0
	creditreel.setvalue(credit)
	if balls="" then balls=3
	if balls<>3 and balls<>5 then balls=5
	if freeplay="" or freeplay<0 or freeplay>1 then freeplay=0
	if matchnumb<0 or matchnumb>9 then matchnumb=0
	if replays="" then replays=2
	if replays<>1 and replays<>2 and replays<>3 then replays=2
	Replay1=Replay1Table(Replays)
	Replay2=Replay2Table(Replays)
	Replay3=Replay3Table(Replays)
	OptionBalls.image="OptionsBalls"&Balls
	OptionReplays.image="OptionsReplays"&replays
	OptionEBmode.image="OptionsEBmode"&ebmode
	if ebmode=0 then
		RepCard.image = "ReplyCard"&replays
	  else
		RepCard.image = "EBCard"&replays
	end if
	OptionFreeplay.image="OptionsFreeplay"&freeplay
	
	if balls=3 then
		bumperlitscore=1000
		bumperoffscore=1000
		InstCard.image="InstCard3balls"
	end if
	if balls=5 then
		bumperlitscore=100
		bumperoffscore=100
		InstCard.image="InstCard5balls"
	end if

	if ballshadows=1 then
        BallShadowUpdate.enabled=1
      else
        BallShadowUpdate.enabled=0
    end if
 
    if flippershadows=1 then
        FlipperLSh.visible=1
        FlipperRSh.visible=1
       else
        FlipperLSh.visible=0
        FlipperRSh.visible=0
     end if

	If B2SOn then
		for each objekt in backdropstuff : objekt.visible = 0 : next
	End If

	if matchnumb=0 then
		matchtxt.text="00"
	  else
		matchtxt.text=matchnumb*10
	end if

    scorereel1.setvalue(score(1) MOD 100000)
	roachreel1.setvalue(RoachScore(1))
	if score(1)>100000 then 
		p1100k.text=100000*Int(score(1)/100000)
		If B2SOn then Controller.B2SSetScoreRollover 24 + 1, Int(score(1)/100000) 
	  else
		p1100k.text=" "
	end if
    scorereel2.setvalue(score(2) MOD 100000)
	roachreel2.setvalue(RoachScore(2))
	if score(2)>100000 then 
		p2100k.text=100000*Int(score(2)/100000)
		If B2SOn then Controller.B2SSetScoreRollover 24 + 2, Int(score(2)/100000)
	  else
		p2100k.text=" "
	end if

	tilt=false
	IF Credit > 0 Then DOF 122, DOFOn
	saahc=1
	set RoachBall = EVAL("RoachHole"&RoachLoc).CreateSizedBall(20)
	RoachBall.visible=0
	set MirrorBall = MirrorDrain.CreateBall
	MirrorBall.visible=0
	set RealBall = Drain.CreateBall
	startGame.enabled=true

	Set Magnet0=New cvpmMagnet
	with Magnet0
		.InitMagnet RoachTrigger0,3
		.CreateEvents "Magnet0"
		.AttractBall RoachBall
		.Grabcenter = 1
		.magneton = 0
	end With
	Set Magnet1=New cvpmMagnet
	with Magnet1
		.InitMagnet RoachTrigger1,3
		.CreateEvents "Magnet1"
		.AttractBall RoachBall
		.Grabcenter = 1
		.magneton = 0
	end With
	Set magnet2=New cvpmMagnet
	with magnet2
		.InitMagnet RoachTrigger2,3
		.CreateEvents "magnet2"
		.AttractBall RoachBall
		.Grabcenter = 1
		.magneton = 0
	end With
	Set Magnet3=New cvpmMagnet
	with Magnet3
		.InitMagnet RoachTrigger3,3
		.CreateEvents "Magnet3"
		.AttractBall RoachBall
		.Grabcenter = 1
		.magneton = 0
	end With
	Set Magnet4=New cvpmMagnet
	with Magnet4
		.InitMagnet RoachTrigger4,3
		.CreateEvents "Magnet4"
		.AttractBall RoachBall
		.Grabcenter = 1
		.magneton = 0
	end With
End sub

sub startGame_timer
	playsound "poweron"
	if B2SOn then
		Controller.B2ssetCredits Credit
		Controller.B2ssetMatch 34, Matchnumb
		Controller.B2SSetGameOver 35,1
		Controller.B2SSetScorePlayer 1, Score(1) MOD 100000
		Controller.B2SSetScorePlayer 2, Score(2) MOD 100000
		Controller.B2SSetScorePlayer 3, RoachScore(1) 
		Controller.B2SSetScorePlayer 4, RoachScore(2) 
	end if
'	lightdelay.enabled=true
	me.enabled=false
end sub

sub lightdelay_timer
	for each light in tlights:light.state=1:next
	initdelay.enabled=1
	me.enabled=0
end Sub

sub initdelay_timer
	saahc=4
	lacucaracha.enabled=1
	if goroaches=0 then 
		roachstep=0
		roachinit.enabled=1
		AnimLoopTimer.enabled=1
		goroaches=1
	end if
	me.enabled=0
end sub

Sub roachtimer_timer
	Dim roachA,dx,dy
	mirrorBall.x = RealBall.x
	mirrorBall.y = RealBall.y
	mirrorBall.z = 359.4-(.06993*(1927-RealBall.y))

	roach.x = roachBall.x
	roach.y = roachBall.y


		'change roach rotation
			dy=-1*roachBall.vely  '-1*(EVAL("roachxy"&xx)(1)(roachstep)-EVAL("roachxy"&xx)(1)(roachstep-1))	'delta Y
			if dy=0 then dy=0.0000001													'avoid divide by zero errors
			dx=RoachBall.velx 	'EVAL("roachxy"&xx)(0)(roachstep)-EVAL("roachxy"&xx)(0)(roachstep-1)		'delta X
			roachA=atn(dX/dY)												'angle in radians
			if roachA<0 then roachA=roachA+PI								'correction for negative angles
			roachA=int(roachA/(pi/180))							'convert to degrees
			if dx<0 then roachA=roachA+180						'correction for negative quadrants
			Roach.roty=roachA
end sub

sub roachmove
	if roachtimer.enabled=0 then
		RoachScore(player)=RoachScore(player)+1
		Rreels(player).setvalue RoachScore(player)
		if B2SOn then Controller.B2SSetScorePlayer (player+2), RoachScore(Player) 
		RoachMagnetsOff
		Do
			roachran=int(rnd*5)
		Loop while roachran=roachloc

		select case roachran
			case 0:
				magnet0.magneton=true
			case 1:
				Magnet1.magneton=true 
			case 2:
				magnet2.magneton=true 
			case 3:
				magnet3.magneton=true 
			case 4:
				magnet4.magneton=true 
		end Select
		roachtimer.enabled=1
		RoachAnimLoopTimer.enabled=1
		select case RoachLoc
			case 0:
				RoachHole0.kick 180,4,0 
			case 1:
				RoachHole1.kick 315,4,0 
			case 2:
				RoachHole2.kick 45,4,0 
			case 3:
				RoachHole3.kick 95,4,0 
			case 4:
				RoachHole4.kick 290,4,0 
		end Select
	end if
end sub

sub RoachAnimLoopTimer_Timer
	Roach.ShowFrame frame
	frame = frame + frameRate
	If frame>18 Then	' the animation is a bit rough so don't play it till the end but stop a bit earlier
		frame=0
	end if
end sub

Sub RoachHole0_hit
	RoachAnimLoopTimer.enabled=0
'	Magnet0.magneton=False
	RoachMagnetsOff
	roachtimer.enabled=0
	roachloc=0
End Sub

Sub RoachHole1_hit
	RoachAnimLoopTimer.enabled=0
'	Magnet1.magneton=False
	RoachMagnetsOff
	roachtimer.enabled=0
	roachloc=1
End Sub

Sub RoachHole2_hit
	RoachAnimLoopTimer.enabled=0
'	magnet2.magneton=False
	RoachMagnetsOff
	roachtimer.enabled=0
	roachloc=2
End Sub

Sub RoachHole3_hit
	RoachAnimLoopTimer.enabled=0
'	Magnet3.magneton=False
	RoachMagnetsOff
	roachtimer.enabled=0
	roachloc=3
End Sub

Sub RoachHole4_hit
	RoachAnimLoopTimer.enabled=0
'	Magnet4.magneton=False
	RoachMagnetsOff
	roachtimer.enabled=0
	roachloc=4
End Sub

Sub RoachMagnetsOff
	Magnet0.magneton=False
	Magnet1.magneton=False
	magnet2.magneton=False
	Magnet3.magneton=False
	Magnet4.magneton=False
End Sub

sub lacucaracha_timer()
	me.interval=100
    Select Case saahc
	    Case 4: PlaySoundAt SoundFXDOF("bell1000",133,DOFPulse,DOFChimes), TrigRout
		Case 5: PlaySoundAt SoundFXDOF("bell1000",133,DOFPulse,DOFChimes), TrigRout
		case 6: PlaySoundAt SoundFXDOF("bell1000",133,DOFPulse,DOFChimes), TrigRout
		case 7: PlaySoundAt SoundFXDOF("bell100",132,DOFPulse,DOFChimes), TrigRout
		Case 9: PlaySoundAt SoundFXDOF("bell10",131,DOFPulse,DOFChimes), TrigRout
		case 10: PlaySoundAt SoundFXDOF("bell1000",133,DOFPulse,DOFChimes), TrigRout
		case 11: PlaySoundAt SoundFXDOF("bell1000",133,DOFPulse,DOFChimes), TrigRout
		case 12: PlaySoundAt SoundFXDOF("bell1000",133,DOFPulse,DOFChimes), TrigRout
		case 13: PlaySoundAt SoundFXDOF("bell100",132,DOFPulse,DOFChimes), TrigRout
		Case 15: PlaySoundAt SoundFXDOF("bell10",131,DOFPulse,DOFChimes), TrigRout
        Case 16: me.enabled=0
    End Select
    saahc = saahc + 1
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
		gamov.timerinterval= (INT (RND*10)+5)*100
		gamov.timerenabled=1
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

Sub CARtoons_KeyDown(ByVal keycode)

	if keycode=AddCreditKey then
		PlaySoundAt "coin3", Drain
		if state=false then
			creditreel.setvalue(credit)
		end if
		coindelay.enabled=true 
    end if

	if keycode = 19 and roachtimer.enabled=0 then roachmove	 			' R Key

	if keycode = 46 then	 			' C Key
		If contball = 1 Then
			contball = 0
		Else
			contball = 1
		End If
	End If
	if keycode = 48 then 				'B Key
		If bcboost = 1 Then
			bcboost = bcboostmulti
		Else
			bcboost = 1
		End If
	End If
	if keycode = 203 then bcleft = 1		' Left Arrow
	if keycode = 200 then bcup = 1			' Up Arrow
	if keycode = 208 then bcdown = 1		' Down Arrow
	if keycode = 205 then bcright = 1		' Right Arrow


    if keycode=StartGameKey and (credit>0 or freeplay=1) and OperatorMenu=0 And Not HSEnterMode=true And Not RSEnterMode=true then
	  if state=false then
		if freeplay=0 then credit=credit-1
		If credit < 1 Then DOF 122, DOFOff
		playsoundat "click", Drain
		creditreel.setvalue(credit)
		ballinplay=1
'		if goroaches=0 then lightdelay.enabled=true
		lightdelay.enabled=true
		'goroaches=1
		If B2SOn Then 
			Controller.B2ssetCredits Credit
			Controller.B2sStartAnimation "Startup"
			Controller.B2ssetballinplay 32, Ballinplay
			Controller.B2ssetplayerup 30, 1
			Controller.B2SSetScoreRollover 24 + 1, 0
			Controller.B2SSetScoreRollover 24 + 2, 0
			Controller.B2ssetcanplay 31, 1
			Controller.B2SSetGameOver 0
		End If
		playsound "initialize"
		PUP1.state=1
		tilt=false
		state=true
		CanPlayReel.setvalue 1
		players=1
		rst=0
		resettimer.enabled=true
	  else if state=true and players < 2 and Ballinplay=1 then
		if freeplay=0 then credit=credit-1
		If credit < 1 Then DOF 122, DOFOff
		players=players+1
		creditreel.setvalue(credit)
		If B2SOn then
			Controller.B2ssetCredits Credit
			Controller.B2ssetcanplay 31, 2
		End If
		CanPlayReel.setvalue 2
		PlaySoundAt "click", Drain
	   end if 
	  end if
	end if

	If HSEnterMode Then 
		HighScoreProcessKey(keycode)
	  elseIf RSEnterMode Then 
		HighRoachProcessKey(keycode)
	end if

	If keycode = PlungerKey Then

		Plunger.PullBack
		PlaySoundAt "plungerpull", Plunger
	End If

	If keycode=LeftFlipperKey and State = false and OperatorMenu=0 then
		OperatorMenuTimer.Enabled = true
	end if

	If keycode=LeftFlipperKey and State = false and OperatorMenu=1 then
		Options=Options+1
		If Options=6 then Options=1
		playsound "target"
		Select Case (Options)
			Case 1:
				Option1.visible=true
				Option5.visible=False
			Case 2:
				Option2.visible=true
				Option1.visible=False
			Case 3:
				Option3.visible=true
				Option2.visible=False
			Case 4:
				Option4.visible=true
				Option3.visible=False
			Case 5:
				Option5.visible=true
				Option4.visible=False
		End Select
	end if

	If keycode=RightFlipperKey and State = false and OperatorMenu=1 then
	  PlaySound "click"
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
			if ebmode=0 Then
				ebmode=1
			  Else
				ebmode=0
			end if
			OptionEBmode.image="OptionsEBmode"&ebmode
			if ebmode=0 then
				repcard.image = "replycard"&replays
			  else
				repcard.image = "ebcard"&replays
			end if
		Case 4:
			Replays=Replays+1
			if Replays>3 then
				Replays=1
			end if
			Replay1=Replay1Table(Replays)
			Replay2=Replay2Table(Replays)
			replay3=Replay3Table(Replays)
			OptionReplays.image="OptionsReplays"&replays
			if ebmode=0 then
				repcard.image = "replycard"&replays
			  else
				repcard.image = "ebcard"&replays
			end if
		Case 5:
			OperatorMenu=0
			savehs
			HideOptions
	  End Select
	End If

  if tilt=false and state=true then
	If keycode = LeftFlipperKey Then
		leftflipperdimmer.enabled=1
		LeftFlipper.RotateToEnd
		PlaySoundat SoundFXDOF("fx_flipperup",101,DOFOn,DOFflippers), Lflip
		PlayXYSound "Buzz",LeftCurvedRail,-1,.05,0,0,0,1
	End If
    
	If keycode = RightFlipperKey Then
		rightflipperdimmer.enabled=1
		RightFlipper.RotateToEnd
		PlaySoundAt SoundFXDOF("fx_flipperup",102,DOFOn,DOFflippers), RFlip
		PlayXYSound "Buzz1",RightCurvedRail, -1,.05,0,0,0,1
	End If
    
	If keycode = LeftTiltKey Then
		Nudge 90, 2
		checktilt
	End If
    
	If keycode = RightTiltKey Then
		Nudge 270, 2
		checktilt
		flickertimer.enabled = 1
	End If
    
	If keycode = CenterTiltKey Then
		Nudge 0, 2
		checktilt
	End If

	If keycode = MechanicalTilt Then
		gametilted
	End If

  end if  
End Sub

sub roachinit_timer
	Dim roachA,dx,dy, xx
	
	roach1.x=roachxy1(0)(roachstep)
	roach1.y=roachxy1(1)(roachstep)
	Roach2.x=roachxy2(0)(roachstep)
	Roach2.y=roachxy2(1)(roachstep)
	roach3.x=roachxy3(0)(roachstep)
	roach3.y=roachxy3(1)(roachstep)
	roach4.x=roachxy4(0)(roachstep)
	roach4.y=roachxy4(1)(roachstep)
	if roachstep>0 Then
		for xx = 1 to 4											'change each roach rotation
			dy=-1*(EVAL("roachxy"&xx)(1)(roachstep)-EVAL("roachxy"&xx)(1)(roachstep-1))	'delta Y
			if dy=0 then dy=0.0000001													'avoid divide by zero errors
			dx=EVAL("roachxy"&xx)(0)(roachstep)-EVAL("roachxy"&xx)(0)(roachstep-1)		'delta X
			roachA=atn(dX/dY)												'angle in radians
			if roachA<0 then roachA=roachA+PI								'correction for negative angles
			roachA=int(roachA/(pi/180))							'convert to degrees
			if dx<0 then roachA=roachA+180						'correction for negative quadrants
			EVAL("roach"&xx).roty=roachA
		Next
	end If
	roachstep=roachstep+1
	if roachstep=24 then 
		AnimLoopTimer.enabled=False
		me.enabled=false
	end if
end Sub

sub AnimLoopTimer_Timer
	Roach1.ShowFrame frame
	Roach2.ShowFrame frame
	Roach3.ShowFrame frame
	Roach4.ShowFrame frame
	frame = frame + frameRate
	If frame>18 Then	' the animation is a bit rough so don't play it till the end but stop a bit earlier
		frame=0
	end if
end sub

Sub DisplayOptions
	OptionsBack.visible = true
	Option1.visible = True
	OptionBalls.visible = True
	OptionEBmode.visible = True
    OptionReplays.visible = True
	OptionFreeplay.visible = True
End Sub

Sub HideOptions
	for each objekt In OptionMenu
		objekt.visible = false
	next
End Sub

Sub OperatorMenuTimer_Timer
	OperatorMenu=1
	Displayoptions
	Options=1
End Sub

Sub StartControl_Hit()
	Set ControlBall = ActiveBall
	contballinplay = true
End Sub

Sub StopControl_Hit()
	contballinplay = false
End Sub	

Dim bcup, bcdown, bcleft, bcright, contball, contballinplay, ControlBall, bcboost
Dim bcvel, bcyveloffset, bcboostmulti

bcboost = 1		'Do Not Change - default setting
bcvel = 4		'Controls the speed of the ball movement
bcyveloffset = -0.01 	'Offsets the force of gravity to keep the ball from drifting vertically on the table, should be negative
bcboostmulti = 3	'Boost multiplier to ball veloctiy (toggled with the B key) 

Sub BallControl_Timer()
	If Contball and ContBallInPlay then
		If bcright = 1 Then
			ControlBall.velx = bcvel*bcboost
		ElseIf bcleft = 1 Then
			ControlBall.velx = - bcvel*bcboost
		Else
			ControlBall.velx=0
		End If

		If bcup = 1 Then
			ControlBall.vely = -bcvel*bcboost
		ElseIf bcdown = 1 Then
			ControlBall.vely = bcvel*bcboost
		Else
			ControlBall.vely= bcyveloffset
		End If
	End If
End Sub

Sub CARtoons_KeyUp(ByVal keycode)
	dim li

	if keycode = 203 then bcleft = 0		' Left Arrow
	if keycode = 200 then bcup = 0			' Up Arrow
	if keycode = 208 then bcdown = 0		' Down Arrow
	if keycode = 205 then bcright = 0		' Right Arrow

	If keycode = PlungerKey Then
		Plunger.Fire
		if PlungeBall=1 then
			playsoundat "plunger", Plunger
		  else
			playsoundat "plungerreleasefree", Plunger
		end if
	End If

   	if keycode = LeftFlipperKey then
		OperatorMenuTimer.Enabled = false
	end if

   If tilt=false and state=true then
	If keycode = LeftFlipperKey Then
		leftflipperdimmer.enabled=0
		if RightFlipper.currentangle>-60 then
			For Each li in GI
				li.Intensityscale = .4
			Next
		  else
			For Each li in GI
				li.Intensityscale = 1
			Next
		end if
		LeftFlipper.RotateToStart
		PlaySoundat SoundFXDOF("fx_flipperdown",101,DOFOff,DOFflippers), Lflip
		StopSound "Buzz"
	End If
    
	If keycode = RightFlipperKey Then
		rightflipperdimmer.enabled=0
		if leftFlipper.currentangle>110 then
			For Each li in GI
				li.Intensityscale = .4
			Next
		  else
			For Each li in GI
				li.Intensityscale = 1
			Next
		end if
		RightFlipper.RotateToStart
		PlaySoundAt SoundFXDOF("fx_flipperdown",102,DOFOff,DOFflippers), RFlip
		StopSound "Buzz1"
	End If
   End if
End Sub

sub LeftFlipperDimmer_timer
	Dim li
	For Each li in GI
		li.Intensityscale = li.Intensityscale * .90
		if li.intensityscale < .1 then li.intensityscale = .1
	Next
End Sub

Sub RightFlipperDimmer_timer
	Dim li
	For Each li in GI
		li.Intensityscale = li.Intensityscale * .90
		if li.intensityscale<.1 then li.intensityscale=.1
	Next
End sub

sub flippertimer_timer()
	LFlip.RotY = LeftFlipper.CurrentAngle
	RFlip.RotY = RightFlipper.CurrentAngle
	pup1r.state=pup1.state
	pup2r.state=pup2.state

	PrimGate1.Rotz = gate.CurrentAngle*.75+20

	if FlipperShadows=1 then
		FlipperLSh.RotZ = LeftFlipper.currentangle
		FlipperRSh.RotZ = RightFlipper.currentangle
	end if

End Sub


Sub EBtimer_timer
	starstate=0
	if starred.state=1 then starstate=starstate+1
	if starwhite.state=1 then starstate=starstate+1
	if stargreen.state=1 then starstate=starstate+1
	if staryellow.state=1 then starstate=starstate+1
	if starstate=4 then ExtraBall.state=1
end sub

Sub PairedlampTimer_timer
	red2.state = red1.state
	white2.state = white1.state
	green2.state = green1.state
	yellow2.state = yellow1.state
	white3.state = white1.state
	green3.state = green1.state
end sub

'mike add->
Sub FlickerTimer_timer
	If TLightRL1.State = 2 Then
		TLightRL1.State = 0:Light4.state = 0
	elseIf TLightRL1.State = 0 Then
		TLightRL1.State = 1:Light4.State = 1
	elseif TLightRL1.State = 1 Then
		TLightRL1.State = 2:Light4.state = 2
	end If
	me.enabled = 0
End Sub
'<- mike add

sub coindelay_timer
	addcredit
    coindelay.enabled=false
end sub

sub resettimer_timer
    rst=rst+1
    scorereel1.resettozero 
    scorereel2.resettozero 
	roachReel1.resettozero
	roachReel2.resettozero
	P1100k.text = " "
	P2100k.text = " "
    If B2SOn then
		Controller.B2SSetScore 1, score(1)
		Controller.B2SSetScore 2, score(2)
		Controller.B2SSetData 3,0
		Controller.B2SSetData 4,0
	End If
    if rst=18 then
		playsoundat "kickerkick", Drain
    end if
    if rst=22 then
		newgame
		resettimer.enabled=false
    end if
end sub

Sub addcredit
  if freeplay=0 then 
      credit=credit+1
	  DOF 122, DOFOn
      if credit>15 then credit=15
	  creditreel.setvalue(credit)
	  If B2SOn Then Controller.B2ssetCredits Credit
  end if
End sub

Sub Drain_Hit()
	PlaySoundat "drain", Drain
	DOF 108, DOFPulse
	scorebonus.enabled=true
End Sub

sub ballhome_hit
	plungeball=1
end sub

sub ballhome_unhit
	DOF 128, DOFPulse
	plungeball=0
end sub

sub ballrel_hit
	if ballrenabled=1 then
		if extraballs>0 then 
			extraballs=extraballs-1
			if extraballs>1 then
				shootagain.state=2
			  else
				shootagain.state=1
			end if
		end if
		if extraballs=0 then shootagain.state=0
		ballrenabled=0
	end if
end sub

sub scorebonus_timer
   if tilt=true then
	bonus=0
	Else
	if bonusx3.state=1 then 
		dbonus=3
		'me.interval=425
	  elseif bonusx2.state=1 then 
			dbonus=2
		'	me.interval=250
		  else 
			dbonus=1
			'me.interval=125
	  end if
	End if
	if bonus>0 then
		chimescount=dbonus
		chimestimer.interval=1
		chimestimer.enabled=true
		checkreplays
		bonuslight(bonus).state=0
		bonus=bonus-1
		if bonus >0 then 
			bonuslight(bonus).state=1
		end if
		me.enabled=False
	  else 
		bonus=0
		for each light in bonuslights: light.state=0: next
	   end if
	if bonus=0 and chimestimer.enabled=0 then 
     if shootagain.state>0 then
	    newball
 	    ballreltimer.enabled=true
      else
		PUP1.state=0
		PUP2.state=0
		if b2son then Controller.B2ssetplayerup 30, 0
		  if players=1 or player=2 then 
			player=1
			nextball
		  else
			player=2
			nextball
		  end if
	 end if
	 scorebonus.enabled=false
	end if
End sub

dim chimescount

sub chimestimer_timer
	chimestimer.interval=100
	addpoints 1000
	checkreplays
    chimescount = chimescount - 1
	if chimescount<1 then
		me.enabled=false
		ScoreBonus.enabled=True
	end if
end Sub

sub newgame
	bumper1.hashitevent=1
	bumper2.hashitevent=1
	DTResetTimer.enabled=True

	player=1
	PUP1.state=1
    score(1)=0
	score(2)=0
	RoachScore(1)=0
	RoachScore(2)=0
	If B2SOn then
		Controller.B2SSetScoreplayer 1, score(1)
		Controller.B2SSetScorePlayer 2, score(2)
		Controller.B2SSetScorePlayer 3, roachscore(1)
		Controller.B2SSetScorePlayer 4, roachscore(2)
	End If
    eg=0
    rep(1)=0
    rep(2)=0
	for each light in bonuslights:light.state=0:next
    special.state=0
	extraballs=0
    ExtraBall.state=0
	shootagain.state=lightstateoff
	for each light in bumperlights:light.state=1:next
	for each light in tlights:light.state=1:next
	for each light in lanelights:light.state=1:next
	for each light in colorlights:light.state=1:next
	TLightRL1.State = 1:Light4.state = 1 'mike add
    gamov.text=" "
    tilttxt.text=" "
    If B2SOn then
		Controller.B2SSetGameOver 35,0
		Controller.B2SSetTilt 33,0
		Controller.B2SSetMatch 34,0
		Controller.B2ssetdata 11, 0
		Controller.B2ssetdata 7,0
		Controller.B2ssetdata 8,0
	End If
    biptxt.text="1"
	matchtxt.text=" "
		ballreltimer.enabled=true
end sub

sub newball
	for each light in colorlights:light.state=1:next
	for each light in starlights:light.state=0:next
	DTResetTimer.enabled=True
	extraball.state=0
	if ballinplay=balls then bonusx2.state=1
End Sub


sub nextball
    if tilt=true then
	  bumper1.hashitevent=1
	  bumper2.hashitevent=1
      tilt=false
      tilttxt.text=" "
		If B2SOn then
			Controller.B2SSetTilt 33,0
			Controller.B2ssetdata 1, 1
		End If
    end if

	bonusx2.state=0
	bonusx3.state=0
	special.state=0
	if player=1 then ballinplay=ballinplay+1
	if ballinplay>balls then
		playsound "motorleer"
		eg=1
		ballreltimer.enabled=true
	  else
		if state=true and tilt=false then
		  newball
		  ballreltimer.enabled=true
		end if
		if player=1 then 
				PUP1.state=1
				PUP2.state=0
			Else
				PUP1.state=0
				PUP2.state=1
		end If
		biptxt.text=ballinplay
		If B2SOn then 
			Controller.B2ssetballinplay 32, Ballinplay
			Controller.B2ssetplayerup 30, player
		end If
	end if
End Sub

sub ballreltimer_timer
  if eg=1 then
      matchnum
	  hiscstate=0
	  hircstate=0
	  biptxt.text=" "
	  state=false
	  saahc=1
	  shavehaircut.interval=1000
	  shavehaircut.enabled=1
	  gamov.timerenabled=1
	  tilttxt.timerenabled=1
	  CanPlayReel.setvalue 0
	  for i=1 to 2
		EVAL("Pup"&i).state=0
	  next

		if score(1)>hisc or score(2)>hisc then 
			if score(2)> score(1) then
				hisc=score(2)
			  else
				hisc=score(1)
			end if
			hiscstate=1
			HighScoreEntryInit()
		end if
		if roachscore(1)>Rhisc or roachscore(2)>Rhisc then 
			if RoachScore(2) > roachscore(1) then
				Rhisc=roachscore(2)
			  else
				Rhisc=roachscore(1)
			end if
			hircstate=1
			if HSEnterMode <> True then HighRoachEntryInit
		end if

	  UpdatePostIt 
	  savehs
	  If B2SOn then 
        Controller.B2SSetGameOver 35,1
        Controller.B2ssetballinplay 32, 0
	    Controller.B2sStartAnimation "EOGame"
	    Controller.B2ssetcanplay 31, 0
	    Controller.B2ssetcanplay 30, 0
	  End If
	  For each light in LaneLights:light.State = 0: Next
	  For each light in BumperLights:light.State = 0: Next
	  For each light in Tlights:light.State = 0: Next
	  For each light in PFlights:light.State = 0: Next
	  TLightRL1.State = 0:Light4.state = 0 'mike add
	  LeftFlipper.RotateToStart
 	  StopSound "Buzz"
	  RightFlipper.RotateToStart
	  StopSound "Buzz1"
	  ballreltimer.enabled=false
  else
	drain.kick 60,12,5
	ballrenabled=1
    PlaySoundAt SoundFXDOF("ballrelease",107,DOFPulse,DOFContactors), Drain
    ballreltimer.enabled=false
  end if
end sub

sub shavehaircut_timer()
	shavehaircut.interval=100
    Select Case saahc
        Case 1: PlaySoundAt SoundFXDOF("bell10",131,DOFPulse,DOFChimes), TrigRout	'playsound"bell10"
		Case 3: PlaySoundAt SoundFXDOF("bell1000",133,DOFPulse,DOFChimes), TrigRout	'playsound"bell1000"
		case 4: PlaySoundAt SoundFXDOF("bell1000",133,DOFPulse,DOFChimes), TrigRout	'playsound"bell1000"
		case 5: PlaySoundAt SoundFXDOF("bell100",132,DOFPulse,DOFChimes), TrigRout	'playsound"bell100"
		Case 7: PlaySoundAt SoundFXDOF("bell1000",133,DOFPulse,DOFChimes), TrigRout	'playsound "bell1000"
		case 11: PlaySoundAt SoundFXDOF("bell10",131,DOFPulse,DOFChimes), TrigRout	'playsound"bell10"
		case 13: PlaySoundAt SoundFXDOF("bell10",131,DOFPulse,DOFChimes), TrigRout	'playsound"bell10"
        Case 16: me.enabled=0
    End Select
    saahc = saahc + 1
end Sub

Sub HStimer_timer
	playsoundat SoundFXDOF("knocker",121,DOFPulse,DOFKnocker), Laneguard
	DOF 139,DOFPulse
	HStimer.uservalue=HStimer.uservalue+1
	if HStimer.uservalue=3 then me.enabled=0
end sub

Sub RStimer_timer
	playsoundat SoundFXDOF("knocker",121,DOFPulse,DOFKnocker), Laneguard
	DOF 139,DOFPulse
	RStimer.uservalue=RStimer.uservalue+1
	if RStimer.uservalue=2 then me.enabled=0
end sub

sub matchnum
	matchnumb=INT (RND*10)
	if matchnumb=0 then
		matchtxt.text="00"
	  else
		matchtxt.text=matchnumb*10
	end if
  If B2SOn then Controller.B2SSetMatch 34,Matchnumb*10
  For i=1 to players
    if (matchnumb*10)=(score(i) mod 100) then 
		addcredit
		PlaySoundat SoundFXDOF("knocker",121,DOFPulse,DOFKnocker), Laneguard
	end if
  next
end sub

'********** Bumpers

Sub Bumper1_Hit
   if tilt=false then
    PlaySoundAt SoundFXDOF("fx_bumper4",105, DOFPulse,DOFContactors), Bumper1
	if bumper1light.state=1 then
		addscore bumperlitscore
	  else
		addscore bumperoffscore
	end if
   end if
	GIDown.enabled = 1
End Sub

Sub Bumper2_Hit
   if tilt=false then
    PlaySoundAt SoundFXDOF("fx_bumper4",106, DOFPulse,DOFContactors), Bumper2
	if bumper2light.state=1 then
		addscore bumperlitscore
	  else
		addscore bumperoffscore
	end if	
   end if
	GIdown.enabled = 1
End Sub


sub FlashBumpers
	For each light in BumperLights
	  Light.State=0
	next
	FlashB.enabled=1
end sub

sub FlashB_timer
	For each light in BumperLights
	  Light.State=1
	next
	FlashB.enabled=0
end sub

Sub RubberDing_hit
	if roachloc=2 then roachmove
end sub

Sub RubberDing1_hit
	if roachloc=1 then roachmove
end sub

Sub RubberDing4_hit
	if roachloc=0 then roachmove
end sub

Sub RubberDing5_hit
	if roachloc=0 then roachmove
end sub


Sub Dingwalls_Hit(idx)
	addscore 10
end sub

'********** Dingwalls - animated - timer 50

sub dingwalla_hit
	if state=true and tilt=false then addscore 10
	SlingA.visible=0
	SlingA1.visible=1
	me.uservalue=1
	Me.timerenabled=1
end sub

sub dingwalla_timer									'default 50 timer
	select case me.uservalue
		Case 1: SlingA1.visible=0: SlingA.visible=1
		case 2:	SlingA.visible=0: SlingA2.visible=1
		Case 3: SlingA2.visible=0: SlingA.visible=1: Me.timerenabled=0
	end Select
	me.uservalue=me.uservalue+1
end sub


sub dingwallb_hit
	if state=true and tilt=false then addscore 10
	SlingB.visible=0
	SlingB1.visible=1
	me.uservalue=1
	Me.timerenabled=1
end sub

sub dingwallb_timer									'default 50 timer
	select case me.uservalue
		Case 1: Slingb1.visible=0: SlingB.visible=1
		case 2:	SlingB.visible=0: Slingb2.visible=1
		Case 3: Slingb2.visible=0: SlingB.visible=1: Me.timerenabled=0
	end Select
	me.uservalue=me.uservalue+1
end sub

sub dingwallc_hit
	if roachloc=1 then roachmove
	if state=true and tilt=false then addscore 10
	SlingC.visible=0
	SlingC1.visible=1
	me.uservalue=1
	Me.timerenabled=1
end sub

sub dingwallc_timer									'default 50 timer
	select case me.uservalue
		Case 1: Slingc1.visible=0: SlingC.visible=1
		case 2:	SlingC.visible=0: Slingc2.visible=1
		Case 3: Slingc2.visible=0: Slingc.visible=1: Me.timerenabled=0
	end Select
	me.uservalue=me.uservalue+1
end sub

sub dingwalld_hit
	if roachloc=2 then roachmove
	if state=true and tilt=false then addscore 10
	SlingD.visible=0
	Slingd1.visible=1
	me.uservalue=1
	Me.timerenabled=1
end sub

sub dingwalld_timer									'default 50 timer
	select case me.uservalue
		Case 1: Slingd1.visible=0: SlingD.visible=1
		case 2:	SlingD.visible=0: Slingd2.visible=1
		Case 3: Slingd2.visible=0: SlingD.visible=1: Me.timerenabled=0
	end Select
	me.uservalue=me.uservalue+1
end sub

'****** Color switches

Sub TRed1_Hit
   if tilt=false then
	DOF 109, DOFPulse
	flashbumpers
	If StarRed.State = 0 then StarRed.State  = 1
	Red1.state = 0
	If balls=3 then
		If StarYellow.State = 0 then StarYellow.State = 1
		Yellow1.state = 0
	end if
	addscore 500
	addbonus
   end if
End Sub

Sub TRed2_Hit
   if tilt=false then
	DOF 110, DOFPulse
	flashbumpers
	If StarRed.State = 0 then StarRed.State  = 1
	Red1.state = 0
	If balls=3 then
		If StarYellow.State = 0 then StarYellow.State = 1
		Yellow1.state = 0
	end if
	addscore 500
	addbonus
   end if
End Sub

Sub TWhite1_Hit
   if tilt=false then
	DOF 111, DOFPulse
	flashbumpers
	If StarWhite.State = 0 then StarWhite.State  = 1
	White1.state = 0
	addscore 500
	addbonus
   end if
End Sub

Sub TWhite2_Hit
   if tilt=false then
	DOF 112, DOFPulse
	flashbumpers
	If StarWhite.State = 0 then StarWhite.State  = 1
	White1.state = 0
	addscore 500
	addbonus
   end if
End Sub

Sub TWhite3_Hit
	if roachloc=3 then roachmove
   if tilt=false then
	DOF 113, DOFPulse
	flashbumpers
	If StarWhite.State = 0 then StarWhite.State  = 1
	White1.state = 0
	addscore 500
	addbonus
   end if
End Sub

Sub TGreen1_Hit
   if tilt=false then
	DOF 114, DOFPulse
	flashbumpers
	If StarGreen.State = 0 then StarGreen.State  = 1
	Green1.state = 0
	addscore 500
	addbonus
   end if
End Sub

Sub TGreen2_Hit
   if tilt=false then
	DOF 115, DOFPulse
	flashbumpers
	If StarGreen.State = 0 then StarGreen.State  = 1
	Green1.state = 0
	addscore 500
	addbonus
   end if
End Sub

Sub TGreen3_Hit
	if roachloc=4 then roachmove
   if tilt=false then
	DOF 116, DOFPulse
	flashbumpers
	If StarGreen.State = 0 then StarGreen.State  = 1
	Green1.state = 0
	addscore 500
	addbonus
   end if
End Sub

Sub TYellow1_Hit
	if roachloc=0 then roachmove
   if tilt=false then
	DOF 117, DOFPulse
	flashbumpers
	If StarYellow.State = 0 then StarYellow.State  = 1
	Yellow1.state = 0
	If balls=3 then
		StarRed.State = 1
		Red1.state = 0
	end if
	addscore 500
	addbonus
   end if
End Sub

Sub TYellow2_Hit
	if roachloc=1 then roachmove
   if tilt=false then
	DOF 118, DOFPulse
	flashbumpers
	If StarYellow.State = 0 then StarYellow.State  = 1
	Yellow1.state = 0
	If balls=3 then
		StarRed.State = 1
		Red1.state = 0
	end if
	addscore 500
	addbonus
   end if
End Sub

Sub TrigRout_Hit
	if tilt=false then
	  DOF 120, DOFPulse
	  flashbumpers
	  if special.state=1 then 
		if ebmode=0 then
			PlaySoundAt SoundFXDOF("knocker",121,DOFPulse,DOFKnocker), Laneguard
			addcredit
		  else
			PlaySoundAt SoundFXDOF("bell",134,DOFPulse,DOFBell), TrigLout
			extraballs=extraballs+1	
			if extraballs>1 then
				shootagain.state=2
			  else
				shootagain.state=1
			end if
		end if
	  end if
	  addscore 5000
	  addbonus
	end if
End Sub

Sub TrigLout_Hit
	if tilt=false then
	  DOF 119, DOFPulse
	  flashbumpers
	  if extraball.state=1 then 
		PlaySoundAt SoundFXDOF("bell",134,DOFPulse,DOFBell), TrigLout
		extraballs=extraballs+1
		if extraballs>1 then
			shootagain.state=2
		  else
			shootagain.state=1
		end if
	  end if
	  addscore 5000
	  addbonus
	end if
End Sub

'****Targets

Sub AwardCheck_timer
	awardcount=0
	For j = 1 to 8
		if target(j).isdropped then awardcount=awardcount+1
	next

	if awardcount=8 and DTResetTimer.enabled=false then 
		if bonusx2.state=1 or bonusx3.state=1 then 
			bonusx3.state=1
			bonusx2.state=0
		  else
			bonusx2.state=1
		end if
		special.state=1
		DTResetTimer.enabled=True
	end if
End Sub

Sub DTResetTimer_timer
	PlaySoundAt SoundFXDOF("bankreset",127,DOFPulse,DOFContactors), DTGreenL
	for i=1 to 8
		target(i).isdropped=false
	next
	for each objekt in DTroachWalls: objekt.isdropped=false: next
	for each objekt in DTlights : objekt.state=0: next
	me.enabled=False
end Sub


Sub DTredL_dropped
	if roachloc=2 then roachmove
	PlaySoundAt SoundFXDOF("target",123,DOFPulse,DOFContactors), DTRedL
	addscore (starstate+1)*1000
	LDTRedL.state=1
	LDTRedL1.state=1
	If StarRed.State = 1 then 
		addbonus
	end if
End Sub

Sub DWredL_hit
	DWredL.isdropped=true
	DTredL.isdropped=true
end sub

Sub DTredR_dropped
	if roachloc=2 then roachmove
	PlaySoundAt SoundFXDOF("target",123,DOFPulse,DOFContactors), DTRedR
	addscore (starstate+1)*1000
	LDTRedR.state=1
	If StarRed.State = 1 then 
		addbonus
	end if
End Sub

Sub DWredR_hit
	DWRedR.isdropped=true
	DTRedR.isdropped=true
end sub

Sub DTwhiteL_dropped
	PlaySoundAt SoundFXDOF("target",124,DOFPulse,DOFContactors), DTwhiteL
	addscore (starstate+1)*1000
	LDTwhiteL.state=1
	If StarWhite.State = 1 then 
		addbonus
	end if
End Sub

Sub DWwhiteL_hit
	DWwhiteL.isdropped=true
	DTwhiteL.isdropped=true
end sub

Sub DTwhiteR_dropped
	PlaySoundAt SoundFXDOF("target",124,DOFPulse,DOFContactors), DTwhiteR
	addscore (starstate+1)*1000
	LDTwhiteR.state=1
	If StarWhite.State = 1 then 
		addbonus
	end if
End Sub

Sub DWwhiteR_hit
	DWWhiteR.isdropped=true
	DTwhiteR.isdropped=true
end sub

Sub DTgreenL_dropped
	PlaySoundAt SoundFXDOF("target",125,DOFPulse,DOFContactors), DTGreenL
	addscore (starstate+1)*1000
	LDTGreenL.state=1
	If StarGreen.State = 1 then 
		addbonus
	end if
End Sub

Sub DWgreenL_hit
	DWgreenL.isdropped=true
	DTgreenL.isdropped=true
end sub

Sub DTGreenR_dropped
	PlaySoundAt SoundFXDOF("target",125,DOFPulse,DOFContactors), DTGreenR
	addscore (starstate+1)*1000
	LDTGreenR.state=1
	If StarGreen.State = 1 then 
		addbonus
	end if
End Sub

Sub DWgreenR_hit
	DWgreenR.isdropped=true
	DTgreenR.isdropped=true
end sub

Sub DTyellowL_dropped
	if roachloc=1 then roachmove
	PlaySoundAt SoundFXDOF("target",126,DOFPulse,DOFContactors), DTYellowL
	addscore (starstate+1)*1000
	LDTYellowL.state=1
	If StarYellow.State = 1 then 
		addbonus
	end if
End Sub

Sub DWyellowL_hit
	DWyellowL.isdropped=true
	DTyellowL.isdropped=true
end sub

Sub DTyellowR_dropped
	if roachloc=1 then roachmove
	PlaySoundAt SoundFXDOF("target",126,DOFPulse,DOFContactors), DTYellowR
	addscore (starstate+1)*1000
	LDTYellowR.state=1
	If StarYellow.State = 1 then 
		addbonus
	end if
End Sub

Sub DWyellowR_hit
	DWyellowR.isdropped=true
	DTyellowR.isdropped=true
end sub

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

Sub AddPoints(Points)			  ' Sounds: there are 3 sounds: tens, hundreds and thousands
    score(player)=score(player)+points
	sreels(player).addvalue(points)
	If B2SOn Then Controller.B2SSetScorePlayer player, score(player) MOD 100000

    If Points = 100 AND(Score(player) MOD 1000) \ 100 = 0 Then  'New 1000 reel
        PlaySoundAt SoundFXDOF("bell1000",133,DOFPulse,DOFChimes), TrigRout
      ElseIf Points = 10 AND(Score(player) MOD 100) \ 10 = 0 Then 'New 100 reel
        PlaySoundAt SoundFXDOF("bell100",132,DOFPulse,DOFChimes), TrigRout
      ElseIf points = 1000 Then
        PlaySoundAt SoundFXDOF("bell1000",133,DOFPulse,DOFChimes), TrigRout
	  elseif Points = 100 Then
        PlaySoundAt SoundFXDOF("bell100",132,DOFPulse,DOFChimes), TrigRout
	  Else
        PlaySoundAt SoundFXDOF("bell10",131,DOFPulse,DOFChimes), TrigRout
    End If
	checkreplays
end sub

Sub checkreplays
    ' check replays and rollover
	if score(player)>99999 then
		If B2SOn then Controller.B2SSetScoreRollover 24 + player, Int(score(player)/100000)
		P100k(player).text= 100000*Int(score(player)/100000) 
	End if
    if score(player)=>replay1 and rep(player)=0 then
		rep(player)=1
		if ebmode=0 then
			PlaySoundAt SoundFXDOF("knocker",121,DOFPulse,DOFKnocker), Laneguard
			addcredit
		  else
			PlaySoundAt SoundFXDOF("bell",134,DOFPulse,DOFBell), TrigLout
			extraballs=extraballs+1	
			if extraballs>1 then
				shootagain.state=2
			  else
				shootagain.state=1
			end if
		end if
    end if
    if score(player)=>replay2 and rep(player)=1 then
		rep(player)=2
		if ebmode=0 then
			PlaySoundAt SoundFXDOF("knocker",121,DOFPulse,DOFKnocker), Laneguard
			addcredit
		  else
			PlaySoundAt SoundFXDOF("bell",134,DOFPulse,DOFBell), TrigLout
			extraballs=extraballs+1	
			if extraballs>1 then
				shootagain.state=2
			  else
				shootagain.state=1
			end if
		end if
    end if
    if score(player)=>replay3 and rep(player)=2 then
		rep(player)=3
		if ebmode=0 then
			PlaySoundAt SoundFXDOF("knocker",121,DOFPulse,DOFKnocker), Laneguard
			addcredit
		  else
			PlaySoundAt SoundFXDOF("bell",134,DOFPulse,DOFBell), TrigLout
			extraballs=extraballs+1	
			if extraballs>1 then
				shootagain.state=2
			  else
				shootagain.state=1
			end if
		end if
    end if
end sub 


Sub CheckTilt
	If Tilttimer.Enabled = True Then 
	 TiltSens = TiltSens + 1
	 if TiltSens = 3 Then GameTilted
	Else
	 TiltSens = 0
	 Tilttimer.Enabled = True
	End If
End Sub

Sub Tilttimer_Timer()
	Tilttimer.Enabled = False
End Sub

Sub GameTilted
	Tilt = True
	tilttxt.text="TILT"
    If B2SOn Then Controller.B2SSetTilt 33,1
    If B2SOn Then Controller.B2ssetdata 1, 0
	playsound "tilt"
	turnoff
End Sub

sub turnoff
    bumper1.hashitevent=0
    bumper2.hashitevent=0
	LeftFlipper.RotateToStart
	DOF 101, DOFOff
	StopSound "Buzz"
	RightFlipper.RotateToStart
	DOF 102, DOFOff
	StopSound "Buzz1"
end sub    


Sub addbonus
	bonus=bonus+1
	if bonus>19 then bonus=19
	if  bonus = 1 then 
		bonuslight(bonus).state=1
	  else
		bonuslight(bonus).state=1
		bonuslight(bonus-1).state=0
		if bonus>10 then bonuslight(10).state=1
	End if
End sub

'**********Sling Shot Animations
' Rstep and Lstep  are the variables that increment the animation
'****************
Dim RStep, Lstep

Sub RightSlingShot_Slingshot
    PlaySoundAt SoundFXDOF("right_slingshot",103,DOFPulse,DOFContactors), slingR
	addscore 10
	if roachloc=4 then roachmove
    RSling.Visible = 0
    RSling1.Visible = 1
    slingR.objroty = -15
	FlickerTimer.Enabled = 1 'mike add
    RStep = 0
    RightSlingShot.TimerEnabled = 1
'	PlightRL.State = 0:TLightRL.State = 0:PlightRL1.State = 0:TLightRL1.State = 0
	GIDown.enabled = 1
End Sub

Sub RightSlingShot_Timer
    Select Case RStep
        Case 3:RSLing1.Visible = 0:RSLing2.Visible = 1:slingR.objroty = -7
        Case 4:RSLing2.Visible = 0:RSLing.Visible = 1:slingR.objroty = 0:RightSlingShot.TimerEnabled = 0 ':PLightRL.State = 1:TLightRL.State = 1:PlightRL1.State = 1:TLightRL1.State = 1
    End Select
    RStep = RStep + 1
End Sub

Sub LeftSlingShot_Slingshot
    PlaySoundAt SoundFXDOF("left_slingshot",104,DOFPulse,DOFContactors), slingL
	addscore 10
	if roachloc=3 then roachmove
    LSling.Visible = 0
    LSling1.Visible = 1
    slingL.objroty = 15
    LStep = 0
    LeftSlingShot.TimerEnabled = 1
'	PLightLL.State = 0:TLightLL.State = 0:PLightLL1.State = 0:TLightLL1.State = 0
	GIDown.enabled = 1
End Sub

Sub LeftSlingShot_Timer
    Select Case LStep
        Case 3:LSLing1.Visible = 0:LSLing2.Visible = 1:slingL.objroty = 7
        Case 4:LSLing2.Visible = 0:LSLing.Visible = 1:slingL.objroty = 0:LeftSlingShot.TimerEnabled = 0 ':PLightLL.State = 1:TLightLL.State = 1:PLightLL1.State = 1:TLightLL1.State = 1
    End Select
    LStep = LStep + 1
End Sub

'mike add ->
Sub Rubber14_Hit
	FlickerTimer.Enabled = 1
End Sub

Dim ii, OI
	For Each ii in GI
	OI = ii.intensity
	Next

Sub UpdateGI(giNo, Object)
Dim li
   Select Case giNo
      Case 1 
		For Each li in GI
		li.Intensity = li.Intensity * .6
		Next
	  Case 0
		For Each li in GI
		li.Intensity = OI
		Next
	End Select
	GIup.Enabled = 1
End Sub

sub GIDown_Timer
	UpdateGI 1, GI
	me.Enabled = 0
End Sub

Sub GIUp_Timer
	UpdateGI 0, GI
	me.Enabled = 0
End Sub

'<-mike add

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


' *********************************************************************
'                      Supporting Ball & Sound Functions
' *********************************************************************

Function AudioFade(tableobj) ' Fades between front and back of the table (for surround systems or 2x2 speakers, etc), depending on the Y position on the table. "table1" is the name of the table
	Dim tmp
    tmp = tableobj.y * 2 / CARtoons.height-1
    If tmp > 0 Then
		AudioFade = Csng(tmp ^10)
    Else
        AudioFade = Csng(-((- tmp) ^10) )
    End If
End Function

Function AudioPan(tableobj) ' Calculates the pan for a tableobj based on the X position on the table. "CARtoons" is the name of the table
    Dim tmp
    tmp = tableobj.x * 2 / CARtoons.width-1
    If tmp > 0 Then
        AudioPan = Csng(tmp ^10)
    Else
        AudioPan = Csng(-((- tmp) ^10) )
    End If
End Function
'Mike change ->
Function Vol(ball) ' Calculates the Volume of the sound based on the ball speed
	Vol =  (Round(BallVel(ball)/10,2))
End Function
' <-
Function Pitch(ball) ' Calculates the pitch of the sound based on the ball speed
    Pitch = BallVel(ball) * 20
End Function

Function BallVel(ball) 'Calculates the ball speed
    BallVel = INT(SQR((ball.VelX ^2) + (ball.VelY ^2) ) )
End Function

'*****************************************
'    JP's VP10 Rolling Sounds
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
        If BallVel(BOT(b) ) > .1 AND BOT(b).z < 30 Then
            rolling(b) = True
            PlaySound("fx_ballrolling" & b), -1, Vol(BOT(b)), AudioPan(BOT(b)), 0, Pitch(BOT(b)), 1, 0, AudioFade(BOT(b))
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
	PlaySound("fx_collide"), 0, Csng(velocity) ^2 / 2000, AudioPan(ball1), 0, Pitch(ball1), 0, 0, AudioFade(ball1)
End Sub



'*****************************************
'			BALL SHADOW
'*****************************************
Dim BallShadow
BallShadow = Array (BallShadow1,BallShadow2,BallShadow3,BallShadow4,BallShadow5)

Sub BallShadowUpdate_timer()
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
        If BOT(b).X < CARtoons.Width/2 Then
            BallShadow(b).X = ((BOT(b).X) - (Ballsize/16) + ((BOT(b).X - (CARtoons.Width/2))/17))' + 13
        Else
            BallShadow(b).X = ((BOT(b).X) + (Ballsize/16) + ((BOT(b).X - (CARtoons.Width/2))/17))' - 13
        End If
        ballShadow(b).Y = BOT(b).Y + 10
        If BOT(b).Z > 20 Then
            BallShadow(b).visible = 1
        Else
            BallShadow(b).visible = 0
        End If
    Next
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

Sub Pins_Hit (idx)
	PlaySound "pinhit_low", 0, Vol(ActiveBall), AudioPan(ActiveBall), 0, Pitch(ActiveBall), 0, 0, AudioFade(ActiveBall)
End Sub

Sub Targets_Hit (idx)
	PlaySound "target", 0, Vol(ActiveBall), AudioPan(ActiveBall), 0, Pitch(ActiveBall), 0, 0, AudioFade(ActiveBall)
End Sub

Sub Metals_Thin_Hit (idx)
	PlaySound "metalhit_thin", 0, Vol(ActiveBall), AudioPan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
End Sub

Sub Metals_Medium_Hit (idx)
	PlaySound "metalhit_medium", 0, Vol(ActiveBall), AudioPan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
End Sub

Sub Metals2_Hit (idx)
	PlaySound "metalhit2", 0, Vol(ActiveBall), AudioPan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
End Sub

Sub Gates_Hit (idx)
	PlaySound "gate4", 0, Vol(ActiveBall), AudioPan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
End Sub

Sub Spinner_Spin
	PlaySound "fx_spinner", 0, .25, AudioPan(Spinner), 0.25, 0, 0, 1, AudioFade(Spinner)
End Sub

Sub Rubbers_Hit(idx)
 	dim finalspeed
  	finalspeed=SQR(activeball.velx * activeball.velx + activeball.vely * activeball.vely)
 	If finalspeed > 20 then 
		PlaySound "fx_rubber2", 0, Vol(ActiveBall), AudioPan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
	End if
	If finalspeed >= 6 AND finalspeed <= 20 then
 		RandomSoundRubber()
 	End If
End Sub

Sub Posts_Hit(idx)
 	dim finalspeed
  	finalspeed=SQR(activeball.velx * activeball.velx + activeball.vely * activeball.vely)
 	If finalspeed > 16 then 
		PlaySound "fx_rubber2", 0, Vol(ActiveBall), AudioPan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
	End if
	If finalspeed >= 6 AND finalspeed <= 16 then
 		RandomSoundRubber()
 	End If
End Sub

Sub RandomSoundRubber()
	Select Case Int(Rnd*3)+1
		Case 1 : PlaySound "rubber_hit_1", 0, Vol(ActiveBall), AudioPan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
		Case 2 : PlaySound "rubber_hit_2", 0, Vol(ActiveBall), AudioPan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
		Case 3 : PlaySound "rubber_hit_3", 0, Vol(ActiveBall), AudioPan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
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
		Case 1 : PlaySound "flip_hit_1", 0, Vol(ActiveBall), AudioPan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
		Case 2 : PlaySound "flip_hit_2", 0, Vol(ActiveBall), AudioPan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
		Case 3 : PlaySound "flip_hit_3", 0, Vol(ActiveBall), AudioPan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
	End Select
End Sub


sub savehs

    savevalue "CARtoonsRC", "credit", credit
    savevalue "CARtoonsRC", "hiscore", hisc
    savevalue "CARtoonsRC", "Roachhiscore", Rhisc
    savevalue "CARtoonsRC", "match", matchnumb
    savevalue "CARtoonsRC", "score1", score(1)
    savevalue "CARtoonsRC", "score2", score(2)
    savevalue "CARtoonsRC", "RCscore1", roachscore(1)
    savevalue "CARtoonsRC", "RCscore2", roachscore(2)
	savevalue "CARtoonsRC", "replays", replays
	savevalue "CARtoonsRC", "balls", balls
	savevalue "CARtoonsRC", "freeplay", freeplay
	savevalue "CARtoonsRC", "ebmode", ebmode
	savevalue "CARtoonsRC", "RoachLoc", roachloc
	savevalue "CARtoonsRC", "hsa1", HSA1
	savevalue "CARtoonsRC", "hsa2", HSA2
	savevalue "CARtoonsRC", "hsa3", HSA3
	savevalue "CARtoonsRC", "rsa1", RSA1
	savevalue "CARtoonsRC", "rsa2", RSA2
	savevalue "CARtoonsRC", "rsa3", RSA3
end sub

sub loadhs
    dim temp
	temp = LoadValue("CARtoonsRC", "credit")
    If (temp <> "") then credit = CDbl(temp)
    temp = LoadValue("CARtoonsRC", "hiscore")
    If (temp <> "") then hisc = CDbl(temp)

    temp = LoadValue("CARtoonsRC", "Roachhiscore")
    If (temp <> "") then Rhisc = CDbl(temp)

    temp = LoadValue("CARtoonsRC", "match")
    If (temp <> "") then matchnumb = CDbl(temp)
    temp = LoadValue("CARtoonsRC", "score1")
    If (temp <> "") then score(1) = CDbl(temp)
    temp = LoadValue("CARtoonsRC", "score2")
    If (temp <> "") then score(2) = CDbl(temp)

    temp = LoadValue("CARtoonsRC", "RCscore1")
    If (temp <> "") then roachscore(1) = CDbl(temp)
    temp = LoadValue("CARtoonsRC", "RCscore2")
    If (temp <> "") then roachscore(2) = CDbl(temp)

    temp = LoadValue("CARtoonsRC", "replays")
    If (temp <> "") then replays = CDbl(temp)
    temp = LoadValue("CARtoonsRC", "balls")
    If (temp <> "") then balls = CDbl(temp)
    temp = LoadValue("CARtoonsRC", "freeplay")
    If (temp <> "") then freeplay = CDbl(temp)
    temp = LoadValue("CARtoonsRC", "ebmode")
    If (temp <> "") then ebmode = CDbl(temp)
    temp = LoadValue("CARtoonsRC", "roachloc")
    If (temp <> "") then RoachLoc = CDbl(temp)
    temp = LoadValue("CARtoonsRC", "hsa1")
    If (temp <> "") then HSA1 = CDbl(temp)
    temp = LoadValue("CARtoonsRC", "hsa2")
    If (temp <> "") then HSA2 = CDbl(temp)
    temp = LoadValue("CARtoonsRC", "hsa3")
    If (temp <> "") then HSA3 = CDbl(temp)
    temp = LoadValue("CARtoonsRC", "rsa1")
    If (temp <> "") then RSA1 = CDbl(temp)
    temp = LoadValue("CARtoonsRC", "rsa2")
    If (temp <> "") then RSA2 = CDbl(temp)
    temp = LoadValue("CARtoonsRC", "rsa3")
    If (temp <> "") then RSA3 = CDbl(temp)
end sub

Sub CARtoons_Exit()
	savehs
	If B2SOn Then Controller.stop
End Sub

'==========================================================================================================================================
'============================================================= START OF HIGH SCORES ROUTINES =============================================================
'==========================================================================================================================================
'
'ADD LINE TO TABLE_KEYDOWN SUB WITH THE FOLLOWING:    If HSEnterMode Then HighScoreProcessKey(keycode) AFTER THE STARTGAME ENTRY
'ADD: And Not HSEnterMode=true TO IF KEYCODE=STARTGAMEKEY
'TO SHOW THE SCORE ON POST-IT ADD LINE AT RELEVENT LOCATION THAT HAS:  UpdatePostIt
'TO INITIATE ADDING INITIALS ADD LINE AT RELEVENT LOCATION THAT HAS:  HighScoreEntryInit()
'ADD THE FOLLOWING LINES TO TABLE_INIT TO SETUP POSTIT
'	if HSA1="" then HSA1=25
'	if HSA2="" then HSA2=25
'	if HSA3="" then HSA3=25
'	UpdatePostIt
'ADD HSA1, HSA2 AND HSA3 TO SAVE AND LOAD VALUES FOR TABLE
'ADD A TIMER NAMED HighScoreFlashTimer WITH INTERVAL 100 TO TABLE
'SET HSSSCOREX BELOW TO WHATEVER VARIABLE YOU USE FOR HIGH SCORE.
'ADD OBJECTS TO PLAYFIELD (EASIEST TO JUST COPY FROM THIS TABLE)
'IMPORT POST-IT IMAGES


Dim HSA1, HSA2, HSA3, RSA1, RSA2, RSA3
Dim HSEnterMode, hsLetterFlash, hsEnteredDigits(3), hsCurrentDigit, hsCurrentLetter, RSEnterMode, rsLetterFlash, rsEnteredDigits(3), rsCurrentDigit, rsCurrentLetter
Dim HSArray  
Dim HSScoreM,HSScore100k, HSScore10k, HSScoreK, HSScore100, HSScore10, HSScore1, HSScorex, RSScore100, RSScore10, RSScore1, RSScorex	'Define 6 different score values for each reel to use
HSArray = Array("Postit0","postit1","postit2","postit3","postit4","postit5","postit6","postit7","postit8","postit9","postitBL","postitCM","Tape")
Const hsFlashDelay = 4
Const rsFlashDelay = 4

' ***********************************************************
'  HiScore DISPLAY 
' ***********************************************************

Sub UpdatePostIt
	dim tempscore
	HSScorex = hisc
	TempScore = HSScorex
	HSScore1 = 0
	HSScore10 = 0
	HSScore100 = 0
	HSScoreK = 0
	HSScore10k = 0
	HSScore100k = 0
	HSScoreM = 0
	if len(TempScore) > 0 Then
		HSScore1 = cint(right(Tempscore,1))
	end If
	if len(TempScore) > 1 Then
		TempScore = Left(TempScore,len(TempScore)-1)
		HSScore10 = cint(right(Tempscore,1))
	end If
	if len(TempScore) > 1 Then
		TempScore = Left(TempScore,len(TempScore)-1)
		HSScore100 = cint(right(Tempscore,1))
	end If
	if len(TempScore) > 1 Then
		TempScore = Left(TempScore,len(TempScore)-1)
		HSScoreK = cint(right(Tempscore,1))
	end If
	if len(TempScore) > 1 Then
		TempScore = Left(TempScore,len(TempScore)-1)
		HSScore10k = cint(right(Tempscore,1))
	end If
	if len(TempScore) > 1 Then
		TempScore = Left(TempScore,len(TempScore)-1)
		HSScore100k = cint(right(Tempscore,1))
	end If
	if len(TempScore) > 1 Then
		TempScore = Left(TempScore,len(TempScore)-1)
		HSScoreM = cint(right(Tempscore,1))
	end If
	Pscore6.image = HSArray(HSScoreM):If HSScorex<1000000 Then PScore6.image = HSArray(10)
	Pscore5.image = HSArray(HSScore100K):If HSScorex<100000 Then PScore5.image = HSArray(10)
	PScore4.image = HSArray(HSScore10K):If HSScorex<10000 Then PScore4.image = HSArray(10)
	PScore3.image = HSArray(HSScoreK):If HSScorex<1000 Then PScore3.image = HSArray(10)
	PScore2.image = HSArray(HSScore100):If HSScorex<100 Then PScore2.image = HSArray(10)
	PScore1.image = HSArray(HSScore10):If HSScorex<10 Then PScore1.image = HSArray(10)
	PScore0.image = HSArray(HSScore1):If HSScorex<1 Then PScore0.image = HSArray(10)
	if HSScorex<1000 then
		PComma.image = HSArray(10)
	else
		PComma.image = HSArray(11)
	end if
	if HSScorex<1000000 then
		PComma2.image = HSArray(10)
	else
		PComma2.image = HSArray(11)
	end if

	HSName1.image = ImgFromCode(HSA1, 1)
	HSName2.image = ImgFromCode(HSA2, 2)
	HSName3.image = ImgFromCode(HSA3, 3)

End Sub

Sub UpdatePostIt2
	dim tempscore
	RSScorex = Rhisc
	TempScore = RSScorex
	RSScore1 = 0
	RSScore10 = 0
	RSScore100 = 0

	if len(TempScore) > 0 Then
		RSScore1 = cint(right(Tempscore,1))
	end If
	if len(TempScore) > 1 Then
		TempScore = Left(TempScore,len(TempScore)-1)
		RSScore10 = cint(right(Tempscore,1))
	end If
	if len(TempScore) > 1 Then
		TempScore = Left(TempScore,len(TempScore)-1)
		RSScore100 = cint(right(Tempscore,1))
	end If


	RScore2.image = HSArray(RSScore100):If RSScorex<100 Then RScore2.image = HSArray(10)
	RScore1.image = HSArray(RSScore10):If RSScorex<10 Then RScore1.image = HSArray(10)
	RScore0.image = HSArray(RSScore1):If RSScorex<1 Then RScore0.image = HSArray(10)

	RSName1.image = RImgFromCode(RSA1, 1)
	RSName2.image = RImgFromCode(RSA2, 2)
	RSName3.image = RImgFromCode(RSA3, 3)

End Sub

Function ImgFromCode(code, digit)
	Dim Image
	if (HighScoreFlashTimer.Enabled = True and hsLetterFlash = 1 and digit = hsCurrentLetter) then
		Image = "postitBL"
	elseif (code + ASC("A") - 1) >= ASC("A") and (code + ASC("A") - 1) <= ASC("Z") then
		Image = "postit" & chr(code + ASC("A") - 1)
	elseif code = 27 Then
		Image = "PostitLT"
    elseif code = 0 Then
		image = "PostitSP"
    Else
      msgbox("Unknown display code: " & code)
	end if
	ImgFromCode = Image
End Function

Function RImgFromCode(code, digit)
	Dim Image
	if (HighRoachFlashTimer.Enabled = True and rsLetterFlash = 1 and digit = rsCurrentLetter) then
		Image = "postitBL"
	elseif (code + ASC("A") - 1) >= ASC("A") and (code + ASC("A") - 1) <= ASC("Z") then
		Image = "postit" & chr(code + ASC("A") - 1)
	elseif code = 27 Then
		Image = "PostitLT"
    elseif code = 0 Then
		image = "PostitSP"
    Else
      msgbox("Unknown display code: " & code)
	end if
	RImgFromCode = Image
End Function

Sub HighScoreEntryInit()
	HStimer.uservalue = 0
	HStimer.enabled=1
	HSA1=0:HSA2=0:HSA3=0
	HSEnterMode = True
	hsCurrentDigit = 0
	hsCurrentLetter = 1:HSA1=1
	HighScoreFlashTimer.Interval = 250
	HighScoreFlashTimer.Enabled = True
	hsLetterFlash = hsFlashDelay
End Sub

Sub HighScoreFlashTimer_Timer()
	hsLetterFlash = hsLetterFlash-1
	UpdatePostIt
	If hsLetterFlash=0 then 'switch back
		hsLetterFlash = hsFlashDelay
	end if
End Sub

Sub HighRoachEntryInit()
	RStimer.uservalue = 0
	RStimer.enabled=1
	RSA1=0:RSA2=0:RSA3=0
	RSEnterMode = True
	rsCurrentDigit = 0
	rsCurrentLetter = 1:RSA1=1
	HighRoachFlashTimer.Interval = 250
	HighRoachFlashTimer.Enabled = True
	rsLetterFlash = rsFlashDelay
End Sub

Sub HighRoachFlashTimer_Timer()
	rsLetterFlash = rsLetterFlash-1
	UpdatePostIt2
	If rsLetterFlash=0 then 'switch back
		rsLetterFlash = rsFlashDelay
	end if
End Sub


' ***********************************************************
'  HiScore ENTER INITIALS 
' ***********************************************************

Sub HighScoreProcessKey(keycode)
    If keycode = LeftFlipperKey Then
		hsLetterFlash = hsFlashDelay
		Select Case hsCurrentLetter
			Case 1:
				HSA1=HSA1-1:If HSA1=-1 Then HSA1=26 'no backspace on 1st digit
				UpdatePostIt
			Case 2:
				HSA2=HSA2-1:If HSA2=-1 Then HSA2=27
				UpdatePostIt
			Case 3:
				HSA3=HSA3-1:If HSA3=-1 Then HSA3=27
				UpdatePostIt
		 End Select
    End If

	If keycode = RightFlipperKey Then
		hsLetterFlash = hsFlashDelay
		Select Case hsCurrentLetter
			Case 1:
				HSA1=HSA1+1:If HSA1>26 Then HSA1=0
				UpdatePostIt
			Case 2:
				HSA2=HSA2+1:If HSA2>27 Then HSA2=0
				UpdatePostIt
			Case 3:
				HSA3=HSA3+1:If HSA3>27 Then HSA3=0
				UpdatePostIt
		 End Select
	End If
	
    If keycode = StartGameKey Then
		Select Case hsCurrentLetter
			Case 1:
				hsCurrentLetter=2 'ok to advance
				HSA2=HSA1 'start at same alphabet spot
'				EMReelHSName1.SetValue HSA1:EMReelHSName2.SetValue HSA2
			Case 2:
				If HSA2=27 Then 'bksp
					HSA2=0
					hsCurrentLetter=1
				Else
					hsCurrentLetter=3 'enter it
					HSA3=HSA2 'start at same alphabet spot
				End If
			Case 3:
				If HSA3=27 Then 'bksp
					HSA3=0
					hsCurrentLetter=2
				Else
					savehs 'enter it
					HighScoreFlashTimer.Enabled = False
					HSEnterMode = False
					if hircstate=1 then 
						HighRoachEntryInit
					end if
					
				End If
		End Select
		UpdatePostIt
    End If
End Sub

' ***********************************************************
'  RoachScore ENTER INITIALS 
' ***********************************************************

Sub HighRoachProcessKey(keycode)
    If keycode = LeftFlipperKey Then
		rsLetterFlash = rsFlashDelay
		Select Case rsCurrentLetter
			Case 1:
				RSA1=RSA1-1:If RSA1=-1 Then RSA1=26 'no backspace on 1st digit
				UpdatePostIt2
			Case 2:
				RSA2=RSA2-1:If RSA2=-1 Then RSA2=27
				UpdatePostIt2
			Case 3:
				RSA3=RSA3-1:If RSA3=-1 Then RSA3=27
				UpdatePostIt2
		 End Select
    End If

	If keycode = RightFlipperKey Then
		rsLetterFlash = rsFlashDelay
		Select Case rsCurrentLetter
			Case 1:
				RSA1=RSA1+1:If RSA1>26 Then RSA1=0
				UpdatePostIt2
			Case 2:
				RSA2=RSA2+1:If RSA2>27 Then RSA2=0
				UpdatePostIt2
			Case 3:
				RSA3=RSA3+1:If RSA3>27 Then RSA3=0
				UpdatePostIt2
		 End Select
	End If
	
    If keycode = StartGameKey Then
		Select Case RsCurrentLetter
			Case 1:
				rsCurrentLetter=2 'ok to advance
				RSA2=RSA1 'start at same alphabet spot
'				EMReelHSName1.SetValue HSA1:EMReelHSName2.SetValue HSA2
			Case 2:
				If RSA2=27 Then 'bksp
					RSA2=0
					rsCurrentLetter=1
				Else
					rsCurrentLetter=3 'enter it
					RSA3=RSA2 'start at same alphabet spot
				End If
			Case 3:
				If RSA3=27 Then 'bksp
					RSA3=0
					rsCurrentLetter=2
				Else
					savehs 'enter it
					HighRoachFlashTimer.Enabled = False
					RSEnterMode = False
				End If
		End Select
		UpdatePostIt2
    End If
End Sub
'	-------------------------------------------------
'				 .               .    .          .              .   .         .
'             .               .    .          .              .   .         .
'               _________________      ____         __________
' .       .    /                 |    /    \    .  |          \
'     .       /    ______   _____| . /      \      |    ___    |     .     .
'             \    \    |   |       /   /\   \     |   |___>   |
'           .  \    \   |   |      /   /__\   \  . |         _/             .
' .     ________>    |  |   | .   /            \   |   |\    \_______    .
'      |            /   |   |    /    ______    \  |   | \           |
'      |___________/    |___|   /____/      \____\ |___|  \__________|    .
'  .     ____    __  . _____   ____      .  __________   .  _________
'       \    \  /  \  /    /  /    \       |          \    /         |      .
'        \    \/    \/    /  /      \      |    ___    |  /    ______|  .
'         \              /  /   /\   \ .   |   |___>   |  \    \
'   .      \            /  /   /__\   \    |         _/.   \    \
'           \    /\    /  /            \   |   |\    \______>    |   .
'            \  /  \  /  /    ______    \  |   | \              /          .
' .       .   \/    \/  /____/      \____\ |___|  \____________/  LS
'                               .
'                                           .               .
'	-------------------------------------------------					.                                   .            .
'	STAR WARS Gottlieb 1978 (well really.. BorgDog 2016)
'	-------------------------------------------------
'
'	thanks to hauntfreaks for graphics help and primitives, and testing
'	thanks to sliderpoint for game play ideas, testing, and primitives
'	thanks to dark for the slotted screw primitive (and possibly others)
'	thanks to GTXjoe by way of mfuegemann's Fast Draw table for the high score code
'	thanks to nFozzy for the flipper "tricks" code
'	thanks to arngrim for checking out the dof and making a couple additions and adding to the config database
'
'	Many options are available.  Hold the left flipper before starting a game to get to options menu and scroll down for a few in the script.
'
'	Instructions:
'
'	Complete three missions to activate the tractor beam.
'		1. Advance Lukes Jedi training by hitting the labeled targets.
'		2. Retrieve the Death Star plans by clearing the top rollovers.
'		3. Rescue the princess by blasting down the drop targets.
'
'	Once the tractor beam is active complete a trench run to
'		  lock a ball in the Death Star.
'
'	Release the tractor beam by hitting the labeled hole in R2-D2,
'		  starting Death Star multi-ball, opening the portal, and
'		  resetting the missions.
'
'	Hit the portal within the multiballs lives to earn extra ball, or the portal closes again..
'
'
'	-------------------------------------------------
'	Layer usage
'		1 - most stuff
'		2 - triggers
'		4 - options menu
'		5 - stuff I don't want to see/turn off first eg shadow ramp
'		6 - GI lighting
'		7 - plastics
'		8 - insert and bumper lighting
'
'	DOF config
'		101 Left Flipper,
'		102 Right Flipper,
'		103 Left sling,
'		104 Left sling flasher,
'
'		107 Bumper1,
'		108 Bumper1 flasher,
'		109 bumper2,
'		110 bumper2 Flasher
'		111 Drop Target Reset,
'		112 Drop Targets Hit
'		113 upper death star roto targets
'		114 lower r2d2 targets
'		115 Death Star kickers
'		116 R2D2 kicker
'		117 top left rollover
'		118 top mid rollover
'		119 top right rollover
'		120 right side rollover
'		121 left outlane rollover
'		122 left inlane rollover
'		123 right outlane rollover
'		124 right inlane rollover
'
'		133 Bell,
'		134 Drain,
'		135 Ball Release
'		136 Shooter Lane/launch ball,
'		137 credit light,
'		138 knocker,
'		139 knocker Flasher
'		141 Chime1-10s,
'		142 Chime2-100s,
'		143 Chime3-1000s
'
'	-------------------------------------------------

' Thalamus 2018-07-24
' Added/Updated "Positional Sound Playback Functions" and "Supporting Ball & Sound Functions"
' No special SSF tweaks yet.


Option Explicit
Randomize

Const cGameName = "StarWars_1978"

' **********************   a few more OPTIONS

Dim DTshadow: DTshadow=0.3    	'set drop target shadow level from 0 to 1, smaller number is more shadow
Dim attractmode: attractmode=1	'set to 0 for no atractmode animation, 1 to animate lights in attractmode
Dim nfozzy: nfozzy=0			'set to 0 for standard vpx flipper physics or 1 for nfozzys adjustments to rate of return
dim showhiscnames: showhiscnames=1		'set to 0 for no high score initials, 1 for initials
dim chimevol: chimevol=1		'set between 0 and 1 to adjust how loud the chimes are, 0 is off, go below 0.1 for noticeable effect

' **********************

Dim operatormenu, options
Dim bumperlitscore, bumperoffscore
Dim balls, replays, replaymode
Dim Replay1Table(3), Replay2Table(3), Replay3Table(3)
dim replay1, replay2, replay3
dim replaytext(3)
Dim Add10, Add100, Add1000
Dim hisc, showhisc, HSA1, HSA2, HSA3
Dim HSEnterMode, hsLetterFlash, hsEnteredDigits(3), hsCurrentDigit, hsCurrentLetter
Dim credit, freeplay, thirdflipper
Dim score
Dim state, eg
Dim tilt, tiltsens
Dim ballinplay
Dim matchnumb
dim ballrenabled
dim lstep, stepa, stepb, stepc, stepd, stepe
Dim dsstep, ds1step, droidstep
Dim rep
Dim Drain1Ball, DrainBall, justlocked, multiball
Dim princess, plans, jedi
Dim bell, chimemarch
Dim i,j, ii, objekt, light
Dim awardcheck, advance

Const hsFlashDelay = 4



'*************nFozzy flipper routine part 1
dim returnspeed, lfstep, rfstep
returnspeed = leftflipper.return
lfstep = 1
rfstep = 1

sub leftflipper_timer()
	select case lfstep
		Case 1: leftflipper.return = returnspeed * 0.6 :lfstep = lfstep + 1
		Case 2: leftflipper.return = returnspeed * 0.7 :lfstep = lfstep + 1
		Case 3: leftflipper.return = returnspeed * 0.8 :lfstep = lfstep + 1
		Case 4: leftflipper.return = returnspeed * 0.9 :lfstep = lfstep + 1
		Case 5: leftflipper.return = returnspeed * 1 :lfstep = lfstep + 1
		Case 6: leftflipper.timerenabled = 0 : lfstep = 1
	end select
end sub

sub rightflipper_timer()
	select case rfstep
		Case 1: rightflipper.return = returnspeed * 0.6 :rfstep = rfstep + 1
		Case 2: rightflipper.return = returnspeed * 0.7 :rfstep = rfstep + 1
		Case 3: rightflipper.return = returnspeed * 0.8 :rfstep = rfstep + 1
		Case 4: rightflipper.return = returnspeed * 0.9 :rfstep = rfstep + 1
		Case 5: rightflipper.return = returnspeed * 1 :rfstep = rfstep + 1
		Case 6: rightflipper.timerenabled = 0 : rfstep = 1
	end select
end sub

'*************nFozzy flipper routine part 1 end

Dim HSArray
Dim HSScoreM,HSScore100k, HSScore10k, HSScoreK, HSScore100, HSScore10, HSScore1, HSScorex	'Define 6 different score values for each reel to use
HSArray = Array("Postit0","postit1","postit2","postit3","postit4","postit5","postit6","postit7","postit8","postit9","postitBL","postitCM","Tape")

On Error Resume Next
ExecuteGlobal GetTextFile("controller.vbs")
If Err Then MsgBox "Can't open controller.vbs"
On Error Goto 0

sub StarWars_init
	LoadEM
	Replay1Table(1)=50000
	Replay2Table(1)=66000
'	Replay3Table(1)=75000
	Replay1Table(2)=60000
	Replay2Table(2)=78000
'	Replay3Table(2)=92000
	Replay1Table(3)=70000
	Replay2Table(3)=92000
'	Replay3Table(3)=109000
	hideoptions
	if attractmode=1 then attract.enabled=1
	loadhs
	if hisc="" then hisc=50000
	if HSA1="" then HSA1=25
	if HSA2="" then HSA2=25
	if HSA3="" then HSA3=25
	gamov.text="Game Over"
	gamov.timerenabled=1
	tilttxt.timerenabled=1
	credittxt.setvalue(credit)
	if thirdflipper<0 or thirdflipper>1 or thirdflipper="" then thirdflipper=1
	if freeplay<0 or freeplay>1 or freeplay="" then freeplay=0
	if showhisc<0 or showhisc>1 or showhisc="" then showhisc=1
	if showhisc=0 then
		for each objekt in hiscsticky:objekt.visible=0:Next
	  Else
		for each objekt in hiscsticky:objekt.visible=1:Next
	end if
	if replaymode<0 or replaymode>1 or replaymode="" then replaymode=0
	if balls="" then balls=5
	if balls<>3 and balls<>5 then balls=5
	if replays="" then replays=3
	if replays<>1 and replays<>2 and replays<>3 then replays=3
	Replay1=Replay1Table(Replays)
	Replay2=Replay2Table(Replays)
	Replay3=Replay3Table(Replays)
	If replaymode=1 then
		OptionReplayMode.image = "OptionsYN0"
		RepCard.image = "ReplayCard"&replays
	  Else
		OptionReplayMode.image = "OptionsYN1"
		RepCard.image = "extraballcard"&replays
	End if
	OptionBalls.image="OptionsBalls"&Balls
	OptionReplays.image="OptionsReplays"&replays
	OptionFree.image="OptionsYN"&freeplay
	OptionHiSc.image="OptionsYN"&showhisc
	OptionFlipper.image="OptionsYN"&thirdflipper
	bumperlitscore=100
	bumperoffscore=100
	if balls=3 then
		InstCard.image="InstCard3balls"
	  else
		InstCard.image="InstCard5balls"
	end if
	If B2SOn then
		setBackglass.enabled=true
		for each objekt in backdropstuff : objekt.visible = 0 : next
	End If
	For each light in GIlights:light.state=1:next
	if matchnumb="" then matchnumb=10
	if matchnumb=0 then
		matchtxt.text="00"
	  else
		matchtxt.text=matchnumb*10
	end if
	ScoreReel1.setvalue(score)
	turnoff
	PlaySound "motorshort1s"
	tilt=false
	If credit>0 or freeplay=1 then DOF 137, DOFOn
    Drain.CreateBall
	Drain1.CreateBall
	Drain1Ball=1
	DrainBall=1
	UpdatePostIt
'	chimemarch=10
'	deathstar.timerenabled=1
End sub

sub setBackglass_timer
	Controller.B2SSetScorePlayer1 score
	Controller.B2ssetCredits Credit
	Controller.B2ssetMatch 34, Matchnumb
    Controller.B2SSetGameOver 35,1
	Controller.B2SSetScoreRollover 25, Int(score/100000)
	me.enabled=0
end Sub

sub Attract_timer
	attractm.Play SeqUpOn,15,1
	attractm.Play SeqDownOn,15,1
	attractm.Play SeqRightOn,15,1
	attractm.Play SeqLeftOn,15,1
	me.enabled=0
end sub

Sub Attractm_PlayDone()
	attract.enabled=1
End Sub

sub DeathStar_timer
	Select Case chimemarch
		Case 20:
			PlaySound SoundFXDOF("bell100",142,DOFPulse,DOFChimes), 0, chimevol
		Case 22:
			PlaySound SoundFXDOF("bell100",142,DOFPulse,DOFChimes), 0, chimevol
		Case 24:
			PlaySound SoundFXDOF("bell100",142,DOFPulse,DOFChimes), 0, chimevol
		Case 26:
			PlaySound SoundFXDOF("bell10",141,DOFPulse,DOFChimes), 0, chimevol
		Case 27:
			PlaySound SoundFXDOF("bell1000",143,DOFPulse,DOFChimes), 0, chimevol
		Case 28:
			PlaySound SoundFXDOF("bell100",142,DOFPulse,DOFChimes), 0, chimevol
		Case 30:
			PlaySound SoundFXDOF("bell10",141,DOFPulse,DOFChimes), 0, chimevol
		Case 31:
			PlaySound SoundFXDOF("bell1000",143,DOFPulse,DOFChimes), 0, chimevol
		Case 32:
			PlaySound SoundFXDOF("bell100",142,DOFPulse,DOFChimes), 0, chimevol
		Case 33:
			me.timerenabled=0
	end select
	chimemarch=chimemarch+1
end sub


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
	if showhisc=1 and showhiscnames=1 then
		for each objekt in hiscname:objekt.visible=1:next
		HSName1.image = ImgFromCode(HSA1, 1)
		HSName2.image = ImgFromCode(HSA2, 2)
		HSName3.image = ImgFromCode(HSA3, 3)
	  else
		for each objekt in hiscname:objekt.visible=0:next
	end if
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

Sub HighScoreEntryInit()
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


' ***********************************************************
'  HS ENTER INITIALS
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
				End If
		End Select
		UpdatePostIt
    End If
End Sub

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
		gamov.timerenabled=1
		gamov.timerinterval= (INT (RND*10)+5)*100
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

Sub StarWars_KeyDown(ByVal keycode)

	if keycode=AddCreditKey then
		playsound "coinin"
		coindelay.enabled=true
    end if

    if keycode=StartGameKey and (credit>0 or freeplay=1) and (drainball=1 and drain1ball=1) And Not HSEnterMode=true then
	  if state=false then
		if freeplay=0 then credit=credit-1
		if credit < 1 and freeplay=0 then DOF 137, DOFOff
'		playsound "cluper"
		credittxt.setvalue(credit)
		ballinplay=1
		If B2SOn Then
			Controller.B2ssetCredits Credit
			Controller.B2sStartAnimation "Startup"
			Controller.B2ssetballinplay 32, Ballinplay
			Controller.B2ssetplayerup 30, 1
			Controller.B2ssetcanplay 31, 1
			Controller.B2SSetGameOver 0
			Controller.B2SSetScorePlayer 5, hisc
		End If
	    Pup1.state=1
		tilt=false
		state=true
		playsound "initialize"
		newgame.enabled=true
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
		If Options=8 then Options=1
		playsound "target"
		Select Case (Options)
			Case 1:
				Option1.visible=true
				Option7.visible=False
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
			Case 6:
				Option6.visible=true
				Option5.visible=False
			Case 7:
				Option7.visible=true
				Option6.visible=False
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
			OptionFree.image = "OptionsYN"&freeplay
		Case 3:
			if showhisc=0 Then
				showhisc=1
				for each objekt in hiscsticky:objekt.visible=true:next
			  Else
				showhisc=0
				for each objekt in hiscsticky:objekt.visible=False:next
			end if
			OptionHiSc.image = "OptionsYN"&showhisc
		Case 4:
			if thirdflipper=0 Then
				thirdflipper=1
			  Else
				thirdflipper=0
			end if
			OptionFlipper.image = "OptionsYN"&thirdflipper
		Case 5:
			if replaymode=0 Then
				replaymode=1
			  Else
				replaymode=0
			end if
			If replaymode=1 then
				OptionReplayMode.image = "OptionsYN0"
				RepCard.image = "ReplayCard"&replays
			  Else
				OptionReplayMode.image = "OptionsYN1"
				RepCard.image = "extraballcard"&replays
			End if
 		Case 6:
			Replays=Replays+1
			if Replays>3 then
				Replays=1
			end if
			Replay1=Replay1Table(Replays)
			Replay2=Replay2Table(Replays)
			replay3=Replay3Table(Replays)
			OptionReplays.image = "OptionsReplays"&replays
			if replaymode=1 Then
				RepCard.image = "ReplayCard"&replays
			  Else
				RepCard.image = "extraballcard"&replays
			end if
		Case 7:
			OperatorMenu=0
			savehs
			HideOptions
	  End Select
	End If

    If HSEnterMode Then HighScoreProcessKey(keycode)

  if tilt=false and state=true then
	If keycode = LeftFlipperKey Then
		LeftFlipper.RotateToEnd
		PlaySound SoundFXDOF("flipperup",101,DOFOn,DOFContactors), 0, .67, -0.05, 0.05
		PlaySound "Buzz",-1,.05,-0.05, 0.05
	End If

	If keycode = RightFlipperKey Then
		RightFlipper.RotateToEnd
		if thirdflipper = 1 then UpperFlipper.RotateToEnd
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
		gametilted
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
	OptionFree.visible = True
	OptionHiSc.visible = True
	OptionFlipper.visible = True
	OptionReplayMode.visible = True
    OptionReplays.visible = True
End Sub

Sub HideOptions
	for each objekt In OptionMenu
		objekt.visible = false
	next
End Sub


Sub StarWars_KeyUp(ByVal keycode)

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
'*************nFozzy flipper routine part 2
		if nfozzy=1 then
			LeftFlipper.TimerEnabled = 1
			LeftFlipper.TimerInterval = 16
			LeftFlipper.return = returnspeed * 0.5
		end if
'*************nFozzy flipper routine part 2 end
		PlaySound SoundFXDOF("flipperdown",101,DOFOff,DOFContactors), 0, 1, -0.05, 0.05
		StopSound "Buzz"
	End If

	If keycode = RightFlipperKey Then
		RightFlipper.RotateToStart
'*************nFozzy flipper routine part 2b
		if nfozzy=1 then
			RightFlipper.TimerEnabled = 1
			RightFlipper.TimerInterval = 16
			RightFlipper.return = returnspeed * 0.5
		end if
'*************nFozzy flipper routine part 2b end
		if thirdflipper = 1 then UpperFlipper.RotateToStart
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
	LrotoL.state=Ldroid2.state
	LrotoR.state=Ldroid1.state
	Ldroid.state=Ldroid1.state
	LinL.state=Ldroid2.state
	LinR.state=Ldroid1.state
	if Ltractorbeam.state=1 then
		LoutL.state=LinR.state
		LoutR.state=LinL.state
	 Else
		LoutL.state=0
		LoutR.state=0
	end if
end sub

sub coindelay_timer
	addcredit
    coindelay.enabled=false
end sub

Sub addcredit
      if freeplay=0 then credit=credit+1
	  DOF 137, DOFOn
      if credit>25 then credit=25
	  credittxt.setvalue(credit)
	  If B2SOn Then Controller.B2ssetCredits Credit
End sub

sub resetmissions
	ResetDT
	For each light in TopLights: light.state=1:Next
	For each light in DTLights:light.state=1:Next
	For each light in AdvanceLights:light.state=0:Next
	For each light in ResetLights:light.state=0:Next
	advance=0
	plans=0
	princess=0
	jedi=0
end sub

'************************Drain/Trough

Sub Drain_Hit()
	DrainBall=1
	if state=True then
		DOF 134, DOFPulse
		PlaySound "drain",0,1,0,0.25
	End if
	if eg=0 then me.timerenabled=1
End Sub

Sub Drain_timer
	If Drain1ball=1 Then
		if multiball=0 then
			if RotoTarget.roty<>0 then
				LDShole1.state=0
				TrotoR.timerenabled=1
				PlaySound "motor"
'				resetmissions
			end if
			if shootagain.state=lightstateon and tilt=false then
				ballreltimer.enabled=true
			  else
				nextball
			end if
		  else
			multiball=multiball-1
		end if
	  me.timerenabled=0
	end if
End Sub

sub Drain1_hit
	Drain1ball=1
end sub

sub Drain1_timer
	If Drain1Ball=0 and DrainBall=1 Then
		Drain.kick 60,55,5
		DrainBall=0
	End If
end sub

sub ballhome_hit
	ballrenabled=1
end sub

sub ballhome_unhit
	DOF 136, DOFPulse
end sub

sub ballrel_hit
	if ballrenabled=1 then
		if justlocked=0 then shootagain.state=lightstateoff
		justlocked=0
		Lplunger.state=0
		ballrenabled=0
	end if
end sub

sub newgame_timer
	Attractm.StopPlay()
	attract.enabled=0
	score=0
	rep=0
	ScoreReel1.resettozero
	for i=1 to 2: EVAL("Bumper"&i).hashitevent = 1: Next
	LeftSlingShot.disabled=False
	P100K.text=""
	pup1.state=1
	if Ldroid1.state=1 Then
		Ldroid1.state=0
		Ldroid2.state=1
	  Else
		Ldroid1.state=1
		Ldroid2.state=0
	end if
    eg=0
	For each light in GIlights:light.state=1:next
	resetmissions
	shootagain.state=0
    tilttxt.text=" "
	gamov.text=" "
    If B2SOn then
		Controller.B2SSetScorePlayer 1, score
		Controller.B2SSetGameOver 35,0
		Controller.B2SSetTilt 33,0
		Controller.B2SSetMatch 34,0
		Controller.B2SSetScoreRollover 25, 0
	End If
    biptext.text="1"
	matchtxt.text=" "
	ResetDT
	ballrelTimer.enabled=true
	newgame.enabled=false
end sub


sub nextball
    if tilt=true then
	  for i=1 to 2
		EVAL("Bumper"&i).hashitevent = 1
	  Next
      tilt=false
      tilttxt.text=" "
		If B2SOn then
			Controller.B2SSetTilt 33,0
			Controller.B2ssetdata 1, 1
		End If
		LeftSlingShot.disabled=False
    end if
	ballinplay=ballinplay+1
	if ballinplay>balls then
		if state=true then playsound "GameOver"
		eg=1
		ballreltimer.enabled=true
	  else
		if state=true then
			ballreltimer.enabled=true
		  else
			biptext.text=ballinplay
			If B2SOn then Controller.B2ssetballinplay 32, Ballinplay
		end if
	end if
End Sub

sub ballreltimer_timer
  if eg=1 then
	if state=true Then
	  turnoff
	  if Ldroidlock.state=1 then DShole.kick 70,5,0
      matchnum
	  state=false
	  biptext.text=" "
	  gamov.text="GAME OVER"
	  gamov.timerenabled=1
	  tilttxt.timerenabled=1
	  if score>hisc then
		hisc=score
		if showhiscnames=1 then HighScoreEntryInit()
	  end if
	  Pup1.state=0
	  UpdatePostIt
	  savehs
	  If B2SOn then
        Controller.B2SSetGameOver 35,1
        Controller.B2ssetballinplay 32, 0
	    Controller.B2SSetScorePlayer 5, hisc
	    Controller.B2ssetcanplay 31, 0
	    Controller.B2ssetcanplay 30, 0
	  End If
	  chimemarch=15
	  deathstar.timerenabled=1
	  For each light in GIlights:light.state=0:next
	  if attractmode=1 then attract.enabled=1
	  ballreltimer.enabled=false
	end if
  else
	playsound SoundFXDOF("kickerkick",135,DOFPulse,DOFContactors)
	biptext.text=ballinplay
	If B2SOn then Controller.B2ssetballinplay 32, Ballinplay
	Drain1.kick 60,5,0
	Drain1Ball=0
    ballreltimer.enabled=false
  end if
end sub

sub matchnum
	if matchnumb=0 then
		matchtxt.text="00"
		If B2SOn then Controller.B2SSetMatch 100
	  else
		matchtxt.text=matchnumb*10
		If B2SOn then Controller.B2SSetMatch 34,Matchnumb*10
	end if
	if (matchnumb*10)=(score mod 100) then
		addcredit
		playsound SoundFXDOF("knock",138,DOFPulse,DOFKnocker)
		DOF 139,DOFPulse
	end if
end sub

'********** Bumpers

Sub Bumper1_Hit
   if tilt=false then
	playsound SoundFXDOF("fx_bumper4",107,DOFPulse,DOFContactors)
	DOF 108,DOFPulse
	FlashBumpers
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
	playsound SoundFXDOF("fx_bumper4",109,DOFPulse,DOFContactors)
	DOF 110,DOFPulse
	FlashBumpers
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

sub FlashBumpers
	FlashB.StopPlay
	FlashB.Play SeqBlinking,,1,25
end sub


'************** Slings


Sub LeftSlingShot_Slingshot
	If tilt=false and state=true then
		playsound SoundFXDOF("left_slingshot",103,DOFPulse,DOFContactors), 0, 1, -0.05, 0.05
		DOF 104,DOFPulse
		addscore 10
		LSling.Visible = 0
		LSling1.Visible = 1
		slingL.objroty = 15
		LStep = 1
		LeftSlingShot.TimerEnabled = 1
	end if
End Sub

Sub LeftSlingShot_Timer
    Select Case LStep
        Case 3:LSLing1.Visible = 0:LSLing2.Visible = 1:slingL.objroty = 7
        Case 4:slingL.objroty = 0:LSLing2.Visible = 0:LSLing.Visible = 1:LeftSlingShot.TimerEnabled = 0
    End Select
    LStep = LStep + 1
End Sub

'******************* Dingwalls (scoring rubbers)

Sub Dingwalls_hit(idx)
	addscore 10
End sub

Sub DingwallA_hit
	stepa=1
	RDWa.visible=0
	rdwa1.visible=1
	me.timerenabled=1
end sub

Sub Dingwalla_timer
	select case stepa
		Case 1: rdwa1.visible=0: rdwa.visible=1
		Case 2: rdwa.visible=0: rdwa2.visible=1
		Case 3: rdwa2.visible=0: rdwa.visible=1: me.timerenabled=0
	end Select
	stepa=stepa+1
end sub

Sub DingwallB_hit
	stepb=1
	RDWb.visible=0
	rdwb1.visible=1
	me.timerenabled=1
end sub

Sub Dingwallb_timer
	select case stepb
		Case 1: rdwb1.visible=0: rdwb.visible=1
		Case 2: RDWB.visible=0: rdwb2.visible=1
		Case 3: rdwb2.visible=0: rdwb.visible=1: me.timerenabled=0
	end Select
	stepb=stepb+1
end sub

Sub DingwallC_hit
	stepc=1
	RDWC.visible=0
	rdwc1.visible=1
	me.timerenabled=1
end sub

Sub Dingwallc_timer
	select case stepc
		Case 1: rdwc1.visible=0: RDWC.visible=1
		Case 2: RDWC.visible=0: rdwc2.visible=1
		Case 3: rdwc2.visible=0: RDWC.visible=1: me.timerenabled=0
	end Select
	stepc=stepc+1
end sub

Sub DingwallD_hit
	stepD=1
	RDWD.visible=0
	rdwD1.visible=1
	me.timerenabled=1
end sub

Sub DingwallD_timer
	select case stepD
		Case 1: rdwD1.visible=0: RDWD.visible=1
		Case 2: RDWD.visible=0: rdwD2.visible=1
		Case 3: rdwD2.visible=0: RDWD.visible=1: me.timerenabled=0
	end Select
	stepD=stepD+1
end sub

Sub DingwallE_hit
	stepE=1
	RDWE.visible=0
	rdwE1.visible=1
	me.timerenabled=1
end sub

Sub DingwallE_timer
	select case stepE
		Case 1: rdwE1.visible=0: RDWE.visible=1
		Case 2: RDWE.visible=0: rdwE2.visible=1
		Case 3: rdwE2.visible=0: RDWE.visible=1: me.timerenabled=0
	end Select
	stepE=stepE+1
end sub

'********** Triggers

sub TGoutL_hit
	addscore 500
	DOF 121, DOFPulse
    if (LoutL.state)=lightstateon then
		ShootAgain.state=1
	end if
end sub

sub TGinL_hit
	DOF 122, DOFPulse
    if (LinL.state)=lightstateon then
		addscore 500
    else
		addscore 50
    end if
end sub

sub TGoutR_hit
	addscore 500
	DOF 123, DOFPulse
    if (LoutR.state)=lightstateon then
		ShootAgain.state=1
	end if
end sub

sub TGinR_hit
	DOF 124, DOFPulse
    if (LinR.state)=lightstateon then
		addscore 500
    else
		addscore 50
    end if
end sub

sub TGtopL_hit
	DOF 117, DOFPulse
	if (LtopL.state)=lightstateon then
		addscore 3000
		LtopL.state=0
    else
		addscore 300
    end if
	CheckAwards
 end sub

sub TGtopC_hit
	DOF 118, DOFPulse
	if (LtopC.state)=lightstateon then
		addscore 3000
		LtopC.state=0
      else
		addscore 300
    end if
	CheckAwards
end sub

sub TGtopR_hit
	DOF 119, DOFPulse
    if (LtopR.state)=lightstateon then
		addscore 1000
		LtopR.state=0
    else
		addscore 100
    end if
	CheckAwards
end sub

sub TGdroid_hit
	DOF 120, DOFPulse
	if Ldroid.state=1 Then
		addscore 5000
		if advance<20 then addadvance
	  Else
		addscore 500
	end if
end sub

'************** Drop Targets

Sub DropTargets_hit(idx)
	playsound SoundFXDOF("drop1",112,DOFPulse,DOFContactors), 0, .05, -0.05, 0.05
End sub

Sub DT1_dropped
	If Ldt1.state=1 Then
		addscore 2000
		Ldt1.state=0
	  Else
		addscore 200
	end if
	LDT1s.intensityscale=1
	LDT1s1.intensityscale=1
	CheckAwards
end sub

Sub DT2_dropped
	If LDT2.state=1 Then
		addscore 2000
		LDT2.state=0
	  Else
		addscore 200
	end if
	LDT2s.intensityscale=1
	CheckAwards
end sub

Sub DT3_dropped
	If LDT3.state=1 Then
		addscore 2000
		LDT3.state=0
	  Else
		addscore 200
	end if
	LDT3s.intensityscale=1
	CheckAwards
end sub

Sub DT4_dropped
	If LDT4.state=1 Then
		addscore 2000
		LDT4.state=0
	  Else
		addscore 200
	end if
	LDT4s.intensityscale=1
	CheckAwards
end sub

'********************* Targets

sub TrotoL_hit
	playsound SoundFXDOF("target",113,DOFPulse,DOFContactors), 0, .05, 0, 0.05
	if LrotoL.state=1 Then
		addadvance
		addscore 4000
	  Else
		addscore 400
	end if
	CheckAwards
end sub

sub TrotoR_hit
	playsound SoundFXDOF("target",113,DOFPulse,DOFContactors), 0, .05, 0, 0.05
	if LrotoR.state=1 Then
		addadvance
		addscore 4000
	  Else
		addscore 400
	end if
	CheckAwards
end sub

sub Tdroid1_hit
	playsound SoundFXDOF("target",114,DOFPulse,DOFContactors), 0, .05, 0.05, 0.05
	if Ldroid1.state=1 Then if advance<20 then addadvance
	addscore 200
end sub

sub Tdroid2_hit
	playsound SoundFXDOF("target",114,DOFPulse,DOFContactors), 0, .05, 0.05, 0.05
	if Ldroid2.state=1 Then if advance<20 then addadvance
	addscore 200
end sub

'*****************Kickers

Sub DShole_Hit()
	PlaySound "cluper"
	if LTractorBeam.state=1 then
		justlocked=1
		addscore 5000
		Ldroidlock.state=1
		Lplunger.state=2
		Drain1.kick 60,5,0
		Drain1Ball=0
	  Else
		addscore 500
		DSstep=1
		me.timerenabled=1
	end if
End Sub

Sub DShole_timer
	select case dsstep
	  case 4:
		playsound SoundFXDOF("holekick",115,DOFPulse,DOFContactors)
		DOF 139, DOFPulse
		DShole.kick 70,12,0
		PkickarmDS.rotz=10
	  case 5:
		PkickarmDS.rotz=0
		me.timerenabled=0
	end Select
	dsstep=dsstep+1
End Sub

Sub DSHole1_hit
	playsound SoundFXDOF("dabell",133,DOFPulse,DOFBell)
	ShootAgain.state=1
	addscore 5000
	ds1step=1
	me.timerenabled=1
end sub

Sub DShole1_timer
	select case ds1step
	  case 2:
		playsound SoundFXDOF("holekick",115,DOFPulse,DOFContactors)
		DOF 139, DOFPulse
		DShole1.kick 140,20,0
		LDShole1.state=0
		resetroto.enabled=1
		PkickarmDS1.rotz=10
	  case 3:
		PkickarmDS1.rotz=0
		me.timerenabled=0
	end Select
	ds1step=ds1step+1
End Sub

Sub resetroto_timer
	TrotoR.timerenabled=1
	me.enabled=0
end sub

Sub DroidHole_Hit()
	PlaySound "cluper"
	if ldroidlock.state=1 Then
		addscore 500
		dsstep=1
		dshole.timerenabled=1
		resetmissions
		LDShole1.state=2
		Ldroidlock.state=2
		TrotoL.timerenabled=1
		PlaySound "motor"
		multiball=1
		me.timerinterval=2000
		droidstep=1
		me.timerenabled=1
	  Else
		me.timerinterval=200
		addscore 50
		droidstep=-1
		me.timerenabled=1
	end if

End Sub

Sub DroidHole_timer
	select case droidstep
	  case 2:
		playsound SoundFXDOF("holekick",116,DOFPulse,DOFContactors)
		DOF 139, DOFPulse
		DroidHole.kick 200,8,0
		if ldroidlock.state=2 then ldroidlock.state=0
		Pkickarm.rotz=10
		me.timerinterval=200
	  case 3:
		Pkickarm.rotz=0
		me.timerenabled=0
	end Select
	droidstep=droidstep+1
End Sub

Sub TrotoL_timer
	Lrotoflash.state=2
	Lrotoflash1.state=2
	Lroto.state=0
	Lroto1.state=0
	if rototarget.roty=-202.5 Then
		Lrotoflash.state=0
		Lrotoflash1.state=0
		Lroto.state=1
		Lroto1.state=1
		StopSound "motor"
		TrotoL.isdropped=True
		TrotoR.isdropped=True
		me.timerenabled=0
	  Else
		RotoTarget.roty=RotoTarget.roty-7.5
	end if
end sub

Sub TrotoR_timer
	Lrotoflash.state=2
	Lrotoflash1.state=2
	Lroto.state=0
	Lroto1.state=0
	if rototarget.roty=0 or rototarget.roty=-360 Then
		Lrotoflash.state=0
		Lrotoflash1.state=0
		Lroto.state=1
		Lroto1.state=1
		rototarget.roty=0
		me.timerenabled=0
		StopSound "motor"
	  Else
		RotoTarget.roty=RotoTarget.roty-7.5
	end if
	TrotoL.isdropped=false
	TrotoR.isdropped=False
end sub

'**********************Check awards

Sub CheckAwards
	awardcheck=LtopL.state+LtopC.state+LtopR.state
	If awardcheck = 0 and plans=0 Then
		If Lyellowalert.state=0 Then
			lyellowalert.state=1
		 else if lredalert.state=0 Then
			lredalert.state=1
		   else if LTractorBeam.state=0 Then
			LTractorBeam.state=1
		   end if
		 end if
		end if
		plans=1
	End If
	awardcheck=DT1.isdropped+dt2.isdropped+dt3.isdropped+dt4.isdropped
	If awardcheck = 4 and princess=0 Then
		If Lyellowalert.state=0 Then
			lyellowalert.state=1
		 else if lredalert.state=0 Then
			lredalert.state=1
		   else if LTractorBeam.state=0 Then
			LTractorBeam.state=1
		   end if
		 end if
		end if
		princess=1
	end if
	If awardcheck=4 then DT1.timerenabled=1
End Sub

Sub DT1_timer
	ResetDT
	me.timerenabled=0
end sub

Sub ResetDT
	playsound SoundFXDOF("BankReset",111,DOFPulse,DOFContactors), 0, .67, -0.05
	for each objekt in droptargets: objekt.isdropped=false: Next
	for each objekt in DTshadows: objekt.intensityscale=DTshadow: Next
End Sub

Sub addadvance
	if advance<6 then
		advance=advance+1
		Eval("advance"&advance).state=1
	end if
	if advance=6 and jedi=0 Then
		Advance6.state=1
		If Lyellowalert.state=0 Then
			lyellowalert.state=1
		 else if lredalert.state=0 Then
			lredalert.state=1
		   else if LTractorBeam.state=0 Then
			LTractorBeam.state=1
		   end if
		 end if
		end if
		jedi=1
	end if
end sub


'**************Scoring

sub addscore(points)
  if tilt=false and state=True then
	If points = 10 then
		matchnumb=matchnumb+1
		if matchnumb>9 then matchnumb=0
		if Ldroid1.state=1 Then
			ldroid2.state=1
			ldroid1.state=0
		   Else
			ldroid2.state=0
			ldroid1.state=1
		end if
	end if
	if points=10 or points=100 or points=1000 then
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

Sub AddPoints(Points)
    score=score+points
	ScoreReel1.addvalue(points)
	If B2SOn Then Controller.B2SSetScorePlayer 1, score

    ' Sounds: there are 3 sounds: tens, hundreds and thousands
    If Points = 100 AND(Score MOD 1000) \ 100 = 0 Then  'New 1000 reel
        PlaySound SoundFXDOF("bell1000",143,DOFPulse,DOFChimes), 0, chimevol
      ElseIf Points = 10 AND(Score MOD 100) \ 10 = 0 Then 'New 100 reel
        PlaySound SoundFXDOF("bell100",142,DOFPulse,DOFChimes), 0, chimevol
      ElseIf points = 1000 Then
        PlaySound SoundFXDOF("bell1000",143,DOFPulse,DOFChimes), 0, chimevol
	  elseif Points = 100 Then
        PlaySound SoundFXDOF("bell100",142,DOFPulse,DOFChimes), 0, chimevol
      Else
        PlaySound SoundFXDOF("bell10",141,DOFPulse,DOFChimes), 0, chimevol
    End If
    ' check replays and rollover
	if score>99999 then
		If B2SOn then Controller.B2SSetScoreRollover 25, 1
		P100K.text="100,000"
		If score>199999 Then
			If B2SOn then Controller.B2SSetScoreRollover 25, 2
			P100K.text="200,000"
			If score>299999 Then
				If B2SOn then Controller.B2SSetScoreRollover 25, 3
				P100K.text="300,000"
			end if
		end if
	End if
    if score=>replay1 and rep=0 then
		if replaymode=1 then
			PlaySound SoundFXDOF("knock",138,DOFPulse,DOFKnocker)
			DOF 139, DOFPulse
			addcredit
		  Else
			playsound SoundFXDOF("dabell",133,DOFPulse,DOFBell)
			ShootAgain.state=1
		end if
		rep=1
    end if
    if score=>replay2 and rep=1 then
		if replaymode=1 then
			PlaySound SoundFXDOF("knock",138,DOFPulse,DOFKnocker)
			DOF 139, DOFPulse
			addcredit
		  Else
			playsound SoundFXDOF("dabell",133,DOFPulse,DOFBell)
			ShootAgain.state=1
		end if
		rep=2
    end if
    if score=>replay3 and rep=2 then
		if replaymode=1 then
			PlaySound SoundFXDOF("knock",138,DOFPulse,DOFKnocker)
			DOF 139, DOFPulse
			addcredit
		  Else
			playsound SoundFXDOF("dabell",133,DOFPulse,DOFBell)
			ShootAgain.state=1
		end if
		rep=3
    end if
end sub


Sub CheckTilt
	If Tilttimer.Enabled = True Then
	 TiltSens = TiltSens + 1
	 if TiltSens = 3 Then
		GameTilted
	 End If
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
	for i=1 to 2
		EVAL("Bumper"&i).hashitevent = 0
	Next
	LeftSlingShot.disabled=true
  	LeftFlipper.RotateToStart
	StopSound "Buzz"
	DOF 101, DOFOff
	RightFlipper.RotateToStart
	UpperFlipper.RotateToStart
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

Sub UpperFlipper_Collide(parm)
 	RandomSoundFlipper()
	if thirdflipper=0 then addscore 10
End Sub

Sub RandomSoundFlipper()
	Select Case Int(Rnd*3)+1
		Case 1 : PlaySound "flip_hit_1", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
		Case 2 : PlaySound "flip_hit_2", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
		Case 3 : PlaySound "flip_hit_3", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
	End Select
End Sub


sub savehs
    savevalue "StarWarsBD", "credit", credit
    savevalue "StarWarsBD", "hiscore", hisc
    savevalue "StarWarsBD", "match", matchnumb
    savevalue "StarWarsBD", "score", score
	savevalue "StarWarsBD", "balls", balls
	savevalue "StarWarsBD", "freeplay", freeplay
	savevalue "StarWarsBD", "showhisc", showhisc
	savevalue "StarWarsBD", "thirdflipper", thirdflipper
	savevalue "StarWarsBD", "replaymode", replaymode
	savevalue "StarWarsBD", "replays", replays
	savevalue "StarWarsBD", "hsa1", HSA1
	savevalue "StarWarsBD", "hsa2", HSA2
	savevalue "StarWarsBD", "hsa3", HSA3
end sub

sub loadhs
    dim temp
	temp = LoadValue("StarWarsBD", "credit")
    If (temp <> "") then credit = CDbl(temp)
    temp = LoadValue("StarWarsBD", "hiscore")
    If (temp <> "") then hisc = CDbl(temp)
    temp = LoadValue("StarWarsBD", "match")
    If (temp <> "") then matchnumb = CDbl(temp)
    temp = LoadValue("StarWarsBD", "score")
    If (temp <> "") then score = CDbl(temp)
    temp = LoadValue("StarWarsBD", "balls")
    If (temp <> "") then balls = CDbl(temp)
    temp = LoadValue("StarWarsBD", "freeplay")
    If (temp <> "") then freeplay = CDbl(temp)
    temp = LoadValue("StarWarsBD", "showhisc")
    If (temp <> "") then showhisc = CDbl(temp)
    temp = LoadValue("StarWarsBD", "thirdflipper")
    If (temp <> "") then thirdflipper = CDbl(temp)
    temp = LoadValue("StarWarsBD", "replaymode")
    If (temp <> "") then replaymode = CDbl(temp)
    temp = LoadValue("StarWarsBD", "replays")
    If (temp <> "") then replays = CDbl(temp)
    temp = LoadValue("StarWarsBD", "hsa1")
    If (temp <> "") then HSA1 = CDbl(temp)
    temp = LoadValue("StarWarsBD", "hsa2")
    If (temp <> "") then HSA2 = CDbl(temp)
    temp = LoadValue("StarWarsBD", "hsa3")
    If (temp <> "") then HSA3 = CDbl(temp)
end sub

Sub StarWars_Exit()
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

Function AudioFade(tableobj) ' Fades between front and back of the table (for surround systems or 2x2 speakers, etc), depending on the Y position on the table. "StarWars" is the name of the table
  Dim tmp
  tmp = tableobj.y * 2 / StarWars.height-1
  If tmp > 0 Then
    AudioFade = Csng(tmp ^10)
  Else
    AudioFade = Csng(-((- tmp) ^10) )
  End If
End Function

Function AudioPan(tableobj) ' Calculates the pan for a tableobj based on the X position on the table. "StarWars" is the name of the table
  Dim tmp
  tmp = tableobj.x * 2 / StarWars.width-1
  If tmp > 0 Then
    AudioPan = Csng(tmp ^10)
  Else
    AudioPan = Csng(-((- tmp) ^10) )
  End If
End Function

Function Pan(ball) ' Calculates the pan for a ball based on the X position on the table. "StarWars" is the name of the table
    Dim tmp
    tmp = ball.x * 2 / StarWars.width-1
    If tmp > 0 Then
        Pan = Csng(tmp ^10)
    Else
        Pan = Csng(-((- tmp) ^10) )
    End If
End Function

Function AudioFade(ball) ' Can this be together with the above function ?
  Dim tmp
  tmp = ball.y * 2 / StarWars.height-1
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
  If StarWars.VersionMinor > 3 OR StarWars.VersionMajor > 10 Then
    PlaySound("fx_collide"), 0, Csng(velocity) ^2 / 200, Pan(ball1), 0, Pitch(ball1), 0, 0, AudioFade(ball1)
  Else
    PlaySound("fx_collide"), 0, Csng(velocity) ^2 / 200, Pan(ball1), 0, Pitch(ball1), 0, 0
  End if
End Sub


' Thalamus : Exit in a clean and proper way
Sub StarWars_exit()
  Controller.Pause = False
  Controller.Stop
End Sub


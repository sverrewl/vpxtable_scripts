'---------------------
' Grand Prix (Williams 1977)
' VPX version by Allknowing2012
'
' Original VP9 Script by ROSVE
' Plastics from Popotte
' DOF Certified by Argrim :-)
' Note tilt_trigger - a rollover found on a couple WPC table is not coded - useless in VP world .. read IPDB rollover tilt for info
'---------------------

' Thalamus 2018-07-23
' Added/Updated "Positional Sound Playback Functions" and "Supporting Ball & Sound Functions"
' No special SSF tweaks yet.
' This is a JP table. He often uses walls as switches so I need to be careful of using PlaySoundAt

Option Explicit
Randomize

Const cGameName = "grandprix_1977"
Const PI = 3.14159265359

Dim tiltsens
Dim roll
Dim BIP  ' what ball are we on?
Dim ball ' actual ball object
Dim RBonus
Dim RBonusQueue
Dim LBonus
Dim LBonusQueue
Dim GameSeq
Dim ReelPulses(24,2)  '(PulseQueue, ReelValue)
Dim MaxPlayers
Dim ActivePlayer
Dim Round
Dim BallInPlay
Dim credit
Dim Tilt
Dim DoubleBonusFlag
Dim CollectL
Dim AltArrows
Dim Rspin,LSpin
Dim AB,CD,Stars
Dim AddBall(4)
Dim Add10, Add100, Add1000, hisc, Matchnumb
Dim i, kickstep
Dim rep(4)
Dim score(4),sreels(4)
Dim balls
Dim replays, Replay1Table(3), Replay2Table(3), Replay3Table(3)
dim replay1, replay2, replay3, freeplay
dim objekt
Dim bumperscore, state, OperatorMenu, Options
dim am


On Error Resume Next
ExecuteGlobal GetTextFile("controller.vbs")
If Err Then MsgBox "You need the controller.vbs in order to run this table, available in the vp10 package"
On Error Goto 0

' Load the core.vbs for supporting Subs and functions
On Error Resume Next
  ExecuteGlobal GetTextFile("core.vbs")
  If Err Then MsgBox "Can't open core.vbs"
On Error Goto 0


Sub Table1_Init()
    LoadEM
    balls=3
    Matchtxt.text="00"
    Tilttext.text="TILT"
	Replay1Table(1)=470000
	Replay2Table(1)=680000
	Replay3Table(1)=820000
	Replay1Table(2)=510000
	Replay2Table(2)=720000
	Replay3Table(2)=860000

    set sreels(1)= ScoreReel1
    set sreels(2)= ScoreReel2
    set sreels(3)= ScoreReel3
    set sreels(4)= ScoreReel4
    am =0
    'bglights.interval=800:bglights.enabled=True   'uncomment for bglights; game didn't have them.
    credit=0
	loadhs
    highscore.text=hisc
    if credit=0 then
      DOF 127, DOFOff
      CreditLight.state=LightStateOff
    else
      DOF 127, DOFOn
      CreditLight.state=LightStateOn
    end if
    state=False
	if balls="" then balls=5
	if balls<>3 and balls<>5 then balls=5
	if freeplay="" or freeplay<0 or freeplay>1 then freeplay=0
	if replays="" then replays=1
	if replays<>1 and replays<>2 then replays=1

	OptionBalls.image="OptionsBalls"&Balls
	OptionReplays.image="OptionsReplays"&replays
	OptionFreeplay.image="OptionsFreeplay"&freeplay

	if balls=3 then
		InstCard.image="InstCard3balls"
		bumperscore=1000
        Replays=1
	  else
		InstCard.image="InstCard5balls"
		bumperscore=100
        replays=2
	end if

	Replay1=Replay1Table(Replays)
	Replay2=Replay2Table(Replays)
	Replay3=Replay3Table(Replays)
	RepCard.image = "ReplayCard"&replays

    GameSeq=0  'Game Over
    CreditTxt.text=credit
    'turnoff  'lights were always on in this EM game.
	 turnon
	If B2SOn then
        Controller.B2SSetGameOver 35,1
		setBackglass.enabled=true
		for each objekt in backdropstuff : objekt.visible = 0 : next
    Else
		setBackglass.enabled=false
		for each objekt in backdropstuff : objekt.visible = 1 : next
	End If

	for i = 1 to 4
		sreels(i).setvalue(score(i))
	next
    BIPTxt.text=""
	gameovertxt.text="Game Over"
    Bumper1.Force=0
    Bumper2.Force=0
    AltArrows=1
    rspin=0
    lspin=0
 End Sub

 Sub InitGame()
	setBackglass.enabled=False  ' Turn off the bg refresh when the game is finally started - allows for b2s to restart
    for i = 1 to 4
      score(i)=0
      sreels(i).resettozero
      if B2SOn Then Controller.B2SSetScorePlayer i, score(i)   ' 0.96
      rep(i)=0
    Next
    BIP=1
    gameovertxt.text=""
    Matchtxt.text=""
    gameovertxt.text=""
End Sub

Sub InitBall
    tiltsens=0
    If B2SOn then Controller.B2SSetShootAgain 0  ' Need The ExtraBall Call here
    RBonus=0
    RBonusQueue=1
    LBonus=0
    LBonusQueue=1

    If B2SOn then Controller.B2ssetballinplay 32, Balls-Round+1
    If B2SOn then Controller.B2SSetGameOver 35,0
    if B2SOn then Controller.B2ssetplayerup 30, ActivePlayer
    turnon()
    Bumper1.Force=10
    Bumper2.Force=10
   ' Sling1.Disabled=0
   ' Sling2.Disabled=0
    Kicker2.Enabled=1
    Kicker3.Enabled=1
    Kicker4.Enabled=1
    DoubleBonusFlag=0
    AltArrows=AltArrows+1
    If AltArrows=2 Then
       AltArrows=0
       Light_LBA.State=1
       Light_RBA.State=0
       Light_LBA1.State=1
       Light_RBA1.State=0
       Light_LBadv.State=0
       Light_RBadv.State=1
    Else
       Light_LBA.State=0
       Light_RBA.State=1
       Light_LBA1.State=0
       Light_RBA1.State=1
       Light_LBadv.State=1
       Light_RBadv.State=0
    End If
    AB=0
    CD=0
    Stars=0
    TARGET2.IsDropped=False
    TARGET1.IsDropped=False
    TARGET3.IsDropped=False
    TARGET4.IsDropped=False
    BIPTxt.text=BIP
 End Sub


 Sub Table1_KeyDown(ByVal keycode)
      If keycode = AddCreditKey Then
         If credit < 15 Then
		   credit=credit+1
           DOF 127, DOFOn
           CreditLight.state=LightStateOn
         End If
        PlaySound "fx_coin"
debug.print "2:" & Credit
	    If B2SOn Then Controller.B2ssetCredits Credit
        CreditTxt.text=credit
	  End If
      If keycode = AddCreditKey2 Then
         If credit < 15 Then
		   credit=credit+1
           CreditLight.state=LightStateOn
           DOF 127, DOFOn
         End If
        PlaySound "fx_coin"

	    If B2SOn Then Controller.B2ssetCredits Credit
        CreditTxt.text=credit
	End If
     If keycode = StartGameKey Then
         state=true
         BallInPlay=1
         If MaxPlayers < 4 And credit>0 Then
		   MaxPlayers=MaxPlayers+1
           AddPlayer
        End If
	End If

	If keycode = PlungerKey Then
		Plunger.PullBack
         PlaySound "fx_plungerpull" ' ,0,1,0.25,0.25
	End If

	If keycode=LeftFlipperKey and State = false and OperatorMenu=0 then
debug.print "left + state:false + OperatorMenu"
		OperatorMenuTimer.interval=100:OperatorMenuTimer.Enabled = true
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
				Option4.visible=true
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
                replays=2
			  else
				Balls=3
				InstCard.image="InstCard3balls"
                replays=1
			end if
			Replay1=Replay1Table(Replays)
			Replay2=Replay2Table(Replays)
            Replay3=Replay3Table(Replays)
			OptionReplays.image = "OptionsReplays"&replays
			repcard.image = "ReplayCard"&replays
			OptionBalls.image = "OptionsBalls"&Balls
		Case 2:
			if freeplay=0 Then
				freeplay=1
			  Else
				freeplay=0
			end if
			OptionFreeplay.image="OptionsFreeplay"&freeplay
		Case 4:
			OperatorMenu=0
			savehs
			HideOptions
	  End Select
	End If

	If keycode = LeftFlipperKey and Round > 0 and not tilt Then
		LeftFlipper.RotateToEnd
        PlaySound SoundFXDOF("fx_flipperup",101,DOFOn,DOFContactors), 0, .67, -0.05, 0.05
	End If

	If keycode = RightFlipperKey and Round > 0 and not tilt Then
		RightFlipper.RotateToEnd
		PlaySound SoundFXDOF("fx_flipperup",102,DOFOn,DOFContactors), 0, .67, 0.05, 0.05
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
		Tilt=true
		mechchecktilt
	End If

End Sub


Sub OperatorMenuTimer_Timer
debug.print "OperatorMenuTimer_Timer()"
	OperatorMenu=1
	Displayoptions
	Options=1
End Sub

Sub DisplayOptions
debug.print "DisplayOptions"
	OptionsBack.visible = true
	Option1.visible = True
	OptionBalls.visible = True
    OptionReplays.visible = True
	OptionFreeplay.visible = True
    OperatorMenuTimer.Enabled = FALSE ' daryl
End Sub

Sub HideOptions
debug.print "HideOptions"
	for each objekt In OptionMenu
		objekt.visible = false
	next
End Sub

Sub Table1_KeyUp(ByVal keycode)

	If keycode = PlungerKey Then
		Plunger.Fire
         PlaySound "fx_plunger" ' ,0,1,0.25,0.25
	End If

	if keycode = LeftFlipperKey then
		OperatorMenuTimer.Enabled = false
	end if

  if not Tilt and Round > 0 Then
	If keycode = LeftFlipperKey Then
	   LeftFlipper.RotateToStart
       LeftFlipper.TimerEnabled = 1    ' nFozzy Flipper Code
       LeftFlipper.TimerInterval = 16
       LeftFlipper.return = returnspeed * 0.5
	   PlaySound SoundFXDOF("fx_flipperdown",101,DOFOff,DOFContactors), 0, 1, -0.05, 0.05
	End If

	If keycode = RightFlipperKey Then
 	   RightFlipper.RotateToStart
       Rightflipper.TimerEnabled = 1  ' nFozzy Flipper Code
       Rightflipper.TimerInterval = 16
       Rightflipper.return = returnspeed * 0.5
	   PlaySound SoundFXDOF("fx_flipperdown",102,DOFOff,DOFContactors), 0, 1, 0.05, 0.05
	End If
  End If
End Sub

Sub RealTimer_Timer
 Pgate.rotz =  Gate4.currentangle+25
 Pgate1.rotz = Gate3.currentangle+25
End Sub

Sub Drain_Hit()
     DOF 130, DOFPulse
     PlaySound "fx_drain",0,1,0,0.25

     BallInPlay=0
	 Drain.DestroyBall
     me.timerenabled=1

End Sub

Sub Drain_timer
     me.timerenabled=0
     GameSeq=3 'Collect Bonus then start new ball
     SeqTimer.Enabled=1
     SeqTimer.Interval=75 'was 125
	' PlaySound "MotorLeer"
End Sub


 '----------------------------------------------------------
 '     TILT
 '-----------------------------------------------------------


 Sub TiltOn()
    If B2SOn then Controller.B2SSetTilt 33,1
    tilttext.text="TILT"
    PlaySound "motor"
    turnon
    turnoff
    Bumper1.Force=0
    Bumper2.Force=0
    'Sling1.Disabled=1
    'Sling2.Disabled=1
    Kicker2.Enabled=0
    Kicker3.Enabled=0
    Kicker4.Enabled=0
    RightFlipper.RotateToStart
    LeftFlipper.RotateToStart
	DOF 101, DOFOff
	DOF 102, DOFOff
    Kicker2.Kick (Rnd*10)+150,9
    Kicker3.Kick (Rnd*10)+200,9
    Kicker4.Kick (Rnd*40)+160,7
	DOF 111, DOFPulse
	DOF 112, DOFPulse
	DOF 113, DOFPulse
 End Sub

Sub CheckTilt
	If Tilttimer.Enabled = True Then
	 TiltSens = TiltSens + 1
	 if TiltSens = 3 Then
	   Tilt = True
       'If B2SOn Then Controller.B2ssetdata 1, 0
	   TiltOn()
	 End If
	Else
	 TiltSens = 0
	 Tilttimer.Enabled = True
	End If
End Sub

Sub MechCheckTilt
	   Tilt = True
       'If B2SOn Then Controller.B2ssetdata 1, 0
	   TiltOn()
End Sub

Sub Tilttimer_Timer()
	Tilttimer.Enabled = False
End Sub

 '----------------------------------------
  'Add Players and start game
 '----------------------------------------
 Sub AddPlayer()
   Dim sr
  If Round=0 Or Round=5 Or Round=3 Then
    If credit>0 Then
     credit=credit-1
debug.print "3:" & Credit
	 If B2SOn Then Controller.B2ssetCredits Credit
     CreditTxt.text=credit
    if credit=0 then
      DOF 127, DOFOff
      CreditLight.state=LightStateOff
    else
      DOF 127, DOFOn
      CreditLight.state=LightStateOn
    end if
    Select Case MaxPlayers
     Case 1
         if B2SOn then Controller.B2ssetcanplay 31, 1
         cp1.state=LightStateOn
         if B2SOn then Controller.B2ssetMatch 34, 0
         matchtxt.text=""
		 PlaySound "initialize"
         Round=Balls
         ActivePlayer=0
         InitGame()
         GameSeq=2
         SeqTimer.Interval=2500:SeqTimer.Enabled=1 'increase from 500
     Case 2
         if B2SOn then Controller.B2ssetcanplay 31, 2
         cp1.state=LightStateOff
         cp2.state=LightStateOn
         PlaySound "addplayer"
     Case 3
         if B2SOn then Controller.B2ssetcanplay 31, 3
         cp2.state=LightStateOff
         cp3.state=LightStateOn
         PlaySound "addplayer"
     Case 4
         if B2SOn then Controller.B2ssetcanplay 31, 4
         cp3.state=LightStateOff
         cp4.state=LightStateOn
         PlaySound "addplayer"
     End Select
     End If
   End If

debug.print "4:" & Credit
   If B2SOn Then  Controller.B2ssetCredits Credit
   CreditTxt.text=credit
   If B2SOn Then
	Controller.B2ssetplayerup 30, 1
	Controller.B2ssetcanplay 31, MaxPlayers
	Controller.B2SSetScorePlayer 5, hisc
    highscore.text=hisc
   End If
 End Sub

Sub addcredit
      credit=credit+1
	  DOF 127, DOFOn
      CreditLight.state=LightStateOn
      if credit>15 then credit=15
debug.print "6:" & Credit
	  If B2SOn Then Controller.B2ssetCredits Credit
      CreditTxt.text=credit
 End sub

 '------------------------------------------------------------------
 '    Game Sequence
 '------------------------------------------------------------------
 Sub SeqTimer_Timer()
    Select Case GameSeq
       Case 0 'Game Over
          ' Attract Mode

       Case 1 'New Ball
              SeqTimer.Enabled=0
              InitBall()
              Set ball = BallRelease.CreateBall
              BallRelease.Kick 45,9
              PlaySound "motor"
			PlaySound "bankreset"
			  PlaySound SoundFXDOF("popper_ball",120,DOFPulse,DOFContactors)
			  gameovertxt.text=""
       Case 2 'Drain
              'Prepare for next ball
              If B2SOn then Controller.B2SSetTilt 33,0
              tilttext.text=""
			  gameovertxt.text=""
              Tilt=false
              IF Light_ShootAgain.State=0 Then
                 If ActivePlayer=MaxPlayers Then
                    BIP=BIP+1
                    Round=Round-1
                    If Round<1 Then
                       ActivePlayer=0
                       plno1.state=LightStateOff
                       plno2.state=LightStateOff
                       plno3.state=LightStateOff
                       plno4.state=LightStateOff
                    Else
                       ActivePlayer=1
                       plno1.state=LightStateOn
                    End If
                 Else
                    plno1.state=LightStateOff
                    plno2.state=LightStateOff
                    plno3.state=LightStateOff
                    plno4.state=LightStateOff
                    ActivePlayer=ActivePlayer+1
                    if ActivePlayer=1 Then
                       plno1.state=LightStateOn
                    else if ActivePlayer=2 then
                       plno2.state=LightStateOn
                    else if ActivePlayer=3 then
                       plno3.state=LightStateOn
                    else if ActivePlayer=4 then
                       plno4.state=LightStateOn
                    end If
                    end If
                    end if
                    end if
                 End If

              End If
              IF Round <1 Then
                  If B2SOn then Controller.B2ssetballinplay 32, 0
                  BIPTxt.text=""
                  'End of game
                  state=false
                  matchnumb=(INT(RND*10))*10
                  ' Check Match Wheel ------
                  For i=1 to MaxPlayers
		             if (matchnumb)=(score(i) mod 100) then
		               addcredit
		               playsound SoundFXDOF("knock",128,DOFPulse,DOFKnocker)
					   DOF 126, DOFPulse
	                 end if
                     if score(i) > hisc then hisc = score(i)
                  next
                  highscore.text=hisc
                  if matchnumb = 0 then
                     matchnumb=100
                     matchtxt.text="00"
                  Else
                     matchtxt.text=Matchnumb
                  End if
                  if B2SOn then
                      Controller.B2ssetMatch 34, Matchnumb
                  Else
                    matchtxt.visible=True
                  end if
                  debug.print "Print Match of " & MatchNumb
                   '--------------------------
                  MaxPlayers=0
				  PlaySound "MotorLeer"
                  SeqTimer.Enabled=0
                  If B2SOn then Controller.B2SSetGameOver 35,1
                  Bumper1.Force=0
                  Bumper2.Force=0
                  BumperLight1.State=LightStateOff
                  BumperLight2.State=LightStateOff
                  if B2SOn then Controller.B2ssetplayerup 30, 0
                  if B2SOn Then Controller.B2ssetcanplay 31, 0
                  If B2SOn then Controller.B2ssetballinplay 32, 0
                  cp1.state=LightStateOff
                  cp2.state=LightStateOff
                  cp3.state=LightStateOff
                  cp4.state=LightStateOff
                  gameovertxt.text="Game Over" 'Game Over
              Else
                 'Start new ball
                 BIPTxt.text=BIP
				 gameovertxt.text=""
                 GameSeq=1
                 SeqTimer.Enabled=1
                 SeqTimer.Interval=500 '500
                 'If ActivePlayer=1 and Round=3 Then
             '       PlaySound "Start"
                  '  SeqTimer.Interval=100 '1800
             '    Else
              '      PlaySound "Start2"
               '  End If
              End If
       Case 3 'Collect Bonus
               If not Tilt Then
                  GameSeq=4
                  If Light_LBA.State=1 Then
                     DoubleBonusFlag=1
                     CollectL=1
                     RBonus=1
                     If Light_Star1.State=LightStateOn Then DoubleBonusFlag=0
                     BonusCollectTimer.Enabled=1
                  Else
                     DoubleBonusFlag=1
                     CollectL=0
                     LBonus=1
                     If Light_Star1.State=LightStateOn Then DoubleBonusFlag=0
                     BonusCollectTimer.Enabled=1
                  End If
               Else
                  SeqTimer.Interval=250
                  ScoreTimer.Interval=250
                  GameSeq=2
               End If
       Case 4 'Wait for bonus
               If LBonus<1 Or RBonus<1 Then
                  SeqTimer.Interval=900
                  ScoreTimer.Interval=150
                  GameSeq=2
               End If
       End Select
 End Sub

 '--------------------------------------------------------------------
 '------- TARGETS
 '--------------------------------------------------------------------
Sub TARGET1_dropped()
   DOF 107, DOFPulse
   'A
   If Tilt Then Exit Sub
   TARGET1.IsDropped=True
   AddScore (1000)
   AB=AB+1
   CheckAB
End Sub

Sub TARGET2_dropped()
   DOF 108, DOFPulse
   'B
   If Tilt Then Exit Sub
   TARGET2.IsDropped=True
   AddScore (1000)
   AB=AB+1
   CheckAB
End Sub

Sub TARGET3_dropped()
   DOF 109, DOFPulse
   'C
   If Tilt Then Exit Sub
   TARGET3.IsDropped=True
   AddScore (1000)
   CD=CD+1
   CheckCD
End Sub

Sub TARGET4_dropped()
   DOF 110, DOFPulse
   'D
   If Tilt Then Exit Sub
   TARGET4.IsDropped=True
   AddScore (1000)
   CD=CD+1
   CheckCD
End Sub

Sub CheckAB()
   If AB=2 Then
     TARGET2.TimerEnabled=1
     CheckStars
   End If
End Sub

Sub TARGET2_Timer()
   TARGET2.TimerEnabled=0
   AB=0:CD=0
   TARGET4.IsDropped=False
   TARGET3.IsDropped=False
   TARGET2.IsDropped=False
   TARGET1.IsDropped=False
   PlaySound SoundFXDOF("BankReset",131,DOFPulse,DOFContactors)
End Sub

Sub CheckCD()
   If CD=2 Then
     TARGET4.TimerEnabled=1
     CheckStars
   End If
End Sub

Sub TARGET4_Timer()
   TARGET4.TimerEnabled=0
   AB=0:CD=0
   TARGET4.IsDropped=False
   TARGET3.IsDropped=False
   TARGET2.IsDropped=False
   TARGET1.IsDropped=False
   PlaySound SoundFXDOF("BankReset",131,DOFPulse,DOFContactors)
End Sub

Sub CheckStars()
  If Stars<4 Then
   Select Case Stars
      Case 0
         Stars=1
         Light_Star1.State=1
         PlaySound "fx_solenoid"
      Case 1
         Stars=2
         Light_Star2.State=1
         PlaySound "fx_solenoid"
         Light_LXB.State=1
      Case 2
         Stars=3
         Light_Star3.State=1
         PlaySound "fx_solenoid"
      Case 3
         Stars=4
         Light_Star4.State=1
         PlaySound "fx_solenoid"
         Light_LSpecial.State=1
   End Select
  End If
End Sub

 '----------------------------------------------------------------------
 '     BUMPER
 '----------------------------------------------------------------------
 Sub Bumper1_Hit()
    PlaySound SoundFXDOF("fx_bumper",105,DOFPulse,DOFContactors)
    if tilt then exit sub
    bumperLight1.state=LightStateOff
    bumper1.timerinterval=200:bumper1.timerenabled=true
    AddScore (bumperscore)
 End Sub

 Sub Bumper2_Hit()
    PlaySound SoundFXDOF("fx_bumper",106,DOFPulse,DOFContactors)
    if tilt then exit sub
    bumperLight2.state=LightStateOff
    bumper2.timerinterval=200:bumper2.timerenabled=true
    AddScore (bumperscore)
 End Sub

Sub Bumper1_Timer()
  Bumper1.timerenabled=False
  bumperLight1.state=LightStateOn
End Sub

Sub Bumper2_Timer()
  Bumper2.timerenabled=False
  bumperLight2.state=LightStateOn
End Sub

 '------------------------------------------------------------------------
 '           SPINNERS
 '------------------------------------------------------------------------


 Sub Spinner1_Spin()
    PlaySound "fx_spinner"
	DOF 114, DOFPulse
    If Tilt Then Exit Sub
    If Light_LSA.State=0 Then
       AddScore (100)
    Else
       AddScore (1000)
    End If
    LSLights(Lspin).State=0
    Lspin=Lspin+1
    If Lspin=10 Then
       Lspin=0
	   If balls=3 then
         LBonusQueue=3
       Else
         LBonusQueue=2
       End If
    End If
    LSLights(Lspin).State=1
 End Sub

 Sub Spinner2_Spin()
    PlaySound "fx_spinner"
	DOF 115, DOFPulse
    If Tilt Then Exit Sub
    If Light_RSA.State=0 Then
       AddScore (100)
    Else
       AddScore (1000)
    End If
    RSLights(Rspin).State=0
    Rspin = Rspin +1
    If Rspin=10 Then
       Rspin=0
	   If balls=3 then
         RBonusQueue=3
       Else
         RBonusQueue=2
       End If
    End If
    RSLights(Rspin).State=1
 End Sub

 '------------------------------------------------------------------------
 '           TRIGGERS
 '------------------------------------------------------------------------
 Sub Trigger1_Hit()
    PlaySound "Target"
	DOF 116, DOFPulse
    If Tilt Then Exit Sub
    If Light_LSpecial.State=1 Then
       'Special
        credit=credit+1
        PlaySound SoundFXDOF("knocker",128,DOFPulse,DOFKnocker)
		DOF 126, DOFPulse
        DOF 127, DOFOn
        CreditLight.state=LightStateOn
    Else
       AddScore (10000)
       LBonusQueue=LBonusQueue+1    ' 0.96
    End if
  End Sub

 Sub Trigger2_Hit()
    PlaySound "Target"
	DOF 117, DOFPulse
    If Tilt Then Exit Sub
    AddScore (10000)
    If Light_LXB.State=1 Then
       Light_ShootAgain.State=1
       Light_LXB.State=0
       If B2SOn then Controller.B2SSetShootAgain 1   ' Need The ExtraBall Call here
       PlaySound "solon"
    End If
 End Sub

 Sub Trigger3_Hit()
    PlaySound "Target"
	DOF 118, DOFPulse
    If Tilt Then Exit Sub
    AddScore (10000)
    If Light_RXB.State=1 Then
       Light_ShootAgain.State=1
       Light_RXB.State=0
       If B2SOn then Controller.B2SSetShootAgain 1   ' Need The ExtraBall Call here
       PlaySound "fx_solenoid"
    End If
 End Sub

 Sub Trigger4_Hit()
    PlaySound "Target"
	DOF 119, DOFPulse
    If Tilt Then Exit Sub
    If Light_RSpecial.State=1 Then
       'Special
        credit=credit+1
        PlaySound SoundFXDOF("knocker",128,DOFPulse,DOFKnocker)
		DOF 126, DOFPulse
        DOF 127, DOFOn
        CreditLight.state=LightStateOn
    Else
       AddScore (10000)
       RBonusQueue=RBonusQueue+1  ' 0.96
    End if
 End Sub

Sub Trigger5_Hit
	DOF 121, DOFPulse
End Sub

 Sub sling3_Slingshot()
    'Left middle rubber
    AddScore (500)
    LBonusQueue=LBonusQueue+1
 End Sub

 Sub sling4_Slingshot()
    'Right middle rubber
    AddScore (500)
    RBonusQueue=RBonusQueue+1
 End Sub

'**********Sling Shot Animations
' Rstep and Lstep  are the variables that increment the animation
'****************
Dim RStep, Lstep

Sub RightSlingShot_Slingshot
	addscore (10)
	playsound SoundFXDOF("right_slingshot",104,DOFPulse,DOFContactors)
    RSling.Visible = 0
    RSling1.Visible = 1
    sling1.TransZ = -10
    RStep = 0
    RightSlingShot.TimerEnabled = 1
    'gi1.state=0
End Sub

Sub RightSlingShot_Timer
    Select Case RStep
        Case 3:RSLing1.Visible = 0:RSLing2.Visible = 1:sling1.TransZ = -20
        Case 4:RSLing2.Visible = 0:RSLing.Visible = 1:sling1.TransZ = 0:RightSlingShot.TimerEnabled = 0 ' :gi1.state=1:gi13.state=1
    End Select
    RStep = RStep + 1
End Sub

Sub LeftSlingShot_Slingshot
	addscore (10)
	playsound SoundFXDOF("left_slingshot",103,DOFPulse,DOFContactors)
    LSling.Visible = 0
    LSling1.Visible = 1
    sling2.TransZ = -10
    LStep = 0
    LeftSlingShot.TimerEnabled = 1
    'gi1.state=0:gi13.state=0
End Sub

Sub LeftSlingShot_Timer
    Select Case LStep
        Case 3:LSLing1.Visible = 0:LSLing2.Visible = 1:sling2.TransZ = -20
        Case 4:LSLing2.Visible = 0:LSLing.Visible = 1:sling2.TransZ = 0 ':LeftSlingShot.TimerEnabled = 0 :gi1.state=1:gi13.state=1
    End Select
    LStep = LStep + 1
End Sub

Sub LeftSlingshot_Init
	LSling1.Visible = 0:LSling2.Visible = 0: Sling2.TransZ = 0
End Sub

Sub RightSlingshot_Init
	RSling1.Visible = 0:RSling2.Visible = 0: Sling1.TransZ = 0
End Sub



 Sub R33_Hit()
    'Left Top rubber
    If Tilt Then Exit Sub
    AddScore (10)
 End Sub

 Sub R34_Hit()
    'Right Top rubber
    If Tilt Then Exit Sub
    AddScore (10)
 End Sub

Sub AltLights()
   If Light_LBadv.State=1 Then
      Light_LBadv.State=0
      Light_RBadv.State=1
   Else
      Light_LBadv.State=1
      Light_RBadv.State=0
   End If
   If Light_Star2.State=1 Then
      'Alt Xball lanes
      if Light_ShootAgain.State=0 then
        If Light_LXB.State=1 Then
           Light_LXB.State=0
           Light_RXB.State=1
        Else
           Light_LXB.State=1
           Light_RXB.State=0
        End If
     End if
   End If
   If Light_Star4.State=1 Then
      'Alt Xball lanes
      If Light_LSpecial.State=1 Then
         Light_LSpecial.State=0
         Light_RSpecial.State=1
      Else
         Light_LSpecial.State=1
         Light_RSpecial.State=0
      End If
   End If
End Sub

 '------------------------------------------------------------------------
 '             KICKERS
 '------------------------------------------------------------------------

 Sub Kicker2_Hit()
	debug.print "Score = " & score(ActivePlayer)
    DoubleBonusFlag=1
    CollectL=1
    If Light_Star1.State=LightStateOn Then DoubleBonusFlag=0
    Kicker2.TimerInterval=900*LBonus*(2-DoubleBonusFlag)+500
	PlaySound "motor"
    BonusCollectTimer.Enabled=1
	kickstep=1
    Kicker2.TimerEnabled = 1
 End Sub

 Sub Kicker2_Timer()
    kicker2.timerEnabled=0
	kickstep=1
    kicker2Timer.Enabled=1
 End Sub

 Sub Kicker2Timer_Timer()
	Select Case kickstep
	  Case 4:
		playsound SoundFXDOF("popper_ball",111,DOFPulse,DOFContactors)
		DOF 126, DOFPulse
		KickArm2.rotz=10
        Kicker2.Kick (Rnd*10)+140,6
	  Case 5:
		kickarm2.rotz=0
		kicker2Timer.enabled=0
	End Select
	kickstep=kickstep+1
 End Sub

 Sub Kicker3_Hit()
	PlaySound "motor"
    debug.print "Score = " & score(ActivePlayer)
    DoubleBonusFlag=1
    CollectL=0
    If Light_Star1.State=LightStateOn Then DoubleBonusFlag=0
    Kicker3.TimerInterval=900*RBonus*(2-DoubleBonusFlag)+500
	PlaySound "motor"
    BonusCollectTimer.Enabled=1
    Kicker3.TimerEnabled = 1
 End Sub

 Sub Kicker3_Timer()
    kicker3.timerEnabled=0
	kickstep=1
    kicker3Timer.Enabled=1
End Sub

Sub kicker3Timer_Timer()
	Select Case kickstep
	  Case 4:
		playsound SoundFXDOF("popper_ball",112,DOFPulse,DOFContactors)
		DOF 126, DOFPulse
		kickArm3.rotz=10
        Kicker3.Kick (Rnd*10)+205,6
	  Case 5:
		kickarm3.rotz=0
		kicker3Timer.enabled=0
	End Select
	kickstep=kickstep+1
 End Sub

 Sub Kicker4_Hit()
	kickstep=1
    Kicker4.TimerEnabled = 1
    AddScore (5000)
	PlaySound "motor"
    If Light_LBadv.State=1 Then
       LBonusQueue=2
    Else
       RBonusQueue=2
    End If
 End Sub

Sub kicker4_Timer()
	Select Case kickstep
	  Case 4:
		playsound SoundFXDOF("popper_ball",113,DOFPulse,DOFContactors)
		DOF 126, DOFPulse
		kickArm4.rotz=10
        Kicker4.Kick (Rnd*40)+160,7
	  Case 5:
		kickarm4.rotz=0
		me.timerenabled=0
	End Select
	kickstep=kickstep+1
 End Sub


 '------------------------------------------------------------------------
 '             SOUNDS
 '------------------------------------------------------------------------

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
	PlaySound "fx_gate", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
End Sub

Sub a_Spinner_Spin(idx)
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

 '--------------------------------------------------------------------------
 '        Scoring
 '--------------------------------------------------------------------------
sub addscore(points)
  if not tilt then
    debug.print "AddScore " & points
    If Points < 100 then ' and AddScore10Timer.enabled = false Then
        Add10 = Add10 + (Points \ 10)
        AddScore10Timer.Enabled = TRUE
      ElseIf Points < 1000 then ' and AddScore100Timer.enabled = false Then
        Add100 = Add100 + (Points \ 100)
        AddScore100Timer.Enabled = TRUE
      Else 'If AddScore1000Timer.enabled = false Then
        Add1000 = Add1000 + (Points \ 1000)
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

Sub AddPoints(Points)
debug.print "BEFORE Player #" & ActivePlayer & " Score : " & score(ActivePlayer)
    score(ActivePlayer)=score(ActivePlayer)+points
    sreels(ActivePlayer).addvalue(Points)

debug.print "Player #" & ActivePlayer & " Score : " & score(ActivePlayer)
	If B2SOn Then Controller.B2SSetScorePlayer ActivePlayer, score(ActivePlayer)

    ' Sounds: there are 3 sounds: tens, hundreds and thousands
    If Points = 100 AND(Score(ActivePlayer) MOD 1000) \ 100 = 0 Then  'New 1000 reel
        PlaySound SoundFXDOF("bell1000",143,DOFPulse,DOFChimes)
      ElseIf Points = 10 AND(Score(ActivePlayer) MOD 100) \ 10 = 0 Then 'New 100 reel
        PlaySound SoundFXDOF("bell100",142,DOFPulse,DOFChimes)
      ElseIf points = 1000 Then
        PlaySound SoundFXDOF("bell1000",143,DOFPulse,DOFChimes)
	  elseif Points = 100 Then
        PlaySound SoundFXDOF("bell100",142,DOFPulse,DOFChimes)
      Else
        PlaySound SoundFXDOF("bell10",141,DOFPulse,DOFChimes)
        AltLights()
    End If
	checkreplays
end Sub

sub checkreplays
    ' check replays and rollover
    if score(ActivePlayer)=>replay1 and rep(ActivePlayer)=0 then
		addcredit
		rep(ActivePlayer)=1
		PlaySound SoundFXDOF("knock",128,DOFPulse,DOFKnocker)
		DOF 126, DOFPulse
        CreditLight.state=LightStateOn
    end if
    if score(ActivePlayer)=>replay2 and rep(ActivePlayer)=1 then
		addcredit
		rep(ActivePlayer)=2
		PlaySound SoundFXDOF("knock",128,DOFPulse,DOFKnocker)
		DOF 126, DOFPulse
        CreditLight.state=LightStateOn
    end if
    if score(ActivePlayer)=>replay3 and rep(ActivePlayer)=2 then
		addcredit
		rep(ActivePlayer)=3
		PlaySound SoundFXDOF("knock",128,DOFPulse,DOFKnocker)
		DOF 126, DOFPulse
        CreditLight.state=LightStateOn
    end if
end sub


'---------------------------------------------------------------------
'       Handle Bonus
'---------------------------------------------------------------------

 Sub AddRightBonus( ByVal bns)
    RBonusQueue=RBonusQueue+bns
 End Sub

 Sub AddLeftBonus( ByVal bns)
    LBonusQueue=LBonusQueue+bns
 End Sub

Sub BonusTimer_Timer()
      'Add Left Bonus
      If LBonusQueue>0 Then
         LBonusQueue=LBonusQueue-1
         If LBonus<10 Then
            LBLights(LBonus).State=1
            LBonus=LBonus+1
            If LBonus=10 Then Light_LSA.State=1
            PlaySound "fx_solenoid"
         Else
            LBonusQueue=0
         End If
      End If
      'Add Right Bonus
      If RBonusQueue>0 Then
         RBonusQueue=RBonusQueue-1
         If RBonus<10 Then
            RBLights(RBonus).State=1
            RBonus=RBonus+1
            If RBonus=10 Then Light_RSA.State=1
            PlaySound "fx_solenoid"
         Else
            RBonusQueue=0
         End If
      End If
End Sub

 Sub BonusCollectTimer_Timer()
    If Tilt Then
       LBonus=0
       RBonus=0
    End If
'msgbox "Left = " & LBonus & " Right=" & RBonus
    If CollectL=1 Then
       'Left Bonus
       If LBonus>0 Then
          AddScore (5000)
          DoubleBonusFlag=DoubleBonusFlag+1
          If DoubleBonusFlag>1 Then
             LBonus=LBonus-1
             LBLights(LBonus).State=0
             DoubleBonusFlag=1
             If Light_Star1.State=LightStateOn Then DoubleBonusFlag=0
          End If
       Else
          BonusCollectTimer.Enabled=0
          Light_RSA.State=0
          debug.print "End Score = " & score(ActivePlayer)
       End If
    Else
       'Right Bonus
       If RBonus>0 Then
          AddScore (5000)
          DoubleBonusFlag=DoubleBonusFlag+1
          If DoubleBonusFlag>1 Then
             RBonus=RBonus-1
             RBLights(RBonus).State=0
             DoubleBonusFlag=1
             If Light_Star1.State=LightStateOn Then DoubleBonusFlag=0
          End If
       Else
          BonusCollectTimer.Enabled=0
          Light_LSA.State=0
          debug.print "End Score = " & score(ActivePlayer)
       End If
    End If
 End Sub

' SpinnerRod code from Cyperpez and http://www.vpforums.org/index.php?showtopic=35497
Sub CheckSpinnerRod_timer()
	SpinnerRod1.TransZ = sin( (spinner1.CurrentAngle+180) * (2*PI/360)) * 5
	SpinnerRod1.TransX = -1*(sin( (spinner1.CurrentAngle- 90) * (2*PI/360)) * 5)
	SpinnerRod2.TransZ = sin( (spinner2.CurrentAngle+180) * (2*PI/360)) * 5
	SpinnerRod2.TransX = -1*(sin( (spinner2.CurrentAngle- 90) * (2*PI/360)) * 5)
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

sub setBackglass_timer
    CreditTxt.text=credit
    gameovertxt.text="Game Over"
    highscore.text=hisc
	If B2SOn Then
	  Controller.B2SSetMatch 34, 0 'Matchnumb
      Controller.B2SSetGameOver 35,1
debug.print "1:" & Credit
      Controller.B2ssetCredits Credit
      Controller.B2SSetScorePlayer 5, hisc
      Controller.B2SSetScorePlayer 1, score(1)
      Controller.B2SSetScorePlayer 2, score(2)
      Controller.B2SSetScorePlayer 3, score(3)
      Controller.B2SSetScorePlayer 4, score(4)
	gameovertxt.text="Game Over" 'fix that stops "Game Over" from being displayed on B2S backglass all the time.
    End if

end Sub

sub loadhs
    dim temp
	temp = LoadValue(cGameName, "credit")
    If (temp <> "") then credit = CDbl(temp)
    temp = LoadValue(cGameName, "hiscore")
    If (temp <> "") then
      hisc = CDbl(temp)
    Else
      hisc = 100000
    End If
    temp = LoadValue(cGameName, "match")
    If (temp <> "") then
      matchnumb = CDbl(temp)
    Else
      matchnumb = Cdbl("10")
    End If
    temp = LoadValue(cGameName, "score1")
    If (temp <> "") then score(1) = CDbl(temp)
    temp = LoadValue(cGameName, "score2")
    If (temp <> "") then score(2) = CDbl(temp)
    temp = LoadValue(cGameName, "score3")
    If (temp <> "") then score(3) = CDbl(temp)
    temp = LoadValue(cGameName, "score4")
    If (temp <> "") then score(4) = CDbl(temp)
    temp = LoadValue(cGameName, "balls")
    If (temp <> "") then balls = CDbl(temp)
end sub

sub savehs
    savevalue cGameName, "credit", credit
    savevalue cGameName, "hiscore", hisc
    savevalue cGameName, "match", matchnumb
    savevalue cGameName, "score1", score(1)
    savevalue cGameName, "score2", score(2)
    savevalue cGameName, "score3", score(3)
    savevalue cGameName, "score4", score(4)
	savevalue cGameName, "balls", balls
end sub

Sub turnoff
    For Each I In CollectionLightsOff
      I.State=LightStateOff
    Next
    For Each I In CollectionLightsOn
      I.State=LightStateOff
    Next
    BumperLight1.State=LightStateOff
    BumperLight2.State=LightStateOff
End Sub

Sub turnon
    For Each I In CollectionLightsOff
      I.State=LightStateOff
    Next
    For Each I In CollectionLightsOn
      I.State=LightStateOn
    Next
    BumperLight1.State=LightStateOn
    BumperLight2.State=LightStateOn
End Sub

Sub Table1_Exit()
	turnoff
	Savehs
	If B2SOn Then Controller.stop
End Sub



' Just some random bg light stuff

Sub bglights_timer()
  if Not B2SOn then Exit Sub
  am=am+1
  if am > 5 then am=1
  Select case am
  case 1:
    Controller.B2SSetData 19,0
    Controller.B2SSetData 20,0
    Controller.B2SSetData 21,0
    Controller.B2SSetData 22,0
    Controller.B2SSetData 23,0
    Controller.B2SSetData 24,1
  case 2:
    Controller.B2SSetData 19,1
    Controller.B2SSetData 20,1
    Controller.B2SSetData 21,1
    Controller.B2SSetData 22,1
    Controller.B2SSetData 23,1
    Controller.B2SSetData 24,0
  case 3:
    Controller.B2SSetData Int(RND*9)+9, 1
    Controller.B2SSetData Int(RND*9)+9, 1
    Controller.B2SSetData Int(RND*9)+9, 0
    Controller.B2SSetData Int(RND*9)+9, 0
  case 4:
    Controller.B2SSetData Int(RND*9)+9, 1
    Controller.B2SSetData Int(RND*9)+9, 0
  case 5:
    Controller.B2SSetData 19,1
    Controller.B2SSetData 20,1
    Controller.B2SSetData 21,0
    Controller.B2SSetData 22,1
    Controller.B2SSetData 23,0
  end Select
End Sub

Sub BonusTimer_Init()

End Sub

Sub ScoreTimer_Timer()

End Sub

Sub Drain_Init()

End Sub

Sub Drain_Unhit()

End Sub

Sub BallRelease_Hit()

End Sub

Sub BallRelease_Timer()

End Sub

Sub BallRelease_Unhit()

End Sub

Sub Scorereel1_Timer()

End Sub

Sub Scorereel1_Init()

End Sub

Sub Lbulb14_Init()

End Sub

Sub Lbulb14_Timer()

End Sub

Sub gameovertxt_Timer()

End Sub

Sub gameovertxt_Init()

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

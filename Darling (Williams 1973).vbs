'   -------------------------------------------------
'   Darling Williams 1973, based on Gottlieb EM 4 player VPX table blank with options menu
'   -------------------------------------------------
'  
'   BorgDog, 2018
'
'   High Score sticky routines from mfuegemann's Fast Draw VPX table
'       - flippers to change letter, Start to Select
'   Ball control script from rothbauerw
'       - press C during play to control ball, use arrows to move the ball
'   Option menu and player up light rotation borrowed from loserman76 and gnance
'       - hold down left shift during Game Over to bring up menu
'   Ball shadows from ninuzzu
'       - set option below to enable or disable
'   primitives from Dark, zany, sliderpoint, hauntfreaks and I'm sure others
'
'   Layer usage generally
'       1 - most stuff
'       2 - triggers
'       3 - under apron walls and ramps
'       4 - options menu
'       5 - hi score sticky
'       6 - GI lighting
'       7 - plastics
'       8 - insert and bumper lighting
'
'   Basic DOF config, may not be accurate and definitely need to change to match a particular game
'       101 Left Flipper, 102 Right Flipper,
'       103 Left sling/kicker, 104 Left sling flasher, 105 right sling/kicker, 106 right sling flasher,
'       107 Bumper1, 108 Bumper1 flasher, 109 bumper2, 110 bumper2 Flasher
'       134 Drain, 135 Ball Release
'       136 Shooter Lane/launch ball, 137 credit light, 138 knocker, 139 knocker/kicker Flasher
'       141 Chime1-10s, 142 Chime2-100s, 143 Chime3-1000s
'
'   -------------------------------------------------
 
Option Explicit
Randomize
 
Const cGameName = "darling"
Const Ballsize = 50
 
 
'**********************************************************
'********       OPTIONS     *******************************
'**********************************************************
 
Dim BallShadows: Ballshadows=1          '******************set to 1 to turn on Ball shadows
Dim FlipperShadows: FlipperShadows=1  '***********set to 1 to turn on Flipper shadows
 
Dim operatormenu, options
Dim bumperlitscore
Dim bumperoffscore
Dim balls, TempPlayerUp, rotatortemp
Dim replays, freeplay
Dim Replay1Table(3), Replay2Table(3), Replay3Table(3)
dim replay1, replay2, replay3
Dim Add10, Add100, Add1000
Dim hisc, LeftBalls, RightBalls
Dim maxplayers, players, player
Dim credit, ebmode
Dim score(4), bonus, extraadvance, chimescount, dbonus
Dim state
Dim tilt, tiltsens
Dim ballinplay
Dim matchnumb
dim rstep, lstep
Dim rep(4)
Dim rst
Dim eg
Dim bell
Dim i,j, ii, objekt, light
Dim awardcheck
 
On Error Resume Next
ExecuteGlobal GetTextFile("controller.vbs")
If Err Then MsgBox "Can't open controller.vbs"
On Error Goto 0
 
 
sub Darling_init
    LoadEM
    maxplayers=2
    Replay1Table(1)=47000
    Replay2Table(1)=63000
    Replay3Table(1)=79000
    Replay1Table(2)=64000
    Replay2Table(2)=73000
    Replay3Table(2)=99000
    hideoptions
    player=1
    RotatorTemp=1
    balls=5         'SET DEFAULT VALUES PRIOR TO LOADING FROM SAVED FILE
    ebmode=0
    freeplay=0
    leftballs=3
    rightballs=2
    hisc=10000
    HSA1=4
    HSA2=15
    HSA3=7
    matchnumb=0
    loadhs      'LOADS SAVED OPTIONS AND HIGH SCORE IF EXISTING
    UpdatePostIt    'UPDATE HIGH SCORE STICKY
    for i = 1 to maxplayers
        EVAL("ScoreReel"&i).setvalue score(i)
        EVAL("Reel100K"&i).setvalue(int(score(i)/100000))
    next
 
    if balls=3 then
        replays=1
      else
        replays=2
    end if
 
    bipreel.setvalue 0
    if freeplay=0 then credittxt.setvalue(credit)
    Replay1=Replay1Table(Replays)
    Replay2=Replay2Table(Replays)
    Replay3=Replay3Table(Replays)
    if ebmode = 0 then
        RepCard.image = "ReplayCard"&Balls&"Balls"
      else
        Repcard.image = "ReplayCard"&balls&"BallsEB"
    end if
    OptionBalls.image="OptionsBalls"&Balls
    OptionReplays.image="OptionReplays"&ebmode
    OptionFreeplay.image="OptionsFreeplay"&freeplay
    bumperlitscore=100
    bumperoffscore=100
    if balls=3 then
        extraadvance=4
      else
        extraadvance=3
    end if
 
    if ebmode=0 then
        InstCard.image="InstCardReplay"
      else
        InstCard.image="InstCardEB"
    end if
 
    If B2SOn then
        for each objekt in backdropstuff : objekt.visible = 0 : next
    End If
    startGame.enabled=true
    tilt=false
    turnoff
 
 
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
    KcaptiveL.uservalue=1
    KcaptiveL.timerenabled=1
    Drain.CreateBall
End sub
 
Sub KcaptiveL_timer
        Select Case (me.uservalue)
            Case 1:
                if leftballs>0 then
                    KcaptiveL.createball
                    KcaptiveL.kick 180,1
                end if
                if rightballs>0 then
                    KcaptiveR.createball
                    kcaptiveR.kick 180,1
                end if
            Case 2:
                if leftballs>1 then
                    KcaptiveL.createball
                    KcaptiveL.kick 180,1
                end if
                if rightballs>1 then
                    KcaptiveR.createball
                    kcaptiveR.kick 180,1
                end if
            Case 3:
                if leftballs>2 then
                    KcaptiveL.createball
                    KcaptiveL.kick 180,1
                end if
                if rightballs>2 then
                    KcaptiveR.createball
                    kcaptiveR.kick 180,1
                end if
            Case 4:
                if leftballs>3 then
                    KcaptiveL.createball
                    KcaptiveL.kick 180,1
                end if
                if rightballs>3 then
                    KcaptiveR.createball
                    kcaptiveR.kick 180,1
                end if
            Case 5:
                if leftballs>4 then
                    KcaptiveL.createball
                    KcaptiveL.kick 180,1
                end if
                if rightballs>4 then
                    KcaptiveR.createball
                    kcaptiveR.kick 180,1
                end if
            Case 6:
                me.timerenabled=0
                PairedlampTimer.enabled=1
        End Select
    me.uservalue=me.uservalue+1
end sub
 
sub startGame_timer
    PlaySoundAtVol "poweron", Plunger, 1
    lightdelay.enabled=true
    me.enabled=false
end sub
 
sub lightdelay_timer
    gamov.text="GAME OVER"
    if ebmode=0 then
		if matchnumb=0 then
			matchtxt.text="00"
			If B2SOn then Controller.B2SSetMatch 34,100
		  else
			matchtxt.text=matchnumb*10
			If B2SOn then Controller.B2SSetMatch 34,Matchnumb*10
		end if
    end if
    If (credit>0 and freeplay=0) or freeplay=1 then
        DOF 137, DOFOn
        Lcredit.state=1
    end if
    For each light in GIlights:light.state=1:Next
    if B2SOn then
        for i = 1 to maxplayers
            Controller.B2SSetScorePlayer i, Score(i) MOD 100000
        next
        Controller.B2SSetData 1,1
        if freeplay=0 then Controller.B2ssetCredits Credit

        Controller.B2SSetGameOver 35,1
    end if
    me.enabled=0
end Sub
 
 
 
Sub Darling_KeyDown(ByVal keycode)
   
    if keycode = 46 then' C Key
        If contball = 1 Then
            contball = 0
          Else
            contball = 1
        End If
    End If
 
    if keycode = 48 then 'B Key
        If bcboost = 1 Then
            bcboost = bcboostmulti
          Else
            bcboost = 1
        End If
    End If
 
    if keycode = 203 then Cleft = 1' Left Arrow
 
    if keycode = 200 then Cup = 1' Up Arrow
 
    if keycode = 208 then Cdown = 1' Down Arrow
 
    if keycode = 205 then Cright = 1' Right Arrow
 
    if keycode=AddCreditKey then
        if freeplay=1 or credit>24 then
            PlaySoundAtVol "coinreturn", Drain, 1
          else
            PlaySoundAtVol "coinin" , Drain, 1
            coindelay.enabled=true
        end if
    end if
 
    if keycode=StartGameKey and (Credit>0 or freeplay=1) And Not lightdelay.enabled and Not startGame.enabled and Not HSEnterMode=true then
      if state=false then
        player=1
        if freeplay=0 then
            credit=credit-1
            if credit < 1 then
                DOF 137, DOFOff
                Lcredit.state=0
            end if
            credittxt.setvalue(credit)
        End if
        PlaySound "cluper"
        ballinplay=1
        If B2SOn Then
            if freeplay=0 then Controller.B2ssetCredits Credit
            Controller.B2ssetballinplay 32, Ballinplay
            Controller.B2ssetplayerup 30, 1
            Controller.B2ssetcanplay 31, 1
            Controller.B2SSetGameOver 0
            Controller.B2SSetScorePlayer 5, hisc
        End If
        for i = 1 to maxplayers
            score(i)=0
            rep(i)=0
            EVAL("Pup"&i).state=0
            EVAL("PupN"&i).state=0
            EVAL("ScoreReel"&i).resettozero
            EVAL("Reel100K"&i).setvalue 0
            If B2SOn then Controller.B2SSetScorePlayer i, score(i)
        next
        Pup1.state=1
        PupN1.state=1
'       PlaySound("RotateThruPlayers"),0,.05,0,0.25
'       TempPlayerUp=Player
'       PlayerUpRotator.enabled=true
        tilt=false
        state=true
        PlaySound "GameStart"
        players=1
        CanPlay1.state=1
        rst=0
        newgame.enabled=true
      else if state=true and players < maxplayers and Ballinplay=1 then
        if freeplay=0 then
            credit=credit-1
            if credit < 1 then
                DOF 137, DOFOff
                Lcredit.state=0
            end if
            credittxt.setvalue(credit)
        End if
        players=players+1
        CanPlay2.state=1
        CanPlay1.state=0
 
        If B2SOn then
            if freeplay=0 then Controller.B2ssetCredits Credit
            Controller.B2ssetcanplay 31, players
        End If
        playsound "cluper"
       end if
      end if
    end if
 
    If keycode = PlungerKey Then
        Plunger.PullBack
        PlaySoundAtVol "plungerpull", Plunger, 1
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
                if ebmode=0 then
                    RepCard.image="ReplayCard5Balls"
                  else
                    RepCard.image="ReplayCard5BallsEB"
                end if
              else
                Balls=3
                if ebmode=0 then
                    RepCard.image="ReplayCard3Balls"
                  else
                    RepCard.image="ReplayCard3BallsEB"
                end if
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
            if ebmode=0 then
                ebmode=1
                instcard.image="InstCardEB"
                OptionReplays.image="OptionReplays1"
              else
                ebmode=0
                instcard.image="InstCardReplay"
                OptionReplays.image="OptionReplays0"
            end if
        Case 4:
            OperatorMenu=0
            savehs
            HideOptions
      End Select
    End If
 
    If HSEnterMode Then HighScoreProcessKey(keycode)
 
  if tilt=false and state=true then
    If keycode = LeftFlipperKey Then
        LeftFlipper.RotateToEnd
        LgiLeftFlipper.state=1
        PlaySoundAtVol SoundFXDOF("fx_flipperup",101,DOFOn,DOFFlippers), LeftFlipper, 1
        PlaySoundAtVolLoops "Buzz", LeftFlipper, 0.01, -1
    End If
   
    If keycode = RightFlipperKey Then
        RightFlipper.RotateToEnd
        LgiRightFlipper.state=1
        PlaySoundAtVol SoundFXDOF("fx_flipperup",102,DOFOn,DOFFlippers), RightFlipper, 1
        PlaySoundAtVolLoops "Buzz1", RightFlipper, 0.01, -1
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
    OptionReplays.visible = True
    OptionFreeplay.visible = True
End Sub
 
Sub HideOptions
    for each objekt In OptionMenu
        objekt.visible = false
    next
End Sub
 
 
Sub Darling_KeyUp(ByVal keycode)
 
    if keycode = 203 then Cleft = 0' Left Arrow
 
    if keycode = 200 then Cup = 0' Up Arrow
 
    if keycode = 208 then Cdown = 0' Down Arrow
 
    if keycode = 205 then Cright = 0' Right Arrow
 
    If keycode = PlungerKey Then
        Plunger.Fire
        if ballhome.BallCntOver=1 then
            PlaySoundAt "plunger", Plunger                  'PLAY WHEN BALL IS HIT
          else
            PlaySoundAt "plungerreleasefree", Plunger       'PLAY WHEN NO BALL TO PLUNGE
        end if
    End If
 
    if keycode = LeftFlipperKey then
        OperatorMenuTimer.Enabled = false
    end if
 
   If tilt=false and state=true then
    If keycode = LeftFlipperKey Then
        LeftFlipper.RotateToStart
        LgiLeftFlipper.duration 1, 100, 0
        PlaySoundAtVol SoundFXDOF("fx_flipperdown",101,DOFOff,DOFFlippers), LeftFlipper, 1
        StopSound "Buzz"
    End If
   
    If keycode = RightFlipperKey Then
        RightFlipper.RotateToStart
        LgiRightFlipper.duration 1, 100, 0
        PlaySoundAtVol SoundFXDOF("fx_flipperdown",102,DOFOff,DOFFlippers), RightFlipper, 1
        StopSound "Buzz1"
    End If
   End if
End Sub
 
 
Dim Cup, Cdown, Cleft, Cright, contball, contballinplay, ControlBall, bcboost
Dim bcvel, bcyveloffset, bcboostmulti
 
bcboost = 1 'Do Not Change - default setting
bcvel = 4 'Controls the speed of the ball movement
bcyveloffset = -0.01 'Offsets the force of gravity to keep the ball from drifting vertically on the table, should be negative
bcboostmulti = 3 'Boost multiplier to ball veloctiy (toggled with the B key)
 
Sub BallControl_Timer()
    If Contball and ContBallInPlay then
        If Cright = 1 Then
            ControlBall.velx = bcvel*bcboost
          ElseIf Cleft = 1 Then
            ControlBall.velx = - bcvel*bcboost
          Else
            ControlBall.velx=0
        End If
        If Cup = 1 Then
            ControlBall.vely = -bcvel*bcboost
          ElseIf Cdown = 1 Then
            ControlBall.vely = bcvel*bcboost
          Else
            ControlBall.vely= bcyveloffset
        End If
    End If
End Sub
 
sub flippertimer_timer()
'testbox.text=bonus
'testbox1.text=LeftBalls
'testbox2.text=RightBalls
    Pgate.rotx = Gate.currentangle*.6
    Pgate1.rotx = Gate1.currentangle*.6
    PgateRollUnder.roty=GateRollUnder.currentangle
	FlipperR.RotZ = RightFlipper.currentAngle
	FlipperL.RotZ = LeftFlipper.CurrentAngle
 
    if FlipperShadows=1 then
        FlipperLSh.RotZ = LeftFlipper.currentangle
        FlipperRSh.RotZ = RightFlipper.currentangle
    end if
 
end sub
 
 
Sub PairedlampTimer_timer
    Lbonus20R.state = Lbonus20.state
    LDbonusR.state = LDbonus.state
    LspecialL1.state = LspecialL.state
    LspecialR1.state = LspecialR.state
    L300WL.state = LspecialL.state
    L300WLR.state = LspecialR.state
 
    '********************track leftballs and RightBalls
 
    select case (TGcaptiveL.BallCntOver+TGcaptiveL1.BallCntOver+TGcaptiveL2.BallCntOver)
        Case 0:
            LeftBalls=0
            RightBalls=5
        Case 1:
            select case (TGcaptiveR.BallCntOver+TGcaptiveR1.BallCntOver+TGcaptiveR2.BallCntOver)
                Case 2:
                    LeftBalls=2
                    RightBalls=3
                Case 3:
                    LeftBalls=1
                    RightBalls=4
            end Select
        case 2:
            LeftBalls=3
            RightBalls=2
        Case 3:
            select case (TGcaptiveR.BallCntOver)
                Case 0:
                    LeftBalls=5
                    RightBalls=0
                Case 1:
                    LeftBalls=4
                    RightBalls=1
            end Select
    end Select
 
end sub
 
sub coindelay_timer
    addcredit
    coindelay.enabled=false
end sub
 
Sub addcredit
    if freeplay=0 then
      credit=credit+1
      DOF 137, DOFOn
      Lcredit.state=1
      if credit>25 then credit=25
      credittxt.setvalue(credit)
      If B2SOn Then Controller.B2ssetCredits Credit
    end if
End sub
 
Sub Drain_Hit()
    DOF 134, DOFPulse
    PlaySoundAt "drain", Drain
    me.timerenabled=1
End Sub
 
Sub Drain_timer
    scorebonus.enabled=true
    me.timerenabled=0
End Sub
 
Sub BallHome_Hit()                  '******* for ball control script
    Set ControlBall = ActiveBall   
    contballinplay = true
End Sub
 
sub ballhome_unhit
    DOF 136, DOFPulse
end sub
 
Sub EndControl_Hit()                '******* for ball control script
    contballinplay = false
End Sub
 
sub ballrel_hit
    lshootagain.state=0
end sub
 
sub scorebonus_timer            '************add bonus scoring if applicable
   if tilt=true then
        bonus=0
    Else
        if LDbonus.state=1 then
            dbonus=2
          else
            dbonus=1
        End if
        if bonus>0 then
            if bonus>1 or dbonus=2 then
                chimescount=2
              else
                chimescount=1
            end if
            chimestimer.interval=1
            chimestimer.enabled=true
            playsound "motor2", 0, 0.04
            me.enabled=False
          else
            bonus=0
            for each light in bonuslights: light.state=0: next
        end if
    end if
    if bonus=0 and chimestimer.enabled=0 then
         if lshootagain.state=lightstateon and tilt=false then
              newball
              ballreltimer.enabled=true
            else
              if players=1 or player=players then
                player=1
               Else
                player=player+1
              end if
              If B2SOn then Controller.B2ssetplayerup 30, player
              nextball
         end if
         me.enabled=false
    end if
End sub
 
sub chimestimer_timer
    chimestimer.interval=100
    addpoints 1000
 
    if (dbonus=2 and chimescount=1) or dbonus=1 then
        EVAL("Lbonus"&bonus).state=0
        bonus=bonus-1
        if bonus >0 then
            EVAL("Lbonus"&bonus).state=1
            if bonus=19 then Lbonus10.state=1
        end if
    end if
    chimescount = chimescount - 1
    if chimescount<1 then
        me.enabled=false
        ScoreBonus.enabled=True
    end if
end Sub
 
 
sub newgame_timer
    eg=0
    if LtopL.state=0 and LtopR.state=0 then LtopL.state=1
'   for each light in GILights: light.state=1: next
    ldbonus.state=0
    Lshootagain.state=0
    tilttxt.text=" "
    gamov.text=" "
    for i=1 to 2
        EVAL("Bumper"&i).hashitevent = 1
    Next
 
    If B2SOn then
        Controller.B2SSetGameOver 35,0
        Controller.B2SSetTilt 33,0
        Controller.B2ssetdata 1, 1
        if ebmode=0 then Controller.B2SSetMatch 34,0
    End If
    bipreel.setvalue 1
    if ebmode=0 then matchtxt.text=" "
    newball
    ballrelTimer.enabled=true
    newgame.enabled=false
end sub
 
 
sub newball
    if ballinplay=Balls then LDbonus.state=1
    bonus=0
    if LeftBalls>RightBalls then
        Lspecialr.state=1
        Lspeciall.state=0
      else
        Lspecialr.state=0
        Lspeciall.state=1
    end if
End Sub
 
 
sub nextball
    if tilt=true then
      for i=1 to 2
        EVAL("Bumper"&i).hashitevent = 1
      Next
      for each light in GILights: light.state=1: next
      for each light in BonusLights: light.state=0: next
      tilt=false
      tilttxt.text=" "
      If B2SOn then
        Controller.B2SSetTilt 33,0
        Controller.B2ssetdata 1, 1
      End If
    end if
    if player=1 then ballinplay=ballinplay+1
    if ballinplay>balls then
        playsound "motor2", 0, 0.04
        eg=1
        ballreltimer.enabled=true
      else
        For each objekt in PlayerHuds
            objekt.State = 0
        next
        For each objekt in PlayerHUDScores
            objekt.state=0
        next
        PlayerHuds(Player-1).State =1
        PlayerHUDScores(Player-1).state=1
'       PlaySound("RotateThruPlayers"),0,.05,0,0.25
'       TempPlayerUp=Player
'       PlayerUpRotator.enabled=true
        if state=true then
          newball
          ballreltimer.enabled=true
        end if
        bipreel.setvalue ballinplay
        If B2SOn then Controller.B2ssetballinplay 32, Ballinplay
    end if
End Sub
 
sub ballreltimer_timer
  dim hiscstate
  if eg=1 then
      hiscstate=0
      if ebmode=0 then matchnum
      state=false
      bipreel.setvalue 0
      gamov.text="GAME OVER"
      gamov.timerenabled=1
      tilttxt.timerenabled=1
        CanPlay1.state=0
        CanPlay2.state=0
      for i=1 to maxplayers
        if score(i)>hisc then
            hisc=score(i)
            hiscstate=1
        end if
        EVAL("Pup"&i).state=0
        EVAL("PupN"&i).state=0
      next
      if hiscstate=1 then
            HighScoreEntryInit()
            HStimer.uservalue = 0
            HStimer.enabled=1
      end if
      UpdatePostIt
      savehs
      If B2SOn then
        Controller.B2SSetGameOver 35,1
        Controller.B2ssetballinplay 32, 0
        Controller.B2SSetScorePlayer 5, hisc
        Controller.B2ssetcanplay 31, 0
        Controller.B2ssetcanplay 30, 0
      End If
      ballreltimer.enabled=false
  else
    Drain.kick 55,15,0
    ballreltimer.enabled=false
    PlaySoundAtVol SoundFXDOF("nextball",135,DOFPulse,DOFContactors), Drain, 1
  end if
end sub
 
Sub HStimer_timer
    PlaySoundAtVol SoundFXDOF("knock",138,DOFPulse,DOFKnocker), BulbTop11, 1
    DOF 139,DOFPulse
    HStimer.uservalue=HStimer.uservalue+1
    if HStimer.uservalue=3 then me.enabled=0
end sub
 
Sub PlayerUpRotator_timer()
        If RotatorTemp<maxplayers+1 then
            TempPlayerUp=TempPlayerUp+1
            If TempPlayerUp>maxplayers then
                TempPlayerUp=1
            end if
            For each objekt in PlayerHuds
                objekt.state = 0
            next
            For each objekt in PlayerHUDScores
                objekt.state=0
            next
            PlayerHuds(TempPlayerUp-1).State = 1
            PlayerHUDScores(TempPlayerUp-1).state=1
            If B2SOn Then Controller.B2SSetPlayerUp TempPlayerUp
        else
            if B2SOn then Controller.B2SSetPlayerUp Player
            PlayerUpRotator.enabled=false
            RotatorTemp=1
            For each objekt in PlayerHuds
                objekt.State = 0
            next
            For each objekt in PlayerHUDScores
                objekt.state=0
            next
            PlayerHuds(Player-1).State = 1
            PlayerHUDScores(Player-1).state=1
        end if
        RotatorTemp=RotatorTemp+1
end sub
 
sub matchnum
    if matchnumb=0 then
        matchtxt.text="00"
        If B2SOn then Controller.B2SSetMatch 34,100
      else
        matchtxt.text=matchnumb*10
        If B2SOn then Controller.B2SSetMatch 34,Matchnumb*10
    end if
    For i=1 to players
        if (matchnumb*10)=(score(i) mod 100) then
          addcredit
          PlaySoundAtVol SoundFXDOF("knock",138,DOFPulse,DOFKnocker), BulbTop11, 1
          DOF 139,DOFPulse
        end if
    next
end sub
 
'********** Bumpers
 
 
Sub Bumper1_Hit
    if tilt=false then
        PlaySoundAtVol SoundFXDOF("fx_bumper4", 107, DOFPulse, DOFContactors), Bumper1, 1
        DOF 108,DOFPulse
        if bumperlight1.state=1 then
            addscore bumperlitscore
          else
            addscore bumperoffscore
        end if
    end if
End Sub
 
Sub Bumper2_Hit
    if tilt=false then
        PlaySoundAtVol SoundFXDOF("fx_bumper4", 109, DOFPulse, DOFContactors), Bumper2, 1
        DOF 110,DOFPulse
        if BumperLight2.state=1 then
            addscore bumperlitscore
          else
            addscore bumperoffscore
        end if 
    end if
End Sub
 
Function SkirtAX(bumper, bumperball)
    skirtAX=cos(skirtA(bumper,bumperball))*(SkirtTilt)      'x component of angle
    if (bumper.y<bumperball.y) then skirtAX=skirtAX*-1      'adjust for ball hit bottom half
End Function
 
Function SkirtAY(bumper, bumperball)
    skirtAY=sin(skirtA(bumper,bumperball))*(SkirtTilt)      'y component of angle
    if (bumper.x>bumperball.x) then skirtAY=skirtAY*-1      'adjust for ball hit left half
End Function
 
Function SkirtA(bumper, bumperball)
    dim hitx, hity, dx, dy
    hitx=bumperball.x
    hity=bumperball.y
    dy=Abs(hity-bumper.y)                   'y offset ball at hit to center of bumper
    if dy=0 then dy=0.0000001
    dx=Abs(hitx-bumper.x)                   'x offset ball at hit to center of bumper
    skirtA=(atn(dx/dy)) '/(PI/180)          'angle in radians to ball from center of Bumper1
End Function
 
'************** Kickers
 
 
Sub Klsling_hit
    if L300WL.state=1 then
        addscore 300
      else
        addscore 30
    end if
    me.uservalue=1
    me.timerenabled=1
End Sub
 
Sub Klsling_timer
    select case me.uservalue
      case 4:
        PlaySoundAtVol SoundFXDOF("holekick",103,DOFPulse,DOFContactors), Klsling, 1
        DOF 139, DOFPulse
        me.kick 23.3,37
        PkickarmL.rotx=20
      case 6:
        PkickarmL.rotx=0
        me.timerenabled=0
    end Select
    me.uservalue=me.uservalue+1
End Sub
 
Sub Krsling_hit
    if L300WLR.state=1 then
        addscore 300
      else
        addscore 30
    end if
    me.uservalue=1
    me.timerenabled=1
End Sub
 
Sub Krsling_timer
    select case me.uservalue
      case 4:
        PlaySoundAtVol SoundFXDOF("holekick",105,DOFPulse,DOFContactors), Krsling, 1
        DOF 139, DOFPulse
        me.kick -23.3,37
        PkickarmR.rotx=20
      case 6:
        PkickarmR.rotx=0
        me.timerenabled=0
    end Select
    me.uservalue=me.uservalue+1
End Sub
 
 
 
 
 
 
 
 
'********** Dingwalls - animated - timer 50
 
sub dingwalla_hit
    if state=true and tilt=false then addscore 50
    SlingA.visible=0
    SlingA1.visible=1
    me.uservalue=1
    Me.timerenabled=1
end sub
 
sub dingwalla_timer                                 'default 50 timer
    select case me.uservalue
        Case 1: SlingA1.visible=0: SlingA.visible=1
        case 2: SlingA.visible=0: SlingA2.visible=1
        Case 3: SlingA2.visible=0: SlingA.visible=1: Me.timerenabled=0
    end Select
    me.uservalue=me.uservalue+1
end sub
 
 
sub dingwallb_hit
    if state=true and tilt=false then addscore 50
    SlingB.visible=0
    SlingB1.visible=1
    me.uservalue=1
    Me.timerenabled=1
end sub
 
sub dingwallb_timer                                 'default 50 timer
    select case me.uservalue
        Case 1: Slingb1.visible=0: SlingB.visible=1
        case 2: SlingB.visible=0: Slingb2.visible=1
        Case 3: Slingb2.visible=0: SlingB.visible=1: Me.timerenabled=0
    end Select
    me.uservalue=me.uservalue+1
end sub
 
sub dingwallc_hit
    if state=true and tilt=false then addscore 10
    SlingC.visible=0
    SlingC1.visible=1
    me.uservalue=1
    Me.timerenabled=1
end sub
 
sub dingwallc_timer                                 'default 50 timer
    select case me.uservalue
        Case 1: SlingC1.visible=0: SlingC.visible=1
        case 2: SlingC.visible=0: SlingC2.visible=1
        Case 3: SlingC2.visible=0: SlingC.visible=1: Me.timerenabled=0
    end Select
    me.uservalue=me.uservalue+1
end sub
 
sub dingwalld_hit
    if state=true and tilt=false then addscore 10
    SlingD.visible=0
    SlingD1.visible=1
    me.uservalue=1
    Me.timerenabled=1
end sub
 
sub dingwalld_timer                                 'default 50 timer
    select case me.uservalue
        Case 1: SlingD1.visible=0: SlingD.visible=1
        case 2: SlingD.visible=0: SlingD2.visible=1
        Case 3: SlingD2.visible=0: SlingD.visible=1: Me.timerenabled=0
    end Select
    me.uservalue=me.uservalue+1
end sub
 
sub dingwalle_hit
    if state=true and tilt=false then addscore 50
    SlingE.visible=0
    SlingE1.visible=1
    me.uservalue=1
    Me.timerenabled=1
end sub
 
sub dingwalle_timer                                 'default 50 timer
    select case me.uservalue
        Case 1: SlingE1.visible=0: SlingE.visible=1
        case 2: SlingE.visible=0: SlingE2.visible=1
        Case 3: SlingE2.visible=0: SlingE.visible=1: Me.timerenabled=0
    end Select
    me.uservalue=me.uservalue+1
end sub
 
sub dingwallf_hit
    if state=true and tilt=false then addscore 50
    SlingF.visible=0
    SlingF1.visible=1
    me.uservalue=1
    Me.timerenabled=1
end sub
 
sub dingwallf_timer                                 'default 50 timer
    select case me.uservalue
        Case 1: SlingF1.visible=0: SlingF.visible=1
        case 2: SlingF.visible=0: SlingF2.visible=1
        Case 3: SlingF2.visible=0: SlingF.visible=1: Me.timerenabled=0
    end Select
    me.uservalue=me.uservalue+1
end sub
 
sub dingwallg_hit
    SlingG.visible=0
    SlingG1.visible=1
    me.uservalue=1
    Me.timerenabled=1
end sub
 
sub dingwallg_timer                                 'default 50 timer
    select case me.uservalue
        Case 1: SlingG1.visible=0: SlingG.visible=1
        case 2: SlingG.visible=0: SlingG2.visible=1
        Case 3: SlingG2.visible=0: SlingG.visible=1: Me.timerenabled=0
    end Select
    me.uservalue=me.uservalue+1
end sub
 
sub dingwallh_hit
    SlingH.visible=0
    SlingH1.visible=1
    me.uservalue=1
    Me.timerenabled=1
end sub
 
sub dingwallh_timer                                 'default 50 timer
    select case me.uservalue
        Case 1: SlingH1.visible=0: SlingH.visible=1
        case 2: SlingH.visible=0: SlingH2.visible=1
        Case 3: SlingH2.visible=0: SlingH.visible=1: Me.timerenabled=0
    end Select
    me.uservalue=me.uservalue+1
end sub
 
'********** Triggers    
 
sub TGoutL_hit
    DOF 113, DOFPulse
    addbonus
end sub    
 
sub TGoutR_hit
    DOF 114, DOFPulse
    addbonus
end sub
 
sub TGbuttonL_hit
    addbonus
end sub
 
sub TGbuttonR_hit
    addbonus
end sub
 
sub TGtopL_hit   '***** top L rollover
    DOF 111, DOFPulse
    if LtopL.state=1 then
        addbonus
        me.uservalue=1
        me.timerenabled=1
      else
        addbonus
    end if
end sub    
 
sub TGtopL_timer
    addbonus
    me.uservalue=me.uservalue+1
    if me.uservalue=extraadvance then me.timerenabled=0
end sub
 
sub TGtopR_hit   '***** top R rollover
    DOF 112, DOFPulse
    if LtopR.state=1 then
        addbonus
        me.uservalue=1
        me.timerenabled=1
    else
        addbonus
    end if
end sub
 
sub TGtopR_timer
    addbonus
    me.uservalue=me.uservalue+1
    if me.uservalue=extraadvance then me.timerenabled=0
end sub
 
sub TGcaptiveL3_hit
    if not tilt then
        if TGcaptiveR3.timerenabled then addbonus
        me.timerenabled=1
        if LspecialL.state=1 and (TGcaptiveL.ballcntover+TGcaptiveL1.ballcntover+TGcaptiveL2.ballcntover=3) then awardspecial
    end if
end sub
 
sub TGcaptiveL3_timer
    me.timerenabled=0
end sub
 
sub TGcaptiveR3_hit
    if not tilt then
        if TGcaptiveL3.timerenabled then addbonus
        me.timerenabled=1
        if LspecialR.state=1 and (TGcaptiveR.ballcntover+TGcaptiveR1.ballcntover+TGcaptiveR2.ballcntover=3) then awardspecial
    end if
end sub
 
sub TGcaptiveR3_timer
    me.timerenabled=0
end sub
 
sub awardspecial
    LShootAgain.state=1
    if lspecialL.state=1 then
        lspeciall.state=0
        lspecialr.state=1
      else
        lspeciall.state=1
        lspecialr.state=0
    end if
end sub
 
'********************* Gates
 
sub GateRollUnder_hit
    addbonus
end sub
 
 
'********** Scoring
 
sub addbonus
    bonus=bonus+1
    if bonus>20 then bonus=20
    EVAL("Lbonus"&bonus).state=1
    if bonus>1 then EVAL("Lbonus"&(bonus-1)).state=0
    if bonus=11 then Lbonus10.state=1
    if bonus=20 then Lbonus10.state=0
    addscore 100
end sub
 
sub addscore(points)
  if tilt=false and state=true then
    If points = 10 then
        matchnumb=matchnumb+1
        if matchnumb>9 then matchnumb=0
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
    score(player)=score(player)+points
    EVAL("ScoreReel"&player).addvalue(points)
    If B2SOn Then Controller.B2SSetScorePlayer player, score(player) MOD 100000
 
    ' Sounds: there are 3 sounds: tens, hundreds and thousands
    If Points = 100 AND(Score(player) MOD 1000) \ 100 = 0 Then  'New 1000 reel
        PlaySoundAtVol SoundFXDOF("bell1000",143,DOFPulse,DOFChimes), chimesound, 1
      ElseIf Points = 10 AND(Score(player) MOD 100) \ 10 = 0 Then 'New 100 reel
        if LtopL.state=1 then
            LtopL.state=0
            LtopR.state=1
          else
            LtopL.state=1
            LtopR.state=0
        end if
        PlaySoundAtVol SoundFXDOF("bell100",142,DOFPulse,DOFChimes), chimesound, 1
      ElseIf points = 1000 Then
        PlaySoundAtVol SoundFXDOF("bell1000",143,DOFPulse,DOFChimes), chimesound, 1
      elseif Points = 100 Then
        PlaySoundAtVol SoundFXDOF("bell100",142,DOFPulse,DOFChimes), chimesound, 1
      Else
        if LtopL.state=1 then
            LtopL.state=0
            LtopR.state=1
          else
            LtopL.state=1
            LtopR.state=0
        end if
        PlaySoundAtVol SoundFXDOF("bell10",141,DOFPulse,DOFChimes), chimesound, 1
    End If
    checkreplays
End Sub
 
Sub checkreplays
    ' check replays and rollover
    if score(player)=>99999 then
        If B2SOn then Controller.B2SSetScoreRollover 24 + player, 1
        EVAL("Reel100K"&player).setvalue(int(score(player)/100000))
    End if
    if score(player)=>replay1 and rep(player)=0 then
        if ebmode=0 then
            addcredit
          else
            lshootagain.state=1
        end if
        rep(player)=1
        PlaySoundAtVol SoundFXDOF("knock",138,DOFPulse,DOFKnocker), BulbTop11, 1
        DOF 139, DOFPulse
    end if
    if score(player)=>replay2 and rep(player)=1 then
        if ebmode=0 then
            addcredit
          else
            lshootagain.state=1
        end if
        rep(player)=2
        PlaySoundAtVol SoundFXDOF("knock",138,DOFPulse,DOFKnocker), BulbTop11, 1
        DOF 139, DOFPulse
    end if
    if score(player)=>replay3 and rep(player)=2 then
        if ebmode=0 then
            addcredit
          else
            lshootagain.state=1
        end if
        rep(player)=3
        PlaySoundAtVol SoundFXDOF("knock",138,DOFPulse,DOFKnocker), BulbTop11, 1
        DOF 139, DOFPulse
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
    turnoff
End Sub
 
sub turnoff
    for i=1 to 2
        EVAL("Bumper"&i).hashitevent = 0
    Next
    for each light in GILights: light.state=0: next
    LeftFlipper.RotateToStart
    LgiLeftFlipper.duration 1, 100, 0
    StopSound "Buzz"
    DOF 101, DOFOff
    RightFlipper.RotateToStart
    LgiRightFlipper.duration 1, 100, 0
    StopSound "Buzz1"
    DOF 102, DOFOff
end sub    
 
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
 
'**************************************************************************
'                 Additional Positional Sound Playback Functions by DJRobX
'**************************************************************************
 
Sub PlaySoundAtVol(sound, tableobj, Vol)
        PlaySound sound, 1, Vol, AudioPan(tableobj), 0,0,0, 1, AudioFade(tableobj)
End Sub
 
'Set position at table object, vol, and loops manually.
 
Sub PlaySoundAtVolLoops(sound, tableobj, Vol, Loops)
        PlaySound sound, Loops, Vol, AudioPan(tableobj), 0,0,0, 1, AudioFade(tableobj)
End Sub
 
'*********************************************************************
'                     Supporting Ball & Sound Functions
'*********************************************************************
 
Function AudioFade(tableobj) ' Fades between front and back of the table (for surround systems or 2x2 speakers, etc), depending on the Y position on the table. "table1" is the name of the table
    Dim tmp
    tmp = tableobj.y * 2 / Darling.height-1
    If tmp > 0 Then
        AudioFade = Csng(tmp ^10)
    Else
        AudioFade = Csng(-((- tmp) ^10) )
    End If
End Function
 
Function AudioPan(tableobj) ' Calculates the pan for a tableobj based on the X position on the table. "table1" is the name of the table
    Dim tmp
    tmp = tableobj.x * 2 / Darling.width-1
    If tmp > 0 Then
        AudioPan = Csng(tmp ^10)
    Else
        AudioPan = Csng(-((- tmp) ^10) )
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
 
Const tnob = 7 ' total number of balls
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
            PlaySound("fx_ballrolling" & b), -1, (Vol(BOT(b))/2), AudioPan(BOT(b)), 0, Pitch(BOT(b)), 1, 0, AudioFade(BOT(b))
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
'           BALL SHADOW
'*****************************************
Dim BallShadow
BallShadow = Array (BallShadow1,BallShadow2,BallShadow3,BallShadow4,BallShadow5,BallShadow6,BallShadow7)
 
Sub BallShadowUpdate_timer()
    Dim BOT, b
    Dim maxXoffset
    maxXoffset=6
    BOT = GetBalls
    ' exit the Sub if no balls on the table
    If UBound(BOT) <1 Then Exit Sub
    ' render the shadow for each ball
    For b = 0 to UBound(BOT)
        BallShadow(b).X = BOT(b).X-maxXoffset*(1-(Bot(1).X)/(Darling.Width/2))
        BallShadow(b).Y = BOT(b).Y + 12
        If BOT(b).Z > 0 and BOT(b).Z < 30 Then
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
 
 
Sub a_Triggers_Hit (idx)
    playsound "sensor", 0,1,AudioPan(ActiveBall),0,0,0,1,AudioFade(ActiveBall)
End sub
 
Sub a_Pins_Hit (idx)
    PlaySound "pinhit_low", 0, Vol(ActiveBall), AudioPan(ActiveBall), 0, Pitch(ActiveBall), 0, 0, AudioFade(ActiveBall)
End Sub
 
Sub a_Targets_Hit (idx)
    PlaySound "target", 0, Vol(ActiveBall), AudioPan(ActiveBall), 0, Pitch(ActiveBall), 0, 0, AudioFade(ActiveBall)
End Sub
 
Sub a_DropTargets_Hit (idx)
    PlaySound "DTDrop", 0, Vol(ActiveBall), AudioPan(ActiveBall), 0, Pitch(ActiveBall), 0, 0, AudioFade(ActiveBall)
End Sub
 
Sub a_Metals_Thin_Hit (idx)
    PlaySound "metalhit_thin", 0, Vol(ActiveBall), AudioPan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
End Sub
 
Sub a_Metals_Medium_Hit (idx)
    PlaySound "metalhit_medium", 0, Vol(ActiveBall), AudioPan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
End Sub
 
Sub a_Metals2_Hit (idx)
    PlaySound "metalhit2", 0, Vol(ActiveBall), AudioPan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
End Sub
 
Sub a_Gates_Hit (idx)
    PlaySound "gate4", 0, Vol(ActiveBall), AudioPan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
End Sub
 
Sub Spinner_Spin
    PlaySound "fx_spinner", 0, .25, AudioPan(Spinner), 0.25, 0, 0, 1, AudioFade(Spinner)
End Sub
 
Sub a_Rubbers_Hit(idx)
    dim finalspeed
    finalspeed=SQR(activeball.velx * activeball.velx + activeball.vely * activeball.vely)
    If finalspeed > 20 then
        PlaySound "fx_rubber2", 0, Vol(ActiveBall), AudioPan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
    End if
    If finalspeed >= 6 AND finalspeed <= 20 then
        RandomSoundRubber()
    End If
End Sub
 
Sub a_Posts_Hit(idx)
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
        Case 1 : PlaySound "rubber_hit_1", 0, Vol(ActiveBall)*2, AudioPan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
        Case 2 : PlaySound "rubber_hit_2", 0, Vol(ActiveBall)*2, AudioPan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
        Case 3 : PlaySound "rubber_hit_3", 0, Vol(ActiveBall)*2, AudioPan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
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
        Case 1 : PlaySound "flip_hit_1", 0, Vol(ActiveBall)*2, AudioPan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
        Case 2 : PlaySound "flip_hit_2", 0, Vol(ActiveBall)*2, AudioPan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
        Case 3 : PlaySound "flip_hit_3", 0, Vol(ActiveBall)*2, AudioPan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
    End Select
End Sub
 
 
sub savehs
    savevalue "Darling", "credit", credit
    savevalue "Darling", "hiscore", hisc
    savevalue "Darling", "match", matchnumb
    savevalue "Darling", "score1", score(1)
    savevalue "Darling", "score2", score(2)
    savevalue "Darling", "score3", score(3)
    savevalue "Darling", "score4", score(4)
    savevalue "Darling", "balls", balls
    savevalue "Darling", "leftballs", leftballs
    savevalue "Darling", "rightballs", rightballs
    savevalue "Darling", "hsa1", HSA1
    savevalue "Darling", "hsa2", HSA2
    savevalue "Darling", "hsa3", HSA3
    savevalue "Darling", "freeplay", freeplay
    savevalue "Darling", "ebmode", ebmode
end sub
 
sub loadhs
    dim temp
    temp = LoadValue("Darling", "credit")
    If (temp <> "") then credit = CDbl(temp)
    temp = LoadValue("Darling", "hiscore")
    If (temp <> "") then hisc = CDbl(temp)
    temp = LoadValue("Darling", "match")
    If (temp <> "") then matchnumb = CDbl(temp)
    temp = LoadValue("Darling", "score1")
    If (temp <> "") then score(1) = CDbl(temp)
    temp = LoadValue("Darling", "score2")
    If (temp <> "") then score(2) = CDbl(temp)
    temp = LoadValue("Darling", "score3")
    If (temp <> "") then score(3) = CDbl(temp)
    temp = LoadValue("Darling", "score4")
    If (temp <> "") then score(4) = CDbl(temp)
    temp = LoadValue("Darling", "balls")
    If (temp <> "") then balls = CDbl(temp)
    temp = LoadValue("Darling", "leftballs")
    If (temp <> "") then leftballs = CDbl(temp)
    temp = LoadValue("Darling", "rightballs")
    If (temp <> "") then rightballs = CDbl(temp)
    temp = LoadValue("Darling", "hsa1")
    If (temp <> "") then HSA1 = CDbl(temp)
    temp = LoadValue("Darling", "hsa2")
    If (temp <> "") then HSA2 = CDbl(temp)
    temp = LoadValue("Darling", "hsa3")
    If (temp <> "") then HSA3 = CDbl(temp)
    temp = LoadValue("Darling", "freeplay")
    If (temp <> "") then freeplay = CDbl(temp)
    temp = LoadValue("Darling", "ebmode")
    If (temp <> "") then ebmode = CDbl(temp)
end sub
 
Sub Darling_Exit()
    Savehs
    turnoff
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
'   if HSA1="" then HSA1=25
'   if HSA2="" then HSA2=25
'   if HSA3="" then HSA3=25
'   UpdatePostIt
'ADD HSA1, HSA2 AND HSA3 TO SAVE AND LOAD VALUES FOR TABLE
'ADD A TIMER NAMED HighScoreFlashTimer WITH INTERVAL 100 TO TABLE
'SET HSSSCOREX BELOW TO WHATEVER VARIABLE YOU USE FOR HIGH SCORE.
'ADD OBJECTS TO PLAYFIELD (EASIEST TO JUST COPY FROM THIS TABLE)
'IMPORT POST-IT IMAGES
 
 
Dim HSA1, HSA2, HSA3
Dim HSEnterMode, hsLetterFlash, hsEnteredDigits(3), hsCurrentDigit, hsCurrentLetter
Dim HSArray  
Dim HSScoreM,HSScore100k, HSScore10k, HSScoreK, HSScore100, HSScore10, HSScore1, HSScorex   'Define 6 different score values for each reel to use
HSArray = Array("Postit0","postit1","postit2","postit3","postit4","postit5","postit6","postit7","postit8","postit9","postitBL","postitCM","Tape")
Const hsFlashDelay = 4
 
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
'   if showhisc=1 and showhiscnames=1 then
'       for each objekt in hiscname:objekt.visible=1:next
        HSName1.image = ImgFromCode(HSA1, 1)
        HSName2.image = ImgFromCode(HSA2, 2)
        HSName3.image = ImgFromCode(HSA3, 3)
'     else
'       for each objekt in hiscname:objekt.visible=0:next
'   end if
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
'               EMReelHSName1.SetValue HSA1:EMReelHSName2.SetValue HSA2
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


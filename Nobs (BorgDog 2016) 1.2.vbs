'   -------------------------------------------------
'   NOBS
'   -------------------------------------------------
'  
'   by BorgDog, 2016
'  
'   Scoring Rules
'
'       Scoring is by standard cribbage rules... mostly... with a "game" to 121 pegs.
'       Tracks cribbage games won, including skunks and double skunks on 2 player games.
'
'       Clear a hand to count cribbage pegs for that "hand" including "turn" card (roto target).
'       Hands resets after clearing and at end of ball, excpet at end of ball when shoot again is lit.
'
'       Hands are as follows and all include the "turn" card (RotoTarget) in the scoring.:
'            4 left side drops  (6, 5, 5, 4)
'            3 center drops and right side rollover (A, 2, 2, 3)
'            2 right drops, right outlane, and left inlane (6, 7, 8, 9)
'            3 top rollovers and left outlane (K, Q, J, 5)
'
'       Clearing all four hands in a single ball turns on Double pegs. Double pegs lights clear at end of ball always.
'
'       Clear K,Q,J and 7,8 rollovers to light star rollovers. Each lit star "pegs" 1 during "play".  
'       Clear any 6 lit wire rollovers to light extra ball (target) and open right outlane gate.
'
'       Clearing the "5" rollover opens left outlane gate, reset if hand 3 is complete or on next ball.
'
'       When a player reaches an end of the cribbage board (>120 pegs) that player scores a game win.
'           If opposing players pegs is less than 90 the player also scores a skunk (another game).
'           If opposing players pegs is less than 60 the player scores a double-skunk (4 games total)
'           New cribbage game, both players game pegs, reset to 0 when a game is won.
'
'       If any rollover or drop target hit during play matches the turn card (roto) a "pair for 2" is pegged.
'       If any rollover or drop target hit during play + turn card = 15, then a "15 for 2" is pegged, face cards count as 10.
'       Hitting roto pegs the displayed card (K,Q,J are worth 10, etc) and spins the roto.
'       Hitting bullseye target pegs 1 and spins the roto.
'       Skill shot pegs 2 for "heels"  (well should only be Jack but where is the fun in that)
'
'   This table uses a simplified and modified version of Wizards_Hat's Alphanumeric Display Tutorial table
'   for the system 1 type displays, expanded to include b2s. Thanks!
'
'   DOF config
'       101 Left Flipper, 102 Right Flipper,
'       105 right sling, 106 right sling flasher,
'       107 Bumper1, 108 Bumper1 flasher, 109 bumper2, 110 bumper2 Flasher
'       115 Roto-gear, 116 left DT reset, 117 mid DT reset, 118 Right DT Reset
'       134 Drain, 135 Ball Release
'       136 Shooter Lane/launch ball, 137 credit light, 138 knocker, 139 knocker Flasher
'       141 Chime1-10s, 142 Chime2-100s, 143 Chime3-1000s
'
'   -------------------------------------------------
 
Option Explicit
Randomize
 
Const cGameName = "Nobs_2016"
Const Ballsize = 50
 
' **********************  OPTIONS
 
dim chimevol: chimevol=1        'set between 0 and 1 to adjust how loud the chimes are, 0 is off, go below 0.1 for noticeable effect
Dim BallShadows: Ballshadows=1          '******************set to 1 to turn on Ball shadows
Dim FlipperShadows: FlipperShadows=1  '***********set to 1 to turn on Flipper shadows
 
' **********************
 
 
Dim operatormenu, options
Dim balls, peg, pegstart, attract
Dim Add10, Add100, Add1000
Dim hipg,higm, NewHG, PlungeBall
Dim maxplayers, players, player
Dim credit, decks, freeplay, chimesounds
Dim Count(2),dispcount(2), dispcount2(2)
Dim pegs(2)
Dim games(2)
Dim peglights(2,120)
Dim cplay(2)
Dim Pups(2)
Dim state, skillshot, shot, skillcheck
Dim tilt, tiltsens
Dim target(9)
Dim ballinplay, dtstep, chimetime, chimetime1
dim rstep, lstep, rotospin, rotopos, rotocount, spinit
dim dw1step, dw2step, dw3step, dw4step, dw5step, dw6step
Dim rep(2)
Dim rst
Dim eg
Dim bell, eball
Dim i,j, ii, objekt, light, temp
 
On Error Resume Next
ExecuteGlobal GetTextFile("controller.vbs")
If Err Then MsgBox "Can't open controller.vbs"
On Error Goto 0
 
sub nobs_init
    LoadEM
    ScoreSetup
    maxplayers=2
    set Pups(1) = Pup1
    set Pups(2) = Pup2
    set target(1)=DT5
    set target(2)=DT5b
    set target(3)=DT6
    set target(4)=DT4
    set target(5)=DTa
    set target(6)=DT2
    set target(7)=DT3
    set target(8)=DT6A
    set target(9)=DT9
    HSEnterMode = False
    HGEnterMode = False
    hideoptions
    player=1
    skillshot=0
    For each light in cardlights:light.State = 0: Next
    loadhs
    if rotopos="" or rotopos<1 or rotopos>13 then rotopos=1
    if rotospin <1 or rotospin>13 then rotospin=1
    RotoTarget.roty=-1*(rotopos-1)*(360/13)
    if hipg="" then hipg=50
    if higm="" then higm=0
    ii=0
    balls=5
    if HSA1="" then HSA1=68
    if HSA2="" then HSA2=79
    if HSA3="" then HSA3=71
    if HGA1="" then HGA1=68
    if HGA2="" then HGA2=79
    if HGA3="" then HGA3=71
    credittxt.setvalue(0)
    BipReel.setvalue(0)
    if decks="" or decks<1 or decks>4 then decks=4
    if freeplay="" or freeplay<0 or freeplay>1 then freeplay=1
    if freeplay=1 then credit=-1
    if chimesounds="" or chimesounds<0 or chimesounds>1 then chimesounds=1
    OptionDeck.image = "OptionsDecks"&decks
    OptionFree.image="OptionsYN"&freeplay
    OptionChimes.image="OptionsYN"&chimesounds
    for each objekt in plastics: objekt.image = "plastics"&decks: Next
    If B2SOn then
        for each objekt in backdropstuff: objekt.visible = 0 : next
    End If
 
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
 
    for i = 1 to 9
        target(i).isdropped=True
    Next
    tilt=false
    Drain.CreateBall
    startGame.enabled=true
    a1.timerenabled=1
    If nightday<=5 then
        For each light in lights:light.intensityscale = 5:Next
        For each light in InsertLights:light.intensityscale = 3:Next
    End If
    If nightday>5 then
        For each light in lights:light.intensityscale = 2:Next
    End If
    If nightday>10 then
        For each light in lights:light.intensityscale = 1.5:Next
        For each light in InsertLights:light.intensityscale = 1:Next
    End If
    If nightday>30 then
        For each light in lights:light.intensityscale = 1:Next
    End If
    If nightday>80 then
        For each light in lights:light.intensityscale = .8:Next
        For each light in InsertLights:light.intensityscale = .8:Next
    End If
End sub
 
sub attractmode
    if state=false then
        pegani.enabled=1
        LightSeq1.Play SeqUpOn,15,1
        LightSeq1.Play SeqDownOn,15,1
        LightSeq1.Play SeqRightOn,15,1
        LightSeq1.Play SeqLeftOn,15,1
        LightSeq1.Play SeqRandom,5,,2000
    end if
end Sub
 
Sub LightSeq1_PlayDone()
    if state=false then attractmode
End Sub
 
Sub Lgi_timer
    attractmode
    Lgi.timerenabled=0
end sub
 
sub pegani_timer
    if ii>119 Then ii=0
    Peg1(ii).duration 1, 500, 0
    Peg2(ii).duration 1, 500, 0
    ii=ii+1
end sub
 
sub startGame_timer
    playsound "poweron"
    lightdelay.enabled=true
    me.enabled=false
end sub
 
sub lightdelay_timer
    For each light in GIlights:light.state=1:Next
    For each light in BumperLights:light.state=1:Next
    for i = 1 to 9
        DTlights(i-1).state=1
    Next
    credittxt.setvalue(credit+1)
    If credit+1>0 or freeplay=1 then DOF 137, DOFOn
    attract=1
    gamov.text="Game Over"
    if B2SOn then
        if freeplay <> 1 then
            if credit<10 then
                Controller.B2ssetreel 30, Credit
                Controller.B2Ssetreel 29, 0
              else
                Controller.B2ssetreel 30, Credit-10
                Controller.B2Ssetreel 29, 1
            end if
        end if
        Controller.B2SSetGameOver 35,1
    end if
    gamov.timerenabled=1
    Lgi.timerenabled=1
    me.enabled=false
end sub
 
' Display text, style(1-norm, 2-flash), screen(1-player 1, 2-player 2,3-both), justify (0 left, 1-right,2-full), time, textR
 
sub makedispcount(dis,sty)  'used to add spaces to put scoring in correct position, dis=player, sty=style, pegs in right 4 digits, games in left 2
    select case len(games(dis))
        case 1: dispcount(dis)=" "&games(dis)
        case 2: dispcount(dis)=games(dis)
        case 3: dispcount(dis)=99
    end Select
    Display dispcount(dis),sty,dis,2,0,pegs(dis)
end sub
 
sub makedispcount2(dis,sty) 'used to add spaces to put scoring in correct position, dis=player, sty=style, pegs in right 4 digits, games in left 2
    select case len(games(dis))
        case 1: dispcount2(dis)=" "&games(dis)
        case 2: dispcount2(dis)=games(dis)
        case 3: dispcount2(dis)=99
    end Select
    Display2 dispcount2(dis),sty,dis,2,0,pegs(dis)
end sub
 
sub a1_timer
    Select Case (attract)
        Case 1:
            makedispcount 1,1
            makedispcount 2,1
        Case 2:
            Display "HI PEG",1,1,0,0,""
            Display hipg,1,2,2,0,chr(HSA1)&chr(HSA2)&chr(HSA3)
        Case 3:
            Display "HI GMS",1,1,0,0,""
            Display higm,1,2,2,0,chr(HGA1)&chr(HGA2)&chr(HGA3)
    End Select
    attract=attract+1
    if attract>3 then attract=1
end sub
 
 
sub gamov_timer
    if state=false then
        gtimer.enabled=true
    end if
    If B2SOn then Controller.B2SSetGameOver 35,0           
    gamov.text=""
    gamov.timerenabled=0
end sub
 
sub gtimer_timer
    if state=false then
        gamov.text="Game Over"
        If B2SOn then Controller.B2SSetGameOver 35,1
        gamov.timerenabled=1
        gamov.timerinterval= (INT (RND*10)+5)*100
      Else
        If B2SOn then Controller.B2SSetGameOver 35,0
    end if
    me.enabled=0
end sub
 
 
Sub nobs_KeyDown(ByVal keycode)
 
'if keycode=leftmagnasave then
'   chimetime=0
'   ChimeInit.enabled=1
'end if
'
'if keycode=rightmagnasave then
'   chimetime1=0
'   ChimeInit1.enabled=1
'end if
 
    if keycode=AddCreditKey then
        if freeplay=1 or credit=9 then
            PlaySoundAt "coinout", Drain
          else
            PlaySoundAt "coinin", Drain
            addcredit
        end if
    end if
 
    if keycode=StartGameKey and (credit>0 or freeplay=1) and operatormenu=0 And Not (HSEnterMode=true OR HGEnterMode=true) then
      if state=false then
        tilt=false
        state=true
        if freeplay=0 then
            credit=credit-1
            credittxt.setvalue(credit+1)
            if credit+1 < 1 then DOF 137, DOFOff
        end if
        ballinplay=1
        for i = 1 to maxplayers
            games(i)=0
            Count(i)=0
            pegs(i)=0
            rep(i)=0
            pups(i).state=0
        next
        If B2SOn Then
          if freeplay <> 1 then
            if credit<10 then
                Controller.B2ssetreel 30, Credit
                Controller.B2Ssetreel 29, 0
              else
                Controller.B2ssetreel 30, Credit-10
                Controller.B2Ssetreel 29, 1
            end if
          end if
            Controller.B2SSetReel 20, ballinplay
            Controller.B2ssetplayerup 30, 1
            Controller.B2ssetcanplay 31, 1
            Controller.B2SSetGameOver 35, 0
            for i = 1 to 4
                Controller.B2sSetScorePlayer i, 0
            next
        End If
 
        pups(1).state=1
        pup1.timerenabled=1
        tilt=false
        state=true
        if chimesounds=1 Then
            chimetime=0
            ChimeInit.enabled=1
            'playsound "initialize"
        end if
        players=1
        rst=0
        newgame.enabled=true
      else if state=true and players < maxplayers and Ballinplay=1 and pup1.timerenabled=0 then
        credit=credit-1
        players=players+1
        makedispcount2 2,1
        if freeplay=0 then
            credittxt.setvalue(credit+1)
            if credit+1 < 1 then DOF 137, DOFOff
        end if
        If B2SOn then
            if credit<10 then
                Controller.B2ssetreel 30, Credit
                Controller.B2Ssetreel 29, 0
              else
                Controller.B2ssetreel 30, Credit-10
                Controller.B2Ssetreel 29, 1
            end if
            Controller.B2ssetcanplay 31, players
        End If
        if chimesounds=1 Then
            chimetime1=0
            ChimeInit1.enabled=1
        end if          'playsound "init2"
       end if
      end if
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
            if freeplay=0 Then
                freeplay=1
                credit=-1
              Else
                freeplay=0
                credit=0
            end if
            credittxt.setvalue(credit+1)
            OptionFree.image = "OptionsYN"&freeplay
        Case 2:
            if decks=1 then
                decks=2
              else if decks=2 Then
                    decks=3
                else if decks=3 Then
                      decks=4
                    Else
                      decks=1
                    end if
                end If
            end if
            For each objekt in plastics: objekt.image = "plastics"&decks: Next
            OptionDeck.image = "OptionsDecks"&decks
        Case 3:
            if chimesounds=0 Then
                chimesounds=1
              Else
                chimesounds=0
            end If
            OptionChimes.image = "OptionsYN"&chimesounds
        Case 4:
            OperatorMenu=0
            savehs
            HideOptions
      End Select
    End If
 
  If HSEnterMode or HGEnterMode Then HighScoreProcessKey(keycode)
 
  if tilt=false and state=true then
    If keycode = LeftFlipperKey Then
        LeftFlipper.RotateToEnd
        PlaySoundAt SoundFXDOF("flipperup",101,DOFOn,DOFContactors), LeftFlipper
    End If
   
    If keycode = RightFlipperKey Then
        RightFlipper.RotateToEnd
        PlaySoundAt SoundFXDOF("flipperup",102,DOFOn,DOFContactors), RightFlipper
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
 
sub ChimeInit_timer
    Select Case Chimetime
        Case 2:
            PlaySound SoundFXDOF("100a",142, DOFPulse, DOFChimes), 1, chimevol, AudioPan(chimesound), 0,0,0, 1, AudioFade(chimesound)
        Case 40:
            PlaySound SoundFXDOF("100a",142, DOFPulse, DOFChimes), 1, chimevol, AudioPan(chimesound), 0,0,0, 1, AudioFade(chimesound)
        Case 68:
            PlaySound SoundFXDOF("1000a",143, DOFPulse, DOFChimes), 1, chimevol, AudioPan(chimesound), 0,0,0, 1, AudioFade(chimesound)
        Case 81:
            PlaySound SoundFXDOF("10a",141, DOFPulse, DOFChimes), 1, chimevol, AudioPan(chimesound), 0,0,0, 1, AudioFade(chimesound)
        Case 99:
            PlaySound SoundFXDOF("1000a",143, DOFPulse, DOFChimes), 1, chimevol, AudioPan(chimesound), 0,0,0, 1, AudioFade(chimesound)
        Case 113:
            PlaySound SoundFXDOF("10a",141, DOFPulse, DOFChimes), 1, chimevol, AudioPan(chimesound), 0,0,0, 1, AudioFade(chimesound)
        Case 126:
            PlaySound SoundFXDOF("100a",142, DOFPulse, DOFChimes), 1, chimevol, AudioPan(chimesound), 0,0,0, 1, AudioFade(chimesound)
        Case 140:
            PlaySound SoundFXDOF("1000a",143, DOFPulse, DOFChimes), 1, chimevol, AudioPan(chimesound), 0,0,0, 1, AudioFade(chimesound)
        Case 165:
            PlaySound SoundFXDOF("10a",141, DOFPulse, DOFChimes), 1, chimevol, AudioPan(chimesound), 0,0,0, 1, AudioFade(chimesound)
        Case 178:
            PlaySound SoundFXDOF("1000a",143, DOFPulse, DOFChimes), 1, chimevol, AudioPan(chimesound), 0,0,0, 1, AudioFade(chimesound)
        Case 180:
            me.enabled=0
    end Select
    Chimetime=Chimetime+1
End sub
 
sub ChimeInit1_timer
    Select Case Chimetime1
        Case 3:
            PlaySound SoundFXDOF("100a",142, DOFPulse, DOFChimes), 1, chimevol, AudioPan(chimesound), 0,0,0, 1, AudioFade(chimesound)
        Case 16:
            PlaySound SoundFXDOF("1000a",143, DOFPulse, DOFChimes), 1, chimevol, AudioPan(chimesound), 0,0,0, 1, AudioFade(chimesound)
        Case 30:
            PlaySound SoundFXDOF("10a",141, DOFPulse, DOFChimes), 1, chimevol, AudioPan(chimesound), 0,0,0, 1, AudioFade(chimesound)
        Case 44:
            PlaySound SoundFXDOF("100a",142, DOFPulse, DOFChimes), 1, chimevol, AudioPan(chimesound), 0,0,0, 1, AudioFade(chimesound)
        Case 68:
            PlaySound SoundFXDOF("10a",141, DOFPulse, DOFChimes), 1, chimevol, AudioPan(chimesound), 0,0,0, 1, AudioFade(chimesound)
        Case 82:
            PlaySound SoundFXDOF("100a",142, DOFPulse, DOFChimes), 1, chimevol, AudioPan(chimesound), 0,0,0, 1, AudioFade(chimesound)
        Case 100:
            me.enabled=0
    end Select
    Chimetime1=Chimetime1+1
End sub
 
sub pup1_timer
    pup1.timerenabled=0
end sub
 
Sub OperatorMenuTimer_Timer
    OperatorMenu=1
    Displayoptions
    Options=1
End Sub
 
Sub DisplayOptions
    OptionsBack.visible = true
    OptionFree.visible = true
    Option1.visible = True
    OptionDeck.visible = True
    OptionChimes.visible = True
End Sub
 
Sub HideOptions
    for each objekt In OptionMenu
        objekt.visible = false
    next
End Sub
 
 
Sub nobs_KeyUp(ByVal keycode)
 
    If keycode = PlungerKey Then
        Plunger.Fire
        if PlungeBall=1 then
            PlaySoundAt "plunger", Plunger
          else
            PlaySoundAt "plungerreleasefree", Plunger
        end if
    End If
 
    if keycode = LeftFlipperKey then
        OperatorMenuTimer.Enabled = false
    end if
 
   If tilt=false and state=true then
    If keycode = LeftFlipperKey Then
        LeftFlipper.RotateToStart
        PlaySoundAt SoundFXDOF("flipperdown",101,DOFOff,DOFContactors), LeftFlipper
    End If
   
    If keycode = RightFlipperKey Then
        RightFlipper.RotateToStart
        PlaySoundAt SoundFXDOF("flipperdown",102,DOFOff,DOFContactors), RightFlipper
    End If
   End if
End Sub
 
sub flippertimer_timer()
    LFlip.RotY = LeftFlipper.CurrentAngle
    RFlip.RotY = RightFlipper.CurrentAngle
    Pgate.rotz=(Gate.currentangle*.75)+25
    diverter.RotY = DiverterFlipper.CurrentAngle+90
    diverterL.RotY = DiverterFlipperL.CurrentAngle+90  
 
    if FlipperShadows=1 then
        FlipperLSh.RotZ = LeftFlipper.currentangle
        FlipperRSh.RotZ = RightFlipper.currentangle
    end if
 
end sub
 
 
Sub addcredit
    if freeplay=0 then
      credit=credit+1
      DOF 137, DOFOn
      if credit>9 then credit=9
      credittxt.setvalue(credit+1)
      If B2SOn Then
        if credit<10 then
            Controller.B2ssetreel 30, Credit
            Controller.B2Ssetreel 29, 0
          else
            Controller.B2ssetreel 30, Credit-10
            Controller.B2Ssetreel 29, 1
        end if
      end if
    end if
End sub
 
Sub Drain_Hit()
    DOF 134, DOFPulse
    PlaySoundAt "drain", Drain
    b2sflash.uservalue=2        'turn off flash
    B2SFlash2.uservalue=2
    if shootagain.state=0 then for each light in GIlights: light.state=0: next
    me.timerenabled=1
End Sub
 
Sub Drain_timer
    makedispcount player,1
    if shootagain.state=1 and tilt=false then
          if LD1.state+LD2.state+LD3.state+LD4.state=4 then
            for each light in doublelights:light.state=0:next
          end if
          skillcheck=0
          skillshot=1
          plunger.timerenabled=true
        else
          if players=1 or player=players then
            player=1
            pups(players).state=0
            pups(player).state=1
           Else
            player=player+1
            pups(player).state=1
            pups(player-1).state=0
          end if
          If B2SOn then Controller.B2ssetplayerup 30, player
          nextball
    end if
    me.timerenabled=0
End Sub
 
sub ballhome_hit
    plungeball=1
end sub
 
sub ballhome_unhit
    DOF 136, DOFPulse
    plungeball=0
end sub
 
 
' Display text, style(1-norm, 2-flash), screen(1-top, 2-bot,3-both), justify (0 left, 1-right,2-full), time, textR
 
sub newgame_timer
    bumper1.hashitevent=True
    bumper2.hashitevent=True
    RightSlingShot.disabled=False
    player=1
    pegani.enabled=0
    a1.timerenabled=0
    makedispcount 1,1
    Display "",1,2,0,0,""
    LightSeq1.StopPlay()
    for each light in BumperLights:light.state=1:Next
    for each light in peg1:light.state=0:Next
    for each light in peg2:light.state=0:Next
    pups(1).state=1
    If B2SOn then
      Controller.B2SSetGameOver 35,0
      Controller.B2SSetTilt 33,0
    End If
    eg=0
    shootagain.state=lightstateoff
    tilttxt.text=" "
    gamov.text=" "
    ballreltimer.enabled=true
    newgame.enabled=false
end sub
 
sub resetDT_timer
    Select Case DTStep
        Case 1:
            PlaySoundAt SoundFXDOF("bankreset", 116, DOFPulse, DOFContactors), DT4
            for i = 1 to 4
                target(i).isdropped=False
                DTlights(i-1).state=0
            Next
        Case 2:
            PlaySoundAt SoundFXDOF("bankreset", 117, DOFPulse, DOFContactors), DT3
            for i = 5 to 7
                target(i).isdropped=False
                DTlights(i-1).state=0
            Next
        Case 3:
            PlaySoundAt SoundFXDOF("bankreset", 118, DOFPulse, DOFContactors), DT6A
            for i = 8 to 9
                target(i).isdropped=False
                DTlights(i-1).state=0
            Next
        Case 4:
            newball
            resetDT.enabled=False
    end Select
    DTStep=DTStep+1
end sub
 
sub newball
    extraball.state=0
    eball=0
    for each light in cardlights:light.state=1:next
    for each light in doublelights:light.state=0:next
    for each light in handlights:light.state=0:next
    for each light in starlights:light.state=0:next
    for each light in GIlights:light.state=1:next
    skillshot=1
    skillcheck=0
    for each light in skillshotlights:light.state=0:next
    Select Case Int(Rnd*3)+1
        Case 1 : LK.state=2
        Case 2 : LQ.state=2
        Case 3 : LJ.state=2
    End Select
End Sub
 
 
sub nextball
    if tilt=true then
      RightSlingShot.disabled=false
      bumper1.hashitevent=True
      bumper2.hashitevent=True
      tilt=false
      tilttxt.text=" "
        If B2SOn then
            Controller.B2SSetTilt 33,0
            Controller.B2ssetdata 1, 1
        End If
    end if
    if player=1 then ballinplay=ballinplay+1
    if ballinplay>balls then
        eg=1
        ballreltimer.enabled=true
      else
        if state=true then
          ballreltimer.enabled=true
        end if
    end if
End Sub
 
sub ballreltimer_timer
  if eg=1 then
      turnoff
      state=false
      BipReel.setvalue(0)
      gamov.text="GAME OVER"
      gamov.timerenabled=1
      ii=0
      for i=1 to players                       
        pups(i).state=0
      next
        if pegs(1)>hipg or pegs(2)>hipg then
            HStimer.uservalue=0
            HStimer.enabled=1
            if pegs(2)>pegs(1) then
                hipg=pegs(2)
                HighPegEntryInit 2
              else
                hipg=pegs(1)
                HighPegEntryInit 1
            end if
        end if
        if games(1)>higm or games(2)>higm then
            if HStimer.enabled = 0 then
                HStimer.uservalue=0
                HStimer.enabled=1
            end if
            if games(2)>games(1) then
                higm=games(2)
                NewHG=2
              else
                higm=games(1)
                NEWHG=1
            end if
            If HSEnterMode <> True then HighGamesEntryInit NewHG
        end if
        If HSEnterMode = False and HGEnterMode = False then
            Lgi.timerenabled=1
            a1.timerenabled=1
        end if
      for each light in GIlights:light.state=1:next
      savehs
      If B2SOn then
        Controller.B2SSetGameOver 35,1
        Controller.B2SSetReel 20, 0
        Controller.B2ssetcanplay 31, 0
        Controller.B2ssetcanplay 30, 0
      End If
      For each light in LaneLights:light.State = 0: Next
      ballreltimer.enabled=false
    else
      DTstep=1
      resetDT.enabled=true
      plunger.timerenabled=true
      ballreltimer.enabled=false
  end if
end sub
 
Sub HStimer_timer
    playsoundat SoundFXDOF("knock",138,DOFPulse,DOFKnocker), Plunger
    DOF 139,DOFPulse
    HStimer.uservalue=HStimer.uservalue+1
    if hstimer.uservalue=3 then me.enabled=0
end sub
 
Sub plunger_timer
    Drain.kick 60, 11,0
    playsoundat SoundFXDOF("kickerkick",135,DOFPulse,DOFContactors), Drain
    DOF 135, DOFPulse
    BipReel.setvalue(ballinplay+1)
    makedispcount player,2
    if ShootAgain.state=1 then
        extraball.state=0
        eball=0
    end if
    If B2SOn then Controller.B2SSetReel 20, ballinplay
    Plunger.timerenabled=false
end sub
 
 
'********** Bumpers
 
Sub Bumper1_Hit
   if tilt=false then
    PlaySoundAt SoundFXDOF("fx_bumper4",107,DOFPulse,DOFContactors), Bumper1
    DOF 108,DOFPulse
    FlashBumpers
    addscore 100
   end if
End Sub
 
Sub Bumper2_Hit
   if tilt=false then
    PlaySoundAt SoundFXDOF("fx_bumper4",109,DOFPulse,DOFContactors), Bumper2
    DOF 110,DOFPulse
    FlashBumpers
    addscore 100
   end if
End Sub
 
 
sub FlashBumpers
    if bumperlight1.state = 1 then
        for each light in BumperLights: light.duration 0, 200, 1:Next
    end If
end sub
 
 
'************** Slings and animated rubbers
 
Sub RightSlingShot_Slingshot
    PlaySoundAt SoundFXDOF("slingshot",105,DOFPulse,DOFContactors), slingR
    DOF 106,DOFPulse
    addscore 10
    RSling.Visible = 0
    RSling1.Visible = 1
    slingR.objroty = -15   
    RStep = 1
    RightSlingShot.TimerEnabled = 1
End Sub
 
Sub RightSlingShot_Timer
    Select Case RStep
        Case 3:RSLing1.Visible = 0:RSLing2.Visible = 1:slingR.objroty = -7
        Case 4: slingR.objroty = 0:RSLing2.Visible = 0:RSLing.Visible = 1:RightSlingShot.TimerEnabled = 0
    End Select
    RStep = RStep + 1
End Sub
 
sub Dingwalls_hit(idx)
    addscore 10
end sub
 
sub dingwall1_hit
    rdw1.visible=0
    RDW1a.visible=1
    dw1step=1
    Me.timerenabled=1
end sub
 
sub dingwall1_timer
    select case dw1step
        Case 1: RDW1a.visible=0: rdw1.visible=1
        case 2: rdw1.visible=0: rdw1b.visible=1
        Case 3: rdw1b.visible=0: rdw1.visible=1: Me.timerenabled=0
    end Select
    dw1step=dw1step+1
end sub
 
sub dingwall2_hit
    rdw2.visible=0
    RDW2a.visible=1
    dw2step=1
    Me.timerenabled=1
end sub
 
sub dingwall2_timer
    select case dw2step
        Case 1: RDW2a.visible=0: rdw2.visible=1
        case 2: rdw2.visible=0: rdw2b.visible=1
        Case 3: rdw2b.visible=0: rdw2.visible=1: me.timerenabled=0
    end Select
    dw2step=dw2step+1
end sub
 
sub dingwall3_hit
    rdw3.visible=0
    RDW3a.visible=1
    dw3step=1
    Me.timerenabled=1
end sub
 
sub dingwall3_timer
    select case dw3step
        Case 1: RDW3a.visible=0: rdw3.visible=1
        case 2: rdw3.visible=0: rdw3b.visible=1
        Case 3: rdw3b.visible=0: rdw3.visible=1: me.timerenabled=0
    end Select
    dw3step=dw3step+1
end sub
 
sub dingwall4_hit
    lsling.visible=0
    lsling1.visible=1
    dw4step=1
    me.timerenabled=1
end sub
 
sub dingwall4_timer
    select case dw4step
        Case 1: lsling1.visible=0: lsling.visible=1
        case 2: lsling.visible=0: lsling2.visible=1
        Case 3: lsling2.visible=0: lsling.visible=1: me.timerenabled=0
    end Select
    dw4step=dw4step+1
end sub
 
sub dingwall5_hit
    rdw5.visible=0
    RDW5a.visible=1
    dw5step=1
    Me.timerenabled=1
end sub
 
sub dingwall5_timer
    select case dw5step
        Case 1: RDW5a.visible=0: rdw5.visible=1
        case 2: rdw5.visible=0: rdw5b.visible=1
        Case 3: rdw5b.visible=0: rdw5.visible=1: me.timerenabled=0
    end Select
    dw5step=dw5step+1
end sub
 
sub dingwall6_hit
    rdw6.visible=0
    RDW6a.visible=1
    dw6step=1
    Me.timerenabled=1
end sub
 
sub dingwall6_timer
    select case dw6step
        Case 1: RDW6a.visible=0: rdw6.visible=1
        case 2: rdw6.visible=0: rdw6b.visible=1
        Case 3: rdw6b.visible=0: rdw6.visible=1: me.timerenabled=0
    end Select
    dw6step=dw6step+1
end sub
 
'********** Triggers    
 
sub skillshotoff(shot)
    b2sflash.uservalue=2
    makedispcount player,1
    if shootagain.state=1 then
        shootagain.state=0
      else
        DiverterFlipper.RotateToStart
        DiverterFlipperL.RotateToStart
        for each light in skillshotlights
            light.state=1
        Next
    end if
    Select Case shot
        Case 1 : LK.state=0
        Case 2 : LQ.state=0
        Case 3 : LJ.state=0
'       Case 4 :
    End Select
    skillshot=0
end Sub
 
 
sub TGK_hit   '***** top King rollover
    DOF 128, DOFPulse
    LBH4K.state=1
    Lstar1.state=1
    skillcheck=1
    hand4check rotopos, 13
    if LK.state=1 or LK.state=2 then
        if LK.state=2 then addcount 2
        eball=eball+1
        checkeball
        addscore 1000
        LK.state=0
    else
        addscore 100
    end if
    if skillshot=1 then skillshotoff 1
end sub    
 
sub TGQ_hit   '***** top Queen rollover
    DOF 129, DOFPulse
    LBH4Q.state=1
    Lstar2.state=1
    skillcheck=1
    hand4check rotopos, 12
    if LQ.state=1 or LQ.state=2 then
        if LQ.state=2 then addcount 2
        eball=eball+1
        checkeball
        addscore 1000
        LQ.state=0
      else
        addscore 100
    end if
    if skillshot=1 then skillshotoff 2
end sub
 
sub TGJ_hit   '***** top Jack rollover
    DOF 130, DOFPulse
    LBH4J.state=1
    Lstar3.state=1
    skillcheck=1
    hand4check rotopos, 11
    if LJ.state=1 or LJ.state=2 then
        if LJ.state=2 then addcount 2
        eball=eball+1
        checkeball
        addscore 1000
        LJ.state=0
 
      else
        addscore 100
    end if
    if skillshot=1 then skillshotoff 3
end sub
 
sub TG2_hit   '***** right 2 rollover
    DOF 131, DOFPulse
    LBH22B.state=1
    hand2check rotopos, 2
    if (L2.state)=lightstateon then
        eball=eball+1
        checkeball
        addscore 1000
        L2.state=0
    else
        addscore 100
    end if
end sub
 
sub TG7_hit   '***** left 7 rollover
    DOF 121, DOFPulse
    LBh37.state=1
    hand3check rotopos, 7
    if (L7.state)=lightstateon then
        eball=eball+1
        checkeball
        addscore 1000
        L7.state=0
        Lstar4.state=1
    else
        addscore 100
    end if
end sub
 
sub TG8_hit   '***** right 8 rollover
    DOF 122, DOFPulse
    LBh38.state=1
    hand3check rotopos, 8
    if (L8.state)=lightstateon then
        eball=eball+1
        checkeball
        addscore 1000
        L8.state=0
        Lstar5.state=1
    else
        addscore 100
    end if
end sub
 
sub checkeball
    if eball>5 and ShootAgain.state=0 Then
        extraball.state=1
        DiverterFlipper.timerenabled=1
    end if
end Sub
 
sub DiverterFlipper_timer
    DiverterFlipper.RotateToEnd
    DiverterFlipper.timerenabled=0
end Sub
 
sub TG5_hit   '***** left 5 rollover
    DOF 120, DOFPulse
    hand4check rotopos, 5
    if (L5.state)=lightstateon then
        addscore 1000
        L5.state=0
    else
        addscore 100
    end if
    LBH45.state=1
    DiverterFlipperL.timerenabled=1
end sub
 
sub DiverterFlipperL_timer
    DiverterFlipperL.RotateToEnd
    DiverterFlipperL.timerenabled=0
end Sub
 
sub TGstar1_hit   '*****star rollover
    DOF 123, DOFPulse
    if (Lstar1.state)=lightstateon then
        addscore 1000
        addcount(1)
    else
        addscore 100
    end if
end Sub
 
sub TGstar2_hit   '*****star rollover
    DOF 124, DOFPulse
    if (Lstar2.state)=lightstateon then
        addscore 1000
        addcount(1)
    else
        addscore 100
    end if
end Sub
 
sub TGstar3_hit   '*****star rollover
    DOF 125, DOFPulse
    if (Lstar3.state)=lightstateon then
        addscore 1000
        addcount(1)
    else
        addscore 100
    end if
end Sub
 
sub TGstar4_hit   '*****star rollover
    DOF 126, DOFPulse
    if (Lstar4.state)=lightstateon then
        addscore 1000
        addcount(1)
    else
        addscore 100
    end if
end Sub
 
sub TGstar5_hit   '*****star rollover
    DOF 127, DOFPulse
    if (Lstar5.state)=lightstateon then
        addscore 1000
        addcount(1)
    else
        addscore 100
    end if
end Sub
 
sub hand1check(pos1, hit1)
    dim hand1, check1
    if pos1>9 then
        check1=10
      else
        check1=pos1
    end if
    hand1=LBH14.state+LBH15b.state+LBH15.state+LBH16.state
    if hand1=4 then
        select case(pos1)
            case 1,3,7,9:
                addcount(14)
            case 2,8:
                addcount(12)
            case 4, 6:
                addcount(24)
            case 5:
                addcount(21)
            case 10,11,12,13:
                addcount(16)
        end Select
        for each light in hand1lights:light.duration 2, 3000, 1:Next
        for each light in hand1Blights:light.duration 2, 3000, 0:Next
        LD1.state=1
        if DT5.timerenabled=false then dt5.timerenabled=True
      else
        IF pos1=hit1 or hit1+check1=15 then addcount 2
    end If
end Sub
 
sub DT5_timer
    for i= 1 to 4:target(i).isdropped=0:Next
    PlaySoundAt SoundFXDOF("bankreset", 116, DOFPulse, DOFContactors), DT4
    for each light in hand1dtlights:light.state=0:Next
    dt5.timerenabled=0
end sub
 
sub hand2check(pos2, hit2)
    dim hand2, check2
    if pos2>9 then
        check2=10
      else
        check2=pos2
    end if
    hand2=LBH23.state+LBH22.state+LBH22B.state+LBH2A.state
    if hand2=4 then
        select case(pos2)
            case 1,3
                addcount(16)
            case 9
                addcount(12)
            case 10,11,12,13:
                addcount(14)
            case 2:
                addcount(15)
            case 4,7,8:
                addcount(10)
            case 5,6:
                addcount(8)
        end Select
        for each light in hand2lights:light.duration 2, 3000, 1:Next
        for each light in hand2Blights:light.duration 2, 3000, 0:Next
        ld2.state=1
        if DTa.timerenabled=false then DTa.timerenabled=True
      else
        IF pos2=hit2 or hit2+check2=15 then addcount 2
    end If
end Sub
 
sub DTa_timer
    for i= 5 to 7:target(i).isdropped=0:Next
    PlaySoundAt SoundFXDOF("bankreset", 117, DOFPulse, DOFContactors), DT3
    for each light in hand2dtlights:light.state=0:Next
    DTa.timerenabled=0
end sub
 
 
sub hand3check(pos3, hit3)
    dim hand3, check3
    if pos3>9 then
        check3=10
      else
        check3=pos3
    end if
    hand3=LBH39.state+LBH38.state+LBH37.state+LbH36.state
    if hand3=4 Then
        select case(pos3)
            case 1, 2:
                addcount(10)
            case 3,4,11,12,13:
                addcount(8)
            case 5,10:
                addcount(9)
            case 6,7,8,9:
                addcount(16)
        end Select
        for each light in hand3lights:light.duration 2, 3000, 1:Next
        for each light in hand3Blights:light.duration 2, 3000, 0:Next
        LD3.state=1
        if DT6a.timerenabled=false then DT6a.timerenabled=True
      else
        IF pos3=hit3 or hit3+check3=15 then addcount 2
    end if
end Sub
 
sub dt6a_timer
    for i= 8 to 9:target(i).isdropped=0:Next
    PlaySoundAt SoundFXDOF("bankreset", 118, DOFPulse, DOFContactors), DT6A
    for each light in hand3DTlights:light.state=0:Next
    DT6a.timerenabled=0
end Sub
 
sub hand4check(pos4, hit4)
    dim hand4, check4, check4b
    if pos4>9 then
        check4=10
      else
        check4=pos4
    end if
    if hit4>9 then
        check4b=10
      else
        check4b=hit4
    end if
    hand4=LBH4K.state+LBH4Q.state+LBH4J.state+LBH45.state
    if hand4=4 Then
        select Case(pos4)
            case 1,2,3,4,6,7,8,9:
                addcount(9)
            case 5:
                addcount(17)
            case 10:
                addcount(12)
    :       case 11,12,13:
                addcount(16)
        end Select
        for each light in hand4lights:light.duration 2, 3000, 1:Next
        for each light in hand4Blights:light.duration 2, 3000, 0:Next
        LD4.state=1
        if LK.timerenabled=false then LK.timerenabled=True
      else
        IF pos4=hit4 or check4b+check4=15 then addcount 2
    end if
end Sub
 
sub lk_timer
    If (LD1.state+LD2.state+LD3.state+LD4.state) <> 4 then DiverterFlipperL.RotateToStart
    LK.timerenabled=0
end Sub
 
'********** Drop Targets
 
sub DropTargets_dropped (idx)
    PlaySoundAtBall "drop1"
end Sub
 
sub DTa_dropped
    DOF 112, DOFPulse
    if state=true then
        addscore 1000
        ldta.state=1
        lbh2a.state=1
        hand2check rotopos, 1
    end If
end sub
 
sub DT2_dropped
    DOF 112, DOFPulse
    if state=true then
        addscore 1000
        Ldt2.state=1
        LBh22.state=1
        hand2check rotopos, 2
    end if
end sub
 
sub DT3_dropped
    DOF 112, DOFPulse
    if state=true then
        addscore 1000
        Ldt3.state=1
        Lbh23.state=1
        hand2check rotopos, 3
    end if
end sub
 
sub DT4_dropped
    DOF 111, DOFPulse
    if state=true then
        addscore 1000
        ldt4.state=1
        Lbh14.state=1
        hand1check rotopos, 4
    end if
end sub
 
sub DT5_dropped
    DOF 111, DOFPulse
    if state=true then
        addscore 1000
        ldt5.state=1
        LBH15.state=1
        hand1check rotopos, 5
    end if
end sub
 
sub DT5b_dropped
    DOF 111, DOFPulse
    if state=true then
        addscore 1000
        ldt5b.state=1
        LBH15B.state=1
        hand1check rotopos, 5
    end if
end sub
 
sub DT6_dropped
    DOF 111, DOFPulse
    if state=true then
        addscore 1000
        Ldt6.state=1
        Lbh16.state=1
        hand1check rotopos, 6
    end If
end sub
 
sub DT6A_dropped
    DOF 112, DOFPulse
    if state=true then
        addscore 1000
        ldt6a.state=1
        LBh36.state=1
        hand3check rotopos, 6
    end If
end sub
 
sub DT9_dropped
    DOF 112, DOFPulse
    if state=true then
        addscore 1000
        Ldt9.state=1
        LBH39.state=1
        hand3check rotopos, 9
    end if
end sub
 
'********** Targets
 
sub Turn_hit()
    if rotopos>9 Then
        addcount(10)
      Else
        addcount(rotopos)
    end If
    if Turn.timerenabled=false Then Turn.timerenabled=True
end Sub
 
sub TSpin_hit()
    addcount(1)
    if Extraball.state=1 then
        playsound "gong"
        ShootAgain.state=1
        extraball.state=0
    end If
    if Turn.timerenabled=false Then turn.timerenabled=True
end Sub
 
sub Turn_timer()
    if moveroto.timerenabled=false then
        if rotopos>rotospin then
            Rotopos=rotopos-rotospin
          Else
            Rotopos=(13+rotopos)-rotospin
        end if
        if rotopos=11 then addcount(2)
        rotocount=0
        spinit=rotospin
        rotospin=rotospin+1
        if rotospin>13 then rotospin=1
        DOF 115, DOFOn
        PlaySoundAt SoundFX("motor",DOFGear), RotoTarget
        moveroto.timerenabled=1
        Turn.timerenabled=false
      Else
        Turn.timerenabled=false
    end if
end Sub
 
sub moveroto_timer
    RotoLight.state=2
    Lgi8.state=0
    rotocount=rotocount+1
    if rotocount=(spinit*4)+53 then
        rotolight.state=0
        lgi8.state=1
        moveroto.timerenabled=False
        DOF 115, DOFOff
        StopSound "motor"
      Else
        RotoTarget.roty=RotoTarget.roty+(360/52)
    end if
end sub
 
 
sub addcount(peg)
    If ld1.state+ld2.state+ld3.state+ld4.state=4 then peg=peg*2
    PlaySoundAt "woodpeg", Plunger
    pegstart=Count(player)
    Count(player)=Count(player)+peg
    pegs(player)=pegs(player)+peg
    if Count(player)>120 Then
        playsoundat soundFXDOF("gong", 144, DOFPulse, DOFBell), TG5
        ShootAgain.state=1
        games(player)=games(player)+1
 
        if players=2 then                  'check for skunks 2 player game only
            if player=1 and count(2)<60 Then
                HStimer.uservalue=1
                HStimer.enabled=1
                games(1)=games(1)+3
              else if player=1 and count(2)<90 Then
                PlaySoundAt SoundFXDOF("knock",138,DOFPulse,DOFKnocker), Plunger
                DOF 139, DOFPulse
                games(1)=games(1)+1
              end If
            end If
            if player=2 and Count(1)<60 Then
                HStimer.uservalue=1
                HStimer.enabled=1
                games(2)=games(2)+3
              else if player=2 and count(1)<90 Then
                PlaySoundAt SoundFXDOF("knock",138,DOFPulse,DOFKnocker), Plunger
                DOF 139, DOFPulse
                games(2)=games(2)+1
              end If
            end If
        end if
 
        for each light in Peg1:light.state=0: Next
        for each light in Peg2:light.state=0:Next
        Count(1)=0
        Count(2)=0
      Else
        for i = pegstart+1 to Count(player)
            if player=1 then Peg1(i-1).duration 2, 3000, 1
            if player=2 then Peg2(i-1).duration 2, 3000, 1
        Next
    end If
 
    makedispcount player,1
end Sub
 
sub addscore(points)
  if skillcheck=0 then
    skillcheck=1
    skillshotoff 4
  end if
  if tilt=false then
    If Points < 100 and AddScore10Timer.enabled = false Then
        Add10 = Points \ 10
        rotospin=rotospin+1
        if rotospin>13 then rotospin=1
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
 
Sub AddPoints(Points)
    if chimesounds=1 Then           'if chimesounds on then play sounds
        ' Sounds: there are 3 sounds: tens, hundreds and thousands
        If points = 1000 Then
            PlaySound SoundFXDOF("bell1000",143, DOFPulse, DOFChimes), 1, chimevol, AudioPan(chimesound), 0,0,0, 1, AudioFade(chimesound)
          elseif Points = 100 Then
            PlaySound SoundFXDOF("bell100",142, DOFPulse, DOFChimes), 1, chimevol, AudioPan(chimesound), 0,0,0, 1, AudioFade(chimesound)
 
          Else
            PlaySound SoundFXDOF("bell10",141, DOFPulse, DOFChimes), 1, chimevol, AudioPan(chimesound), 0,0,0, 1, AudioFade(chimesound)
         End If
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
    PlaySoundAt "tilt", Plunger
    turnoff
End Sub
 
sub turnoff
    RightSlingShot.disabled=true
    bumper1.hashitevent=False
    bumper2.hashitevent=False
    LeftFlipper.RotateToStart
    RightFlipper.RotateToStart
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
 
 
'*********************************************************************
'                     Supporting Ball & Sound Functions
'*********************************************************************
 
Function AudioFade(tableobj) ' Fades between front and back of the table (for surround systems or 2x2 speakers, etc), depending on the Y position on the table. "table1" is the name of the table
    Dim tmp
    tmp = tableobj.y * 2 / Nobs.height-1
    If tmp > 0 Then
        AudioFade = Csng(tmp ^10)
    Else
        AudioFade = Csng(-((- tmp) ^10) )
    End If
End Function
 
Function AudioPan(tableobj) ' Calculates the pan for a tableobj based on the X position on the table. "table1" is the name of the table
    Dim tmp
    tmp = tableobj.x * 2 / Nobs.width-1
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
        If BallVel(BOT(b) ) > 1 AND BOT(b).z < 30 Then
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
'           BALL SHADOW
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
        If BOT(b).X < Nobs.Width/2 Then
            BallShadow(b).X = ((BOT(b).X) - (Ballsize/16) + ((BOT(b).X - (Nobs.Width/2))/17))' + 13
        Else
            BallShadow(b).X = ((BOT(b).X) + (Ballsize/16) + ((BOT(b).X - (Nobs.Width/2))/17))' - 13
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
    savevalue "nobs", "credit", credit
    savevalue "nobs", "hipeg", hipg
    savevalue "nobs", "higame", higm
    savevalue "nobs", "count1", pegs(1)
    savevalue "nobs", "count2", pegs(2)
    savevalue "nobs", "games1", games(1)
    savevalue "nobs", "games2", games(2)
    savevalue "nobs", "decks", decks
    savevalue "nobs", "rotopos", rotopos
    savevalue "nobs", "freeplay", freeplay
    savevalue "nobs", "chimesounds", chimesounds
    savevalue "nobs", "hsa1", HSA1
    savevalue "nobs", "hsa2", HSA2
    savevalue "nobs", "hsa3", HSA3
    savevalue "nobs", "hga1", HGA1
    savevalue "nobs", "hga2", HGA2
    savevalue "nobs", "hga3", HGA3
end sub
 
sub loadhs
    temp = LoadValue("nobs", "credit")
    If (temp <> "") then credit = CDbl(temp)
    temp = LoadValue("nobs", "hipeg")
    If (temp <> "") then hipg = CDbl(temp)
    temp = LoadValue("nobs", "higame")
    If (temp <> "") then higm = CDbl(temp)
    temp = LoadValue("nobs", "count1")
    If (temp <> "") then pegs(1) = CDbl(temp)
    temp = LoadValue("nobs", "count2")
    If (temp <> "") then pegs(2) = CDbl(temp)
    temp = LoadValue("nobs", "games1")
    If (temp <> "") then games(1) = CDbl(temp)
    temp = LoadValue("nobs", "games2")
    If (temp <> "") then games(2) = CDbl(temp)
    temp = LoadValue("nobs", "decks")
    If (temp <> "") then decks = CDbl(temp)
    temp = LoadValue("nobs", "rotopos")
    If (temp <> "") then rotopos = CDbl(temp)
    temp = LoadValue("nobs", "freeplay")
    If (temp <> "") then freeplay = CDbl(temp)
    temp = LoadValue("nobs", "chimesounds")
    If (temp <> "") then chimesounds = CDbl(temp)
    temp = LoadValue("nobs", "hsa1")
    If (temp <> "") then HSA1 = CDbl(temp)
    temp = LoadValue("nobs", "hsa2")
    If (temp <> "") then HSA2 = CDbl(temp)
    temp = LoadValue("nobs", "hsa3")
    If (temp <> "") then HSA3 = CDbl(temp)
    temp = LoadValue("nobs", "hga1")
    If (temp <> "") then HGA1 = CDbl(temp)
    temp = LoadValue("nobs", "hga2")
    If (temp <> "") then HGA2 = CDbl(temp)
    temp = LoadValue("nobs", "hga3")
    If (temp <> "") then HGA3 = CDbl(temp)
end sub
 
Sub nobs_Exit()
    Savehs
    turnoff
    If B2SOn Then Controller.stop
End Sub
 
'==========================================================================================================================================
'============================================================= START OF 6-DIGIT GOTTLIEB DISPLAY =============================================================
'==========================================================================================================================================
'
'  Segmented score display originally developed by Wizards Hat, modified to Gottlieb 6 digit display and simplified by BorgDog
'
'
'Command as follows:
'1) Display             = simply displays a message
 
'Each command is then followed by text, style, screen, full, time, text2   - where text is a string. Style, screen, full & time are numbers
'   - style: 1-2               1=Still text
'                              2=Flashing Text
'                              
'   - screen: 1-3              1=Player 1 display
'                              2=Player 2 display
'                              3=both
 
'   - full: 0 or 1             0=Left justify
'                              1=Right Justify
'                              2=Full Justify (text to left, text2 to right)
'
'   - time: 0+                 Time in seconds for display to show before screen (1-3) is cleared to blank.  O=permanent (until next Display call)
'
 
'As a more advanced example: To display the same message but have "WIZARDS" scroll in from the left then stop, and "HAT" scroll in from the right and then stop and flash
'You would use the following 2 commands:
'       Display "  WIZARDS ",6,1,0,5
'       Display2 "    HAT   ",5,2,0,5
'
'
'the allowed characters are A-Z 0-9 and + - * " ' ( ) / < = > [ \ ] _ `
'lower case is converted to upper case
'invalid characters are converted to spaces !@#$%^&{}|;:?., - in fact any ASCII character not listed in allowed characters
'
'
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'++++++++++++++++++++++++ You do not need to change or understand anything below this line in order to now be able to use the display system +++++++++++++++++++++++++++++++++++++
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
 
 'All variables have had "HAT" prefixed to them to ensure that they are unique if this part of the script is copied directly to another table - it looks confusing to me, but at least
 'it should avoid any duplication of variable names! (thanks to Bob5453 for this suggestion)
 
 'Variables used to count, or keep track
Dim HATn,HATi,HATj,HATV,HATW,HATA,HATB,HATC,HATD,HATE,HATF,HATG,HATV2,HATW2,HATA2,HATB2,HATC2,HATD2,HATE2,HATF2,HATG2
'Variables used to track pre-stated values
Dim HATCurrent,HATPosition,HATCharacter,HATText,HATStyle,HATScreen,HATFull,HATTime,HATtextR
Dim HATCurrent2,HATPosition2,HATCharacter2,HATText2,HATStyle2,HATScreen2,HATFull2,HATTime2,HATtextR2
'Arrays and flags
 
Dim HATGasSeg(12,8),HATRom(255)    ' HATGasSeq is 12 digits, 8 segments per digit, HATRom is ascii table
 
'Initilise Settings For The Display - this is done using the initialization of the first light in the display
Sub ScoreSetup()
 
    For HATA=1 To 12                                    'Set the lights to an array
        For HATB=1 To 8
            Set HATGasSeg(HATA,HATB)=Eval("a"& HATB+((HATA-1)*8))
        Next
    Next
    For HATA=1 To 96                                    'Set all display lights to off (again!)
        Eval("a"& HATA).State=0
    Next
 
'the allowed characters are A-Z 0-9 and + - * " ' ( ) / < = > [ \ ] _ `
 
    For HATA=0 To 255:HATRom(HATA)="00000000":Next      'Set all the possbile characters as spaces (blank)
    HATRom(32)="00000000":HATRom(34)="01000100":HATRom(39)="01000000"       'Settings for each character to be displayed
    HATRom(40)="01100000":HATRom(41)="00001100":HATRom(42)="00000011"
    HATRom(43)="00000011":HATRom(44)="00001000":HATRom(45)="00000010"
    HATRom(47)="01001010"
    HATRom(48)="11111100":HATRom(49)="00000001":HATRom(50)="11011010"
    HATRom(51)="11110010":HATRom(52)="01100110":HATRom(53)="10110110"
    HATRom(54)="10111110":HATRom(55)="11100000":HATRom(56)="11111110"
    HATRom(57)="11110110":HATRom(60)="01100010":HATRom(61)="00010010"
    HATRom(62)="00001110":HATRom(65)="11101110":HATRom(66)="11110011"
    HATRom(67)="10011100":HATRom(68)="11110001":HATRom(69)="10011110"
    HATRom(70)="10001110":HATRom(71)="10111100":HATRom(72)="01101110"
    HATRom(73)="00000001":HATRom(74)="01111000":HATRom(75)="00001111"
    HATRom(76)="00011100":HATRom(77)="11101101":HATRom(78)="11101100"
    HATRom(79)="11111100":HATRom(80)="11001110":HATRom(81)="11100110"
    HATRom(82)="11001100":HATRom(83)="10110110":HATRom(84)="10000001"
    HATRom(85)="01111100":HATRom(86)="01010110":HATRom(87)="01111101"
    HATRom(88)="00100111":HATRom(89)="01100110":HATRom(90)="11011010"
    HATRom(91)="01100000":HATRom(92)="00100110":HATRom(93)="00001100"
    HATRom(95)="00010000":HATRom(96)="00000100"
 
    HATRom(163)="11111111"  'Test all on (ASCII 163 = £)
 
End Sub
 
''''''
' B2S ******************************************** Adapted from JPs ScarFace for 10 segment display
Sub B2SDisplayChar(achar, adigit)
    dim ledvalue, ascchar
    ledvalue = 0
    If achar = "" Then achar = " "
    ascchar = ASC(achar)
    Select Case ascchar
        Case 34: ledvalue = 2+32                'double quote
        Case 39: ledvalue = 2                   'single quote
        Case 40: ledvalue = 2+4                 '(
        Case 41: ledvalue = 16+32               ')
        Case 42: ledvalue = 64+256+512          '*
 
        Case 43: ledvalue = 64+256+512          '+
        Case 44: ledvalue = 16                  ',
        Case 45: ledvalue = 8                   '-
        Case 47: ledvalue = 2+16+64             '/
 
        Case 48: ledvalue = 1+2+4+8+16+32       '0
        Case 49: ledvalue = 256+512             '1
        Case 50: ledvalue = 1+2+8+16+64         '2
        Case 51: ledvalue = 1+2+4+8+64          '3
        Case 52: ledvalue = 2+4+32+64           '4
 
        Case 53: ledvalue = 1+4+8+32+64             '5
        Case 54: ledvalue = 1+4+8+16+32+64          '6
        Case 55: ledvalue = 1+2+4                   '7
        Case 56: ledvalue = 1+2+4+8+16+32+64        '8
        Case 57: ledvalue = 1+2+4+8+32+64           '9
 
        Case 60: ledvalue = 2+4+64                  '<
        Case 61: ledvalue = 8+64                    '=
        Case 62: ledvalue = 16+32+64                '>
 
        Case 65: ledvalue = 1+2+4+16+32+64          'A
        Case 66: ledvalue = 1+2+4+8+64+256+512      'B
        Case 67: ledvalue = 1+8+16+32               'C
        Case 68: ledvalue = 1+2+4+8+256+512         'D
        Case 69: ledvalue = 1+8+16+32+64            'E
 
        Case 70: ledvalue = 1+16+32+64              'F
        Case 71: ledvalue = 1+4+8+16+32             'G
        Case 72: ledvalue = 2+4+16+32+64            'H
        Case 73: ledvalue = 1+8+256+512             'I
        Case 74: ledvalue = 2+4+8+16                'J
 
        Case 75: ledvalue = 16+32+64+256+512        'K
        Case 76: ledvalue = 8+16+32                 'L
        Case 77: ledvalue = 1+2+4+16+32+256+512     'M
        Case 78: ledvalue = 1+2+4+16+32             'N
        Case 79: ledvalue = 1+2+4+8+16+32           'O
 
        Case 80: ledvalue = 1+2+16+32+64            'P
        Case 81: ledvalue = 1+2+4+32+64             'Q
        Case 82: ledvalue = 1+2+16+32               'R
        Case 83: ledvalue = 1+4+8+32+64             'S
        Case 84: ledvalue = 1+256+512               'T
 
        Case 85: ledvalue = 2+4+8+16+32             'U
        Case 86: ledvalue = 2+8+32+64               'V
        Case 87: ledvalue = 2+4+8+16+32+256+512     'W
        Case 88: ledvalue = 4+32+64+256+512         'X
        Case 89: ledvalue = 2+4+32+64               'Y
 
        Case 90: ledvalue = 1+2+8+16+64             'Z
        Case 91: ledvalue = 2+4             '[
        Case 92: ledvalue = 4+32+64         '\
        Case 93: ledvalue = 16+32           ']
        Case 95: ledvalue = 8               '_
        Case 96: ledvalue = 32              ' 'accent
 
        Case Else: ledvalue = 0
    End Select
    Controller.B2SSetLED adigit, ledvalue
End Sub
 
 
Dim DOGtx, DOGst, DOGsc
 
Sub DisplayOne(tx,st,sc)    'this routine will hopefully change just one character in the displays.tx is one character to display, st is style, sc is position
    DOGtx=tx: DOGst=st: DOGsc=sc    'set to global variables
    If b2son then
        if st=2 then
            B2SDisplayChar DOGtx, DOGsc
            FlashOne.uservalue=1
            Flashone.enabled=1
          else
            B2SDisplayChar DOGtx, DOGsc
        end if
      else
        For i=1 to 8
            if st=2 then
                HATGasSeg(sc,i).duration (Mid(HATRom(Asc(tx)),i,1)),300,(Mid(HATRom(Asc(tx)),i,1)*st)
              Else
                HATGasSeg(sc,i).State=(Mid(HATRom(Asc(tx)),i,1)*st)
            end If
        Next
    end if
End Sub
 
Sub FlashOne_timer  ' flash single character on the B2S
 
    Select Case Flashone.uservalue
        Case 0:
            B2SDisplayChar "", DOGsc
            if Flashone.uservalue<>2 then Flashone.uservalue=1
        Case 1:
            B2SDisplayChar DOGtx, DOGsc
            if Flashone.uservalue<>2 then Flashone.uservalue=0
        Case 2:
            B2SDisplayChar DOGtx, DOGsc
            Flashone.enabled=0
    end Select
End sub
 
Sub Display(tx,st,sc,fl,tm,txr)                         'This routine is followed whenever a message is to be displayed using the command Display
    On error resume next
    HATtext=tx:HATstyle=st:HATscreen=sc:HATfull=fl:HATTime=tm:HATtextR=txr      'Set global variables from numbers passed to this routine
 
'step 1 - clear screen(s)
    Clear
 
'step 2 - Add space if needed to right or full justify (1=right justify, 0=left justify, 2=full justify tx and txr)
    If HATFull=1 and (HATStyle=1 or HATStyle=2) then
        If HATScreen=3 then :FullLengthText HATText:Else:HalfLengthText HATText:End If
    End If
 
    If HATFull=2 and (HATStyle=1 or HATStyle=2) then
        If HATScreen=3 then :FullJustify HATText,HATtextR:Else:FullJustify HATText,HATtextR:End If
    End If
 
'step 3 - set variables for 1st & last position
    If HATScreen=3 or HATScreen=1 then HATC=1:HATE=0 Else HATC=7:HATE=6
    If HATScreen=1 then HATD=6 Else HATD=12
 
'step 4 - put display into action
    If HATStyle=1 then  'Still
        For HATPosition=HATC to HATD
            If HATPosition-HATE<=Len(HATText) then
                HATCharacter=Ucase(Mid(HATText,HATPosition-HATE,1))
                if b2son then
                    B2SDisplayChar HATCharacter, HATPosition
                  else
                    For HATA=1 to 8
                        HATGasSeg(HATPosition,HATA).State=(Mid(HATRom(Asc(HATCharacter)),HATA,1)*HATStyle)
                    Next
                end if
            End If
        Next
    End If
 
    If HATStyle=2 then  'Flashing
        if b2son then
            B2SFlash.uservalue = 1
            B2SFlash.enabled= true  
          else
            For HATPosition=HATC to HATD
                If HATPosition-HATE<=Len(HATText) then
                    HATCharacter=Ucase(Mid(HATText,HATPosition-HATE,1))
                    For HATA=1 to 8
                            HATGasSeg(HATPosition,HATA).State=(Mid(HATRom(Asc(HATCharacter)),HATA,1)*HATStyle)
                    Next
                End If
            Next
        end if
    End If
 
End Sub
 
Sub FullJustify(Tx,TxR)                         'Used to right justify for style 1,2
    Select Case 6-Len(Tx)-Len(TxR)
        Case -1:Tx="999"&TxR
        Case 0:Tx=Tx&TxR
        Case 1:Tx=Tx&" "&TxR
        Case 2:Tx=Tx&"  "&TxR
        Case 3:Tx=Tx&"   "&TxR
        Case 4:Tx=Tx&"    "&TxR
        Case 5:Tx=Tx&"     "&TxR
        Case 6:Tx=Tx&"      "&TxR
    End Select
End Sub
 
Sub FullLengthText(Tx)                              'Used to right justify for style 1,2
    Select Case Len(Tx)
        Case 1:Tx="           "&Tx
        Case 2:Tx="          "&Tx
        Case 3:Tx="         "&Tx
        Case 4:Tx="        "&Tx
        Case 5:Tx="       "&Tx
        Case 6:Tx="      "&Tx
        Case 7:Tx="     "&Tx
        Case 8:Tx="    "&Tx
        Case 9:Tx="   "&Tx
        Case 10:Tx="  "&Tx
        Case 11:Tx=" "&Tx
    End Select
End Sub
 
 
Sub HalfLengthText(Tx)                              'Used to right justify for style 1,2
    Select Case Len(Tx)
        Case 1:Tx="     "&Tx
        Case 2:Tx="    "&Tx
        Case 3:Tx="   "&Tx
        Case 4:Tx="  "&Tx
        Case 5:Tx=" "&Tx
    End Select
End Sub
 
 
Sub Clear                                           'Used to clear screen (top, bottom, or both)
    dim dogb, dogg
    if b2son then
        If HATScreen=1 then dogb=6 Else dogb=12
        If HATScreen=2 then dogg=7 Else dogg=1
        For HATA=dogg to dogb
            B2SDisplayChar "", HATA
        Next
      else
        If HATScreen=1 then HATB=48 Else HATB=96
        If HATScreen=2 then HATG=49 Else HATG=1
        For HATA=HATG To HATB
            Eval("a"& HATA).State=0
        Next
    end if
End Sub
 
sub B2SFlash_timer
    Select Case b2sflash.uservalue
        Case 0:
            For HATPosition=HATC to HATD
                If HATPosition-HATE<=Len(HATText) then
                    HATCharacter=Ucase(Mid(HATText,HATPosition-HATE,1))
                    B2SDisplayChar "", HATPosition
                End If
            Next
            if b2sflash.uservalue<>2 then B2SFlash.uservalue=1
        Case 1:
            For HATPosition=HATC to HATD
                If HATPosition-HATE<=Len(HATText) then
                    HATCharacter=Ucase(Mid(HATText,HATPosition-HATE,1))
                    B2SDisplayChar HATCharacter, HATPosition
                End If
            Next
            if b2sflash.uservalue<>2 then B2SFlash.uservalue=0
        Case 2:                                                                                         'turn b2s display to on and turn off flashing
            For HATPosition=HATC to HATD
                If HATPosition-HATE<=Len(HATText) then
                    HATCharacter=Ucase(Mid(HATText,HATPosition-HATE,1))
                    B2SDisplayChar HATCharacter, HATPosition
                End If
            Next
            b2sflash.enabled=0
    end Select
end Sub
 
'------------------------------------
'Basically all the same again, so that the 2 halves of the screen can be treated separately simultaneously
'------------------------------------
Sub Display2(tx,st,sc,fl,tm,txr)                            'This routine is followed whenever a message is to be displayed using the command Display
    On error resume next
    HATtext2=tx:HATstyle2=st:HATscreen2=sc:HATfull2=fl:HATTime2=tm:HATtextR2=txr        'Set global variables from numbers passed to this routine
 
'step 1 - clear screen(s)
    Clear2
 
'step 2 - Add space if needed to right or full justify (1=right justify, 0=left justify, 2=full justify)
    If HATFull2=1 and (HATStyle2=1 or HATStyle2=2) then
        If HATScreen2=3 then :FullLengthText2 HATText2:Else:HalfLengthText2 HATText2:End If
    End If
 
    If HATFull2=2 and (HATStyle2=1 or HATStyle2=2) then
        If HATScreen2=3 then :FullJustify2 HATText2,HATtextR2:Else:FullJustify2 HATText2,HATtextR2:End If
    End If
 
 
'step 3 - set variables for 1st & last position
    If HATScreen2=3 or HATScreen2=1 then HATC2=1:HATE2=0 Else HATC2=7:HATE2=6
    If HATScreen2=1 then HATD2=6 Else HATD2=12
 
'step 4 - put display into action
    If HATStyle2=1 then 'Still
        For HATPosition2=HATC2 to HATD2
            If HATPosition2-HATE2<=Len(HATText2) then
                HATCharacter2=Ucase(Mid(HATText2,HATPosition2-HATE2,1))
                if b2son then
                    B2SDisplayChar HATCharacter2, HATPosition2
                  else
                    For HATA2=1 to 8
                        HATGasSeg(HATPosition2,HATA2).State=(Mid(HATRom(Asc(HATCharacter2)),HATA,1)*HATStyle2)
                    Next
                end if
            End If
        Next
    End If
 
    If HATStyle=2 then  'Flashing
        if b2son then
            B2SFlash.uservalue = 1
            B2SFlash.enabled= true  
          else
            For HATPosition2=HATC2 to HATD2
                If HATPosition2-HATE2<=Len(HATText2) then
                    HATCharacter2=Ucase(Mid(HATText2,HATPosition2-HATE2,1))
                    For HATA2=1 to 8
                            HATGasSeg(HATPosition2,HATA2).State=(Mid(HATRom(Asc(HATCharacter2)),HATA2,1)*HATStyle2)
                    Next
                End If
            Next
        end if
    End If
 
End Sub
 
Sub FullJustify2(Tx,TxR)
    Select Case 6-Len(Tx)-Len(TxR)
        Case -1:Tx="999"&TxR
        Case 0:Tx=Tx&TxR
        Case 1:Tx=Tx&" "&TxR
        Case 2:Tx=Tx&"  "&TxR
        Case 3:Tx=Tx&"   "&TxR
        Case 4:Tx=Tx&"    "&TxR
        Case 5:Tx=Tx&"     "&TxR
        Case 6:Tx=Tx&"      "&TxR
    End Select
End Sub
 
Sub FullLengthText2(Tx)                             'Used to right justify for style 1,2
    Select Case Len(Tx)
        Case 1:Tx="           "&Tx
        Case 2:Tx="          "&Tx
        Case 3:Tx="         "&Tx
        Case 4:Tx="        "&Tx
        Case 5:Tx="       "&Tx
        Case 6:Tx="      "&Tx
        Case 7:Tx="     "&Tx
        Case 8:Tx="    "&Tx
        Case 9:Tx="   "&Tx
        Case 10:Tx="  "&Tx
        Case 11:Tx=" "&Tx
    End Select
End Sub
 
 
Sub HalfLengthText2(Tx)                             'Used to right justify for style 1,2
    Select Case Len(Tx)
        Case 1:Tx="     "&Tx
        Case 2:Tx="    "&Tx
        Case 3:Tx="   "&Tx
        Case 4:Tx="  "&Tx
        Case 5:Tx=" "&Tx
    End Select
End Sub
 
 
Sub Clear2                                          'Used to clear screen (top, bottom, or both)
    dim dogb2, dogg2
    if b2son then
        If HATScreen2=1 then dogb2=6 Else dogb2=12
        If HATScreen2=2 then dogg2=7 Else dogg2=1
        For HATA2=dogg2 to dogb2
            B2SDisplayChar "", HATA2
        Next
      else
        If HATScreen2=1 then HATB2=48 Else HATB2=96
        If HATScreen2=2 then HATG2=49 Else HATG2=1
        For HATA2=HATG2 To HATB2
            Eval("a"& HATA2).State=0
        Next
    end if
End Sub
 
 
sub B2SFlash2_timer
 
    Select Case B2SFlash2.uservalue
         Case 0:
            For HATPosition2=HATC2 to HATD2
                If HATPosition2-HATE2<=Len(HATText2) then
                    HATCharacter2=Ucase(Mid(HATText2,HATPosition2-HATE2,1))
                    B2SDisplayChar "", HATPosition2
                End If
            Next
            if B2sFlash2.uservalue<>2 then B2SFlash2.uservalue=1
        Case 1:
            For HATPosition2=HATC2 to HATD2
                If HATPosition2-HATE2<=Len(HATText2) then
                    HATCharacter2=Ucase(Mid(HATText2,HATPosition2-HATE2,1))
                    B2SDisplayChar HATCharacter2, HATPosition2
                End If
            Next
            if B2sFlash2.uservalue<>2 then B2SFlash2.uservalue=0
        Case 2:
            For HATPosition2=HATC2 to HATD2
                If HATPosition2-HATE2<=Len(HATText2) then
                    HATCharacter2=Ucase(Mid(HATText2,HATPosition2-HATE2,1))
                    B2SDisplayChar HATCharacter2, HATPosition2
                End If
            Next
            B2sFlash2.enabled=0
    End Select
end Sub
 
 
 
'==========================================================================================================================================
'============================================================= END OF 6-DIGIT Gottlieb Display =============================================================
'==========================================================================================================================================
 
'==========================================================================================================================================
'============================================================= START OF HIGH SCORES ROUTINES =============================================================
'==========================================================================================================================================
'
'ADD LINE TO TABLE_KEYDOWN SUB WITH THE FOLLOWING:    If HSEnterMode Then HighScoreProcessKey(keycode) AFTER THE STARTGAME ENTRY
'ADD And Not HSEnterMode=true TO IF KEYCODE=STARTGAMEKEY
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
'IMPORT POST-IT IMAGES
'
Dim HSA1, HSA2, HSA3, HGA1, HGA2, HGA3
Dim HSEnterMode, HGEnterMode, hsEnteredDigits(3), hsCurrentDigit, hsCurrentLetter
'Dim HSArray  
'Dim HSScoreM,HSScore100k, HSScore10k, HSScoreK, HSScore100, HSScore10, HSScore1, HSScorex  'Define 6 different score values for each reel to use
 
' ***********************************************************
'  HiPeg Record
' ***********************************************************
 
'Each Display command is then followed by text, style, screen, full, time, text2
 
Sub HighPegEntryInit(pl)
    HSA1=0:HSA2=0:HSA3=0
    Display "Hi Peg",2,1,0,0,""
    Display2 "P"&pl,1,2,0,0,""
    DisplayOne " ",0,10
    DisplayOne " ",0,11
    DisplayOne " ",0,12
    HSEnterMode = True
    hsCurrentLetter = 1:HSA1=65
    DisplayOne Chr(HSA1),2,10
End Sub
 
 
' ***********************************************************
'  HiGames Record
' ***********************************************************
 
'Each Display command is then followed by text, style, screen, full, time, text2
 
 
Sub HighGamesEntryInit(pg)
    HGA1=0:HGA2=0:HGA3=0
    Display "Hi Gms",2,1,0,0,""
    Display2 "P"&pg,1,2,0,0,""
    DisplayOne " ",0,10
    DisplayOne " ",0,11
    DisplayOne " ",0,12
    HGEnterMode = True
    hsCurrentLetter = 1:HGA1=65
    DisplayOne Chr(HGA1),2,10
End Sub
 
' ***********************************************************
'  HiPeg ENTER INITIALS
' ***********************************************************
 
Sub HighScoreProcessKey(keycode)
    if HSEnterMode=True then
        If keycode = LeftFlipperKey Then
            Select Case hsCurrentLetter
                Case 1:
                    HSA1=HSA1-1:If HSA1=64 Then HSA1=90 'no backspace on 1st digit
                    DisplayOne Chr(HSA1),2,10
                Case 2:
                    HSA2=HSA2-1:If HSA2=64 Then HSA2=60: If HSA2=59 Then HSA2=90
                    DisplayOne Chr(HSA2),2,11
                Case 3:
                    HSA3=HSA3-1:If HSA3=64 Then HSA3=60: If HSA3=59 then HSA3=90
                    DisplayOne Chr(HSA3),2,12
             End Select
        End If
 
        If keycode = RightFlipperKey Then
            Select Case hsCurrentLetter
                Case 1:
                    HSA1=HSA1+1:If HSA1>90 Then HSA1=65
                    DisplayOne Chr(HSA1),2,10
                Case 2:
                    HSA2=HSA2+1:If HSA2>90 Then HSA2=60: If HSA2=61 then HSA2=65
                    DisplayOne Chr(HSA2),2,11
                Case 3:
                    HSA3=HSA3+1:If HSA3>90 Then HSA3=60: If HSA3=61 then HSA3=65
                    DisplayOne Chr(HSA3),2,12
             End Select
        End If
       
        If keycode = StartGameKey Then
            Select Case hsCurrentLetter
                Case 1:
                    hsCurrentLetter=2       'ok to advance
                    DisplayOne Chr(HSA1),1,10
                    HSA2=HSA1               'start at same alphabet spot
                    DisplayOne Chr(HSA2),2,11
                Case 2:
                    If HSA2=60 Then         'bksp
                        DisplayOne " ",0,11
                        hsCurrentLetter=1
                        DisplayOne Chr(HSA1),2,10
                    Else
                        hsCurrentLetter=3   'enter it
                        DisplayOne Chr(HSA2),1,11
                        HSA3=HSA2           'start at same alphabet spot
                        DisplayOne Chr(HSA3),2,12
                    End If
                Case 3:
                    If HSA3=60 Then 'bksp
                        DisplayOne " ",0,12
                        hsCurrentLetter=2
                        DisplayOne Chr(HSA2),2,11
                    Else
                        DisplayOne Chr(HSA3),1,12
                        savehs 'enter it
                        HSEnterMode = False
                        b2sflash.uservalue=2
                        FlashOne.uservalue=2        'turn off flash
                        If NewHG>0 then
                            HighGamesEntryInit NewHG
                          else
                            Lgi.timerenabled=1
                            a1.timerenabled=1
                        End If
                    End If
            End Select
        End If
'   End If
 
' ***********************************************************
'  HiGames ENTER INITIALS
' ***********************************************************
   
    elseif HGEnterMode=True then
        If keycode = LeftFlipperKey Then
            Select Case hsCurrentLetter
                Case 1:
                    HGA1=HGA1-1:If HGA1=64 Then HGA1=90 'no backspace on 1st digit
                    DisplayOne Chr(HGA1),2,10
                Case 2:
                    HGA2=HGA2-1:If HGA2=64 Then HGA2=60: If HGA2=59 Then HGA2=90
                    DisplayOne Chr(HGA2),2,11
                Case 3:
                    HGA3=HGA3-1:If HGA3=64 Then HGA3=60: If HGA3=59 then HGA3=90
                    'DisplayOne "",0,12
                    DisplayOne Chr(HGA3),2,12
             End Select
        End If
 
        If keycode = RightFlipperKey Then
            Select Case hsCurrentLetter
                Case 1:
                    HGA1=HGA1+1:If HGA1>90 Then HGA1=65
                    DisplayOne Chr(HGA1),2,10
                Case 2:
                    HGA2=HGA2+1:If HGA2>90 Then HGA2=60: If HGA2=61 then HGA2=65
                    DisplayOne Chr(HGA2),2,11
                Case 3:
                    HGA3=HGA3+1:If HGA3>90 Then HGA3=60: If HGA3=61 then HGA3=65
                    DisplayOne Chr(HGA3),2,12
             End Select
        End If
       
        If keycode = StartGameKey Then
            Select Case hsCurrentLetter
                Case 1:
                    hsCurrentLetter=2 'ok to advance
                    DisplayOne Chr(HGA1),1,10
                    HGA2=HGA1 'start at same alphabet spot
                    DisplayOne Chr(HGA2),2,11
                Case 2:
                    If HGA2=60 Then 'bksp
                        DisplayOne " ",0,11
                        hsCurrentLetter=1
                        DisplayOne Chr(HGA1),2,10
                    Else
                        hsCurrentLetter=3 'enter it
                        DisplayOne Chr(HGA2),1,11
                        HGA3=HGA2 'start at same alphabet spot
                        DisplayOne Chr(HGA3),2,12
                    End If
                Case 3:
                    If HGA3=60 Then 'bksp
                        DisplayOne " ",0,12
                        hsCurrentLetter=2
                        DisplayOne Chr(HGA2),2,11
                    Else
                        DisplayOne Chr(HGA3),1,12
                        savehs 'enter it
                        HGEnterMode = False
                        b2sflash.uservalue=2
                        FlashOne.uservalue=2
                        Lgi.timerenabled=1
                        a1.timerenabled=1
                    End If
            End Select
        End If
    End If
 
End Sub
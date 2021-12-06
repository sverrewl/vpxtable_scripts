' -------------------------------------------------
' NOBS
' -------------------------------------------------
'
Option Explicit
Randomize

' by BorgDog, 2016
'
' Scoring Rules
'
'   Two seperate scoring mechanisms, Traditional pinball scoring (desktop mode only) and counting "hands"
'   to add to cribbage board count. Game tracks cribbage games won as well.
'
'   Drops score 1000 pts. Bumpers score 100. Rollovers score 100 or 1000 when lit.
'
'   Clear a hand to count cribbage points for that "hand"
'   including "turn" card (roto target), bank resets after clear.
'
'   Hands are as follows and all include the "turn" card (RotoTarget).:
'     1. 4 left side drops
'     2. 3 center drops and right side rollover
'     3. 3 top rollovers and left outlane
'     4. 2 right drops, right outlane, and left inlane
'
'   Clear K,Q,J and 7,8 rollovers to light star rollovers. Each lit star "pegs" 1 point during "play".
'     Clear any 6 lit wire rollovers to light extra ball (target) and open right outlane gate.
'
'   Clearing the 5 rollover opens left outlane gate, reset if hand 4 is complete or next ball.
'
'   When a player reaches an end of the cribbage board that player scores a game win.
'     If opposing players score is less than 90 the player also scores a skunk (another game).
'     Both players board pegging are reset to 0 when one scores a game.
'
'   If any rollover or drop target hit during play matches the turn card (roto) a pair (2pts) is pegged.
'   If any rollover or drop target hit during play + turn card = 15, then a "15 for 2" is pegged.
'   Hitting bullseye target or roto target spins roto.
'   Hitting roto pegs the displayed card (K,Q,J are worth 10, etc)
'   Hitting bullseye pegs 1.
'
'
' -------------------------------------------------

' Thalamus 2018-07-24
' Added/Updated "Positional Sound Playback Functions" and "Supporting Ball & Sound Functions"
' Due to ramp changes Trough1 and Trough2 needs Physics->Hit Threshold set to 0 on both left/right walls.

' Thalamus 2018-11-01 : Improved directional sounds
' Can still become better - mainly by adjusting for droptargets and reset
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
Const VolGates  = 1    ' Gates volume.
Const VolMetal  = 1    ' Metals volume.
Const VolRH     = 1    ' Rubber hits volume.
Const VolPi     = 1    ' Rubber pins volume.
Const VolTarg   = 1    ' Targets volume.
Const VolSpin   = 1.5  ' Spinners volume.
Const VolFlip   = 1    ' Flipper volume.

Const cGameName = "Nobs_2016"

Dim operatormenu, options
Dim balls, peg, pegstart
Dim Add10, Add100, Add1000
Dim hisc,hipg,higm
Dim maxplayers
Dim players
Dim player
Dim credit, decks, scoring
Dim score(2)
Dim Count(2)
Dim pegs(2)
Dim games(2)
Dim sreels(2)
Dim pegreels(2)
Dim gamereels(2)
Dim peglights(2,120)
Dim cplay(2)
Dim Pups(2)
Dim state
Dim tilt
Dim tiltsens
Dim target(9)
Dim ballinplay
dim ballrenabled, dtstep
dim rstep, lstep, rotospin, rotopos, rotocount, spinit
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
  maxplayers=2
  set sreels(1) = ScoreReel1
  set sreels(2) = ScoreReel2
  set pegreels(1) = PegReel1
  set pegreels(2) = PegReel2
  set gamereels(1) = GameReel1
  set gamereels(2) = GameReel2
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
  hideoptions
  player=1
  For each light in lights:light.State = 0: Next
  loadhs
  if rotopos="" or rotopos<1 or rotopos>13 then rotopos=1
  if rotospin <1 or rotospin>13 then rotospin=1
  RotoTarget.roty=-1*(rotopos-1)*(360/13)
  if hisc="" then hisc=10000
  if hipg="" then hipg=0
  if higm="" then higm=0
  hstxt.text=hisc
  hptxt.text=hipg
  hgtxt.text=higm
  gamov.text="Game Over"
  gamov.timerenabled=1
  tilttxt.timerenabled=1
  ii=0
  attractmode
  balls=5
  credittxt.setvalue(credit)
  if decks="" or decks<1 or decks>4 then decks=1
  OptionDeck.image = "OptionsDecks"&decks
  if scoring="" or scoring<1 or scoring>2 then scoring=1
  OptionScoring.image = "OptionsScoring"&scoring
  if scoring=1 then
    for each objekt in pinscoring: objekt.visible=false: Next
  end If
  for each objekt in plastics: objekt.image = "plastics"&decks: Next
  if decks=4 Then
    pcover.image = "PcardW"
    instcard1.visible= true
    instcard.visible= False
    Else
    pcover.image = "PcoverCard"
    instcard1.visible= False
    instcard.visible= True
  end if
  If B2SOn then
    setBackglass.enabled=true
    for each objekt in backdropstuff: objekt.visible = 0 : next
  End If
  for i = 1 to maxplayers
    sreels(i).setvalue(score(i))
    pegreels(i).setvalue(Count(i))
    gamereels(i).setvalue(games(i))
  next
  for i = 1 to 9
    target(i).isdropped=True
    DTlights(i-1).state=1
  Next
  PlaySound "motor"
  tilt=false
  Drain.CreateBall
End sub

sub attractmode
  pegani.enabled=1
  LightSeq1.Play SeqUpOn,15,1
  LightSeq1.Play SeqDownOn,15,1
  LightSeq1.Play SeqRightOn,15,1
  LightSeq1.Play SeqLeftOn,15,1
  LightSeq1.Play SeqRandom,5,,2000
end Sub

Sub LightSeq1_PlayDone()
  attractmode
End Sub

sub pegani_timer
  if ii>119 then
    ii=0
    for each light in peg1:light.state=0:Next
    for each light in peg2:light.state=0:Next
  end if
  Peg1(ii).state=1
  Peg2(ii).state=1
  ii=ii+1
end sub


sub setBackglass_timer
  Controller.B2ssetCredits Credit
    Controller.B2SSetGameOver 35,1
    Controller.B2SSetScorePlayer 5, hipg
    Controller.B2SSetScorePlayer 6, higm
  setBackglass.enabled=0
end Sub


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

Sub nobs_KeyDown(ByVal keycode)

  if keycode=AddCreditKey then
    playsound "coinin"
    addcredit
    end if

    if keycode=StartGameKey and credit>0 then
    if state=false then
    tilt=false
    state=true
    credit=credit-1
    playsound "cluper"
    credittxt.setvalue(credit)
    ballinplay=1
    If B2SOn Then
      Controller.B2ssetCredits Credit
      Controller.B2sStartAnimation "Startup"
      Controller.B2ssetballinplay 32, Ballinplay
      Controller.B2ssetplayerup 30, 1
      Controller.B2ssetcanplay 31, 1
      Controller.B2SSetGameOver 35, 0
      Controller.B2SSetScorePlayer 5, hipg
      Controller.B2SSetScorePlayer 6, higm
    End If
      pups(1).state=1
    tilt=false
    state=true
    playsound "initialize"
    players=1
    canplay1.state=1
    rst=0
    newgame.enabled=true
    else if state=true and players < maxplayers and Ballinplay=1 then
    credit=credit-1
    players=players+1
    CanPlay2.state=1
    credittxt.setvalue(credit)
    If B2SOn then
      Controller.B2ssetCredits Credit
      Controller.B2ssetcanplay 31, players
    End If
    playsound "cluper"
     end if
    end if
  end if

  If keycode = PlungerKey Then
    Plunger.PullBack
    PlaySoundAtVol "plungerpull",plunger,1
  End If

  If keycode=LeftFlipperKey and State = false and OperatorMenu=0 then
    OperatorMenuTimer.Enabled = true
  end if

  If keycode=LeftFlipperKey and State = false and OperatorMenu=1 then
    Options=Options+1
    If Options=4 then Options=1
    playsound "target"
    Select Case (Options)
      Case 1:
        Option1.visible=true
        Option3.visible=False
      Case 2:
        Option2.visible=true
        Option1.visible=False
      Case 3:
        Option3.visible=true
        Option2.visible=False
    End Select
  end if

  If keycode=RightFlipperKey and State = false and OperatorMenu=1 then
    PlaySound "metalhit2"
    Select Case (Options)
    Case 1:
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
      if decks=4 Then
        pcover.image = "PcardW"
        instcard1.visible= true
        instcard.visible= False
        Else
        pcover.image = "PcoverCard"
        instcard1.visible= False
        instcard.visible= True
      end if
      OptionDeck.image = "OptionsDecks"&decks
    Case 2:
      if scoring=1 then
        scoring=2
        For each objekt in pinscoring: objekt.visible=True: Next
        else
        scoring=1
        For each objekt in pinscoring: objekt.visible=false: Next
      end If
      OptionScoring.image = "OptionsScoring"&scoring
    Case 3:
      OperatorMenu=0
      savehs
      HideOptions
    End Select
  End If

  if tilt=false and state=true then
  If keycode = LeftFlipperKey Then
    LeftFlipper.RotateToEnd
    PlaySoundAtVol "flipperup", LeftFlipper, VolFlip
    PlaySoundAtVol "Buzz", LeftFlipper, VolFlip
  End If

  If keycode = RightFlipperKey Then
    RightFlipper.RotateToEnd
    PlaySoundAtVol "flipperup", RightFlipper, VolFlip
    PlaySoundAtVol "Buzz1", RightFlipper, VolFlip
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
  OptionDeck.visible = True
  OptionScoring.visible = True
End Sub

Sub HideOptions
  for each objekt In OptionMenu
    objekt.visible = false
  next
End Sub


Sub nobs_KeyUp(ByVal keycode)

  If keycode = PlungerKey Then
    Plunger.Fire
    PlaySoundAtVol "plunger", Plunger, 1
  End If

  if keycode = LeftFlipperKey then
    OperatorMenuTimer.Enabled = false
  end if

   If tilt=false and state=true then
  If keycode = LeftFlipperKey Then
    LeftFlipper.RotateToStart
    PlaySoundAtVol "flipperdown", LeftFlipper, VolFlip
    StopSound "Buzz"
  End If

  If keycode = RightFlipperKey Then
    RightFlipper.RotateToStart
    PlaySoundAtVol "flipperdown", RightFlipper, VolFlip
    StopSound "Buzz1"
  End If
   End if
End Sub

sub flippertimer_timer()
  LFlip.RotY = LeftFlipper.CurrentAngle
  RFlip.RotY = RightFlipper.CurrentAngle
  diverter.RotY = DiverterFlipper.CurrentAngle+90
  diverterL.RotY = DiverterFlipperL.CurrentAngle+90
end sub


Sub addcredit
      credit=credit+1
      if credit>15 then credit=15
    credittxt.setvalue(credit)
    If B2SOn Then Controller.B2ssetCredits Credit
End sub

Sub Drain_Hit()
  PlaySoundAtVol "drain", Drain, 1
  for each light in GIlights: light.state=0: next
  me.timerenabled=1
End Sub

Sub Drain_timer
    if shootagain.state=lightstateon and tilt=false then
        newball
        ballreltimer.enabled=true
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
  ballrenabled=1
end sub

sub ballrel_hit
  if ballrenabled=1 then
    shootagain.state=lightstateoff
    DiverterFlipper.RotateToStart
    DiverterWall.isdropped=False
    DiverterFlipperL.RotateToStart
    ballrenabled=0
  end if
end sub


sub newgame_timer
  bumper1.force=12
  bumper2.force=12
  player=1
  for i = 1 to maxplayers
    sreels(i).resettozero
    pegreels(i).resettozero
    gamereels(i).resettozero
    games(i)=0
      score(i)=0
    Count(i)=0
    pegs(i)=0
    rep(i)=0
    pups(i).state=0
  next
  pegani.enabled=0
  LightSeq1.StopPlay()
  for each light in BumperLights:light.state=1:Next
  for each light in peg1:light.state=0:Next
  for each light in peg2:light.state=0:Next
    pups(1).state=1
  If B2SOn then
    for i = 1 to maxplayers
    Controller.B2SSetScorePlayer i, score(i)
    next
    Controller.B2SSetGameOver 35,0
    Controller.B2SSetTilt 33,0
  End If
    eg=0
  shootagain.state=lightstateoff
    tilttxt.text=" "
  gamov.text=" "
    biptext.text="1"
  newball
  ballreltimer.enabled=true
  newgame.enabled=false
end sub

sub resetDT_timer
  Select Case DTStep
    Case 1:
      playsound "bankreset"
      for i = 1 to 4
        target(i).isdropped=False
        DTlights(i-1).state=0
      Next
    Case 2:
      playsound "bankreset"
      for i = 5 to 7
        target(i).isdropped=False
        DTlights(i-1).state=0
      Next
    Case 3:
      playsound "bankreset"
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
  for each light in lights:light.state=1:next
  for each light in starlights:light.state=0:next
  for each light in GIlights:light.state=1:next
End Sub


sub nextball
    if tilt=true then
    bumper1.force=12
    bumper2.force=12
      tilt=false
      tilttxt.text=" "
    If B2SOn then
      Controller.B2SSetTilt 33,0
      Controller.B2ssetdata 1, 1
    End If
    end if
  if player=1 then ballinplay=ballinplay+1
  if ballinplay>balls then
    playsound "GameOver"
    eg=1
    ballreltimer.enabled=true
    else
    if state=true then
      ballreltimer.enabled=true
    end if
    biptext.text=ballinplay
    If B2SOn then Controller.B2ssetballinplay 32, Ballinplay
  end if
End Sub

sub ballreltimer_timer
  if eg=1 then
    turnoff
    state=false
    biptext.text=" "
    gamov.text="GAME OVER"
    gamov.timerenabled=1
    tilttxt.timerenabled=1
    ii=0
    attractmode
    canplay1.state=0
    canplay2.state=0
    for i=1 to players
    if score(i)>hisc then hisc=score(i)
    if pegs(i)>hipg then hipg=pegs(i)
    if games(i)>higm then higm=games(i)
    pups(i).state=0
    next
    hstxt.text=hisc
    hptxt.text=hipg
    hgtxt.text=higm
    savehs
    If B2SOn then
        Controller.B2SSetGameOver 35,1
        Controller.B2ssetballinplay 32, 0
      Controller.B2sStartAnimation "EOGame"
      Controller.B2SSetScorePlayer 5, hipg
      Controller.B2SSetScorePlayer 6, higm
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

Sub plunger_timer
  Drain.kick 60, 11,0
  playsound "kickerkick"
  Plunger.timerenabled=false
end sub


'********** Bumpers

Sub Bumper1_Hit
   if tilt=false then
  PlaySoundAtVol "fx_bumper4", Bumper1, VolBump
  FlashBumpers
  addscore 100
  me.timerenabled=1
   end if
End Sub

Sub Bumper1_timer
  BumperTimerRing1.Enabled=0
  if bumperring1.transz>-36 then  BumperRing1.transz=BumperRing1.transz-4
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
  PlaySoundAtVol "fx_bumper4", Bumper2, VolBump
  FlashBumpers
  addscore 100
  me.timerenabled=1
   end if
End Sub

Sub Bumper2_timer
  BumperTimerRing2.enabled=0
  if BumperRing2.transz>-36 then  BumperRing2.transz=BumperRing2.transz-4
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
  if bumperlight1.timerenabled=false then
    for each light in BumperLights: light.state=0:Next
    BumperLight1.timerenabled=True
  end If
end sub

sub BumperLight1_timer
  for each light in BumperLights: light.state=1:Next
  BumperLight1.timerenabled=False
end Sub


'************** Slings

Sub RightSlingShot_Slingshot
    PlaySoundAtVol "slingshot", sling1, 1
  addscore 10
    RSling.Visible = 0
    RSling1.Visible = 1
    sling1.TransZ = -8
    RStep = 1
    RightSlingShot.TimerEnabled = 1
End Sub

Sub RightSlingShot_Timer
    Select Case RStep
        Case 3:RSLing1.Visible = 0:RSLing2.Visible = 1:sling1.TransZ = -8
        Case 4:sling1.TransZ = 0:RSLing2.Visible = 0:RSLing.Visible = 1:RightSlingShot.TimerEnabled = 0
    End Select
    RStep = RStep + 1
End Sub

sub Dingwalls_hit(idx)
  addscore 10
end sub


'********** Triggers


sub TGK_hit   '***** top King rollover
  if rotopos=13 then addcount 2
  if rotopos=5 then addcount 2  '15 for 2
  if (LK.state)=lightstateon then
    eball=eball+1
    checkeball
    addscore 1000
    LK.state=0
    hand4check rotopos
    Lstar1.state=1
    else
    addscore 100
    end if
 end sub

sub TGQ_hit   '***** top Queen rollover
  if rotopos=12 then addcount 2
  if rotopos=5 then addcount 2  '15 for 2
  if (LQ.state)=lightstateon then
    eball=eball+1
    checkeball
    addscore 1000
    LQ.state=0
    hand4check rotopos
    Lstar2.state=1
      else
    addscore 100
    end if

end sub

sub TGJ_hit   '***** top Jack rollover
  if rotopos=11 then addcount 2
  if rotopos=5 then addcount 2  '15 for 2
    if (LJ.state)=lightstateon then
    eball=eball+1
    checkeball
    addscore 1000
    LJ.state=0
    hand4check rotopos
    Lstar3.state=1
    else
    addscore 100
    end if
end sub

sub TG2_hit   '***** right 2 rollover
  if rotopos=2 then addcount 2
    if (L2.state)=lightstateon then
    eball=eball+1
    checkeball
    addscore 1000
    L2.state=0
    hand2check rotopos
    else
    addscore 100
    end if
end sub

sub TG7_hit   '***** left 7 rollover
  if rotopos=7 then addcount 2
  if rotopos=8 then addcount 2  '15 for 2
    if (L7.state)=lightstateon then
    eball=eball+1
    checkeball
    addscore 1000
    L7.state=0
    hand3check rotopos
    Lstar4.state=1
    else
    addscore 100
    end if
end sub

sub TG8_hit   '***** right 8 rollover
  if rotopos=8 then addcount 2
  if rotopos=7 then addcount 2  '15 for 2
    if (L8.state)=lightstateon then
    eball=eball+1
    checkeball
    addscore 1000
    L8.state=0
    hand3check rotopos
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
  DiverterWall.isdropped=True
  DiverterFlipper.RotateToEnd
  DiverterFlipper.timerenabled=0
end Sub

sub TG5_hit   '***** left 5 rollover
  if rotopos=5 then addcount 2
  if rotopos>9 then addcount 2
    if (L5.state)=lightstateon then
    addscore 1000
    L5.state=0
    hand4check rotopos
    else
    addscore 100
    end if
  DiverterFlipperL.timerenabled=1
end sub

sub DiverterFlipperL_timer
  DiverterFlipperL.RotateToEnd
  DiverterFlipperL.timerenabled=0
end Sub

sub TGstar1_hit   '*****star rollover
    if (Lstar1.state)=lightstateon then
    addscore 1000
    addcount(1)
    else
    addscore 100
    end if
end Sub

sub TGstar2_hit   '*****star rollover
    if (Lstar2.state)=lightstateon then
    addscore 1000
    addcount(1)
    else
    addscore 100
    end if
end Sub

sub TGstar3_hit   '*****star rollover
    if (Lstar3.state)=lightstateon then
    addscore 1000
    addcount(1)
    else
    addscore 100
    end if
end Sub

sub TGstar4_hit   '*****star rollover
    if (Lstar4.state)=lightstateon then
    addscore 1000
    addcount(1)
    else
    addscore 100
    end if
end Sub

sub TGstar5_hit   '*****star rollover
    if (Lstar5.state)=lightstateon then
    addscore 1000
    addcount(1)
    else
    addscore 100
    end if
end Sub

sub hand1check(pos1)
  dim hand1: hand1=0
  hand1=ldt5.state+ldt5b.state+ldt6.state+ldt4.state
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
    for each light in hand1lights:light.state=2:Next
    if DT5.timerenabled=false then dt5.timerenabled=True
  end If
end Sub

sub DT5_timer
  for i= 1 to 4:target(i).isdropped=0:Next
  playsound "bankreset"
  for each light in hand1lights:light.state=1:Next
  for each light in hand1dtlights:light.state=0:Next
  dt5.timerenabled=0
end sub

sub hand2check(pos2)
  dim hand2: hand2=0
  hand2=ldta.state+ldt2.state+ldt3.state
  if l2.state=0 then hand2=hand2+1
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
    for each light in hand2lights:light.state=2:Next
    if DTa.timerenabled=false then DTa.timerenabled=True
  end If
end Sub

sub DTa_timer
  for i= 5 to 7:target(i).isdropped=0:Next
  playsound "bankreset"
  for each light in hand2lights:light.state=1:Next
  for each light in hand2dtlights:light.state=0:Next
  DTa.timerenabled=0
end sub


sub hand3check(pos3)
  dim hand3
  hand3=0
  hand3=ldt6a.state+ldt9.state
  if l8.state=0 then hand3=hand3+1
  if l7.state=0 then hand3=hand3+1
  if hand3=4 Then
    select case(pos3)
      case 1:
        addcount(12)
      case 2:
        addcount(10)
      case 3,4,11,12,13:
        addcount(8)
      case 5,10:
        addcount(9)
      case 6,7,8,9:
        addcount(16)
    end Select
    for each light in hand3lights:light.state=2:Next
    if DT6a.timerenabled=false then DT6a.timerenabled=True
  end if
end Sub

sub dt6a_timer
  for i= 8 to 9:target(i).isdropped=0:Next
  playsound "bankreset"
  for each light in hand3lights:light.state=1:Next
  for each light in hand3DTlights:light.state=0:Next
  DT6a.timerenabled=0
end Sub

sub hand4check(pos4)
  dim hand4: hand4=0
  hand4=LK.state+LQ.state+LJ.state+L5.state
  if hand4=0 Then
    select Case(pos4)
      case 1,2,3,4,6,7,8,9:
        addcount(9)
      case 5:
        addcount(17)
      case 10:
        addcount(12)
  :   case 11,12,13:
        addcount(16)
    end Select
    for each light in hand4lights:light.state=2:Next
    if LK.timerenabled=false then LK.timerenabled=True
  end if
end Sub

sub lk_timer
  for each light in hand4lights:light.state=1:Next
  DiverterFlipperL.RotateToStart
  LK.timerenabled=0
end Sub

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
testbox.text=rotopos
    rotospin=rotospin+1
    if rotospin>13 then rotospin=1
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
    Else
    RotoTarget.roty=RotoTarget.roty+(360/52)
  end if
end sub

'********** Drop Targets

sub DTa_dropped
  if state=true then
    if rotopos=1 then addcount 2
    addscore 1000
    ldta.state=1
    hand2check rotopos
  end If
end sub

sub DT2_dropped
  if state=true then
    if rotopos=2 then addcount 2
    addscore 1000
    Ldt2.state=1
    hand2check rotopos
  end if
end sub

sub DT3_dropped
  if state=true then
    if rotopos=3 then addcount 2
    addscore 1000
    Ldt3.state=1
    hand2check rotopos
  end if
end sub

sub DT4_dropped
  if state=true then
    if rotopos=4 then addcount 2
    addscore 1000
    ldt4.state=1
    hand1check rotopos
  end if
end sub

sub DT5_dropped
  if state=true then
    if rotopos=5 then addcount 2
    if rotopos>9 then addcount 2
    addscore 1000
    ldt5.state=1
    hand1check rotopos
  end if
end sub

sub DT5b_dropped
  if state=true then
    if rotopos=5 then addcount 2
    if rotopos>9 then addcount 2
    addscore 1000
    ldt5b.state=1
    hand1check rotopos
  end if
end sub

sub DT6_dropped
  if state=true then
    if rotopos=6 then addcount 2
    if rotopos=9 then addcount 2
    addscore 1000
    Ldt6.state=1
    hand1check rotopos
  end If
end sub

sub DT6A_dropped
  if state=true then
    if rotopos=6 then addcount 2
    if rotopos=9 then addcount 2
    addscore 1000
    ldt6a.state=1
    hand3check rotopos
  end If
end sub

sub DT9_dropped
  if state=true then
    if rotopos=9 then addcount 2
    if rotopos=6 then addcount 2
    addscore 1000
    Ldt9.state=1
    hand3check rotopos
  end if
end sub

sub addcount(peg)
  pegbox.text="Pegged "&peg
  playsound "woodpeg"
  PegBox.timerenabled=1
  pegstart=Count(player)+1
  Count(player)=Count(player)+peg
  pegs(player)=pegs(player)+peg
  pegreels(player).setvalue(pegs(player))
  If B2SOn Then Controller.B2SSetScorePlayer player, pegs(player)
  if Count(player)>120 Then
    playsound "gong"
    games(player)=games(player)+1
    gamereels(player).setvalue(games(player))
    if players=2 then                  'check for skunk 2 player game only
      if player=1 and count(2)<60 Then
        playsound "knockknock"
        games(1)=games(1)+3
        else if player=1 and count(2)<90 Then
        playsound "knock"
        games(1)=games(1)+1
        end If
      end If
      if player=2 and Count(1)<60 Then
        playsound "knockknock"
        games(2)=games(2)+3
        else if player=2 and count(1)<90 Then
        playsound "knock"
        games(2)=games(2)+1
        end If
      end If
    end if
    If B2SOn Then Controller.B2SSetScorePlayer player+2, games(player)
    for each light in Peg1:light.state=0: Next
    for each light in Peg2:light.state=0:Next
    Count(1)=0
    Count(2)=0
    Else
    for i = pegstart to Count(player)
      if player=1 then Peg1(i-1).state = 2
      if player=2 then Peg2(i-1).state = 2
    Next
  end If
end Sub

sub pegbox_timer
  pegbox.text=" "
  for i = 1 to Count(player)
    if player=1 then Peg1(i-1).state = 1
    if player=2 then Peg2(i-1).state = 1
  Next
  pegbox.timerenabled=0
end sub

sub addscore(points)
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
    score(player)=score(player)+points
  sreels(player).addvalue(points)
' If B2SOn Then Controller.B2SSetScorePlayer player, score(player)

    ' Sounds: there are 3 sounds: tens, hundreds and thousands
    If Points = 100 AND(Score(player) MOD 1000) \ 100 = 0 Then  'New 1000 reel
        PlaySound SoundFXDOF("bell1000",143,DOFPulse,DOFChimes)
      ElseIf Points = 10 AND(Score(player) MOD 100) \ 10 = 0 Then 'New 100 reel
        PlaySound SoundFXDOF("bell100",142,DOFPulse,DOFChimes)
      ElseIf points = 1000 Then
        PlaySound SoundFXDOF("bell1000",143,DOFPulse,DOFChimes)
    elseif Points = 100 Then
        PlaySound SoundFXDOF("bell100",142,DOFPulse,DOFChimes)
      Else
        PlaySound SoundFXDOF("bell10",141,DOFPulse,DOFChimes)
    End If
end sub




Sub CheckTilt
  If Tilttimer.Enabled = True Then
   TiltSens = TiltSens + 1
   if TiltSens = 3 Then
     Tilt = True
     tilttxt.text="TILT"
        If B2SOn Then Controller.B2SSetTilt 33,1
        If B2SOn Then Controller.B2ssetdata 1, 0
     playsound "tilt"
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
    bumper1.force=0
    bumper2.force=0
    LeftFlipper.RotateToStart
  StopSound "Buzz"
  RightFlipper.RotateToStart
  StopSound "Buzz1"
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

Sub Pins_Hit (idx)
  PlaySound "pinhit_low", 0, Vol(ActiveBall)*VolPi, AudioPan(ActiveBall), 0, Pitch(ActiveBall), 0, 0, AudioFade(ActiveBall)
End Sub

Sub Targets_Hit (idx)
  PlaySound "target", 0, Vol(ActiveBall)*VolTarg, AudioPan(ActiveBall), 0, Pitch(ActiveBall), 0, 0, AudioFade(ActiveBall)
End Sub

Sub Metals_Thin_Hit (idx)
  PlaySound "metalhit_thin", 0, Vol(ActiveBall)*VolMetal, AudioPan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
End Sub

Sub Metals_Medium_Hit (idx)
  PlaySound "metalhit_medium", 0, Vol(ActiveBall)*VolMetal, AudioPan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
End Sub

Sub Metals2_Hit (idx)
  PlaySound "metalhit2", 0, Vol(ActiveBall)*VolMetal, AudioPan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
End Sub

Sub Gates_Hit (idx)
  PlaySound "gate4", 0, Vol(ActiveBall)*VolGates, AudioPan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
End Sub

Sub Spinner_Spin
  PlaySoundAtVol "fx_spinner", Spinner, VolSpin
End Sub

Sub a_Rubbers_Hit(idx)
  dim finalspeed
    finalspeed=SQR(activeball.velx * activeball.velx + activeball.vely * activeball.vely)
  If finalspeed > 20 then
    PlaySound "fx_rubber2", 0, Vol(ActiveBall)*VolRH, Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
  End if
  If finalspeed >= 6 AND finalspeed <= 20 then
    RandomSoundRubber()
  End If
End Sub

Sub RubberWheel_hit
  dim finalspeed
    finalspeed=SQR(activeball.velx * activeball.velx + activeball.vely * activeball.vely)
  If finalspeed > 20 then
    PlaySound "fx_rubber2", 0, Vol(ActiveBall)*VolRH, Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
  End if
  If finalspeed >= 6 AND finalspeed <= 20 then
    RandomSoundRubber()
  End If
End sub

Sub a_Posts_Hit(idx)
  dim finalspeed
    finalspeed=SQR(activeball.velx * activeball.velx + activeball.vely * activeball.vely)
  If finalspeed > 16 then
    PlaySound "fx_rubber2", 0, Vol(ActiveBall)*VolRH, Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
  End if
  If finalspeed >= 6 AND finalspeed <= 16 then
    RandomSoundRubber()
  End If
End Sub

Sub RandomSoundRubber()
  Select Case Int(Rnd*3)+1
    Case 1 : PlaySound "rubber_hit_1", 0, Vol(ActiveBall)*VolRH, Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
    Case 2 : PlaySound "rubber_hit_2", 0, Vol(ActiveBall)*VolRH, Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
    Case 3 : PlaySound "rubber_hit_3", 0, Vol(ActiveBall)*VolRH, Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
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
    Case 1 : PlaySound "flip_hit_1", 0, Vol(ActiveBall)*VolRH, Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
    Case 2 : PlaySound "flip_hit_2", 0, Vol(ActiveBall)*VolRH, Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
    Case 3 : PlaySound "flip_hit_3", 0, Vol(ActiveBall)*VolRH, Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
  End Select
End Sub


sub savehs

    savevalue "nobs", "credit", credit
    savevalue "nobs", "hiscore", hisc
    savevalue "nobs", "hipeg", hipg
    savevalue "nobs", "higame", higm
    savevalue "nobs", "score1", score(1)
    savevalue "nobs", "score2", score(2)
    savevalue "nobs", "count1", Count(1)
    savevalue "nobs", "count2", Count(2)
  savevalue "nobs", "games1", games(1)
  savevalue "nobs", "games2", games(2)
  savevalue "nobs", "decks", decks
  savevalue "nobs", "scoring", scoring
  savevalue "nobs", "rotopos", rotopos
end sub

sub loadhs
  temp = LoadValue("nobs", "credit")
    If (temp <> "") then credit = CDbl(temp)
    temp = LoadValue("nobs", "hiscore")
    If (temp <> "") then hisc = CDbl(temp)
    temp = LoadValue("nobs", "hipeg")
    If (temp <> "") then hipg = CDbl(temp)
    temp = LoadValue("nobs", "higame")
    If (temp <> "") then higm = CDbl(temp)
    temp = LoadValue("nobs", "score1")
    If (temp <> "") then score(1) = CDbl(temp)
    temp = LoadValue("nobs", "score2")
    If (temp <> "") then score(2) = CDbl(temp)
    temp = LoadValue("nobs", "count1")
    If (temp <> "") then count(1) = CDbl(temp)
    temp = LoadValue("nobs", "count2")
    If (temp <> "") then count(2) = CDbl(temp)
    temp = LoadValue("nobs", "games1")
    If (temp <> "") then games(1) = CDbl(temp)
    temp = LoadValue("nobs", "games2")
    If (temp <> "") then games(2) = CDbl(temp)
    temp = LoadValue("nobs", "decks")
    If (temp <> "") then decks = CDbl(temp)
    temp = LoadValue("nobs", "scoring")
    If (temp <> "") then scoring = CDbl(temp)
  temp = LoadValue("nobs", "rotopos")
    If (temp <> "") then rotopos = CDbl(temp)
end sub

Sub nobs_Exit()
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

Sub PlaySoundAtVol(sound, tableobj, Volum)
  PlaySound sound, 1, Volum, Pan(tableobj), 0,0,0, 1, AudioFade(tableobj)
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

Function AudioFade(tableobj) ' Fades between front and back of the table (for surround systems or 2x2 speakers, etc), depending on the Y position on the table. "nobs" is the name of the table
  Dim tmp
  tmp = tableobj.y * 2 / nobs.height-1
  If tmp > 0 Then
    AudioFade = Csng(tmp ^10)
  Else
    AudioFade = Csng(-((- tmp) ^10) )
  End If
End Function

Function AudioPan(tableobj) ' Calculates the pan for a tableobj based on the X position on the table. "nobs" is the name of the table
  Dim tmp
  tmp = tableobj.x * 2 / nobs.width-1
  If tmp > 0 Then
    AudioPan = Csng(tmp ^10)
  Else
    AudioPan = Csng(-((- tmp) ^10) )
  End If
End Function

Function Pan(ball) ' Calculates the pan for a ball based on the X position on the table. "nobs" is the name of the table
    Dim tmp
    tmp = ball.x * 2 / nobs.width-1
    If tmp > 0 Then
        Pan = Csng(tmp ^10)
    Else
        Pan = Csng(-((- tmp) ^10) )
    End If
End Function

Function AudioFade(ball) ' Can this be together with the above function ?
  Dim tmp
  tmp = ball.y * 2 / nobs.height-1
  If tmp > 0 Then
    AudioFade = Csng(tmp ^10)
  Else
    AudioFade = Csng(-((- tmp) ^10) )
  End If
End Function

Function Vol(ball) ' Calculates the Volume of the sound based on the ball speed
  Vol = Csng(BallVel(ball) ^2 / VolDiv)
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
    PlaySound("fx_collide"), 0, Csng(velocity) ^2 / (VolDiv/VolCol), Pan(ball1), 0, Pitch(ball1), 0, 0, AudioFade(ball1)
End Sub


' Thalamus : Exit in a clean and proper way
Sub nobs_exit()
  Controller.Pause = False
  Controller.Stop
End Sub


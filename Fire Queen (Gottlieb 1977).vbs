' -------------------------------------------------'
' FIRE QUEEN  Gottlieb 1977
' -------------------------------------------------
' Original table by D. Gottlieb & Co., October 1977
' VP   Vulcan adaptation by Dave Sanders, December 2002
' VP10 Vulcan adaptation started by MaX and hauntfreaks, July 2015
'   VP10 Fire Queen and Vulcan by hauntfreaks and BorgDog, Sept 2015
' DOF coded by BorgDog with instruction(and correction) by arngrim - thanks!
' -------------------------------------------------

Option Explicit
Randomize

' Thalamus 2018-11-01 : Improved directional sounds
' !! NOTE : Table not verified yet !!

' Thalamus, table uses its own SSF routines.

' Options
' Volume devided by - lower gets higher sound

Const VolDiv = 2000    ' Lower number, louder ballrolling/collition sound
Const VolCol = 10      ' Ball collition divider ( voldiv/volcol )

Const VolFlip   = 1    ' Flipper volume.


Const cGameName = "firequeen_1977"
Const Ballsize = 50     'used by ball shadow routine
Dim operatormenu, options, optionsX
Dim bumperlitscore
Dim bumperoffscore
Dim balls
dim ebcount
Dim replays
Dim Replay1Table(3)
Dim Replay2Table(3)
Dim Replay3Table(3)
dim replay1
dim replay2
dim replay3
Dim hisc, hiscstate
Dim Controller
Dim maxplayers
Dim players
Dim player
Dim credit, dbonus, freeplay, ballshadows, flippershadows
Dim score(4)
Dim Add10, Add100, Add1000
Dim sreels(4)
Dim state
Dim tilt
Dim tiltsens
Dim holebonus
Dim holeb(4)
Dim StarState
Dim Bonus, chimescount
Dim ballinplay
Dim matchnumb
dim ballrenabled
Dim rep(4)
Dim rst
Dim eg
Dim bell
Dim i,j,tnumber, objekt, light
Dim awardcheck

On Error Resume Next
ExecuteGlobal GetTextFile("controller.vbs")
If Err Then MsgBox "Can't open controller.vbs"
On Error Goto 0

sub FireQueen_init
  LoadEM
  maxplayers=2
  Replay1Table(1)=70000
  Replay1Table(2)=90000
  Replay1Table(3)=100000
  Replay2Table(1)=120000
  Replay2Table(2)=110000
  Replay2Table(3)=120000
  Replay3Table(1)=150000
  Replay3Table(2)=160000
  Replay3Table(3)=170000
  set sreels(1) = ScoreReel1
  set sreels(2) = ScoreReel2
  set sreels(3) = ScoreReel3
  set sreels(4) = ScoreReel4

  set holeb(1)=Hole1
  set holeb(2)=Hole2
  set holeb(3)=Hole3
  set holeb(4)=Hole4
  hideoptions
  player=1


  hisc=50000    'SET DEFAULT VALUES PRIOR TO LOADING FROM SAVED FILE
  balls=5
  replays=2
  credit=0
  freeplay=0
  flippershadows=1
  ballshadows=1
  matchnumb=0
  HSA1=4
  HSA2=15
  HSA3=7

  if ShowDT then
    OptionReel.objrotx=10
    OptionsReel.objrotx=0
    OptionBox.objrotx=-14
    OptionBox.z=25
    OptionReel.z=30
    OptionsReel.z=30
    else
    OptionReel.objrotx=28
    OptionsReel.objrotx=15
    OptionBox.objrotx=0
    OptionBox.z=0
    OptionReel.z=5
    OptionsReel.z=5
  end If

  loadhs      'LOADS SAVED OPTIONS AND HIGH SCORE IF EXISTING
  UpdatePostIt  'UPDATE HIGH SCORE STICKY



  credittxt.setvalue(credit)


  Replay1=Replay1Table(Replays)
  Replay2=Replay2Table(Replays)
  Replay3=Replay3Table(Replays)

  RepCard.image = "ReplayCard"&replays
  bumperlitscore=100
  bumperoffscore=100
  if balls=3 then
    InstCard.image="InstCard3balls"
    Else
    InstCard.image="InstCard5balls"
  end if
  If B2SOn Then
    setBackglass.enabled=true
    For each objekt in Backdropstuff: objekt.visible=false: next
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

  for i = 1 to maxplayers
    sreels(i).setvalue(score(i))

  next

  tilt=false
  startGame.enabled=true
    Drain.CreateBall
End sub

sub startGame_timer
  PlaySoundAt "poweron", Plunger
  setBackglass.enabled=true
  me.enabled=false
end sub

sub setBackglass_timer
  GOReel.setvalue 1
  GOReel.timerenabled=1
  if matchnumb=0 then
    matchtxt.text="00"
    else
    matchtxt.text=matchnumb*10
  end if

  If credit>0 then DOF 133, DOFOn
  for each light in bumperlights:light.state=1:next
  for each light in GIlights:light.state=1:next
  shadowsGION.visible=true
  outerleft_prim.image = "outerlefton"
  outerright_prim.image = "outerrighton"
  for i=1 to maxplayers
    EVAL("r100k"&i).setvalue INT(score(i)/100000)
  next
  if b2son then
    Controller.B2ssetCredits Credit
    if matchnumb=0 then
      Controller.B2ssetMatch 34, 100
      else
      Controller.B2ssetMatch 34, Matchnumb*10
    end if
    Controller.B2ssetdata 1, 1
    Controller.B2SSetGameOver 35,1
    for i = 1 to maxplayers
      Controller.B2SSetScorePlayer i, Score(i) MOD 100000
      if score(i)>99999 then Controller.B2SSetScoreRollover 24 + i, 1
    next
  end if
  me.enabled=false
end sub

sub GOreel_timer
  if state=false then
    If B2SOn then Controller.B2SSetGameOver 35,0
    GOReel.setvalue 0
    gtimer.enabled=true
  end if
  GOReel.timerenabled=0
end sub

sub gtimer_timer
  if state=false then
    GOReel.setvalue 1
    If B2SOn then Controller.B2SSetGameOver 35,1
    GOReel.timerinterval= (INT (RND*10)+5)*100
    GOReel.timerenabled=1
  end if
  me.enabled=0
end sub



Sub FireQueen_KeyDown(ByVal keycode)

  if keycode=AddCreditKey then
    if freeplay=1 or credit>15 Then
      playsoundat "coin3", Drain
      Else
      playsoundat "coin6", Drain
      addcredit
    end If
    end if

   if keycode=StartGameKey and (Credit>0 or freeplay=1) And Not HSEnterMode=true and Not Startgame.enabled then
    if state=false then
    if freeplay=0 then
      credit=credit-1
      if credit < 1 then DOF 133, DOFOff
      credittxt.setvalue(credit)
    end if
    playsound "cluper"

    ballinplay=1
    If b2son Then
      Controller.B2ssetCredits Credit
      Controller.B2sStartAnimation "Startup"
      Controller.B2ssetballinplay 32, Ballinplay
      Controller.B2ssetplayerup 30, 1
      Controller.B2ssetcanplay 31, 1
      Controller.B2SSetGameOver 0
    End If
    shoot1.state=1
    tilt=false
    state=true
    CanPlay1.state=1
    playsound "initialize"
    players=1
    rst=0
    resettimer.enabled=true
    else if state=true and players < maxplayers and Ballinplay=1 and (Credit>0 or freeplay=1) then
    if freeplay=0 then
      credit=credit-1
      if credit < 1 then DOF 133, DOFOff
      credittxt.setvalue(credit)
    end if
    players=players+1
    If b2son then
      Controller.B2ssetCredits Credit
      Controller.B2ssetcanplay 31, 2
    End If
    CanPlay2.state=1
    playsound "cluper"
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

  If keycode=LeftFlipperKey and State = false and OperatorMenu=1 and Not OptionReelTimer.enabled and Not OptionsReelTimer.enabled then
    Options=Options+1
    If Options=7 then Options=1
    PlaySoundAt SoundFXDOF("reelclick",104,DOFPulse,DOFContactors), OptionReel
    OptionReelTimer.enabled=true
    Select Case (Options)
      Case 1:
        if Balls=3 then
          OptionsX=1
          else
          OptionsX=2
        end if
      Case 2:
        if ballshadows=0 Then
          OptionsX=4      'NO
          Else
          OptionsX=3    'yes
        end If
      Case 3:
        if freeplay=0 Then
          OptionsX=4      'no
          Else
          OptionsX=3    'yes
        end if
      Case 4:
        Select Case (replays)
          Case 1:
            OptionsX=5    'Low
          Case 2:
            OptionsX=6    'Med
          Case 3:
            OptionsX=7    'High
        End Select
      Case 5:
        if flippershadows=0 Then
          OptionsX=4      'NO
          Else
          OptionsX=3    'yes
        end If
      Case 6:
        OptionsX=8
    End Select
    OptionsReelTimer.enabled=true
  end if

  If keycode=RightFlipperKey and State = false and OperatorMenu=1 and Not OptionReelTimer.enabled and Not OptionsReelTimer.enabled then
    PlaySoundAt SoundFXDOF("reelclick",103,DOFPulse,DOFContactors), OptionsReel
    Select Case (Options)
      Case 1:
      if Balls=3 then
        Balls=5
        InstCard.image="InstCard5balls"
        OptionsX=2
        else
        Balls=3
        InstCard.image="InstCard3balls"
        OptionsX=1
      end if
      OptionsReelTimer.enabled=true
    Case 2:
      if ballshadows=0 Then
        ballshadows=1
        BallShadowUpdate.enabled=1
        OptionsX=3      'yes
        Else
        ballshadows=0
        BallShadowUpdate.enabled=0
        OptionsX=4    'no
      end If
      OptionsReelTimer.enabled=true
    Case 3:
      if freeplay=0 Then
        freeplay=1
        OptionsX=3      'yes
        Else
        freeplay=0
        OptionsX=4    'no
      end if
      OptionsReelTimer.enabled=true
    Case 4:
      Replays=Replays+1
      if Replays>3 then
        Replays=1
      end if
      Replay1=Replay1Table(Replays)
      Replay2=Replay2Table(Replays)
      replay3=Replay3Table(Replays)
      repcard.image = "replaycard"&replays
      Select Case (replays)
        Case 1:
          OptionsX=5    'Low
        Case 2:
          OptionsX=6    'Med
        Case 3:
          OptionsX=7    'High
      End Select
      OptionsReelTimer.enabled=true
    Case 5:
      if flippershadows=0 Then
        flippershadows=1
        FlipperLSh.visible=1
        FlipperRSh.visible=1
        OptionsX=3      'yes
        Else
        flippershadows=0
        FlipperLSh.visible=0
        FlipperRSh.visible=0
        OptionsX=4    'no
      end If
      OptionsReelTimer.enabled=true
    Case 6:
      OperatorMenu=0
      savehs
      HideOptions
    End Select
  End If

    If HSEnterMode Then HighScoreProcessKey(keycode)

  if tilt=false and state=true then
  If keycode = LeftFlipperKey Then
    LeftFlipper.RotateToEnd
    PlaySoundat SoundFXDOF("flipperup",101,DOFOn,DOFContactors), LeftFlipper
    PlaySoundAtVol "Buzz", LeftFlipper, VolFlip
  End If

  If keycode = RightFlipperKey Then
    RightFlipper.RotateToEnd
    PlaySoundat SoundFXDOF("flipperup",102,DOFOn,DOFContactors), RightFlipper
    PlaySoundAtVol "Buzz1",RightFlipper, VolFlip
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

Sub OptionReelTimer_timer
  PlaySoundAt "fx_switch", OptionsReel
  if OptionReel.roty <> ((Options*60)-60) Then
    optionreel.roty=(OptionReel.roty+10) mod 360
    Else
    me.enabled=False
  end If
end Sub

Sub OptionsReelTimer_timer
  PlaySoundAt "fx_switch", OptionsReel
  if OptionsReel.roty <> ((OptionsX*45)-45) Then
    OptionsReel.roty=(OptionsReel.roty+9) mod 360
    Else
    me.enabled=False
  end If
end Sub

Sub DisplayOptions

  OptionReel.visible = True
  OptionsReel.visible = True
  if Balls=3 then
    OptionsReel.roty=0
    else
    OptionsReel.roty=45
  end if
  OptionReel.roty=0
  OptionBox.visible = true
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


Sub FireQueen_KeyUp(ByVal keycode)

  If keycode = PlungerKey Then
    Plunger.Fire
    if ballhome.BallCntOver>0 then
      PlaySoundAt "plungerreleaseball", Plunger
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
    StopSound "Buzz"
  End If

  If keycode = RightFlipperKey Then
    RightFlipper.RotateToStart
    PlaySoundAt SoundFXDOF("flipperdown",102,DOFOff,DOFContactors), RightFlipper
    StopSound "Buzz1"
  End If
   End if
End Sub

sub flippertimer_timer()
  LFlip.RotZ = LeftFlipper.CurrentAngle
  RFlip.RotZ = RightFlipper.CurrentAngle
  LflipRubber.RotZ = LeftFlipper.CurrentAngle
  rflipRubber.RotZ = RightFlipper.CurrentAngle
  Pgate.rotx = Gate.currentangle*.6
  if w1a.isdropped=1 then dropplate1.image = "blank"
  if w2a.isdropped=1 then dropplate2.image = "blank"
  if w3a.isdropped=1 then dropplate3.image = "blank"
  if w4a.isdropped=1 then dropplate4.image = "blank"
  if w5a.isdropped=1 then dropplate5.image = "blank"
  if g4a.isdropped=1 then dropplate6.image = "blank"
  if g3a.isdropped=1 then dropplate7.image = "blank"
  if g2a.isdropped=1 then dropplate8.image = "blank"
  if g1a.isdropped=1 then dropplate9.image = "blank"
  if llane3.state = 1 and w1a.isdropped=0 then dropplate1.image = "drop1"
  if llane3.state = 1 and w2a.isdropped=0 then dropplate2.image = "drop2"
  if llane3.state = 1 and w3a.isdropped=0 then dropplate3.image = "drop3"
  if llane3.state = 1 and w4a.isdropped=0 then dropplate4.image = "drop4"
  if llane3.state = 1 and w5a.isdropped=0 then dropplate5.image = "drop5"
  if llane3.state = 1 and g4a.isdropped=0 then dropplate6.image = "drop6"
  if llane3.state = 1 and g3a.isdropped=0 then dropplate7.image = "drop7"
  if llane3.state = 1 and g2a.isdropped=0 then dropplate8.image = "drop8"
  if llane3.state = 1 and g1a.isdropped=0 then dropplate9.image = "drop9"
  if llane3.state = 1 then outerleft_prim.blenddisablelighting = 0.5
  if llane3.state = 0 then outerleft_prim.blenddisablelighting = 0.2
  if llane3.state = 1 then outerright_prim.blenddisablelighting = 0.5
  if llane3.state = 0 then outerright_prim.blenddisablelighting = 0.2
  if llane3.state = 1 then metaloutlaneR_prim.blenddisablelighting = 0.5
  if llane3.state = 0 then metaloutlaneR_prim.blenddisablelighting = 0.1
  if llane3.state = 1 then metaloutlaneL_prim.blenddisablelighting = 0.5
  if llane3.state = 0 then metaloutlaneL_prim.blenddisablelighting = 0.2

  if Llane3.state=0 then shadowsGIOFF.visible=1
  if Llane3.state=0 then shadowsGION.visible=0
  if Llane3.state=1 then shadowsGIOFF.visible=0
  if Llane3.state=1 then shadowsGION.visible=1

  if flippershadows=1 then
    FlipperLSh.RotZ = LeftFlipper.currentangle
    FlipperRSh.RotZ = RightFlipper.currentangle
  end if

end sub

Sub SPtimer_timer
  starstate=0
  if top1.state=0 then starstate=starstate+1
  if top2.state=0 then starstate=starstate+1
  if top3.state=0 then starstate=starstate+1
  if top4.state=0 then starstate=starstate+1
  if spot5.state=0 then starstate=starstate+1
  if starstate=5 then LeftSP.state=1
end sub

Sub PairedlampTimer_timer
  bottom1.state = top1.state
  bottom2.state = top2.state
  bottom3.state = top3.state
  bottom4.state = top4.state
  roll1.state = drop1.state
  roll2.state = drop2.state
  roll3.state = drop3.state
  roll4.state = drop4.state
  roll5.state = drop5.state
  rightSP.state = leftSP.state
  ShootA1.state = Shoot1.state
  ShootA2.state = Shoot2.state
end sub


sub resettimer_timer
    rst=rst+1
  for i = 1 to maxplayers
    sreels(i).resettozero
    EVAL("r100k"&i).setvalue 0
    next
    If b2son then
    for i = 1 to maxplayers
      Controller.B2SSetScorePlayer i, score(i)  MOD 100000
      Controller.B2SSetScoreRollover 24 + i, 0
    next
  End If
    if rst=22 then
    newgame
    resettimer.enabled=false
    end if
end sub

Sub addcredit
  if freeplay=0 then
      credit=credit+1
    DOF 133, DOFOn
      if credit>15 then credit=15
    credittxt.setvalue(credit)
    If b2son Then Controller.B2ssetCredits Credit
  end if
End sub

Sub Drain_Hit()
  DOF 129, DOFPulse
  PlaySoundAt "drain", Drain
  me.timerenabled=1
End Sub

Sub Drain_timer
  scorebonus.enabled=true
  me.timerenabled=0
End Sub


sub ballhome_unhit
  DOF 134, DOFPulse
end sub



sub ScoreBonus_timer
  if tilt=true then
    bonus=0
    else
    if bonusx2.state=1 then
      dbonus=2
      else
      dbonus=1
    End if
        if bonus>0 then
       if (bonus*dbonus)>3 then
        chimescount=4
        elseif (bonus*dbonus)>2 then
        chimescount=3
        elseif (bonus*dbonus)>1 then
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
     if shootagain.state=1 then
      ballreltimer.enabled=true
    else
'     if players=1 or player=players then
'       player=1
'       If b2son then Controller.B2ssetplayerup 30, 1
'       shoot1.state=1
'       shoot2.state=0
'       nextball
'     else
'       player=player+1
'       If b2son then Controller.B2ssetplayerup 30, player
'       shoot1.state=0
'       shoot2.state=1
        nextball
'     end if
   end if
   me.enabled=false
   end if
End sub

sub chimestimer_timer
    addpoints 1000

  select case chimescount
    case 4,2:
      if dbonus=1 then
        BonusLights(Bonus).state=0
        bonus=bonus-1
        if bonus>0 then
          BonusLights(Bonus).state=1
          if bonus=19 then Bonus10.state=1
        end if
      end if
    case 3,1:
      BonusLights(Bonus).state=0
      bonus=bonus-1
      if bonus>0 then
        BonusLights(Bonus).state=1
        if bonus=19 then Bonus10.state=1
      end if
  end select
'    if (dbonus>1 and chimescount=1) or dbonus=1 then
'        EVAL("Bonus"&bonus).state=0
'        bonus=bonus-1
'        if bonus > 0 then
'            EVAL("Bonus"&bonus).state=1
'            if bonus=19 then bonus10.state=1
'        end if
'    end if
    chimescount = chimescount - 1
    if chimescount<1 then
        me.enabled=false
        ScoreBonus.enabled=True
    end if
    chimestimer.interval=150
end Sub

sub newgame
  SPtimer.enabled=1
  bumper1.hashitevent=1
  bumper2.hashitevent=1
  player=1
  shoot1.state=1
  ebcount=0
  for i = 1 to 4:
    holeb(i).state=0
    score(i)=0
  next
  If b2son then
    for i = 1 to maxplayers
    Controller.B2SSetScorePlayer i, score(i)
    next
  End If
    eg=0
    rep(1)=0
    rep(2)=0
    rep(3)=0
    rep(4)=0
  for each light in bonuslights:light.state=0:next
    leftsp.state=0
    ExtraBall.state=0
  shootagain.state=lightstateoff
  for each light in bumperlights:light.state=1:next

  for each light in GIlights:light.state=1:next
  outerleft_prim.image = "outerlefton"
  outerright_prim.image = "outerrighton"
  for each light in numberlights:light.state=1:next
  GOReel.setvalue 0
  TiltReel.setvalue 0

    If b2son then
    Controller.B2SSetGameOver 35,0
    Controller.B2SSetTilt 33,0
    Controller.B2SSetMatch 34,0
  End If
    biptext.text="1"
  matchtxt.text=" "
  newball
end sub

sub newball
  for each light in numberlights:light.state=1:next
  for each light in starlights:light.state=0:next
  if ebcount=0 then
      resetDT
    else
      resetWT
  end if
  extraball.state=0
  ShootAgain.state=0
  bonusx2.state=0
  leftsp.state=0
  if ballinplay=balls then bonusx2.state=1
  TroughTimer.enabled=1
End Sub

Sub TroughTimer_timer
  EVAL("Shoot"&player).duration 0,300,1
    Drain.kick 60,45,0
  PlaySoundAt SoundFXDOF("drainkick",128,DOFPulse,DOFContactors), Drain
  me.enabled=0
end sub

Sub resetWT           'reset only white drop targets
  DTreset.uservalue=2
  DTreset.enabled=true
end sub

Sub resetDT           'reset all drop targets
  DTreset.uservalue=0
  DTreset.enabled=true
end sub

Sub DTreset_timer
    Select Case me.uservalue
        Case 1:
      playsoundat SoundFXDOF("bankreset",114,DOFPulse,DOFContactors), G4a   'green
      for each objekt in droptargetsGreen: objekt.isdropped=false: Next
        for each light in GDTlights:light.state=0:next
    Case 3:
      playsoundat SoundFXDOF("bankreset",111,DOFPulse,DOFContactors), W1a   'white
      for each objekt in droptargetsWhite: objekt.isdropped=false: Next
      for each light in WDTlights:light.state=0:next
      me.enabled=False
    End Select
    me.uservalue = me.uservalue + 1
End Sub

sub nextball
  if players=1 or player=players then
    player=1
    else
    player=player+1
  end if

  if player=1 then ballinplay=ballinplay+1
  if ballinplay>balls then
    playsound "endgame"
    eg=1
    ballreltimer.enabled=true
    else
    if tilt=true then
      bumper1.hashitevent=1
      bumper2.hashitevent=1
      for each light in GIlights:light.state=1:next
      tilt=false
      TiltReel.setvalue 0
      If b2son then
        Controller.B2SSetTilt 33,0
        Controller.B2ssetdata 1, 1
      End If
    end if
    ebcount=0
    for i = 1 to 4
      holeb(i).state=0
    next
    If b2son then Controller.B2ssetplayerup 30, player
    if player=1 then
      shoot1.state=1
      shoot2.state=0
      else
      shoot1.state=0
      shoot2.state=1
    end if
    if state=true and tilt=false then
      ballreltimer.enabled=true
    end if
    biptext.text=ballinplay
    If b2son then Controller.B2ssetballinplay 32, Ballinplay
  end if
End Sub

sub ballreltimer_timer
  if eg=1 then
    SPtimer.enabled=0
    hiscstate=0
    turnoff
      matchnum
    biptext.text=" "
    state=false
    GOReel.setvalue 1
    GOReel.timerenabled=1
    for i=1 to maxplayers
    EVAL("CanPlay"&i).state=0
    if score(i)>hisc then
      hisc=score(i)
      hiscstate=1
    end if
    EVAL("Shoot"&i).state=0
    next
    if hiscstate=1 then
      HighScoreEntryInit()
      HStimer.uservalue = 0
      HStimer.enabled=1
    end if
    UpdatePostIt
    savehs
    If b2son then
        Controller.B2SSetGameOver 35,1
        Controller.B2ssetballinplay 32, 0
      Controller.B2sStartAnimation "EOGame"
'     Controller.B2SSetScorePlayer 5, hisc
      Controller.B2ssetcanplay 31, 0
      Controller.B2ssetcanplay 30, 0
    End If

    ballreltimer.enabled=false
  else
  newball
    ballreltimer.enabled=false
  end if
end sub

Sub HStimer_timer
  PlaySoundAtVol SoundFXDOF("knock",130,DOFPulse,DOFKnocker), Plunger, 1
  DOF 131,DOFPulse
  HStimer.uservalue=HStimer.uservalue+1
  if HStimer.uservalue=3 then me.enabled=0
end sub

sub matchnum
  if matchnumb=0 then
    matchtxt.text="00"
    else
    matchtxt.text=matchnumb*10
  end if
  If b2son then Controller.B2SSetMatch 34,Matchnumb*10
  For i=1 to players
    if (matchnumb*10)=(score(i) mod 100) then
      addcredit
      PlaySoundAtVol SoundFXDOF("knock",130,DOFPulse,DOFKnocker), Plunger, 1
      DOF 131,DOFPulse
    end if
  next
end sub

'********** Bumpers

Sub Bumper1_Hit
   if tilt=false then
  bumper2.playhit
  PlaySoundAtVol SoundFXDOF("fx_bumper4",107,DOFPulse,DOFContactors), Bumper1, 1
  DOF 108,DOFPulse
  DOF 109,DOFPulse
  if bumperlight1.state=1 then
    addscore bumperlitscore
    else
    addscore bumperoffscore
  end if
   end if
End Sub

Sub Bumper2_Hit
   if tilt=false then
  bumper1.playhit
  PlaySoundAtVol SoundFXDOF("fx_bumper4",109,DOFPulse,DOFContactors), Bumper2, 1
  DOF 110,DOFPulse
  DOF 107,DOFPulse
  if bumperlight2.state=1 then
    addscore bumperlitscore
    else
    addscore bumperoffscore
  end if
   end if
End Sub

sub FlashBumpers
  For each light in BumperLights
    Light.duration 0, 500, 1
  next
  For each light in starlights
    if light.state=1 then light.duration 0, 500, 1
  next
end sub


'********** Dingwalls - animated - timer 50

sub dingwalla_hit
  if state=true and tilt=false then addscore 10
  SlingA.visible=0
  SlingA1.visible=1
  me.uservalue=1
  Me.timerenabled=1
end sub

sub dingwalla_timer                 'default 50 timer
  select case me.uservalue
    Case 1: SlingA1.visible=0: SlingA.visible=1
    case 2: SlingA.visible=0: SlingA2.visible=1
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

sub dingwallb_timer                 'default 50 timer
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

sub dingwallC_timer                 'default 50 timer
  select case me.uservalue
    Case 1: SlingC1.visible=0: SlingC.visible=1
    case 2: SlingC.visible=0: SlingC2.visible=1
    Case 3: SlingC2.visible=0: SlingC.visible=1: Me.timerenabled=0
  end Select
  me.uservalue=me.uservalue+1
end sub

sub dingwalld_hit
  if state=true and tilt=false then addscore 10
  Slingd.visible=0
  Slingd1.visible=1
  me.uservalue=1
  Me.timerenabled=1
end sub

sub dingwalld_timer                 'default 50 timer
  select case me.uservalue
    Case 1: Slingd1.visible=0: Slingd.visible=1
    case 2: Slingd.visible=0: Slingd2.visible=1
    Case 3: Slingd2.visible=0: Slingd.visible=1: Me.timerenabled=0
  end Select
  me.uservalue=me.uservalue+1
end sub

sub dingwalle_hit
  if state=true and tilt=false then addscore 10
  SlingE.visible=0
  Slinge1.visible=1
  me.uservalue=1
  Me.timerenabled=1
end sub

sub dingwalle_timer                 'default 50 timer
  select case me.uservalue
    Case 1: Slinge1.visible=0: SlingE.visible=1
    case 2: SlingE.visible=0: Slinge2.visible=1
    Case 3: Slinge2.visible=0: SlingE.visible=1: Me.timerenabled=0
  end Select
  me.uservalue=me.uservalue+1
end sub

sub dingwallf_hit
  if state=true and tilt=false then addscore 10
  SlingF.visible=0
  Slingf1.visible=1
  me.uservalue=1
  Me.timerenabled=1
end sub

sub dingwallf_timer                 'default 50 timer
  select case me.uservalue
    Case 1: Slingf1.visible=0: SlingF.visible=1
    case 2: SlingF.visible=0: Slingf2.visible=1
    Case 3: Slingf2.visible=0: SlingF.visible=1: Me.timerenabled=0
  end Select
  me.uservalue=me.uservalue+1
end sub

sub dingwallg_hit
  if state=true and tilt=false then addscore 10
  SlingG.visible=0
  Slingg1.visible=1
  me.uservalue=1
  Me.timerenabled=1
end sub

sub dingwallg_timer                 'default 50 timer
  select case me.uservalue
    Case 1: Slingg1.visible=0: SlingG.visible=1
    case 2: SlingG.visible=0: Slingg2.visible=1
    Case 3: Slingg2.visible=0: SlingG.visible=1: Me.timerenabled=0
  end Select
  me.uservalue=me.uservalue+1
end sub

'********* Targets

sub SpotTarget_hit
  spot5.state=0
  drop5.state=1
  PlaySoundAtVol SoundFXDOF("target",135,DOFPulse,DOFContactors),SpotTarget, 0.5
  if balls = 3 then
    addscore 3000
    else
    addscore 500
  end if
  if extraball.state=1 then
    if ebcount=0 then
      shootagain.state=lightstateon
      extraball.state=0
      ebcount=1
     end if
  end if

end sub

'******** Kick Holes

sub TopHole_hit
  addscore 1000
  holebonus=0
  for i = 1 to 4
    if holeb(i).state=1 then holebonus=holebonus+1
  next
  addscore 1000*holebonus
  me.uservalue=1
  me.timerenabled=1
end sub

Sub TopHole_timer
  select case me.uservalue
    case 4:
    TopHole.Kick 205, 12
    Pkickarm.rotz=15
    PlaySoundAtVol SoundFXDOF("HoleKick",132,DOFPulse,DOFContactors),TopHole, 0.25
    DOF 131, DOFPulse
    case 6:
    Pkickarm.rotz=0
    me.timerenabled=0
  end Select
  me.uservalue=me.uservalue+1
End Sub

'****** Wire Triggers

Sub Toplane1_Hit
   if tilt=false then
  DOF 115,DOFPulse
  flashbumpers
  Drop1.State  = 1
  Top1.state = 0
  if balls = 3 then
    addscore 3000
    else
    addscore 500
  end if
   end if
End Sub

Sub Toplane2_Hit
   if tilt=false then
  DOF 116,DOFPulse
  flashbumpers
  Drop2.State  = 1
  if balls = 3 then
    addscore 3000
    else
    addscore 500
  end if
  Top2.state = 0
   end if
End Sub

Sub Toplane3_Hit
   if tilt=false then
  DOF 117,DOFPulse
  flashbumpers
  Drop3.State  = 1
  if balls = 3 then
    addscore 3000
    else
    addscore 500
  end if
  Top3.state = 0
   end if
End Sub

Sub Toplane4_Hit
   if tilt=false then
  DOF 118,DOFPulse
  flashbumpers
  Drop4.State  = 1
  if balls = 3 then
    addscore 3000
    else
    addscore 500
  end if
  Top4.state = 0
   end if
End Sub

Sub LeftOutlane_Hit
   if tilt=false then
  DOF 124,DOFPulse
  flashbumpers
    if LeftSP.state=1 then
    PlaySoundAt SoundFXDOF("knock",130,DOFPulse,DOFKnocker), LeftOutlane
    DOF 131,DOFPulse
    addcredit
    end if
  Drop1.State  = 1
  if balls = 3 then
    addscore 3000
    else
    addscore 500
  end if
  Top1.state = 0
   end if
End Sub

Sub LeftInlane_Hit
   if tilt=false then
  DOF 125,DOFPulse
  flashbumpers
  Drop2.State  = 1
  if balls = 3 then
    addscore 3000
    else
    addscore 500
  end if
  Top2.state = 0
   end if
End Sub

Sub RightInlane_Hit
   if tilt=false then
  DOF 127,DOFPulse
  flashbumpers
  Drop3.State  = 1
  if balls = 3 then
    addscore 3000
    else
    addscore 500
  end if
  Top3.state = 0
   end if
End Sub

Sub RightOutlane_Hit
   if tilt=false then
  DOF 126,DOFPulse
  flashbumpers
    if LeftSP.state=1 then
    PlaySoundAt SoundFXDOF("knock",130,DOFPulse,DOFKnocker), RightOutlane
    DOF 131,DOFPulse
    addcredit
    end if
  Drop4.State  = 1
  if balls = 3 then
    addscore 3000
    else
    addscore 500
  end if
  Top4.state = 0
   end if
End Sub

'**** Star Rollovers

Sub Rollover1_Hit
  if tilt=false then
    DOF 119, DOFPulse
    if roll1.state=1 then
      addscore 1000
      else
      addscore 100
    end if
  end if
end sub

Sub Rollover2_Hit
  if tilt=false then
    DOF 120, DOFPulse
    if roll2.state=1 then
      addscore 1000
      else
      addscore 100
    end if
  end if
end sub

Sub Rollover3_Hit
  if tilt=false then
    DOF 121, DOFPulse
    if roll3.state=1 then
      addscore 1000
      else
      addscore 100
    end if
  end if
end sub

Sub Rollover4_Hit
  if tilt=false then
    DOF 122, DOFPulse
    if roll4.state=1 then
      addscore 1000
      else
      addscore 100
    end if
  end if
end sub

Sub Rollover5_Hit
  if tilt=false then
    DOF 123, DOFPulse
    if roll5.state=1 then
      addscore 1000
      else
      addscore 100
    end if
  end if
end sub


'**** Drop Targets

Sub CheckAwards
  awardcheck=0
  For j = 0 to 4
    if droptargets(j).isdropped then awardcheck=awardcheck+1
  next
  if awardcheck=5 then
    bonusx2.state=1
    resetWT
  end if
  awardcheck=0
  For j = 5 to 8
    if droptargets(j).isdropped then awardcheck=awardcheck+1
  next
  if awardcheck=4 and ebcount=0 then extraball.state=1
End Sub

Sub W1a_dropped
  DOF 112, DOFpulse
  if tilt=false then
    addscore 500
'   FlashBumpers
    If Top1.State = 0 then
      if balls =3 then
        addbonus(3)
        else
        addbonus(2)
      end if
      else
      addbonus(1)
    end if
    CheckAwards
  end if
End Sub

Sub W2a_dropped
  DOF 112, DOFpulse
  if tilt=false then
    addscore 500
    FlashBumpers
    If Top2.State = 0 then
      if balls =3 then
        addbonus(3)
        else
        addbonus(2)
      end if
      else
      addbonus(1)
    end if
    CheckAwards
  end if
End Sub

Sub W3a_dropped
  DOF 112, DOFpulse
  if tilt=false then
    addscore 500
    FlashBumpers
    If Top3.State = 0 then
      if balls =3 then
        addbonus(3)
        else
        addbonus(2)
      end if
      else
      addbonus(1)
    end if
    CheckAwards
  end if
End Sub

Sub W4a_dropped
  DOF 112, DOFpulse
  if tilt=false then
    addscore 500
    FlashBumpers
    If Top4.State = 0 then
      if balls =3 then
        addbonus(3)
        else
        addbonus(2)
      end if
      else
      addbonus(1)
    end if
    CheckAwards
  end if
End Sub

Sub W5a_dropped
  DOF 112, DOFpulse
  if tilt=false then
    addscore 500
    FlashBumpers
    If Spot5.State = 0 then
      if balls =3 then
        addbonus(3)
        else
        addbonus(2)
      end if
      else
      addbonus(1)
    end if
    CheckAwards
  end if
End Sub

Sub G1a_dropped
  DOF 113, DOFpulse
  if tilt=false then
    FlashBumpers
    addscore 500
    addbonus(1)
    CheckAwards
    Hole1.state =1
  end if
End Sub

Sub G2a_dropped
  DOF 113, DOFpulse
  if tilt=false then
    FlashBumpers
    addscore 500
    addbonus(1)
    CheckAwards
    Hole2.state =1
  end if
End Sub

Sub G3a_dropped
  DOF 113, DOFpulse
  if tilt=false then
    FlashBumpers
    addscore 500
    addbonus(1)
    Hole3.state =1
    CheckAwards
  end if
End Sub

Sub G4a_dropped
  DOF 113, DOFpulse
  if tilt=false then
    FlashBumpers
    addscore 500
    addbonus(1)
    Hole4.state =1
    CheckAwards
  end if
End Sub


'********** Scoring

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
        PlaySoundAtVol SoundFXDOF("bell100",142,DOFPulse,DOFChimes), chimesound, 1
      ElseIf points = 1000 Then
        PlaySoundAtVol SoundFXDOF("bell1000",143,DOFPulse,DOFChimes), chimesound, 1
    elseif Points = 100 Then
        PlaySoundAtVol SoundFXDOF("bell100",142,DOFPulse,DOFChimes), chimesound, 1
      Else
        PlaySoundAtVol SoundFXDOF("bell10",141,DOFPulse,DOFChimes), chimesound, 1
    End If
  checkreplay

end sub

sub checkreplay
  if score(player)=>99999 then
    if b2son then Controller.B2SSetScoreRollover 24 + player, 1
    EVAL("r100k"&player).setvalue INT(score(player)/100000)
  End if
    if score(player)=>replay1 and rep(player)=0 then
    PlaySoundAt SoundFXDOF("knock",130,DOFPulse,DOFKnocker), Plunger
    DOF 131,DOFPulse
    addcredit
      rep(player)=1
    end if
    if score(player)=>replay2 and rep(player)=1 then
    PlaySoundAt SoundFXDOF("knock",130,DOFPulse,DOFKnocker), Plunger
    DOF 131,DOFPulse
      addcredit
      rep(player)=2
    end if
    if score(player)=>replay3 and rep(player)=2 then
    PlaySoundAt SoundFXDOF("knock",130,DOFPulse,DOFKnocker), plunger
    DOF 131,DOFPulse
      addcredit
      rep(player)=3
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
  TiltReel.setvalue 1
    If B2SOn Then Controller.B2SSetTilt 33,1
    If B2SOn Then Controller.B2ssetdata 1, 0
  for each light in GIlights:light.state=0:next
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


Sub addbonus(count)
  if tilt=false and addbonustimer.enabled=false then
    addbonustimer.enabled=true
    bonus=bonus+count
    if bonus > 19 then bonus = 19
    BonusLights(bonus).state=1
    if bonus-count>0 then Bonuslights(bonus-count).state=0
    if bonus>10 then BonusLights(10).state=1
  end if
End sub

sub addbonustimer_timer()
  me.enabled=false
end sub


sub savehs
    savevalue "FireQueen", "credit", credit
    savevalue "FireQueen", "hiscore", hisc
    savevalue "FireQueen", "match", matchnumb
    savevalue "FireQueen", "score1", score(1)
    savevalue "FireQueen", "score2", score(2)
    savevalue "FireQueen", "score3", score(3)
    savevalue "FireQueen", "score4", score(4)
  savevalue "FireQueen", "replays", replays
  savevalue "FireQueen", "balls", balls
  savevalue "FireQueen", "hsa1", HSA1
  savevalue "FireQueen", "hsa2", HSA2
  savevalue "FireQueen", "hsa3", HSA3
  savevalue "FireQueen", "freeplay", freeplay
  savevalue "FireQueen", "ballshadows", ballshadows
  savevalue "FireQueen", "flippershadows", flippershadows
end sub

sub loadhs
    dim temp
  temp = LoadValue("FireQueen", "credit")
    If (temp <> "") then credit = CDbl(temp)
    temp = LoadValue("FireQueen", "hiscore")
    If (temp <> "") then hisc = CDbl(temp)
    temp = LoadValue("FireQueen", "match")
    If (temp <> "") then matchnumb = CDbl(temp)
    temp = LoadValue("FireQueen", "score1")
    If (temp <> "") then score(1) = CDbl(temp)
    temp = LoadValue("FireQueen", "score2")
    If (temp <> "") then score(2) = CDbl(temp)
    temp = LoadValue("FireQueen", "score3")
    If (temp <> "") then score(3) = CDbl(temp)
    temp = LoadValue("FireQueen", "score4")
    If (temp <> "") then score(4) = CDbl(temp)
    temp = LoadValue("FireQueen", "replays")
    If (temp <> "") then replays = CDbl(temp)
    temp = LoadValue("FireQueen", "balls")
    If (temp <> "") then balls = CDbl(temp)
    temp = LoadValue("FireQueen", "hsa1")
    If (temp <> "") then HSA1 = CDbl(temp)
    temp = LoadValue("FireQueen", "hsa2")
    If (temp <> "") then HSA2 = CDbl(temp)
    temp = LoadValue("FireQueen", "hsa3")
    If (temp <> "") then HSA3 = CDbl(temp)
    temp = LoadValue("FireQueen", "freeplay")
    If (temp <> "") then freeplay = CDbl(temp)
    temp = LoadValue("FireQueen", "ballshadows")
    If (temp <> "") then ballshadows = CDbl(temp)
    temp = LoadValue("FireQueen", "flippershadows")
    If (temp <> "") then flippershadows = CDbl(temp)
end sub

Sub FireQueen_Exit()
  turnoff
  Savehs
  If B2SOn Then Controller.stop
End Sub

' *******************************************************************************************************
' Positional Sound Playback Functions by DJRobX
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

'Set position as bumperX and Vol manually.

Sub PlaySoundAtBumperVol(sound, tableobj, Vol)
  PlaySound sound, 1, Vol, AudioPan(tableobj), 0,0,1, 1, AudioFade(tableobj)
End Sub

'*********************************************************************
'                     Supporting Ball & Sound Functions
'*********************************************************************

Function RndNum(min, max)
    RndNum = Int(Rnd() * (max-min + 1) ) + min ' Sets a random number between min and max
End Function

Function AudioFade(tableobj) ' Fades between front and back of the table (for surround systems or 2x2 speakers, etc), depending on the Y position on the table. "FireQueen" is the name of the table
  Dim tmp
  tmp = tableobj.y * 2 / FireQueen.height-1
  If tmp > 0 Then
    AudioFade = Csng(tmp ^10)
  Else
    AudioFade = Csng(-((- tmp) ^10) )
  End If
End Function

Function AudioPan(tableobj) ' Calculates the pan for a tableobj based on the X position on the table. "FireQueen" is the name of the table
  Dim tmp
  tmp = tableobj.x * 2 / FireQueen.width-1
  If tmp > 0 Then
    AudioPan = Csng(tmp ^10)
  Else
    AudioPan = Csng(-((- tmp) ^10) )
  End If
End Function

Function Pan(ball) ' Calculates the pan for a ball based on the X position on the table. "FireQueen" is the name of the table
    Dim tmp
    tmp = ball.x * 2 / FireQueen.width-1
    If tmp > 0 Then
        Pan = Csng(tmp ^10)
    Else
        Pan = Csng(-((- tmp) ^10) )
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

'*****************************************
'     BALL SHADOW
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
        If BOT(b).X < FireQueen.Width/2 Then
            BallShadow(b).X = ((BOT(b).X) - (Ballsize/16) + ((BOT(b).X - (FireQueen.Width/2))/17))' + 13
        Else
            BallShadow(b).X = ((BOT(b).X) + (Ballsize/16) + ((BOT(b).X - (FireQueen.Width/2))/17))' - 13
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

'==========================================================================================================================================
'============================================================= START OF HIGH SCORES ROUTINES =============================================================
'==========================================================================================================================================
'
'ADD LINE TO TABLE_KEYDOWN SUB WITH THE FOLLOWING:    If HSEnterMode Then HighScoreProcessKey(keycode) AFTER THE STARTGAME ENTRY
'ADD: And Not HSEnterMode=true TO IF KEYCODE=STARTGAMEKEY
'TO SHOW THE SCORE ON POST-IT ADD LINE AT RELEVENT LOCATION THAT HAS:  UpdatePostIt
'TO INITIATE ADDING INITIALS ADD LINE AT RELEVENT LOCATION THAT HAS:  HighScoreEntryInit()
'ADD THE FOLLOWING LINES TO TABLE_INIT TO SETUP POSTIT
' if HSA1="" then HSA1=25
' if HSA2="" then HSA2=25
' if HSA3="" then HSA3=25
' UpdatePostIt
'ADD HSA1, HSA2 AND HSA3 TO SAVE AND LOAD VALUES FOR TABLE
'ADD A TIMER NAMED HighScoreFlashTimer WITH INTERVAL 100 TO TABLE
'SET HSSSCOREX BELOW TO WHATEVER VARIABLE YOU USE FOR HIGH SCORE.
'ADD OBJECTS TO PLAYFIELD (EASIEST TO JUST COPY FROM THIS TABLE)
'IMPORT POST-IT IMAGES


Dim HSA1, HSA2, HSA3
Dim HSEnterMode, hsLetterFlash, hsEnteredDigits(3), hsCurrentDigit, hsCurrentLetter
Dim HSArray
Dim HSScoreM,HSScore100k, HSScore10k, HSScoreK, HSScore100, HSScore10, HSScore1, HSScorex 'Define 6 different score values for each reel to use
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
' if showhisc=1 and showhiscnames=1 then
'   for each objekt in hiscname:objekt.visible=1:next
    HSName1.image = ImgFromCode(HSA1, 1)
    HSName2.image = ImgFromCode(HSA2, 2)
    HSName3.image = ImgFromCode(HSA3, 3)
'   else
'   for each objekt in hiscname:objekt.visible=0:next
' end if
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
'       EMReelHSName1.SetValue HSA1:EMReelHSName2.SetValue HSA2
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



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
dim rst
dim ballinplay
dim match(10)
dim tilt
dim tiltsens
dim rep(4)
dim plm(4)
dim matchnumb
dim bell
dim points
dim tempscore
dim replay1
dim replay2
dim replay3
dim replay4
dim hisc
dim scn
dim scn1
dim scn2
dim scn3
dim wheelvalue
dim wv
dim holevalue
dim hv
dim wn
dim tt
dim i
dim colorvalue
dim cv
dim qs

' Thalamus 2020 June : Improved directional sounds
' !! NOTE : Table not verified yet !!

' Options
' Volume devided by - lower gets higher sound

Const VolDiv = 2000    ' Lower number, louder ballrolling/collition sound
Const VolCol = 10      ' Ball collition divider ( voldiv/volcol )

Dim B2SOn
Dim controller
ExecuteGlobal GetTextFile("core.vbs")


Sub Table1_Init
  If ShowDT=false Then
    for each obj in DesktopItems
      obj.visible=False
    Next
  end If
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
    replay1=1400
    replay2=1800
    replay3=2200
    replay4=2600
    loadhs
    if hisc="" then hisc=1000
    hisctxt.text=hisc
    if wv="" then wv=0
    if credit="" then credit=0
    credittxt.text=credit
    if matchnumb="" then matchnumb=int(rnd(1)*9)
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
  B2SOn=True
  if B2SOn then
    Set Controller = CreateObject("B2S.Server")
    Controller.B2SName = "Gay90s_1970"
    Controller.Run()
    If Err Then MsgBox "Can't Load B2S.Server."
  end if
  If B2SOn Then

    if matchnumb=0 then
      Controller.B2SSetMatch 10
    else
      Controller.B2SSetMatch matchnumb
    end if
    Controller.B2SSetScoreRolloverPlayer1 0
    Controller.B2SSetScoreRolloverPlayer2 0
    Controller.B2SSetScoreRolloverPlayer3 0
    Controller.B2SSetScoreRolloverPlayer4 0
    Controller.B2SSetShootAgain 0
    Controller.B2SSetTilt 1
    Controller.B2SSetCredits Credit
    Controller.B2SSetGameOver 1
    Controller.B2SSetData 81,0
    Controller.B2SSetData 82,0
    Controller.B2SSetData 83,0
    Controller.B2SSetData 84,0
  End If
  for i=1 to 4
    If B2SOn Then
      Controller.B2SSetScorePlayer i, 0
    End If
  next
End Sub


Sub Table1_KeyDown(ByVal keycode)

  If keycode = PlungerKey Then
  Plunger.PullBack
  End If

  if keycode = 6 then
  playsound "coin3"
  coindelay.enabled=true
  If B2SOn Then
    Controller.B2SSetCredits Credit
  end if
  end if

  if keycode = 5 then
  playsound "coin3"
  coindelay1.enabled=true
  If B2SOn Then
    Controller.B2SSetCredits Credit
  end if
  end if

  if keycode = 2 and credit>0 and state=false and playno=0 then
  credit=credit-1
  qs=qs-1
  If qs<0 then qs=0
  credittxt.text=credit
    eg=0
    playno=1
    playno1.state=1
    currpl=1
    pno.setvalue(playno)
    play(currpl).state=1
    playsound "click"
    playsound "initialize"
    rst=0
    ballinplay=1
    resettimer.enabled=true
    If B2SOn Then
      Controller.B2SSetTilt 0
      Controller.B2SSetGameOver 0
      Controller.B2SSetMatch 0
      Controller.B2SSetCredits Credit
      Controller.B2SSetScore 3,HighScore
      Controller.B2SSetCanPlay 1
      Controller.B2SSetPlayerUp 1
      Controller.B2SSetData 81,1
      Controller.B2SSetData 82,0
      Controller.B2SSetData 83,0
      Controller.B2SSetData 84,0
      Controller.B2SSetBallInPlay BallInPlay
      Controller.B2SSetScoreRolloverPlayer1 0

    End If
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
    Controller.B2SSetCredits Credit
    Controller.B2SSetCanPlay playno
  end if
    end if

    if state=true and tilt=false then

  If keycode = LeftFlipperKey Then
  LeftFlipper.RotateToEnd
  PlaySoundAtVol "FlipperUp", LeftFlipper, 1
  End If

  If keycode = RightFlipperKey Then
  RightFlipper.RotateToEnd
  PlaySoundAtVol "FlipperUp", RightFlipper, 1
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
  End If

End Sub

Sub Table1_KeyUp(ByVal keycode)

  If keycode = PlungerKey Then
  Plunger.Fire
  playsoundAtVol "plungerrelease", Plunger, 1
  End If

  If keycode = LeftFlipperKey Then
  LeftFlipper.RotateToStart
  stopsound "buzz"
  if state=true and tilt=false then PlaySoundAtVol "FlipperDown", LeftFlipper, 1
  End If

  If keycode = RightFlipperKey Then
  RightFlipper.RotateToStart
  stopsound "buzz"
  if state=true and tilt=false then PlaySoundAtVol "FlipperDown", RightFlipper, 1
  End If

End Sub

sub flippertimer_timer()
  LFlip.RotY = LeftFlipper.CurrentAngle+90
  RFlip.RotY = RightFlipper.CurrentAngle+90
end sub


sub newgame
  wheelchange.enabled=False
  BumperLight4.State=1
  Light1.State=1
  Light3.State=1
  Light7.State=0
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
  ballinplay=1
  nb.CreateBall
  nb.kick 90,6
  If B2SOn then Controller.B2SSetBallInPlay 1
    wheelchange.enabled=True
end sub

sub resettimer_timer
    wheelchange.enabled=False
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
  end if
    if rst=14 then
    playsound "kickerkick"
    end if
    if rst=18 then
    newgame
    resettimer.enabled=false
    wheelchange.enabled=True
    end if
end sub

sub coindelay_timer
  playsound "click"
  credit=credit+5
  credittxt.text=credit
    coindelay.enabled=false
  If B2SOn Then
    Controller.B2SSetCredits Credit
  end if
end sub

sub coindelay1_timer
  playsound "click"
  credit=credit+1
  credittxt.text=credit
    coindelay1.enabled=false
  If B2SOn Then
    Controller.B2SSetCredits Credit
  end if
end sub

sub addscore(points)
  if tilt=false then
  bell=0
  scn=0
  scn1=0
  if points = 1 or points = 10 or points=100 then scn=1

  if points = 300 then
  reel(currpl).addvalue(300)
  bell=3
  scn=3
  end if

  if points = 200 then
  reel(currpl).addvalue(200)
  bell=3
  scn=2
  end if

  if points = 100 then
  reel(currpl).addvalue(100)
  bell=3
  end if

  if points = 50 then
  reel(currpl).addvalue(50)
  bell=2
  scn=5
    end if

  if points = 40 then
  reel(currpl).addvalue(40)
  bell=2
  scn=4
    end if

  if points = 30 then
  reel(currpl).addvalue(30)
  bell=2
  scn=3
    end if

  if points = 20 then
  reel(currpl).addvalue(20)
  bell=2
  scn=2
    end if

  if points = 10 then
  reel(currpl).addvalue(10)
  bell=2
    end if

  if points = 1 then
  reel(currpl).addvalue(1)
  bell=1
    end if

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
  playsound "knocke"
  credittxt.text=credit
  If B2SOn Then
    Controller.B2SSetCredits Credit
  end if
  qs=qs+2
  If qs>32 then qs=32
  rep(currpl)=1
  playsound "click"
  end if

  if score(currpl)=>replay2 and rep(currpl)=1 then
  credit=credit+1
  playsound "knocke"
  credittxt.text=credit
  If B2SOn Then
    Controller.B2SSetCredits Credit
  end if
  qs=qs+2
  If qs>32 then qs=32
  rep(currpl)=2
  playsound "click"
  end if

  if score(currpl)=>replay3 and rep(currpl)=2 then
  credit=credit+1
  playsound "knocke"
  credittxt.text=credit
  If B2SOn Then
    Controller.B2SSetCredits Credit
  end if
  qs=qs+2
  If qs>32 then qs=32
  rep(currpl)=3
  playsound "click"

  if score(currpl)=>replay4 and rep(currpl)=3 then
  credit=credit+1
  playsound "knocke"
  credittxt.text=credit
  If B2SOn Then
    Controller.B2SSetCredits Credit
  end if
  qs=qs+2
  If qs>32 then qs=32
  rep(currpl)=4
  playsound "click"
  end if
  end if
    end if
end sub

sub scntimer_timer
  scn1=scn1 + 1
  matchnumb=matchnumb+1
  wv=wv+1
  if wv>9 then wv=0
  if matchnumb=10 then matchnumb=0
  if bell=3 then playsound "bell1000",0, 0.15, 0, 0
  if bell=2 then playsound "bell100",0, 0.15, 0, 0
  if bell=1 then playsound "bell10",0, 0.30, 0, 0
  if scn1=scn then
  bell=0
    wheelcheck
    scntimer.enabled=false
    end if
end sub

sub wheelcheck
  If wv=0 or wv=1 then
  BumperLight1.state=1
  Light8.state=1
  BumperLight2.state=0
  BumperLight3.state=0
  Light9.state=0
  Light1.State=1
  Light3.State=1
  Light2.State=0
  Light4.State=0
  Light5.State=1
  Light6.State=0
    quickstep
    End If

    If wv=2 or wv=3 then
    BumperLight1.state=1
  Light8.state=1
  BumperLight2.state=1
  BumperLight3.state=0
  Light9.state=0
  Light1.State=0
  Light3.State=0
  Light2.State=1
  Light4.State=1
  Light5.State=0
  Light6.State=1
    quickstep
    End If

    If wv=4 or wv=5 then
  BumperLight1.state=1
  Light8.state=1
  BumperLight2.state=1
  BumperLight3.state=1
  Light9.state=1
  Light1.State=1
  Light3.State=1
  Light2.State=0
  Light4.State=0
  Light5.State=1
  Light6.State=0
  quickstep
    End If

    If wv=6 or wv=7 then
  BumperLight1.state=0
  Light8.state=0
  BumperLight2.state=1
  BumperLight3.state=1
  Light9.state=1
  Light1.State=0
  Light3.State=0
  Light2.State=1
  Light4.State=1
  Light5.State=0
  Light6.State=1
    quickstep
    End If

    If wv=8 or wv=9 then
  BumperLight1.state=0
  Light8.state=0
  BumperLight2.state=0
  BumperLight3.state=1
  Light9.state=1
  Light1.State=1
  Light3.State=1
  Light2.State=0
  Light4.State=0
  Light5.State=1
  Light6.State=0
    quickstep
    End If
end sub


Sub quickstep
    If wv=1 and qs>=8 then
    BumperLight1.state=0
    Light8.state=0
    BumperLight2.state=0
    BumperLight3.state=0
    Light9.state=0
    End If

    If wv=3 and qs>=12 then
    BumperLight1.state=0
    Light8.state=0
    BumperLight2.state=0
    BumperLight3.state=0
    Light9.state=0
    End If

    If (wv=4 or wv=5) and qs>=16 then
    BumperLight1.state=0
    Light8.state=0
    BumperLight2.state=0
    BumperLight3.state=0
    Light9.state=0
    End If

    If wv=7 and qs>=20 then
    BumperLight1.state=0
    Light8.state=0
    BumperLight2.state=0
    BumperLight3.state=0
    Light9.state=0
    End If

    If wv=9 and qs>=24 then
    BumperLight1.state=0
    Light8.state=0
    BumperLight2.state=0
    BumperLight3.state=0
    Light9.state=0
    End If

End Sub


Sub motorcheck
    If mv=0 or mv=1 then
    Light1.state=1
    Light2.state=0
    End If

    If mv=2 then
    Light2.state=1
    Light1.state=0
    End If
End Sub


sub wheelchange_timer
    wn=wn+1
    if tilt=false then
    If points=1 or points=10 or point=100 then wv=wv+1
    if wv>9 then wv=0
    wheelcheck
    end if
end sub


sub bumper1_hit
    if tilt=false then playsoundAtVol "jet1", ActiveBall, 1
    if BumperLight1.state=1 then addscore 10 else addscore 1
end sub

sub bumper2_hit
    if tilt=false then playsoundAtVol "jet1", ActiveBall, 1
    if BumperLight2.state=1 then addscore 10 else addscore 1
end sub

sub bumper3_hit
    if tilt=false then playsoundAtVol "jet1", ActiveBall, 1
    if BumperLight3.state=1 then addscore 10 else addscore 1
end sub

sub bumper4_hit
    if tilt=false then playsoundAtVol "jet1", ActiveBall, 1
    if BumperLight4.state=1 then addscore 10 else addscore 1
end sub

sub matchnum
  match(matchnumb).text=matchnumb
  If B2SOn Then

    if matchnumb=0 then
      Controller.B2SSetMatch 10
    else
      Controller.B2SSetMatch matchnumb
    end if
  end if
  for i=0 to playno
  if matchnumb=(score(i) mod 10) then
  credit=credit+1
  If B2SOn Then
    Controller.B2SSetCredits Credit
  end if
  playsound "knocke"
  credittxt.text= credit
  qs=qs+2
  If qs>32 then qs=32
    playsound "click"
    end if
    next
end sub

sub turnoff
    tiltseq.play seqalloff
    lsling.isdropped=true
    rsling.isdropped=true
    wheelchange.enabled=False
end sub

Sub Drain_Hit()
  playsoundAtVol "drainshorter", Drain, 1
  Drain.DestroyBall
    nextball
End Sub


sub nextball
  BumperLight4.State=1
  Light7.State=0
    LightReset

  if tilt=true then
  tilt=false
  tilttext.text=" "
  If B2SOn Then Controller.B2SSetTilt 0:Controller.B2SSetShootAgain 0
  tiltseq.stopplay
  pno.setvalue(playno)
  end if

  currpl=currpl+1
  If B2SOn Then
    Controller.B2SSetData 81,0
    Controller.B2SSetData 82,0
    Controller.B2SSetData 83,0
    Controller.B2SSetData 84,0
    Controller.B2SSetData (80+currpl),1
  end If
  if currpl>playno then
  ballinplay=ballinplay+1
  if ballinplay>5 then
  playsound "motorleer"
  If B2SOn then Controller.B2SSetBallInPlay 0
  eg=1
  ballreltimer.enabled=true
  else
  'if state=true and tilt=false then
  play(currpl-1).state=0
  currpl=1
  play(currpl).state=1
  If B2SOn Then
    Controller.B2SSetData 81,0
    Controller.B2SSetData 82,0
    Controller.B2SSetData 83,0
    Controller.B2SSetData 84,0
    Controller.B2SSetData (80+currpl),1
  end If
  playsound "kickerkick"
  ballreltimer.enabled=true
  end if
  If B2SOn then Controller.B2SSetBallInPlay ballinplay
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
  'end if
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
  end If
  playsound "kickerkick"
  ballreltimer.enabled=true
  end if
  end if
end Sub

Sub LightReset
  Light11.State=0
  Light12.State=0
  Light13.State=0
  Light14.State=0
  Light15.State=0
  Light16.State=0
  Light17.State=0
  Light18.State=0
  Light19.State=0
  Light20.State=0
  Light21.State=0
  Light22.State=0
  Light23.State=0
  Light24.State=0
  Light25.State=0
End Sub

sub ballreltimer_timer
  if eg=1 then
  matchnum
  bip3.text=" "
  bip5.text=" "
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
  gamov.text="SPIEL AUS"
  If B2SOn Then
    Controller.B2SSetGameOver 1
    Controller.B2SSetData 81,0
    Controller.B2SSetData 82,0
    Controller.B2SSetData 83,0
    Controller.B2SSetData 84,0
  end if
  wheelchange.enabled=False
  savehs
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
  end if
  for i=1 to playno
  if matchnumb=(score(i) mod 10) then
  credit=credit+1
  If B2SOn Then
    Controller.B2SSetCredits Credit
  end if
  qs=qs+2
  If qs>32 then qs=32
    playsound "knocke"
    credittxt.text= credit
    playsound "click"
    wheelchange.enabled=False
    end if
    next
end sub

Sub CheckTilt
  If Tilttimer.Enabled = True Then
  TiltSens = TiltSens + 1
  if TiltSens = 3 Then
  Tilt = True
  tilttext.text="GEKIPPT"
  If B2SOn Then Controller.B2SSetTilt 1
  tiltsens = 0
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
    tiltseq.play seqalloff
    pno.setvalue(5)
end sub



'**********Sling Shot Animations
' Rstep and Lstep  are the variables that increment the animation
'****************
Dim RStep, Lstep

Sub RightSlingShot_Slingshot
    If Light9.State=1 then Addscore 10 else Addscore 1
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
    If Light8.State=1 then Addscore 10 else Addscore 1
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
' Positional Sound Playback Functions by DJRobX, Rothbauerw, Thalamus and Herweh
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

Function VolMulti(ball,Multiplier) ' Calculates the Volume of the sound based on the ball speed
  VolMulti = Csng(BallVel(ball) ^2 / 150 ) * Multiplier
End Function

Function DVolMulti(ball,Multiplier) ' Calculates the Volume of the sound based on the ball speed
  DVolMulti = Csng(BallVel(ball) ^2 / 150 ) * Multiplier
  debug.print DVolMulti
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
  Set ScoreFile=FileObj.CreateTextFile(UserDirectory & "Aces_High.txt",True)
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
  If Not FileObj.FileExists(UserDirectory & "Aces_High.txt") then
  Exit Sub
  End if
  Set ScoreFile=FileObj.GetFile(UserDirectory & "Aces_High.txt")
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
  Credit = CDbl(temp1)
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


Sub LeftSlingShot1_Slingshot()
  Addscore 1
End Sub

Sub LeftSlingShot2_Slingshot()
  Addscore 1
End Sub

Sub LeftSlingShot3_Slingshot()
  Addscore 1
End Sub

Sub RightSlingShot1_Slingshot()
  Addscore 1
End Sub

Sub RightSlingShot2_Slingshot()
  Addscore 1
End Sub

Sub RightSlingShot3_Slingshot()
  Addscore 1
End Sub

Sub Bumper5_Hit()
  If Light11.State=1 and Light14.State=1 and Light17.State=1 and Light20.State=1 and Light23.State=1 then addscore 300: LightReset: Light7.State=1: PlaysoundAtVol "Motorshort", ActiveBall, 1
  If Light14.State=1 and Light15.State=1 and Light16.State=1 then Addscore 200: LightReset: PlaysoundAtVol "Motorshort", ActiveBall, 1
  If Light15.State=1 then Light16.State=1
  If Light14.State=1 then Light15.State=1
  Light14.State=1
  Addscore 10
End Sub

Sub Bumper6_Hit()
  If Light11.State=1 and Light14.State=1 and Light17.State=1 and Light20.State=1 and Light23.State=1 then addscore 300: LightReset: Light7.State=1: PlaysoundAtVol "Motorshort", ActiveBall, 1
  If Light11.State=1 and Light12.State=1 and Light13.State=1 then Addscore 300: LightReset: PlaysoundAtVol "Motorshort", ActiveBall, 1
  If Light12.State=1 then Light13.State=1
  If Light11.State=1 then Light12.State=1
  Light11.State=1
  Addscore 10
End Sub

Sub Bumper9_Hit()
  If Light11.State=1 and Light14.State=1 and Light17.State=1 and Light20.State=1 and Light23.State=1 then addscore 300: LightReset: Light7.State=1: PlaysoundAtVol "Motorshort", ActiveBall, 1
  If Light20.State=1 and Light21.State=1 and Light22.State=1 then Addscore 100: LightReset: PlaysoundAtVol "Motorshort", ActiveBall, 1
  If Light21.State=1 then Light22.State=1
  If Light20.State=1 then Light21.State=1
  Light20.State=1
  Addscore 10
End Sub

Sub Bumper7_Hit()
  If Light11.State=1 and Light14.State=1 and Light17.State=1 and Light20.State=1 and Light23.State=1 then addscore 300: LightReset: Light7.State=1: PlaysoundAtVol "Motorshort", ActiveBall, 1
  If Light17.State=1 and Light18.State=1 and Light19.State=1 then Addscore 100: LightReset: PlaysoundAtVol "Motorshort", ActiveBall, 1
  If Light18.State=1 then Light19.State=1
  If Light17.State=1 then Light18.State=1
  Light17.State=1
  Addscore 10
End Sub

Sub Bumper8_Hit()
  If Light23.State=1 and Light24.State=1 and Light25.State=1 then addscore 50: LightReset
  If Light11.State=1 and Light14.State=1 and Light17.State=1 and Light20.State=1 and Light23.State=1 then addscore 300: LightReset: Light7.State=1: PlaysoundAtVol "Motorshort", ActiveBall, 1
  If Light24.State=1 then Light25.State=1
  If Light23.State=1 then Light24.State=1
  Light23.State=1
  Addscore 10
End Sub

Sub Trigger1_Hit()
  If Light1.State=1 then Addscore 50 else Addscore 10
End Sub

Sub Trigger2_Hit()
  If Light2.State=1 then Addscore 50 else Addscore 10
End Sub

Sub Trigger3_Hit()
  If Light3.State=1 then Addscore 50 else Addscore 10
End Sub

Sub Trigger4_Hit()
  If Light4.State=1 then Addscore 50 else Addscore 10
End Sub

Sub Trigger5_Hit()
  If Light5.State=1 then Addscore 100 else Addscore 10
End Sub

Sub Trigger6_Hit()
  If Light6.State=1 then Addscore 100 else Addscore 10
End Sub

Sub Trigger7_Hit()
  Addscore 50
  If Light7.State=1 then
  credit=credit+1
  Playsound "Knocke"
  qs=qs+2
  If qs>32 then qs=32
    End If
End Sub

Sub Trigger8_Hit()
  Addscore 50
End Sub

Sub Trigger9_Hit()
  If Light11.State=1 and Light14.State=1 and Light17.State=1 and Light20.State=1 and Light23.State=1 then addscore 300: Light7.State=1: LightReset: PlaysoundAtVol "Motorshort", ActiveBall, 1
  If Light17.State=1 and Light18.State=1 and Light19.State=1 then Addscore 100: LightReset: PlaysoundAtVol "Motorshort", ActiveBall, 1
  If Light18.State=1 then Light19.State=1
  If Light17.State=1 then Light18.State=1
  Light17.State=1
  Addscore 10
End Sub

Sub Trigger10_Hit()
  If Light11.State=1 and Light14.State=1 and Light17.State=1 and Light20.State=1 and Light23.State=1 then addscore 300: Light7.State=1: LightReset: PlaysoundAtVol "Motorshort", ActiveBall, 1
  If Light20.State=1 and Light21.State=1 and Light22.State=1 then Addscore 100: LightReset: PlaysoundAtVol "Motorshort", ActiveBall, 1
  If Light21.State=1 then Light22.State=1
  If Light20.State=1 then Light21.State=1
  Light20.State=1
  Addscore 10
End Sub


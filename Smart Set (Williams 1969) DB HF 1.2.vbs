'*********************************************************************************
' _______  __   __  _______  ______    _______        _______  _______  _______
'|       ||  |_|  ||   _   ||    _ |  |       |      |       ||       ||       |
'|  _____||       ||  |_|  ||   | ||  |_     _|      |  _____||    ___||_     _|
'| |_____ |       ||       ||   |_||_   |   |        | |_____ |   |___   |   |
'|_____  ||       ||       ||    __  |  |   |        |_____  ||    ___|  |   |
' _____| || ||_|| ||   _   ||   |  | |  |   |         _____| ||   |___   |   |
'|_______||_|   |_||__| |__||___|  |_|  |___|        |_______||_______|  |___|
'
'                                  Wiliams (1969)
'*********************************************************************************
'Credits:
'Kees (original table)
'DryBonz (PF Redraw)
'Hauntfreaks (Plastics redraw/primitives)
'Loserman76 (B2S & coding updates)

' Thalamus 2018-07-24
' Added/Updated "Positional Sound Playback Functions" and "Supporting Ball & Sound Functions"
' Thalamus 2018-11-01 : Improved directional sounds
' !! NOTE : Table not verified yet !!

Randomize

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
Const VolPo     = 1    ' Rubber posts volume.
Const VolPi     = 1    ' Rubber pins volume.
Const VolTarg   = 1    ' Targets volume.
Const VolSpin   = 1.5  ' Spinners volume.
Const VolFlip   = 1    ' Flipper volume.


On Error Resume Next
ExecuteGlobal GetTextFile("controller.vbs")
If Err Then MsgBox "You need the controller.vbs in order to run this table, available in the vp10 package"
On Error Goto 0

Const cGameName = "smart_set_1969"

dim score(4)
dim truesc(4)
dim reel(5)
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
dim wheelvalue
dim wv
dim holevalue
dim hv
dim wn
dim tt
dim i
dim colorvalue
dim cv
dim quickstep
dim qs

ExecuteGlobal GetTextFile("core.vbs")

Sub Table1_Init
  LoadEM
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
    replay1=4300
    replay2=5400
    replay3=6500
    replay4=7600
    loadhs
    if hisc="" then hisc=1000
    reel5.setvalue(hisc)
    if wv="" then wv=0
    if credit="" then
    credit=0
  Else
    DOF 128, DOFOn
  End If
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
  playsoundAtVol "coin3", drain, 1
  coindelay.enabled=true
  If B2SOn Then
    Controller.B2SSetCredits Credit
  end if
  end if

  if keycode = 5 then
  playsoundAtVol "coin3", drain, 1
  coindelay1.enabled=true
  If B2SOn Then
    Controller.B2SSetCredits Credit
  end if
  end if

  if keycode = 2 and credit>0 and state=false and playno=0 then
  credit=credit-1
  If credit < 1 Then DOF 128, DOFOff
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
  If credit < 1 Then DOF 128, DOFOn
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
  PlaySoundAtVol SoundFXDOF("FlipperUp",101,DOFOn,DOFFlippers), LeftFlipper, VolFlip
  End If

  If keycode = RightFlipperKey Then
  RightFlipper.RotateToEnd
  PlaySoundAtVol SoundFXDOF("FlipperUp",102,DOFOn,DOFFlippers), RightFlipper, VolFlip
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

  End If

End Sub

Sub Table1_KeyUp(ByVal keycode)

  If keycode = PlungerKey Then
  Plunger.Fire
  playsoundAtVol "plungerrelease", plunger, 1
  End If

  If keycode = LeftFlipperKey Then
  LeftFlipper.RotateToStart
  stopsound "buzz"
  if state=true and tilt=false then PlaySoundAtVol SoundFXDOF("FlipperDown",101,DOFOff,DOFFlippers), LeftFlipper, VolFlip
  End If

  If keycode = RightFlipperKey Then
  RightFlipper.RotateToStart
  stopsound "buzz"
  if state=true and tilt=false then PlaySoundAtVol SoundFXDOF("FlipperDown",102,DOFOff,DOFFlippers), RightFlipper, VolFlip
  End If

End Sub

sub newgame
  wheelchange.enabled=False
  BumperLight1.State=0
  BumperLight2.State=0
  BumperLight3.State=0
  BumperLight4.State=0
  Post.IsDropped=True
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
  nb.kick 135,6
  DOF 123, DOFPulse
    wheelchange.enabled=True
  If B2SOn then Controller.B2SSetBallInPlay 1
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
  DOF 128, DOFOn
  credittxt.text=credit
    coindelay.enabled=false
  If B2SOn Then
    Controller.B2SSetCredits Credit
  end if
end sub

sub coindelay1_timer
  playsound "click"
  credit=credit+1
  DOF 128, DOFOn
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

  if points = 500 then
  reel(currpl).addvalue(500)
  bell=3
  scn=5
    end if

  if points = 400 then
  reel(currpl).addvalue(400)
  bell=3
  scn=4
    end if

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
  if score(currpl)>99999 then
  score(currpl)=score(currpl)-100000
  rep(currpl)=0
  end if

  if score(currpl)=>replay1 and rep(currpl)=0 then
  credit=credit+1
  DOF 128, DOFOn
    playsound SoundFXDOF("knocke",126,DOFPulse,DOFKnocker)
  DOF 125, DOFPulse
  credittxt.text=credit
  If B2SOn Then
    Controller.B2SSetCredits Credit
  end if
  rep(currpl)=1
  playsound "click"
  end if

  if score(currpl)=>replay2 and rep(currpl)=1 then
  credit=credit+1
  DOF 128, DOFOn
    playsound SoundFXDOF("knocke",126,DOFPulse,DOFKnocker)
  DOF 125, DOFPulse
  credittxt.text=credit
  If B2SOn Then
    Controller.B2SSetCredits Credit
  end if
  rep(currpl)=2
  playsound "click"
  end if

  if score(currpl)=>replay3 and rep(currpl)=2 then
  credit=credit+1
  DOF 128, DOFOn
    playsound SoundFXDOF("knocke",126,DOFPulse,DOFKnocker)
  DOF 125, DOFPulse
  credittxt.text=credit
  If B2SOn Then
    Controller.B2SSetCredits Credit
  end if
  rep(currpl)=3
  playsound "click"

  if score(currpl)=>replay4 and rep(currpl)=3 then
  credit=credit+1
  DOF 128, DOFOn
    playsound SoundFXDOF("knocke",126,DOFPulse,DOFKnocker)
  DOF 125, DOFPulse
  credittxt.text=credit
  If B2SOn Then
    Controller.B2SSetCredits Credit
  end if
  rep(currpl)=4
  playsound "click"
  end if
  end if
    end if
end sub

sub scntimer_timer
  scn1=scn1 + 1
  matchnumb=matchnumb+1
  if matchnumb=10 then matchnumb=0
  if bell=3 then playsound SoundFXDOF("bell1000",143,DOFPule,DOFChimes),0, 0.15, 0, 0
  if bell=2 then playsound SoundFXDOF("bell100",142,DOFPule,DOFChimes),0, 0.15, 0, 0
  if bell=1 then playsound SoundFXDOF("bell10",141,DOFPule,DOFChimes),0, 0.30, 0, 0
  if scn1=scn then
  bell=0
    scntimer.enabled=false
    end if
end sub


sub bumper1_hit
    if tilt=false then playsoundAtVol SoundFXDOF("jet1",107,DOFPulse,DOFContactors), bumper1, volbump
  DOF 108,DOFPulse
    if BumperLight1.state=1 then addscore 10 else addscore 1
end sub

sub bumper2_hit
    if tilt=false then playsoundAtVol SoundFXDOF("jet1",113,DOFPulse,DOFContactors), bumper2, volbump
  DOF 114,DOFPulse
    if BumperLight2.state=1 then addscore 10 else addscore 1
end sub

sub bumper3_hit
    if tilt=false then playsoundAtVol SoundFXDOF("jet1",109,DOFPulse,DOFContactors), bumper3,volbump
  DOF 110,DOFPulse
    if BumperLight3.state=1 then addscore 10 else addscore 1
end sub

sub bumper4_hit
    if tilt=false then playsoundAtVol SoundFXDOF("jet1",111,DOFPulse,DOFContactors), bumper4, volbump
  DOF 112,DOFPulse
    if BumperLight4.state=1 then addscore 10 else addscore 1
end sub

sub wheelchange_timer
  wv=wv+1
  If wv>9 then wv=0
  If wv=9 then Light24.State=1: Light23.State=0
  If wv=8 then Light23.State=1: Light22.State=0
  If wv=7 then Light22.State=1: Light21.State=0
  If wv=6 then Light21.State=1: Light20.State=0
  If wv=5 then Light20.State=1: Light19.State=0
  If wv=4 then Light19.State=1: Light18.State=0
  If wv=3 then Light18.State=1: Light17.State=0
  If wv=2 then Light17.State=1: Light16.State=0
  If wv=1 then Light16.State=1: Light15.State=0
  If wv=0 then Light15.State=1: Light24.State=0
  If (Light11.State=1 and (wv=1 or wv=5 or wv=9)) or (Light12.State=1 and (wv=2 or wv=6)) or (Light13.State=1 and (wv=3 or wv=7)) or (Light14.State=1 and (wv=0 or wv=4 or wv=8)) then Light10.State=1 Else Light10.State=0
end sub

sub wheelcheck
    If wv=9 and Light10.State=1 then Addscore 300
    If wv=9 and Light10.State=0 then Addscore 30
    If wv=9 then Light11.State=1

    If wv=8 and Light10.State=1 then Addscore 400
    If wv=8 and Light10.State=0 then Addscore 40
    If wv=8 then Light14.State=1

    If wv=7 and Light10.State=1 then Addscore 500
    If wv=7 and Light10.State=0 then Addscore 50
    If wv=7 then Light13.State=1

    If wv=6 and Light10.State=1 then Addscore 400
    If wv=6 and Light10.State=0 then Addscore 40
    If wv=6 then Light12.State=1

    If wv=5 and Light10.State=1 then Addscore 300
    If wv=5 and Light10.State=0 then Addscore 30
    If wv=5 then Light11.State=1

    If wv=4 and Light10.State=1 then Addscore 400
    If wv=4 and Light10.State=0 then Addscore 40
    If wv=4 then Light14.State=1

    If wv=3 and Light10.State=1 then Addscore 500
    If wv=3 and Light10.State=0 then Addscore 50
    If wv=3 then Light13.State=1

    If wv=2 and Light10.State=1 then Addscore 400
    If wv=2 and Light10.State=0 then Addscore 40
    If wv=2 then Light12.State=1

    If wv=1 and Light10.State=1 then Addscore 300
    If wv=1 and Light10.State=0 then Addscore 30
    If wv=1 then Light11.State=1

    If wv=0 and Light10.State=1 then Light3.State=1: Addscore 200: if B2SOn then Controller.B2SSetShootAgain 1
    If wv=0 and Light10.State=0 then Light3.State=1: Addscore 20: if B2SOn then Controller.B2SSetShootAgain 1
    If wv=0 then Light14.State=1
    Playsound "motorshort1s"
end sub


sub Kicker1_hit
    if tilt=false then
    hv=hv+1
    If hv>6 then hv=6
    If hv=6 then addscore 500 else addscore 50
    If hv=5 then Light9.State=1
    If hv=4 then Light8.State=1
    If hv=3 then Light7.State=1
    If hv=2 then Light6.State=1
    If hv=1 then Light5.State=1
    If hv=0 then Light5.State=0: Light6.State=0: Light7.State=0: Light8.State=0: Light9.State=0
    playsound "motorshort1s"
    end if
    playsound "kickerkick"
    kicktimer.enabled=true
    playsound "click"
end sub

sub Kicker2_hit
    if tilt=false then
    wheelchange.enabled=False
    wheelcheck
    end if
    playsound "kickerkick"
    kicktimer.enabled=true
    playsound "click"
end sub

sub Kicker3_hit
    if tilt=false then
    wheelchange.enabled=False
    wheelcheck
    end if
    playsound "kickerkick"
    kicktimer.enabled=true
    playsound "click"
end sub


sub kicktimer_timer
    Kicker1.kick (int(rnd(1)*5)+195),15
    Kicker2.kick (int(rnd(1)*10)+130),15
    Kicker3.kick (int(rnd(1)*10)+220),15
  DOF 124, DOFPulse
  DOF 125, DOFPulse
    kicktimer.enabled=false
    wheelchange.enabled=true
end sub

sub ballhome_hit
    Flipper1.RotateToStart
    BumperLight2.state=0
    BumperLight3.state=0
    BumperLight1.state=0
    BumperLight4.state=0
    Light2.State=LifgtStateOff
    Light4.State=LifgtStateOff
    wheelchange.enabled=True
end Sub

sub ballhome_unhit
  DOF 129, DOFPulse
end sub

' Thalamus : This sub is used twice - this means ... this one IS NOT USED
'
' sub matchnum
'     match(matchnumb).text=matchnumb
'     for i=0 to playno
'     if matchnumb=(score(i) mod 10) then
'     credit=credit+1
'   DOF 128, DOFOn
'     playsound SoundFXDOF("knocke",126,DOFPulse,DOFKnocker)
'   DOF 125, DOFPulse
'     credittxt.text= credit
'     playsound "click"
'     end if
'     next
' end sub


' Thalamus : This sub is used twice - this means ... this one IS NOT USED
'
' sub turnoff
'     if (post.isdropped)=false then playsound "postdown"
'     post.isdropped=true
'     tiltseq.play seqalloff
'     lsling.isdropped=true
'     rsling.isdropped=true
'     wheelchange.enabled=False
' end sub

Sub Drain_Hit()
  DOF 127, DOFPulse
  playsoundAtVol "drainshorter", drain ,1
  Drain.DestroyBall
  nextball
End Sub

sub nextball
    hv=0
    BumperLight1.State=0
    BumperLight2.State=0
    BumperLight3.State=0
    BumperLight4.State=0
    Light5.State=LifgtStateOff
    Light6.State=LifgtStateOff
    Light7.State=LifgtStateOff
    Light8.State=LifgtStateOff
    Light9.State=LifgtStateOff
    Light15.State=LifgtStateOff
    Light16.State=LifgtStateOff
    Light17.State=LifgtStateOff
    Light18.State=LifgtStateOff
    Light19.State=LifgtStateOff
    Light20.State=LifgtStateOff
    Light21.State=LifgtStateOff
    Light22.State=LifgtStateOff
    Light23.State=LifgtStateOff
    Light24.State=LifgtStateOff
    Light10.State=LifgtStateOff
    Light11.State=LifgtStateOff
    Light12.State=LifgtStateOff
    Light13.State=LifgtStateOff
    Light14.State=LifgtStateOff
    Post.IsDropped=True
  if tilt=true then
  tilt=false
  If B2SOn Then Controller.B2SSetTilt 0
  tilttext.text=" "
  tiltseq.stopplay
    bipreel.setvalue(1)
    pno.setvalue(playno)
    end if

  if (Light3.state)=1 then
  playsound "kickerkick"
  ballreltimer.enabled=true
  Light3.state=0 : if B2SOn then Controller.B2SSetShootAgain 0
  else
  currpl=currpl+1
  If B2SOn Then
    Controller.B2SSetData 81,0
    Controller.B2SSetData 82,0
    Controller.B2SSetData 83,0
    Controller.B2SSetData 84,0
    Controller.B2SSetData (80+currpl),1
  end If
  end if

  if currpl>playno then
  ballinplay=ballinplay+1
  if ballinplay>5 then
  playsound "motorleer"
  If B2SOn then Controller.B2SSetBallInPlay 0
  eg=1
  ballreltimer.enabled=true
  else
  if state=true and tilt=false then
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
  bipreel.setvalue(0)
  state=false
  for i=1 to 4
  if truesc(i)>hisc then
  hisc=truesc(i)
  reel5.setvalue(hisc)
  end if
  next
  pno.setvalue(0)
  play(currpl-1).state=0
  playno=0
  gamov.text="GAME OVER"
  If B2SOn Then
    Controller.B2SSetGameOver 1
  end if
  wheelchange.enabled=False
  savehs
  ballreltimer.enabled=false
  else
  nb.CreateBall
  nb.kick 135,6
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
    if (matchnumb*10)=(score(i) mod 10) then
    credit=credit+1
  DOF 128, DOFOn
    playsound SoundFXDOF("knocke",126,DOFPulse,DOFKnocker)
  DOF 125, DOFPulse
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
  tilttext.text="TILT"
  If B2SOn Then Controller.B2SSetTilt 1
  post.isdropped=true
  tiltsens = 0
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
  tilttext.text="TILT"
  If B2SOn Then Controller.B2SSetTilt 1
  post.isdropped=true
  tiltsens = 0
  playsound "tilt"
  turnoff
End Sub

Sub Tilttimer_Timer()
  Tilttimer.Enabled = False
End Sub

sub turnoff
    tiltseq.play seqalloff
    bipreel.setvalue(0)
    pno.setvalue(5)
end sub



'**********Sling Shot Animations
' Rstep and Lstep  are the variables that increment the animation
'****************
Dim RStep, Lstep

Sub RightSlingShot_Slingshot
    Addscore 1
    PlaySoundAtVol SoundFXDOF("right_slingshot",105,DOFPulse,DOFContactors), sling1, 1
  DOF 106, DOFPulse
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
    Addscore 1
    PlaySoundAtVol SoundFXDOF("left_slingshot",103,DOFPulse,DOFContactors), sling2, 1
  DOF 104, DOFPulse
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
  PlaySound "pinhit_low", 0, Vol(ActiveBall)*VolPi, AudioPan(ActiveBall), 0, Pitch(ActiveBall), 0, 0, AudioFade(ActiveBall)
End Sub

Sub a_Targets_Hit (idx)
  PlaySound "target", 0, Vol(ActiveBall)*VolTarg, AudioPan(ActiveBall), 0, Pitch(ActiveBall), 0, 0, AudioFade(ActiveBall)
End Sub

Sub a_Metals_Thin_Hit (idx)
  PlaySound "metalhit_thin", 0, Vol(ActiveBall)*VolMetal, AudioPan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
End Sub

Sub a_Metals_Medium_Hit (idx)
  PlaySound "metalhit_medium", 0, Vol(ActiveBall)*VolMetal, AudioPan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
End Sub

Sub a_Metals2_Hit (idx)
  PlaySound "metalhit2", 0, Vol(ActiveBall)*VolMetal, AudioPan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
End Sub

Sub a_Gates_Hit (idx)
  PlaySound "gate4", 0, Vol(ActiveBall)*VolGates, AudioPan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
End Sub

Sub a_Spinner_Spin
  PlaySound "fx_spinner", a_Spinner, VolSpin
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
    PlaySound "fx_rubber2", 0, Vol(ActiveBall)*VolPo, Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
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
  ' Based on Black's Highscore routines
  Dim FileObj
  Dim ScoreFile
  Set FileObj=CreateObject("Scripting.FileSystemObject")
  If Not FileObj.FolderExists(UserDirectory) then
    Exit Sub
  End if
  Set ScoreFile=FileObj.CreateTextFile(UserDirectory & "Smart_Set.txt",True)
    scorefile.writeline credit
    ScoreFile.WriteLine score(0)
        ScoreFile.WriteLine score(1)
        ScoreFile.WriteLine score(2)
        ScoreFile.WriteLine score(3)
        scorefile.writeline hisc
    scorefile.writeline matchnumb
        scorefile.writeline wv
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
    Set FileObj=CreateObject("Scripting.FileSystemObject")
  If Not FileObj.FolderExists(UserDirectory) then
    Exit Sub
  End if
  If Not FileObj.FileExists(UserDirectory & "Smart_Set.txt") then
    Exit Sub
  End if
  Set ScoreFile=FileObj.GetFile(UserDirectory & "Smart_Set.txt")
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
        TextStr.Close
    credit = CDbl(temp1)
    score(0) = CDbl(temp2)
      score(1) = CDbl(temp3)
      score(2) = CDbl(temp4)
      score(3) = CDbl(temp5)
      hisc = CDbl(temp6)
      matchnumb = CDbl(temp7)
      wv = CDbl(temp8)
      Set ScoreFile=Nothing
      Set FileObj=Nothing
end sub

Sub Trigger1_Hit()
    DOF 119, DOFPulse
        Light2.State=1
        BumperLight1.state=1
        BumperLight4.state=1
        If Light2.State=1 and Light4.State=1 then
        Flipper1.RotateToEnd
        End If
  Addscore 100
End Sub

Sub Trigger2_Hit()
    DOF 120, DOFPulse
        Light4.State=1
        BumperLight2.state=1
        BumperLight3.state=1
        If Light2.State=1 and Light4.State=1 then
        Flipper1.RotateToEnd
        End If
  Addscore 100
End Sub

Sub Target1_Hit()
  DOF 115, DOFPulse
  Addscore 10
    Post.IsDropped=False
  DOF 130, DOFPulse
End Sub

Sub Target2_Hit()
  DOF 117, DOFPulse
  Addscore 10
    Post.IsDropped=False
  DOF 130, DOFPulse
End Sub

Sub Target3_Hit()
  DOF 115, DOFPulse
  Addscore 10
    Post.IsDropped=True
End Sub

Sub Target4_Hit()
  DOF 117, DOFPulse
  Addscore 10
    Post.IsDropped=True
End Sub

Sub Target5_Hit()
  DOF 116, DOFPulse
  Addscore 10
        BumperLight1.state=1
        BumperLight4.state=1
        Light2.State=1
        If Light2.State=1 and Light4.State=1 then
        Flipper1.RotateToEnd
        End If
End Sub

Sub Target6_Hit()
  DOF 118, DOFPulse
  Addscore 10
        BumperLight2.state=1
        BumperLight3.state=1
        Light4.State=1
        If Light2.State=1 and Light4.State=1 then
        Flipper1.RotateToEnd
        End If
End Sub

Sub Trigger3_Hit()
  DOF 121, DOFPulse
    wheelchange.enabled=False
    wheelcheck
    kicktimer.enabled=true
End Sub

Sub Trigger4_Hit()
  DOF 122, DOFPulse
    wheelchange.enabled=False
    wheelcheck
    kicktimer.enabled=true
End Sub

Sub Trigger5_Hit()
  Addscore 100
End Sub


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

Sub Table1_Exit
  If B2SOn Then Controller.Stop
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
Sub Table1_exit()
  Controller.Pause = False
  Controller.Stop
End Sub


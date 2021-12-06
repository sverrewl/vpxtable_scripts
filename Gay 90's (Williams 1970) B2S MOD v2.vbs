
' Thalamus 2018-07-23
' Added/Updated "Positional Sound Playback Functions" and "Supporting Ball & Sound Functions"
' Thalamus 2018-11-01 : Improved directional sounds
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
Const VolPo     = 1    ' Rubber posts volume.
Const VolPi     = 1    ' Rubber pins volume.
Const VolTarg   = 1    ' Targets volume.
Const VolKick   = 1    ' Kicker volume.
Const VolSpin   = 1.5  ' Spinners volume.
Const VolFlip   = 1    ' Flipper volume.


Dim Score(4)
dim truesc(4)
dim reel(5)
dim ballrelenabled
dim state
dim playno
dim credit
dim eg
dim currpl
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
dim scn
dim scn1
dim bell
dim points
dim tempscore
dim replay1
dim replay2
dim replay3
dim hisc
dim up(4)
dim hv
dim wv
dim laf
dim target4X
dim i
dim movetarget
dim movedirection
dim kickstep
dim rdwastep, rdwbstep, rdwcstep, rdwdstep, rdwestep

Dim B2SOn
Dim controller
ExecuteGlobal GetTextFile("core.vbs")

sub MovingTargetTimer_timer
  if movedirection=1 Then
    Collection2(movetarget).visible=0
    Collection2(movetarget).collidable=False
    Collection2(movetarget).hashitevent=False
    movetarget=movetarget+1
'   Collection2(movetarget).visible=1
    Collection2(movetarget).collidable=True
    Collection2(movetarget).hashitevent=True

    if movetarget>48 then
      movedirection=0
    end If
    PtargetC.objroty=((movetarget*.8)-20.5)*-1
    exit sub
  Else
    Collection2(movetarget).Visible=0
    Collection2(movetarget).collidable=False
    Collection2(movetarget).hashitevent=False

    movetarget=movetarget-1
'   Collection2(movetarget).visible=1
    Collection2(movetarget).collidable=True
    Collection2(movetarget).hashitevent=True

    if movetarget<1 then
      movedirection=1
    end If
  end If
  PtargetC.objroty=((movetarget*.8)-20.5)*-1
end sub


sub table1_init
  If ShowDT=false Then
    for each obj in DesktopItems
      obj.visible=False
    Next
  end If
  for i = 0 to 49
    Collection2(i).Visible=0
    Collection2(i).collidable=false
    Collection2(i).hashitevent=False
  Next
' CenterT50.Visible=1
  CenterT50.Collidable=True
  CenterT50.Hashitevent=True

  movetarget=0
  movetargetup=0
  movetargetupcount=0
  movedirection=1
  MovingTargetTimer.enabled=false
  PtargetC.objroty=((movetarget*.8)-20.5)*-1
  StopperFlip.rotatetoend
    wv=0
    hv=0
    targetsdown=0
    set play(1)=up1
    set play(2)=up2
    set play(3)=up3
    set play(4)=up4
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
    set up(1)=up1
    set up(2)=up2
    set up(3)=up3
    set up(4)=up4
    set reel(1)=reel1
    set reel(2)=reel2
    set reel(3)=reel3
    set reel(4)=reel4
    set reel(5)=reel5
    replay1=5000
    replay2=7400
    replay3=9800
    loadhs
    if hisc="" then hisc=1000
    reel5.setvalue(hisc)
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
end sub


Sub Table1_KeyDown(ByVal keycode)

    If keycode = PlungerKey Then
  Plunger.PullBack
  End If

    if keycode = AddCreditKey then
  playsoundAtVol "coin3", drain, 1
  coindelay.enabled=true
    credit=credit+4
  savehs
  If B2SOn Then
    Controller.B2SSetCredits Credit
  end if
  end if

    if keycode = 5 then
  playsoundAtVol "coin3", drain, 1
  coindelay.enabled=true
  If B2SOn Then
    Controller.B2SSetCredits Credit
  end if
  end if

    If Keycode = 4 Then
    Target4Mover
    End If

  if keycode = 2 and credit>0 and state=false and playno=0 then
  credit=credit-1
  credittxt.text=credit
    eg=0
    playno=1
  MovingTargetTimer.enabled=true
    currpl=1
    pno.setvalue(playno)
    if playno=1 then pno1.state=lightstateon
    pno2.state=lightstateoff
    pno3.state=lightstateoff
    pno4.state=lightstateoff
    play(currpl).state=lightstateon
    playsound "Switch"
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
    if playno=2 then pno2.state=lightstateon
    if playno=3 then pno3.state=lightstateon
    if playno=4 then pno4.state=lightstateon
    playsound "Switch"
    If B2SOn Then
      Controller.B2SSetCredits Credit
      Controller.B2SSetCanPlay playno
    end if
    end if

    if state=true and tilt=false then

  If keycode = LeftFlipperKey Then
    LeftFlipper.RotateToEnd
    PlaySoundAtVol "FlipperUp", LeftFlipper, VolFlip
        playsoundAtVol "buzz", LeftFlipper, VolFlip
  End If

  If keycode = RightFlipperKey Then
    RightFlipper.RotateToEnd
        PlaySoundAtVol "FlipperUp", RightFlipper, VolFlip
        playsoundAtVol "buzz", RightFlipper, VolFlip
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

  end if

End Sub

Sub Table1_KeyUp(ByVal keycode)

  If keycode = PlungerKey Then
    Plunger.Fire
    playsoundAtVol "plunger", Plunger, 1
  End If

  If keycode = LeftFlipperKey Then
    LeftFlipper.RotateToStart
    stopsound "buzz"
    if state=true and tilt=false then PlaySoundAtVol "FlipperDown", LeftFlipper, VolFlip
  End If

  If keycode = RightFlipperKey Then
    RightFlipper.RotateToStart
    stopsound "buzz"
        if state=true and tilt=false then PlaySoundAtVol "FlipperDown", RightFlipper, VolFlip
  End If

End Sub

sub flippertimer_timer()
  BottomGate.RotY = Flipper2.CurrentAngle+90
  BottomGate1.RotY = Flipper1.CurrentAngle+90
  Pstopper.Transz= StopperFlip.CurrentAngle
end sub


sub coindelay_timer
    playsound "Switch"
    credit=credit+1
  credittxt.text=credit
    coindelay.enabled=false
  If B2SOn Then
    Controller.B2SSetCredits Credit
  end if
end sub

sub resettimer_timer
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
    end if
end sub


Sub Target4Mover
    Target4.PositionX=360
    Target4.Visible=False
    Target4X=Target4.X
    Target4X=Target4X+10
    Target4.Visible=True
End Sub


sub addscore(points)
    if tilt=false then
    bell=0
    if points=10 or points=1 then
    matchnumb=matchnumb+1
    end if
    if matchnumb=10 then matchnumb=0
    end if
    if points = 1 or points = 10 or points = 100 then scn=1
    if points = 100 then
    reel(currpl).addvalue(100)
    bell=100
    end if
    if points = 10 then
    reel(currpl).addvalue(10)
    bell=10
    end if
    if points = 1 then
    reel(currpl).addvalue(1)
    bell=1
    end if
    if points = 300 then
    reel(currpl).addvalue(300)
    scn=3
    bell=100
    end if
    if points = 200 then
    reel(currpl).addvalue(200)
    scn=2
    bell=100
    end if
    if points = 500 then
    reel(currpl).addvalue(500)
    scn=5
    bell=100
    end if
    if points = 400 then
    reel(currpl).addvalue(400)
    scn=4
    bell=100
    end if
    if points = 30 then
    reel(currpl).addvalue(30)
    scn=3
    bell=10
    end if
    if points = 20 then
    reel(currpl).addvalue(20)
    scn=2
    bell=10
    end if
    if points = 50 then
    reel(currpl).addvalue(50)
    scn=5
    bell=10
    end if
    if points = 40 then
    reel(currpl).addvalue(40)
    scn=4
    bell=10
    end if
    scn1=0
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
    rep(currpl)=1
    playsound "Switch"
    end if
    if score(currpl)=>replay2 and rep(currpl)=1 then
    credit=credit+1
    playsound "knocke"
    credittxt.text=credit
  If B2SOn Then
    Controller.B2SSetCredits Credit
  end if
    rep(currpl)=2
    playsound "Switch"
    end if
    if score(currpl)=>replay3 and rep(currpl)=2 then
    credit=credit+1
    playsound "knocke"
    credittxt.text=credit
  If B2SOn Then
    Controller.B2SSetCredits Credit
  end if
    rep(currpl)=3
    playsound "Switch"
    end if
end sub

sub scntimer_timer
    scn1=scn1 + 1
    if bell=1 then playsound "bell1",0, 0.10, 0, 0
    if bell=10 then playsound "bell10",0, 0.15, 0, 0
    if bell=100 then playsound "bell100",0, 0.30, 0, 0
    if scn1=scn then scntimer.enabled=false
end sub


sub newgame
  wv=0
  wheelcheck
  If Target1.IsDropped=True then Target1.IsDropped=False
  If Target2.IsDropped=True then Target2.IsDropped=False
' DropTarget3.IsDropped = False ' Reset all three drop targets
' DropTarget4.IsDropped = False
' DropTarget5.IsDropped = False
  Playsound "solenoid"
  TargetsDown = 0
  Light8.State=LightStateOff
  Light9.State=LightStateOff
  Stopper.IsDropped=True
  StopperFlip.rotatetostart
  Lstopper.state=0
  state=true
  wv=0
  wheelcheck
  rep(1)=0
  rep(2)=0
  rep(3)=0
  eg=0
  laf=0
  score(1)=0
  score(2)=0
  score(3)=0
  score(4)=0
  truesc(1)=0
  truesc(2)=0
  truesc(3)=0
  truesc(4)=0
  bumper1.force=5
  bumper2.force=5
  bumper3.force=5
  bumper4.force=1
  bumper5.force=1
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
end sub


Sub Drain_Hit()
  playsoundAtVol "drainshorter", drain, 1
  Drain.DestroyBall
    wv=0
    wheelcheck
    Flipper1.RotateToStart
    Flipper2.RotateToStart
    nextball
End Sub

sub nextball
  if tilt=true then
  tilt=false
  If B2SOn Then Controller.B2SSetTilt 0:Controller.B2SSetShootAgain 0
  tilttext.text=" "
    Stopper.IsDropped=True
  StopperFlip.rotatetostart
  Lstopper.state=0
'    DropTarget3.IsDropped = False ' Reset all three drop targets
'    DropTarget4.IsDropped = False
'    DropTarget5.IsDropped = False
    TargetsDown = 0
    If Target1.IsDropped=True then Target1.IsDropped=False
    If Target2.IsDropped=True then Target2.IsDropped=False
    Light8.State=LightStateOff
    Light9.State=LightStateOff
    Bumper1.force=5
    bumper2.force=5
    bumper3.force=5
    bumper4.force=1
    bumper5.force=1
    tiltseq.stopplay
    pno.setvalue(playno)
'    LeftSlingShot.isdropped=false
'    RightSlingShot.isdropped=false
  end if
    newball
'    DropTarget3.IsDropped = False ' Reset all three drop targets
'    DropTarget4.IsDropped = False
'    DropTarget5.IsDropped = False
    TargetsDown = 0
  ballreltimer.enabled=true
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
  if state=true and tilt=false then
  play(currpl-1).state=lightstateoff
  currpl=1
  play(currpl).state=lightstateon
  If B2SOn Then
    Controller.B2SSetData 81,0
    Controller.B2SSetData 82,0
    Controller.B2SSetData 83,0
    Controller.B2SSetData 84,0
    Controller.B2SSetData (80+currpl),1
  end If
  newball
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
  play(currpl-1).state=lightstateoff
  play(currpl).state=lightstateon
  If B2SOn Then
    Controller.B2SSetData 81,0
    Controller.B2SSetData 82,0
    Controller.B2SSetData 83,0
    Controller.B2SSetData 84,0
    Controller.B2SSetData (80+currpl),1
  end If
  newball
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
    'bipreel.setvalue(0)
    state=false
    for i=1 to 4
    if truesc(i)>hisc then
      hisc=truesc(i)
      reel5.setvalue(hisc)
    end if
    next
    pno.setvalue(0)
    play(currpl-1).state=lightstateoff
    playno=0
    gamov.text="GAME OVER"
    If B2SOn Then
      Controller.B2SSetGameOver 1
    end if
    MovingTargetTimer.enabled=false
    savehs
    ballreltimer.enabled=false
    else
    nb.CreateBall
    nb.kick 90,6
    Stopper.IsDropped=True
    StopperFlip.rotatetostart
    Lstopper.state=0
    ballreltimer.enabled=false
    end if
end sub

sub newball
    If Target1.IsDropped=True then Target1.IsDropped=False
    If Target2.IsDropped=True then Target2.IsDropped=False
    Light8.State=LightStateOff
    Light9.State=LightStateOff
    loutl.state=lightstateoff
    routl.state=lightstateoff
    hv=0
    Light1.State=LightStateOn
    Light2.State=LightStateOff
    Light3.State=LightStateOff
    Light4.State=LightStateOff
    Light5.State=LightStateOff
    Light6.State=LightStateOff
    Light7.State=LightStateOff
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
  If B2SOn Then
    Controller.B2SSetCredits Credit
  end if
    playsound "knocke"
    credittxt.text= credit
    playsound "Switch"
    end if
    next
end sub

Sub CheckTilt
  If Tilttimer.Enabled = True Then
  TiltSens = TiltSens + 1
  if TiltSens = 2 Then
  Tilt = True
  tilttext.text="TILT"
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

Sub MechCheckTilt
  Tilt = True
  tilttext.text="TILT"
  If B2SOn Then Controller.B2SSetTilt 1
  tiltsens = 0
  playsound "tilt"
  turnoff
  LeftFlipper.RotateToStart
  RightFlipper.RotateToStart
  Stopper.IsDropped=True
    StopperFlip.rotatetostart
  Lstopper.state=0
End Sub

Sub Tilttimer_Timer()
  Tilttimer.Enabled = False
End Sub

sub turnoff
  bumper1.force=0
  bumper2.force=0
  bumper3.force=0
  bumper4.force=0
  bumper5.force=0
  tiltseq.play seqalloff
    pno.setvalue(5)
'    LeftSlingShot.isdropped=true
'    RightSlingShot.isdropped=true
end sub


Sub lsling2_Slingshot()
  if tilt=false then addscore 1
End Sub

Sub rsling2_Slingshot()
  if tilt=flase then addscore 1
End Sub

Sub lsling3_Slingshot()
  if tilt=false then addscore 1
End Sub

Sub rsling3_Slingshot()
  if tilt=flase then addscore 1
End Sub

Sub lsling4_Slingshot()
  if tilt=false then addscore 1
End Sub

Sub rsling4_Slingshot()
  if tilt=flase then addscore 1
End Sub

Sub lsling5_Slingshot()
  if tilt=false then addscore 1
End Sub

Sub rsling5_Slingshot()
  if tilt=flase then addscore 1
End Sub

sub bumper1_hit
    if tilt=false then playsoundAtVol "jet1", bumper1, VolBump
    addscore 10
  me.timerenabled=1
end sub

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

sub bumper2_hit
    if tilt=false then playsoundAtVol "jet1", bumper2, VolBump
    addscore 10
  me.timerenabled=1
end sub

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

sub bumper3_hit
    if tilt=false then playsoundAtVol "jet1", bumper3, VolBump
    addscore 10
  me.timerenabled=1
end sub

Sub Bumper3_timer
  BumperTimerRing3.enabled=0
  if bumperring3.transz>-36 then BumperRing3.transz=BumperRing3.transz-4
  if BumperRing3.transz=-36 then
    BumperTimerRing3.enabled=1
    me.timerenabled=0
  end if
End Sub

Sub BumperTimerRing3_timer
  if bumperring3.transz<0 then BumperRing3.transz=BumperRing3.transz+4
  If BumperRing3.transz=0 then BumperTimerRing3.enabled=0
End sub

sub bumper4_hit
    if tilt=false then playsoundAtVol "buzz", Bumper4, VolBump
    addscore 10
    wv=wv+1
    wheelcheck
end sub

sub bumper5_hit
    if tilt=false then playsoundAtVol "buzz", Bumper5, VolBump
    addscore 10
    wv=wv+1
    wheelcheck
end sub

Sub Trigger1_Hit()
  Addscore 100
    HoleValue
End Sub

sub Kicker1_Hit()
    if tilt=false then
    holecheck
    playsoundAtVol "motorshort1s", Kicker1, VolKick
    end if
  kickstep=1
    me.timerenabled=1
end sub

sub Kicker1_timer
  Select Case kickstep
    Case 4:
    playsoundAtVol "kickerkick", Pkickarm, VolKick
    Pkickarm.rotz=10
    kicker1.kick 240,10
    Case 5:
    Pkickarm.rotz=0
    me.timerenabled=0
  End Select
  kickstep=kickstep+1
end Sub

Sub Kicker2_Hit()
    playsoundAtVol "motorshort1s", Kicker2, VolKick
    me.timerenabled=1
    loutl.State=LightStateOff
    Flipper1Timer.enabled=True
End Sub

sub kicker2_timer
    playsoundAtVol "kickerkick", Kicker2, VolKick
    Kicker2.kick 0, 20
    me.timerenabled=0
end sub

Sub Flipper1Timer_Timer
    Flipper1.RotateToStart
    Flipper1Timer.enabled=False
End Sub

Sub Flipper2Timer_Timer
    Flipper2.RotateToStart
    Flipper2Timer.enabled=False
End Sub


sub wheelcheck
    If wv=>5 then wv=0
    If wv=0 then
    Light10.State=LightStateOn
    Light11.State=LightStateOff
    Light12.State=LightStateOff
    Light13.State=LightStateOn
    Light14.State=LightStateOff
    Light15.State=LightStateOff
    Light16.State=LightStateOff
    Light17.State=LightStateOff
    Light18.State=LightStateOff
    Light19.State=LightStateOff
    End If
    If wv=1 then
    Light10.State=LightStateOff
    Light11.State=LightStateOff
    Light12.State=LightStateOff
    Light13.State=LightStateOff
    Light14.State=LightStateOff
    Light15.State=LightStateOff
    Light16.State=LightStateOn
    Light17.State=LightStateOn
    Light18.State=LightStateOff
    Light19.State=LightStateOff
    End If
    If wv=2 then
    Light10.State=LightStateOff
    Light11.State=LightStateOn
    Light12.State=LightStateOff
    Light13.State=LightStateOff
    Light14.State=LightStateOn
    Light15.State=LightStateOff
    Light16.State=LightStateOff
    Light17.State=LightStateOff
    Light18.State=LightStateOff
    Light19.State=LightStateOff
    End If
    If wv=3 then
    Light10.State=LightStateOff
    Light11.State=LightStateOff
    Light12.State=LightStateOff
    Light13.State=LightStateOff
    Light14.State=LightStateOff
    Light15.State=LightStateOff
    Light16.State=LightStateOff
    Light17.State=LightStateOff
    Light18.State=LightStateOn
    Light19.State=LightStateOn
    End If
    If wv=4 then
    Light10.State=LightStateOff
    Light11.State=LightStateOff
    Light12.State=LightStateOn
    Light13.State=LightStateOff
    Light14.State=LightStateOff
    Light15.State=LightStateOn
    Light16.State=LightStateOff
    Light17.State=LightStateOff
    Light18.State=LightStateOff
    Light19.State=LightStateOff
    End If

end sub


Sub HoleValue
    If Light16.State=LightStateOn then loutl.State=LightStateOn: Flipper1.RotateToEnd
    If Light19.State=LightStateOn then routl.State=LightStateOn: Flipper2.RotateToEnd
    If wv=0 or wv=2 or wv=4 then hv=hv+1
    If hv=0 then Light1.State=LightStateOn
    If hv=1 then Light1.state=LightstateOff: Light2.State=LightStateOn
    If hv=2 then Light2.state=LightstateOff: Light3.State=LightStateOn
    If hv=3 then Light3.state=LightstateOff: Light4.State=LightStateOn
    If hv=4 then Light4.state=LightstateOff: Light5.State=LightStateOn
    If hv=5 then Light5.state=LightstateOff: Light6.State=LightStateOn
    If hv=6 then Light6.state=LightstateOn: Light7.State=LightStateOn
    If hv>7 then hv=7

End Sub


Sub HoleCheck
    If Light1.State=LightStateOn then Addscore 50
    If Light2.State=LightStateOn then Addscore 100
    If Light3.State=LightStateOn then Addscore 200
    If Light4.State=LightStateOn then Addscore 300
    If Light5.State=LightStateOn then Addscore 400
    If Light6.State=LightStateOn then Addscore 500
    If Light7.State=LightStateOn then Credit=Credit+1: Playsound "Knock": credittxt.text= credit
End Sub

'**********Sling Shot Animations
' Rstep and Lstep  are the variables that increment the animation
'****************
Dim RStep, Lstep

Sub RightSlingShot_Slingshot
    PlaySoundAtVol "right_slingshot", sling1, 1
    RSling.Visible = 0
    RSling1.Visible = 1
    sling1.TransZ = -20
    RStep = 0
    RightSlingShot.TimerEnabled = 1
    addscore 1
End Sub

Sub RightSlingShot_Timer
    Select Case RStep
        Case 3:RSLing1.Visible = 0:RSLing2.Visible = 1:sling1.TransZ = -10
        Case 4:RSLing2.Visible = 0:RSLing.Visible = 1:sling1.TransZ = 0:RightSlingShot.TimerEnabled = 0
    End Select
    RStep = RStep + 1
End Sub

Sub LeftSlingShot_Slingshot
    PlaySoundAtVol "left_slingshot", sling2, 1
    LSling.Visible = 0
    LSling1.Visible = 1
    sling2.TransZ = -20
    LStep = 0
    LeftSlingShot.TimerEnabled = 1
    addscore 1
End Sub

Sub LeftSlingShot_Timer
    Select Case LStep
        Case 3:LSLing1.Visible = 0:LSLing2.Visible = 1:sling2.TransZ = -10
        Case 4:LSLing2.Visible = 0:LSLing.Visible = 1:sling2.TransZ = 0:LeftSlingShot.TimerEnabled = 0
    End Select
    LStep = LStep + 1
End Sub


sub savehs
    savevalue "Gay90s", "credit", credit
    savevalue "Gay90s", "hiscore", hisc
    savevalue "Gay90s", "match", matchnumb
    savevalue "Gay90s", "score1", truesc(1)
    savevalue "Gay90s", "score2", truesc(2)
    savevalue "Gay90s", "score2", truesc(3)
    savevalue "Gay90s", "score2", truesc(4)
end sub

sub loadhs
    dim temp
  temp = LoadValue("Gay90s", "credit")
    If (temp <> "") then credit = CDbl(temp)
    temp = LoadValue("Gay90s", "hiscore")
    If (temp <> "") then hisc = CDbl(temp)
    temp = LoadValue("Gay90s", "match")
    If (temp <> "") then matchnumb = CDbl(temp)
    temp = LoadValue("Gay90s", "score1")
    If (temp <> "") then score(1) = CDbl(temp)
    temp = LoadValue("Gay90s", "score2")
    If (temp <> "") then score(2) = CDbl(temp)
    temp = LoadValue("Gay90s", "score3")
    If (temp <> "") then score(3) = CDbl(temp)
    temp = LoadValue("Gay90s", "score4")
    If (temp <> "") then score(4) = CDbl(temp)
end sub


sub ballrel_hit
    if ballrelenabled=1 then
    playsoundAtVol "launchball", ballrel, 1
    ballrelenabled=0
    end if
end sub

Sub Trigger2_Hit()
  Addscore 100
End Sub

Sub Trigger3_Hit()
  Addscore 100
End Sub

Sub Trigger4_Hit()
  Addscore 100
    routl.State=LightStateOff
    Flipper2Timer.enabled=True
End Sub

Sub Target1_dropped()
  Target1.Isdropped=True
  Light8.State=LightStateOn
  Stopper.IsDropped=False
  StopperFlip.rotatetoend
  Lstopper.state=1
    Addscore 10
    TargetTimer.Enabled=True
End Sub

Sub Target2_dropped()
  Target2.Isdropped=True
  Light9.State=LightStateOn
  Stopper.IsDropped=False
  StopperFlip.rotatetoend
  Lstopper.state=1
    Addscore 10
    TargetTimer.Enabled=True
End Sub

Sub TargetTimer_Timer
    If Target2.IsDropped and Target1.IsDropped then Target2.IsDropped=False: Target1.IsDropped=False: Light8.State=LightStateOff: Light9.State=LightStateOff: Addscore 200: hv=hv+1: Playsound "solenoid":end if
    If hv=0 then Light1.State=LightStateOn
    If hv=1 then Light1.state=LightstateOff: Light2.State=LightStateOn
    If hv=2 then Light2.state=LightstateOff: Light3.State=LightStateOn
    If hv=3 then Light3.state=LightstateOff: Light4.State=LightStateOn
    If hv=4 then Light4.state=LightstateOff: Light5.State=LightStateOn
    If hv=5 then Light5.state=LightstateOff: Light6.State=LightStateOn
    If hv=6 then Light6.state=LightstateOn: Light7.State=LightStateOn
    If hv>7 then hv=7
    TargetTimer.Enabled=False
End Sub

Sub TargetTimer1_Timer
' DropTarget3.IsDropped = False ' Reset all three drop targets
' DropTarget4.IsDropped = False
' DropTarget5.IsDropped = False
  Playsound "solenoid"
    TargetTimer1.enabled=False
End Sub


'******************************************
'***    CenterTargets Handler   ***
'******************************************

Sub Collection2_hit(idx)
  AddScore 100
  HoleValue
end sub

Dim TargetsDown ' number of targets hit in the bank. Set this to 0 at start of your game
    Targetsdown=0


Sub DropTarget3_Hit()
    AddScore 10 ' Add some points
    TargetsDown = TargetsDown + 1 ' Bump the number of targets down by 1
    DropTarget3.IsDropped = True ' This actually drops the target
    CheckDropTargets ' Check to see if we got all of them
End Sub

' Hit Event for Drop Target2

Sub DropTarget4_Hit()
    AddScore 10 ' Add some points
    TargetsDown = TargetsDown + 1 ' Bump the number of targets down by 1
    DropTarget4.IsDropped = True ' This actually drops the target
    CheckDropTargets ' Check to see if we got all of them
End Sub

' Hit Event for Drop Target3

Sub DropTarget5_Hit()
    AddScore 10 ' Add some points
    TargetsDown = TargetsDown + 1 ' Bump the number of targets down by 1
    DropTarget5.IsDropped = True ' This actually drops the target
    CheckDropTargets ' Check to see if we got all of them
End Sub

Sub CheckDropTargets()
    If TargetsDown => 3 then
    Holevalue ' Add some points for clearing the bank if desired
    Addscore 300
    TargetTimer1.enabled=True
    TargetsDown = 0 ' Reset the counter to 0
End If
End Sub

Sub LeftSlingShot1_Slingshot()
  Addscore 1
    stopper.IsDropped=True
  StopperFlip.rotatetostart
  Lstopper.state=0
    wv=wv+1
    wheelcheck
  rdwa.visible=0
  RdwA1.visible=1
  rdwastep=1
  me.timerenabled=1
End Sub

Sub Leftslingshot1_timer
  select case rdwastep
    Case 1: RDWa1.visible=0: rdwa.visible=1
    case 2: rdwa.visible=0: rdwa2.visible=1
    Case 3: rdwa2.visible=0: rdwa.visible=1: Me.timerenabled=0
  end Select
  rdwastep=rdwastep+1
End sub

Sub LeftSlingShot2_Slingshot()
  Addscore 1
  rdwb.visible=0
  Rdwb1.visible=1
  rdwbstep=1
  me.timerenabled=1
End Sub

Sub Leftslingshot2_timer
  select case rdwbstep
    Case 1: RDWb1.visible=0: rdwb.visible=1
    case 2: rdwb.visible=0: rdwb2.visible=1
    Case 3: rdwb2.visible=0: rdwb.visible=1: Me.timerenabled=0
  end Select
  rdwbstep=rdwbstep+1
End sub

Sub LeftSlingShot3_Slingshot()
  Addscore 1
End Sub

Sub LeftSlingShot4_Slingshot()
  Addscore 1
End Sub

Sub LeftSlingShot5_Slingshot()
  Addscore 1
End Sub

Sub RightSlingShot1_Slingshot()
  Addscore 1
    stopper.IsDropped=True
  StopperFlip.rotatetostart
  Lstopper.state=0
    wv=wv+1
    wheelcheck
  rdwc.visible=0
  Rdwc1.visible=1
  rdwcstep=1
  me.timerenabled=1
End Sub

Sub Rightslingshot1_timer
  select case rdwcstep
    Case 1: RDWc1.visible=0: rdwc.visible=1
    case 2: rdwc.visible=0: rdwc2.visible=1
    Case 3: rdwc2.visible=0: rdwc.visible=1: Me.timerenabled=0
  end Select
  rdwcstep=rdwcstep+1
End sub

Sub RightSlingShot2_Slingshot()
  Addscore 1
  rdwd.visible=0
  Rdwd1.visible=1
  rdwdstep=1
  me.timerenabled=1
End Sub

Sub Rightslingshot2_timer
  select case rdwdstep
    Case 1: RDWd1.visible=0: rdwd.visible=1
    case 2: rdwd.visible=0: rdwd2.visible=1
    Case 3: rdwd2.visible=0: rdwd.visible=1: Me.timerenabled=0
  end Select
  rdwdstep=rdwdstep+1
End sub

Sub RightSlingShot3_Slingshot()
  Addscore 1
End Sub

Sub RightSlingShot4_Slingshot()
  Addscore 1
End Sub

Sub RightSlingShot5_Slingshot()
  Addscore 1
End Sub

Sub RightSlingShot6_Slingshot()
  Addscore 1
  rdwe.visible=0
  Rdwe1.visible=1
  rdwestep=1
  me.timerenabled=1
End Sub

Sub Rightslingshot6_timer
  select case rdwestep
    Case 1: RDWe1.visible=0: rdwe.visible=1
    case 2: rdwe.visible=0: rdwe2.visible=1
    Case 3: rdwe2.visible=0: rdwe.visible=1: Me.timerenabled=0
  end Select
  rdwestep=rdwestep+1
End sub

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
  PlaySoundAtVol "fx_spinner", a_Spinner, VolSpin
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

Sub Table1_Exit()
  Savehs
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
    PlaySound("fx_collide"), 0, Csng(velocity) ^2 / (VolDiv/VolCol), Pan(ball1), 0, Pitch(ball1), 0, 0, AudioFade(ball1)
End Sub


' Thalamus : Exit in a clean and proper way
Sub table1_exit()
  Controller.Pause = False
  Controller.Stop
End Sub


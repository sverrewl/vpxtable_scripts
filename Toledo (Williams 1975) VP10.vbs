'Williams Toledo 1975 VP10 table
'original vp8 script by Leon Spalding
'VP10 build and PF, plastics and backglass redraw by bodydump
dim ballrelease
dim credit
dim score(2)
dim truesc(2)
dim playno
dim currpl
dim plno(2)
dim play(2)
dim rst
dim ballinplay
dim eg
dim match(9)
dim rep(2)
dim tilt
dim tiltsens
dim state
dim reel(2)
dim cred
dim hisc
dim scn
dim scn1
dim bell
dim points
dim matchnumb
dim replay1
dim replay2
dim replay3
dim tl(8)
dim tla(8)
dim ba
dim batemp
dim b(10)
dim targ
dim fulleb
dim sa

Dim Controller
If showdt = false then LoadController

  Sub LoadController()
  Set Controller = CreateObject("B2S.Server")
  Controller.B2SName = "Toledo"
  Controller.Run()
  End Sub


sub Table1_init
    set plno(1)=plno1
    set plno(2)=plno2
    set play(1)=play1
    set play(2)=play2
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
    set tl(1)=l1
    set tl(2)=l2
    set tl(3)=l3
    set tl(4)=l4
    set tl(5)=l5
    set tl(6)=l6
    set tl(7)=l7
    set tl(8)=l8
    set tla(1)=l1a
    set tla(2)=l2a
    set tla(3)=l3a
    set tla(4)=l4a
    set tla(5)=l5a
    set tla(6)=l6a
    set tla(7)=l7a
    set tla(8)=l8a
    set b(10)=b10k
    set b(1)=b1k
    set b(2)=b2k
    set b(3)=b3k
    set b(4)=b4k
    set b(5)=b5k
    set b(6)=b6k
    set b(7)=b7k
    set b(8)=b8k
    set b(9)=b9k
    credtimer.enabled=true
    loadhs

    if credit="" then credit=0
    If showdt = true then credtxt.text=credit
  If showdt = false then
    dthigh.text = ""
    dtinitial1.visible = 0
    dtinitial2.visible = 0
    dtinitial3.visible = 0
  End If
  If showdt = true then
    EMReelHSBackdrop.visible = 0
    EMReelHSTitle.visible = 0
    EMReelHSName1.visible = 0
    EMReelHSName2.visible = 0
    EMReelHSName3.visible = 0
    EMReelHSNum1.visible = 0
    EMReelHSNum2.visible = 0
    EMReelHSNum3.visible = 0
    EMReelHSNum4.visible = 0
    EMReelHSNum5.visible = 0
    EMReelHSNum6.visible = 0
    EMReelHSComma.visible = 0
    playuptimer.enabled = 1
  End If
  If showdt = false then
    backdropimage = "toledoredraw"
    primitive22.visible = 0
    primitive40.visible = 0
'   primitive107.visible = 0
    reel1.visible = 0
    reel2.visible = 0
    tilttxt.text = ""
    gamov.text = ""
    credtxt.text = ""
    bip1.text = ""
    bip2.text = ""
    bip3.text = ""
  End If
    replay1=47000
    replay2=66000
    replay3=80000
    If showdt = true then rep1.text=replay1
    If showdt = true then rep2.text=replay2
    If showdt = true then rep3.text=replay3
    select case(matchnumb)
    case 0:
    If showdt = true then m0.text="00"
    case 1:
    If showdt = true then m1.text="10"
    case 2:
    If showdt = true then m2.text="20"
    case 3:
    If showdt = true then m3.text="30"
    case 4:
    If showdt = true then m4.text="40"
    case 5:
    If showdt = true then m5.text="50"
    case 6:
    If showdt = true then m6.text="60"
    case 7:
    If showdt = true then m7.text="70"
    case 8:
    If showdt = true then m8.text="80"
    case 9:
    If showdt = true then m9.text="90"
    end select
    for i=1 to 2
    currpl=i
    If showdt = true then
    reel(i).setvalue(score(i))
  else


  End If
    next
    currpl=0
  If nightday<=10 then gi10
  If nightday>10 and nightday<=20 then gi20
  If nightday>40 then gidimmer

end sub

'*****GI Lights On
dim xx
Sub gion()
  For each xx in GI:xx.State = 1: Next
  If showdt = true then
    B2l1.State = 1
    B1L2.State = 1
    B1L3.State = 0
    B2L3.State = 0
  Else
    B2l1.State = 0
    B1L2.State = 0
    B1L3.State = 1
    B2L3.State = 1
  End If
End Sub

Sub gioff()
  For each xx in GI:xx.State = 0: Next
  If showdt = true then
    B2l1.State = 0
    B1L2.State = 0
    B1L3.State = 0
    B2L3.State = 0
  Else
    B2l1.State = 0
    B1L2.State = 0
    B1L3.State = 0
    B2L3.State = 0
  End If
End Sub

Sub gi10()
  For each xx in GIground:xx.Intensity = 3:Next
  For each xx in GIhigh:xx.Intensity = 3:Next
  For each xx in GIhigh:xx.FallOff = 350:Next
  giLight31.Intensity = 3:giLight31.FallOff = 395
  giLight32.Intensity = 3:giLight32.FallOff = 395
  Light17.Intensity = 3
  Light25.Intensity = 3
  Light19.Intensity = 3
  Light26.Intensity = 3
  b1l2.Intensity = 20
  b2l1.Intensity = 20
  b1l3.Intensity = 20
  b2l3.Intensity = 20
End Sub
Sub gi20()
  For each xx in GIground:xx.Intensity = 2:Next
  For each xx in GIhigh:xx.Intensity = 2:Next
  For each xx in GIhigh:xx.FallOff = 300:Next
  giLight31.Intensity = 2:giLight31.FallOff = 350
  giLight32.Intensity = 2:giLight32.FallOff = 350
  Light17.Intensity = 2
  Light25.Intensity = 2
  Light19.Intensity = 2
  Light26.Intensity = 2
End Sub
Sub gidimmer()
  For each xx in GIground:xx.Intensity = .5:Next
  For each xx in GIhigh:xx.Intensity = .5:Next
  For each xx in GIhigh:xx.FallOff = 220:Next
  For each xx in GIplastics:xx.Intensity = 3:Next
  giLight31.Intensity = 1:giLight31.FallOff = 300
  giLight32.Intensity = 1:giLight32.FallOff = 300
  Light17.Intensity = .5
  Light25.Intensity = .5
  Light19.Intensity = .5
  Light26.Intensity = .5
  b1l2.Intensity = 3
  b2l1.Intensity = 3
  b1l3.Intensity = 3
  b2l3.Intensity = 3
End Sub


Sub Table1_KeyDown(ByVal keycode)
    If PostItHighScoreCheck(keycode) then Exit Sub
    if keycode = 6 then
  PlaySoundAtVol "coin3", Drain, 1
  coindelay.enabled=true
  end if

    if keycode = 2 and credit>0 and state=false and playno=0 then
  credit=credit-1
  If showdt = true then credtxt.text=credit
  gion
  eg=0
    playno=1
    currpl=1
    plno(playno).state=lightstateon
    If showdt = true then play(currpl).state=lightstateon
    playsound "click"
    playsound "initialize"
    rst=0
    ballinplay=1
    resettimer.enabled=true
    end if

    if keycode = 2 and credit>0 and state=true and playno>0 and playno<2 and ballinplay<2 then
    credit=credit-1
    If showdt = true then credtxt.text=credit
    plno(playno).state=lightstateoff
    playno=playno+1
    If showdt = true then plno(playno).state=lightstateon
    playsound "click"
    end if

  If keycode = PlungerKey Then
    Plunger.PullBack
  End If

    if tilt=false and state=true then
  If keycode = LeftFlipperKey Then
    LeftFlipper.RotateToEnd
    upl.rotatetoend
    Light17.state = 0: Light25.state=1
    PlaySoundAtVol "FlipperUp", ActiveBall, 1
  End If

  If keycode = RightFlipperKey Then
    RightFlipper.RotateToEnd
        upr.rotatetoend
    Light19.state = 0: Light26.state=1
    PlaySoundAtVol "FlipperUp", ActiveBall, 1
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

Sub Table1_KeyUp(ByVal keycode)

  If keycode = PlungerKey Then
    Plunger.Fire
    PlaySoundAtVol "plunger", Plunger, 1
  End If

  If keycode = LeftFlipperKey Then
    LeftFlipper.RotateToStart
    upl.rotatetostart
    if tilt=false and state=true then PlaySoundAtVol "FlipperDown", LeftFlipper, 1:Light17.state = 1:Light25.state=0
  End If

  If keycode = RightFlipperKey Then
    RightFlipper.RotateToStart
        upr.rotatetostart
    if tilt=false and state=true then PlaySoundAtVol "FlipperDown", RightFlipper, 1:Light19.state = 1:Light26.state=0
  End If

End Sub

sub coindelay_timer
    playsound "click"
    credit=credit+1
    if credit>24 then credit=24
    If showdt = true then credtxt.text=credit
  coindelay.enabled=false
end sub

Sub Drain_Hit()
  Drain.DestroyBall
  PlaySoundAtVol "drainshort", Drain, 1
  if tilt=false then
  bonuscount.enabled=true
  else
  nextball
    end if
End Sub

sub nextball
  if tilt=true then
  tilt=false
  tilttxt.text=" "
  If showdt = false then Controller.B2SSetTilt 33,0
  end if
  if (shootagain.state)=lightstateon then
  shootagain.state=lightstateoff
  If showdt = false then Controller.B2SSetShootAgain 36, 0
  playsound "kickerkick"
    newball
  ballreltimer.enabled=true
  else
  currpl=currpl+1
  end if
  if currpl>playno then
  ballinplay=ballinplay+1
  if ballinplay>5 then
  playsound "motorleer"
  eg=1
  ballinplay = 0
  ballreltimer.enabled=true
  else
  if state=true and tilt=false then
  If showdt = true then play(currpl-1).state=lightstateoff
  currpl=1
  If showdt = true then play(currpl).state=lightstateon
  newball
  playsound "kickerkick"
  ballreltimer.enabled=true
  end if
  select case (ballinplay)
  case 1:
  If showdt = true then bip1.text="1"
  case 2:
  bip1.text=" "
  If showdt = true then bip2.text="2"
  case 3:
  bip2.text=" "
  If showdt = true then bip3.text="3"
  end select
  end if
  end if
  if currpl>1 and currpl<(playno+1) then
  if state=true and tilt=false then
  If showdt = true then play(currpl-1).state=lightstateoff
  If showdt = true then play(currpl).state=lightstateon
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
  state=false
  plno(playno).state=lightstateoff
  play(1).state=lightstateoff
  play(2).state=lightstateoff
  playno=0
  currpl = 0
  If showdt = true then gamov.text="   GAME OVER"
  If showdt = false then Controller.B2SSetGameOver 35,1
  for i=1 to 2
  if truesc(i)>hisc then hisc=truesc(i)
  next
  If showdt = true then hstxt.text=HSAHighScore
  CheckNewHighScorePostIt2Player truesc(1), truesc(2)
  savehs
  cred=0
  credtimer.enabled=true
  ballreltimer.enabled=false
    else
    nb.CreateBall
  nb.kick 135,4
    ballreltimer.enabled=false
    end if
end sub

sub resettimer_timer
    rst=rst+1
    for i=1 to 2
    reel(i).resettozero
  score(1) = 0
  score(2) = 0
    next
    if rst=20 then
    playsound "kickerkick"
    end if
    if rst=24 then
    newgame
    resettimer.enabled=false
    end if
end sub

sub newgame
    credtimer.enabled=false
    If showdt = true then credittxt.text="Ported To VP By Leon Spalding, VP10mod by bodydump"
    state=true
    tilt=false
    bumper1.force=8
    bumper2.force=8
  Wall35.IsDropped = 1
  Wall36.IsDropped = 1
    rep(1)=0
    rep(2)=0
    eg=0
    ba=1
    fulleb=0
    sa=0
    targ=0
    score(1)=0
    score(2)=0
    truesc(1)=0
    truesc(2)=0
    If showdt = true then bip1.text="1"
    for i=0 to 9
    match(i).text=" "
    next
    k3k.state=lightstateoff
    tabl.state=lightstateon
    snl.state=lightstateoff
    abbon1.isdropped=true:rolloverprim4.image = "Red peg1"
    abbon2.isdropped=false:rolloverprim7.image = "Red peg2"
    for i=1 to 8
    tl(i).state=lightstateoff
    tla(i).state=lightstateoff
    next
    l1.state=lightstateon
    t1kl.state=lightstateoff
    mabl.state=lightstateoff
    b1k.state=lightstateon
    dbl.state=lightstateoff
    al.state=lightstateoff
    bl.state=lightstateoff
    cl.state=lightstateoff
    dl.state=lightstateoff
    shootagain.state=lightstateoff
  If showdt = false then Controller.B2SSetShootAgain 36, 0
    eb1.state=lightstateoff
    eb2.state=lightstateoff
    sp1.state=lightstateoff
    sp2.state=lightstateoff
    tilttxt.text=" "
  If showdt = false then Controller.B2SSetTilt 33,0
    gamov.text=" "
  If showdt = false then Controller.B2SSetGameOver 35,0
    tilt=false
    tiltsens=0
    ballinplay=1
    nb.CreateBall
  nb.kick 135,4
end sub

sub newball
  gion
    ba=1
    targ=0
    fulleb=0
    sa=0
    k3k.state=lightstateoff
    tabl.state=lightstateon
    snl.state=lightstateoff
    abbon1.isdropped=true:rolloverprim4.image = "Red peg1"
   abbon2.isdropped=false:rolloverprim7.image = "Red peg2"
    for i=1 to 8
    tl(i).state=lightstateoff
    tla(i).state=lightstateoff
    next
    l1.state=lightstateon
    t1kl.state=lightstateoff
    mabl.state=lightstateoff
    b1k.state=lightstateon
    dbl.state=lightstateoff
    al.state=lightstateoff
    bl.state=lightstateoff
    cl.state=lightstateoff
    dl.state=lightstateoff
    shootagain.state=lightstateoff
  If showdt = false then Controller.B2SSetShootAgain 36, 0
    eb1.state=lightstateoff
    eb2.state=lightstateoff
    sp1.state=lightstateoff
    sp2.state=lightstateoff
end sub

sub bonadv
    if tilt=false then
    ba=ba+1
    if ba>19 then ba=19
    batemp=ba
    if batemp>9 and batemp<20 then
    b10k.state=lightstateon
    batemp=batemp-10
    end if
    select case (batemp)
    case 0:
    b9k.state=lightstateoff
    case 1:
    b1k.state=lightstateon
    case 2:
    b1k.state=lightstateoff
    b2k.state=lightstateon
    case 3:
    b2k.state=lightstateoff
    b3k.state=lightstateon
    case 4:
    b3k.state=lightstateoff
    b4k.state=lightstateon
    case 5:
    b4k.state=lightstateoff
    b5k.state=lightstateon
    case 6:
    b5k.state=lightstateoff
    b6k.state=lightstateon
    case 7:
    b6k.state=lightstateoff
    b7k.state=lightstateon
    case 8:
    b7k.state=lightstateoff
    b8k.state=lightstateon
    case 9:
    b8k.state=lightstateoff
    b9k.state=lightstateon
    end select
    end if
end sub

sub bonuscount_timer
    playsound "motorshort"
    batemp=ba
    if batemp<10 then b10k.state=lightstateoff
    if batemp>9 and batemp<19 then
    b10k.state=lightstateon
    batemp=batemp-10
    end if
    select case (batemp)
    case 0:
    b1k.state=lightstateoff
    case 1:
    b2k.state=lightstateoff
    b1k.state=lightstateon
    case 2:
    b3k.state=lightstateoff
    b2k.state=lightstateon
    case 3:
    b4k.state=lightstateoff
    b3k.state=lightstateon
    case 4:
    b5k.state=lightstateoff
    b4k.state=lightstateon
    case 5:
    b6k.state=lightstateoff
    b5k.state=lightstateon
    case 6:
    b7k.state=lightstateoff
    b6k.state=lightstateon
    case 7:
    b8k.state=lightstateoff
    b7k.state=lightstateon
    case 8:
    b9k.state=lightstateoff
    b8k.state=lightstateon
    case 9:
    b9k.state=lightstateon
    end select
    if (dbl.state)=lightstateon then
    addscore 2000
    else
    addscore 1000
    end if
    ba=ba-1
    if ba=0 then
    nextball
    bonuscount.enabled=false
    end if
end sub

sub matchnum
    select case(matchnumb)
    case 0:
    If showdt = true then m0.text="00" else Controller.B2SSetMatch 100
    case 1:
    If showdt = true then m1.text="10" else Controller.B2SSetMatch 10
    case 2:
    If showdt = true then m2.text="20" else Controller.B2SSetMatch 20
    case 3:
    If showdt = true then m3.text="30" else Controller.B2SSetMatch 30
    case 4:
    If showdt = true then m4.text="40" else Controller.B2SSetMatch 40
    case 5:
    If showdt = true then m5.text="50" else Controller.B2SSetMatch 50
    case 6:
    If showdt = true then m6.text="60" else Controller.B2SSetMatch 60
    case 7:
    If showdt = true then m7.text="70" else Controller.B2SSetMatch 70
    case 8:
    If showdt = true then m8.text="80" else Controller.B2SSetMatch 80
    case 9:
    If showdt = true then m9.text="90" else Controller.B2SSetMatch 90
    end select
    for i=1 to playno
    if (matchnumb*10)=(score(i) mod 100) then
    credit=credit+1
    playsound "knocker"
    if credit>24 then credit=24
    If showdt = true then credtxt.text=credit
    playsound "click"
    end if
    next
end sub

sub addscore(points)
    if tilt=false then
    bell=0
    if points=10 then
    if (tabl.state)=lightstateon then
    tabl.state=lightstateoff
    snl.state=lightstateon
    abbon2.isdropped=true:rolloverprim7.image = "Red peg1"
    abbon1.isdropped=false:rolloverprim4.image = "Red peg2"
    else
    snl.state=lightstateoff
    tabl.state=lightstateon
    abbon1.isdropped=true:rolloverprim4.image = "Red peg1"
    abbon2.isdropped=false:rolloverprim7.image = "Red peg2"
    end if
    if (sp1.state)=lightstateon and targ>7 then
    sp1.state=lightstateoff
    sp2.state=lightstateon
    else
    if targ>7 then
    sp2.state=lightstateoff
    sp1.state=lightstateon
    end if
    end if
    if (eb1.state)=lightstateon and targ>4 and fulleb=0 and sa=0 then
    eb1.state=lightstateoff
    eb2.state=lightstateon
    else
    if targ>4 and fulleb=0 and sa=0 then
    eb2.state=lightstateoff
    eb1.state=lightstateon
    end if
    end if
    end if
    if points=10 or points=100 then
    matchnumb=matchnumb+1
    if matchnumb=10 then matchnumb=0
    end if
    if points = 10 or points = 100 or points = 1000 then scn=1
    if points = 100 then
    If showdt = true then reel(currpl).addvalue(100)
    bell=100
    end if
    if points = 10 then
    If showdt = true then reel(currpl).addvalue(10)
    bell=10
    end if
    if points = 1000 then
    If showdt = true then reel(currpl).addvalue(1000)
    bell=1000
    end if
    if points = 3000 then
    If showdt = true then reel(currpl).addvalue(3000)
    scn=3
    bell=1000
    end if
    if points = 2000 then
    If showdt = true then reel(currpl).addvalue(2000)
    scn=2
    bell=1000
    end if
    if points = 500 then
    If showdt = true then reel(currpl).addvalue(500)
    scn=5
    bell=100
    end if
    scn1=0
    scntimer.enabled=true
    score(currpl)=score(currpl)+points
    truesc(currpl)=truesc(currpl)+points
    if score(currpl)>99999 then
    score(currpl)=score(currpl)-100000
    rep(currpl)=0
    end if
    if score(currpl)=>replay1 and rep(currpl)=0 then
    credit=credit+1
    playsound "knocker"
    if credit>24 then credit=24
    If showdt = true then credtxt.text=credit
    rep(currpl)=1
    playsound "click"
    end if
    if score(currpl)=>replay2 and rep(currpl)=1 then
    credit=credit+1
    playsound "knocker"
    if credit>24 then credit=24
    If showdt = true then credtxt.text=credit
    rep(currpl)=2
    playsound "click"
    end if
    if score(currpl)=>replay3 and rep(currpl)=2 then
    credit=credit+1
    playsound "knocker"
    if credit>24 then credit=24
    If showdt = true then credtxt.text=credit
    rep(currpl)=3
    playsound "click"
    end if
    end if
end sub

sub scntimer_timer
    scn1=scn1 + 1
    if bell=10 then playsound "bell10"
    if bell=100 then playsound "bell100"
    if bell=1000 then playsound "bell1000"
    if scn1=scn then scntimer.enabled=false
end sub

Sub CheckTilt
  If Tilttimer.Enabled = True Then
  TiltSens = TiltSens + 1
  if TiltSens = 2 Then
  Tilt = True
  If showdt = true then tilttxt.text="TILT"
  If showdt = false then Controller.B2SSetTilt 33,1
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
  gioff
  Light17.State = 0
  Light19.State = 0
  Light25.State = 0
  Light26.State = 0
    bumper1.force=0
    bumper2.force=0
  Wall35.IsDropped = 0
  Wall36.IsDropped = 0
    k3k.state=lightstateoff
    tabl.state=lightstateoff
    snl.state=lightstateoff
    abbon1.isdropped=true:rolloverprim4.image = "Red peg1"
    abbon2.isdropped=true:rolloverprim7.image = "Red peg1"
    for i=1 to 8
    tl(i).state=lightstateoff
    tla(i).state=lightstateoff
    next
    l1.state=lightstateoff
    t1kl.state=lightstateoff
    mabl.state=lightstateoff
    for i= 1 to 10
    b(i).state=lightstateoff
    next
    dbl.state=lightstateoff
    al.state=lightstateoff
    bl.state=lightstateoff
    cl.state=lightstateoff
    dl.state=lightstateoff
    shootagain.state=lightstateoff
  If showdt = false then Controller.B2SSetShootAgain 36, 0
    eb1.state=lightstateoff
    eb2.state=lightstateoff
    sp1.state=lightstateoff
    sp2.state=lightstateoff
end sub

sub credtimer_timer
end sub

sub savehs
  ' Based on Black's Highscore routines
  Dim FileObj
  Dim ScoreFile
  Set FileObj=CreateObject("Scripting.FileSystemObject")
  If Not FileObj.FolderExists(UserDirectory) then
    Exit Sub
  End if
  Set ScoreFile=FileObj.CreateTextFile(UserDirectory & "Toledo.txt",True)
    ScoreFile.WriteLine score(1)
    scorefile.writeline score(2)
    scorefile.writeline credit
    scorefile.writeline matchnumb
'   scorefile.writeline HSAHighScore
    ScoreFile.Close
  Set ScoreFile=Nothing
  Set FileObj=Nothing
end sub

sub loadhs
    ' Based on Black's Highscore routines
  Dim FileObj
  Dim ScoreFile
    dim temp1
    dim temp2
    dim temp3
    dim temp4
    dim temp5
    Set FileObj=CreateObject("Scripting.FileSystemObject")
  If Not FileObj.FolderExists(UserDirectory) then
    Exit Sub
  End if
  If Not FileObj.FileExists(UserDirectory & "Toledo.txt") then
    Exit Sub
  End if
  Set ScoreFile=FileObj.GetFile(UserDirectory & "Toledo.txt")
  Set TextStr=ScoreFile.OpenAsTextStream(1,0)
    If (TextStr.AtEndOfStream=True) then
      Exit Sub
    End if
    temp1=TextStr.ReadLine
        temp2=textstr.readline
        temp3=textstr.readline
        temp4=textstr.readline
'        temp5=textstr.readline
    TextStr.Close
      score(1) = CDbl(temp1)
      score(2) = CDbl(temp2)
      credit= CDbl(temp3)
      matchnumb= CDbl(temp4)
'     HSAHighScore=cdbl(temp5)
      Set ScoreFile=Nothing
      Set FileObj=Nothing
end sub

Dim RStep, Lstep

Sub RightSlingShot_Slingshot
  if tilt=false then
    addscore 10
    PlaySoundAtVol "fx_slingshot", ActiveBall, 1
    RSling.Visible = 0
    RSling1.Visible = 1
    sling1.TransZ = -20
    RStep = 0
    RightSlingShot.TimerEnabled = 1
  End If
End Sub

Sub RightSlingShot_Timer
    Select Case RStep
        Case 3:RSLing1.Visible = 0:RSLing2.Visible = 1:sling1.TransZ = -10
        Case 4:RSLing2.Visible = 0:RSLing.Visible = 1:sling1.TransZ = 0:RightSlingShot.TimerEnabled = 0
    End Select
    RStep = RStep + 1
End Sub
Sub LeftSlingShot_Slingshot
  if tilt=false then
    PlaySoundAtVol "fx_slingshot", ActiveBall, 1
    addscore 10
    LSling.Visible = 0
    LSling1.Visible = 1
    sling2.TransZ = -20
    LStep = 0
    LeftSlingShot.TimerEnabled = 1
  End If
End Sub

Sub LeftSlingShot_Timer
    Select Case LStep
        Case 3:LSLing1.Visible = 0:LSLing2.Visible = 1:sling2.TransZ = -10
        Case 4:LSLing2.Visible = 0:LSLing.Visible = 1:sling2.TransZ = 0:LeftSlingShot.TimerEnabled = 0
    End Select
    LStep = LStep + 1
End Sub


Sub Bumper3_Hit
    addscore 10
    if tilt=false then
    PlaySoundAtVol "Bumper", ActiveBall, 1
  End If
End Sub


Sub Bumper4_Hit
    addscore 10
    if tilt=false then
    PlaySoundAtVol "Bumper", ActiveBall, 1

  End If
End Sub

sub hole_hit
    if (k3k.state)=lightstateon then
    addscore 3000
    else
    addscore 500
    end if
    if (tabl.state)=lightstateon then bonadv
    if (snl.state)=lightstateon then advnum
    PlaySoundAtVol "kickerkick", ActiveBall, 1
    kicktimer.enabled=true
end sub

sub kicktimer_timer
    hole.kick 200,7
    kicktimer.enabled=false
end sub

sub a500_hit
    addscore 500
  rolloverprim5.transz = -4
end sub
Sub a500_UnHit
    rolloverprim5.transz = 0
End Sub
sub b500_hit
    addscore 500
  rolloverprim6.transz = -4
end sub

Sub b500_Unhit
  rolloverprim6.transz = 0
End Sub

sub ab1001_hit
  rolloverprim4.transz = -4
    if (abbon1.isdropped)=false then bonadv
    addscore 100
end sub
Sub ab1001_Unhit
  rolloverprim4.transz = 0
End Sub

sub ab1002_hit
  rolloverprim7.transz = -4
    if (abbon2.isdropped)=false then bonadv
    addscore 100
end sub

Sub ab1002_UnHit
  rolloverprim7.transz = 0
End Sub

sub ten1_hit
  rolloverprim3.transz = -4
    addscore 10
end sub

Sub ten1_UnHit
  rolloverprim3.transz = 0
End Sub

sub ten2_hit
  rolloverprim8.transz = -4
    addscore 10
end sub

Sub ten2_UnHit
  rolloverprim8.transz = 0
End Sub

sub t1_hit
  PlaySound "target", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0, AudioFade(ActiveBall)
  t1p.transx = -10
  Me.TimerEnabled = 1
    if (t1kl.state)=lightstateon then
    addscore 1000
    else
    if (l1a.state)=lightstateon then
    addscore 500
    else
    addscore 100
    end if
    end if
    if (l1.state)=lightstateon then
    targ=targ+1
    l1.state=lightstateoff
    l2.state=lightstateon
    l1a.state=lightstateon
    end if
    if (mabl.state)=lightstateon then bonadv
    checkaward
end sub

Sub t1_Timer
  t1p.transx = 0
  Me.TimerEnabled = 0
End Sub

sub t2_hit
  PlaySound "target", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0, AudioFade(ActiveBall)
  t2p.transx = -10
  Me.TimerEnabled = 1
    if (t1kl.state)=lightstateon then
    addscore 1000
    else
    if (l2a.state)=lightstateon then
    addscore 500
    else
    addscore 100
    end if
    end if
    if (l2.state)=lightstateon then
    targ=targ+1
    l2.state=lightstateoff
    l3.state=lightstateon
    l2a.state=lightstateon
    end if
    if (mabl.state)=lightstateon then bonadv
    checkaward
end sub

Sub t2_Timer
  t2p.transx = 0
  Me.TimerEnabled = 0
End Sub


sub t3_hit
  PlaySound "target", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0, AudioFade(ActiveBall)
  t3p.transx = -10
  Me.TimerEnabled = 1
    if (t1kl.state)=lightstateon then
    addscore 1000
    else
    if (l3a.state)=lightstateon then
    addscore 500
    else
    addscore 100
    end if
    end if
    if (l3.state)=lightstateon then
    targ=targ+1
    l3.state=lightstateoff
    l4.state=lightstateon
    l3a.state=lightstateon
    end if
    if (mabl.state)=lightstateon then bonadv
    checkaward
end sub

Sub t3_Timer
  t3p.transx = 0
  Me.TimerEnabled = 0
End Sub

sub t4_hit
  PlaySound "target", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0, AudioFade(ActiveBall)
  t4p.transx = -10
  Me.TimerEnabled = 1
    if (t1kl.state)=lightstateon then
    addscore 1000
    else
    if (l4a.state)=lightstateon then
    addscore 500
    else
    addscore 100
    end if
    end if
    if (l4.state)=lightstateon then
    targ=targ+1
    l4.state=lightstateoff
    l5.state=lightstateon
    l4a.state=lightstateon
    end if
    if (mabl.state)=lightstateon then bonadv
    checkaward
end sub

Sub t4_Timer
  t4p.transx = 0
  Me.TimerEnabled = 0
End Sub


sub t5_hit
  PlaySound "target", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0, AudioFade(ActiveBall)
  t5p.transx = -10
  Me.TimerEnabled = 1
    if (t1kl.state)=lightstateon then
    addscore 1000
    else
    if (l5a.state)=lightstateon then
    addscore 500
    else
    addscore 100
    end if
    end if
    if (l5.state)=lightstateon then
    targ=targ+1
    l5.state=lightstateoff
    l6.state=lightstateon
    l5a.state=lightstateon
    end if
    if (mabl.state)=lightstateon then bonadv
    checkaward
end sub

Sub t5_Timer
  t5p.transx = 0
  Me.TimerEnabled = 0
End Sub

sub t6_hit
  PlaySound "target", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0, AudioFade(ActiveBall)
  t6p.transx = -10
  Me.TimerEnabled = 1
    if (t1kl.state)=lightstateon then
    addscore 1000
    else
    if (l6a.state)=lightstateon then
    addscore 500
    else
    addscore 100
    end if
    end if
    if (l6.state)=lightstateon then
    targ=targ+1
    l6.state=lightstateoff
    l7.state=lightstateon
    l6a.state=lightstateon
    end if
    if (mabl.state)=lightstateon then bonadv
    checkaward
end sub

Sub t6_Timer
  t6p.transx = 0
  Me.TimerEnabled = 0
End Sub

sub t7_hit
  PlaySound "target", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0, AudioFade(ActiveBall)
  t7p.transx = -10
  Me.TimerEnabled = 1
    if (t1kl.state)=lightstateon then
    addscore 1000
    else
    if (l7a.state)=lightstateon then
    addscore 500
    else
    addscore 100
    end if
    end if
    if (l7.state)=lightstateon then
    targ=targ+1
    l7.state=lightstateoff
    l8.state=lightstateon
    l7a.state=lightstateon
    end if
    if (mabl.state)=lightstateon then bonadv
    checkaward
end sub

Sub t7_Timer
  t7p.transx = 0
  Me.TimerEnabled = 0
End Sub

sub t8_hit
  PlaySound "target", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0, AudioFade(ActiveBall)
  t8p.transx = -10
  Me.TimerEnabled = 1
    if (t1kl.state)=lightstateon then
    addscore 1000
    else
    if (l8a.state)=lightstateon then
    addscore 500
    else
    addscore 100
    end if
    end if
    if (l8.state)=lightstateon then
    targ=targ+1
    l8.state=lightstateoff
    t1kl.state=lightstateon
    l8a.state=lightstateon
    end if
    if (mabl.state)=lightstateon then bonadv
    checkaward
end sub

Sub t8_Timer
  t8p.transx = 0
  Me.TimerEnabled = 0
End Sub

sub lt1_hit
  PlaySound "target", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0, AudioFade(ActiveBall)
  rt1p.transx = -10
  Me.TimerEnabled = 1
    addscore 500
    al.state=lightstateon
    checkaward
end sub

Sub lt1_Timer
  lt1p.transx = 0
  Me.TimerEnabled = 0
End Sub

sub lt2_hit
  PlaySound "target", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0, AudioFade(ActiveBall)
  lt2p.transx = -10
  Me.TimerEnabled = 1
    addscore 500
    cl.state=lightstateon
    checkaward
end sub

Sub lt2_Timer
  lt2p.transx = 0
  Me.TimerEnabled = 0
End Sub

sub rt1_hit
  PlaySound "target", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0, AudioFade(ActiveBall)
  rt1p.transx = -10
  Me.TimerEnabled = 1
    addscore 500
    bl.state=lightstateon
    checkaward
end sub

Sub rt1_Timer
  rt1p.transx = 0
  Me.TimerEnabled = 0
End Sub

sub rt2_hit
  PlaySound "target", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0, AudioFade(ActiveBall)
  rt2p.transx = -10
  Me.TimerEnabled = 1
    addscore 500
    dl.state=lightstateon
    checkaward
end sub

Sub rt2_Timer
  rt2p.transx = 0
  Me.TimerEnabled = 0
End Sub

sub ab1_hit
  rolloverprim1.transz = -4
    if tilt=false then PlaySoundAtVol "clerker", ActiveBall, 1
    bonadv
end sub

Sub ab1_UnHit
  rolloverprim1.transz = 0
End Sub

sub ab2_hit
  rolloverprim2.transz = -4
    if tilt=false then PlaySoundAtVol "clerker", ActiveBall, 1
    bonadv
end sub

Sub ab2_Unhit
  rolloverprim2.transz = 0
End Sub

sub lin_hit
    if (eb1.state)=lightstateon and tilt=false then
    eb1.state=lightstateoff
    eb2.state=lightstateoff
    shootagain.state=lightstateon
  If showdt = false then Controller.B2SSetShootAgain 36, 1
    sa=1
    end if
    addscore 500
end sub

sub rin_hit
    if (eb2.state)=lightstateon and tilt=false then
    eb1.state=lightstateoff
    eb2.state=lightstateoff
    shootagain.state=lightstateon
  If showdt = false then Controller.B2SSetShootAgain 36, 1
    sa=1
    end if
    addscore 500
end sub

sub lout_hit
    if (sp1.state)=lightstateon and tilt=false then
    credit=credit+1
    playsound "knocker"
    if credit>24 then credit=24
    If showdt = true then credtxt.text=credit
    playsound "click"
    end if
    bonadv
    addscore 1000
end sub

sub rout_hit
    if (sp2.state)=lightstateon and tilt=false then
    credit=credit+1
    playsound "knocker"
    if credit>24 then credit=24
    If showdt = true then credtxt.text=credit
    playsound "click"
    end if
    bonadv
    addscore 1000
end sub

sub sn_hit
  rolloverprim.transz = -5
    addscore 10
    advnum
end sub

Sub sn_UnHit
  rolloverprim.transz = 0
End Sub
sub advnum
    if targ<8 then
    tl(targ+1).state=lightstateoff
    tla(targ+1).state=lightstateon
    targ=targ+1
    end if
    if targ<8 then tl(targ+1).state=lightstateon
    checkaward
end sub

sub checkaward
    if targ=>5 and sa=0 and (eb1.state)=lightstateoff and (eb2.state)=lightstateoff then eb2.state=lightstateon
    if targ=>8 and (sp1.state)=lightstateoff and (sp2.state)=lightstateoff then sp1.state=lightstateon
    if targ=>8 and (al.state)=lightstateon and (bl.state)=lightstateon then mabl.state=lightstateon
    if (al.state)=lightstateon and (bl.state)=lightstateon then dbl.state=lightstateon
    if sa=0 and (al.state)=lightstateon and (bl.state)=lightstateon and (cl.state)=lightstateon and (dl.state)=lightstateon then
    fulleb=1
    eb1.state=lightstateon
    eb2.state=lightstateon
    k3k.state=lightstateon
    end if
end sub

sub ballhome_hit
    bb1.isdropped=true
    ballrelease=1
end sub

sub ballrel_hit
    if ballrelease=1 then
    PlaySoundAtVol "launchball", ActiveBall, 1
    ballrelease=0
    end if
end sub

sub ballout_hit
    bb1.isdropped=false
    bb2.isdropped=false
end sub

sub ballapr_hit
    bb2.isdropped=true
end sub

sub ballout2_hit
    bb2.isdropped=false
end sub

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
'            Supporting Ball, Sound Functions and Math
'*********************************************************************

Function RndNum(min, max)
  RndNum = Int(Rnd() * (max-min + 1) ) + min ' Sets a random number between min and max
End Function

Const Pi = 3.1415927

Function dSin(degrees)
  dsin = sin(degrees * Pi/180)
End Function

Function dCos(degrees)
  dcos = cos(degrees * Pi/180)
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
  Vol = Csng(BallVel(ball) ^2 / 2000)
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
				  ' PlaySound("fx_ballrolling" & b), -1, BallRollVol(BOT(b) )*.8, AudioPan(BOT(b) ), 0, Pitch(BOT(b) ), 1, 0, AudioFade(BOT(b) )
        Else ' Ball on raised ramp
          PlaySound("fx_ballrolling" & b), -1, Vol(BOT(b) )*.5, AudioPan(BOT(b) ), 0, Pitch(BOT(b) )+50000, 1, 0, AudioFade(BOT(b) )
				  ' PlaySound("fx_ballrolling" & b), -1, BallRollVol(BOT(b) )*.2, AudioPan(BOT(b) ), 0, Pitch(BOT(b) )+50000, 1, 0, AudioFade(BOT(b) )
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
    PlaySound("fx_collide"), 0, Csng(velocity) ^2 / 2000, AudioPan(ball1), 0, Pitch(ball1), 0, 0, AudioFade(ball1)
End Sub

'************************************
' What you need to add to your table
'************************************

' a timer called CollisionTimer. With a fast interval, like from 1 to 10
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

' The Double For loop: This is a double cycle used to check the collision between a ball and the other ones.
' If you look at the parameters of both cycles, youll notice they are designed to avoid checking
' collision twice. For example, I will never check collision between ball 2 and ball 1,
' because I already checked collision between ball 1 and 2. So, if we have 4 balls,
' the collision checks will be: ball 1 with 2, 1 with 3, 1 with 4, 2 with 3, 2 with 4 and 3 with 4.

' Sum first the radius of both balls, and if the height between them is higher then do not calculate anything more,
' since the balls are on different heights so they can't collide.

' The next 3 lines calculates distance between xth and yth balls with the Pytagorean theorem,

' The first "If": Checking if the distance between the two balls is less than the sum of the radius of both balls,
' and both balls are not already colliding.

' Why are we checking if balls are already in collision?
' Because we do not want the sound repeting when two balls are resting closed to each other.

' Set the collision property of both balls to True, and we assign the number of the ball colliding

' Play the collide sound of your choice using the VOL, PAN and PITCH functions

' Last lines: If the distance between 2 balls is more than the radius of a ball,
' then there is no collision and then set the collision property of the ball to False (-1).

Sub Pins_Hit (idx)
  PlaySound "pinhit_low", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0
End Sub

'Sub atargets_Hit (idx)
' PlaySound "target", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0
'End Sub

Sub metals_Hit (idx)
  PlaySound "metalhit_thin", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0
End Sub

Sub Metals_Medium_Hit (idx)
  PlaySound "metalhit_medium", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0
End Sub

Sub metals2_Hit (idx)
  PlaySound "metalhit2", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0
End Sub

Sub agates_Hit (idx)
  PlaySound "gate4", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0
End Sub

'Sub Spinner_Spin
' PlaySound "fx_spinner",0,.25,0,0.25
'End Sub

Sub arubbers_Hit(idx)
  dim finalspeed
    finalspeed=SQR(activeball.velx * activeball.velx + activeball.vely * activeball.vely)
  If finalspeed > 20 then
    PlaySound "fx_rubber2", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0
  End if
  If finalspeed >= 6 AND finalspeed <= 20 then
    RandomSoundRubber()
  End If
End Sub

Sub aposts_Hit(idx)
  dim finalspeed
    finalspeed=SQR(activeball.velx * activeball.velx + activeball.vely * activeball.vely)
  If finalspeed > 16 then
    PlaySound "fx_rubber2", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0
  End if
  If finalspeed >= 6 AND finalspeed <= 16 then
    RandomSoundRubber()
  End If
End Sub

Sub RandomSoundRubber()
  Select Case Int(Rnd*3)+1
    Case 1 : PlaySound "rubber_hit_1", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0
    Case 2 : PlaySound "rubber_hit_2", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0
    Case 3 : PlaySound "rubber_hit_3", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0
  End Select
End Sub

Sub LeftFlipper_Collide(parm)
  RandomSoundFlipper()
End Sub

Sub RightFlipper_Collide(parm)
  RandomSoundFlipper()
End Sub

Sub upl_Collide(parm)
  RandomSoundFlipper()
End Sub

Sub upr_Collide(parm)
  RandomSoundFlipper()
End Sub

Sub RandomSoundFlipper()
  Select Case Int(Rnd*3)+1
    Case 1 : PlaySound "flip_hit_1", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0
    Case 2 : PlaySound "flip_hit_2", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0
    Case 3 : PlaySound "flip_hit_3", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0
  End Select
End Sub

Sub playuptimer_timer
    if currpl = 1 then up1.text = ">":up2.text = ""
    if currpl = 2 then up2.text = "<":up1.text = ""
    if currpl = 0 then up1.text = "":up2.text = ""
End Sub
Sub UpdateFlipperLogos
    Flipperbatleft.RotAndTra8 = upl.CurrentAngle + 90
    Flipperbatright.RotAndTra8 = upr.CurrentAngle + 90
  leftflipprim.RotAndTra8 = LeftFlipper.CurrentAngle
  rightflipprim.RotAndTra8 = RightFlipper.CurrentAngle
End Sub
Sub FlippersTimer_Timer()
  UpdateFlipperLogos
  GateHeavyT1.RotZ = -(Gate.currentangle)
  GateHeavyT2.RotZ = -(Gate1.currentangle)
  If showdt = false then B2SSetScore
End Sub

Sub B2SSetScore
    Controller.B2SSetScorePlayer1 score(1)
    Controller.B2SSetScorePlayer2 score(2)
    Controller.B2SSetScorePlayer3 credit
    Controller.B2SSetBallInPlay 32, ballinplay
    Controller.B2SSetCanPlay 31, playno
    Controller.B2SSetPlayerUp 30, currpl
  If truesc(1)> 99999 then
    Controller.B2SSetScoreRollover 25 ,1
  Else
    Controller.B2SSetScoreRollover 25 ,0
  End If
  If truesc(2)> 99999 then
    Controller.B2SSetScoreRollover 26 ,1
  Else
    Controller.B2SSetScoreRollover 26 ,0
  End If
End Sub

Const HighScoreFilename = "ToledoHighscorePostIt.txt"

Dim HSAHighScore, HSA1, HSA2, HSA3
Dim HSScore100k, HSScore10k, HSScoreK, HSScore100, HSScore10, HSScore1, HSScorex  'Define 5 different score values for each reel to use
Dim HSEnterMode, hsLetterFlash, hsEnteredDigits(3), hsCurrentDigit, hsCurrentLetter
Const hsFlashDelay = 4
Const DefaultHighScore = 100
Const DefaultHSA1 = 7
Const DefaultHSA2 = 20
Const DefaultHSA3 = 24

LoadHighScore
Sub LoadHighScore
  Dim FileObj
  Dim ScoreFile
  Dim TextStr
    Dim SavedDataTemp3 'HighScore
    Dim SavedDataTemp4 'HSA1
    Dim SavedDataTemp5 'HSA2
    Dim SavedDataTemp6 'HSA3
    Set FileObj=CreateObject("Scripting.FileSystemObject")
  If Not FileObj.FolderExists(UserDirectory) then
    Exit Sub
  End if
  If Not FileObj.FileExists(UserDirectory & HighScoreFilename) then
    SetDefaultHSTD:UpdatePostIt:SaveHighScore
    Exit Sub
  End if
  Set ScoreFile=FileObj.GetFile(UserDirectory & HighScoreFilename)
  Set TextStr=ScoreFile.OpenAsTextStream(1,0)
    If (TextStr.AtEndOfStream=True) then
      SetDefaultHSTD:UpdatePostIt:SaveHighScore
      Exit Sub
    End if
    SavedDataTemp3=Textstr.ReadLine ' HighScore
    SavedDataTemp4=Textstr.ReadLine ' HSA1
    SavedDataTemp5=Textstr.ReadLine ' HSA2
    SavedDataTemp6=Textstr.ReadLine ' HSA3
    TextStr.Close
    HSAHighScore=SavedDataTemp3
    HSA1=SavedDataTemp4
    HSA2=SavedDataTemp5
    HSA3=SavedDataTemp6
    UpdatePostIt
      Set ScoreFile = Nothing
      Set FileObj = Nothing
End Sub

Sub SetDefaultHSTD  'bad data or missing file - reset and resave
  HSAHighScore = DefaultHighScore
  HSA1 = DefaultHSA1
  HSA2 = DefaultHSA2
  HSA3 = DefaultHSA3
  SaveHighScore
End Sub

Sub LastScoreReels
    HSScorex = LastScore
    HSScore100K=Int (HSScorex/100000)'Calculate the value for the 100,000's digit
    HSScore10K=Int ((HSScorex-(HSScore100k*100000))/10000) 'Calculate the value for the 10,000's digit
    HSScoreK=Int((HSScorex-(HSScore100k*100000)-(HSScore10K*10000))/1000) 'Calculate the value for the 1000's digit
    HSScore100=Int((HSScorex-(HSScore100k*100000)-(HSScore10K*10000)-(HSScoreK*1000))/100) 'Calculate the value for the 100's digit
    HSScore10=Int((HSScorex-(HSScore100k*100000)-(HSScore10K*10000)-(HSScoreK*1000)-(HSScore100*100))/10) 'Calculate the value for the 10's digit
    HSScore1=Int(HSScorex-(HSScore100k*100000)-(HSScore10K*10000)-(HSScoreK*1000)-(HSScore100*100)-(HSScore10*10)) 'Calculate the value for the 1's digit
End Sub

Sub UpdatePostIt
    HSScorex = HSAHighScore
    HSScore100K=Int (HSScorex/100000)'Calculate the value for the 100,000's digit
    HSScore10K=Int ((HSScorex-(HSScore100k*100000))/10000) 'Calculate the value for the 10,000's digit
    HSScoreK=Int((HSScorex-(HSScore100k*100000)-(HSScore10K*10000))/1000) 'Calculate the value for the 1000's digit
    HSScore100=Int((HSScorex-(HSScore100k*100000)-(HSScore10K*10000)-(HSScoreK*1000))/100) 'Calculate the value for the 100's digit
    HSScore10=Int((HSScorex-(HSScore100k*100000)-(HSScore10K*10000)-(HSScoreK*1000)-(HSScore100*100))/10) 'Calculate the value for the 10's digit
    HSScore1=Int(HSScorex-(HSScore100k*100000)-(HSScore10K*10000)-(HSScoreK*1000)-(HSScore100*100)-(HSScore10*10)) 'Calculate the value for the 1's digit
    EMReelHSNum1.SetValue(HSScore100K):If HSScorex<100000 Then EMReelHSNum1.SetValue 10
    EMReelHSNum2.SetValue(HSScore10K):If HSScorex<10000 Then EMReelHSNum2.SetValue 10
    EMReelHSNum3.SetValue(HSScoreK):If HSScorex<1000 Then EMReelHSNum3.SetValue 10:EMReelHSComma.SetValue 0:Else EMReelHSComma.SetValue 1
    EMReelHSNum4.SetValue(HSScore100):If HSScorex<100 Then EMReelHSNum4.SetValue 10
    EMReelHSNum5.SetValue(HSScore10):If HSScorex<10 Then EMReelHSNum5.SetValue 10
    EMReelHSNum6.SetValue(HSScore1)
    EMReelHSName1.SetValue HSA1
    EMReelHSName2.SetValue HSA2
    EMReelHSName3.SetValue HSA3
    dthigh.text = HSAHighScore
    dtinitial1.setvalue HSA1
    dtinitial2.setvalue HSA2
    dtinitial3.setvalue HSA3
End Sub

Sub SaveHighScore
  Dim FileObj
  Dim ScoreFile
  Set FileObj=CreateObject("Scripting.FileSystemObject")
  If Not FileObj.FolderExists(UserDirectory) then
    Exit Sub
  End if
  Set ScoreFile=FileObj.CreateTextFile(UserDirectory & HighScoreFilename,True)
    ScoreFile.WriteLine HSAHighScore
    ScoreFile.WriteLine HSA1
    ScoreFile.WriteLine HSA2
    ScoreFile.WriteLine HSA3
    ScoreFile.Close
  Set ScoreFile=Nothing
  Set FileObj=Nothing
End Sub

Sub HighScoreEntryInit()
  HSEnterMode = True
  hsCurrentDigit = 0
  hsCurrentLetter = 1:HSA1=1
  HighScoreFlashTimer.Interval = 250
  HighScoreFlashTimer.Enabled = True
  hsLetterFlash = hsFlashDelay
End Sub

Sub HighScoreFlashTimer_Timer()
  hsLetterFlash = hsLetterFlash-1
  If hsLetterFlash=1 then 'switch to underscore
    Select Case hsCurrentLetter
      Case 1:
        EMReelHSName1.SetValue 28:dtinitial1.setvalue 28
      Case 2:
        EMReelHSName2.SetValue 28:dtinitial2.setvalue 28
      Case 3:
        EMReelHSName3.SetValue 28:dtinitial3.setvalue 28
    End Select
  End If
  If hsLetterFlash=0 then 'switch back
    hsLetterFlash = hsFlashDelay
    Select Case hsCurrentLetter
      Case 1:
        EMReelHSName1.SetValue HSA1:dtinitial1.setvalue HSA1
      Case 2:
        EMReelHSName2.SetValue HSA2:dtinitial2.setvalue HSA2
      Case 3:
        EMReelHSName3.SetValue HSA3:dtinitial3.setvalue HSA3
    End Select
  End If
End Sub

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
        HSA1=HSA1+1:If HSA1>27 Then HSA1=0
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
        EMReelHSName1.SetValue HSA1:EMReelHSName2.SetValue HSA2
        dtinitial1.setvalue HSA1::dtinitial2.setvalue HSA2
      Case 2:
        If HSA2=27 Then 'bksp
          HSA2=0
          hsCurrentLetter=1
        Else
          hsCurrentLetter=3 'enter it
          HSA3=HSA2 'start at same alphabet spot
        End If
        EMReelHSName2.SetValue HSA2:EMReelHSName3.SetValue HSA3
        dtinitial2.setvalue HSA2:dtinitial3.setvalue HSA3
      Case 3:
        If HSA3=27 Then 'bksp
          HSA3=0
          hsCurrentLetter=2
        Else
          SaveHighScore 'enter it
          HighScoreFlashTimer.Enabled = False
          HSEnterMode = False

          UpdatePostIt
          EMReelHSTitle.SetValue 0
        End If
        EMReelHSName3.SetValue HSA3
        dtinitial3.setvalue HSA3
    End Select
    End If
End Sub

Function PostItHighScoreCheck (keycode)
  PostItHighScoreCheck = 0
    If HSEnterMode Then
    HighScoreProcessKey(keycode)

    Select Case keycode
      Case LeftFlipperKey, RightFlipperKey, 2, StartGameKey
        PostItHighScoreCheck = 1
    End Select
  End If
End Function

Sub CheckNewHighScorePostIt (newScore)
    If CLng(newScore) > CLng(HSAHighScore) Then
      HSAHighScore=newScore:HSA1 = 0:HSA2 = 0:HSA3 = 0:UpdatePostIt
      EMReelHSTitle.SetValue 1
      HighScoreEntryInit()
    End If
End Sub
Sub CheckNewHighScorePostIt1Player (newScore1)
  CheckNewHighScorePostIt newScore1
End Sub
Sub CheckNewHighScorePostIt2Player (newScore1, newScore2)
  Dim bestscore
  bestscore = newScore1
  If newScore2 > bestscore then bestscore = newScore2
  CheckNewHighScorePostIt bestscore
End Sub
Sub CheckNewHighScorePostIt3Player (newScore1, newScore2, newScore3)
  Dim bestscore
  bestscore = newScore1
  If newScore2 > bestscore then bestscore = newScore2
  If newScore3 > bestscore then bestscore = newScore3
  CheckNewHighScorePostIt bestscore
End Sub
Sub CheckNewHighScorePostIt4Player (newScore1, newScore2, newScore3, newScore4)
  Dim bestscore
  bestscore = newScore1
  If newScore2 > bestscore then bestscore = newScore2
  If newScore3 > bestscore then bestscore = newScore3
  If newScore4 > bestscore then bestscore = newScore4
  CheckNewHighScorePostIt bestscore
End Sub

' Thalamus : Exit in a clean and proper way
Sub Table1_exit
  Controller.Pause = False
  Controller.Stop
End Sub

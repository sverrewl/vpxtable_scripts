Const cGameName = "aztec_1976"
'Const ReflectionMod=True     'enable JPJ reflection mod

' Thalamus 2018-07-19
' Added/Updated "Positional Sound Playback Functions" and "Supporting Ball & Sound Functions"
' No special SSF tweaks yet.

Dim Score(4)
dim truesc(4)
dim ballrelenabled
dim state
dim playno
dim credit
dim eg
dim hisc1
dim hisc2
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
dim cred
dim scn
dim scn1
dim bell
dim points
dim tempscore
dim digit
dim update
dim reel(4)
dim replay1
dim replay2
dim replay3
dim up(4)
dim bl(10)
dim bv
dim balls
dim azlet(5)
dim az1(5)
dim az2(5)
dim sm
dim br
dim bumperchoice
dim rstep, rstepa, rstepb, rstepc, rstepd, rstepe, lstep, lstepa, lstepb, lstepc, lstepd
Dim GlobalSoundLevel
Dim coin, coinc

GlobalSoundLevel = 2



On Error Resume Next
ExecuteGlobal GetTextFile("controller.vbs")
If Err Then MsgBox "You need the controller.vbs in order to run this table, available in the vp10 package"
On Error Goto 0

'************************ Table ini ****************************
sub table1_init

'*********** JPJ's automatic random rotation of screws *******
  dim z
    For Each x in AutomaticRot
  z = Int(Rnd*360)+1
    x.RotZ = z
  Next
    For Each x in AutomaticRot2
  z = Int(Rnd*360)+1
    x.RotY = z
  Next
'**************************************************************

  LoadEM
  If ShowDT=false Then
    for each obj in DTItems
      obj.visible=False
    Next
  end if
rmod.enabled = 1

' If ShowDT=false Then
'   for each obj in DTItems
'     obj.visible=False
'   Next
' Primitive78.visible = 0
' Primitive77.visible = 0
' Primitive81.visible = 0
' Primitive82.visible = 0
' Primitive83.visible = 0
' Primitive84.visible = 0
' RampLBlack.visible = 1
' RampRBlack.visible = 1
' RampLBlack1.visible = 1
' RampRBlack1.visible = 1
'   Else
'
' RampLBlack.visible = 0
' RampRBlack.visible = 0
' RampLBlack1.visible = 0
' RampRBlack1.visible = 0
'
' end If

    set plno(1)=plno1
    set plno(2)=plno2
    set plno(3)=plno3
    set plno(4)=plno4
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
    set bl(1)=b5k
    set bl(2)=b10k
    set bl(3)=b15k
    set bl(4)=b20k
    set bl(5)=b25k
    set bl(6)=b30k
    set bl(7)=b35k
    set bl(8)=b40k
    set bl(9)=b45k
    set bl(10)=b50k
    set azlet(1)=upcl
    set azlet(2)=midll
    set azlet(3)=midrl
    set azlet(4)=lowll
    set azlet(5)=lowrl
    set az1(1)=al1
    set az1(2)=zl1
    set az1(3)=tl1
    set az1(4)=el1
    set az1(5)=cl1
    set az2(1)=al2
    set az2(2)=zl2
    set az2(3)=tl2
    set az2(4)=el2
    set az2(5)=cl2
    set reel(1)=reel1
    set reel(2)=reel2
    set reel(3)=reel3
    set reel(4)=reel4
    credtimer.enabled=true
    loadhs
    if hisc1="" then hisc1=300250
    if hisc2="" then hisc2=450250
    if credit="" then credit=0
    credtxt.text=credit
'************************************************************************
'**  Changing number of balls, 3 or 5 manually in the script           **
'**  Or Using Left Magna Button before starting game                   **
   balls=3        '                                              **
'************************************************************************

  apron3d.image = "apron01"
    ballstxt.text=balls
    if balls=3 then
    replay1=300000
    replay2=450000
    replay3=575000
    if balls=3 then hstxt.text=hisc1
    else
    replay1=450000
    replay2=600000
    replay3=850000
    if balls=5 then hstxt.text=hisc2
    end if
'    rep1.text=replay1
'    rep2.text=replay2
'    rep3.text=replay3
    select case(matchnumb)
    case 0:
    m0.text="00"
  if b2son Then
    Controller.b2ssetdata 34,1
  end if
    case 1:
    m1.text="10"
  if b2son Then
    Controller.b2ssetdata 34,2
  end if
    case 2:
    m2.text="20"
  if b2son Then
    Controller.b2ssetdata 34,3
  end if
    case 3:
    m3.text="30"
  if b2son Then
    Controller.b2ssetdata 34,4
  end if
    case 4:
    m4.text="40"
  if b2son Then
    Controller.b2ssetdata 34,5
  end if
    case 5:
    m5.text="50"
  if b2son Then
    Controller.b2ssetdata 34,6
  end if
    case 6:
    m6.text="60"
  if b2son Then
    Controller.b2ssetdata 34,7
  end if
    case 7:
    m7.text="70"
  if b2son Then
    Controller.b2ssetdata 34,8
  end if
    case 8:
    m8.text="80"
  if b2son Then
    Controller.b2ssetdata 34,9
  end if
    case 9:
    m9.text="90"
  if b2son Then
    Controller.b2ssetdata 34,10
  end if
    end select
    for i=1 to 4
    currpl=i
    reel(currpl).setvalue(score(currpl))
    next
    currpl=0

  coin = 1
'******** lights Ini ***********
    bv=1
turnoff
gioff

end sub
'**********************************************************

Sub Table1_KeyDown(ByVal keycode)
          if keycode=50 then dbl.state=lightstateon
  If keycode = PlungerKey Then
    Plunger.PullBack
  End If

  if keycode = leftmagnasave then
    If ballinplay<1 or ballinplay>balls then
  '    if keycode=59 and state=false then
      if balls=3 then
        balls=5
        apron3d.image = "apron02"
        ballstxt.text=balls
        replay1=450000
        replay2=600000
        replay3=850000
'       rep1.text=replay1
'       rep2.text=replay2
'       rep3.text=replay3
        hstxt.text=hisc2
      else
        balls=3
        apron3d.image = "apron01"
        ballstxt.text=balls
        replay1=300000
        replay2=450000
        replay3=575000
'       rep1.text=replay1
'       rep2.text=replay2
'       rep3.text=replay3
        hstxt.text=hisc1
      end if
    end if
  end if

    if keycode = addcreditkey then
  playsound "coin3"
  coindelay.enabled=true
  end if

    if keycode = AddCreditKey2 then
  playsound "coin3"
  coindelay2.enabled=true
  end if

  if keycode = StartGameKey and credit>0 and state=false and playno=0 then
  credit=credit-1
If credit < 1 Then DOF 128, DOFOff
  credtxt.text=credit
  GION
    eg=0
    playno=1
    currpl=1
    plno(playno).state=lightstateon
    play(currpl).state=lightstateon
    playsound "click"
    playsound "initialize"
    rst=0
    ballinplay=1
    resettimer.enabled=true
    end if

    if keycode = StartGameKey and credit>0 and state=true and playno>0 and playno<4 and ballinplay<2 then
    credit=credit-1
If credit < 1 Then DOF 128, DOFOff
    credtxt.text=credit
    plno(playno).state=lightstateoff
    playno=playno+1
    plno(playno).state=lightstateon
    playsound "click"
    end if

  if state=true and tilt=false then
  If keycode = LeftFlipperKey Then
    LeftFlipper.RotateToEnd
    PlaySound SoundFXDOF("FlipperUp",101,DOFOn,DOFFlippers)
  End If

  If keycode = RightFlipperKey Then
    RightFlipper.RotateToEnd
        PlaySound SoundFXDOF("FlipperUp",102,DOFOn,DOFFlippers)
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
    playsound "plunger"
  End If

  If keycode = LeftFlipperKey Then
    LeftFlipper.RotateToStart
    if tilt=false and state=true then PlaySound SoundFXDOF("FlipperDown",101,DOFOff,DOFFlippers)
  End If

  If keycode = RightFlipperKey Then
    RightFlipper.RotateToStart
    if tilt=false and state=true then PlaySound SoundFXDOF("FlipperDown",102,DOFOff,DOFFlippers)
  End If

End Sub

sub coindelay_timer
playsound "click"
credit=credit+1
DOF 128, DOFOn
if credit>25 then credit=25
credtxt.text=credit
  coindelay.enabled=false
end sub

sub coindelay2_timer
playsound "click"
credit=credit+1
DOF 128, DOFOn
if credit>25 then credit=25
credtxt.text=credit
coindelay3.enabled=true
coindelay2.enabled=false
end sub

sub coindelay3_timer
playsound "click"
credit=credit+1
DOF 128, DOFOn
if credit>25 then credit=25
credtxt.text=credit
coindelay4.enabled=true
coindelay3.enabled=false
end sub

sub coindelay4_timer
playsound "click"
credit=credit+1
DOF 128, DOFOn
if credit>25 then credit=25
credtxt.text=credit
coindelay4.enabled=false
end sub

sub resettimer_timer
    rst=rst+1
    for i=1 to 4
    reel(i).resettozero
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
'    credittxt.text="Ported To VP By Leon Spalding"
    state=true
    for i=1 to 4
    score(i)=0
    truesc(i)=0
    rep(i)=0
    next
    eg=0
    sm=0
    bip5.text=" "
    bip1.text="1"
    for i=0 to 9
    match(i).text=" "
    next
    tilttxt.text=" "
    gamov.text=" "
    bv=1
    bl(bv).state=lightstateon
    for i=1 to 5
    azlet(i).state=lightstateon
    az1(i).state=lightstateoff
    az2(i).state=lightstateoff
    next
    sp1.state=lightstateoff
    sp2.state=lightstateoff
    ebl.state=lightstateoff
    spinl.state=lightstateoff
    leftsidel.state=lightstateoff
    topl.state=lightstateoff
    lanekl.state=lightstateoff
    Bumper3l.state = 0
    Bumper2l.state = 0
    Bumper1l.state = 1
    uprl.state = 1
    upll.state = 0
    dbl.state=lightstateoff
    tilt=false
    tiltsens=0
    ballinplay=1
    nb.CreateBall.image="ballimage7"
  nb.kick 135,4
  DOF 113, DOFPulse
end sub

Sub Drain_Hit()
  Drain.DestroyBall
  DOF 114, DOFPulse
  if tilt=false then
  br=1
  for i=0 to 13
    StopSound("fx_ballrolling" & (i))
  Next
  bonuscount.interval=200
  bonuscount.enabled=true
  else
  for i=0 to 13
    StopSound("fx_ballrolling" & (i))
  Next
  nextball
    end if
end sub

sub bonuscount_timer
    if (dbl.state)=lightstateoff then bonuscount.interval=1000
    if bv>0 then
    if (dbl.state)=lightstateon then
    addscore 10000
    else
    addscore 5000
    end if
    end if
    if bv=<0 then
    nextball
    br=0
    bonuscount.enabled=false
    end if
end sub

sub bld
    bl(bv).state=lightstateoff
    if bv>1 then bl(bv-1).state=lightstateon
    bv=bv-1
end sub

sub nextball
  if tilt=true then
  tilt=false
  tilttxt.text=" "
' bumper1.force=7
'    bumper2.force=7
'    bumper3.force=7
    leftsling.slingshotstrength=6
    rightsling.slingshotstrength=6
  end if
  if (shootagain.state)=lightstateon then
  shootagain.state=lightstateoff
  playsound "kickerkick"
    newball
  ballreltimer.enabled=true
  else
  currpl=currpl+1
  end if
  if currpl>playno then
  ballinplay=ballinplay+1
  if ballinplay>balls then
  GIOFF
  playsound "motorleer"
  eg=1
  ballreltimer.enabled=true
  else
  if state=true and tilt=false then
  play(currpl-1).state=lightstateoff
  currpl=1
  play(currpl).state=lightstateon
  newball
  playsound "kickerkick"
  ballreltimer.enabled=true
  end if
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
  state=false
  plno(playno).state=lightstateoff
  play(currpl-1).state=lightstateoff
  playno=0
  gamov.text="GAME OVER"
  for i=1 to 4
  if balls=3 and truesc(i)>hisc1 then
  hisc1=truesc(i)
  hstxt.text=hisc1
  end if
  if balls=5 and truesc(i)>hisc2 then
  hisc2=truesc(i)
  hstxt.text=hisc2
  end if
  next
  savehs
  cred=0
  credtimer.enabled=true
  ballreltimer.enabled=false
    else
    nb.CreateBall.image="ballimage7"
  nb.kick 135,4
  DOF 113, DOFPulse
    ballreltimer.enabled=false
    end if
end sub

sub newball
    bv=1
    sm=0
    bl(bv).state=lightstateon
    for i=1 to 5
    azlet(i).state=lightstateon
    az1(i).state=lightstateoff
    az2(i).state=lightstateoff
    next
    Bumper3l.state = 0
    Bumper2l.state = 0
    Bumper1l.state = 1
    uprl.state = 1
    upll.state = 0
    topl.state=lightstateoff
    sp1.state=lightstateoff
    sp2.state=lightstateoff
    ebl.state=lightstateoff
    spinl.state=lightstateoff
    leftsidel.state=lightstateoff
    lanekl.state=lightstateoff
    dbl.state=lightstateoff
end sub

sub addscore(points)
    if tilt=false then
    bell=0
'    bb2.isdropped=true
    if points=100 or points=10 then
    matchnumb=matchnumb+1
    if matchnumb=10 then matchnumb=0
    end if
    if points = 10 or points = 100 or points = 1000 or points=10000 then scn=1
    if points = 100 then
    reel(currpl).addvalue(100)
    bell=100
    end if
    if points = 10 then
    reel(currpl).addvalue(10)
    bell=10
    end if
    if points = 1000 then
    reel(currpl).addvalue(1000)
    bell=1000
    end if
    if points = 10000 then
    reel(currpl).addvalue(10000)
    bell=10000
    end if
    if points = 500 then
    reel(currpl).addvalue(500)
    scn=5
    bell=100
    end if
    if points = 20000 then
    reel(currpl).addvalue(20000)
    scn=2
    bell=10000
    end if
    if points=2000 then
    reel(currpl).addvalue(2000)
    scn=2
    bell=1000
    end if
    if points = 30000 then
    reel(currpl).addvalue(30000)
    scn=3
    bell=10000
    end if
    if points=3000 then
    reel(currpl).addvalue(3000)
    scn=3
    bell=1000
    end if
    if points = 40000 then
    reel(currpl).addvalue(40000)
    scn=4
    bell=10000
    end if
    if points=4000 then
    reel(currpl).addvalue(4000)
    scn=4
    bell=1000
    end if
    if points=50000 then
    reel(currpl).addvalue(50000)
    scn=5
    bell=10000
    end if
    if points=5000 then
    reel(currpl).addvalue(5000)
    scn=5
    bell=1000
    end if
    scn1=0
    scntimer.enabled=true
    score(currpl)=score(currpl)+points
    truesc(currpl)=truesc(currpl)+points
    if score(currpl)=>1000000 then
    score(currpl)=score(currpl)-1000000
    rep(currpl)=0
    end if
    if score(currpl)=>replay1 and rep(currpl)=0 then
    credit=credit+1
  DOF 128, DOFOn
    playsound SoundFXDOF("knocker",126,DOFPulse,DOFKnocker)
  DOF 127, DOFPulse
    if credit>25 then credit=25
  credtxt.text=credit
    rep(currpl)=1
    playsound "click"
    end if
    if score(currpl)=>replay2 and rep(currpl)=1 then
    credit=credit+1
  DOF 128, DOFOn
  playsound SoundFXDOF("knocker",126,DOFPulse,DOFKnocker)
  DOF 127, DOFPulse
    if credit>25 then credit=25
  credtxt.text=credit
    rep(currpl)=2
    playsound "click"
    end if
    if score(currpl)=>replay3 and rep(currpl)=2 then
    credit=credit+1
  DOF 128, DOFOn
    playsound SoundFXDOF("knocker",126,DOFPulse,DOFKnocker)
  DOF 127, DOFPulse
    if credit>25 then credit=25
  credtxt.text=credit
    rep(currpl)=3
    playsound "click"
    end if
    end if
end sub

sub scntimer_timer
    scn1=scn1 + 1
    if bell=10 then playsound SoundFXDOF("bell10",141,DOFPulse,DOFChimes)
    if bell=100 then playsound SoundFXDOF("bell100",142,DOFPulse,DOFChimes)
    if bell=1000 then playsound SoundFXDOF("bell1000",143,DOFPulse,DOFChimes)
    if bell=10000 then playsound SoundFXDOF("bell1000low",144,DOFPulse,DOFChimes)
    if scn1=scn then
    if br=1 then bld
    scntimer.enabled=false
    end if
end sub

sub matchnum
    select case(matchnumb)
    case 0:
    m0.text="00"
  if b2son Then
    Controller.b2ssetdata 34,1
  end if
    case 1:
    m1.text="10"
  if b2son Then
    Controller.b2ssetdata 34,2
  end if
    case 2:
    m2.text="20"
  if b2son Then
    Controller.b2ssetdata 34,3
  end if
    case 3:
    m3.text="30"
  if b2son Then
    Controller.b2ssetdata 34,4
  end if
    case 4:
    m4.text="40"
  if b2son Then
    Controller.b2ssetdata 34,5
  end if
    case 5:
    m5.text="50"
  if b2son Then
    Controller.b2ssetdata 34,6
  end if
    case 6:
    m6.text="60"
  if b2son Then
    Controller.b2ssetdata 34,7
  end if
    case 7:
    m7.text="70"
  if b2son Then
    Controller.b2ssetdata 34,8
  end if
    case 8:
    m8.text="80"
  if b2son Then
    Controller.b2ssetdata 34,9
  end if
    case 9:
    m9.text="90"
  if b2son Then
    Controller.b2ssetdata 34,10
  end if
    end select
    for i=1 to playno
    if (matchnumb*10)=(score(i) mod 100) then
    credit=credit+1
  DOF 128, DOFOn
    playsound SoundFXDOF("knocker",126,DOFPulse,DOFKnocker)
  DOF 127, DOFPulse
    if credit>25 then credit=25
  credtxt.text=credit
    playsound "click"
    end if
    next
end sub

sub credtimer_timer
'    cred=cred+1
'    if cred=1 then credittxt.text="Ported To VP By Leon Spalding"
'    if cred=2 then credittxt.text="At The Request Of Bogy"
'    if cred=3 then credittxt.text="'F1' 3 or 5 Ball Game"
'    if cred=4 then credittxt.text="'5' Add Coin : 'S' Start game"
'    if cred=5 then credittxt.text="Thanx To Black (load&save)"
'    if cred=6 then credittxt.text="Thanx To Randy Davis (guess)"
'    if cred=7 then credittxt.text="Thanx To All At Shivas VP Forum"
'    if cred=8 then credittxt.text="'F1' 3 or 5 Ball Game"
'    if cred=9 then credittxt.text="'5' Add Coin : 'S' Start game"
'    if cred=10 then credittxt.text="An 'IR Pinball' Preserved Classic"
'    if cred=11 then credittxt.text="'IR Pinball' Are...."
'    if cred=12 then credittxt.text="Steveir, Duglis, Danz,"
'    if cred=13 then credittxt.text="Robair, Jay & Leon."
'    if cred=14 then credittxt.text="www.hippie.net/shivasite/irpinball"
'    if cred=14 then cred=0
end sub


'**************** Bonus 10 lights ***************
sub bonadv
    bl(bv).state=lightstateoff
    bv=bv+1
    if bv>10 then bv=10
    bl(bv).state=lightstateon
    playsound "clerker"
end sub
'************************************************


'************** Slingshots ***********************

Sub RightSling_Slingshot
  light
  playsound SoundFXDOF("rightsling",104,DOFPulse,DOFContactors)
  DOF 106, DOFPulse
' if tilt=false then PlaySound "Bumper"
  addscore 10
    Rubber3.Visible = 0
    Rubber3a.Visible = 1
  slingR.objroty = -15
  RStep = 2
    RightSling.TimerEnabled = 1
End Sub

Sub RightSling_Timer
    Select Case RStep
        Case 3:Rubber3a.Visible = 0:Rubber3b.Visible = 1:slingR.objroty = -7
        Case 4: slingR.objroty = 0:Rubber3b.Visible = 0:Rubber3.Visible = 1:RightSling.TimerEnabled = 0
    End Select
    RStep = RStep + 1
End Sub

Sub LeftSling_Slingshot
  light
  playsound SoundFXDOF("leftslingshot",103,DOFPulse,DOFContactors)
  DOF 105, DOFPulse
' if tilt=false then PlaySound "Bumper"
  addscore 10
    Rubber2.Visible = 0
    Rubber2a.Visible = 1
  slingL.objroty = 15
    LStep = 2
    LeftSling.TimerEnabled = 1
End Sub

Sub LeftSling_Timer()
    Select Case LStep
        Case 3:Rubber2a.Visible = 0:Rubber2b.Visible = 1:slingL.objroty = 7
        Case 4:slingL.objroty = 0:Rubber2b.Visible = 0:Rubber2.Visible = 1:LeftSling.TimerEnabled = 0
    End Select
    LStep = LStep + 1
End Sub

Sub topSling_Slingshot
  light
  playsound SoundFXDOF("leftslingshot",103,DOFPulse,DOFContactors)
  DOF 105, DOFPulse
  addscore 10
    LStepa = 2
    topSling.TimerEnabled = 1
End Sub

Sub topSling_Timer()
    Select Case LStepa
        Case 2:Rubber12.Visible = 0:Rubber12a.Visible = 1
        Case 3:Rubber12a.Visible = 0:Rubber12b.Visible = 1
        Case 4:Rubber12b.Visible = 0:Rubber12c.Visible = 1
        Case 5:Rubber12c.Visible = 0:Rubber12b.Visible = 1
        Case 6:Rubber12b.Visible = 0:Rubber12c.Visible = 1
        Case 7:Rubber12c.Visible = 0:Rubber12.Visible = 1:topSling.TimerEnabled = 0
    End Select
    LStepa = LStepa + 1
End Sub

sub topslingL_slingshot
  light
  playsound SoundFXDOF("leftslingshot",103,DOFPulse,DOFContactors)
  addscore 10
end sub

Sub topslingL2_Slingshot
  light
  playsound SoundFXDOF("leftslingshot",103,DOFPulse,DOFContactors)
  DOF 105, DOFPulse
  addscore 10
    LStepb = 2
    topslingL2.TimerEnabled = 1
End Sub

Sub topslingL2_Timer()
    Select Case LStepb
        Case 2:Rubber8.Visible = 0:Rubber8a.Visible = 1
        Case 3:Rubber8a.Visible = 0:Rubber8b.Visible = 1
        Case 4:Rubber8b.Visible = 0:Rubber8c.Visible = 1
        Case 5:Rubber8c.Visible = 0:Rubber8b.Visible = 1
        Case 6:Rubber8b.Visible = 0:Rubber8c.Visible = 1
        Case 7:Rubber8c.Visible = 0:Rubber8.Visible = 1:topslingL2.TimerEnabled = 0
    End Select
    LStepb = LStepb + 1
End Sub


sub topslingR_slingshot
  light
  playsound SoundFXDOF("rightslingshot",103,DOFPulse,DOFContactors)
  addscore 10
end sub


Sub sideSling_Slingshot
  light
  playsound SoundFXDOF("leftslingshot",103,DOFPulse,DOFContactors)
  DOF 105, DOFPulse
  addscore 10
    RStepa = 2
    sideSling.TimerEnabled = 1
End Sub

Sub sideSling_Timer()
    Select Case RStepa
        Case 2:Rubber22.Visible = 0:Rubber22a.Visible = 1
        Case 3:Rubber22a.Visible = 0:Rubber22b.Visible = 1
        Case 4:Rubber22b.Visible = 0:Rubber22c.Visible = 1
        Case 5:Rubber22c.Visible = 0:Rubber22b.Visible = 1
        Case 6:Rubber22b.Visible = 0:Rubber22c.Visible = 1
        Case 7:Rubber22c.Visible = 0:Rubber22.Visible = 1:sideSling.TimerEnabled = 0
    End Select
    RStepa = RStepa + 1
End Sub

Sub sideSlingR1_Slingshot
  light
  playsound SoundFXDOF("Rightslingshot",103,DOFPulse,DOFContactors)
  DOF 105, DOFPulse
  addscore 10
    RStepb = 2
    sideSlingR1.TimerEnabled = 1
End Sub

Sub sideSlingR1_Timer()
    Select Case RStepb
        Case 2:Rubber23.Visible = 0:Rubber23a.Visible = 1
        Case 3:Rubber23a.Visible = 0:Rubber23b.Visible = 1
        Case 4:Rubber23b.Visible = 0:Rubber23c.Visible = 1
        Case 5:Rubber23c.Visible = 0:Rubber23b.Visible = 1
        Case 6:Rubber23b.Visible = 0:Rubber23c.Visible = 1
        Case 7:Rubber23c.Visible = 0:Rubber23.Visible = 1:sideSlingR1.TimerEnabled = 0
    End Select
    RStepb = RStepb + 1
End Sub

Sub sideSlingR2_Slingshot
    RStepc = 2
    sideSlingR2.TimerEnabled = 1
End Sub

Sub sideSlingR2_Timer()
    Select Case RStepc
        Case 2:Rubber19.Visible = 0:Rubber19a.Visible = 1
        Case 3:Rubber19a.Visible = 0:Rubber19b.Visible = 1
        Case 4:Rubber19b.Visible = 0:Rubber19c.Visible = 1
        Case 5:Rubber19c.Visible = 0:Rubber19b.Visible = 1
        Case 6:Rubber19b.Visible = 0:Rubber19c.Visible = 1
        Case 7:Rubber19c.Visible = 0:Rubber19.Visible = 1:sideSlingR2.TimerEnabled = 0
    End Select
    RStepc = RStepc + 1
End Sub

Sub sideSlingL1_Slingshot
    LStepc = 2
    sideSlingL1.TimerEnabled = 1
End Sub

Sub sideSlingL1_Timer()
    Select Case LStepc
        Case 2:Rubber5.Visible = 0:Rubber5a.Visible = 1
        Case 3:Rubber5a.Visible = 0:Rubber5b.Visible = 1
        Case 4:Rubber5b.Visible = 0:Rubber5c.Visible = 1
        Case 5:Rubber5c.Visible = 0:Rubber5b.Visible = 1
        Case 6:Rubber5b.Visible = 0:Rubber5c.Visible = 1
        Case 7:Rubber5c.Visible = 0:Rubber5.Visible = 1:sideSlingL1.TimerEnabled = 0
    End Select
    LStepc = LStepc + 1
End Sub

'******************************************************************************

sub rbut_hit
  lbutt01.state = 1
    addscore 100
  butt03.z=-1.5
end sub
sub rbut_unhit
  'butt03.z=0.5
  cc=0
  rbut0.Enabled = 1
  lbutt01.state = 0
end sub

dim cc
cc=0

sub rbut0_Timer()
    Select Case cc
        Case 0:butt03.z=-1.3
        Case 1:butt03.z=-0.8
        Case 2:butt03.z=-0.3
        Case 3:butt03.z=0.2
        Case 4:butt03.z=0.4
        Case 5:butt03.z=0.5:rbut0.enabled = 0
    End Select
CC=cc+1
end sub

sub lbut_hit
  lbutt03.state = 1
    addscore 100
  butt01.z=-1.5
end sub
sub lbut_unhit
' butt01.z=0.5
  ccc=0
  rbut00.Enabled = 1
  lbutt03.state = 0
end sub

dim ccc
ccc=0

sub rbut00_Timer()
    Select Case ccc
        Case 0:butt01.z=-1.3
        Case 1:butt01.z=-0.8
        Case 2:butt01.z=-0.3
        Case 3:butt01.z=0.2
        Case 4:butt01.z=0.4
        Case 5:butt01.z=0.5:rbut00.enabled = 0
    End Select
CCC=ccc+1
end sub


sub leftin_hit
  DOF 119, DOFPulse
    addscore 5000
end sub

sub rightin_hit
  DOF 120, DOFPulse
    addscore 5000
end sub

sub rightout_hit
  DOF 121, DOFPulse
    if (sp2.state)=lightstateon then
    credit=credit+1
  DOF 128, DOFOn
    playsound SoundFXDOF("knocker",126,DOFPulse,DOFKnocker)
  DOF 127, DOFPulse
    if credit>25 then credit=25
  credtxt.text=credit
    playsound "click"
    else
    addscore 10000
    bonadv
    end if
end sub

sub leftout_hit
  DOF 118, DOFPulse
    if (sp1.state)=lightstateon then
    credit=credit+1
  DOF 128, DOFOn
    playsound SoundFXDOF("knocker",126,DOFPulse,DOFKnocker)
  DOF 127, DOFPulse
    if credit>25 then credit=25
  credtxt.text=credit
    playsound "click"
    else
    addscore 10000
    bonadv
    end if
end sub

sub spinner1_spin
  DOF 129, DOFPulse
  PlaySound "spinner", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0
    if (spinl.state)=lightstateon then
    addscore 1000
    else
    addscore 100
    end if
end sub

sub cbut_hit
  lbutt02.state = 1
    bonadv
  butt02.z=-1.5
end sub
sub cbut_unhit
  lbutt02.state = 0
' butt02.z=0.5
  cccc=0
  rbut000.Enabled = 1
end sub

dim cccc
cccc=0

sub rbut000_Timer()
    Select Case cccc
        Case 0:butt02.z=-1.3
        Case 1:butt02.z=-0.8
        Case 2:butt02.z=-0.3
        Case 3:butt02.z=0.2
        Case 4:butt02.z=0.4
        Case 5:butt02.z=0.5:rbut000.enabled = 0
    End Select
CCCC=cccc+1
end sub


sub laneK_hit
    if (lanekl.state)=lightstateon then dbl.state=lightstateon
    if sm>0 then
    addscore (10000*sm)
    else
    addscore 1000
    end if
    playsound "kickerkick"
    kickertimer.enabled=true
end sub

sub kickertimer_timer
    lanek.kick 0,18
    kickertimer.enabled=false
end sub

sub leftside_hit
    if (leftsidel.state)=lightstateon then dbl.state=lightstateon
    bonadv
end sub

sub bumper1_hit
  if tilt=false then playsound SoundFXDOF("jet2",109,DOFPulse,DOFContactors)
  DOF 110, DOFPulse
    if (Bumper1l.state)=lightstateon then
    addscore 1000
    else
    addscore 100
    end if
end sub

sub bumper2_hit
    if tilt=false then playsound SoundFXDOF("jet2",107,DOFPulse,DOFContactors)
  DOF 108, DOFPulse
    if (bumper2l.state)=lightstateon then
    addscore 1000
    else
    addscore 100
    end if
end sub

sub bumper3_hit
    if tilt=false then playsound SoundFXDOF("jet2",111,DOFPulse,DOFContactors)
  DOF 112, DOFPulse
    if (bumper3l.state)=lightstateon then
    addscore 1000
    else
    addscore 100
    end if
end sub

sub top_hit
  topt.transx = -10:Me.TimerEnabled = 1:PlaySound SoundFXDOF("target",125,DOFPulse,DOFTargets)
    if (topl.state)=lightstateon then bonadv
    addscore 1000
end sub
Sub top_Timer():topt.transx = 0:Me.TimerEnabled = 0:End Sub


sub upl_hit
  DOF 115, DOFPulse
    if (upll.state)=lightstateon then bonadv
    addscore 1000
end sub

sub upr_hit
  DOF 117, DOFPulse
    if (uprl.state)=lightstateon then bonadv
    addscore 1000
end sub

sub upc_hit
  DOF 116, DOFPulse
    addscore 1000
'    bumper1.state=lightstateon
    spinl.state=lightstateon
    if (upcl.state)=lightstateon then
'    upll.state=lightstateon
    upcl.state=lightstateoff
    al1.state=lightstateon
    al2.state=lightstateon
    sm=sm+1
    end if
    checkaward
end sub

sub midl_hit
  midlt.transx = -10:Me.TimerEnabled = 1:PlaySound SoundFXDOF("target",122,DOFPulse,DOFTargets)
    addscore 1000
'    bumper2.state=lightstateon
    if (midll.state)=lightstateon then
    uprl.state=lightstateon
    midll.state=lightstateoff
    zl1.state=lightstateon
    zl2.state=lightstateon
    sm=sm+1
    end if
    checkaward
end sub
Sub midl_Timer():midlt.transx = 0:Me.TimerEnabled = 0:End Sub

sub midr_hit
  midrt.transx = -10:Me.TimerEnabled = 1:PlaySound SoundFXDOF("target",124,DOFPulse,DOFTargets)
    addscore 1000
'    bumper3.state=lightstateon
    if (midrl.state)=lightstateon then
    topl.state=lightstateon
    midrl.state=lightstateoff
    tl1.state=lightstateon
    tl2.state=lightstateon
    sm=sm+1
    end if
    checkaward
end sub
Sub midr_Timer():midrt.transx = 0:Me.TimerEnabled = 0:End Sub

sub lowl_hit
  lowlt.transx = -10:Me.TimerEnabled = 1:PlaySound SoundFXDOF("target",122,DOFPulse,DOFTargets)
    addscore 1000
    if (lowll.state)=lightstateon then
    lowll.state=lightstateoff
    el1.state=lightstateon
    el2.state=lightstateon
    sm=sm+1
    end if
    checkaward
end sub
Sub lowl_Timer():lowlt.transx = 0:Me.TimerEnabled = 0:End Sub

sub lowr_hit
  lowrt.transx = -10:Me.TimerEnabled = 1:PlaySound SoundFXDOF("target",124,DOFPulse,DOFTargets)
    addscore 1000
    if (lowrl.state)=lightstateon then
    lowrl.state=lightstateoff
    cl1.state=lightstateon
    cl2.state=lightstateon
    sm=sm+1
    end if
    checkaward
end sub
Sub lowr_Timer():lowrt.transx = 0:Me.TimerEnabled = 0:End Sub

sub midc_hit
  midct.transx = -10:Me.TimerEnabled = 1:PlaySound SoundFXDOF("target",123,DOFPulse,DOFTargets)
    if (ebl.state)=lightstateon then
    shootagain.state=lightstateon
    ebl.state=lightstateoff
    end if
    if sm>0 then
    addscore (1000*sm)
    else
    addscore 500
    end if
    bonadv
end sub
Sub midc_Timer():midct.transx = 0:Me.TimerEnabled = 0:End Sub


sub checkaward
    if (al1.state)=lightstateon and (zl1.state)=lightstateon and (tl1.state)=lightstateon then
    ebl.state=lightstateon
    end if
    if (al1.state)=lightstateon and (zl1.state)=lightstateon and (tl1.state)=lightstateon and bv=10 then
    sp1.state=lightstateon
    sp2.state=lightstateon
    end if
    if (zl1.state)=lightstateon and (tl1.state)=lightstateon then
    leftsidel.state=lightstateon
    lanekl.state=lightstateon
    end if
end sub

Sub CheckTilt
  If Tilttimer.Enabled = True Then
  TiltSens = TiltSens + 1
  if TiltSens = 10 Then
  Tilt = True
  tilttxt.text="TILT"
  playsound "tilt"
  savehs
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
    bl(bv).state=lightstateoff
    for i=1 to 5
    azlet(i).state=lightstateoff
    az1(i).state=lightstateoff
    az2(i).state=lightstateoff
    next
    sp1.state=lightstateoff
    sp2.state=lightstateoff
    ebl.state=lightstateoff
    spinl.state=lightstateoff
    leftsidel.state=lightstateoff
    topl.state=lightstateoff
    lanekl.state=lightstateoff
    Bumper3l.state = 0
    Bumper2l.state = 0
    Bumper1l.state = 0
    uprl.state = 0
    upll.state = 0
    dbl.state=lightstateoff
    leftsling.slingshotstrength=0
    rightsling.slingshotstrength=0
end sub

sub savehs
  ' Based on Black's Highscore routines
  Dim FileObj
  Dim ScoreFile
  Set FileObj=CreateObject("Scripting.FileSystemObject")
  If Not FileObj.FolderExists(UserDirectory) then
    Exit Sub
  End if
  Set ScoreFile=FileObj.CreateTextFile(UserDirectory & "aztec.txt",True)
    ScoreFile.WriteLine credit
    ScoreFile.WriteLine score(1)
    ScoreFile.WriteLine matchnumb
        scorefile.writeline balls
        scorefile.writeline score(2)
        scorefile.writeline score(3)
        scorefile.writeline score(4)
        scorefile.writeline hisc1
        scorefile.writeline hisc2
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
    dim temp6
    dim temp7
    dim temp8
    dim temp9
    Set FileObj=CreateObject("Scripting.FileSystemObject")
  If Not FileObj.FolderExists(UserDirectory) then
    Exit Sub
  End if
  If Not FileObj.FileExists(UserDirectory & "aztec.txt") then
    Exit Sub
  End if
  Set ScoreFile=FileObj.GetFile(UserDirectory & "aztec.txt")
  Set TextStr=ScoreFile.OpenAsTextStream(1,0)
    If (TextStr.AtEndOfStream=True) then
      Exit Sub
    End if
    temp1=TextStr.ReadLine
    temp2=Textstr.ReadLine
    temp3=Textstr.ReadLine
        temp4=textstr.readline
        temp5=textstr.readline
        temp6=Textstr.ReadLine
        temp7=textstr.readline
        temp8=textstr.readline
        temp9=textstr.readline
    TextStr.Close
      credit = CDbl(temp1)
    If credit > 0 Then DOF 128, DOFOn
      score(1) = CDbl(temp2)
      matchnumb=cdbl(temp3)
      balls=cdbl(temp4)
      score(2)=cdbl(temp5)
      score(3)=cdbl(temp6)
      score(4)=cdbl(temp7)
      hisc1=cdbl(temp8)
      hisc2=cdbl(temp9)
      Set ScoreFile=Nothing
      Set FileObj=Nothing
end sub

sub ballrel_hit
    if ballrelenabled=1 then
    playsound "launchball"
    ballrelenabled=0
    end if
end sub

sub ballhome_hit
    ballrelenabled=1
end sub




'********************** Backglass script *****************

sub Scoreb2s_timer()

'********** Scores B2S ************
if B2SOn Then
  Controller.B2SSetScore 1, Score(1)
  Controller.B2SSetScore 2, Score(2)
  Controller.B2SSetScore 3, Score(3)
  Controller.B2SSetScore 4, Score(4)
  Controller.B2SSetCredits credit
  Controller.B2SSetCanPlay playno
  if ballinplay>0 and ballinplay<(balls+1) then
    Controller.B2SSetPlayerUp currpl
    Controller.B2SSetBallInPlay ballinplay
  end if

  '********** active PLayer's light in game ************
  if currpl=0 Then
    if ballinplay>0 and ballinplay<(balls+1) then
      Controller.B2SSetData 25, 1
    end if
    else
      Controller.B2SSetData 25, 0
  end If
  if currpl=1 Then
    if ballinplay>0 and ballinplay<(balls+1) then
      Controller.B2SSetData 26, 1
    end if
    else
      Controller.B2SSetData 26, 0
  end If
  if currpl=2 Then
    if ballinplay>0 and ballinplay<(balls+1) then
      Controller.B2SSetData 27, 1
    end if
    else
      Controller.B2SSetData 27, 0
  end If
  if currpl=3 Then
    if ballinplay>0 and ballinplay<(balls+1) then
      Controller.B2SSetData 28, 1
    end if
    else
      Controller.B2SSetData 28, 0
  end If

  '********** Tilt & ball in pley ************

  if Tilt = True then
    Controller.B2SSetTilt 33, 1
  Else
      Controller.B2SSetTilt 33, 0
  end If
End If

'********** Game Over ************
If ballinplay<1 or ballinplay>balls then
  If B2SOn Then Controller.B2SSetBallInPlay 0
  If B2SOn Then Controller.B2SSetData 35, 1: Controller.B2SSetPlayerUp 0

Else
  If B2SOn Then
    Controller.B2SSetData 35, 0
    Controller.B2SSetData 10, 0
    Controller.B2SSetData 11, 0
  End If

end If
end Sub
'**************************************************



'***************** Light routine ******************
sub light
  if Uprl.state = 1 then
    Bumper3l.state = 1
    Bumper2l.state = 1
    Bumper1l.state = 0
    uprl.state = 0
    upll.state = 1
  else
    Bumper3l.state = 0
    Bumper2l.state = 0
    Bumper1l.state = 1
    uprl.state = 1
    upll.state = 0
  end if
end Sub

'*********************** GI **************************
sub GIOFF
  shadows.visible = 0
  For each xx in GI:xx.state=0:next
end Sub

sub GION
  shadows.visible = 1
  For each xx in GI:xx.state=1:next
end Sub

'************ Flipper Shadows from Ninuzzu ************
Sub flippers_Timer()
  RollingSoundUpdate
  BallShadowUpdate
  LeftFlipperSh.RotZ = leftflipper.currentangle
  RightFlipperSh.RotZ = RightFlipper.currentangle
  if credit = 0 then
      creditlight.state = 0
    Else
      creditlight.state = 1
  end if
end Sub





'**************************************************************
'************************    Sounds    ************************
'**************************************************************

 'Sounds
dim rt, rb 'ramp top and ramp bottom

 dim speedx
 dim speedy
 dim finalspeed
  Sub RHSND_Hit(IDX)
  finalspeed=SQR(activeball.velx * activeball.velx + activeball.vely * activeball.vely)
  if finalspeed > 11 then PlaySound "rubber" else PlaySound "rubberFlipper":end if
   End Sub
  Sub Gates_Hit(IDX):RandomSoundGates():End Sub
  Sub LeftFlipper_Collide(parm)
  RandomSoundRubber()
   End Sub
   Sub RightFlipper_Collide(parm)
  RandomSoundRubber()
   End Sub

   Sub Ramp42_hit()
  RandomSoundMetal()
   End Sub
   Sub Ramp5_hit()
  RandomSoundMetal()
   End Sub

Sub Metal_Hit (idx)
  RandomSoundMetal()
End Sub

Sub plastic_Hit (idx)
  PlaySound "plastic", 0, Vol(activeball), Pan(activeball), 0, Pitch(activeball), 1, 0, AudioFade(ActiveBall)
End Sub

 Sub Gates_Hit(IDX):RandomSoundGates():End Sub


Sub Rubber_Hit(idx)
  dim finalspeed
    finalspeed=SQR(activeball.velx * activeball.velx + activeball.vely * activeball.vely)
  If finalspeed > 20 then
    PlaySound "rubber_hit_1", 0, Vol(activeball)*20*GlobalSoundLevel , pan(activeball), 0, Pitch(activeball), 0, 0, AudioFade(ActiveBall)
  End if
  If finalspeed >= 1 AND finalspeed <= 20 then
    RandomSoundRubber()
  End If
End Sub

  sub Triggermetal1_hit()
    PlaySound "metalhit2", 0, Vol(activeball)*10*GlobalSoundLevel , pan(activeball), 0, Pitch(activeball), 0, 0, AudioFade(ActiveBall)
  end sub
  sub Triggermetal2_hit()
    PlaySound "metalhit3", 0, Vol(activeball)*10*GlobalSoundLevel , pan(activeball), 0, Pitch(activeball), 0, 0, AudioFade(ActiveBall)
  end sub

Sub RandomSoundMetal()
  Select Case Int(Rnd*3)+1
    Case 1 : PlaySound "metalhit1", 0, Vol(activeball)*10*GlobalSoundLevel , pan(activeball), 0, Pitch(activeball), 0, 0, AudioFade(ActiveBall)
    Case 2 : PlaySound "metalhit2", 0, Vol(activeball)*10*GlobalSoundLevel , pan(activeball), 0, Pitch(activeball), 0, 0, AudioFade(ActiveBall)
    Case 3 : PlaySound "metalhit3", 0, Vol(activeball)*10*GlobalSoundLevel , pan(activeball), 0, Pitch(activeball), 0, 0, AudioFade(ActiveBall)
  End Select
End Sub

Sub RandomSoundRubber()
  Select Case Int(Rnd*3)+1
    Case 1 : PlaySound "rubber_hit_1", 0, Vol(activeball)*13*GlobalSoundLevel , pan(activeball), 0, Pitch(activeball), 0, 0, AudioFade(ActiveBall)
    Case 2 : PlaySound "rubber_hit_2", 0, Vol(activeball)*11*GlobalSoundLevel , pan(activeball), 0, Pitch(activeball), 0, 0, AudioFade(ActiveBall)
    Case 3 : PlaySound "rubber_hit_3", 0, Vol(activeball)*12*GlobalSoundLevel , pan(activeball), 0, Pitch(activeball), 0, 0, AudioFade(ActiveBall)
  End Select
End Sub

Sub RandomSoundGates()
  Select Case Int(Rnd*2)+1
    Case 1 : PlaySound "Gate", 0, Vol(activeball)*10*GlobalSoundLevel , pan(activeball), 0, Pitch(activeball), 0, 0, AudioFade(ActiveBall)
    Case 2 : PlaySound "Gate4", 0, Vol(activeball)*10*GlobalSoundLevel , pan(activeball), 0, Pitch(activeball), 0, 0, AudioFade(ActiveBall)
  End Select
End Sub

Sub RandomSoundFlipper()
  Select Case Int(Rnd*3)+1
    Case 1 :PlaySound "flip_hit_1", 0, GlobalSoundLevel * ballvel(activeball) / 50, -0.1, 0.25
    Case 2 :PlaySound "flip_hit_2", 0, GlobalSoundLevel * ballvel(activeball) / 50, -0.1, 0.25
    Case 3 :PlaySound "flip_hit_3", 0, GlobalSoundLevel * ballvel(activeball) / 50, -0.1, 0.25
  End Select
End Sub

'*******************************************************************************************



'*********** BALL SHADOW *********************************
Const tnob = 1 ' total number of balls for this table (at the same time)
ReDim BallShadow(tnob-1)
InitBallShadow

Sub InitBallShadow
  Dim i:For i = 0 to tnob-1
    ExecuteGlobal "Set BallShadow(" & i & ") = BallShadow" & (i+1) & " :"
  Next
End Sub

Sub BallShadowUpdate
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
    If BOT(b).X < Table1.Width/2 Then
      BallShadow(b).X = ((BOT(b).X))+10' - (50/6) + ((BOT(b).X - (Table1.Width/2))/7)) -2'+ 10
    Else
      BallShadow(b).X = ((BOT(b).X))-10' + (50/6) + ((BOT(b).X - (Table1.Width/2))/7)) +2'- 10
    End If
      ballShadow(b).Y = BOT(b).Y + 20
    If BOT(b).Z > 20 Then
      BallShadow(b).visible = 1
    Else
      BallShadow(b).visible = 0
    End If
  Next
End Sub

'************************ Reflection Mod ***************************

Sub RMod_Timer()
if shootagain.state = 1 then shootagainz.state = 1 else shootagainz.state = 0:end If
if sp1.state = 1 then sp1z.state = 1 else sp1z.state = 0:end If
if Lowll.state = 1 then Lowllz.state = 1 else Lowllz.state = 0:end If
if midll.state = 1 then midllz.state = 1 else midllz.state = 0:end If
if spinl.state = 1 then spinlz.state = 1 else spinlz.state = 0:end If
if b5k.state = 1 then b5kz.state = 1 else b5kz.state = 0:end If
if b10k.state = 1 then b10kz.state = 1 else b10kz.state = 0:end If
if b15k.state = 1 then b15kz.state = 1 else b15kz.state = 0:end If
if b20k.state = 1 then b20kz.state = 1 else b20kz.state = 0:end If
if b25k.state = 1 then b25kz.state = 1 else b25kz.state = 0:end If
if dbl.state = 1 then dblz.state = 1 else dblz.state = 0:end If
if b30k.state = 1 then b30kz.state = 1 else b30kz.state = 0:end If
if b35k.state = 1 then b35kz.state = 1 else b35kz.state = 0:end If
if b40k.state = 1 then b40kz.state = 1 else b40kz.state = 0:end If
if b45k.state = 1 then b45kz.state = 1 else b45kz.state = 0:end If
if b50k.state = 1 then b50kz.state = 1 else b50kz.state = 0:end If
if sp2.state = 1 then sp2z.state = 1 else sp2z.state = 0:end If
if lowrl.state = 1 then lowrlz.state = 1 else lowrlz.state = 0:end If
if midrl.state = 1 then midrlz.state = 1 else midrlz.state = 0:end If
if ebl.state = 1 then eblz.state = 1 else eblz.state = 0:end If
if al1.state = 1 then al1z.state = 1 else al1z.state = 0:end If
if zl1.state = 1 then zl1z.state = 1 else zl1z.state = 0:end If
if tl1.state = 1 then tl1z.state = 1 else tl1z.state = 0:end If
if el1.state = 1 then el1z.state = 1 else el1z.state = 0:end If
if cl1.state = 1 then cl1z.state = 1 else cl1z.state = 0:end If
if leftsidel.state = 1 then leftsidelz.state = 1 else leftsidelz.state = 0:end If
if topl.state = 1 then toplz.state = 1 else toplz.state = 0:end If
if upll.state = 1 then upllz.state = 1 else upllz.state = 0:end If
if upcl.state = 1 then upclz.state = 1 else upclz.state = 0:end If
if uprl.state = 1 then uprlz.state = 1 else uprlz.state = 0:end If
if lbutt03.state = 1 then lbutt03z.state = 1 else lbutt03z.state = 0:end If
if lbutt02.state = 1 then lbutt02z.state = 1 else lbutt02z.state = 0:end If
if lbutt01.state = 1 then lbutt01z.state = 1 else lbutt01z.state = 0:end If
if lanekl.state = 1 then laneklz.state = 1 else laneklz.state = 0:end If
if al2.state = 1 then al2z.state = 1 else al2z.state = 0:end If
if zl2.state = 1 then zl2z.state = 1 else zl2z.state = 0:end If
if tl2.state = 1 then tl2z.state = 1 else tl2z.state = 0:end If
if el2.state = 1 then el2z.state = 1 else el2z.state = 0:end If
if cl2.state = 1 then cl2z.state = 1 else cl2z.state = 0:end If
eND Sub


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
'*********** ROLLING SOUND *********************************

ReDim rolling(tnob)
InitRolling
Sub InitRolling:Dim i:For i=0 to (tnob-1):rolling(i) = False:Next:End Sub

Sub RollingSoundUpdate
    Dim BOT, b
    BOT = GetBalls
  ' stop the sound of deleted balls
  If UBound(BOT)<(tnob-1) Then
    For b = (UBound(BOT) + 1) to (tnob-1)
      rolling(b) = False
      StopSound("fx_ballrolling" & (b+1))
    Next
  End If
  ' exit the Sub if no balls on the table
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


' Thalamus : Exit in a clean and proper way
Sub table1_exit()
  Controller.Pause = False
  Controller.Stop
End Sub


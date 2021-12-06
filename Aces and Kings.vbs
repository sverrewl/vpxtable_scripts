Option Explicit
Randomize

  Const cGameName = "Aces_and_Kings"

'*****************************************************************************************************
' CREDITS
' Initial table created by fuzzel, jimmyfingers, jpsalas, toxie & unclewilly (in alphabetical order)
' Flipper primitives by zany
' Ball rolling sound script by jpsalas
' Ball shadow by ninuzzu
' Ball control by rothbauerw
' DOF by arngrim
' Positional sound helper functions by djrobx
' Plus a lot of input from the whole community (sorry if we forgot you :/)
'*****************************************************************************************************

On Error Resume Next
ExecuteGlobal GetTextFile("controller.vbs")
  If Err Then MsgBox "You need the Controller.vbs file"
  On Error Goto 0
ExecuteGlobal GetTextFile("core.vbs")

' Options
' Volume devided by - lower gets higher sound

Const VolDiv = 2000    ' Lower number, louder ballrolling/collition sound
Const VolCol = 10      ' Ball collition divider ( voldiv/volcol )

Dim Controller
Dim B2SScore
Dim B2SOn
dim ballinplay
dim ballrelenabled
dim rst
dim score(4)
dim truesc(4)
dim reel(5)
dim state
dim playno
dim credit
dim eg
dim currpl
dim plno(4)
dim play(4)
dim match(10)
dim tilt
dim tiltsens
dim rep(4)
dim plm(4)
dim matchnumb
dim obj
dim update
dim scn
dim scn1
dim bell
dim points
dim tempscore
dim replay1
dim replay2
dim replay3
dim replay4
dim hisc
dim bn(10)
dim spl(10)
dim up(4)
dim lp
dim ba
dim ca
dim batemp
dim catemp
dim wv
dim i


If Table1.ShowDT=false then
  reel1.Visible=false
  reel2.Visible=false
  reel3.Visible=false
  reel4.Visible=false
  reel5.Visible=false
  pno1.Visible=false
  pno2.Visible=false
  pno3.Visible=false
  pno4.Visible=false
  up1.Visible=false
  up2.Visible=false
  up3.Visible=false
  up4.Visible=false
  bip1.Visible=false
  bip2.Visible=false
  bip3.Visible=false
  bip4.Visible=false
  bip5.Visible=false
  m0.Visible=false
  m1.Visible=false
  m2.Visible=false
  m3.Visible=false
  m4.Visible=false
  m5.Visible=false
  m6.Visible=false
  m7.Visible=false
  m8.Visible=false
  m9.Visible=false
  credittxt.Visible=false
  gamov.Visible=false
Tilttxt.Visible=false
End If

sub table1_init
  InitAngles
  InitLamp
  ba=1
  ca=1
b1.state =1
c1.state =1
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
  replay1=4300
  replay2=5700
  replay3=7100
  replay4=8500
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
  Controller.B2SName = "Aces_and_Kings"
  Controller.Run()
  If Err Then MsgBox "Can't Load B2S.Server."
  end if
  If B2SOn Then
  if matchnumb=0 then
  Controller.B2SSetMatch 10
  else
  Controller.B2SSetMatch matchnumb
  end if
  Controller.B2SSetShootAgain 0
  Controller.B2SSetTilt 0
  Controller.B2SSetCredits Credit
  Controller.B2SSetGameOver 1
  Controller.B2SSetScorePlayer1 0
  Controller.B2SSetScorePlayer2 0
  Controller.B2SSetScorePlayer3 0
  Controller.B2SSetScorePlayer4 0
' Controller.B2SSetScorePlayer5 hisc
  End If
  for i=1 to 4
  If B2SOn Then
  Controller.B2SSetScorePlayer(i), 0
  End If
  next
  Light21.state=1
  Light28.state=1
  wv=1
  centerpost.visible=true
  centerpost1.visible=false
  stopper.isdropped=true
  postlight1.state=0
end sub


Sub Table1_KeyDown(ByVal keycode)

  If keycode = PlungerKey Then
  Plunger.PullBack
  PlaySound "plungerpull",0,1,AudioPan(Plunger),0.25,0,0,1,AudioFade(Plunger)
  End If

  if keycode = 6 then
  playsound "coin3"
  coindelay.enabled=true
  credit=credit+1
If credit >16 then credit =16
  If B2SOn Then
  Controller.B2SSetCredits Credit
  end if
  end if

  if keycode = 5 then
  playsound "coin3"
  coindelay.enabled=true
  credit=credit+1
If credit >16 then credit =16
  If B2SOn Then
  Controller.B2SSetCredits Credit
  End If
  end if

  if keycode = 2 and credit>0 and state=false and playno=0 then
  credit=credit-1
  credittxt.text=credit
  eg=0
  playno=1
  currpl=1
  pno.setvalue(playno)
  if playno=1 then pno1.state=1
  pno2.state=0
  pno3.state=0
  pno4.state=0
  play(currpl).state=1
  playsound "click"
  'playsound "motorshort1s",0, 0.15, 0, 0
  rst=0
  ballinplay=1
  resettimer.enabled=true
  If B2SOn Then
  Controller.B2SSetTilt 0
  Controller.B2SSetGameOver 0
  Controller.B2SSetMatch 0
  Controller.B2SSetCredits Credit
  Controller.B2SSetCanPlay 1
  Controller.B2SSetPlayerUp 1
  Controller.B2SSetBallInPlay BallInPlay
  Controller.B2SSetScorePlayer1 0
  Controller.B2SSetScorePlayer2 0
  Controller.B2SSetScorePlayer3 0
  Controller.B2SSetScorePlayer4 0
' Controller.B2SSetScorePlayer5 hisc
  End If
    end if

  if keycode = 2 and credit>0 and state=true and playno>0 and playno<4 and ballinplay<2 then
  credit=credit-1
  credittxt.text=credit
  playno=playno+1
  pno.setvalue(playno)
  if playno=2 then pno2.state=1: pno1.state=0
  if playno=3 then pno3.state=1: pno2.state=0
  if playno=4 then pno4.state=1: pno3.state=0
  playsound "click"
  If B2SOn Then
  Controller.B2SSetCredits Credit
  Controller.B2SSetCanPlay playno
  end if
    end if

  if state=true and tilt=false then

  If keycode = LeftFlipperKey Then
  LeftFlipper.RotateToEnd
  PlaySound SoundFX("flipperup",DOFFlippers), 0, .67, AudioPan(LeftFlipper), 0.05,0,0,1,AudioFade(LeftFlipper)
  PlayLoopSoundAtVol "buzzl", LeftFlipper, 1
  End If

  If keycode = RightFlipperKey Then
  RightFlipper.RotateToEnd
  PlaySound SoundFX("flipperup",DOFFlippers), 0, .67, AudioPan(RightFlipper), 0.05,0,0,1,AudioFade(RightFlipper)
  PlayLoopSoundAtVol "buzzr", RightFlipper, 1
  End If

  If keycode = LeftTiltKey Then
  Nudge 90, 6
  checktilt
  End If

  If keycode = RightTiltKey Then
  Nudge 270, 6
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
  PlaySound "plunger",0,1,AudioPan(Plunger),0.25,0,0,1,AudioFade(Plunger)
  End If

  If keycode = LeftFlipperKey Then
  LeftFlipper.RotateToStart
Stopsound "buzzl"
If state = True then
  PlaySound SoundFX("fx_flipperdown",DOFFlippers), 0, 1, AudioPan(LeftFlipper), 0.05,0,0,1,AudioFade(LeftFlipper)
End If
End If

  If keycode = RightFlipperKey Then
  RightFlipper.RotateToStart
Stopsound "buzzr"
If state = True then
  PlaySound SoundFX("fx_flipperdown",DOFFlippers), 0, 1, AudioPan(RightFlipper), 0.05,0,0,1,AudioFade(RightFlipper)
End If
End If
End Sub


sub coindelay_timer
  playsound "click"
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
' Controller.B2SSetScorePlayer5 hisc
  end if
  if rst=1 then
Playsound "motorleise"
  end if
  if rst=12 then
  newgame
  resettimer.enabled=false
  end if
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
If credit >16 then credit =16
  playsound "knocke"
  credittxt.text=credit
  If B2SOn Then
  Controller.B2SSetCredits Credit
  end if
  rep(currpl)=1
  playsound "click"
  end if
  if score(currpl)=>replay2 and rep(currpl)=1 then
  credit=credit+1
If credit >16 then credit =16
  playsound "knocke"
  credittxt.text=credit
  If B2SOn Then
  Controller.B2SSetCredits Credit
  end if
  rep(currpl)=2
  playsound "click"
  end if
  if score(currpl)=>replay3 and rep(currpl)=2 then
  credit=credit+1
If credit >16 then credit =16
  playsound "knocke"
  credittxt.text=credit
  If B2SOn Then
  Controller.B2SSetCredits Credit
  end if
  rep(currpl)=3
  playsound "click"
  end if
  if score(currpl)=>replay4 and rep(currpl)=3 then
  credit=credit+1
If credit >16 then credit =16
  playsound "knocke"
  credittxt.text=credit
  If B2SOn Then
  Controller.B2SSetCredits Credit
  end if
    rep(currpl)=4
    playsound "click"
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
' Playsound "solenoid"
  Stopper.IsDropped=True
  state=true
  ba=1
  ca=1
b1.state =1
c1.state =1
  rep(1)=0
  rep(2)=0
  rep(3)=0
  eg=0
  score(1)=0
  score(2)=0
  score(3)=0
  score(4)=0
  truesc(1)=0
  truesc(2)=0
  truesc(3)=0
  truesc(4)=0
' bumper1.force=12
' bumper2.force=12
' bumper3.force=12
  bip5.text=" "
  bip1.text="1"
  for i=0 to 9
  match(i).text=" "
  next
  tilttxt.text=" "
  gamov.text=" "
  tilt=false
  tiltsens=0
  ballinplay=1
  lightreset
  centerpost.visible=true
  centerpost1.visible=false
  stopper.isdropped=true
  postlight1.state=0
  nb.CreateSizedBall (26)
  nb.kick 135,6
Stopsound "motorleise"
Playsound "ballrel"
    bumper1.HasHitEvent=1
  bumper2.HasHitEvent=1
  bumper3.HasHitEvent=1
LeftSlingShot.Disabled =0
RightSlingShot.Disabled =0
  If B2SOn then
  Controller.B2SSetBallInPlay 1
  Controller.B2SSetPlayerUp 1
  end if
end sub


Sub Drain_Hit()
Playsound "drain"
Playsound "motorleise"
' ba=0
' ca=0
  Drain.DestroyBall
  Lightreset
   ' nextball
Timer1.enabled =1
Timer2.enabled =1
End Sub

sub nextball
  if tilt=true then
  tilt=false
  If B2SOn Then Controller.B2SSetTilt 0':Controller.B2SSetShootAgain 0
  tilttxt.text=" "
' Bumper1.force=12
' bumper2.force=12
' bumper3.force=12
  ba=1
  ca=1
  tiltseq.stopplay
  pno.setvalue(playno)
  LeftSlingShot.isdropped=false
  RightSlingShot.isdropped=false
  end if
  if (Light27.state)=1 then
  'playsound "kickerkick"
  ballreltimer.enabled=true
  else
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
  if ballinplay>3 then
  playsound "motorleer"
Stopsound "motorleise"
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
  'playsound "kickerkick"
  ballreltimer.enabled=true
  end if
  If B2SOn then Controller.B2SSetBallInPlay ballinplay
  If B2SOn then Controller.B2SSetPlayerUp currpl
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
    Controller.B2SSetPlayerUp currpl
  end If
  'playsound "kickerkick"
  ballreltimer.enabled=true
  end if
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
  If B2SOn Then
' Controller.B2SSetScorePlayer5 hisc
  End if
  end if
  next
  pno.setvalue(0)
  play(currpl-1).state=0
  playno=0
  gamov.text="GAME OVER"
  If B2SOn Then
  Controller.B2SSetGameOver 1
  Controller.B2SSetBallInPlay 0
  Controller.B2SSetPlayerUp 0
  end if
  savehs
  ballreltimer.enabled=false
  else
  nb.CreateSizedBall (26)
  nb.kick 135,6
Stopsound "motorleise"
Playsound "ballrel"
    bumper1.HasHitEvent=1
  bumper2.HasHitEvent=1
  bumper3.HasHitEvent=1
LeftSlingShot.Disabled =0
RightSlingShot.Disabled =0
If Stopper.IsDropped=False then Playsound "postdown"
  Stopper.IsDropped=True
  ballreltimer.enabled=false
' Light27.state=0
    end if
end sub

Sub wheelvalue
  if wv>1 then wv=0
  if wv=0 then Light21.state=1: Light22.state=0: Light28.state=1: Light29.state=0
  if wv=1 then Light21.state=0: Light22.state=1: Light28.state=0: Light29.state=1
  wv=wv+1
End Sub

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
If credit >16 then credit =16
  If B2SOn Then
  Controller.B2SSetCredits Credit
  end if
    playsound "knocke"
    credittxt.text= credit
    playsound "click"
    end if
    next
end sub

Sub CheckTilt
  If Tilttimer.Enabled = True Then
  TiltSens = TiltSens + 1
  if TiltSens = 2 Then
  Tilt = True
Playsound "tilt"
    bumper1.HasHitEvent=0
  bumper2.HasHitEvent=0
  bumper3.HasHitEvent=0
LeftSlingShot.Disabled =1
RightSlingShot.Disabled =1
LeftFlipper.RotateToStart
RightFlipper.RotateToStart
Stopsound "buzzl"
Stopsound "buzzr"
  tilttxt.text="TILT"
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

Sub lightreset
  BLight1.state=1
  BLight2.state=0
  BLight3.state=0
  TLight1.state=1
  TLight2.state=1
  TLight3.state=1
  TLight4.state=1
  TLight5.state=1
  TLight6.state=1
  TLight7.state=1
  TLight8.state=1
' b1.state=1
' b2.state=0
' b3.state=0
' b4.state=0
' b5.state=0
' b6.state=0
' b7.state=0
' b8.state=0
' b9.state=0
' b10.state=0
' c1.state=1
' c2.state=0
' c3.state=0
' c4.state=0
' c5.state=0
' c6.state=0
' c7.state=0
' c8.state=0
' c9.state=0
' c10.state=0
  Light23.state=0
  Light24.state=0
  Light25.state=0
  Light26.state=0
If Stopper.IsDropped=False then Playsound "postdown"
  centerpost.visible=true
  centerpost1.visible=false
  stopper.isdropped=true
  postlight1.state=0
End Sub

sub turnoff
' bumper1.force=0
' bumper2.force=0
' bumper3.force=0
  tiltseq.play seqalloff
    pno.setvalue(5)
    LeftSlingShot.isdropped=true
    RightSlingShot.isdropped=true
end sub



'Bumpers
Sub Bumper1_Hit
  If tilt=false then
  PlaySoundAtVol SoundFX("jet2",DOFContactors), Bumper1, 1
  Addscore 10
  End If
End Sub

Sub Bumper2_Hit
  If tilt=false then
  PlaySoundAtVol SoundFX("jet2",DOFContactors), Bumper2, 1
  If Blight2.state=1 then Addscore 10 else Addscore 1
  End If
End Sub

Sub Bumper3_Hit
  If tilt=false then
  PlaySoundAtVOl SoundFX("jet2",DOFContactors), Bumper3, 1
  If Blight3.state=1 then Addscore 10 else Addscore 1
  End If
End Sub


'Targets
Sub sw1_Hit
  If tilt=false then
  Addscore 10
  TLight1.state=0
  Blight2.state=1
  If TLight1.state=0 and TLight2.state=0 and TLight3.state=0 and TLight4.state=0 then Light23.state=1
  End If
End Sub

Sub sw2_Hit
  If tilt=false then
  Addscore 10
  TLight2.state=0
  Blight2.state=1
  If TLight1.state=0 and TLight2.state=0 and TLight3.state=0 and TLight4.state=0 then Light23.state=1
  End If
End Sub

Sub sw3_Hit
  If tilt=false then
  Addscore 10
  TLight3.state=0
  Blight2.state=1
  If TLight1.state=0 and TLight2.state=0 and TLight3.state=0 and TLight4.state=0 then Light23.state=1
  End If
End Sub

Sub sw4_Hit
  If tilt=false then
  Addscore 10
  TLight4.state=0
  Blight2.state=1
  If TLight1.state=0 and TLight2.state=0 and TLight3.state=0 and TLight4.state=0 then Light23.state=1
  End If
End Sub

Sub sw5_Hit
  If tilt=false then
  Addscore 10
  TLight5.state=0
  Blight3.state=1
  If TLight5.state=0 and TLight6.state=0 and TLight7.state=0 and TLight8.state=0 then Light24.state=1
  End If
End Sub

Sub sw6_Hit
  If tilt=false then
  Addscore 10
  TLight6.state=0
  Blight3.state=1
  If TLight5.state=0 and TLight6.state=0 and TLight7.state=0 and TLight8.state=0 then Light24.state=1
  End If
End Sub

Sub sw7_Hit
  If tilt=false then
  Addscore 10
  TLight7.state=0
  Blight3.state=1
  If TLight5.state=0 and TLight6.state=0 and TLight7.state=0 and TLight8.state=0 then Light24.state=1
  End If
End Sub

Sub sw8_Hit
  If tilt=false then
  Addscore 10
  TLight8.state=0
  Blight3.state=1
  If TLight5.state=0 and TLight6.state=0 and TLight7.state=0 and TLight8.state=0 then Light24.state=1
  End If
End Sub


'Triggers
Sub Trigger1_Hit()
  If tilt=false then
  If Light21.state=1 then Addscore 300 else Addscore 100
  Light25.state=1
  If TLight5.state=0 and TLight6.state=0 and TLight7.state=0 and TLight8.state=0 then Light24.state=1
  End If
End Sub

Sub Trigger2_Hit()
  If tilt=false then
  If Light22.state=1 then Addscore 300 else Addscore 100
  Light26.state=1
  If TLight5.state=0 and TLight6.state=0 and TLight7.state=0 and TLight8.state=0 then Light24.state=1
  End If
End Sub

Sub Trigger3_Hit()
  If tilt=false then
  Addscore 100
  End If
End Sub

Sub Trigger4_Hit()
  If tilt=false then
  Addscore 100
  End If
End Sub

Sub Trigger5_Hit()
PlaysoundAtVol "rollover", ActiveBall, 1
  If tilt=false and postlight1.state=0 then
  centerpost.visible=false
  centerpost1.visible=true
  stopper.isdropped=false
  postlight1.state=1
Playsound "postup"
  End If
End Sub

Sub Trigger6_Hit()
PlaysoundAtVol "rollover", ActiveBall, 1
If postlight1.state=1 then
  centerpost.visible=true
  centerpost1.visible=false
  stopper.isdropped=true
  postlight1.state=0
PlaysoundAt "postdown", centerpost
End If
End Sub


'Contactors
Sub LeftSlingShot1_Slingshot()
  If tilt=false then
  Addscore 1
  End If
End Sub

Sub LeftSlingShot2_Slingshot()
  If tilt=false then
  Addscore 10
  wheelvalue
  End If
End Sub

Sub LeftSlingShot3_Slingshot()
  If tilt=false then
  Addscore 10
  wheelvalue
  End If
End Sub

Sub RightSlingShot1_Slingshot()
  If tilt=false then
  Addscore 1
  End If
End Sub

Sub RightSlingShot2_Slingshot()
  If tilt=false then
  Addscore 10
  wheelvalue
  End If
End Sub

Sub RightSlingShot3_Slingshot()
  If tilt=false then
  Addscore 10
  wheelvalue
  End If
End Sub


'Kickers
sub Kicker1_Hit()
PlaysoundAtVol "ballrein", Kicker1, 1
PlayLoopSoundAtVol "motorleise",Kicker1,  1
If Tilt =True then kicktimer.enabled =1 : Exit Sub
If Light23.state=1 then
    Light27.state=1
    If B2SOn Then
        Controller.B2SSetShootAgain 1
    End If
    End If
    bonuscount.enabled=true
end sub

Sub Kicker2_Hit()
PlaysoundAtVol "ballrein", Kicker2, 1
PlayLoopSoundAtVol "motorleise", Kicker2, 1
If Tilt =True then kicktimer.enabled =1 : Exit Sub
  If Light24.state=1 then
    Light27.state=1
    If B2SOn Then
        Controller.B2SSetShootAgain 1
    End If
    End If
    bonuscount1.enabled=true
End Sub

sub kicktimer_timer
'Kicker1.kick 135,14
'Kicker2.kick 225,14
'Kicker1.Kick (int(rnd(1)*10)+90),15
'Kicker2.Kick (int(rnd(1)*5)+165),15
Kicker1.Kick (int(rnd(1)*3)+155),15
Kicker2.Kick (int(rnd(1)*3)+205),15
hebel1.rotz =15
hebel2.rotz =15
Timer4.enabled =1
    kicktimer.enabled=false
    playsound "ballout"
Stopsound "motorleise"
end sub

Sub Timer4_Timer
hebel1.rotz =0
hebel2.rotz =0
Me.enabled =0
End Sub


'**********Sling Shot Animations
' Rstep and Lstep  are the variables that increment the animation
'****************
Dim RStep, Lstep

Sub RightSlingShot_Slingshot
  If tilt=false then Addscore 1
  PlaySoundAtVol SoundFX("right_slingshot",DOFContactors), sling1, 1
  RSling.Visible = 0
  RSling1.Visible = 1
  sling1.rotx = 20
  RStep = 0
  RightSlingShot.TimerEnabled = 1
End Sub

Sub RightSlingShot_Timer
    Select Case RStep
        Case 3:RSLing1.Visible = 0:RSLing2.Visible = 1:sling1.rotx = 10
        Case 4:RSLing2.Visible = 0:RSLing.Visible = 1:sling1.rotx = 0:RightSlingShot.TimerEnabled = 0
    End Select
    RStep = RStep + 1
End Sub

Sub LeftSlingShot_Slingshot
  If tilt=false then Addscore 1
  PlaySoundAtVol SoundFX("left_slingshot",DOFContactors), sling2, 1
  LSling.Visible = 0
  LSling1.Visible = 1
  sling2.rotx = 20
  LStep = 0
  LeftSlingShot.TimerEnabled = 1
End Sub

Sub LeftSlingShot_Timer
    Select Case LStep
        Case 3:LSLing1.Visible = 0:LSLing2.Visible = 1:sling2.rotx = 10
        Case 4:LSLing2.Visible = 0:LSLing.Visible = 1:sling2.rotx = 0:LeftSlingShot.TimerEnabled = 0
    End Select
    LStep = LStep + 1
End Sub

' *******************************************************************************************************
' Positional Sound Playback Functions by DJRobX, Rothbauerw and Herweh
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

' The sound is played using the VOL, AUDIOPAN, AUDIOFADE and PITCH functions, so the volume and pitch of the sound
' will change according to the ball speed, and the AUDIOPAN & AUDIOFADE functions will change the stereo position
' according to the position of the ball on the table.


'**************************************
' Explanation of the collision routine
'**************************************

' The collision is built in VP.
' You only need to add a Sub OnBallBallCollision(ball1, ball2, velocity) and when two balls collide they
' will call this routine. What you add in the sub is up to you. As an example is a simple Playsound with volume and paning
' depending of the speed of the collision.


Sub Pins_Hit (idx)
  PlaySound "pinhit_low", 0, Vol(ActiveBall), AudioPan(ActiveBall), 0, Pitch(ActiveBall), 0, 0, AudioFade(ActiveBall)
End Sub

Sub Targets_Hit (idx)
  PlaySound "target", 0, Vol(ActiveBall), AudioPan(ActiveBall), 0, Pitch(ActiveBall), 0, 0, AudioFade(ActiveBall)
End Sub

Sub Metals_Thin_Hit (idx)
  PlaySound "metalhit_thin", 0, Vol(ActiveBall), AudioPan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
End Sub

Sub Metals_Medium_Hit (idx)
  PlaySound "metalhit_medium"', 0, Vol(ActiveBall), AudioPan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
End Sub

Sub Metals2_Hit (idx)
  PlaySound "metalhit2", 0, Vol(ActiveBall), AudioPan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
End Sub

Sub Gates_Hit (idx)
  PlaySound "gate4", 0, Vol(ActiveBall), AudioPan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
End Sub

Sub Spinner_Spin
  PlaySound "fx_spinner", 0, .25, AudioPan(Spinner), 0.25, 0, 0, 1, AudioFade(Spinner)
End Sub

Sub Rubbers_Hit(idx)
  dim finalspeed
    finalspeed=SQR(activeball.velx * activeball.velx + activeball.vely * activeball.vely)
  If finalspeed > 5 then
    PlaySound "fx_rubber2", 0, Vol(ActiveBall), AudioPan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
  End if
  If finalspeed >= 1 AND finalspeed <= 5 then
    RandomSoundRubber()
  End If
End Sub

Sub Posts_Hit(idx)
  dim finalspeed
    finalspeed=SQR(activeball.velx * activeball.velx + activeball.vely * activeball.vely)
  If finalspeed > 5 then
    PlaySound "fx_rubber2", 0, Vol(ActiveBall), AudioPan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
  End if
  If finalspeed >= 1 AND finalspeed <= 5 then
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

Sub Stopper_Hit
PlaysoundAt "pfosten", Stopper
End Sub


sub savehs
  ' Based on Black's Highscore routines
  Dim FileObj
  Dim ScoreFile
  Set FileObj=CreateObject("Scripting.FileSystemObject")
  If Not FileObj.FolderExists(UserDirectory) then
  Exit Sub
  End if
  Set ScoreFile=FileObj.CreateTextFile(UserDirectory & "Aces_and_Kings.txt",True)
  ScoreFile.WriteLine score(1)
  ScoreFile.WriteLine score(2)
  ScoreFile.WriteLine score(3)
  ScoreFile.WriteLine score(4)
  scorefile.writeline credit
  scorefile.writeline matchnumb
  scorefile.writeline hisc
  ScoreFile.Close
  Set ScoreFile=Nothing
  Set FileObj=Nothing
end sub

sub loadhs
    ' Based on Black's Highscore routines
  Dim FileObj
  Dim ScoreFile
  DIM TextStr
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
  If Not FileObj.FileExists(UserDirectory & "Aces_and_Kings.txt") then
  Exit Sub
  End if
  Set ScoreFile=FileObj.GetFile(UserDirectory & "Aces_and_Kings.txt")
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
  TextStr.Close
  score(1) = CDbl(temp1)
  score(2)= CDbl(temp2)
  score(3) = CDbl(temp3)
  score(4) = CDbl(temp4)
  credit= CDbl(temp5)
  matchnumb= CDbl(temp6)
  hisc=cdbl(temp7)
  Set ScoreFile=Nothing
  Set FileObj=Nothing
end sub


'-----------------------------------
' *Dorsola's lamp script (thanks!)*
'-----------------------------------
Dim LampPosition            ' In Degrees (0-180)
Dim SpinSpeed             ' SpinSpeed is in Deg/Sec
Dim LegNum
Dim ii

const pi = 3.1415926535897932
const SpeedMultiplier = 18        ' Affects speed transfer to object (deg/sec)
const Friction = 1.5          ' Friction coefficient (deg/sec/sec)
const MinSpeed = 100' 200         ' Object stops at this speed. (deg/sec) früher: 200, Orig.: 25
const StartAngle = 15         ' Start at 30 degrees rotation.
const ObjectRadius = 61         ' Distance from center of circle to center of post

dim ObjectX, ObjectY
ObjectX = t1.X
ObjectY = t1.Y

Dim AngleRef(32), LegPos1, LegPos2, PostX(32), PostY(32)
LegPos1 = Array( LampLeg1a, LampLeg2a, LampLeg3a, LampLeg4a, LampLeg5a, LampLeg6a, LampLeg7a, LampLeg8a, LampLeg9a, LampLeg10a, LampLeg11a, LampLeg12a, LampLeg13a, LampLeg14a, LampLeg15a, LampLeg16a )
LegPos2 = Array( LampLeg1b, LampLeg2b, LampLeg3b, LampLeg4b, LampLeg5b, LampLeg6b, LampLeg7b, LampLeg8b, LampLeg9b, LampLeg10b, LampLeg11b, LampLeg12b, LampLeg13b, LampLeg14b, LampLeg15b, LampLeg16b )

'--------------------------------
' Construction notes:
' With respect to the grid at standard size:
'   wall-centers are 1.5 units from center of circle.
'   Approach triggers are 2.75 units out.
'--------------------------------

Sub InitLamp()
  Dim obj
  for each obj in LegPos1 : obj.IsDropped = true : next
  for each obj in LegPos2 : obj.IsDropped = true : next

  LampPosition = StartAngle
  SpinSpeed = 0
  LegNum = 0
  DisplayLamp
End Sub

Sub InitAngles
  Dim ii
  for ii = 0 to 31
  AngleRef(ii) = ((11.25*ii) * pi/180)
  PostX(ii) = ObjectX + ObjectRadius*cos(AngleRef(ii))
  PostY(ii) = ObjectY - ObjectRadius*sin(AngleRef(ii))
  next
End Sub

Function GetAngle( x, y )
  ' Angle quadrants:
  '  Q2 (-x,-y)  |  Q1 (+x,-y)
  '--------------+--------------
  '  Q3 (-x,+y)  |  Q4 (+x,+y)

  if x = 0 And y < 0 then
    GetAngle = pi / 180             ' -90 degrees
  elseif x = 0 And y > 0 then
    GetAngle = -pi / 180            ' 90 degrees
  else
    GetAngle = atn( -y / x )
    if x < 0 then GetAngle = GetAngle + pi    ' Add 180 deg if in quadrant 2 or 3
  End if

  if GetAngle < 0 then GetAngle = GetAngle + 2*pi
  if GetAngle > 2*pi then GetAngle = GetAngle - 2*pi
End Function

Dim curPos, lastPos
lastPos = 7

Sub DisplayLamp()
  curPos = CInt(LampPosition / 11.25) Mod 16

  if curPos <> lastPos then
    LegPos1(lastPos).IsDropped = true
    LegPos2(lastPos).IsDropped = true
    LegPos1(curPos).IsDropped = false
    LegPos2(curPos).IsDropped = false

    if curPos = 0 then movelight  ' trip lamp switch


    lastPos = curPos
  end if

End Sub



Sub SpinTimer_Timer()
  if Abs(SpinSpeed) < MinSpeed then SpinSpeed = 0     ' Stop lamp below minimum speed.
  ComputePosition
  DisplayLamp
' DebugBox2.Text = SpinSpeed
' DebugBox2.Text = curPos
End Sub

' Definitions:
' SpinSpeed = Rotational speed in degrees per second.
' Diff/Position = degrees + SpinSpeed / nIntervals p.sec
' Friction applied: Speed - Speed*Frict/ nIntervals

Sub ComputePosition()
  ' Theory: Position is expressed in degrees.
  ' Add/Subtract to degrees with speed value.

  LampPosition = LampPosition + SpinSpeed / (1000 / SpinTimer.Interval)

  while LampPosition < 0.0
    LampPosition = LampPosition + 180.0
  wend

  while LampPosition > 180
    LampPosition = LampPosition - 180
  wend

  SpinSpeed = SpinSpeed - SpinSpeed * Friction / (1000/SpinTimer.Interval)
End Sub

'--------------------------
' Construction notes:
'   Angles are counter-clockwise, 0 degrees = right, 90 degrees = up
'--------------------------

Sub LampLeg1a_Hit() : ComputeLegHit 0 , ActiveBall : End Sub
Sub LampLeg2a_Hit() : ComputeLegHit 1 , ActiveBall : End Sub
Sub LampLeg3a_Hit() : ComputeLegHit 2 , ActiveBall : End Sub
Sub LampLeg4a_Hit() : ComputeLegHit 3 , ActiveBall : End Sub
Sub LampLeg5a_Hit() : ComputeLegHit 4 , ActiveBall : End Sub
Sub LampLeg6a_Hit() : ComputeLegHit 5 , ActiveBall : End Sub
Sub LampLeg7a_Hit() : ComputeLegHit 6 , ActiveBall : End Sub
Sub LampLeg8a_Hit() : ComputeLegHit 7 , ActiveBall : End Sub
Sub LampLeg9a_Hit() : ComputeLegHit 8 , ActiveBall : End Sub
Sub LampLeg10a_Hit() : ComputeLegHit 9 , ActiveBall : End Sub
Sub LampLeg11a_Hit() : ComputeLegHit 10 , ActiveBall : End Sub
Sub LampLeg12a_Hit() : ComputeLegHit 11 , ActiveBall : End Sub
Sub LampLeg13a_Hit() : ComputeLegHit 12 , ActiveBall : End Sub
Sub LampLeg14a_Hit() : ComputeLegHit 13 , ActiveBall : End Sub
Sub LampLeg15a_Hit() : ComputeLegHit 14 , ActiveBall : End Sub
Sub LampLeg16a_Hit() : ComputeLegHit 15 , ActiveBall : End Sub

Sub LampLeg1b_Hit() : ComputeLegHit 16 , ActiveBall : End Sub
Sub LampLeg2b_Hit() : ComputeLegHit 17 , ActiveBall : End Sub
Sub LampLeg3b_Hit() : ComputeLegHit 18, ActiveBall : End Sub
Sub LampLeg4b_Hit() : ComputeLegHit 19, ActiveBall : End Sub
Sub LampLeg5b_Hit() : ComputeLegHit 20, ActiveBall : End Sub
Sub LampLeg6b_Hit() : ComputeLegHit 21, ActiveBall : End Sub
Sub LampLeg7b_Hit() : ComputeLegHit 22, ActiveBall : End Sub
Sub LampLeg8b_Hit() : ComputeLegHit 23, ActiveBall : End Sub
Sub LampLeg9b_Hit() : ComputeLegHit 24 , ActiveBall : End Sub
Sub LampLeg10b_Hit() : ComputeLegHit 25 , ActiveBall : End Sub
Sub LampLeg11b_Hit() : ComputeLegHit 26, ActiveBall : End Sub
Sub LampLeg12b_Hit() : ComputeLegHit 27, ActiveBall : End Sub
Sub LampLeg13b_Hit() : ComputeLegHit 28, ActiveBall : End Sub
Sub LampLeg14b_Hit() : ComputeLegHit 29, ActiveBall : End Sub
Sub LampLeg15b_Hit() : ComputeLegHit 30, ActiveBall : End Sub
Sub LampLeg16b_Hit() : ComputeLegHit 31, ActiveBall : End Sub


Sub ComputeLegHit( LegNum, BallObj )

  ' Method: Determine coordinates of ball relative to leg and
  ' use those to determine the angle the ball hit the leg, then
  ' compute force applied to whole body.
  ' Construction notes: Items are all arranged CCW.

  ' Step 1: Determine angle of post to center, angle of hit point to
  ' center of post, and velocity/angle of ball travel.

  Dim postAngle, hitAngle, velAngle, velocity, ballAngle
  Dim dX, dY
  Dim tX, tY, lX, lY

  postAngle = AngleRef(LegNum)

  dX = BallObj.X - PostX(LegNum)
  dY = BallObj.Y - PostY(LegNum)
  hitAngle = GetAngle( dX, dY )

  dX = BallObj.X - ObjectX
  dY = BallObj.Y - ObjectY
  ballAngle = GetAngle( dX, dY )

  dX = BallObj.VelX
  dY = BallObj.VelY
  velocity = Sqr( dX*dX + dY*dY )
  velAngle = GetAngle( dX, dY )

  ' Step 2: Compute amount of force transferred to leg at hit point.
  ' This is determined by how much of the ball's velocity faces toward
  ' the center of the post.

  Dim speed, forceRatio, dAngle
  dAngle = velAngle - hitAngle
  forceRatio = cos(dAngle)    ' How close to a direct-line course is the ball on with the post?

  speed = SpeedMultiplier * 10 * forceRatio * velocity

  ' Step 3: Now, take the hit angle and compare it to the post's angle
  ' relative to the center of the object, and use that to determine the
  ' amount of force transferred to the object as a whole.

  dAngle = hitAngle - postAngle
  forceRatio = -sin(dAngle)
  speed = speed * forceRatio

' DebugBox3.Text = _
'   "vel= " & velocity & vbNewLine &_
'   "dirAng=" & velAngle & vbNewLine &_
'   "hitAng=" & hitAngle * 180 / pi & vbNewLine &_
'   "postAng=" & postAngle * 180 / pi & vbNewLine &_
'   "ballAng=" & ballAngle * 180 / pi & vbNewLine &_
'   "(Diff)=" & Abs(postAngle - ballAngle) * 180 / pi & vbNewLine &_
'   "frat=" & forceRatio & vbNewLine &_
'   "spd=" & speed

  SpinSpeed = SpinSpeed + speed
End Sub


sub movelight
If Tilt = True then Exit Sub
  addscore 10
  if Light28.state=1 then bonadv
  if Light29.state=1 then bonadv1
end sub

sub bonadv
    if tilt=false then
    ba=ba+1
    if ba>10 then ba=10
    batemp=ba
    if batemp>9 and batemp<11 then
    b10.state=1
    batemp=batemp-10
    end if
    select case (batemp)
    case 0:
    b9.state=0
    case 1:
    b1.state=1
    case 2:
    b1.state=0
    b2.state=1
    case 3:
    b2.state=0
    b3.state=1
    case 4:
    b3.state=0
    b4.state=1
    case 5:
    b4.state=0
    b5.state=1
    case 6:
    b5.state=0
    b6.state=1
    case 7:
    b6.state=0
    b7.state=1
    case 8:
    b7.state=0
    b8.state=1
    case 9:
    b8.state=0
    b9.state=1
    end select
    playsound "relay2"
    end if
end sub

sub bonuscount_timer
  playsound "motorshorter"
playsound "relay2"
  batemp=ba
  if batemp<10 then b10.state=0
  select case (batemp)
  case 0:
  b1.state=0
  case 1:
  b2.state=0
  b1.state=1
  case 2:
  b3.state=0
  b2.state=1
  case 3:
  b4.state=0
  b3.state=1
  case 4:
  b5.state=0
  b4.state=1
  case 5:
  b6.state=0
  b5.state=1
  case 6:
  b7.state=0
  b6.state=1
  case 7:
  b8.state=0
  b7.state=1
  case 8:
  b9.state=0
  b8.state=1
  case 9:
  b9.state=1
  end select
  if (light25.state)=1 then addscore 100 else addscore 10
  ba=ba-1
  if ba<1 then
ba =1
  bonuscount.enabled=false
  kicktimer.enabled=true
  b1.state=1
    end if
end sub


sub bonadv1
    if tilt=false then
    ca=ca+1
    if ca>10 then ca=10
    catemp=ca
    if catemp>9 and catemp<11 then
    c10.state=1
    catemp=catemp-10
    end if
    select case (catemp)
    case 0:
    c9.state=0
    case 1:
    c1.state=1
    case 2:
    c1.state=0
    c2.state=1
    case 3:
    c2.state=0
    c3.state=1
    case 4:
    c3.state=0
    c4.state=1
    case 5:
    c4.state=0
    c5.state=1
    case 6:
    c5.state=0
    c6.state=1
    case 7:
    c6.state=0
    c7.state=1
    case 8:
    c7.state=0
    c8.state=1
    case 9:
    c8.state=0
    c9.state=1
    end select
    playsound "relay2"
    end if
end sub


sub bonuscount1_timer
  playsound "motorshorter"
playsound "relay2"
  catemp=ca
  if catemp<10 then c10.state=0
  select case (catemp)
  case 0:
  c1.state=0
  case 1:
  c2.state=0
  c1.state=1
  case 2:
  c3.state=0
  c2.state=1
  case 3:
  c4.state=0
  c3.state=1
  case 4:
  c5.state=0
  c4.state=1
  case 5:
  c6.state=0
  c5.state=1
  case 6:
  c7.state=0
  c6.state=1
  case 7:
  c8.state=0
  c7.state=1
  case 8:
  c9.state=0
  c8.state=1
  case 9:
  c9.state=1
  end select
  if (light26.state)=1 then addscore 100 else addscore 10
  ca=ca-1
  if ca<1 then
ca =1
  bonuscount1.enabled=false
  kicktimer.enabled=true
  c1.state=1
    end if
end sub

Sub Timer1_Timer
  playsound "relay2"
  batemp=ba
  if batemp<10 then b10.state=0
  select case (batemp)
  case 0:
  b1.state=0
  case 1:
  b2.state=0
  b1.state=1
  case 2:
  b3.state=0
  b2.state=1
  case 3:
  b4.state=0
  b3.state=1
  case 4:
  b5.state=0
  b4.state=1
  case 5:
  b6.state=0
  b5.state=1
  case 6:
  b7.state=0
  b6.state=1
  case 7:
  b8.state=0
  b7.state=1
  case 8:
  b9.state=0
  b8.state=1
  case 9:
  b9.state=1
  end select
  'if (light25.state)=1 then addscore 100 else addscore 10
  ba=ba-1
  if ba<1 then
ba =1
  b1.state=1
Me.enabled =0
Timer3.enabled =1
    end if
End Sub

Sub Timer2_Timer
playsound "relay2"
  catemp=ca
  if catemp<10 then c10.state=0
  select case (catemp)
  case 0:
  c1.state=0
  case 1:
  c2.state=0
  c1.state=1
  case 2:
  c3.state=0
  c2.state=1
  case 3:
  c4.state=0
  c3.state=1
  case 4:
  c5.state=0
  c4.state=1
  case 5:
  c6.state=0
  c5.state=1
  case 6:
  c7.state=0
  c6.state=1
  case 7:
  c8.state=0
  c7.state=1
  case 8:
  c9.state=0
  c8.state=1
  case 9:
  c9.state=1
  end select
  'if (light26.state)=1 then addscore 100 else addscore 10
  ca=ca-1
  if ca<1 then
ca =1
  c1.state=1
Me.enabled =0
Timer3.enabled =1
    end if
End Sub

Sub Timer3_Timer
If b1.state=1 and c1.state=1 then
Me.enabled =0
nextball
End If
End Sub

Sub Gate_Hit
PlaysoundAtVol "gate", ActiveBall, 1
Light27.state=0
If B2SOn then Controller.B2SSetShootAgain 0
End Sub

licht.enabled =0

Dim zeit
Dim lampe1
Dim lampe2
Dim lampe3
Dim lampe4
Dim min
Dim max

Sub licht_Timer
min =0
max =4
lampe1=(Int((max-min+1)*Rnd+min))
min =0
max =4
lampe2=(Int((max-min+1)*Rnd+min))
min =0
max =4
lampe3=(Int((max-min+1)*Rnd+min))
min =0
max =4
lampe4=(Int((max-min+1)*Rnd+min))
If B2SOn then
If lampe1 >0 then Controller.B2SSetData 91,1
If lampe1 =0 then Controller.B2SSetData 91,0
If lampe2 >0 then Controller.B2SSetData 92,1
If lampe2 =0 then Controller.B2SSetData 92,0
If lampe3 >0 then Controller.B2SSetData 93,1
If lampe3 =0 then Controller.B2SSetData 93,0
If lampe4 >0 then Controller.B2SSetData 94,1
If lampe4 =0 then Controller.B2SSetData 94,0
End If
min =250
max =500
zeit=(Int((max-min+1)*Rnd+min))
licht.Interval =zeit
End Sub

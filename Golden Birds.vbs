'**************************************************************
'          GOLDEN BIRDS
'
'
' vp10 assembled and scripted by BorgDog, 2015, based on Golden Arrow (Gottlieb 1977)
' ball trough, captured pig routines, 3d model manipulations and graphics updates by sliderpoint
'
'**************************************************************

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
Const VolFlip   = 1    ' Flipper volume.


Dim DesktopMode: DesktopMode = Table1.ShowDT
DIM jump1,jump2,jump3, jump4, jump5
dim time1, time2, time3, time4, time5
dim count1, count2, count3, count4, count5, birdcount
dim ballinplay
dim ballrenabled
dim rst
dim eg
dim credit
dim score(2)
dim player
dim players
dim match(9)
dim rep
dim tilt
dim tiltsens
dim state
dim scn
dim scn1
dim bell
dim points, add10, add100, add1000
dim matchnumb
dim replay1
dim replay2
dim replay3
dim hisc
dim arrow(10)
dim numb(10)
dim numbstate(2,10)
dim apos(2)
dim ac(2)
dim sa(2)
dim spstate(2)
Dim RStep, Lstep
Dim objekt
'/////////////////Mike Add
'////////////////Ball Trough Dims
Dim f,g,h,i,j
Dim Ball(6)
Dim InitTime
Dim TroughTime
Dim EjectTime
Dim MaxBalls
Dim TroughCount
Dim TroughBall(7)
Dim TroughEject
'/////////////////BTrough Set
Dim BallB(6)
Dim gB
Dim InitTimeB
Dim TroughTimeB
Dim EjectTimeB
Dim TroughCountB
Dim TroughBallB(7)
Dim TroughEjectB
Dim Flip
'////////////////

 Dim Controller
 LoadController

 Sub LoadController()
  If DesktopMode = False Then
    Set Controller = CreateObject("B2S.Server")
    Controller.B2SName = "Golden Birds"
    Controller.Run()
  End if
 End Sub

sub table1_init
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
  set arrow(1)=a1a
  set arrow(2)=a2
  set arrow(3)=a3
  set arrow(4)=a4a
  set arrow(5)=a5a
  set arrow(6)=a6a
  set arrow(7)=a7
  set arrow(8)=a8
  set arrow(9)=a9a
  set arrow(10)=a10a
  set numb(1)=n1a
  set numb(2)=n2
  set numb(3)=n3
  set numb(4)=n4a
  set numb(5)=n5a
  set numb(6)=n6a
  set numb(7)=n7
  set numb(8)=n8
  set numb(9)=n9a
  set numb(10)=n10a
  player=1
    replay1=120000
    replay2=140000
    replay3=160000
    bumper1light.state=lightstateoff
    bumper2light.state=lightstateoff
    bumper3light.state=lightstateoff
    loadhs
    if hisc="" then hisc=50000
    hstxt.text=hisc
    if credit="" then credit=0
    credittxt.text=credit
  If DesktopMode = False Then
    setBackglass.enabled = true
    For each objekt in Backdropstuff: objekt.visible=false: next
  End If

  If FlipWall.transz=-80 then
    Flip = 8
    FlipWallTimer.Enabled = 1
  End if

    select case(matchnumb)
    case 0:
    m0.text="00"
    case 1:
    m1.text="10"
    case 2:
    m2.text="20"
    case 3:
    m3.text="30"
    case 4:
    m4.text="40"
    case 5:
    m5.text="50"
    case 6:
    m6.text="60"
    case 7:
    m7.text="70"
    case 8:
    m8.text="80"
    case 9:
    m9.text="90"
    end select
    scorereel1.setvalue(score(1))
    scorereel2.setvalue(score(2))
    if apos(1)<1 or apos(1)>10 then apos(1)=1
    if apos(2)<1 or apos(2)>10 then apos(2)=1

'///////////////////////////////////////////Mike add
'Set MaxBalls value from 1 to 6 to configure trough to handle
'required number of balls. No other script change is needed.

  MaxBalls=5
  InitTime=91
  EjectTime=0
  TroughEject=0
  TroughCount=0
  InitTimeB=91
  EjectTimeB=0
  TroughEjectB=0
  TroughCountB=0
'//////////////////////////

playsound "gameover"
end sub


'////////////////////////////////////Mike Add
Dim cBall, cBall2, cBall3
cBallInit

Sub cBallInit
  set cBall = captiveBall.createsizedball(24)
  cball.FrontDecal = "pigball"
  captiveball.kick 90,7

  set cBall2 = captiveBall2.createsizedball(15)
  cBall2.FrontDecal = "pigball"
  captiveball2.kick 90,2

  set cBall3 = captiveBall3.createsizedball(15)
  cBall3.FrontDecal = "pigball"
  captiveball3.kick 90,12
end Sub
'///////////////////////////


sub setBackglass_timer
  Controller.B2ssetCredits Credit
  Controller.B2ssetMatch 34, Matchnumb
    Controller.B2SSetGameOver 35,1
    Controller.B2SSetScorePlayer 5, hisc
    Controller.B2SSetScorePlayer 1, Score(1)
    Controller.B2SSetScorePlayer 2, Score(2)
  me.enabled=false
end sub


Sub Table1_KeyDown(ByVal keycode)

    if keycode=AddCreditKey then
    playsound "coin3"
    if state=false then
      If DesktopMode = False Then Controller.B2sStartAnimation "Creditss"
       credittxt.text=credit
    end if
    If DesktopMode = False Then Controller.B2SSetScore 4, hisc
    coindelay.enabled=true
    end if


    if keycode=StartGameKey and credit>0 then
    If DesktopMode = False then Controller.B2SSetScore 4, hisc
    if state=false then
    credit=credit-1
    playsound "click"
    credittxt.text=credit
    If DesktopMode = False then
      Controller.B2ssetCredits Credit
      Controller.B2sStartAnimation "Startup"
      Controller.B2ssetballinplay 32, Ballinplay
      Controller.B2ssetplayerup 30, 1
      Controller.B2ssetcanplay 31, 1
        Controller.B2ssetdata 1, 1
    End If
    tilt=false
    state=true
    rst=0
    ballinplay=1
    CanPlay.Text="One Player"
    playsound "initialize"
    resettimer.enabled=true
    players=1
    If FlipWall.transz=-80 then
      Flip = 8
      FlipWallTimer.Enabled = 1
    End if
'////////Mike add
    InitTimer.Enabled = True
'///////
    else if state=true and players < 2 and Ballinplay=1 then
    credit=credit-1
    players=players+1
    credittxt.text=credit
    If DesktopMode = False then
      Controller.B2ssetCredits Credit
      Controller.B2sStartAnimation "Startup"
      Controller.B2ssetcanplay 31, 2
    End If
    CanPlay.Text="Two Players"
    playsound "click"
    InitTimerB.Enabled = True
'   flipwall.transz=-80
    Flip = 1
    FlipWallTimer.Enabled = 1
     end if
    end if
  end if

  If keycode = PlungerKey Then
if birdjump.enabled=false then birdcount=0: birdjump.enabled=true  '*********test jumping
    Plunger.PullBack
    PlaySoundAtVol "slingshot", Plunger, 1
  End If

    if tilt=false and state=true then
    If keycode = LeftFlipperKey Then
    LeftFlipper.RotateToEnd
    PlaySoundAtVol "FlipperUp", LeftFlipper, VolFlip
    PlaySoundAtVol "Buzz", LeftFlipper, VolFlip
    End If

    If keycode = RightFlipperKey Then
    RightFlipper.RotateToEnd
        PlaySoundAtVol "FlipperUp", RightFlipper, VolFlip
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
    end If

  If keycode = MechanicalTilt Then
    mechchecktilt
  End If

    end if

End Sub

Sub Table1_KeyUp(ByVal keycode)

  If keycode = PlungerKey Then
    Plunger.Fire

  End If

  If keycode = LeftFlipperKey Then
    LeftFlipper.RotateToStart
    if tilt= false and state=true then
      PlaySoundAtVol "FlipperDown", LeftFlipper, Volflip
      StopSound "Buzz"
    end if
  End If

  If keycode = RightFlipperKey Then
    RightFlipper.RotateToStart
          if tilt= false and state=true then
      PlaySoundAtVol "FlipperDown", RightFlipper, VolFlip
      StopSound "Buzz1"
    end if
  End If

End Sub

Sub PairedlampTimer_timer
  a1b.state = a1a.state
  a4b.state = a4a.state
  a5b.state = a5a.state
  a6b.state = a6a.state
  a9b.state = a9a.state
  a10b.state = a10a.state
  n1b.state = n1a.state
  n4b.state = n4a.state
  n5b.state = n5a.state
  n6b.state = n6a.state
  n9b.state = n9a.state
  n10b.state = n10a.state
  rtl.state = ltl.state
end sub


sub coindelay_timer
  addcredit
  coindelay.enabled=false
end sub

sub resettimer_timer
    rst=rst+1
    scorereel1.resettozero
    scorereel2.resettozero
    If DesktopMode = False then
    Controller.B2SSetScore 1, score(1)
    Controller.B2SSetScore 2, score(2)
  End If
    if rst=18 then
    playsound "kickerkick"
    end if
    if rst=22 then
    newgame
    resettimer.enabled=false
    end if
end sub

'//////////////////Mike Add
Sub InitTimer_Timer()

'Sets up required ball array in trough. Game is effectively in
'bootup mode while InitTime>0. InitTime must be allowed to
'complete properly for ball-checking routines to function.

  If InitTime>0 then
    InitTime=InitTime-1
    If (InitTime/10 = Int(InitTime/10)) and InitTime>0 then
      TroughTime=TroughTime+8
      If TroughCount<MaxBalls then
        If TroughCount=0 then Drain.createsizedball(26).FrontDecal = "redbirdball"
        If TroughCount=1 then Drain.createsizedball(26).FrontDecal = "yellowbirdball"
        If TroughCount=2 then Drain.createsizedball(26).FrontDecal = "bluebirdball"
        If TroughCount=3 then Drain.createsizedball(26).FrontDecal = "whitebirdball"
        If TroughCount=4 then Drain.createsizedball(26).FrontDecal = "blackbirdball"
        TroughCount=TroughCount+1
      End If
    End If
  End If
End Sub

Sub InitTimerB_Timer()  '\\\\\\BSET FOR PLAYER 2

  If InitTimeB>0 then
    InitTimeB=InitTimeB-1
    If (InitTimeB/10 = Int(InitTimeB/10)) and InitTimeB>0 then
      TroughTimeB=TroughTimeB+8
      If TroughCountB<MaxBalls then
        If TroughCountB=0 then nb.createsizedball(26).FrontDecal = "redbirdball"
        If TroughCountB=1 then nb.createsizedball(26).FrontDecal = "yellowbirdball"
        If TroughCountB=2 then nb.createsizedball(26).FrontDecal = "bluebirdball"
        If TroughCountB=3 then nb.createsizedball(26).FrontDecal = "whitebirdball"
        If TroughCountB=4 then nb.createsizedball(26).FrontDecal = "blackbirdball"
        TroughCountB=TroughCountB+1
      End If
    End If
  End If
End Sub

Sub TroughTimer_Timer()

'Timer-based bubble-sort routine checks each six-ball array position
'in turn from right to left. TroughEject variable tells script when
'to eject new ball in play to the shooter lane. Note that drain hit
'subroutine must be part of this section.


  If TroughEject=1 then
    EjectTime=61
    TroughEject=0
  End If

  If EjectTime>0 then
    EjectTime=EjectTime-1
    If EjectTime=0 then
      If TroughBall(1)=1 then
        TroughTime=TroughTime+8
        Trough1.Kickz 50,6, .5, 10
        TroughCount=TroughCount-1
        TroughBall(1)=0
                if ballinplay=1 then PlaySound "redbirdselect"
                if ballinplay=2 then PlaySound "yellowbirdselect"
                if ballinplay=3 then PlaySound "bluebirdselect"
                if ballinplay=4 then PlaySound "whitebirdselect"
                if ballinplay=5 then PlaySound "blackbirdselect"
      End If
    End If
  End If

  If TroughTime>0 then
    TroughTime=TroughTime-1
    g=(Int(TroughTime/8))*8
    If (TroughTime-g)=6 and TroughBall(1)=0 then
      TroughBall(2)=0
      Trough2.Kick 45,10
    End If
    If (TroughTime-g)=5 and TroughBall(2)=0 then
      TroughBall(3)=0
      Trough3.Kick 45,10
    End If
    If (TroughTime-g)=4 and TroughBall(3)=0 then
      TroughBall(4)=0
      Trough4.Kick 45,10
    End If
    If (TroughTime-g)=3 and TroughBall(4)=0 then
      TroughBall(5)=0
      Trough5.Kick 45,10
    End If
    If (TroughTime-g)=2 and TroughBall(5)=0 then
      TroughBall(6)=0
      Trough6.Kick 45,10
    End If
    If (TroughTime-g)=1 and TroughBall(6)=0 then
      TroughBall(7)=0
      Drain.Kick 45,8
    End If
  End If
End Sub

Sub Trough1_Hit()
  TroughBall(1)=1
  If InitTime>0 then
    Set Ball(TroughCount)=ActiveBall
  End If
End Sub

Sub Trough2_Hit()
  TroughBall(2)=1
End Sub

Sub Trough3_Hit()
  TroughBall(3)=1
End Sub

Sub Trough4_Hit()
  TroughBall(4)=1
End Sub

Sub Trough5_Hit()
  TroughBall(5)=1
End Sub

Sub Trough6_Hit()
  TroughBall(6)=1
End Sub


'////////////////////////////BSET Trough FOR PLAYER 2
Sub TroughTimerB_Timer()

'Timer-based bubble-sort routine checks each six-ball array position
'in turn from right to left. TroughEject variable tells script when
'to eject new ball in play to the shooter lane. Note that drain hit
'subroutine must be part of this section.


  If TroughEjectB=1 then
    EjectTimeB=61
    TroughEjectB=0
  End If

  If EjectTimeB>0 then
    EjectTimeB=EjectTimeB-1
    If EjectTimeB=0 then
      If TroughBallB(1)=1 then
        TroughTimeB=TroughTimeB+8
        TroughB1.Kickz 50,6, .5, 10
        TroughCountB=TroughCountB-1
        TroughBallB(1)=0
                if ballinplay=1 then PlaySound "redbirdselect"
                if ballinplay=2 then PlaySound "yellowbirdselect"
                if ballinplay=3 then PlaySound "bluebirdselect"
                if ballinplay=4 then PlaySound "whitebirdselect"
                if ballinplay=5 then PlaySound "blackbirdselect"
      End If
    End If
  End If

  If TroughTimeB>0 then
    TroughTimeB=TroughTimeB-1
    gB=(Int(TroughTimeB/8))*8
    If (TroughTimeB-gB)=6 and TroughBallB(1)=0 then
      TroughBallB(2)=0
      TroughB2.Kick 45,10
    End If
    If (TroughTimeB-gB)=5 and TroughBallB(2)=0 then
      TroughBallB(3)=0
      TroughB3.Kick 45,10
    End If
    If (TroughTimeB-gB)=4 and TroughBallB(3)=0 then
      TroughBallB(4)=0
      TroughB4.Kick 45,10
    End If
    If (TroughTimeB-gB)=3 and TroughBallB(4)=0 then
      TroughBallB(5)=0
      TroughB5.Kick 45,10
    End If
    If (TroughTimeB-gB)=2 and TroughBallB(5)=0 then
      TroughBallB(6)=0
      TroughB6.Kick 45,10
    End If
    If (TroughTimeB-gB)=1 and TroughBallB(6)=0 then
      TroughBallB(7)=0
      nb.Kick 50,8
    End If
  End If
End Sub

Sub TroughB1_Hit()
  TroughBallB(1)=1
  If InitTimeB>0 then
    Set BallB(TroughCount)=ActiveBall
  End If
End Sub

Sub TroughB2_Hit()
  TroughBallB(2)=1
End Sub

Sub TroughB3_Hit()
  TroughBallB(3)=1
End Sub

Sub TroughB4_Hit()
  TroughBallB(4)=1
End Sub

Sub TroughB5_Hit()
  TroughBallB(5)=1
End Sub

Sub TroughB6_Hit()
  TroughBallB(6)=1
End Sub

'/////////////////////END PLAYER 2 TROUGH


sub newgame
    bumper1.force=7
    bumper2.force=7
    bumper3.force=7
  player=1
  shoot1.text="Player 1"
    score(1)=0
  score(2)=0
  If DesktopMode = False then
    Controller.B2SSetScore 1, score(1)
    Controller.B2SSetScore 2, score(2)
  End If
    eg=0
    rep=0
    sa(1)=0
  sa(2)=0
' sm(1)=1
' sm(2)=1
  spstate(1)=0
  spstate(2)=0
  Star1L.state=0: Star2L.state=0:Star3L.state=0
    sp.state=lightstateoff
    ltl.state=lightstateoff
    rtl.state=lightstateoff
    fivek.state=lightstateon
    bumper1light.state=lightstateon
    bumper2light.state=lightstateon
    bumper3light.state=lightstateon
  PlightRM.state=lightstateon
  TlightRM.State=lightstateon
  PlightLM.state=lightstateon
  TlightLM.state=lightstateon
  PlightLL.state=lightstateon
  PlightRL.state=lightstateon
  TlightLL.state=lightstateon
  TlightRL.state=lightstateon
  PlightMR.state=lightstateon
  TlightMR.state=lightstateon
  PlightML.state=lightstateon
  TlightML.state=lightstateon
  PlightLU.state=lightstateon
  TlightLU.state=lightstateon
  PlightRU.state=lightstateon
  TlightRU.state=lightstateon
  For Each Light in LaneLights
    Light.state=lightstateon
  Next
  for i=1 to 10
    arrow(i).state=lightstateoff
    numb(i).state=lightstateon
    numbstate(1,i)=1
    numbstate(2,i)=1
  next
    movearrow
    gamov.text=" "
    tilttxt.text=" "
    If DesktopMode = False then
    Controller.B2SSetGameOver 35,0
    Controller.B2SSetTilt 33,0
    Controller.B2SSetMatch 34,0
    Controller.B2ssetdata 11, 0
    Controller.B2ssetdata 7,0
    Controller.B2ssetdata 8,0
  End If
    bip5.text=" "
    bip1.text="1"
    for i=0 to 9
      match(i).text=" "
    next
'////////////Mike add
   TroughEject = 1
'////////////
end sub

sub newball
    bumper1light.state=lightstateon
    bumper3light.state=lightstateon
    bumper2light.state=lightstateon
    if sa(player)=1 then
'      for i=1 to 10
'   numb(i).state=lightstateon
'   numbstate(player,i)=1
'      next
'      sp.state=lightstateoff
'   spstate(player)=0
'      rtl.state=lightstateoff
'      ltl.state=lightstateoff
'      sa(player)=0
   else
    for i=1 to 10
      if numbstate(player,i)=1 then
          numb(i).state=lightstateon
        else
        numb(i).state=lightstateoff
      end if
      arrow(i).state=lightstateoff
    next
    arrow(apos(player)).state=lightstateon
    end if
end sub

Sub Drain_Hit()
  if Star3L.state = 2 then
    star3.timerenabled=0
    stopsound "StarWizard"
    star3L.state=1
  end if
  if shootagain.state=2 then
    drain.destroyball
    select case (ballinplay)
      case 1:
      PlaySound "redbirdselect"
      savedrain.createsizedball(26).FrontDecal = "redbirdball"
      case 2:
      PlaySound "yellowbirdselect"
      savedrain.createsizedball(26).FrontDecal = "yellowbirdball"
      case 3:
      PlaySound "bluebirdselect"
      savedrain.createsizedball(26).FrontDecal = "bluebirdball"
          case 4:
      PlaySound "whitebirdselect"
      savedrain.createsizedball(26).FrontDecal = "whitebirdball"
      case 5:
            PlaySound "blackbirdselect"
      savedrain.createsizedball(26).FrontDecal = "blackbirdball"
    end select
    savedrain.kick 180,1,0
    shootagain.state=0
  else

    playsoundAtVol "drainshorter", drain, 1
  '////////////Mike Add
    If players=1 or player=2 then
      TroughBall(7)=1
      TroughTime=TroughTime+48
      TroughCount=TroughCount+1
      If TroughCount=MaxBalls Then TroughEject=1
    Else
      TroughBallB(7)=1
      TroughTimeB=TroughTimeB+48
      TroughCountB=TroughCountB+1
      If TroughCountB=MaxBalls Then TroughEjectB=1
    End if
  '///////////
    Drain.DestroyBall
      playsoundAtVol "kickerkick", drain, 1
      if players=1 or player=2 then
        player=1
        If DesktopMode = False then Controller.B2ssetplayerup 30, 1
        shoot1.text="Player 1"
        shoot2.text=" "
        nextball
        else
        player=2
        If DesktopMode = False then Controller.B2ssetplayerup 30, 2
        shoot2.text="Player 2"
        shoot1.text=" "
        nextball
      end if
  end if
End Sub

sub nextball
    if tilt=true then
    bumper1.force=7
    bumper2.force=7
    bumper3.force=7
    tiltseq.stopplay
    tilt=false
    tilttxt.text=" "
    If DesktopMode = False then
      Controller.B2SSetTilt 33,0
      Controller.B2ssetdata 1, 1
    End If
    end if
  if player=1 then ballinplay=ballinplay+1
  If DesktopMode = False then Controller.B2ssetballinplay 32, Ballinplay
  if ballinplay>5 then
    playsound "motorleer"
    eg=1
    ballreltimer.enabled=true
    else
    if state=true and tilt=false then
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
End Sub

sub ballreltimer_timer
    if eg=1 then
      matchnum
    bip5.text=" "
    state=false
    gamov.text="GAME OVER"
    CanPlay.text=" "
    shoot1.text=" "
    shoot2.text=" "
    if score(1)>hisc then hisc=score(1)
    if score(2)>hisc then hisc=score(2)
    hstxt.text=hisc
  If DesktopMode = False then
      Controller.B2SSetGameOver 35,1
      Controller.B2ssetballinplay 32, 0
    Controller.B2sStartAnimation "EOGame"
    Controller.B2SSetScore 4, hisc
    Controller.B2ssetcanplay 31, 0
    Controller.B2ssetcanplay 30, 0
  End If
      If score(1)>replay1 or score(1)=hisc or score(2)>replay1 or score(2)=hisc then
    playsound "levelcleared"
    else
    playsound "levelfailed"
    end if
'//////////Mike Add

    InitTimer.Enabled = False
    TroughCount=0
    InitTime=91
    InitTimerB.Enabled = False
    TroughCountB=0
    InitTimeB=91

'///////////
      bumper1light.state=lightstateoff
      bumper2light.state=lightstateoff
      bumper3light.state=lightstateoff
    PlightRM.state=lightstateoff
    TlightRM.State=lightstateoff
    PlightLM.state=lightstateoff
    TlightLM.state=lightstateoff
    PlightLL.state=lightstateoff
    PlightRL.state=lightstateoff
    TlightLL.state=lightstateoff
    TlightRL.state=lightstateoff
    PlightMR.state=lightstateoff
    TlightMR.state=lightstateoff
    PlightML.state=lightstateoff
    TlightML.state=lightstateoff
    PlightLU.state=lightstateoff
    TlightLU.state=lightstateoff
    PlightRU.state=lightstateoff
    TlightRU.state=lightstateoff
    For each Light in LaneLights
      Light.state=lightstateoff
    next
    savehs
    cred=0
    ballreltimer.enabled=false
    else
'//////////Mike add
    If player = 1 then TroughEject = 1
    If player = 2 then TroughEjectB = 1
'//////////
        ballreltimer.enabled=false
    end if
end sub


    Sub FlipWallTimer_Timer()
        Select Case Flip
               Case 1:FlipWall.Transz=-11:Flip = 2
               Case 2:FlipWall.Transz=-22:Flip = 3
               Case 3:FlipWall.Transz=-33:Flip = 4
               Case 4:FlipWall.Transz=-44:Flip = 5
               Case 5:FlipWall.Transz=-56:Flip = 6
               Case 6:FlipWall.Transz=-68:Flip = 7
               Case 7:FlipWall.Transz=-80:Me.Enabled = 0
               Case 8:FlipWall.Transz=-68:Flip = 9
               Case 9:FlipWall.Transz=-56:Flip = 10
               Case 10:FlipWall.Transz=-44:Flip = 11
               Case 11:FlipWall.Transz=-33:Flip = 12
               Case 12:FlipWall.Transz=-22:Flip = 13
               Case 13:FlipWall.Transz=-11:Flip = 14
               Case 14:FlipWall.Transz=-0:Me.Enabled = 0
           End Select
     End Sub


sub matchnum
    select case(matchnumb)
      case 0:
      m0.text="00"
      case 1:
      m1.text="10"
      case 2:
      m2.text="20"
      case 3:
      m3.text="30"
      case 4:
      m4.text="40"
      case 5:
      m5.text="50"
      case 6:
      m6.text="60"
      case 7:
      m7.text="70"
      case 8:
      m8.text="80"
      case 9:
      m9.text="90"
    end select
  If DesktopMode = False then Controller.B2SSetMatch 34,Matchnumb*10
  For i=1 to players
    if (matchnumb*10)=(score(i) mod 100) then
    playsound "knocke"
    addcredit
  end if
  next
end sub

'**************
'   Score
'**************

'jp: the points to add must be: 10 .20. 30...or 100, 200, 300...or 1000, 2000, 3000...
'jp: points 150 or 1500 will be reduced to 100 or 1000.

Sub AddScore(Points)
    If Tilted Then Exit Sub
    If Points < 100 Then
        Add10 = Points \ 10
        AddScore10Timer.Enabled = TRUE
    ElseIf Points < 1000 Then
        Add100 = Points \ 100
        AddScore100Timer.Enabled = TRUE
    ElseIf Points < 10000 Then
        Add1000 = Points \ 1000
        AddScore1000Timer.Enabled = TRUE
    Else
        Addpoints 10000
    End If
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
    Score(player) = Score(player) + Points
  if player=1 then scorereel1.addvalue(points)
    if player=2 then scorereel2.addvalue(points)
  If DesktopMode = False Then Controller.B2SSetScore player, score(player) + Points

    ' Sounds: there are 3 sounds: tens, hundreds and thousands
    If Points = 100 AND(Score(player) MOD 1000) \ 100 = 0 Then  'New 1000 reel
        PlaySound "bell1000"
    ElseIf Points = 10 AND(Score(player) MOD 100) \ 10 = 0 Then 'New 100 reel
        PlaySound "bell100"
    Else
        PlaySound "bell" & Points
    End If
    ' check replays
    if score(player)=>replay1 and rep=0 then
      playsound "levelcleared"
    addcredit
      rep=1
    end if
    if score(player)=>replay2 and rep=1 then
      playsound "levelcleared"
      addcredit
      rep=2
    end if
    if score(player)=>replay3 and rep=2 then
      playsound "levelcleared"
      addcredit
      rep=3
    end if
End Sub

Sub addcredit
      credit=credit+1
      if credit>25 then credit=25
    credittxt.text=credit
    If DesktopMode = False Then Controller.B2ssetCredits Credit
      playsound "click"
End sub

Sub CheckTilt
  If Tilttimer.Enabled = True Then
   TiltSens = TiltSens + 1
   if TiltSens = 2 Then
     Tilt = True
    ballsave.timerenabled=0
    shootagain.state=0
     tilttxt.text="TILT"
        If DesktopMode = False Then Controller.B2SSetTilt 33,1
        If DesktopMode = False Then Controller.B2ssetdata 1, 0
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
    ballsave.timerenabled=0
    shootagain.state=0
     tilttxt.text="TILT"
        If DesktopMode = False Then Controller.B2SSetTilt 33,1
        If DesktopMode = False Then Controller.B2ssetdata 1, 0
     playsound "tilt"
     turnoff
End Sub

Sub Tilttimer_Timer()
  Tilttimer.Enabled = False
End Sub

sub turnoff
    bumper1.force=0
    bumper2.force=0
    bumper3.force=0
    tiltseq.play seqalloff
end sub

Sub RightSlingShot_Slingshot
    PlaySoundAtVol "left_slingshot", sling1, 1
  addscore 10
'//////Mike Add
  cball.velx = 10 + 2*RND(1)
    cball.vely = 2*(RND(1)-RND(1))
'//////////
    RSling.Visible = 0
    RSling1.Visible = 1
    sling1.TransZ = -8
    RStep = 1
    RightSlingShot.TimerEnabled = 1
  PLightRL.State = 0
  TLightRL.State = 0
End Sub

Sub RightSlingShot_Timer
    Select Case RStep
        Case 3:RSLing1.Visible = 0:RSLing2.Visible = 1:sling1.TransZ = -8
        Case 4:sling1.TransZ = 0:RSLing2.Visible = 0:RSLing.Visible = 1:RightSlingShot.TimerEnabled = 0:PLightRL.State = 1:TLightRL.State = 1
    End Select
    RStep = RStep + 1
End Sub

Sub LeftSlingShot_Slingshot
    PlaySoundAtVol "right_slingshot", sling2, 1
  addscore 10
'///////////Mike add
  cball2.velx = 10 + 2*RND(1)
    cball2.vely = 2*(RND(1)-RND(1))
'//////////
    LSling.Visible = 0
    LSling1.Visible = 1
    sling2.TransZ = -8
    LStep = 1
    LeftSlingShot.TimerEnabled = 1
  PLightLL.State = 0
  TLightLL.State = 0
End Sub

Sub LeftSlingShot_Timer
    Select Case LStep
        Case 3:LSLing1.Visible = 0:LSLing2.Visible = 1:sling2.TransZ = -8
        Case 4:sling2.TransZ = 0:LSLing2.Visible = 0:LSLing.Visible = 1:LeftSlingShot.TimerEnabled = 0:PLightLL.State = 1:TLightLL.State = 1
    End Select
    LStep = LStep + 1
End Sub

sub bumper1_hit
    if tilt=false then playsoundAtVol "jet2", bumper1, VolBump
    addscore 100
  If FlashB.Enabled = False then
    Bumper1Light.State = 0
    Me.TimerEnabled = 1
  End if
  end sub

sub Bumper1_Timer
  Bumper1Light.State = 1
  Me.Timerenabled = 0
End Sub

sub bumper3_hit
    if tilt=false then playsoundAtVol "jet2", bumper3, VolBump
    addscore 100
  If FlashB.Enabled = False then
    Bumper3Light.State = 0
    Me.TimerEnabled = 1
  End if
  end sub

sub Bumper3_Timer
  Bumper3Light.State = 1
  Me.Timerenabled = 0
End Sub

sub bumper2_hit
    if tilt=false then playsoundAtVol "jet2", Bumper2, VolBump
  if (bumper2light.state)=lightstateon then
    addscore 1000
  else
    addscore 100
    end if
  If FlashB.Enabled = False then
    Bumper2Light.State = 0
    Me.TimerEnabled = 1
  End if
end sub

sub Bumper2_Timer
  Bumper2Light.State = 1
  Me.Timerenabled = 0
End Sub

sub ctrig_hit
    movearrow
    addscore 100
end sub

sub spinner1_spin
    movearrow
    addscore 100
end sub

sub spinner2_spin
    movearrow
    addscore 100
end sub

sub tlt_hit
    if tilt=false then
  FlashBumpers
   playsound "oinkoink"
    If DesktopMode = False Then Controller.B2SStartAnimation "PigNose"
   PlightLU.state=lightstateoff
   TlightLU.state=lightstateoff
   Me.TimerEnabled = 1
   if (ltl.state)=lightstateon then
     addscore 10000
    else
     addscore 1000
   end if
  end if
end sub

sub tlt_Timer
  PlightLU.State=lightstateon
  TlightLU.state=lightstateon
  Me.Timerenabled = 0
End Sub

sub trt_hit
    if tilt=false then
  FlashBumpers
  playsound "oinkoink"
  If DesktopMode = False Then Controller.B2SStartAnimation "PigNose"
  PlightRU.state=lightstateoff
  TlightRU.state=lightstateoff
  Me.TimerEnabled = 1
    if (rtl.state)=lightstateon then
    addscore 10000
    else
    addscore 1000
    end if
  end if
end sub

sub trt_Timer
  PlightRU.State=lightstateon
  TlightRU.state=lightstateon
  Me.Timerenabled = 0
End Sub

sub tlt1_hit
    if tilt=false then
    FlashBumpers
    if (n1b.state)=lightstateon then
      n1a.state=lightstateoff
      numbstate(player,1)=0
      addscore 5000
      checkaward
      else
      addscore 500
    end if
    if apos(player)=1 then
      if spstate(player)=1 then
        playsound "knocke"
        addcredit
        fivekdelay.enabled=true
      else
        fivekdelay.enabled=true
      end if
    end if
    end if
end sub

sub tlt2_hit
    if tilt=false then
    FlashBumpers
    if (n2.state)=lightstateon then
      n2.state=lightstateoff
      numbstate(player,2)=0
      addscore 5000
      checkaward
    else
      addscore 500
    end if
    if apos(player)=2 then
      if spstate(player)=1 then
        playsound "knocke"
        addcredit
        fivekdelay.enabled=true
        else
        fivekdelay.enabled=true
      end if
    end if
    end if
end sub

sub trt2_hit
    if tilt=false then
  FlashBumpers
    if (n3.state)=lightstateon then
    n3.state=lightstateoff
  numbstate(player,3)=0
    addscore 5000
    checkaward
    else
    addscore 500
    end if
    if apos(player)=3 then
    if spstate(player)=1 then
    playsound "knocke"
  addcredit
    fivekdelay.enabled=true
    else
    fivekdelay.enabled=true
    end if
    end if
    end if
end sub

sub trt1_hit
    if tilt=false then
  FlashBumpers
    if (n4b.state)=lightstateon then
    n4a.state=lightstateoff
    numbstate(player,4)=0
    addscore 5000
    checkaward
  else
    addscore 500
    end if
    if apos(player)=4 then
    if spstate(player)=1 then
    playsound "knocke"
  addcredit
    fivekdelay.enabled=true
    else
    fivekdelay.enabled=true
    end if
    end if
    end if
end sub

sub mlt1_hit
    if tilt=false then
  FlashBumpers
    if (n5b.state)=lightstateon then
    n5a.state=lightstateoff
    numbstate(player,5)=0
    addscore 5000
    checkaward
    else
    addscore 500
    end if
    if apos(player)=5 then
    if spstate(player)=1 then
    playsound "knocke"
  addcredit
    fivekdelay.enabled=true
    else
    fivekdelay.enabled=true
    end if
    end if
    end if
end sub

sub mlt2_hit
 if tilt=false then
  FlashBumpers
   if (n6b.state)=lightstateon then
     n6a.state=lightstateoff
  numbstate(player,6)=0
    addscore 5000
    checkaward
    else
    addscore 500
   end if
    if apos(player)=6 then
      if spstate(player)=1 then
        playsound "knocke"
    addcredit
        fivekdelay.enabled=true
      else
        fivekdelay.enabled=true
      end if
    end if
 end if
end sub

sub mlt_hit
    if tilt=false then
  FlashBumpers
  playsound "oink"
  If DesktopMode = False Then Controller.B2SStartAnimation "PigNoseLeft"
  PlightML.state=lightstateoff
  TlightML.state=lightstateoff
  Me.TimerEnabled = 1
    if (n4a.state)=lightstateon then
    n4a.state=lightstateoff
    numbstate(player,4)=0
    addscore 5000
    checkaward
    else
    addscore 500
    end if
    if apos(player)=4 then
    if spstate(player)=1 then
    playsound "knocke"
  addcredit
    fivekdelay.enabled=true
    else
    fivekdelay.enabled=true
    end if
    end if
    end if
end sub

sub mlt_Timer
  PlightML.State=lightstateon
  TlightML.state=lightstateon
  Me.Timerenabled = 0
End Sub

sub mrt_hit
    if tilt=false then
  FlashBumpers
  playsound "oink"
  If DesktopMode = False Then Controller.B2SStartAnimation "PigNoseRight"
  PlightMR.state=lightstateoff
  TlightMR.state=lightstateoff
  Me.TimerEnabled = 1
    if (n1a.state)=lightstateon then
    n1a.state=lightstateoff
    numbstate(player,1)=0
    addscore 5000
    checkaward
    else
    addscore 500
    end if
    if apos(player)=1 then
    if spstate(player)=1 then
    playsound "knocke"
    addcredit
    fivekdelay.enabled=true
    else
    fivekdelay.enabled=true
    end if
    end if
    end if
end sub

sub mrt_Timer
  PlightMR.State=lightstateon
  TlightMR.state=lightstateon
  Me.Timerenabled = 0
End Sub

sub mrt1_hit
    if tilt=false then
  FlashBumpers
    if (n9b.state)=lightstateon then
    n9a.state=lightstateoff
  numbstate(player,9)=0
    addscore 5000
    checkaward
    else
    addscore 500
    end if
    if apos(player)=9 then
    if spstate(player)=1 then
    playsound "knocke"
  addcredit
    fivekdelay.enabled=true
    else
    fivekdelay.enabled=true
    end if
    end if
    end if
end sub

sub mrt2_hit
    if tilt=false then
  FlashBumpers
    if (n10b.state)=lightstateon then
    n10a.state=lightstateoff
  numbstate(player,10)=0
    addscore 5000
    checkaward
    else
    addscore 500
    end if
    if apos(player)=10 then
    if spstate(player)=1 then
      playsound "knocke"
      addcredit
      fivekdelay.enabled=true
      else
      fivekdelay.enabled=true
      end if
    end if
    end if
end sub

sub llt_hit
    if tilt=false then
  FlashBumpers
  playsound "oink"
  If DesktopMode = False Then Controller.B2SStartAnimation "PigNoseLeft"
  PlightLM.state=lightstateoff
  TlightLM.state=lightstateoff
  Me.TimerEnabled = 1
    if (n8.state)=lightstateon then
    n8.state=lightstateoff
  numbstate(player,8)=0
    addscore 5000
    checkaward
    else
    addscore 500
    end if
    if apos(player)=8 then
    if spstate(player)=1 then
    playsound "knocke"
  addcredit
    fivekdelay.enabled=true
    else
    fivekdelay.enabled=true
    end if
    end if
    end if
end sub

sub llt_Timer
  PlightLM.State=lightstateon
  TlightLM.State=lightstateon
  Me.Timerenabled = 0
End Sub

sub lrt_hit
    if tilt=false then
    FlashBumpers
    playsound "oink"
    If DesktopMode = False Then Controller.B2SStartAnimation "PigNoseRight"
    PlightRM.state=lightstateoff
    TlightRM.state=lightstateoff
    Me.TimerEnabled = 1
    if (n7.state)=lightstateon then
      n7.state=lightstateoff
      numbstate(player,7)=0
      addscore 5000
      checkaward
      else
      addscore 500
    end if
    if apos(player)=7 then
      if spstate(player)=1 then
        playsound "knocke"
        addcredit
        fivekdelay.enabled=true
      else
        fivekdelay.enabled=true
    end if
    end if
    end if
end sub

sub lrt_Timer
  PlightRM.State=lightstateon
  TlightRM.State=lightstateon
  Me.Timerenabled = 0
End Sub

sub lout_hit
    if tilt=false then
  FlashBumpers
    if (n10a.state)=lightstateon then
    n10a.state=lightstateoff
  numbstate(player,10)=0
    addscore 5000
    checkaward
    else
    addscore 500
    end if
    if apos(player)=10 then
    if spstate(player)=1 then
    playsound "knocke"
  addcredit
    fivekdelay.enabled=true
    else
    fivekdelay.enabled=true
    end if
    end if
    end if
end sub

sub lin_hit
    if tilt=false then
  FlashBumpers
    if (n9a.state)=lightstateon then
    n9a.state=lightstateoff
  numbstate(player,9)=0
    addscore 5000
    checkaward
    else
    addscore 500
    end if
    if apos(player)=9 then
    if spstate(player)=1 then
    playsound "knocke"
  addcredit
    fivekdelay.enabled=true
    else
    fivekdelay.enabled=true
    end if
    end if
    end if
end sub

sub rin_hit
    if tilt=false then
  FlashBumpers
    if (n6a.state)=lightstateon then
    n6a.state=lightstateoff
    numbstate(player,6)=0
    addscore 5000
    checkaward
    else
    addscore 500
    end if
    if apos(player)=6 then
    if spstate(player)=1 then
    playsound "knocke"
    addcredit
    fivekdelay.enabled=true
    else
    fivekdelay.enabled=true
    end if
    end if
    end if
end sub

sub rout_hit
    if tilt=false then
  FlashBumpers
    if (n5a.state)=lightstateon then
    n5a.state=lightstateoff
    numbstate(player,5)=0
    addscore 5000
    checkaward
    if apos(player)=5 then
      addscore 5000
    end if
  else
    addscore 500
    end if
    if apos(player)=5 then
    if spstate(player)=1 then
    playsound "knocke"
    addcredit
    fivekdelay.enabled=true
    else
    fivekdelay.enabled=true
    end if
    end if
    end if
end sub

sub FlashBumpers
  Bumper1Light.State=0
  Bumper2Light.State=0
  Bumper3Light.State=0
  FlashB.enabled=1
end sub

sub FlashB_timer
  Bumper1Light.State=1
  Bumper2Light.State=1
  Bumper3Light.State=1
  FlashB.enabled=0
end sub

sub fivekdelay_timer
    addscore 5000
  birdcount=0
  playsound "bonuscheer"
  birdjump.enabled=true
    fivekdelay.enabled=false
end sub

sub birdjump_timer
  if toy1timer.enabled + toy1down.enabled + toy2timer.enabled + toy2down.enabled + toy3timer.enabled + toy3down.enabled + toy4timer.enabled + toy4down.enabled + toy5timer.enabled + toy5down.enabled = 0 then
    count1=0
    jump1 = int(rnd*5)+3
    time1 = int(rnd*5)+1
    toy1timer.enabled=true

    count2=0
    jump2 = int(rnd*5)+3
    time2 = int(rnd*5)+1
    toy2timer.enabled=true

    count3=0
    jump3 = int(rnd*5)+3
    time3 = int(rnd*5)+1
    toy3timer.enabled=true

    count4=0
    jump4 = int(rnd*5)+3
    time4 = int(rnd*5)+1
    toy4timer.enabled=true

    count5=0
    jump5 = int(rnd*5)+3
    time5 = int(rnd*5)+1
    toy5timer.enabled=true

  birdcount=birdcount+1
  end if
  if birdcount=5 then me.enabled=0
end sub

sub toy1timer_timer
  toy1.transy=toy1.transy+jump1
  if count1=time1 then
    count1=0
    toy1down.enabled=true
    me.enabled=0
    else
    count1=count1+1
  end if
end sub

sub toy1down_timer
  toy1.transy=toy1.transy-jump1
  if count1=time1 then
    me.enabled=0
    else
    count1=count1+1
  end if
end sub

sub toy2timer_timer
  toy2.transy=toy2.transy+jump2
  if count2=time2 then
    count2=0
    toy2down.enabled=true
    me.enabled=0
    else
    count2=count2+1
  end if
end sub

sub toy2down_timer
  toy2.transy=toy2.transy-jump2
  if count2=time2 then
    me.enabled=0
    else
    count2=count2+1
  end if
end sub

sub toy3timer_timer
  toy3.transy=toy3.transy+jump3
  if count3=time3 then
    count3=0
    toy3down.enabled=true
    me.enabled=0
    else
    count3=count3+1
  end if
end sub

sub toy3down_timer
  toy3.transy=toy3.transy-jump3
  if count3=time3 then
    me.enabled=0
    else
    count3=count3+1
  end if
end sub

sub toy4timer_timer
  toy4.transy=toy4.transy+jump4
  if count4=time4 then
    count4=0
    toy4down.enabled=true
    me.enabled=0
    else
    count4=count4+1
  end if
end sub

sub toy4down_timer
  toy4.transy=toy4.transy-jump4
  if count4=time4 then
    me.enabled=0
    else
    count4=count4+1
  end if
end sub

sub toy5timer_timer
  toy5.transy=toy5.transy+jump5
  if count5=time5 then
    count5=0
    toy5down.enabled=true
    me.enabled=0
    else
    count5=count5+1
  end if
end sub

sub toy5down_timer
  toy5.transy=toy5.transy-jump5
  if count5=time5 then
    me.enabled=0
    else
    count5=count5+1
  end if
end sub

sub movearrow
    if tilt=false then
    for i = 1 to 10
    arrow(apos(player)).state=lightstateoff
    next
      apos(player)=apos(player)+1
      if apos(player)>10 then apos(player)=1
    arrow(apos(player)).state=lightstateon
  end if
end sub

sub checkaward
    ac(player)=0
    for i=1 to 10
      if numbstate(player,i)=0 then ac(player)=ac(player)+1
    next
  if ac(player)>2 then Star1L.state=1
  if ac(player)>5 then Star2L.state=1
    if ac(player)=10 then
    Star3L.state=1
      sp.state=lightstateon
    spstate(player)=1
      sa(player)=1
   else
    Star3L.state=0
      sp.state=0
    spstate(player)=0
      sa(player)=0
    end if
  if star3l.state=2 then
    addscore 10000
    playsound "bonuscheer"
  end if
end sub

sub Star1_hit
  if star1l.state=1 then
    addscore 2000
    else
    addscore 10
  end if
end sub

sub Star2_hit
  if star2l.state=1 then
    addscore 2000
    else
    addscore 10
  end if
'         if star3.timerenabled=0 then        '***** if then used for testing StarWizard mode
'            for i=1 to 10
'             numb(i).state=lightstateoff
'             numbstate(player,i)=0
'             next
'           checkaward
'         end if
end sub

sub Star3_hit
  if star3l.state=1 then
    star3l.state=2
    playsound "StarWizard"
    me.timerenabled=1
  end if
end sub

sub Star3_timer
    Star3L.state=0
    star2L.state=0
    star1l.state=0
      for i=1 to 10
    numb(i).state=lightstateon
    numbstate(player,i)=1
      next
    checkaward
    star3.timerenabled=0
end sub

sub savehs

    savevalue "GoldenBirds", "credit", credit
    savevalue "GoldenBirds", "hiscore", hisc
    savevalue "GoldenBirds", "match", matchnumb

    savevalue "GoldenBirds", "score1", score(1)
    savevalue "GoldenBirds", "arrowpos1", apos(1)
    savevalue "GoldenBirds", "score2", score(2)
    savevalue "GoldenBirds", "arrowpos2", apos(2)

end sub

sub loadhs
    dim temp

  temp = LoadValue("GoldenBirds", "credit")
    If (temp <> "") then credit = CDbl(temp)
    temp = LoadValue("GoldenBirds", "hiscore")
    If (temp <> "") then hisc = CDbl(temp)
    temp = LoadValue("GoldenBirds", "match")
    If (temp <> "") then matchnumb = CDbl(temp)

    temp = LoadValue("GoldenBirds", "score1")
    If (temp <> "") then score(1) = CDbl(temp)
    temp = LoadValue("GoldenBirds", "arrowpos1")
    If (temp <> "") then apos(1) = CDbl(temp)
    temp = LoadValue("GoldenBirds", "score2")
    If (temp <> "") then score(2) = CDbl(temp)
    temp = LoadValue("GoldenBirds", "arrowpos2")
    If (temp <> "") then apos(2) = CDbl(temp)
end sub

sub ballsave_hit
  shootagain.state=2
  me.timerenabled=1
end sub

sub ballsave_timer
  shootagain.state=0
  me.timerenabled=0
end sub

sub ballhome_hit
  ballrenabled=1
end sub

sub ballrel_hit
  if ballrenabled=1 then
    if ballinplay=1 then playsound "redbird"
    if ballinplay=2 then playsound "yellowbird"
    if ballinplay=3 then playsound "bluebird"
    if ballinplay=4 then playsound "whitebird"
    if ballinplay=5 then playsound "blackbird"
    ballrenabled=0
  end if
end sub

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

Const tnob = 13 ' total number of balls
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


'   -------------------------------------------------
'   Super Straight  SONIC
'   -------------------------------------------------
'
' Thalamus 2018-07-24
' Added/Updated "Positional Sound Playback Functions" and "Supporting Ball & Sound Functions"
' No special SSF tweaks yet.

Option Explicit
Randomize

Const cGameName = "SuperStraight"
Const Ballsize = 50

Dim bumperlitscore
Dim bumperoffscore
Dim balls
dim ebcount
Dim replays
Dim Replay1Table
Dim Replay2Table
dim replay1
dim replay2
Dim hisc
Dim Controller
Dim maxplayers
Dim players
Dim player
Dim credit, dbonus
Dim score(4)
Dim Add100, Add1000, Add10000
Dim sreels(4)
Dim Shoot(4), CanPlay(4)
Dim state
Dim tilt
Dim tiltsens
Dim target(9)
Dim holebonus
Dim StarState
Dim Bonus
Dim Bonuslight(19)
Dim ballinplay
Dim matchnumb
dim ballrenabled
Dim rep
Dim rst
Dim eg
Dim bell
Dim i,j,tnumber, objekt, light
Dim awardcheck
Dim DW1, DW2, DW3

On Error Resume Next
ExecuteGlobal GetTextFile("controller.vbs")
If Err Then MsgBox "Can't open controller.vbs"
On Error Goto 0

sub SuperStraight_init
    LoadEM
    maxplayers=4
    Replay1Table=450000
    Replay2Table=630000

    set sreels(1) = ScoreReel1
    set sreels(2) = ScoreReel2
    set sreels(3) = ScoreReel3
    set sreels(4) = ScoreReel4
    set Shoot(1)= Shoot1
    set Shoot(2)= Shoot2
    set Shoot(3)= Shoot3
    set Shoot(4)= Shoot4
    set CanPlay(1) = CanPlay1
    set CanPlay(2) = CanPlay2
    set CanPlay(3) = CANPLAY3
    set CanPlay(4) = CANPLAY4
    set target(1)=Tg1
    set target(2)=Tg2
    set target(3)=Tg3
    set target(4)=Tg4
    set target(5)=Tg5
    set target(6)=Tg6
    set target(7)=Tg7
    set target(8)=Drop1
    set target(9)=Drop2
    set bonuslight(1)=bonus1
    set bonuslight(2)=bonus2
    set bonuslight(3)=bonus3
    set bonuslight(4)=bonus4
    set bonuslight(5)=bonus5
    set bonuslight(6)=bonus6
    set bonuslight(7)=bonus7
    set bonuslight(8)=bonus8
    set bonuslight(9)=bonus9
    set bonuslight(10)=bonus10
    LTop1.state=0
    LRinlane.state=0
    LLinlane.state=0
    LRightSP.state=0
    LLeftSP.state=0
    LABspin.state=0
    Bonus10.state=0
    LSpin.state=0
    LRsideout.state=0
    player=1
    For each light in BumperLights:light.State = 0: Next
    For each light in Cardlights:light.State = 0: Next
    For each light in TGlights:light.State = 1: Next
    For each light in Lights:light.state=1:next
    loadhs
    if hisc="" then hisc=50000
    hstxt.text=hisc
    gamov.text="Game Over"
    gamov.timerenabled=1
    tilttxt.timerenabled=1
    credittxt.setvalue(credit)
    If credit =0 then light23.state=0 else light23.state=1
    if balls="" then balls=5
    if balls<>3 and balls<>5 then balls=5
    if replays="" then replays=2
    if replays<>1 and replays<>2 and replays<>3 then replays=2
    Replay1=Replay1Table
    bumperlitscore=1000
    bumperoffscore=100

    If ShowDT = True Then
        For each objekt in backdropstuff
        Objekt.visible = 1
        Next
    End If

    If ShowDt = False Then
        For each objekt in backdropstuff
        Objekt.visible = 0
        Next
    End If

If B2SOn Then
        setBackglass.enabled=true
    End If
    PlaySound "motor"
    tilt=false
    If credit>0 then DOF 133, DOFOn
End sub

sub setBackglass_timer
    Controller.B2ssetCredits Credit
    Controller.B2ssetMatch 34, Matchnumb*10
    Controller.B2SSetGameOver 35,1
    Controller.B2SSetScorePlayer 5, hisc
    Controller.B2SSetScorePlayer 1, Score(1)
    Controller.B2SSetScorePlayer 2, Score(2)
    me.enabled=false
end sub

sub gamov_timer
    if state=false then
        If B2SOn then Controller.B2SSetGameOver 35,0
        gamov.text=""
        gtimer.enabled=true
    end if
    gamov.timerenabled=0
end sub

sub gtimer_timer
    if state=false then
        gamov.text="Game Over"
        If B2SOn then Controller.B2SSetGameOver 35,1
        gamov.timerenabled=1
        gamov.timerinterval= (INT (RND*10)+5)*100
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

Sub SuperStraight_KeyDown(ByVal keycode)

    if keycode=AddCreditKey then
        playsound "coinin"
        coindelay.enabled=true
    end if

    if keycode=StartGameKey and credit>0 then
      if state=false then
        credit=credit-1
        if credit < 1 then DOF 133, DOFOff
        playsound "cluper"
        credittxt.setvalue(credit)
        If credit =0 then light23.state=0 else light23.state=1
        ballinplay=1
        player=1
        If b2son Then
            Controller.B2ssetCredits Credit
            Controller.B2ssetballinplay 32, Ballinplay
            Controller.B2ssetplayerup 30, 1
            Controller.B2ssetcanplay 31, 1
            Controller.B2SSetGameOver 0
        End If
        shoot(player).state=1
        tilt=false
        state=true
        CanPlay(player).state=1
        playsound "initialize"
        players=1
        rst=0
        resettimer.enabled=true
      else if state=true and players < maxplayers and Ballinplay=1 then
        credit=credit-1
        if credit < 1 then DOF 133, DOFOff
        players=players+1
        credittxt.setvalue(credit)
        if credit =0 then light23.state=0
        If b2son then
            Controller.B2ssetCredits Credit
            Controller.B2ssetcanplay 31, players
        End If
        CanPlay(players).state=1
        playsound "cluper"
       end if
      end if
    end if

    If keycode = PlungerKey Then
        Plunger.PullBack
        PlaySound "plungerpull",0,1,0.25,0.25
    End If

  if tilt=false and state=true then
    If keycode = LeftFlipperKey Then
        LeftFlipper.RotateToEnd
        PlaySound SoundFXDOF("flipperup",101,DOFOn,DOFFlippers), 0, .67, -0.05, 0.05
        PlaySound "Buzz",-1,.05,-0.05, 0.05
    End If

    If keycode = RightFlipperKey Then
        RightFlipper.RotateToEnd
        PlaySound SoundFXDOF("flipperup",102,DOFOn,DOFFlippers), 0, .67, 0.05, 0.05
        PlaySound "Buzz1",-1,.05,0.05,0.05
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
  if keycode = 46 then ' C Key
     If contball = 1 Then
          contball = 0
     Else
          contball = 1
     End If
End If
if keycode = 48 then 'B Key
     If bcboost = 1 Then
          bcboost = bcboostmulti
     Else
          bcboost = 1
     End If
End If
if keycode = 203 then bcleft = 1 ' Left Arrow
if keycode = 200 then bcup = 1 ' Up Arrow
if keycode = 208 then bcdown = 1 ' Down Arrow
if keycode = 205 then bcright = 1 ' Right Arrow
end if
End Sub

Sub SuperStraight_KeyUp(ByVal keycode)

    If keycode = PlungerKey Then
        Plunger.Fire
        PlaySound "plunger",0,1,0.25,0.25
    End If

   If tilt=false and state=true then
    If keycode = LeftFlipperKey Then
        LeftFlipper.RotateToStart
        PlaySound SoundFXDOF("flipperdown",101,DOFOff,DOFFlippers), 0, 1, -0.05, 0.05
        StopSound "Buzz"
    End If

    If keycode = RightFlipperKey Then
        RightFlipper.RotateToStart
        PlaySound SoundFXDOF("flipperdown",102,DOFOff,DOFFlippers), 0, 1, 0.05, 0.05
        StopSound "Buzz1"
    End If
if keycode = 203 then bcleft = 0 ' Left Arrow
if keycode = 200 then bcup = 0 ' Up Arrow
if keycode = 208 then bcdown = 0 ' Down Arrow
if keycode = 205 then bcright = 0 ' Right Arrow
   End if
End Sub

Sub StartControl_Hit()
     Set ControlBall = ActiveBall
     contballinplay = true
End Sub

Sub StopControl_Hit()
     contballinplay = false
End Sub

Dim bcup, bcdown, bcleft, bcright, contball, contballinplay, ControlBall, bcboost
Dim bcvel, bcyveloffset, bcboostmulti

bcboost = 1 'Do Not Change - default setting
bcvel = 4 'Controls the speed of the ball movement
bcyveloffset = -0.01 'Offsets the force of gravity to keep the ball from drifting vertically on the table, should be negative
bcboostmulti = 3 'Boost multiplier to ball veloctiy (toggled with the B key)

Sub BallControl_Timer()
     If Contball and ContBallInPlay then
          If bcright = 1 Then
               ControlBall.velx = bcvel*bcboost
          ElseIf bcleft = 1 Then
               ControlBall.velx = - bcvel*bcboost
          Else
               ControlBall.velx=0
          End If

         If bcup = 1 Then
              ControlBall.vely = -bcvel*bcboost
         ElseIf bcdown = 1 Then
              ControlBall.vely = bcvel*bcboost
         Else
              ControlBall.vely= bcyveloffset
         End If
     End If
End Sub

sub flippertimer_timer()
    LFlip.RotY = LeftFlipper.CurrentAngle
    RFlip.RotY = RightFlipper.CurrentAngle
    FlipperLSh.RotZ = LeftFlipper.currentangle
    FlipperRSh.RotZ = RightFlipper.currentangle
end sub


sub coindelay_timer
    addcredit
    coindelay.enabled=false
end sub

sub resettimer_timer
    rst=rst+1
    for i = 1 to maxplayers
        sreels(i).resettozero

    next
    If b2son then
        for i = 1 to maxplayers
          Controller.B2SSetScorePlayer i, score(i)

        next
    End If
    if rst=18 then
        playsound "ballrelease"
    end if
    if rst=22 then
        newgame
        resettimer.enabled=false
    end if
end sub

Sub addcredit
      credit=credit+1
      DOF 133, DOFOn
      if credit>15 then credit=15
      credittxt.setvalue(credit)
      If b2son Then Controller.B2ssetCredits Credit
      If credit=0 then light23.state=0 else light23.state=1
End sub

Sub Drain_Hit()
    DOF 129, DOFPulse
    PlaySound "drain",0,1,0,0.25
    Drain.DestroyBall
    LTop1.state=0
    Ltop2.state=0
    LRinlane.state=0
    LLinlane.state=0
    LRightSP.state=0
    LLeftSP.state=0
    me.timerenabled=1
End Sub

Sub Drain_timer
    scorebonus.enabled=true
    me.timerenabled=0
End Sub

sub ballhome_hit
    ballrenabled=1
end sub

sub ballhome_unhit
    DOF 134, DOFPulse
end sub

sub ballrel_hit
    if ballrenabled=1 then
        Lshootagain.state=0
        ballrenabled=0
    end if
end sub

sub ScoreBonus_timer
   if tilt=false and bonus>0 then
        if Lbonusx2.state=1 then
            dbonus=2
            me.interval=200
          else
            dbonus=1
            me.interval=125
        End if
        score(player)=score(player)+(1000*dbonus)
        sreels(player).addvalue(1000*dbonus)
        If B2SOn Then Controller.B2SSetScorePlayer player, score(player)
        If dbonus=2 Then
            PlaySound SoundFXDOF("bell2000",143,DOFPulse,DOFChimes)
          Else
            PlaySound SoundFXDOF("bell1000",143,DOFPulse,DOFChimes)
        End If
        checkreplay
        bonuslight(bonus).state=0
        bonus=bonus-1
        LABspin.state=bonus mod 2
        if bonus>0 then bonuslight(bonus).state=1
   else
        bonus=0
        for each light in bonuslights: light.state=0: next
   end if
   if bonus=0 then
     if Lshootagain.state=lightstateon then
        newball
        ballreltimer.enabled=true
     else
      if players=1 or player=players then
        player=1
       Else
        player=player+1
      end if
      If b2son then Controller.B2ssetplayerup 30, player
      for i = 1 to 4: Shoot(i).state=0: Next
      shoot(player).state=1
      nextball
     end if
     scorebonus.enabled=false
   end if
End sub


sub newgame
    bumper1.force=13
    bumper2.force=13
    bumper3.force=13
    resetDT
    player=1
    shoot(player).state=1
    ebcount=0
    for i = 1 to 4:
        score(i)=0
    next
    If b2son then
      for i = 1 to maxplayers
        Controller.B2SSetScorePlayer i, score(i)
      next
    End If
    eg=0
    rep=0
    for each light in bonuslights:light.state=0:next
    addbonus
    LRightSP.state=0
    Lleftsp.state=0
    LExtraBall.state=0
    Lshootagain.state=lightstateoff
    LTop1.state=1
    LRinlane.state=0
    LLinlane.state=0
    LRightSP.state=0
    LLeftSP.state=0
    for each light in bumperlights:light.state=0:next
    for each light in TGlights:light.state=1:next
    for each light in Cardlights:light.state=0:next
    gamov.text=" "
    tilttxt.text=" "
    If b2son then
        Controller.B2SSetGameOver 35,0
        Controller.B2SSetTilt 33,0
        Controller.B2SSetMatch 34,0
    End If
    biptext.text="1"
    matchtxt.text=" "
    BallRelease.CreateBall
    BallRelease.kick 135,4,0
    playsound SoundFXDOF("ballrelease",135,DOFPulse,DOFContactors)
end sub

sub newball
    for each light in bumperlights:light.state=0:next
    for each light in TGlights:light.state=1:next
    for each light in Cardlights:light.state=0:next
    for each light in starlights:light.state=0:next
    for each light in BonusLights:light.state=0:next
    if ebcount=0 then
            resetDT
    end if
    addbonus
    Lextraball.state=0
    LBonusx2.state=0
    Lleftsp.state=0
    LTop1.state=1
    Primitive4.RotX=0
    Lgate.state=0
    LRsideout.state=0
    Lkicker.state=0
    gatewall.isdropped=False
End Sub

Sub resetDT
    playsound SoundFXDOF("BankReset",114,DOFPulse,DOFContactors)
    for each objekt in droptargets: objekt.isdropped=false: Next
    for each light in TGlights:light.state=1:next
end sub

sub nextball
    if tilt=true then
      bumper1.force=0
      bumper2.force=0
      bumper3.force=0
        tilt=false
      tilttxt.text=" "
        If b2son then
            Controller.B2SSetTilt 33,0
            Controller.B2ssetdata 1, 1
        End If
    end if
    ebcount=0
    for i = 1 to 4

    next
    if player=1 then ballinplay=ballinplay+1
    if ballinplay>balls then
        playsound "GameOver"
        eg=1
        ballreltimer.enabled=true
      else
        if state=true and tilt=false then
          newball
          ballreltimer.enabled=true
        end if
        biptext.text=ballinplay
        If b2son then Controller.B2ssetballinplay 32, Ballinplay
    end if
End Sub

sub ballreltimer_timer
  if eg=1 then
      turnoff
      matchnum
      biptext.text=" "
      state=false
      gamov.text="GAME OVER"
      for i=1 to maxplayers
        CanPlay(i).state=0
        Shoot(i).state=0
        if score(i)>hisc then hisc=score(i)
      next
      hstxt.text=hisc
      savehs
      If b2son then
        Controller.B2SSetGameOver 35,1
        Controller.B2ssetballinplay 32, 0
        Controller.B2sStartAnimation "EOGame"
        Controller.B2SSetScorePlayer 5, hisc
        Controller.B2ssetcanplay 31, 0
        Controller.B2ssetPlayerUp 30, 0
      End If
  else
    BallRelease.CreateBall
    BallRelease.kick 135,4,0
    playsound SoundFXDOF("ballrelease",135,DOFPulse,DOFContactors)
  end if
    ballreltimer.enabled=false
end sub
sub matchnum
    'matchnumb=INT(RND*10)
    if matchnumb=0 then
        matchtxt.text="000"
      else
        matchtxt.text=matchnumb*100
    end if
  If b2son then
        If Matchnumb = 0 then
               Controller.B2sSetMatch 100
           else
               Controller.B2SSetMatch Matchnumb*10
       end if
  end if
  For i=1 to players
    if (matchnumb*100)=(score(i) mod 1000) then
    addcredit
    playsound SoundFXDOF("knock",130,DOFPulse,DOFKnocker)
    DOF 131,DOFPulse
    end if
  next
end sub


'slings
Dim LStep, RStep

Sub RSling_Slingshot
    PlaySound SoundFXDOF("Slingshot",104,DOFPulse,DOFContactors), 0, 1, 0.05, 0.05
        DOF 106, DOFPulse
        addscore 100
        if LTop1.state=0 then
            LTop1.state=1
            LTop2.state=0
        else
            LTop1.state=0
            LTop2.state=1
        end if
        if LRinlane.state=1 then
            LRinlane.state=0
            LLinlane.state=1
        else
    if LLinlane.state=1 then
            LRinlane.state=1
            LLinlane.state=0
        end if
        end if
    if LRightSP.state=1 then
            LRightSP.state=0
            LLeftSP.state=1
        else
    if LLeftSP.state=1 then
            LRightSP.state=1
            LLeftSP.state=0
        end if
        end if
    RightSling3.Visible = 1
    Rightsling.visible=0
    SlingR.RotX = 26
    RStep = 0
    RSling.TimerEnabled = 1
End Sub

Sub RSling_Timer
    Select Case RStep
        Case 1:RightSLing3.Visible = 0:RightSLing2.Visible = 1:SlingR.RotX = 14
        Case 2:RightSLing2.Visible = 0:RightSLing.Visible = 1:SlingL.RotX = 2
        Case 3:RightSling.visible = 1:SlingR.RotX = 10:RSling.TimerEnabled = 0
    End Select
    RStep = RStep + 1
End Sub

Sub LSling_Slingshot
    PlaySound SoundFXDOF("Slingshot",103,DOFPulse,DOFContactors),0,1,-0.05,0.05
    DOF 105, DOFPulse
    addscore 100
    if LTop1.state=0 then
            LTop1.state=1
            LTop2.state=0
        else
            LTop1.state=0
            LTop2.state=1
        end if
    if LRinlane.state=1 then
            LRinlane.state=0
            LLinlane.state=1
        else
    if LLinlane.state=1 then
            LRinlane.state=1
            LLinlane.state=0
        end if
        end if
    if LRightSP.state=1 then
            LRightSP.state=0
            LLeftSP.state=1
        else
    if LLeftSP.state=1 then
            LRightSP.state=1
            LLeftSP.state=0
        end if
        end if
    LeftSling2.Visible = 1
    leftsling.visible=0
    SlingL.Rotx=26
    LStep = 0
    Lsling.TimerEnabled = 1
End Sub

Sub LSling_Timer
    Select Case LStep
        Case 1:LeftSLing3.Visible = 0:LeftSLing2.Visible = 1:SlingL.RotX = 14
        Case 2:LeftSling2.Visible = 0:LeftSling.Visible= 1:SlingL.RotX = 2
        Case 3:LeftSLing.Visible = 1:SlingL.RotX = 10:LSling.TimerEnabled = 0
    End Select

    LStep = LStep + 1
End Sub



'********** Bumpers

Sub Bumper1_Hit
   if tilt=false then
    playsound SoundFXDOF("fx_bumper4",107,DOFPulse,DOFContactors)
    DOF 108,DOFPulse
    addscore bumperoffscore
    me.timerenabled=1
   end if
End Sub

Sub Bumper1_timer
    Bumper1Ring.Enabled=0
    BumperRing1.transz=BumperRing1.transz-4
    if BumperRing1.transz=-36 then
        Bumper1Ring.enabled=1
        me.timerenabled=0
    end if
End Sub

Sub Bumper1Ring_timer
    BumperRing1.transz=BumperRing1.transz+4
    If BumperRing1.transz=0 then Bumper1Ring.enabled=0
End sub


Sub Bumper2_Hit
   if tilt=false then
    playsound SoundFXDOF("fx_bumper4",111,DOFPulse,DOFContactors)
    DOF 112,DOFPulse
    if bumper2light.state=1 then
        addscore bumperlitscore
      else
        addscore bumperoffscore
    end if
    me.timerenabled=1
   end if
End Sub

Sub Bumper2_timer
    Bumper2Ring.enabled=0
    BumperRing2.transz=BumperRing2.transz-4
    if BumperRing2.transz=-36 then
        Bumper2Ring.enabled=1
        me.timerenabled=0
    end if
End Sub

Sub Bumper2Ring_timer
    BumperRing2.transz=BumperRing2.transz+4
    If BumperRing2.transz=0 then Bumper2Ring.enabled=0
End sub

Sub Bumper3_Hit
   if tilt=false then
    playsound SoundFXDOF("fx_bumper4",109,DOFPulse,DOFContactors)
    DOF 110,DOFPulse
    if bumper3light.state=1 then
        addscore bumperlitscore
      else
        addscore bumperoffscore
    end if
    me.timerenabled=1
   end if
End Sub

Sub Bumper3_timer
    Bumper3Ring.enabled=0
    BumperRing3.transz=BumperRing3.transz-4
    if BumperRing3.transz=-36 then
        Bumper3Ring.enabled=1
        me.timerenabled=0
    end if
End Sub

Sub Bumper3Ring_timer
    BumperRing3.transz=BumperRing3.transz+4
    If BumperRing3.transz=0 then Bumper3Ring.enabled=0
End sub

'Dingwalls

Sub DW1_Hit
    If tilt=False Then
    addscore 100
    LRsideout.state=1
    if LTop1.state=0 and LTop2.state=1 then
            LTop1.state=1
            LTop2.state=0
        else
            LTop1.state=0
            LTop2.state=1
        end if
    if LRinlane.state=1 then
            LRinlane.state=0
            LLinlane.state=1
        else
    if LLinlane.state=1 then
            LRinlane.state=1
            LLinlane.state=0
        end if
        end if
    if LRightSP.state=1 then
            LRightSP.state=0
            LLeftSP.state=1
        else
    if LLeftSP.state=1 then
            LRightSP.state=1
            LLeftSP.state=0
        end if
        end if
    end if
end sub

Sub DW2_hit
    If tilt=False Then
    addscore 100
    if LTop1.state=0 and LTop2.state=1 then
            LTop1.state=1
            LTop2.state=0
        else
            LTop1.state=0
            LTop2.state=1
        end if
    if LRinlane.state=1 then
            LRinlane.state=0
            LLinlane.state=1
        else
    if LLinlane.state=1 then
            LRinlane.state=1
            LLinlane.state=0
        end if
        end if
    if LRightSP.state=1 then
            LRightSP.state=0
            LLeftSP.state=1
        else
    if LLeftSP.state=1 then
            LRightSP.state=1
            LLeftSP.state=0
        end if
        end if
    end if
end Sub
Sub DW3_hit
    If tilt=False Then
    addscore 100
    if LTop1.state=0 and LTop2.state=1 then
            LTop1.state=1
            LTop2.state=0
        else
            LTop1.state=0
            LTop2.state=1
        end if
    if LRinlane.state=1 then
            LRinlane.state=0
            LLinlane.state=1
        else
    if LLinlane.state=1 then
            LRinlane.state=1
            LLinlane.state=0
        end if
        end if
    if LRightSP.state=1 then
            LRightSP.state=0
            LLeftSP.state=1
        else
    if LLeftSP.state=1 then
            LRightSP.state=1
            LLeftSP.state=0
        end if
        end if
    end if
end sub

sub TopHole_hit
    if Lkicker.state=1 then
    scorebonus1.enabled=true
    else
    addscore 5000
    addbonus
    Pkickarm.rotz=15
    me.timerenabled=1
    end if
end sub

Sub TopHole_timer
    TopHole.Kick 135, 20 + (3*rnd())
    playsound SoundFXDOF("HoleKick",132,DOFPulse,DOFContactors),0,1,0,0.25
    DOF 131, DOFPulse
    Pkickarm.rotz=0
    me.timerenabled=0
End Sub

sub ScoreBonus1_timer
  if tilt=false and bonus>0 then
        if Lbonusx2.state=1 then
            dbonus=2
            me.interval=200
          else
            dbonus=1
            me.interval=125
        End if
        score(player)=score(player)+(1000*dbonus)
        sreels(player).addvalue(1000*dbonus)
        If B2SOn Then Controller.B2SSetScorePlayer player, score(player)
        If dbonus=2 Then
            PlaySound SoundFXDOF("bell2000",143,DOFPulse,DOFChimes)
          Else
            PlaySound SoundFXDOF("bell1000",143,DOFPulse,DOFChimes)
        End If
        bonuslight(bonus).state=0
        bonus=bonus-1
        LABspin.state=bonus mod 2
        if bonus>0 then bonuslight(bonus).state=1
   else
        bonus=0
        scorebonus1.enabled=false
        TopHole.timerenabled=true
        LSpin.state=0
end if

End sub


sub Kicker1_hit
    if LRsideout.state=1 Then
    addscore 5000
    addbonus
    End If
    if LExtraball.state=1 Then
    LExtraball.state=0
    LShootAgain.state=1
End if
    me.timerenabled=1
end sub

sub Kicker1_timer
    Kicker1.kick 20, 40 + (3*rnd())
    playsound SoundFXDOF("HoleKick",139,DOFPulse,DOFContactors),0,1,0,0.25
    DOF 131, DOFPulse
    me.timerenabled=0
end sub
'****** Wire Triggers

Sub Toplane1_Hit
   if tilt=false then
    DOF 115,DOFPulse
        LRsideout.state=1
        if LTop1.state=1 then
        addscore 1000
        addbonus
    else
        addscore 1000
    end if
   end if
End Sub

Sub Toplane2_Hit
   if tilt=false then
    DOF 116,DOFPulse
        addscore 1000
        addbonus
        LRsideout.state=1
    end if
End Sub

Sub Toplane3_Hit
   if tilt=false then
    DOF 117,DOFPulse
        LRsideout.state=1
    if LTop2.state=1 then
        addscore 1000
        addbonus
    else
        addscore 1000
    end if
   end if
End Sub

Sub shooterlane_Hit
   if tilt=false then
    if Lgate.state=1 then
        addscore 10000
        Lgate.state=0
        Primitive4.RotX=0
        gatewall.isdropped=false
    end if
   end if
End Sub

Sub LeftOutlane_Hit
   if tilt=false then
    DOF 124,DOFPulse
      if LLeftSP.state=1 then
        playsound SoundFXDOF("knock",130,DOFPulse,DOFKnocker)
        DOF 131,DOFPulse
        addcredit
      else
        addscore 10000
        addbonus
    end if
   end if
End Sub

Sub LeftInlane_Hit
   if tilt=false then
    DOF 125,DOFPulse
    if LLinlane.state=1 then
        addscore 5000
        addbonus
        LBonusx2.state=1
    else
        addscore 5000
        addbonus
    end if
   end if
End Sub

Sub RightInlane_Hit
   if tilt=false then
    DOF 126,DOFPulse
    if LRinlane.state=1 then
        addscore 5000
        addbonus
        LBonusx2.State=1
      else
        addscore 5000
        addbonus
    end if
   end if
End Sub

Sub RightOutlane_Hit
   if tilt=false then
    DOF 127,DOFPulse
      if LRightSP.state=1 then
        playsound SoundFXDOF("knock",130,DOFPulse,DOFKnocker)
        DOF 131,DOFPulse
        addcredit
      else
        addscore 10000
        addbonus
    end if
   end if
End Sub

'**** Rollovers

Sub Rollover1_Hit
    if tilt=false then
        if Lroll1.state=1 then
            addscore 1000
            LRsideout.state=1
            else
            addscore 100
            LRsideout.state=1
            end if
    end if
end sub

Sub Rollover2_Hit
    if tilt=false then
        if Lroll2.state=1 then
            addscore 1000
          else
            addscore 100
        end if
    end if
end sub

Sub Rollover3_Hit
    if tilt=false then
        if Lroll3.state=1 then
            addscore 1000
          else
            addscore 100
        end if
    end if
end sub

Sub spinner1_spin
    if tilt=False Then
        DOF 119, DOFPulse
        PlaySound "fx_spinner",0,.25,0,0.25
        if lspin.state=1 Then
        addscore 1000
        Else
        addscore 100
        If LABspin.State=1 Then
        LABspin.state=0
        addbonus
        End If
    end If
end if
end Sub

' Targets
Sub Tg1_Hit
    playsound SoundFXDOF("target",120,DOFPulse,DOFContactors)
    If LTg1a.state=0 Then
    addscore 1000
    else
    LTg1a.state=0
    Ltg1b.State=1
    addscore 1000
    if LTg4b.state=1 and LTg3b.State=1 and LTg7b.state=1 and Ldrop2.state=1 Then LRightSP.state=1
    end if
    if ltg5b.state=1 then
    for each light in bumperlights:light.state=1:Next
    end If
    'if LTg4b.state=1 and LTg3b.State=1 and LTg7b.state=1 and Ldrop2.state=1 Then LRightSP.state=1
End Sub

Sub Tg2_Hit
    playsound SoundFXDOF("drop1",120,DOFPulse,DOFContactors)
    If Ltg2a.state=0 then
    addscore 1000
    else
    Ltg2a.state=0
    LTg2b.State=1
    addscore 1000
    if LTg5b.state=1 and LTg3b.State=1 and LTg6b.state=1 and Ldrop1.state=1 Then LExtraBall.state=1
    end if
    if LTg4b.state=1 then Lkicker.state=1
    'if LTg5b.state=1 and LTg3b.State=1 and LTg6b.state=1 and Ldrop1.state=1 Then LExtraBall.state=1
End Sub

Sub Tg3_Hit
    playsound SoundFXDOF("drop1",120,DOFPulse,DOFContactors)
    If LTg3a.state=0 Then
    addscore 1000
    else
    LTg3a.state=0
    LTg3b.state=1
    addscore 1000
    if LTg1b.state=1 and LTg4b.State=1 and LTg7b.state=1 and Ldrop2.state=1 Then LRightSP.state=1
    end if
    For each light in starlights:light.State = 1:next
    if LTg5b.state=1 and LTg2b.State=1 and LTg6b.state=1 and Ldrop1.state=1 Then LExtraBall.state=1
    'if LTg1b.state=1 and LTg4b.State=1 and LTg7b.state=1 and Ldrop2.state=1 Then LRightSP.state=1
End Sub

Sub Tg4_Hit
    playsound SoundFXDOF("drop1",120,DOFPulse,DOFContactors)
    if LTg4a.state=0 Then
    addscore 1000
    else
    LTg4a.state=0
    LTg4b.state=1
    addscore 1000
    if LTg1b.state=1 and LTg3b.State=1 and LTg7b.state=1 and Ldrop2.state=1 Then LRightSP.state=1
    end if
    if LTg2b.state=1 then Lkicker.state=1
    'if LTg1b.state=1 and LTg3b.State=1 and LTg7b.state=1 and Ldrop2.state=1 Then LRightSP.state=1
End Sub

Sub Tg5_Hit
    playsound SoundFXDOF("drop1",120,DOFPulse,DOFContactors)
    if LTg5a.state=0 Then
    addscore 1000
    else
    LTg5a.state=0
    Ltg5b.state=1
    addscore 1000
    if LTg2b.state=1 and LTg3b.State=1 and LTg6b.state=1 and Ldrop1.state=1 Then LExtraBall.state=1
    end if
    if ltg1b.state=1 then
    for each light in bumperlights:light.state=1:Next
    end if
    'if LTg2b.state=1 and LTg3b.State=1 and LTg6b.state=1 and Ldrop1.state=1 Then LExtraBall.state=1
End Sub

Sub Tg6_Hit
    playsound SoundFXDOF("drop1",121,DOFPulse,DOFContactors)
    If LTg6a.state=0 then
    addscore 1000
    else
    LTg6a.state=0
    Ltg6b.state=1
    addscore 1000
    if LShootAgain.state=0 and LTg5b.state=1 and LTg2b.State=1 and LTg3b.state=1 and Ldrop1.state=1 Then LExtraBall.state=1
    if Ltg7b.state=1 Then
    Primitive4.RotX=60
    gatewall.isdropped=true
    Lgate.state=1
    end if
    end if
    'if LShootAgain.state=0 and LTg5b.state=1 and LTg2b.State=1 and LTg3b.state=1 and Ldrop1.state=1 Then LExtraBall.state=1
End Sub

Sub Tg7_Hit
    Playsound SoundFXDOF("drop1",121,DOFPulse,DOFContactors)
    if LTg7a.state=0 then
    addscore 1000
    else
    LTg7a.state=0
    Ltg7b.state=1
    addscore 1000
    if LTg1b.state=1 and LTg4b.State=1 and LTg3b.state=1 and Ldrop2.state=1 Then LRightSP.state=1
    if Ltg6b.state=1 Then
    Primitive4.RotX=60
    gatewall.isdropped=true
    Lgate.state=1
    end if
    end if
    'if LTg1b.state=1 and LTg4b.State=1 and LTg3b.state=1 and Ldrop2.state=1 Then LRightSP.state=1


End Sub

'**** Drop Targets
Sub Drop1_Hit
    If Ldrop2.state=1 then LRinlane.state=1
    if LTg5b.state=1 and LTg2b.State=1 and LTg3b.state=1 and LTg6b.state=1 then LExtraBall.state=1
    playsound SoundFXDOF("drop1",113,DOFPulse,DOFContactors)
    addscore 10000
    Ldrop1.state=1

End Sub

Sub Drop2_Hit
    If Ldrop1.state=1 then LRinlane.state=1
    if LTg1b.state=1 and LTg4b.State=1 and LTg3b.state=1 and LTg7b.state=1 Then LRightSP.state=1
    playsound SoundFXDOF("drop1",113,DOFPulse,DOFContactors)
    addscore 10000
    Ldrop2.state=1

End Sub


sub addscore(points)
  if tilt=false then
      if points=100 then
      matchnumb=matchnumb+1
      if matchnumb>9 then matchnumb=0
     end if
      If Points < 1000 and AddScore100Timer.enabled = false Then
        Add100 = Points \ 100
        AddScore100Timer.Enabled = TRUE
      ElseIf Points < 10000 and AddScore1000Timer.enabled = false Then
        Add1000 = Points \ 1000
        AddScore1000Timer.Enabled = TRUE
      ElseIf AddScore10000Timer.enabled = false Then
        Add10000 = Points \ 10000
        AddScore10000Timer.Enabled = TRUE
    End If
  end if
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

Sub AddScore10000Timer_Timer()
    if Add10000 > 0 then
        AddPoints 10000
        Add10000 = Add10000 - 1
    Else
        Me.Enabled = FALSE
    End If
End Sub

Sub AddPoints(Points)
    score(player)=score(player)+points
    sreels(player).addvalue(points)
    If B2SOn Then Controller.B2SSetScorePlayer player, score(player)

    ' Sounds: there are 3 sounds: tens, hundreds and thousands
     If Points = 1000 AND(Score(player) MOD 100) \ 1000 = 0 Then
        PlaySound SoundFXDOF("bell1000",143,DOFPulse,DOFChimes)
      ElseIf Points = 100 AND(Score(player) MOD 100) \ 100 = 0 Then
        PlaySound SoundFXDOF("bell100",142,DOFPulse,DOFChimes)
      ElseIf points = 10000 Then
        PlaySound SoundFXDOF("bell1000",143,DOFPulse,DOFChimes)
      elseif Points = 1000 Then
        PlaySound SoundFXDOF("bell100",142,DOFPulse,DOFChimes)
      Else
        PlaySound SoundFXDOF("bell10",141,DOFPulse,DOFChimes)

    End If
    checkreplay
end sub

sub checkreplay

    if score(player)=>replay1 and rep=0 then
        playsound SoundFXDOF("knock",130,DOFPulse,DOFKnocker)
        DOF 131,DOFPulse
      addcredit
      rep=1
    end if
    if score(player)=>replay2 and rep=1 then
        playsound SoundFXDOF("knock",130,DOFPulse,DOFKnocker)
        DOF 131,DOFPulse
      addcredit
      rep=2
    end if

end sub

Sub CheckTilt
    If Tilttimer.Enabled = True Then
     TiltSens = TiltSens + 1
     if TiltSens = 3 Then
       Tilt = True
       tilttxt.text="TILT"
        If b2son Then Controller.B2SSetTilt 33,1
        If b2son Then Controller.B2ssetdata 1, 0
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
    DOF 101, DOFOff
    StopSound "Buzz"
    RightFlipper.RotateToStart
    DOF 102, DOFOff
    StopSound "Buzz1"
end sub


Sub addbonus
    bonus=bonus+1
    if bonus > 10 then bonus = 10
    Bonuslight(bonus).state=1
    LABspin.state=bonus mod 2
    LSpin.state=bonus10.state
    if bonus>1 then Bonuslight(bonus-1).state=0
    if bonus>10 then Bonuslight(10).state=1
End sub

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
        If BOT(b).X < SuperStraight.Width/2 Then
            BallShadow(b).X = ((BOT(b).X) - (Ballsize/16) + ((BOT(b).X - (SuperStraight.Width/2))/17))' + 13
        Else
            BallShadow(b).X = ((BOT(b).X) + (Ballsize/16) + ((BOT(b).X - (SuperStraight.Width/2))/17))' - 13
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


Sub Metals_Thin_Hit (idx)
    PlaySound "metalhit_thin", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
End Sub

Sub Metals_Medium_Hit (idx)
    PlaySound "metalhit_medium", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
End Sub

Sub Metals2_Hit (idx)
    PlaySound "metalhit2", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
End Sub

Sub Gates_Hit (idx)
    PlaySound "gate4", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
End Sub

Sub Rubbers_Hit(idx)
    dim finalspeed
    finalspeed=SQR(activeball.velx * activeball.velx + activeball.vely * activeball.vely)
    If finalspeed > 20 then
        PlaySound "fx_rubber2", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
    End if
    If finalspeed >= 6 AND finalspeed <= 20 then
        RandomSoundRubber()
    End If
End Sub

Sub RubberWheel_Hit()
    dim finalspeed
    finalspeed=SQR(activeball.velx * activeball.velx + activeball.vely * activeball.vely)
    If finalspeed > 20 then
        PlaySound "fx_rubber2", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
    End if
    If finalspeed >= 6 AND finalspeed <= 20 then
        RandomSoundRubber()
    End If
End Sub

Sub Posts_Hit(idx)
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

    savevalue "SuperStraight", "credit", credit
    savevalue "SuperStraight", "hiscore", hisc
    savevalue "SuperStraight", "match", matchnumb
    savevalue "SuperStraight", "score1", score(1)
    savevalue "SuperStraight", "score2", score(2)
    savevalue "SuperStraight", "score3", score(3)
    savevalue "SuperStraight", "score4", score(4)
    savevalue "SuperStraight", "replays", replays
    savevalue "SuperStraight", "balls", balls

end sub

sub loadhs
    dim temp
    temp = LoadValue("SuperStraight", "credit")
    If (temp <> "") then credit = CDbl(temp)
    temp = LoadValue("SuperStraight", "hiscore")
    If (temp <> "") then hisc = CDbl(temp)
    temp = LoadValue("SuperStraight", "match")
    If (temp <> "") then matchnumb = CDbl(temp)
    temp = LoadValue("SuperStraight", "score1")
    If (temp <> "") then score(1) = CDbl(temp)
    temp = LoadValue("SuperStraight", "score2")
    If (temp <> "") then score(2) = CDbl(temp)
    temp = LoadValue("SuperStraight", "score3")
    If (temp <> "") then score(3) = CDbl(temp)
    temp = LoadValue("SuperStraight", "score4")
    If (temp <> "") then score(4) = CDbl(temp)
    temp = LoadValue("SuperStraight", "replays")
    If (temp <> "") then replays = CDbl(temp)
    temp = LoadValue("SuperStraight", "balls")
    If (temp <> "") then balls = CDbl(temp)
end sub

Sub SuperStraight_Exit()
    turnoff
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

'*****************************************
'      JP's VP10 Rolling Sounds
'*****************************************

Const tnob = 2 ' total number of balls
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
  If Table1.VersionMinor > 3 OR Table1.VersionMajor > 10 Then
    PlaySound("fx_collide"), 0, Csng(velocity) ^2 / 200, Pan(ball1), 0, Pitch(ball1), 0, 0, AudioFade(ball1)
  Else
    PlaySound("fx_collide"), 0, Csng(velocity) ^2 / 200, Pan(ball1), 0, Pitch(ball1), 0, 0
  End if
End Sub


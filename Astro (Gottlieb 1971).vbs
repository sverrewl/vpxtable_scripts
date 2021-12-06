'///////////////////////////////////////////////////////////
'     ASTRO by Gottlieb (1971)
'
'
'
' vp10 assembled and scripted by BorgDog, 2015
' thanks to Plumb for the playfield scan on VPF
' DOF by BorgDog, reviewed by Arngrim :)
'
'///////////////////////////////////////////////////////////

Option Explicit
Randomize

' Thalamus 2018-07-19
' Added/Updated "Positional Sound Playback Functions" and "Supporting Ball & Sound Functions"

' Thalamus 2018-10-25 : Improved directional sounds
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
Const VolRol    = 1    ' Rollovers volume.
Const VolGates  = 1    ' Gates volume.
Const VolMetal  = 1    ' Metals volume.
Const VolRB     = 1    ' Rubber bands volume.
Const VolRH     = 1    ' Rubber hits volume.
Const VolPo     = 1    ' Rubber posts volume.
Const VolPi     = 1    ' Rubber pins volume.
Const VolPlast  = 1    ' Plastics volume.
Const VolTarg   = 1    ' Targets volume.
Const VolWood   = 1    ' Woods volume.
Const VolKick   = 1    ' Kicker volume.
Const VolSpin   = 1.5  ' Spinners volume.
Const VolFlip   = 1    ' Flipper volume.

Const cGameName = "startrek_1971"

Dim operatormenu, options
Dim bumperlitscore
Dim bumperoffscore
Dim balls
dim award
Dim replays
Dim Add10, Add100, Add1000
Dim Replay1Table(2)
Dim Replay2Table(2)
dim replay1
dim replay2
Dim hisc
Dim Controller
Dim maxplayers
Dim players
Dim player
Dim credit
Dim score(4)
Dim sreels(4)
Dim cplay(4)
Dim state
Dim tilt
Dim tiltsens
Dim target(9)
Dim DTprim(9)
Dim holebonus
Dim holeb(4)
Dim StarState
Dim Bonus
Dim ballstoplay
dim ballrenabled
dim rlight
Dim rep(2)
Dim rst
Dim eg
Dim bell
Dim i,j, objekt, light
Dim awardcheck
Dim rstep, lstep

On Error Resume Next
ExecuteGlobal GetTextFile("controller.vbs")
If Err Then MsgBox "Can't open controller.vbs"
On Error Goto 0


sub Astro_init
  LoadEM
  maxplayers=1
  Replay1Table(1)=20000
  Replay1Table(2)=30000
  Replay2Table(1)=40000
  Replay2Table(2)=50000
  set sreels(1) = ScoreReel1
  hideoptions
  player=1
  For each light in LaneLights:light.State = 0: Next
  For each light in BumperLights:light.State = 1: Next
  For each light in Tlights:light.State = 0: Next
  For each light in starlights:light.State = 0: Next
  loadhs
  if hisc="" then hisc=20000
  hstxt.text=hisc
  gamov.text="Game Over"
  tilttxt.text="TILT"
  gamov.timerenabled=1
  tilttxt.timerenabled=1
  if balls="" then balls=5
  if balls<>3 and balls<>5 and balls<>8 then balls=5
  if replays="" then replays=2
  if replays<>1 and replays<>2 then replays=2
  Replay1=Replay1Table(Replays)
  Replay2=Replay2Table(Replays)
  RepCard.image = "ReplayCard1"
  OptionBalls.image="OptionsBalls"&Balls
  OptionReplays.image="OptionsReplays"&replays
  RepCard.image = "ReplayCard"&replays
  if balls=3 then
    bumperlitscore=1000
    bumperoffscore=100
    InstCard.image="InstCard3balls"
  end if
  if balls=5 then
    bumperlitscore=1000
    bumperoffscore=100
    InstCard.image="InstCard5balls"
  end if
  if balls=8 then
    bumperlitscore=1000
    bumperoffscore=100
    InstCard.image="InstCard8balls"
  end if
  If B2SOn then
    setBackglass.enabled=true
    for each objekt in backdropstuff : objekt.visible = 0 : next
  End If
  for i = 1 to maxplayers
    sreels(i).setvalue(score(i))
  next
  PlaySound "motor"
  tilt=false
  If credit>0 then DOF 136, DOFOn
End sub

sub setBackglass_timer
    Controller.B2SSetGameOver 35,1
    Controller.B2SSetScorePlayer 5, hisc
    Controller.B2SSetScorePlayer 1, Score(1) MOD 100000
  dim objekt : for each objekt in backdropstuff : objekt.visible = 0 : next
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

Sub Astro_KeyDown(ByVal keycode)

  if keycode=AddCreditKey then
    playsoundAtVol "coin3", Drain ,1
    coindelay.enabled=true
    end if

    if keycode=StartGameKey and credit>0 and state=false then
    credit=credit-1
    if credit < 1 then DOF 136, DOFOff
    ballstoplay=balls
    If B2SOn Then
      Controller.B2sStartAnimation "startup"
      Controller.B2ssetballinplay 32, Ballstoplay
      Controller.B2ssetplayerup 30, 1
      Controller.B2SSetGameOver 0
      Controller.B2SSetScorePlayer 5, hisc
    End If
    tilt=false
    state=true
    playsound "initialize"
    players=1
    rst=0
    resettimer.enabled=true
  end if

  If keycode = PlungerKey Then
    Plunger.PullBack
    PlaySoundAtVol "plungerpull", Plunger, 1
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
      if Balls=3 then
        Balls=5
        InstCard.image="InstCard5balls"
        elseif balls=5 Then
        Balls=8
        InstCard.image="InstCard8balls"
        else
        Balls=3
        InstCard.image="InstCard3balls"
      end if
      OptionBalls.image = "OptionsBalls"&Balls
    Case 2:
      Replays=Replays+1
      if Replays>2 then
        Replays=1
      end if
      Replay1=Replay1Table(Replays)
      Replay2=Replay2Table(Replays)
      OptionReplays.image = "OptionsReplays"&replays
      RepCard.image = "ReplayCard"&replays
    Case 3:
      OperatorMenu=0
      savehs
      HideOptions
    End Select
  End If

  if tilt=false and state=true then
  If keycode = LeftFlipperKey Then
    LeftFlipper.RotateToEnd
    PlaySoundAtVol SoundFXDOF("flipperup",101,DOFOn,DOFContactors), LeftFlipper, VolFlip
    PlaySoundAtVol "Buzz", LeftFlipper, VolFlip
  End If

  If keycode = RightFlipperKey Then
    RightFlipper.RotateToEnd
    PlaySoundAtVol SoundFXDOF("flipperup",102,DOFOn,DOFContactors), RightFlipper, VolFlip
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

  If keycode = MechanicalTilt Then
    mechchecktilt
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
  OptionBalls.visible = True
    OptionReplays.visible = True
End Sub

Sub HideOptions
  for each objekt In OptionMenu
    objekt.visible = false
  next
End Sub


Sub Astro_KeyUp(ByVal keycode)

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
    PlaySoundAtVol SoundFXDOF("flipperdown",101,DOFOff,DOFContactors), LeftFlipper, VolFlip
    StopSound "Buzz"
  End If

  If keycode = RightFlipperKey Then
    RightFlipper.RotateToStart
    PlaySoundAtVol SoundFXDOF("flipperdown",102,DOFOff,DOFContactors), RightFlipper, VolFlip
    StopSound "Buzz1"
  End If
   End if
End Sub

sub flippertimer_timer()
  LFlip.RotY = LeftFlipper.CurrentAngle-90
  RFlip.RotY = RightFlipper.CurrentAngle-90
  PGate.Rotz = gate.CurrentAngle+25
end sub

Sub PairedlampTimer_timer

  if credit>0 then
    creditlight.state=1
    else
    creditlight.state=0
  end if
  LStarS2.state = LStarS1.state
  LStarT2.state = LStarT1.state
  LStarA2.state = LStarA1.state
  LStarR2.state = LStarR1.state
  LTrekT2.state = LTrekT1.state
  LTrekR2.state = LTrekR1.state
  LTrekE2.state = LTrekE1.state
  LTrekK2.state = LTrekK1.state
  LGreenStar1.state = LGreenStar.state
  LPurpleStar1.state = LPurpleStar.state
  LRwow.state = LLwow.state
  Bumper1Light1.state = bumper1light.state
  Bumper2Light1.state = bumper2light.state
  Bumper3Light.state = bumper2light.state
  Bumper3Light1.state = bumper3light.state
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
    If B2SOn then
    for i = 1 to maxplayers
      Controller.B2SSetScorePlayer i, score(i)
    next
  End If
    if rst=18 then
    playsoundat "kickerkick", BallRelease
    end if
    if rst=22 then
    newgame
    resettimer.enabled=false
    end if
end sub

Sub addcredit
      credit=credit+1
    DOF 136, DOFOn
      if credit>15 then credit=15
End sub

Sub Drain_Hit()
  DOF 132, DOFPulse
  PlaySoundAtVol "drain", Drain, 1
  Drain.DestroyBall
  me.timerenabled=1
End Sub

Sub Drain_timer
  nextball
  me.timerenabled=0
End Sub

sub ballhome_hit
  ballrenabled=1
end sub

sub ballhome_unhit
  DOF 137, DOFPulse
end sub

sub ballrel_hit
  if ballrenabled=1 then
    ballrenabled=0
  end if
end sub

sub newgame
  bumper1.force=8
  bumper2.force=8
  bumper3.force=8
  player=1
    score(1)=0
  award=0
  If B2SOn then
    for i = 1 to maxplayers
    Controller.B2SSetScorePlayer i, score(i)
    next
  End If
    eg=0
    rep(player)=0
  for each light in bumperlights:light.state=0:next
  for each light in tlights:light.state=1:next
  for each light in lanelights:light.state=1:next
  for each light in treklights:light.state=1:next
  for each light in starlights:light.state=0:next
    gamov.text=" "
    tilttxt.text=" "
    If B2SOn then
    Controller.B2SSetGameOver 35,0
    Controller.B2SSetTilt 33,0
  End If
  biptext.text=ballstoplay
    BallRelease.CreateBall
    BallRelease.kick 45,8,0
  playsoundAtVol SoundFXDOF("kickerkick",131,DOFPulse,DOFContactors), BallRelease, VolKick
end sub

sub nextball
    if tilt=true then
    bumper1.force=8
    bumper2.force=8
    bumper3.force=8
      tilt=false
      tilttxt.text=" "
    If B2SOn then
      Controller.B2SSetTilt 33,0
      Controller.B2ssetdata 1, 1
    End If
    end if
  if player=1 then ballstoplay=ballstoplay-1
  if ballstoplay=0 then
    playsound "motorleer"
    eg=1
    ballreltimer.enabled=true
    else
    if state=true and tilt=false then
      ballreltimer.enabled=true
    end if
    biptext.text=ballstoplay
    If B2SOn then Controller.B2ssetballinplay 32, Ballstoplay
  end if
End Sub

sub ballreltimer_timer
  if eg=1 then
    biptext.text=" "
    turnoff
    state=false
    gamov.text="GAME OVER"
    gamov.timerinterval=1000
    tilttxt.timerinterval=1000
      for i = 1 to maxplayers
    if score(i)>hisc then hisc=score(i)
    next
    hstxt.text=hisc
    savehs
    If B2SOn then
        Controller.B2SSetGameOver 35,1
        Controller.B2ssetballinplay 32, 0
      Controller.B2sStartAnimation "EOGame"
      Controller.B2SSetScorePlayer 5, hisc
      Controller.B2ssetcanplay 31, 0
      Controller.B2ssetcanplay 30, 0
    End If
    For each light in LaneLights:light.State = 0: Next
    For each light in Tlights:light.State = 0: Next
    gamov.timerenabled=1
    tilttxt.timerenabled=1
  else
  if award=1 then
    for each light in spotlights:light.state = 1: Next
    for each light in spotoff:light.state = 0: Next
    award = 0
  end if
    BallRelease.CreateBall
  BallRelease.kick 45,8,0
  playsoundAtVol SoundFXDOF("kickerkick",131,DOFPulse,DOFContactors), BallRelease, VolKick
  end if
  ballreltimer.enabled=false
end sub

'********** Bumpers

Sub Bumper1_Hit
   if tilt=false then
  playsoundAtVol SoundFXDOF("fx_bumper4",107,DOFPulse,DOFContactors), Bumper1, VolBump
  DOF 108,DOFPulse
  if bumper1light.state=1 then
    addscore bumperlitscore
    else
    addscore bumperoffscore
  end if
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
  playsoundAtVol SoundFXDOF("fx_bumper4",109,DOFPulse,DOFContactors), Bumper2, VolBump
  DOF 110,DOFPulse
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
  playsoundAtVol SoundFXDOF("fx_bumper4",111,DOFPulse,DOFContactors), Bumper3, VolBump
  DOF 112,DOFPulse
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

sub FlashBumpers
  For each light in BumperLights
    Light.State=0
  next
  FlashB.enabled=1
end sub

sub FlashB_timer
  For each light in BumperLights
    Light.State=1
  next
  FlashB.enabled=0
end sub

Sub RightSlingShot_Slingshot
  playsoundAtVol SoundFXDOF("left_slingshot",105,DOFPulse,DOFContactors), slingR, 1
  DOF 106,DOFPulse
  addscore 10
    RSling.Visible = 0
    RSling1.Visible = 1
    slingR.objroty = -15
    RStep = 1
    RightSlingShot.TimerEnabled = 1
End Sub

Sub RightSlingShot_Timer
    Select Case RStep
        Case 3:RSLing1.Visible = 0:RSLing2.Visible = 1:slingR.objroty = -7
        Case 4:slingR.objroty = 0:RSLing2.Visible = 0:RSLing.Visible = 1:RightSlingShot.TimerEnabled = 0
    End Select
    RStep = RStep + 1
End Sub

Sub LeftSlingShot_Slingshot
  playsoundAtVol SoundFXDOF("right_slingshot",103,DOFPulse,DOFContactors), slingL, 1
  DOF 104,DOFPulse
  addscore 10
    LSling.Visible = 0
    LSling1.Visible = 1
    slingL.objroty = 15
    LStep = 1
    LeftSlingShot.TimerEnabled = 1
End Sub

Sub LeftSlingShot_Timer
    Select Case LStep
        Case 3:LSLing1.Visible = 0:LSLing2.Visible = 1:slingL.objroty = 7
        Case 4:slingL.objroty = 0:LSLing2.Visible = 0:LSLing.Visible = 1:LeftSlingShot.TimerEnabled = 0
    End Select
    LStep = LStep + 1
End Sub


'****** Triggers & Targets

Sub TStarS1_Hit
   if tilt=false then
  DOF 113,DOFPulse
  addscore 500
  LStarS1.state = 0
  LStarS3.state = 1
   end if
 End Sub

Sub TStarT1_Hit
   if tilt=false then
  DOF 114,DOFPulse
  addscore 500
  LStarT1.state = 0
  LStarT3.state = 1
   end if
End Sub

Sub TStarA1_Hit
   if tilt=false then
  DOF 115,DOFPulse
  addscore 500
  LStarA1.state = 0
  LStarA3.state = 1
   end if
End Sub

Sub TStarR1_Hit
   if tilt=false then
  DOF 116,DOFPulse
  addscore 500
  LStarR1.state = 0
  LStarR3.state = 1
   end if
End Sub

Sub TStarS2_Hit
   if tilt=false then
  DOF 121,DOFPulse
  addscore 500
  LStarS1.state = 0
  LStarS3.state = 1
   end if
End Sub

Sub TStarT2_Hit
   if tilt=false then
  DOF 122,DOFPulse
  addscore 500
  LStarT1.state = 0
  LStarT3.state = 1
   end if
End Sub

Sub TStarA2_Hit
   if tilt=false then
  DOF 124,DOFPulse
  addscore 500
  LStarA1.state = 0
  LStarA3.state = 1
   end if
End Sub

Sub TStarR2_Hit
   if tilt=false then
  DOF 123,DOFPulse
  addscore 500
  LStarR1.state = 0
  LStarR3.state = 1
   end if
End Sub

Sub TTrekT1_Hit
   if tilt=false then
  DOF 125,DOFPulse
  addscore 500
  LTrekT1.state = 0
  LTrekT3.state = 1
   end if
End Sub

Sub TTrekR1_Hit
   if tilt=false then
  DOF 125,DOFPulse
  addscore 500
  LTrekR1.state = 0
  LTrekR3.state = 1
   end if
End Sub

Sub TTrekE1_Hit
   if tilt=false then
  DOF 126,DOFPulse
  addscore 500
  LTrekE1.state = 0
  LTrekE3.state = 1
   end if
End Sub

Sub TTrekK1_Hit
   if tilt=false then
  DOF 126,DOFPulse
  addscore 500
  LTrekK1.state = 0
  LTrekK3.state = 1
   end if
End Sub

Sub TTrekT2_Hit
   if tilt=false then
  DOF 118,DOFPulse
  addscore 500
  LTrekT1.state = 0
  LTrekT3.state = 1
   end if
End Sub

Sub TTrekR2_Hit
   if tilt=false then
  DOF 129,DOFPulse
  addscore 500
  LTrekR1.state = 0
  LTrekR3.state = 1
   end if
End Sub

Sub TTrekE2_Hit
   if tilt=false then
  DOF 130,DOFPulse
  addscore 500
  LTrekE1.state = 0
  LTrekE3.state = 1
   end if
End Sub

Sub TTrekK2_Hit
   if tilt=false then
  DOF 120,DOFPulse
  addscore 500
  LTrekK1.state = 0
  LTrekK3.state = 1
   end if
End Sub

Sub TLwow_Hit
   if tilt=false then
  DOF 127,DOFPulse
  addscore 10
  LGreenStar.state = 0
  LGreenStar3.state = 1
  if LLwow.state=1 then addballs
   end if
End Sub

Sub TLstar_Hit
   if tilt=false then
  DOF 117,DOFPulse
  addscore 10
  LGreenStar.state = 0
  LGreenStar3.state = 1
   end if
End Sub

Sub TRwow_Hit
   if tilt=false then
  DOF 128,DOFPulse
  addscore 10
  LPurpleStar.state = 0
  LPurpleStar3.state = 1
  if LLwow.state=1 then addballs
   end if
End Sub

Sub TRstar_Hit
   if tilt=false then
  DOF 119,DOFPulse
  addscore 10
  LPurpleStar.state = 0
  LPurpleStar3.state = 1
   end if
End Sub

Sub AwardCheckTimer_timer
  Dim starcheck, trekcheck
  starcheck = LStarS3.state + LStarT3.state + LStarA3.state + LStarR3.state
  trekcheck = LTrekT3.state + LTrekR3.state + LTrekE3.state + LTrekK3.state
  if starcheck = 4 then Bumper1Light.state = 1
  if trekcheck = 4 then Bumper2Light.state = 1
  if starcheck + trekcheck + LGreenStar3.state + LPurpleStar3.state = 10 then
    if award=0 then
      addballs
      LLwow.state = 1
      award=1
    end if
  end if
End Sub

sub addballs
  ballstoplay=ballstoplay+1
  if ballstoplay>10 then ballstoplay=10
  biptext.text=ballstoplay
  If B2SOn then Controller.B2ssetballinplay 32, Ballstoplay
  playsound SoundFXDOF("knocker",133,DOFPulse,DOFKnocker)
  DOF 134,DOFPulse
end sub

sub addscore(points)
  if tilt=false then
  If points=10 or points=100 or points=1000 Then
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
  end If
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
  If B2SOn Then Controller.B2SSetScorePlayer player, score(player)

    ' Sounds: there are 3 sounds: tens, hundreds and thousands
    If Points = 100 AND(Score(player) MOD 1000) \ 100 = 0 Then  'New 1000 reel
        PlaySoundat SoundFXDOF("ding3",143,DOFPulse,DOFChimes),Primitive21
      ElseIf Points = 10 AND(Score(player) MOD 100) \ 10 = 0 Then 'New 100 reel
        PlaySoundat SoundFXDOF("ding2",142,DOFPulse,DOFChimes),Primitive21
      ElseIf points = 1000 Then
        PlaySoundat SoundFXDOF("ding3",143,DOFPulse,DOFChimes),Primitive21
    elseif Points = 100 Then
        PlaySoundat SoundFXDOF("ding2",142,DOFPulse,DOFChimes),Primitive21
      Else
        PlaySoundat SoundFXDOF("ding1",141,DOFPulse,DOFChimes),Primitive21
  End If
  checkreplay
end sub

sub checkreplay
    if score(player)=>replay1 and rep(player)=0 then
    addballs
    rep(player)=1
    end if
    if score(player)=>replay2 and rep(player)=1 then
    addballs
    rep(player)=2
    end if
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

Sub MechCheckTilt
  Tilt = True
  tilttxt.text="TILT"
    If B2SOn Then Controller.B2SSetTilt 33,1
    If B2SOn Then Controller.B2ssetdata 1, 0
  playsound "tilt"
  turnoff
  LeftFlipper.RotateToStart
  RightFlipper.RotateToStart
End Sub

Sub Tilttimer_Timer()
  Tilttimer.Enabled = False
End Sub

sub turnoff
    bumper1.force=0
    bumper2.force=0
  bumper3.force=0
    LeftFlipper.RotateToStart
  DOF 101, DOFOff
  StopSound "Buzz"
  RightFlipper.RotateToStart
  DOF 102, DOFOff
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
  PlaySound "fx_spinner", Spinner, VolSpin
End Sub

Sub Rubbers_Hit(idx)
  dim finalspeed
    finalspeed=SQR(activeball.velx * activeball.velx + activeball.vely * activeball.vely)
  If finalspeed > 20 then
    PlaySound "fx_rubber2", 0, Vol(ActiveBall)*VolRH, AudioPan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
  End if
  If finalspeed >= 6 AND finalspeed <= 20 then
    RandomSoundRubber()
  End If
End Sub

Sub Posts_Hit(idx)
  dim finalspeed
    finalspeed=SQR(activeball.velx * activeball.velx + activeball.vely * activeball.vely)
  If finalspeed > 16 then
    PlaySound "fx_rubber2", 0, Vol(ActiveBall)*VolRH, AudioPan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
  End if
  If finalspeed >= 6 AND finalspeed <= 16 then
    RandomSoundRubber()
  End If
End Sub

Sub RandomSoundRubber()
  Select Case Int(Rnd*3)+1
    Case 1 : PlaySound "rubber_hit_1", 0, Vol(ActiveBall)*VolRH, AudioPan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
    Case 2 : PlaySound "rubber_hit_2", 0, Vol(ActiveBall)*VolRH, AudioPan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
    Case 3 : PlaySound "rubber_hit_3", 0, Vol(ActiveBall)*VolRH, AudioPan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
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
    Case 1 : PlaySound "flip_hit_1", 0, Vol(ActiveBall)*VolRH, AudioPan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
    Case 2 : PlaySound "flip_hit_2", 0, Vol(ActiveBall)*VolRH, AudioPan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
    Case 3 : PlaySound "flip_hit_3", 0, Vol(ActiveBall)*VolRH, AudioPan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
  End Select
End Sub



sub savehs
    savevalue "Astro", "credit", credit
    savevalue "Astro", "hiscore", hisc
    savevalue "Astro", "score1", score(1)
  savevalue "Astro", "replays", replays
  savevalue "Astro", "balls", balls
end sub

sub loadhs
    dim temp
  temp = LoadValue("Astro", "credit")
    If (temp <> "") then credit = CDbl(temp)
    temp = LoadValue("Astro", "hiscore")
    If (temp <> "") then hisc = CDbl(temp)
    temp = LoadValue("Astro", "score1")
    If (temp <> "") then score(1) = CDbl(temp)
    temp = LoadValue("Astro", "replays")
    If (temp <> "") then replays = CDbl(temp)
    temp = LoadValue("Astro", "balls")
    If (temp <> "") then balls = CDbl(temp)
end sub

Sub Astro_Exit()
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

Function AudioFade(tableobj) ' Fades between front and back of the table (for surround systems or 2x2 speakers, etc), depending on the Y position on the table. "Astro" is the name of the table
  Dim tmp
  tmp = tableobj.y * 2 / Astro.height-1
  If tmp > 0 Then
    AudioFade = Csng(tmp ^10)
  Else
    AudioFade = Csng(-((- tmp) ^10) )
  End If
End Function

Function AudioPan(tableobj) ' Calculates the pan for a tableobj based on the X position on the table. "Astro" is the name of the table
  Dim tmp
  tmp = tableobj.x * 2 / Astro.width-1
  If tmp > 0 Then
    AudioPan = Csng(tmp ^10)
  Else
    AudioPan = Csng(-((- tmp) ^10) )
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

' Thalamus : Exit in a clean and proper way
Sub Astro_exit()
  Controller.Pause = False
  Controller.Stop
End Sub


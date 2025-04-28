'-----------------------------------------------------------------------------------------------
'-¦¦¦¦¦¦----------¦¦¦¦¦¦-¦¦¦¦¦¦¦¦¦¦-¦¦¦¦¦¦¦¦¦¦¦¦¦¦-¦¦¦¦¦¦¦¦¦¦¦¦¦¦----------------¦¦¦¦¦¦¦¦¦¦¦¦¦¦-
'-¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦-¦¦¦¦¦¦¦¦¦¦-¦¦¦¦¦¦¦¦¦¦¦¦¦¦-¦¦¦¦¦¦¦¦¦¦¦¦¦¦----------------¦¦¦¦¦¦¦¦¦¦¦¦¦¦-
'-¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦-¦¦¦¦¦¦¦¦¦¦-¦¦¦¦¦¦¦¦¦¦¦¦¦¦-¦¦¦¦¦¦¦¦¦¦¦¦¦¦----------------¦¦¦¦¦¦¦¦¦¦¦¦¦¦-
'-¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦---¦¦¦¦¦¦---¦¦¦¦¦¦---------¦¦¦¦¦¦------------------------¦¦¦¦¦¦--¦¦¦¦¦¦-
'-¦¦¦¦¦¦--¦¦¦¦¦¦--¦¦¦¦¦¦---¦¦¦¦¦¦---¦¦¦¦¦¦¦¦¦¦¦¦¦¦-¦¦¦¦¦¦¦¦¦¦¦¦¦¦-¦¦¦¦¦¦¦¦¦¦¦¦¦¦-¦¦¦¦¦¦--¦¦¦¦¦¦-
'-¦¦¦¦¦¦--¦¦¦¦¦¦--¦¦¦¦¦¦---¦¦¦¦¦¦---¦¦¦¦¦¦¦¦¦¦¦¦¦¦-¦¦¦¦¦¦¦¦¦¦¦¦¦¦-¦¦¦¦¦¦¦¦¦¦¦¦¦¦-¦¦¦¦¦¦--¦¦¦¦¦¦-
'-¦¦¦¦¦¦--¦¦¦¦¦¦--¦¦¦¦¦¦---¦¦¦¦¦¦---¦¦¦¦¦¦¦¦¦¦¦¦¦¦-¦¦¦¦¦¦¦¦¦¦¦¦¦¦-¦¦¦¦¦¦¦¦¦¦¦¦¦¦-¦¦¦¦¦¦--¦¦¦¦¦¦-
'-¦¦¦¦¦¦----------¦¦¦¦¦¦---¦¦¦¦¦¦-----------¦¦¦¦¦¦---------¦¦¦¦¦¦----------------¦¦¦¦¦¦--¦¦¦¦¦¦-
'-¦¦¦¦¦¦----------¦¦¦¦¦¦-¦¦¦¦¦¦¦¦¦¦-¦¦¦¦¦¦¦¦¦¦¦¦¦¦-¦¦¦¦¦¦¦¦¦¦¦¦¦¦----------------¦¦¦¦¦¦¦¦¦¦¦¦¦¦-
'-¦¦¦¦¦¦----------¦¦¦¦¦¦-¦¦¦¦¦¦¦¦¦¦-¦¦¦¦¦¦¦¦¦¦¦¦¦¦-¦¦¦¦¦¦¦¦¦¦¦¦¦¦----------------¦¦¦¦¦¦¦¦¦¦¦¦¦¦-
'-¦¦¦¦¦¦----------¦¦¦¦¦¦-¦¦¦¦¦¦¦¦¦¦-¦¦¦¦¦¦¦¦¦¦¦¦¦¦-¦¦¦¦¦¦¦¦¦¦¦¦¦¦----------------¦¦¦¦¦¦¦¦¦¦¦¦¦¦-
'-----------------------------------------------------------------------------------------------
'MISS-O
' Williams, 1969
'
' -------------------------------------------------
'  based on _Gottlieb EM 4 player VPX table blank with options menu
' stripped down to single player
'   scripting/animation/menu/tweaking by BorgDog, 2016
'   originals images created by Popotte, vetorized/modified by HauntFreaks
'   layout/lighting Hauntfreaks
'   DOF scripting by Arngrim
' -------------------------------------------------
'
'
'
' Layer usage
'   1 - most stuff
'   2 - triggers
'   4 - options menu
'   6 - GI lighting
'   7 - plastics
'   8 - insert and bumper lighting
'
' Basic DOF config
'   101 Left Flipper, 102 Right Flipper,
'   103 Left sling, 104 Left sling flasher, 105 right sling, 106 right sling flasher,
'   107 Bumper1, 108 Bumper1 flasher, 109 bumper2, 110 bumper2 Flasher
'   111 Bumper3, 112 Bumper3 flasher, 113 Bumper4, 114 Bumper4 flasher, 134 Drain, 135 Ball Release
'   136 Shooter Lane/launch ball, 137 credit light, 138 knocker, 139 knocker Flasher
'   141 Chime1-10s, 142 Chime2-100s, 143 Chime3-1000s
'
'  -------------------------------------------------

Option Explicit
Randomize

Const cGameName = "miss_o_1969"

Dim operatormenu, options
Dim bumperlitscore, bumperoffscore
Dim replays
Dim Replay1Table(2), Replay2Table(2), Replay3Table(2), Replay4Table(2)
dim replay1, replay2, replay3, replay4
Dim Add10, Add100, Add1000
Dim hisc, credit
Dim score, state
Dim tilt, tiltsens
Dim balls, ballinplay
Dim matchnumb
dim ballrenabled
dim rstep, lstep
Dim rep, rst, eg
Dim dw1step, dw2step, dw3step, dw4step, dw5step, dw6step
Dim rw1step, rw2step
Dim advance, tempadvance, shootagain
Dim i,j, ii, objekt, light

On Error Resume Next
ExecuteGlobal GetTextFile("controller.vbs")
If Err Then MsgBox "Can't open controller.vbs"
On Error Goto 0

sub MissO_init
  LoadEM
  Replay1Table(1)=3300
  Replay2Table(1)=4500
  Replay3Table(1)=5600
  Replay4Table(1)=6700
  Replay1Table(2)=4800
  Replay2Table(2)=6000
  Replay3Table(2)=7100
  Replay4Table(2)=8200
  hideoptions
  For each light in GILights:light.State = 0: Next
  balls=5
  hisc=2000
  loadhs

  if HSA1="" then HSA1=4
  if HSA2="" then HSA2=15
  if HSA3="" then HSA3=7
  UpdatePostIt

  if not b2son then gamov.visible=1
  gamov.timerenabled=1
  tilttxt.timerenabled=1
  credittxt.setvalue(credit)
  if balls=3 then
    replays=1
    Else
    replays=2
  end If
  Replay1=Replay1Table(Replays)
  Replay2=Replay2Table(Replays)
  Replay3=Replay3Table(Replays)
  Replay4=Replay4Table(Replays)
  OptionBalls.image="OptionsBalls"&Balls
  bumperlitscore=10
  bumperoffscore=1
  advance=1
  moveadvance
  if balls=3 then
    InstCard.image="InstCard3balls"
    else
    InstCard.image="InstCard5balls"
  end if
  If B2SOn then
    setBackglass.enabled=true
    for each objekt in backdropstuff : objekt.visible = 0 : next
  End If
  if matchnumb="" then matchnumb=100
  if matchnumb=0 then
    matchtxt.text="00"
    else
    matchtxt.text=matchnumb
  end if
  scorereel.setvalue(score)
  PlaySound "motor"
  tilt=false
  If credit>0 then DOF 137, DOFOn
    Drain.CreateBall
End sub

sub setBackglass_timer
  Controller.B2ssetCredits Credit
  Controller.B2ssetMatch 34, Matchnumb
    Controller.B2SSetGameOver 35,1
    Controller.B2SSetScorePlayer 5, hisc
  me.enabled=0
end Sub


sub gamov_timer
  if state=false then
    If B2SOn then Controller.B2SSetGameOver 35,0
    gamov.visible=0
    gtimer.enabled=true
  end if
  gamov.timerenabled=0
end sub

sub gtimer_timer
  if state=false then
    if not b2son then gamov.visible=1
    If B2SOn then Controller.B2SSetGameOver 35,1
    gamov.timerenabled=1
    gamov.timerinterval= (INT (RND*10)+5)*100
  end if
  me.enabled=0
end sub

sub tilttxt_timer
  if state=false then
    tilttxt.visible=0
    If B2SOn then Controller.B2SSetTilt 33,0
    ttimer.enabled=true
  end if
  tilttxt.timerenabled=0
end sub

sub ttimer_timer
  if state=false then
    if not b2son then tilttxt.visible=1
    If B2SOn then Controller.B2SSetTilt 33,1
    tilttxt.timerenabled=1
    tilttxt.timerinterval= (INT (RND*10)+5)*100
  end if
  me.enabled=0
end sub

Sub MissO_KeyDown(ByVal keycode)

  if keycode=AddCreditKey then
    PlaySoundAtVol "coinin", Drain, 1
    coindelay.enabled=true
    end if

    if keycode=StartGameKey and credit>0 And Not HSEnterMode=true then
    if state=false then
    credit=credit-1
    if credit < 1 then DOF 137, DOFOff
    playsound "cluper"
    credittxt.setvalue(credit)
    ballinplay=1
    If B2SOn Then
      Controller.B2ssetCredits Credit
      Controller.B2sStartAnimation "Startup"
      Controller.B2ssetballinplay 32, Ballinplay
      Controller.B2ssetplayerup 30, 1
      Controller.B2SSetGameOver 0
    End If
      pup.state=1
    tilt=false
    state=true
    playsound "initialize"
    rst=0
    newgame.enabled=true
    end if
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
    If Options=3 then Options=1
    playsound "target"
    Select Case (Options)
      Case 1:
        Option1.visible=true
        Option2.visible=False
      Case 2:
        Option2.visible=true
        Option1.visible=False
    End Select
  end if

  If keycode=RightFlipperKey and State = false and OperatorMenu=1 then
    PlaySound "metalhit2"
    Select Case (Options)
    Case 1:
      if Balls=3 then
        Balls=5
        InstCard.image="InstCard5balls"
        else
        Balls=3
        InstCard.image="InstCard3balls"
      end if
      OptionBalls.image = "OptionsBalls"&Balls
    Case 2:
      OperatorMenu=0
      savehs
      HideOptions
    End Select
  End If

  If HSEnterMode Then HighScoreProcessKey(keycode)

  if tilt=false and state=true then
  If keycode = LeftFlipperKey Then
    LeftFlipper.RotateToEnd
    PlaySoundAtVol SoundFXDOF("flipperup",101,DOFOn,DOFContactors), LeftFlipper, 1
    PlayLoopSoundAtVol "Buzz", LeftFlipper, 1
  End If

  If keycode = RightFlipperKey Then
    RightFlipper.RotateToEnd
    PlaySoundAtVol SoundFXDOF("flipperup",102,DOFOn,DOFContactors), RightFlipper, 1
    PlayLoopSoundAtVol "Buzz1", RightFlipper, 1
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

Sub OperatorMenuTimer_Timer
  OperatorMenu=1
  Displayoptions
  Options=1
End Sub

Sub DisplayOptions
  OptionsBack.visible = true
  Option1.visible = True
  OptionBalls.visible = True
End Sub

Sub HideOptions
  for each objekt In OptionMenu
    objekt.visible = false
  next
End Sub


Sub MissO_KeyUp(ByVal keycode)

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
    PlaySoundAtVol SoundFXDOF("flipperdown",101,DOFOff,DOFContactors), LeftFlipper, 1
    StopSound "Buzz"
  End If

  If keycode = RightFlipperKey Then
    RightFlipper.RotateToStart
    PlaySoundAtVol SoundFXDOF("flipperdown",102,DOFOff,DOFContactors), RightFlipper, 1
    StopSound "Buzz1"
  End If
   End if
End Sub

sub flippertimer_timer()
  LFlip.RotY = LeftFlipper.CurrentAngle-90
  RFlip.RotY = RightFlipper.CurrentAngle+90
  Pgate.rotz = Gate.currentangle+25
end sub


Sub PairedlampTimer_timer
  HSR.state=HSL.state
  Ltarget.state=LkickL.state
end sub

sub coindelay_timer
  addcredit
    coindelay.enabled=false
end sub

Sub addcredit
      credit=credit+1
    DOF 137, DOFOn
      if credit>25 then credit=25
    credittxt.setvalue(credit)
    If B2SOn Then Controller.B2ssetCredits Credit
End sub

Sub Drain_Hit()
  DOF 134, DOFPulse
  PlaySoundAtVol "drain", Drain, 1
  me.timerenabled=1
End Sub

Sub Drain_timer
  if shootagain=true and tilt=false Then
    newball
    ballreltimer.enabled=True
    Else
    nextball
  end if
  me.timerenabled=0
End Sub

sub ballhome_hit
  ballrenabled=1
end sub

sub ballhome_unhit
  DOF 136, DOFPulse
end sub

sub ballrel_hit
  if ballrenabled=1 then
    ballrenabled=0
  end if
end sub

sub newgame_timer
    score=0
  rep=0
  advance=1
  lt1.state=1
  hsl.state=0
  for i=1 to 4: EVAL("Bumper"&i).hashitevent = 1: Next
  shootagain=False
  ScoreReel.resettozero
  pup.state=1
  If B2SOn then Controller.B2SSetScorePlayer 1, score
    eg=0
  For each light in GIlights:light.state=1:next
  for each light in bumperlights:light.state=0: next
  bumperlight2.state=1
    tilttxt.visible=0
  gamov.visible=0
    If B2SOn then
    Controller.B2SSetGameOver 35,0
    Controller.B2SSetTilt 33,0
    Controller.B2SSetMatch 34,0
  End If
    biptext.text="1"
  matchtxt.text=" "
  newball
  ballrelTimer.enabled=true
  newgame.enabled=false
end sub


sub newball
  if shootagain=true Then
    for each light in bumperlights:light.state=0:Next
    bumperlight2.state=1
    shootagain=False
  end if
End Sub


sub nextball
    if tilt=true then
    for i=1 to 4: EVAL("Bumper"&i).hashitevent = 1: Next
      tilt=false
      tilttxt.visible=0
    If B2SOn then
      Controller.B2SSetTilt 33,0
      Controller.B2ssetdata 1, 1
    End If
    end if
  ballinplay=ballinplay+1
  if ballinplay>balls then
    playsound "GameOver"
    eg=1
    ballreltimer.enabled=true
    else
    if state=true then
      newball
      ballreltimer.enabled=true
    end if
    biptext.text=ballinplay
    If B2SOn then Controller.B2ssetballinplay 32, Ballinplay
  end if
End Sub

sub ballreltimer_timer
  if eg=1 then
    turnoff
      matchnum
    state=false
    biptext.text=" "
    gamov.visible=1
    gamov.timerenabled=1
    tilttxt.timerenabled=1
    if score>hisc then
    hisc=score
    HighScoreEntryInit()
    UpdatePostIt
    end If
    pup.state=0

    savehs
    If B2SOn then
        Controller.B2SSetGameOver 35,1
        Controller.B2ssetballinplay 32, 0
      Controller.B2SSetScorePlayer 5, hisc
    End If
    For each light in GIlights:light.state=0:next
    ballreltimer.enabled=false
  else
  Drain.kick 60,35,0
    ballreltimer.enabled=false
  playsoundAtVol SoundFXDOF("kickerkick",135,DOFPulse,DOFContactors), Drain, 1
  end if
end sub

sub matchnum
  if matchnumb=0 then
    matchtxt.text=matchnumb
    If B2SOn then Controller.B2SSetMatch 100
    else
    matchtxt.text=matchnumb
    If B2SOn then Controller.B2SSetMatch 34,Matchnumb*10
  end if
  if (matchnumb)=(score mod 10) then
    addcredit
    playsound SoundFXDOF("knock",138,DOFPulse,DOFKnocker)
    DOF 139,DOFPulse
    end if
end sub

'********** Bumpers

Sub Bumper1_Hit
   if tilt=false then
  PlaySoundAtVol SoundFXDOF("fx_bumper4",107,DOFPulse,DOFContactors), ActiveBall, 1
  DOF 108,DOFPulse
  if bumperlight1.state=1 then
    addscore bumperlitscore
    else
    addscore bumperoffscore
  end if
  moveadvance
  me.timerenabled=1
   end if
End Sub

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


Sub Bumper2_Hit
   if tilt=false then
  PlaySoundAtVol SoundFXDOF("fx_bumper4",109,DOFPulse,DOFContactors), ActiveBall, 1
  DOF 110,DOFPulse
  if BumperLight2.state=1 then
    addscore bumperlitscore
    else
    addscore bumperoffscore
  end if
  moveadvance
  me.timerenabled=1
   end if
End Sub

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

Sub Bumper3_Hit
   if tilt=false then
  PlaySoundAtVol SoundFXDOF("fx_bumper4",111,DOFPulse,DOFContactors), ActiveBall, 1
  DOF 112,DOFPulse
  if BumperLight3.state=1 then
    addscore bumperlitscore
    else
    addscore bumperoffscore
  end if
  moveadvance
  me.timerenabled=1
   end if
End Sub

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

Sub Bumper4_Hit
   if tilt=false then
  PlaySoundAtVol SoundFXDOF("fx_bumper4",113,DOFPulse,DOFContactors), ActiveBall, 1
  DOF 114,DOFPulse
  if BumperLight4.state=1 then
    addscore bumperlitscore
    else
    addscore bumperoffscore
  end if
  moveadvance
  me.timerenabled=1
   end if
End Sub

Sub Bumper4_timer
  BumperTimerRing4.enabled=0
  if BumperRing4.transz>-36 then BumperRing4.transz=BumperRing4.transz-4
  if BumperRing4.transz=-36 then
    BumperTimerRing4.enabled=1
    me.timerenabled=0
  end if
End Sub

Sub BumperTimerRing4_timer
  if BumperRing4.transz<0 then BumperRing4.transz=BumperRing4.transz+4
  If BumperRing4.transz=0 then BumperTimerRing4.enabled=0
End sub

Sub RubberA_hit
  AddScore 1
  moveadvance
End Sub

Sub RubberB_hit
  AddScore 1
  moveadvance
End Sub

Sub RubberC_hit
  AddScore 1
  moveadvance
End Sub

Sub RubberD_hit
  AddScore 1
  moveadvance
End Sub

sub moveadvance
  if LkickL.state=1 Then
    LkickR.state=1
    LkickL.state=0
    if advance>15 then
      LspecialL.state=1
      LspecialR.state=0
    end If
    Else
    LkickL.state=1
    LkickR.state=0
    if advance>15 then
      LspecialR.state=1
      LspecialL.state=0
    end if
  end If
end sub

sub FlashBumpers
  FlashB.Play SeqAllOff
  FlashB.timerenabled=1
end sub

sub FlashB_timer
  FlashB.Play SeqAllOn
  me.timerenabled=0
end sub

'************** Slings

Sub RightSlingShot_Slingshot
  moveadvance
  PlaySoundAtVol SoundFXDOF("right_slingshot",105,DOFPulse,DOFContactors), ActiveBall, 1
  DOF 106,DOFPulse
  addscore 1
    RSling.Visible = 0
    RSling1.Visible = 1
  slingR.objroty = -15
    RStep = 1
    RightSlingShot.TimerEnabled = 1
End Sub

Sub RightSlingShot_Timer
    Select Case RStep
        Case 3:RSLing1.Visible = 0:RSLing2.Visible = 1:slingR.objroty = -7
        Case 4: slingR.objroty = 0:RSLing2.Visible = 0:RSLing.Visible = 1:RightSlingShot.TimerEnabled = 0
    End Select
    RStep = RStep + 1
End Sub

Sub LeftSlingShot_Slingshot
  moveadvance
  PlaySoundAtVol SoundFXDOF("left_slingshot",103,DOFPulse,DOFContactors), ActiveBall, 1
  DOF 104,DOFPulse
  addscore 1
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

'************** Dingwalls

sub dingwall1_hit
  addscore 1
  rdw1.visible=0
  RDW1a.visible=1
  dw1step=1
  Me.timerenabled=1
end sub

sub dingwall1_timer
  select case dw1step
    Case 1: RDW1a.visible=0: rdw1.visible=1
    case 2: rdw1.visible=0: rdw1b.visible=1
    Case 3: rdw1b.visible=0: rdw1.visible=1: Me.timerenabled=0
  end Select
  dw1step=dw1step+1
end sub

sub dingwall2_hit
  addscore 1
  rdw2.visible=0
  RDW2a.visible=1
  dw2step=1
  Me.timerenabled=1
end sub

sub dingwall2_timer
  select case dw2step
    Case 1: RDW2a.visible=0: rdw2.visible=1
    case 2: rdw2.visible=0: rdw2b.visible=1
    Case 3: rdw2b.visible=0: rdw2.visible=1: me.timerenabled=0
  end Select
  dw2step=dw2step+1
end sub

sub dingwall3_hit
  addscore 1
  Rdw3.visible=0
  RDW3a.visible=1
  dw3step=1
  Me.timerenabled=1
end sub

sub dingwall3_timer
  select case dw3step
    Case 1: RDW3a.visible=0: Rdw3.visible=1
    case 2: Rdw3.visible=0: rdw3b.visible=1
    Case 3: rdw3b.visible=0: Rdw3.visible=1: me.timerenabled=0
  end Select
  dw3step=dw3step+1
end sub

sub dingwall4_hit
  addscore 1
  Rdw4.visible=0
  RDW4a.visible=1
  dw4step=1
  Me.timerenabled=1
end sub

sub dingwall4_timer
  select case dw4step
    Case 1: RDW4a.visible=0: Rdw4.visible=1
    case 2: Rdw4.visible=0: rdw4b.visible=1
    Case 3: rdw4b.visible=0: Rdw4.visible=1: me.timerenabled=0
  end Select
  dw4step=dw4step+1
end sub

sub dingwall5_hit
  addscore 1
  Rdw5.visible=0
  RDW5a.visible=1
  dw5step=1
  Me.timerenabled=1
end sub

sub dingwall5_timer
  select case dw5step
    Case 1: RDW5a.visible=0: Rdw5.visible=1
    case 2: Rdw5.visible=0: rdw5b.visible=1
    Case 3: rdw5b.visible=0: Rdw5.visible=1: me.timerenabled=0
  end Select
  dw5step=dw5step+1
end sub

sub dingwall6_hit
  addscore 1
  Rdw6.visible=0
  RDW6a.visible=1
  dw6step=1
  Me.timerenabled=1
end sub

sub dingwall6_timer
  select case dw6step
    Case 1: RDW6a.visible=0: Rdw6.visible=1
    case 2: Rdw6.visible=0: rdw6b.visible=1
    Case 3: rdw6b.visible=0: Rdw6.visible=1: me.timerenabled=0
  end Select
  dw6step=dw6step+1
end sub

'********** Rubber non-scoring wall animations

sub Rwall1_hit
  Rw1.visible=0
  RW1a.visible=1
  rw1step=1
  Me.timerenabled=1
end sub

sub Rwall1_timer
  select case rw1step
    Case 1: RW1a.visible=0: Rw1.visible=1
    case 2: RW1.visible=0: rw1b.visible=1
    Case 3: rw1b.visible=0: Rw1.visible=1: me.timerenabled=0
  end Select
  rw1step=rw1step+1
end sub

sub Rwall2_hit
  Rw2.visible=0
  RW2a.visible=1
  rw2step=1
  Me.timerenabled=1
end sub

sub Rwall2_timer
  select case rw2step
    Case 1: RW2a.visible=0: Rw2.visible=1
    case 2: RW2.visible=0: rw2b.visible=1
    Case 3: rw2b.visible=0: Rw2.visible=1: me.timerenabled=0
  end Select
  rw2step=rw2step+1
end sub

'********** Triggers

sub TGtopL_hit
  DOF 115, DOFPulse
    if tilt=false then
    addscore 100
    end if
end sub

sub TGtopC_hit
  DOF 116, DOFPulse
    if tilt=false then
    addscore 200
    spotballs 2
    end if
end sub

sub TGtopR_hit
  DOF 117, DOFPulse
    if tilt=false then
    addscore 100
    end if
end sub

sub TGoutL_hit
   DOF 118, DOFPulse
   if tilt=False then
  if LspecialL.state=1 then
    addcredit
    PlaySound SoundFXDOF("knock",138,DOFPulse,DOFKnocker)
    DOF 139, DOFPulse
  end If
   end if
end Sub

sub TGoutR_hit
   DOF 119, DOFPulse
   if tilt=False then
  if LspecialR.state=1 then
    addcredit
    PlaySound SoundFXDOF("knock",138,DOFPulse,DOFKnocker)
    DOF 139, DOFPulse
  end If
   end if
end Sub

sub TGoutL1_hit
   DOF 120, DOFPulse
   if tilt=False then
    addscore 200
   end if
end Sub

sub TGoutR1_hit
   DOF 121, DOFPulse
   if tilt=False then
    addscore 200
   end if
end Sub

sub TGinL_hit
   DOF 122, DOFPulse
   if tilt=False then
    addscore 100
   end if
end Sub

sub TGinR_hit
   DOF 123, DOFPulse
   if tilt=False then
    addscore 100
   end if
end Sub

sub TGinC_hit
   DOF 124, DOFPulse
   if tilt=False then
    addscore 200
    spotballs 2
   end if
end Sub

sub TGhorseshoe_hit
   DOF 125, DOFPulse
   if tilt=False then
    addscore 500
    if HSL.state=1 then
      shootagain=True
      HSL.state=0
    end if
   end if
End Sub

sub spotballs(spot)
  tempadvance=advance
  advance=advance+spot
  lt1.timerenabled=1
  if advance>15 Then
    advance=16
    if lkickl.state=1 Then
      LspecialR.state=1
      Else
      LspecialL.state=1
    end If
  end If
end Sub

sub lt1_timer
  tempadvance=tempadvance+1
  for each light in advancelights:light.state=0:Next
  if tempadvance<=advance and tempadvance<16 Then EVAL("Lt"&tempadvance).state=1
' if tempadvance>1 and tempadvance<17 Then EVAL("Lt"&tempadvance-1).state=0
  if tempadvance=advance Then lt1.timerenabled=0
end Sub

'************* Targets

Sub TargetA_hit
  DOF 126, DOFPulse
  addscore 1
  bumperlighta.state=1
  bumperlight1.state=1
  bumperlight3.state=1
  checkeb
end Sub

Sub TargetB_hit
  DOF 127, DOFPulse
  addscore 1
  BumperLightB.state=1
  checkeb
end Sub

Sub TargetC_hit
  DOF 127, DOFPulse
  addscore 1
  BumperLightC.state=1
  checkeb
end Sub

Sub TargetD_hit
  DOF 128, DOFPulse
  addscore 1
  BumperLightD.state=1
  bumperlight4.state=1
  checkeb
end Sub

Sub Target_hit
  DOF 129, DOFPulse
  addscore 1
  if Ltarget.state=1 then spotballs 1
end sub

sub checkeb
  if (bumperlighta.state+bumperlightb.state+bumperlightc.state+bumperlightd.state=4) and shootagain=false then hsl.state=1
end Sub

'************* Kickers

sub KickerL_hit
  addscore 100
  if LkickL.state=1 Then
    select case (matchnumb)
      case 0,1,2:
        spotballs 1
      case 3,4,5:
        spotballs 2
      case 6,7,8,9:
        spotballs 3
      case 10:
        spotballs 5
    end Select
  end If
  me.timerenabled=1
end Sub

sub kickerl_timer
  PlaySoundAtVol SoundFXDOF("holekick",130,DOFPulse,DOFContactors), KickerL, 1
  DOF 139, DOFPulse
  KickerL.kick 120, 15, 0
  kickerl.timerenabled=0
end Sub

sub KickerR_hit
  addscore 100
  if LkickR.state=1 Then
    select case (matchnumb)
      case 0,1:
        spotballs 1
      case 2,3,4,5:
        spotballs 2
      case 6,7,8,9:
        spotballs 3
      case 10:
        spotballs 5
    end Select
  end If
  me.timerenabled=1
end Sub

sub kickerR_timer
  PlaySoundAtVol SoundFXDOF("holekick",131,DOFPulse,DOFContactors), KickerR, 1
  DOF 139, DOFPulse
  KickerR.kick 240, 15, 0
  kickerr.timerenabled=0
end Sub

'************* Scoring

sub addscore(points)
  if tilt=false then
  If points = 1 then
    matchnumb=matchnumb+1
    if matchnumb>9 then matchnumb=0
  end if
  if points=1 or points=10 or points=100 then
    addpoints Points
    else
    If Points < 10 and AddScore10Timer.enabled = false Then
      Add10 = Points
      AddScore10Timer.Enabled = TRUE
      ElseIf Points < 100 and AddScore100Timer.enabled = false Then
      Add100 = Points \ 10
      AddScore100Timer.Enabled = TRUE
      ElseIf AddScore1000Timer.enabled = false Then
      Add1000 = Points \ 100
      AddScore1000Timer.Enabled = TRUE
    End If
  End If
  end if
End Sub

Sub AddScore10Timer_Timer()
    if Add10 > 0 then
        AddPoints 1
        Add10 = Add10 - 1
    Else
        Me.Enabled = FALSE
    End If
End Sub

Sub AddScore100Timer_Timer()
    if Add100 > 0 then
        AddPoints 10
        Add100 = Add100 - 1
    Else
        Me.Enabled = FALSE
    End If
End Sub

Sub AddScore1000Timer_Timer()
    if Add1000 > 0 then
        AddPoints 100
        Add1000 = Add1000 - 1
    Else
        Me.Enabled = FALSE
    End If
End Sub

Sub AddPoints(Points)
    score=score+points
  ScoreReel.addvalue(points)
  If B2SOn Then Controller.B2SSetScorePlayer1 score

    ' Sounds: there are 3 sounds: tens, hundreds and thousands
    If Points = 10 AND(Score MOD 100) \ 10 = 0 Then  'New 100 reel
        PlaySound SoundFXDOF("bell1000",143,DOFPulse,DOFChimes)
      ElseIf Points = 1 AND(Score MOD 10) = 0 Then 'New 10 reel
        PlaySound SoundFXDOF("bell100",142,DOFPulse,DOFChimes)
      ElseIf points = 100 Then
        PlaySound SoundFXDOF("bell1000",143,DOFPulse,DOFChimes)
    elseif Points = 10 Then
        PlaySound SoundFXDOF("bell100",142,DOFPulse,DOFChimes)
      Else
        PlaySound SoundFXDOF("bell10",141,DOFPulse,DOFChimes)
    End If
    ' check replays
    if score=>replay1 and rep=0 then
    addcredit
    rep=1
    PlaySound SoundFXDOF("knock",138,DOFPulse,DOFKnocker)
    DOF 139, DOFPulse
    end if
    if score=>replay2 and rep=1 then
    addcredit
    rep=2
    PlaySound SoundFXDOF("knock",138,DOFPulse,DOFKnocker)
    DOF 139, DOFPulse
    end if
    if score=>replay3 and rep=2 then
    addcredit
    rep=3
    PlaySound SoundFXDOF("knock",138,DOFPulse,DOFKnocker)
    DOF 139, DOFPulse
    end if
end sub


Sub CheckTilt
  If Tilttimer.Enabled = True Then
   TiltSens = TiltSens + 1
   if TiltSens = 3 Then
     Tilt = True
     eg=1
     tilttxt.visible=1
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

Sub Tilttimer_Timer()
  Tilttimer.Enabled = False
End Sub

sub turnoff
  for i=1 to 4
    EVAL("Bumper"&i).hashitevent = 0
  Next
    LeftFlipper.RotateToStart
  StopSound "Buzz"
  DOF 101, DOFOff
  RightFlipper.RotateToStart
  StopSound "Buzz1"
  DOF 102, DOFOff
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
        If BallVel(BOT(b) ) > 1 AND BOT(b).z < 30 Then
            rolling(b) = True
            PlaySound("fx_ballrolling" & b), -1, Vol(BOT(b))*4, Pan(BOT(b) ), 0, Pitch(BOT(b) ), 1, 0
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
  PlaySound("fx_collide"), 0, Csng(velocity) ^2 / 2000, Pan(ball1), 0, Pitch(ball1), 0, 0
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
  PlaySound "pinhit_low", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0
End Sub

Sub a_Targets_Hit (idx)
  PlaySound "target", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0
End Sub

Sub a_Metals_Thin_Hit (idx)
  PlaySound "metalhit_thin", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0
End Sub

Sub a_Metals_Medium_Hit (idx)
  PlaySound "metalhit_medium", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0
End Sub

Sub a_Metals2_Hit (idx)
  PlaySound "metalhit2", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0
End Sub

Sub a_Gates_Hit (idx)
  PlaySound "gate4", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0
End Sub

Sub a_Spinner_Spin
  PlaySound "fx_spinner",0,.25,0,0.25
End Sub

Sub a_Rubbers_Hit(idx)
  dim finalspeed
    finalspeed=SQR(activeball.velx * activeball.velx + activeball.vely * activeball.vely)
  If finalspeed > 20 then
    PlaySound "fx_rubber2", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0
  End if
  If finalspeed >= 6 AND finalspeed <= 20 then
    RandomSoundRubber()
  End If
End Sub

Sub RubberWheel_hit
  dim finalspeed
    finalspeed=SQR(activeball.velx * activeball.velx + activeball.vely * activeball.vely)
  If finalspeed > 20 then
    PlaySound "fx_rubber2", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0
  End if
  If finalspeed >= 6 AND finalspeed <= 20 then
    RandomSoundRubber()
  End If
End sub

Sub a_Posts_Hit(idx)
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

Sub RandomSoundFlipper()
  Select Case Int(Rnd*3)+1
    Case 1 : PlaySound "flip_hit_1", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0
    Case 2 : PlaySound "flip_hit_2", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0
    Case 3 : PlaySound "flip_hit_3", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0
  End Select
End Sub


sub savehs

    savevalue "MissO", "credit", credit
    savevalue "MissO", "hiscore", hisc
    savevalue "MissO", "match", matchnumb
    savevalue "MissO", "score", score
  savevalue "MissO", "balls", balls
  savevalue "MissO", "hsa1", HSA1
  savevalue "MissO", "hsa2", HSA2
  savevalue "MissO", "hsa3", HSA3
end sub

sub loadhs
    dim temp
  temp = LoadValue("MissO", "credit")
    If (temp <> "") then credit = CDbl(temp)
    temp = LoadValue("MissO", "hiscore")
    If (temp <> "") then hisc = CDbl(temp)
    temp = LoadValue("MissO", "match")
    If (temp <> "") then matchnumb = CDbl(temp)
    temp = LoadValue("MissO", "score")
    If (temp <> "") then score = CDbl(temp)
    temp = LoadValue("MissO", "balls")
    If (temp <> "") then balls = CDbl(temp)
    temp = LoadValue("MissO", "hsa1")
    If (temp <> "") then HSA1 = CDbl(temp)
    temp = LoadValue("MissO", "hsa2")
    If (temp <> "") then HSA2 = CDbl(temp)
    temp = LoadValue("MissO", "hsa3")
    If (temp <> "") then HSA3 = CDbl(temp)
end sub

Sub MissO_Exit()
  Savehs
  turnoff
  If B2SOn Then Controller.stop
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

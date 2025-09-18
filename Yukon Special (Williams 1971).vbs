'*** Jackpot ***'

Option Explicit  'Force explicit variable declaration.
Randomize



Const cGameName = "YukonSpecial_1971"
Const B2STableName = "YukonSpecial_1971"
ExecuteGlobal GetTextFile("core.vbs")

On Error Resume Next
ExecuteGlobal GetTextFile("Controller.vbs")
If Err Then MsgBox "Unable to open Controller.vbs. Ensure that it is in the scripts folder."
On Error Goto 0
Dim Controller  ' B2S
Dim B2SScore  ' B2S Score Displayed
dim score(4)    'Reel score
dim truesc(4)   'True reel score if overring the numer og digits
dim reel(8)     'Maximum number of scoring reels
dim state
dim credit      'Credits
dim eg        'End Game
dim currpl      'Current player status
dim playno      'Number of players
dim up(4)     'Player up
dim play(4)     'Player status
dim rst       'Reset Variable
dim bip       'Ball in play
dim ballrelenabled  'Used for shooter sound
dim match(10)
dim tilt
dim tiltsens
dim rep(4)
dim matchnumb
dim bell      'Global scoring sound variable
dim points
dim replay1
dim replay2
dim replay3
dim replay4
dim hisc      'High Score to date
dim wv        'Wheel Value (positiion)
dim wn        'Wheel Change animation steps
dim rv(8)
dim i       'Indexing variable
dim scn       'Single points score timer
dim scn1      'Single points end score timer
dim scn2      'Multiple points score timer
dim scn3      'Multiple points end score timer
dim obj
Dim B2SOn



Sub Table1_Init
  If Table1.ShowDT = false then
    For each obj in DesktopStuff
      obj.visible=False
    next
  End If
  LoadEM
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
  set reel(1)=reel1
  set reel(2)=reel2
  set reel(3)=reel3
  set reel(4)=reel4
  set reel(5)=reel5
  set reel(6)=reel6
  set reel(7)=reel7
  set reel(8)=reel8
  replay1=40000
  replay2=80000
  replay3=9980000
  replay4=9990000
  loadhs
  if hisc="" then hisc=10000
  reel5.setvalue(hisc)
  reel6.setvalue(rv(6))
  reel7.setvalue(rv(7))
  reel8.setvalue(rv(8))
  Ramp1.image = "GoldRushReel_1_"&(rv(6)+1)
  Ramp2.image = "GoldRushReel_1_"&(rv(7)+1)
  Ramp3.image = "GoldRushReel_1_"&(rv(8)+1)
  checkreels
    if credit="" then credit=0
    credittxt.text=credit

    for i=1 to 4
    currpl=i
    reel(i).setvalue(score(i))
    next
    currpl=0
  If B2SOn then
    Controller.B2SSetScorePlayer1 hisc
    Controller.B2SSetGameOver 1
    Controller.B2SSetTilt 1
    Controller.B2SSetCredits credit
    If matchnumb =100 then
      Controller.B2SSetMatch 100
    else
      Controller.B2SSetMatch matchnumb*10
    end if
  end if
End Sub


Sub Table1_KeyDown(ByVal keycode)

  If keycode = PlungerKey Then
  Plunger.PullBack
  End If

    if keycode = 6 then
  PlaySoundAt "coin3", Drain, 1
  coindelay.enabled=true
  end if

    if keycode = 5 then
  PlaySoundAt "coin3", Drain, 1
  coindelay1.enabled=true
  end if

  if keycode = 2 and credit>0 and state=false and playno=0 then
  credit=credit-1
  credittxt.text=credit
  If B2SOn then
    Controller.B2SSetCredits credit
    Controller.B2SSetCanPlay 1
    Controller.B2SSetScoreRolloverPlayer1 0
  end if
    eg=0
    playno=1
    playno1.state=1
    currpl=1
    play(currpl).state=1
    playsound "click"
    playsound "initialize"
    rst=0
    bip=5
    resettimer.enabled=true
    end if



    if state=true and tilt=false then

  If keycode = LeftFlipperKey Then
    LeftFlipper.RotateToEnd
    PlaySoundAt "FlipperUp", LeftFlipper, 1
  End If

  If keycode = RightFlipperKey Then
    RightFlipper.RotateToEnd
        PlaySoundAt "FlipperUp", RightFlipper, 1
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
  End If

  If keycode = LeftFlipperKey Then
    LeftFlipper.RotateToStart
    stopsound "buzz"
    if state=true and tilt=false then PlaySoundAt "FlipperDown", LeftFlipper, 1
  End If

  If keycode = RightFlipperKey Then
    RightFlipper.RotateToStart
    stopsound "buzz"
        if state=true and tilt=false then PlaySoundAt "FlipperDown", RightFlipper, 1
  End If

End Sub

'***********************
'     Flipper Logos
'***********************

Sub UpdateFlipperLogos_Timer
  LFlip.ObjRotZ = LeftFlipper.CurrentAngle-90
  RFlip.ObjRotZ = RightFlipper.CurrentAngle+90


End Sub


sub newgame
  state=true
  eg=0
  for i=0 to 3
  score(i)=0
  truesc(i)=0
  rep(i)=0
  next
  bip5.text="5"
  bip1.text=" "
  for i=0 to 9
  match(i).text=" "
  next
  tilttext.text=" "
  gamov.text=" "
  tilt=false
  tiltsens=0
  bip=5
  checkreels
  If B2SOn then
    Controller.B2SSetTilt 0
    Controller.B2SSetGameOver 0
    Controller.B2SSetMatch 0
    Controller.B2SSetBallInPlay bip
    Controller.B2SSetPlayerUp 1
  end if
  bumperlight1.state=0
  bumperlight2.state=0
  bumperlight3.state=0
  bumperlight4.state=0
  TriggerLight1.state=1
  TriggerLight2.state=1
  TriggerLight3.state=1
  TriggerLight7.state=1
  TriggerLight8.state=1
  TriggerLight12.state=1
  TriggerLight13.state=1
  TargetLight1.state=1
  Targetlight2.state=1
  TargetLight3.state=1
  TargetLight4.state=1
  TargetLight5.state=1
  TargetLightSpecial.state=0
  leftkickerlight.state=0
  rightkickerlight.state=0


  nb.CreatesizedBall 27
  nb.kick 90,6
end sub

sub resettimer_timer
  rst=rst+1
  reel1.resettozero
  reel2.resettozero
  reel3.resettozero
  reel4.resettozero
  reel6.resettozero
  reel7.resettozero
  reel8.resettozero
  if rv(6)>0 then
    rv(6)=rv(6)-1
    Ramp1.image = "GoldRushReel_1_"&(rv(6)+1)
  end if
  if rv(7)>0 then
    rv(7)=rv(7)-1
    Ramp2.image = "GoldRushReel_1_"&(rv(7)+1)
  end if
  if rv(8)>0 then
    rv(8)=rv(8)-1
    Ramp3.image = "GoldRushReel_1_"&(rv(8)+1)
  end if

  If B2SOn then
    Controller.B2SSetScorePlayer1 0
    Controller.B2SSetScorePlayer2 0
    Controller.B2SSetScorePlayer3 0
    Controller.B2SSetScorePlayer4 0
  end if
  if rst=14 then
  playsound "newball"
  end if
  if rst=18 then
  newgame
  resettimer.enabled=false
    end if
end sub

sub coindelay_timer
  playsound "click"
  credit=credit+5
  credittxt.text=credit
  If B2SOn then
    Controller.B2SSetCredits credit
  end if
    coindelay.enabled=false
end sub

sub coindelay1_timer
  playsound "click"
  credit=credit+1
  credittxt.text=credit
  If B2SOn then
    Controller.B2SSetCredits credit
  end if
    coindelay1.enabled=false
end sub


'*** General points scoring ***

sub addscore(points)

    if tilt=false then
    bell=0
    if points=100 then matchnumb=matchnumb+1
    if matchnumb=10 then matchnumb=0

    if points = 10 or points = 100 or points = 1000 or points = 10000 then scn=1
    if points = 5000 then scn2=5
    if points = 4000 then scn2=4
    if points = 3000 then scn2=3
    if points = 500 then scn2=5

    if points = 5000 then
      reel(currpl).addvalue(5000)
      'playsound "motorshort1s"
      scn3=5
      bell=1000
    end if

    if points = 4000 then
      reel(currpl).addvalue(4000)
      'playsound "motorshort1s"
      scn3=4
      bell=1000
    end if

    if points = 3000 then
      reel(currpl).addvalue(3000)
      'playsound "motorshort1s"
      scn3=3
      bell=1000
    end if


    if points = 500 then
      reel(currpl).addvalue(500)
      'playsound "motorshort1s"
      scn3=5
      bell=100
    end if

    if points = 10000 then
      reel(currpl).addvalue(10000)
      bell=1000
      scn1=1
    end if

    if points = 1000 then
      reel(currpl).addvalue(1000)
      bell=1000
      scn1=1
    end if

    if points = 100 then
      reel(currpl).addvalue(100)
      bell=100
      scn1=1
    end if

    if points = 10 then
      reel(currpl).addvalue(10)
      bell=10
      scn1=1
    end if

    if points = 10 or points = 100 or points = 1000 or points = 10000 then scn1=0 : scntimer.enabled=true : end if
    if points = 5000 then scn3=0 : scntimer1.enabled=true : end if
    if points = 4000 then scn3=0 : scntimer1.enabled=true : end if
    if points = 3000 then scn3=0 : scntimer1.enabled=true : end if
    if points = 500 then scn3=0 : scntimer1.enabled=true : end if


    score(currpl)=score(currpl)+points
    If B2SOn then
      Controller.B2SSetScorePlayer currpl,score(currpl)
    end if
    truesc(currpl)=truesc(currpl)+points
  end if

    if score(currpl)>99990 then
    If B2SOn then
      Controller.B2SSetScoreRolloverPlayer1 (int(truesc(currpl)/100000) )
    end if
    score(currpl)=score(currpl)-100000
    rep(currpl)=0
    end if

    if truesc(currpl)=>replay1 and rep(currpl)=0 then

    playsound "knocke"
    bip=bip+1
    if bip>9 then
      bip=9
    end if
    If B2SOn then
      Controller.B2SSetBallInPlay bip
    End If
    rep(currpl)=1
    playsound "click"
    end if

    if truesc(currpl)=>replay2 and rep(currpl)=1 then
    playsound "knocke"
    bip=bip+1
    if bip>9 then
      bip=9
    end if
    If B2SOn then
      Controller.B2SSetBallInPlay bip
    End If
    rep(currpl)=2
    playsound "click"
    end if



end sub

sub scntimer_timer
    scn1=scn1 + 1
    if bell=1000 then playsound "bell1000",0, 0.45, 0, 0
    if bell=100 then playsound "bell100",0, 0.3, 0, 0
    if bell=10 then playsound "bell10",0, 0.15, 0, 0
    if scn1=scn then scntimer.enabled=false
end sub

sub scntimer1_timer
  scn3=scn3 + 1
    if bell=1000 then playsound "bell1000",0, 0.45, 0, 0
    if bell=100 then playsound "bell100",0, 0.3, 0, 0
    if bell=10 then playsound "bell10",0, 0.15, 0, 0
    if scn3=scn2 then scntimer1.enabled=false
end sub


'*** Ball handling ***

sub ballhome_hit
    ballrelenabled=1
end sub

sub ballrelease_hit
  if ballrelenabled=1 then PlaySoundAtVol  "launchball", ActiveBall, 1: ballrelenabled=0: end if
end sub

Sub Drain_Hit()
  PlaySoundAt "drainshorter", Drain, 1
  Drain.DestroyBall
  nextball
End Sub

sub nextball
  bumperlight1.state=0
  bumperlight2.state=0
  bumperlight3.state=0
  bumperlight4.state=0

  if tilt=true then
  tilt=false
  tilttext.text=" "
  tiltseq.stopplay
  end if
  if (Light6.state)=1 then
  playsound "newball"
  ballreltimer.enabled=true
  light6.state=0

  else
  currpl=currpl+1
  end if
  if currpl>playno then
  bip=bip-1
  if bip<1 then
  playsound "motorleer"
  eg=1
  ballreltimer.enabled=true
  else
  if state=true and tilt=false then
  play(currpl-1).state=0
  currpl=1
  play(currpl).state=1
  playsound "newball"
  ballreltimer.enabled=true
  end if
  select case (bip)
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
  playsound "newball"
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
  for i=1 to 4
  if truesc(i)>hisc then
  hisc=truesc(i)
  reel5.setvalue(hisc)
  end if
  next
  play(currpl-1).state=0
  playno=0
  gamov.text="GAME OVER"
  If B2SOn then
    Controller.B2SSetTilt 0
    Controller.B2SSetGameOver 1
    Controller.B2SSetBallInPlay 0
    Controller.B2SSetPlayerUp 0
  end if
  savehs
  ballreltimer.enabled=false
  else
  If B2SOn then
    Controller.B2SSetTilt 0
    Controller.B2SSetGameOver 0
    Controller.B2SSetShootAgain 0
    Controller.B2SSetBallInPlay bip
    Controller.B2SSetPlayerUp currpl
  end if
  nb.CreatesizedBall 27
  nb.kick 90,6
    ballreltimer.enabled=false
    end if
end sub

sub matchnum
  exit sub
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
  if B2SOn Then
    if matchnumb=0 then
      Controller.B2SSetMatch 100
    else
      Controller.B2SSetMatch matchnumb*10
    end if
  end if
    for i=1 to playno
    if matchnumb*10=(score(i) mod 100) then
    credit=credit+1
    playsound "knocke"
    credittxt.text= credit
  If B2SOn then
    Controller.B2SSetCredits credit
  end if
    end if
    next
end sub

Sub CheckTilt
  If Tilttimer.Enabled = True Then
  TiltSens = TiltSens + 1
  if TiltSens = 3 Then
  Tilt = True
  tilttext.text="TILT"
  If B2SOn then
    Controller.B2SSetTilt 1
  end if

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
end sub

Sub Gate1_hit()
  PlaySoundAt "gate", ActiveBall, 1
End Sub

Sub AddSpecial
  playsound "knocke"
  bip=bip+1
  if bip>9 then
    bip=9
  end if
  If B2SOn then
    Controller.B2SSetBallInPlay bip
  End If
  TargetLightSpecial.state=0
  leftkickerlight.state=0
  rightkickerlight.state=0
  reel7add
  TargetLight2.state=1
  TriggerLight2.state=1
  TriggerLight7.state=1
  TriggerLight8.state=1
  TriggerLight12.state=1
  TriggerLight13.state=1

end sub

'**********Sling Shot Animations
' Rstep and Lstep  are the variables that increment the animation
'****************
Dim RStep, Lstep

Sub RightSlingShot_Slingshot
  if tilt = false then
    PlaySoundAtVol "right_slingshot", ActiveBall, 1
    RSling.Visible = 0
    RSling1.Visible = 1
    sling1.TransZ = -20
    RStep = 0
    RightSlingShot.TimerEnabled = 1
    addscore 10
  end if
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
  PlaySoundAtVol "left_slingshot",ActiveBall, 1
  LSling.Visible = 0
  LSling1.Visible = 1
  sling2.TransZ = -20
  LStep = 0
  LeftSlingShot.TimerEnabled = 1
  addscore 10
  end if
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
' Thalamus, AudioFade - Patched
  If tmp > 0 Then
    AudioFade = Csng(tmp ^5) 'was 10
  Else
    AudioFade = Csng(-((- tmp) ^5) ) 'was 10
  End If
End Function

Function AudioPan(tableobj) ' Calculates the pan for a tableobj based on the X position on the table. "table1" is the name of the table
  Dim tmp
  On Error Resume Next
  tmp = tableobj.x * 2 / table1.width-1
' Thalamus, AudioPan - Patched
  If tmp > 0 Then
    AudioPan = Csng(tmp ^5) 'was 10
  Else
    AudioPan = Csng(-((- tmp) ^5) ) 'was 10
  End If
End Function

Function Pan(ball) ' Calculates the pan for a ball based on the X position on the table. "table1" is the name of the table
  Dim tmp
  On Error Resume Next
  tmp = ball.x * 2 / table1.width-1
' Thalamus, Pan - Patched
  If tmp > 0 Then
    Pan = Csng(tmp ^5) 'was 10
  Else
    Pan = Csng(-((- tmp) ^5) ) ' was 10
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
  PlaySoundAtVol "fx_spinner", a_Spinner, 1
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


'*** High Score Handling ***

sub loadhs
    ' Based on Black's Highscore routines
  dim FileObj
  dim ScoreFile
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
  dim temp10
  Set FileObj=CreateObject("Scripting.FileSystemObject")
  If Not FileObj.FolderExists(UserDirectory) then
  Exit Sub
  End if
  If Not FileObj.FileExists(UserDirectory & "yukonspecial.txt") then
  Exit Sub
  End if
  Set ScoreFile=FileObj.GetFile(UserDirectory & "yukonspecial.txt")
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
  temp10=Textstr.ReadLine
  TextStr.Close
  credit = CDbl(temp1)
  score(0) = CDbl(temp2)
  score(1) = CDbl(temp3)
  score(2) = CDbl(temp4)
  score(3) = CDbl(temp5)
  hisc = CDbl(temp6)
  matchnumb = CDbl(temp7)
  rv(6) = CDbl(temp8)
  rv(7) = CDbl(temp9)
  rv(8) = CDbl(temp10)
  Set ScoreFile=Nothing
  Set FileObj=Nothing
end sub

sub savehs
  ' Based on Black's Highscore routines
  Dim FileObj
  Dim ScoreFile
  Set FileObj=CreateObject("Scripting.FileSystemObject")
  If Not FileObj.FolderExists(UserDirectory) then
  Exit Sub
  End if
  Set ScoreFile=FileObj.CreateTextFile(UserDirectory & "Jackpot.txt",True)
  scorefile.writeline credit
  ScoreFile.WriteLine score(0)
  ScoreFile.WriteLine score(1)
  ScoreFile.WriteLine score(2)
  ScoreFile.WriteLine score(3)
  scorefile.writeline hisc
  scorefile.writeline matchnumb
  scorefile.writeline rv(6)
  scorefile.writeline rv(7)
  scorefile.writeline rv(8)
  ScoreFile.Close
  Set ScoreFile=Nothing
  Set FileObj=Nothing
end sub


'*** Tabel specific score handling ***


sub lk_hit
  if tilt=false then
    if leftkickerlight.state=1 then
      AddSpecial
    else
      AddScore 5000
    end if

    end if
    PlaySoundAt "kickerkick", ActiveBall, 1
    kicktimer.enabled=true
 end sub

sub rk_hit
  if tilt=false then
    if rightkickerlight.state=1 then
      AddSpecial
    else
      AddScore 5000
    end if
  end if
  PlaySoundAt "kickerkick", ActiveBall, 1
  kicktimer.enabled=true
end sub


sub kicktimer_timer

  lk.kick (int(rnd(1)*5)+165),15
  rk.kick (int(rnd(1)*5)+190),15
    kicktimer.enabled=false
end sub



Sub Bumper1_Hit()
  If Bumperlight1.state=1 then addscore 100 else addscore 10
  PlaySoundAt "jet1", ActiveBall, 1
End Sub

Sub Bumper2_Hit()
  If Bumperlight2.state=1 then addscore 100 else addscore 10
  PlaySoundAt "jet1", ActiveBall, 1
End Sub

Sub Bumper3_Hit()
  If Bumperlight3.state=1 then addscore 100 else addscore 10
  PlaySoundAt "jet1", ActiveBall, 1
End Sub

Sub Bumper4_Hit()
  If Bumperlight4.state=1 then addscore 100 else addscore 10
  PlaySoundAt "jet1", ActiveBall, 1
End Sub


Sub Trigger1_Hit()
  addscore 1000
  if TriggerLight1.state=1 then
    reel6add
  end if
End Sub

Sub Trigger2_Hit()
  addscore 1000
  if TriggerLight2.state=1 then
    reel7add
  end if
End Sub


Sub Trigger3_Hit()
  addscore 1000
  if TriggerLight3.state=1 then
    reel8add
  end if
End Sub


Sub Trigger5_Hit()
  AddScore 10000
End Sub

Sub Trigger6_Hit()
  AddScore 10000
End Sub

Sub Trigger7_Hit()
  Addscore 1000
  if tilt=false then
    BumperLight1.state=1
    BumperLight2.state=1
    BumperLight3.state=1
    BumperLight4.state=1
    if TriggerLight7.state=1 then
      reel7add
    end if
  end if
End Sub

Sub Trigger8_Hit()
  Addscore 1000
  if tilt=false then
    BumperLight1.state=1
    BumperLight2.state=1
    BumperLight3.state=1
    BumperLight4.state=1
    if TriggerLight8.state=1 then
      reel7add
    end if
  end if
End Sub

Sub Trigger9_Hit()
  AddScore 10000
End Sub

Sub Trigger10_Hit()
  AddScore 10000
End Sub

Sub Trigger12_Hit()
  Addscore 1000
  if tilt=false then
    if TriggerLight12.state=1 then
      reel7add
    end if
  end if
End Sub

Sub Trigger13_Hit()
  Addscore 1000
  if tilt=false then
    if TriggerLight13.state=1 then
      reel7add
    end if
  end if
End Sub

Sub Target1_Hit()
  if TargetLight1.state=1 then
    reel6add
  end if
  Addscore 1000
End Sub

Sub Target2_Hit()
  if TargetLight2.state=1 then
    reel7add
  else
    if TargetLightSpecial.state=1 then
      AddSpecial
    end if
  end if
  Addscore 1000
End Sub

Sub Target3_Hit()
  if TargetLight3.state=1 then
    reel8add
  end if
  Addscore 1000
End Sub

Sub Target4_Hit()
  if TargetLight4.state=1 then
    reel6add
  end if
  Addscore 1000
End Sub

Sub Target5_Hit()
  if TargetLight5.state=1 then
    reel8add
  end if
  Addscore 1000
End Sub


sub reel6add

    reel6.addvalue(1)
    rv(6)=rv(6)+1
    if rv(6)=9 then
    TargetLight1.state=0
    TargetLight4.state=0
    TriggerLight1.state=0
  end if
  Ramp1.image = "GoldRushReel_1_"&(rv(6)+1)
    checkreels
end sub

sub reel7add

    reel7.addvalue(1)
    rv(7)=rv(7)+1
  if rv(7)>9 then rv(7)=0
    if rv(7)=9 then
    TargetLight2.state=0
    TriggerLight7.state=0
    TriggerLight8.state=0
    TriggerLight12.state=0
    TriggerLight13.state=0
    TriggerLight2.state=0
  end if
  Ramp2.image = "GoldRushReel_1_"&(rv(7)+1)
    checkreels
end sub

sub reel8add

    reel8.addvalue(1)
    rv(8)=rv(8)+1
    if rv(8)=9 then
    TargetLight3.state=0
    TargetLight5.state=0
    TriggerLight3.state=0
  end if
  Ramp3.image = "GoldRushReel_1_"&(rv(8)+1)
    checkreels
end sub


sub checkreels
  if rv(6)=9 AND rv(7)=9 AND rv(8)=9 then
    LeftKickerLight.state=1
    RightKickerLight.state=1
    TargetLightSpecial.state=1
  end if
end sub

Sub Table1_exit()
  savehs
  If B2SOn Then Controller.Stop
end sub

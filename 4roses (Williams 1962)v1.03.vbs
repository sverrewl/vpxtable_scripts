'-----------------------------------------------
'                  4 ROSES WILLIAMS 1962
'                      BALATER .Version VPX
'                        DOF Config Foxyt
'            Script Adj - abhcoide
'-----------------------------------------------

Randomize

On Error Resume Next
ExecuteGlobal GetTextFile("controller.vbs")
If Err Then MsgBox "You need the controller.vbs in order to run this table, available in the vp10 package"
On Error Goto 0

Const cGameName = "roses_1962"
Const B2STableName="4roses_1962"
Dim Score
Dim BallstoPlay
Dim Credits
Dim GameinProgress
Dim GameStart
Dim BallsPerGame
Dim Highscore
Dim Tilt
Dim TiltSensor
Dim LPR, LPY,KickerCounterR, KickerCounterY
Dim Replay
Dim BallinLane
Dim Playedballs
Dim aXpos
Dim aXadd
Dim aYpos
Dim aYadd
Dim aZpos
dim aBall
'-----------------------------------------------------
Dim Matchnum

 ExecuteGlobal GetTextFile("core.vbs")

Dim DesktopMode: DesktopMode = Table1.ShowDT
If DesktopMode = True Then
  EMReel1.Visible=True
  EMReel0.Visible=True
  match0.visible=true
  match1.visible=true
  match2.visible=true
  match3.visible=true
  match4.visible=true
  match5.visible=true
  match6.visible=true
  match7.visible=true
  match8.visible=true
  match9.visible=true
  GameOver.visible=True
  TiltLight.visible=True
  credittext.visible=True
  GameOver.SetValue 1
Else
  EMReel1.Visible=False
  EMReel0.Visible=False
  match0.visible=false
  match1.visible=false
  match2.visible=false
  match3.visible=false
  match4.visible=false
  match5.visible=false
  match6.visible=false
  match7.visible=false
  match8.visible=false
  match9.visible=false
  GameOver.visible=false
  TiltLight.visible=false
  Credittext.visible=false
end If
  kicker2.CreateSizedBall(22)
  kicker2.Kick 180,1
  kicker003.CreateSizedBall(22)
  Kicker003.Kick 180,1
  kicker003.enabled=false
  Kicker004.CreateSizedBall(22)
  kicker004.Kick 180,1
  kicker004.enabled=false
  Kicker005.CreateSizedBall(22)
  kicker005.Kick 180,1
  kicker005.enabled=false
  kicker006.CreateSizedBall(22)
  kicker006.Kick 180,1
  kicker006.enabled=false
Sub Table1_Init()
LoadEM
loadhs
UpdatePostIt
ballinlane=1
If B2SOn then
  Controller.B2SSetGameOver 1
  Controller.B2SSetScoreRolloverPlayer1 0
  Controller.B2SSetScorePlayer1 Score
  Controller.B2SSetTilt 0
  Controller.B2SSetMatch Credits+1
  Controller.B2SSetData 61, 0
  Controller.B2SSetData 62, 0
  Controller.B2SSetData 63, 0
  Controller.B2SSetData 64, 0
  Controller.B2SSetData 65, 0
  Controller.B2SSetData 66, 0
  Controller.B2SSetData 67, 0
  Controller.B2SSetData 68, 0
  Controller.B2SSetData 69, 0
  Controller.B2SSetData 70, 0
  Controller.B2SSetData 71, 0
  Controller.B2SSetCredits Credits
End If
motor.enabled=false
Replay=0
TiltSensor=0
Matchnum=0
GameinProgress=0
GameStart=0
Score=0
BallsPerGame=5
BallstoPlay=0
'Credittext.Text=Credits
If B2SOn Then
    Controller.B2SSetCredits Credits
End If
if Credits > 0 Then DOF 121, DOFOn
End Sub

Sub EndofGame()
  Dim numer0
    If B2SOn Then
        Controller.B2SSetCredits Credits
    End If
  If B2SOn then Controller.B2SSetMatch Credits+1
  GameinProgress=0

  DOF 101, DOFOff
  DOF 102, DOFOff
  LeftFlipper.RotateToStart
  RightFlipper.RotateToStart
  stopSound "buzzl"
  stopSound "buzz"

  select Case Matchnum
    case 1: Match1.SetValue 1:If B2SOn then Controller.B2SSetdata 61,1
    case 2: Match2.SetValue 1:If B2SOn then Controller.B2SSetdata 62,1
    case 3: Match3.SetValue 1:If B2SOn then Controller.B2SSetdata 63,1
    case 4: Match4.SetValue 1:If B2SOn then Controller.B2SSetdata 64,1
    case 5: Match5.SetValue 1:If B2SOn then Controller.B2SSetdata 65,1
    case 6: Match6.SetValue 1:If B2SOn then Controller.B2SSetdata 66,1
    case 7: Match7.SetValue 1:If B2SOn then Controller.B2SSetdata 67,1
    case 8: Match8.SetValue 1:If B2SOn then Controller.B2SSetdata 68,1
    case 9: Match9.SetValue 1:If B2SOn then Controller.B2SSetdata 69,1
    case 10: Match0.SetValue 1:If B2SOn then Controller.B2SSetdata 70,1
  end Select
  numer0= Score Mod 10
  if Matchnum = numer0   Then
    If B2SOn Then
        Controller.B2SSetCredits Credits
    End If
    playsound SoundFXDOF("knock",119,DOFPulse,DOFKnocker)
    DOF 120, DOFPulse
  end If
End Sub


Sub Table1_KeyDown(ByVal keycode)
  If keycode = PlungerKey Then
    Plunger.PullBack
  End If
  If HSEnterMode Then HighScoreProcessKey(keycode)
  If keycode = LeftFlipperKey Then
    If Tilt=0 then
      If GameInProgress=1 then
        LeftFlipper.RotateToEnd
        PlaySoundAtVol SoundFXDOF("FlipperUp",101,DOFOn,DOFContactors), LeftFlipper, 1
        PlayLoopSoundAtVol "buzzl",LeftFlipper, 1
      End If
    End If
  End If
  If keycode = RightFlipperKey Then
    If Tilt=0 then
      If GameinProgress=1 then
        RightFlipper.RotateToEnd
        PlaySoundAtVol SoundFXDOF("FlipperUp",102,DOFOn,DOFContactors), RightFlipper, 1
        PlayLoopSoundAtVol"buzz", RightFlipper, 1
      End If
    End If
  End If
  If keycode = LeftTiltKey Then
    Nudge 90, 2
    TiltSensor=TiltSensor+50
  End If
  If keycode = RightTiltKey Then
    Nudge 270, 2
    TiltSensor=TiltSensor+50
  End If
  If keycode = CenterTiltKey Then
    Nudge 0, 2
    TiltSensor=TiltSensor+50
  End If
  If keycode = MechanicalTilt Then
    TiltSensor=TiltSensor+50
  End If

  If Keycode = StartGameKey And Not HSEnterMode=true then
    If GameinProgress=0 and credits>0 then
      cball=0
      Wall49.IsDropped=1
      playsound "motor"
      motor.enabled=True
      ballinlane=0
      playedballs=0
      Credits=Credits-1
      If Credits < 1 Then DOF 121, DOFOff
    If B2SOn Then
        Controller.B2SSetCredits Credits
    End If
      If B2SOn then
        Controller.B2SSetMatch Credits+1
        Controller.B2SSetData 35, 0 ' GameOver off
        Controller.B2SSetPlayerUp 1
        Controller.B2SSetTilt 0
        Controller.B2SSetScorePlayer1 0
      End If
      GameStart=1
      playsound "DrainShorter"
      GameinProgress=1
      GameOver.SetValue 0
      Start_Game()
    End If
  End IF
  If Keycode= 6 then
    playsound "coin"
    If credits<47 then Credits=Credits+1:DOF 121, DOFOn:end if
    If B2SOn Then
        Controller.B2SSetCredits Credits
    End If
    If B2SOn then Controller.B2SSetMatch Credits+1
  End if
End Sub

Sub Table1_KeyUp(ByVal keycode)
  If keycode = PlungerKey Then
    Plunger.Fire
    PlaySoundAtVol "Plunger", Plunger, 1
  End If
  If keycode = LeftFlipperKey Then
    If Tilt=0 then
      If GameInProgress=1 then
        LeftFlipper.RotateToStart
        PlaySoundAtVol SoundFXDOF("FlipperDown",101,DOFOff,DOFContactors), LeftFlipper, 1
        stopsound"buzzl"
      End If
    End If
  End If
  If keycode = RightFlipperKey Then
    If GameinProgress=1 then
      If Tilt=0 then
        RightFlipper.RotateToStart
        PlaySoundAtVol SoundFXDOF("FlipperDown",102,DOFOff,DOFContactors), RightFlipper, 1
        stopsound "buzz"
      End If
    End If
  End If
  if keycode=24 and GameStart=0 and GameinProgress=0then
    call Toggleroutine()
    playsound "metal4"
  End If
End Sub

Sub Toggleroutine()
  if ballspergame=3 then
    ballspergame=5
  else ballspergame=3
  end if
end sub

sub Motor_Timer
  'If GameStart=1 then playsound "motor"

End Sub

Sub BallLifter_Timer
  BallLifter.enabled=false
  Kicker1.CreateBall.image="pinball"
  Kicker1.kick 270,2
  ballinlane=1
  BallLifter2.enabled=true
End Sub
Sub BallLifter2_Timer
  BallRelease.IsDropped=False
  BallLifter2.enabled=false
End Sub

Sub TiltTimer_Timer
  If TiltSensor>1 then TiltSensor=TiltSensor-1
  If TiltSensor>99 then
    TiltSensor=0
    If Tilt=0 then playsound "buzzer" End If
    Tilt=1
    TiltLight.SetValue 1
    If B2SOn then Controller.B2SSetTilt 1
  End If
End Sub

Sub UpdateFlipperLogos_Timer
  LFLogo.ObjRotZ = LeftFlipper.CurrentAngle+80
  RFlogo.ObjRotZ = RightFlipper.CurrentAngle+80
End Sub

Sub Start_Game()
  If GameStart=1 and GameinProgress=1 then
    Score=0
    emreel1.ResetToZero()
    EMReel0.ResetToZero()
    Match1.SetValue 0
    Match2.SetValue 0
    Match3.SetValue 0
    Match4.SetValue 0
    Match5.SetValue 0
    Match6.SetValue 0
    Match7.SetValue 0
    Match8.SetValue 0
    Match9.SetValue 0
    Match0.SetValue 0
    If B2SOn then
      Controller.B2SSetdata 61,0
      Controller.B2SSetdata 62,0
      Controller.B2SSetdata 63,0
      Controller.B2SSetdata 64,0
      Controller.B2SSetdata 65,0
      Controller.B2SSetdata 66,0
      Controller.B2SSetdata 67,0
      Controller.B2SSetdata 68,0
      Controller.B2SSetdata 69,0
      Controller.B2SSetdata 70,0
      Controller.B2SSetdata 71,0
    End If
    playsound SoundFXDOF("droptargetreset",106,DOFPulse,DOFContactors)
    Ballstoplay=5
    BallstoPlay=BallsPerGame
    New_Ball
    lBumper1.State=1
    lBumper2.State=0
    lBumper3.State=0
    lbumper4.state=1
    lbumper5.state=1
    light409.state=1
    light410.state=0
    light411.state=0
    light412.state=0
    light413.state=0
    light414.state=1
    light415.state=0
    light416.state=0
    light417.state=0
    light418.state=0
    light419.state=1
    TiltLight.SetValue 0
    light50r.state=1
    light100r.state=0
    light200r.state=0
    light300r.state=0
    Light50Y.state=1
    Light100Y.state=0
    Light200Y.state=0
    Light300Y.state=0
    lightspecial.state=0
    lightspr.state=0
    lightspy.state=0
    lighttriple.state=0
    rpoints=1
    ypoints=1
    Tilt=0
    lpr=0
    lpy=0
  End If
End Sub
' Quick Ball Sound V1.1 by STAT, stefanaustria
' -----------------------
Dim aX, aY, sX, RSound, SBall
Sub TriggerS_Timer()
  aX = int(SBall.VelX): aY = int(SBall.VelY)
  sX = -1.0: If int(Sball.X)>500 Then sX = 1.0
  If (aX>5 OR aY>5) AND Rsound = 0 Then
    RSound=int(RND*4)+1
    PlaySound "Roll"&RSound,0,1.0,sX,0.2 'replace sX to 0 for FullSound
  Elseif (aX<6 AND aY<6) AND Rsound > 0 Then
    StopSound "Roll"&Rsound
    Rsound = 0
  End If
End Sub

'Sub TriggerS_Hit()
Sub TriggerLaunch_Hit()
  DOF 222, DOFPulse
  Set SBall = Activeball
  PlaySound "roll1"
End Sub
' -----------------------

Sub New_Ball
  lBumper1.State =1
  lBumper2.State =0
  lBumper3.State =0
  lightspecial.state=0
' lightspr.state=0
' lightspy.state=0
' lighttriple.state=0
  ballinlane=0
End Sub
dim cball
Sub Drain1_hit
  Drain1.DestroyBall
  PlaySound "drain"
  cball=cball+1
  If cball = 5 Then
    If Ballinlane=0 and ballstoplay>0 then
      'playsound "ballup"
            playsound SoundFXDOF("ballup",114,DOFPulse,DOFContactors)
      BallRelease.isdropped=True
      balllifter.enabled=true
    End If
  end If
End Sub
Sub Trigger1_hit()
' Wall49.IsDropped=0
End Sub

Sub Drain_Hit()
  PlaySound "drain"
  Wall49.IsDropped=0
  DOF 117, DOFPulse
  TriggerS.TimerEnabled = False
  TiltLight.SetValue 0
  Tilt=0
  Drain.DestroyBall
  If B2SOn then Controller.B2SSetTilt 0
  Kicker2.CreateSizedBall(22)
  Kicker2.Kick 90,2
  BallstoPlay=BallstoPlay-1
  If BallstoPlay=0 then
    EndofGame()
    GameOver.SetValue 1
    LeftFlipper.RotateToStart
    RightFlipper.RotateToStart
    motor.enabled=False
    stopsound "motor"
    If B2SOn then Controller.B2SSetGameOver 1
    playsound "motorleer"
    If score>HighScore Then
      HighScore=score
      HighScoreEntryInit()
      UpdatePostIt
      savehs
    End if
  End If
  If BallstoPlay>0 then
'   Credittext.Text=Credits
    If B2SOn Then
        Controller.B2SSetCredits Credits
    End If
    If B2SOn then Controller.B2SSetMatch Credits+1
    New_Ball
    If Ballinlane=0 and ballstoplay>0 then
      'playsound "ballup"
            playsound SoundFXDOF("ballup",114,DOFPulse,DOFContactors)
      BallRelease.isdropped=True
      balllifter.enabled=true
    End If
  End If
End Sub


sub trigger14_hit()
  addscore (10)
    DOF 215, DOFPulse
  If LightSPY.state=0 Then yellowslights
  If LightSPR.state=0 Then redslights
end sub

sub bumper2_hit()
  PlayBumperSound
    playsound SoundFXDOF("bumper_2",106,DOFPulse,DOFContactors)
    DOF 206, DOFPulse
  If lBumper2.State =1 then
    AddScore (10)
  Else
    AddScore (1)
  End If
end sub
sub bumper3_hit()
  PlayBumperSound
    playsound SoundFXDOF("bumper_3",107,DOFPulse,DOFContactors)
    DOF 207, DOFPulse
  If lBumper3.State =1 then
    AddScore (10)
  Else
    AddScore (1)
  End If
end sub

sub bumper4_hit()
  PlayBumperSound
    DOF 208, DOFPulse
  if LightSPY.state=0 Then
    AddScore (10)
    yellowslights
  End If
end sub
sub bumper5_hit()
  PlayBumperSound
    DOF 209, DOFPulse
  If lightspr.state=0 Then
    AddScore (10)
    redslights
  End If
end sub

sub bumper1_hit()
  PlayBumperSound
    playsound SoundFXDOF("bumper_1",105,DOFPulse,DOFContactors)
    DOF 205, DOFPulse
  addscore (1)
  matchnum=matchnum+1
  if matchnum>10 then matchnum=1
end sub

Sub PlayBumperSound()
  Select Case Int(Rnd*3)+1
    Case 1 : PlaySound SoundFX("bumper_1",DOFContactors)
    Case 2 : PlaySound SoundFX("bumper_2",DOFContactors)
    Case 3 : PlaySound SoundFX("bumper_3",DOFContactors)
  End Select
End Sub

sub trigger13_hit()
  PlaysoundAtVol "fx_sensor", ActiveBall, 1
    If lbumper2.state=0 And lbumper3.state=0 Then
    DOF 214, DOFPulse

   end If
  lbumper2.state=1
  lbumper3.state=1
  if lightspr.state=1 and lightspy.state=1 Then
      Lightspecial.state=1
    Lighttriple.state=0
  Else
    lighttriple.state =1
  end If
end sub


dim rpoints, ypoints


sub trigger3_hit()
  PlaysoundAtVol "fx_sensor", ActiveBall, 1
  if light409.state=1 then
     DOF 210, DOFPulse
    addscore 10
    redslights
  else
    addscore rpoints
  end if
end sub
sub trigger4_hit()
  PlaysoundAtVol "fx_sensor", ActiveBall, 1
  if light410.state=1 then
     DOF 210, DOFPulse
    addscore 10
    redslights
  else
    addscore rpoints
  end if
end sub
sub trigger5_hit()
  PlaysoundAtVol "fx_sensor", ActiveBall, 1
  if light411.state=1 then
     DOF 210, DOFPulse
    addscore 10
    redslights
  else
    addscore rpoints
  end if
end sub
sub trigger6_hit()
  PlaysoundAtVol "fx_sensor", ActiveBall, 1
  if light412.state=1 then
     DOF 210, DOFPulse
    addscore 10
    redslights
  else
    addscore rpoints
  end if
end sub
sub trigger7_hit()
  PlaysoundAtVol "fx_sensor", ActiveBall, 1
  if light413.state=1 then
     DOF 210, DOFPulse
    addscore 10
    redslights
    targ=1
    rotatetarget
  else
    addscore rpoints
  end if
end sub
sub trigger8_hit()
  PlaysoundAtVol "fx_sensor", ActiveBall, 1
  if light414.state=1 then
     DOF 211, DOFPulse
    addscore 10
    yellowslights
  else
    addscore ypoints
  end if
end sub
sub trigger9_hit()
  PlaysoundAtVol "fx_sensor", ActiveBall, 1
  if light415.state=1 then
     DOF 211, DOFPulse
    addscore 10
    yellowslights
  else
    addscore ypoints
  end if
end sub
sub trigger10_hit()
  PlaysoundAtVol "fx_sensor", ActiveBall, 1
  if light416.state=1 then
     DOF 211, DOFPulse
    addscore 10
    yellowslights
  else
    addscore ypoints
  end if
end sub
sub trigger11_hit()
  PlaysoundAtVol "fx_sensor", ActiveBall, 1
  if light417.state=1 then
     DOF 211, DOFPulse
    addscore 10
    yellowslights
  else
    addscore ypoints
  end if
end sub
sub trigger12_hit()
  PlaysoundAtVol "fx_sensor", ActiveBall, 1
  if light418.state=1 then
     DOF 211, DOFPulse
    addscore 10
    yellowslights
  else
    addscore ypoints
  end if
end sub
sub trigger15_hit()
  PlaysoundAtVol "fx_sensor", ActiveBall, 1
    DOF 212, DOFPulse
  addscore 10
  yellowslights
end sub
sub trigger16_hit()
  PlaysoundAtVol "fx_sensor", ActiveBall, 1
    DOF 213, DOFPulse
  addscore 10
  redslights
end sub


sub rotatetarget()
  if rotafl=0 Then
    playsound "motorleer"
    rotatet.Interval = 40
    rotatet.Enabled = True
    rotdeg = 0
  End If
End Sub

dim rotdeg, rotafl, gobr, goby, Targ,compry


sub rotatet_timer
  rotafl=1
  rotdeg =rotdeg+5
  If targ=1 Then
    Primitive001.rotz=Primitive001.rotz+5
    Primitive002.rotz=Primitive002.rotz+5
    Primitive003.rotz=Primitive003.rotz+5
    Primitive004.rotz=Primitive004.rotz+5
    Primitive005.rotz=Primitive005.rotz+5
    Primitive006.rotz=Primitive006.rotz+5
  End If
  if rotdeg >=90 Then
    rotatet.Enabled = 0
    rotafl=0
    targ=0
    If compry=1 Then redyellow=redyellow*(-1)
    compry=0
  End If
  If gobr=1 Then
    If Light50R.state =1 Then
      if rotdeg=5 or rotdeg =25 or rotdeg= 45 or rotdeg= 65 Or rotdeg=85 Then AddScore (10)
    Else
      if Light100R.state=1 Then
        if rotdeg = 45 Then AddScore(100)
      Else
        if Light200R.state=1 Then
          if rotdeg=5 or rotdeg=85 Then AddScore(100)
        Else
          if Light300R.state=1 Then
            if rotdeg=5 or rotdeg=45 or rotdeg=85 Then addScore(100)
          Else
            If LightSPR.state=1 Then
              playsound "special1"
              If Credits<47 then Credits=Credits+1:DOF 121, DOFOn:end if
            End If
          End If
        End If
      End If
    End If
    If rotdeg>=90 Then
      gobr=0
    End If
  End If
  If goby=1 Then
    If Light50Y.state=1 Then
      if rotdeg=5 or rotdeg =25 or rotdeg= 45 or rotdeg= 65 Or rotdeg=85 Then AddScore (10)
    Else
      if Light100Y.state=1 Then
        if rotdeg = 45 Then AddScore(100)
      Else
        if Light200Y.state=1 Then
          if rotdeg=5 or rotdeg=85 Then AddScore(100)
        Else
          if Light300Y.state=1 Then
            if rotdeg=5 or rotdeg=45 or rotdeg=85 Then addScore(100)
          Else
            If LightSPY.state=1 Then
              playsound "special1"
              If Credits<47 then Credits=Credits+1:DOF 121, DOFOn:end if
            End If
          End If
        End If
      End If
    End If
    If rotdeg>=90 Then
      goby=0
    End If
  End If
  If compry=1 Then
    if Lighttriple.state = 0 Then
      If rotdeg= 5 Then
        AddScore(10)
        if redyellow= -1 then
          yellowslights()
        Else
          redslights()
        End If
      End If
    Else
      if rotdeg=5 or rotdeg= 45 or rotdeg =85 Then
        AddScore (10)
        if redyellow= -1 then
          yellowslights
        Else
          redslights
        End If
      End If
    End If
  End If
  If compry=2 Then
    playsound "special1"
    If Credits<47 then Credits=Credits+1
        DOF 121, DOFOn
    special.state=0
  End If
  if Primitive001.rotz=360 Then
    Primitive001.rotz=0
    Primitive002.rotz=0
    Primitive003.rotz=0
    Primitive004.rotz=0
    Primitive005.rotz=0
    Primitive006.rotz=0
  End If
end sub

sub redslights()
  If lpr = 4 then Exit Sub
  if light409.state=1 then
    light409.state=0
    light410.state=1
    exit sub
  end if
  if light410.state=1 then
    light410.state=0
    light411.state=1
    exit sub
  end if
  if light411.state=1 then
    light411.state=0
    light412.state=1
    exit sub
  end if
  if light412.state=1 then
    light412.state=0
    light413.state=1
    exit sub
  end if
  if light413.state=1 then
    light413.state=0
    light409.state=1
    targ=1
    rotatetarget
    LPR=LPR+1
    if LPR=1 then
      light100r.state=1
      light50r.state=0
      exit sub
    end if
    if LPR=2 then
      light200r.state=1
      light100r.state=0
      exit sub
    end if
    if LPR=3 then
      light300r.state=1
      light200r.state=0
      exit sub
    end if
      if LPR=4 then
      lightSPr.state=1
      light300r.state=0
      Light50R.state=0
      light100r.state=0
      Light200R.state=0
      Light409.state=0
      rpoints=0
      exit sub
    end if
    exit sub
  end if
end sub


sub yellowslights()
  If lpy=4 Then exit Sub
  if light414.state=1 then
    light414.state=0
    light415.state=1
    exit sub
  end if
  if light415.state=1 then
    light415.state=0
    light416.state=1
    exit sub
  end if
  if light416.state=1 then
    light416.state=0
    light417.state=1
    exit sub
  end if
  if light417.state=1 then
    light417.state=0
    light418.state=1
    exit sub
  end if
  if light418.state=1 then
    light418.state=0
    light414.state=1
    targ=1
    rotatetarget
    LPY=LPY+1
    if LPY=1 then
      light100y.state=1
      light50y.state=0
      exit sub
    end if
    if LPY=2 then
      light200y.state=1
      light100y.state=0
      exit sub
    end if
    if LPY=3 then
      light300y.state=1
      light200y.state=0
      exit sub
    end if
    if LPY=4 then
      lightSPy.state=1
      light300y.state=0
      Light50Y.state=0
      Light100Y.state=0
      Light200Y.state=0
      Light414.state=0
      ypoints=0
      exit sub
    end if
    exit sub
  end if
end sub



Sub gobblerr_Hit()
  gobblerR.TimerEnabled=True
  gobblerr.TimerInterval=2000
  Set aBall=ActiveBall      ' aBall = the ball Being Gobbled
  aXpos=aBall.X
  aXadd=(gobblerR.X-aXpos)/15     ' 15 ticks to roll ball into hole
  aYpos=aBall.Y
  aYadd=(gobblerR.Y-aYpos)/15
  aBall.X=aXpos         ' Put Ball Back
  aBall.Y=aYpos
  aZpos=50            ' Including Starting Height
  aBall.Z=aZPos
  aBall.VelX=0          ' Stop It Rolling
  aBall.VelY=0
  gobblerR.TimerInterval=3      ' Set timer to 3ms
  gobblerR.TimerEnabled=True      ' And Turn It On
  compry=0
  gobr=1
  rotatetarget
End Sub


Sub gobblerr_Timer()
  playsound "kick"
  aZpos = aZpos - 1       'Subtract 1 from ball Z position and repeat the line above
  If aZpos > 25 Then
    aYpos = aYpos + aYadd   ' Move X & X towards Hole Center
    aXpos = aXpos + aXadd
    wall016.visible=0
  Else
    aYpos=gobblerR.Y          'Ball Now in the hole
    aXpos=gobblerR.X
  End If
  aBall.Z=aZpos         ' Move the ball Z position to the value of variable aZpos
  aBall.X=aXpos         ' Roll to Center of Kicker
  aBall.Y=aYpos
  aBall.VelX=0          ' Stop It Really Rolling
  aBall.VelY=0
  If aZpos > -40 Then       'If the ball Z position is above -30 cycle
    Exit Sub
  End If
'
'------------------------------------
' Gobble Finished, Do Scoring & Drain
'------------------------------------
'
  Wall016.visible=1
  gobblerR.TimerEnabled=False     ' Stop this timer
  gobblerR.DestroyBall            ' Delete Ball From Table
  Gob_End()
End Sub


Sub gobblery_Hit()
  gobblery.TimerEnabled=True
  gobblery.TimerInterval=2000
  Set aBall=ActiveBall      ' aBall = the ball Being Gobbled
  aXpos=aBall.X
  aXadd=(gobblery.X-aXpos)/15     ' 15 ticks to roll ball into hole
  aYpos=aBall.Y
  aYadd=(gobblery.Y-aYpos)/15
  aBall.X=aXpos         ' Put Ball Back
  aBall.Y=aYpos
  aZpos=50            ' Including Starting Height
  aBall.Z=aZPos
  aBall.VelX=0          ' Stop It Rolling
  aBall.VelY=0
  gobblery.TimerInterval=3      ' Set timer to 3ms
  gobblery.TimerEnabled=True      ' And Turn It On
  compry=0
  goby=1
  rotatetarget
End Sub


Sub gobblery_Timer()
  playsound "kick"
  aZpos = aZpos - 1       'Subtract 1 from ball Z position and repeat the line above
  If aZpos > 25 Then
    aYpos = aYpos + aYadd   ' Move X & X towards Hole Center
    aXpos = aXpos + aXadd
    Wall017.visible=0
  Else
    aYpos=gobblery.Y          'Ball Now in the hole
    aXpos=gobblery.X
  End If
  aBall.Z=aZpos         ' Move the ball Z position to the value of variable aZpos
  aBall.X=aXpos         ' Roll to Center of Kicker
  aBall.Y=aYpos
  aBall.VelX=0          ' Stop It Really Rolling
  aBall.VelY=0
  If aZpos > -40 Then       'If the ball Z position is above -30 cycle
    Exit Sub
  End If

'
'------------------------------------
' Gobble Finished, Do Scoring & Drain
'------------------------------------
'
  Wall017.visible=1
  gobblery.TimerEnabled=False     ' Stop this timer
  gobblery.DestroyBall            ' Delete Ball From Table
  Gob_End()
End Sub


Sub Gob_End()
  Lighttriple.state=0
  Lbumper2.state=0
  Lbumper3.state=0
  playsound"fx_drain"
    DOF 218, DOFPulse
  gobend.Interval=1000      ' Set timer to 3ms
  gobend.Enabled=True
End Sub

sub gobend_timer()
  gobend.Enabled=0
  playsound "ballup"
Drain_Hit
end Sub







dim redyellow
redyellow=-1


Sub MainTarget_Hit()
  if rotafl=0 Then
    If Tilt=0 then
         DOF 116, DOFPulse
         DOF 216, DOFPulse

      targ=1
      compry=1
      if Lightspecial.state=1 then compry=2
      rotatetarget
    End If
  End If
End Sub



Sub Gate_hit
  PlaysoundAtVol "gate", ActiveBall, 1
End Sub


Sub RightSlingShot_Slingshot
    playsoundAtVol SoundFXDOF("right_slingshot",104,DOFPulse,DOFContactors),sling1, 1
  DOF 204,DOFPulse
    RSling0.Visible = 0
    RSling1.Visible = 1
    sling1.TransZ = -20
    RStep = 0
    RightSlingShot.TimerEnabled = 1
  If Tilt=0 then
    Addscore (1)
  End If
End Sub
dim lstep, rstep
Sub RightSlingShot_Timer
    Select Case RStep
        Case 3:RSLing1.Visible = 0:RSLing2.Visible = 1:sling1.TransZ = -10
        Case 4:RSLing2.Visible = 0:RSling0.Visible = 1:sling1.TransZ = 0:RightSlingShot.TimerEnabled = 0
    End Select
    RStep = RStep + 1
End Sub

Sub LeftSlingShot_Slingshot
    playsoundAtVol SoundFXDOF("left_slingshot",103,DOFPulse,DOFContactors), sling2, 1
  DOF 203,DOFPulse
    LSling0.Visible = 0
    LSling1.Visible = 1
    sling2.TransZ = -20
    LStep = 0
    LeftSlingShot.TimerEnabled = 1
  If Tilt=0 then
    Addscore (1)
  End If
End Sub

Sub LeftSlingShot_Timer
    Select Case LStep
        Case 3:LSLing1.Visible = 0:LSLing2.Visible = 1:sling2.TransZ = -10
        Case 4:LSLing2.Visible = 0:LSLing0.Visible = 1:sling2.TransZ = 0:LeftSlingShot.TimerEnabled = 0
    End Select
    LStep = LStep + 1
End Sub


Sub AddScore (points)
  Select Case points
    Case 1:Playsound SoundFXDOF("10p",141,DOFPulse,DOFChimes)
    Case 10:Playsound SoundFXDOF("100p",142,DOFPulse,DOFChimes)
    Case 100:Playsound SoundFXDOF("100p",143,DOFPulse,DOFChimes)
  End Select
  Score=Score + points
  If B2SOn then Controller.B2SSetScorePlayer1 Score
  If score>999 and Replay=0 then
    Playsound SoundFXDOF("knock",119,DOFPulse,DOFKnocker)
    DOF 120, DOFPulse
    If Credits<47 then Credits=Credits+1:DOF 121, DOFOn:end if
    If B2SOn Then
        Controller.B2SSetCredits Credits
    End If
    If B2SOn then Controller.B2SSetMatch Credits+1
    Replay=1
  End If
  If score>1400 and Replay=1 then
    Playsound SoundFXDOF("knock",119,DOFPulse,DOFKnocker)
    DOF 120, DOFPulse
    If Credits<47 then Credits=Credits+1:DOF 121, DOFOn:end if
    If B2SOn Then
        Controller.B2SSetCredits Credits
    End If
    If B2SOn then Controller.B2SSetMatch Credits+1
    Replay=2
  End If
  If score>1600 and Replay=2 then
    Playsound SoundFXDOF("knock",119,DOFPulse,DOFKnocker)
    DOF 120, DOFPulse
    If Credits<47 then Credits=Credits+1:DOF 121, DOFOn:end if
    If B2SOn Then
        Controller.B2SSetCredits Credits
    End If
    If B2SOn then Controller.B2SSetMatch Credits+1
    Replay=3
  End If
  If score>1800 and Replay=3 then
    Playsound SoundFXDOF("knock",119,DOFPulse,DOFKnocker)
    DOF 120, DOFPulse
    If Credits<47 then Credits=Credits+1:DOF 121, DOFOn:end if
    If B2SOn Then
        Controller.B2SSetCredits Credits
    End If
    If B2SOn then Controller.B2SSetMatch Credits+1
    Replay=4
  End If
  emreel1.AddValue(points)
  if score >=1000 Then
    EMReel0.setvalue (1)
    If B2SOn then Controller.B2SSetdata 71,1
  Else
    EMReel0.setvalue (0)
    If B2SOn then Controller.B2SSetdata 71,0
  end If
End Sub






Sub Table1_Exit()
  If B2SOn Then Controller.Stop
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
'                      Supporting Ball & Sound Functions
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

Const tnob = 6 ' total number of balls
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
  PlaySound("fx_collide"), 0, Csng(velocity) ^2 / 2000, AudioPan(ball1), 0, Pitch(ball1), 0, 0, AudioFade(ball1)
End Sub

Sub Rubbers_Hit(idx)
  Select Case Int(Rnd*2)+1
    Case 1 : PlaySound "rubber_hit_1", 0, Vol(ActiveBall)*400, Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
    Case 1 : PlaySound "rubber_hit_2", 0, Vol(ActiveBall)*400, Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
    Case 2 : PlaySound "rubber_hit_3", 0, Vol(ActiveBall)*400, Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
  End Select
End Sub

Sub Slings1_slingshot(idx)
  PlaySound "rubber_hit_1", 0, Vol(ActiveBall)*400, Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
  Addscore(1)
End Sub
Sub Slings10_slingshot(idx)
  PlaySound "rubber_hit_1", 0, Vol(ActiveBall)*400, Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
  Addscore(10)
End Sub

Sub LeftFlipper_Collide(parm)
  RandomSoundFlipper()
End Sub

Sub RightFlipper_Collide(parm)
  RandomSoundFlipper()
End Sub

Sub RandomSoundFlipper()
  Select Case Int(Rnd*3)+1
    Case 1 : PlaySound "flip_hit_1", 0, Vol(ActiveBall)*400, Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
    Case 2 : PlaySound "flip_hit_2", 0, Vol(ActiveBall)*400, Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
    Case 3 : PlaySound "flip_hit_3", 0, Vol(ActiveBall)*400, Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
  End Select
End Sub

'******************************
' Diverse Collection Hit Sounds
'******************************

Sub aMetals_Hit(idx):PlaySound "fx_MetalHit", 0, Vol(ActiveBall), pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0, AudioFade(ActiveBall):End Sub
Sub aRubber_Bands_Hit(idx):PlaySound "fx_rubber", 0, Vol(ActiveBall), pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0, AudioFade(ActiveBall):End Sub
Sub aRubber_Posts_Hit(idx):PlaySound "fx_postrubber", 0, Vol(ActiveBall), pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0, AudioFade(ActiveBall):End Sub
Sub aRubber_Pins_Hit(idx):PlaySound "fx_postrubber", 0, Vol(ActiveBall), pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0, AudioFade(ActiveBall):End Sub
Sub aYellowPins_Hit(idx):PlaySound "fx_postrubber", 0, Vol(ActiveBall), pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0, AudioFade(ActiveBall):End Sub
Sub aPlastics_Hit(idx):PlaySound "fx_PlasticHit", 0, Vol(ActiveBall), pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0, AudioFade(ActiveBall):End Sub
Sub aGates_Hit(idx):PlaySound "fx_Gate", 0, Vol(ActiveBall), pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0, AudioFade(ActiveBall):End Sub
Sub aWoods_Hit(idx):PlaySound "fx_Woodhit", 0, Vol(ActiveBall), pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0, AudioFade(ActiveBall):End Sub
Sub aCaptiveWalls_Hit(idx):PlaySound "fx_collide", 0, Vol(ActiveBall), pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0, AudioFade(ActiveBall):End Sub



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

dim brarray
brarray = array("ball go", "ball 1","ball 2","ball 3","ball 4","ball 5","ball")
sub ballaff(nbball)
dim nbll
  nbll=nbball
  if nbll>5 then nbll =6
  ballin.image = brarray(nbll)
end sub


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
  HSScorex = highscore
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

    If keycode = PlungerKey Then
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
sub savehs
    savevalue "4roses", "credit", Credits
    savevalue "4roses", "hiscore", Highscore
    savevalue "4roses", "score1", score
' savevalue "4roses4", "balls", balls
  savevalue "4roses", "hsa1", HSA1
  savevalue "4roses", "hsa2", HSA2
  savevalue "4roses", "hsa3", HSA3
end sub

sub loadhs
    dim temp
    dim temp2
  temp = LoadValue("4roses", "credit")
    If (temp <> "") then Credits = CDbl(temp)
    temp = LoadValue("4roses", "hiscore")
    If (temp <> "") then Highscore = CDbl(temp)
    temp = LoadValue("4roses", "score1")
    If (temp <> "") then score = CDbl(temp)
'    temp = LoadValue("4roses4", "balls")
'    If (temp <> "") then balls = CDbl(temp)
  temp = LoadValue("4roses", "hsa1")
    If (temp <> "") then HSA1 = CDbl(temp)
    temp = LoadValue("4roses", "hsa2")
    If (temp <> "") then HSA2 = CDbl(temp)
    temp = LoadValue("4roses", "hsa3")
    If (temp <> "") then HSA3 = CDbl(temp)
end sub

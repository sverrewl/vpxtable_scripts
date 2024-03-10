' Volcano Gottlieb 1981
' Allknowing2012
' Cool site for Switches and lamps http://replay.pinball.ch/system80pins/Volcano-tec.html
'
' Credits to original Volcano VP9 for code segments
'
' BUGS:  Tilt doesnt seem to work
'
'1.02 - graphic updates from Hauntfreaks
'1.03 - Fixes from 32assassin & dof updates from Arngrim
'1.04 - remove dbug code and fix slingshot sounds
'
Option Explicit
Randomize

Const cGameName="vlcno_ax"
Const UseSolenoids=2,UseLamps=True,UseGI=0,UseSyn=1,SSolenoidOn="SolOn",SSolenoidOff="Soloff",SFlipperOn="FlipperUp",SFlipperOff="FlipperDown"
Const SCoin="coin3",cCredits=""

On Error Resume Next
ExecuteGlobal GetTextFile("controller.vbs")
If Err Then MsgBox "Can't open controller.vbs"
On Error Goto 0

LoadVPM "01560000","sys80.vbs",2.33
Dim DesktopMode: DesktopMode = Table1.ShowDT
Dim rcpos
Dim rot

SolCallback(1)="SubwayRelease"      'Subway Ball Release
SolCallback(2)="dtL_SolDropUp"    'dtL.SolDropUp
SolCallback(5)="SubWayKick"       'Subway Kicker
SolCallback(6)="dtR_SolDropUp"   'dtR.SolDropUp
SolCallback(8)="vpmSolSound SoundFX(""Knocker"",DOFKnocker),"
SolCallback(9)="SolTroughIn"
'SolCallback(sLLFlipper)="SolLFlipper"  'VpmSolFlipper Flipper1,nothing,"
'SolCallback(sLRFlipper)="SolRFlipper"   'VpmSolFlipper Flipper2,Flipper3,"

SolCallback(sLRFlipper) = "SolRFlipper"
SolCallback(sLLFlipper) = "SolLFlipper"

Dim LockBalls
LockBalls=0
Dim BallsInPlay
BallsInPlay=0
Dim DrainBall
DrainBall=False

Dim bsSaucer,dtL,dtR,bsSaucer2
Dim rc
Sub Table1_Init

' Thalamus : Was missing 'vpminit me'
  vpminit me

    rot = 0
    for each rc in RCWalls
      rc.isdropped=True
    Next
    RCWalls(0).isdropped=False
  CheckGate
 On Error Resume Next
  With Controller
   .GameName=cGameName
   If Err Then MsgBox "Can't start Game" & cGameName & vbNewLine & Err.Description : Exit Sub
    .SplashInfoLine=cCredits
    .HandleMechanics=0
    .HandleKeyboard=0
   .ShowDMDOnly=1
    .ShowFrame=0
    .Hidden=True
    .ShowTitle=0
  End With
  Controller.SolMask(0)=0
 vpmTimer.AddTimer 1000,"Controller.SolMask(0)=&Hffffffff'"
  Controller.Run
  If Err Then MsgBox Err.Description
  On Error Goto 0

 PinMAMETimer.Interval=PinMAMEInterval
 PinMAMETimer.Enabled=1

  vpmNudge.TiltSwitch=46     'swTilt
 vpmNudge.Sensitivity=2
  vpmNudge.TiltObj=Array(LeftSlingshot,RightSlingshot,Bumper1,Bumper2,Bumper3,Bumper4)
  vpmNudge.SolGameOn(True)

  Set bsSaucer=New cvpmBallStack
  bsSaucer.InitSaucer Kicker1,43,0,15
 bsSaucer.InitExitSnd SoundFX("popper",DOFContactors),SoundFX("popper",DOFContactors)
  bsSaucer.KickForceVar = 3
  bsSaucer.KickAngleVar = 3

  Set bsSaucer2=New cvpmBallStack
 bsSaucer2.InitSaucer Kicker2,23,200+Int(RND*4),7
  bsSaucer2.InitExitSnd SoundFX("popper",DOFContactors),SoundFX("popper",DOFContactors)
  bsSaucer2.KickForceVar = 3
  bsSaucer2.KickAngleVar = 3

 set dtL=New cvpmDropTarget
  dtL.InitDrop Array(target1,target2,target3,target4,target5),Array(2,12,22,32,42)
  dtL.initsnd SoundFX("DTC",DOFContactors),SoundFX("DTReset",DOFContactors)

 set dtR=New cvpmDropTarget
  dtR.InitDrop Array(target6,target7,target8,target9,target10),Array(1,11,21,31,41)
 dtR.initsnd SoundFX("DTC",DOFContactors),SoundFX("DTReset",DOFContactors)

    Plunger1.Pullback
    vpmMapLights InsertLights
    sLight94.state=LightStateOff   'BallSave

    rotatecam.interval=300
  rotatecam.enabled=false
End Sub

Sub Table1_Exit
  If B2SOn then Controller.Stop
End Sub

Sub Table1_KeyDown(ByVal keycode)
  If KeyCode=LeftMagnaSave AND CurrentE=1 Then   ' passing thru the lane and ballsave is lit
    Plunger1.Fire:PlaySoundAtVol "Plunger", Plunger, 1
    Plunger1.TimerEnabled=1
        sLight94.state=LightStateOff   'BallSave
  End If

  If KeyCode=RightMagnaSave Then Controller.Switch(56)=1 'Right cabinet green button -- '
  If KeyCode=PlungerKey Then Plunger.Pullback

    If vpmKeyDown(KeyCode) Then Exit Sub
End Sub

Sub Table1_KeyUp(ByVal keycode)
  If KeyCode=RightMagnaSave Then Controller.Switch(56)=0
  If KeyCode=PlungerKey Then Plunger.Fire:PlaySoundAtVol "Plunger", Plunger, 1
  If vpmKeyUp(KeyCode) Then Exit Sub
End Sub

Sub SolTroughIn(Enabled)
  If Enabled Then
  Controller.Switch(20)=0
  If DrainBall=True Then Drain.DestroyBall:BallsInPlay=BallsInPlay-1
  If BallsInPlay=0 Then
        sLight94.state=LightStateOff ' reset BallSave
    Controller.Switch(30)=1
    Controller.Switch(40)=1
  Else
    Controller.Switch(30)=1
    Controller.Switch(40)=0
  End If
  DrainBall=False
    debug.print "Trough Release Lock=" & LockBalls & " BIP:" & BallsInPlay
  End If
End Sub

Sub dtL_SolDropUp(enabled)
  if enabled Then
    dtLTimer.interval=500
    dtLTimer.enabled=True
  end if
End Sub

Sub dtLTimer_Timer
    dtLTimer.enabled=False
    dtL.DropSol_On
End Sub

Sub dtR_SolDropUp(enabled)
  if enabled Then
    dtRTimer.interval=500
    dtRTimer.enabled=True
  end if
End Sub

Sub dtRTimer_Timer
    dtRTimer.enabled=False
    dtR.DropSol_On
End Sub

Sub SolLFlipper(Enabled)
     If Enabled Then
         PlaySoundAtVol SoundFx("FlipperUpLeft",DOFContactors), Flipper1, 1:Flipper1.RotateToEnd
     Else
         PlaySoundAtVol SoundFx("FlipperDown",DOFContactors), Flipper1, 1:Flipper1.RotateToStart
     End If
End Sub

Sub SolRFlipper(Enabled)
     If Enabled Then
         PlaySoundAtVol SoundFx("FlipperUpRight",DOFContactors), Flipper2, 1:Flipper2.RotateToEnd:Flipper3.RotateToEnd
     Else
         PlaySoundAtVol SoundFx("FlipperDown",DOFContactors), Flipper2, 1:Flipper2.RotateToStart:Flipper3.RotateToStart
     End If
End Sub

' Left Targets
Sub target1_Hit
 If light19b.State=1 then sLight94.state=1 'ball save when lit
 dtL.Hit 1
End Sub
Sub target2_Hit
  If light20b.State=1 then sLight94.state=1 'ball save when lit
  dtL.Hit 2
  End Sub
Sub target3_Hit
  If light21b.State=1 then sLight94.state=1 'ball save when lit
  dtL.Hit 3
End Sub
Sub target4_Hit
  If light22b.State=1 then sLight94.state=1 'ball save when lit
    dtL.Hit 4
End Sub
Sub target5_Hit
  If light23b.State=1 then sLight94.state=1 'ball save when lit
    dtL.Hit 5
End Sub

' Right Targets
Sub target6_Hit
  If light19.State=1 then sLight94.state=1 'ball save when lit
    dtR.Hit 1
End Sub
Sub target7_Hit
  If light20.State=1 then sLight94.state=1 'ball save when lit
    dtR.Hit 2
End Sub
Sub Target8_Hit
  If light21.State=1 then sLight94.state=1 'ball save when lit
    dtR.Hit 3
End Sub
Sub target9_Hit
  If light22.State=1 then sLight94.state=1 'ball save when lit
    dtR.Hit 4
End Sub
Sub target10_Hit
  If light23.State=1 then sLight94.state=1 'ball save when lit
    dtR.Hit 5
End Sub

Sub sw15_walls_hit(idx)
  vpmTimer.PulseSw 15
  RandomSoundRubber()
End Sub

Sub sw5_Hit:Controller.Switch(5)=1:End Sub              '5
Sub sw5_unHit:Controller.Switch(5)=0:End Sub
Sub sw6_Hit:Controller.Switch(6)=1:End Sub        '6
Sub sw6_unHit:Controller.Switch(6)=0:End Sub

Dim dBall1,dZpos1,dBall2,dZpos2,dBall3,dZpos3,dBall4,dZpos4

Sub Kicker3_Hit:crater4.isdropped=1:set dBall4 = ActiveBall:dZPos4 = 60:Me.TimerInterval=4:Me.TimerEnabled=1:LastHit=1:vpmTimer.PulseSwitch(4),100,"WorkLock":End Sub     '4
Sub Kicker4_Hit:crater3.isdropped=1:set dBall3 = ActiveBall:dZPos3 = 50:Me.TimerInterval=4:Me.TimerEnabled=1:LastHit=2:vpmTimer.PulseSwitch(14),100,"WorkLock":End Sub      '14
Sub Kicker5_Hit:crater2.isdropped=1:set dBall2 = ActiveBall:dZPos2 = 40:Me.TimerInterval=4:Me.TimerEnabled=1:LastHit=3:vpmTimer.PulseSwitch(24),100,"WorkLock":End Sub     '24
Sub Kicker6_Hit:crater1.isdropped=1:set dBall1 = ActiveBall:dZPos1 = 30:Me.TimerInterval=4:Me.TimerEnabled=1:LastHit=4:vpmTimer.PulseSwitch(34),100,"WorkLock":End Sub      '34

Sub kicker3_timer
  dBall4.Z = dZpos4
  dZpos4 = dZpos4-4
  if dZPos4 < 20 Then
    Me.TimerEnabled=0
    Me.DestroyBall
    crater4.isdropped=False
  end if
End sub

Sub kicker4_timer
  dBall3.Z = dZpos3
  dZpos3 = dZpos3-4
  if dZPos3 < 10 Then
    Me.TimerEnabled=0
    Me.DestroyBall
    crater3.isdropped=False
  end if
End sub
Sub kicker5_timer
  dBall2.Z = dZpos2
  dZpos2 = dZpos2-4
  if dZPos2 <  0 Then
    Me.TimerEnabled=0
    Me.DestroyBall
    crater2.isdropped=False
  end if
End sub
Sub kicker6_timer
  dBall1.Z = dZpos1
  dZpos1 = dZpos1-4
  if dZPos1 < 0 Then
    Me.TimerEnabled=0
    Me.DestroyBall
    crater1.isdropped=False
  end if
End sub

Sub sw16_Hit:Controller.Switch(16)=1:End Sub              '16
Sub sw16_unHit:Controller.Switch(16)=0:End Sub
Sub Drain_Hit:Controller.Switch(20)=1:DrainBall=True:End Sub       '20/30/40

                      '22
Sub Kicker2_Hit:bsSaucer2.AddBall 0:End Sub                  '23
Sub gate3_Hit:Controller.Switch(25)=1:End Sub                '25

Sub sw25_unHit:Controller.Switch(25)=0:End Sub
Sub sw26_Hit:vpmTimer.PulseSw 26:End Sub                  '26
                                     '31
                                     '32
Sub sw36_Hit:vpmTimer.PulseSw 36:End Sub                 '36
                                     '41
                                     '42
Sub Kicker1_Hit:bsSaucer.AddBall 0:End Sub                 '43
Sub Spinner1_Spin:vpmTimer.PulseSw 44:End Sub

'***Lighting***

Dim yy:For each yy in GIlights:yy.intensity=40:Next   'Global Lighting Adjustment Optional
       For each yy in alights:yy.intensity=180:Next   'Bumpers

Sub GIOn
  dim bulb
  for each bulb in GILights
  bulb.state = 1
  next
    for each bulb in aLights
    bulb.state = 1
  next
End Sub

Sub GIOff
  dim bulb
  for each bulb in GILights
  bulb.state = 0
  next
    for each bulb in aLights
  bulb.state = 0
  next
End Sub
Sub trigger1_hit:light1.state=2: End Sub
Sub trigger2_hit:light1.state=0: End Sub
Sub trigger3_hit
dim bulb
l5g.state=2
for each bulb in aLights
    bulb.state = 2
  next
 End Sub
Sub trigger4_hit
dim bulb
l5g.state=0
for each bulb in aLights
    bulb.state = 1
  next
End Sub
Sub trigger5_hit
dim bulb
l5g.state=0
for each bulb in aLights
    bulb.state = 1
  next
End Sub
Sub Trigger6_hit
dim bulb
  for each bulb in GILights
  bulb.state = 2
  next
End Sub
Sub Trigger6_Unhit
dim bulb
  for each bulb in GILights
  bulb.state = 1
  next
End Sub

' These are for testing as they may trigger events
Light12.visible=0 ' Enable RotateCam,  Ball in Plunger
Light2.visible=0    'Mostly on GI using
Light8.visible=0    'left side hole eject
Light13.visible=0    ' unused
Light14.visible=0  'gate open
Light15.visible=0   'ball release trough
Light16.visible=0  'top hole eject

'1 Match Light

'8  Fire Pit
'14 Ball Gate
'15 Ball Release
'16 Hole Kicker
'13 Ball Saver Relay
'12 Motor Relay

Set LampCallback=GetRef("UpdateMultipleLamps")

Dim CurrentA,OldA,CurrentB,OldB,CurrentC,OldC,CurrentD,OldD,CurrentE,OldE,CurrentF,OldF,Current2,Old2
OldB=0:OldC=0:OldD=0:OldA=0:OldE=0:Old2=0:OldF=0

Sub UpdateMultipleLamps
  L1=Light24.State  ' Crater PF
  L2=Light25.State
  L3=Light26.State
  L4=Light27.State

    CurrentA=Light8.State
   If CurrentA<>OldA Then
      If CurrentA=1 Then
        bsSaucer.ExitSol_On
     End If
    End If
  OldA=CurrentA
 CurrentB=Light15.State
    If CurrentB<>OldB Then
      If CurrentB=1 Then
        If BallsInPlay<3 Then
       BallsInPlay=BallsInPlay+1
       BallRelease.CreateBall
        BallRelease.Kick 75,5
                debug.print "Light15 Release Lock=" & LockBalls & " BIP:" & BallsInPlay
        Select Case BallsInPlay
         Case 1:Controller.Switch(30)=1:Controller.Switch(40)=0
          Case 3:Controller.Switch(30)=0:Controller.Switch(40)=0
        End Select
        End If
      End If
    End If
  OldB=CurrentB

  CurrentC=Light14.State
  If CurrentC<>OldC Then
        OldC=CurrentC
   CheckGate
 End If

  CurrentD=Light16.State
    If CurrentD<>OldD Then
      If CurrentD=1 Then
        bsSaucer2.ExitSol_On
      End If
    End If
  OldD=CurrentD

  CurrentE=sLight94.State

    CurrentF=Light12.State    ' RotateCam
 If CurrentF<>OldF Then
    OldF=CurrentF
      If CurrentF=1 then
        rotatecam.interval=300
    rotatecam.enabled=true
      Else
        rotatecam.enabled=False
      End If
    End if

    Current2=Light2.State    ' GI Lighting
    if Current2 <> Old2 Then
      Old2=Current2
      if Current2=1 Then
        GIOn()
      Else
        GIOff()
      End if
    End If
End Sub

Sub RotateCam_Timer() 'Shooter Lane Guide
  rcpos = round(rot/45)
  RCWalls(rcpos).isdropped=True
  rot=rot + 4
  if rot > 360 Then  rot = 0
  RotateCamP.ObjRotZ = rot
  rcpos = round(rot/45)
  RCWalls(rcpos).isdropped=False
'msgbox "rcpos=" & rcpos
End Sub

Dim L1,L2,L3,L4
Dim LockOn
LockOn=0
Dim LastHit
LastHit=0
L1=0:L2=0:L3=0:L4=0

Sub WorkLock(swNo)
  LockOn=0
  Select Case LastHit
  Case 1:If L1>0 Then LockOn=1   ' The LockLight was On
 Case 2:If L2>0 Then LockOn=1
  Case 3:If L3>0 Then LockOn=1
  Case 4:If L4>0 Then LockOn=1
  End Select
  LockBalls=LockBalls+1
  Controller.Switch(0)=1
  If LockOn=1 And LockBalls>2 Then LockOn=0 ' Already have 2 locked balls
  debug.print "WorkLock Lock=" & LockBalls & " BIP:" & BallsInPlay & "LockedBall:" & LockOn
End Sub

Sub SubWayKick(Enabled)
  If Enabled Then
  Controller.Switch(0)=0
  Controller.Switch(10)=1
  End If
End Sub

Sub SubwayRelease(Enabled)
  debug.print "SR: 0"
  If Enabled Then
    debug.print "SR: 1"
  '  If LockOn=0 Then
      debug.print "SR: 2"
   If LockBalls>0 Then
        debug.print "SR: 3"
   Kicker7.CreateBall
    Kicker7.Kick 180,5:playsoundAtVol SoundFX("popper",DOFContactors), Kicker7, 1
    LockBalls=LockBalls-1
   End If
    If LockBalls=0 Then Controller.Switch(10)=0
  '  End If
  End If
  debug.print "Subway Release Lock=" & LockBalls & " BIP:" & BallsInPlay
End Sub

Sub Plunger1_Timer
  Plunger1.TimerEnabled=0
  Plunger1.PullBack
End Sub

Sub FlipperTimer_Timer()
   Primitiveflipper.RotY = OutGate.Currentangle  ' Ballsave
   LFlip.RotY = Flipper1.CurrentAngle
   RFlip.RotY = Flipper2.CurrentAngle
   RFlip1.RotY = Flipper3.CurrentAngle
lfs.RotZ = Flipper1.CurrentAngle
  rfs.RotZ = Flipper2.CurrentAngle
 tfs.RotZ = Flipper3.CurrentAngle

End Sub

Sub CheckGate                 'Determine which direction the gate has to move
 If Light14.State=0 Then           'If gate is closing then
        OutGate.RotateToStart
        debug.print "Gate to Start"
        Q15.state=LightStateOff
 Else                    'If gate is opening then
        OutGate.RotateToEnd
        debug.print "Gate to End"
        Q15.state=LightStateOn
  End If
End Sub

Dim bump1,bump2,bump3,bump4

Sub Bumper1_Hit:vpmTimer.PulseSw 35:bump1 = 1:Me.TimerEnabled = 1:DOF 106, DOFPulse:light10.state=0:vpmtimer.addtimer 200, "light10.state=1'":End Sub
Sub Bumper1_Timer()
  Select Case bump1
        Case 1:Ring1.Z = -30:bump1 = 2
        Case 2:Ring1.Z = -20:bump1 = 3
        Case 3:Ring1.Z = -10:bump1 = 4
        Case 4:Ring1.Z = 0:Me.TimerEnabled = 0
  End Select
End Sub

Sub Bumper2_Hit:vpmTimer.PulseSw 13:bump2 = 1:Me.TimerEnabled = 1:DOF 103, DOFPulse:light48.state=0:vpmtimer.addtimer 200, "light48.state=1'":End Sub
Sub Bumper2_Timer()
  Select Case bump2
        Case 1:Ring2.Z = -30:bump2 = 2
        Case 2:Ring2.Z = -20:bump2 = 3
        Case 3:Ring2.Z = -10:bump2 = 4
        Case 4:Ring2.Z = 0:Me.TimerEnabled = 0
  End Select
End Sub

Sub Bumper3_Hit:vpmTimer.PulseSw 13:bump3 = 1:Me.TimerEnabled = 1:DOF 104, DOFPulse:light11.state=0:vpmtimer.addtimer 200, "light11.state=1'":End Sub
Sub Bumper3_Timer()
  Select Case bump3
        Case 1:Ring3.Z = -30:bump3 = 2
        Case 2:Ring3.Z = -20:bump3 = 3
        Case 3:Ring3.Z = -10:bump3 = 4
        Case 4:Ring3.Z = 0:Me.TimerEnabled = 0
  End Select
End Sub

Sub Bumper4_Hit:vpmTimer.PulseSw 35:bump4 = 1:Me.TimerEnabled = 1:DOF 105, DOFPulse:light9.state=0:vpmtimer.addtimer 200, "light9.state=1'":End Sub
Sub Bumper4_Timer()
  Select Case bump4
        Case 1:Ring4.Z = -30:bump4 = 2
        Case 2:Ring4.Z = -20:bump4 = 3
        Case 3:Ring4.Z = -10:bump4 = 4
        Case 4:Ring4.Z = 0:Me.TimerEnabled = 0
  End Select
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

' *********************************************************************
'                      Supporting Ball & Sound Functions
' *********************************************************************

Sub Pins_Hit (idx)
  PlaySound "pinhit_low", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0, AudioFade(ActiveBall)
End Sub

Sub Targets_Hit (idx)
  PlaySound SoundFX("target",DOFContactors), 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0, AudioFade(ActiveBall)
End Sub

Sub TargetBankWalls_Hit (idx)
  PlaySound "target", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0, AudioFade(ActiveBall)
End Sub

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
 '  debug.print "rubber"
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
  '  debug.print "Posts"
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

Sub Bumpers_Hit(idx)
  Select Case Int(Rnd*4)+1
    Case 1 : PlaySound SoundFx("fx_bumper2",DOFContactors), 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
    Case 2 : PlaySound SoundFx("fx_bumper2",DOFContactors), 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
    Case 3 : PlaySound SoundFx("fx_bumper3",DOFContactors), 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
    Case 4 : PlaySound SoundFx("fx_bumper4",DOFContactors), 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
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

Function Vol(ball) ' Calculates the Volume of the sound based on the ball speed
    Vol = Csng((BallVel(ball)*0.3 + 4)/10)
End Function

Function Pan(ball) ' Calculates the pan for a ball based on the X position on the table. "table" is the name of the table
    Dim tmp
    tmp = ball.x * 2 / table1.width-1
    If tmp > 0 Then
        Pan = Csng(tmp ^10)
    Else
        Pan = Csng(-((- tmp) ^10) )
    End If
End Function

Function Pitch(ball) ' Calculates the pitch of the sound based on the ball speed
    Pitch = BallVel(ball) * 20
End Function

Function BallVel(ball) 'Calculates the ball speed
    BallVel = INT(SQR((ball.VelX ^2) + (ball.VelY ^2) ) )
End Function

'**********Sling Shot Animations
' Rstep and Lstep  are the variables that increment the animation
'****************
Dim RStep, Lstep, Tstep

Sub RightSlingShot_Slingshot
    PlaySoundAtVol SoundFXDOF("slingshotright",102,DOFPulse,DOFContactors), sling1, 1
    vpmTimer.PulseSw 15
    RSling.Visible = 0
    RSling1.Visible = 1
    sling1.TransZ = -20
    RStep = 0
    RightSlingShot.TimerEnabled = 1
End Sub

Sub RightSlingShot_Timer
    Select Case RStep
        Case 3:RSLing1.Visible = 0:RSLing2.Visible = 1:sling1.TransZ = -10
        Case 4:RSLing2.Visible = 0:RSLing.Visible = 1:sling1.TransZ = 0:RightSlingShot.TimerEnabled = 0
    End Select
    RStep = RStep + 1
End Sub

Sub LeftSlingShot_Slingshot
    PlaySoundAtVol SoundFXDOF("slingshotleft",101,DOFPulse,DOFContactors), sling2, 1
    vpmTimer.PulseSw 15
    LSling.Visible = 0
    LSling1.Visible = 1
    sling2.TransZ = -20
    LStep = 0
    LeftSlingShot.TimerEnabled = 1
End Sub

Sub LeftSlingShot_Timer
    Select Case LStep
        Case 3:LSLing1.Visible = 0:LSLing2.Visible = 1:sling2.TransZ = -10
        Case 4:LSLing2.Visible = 0:LSLing.Visible = 1:sling2.TransZ = 0:LeftSlingShot.TimerEnabled = 0
    End Select
    LStep = LStep + 1
End Sub

Sub TopSlingShot_Slingshot
    PlaySoundAtVol SoundFXDOF("left_slingshot",107,DOFPulse,DOFContactors), sling3, 1
    vpmTimer.PulseSw 15
    TSling.Visible = 0
    TSling1.Visible = 1
    sling3.TransZ = -20
    TStep = 0
    TopSlingShot.TimerEnabled = 1
End Sub

Sub TopSlingShot_Timer
    Select Case TStep
        Case 3:TSLing1.Visible = 0:TSLing2.Visible = 1:sling3.TransZ = -10
        Case 4:TSLing2.Visible = 0:TSLing.Visible = 1:sling3.TransZ = 0:TopSlingShot.TimerEnabled = 0
    End Select
    TStep = TStep + 1
End Sub

'*****************************************
'      JP's VP10 Rolling Sounds
'*****************************************

Const tnob = 6 ' total number of balls
ReDim rolling(tnob)
InitRolling

Sub InitRolling
    Dim i
    RollingTimer.interval=100
    RollingTimer.enabled=True
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
        If BallVel(BOT(b) ) > 1 AND BOT(b).z < 30 AND BOT(b).z > 0 Then
            rolling(b) = True
            PlaySound("fx_ballrolling" & b), -1, Vol(BOT(b) ), Pan(BOT(b) ), 0, Pitch(BOT(b) ), 1, 0, AudioFade(ActiveBall)
        Else
            If rolling(b) = True Then
                StopSound("fx_ballrolling" & b)
                rolling(b) = False
            End If
        End If
    Next
End Sub

'Gottlieb Volcano
'added by Inkochnito
'Added Coins chute by Mike da Spike
Sub editDips
  Dim vpmDips : Set vpmDips = New cvpmDips
  With vpmDips
    .AddForm 700,400,"Volcano - DIP switches"
    .AddFrame 2,5,190,"Coin Chute 1 (Coins/Credit)",&H0000000F,Array("1/2",&H00000008,"1/1",&H00000000,"2/1",&H00000009) 'Dip 1-4
    .AddFrame 2,67,190,"Coin Chute 2 (Coins/Credit)",&H000000F0,Array("1/2",&H00000080,"1/1",&H00000000,"2/1",&H00000090) 'Dip 5-8
    .AddFrame 2,129,190,"Coin Chute 3 (Coins/Credit)",&H00000F00,Array("1/2",&H00000800,"1/1",&H00000000,"2/1",&H00000900) 'Dip 9-12
    .AddFrame 2,191,190,"Coin Chute 3 extra credits",&H00001000,Array("no effect",0,"add 9",&H00001000)'dip 13
    .AddFrame 207,5,190,"Maximum credits",49152,Array("8 credits",0,"10 credits",32768,"15 credits",&H00004000,"25 credits",49152)'dip 15&16
    .AddFrame 207,81,190,"Coin chute 1 and 2 control",&H00002000,Array("seperate",0,"same",&H00002000)'dip 14
    .AddFrame 207,127,190,"Playfield special",&H00200000,Array("replay",0,"extra ball",&H00200000)'dip 22
    .AddFrame 207,173,190,"High game to date awards",&H00C00000,Array("not displayed and no award",0,"displayed and no award",&H00800000,"displayed and 2 credits",&H00400000,"displayed and 3 credits",&H00C00000)'dip 23&24
    .AddChk 2,245,190,Array("Match feature",&H00020000)'dip 18
    .AddChkExtra 2,260,190,Array("Speech",&H0020)'SS-board dip 6
    .AddChkExtra 2,275,190,Array("Background sound",&H0010)'SS-board dip 5
    .AddChk 2,290,190,Array("Must be on",&H01000000)'dip 25
    .AddChk 2,305,190,Array("Replay button tune?",&H02000000)'dip 26
    .AddFrameExtra 412,5,190,"Attract sound",&H000c,Array("off",0,"every 10 seconds",&H0004,"every 2 minutes",&H0008,"every 4 minutes",&H000C)'SS-board dip 3&4
    .AddFrame 412,81,190,"Balls per game",&H00010000,Array("5 balls",0,"3 balls",&H00010000)'dip 17
    .AddFrame 412,127,190,"Replay limit",&H00040000,Array("no limit",0,"one per game",&H00040000)'dip 19
    .AddFrame 412,173,190,"Novelty mode",&H00080000,Array("normal game mode",0,"50,000 points per special/extra ball",&H00080000)'dip 20
    .AddFrame 412,219,190,"Game mode",&H00100000,Array("replay",0,"extra ball",&H00100000)'dip 21
    .AddFrame 412,265,190,"Outlane extra ball one shot lite",&H40000000,Array("no extra ball through right outlane",0,"extra ball one shot light active",&H40000000)'dip 31
    .AddChk 207,275,190,Array("Coin switch tune?",&H04000000)'dip 27
    .AddChk 207,290,190,Array("Credits displayed?",&H08000000)'dip 28
    .AddChk 207,305,190,Array("Attract features",&H20000000)'dip 30
    .AddLabel 150,340,350,20,"After hitting OK, press F3 to reset game with new settings."
  End With
  Dim extra
  extra = Controller.Dip(4) + Controller.Dip(5)*256
  extra = vpmDips.ViewDipsExtra(extra)
  Controller.Dip(4) = extra And 255
  Controller.Dip(5) = (extra And 65280)\256 And 255
End Sub
Set vpmShowDips = GetRef("editDips")
'Digital LED Display

Dim Digits(28)
Digits(0)=Array(a00,a01,a02,a03,a04,a05,a06,n,a08)
Digits(1)=Array(a10,a11,a12,a13,a14,a15,a16,n,a18)
Digits(2)=Array(a20,a21,a22,a23,a24,a25,a26,n,a28)
Digits(3)=Array(a30,a31,a32,a33,a34,a35,a36,n,a38)
Digits(4)=Array(a40,a41,a42,a43,a44,a45,a46,n,a48)
Digits(5)=Array(a50,a51,a52,a53,a54,a55,a56,n,a58)
Digits(6)=Array(b00,b01,b02,b03,b04,b05,b06,n,b08)
Digits(7)=Array(b10,b11,b12,b13,b14,b15,b16,n,b18)
Digits(8)=Array(b20,b21,b22,b23,b24,b25,b26,n,b28)
Digits(9)=Array(b30,b31,b32,b33,b34,b35,b36,n,b38)
Digits(10)=Array(b40,b41,b42,b43,b44,b45,b46,n,b48)
Digits(11)=Array(b50,b51,b52,b53,b54,b55,b56,n,b58)
Digits(12)=Array(c00,c01,c02,c03,c04,c05,c06,n,c08)
Digits(13)=Array(c10,c11,c12,c13,c14,c15,c16,n,c18)
Digits(14)=Array(c20,c21,c22,c23,c24,c25,c26,n,c28)
Digits(15)=Array(c30,c31,c32,c33,c34,c35,c36,n,c38)
Digits(16)=Array(c40,c41,c42,c43,c44,c45,c46,n,c48)
Digits(17)=Array(c50,c51,c52,c53,c54,c55,c56,n,c58)
Digits(18)=Array(d00,d01,d02,d03,d04,d05,d06,n,d08)
Digits(19)=Array(d10,d11,d12,d13,d14,d15,d16,n,d18)
Digits(20)=Array(d20,d21,d22,d23,d24,d25,d26,n,d28)
Digits(21)=Array(d30,d31,d32,d33,d34,d35,d36,n,d38)
Digits(22)=Array(d40,d41,d42,d43,d44,d45,d46,n,d48)
Digits(23)=Array(d50,d51,d52,d53,d54,d55,d56,n,d58)
Digits(24)=Array(f00,f01,f02,f03,f04,f05,f06,n,f08)
Digits(25)=Array(f10,f11,f12,f13,f14,f15,f16,n,f18)

Digits(26)=Array(e00,e01,e02,e03,e04,e05,e06,n,e08)
Digits(27)=Array(e10,e11,e12,e13,e14,e15,e16,n,e18)




Sub DisplayTimer_Timer
  Dim ChgLED,ii,num,chg,stat,obj
  ChgLed = Controller.ChangedLEDs(&Hffffffff, &Hffffffff)
If Not IsEmpty(ChgLED) Then
    If DesktopMode = True Then
    For ii = 0 To UBound(chgLED)
      num = chgLED(ii, 0) : chg = chgLED(ii, 1) : stat = chgLED(ii, 2)
      if (num < 28) then
        For Each obj In Digits(num)
          If chg And 1 Then obj.State = stat And 1
          chg = chg\2 : stat = stat\2
        Next
      else

      end if
    next
    end if
end if
End Sub

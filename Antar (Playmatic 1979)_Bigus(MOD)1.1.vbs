
 '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
 '                                                               '
 '                              Antar                            '
 '                       (C) 1979 Playmatic                      '
 '                       VP9-Table by Mickey                     '
 '                                                               '
 '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

 Option Explicit

ExecuteGlobal GetTextFile("controller.vbs")
If Err Then MsgBox "You need the controller.vbs in order to run this table, available in the vp10 package"
On Error Goto 0

 LoadVPM "01560000", "play2.vbs", 3.32

if Table1.showdt = false then ramp1.visible = 0:ramp1a.visible = 0

 ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
 '                    Standard definitions                        '
 ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

 Const cGameName     = "antar2"
 Const UseSolenoids  = 1
 Const UseLamps      = 1
 Const UseGI         = 0

 ' Standard Sounds
 Const SSolenoidOn   = "Solenoid"
 Const SSolenoidOff  = ""
 Const SCoin         = "Coin"

 ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
 '                              Keys                              '
 ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

 Sub Table1_KeyDown(ByVal Keycode)
   If vpmKeyDown(keycode) Then Exit Sub
   If keycode = PlungerKey Then : Plunger1.Pullback: End If
 End Sub

 Sub Table1_KeyUp(ByVal Keycode)
     If vpmKeyUp(keycode) Then Exit Sub
     If keycode = PlungerKey Then : Plunger1.Fire: End If
 End Sub

 ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
 '                           Flippers                             '
 ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

 Sub SolLFlipper(Enabled)
     If Enabled AND bsTrough.balls <1 Then
         PlaySoundAtVol "flipperup", LeftFlipper, 1:LeftFlipper.RotateToEnd:UpperleftFlipper.RotateToEnd
     Else
     If bsTrough.balls >0 Then LeftFlipper.RotateToStart:UpperLeftFlipper.RotateToStart
     If bsTrough.balls <1 Then PlaySoundAtVol "flipperdown", LeftFlipper, 1:LeftFlipper.RotateToStart:UpperLeftFlipper.RotateToStart
     End If
 End Sub

 Sub SolRFlipper(Enabled)
     If Enabled AND bsTrough.balls <1 Then
         PlaySoundAtVol "flipperup", RightFlipper, 1:RightFlipper.RotateToEnd
     Else
     If bsTrough.balls >0 Then RightFlipper.RotateToStart
     If bsTrough.balls <1 Then PlaySoundAtVol "flipperdown", ActiveBall, 1:RightFlipper.RotateToStart
     End If
 End Sub

 ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
 '                      Table definitions                         '
 ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
 ' Solenoid 3 is OFF if buyin>15 (standard) -> no more coins will be accepted

 Const sSaucer = 1     'OK
 Const sResetDrop = 4  'OK
 Const sOutHole = 5    'OK
 Const sKnocker = 8    'OK

 SolCallback(sOutHole) = "bsTrough.SolOut"
 SolCallback(sResetDrop) = "dtL.SolDropUp"
 SolCallback(sSaucer) = "SolKickers"
 'SolCallback(sKnocker) = "vpmSolSound ""knocker""," 'disable if you don't want knocker sound at beginning of each game
 SolCallback(sLRFlipper) = "SolRFlipper"
 SolCallback(sLLFlipper) = "SolLFlipper"

 Dim bsTrough,bsSaucerR,bsSaucerR1,bsSaucerL,bsSaucerL1,dtL,Bump,plungerIM



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'                         Table init                             '
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

 Sub Table1_Init

     vpmInit me
     With Controller
         .GameName = cGameName
         If Err Then MsgBox "Can't start Game" & cGameName & vbNewLine & Err.Description:Exit Sub
         .SplashInfoLine = "Antar - Playmatic 1979"
         .HandleKeyboard = 0
         .ShowTitle = 0
         .ShowDMDOnly = 1
         .ShowFrame = 0
         .HandleMechanics = 0
         .hidden = 0
'         .Games(cGameName).Settings.Value("rol")=0
         '.SetDisplayPosition 0, 0, GetPlayerHWnd
         On Error Resume Next
         .Run GetPlayerHWnd
         If Err Then MsgBox Err.Description
         On Error Goto 0
     End With


 Controller.Games(cGameName).Settings.Value("volume") = -8.0


   ' Nudge init
   With vpmNudge
       .TiltSwitch = swTilt
       .Sensitivity = 5
     .TiltObj=Array(SW36_Slingshot,SW44_Slingshot,SW45_Slingshot,SW46_Slingshot,SW35_Bumper,SW8_Trigger,LeftFlipper,RightFlipper)
   End With

     ' Trough init
     Set bsTrough = New cvpmBallStack
     With bsTrough
         .InitSw 0, 5, 0, 0, 0, 0, 0, 0
         .InitKick BallRelease, 80, 6
         .InitEntrySnd "solenoid", "solenoid"
         .InitExitSnd "ballrel", "solenoid"
         .Balls = 1
     End With

   ' Saucers init
     Set bsSaucerR = New cvpmBallStack
   With bsSaucerR
     .InitSaucer SW32_Kicker, 32, 10, 15
         .InitExitSnd "popper","popper"
     .Balls = 0
     End With

      Set bsSaucerR1 = New cvpmBallStack
    With bsSaucerR1
         .InitSaucer kicker1, 37, 337, 24
        .InitExitSnd "popper","popper"
      .Balls = 0
     End With


     Set bsSaucerL = New cvpmBallStack
   With bsSaucerL
     .InitSaucer SW38_Kicker, 38, 10, 20 'if speed is to high, SW32 cannot initialize and ball will stay in kicker...
       .InitExitSnd "popper","popper"
     .Balls = 0
     End With

     Set bsSaucerL1 = New cvpmBallStack
   With bsSaucerL1
         .InitSaucer kicker2, 31, 20, 28
        .InitExitSnd "popper","popper"
      .Balls = 0
     End With

   ' Droptargets init
   Set dtL = New cvpmDropTarget
   dtL.InitDrop Array(SW11_Droptarget,SW12_Droptarget,SW13_Droptarget,SW14_Droptarget,SW15_Droptarget,SW16_Droptarget,SW17_Droptarget,SW18_Droptarget,SW21_Droptarget,SW22_Droptarget),Array(11,12,13,14,15,16,17,18,21,22)
   dtL.initsnd "fx_droptarget","resetdrop"

     ' Impulse Plunger init
   Const IMPowerSetting = 38    ' Plunger power
   Const IMTime     = 0.5   ' Time in seconds for full plunge
   'Set plungerIM = New cvpmImpulseP : With plungerIM
    '.InitImpulseP IMPlunger, IMPowerSetting, IMTime
    '.Random 0.3
    '.InitExitSnd  "plunger2", "plunger"
    '.CreateEvents "plungerIM"
   'End With

     ' Main Timer init
     PinMAMETimer.Interval = PinMAMEInterval
     PinMAMETimer.Enabled = 1

     ' Other Timers init
   Kicker1.TimerInterval = 1500
   Kicker2.Timerinterval = 1500
   TriggerGate1_open.Timerinterval = 1250
   TriggerGate2_open.Timerinterval = 1000

   ' Animation init
     Gate1_active.IsDropped = 1:Gate2_active.IsDropped = 1
   Kicklever1_active.Isdropped = 1:Kicklever2_active.IsDropped = 1
  vpmMapLights AllLights
 End Sub

 Sub Table1_Paused:Controller.Pause = 1:End Sub
 Sub Table1_unPaused:Controller.Pause = 0:End Sub

 ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
 '                        Switchhandling                          '
 ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Sub SolKickers(Enabled)
If enabled Then
    bsSaucerR.ExitSol_On
    bsSaucerR1.ExitSol_On
    bsSaucerL.ExitSol_On
    bsSaucerL1.ExitSol_On
End If
End Sub


   'Drain
  Sub Drain_Hit:PlaySoundAtVol "drain", Drain, 1:bsTrough.AddBall Me:vpmtimer.addtimer 1000,"SW23_Target.isdropped = 0'":Controller.Switch(23)=0:Light5c.state = 0:End Sub
  Sub SW8_Trigger_Hit:Controller.Switch(8)=1:PlaySoundAtVol "rollover", ActiveBall, 1:End Sub
  Sub SW8_Trigger_unHit:Controller.Switch(8)=0:End Sub
   'Gates
  Sub TriggerGate1_open_Hit:Gate1_norm.IsDropped = 1:Gate1_active.IsDropped = 0:StopSound "gate":Me.TimerEnabled = 1:PlaySoundAtVol "gate", ActiveBall, 1:End Sub
  Sub TriggerGate1_open_Timer:Gate1_norm.IsDropped = 0:Gate1_active.Isdropped = 1:Me.TimerEnabled = 0:End Sub
  Sub TriggerGate1_close_Hit:Gate1_norm.IsDropped = 0:Gate1_active.Isdropped = 1:End Sub
  Sub TriggerGate2_open_Hit:Gate2_norm.IsDropped = 1:Gate2_active.IsDropped = 0:StopSound "gate":Me.TimerEnabled = 1:PlaySoundAtVol "gate", ActiveBall, 1:End Sub
  Sub TriggerGate2_open_Timer:Gate2_norm.IsDropped = 0:Gate2_active.Isdropped = 1:Me.TimerEnabled = 0:End Sub
  Sub TriggerGate2_close_Hit:Gate2_norm.IsDropped = 0:Gate2_active.Isdropped = 1:End Sub
   'Slingshots
  Sub SW36_Slingshot_Slingshot:VpmTimer.PulseSw 36:PlaySoundAtVol "slingshot", ActiveBall, 1:End Sub
  Sub SW44_Slingshot_Slingshot:VpmTimer.PulseSw 44:PlaySoundAtVol "slingshot", ActiveBall, 1:End Sub
  Sub SW45_Slingshot_Slingshot:VpmTimer.PulseSw 45:PlaySoundAtVol "slingshot", ActiveBall, 1:End Sub
  Sub SW46_Slingshot_Slingshot:VpmTimer.PulseSw 46:PlaySoundAtVol "slingshot", ActiveBall, 1:End Sub
   'Bumper
  Sub SW35_Bumper_Hit:vpmTimer.PulseSw 35:PlaySoundAtVol "bumper1", ActiveBall, 1:End Sub
     'Spinner
  Sub Sw33_Spinner_Spin:vpmTimer.PulseSw 33:PlaySoundAtVol"spinner", ActiveBall, 1:End Sub
   'Saucers
  Sub Kicker1_Timer:Kicklever1_norm.IsDropped = 0:Kicklever1_active.IsDropped = 1:Me.TimerEnabled = 0:End Sub'
  Sub Kicker2_Timer:Kicklever2_norm.IsDropped = 0:Kicklever2_active.IsDropped = 1:Me.TimerEnabled = 0:End Sub
  Sub KickerTimer1_Timer:Kicklever1_norm.IsDropped = 1:Kicklever1_active.IsDropped = 0:KickerTimer1.enabled = 0:End Sub
  Sub KickerTimer2_Timer:Kicklever2_norm.IsDropped = 1:Kicklever2_active.IsDropped = 0:KickerTimer2.enabled = 0:End Sub
  Sub SW32_Kicker_Hit:bsSaucerR.AddBall 0:Controller.Switch(32) = 0:SolCallback(sSaucer) = "bsSaucerR.SolOut":End Sub 'ROM do not call SolCallback by now
  Sub SW38_Kicker_Hit:bsSaucerL.AddBall 0:Controller.Switch(38) = 0:SolCallback(sSaucer) = "bsSaucerL.SolOut":End Sub
  Sub Kicker1_Hit:bsSaucerR1.AddBall 0:Controller.Switch(37) = 0:SolCallback(sSaucer) = "bsSaucerR1.SolOut":End Sub
  Sub Kicker2_Hit:bsSaucerL1.AddBall 0:Controller.Switch(31) = 0:SolCallback(sSaucer) = "bsSaucerL1.SolOut":End Sub

   'Droptargets
  Sub SW11_Droptarget_Hit:dtL.Hit 1:End Sub
  Sub SW12_Droptarget_Hit:dtL.Hit 2:End Sub
  Sub SW13_Droptarget_Hit:dtL.Hit 3:End Sub
  Sub SW14_Droptarget_Hit:dtL.Hit 4:End Sub
  Sub SW15_Droptarget_Hit:dtL.Hit 5:End Sub
  Sub SW16_Droptarget_Hit:dtL.Hit 6:End Sub
  Sub SW17_Droptarget_Hit:dtL.Hit 7:End Sub
  Sub SW18_Droptarget_Hit:dtL.Hit 8:End Sub
  Sub SW21_Droptarget_Hit:dtL.Hit 9:End Sub
  Sub SW22_Droptarget_Hit:dtL.Hit 10:End Sub
   'Rollovers
  Sub SW28_Trigger_Hit:Controller.Switch(28)=1:PlaySoundAtVol "rollover", ActiveBall, 1:End Sub
  Sub SW28_Trigger_UnHit:Controller.Switch(28)=0:End Sub
  Sub SW31_Trigger_Hit:Controller.Switch(31)=1:KickerTimer2.enabled = 1:Kicker2.TimerEnabled = 1:vpmtimer.addtimer 1000,"SW23_Target.isdropped = 0'":vpmtimer.addtimer 1000,"Controller.Switch(23)=0'":vpmtimer.addtimer 1000,"Light5c.state = 0'":End Sub
  Sub SW31_Trigger_UnHit:Controller.Switch(31)=0:End Sub
  Sub SW34_Trigger_Hit:Controller.Switch(34)=1:PlaySoundAtVol "rollover", ActiveBall, 1:End Sub
  Sub SW34_Trigger_UnHit:Controller.Switch(34)=0:End Sub
  Sub SW37_Trigger_Hit:Controller.Switch(37)=1:KickerTimer1.enabled = 1:Kicker1.TimerEnabled = 1:vpmtimer.addtimer 1000,"SW23_Target.isdropped = 0'":vpmtimer.addtimer 1000,"Controller.Switch(23)=0'":vpmtimer.addtimer 1000,"Light5c.state = 0'":End Sub
  Sub SW37_Trigger_UnHit:Controller.Switch(37)=0:End Sub
  Sub SW41_Trigger_Hit:Controller.Switch(41)=1:PlaySoundAtVol "rollover", ActiveBall, 1:End Sub
  Sub SW41_Trigger_UnHit:Controller.Switch(41)=0:End Sub
  Sub SW42_Trigger_Hit:Controller.Switch(42)=1:PlaySoundAtVol "rollover", ActiveBall, 1:End Sub
  Sub SW42_Trigger_UnHit:Controller.Switch(42)=0:End Sub
  Sub SW43_Trigger_Hit:Controller.Switch(43)=1:PlaySoundAtVol "rollover", ActiveBall, 1:End Sub
  Sub SW43_Trigger_UnHit:Controller.Switch(43)=0:End Sub
   'Targets
  Sub SW23_Target_Hit:Controller.Switch(23)=1:SW23_Target.isdropped = 1:PlaySoundAtVol"fx_droptarget", ActiveBall, 1:Light5c.state = 1:End Sub
  Sub SW24_Target_Hit:VpmTimer.PulseSw 24:PlaySoundAtVol"target", ActiveBall, 1:End Sub
  Sub SW25_Target_Hit:VpmTimer.PulseSw 25:PlaySoundAtVol"target", ActiveBall, 1:End Sub
  Sub SW26_Target_Hit:VpmTimer.PulseSw 26:PlaySoundAtVol"target", ActiveBall, 1:End Sub
  Sub SW27_Target_Hit:VpmTimer.PulseSw 27:PlaySoundAtVol"target", ActiveBall, 1:End Sub

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


 '**********************
'Flipper Shadows
'***********************
Sub RealTime_Timer
  lfs.RotZ = LeftFlipper.CurrentAngle
  rfs.RotZ = RightFlipper.CurrentAngle
BallShadowUpdate
End Sub


Sub BallShadowUpdate()
Dim BallShadow
BallShadow = Array (BallShadow1,BallShadow2)
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
    BallShadow(b).X = BOT(b).X
    ballShadow(b).Y = BOT(b).Y + 10
        If BOT(b).Z > 20 and BOT(b).Z < 200 Then
            BallShadow(b).visible = 1
        Else
            BallShadow(b).visible = 0
        End If
if BOT(b).z > 30 Then
ballShadow(b).height = BOT(b).Z - 20
ballShadow(b).opacity = 80
Else
ballShadow(b).height = BOT(b).Z - 24
ballShadow(b).opacity = 90
End If
    Next
End Sub



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
  PlaySound "metalhit_medium", 0, Vol(ActiveBall), AudioPan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
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
   If finalspeed > 20 then
    PlaySound "fx_rubber2", 0, Vol(ActiveBall), AudioPan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
 End if
  If finalspeed >= 6 AND finalspeed <= 20 then
    RandomSoundRubber()
   End If
End Sub

Sub Posts_Hit(idx)
   dim finalspeed
    finalspeed=SQR(activeball.velx * activeball.velx + activeball.vely * activeball.vely)
   If finalspeed > 16 then
    PlaySound "fx_rubber2", 0, Vol(ActiveBall), AudioPan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
 End if
  If finalspeed >= 6 AND finalspeed <= 16 then
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

' Thalamus : Exit in a clean and proper way
Sub Table1_exit
  Controller.Pause = False
  Controller.Stop
End Sub


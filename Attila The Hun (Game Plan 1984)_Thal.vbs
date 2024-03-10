Option Explicit
On Error Resume Next

ExecuteGlobal GetTextFile("controller.vbs")
If Err Then MsgBox "You need the Controller.vbs file in order to run this table (installed with the VPX package in the scripts folder)"
On Error Goto 0
LoadVPM "01120100","GamePlan.vbs",3.1

Dim cCredits
cCredits="Attila The Hun" & vbNewline & "GamePlan 1984"
Const cGameName="attila",UseSolenoids=1,UseLamps=1,UseGI=0,UseSync=1,SCoin="coin3"

Dim EnableBallControl
EnableBallControl = False 'Change to true to enable manual ball control (or press C in-game) via the arrow keys and B (boost movement) keys

BallMass = 1
BallSize = 50

SolCallback(8)="bsTrough.SolOut"
SolCallback(11)="dtL.SolDropUp"
SolCallback(12)="dtR.SolDropUp"
SolCallback(3)="vpmSolSound ""Knocker"","
SolCallback(16)="vpmNudge.SolGameOn"
SolCallback(sLRFlipper) = "SolRFlipper"
SolCallback(sLLFlipper) = "SolLFlipper"

Sub SolLFlipper(Enabled)
     If Enabled Then
         PlaySoundAtVol SoundFX("fx_Flipperup",DOFContactors), LeftFlipper, 1:LeftFlipper.RotateToEnd
     Else
         PlaySoundAtVol SoundFX("fx_Flipperdown",DOFContactors), LeftFlipper, 1:LeftFlipper.RotateToStart
     End If
  End Sub

Sub SolRFlipper(Enabled)
     If Enabled Then
         PlaySoundAtVol SoundFX("fx_Flipperup",DOFContactors), RightFlipper, 1:RightFlipper.RotateToEnd:RightFlipper.RotateToEnd
     Else
         PlaySoundAtVol SoundFX("fx_Flipperdown",DOFContactors), RightFlipper, 1:RightFlipper.RotateToStart:RightFlipper.RotateToStart
     End If
End Sub

Dim bsTrough,bsSaucer,dtR,dtL,cbCaptive,Objekt

Sub Table1_Init

' Thalamus : Was missing 'vpminit me'
  vpminit me

  On Error Resume Next
  With Controller
    .GameName=cGameName
    If Err Then MsgBox "Can't start Game" & cGameName & vbNewLine & Err.Description:Exit Sub
    .SplashInfoLine=cCredits
    .HandleMechanics=0
    .ShowDMDOnly=1
    .ShowFrame=0
    .ShowTitle=0
    .Hidden=0
    .Run
    If Err Then MsgBox Err.Description
  End With
  On Error Goto 0

  'F6 to access in game.  (23 on for 5 balls, off for 3)
  Controller.Dip(0) = (0*1 + 1*2 + 0*4 + 0*8 + 0*16 + 0*32 + 0*64 + 0*128) '01-08
  Controller.Dip(1) = (0*1 + 0*2 + 0*4 + 1*8 + 1*16 + 1*32 + 1*64 + 1*128) '09-16
  Controller.Dip(2) = (0*1 + 1*2 + 0*4 + 0*8 + 0*16 + 0*32 + 0*64 + 1*128) '17-24
  Controller.Dip(3) = (0*1 + 1*2 + 1*4 + 0*8 + 0*16 + 1*32 + 1*64 + 0*128) '25-32

  PinMAMETimer.Interval=PinMAMEInterval
  PinMAMETimer.Enabled=True
  vpmMapLights aLights

  vpmNudge.TiltSwitch=swTilt
  vpmNudge.Sensitivity=5
  vpmNudge.TiltObj=Array(Bumper1,Bumper2,Bumper3,LeftSlingshot,RightSlingshot)
  Kicker1.createball
  Kicker1.kick 155,10

  Set bsTrough=New cvpmBallStack
  bsTrough.InitSw 0,11,0,0,0,0,0,0
  bsTrough.InitKick BallRelease,60,3
  bsTrough.InitExitSnd "fx_ballrel","solon"
  bsTrough.Balls=1

  Set dtL=New cvpmDropTarget
  dtL.InitDrop Array(Target1,Target2,Target3),Array(27,28,29)
  dtL.InitSnd "target","fx_resetdrop"

  Set dtR=New cvpmDropTarget
  dtR.InitDrop Array(Target4,Target5,Target6),Array(30,31,32)
  dtR.InitSnd "target","fx_resetdrop"

  If Table1.ShowDT = False then
  for each objekt in backdrop:objekt.visible = False:Next
  End If

End Sub

Sub Table1_KeyDown(ByVal KeyCode)
    If vpmKeyDown(KeyCode) Then Exit Sub
    If KeyCode=PlungerKey Then playsoundAtVol "fx_plungerpull", Plunger, 1:Plunger.Pullback:End if

' Manual Ball Control
  If keycode = 46 Then          ' C Key
    If EnableBallControl = 1 Then
      EnableBallControl = 0
    Else
      EnableBallControl = 1
    End If
  End If
    If EnableBallControl = 1 Then
    If keycode = 48 Then        ' B Key
      If BCboost = 1 Then
        BCboost = BCboostmulti
      Else
        BCboost = 1
      End If
    End If
    If keycode = 203 Then BCleft = 1  ' Left Arrow
    If keycode = 200 Then BCup = 1    ' Up Arrow
    If keycode = 208 Then BCdown = 1  ' Down Arrow
    If keycode = 205 Then BCright = 1 ' Right Arrow
  End If
End Sub


Sub Table1_KeyUp(ByVal KeyCode)
    If vpmKeyUp(KeyCode) Then Exit Sub
    If KeyCode=PlungerKey Then PlaySoundAtVol"fx_plunger", Plunger, 1:Plunger.Fire

'Manual Ball Control
  If EnableBallControl = 1 Then
    If keycode = 203 Then BCleft = 0  ' Left Arrow
    If keycode = 200 Then BCup = 0    ' Up Arrow
    If keycode = 208 Then BCdown = 0  ' Down Arrow
    If keycode = 205 Then BCright = 0 ' Right Arrow
  End If
End Sub


Sub Spinner1_Spin:vpmTimer.PulseSw 4:PlaySound "fx_spinner", 0, .25, AudioPan(Spinner1), 0.25, 0, 0, 1, AudioFade(Spinner1):End Sub
Sub Drain_Hit:Controller.Switch(11)=1:PlaySound SoundFX("drain",DOFContactors),0,1,AudioPan(Drain),0.25,0,0,1,AudioFade(Drain):bsTrough.AddBall Me:End Sub
Sub Trigger2_Hit:Controller.Switch(12)=1:End Sub
Sub Trigger2_unHit:Controller.Switch(12)=0:End Sub
Sub Trigger3_Hit:Controller.Switch(13)=1:End Sub
Sub Trigger3_unHit:Controller.Switch(13)=0:End Sub
Sub Trigger4_Hit:Controller.Switch(14)=1:End Sub
Sub Trigger4_unHit:Controller.Switch(14)=0:End Sub
Sub Trigger1_Hit:Controller.Switch(15)=1:End Sub
Sub Trigger1_unHit:Controller.Switch(15)=0:End Sub
Sub Trigger5_Hit:Controller.Switch(16)=1:End Sub
Sub Trigger5_unHit:Controller.Switch(16)=0:End Sub
Sub Trigger6_Hit:Controller.Switch(17)=1:End Sub
Sub Trigger6_unHit:Controller.Switch(17)=0:End Sub
Sub LeftOutlane_Hit:Controller.Switch(18)=1:End Sub
Sub LeftOutlane_unHit:Controller.Switch(18)=0:End Sub
Sub LeftInlane_Hit:Controller.Switch(19)=1:End Sub
Sub LeftInlane_unHit:Controller.Switch(19)=0:End Sub
Sub RightInlane_Hit:Controller.Switch(20)=1:End Sub
Sub RightInlane_unHit:Controller.Switch(20)=0:End Sub
Sub RightOutlane_Hit:Controller.Switch(21)=1:End Sub
Sub RightOutlane_unHit:Controller.Switch(21)=0:End Sub
Sub Bumper1_Hit:vpmTimer.PulseSw 22:PlaySound SoundFX("fx_bumper4",DOFContactors), 0,1,AudioPan(Bumper1),0,0,0,1,AudioFade(Bumper1):End Sub
Sub Bumper3_Hit:vpmTimer.PulseSw 23:PlaySound SoundFX("fx_bumper3",DOFContactors), 0,1,AudioPan(Bumper3),0,0,0,1,AudioFade(Bumper3):End Sub
Sub Bumper2_Hit:vpmTimer.PulseSw 24:PlaySound SoundFX("fx_bumper2",DOFContactors), 0,1,AudioPan(Bumper2),0,0,0,1,AudioFade(Bumper2):End Sub
Sub Wall4_Hit:vpmTimer.PulseSw 25:End Sub
Sub Wall26_Hit:vpmTimer.PulseSw 25:End Sub
Sub Wall14_Hit:vpmTimer.PulseSw 25:End Sub
Sub Wall16_Hit:vpmTimer.PulseSw 25:End Sub
Sub Wall18_Hit:vpmTimer.PulseSw 25:End Sub
Sub Wall19_Hit:vpmTimer.PulseSw 25:End Sub
Sub Wall20_Hit:vpmTimer.PulseSw 25:End Sub
Sub Wall21_Hit:vpmTimer.PulseSw 25:End Sub
Sub Wall24_Hit:vpmTimer.PulseSw 25:End Sub
Sub Target1_Hit:dtL.Hit 1:End Sub
Sub Target2_Hit:dtL.Hit 2:End Sub
Sub Target3_Hit:dtL.Hit 3:End Sub
Sub Target4_Hit:dtR.Hit 1:End Sub
Sub Target5_Hit:dtR.Hit 2:End Sub
Sub Target6_Hit:dtR.Hit 3:End Sub
Sub Table1_Paused:Controller.Pause=True:End Sub
Sub Table1_UnPaused:Controller.Pause=False:End Sub
Sub Table1_Exit:Controller.Stop:End Sub

'**********Sling Shot Animations
' Rstep and Lstep  are the variables that increment the animation
'****************
Dim RStep, Lstep

Sub RightSlingShot_Slingshot: vpmTimer.PulseSw 10
    PlaySoundAtVol SoundFX("right_slingshot",DOFContactors), sling1, 1
    RSling.Visible = 0
    RSling1.Visible = 1
    sling1.TransZ = -20
    RStep = 0
    RightSlingShot.TimerEnabled = 1
  gi1.State = 0:Gi2.State = 0
End Sub

Sub RightSlingShot_Timer
    Select Case RStep
        Case 3:RSLing1.Visible = 0:RSLing2.Visible = 1:sling1.TransZ = -10
        Case 4:RSLing2.Visible = 0:RSLing.Visible = 1:sling1.TransZ = 0:RightSlingShot.TimerEnabled = 0:gi1.State = 1:Gi2.State = 1
    End Select
    RStep = RStep + 1
End Sub

Sub LeftSlingShot_Slingshot:vpmTimer.PulseSw 9
    PlaySoundAtVol SoundFX("left_slingshot",DOFContactors), sling2, 1
    LSling.Visible = 0
    LSling1.Visible = 1
    sling2.TransZ = -20
    LStep = 0
    LeftSlingShot.TimerEnabled = 1
  gi3.State = 0:Gi4.State = 0
End Sub

Sub LeftSlingShot_Timer
    Select Case LStep
        Case 3:LSLing1.Visible = 0:LSLing2.Visible = 1:sling2.TransZ = -10
        Case 4:LSLing2.Visible = 0:LSLing.Visible = 1:sling2.TransZ = 0:LeftSlingShot.TimerEnabled = 0:gi3.State = 1:Gi4.State = 1
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


' '*****************************************
' '   rothbauerw's Manual Ball Control
' '*****************************************
'
' Dim BCup, BCdown, BCleft, BCright
' Dim ControlBallInPlay, ControlActiveBall
' Dim BCvel, BCyveloffset, BCboostmulti, BCboost
'
' BCboost = 1       'Do Not Change - default setting
' BCvel = 4       'Controls the speed of the ball movement
' BCyveloffset = -0.01  'Offsets the force of gravity to keep the ball from drifting vertically on the table, should be negative
' BCboostmulti = 3    'Boost multiplier to ball veloctiy (toggled with the B key)
'
' ControlBallInPlay = false
'
' Sub StartBallControl_Hit()
'   Set ControlActiveBall = ActiveBall
'   ControlBallInPlay = true
' End Sub
'
' Sub StopBallControl_Hit()
'   ControlBallInPlay = false
' End Sub
'
' Sub BallControlTimer_Timer()
'   If EnableBallControl and ControlBallInPlay then
'     If BCright = 1 Then
'       ControlActiveBall.velx =  BCvel*BCboost
'     ElseIf BCleft = 1 Then
'       ControlActiveBall.velx = -BCvel*BCboost
'     Else
'       ControlActiveBall.velx = 0
'     End If
'
'     If BCup = 1 Then
'       ControlActiveBall.vely = -BCvel*BCboost
'     ElseIf BCdown = 1 Then
'       ControlActiveBall.vely =  BCvel*BCboost
'     Else
'       ControlActiveBall.vely = bcyveloffset
'     End If
'   End If
' End Sub

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
      If BOT(b).VelZ < -1 and BOT(b).z < 55 and BOT(b).z > 27 Then 'height adjust for ball drop sounds
        PlaySoundAtBOTBallZ "fx_ball_drop" & b, BOT(b)
        'debug.print BOT(b).velz
      End If
    Next
End Sub

'**********************
' Ball Collision Sound
'**********************

Sub OnBallBallCollision(ball1, ball2, velocity)
    PlaySound("fx_collide"), 0, Csng(velocity) ^2 / 2000, AudioPan(ball1), 0, Pitch(ball1), 0, 0, AudioFade(ball1)
End Sub


'*****************************************
' ninuzzu's FLIPPER SHADOWS
'*****************************************

sub FlipperTimer_Timer()
  FlipperLSh.RotZ = LeftFlipper.currentangle
  FlipperRSh.RotZ = RightFlipper.currentangle

End Sub

'*****************************************
' ninuzzu's BALL SHADOW
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
        If BOT(b).X < Table1.Width/2 Then
            BallShadow(b).X = ((BOT(b).X) - (Ballsize/6) + ((BOT(b).X - (Table1.Width/2))/21)) + 6
        Else
            BallShadow(b).X = ((BOT(b).X) + (Ballsize/6) + ((BOT(b).X - (Table1.Width/2))/21)) - 6
        End If
        ballShadow(b).Y = BOT(b).Y + 4
        If BOT(b).Z > 20 Then
            BallShadow(b).visible = 1
        Else
            BallShadow(b).visible = 0
        End If
    Next
End Sub

Sub aGates_Hit(idx):PlaySound "fx_Gate", 0, Vol(ActiveBall), Audiopan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall):End Sub
Sub aMetals_Hit(idx):PlaySound "fx_MetalHit", 0, Vol(ActiveBall), Audiopan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall):End Sub

Sub aRubbers_Hit(idx)
  dim finalspeed
    finalspeed=SQR(activeball.velx * activeball.velx + activeball.vely * activeball.vely)
  If finalspeed > 20 then
    PlaySound "fx_rubber_band", 0, Vol(ActiveBall), AudioPan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
  End if
  If finalspeed >= 6 AND finalspeed <= 20 then
    RandomSoundRubber()
  End If
End Sub

Sub aPosts_Hit(idx)
  dim finalspeed
    finalspeed=SQR(activeball.velx * activeball.velx + activeball.vely * activeball.vely)
  If finalspeed > 16 then
    PlaySound "fx_postrubber", 0, Vol(ActiveBall), AudioPan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
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
  dim finalspeed
    finalspeed=SQR(activeball.velx * activeball.velx + activeball.vely * activeball.vely)
  If finalspeed > 4 then
    RandomSoundFlipper()
  Else
    RandomSoundFlipperLowVolume()
  End If

End Sub

Sub RightFlipper_Collide(parm)
  dim finalspeed
    finalspeed=SQR(activeball.velx * activeball.velx + activeball.vely * activeball.vely)
  If finalspeed > 4 then
    RandomSoundFlipper()
  Else
    RandomSoundFlipperLowVolume()
  End If
End Sub
Sub RandomSoundFlipper()
  Select Case Int(Rnd*3)+1
    Case 1 : PlaySoundAtBall "fx2_flip_hit_1"
    Case 2 : PlaySoundAtBall "fx2_flip_hit_2"
    Case 3 : PlaySoundAtBall "fx2_flip_hit_3"
  End Select
End Sub

Sub RandomSoundFlipperLowVolume()
  Select Case Int(Rnd*3)+1
    Case 1 : PlaySoundAtBall "fx2_flip_hit_1_low"
    Case 2 : PlaySoundAtBall "fx2_flip_hit_2_low"
    Case 3 : PlaySoundAtBall "fx2_flip_hit_3_low"
  End Select
End Sub

'************************************
'          LEDs Display
'************************************
Dim Digits(32)
Dim Patterns(11)
Dim Patterns2(11)

Patterns(0) = 0     'empty
Patterns(1) = 63    '0
Patterns(2) = 6     '1
Patterns(3) = 91    '2
Patterns(4) = 79    '3
Patterns(5) = 102   '4
Patterns(6) = 109   '5
Patterns(7) = 124   '6
Patterns(8) = 7     '7
Patterns(9) = 127   '8
Patterns(10) = 103  '9

Patterns2(0) = 128  'empty
Patterns2(1) = 191  '0
Patterns2(2) = 134  '1
Patterns2(3) = 219  '2
Patterns2(4) = 207  '3
Patterns2(5) = 230  '4
Patterns2(6) = 237  '5
Patterns2(7) = 252  '6
Patterns2(8) = 135  '7
Patterns2(9) = 255  '8
Patterns2(10) = 239 '9

Set Digits(0) = a0
Set Digits(1) = a1
Set Digits(2) = a2
Set Digits(3) = a3
Set Digits(4) = a4
Set Digits(5) = a5

Set Digits(6) = b0
Set Digits(7) = b1
Set Digits(8) = b2
Set Digits(9) = b3
Set Digits(10) = b4
Set Digits(11) = b5

Set Digits(12) = c0
Set Digits(13) = c1
Set Digits(14) = c2
Set Digits(15) = c3
Set Digits(16) = c4
Set Digits(17) = c5

Set Digits(18) = d0
Set Digits(19) = d1
Set Digits(20) = d2
Set Digits(21) = d3
Set Digits(22) = d4
Set Digits(23) = d5

Set Digits(24) = e0
Set Digits(25) = e1
Set Digits(26) = e2
Set Digits(27) = e3
Set Digits(28) = e4
Set Digits(29) = e5
Set Digits(30) = e6
Set Digits(31) = e7


Sub Leds_Timer()
     On Error Resume Next
    Dim ChgLED, ii, jj, chg, stat
    ChgLED = Controller.ChangedLEDs(&Hffffffff, &Hffffffff)
    If Not IsEmpty(ChgLED)Then
        For ii = 0 To UBound(ChgLED)
            chg = chgLED(ii, 1):stat = chgLED(ii, 2)
            For jj = 0 to 10
                If stat = Patterns(jj)OR stat = Patterns2(jj)then Digits(chgLED(ii, 0)).SetValue jj
            Next
        Next
    End IF
End Sub

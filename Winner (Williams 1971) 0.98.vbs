' ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' Willimas' "Winner" (1972)
' Created by: CactusDude
' Made in 2021
' Version: 0.97
'
' Thanks to
' BorgDog & team for creating Hi-Score Pool, which served as the jumping off point
' bord for using his Blender magic on the backglass images
' Pmax65 for helping squash bugs in in the table & providing code for a better horse reset animation
' ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

Option Explicit
Randomize

On Error Resume Next
ExecuteGlobal GetTextFile("controller.vbs")
If Err Then MsgBox "Can't open controller.vbs"
On Error Goto 0

Const cGameName = "WilliamsWinner"

Const Purse = 10000 'Points awarded for winning the race

Dim BIP, Balls, gLight
Dim PlayerScores(2), SubScores(3), ScoresRoll, SubPlayer
Dim Player, Players, Credits
Dim j, GameMode, Objekt, wdown
Dim BallsLit, PlayBall, D, R
Dim Rack, Racks, Tilt, tiltsens
Dim StarLit
Dim motorsound, moveballs, KnockerOn
Dim ballangle(15), direction(15)
Dim BIPonBG
Dim hiscore

Const BallSize = 22




Sub Table1_Init()
    LoadEM
  Players = 0
  D = 0 'for Disc Degree
  R = -1
  loadcredits
  EndScoresTimer.Enabled = false
  If Credits="" or Credits<0 or Credits>25 then Credits=0
  If Hiscore="" then hiscore=1000
  If B2SOn Then
    Backglasstimer.enabled=1
    for each objekt in BackDropStuff: Objekt.visible=0:Next
  end if
  EMReel3.SetValue Credits
  GOReel.setvalue 1
  EMReel1.Image = "NumbersDark"
  EMReel2.Image = "NumbersDark"
  Tilt=False

  For each objekt in RotWalls: objekt.isdropped=True: Next

  rotWall0.IsDropped = false

  If Credits > 0 Then DOF 103, DOFOn
  If Credits > 0 Then CreditLight.State = 1 'should be 1
  If Credits = 0 Then CreditLight.State = 0 'should be 0

  ResetGame
End Sub


sub backglasstimer_timer
  Controller.B2SSetCredits Credits
  Controller.B2SSetGameover 1
  Controller.B2SSetScorePlayer 5, hiscore
  Controller.B2SSetData 200,1 'Cartoon drawings
  Controller.B2SSetData 201,1 'Reel background light & "High Score"
  If Credits > 1 Then CreditLight.State = 1:Controller.B2SSetData 101, 1:Controller.B2SSetData 102, 1 'State should be 1
  If Credits = 1 Then CreditLight.State = 1:Controller.B2SSetData 101, 1 'State should be 1
  If Credits = 0 Then CreditLight.State = 0 'State should be 0
  me.enabled=0
end sub


Sub Table1_KeyDown(ByVal keycode)

  If keycode = LeftFlipperKey and Rack = 0 and BIP = 1 and tilt=False Then
    LeftFlipper.RotateToEnd
    LeftFlipper001.RotateToEnd
    PlaySoundAtVol SoundFXDOF("flip1up",101,DOFOn,DOFContactors), LeftFlipper, 1
    PlayLoopSoundAtVol "buzzL", LeftFlipper, 1
  End If

  If keycode = RightFlipperKey and Rack = 0 and BIP = 1 and tilt=False Then
    RightFlipper.RotateToEnd
    RightFlipper001.RotatetoEnd
    PlaySoundAtVol SoundFXDOF("flip1up",102,DOFOn,DOFContactors), RightFlipper, 1
    PlayLoopSoundAtVol "buzz", RightFlipper, 1
  End If

  if HorseReturn.Enabled = False then
  If (keycode = PlungerKey or keycode = 57) then '57=Spacebar
    'Plunger.Fire
    If Balls > 0 and Player > 0 and BIP = 0 Then
      PlaySoundAtVol SoundFXDOF("BallLaunch",106,DOFPulse,DOFContactors), BallRelease, 1
      DOF 107, DOFPulse
      BallRelease.CreateSizedBallWithMass BallSize, BallSize^3/15625
      BallRelease.Kick D, 40

      BIP = 1
      'Disc.Image = "rotDisc"
      wallsup.enabled=1
      BIPonBG = BIPonBG + 1
      UpdateBallinBallDisplay
    End If
  End If
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
End Sub

Sub CheckTilt
  If Tilttimer.Enabled = True Then
   TiltSens = TiltSens + 1
   if TiltSens = 3 Then
     Tilt = True
'   TILTreel.setvalue 1
'         If B2SOn Then Controller.B2SSetTilt 33,1

' Add a new line here that triggers "Tilt" on backglass once I add it
      playsound "Neigh"
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

Sub turnoff
  LeftFlipper.RotateToStart:DOF 101, DOFOff
  LeftFlipper001.RotateToStart
  RightFlipper.RotateToStart:DOF 102, DOFOff
  RightFlipper001.RotateToStart
end sub

Sub Table1_KeyUp(ByVal keycode)

  If Keycode = AddCreditKey Then
        Credits = Credits + 1: If Credits > 25 Then Credits = 25
    DOF 103, DOFOn
    EMReel3.SetValue Credits
    PlaySoundAtVol "coinin", Drain, 1
    Controller.B2SSetCredits Credits
    If Credits > 1 Then CreditLight.State = 1:Controller.B2SSetData 101, 1:Controller.B2SSetData 102, 1 'State should be 1
    If Credits = 1 Then CreditLight.State = 1:Controller.B2SSetData 101, 1 'State should be 1
    If Credits = 0 Then CreditLight.State = 0 'State should be 0
  End If

  If keycode = StartGameKey and Credits > 0 and Players < 2 Then
    ResetGame
'   WinningNumberLottery
    PlaySound "BallLoad",0,1,0.25,0.25
    tilt=False
    wdown=1
    WallsDown
    DiscTimer.Enabled = true
    if motorsound=1 then playsound SoundFXDOF("motor",105,DOFOn,DOFGear), 1,.1
    GOReel.setvalue 0
    Players = Players + 1
    Credits = Credits - 1
    If Credits > 1 Then CreditLight.State = 1:Controller.B2SSetData 101, 1:Controller.B2SSetData 102, 1 'State should be 1
    If Credits = 1 Then CreditLight.State = 1:Controller.B2SSetData 101, 1 'State should be 1
    If Credits = 0 Then CreditLight.State = 0 'State should be 0
'   If Credits > 0 Then CreditLight.State = 1
'   If Credits = 0 Then CreditLight.State = 0
    If Credits < 1 Then DOF 103, DOFOff
    EMReel3.SetValue Credits
    EMReel1.Image = "Numbers"
    Player = 1
    Balls = 5
    BIP = 0
    Rack = 0
    Racks = 0
    PlayerScores(1) = 0: Scores 1,0
    PlayerScores(2) = 0: Scores 2,0
    If Players = 1 Then WinningNumberLottery
    If Players = 2 Then WinningNumberLottery2':WinningNumberLottery
    If B2SOn Then
      B2SBIP(1)
      Controller.B2SSetCredits Credits
      Controller.B2SSetData 202,1
'     Controller.B2SSetData 39+Players, 1
      Controller.B2SSetGameover 0
'     Controller.B2SSetCanPlay Players
    End if
  End If

  if BIP=1 Then
  If keycode = LeftFlipperKey and tilt=False Then
    LeftFlipper.RotateToStart
    LeftFlipper001.RotateToStart
    PlaySoundAtVol SoundFXDOF("flip1down",101,DOFOff,DOFContactors), LeftFlipper, 1
    StopSound "buzzL"
  End If

  If keycode = RightFlipperKey and tilt=False Then
    RightFlipper.RotateToStart
    RightFlipper001.RotateToStart
    PlaySoundAtVol SoundFXDOF("flip1down",102,DOFOff,DOFContactors), RightFlipper, 1
    StopSound "buzz"
  End If
  end if

End Sub

Sub WallsDown
  wdown=1
  For each objekt in RotWalls: objekt.isdropped=True: Next
End Sub

Sub WallsUp_Timer
  wdown=0
  me.enabled=0
End Sub

Sub DiscTimer_Timer()
' Eval("rotWall"&D).IsDropped = true


  If R = -1  Then D = D - 2
  If D < 0   Then D = 358
  If D = 344 Then 'was 308
    if motorsound=1 then stopsound "motor":DOF 105, DOFOff
    if motorsound=1 then PlaySoundAtVol SoundFXDOF("motor",105,DOFOn,DOFGear), Disc, 1
    R = 1
  end if
  If R =  1  Then D = D + 2
  If D > 359 Then D = 0
  If D = 14  Then 'was 50
    if motorsound=1 then stopsound "motor":DOF 105, DOFOff
    if motorsound=1 then PlaySoundAtVol SoundFXDOF("motor",105,DOFOn,DOFGear), Disc, 1
    R = -1
  end if
  Disc.RotZ = D

' if wdown=0 then Eval("rotWall"&D).IsDropped = false

End Sub

Sub B2SBIP(b)
  For j = 1 to 5
  Controller.B2SSetData j, 0
  Next
  If b > 0 Then Controller.B2SSetData b, 3
End Sub

Sub Scores(pl, points)
  PlayerScores(pl) = PlayerScores(pl) + points
  If B2SOn Then Controller.B2SSetScorePlayer pl, PlayerScores(pl)
  EMReel1.SetValue PlayerScores(1)
  EMReel2.SetValue PlayerScores(2)
End Sub

Sub EndScoresTimer_Timer()
  ScoresRoll = 1
    If SubScores(1)+SubScores(2)+SubScores(3) = 0 Then
      If Rack = 0 Then
        If Players = 2 and Player = 1 Then
          Player = 2
          If B2SOn Then
            Controller.B2SSetData 203,1
            Controller.B2SSetData 202,0
          End If
          EMReel1.Image = "NumbersDark"
          EMReel2.Image = "Numbers"
        Else
          Player = 1: Balls = Balls - 1
          If Balls = 0 Then GameOver
          If Balls > 0 Then
            EMReel1.Image = "Numbers"
            EMReel2.Image = "NumbersDark"
          End If
          If B2SOn and Balls > 0 Then
            Controller.B2SSetData 202,1
            Controller.B2SSetData 203,0
            B2SBIP(6-Balls)
          End If
        End If
      End If 'If Rack = 0
      ScoresRoll = 0

      If Balls <= 5 then EndScoresTimer.Enabled = false
      End if

End Sub

Sub Drain_Hit()
  PlaySoundAtVol "Drains", Drain, 1
  Drain.DestroyBall
  WallsDown
  StopSound "buzzL"
  StopSound "buzz"

  Tilt=False 'have tilt on BG go away

  SubPlayer = Player
  EndScoresTimer.Enabled = true

  BIP = 0
  turnoff

End Sub

Sub GameOver
  PlaySound "GameOver",0,1,0.25,0.25
' WinningTicketCheck 'I think this might force it to always pay out 1 credit. Let's try putting it back in all the horse checks
  DiscTimer.Enabled = false
  stopsound "motor":DOF 105, DOFOff
' LightOn2.State = 0
' LightOn3.State = 0
  Players = 0
  Player = 0
  Controller.B2SSetGameOver 1

  ' ------------------
  ' checks for highscore
  if PlayerScores(1)>hiscore then hiscore=PlayerScores(1)
  if PlayerScores(2)>hiscore then hiscore=PlayerScores(2)
' TextHiScore.text = hiscore
  savecredits
  If B2SOn Then
'   Controller.B2SSetData 16, 0
'   Controller.B2SSetGameover 1
'   Controller.B2SSetPlayerUp 0
'   Controller.B2SSetCanPlay 0
    Controller.B2SSetScorePlayer 5, hiscore
    B2SBIP(0)

  Balls = 0
  BIP = 0 'Not entirely sure if this is needed here?
' ResetGame - don't think it's good to have this here
  End If

End Sub

Sub Triggers_hit(idx)
  dim rotball
  rotball=idx+1
  if moveballs=1 Then
    direction(rotball)=Int(Rnd*2)-1
    if direction(rotball)=0 then direction(rotball)=1
    if ballangle(rotball)=0 then
      ballangle(rotball)=(12*direction(rotball))
'     EVAL("Primball"&rotball).roty=ballangle(rotball)
    end if
  end If

End Sub

Sub BallTurnTimer_Timer()
  dim x
  for x=1 to 15
    if ballangle(x)<>0 Then
      ballangle(x)=ballangle(x)+(12*direction(x))
      if (ballangle(x) > 359) or (ballangle(x) < -359) Then
        ballangle(x) = 0
      End If
'     EVAL("Primball"&x).roty=ballangle(x)
    end if
  next
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


Sub Pins_Hit (idx)
  PlaySound "pinhit_low", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0, AudioFade(ActiveBall)
End Sub

Sub Targets_Hit (idx)
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

Sub Spinner_Spin
  PlaySound "fx_spinner",0,.25,0,0.25
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

Sub LeftFlipper001_Collide(parm)
  RandomSoundFlipper()
End Sub

Sub RightFlipper_Collide(parm)
  RandomSoundFlipper()
End Sub

Sub RightFlipper001_Collide(parm)
  RandomSoundFlipper()
End Sub

Sub RandomSoundFlipper()
  Select Case Int(Rnd*3)+1
    Case 1 : PlaySound "flip_hit_1", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
    Case 2 : PlaySound "flip_hit_2", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
    Case 3 : PlaySound "flip_hit_3", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
  End Select
End Sub

sub savecredits
    savevalue "WilliamsWinner", "credit", Credits
    savevalue "WilliamsWinner", "hiscore", hiscore
end sub

sub loadcredits
    dim temp
  temp = LoadValue("WilliamsWinner", "credit")
    If (temp <> "") then Credits = CDbl(temp)
    temp = LoadValue("WilliamsWinner", "hiscore")
    If (temp <> "") then hiscore = CDbl(temp)
end sub

Sub Table1_Exit()
  turnoff
  savecredits
  If B2SOn Then Controller.Stop
End Sub

'*****************************************
'           BALL SHADOW
'*****************************************
Dim BallShadow
BallShadow = Array (BallShadow1,BallShadow2,BallShadow3,BallShadow4,BallShadow5,BallShadow6)

Sub BallShadowUpdate_timer()
Dim BOT, b
BOT = GetBalls

' hide shadow of deleted balls
If UBound(BOT)<(tnob-1) Then
For b = (UBound(BOT) + 1) to (tnob-1)
BallShadow(b).visible = 0
Next
End If

'exit the Sub if no balls on the table
If UBound(BOT) = -1 Then Exit Sub

'render the shadow for each ball
For b = 0 to UBound(BOT)
If BOT(b).X < Table1.Width/2 Then
BallShadow(b).X = ((BOT(b).X) - (Ballsize/6) + ((BOT(b).X - (Table1.Width/2))/7)) + 10
Else
BallShadow(b).X = ((BOT(b).X) + (BAllsize/6) + ((BOT(b).X - (Table1.Width/2))/7)) - 10
End If
ballShadow(b).Y = BOT(b).Y + 20
If BOT(b).Z > 20 Then
BallShadow(b).visible = 1
Else
BallShadow(b).visible = 0
End If
Next
End Sub

'***********************
' slingshots (adapted from loserman's "Big Deal" table)
'This takes card of the bottom slings. I don't *think* there are slings elsewhere, but I could be wrong.
'I am building this table off eBay images and low res YouTube videos, after all

Dim LStep, RStep, i, xx

Sub RightSlingShot_Hit
    PlaySoundAtVol SoundFXDOF("right_slingshot",205,DOFPulse,DOFContactors), ActiveBall, 1
  DOF 206,2
    RSling0.Visible = 0
    RSling1.Visible = 1
    sling1.TransZ = -20
    RStep = 0
    RSling0.TimerEnabled = 1

End Sub

Sub RSling0_Timer
    Select Case RStep
        Case 3:RSLing1.Visible = 0:RSLing2.Visible = 1:sling1.TransZ = -10
        Case 4:RSLing2.Visible = 0:RSling0.Visible = 1:sling1.TransZ = 0
    End Select
    RStep = RStep + 1
End Sub

Sub LeftSlingShot_Hit 'was LeftSlingShot_Slingshot
    PlaySoundAtVol SoundFXDOF("left_slingshot",203,DOFPulse,DOFContactors), ActiveBall, 1
  DOF 204,2
    LSling0.Visible = 0
    LSling1.Visible = 1
    sling2.TransZ = -20
    LStep = 0
    LSling0.TimerEnabled = 1
End Sub

Sub LSling0_Timer
    Select Case LStep
        Case 3:LSLing1.Visible = 0:LSLing2.Visible = 1:sling2.TransZ = -10
        Case 4:LSLing2.Visible = 0:LSLing0.Visible = 1:sling2.TransZ = 0
    End Select
    LStep = LStep + 1
End Sub

'***********************
'Lights on Playfield
Dim ABDouble, CDDouble

Sub ABTarget_Hit
  Playsound "Bell"
  LightA.State = 1
  LightB.State = 1
  ABDouble = True
End Sub

Sub CDTarget_Hit
  Playsound "Bell"
  LightC.State = 1
  LightD.State = 1
  CDDouble = True
End Sub


'Horse Moving
Dim HorseSpace1, HorseSpace2, HorseSpace3, HorseSpace4, HorseSpace5, HorseSpace6
Dim DoubleSpace135, DoubleSpace246
Dim HoldYourHorses

Sub HeldHorseCheck
  If HorseSpace1 = 7 Then HoldYourHorses = True end if
  If HorseSpace2 = 7 Then HoldYourHorses = True end if
  If HorseSpace3 = 7 Then HoldYourHorses = True end if
  If HorseSpace4 = 7 Then HoldYourHorses = True end if
  If HorseSpace5 = 7 Then HoldYourHorses = True end if
  If HorseSpace6 = 7 Then HoldYourHorses = True end if
End Sub

Sub BackHorse1_Hit
  HeldHorseCheck
  If HoldYourHorses = False Then Playsound "HorseMovement" end if
' If Winner1 = 1 and DoubleSpace135 = True then Scores Player,5000 end If
  If LeftOne = True and DoubleSpace135 = True then PlayerScores(1) = PlayerScores(1) + 5000 end if
  If Right1 = True and DoubleSpace135 = True then PlayerScores(2) = PlayerScores(2) + 5000 end if
  Controller.B2SSetScorePlayer 1, PlayerScores(1)
  Controller.B2SSetScorePlayer 2, PlayerScores(2)
  If HoldYourHorses = True Then Playsound "Hit1":Exit Sub End if
  HorseSpace1 = HorseSpace1 + 1
  If HorseSpace1 > 7 Then HorseSpace1 = 7 End if
  Horse1Move
End Sub

Sub Horse1Move
  If HorseSpace1 = 1 and DoubleSpace135 = False Then Horse1.Transx = -50:Horse1a.enabled = True  end If
  If HorseSpace1 = 2 and DoubleSpace135 = False Then Horse1.Transx = -150:Horse1b.enabled = True end If
  If HorseSpace1 = 3 and DoubleSpace135 = False Then Horse1.Transx = -250:Horse1c.enabled = True end If
  If HorseSpace1 = 4 and DoubleSpace135 = False Then Horse1.Transx = -350:Horse1d.enabled = True end If
  If HorseSpace1 = 5 and DoubleSpace135 = False Then Horse1.Transx = -450:Horse1e.enabled = True end If
  If HorseSpace1 = 6 and DoubleSpace135 = False Then Horse1.Transx = -550:Horse1f.enabled = True end If
  If HorseSpace1 = 7 and DoubleSpace135 = False Then Horse1.Transx = -650:Horse1g.enabled = True end If
  If HorseSpace1 = 1 and DoubleSpace135 = True Then Horse1.Transx = -100:Horse1a.enabled = True  end If
  If HorseSpace1 = 2 and DoubleSpace135 = True Then Horse1.Transx = -200:Horse1b.enabled = True end If
  If HorseSpace1 = 3 and DoubleSpace135 = True Then Horse1.Transx = -300:Horse1c.enabled = True end If
  If HorseSpace1 = 4 and DoubleSpace135 = True Then Horse1.Transx = -400:Horse1d.enabled = True end If
  If HorseSpace1 = 5 and DoubleSpace135 = True Then Horse1.Transx = -500:Horse1e.enabled = True end If
  If HorseSpace1 = 6 and DoubleSpace135 = True Then Horse1.Transx = -600:Horse1f.enabled = True end If
  If HorseSpace1 = 7 and DoubleSpace135 = True Then Horse1.Transx = -700:HoldYourHorses = True end If
  Horse1Check
End Sub

Sub BackHorse2_Hit
  HeldHorseCheck
  If HoldYourHorses = False Then Playsound "HorseMovement" end if
' If Winner2 = 1 and DoubleSpace246 = True then Scores Player,5000 end If
  If LeftTwo = True and DoubleSpace135 = True then PlayerScores(1) = PlayerScores(1) + 5000 end if
  If Right2 = True and DoubleSpace135 = True then PlayerScores(2) = PlayerScores(2) + 5000 end if
  Controller.B2SSetScorePlayer 1, PlayerScores(1)
  Controller.B2SSetScorePlayer 2, PlayerScores(2)
  If HoldYourHorses = True Then Playsound "Hit1":Exit Sub End if
  HorseSpace2 = HorseSpace2 + 1
  If HorseSpace2 > 7 Then HorseSpace2 = 7 End if
  Horse2Move
End Sub

Sub Horse2Move
  If HorseSpace2 = 1 and DoubleSpace246 = False Then Horse2.Transx = -50:Horse2a.enabled = True  end If
  If HorseSpace2 = 2 and DoubleSpace246 = False Then Horse2.Transx = -150:Horse2b.enabled = True end If
  If HorseSpace2 = 3 and DoubleSpace246 = False Then Horse2.Transx = -250:Horse2c.enabled = True end If
  If HorseSpace2 = 4 and DoubleSpace246 = False Then Horse2.Transx = -350:Horse2d.enabled = True end If
  If HorseSpace2 = 5 and DoubleSpace246 = False Then Horse2.Transx = -450:Horse2e.enabled = True end If
  If HorseSpace2 = 6 and DoubleSpace246 = False Then Horse2.Transx = -550:Horse2f.enabled = True end If
  If HorseSpace2 = 7 and DoubleSpace246 = False Then Horse2.Transx = -650:Horse2g.enabled = True end If
  If HorseSpace2 = 1 and DoubleSpace246 = True Then Horse2.Transx = -100:Horse2a.enabled = True  end If
  If HorseSpace2 = 2 and DoubleSpace246 = True Then Horse2.Transx = -200:Horse2b.enabled = True end If
  If HorseSpace2 = 3 and DoubleSpace246 = True Then Horse2.Transx = -300:Horse2c.enabled = True end If
  If HorseSpace2 = 4 and DoubleSpace246 = True Then Horse2.Transx = -400:Horse2d.enabled = True end If
  If HorseSpace2 = 5 and DoubleSpace246 = True Then Horse2.Transx = -500:Horse2e.enabled = True end If
  If HorseSpace2 = 6 and DoubleSpace246 = True Then Horse2.Transx = -600:Horse2f.enabled = True end If
  If HorseSpace2 = 7 and DoubleSpace246 = True Then Horse2.Transx = -700:HoldYourHorses = True end If
  Horse2Check
End Sub

Sub BackHorse3_Hit
  HeldHorseCheck
  If HoldYourHorses = False Then Playsound "HorseMovement" end if
' If Winner3 = 1 and DoubleSpace135 = True then Scores Player,5000 end If
  If LeftThree = True and DoubleSpace135 = True then PlayerScores(1) = PlayerScores(1) + 5000 end if
  If Right3 = True and DoubleSpace135 = True then PlayerScores(2) = PlayerScores(2) + 5000 end if
  Controller.B2SSetScorePlayer 1, PlayerScores(1)
  Controller.B2SSetScorePlayer 2, PlayerScores(2)
  If HoldYourHorses = True Then Playsound "Hit1":Exit Sub End if
  HorseSpace3 = HorseSpace3 + 1
  If HorseSpace3 > 7 Then HorseSpace3 = 7 End if
  Horse3Move
End Sub

Sub Horse3Move
  If HorseSpace3 = 1 and DoubleSpace135 = False Then Horse3.Transx = -50:Horse3a.enabled = True  end If
  If HorseSpace3 = 2 and DoubleSpace135 = False Then Horse3.Transx = -150:Horse3b.enabled = True end If
  If HorseSpace3 = 3 and DoubleSpace135 = False Then Horse3.Transx = -250:Horse3c.enabled = True end If
  If HorseSpace3 = 4 and DoubleSpace135 = False Then Horse3.Transx = -350:Horse3d.enabled = True end If
  If HorseSpace3 = 5 and DoubleSpace135 = False Then Horse3.Transx = -450:Horse3e.enabled = True end If
  If HorseSpace3 = 6 and DoubleSpace135 = False Then Horse3.Transx = -550:Horse3f.enabled = True end If
  If HorseSpace3 = 7 and DoubleSpace135 = False Then Horse3.Transx = -650:Horse3g.enabled = True end If
  If HorseSpace3 = 1 and DoubleSpace135 = True Then Horse3.Transx = -100:Horse3a.enabled = True  end If
  If HorseSpace3 = 2 and DoubleSpace135 = True Then Horse3.Transx = -200:Horse3b.enabled = True end If
  If HorseSpace3 = 3 and DoubleSpace135 = True Then Horse3.Transx = -300:Horse3c.enabled = True end If
  If HorseSpace3 = 4 and DoubleSpace135 = True Then Horse3.Transx = -400:Horse3d.enabled = True end If
  If HorseSpace3 = 5 and DoubleSpace135 = True Then Horse3.Transx = -500:Horse3e.enabled = True end If
  If HorseSpace3 = 6 and DoubleSpace135 = True Then Horse3.Transx = -600:Horse3f.enabled = True end If
  If HorseSpace3 = 7 and DoubleSpace135 = True Then Horse3.Transx = -700:HoldYourHorses = True end If
  Horse3Check
End Sub

Sub BackHorse4_Hit
  HeldHorseCheck
  If HoldYourHorses = False Then Playsound "HorseMovement" end if
' If Winner4 = 1 and DoubleSpace246 = True then Scores Player,5000 end If
  If LeftFour = True and DoubleSpace135 = True then PlayerScores(1) = PlayerScores(1) + 5000 end if
  If Right4 = True and DoubleSpace135 = True then PlayerScores(2) = PlayerScores(2) + 5000 end if
  Controller.B2SSetScorePlayer 1, PlayerScores(1)
  Controller.B2SSetScorePlayer 2, PlayerScores(2)
  If HoldYourHorses = True Then Playsound "Hit1":Exit Sub End if
  HorseSpace4 = HorseSpace4 + 1
  If HorseSpace4 > 7 Then HorseSpace4 = 7 End if
  Horse4Move
End Sub

Sub Horse4Move
  If HorseSpace4 = 1 and DoubleSpace246 = False Then Horse4.Transx = -50:Horse4a.enabled = True  end If
  If HorseSpace4 = 2 and DoubleSpace246 = False Then Horse4.Transx = -150:Horse4b.enabled = True end If
  If HorseSpace4 = 3 and DoubleSpace246 = False Then Horse4.Transx = -250:Horse4c.enabled = True end If
  If HorseSpace4 = 4 and DoubleSpace246 = False Then Horse4.Transx = -350:Horse4d.enabled = True end If
  If HorseSpace4 = 5 and DoubleSpace246 = False Then Horse4.Transx = -450:Horse4e.enabled = True end If
  If HorseSpace4 = 6 and DoubleSpace246 = False Then Horse4.Transx = -550:Horse4f.enabled = True end If
  If HorseSpace4 = 7 and DoubleSpace246 = False Then Horse4.Transx = -650:Horse4g.enabled = True end If
  If HorseSpace4 = 1 and DoubleSpace246 = True Then Horse4.Transx = -100:Horse4a.enabled = True  end If
  If HorseSpace4 = 2 and DoubleSpace246 = True Then Horse4.Transx = -200:Horse4b.enabled = True end If
  If HorseSpace4 = 3 and DoubleSpace246 = True Then Horse4.Transx = -300:Horse4c.enabled = True end If
  If HorseSpace4 = 4 and DoubleSpace246 = True Then Horse4.Transx = -400:Horse4d.enabled = True end If
  If HorseSpace4 = 5 and DoubleSpace246 = True Then Horse4.Transx = -500:Horse4e.enabled = True end If
  If HorseSpace4 = 6 and DoubleSpace246 = True Then Horse4.Transx = -600:Horse4f.enabled = True end If
  If HorseSpace4 = 7 and DoubleSpace246 = True Then Horse4.Transx = -700:HoldYourHorses = True end If
  Horse4Check
End Sub

Sub BackHorse5_Hit
  HeldHorseCheck
  If HoldYourHorses = False Then Playsound "HorseMovement" end if
' If Winner5 = 1 and DoubleSpace135 = True then Scores Player,5000 end If
  If LeftFive = True and DoubleSpace135 = True then PlayerScores(1) = PlayerScores(1) + 5000 end if
  If Right5 = True and DoubleSpace135 = True then PlayerScores(2) = PlayerScores(2) + 5000 end if
  Controller.B2SSetScorePlayer 1, PlayerScores(1)
  Controller.B2SSetScorePlayer 2, PlayerScores(2)
  If HoldYourHorses = True Then Playsound "Hit1":Exit Sub End if
  HorseSpace5 = HorseSpace5 + 1
  If HorseSpace5 > 7 Then HorseSpace5 = 7 End if
  Horse5Move
End Sub

Sub Horse5Move
  If HorseSpace5 = 1 and DoubleSpace135 = False Then Horse5.Transx = -50:Horse5a.enabled = True  end If
  If HorseSpace5 = 2 and DoubleSpace135 = False Then Horse5.Transx = -150:Horse5b.enabled = True end If
  If HorseSpace5 = 3 and DoubleSpace135 = False Then Horse5.Transx = -250:Horse5c.enabled = True end If
  If HorseSpace5 = 4 and DoubleSpace135 = False Then Horse5.Transx = -350:Horse5d.enabled = True end If
  If HorseSpace5 = 5 and DoubleSpace135 = False Then Horse5.Transx = -450:Horse5e.enabled = True end If
  If HorseSpace5 = 6 and DoubleSpace135 = False Then Horse5.Transx = -550:Horse5f.enabled = True end If
  If HorseSpace5 = 7 and DoubleSpace135 = False Then Horse5.Transx = -650:Horse5g.enabled = True end If
  If HorseSpace5 = 1 and DoubleSpace135 = True Then Horse5.Transx = -100:Horse5a.enabled = True  end If
  If HorseSpace5 = 2 and DoubleSpace135 = True Then Horse5.Transx = -200:Horse5b.enabled = True end If
  If HorseSpace5 = 3 and DoubleSpace135 = True Then Horse5.Transx = -300:Horse5c.enabled = True end If
  If HorseSpace5 = 4 and DoubleSpace135 = True Then Horse5.Transx = -400:Horse5d.enabled = True end If
  If HorseSpace5 = 5 and DoubleSpace135 = True Then Horse5.Transx = -500:Horse5e.enabled = True end If
  If HorseSpace5 = 6 and DoubleSpace135 = True Then Horse5.Transx = -600:Horse5f.enabled = True end If
  If HorseSpace5 = 7 and DoubleSpace135 = True Then Horse5.Transx = -700:HoldYourHorses = True end If
  Horse5Check
End Sub

Sub BackHorse6_Hit
  HeldHorseCheck
  If HoldYourHorses = False Then Playsound "HorseMovement" end if
' If Winner6 = 1 and DoubleSpace246 = True then Scores Player,5000 end If
  If LeftSix = True and DoubleSpace135 = True then PlayerScores(1) = PlayerScores(1) + 5000 end if
  If Right6 = True and DoubleSpace135 = True then PlayerScores(2) = PlayerScores(2) + 5000 end if
  Controller.B2SSetScorePlayer 1, PlayerScores(1)
  Controller.B2SSetScorePlayer 2, PlayerScores(2)
  If HoldYourHorses = True Then Playsound "Hit1":Exit Sub End if
  HorseSpace6 = HorseSpace6 + 1
  If HorseSpace6 > 7 Then HorseSpace6 = 7 End if
  Horse6Move
End Sub

Sub Horse6Move
  If HorseSpace6 = 1 and DoubleSpace246 = False Then Horse6.Transx = -50:Horse6a.enabled = True  end If
  If HorseSpace6 = 2 and DoubleSpace246 = False Then Horse6.Transx = -150:Horse6b.enabled = True end If
  If HorseSpace6 = 3 and DoubleSpace246 = False Then Horse6.Transx = -250:Horse6c.enabled = True end If
  If HorseSpace6 = 4 and DoubleSpace246 = False Then Horse6.Transx = -350:Horse6d.enabled = True end If
  If HorseSpace6 = 5 and DoubleSpace246 = False Then Horse6.Transx = -450:Horse6e.enabled = True end If
  If HorseSpace6 = 6 and DoubleSpace246 = False Then Horse6.Transx = -550:Horse6f.enabled = True end If
  If HorseSpace6 = 7 and DoubleSpace246 = False Then Horse6.Transx = -650:Horse6g.enabled = True end If
  If HorseSpace6 = 1 and DoubleSpace246 = True Then Horse6.Transx = -100:Horse6a.enabled = True  end If
  If HorseSpace6 = 2 and DoubleSpace246 = True Then Horse6.Transx = -200:Horse6b.enabled = True end If
  If HorseSpace6 = 3 and DoubleSpace246 = True Then Horse6.Transx = -300:Horse6c.enabled = True end If
  If HorseSpace6 = 4 and DoubleSpace246 = True Then Horse6.Transx = -400:Horse6d.enabled = True end If
  If HorseSpace6 = 5 and DoubleSpace246 = True Then Horse6.Transx = -500:Horse6e.enabled = True end If
  If HorseSpace6 = 6 and DoubleSpace246 = True Then Horse6.Transx = -600:Horse6f.enabled = True end If
  If HorseSpace6 = 7 and DoubleSpace246 = True Then Horse6.Transx = -700:HoldYourHorses = True end If
  Horse6Check
End Sub

'Horse Movement Timers
'Horse 1 Timers

Sub Horse1a_Timer()
  If DoubleSpace135 = False Then Horse1.Transx = -100 end if
  If DoubleSpace135 = True Then Horse1.Transx = -200:HorseSpace1 = 2 end if
  Horse1a.enabled = False
End Sub

Sub Horse1b_Timer()
  If DoubleSpace135 = False Then Horse1.Transx = -200 end if
  If DoubleSpace135 = True Then Horse1.Transx = -300:HorseSpace1 = 3 end if
  Horse1b.enabled = False
End Sub

Sub Horse1c_Timer()
  If DoubleSpace135 = False Then Horse1.Transx = -300 end if
  If DoubleSpace135 = True Then Horse1.Transx = -400:HorseSpace1 = 4 end if
  Horse1c.enabled = False
End Sub

Sub Horse1d_Timer()
  If DoubleSpace135 = False Then Horse1.Transx = -400 end if
  If DoubleSpace135 = True Then Horse1.Transx = -500:HorseSpace1 = 5 end if
  Horse1d.enabled = False
End Sub

Sub Horse1e_Timer()
  If DoubleSpace135 = False Then Horse1.Transx = -500 end if
  If DoubleSpace135 = True Then Horse1.Transx = -600:HorseSpace1 = 6 end if
  Horse1e.enabled = False
End Sub

Sub Horse1f_Timer()
  If DoubleSpace135 = False Then Horse1.Transx = -600 end if
  If DoubleSpace135 = True Then Horse1.Transx = -700:HoldYourHorses = True:HorseSpace1 = 7:Controller.B2SSetData 115,1 end if
  Horse1f.enabled = False
End Sub

Sub Horse1g_Timer()
  Horse1.Transx = -700
  Horse1g.enabled = False
  HoldYourHorses = True
  Controller.B2SSetData 115,1
End Sub

'Horse 2 Timers

Sub Horse2a_Timer()
  If DoubleSpace246 = False Then Horse2.Transx = -100 end if
  If DoubleSpace246 = True Then Horse2.Transx = -200:HorseSpace2 = 2 end if
  Horse2a.enabled = False
End Sub

Sub Horse2b_Timer()
  If DoubleSpace246 = False Then Horse2.Transx = -200 end if
  If DoubleSpace246 = True Then Horse2.Transx = -300:HorseSpace2 = 3 end if
  Horse2b.enabled = False
End Sub

Sub Horse2c_Timer()
  If DoubleSpace246 = False Then Horse2.Transx = -300 end if
  If DoubleSpace246 = True Then Horse2.Transx = -400:HorseSpace2 = 4 end if
  Horse2c.enabled = False
End Sub

Sub Horse2d_Timer()
  If DoubleSpace246 = False Then Horse2.Transx = -400 end if
  If DoubleSpace246 = True Then Horse2.Transx = -500:HorseSpace2 = 5 end if
  Horse2d.enabled = False
End Sub

Sub Horse2e_Timer()
  If DoubleSpace246 = False Then Horse2.Transx = -500 end if
  If DoubleSpace246 = True Then Horse2.Transx = -600:HorseSpace2 = 6 end if
  Horse2e.enabled = False
End Sub

Sub Horse2f_Timer()
  If DoubleSpace246 = False Then Horse2.Transx = -600 end if
  If DoubleSpace246 = True Then Horse2.Transx = -700:HoldYourHorses = True:HorseSpace2 = 7:Controller.B2SSetData 116,1 end if
  Horse2f.enabled = False
End Sub

Sub Horse2g_Timer()
  Horse2.Transx = -700
  Horse2g.enabled = False
  HoldYourHorses = True
  Controller.B2SSetData 116,1
End Sub

'Horse 3 Timers

Sub Horse3a_Timer()
  If DoubleSpace135 = False Then Horse3.Transx = -100 end if
  If DoubleSpace135 = True Then Horse3.Transx = -200:HorseSpace3 = 2 end if
  Horse3a.enabled = False
End Sub

Sub Horse3b_Timer()
  If DoubleSpace135 = False Then Horse3.Transx = -200 end if
  If DoubleSpace135 = True Then Horse3.Transx = -300:HorseSpace3 = 3 end if
  Horse3b.enabled = False
End Sub

Sub Horse3c_Timer()
  If DoubleSpace135 = False Then Horse3.Transx = -300 end if
  If DoubleSpace135 = True Then Horse3.Transx = -400:HorseSpace3 = 4 end if
  Horse3c.enabled = False
End Sub

Sub Horse3d_Timer()
  If DoubleSpace135 = False Then Horse3.Transx = -400 end if
  If DoubleSpace135 = True Then Horse3.Transx = -500:HorseSpace3 = 5 end if
  Horse3d.enabled = False
End Sub

Sub Horse3e_Timer()
  If DoubleSpace135 = False Then Horse3.Transx = -500 end if
  If DoubleSpace135 = True Then Horse3.Transx = -600:HorseSpace3 = 6 end if
  Horse3e.enabled = False
End Sub

Sub Horse3f_Timer()
  If DoubleSpace135 = False Then Horse3.Transx = -600 end if
  If DoubleSpace135 = True Then Horse3.Transx = -700:HoldYourHorses = True:HorseSpace3 = 7::Controller.B2SSetData 117,1 end if
  Horse3f.enabled = False
End Sub

Sub Horse3g_Timer()
  Horse3.Transx = -700
  Horse3g.enabled = False
  HoldYourHorses = True
  Controller.B2SSetData 117,1
End Sub

'Horse 4 Timers

Sub Horse4a_Timer()
  If DoubleSpace246 = False Then Horse4.Transx = -100 end if
  If DoubleSpace246 = True Then Horse4.Transx = -200:HorseSpace4 = 2 end if
  Horse4a.enabled = False
End Sub

Sub Horse4b_Timer()
  If DoubleSpace246 = False Then Horse4.Transx = -200 end if
  If DoubleSpace246 = True Then Horse4.Transx = -300:HorseSpace4 = 3 end if
  Horse4b.enabled = False
End Sub

Sub Horse4c_Timer()
  If DoubleSpace246 = False Then Horse4.Transx = -300 end if
  If DoubleSpace246 = True Then Horse4.Transx = -400:HorseSpace4 = 4 end if
  Horse4c.enabled = False
End Sub

Sub Horse4d_Timer()
  If DoubleSpace246 = False Then Horse4.Transx = -400 end if
  If DoubleSpace246 = True Then Horse4.Transx = -500:HorseSpace4 = 5 end if
  Horse4d.enabled = False
End Sub

Sub Horse4e_Timer()
  If DoubleSpace246 = False Then Horse4.Transx = -500 end if
  If DoubleSpace246 = True Then Horse4.Transx = -600:HorseSpace4 = 6 end if
  Horse4e.enabled = False
End Sub

Sub Horse4f_Timer()
  If DoubleSpace246 = False Then Horse4.Transx = -600 end if
  If DoubleSpace246 = True Then Horse4.Transx = -700:HoldYourHorses = True:HorseSpace4 = 7:Controller.B2SSetData 118,1 end if
  Horse4f.enabled = False
End Sub

Sub Horse4g_Timer()
  Horse4.Transx = -700
  Horse4g.enabled = False
  HoldYourHorses = True
  Controller.B2SSetData 118,1
End Sub

'Horse 5 Timers

Sub Horse5a_Timer()
  If DoubleSpace135 = False Then Horse5.Transx = -100 end if
  If DoubleSpace135 = True Then Horse5.Transx = -200:HorseSpace5 = 2 end if
  Horse5a.enabled = False
End Sub

Sub Horse5b_Timer()
  If DoubleSpace135 = False Then Horse5.Transx = -200 end if
  If DoubleSpace135 = True Then Horse5.Transx = -300:HorseSpace5 = 3 end if
  Horse5b.enabled = False
End Sub

Sub Horse5c_Timer()
  If DoubleSpace135 = False Then Horse5.Transx = -300 end if
  If DoubleSpace135 = True Then Horse5.Transx = -400:HorseSpace5 = 4 end if
  Horse5c.enabled = False
End Sub

Sub Horse5d_Timer()
  If DoubleSpace135 = False Then Horse5.Transx = -400 end if
  If DoubleSpace135 = True Then Horse5.Transx = -500:HorseSpace5 = 5 end if
  Horse5d.enabled = False
End Sub

Sub Horse5e_Timer()
  If DoubleSpace135 = False Then Horse5.Transx = -500 end if
  If DoubleSpace135 = True Then Horse5.Transx = -600:HorseSpace5 = 6 end if
  Horse5e.enabled = False
End Sub

Sub Horse5f_Timer()
  If DoubleSpace135 = False Then Horse5.Transx = -600 end if
  If DoubleSpace135 = True Then Horse5.Transx = -700:HoldYourHorses = True:HorseSpace5 = 7:Controller.B2SSetData 119,1 end if
  Horse5f.enabled = False
End Sub

Sub Horse5g_Timer()
  Horse5.Transx = -700
  Horse5g.enabled = False
  HoldYourHorses = True
  Controller.B2SSetData 119,1
End Sub

'Horse 6 Timers

Sub Horse6a_Timer()
  If DoubleSpace246 = False Then Horse6.Transx = -100 end if
  If DoubleSpace246 = True Then Horse6.Transx = -200:HorseSpace6 = 2 end if
  Horse6a.enabled = False
End Sub

Sub Horse6b_Timer()
  If DoubleSpace246 = False Then Horse6.Transx = -200 end if
  If DoubleSpace246 = True Then Horse6.Transx = -300:HorseSpace6 = 3 end if
  Horse6b.enabled = False
End Sub

Sub Horse6c_Timer()
  If DoubleSpace246 = False Then Horse6.Transx = -300 end if
  If DoubleSpace246 = True Then Horse6.Transx = -400:HorseSpace6 = 4 end if
  Horse6c.enabled = False
End Sub

Sub Horse6d_Timer()
  If DoubleSpace246 = False Then Horse6.Transx = -400 end if
  If DoubleSpace246 = True Then Horse6.Transx = -500:HorseSpace6 = 5 end if
  Horse6d.enabled = False
End Sub

Sub Horse6e_Timer()
  If DoubleSpace246 = False Then Horse6.Transx = -500 end if
  If DoubleSpace246 = True Then Horse6.Transx = -600:HorseSpace6 = 6 end if
  Horse6e.enabled = False
End Sub

Sub Horse6f_Timer()
  If DoubleSpace246 = False Then Horse6.Transx = -600 end if
  If DoubleSpace246 = True Then Horse6.Transx = -700::HoldYourHorses = True:HorseSpace6 = 7::Controller.B2SSetData 120,1 end if
  Horse6f.enabled = False
End Sub

Sub Horse6g_Timer()
  Horse6.Transx = -700
  Horse6g.enabled = False
  HoldYourHorses = True
  Controller.B2SSetData 120,1
End Sub

'*********************
'Picking Winner
'*********************

Dim WinnerOne, WinnerTwo, WinnerThree, WinnerFour, WinnerFive, WinnerSix 'Player1 Goal variable
Dim Winner1, Winner2, Winner3, Winner4, Winner5, Winner6 'Player2 Goal variable
Dim WinnerMatch
Dim LeftOne, LeftTwo, LeftThree, LeftFour, LeftFive, LeftSix
Dim Right1, Right2, Right3, Right4, Right5, Right6

Sub WinningNumberLottery    'This selects a random number on the left horseshoe on the backglass as Player 1's goal
  Select Case Int(Rnd*6)+1
    Case 1 : HorseOneWinner
    Case 2 : HorseTwoWinner
    Case 3 : HorseThreeWinner
    Case 4 : HorseFourWinner
    Case 5 : HorseFiveWinner
    Case 6 : HorseSixWinner
  End Select
End Sub

Sub WinningNumberLottery2   'This selects a random number on the right horseshoe on the backglass as Player 2's goal
  Select Case Int(Rnd*30)+1
Case 1 : HorseOneWinner:Horse2Winner
Case 2 : HorseOneWinner:Horse3Winner
Case 3 : HorseOneWinner:Horse4Winner
Case 4 : HorseOneWinner:Horse5Winner
Case 5 : HorseOneWinner:Horse6Winner
Case 6 : HorseTwoWinner:Horse1Winner
Case 7 : HorseTwoWinner:Horse3Winner
Case 8 : HorseTwoWinner:Horse4Winner
Case 9 : HorseTwoWinner:Horse5Winner
Case 10 : HorseTwoWinner:Horse6Winner
Case 11 : HorseThreeWinner:Horse1Winner
Case 12 : HorseThreeWinner:Horse2Winner
Case 13 : HorseThreeWinner:Horse4Winner
Case 14 : HorseThreeWinner:Horse5Winner
Case 15 : HorseThreeWinner:Horse6Winner
Case 16 : HorseFourWinner:Horse1Winner
Case 17 : HorseFourWinner:Horse2Winner
Case 18 : HorseFourWinner:Horse3Winner
Case 19 : HorseFourWinner:Horse5Winner
Case 20 : HorseFourWinner:Horse6Winner
Case 21 : HorseFiveWinner:Horse1Winner
Case 22 : HorseFiveWinner:Horse2Winner
Case 23 : HorseFiveWinner:Horse3Winner
Case 24 : HorseFiveWinner:Horse4Winner
Case 25 : HorseFiveWinner:Horse6Winner
Case 26 : HorseSixWinner:Horse1Winner
Case 27 : HorseSixWinner:Horse2Winner
Case 28 : HorseSixWinner:Horse3Winner
Case 29 : HorseSixWinner:Horse4Winner
Case 30 : HorseSixWinner:Horse5Winner

  End Select
End Sub

Sub HorseOneWinner
  WinnerOne = 1:WinnerTwo = 0:WinnerThree = 0:WinnerFour = 0:WinnerFive = 0:WinnerSix = 0
  LeftOne = True:LeftTwo = False:LeftThree = False:LeftFour = False:LeftFive = False:LeftSix = False
  Controller.B2SSetData 121,1
End Sub

Sub HorseTwoWinner
  WinnerOne = 0:WinnerTwo = 1:WinnerThree = 0:WinnerFour = 0:WinnerFive = 0:WinnerSix = 0
  LeftOne = False:LeftTwo = True:LeftThree = False:LeftFour = False:LeftFive = False:LeftSix = False
  Controller.B2SSetData 122,1
End Sub

Sub HorseThreeWinner
  WinnerOne = 0:WinnerTwo = 0:WinnerThree = 1:WinnerFour = 0:WinnerFive = 0:WinnerSix = 0
  LeftOne = False:LeftTwo = False:LeftThree = True:LeftFour = False:LeftFive = False:LeftSix = False
  Controller.B2SSetData 123,1
End Sub

Sub HorseFourWinner
  WinnerOne = 0:WinnerTwo = 0:WinnerThree = 1:WinnerFour = 0:WinnerFive = 0:WinnerSix = 0
  LeftOne = False:LeftTwo = False:LeftThree = False:LeftFour = True:LeftFive = False:LeftSix = False
  Controller.B2SSetData 124,1
End Sub

Sub HorseFiveWinner
  WinnerOne = 0:WinnerTwo = 0:WinnerThree = 0:WinnerFour = 0:WinnerFive = 1:WinnerSix = 0
  LeftOne = False:LeftTwo = False:LeftThree = False:LeftFour = False:LeftFive = True:LeftSix = False
  Controller.B2SSetData 125,1
End Sub

Sub HorseSixWinner
  WinnerOne = 0:WinnerTwo = 0:WinnerThree = 0:WinnerFour = 0:WinnerFive = 0:WinnerSix = 1
  LeftOne = False:LeftTwo = False:LeftThree = False:LeftFour = False:LeftFive = False:LeftSix = True
  Controller.B2SSetData 126,1
End Sub

Sub Horse1Winner
  Winner1 = 1:Winner2 = 0:Winner3 = 0:Winner4 = 0:Winner5 = 0:Winner6 = 0
  Right1 = True:Right2 = False:Right3 = False:Right4 = False:Right5 = False:Right6 = False
  Controller.B2SSetData 131,1
End Sub

Sub Horse2Winner
  Winner1 = 0:Winner2 = 1:Winner3 = 0:Winner4 = 0:Winner5 = 0:Winner6 = 0
  Right1 = False:Right2 = True:Right3 = False:Right4 = False:Right5 = False:Right6 = False
  Controller.B2SSetData 132,1
End Sub

Sub Horse3Winner
  Winner1 = 0:Winner2 = 0:Winner3 = 1:Winner4 = 0:Winner5 = 0:Winner6 = 0
  Right1 = False:Right2 = False:Right3 = True:Right4 = False:Right5 = False:Right6 = False
  Controller.B2SSetData 133,1
End Sub

Sub Horse4Winner
  Winner1 = 0:Winner2 = 0:Winner3 = 0:Winner4 = 1:Winner5 = 0:Winner6 = 0
  Right1 = False:Right2 = False:Right3 = False:Right4 = True:Right5 = False:Right6 = False
  Controller.B2SSetData 134,1
End Sub

Sub Horse5Winner
  Winner1 = 0:Winner2 = 0:Winner3 = 0:Winner4 = 0:Winner5 = 1:Winner6 = 0
  Right1 = False:Right2 = False:Right3 = False:Right4 = False:Right5 = True:Right6 = False
  Controller.B2SSetData 135,1
End Sub

Sub Horse6Winner
  Winner1 = 0:Winner2 = 0:Winner3 = 0:Winner4 = 0:Winner5 = 0:Winner6 = 1
  Right1 = False:Right2 = False:Right3 = False:Right4 = False:Right5 = False:Right6 = True
  Controller.B2SSetData 136,1
End Sub

Sub Horse1Check
  If HorseSpace1 < 7 Then Exit Sub End If
  If WinnerOne = 1 and HorseSpace1 >= 7 Then PlayerScores(1) = PlayerScores(1) + Purse:WinnerMatch = 1 End if
  If Winner1 = 1 and HorseSpace1 >= 7 Then PlayerScores(2) = PlayerScores(2) + Purse:WinnerMatch = 1 End if
  If NOT WinnerOne = 1 OR Winner1 = 1 Then WinnerMatch = 0 End if
  HoldYourHorses = True
End Sub

Sub Horse2Check
  If HorseSpace2 < 7 Then Exit Sub End If
  If WinnerTwo = 1 and HorseSpace2 >= 7 Then PlayerScores(1) = PlayerScores(1) + Purse:WinnerMatch = 1 End if
  If Winner2 = 1 and HorseSpace2 >= 7 Then PlayerScores(2) = PlayerScores(2) + Purse:WinnerMatch = 1 End if
  If NOT WinnerTwo = 1 OR Winner2 = 1 Then WinnerMatch = 0 End if
  HoldYourHorses = True
End Sub

Sub Horse3Check
  If HorseSpace3 < 7 Then Exit Sub End If
  If WinnerThree = 1 and HorseSpace3 >= 7 Then PlayerScores(1) = PlayerScores(1) + Purse:WinnerMatch = 1 End if
  If Winner3 = 1 and HorseSpace3 >= 7 Then PlayerScores(2) = PlayerScores(2) + Purse:WinnerMatch = 1 End if
  If NOT WinnerThree = 1 OR Winner3 = 1 Then WinnerMatch = 0 End if
  HoldYourHorses = True
End Sub

Sub Horse4Check
  If HorseSpace4 < 7 Then Exit Sub End If
  If WinnerFour = 1 and HorseSpace4 >= 7 Then PlayerScores(1) = PlayerScores(1) + Purse:WinnerMatch = 1 End if
  If Winner4 = 1 and HorseSpace4 >= 7 Then PlayerScores(2) = PlayerScores(2) + Purse:WinnerMatch = 1 End if
  If NOT WinnerFour = 1 OR Winner4 = 1 Then WinnerMatch = 0 End if
  HoldYourHorses = True
End Sub

Sub Horse5Check
  If HorseSpace5 < 7 Then Exit Sub End If
  If WinnerFive = 1 and HorseSpace5 >= 7 Then PlayerScores(1) = PlayerScores(1) + Purse:WinnerMatch = 1 End if
  If Winner5 = 1 and HorseSpace5 >= 7 Then PlayerScores(2) = PlayerScores(2) + Purse:WinnerMatch = 1 End if
  If NOT WinnerFive = 1 OR Winner5 = 1 Then WinnerMatch = 0 End if
  HoldYourHorses = True
End Sub

Sub Horse6Check
  If HorseSpace6 < 7 Then Exit Sub End If
  If WinnerSix = 1 and HorseSpace6 >= 7 Then PlayerScores(1) = PlayerScores(1) + Purse:WinnerMatch = 1 End if
  If Winner6 = 1 and HorseSpace6 >= 7 Then PlayerScores(2) = PlayerScores(2) + Purse:WinnerMatch = 1 End if
  If NOT WinnerSix = 1 OR Winner6 = 1 Then WinnerMatch = 0 End if
  HoldYourHorses = True
End Sub

Sub Bookie_Timer() 'Not 100% sure it's paying out correct amounts yet, but it *is* paying out
  Horse1Check
  Horse2Check
  Horse3Check
  Horse4Check
  Horse5Check
  Horse6Check
  If WinnerMatch = 0 Then exit sub end If
  Playsound "RaceOver"
  If Balls < 4 Then Credits = (Credits + 1) End if
  If Balls = 4 Then Credits = (Credits + 2) End If
  If Balls > 4 Then Credits = (Credits + 3) End If
  Controller.B2SSetCredits Credits
  If Credits > 0 Then CreditLight.State = 1
  If Credits = 0 Then CreditLight.State = 0 'not sure this line is super necessary here, but whatever
  Bookie.enabled = False
End Sub

'Ball in Play Numbers on Backglass

Sub UpdateBallinBallDisplay
  If BIPonBG = 1 then Controller.B2SSetData 141,1 end If
  If BIPonBG = 2 and Players = 1 then Controller.B2SSetData 141,0:Controller.B2SSetData 142,1 end If
  If BIPonBG = 3 and Players = 1 then Controller.B2SSetData 142,0:Controller.B2SSetData 143,1 end If
  If BIPonBG = 4 and Players = 1 then Controller.B2SSetData 143,0:Controller.B2SSetData 144,1 end If
  If BIPonBG = 5 and Players = 1 then Controller.B2SSetData 144,0:Controller.B2SSetData 145,1 end If

  If BIPonBG = 2 and Players = 2 then Controller.B2SSetData 141,1 end If
  If BIPonBG = 3 and Players = 2 then Controller.B2SSetData 141,0:Controller.B2SSetData 142,1 end If
  If BIPonBG = 4 and Players = 2 then Controller.B2SSetData 142,1 end If
  If BIPonBG = 5 and Players = 2 then Controller.B2SSetData 142,0:Controller.B2SSetData 143,1 end If
  If BIPonBG = 6 and Players = 2 then Controller.B2SSetData 143,1 end If
  If BIPonBG = 7 and Players = 2 then Controller.B2SSetData 143,0:Controller.B2SSetData 144,1 end If
  If BIPonBG = 8 and Players = 2 then Controller.B2SSetData 144,1 end If
  If BIPonBG = 9 and Players = 2 then Controller.B2SSetData 144,0:Controller.B2SSetData 145,1 end If
  If BIPonBG = 10 and Players = 2 then Controller.B2SSetData 145,1 end If
End Sub

'**********************
'My Scoring
'**********************

Sub LeftOutTrigger_Hit
  PlaySound "Bell"
  Scores Player,1000
End Sub

Sub RightOutTrigger_Hit
  PlaySound "Bell"
  Scores Player,1000
End Sub

Sub BLeft_Hit
  Playsound "Machine1"
  If  ABDouble = False Then Scores Player,50 end If
  If  ABDouble = True Then Scores Player,50 end If
End Sub

Sub DRight_Hit
  Playsound "Machine1"
  If  CDDouble = False Then Scores Player,50 end If
  If  CDDouble = True Then Scores Player,50 end If
End Sub

'Double Speed Turned On

Sub BackA_Hit
  Playsound "Hit1"
  Scores Player,500
  If ABDouble = False Then Exit Sub End If
  DoubleLight1.State = True
  DoubleLight3.State = True
  DoubleLight5.State = True
  DoubleSpace135 = True
  Scores Player,4500
End Sub

Sub BackB_Hit
  Playsound "Hit1"
  Scores Player,500
  If CDDouble = False Then Exit Sub End If
  DoubleLight2.State = True
  DoubleLight4.State = True
  DoubleLight6.State = True
  DoubleSpace246 = True
  Scores Player,4500
End Sub

'**********************
'End of game clean up
'**********************
Sub ResetGame
  BIPonBG = 0
  Controller.B2SSetData 141,0 'Ball in Play # on backglass lights
  Controller.B2SSetData 142,0
  Controller.B2SSetData 143,0
  Controller.B2SSetData 144,0
  Controller.B2SSetData 145,0

  Controller.B2SSetData 115,0 'Big numbers that declare the winning horse
  Controller.B2SSetData 116,0
  Controller.B2SSetData 117,0
  Controller.B2SSetData 118,0
  Controller.B2SSetData 119,0
  Controller.B2SSetData 120,0

' Controller.B2SSetData 101,0 '1 & 2 in the 1 or 2 can play oval
' Controller.B2SSetData 102,0

  If Credits > 1 Then CreditLight.State = 1:Controller.B2SSetData 101, 1:Controller.B2SSetData 102, 1 'State should be 1
  If Credits = 1 Then CreditLight.State = 1:Controller.B2SSetData 101, 1 'State should be 1
  If Credits = 0 Then CreditLight.State = 0 'State should be 0

  HorseReturn.Enabled = True

  HorseSpace1 = 0
  HorseSpace2 = 0
  HorseSpace3 = 0
  HorseSpace4 = 0
  HorseSpace5 = 0
  HorseSpace6 = 0

  LightA.state = 0
  LightB.state = 0
  LightC.state = 0
  LightD.state = 0

  DoubleSpace135 = False
  DoubleSpace246 = False

  ABDouble = False
  CDDouble = False

  Controller.B2SSetGameOver 0

  HoldYourHorses = False

  ClearLights

  Bookie.enabled = True

' WinningNumberLottery
End Sub

Sub ClearLights
  DoubleLight1.State = False
  DoubleLight2.State = False
  DoubleLight3.State = False
  DoubleLight4.State = False
  DoubleLight5.State = False
  DoubleLight6.State = False

  Controller.B2SSetData 121,0 'left horseshoe numbers
  Controller.B2SSetData 122,0
  Controller.B2SSetData 123,0
  Controller.B2SSetData 124,0
  Controller.B2SSetData 125,0
  Controller.B2SSetData 126,0

  Controller.B2SSetData 131,0 'right horseshoe numbers
  Controller.B2SSetData 132,0
  Controller.B2SSetData 133,0
  Controller.B2SSetData 134,0
  Controller.B2SSetData 135,0
  Controller.B2SSetData 136,0
End Sub

Sub CreditCheck_Timer()
  If Credits > 1 Then CreditLight.State = 1:Controller.B2SSetData 101, 1:Controller.B2SSetData 102, 1 'State should be 1
  If Credits = 1 Then CreditLight.State = 1:Controller.B2SSetData 101, 1 'State should be 1
  If Credits = 0 Then CreditLight.State = 0 'State should be 0
End Sub

Sub HorseReturn_Timer() 'Thanks to Pmax65 for this code
  if (Horse1.Transx < 0) Or (Horse2.Transx < 0) Or (Horse3.Transx < 0) Or (Horse4.Transx < 0) Or (Horse5.Transx < 0) Or (Horse6.Transx < 0) Then
    If Horse1.Transx < 0 then Horse1.Transx = Horse1.Transx + 50: else Horse1.Transx = 0: End If
    If Horse2.Transx < 0 then Horse2.Transx = Horse2.Transx + 50: else Horse2.Transx = 0: End If
    If Horse3.Transx < 0 then Horse3.Transx = Horse3.Transx + 50: else Horse3.Transx = 0: End If
    If Horse4.Transx < 0 then Horse4.Transx = Horse4.Transx + 50: else Horse4.Transx = 0: End If
    If Horse5.Transx < 0 then Horse5.Transx = Horse5.Transx + 50: else Horse5.Transx = 0: End If
    If Horse6.Transx < 0 then Horse6.Transx = Horse6.Transx + 50: else Horse6.Transx = 0: End If
    Playsound "HorsesReturning"
  Else
    HorseReturn.Enabled = False
  End If
End Sub

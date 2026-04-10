Option Explicit

'*****************************************************************************************************
' BunnyBoard
' Marble Games Co 1931
' IPD 407
' https://www.ipdb.org/machine.cgi?id=407
'
' Attribution
'
' History
' 1.0 - initial release

'*****************************************************************************************************

On Error Resume Next
ExecuteGlobal GetTextFile("Controller.vbs")
If Err Then MsgBox "Unable to open Controller.vbs. Ensure that it is in the Scripts folder of Visual Pinball."
On Error Goto 0

Const cGameName = "Bunnyboard"
Const BallSize = 18
Dim UseDebugWindow : UseDebugWindow = false

Dim Controller  ' B2S
Dim B2SScore  ' B2S Score Displayed
Dim UseB2S : UseB2S = 1

Dim IsFirstLoad : IsFirstLoad = True
'Dim IsGameStarting : IsGameStarting = False
Dim IsGameOver : IsGameOver = False
Dim IsGameReady : IsGameReady = False

Dim EnableRetractPlunger
EnableRetractPlunger = false 'Change to true to enable retracting the plunger at a linear speed; wait for 1 second at the maximum position; move back towards the resting position; nice for button/key plungers

' Tilt not implemented on this table
Dim TiltWarnings : TiltWarnings = 1
Dim CurrentWarnings : CurrentWarnings = 0
Dim IsTilted : IsTilted = False
Dim TiltWarningTime
Dim tiltball

Dim BallsPerGame : BallsPerGame = 7
Dim Score, BIP, BallsPlayed, HighScore, DefaultHighScore

Dim EnableBallControl : EnableBallControl = false 'Change to true to enable manual ball control (or press C in-game) via the arrow keys and B (boost movement) keys
Dim LightLevel : LightLevel = 0.7
Dim StandardBallBrightness : StandardBallBrightness = 255 ' this is overriden by day/night setting in table

' for resetting high score: hold start key for 5+ seconds
Dim StartKeyBeginTime
Dim StartKeyEndTime

Sub Table1_Exit()
  Controller.Stop
End Sub

Sub Table1_Init()
  BIP = 0
  BallsPlayed = 0
  Score = 0
  HighScore = 0
  DefaultHighScore = 150
  ResetBallCounts

  LoadEM
  Debug.print "init"
  Randomize

  HighScore=LoadValue(cGameName,"HighScore")
  If HighScore="" then
    HighScore=DefaultHighScore
  Else
    HighScore=Cdbl(HighScore)
  End If

  B2SLoadedTimer.Enabled = True
  B2S_SetGameOver

  SetRampsLow

  CreateInitialBalls


  Dim wrapLighting, roughness, glossyImageLerp, thickness, edge, edgeAlpha, opacity, base, glossy, clearcoat, isMetal, opacityActive, elasticity, elasticityFalloff, friction, scatterAngle
  GetMaterial "VLM.Bake.Solid_T", wrapLighting, roughness, glossyImageLerp, thickness, edge, edgeAlpha, opacity, base, glossy, clearcoat, isMetal, opacityActive, elasticity, elasticityFalloff, friction, scatterAngle


' UpdateMaterial "VLM.Bake.Solid_T",0,0,1,0, 0,0,0, RGB(127,127,127),RGB(0,0,0),RGB(0,0,0), False,True,0,0,0,0

  UpdateMaterial "VLM.Bake.Solid_T", wrapLighting, roughness, glossyImageLerp, thickness, edge, edgeAlpha, 0, base, glossy, clearcoat, isMetal, opacityActive, elasticity, elasticityFalloff, friction, scatterAngle

  SaveMtlColors
End Sub

Sub Table1_KeyDown(ByVal keycode)

  If keycode = PlungerKey Then
        If EnableRetractPlunger Then
            Plunger.PullBackandRetract
        Else
        Plunger.PullBack
        End If
    PlaySound "plungerpull",0,1,AudioPan(Plunger),0.25,0,0,1,AudioFade(Plunger)
  End If

  if keycode = MechanicalTilt then
    'SetMechTilt
  end If

  if keycode = StartGameKey Then
    StartKeyBeginTime = Now
  end If

  If (keycode = StartGameKey and IsGameOver = True) Then
    IsGameOver = False
    IsGameReady = False
    BallsPlayed = 0
    BIP = 0
    Score = 0
    IsTilted = False

    StartGameTimer.enabled = True

    B2S_GameStart
    B2S_UpdateHighScoreReel
  End If

  if keycode = RightMagnaSave Then
    'SaveMtlColors
    LowerRamps
  end if

  if keycode = LeftMagnaSave Then
    RaiseRamps
  end if

  If keycode = RightFlipperKey and not IsGameOver and IsGameReady and _
    BallsPlayed < BallsPerGame and not IsShooterLaneFull Then

    PlaySound SoundFX("fx_kicker_enter",DOFContactors), 0,1,AudioPan(BallRelease),0.25,0,0,1,AudioFade(BallRelease)

    CreateNewBall BallRelease
    BallRelease.Kick 90, 5

    AddBIP(1)
    BallsPlayed = BallsPlayed + 1


    B2S_UpdateBallReel
    'Debug.print "RightFlip: BIP:" & BIP & " Played:" & BallsPlayed
  End If

  If keycode = LeftTiltKey Then
    Nudge 90, 2
    'CheckManualTilt 90
  End If

  If keycode = RightTiltKey Then
    Nudge 270, 2
    'CheckManualTilt 270
  End If

  If keycode = CenterTiltKey Then
    Nudge 0, 2
    'CheckManualTilt 0
  End If

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

Sub Table1_KeyUp(ByVal keycode)
  If keycode = PlungerKey Then
    Plunger.Fire
    PlaySound "plunger",0,1,AudioPan(Plunger),0.25,0,0,1,AudioFade(Plunger)
  End If

  If keycode = LeftFlipperKey Then

  End If

  If keycode = RightFlipperKey Then

  End If

  if keycode = StartGameKey Then
    StartKeyEndTime = Now

    if DateDiff("s",StartKeyBeginTime,StartKeyEndTime) >= 5 Then
      ResetHighScore
    end If
  end If

    'Manual Ball Control
  If EnableBallControl = 1 Then
    If keycode = 203 Then BCleft = 0  ' Left Arrow
    If keycode = 200 Then BCup = 0    ' Up Arrow
    If keycode = 208 Then BCdown = 0  ' Down Arrow
    If keycode = 205 Then BCright = 0 ' Right Arrow
  End If

End Sub

sub CreateInitialBalls
  'create 2 balls in the lower ring, the rest above the drain

  BallCreateTimer.UserValue = BallsPerGame
  BallCreateTimer.enabled=True

end Sub

Sub BallCreateTimer_Timer()
  dim newball

  BallCreateTimer.UserValue = BallCreateTimer.UserValue - 1

  if BallCreateTimer.UserValue >4 then
    CreateNewBall BallCreate2
    BallCreate2.Kick 100,2
  else
    CreateNewBall BallCreate1
    BallCreate1.Kick 100,2
  end if

  if BallCreateTimer.UserValue = 0 then
    BallCreateTimer.enabled = false
  end if

End Sub



' ***** SHOOTER LANE CONTROL *****

dim shooterx1 : shooterx1 = 843
dim shooterx2: shooterx2 = 900
dim shootery1 : shootery1 = 540
dim shootery2 : shootery2 = 1700
dim MaxBallsInShooterLane: MaxBallsInShooterLane = 2

function BallsInShooterLane

  dim BOT, b, ballsInShooter
  ballsInShooter = 0

  BOT = GetBalls
  For b = 0 to UBound(BOT)
    'define rough area of shooter lane
    if BOT(b).X > shooterx1 and BOT(b).X < shooterx2 and BOT(b).Y > shootery1 and BOT(b).Y < shootery2 Then
      ballsInShooter = ballsInShooter + 1
    end If
    Next

  BallsInShooterLane = ballsInShooter
end Function

function IsShooterLaneFull

  IsShooterLaneFull = False

  if BallsInShooterLane >= MaxBallsInShooterLane Then
    IsShooterLaneFull = True
  end If

end Function

' ***** END SHOOTER LANE CONTROL *****


'*****************
' Maths
'*****************

Dim PI
PI = 4 * Atn(1)

Function dSin(degrees)
  dsin = Sin(degrees * Pi / 180)
End Function

Function dCos(degrees)
  dcos = Cos(degrees * Pi / 180)
End Function

Function Atn2(dy, dx)
  If dx > 0 Then
    Atn2 = Atn(dy / dx)
  ElseIf dx < 0 Then
    If dy = 0 Then
      Atn2 = pi
    Else
      Atn2 = Sgn(dy) * (pi - Atn(Abs(dy / dx)))
    End If
  ElseIf dx = 0 Then
    If dy = 0 Then
      Atn2 = 0
    Else
      Atn2 = Sgn(dy) * pi / 2
    End If
  End If
End Function

Function max(a,b)
  If a > b Then
    max = a
  Else
    max = b
  End If
End Function

Function min(a,b)
  If a > b Then
    min = b
  Else
    min = a
  End If
End Function

Function RndInt(min, max)
    RndInt = Int(Rnd() * (max-min + 1) + min)' Sets a random number integer between min and max
End Function

' RAMPS

dim ramp, pointTrigger
dim IsRampUp : IsRampUp = True ' default state is up

Dim wrapLighting, roughness, glossyImageLerp, thickness, edge, edgeAlpha, opacity, base, glossy, clearcoat, isMetal, opacityActive, elasticity, elasticityFalloff, friction, scatterAngle

sub RaiseRamps

  'Dim wrapLighting, roughness, glossyImageLerp, thickness, edge, edgeAlpha, opacity, base, glossy, clearcoat, isMetal, opacityActive, elasticity, elasticityFalloff, friction, scatterAngle
  GetMaterial "VLM.Bake.Solid_T", wrapLighting, roughness, glossyImageLerp, thickness, edge, edgeAlpha, opacity, base, glossy, clearcoat, isMetal, opacityActive, elasticity, elasticityFalloff, friction, scatterAngle

  if not IsRampUp Then
    RampUpTimer.enabled = True
  end If

  for each pointTrigger in PointsTriggers
    pointTrigger.enabled = False
  Next

end Sub

sub LowerRamps

  if IsRampUp Then
    RampDownTimer.enabled = True
  end If

  for each pointTrigger in PointsTriggers
    pointTrigger.enabled = True
  Next

end Sub

dim curZ : curZ = -22
dim shadowOpacity : shadowOpacity = 0
dim shadowStepSize : shadowStepSize = 12
Sub RampUpTimer_Timer

  curZ = curZ + 4
  for each ramp in AllRamps
    ramp.TransZ = -curZ
  Next

  BM_center_front.transx = -curZ
  BM_center_back.transz = curZ + 8

  'for each ramp in CentralRamps
  ' ramp.TransX = -curZ
  'Next

  shadowOpacity = shadowOpacity + shadowStepSize
  if shadowOpacity > 100 then shadowOpacity = 100
  'debug.print shadowOpacity

  'UpdateMaterial BM_Playfield_s.material,0,0,1,0, 0,0,cDbl(shadowOpacity/100), RGB(127,127,127),RGB(0,0,0),RGB(0,0,0), False,True,0,0,0,0
  UpdateMaterial "VLM.Bake.Solid_T", wrapLighting, roughness, glossyImageLerp, thickness, edge, edgeAlpha, cDbl(shadowOpacity/100), base, glossy, clearcoat, isMetal, opacityActive, elasticity, elasticityFalloff, friction, scatterAngle


  if curZ >= 0 Then
    for each ramp in RampWalls
      ramp.Collidable = False
    Next

    IsRampUp = True
    'BM_Playfield_S.opacity = 1
    RampUpTimer.Enabled = False
  End If

End Sub

Sub RampDownTimer_Timer

  curZ = curZ - 4
  for each ramp in AllRamps
    ramp.TransZ = -curZ
  Next

  BM_center_front.transx = -curZ
  BM_center_back.transz = curZ + 8

  'for each ramp in CentralRamps
  ' ramp.TransX = -curZ
  'Next

  shadowOpacity = shadowOpacity - shadowStepSize
  if shadowOpacity < 0 then shadowOpacity = 100
  'debug.print shadowOpacity
  'UpdateMaterial BM_Playfield_s.material,0,0,1,0, 0,0,cDbl(shadowOpacity/100), RGB(127,127,127),RGB(0,0,0),RGB(0,0,0), False,True,0,0,0,0
  UpdateMaterial "VLM.Bake.Solid_T", wrapLighting, roughness, glossyImageLerp, thickness, edge, edgeAlpha, cDbl(shadowOpacity/100), base, glossy, clearcoat, isMetal, opacityActive, elasticity, elasticityFalloff, friction, scatterAngle


  if curZ < LowerRampZ+3 Then
    for each ramp in RampWalls
      ramp.Collidable = True
    Next
  End If

  if curZ <= LowerRampZ Then
    IsRampUp = False
    RampDownTimer.Enabled = False
    'UpdateMaterial BM_Playfield_s.material,0,0,1,0, 0,0,0, RGB(127,127,127),RGB(0,0,0),RGB(0,0,0), False,True,0,0,0,0
    UpdateMaterial "VLM.Bake.Solid_T", wrapLighting, roughness, glossyImageLerp, thickness, edge, edgeAlpha, 0, base, glossy, clearcoat, isMetal, opacityActive, elasticity, elasticityFalloff, friction, scatterAngle

  End If

End Sub


Const LowerRampZ = -22
Sub SetRampsLow

  BM_center_front.transx = -LowerRampZ
  BM_center_back.transz = LowerRampZ + 8 ' hack to fine tune height

  for each ramp in AllRamps
    ramp.TransZ = -LowerRampZ
  Next

  IsRampUp = False
end Sub

Sub SetRampsHigh

  BM_center_front.transx = 0
  BM_center_back.transz = 0 + 3 ' hack to fine tune height

  for each ramp in AllRamps
    ramp.TransZ = 0
  Next

  IsRampUp = False
end Sub


' END RAMPS

dim startGameTimerIndex : startGameTimerIndex = 0
sub StartGameTimer_Timer
  dim trigger
  if startGameTimerIndex=0 Then
    Debug.print "Start game timer enabled"
    PlaySound "coinslidein"
    BIP = BallsPerGame

    for each trigger in AllTriggers
      trigger.enabled = false
    Next

    RaiseRamps
  end If

  startGameTimerIndex = startGameTimerIndex + 1

  if BIP = 0 Then ' all balls drained
    PlaySound "coinslideout"
    ResetBallCounts
    StartGameTimer.enabled = False
    LowerRamps
    IsGameReady = True
    startGameTimerIndex = 0
    B2S_GameStart

    for each trigger in AllTriggers
      trigger.enabled = true
    Next
  end If

end Sub

sub ResetBallCounts
  dim i

  for i = 0 to ubound(BallCount,1)
    BallCount(i,1) = 0
  next

end Sub

Sub CheckHighScore
  If Score>HighScore and not IsTilted then
    HighScore=Score
      SaveValue cGameName, "HighScore", HighScore
    B2S_UpdateHighScoreReel
  End If
End Sub

Sub ResetHighScore

  HighScore = DefaultHighScore
  SaveValue cGameName, "HighScore", HighScore
  B2S_UpdateHighScoreReel

End Sub


Sub BallRelease_Hit()
  BallRelease.DestroyBall

  AddBIP(-1)
End Sub


' ***** Test *****

sub AddTestBalls

  CreateNewBall TestKicker000
  TestKicker000.Kick 180,1

  CreateNewBall TestKicker005
  TestKicker005.Kick 110,4

  CreateNewBall TestKicker001
  TestKicker001.Kick 180,1

  CreateNewBall TestKicker002
  TestKicker002.Kick 180,1

  CreateNewBall TestKicker003
  TestKicker003.Kick 180,1

  CreateNewBall TestKicker004
  TestKicker004.Kick 180,1

end Sub


' ***** End Test *****


' ***** Triggers *****

sub Triggers_Bottom_Hit(idx)
  Debug.print "Bottom trigger hit"
  AddBIP(-1)
  EndCheck
end Sub

sub Triggers_50_Hit(idx)
  Debug.print "50 hit"
  AddScore(50)
  AddBIP(-1)
  EndCheck
end Sub

sub Triggers_75_Hit(idx)
  Debug.print "75 hit"
  AddScore(75)
  AddBIP(-1)
  EndCheck
end Sub

sub Triggers_100_Hit(idx)
  Debug.print "100 hit"
  AddScore(100)
  AddBIP(-1)
  EndCheck
end Sub

sub Triggers_150_Hit(idx)
  Debug.print "150 hit"
  AddScore(150)
  AddBIP(-1)
  EndCheck
end Sub

sub AddBIP(val)
  BIP = BIP + val
  if BIP < 0 then BIP = 0
  debug.print "BIP: " & BIP
end Sub

Sub Plunger_Init()
End Sub

'*********************************************************************
'                 Positional Sound Playback Functions
'*********************************************************************

' Play a sound, depending on the X,Y position of the table element (especially cool for surround speaker setups, otherwise stereo panning only)
' parameters (defaults): loopcount (1), volume (1), randompitch (0), pitch (0), useexisting (0), restart (1))
' Note that this will not work (currently) for walls/slingshots as these do not feature a simple, single X,Y position
Sub PlayXYSound(soundname, tableobj, loopcount, volume, randompitch, pitch, useexisting, restart)
  PlaySound soundname, loopcount, volume, AudioPan(tableobj), randompitch, pitch, useexisting, restart, AudioFade(tableobj)
End Sub

' Similar subroutines that are less complicated to use (e.g. simply use standard parameters for the PlaySound call)
Sub PlaySoundAt(soundname, tableobj)
    PlaySound soundname, 1, 1, AudioPan(tableobj), 0,0,0, 1, AudioFade(tableobj)
End Sub

Sub PlaySoundAtBall(soundname)
    PlaySoundAt soundname, ActiveBall
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
'   rothbauerw's Manual Ball Control
'*****************************************

Dim BCup, BCdown, BCleft, BCright
Dim ControlBallInPlay, ControlActiveBall
Dim BCvel, BCyveloffset, BCboostmulti, BCboost

BCboost = 1       'Do Not Change - default setting
BCvel = 4       'Controls the speed of the ball movement
BCyveloffset = -0.01  'Offsets the force of gravity to keep the ball from drifting vertically on the table, should be negative
BCboostmulti = 3    'Boost multiplier to ball veloctiy (toggled with the B key)

ControlBallInPlay = false

Sub StartBallControl_Hit()
  Set ControlActiveBall = ActiveBall
  ControlBallInPlay = true
End Sub

Sub StopBallControl_Hit()
  ControlBallInPlay = false
End Sub

Sub BallControlTimer_Timer()
  If EnableBallControl and ControlBallInPlay then
    If BCright = 1 Then
      ControlActiveBall.velx =  BCvel*BCboost
    ElseIf BCleft = 1 Then
      ControlActiveBall.velx = -BCvel*BCboost
    Else
      ControlActiveBall.velx = 0
    End If

    If BCup = 1 Then
      ControlActiveBall.vely = -BCvel*BCboost
    ElseIf BCdown = 1 Then
      ControlActiveBall.vely =  BCvel*BCboost
    Else
      ControlActiveBall.vely = bcyveloffset
    End If
  End If
End Sub


'********************************************************************
'      JP's VP10 Rolling Sounds (+rothbauerw's Dropping Sounds)
'********************************************************************

Const tnob = 10 ' total number of balls
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

  If UBound(BOT) > tnob Then Exit Sub

    For b = 0 to UBound(BOT)
        ' play the rolling sound for each ball
        If BallVel(BOT(b) ) > 1 AND BOT(b).z < 30 Then
            rolling(b) = True
            PlaySound("fx_ballrolling" & b), -1, Vol(BOT(b)), AudioPan(BOT(b)), 0, Pitch(BOT(b)), 1, 0, AudioFade(BOT(b))
        Else
            If rolling(b) = True Then
                StopSound("fx_ballrolling" & b)
                rolling(b) = False
            End If
        End If

        ' play ball drop sounds
        'If BOT(b).VelZ < -1 and BOT(b).z < 55 and BOT(b).z > 1 Then 'height adjust for ball drop sounds
       '     PlaySound "fx_ball_drop" & b, 0, ABS(BOT(b).velz)/1200, AudioPan(BOT(b)), 0, Pitch(BOT(b)), 1, 0, AudioFade(BOT(b))
       ' End If
    Next
End Sub

'**********************
' Ball Collision Sound
'**********************

Sub OnBallBallCollision(ball1, ball2, velocity)
  PlaySound("fx_collide"), 0, Csng(velocity) ^2 / 2000, AudioPan(ball1), 0, Pitch(ball1), 0, 0, AudioFade(ball1)
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

Sub BumperWall_Hit
  PlaySound "metalhit2", 0, Vol(ActiveBall), AudioPan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
End Sub

Sub ContactWalls_Hit(idx)
  PlaySound "metalhit_thin", 0, Vol(ActiveBall), AudioPan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
End Sub

'Sub ApronTop_Hit
' PlaySound "metalhit_thin", 0, Vol(ActiveBall), AudioPan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
'End Sub

'Sub ApronBottom_Hit
' PlaySound "metalhit_thin", 0, Vol(ActiveBall), AudioPan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
'End Sub

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

Sub Pins_Hit (idx)
  'PlaySound(string, int loopcount, float volume, float pan, float randompitch, int pitch, bool useexisting, bool restart, float front_rear_fade)
  PlaySound "pintap3", 0, PinTapVol(ActiveBall), AudioPan(ActiveBall), 0.1, Pitch(ActiveBall), 0, 0, AudioFade(ActiveBall)
End Sub

Sub RampWalls_Hit (idx)
  'debug.print "RampWalls_Hit " & idx
  'PlaySound(string, int loopcount, float volume, float pan, float randompitch, int pitch, bool useexisting, bool restart, float front_rear_fade)
  PlaySound "pintap3", 0, PinTapVol(ActiveBall), AudioPan(ActiveBall), 0.1, 0.6, 0, 0, AudioFade(ActiveBall)
End Sub

Function PinTapVol(ball) ' Calculates the Volume of the sound based on the ball speed
    PinTapVol = Csng(BallVel(ball) ^2 / 60)
End Function

Sub ApronTopRail_Hit
  'debug.print "ApronTopRail_Hit"
  'PlaySound(string, int loopcount, float volume, float pan, float randompitch, int pitch, bool useexisting, bool restart, float front_rear_fade)
  'PlaySound "RampLoop"&RndInt(1,1), 2, PinTapVol(ActiveBall), AudioPan(ActiveBall), 0.1, 0.6, 0, 0, AudioFade(ActiveBall)
End Sub

Sub ApronTopRail_Unhit()
  'debug.print "ApronTopRail_Unhit"
  'PlaySound(string, int loopcount, float volume, float pan, float randompitch, int pitch, bool useexisting, bool restart, float front_rear_fade)
  'PlaySound "pintap3", 0, PinTapVol(ActiveBall), AudioPan(ActiveBall), 0.1, 0.6, 0, 0, AudioFade(ActiveBall)
  'StopSound "RampLoop1"
  'StopSound "Arch_L2"
  'StopSound "Arch_L3"
  'StopSound "Arch_L4"
End Sub

Sub ApronBottom_Hit
  'PlaySound(string, int loopcount, float volume, float pan, float randompitch, int pitch, bool useexisting, bool restart, float front_rear_fade)
  PlaySound "Apron_Soft_" & RndInt(1,7), 0, PinTapVol(ActiveBall), AudioPan(ActiveBall), 0.1, 1, 0, 0, AudioFade(ActiveBall)

End Sub

Sub WoodRails_Hit (idx)
  'debug.print "WoodRails_Hit"
  'PlaySound(string, int loopcount, float volume, float pan, float randompitch, int pitch, bool useexisting, bool restart, float front_rear_fade)
  PlaySound "Wall_Hit_8", 0, PincTapVol(ActiveBall), AudioPan(ActiveBall), 0.1, 0.8, 0, 0, AudioFade(ActiveBall)
End Sub

Sub WoodWalls_Hit (idx)
  'debug.print "WoodWalls hit"
  'PlaySound(string, int loopcount, float volume, float pan, float randompitch, int pitch, bool useexisting, bool restart, float front_rear_fade)
  PlaySound "Wall_Hit_8", 0, PinTapVol(ActiveBall)*.6, AudioPan(ActiveBall), 0.3, 1.9, 0, 0, AudioFade(ActiveBall)
End Sub

Sub BumperTrigger_Hit
  'debug.print "Bumper hit"
  'PlaySound(string, int loopcount, float volume, float pan, float randompitch, int pitch, bool useexisting, bool restart, float front_rear_fade)
  PlaySound "Wall_Hit_8", 0, PinTapVol(ActiveBall)*.15, AudioPan(ActiveBall), 0.1, 0.8, 0, 0, AudioFade(ActiveBall)
End Sub

Sub BumperFlipper_Hit
  'debug.print "Flipper hit"
  'PlaySound(string, int loopcount, float volume, float pan, float randompitch, int pitch, bool useexisting, bool restart, float front_rear_fade)
  PlaySound "Wall_Hit_8", 0, PinTapVol(ActiveBall), AudioPan(ActiveBall), 0.1, 0.8, 0, 0, AudioFade(ActiveBall)
End Sub



' ** Ramp Rolling Sounds

Sub RailTrigger_Hit
  if IsRampUp then exit Sub
  'debug.print "RailTrigger_Hit"
  WireRampOn True
  'PlaySound(string, int loopcount, float volume, float pan, float randompitch, int pitch, bool useexisting, bool restart, float front_rear_fade)
  'PlaySound "RampLoop"&RndInt(1,10), 2, .002, AudioPan(ActiveBall), 0.1, 1.2, 0, 0, AudioFade(ActiveBall)
End Sub

Sub RailTrigger_Unhit
  'debug.print "RailTrigger_Unhit"
  WireRampOff
  'PlaySound(string, int loopcount, float volume, float pan, float randompitch, int pitch, bool useexisting, bool restart, float front_rear_fade)
  'dim i
  'for i = 1 to 10
  ' StopSound "RampLoop"&i
  'next
End Sub

Sub zzRampTrigger_Drain_Hit
  if IsRampUp then exit Sub
  'debug.print "RampTrigger_Drain_Hit"
  WireRampOn False
End Sub

Sub zzRampTrigger_Drain_Unhit
  'debug.print "RampTrigger_Drain_Unhit"
  WireRampOff
End Sub

Sub RampTriggers_Hit(idx)
  if IsRampUp then exit Sub
  'debug.print "RampTriggers_Hit"
  WireRampOn False
End Sub

Sub RampTriggers_Unhit(idx)
  'debug.print "RampTriggers_Unhit"
  WireRampOff
End Sub


' ** End Ramp Rolling Sounds



'******************************************************
'**** RAMP ROLLING SFX
'******************************************************

Dim ApronAdjustment : ApronAdjustment = 0.3 'Adjust for apron rail to lower volume
Dim VolumeDial : VolumeDial = 0.9     'Overall Mechanical sound effect volume. Recommended values should be no greater than 1.
Dim RampRollVolume : RampRollVolume = 1   'Level of ramp rolling volume. Value between 0 and 1

'Ball tracking ramp SFX 1.0
'   Reqirements:
'          * Import A Sound File for each ball on the table for plastic ramps.  Call It RampLoop<Ball_Number> ex: RampLoop1, RampLoop2, ...
'          * Import a Sound File for each ball on the table for wire ramps. Call it WireLoop<Ball_Number> ex: WireLoop1, WireLoop2, ...
'          * Create a Timer called RampRoll, that is enabled, with a interval of 100
'          * Set RampBAlls and RampType variable to Total Number of Balls
' Usage:
'          * Setup hit events and call WireRampOn True or WireRampOn False (True: lower volume, False: normal volume)
'          * To stop tracking ball
'                 * call WireRampOff
'                 * Otherwise, the ball will auto remove if it's below 30 vp units
'

dim RampMinLoops : RampMinLoops = 4

' RampBalls
'      Setup:        Set the array length of x in RampBalls(x,2) Total Number of Balls on table + 1:  if tnob = 5, then RammBalls(6,2)
'      Description:
dim RampBalls(8,2)
'x,0 = ball x,1 = ID, 2 = Protection against ending early (minimum amount of updates)
'0,0 is boolean on/off, 0,1 unused for now
RampBalls(0,0) = False

' RampType
'     Setup: Set this array to the number Total number of balls that can be tracked at one time + 1.  5 ball multiball then set value to 6
'     Description: Array type indexed on BallId and a values used to deterimine what type of ramp the ball is on: False = Wire Ramp, True = Plastic Ramp
dim RampType(8)

Sub WireRampOn(input)  : Waddball ActiveBall, input : RampRollUpdate: End Sub
Sub WireRampOff() : WRemoveBall ActiveBall.ID : End Sub


' WaddBall (Active Ball, Boolean)
'     Description: This subroutine is called from WireRampOn to Add Balls to the RampBalls Array
Sub Waddball(input, RampInput)  'Add ball
  ' This will loop through the RampBalls array checking each element of the array x, position 1
  ' To see if the the ball was already added to the array.
  ' If the ball is found then exit the subroutine
  dim x : for x = 1 to uBound(RampBalls)  'Check, don't add balls twice
    if RampBalls(x, 1) = input.id then
      if Not IsEmpty(RampBalls(x,1) ) then Exit Sub 'Frustating issue with BallId 0. Empty variable = 0
    End If
  Next

  ' This will iterate through the RampBalls Array.
  ' The first time it comes to a element in the array where the Ball Id (Slot 1) is empty.  It will add the current ball to the array
  ' The RampBalls assigns the ActiveBall to element x,0 and ball id of ActiveBall to 0,1
  ' The RampType(BallId) is set to RampInput
  ' RampBalls in 0,0 is set to True, this will enable the timer and the timer is also turned on
  For x = 1 to uBound(RampBalls)
    if IsEmpty(RampBalls(x, 1)) then
      Set RampBalls(x, 0) = input
      RampBalls(x, 1) = input.ID
      RampType(x) = RampInput
      RampBalls(x, 2) = 0
      'exit For
      RampBalls(0,0) = True
      RampRoll.Enabled = 1   'Turn on timer
      'RampRoll.Interval = RampRoll.Interval 'reset timer
      exit Sub
    End If
    if x = uBound(RampBalls) then   'debug
      Debug.print "WireRampOn error, ball queue is full: " & vbnewline & _
      RampBalls(0, 0) & vbnewline & _
      Typename(RampBalls(1, 0)) & " ID:" & RampBalls(1, 1) & "type:" & RampType(1) & vbnewline & _
      Typename(RampBalls(2, 0)) & " ID:" & RampBalls(2, 1) & "type:" & RampType(2) & vbnewline & _
      Typename(RampBalls(3, 0)) & " ID:" & RampBalls(3, 1) & "type:" & RampType(3) & vbnewline & _
      Typename(RampBalls(4, 0)) & " ID:" & RampBalls(4, 1) & "type:" & RampType(4) & vbnewline & _
      Typename(RampBalls(5, 0)) & " ID:" & RampBalls(5, 1) & "type:" & RampType(5) & vbnewline & _
      " "
    End If
  next
End Sub

' WRemoveBall (BallId)
'    Description: This subroutine is called from the RampRollUpdate subroutine
'                 and is used to remove and stop the ball rolling sounds
Sub WRemoveBall(ID)   'Remove ball
  'Debug.Print "In WRemoveBall() + Remove ball from loop array"
  dim ballcount : ballcount = 0
  dim x : for x = 1 to Ubound(RampBalls)
    if ID = RampBalls(x, 1) then 'remove ball
      Set RampBalls(x, 0) = Nothing
      RampBalls(x, 1) = Empty
      RampType(x) = Empty
      StopSound("wireloop" & x)
    end If
    'if RampBalls(x,1) = Not IsEmpty(Rampballs(x,1) then ballcount = ballcount + 1
    if not IsEmpty(Rampballs(x,1)) then ballcount = ballcount + 1
  next
  if BallCount = 0 then RampBalls(0,0) = False  'if no balls in queue, disable timer update
End Sub

Sub RampRoll_Timer():RampRollUpdate:End Sub

Sub RampRollUpdate()    'Timer update
  dim x : for x = 1 to uBound(RampBalls)
    if Not IsEmpty(RampBalls(x,1) ) then
      if BallVel(RampBalls(x,0) ) > 1 then ' if ball is moving, play rolling sound
        If RampType(x) then
          PlaySound("wireloop" & x), -1, VolPlayfieldRoll(RampBalls(x,0)) * RampRollVolume * VolumeDial * ApronAdjustment, AudioPan(RampBalls(x,0)), 0, BallPitchV(RampBalls(x,0)), 1, 0, AudioFade(RampBalls(x,0))
        Else
          PlaySound("wireloop" & x), -1, VolPlayfieldRoll(RampBalls(x,0)) * RampRollVolume * VolumeDial, AudioPan(RampBalls(x,0)), 0, BallPitch(RampBalls(x,0)), 1, 0, AudioFade(RampBalls(x,0))
        End If
        RampBalls(x, 2) = RampBalls(x, 2) + 1
      Else
        StopSound("wireloop" & x)
      end if
      if RampBalls(x,0).Z < 30 and RampBalls(x, 2) > RampMinLoops then  'if ball is on the PF, remove  it
        StopSound("wireloop" & x)
        Wremoveball RampBalls(x,1)
      End If
    Else
      StopSound("wireloop" & x)
    end if
  next
  if not RampBalls(0,0) then RampRoll.enabled = 0

End Sub

dim RollingSoundFactor : RollingSoundFactor = 1.1/5

Function VolPlayfieldRoll(ball) ' Calculates the roll volume of the sound based on the ball speed
    VolPlayfieldRoll = RollingSoundFactor * 0.0005 * Csng(BallVel(ball) ^3)
End Function


Function BallPitch(ball) ' Calculates the pitch of the sound based on the ball speed
    BallPitch = pSlope(BallVel(ball), 1, -1000, 60, 10000)
End Function

Function BallPitchV(ball) ' Calculates the pitch of the sound based on the ball speed Variation
  BallPitchV = pSlope(BallVel(ball), 1, -4000, 60, 7000)
End Function

' Used for flipper correction and rubber dampeners
Function PSlope(Input, X1, Y1, X2, Y2)    'Set up line via two points, no clamping. Input X, output Y
  Dim x, y, b, m
  x = input
  m = (Y2 - Y1) / (X2 - X1)
  b = Y2 - m*X2
  Y = M*x+b
  PSlope = Y
End Function

'******************************************************



' ***** GAME LOGIC *****

dim index : index = 0
dim pfball
dim specialball

dim BallCount(10,1) ' keeps track of which ball index to use

dim currentBallCount : currentBallCount = 0
sub CreateNewBall(location)

  dim imageId, tempId ': imageId = ballimageid

  if ballimageid = 6 then
    tempId = GetRandomBallImageId
  else
    tempId = ballimageid
  end if

  Set pfball = location.CreateBall
  pfball.radius = BallSize
  pfball.mass = 1
  pfball.ForceReflection = True
  pfball.UserValue = tempId

  currentBallCount = currentBallCount + 1

  pfball.FrontDecal = GetNextBallName(tempId)
  SetBallColor pfball, tempId

end sub

function GetNextBallName(id)

  if id=1 then ' mashup, always pick a new one
    id = GetRandomBallImageId
  end if

  BallCount(id, 1) = BallCount(id, 1) + 1
  if BallCount(id, 1) > 7 then BallCount(id, 1) = 1

  select case id
    case 3:
      GetNextBallName = "ball algae"
    case 4:
      GetNextBallName = "ball bluewave"
    case 5:
      GetNextBallName = "ball fire"
    case 6:
      GetNextBallName = "ball gold"
    case 7:
      GetNextBallName = "ball violet"
    case 8:
      GetNextBallName = "ball rainbow"
    case else:
      GetNextBallName = "ball0"
      exit function
  end select

  GetNextBallName = GetNextBallName & " " & BallCount(id, 1)
  'GetNextBallName = "ball rainbow 1"
end function


Sub AddScore (points)

  if not IsGameReady then
    Debug.print "Points when not ready:" & points
    'ignore score updates when resetting
    exit Sub
  end If

  Debug.print "Score modified: " & points
  if Not(IsTilted) Then
    Score = Score + points
    B2S_UpdateScoreReel
  End If

End Sub

Sub EndCheck

  if BallsPlayed >= BallsPerGame and BIP <= 0 Then
    IsGameOver = True
    IsGameReady = False
  end If

  if IsGameOver Then
    B2S_SetGameOver
    CheckHighScore
  end If

End Sub

' ***** END GAME LOGIC *****

' ***** TEST *****

dim alternate : alternate = False
Sub TestTimer_Timer
  Dim ballAngle, ballPower
  alternate = not alternate

  if alternate then
    ballAngle = Int((180-1+1)*Rnd+1)
  Else
    ballAngle = Int((360-180+1)*Rnd+180)
  end If

  'ballAngle = 45

  ballPower = Int((9-2+1)*Rnd+2)

  CreateNewBall TestKicker
  TestKicker.kick ballAngle, ballPower
End Sub

Sub RunSim

  if TestTimer.Enabled = True Then
    DisableKickers
    'ClearAllKickers
    'FillAllKickers
    'TestGate.Collidable = True
    TestKicker.enabled = False
    TestTimer.enabled = False
    TestDrain.Enabled = False
  Else
    DisableKickers
    'ClearAllKickers
    'FillAllKickers
    'TestGate.Collidable = True
    TestKicker.enabled = False
    TestTimer.enabled = True
    TestDrain.Enabled = True
  end If


  'TestWall.Dropped = False
End Sub

Sub TestDrain_Hit
  TestDrain.DestroyBall
End Sub

' ***** END TEST *****



' ***** B2S *****

sub B2SLoadedTimer_Timer

  B2S_UpdateHighScoreReel
  B2S_HighScoreLight
  IsGameOver = True

  B2SLoadedTimer.Enabled = False

end Sub

Sub B2S_HighScoreLight
  if UseB2S = 1 then
    Controller.B2SSetData 105, 1  ' High Score reel
    Controller.B2SSetData 106, 1  ' High Score light
  end If
End Sub

Sub B2S_SetTiltedOn
  '33 b2s id, default tilt
  if UseB2S = 1 then
    Controller.B2SSetTilt 1
  end If
End Sub

Sub B2S_SetTiltedOff
  '33 b2s id, default tilt
  if UseB2S = 1 then
    Controller.B2SSetTilt 0
  end If
End Sub

Sub B2S_LightbulbAnimation
  if UseB2S = 1 then
    Controller.B2SStartAnimation "globes"
  end If
End Sub

Sub B2S_Set21000LightsOn
  '130,131,132
  if UseB2S = 1 then
    'Controller.B2SSetGameOver 1
    '21k lights on dmd
    Controller.B2SSetData 130, 1
    Controller.B2SSetData 131, 1
    Controller.B2SSetData 132, 1

  end If
End Sub

Sub B2S_Setsomething
  if UseB2S = 1 then
    'Controller.B2SSetGameOver 1
    '21k lights on dmd
    Controller.B2SSetData 130, 0
    Controller.B2SSetData 131, 0
    Controller.B2SSetData 132, 0
  end If
End Sub


Sub B2S_SetGameOver
  'Controller.B2SSetData 100, 1
  if UseB2S = 1 then
    'Debug.print "B2S_SetGameOver"
    Controller.B2SSetGameOver 1

    'Controller.B2SSetData 160, 0  ' Ball reel light
    'Controller.B2SSetData 161, 0  ' Winnings reel light
    'Controller.B2SSetData 162, 0  ' BALL light
    'Controller.B2SSetData 163, 0  ' WINNINGS light
  end If
End Sub

Sub B2S_GameStart

  if UseB2S = 1 then
    'Debug.print "B2S_GameStart"
    Controller.B2SSetGameOver 0

    Controller.B2SSetData 100, 1  ' Score reel
    Controller.B2SSetData 101, 1  ' Score light

    Controller.B2SSetData 103, 1  ' Ball reel light
    Controller.B2SSetData 104, 1  ' Ball light

    Controller.B2SSetData 105, 1  ' High Score reel
    Controller.B2SSetData 106, 1  ' High Score light

    Controller.B2SSetScorePlayer1 000
    Controller.B2SSetScore 2,0 ' Ball Reel

    B2S_SetTiltedOff

  end If

End Sub

Sub B2S_UpdateBallReel
  dim currBall

  if UseB2S = 1 then
    Controller.B2SSetScore 2,BallsPlayed ' Ball Reel
  end If

End Sub

Sub B2S_UpdateScoreReel

  if UseB2S = 1 then
    Controller.B2SSetScorePlayer1 Score
  end If

End Sub

Sub B2S_UpdateHighScoreReel

  if UseB2S = 1 then
    Controller.B2SSetScore 3,HighScore
  end If

End Sub

' ***** END B2S *****

' ***** DEBUG / SETTINGS *****

Dim TableSlope : TableSlope = 4

dim ballimageid : ballimageid = 1
Dim BallBrightnessModifier : BallBrightnessModifier = 1.0

' Called when options are tweaked by the player.
' - 0: game has started, good time to load options and adjust accordingly
' - 1: an option has changed
' - 2: options have been reset
' - 3: player closed the tweak UI, good time to update staticly prerendered parts
' Table1.Option arguments are:
' - option name, minimum value, maximum value, step between valid values, default value, unit (0=None, 1=Percent), an optional arry of literal strings
Dim dspTriggered : dspTriggered = False
Sub Table1_OptionEvent(ByVal eventId)
  dim ballOption, ballBrightness

  If eventId = 1 And Not dspTriggered Then
    dspTriggered = True
    DisableStaticPreRendering = True
  End If

  dim shooterscale
    TableSlope = Table1.Option("Slope", 1, 10, 1, 4, 0)
  Table1.SlopeMin = TableSlope
    MaxBallsInShooterLane = Table1.Option("Max Balls in Lane", 1, 5, 1, 2, 0)

  ballOption =  Table1.Option("Ball", 0, 8, 1, 3, 0, Array("Random", "Mashup", "White Marble", "Algae", "Blue Wave", "Fire", "Gold", "Violet", "Rainbow"))

  if ballimageid <> ballOption then
    debug.print "Ball image has changed"

    if ballOption = 0 then ballimageid = GetRandomBallImageId()
    ballimageid = ballOption

    if ballOption = 1 then
      RandomizeAll
    else
      SetBall(ballimageid)
    end if
  end if

  ballBrightness = Table1.Option("Ball Brightness", 0, 2, .05, 1, 1)

  if ballBrightness <> BallBrightnessModifier then
    BallBrightnessModifier = ballBrightness
    debug.print "Ball brightness has changed"
    SetBrightnessForAllBalls
  end if

  ' Room brightness
  LightLevel = NightDay/100
  SetRoomBrightness LightLevel

  If eventId = 3 And dspTriggered Then
    dspTriggered = False
    DisableStaticPreRendering = False
  End If
End Sub

Function GetRandomBallImageId
  'gets random ball image based on number of options
  GetRandomBallImageId = Int(7*Rnd) + 2
  debug.print "Random Ball Image ID:" & GetRandomBallImageId
end Function

Sub SetBall(ballid)
  dim BOT
  dim b

  BOT = GetBalls
  For b = 0 to UBound(BOT)
    BOT(b).FrontDecal = GetNextBallName(ballid)
    BOT(b).UserValue = ballid
    SetBallColor BOT(b), ballid
    Next
End Sub

Sub RandomizeAll
  dim BOT
  dim b
  dim tempId
  'Table1.BallFrontDecal = "marble " & ballid

  BOT = GetBalls
  For b = 0 to UBound(BOT)
    tempId = GetRandomBallImageId
    BOT(b).FrontDecal = GetNextBallName(tempId)
    BOT(b).UserValue = tempId
    SetBallColor BOT(b), tempId
    Next
End Sub

Sub SetBrightnessForAllBalls
  dim BOT
  dim b

  BOT = GetBalls
  For b = 0 to UBound(BOT)
    'debug.print "Modifying: " & BOT(b).UserValue
    SetBallColor BOT(b), cInt(BOT(b).UserValue)
    Next
End Sub

Sub SetBallDecal(ball, ballid)
  ball.FrontDecal = "ball" & ballid
End Sub

sub SetBallColor(ball, imageId)
  dim r,g,b
  r = 125 : g = 125 : b = 125
  Select Case imageId
    Case 2 : r = 135 : g = 135 : b = 135
    Case 3 : r = 125 : g = 175 : b = 125
    Case 4 : r = 125 : g = 125 : b = 175
    Case 5 : r = 235 : g = 135 : b = 135
    Case 6 : r = 155 : g = 155 : b = 115
    Case 7 : r = 200 : g = 85 : b = 150
  End Select

  r = r * BallBrightnessModifier
  g = g * BallBrightnessModifier
  b = b * BallBrightnessModifier

  ' ensure min/max values
  If r < 5 Then r = 5 : If r > 255 Then r = 255
  If g < 5 Then g = 5 : If g > 255 Then g = 255
  If b < 5 Then b = 5 : If b > 255 Then b = 255

  ball.color = rgb(r,g,b)
end sub

function GetStandardBallBrightness

  dim b_base

  'ball brightness needs to be moderated by type (white marble, silver, etc.)
  'and adjusted for light level, each handled differently.

  if BallType = 0 Then 'marble
    b_base = 100 * LightLevel + 5
  elseif BallType = 1 Then 'plain marble
    b_base = 90 * LightLevel + 0
  Else ' must be 2, silver ball
    b_base = 120 * LightLevel + 90
  end If

  b_base = b_base * BallBrightnessModifier

  ' ensure min/max values
  If b_base < 10 Then b_base = 10
  If b_base > 255 Then b_base = 255

  GetStandardBallBrightness = b_base

end Function

' ***** END DEBUG / SETTINGS *****

' ***** ROOM BRIGHTNESS *****

' Update these arrays if you want to change more materials with room light level
Dim RoomBrightnessMtlArray: RoomBrightnessMtlArray = Array("VLM.Bake.Active","VLM.Bake.Solid","VLM.Bake.Solid_T")

Sub SetRoomBrightness(lvl)
  If lvl > 1 Then lvl = 1
  If lvl < 0 Then lvl = 0

  ' Lighting level
  Dim v: v=(lvl * 245 + 10)/255

  Dim i: For i = 0 to UBound(RoomBrightnessMtlArray)
    ModulateMaterialBaseColor RoomBrightnessMtlArray(i), i, v
  Next
End Sub

Dim SavedMtlColorArray

'SaveMtlColors

Sub SaveMtlColors
  ReDim SavedMtlColorArray(UBound(RoomBrightnessMtlArray))
  Dim i: For i = 0 to UBound(RoomBrightnessMtlArray)
    SaveMaterialBaseColor RoomBrightnessMtlArray(i), i
  Next
End Sub

Sub SaveMaterialBaseColor(name, idx)
  Dim wrapLighting, roughness, glossyImageLerp, thickness, edge, edgeAlpha, opacity, base, glossy, clearcoat, isMetal, opacityActive, elasticity, elasticityFalloff, friction, scatterAngle
  GetMaterial name, wrapLighting, roughness, glossyImageLerp, thickness, edge, edgeAlpha, opacity, base, glossy, clearcoat, isMetal, opacityActive, elasticity, elasticityFalloff, friction, scatterAngle
  SavedMtlColorArray(idx) = round(base,0)
  debug.print "Save " & idx & ": " & SavedMtlColorArray(idx)
End Sub

Sub ModulateMaterialBaseColor(name, idx, val)
  Debug.print "ModulateMaterialBaseColor " & name & " " & idx & " " & val
  Dim wrapLighting, roughness, glossyImageLerp, thickness, edge, edgeAlpha, opacity, base, glossy, clearcoat, isMetal, opacityActive, elasticity, elasticityFalloff, friction, scatterAngle
  Dim red, green, blue, saved_base, new_base

  'First get the existing material properties
  GetMaterial name, wrapLighting, roughness, glossyImageLerp, thickness, edge, edgeAlpha, opacity, base, glossy, clearcoat, isMetal, opacityActive, elasticity, elasticityFalloff, friction, scatterAngle

  'Get saved color
  saved_base = SavedMtlColorArray(idx)

  'Next extract the r,g,b values from the base color
  red = saved_base And &HFF
  green = (saved_base \ &H100) And &HFF
  blue = (saved_base \ &H10000) And &HFF
  'msgbox red & " " & green & " " & blue

  'Create new color scaled down by 'val', and update the material
  new_base = RGB(red*val, green*val, blue*val)
  'if idx=2 then new_base = RGB(1, 1, 1)
  UpdateMaterial name, wrapLighting, roughness, glossyImageLerp, thickness, edge, edgeAlpha, opacity, new_base, glossy, clearcoat, isMetal, opacityActive, elasticity, elasticityFalloff, friction, scatterAngle
End Sub

Sub BumperFlipper_Animate

  Dim a : a = BumperFlipper.CurrentAngle 'at rest is 92
  Dim BL
  For each BL in BP_Bumper

    BL.roty = (92-a)*13
    'debug.print BumperFlipper.CurrentAngle & " - " & BL.roty
  Next

end Sub

' VLM  Arrays - Start
' Arrays per baked part
Dim BP_Bumper: BP_Bumper=Array(BM_Bumper)
Dim BP_Cab: BP_Cab=Array(BM_Cab)
Dim BP_Parts: BP_Parts=Array(BM_Parts)
Dim BP_Pins: BP_Pins=Array(BM_Pins)
Dim BP_Playfield: BP_Playfield=Array(BM_Playfield)
Dim BP_bottom: BP_bottom=Array(BM_bottom)
Dim BP_center_back: BP_center_back=Array(BM_center_back)
Dim BP_center_front: BP_center_front=Array(BM_center_front)
Dim BP_drainramp: BP_drainramp=Array(BM_drainramp)
Dim BP_left: BP_left=Array(BM_left)
Dim BP_rightramp: BP_rightramp=Array(BM_rightramp)
Dim BP_top: BP_top=Array(BM_top)
' Arrays per lighting scenario
Dim BL_World: BL_World=Array(BM_Bumper, BM_Cab, BM_Parts, BM_Pins, BM_Playfield, BM_bottom, BM_center_back, BM_center_front, BM_drainramp, BM_left, BM_rightramp, BM_top)
' Global arrays
Dim BG_Bakemap: BG_Bakemap=Array(BM_Bumper, BM_Cab, BM_Parts, BM_Pins, BM_Playfield, BM_bottom, BM_center_back, BM_center_front, BM_drainramp, BM_left, BM_rightramp, BM_top)
Dim BG_Lightmap: BG_Lightmap=Array()
Dim BG_All: BG_All=Array(BM_Bumper, BM_Cab, BM_Parts, BM_Pins, BM_Playfield, BM_bottom, BM_center_back, BM_center_front, BM_drainramp, BM_left, BM_rightramp, BM_top)
' VLM  Arrays - End

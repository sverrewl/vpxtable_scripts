Option Explicit
Randomize

'*****************************************************************************************************
' Junior (Genco 1937)
' Author : 7he S4ge (also known as jejepinball or jejevpuniverse)
'
' Versions :
' 1.0 : first version
'*****************************************************************************************************

'First, try to load the Controller.vbs (DOF), which helps controlling additional hardware like lights, gears, knockers, bells and chimes (to increase realism)
'This table uses DOF via the 'SoundFX' calls that are inserted in some of the PlaySound commands, which will then fire an additional event, instead of just playing a sample/sound effect
On Error Resume Next
ExecuteGlobal GetTextFile("controller.vbs")
If Err Then MsgBox "You need the Controller.vbs file in order to run this table (installed with the VPX package in the scripts folder)"
On Error Goto 0

' Options
Dim displayB2S:displayB2S = True ' True for display B2S backglass
Dim BackglassTimerInterval:BackglassTimerInterval = 500 'ms before launching the backglass, can avoid the backglass to be hidden by the playfield in desktop mode
Dim FreeToPlay:FreeToPlay = False ' If true, you don't need to add credits
Dim MusicOn:MusicOn = True ' Enable music playing (see folder in Visual Pinball folder\Music\JuniorGenco\)
Dim UseFlexDMD:UseFlexDMD = True ' Enable Flex DMD display
Dim FlexDMDHighQuality:FlexDMDHighQuality = True ' True for 256x64, False for 128x32
Dim DMDDelay:DMDDelay = 800 ' Delay before the launching of the DMD (for have the DMD on top of all screens)
Dim NumberOfShakesBeforeTilt:NumberOfShakesBeforeTilt = 3 'After this number, it's a tilt and you lose the ball
Dim TimeReminderTilt:TimeReminderTilt = 3000 ' In ms. If you tilt several times in this period, it's a tilt (check also NumberOfShakesBeforeTilt)
Dim TimeBeforeRemovingTilt:TimeBeforeRemovingTilt = 5000 'In ms. After a tilt, we disable tilt after this time.
Dim EnableBallShadows:EnableBallShadows = True ' If true we add shadows below the balls (can be a problem for slow computers)
Dim EnableRollingSound:EnableRollingSound = True ' If true we add sound when balls rolls on the playfield
Dim EnableBallControl:EnableBallControl = false 'Change to true to enable manual ball control (or press C in-game) via the arrow keys and B (boost movement) keys
Dim SpringAnimationTimerInterval:SpringAnimationTimerInterval = 50 'in ms
Dim SpringBarAnimationTimerInterval:SpringBarAnimationTimerInterval = 50 'in ms
Dim SpringBumperAnimationTimerInterval:SpringBumperAnimationTimerInterval = 30 'in ms
Dim TrapAnimationTimerInterval:TrapAnimationTimerInterval = 500 'in ms
Dim GameOverCheckTimerInterval:GameOverCheckTimerInterval = 2000 'in ms, check when all balls are at the bottom of the playfield for detect GAME OVER
Dim RemoveWoodRails:RemoveWoodRails = False 'Better view of the playfield if true


' Game variables
Const TableName = "JuniorGenco" ' file name to save highscores and other variables
Const cGameName = "JuniorGenco" ' for FlexDMD and B2S
Dim Credits:Credits = 0
Dim BIP:BIP = 0 'Number of balls in the playfield
Dim GameOver:GameOver = True
Dim balls
Set balls = CreateObject("System.Collections.ArrayList")
Dim Score:Score = 0 ' Bumpers hits 100 points, rollover switch 1000 points
Dim Score1000:Score1000 = 0 'Only hundred points
Dim Score100:Score100 = 0 'Only thousand points
Dim TiltOn:TiltOn = False 'If tilt, points are not added to the score
Dim NbShakes:NbShakes = 0 'Number of shakes for check if it's a tilt
Dim StartTimeWhenFirstTilt:StartTimeWhenFirstTilt = Timer

Dim EnableRetractPlunger
EnableRetractPlunger = false 'Change to true to enable retracting the plunger at a linear speed; wait for 1 second at the maximum position; move back towards the resting position; nice for button/key plungers

If Table1.ShowDT = false then
    InfosText.Visible = false
  ScoreText.Visible = false
End If

Const BallSize = 23.5  'Ball radius. Junior have 1" inch ball (47px with this table)

Sub Table1_Init
  If RemoveWoodRails Then
    WoodRailLeft.IsDropped = True
    WoodRailRight.IsDropped = True
    WoodRailTop.IsDropped = True
    WoodRailBottom.IsDropped = True
  End If
  InitSpringBars

  ' Load HighScores
    Loadhs
  Savehs ' If the file don't exists, it create the file

  ' Launch B2S server
  If displayB2S Then
    BackglassTimer.Interval = BackglassTimerInterval
    BackglassTimer.Enabled = True
  End If

  Trap_CreateBalls

  If Not FreeToPlay Then
    InfosText.Text = "WAITING COIN"
  Else
    InfosText.Text = "PRESS START"
  End If

  If MusicOn Then
    StartRandomMusic
  End If

  ' Initalise the DMD display
  If UseFlexDMD Then
    LaunchDMD.Interval = DMDDelay
    LaunchDMD.Enabled = True
  End If
End Sub

Sub Table1_Exit()
  If UseFlexDMD And FlexDMDEnable Then
    If Not FlexDMD is Nothing Then
      FlexDMD.Show = False
      FlexDMD.Run = False
      FlexDMD = NULL
    End If
  End If
  If displayB2S Then
    Controller.Pause
    Controller.Stop
  End If
  ' Save high scores
  Savehs
End Sub

' Launch B2S server
Sub BackglassTimer_Timer
  BackglassTimer.Enabled = False
  LoadEM
End Sub

Sub Table1_KeyDown(ByVal keycode)
  ' High score initials enter
  If EnteringInitials then
        CollectInitials(keycode)
        Exit Sub
    End If

  If keycode = AddCreditKey Then
    If Not GameOver Or Credits >= 1 Then
      ' only one coin for this game
      Exit Sub
    End If
        Credits = Credits + 1
    PlaySound "coinslidein"
    InfosText.Text = "COIN INSERTED"
    FlexDMD_DisplayText "COIN INSERTED", "", 5
    'PlaySound "coinslideout"
  End If
  If keycode = StartGameKey Then
    If (Credits > 0 OR FreeToPlay) And TrapAnimationIndex = 0 Then
      Credits = Credits - 1
      GameReset
    Else
      InfosText.Text = "NO COIN"
    End If
  End If
  If keycode = RightMagnaSave Then
    If Not GameOver Then
      Plunger_ReleaseBall
    End If
  End If
  If keycode = LeftMagnaSave And MusicOn Then
    NextMusic
  End If
  If keycode = PlungerKey Then
        If EnableRetractPlunger Then
            Plunger.PullBackandRetract
        Else
        Plunger.PullBack
        End If
    PlaySound "plungerpull",0,1,AudioPan(Plunger),0.25,0,0,1,AudioFade(Plunger)
  End If

  If keycode = LeftTiltKey Then
    Nudge 90, 2
    CheckTilt
  End If

  If keycode = RightTiltKey Then
    Nudge 270, 2
    CheckTilt
  End If

  If keycode = CenterTiltKey Then
    Nudge 0, 2
    CheckTilt
  End If

  If keycode = MechanicalTilt Then
    CheckTilt
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

    'Manual Ball Control
  If EnableBallControl = 1 Then
    If keycode = 203 Then BCleft = 0  ' Left Arrow
    If keycode = 200 Then BCup = 0    ' Up Arrow
    If keycode = 208 Then BCdown = 0  ' Down Arrow
    If keycode = 205 Then BCright = 0 ' Right Arrow
  End If
End Sub


'*****GI Lights On
dim xx

For each xx in GI:xx.State = 1: Next

'**********Sling Shot Animations
' Rstep and Lstep  are the variables that increment the animation
'****************
Dim RStep, Lstep

Sub RightSlingShot_Slingshot
    PlaySound SoundFX("right_slingshot",DOFContactors), 0,1, 0.05,0.05 '0,1, AudioPan(RightSlingShot), 0.05,0,0,1,AudioFade(RightSlingShot)
    RSling.Visible = 0
    RSling1.Visible = 1
    sling1.rotx = 20
    RStep = 0
    RightSlingShot.TimerEnabled = 1
  gi1.State = 0:Gi2.State = 0
End Sub

Sub RightSlingShot_Timer
    Select Case RStep
        Case 3:RSLing1.Visible = 0:RSLing2.Visible = 1:sling1.rotx = 10
        Case 4:RSLing2.Visible = 0:RSLing.Visible = 1:sling1.rotx = 0:RightSlingShot.TimerEnabled = 0:gi1.State = 1:Gi2.State = 1
    End Select
    RStep = RStep + 1
End Sub

Sub LeftSlingShot_Slingshot
    PlaySound SoundFX("left_slingshot",DOFContactors), 0,1, -0.05,0.05 '0,1, AudioPan(LeftSlingShot), 0.05,0,0,1,AudioFade(LeftSlingShot)
    LSling.Visible = 0
    LSling1.Visible = 1
    sling2.rotx = 20
    LStep = 0
    LeftSlingShot.TimerEnabled = 1
  gi3.State = 0:Gi4.State = 0
End Sub

Sub LeftSlingShot_Timer
    Select Case LStep
        Case 3:LSLing1.Visible = 0:LSLing2.Visible = 1:sling2.rotx = 10
        Case 4:LSLing2.Visible = 0:LSLing.Visible = 1:sling2.rotx = 0:LeftSlingShot.TimerEnabled = 0:gi3.State = 1:Gi4.State = 1
    End Select
    LStep = LStep + 1
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

Const tnob = 5 ' total number of balls
ReDim rolling(tnob)
If EnableRollingSound Then
  InitRolling
End If

Sub RollingTimer_Init
  If Not EnableRollingSound Then
    RollingTimer.Enabled = False
  End If
End Sub

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
        If BOT(b).VelZ < -1 and BOT(b).z < 55 and BOT(b).z > 27 Then 'height adjust for ball drop sounds
            PlaySound "fx_ball_drop" & b, 0, ABS(BOT(b).velz)/17, AudioPan(BOT(b)), 0, Pitch(BOT(b)), 1, 0, AudioFade(BOT(b))
        End If
    Next
End Sub

'**********************
' Ball Collision Sound
'**********************

Sub OnBallBallCollision(ball1, ball2, velocity)
  If GameOver Then
    Exit Sub
  End If
  PlaySound("fx_collide"), 0, Csng(velocity) ^2 / 2000, AudioPan(ball1), 0, Pitch(ball1), 0, 0, AudioFade(ball1)
End Sub

'*****************************************
' ninuzzu's BALL SHADOW
'*****************************************
Dim BallShadow
' List of primitives : the Z position must be 1px on top of the playfield
' You must have as many primitives as balls
BallShadow = Array (BallShadow1,BallShadow2,BallShadow3,BallShadow4,BallShadow5)

Sub BallShadowUpdate_Init
  If Not EnableBallShadows Then
    BallShadowUpdate.Enabled = False
  End If
End Sub

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
        'If BOT(b).X < Table1.Width/2 Then
        '    BallShadow(b).X = ((BOT(b).X) - (Ballsize/6) + ((BOT(b).X - (Table1.Width/2))/7)) + 6
        'Else
        '    BallShadow(b).X = ((BOT(b).X) + (Ballsize/6) + ((BOT(b).X - (Table1.Width/2))/7)) - 6
        'End If
    BallShadow(b).X = BOT(b).X + (BOT(b).X - (Table1.Width/2)) * 1.25 / BallSize
        BallShadow(b).Y = BOT(b).Y + 12
    BallShadow(b).Size_X = 5
    BallShadow(b).Size_Y = 5
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

' The sound is played using the VOL, AUDIOPAN, AUDIOFADE and PITCH functions, so the volume and pitch of the sound
' will change according to the ball speed, and the AUDIOPAN & AUDIOFADE functions will change the stereo position
' according to the position of the ball on the table.


'**************************************
' Explanation of the collision routine
'**************************************

' The collision is built in VP.
' You only need to add a Sub OnBallBallCollision(ball1, ball2, velocity) and when two balls collide they
' will call this routine. What you add in the sub is up to you. As an example is a simple Playsound with volume and paning
' depending of the speed of the collision.


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

Sub Springs_Hit(idx)
  RandomSoundBumper
End Sub

Sub RandomSoundRubber()
  Select Case Int(Rnd*3)+1
    Case 1 : PlaySound "rubber_hit_1", 0, Vol(ActiveBall), AudioPan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
    Case 2 : PlaySound "rubber_hit_2", 0, Vol(ActiveBall), AudioPan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
    Case 3 : PlaySound "rubber_hit_3", 0, Vol(ActiveBall), AudioPan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
  End Select
End Sub

Sub RandomSoundBumper()
  Select Case Int(Rnd*3)+1
    Case 1 : PlaySound "spring", 0, Vol(ActiveBall), AudioPan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
    Case 2 : PlaySound "spring2", 0, Vol(ActiveBall), AudioPan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
    Case 3 : PlaySound "spring3", 0, Vol(ActiveBall), AudioPan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
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

'**********************************
' Game init and reset
'**********************************

' Add 5 balls on top of the trap
Sub Trap_CreateBalls
  balls.Add TrapBallRelease1.CreateSizedBall(BallSize)
  TrapBallRelease1.Kick 180, 1
  balls.Add TrapBallRelease2.CreateSizedBall(BallSize)
  TrapBallRelease2.Kick 180, 1
  balls.Add TrapBallRelease3.CreateSizedBall(BallSize)
  TrapBallRelease3.Kick 180, 1
  balls.Add TrapBallRelease4.CreateSizedBall(BallSize)
  TrapBallRelease4.Kick 180, 1
  balls.Add TrapBallRelease5.CreateSizedBall(BallSize)
  TrapBallRelease5.Kick 180, 1
  BIP = 5
End Sub

' Open the trap and destroy the balls
Sub GameReset
  PlaySound "table reset"
  TrapAnimationTimer.Interval = TrapAnimationTimerInterval
  TrapAnimationIndex = 1
  TrapAnimationTimer.Enabled = True
End Sub

' Trap animation
Dim TrapAnimationIndex:TrapAnimationIndex = 0
Sub TrapAnimationTimer_Timer
  If TrapAnimationIndex = 1 Then
    InfosText.Text = "OPEN TRAP..."
    FlexDMD_DisplayText "OPEN TRAP...", "", 5
    ' Move the trap
    TrapWall1.IsDropped = True
  End If
  TrapAnimationIndex = TrapAnimationIndex + 1
  If TrapAnimationIndex = 3 Then
    TrapAnimationTimer.Enabled = False
    TrapAnimationIndex = 0
  End If
End Sub

' Destroy balls after trap opened
Sub TrapDrain_Hit()
  PlaySound "fx_drain",0,1,AudioPan(TrapDrain),0.25,0,0,1,AudioFade(TrapDrain)
  TrapDrain.DestroyBall
  BIP = BIP -1
  balls.RemoveAt(balls.Count - 1)

  If BIP = 0 Then
    ' Move the trap
    TrapWall1.IsDropped = False
    GameOver = False
    InfosText.Text = "GAME READY"
    Score = 0
    Score1000 = 0
    Score100 = 0
    DisplayScore
    DisableTilt
    FlexDMDDefaultScreen = "Score"
    If FlexDMDHighQuality Then
      FlexDMD_DisplayText "JUNIOR BY GENCO", "GAME READY", 5
    Else
      FlexDMD_DisplayText "JUNIOR", "GAME READY", 5
    End If
  End If
End Sub

' Add a new ball in the plunger lane
Sub Plunger_ReleaseBall
  If BIP >= 5 Then
    Exit Sub
  End If
  If BallAlreadyInPlunger Then
    ' A ball already here
    Exit Sub
  End If

  PlaySound SoundFX("fx_ejectball",DOFContactors), 0,1,AudioPan(BallRelease),0.25,0,0,1,AudioFade(BallRelease)
  'Plunger.CreateBall
  balls.Add BallRelease.CreateSizedBall(BallSize)
  BallRelease.Kick 90, 7
  BIP = BIP + 1

  ' If this is the last ball, we check if all balls are at bottom for detect GAME OVER
  If BIP = 5 Then
    GameOverCheckTimer.Interval = GameOverCheckTimerInterval
    GameOverCheckTimer.Enabled = True
  End If
End Sub

' Check if the 5 balls are at bottom of the table
Sub GameOverCheckTimer_Timer
  If GameOver Or balls.Count < 5 Then
    ' already game over
    Exit Sub
  End If

  'InfosText.Text = balls(4).X & " x " & balls(4).Y

  ' check the position Y of all the 5 balls
  Dim i
  For i = 0 to (balls.Count - 1)
    If balls(i).X > 600 Or balls(i).Y < 1000 Then
      Exit Sub
    End If
  Next

  GameOver = True
  DisableTilt
  InfosText.Text = "GAME OVER"
  FlexDMD_DisplayText FormatScore(Score), "GAME OVER", 5
  CheckHighScore()
  GameOverCheckTimer.Enabled = False
End Sub

' A ball is in the plunger lane
Dim BallAlreadyInPlunger:BallAlreadyInPlunger = False
Sub BallReleaseTrigger_Hit
  BallAlreadyInPlunger = True
End Sub

' No ball in the plunger lane
Sub BallReleaseTrigger_UnHit
  BallAlreadyInPlunger = False
End Sub

'******************************************************
' Springs animations : when the ball touch the springs
'******************************************************

Sub Spring1Rubber_Hit
  LaunchSpringAnimation1
End Sub

' Launch the spring animation
Dim SpringAnimation1Index:SpringAnimation1Index = -1
Sub LaunchSpringAnimation1
  If SpringAnimation1Index >= 0 Then
    ' Animation already launched
    Exit Sub
  End If

  SpringTimer1.Interval = SpringAnimationTimerInterval
  SpringTimer1.Enabled = True
  SpringAnimation1Index = 1
End Sub

Sub SpringTimer1_Timer
  If SpringAnimation1Index = 1 Then
    Spring1.Visible = False
    Spring1Animation2.Visible = True
  End If
  If SpringAnimation1Index = 2 Then
    Spring1Animation2.Visible = False
    Spring1Animation3.Visible = True
  End If
  If SpringAnimation1Index = 3 Then
    Spring1Animation3.Visible = False
    Spring1Animation2.Visible = True
  End If
  If SpringAnimation1Index = 4 Then
    Spring1Animation2.Visible = False
    Spring1.Visible = True
  End If
  SpringAnimation1Index = SpringAnimation1Index + 1
  If SpringAnimation1Index >= 6 Then
    SpringTimer1.Enabled = False
    SpringAnimation1Index = -1
  End If
End Sub

Sub Spring2Rubber_Hit
  LaunchSpringAnimation2
End Sub

' Launch the spring animation
Dim SpringAnimation2Index:SpringAnimation2Index = -1
Sub LaunchSpringAnimation2
  If SpringAnimation2Index >= 0 Then
    ' Animation already launched
    Exit Sub
  End If

  SpringTimer2.Interval = SpringAnimationTimerInterval
  SpringTimer2.Enabled = True
  SpringAnimation2Index = 1
End Sub

Sub SpringTimer2_Timer
  If SpringAnimation2Index = 1 Then
    Spring2.Visible = False
    Spring2Animation2.Visible = True
  End If
  If SpringAnimation2Index = 2 Then
    Spring2Animation2.Visible = False
    Spring2Animation3.Visible = True
  End If
  If SpringAnimation2Index = 3 Then
    Spring2Animation3.Visible = False
    Spring2Animation2.Visible = True
  End If
  If SpringAnimation2Index = 4 Then
    Spring2Animation2.Visible = False
    Spring2.Visible = True
  End If
  SpringAnimation2Index = SpringAnimation2Index + 1
  If SpringAnimation2Index >= 6 Then
    SpringTimer2.Enabled = False
    SpringAnimation2Index = -1
  End If
End Sub

Sub Spring3Rubber_Hit
  LaunchSpringAnimation3
End Sub

' Launch the spring animation
Dim SpringAnimation3Index:SpringAnimation3Index = -1
Sub LaunchSpringAnimation3
  If SpringAnimation3Index >= 0 Then
    ' Animation already launched
    Exit Sub
  End If

  SpringTimer3.Interval = SpringAnimationTimerInterval
  SpringTimer3.Enabled = True
  SpringAnimation3Index = 1
End Sub

Sub SpringTimer3_Timer
  If SpringAnimation3Index = 1 Then
    Spring3.Visible = False
    Spring3Animation2.Visible = True
  End If
  If SpringAnimation3Index = 2 Then
    Spring3Animation2.Visible = False
    Spring3Animation3.Visible = True
  End If
  If SpringAnimation3Index = 3 Then
    Spring3Animation3.Visible = False
    Spring3Animation2.Visible = True
  End If
  If SpringAnimation3Index = 4 Then
    Spring3Animation2.Visible = False
    Spring3.Visible = True
  End If
  SpringAnimation3Index = SpringAnimation3Index + 1
  If SpringAnimation3Index >= 6 Then
    SpringTimer3.Enabled = False
    SpringAnimation3Index = -1
  End If
End Sub

Sub Spring4Rubber_Hit
  LaunchSpringAnimation4
End Sub

' Launch the spring animation
Dim SpringAnimation4Index:SpringAnimation4Index = -1
Sub LaunchSpringAnimation4
  If SpringAnimation4Index >= 0 Then
    ' Animation already launched
    Exit Sub
  End If

  SpringTimer4.Interval = SpringAnimationTimerInterval
  SpringTimer4.Enabled = True
  SpringAnimation4Index = 1
End Sub

Sub SpringTimer4_Timer
  If SpringAnimation4Index = 1 Then
    Spring4.Visible = False
    Spring4Animation2.Visible = True
  End If
  If SpringAnimation4Index = 2 Then
    Spring4Animation2.Visible = False
    Spring4Animation3.Visible = True
  End If
  If SpringAnimation4Index = 3 Then
    Spring4Animation3.Visible = False
    Spring4Animation2.Visible = True
  End If
  If SpringAnimation4Index = 4 Then
    Spring4Animation2.Visible = False
    Spring4.Visible = True
  End If
  SpringAnimation4Index = SpringAnimation4Index + 1
  If SpringAnimation4Index >= 6 Then
    SpringTimer4.Enabled = False
    SpringAnimation4Index = -1
  End If
End Sub

Sub Spring5Rubber_Hit
  LaunchSpringAnimation5
End Sub

' Launch the spring animation
Dim SpringAnimation5Index:SpringAnimation5Index = -1
Sub LaunchSpringAnimation5
  If SpringAnimation5Index >= 0 Then
    ' Animation already launched
    Exit Sub
  End If

  SpringTimer5.Interval = SpringAnimationTimerInterval
  SpringTimer5.Enabled = True
  SpringAnimation5Index = 1
End Sub

Sub SpringTimer5_Timer
  If SpringAnimation5Index = 1 Then
    Spring5.Visible = False
    Spring5Animation2.Visible = True
  End If
  If SpringAnimation5Index = 2 Then
    Spring5Animation2.Visible = False
    Spring5Animation3.Visible = True
  End If
  If SpringAnimation5Index = 3 Then
    Spring5Animation3.Visible = False
    Spring5Animation2.Visible = True
  End If
  If SpringAnimation5Index = 4 Then
    Spring5Animation2.Visible = False
    Spring5.Visible = True
  End If
  SpringAnimation5Index = SpringAnimation5Index + 1
  If SpringAnimation5Index >= 6 Then
    SpringTimer5.Enabled = False
    SpringAnimation5Index = -1
  End If
End Sub

Sub Spring6Rubber_Hit
  LaunchSpringAnimation6
End Sub

' Launch the spring animation
Dim SpringAnimation6Index:SpringAnimation6Index = -1
Sub LaunchSpringAnimation6
  If SpringAnimation6Index >= 0 Then
    ' Animation already launched
    Exit Sub
  End If

  SpringTimer6.Interval = SpringAnimationTimerInterval
  SpringTimer6.Enabled = True
  SpringAnimation6Index = 1
End Sub

Sub SpringTimer6_Timer
  If SpringAnimation6Index = 1 Then
    Spring6.Visible = False
    Spring6Animation2.Visible = True
  End If
  If SpringAnimation6Index = 2 Then
    Spring6Animation2.Visible = False
    Spring6Animation3.Visible = True
  End If
  If SpringAnimation6Index = 3 Then
    Spring6Animation3.Visible = False
    Spring6Animation2.Visible = True
  End If
  If SpringAnimation6Index = 4 Then
    Spring6Animation2.Visible = False
    Spring6.Visible = True
  End If
  SpringAnimation6Index = SpringAnimation6Index + 1
  If SpringAnimation6Index >= 6 Then
    SpringTimer6.Enabled = False
    SpringAnimation6Index = -1
  End If
End Sub

'**************************************************************
' Spring bars animations : when the ball touch the spring bars
'**************************************************************

Sub InitSpringBars
  SpringBar1Animation2.IsDropped = True
  SpringBar1Animation3.IsDropped = True
  SpringBar2Animation2.IsDropped = True
  SpringBar2Animation3.IsDropped = True
  SpringBar3Animation2.IsDropped = True
  SpringBar3Animation3.IsDropped = True
  SpringBar4Animation2.IsDropped = True
  SpringBar4Animation3.IsDropped = True
End Sub

Sub SpringBar1_Hit
  LaunchSpringBarAnimation1
End Sub

' Launch the spring bar animation
Dim SpringBarAnimation1Index:SpringBarAnimation1Index = -1
Sub LaunchSpringBarAnimation1
  If SpringBarAnimation1Index >= 0 Then
    ' Animation already launched
    Exit Sub
  End If

  SpringBarTimer1.Interval = SpringBarAnimationTimerInterval
  SpringBarTimer1.Enabled = True
  SpringBarAnimation1Index = 1
End Sub

Sub SpringBarTimer1_Timer
  If SpringBarAnimation1Index = 1 Then
    SpringBar1.IsDropped = True
    SpringBar1Animation2.IsDropped = False
  End If
  If SpringBarAnimation1Index = 2 Then
    SpringBar1Animation2.IsDropped = True
    SpringBar1Animation3.IsDropped = False
  End If
  If SpringBarAnimation1Index = 3 Then
    SpringBar1Animation3.IsDropped = True
    SpringBar1Animation2.IsDropped = False
  End If
  If SpringBarAnimation1Index = 4 Then
    SpringBar1Animation2.IsDropped = True
    SpringBar1.IsDropped = False
  End If
  SpringBarAnimation1Index = SpringBarAnimation1Index + 1
  If SpringBarAnimation1Index >= 6 Then
    SpringBarTimer1.Enabled = False
    SpringBarAnimation1Index = -1
  End If
End Sub

Sub SpringBar2_Hit
  LaunchSpringBarAnimation2
End Sub

' Launch the spring bar animation
Dim SpringBarAnimation2Index:SpringBarAnimation2Index = -1
Sub LaunchSpringBarAnimation2
  If SpringBarAnimation2Index >= 0 Then
    ' Animation already launched
    Exit Sub
  End If

  SpringBarTimer2.Interval = SpringBarAnimationTimerInterval
  SpringBarTimer2.Enabled = True
  SpringBarAnimation2Index = 1
End Sub

Sub SpringBarTimer2_Timer
  If SpringBarAnimation2Index = 1 Then
    SpringBar2.IsDropped = True
    SpringBar2Animation2.IsDropped = False
  End If
  If SpringBarAnimation2Index = 2 Then
    SpringBar2Animation2.IsDropped = True
    SpringBar2Animation3.IsDropped = False
  End If
  If SpringBarAnimation2Index = 3 Then
    SpringBar2Animation3.IsDropped = True
    SpringBar2Animation2.IsDropped = False
  End If
  If SpringBarAnimation2Index = 4 Then
    SpringBar2Animation2.IsDropped = True
    SpringBar2.IsDropped = False
  End If
  SpringBarAnimation2Index = SpringBarAnimation2Index + 1
  If SpringBarAnimation2Index >= 6 Then
    SpringBarTimer2.Enabled = False
    SpringBarAnimation2Index = -1
  End If
End Sub

Sub SpringBar3_Hit
  LaunchSpringBarAnimation3
End Sub

' Launch the spring bar animation
Dim SpringBarAnimation3Index:SpringBarAnimation3Index = -1
Sub LaunchSpringBarAnimation3
  If SpringBarAnimation3Index >= 0 Then
    ' Animation already launched
    Exit Sub
  End If

  SpringBarTimer3.Interval = SpringBarAnimationTimerInterval
  SpringBarTimer3.Enabled = True
  SpringBarAnimation3Index = 1
End Sub

Sub SpringBarTimer3_Timer
  If SpringBarAnimation3Index = 1 Then
    SpringBar3.IsDropped = True
    SpringBar3Animation2.IsDropped = False
  End If
  If SpringBarAnimation3Index = 2 Then
    SpringBar3Animation2.IsDropped = True
    SpringBar3Animation3.IsDropped = False
  End If
  If SpringBarAnimation3Index = 3 Then
    SpringBar3Animation3.IsDropped = True
    SpringBar3Animation2.IsDropped = False
  End If
  If SpringBarAnimation3Index = 4 Then
    SpringBar3Animation2.IsDropped = True
    SpringBar3.IsDropped = False
  End If
  SpringBarAnimation3Index = SpringBarAnimation3Index + 1
  If SpringBarAnimation3Index >= 6 Then
    SpringBarTimer3.Enabled = False
    SpringBarAnimation3Index = -1
  End If
End Sub

Sub SpringBar4_Hit
  LaunchSpringBarAnimation4
End Sub

' Launch the spring bar animation
Dim SpringBarAnimation4Index:SpringBarAnimation4Index = -1
Sub LaunchSpringBarAnimation4
  If SpringBarAnimation4Index >= 0 Then
    ' Animation already launched
    Exit Sub
  End If

  SpringBarTimer4.Interval = SpringBarAnimationTimerInterval
  SpringBarTimer4.Enabled = True
  SpringBarAnimation4Index = 1
End Sub

Sub SpringBarTimer4_Timer
  If SpringBarAnimation4Index = 1 Then
    SpringBar4.IsDropped = True
    SpringBar4Animation2.IsDropped = False
  End If
  If SpringBarAnimation4Index = 2 Then
    SpringBar4Animation2.IsDropped = True
    SpringBar4Animation3.IsDropped = False
  End If
  If SpringBarAnimation4Index = 3 Then
    SpringBar4Animation3.IsDropped = True
    SpringBar4Animation2.IsDropped = False
  End If
  If SpringBarAnimation4Index = 4 Then
    SpringBar4Animation2.IsDropped = True
    SpringBar4.IsDropped = False
  End If
  SpringBarAnimation4Index = SpringBarAnimation4Index + 1
  If SpringBarAnimation4Index >= 6 Then
    SpringBarTimer4.Enabled = False
    SpringBarAnimation4Index = -1
  End If
End Sub

'********************************************************************
' Spring bumpers animations : when the ball touch the spring bumpers
'********************************************************************

Sub SpringBumperTop1_Hit
  SpringBumperAnimation1Direction = "bottom"
  LaunchSpringBumperAnimation1
  BumperHit
End Sub

Sub SpringBumperBottom1_Hit
  SpringBumperAnimation1Direction = "top"
  LaunchSpringBumperAnimation1
  BumperHit
End Sub

Sub SpringBumperLeft1_Hit
  SpringBumperAnimation1Direction = "right"
  LaunchSpringBumperAnimation1
  BumperHit
End Sub

Sub SpringBumperRight1_Hit
  SpringBumperAnimation1Direction = "left"
  LaunchSpringBumperAnimation1
  BumperHit
End Sub

' Launch the spring bumper animation
Dim SpringBumperAnimation1Index:SpringBumperAnimation1Index = -1
Dim SpringBumperAnimation1Direction:SpringBumperAnimation1Direction = "top" 'Direction of the spring : top, bottom, left, right
Dim SpringBumperSpring1rotY:SpringBumperSpring1rotY = -1
Dim SpringBumperSpring1rotZ:SpringBumperSpring1rotZ = -1
Sub LaunchSpringBumperAnimation1
  If SpringBumperAnimation1Index >= 0 Then
    ' Animation already launched
    Exit Sub
  End If

  SpringBumperTimer1.Interval = SpringBumperAnimationTimerInterval
  SpringBumperAnimation1Index = 1
  ' Keep start position of the spring bumper
  SpringBumperSpring1rotY = SpringBumperSpring1.rotY
  SpringBumperSpring1rotZ = SpringBumperSpring1.rotZ
  SpringBumperTimer1.Enabled = True
End Sub

Sub SpringBumperTimer1_Timer
  'SpringBumperSpring1.rotY = 0
  'SpringBumperSpring1.rotZ = SpringBumperSpring1.rotZ + 5 'right
  'SpringBumperSpring1.rotY = 0
  'SpringBumperSpring1.rotZ = SpringBumperSpring1.rotZ - 5 'left
  'SpringBumperSpring1.rotY = 90
  'SpringBumperSpring1.rotZ = SpringBumperSpring1.rotZ + 5 'bottom
  'SpringBumperSpring1.rotY = 270
  'SpringBumperSpring1.rotZ = SpringBumperSpring1.rotZ + 5 'top
  If SpringBumperAnimation1Index <= 5 Then
    If SpringBumperAnimation1Direction = "top" Then
      SpringBumperSpring1.rotY = 270
      SpringBumperSpring1.rotZ = SpringBumperSpring1.rotZ + 5 'top
    End If
    If SpringBumperAnimation1Direction = "bottom" Then
      SpringBumperSpring1.rotY = 90
      SpringBumperSpring1.rotZ = SpringBumperSpring1.rotZ + 5 'bottom
    End If
    If SpringBumperAnimation1Direction = "left" Then
      SpringBumperSpring1.rotY = 0
      SpringBumperSpring1.rotZ = SpringBumperSpring1.rotZ - 5 'left
    End If
    If SpringBumperAnimation1Direction = "right" Then
      SpringBumperSpring1.rotY = 0
      SpringBumperSpring1.rotZ = SpringBumperSpring1.rotZ + 5 'right
    End If
  End If
  If SpringBumperAnimation1Index > 5 Then
    If SpringBumperAnimation1Direction = "top" Then
      SpringBumperSpring1.rotY = 270
      SpringBumperSpring1.rotZ = SpringBumperSpring1.rotZ - 5 'top
    End If
    If SpringBumperAnimation1Direction = "bottom" Then
      SpringBumperSpring1.rotY = 90
      SpringBumperSpring1.rotZ = SpringBumperSpring1.rotZ - 5 'bottom
    End If
    If SpringBumperAnimation1Direction = "left" Then
      SpringBumperSpring1.rotY = 0
      SpringBumperSpring1.rotZ = SpringBumperSpring1.rotZ + 5 'left
    End If
    If SpringBumperAnimation1Direction = "right" Then
      SpringBumperSpring1.rotY = 0
      SpringBumperSpring1.rotZ = SpringBumperSpring1.rotZ - 5 'right
    End If
  End If
  SpringBumperAnimation1Index = SpringBumperAnimation1Index + 1
  If SpringBumperAnimation1Index > 10 Then
    SpringBumperTimer1.Enabled = False
    SpringBumperAnimation1Index = -1
    ' Reload position infos of the spring bumper, because sometimes errors
    SpringBumperSpring1.rotY = SpringBumperSpring1rotY
    SpringBumperSpring1.rotZ = SpringBumperSpring1rotZ
  End If
End Sub

Sub SpringBumperTop2_Hit
  SpringBumperAnimation2Direction = "bottom"
  LaunchSpringBumperAnimation2
  BumperHit
End Sub

Sub SpringBumperBottom2_Hit
  SpringBumperAnimation2Direction = "top"
  LaunchSpringBumperAnimation2
  BumperHit
End Sub

Sub SpringBumperLeft2_Hit
  SpringBumperAnimation2Direction = "right"
  LaunchSpringBumperAnimation2
  BumperHit
End Sub

Sub SpringBumperRight2_Hit
  SpringBumperAnimation2Direction = "left"
  LaunchSpringBumperAnimation2
  BumperHit
End Sub

' Launch the spring bumper animation
Dim SpringBumperAnimation2Index:SpringBumperAnimation2Index = -1
Dim SpringBumperAnimation2Direction:SpringBumperAnimation2Direction = "top" 'Direction of the spring : top, bottom, left, right
Dim SpringBumperSpring2rotY:SpringBumperSpring2rotY = -1
Dim SpringBumperSpring2rotZ:SpringBumperSpring2rotZ = -1
Sub LaunchSpringBumperAnimation2
  If SpringBumperAnimation2Index >= 0 Then
    ' Animation already launched
    Exit Sub
  End If

  SpringBumperTimer2.Interval = SpringBumperAnimationTimerInterval
  SpringBumperAnimation2Index = 1
  ' Keep start position of the spring bumper
  SpringBumperSpring2rotY = SpringBumperSpring2.rotY
  SpringBumperSpring2rotZ = SpringBumperSpring2.rotZ
  SpringBumperTimer2.Enabled = True
End Sub

Sub SpringBumperTimer2_Timer
  'SpringBumperSpring1.rotY = 0
  'SpringBumperSpring1.rotZ = SpringBumperSpring1.rotZ + 5 'right
  'SpringBumperSpring1.rotY = 0
  'SpringBumperSpring1.rotZ = SpringBumperSpring1.rotZ - 5 'left
  'SpringBumperSpring1.rotY = 90
  'SpringBumperSpring1.rotZ = SpringBumperSpring1.rotZ + 5 'bottom
  'SpringBumperSpring1.rotY = 270
  'SpringBumperSpring1.rotZ = SpringBumperSpring1.rotZ + 5 'top
  If SpringBumperAnimation2Index <= 5 Then
    If SpringBumperAnimation2Direction = "top" Then
      SpringBumperSpring2.rotY = 270
      SpringBumperSpring2.rotZ = SpringBumperSpring2.rotZ + 5 'top
    End If
    If SpringBumperAnimation2Direction = "bottom" Then
      SpringBumperSpring2.rotY = 90
      SpringBumperSpring2.rotZ = SpringBumperSpring2.rotZ + 5 'bottom
    End If
    If SpringBumperAnimation2Direction = "left" Then
      SpringBumperSpring2.rotY = 0
      SpringBumperSpring2.rotZ = SpringBumperSpring2.rotZ - 5 'left
    End If
    If SpringBumperAnimation2Direction = "right" Then
      SpringBumperSpring2.rotY = 0
      SpringBumperSpring2.rotZ = SpringBumperSpring2.rotZ + 5 'right
    End If
  End If
  If SpringBumperAnimation2Index > 5 Then
    If SpringBumperAnimation2Direction = "top" Then
      SpringBumperSpring2.rotY = 270
      SpringBumperSpring2.rotZ = SpringBumperSpring2.rotZ - 5 'top
    End If
    If SpringBumperAnimation2Direction = "bottom" Then
      SpringBumperSpring2.rotY = 90
      SpringBumperSpring2.rotZ = SpringBumperSpring2.rotZ - 5 'bottom
    End If
    If SpringBumperAnimation2Direction = "left" Then
      SpringBumperSpring2.rotY = 0
      SpringBumperSpring2.rotZ = SpringBumperSpring2.rotZ + 5 'left
    End If
    If SpringBumperAnimation2Direction = "right" Then
      SpringBumperSpring2.rotY = 0
      SpringBumperSpring2.rotZ = SpringBumperSpring2.rotZ - 5 'right
    End If
  End If
  SpringBumperAnimation2Index = SpringBumperAnimation2Index + 1
  If SpringBumperAnimation2Index > 10 Then
    SpringBumperTimer2.Enabled = False
    SpringBumperAnimation2Index = -1
    ' Reload position infos of the spring bumper, because sometimes errors
    SpringBumperSpring2.rotY = SpringBumperSpring2rotY
    SpringBumperSpring2.rotZ = SpringBumperSpring2rotZ
  End If
End Sub

Sub SpringBumperTop3_Hit
  SpringBumperAnimation3Direction = "bottom"
  LaunchSpringBumperAnimation3
  BumperHit
End Sub

Sub SpringBumperBottom3_Hit
  SpringBumperAnimation3Direction = "top"
  LaunchSpringBumperAnimation3
  BumperHit
End Sub

Sub SpringBumperLeft3_Hit
  SpringBumperAnimation3Direction = "right"
  LaunchSpringBumperAnimation3
  BumperHit
End Sub

Sub SpringBumperRight3_Hit
  SpringBumperAnimation3Direction = "left"
  LaunchSpringBumperAnimation3
  BumperHit
End Sub

' Launch the spring bumper animation
Dim SpringBumperAnimation3Index:SpringBumperAnimation3Index = -1
Dim SpringBumperAnimation3Direction:SpringBumperAnimation3Direction = "top" 'Direction of the spring : top, bottom, left, right
Dim SpringBumperSpring3rotY:SpringBumperSpring3rotY = -1
Dim SpringBumperSpring3rotZ:SpringBumperSpring3rotZ = -1
Sub LaunchSpringBumperAnimation3
  If SpringBumperAnimation3Index >= 0 Then
    ' Animation already launched
    Exit Sub
  End If

  SpringBumperTimer3.Interval = SpringBumperAnimationTimerInterval
  SpringBumperAnimation3Index = 1
  ' Keep start position of the spring bumper
  SpringBumperSpring3rotY = SpringBumperSpring3.rotY
  SpringBumperSpring3rotZ = SpringBumperSpring3.rotZ
  SpringBumperTimer3.Enabled = True
End Sub

Sub SpringBumperTimer3_Timer
  'SpringBumperSpring1.rotY = 0
  'SpringBumperSpring1.rotZ = SpringBumperSpring1.rotZ + 5 'right
  'SpringBumperSpring1.rotY = 0
  'SpringBumperSpring1.rotZ = SpringBumperSpring1.rotZ - 5 'left
  'SpringBumperSpring1.rotY = 90
  'SpringBumperSpring1.rotZ = SpringBumperSpring1.rotZ + 5 'bottom
  'SpringBumperSpring1.rotY = 270
  'SpringBumperSpring1.rotZ = SpringBumperSpring1.rotZ + 5 'top
  If SpringBumperAnimation3Index <= 5 Then
    If SpringBumperAnimation3Direction = "top" Then
      SpringBumperSpring3.rotY = 270
      SpringBumperSpring3.rotZ = SpringBumperSpring3.rotZ + 5 'top
    End If
    If SpringBumperAnimation3Direction = "bottom" Then
      SpringBumperSpring3.rotY = 90
      SpringBumperSpring3.rotZ = SpringBumperSpring3.rotZ + 5 'bottom
    End If
    If SpringBumperAnimation3Direction = "left" Then
      SpringBumperSpring3.rotY = 0
      SpringBumperSpring3.rotZ = SpringBumperSpring3.rotZ - 5 'left
    End If
    If SpringBumperAnimation3Direction = "right" Then
      SpringBumperSpring3.rotY = 0
      SpringBumperSpring3.rotZ = SpringBumperSpring3.rotZ + 5 'right
    End If
  End If
  If SpringBumperAnimation3Index > 5 Then
    If SpringBumperAnimation3Direction = "top" Then
      SpringBumperSpring3.rotY = 270
      SpringBumperSpring3.rotZ = SpringBumperSpring3.rotZ - 5 'top
    End If
    If SpringBumperAnimation3Direction = "bottom" Then
      SpringBumperSpring3.rotY = 90
      SpringBumperSpring3.rotZ = SpringBumperSpring3.rotZ - 5 'bottom
    End If
    If SpringBumperAnimation3Direction = "left" Then
      SpringBumperSpring3.rotY = 0
      SpringBumperSpring3.rotZ = SpringBumperSpring3.rotZ + 5 'left
    End If
    If SpringBumperAnimation3Direction = "right" Then
      SpringBumperSpring3.rotY = 0
      SpringBumperSpring3.rotZ = SpringBumperSpring3.rotZ - 5 'right
    End If
  End If
  SpringBumperAnimation3Index = SpringBumperAnimation3Index + 1
  If SpringBumperAnimation3Index > 10 Then
    SpringBumperTimer3.Enabled = False
    SpringBumperAnimation3Index = -1
    ' Reload position infos of the spring bumper, because sometimes errors
    SpringBumperSpring3.rotY = SpringBumperSpring3rotY
    SpringBumperSpring3.rotZ = SpringBumperSpring3rotZ
  End If
End Sub

Sub SpringBumperTop4_Hit
  SpringBumperAnimation4Direction = "bottom"
  LaunchSpringBumperAnimation4
  BumperHit
End Sub

Sub SpringBumperBottom4_Hit
  SpringBumperAnimation4Direction = "top"
  LaunchSpringBumperAnimation4
  BumperHit
End Sub

Sub SpringBumperLeft4_Hit
  SpringBumperAnimation4Direction = "right"
  LaunchSpringBumperAnimation4
  BumperHit
End Sub

Sub SpringBumperRight4_Hit
  SpringBumperAnimation4Direction = "left"
  LaunchSpringBumperAnimation4
  BumperHit
End Sub

' Launch the spring bumper animation
Dim SpringBumperAnimation4Index:SpringBumperAnimation4Index = -1
Dim SpringBumperAnimation4Direction:SpringBumperAnimation4Direction = "top" 'Direction of the spring : top, bottom, left, right
Dim SpringBumperSpring4rotY:SpringBumperSpring4rotY = -1
Dim SpringBumperSpring4rotZ:SpringBumperSpring4rotZ = -1
Sub LaunchSpringBumperAnimation4
  If SpringBumperAnimation4Index >= 0 Then
    ' Animation already launched
    Exit Sub
  End If

  SpringBumperTimer4.Interval = SpringBumperAnimationTimerInterval
  SpringBumperAnimation4Index = 1
  ' Keep start position of the spring bumper
  SpringBumperSpring4rotY = SpringBumperSpring4.rotY
  SpringBumperSpring4rotZ = SpringBumperSpring4.rotZ
  SpringBumperTimer4.Enabled = True
End Sub

Sub SpringBumperTimer4_Timer
  'SpringBumperSpring1.rotY = 0
  'SpringBumperSpring1.rotZ = SpringBumperSpring1.rotZ + 5 'right
  'SpringBumperSpring1.rotY = 0
  'SpringBumperSpring1.rotZ = SpringBumperSpring1.rotZ - 5 'left
  'SpringBumperSpring1.rotY = 90
  'SpringBumperSpring1.rotZ = SpringBumperSpring1.rotZ + 5 'bottom
  'SpringBumperSpring1.rotY = 270
  'SpringBumperSpring1.rotZ = SpringBumperSpring1.rotZ + 5 'top
  If SpringBumperAnimation4Index <= 5 Then
    If SpringBumperAnimation4Direction = "top" Then
      SpringBumperSpring4.rotY = 270
      SpringBumperSpring4.rotZ = SpringBumperSpring4.rotZ + 5 'top
    End If
    If SpringBumperAnimation4Direction = "bottom" Then
      SpringBumperSpring4.rotY = 90
      SpringBumperSpring4.rotZ = SpringBumperSpring4.rotZ + 5 'bottom
    End If
    If SpringBumperAnimation4Direction = "left" Then
      SpringBumperSpring4.rotY = 0
      SpringBumperSpring4.rotZ = SpringBumperSpring4.rotZ - 5 'left
    End If
    If SpringBumperAnimation4Direction = "right" Then
      SpringBumperSpring4.rotY = 0
      SpringBumperSpring4.rotZ = SpringBumperSpring4.rotZ + 5 'right
    End If
  End If
  If SpringBumperAnimation4Index > 5 Then
    If SpringBumperAnimation4Direction = "top" Then
      SpringBumperSpring4.rotY = 270
      SpringBumperSpring4.rotZ = SpringBumperSpring4.rotZ - 5 'top
    End If
    If SpringBumperAnimation4Direction = "bottom" Then
      SpringBumperSpring4.rotY = 90
      SpringBumperSpring4.rotZ = SpringBumperSpring4.rotZ - 5 'bottom
    End If
    If SpringBumperAnimation4Direction = "left" Then
      SpringBumperSpring4.rotY = 0
      SpringBumperSpring4.rotZ = SpringBumperSpring4.rotZ + 5 'left
    End If
    If SpringBumperAnimation4Direction = "right" Then
      SpringBumperSpring4.rotY = 0
      SpringBumperSpring4.rotZ = SpringBumperSpring4.rotZ - 5 'right
    End If
  End If
  SpringBumperAnimation4Index = SpringBumperAnimation4Index + 1
  If SpringBumperAnimation4Index > 10 Then
    SpringBumperTimer4.Enabled = False
    SpringBumperAnimation4Index = -1
    ' Reload position infos of the spring bumper, because sometimes errors
    SpringBumperSpring4.rotY = SpringBumperSpring4rotY
    SpringBumperSpring4.rotZ = SpringBumperSpring4rotZ
  End If
End Sub

Sub SpringBumperTop5_Hit
  SpringBumperAnimation5Direction = "bottom"
  LaunchSpringBumperAnimation5
  BumperHit
End Sub

Sub SpringBumperBottom5_Hit
  SpringBumperAnimation5Direction = "top"
  LaunchSpringBumperAnimation5
  BumperHit
End Sub

Sub SpringBumperLeft5_Hit
  SpringBumperAnimation5Direction = "right"
  LaunchSpringBumperAnimation5
  BumperHit
End Sub

Sub SpringBumperRight5_Hit
  SpringBumperAnimation5Direction = "left"
  LaunchSpringBumperAnimation5
  BumperHit
End Sub

' Launch the spring bumper animation
Dim SpringBumperAnimation5Index:SpringBumperAnimation5Index = -1
Dim SpringBumperAnimation5Direction:SpringBumperAnimation5Direction = "top" 'Direction of the spring : top, bottom, left, right
Dim SpringBumperSpring5rotY:SpringBumperSpring5rotY = -1
Dim SpringBumperSpring5rotZ:SpringBumperSpring5rotZ = -1
Sub LaunchSpringBumperAnimation5
  If SpringBumperAnimation5Index >= 0 Then
    ' Animation already launched
    Exit Sub
  End If

  SpringBumperTimer5.Interval = SpringBumperAnimationTimerInterval
  SpringBumperAnimation5Index = 1
  ' Keep start position of the spring bumper
  SpringBumperSpring5rotY = SpringBumperSpring5.rotY
  SpringBumperSpring5rotZ = SpringBumperSpring5.rotZ
  SpringBumperTimer5.Enabled = True
End Sub

Sub SpringBumperTimer5_Timer
  'SpringBumperSpring1.rotY = 0
  'SpringBumperSpring1.rotZ = SpringBumperSpring1.rotZ + 5 'right
  'SpringBumperSpring1.rotY = 0
  'SpringBumperSpring1.rotZ = SpringBumperSpring1.rotZ - 5 'left
  'SpringBumperSpring1.rotY = 90
  'SpringBumperSpring1.rotZ = SpringBumperSpring1.rotZ + 5 'bottom
  'SpringBumperSpring1.rotY = 270
  'SpringBumperSpring1.rotZ = SpringBumperSpring1.rotZ + 5 'top
  If SpringBumperAnimation5Index <= 5 Then
    If SpringBumperAnimation5Direction = "top" Then
      SpringBumperSpring5.rotY = 270
      SpringBumperSpring5.rotZ = SpringBumperSpring5.rotZ + 5 'top
    End If
    If SpringBumperAnimation5Direction = "bottom" Then
      SpringBumperSpring5.rotY = 90
      SpringBumperSpring5.rotZ = SpringBumperSpring5.rotZ + 5 'bottom
    End If
    If SpringBumperAnimation5Direction = "left" Then
      SpringBumperSpring5.rotY = 0
      SpringBumperSpring5.rotZ = SpringBumperSpring5.rotZ - 5 'left
    End If
    If SpringBumperAnimation5Direction = "right" Then
      SpringBumperSpring5.rotY = 0
      SpringBumperSpring5.rotZ = SpringBumperSpring5.rotZ + 5 'right
    End If
  End If
  If SpringBumperAnimation5Index > 5 Then
    If SpringBumperAnimation5Direction = "top" Then
      SpringBumperSpring5.rotY = 270
      SpringBumperSpring5.rotZ = SpringBumperSpring5.rotZ - 5 'top
    End If
    If SpringBumperAnimation5Direction = "bottom" Then
      SpringBumperSpring5.rotY = 90
      SpringBumperSpring5.rotZ = SpringBumperSpring5.rotZ - 5 'bottom
    End If
    If SpringBumperAnimation5Direction = "left" Then
      SpringBumperSpring5.rotY = 0
      SpringBumperSpring5.rotZ = SpringBumperSpring5.rotZ + 5 'left
    End If
    If SpringBumperAnimation5Direction = "right" Then
      SpringBumperSpring5.rotY = 0
      SpringBumperSpring5.rotZ = SpringBumperSpring5.rotZ - 5 'right
    End If
  End If
  SpringBumperAnimation5Index = SpringBumperAnimation5Index + 1
  If SpringBumperAnimation5Index > 10 Then
    SpringBumperTimer5.Enabled = False
    SpringBumperAnimation5Index = -1
    ' Reload position infos of the spring bumper, because sometimes errors
    SpringBumperSpring5.rotY = SpringBumperSpring5rotY
    SpringBumperSpring5.rotZ = SpringBumperSpring5rotZ
  End If
End Sub

Sub SpringBumperTop6_Hit
  SpringBumperAnimation6Direction = "bottom"
  LaunchSpringBumperAnimation6
  BumperHit
End Sub

Sub SpringBumperBottom6_Hit
  SpringBumperAnimation6Direction = "top"
  LaunchSpringBumperAnimation6
  BumperHit
End Sub

Sub SpringBumperLeft6_Hit
  SpringBumperAnimation6Direction = "right"
  LaunchSpringBumperAnimation6
  BumperHit
End Sub

Sub SpringBumperRight6_Hit
  SpringBumperAnimation6Direction = "left"
  LaunchSpringBumperAnimation6
  BumperHit
End Sub

' Launch the spring bumper animation
Dim SpringBumperAnimation6Index:SpringBumperAnimation6Index = -1
Dim SpringBumperAnimation6Direction:SpringBumperAnimation6Direction = "top" 'Direction of the spring : top, bottom, left, right
Dim SpringBumperSpring6rotY:SpringBumperSpring6rotY = -1
Dim SpringBumperSpring6rotZ:SpringBumperSpring6rotZ = -1
Sub LaunchSpringBumperAnimation6
  If SpringBumperAnimation6Index >= 0 Then
    ' Animation already launched
    Exit Sub
  End If

  SpringBumperTimer6.Interval = SpringBumperAnimationTimerInterval
  SpringBumperAnimation6Index = 1
  ' Keep start position of the spring bumper
  SpringBumperSpring6rotY = SpringBumperSpring6.rotY
  SpringBumperSpring6rotZ = SpringBumperSpring6.rotZ
  SpringBumperTimer6.Enabled = True
End Sub

Sub SpringBumperTimer6_Timer
  'SpringBumperSpring1.rotY = 0
  'SpringBumperSpring1.rotZ = SpringBumperSpring1.rotZ + 5 'right
  'SpringBumperSpring1.rotY = 0
  'SpringBumperSpring1.rotZ = SpringBumperSpring1.rotZ - 5 'left
  'SpringBumperSpring1.rotY = 90
  'SpringBumperSpring1.rotZ = SpringBumperSpring1.rotZ + 5 'bottom
  'SpringBumperSpring1.rotY = 270
  'SpringBumperSpring1.rotZ = SpringBumperSpring1.rotZ + 5 'top
  If SpringBumperAnimation6Index <= 5 Then
    If SpringBumperAnimation6Direction = "top" Then
      SpringBumperSpring6.rotY = 270
      SpringBumperSpring6.rotZ = SpringBumperSpring6.rotZ + 5 'top
    End If
    If SpringBumperAnimation6Direction = "bottom" Then
      SpringBumperSpring6.rotY = 90
      SpringBumperSpring6.rotZ = SpringBumperSpring6.rotZ + 5 'bottom
    End If
    If SpringBumperAnimation6Direction = "left" Then
      SpringBumperSpring6.rotY = 0
      SpringBumperSpring6.rotZ = SpringBumperSpring6.rotZ - 5 'left
    End If
    If SpringBumperAnimation6Direction = "right" Then
      SpringBumperSpring6.rotY = 0
      SpringBumperSpring6.rotZ = SpringBumperSpring6.rotZ + 5 'right
    End If
  End If
  If SpringBumperAnimation6Index > 5 Then
    If SpringBumperAnimation6Direction = "top" Then
      SpringBumperSpring6.rotY = 270
      SpringBumperSpring6.rotZ = SpringBumperSpring6.rotZ - 5 'top
    End If
    If SpringBumperAnimation6Direction = "bottom" Then
      SpringBumperSpring6.rotY = 90
      SpringBumperSpring6.rotZ = SpringBumperSpring6.rotZ - 5 'bottom
    End If
    If SpringBumperAnimation6Direction = "left" Then
      SpringBumperSpring6.rotY = 0
      SpringBumperSpring6.rotZ = SpringBumperSpring6.rotZ + 5 'left
    End If
    If SpringBumperAnimation6Direction = "right" Then
      SpringBumperSpring6.rotY = 0
      SpringBumperSpring6.rotZ = SpringBumperSpring6.rotZ - 5 'right
    End If
  End If
  SpringBumperAnimation6Index = SpringBumperAnimation6Index + 1
  If SpringBumperAnimation6Index > 10 Then
    SpringBumperTimer6.Enabled = False
    SpringBumperAnimation6Index = -1
    ' Reload position infos of the spring bumper, because sometimes errors
    SpringBumperSpring6.rotY = SpringBumperSpring6rotY
    SpringBumperSpring6.rotZ = SpringBumperSpring6rotZ
  End If
End Sub

Sub SpringBumperTop7_Hit
  SpringBumperAnimation7Direction = "bottom"
  LaunchSpringBumperAnimation7
  BumperHit
End Sub

Sub SpringBumperBottom7_Hit
  SpringBumperAnimation7Direction = "top"
  LaunchSpringBumperAnimation7
  BumperHit
End Sub

Sub SpringBumperLeft7_Hit
  SpringBumperAnimation7Direction = "right"
  LaunchSpringBumperAnimation7
  BumperHit
End Sub

Sub SpringBumperRight7_Hit
  SpringBumperAnimation7Direction = "left"
  LaunchSpringBumperAnimation7
  BumperHit
End Sub

' Launch the spring bumper animation
Dim SpringBumperAnimation7Index:SpringBumperAnimation7Index = -1
Dim SpringBumperAnimation7Direction:SpringBumperAnimation7Direction = "top" 'Direction of the spring : top, bottom, left, right
Dim SpringBumperSpring7rotY:SpringBumperSpring7rotY = -1
Dim SpringBumperSpring7rotZ:SpringBumperSpring7rotZ = -1
Sub LaunchSpringBumperAnimation7
  If SpringBumperAnimation7Index >= 0 Then
    ' Animation already launched
    Exit Sub
  End If

  SpringBumperTimer7.Interval = SpringBumperAnimationTimerInterval
  SpringBumperAnimation7Index = 1
  ' Keep start position of the spring bumper
  SpringBumperSpring7rotY = SpringBumperSpring7.rotY
  SpringBumperSpring7rotZ = SpringBumperSpring7.rotZ
  SpringBumperTimer7.Enabled = True
End Sub

Sub SpringBumperTimer7_Timer
  'SpringBumperSpring1.rotY = 0
  'SpringBumperSpring1.rotZ = SpringBumperSpring1.rotZ + 5 'right
  'SpringBumperSpring1.rotY = 0
  'SpringBumperSpring1.rotZ = SpringBumperSpring1.rotZ - 5 'left
  'SpringBumperSpring1.rotY = 90
  'SpringBumperSpring1.rotZ = SpringBumperSpring1.rotZ + 5 'bottom
  'SpringBumperSpring1.rotY = 270
  'SpringBumperSpring1.rotZ = SpringBumperSpring1.rotZ + 5 'top
  If SpringBumperAnimation7Index <= 5 Then
    If SpringBumperAnimation7Direction = "top" Then
      SpringBumperSpring7.rotY = 270
      SpringBumperSpring7.rotZ = SpringBumperSpring7.rotZ + 5 'top
    End If
    If SpringBumperAnimation7Direction = "bottom" Then
      SpringBumperSpring7.rotY = 90
      SpringBumperSpring7.rotZ = SpringBumperSpring7.rotZ + 5 'bottom
    End If
    If SpringBumperAnimation7Direction = "left" Then
      SpringBumperSpring7.rotY = 0
      SpringBumperSpring7.rotZ = SpringBumperSpring7.rotZ - 5 'left
    End If
    If SpringBumperAnimation7Direction = "right" Then
      SpringBumperSpring7.rotY = 0
      SpringBumperSpring7.rotZ = SpringBumperSpring7.rotZ + 5 'right
    End If
  End If
  If SpringBumperAnimation7Index > 5 Then
    If SpringBumperAnimation7Direction = "top" Then
      SpringBumperSpring7.rotY = 270
      SpringBumperSpring7.rotZ = SpringBumperSpring7.rotZ - 5 'top
    End If
    If SpringBumperAnimation7Direction = "bottom" Then
      SpringBumperSpring7.rotY = 90
      SpringBumperSpring7.rotZ = SpringBumperSpring7.rotZ - 5 'bottom
    End If
    If SpringBumperAnimation7Direction = "left" Then
      SpringBumperSpring7.rotY = 0
      SpringBumperSpring7.rotZ = SpringBumperSpring7.rotZ + 5 'left
    End If
    If SpringBumperAnimation7Direction = "right" Then
      SpringBumperSpring7.rotY = 0
      SpringBumperSpring7.rotZ = SpringBumperSpring7.rotZ - 5 'right
    End If
  End If
  SpringBumperAnimation7Index = SpringBumperAnimation7Index + 1
  If SpringBumperAnimation7Index > 10 Then
    SpringBumperTimer7.Enabled = False
    SpringBumperAnimation7Index = -1
    ' Reload position infos of the spring bumper, because sometimes errors
    SpringBumperSpring7.rotY = SpringBumperSpring7rotY
    SpringBumperSpring7.rotZ = SpringBumperSpring7rotZ
  End If
End Sub

'********
' Score
'********

Sub DisplayScore
  ScoreText.Text = CStr(Score)
  If displayB2S Then
    ' Reset illuminations
    Controller.B2SSetData 51,0
    Controller.B2SSetData 50,0

    If Score1000 > 0 Then
      ' This game can't display up to 8000 points
      If Score1000 <= 8000 Then
        Controller.B2SSetData 50,(Score1000 / 1000)
      Else
        Controller.B2SSetData 50,8 'stay to 8000
      End If
    End If
    If Score100 > 0 And Score100 <= 900 Then
      Controller.B2SSetData 51,(Score100 / 100)
    Else
      Controller.B2SSetData 51,10 'Illumination for 0 has B2SID=1
    End If
  End If
  FlexDMD_Score 0
End Sub

Sub Trigger1_Hit
  If TiltOn Then
    ' No points during tilt
    Exit Sub
  End If
  Score = Score + 1000
  PlaySound "BellA"
  BellSound
  Score1000 = Score1000 + 1000
  DisplayScore
End Sub

Sub BumperHit
  RandomSoundBumper()
  If TiltOn Then
    ' No points during tilt
    Exit Sub
  End If
  Score = Score + 100
  BellSound
  Score100 = Score100 + 100
  If Score100 > 900 Then
    Score100 = 0
    Score1000 = Score1000 + 1000
  End If
  DisplayScore
End Sub

' Bell sound when score corresponds to instruction card goals
Sub BellSound
  If Score = 6000 Or Score = 6500 Or Score = 7000 Or Score = 7500 Or Score = 8000 Then
    PlaySound "BellA5"
  End If
End Sub

'********************
' Music
'********************

' Random music : add a folder with .mp3 files in Visual Pinball folder\Music\XXX\
Dim FileSystemObject, folderRM, rRM, ctRM, fileRM, musicPath, musicFileName, musicFolder
musicFolder = "JuniorGenco"
musicPath = ""
Sub StartRandomMusic
  Set FileSystemObject = CreateObject("Scripting.FileSystemObject")
  If musicPath = "" Then
    If (FileSystemObject.FolderExists(UserDirectory)) Then 'User directory. Ex : C:\VisualPinball\User
      musicPath = Left(UserDirectory, Len(UserDirectory) - 5)
      musicPath = musicPath & "\Music\" & musicFolder & "\"
    End If
    If Not (FileSystemObject.FolderExists(musicPath)) Then
      musicPath = "-1"
    End If
  End If
  If musicPath <> "-1" Then
    Set folderRM = FileSystemObject.GetFolder(musicPath)
    Randomize ' randow motor
    rRM = INT(folderRM.Files.Count * Rnd + 1) 'Define a random number
    ctRM=1
    For Each fileRM in folderRM.Files
      If ctRM = rRM Then
        If (LCase(Right(fileRM,4))) = ".mp3" Then 'Search a .mp3 file
          PlayMusic musicFolder & "\" & Right(fileRM, Len(fileRM) - Len(musicPath) + 1), 1 'Ex : XXX\fileName.mp3
          ' Here you can get the name of the file for display it wherever you want
          Dim musicFilename 'music name without .mp3 extension
          musicFilename = Right(fileRM, Len(fileRM) - Len(musicPath) + 1)
          musicFilename = Left(musicFilename, Len(musicFilename) - 4)
          'MsgBox musicFilename
          Exit For
        End If
      End If
    ctRM = ctRM + 1
    Next
  End If
End Sub

' When music stops, we relaunch a new music
Sub Table1_MusicDone
  StartRandomMusic
End Sub

Sub EndRandomMusic
  EndMusic()
End Sub

Sub NextMusic
  EndRandomMusic
  StartRandomMusic
End Sub

'******
' Tilt
'******

' Add a shake and check if it's a tilt or not
Sub CheckTilt
  If GameOver Then
    Exit Sub
  End If
  If TiltOn Then
    ' It's already tilted
    Exit Sub
  End If
  ' NumberOfShakesBeforeTilt
  ' StartTimeWhenFirstTilt
  ' TimeReminderTilt
  If NbShakes > 0 And ((Timer - StartTimeWhenFirstTilt) * 1000) > TimeReminderTilt Then
    NbShakes = 0
  End If
  If NbShakes > 0 Then
    NbShakes = NbShakes + 1
  Else
    StartTimeWhenFirstTilt = Timer
    NbShakes = NbShakes + 1
  End If
  If NbShakes >= NumberOfShakesBeforeTilt Then
    TiltOn = True
    TimeBeforeRemovingTiltTimer.Interval = TimeBeforeRemovingTilt
    TimeBeforeRemovingTiltTimer.Enabled = True
    End If
  If TiltOn Then
    InfosText.Text = "TILT"
    FlexDMD_DisplayText "TILT", "", 5
    If displayB2S Then
      Controller.B2SSetTilt 1
    End If
  End If
End Sub

Sub DisableTilt
  TiltOn = False
  If InfosText.Text = "TILT" Then
    InfosText.Text = ""
    FlexDMD_DisplayText "", "", 5
  End If
  If displayB2S Then
    Controller.B2SSetTilt 0
  End If
End Sub

' Disable tilt after this interval
Sub TimeBeforeRemovingTiltTimer_Timer
  TimeBeforeRemovingTiltTimer.Enabled = False
  DisableTilt
End Sub

'***************
' DMD
'***************

Dim FlexDMD, FlexDMDEnable, FlexDMDFontDefault, FlexDMDFont26, FlexDMDFont24, FlexDMDFont12, FlexDMDFont7, FlexDMDFont5, FlexDMDDefaultScreen, FlexDMDDefaultBackgroundImage
FlexDMDEnable = False
FlexDMDDefaultScreen = "Logo" ' Indicates the screen displayed after a quick display : Score, Logo
FlexDMDDefaultBackgroundImage = "dmd_border3" ' Indicates the screen displayed behind the text and score : dmd_border, dmd_border2, dmd_border3

Const vbAliceBlue = 1677540
Const vbAntiqueWhite = 14150650
Const vbAqua = 16776960
Const vbAquamarine = 13959039
Const vbAzure = 16777200
Const vbBeige = 14480885
Const vbBisque = 12903679
Const vbBlancheDalmond = 13495295
Const vbBlueViolet = 14822282
Const vbBrown = 2763429
Const vbBurlyWood = 8894686
Const vbCadetBlue = 10526303
Const vbChartReuse = 65407
Const vbChocolate = 1993170
Const vbCoral = 5275647
Const vbCornFlowerBlue = 15570276
Const vbCornSilk = 14481663
Const vbCream = 15793151
Const vbCrimson = 3937500
Const vbDarkBlue = 9109504
Const vbDarkCyan = 9145088
Const vbDarkerGray = 8421504
Const vbDarkGreen = 25600
Const vbDarkGoldenRod = 755384
Const vbDarkGray = 11119017
'Const vbGreen = 32512
Const vbDarkKhaki = 7059389
Const vbDarkMagenta = 9109643
Const vbDarkOliveGreen = 3107669
Const vbDarkOrange = 36095
Const vbDarkOrchid = 13382297
Const vbDarkRed = 139
Const vbDarkSalmon = 8034025
Const vbDarkSeaGreen = 9419919
Const vbDarkSlateBlue = 9125192
Const vbDarkSlateGray = 5197615
Const vbDarkSlateGrey = 5197615
Const vbDarkTurquoise = 13749760
Const vbDarkViolet = 13828244
Const vbDeepPink = 9639167
Const vbDeepSkyBlue = 16760576
Const vbDimGray = 4144959
Const vbDodgerBlue = 16748574
Const vbFireBrick = 2237106
Const vbFloralWhite = 15792895
Const vbFuchsia = 16711935
Const vbGainsBoro = 14474460
Const vbGhostWhite = 16775416
Const vbGold = 55295
Const vbGoldenRod = 2139610
Const vbGray = 8355711
Const vbGreenYellow = 3145645
Const vbHoneyDew = 15794160
Const vbHotPink = 11823615
Const vbIndianRed = 6053069
Const vbIndigo = 8519755
Const vbIvory = 15794175
Const vbKhaki = 9234160
Const vbLavender = 16443110
Const vbLavenderBlush = 16118015
Const vbLawnGreen = 64636
Const vbLegacySkyBlue = 15780518
Const vbLemonChiffon = 13499135
Const vbLightBlue = 15128749
Const vbLightCoral = 8421616
Const vbLightCyan = 16777184
Const vbLightGrey = 13882323
Const vbLightGoldenRodYellow = 13826810
Const vbLightGray = 12632256
Const vbLightGreen = 9498256
Const vbLightPink = 12695295
Const vbLightSalmon = 8036607
Const vbLightSeaGreen = 11186720
Const vbLightSkyBlue = 16436871
Const vbLightSlateGray = 10061943
Const vbLightSteelBlue = 14599344
Const vbLightYellow = 14745599
Const vbLime = 65280
Const vbLimeGreen = 3329330
Const vbLinen = 15134970
Const vbMaroon = 127
Const vbMediumAquamarine = 11193702
Const vbMediumBlue = 13434880
Const vbMediumGray = 10789024
Const vbMediumOrchid = 13850042
Const vbMediumPurple = 14381203
Const vbMediumSeaGreen = 7451452
Const vbMediumSlateBlue = 15624315
Const vbMediumSpringGreen = 10156544
Const vbMediumTurquoise = 13422920
Const vbMediumVioletRed = 8721863
Const vbMidnightBlue = 7346457
Const vbMintCream = 16449525
Const vbMistyRose = 14804223
Const vbMoccasin = 11920639
Const vbMoneyGreen = 12639424
Const vbNavajoWhite = 11394815
Const vbNavy = 8323072
Const vbOldLace = 15136253
Const vbOlive = 32639
Const vbOliveDrab = 2330219
Const vbOrange = 42495
Const vbOrangeRed = 17919 'don't work
Const vbOrchid = 14053594
Const vbPaleGoldenRod = 11200750
Const vbPaleGreen = 10025880
Const vbPaleTurquoise = 15658671
Const vbPaleVioletRed = 9662683
Const vbPapayaWhip = 14020607
Const vbPeachPuff = 12180223
Const vbPeru = 4163021
Const vbPlum = 14524637
Const vbPowderBlue = 15130800
Const vbPurple = 8323199
Const vbRosyBrown = 9408444
Const vbRoyalBlue = 14772545
Const vbSaddleBrown = 1262987
Const vbSalmon = 7504122
Const vbSandyBrown = 6333684
Const vbSeaGreen = 5737262
Const vbSeaShell = 15660543
Const vbSienna = 2970272
Const vbSilver = 12632256
Const vbSkyBlue = 15453831
Const vbSlateBlue = 13458026
Const vbSlateGray = 9470064
Const vbSlateGrey = 9470064
Const vbSnow = 16448255
Const vbSpringGreen = 8388352
Const vbSteelBlue = 11829830
Const vbTan = 9221330
Const vbTeal = 8355584
Const vbThistle = 14204888
Const vbTomato = 4678655
Const vbTurquoise = 13688896
Const vbViolet = 15631086
Const vbWheat = 11788021
Const vbWhiteSmoke = 16119285
Const vbYellowGreen = 3329434

' Delay before the launching of the DMD
Sub LaunchDMD_Timer
  LaunchDMD.Enabled = False
  FlexDMD_Load
End Sub

' Load Flex DMD display
Dim DMDScene
Dim FlexDMDWidth:FlexDMDWidth = 256
Dim FlexDMDHeight:FlexDMDHeight = 64
Sub FlexDMD_Load
  Set FlexDMD = CreateObject("FlexDMD.FlexDMD")
  If Not FlexDMD is Nothing Then
    FlexDMD.TableFile = Table1.Filename & ".vpx"
    FlexDMD.RenderMode = 2 ' monochrome with 4 shades (value = 0) / monochrome with 16 shades (value = 1) / full color (value = 2)
    If FlexDMDHighQuality Then
      FlexDMDWidth = 256
      FlexDMDHeight = 64
      FlexDMD.Width = 256
      FlexDMD.Height = 64
    Else
      FlexDMDWidth = 128
      FlexDMDHeight = 32
      FlexDMD.Width = 128
      FlexDMD.Height = 32
    End If
    FlexDMD.Clear = True
    FlexDMD.GameName = cGameName
    FlexDMD.Run = True

    ' Fonts
    If FlexDMDHighQuality Then
      FlexDMDFontDefault = FlexDMD.NewFont("FlexDMD.Resources.udmd-f12by24.fnt", vbBrown, vbGold, 1)
    Else
      FlexDMDFontDefault = FlexDMD.NewFont("FlexDMD.Resources.bm_army-12.fnt", vbBrown, vbGold, 1)
    End If
    FlexDMDFont26 = FlexDMD.NewFont("FlexDMD.Resources.udmd-f14by26.fnt", vbRed, vbGreen, 1)
    FlexDMDFont24 = FlexDMD.NewFont("FlexDMD.Resources.udmd-f12by24.fnt", vbRed, vbGreen, 1)
    FlexDMDFont12 = FlexDMD.NewFont("FlexDMD.Resources.bm_army-12.fnt", vbRed, vbGreen, 1)
    FlexDMDFont7 = FlexDMD.NewFont("FlexDMD.Resources.zx_spectrum-7.fnt", vbRed, vbGreen, 1)
    FlexDMDFont5 = FlexDMD.NewFont("FlexDMD.Resources.teeny_tiny_pixls-5.fnt", vbRed, vbGreen, 1)

    ' Default background image
    Set DMDScene = FlexDMD.NewGroup("Scene")
    FlexDMD.LockRenderThread
    FlexDMD.Stage.AddActor DMDScene
    FlexDMD.UnlockRenderThread
  End If
  FlexDMDEnable = True

  FlexDMD_DisplayDefaultScreen
End Sub

' Display score of the current player
' seconds : nb of seconds to force the display, after that we show FlexDMDDefaultScreen. 0 to infinite display
Sub FlexDMD_Score(seconds)
  If Not FlexDMDEnable Then
    Exit Sub
  End If

  Dim textLine2
  If Not GameOver Then
    'textLine2 = "Ball " & CStr(CurrentBallNumber)
  Else
    textLine2 = "GAME OVER"
  End If
  FlexDMD_DisplayText FormatScore(Score), textLine2, seconds
End Sub

' Called when the screen can be replaced by the FlexDMDDefaultScreen
Sub FlexDMDScreenTimer_Timer
  FlexDMDScreenTimer.Enabled = False

  FlexDMD_DisplayDefaultScreen
End Sub

Sub FlexDMD_DisplayDefaultScreen
  If FlexDMDDefaultScreen = "Score" Then
    FlexDMD_Score 0
  End If
  If FlexDMDDefaultScreen = "Logo" Then
    FlexDMD_DisplayImage "dmd_logo", 0, True
  End If
End Sub

' Display text (centered in width, on two lines)
' seconds : nb of seconds to force the display, after that we show FlexDMDDefaultScreen. 0 to infinite display
Sub FlexDMD_DisplayText(textLine1, textLine2, seconds)
  If Not FlexDMDEnable Then
    Exit Sub
  End If

  FlexDMD.LockRenderThread
  DMDScene.RemoveAll 'TODO : use .Update for avoid the removall of actors ?
  ' Default background image
  FlexDMD_DisplayImage FlexDMDDefaultBackgroundImage, 0, False
  'Dim AlignWidth:AlignWidth = CInt(CLng(FlexDMD.Width) / CLng(42.7))
  'Dim AlignHeight:AlignHeight = CInt(CLng(FlexDMD.Height) / CLng(10.7))
  'Dim AlignHeightLow:AlignHeightLow = CInt(CLng(FlexDMD.Height) / CLng(1.8))
  If textLine1 <> "" Then
    DMDScene.AddActor(FlexDMD.NewLabel("TextLine1", FlexDMDFontDefault, textLine1))
    'Alignment : 0 (TopLeft), 1 (Top), 2 (TopRight), 3 (Left), 4 (Center), 5 (Right), 6 (BottomLeft), 7 (Bottom), 8 (BottomRight)
    'DMDScene.GetLabel("TextLine1").SetAlignedPosition AlignWidth, AlignHeight, 0 ' x, y, alignment (0 for TopLeft)
    DMDScene.GetLabel("TextLine1").SetAlignedPosition CInt(CLng(FlexDMDWidth / 2)), CInt(CLng(FlexDMDHeight / 4)), 4
  End If
  If textLine2 <> "" Then
    DMDScene.AddActor(FlexDMD.NewLabel("TextLine2", FlexDMDFontDefault, textLine2))
    'DMDScene.GetLabel("TextLine2").SetAlignedPosition AlignWidth, AlignHeightLow, 0 ' x, y, alignment (0 for TopLeft)
    DMDScene.GetLabel("TextLine2").SetAlignedPosition CInt(CLng(FlexDMDWidth / 2)), CInt(CLng(FlexDMDHeight * 3 / 4)), 4
  End If
  FlexDMD.UnlockRenderThread

  If seconds > 0 Then
    ' Start the timer
    FlexDMDScreenTimer.Interval = 1000 * seconds
    FlexDMDScreenTimer.Enabled = True
  End If
End Sub

' Display image from Table\Image Manager
' seconds : nb of seconds to force the display, after that we show FlexDMDDefaultScreen. 0 to infinite display
Sub FlexDMD_DisplayImage(imageName, seconds, removeAll)
  If Not FlexDMDEnable Then
    Exit Sub
  End If

  FlexDMD.LockRenderThread
  If removeAll Then
    DMDScene.RemoveAll 'TODO : use .Update for avoid the removall of actors ?
  End If
  DMDScene.AddActor FlexDMD.NewImage(imageName, "VPX." & imageName)
  DMDScene.GetImage(imageName).SetSize FlexDMD.Width, FlexDMD.Height
  FlexDMD.UnlockRenderThread

  If seconds > 0 Then
    ' Start the timer
    FlexDMDScreenTimer.Interval = 1000 * seconds
    FlexDMDScreenTimer.Enabled = True
  End If
End Sub

Function ExpandLine(TempStr) 'id is the number of the dmd line
    If TempStr = "" Then
        TempStr = Space(20)
    Else
        if Len(TempStr) > Space(20) Then
            TempStr = Left(TempStr, Space(20) )
        Else
            if(Len(TempStr) < 20) Then
                TempStr = TempStr & Space(20 - Len(TempStr) )
            End If
        End If
    End If
    ExpandLine = TempStr
End Function

'Convert numbers with commas separators (as in Black's original font)
Function FormatScore(ByVal Num)
    dim i
    dim NumString

    NumString = CStr(abs(Num)) ' remove commas and negative sign

    For i = Len(NumString) -3 to 1 step -3
    if IsNumeric(mid(NumString, i, 1)) then ' mid(string, start, length)
            'NumString = left(NumString, i-1) & chr(asc(mid(NumString, i, 1)) + 48) & right(NumString, Len(NumString) - i) 'seems to work only with a specific font
      NumString = left(NumString, i-1) & mid(NumString, i, 1) & "," & right(NumString, Len(NumString) - i)
        end if
    Next
    FormatScore = NumString
End function

Function FL(NumString1, NumString2) 'Fill line
    Dim Temp, TempStr
    Temp = 17 - Len(NumString1) - Len(NumString2)
    TempStr = NumString1 & Space(Temp) & NumString2
    FL = TempStr
End Function

' Returns the text centered with spaces on left and right
Function CL(NumString)
    Dim Temp, TempStr
    Temp = (17 - Len(NumString) ) \ 2
    TempStr = Space(Temp) & NumString & Space(Temp)
    CL = TempStr
End Function

Function RL(NumString) 'right line
    Dim Temp, TempStr
    Temp = 17 - Len(NumString)
    TempStr = Space(Temp) & NumString
    RL = TempStr
End Function

'*********************************
'    Load / Save / Highscore
'*********************************

Dim HighScore

' Load high scores in User directory of Visual Pinball
Sub Loadhs
    ' Based on Black's Highscore routines
    Dim FileObj
    Dim ScoreFile, TextStr
    Dim temp1
    Dim temp2
    Dim temp3
    Dim temp4
    Dim temp5
    Dim temp6
    Dim temp8
    Dim temp9
    Dim temp10
    Dim temp11
    Dim temp12
    Dim temp13
    Dim temp14
    Dim temp15
    Dim temp16
    Dim temp17

    Set FileObj = CreateObject("Scripting.FileSystemObject")
    If Not FileObj.FolderExists(UserDirectory) then
        'Credits = 0
        Exit Sub
    End If
    If Not FileObj.FileExists(UserDirectory & TableName& ".txt") then
        'Credits = 0
        Exit Sub
    End If
    Set ScoreFile = FileObj.GetFile(UserDirectory & TableName& ".txt")
    Set TextStr = ScoreFile.OpenAsTextStream(1, 0)
    If(TextStr.AtEndOfStream = True) then
        Exit Sub
    End If
    temp1 = TextStr.ReadLine
    temp2 = textstr.readline

    HighScore = cdbl(temp1)
    If HighScore <1 then
        temp8 = textstr.readline
        temp9 = textstr.readline
        temp10 = textstr.readline
        temp11 = textstr.readline
        temp12 = textstr.readline
        temp13 = textstr.readline
        temp14 = textstr.readline
        temp15 = textstr.readline
        temp16 = textstr.readline
        temp17 = textstr.readline
    End If
    TextStr.Close
    ' No credits to load for this table, there is only one credit
  'Credits = cdbl(temp2)

    If HighScore <1 then
        HSScore(1) = int(temp8)
        HSScore(2) = int(temp9)
        HSScore(3) = int(temp10)
        HSScore(4) = int(temp11)
        HSScore(5) = int(temp12)

        HSName(1) = temp13
        HSName(2) = temp14
        HSName(3) = temp15
        HSName(4) = temp16
        HSName(5) = temp17
    End If
    Set ScoreFile = Nothing
    Set FileObj = Nothing
    SortHighscore 'added to fix a previous error
End Sub

' Save high scores on User directory of Visual Pinball
Sub Savehs
    ' Based on Black's Highscore routines
    Dim FileObj
    Dim ScoreFile
    Dim xx
    Set FileObj = CreateObject("Scripting.FileSystemObject")
    If Not FileObj.FolderExists(UserDirectory) then
        Exit Sub
    End If
    Set ScoreFile = FileObj.CreateTextFile(UserDirectory & TableName& ".txt", True)
    ScoreFile.WriteLine 0
    ' we save credits for this table, but we don't use it
  ScoreFile.WriteLine Credits
    For xx = 1 to 5
        scorefile.writeline HSScore(xx)
    Next
    For xx = 1 to 5
        scorefile.writeline HSName(xx)
    Next
    ScoreFile.Close
    Set ScoreFile = Nothing
    Set FileObj = Nothing
End Sub

Sub SortHighscore
    Dim tmp, tmp2, i, j
    For i = 1 to 5
        For j = 1 to 4
            If HSScore(j) <HSScore(j + 1) Then
                tmp = HSScore(j + 1)
                tmp2 = HSName(j + 1)
                HSScore(j + 1) = HSScore(j)
                HSName(j + 1) = HSName(j)
                HSScore(j) = tmp
                HSName(j) = tmp2
            End If
        Next
    Next
End Sub

' check the highscore
Dim PlayersPlayingGame:PlayersPlayingGame = 1
Sub CheckHighscore
    Dim playertops, si, sj, i, stemp, stempplayers
    For i = 1 to 4
        sortscores(i) = 0
        sortplayers(i) = 0
    Next
    playertops = 0
    For i = 1 to PlayersPlayingGame
        sortscores(i) = Score' Score(i) if multiple players
        sortplayers(i) = i
    Next
    For si = 1 to PlayersPlayingGame
        For sj = 1 to PlayersPlayingGame-1
            If sortscores(sj)> sortscores(sj + 1) then
                stemp = sortscores(sj + 1)
                stempplayers = sortplayers(sj + 1)
                sortscores(sj + 1) = sortscores(sj)
                sortplayers(sj + 1) = sortplayers(sj)
                sortscores(sj) = stemp
                sortplayers(sj) = stempplayers
            End If
        Next
    Next
    HighScoreTimer.interval = 100
    HighScoreTimer.enabled = True
    ScoreChecker = 4
    CheckAllScores = 1
    NewHighScore sortscores(ScoreChecker), sortplayers(ScoreChecker)
End Sub

' ***************************************************
' GNMOD - Multiple High Score Display and Collection
' jpsalas: changed ramps by flashers to remove extra shadow
'
' How to install :
' Define High scores default in HSScore and HSName variables
' Add this to Table1_KeyDown() :
' If EnteringInitials then
'     CollectInitials(keycode)
'     Exit Sub
' End If
' ***************************************************

Dim EnteringInitials ' Normally zero, set to non-zero to enter initials
EnteringInitials = False ' When true, we must give keys control to the High score script
Dim ScoreChecker
ScoreChecker = 0
Dim CheckAllScores
CheckAllScores = 0
Dim sortscores(4)
Dim sortplayers(4)

Dim PlungerPulled
PlungerPulled = 0

Dim SelectedChar   ' character under the "cursor" when entering initials

Dim HSTimerCount   ' Pass counter For HS timer, scores are cycled by the timer
HSTimerCount = 5   ' Timer is initially enabled, it'll wrap from 5 to 1 when it's displayed

Dim InitialString  ' the string holding the player's initials as they're entered

Dim AlphaString    ' A-Z, 0-9, space (_) and backspace (<)
Dim AlphaStringPos ' pointer to AlphaString, move Forward and backward with flipper keys
AlphaString = "ABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789_<"

Dim HSNewHigh      ' The new score to be recorded

Dim HSScore(5)     ' High Scores read in from config file
Dim HSName(5)      ' High Score Initials read in from config file

' default high scores, remove this when the scores are available from the config file
HSScore(1) = 3000
HSScore(2) = 4000
HSScore(3) = 5000
HSScore(4) = 6000
HSScore(5) = 7000

HSName(1) = "AAA"
HSName(2) = "ZZZ"
HSName(3) = "XXX"
HSName(4) = "ABC"
HSName(5) = "BBB"

Sub HighScoreTimer_Timer
    If EnteringInitials then
        If HSTimerCount = 1 then
            SetHSLine 3, InitialString & MID(AlphaString, AlphaStringPos, 1)
            HSTimerCount = 2
        Else
            SetHSLine 3, InitialString
            HSTimerCount = 1
        End If
    ElseIf Not GameOver then
        SetHSLine 1, "HIGH SCORE1"
        SetHSLine 2, HSScore(1)
        SetHSLine 3, HSName(1)
        HSTimerCount = 5 ' set so the highest score will show after the game is over
        HighScoreTimer.enabled = false
    ElseIf CheckAllScores then
        NewHighScore sortscores(ScoreChecker), sortplayers(ScoreChecker)
    Else
        ' cycle through high scores
        HighScoreTimer.interval = 2000
        HSTimerCount = HSTimerCount + 1
        If HsTimerCount> 5 then
            HSTimerCount = 1
        End If
        SetHSLine 1, "HIGH SCORE" + FormatNumber(HSTimerCount, 0)
        SetHSLine 2, HSScore(HSTimerCount)
        SetHSLine 3, HSName(HSTimerCount)
    End If
End Sub

Function GetHSChar(String, Index)
    Dim ThisChar
    Dim FileName
    ThisChar = Mid(String, Index, 1)
    FileName = "PostIt"
    If ThisChar = " " or ThisChar = "" then
        FileName = FileName & "BL"
    ElseIf ThisChar = "<" then
        FileName = FileName & "LT"
    ElseIf ThisChar = "_" then
        FileName = FileName & "SP"
    Else
        FileName = FileName & ThisChar
    End If
    GetHSChar = FileName
End Function

Sub SetHsLine(LineNo, String)
    Dim Letter
    Dim ThisDigit
    Dim ThisChar
    Dim StrLen
    Dim LetterLine
    Dim Index
    Dim StartHSArray
    Dim EndHSArray
    Dim LetterName
    Dim xFor
    StartHSArray = array(0, 1, 12, 22)
    EndHSArray = array(0, 11, 21, 31)
    StrLen = len(string)
    Index = 1

    For xFor = StartHSArray(LineNo) to EndHSArray(LineNo)
        Eval("HS" &xFor).imageA = GetHSChar(String, Index)
        Index = Index + 1
    Next
End Sub

Sub NewHighScore(NewScore, PlayNum)
    'debug.print "NewHighScore " & NewScore & " - player " & PlayNum
  If NewScore> HSScore(5) then
        HighScoreTimer.interval = 500
        HSTimerCount = 1
        AlphaStringPos = 1      ' start with first character "A"
        EnteringInitials = true ' intercept the control keys while entering initials
        InitialString = ""      ' initials entered so far, initialize to empty
        SetHSLine 1, "PLAYER " + FormatNumber(PlayNum, 0)
        SetHSLine 2, "ENTER NAME"
        SetHSLine 3, MID(AlphaString, AlphaStringPos, 1)
        HSNewHigh = NewScore
        PlaySound "knocker" 'but do not give an extra credit
    End If
    ScoreChecker = ScoreChecker-1
    If ScoreChecker = 0 then
        CheckAllScores = 0
    End If
End Sub

Sub CollectInitials(keycode)
    Dim i
    If keycode = LeftFlipperKey Then
        ' back up to previous character
        AlphaStringPos = AlphaStringPos - 1
        If AlphaStringPos <1 then
            AlphaStringPos = len(AlphaString) ' handle wrap from beginning to End
            If InitialString = "" then
                ' Skip the backspace If there are no characters to backspace over
                AlphaStringPos = AlphaStringPos - 1
            End If
        End If
        SetHSLine 3, InitialString & MID(AlphaString, AlphaStringPos, 1)
        PlaySound "fx_Previous"
    ElseIf keycode = RightFlipperKey Then
        ' advance to Next character
        AlphaStringPos = AlphaStringPos + 1
        If AlphaStringPos> len(AlphaString) or(AlphaStringPos = len(AlphaString) and InitialString = "") then
            ' Skip the backspace If there are no characters to backspace over
            AlphaStringPos = 1
        End If
        SetHSLine 3, InitialString & MID(AlphaString, AlphaStringPos, 1)
        PlaySound "fx_Next"
    ElseIf keycode = StartGameKey or keycode = PlungerKey Then
        SelectedChar = MID(AlphaString, AlphaStringPos, 1)
        If SelectedChar = "_" then
            InitialString = InitialString & " "
            PlaySound("fx_Esc")
        ElseIf SelectedChar = "<" then
            InitialString = MID(InitialString, 1, len(InitialString) - 1)
            If len(InitialString) = 0 then
                ' If there are no more characters to back over, don't leave the < displayed
                AlphaStringPos = 1
            End If
            PlaySound("fx_Esc")
        Else
            InitialString = InitialString & SelectedChar
            PlaySound("fx_Enter")
        End If
        If len(InitialString) <3 then
            SetHSLine 3, InitialString & SelectedChar
        End If
    End If
    If len(InitialString) = 3 then
        ' save the score
        For i = 5 to 1 step -1
            If i = 1 or(HSNewHigh> HSScore(i) and HSNewHigh <= HSScore(i - 1) ) then
                ' Replace the score at this location
                If i <5 then
                    HSScore(i + 1) = HSScore(i)
                    HSName(i + 1) = HSName(i)
                End If
                EnteringInitials = False
                HSScore(i) = HSNewHigh
                HSName(i) = InitialString
                HSTimerCount = 5
                HighScoreTimer_Timer
                HighScoreTimer.interval = 2000
                Exit Sub
            ElseIf i <5 then
                ' move the score in this slot down by 1, it's been exceeded by the new score
                HSScore(i + 1) = HSScore(i)
                HSName(i + 1) = HSName(i)
            End If
        Next
    End If
End Sub
' End GNMOD

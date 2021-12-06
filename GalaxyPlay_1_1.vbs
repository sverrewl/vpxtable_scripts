'############################################################################################
'############################################################################################
'#######                                                                             ########
'#######          Galaxy Play                                                        ########
'#######          (CIC Play 1986)                                                    ########
'#######                                                                             ########
'############################################################################################
'############################################################################################
' Version 1.0 FS mfuegemann 2017
Option Explicit
Randomize
'
' Version 1.1
' - adjusted friction values (found by JPSalas)
' - Backdrop image and settings contributed by Batch

' Thalamus 2018-07-23
' Added/Updated "Positional Sound Playback Functions" and "Supporting Ball & Sound Functions"
' Thalamus 2018-11-01 : Improved directional sounds
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
Const VolGates  = 1    ' Gates volume.
Const VolMetal  = 1    ' Metals volume.
Const VolRH     = 1    ' Rubber hits volume.
Const VolPo     = 1    ' Rubber posts volume.
Const VolPi     = 1    ' Rubber pins volume.
Const VolTarg   = 1    ' Targets volume.
Const VolSpin   = 1.5  ' Spinners volume.
Const VolFlip   = 1    ' Flipper volume.


On Error Resume Next
ExecuteGlobal GetTextFile("controller.vbs")
If Err Then MsgBox "You need the controller.vbs in order to run this table, available in the vp10 package"
On Error Goto 0

'-----------------------------------
'------  Global Cofigurations ------
'-----------------------------------
Const FreePlay=False      'set table to Free Play
Const DimGI=2       'set to dim or brighten GI lights (base value is 10)
Const ForceNoB2S=False    'set to True to skip the B2S calls - B2S errors lead to bad solenoid and lamp init sequences
Const VolumeMultiplier = 3    'adjusts table sound volume




Const cgamename = "glxplay"
Const UseSolenoids=1,UseLamps=1,SSolenoidOn="SolOn",SSolenoidOff="SolOff",SFlipperOn="FlipperUp_Akiles",SFlipperOff="FlipperDown_Akiles",SCoin="coin3"

if ForceNoB2S Then
  LoadVPMALT "01560000","mac.VBS",3.2
else
  LoadVPM "01560000","mac.VBS",3.2
end If

BallSize = 26           'Ball radius

'-----------------------------------
'------  Solenoid Assignment  ------
'-----------------------------------
SolCallback(8)="bsTrough.SolOut"        'OK
SolCallback(7)="Sol7"             'Drop Target Bank 'OK
SolCallback(3)="bsRightHole.SolOut"       'OK
SolCallback(10)="Sol10"             'GameOn OK
SolCallback(17)="vpmSolSound ""knocker"","

' Fix from Arngrim
' https://www.vpforums.org/index.php?showtopic=39101&page=2

'SolCallback(sLLFlipper)="vpmSolFlipper LeftFlipper,Nothing,"
'SolCallback(sLRFlipper)="vpmSolFlipper RightFlipper,Nothing,"

SolCallback(sLRFlipper) = "SolRFlipper"
SolCallback(sLLFlipper) = "SolLFlipper"

Sub SolLFlipper(Enabled)
  If Enabled Then
    PlaySound SoundFX("FlipperUp",DOFContactors):LeftFlipper.RotateToEnd
  Else
    PlaySound SoundFX("FlipperDown",DOFContactors):LeftFlipper.RotateToStart
  End If
End Sub

Sub SolRFlipper(Enabled)
  If Enabled Then
    PlaySound SoundFX("FlipperUp",DOFContactors):RightFlipper.RotateToEnd
  Else
    PlaySound SoundFX("FlipperDown",DOFContactors):RightFlipper.RotateToStart
  End If
End Sub


Sub Sol7(enabled)
  if enabled then
    DropTargetBank.DropSol_On
  end if
End Sub

Sub Sol10(enabled)
  for each obj in GI
    obj.State = abs(enabled)
  next
End Sub

'--------------------------
'------  Table Init  ------
'--------------------------
Dim bsTrough,bsRightHole,DropTargetBank,cRightCaptive,Flipperactive,obj,cLCaptive,cRCaptive
Dim DesktopMode: DesktopMode = GalaxyPlay.ShowDT

If DesktopMode = True Then 'Show Desktop components
  Ramp16.visible=1
  Ramp15.visible=1
  SideWood.visible=1
  for each obj in DTLights
    obj.visible = True
  Next
Else
  Ramp16.visible=0
  Ramp15.visible=0
  SideWood.visible=0
  for each obj in DTLights
    obj.visible = False
  Next
End if

for each obj in GI
  obj.intensity = obj.intensity + DimGI
next

Sub GalaxyPlay_Init
  vpminit me
    vpmMapLights AllLights
  cWall.isdropped = True

    Controller.GameName=cGameName
    Controller.SplashInfoLine="Galaxy Play" & vbNewLine & "created by mfuegemann"
    Controller.HandleKeyboard=False
    Controller.ShowTitle=0
    Controller.ShowFrame=0
    Controller.ShowDMDOnly=1
  'Controller.Hidden = 1      'enable to hide DMD if You use a B2S backglass

'    'DMD position for 3 Monitor Setup
'    'Controller.Games(cGameName).Settings.Value("dmd_pos_x")=3850    'set this to 0 if You cannot find the DMD
'    'Controller.Games(cGameName).Settings.Value("dmd_pos_y")=300   'set this to 0 if You cannot find the DMD
'    'Controller.Games(cGameName).Settings.Value("dmd_width")=505
'    'Controller.Games(cGameName).Settings.Value("dmd_height")=155
'    'Controller.Games(cGameName).Settings.Value("rol")=0
'
' 'Controller.Games(cGameName).Settings.Value("ddraw") = 0             'set to 0 if You have problems with DMD showing or table stutter

    Controller.HandleMechanics=0

' vpmTimer.AddTimer 2000, "Controller.SolMask(0)=&Hffffffff'" 'ignore all solenoids - then add the Timer to renable all the solenoids after 2 seconds

    Controller.Run
    If Err Then MsgBox Err.Description
    On Error Goto 0

    PinMAMETimer.Interval=PinMAMEInterval
  PinMAMETimer.Enabled = true

    vpmNudge.TiltSwitch=5
    vpmNudge.Sensitivity=5
  vpmNudge.TiltObj = Array(LeftFlipper,RightFlipper,Bumper1,Bumper2,Bumper3)

   Set bsTrough=New cvpmBallStack
        bsTrough.InitSw 0,9,0,0,0,0,0,0     '0.1 = Switch 0
        bsTrough.InitKick BallRelease,90,5
        bsTrough.InitExitSnd SoundFX("BallRel",DOFContactors),SoundFX("Solenoid",DOFContactors)
        bsTrough.Balls=1

  set DropTargetBank = new cvpmDropTarget
    DropTargetBank.InitDrop Array(DT1,DT2,DT3,DT4), Array(33,37,38,39)
    DropTargetBank.InitSnd SoundFX("Targetdrop1",DOFContactors),SoundFX("TargetBankreset1",DOFContactors)
    DropTargetBank.CreateEvents "DropTargetBank"

  Set bsRightHole = New cvpmBallStack
    bsRightHole.InitExitSnd SoundFX("Popper",DOFContactors),SoundFX("Solenoid",DOFContactors)
    bsRightHole.initsaucer RHole,10,230,15

  Set cLCaptive=New cvpmCaptiveBall
    cLCaptive.InitCaptive LCaptiveTrigger,LCaptiveWall,LCaptiveKicker,45
    controller.switch(34) = 1
    cLCaptive.Start
    cLCaptive.ForceTrans = 0.7
    cLCaptive.MinForce = 1.5
    cLCaptive.CreateEvents "cLCaptive"
    LCaptiveKicker.destroyball

  Set cRCaptive=New cvpmCaptiveBall
    cRCaptive.InitCaptive RCaptiveTrigger,RCaptiveWall,RCaptiveKicker,-45
    cRCaptive.Start
    cRCaptive.ForceTrans = 0.7
    cRCaptive.MinForce = 1.5
    cRCaptive.CreateEvents "cRCaptive"
    RCaptiveKicker.destroyball

  CaptiveInitKicker.createBall
  CaptiveInitKicker.kick 180,5

End Sub

Sub GalaxyPlay_Exit()
  Controller.Pause = False
  Controller.Stop
End Sub

'------------------------------
'------  Trough Handler  ------
'------------------------------
Sub Drain_Hit()
  bsTrough.AddBall Me
  playsoundAtVol "Drain5", drain, 1
End Sub


'------------------------------
'------  Switch Handler  ------
'------------------------------
Dim CBall
Sub cTrigger_Hit
  Set CBall = Activeball
End Sub

Sub cTrigger_UnHit
  cWall.isdropped = False

  if CBall.x < Captiveinitkicker.x then
    controller.switch(34) = 0
    controller.switch(41) = 1
  else
    controller.switch(34) = 1
    controller.switch(41) = 0
  end if
End Sub


Sub RHole_Hit:bsRightHole.addball 0:End Sub


'sub Trigger1_hit:Controller.Switch(9)=1:End Sub
'sub Trigger1_unhit:Controller.Switch(9)=0:End Sub

'Top Rollover
sub TopRollover_Left_hit:Controller.Switch(29)=1:End Sub
sub TopRollover_Left_unhit:Controller.Switch(29)=0:End Sub
sub TopRollover_Middle_hit:Controller.Switch(30)=1:End Sub
sub TopRollover_Middle_unhit:Controller.Switch(30)=0:End Sub
sub TopRollover_Right_hit:Controller.Switch(31)=1:End Sub
sub TopRollover_Right_unhit:Controller.Switch(31)=0:End Sub

Sub Bumper1_Hit:vpmTimer.PulseSw 12:PlaySoundAtVol SoundFX("Jet1",DOFContactors),Bumper1, VolBump:End Sub
Sub Bumper2_Hit:vpmTimer.PulseSw 13:PlaySoundAtVol SoundFX("Jet1",DOFContactors),Bumper2, VolBump:End Sub
Sub Bumper3_Hit:vpmTimer.PulseSw 11:PlaySoundAtVol SoundFX("Jet2",DOFContactors),Bumper3, VolBump:End Sub

sub LOutlane_hit:Controller.Switch(25)=1:End Sub
sub LOutlane_unhit:Controller.Switch(25)=0:End Sub
sub LInlane_hit:Controller.Switch(26)=1:End Sub
sub LInlane_unhit:Controller.Switch(26)=0:End Sub

sub ROutlane_hit:Controller.Switch(28)=1:End Sub
sub ROutlane_unhit:Controller.Switch(28)=0:End Sub
sub RInlane_hit:Controller.Switch(27)=1:End Sub
sub RInlane_unhit:Controller.Switch(27)=0:End Sub

sub RPassage_hit:Controller.Switch(44)=1:End Sub
sub RPassage_unhit:Controller.Switch(44)=0:End Sub

Sub Rubber_17_Hit:vpmTimer.PulseSw 17:End Sub
Sub Rubber_18_Hit:vpmTimer.PulseSw 18:End Sub
Sub Rubber_19_Hit:vpmTimer.PulseSw 19:End Sub
Sub Rubber_20_Hit:vpmTimer.PulseSw 20:End Sub
Sub Rubber_21_Hit:vpmTimer.PulseSw 21:End Sub
Sub Rubber_22_Hit:vpmTimer.PulseSw 22:End Sub
Sub Rubber_23_Hit:vpmTimer.PulseSw 23:End Sub
Sub Rubber_44_Hit:vpmTimer.PulseSw 44:End Sub

sub LeftTarget_hit:vpmTimer.PulseSw 35:End Sub

sub RightTarget_hit:vpmTimer.PulseSw 36:End Sub

Sub LaunchTrigger_Hit
  if activeball.vely < -4 Then
    PlaysoundAtVol "Launch", ActiveBall, 1
  end if
End Sub

'-------------------------------
'------  Keybord Handler  ------
'-------------------------------

Sub GalaxyPlay_KeyDown(ByVal keycode)
  If keycode = PlungerKey Then
    Plunger.PullBack
    PlaysoundAtVol "PlungerPull", Plunger, 1
  End If
  if keycode = StartGameKey and FreePlay then
    vpmTimer.PulseSw 3
  end if
  If vpmKeyDown(KeyCode) Then Exit Sub
End Sub

Sub GalaxyPlay_KeyUp(ByVal keycode)
  If keycode = PlungerKey Then
    Plunger.Fire
    PlaysoundAtVol "Plunger", Plunger ,1
  End If
  If vpmKeyUp(KeyCode) Then Exit Sub
End Sub


'-----------------------------------
' Flipper Primitives
'-----------------------------------
sub FlipperTimer_Timer()
  DT4Wall.isdropped = DT4.isdropped

'*****************************************
' ninuzzu's FLIPPER SHADOWS
'*****************************************
  FlipperLSh.RotZ = LeftFlipper.currentangle
  FlipperRSh.RotZ = RightFlipper.currentangle
end sub


'**********Sling Shot Animations
' Rstep and Lstep  are the variables that increment the animation
'****************
Dim RStep, Lstep

Sub RightSlingShot_Slingshot
  vpmtimer.pulsesw 15
    PlaySoundAtVol SoundFX("rsling",DOFContactors), sling1, 1
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
  vpmtimer.pulsesw 14
    PlaySoundAtVol SoundFX("lsling",DOFContactors), sling2, 1
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


'-----------------------------------
' DIP Switch Menu
'-----------------------------------

Sub editDips
  Dim vpmDips:Set vpmDips=New cvpmDips
  With vpmDips
    .AddForm 400,200,"Galaxy Play - DIP switches"
    .AddFrame 0,0,110,"Balls per game",&H00000001,Array("5 Balls",0,"3 Balls",&H00000001)'dip 1 (A)
'   .AddFrame 0,50,110,"Coin Chute settings",&H00000006,Array("1/2 - 3",0,"1/2 - 4",&H00000002,"1 - 5",&H00000006)'dip 2&3 (A)
'   .AddChk 0,115,110,Array("Enable Music",&H00000020)'dip 6 (A)
    .AddLabel 0,140,280,20,"After hitting OK, press F3 to reset game with new settings."
    .ViewDips
  End With
End Sub
Set vpmShowDips=GetRef("editDips")


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
  PlaySoundAtVol "fx_spinner", Spinner, VolSpin
End Sub

Sub Rubbers_Hit(idx)
  dim finalspeed
    finalspeed=SQR(activeball.velx * activeball.velx + activeball.vely * activeball.vely)
  If finalspeed > 20 then
    PlaySound "fx_rubber2", 0, Vol(ActiveBall)*VolRH, Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
  End if
  If finalspeed >= 6 AND finalspeed <= 20 then
    RandomSoundRubber()
  End If
End Sub

Sub Posts_Hit(idx)
  dim finalspeed
    finalspeed=SQR(activeball.velx * activeball.velx + activeball.vely * activeball.vely)
  If finalspeed > 16 then
    PlaySound "fx_rubber2", 0, Vol(ActiveBall)*VolPo, Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
  End if
  If finalspeed >= 6 AND finalspeed <= 16 then
    RandomSoundRubber()
  End If
End Sub

Sub RandomSoundRubber()
  Select Case Int(Rnd*3)+1
    Case 1 : PlaySound "rubber_hit_1", 0, Vol(ActiveBall)*VolRH, Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
    Case 2 : PlaySound "rubber_hit_2", 0, Vol(ActiveBall)*VolRH, Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
    Case 3 : PlaySound "rubber_hit_3", 0, Vol(ActiveBall)*VolRH, Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
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
    Case 1 : PlaySound "flip_hit_1", 0, Vol(ActiveBall)*VolRH, Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
    Case 2 : PlaySound "flip_hit_2", 0, Vol(ActiveBall)*VolRH, Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
    Case 3 : PlaySound "flip_hit_3", 0, Vol(ActiveBall)*VolRH, Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
  End Select
End Sub

' *******************************************************************************************************
' Positional Sound Playback Functions by DJRobX
' PlaySound sound, 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 0, 1, AudioFade(ActiveBall)
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
  PlaySound sound, 1, Volume, Pan(tableobj), 0,0,0, 1, AudioFade(tableobj)
End Sub

'Set all as per ball position & speed, but Vol Multiplier may be used eg; PlaySoundAtBallVol "sound",3

Sub PlaySoundAtBallVol(sound, VolMult)
  PlaySound sound, 0, Vol(ActiveBall) * VolMult, Pan(ActiveBall), 0, Pitch(ActiveBall), 0, 1, AudioFade(ActiveBall)
End Sub

'Set position as bumperX and Vol manually.

Sub PlaySoundAtBumperVol(sound, tableobj, Vol)
  PlaySound sound, 1, Vol, Pan(tableobj), 0,0,1, 1, AudioFade(tableobj)
End Sub

'*********************************************************************
'                     Supporting Ball & Sound Functions
'*********************************************************************

Function AudioFade(tableobj) ' Fades between front and back of the table (for surround systems or 2x2 speakers, etc), depending on the Y position on the table. "GalaxyPlay" is the name of the table
  Dim tmp
  tmp = tableobj.y * 2 / GalaxyPlay.height-1
  If tmp > 0 Then
    AudioFade = Csng(tmp ^10)
  Else
    AudioFade = Csng(-((- tmp) ^10) )
  End If
End Function

Function AudioPan(tableobj) ' Calculates the pan for a tableobj based on the X position on the table. "GalaxyPlay" is the name of the table
  Dim tmp
  tmp = tableobj.x * 2 / GalaxyPlay.width-1
  If tmp > 0 Then
    AudioPan = Csng(tmp ^10)
  Else
    AudioPan = Csng(-((- tmp) ^10) )
  End If
End Function

Function Pan(ball) ' Calculates the pan for a ball based on the X position on the table. "GalaxyPlay" is the name of the table
    Dim tmp
    tmp = ball.x * 2 / GalaxyPlay.width-1
    If tmp > 0 Then
        Pan = Csng(tmp ^10)
    Else
        Pan = Csng(-((- tmp) ^10) )
    End If
End Function

Function AudioFade(ball) ' Can this be together with the above function ?
  Dim tmp
  tmp = ball.y * 2 / GalaxyPlay.height-1
  If tmp > 0 Then
    AudioFade = Csng(tmp ^10)
  Else
    AudioFade = Csng(-((- tmp) ^10) )
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
          PlaySound("fx_ballrolling" & b), -1, Vol(BOT(b) ), Pan(BOT(b) ), 0, Pitch(BOT(b) ), 1, 0, AudioFade(BOT(b) )
        Else ' Ball on raised ramp
          PlaySound("fx_ballrolling" & b), -1, Vol(BOT(b) )*.5, Pan(BOT(b) ), 0, Pitch(BOT(b) )+50000, 1, 0, AudioFade(BOT(b) )
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
    PlaySound("fx_collide"), 0, Csng(velocity) ^2 / (VolDiv/VolCol), Pan(ball1), 0, Pitch(ball1), 0, 0, AudioFade(ball1)
End Sub

' Thalamus : Exit in a clean and proper way
Sub GalaxyPlay_exit()
  Controller.Pause = False
  Controller.Stop
End Sub


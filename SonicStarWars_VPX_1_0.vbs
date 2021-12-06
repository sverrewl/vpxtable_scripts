'############################################################################################
'############################################################################################
'#######                                                                             ########
'#######          Star Wars                                                          ########
'#######          (Sonic 1987)                                                       ########
'#######                                                                             ########
'############################################################################################
'############################################################################################
' Version 1.0 FS mfuegemann 2018
Option Explicit

' Thanks to
' Akiles for providing the playfield image and plastics scans
' TAB for the VP8 version that gave me a good start with the solenoid, lamp and switch numbers

' Thalamus 2018-07-24
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

Const VolBump   = .5   ' Bumpers volume.
Const VolRol    = 1    ' Rollovers volume.
Const VolGates  = 1    ' Gates volume.
Const VolMetal  = 1    ' Metals volume.
Const VolRB     = 1    ' Rubber bands volume.
Const VolRH     = 1    ' Rubber hits volume.
Const VolPo     = 1    ' Rubber posts volume.
Const VolPi     = 1    ' Rubber pins volume.
Const VolPlast  = 1    ' Plastics volume.
Const VolTarg   = 1    ' Targets volume.
Const VolWood   = 1    ' Woods volume.
Const VolKick   = 1    ' Kicker volume.
Const VolSpin   = .5   ' Spinners volume.
Const VolFlip   = 1    ' Flipper volume.


On Error Resume Next
ExecuteGlobal GetTextFile("controller.vbs")
If Err Then MsgBox "You need the controller.vbs in order to run this table, available in the vp10 package"
On Error Goto 0

'-----------------------------------
'------  Global Cofigurations ------
'-----------------------------------
Const FreePlay=False      'set table to Free Play
Const DimGI=0         'set to dim or brighten GI lights (base value is 10)
Const VolumeMultiplier = 3    'adjusts table sound volume
Const OutlanePostPosition = 0 '0=easy, 1=difficult

Const TopScoopHelper=False    'enable to use helper kicker instead of the collidable Ramp-Up scoop at the right ramp

Const ForceNoB2S=False      'set to True to skip the B2S calls - B2S errors lead to bad solenoid and lamp init sequences

Const cgamename = "sonstwar"
Const UseSolenoids=1,UseLamps=1,SSolenoidOn="SolOn",SSolenoidOff="SolOff",SFlipperOn="FlipperUp_Akiles",SFlipperOff="FlipperDown_Akiles",SCoin="coin3"

if ForceNoB2S Then
  LoadVPMALT "01560000","peyper.VBS",3.2
else
  LoadVPM "01560000","peyper.VBS",3.2
end If

if OutlanePostPosition = 1 Then
  RightOutlane_Dificil.isdropped = False
  RightOutlane_Facil.isdropped = True
  P_RightOutlaneCap.transx = 0
  P_RightOutlanePost.transx = 0
  P_RightOutlaneRubber_Facil.visible = False
  P_RightOutlaneRubber_Dificil.visible = True

  LeftOutlane_Dificil.isdropped = False
  LeftOutlane_Facil.isdropped = True
  P_LeftOutlanePost.transx = 0
  P_LeftOutlanePost.transz = 0
  P_LeftOutlaneRubber_Facil.visible = False
  P_LeftOutlaneRubber_Dificil.visible = True
Else
  RightOutlane_Dificil.isdropped = True
  RightOutlane_Facil.isdropped = False
  P_RightOutlaneCap.transx = -14
  P_RightOutlanePost.transx = -14
  P_RightOutlaneRubber_Facil.visible = True
  P_RightOutlaneRubber_Dificil.visible = False

  LeftOutlane_Dificil.isdropped = True
  LeftOutlane_Facil.isdropped = False
  P_LeftOutlanePost.transx = 11
  P_LeftOutlanePost.transz = -5
  P_LeftOutlaneRubber_Facil.visible = True
  P_LeftOutlaneRubber_Dificil.visible = False
End If

if TopScoopHelper Then
  TopScoop.enabled = True
  P_TopScoop.collidable = False
  SW3.enabled = False
else
  TopScoop.enabled = False
  P_TopScoop.collidable = True
  SW3.enabled = True
end If

'-----------------------------------
'------  Solenoid Assignment  ------
'-----------------------------------

'SolCallback(1) - Bumper
'SolCallback(2) - Bumper
'SolCallback(3) - Bumper
'SolCallback(4) - Sling
'SolCallback(5) - Sling
SolCallback(6)="Sol6"   'Ramp Up
SolCallback(7)="Sol7"   'Ramp Down
SolCallback(8)="vpmSolSound ""knocker"","
SolCallback(9)="vpmFlasher Array(Laser1,Laser2),"    'left Laser
SolCallback(10)="vpmFlasher Array(Laser3,Laser4),"    'right Laser
SolCallback(11)="bsTrough.SolOut"
SolCallback(30)="Sol_GameOn"

Sub Sol_GameOn(enabled)
  VpmNudge.SolGameOn enabled
  Flipperactive = enabled
  for each obj in GI
    obj.State = abs(enabled)
  next
end sub

'Ramp Up
Sub Sol6(enabled)
  PlaysoundAtVol SoundFX("solon",DOFContactors), P_BridgeRamp, .5
  Ramp_Down.collidable = False
  if RampLimit <> RampUp then
    P_BridgeRamp.Rotx = RampUp + 15
  end if
  RampLimit = RampUp
  CycloneRampTimer.enabled = True
  Primitive_MechArm.Rotx = 120
End Sub

'Ramp Down
Sub Sol7(enabled)
  PlaysoundAtVol SoundFX("solon",DOFContactors), Primitive_MechArm, .5
  Ramp_Down.collidable = True
  RampLimit = RampDown
  CycloneRampTimer.enabled = True
  Primitive_MechArm.Rotx = 80
End Sub

'Cyclone Ramp Animation
Dim RampLimit
Const RampDown = -14
Const RampUp = 2

Sub CycloneRampTimer_Timer
  P_BridgeRamp.Rotx = P_BridgeRamp.Rotx - 1.5
  if P_BridgeRamp.Rotx <= RampLimit then
    CycloneRampTimer.enabled = False
    P_BridgeRamp.Rotx = RampLimit
  end if
End Sub

Ramp_Down.collidable = True
P_BridgeRamp.Rotx = RampDown
Primitive_MechArm.Rotx = 80


SolCallback(sLLFlipper)="vpmSolFlipper LeftFlipper,Nothing,"
SolCallback(sLRFlipper)="vpmSolFlipper RightFlipper,URightFlipper,"


'--------------------------
'------  Table Init  ------
'--------------------------
Dim bsTrough,LeftDropTargetBank,RightDropTargetBank,cCaptive,Flipperactive,obj
Dim DesktopMode: DesktopMode = SSW.ShowDT

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

Sub SSW_Init
  vpminit me
    vpmMapLights AllLights

    Controller.GameName=cGameName
    Controller.SplashInfoLine="Sonic Star Wars" & vbNewLine & "created by mfuegemann"
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

    vpmNudge.TiltSwitch=-5
    vpmNudge.Sensitivity=5
  vpmNudge.TiltObj = Array(Bumper1,Bumper2,Bumper3,LeftSlingshot,RightSlingshot)

    Set bsTrough=New cvpmBallStack
        bsTrough.InitSw 0,0.1,0,0,0,0,0,0     '0.1 = Switch 0
        bsTrough.InitKick BallRelease,90,5
        bsTrough.InitExitSnd SoundFX("BallRel",DOFContactors),SoundFX("Solenoid",DOFContactors)
        bsTrough.Balls=1

End Sub

Sub SSW_Exit()
  Controller.Pause = False
  Controller.Stop
End Sub

'------------------------------
'------  Trough Handler  ------
'------------------------------
Sub Drain_Hit()
  bsTrough.AddBall Me
  PlaySoundAtVol "drain5", drain, 1
End Sub


'------------------------------
'------  Switch Handler  ------
'------------------------------
Sub TopScoop_Hit
  if activeball.vely > -4 Then
    TopScoop.Kick 180,2
  Else
    vpmTimer.PulseSw 27
    TopScoop.DestroyBall
    TopScoopExit.CreateBall
    TopScoopExit.Kick 180,4
  end if
End Sub

Sub RubberSW1_Hit:vpmtimer.pulsesw(23):End Sub
Sub RubberSW2_Hit:vpmtimer.pulsesw(24):End Sub
Sub RubberSW6_Hit:vpmtimer.pulsesw(30):End Sub
Sub RubberSW16_Hit:vpmtimer.pulsesw(32):End Sub

Sub Target10_Hit:vpmtimer.pulsesw(20):End Sub
Sub Target11_Hit:vpmtimer.pulsesw(21):End Sub
Sub Target12_Hit:vpmtimer.pulsesw(22):End Sub
Sub Target15_Hit:vpmtimer.pulsesw(11):End Sub
Sub Target17_Hit:vpmtimer.pulsesw(12):End Sub
Sub Target19_Hit:vpmtimer.pulsesw(13):End Sub
Sub Target20_Hit:vpmtimer.pulsesw(14):End Sub
Sub Target21_Hit:vpmtimer.pulsesw(16):End Sub
Sub Target22_Hit:vpmtimer.pulsesw(17):End Sub
Sub Target23_Hit:vpmtimer.pulsesw(10):End Sub

Sub Spinner14_Spin:vpmtimer.pulsesw(6):PlaySoundAtVol "soloff", Spinner14, VolSpin:End Sub
Sub Spinner18_Spin:vpmtimer.pulsesw(7):PlaySoundAtVol "soloff", Spinner18, VolSpin:End Sub

Sub SW13_Hit:Controller.Switch(31) = 1:End Sub
Sub SW13_unhit:Controller.Switch(31) = 0:End Sub
Sub SW30_Hit:Controller.Switch(37) = 1:End Sub
Sub SW30_unhit:Controller.Switch(37) = 0:End Sub

Sub SW3_Hit:Controller.Switch(27) = 1:End Sub
Sub SW3_unhit:Controller.Switch(27) = 0:PlaySoundAtVol "WireRampR",ActiveBall,.5:End Sub
Sub SW4_Hit:Controller.Switch(25) = 1:End Sub
Sub SW4_unhit:Controller.Switch(25) = 0:End Sub
Sub SW5_Hit:Controller.Switch(26) = 1:LightCyclotronLamps:End Sub
Sub SW5_unhit:Controller.Switch(26) = 0:End Sub

Sub Bumper1_Hit:vpmtimer.pulsesw(1):PlaySoundAtVol SoundFX("fx_bumper4",DOFContactors),ActiveBall,VolBump:End Sub
Sub Bumper2_Hit:vpmtimer.pulsesw(2):PlaySoundAtVol SoundFX("fx_bumper4",DOFContactors),ActiveBall,VolBump:End Sub
Sub Bumper3_Hit:vpmtimer.pulsesw(3):PlaySoundAtVol SoundFX("fx_bumper4",DOFContactors),ActiveBall,VolBump:End Sub

Sub LInlane_Hit:Controller.Switch(35) = 1:End Sub
Sub LInlane_Unhit:Controller.Switch(35) = 0:End Sub
Sub LOutlane_Hit:Controller.Switch(36) = 1:End Sub
Sub LOutlane_Unhit:Controller.Switch(36) = 0:End Sub
Sub RInlane_Hit:Controller.Switch(34) = 1:End Sub
Sub RInlane_Unhit:Controller.Switch(34) = 0:End Sub
Sub ROutlane_Hit:Controller.Switch(33) = 1:End Sub
Sub ROutlane_Unhit:Controller.Switch(33) = 0:End Sub

Sub LaunchTrigger_Hit
  if activeball.vely < -4 Then
    PlaysoundAtVol "Launch", LaunchTrigger, 1
  end if
End Sub

'-------------------------------
'------  Keybord Handler  ------
'-------------------------------

Sub SSW_KeyDown(ByVal keycode)
  If keycode = PlungerKey Then
    Plunger.PullBack
    PlaysoundAtVol "PlungerPull", Plunger, 1
  End If
  if keycode = StartGameKey and FreePlay then
    vpmTimer.PulseSw -3
  end if
  If KeyCode=LeftFlipperKey Then Controller.Switch(103)=1
  If KeyCode=RightFlipperKey Then Controller.Switch(101)=1
  If vpmKeyDown(KeyCode) Then Exit Sub
End Sub

Sub SSW_KeyUp(ByVal keycode)
  If keycode = PlungerKey Then
    Plunger.Fire
    PlaysoundAtVol "Plunger", Plunger, 1
  End If

  If KeyCode=LeftFlipperKey Then Controller.Switch(103)=0
  If KeyCode=RightFlipperKey Then Controller.Switch(101)=0
  If vpmKeyUp(KeyCode) Then Exit Sub
End Sub

Sub LightCyclotronLamps
  for each obj in Cyclotronlamps
    obj.state = Lightstateon
  Next
  CyclotronTimer.enabled = True
End Sub

Sub CyclotronTimer_Timer
  CyclotronTimer.enabled = False
  for each obj in Cyclotronlamps
    obj.state = Lightstateoff
  Next
End Sub



'-----------------------------------
' Flipper Primitives
'-----------------------------------
sub FlipperTimer_Timer()

'*****************************************
' ninuzzu's FLIPPER SHADOWS
'*****************************************
  FlipperLSh.RotZ = LeftFlipper.currentangle
  FlipperRSh.RotZ = RightFlipper.currentangle
  FlipperURSh.RotZ = URightFlipper.currentangle
end sub


'**********Sling Shot Animations
' Rstep and Lstep  are the variables that increment the animation
'****************
Dim RStep, Lstep

Sub RightSlingShot_Slingshot
  vpmtimer.pulsesw 5
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
  vpmtimer.pulsesw 4
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

'Sonic Star Wars
Sub editDips
  Dim vpmDips:Set vpmDips=New cvpmDips
  With vpmDips
    .AddForm 700,280,"Sonic Star Wars - DIP switches"
    .AddFrame 205,0,190,"Coins per game",&H00000018,Array("(3x)1-2",0,"(2x)1-3",&H00000008,"(3x)2-4",&H00000010,"1-5",&H00000018)'dip 4&5
    .AddFrame 205,75,190,"Handicap",192,Array("3,300,000 points",192,"3,600,000 points",&H00000040,"3,900,000 points",&H00000080,"4,200,000 points",0)'dip 7&8
    .AddFrame 205,150,190,"Balls per game",&H00000020,Array("3 balls",&H00000020,"5 balls",0)'dip 6
    .AddFrame 0,0,190,"Outlane Special lit after",&H00000600,Array("1 attack",&H00000600,"2 attacks",&H00000200,"3 attacks",&H00000400,"4 attacks",0)'dip 10&11
    .AddFrame 0,75,190,"Extra ball (lit with right ramp)",&H00001800,Array("STAR targets completed once",&H00001800,"2 times STAR targets completed",&H00000800,"3 times STAR targets completed",&H00001000,"4 times STAR targets completed",0)'dip 12&13
    .AddChk 0,160,180,Array("Match feature disabled",&H00000002) 'dip 2
    .AddChk 0,175,180,Array("Attract mode (not working)",&H00000004)'dip 3
    .AddChk 0,190,180,Array("Erase memory",&H00000100)'dip 9
    .AddChk 0,205,180,Array("TEST disabled (must be ON)",&H00002000)'dip 14
    .AddLabel 50,230,280,20,"After hitting OK, press F3 to reset game with new settings."
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

Sub URightFlipper_Collide(parm)
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

Sub PlaySoundAtVol(sound, tableobj, Volum)
  PlaySound sound, 1, Volum, Pan(tableobj), 0,0,0, 1, AudioFade(tableobj)
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

Function AudioFade(tableobj) ' Fades between front and back of the table (for surround systems or 2x2 speakers, etc), depending on the Y position on the table. "SSW" is the name of the table
  Dim tmp
  tmp = tableobj.y * 2 / SSW.height-1
  If tmp > 0 Then
    AudioFade = Csng(tmp ^10)
  Else
    AudioFade = Csng(-((- tmp) ^10) )
  End If
End Function

Function AudioPan(tableobj) ' Calculates the pan for a tableobj based on the X position on the table. "SSW" is the name of the table
  Dim tmp
  tmp = tableobj.x * 2 / SSW.width-1
  If tmp > 0 Then
    AudioPan = Csng(tmp ^10)
  Else
    AudioPan = Csng(-((- tmp) ^10) )
  End If
End Function

Function Pan(ball) ' Calculates the pan for a ball based on the X position on the table. "SSW" is the name of the table
    Dim tmp
    tmp = ball.x * 2 / SSW.width-1
    If tmp > 0 Then
        Pan = Csng(tmp ^10)
    Else
        Pan = Csng(-((- tmp) ^10) )
    End If
End Function

Function AudioFade(ball) ' Can this be together with the above function ?
  Dim tmp
  tmp = ball.y * 2 / SSW.height-1
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
Sub SSW_exit()
  Controller.Pause = False
  Controller.Stop
End Sub


'RAT FINK (2014) -  rethemed Star Trek - Bally 1979
'VP9 original table by jpsalas
'VP10 table layout/code by rothbauerw
'pop bumper code and mechanics by BorgDog
'VP10 retheme by HauntFreaks (2016)
Option Explicit
Randomize


'NOT FOR RESALE OR COMMERCIAL USE!!

'Version 1.0

' Thalamus 2018-07-24
' Added/Updated "Positional Sound Playback Functions" and "Supporting Ball & Sound Functions"
' Thalamus 2018-11-01 : Improved directional sounds
' Changed useSolenoids=2 - thanks Carny
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
Const VolSpin   = 1.5  ' Spinners volume.
Const VolFlip   = 1    ' Flipper volume.


Dim cBall1, cBall2, cBall3, direction1, direction2, direction3
Dim bMusicOn


On Error Resume Next
ExecuteGlobal GetTextFile("controller.vbs")
If Err Then MsgBox "You need the controller.vbs in order to run this table, available in the vp10 package"
On Error Goto 0

Const cGameName="ratfink",UseSolenoids=2,UseLamps=1,UseGI=0,SSolenoidOn="SolOn",SSolenoidOff="SolOff",SFlipperOn="fx_Flipperup",SFlipperOff="fx_Flipperdown"
'Const cGameName="startreb",UseSolenoids=1,UseLamps=1,UseGI=0,SSolenoidOn="SolOn",SSolenoidOff="SolOff",SFlipperOn="fx_Flipperup",SFlipperOff="fx_Flipperdown"
Const SCoin="Coin",cCredits=""

LoadVPM "01550000", "Bally.vbs", 3.26
Dim DesktopMode: DesktopMode = Table1.ShowDT


'Initiate Table
'**********************************************************************************************************
Dim bsTrough,dtBank,bsSaucer

Sub Table1_Init
  vpmInit Me
  On Error Resume Next
    With Controller
    .GameName = cGameName
    If Err Then MsgBox "Can't start Game" & cGameName & vbNewLine & Err.Description : Exit Sub
    .SplashInfoLine="Star Trek (Bally 1979)"
    .HandleMechanics=0
    .HandleKeyboard=0
    .ShowDMDOnly=1
    .ShowFrame=0
    .ShowTitle=0
    .Hidden = 1
    If Err Then MsgBox Err.Description
  End With
  On Error Goto 0
    Controller.SolMask(0)=0
      vpmTimer.AddTimer 2000,"Controller.SolMask(0)=&Hffffffff'" 'ignore all solenoids - then add the timer to renable all the solenoids after 2 seconds
    Controller.Run
  If Err Then MsgBox Err.Description
  On Error Goto 0

  PinMAMETimer.Interval=PinMAMEInterval
  PinMAMETimer.Enabled=1


     ' Nudging
     vpmNudge.TiltSwitch = 7
     vpmNudge.Sensitivity = 1
     vpmNudge.TiltObj = Array(Bumper2, Bumper3, Bumper1, LeftSlingshot, RightSlingshot)

     ' Trough
     Set bsTrough = New cvpmBallStack
     With bsTrough
         .InitSw 0, 8, 0, 0, 0, 0, 0, 0
         .InitKick BallRelease, 90, 7
         .InitEntrySnd SoundFX("Solenoid",DOFContactors), SoundFX("Solenoid",DOFContactors)
         .InitExitSnd SoundFX("ballrelease",DOFContactors), SoundFX("Solenoid",DOFContactors)
         .Balls = 1
     End With

     ' Drop targets
     set dtBank = new cvpmdroptarget
     With dtBank
         .initdrop array(sw1, sw2, sw3, sw4), array(1, 2, 3, 4)
         .initsnd SoundFX("droptarget",DOFContactors), SoundFX("resetdrop",DOFContactors)
     End With

     ' Saucer
     Set bsSaucer = New cvpmBallStack
     With bsSaucer
         .InitSaucer Kicker, 32, 202, 16
         .InitExitSnd SoundFX("HoleKick",DOFContactors),SoundFX("SolOn",DOFContactors)
         .InitExitSnd SoundFX("wild",DOFContactors),SoundFX("SolOn",DOFContactors)
         .KickAngleVar = 2
         .KickForceVar = 2
     End With
     sw32a.IsDropped = 1


  cBallInit
    bMusicOn = TRUE 'TRUE

End Sub

Sub cBallInit
  set cBall1 = captiveye1.createsizedball(20)
  cBall1.image = "black"
  cBall1.frontdecal = "76_ball"
  captiveye1.kick 0,3

  set cBall2 = captiveye2.createsizedball(20)
  cBall2.image = "black"
  cBall2.frontdecal = "eye_ball_ball"
  captiveye2.kick 45,5

  set cBall3 = captiveye3.createsizedball(20)
  cBall3.image = "black"
  cBall3.frontdecal = "poolballuv8"
  captiveye3.kick 270,7
end Sub


Sub table1_Paused:Controller.Pause = 1:End Sub
Sub table1_unPaused:Controller.Pause = 0:End Sub

'***************
' Random Music
'***************

Dim song

Sub PlaySong 'random
    If bMusicOn Then
        song = "bgout_rat-fink-vpx" &INT(4 * RND(1) ) & ".mp3"
        PlayMusic Song
    End If
End Sub

Sub Table1_MusicDone
    If bMusicOn Then
        PlaySong
    End If
End Sub

Sub ballRelease_Unhit()
    If song="" Then
        PlaySong
    End If
End Sub


'************************************************************

'**********************************************************************************************************
'Key Handling
'**********************************************************************************************************
 Sub Table1_KeyDown(ByVal keycode)
  If vpmKeyDown(KeyCode) Then Exit Sub

    If keycode=PlungerKey Then Plunger.Pullback:playsoundAtVol"plungerpull",plunger,1

  If keycode = LeftTiltKey Then
    LeftNudge 80, 1.2, 20
    PlaySound SoundFX("nudge_left",DOFContactors)
  End If

    If keycode = RightTiltKey Then
    RightNudge 280, 1.2, 20
    PlaySound SoundFX("nudge_right",DOFContactors)
  End If

    If keycode = CenterTiltKey Then
    CenterNudge 0, 1.6, 25
    PlaySound SoundFX("nudge_forward",DOFContactors)
  End If
End Sub

Sub Table1_KeyUp(ByVal keycode)
  If vpmKeyUp(KeyCode) Then Exit Sub
  If keycode=PlungerKey Then Plunger.Fire:playsoundAtVol"plunger",Plunger, 1

End Sub


'**********Sling Shot Animations
' Rstep and Lstep  are the variables that increment the animation
'****************
Dim RStep, Lstep

Sub RightSlingShot_Slingshot
    PlaySoundAtVol SoundFX("right_slingshot",DOFContactors), sling1, 1
    vpmTimer.PulseSw 36
    RightSling.Visible = 0
    RightSling1.Visible = 1
    sling1.TransZ = -20
    RStep = 0
    RightSlingShot.TimerEnabled = 1
    Light41.State = 0:Light42.State = 0:Light43.State = 0:Light40.State = 0
End Sub

Sub RightSlingShot_Timer
    Select Case RStep
        Case 1:Light41.State = 1:Light42.State = 1:Light43.State = 1:Light40.State = 1
        Case 3:RightSLing1.Visible = 0:RightSLing2.Visible = 1:sling1.TransZ = -10
        Case 4:RightSLing2.Visible = 0:RightSLing.Visible = 1:sling1.TransZ = 0:RightSlingShot.TimerEnabled = 0
    End Select
    RStep = RStep + 1
End Sub

Sub LeftSlingShot_Slingshot
    PlaySoundAtVol SoundFX("left_slingshot",DOFContactors), sling2, 1
    vpmTimer.PulseSw 37
    LeftSling.Visible = 0
    LeftSling1.Visible = 1
    sling2.TransZ = -20
    LStep = 0
    LeftSlingShot.TimerEnabled = 1
    Light6.State = 0:Light12.State = 0:Light5.State = 0:Light15.State = 0
End Sub

Sub LeftSlingShot_Timer
    Select Case LStep
        Case 1:Light6.State = 1:Light12.State = 1:Light5.State = 1:Light15.State = 1
        Case 3:LeftSLing1.Visible = 0:LeftSLing2.Visible = 1:sling2.TransZ = -10
        Case 4:LeftSLing2.Visible = 0:LeftSLing.Visible = 1:sling2.TransZ = 0:LeftSlingShot.TimerEnabled = 0
    End Select
    LStep = LStep + 1
End Sub


'**********Rubber Compression
' RUstep, RLstep, and LLstep  are the variables that increment the animation
'****************
Dim RUStep, RLstep, LLstep, LUstep, LDstep

Sub sw23a_Timer
    Select Case LDstep
        Case 1:RubberDrop1.Visible = 0:RubberDrop.Visible = 1:sw23a.TimerEnabled = 0
    End Select
    LDStep = LDStep + 1
End Sub

Sub sw23d_Timer
    Select Case RUstep
        Case 1:RubberRightUpper.Visible = 0:RubberRight.Visible = 1:sw23d.TimerEnabled = 0
    End Select
    RUStep = RUStep + 1
End Sub


Sub RubberRL_Hit
    'Rubber Animation
    RubberRight.Visible = 0
    RubberRightLower.Visible = 1
    RLstep = 0
    RubberRL.TimerEnabled = 1
End Sub


Sub RubberRL_Timer
    Select Case RLstep
        Case 1:RubberRightLower.Visible = 0:RubberRight.Visible = 1:RubberRL.TimerEnabled = 0
    End Select
    RLStep = RLStep + 1
End Sub

Sub RubberLL_Hit
    'Rubber Animation
    RubberLeftLower.Visible = 0
    RubberLeftLower1.Visible = 1
    LLstep = 0
    RubberLL.TimerEnabled = 1
End Sub


Sub RubberLL_Timer
    Select Case LLstep
        Case 1:RubberLeftLower1.Visible = 0:RubberLeftLower.Visible = 1:RubberLL.TimerEnabled = 0
    End Select
    LLStep = LLStep + 1
End Sub


Sub RubberLU_Hit
    'Rubber Animation
    RubberLeftUpper.Visible = 0
    RubberLeftUpper1.Visible = 1
    LUstep = 0
    RubberLU.TimerEnabled = 1
End Sub


Sub RubberLU_Timer
    Select Case LUstep
        Case 1:RubberLeftUpper1.Visible = 0:RubberLeftUpper.Visible = 1:RubberLU.TimerEnabled = 0
    End Select
    LUStep = LUStep + 1
End Sub



'*********
' Switches
'*********

Sub sw23a_Hit
    vpmTimer.PulseSw 23
    'Rubber Animation
    RubberDrop.Visible = 0
    RubberDrop1.Visible = 1
    LDstep = 0
    sw23a.TimerEnabled = 1

End Sub
Sub sw23b_Hit:vpmTimer.PulseSw 23:End Sub
Sub sw23c_Hit:vpmTimer.PulseSw 23:End Sub
Sub sw23d_Hit
    vpmTimer.PulseSw 23
    'Rubber Animation
    RubberRight.Visible = 0
    RubberRightUpper.Visible = 1
    RUstep = 0
    sw23d.TimerEnabled = 1
End Sub

' Bumpers
Sub Bumper1_Hit
  vpmTimer.PulseSw 40
  PlaySoundAtVol SoundFX("fx_bumper1",DOFContactors), Bumper1, VolBump
  direction1=Int(Rnd*2)-1
  if direction1=0 then direction1=1
  cBall1.velx = (10+2*RND(1))*direction1
  cBall1.vely = 2*(RND(1)-RND(1))
End Sub

Sub Bumper2_Hit
  vpmTimer.PulseSw 38
  PlaySoundAtVol SoundFX("fx_bumper2",DOFContactors), Bumper2, VolBump
  direction2=Int(Rnd*2)-1
  if direction2=0 then direction2=1
  cBall2.velx = (10+2*RND(1))*direction2
  cBall2.vely = 2*(RND(1)-RND(1))
End Sub

Sub Bumper3_Hit
  vpmTimer.PulseSw 39
  PlaySoundAtVol SoundFX("fx_bumper3",DOFContactors), Bumper3, VolBump
  direction3=Int(Rnd*2)-1
  if direction2=0 then direction3=1
  cBall3.velx = (10+2*RND(1))*direction3
  cBall3.vely = 2*(RND(1)-RND(1))
End Sub


' Drain holes
Sub Drain_Hit()
    bsTrough.AddBall Me
End Sub

'Saucer
Sub Kicker_Hit:PlaySoundAtVol "metalhit2",Kicker,VolKick:bsSaucer.AddBall 0:End Sub

' Rollovers
Sub sw35_Hit:Controller.Switch(35) = 1:PlaySoundAtVol "sensor",ActiveBall, 1:PlaySoundAtVol "outlane",ActiveBall, 1:End Sub
Sub sw35_UnHit:Controller.Switch(35) = 0:End Sub

Sub sw20_Hit:Controller.Switch(20) = 1:PlaySoundAtVol "sensor",ActiveBall, 1:End Sub
Sub sw20_UnHit:Controller.Switch(20) = 0:End Sub

Sub sw19_Hit:Controller.Switch(19) = 1:PlaySoundAtVol "sensor",ActiveBall, 1:End Sub
Sub sw19_UnHit:Controller.Switch(19) = 0:End Sub

Sub sw34_Hit:Controller.Switch(34) = 1:PlaySoundAtVol "sensor",ActiveBall, 1:PlaySoundAtVol "outlane",ActiveBall, 1:End Sub
Sub sw34_UnHit:Controller.Switch(34) = 0:End Sub

Sub sw5_Hit:Controller.Switch(5) = 1:PlaySoundAtVol "sensor",ActiveBall, 1:End Sub
Sub sw5_UnHit:Controller.Switch(5) = 0:End Sub

Sub sw26_Hit:Controller.Switch(26) = 1:PlaySoundAtVol "sensor",ActiveBall, 1:End Sub
Sub sw26_UnHit:Controller.Switch(26) = 0:End Sub

Sub sw25_Hit:Controller.Switch(25) = 1:PlaySoundAtVol "sensor",ActiveBall, 1:End Sub
Sub sw25_UnHit:Controller.Switch(25) = 0:End Sub

Sub sw31_Hit:Controller.Switch(31) = 1:PlaySoundAtVol "sensor",ActiveBall, 1:End Sub
Sub sw31_UnHit:Controller.Switch(31) = 0:End Sub

Sub sw30_Hit:Controller.Switch(30) = 1:PlaySoundAtVol "sensor",ActiveBall, 1:End Sub
Sub sw30_UnHit:Controller.Switch(30) = 0:End Sub

Sub sw18_Hit:Controller.Switch(18) = 1:PlaySoundAtVol "sensor",ActiveBall, 1:End Sub
Sub sw18_UnHit:Controller.Switch(18) = 0:End Sub

Sub sw22_Hit:Controller.Switch(22) = 1:PlaySoundAtVol "sensor",ActiveBall, 1:End Sub
Sub sw22_UnHit:Controller.Switch(22) = 0:End Sub

' Droptargets
Sub sw1_Hit:dtbank.Hit 1:lightdt1.state = 1:End Sub
Sub sw2_Hit:dtbank.Hit 2:lightdt2.state = 1:End Sub
Sub sw3_Hit:dtbank.Hit 3:lightdt3.state = 1:End Sub
Sub sw4_Hit:dtbank.Hit 4:lightdt4.state = 1:End Sub

' Targets
Sub sw27_Hit:vpmTimer.PulseSw 27:PlaySoundAtVol SoundFX("target",DOFContactors),ActiveBall, VolTarg:End Sub
Sub sw28_Hit:vpmTimer.PulseSw 28:PlaySoundAtVol SoundFX("target",DOFContactors),ActiveBall, VolTarg:End Sub
Sub sw29_Hit:vpmTimer.PulseSw 29:PlaySoundAtVol SoundFX("target",DOFContactors),ActiveBall, VolTarg:End Sub
Sub sw21_Hit:vpmTimer.PulseSw 21:PlaySoundAtVol SoundFX("target",DOFContactors),ActiveBall, VolTarg:End Sub
Sub sw24_Hit:vpmTimer.PulseSw 24:PlaySoundAtVol SoundFX("target",DOFContactors),ActiveBall, VolTarg:End Sub


 ' Extra Sounds

 'Sub Rubbers_Hit(idx):PlaySound "rubber":End Sub
 'Sub Woods_Hit(idx):PlaySound "woodhit":End Sub
 'Sub Rolling_Hit():PlaySound "ballrolling":End Sub


'*********
'Solenoids
'*********

 SolCallback(6) = "vpmSolSound ""knocker"","
 SolCallback(7) = "bsTrough.SolOut"
 SolCallBack(8) = "SolTopHole"
 SolCallback(14) = "SolRaiseDrop"
 SolCallback(19) = "vpmNudge.SolGameOn"

 Sub SolTopHole(Enabled)
     If Enabled Then
         bsSaucer.SolOut 1
         sw32a.IsDropped = 0
         sw32a.TimerEnabled = 1
     End If
 End Sub

 Sub sw32a_Timer:sw32a.IsDropped = 1:Me.TimerEnabled = 0:End Sub


 Sub SolRaiseDrop(enabled)
   If enabled Then
      dttimer.interval=500
      dttimer.enabled=True
   End if
 End Sub


 Sub dttimer_Timer()
   dttimer.enabled=False
   dtBank.DropSol_On
   lightdt1.state = 0
   lightdt2.state = 0
   lightdt3.state = 0
   lightdt4.state = 0
 End Sub


'**************
' Flipper Subs
'**************

 SolCallback(sLRFlipper) = "SolRFlipper"
 SolCallback(sLLFlipper) = "SolLFlipper"

 Sub SolLFlipper(Enabled)
     If Enabled Then
         PlaySoundAtVol SoundFX("fx_Flipperup",DOFContactors),LeftFlipper,VolFlip:LeftFlipper.RotateToEnd
     Else
         PlaySoundAtVol SoundFX("fx_Flipperdown",DOFContactors),LeftFlipper,VolFlip:LeftFlipper.RotateToStart
     End If
  End Sub

 Sub SolRFlipper(Enabled)
     If Enabled Then
         PlaySoundAtVol SoundFX("fx_Flipperup",DOFContactors),RightFlipper,VolFlip:RightFlipper.RotateToEnd
     Else
         PlaySoundAtVol SoundFX("fx_Flipperdown",DOFContactors),RightFlipper,VolFlip:RightFlipper.RotateToStart
     End If
 End Sub

Sub LeftFlipper_Collide(parm)
  RandomSoundFlipper()
End Sub

Sub RightFlipper_Collide(parm)
  RandomSoundFlipper()
End Sub

Sub RandomSoundFlipper()
  Select Case Int(Rnd*3)+1
    Case 1 : PlaySound "flip_hit_1", 0, Vol(ActiveBall)*VolRH, Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0
    Case 2 : PlaySound "flip_hit_2", 0, Vol(ActiveBall)*VolRH, Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0
    Case 3 : PlaySound "flip_hit_3", 0, Vol(ActiveBall)*VolRH, Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0
  End Select
End Sub


'**************
' Edit Dips
'**************

 Sub editDips
     Dim vpmDips:Set vpmDips = New cvpmDips
     With vpmDips
         .AddForm 320, 258, "Star Trek - DIP switch settings"
         .AddChk 214, 218, 95, Array("Match feature", &H00100000)
         .AddChk 214, 234, 95, Array("Credits display", &H00080000)
         .AddFrame 2, 5, 88, "HyperSpace lane", &H00200000, Array("Start at 2K", 0, "Start at 4K", &H200000)
         .AddFrame 2, 55, 88, "BALLY value", &H10000000, Array("Start at 10K", 0, "Start at 25K", &H10000000)
         .AddFrame 2, 105, 88, "BALLY Special", &H00400000, Array("Each time", &H400000, "Only once", 0)
         .AddFrame 2, 155, 88, "Center target", &H00800000, Array("Always lit", &H800000, "Alternating", 0)
         .AddFrame 2, 205, 88, "Outlanes", &H20000000, Array("Both lit", &H20000000, "Alternating", 0)
         .AddFrame 105, 5, 88, "Balls per game", &H0F009F1F, Array("3 balls", &H01000A04, "5 balls", &H01008A04)
         .AddFrame 105, 53, 88, "Play mode", &H00006000, Array("Replay", &H6000, "Extra Ball", &H4000, "Novelty", 0, "Unknown", &H2000)
         .AddFrame 105, 129, 88, "Sound settings", &H80000080, Array("Few", 0, "More", &H80, "Most", &H80000000, "Full", &H80000080)
         .AddFrame 105, 205, 88, "Return lanes", &H40000000, Array("Both lit", &H40000000, "Alternating", 0)
         .AddFrame 208, 5, 93, "High game to date", &H00000060, Array("No award", 0, "1 credit", &H20, "2 credits", &H40, "3 credits", &H60)
         .AddFrame 208, 83, 93, "Max. credits", &H00070000, Array("5 credits", 0, "10 credits", &H10000, "15 credits", &H20000, "20 credits", &H30000, "25 credits", &H40000, "30 credits", &H50000, "35 credits", &H60000, "40 credits", &H70000)
         .ViewDips
     End With
 End Sub
 Set vpmShowDips = GetRef("editDips")


'**************
' GI Lights On
'**************
dim xx

For each xx in GI:xx.State = 1: Next
For each xx in GI_low:xx.State = 1: Next
For each xx in GI_Other:xx.State = 1: Next

'**********************************************************************************************************
'Map lights to an array
'**********************************************************************************************************

Set Lights(1) = l1
Set Lights(2) = l2
Set Lights(3) = l3
Set Lights(5) = l5
Set Lights(6) = l6
Set Lights(7) = l7
Set Lights(8) = l8
Set Lights(9) = l9
Set Lights(10) = l10
Set Lights(11) = l11
Set Lights(12) = l12
Set Lights(13) = l13 'Ball in Play
Set Lights(14) = l14
Set Lights(15) = l15 'Player 1
Set Lights(17) = l17
Set Lights(18) = l18
Set Lights(19) = l19
Lights(20)=Array(l20,l20b)
Set Lights(21) = l21
Set Lights(22) = l22
Set Lights(23) = l23
Set Lights(24) = l24
Set Lights(25) = l25
Set Lights(26) = l26
Set Lights(27) = l27 'Match
Set Lights(28) = l28
Set Lights(29) = l29 'High Score to Date
Set Lights(30) = l30
Set Lights(31) = l31 'Player 2
Set Lights(33) = l33
Set Lights(34) = l34
Set Lights(35) = l35
Lights(36)= Array(l36,l36b)
Set Lights(37) = l37
Set Lights(38) = l38
Set Lights(39) = l39
Set Lights(40) = l40
Set Lights(41) = l41
Set Lights(42) = l42
Set Lights(43) = l43 'SAME PLAYER SHOOTS AGAIN
Set Lights(44) = l44
Set Lights(45) = l45 'GAME OVER
Set Lights(46) = l46
Set Lights(47) = l47 'Player 3
Set Lights(49) = l49
Set Lights(50) = l50
Set Lights(52) = l52
Set Lights(53) = l53
Set Lights(54) = l54
Set Lights(55) = l55
Set Lights(56) = l56
Set Lights(57) = l57
Set Lights(58) = l58
Set Lights(59) = l59 'coin light
Set Lights(60) = l60
Set Lights(61) = l61 'TILT
Set Lights(62) = l62
Set Lights(63) = l63 'Player 4
'Lights(65) = Array(bumperlight1,bumperlight1a)  'Bumper L
'Lights(66) = Array(bumperlight2,bumperlight2a)  'Bumper R
'Lights(67) = Array(bumperlight3,bumperlight3a)  'Bumper C

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


Sub Rubbers_Hidden_Hit(idx)
  dim finalspeed
    finalspeed=SQR(activeball.velx * activeball.velx + activeball.vely * activeball.vely)
  If finalspeed > 20 then
    PlaySound "fx_rubber2", 0, Vol(ActiveBall)*VolRH, Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
  End if
  If finalspeed >= 6 AND finalspeed <= 20 then
    RandomSoundRubber()
  End If
End Sub

Sub Rubbers_Hidden_Wall_Hit(idx)
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
  Select Case Int(Rnd*2)+1
    Case 1 : PlaySound "rubber_hit_1", 0, Vol(ActiveBall)*VolRH, Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
    Case 2 : PlaySound "rubber_hit_2", 0, Vol(ActiveBall)*VolRH, Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
    Case 3 : PlaySound "rubber_hit_3", 0, Vol(ActiveBall)*VolRH, Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
  End Select
End Sub


 '************************************
 '          LEDs Display
 'Based on Scapino's 7 digit Reel LEDs
 '************************************


 LampTimer.Interval = 35
 LampTimer.Enabled = 1

 Sub LampTimer_Timer()
     UpdateLeds
     UpdateTextBoxes
 End Sub


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
 Patterns(7) = 125   '6
 Patterns(8) = 7     '7
 Patterns(9) = 127   '8
 Patterns(10) = 111  '9

 Patterns2(0) = 128  'empty
 Patterns2(1) = 191  '0
 Patterns2(2) = 134  '1
 Patterns2(3) = 219  '2
 Patterns2(4) = 207  '3
 Patterns2(5) = 230  '4
 Patterns2(6) = 237  '5
 Patterns2(7) = 253  '6
 Patterns2(8) = 135  '7
 Patterns2(9) = 255  '8
 Patterns2(10) = 239 '9

 Set Digits(0) = a0
 Set Digits(1) = a1
 Set Digits(2) = a2
 Set Digits(3) = a3
 Set Digits(4) = a4
 Set Digits(5) = a5
 Set Digits(6) = a6

 Set Digits(7) = b0
 Set Digits(8) = b1
 Set Digits(9) = b2
 Set Digits(10) = b3
 Set Digits(11) = b4
 Set Digits(12) = b5
 Set Digits(13) = b6

 Set Digits(14) = c0
 Set Digits(15) = c1
 Set Digits(16) = c2
 Set Digits(17) = c3
 Set Digits(18) = c4
 Set Digits(19) = c5
 Set Digits(20) = c6

 Set Digits(21) = d0
 Set Digits(22) = d1
 Set Digits(23) = d2
 Set Digits(24) = d3
 Set Digits(25) = d4
 Set Digits(26) = d5
 Set Digits(27) = d6

 Set Digits(28) = e0
 Set Digits(29) = e1
 Set Digits(30) = e2
 Set Digits(31) = e3

 Sub UpdateLeds
     On Error Resume Next
     Dim ChgLED, ii, jj, chg, stat
     ChgLED = Controller.ChangedLEDs(&HFF, &HFFFF)
     If Not IsEmpty(ChgLED)Then
         For ii = 0 To UBound(ChgLED)
             chg = chgLED(ii, 1):stat = chgLED(ii, 2)

             For jj = 0 to 10
                 If stat = Patterns(jj)OR stat = Patterns2(jj)then Digits(chgLED(ii, 0)).SetValue jj
             Next
         Next
     End IF
 End Sub


 Sub UpdateTextBoxes()
     NFadeT 13, l13, "BALL IN PLAY"
     NFadeT 27, l27, "MATCH"
     NFadeT 29, l29, "HIGH SCORE TO DATE"
     NFadeT 43, l43, "SAME PLAYER SHOOTS AGAIN"
     NFadeT 45, l45, "GAME OVER"
     NFadeT 61, l61, "TILT"
 End Sub

 Sub NFadeT(nr, a, b)
     Select Case controller.lamp(nr)
         Case False:a.Text = ""
         Case True:a.Text = b
     End Select
End Sub

CREDIT.Text = "CREDIT"

If Table1.ShowDT = false then
    dim zz
    For each zz in DT:zz.Visible = false: Next
End If

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

Function Pan(ball) ' Calculates the pan for a ball based on the X position on the table. "table1" is the name of the table
    Dim tmp
    tmp = ball.x * 2 / table1.width-1
    If tmp > 0 Then
        Pan = Csng(tmp ^10)
    Else
        Pan = Csng(-((- tmp) ^10) )
    End If
End Function

Function AudioFade(ball) ' Can this be together with the above function ?
  Dim tmp
  tmp = ball.y * 2 / Table1.height-1
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
Sub Table1_exit()
  Controller.Pause = False
  Controller.Stop
End Sub


Option Explicit
Randomize

' Thalamus 2018-07-23
' Added/Updated "Positional Sound Playback Functions" and "Supporting Ball & Sound Functions"
' Changed UseSolenoids=1 to 2
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

Const cGameName="httip_l1",UseSolenoids=2,UseLamps=1,UseGI=0,SSolenoidOn="SolOn",SSolenoidOff="SolOff",SCoin="coin"

LoadVPM "00990100", "S4.VBS", 1.2
Dim DesktopMode: DesktopMode = Table1.ShowDT

'Solenoid Call backs
'**********************************************************************************************************
SolCallback(1)    = "bsTop.SolOut"
SolCallback(2)    = "dtRReset" 'right drop targets
SolCallback(3)    = "dtLReset" 'left drop targets
SolCallback(4)    = "bsLeft.SolOut"
SolCallback(5)    = "bsTrough.SolOut"
SolCallback(9)      = "vpmSolSound SoundFX(""chime1"",DOFChimes),"
SolCallback(10)     = "vpmSolSound SoundFX(""chime2"",DOFChimes),"
SolCallback(11)     = "vpmSolSound SoundFX(""chime3"",DOFChimes),"
SolCallback(12)     = "vpmSolSound SoundFX(""chime4"",DOFChimes),"
SolCallback(13)   = "vpmSolSound ""solon"","   'Noise Drum
'SolCallback(14)   = "vpmSolSound ""Knocker""," 'coin in knocker
SolCallback(15)   = "vpmSolSound SoundFX(""solon"",DOFBell),"   'Buzzer


SolCallback(sLRFlipper) = "SolRFlipper"
SolCallback(sLLFlipper) = "SolLFlipper"

Sub SolLFlipper(Enabled)
     If Enabled Then
         PlaySoundAtVol SoundFx("fx_Flipperup",DOFContactors), LeftFlipper, VolFlip:LeftFlipper.RotateToEnd
     Else
         PlaySoundAtVol SoundFx("fx_Flipperdown",DOFContactors), LeftFlipper, VolFlip:LeftFlipper.RotateToStart
     End If
  End Sub

Sub SolRFlipper(Enabled)
     If Enabled Then
         PlaySoundAtVol SoundFx("fx_Flipperup",DOFContactors), RightFlipper, VolFlip:RightFlipper.RotateToEnd
     Else
         PlaySoundAtVol SoundFx("fx_Flipperdown",DOFContactors), RightFlipper, VolFlip:RightFlipper.RotateToStart
     End If
End Sub

'Fantasy
 SolCallback(17)  = "vpmFlasher Array(Flasher17,Flasher17a),"
'**********************************************************************************************************
 Sub dtLReset(Enabled)
  If Enabled Then
    dtL.DropSol_On
    Controller.Switch(37)=0
    LCount=0
  End If
End Sub

Sub dtRReset(Enabled)
  If Enabled Then
    dtR.DropSol_On
    Controller.Switch(37)=0
    RCount=0
  End If
End Sub


'Initiate Table
'**********************************************************************************************************

Dim bsTrough, bsLeft, bsTop, dtL, dtR, Count, LCount, RCount
Count=0

Sub Table1_Init
  vpmInit Me
  On Error Resume Next
    With Controller
    .GameName = cGameName
    If Err Then MsgBox "Can't start Game" & cGameName & vbNewLine & Err.Description : Exit Sub
    .SplashInfoLine = "Hot Tip (Williams)"&chr(13)&"You Suck"
    .HandleMechanics=0
    .HandleKeyboard=0
    .ShowDMDOnly=1
    .ShowFrame=0
    .ShowTitle=0
        .hidden = 1
    If Err Then MsgBox Err.Description
  End With
  On Error Goto 0
    Controller.SolMask(0)=0
      vpmTimer.AddTimer 2000,"Controller.SolMask(0)=&Hffffffff'" 'ignore all solenoids - then add the timer to renable all the solenoids after 2 seconds
    Controller.Run
  If Err Then MsgBox Err.Description
  On Error Goto 0

  PinMAMETimer.Interval = PinMAMEInterval
  PinMAMETimer.Enabled = True

  vpmNudge.TiltSwitch = swTilt
  vpmNudge.Sensitivity = 3
  vpmNudge.TiltObj = Array(LeftSlingshot, RightSlingshot, Bumper1)

  Set bsTrough = New cvpmBallStack ' Trough handler
    bsTrough.InitSw 0,9,0,0,0,0,0,0
    bsTrough.InitKick BallRelease, 180, 1
  bsTrough.InitExitSnd SoundFX("BallRelease",DOFContactors),SoundFX("solenoid",DOFContactors)
    bsTrough.Balls = 1

  ' Left Kicker
  Set bsLeft = New cvpmBallStack
  bsLeft.InitSaucer sw29,29,110,10
  bsLeft.KickAngleVar = 15
  bsLeft.InitExitSnd SoundFX("Popper",DOFContactors),SoundFX("solon",DOFContactors)

  ' Top Kicker
  Set bsTop = New cvpmBallStack
  bsTop.InitSaucer sw22,22,170,7
  bsTop.KickAngleVar = 20
  bsTop.InitExitSnd SoundFX("Popper",DOFContactors),SoundFX("solon",DOFContactors)

  Set dtR = New cvpmDropTarget
  dtR.InitDrop Array(sw17,sw18,sw19),Array(17,18,19)
  dtR.InitSnd SoundFX("DTDrop",DOFContactors),SoundFX("DTReset",DOFContactors)

  Set dtL = New cvpmDropTarget
  dtL.InitDrop Array(sw25,sw26,sw27),Array(25,26,27)
  dtL.InitSnd SoundFX("DTDrop",DOFContactors),SoundFX("DTReset",DOFContactors)

  CaptiveBall.CreateBall
  CaptiveBall.Kick 270,1
End Sub

'**********************************************************************************************************
'Plunger code
'**********************************************************************************************************

Sub Table1_KeyDown(ByVal KeyCode)
  If KeyDownHandler(keycode) Then Exit Sub
  If keycode = PlungerKey Then Plunger.Pullback:playsoundAtVol"plungerpull", Plunger, 1
End Sub

Sub Table1_KeyUp(ByVal KeyCode)
  If KeyUpHandler(keycode) Then Exit Sub
  If keycode = PlungerKey Then Plunger.Fire:PlaySoundAtVol"plunger", Plunger, 1
End Sub

'**********************************************************************************************************

'SWITCH HANDLING
'**********************************************************************************************************
 ' Drain hole and kickers
Sub Drain_Hit:playsoundAtVol"drain", drain, 1:bsTrough.addball me:End Sub
Sub sw22_Hit : bsTop.addBall 0 : PlaySoundAtVol "Wave6", sw22,1: End Sub
Sub sw29_Hit : bsLeft.addBall 0 : PlaySoundAtVol "Wave5", sw29,1:End Sub

'scoring rubbers
Sub sw21_Slingshot:vpmTimer.PulseSw 21:playsound SoundFX("slingshot",DOFContactors):End Sub ' TODO
Sub sw28_Slingshot:vpmTimer.PulseSw 28:playsound SoundFX("slingshot",DOFContactors):End Sub
Sub sw23_Slingshot:vpmTimer.PulseSw 23:playsound SoundFX("slingshot",DOFContactors):End Sub

'Stand Up Targets
Sub Trigger16_Hit : vpmTimer.PulseSwitch 16, 100, 0 : End Sub
Sub Trigger20_Hit : vpmTimer.PulseSwitch 20, 100, 0 : End Sub

'Wire Triggers
Sub sw11_Hit : Controller.Switch(11) = 1 : playsoundAtVol"rollover", Activeball, 1 : End Sub
Sub sw11_UnHit : Controller.Switch(11) = 0 :End Sub
Sub sw12_Hit : Controller.Switch(12) = 1 : playsoundAtVol"rollover", ActiveBall, 1 : End Sub
Sub sw12_UnHit : Controller.Switch(12) = 0 :End Sub
Sub sw15_Hit : Controller.Switch(15) = 1 : playsoundAtVol"rollover", ActiveBall, 1 : End Sub
Sub sw15_UnHit : Controller.Switch(15) = 0 : End Sub
Sub sw24_Hit : Controller.Switch(24) = 1 : playsoundAtVol"rollover", ActiveBall, 1 : End Sub
Sub sw24_UnHit : Controller.Switch(24) = 0 : End Sub
Sub sw33_Hit : Controller.Switch(33) = 1 : playsoundAtVol"rollover", ActiveBall, 1 : End Sub
Sub sw33_UnHit : Controller.Switch(33) = 0 : End Sub
Sub sw34_Hit : Controller.Switch(34) = 1 : playsoundAtVol"rollover", ActiveBall, 1 : End Sub
Sub sw34_UnHit : Controller.Switch(34) = 0 :End Sub

'Bumpers
Sub Bumper1_Hit : vpmTimer.PulseSw(31) : playsoundAtVol SoundFX("fx_bumper1",DOFContactors), Bumper1, VolBump: End Sub

'Spinners
Sub Spinner_Spin:vpmTimer.PulseSw (30) : playsoundAtVol"fx_spinner", Spinner, VolSpin : End Sub

'Drop Targets
Sub sw17_Hit : dtR.Hit 1 :CheckRCount: End Sub
Sub sw18_Hit : dtR.Hit 2 :CheckRCount: End Sub
Sub sw19_Hit : dtR.Hit 3 :CheckRCount: End Sub

Sub sw25_Hit : dtL.Hit 1 :CheckLCount: End Sub
Sub sw26_Hit : dtL.Hit 2 :CheckLCount: End Sub
Sub sw27_Hit : dtL.Hit 3 :CheckLCount: End Sub

'Switch 37=1 when  ALL drop targets down
 Sub CheckLCount
  LCount=LCount+1
  If 6 = LCount + RCount Then Controller.Switch(37)=1
 End Sub

 Sub CheckRCount
  RCount=RCount+1
  If 6 = LCount + RCount Then Controller.Switch(37)=1
 End Sub


'CaptiveBall

'**********************************************************************************************************
'**********************************************************************************************************

' Map lights into array
 '**********************************************************************************************************
Set Lights(11) = Light11
Set Lights(12) = Light12
Set Lights(13) = Light13
Set Lights(14) = Light14
Set Lights(15) = Light15
Set Lights(16) = Light16
Set Lights(17) = Light17
Set Lights(18) = Light18
Set Lights(19) = Light19
Set Lights(20) = Light20
Set Lights(21) = Light21
Set Lights(22) = Light22
Set Lights(23) = Light23
Set Lights(24) = Light24
Set Lights(25) = Light25
Set Lights(26) = Light26
Set Lights(27) = Light27
Set Lights(28) = Light28
Set Lights(29) = Light29
Set Lights(39) = Light39
Set Lights(40) = Light40
Set Lights(41) = Light41
Set Lights(42) = Light42
Set Lights(43) = Light43
Set Lights(44) = Light44
Set Lights(45) = Light45
Set Lights(46) = Light46
Set Lights(47) = Light47
Set Lights(48) = Light48
Set Lights(49) = Light49
Set Lights(56) = Light56

'Backglass
'Set Lights(50) = '1 Can Play
'Set Lights(51) = '2 Can Play
'Set Lights(52) = '3 Can Play
'Set Lights(53) = '4 Can Play
'Set Lights(54) = ' Match
'Set Lights(55) = ' Ball In Play
'Set Lights(57) ='1 Up
'Set Lights(58) ='2 Up
'Set Lights(59) ='3 Up
'Set Lights(60) ='4 Up
'Set Lights(61) = ' Tilt
'Set Lights(62) = ' Game Over
'Set Lights(63) = ' Same Player Shoots Again
'Set Lights(64) = '  High Score

   Sub DisplayTimer_Timer
  Dim ChgLED,ii,jj,num,chg,stat,obj,b,x
  ChgLED=Controller.ChangedLEDs(&Hffffffff,&Hffffffff)
  If Not IsEmpty(ChgLED) Then
    If DesktopMode = True Then
    For ii=0 To UBound(chgLED)
      num=chgLED(ii,0):chg=chgLED(ii,1):stat=chgLED(ii,2)
      For Each obj In Digits(num)
        If chg And 1 Then obj.State=stat And 1
        chg=chg\2:stat=stat\2
      Next
    Next
    end if
  End If
End Sub

Dim Digits(27)

Digits(0)=Array(D1,D2,D3,D4,D5,D6,D7)
Digits(1)=Array(D11,D12,D13,D14,D15,D16,D17)
Digits(2)=Array(D21,D22,D23,D24,D25,D26,D27)
Digits(3)=Array(D31,D32,D33,D34,D35,D36,D37)
Digits(4)=Array(D41,D42,D43,D44,D45,D46,D47)
Digits(5)=Array(D51,D52,D53,D54,D55,D56,D57)
Digits(6)=Array(D61,D62,D63,D64,D65,D66,D67)
Digits(7)=Array(D71,D72,D73,D74,D75,D76,D77)
Digits(8)=Array(D81,D82,D83,D84,D85,D86,D87)
Digits(9)=Array(D91,D92,D93,D94,D95,D96,D97)
Digits(10)=Array(D101,D102,D103,D104,D105,D106,D107)
Digits(11)=Array(D111,D112,D113,D114,D115,D116,D117)
Digits(12)=Array(D121,D122,D123,D124,D125,D126,D127)
Digits(13)=Array(D131,D132,D133,D134,D135,D136,D137)
Digits(14)=Array(D141,D142,D143,D144,D145,D146,D147)
Digits(15)=Array(D151,D152,D153,D154,D155,D156,D157)
Digits(16)=Array(D161,D162,D163,D164,D165,D166,D167)
Digits(17)=Array(D171,D172,D173,D174,D175,D176,D177)
Digits(18)=Array(D181,D182,D183,D184,D185,D186,D187)
Digits(19)=Array(D191,D192,D193,D194,D195,D196,D197)
Digits(20)=Array(D201,D202,D203,D204,D205,D206,D207)
Digits(21)=Array(D211,D212,D213,D214,D215,D216,D217)
Digits(22)=Array(D221,D222,D223,D224,D225,D226,D227)
Digits(23)=Array(D231,D232,D233,D234,D235,D236,D237)
Digits(24)=Array(D241,D242,D243,D244,D245,D246,D247)
Digits(25)=Array(D251,D252,D253,D254,D255,D256,D257)
Digits(26)=Array(D261,D262,D263,D264,D265,D266,D267)
Digits(27)=Array(D271,D272,D273,D274,D275,D276,D277)



'*****GI Lights On
dim xx
For each xx in GI:xx.State = 1: Next

'**********Sling Shot Animations
' Rstep and Lstep  are the variables that increment the animation
'****************
Dim RStep, Lstep, R2Step, L2step

Sub RightSlingShot_Slingshot
  vpmTimer.PulseSw 35
    PlaySoundAtVol SoundFX("left_slingshot",DOFContactors), sling1, 1
    RSling.Visible = 0
    RSling1.Visible = 1
    sling1.TransZ = -20
    RStep = 0
    RightSlingShot.TimerEnabled = 1
End Sub

Sub RightSlingShot_Timer
    Select Case RStep
        Case 3:RSLing1.Visible = 0:RSLing2.Visible = 1:sling1.TransZ = -10
        Case 4:RSLing2.Visible = 0:RSLing.Visible = 1:sling1.TransZ = 0:RightSlingShot.TimerEnabled = 0:
    End Select
    RStep = RStep + 1
End Sub

Sub LeftSlingShot_Slingshot
  vpmTimer.PulseSw 13
    PlaySoundAtVol SoundFX("right_slingshot",DOFContactors), sling2, 1
    LSling.Visible = 0
    LSling1.Visible = 1
    sling2.TransZ = -20
    LStep = 0
    LeftSlingShot.TimerEnabled = 1
End Sub

Sub LeftSlingShot_Timer
    Select Case LStep
        Case 3:LSLing1.Visible = 0:LSLing2.Visible = 1:sling2.TransZ = -10
        Case 4:LSLing2.Visible = 0:LSLing.Visible = 1:sling2.TransZ = 0:LeftSlingShot.TimerEnabled = 0:
    End Select
    LStep = LStep + 1
End Sub

Sub Right2SlingShot_Slingshot
  vpmTimer.PulseSw 14
    PlaySoundAtVol SoundFX("left_slingshot",DOFContactors), sling3, 1
    R2Sling.Visible = 0
    R2Sling1.Visible = 1
    sling3.TransZ = -20
    R2Step = 0
    Right2SlingShot.TimerEnabled = 1
End Sub

Sub Right2SlingShot_Timer
    Select Case R2Step
        Case 3:R2SLing1.Visible = 0:R2SLing2.Visible = 1:sling3.TransZ = -10
        Case 4:R2SLing2.Visible = 0:R2SLing.Visible = 1:sling3.TransZ = 0:Right2SlingShot.TimerEnabled = 0:
    End Select
    R2Step = R2Step + 1
End Sub

Sub Left2SlingShot_Slingshot
  vpmTimer.PulseSw 32
    PlaySoundAtVol SoundFX("right_slingshot",DOFContactors), sling4, 1
    L2Sling.Visible = 0
    L2Sling1.Visible = 1
    sling4.TransZ = -20
    L2Step = 0
    Left2SlingShot.TimerEnabled = 1
End Sub

Sub Left2SlingShot_Timer
    Select Case L2Step
        Case 3:L2SLing1.Visible = 0:L2SLing2.Visible = 1:sling4.TransZ = -10
        Case 4:L2SLing2.Visible = 0:L2SLing.Visible = 1:sling4.TransZ = 0:Left2SlingShot.TimerEnabled = 0:
    End Select
    L2Step = L2Step + 1
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
' will call this routine. What you add in the sub is up to you. As an example is a simple PlaysoundAtVol with volume and paning
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
    Case 1 : PlaySound "flip_hit_1", 0, Vol(ActiveBall)*VolRh, Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
    Case 2 : PlaySound "flip_hit_2", 0, Vol(ActiveBall)*VolRh, Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
    Case 3 : PlaySound "flip_hit_3", 0, Vol(ActiveBall)*VolRh, Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
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


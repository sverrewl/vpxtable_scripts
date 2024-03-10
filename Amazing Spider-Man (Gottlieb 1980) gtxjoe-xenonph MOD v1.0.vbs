'VPX Amazing Spider-Man
' gtxjoe v1.0
'
'  Thanks to all VP contributors present and past
'  Inspiration from Bob5453 and Gaston for their previous Spider-man versions

' Thalamus 2018-07-18
' Added/Updated "Positional Sound Playback Functions" and "Supporting Ball & Sound Functions"
' Changed UseSolenoids=1 to 2
' No special SSF tweaks yet.

Option Explicit
Randomize

Const cGameName     = "spiderm7"

On Error Resume Next
ExecuteGlobal GetTextFile("controller.vbs")
If Err Then MsgBox "Can't open controller.vbs"
On Error Goto 0

LoadVPM "01000100", "sys80.vbs", 2.31

Const UseSolenoids  = 2
Const UseLamps      = False
Const UseGI         = False

' Standard Sounds
Const SSolenoidOn  = ""
Const SSolenoidOff = ""
Const SCoin        = "fx_coin"

Dim bsTrough, initContacts, dtRDrop, dtLDrop, bsLSaucer, bsMSaucer, bsRSaucer

Sub Table1_Init()

  GiLights 0

' Thalamus : Was missing 'vpminit me'
  vpminit me

  On Error Resume Next

  With Controller
    .GameName               = cGameName
    .SplashInfoLine         = "Amazing Spider-Man (Gottlieb 1980) gtxjoe-Xenonph MOD v1.0"
    .HandleKeyboard         = False
    .ShowTitle              = False
    .ShowDMDOnly            = True
    .ShowFrame              = False
    .ShowTitle              = False
    If Table1.ShowDT = False Then
      .Hidden = 1
    Else
      .Hidden = 0
    End If
    .Games(cGameName).Settings.Value("rol") = 0

    On Error Resume Next
    .Run GetPlayerHWnd
    If Err Then MsgBox Err.Description
    End With
    On Error Goto 0

    ' basic pinmame timer
    PinMAMETimer.Interval = PinMAMEInterval
    PinMAMETimer.Enabled  = True

    ' Main Timer init
    PinMAMETimer.Interval = PinMAMEInterval
    PinMAMETimer.Enabled = True


    ' Nudging
    vpmNudge.TiltSwitch = 57
    vpmNudge.Sensitivity = 7
    vpmNudge.TiltObj = Array(Bumper1,Bumper2,LeftSlingshot,RightSlingShot)


    ' ball stack for trough: Gottlieb System 80
    Set bsTrough = New cvpmBallStack
    bsTrough.InitSw 0, 67, 0, 0, 0, 0, 0, 0
    bsTrough.InitKick BallRelease, 85,4
    bsTrough.InitExitSnd Soundfx("BallRelease",DOFContactors), Soundfx("SolOn",DOFContactors)
    bsTrough.Balls = 1

    Set dtLDrop = New cvpmDropTarget
    dtLDrop.InitDrop Array(sw1,sw11,sw21),Array(1,11,21)
    dtLDrop.InitSnd "DropTarget", Soundfx("DropTargetreset",DOFContactors)

    Set dtRDrop = New cvpmDropTarget
    dtRDrop.InitDrop Array(sw0,sw10,sw20,sw30,sw40),Array(0,10,20,30,40)
    dtRDrop.InitSnd "DropTarget",Soundfx("DropTargetreset",DOFContactors)

End Sub


Sub Table1_KeyDown(ByVal keycode)

  If vpmKeyDown(keycode) Then Exit Sub
    If keycode = PlungerKey Then
      Plunger.PullBack
      PlaySound "plungerpull",0,1,0.25,0.25
    End If

    ' If keycode = LeftFlipperKey Then
    '   LeftFlipper1.RotateToEnd
    '   LeftFlipper.RotateToEnd
    '   PlaySound "fx_flipperup", 0, .67, -0.05, 0.05
    ' End If
    '
    ' If keycode = RightFlipperKey Then
    '   RightFlipper.RotateToEnd
    '   RightFlipper1.RotateToEnd
    '   PlaySound "fx_flipperup", 0, .67, 0.05, 0.05
    ' End If

    If keycode = LeftTiltKey Then
      Nudge 90, 2
    End If

    If keycode = RightTiltKey Then
      Nudge 270, 2
    End If

    If keycode = CenterTiltKey Then
      Nudge 0, 2
    End If

End Sub

Sub Table1_KeyUp(ByVal keycode)

  If keycode = PlungerKey Then
    Plunger.Fire
    PlaySound "plunger",0,1,0.25,0.25
  End If

  ' If keycode = LeftFlipperKey Then
  '   LeftFlipper.RotateToStart
  '   LeftFlipper1.RotateToStart
  '   PlaySound "fx_flipperdown", 0, 1, -0.05, 0.05
  ' End If
  '
  ' If keycode = RightFlipperKey Then
  '   RightFlipper.RotateToStart
  '   RightFlipper1.RotateToStart
  '   PlaySound "fx_flipperdown", 0, 1, 0.05, 0.05
  ' End If

  If vpmKeyUp(keycode) Then Exit Sub

End Sub

Sub Drain_Hit()
  PlaySoundAt "drain", Drain
  bsTrough.AddBall Me
End Sub

Sub GILights (enabled)
  Dim light
  For each light in GI:light.State = Enabled: Next
End Sub

Sub Trigger1_hit
  GILights 1
End Sub
'*****************************************
'Solenoids
'*****************************************
SolCallBack(1)  = "SolSaucerMid"
SolCallBack(2)  = "SolSaucerOuter"
SolCallback(5)  = "dtLDrop.SolDropUp"
SolCallback(6)  = "dtRDrop.SolDropUp"
SolCallback(8)  = "SolKnocker"
SolCallback(9)  = "SolTrough"

SolCallback(sLRFlipper) = "SolRFlipper"
SolCallback(sLLFlipper) = "SolLFlipper"

Sub SolLFlipper(Enabled)
  If Enabled Then
    PlaySound SoundFX("fx_flipperup",DOFContactors), 0, 1, -0.1, 0.25
    LeftFlipper.RotateToEnd
    LeftFlipper1.RotateToEnd
  Else
    PlaySound SoundFX("fx_flipperdown",DOFContactors), 0, 1, -0.1, 0.25
    LeftFlipper.RotateToStart
    LeftFlipper1.RotateToStart
  End If
End Sub

Sub SolRFlipper(Enabled)
  If Enabled Then
    PlaySound SoundFX("fx_flipperup",DOFContactors), 0, 1, 0.1, 0.25
    RightFlipper.RotateToEnd
    RightFlipper1.RotateToEnd
  Else
    PlaySound SoundFX("fx_flipperdown",DOFContactors), 0, 1, 0.1, 0.25
    RightFlipper.RotateToStart
    RightFlipper1.RotateToStart
  End If
End Sub

Sub SolTrough(Enabled)
  If Enabled Then bsTrough.ExitSol_On
End Sub

Sub SolKnocker(Enabled)
  If Enabled Then PlaySound SoundFX("Knocker",DOFKnocker)
End Sub

Sub SolSaucerMid(Enabled)
  If Enabled Then sw12.kick 165,6:Controller.Switch(12) = False:PlaySound SoundFX("popper_ball",DOFContactors)

End Sub
Sub SolSaucerOuter(Enabled)
  If Enabled Then Sw2.kick 180,6:Controller.Switch(2) = False: PlaySound SoundFX("popper_ball",DOFContactors)
  If Enabled Then sw22.kick 160,6:Controller.Switch(22) = False: PlaySound SoundFX("popper_ball",DOFContactors)
End Sub
'*****************************************
'Switches
'*****************************************

Sub sw0_Hit  : dtRDrop.Hit 1 : End Sub
Sub sw10_Hit : dtRDrop.Hit 2 : End Sub
Sub sw20_Hit : dtRDrop.Hit 3 : End Sub
Sub sw30_Hit : dtRDrop.Hit 4 : End Sub
Sub sw40_Hit : dtRDrop.Hit 5 : End Sub
Sub sw1_Hit  : dtLDrop.Hit 1 : End Sub
Sub sw11_Hit : dtLDrop.Hit 2 : End Sub
Sub sw21_Hit : dtLDrop.Hit 3 : End Sub

Sub sw31_Hit  : debug.print "31":vpmTimer.PulseSw 31: End Sub
Sub sw31a_Hit : vpmTimer.PulseSw 31: End Sub

Sub sw32a_Hit : vpmTimer.PulseSw 32: End Sub
Sub sw32b_Hit : vpmTimer.PulseSw 32: End Sub
Sub sw32c_Hit : vpmTimer.PulseSw 32: End Sub
Sub sw32d_Hit : vpmTimer.PulseSw 32: End Sub

Sub sw41_Hit :  debug.print "41": vpmTimer.PulseSw 41: End Sub
Sub sw41a_Hit : vpmTimer.PulseSw 41: End Sub

Sub sw42a_Hit : vpmTimer.PulseSw 42: End Sub
Sub sw42b_Hit : vpmTimer.PulseSw 42: End Sub
Sub sw42c_Hit : vpmTimer.PulseSw 42: End Sub
Sub sw42d_Hit : vpmTimer.PulseSw 42: End Sub
Sub sw42e_Hit : vpmTimer.PulseSw 42: End Sub
Sub sw42f_Hit : vpmTimer.PulseSw 42: End Sub

Sub sw50_Hit : vpmTimer.PulseSw 50: End Sub
Sub sw51_Hit : vpmTimer.PulseSw 51: End Sub
Sub sw51a_Hit : vpmTimer.PulseSw 51: End Sub
Sub sw61_Hit : vpmTimer.PulseSw 61: End Sub
Sub sw61a_Hit : vpmTimer.PulseSw 61: End Sub
Sub sw70_Hit : vpmTimer.PulseSw 70: PlaySoundAt "target",sw70: End Sub
Sub sw71_Hit : vpmTimer.PulseSw 71: PlaySoundAt "target",sw71: End Sub

Sub sw60_spin
  vpmTimer.PulseSw 60
  PlaySoundAt "spinner", sw60
End Sub

Sub sw2_Hit
  Controller.Switch(2) = True
  PlaySoundAt "kicker_enter", sw2
End Sub

Sub sw12_Hit
  Controller.Switch(12) = True
  PlaySoundAt "kicker_enter", sw12
End Sub

Sub sw22_Hit
  Controller.Switch(22) = True
  PlaySoundAt "kicker_enter", sw22
End Sub

Dim flipcoin
Sub saucercheck1_hit
  'debug.print activeball.vely
  if activeball.vely < -6 then
    flipcoin = Int((Rnd) + .5)
    sw2.enabled = flipcoin
    DEBUG.PRINT FLIPCOIN
  End If
End Sub
Sub saucercheck1_unhit
  sw2.enabled = true
DEBUG.PRINT "UNHIT"
End Sub
Sub saucercheck2_hit
  if activeball.vely < -6 then
    flipcoin = Int((Rnd) + .3)
    sw12.enabled = flipcoin
    DEBUG.PRINT FLIPCOIN
  End If
End Sub
Sub saucercheck2_unhit
  sw12.enabled = true
DEBUG.PRINT "UNHIT"
End Sub
Sub saucercheck3_hit
  'debug.print activeball.vely
  if activeball.vely < -6 then
    flipcoin = Int((Rnd) + .5)
    sw22.enabled = flipcoin
    DEBUG.PRINT FLIPCOIN
  End If
End Sub
Sub saucercheck3_unhit
  sw22.enabled = true
DEBUG.PRINT "UNHIT"
End Sub
Sub Bumper1_Hit
  vpmTimer.PulseSw 72
  PlaySound SoundFX("fx_bumper1",DOFContactors)
  DOF 104,2
End Sub

Sub Bumper2_Hit
  vpmTimer.PulseSw 72
  PlaySound SoundFX("fx_bumper2",DOFContactors)
  DOF 103,2
End Sub

'**********
' Sling Shot Animations
' Rstep and Lstep  are the variables that increment the animation
'****************
Dim RStep, Lstep

Sub RightSlingShot_Slingshot
  vpmTimer.PulseSw 42
  DOF 102,2
    PlaySound SoundFX("right_slingshot",DOFContactors), 0, 1, 0.05, 0.05
    RSling.Visible = 0
    RSling1.Visible = 1
    sling1.TransZ = -20
    RStep = 0
    RightSlingShot.TimerEnabled = 1
End Sub

Sub RightSlingShot_Timer
    Select Case RStep
        Case 2:RSLing1.Visible = 0:RSLing2.Visible = 1:sling1.TransZ = -10
        Case 3:RSLing2.Visible = 0:RSLing.Visible = 1:sling1.TransZ = 0:RightSlingShot.TimerEnabled = 0
    End Select
    RStep = RStep + 1
End Sub

Sub LeftSlingShot_Slingshot
    vpmTimer.PulseSw 42
  DOF 101,2
    PlaySound SoundFX("left_slingshot",DOFContactors),0,1,-0.05,0.05
    LSling.Visible = 0
    LSling1.Visible = 1
    sling2.TransZ = -20
    LStep = 0
    LeftSlingShot.TimerEnabled = 1
End Sub

Sub LeftSlingShot_Timer
    Select Case LStep
        Case 2:LSLing1.Visible = 0:LSLing2.Visible = 1:sling2.TransZ = -10
        Case 3:LSLing2.Visible = 0:LSLing.Visible = 1:sling2.TransZ = 0:LeftSlingShot.TimerEnabled = 0
    End Select
    LStep = LStep + 1
End Sub

Sub FlipperTimer_Timer()

End Sub

Sub UpdateLeftFlipperLogo()
  LFLogo.RotY = LeftFlipper.CurrentAngle
  LFLogo1.RotY = LeftFlipper1.CurrentAngle
End Sub
Sub UpdateRightFlipperLogo()
  RFLogo.RotY = RightFlipper.CurrentAngle
  RFLogo1.RotY = RightFlipper1.CurrentAngle
End Sub




'************************************
' What you need to add to your table
'************************************

' a timer called CollisionTimer. With a fast interval, like from 1 to 10
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

' The Double For loop: This is a double cycle used to check the collision between a ball and the other ones.
' If you look at the parameters of both cycles, you’ll notice they are designed to avoid checking
' collision twice. For example, I will never check collision between ball 2 and ball 1,
' because I already checked collision between ball 1 and 2. So, if we have 4 balls,
' the collision checks will be: ball 1 with 2, 1 with 3, 1 with 4, 2 with 3, 2 with 4 and 3 with 4.

' Sum first the radius of both balls, and if the height between them is higher then do not calculate anything more,
' since the balls are on different heights so they can't collide.

' The next 3 lines calculates distance between xth and yth balls with the Pytagorean theorem,

' The first "If": Checking if the distance between the two balls is less than the sum of the radius of both balls,
' and both balls are not already colliding.

' Why are we checking if balls are already in collision?
' Because we do not want the sound repeting when two balls are resting closed to each other.

' Set the collision property of both balls to True, and we assign the number of the ball colliding

' Play the collide sound of your choice using the VOL, PAN and PITCH functions

' Last lines: If the distance between 2 balls is more than the radius of a ball,
' then there is no collision and then set the collision property of the ball to False (-1).

Sub Rollovers_Hit (idx)
  PlaySound "rollover", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0, AudioFade(ActiveBall)
  'PlaySound "rollover"
End Sub

Sub Pins_Hit (idx)
  PlaySound "pinhit_low", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0, AudioFade(ActiveBall)
  'PlaySound "pinhit_low"
End Sub

Sub Targets_Hit (idx)
  PlaySound "target", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0, AudioFade(ActiveBall)
  'PlaySound "target"
End Sub

Sub Metals_Thin_Hit (idx)
  PlaySound "metalhit_thin", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
  'PlaySound "metalhit_thin"
End Sub

Sub Metals_Medium_Hit (idx)
  PlaySound "metalhit_medium", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
  'PlaySound "metalhit_medium"
End Sub

Sub Metals2_Hit (idx)
  PlaySound "metalhit2", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
End Sub

Sub Gates_Hit (idx)
  PlaySound "fx_gate", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
End Sub

Sub Spinner_Spin
  PlaySoundAt "fx_spinner", Spinner
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
Sub LeftFlipper1_Collide(parm)
  RandomSoundFlipper()
End Sub

Sub RightFlipper_Collide(parm)
  RandomSoundFlipper()
End Sub
Sub RightFlipper1_Collide(parm)
  RandomSoundFlipper()
End Sub

Sub RandomSoundFlipper()
  Select Case Int(Rnd*3)+1
    Case 1 : PlaySound "flip_hit_1", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
    Case 2 : PlaySound "flip_hit_2", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
    Case 3 : PlaySound "flip_hit_3", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
  End Select
End Sub


Sub editDips
  Dim vpmDips : Set vpmDips = New cvpmDips
  With vpmDips
    .AddForm 700,400,"Spiderman - DIP switches"
    .AddFrame 2,0,190,"Maximum credits",49152,Array("8 credits",0,"10 credits",32768,"15 credits",&H00004000,"25 credits",49152)'dip 15&16
    .AddFrame 2,76,190,"Coin chute 1 and 2 control",&H00002000,Array("seperate",0,"same",&H00002000)'dip 14
    .AddFrame 2,122,190,"Playfield special",&H00200000,Array("replay",0,"extra ball",&H00200000)'dip 22
    .AddFrame 2,168,190,"Game mode",&H00100000,Array("replay",0,"extra ball",&H00100000)'dip 21
    .AddFrame 2,214,190,"3rd coin chute credits control",&H00001000,Array("no effect",0,"add 9",&H00001000)'dip 13
    .AddFrame 2,260,190,"Tilt penalty",&H10000000,Array("game over",0,"ball in play",&H10000000)'dip 29
    .AddChk 2,310,140,Array("Sound when scoring?",&H01000000)'dip 25
    .AddChk 2,325,140,Array("Replay button tune?",&H02000000)'dip 26
    .AddChk 2,340,140,Array("Coin switch tune?",&H04000000)'dip 27
    .AddFrame 205,0,190,"High game to date awards",&H00C00000,Array("not displayed and no award",0,"displayed and no award",&H00800000,"displayed and 2 replays",&H00400000,"displayed and 3 replays",&H00C00000)'dip 23&24
    .AddFrameExtra 205,76,190,"Attract tune",&H0200,Array("no attract tune",0,"attract tune played every 6 minutes",&H0200)'S-board dip 2
    .AddFrame 205,122,190,"Balls per game",&H00010000,Array("5 balls",0,"3 balls",&H00010000)'dip 17
    .AddFrame 205,168,190,"Replay limit",&H00040000,Array("no limit",0,"one per game",&H00040000)'dip 19
    .AddFrame 205,214,190,"Novelty",&H00080000,Array("normal game mode",0,"50K for special/extra ball",&H0080000)'dip 20
    .AddFrameExtra 205,260,190,"Sound option",&H0100,Array("sound mode",0,"tone mode",&H0100)'S-board dip 1
    .AddChk 205,310,140,Array("Credits displayed?",&H08000000)'dip 28
    .AddChk 205,325,140,Array("Match feature",&H00020000)'dip 18
    .AddChk 205,340,140,Array("Attract features",&H20000000)'dip 30
    .AddLabel 50,360,300,20,"After hitting OK, press F3 to reset game with new settings."
  End With
  Dim extra
  extra = Controller.Dip(4) + Controller.Dip(5)*256
  extra = vpmDips.ViewDipsExtra(extra)
  Controller.Dip(4) = extra And 255
  Controller.Dip(5) = (extra And 65280)\256 And 255
End Sub
Set vpmShowDips = GetRef("editDips")

'*************************************************************
'Lamp and Flasher routines
'*************************************************************

Dim LampState(200), FadingLevel(200), FadingState(200)
Dim FlashState(200), FlashLevel(200)
Dim FlashSpeedUp, FlashSpeedDown
Dim x

AllLampsOff()
LampTimer.Interval = 30
LampTimer.Enabled = 1
FlashInit()
FlasherTimer.Interval = 10 'flash fading speed
FlasherTimer.Enabled = 1

Sub LampTimer_Timer()
    Dim chgLamp, num, chg, ii
    chgLamp = Controller.ChangedLamps
    If Not IsEmpty(chgLamp) Then
        For ii = 0 To UBound(chgLamp)
            LampState(chgLamp(ii, 0) ) = chgLamp(ii, 1)
            FadingLevel(chgLamp(ii, 0) ) = chgLamp(ii, 1) + 4
      FlashState(chgLamp(ii, 0) ) = chgLamp(ii, 1)
        Next
    End If

    UpdateLamps
End Sub

Sub FlashInit
    Dim i
    For i = 0 to 200
        FlashState(i) = 0
        FlashLevel(i) = 0
    Next

    FlashSpeedUp = 50   ' fast speed when turning on the flasher
    FlashSpeedDown = 10 ' slow speed when turning off the flasher, gives a smooth fading
    AllFlashOff()
End Sub

Sub AllFlashOff
    Dim i
    For i = 0 to 200
        FlashState(i) = 0
    Next
End Sub

Sub UpdateLamps()
  NFadeL 1, l1
  NFadeL 2, l2
  NFadeL 3, l3
  NFadeL 4, l4
  NFadeL 5, l5
  NFadeL 6, l6
  NFadeL 7, l7
  NFadeL 8, l8
  NFadeL 9, l9
  NFadeL 10, l10
  NFadeL 11, l11
  NFadeL 12, l12
  NFadeL 13, l13
  NFadeL 14, l14
  NFadeL 15, l15
  NFadeL 16, l16
  NFadeL 17, l17
  NFadeL 18, l18
  NFadeL 19, l19
  NFadeL 20, l20
  NFadeL 21, l21
  NFadeL 22, l22
  NFadeL 23, l23
  NFadeL 24, l24
  NFadeLm 25, l25b
  NFadeLm 25, l25c
  NFadeL 25, l25
  NFadeL 26, l26
  NFadeL 27, l27
  NFadeL 28, l28
  NFadeL 29, l29
  NFadeL 30, l30
  NFadeL 31, l31
  NFadeL 32, l32
  NFadeL 33, l33
  NFadeL 34, l34
  NFadeL 35, l35
  NFadeL 36, l36
  NFadeL 37, l37
  NFadeL 38, l38
  NFadeL 39, l39
  NFadeL 40, l40
  NFadeL 41, l41
  NFadeL 42, l42
  NFadeL 43, l43
  NFadeL 44, l44
  NFadeLm 45, l45b
  NFadeL 45, l45
  NFadeLm 46, l46b
  NFadeL 46, l46
  NFadeLm 47, l47zb
  NFadeL 47, l47z
  NFadeL 48, l48
  NFadeLm 49, l49b
  NFadeL 49, l49
  NFadeLm 50, l50b
  NFadeL 50, l50
  NFadeL 51, l51
' NFadeL 52, l52
' NFadeL 53, l53
' NFadeL 54, l54
' NFadeL 55, l55
' NFadeL 56, l56
' NFadeL 57, l57
' NFadeL 58, l58
' NFadeL 59, l59
' NFadeL 60, l60
' NFadeL 61, l61
' NFadeL 62, l62
' NFadeL 63, l63
' NFadeL 64, l64
' NFadeL 65, l65
' NFadeL 66, l66
' NFadeL 67, l67
' NFadeL 68, l68
' NFadeL 69, l69
' NFadeL 70, l70
' NFadeL 71, l71
' NFadeL 72, l72
' NFadeL 73, l73
' NFadeL 74, l74
' NFadeL 75, l75
' NFadeL 76, l76
' NFadeL 77, l77
' NFadeL 78, l78
' NFadeL 79, l79
' NFadeL 125, l125
' NFadeL 126, l126
' NFadeL 127, l127
' NFadeL 131, l131
' NFadeL 132, l132
End Sub

Sub FlasherTimer_Timer()

' Flash 1, F1
' Flash 2, F2
' Flash 3, F3
' Flash 4, F4
' Flash 5, F5
' Flash 9, F9
' Flash 10, F10
' Flash 11, F11
' Flash 12, F12
' Flash 17, F17
' Flash 18, F18
' Flash 19, F19
' Flash 20, F20
' Flash 21, F21
' Flash 33, F33
' Flash 34, F34
' Flash 35, F35
' Flash 36, F36
' Flash 37, F37
' Flash 38, F38
' Flash 57, F57

End Sub

Sub AllLampsOff():For x = 1 to 200:LampState(x) = 4:FadingLevel(x) = 4:Next:UpdateLamps:UpdateLamps:Updatelamps:End Sub

Sub SetLamp(nr, value)
    If value = 0 AND LampState(nr) = 0 Then Exit Sub
    If value = 1 AND LampState(nr) = 1 Then Exit Sub
    LampState(nr) = abs(value) + 4
  FadingLevel(nr) = abs(value) + 4: FadingState(nr ) = abs(value) + 4
  FlashState(nr) = abs(value)
End Sub

Sub NFadeL(nr, a)
    Select Case FadingLevel(nr)
        Case 4:a.state = 0:FadingLevel(nr) = 0
        Case 5:a.State = 1:FadingLevel(nr) = 1
    End Select
End Sub

Sub NFadeLm(nr, a)
    Select Case FadingLevel(nr)
        Case 4:a.state = 0
        Case 5:a.State = 1
    End Select
End Sub

Sub Flash(nr, object)
    Select Case FlashState(nr)
        Case 0 'off
            FlashLevel(nr) = FlashLevel(nr) - FlashSpeedDown
            If FlashLevel(nr) < 0 Then
                FlashLevel(nr) = 0
                FlashState(nr) = -1 'completely off
            End if
            Object.opacity = FlashLevel(nr)
        Case 1 ' on
            FlashLevel(nr) = FlashLevel(nr) + FlashSpeedUp
            If FlashLevel(nr) > 255 Then
                FlashLevel(nr) = 255
                FlashState(nr) = -2 'completely on
            End if
            Object.opacity = FlashLevel(nr)
    End Select
End Sub

Sub Flashm(nr, object) 'multiple flashers, it doesn't change the flashstate
Dim TempLevel
    Select Case FlashState(nr)
        Case 0 'off
            TempLevel = FlashLevel(nr) - FlashSpeedDown
            If TempLevel < 0 Then
                TempLevel = 0
            End if
            Object.opacity = TempLevel
        Case 1 ' on
            TempLevel = FlashLevel(nr) + FlashSpeedUp
            If TempLevel > 255 Then
                TempLevel = 255
            End if
            Object.opacity = TempLevel
    End Select
End Sub

Sub Strobe(nr, object)
    Select Case FlashState(nr)
        Case 0 'off
            FlashLevel(nr) = FlashLevel(nr) - 130
            If FlashLevel(nr) < 0 Then
                FlashLevel(nr) = 0
                FlashState(nr) = -1 'completely off
            End if
            Object.opacity = FlashLevel(nr)
        Case 1 ' on
            FlashLevel(nr) = FlashLevel(nr) + 255
            If FlashLevel(nr) > 255 Then
                FlashLevel(nr) = 255
                FlashState(nr) = -2 'completely on
            End if
            Object.opacity = FlashLevel(nr)
    End Select
End Sub

Sub Strobem(nr, object) 'multiple flashers, it doesn't change the flashstate
Dim TempLevel
    Select Case FlashState(nr)
        Case 0 'off
            TempLevel = FlashLevel(nr) - 130
            If TempLevel < 0 Then
                TempLevel = 0
            End if
            Object.opacity = TempLevel
        Case 1 ' on
            TempLevel = FlashLevel(nr) + 255
            If TempLevel > 255 Then
                TempLevel = 255
            End if
            Object.opacity = TempLevel
    End Select
End Sub




Sub B2SCommand(nr, state)
  If cController = 3 Then
    Controller.B2SSetData nr, state
  End If
End Sub

Sub DOF(dofevent, dofstate)
  '*******Use DOF 1**, 1 to activate a ledwiz output*******************
  '*******Use DOF 1**, 0 to deactivate a ledwiz output*****************
  '*******Use DOF 1**, 2 to pulse a ledwiz output**********************
  If cController > 2 Then
    If dofstate = 2 Then
      Controller.B2SSetData dofevent, 1
      Controller.B2SSetData dofevent, 0
    Else
      Controller.B2SSetData dofevent, dofstate
    End If
  End If
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

Sub PlaySoundAtVol(sound, tableobj, Vol)
  PlaySound sound, 1, Vol, Pan(tableobj), 0,0,0, 1, AudioFade(tableobj)
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
  Vol = Csng(BallVel(ball) ^2 / 2000)
End Function

Function Pitch(ball) ' Calculates the pitch of the sound based on the ball speed
  Pitch = BallVel(ball) * 20
End Function

Function BallVel(ball) 'Calculates the ball speed
  BallVel = INT(SQR((ball.VelX ^2) + (ball.VelY ^2) ) )
End Function


'*****************************************
'    JP's VP10 Collision & Rolling Sounds
'*****************************************

Const tnob = 16 ' total number of balls
ReDim rolling(tnob)
ReDim collision(tnob)
Initcollision

Sub Initcollision
    Dim i
    For i = 0 to tnob
        collision(i) = -1
        rolling(i) = False
    Next
End Sub

Sub CollisionTimer_Timer()

  'Added flipper and gate and spinner rotation here
  UpdateLeftFlipperLogo
  UpdateRightFlipperLogo
  PrimGate3.Rotz = Gate3.CurrentAngle * 70/90
  PrimGate1.Rotz = Gate1.CurrentAngle * 70/90
  PrimSw60.Rotx = -(sw60.CurrentAngle-90)



  Dim BOT, B, B1, B2, dx, dy, dz, distance, radii
  BOT = GetBalls

  ' rolling

  For B = UBound(BOT) +1 to tnob
        rolling(b) = False
        StopSound("fx_ballrolling" & b)
  Next

    If UBound(BOT) = -1 Then Exit Sub

    For B = 0 to UBound(BOT)
        If BallVel(BOT(b) ) > 1 AND BOT(b).z < 30 Then
            rolling(b) = True
            PlaySound("fx_ballrolling" & b), -1, Vol(BOT(b) ), Pan(BOT(b) ), 0, Pitch(BOT(b) ), 1, 0
        Else
            If rolling(b) = True Then
                StopSound("fx_ballrolling" & b)
                rolling(b) = False
            End If
        End If
    Next

    'collision

    If UBound(BOT) < 1 Then Exit Sub

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
  If Table1.VersionMinor > 3 OR Table1.VersionMajor > 10 Then
    PlaySound("fx_collide"), 0, Csng(velocity) ^2 / 200, Pan(ball1), 0, Pitch(ball1), 0, 0, AudioFade(ball1)
  Else
    PlaySound("fx_collide"), 0, Csng(velocity) ^2 / 200, Pan(ball1), 0, Pitch(ball1), 0, 0
  End if
End Sub


' Thalamus : Exit in a clean and proper way
Sub Table1_exit()
  Controller.Pause = False
  Controller.Stop
End Sub


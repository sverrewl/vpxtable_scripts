'############################################################################################
'############################################################################################
'#######                                                                             ########
'#######          Catacomb                                                           ########
'#######          (Stern 1981)           VPX conversion HSM                          ########
'#######                                                                             ########
'############################################################################################
'############################################################################################
'
' Version 1.4 mfuegemann 2014
'
' Thanks to Destruk/Lander/Skalar for light numbers and solenoid corrections in their VP8 table
' Thanks to Inkochnito for the DIP switch settings code
' Thalmus - seems ok
Option Explicit
On Error Resume Next
ExecuteGlobal GetTextFile("controller.vbs")
If Err Then MsgBox "You need the Controller.vbs file in order to run this table (installed with the VPX package in the scripts folder)"
On Error Goto 0

LoadVPM "01560000","STERN.VBS",3.1

Const UseSolenoids=2,UseLamps=1,SSolenoidOn="SolOn",SSolenoidOff="SolOff",SCoin="coin3"
Const UseSync = 0

'-----------------------------------
'------  Solenoid Assignment  ------
'-----------------------------------
SolCallback(6)  = "vpmSolSound ""Knocker"","            'Sol3 Knocker --> 6 Lander
SolCallback(7)  = "bsLeftKicker.SolOut"         'Sol4 Left Saucer --> 7 Lander
SolCallback(8)  = "bsRightKicker.SolOut"          'Sol8 Right Saucer
SolCallback(11) = "DropTargetBank_A.SolDropUp"          'Sol9 Drop Target A --> 11 Lander
SolCallback(12) = "DropTargetBank_B.SolDropUp"          'Sol10 Drop Target B --> 12 Lander
SolCallback(14) = "DropTargetBank_D.SolDropUp"          'Sol11 Drop Target D --> 14 Lander
SolCallback(13) = "DropTargetBank_C.SolDropUp"          'Sol12 Drop Target C --> 13 Lander
SolCallback(9)  = "bsTrough.SolIn"                      'Sol13 Outhole --> 9 Lander
SolCallback(10) = "SolBallRelease"                      'Sol14 Ball Release --> 10 Lander
SolCallback(19) = "vpmNudge.SolGameOn"          'Sol15 --> 19 Lander
SolCallback(20) = "SolBackbox"              'Sol18 BackBox Game --> 20 Lander
SolCallback(sLRFlipper) = "SolRFlipper"
SolCallback(sLLFlipper) = "SolLFlipper"


Set LampCallback= GetRef("UpdateMultipleLamps")
Set vpmShowDips = GetRef("editDips")

Dim Objekt
  If Catacomb.ShowDT = False then
  for each Objekt in backdropstuff:Objekt.visible = False:Next
  End If


'--------------------------
'------  Table Init  ------
'--------------------------
Dim bsTrough,DropTargetBank_A,DropTargetBank_B,DropTargetBank_C,DropTargetBank_D,GIState,obj,bsRightKicker,bsLeftKicker,InitGravity,TableInitOK
Dim X,Y,YY,XX,Square,OldSquare,PrevOldSquare,Prev2OldSquare,Flash
Const BackBoxGravity = 3
Sub Catacomb_Init
  TableInitOK = False
  vpminit me
  InitGravity = Catacomb.gravity
  HideBackBox
  Const cgamename = "catacomb"
    Controller.GameName=cGameName
    Controller.SplashInfoLine="Catacomb" & vbNewLine & "Created by mfuegemann"
    Controller.HandleKeyboard=False
    Controller.ShowTitle=0
    Controller.ShowFrame=0
    Controller.ShowDMDOnly=1
    Controller.HandleMechanics=0
    Controller.Run
    If Err Then MsgBox Err.Description
    On Error Goto 0

    PinMAMETimer.Interval=PinMAMEInterval
  PinMAMETimer.Enabled = true

    vpmNudge.TiltSwitch=7
    vpmNudge.Sensitivity=2
  vpmNudge.TiltObj = Array(Bumper1,Bumper2,LeftslingShot,RightslingShot,UpperLeftslingshot)

    vpmMapLights AllLights

    Set bsTrough=New cvpmBallStack
        bsTrough.InitSw 17,33,34,35,0,0,0,0
        bsTrough.InitKick BallRelease,90,5
        bsTrough.InitExitSnd "BallRelease","Solenoid"
        bsTrough.Balls=3

  set DropTargetBank_A = new cvpmDropTarget
    DropTargetBank_A.InitDrop Array(DT_A_Left,DT_A_Middle,DT_A_Right), Array(16,15,14)
    DropTargetBank_A.InitSnd "Target","fx_resetdrop"
  set DropTargetBank_B = new cvpmDropTarget
    DropTargetBank_B.InitDrop Array(DT_B_Left,DT_B_Middle,DT_B_Right), Array(24,23,22)
    DropTargetBank_B.InitSnd "Target","fx_resetdrop"
  set DropTargetBank_C = new cvpmDropTarget
    DropTargetBank_C.InitDrop Array(DT_C_Left,DT_C_Middle,DT_C_Right), Array(32,31,30)
    DropTargetBank_C.InitSnd "Target","fx_resetdrop"
  set DropTargetBank_D = new cvpmDropTarget
    DropTargetBank_D.InitDrop Array(DT_D_Left,DT_D_Middle,DT_D_Right), Array(40,39,38)
    DropTargetBank_D.InitSnd "Target","fx_resetdrop"

    ' Thalamus - more randomness to kickers pls
    Set bsLeftKicker=New cvpmBallStack
        bsLeftKicker.KickForceVar = 3
        bsLeftKicker.KickAngleVar = 3
        bsLeftKicker.InitSaucer LeftKicker,36,90,10
        bsLeftKicker.InitExitSnd "popper_ball","solon"

    Set bsRightKicker=New cvpmBallStack
        bsRightKicker.KickForceVar = 3
        bsRightKicker.KickAngleVar = 3
        bsRightKicker.InitSaucer RightKicker,37,260,5
        bsRightKicker.InitExitSnd "popper_ball","solon"

    GIState = LightstateOff
End Sub

Sub DropAreset()
  DT_A_Left.isdropped=0
  DT_A_Middle.isdropped=0
  DT_A_Right.isdropped=0
End Sub

Sub Catacomb_exit()
  Controller.Pause = False
  Controller.Stop
End Sub

'------------------------------
'------  Trough Handler  ------
'------------------------------
Sub Drain_Hit()
  SoundTimer.enabled = 0
  bsTrough.AddBall Me
  PlaySoundAtVol "Drain", Drain, 1
End Sub

'---------------------------------
'------  Backbox Functions  ------
'---------------------------------

Sub SolBallRelease(Enabled)
  If Enabled Then
    HideBackBox
    bsTrough.ExitSol_On
    TableInitOK = True
  End If
End Sub

Dim BackboxActive,BBall

Sub SolBackbox(Enabled)
  If Enabled Then
    ActivateBackBox
  else
    DeactivateBackBox
  End If
End Sub

Sub ActivateBackBox
  if B2Son then
    Controller.B2SSetData "MaskAboveFlipper", 1
    Controller.B2SSetData "BackboxActive", 1
    BallPosTimer.enabled = True
  end if
  if TableInitOK = True then
    BackboxActive = True
    for each obj in Backbox
      obj.visible = True
    next
    for each obj in BackboxWalls
      obj.visible = True
      obj.Sidevisible = True
      obj.collidable = True
    next
    for each obj in BackBoxRubber
      obj.collidable = True
      obj.visible = True
    Next

    BackBoxLight.State = LightstateOn
    Catacomb.gravity = BackBoxGravity
    BackDestroyBall.enabled=False
    set BBall = BKicker.Createsizedball(13)
    BBall.Image = "Yball"
    BKicker.kick 180,1
  end if
End Sub

Sub DeactivateBackBox
  BackDestroyBall.enabled=True
  HideBackboxTimer.enabled = True
  BallPosTimer.enabled = False
  BackBoxActive = False
  BackBoxLight.State = LightstateOff
  BFlipper.RotateToStart
  Catacomb.gravity = InitGravity

  if B2Son then
    Controller.B2SSetData "BackboxActive", 0
    for xx = 1 to 23
      for yy = 1 to 23
        Controller.B2SSetData CStr(10000 + (XX * 100) + YY), 0
      next
    next
    Square = 10101
    Controller.B2SSetData "Flipper_up", 0
    Controller.B2SSetData "Flipper_down", 1
    Controller.B2SSetData CStr(Square), 1
  End If
End Sub

Sub HideBackboxTimer_Timer
  HideBackboxTimer.enabled = false
  HideBackBox
End Sub

Sub HideBackBox
  If B2Son Then
    for xx = 1 to 23
      for yy = 1 to 23
        Controller.B2SSetData CStr(10000 + (XX * 100) + YY), 0
      next
    next
    Square = 10101
    Controller.B2SSetData "Flipper_up", 0
    Controller.B2SSetData "Flipper_down", 1
    Controller.B2SSetData CStr(Square), 1
  End If
  BackBoxActive = False
  for each obj in Backbox
    obj.visible = False
  next
  for each obj in BackBoxRubber
    obj.collidable = False
    obj.visible = False
  Next
  for each obj in BackboxWalls
    obj.visible = False
    obj.Sidevisible = False
    obj.collidable = False
  next
  BackBoxLight.State = LightstateOff
  BFlipper.RotateToStart
  Catacomb.gravity = InitGravity
End Sub

Sub BackDestroyBall_Hit
  BackDestroyBall.destroyball
End Sub

'---------------------------------
'-----  B2S Ball Simulation  -----
'---------------------------------
OldSquare = 10101     'Rest Position
PrevOldSquare = 10101
Prev2OldSquare = 10101

Sub BallPosTimer_Timer
  if Backboxactive then
    X = Round((BOrigin.X - BBall.X)/21.3)         '21.3 is the x grid size
    Y = Round((BOrigin.Y - BBall.Y)/21.3)         '21.3 is the y grid size

    if (Y<20) and (Y>4) and (X<3) then X = 1        'shooter lane exception
    if (Y<18) and (Y>4) and (X=13) then X = 12        'lane A exception
    if (Y<18) and (Y>4) and ((X=9) or (X=11)) then X = 10 'lane B exception
    if (Y<18) and (Y>4) and (X=7) then X = 8        'lane C exception
    if (Y<18) and (Y>4) and ((X=4) or (X=6)) then X = 5   'lane D exception
    if (Y<5) and (Y>2) and ((X>1) and (X<5)) then Y = 4   'return lane exception
    if X<1 then X = 1                   'boundary check
    if X>15 then X = 15
    if Y<1 then Y = 1
    if Y>23 then Y = 23

    Square = 10000 + (X * 100) + Y              'calculates the current grid position string
    if Square <> OldSquare then
      Controller.B2SSetData CStr(Prev2OldSquare), 0
      Controller.B2SSetData CStr(PrevOldSquare), 0
      Controller.B2SSetData CStr(OldSquare), 0
      Controller.B2SSetData CStr(Square), 1
      Prev2OldSquare = PrevOldSquare
      PrevOldSquare = OldSquare
      Oldsquare = Square
    end if
  end if
End Sub

Sub SolLFlipper(Enabled)
    If Enabled Then
        PlaySoundAtVol SoundFX("fx_flipperup", DOFFlippers), LeftFlipper, 1
        LeftFlipper.RotateToEnd
    Else
        PlaySoundAtVol SoundFX("fx_flipperdown", DOFFlippers), LeftFlipper, 1
        LeftFlipper.RotateToStart
    End If
End Sub

Sub SolRFlipper(Enabled)
    If Enabled Then
        PlaySoundAtVol SoundFX("fx_flipperup", DOFFlippers), RightFlipper, 1
        RightFlipper.RotateToEnd
    Else
        PlaySoundAtVol SoundFX("fx_flipperdown", DOFFlippers), RightFlipper, 1
        RightFlipper.RotateToStart
    End If
End Sub

'-------------------------------
'------  Keybord Handler  ------
'-------------------------------
Sub catacomb_KeyDown(ByVal keycode)
  If KeyDownHandler(keycode) Then Exit Sub
  If keycode = PlungerKey Then Plunger.PullBack:PlaySoundAtVol "plungerpull", Plunger, 1

  If keycode = LeftFlipperKey Then
    If BackBoxActive = True Then
      Controller.Switch(8)=1      '8 Lander
    End If
  End If

  If keycode = RightFlipperKey Then
    If BackBoxActive = True Then
      BFlipper.RotateToEnd
      PlaySoundAtVol SoundFX("fx_flipperup", DOFFlippers), BFlipper, 1
      If B2Son=True Then
        Controller.B2SSetData "Flipper_down", 0   'B2S
        Controller.B2SSetData "Flipper_up", 1   'B2S
      End If
    End If
  End If
End Sub

Sub catacomb_KeyUp(ByVal keycode)
  If KeyUpHandler(keycode) Then Exit Sub
  If keycode = PlungerKey Then Plunger.Fire: PlaySoundAtVol "plunger", Plunger, 1

  If keycode = LeftFlipperKey Then
    If BackBoxActive = True Then
      Controller.Switch(8)=0       '8 Lander
    End If
  End If
  If keycode = RightFlipperKey Then
    If BackBoxActive = True Then
      PlaySoundAtVol SoundFX("fx_flipperdown", DOFFlippers), BFlipper, 1
      BFlipper.RotateToStart
      If B2Son=True Then
        Controller.B2SSetData "Flipper_up", 0     'B2S
        Controller.B2SSetData "Flipper_down", 1     'B2S
      End If
    End If
  End If
End Sub

'------------------------------
'------  Switch Handler  ------
'------------------------------
Sub Spinner4_Spin:vpmTimer.PulseSwitch 4,0,0:PlaySound "fx_spinner", 0, .25, AudioPan(Spinner4), 0.25, 0, 0, 1, AudioFade(Spinner4):End Sub
sub Trigger5_hit:Controller.Switch(5)=1:End Sub   '5 Left Gate
sub Trigger5_unhit:Controller.Switch(5)=0:End Sub

sub Bumper2_hit                   '9 Right Bumper
  vpmTimer.pulseSw 9
  RandomSoundBumper
End Sub

sub Bumper1_hit                   '10 Left Bumper
  vpmTimer.pulseSw 10
  RandomSoundBumper
End Sub


Sub RandomSoundBumper()
  Select Case Int(Rnd*4)+1
    Case 1 : PlaySoundAtVol "fx_bumper1", ActiveBall, 1
    Case 2 : PlaySoundAtVol "fx_bumper2", ActiveBall, 1
    Case 3 : PlaySoundAtVol "fx_bumper3", ActiveBall, 1
    Case 4 : PlaySoundAtVol "fx_bumper4", ActiveBall, 1
  End Select
End Sub

'**********Sling Shot Animations
' Rstep and Lstep  are the variables that increment the animation
'****************
Dim RStep, Lstep

Sub RightSlingShot_Slingshot
  vpmTimer.pulseSw 13
    PlaySoundAtVol SoundFX("right_slingshot",DOFContactors), ActiveBall, 1
    RSling.Visible = 0
    RSling1.Visible = 1
    sling1.rotx = 20
    RStep = 0
    RightSlingShot.TimerEnabled = 1
End Sub

Sub RightSlingShot_Timer
    Select Case RStep
        Case 3:RSLing1.Visible = 0:RSLing2.Visible = 1:sling1.rotx = 10
        Case 4:RSLing2.Visible = 0:RSLing.Visible = 1:sling1.rotx = 0:RightSlingShot.TimerEnabled = 0
    End Select
    RStep = RStep + 1
End Sub

Sub LeftSlingShot_Slingshot
  vpmTimer.pulseSw 12
    PlaySoundAtVol SoundFX("left_slingshot",DOFContactors), ActiveBall, 1
    LSling.Visible = 0
    LSling1.Visible = 1
    sling2.rotx = 20
    LStep = 0
    LeftSlingShot.TimerEnabled = 1
End Sub

Sub LeftSlingShot_Timer
    Select Case LStep
        Case 3:LSLing1.Visible = 0:LSLing2.Visible = 1:sling2.rotx = 10
        Case 4:LSLing2.Visible = 0:LSLing.Visible = 1:sling2.rotx = 0:LeftSlingShot.TimerEnabled = 0
    End Select
    LStep = LStep + 1
End Sub

Sub UpperLeftSlingShot_Slingshot
  vpmTimer.pulseSw 11
    PlaySoundAtVol SoundFX("left_slingshot",DOFContactors), ActiveBall, 1
    ULSling.Visible = 0
    ULSling1.Visible = 1
    sling3.rotx = 20
    LStep = 0
    UpperLeftSlingShot.TimerEnabled = 1
End Sub

Sub UpperLeftSlingShot_Timer
    Select Case LStep
        Case 3:ULSling1.Visible = 0:ULSling2.Visible = 1:sling3.rotx = 10
        Case 4:ULSling2.Visible = 0:ULSling.Visible = 1:sling3.rotx = 0:UpperLeftSlingShot.TimerEnabled = 0
    End Select
    LStep = LStep + 1
End Sub

Sub DT_A_Right_Dropped:DropTargetBank_A.Hit 3:End Sub      '14
Sub DT_A_Middle_Dropped:DropTargetBank_A.Hit 2:End Sub       '15
Sub DT_A_Left_Dropped:DropTargetBank_A.Hit 1:End Sub       '16

sub Trigger18_hit:Controller.Switch(18)=1:PlaySoundAtVol "Gate1", ActiveBall, 1:End Sub   '18 Left Outlane
sub Trigger18_unhit:Controller.Switch(18)=0:Trigger18.timerenabled=True:End Sub
sub Trigger19_hit:Controller.Switch(19)=1:PlaySoundAtVol "Gate1", ActiveBall, 1:End Sub    '19 Right Outlane
sub Trigger19_unhit:Controller.Switch(19)=0:Trigger19.timerenabled=True:End Sub
sub Trigger20_hit:Controller.Switch(20)=1:PlaySoundAtVol "Gate1", ActiveBall, 1:End Sub   '20 Left Inlane
sub Trigger20_unhit:Controller.Switch(20)=0:Trigger20.timerenabled=True:End Sub
sub Trigger21_hit:Controller.Switch(21)=1:PlaySoundAtVol "Gate1", ActiveBall, 1:End Sub    '21 Right Inlane
sub Trigger21_unhit:Controller.Switch(21)=0:Trigger21.timerenabled=True:End Sub

Sub DT_B_Right_Dropped:DropTargetBank_B.Hit 3:End Sub      '22
Sub DT_B_Middle_Dropped:DropTargetBank_B.Hit 2:End Sub       '23
Sub DT_B_Left_Dropped:DropTargetBank_B.Hit 1:End Sub       '24

sub Trigger25_hit:Controller.Switch(25)=1:PlaySoundAtVol "Gate1", ActiveBall, 1:End Sub    '25 Spot 6 Rollover
sub Trigger25_unhit:Controller.Switch(25)=0:Trigger25.timerenabled=True:End Sub
sub Trigger26_hit:Controller.Switch(26)=1:PlaySoundAtVol "Gate1", ActiveBall, 1:End Sub   '26 Spot 1 Rollover
sub Trigger26_unhit:Controller.Switch(26)=0:Trigger26.timerenabled=True:End Sub
Sub Target27_Hit:vpmTimer.PulseSw 27:PlaySoundAtVol "Target1", ActiveBall, 1:End Sub      '27 Spot 2 Target
sub Trigger28_hit:Controller.Switch(28)=1:End Sub   '28 Highscore Rollover
sub Trigger28_unhit:Controller.Switch(28)=0:End Sub
sub Trigger29_hit:Controller.Switch(29)=1:PlaySoundAtVol "Gate", ActiveBall, 1:End Sub   '29 Spinner Value Rollover
sub Trigger29_unhit:Controller.Switch(29)=0:End Sub ':Trigger29.timerenabled=True

Sub DT_C_Right_Dropped:DropTargetBank_C.Hit 3:End Sub      '30
Sub DT_C_Middle_Dropped:DropTargetBank_C.Hit 2:End Sub       '31
Sub DT_C_Left_Dropped:DropTargetBank_C.Hit 1:End Sub       '32

Sub LeftKicker_Hit:bsLeftKicker.AddBall 0:End Sub          '36
Sub RightKicker_Hit:bsRightKicker.AddBall 0:End Sub        '37

Sub DT_D_Right_Dropped:DropTargetBank_D.Hit 3:End Sub      '38
Sub DT_D_Middle_Dropped:DropTargetBank_D.Hit 2:End Sub       '39
Sub DT_D_Left_Dropped:DropTargetBank_D.Hit 1:End Sub       '40

sub TriggerB14_hit:Controller.Switch(14)=1:End Sub    '14 Backbox
sub TriggerB14_unhit:Controller.Switch(14)=0:End Sub
sub TriggerB15_hit:Controller.Switch(15)=1:End Sub    '15 Backbox
sub TriggerB15_unhit:Controller.Switch(15)=0:End Sub
sub TriggerB16_hit:Controller.Switch(16)=1:End Sub    '16 Backbox
sub TriggerB16_unhit:Controller.Switch(16)=0:End Sub
sub TriggerB22_hit:Controller.Switch(22)=1:End Sub    '22 Backbox
sub TriggerB22_unhit:Controller.Switch(22)=0:End Sub
sub TriggerB23_hit:Controller.Switch(23)=1:End Sub    '23 Backbox
sub TriggerB23_unhit:Controller.Switch(23)=0:End Sub
sub TriggerB24_hit:Controller.Switch(24)=1:End Sub    '24 Backbox
sub TriggerB24_unhit:Controller.Switch(24)=0:End Sub
sub TriggerB30_hit:Controller.Switch(30)=1:End Sub    '30 Backbox
sub TriggerB30_unhit:Controller.Switch(30)=0:End Sub
sub TriggerB31_hit:Controller.Switch(31)=1:End Sub    '31 Backbox
sub TriggerB31_unhit:Controller.Switch(31)=0:End Sub
sub TriggerB32_hit:Controller.Switch(32)=1:End Sub    '32 Backbox
sub TriggerB32_unhit:Controller.Switch(32)=0:End Sub
sub TriggerB38_hit:Controller.Switch(38)=1:End Sub    '38 Backbox
sub TriggerB38_unhit:Controller.Switch(38)=0:End Sub
sub TriggerB39_hit:Controller.Switch(39)=1:End Sub    '39 Backbox
sub TriggerB39_unhit:Controller.Switch(39)=0:End Sub
sub TriggerB40_hit:Controller.Switch(40)=1:End Sub    '40 Backbox
sub TriggerB40_unhit:Controller.Switch(40)=0:End Sub

'________________________________________
'Drop Target Lighting
'-------------------------------------

Sub DropTargetLight_Timer()
  If DT_D_Left.isdropped then
  DTL3.state=1
    Else
  DTL3.state=0
  End If
  If DT_D_Middle.isdropped then
  DTL2.state=1
    Else
  DTL2.state=0
  End If
  If DT_D_Right.isdropped then
  DTL1.state=1
    Else
  DTL1.state=0
  End If
  If DT_C_Left.isdropped then
  DTL4.state=1
    Else
  DTL4.state=0
  End If
  If DT_C_Middle.isdropped then
  DTL5.state=1
    Else
  DTL5.state=0
  End If
  If DT_C_Right.isdropped then
  DTL6.state=1
    Else
  DTL6.state=0
  End If
  If DT_B_Left.isdropped then
  DTL7.state=1
    Else
  DTL7.state=0
  End If
  If DT_B_Middle.isdropped then
  DTL8.state=1
    Else
  DTL8.state=0
  End If
  If DT_B_Right.isdropped then
  DTL9.state=1
    Else
  DTL9.state=0
  End If
  If DT_A_Left.isdropped then
  DTL10.state=1
    Else
  DTL10.state=0
  End If
  If DT_A_Middle.isdropped then
  DTL11.state=1
    Else
  DTL11.state=0
  End If
  If DT_A_Right.isdropped then
  DTL12.state=1
    Else
  DTL12.state=0
  End If
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
        If BOT(b).X < Catacomb.Width/2 Then
            BallShadow(b).X = ((BOT(b).X) - (Ballsize/6) + ((BOT(b).X - (Catacomb.Width/2))/21)) + 6
        Else
            BallShadow(b).X = ((BOT(b).X) + (Ballsize/6) + ((BOT(b).X - (Catacomb.Width/2))/21)) - 6
        End If
        ballShadow(b).Y = BOT(b).Y + 4
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
  PlaySound "gate", 0, Vol(ActiveBall), AudioPan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
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


'Stern Catacomb
' added by Inkochnito
Sub editDips
  Dim vpmDips:Set vpmDips = New cvpmDips
    With vpmDips
    .AddForm 700,400,"Catacomb - DIP switches"
    .AddChk 2,10,180,Array("Match feature",&H00100000)'dip 21
    .AddChk 2,25,115,Array("Credits display",&H00080000)'dip 20
    .AddFrame 2,45,190,"Maximum credits",&H00060000,Array("10 credits",0,"15 credits",&H00020000,"25 credits",&H00040000,"40 credits",&H00060000)'dip 18&19
    .AddFrame 2,120,190,"High game to date",49152,Array("points",0,"1 free game",&H00004000,"2 free games",32768,"3 free games",49152)'dip 15&16
    .AddFrame 2,195,190,"High score feature",&H00000020,Array("extra ball",0,"replay",&H00000020)'dip 6
    .AddChk 205,10,190,Array("Talking sound",&H00010000)'dip 17
    .AddChk 205,25,190,Array("Background sound",&H00000080)'dip 8
    .AddFrame 205,45,190,"Special award",&HC0000000,Array("no award",0,"100,000 points",&H40000000,"free ball",&H80000000,"free game",&HC0000000)'dip 31&32
    .AddFrame 205,120,190,"Add a ball memory",&H00E00000,Array ("no ball",0,"one ball",&H00800000,"three balls",&H00C00000,"five balls",&H00E00000)'dip 22&23&24
    .AddFrame 205,195,190,"Balls per game",&H00000040,Array("3 balls",0,"5 balls",&H00000040)'dip 7
    .AddLabel 50,245,300, 20,"After hitting OK, press F3 to reset game with new settings."
    .ViewDips
  End With
End Sub



  '************************************
 '          LEDs Display
 'Based on Scapino's 7 digit Reel LEDs
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

 Sub LED_Timer
     On Error Resume Next
     Dim ChgLED, ii, jj, chg, stat
     ChgLED = Controller.ChangedLEDs(&HFF, &HFFFF)
     If Not IsEmpty(ChgLED) Then
         For ii = 0 To UBound(ChgLED)
             chg = chgLED(ii, 1):stat = chgLED(ii, 2)

             For jj = 0 to 10
                 If stat = Patterns(jj) OR stat = Patterns2(jj) then Digits(chgLED(ii, 0) ).SetValue jj
             Next
         Next
     End IF
 End Sub
'BackDrop Lights
 Dim BDL

Sub UpdateMultipleLamps

  BDL=Controller.Lamp(13) 'HS to Date
  If BDL Then
    HighScore_Box.text="HIGH SCORE"
    Else
    HighScore_Box.text=""
  End If

  BDL=Controller.Lamp(45) 'Game Over
  If BDL Then
    GameOver_Box.text="GAME OVER"
    Else
    GameOver_Box.text=""
  End If

  BDL=Controller.Lamp(61) 'Tilt
  If BDL Then
    Tilt_Box.text="TILT"
    Else
    Tilt_Box.text=""
  End If

  BDL=Controller.Lamp(11) 'Shoot Again
  If BDL Then
    ShootAgain_Box.text="SHOOT AGAIN"
    Else
    ShootAgain_Box.text=""
  End If
End Sub

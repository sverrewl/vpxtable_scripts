Option Explicit
Randomize

On Error Resume Next
ExecuteGlobal GetTextFile("controller.vbs")
If Err Then MsgBox "You need the controller.vbs in order to run this table, available in the vp10 package"
On Error Goto 0

Const ballSize = 52

Dim VarHidden, UseVPMDMD
If goldeneye.ShowDT = true then
UseVPMDMD = true
VarHidden = 1
else
UseVPMDMD = false
VarHidden = 0
end if

LoadVPM "01120100", "SEGA2.VBS", 3.02

Const cGameName="gldneye",UseLamps=1
Const SSolenoidOn="solon",SSolenoidOff="soloff",SFlipperOn="FlipperUp",SFlipperOff="FlipperDown",sCoin="coin3"


SolCallback(1)="bsTrough.SolOut"
SolCallback(2)="bsPlunger.SolOut"
SolCallback(4)="bsScoop.SolOut"
SolCallback(8)="vpmSolSound ""fx_knocker"","
'SolCallback(9)="vpmSolSound ""jet"","
'SolCallback(10)="vpmSolSound ""jet"","
'SolCallback(11)="vpmSolSound ""jet"","
'SolCallback(12)="vpmSolSound ""sling"","
'SolCallback(13)="vpmSolSound ""sling"","
SolCallback(14)="Tank_Sol"
SolCallback(17)="SolLockOut"
SolCallback(18)="SolUpDownRamp"
SolCallback(20)="SolSatLaunchRamp"
SolCallback(21)="SolSatMotorRelay"
SolCallback(22)="Tank_Trapdoor"
SolCallBack(33) = "SatMag_Sol"
SolCallback(34)= "FlipMag_Sol"
SolCallback(25)="Sol25" '
SolCallback(26)="Sol26" '
SolCallback(27)="Sol27" '
SolCallback(28)="Sol28" '
SolCallback(29)="Sol29" '
SolCallback(30)="Sol30"
SolCallback(31)="Sol31" '
SolCallback(32)="Sol32" '


'General Illumination
set GICallback = GetRef("UpdateGI")
Sub UpdateGI(no, Enabled)
  If Enabled Then
    dim xx
    For each xx in GI:xx.State = 1: Next
        PlaySound "fx_relay"
        MiddleWheel.disablelighting = 1
        Tank.disablelighting = 1
        Chopper.DisableLighting = 1:rotornew.DisableLighting = 1
        rotornew001.visible = 1
        MiddleWheel001.visible = 1
        Chopper001.visible = 1
    DOF 100, DOFOn
        goldeneye.ColorGradeImage = "ColorGradeEX_8"
  Else For each xx in GI:xx.State = 0: Next
        PlaySound "fx_relay"
        MiddleWheel.disablelighting = 0
        Tank.disablelighting = 0
        Chopper.DisableLighting = 0:rotornew.DisableLighting = 0
    DOF 100, DOFOff
        rotornew001.visible = 0
        MiddleWheel001.visible = 0
        Chopper001.visible = 0
        goldeneye.ColorGradeImage = "ColorGradeEX_5"
  End If
End Sub


'FLASHERS
Sub Sol25(Enabled)
F25_1.State = Enabled
F25_2.State = Enabled
F25_3.State = Enabled
F25_4.State = Enabled
F25_5.State = Enabled
F25_6.State = Enabled
End Sub

Sub Sol26(Enabled)
F26_1.State = Enabled
F26_2.State = Enabled
End Sub

Sub Sol30(Enabled)
F30_H1.State = Enabled
F30_H2.State = Enabled
F30_H3.State = Enabled
F30_H4.State = Enabled
End Sub

Sub Sol28(Enabled)
F28_1.State = Enabled
F28_2.State = Enabled
F28_3.State = Enabled
F28_4.State = Enabled
'MiddleWheel001.visible = Enabled
End Sub

Sub Sol31(Enabled)
F31_1.State = Enabled
F31_2.State = Enabled
F31_3.State = Enabled
F31_4.State = Enabled
F31_5.State = Enabled
End Sub

Sub Sol32(Enabled)
F32_1.State = Enabled
F32_2.State = Enabled
F32_3.State = Enabled
'MiddleWheel001.visible = Enabled
End Sub

Sub Sol29(Enabled)
F29_1.State = Enabled
F29_2.State = Enabled
F29_3.State = Enabled
F29_4.State = Enabled
F29_5.State = Enabled
End Sub

Sub Sol27(Enabled)
F27_1.State = Enabled
F27_2.State = Enabled
F27_3.State = Enabled
F27_4.State = Enabled
F27_5.State = Enabled
End Sub


'Flipper Callbacks
SolCallback(sLRFlipper) = "SolRFlipper"
SolCallback(sLLFlipper) = "SolLFlipper"


'# Tank Trapdoor

Sub Tank_Trapdoor(enabled)
If enabled Then
DropRamp1.Enabled=1
PlaySoundAt "fx_solenoid", Tank
F31_003.State = 2
Else
PlaySoundAt "soloff", Tank
DropRamp1.Enabled=0
F31_003.State = 0
End If
End Sub

Sub Tank_Sol(enabled)
If enabled Then
bsTank.ExitSol_on
F31_003.State = 1
Else
F31_003.State = 0
End If
End Sub

'########################################

'Magnet Solenoid for Flipper Ball Save

Dim fmstate

Sub FlipMag_Sol(Enabled)
If Enabled Then
FlipMag.MagnetOn = 1
FlipMag_Kick.enabled = 1
FM_Flash.State = 2
fmstate = 1
Else
FlipMag.MagnetOn = 0
If fmta = 0 Then
FlipMag_Kick.enabled = 0
End If
FM_Flash.State = 0
fmstate = 0
End If
End Sub


'#Flipper Magnet Trigger Hit.
Sub FlipMag_kick_Hit()
Controller.Switch(24) = 1
Light006.State = 1
fmts = 1
Flip_Mag_Timer.enabled = 1
End Sub

Dim fmts,fmta
Sub Flip_Mag_Timer_Timer()
Select Case fmts
Case 1:
fmta = 1
FlipMag.MagnetOn = 0
fmts = 2
Case 2:
FlipMag_Kick.kick 0+(Rnd*16),15+(Rnd*16)
fmts = 3
Case 3:
Controller.Switch(24) = 0
fmts = 4
Case 4:
If fmstate = 1 Then
FlipMag.MagnetOn = 1
End If
Light006.State = 0
fmta = 0
me.enabled = 0
End Select
End Sub
'########################################

'Satellite Centre Magnet Solenoid
Dim smagon, smagball

Sub SatMag_Sol(Enabled)
If enabled Then
Sat_Kick.enabled = 1
smagon = 1
Light006.State = 1
Else
Light006.State = 0
If smagball = 1 Then
Sat_Kick.enabled = 0
Sat_Kick.createball
'Sat_Kick.kick 190,10
Sat_Kick.kick 180+(Rnd*10),6+(Rnd*4)
Controller.Switch(23) = 0
Sat_Throw.enabled = 1
Ball.Visible = 0
smagball = 0
smagon = 0
End If
End If
End Sub


'Logic for ball behaviour on Satellite.
Sub Sat_Kick_Hit
PlaySoundAt "target",MiddleWheel
If smagball = 1 Then
me.destroyball
ball.visible = 0
Sat_Kick.createball
'Sat_Kick.kick 190,10
Sat_Kick.kick 180+(Rnd*10),6+(Rnd*4)
Controller.Switch(23) = 0
Sat_Throw.enabled = 1
Else
If smagon = 1 AND smagball = 0 Then
smagball = 1
me.destroyball
Ball2.Visible = 1
shakeball = 1
sball_shake.enabled = 1
Controller.Switch(23) = 1
Else
If smagball = 0 AND smagon = 0 Then
vpmTimer.PulseSw 23
me.destroyball
Sat_Kick.createball
'Sat_Kick.kick 190,10
Sat_Kick.kick 180+(Rnd*10),6+(Rnd*4)
smagball = 0
Sat_Throw.enabled = 1
End If
End If
End If
End Sub

'Satellite Prim Ball Shake
Dim shakeball
shakeball = 1
Sub SBall_Shake_Timer()
Select Case shakeball
Case 1:
If Ball2.ObjRotZ <= -10 Then
shakeball = 2
End If
Ball2.ObjRotZ = Ball2.ObjRotZ - 1
Case 2:
If Ball2.ObjRotZ => 0 Then
shakeball = 3
End If
Ball2.ObjRotZ = Ball2.ObjRotZ + 1
Case 3:
If Ball2.ObjRotZ => 10 Then
shakeball = 4
End If
Ball2.ObjRotZ = Ball2.ObjRotZ + 1
Case 4:
If Ball2.ObjRotX <= -6 Then
shakeball = 5
End If
Ball2.ObjRotX = Ball2.ObjRotX - 1
Case 5:
If Ball2.ObjRotX => 2 Then
shakeball = 6
End If
Ball2.ObjRotX = Ball2.ObjRotX + 1
Case 6:
If Ball2.ObjRotX = 0 Then
Ball2.ObjRotX = 0
shakeball = 7
Else
If Ball2.ObjRotX < 0 Then
Ball2.ObjRotX = Ball2.ObjRotX + 1
Else
If Ball2.ObjRotX > 0 Then
Ball2.ObjRotX = Ball2.ObjRotX - 1
End If
End If
End If
Case 7:
If Ball2.ObjRotZ = -4 Then
Ball2.ObjRotZ = -4
Ball2.visible = 0
If smagon = 1 Then
Ball.visible = 1
End If
me.enabled = 0
Else
If Ball2.ObjRotZ < -4 Then
Ball2.ObjRotZ = Ball2.ObjRotZ + 1
Else
If Ball2.ObjRotZ > -4 Then
Ball2.ObjRotZ = Ball2.ObjRotZ - 1
End If
End If
End If
End Select
End Sub

Sub Trigger005_Hit()
Sat_Throw.enabled = 1
End Sub

Sub Sat_Throw_Timer()
'PlaySound "ball_bounce"
me.Enabled = 0
End Sub

'Satellite Animation Code / home switch position.
Dim sdir
sdir = 1
Sub Sat_Timer()
Select Case sdir
Case 1:
If MiddleWheel.RotY => 30 Then
MiddleWheel.RotY = 30
Ball.ObjRotZ = 17
sdir = 2
End If
MiddleWheel.Roty = MiddleWheel.RotY + 0.5
MiddleWheel001.Roty = MiddleWheel001.RotY + 0.5
Ball.ObjRotZ = Ball.ObjRotz + 0.35
Sat_Back.ObjRotZ = Sat_Back.ObjRotZ + 0.55
'Ball001.ObjRotZ = Ball001.ObjRotz + 0.28
Case 2:
If MiddleWheel.RotY <= -20 Then
MiddleWheel.RotY = -20
Ball.ObjRotZ = -18
sdir = 1
End If
MiddleWheel.Roty = MiddleWheel.RotY - 0.5
MiddleWheel001.Roty = MiddleWheel001.RotY - 0.5
Ball.ObjRotZ = Ball.ObjRotz - 0.35
'Ball001.ObjRotZ = Ball001.ObjRotz - 0.28
Sat_Back.ObjRotZ = Sat_Back.ObjRotZ - 0.55
End Select
If Ball.ObjRotZ > -1 AND Ball.ObjRotZ < 8 Then
Controller.Switch(20)=1
Else
Controller.Switch(20)=0
End If
End Sub


'Satellite Motor Relay
Sub SolSatMotorRelay(enabled)   'rotate satelite
  if enabled then
  sat.enabled = 1
Else
    sat.enabled = 0
  End If
End Sub

'Flipper Subs
Sub SolLFlipper(Enabled)
     If Enabled Then
     'PlaySound SoundFX("Flipperup",DOFFlippers):LeftFlipper.RotateToEnd
     PlaySoundatvol "flipperup",leftflipper, 0.5
     LeftFlipper.RotateToEnd
     DOF 102,1
     Else
     PlaySoundatvol "flipperdown",leftflipper, 0.5
     'PlaySound SoundFX("Flipperdown",DOFFlippers):LeftFlipper.RotateToStart
   LeftFlipper.RotateToStart
     DOF 102,0
     End If
  End Sub

Sub SolRFlipper(Enabled)
     If Enabled Then
     'PlaySound SoundFX("Flipperup",DOFFlippers):RightFlipper.RotateToEnd
   PlaySoundatvol "flipperup",rightflipper, 0.5
     RightFlipper.RotateToEnd
     DOF 101,1
     Else
     PlaySoundatvol "flipperdown",rightflipper, 0.5
     'PlaySound SoundFX("Flipperdown",DOFFlippers):RightFlipper.RotateToStart
   RightFlipper.RotateToStart
     DOF 101,0
     End If
End Sub


'Core Timer for Subs that need refresh.
Sub SS19_Hit
vpmTimer.PulseSw 19
End Sub

Sub Chop_Trig_Hit()
timer1.enabled = 1
End Sub

Dim rcount
Sub timer1_timer()
If rcount => 2000 Then
rcount = 0
me.enabled = 0
End If
If rcount <= 500 Then
me.interval = 2
Else
If rcount > 500 AND rcount < 1500 Then
me.interval = 1
Else
If rcount > 1500 Then
me.interval = 2
End If
End If
End If
rcount = rcount + 1
rotornew.ObjRotZ = rotornew.ObjRotZ + 1
rotornew001.ObjRotZ = rotornew001.ObjRotZ + 1
End Sub


'Pulse Switch 15 on trough upkick.
Sub SolLockOut(Enabled)
 If Enabled Then vpmTimer.PulseSw 15
    PlaySoundAt "fx_solenoid", drain
End Sub


Sub GoldenEye_Exit  '  in some tables this needs to be Table1_Exit
    Controller.Stop
End Sub


Sub GoldenEye_KeyDown(ByVal Keycode)

    If Keycode = 3 Then
    'SatMag.MagnetOn = 0
    'shakeball = 1:sball_shake.enabled = 1
    'Sat_Ramp.heighttop = 0:Sat_Ramp.collidable = 0
    'Sat_Ramp_Wall.isdropped = 1
    'ball.visible = 0
    'Sat_Kick.createball
    'Sat_Kick.kick 190,10
    'Controller.Switch(23) = 0
    'mag_stub.enabled = 1
    End If

  If KeyCode=StartGameKey Then
    Controller.Switch(3)=1
    Exit Sub
  End If
  If KeyCode=KeySlamDoorHit Then
    Controller.Switch(7)=1
    Exit Sub
  End If
  If KeyCode=PlungerKey or KeyCode=LockBarKey Then Controller.Switch(9)=1
 If vpmKeyDown(KeyCode) Then Exit Sub
End Sub

Sub GoldenEye_KeyUp(ByVal KeyCode)
 If KeyCode=StartGameKey Then
    Controller.Switch(3)=0
    Exit Sub
  End If
  If KeyCode=KeySlamDoorHit Then
    Controller.Switch(7)=0
    Exit Sub
  End If
  If KeyCode=PlungerKey or KeyCode=LockBarKey Then Controller.Switch(9)=0
 If vpmKeyUp(KeyCode) Then Exit Sub
End Sub

Dim bsTrough,bsScoop,bsTank,SatMech,bsPlunger,bsLockOut,BallInSat, SatMag, FlipMag
BallInSat=0

Sub GoldenEye
  vpmInit Me
  On Error Resume Next
    With Controller
     .GameName=cGameName
             .Games(cGameName).Settings.Value("rol")=0
       .Games(cGameName).Settings.Value("sound")=1
      If Err Then MsgBox "Can't start Game" & cGameName & vbNewLine & Err.Description : Exit Sub
      .SplashInfoLine="007 Goldeneye - Sega, 1996"&vbNewLine&"Initial Table Releace by Pingod & Kid Charlemagne"&vbNewLine&"Mod by UncleReamus & Destruk"&vbNewLine&"Further Modding by The Trout"
      .HandleMechanics=0
      .HandleKeyboard=0
     .ShowDMDOnly=1
      .ShowFrame=0
      .ShowTitle=0
      '.Run
     'End With
     'If Err Then MsgBox Err.Description
         'On Error Goto 0
            On Error Resume Next
            .Run GetPlayerHWnd
            If Err Then MsgBox Err.Description
            On Error Goto 0
     End With
     On Error Goto 0

  vpmNudge.TiltSwitch=1
 vpmNudge.Sensitivity=5
  vpmNudge.TiltObj=Array(LeftSlingshot,RightSlingshot,LeftTurboBumper,RightTurboBumper)


    Set FlipMag = New cvpmMagnet
    With FlipMag
        .InitMagnet Trigger002, 90
        .GrabCenter = True
        .CreateEvents "FlipMag"
    End With

    Set bsTrough=New cvpmBallStack
    bsTrough.InitSw 0,14,13,12,11,10,0,0
    bsTrough.InitKick BallRelease,40,8
    'bsTrough.InitExitSnd "BallRelease","Solon"
        bsTrough.InitExitSnd SoundFX("ballrelease", DOFContactors), SoundFX("fx_Solenoid", DOFContactors)
   bsTrough.Balls=5

  'Set bsScoop=New cvpmBallStack
  ' bsScoop.InitSw 0,50,0,0,0,0,0,0
 ' bsScoop.InitKick Scoop,213,10
 ' bsScoop.InitExitSnd "scoopexit","Solon"
    '   bsScoop.InitExitSnd SoundFX("scoopexit", DOFContactors), SoundFX("fx_Solenoid", DOFContactors)

    Set bsScoop = New cvpmBallStack
    bsScoop.InitSaucer Scoop, 50,213,10
    'bsLeft.InitExitSnd BallRel,SolOn
        bsScoop.InitExitSnd SoundFX("scoopexit", DOFContactors), SoundFX("fx_Solenoid", DOFContactors)

  Set bsTank=New cvpmBallStack
    bsTank.InitSw 0,56,0,0,0,0,0,0
    bsTank.InitKick TankKickBig,0,60
    bsTank.InitExitSnd "popper","fx_solenoid"
   bsTank.KickBalls=2
    bsTank.KickForceVar=4

 Set bsPlunger=New cvpmBallStack
   bsPlunger.InitSaucer Plunger,16,0,55
    bsPlunger.InitExitSnd "fx_solenoid","SolOff"
    bsPlunger.KickForceVar=6

  vpmCreateEvents AllSwitches
 vpmMapLights AllLights

Sat_Ramp_Wall.isdropped = 1
TL_Wall.isdropped = 1
Controller.Switch(20)=1
'UDRamp001.Collidable = 1
'FlipMag.MagnetOn = 1
End Sub

'Additional Solenoid Subs

Sub LeftTurboBumper_Hit: vpmTimer.PulseSw 41:PlaySoundAtVol "bumper_slap", Primitive73, 0.03:End Sub
Sub BottomTurboBumper_Hit: vpmTimer.PulseSw 42:PlaySoundAtVol "bumper_slap", Primitive004, 0.03:End Sub
Sub RightTurboBumper_Hit: vpmTimer.PulseSw 43:PlaySoundAtVol "bumper_slap", Primitive003, 0.03:End Sub

'**********Sling Shot Animations
' Rstep and Lstep  are the variables that increment the animation
'****************
Dim RStep, Lstep

Sub RightSlingShot_Slingshot
  vpmTimer.PulseSw 62
    PlaySoundAt "slingshot_akiles", Sling1
    'PlaySound SoundFX("slingshot_akiles",DOFContactors), 0,1, 0.05,0.05,0,0,1
    RSling.Visible = 0
    RSling1.Visible = 1
    sling1.rotx = 20
    RStep = 0
    RightSlingShot.TimerEnabled = 1
  RightSlingShot.TimerInterval = 10
End Sub

Sub RightSlingShot_Timer
    Select Case RStep
        Case 3:RSLing1.Visible = 0:RSLing2.Visible = 1:sling1.rotx = 10
        Case 4:RSLing2.Visible = 0:RSLing.Visible = 1:sling1.rotx = 0:RightSlingShot.TimerEnabled = 0
    End Select
    RStep = RStep + 1
End Sub

Sub LeftSlingShot_Slingshot
  vpmTimer.PulseSw 61
  PlaySoundAt "slingshot_akiles", Sling2
    'PlaySound SoundFX("slingshot_akiles",DOFContactors), 0,1, -0.05,0.05,0,0,1
    LSling.Visible = 0
    LSling1.Visible = 1
    sling2.rotx = 20
    LStep = 0
    LeftSlingShot.TimerEnabled = 1
  LeftSlingShot.TimerInterval = 10
End Sub

Sub LeftSlingShot_Timer
    Select Case LStep
        Case 3:LSLing1.Visible = 0:LSLing2.Visible = 1:sling2.rotx = 10
        Case 4:LSLing2.Visible = 0:LSLing.Visible = 1:sling2.rotx = 0:LeftSlingShot.TimerEnabled = 0
    End Select
    LStep = LStep + 1
End Sub

Sub Drain_Hit:PlaySoundAtBall "drain5":bsTrough.AddBall Me:End Sub
Sub Scoop_Hit:PlaySoundAtVol "scoopenter", PegPlasticT003, 0.25:bsScoop.AddBall Me:End Sub
Sub TankKickBig_Hit:bsTank.AddBall Me:End Sub
Sub DropRamp1_Hit:PlaySoundAt "popper_ball", Tank:bsTank.AddBall Me:End Sub
Sub Plunger_Hit:bsPlunger.Addball 0:End Sub
Sub Kicker1_Hit:bsLockOut.AddBall 0:End Sub


'Satellite Ramp (Add animation when you CBF)
Sub SolSatLaunchRamp(Enabled)
 If Enabled then
'Sat_Ramp.heighttop = 70
satrampdir = 2:SatUD.enabled = 1
Sat_Ramp001.collidable = 1
Sat_Ramp_Wall.isdropped = 0
PlaySoundAt "SolOn", MiddleWheel001
DOF 131,2
 Else
'Sat_Ramp.heighttop = 0
satrampdir = 1:SatUD.enabled = 1
Sat_Ramp001.collidable = 0
Sat_Ramp_Wall.isdropped = 1
PlaySoundAt "SolOff", MiddleWheel001
DOF 131,2
    End If
End Sub

Dim satrampdir

Sub SatUD_Timer
Select Case satrampdir
Case 1:
If Sat_Ramp.HeightTop <= 0 Then
Sat_Ramp.HeightTop = 0
me.enabled = 0
End If
Sat_Ramp.HeightTop = Sat_Ramp.HeightTop - 1
Case 2:
If Sat_Ramp.HeightTop => 70 Then
Sat_Ramp.HeightTop = 70
me.enabled = 0
End If
Sat_Ramp.HeightTop = Sat_Ramp.HeightTop + 1
End Select
End Sub

'Shooterlane Top Drop Ramp - Weird design (Add granular animation when you CBF)

Sub SolUpDownRamp(Enabled)
If Enabled Then
 'UDRamp.HeightTop = 0
    udrampdir = 1:RampUD.enabled = 1
    UDRamp001.Collidable = 1
    UDRamp.Collidable = 0
    TL_Wall.isdropped = 0
    PlaySoundAt "SolOn", Bolt049
    DOF 130,2
Else
 'UDRamp.HeightTop = 70
    udrampdir = 2:RampUD.enabled = 1
    UDRamp001.Collidable = 0
  UDRamp.Collidable = 1
  TL_Wall.isdropped = 1
    PlaySoundAt "SolOff", Bolt049
    DOF 130,2
End If
End Sub

Dim udrampdir

Sub RampUD_Timer
Select Case udrampdir
Case 1:
If UDRamp.HeightTop <= 0 Then
UDRamp.HeightTop = 0
me.enabled = 0
End If
UDRamp.HeightTop = UDRamp.HeightTop - 1
Case 2:
If UDRamp.HeightTop => 70 Then
UDRamp.HeightTop = 70
me.enabled = 0
End If
UDRamp.HeightTop = UDRamp.HeightTop + 1
End Select
End Sub


'Sub Trigger2_Hit()
' stopsound "fx_metalrolling"
' playsoundatBall "balldrop"
'End Sub

'Sub Trigger3_Hit()
' playsound "fx_metalrolling"
'    xtype = 0
'End Sub

'Sub Trigger4_Hit()
' playsound "fx_ballhit"
' playsound "fx_metalrolling"
'    xtype = 0
'End Sub

'Sub Trigger5_Hit()
' stopsound "fx_metalrolling"
' playsoundatBall "balldrop"
'    xtype = 0
'End Sub

'Sub Trigger001_Hit()
' xtype = 1
'End Sub

'Sub Trigger003_Hit()
' xtype = 1
'End Sub

'Sub Trigger004_Hit()
' xtype = 1
'End Sub

Sub aTargets_hit(idx)
PlaySoundAtBall "target"
End Sub

Sub aRollOvers_hit(idx)
PlaySoundAtBall "fx_sensor"
End Sub

Sub aGates_hit(idx)
PlaySoundAtBall "fx_gate"
End Sub
'*********************************************************************
'                     Supporting Ball & Sound Functions
'*********************************************************************

Function AudioFade(tableobj) ' Fades between front and back of the table (for surround systems or 2x2 speakers, etc), depending on the Y position on the table. "goldeneye" is the name of the table
  Dim tmp
    tmp = tableobj.y * 2 / goldeneye.height-1
    If tmp > 0 Then
    AudioFade = Csng(tmp ^10)
    Else
        AudioFade = Csng(-((- tmp) ^10) )
    End If
End Function

Function AudioPan(tableobj) ' Calculates the pan for a tableobj based on the X position on the table. "goldeneye" is the name of the table
    Dim tmp
    tmp = tableobj.x * 2 / goldeneye.width-1
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

Function Pan(ball) ' Calculates the pan for a ball based on the X position on the table. "CP" is the name of the table
    Dim tmp
    tmp = ball.x * 2 / goldeneye.width-1
    If tmp > 0 Then
        Pan = Csng(tmp ^10)
    Else
        Pan = Csng(-((- tmp) ^10) )
    End If
End Function

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

Sub PlaySoundAtVol(soundname, tableobj, vol)
    PlaySound soundname, 1, (vol), AudioPan(tableobj), 0,0,0, 1, AudioFade(tableobj)
End Sub

Sub PlaySoundAtBall(soundname)
    PlaySoundAt soundname, ActiveBall
End Sub



'**********************
' Ball Collision Sound
'**********************

Sub OnBallBallCollision(ball1, ball2, velocity)
  PlaySound("fx_collide"), 0, Csng(velocity) ^2 / 2000, AudioPan(ball1), 0, Pitch(ball1), 0, 0, AudioFade(ball1)
End Sub


'*****************************************
'      Other Table Sounds
'*****************************************

Sub aRubbers_Hit(idx)
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

Sub RightFlipper2_Collide(parm)
  RandomSoundFlipper()
End Sub

Sub RandomSoundFlipper()
  Select Case Int(Rnd*3)+1
    Case 1 : PlaySound "flip_hit_1", 0, Vol(ActiveBall)*3, AudioPan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
    Case 2 : PlaySound "flip_hit_2", 0, Vol(ActiveBall)*3, AudioPan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
    Case 3 : PlaySound "flip_hit_3", 0, Vol(ActiveBall)*3, AudioPan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
  End Select
End Sub

Sub Dampen(dt,df,r)           'dt is threshold speed, df is dampen factor 0 to 1 (higher more dampening), r is randomness
  Dim dfRandomness
  r=cint(r)
  dfRandomness=INT(RND*(2*r+1))
  df=df+(r-dfRandomness)*.01
  If ABS(activeball.velx) > dt Then activeball.velx=activeball.velx*(1-df*(ABS(activeball.velx)/100))
  If ABS(activeball.vely) > dt Then activeball.vely=activeball.vely*(1-df*(ABS(activeball.vely)/100))
End Sub

'*****************************************
'      JP's VP10 Rolling Sounds
'*****************************************

'Const tnob = 5 ' total number of balls in this table is 4, but always use a higher number here because of the timing

Const tnob = 5 'ninuzzu - why 5 balls? Bad Cats has only one ball
ReDim rolling(tnob)
InitRolling

Sub InitRolling
    Dim i
    For i = 0 to tnob
        rolling(i) = False
    Next
End Sub

Sub RollingTimer_Timer()
    Dim BOT, b, ballpitch
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
            If BOT(b).z < 30 Then
                ballpitch = Pitch(BOT(b) )
            Else
                ballpitch = Pitch(BOT(b) ) * 100
            End If
            rolling(b) = True
            PlaySound("fx_ballrolling" & b), -1, Vol(BOT(b) ), Pan(BOT(b) ), 0, ballpitch, 1, 0
        Else
            If rolling(b) = True Then
                StopSound("fx_ballrolling" & b)
                rolling(b) = False
            End If
        End If

     ' rothbauerw's Dropping Sounds
        If BOT(b).VelZ < -1 and BOT(b).z < 55 and BOT(b).z > 27 Then 'height adjust for ball drop sounds
            PlaySound "fx_balldrop", 0, ABS(BOT(b).velz)/17, Pan(BOT(b)), 0, Pitch(BOT(b)), 1, 0, AudioFade(BOT(b))
        End If

    Next
End Sub


'*****************************************
' ninuzzu's BALL SHADOW
'*****************************************
Dim BallShadow
BallShadow = Array (BallShadow1,BallShadow2,BallShadow3,BallShadow4,BallShadow5,BallShadow6,BallShadow7)

Sub BallShadowUpdate_Timer
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
        If BOT(b).X < goldeneye.Width/2 Then
            BallShadow(b).X = ((BOT(b).X) - (Ballsize/6) + ((BOT(b).X - (goldeneye.Width/2))/21)) + 6
        Else
            BallShadow(b).X = ((BOT(b).X) + (Ballsize/6) + ((BOT(b).X - (goldeneye.Width/2))/21)) - 6
        End If
        ballShadow(b).Y = BOT(b).Y + 4
        If BOT(b).Z > 20 Then
            BallShadow(b).visible = 1
        Else
            BallShadow(b).visible = 0
        End If
    Next
End Sub

' Theatre Of Magic - IPDB No. 2845
' © Bally/Midway 1995
' VPX recreation by ninuzzu & Tom Tower
' Thanks to all the authors (JPSalas, Jamin, PacDude, Fuseball, Sacha) who made this table before us.
' Thanks to Clark Kent for the pics and the advices
' Thanks to Hauntfreaks for the magnet decal
' Thanks to Flupper for retexturing and remeshing the plastic ramps
' Thanks to zany for the domes, flippers and bumpers models
' Thanks to knorr for some sound effects I borrowed from his tables
' Thanks to VPDev Team for the freaking amazing VPX

' Wob 2018-08-09
' Transfered from the Hat Mod Edits by Thalamus
' Added/Updated "Positional Sound Playback Functions" and "Supporting Ball & Sound Functions"
' Changed UseSolenoids=1 to 2
' Added vpmInit Me to table init and Moved cSingleLFlip and /cSingleRFlip
' No special SSF tweaks yet.

Option Explicit
Randomize

Dim RollingVol, FlipperType, TDFlasher, TDFlasherColor, TrunkShake, TigerSaw, CenterPost

'************************************************************************
'* TABLE OPTIONS ********************************************************
'************************************************************************

'***********  Ball Rolling Sound Volume *****************************

RollingVol = 4000       '(decrease to higher the volume)

'***********  Set the Flippers Type *********************************

FlipperType = 0       '0 = white , 1 = yellow , 2 = random

'***********  Enable Flasher under the TrapDoor   *********************

TDFlasher = 1       '0 = no , 1 = yes
TDFlasherColor = 4      '0 = red, 1 = green, 2 = blue, 3 = yellow, 4 = purple, 5= cyan

'***********  Shake the Trunk on Hit    *****************************

TrunkShake = 1        '0 = no , 1 = yes

'***********  Rotate the TigerSaw (only in prototypes) ****************

TigerSaw = 1        '0 = no , 1 = yes

'***********  Enable Center Post (only in prototypes) *****************

CenterPost = 0      '0 = no , 1 = yes


'************************************************************************
'* END OF TABLE OPTIONS *************************************************
'************************************************************************

Dim Ballsize,BallMass
BallSize = 50
BallMass = (Ballsize^3)/125000

On Error Resume Next
ExecuteGlobal GetTextFile("controller.vbs")
If Err Then MsgBox "You need the controller.vbs in order to run this table, available in the vp10 package"
On Error Goto 0

Dim DesktopMode:DesktopMode = Table1.ShowDT
Dim UseVPMDMD:UseVPMDMD = DesktopMode

LoadVPM "01560000", "WPC.VBS", 3.50

'********************
'Standard definitions
'********************
 Const UseSolenoids = 2
' Wob: Added for Fast Flips (No upper Flippers)
 Const cSingleLFlip = 0
 Const cSingleRFlip = 0
 Const UseLamps = 0
 Const UseSync = 0
 Const HandleMech = 1

 ' Standard Sounds
 Const SSolenoidOn = "fx_solon"
 Const SSolenoidOff = ""
 Const SFlipperOn = "FlipperUp"
 Const SFlipperOff = "FlipperDown"
 Const SCoin = "fx_coin"

 Set GiCallback2 = GetRef("UpdateGI")

 Dim LeftMagnet, RightMagnet, cbCaptive

'******************************************************
'           TABLE INIT
'******************************************************

 Const cGameName = "tom_14hb" '1.3x arcade rom - with credits

 Sub Table1_Init
     vpmInit Me
     With Controller
         .GameName = cGameName
         .SplashInfoLine = "Theatre of Magic - Bally 1995"
         .HandleKeyboard = 0
         .ShowTitle = 0
         .ShowDMDOnly = 1
         .ShowFrame = 0
         .HandleMechanics = 1
         .Hidden = DesktopMode
         On Error Resume Next
         .Run GetPlayerHWnd
         If Err Then MsgBox Err.Description
         On Error Goto 0
    .Switch(22) = 1     'close coin door
    .Switch(24) = 1     'and keep it close
     End With

     ' Nudging
     vpmNudge.TiltSwitch = 14
     vpmNudge.Sensitivity = 3
     vpmNudge.TiltObj = Array(Bumper1, Bumper2, Bumper3, LeftSlingshot, RightSlingshot)

     ' Left Magnet
     Set LeftMagnet = New cvpmMagnet
     With LeftMagnet
         .InitMagnet LMagnet, 20
         .CreateEvents "LeftMagnet"
         .GrabCenter = 1
     End With

     ' Right Magnet
     Set RightMagnet = New cvpmMagnet
     With RightMagnet
         .InitMagnet RMagnet, 20
         .CreateEvents "RightMagnet"
         .GrabCenter = 1
     End With

' ' Trunk Mech
' Set mechTrunk = New cvpmMech
' With mechTrunk
'   .MType = vpmMechTwoDirSol+vpmMechLinear
'   .Sol1 = 17
'   .Sol2 = 18
'   .Length = 120
'   .Steps = 271
'   .AddSw 56, 1, 271
'   .AddSw 55, 0, 89
'   .AddSw 55, 91, 271
'   .AddSw 58, 0, 179
'   .AddSw 58, 181, 271
'   .AddSw 57, 0, 267
'        .CallBack = GetRef("UpdateTrunk")
'        .Start
' End With

     ' Captive Balls
  Set cbCaptive = New cvpmCaptiveBall
  With cbCaptive
    .InitCaptive CapTrigger, CapWall, Array(CapKicker, CapKicker1),350
    .NailedBalls = 1
    .Start
    CapKicker1.CreateSizedBallWithMass (Ballsize/2,BallMass).ID=5
    .ForceTrans = 0.8
    .MinForce = 3.5
    .RestSwitch = 52
    .CreateEvents "cbCaptive"
  End With

     ' Main Timer init
     PinMAMETimer.Interval = PinMAMEInterval
     PinMAMETimer.Enabled = 1

'************  Trough init
  sw32.CreateSizedBallWithMass Ballsize/2, BallMass
  sw33.CreateSizedBallWithMass Ballsize/2, BallMass
  sw34.CreateSizedBallWithMass Ballsize/2, BallMass
  sw35.CreateSizedBallWithMass Ballsize/2, BallMass
  Controller.Switch(35) = 1
  Controller.Switch(34) = 1
  Controller.Switch(33) = 1
  Controller.Switch(32) = 1

'************  Mirrored Primitives and lights init
  InitMirror sw66p, sw66m
  InitMirror sw67p, sw67m
  InitMirror LLoopGateP, LLoopGatePm
  InitMirror RLoopGateP, RLoopGatePm
  InitMirror LLoopGateSw, LLoopGateSwm
  InitMirror RLoopGateSw, RLoopGateSwm
  InitMirror Wiregate, Wiregatem
  InitMirror Prim_Lane1, Prim_Lane1m
  InitMirror Prim_Lane2, Prim_Lane2m
  InitMirror Prim_Lane3, Prim_Lane3m
  InitMirror PegMetal10, PegMetal10m
  InitMirror PegPlastic19, PegPlastic19m
  InitMirror PegPlastic20, PegPlastic20m
  InitMirror PegPlastic21, PegPlastic21m
  InitMirror PegPlastic22, PegPlastic22m
  InitMirror PegPlastic23, PegPlastic23m
  InitMirror bracket14, bracket14m
  InitMirror Screw3, Screw3m
  InitMirror Screw4, Screw4m
  InitMirror Screw5, Screw5m
  InitMirror Screw6, Screw6m
  InitMirror Screw7, Screw7m
  InitMirror Screw8, Screw8m
  InitMirror Plastics10, Plastics10m
  InitMirror Plastics13, Plastics13m
  InitMirror Plastics13, Plastics13m
  InitMirror BS1, BS4
  InitMirror BB1, BB4
  InitMirror BR1, BR4
  InitMirror BC1, BC4
  InitMirror Rub7, Rub7m
  InitMirror Rub8, Rub8m
  InitMirror Rub9, Rub9m
  InitMirror Rub10, Rub10m
  InitMirror Rub11, Rub11m
  InitMirror Rub12, Rub12m
  InitMirror Rub13, Rub13m
  InitMLights l37a, - 345.9, 1
  InitMLights l38a, - 345.9, 1
  InitMLights GIa, - 363, 58
  InitMLights GIa1, - 258, 58
  InitMLights GIa2, - 263, 58
  InitMLights GIa3, - 240.5, 58
  InitMLights GIa4, - 223, 58
  GIa5.visible = NOT DesktopMode
  GIa6.visible = NOT DesktopMode
  RLoopGatePm.visible = DesktopMode
  RLoopGateSwm.visible = DesktopMode
  RailL.visible = DesktopMode
  RailR.visible = DesktopMode

'************  Misc Stuff init
  TrapDoor.IsDropped = 1
  DivPost.IsDropped = 1
  LoopPost.IsDropped = 1
  sw85.IsDropped = 0      ' Default start position
  TrunkBlock.IsDropped = 1
  CPC.IsDropped = 1
 End Sub


'******************************************************
'             KEYS
'******************************************************

Sub Table1_KeyDown(ByVal Keycode)
  If Keycode = KeyFront Then Controller.Switch(23) = 1
  If keycode = PlungerKey Then Plunger.Pullback:Playsound "fx_plungerpull"
  If keycode = LeftTiltKey Then Nudge 90, 4:PlaySound SoundFX("fx_nudge",0)
  If keycode = RightTiltKey Then Nudge 270, 4:PlaySound SoundFX("fx_nudge",0)
  If keycode = CenterTiltKey Then Nudge 0, 5:PlaySound SoundFX("fx_nudge",0)
  If vpmKeyDown(keycode) Then Exit Sub
End Sub

Sub Table1_KeyUp(ByVal Keycode)
  If Keycode = KeyFront Then Controller.Switch(23) = 0
  If keycode = PlungerKey Then Plunger.Fire:Playsound "fx_plunger"
  If vpmKeyUp(keycode) Then Exit Sub
End Sub

Sub Table1_Paused: Controller.Pause = 1:End Sub
Sub Table1_UnPaused: Controller.Pause = 0:End Sub
Sub Table1_Exit(): Controller.Stop:End Sub

'******************************************************
'           SWITCHES
'******************************************************

'**********Rollovers
 Sub Switches_Hit(idx)
  Controller.Switch(Switches(idx).timerinterval)=1
  If Switches(idx).timerinterval = 66 Then sw66m.transZ = -30 * dCos(mAngle) : sw66m.transY =  -30 * dSin(mAngle)
  If Switches(idx).timerinterval = 67 Then sw67m.transZ = -30 * dCos(mAngle): sw66m.transY = - 30 * dSin(mAngle)
  If Switches(idx).timerinterval = 77 Then If TigerSaw = 0 Then RotateSaw : End If
  If Switches(idx).timerinterval = 78 Then If ActiveBall.VelY < -25 Then ActiveBall.VelY = -25 : End If
  If Switches(idx).timerinterval = 81 Then If ActiveBall.VelY < -15 Then ActiveBall.VelY = -15 : End If
 End Sub

 Sub Switches_UnHit(idx)
  Controller.Switch(Switches(idx).timerinterval)=0
  If Switches(idx).timerinterval = 66 Then sw66m.transZ = 0 : sw66m.transY = 0
  If Switches(idx).timerinterval = 67 Then sw67m.transZ = 0 : sw67m.transY = 0
 End Sub

'**********Ramp Switches
 Sub RampSwitches_Hit(idx)
  vpmTimer.PulseSw (RampSwitches(idx).timerinterval)
  If RampSwitches(idx).timerinterval = 73 Then sw73flip.RotatetoEnd
  If RampSwitches(idx).timerinterval = 74 Then sw74flip.RotatetoEnd
  If RampSwitches(idx).timerinterval = 75 Then If ActiveBall.VelY < 0 Then Playsound "fx_rlenter" : Else StopSound "fx_rlenter" : End If
  If RampSwitches(idx).timerinterval = 76 Then If ActiveBall.VelY < 0 Then Playsound "fx_rlenter" : Else StopSound "fx_rlenter" : End If
  If RampSwitches(idx).timerinterval = 87 Then sw87flip.RotatetoEnd
 End Sub

 Sub RampSwitches_UnHit(idx)
  If RampSwitches(idx).timerinterval = 73 Then sw73flip.RotatetoStart
  If RampSwitches(idx).timerinterval = 74 Then sw74flip.RotatetoStart
  If RampSwitches(idx).timerinterval = 87 Then sw87flip.RotatetoStart
 End Sub
'**********Spinner
 Sub sw37_Spin():vpmTimer.PulseSw 37:PlaySound "fx_spinner":End Sub

'**********Targets
Sub sw51_Hit : vpmTimer.PulseSw 51 : PlaySound SoundFX("fx_target",DOFContactors) : End Sub
Sub sw51b_Hit: vpmTimer.PulseSw 51 : PlaySound SoundFX("fx_target",DOFContactors) : End Sub
Sub sw38_Hit : vpmTimer.PulseSw 38 : PlaySound SoundFX("fx_target",DOFContactors) : End Sub
Sub sw82_Hit : vpmTimer.PulseSw 82 : PlaySound SoundFX("fx_target",DOFContactors) : End Sub
Sub sw82b_Hit: vpmTimer.PulseSw 82 : PlaySound SoundFX("fx_target",DOFContactors) : End Sub

Sub LLoopGate_Hit() : RandomSoundMetal() : End Sub
Sub RLoopGate_Hit() : RandomSoundMetal() :End Sub

'**********Captive Balls
Sub CapTrigger_Hit : cbCaptive.TrigHit ActiveBall : End Sub
Sub Caprigger_UnHit : cbCaptive.TrigHit 0 : End Sub
Sub CapWall_Hit : cbCaptive.BallHit ActiveBall : Playsound "fx_collide" : End Sub
Sub CapKicker1_Hit : cbCaptive.BallReturn Me : End Sub

'******************************************************
'         SOLENOIDS
'******************************************************

'standard coils
SolCallback(1) = "SolRelease"             'Ball Trough
SolCallback(2) = "SolSpiritRing"            'Magnet Diverter
'SolCallback(3) = "SolTrapDoorUp"           'Trap Door Up
SolCallback(4) = "SolSubwayPopper"            'Subway Popper
SolCallback(5) = "SolRightMagnet"           'Right Drain Magnet
SolCallback(6) = "SolLoopPost"              'Center Loop Post
SolCallback(7) = "SolKnocker"             'Knocker
SolCallback(8) = "SolDivPost"             'Top Diverter Post
'SolCallback(9) = "SolLSling"             'Left Slingshot
'SolCallback(10) = "SolRSling"              'RIght Slingshot
'SolCallback(11) = "SolBottomJet"           'Bottom Bumper
'SolCallback(12) = "SolMiddleJet"           'Middle Bumper
'SolCallback(13) = "SolTopJet"              'Top Bumper
SolCallback(14) = "SolTrapDoorHold"           'Trap Door Hold
SolCallBack(15) = "vpmSolGate Gate3,0,"         'Left Up/Down Gate
SolCallback(16) = "vpmSolGate Gate1,0,"         'Right Up/Down Gate
SolCallback(17) = "SolTrunkMotor"           'Move Trunk Clockwise
SolCallback(18) = "SolTrunkMotor"           'Move Trunk Counter Clockwise
SolCallBack(20) = "SetLamp 120,"            'Return Lane Flasher (x2)
SolCallback(21) = "SolTopKickout"           'Top Kickout

SolCallback(24) = "SetLamp 124,"            'Trap Door Flasher (x2)
SolCallback(25) = "SetLamp 125,"            'Spirit Ring Flasher (x2)
SolCallback(26) = "SetLamp 126,"            'Saw Flasher (x4)
SolCallback(27) = "SetLamp 127,"            'Bumper Flasher (x4)
SolCallBack(28) = "SetLamp 128,"            'Trunk Flasher (x3)

'aux board
SolCallback(33) = "SolTrunkMagnet"            'Trunk Magnet
SolCallback(34) = "SolSubBallRelease"         'Subway BallRelease
SolCallback(35) = "SolLeftMagnet"           'Left Drain Magnet

'fliptronic board
 SolCallback(sLLFlipper) = "SolLFlipper"        'Left Flipper
 SolCallback(sLRFlipper) = "SolRFlipper"        'Right Flipper

'prototype extra coils
SolCallback(23) = "SetLamp 123,"            'Magic Post Flasher (***)
If CenterPost Then SolCallback(36) = "SolMagicPost"   'Magic Post Up/Down (***)
If TigerSaw Then SolCallBack(19) = "SolTigerSaw"    'Move the TigerSaw (***)

'******************************************************
'     TROUGH BASED ON FOZZY'S
'******************************************************

Dim BIP : BIP = 0

Sub sw35_Hit():Controller.Switch(35) = 1:UpdateTrough: End Sub
Sub sw35_UnHit():Controller.Switch(35) = 0:UpdateTrough:End Sub

Sub sw34_Hit():Controller.Switch(34) = 1:End Sub
Sub sw34_UnHit():Controller.Switch(34) = 0:UpdateTrough:End Sub

Sub sw33_Hit():Controller.Switch(33) = 1:End Sub
Sub sw33_UnHit():Controller.Switch(33) = 0:UpdateTrough:End Sub

Sub sw32_Hit():Controller.Switch(32) = 1:End Sub
Sub sw32_UnHit():Controller.Switch(32) = 0:UpdateTrough:End Sub

Sub sw31_Hit():vpmTimer.PulseSw 31 : BIP = BIP + 1:End Sub

Sub UpdateTrough()
  Subway.TimerInterval = 300
  Subway.TimerEnabled = 1
End Sub

Sub Subway_Timer()
  If sw32.BallCntOver = 0 Then sw33.kick 65, 6
  If sw33.BallCntOver = 0 Then sw34.kick 65, 6
  If sw34.BallCntOver = 0 Then sw35.Kick 65, 6
  If sw35.BallCntOver = 0 Then OutHole.kick 65,8
  If sw41.BallCntOver = 0 Then sw42.kick 120, 6
  If sw42.BallCntOver = 0 Then sw43.kick 120, 6
  If sw83.BallCntOver = 0 Then sw84.kick 280, 6
  Me.TimerEnabled = 0
End Sub

'******************************************************
'       DRAIN & RELEASE
'******************************************************

Sub OutHole_Hit()
  PlaySound "fx_drain"
  UpdateTrough
  BIP = BIP - 1
End Sub


Sub SolRelease(enabled)
  If Enabled Then
    PlaySound SoundFX("fx_ballrel",DOFContactors)
    sw32.kick 65, 8
    UpdateTrough
  End If
End Sub

'******************************************************
'         DIVERTERS & GATES
'******************************************************

Sub SolLoopPost(enabled)
  If Enabled Then
    LoopPost.IsDropped = 0
    PlaySound SoundFX("fx_PostUp",DOFContactors)
  Else
    LoopPost.IsDropped = 1
    PlaySound SoundFX("fx_PostDown",0)
  End If
End Sub

Sub SolDivPost(enabled)
  If Enabled Then
    DivPost.IsDropped = 0
    PlaySound SoundFX("fx_PostUp",DOFContactors)
  Else
    DivPost.IsDropped = 1
    PlaySound SoundFX("fx_PostDown",0)
  End If
End Sub

 Sub Gate1_Hit():PlaySound "fx_gate":End Sub
 Sub Gate2_Hit():PlaySound "fx_gate":End Sub
 Sub Gate3_Hit():PlaySound "fx_gate":End Sub
 Sub Gate5_Hit():PlaySound "fx_gate":End Sub

'******************************************************
'         KNOCKER
'******************************************************

Sub SolKnocker(enabled)
  If enabled Then
    PlaySound SoundFX("fx_knocker",DOFKnocker)
  End If
End Sub

'******************************************************
'       VANISH LOCK
'******************************************************

Dim BIV: BIV = 0

Sub VanishHole_Hit:PlaySound "fx_hole2": BIV = BIV + 1 :End Sub
Sub sw84_Hit():Controller.Switch(84) = 1:UpdateTrough:End Sub
Sub sw84_UnHit():Controller.Switch(84) = 0:End Sub
Sub sw83_Hit():Controller.Switch(83) = 1:End Sub
Sub sw83_UnHit():Controller.Switch(83) = 0:End Sub

Sub SolTopKickout(Enabled)
If Enabled then
  Select Case BIV
  Case 0:                         'no ball in the vanish lock
    Playsound SoundFX(SSolenoidOn,DOFContactors)
  Case 1:                         ' only one ball in the vanish lock
    Playsound SoundFX("fx_saucer_exit",DOFContactors)
    sw83.kick 105,55
    sw84.Enabled = 0
    VanishHole.Enabled = 0
    vpmtimer.addtimer 100, "sw84.Enabled = 1'"
    vpmtimer.addtimer 300, "VanishHole.Enabled = 1'"
    TrunkMagnets 0
    BIV = BIV - 1
  Case 2:                         'two balls in the vanish lock
    Playsound SoundFX("fx_saucer_exit",DOFContactors)
    sw84.kick 105,50
    VanishHole.Enabled = 0
    vpmtimer.addtimer 300, "VanishHole.Enabled = 1'"
    TrunkMagnets 0
    BIV = BIV - 1
  End Select
End If
End Sub

'******************************************************
'         LOCK
'******************************************************

Dim BIK: BIK = 0

Sub sw44_Hit():Controller.Switch(44) = 1: BIK = BIK + 1 : End Sub
Sub sw44_UnHit():Controller.Switch(44) = 0: BIK = BIK - 1 : End Sub

Sub sw43_Hit():Controller.Switch(43) = 1:UpdateTrough:End Sub
Sub sw43_UnHit():Controller.Switch(43) = 0:UpdateTrough:End Sub

Sub sw42_Hit():Controller.Switch(42) = 1:End Sub
Sub sw42_UnHit():Controller.Switch(42) = 0:UpdateTrough:End Sub

Sub sw41_Hit():Controller.Switch(41) = 1:PlaySound "fx_kicker_catch":End Sub
Sub sw41_UnHit():Controller.Switch(41) = 0:UpdateTrough:End Sub

 Sub SolSubBallRelease(Enabled)
     If Enabled Then
    Playsound SoundFX("fx_popper",DOFContactors)
    sw41.kick 135,10
     End If
 end sub

 Sub SolSubwayPopper(Enabled)
  If Enabled Then
    sw44.kick 0, 40, 1.56
    Playsound SoundFX("fx_td-exit",DOFContactors)
  End If
 End Sub

'******************************************************
'       SLINGSHOTS
'******************************************************

Dim LStep, RStep

Sub LeftSlingshot_Slingshot
    PlaySound SoundFX("fx_slingshotL",DOFContactors)
  vpmTimer.PulseSw 61
    LSling.Visible = 0
    LSling1.Visible = 1
    sling2.TransZ = -16
    LStep = 0
    Me.TimerEnabled = 1
End Sub

Sub LeftSlingshot_Timer
    Select Case LStep
        Case 3:LSLing1.Visible = 0:LSLing2.Visible = 1:sling2.TransZ = -8
        Case 4:LSLing2.Visible = 0:LSLing.Visible = 1:sling2.TransZ = 0:Me.TimerEnabled = 0
    End Select
    LStep = LStep + 1
End Sub

Sub RightSlingshot_Slingshot
    PlaySound SoundFX("fx_slingshotR",DOFContactors)
  vpmTimer.PulseSw 62
    RSling.Visible = 0
    RSling1.Visible = 1
    sling1.TransZ = -16
    RStep = 0
    Me.TimerEnabled = 1
End Sub

Sub RightSlingshot_Timer
    Select Case RStep
        Case 3:RSLing1.Visible = 0:RSLing2.Visible = 1:sling1.TransZ = -8
        Case 4:RSLing2.Visible = 0:RSLing.Visible = 1:sling1.TransZ = 0:Me.TimerEnabled = 0
    End Select
    RStep = RStep + 1
End Sub

'******************************************************
'         BUMPERS
'******************************************************

Dim dirRing1 : dirRing1 = -1
Dim dirRing2 : dirRing2 = -1
Dim dirRing3 : dirRing3 = -1

Sub Bumper1_Hit : vpmTimer.PulseSw 65 : PlaySound SoundFX("fx_Bumper1",DOFContactors) : Me.TimerEnabled = 1 : End Sub
Sub Bumper2_Hit : vpmTimer.PulseSw 64 : PlaySound SoundFX("fx_Bumper2",DOFContactors) : Me.TimerEnabled = 1 : End Sub
Sub Bumper3_Hit : vpmTimer.PulseSw 63 : PlaySound SoundFX("fx_Bumper3",DOFContactors) : Me.TimerEnabled = 1 : End Sub

Sub Bumper1_timer()
  BR1.Z = BR1.Z + (5 * dirRing1)
  BR4.Y = mdistY - SQR((ABS(BR1.Y)-mdistY)^2 +(BR1.Z)^2) * dCos(mangle) + BR1.Z * dSin(mangle)
  BR4.Z = BR1.Z * dCos(mangle) + SQR((ABS(BR1.Y)-mdistY)^2 +(BR1.Z)^2) * dSin(mangle)
  If BR1.Z <= 0 Then dirRing1 = 1
  If BR1.Z >= 40 Then
    dirRing1 = -1
    BR1.Z = 40
    Me.TimerEnabled = 0
  End If
End Sub

Sub Bumper2_timer()
  BR2.Z = BR2.Z + (5 * dirRing2)
  If BR2.Z <= 0 Then dirRing2 = 1
  If BR2.Z >= 40 Then
    dirRing2 = -1
    BR2.Z = 40
    Me.TimerEnabled = 0
  End If
End Sub

Sub Bumper3_timer()
  BR3.Z = BR3.Z + (5 * dirRing3)
  If BR3.Z <= 0 Then dirRing3 = 1
  If BR3.Z >= 40 Then
    dirRing3 = -1
    BR3.Z = 40
    Me.TimerEnabled = 0
  End If
End Sub

'******************************************************
'       NFOZZY'S FLIPPERS
'******************************************************

Dim returnspeed, lfstep, rfstep
returnspeed = LeftFlipper.return
lfstep = 1
rfstep = 1

Sub SolLFlipper(Enabled)
    If Enabled Then
        PlaySound SoundFX("fx_flipperup",DOFContactors), 0, 1, -0.1, 0.25
        LeftFlipper.RotateToEnd
    Else
        PlaySound SoundFX("fx_flipperdown",DOFContactors), 0, 1, -0.1, 0.25
        LeftFlipper.RotateToStart
        LeftFlipper.TimerEnabled = 1
        LeftFlipper.TimerInterval = 16
        LeftFlipper.return = returnspeed * 0.5
    End If
End Sub

Sub SolRFlipper(Enabled)
    If Enabled Then
        PlaySound SoundFX("fx_flipperup",DOFContactors), 0, 1, 0.1, 0.25
        RightFlipper.RotateToEnd
    Else
        PlaySound SoundFX("fx_flipperdown",DOFContactors), 0, 1, 0.1, 0.25
        RightFlipper.RotateToStart
        RightFlipper.TimerEnabled = 1
        RightFlipper.TimerInterval = 16
        RightFlipper.return = returnspeed * 0.5
    End If
End Sub

Sub LeftFlipper_timer()
  select case lfstep
    Case 1: LeftFlipper.return = returnspeed * 0.6 :lfstep = lfstep + 1
    Case 2: LeftFlipper.return = returnspeed * 0.7 :lfstep = lfstep + 1
    Case 3: LeftFlipper.return = returnspeed * 0.8 :lfstep = lfstep + 1
    Case 4: LeftFlipper.return = returnspeed * 0.9 :lfstep = lfstep + 1
    Case 5: LeftFlipper.return = returnspeed * 1 :lfstep = lfstep + 1
    Case 6: LeftFlipper.timerenabled = 0 : lfstep = 1
  end select
end sub

Sub RightFlipper_timer()
  select case rfstep
    Case 1: RightFlipper.return = returnspeed * 0.6 :rfstep = rfstep + 1
    Case 2: RightFlipper.return = returnspeed * 0.7 :rfstep = rfstep + 1
    Case 3: RightFlipper.return = returnspeed * 0.8 :rfstep = rfstep + 1
    Case 4: RightFlipper.return = returnspeed * 0.9 :rfstep = rfstep + 1
    Case 5: RightFlipper.return = returnspeed * 1 :rfstep = rfstep + 1
    Case 6: RightFlipper.timerenabled = 0 : rfstep = 1
  end select
end sub

'******************************************************
'       MAGNET DIVERTER
'******************************************************

Sub SolSpiritRing(Enabled)
  SpiritH1.Enabled = Enabled
  If Enabled Then
    PlaySound SoundFX("fx_magnet",DOFShaker), -1, 0.1, 0, 0.25
  Else
    StopSound "fx_magnet"
    SpiritMagnet.kick 0,0,-1.56
    vpmtimer.addtimer 150, "BalldropSound"
  End If
End Sub

Dim wball,xball, yball, zball
Dim bsteps: bsteps = 0

Sub SpiritH1_Hit()
  StopSound "fx_metalrolling"
    set wball = ActiveBall
    xball = SpiritH1.X
    yball = SpiritH1.Y
    zball = 150
    Me.TimerInterval = 5
    Me.TimerEnabled = 1
End Sub

Sub SpiritH1_Timer()
  wball.x = xball
  wball.y = yball
  wball.z = zball
  xball = xball - 1
  yball = yball + 2.12
  zball = zball + 3.5
  bsteps = bsteps + 1
  If bsteps > 28 then Me.timerenabled = 0 : bsteps = 0 : Me.DestroyBall: SpiritMagnet.CreateSizedBallWithMass Ballsize/2,BallMass
End Sub

'******************************************************
'     PLAYFIELD MAGNETS
'******************************************************

Sub SolLeftMagnet(Enabled)
  If Enabled Then
    LeftMagnet.MagnetOn = 1
    PlaySound SoundFX("fx_magnet_catch",DOFShaker)
  Else
    LeftMagnet.MagnetOn = 0
  End If
End Sub

Sub SolRightMagnet(Enabled)
  If Enabled Then
    RightMagnet.MagnetOn = 1
    PlaySound SoundFX("fx_magnet_catch",DOFShaker)
  Else
    RightMagnet.MagnetOn = 0
  End If
End Sub

'******************************************************
'         TRAP DOOR
'******************************************************

Dim TDpos

Sub ShakeTrapDoor
  TDpos = 5
  TrapDoor.TimerEnabled = 1
  TrapDoor.TimerInterval = 40
End Sub

Sub TrapDoor_Timer
    TrapDoorP.transZ = TDpos
    If TDpos = 0 Then Me.TimerEnabled = 0:Exit Sub
    If TDpos < 0 Then
        TDpos = ABS(TDpos) - 1
    Else
        TDpos = - TDpos + 1
    End If
End Sub

Sub SolTrapDoorHold(Enabled)
  If Enabled then
    PlaySound SoundFX("fx_td-up",DOFContactors)
    TrapDoorP.Z = TrapDoorP.Z + 70
        ShakeTrapDoor
    TrapDoor.IsDropped = 0
    TrapDoor1.collidable= 1
    TrapDoor2.collidable= 1
    ScoopLight.state = 2
    If BIK >  0 Then
      TDHole.Enabled = 0
    Else
      TDHole.Enabled = 1
    End If
    Else
        TrapDoorP.Z = TrapDoorP.Z - 70
    TDHole.Enabled = 0
    TrapDoor.IsDropped = 1
    TrapDoor1.collidable= 0
    TrapDoor2.collidable= 0
    PlaySound "fx_td-down"
    ScoopLight.state = 0
    End If
End Sub

 Sub TDHole_Hit : PlaySound "fx_hole1" : End Sub

'******************************************************
'             MAGIC TRUNK
'******************************************************

Dim MagnetBall

Sub sw85_Hit
vpmTimer.PulseSw 85
If TrunkShake Then ShakeTrunk
RandomSoundMetal()
End Sub

Sub TrunkBlock_Hit
If TrunkShake Then ShakeTrunk
RandomSoundMetal()
End Sub

Sub TrunkOpen_Hit
If TrunkShake Then ShakeTrunk
RandomSoundMetal()
End Sub

 Sub UpdateTrunk (aNewPos, aSpeed, aOldPos)
  If aNewPos > 267 Then aNewPos = 270
  Trunk.RotZ = aNewPos
  Trunk1.RotZ = Trunk.RotZ
  Trunk2.RotZ = Trunk.RotZ
  TrunkShadow.RotZ = Trunk.RotZ
  TrunkScrews.RotZ = Trunk.Rotz
  L85.X = Trunk.x - dSin(Trunk.RotZ+175)*90
  L85.Y = Trunk.y - dCos(Trunk.RotZ-5)*90
  L85.RotY = 90 * ABS(dSin(Trunk.RotZ)) + 5
  TrunkBlock.IsDropped = Not cbool(aSpeed)      'only drop TrunkBlock if trunk is not moving
  sw85.IsDropped = NOT Controller.Switch(57)      'if Trunk is rotated to last position, then drop sw85
  If Not IsEmpty(MagnetBall) Then
    BallPos dSin(Trunk.RotZ+85)*105.7 + Trunk.x, dCos(Trunk.Rotz-95)*105.7 + Trunk.y, 103
  End If
  aOldPos = aNewPos
 End Sub

Sub solTrunkMotor(Enabled)
  If Enabled Then
    Playsound SoundFX("fx_motor",DOFGear), 0,0.5,0,0.25
    sw85.TimerEnabled = 1
    TrunkBlock.IsDropped = 0        'only drop TrunkBlock if trunk is not moving
  Else
    StopSound "fx_motor"
    sw85.TimerEnabled = 0
    TrunkBlock.IsDropped = 1        'only drop TrunkBlock if trunk is not moving
  End If
End Sub

Sub sw85_Timer()
  Trunk.RotZ = Controller.GetMech(0) -8
  Trunk1.RotZ = Trunk.RotZ
  Trunk2.RotZ = Trunk.RotZ
  TrunkShadow.RotZ = Trunk.RotZ
  TrunkScrews.RotZ = Trunk.Rotz
  L85.X = Trunk.x - dSin(Trunk.RotZ+175)*90
  L85.Y = Trunk.y - dCos(Trunk.RotZ-5)*90
  L85.RotY = 90 * ABS(dSin(Trunk.RotZ)) + 5
  sw85.IsDropped = NOT Controller.Switch(57)      'if Trunk is rotated to last position, then drop sw85
  If Not IsEmpty(MagnetBall) Then
    BallPos dSin(Trunk.RotZ+85)*105.7 + Trunk.x, dCos(Trunk.Rotz-95)*105.7 + Trunk.y, 103
  End If
End Sub

 Sub SolTrunkMagnet(aEnabled)
  If aEnabled Then
        PlaySound SoundFX("fx_magnet",DOFShaker), -1, 0.1, 0, 0.25
    TrunkMagnets 1
  Else
    If Not IsEmpty(MagnetBall) Then
      MagnetHold.kick 180,1
      MagnetBall = Empty
    End If
    StopSound "fx_magnet"
    TrunkMagnets 0
  End If
 End Sub

 Sub TrunkMagnet_Hit(idx)
  PlaySound SoundFX("fx_magnet_catch",DOFShaker)
  vpmTimer.PulseSw 85
  TrunkMagnet(idx).DestroyBall
  Set MagnetBall = MagnetHold.CreateSizedBallWithMass (Ballsize/2,BallMass)
  TrunkMagnets 0
 End Sub

 Sub TrunkMagnets(aEnabled)
  CaughtByMagnet1.enabled = aEnabled
    CaughtByMagnet2.enabled = aEnabled
    CaughtByMagnet3.enabled = aEnabled
 End Sub

 Sub BallPos(Xpos,Ypos,Zpos)
  MagnetBall.x = xpos
  MagnetBall.y = ypos
  MagnetBall.z = zpos
 End Sub

Sub sw36_Hit():vpmTimer.PulseSw 36: End Sub             ' Opto subway switch

Sub sw47_Hit():vpmTimer.PulseSw 47:End Sub

Sub TrunkEntrance1_Hit():PlaySound "fx_mine_enter":End Sub      ' Enter trunk from front
Sub TrunkEntrance2_Hit():PlaySound "fx_hole4":End Sub       ' Enter trunk from rear

'******************************************************
'* TRUNK SHAKE CODE BASED ON JP'S ********************
'******************************************************

Dim TrunkPos, tAngle

Sub ShakeTrunk
Dim finalspeed
finalspeed=INT(1 + SQR(ActiveBall.VelX * ActiveBall.VelX + ActiveBall.VelY * ActiveBall.VelY))
tAngle = dAtn(ActiveBall.VelY/ActiveBall.VelX)
If finalspeed >= 8 then
TrunkPos = 8
Else
TrunkPos = INT(9 - 8 / finalspeed)
End If
TrunkShakeTimer.Enabled = 1
End Sub

Sub TrunkShakeTimer_Timer
    Trunk.TransX = TrunkPos * dCos(tAngle)
    Trunk.TransY = TrunkPos * dSin(tAngle)
  Trunk1.TransX = Trunk.TransX
  Trunk2.TransX = Trunk.TransX
  TrunkShadow.TransX = Trunk.TransX
    TrunkScrews.TransX = Trunk.TransX
  Trunk1.TransY = Trunk.TransY
  Trunk2.TransY = Trunk.TransY
  TrunkShadow.TransY = Trunk.TransY
    TrunkScrews.TransY = Trunk.TransY
    If TrunkPos = 0 Then TrunkShakeTimer.Enabled = 0:Exit Sub
    If TrunkPos < 0 Then
        TrunkPos = ABS(TrunkPos) - 1
    Else
        TrunkPos = - TrunkPos + 1
    End If
End Sub

'******************************************************
'     TIGER SAW ANIMATION
'******************************************************

Dim tsAngle:tsAngle=0
Dim stepAngle, stopRotation

'******   VPM controlled (only in prototypes)
Sub SolTigerSaw(enabled)
  If Enabled Then
    CPC.TimerEnabled = 1
    CPC.TimerInterval = 15
    Playsound SoundFX("fx_motor",DOFGear), 0,0.25,0,0.25
    stepAngle=15
    stopRotation=0
  Else
    StopSound "fx_motor"
    stopRotation=1
  End If
End Sub

Sub CPC_Timer()
  tsAngle = tsAngle + stepAngle
  If tsAngle >= 360 Then
    tsAngle = tsAngle - 360
  End If
  Saw.rotz = tsAngle
  If stopRotation Then
    stepAngle = stepAngle - 0.1
    If stepAngle <= 0 Then
      CPC.TimerEnabled  = 0
    End If
  End If
End Sub

'******  NOT VPM controlled
Sub RotateSaw()
  Playsound SoundFX("fx_motor",0), 0,0.25,0,0.25
  RotateSawTimer.Enabled = 1
  RotateSawTimer.Interval = 20
  stepAngle=10
End Sub

Sub RotateSawTimer_Timer()
  Saw.rotz = Saw.rotz + stepAngle
  If Saw.rotz >= 3*360 Then stepAngle = stepAngle - 0.2
  If stepAngle <= 0 Then Me.Enabled = 0 : Saw.rotz = Saw.rotz -3*360 : StopSound "fx_motor"
End Sub

'******************************************************
'       GENERAL ILLUMINATION
'******************************************************

Dim gistep

Sub UpdateGI(no, step)
  'TextBox.text = no & " " &  step
    Dim xx
    If step = 0 then exit sub 'only values from 1 to 8 are visible and reliable. 0 is not reliable and 7 & 8 are the same so...
    gistep = (step-1) / 7
    Select Case no
        Case 0    ' top
            For each xx in GITop:xx.IntensityScale = gistep:next
      For each xx in GIBumpers:xx.IntensityScale = gistep:next
        Case 1    ' bottom left
            For each xx in GILeft:xx.IntensityScale = gistep:next
        Case 2    ' bottom right
            For each xx in GIRight:xx.IntensityScale = gistep:next
        Case 3    ' middle
            For each xx in GIMiddle:xx.IntensityScale = gistep:next
    End Select

  Table1.ColorGradeImage = "ColorGrade_" & step

    ' change the intensity of the flasher depending on the gi to compensate for the gi lights being off
    For xx = 0 to 200
        FlashMax(xx) = 6 - gistep * 3 ' the maximum value of the flashers
    Next
End Sub

'***************************************************
'       JP's VP10 Fading Lamps & Flashers
'       Based on PD's Fading Light System
' SetLamp 0 is Off
' SetLamp 1 is On
' fading for non opacity objects is 4 steps
'***************************************************

Dim LampState(200), FadingLevel(200)
Dim FlashSpeedUp(200), FlashSpeedDown(200), FlashMin(200), FlashMax(200), FlashLevel(200)

InitLamps()             ' turn off the lights and flashers and reset them to the default parameters
LampTimer.Interval = 10 'lamp fading speed
LampTimer.Enabled = 1

Sub LampTimer_Timer()
    Dim chgLamp, num, chg, ii
    chgLamp = Controller.ChangedLamps
    If Not IsEmpty(chgLamp) Then
        For ii = 0 To UBound(chgLamp)
            LampState(chgLamp(ii, 0) ) = chgLamp(ii, 1)       'keep the real state in an array
            FadingLevel(chgLamp(ii, 0) ) = chgLamp(ii, 1) + 4 'actual fading step
        Next
    End If
    UpdateLamps
End Sub

Sub InitLamps()
    Dim x
    For x = 0 to 200
        LampState(x) = 0         ' current light state, independent of the fading level. 0 is off and 1 is on
        FadingLevel(x) = 4       ' used to track the fading state
        FlashSpeedUp(x) = 0.5    ' faster speed when turning on the flasher
        FlashSpeedDown(x) = 0.35 ' slower speed when turning off the flasher
        FlashMax(x) = 1          ' the maximum value when on, usually 1
        FlashMin(x) = 0          ' the minimum value when off, usually 0
        FlashLevel(x) = 0        ' the intensity of the flashers, usually from 0 to 1
    Next
End Sub

 Sub UpdateLamps
  NFadeL 11,  l11
  NFadeL 12,  l12
  NFadeL 13,  l13
  NFadeL 14,  l14
  NFadeL 15,  l15
  NFadeL 16,  l16
  NFadeL 17,  l17
  NFadeL 18,  l18
  NFadeL 21,  l21
  NFadeL 22,  l22
  NFadeL 23,  l23
  NFadeL 24,  l24
  NFadeL 25,  l25
  NFadeLm 26,  l26
  NFadeL 26,  l26a
  NFadeL 27,  l27
  NFadeL 28,  l28
  NFadeL 31,  l31
  NFadeL 32,  l32
  NFadeL 33,  l33
  NFadeL 34,  l34
  NFadeL 35,  l35
  NFadeL 36,  l36
  NFadeLm 37,  l37
  Flash 37,  l37a
  NFadeLm 38,  l38
  Flash 38,  l38a
  NFadeL 41,  l41
  NFadeL 42,  l42
  NFadeL 43,  l43
  NFadeL 44,  l44
  NFadeLm 45,  l45
  NFadeL 45,  l45a
  NFadeL 46,  l46
  NFadeL 47,  l47
  NFadeL 48,  l48
  NFadeL 51,  l51
  NFadeL 52,  l52
  NFadeL 53,  l53
  NFadeLm 54,  l54
  NFadeL 54,  l54a
  NFadeL 55,  l55
  NFadeL 56,  l56
  NFadeL 57,  l57
  NFadeL 58,  l58
  NFadeL 61,  l61
  NFadeL 62,  l62
  NFadeL 63,  l63
  NFadeL 64,  l64
  NFadeL 65,  l65
  NFadeL 66,  l66
  NFadeL 67,  l67
  NFadeL 68,  l68
  NFadeL 71,  l71
  NFadeL 72,  l72
  NFadeL 73,  l73
  NFadeL 74,  l74
  NFadeL 75,  l75
  NFadeL 76,  l76
  NFadeL 77,  l77
  NFadeL 78,  l78
  NFadeLm 81,  l81
  NFadeL 81,  l81a
  Flashm 85,  l85
  Flashm 85,  l85a
  Flashm 85,  l85a1
  Flashm 85,  l85a2
  Flash 85,  l85a3
  NFadeLm 86,  l86
  FadeObj 86, CenterPostP, "3D_Centerpost_3", "3D_Centerpost_2", "3D_Centerpost_1", "3D_Centerpost"

'flashers

  NFadeLm 120,F20
  NFadeLm 120,F20a
  NFadeLm 120,F20b
  NFadeLm 120,F20c
  Flashm 120, F20d
  Flash 120, F20e

  Flash 123, F23

  NFadeLm 124,F24a
  NFadeLm 124,F24b
  NFadeLm 124,F24c
  NFadeL 124, F24d

  NFadeLm 125,F25
  NFadeLm 125,F25a
  NFadeLm 125,F25b
  NFadeLm 125, F25c
  NFadeLm 125, F25d
  NFadeLm 125, F25e
  NFadeLm 125, F25f
  Flashm 125, F25g
  Flashm 125, F25h
  Flash 125, F25i

  NFadeLm 126,F26
  NFadeLm 126,F26a
  NFadeLm 126,F26b
  NFadeLm 126, F26c
  NFadeLm 126, F26d
  Flashm 126, F26e
  Flash 126, F26f

  NFadeLm 127,F27
  NFadeLm 127,F27a
  NFadeLm 127,F27b
  NFadeLm 127, F27c
  NFadeLm 127, F27d
  NFadeLm 127, F27e
  Flash 127, F27f

  NFadeLm 128,F28
  NFadeLm 128,F28a
  NFadeLm 128,F28b
  NFadeLm 128, F28c
  NFadeLm 128, F28d
  Flashm 128, F28e
  Flashm 128, F28f
  Flash 128, F28g

  F23.Height = CenterPostP.TransZ + 1.2
 End Sub

Sub SetLamp(nr, value)
    If value <> LampState(nr) Then
        LampState(nr) = abs(value)
        FadingLevel(nr) = abs(value) + 4
    End If
End Sub

' Lights: used for VP10 standard lights, the fading is handled by VP itself

Sub NFadeL(nr, object)
    Select Case FadingLevel(nr)
        Case 4:object.state = 0:FadingLevel(nr) = 0
        Case 5:object.state = 1:FadingLevel(nr) = 1
    End Select
End Sub

Sub NFadeLm(nr, object) ' used for multiple lights
    Select Case FadingLevel(nr)
        Case 4:object.state = 0
        Case 5:object.state = 1
    End Select
End Sub

'Lights, Ramps & Primitives used as 4 step fading lights
'a,b,c,d are the images used from on to off

Sub FadeObj(nr, object, a, b, c, d)
  Select Case FadingLevel(nr)
    Case 2:object.image = d:FadingLevel(nr) = 0 'Off
    Case 3:object.image = c:FadingLevel(nr) = 2 'fading...
    Case 4:object.image = b:FadingLevel(nr) = 3 'fading...
    Case 5:object.image = a:FadingLevel(nr) = 1 'ON
 End Select
End Sub

Sub FadeObjm(nr, object, a, b, c, d)
    Select Case FadingLevel(nr)
        Case 2:object.image = d
        Case 3:object.image = c
        Case 4:object.image = b
        Case 5:object.image = d
    End Select
End Sub

Sub NFadeObj(nr, object, a, b)
    Select Case FadingLevel(nr)
        Case 4:object.image = b:FadingLevel(nr) = 0 'off
        Case 5:object.image = a:FadingLevel(nr) = 1 'on
    End Select
End Sub

Sub NFadeObjm(nr, object, a, b)
    Select Case FadingLevel(nr)
        Case 4:object.image = b
        Case 5:object.image = a
    End Select
End Sub

' Flasher objects

Sub Flash(nr, object)
    Select Case FadingLevel(nr)
        Case 4 'off
            FlashLevel(nr) = FlashLevel(nr) - FlashSpeedDown(nr)
            If FlashLevel(nr) < FlashMin(nr) Then
                FlashLevel(nr) = FlashMin(nr)
                FadingLevel(nr) = 0 'completely off
            End if
            Object.IntensityScale = FlashLevel(nr)
        Case 5 ' on
            FlashLevel(nr) = FlashLevel(nr) + FlashSpeedUp(nr)
            If FlashLevel(nr) > FlashMax(nr) Then
                FlashLevel(nr) = FlashMax(nr)
                FadingLevel(nr) = 1 'completely on
            End if
            Object.IntensityScale = FlashLevel(nr)
    End Select
End Sub

Sub Flashm(nr, object) 'multiple flashers, it just sets the flashlevel
    Object.IntensityScale = FlashLevel(nr)
End Sub

'*****************
' Maths
'*****************

Const Pi = 3.1415927

Function dSin(degrees)
  dsin = sin(degrees * Pi/180)
  if ABS(dSin) < 0.000001 Then dSin = 0
  if ABS(dSin) > 0.999999 Then dSin = 1
End Function

Function dCos(degrees)
  dcos = cos(degrees * Pi/180)
  if ABS(dCos) < 0.000001 Then dCos = 0
  if ABS(dCos) > 0.999999 Then dCos = 1
End Function

Function dAtn(degrees)
  dAtn = atn(degrees * Pi/180)
End Function

Function RndNum(min, max)
    RndNum = Int(Rnd() * (max-min + 1) ) + min ' Sets a random number between min and max
End Function


'*****************************************
' Ball Shadow
'*****************************************

Dim BallShadow
BallShadow = Array (BallShadow1, BallShadow1, BallShadow2, BallShadow3, BallShadow4, BallShadow5, BallShadow6, BallShadow7, BallShadow8)


Sub BallShadowUpdate()
    Dim BOT, b, shadowZ
    BOT = GetBalls

  ' render the shadow for each ball
    For b = 0 to UBound(BOT)
    If BOT(b).X < Table1.Width/2 Then
      BallShadow(b).X = ((BOT(b).X) - (Ballsize/6) + ((BOT(b).X - (Table1.Width/2))/7)) + 10
    Else
      BallShadow(b).X = ((BOT(b).X) + (Ballsize/6) + ((BOT(b).X - (Table1.Width/2))/7)) - 10
    End If

    If BOT(b).X > 875 AND BOT(b).Y > 935 Then shadowZ = BOT(b).Z : BallShadow(b).X = BOT(b).X Else shadowZ = 1

      BallShadow(b).Y = BOT(b).Y + 20
      BallShadow(b).Z = shadowZ
    If BOT(b).Z > 20 Then
      BallShadow(b).visible = 1
    Else
      BallShadow(b).visible = 0
    End If
  Next
End Sub

'*****************************************
' JP's VP10 Ball Collision Sound
'*****************************************

Sub OnBallBallCollision(ball1, ball2, velocity)
  PlaySound("fx_collide"), 0, Csng(velocity) ^2 / 500, Pan(ball1), 0, Pitch(ball1), 0, 0
End Sub

'*****************************************
' JF's Sound Routines
'*****************************************

Sub Rubbers_Hit(idx)
  dim finalspeed
    finalspeed=SQR(activeball.velx * activeball.velx + activeball.vely * activeball.vely)
  If finalspeed > 20 then
    PlaySound "fx_rubber", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0
  End if
  If finalspeed >= 6 AND finalspeed <= 20 then
    RandomSoundRubber()
  End If
End Sub

Sub Posts_Hit(idx)
  dim finalspeed
    finalspeed=SQR(activeball.velx * activeball.velx + activeball.vely * activeball.vely)
  If finalspeed > 16 then
    PlaySound "fx_rubber", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0
  End if
  If finalspeed >= 6 AND finalspeed <= 16 then
    RandomSoundRubber()
  End If
End Sub

Sub RandomSoundRubber()
  Select Case Int(Rnd*3)+1
    Case 1 : PlaySound "fx_rubber_hit_1", 0, 2*Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0
    Case 2 : PlaySound "fx_rubber_hit_2", 0, 2*Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0
    Case 3 : PlaySound "fx_rubber_hit_3", 0, 2*Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0
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
    Case 1 : PlaySound "fx_flip_hit_1", 0, 2*Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0
    Case 2 : PlaySound "fx_flip_hit_2", 0, 2*Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0
    Case 3 : PlaySound "fx_flip_hit_3", 0, 2*Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0
  End Select
End Sub

Sub RandomSoundMetal()
  Select Case Int(Rnd*3)+1
    Case 1 : PlaySound "fx_metal_hit_1"
    Case 2 : PlaySound "fx_metal_hit_2"
    Case 3 : PlaySound "fx_metal_hit_3"
  End Select
End Sub

'**********************
' Balldrop & Ramp Sound
'**********************

Sub BallDropSound(dummy)
    PlaySound "fx_balldrop"
End Sub

Sub WireDropSound(dummy)
    PlaySound "fx_balldrop_SSRamp"
End Sub

Sub RHelp1_Hit()
  StopSound "fx_lr7"
  vpmtimer.addtimer 200, "BallDropSound"
End Sub

Sub RHelp2_Hit()
  StopSound "fx_metalrolling"
  StopSound "fx_balldrop_SSRamp"
  vpmtimer.addtimer 200, "BallDropSound"
End Sub

Sub RHelp3_Hit()
  StopSound "fx_metalrolling"
  vpmtimer.addtimer 300, "WireDropSound"
End Sub

Sub RHelp4_Hit():vpmtimer.addtimer 200, "WireDropSound" : End Sub

Sub RHelp5_Hit()
  ActiveBall.VelZ = -1
  ActiveBall.VelY = 0
  ActiveBall.VelX = 0
  vpmtimer.addtimer 300, "BallDropSound"
End Sub

Sub RHelp6_Hit(): PlaySound "fx_metalrolling" : End Sub

Sub RHelp7_Hit()
  If ActiveBall.VelY > 0 Then PlaySound "fx_metalrolling"
End Sub

Sub RHelp7_UnHit()
  If ActiveBall.VelY < 0 Then StopSound "fx_metalrolling"
End Sub

Sub LRHit1_Hit() : PlaySound "fx_lr1" : End Sub
Sub LRHit2_Hit() : PlaySound "fx_lr2" : End Sub
Sub LRHit3_Hit() : PlaySound "fx_lr3" : End Sub
Sub LRHit4_Hit() : PlaySound "fx_lr4" : End Sub
Sub LRHit5_Hit() : PlaySound "fx_lr5" : End Sub
Sub LRHit6_Hit() : PlaySound "fx_lr6" : End Sub
Sub LRHit7_Hit() : PlaySound "fx_lr7" : End Sub

Sub RRHit1_Hit() : PlaySound "fx_lr1" : End Sub
Sub RRHit2_Hit() : PlaySound "fx_lr2" : End Sub

'******************************************************
'       Mirrored Balls (based on Koadic's )
'******************************************************

Dim mDistY:mDistY = ABS (PFmirror.Y)                ' mDistY is the the mirroring plane y component
Dim mangle                              ' mAngle angle between the mirrored pf and the pf itself

If DesktopMode Then
  mAngle= 22
Else
  mAngle= 80
End If

PFmirror.RotX = -mangle
Dim mballs : mballs = 4                   ' mballs is the maximum number of balls being reflected;
'                               ' this number must match the number of kickers defined in the array.
Dim iball, cnt, mb, MKick
MKick = Array("", MKick1, MKick2, MKick3, MKick4)     ' array of kickers that create/destroy balls
ReDim MBall(mballs),MBallStatus(mballs), BallM(mballs)
MBallStatus(0) = 0

Sub MirrorTrigger_Hit:NewMBall:End Sub            ' on hit trigger the mirrored ball, on unhit destroy the ball and disable the timer
Sub MirrorTrigger_UnHit:ClearMBall:If MBallStatus(0) = 0 Then Me.TimerEnabled=0:End If:End Sub

Sub MirrorTrigger_Timer
  If MBallStatus(0) = 0 Then Me.TimerEnabled=0:Exit Sub
  For mb = 1 to mballs
    If MBallStatus(mb) > 0 Then
      BallM(mb).x = MBall(mb).x                                               ' mirrored ball x position
      BallM(mb).y = mDistY - SQR ((MBall(mb).y - mDistY)^2 + (MBall(mb).z)^2)* dCos (mAngle) + MBall(mb).z * dSin(mAngle)     ' mirrored ball y position
      BallM(mb).z = SQR ((MBall(mb).y - mDistY)^2 + (MBall(mb).z)^2) * dSin(mAngle) + MBall(mb).z * dCos(mAngle)          ' mirrored ball z position
    End If
  Next
End Sub

Sub NewMBall                        ' this creates the mirrored ball(s)
  For cnt = 1 to mballs
    If MBallStatus(cnt) = 0 Then
      Set MBall(cnt) = ActiveBall
      MBall(cnt).uservalue = cnt
      Set BallM(cnt) = MKick(cnt).CreateSizedBallWithMass (Ballsize/2, BallMass)
      MBallStatus(cnt) = 1
      MBallStatus(0) = MBallStatus(0)+1
      Exit For
    End If
    Next
  MirrorTrigger.TimerEnabled = 1
End Sub

Sub ClearMBall                        ' this destroys the mirrored ball(s)
    iball = ActiveBall.uservalue
    MBall(iball).UserValue = 0
    MBallStatus(iBall) = 0
    MBallStatus(0) = MBallStatus(0)-1
  MKick(iball).DestroyBall
End Sub

Sub InitMirror (Prim, mPrim)
    mPrim.Y = mDistY -  SQR ((ABS(Prim.y) - mDistY)^2 + (Prim.z)^2)* dCos(mAngle) + Prim.z * dSin(mAngle)
    mPrim.Z =  SQR ((ABS(Prim.y) - mDistY)^2 + (Prim.z)^2) * dSin(mAngle) + Prim.z * dCos(mAngle)
    mPrim.objRotX = Prim.objRotX -mAngle
End Sub

Sub InitMLights(obj, ypos, height)
  obj.Y = mDistY -  SQR((ABS(ypos) - mDistY)^2 + height^2)* dCos(mAngle)
  obj.Height =  SQR((ABS(ypos) - mDistY)^2 + height^2) * dSin(mAngle) + height * dCos(mAngle)
  obj.RotX = obj.RotX - mAngle

End Sub

'******************************************************
'         RealTime Updates
'******************************************************

Set MotorCallback = GetRef("GameTimer")

Sub GameTimer
  FlipperL.RotZ = LeftFlipper.currentangle
  FlipperR.RotZ = RightFlipper.currentangle
  FlipperLSh.RotZ = LeftFlipper.currentangle
  FlipperRSh.RotZ = RightFlipper.currentangle
  WireGateL.RotX = Spinner1.currentangle
  WireGateR.RotX = Gate4.currentangle
  sw37p.RotX= -sw37.currentangle
  sw73p.RotZ= sw73flip.currentangle
  sw74p.RotZ= sw74flip.currentangle
  sw87p.RotZ= sw87flip.currentangle
  LLoopGateSw.RotY = -Gate3.currentangle
  LLoopGateSwm.RotY = -Gate3.currentangle
  RLoopGateSw.RotY = Gate1.currentangle
  RLoopGateSwm.RotY = Gate1.currentangle
  WireGate.RotX = -Gate2.currentangle
  WireGatem.RotX = -Gate2.currentangle
  RollingSoundUpdate
  BallShadowUpdate
End Sub

'******************************************************
'       Center Post (only in prototype)
'******************************************************

Sub SolMagicPost(enabled)
  If enabled then
    CenterPostP.TransZ = 26
    Playsound SoundFX("Centerpost_Up",DOFContactors)
    CPC.IsDropped = 0
  Else
    Playsound SoundFX("Centerpost_Down",DOFContactors)
    CenterPostP.TransZ = 0
    CPC.IsDropped = 1
  End If
End Sub

'*******  TABLE OPTIONS   ***********************************

CenterPostP.visible = CenterPost
CPC.collidable = CenterPost
F23.visible = CenterPost
ScoopLight.visible = TDFlasher

If FlipperType = 2 then FlipperType = RndNum(0, 1)

Select Case FlipperType
    Case 0: FlipperL.image = "flippers white"
      FlipperR.image = "flippers white"
    Case 1: FlipperL.image = "flippers"
      FlipperR.image = "flippers"
End Select

Select Case TDFlasherColor
  'Red
    Case 0 :  ScoopLight.color = RGB (255,0,0):ScoopLight.colorfull = RGB (255,170,170)
  'Green
    Case 1 :  ScoopLight.color = RGB (0,255,0):ScoopLight.colorfull = RGB (170,255,170)
  'Blue
    Case 2 :  ScoopLight.color = RGB (0,0,255):ScoopLight.colorfull = RGB (170,170,255)
  'Yellow
    Case 3 :  ScoopLight.color = RGB (255,197,143):ScoopLight.colorfull = RGB (255,255,170)
  'Magenta
    Case 4 :  ScoopLight.color = RGB (255,0,255):ScoopLight.colorfull = RGB (255,170,255)
  'Cyan
    Case 5 :  ScoopLight.color = RGB (0,255,255):ScoopLight.colorfull = RGB (170,255,255)
End Select

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
'      JP's VP10 Rolling Sounds
'*****************************************

Const tnob = 8 ' total number of balls
ReDim rolling(tnob)
InitRolling

Sub InitRolling
    Dim i
    For i = 0 to tnob
        rolling(i) = False
    Next
End Sub

Sub RollingSoundUpdate()
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
  If Table1.VersionMinor > 3 OR Table1.VersionMajor > 10 Then
    PlaySound("fx_collide"), 0, Csng(velocity) ^2 / 200, Pan(ball1), 0, Pitch(ball1), 0, 0, AudioFade(ball1)
  Else
    PlaySound("fx_collide"), 0, Csng(velocity) ^2 / 200, Pan(ball1), 0, Pitch(ball1), 0, 0
  End if
End Sub

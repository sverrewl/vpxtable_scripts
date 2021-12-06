'              ______           _
'              |  ___|         | |
'              | |_ _   _ _ __ | |__   ___  _   _ ___  ___
'              |  _| | | | '_ \| '_ \ / _ \| | | / __|/ _ \ 
'              | | | |_| | | | | | | | (_) | |_| \__ \  __/
'              \_|  \__,_|_| |_|_| |_|\___/ \__,_|___/\___|
'
'
' _    _ _ _ _ _                        __  __   _____  _____  _____  __
'| |  | (_) | (_)                      / / /  | |  _  ||  _  ||  _  | \ \ 
'| |  | |_| | |_  __ _ _ __ ___  ___  | |  `| | | |_| || |_| || |/' |  | |
'| |/\| | | | | |/ _` | '_ ` _ \/ __| | |   | | \____ |\____ ||  /| |  | |
'\  /\  / | | | | (_| | | | | | \__ \ | |  _| |_.___/ /.___/ /\ |_/ /  | |
' \/  \/|_|_|_|_|\__,_|_| |_| |_|___/ | |  \___/\____/ \____/  \___/   | |
'                                      \_\                            /_/

'Funhouse (Williams 1990) Rev1.3 for VP10 (Requires VP10.2 or greater to play)

'***The very talented FH development team.***
'Original VP10 beta by "Shoopity" and completed by "wrd1972"
'Original content borrowed from "JPsalas" VP9 table
'Additional scripting by "cyberpez", "rothbauerw", "32assassin"
'Plastics prims by "cyberpez"
'Popcorn, balloons and hotdog cart, marble ball and faceless Rudy mods by "Cyberpez"
'Creepy Rudy artwork by "Rothbauerw"
'Clear ramps by "dark"
'Wire ramps by "ninuzzu"
'Subway by "ninuzzu"
'Rudy prims by "vanlion"
'Mystery mirror by "cyberpez"
'Flasher domes by "Flupper"
'PF lighting by "wrd1972"
'Flashers and bulbs by "wrd1972"
'LFMH Physics by "wrd1972"
'PF image refresh by "clarkkent"
'DT view scoring reels and background by "32assissin"
'DOF by "arngrim"

' Thalamus 2018-07-23
' Added/Updated "Positional Sound Playback Functions" and "Supporting Ball & Sound Functions"
' Changed UseSolenoids=1 to 2
' No special SSF tweaks yet.

'***An extra special thanks to "cyberpez" and "Rothbauerw" for the hard work and countless hours of development on this table.***

Option Explicit
Randomize

'*************************************************************************************************************************
Dim OptReset
'OptReset = 1  'Uncomment to reset to default options in case of error OR keep all changes temporary. Re-comment after table start-up.
'*************************************************************************************************************************


Dim RudyMod, LazyEye, BallMod, FlipperColor, ClockMod, BalloonMod, PopCornMod, HotDogCartMod, PunchMod, TwitchMod, LevelMod, DrainPostMod, MirrorMod, SubwayColorMod, ApronMod, CardMod, MirrorRWBMod
' _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _
On Error Resume Next
ExecuteGlobal GetTextFile("controller.vbs")
If Err Then MsgBox "You need the controller.vbs in order to run this table, available in the vp10 package"
On Error Goto 0


MirrorRWBMod = 1



Const DMDRotation= 1          '0= normal,  1= rotated by 90°
'Const cGameName="fh_905"
Const cGameName="fh_905h"
Const UseSolenoids = 2
Const UseLamps = 0
Const UseGI=0
Const UseSync = 0
Const HandleMech = 0
Const SSolenoidOn="fx_solon"
Const SSolenoidOff=""
Const SCoin="coin"

Const ballsize = 25  'radius
Const ballmass = 1.7

LoadVPM "02060000", "WPC.VBS", 3.50
Dim DesktopMode: DesktopMode = Table1.ShowDT

If DesktopMode = True Then 'Show Desktop components
  leftrail.visible=1
  rightrail.visible=1
  frontlockbar.visible=1
  rearbar.visible=1
  BulbTop1DT.visible=1
  BulbTop1FS.visible=0
  BulbTop2DT.visible=1
  BulbTop2FS.visible=0
  F58_DT.visible=1
  F58a_DT.visible=1
  F58_FS.visible=0
  F58a_FS.visible=0
  F57_DT.visible=1
  F57a_DT.visible=1
  F57_FS.visible=0
  F57a_FS.visible=0
  Flasher17_FS.visible=0
  Flasher23_FS.visible=0
  Flasher24_FS.visible=0
  Flasher17_DT.visible=1
  Flasher23_DT.visible=1
  Flasher24_DT.visible=1
'    pSidewall_DT.visible=1
'    pSidewall_FS.visible=0

Else
  leftrail.visible=0
  rightrail.visible=0
  frontlockbar.visible=0
  rearbar.visible=0
  BulbTop1DT.visible=0
  BulbTop1FS.visible=1
  BulbTop2DT.visible=0
  BulbTop2FS.visible=1
  F58_DT.visible=0
  F58a_DT.visible=0
  F58_FS.visible=1
  F58a_FS.visible=1
  F57_DT.visible=0
  F57a_DT.visible=0
  F57_FS.visible=1
  F57a_FS.visible=1
  Flasher17_FS.visible=1
  Flasher23_FS.visible=1
  Flasher24_FS.visible=1
  Flasher17_DT.visible=0
  Flasher23_DT.visible=0
  Flasher24_DT.visible=0
'    pSidewall_DT.visible=0
'    pSidewall_FS.visible=1
End if

Set GiCallback2 = GetRef("UpdateGI")

'**************************************************
'     Solenoid callbacks
'**************************************************

SolCallback(1) = "kisort"                       'Out hole
SolCallback(2) = "SolRampDiverter"                    'Ramp Diverter
SolCallback(3) = "bsHideout.SolOut"                   'Rudy's Hideout kicker
SolCallback(4) = "SolKickout"                     'Main kickout
SolCallback(5) = "SolTrapDoorO"                     'Open Trap Door
SolCallback(6) = "SolTrapDoorC"                     'Close Trap Door
SolCallback(7) = "vpmSolSound SoundFX(""fx_knocker"",DOFKnocker),"    'Knocker
SolCallback(8) = "MBRelease"                      'Multi-ball release
'SolCallback(9) =                           'Left Bumper - Red
'SolCallback(10) =                            'Right Bumper - White
'SolCallback(11) =                            'Bottom Bumper - Blue
'SolCallback(12) =                            'Left Sling
'SolCallback(13) =                            'Right Sling
SolCallback(14) = "SolFlipperDiverter"                  'Steps shooter lane diverter
SolCallback(15) = "KickBallToLane"                    'Main trough kickout
SolCallback(16) = "bsRudy.SolOut"                   'Rudy's mouth kickout
'SolCallback(21) = "SolMouthMotor"                    'Rudy Mouth On/Off
'SolCallback(22) = "SolMouthUpDown"                   'Rudy Mouth Up/Down
SolCallback(25) = "SolEyesRight"                    'Rudy eyes right
SolCallback(26) = "SolEyesOpen"                     'Rudy lids open
SolCallback(27) = "SolEyesClosed"                   'Rudy lids closed
SolCallback(28) = "SolEyesLeft"                     'Rudy eyes left

'***Flashers***
SolCallback(17) = "SetBlueDome"                     'Blue Dome Flasher and X2 PF lights
SolCallback(18) = "setlamp 118,"                    'Flasher in front Rudy
SolCallback(19) = "setlamp 119,"                    'Center clock flasher
SolCallback(20) = "setlamp 120,"                    'Hot Dog Flasher
SolCallback(23) = "SetRedDome"                      'Red Dome Flasher and X2 PF lights
SolCallback(24) = "SetWhiteDome"                    'White Dome Flasher and X2 PF lights

SolCallback(sLRFlipper) = "SolRFlipper"                 'Right Flipper
SolCallback(sLLFlipper) = "SolLFlipper"                 'Left Flipper

Sub SetBlueDome(value)
  SetLamp 117, value
  if value Then
    Flasher1.image="domeblue lit"
    Flasher1.material="Flasherlit"
    Flasher1.disablelighting=1
  Else
    Flasher1.image="domeblue unlit"
    Flasher1.material="Flasherunlit"
    Flasher1.disablelighting=0
  End If
End Sub

Sub SetRedDome(value)
  SetLamp 123, value
  if value Then
    Flasher2.image="domered lit"
    Flasher2.material="Flasherlit"
    Flasher2.disablelighting=1
  Else
    Flasher2.image="domered unlit"
    Flasher2.material="Flasherunlit"
    Flasher2.disablelighting=0
  End If
End Sub

Sub SetWhiteDome(value)
  SetLamp 124, value
  if value Then
    Flasher6.image="domewhite lit"
    Flasher6.material="Flasherlit"
    Flasher6.disablelighting=1
  Else
    Flasher6.image="domewhite unlit"
    Flasher6.material="Flasherunlit"
    Flasher6.disablelighting=0
  End If
End Sub




'**************************************************
'     Initiate Table
'**************************************************

Dim bsHideout, bsRudy, rudyjawmech, MaxBalls
MaxBalls=3

Sub Table1_Init()
  vpmInit Me
  With Controller
    .GameName = cGameName
    If Err Then MsgBox "Can't start Game " & cGameName & vbNewLine & Err.Description:Exit Sub
    .SplashInfoLine = "Funhouse - Williams 1990"
    .Games(cGameName).Settings.Value("rol") = DMDRotation
    .HandleKeyboard = 0
    .ShowTitle = 0
    .ShowDMDOnly = 1
    .ShowFrame = 0
    .HandleMechanics = 0
    .Hidden = 1
    On Error Resume Next
    .Run GetPlayerHWnd
    If Err Then MsgBox Err.Description
    On Error Goto 0
  End With

  PinMAMETimer.Interval = PinMAMEInterval
  PinMAMETimer.Enabled = 1

  vpmNudge.TiltSwitch = 14
  vpmNudge.Sensitivity = 1
  vpmNudge.TiltObj=Array(LeftSlingshot,RightSlingshot,Bumper1,Bumper2,Bumper3)

  Set bsHideout = New cvpmSaucer
  With bsHideout
    .InitKicker sw46, 46, 290, 15, 0
    .InitSounds "", SoundFX("fx_solon",DOFContactors), SoundFX("fx_Popper",DOFContactors)
  End With

  Set bsRudy = New cvpmSaucer
  With bsRudy
    .InitKicker sw65, 65, 190, 12, 0
    .InitSounds "", SoundFX("fx_solon",DOFContactors), SoundFX("fx_Popper",DOFContactors)
    pRudyKick.TransY = 50
    sw65.timerenabled = true
  End With

  UpdateGI 1,1 : UpdateGI 2,1 : UpdateGI 4,1
  WaSw25.IsDropped = 1
  WaSw27.IsDropped = 1
  TrapWall.IsDropped = 1

  CheckMaxBalls 'Allow balls to be created at table start up

    Set RudyJawMech = New cvpmMech
  With RudyJawMech
    .MType = vpmMechOneDirSol + vpmMechStopEnd + vpmMechNonLinear + vpmMechFast
    .Sol1 = 21
    .Sol2 = 22
    .length = 95
    .steps = 24
    .callback = getRef("UpdateJawRudy")
    .start
  End With
End Sub

'******************************
'   Keys
'******************************

Dim BGSounds


Sub Table1_KeyDown(ByVal keycode)
  If Keycode = KeyFront Then Controller.Switch(23) = 1
  If keycode = PlungerKey Then : Plunger1.Pullback :Plunger2.Pullback : PlaySound "fx_plungerpull"
  If keycode = LeftTiltKey Then Nudge 90, 2:PlaySound SoundFX("fx_nudge",0)
  If keycode = RightTiltKey Then Nudge 270, 2:PlaySound SoundFX("fx_nudge",0)
  If keycode = CenterTiltKey Then Nudge 0, 3:PlaySound SoundFX("fx_nudge",0)
  If keycode = LeftMagnaSave then
    If BGSounds = 1 then
      StopSound "arcade"
      BGSounds = 0
    Else
      PlaySound "arcade",-1
      BGSounds = 1
    End If
  End If

  If keycode = RightMagnaSave then
    RudyType = RudyType + 1
    If RudyType = 4 then RudyType = 1
    CheckRudyType
  End If

  'uncomment and use the "L" key for testkick
  'If keycode = 38 Then testkick

'   If Keycode = LeftFlipperKey then
'   SolLFlipper true
'   Exit Sub
' End If
'   If Keycode = RightFlipperKey then
'   SolRFlipper true
'   Exit Sub
' End If

  If vpmKeyDown(keycode) Then Exit Sub
End Sub

Sub Table1_KeyUp(ByVal keycode)
  If Keycode = KeyFront Then Controller.Switch(23) = 0
  If keycode = PlungerKey Then : Plunger1.Fire :Plunger2.Fire : PlaySound "plunger2"

'   If Keycode = LeftFlipperKey then
'   SolLFlipper False
'   Exit Sub
' End If
'   If Keycode = RightFlipperKey then
'   SolRFlipper False
'   Exit Sub
' End If

  If vpmKeyUp(keycode) Then Exit Sub
End Sub

Sub testkick()
  test.CreateSizedBallWithMass BallSize, BallMass
  test.kick 80,15
End Sub

Sub table1_Paused:Controller.Pause = 1:End Sub
Sub table1_unPaused:Controller.Pause = 0:End Sub
Sub table1_exit:Controller.Stop:End Sub

'************************
'      RealTime Updates
'************************

Const PI = 3.14
Dim Gate4Angle,GateSpeed,Gate1Open,Gate1Angle,OldGameTime
Gate1Open=0:Gate1Angle=0:GateSpeed = 5

Sub Gate1_Hit():Gate1Open=1:Gate1Angle=0:PlaySound "fx_gate":End Sub

Set MotorCallback = GetRef("GameTimer")

dim defaultEOS, defaultUpperEOS,EOSAngle,EOSTorque
defaulteos = leftflipper.eostorque
defaultUpperEOS = leftflipper1.eostorque
EOSAngle = 3
EOSTorque = .9

Sub GameTimer()

  PrStepGate.ObjRotZ = StepGate2.CurrentAngle + 90
  Prim_Diverter.RotZ = RampDiv.CurrentAngle
  FlipperL.RotZ = LeftFlipper.CurrentAngle
  FlipperR.RotZ = RightFlipper.CurrentAngle
  FlipperUL.RotZ = LeftFlipper1.CurrentAngle

  If LeftFlipper.CurrentAngle < LeftFlipper.EndAngle + EOSAngle Then
    LeftFlipper.eostorque = EOSTorque
  Else
    LeftFlipper.eostorque = defaultEOS
  End If

  If RightFlipper.CurrentAngle > RightFlipper.EndAngle - EOSAngle Then
    RightFlipper.eostorque = EOSTorque
  Else
    RightFlipper.eostorque = defaultEOS
  End If

  If LeftFlipper1.CurrentAngle < LeftFlipper1.EndAngle + EOSAngle Then
    LeftFlipper1.eostorque = EOSTorque
  Else
    LeftFlipper.eostorque = defaultUpperEOS
  End If


  Gate4Angle = Int(Gate4.CurrentAngle)
  If Gate4Angle > 0 then
    pGate4_switch.ObjRotY = sin( (Gate4Angle * -1) * (2*PI/180)) * 5
  Else
    pGate4_switch.ObjRotY = sin( (Gate4Angle * 1) * (2*PI/180)) * 5
  End If

    pGate4.Rotx = Gate4.CurrentAngle' + 90

    P_gate1.RotZ = Gate.CurrentAngle +90
End Sub



'**********************************
'  Flippers
'**********************************

Sub SolLFlipper(Enabled)
     If Enabled Then
         PlaySound SoundFX("fx_Flipperup",DOFContactors):LeftFlipper.RotateToEnd:LeftFlipper1.RotateToEnd
     Else
         PlaySound SoundFX("fx_Flipperdown",DOFContactors):LeftFlipper.RotateToStart:LeftFlipper1.RotateToStart
     End If
  End Sub

Sub SolRFlipper(Enabled)
     If Enabled Then
         PlaySound SoundFX("fx_Flipperup",DOFContactors):RightFlipper.RotateToEnd
     Else
         PlaySound SoundFX("fx_Flipperdown",DOFContactors):RightFlipper.RotateToStart
     End If
End Sub

' **********************************
'       Holes & Subway
' **********************************

 Dim aBall, aZpos
 Dim bBall, bZpos

'***Drain, VUKs & Saucers***
Sub sw46_hit():bsHideOut.AddBall Me:End Sub
Sub sw65_Hit():bsRudy.AddBall Me:End Sub

Sub sw67_Hit()
  Playsound "fx_hole"
  vpmTimer.PulseSw 67
End Sub

'***Wind Tunnel Hole***
Sub subwayenter_hit()
  PlaySound "Fx_subwayenter"
End Sub

Sub sw44_hit()
  PlaySound "kicker_enter_center"
  vpmTimer.PulseSw 44
End Sub

'****Trap Door***
Sub TrigSub1_hit
Playsound "fx_subway"
End Sub

Sub TrigSub2_hit
Playsound "fx_subway"
End Sub





'**************************************
'   Tunnel Kickout
'**************************************
Sub Destroyer_hit():me.destroyball:end sub  'debug
Sw58k1.enabled = 0
Sw58k.enabled = 1

Sub SolKickout(enabled)
  If Enabled then
    sw58k1.enabled = 0  'disable second chute kicker
    vpmtimer.addtimer 10, "sw58k.kick 20,38'" '38 = strength
    sw58k1.kick 22,30 '22.5, 40
'   bsChute.SolExit true
    sw58.enabled= 0
    vpmtimer.addtimer 600, "sw58.enabled= 1'"
    Playsound SoundFx("fx_Popper",DOFContactors)
  End If
End Sub

Sub sw58_Hit()
  Playsound "fx_vuk_enter"
End Sub

Sub sw58k_Hit()
  Stopsound "fx_subway"
  'Playsound "fx_kicker_catch"
  Controller.Switch(58) = 1
  Sw58k1.enabled = 1  'enable second chute kicker
End Sub

Sub sw58k_UnHit()
  Controller.Switch(58) = 0
End Sub



'**************************************
'   Lock Mech
'**************************************

Dim lockdir

Sub MBRelease(enabled)
  If enabled then
    PlaySound SoundFX("fx_lock_exit",DOFContactors)
    WaSw28.IsDropped = 1
    flipper1.rotatetoend
    MoveLock.enabled = 1
    lockdir=-30
  End If
End Sub

Sub MoveLock_Timer()
  Lock_Release_Prim.objRotZ = Lock_Release_Prim.objRotZ + lockdir
  If Lock_Release_Prim.objRotZ <=-210 Then:lockdir = 30:flipper1.rotatetostart::end if
  If Lock_Release_Prim.objRotZ >=0 Then Lock_Release_Prim.objRotZ=0:WaSw28.IsDropped = 0:WaSw27.IsDropped = 1: Me.Enabled = 0
End Sub

Sub Sw25_Hit:Controller.Switch(25) = 1 : Playsound "fx_sensor" :End Sub
Sub Sw25_UnHit:Controller.Switch(25) = 0:End Sub
Sub Sw27_hit:Controller.Switch(27) = 1:WaSw25.IsDropped = 0 : Playsound "fx_sensor" :End Sub
Sub Sw27_unhit:Controller.Switch(27) = 0:End Sub
Sub Sw28_Hit:Controller.Switch(28) = 1:WaSw27.IsDropped = 0:WaSw25.IsDropped = 1 : Playsound "fx_sensor" :End Sub
Sub Sw28_UnHit:Controller.Switch(28) = 0:End Sub

'**************************************
'       Diverters
'**************************************

'******** Step Gate Diverter
Sub SolFlipperDiverter(enabled)
     If Enabled Then
         PlaySound SoundFX("fx_divRR",DOFContactors)
    StepGate2.RotateToEnd
     Else
    StepGate2.RotateToStart
     End If
End Sub

'******** Ramp Diverter
Sub SolRampDiverter(enabled)
  If Enabled Then
    PlaySound SoundFX("fx_divLR",DOFContactors)
    RampDiv.RotateToEnd
     Else
    RampDiv.RotateToStart
     End If
End Sub

' **********************************
'    Bumpers
' **********************************

Sub Bumper1_Hit: vpmTimer.PulseSw(18) : Playsound SoundFX("fx_bumper1",DOFContactors): End Sub
Sub Bumper2_Hit: vpmTimer.PulseSw(77) : Playsound SoundFX("fx_bumper2",DOFContactors): End Sub
Sub Bumper3_Hit: vpmTimer.PulseSw(68) : Playsound SoundFX("fx_bumper3",DOFContactors): End Sub

' **********************************
'   SlingShots Animation
' **********************************

Dim LStep, RStep

Sub LeftSlingShot_Slingshot
  vpmTimer.PulseSw 41
    PlaySound SoundFX("fx_slingshotL",DOFContactors), 0, 1, 0.05, 0.05
    LSling.Visible = 0
    LSling1.Visible = 1
    sling2.TransZ = -20
    LStep = 0
    LeftSlingShot.TimerEnabled = 1
End Sub

Sub LeftSlingShot_Timer
    Select Case LStep
        Case 2:LSLing1.Visible = 0:LSLing2.Visible = 1:sling2.TransZ = -10
        Case 3:LSLing2.Visible = 0:LSLing.Visible = 1:sling2.TransZ = 0:LeftSlingShot.TimerEnabled = 0:
    End Select
    LStep = LStep + 1
End Sub

Sub RightSlingShot_Slingshot
  vpmTimer.PulseSw 53
    PlaySound SoundFX("fx_slingshotR",DOFContactors),0,1,-0.05,0.05
    RSling.Visible = 0
    RSling1.Visible = 1
    sling1.TransZ = -20
    RStep = 0
    RightSlingShot.TimerEnabled = 1
End Sub

Sub RightSlingShot_Timer
    Select Case RStep
        Case 2:RSLing1.Visible = 0:RSLing2.Visible = 1:sling1.TransZ = -10
        Case 3:RSLing2.Visible = 0:RSLing.Visible = 1:sling1.TransZ = 0:RightSlingShot.TimerEnabled = 0:
    End Select
    RStep = RStep + 1
End Sub

' **********************************
'  Trap Door Animation -Shoopity-
' **********************************

Dim TrapDir, PauseTrap: PauseTrap=0
Dim TrapSpeed : TrapSpeed = 2

Sub SolTrapDoorO(enabled)
  If PauseTrap Then
    TrapDoorPause.timerenabled=1
  Else
    TrapDoorRamp.Collidable = 1
    Trapwall.isdropped=0
    TrapDoorPause.enabled=0
    TrapDir = 1
    TrapMover.Enabled = 1
    Controller.Switch(76) = 0
    If enabled Then PlaySound SoundFX("fx_solon",DOFContactors)
  End If
End Sub

Sub SolTrapDoorC(enabled)
  TrapDir = -1
  TrapMover.Enabled = 1
  If enabled Then PlaySound SoundFX("fx_solon",DOFContactors)
End Sub

Sub TrapMover_Timer()
  PrTrap.RotX = PrTrap.RotX + TrapSpeed*TrapDir
  If PrTrap.RotX >= 120 Then
    PrTrap.RotX = 120
    me.enabled = 0
  End If
  If PrTrap.RotX >= 102 Then
    TrapDrop.enabled = 1
    'TrapDoorEnter1.Enabled = 1
    'TrapDoorEnter2.Enabled = 1
    trapwall.isdropped = 1
  Else
    TrapDrop.enabled = 0
    TrapDoorEnter1.Enabled = 0
    TrapDoorEnter2.Enabled = 0
    trapwall.isdropped = 0
  End If
  If PrTrap.RotX <= 90 Then
    PrTrap.RotX = 90
    trapwall.isdropped = 1
    TrapDoorRamp.Collidable = 0
    TrapDoorPause.enabled=1
    Controller.Switch(76) = 1
    me.enabled = 0
  End If
End Sub

'Sub TrapDrop_Hit():TrapDrop.timerenabled=1:debug.print "hit":End Sub
Sub TrapDrop_unHit()
  TrapDrop.timerenabled=1
  'debug.print "unhit"
End Sub

Sub TrapDrop_Timer()
  TrapDoorEnter1.Enabled = 1
  TrapDoorEnter2.Enabled = 1
  TrapDrop.timerenabled = 0
End Sub

Sub TrapDoorPause_Timer()
  If PauseTrap = 0 and TrapDir = 1 Then
    SolTrapDoorO(1)
    TrapDoorPause.timerenabled = 0
  ElseIf TrapDir = -1 Then
    TrapDoorPause.timerenabled = 0
  End If
End Sub

Sub TrapDoorPause_hit:PauseTrap=1:TrapDoorEnter1.Enabled = 0:TrapDoorEnter2.Enabled = 0:end Sub
Sub TrapDoorPause_unhit:PauseTrap=0::end Sub

Sub TrapDoorEnter1_Hit():TrapDoorHit:End Sub
Sub TrapDoorEnter2_Hit():TrapDoorHit:End Sub

Sub TrapDoorHit()
  PlaySound "kicker_enter_center"
  PauseTrap=0
End Sub


' **********************************
'  Rudy's Mouth Animation
' **********************************

Sub UpdateJawRudy(aNewPos,aSpeed,aLastPos)
  MoveMouth.Enabled = 1
End Sub

Sub MoveMouth_Timer()
  If RudyJawMech.position > PrMouth.RotX then:PrMouth.RotX = PrMouth.RotX +(0.5):End if
  If RudyJawMech.position < PrMouth.RotX then:PrMouth.Rotx = PrMouth.RotX -(0.5):End if
  If PrMouth.RotX < 5 then WaMouth.isDropped = 1:Else WaMouth.isDropped = 0
  prMouthb.Rotx = prMouth.Rotx
  prMouthSpringA.Size_Z = 29 + (PrMouth.RotX / 10)
  prMouthSpringB.Size_Z = 29 + (PrMouth.RotX / 10)
End Sub

' **********************************
'  Rudy's Eyes Animation -Shoopity-
' **********************************

Dim EyeDest
Dim EyeSpeed : EyeSpeed = 10

Sub SolEyesRight(enabled)
  Playsound "FX_Rudysol"
  MoveEyes.Enabled = 1
  If enabled Then EyeDest = 1 Else EyeDest = 0
End Sub
Sub SolEyesLeft(enabled)
  Playsound "FX_Rudysol1"
  MoveEyes.Enabled = 1
  If enabled Then EyeDest = -1 Else EyeDest = 0
End Sub

Sub MoveEyes_Timer()
  Select Case EyeDest
  Case -1:
    PrEyeL.RotZ = PrEyeL.RotZ + EyeSpeed
'   PrEyeR.RotZ = PrEyeR.RotZ + EyeSpeed
    If PrEyeL.RotZ >= 22 Then
      PrEyeL.RotZ = 22
'     PrEyeR.RotZ = 22
      me.enabled = 0
    End If
  Case 0:
    If PrEyeL.RotZ <= 2 Then
      PrEyeL.RotZ = PrEyeL.RotZ + EyeSpeed
'     PrEyeR.RotZ = PrEyeR.RotZ + EyeSpeed
    ElseIf PrEyeL.RotZ >= 2 Then
      PrEyeL.RotZ = PrEyeL.RotZ - EyeSpeed
'     PrEyeR.RotZ = PrEyeR.RotZ - EyeSpeed
    End If
    If PrEyeL.RotZ <= 2+(EyeSpeed+1) AND PrEyeL.RotZ >= 2-(EyeSpeed+1) Then
      PrEyeL.RotZ = 2
'     PrEyeR.RotZ = 2
      me.enabled = 0
    End If
  Case 1:
    PrEyeL.RotZ = PrEyeL.RotZ - EyeSpeed
'   PrEyeR.RotZ = PrEyeR.RotZ - EyeSpeed
    If PrEyeL.RotZ <= -22 Then
      PrEyeL.RotZ = -22
'     PrEyeR.RotZ = -22
      me.enabled = 0
    End If
  End Select
  If LazyEye = 1 Then
    PrEyeR.RotZ = PrEyeL.RotZ / 8
  Else
    PrEyeR.RotZ = PrEyeL.RotZ
  End If
  PrEye_Slider.RotY = PrEyeL.RotZ / 3
End Sub

' **********************************
'  Rudy's Lids Animation -Shoopity-
' **********************************

Dim LidSpeed : LidSpeed = 10
Dim LidDest


Sub SolEyesOpen(enabled)
  MoveLids.Enabled = 1
  Playsound "FX_Rudysol"
  If enabled Then
    LidDest = 1
  Else
    LidDest = 0
    prLidsThingerA.TransZ = -20
    prLidsThingerB.RotX = 95
  End If
End Sub

Sub SolEyesClosed(enabled)
  MoveLids.Enabled = 1
  Playsound "FX_Rudysol1"
  If enabled Then
    LidDest = -1
    prLidsThingerA.TransZ = 0
    prLidsThingerB.RotX = 90
  Else
'   LidDest = 0
  End If
End Sub

Sub MoveLids_Timer()
  Select Case LidDest
''''Lids Raised
  Case 1:
    PrLids.RotX = PrLids.RotX + LidSpeed
    If PrLids.RotX >= 5 Then  'was 55
      PrLids.RotX = 5   'was 55
      me.enabled = 0
    End If
    prLidsThingerB.TransY = (PrLids.RotX / 8) * -1
    prLidsThingerC.TransY = (PrLids.RotX / 8) * -1
    pSmallSpring.Size_Y = (24 - (prLidsThingerC.TransY * -2)) *.75

''''Lids Normal -midpoint
  Case 0:
    If PrLids.RotX <= -30 Then
      PrLids.RotX = PrLids.RotX + LidSpeed
    ElseIf PrLids.RotX >= -30 Then
      PrLids.RotX = PrLids.RotX - LidSpeed
    End If
    If PrLids.RotX <= (LidSpeed+1) AND PrLids.RotX >= -(LidSpeed+1) Then
      PrLids.RotX = -30
      me.enabled = 0
    End If
    prLidsThingerB.TransY = (PrLids.RotX / 8) * -1
    prLidsThingerC.TransY = (PrLids.RotX / 8) * -1
    pSmallSpring.Size_Y = (24 - (prLidsThingerC.TransY * -2)) *.75

''''Lids Lowered
  Case -1:
    PrLids.RotX = PrLids.RotX - LidSpeed
    If PrLids.RotX <= -75 Then  'was 115
      PrLids.RotX = -75   'was 115
      me.enabled = 0
    End If
  End Select
' prLidsThingerB.TransY = (PrLids.RotX / 8) * -1
' prLidsThingerC.TransY = (PrLids.RotX / 8) * -1
' pSmallSpring.Size_Y = (24 - (prLidsThingerC.TransY * -2)) *.75
' PrLids1.RotX = PrLids.RotX
End Sub

' **********************************
'  Rudy Twitch and Punch Sounds
' **********************************

Dim punchtype,TwitchCounter

Sub WaMouth_hit()

  If TwitchMod = 1 Then:TwitchTimer.enabled=1:TwitchCounter=0

    Select Case PunchMod
        Case 0:Playsound "fx_Flipperdown" 'normal sound
    Case 1 'random sounds
          Select Case Int(Rnd * 13) + 1
          Case 1:Playsound "punch"
          Case 2:Playsound "toasty"
          Case 3:Playsound "finish"
          Case 4:Playsound "Coocoo"
          Case 5:Playsound "glass"
          Case 6:Playsound "ricochet1"
          Case 7:Playsound "ricochet2"
          Case 8:Playsound "doink"
          Case 9:Playsound "drama"
          Case 10:Playsound "Cry"
          Case 11:Playsound "Excellent"
          Case 12:Playsound "Silly"
          Case 13:Playsound "Boing"
        End Select
    Case 2:Playsound "punch"
    Case 3:Playsound "toasty"
    Case 4:Playsound "finish"
    Case 5:Playsound "Coocoo"
    Case 6:Playsound "glass"
    Case 7:Playsound "ricochet1"
    Case 8:Playsound "ricochet2"
    Case 9:Playsound "doink"
    Case 10:Playsound "drama"
    Case 11:Playsound "Cry"
    Case 12:Playsound "Excellent"
    Case 13:Playsound "Silly"
    Case 14:Playsound "Boing"

    End Select

End Sub

Sub TwitchTimer_Timer()
  Dim TwitchMove
  Select Case TwitchCounter
    Case 0:TwitchMove=1:if RudyType=3 Then:PrLids.visible=false
    Case 1:TwitchMove=2
    Case 2:TwitchMove=3
    Case 3:TwitchMove=4
    Case 4:TwitchMove=5
    Case 5:TwitchMove=5
    Case 6:TwitchMove=4
    Case 7:TwitchMove=3
    Case 8:TwitchMove=2
    Case 9:TwitchMove=1
    Case 10:TwitchMove=0:if RudyType=3 Then:PrLids.visible=true:Me.Enabled = 0
  End Select

  PrRudy.TransX=TwitchMove/2:PrRudy.TransY=-TwitchMove:PrRudy.TransZ=TwitchMove
  PrRudy1.TransX=TwitchMove/2:PrRudy1.TransY=-TwitchMove:PrRudy1.TransZ=TwitchMove

  PrMouth.TransX=TwitchMove/2:PrMouth.TransY=-TwitchMove:PrMouth.TransZ=TwitchMove
  PrMouthb.TransX=TwitchMove/2:PrMouthb.TransY=-TwitchMove:PrMouthb.TransZ=TwitchMove
  PrMouthSpringA.TransX=TwitchMove/2:PrMouthSpringA.TransY=-TwitchMove:PrMouthSpringA.TransZ=TwitchMove
  PrMouthSpringB.TransX=TwitchMove/2:PrMouthSpringB.TransY=-TwitchMove:PrMouthSpringB.TransZ=TwitchMove

  If RudyType=3 Then
    PrEyeL.TransX=-TwitchMove:PrEyeL.TransY=TwitchMove*2:PrEyeL.TransZ=TwitchMove
    PrEyeR.TransX=-TwitchMove:PrEyeR.TransY=TwitchMove*2:PrEyeR.TransZ=TwitchMove
  End If

  TwitchCounter = TwitchCounter + 1
End Sub


' **********************************
'         Switches
' **********************************

'***Wire Triggers***
Sub Sw42_Hit:Controller.Switch(42) = 1 : playsound"fx_sensor" : End Sub
Sub Sw42_UnHit:Controller.Switch(42) = 0:End Sub
Sub Sw43_Hit:Controller.Switch(43) = 1 : playsound"fx_sensor" : End Sub
Sub Sw43_UnHit:Controller.Switch(43) = 0:End Sub
Sub Sw47_Hit:Controller.Switch(47) = 1 : PlaySound "fx_sensor" : End Sub
Sub Sw47_UnHit:Controller.Switch(47) = 0:End Sub
Sub Sw52_Hit:Controller.Switch(52) = 1 : PlaySound "fx_sensor" : End Sub
Sub Sw52_UnHit:Controller.Switch(52) = 0:End Sub
Sub Sw57_Hit:Controller.Switch(57) = 1 : PlaySound "fx_sensor" : End Sub
Sub Sw57_UnHit:Controller.Switch(57) = 0:End Sub
Sub Sw61_Hit:Controller.Switch(61) = 1 : PlaySound "fx_sensor" : End Sub
Sub Sw61_UnHit:Controller.Switch(61) = 0:End Sub
Sub Sw62_Hit:Controller.Switch(62) = 1 : PlaySound "fx_sensor" : End Sub
Sub Sw62_UnHit:Controller.Switch(62) = 0:End Sub
Sub Sw66_Hit:Controller.Switch(66) = 1 : PlaySound "fx_sensor" : End Sub
Sub Sw66_UnHit:Controller.Switch(66) = 0:End Sub
Sub Sw71_Hit:Controller.Switch(71) = 1 : PlaySound "fx_sensor" : End Sub
Sub Sw71_UnHit:Controller.Switch(71) = 0:End Sub
Sub sw75_Hit:Controller.Switch(75) = 1 : PlaySound "fx_sensor" : End Sub
Sub sw75_UnHit:Controller.Switch(75) = 0:End Sub

'***Standup Targets***
Sub sw17_Hit:vpmTimer.PulseSw(17):Playsound SoundFX("fx_target",DOFContactors):End Sub
Sub sw31_Hit:vpmTimer.PulseSw(31):Playsound SoundFX("fx_target",DOFContactors):End Sub
Sub sw32_Hit:vpmTimer.PulseSw(32):Playsound SoundFX("fx_target",DOFContactors):End Sub
Sub sw34_Hit:vpmTimer.PulseSw(34):Playsound SoundFX("fx_target",DOFContactors):End Sub
Sub sw37_Hit:vpmTimer.PulseSw(37):Playsound SoundFX("fx_target",DOFContactors):End Sub
Sub sw54_Hit:vpmTimer.PulseSw(54):Playsound SoundFX("fx_target",DOFContactors):End Sub
Sub sw64_Hit:vpmTimer.PulseSw(64):Playsound SoundFX("fx_target",DOFContactors):End Sub

'***Ramp Triggers***
Sub sw15_Hit:vpmTimer.PulseSw(15):PlaySound "fx_sensor":End Sub 'Step Lights Frenczy
Sub sw16_Hit:vpmTimer.PulseSw(16):PlaySound "fx_sensor":End Sub 'Upper Ramp Switch
Sub sw26_Hit:vpmTimer.PulseSw(26):PlaySound "fx_sensor":End Sub 'Step Light Exit Ball
Sub sw36_Hit:vpmTimer.PulseSw(36):PlaySound "fx_sensor":End Sub 'Step 500,000
Sub sw35_Hit:vpmTimer.PulseSw(35):PlaySound "fx_sensor":End Sub 'Step Tracker lower
Sub sw38_Hit:vpmTimer.PulseSw(38):PlaySound "fx_metalrolling":End Sub 'Step Tracker upper
Sub sw48_Hit:vpmTimer.PulseSw(48):PlaySound "fx_metalrolling":End Sub 'Ramp Exit Track
Sub sw45_Hit:vpmTimer.PulseSw(45):PlaySound "fx_sensor":End Sub 'Trap door score
Sub sw55_Hit:vpmTimer.PulseSw(55):End Sub 'Steps Superdog

'***Gate Triggers***
Sub sw33_Hit:vpmTimer.PulseSw(33):PlaySound "fx_gate":End Sub 'Upper left gangway roll under
Sub sw56_Hit:vpmTimer.PulseSw(56):PlaySound "fx_gate":End Sub 'Main Ramp Entrance
'Sub sw56_Hit:vpmTimer.PulseSw 78:sw56.timerenabled = 1:End Sub
'Sub sw56_timer:sw56.timerenabled= 0:End Sub


'***Rudy's mouth kicker animation***
Dim RKStep

Sub sw65_timer()
  Select Case RKStep
    Case 0:pRudyKick.TransY = 35
    Case 1:pRudyKick.TransY = 18
    Case 2:pRudyKick.TransY = 0:me.TimerEnabled = false:RKStep = 0
  End Select
  RKStep = RKStep + 1
End Sub

'******** Hidden switches
Sub sw51_Hit:vpmTimer.PulseSw(51):End Sub 'Dummy Jaw (opto)

'***************************************************
'      GI lights controlled by Strings
' 01 Upper BackGlass    'Case 0
' 02 Rudy         'Case 1
' 03 Upper Playfield      'Case 2
' 04 Center BackGlass   'Case 3
' 05 Lower Playfield    'Case 4

'***************************************************

Dim gistep

Set gicallback = GetRef("UpdateGI2")

sub UpdateGI2(no,enabled)
    Dim xx
    Select Case no
        Case 1
      If enabled then
        RudySign1.IntensityScale=1
        RudySign2.IntensityScale=1
        Rudylight.IntensityScale=1
        RudyShade.state=1

      Else
        RudySign1.IntensityScale=0
        RudySign2.IntensityScale=0
        Rudylight.IntensityScale=0
        RudyShade.state=0

      End If
        Case 2
      If enabled then
        For each xx in GI_Upper:xx.state=1:next
        If PopCornMod Then: lPopcornLight.state=1
      Else
        For each xx in GI_Upper:xx.state=0:next
        lPopcornLight.state=0
      End If
        Case 4
      If enabled then
        For each xx in GI_Lower:xx.state=1:next
        If HotDogCartMod Then: lHotDogCartA.state=1
      Else
        For each xx in GI_Lower:xx.state=0:next
        lHotDogCartA.state=0
      End If
    End Select
  If no = 2 OR no = 4 Then
    If NOT enabled then Table1.ColorGradeImage = "ColorGradeBOP_1"
  End If
End sub




Sub UpdateGI(no, step)
    Dim xx
    If step = 0 then exit sub 'only values from 1 to 8 are visible and reliable. 0 is not reliable and 7 & 8 are the same so...
    gistep = (step-1) / 7
    Select Case no
        Case 1
            For each xx in GI_Rudy:xx.IntensityScale = gistep:next
      If step >= 4 Then
        If RudyType = 3 Then
          PrRudy.Image = "Rudy_Face_On_2c"
          PrRudy1.Image = "Rudy_Back_On_2c"
          PrMouth.Image = "Rudy mouth baked on c"
        Else
          PrRudy.Image = "Rudy_Face_On_2"
          PrRudy1.Image = "Rudy_Back_On_2"
          PrMouth.Image = "Rudy mouth baked on"
        End If
      Else
        If RudyType = 3 then
          PrRudy.Image = "Rudy_Face_On_2c"
          PrRudy1.Image = "Rudy_Back_On_2c"
          PrMouth.Image = "Rudy mouth baked on c"
        Else
          PrRudy.Image = "Rudy_Face_Off_2"
          PrRudy1.Image = "Rudy_Back_Off_2"
          PrMouth.Image = "Rudy mouth baked off"
        End If
      End If
        Case 2
            For each xx in GI_Upper:xx.IntensityScale = gistep:next
      lPopcornLight.IntensityScale = gistep
        Case 4
            For each xx in GI_Lower:xx.IntensityScale = gistep:next
      lHotDogCartA.IntensityScale = gistep
    End Select

  If no = 2 OR no = 4 Then
    ' change the intensity of the flasher depending on the gi to compensate for the gi lights being off
    For xx = 0 to 200:FlashMax(xx) = 6 - gistep * 3 : Next  ' the maximum value of the flashers

    Table1.ColorGradeImage = "ColorGradeBOP_" & step
  End If
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
Dim TextureArray1: TextureArray1 = Array("Plastic with an image trans", "Plastic with an image")



InitLamps()             ' turn off the lights and flashers and reset them to the default parameters
LampTimer.Interval = 10 'lamp fading speed
LampTimer.Enabled = 1

' Lamp & Flasher Timers

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

Sub UpdateLamps

'***Inserts***

  FadeFlashm 11, fmfl11
  NFadeL 11, l11
  FadeFlashm 12, fmfl12
  NFadeL 12, l12
  FadeFlashm 13, fmfl13
  NFadeL 13, l13
  FadeFlashm 14, fmfl14
  NFadeL 14, l14
  NFadeL 15, l15
  NFadeL 16, l16
  NFadeL 17, l17
  NFadeL 18, l18
  If ClockMod = 1 then
    FadeFlashm 21, cfs45
  End If
  FadeFlashm 21, fmfl21
  NFadeL 21, l21
  If ClockMod = 1 then
    FadeFlashm 22, cfh8
  End If
  FadeFlashm 22, fmfl22
  NFadeL 22, l22
  If ClockMod = 1 then
    FadeFlashm 23, cfh6
  End If
  FadeFlashm 23, fmfl23
  NFadeL 23, l23
  If ClockMod = 1 then
    FadeFlashm 24, cfs25
  End If
  FadeFlashm 24, fmfl24
  NFadeL 24, l24
  If ClockMod = 1 then
    FadeFlashm 25, cfs15
  End If
  NFadeL 25, l25
  If ClockMod = 1 then
    FadeFlashm 26, cfs10
  End If
  NFadeL 26, l26
  If ClockMod = 1 then
    FadeFlashm 27, cfh12
  End If
  FadeFlashm 27, fmfl27
  NFadeL 27, l27
  If ClockMod = 1 then
    NFadeLm 28, l28b
  End If
  FadeFlashm 28, fmfl28
  NFadeL 28, l28
  If ClockMod = 1 then
    FadeFlashm 31, cfs40
  End If
  FadeFlashm 31, fmfl31
  NFadeL 31, l31
  If ClockMod = 1 then
    FadeFlashm 32, cfs35
  End If
  NFadeL 32, l32
  If ClockMod = 1 then
    FadeFlashm 33, cfs30
  End If
  FadeFlashm 33, fmfl33
  NFadeL 33, l33
  If ClockMod = 1 then
    FadeFlashm 34, cfs20
  End If
  NFadeL 34, l34
  If ClockMod = 1 then
    FadeFlashm 35, cfh3
  End If
  NFadeL 35, l35
  If ClockMod = 1 then
    FadeFlashm 36, cfh1
  End If
  NFadeL 36, l36
  If ClockMod = 1 then
    FadeFlashm 37, cfh11
  End If
  FadeFlashm 37, fmfl37
  NFadeL 37, l37
  If ClockMod = 1 then
    FadeFlashm 38, cfs50
  End If
  FadeFlashm 38, fmfl38
  NFadeL 38, l38
  If ClockMod = 1 then
    FadeFlashm 41, cfh9
  End If
  FadeFlashm 41, fmfl41
  NFadeL 41, l41
  If ClockMod = 1 then
    FadeFlashm 42, cfh7
  End If
  FadeFlashm 42, fmfl42
  NFadeL 42, l42
  If ClockMod = 1 then
    FadeFlashm 43, cfh5
  End If
  NFadeL 43, l43
  If ClockMod = 1 then
    FadeFlashm 44, cfh4
  End If
  NFadeL 44, l44
  If ClockMod = 1 then
    FadeFlashm 45, cfh2
  End If
  NFadeL 45, l45
  If ClockMod = 1 then
    FadeFlashm 46, cfs5
  End If
  NFadeL 46, l46
  If ClockMod = 1 then
    FadeFlashm 47, cfs55
  End If
  FadeFlashm 47, fmfl47
  NFadeL 47, l47
  If ClockMod = 1 then
    FadeFlashm 48, cfh10
  End If
  FadeFlashm 48, fmfl48
  NFadeL 48, l48
  If BalloonMod = 1 Then
    FadeMaterialP 51, prballoon_Blue, TextureArray1
    NFadeLm 51, LBballoon
  End If
  NFadeLm 51, l51
  NFadeLm 51, l51a
  NFadeL 51, l51b
  If BalloonMod = 1 Then
    FadeMaterialP 52, prballoon_Red, TextureArray1
    NFadeLm 52, LRballoon
  End If
  NFadeLm 52, l52
  NFadeLm 52, l52a
  NFadeL 52, l52b
  If HotDogCartMod = 1 then
  FadeMaterialP 53, prHotDogCartC, TextureArray1
  NFadeLm 53, lHotDogCartB
  End If
  NFadeLm 53, l53
  NFadeL 53, l53a
  NFadeLm 54, Bot_finger_1
  NFadeLm 54, Bot_finger_2
  NFadeLm 54, Bot_finger_3
  NFadeL 54, Bot_finger_4
  NFadeLm 55, Mid_finger_1
  NFadeLm 55, Mid_finger_2
  NFadeLm 55, Mid_finger_3
  NFadeL 55, Mid_finger_4
  NFadeLm 56, Top_finger_1
  NFadeLm 56, Top_finger_2
  NFadeLm 56, Top_finger_3
  NFadeL 56, Top_finger_4
    Flashm 57, F57_FS
    Flashm 57, F57a_FS
    Flashm 57, F57_DT
    Flash 57, F57a_DT
    Flashm 58, F58_DT
    Flashm 58, F58a_DT
    Flashm 58, F58a_FS
    Flash 58, F58_FS
    NFadeLm 61, l61a
    NFadeL 61, l61
    NFadeL 62, l62
  FadeFlashm 63, fmfl63
    NFadeL 63, l63
    NFadeL 64, l64
    NFadeL 65, l65
    NFadeL 66, l66
    NFadeL 67, l67
    NFadeL 68, l68
    Flashm 71, F71
  FadeFlashm 71, fmfl71
    Flash 71, F71a
    Flashm 74, F74
  FadeFlashm 74, fmfl74
    Flash 74, F74a
    Flashm 75, F75
  FadeFlashm 75, fmfl75
    Flash 75, F75a
    Flashm 76, F76
  FadeFlashm 76, fmfl76
    Flash 76, F76a
    Flashm 77, F77
  FadeFlashm 77, fmfl77
    Flash 77, F77a
    Flashm 78, F78
  FadeFlashm 78, fmfl78
    Flash 78, F78a
  If BalloonMod = 1 Then
    FadeMaterialP 72, prballoon_Yellow, TextureArray1
    NFadeLm 72, LYballoon
  End If
    NFadeLm 72, l72
    NFadeLm 72, l72a
    NFadeL 72, l72b
    NFadeL 73, l73
    NFadeL 81, l81
    NFadeLm 82, l82a
    NFadeL 82, l82
  FadeFlashm 83, fmfl83
    NFadeL 83, l83
    NFadeL 84, l84
    NFadeL 85, l85
    NFadeL 86, l86
    NFadeL 87, l87
    'NFadeL 88, l88   'Start Button Cabinet

'***Flashers***

  NFadeLm 117, F17
  NFadeLm 117, F17a
  Flashm 117, F17b
  Flashm 117, F17b1
  Flashm 117, Flasher17_FS
  Flashm 117, Flasher17_DT
  Flash 117, Flasher3
  NFadeLm 118, F18
  Flash 118, F18a
  If ClockMod = 1 then
    FadeFlashm 119, cfcenter
  End If
  FadeFlashm 119, fmfl19
  NFadeLm 119, L19
  Flash 119, F19
  NFadeLm 123, F23
  NFadeLm 123, F23a
  Flashm 123, F23b
  Flashm 123, F23b1
  Flashm 123, Flasher23_FS
  Flashm 123, Flasher23_DT
  Flash 123, Flasher4
  NFadeLm 124, F24
  NFadeLm 124, F24a
  Flashm 124, F24b
  Flashm 124, F24b1
  Flashm 124, Flasher24_FS
  Flashm 124, Flasher24_DT
  Flash 124, Flasher5
End Sub

'<<<<<<<<<<<<<Is this scripting needed???>>>>>>>>>Lines 1265 - 1414
' div lamp subs

Sub InitLamps()
    Dim x
    For x = 0 to 200
        LampState(x) = 0        ' current light state, independent of the fading level. 0 is off and 1 is on
        FadingLevel(x) = 4      ' used to track the fading state
        FlashSpeedUp(x) = 0.4   ' faster speed when turning on the flasher
        FlashSpeedDown(x) = 0.2 ' slower speed when turning off the flasher
        FlashMax(x) = 1         ' the maximum value when on, usually 1
        FlashMin(x) = 0         ' the minimum value when off, usually 0
        FlashLevel(x) = 0       ' the intensity of the flashers, usually from 0 to 1
    Next
End Sub

Sub AllLampsOff
    Dim x
    For x = 0 to 200
        SetLamp x, 0
    Next
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
        Case 4:object.image = b:FadingLevel(nr) = 6                   'fading to off...
        Case 5:object.image = a:FadingLevel(nr) = 1                   'ON
        Case 6, 7, 8:FadingLevel(nr) = FadingLevel(nr) + 1             'wait
        Case 9:object.image = c:FadingLevel(nr) = FadingLevel(nr) + 1 'fading...
        Case 10, 11, 12:FadingLevel(nr) = FadingLevel(nr) + 1         'wait
        Case 13:object.image = d:FadingLevel(nr) = 0                  'Off
    End Select
End Sub

Sub FadeObjm(nr, object, a, b, c, d)
    Select Case FadingLevel(nr)
        Case 4:object.image = b
        Case 5:object.image = a
        Case 9:object.image = c
        Case 13:object.image = d
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


'***Fade Materials***

dim itemw, itemp

Sub FadeMaterialW(nr, itemw, group)
    Select Case FadingLevel(nr)
        Case 4:itemw.TopMaterial = group(1):itemw.SideMaterial = group(1)
        Case 5:itemw.TopMaterial = group(0):itemw.SideMaterial = group(0)
    End Select
End Sub

Sub FadeMaterialP(nr, itemp, group)
    Select Case FadingLevel(nr)
        Case 4:itemp.Material = group(1)
        Case 5:itemp.Material = group(0)
    End Select
End Sub



'***Flasher objects***

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



Sub FadeFlash(nr, object)
    Select Case FadingLevel(nr)
        Case 4:object.Opacity = 0:FadingLevel(nr) = 6                   'fading to off...
        Case 5:object.Opacity = 100:FadingLevel(nr) = 1                   'ON
        Case 6, 7, 8:FadingLevel(nr) = FadingLevel(nr) + 1             'wait
        Case 9:object.Opacity = 66:FadingLevel(nr) = FadingLevel(nr) + 1 'fading...
        Case 10, 11, 12:FadingLevel(nr) = FadingLevel(nr) + 1         'wait
        Case 13:object.Opacity = 33:FadingLevel(nr) = 0                  'Off
    End Select
End Sub

Sub FadeFlashm(nr, object)
    Select Case FadingLevel(nr)
        Case 4:object.Opacity = 0
        Case 5:object.Opacity = 100
        Case 9:object.Opacity = 66
        Case 13:object.Opacity = 33
    End Select
End Sub

Sub BallHitSound(dummy):PlaySound "ball_bounce":End Sub

'*****************************************
' JF's Sound Routines
'*****************************************

Sub RubbersSmallRings_Hit(idx)
    dim finalspeed
    finalspeed = SQR(activeball.velx * activeball.velx + activeball.vely * activeball.vely)
    If finalspeed > 20 then
        PlaySound "fx_rubber", 0, 3*Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
    End if
    If finalspeed >= 6 AND finalspeed <= 20 then
        RandomSoundRubber()
    End If
End Sub

Sub RubbersBandsLargeRings_Hit(idx)
    dim finalspeed
    finalspeed = SQR(activeball.velx * activeball.velx + activeball.vely * activeball.vely)
    If finalspeed > 20 then
        PlaySound "fx_rubber", 0, 3*Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
    End if
    If finalspeed >= 6 AND finalspeed <= 20 then
        RandomSoundRubber()
    End If
End Sub


Sub RandomSoundRubber()
    Select Case Int(Rnd * 3) + 1
        Case 1:PlaySound "fx_rubber_hit_1", 0, 2*Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
        Case 2:PlaySound "fx_rubber_hit_2", 0, 2*Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
        Case 3:PlaySound "fx_rubber_hit_3", 0, 2*Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
    End Select
End Sub

Sub Rampdiv_Collide(parm)
PlaySound "fx_metalhit2": End Sub

Sub LeftFlipper_Collide(parm)
  RandomSoundFlipper()
End Sub

Sub LeftFlipper1_Collide(parm)
  RandomSoundFlipper()
End Sub

Sub RightFlipper_Collide(parm)
  RandomSoundFlipper()
End Sub

Sub RandomSoundFlipper()
  Select Case Int(Rnd*3)+1
    Case 1 : PlaySound "fx_flip_hit_1", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
    Case 2 : PlaySound "fx_flip_hit_2", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
    Case 3 : PlaySound "fx_flip_hit_3", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
  End Select
End Sub

Sub waSW25_Hit : Playsound "fx_metalhit2" : End Sub
Sub waSW27_Hit : Playsound "fx_metalhit2" : End Sub
Sub waSW28_Hit : Playsound "fx_metalhit2" : End Sub

Sub tr27_Hit : Playsound "fx_stepdrop" : End Sub
Sub tr26_Hit : Playsound "fx_stepdrop" : End Sub
Sub tr25_Hit : Playsound "fx_stepdrop" : End Sub
Sub tr24_Hit : Playsound "fx_stepdrop" : End Sub
Sub tr23_Hit : Playsound "fx_stepdrop" : End Sub
Sub tr23a_Hit : Playsound "fx_stepdrop" : End Sub
Sub tr3_Hit : Playsound "fx_stepdrop" : End Sub
Sub tr4_Hit : Playsound "fx_stepdrop" : End Sub
Sub tr6_Hit : Playsound "fx_stepdrop" : End Sub

Sub Trigger1_Hit():PlaySound "fx_lr5":End Sub
Sub Trigger2_Hit():PlaySound "fx_lr5":End Sub
Sub Trigger3_Hit():PlaySound "fx_lr5":End Sub
Sub Trigger4_Hit():PlaySound "fx_lr5":End Sub
Sub Trigger5_Hit():PlaySound "fx_lr1":End Sub
Sub Trigger6_Hit():PlaySound "fx_lr5" End Sub
Sub Trigger7_Hit():PlaySound "fx_lr5":End Sub
Sub Trigger8_Hit():PlaySound "fx_lr2":End Sub
Sub Trigger9_Hit():PlaySound "fx_lr6":End Sub
Sub Trigger10_Hit():PlaySound "fx_lr6":End Sub

'***Wire ramp sounds***
Sub LWireStart_Hit():PlaySound "fx_metalrolling_FH":End Sub
Sub RWireStart_Hit():PlaySound "fx_metalrolling_FH":End Sub

Sub LWireEnd_Hit()
     vpmTimer.AddTimer 150, "BallHitSound"
   StopSound "fx_metalrolling_FH"
 End Sub

Sub RWireEnd_Hit()
     vpmTimer.AddTimer 150, "BallHitSound"
   StopSound "fx_metalrolling_FH"
 End Sub

'Sub RWireStart_Hit()
'If ActiveBall.VelY < 0 Then Playsound "fx_metalrolling"
'End Sub

'Sub RWireEnd_Hit()
 '    vpmTimer.AddTimer 150, "BallHitSound"
  ' StopSound "fx_metalrolling"
 'End Sub

'Sub RWireStart_Hit()
'If ActiveBall.VelY < 0 Then Playsound "fx_metalrolling"
'End Sub

'Sub RWireEnd_Hit()
 '    vpmTimer.AddTimer 150, "BallHitSound"
'  StopSound "fx_metalrolling"
 'End Sub
'**************************

Sub Metals_Hit(idx):PlaySound "fx_metalhit2", 0, Vol(ActiveBall), pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0, AudioFade(ActiveBall):End Sub
Sub Metals1_Hit(idx):PlaySound "fx_metalhit2": End Sub
Sub SpotTargets_Hit(idx):PlaySound "fx_target", 0, Vol(ActiveBall), pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0, AudioFade(ActiveBall):End Sub
Sub Gates_Hit (idx): PlaySound "fx_gate", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall): End Sub

Sub plungeballdrop_Hit()
If ActiveBall.VelY > 0 Then Playsound "wirerampdrop"
End Sub

'**********************
' Balldrop & Ramp Sound
'**********************

Sub BallDropSound(dummy)
    PlaySound "fx_ball_drop"
End Sub

Sub Balldrop1_Hit()
    vpmtimer.addtimer 300, "BallDropSound"
    StopSound "fx_metalrolling"
End Sub

Sub Balldrop2_Hit()
    vpmtimer.addtimer 300, "BallDropSound"
    StopSound "fx_metalrolling"
End Sub

Sub Balldrop3_Hit()
    vpmtimer.addtimer 200, "BallDropSound"
End Sub

'******************
'Switch animations
'*******************

Sub sw35_Hit:Controller.Switch(35) = 1:sw35.timerenabled = true:End Sub
Sub sw35_UnHit:Controller.Switch(35) = 0:End Sub

'***switch 35 animation***

Const Switch35min = 0
Const Switch35max = -20
Dim Switch35dir
Switch35dir = -2

Sub sw35_timer()
 pRampSwitch1B.RotY = pRampSwitch1B.RotY + Switch35dir
  If pRampSwitch1B.RotY >= Switch35min Then
    sw35.timerenabled = False
    pRampSwitch1B.RotY = Switch35min
    Switch35dir = -2
  End If
  If pRampSwitch1B.RotY <= Switch35max Then
    Switch35dir = 4
  End If
End Sub


Sub sw48_Hit:Controller.Switch(48) = 1:sw48.timerenabled = true:End Sub
Sub sw48_UnHit:Controller.Switch(48) = 0:End Sub

'***switch 48 animation***

Const Switch48min = 0
Const Switch48max = -20
Dim Switch48dir
Switch48dir = -2

Sub sw48_timer()
  pRampSwitch3B.RotY = pRampSwitch3B.RotY + Switch48dir

  If pRampSwitch3B.RotY >= Switch48min Then
    sw48.timerenabled = False
    pRampSwitch3B.RotY = Switch48min
    Switch48dir = -2
  End If

  If pRampSwitch3B.RotY <= Switch48max Then
    Switch48dir = 4
  End If
End Sub

Sub sw38_Hit:Controller.Switch(38) = 1:sw38.timerenabled = true:End Sub
Sub sw38_UnHit:Controller.Switch(38) = 0:End Sub

'***switch 38 animation***

Const Switch38min = 0
Const Switch38max = -20
Dim Switch38dir
Switch38dir = -2

Sub sw38_timer()
  pRampSwitch2B.RotY = pRampSwitch2B.RotY + Switch38dir
  If pRampSwitch2B.RotY >= Switch38min Then
    sw38.timerenabled = False
    pRampSwitch2B.RotY = Switch38min
    Switch38dir = -2
  End If
  If pRampSwitch2B.RotY <= Switch38max Then
    Switch38dir = 4
  End If
End Sub


Sub sw15_Hit:Controller.Switch(15) = 1:sw15.timerenabled = true:End Sub
Sub sw15_UnHit:Controller.Switch(15) = 0:End Sub

'***switch 15 animation***

Const Switch15min = 0
Const Switch15max = -20
Dim Switch15dir
Switch15dir = -2

Sub sw15_timer()
  pRampSwitch4B.RotX = pRampSwitch4B.RotX + Switch15dir
  If pRampSwitch4B.RotX >= Switch15min Then
    sw15.timerenabled = False
    pRampSwitch4B.RotX = Switch15min
    Switch15dir = -2
  End If
  If pRampSwitch4B.RotX <= Switch15max Then
    Switch15dir = 4
  End If
End Sub


Sub sw26_Hit:Controller.Switch(26) = 1:sw26.timerenabled = true:End Sub
Sub sw26_UnHit:Controller.Switch(26) = 0:End Sub

'***switch 26 animation***

Const Switch26min = 0
Const Switch26max = -20
Dim Switch26dir
Switch26dir = -2

Sub sw26_timer()
  pRampSwitch5B.RotX = pRampSwitch5B.RotX + Switch26dir
  If pRampSwitch5B.RotX >= Switch26min Then
    sw26.timerenabled = False
    pRampSwitch5B.RotX = Switch26min
    Switch26dir = -2
  End If
  If pRampSwitch5B.RotX <= Switch26max Then
    Switch26dir = 4
  End If
End Sub


Sub sw36_Hit:Controller.Switch(36) = 1:sw36.timerenabled = true:End Sub
Sub sw36_UnHit:Controller.Switch(36) = 0:End Sub

'***switch 36 animation***

Const Switch36min = 0
Const Switch36max = -20
Dim Switch36dir
Switch36dir = -2

Sub sw36_timer()
  pRampSwitch6B.RotX = pRampSwitch6B.RotX + Switch36dir
  If pRampSwitch6B.RotX >= Switch36min Then
    sw36.timerenabled = False
    pRampSwitch6B.RotX = Switch36min
    Switch36dir = -2
  End If
  If pRampSwitch6B.RotX <= Switch36max Then
    Switch36dir = 4
  End If
End Sub


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''  Ball Trough system ''''''''''''''''''''''''''
'''''''''''''''''''''by cyberpez''''''''''''''''''''''''''''''''
''''''''''''''''based off of EalaDubhSidhe's''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Dim BallCount

Sub CheckMaxBalls()
  BallCount = MaxBalls
  TroughWall1.isDropped = true
  TroughWall2.isDropped = true
End Sub

Dim cBall1, cBall2, cBall3

Sub CreatBalls_timer()
  If BallCount > 0 then
    If BallMod = 1 Then
      If BallCount = 3 Then
        Set cBall1 = drain.CreateSizedBallWithMass(BallSize, BallMass)
        cBall1.Image = "Chrome_Ball_29"
        cBall1.FrontDecal = "FunhouseBall1"
      End If
      If BallCount = 2 Then
        Set cBall2 = drain.CreateSizedBallWithMass(BallSize, BallMass)
        cBall2.Image = "Chrome_Ball_29"
        cBall2.FrontDecal = "FunhouseBall2"
      End If
      If BallCount = 1 Then
        Set cBall3 = drain.CreateSizedBallWithMass(BallSize, BallMass)
        cBall3.Image = "Chrome_Ball_29"
        cBall3.FrontDecal = "FunhouseBall3"
      End If
    Else
      If BallCount = 3 Then
        Set cBall1 = drain.CreateSizedBallWithMass(BallSize, BallMass)
      End If
      If BallCount = 2 Then
        Set cBall2 = drain.CreateSizedBallWithMass(BallSize, BallMass)
      End If
      If BallCount = 1 Then
        Set cBall3 = drain.CreateSizedBallWithMass(BallSize, BallMass)
      End If
    End If
    Drain.kick 70,30
    BallCount = BallCount - 1
  End If

  If BallCount = 0 Then
    CreatBalls.enabled = false
    setoptions
  End If
End Sub

Dim DRSstep

Sub DelayRollingStart_timer()
  Select Case DRSstep
    Case 5: RollingSoundTimer.enabled = true
  End Select
  DRSstep = DRSstep + 1
End Sub

Sub ballrelease_hit()
' Kicker1active = 1
  Controller.Switch(63)=1
  TroughWall1.isDropped = false

End Sub

Sub sw74_Hit()
  Controller.Switch(74)=1
  TroughWall2.isDropped = false
End Sub

Sub sw74_unHit()
  Controller.Switch(74)=0
  TroughWall2.isDropped = true
End Sub

Sub sw72_Hit()
  Controller.Switch(72)=1
End Sub

Sub sw72_unHit()
  Controller.Switch(72)=0
End Sub

Dim DontKickAnyMoreBalls,DKTMstep

Sub KickBallToLane(Enabled)
  If DontKickAnyMoreBalls = 0 then
    PlaySound SoundFX("fx_ballrel",DOFContactors)
    PlaySound SoundFX("Solenoid",DOFContactors)
    ballrelease.Kick 60,10
    TroughWall1.isDropped = true
    Controller.Switch(63)=0
    DontKickAnyMoreBalls = 1
    DKTMstep = 1
    DontKickToMany.enabled = true
  End If
End Sub

Sub DontKickToMany_timer()
  Select Case DKTMstep
    Case 1:
    Case 2:
    Case 3: DontKickAnyMoreBalls = 0:DontKickToMany.Enabled = False: DontKickAnyMoreBalls = 0
  End Select
  DKTMstep = DKTMstep + 1
End Sub

sub kisort(enabled)
  Drain.Kick 70,30
  controller.switch(73) = false
end sub

Sub Drain_hit()
  PlaySound "drain"
  controller.switch(73) = true
End Sub

'***Ball brakes***
Sub ramp_brake1_Hit()
    ActiveBall.vely = Activeball.vely/5
End Sub

Sub ramp_brake2_Hit()
    ActiveBall.vely = Activeball.vely/5
End Sub

Sub ramp_brake3_Hit()
    ActiveBall.velx = Activeball.velx/2
End Sub

Sub ramp_brake4_Hit()
    ActiveBall.vely = Activeball.vely/2
End Sub

Sub ramp_brake5_Hit()
    ActiveBall.velx = Activeball.velx/10
End Sub

Sub ramp_brake6_Hit()
    ActiveBall.vely = Activeball.vely/5
End Sub

Sub ramp_brake7_Hit()
    ActiveBall.vely = Activeball.vely/5
End Sub


'***Rubber animations***
Sub wall41_Hit:rubber16.visible = 0::rubber16a.visible = 1:rubber18.visible = 0::rubber18a.visible = 1:wall41.timerenabled = 1:End Sub
Sub wall41_timer:rubber16.visible = 1::rubber16a.visible = 0:rubber18.visible = 1::rubber18a.visible = 0: wall41.timerenabled= 0:End Sub
Sub wall58_Hit:rubber31.visible = 0::rubber31a.visible = 1::rubber5.visible = 0:rubber5a.visible = 1:wall58.timerenabled = 1:End Sub
Sub wall58_timer:rubber31.visible = 1::rubber31a.visible = 0:rubber5.visible = 1:rubber5a.visible = 0:wall58.timerenabled= 0:End Sub
' _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _
'(_|_|_|_|_|_|_|_|_|_|_|_|_|_|_|_|_|_|_|_|_|_|_|_|_|_|_|_|_|_|_|_|_|_|_|_|_|_|_|_|_|_|_|_)


'__          __   _____       _   _                   _                     __          __
'\ \        / /  |  _  |     | | (_)                 | |                    \ \        / /
' \ \      / /   | | | |_ __ | |_ _  ___  _ __  ___  | |__   ___ _ __ ___    \ \      / /
'  \ \    / /    | | | | '_ \| __| |/ _ \| '_ \/ __| | '_ \ / _ \ '__/ _ \    \ \    / /
'   \ \  / /     \ \_/ / |_) | |_| | (_) | | | \__ \ | | | |  __/ | |  __/     \ \  / /
'    \_\/_/       \___/| .__/ \__|_|\___/|_| |_|___/ |_| |_|\___|_|  \___|      \_\/_/
'                      | |
'                      |_|

Dim TableOptions, TableName
Private vpmShowDips1, vpmDips1

Sub InitializeOptions
  TableName="Funhouse_VPX"                  'Your descriptive table name, it will be used to save settings in VPReg.stg file
  Set vpmShowDips1 = vpmShowDips                'Reassigns vpmShowDips to vpmShowDips1 to allow usage of default dips menu
  Set vpmShowDips = GetRef("TableShowDips")         'Assigns new sub to vmpShowDips
  TableOptions = LoadValue(TableName,"Options")       'Load saved table options

  Set Controller = CreateObject("VPinMAME.Controller")    'Load vpm controller temporarily so options menu can be loaded if needed
  If TableOptions = "" Or optReset Then           'If no existing options, reset to default through optReset, then open Options menu
    TableOptions = 1                    'clear any existing settings and set table options to default options
    TableShowOptions
  ElseIf (TableOptions And 1) Then              'If Enable Next Start was selected then
    'TableOptions = TableOptions - 1              'clear setting to avoid future executions
    TableShowOptions
  Else
    TableSetOptions
  End If

  Set Controller = Nothing                  'Unload vpm controller so selected controller can be loaded
End Sub

Private Sub TableShowDips
  vpmShowDips1                        'Show original Dips menu
  TableShowOptions                      'Show new options menu
End Sub

Private Sub TableShowOptions          'New options menu
  Dim oldOptions : oldOptions = TableOptions
  If Not IsObject(vpmDips1) Then
    Set vpmDips1 = New cvpmDips
    With vpmDips1
      .AddForm 700, 500, "TABLE OPTIONS MENU"
      .AddFrameExtra 0,0,155,"Mirror Lights color Mod",2^1, Array("Normal Mirror", 0, "RWB Mirror", 2^1)
      .AddFrameExtra 0,45,155,"Lazy Eye Mod",16, Array("Normal", 0, "Lazy Eye", 16)
      .AddFrameExtra 0,90,155,"Ball Type Mod",32, Array("Standard Ball", 0, "Marbled Ball", 32)
      .AddFrameExtra 0,135,155,"Rudy Mouth Hit Twitch", 64, Array("Normal", 0, "Face Twitch", 64)
      .AddFrameExtra 0,181,155,"Drain Post", 2^29, Array("No Drain Post", 0, "Add Drain Post", 2^29)
      .AddFrameExtra 0,227,155,"Mirror Lettering Mod", 2^30, Array("Normal Letters", 0, "Background Letters", 2^30)
      .AddFrameExtra 0,273,155,"Custom Apron/Walls Mod", 2^22, Array("Normal Apron/Walls", 0, "Custom Apron/Walls", 2^22)
      .AddFrameExtra 0,320,155,"Instruction Card Mod", 2^24, Array("Standard Cards", 0, "Random Cards", 2^24)
      .AddFrameExtra 175,0,155,"Flipper Color Mod",128, Array("Yellow/Red", 0, "Yellow/Blue", 128)
      .AddFrameExtra 175,45,155,"Clock Mod",256, Array("No Clock", 0, "Show Clock", 256)
      .AddFrameExtra 175,90,155,"Balloons Mod", 512, Array("No Balloons", 0, "Show Balloons", 512)
      .AddFrameExtra 175,135,155,"Popcorn Bucket Mod", 1024, Array("No Popcorn", 0, "Show Popcorn", 1024)
      .AddFrameExtra 175,180,155,"Hotdog Cart Mod", 2048, Array("No Hotdog Cart", 0, "Hotdog Cart", 2048)
      .AddFrameExtra 175,225,155,"Bubble Level Mod", 4096, Array("No Level", 0, "Show Level", 4096)
      .AddFrameExtra 175,270,155,"Subway Color Mod", 2^13+2^14, Array("No Color Added", 0, "Blue", 2^13, "Red", 2^14)
      .AddFrameExtra 350,0,155,"Mouth Hit Sound Mod", 2^15+2^16+2^17+2^18+2^19+2^20+2^21+2^23+2^25+2^26+2^27+2^28, Array("No Sound Effect", 0, "Random", 2^15, "Punch Sound", 2^16, "Toasty", 2^17, "Finish Him", 2^18, "Coo Coo", 2^19, "Glass Break", 2^20, "Richochet1", 2^21, "Doink", 2^23, "Cry", 2^25, "Excellent", 2^26, "Silly", 2^27, "Boing", 2^28)

      .Addlabel 350,205,155,20,"Left Magna-save Button"
      .Addlabel 350,220,155,21,"Toggles:"
      .Addlabel 350,235,155,21,"Arcade Ambiant Sounds"

      .Addlabel 350,265,155,20,"Right Magna-save Button"
      .Addlabel 350,280,155,21,"Toggles:"
      .Addlabel 350,295,155,21,"Rudy Face Mod"


      .AddChkExtra 350,330,155, Array("Enable Menu Next Start", 1)
      .Addlabel 350,350,155,21,"To Re-activate Menu,"
      .Addlabel 350,365,155,21,"See Script line 46"

    End With
  End If
  TableOptions = vpmDips1.ViewDipsExtra(TableOptions)
  SaveValue TableName,"Options",TableOptions
  TableSetOptions
  SetOptions
End Sub

Sub TableSetOptions   'defines required settings before table is run
  MirrorRWBMod = (TableOptions And 2^1):If MirrorRWBMod = 2^1 Then MirrorRWBMod = 1
  LazyEye = (TableOptions And 16):If LazyEye=16 Then:LazyEye=1
  BallMod = (TableOptions And 32):If BallMod=32 Then:BallMod=1
  TwitchMod = (TableOptions And 64):If TwitchMod=64 Then:TwitchMod=1
    DrainPostMod = (TableOptions And 2^29):If DrainPostMod=2^29 Then:DrainPostMod=1
    MirrorMod = (TableOptions And 2^30):If MirrorMod=2^30 Then:MirrorMod=1
    ApronMod = (TableOptions And 2^22):If ApronMod=2^22 Then:ApronMod=1
  FlipperColor = (TableOptions And 128):If FlipperColor=128 Then:FlipperColor=1
  ClockMod = (TableOptions And 256):If ClockMod=256 Then:ClockMod=1
  BalloonMod = (TableOptions And 512):If BalloonMod=512 Then:BalloonMod=1
  PopcornMod = (TableOptions And 1024):If PopcornMod=1024 Then:PopcornMod=1
  HotDogCartMod = (TableOptions And 2048):If HotDogCartMod=2048 Then:HotDogCartMod=1
  LevelMod = (TableOptions And 4096):If LevelMod=4096 Then:LevelMod=1
    SubwayColorMod = (TableOptions And (2^13+2^14))
    Select Case SubwayColorMod
      Case 0: SubwayColorMod = 0
      Case 2^13: SubwayColorMod = 1
      Case 2^14: SubwayColorMod = 2
    End Select
  PunchMod= (TableOptions And (2^15+2^16+2^17+2^18+2^19+2^20+2^21+2^23+2^25+2^26+2^27+2^28))
    Select Case PunchMod
      Case 0: PunchMod = 0
      Case 2^15: PunchMod = 1
      Case 2^16: PunchMod = 2
      Case 2^17: PunchMod = 3
      Case 2^18: PunchMod = 4
      Case 2^19: PunchMod = 5
      Case 2^20: PunchMod = 6
      Case 2^21: PunchMod = 7
      'Case 2^22: PunchMod = 8
      Case 2^23: PunchMod = 9
      'Case 2^24: PunchMod = 10
      Case 2^25: PunchMod = 11
      Case 2^26: PunchMod = 12
      Case 2^27: PunchMod = 13
      Case 2^28: PunchMod = 14
    End Select
  CardMod = (TableOptions And 2^24):If CardMod=2^24 Then:CardMod=1
  SaveValue TableName,"Options",TableOptions
  SetOptions
End Sub


'''''''Set Options
Dim RudyType, cheaterpost, CardType

Sub SetOptions()
  If MirrorMod = 1 Then
    pMirrorFrontB.image = "Mirror_FrontB_texture2"
  Else
    pMirrorFrontB.image = "Mirror_FrontB_texture"
  End If

  If ApronMod = 1 Then
    pApronOverlay.visible = 1
    If DesktopMode = True Then 'Show Desktop components
      pSidewall_DT.visible = 1
      pSidewall_FS.visible = 0
    Else
      pSidewall_FS.visible = 1
      pSidewall_DT.visible = 0
    End If
  Else
    pApronOverlay.visible = 0
    pSidewall_DT.visible = 0
    pSidewall_FS.visible = 0
  End If

  If DrainPostMod = 1 Then
    cpost.visible = 1
    crubber.collidable = 1
    crubber.visible = 1
  Else
    cpost.visible = 0
    crubber.collidable = 0
    crubber.visible = 0
  End If

' If RudyMod = 0 then
'   RudyType = Int(Rnd*3)+1
' Else
'   RudyType = RudyMod
' End If

' If RudyType = 1 Then
'   PrRudy.Visible = True
'   PrLids.Image = "Rudy eyelid1"
'   PrEyeL.Image = "eye_texture"
'   PrEyeR.Image = "eye_texture"
'   PrRudy.Image = "Rudy_Face_Off_2"
'   PrRudy1.Image = "Rudy_Back_Off_2"
'   PrMouth.Image = "Rudy mouth baked off"
'   prRIWCage.Material = "Metal with an image Dark"
'   prRudyScoop.Material = "Metal with an image Dark"
'   PrMouthb.Material = "Metal with an image Dark"
' End If

' If RudyType = 2 Then
'   PrRudy.Visible = False
'   prRIWCage.Material = "Metal with an image"
'   prRudyScoop.Material = "Metal with an image"
'   PrMouthb.Material = "Metal with an image"
' End If

' If RudyType = 3 Then
'   PrRudy.Visible = True
'   PrLids.Image = "Rudy eyelid1c"
'   PrEyeL.Image = "eye_texture2"
'   PrEyeR.Image = "eye_texture2"
'   PrRudy.Image = "Rudy_Face_Off_2c"
'   PrRudy1.Image = "Rudy_Back_Off_2c"
'   PrMouth.Image = "Rudy mouth baked off c"
'   prRIWCage.Material = "Metal with an image Dark"
'   prRudyScoop.Material = "Metal with an image Dark"
'   PrMouthb.Material = "Metal with an image Dark"
' End If

If FlipperColor = 1 Then

  LeftFlipper.Material = "Plastic Yellow"
  LeftFlipper.RubberMaterial = "Blue Rubber"

  RightFlipper.Material = "Plastic Yellow"
  RightFlipper.RubberMaterial = "Blue Rubber"

  LeftFlipper1.Material = "Plastic Yellow"
  LeftFlipper1.RubberMaterial = "Blue Rubber"



Else

  LeftFlipper.Material = "Plastic Yellow"
  LeftFlipper.RubberMaterial = "Red Rubber"

  RightFlipper.Material = "Plastic Yellow"
  RightFlipper.RubberMaterial = "Red Rubber"

  LeftFlipper1.Material = "Plastic Yellow"
  LeftFlipper1.RubberMaterial = "Red Rubber"



End If

  If ClockMod = 1 Then
    pClock.Visible = True
  Else
    pClock.Visible = False

    cfs55.Opacity = 0
    cfs50.Opacity = 0
    cfs45.Opacity = 0
    cfs40.Opacity = 0
    cfs35.Opacity = 0
    cfs30.Opacity = 0
    cfs25.Opacity = 0
    cfs20.Opacity = 0
    cfs15.Opacity = 0
    cfs10.Opacity = 0
    cfs5.Opacity = 0

    cfh12.Opacity = 0
    cfh11.Opacity = 0
    cfh10.Opacity = 0
    cfh9.Opacity = 0
    cfh8.Opacity = 0
    cfh7.Opacity = 0
    cfh6.Opacity = 0
    cfh5.Opacity = 0
    cfh4.Opacity = 0
    cfh3.Opacity = 0
    cfh2.Opacity = 0
    cfh1.Opacity = 0

    l28b.state = 0
    cfcenter.Opacity = 0
        l28b.ShowBulbMesh = True
        l28b.ShowBulbMesh = False

  End If

  If BalloonMod = 1 Then
    prballoon_yellow.Visible = True
    prballoon_Red.Visible = True
    prballoon_Blue.Visible = True
    debug.print "On"
  Else
    prballoon_yellow.Visible = False
    prballoon_blue.Visible = False
    prballoon_red.Visible = False
    lbballoon.state=0
    lrballoon.state=0
    lyballoon.state=0
    debug.print "Off"
  End If

  If PopCornMod = 1 Then
    prPopCorn.Visible = True
    lPopcornLight.state = 1
  Else
    prPopCorn.Visible = False
    lPopcornLight.state = 0
  End If

  If HotDogCartMod = 1 Then
    prHotDogCartA.Visible = True
    prHotDogCartB.Visible = True
    prHotDogCartC.Visible = True
    lHotDogCartA.state = 1
  Else
    prHotDogCartA.Visible = False
    prHotDogCartB.Visible = False
    prHotDogCartC.Visible = False
    lHotDogCartA.state = 0
    lHotDogCartB.state = 0
  End If

  If LevelMod = 1 Then

    wall31.Visible = True
    wall29.Visible = True
    wall11.Visible = True
    level.Visible = True

  Else

    wall31.Visible = false
    wall29.Visible = false
    wall11.Visible = false
    level.Visible = false
  End If

  If BallMod = 1 Then
    cBall1.Image = "Chrome_Ball_29"
    cBall1.FrontDecal = "FunhouseBall1"
    cBall2.Image = "Chrome_Ball_29"
    cBall2.FrontDecal = "FunhouseBall2"
    cBall3.Image = "Chrome_Ball_29"
    cBall3.FrontDecal = "FunhouseBall3"
  Else
    cBall1.Image = "Pinball"
    cBall1.FrontDecal = "Scratches"
    cBall2.Image = "Pinball"
    cBall2.FrontDecal = "Scratches"
    cBall3.Image = "Pinball"
    cBall3.FrontDecal = "Scratches"
  End If

  If SubwayColorMod = 0 Then 'No color
    Up_subway_red.state=0
    Low_subway_red.state=0
    Up_subway_blue.state=0
    Low_subway_blue.state=0
    TD_subway_blue.state=0
    TD_subway_red.state=0
  ElseIf SubwayColorMod = 1 Then 'Blue color
    Up_subway_red.state = 0
    Low_subway_red.state = 0
    Up_subway_blue.state = 1
    Low_subway_blue.state = 1
    TD_subway_blue.state=1
    TD_subway_red.state=0
  ElseIf SubwayColorMod = 2 Then 'Red color
    Up_subway_red.state = 1
    Low_subway_red.state=1
    Up_subway_blue.state = 0
    Low_subway_blue.state = 0
    TD_subway_blue.state=0
    TD_subway_red.state=1
  End if

  If CardMod = 0 then
    CardType = CardMod
  Else
    CardType = Int(Rnd*6)
  End If

  If CardType = 0 Then
    pIC_Right.image="FH_IC1-R"
    pIC_Left.image="FH_IC1-L"
  End If

  If CardType = 1 Then
    pIC_Right.image="FH_IC2-R"
    pIC_Left.image="FH_IC2-L"
  End If

  If CardType = 2 Then
    pIC_Right.image="FH_IC3-R"
    pIC_Left.image="FH_IC3-L"
  End If

  If CardType = 3 Then
    pIC_Right.image="FH_IC4-R"
    pIC_Left.image="FH_IC4-L"
  End If

  If CardType = 4 Then
    pIC_Right.image="FH_IC5-R"
    pIC_Left.image="FH_IC5-L"
  End If

  If CardType = 5 Then
    pIC_Right.image="FH_IC6-R"
    pIC_Left.image="FH_IC6-L"
  End If

  If MirrorRWBMod = 1 Then
    L71.Material = "Lamps Glass Red"
    F71.Color = RGB(255,0,0)

    L74.Material = "Lamps Glass"
    F74.Color = RGB(255,255,255)

    L75.Material = "Lamps Glass"
    F75.Color = RGB(255,255,255)

    L76.Material = "Lamps Glass"
    F76.Color = RGB(255,255,255)

    L77.Material = "Lamps Glass"
    F77.Color = RGB(255,255,255)

    L78.Material = "Lamps Glass Blue"
    F78.Color = RGB(0,150,255)
  Else
    L71.Material = "Lamps Glass Red"
    F71.Color = RGB(255,0,0)

    L74.Material = "Lamps Glass Yellow"
    F74.Color = RGB(255,255,0)

    L75.Material = "Lamps Glass Yellow"
    F75.Color = RGB(255,255,0)

    L76.Material = "Lamps Glass Yellow"
    F76.Color = RGB(255,255,0)

    L77.Material = "Lamps Glass Yellow"
    F77.Color = RGB(255,255,0)

    L78.Material = "Lamps Glass Green"
    F78.Color = RGB(0,255,0)
  End If

End Sub

' _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _
'(_|_|_|_|_|_|_|_|_|_|_|_|_|_|_|_|_|_|_|_|_|_|_|_|_|_|_|_|_|_|_|_|_|_|_|_|_|_|_|_|_|_|_|_)

Sub CheckRudyType()
  If RudyType = 1 Then
    PrRudy.Visible = True
    PrLids.Image = "Rudy eyelid1"
    PrEyeL.Image = "eye_texture"
    PrEyeR.Image = "eye_texture"
    PrRudy.Image = "Rudy_Face_Off_2"
    PrRudy1.Image = "Rudy_Back_Off_2"
    PrMouth.Image = "Rudy mouth baked off"
    prRIWCage.Material = "Metal with an image Dark"
    prRudyScoop.Material = "Metal with an image Dark"
    PrMouthb.Material = "Metal with an image Dark"
  End If

  If RudyType = 2 Then
    PrRudy.Visible = False
    prRIWCage.Material = "Metal with an image"
    prRudyScoop.Material = "Metal with an image"
    PrMouthb.Material = "Metal with an image"
  End If

  If RudyType = 3 Then
    PrRudy.Visible = True
    PrLids.Image = "Rudy eyelid1c"
    PrEyeL.Image = "eye_texture2"
    PrEyeR.Image = "eye_texture2"
    PrRudy.Image = "Rudy_Face_Off_2c"
    PrRudy1.Image = "Rudy_Back_Off_2c"
    PrMouth.Image = "Rudy mouth baked off c"
    prRIWCage.Material = "Metal with an image Dark"
    prRudyScoop.Material = "Metal with an image Dark"
    PrMouthb.Material = "Metal with an image Dark"
  End If
End Sub

' _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _
'(_|_|_|_|_|_|_|_|_|_|_|_|_|_|_|_|_|_|_|_|_|_|_|_|_|_|_|_|_|_|_|_|_|_|_|_|_|_|_|_|_|_|_|_)


'*************************
'* Alpha-numeric display
'*************************

Dim Digits(32)
Digits(0) =  Array(LiScore1, LiScore2, LiScore3, LiScore4, LiScore5, LiScore6, LiScore7, LiScore8, LiScore9, LiScore10, LiScore11, LiScore12, LiScore13, LiScore14, LiScore15, LiScore16)
Digits(1) =  Array(LiScore17, LiScore18, LiScore19, LiScore20, LiScore21, LiScore22, LiScore23, LiScore24, LiScore25, LiScore26, LiScore27, LiScore28, LiScore29, LiScore30, LiScore31, LiScore32)
Digits(2) =  Array(LiScore33, LiScore34, LiScore35, LiScore36, LiScore37, LiScore38, LiScore39, LiScore40, LiScore41, LiScore42, LiScore43, LiScore44, LiScore45, LiScore46, LiScore47, LiScore48)
Digits(3) =  Array(LiScore49, LiScore50, LiScore52, LiScore53, LiScore54, LiScore55, LiScore56, LiScore63, LiScore57, LiScore58, LiScore59, LiScore51, LiScore60, LiScore61, LiScore62, LiScore64)
Digits(4) =  Array(LiScore65, LiScore66, LiScore68, LiScore69, LiScore70, LiScore71, LiScore72, LiScore79, LiScore73, LiScore74, LiScore75, LiScore67, LiScore76, LiScore77, LiScore78, LiScore80)
Digits(5) =  Array(LiScore81, LiScore82, LiScore84, LiScore85, LiScore86, LiScore87, LiScore88, LiScore95, LiScore89, LiScore90, LiScore91, LiScore83, LiScore92, LiScore93, LiScore94, LiScore96)
Digits(6) =  Array(LiScore97, LiScore98, LiScore100, LiScore101, LiScore102, LiScore103, LiScore104, LiScore111, LiScore105, LiScore106, LiScore107, LiScore99, LiScore108, LiScore109, LiScore110, LiScore112)
Digits(7) =  Array(LiScore113, LiScore114, LiScore116, LiScore117, LiScore118, LiScore119, LiScore120, LiScore127, LiScore121, LiScore122, LiScore123, LiScore115, LiScore124, LiScore125, LiScore126, LiScore128)
Digits(8) =  Array(LiScore129, LiScore130, LiScore132, LiScore133, LiScore134, LiScore135, LiScore136, LiScore143, LiScore137, LiScore138, LiScore139, LiScore131, LiScore140, LiScore141, LiScore142, LiScore144)
Digits(9) =  Array(LiScore145, LiScore146, LiScore148, LiScore149, LiScore150, LiScore151, LiScore152, LiScore159, LiScore153, LiScore154, LiScore155, LiScore147, LiScore156, LiScore157, LiScore158, LiScore160)
Digits(10) = Array(LiScore161, LiScore162, LiScore164, LiScore165, LiScore166, LiScore167, LiScore168, LiScore175, LiScore169, LiScore170, LiScore171, LiScore163, LiScore172, LiScore173, LiScore174, LiScore176)
Digits(11) = Array(LiScore177, LiScore178, LiScore180, LiScore181, LiScore182, LiScore183, LiScore184, LiScore191, LiScore185, LiScore186, LiScore187, LiScore179, LiScore188, LiScore189, LiScore190, LiScore192)
Digits(12) = Array(LiScore193, LiScore194, LiScore196, LiScore197, LiScore198, LiScore199, LiScore200, LiScore207, LiScore201, LiScore202, LiScore203, LiScore195, LiScore204, LiScore205, LiScore206, LiScore208)
Digits(13) = Array(LiScore209, LiScore210, LiScore212, LiScore213, LiScore214, LiScore215, LiScore216, LiScore223, LiScore217, LiScore218, LiScore219, LiScore211, LiScore220, LiScore221, LiScore222, LiScore224)
Digits(14) = Array(LiScore225, LiScore226, LiScore228, LiScore229, LiScore230, LiScore231, LiScore232, LiScore239, LiScore233, LiScore234, LiScore235, LiScore227, LiScore236, LiScore237, LiScore238, LiScore240)
Digits(15) = Array(LiScore241, LiScore242, LiScore244, LiScore245, LiScore246, LiScore247, LiScore248, LiScore255, LiScore249, LiScore250, LiScore251, LiScore243, LiScore252, LiScore253, LiScore254, LiScore256)
Digits(16) = Array(LiScore257, LiScore258, LiScore260, LiScore261, LiScore262, LiScore263, LiScore264, LiScore271, LiScore265, LiScore266, LiScore267, LiScore259, LiScore268, LiScore269, LiScore270, LiScore272)
Digits(17) = Array(LiScore273, LiScore274, LiScore276, LiScore277, LiScore278, LiScore279, LiScore280, LiScore287, LiScore281, LiScore282, LiScore283, LiScore275, LiScore284, LiScore285, LiScore286, LiScore288)
Digits(18) = Array(LiScore289, LiScore290, LiScore292, LiScore293, LiScore294, LiScore295, LiScore296, LiScore303, LiScore297, LiScore298, LiScore299, LiScore291, LiScore300, LiScore301, LiScore302, LiScore304)
Digits(19) = Array(LiScore305, LiScore306, LiScore308, LiScore309, LiScore310, LiScore311, LiScore312, LiScore319, LiScore313, LiScore314, LiScore315, LiScore307, LiScore316, LiScore317, LiScore318, LiScore320)
Digits(20) = Array(LiScore321, LiScore322, LiScore324, LiScore325, LiScore326, LiScore327, LiScore328, LiScore335, LiScore329, LiScore330, LiScore331, LiScore323, LiScore332, LiScore333, LiScore334, LiScore336)
Digits(21) = Array(LiScore337, LiScore338, LiScore340, LiScore341, LiScore342, LiScore343, LiScore344, LiScore351, LiScore345, LiScore346, LiScore347, LiScore339, LiScore348, LiScore349, LiScore350, LiScore352)
Digits(22) = Array(LiScore353, LiScore354, LiScore356, LiScore357, LiScore358, LiScore359, LiScore360, LiScore367, LiScore361, LiScore362, LiScore363, LiScore355, LiScore364, LiScore365, LiScore366, LiScore368)
Digits(23) = Array(LiScore369, LiScore370, LiScore372, LiScore373, LiScore374, LiScore375, LiScore376, LiScore383, LiScore377, LiScore378, LiScore379, LiScore371, LiScore380, LiScore381, LiScore382, LiScore384)
Digits(24) = Array(LiScore385, LiScore386, LiScore388, LiScore389, LiScore390, LiScore391, LiScore392, LiScore399, LiScore393, LiScore394, LiScore395, LiScore387, LiScore396, LiScore397, LiScore398, LiScore400)
Digits(25) = Array(LiScore401, LiScore402, LiScore404, LiScore405, LiScore406, LiScore407, LiScore408, LiScore415, LiScore409, LiScore410, LiScore411, LiScore403, LiScore412, LiScore413, LiScore414, LiScore416)
Digits(26) = Array(LiScore417, LiScore418, LiScore420, LiScore421, LiScore422, LiScore423, LiScore424, LiScore431, LiScore425, LiScore426, LiScore427, LiScore419, LiScore428, LiScore429, LiScore430, LiScore432)
Digits(27) = Array(LiScore433, LiScore434, LiScore436, LiScore437, LiScore438, LiScore439, LiScore440, LiScore447, LiScore441, LiScore442, LiScore443, LiScore435, LiScore444, LiScore445, LiScore446, LiScore448)
Digits(28) = Array(LiScore449, LiScore450, LiScore452, LiScore453, LiScore454, LiScore455, LiScore456, LiScore463, LiScore457, LiScore458, LiScore459, LiScore451, LiScore460, LiScore461, LiScore462, LiScore464)
Digits(29) = Array(LiScore465, LiScore466, LiScore468, LiScore469, LiScore470, LiScore471, LiScore472, LiScore479, LiScore473, LiScore474, LiScore475, LiScore467, LiScore476, LiScore477, LiScore478, LiScore480)
Digits(30) = Array(LiScore481, LiScore482, LiScore484, LiScore485, LiScore486, LiScore487, LiScore488, LiScore495, LiScore489, LiScore490, LiScore491, LiScore483, LiScore492, LiScore493, LiScore494, LiScore496)
Digits(31) = Array(LiScore497, LiScore498, LiScore500, LiScore501, LiScore502, LiScore503, LiScore504, LiScore511, LiScore505, LiScore506, LiScore507, LiScore499, LiScore508, LiScore509, LiScore510, LiScore512)


Sub TiDisplay_Timer()
  Dim ChgLED, ii, num, chg, stat, obj
  ChgLED=Controller.ChangedLEDs(&H00000000, &Hffffffff)
  If Not IsEmpty(ChgLED) Then
    If DesktopMode = True Then
    For ii=0 To UBound(chgLED)
      num=chgLED(ii,0)
      chg=chgLED(ii,1)
      stat=chgLED(ii,2)
      For Each obj In Digits(num)
        If chg And 1 Then obj.State=stat And 1
        chg=chg\2
        stat=stat\2
      Next
    Next
     end if
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
'    JP's VP10 Rolling Sounds
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

Sub RollingSoundTimer_Timer()
    Dim BOT, b
    BOT = GetBalls
'End sub

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


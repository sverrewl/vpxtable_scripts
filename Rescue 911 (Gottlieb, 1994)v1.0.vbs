Option Explicit
Randomize

'*****************************************************************************************************

' $$$$$$$\                                                           $$$$$$\    $$\     $$\
' $$  __$$\                                                         $$  __$$\ $$$$ |  $$$$ |
' $$ |  $$ | $$$$$$\   $$$$$$$\  $$$$$$$\ $$\   $$\  $$$$$$\        $$ /  $$ |\_$$ |  \_$$ |
' $$$$$$$  |$$  __$$\ $$  _____|$$  _____|$$ |  $$ |$$  __$$\       \$$$$$$$ |  $$ |    $$ |
' $$  __$$< $$$$$$$$ |\$$$$$$\  $$ /      $$ |  $$ |$$$$$$$$ |       \____$$ |  $$ |    $$ |
' $$ |  $$ |$$   ____| \____$$\ $$ |      $$ |  $$ |$$   ____|      $$\   $$ |  $$ |    $$ |
' $$ |  $$ |\$$$$$$$\ $$$$$$$  |\$$$$$$$\ \$$$$$$  |\$$$$$$$\       \$$$$$$  |$$$$$$\ $$$$$$\
' \__|  \__| \_______|\_______/  \_______| \______/  \_______|       \______/ \______|\______|

'**************************************** Gottlieb, 1994 **********************************************


'*****************************************************************************************************
' CREDITS
' Rescue 911 by Antisect 2021
' Original Helicopter script by Dozer316 & DJRobX
' Ball rolling sound script by jpsalas
' Ball shadow by ninuzzu
' Ball dropping sound by rothbauerw
' DOF by arngrim
' Plus a lot of input from the whole community (sorry if we forgot you :/)
'*****************************************************************************************************


'First, try to load the Controller.vbs (DOF), which helps controlling additional hardware like lights, gears, knockers, bells and chimes (to increase realism)
'This table uses DOF via the 'SoundFX' calls that are inserted in some of the PlaySound commands, which will then fire an additional event, instead of just playing a sample/sound effect
On Error Resume Next
ExecuteGlobal GetTextFile("controller.vbs")
If Err Then MsgBox "You need the Controller.vbs file in order to run this table (installed with the VPX package in the scripts folder)"
On Error Goto 0

'If using Visual PinMAME (VPM), place the ROM/game name in the constant below,
'both for VPM, and DOF to load the right DOF config from the Configtool, whether it's a VPM or an Original table

Dim UseVPMDMD


If Table1.ShowDT = false then
  UseVPMDMD = 0
  else
  UseVPMDMD = 1
End If

Dim Ballsize,BallMass
BallSize = 50
BallMass = 1
Const cGameName = "rescu911"
Const UseSolenoids=2,UseLamps=0,UseGI=0


LoadVPM "01560000", "GTS3.VBS", 3.26

'Dont use Table1.width or Table1.height in script as it can effect performance
dim tablewidth: tablewidth = Table1.width
dim tableheight: tableheight = Table1.height

'switches
'switch 0-7 handled by pinmame
'8-9 not used
Const swLeftCoinShute = 0
Const swRightCoinShute = 1
Const swCenterCoinChute = 2
Const swCoinChute = 3
Const swStartButton = 4
Const swTournament = 5

Const sElectroMagnet = 13
Const sBallLiftMotor = 22
Const sHeliBladeMotor = 23
Const sHeliARMmotorRelay = 24
Const sHeliDirectionalRelay = 25
Const stiltrelay = 31
Const sGameOverRelay = 32

'****************************************************************************************************************

'  $$$$$$$$\        $$\       $$\                  $$$$$$\             $$\     $$\
'  \__$$  __|       $$ |      $$ |                $$  __$$\            $$ |    \__|
'     $$ | $$$$$$\  $$$$$$$\  $$ | $$$$$$\        $$ /  $$ | $$$$$$\ $$$$$$\   $$\  $$$$$$\  $$$$$$$\   $$$$$$$\
'     $$ | \____$$\ $$  __$$\ $$ |$$  __$$\       $$ |  $$ |$$  __$$\\_$$  _|  $$ |$$  __$$\ $$  __$$\ $$  _____|
'     $$ | $$$$$$$ |$$ |  $$ |$$ |$$$$$$$$ |      $$ |  $$ |$$ /  $$ | $$ |    $$ |$$ /  $$ |$$ |  $$ |\$$$$$$\
'     $$ |$$  __$$ |$$ |  $$ |$$ |$$   ____|      $$ |  $$ |$$ |  $$ | $$ |$$\ $$ |$$ |  $$ |$$ |  $$ | \____$$\
'     $$ |\$$$$$$$ |$$$$$$$  |$$ |\$$$$$$$\        $$$$$$  |$$$$$$$  | \$$$$  |$$ |\$$$$$$  |$$ |  $$ |$$$$$$$  |
'     \__| \_______|\_______/ \__| \_______|       \______/ $$  ____/   \____/ \__| \______/ \__|  \__|\_______/
'                                                           $$ |
'                                                           $$ |
'                                                           \__|



'///////////////////////-----General Sound Options-----///////////////////////
'// VolumeDial:
'// VolumeDial is the actual global volume multiplier for the mechanical sounds.
'// Values smaller than 1 will decrease mechanical sounds volume.
'// Recommended values should be no greater than 1.
Const VolumeDial = 0.5

'////////////////////////////////////////////////////////////////////////////
'Render the Mars Bulb light that sits on top of the real machine's
'bacbox over the top 1/4 of the playfield
Const marson = 1


'///////////////////////-----Debug-----///////////////////////
'// Look under the hood
'// 1 = on 0 = off
Const  aprondebug = 0

'///////////////////////-----Authentic Flashers-----///////////////////////
'// Authentic Flashers (more apron flashes when set to off)
'// 1 = on 0 = off
Const  Flashon = 0

'****************************************************************************************************************


'******************************************************
'       Solenoids
'******************************************************

SolCallback(sLRFlipper) = "SolRFlipper"
SolCallback(sLLFlipper) = "SolLFlipper"
SolCallback(sURFlipper) = "SolURFlipper"
'SolCallback(1) = "vpmSolSound ""jet1"","
'SolCallback(2) = "vpmSolSound ""jet1"","
'SolCallback(3) = "vpmSolSound ""Jet1"","
'SolCallback(4) = "vpmSolSound ""lSling"","
'SolCallback(5) = "vpmSolSound ""lSling"","
'SolCallback(6) = "vpmSolSound ""lSling"","
'SolCallback(7) = "vpmSolSound ""lSling"","
SolCallback(8) = "dtLDrop.SolDropUp"
SolCallback(9) = "BallLeftShooter"
SolCallback(10) = "VukTopPop"
SolCallback(11) = "Tpost"
SolCallback(12) = "LockRelease"
SolCallback(13) = "ElectroMagnet"
SolCallback(14) = "SetLamp 114,"  'Left Lane and Heli Scoop Flashers
SolCallBack(15) = "SetLamp 115,"  'Right Lane and Car Flashers
SolCallBack(16) = "SetLamp 116,"  'Searchlight Flasher
SolCallBack(17) = "Sol17" 'Yellow Flasher
SolCallBack(18) = "Sol18" 'Top Right Flasher Red
SolCallBack(19) = "Sol19" 'Top Right Small Dome Flasher Red
SolCallBack(20) = "SetLamp 130,"  'Cave Flasher
SolCallback(21) = "MBulb" 'Rotating Red Light
SolCallback(22) = "BallLiftMotor"
SolCallback(23) = "Blades_Motor"
SolCallback(24) = "HeliArm"
SolCallback(25) = "HeliArmRelay"
SolCallback(26)  = "GI_Update"
'SolCallback(27)  = "vpmSolSound ""solenoid"","
SolCallback(28) = "SolReleaseBall"  'Ball Release
SolCallback(29) = "SolOutHole"    'Outhole
SolCallback(30) = "vpmSolSound SoundFX(""knocker"",DOFKnocker),"
SolCallback(31) = "SolTilt"
SolCallback(32) = "GameOn"

'******************************************************
'       Table init
'******************************************************

Dim vlLock, bsVUK, cbCaptive, dtLDrop, bsTrough

Sub Table1_Init
    vpmInit Me
    With Controller
        .GameName = cGameName
        If Err Then MsgBox "Can't start Game " & cGameName & vbNewLine & Err.Description:Exit Sub
        .SplashInfoLine = "" & vbNewLine & ""
    .HandleKeyboard = 0
    .ShowTitle = 0
    .ShowDMDOnly = 1
    .ShowFrame = 0
    .HandleMechanics = 0
    .Hidden = 0
    .Games(cGameName).Settings.Value("sound") = 1
    On Error Resume Next
    .Run GetPlayerHWnd
    If Err Then MsgBox Err.Description
  End With
    On Error Goto 0

'/////////////////////////////  PinMAME  ////////////////////////////

  PinMAMETimer.Interval=PinMAMEInterval
  PinMAMETimer.Enabled=1
    RealTime.Enabled = 1

'/////////////////////////////  Tilt  ////////////////////////////

    vpmNudge.TiltSwitch = 151
    vpmNudge.Sensitivity = 0.25
    vpmNudge.TiltObj = Array(bumper1, bumper2, bumper3, LeftSlingshot, RightSlingshot)

'/////////////////////////////  Trough  ////////////////////////////

    Set bsTrough = New cvpmTrough
    With bsTrough
    .EntrySw  = 23
        .size = 3
        .initSwitches Array(201,200,33)
        .Initexit BallRelease, 55, 10
        .Balls = 3
    End With
  DrainFill.CreateSizedballWithMass Ballsize/2,Ballmass
  DrainFill.kick 225, 0.2
  TroughGuard.isdropped = 1

'/////////////////////////////  Visible Lock  ////////////////////////////

    Set vlLock = New cvpmVLock
    With vllock
        .InitVLock Array(RL1, RL2), Array(RLK1, RLK2), 0
        .InitSnd SoundFX("flapopen", DOFContactors), SoundFX("solenoid", DOFContactors)
    .ExitForce = 0
        .CreateEvents "vlLock"
    End With

'/////////////////////////////  Captive Balls  ////////////////////////////

  Set cbCaptive = New cvpmCaptiveBall : With cbCaptive
        .InitCaptive CaptiveTrigger1, CaptiveWall, Array(Captive1), 360
    .CreateEvents "cbCaptive"
        .Start
    .ForceTrans = 0.6
    '.MinForce = 3.5
    End With

  Captive2.CreateBall.Image = "JPBall-Dark2"
  Captive2.kick 180, 1
  Captive2.enabled = 0
  Captive3.CreateBall.Image = "JPBall-Dark2"
  Captive3.kick 180, 1
  Captive3.enabled = 0

'/////////////////////////////  Drop Targets  ////////////////////////////

  set dtLDrop = New cvpmDropTarget
  dtlDrop.InitDrop Array(sw20, sw30), Array(20, 30)
    dtlDrop.InitSnd SoundFX("target_drop", DOFContactors), SoundFX("DTReset", DOFContactors)

  Controller.Switch(34) = 1 'ball lift down
    PUBlock.isdropped = 1

'/////////////////////////////  Debug  ////////////////////////////

  If aprondebug = 1 Then
    Primitive13.visible=0
  Else
    Primitive13.visible=1
  End If

  Dim DesktopMode: DesktopMode = Table1.ShowDT

  If DesktopMode = True Then 'Show Desktop components
    PlasticTop001.Visible=1
    Primitive017.visible=1
  Else
    PlasticTop001.Visible=0
    Primitive017.visible=0
  End if

'/////////////////////////////  Load LUT  ////////////////////////////
  TextBox001.visible = 0
  LoadLUT

End Sub

'******************
' RealTime Updates
'******************

Sub RealTime_Timer
    RollingUpdate
End Sub

'******************************************************
'       Game-On
'******************************************************

Dim gon, ton

gon = 0
ton = 0

Sub GameOn(enabled)
    If enabled Then
    GO.State = 1
    gon = 1
    Else
    GO.State = 0
        gon = 0
    End If
End Sub

Sub SolTilt(enabled)
    If enabled Then
    XTO2.State = 1
  Else
    XTO2.State = 0
    End If
End Sub

'******************************************************
'       DRAIN & RELEASE
'******************************************************

Sub SWTroughGuard_Hit()
  DrainFill.Enabled = 0
  TroughGuard.isdropped = 0
End Sub

Sub Drain_Hit()
  RandomSoundDrain Drain
  bsTrough.AddBall Me
End Sub

Sub SolReleaseBall(enabled)
  If enabled Then
    RandomSoundBallRelease BallRelease
    bsTrough.ExitSol_On
    'debug.print bsTrough.BallsPending
  End If
End Sub

Sub SolOutHole(enabled)
  If enabled Then
    PlaySound "Saucer_Kick"
    bsTrough.EntrySol_On
    'debug.print bsTrough.BallsPending
  End If
End Sub

'******************************************************
'       Mars Flasher
'******************************************************

Flasher_M.visible = 0

Sub Mbulb(enabled)
    If enabled And marson = 1 Then
        Flasher_M.visible = 1
        SpinBulb.enabled = 1
    Else
        Flasher_M.visible = 0
        SpinBulb.enabled = 0
    End If
End Sub

Sub SpinBulb_Timer()
    Flasher_M.RotZ = Flasher_M.RotZ + 1
End Sub

'******************************************************
'       Switches
'******************************************************

dim BIPL

Sub sw21_Hit:Controller.Switch(21)=1 : RandomSoundRollover:BIPL = 1: End Sub
Sub sw21_unHit:Controller.Switch(21)=0 : RandomSoundRollover:BIPL = 0:End Sub
Sub sw22_Hit : vpmTimer.PulseSw 22 : End Sub
Sub sw22a_Hit : vpmTimer.PulseSw 22 : End Sub

Sub sw32_Hit:vpmTimer.PulseSw 32: End Sub

Sub sw92_Hit:Controller.Switch(92)=1 : End Sub
Sub sw92_unHit:Controller.Switch(92)=0:End Sub
Sub sw93_Hit:Controller.Switch(93)=1 : End Sub
Sub sw93_unHit:Controller.Switch(93)=0:End Sub
Sub swA2_Hit:Controller.Switch(102)=1 : End Sub
Sub swA2_unHit:Controller.Switch(102)=0:End Sub
Sub swA3_Hit:Controller.Switch(103)=1 : End Sub
Sub swA3_unHit:Controller.Switch(103)=0:End Sub

Sub sw81_Hit:vpmTimer.PulseSw 81: End Sub


Sub sw101_Hit : Controller.Switch(101) = 1 : End Sub
Sub sw101_UnHit : Controller.Switch(101) = 0 : End Sub
Sub sw111_Hit : Controller.Switch(111) = 1 : End Sub
Sub sw111_UnHit : Controller.Switch(111) = 0 : End Sub

Sub sw113_Hit : Controller.Switch(113) = 1 : End Sub
Sub sw113_UnHit : Controller.Switch(113) = 0 : End Sub

Sub T15_Slingshot:vpmTimer.PulseSw 15: T912.Enabled=1: RandomSoundSlingshotLeft T911Sound002 : End Sub
Sub T16_Slingshot:vpmTimer.PulseSw 16: T911.Enabled=1: RandomSoundSlingshotRight T911Sound001 : End Sub
Sub T31_Hit:vpmTimer.PulseSw 31: End Sub
Sub T85_Hit:vpmTimer.PulseSw 85: End Sub
Sub T94_Hit:vpmTimer.PulseSw 94: End Sub
Sub T95_Hit:vpmTimer.PulseSw 95: End Sub
Sub T96_Hit:vpmTimer.PulseSw 96: End Sub
Sub T97_Hit:vpmTimer.PulseSw 97: End Sub
Sub T114_Hit:vpmTimer.PulseSw 114: End Sub
Sub T104_Hit:vpmTimer.PulseSw 104: End Sub
Sub T105_Hit:vpmTimer.PulseSw 105: End Sub
Sub T106_Hit:vpmTimer.PulseSw 106: End Sub
Sub T107_Hit:vpmTimer.PulseSw 107: End Sub
Sub T115_Hit:vpmTimer.PulseSw 115: End Sub
Sub T117_Hit:vpmTimer.PulseSw 117: End Sub

'/////////////////////////////  Bumpers  ////////////////////////////

Sub Bumper1_Hit()
    RandomSoundBumperMiddle Bumper1
    vpmTimer.PulseSw 10
End Sub

Sub Bumper2_Hit()
    RandomSoundBumperTop Bumper2
    vpmTimer.PulseSw 12
End Sub

Sub Bumper3_Hit()
    RandomSoundBumperBottom Bumper3
    vpmTimer.PulseSw 11
End Sub

'/////////////////////////////  Drop Targets  ////////////////////////////

Sub sw20_Hit() : dtlDrop.Hit 1: sw20.IsDropped = 1 : PlaySound SoundFX("Target_Drop", DOFContactors) : End Sub
Sub sw30_Hit() : dtlDrop.Hit 2: sw30.Isdropped = 1 : PlaySound SoundFX("Target_Drop", DOFContactors) : End Sub


'/////////////////////////////  911 Targets  ////////////////////////////

dim thit, thit2: thit=0: thit2=0
dim hitreturn, hitreturn2: hitreturn=0: hitreturn2=0

Sub T911_timer()

  if hitreturn = 0 then
    thit = thit + 3
  end if

  if hitreturn = 1 then
    thit = thit - 6

    if thit = 0 then
      hitreturn = 0
      thit=0
      T911.Enabled=0
    end if

  end if

  T9111.x = thit
  T9111.y = -thit

  if thit > 21 then
    hitreturn = 1
  end if

End Sub

Sub T912_timer()

  if hitreturn2 = 0 then
    thit2 = thit2 - 3
  end if

  if hitreturn2 = 1 then
    thit2 = thit2 + 6

    if thit2 => 0 then
      hitreturn2 = 0
      thit2=0
      T912.Enabled=0
    end if

  end if

  T9112.x = thit2
  T9112.y = thit2

  if thit2 < -21 then
    hitreturn2 = 1
  end if

End Sub

'******************************************************
'         Ramp Accelerators
'******************************************************

Sub swSC01_Hit
    activeball.velY=activeball.velY+5
    activeball.velX=activeball.velX+3
End Sub

'******************************************************
'           Vertical Kick
'******************************************************

 Dim raiseballsw, raiseball

 Sub TopVUK_Hit()
  TopVUK.Enabled=FALSE
  Controller.switch (91) = True
  PlaySound "Hole"
  'playsound "popper_ball"
 End Sub

 Sub VukTopPop(enabled)
  if(enabled and Controller.switch (91)) then
    TopVUK.DestroyBall
    Set raiseball = TopVUK.CreateBall
        PlaySound SoundFX("popper_ball",DOFContactors)
    playsound SoundFX("vukout",DOFContactors)
    raiseballsw = True
    TopVukraiseball.Enabled = True 'Added by Rascal
    TopVUK.Enabled=TRUE
    Controller.switch (91) = False
  else

  end if
End Sub

 Sub TopVukraiseball_Timer()
  If raiseballsw = True then
    raiseball.z = raiseball.z + 6   '12
    'raiseball.x = raiseball.x + 0
    If raiseball.z > 126 then
      'playsound "vukout"
      TopVUK.Kick 135, 14
      Set raiseball = Nothing
      TopVukraiseball.Enabled = False
      raiseballsw = False
    End If
  End If
 End Sub

'/////////////////////////////  Subways  ////////////////////////////

Sub Lock_Drop_Hit()
    Playsound "Hole"
    Playsound "subway"
    vpmTimer.PulseSw 110
End Sub

Sub Heli_Mid_Drop_Hit()
    Playsound "Hole"
    Playsound "subway"
    vpmTimer.PulseSw 110
End Sub

'/////////////////////////////  TPost  ////////////////////////////

Sub Tpost(enabled)
    If enabled Then
        TopPost.isdropped = 1
        PlaySound SoundFX("solenoid", DOFContactors)
Else
        TopPost.isdropped = 0
        PlaySound SoundFX("solenoid", DOFContactors)
    End If
End Sub

'******************************************************
'           Ball Lock
'******************************************************

Sub LockRelease(Enabled)
  If Enabled Then
    vlLock.SolExit True
    TopPost2.isdropped = 1
    PlaySound SoundFX("solenoid", DOFContactors)
  Else
    vlLock.SolExit False
    PlaySound SoundFX("solenoid", DOFContactors)
    TopPost2.isdropped = 0
  End If
End Sub

'******************************************************
'           Keyboard
'******************************************************

Sub Table1_KeyDown(ByVal keycode)
  If keycode = LeftFlipperKey Then FlipperActivate LeftFlipper, LFPress
  If keycode = RightFlipperKey Then FlipperActivate RightFlipper, RFPress
  If KeyCode = PlungerKey Then Plunger.PullBack:SoundPlungerPull()

  If keycode = LeftTiltKey Then Nudge 90, 3 : SoundNudgeLeft()
  If keycode = RightTiltKey Then Nudge 270, 3 : SoundNudgeRight()
  If keycode = CenterTiltKey Then Nudge 0, 2 : SoundNudgeCenter()

  If keycode = LeftMagnaSave Then bLutActive = True
  If keycode = RightMagnaSave Then
    textbox001.visible = 1
    If bLutActive Then NextLUT: End If
    dim tekst : tekst = "LUT No#:" & LUTImage
    textbox001.text = tekst
    vpmTimer.AddTimer 2000, "If textbox001.text =" + chr(34) + tekst + chr(34) + " then textbox001.visible = 0'"
  End If
  If keycode = keyInsertCoin1 or keycode = keyInsertCoin2 or keycode = keyInsertCoin3 or keycode = keyInsertCoin4 Then
    Select Case Int(rnd*3)
      Case 0: PlaySound ("Coin_In_1"), 0, CoinSoundLevel, 0, 0.25
      Case 1: PlaySound ("Coin_In_2"), 0, CoinSoundLevel, 0, 0.25
      Case 2: PlaySound ("Coin_In_3"), 0, CoinSoundLevel, 0, 0.25
    End Select
  End If
  if keycode=StartGameKey then soundStartButton()
  If vpmKeyDown(keycode) Then Exit Sub
  If KeyDownHandler(keycode) Then Exit Sub
End Sub

Sub Table1_KeyUp(ByVal keycode)
  If keycode = LeftFlipperKey Then FlipperDeActivate LeftFlipper, LFPress
  If keycode = RightFlipperKey Then FlipperDeActivate RightFlipper, RFPress
  If KeyCode = PlungerKey Then
    Plunger.Fire
    If BIPL = 1 Then
      SoundPlungerReleaseBall()     'Plunger release sound when there is a ball in shooter lane
    Else
      SoundPlungerReleaseNoBall()     'Plunger release sound when there is no ball in shooter lane
    End If
  End If
  If keycode = LeftMagnaSave Then bLutActive = False
    If vpmKeyUp(keycode) Then Exit Sub
  If KeyUpHandler(keycode) Then Exit Sub
End Sub

'*********
'   LUT
'*********

Dim bLutActive, LUTImage
Sub LoadLUT
  dim x
  bLutActive = False
    x = LoadValue(cGameName, "LUTImage")
    If(x <> "") Then LUTImage = x Else LUTImage = 0
  UpdateLUT
End Sub

Sub SaveLUT
    SaveValue cGameName, "LUTImage", LUTImage
End Sub

Sub NextLUT: LUTImage = (LUTImage +1 ) MOD 12: UpdateLUT: SaveLUT: End Sub

Sub UpdateLUT
Select Case LutImage
Case 0: table1.ColorGradeImage = "LUT0"
Case 1: table1.ColorGradeImage = "LUT1"
Case 2: table1.ColorGradeImage = "LUT2"
Case 3: table1.ColorGradeImage = "LUT3"
Case 4: table1.ColorGradeImage = "LUT4"
Case 5: table1.ColorGradeImage = "LUT5"
Case 6: table1.ColorGradeImage = "LUT6"
Case 7: table1.ColorGradeImage = "LUT7"
Case 8: table1.ColorGradeImage = "LUT8"
Case 9: table1.ColorGradeImage = "LUT9"
Case 10: table1.ColorGradeImage = "LUT10"
Case 11: table1.ColorGradeImage = "LUT11"
End Select
End Sub

' #####################################
' ######    Flippers      #####
' #####################################

Const ReflipAngle = 20

Sub SolLFlipper(Enabled)
  If Enabled Then
    LF.Fire  'leftflipper.rotatetoend

    If leftflipper.currentangle < leftflipper.endangle + ReflipAngle Then
      RandomSoundReflipUpLeft LeftFlipper
    Else
      SoundFlipperUpAttackLeft LeftFlipper
      RandomSoundFlipperUpLeft LeftFlipper
    End If
  Else
    LeftFlipper.RotateToStart
    If LeftFlipper.currentangle < LeftFlipper.startAngle - 5 Then
      RandomSoundFlipperDownLeft LeftFlipper
    End If
    FlipperLeftHitParm = FlipperUpSoundLevel
    End If
End Sub

Sub SolRFlipper(Enabled)
  If Enabled Then
    RF.Fire 'rightflipper.rotatetoend

    If rightflipper.currentangle > rightflipper.endangle - ReflipAngle Then
      RandomSoundReflipUpRight RightFlipper
    Else
      SoundFlipperUpAttackRight RightFlipper
      RandomSoundFlipperUpRight RightFlipper
    End If
  Else
    RightFlipper.RotateToStart
    If RightFlipper.currentangle > RightFlipper.startAngle + 5 Then
      RandomSoundFlipperDownRight RightFlipper
    End If
    FlipperRightHitParm = FlipperUpSoundLevel
  End If
End Sub

Sub SolURFlipper(Enabled)
  If Enabled Then
'   PlaySound "fx_flipperup", 0, 0.5, -0.06, 0.15
    RightFlipper002.RotateToEnd
  Else
'   PlaySound "fx_flipperdown", 0, 0.5, -0.06, 0.15
    RightFlipper002.RotateToStart
  End If
End Sub

Sub LeftFlipper_Collide(parm)
  LeftFlipperCollide parm
  CheckLiveCatch Activeball, LeftFlipper, LFCount, parm
End Sub

Sub RightFlipper_Collide(parm)
  RightFlipperCollide parm
  CheckLiveCatch Activeball, RightFlipper, RFCount, parm
End Sub

'**************
' Flasher Subs
'**************

Sub Sol19(flstate)
  If Flstate Then
    Objlevel(1) = 1 : FlasherFlash1_Timer
  End If
End Sub

Sub Sol18(flstate)
  If Flstate Then
    Objlevel(2) = 1 : FlasherFlash2_Timer
  End If
End Sub

Sub Sol17(flstate)
  If Flstate Then
    Objlevel(3) = 1 : FlasherFlash3_Timer
  End If
End Sub

' #####################################
' ###### Flupper1 Flashers #####
' #####################################

Dim TestFlashers, TableRef, FlasherLightIntensity, FlasherFlareIntensity, FlasherOffBrightness

                ' *********************************************************************
TestFlashers = 0        ' *** set this to 1 to check position of flasher object       ***
Set TableRef = Table1       ' *** change this, if your table has another name             ***
FlasherLightIntensity = 0.3   ' *** lower this, if the VPX lights are too bright (i.e. 0.1)   ***
FlasherFlareIntensity = 0.2     ' *** lower this, if the flares are too bright (i.e. 0.1)     ***
FlasherOffBrightness = 0.3    ' *** brightness of the flasher dome when switched off (range 0-2)  ***
                ' *********************************************************************

Dim ObjLevel(20), objbase(20), objlit(20), objflasher(20), objlight(20)

'initialise the flasher color, you can only choose from "green", "red", "purple", "blue", "white" and "yellow"
InitFlasher 1, "red" : InitFlasher 2, "red" : : InitFlasher 3, "yellow"

' rotate the flasher with the command below (first argument = flasher nr, second argument = angle in degrees)
'RotateFlasher 4,17 : RotateFlasher 5,0 : RotateFlasher 6,90
RotateFlasher 1,-32

Sub InitFlasher(nr, col)
  ' store all objects in an array for use in FlashFlasher subroutine
  Set objbase(nr) = Eval("Flasherbase" & nr) : Set objlit(nr) = Eval("Flasherlit" & nr)
  Set objflasher(nr) = Eval("Flasherflash" & nr) : Set objlight(nr) = Eval("Flasherlight" & nr)
  ' If the flasher is parallel to the playfield, rotate the VPX flasher object for POV and place it at the correct height
  If objbase(nr).RotY = 0 Then
    objbase(nr).ObjRotZ =  atn( (tablewidth/2 - objbase(nr).x) / (objbase(nr).y - tableheight*1.1)) * 180 / 3.14159
    objflasher(nr).RotZ = objbase(nr).ObjRotZ : objflasher(nr).height = objbase(nr).z + 60
  End If
  ' set all effects to invisible and move the lit primitive at the same position and rotation as the base primitive
  objlight(nr).IntensityScale = 0 : objlit(nr).visible = 0 : objlit(nr).material = "Flashermaterial" & nr
  objlit(nr).RotX = objbase(nr).RotX : objlit(nr).RotY = objbase(nr).RotY : objlit(nr).RotZ = objbase(nr).RotZ
  objlit(nr).ObjRotX = objbase(nr).ObjRotX : objlit(nr).ObjRotY = objbase(nr).ObjRotY : objlit(nr).ObjRotZ = objbase(nr).ObjRotZ
  objlit(nr).x = objbase(nr).x : objlit(nr).y = objbase(nr).y : objlit(nr).z = objbase(nr).z
  objbase(nr).BlendDisableLighting = FlasherOffBrightness
  ' set the texture and color of all objects
  select case objbase(nr).image
    Case "dome2basewhite" : objbase(nr).image = "dome2base" & col : objlit(nr).image = "dome2lit" & col :
    Case "ronddomebasewhite" : objbase(nr).image = "ronddomebase" & col : objlit(nr).image = "ronddomelit" & col
    Case "domeearbasewhite" : objbase(nr).image = "domeearbase" & col : objlit(nr).image = "domeearlit" & col
  end select
  If TestFlashers = 0 Then objflasher(nr).imageA = "domeflashwhite" : objflasher(nr).visible = 0 : End If
  select case col
    Case "blue" :   objlight(nr).color = RGB(4,120,255) : objflasher(nr).color = RGB(200,255,255) : objlight(nr).intensity = 5000
    Case "green" :  objlight(nr).color = RGB(12,255,4) : objflasher(nr).color = RGB(12,255,4)
    Case "red" :    objlight(nr).color = RGB(255,32,4) : objflasher(nr).color = RGB(255,32,4)
    Case "purple" : objlight(nr).color = RGB(230,49,255) : objflasher(nr).color = RGB(255,64,255)
    Case "yellow" : objlight(nr).color = RGB(200,173,25) : objflasher(nr).color = RGB(255,200,50)
    Case "white" :  objlight(nr).color = RGB(255,240,150) : objflasher(nr).color = RGB(100,86,59)
  end select
  objlight(nr).colorfull = objlight(nr).color
  If TableRef.ShowDT and ObjFlasher(nr).RotX = -45 Then
    objflasher(nr).height = objflasher(nr).height - 20 * ObjFlasher(nr).y / tableheight
    ObjFlasher(nr).y = ObjFlasher(nr).y + 10
  End If
End Sub

Sub RotateFlasher(nr, angle) : angle = ((angle + 360 - objbase(nr).ObjRotZ) mod 180)/30 : objbase(nr).showframe(angle) : objlit(nr).showframe(angle) : End Sub

Sub FlashFlasher(nr)
  If not objflasher(nr).TimerEnabled Then objflasher(nr).TimerEnabled = True : objflasher(nr).visible = 1 : objlit(nr).visible = 1 : End If
  objflasher(nr).opacity = 1000 *  FlasherFlareIntensity * ObjLevel(nr)^2.5
  objlight(nr).IntensityScale = 0.5 * FlasherLightIntensity * ObjLevel(nr)^3
  objbase(nr).BlendDisableLighting =  FlasherOffBrightness + 10 * ObjLevel(nr)^3
  objlit(nr).BlendDisableLighting = 10 * ObjLevel(nr)^2
  UpdateMaterial "Flashermaterial" & nr,0,0,0,0,0,0,ObjLevel(nr),RGB(255,255,255),0,0,False,True,0,0,0,0
  ObjLevel(nr) = ObjLevel(nr) * 0.9 - 0.01
  If ObjLevel(nr) < 0 Then objflasher(nr).TimerEnabled = False : objflasher(nr).visible = 0 : objlit(nr).visible = 0 : End If
End Sub

Sub FlasherFlash1_Timer() : FlashFlasher(1) : End Sub
Sub FlasherFlash2_Timer() : FlashFlasher(2) : End Sub
Sub FlasherFlash3_Timer() : FlashFlasher(3) : End Sub


'******************************************************
'            Gi & GI Effects
'******************************************************

dim GIState: GIState = 1
dim fon,foff
GiEffects.Enabled=1


'Playfield GI
Sub GI_Update(Enabled)
  dim xx
  If Enabled Then
    For each xx in GI:xx.State = 0: Next
    GiMaskOn.Enabled=1:fon=1
    PlaySound "fx_relay",0,0.2,0,0.02
    GIState = 0
  Else
    For each xx in GI:xx.State = 1: Next
    GiMaskOff.Enabled=1:foff=1
    PlaySound "fx_relay",0,0.2,0,0.02
    GIState = 1
  End If
End Sub

Sub GiEffects_timer()

'debug.print gi001.Statev
    FadeDisableLighting Cab, 1, 0.1 'Cab
    FadeDisableLighting Primitive018, 2, 0 'Top Bulbs
    FadeDisableLighting PrampLower, 1.1, 0.2 'Pramp Lower
    FadeDisableLighting PrampUpper, 1, 0.2 'Pramp Upper
    FadeDisableLighting Plastics, 0.3, 0 'Plastics
    FadeDisableLighting PlasticTop, 1.2, 0 'Plastics
    FadeDisableLighting Rocks, 0.3, 0 'Rocks
    FadeDisableLighting Woods, 0.9, 0 'Wood
    FadeDisableLighting Wire001, 0.65, 0 'WRamp 1
    FadeDisableLighting Wire002, 0.65, 0 'WRamp 2
    FadeDisableLighting ApronWire, 0.2, 0 'Apron Wire
    FadeDisableLighting PrimBulb002, 2, 0 'Bulb
    FadeDisableLighting PrimBulb003, 2, 0 'Bulb
    FadeDisableLighting PrimBulb005, 2, 0 'Bulb
    FadeDisableLighting PrimBulb006, 2, 0 'Bulb
    FadeDisableLighting PrimBulb007, 2, 0 'Bulb
    FadeDisableLighting PrimBulb008, 2, 0 'Bulb
    FadeDisableLighting PrimBulb009, 2, 0 'Bulb
    FadeDisableLighting PrimBulb010, 2, 0 'Bulb
    FadeDisableLighting PrimBulb011, 2, 0 'Bulb
    FadeDisableLighting PrimBulb012, 2, 0 'Bulb
    FadeDisableLighting PrimBulb013, 2, 0 'Bulb
    FadeDisableLighting PrimBulb014, 2, 0 'Bulb
    FadeDisableLighting PrimBulb016, 2, 0 'Bulb
    FadeDisableLighting RampSupports, 0.5, 0
    FadeDisableLighting Metals20, 0.8, 0
    FadeDisableLighting GottliebLF, 0.3, 0 'Flipper
    FadeDisableLighting GottliebRF, 0.3, 0 'Flipper
    FadeDisableLighting GottliebRFU, 0.3, 0 'Flipper
    FadeDisableLighting Heli_Base, 0.3, 0
    FadeDisableLighting Chopper2, 0.2, 0 'Helicopter
    FadeDisableLighting Primitive003, 1, 0 'opto1
    FadeDisableLighting Primitive010, 0.3, 0 'opto2
    FadeDisableLighting FlashGS, 1.5, 0.4
    FadeDisableLighting FlashG, 1, 0.4
    FadeDisableLighting Prim_Level1, 0.1, 0
    FadeDisableLighting Fixings, 0.3, 0
    FadeDisableLighting Primitive009, 0.3, 0 'Flipper



End Sub


Sub GiMaskOn_timer()

  Select Case fon
    Case 1:FlasherGI001.IntensityScale = 0.66:fon=2
    Case 2:FlasherGI001.IntensityScale = 0.33:fon=3
    Case 3:FlasherGI001.IntensityScale = 0:fon=4
    Case 4:GiMaskOn.Enabled=0
  End Select

End Sub

Sub GiMaskOff_timer()

  Select Case foff
    Case 1:FlasherGI001.IntensityScale = 0.33:foff=2
    Case 2:FlasherGI001.IntensityScale = 0.66:foff=3
    Case 3:FlasherGI001.IntensityScale = 1:foff=4
    Case 4:GiMaskOff.Enabled=0
  End Select

End Sub


'Fade DisableLighting (object, starting brightness, ending brightness)

Sub FadeDisableLighting(a, alvlu, alvld)

  Dim FTime
  FTime = (alvlu - alvld) / 3

  Select Case gi002.State

    Case 1:
      a.UserValue = a.UserValue + FTime
      If a.UserValue > alvlu Then
        a.UserValue = alvlu
      end If

      a.BlendDisableLighting = a.UserValue 'On
      'PlaySound "fx_relay",0,0.2,0,0.02

    Case 0:
      a.UserValue = a.UserValue - FTime
      If a.UserValue < alvld Then
        a.UserValue = alvld
      end If

      a.BlendDisableLighting = a.UserValue 'Off
      'PlaySound "fx_relay",0,0.2,0,0.02

    End Select
End Sub

'***************************************************
'       JP's VP10 Fading Lamps & Flashers
'       Based on PD's Fading Light System
'       Plus Anti fading disabled lighting
' SetLamp 0 is Off
' SetLamp 1 is On
' fading for non opacity objects is 4 steps
'***************************************************

Dim LampState(200), FadingState(200)
Dim FlashSpeedUp(200), FlashSpeedDown(200), FlashMin(200), FlashMax(200), FlashLevel(200)
Dim PO199:PO199 = 1

InitLamps()             ' turn off the lights and flashers and reset them to the default parameters

' Lamp & Flasher Timers

Sub LampTimer_Timer()
    Dim chgLamp, num, chg, ii
    chgLamp = Controller.ChangedLamps
    If Not IsEmpty(chgLamp)Then
        For ii = 0 To UBound(chgLamp)
            LampState(chgLamp(ii, 0)) = chgLamp(ii, 1)       'keep the real state in an array
            FadingState(chgLamp(ii, 0)) = chgLamp(ii, 1) + 3 'fading step
        Next
    End If
    'UpdateLeds
    UpdateLamps
End Sub

Sub UpdateLamps


  Lamp 0, L0

  Lamp 2, L2
  Lamp 3, L3
  Lamp 4, L4
  Lamp 5, L5
  Lamp 6, L6
  Lamp 7, L7

  Lamp 12, L12
  Lamp 13, L13
  Lamp 14, L14
  Lamp 15, L15
  Lamp 16, L16
  Lamp 17, L17

  Lamp 22, L22
  Lamp 23, L23
  Lamp 24, L24
  Lamp 25, L25
  Lamp 26, L26
  Lamp 27, L27

  Lamp 32, L32
  Lamp 33, L33
  Lamp 34, L34
  Lamp 35, L35
  Lamp 36, L36
  Lamp 37, L37

  Lamp 42, L42
  Lamp 43, L43
  Lamp 44, L44
  Lamp 45, L45
  Lamp 46, L46
  Lamp 47, L47

  Lampm 52, L52a
  FDL 52, L52, 2,0
  Lampm 53, L53a
  FDL 53, L53, 2,0


  Lampm 54, L54
  FadeObjm 54, City, "City-Left-ON", "City-Left-Mid", "City-Left-Fade", "City-Left"
  FDL 54, City, 0.25,0.08

  Lampm 55, L55
  FadeObjm 55, City001, "City-Left-ON", "City-Left-Mid", "City-Left-Fade", "City-Left"
  FDL 55, City001, 0.25,0.08


  Lampm 56, L56
  FadeObjm 56, City002, "City-Left-ON", "City-Left-Mid", "City-Left-Fade", "City-Left"
  FDL 56, City002, 0.25,0.08

  Lampm 57, L57
  FadeObjm 57, City003, "City-Left-ON", "City-Left-Mid", "City-Left-Fade", "City-Left"
  FDL 57, City003, 0.25,0.08

  Lamp 60, L60
  Lamp 61, L61
  Lamp 62, L62
  Lamp 63, L63
  Lamp 64, L64
  Lamp 65, L65
  Lamp 66, L66
  Lamp 67, L67

  Lamp 70, L70
  Lamp 71, L71
  Lamp 72, L72
  Lamp 73, L73

  Lamp 74, L74

  Lampm 75, L75
  Flash 75, F75

  Lampm 77, L77

  Lampm 80, B80a
  Lampm 80, B80
  FDLm 80, BumpL001, 0.6,0
  FDL 80, BumpL, 0.8,0
  Lampm 81, B81a
  Lampm 81, B81
  FDLm 81, BumpR001, 0.6,0
  FDL 81, BumpR, 2,0
  Lampm 82, B82a
  Lampm 82, B82
  FDLm 82, BumpC001, 0.6,0
  FDL 82, BumpC, 2,0

  Lampm 83, L83
  Flash 83, F83

  Flashm 114, F114
  Flashm 114, L114
  Lampm 114, L114a
  FDL 114, PrimBulb001, 1,0

  Flashm 115, L115
  Flashm 115, F115
  Lampm 115, L115a
  FDL 115, PrimBulb004, 1,0

  Lampm 116, L116
  Flash 116, F116

  Lampm 120, l120
  if Flashon = 0 then
    Lampm 120, l124
    FadeObjm 120, ApronBFlash, "dome2litblue", "dome2baseblue_2", "dome2baseblue_2", "dome2baseblue"
    FDLm 120, ApronBFlash, 0.5,0
    Flashm 120, F124
    Flashm 120, FlasherGI002
  end if
  Flash 120, F120

  Lampm 124, l124
  FadeObjm 120, ApronBFlash, "dome2litblue", "dome2baseblue_2", "dome2baseblue_2", "dome2baseblue"
  FDLm 124, ApronBFlash, 0.5,0
  Flashm 124, FlasherGI002
  Flash 124, F124

  Lampm 121, l121
  if Flashon = 0 then
    Lampm 121, l125
    FadeObjm 121, ApronRFlash, "dome2litred", "dome2basered_2", "dome2basered_2", "dome2basered"
    FDLm 121, ApronRFlash, 0.5,0
    Flashm 121, F125
    Flashm 121, FlasherGI003
  end if
  Flash 121, F121

  Lampm 125, l125
  FadeObjm 121, ApronRFlash, "dome2litred", "dome2basered_2", "dome2basered_2", "dome2basered"
  FDLm 125, ApronRFlash, 0.5,0
  Flashm 125, FlasherGI003
  Flash 125, F125







  Flash 125, F125

  Flashm 130, F130
  FDLm 130, PrimBulb015, 4,0
  FDL 130, Rocks, 1,0


End Sub

'*********************************************************************

' Lamp subs

' Normal Lamp & Flasher subs

Sub InitLamps()
    Dim x
    LampTimer.Interval = 30 ' flasher fading speed
    LampTimer.Enabled = 1
    For x = 0 to 200
        LampState(x) = 0
        FadingState(x) = 3 ' used to track the fading state
        FlashLevel(x) = 0
    Next
End Sub

Sub SetLamp(nr, value) ' 0 is off, 1 is on
    FadingState(nr) = abs(value) + 3
End Sub

' Lights: used for VPX standard lights, the fading is handled by VPX itself, they are here to be able to make them work together with the flashers

Sub Lamp(nr, object)
    Select Case FadingState(nr)
        Case 4:object.state = 1:FadingState(nr) = 0
        Case 3:object.state = 0:FadingState(nr) = 0
    End Select
End Sub

Sub Lampm(nr, object) ' used for multiple lights, it doesn't change the fading state
    Select Case FadingState(nr)
        Case 4:object.state = 1
        Case 3:object.state = 0
    End Select
End Sub


' Flashers: 4 is on,3,2,1 fade steps. 0 is off

Sub Flash(nr, object)
    Select Case FadingState(nr)
        Case 4:Object.IntensityScale = 1:FadingState(nr) = 0
        Case 3:Object.IntensityScale = 0.66:FadingState(nr) = 2
        Case 2:Object.IntensityScale = 0.33:FadingState(nr) = 1
        Case 1:Object.IntensityScale = 0:FadingState(nr) = 0
    End Select
End Sub

Sub Flashm(nr, object) 'multiple flashers, it doesn't change the fading state
    Select Case FadingState(nr)
        Case 4:Object.IntensityScale = 1
        Case 3:Object.IntensityScale = 0.66
        Case 2:Object.IntensityScale = 0.33
        Case 1:Object.IntensityScale = 0
    End Select
End Sub


'Fade DisableLighting (object, start state, end state)


Sub FDL(nr, object, ss, es)

  Select Case FadingState(nr)

    Case 4:Object.BlendDisableLighting = ss:FadingState(nr) = 0
        Case 3:Object.BlendDisableLighting = ss / 1.4:FadingState(nr) = 2
        Case 2:Object.BlendDisableLighting = ss / 2.5:FadingState(nr) = 1
        Case 1:Object.BlendDisableLighting = es:FadingState(nr) = 0

  End Select

End Sub


Sub FDLm(nr, object, ss, es)

  Select Case FadingState(nr)

    Case 4:Object.BlendDisableLighting = ss
        Case 3:Object.BlendDisableLighting = ss / 1.4
        Case 2:Object.BlendDisableLighting = ss / 2.5
        Case 1:Object.BlendDisableLighting = es

  End Select

End Sub


'Fade DisableLighting two state (object, start state, end state)

Sub NFDLm(nr, object, starts, ends)

  Select Case FadingState(nr)

    Case 4:Object.BlendDisableLighting = starts
        Case 3:Object.BlendDisableLighting = ends


  End Select

End Sub


' Desktop Objects: Reels & texts (you may also use lights on the desktop)

' Reels

Sub Reel(nr, object)
    Select Case FadingState(nr)
        Case 4:object.SetValue 1:FadingState(nr) = 0
        Case 3:object.SetValue 2:FadingState(nr) = 2
        Case 2:object.SetValue 3:FadingState(nr) = 1
        Case 1:object.SetValue 0:FadingState(nr) = 0
    End Select
End Sub

Sub Reelm(nr, object)
    Select Case FadingState(nr)
        Case 4:object.SetValue 1
        Case 3:object.SetValue 2
        Case 2:object.SetValue 3
        Case 1:object.SetValue 0
    End Select
End Sub

Sub NFadeReel(nr, object)
    Select Case FadingState(nr)
        Case 4:object.SetValue 1:FadingState(nr) = 1
        Case 3:object.SetValue 0:FadingState(nr) = 0
    End Select
End Sub

Sub NFadeReelm(nr, object)
    Select Case FadingState(nr)
        Case 4:object.SetValue 1
        Case 3:object.SetValue 0
    End Select
End Sub

'Texts

Sub Text(nr, object, message)
    Select Case FadingState(nr)
        Case 4:object.Text = message:FadingState(nr) = 0
        Case 3:object.Text = "":FadingState(nr) = 0
    End Select
End Sub

Sub Textm(nr, object, message)
    Select Case FadingState(nr)
        Case 4:object.Text = message
        Case 3:object.Text = ""
    End Select
End Sub

' Modulated Subs for the WPC tables

Sub SetModLamp(nr, level)
    FlashLevel(nr) = level / 150 'lights & flashers
End Sub

Sub LampMod(nr, object)          ' modulated lights used as flashers
    Object.IntensityScale = FlashLevel(nr)
    Object.State = 1             'in case it was off
End Sub

Sub FlashMod(nr, object)         'sets the flashlevel from the SolModCallback
    Object.IntensityScale = FlashLevel(nr)
End Sub

'Walls and mostly Primitives used as 4 step fading lights
'a,b,c,d are the images used from on to off

Sub FadeObj(nr, object, a, b, c, d)
    Select Case FadingState(nr)
        Case 4:object.image = a:FadingState(nr) = 0 'fading to off...
        Case 3:object.image = b:FadingState(nr) = 2
        Case 2:object.image = c:FadingState(nr) = 1
        Case 1:object.image = d:FadingState(nr) = 0
    End Select
End Sub

Sub FadeObjm(nr, object, a, b, c, d)
    Select Case FadingState(nr)
        Case 4:object.image = a
        Case 3:object.image = b
        Case 2:object.image = c
        Case 1:object.image = d
    End Select
End Sub

Sub NFadeObj(nr, object, a, b)
    Select Case FadingState(nr)
        Case 4:object.image = a:FadingState(nr) = 0 'off
        Case 3:object.image = b:FadingState(nr) = 0 'on
    End Select
End Sub

Sub NFadeObjm(nr, object, a, b)
    Select Case FadingState(nr)
        Case 4:object.image = a
        Case 3:object.image = b
    End Select
End Sub

'**********Sling Shot Animations
' Rstep and Lstep  are the variables that increment the animation
'****************
Dim RStep, Lstep

Sub RightSlingShot_Slingshot
  vpmTimer.PulseSw 14
    RandomSoundSlingshotRight Sling1
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
  vpmTimer.PulseSw 13
    RandomSoundSlingshotLeft Sling2
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

'*****************************************
' ninuzzu's FLIPPER SHADOWS
'*****************************************

sub FlipperTimer_Timer()
  FlipperLSh.RotZ = LeftFlipper.currentangle
  FlipperRSh.RotZ = RightFlipper.currentangle
  FlipperRSh002.RotZ = RightFlipper002.currentangle
'Rotate Flipper prims
    GottliebLF.RotY = LeftFlipper.CurrentAngle
    GottliebRF.RotY = RightFlipper.CurrentAngle
  GottliebRFU.RotY = RightFlipper002.CurrentAngle
  GottliebLFX.RotZ = LeftFlipper.CurrentAngle
  GottliebRFX.RotZ = RightFlipper.CurrentAngle

End Sub


'***********************************************
'   JP's VP10 Rolling Sounds + Ballshadow v3.0
'   uses a collection of shadows, aBallShadow
'***********************************************

Const tnob = 7  'total number of balls, 20 balls, from 0 to 19
Const lob = 3    'number of locked balls
Const maxvel = 48 'max ball velocity 25-50

Dim DropCount
ReDim DropCount(tnob)

ReDim rolling(tnob)
InitRolling

Sub InitRolling
    Dim i
    For i = 0 to tnob
        rolling(i) = False
    Next
End Sub

Sub RollingUpdate()
    Dim BOT, b, ballpitch, ballvol, speedfactorx, speedfactory
    BOT = GetBalls

    ' stop the sound of deleted balls and hide the shadow
    For b = UBound(BOT) + 1 to tnob
        rolling(b) = False
        StopSound("BallRoll_" & b)
        aBallShadow(b).Y = 3000
    Next


    ' exit the sub if no balls on the table
    If UBound(BOT) = lob - 1 Then Exit Sub 'there no extra balls on this table

    ' play the rolling sound for each ball and draw the shadow
    For b = lob to UBound(BOT)

    'Angled Shadow
    If BOT(b).z >=0 Then
        If BOT(b).X < tablewidth/2 Then
          aBallShadow(b).X = ((BOT(b).X) - (Ballsize/6) + ((BOT(b).X - (tablewidth/2))/21)) + 6
        Else
          aBallShadow(b).X = ((BOT(b).X) + (Ballsize/6) + ((BOT(b).X - (tablewidth/2))/21)) - 6
        End If
        aBallShadow(b).Y = BOT(b).Y + 12
    End If

    'Simple Shadow
'        aBallShadow(b).X = BOT(b).X
'        aBallShadow(b).Y = BOT(b).Y


        If BallVel(BOT(b)) > 1 Then
            If BOT(b).z < 30 Then
                ballpitch = PitchPlayfieldRoll(BOT(b))
                ballvol = VolPlayfieldRoll(BOT(b)) * VolumeDial
            Else
                ballpitch = PitchPlayfieldRoll(BOT(b)) + 25000 'increase the pitch on a ramp
                ballvol = VolPlayfieldRoll(BOT(b)) * 10 * VolumeDial
            End If
            rolling(b) = True
            PlaySound ("BallRoll_" & b), -1, ballvol, AudioPan(BOT(b)), 0, ballpitch, 1, 0, AudioFade(BOT(b))
        Else
            If rolling(b) = True Then
                StopSound("BallRoll_" & b)
                rolling(b) = False
            End If
        End If

        '***Ball Drop Sounds***
        If BOT(b).VelZ < -1 and BOT(b).z < 55 and BOT(b).z > 27 Then 'height adjust for ball drop sounds
            If DropCount(b) >= 5 Then
               DropCount(b) = 0
               If BOT(b).velz > -7 Then
                  RandomSoundBallBouncePlayfieldSoft BOT(b)
               Else
                  RandomSoundBallBouncePlayfieldHard BOT(b)
               End If
      End If
        End If
        If DropCount(b) < 5 Then
            DropCount(b) = DropCount(b) + 1
        End If

        ' jps ball speed control
        If BOT(b).VelX AND BOT(b).VelY <> 0 Then
            speedfactorx = ABS(maxvel / BOT(b).VelX)
            speedfactory = ABS(maxvel / BOT(b).VelY)
            If speedfactorx < 1 Then
                BOT(b).VelX = BOT(b).VelX * speedfactorx
                BOT(b).VelY = BOT(b).VelY * speedfactorx
            End If
            If speedfactory < 1 Then
                BOT(b).VelX = BOT(b).VelX * speedfactory
                BOT(b).VelY = BOT(b).VelY * speedfactory
            End If
        End If
    Next
End Sub



'/////////////////////////////  SOUND PLAYBACK FUNCTIONS  ////////////////////////////
'/////////////////////////////  POSITIONAL SOUND PLAYBACK METHODS  ////////////////////////////
' Positional sound playback methods will play a sound, depending on the X,Y position of the table element or depending on ActiveBall object position
' These are similar subroutines that are less complicated to use (e.g. simply use standard parameters for the PlaySound call)
' For surround setup - positional sound playback functions will fade between front and rear surround channels and pan between left and right channels
' For stereo setup - positional sound playback functions will only pan between left and right channels
' For mono setup - positional sound playback functions will not pan between left and right channels and will not fade between front and rear channels

' PlaySound full syntax - PlaySound(string, int loopcount, float volume, float pan, float randompitch, int pitch, bool useexisting, bool restart, float front_rear_fade)
' Note - These functions will not work (currently) for walls/slingshots as these do not feature a simple, single X,Y position
Sub PlaySoundAtLevelStatic(playsoundparams, aVol, tableobj)
    PlaySound playsoundparams, 0, aVol * VolumeDial, AudioPan(tableobj), 0, 0, 0, 0, AudioFade(tableobj)
End Sub

Sub PlaySoundAtLevelExistingStatic(playsoundparams, aVol, tableobj)
    PlaySound playsoundparams, 0, aVol * VolumeDial, AudioPan(tableobj), 0, 0, 1, 0, AudioFade(tableobj)
End Sub

Sub PlaySoundAtLevelStaticLoop(playsoundparams, aVol, tableobj)
    PlaySound playsoundparams, -1, aVol * VolumeDial, AudioPan(tableobj), 0, 0, 0, 0, AudioFade(tableobj)
End Sub

Sub PlaySoundAtLevelStaticRandomPitch(playsoundparams, aVol, randomPitch, tableobj)
    PlaySound playsoundparams, 0, aVol * VolumeDial, AudioPan(tableobj), randomPitch, 0, 0, 0, AudioFade(tableobj)
End Sub

Sub PlaySoundAtLevelActiveBall(playsoundparams, aVol)
        PlaySound playsoundparams, 0, aVol * VolumeDial, AudioPan(ActiveBall), 0, 0, 0, 0, AudioFade(ActiveBall)
End Sub

Sub PlaySoundAtLevelExistingActiveBall(playsoundparams, aVol)
        PlaySound playsoundparams, 0, aVol * VolumeDial, AudioPan(ActiveBall), 0, 0, 1, 0, AudioFade(ActiveBall)
End Sub

Sub PlaySoundAtLeveTimerActiveBall(playsoundparams, aVol, ballvariable)
        PlaySound playsoundparams, 0, aVol * VolumeDial, AudioPan(ballvariable), 0, 0, 0, 0, AudioFade(ballvariable)
End Sub

Sub PlaySoundAtLevelTimerExistingActiveBall(playsoundparams, aVol, ballvariable)
        PlaySound playsoundparams, 0, aVol * VolumeDial, AudioPan(ballvariable), 0, 0, 1, 0, AudioFade(ballvariable)
End Sub

Sub PlaySoundAtLevelRoll(playsoundparams, aVol, pitch)
    PlaySound playsoundparams, -1, aVol * VolumeDial, AudioPan(tableobj), randomPitch, 0, 0, 0, AudioFade(tableobj)
End Sub

' Previous Positional Sound Subs

Sub PlayXYSound(soundname, tableobj, loopcount, volume, randompitch, pitch, useexisting, restart)
  PlaySound soundname, loopcount, volume, AudioPan(tableobj), randompitch, pitch, useexisting, restart, AudioFade(tableobj)
End Sub

Sub PlaySoundAt(soundname, tableobj)
    PlaySound soundname, 1, 1 * VolumeDial, AudioPan(tableobj), 0,0,0, 1, AudioFade(tableobj)
End Sub

Sub PlaySoundAtVol(soundname, tableobj, aVol)
    PlaySound soundname, 1, aVol * VolumeDial, AudioPan(tableobj), 0,0,0, 1, AudioFade(tableobj)
End Sub

Sub PlaySoundAtBall(soundname)
    PlaySoundAt soundname, ActiveBall
End Sub

Sub PlaySoundAtBallVol (Soundname, aVol)
        Playsound soundname, 1,aVol * VolumeDial, AudioPan(ActiveBall), 0,0,0, 1, AudioFade(ActiveBall)
End Sub

Sub PlaySoundAtBallVolM (Soundname, aVol)
        Playsound soundname, 1,aVol * VolumeDial, AudioPan(ActiveBall), 0,0,0, 0, AudioFade(ActiveBall)
End Sub

Sub PlaySoundAtVolLoops(sound, tableobj, Vol, Loops)
        PlaySound sound, Loops, Vol * VolumeDial, AudioPan(tableobj), 0,0,0, 1, AudioFade(tableobj)
End Sub


'////////////////////////////  MECHANICAL SOUNDS  ///////////////////////////
'//  This part in the script is an entire block that is dedicated to the physics sound system.
'//  Various scripts and sounds that may be pretty generic and could suit other WPC systems, but the most are tailored specifically for this table.

'///////////////////////////////  SOUNDS PARAMETERS  //////////////////////////////
Dim GlobalSoundLevel, CoinSoundLevel, PlungerReleaseSoundLevel, PlungerPullSoundLevel, NudgeLeftSoundLevel
Dim NudgeRightSoundLevel, NudgeCenterSoundLevel, StartButtonSoundLevel, RollingSoundFactor

CoinSoundLevel = 1                                                                                                                'volume level; range [0, 1]
NudgeLeftSoundLevel = 1                                                                                                        'volume level; range [0, 1]
NudgeRightSoundLevel = 1                                                                                                'volume level; range [0, 1]
NudgeCenterSoundLevel = 1                                                                                                'volume level; range [0, 1]
StartButtonSoundLevel = 0.1                                                                                                'volume level; range [0, 1]
PlungerReleaseSoundLevel = 0.8 '1 wjr                                                                                        'volume level; range [0, 1]
PlungerPullSoundLevel = 1                                                                                                'volume level; range [0, 1]
RollingSoundFactor = 4

'///////////////////////-----Solenoids, Kickers and Flash Relays-----///////////////////////
Dim FlipperUpAttackMinimumSoundLevel, FlipperUpAttackMaximumSoundLevel, FlipperUpAttackLeftSoundLevel, FlipperUpAttackRightSoundLevel
Dim FlipperUpSoundLevel, FlipperDownSoundLevel, FlipperLeftHitParm, FlipperRightHitParm
Dim SlingshotSoundLevel, BumperSoundFactor, KnockerSoundLevel

FlipperUpAttackMinimumSoundLevel = 0.010                                                           'volume level; range [0, 1]
FlipperUpAttackMaximumSoundLevel = 0.635                                                                'volume level; range [0, 1]
FlipperUpSoundLevel = 1.0                                                                        'volume level; range [0, 1]
FlipperDownSoundLevel = 0.45                                                                      'volume level; range [0, 1]
FlipperLeftHitParm = FlipperUpSoundLevel                                                                'sound helper; not configurable
FlipperRightHitParm = FlipperUpSoundLevel                                                                'sound helper; not configurable
SlingshotSoundLevel = 0.95                                                                                                'volume level; range [0, 1]
BumperSoundFactor = 4.25                                                                                                'volume multiplier; must not be zero
KnockerSoundLevel = 1                                                                                                         'volume level; range [0, 1]

'///////////////////////-----Ball Drops, Bumps and Collisions-----///////////////////////
Dim RubberStrongSoundFactor, RubberWeakSoundFactor, RubberFlipperSoundFactor,BallWithBallCollisionSoundFactor
Dim BallBouncePlayfieldSoftFactor, BallBouncePlayfieldHardFactor, PlasticRampDropToPlayfieldSoundLevel, WireRampDropToPlayfieldSoundLevel, DelayedBallDropOnPlayfieldSoundLevel
Dim WallImpactSoundFactor, MetalImpactSoundFactor, SubwaySoundLevel, SubwayEntrySoundLevel, ScoopEntrySoundLevel
Dim SaucerLockSoundLevel, SaucerKickSoundLevel

BallWithBallCollisionSoundFactor = 3.2                                                                        'volume multiplier; must not be zero
RubberStrongSoundFactor = 0.055/5                                                                                        'volume multiplier; must not be zero
RubberWeakSoundFactor = 0.075/5                                                                                        'volume multiplier; must not be zero
RubberFlipperSoundFactor = 0.075/5                                                                                'volume multiplier; must not be zero
BallBouncePlayfieldSoftFactor = 0.025                                                                        'volume multiplier; must not be zero
BallBouncePlayfieldHardFactor = 0.025                                                                        'volume multiplier; must not be zero
DelayedBallDropOnPlayfieldSoundLevel = 0.8                                                                        'volume level; range [0, 1]
WallImpactSoundFactor = 0.075/7                                                                                       'volume multiplier; must not be zero
MetalImpactSoundFactor = 0.075/7
SaucerLockSoundLevel = 0.8
SaucerKickSoundLevel = 0.8

'///////////////////////-----Gates, Spinners, Rollovers, Ramps and Targets-----///////////////////////

Dim GateSoundLevel, TargetSoundFactor, SpinnerSoundLevel, RolloverSoundLevel, DTSoundLevel, WireRampSoundFactor, LoopSoundFactor

GateSoundLevel = 0.01                                                                                                       'volume level; range [0, 1]
TargetSoundFactor = 0.0025 * 10                                                                                        'volume multiplier; must not be zero
DTSoundLevel = 0.25                                                                                                                'volume multiplier; must not be zero
RolloverSoundLevel = 0.25
WireRampSoundFactor = 0.3
LoopSoundFactor = 1 * 20


'///////////////////////-----Ball Release, Guides and Drain-----///////////////////////
Dim DrainSoundLevel, BallReleaseSoundLevel, BottomArchBallGuideSoundFactor, FlipperBallGuideSoundFactor

DrainSoundLevel = 0.8                                                                                                                'volume level; range [0, 1]
BallReleaseSoundLevel = 1                                                                                                'volume level; range [0, 1]
BottomArchBallGuideSoundFactor = 0.2                                                                        'volume multiplier; must not be zero
FlipperBallGuideSoundFactor = 0.015                                                                                'volume multiplier; must not be zero

'///////////////////////-----Loops and Lanes-----///////////////////////
Dim ArchSoundFactor
ArchSoundFactor = 0.025/5                                                                                                        'volume multiplier; must not be zero


' *********************************************************************
'                     Fleep  Supporting Ball & Sound Functions
' *********************************************************************



Function AudioFade(tableobj) ' Fades between front and back of the table (for surround systems or 2x2 speakers, etc), depending on the Y position on the table. "table1" is the name of the table
        Dim tmp
    tmp = tableobj.y * 2 / tableheight-1
    If tmp > 0 Then
                AudioFade = Csng(tmp ^10)
    Else
        AudioFade = Csng(-((- tmp) ^10) )
    End If
End Function

Function AudioPan(tableobj) ' Calculates the pan for a tableobj based on the X position on the table. "table1" is the name of the table
    Dim tmp
    tmp = tableobj.x * 2 / tablewidth-1
    If tmp > 0 Then
        AudioPan = Csng(tmp ^10)
    Else
        AudioPan = Csng(-((- tmp) ^10) )
    End If
End Function

Function Vol(ball) ' Calculates the volume of the sound based on the ball speed
        Vol = Csng(BallVel(ball) ^2)
End Function

Function Volz(ball) ' Calculates the volume of the sound based on the ball speed
        Volz = Csng((ball.velz) ^2)
End Function

Function Pitch(ball) ' Calculates the pitch of the sound based on the ball speed
    Pitch = BallVel(ball) * 20
End Function

Function BallVel(ball) 'Calculates the ball speed
    BallVel = INT(SQR((ball.VelX ^2) + (ball.VelY ^2) ) )
End Function

Function VolPlayfieldRoll(ball) ' Calculates the roll volume of the sound based on the ball speed
        VolPlayfieldRoll = RollingSoundFactor * 0.0005 * Csng(BallVel(ball) ^3)
End Function

Function PitchPlayfieldRoll(ball) ' Calculates the roll pitch of the sound based on the ball speed
    PitchPlayfieldRoll = BallVel(ball) ^2 * 15
End Function

Function RndInt(min, max)
    RndInt = Int(Rnd() * (max-min + 1) + min)' Sets a random number integer between min and max
End Function

Function RndNum(min, max)
    RndNum = Rnd() * (max-min) + min' Sets a random number between min and max
End Function

'/////////////////////////////  GENERAL SOUND SUBROUTINES  ////////////////////////////
Sub SoundStartButton()
        PlaySound ("Start_Button"), 0, StartButtonSoundLevel, 0, 0.25
End Sub

Sub SoundNudgeLeft()
        PlaySound ("Nudge_" & Int(Rnd*2)+1), 0, NudgeLeftSoundLevel * VolumeDial, -0.1, 0.25
End Sub

Sub SoundNudgeRight()
        PlaySound ("Nudge_" & Int(Rnd*2)+1), 0, NudgeRightSoundLevel * VolumeDial, 0.1, 0.25
End Sub

Sub SoundNudgeCenter()
        PlaySound ("Nudge_" & Int(Rnd*2)+1), 0, NudgeCenterSoundLevel * VolumeDial, 0, 0.25
End Sub


Sub SoundPlungerPull()
        PlaySoundAtLevelStatic ("Plunger_Pull_1"), PlungerPullSoundLevel, Plunger
End Sub

Sub SoundPlungerReleaseBall()
        PlaySoundAtLevelStatic ("Plunger_Release_Ball"), PlungerReleaseSoundLevel, Plunger
End Sub

Sub SoundPlungerReleaseNoBall()
        PlaySoundAtLevelStatic ("Plunger_Release_No_Ball"), PlungerReleaseSoundLevel, Plunger
End Sub


'/////////////////////////////  KNOCKER SOLENOID  ////////////////////////////
Sub KnockerSolenoid()
        PlaySoundAtLevelStatic SoundFX("Knocker_1",DOFKnocker), KnockerSoundLevel, KnockerPosition
End Sub

'/////////////////////////////  DRAIN SOUNDS  ////////////////////////////
Sub RandomSoundDrain(drainswitch)
        PlaySoundAtLevelStatic ("Drain_" & Int(Rnd*11)+1), DrainSoundLevel, drainswitch
End Sub

'/////////////////////////////  TROUGH BALL RELEASE SOLENOID SOUNDS  ////////////////////////////

Sub RandomSoundBallRelease(drainswitch)
        PlaySoundAtLevelStatic SoundFX("BallRelease" & Int(Rnd*7)+1,DOFContactors), BallReleaseSoundLevel, drainswitch
End Sub

'/////////////////////////////  SLINGSHOT SOLENOID SOUNDS  ////////////////////////////
Sub RandomSoundSlingshotLeft(sling)
        PlaySoundAtLevelStatic SoundFX("Sling_L" & Int(Rnd*10)+1,DOFContactors), SlingshotSoundLevel, Sling
End Sub

Sub RandomSoundSlingshotRight(sling)
        PlaySoundAtLevelStatic SoundFX("Sling_R" & Int(Rnd*8)+1,DOFContactors), SlingshotSoundLevel, Sling
End Sub

'/////////////////////////////  BUMPER SOLENOID SOUNDS  ////////////////////////////
Sub RandomSoundBumperTop(Bump)
        PlaySoundAtLevelStatic SoundFX("Bumpers_Top_" & Int(Rnd*5)+1,DOFContactors), Vol(ActiveBall) * BumperSoundFactor, Bump
End Sub

Sub RandomSoundBumperMiddle(Bump)
        PlaySoundAtLevelStatic SoundFX("Bumpers_Middle_" & Int(Rnd*5)+1,DOFContactors), Vol(ActiveBall) * BumperSoundFactor, Bump
End Sub

Sub RandomSoundBumperBottom(Bump)
        PlaySoundAtLevelStatic SoundFX("Bumpers_Bottom_" & Int(Rnd*5)+1,DOFContactors), Vol(ActiveBall) * BumperSoundFactor, Bump
End Sub

'/////////////////////////////  FLIPPER BATS SOUND SUBROUTINES  ////////////////////////////
'/////////////////////////////  FLIPPER BATS SOLENOID ATTACK SOUND  ////////////////////////////
Sub SoundFlipperUpAttackLeft(flipper)
        FlipperUpAttackLeftSoundLevel = RndNum(FlipperUpAttackMinimumSoundLevel, FlipperUpAttackMaximumSoundLevel)
        PlaySoundAtLevelStatic ("Flipper_Attack-L01"), FlipperUpAttackLeftSoundLevel, flipper
End Sub

Sub SoundFlipperUpAttackRight(flipper)
        FlipperUpAttackRightSoundLevel = RndNum(FlipperUpAttackMinimumSoundLevel, FlipperUpAttackMaximumSoundLevel)
                PlaySoundAtLevelStatic ("Flipper_Attack-R01"), FlipperUpAttackLeftSoundLevel, flipper
End Sub

'/////////////////////////////  FLIPPER BATS SOLENOID CORE SOUND  ////////////////////////////
Sub RandomSoundFlipperUpLeft(flipper)
        PlaySoundAtLevelStatic SoundFX("Flipper_L0" & Int(Rnd*9)+1,DOFFlippers), FlipperLeftHitParm, Flipper
End Sub

Sub RandomSoundFlipperUpRight(flipper)
        PlaySoundAtLevelStatic SoundFX("Flipper_R0" & Int(Rnd*9)+1,DOFFlippers), FlipperRightHitParm, Flipper
End Sub

Sub RandomSoundReflipUpLeft(flipper)
        PlaySoundAtLevelStatic SoundFX("Flipper_ReFlip_L0" & Int(Rnd*3)+1,DOFFlippers), (RndNum(0.8, 1))*FlipperUpSoundLevel, Flipper
End Sub

Sub RandomSoundReflipUpRight(flipper)
        PlaySoundAtLevelStatic SoundFX("Flipper_ReFlip_R0" & Int(Rnd*3)+1,DOFFlippers), (RndNum(0.8, 1))*FlipperUpSoundLevel, Flipper
End Sub

Sub RandomSoundFlipperDownLeft(flipper)
        PlaySoundAtLevelStatic SoundFX("Flipper_Left_Down_" & Int(Rnd*7)+1,DOFFlippers), FlipperDownSoundLevel, Flipper
End Sub

Sub RandomSoundFlipperDownRight(flipper)
        PlaySoundAtLevelStatic SoundFX("Flipper_Right_Down_" & Int(Rnd*8)+1,DOFFlippers), FlipperDownSoundLevel, Flipper
End Sub

'/////////////////////////////  FLIPPER BATS BALL COLLIDE SOUND  ////////////////////////////

Sub LeftFlipperCollide(parm)
        FlipperLeftHitParm = parm/10
        If FlipperLeftHitParm > 1 Then
                FlipperLeftHitParm = 1
        End If
        FlipperLeftHitParm = FlipperUpSoundLevel * FlipperLeftHitParm
        RandomSoundRubberFlipper(parm)
End Sub

Sub RightFlipperCollide(parm)
        FlipperRightHitParm = parm/10
        If FlipperRightHitParm > 1 Then
                FlipperRightHitParm = 1
        End If
        FlipperRightHitParm = FlipperUpSoundLevel * FlipperRightHitParm
         RandomSoundRubberFlipper(parm)
End Sub

Sub RandomSoundRubberFlipper(parm)
        PlaySoundAtLevelActiveBall ("Flipper_Rubber_" & Int(Rnd*7)+1), parm  * RubberFlipperSoundFactor
End Sub

'/////////////////////////////  ROLLOVER SOUNDS  ////////////////////////////
Sub RandomSoundRollover()
        PlaySoundAtLevelActiveBall ("Rollover_" & Int(Rnd*4)+1), RolloverSoundLevel
End Sub

Sub Rollovers_Hit(idx)
        RandomSoundRollover
End Sub

'/////////////////////////////  VARIOUS PLAYFIELD SOUND SUBROUTINES  ////////////////////////////
'/////////////////////////////  RUBBERS AND POSTS  ////////////////////////////
'/////////////////////////////  RUBBERS - EVENTS  ////////////////////////////
Sub Rubbers_Hit(idx)
         dim finalspeed
          finalspeed=SQR(activeball.velx * activeball.velx + activeball.vely * activeball.vely)
         If finalspeed > 5 then
                 RandomSoundRubberStrong 1
        End if
        If finalspeed <= 5 then
                 RandomSoundRubberWeak()
         End If
End Sub

'/////////////////////////////  RUBBERS AND POSTS - STRONG IMPACTS  ////////////////////////////
Sub RandomSoundRubberStrong(voladj)
        Select Case Int(Rnd*10)+1
                Case 1 : PlaySoundAtLevelActiveBall ("Rubber_Strong_1"), Vol(ActiveBall) * RubberStrongSoundFactor*voladj
                Case 2 : PlaySoundAtLevelActiveBall ("Rubber_Strong_2"), Vol(ActiveBall) * RubberStrongSoundFactor*voladj
                Case 3 : PlaySoundAtLevelActiveBall ("Rubber_Strong_3"), Vol(ActiveBall) * RubberStrongSoundFactor*voladj
                Case 4 : PlaySoundAtLevelActiveBall ("Rubber_Strong_4"), Vol(ActiveBall) * RubberStrongSoundFactor*voladj
                Case 5 : PlaySoundAtLevelActiveBall ("Rubber_Strong_5"), Vol(ActiveBall) * RubberStrongSoundFactor*voladj
                Case 6 : PlaySoundAtLevelActiveBall ("Rubber_Strong_6"), Vol(ActiveBall) * RubberStrongSoundFactor*voladj
                Case 7 : PlaySoundAtLevelActiveBall ("Rubber_Strong_7"), Vol(ActiveBall) * RubberStrongSoundFactor*voladj
                Case 8 : PlaySoundAtLevelActiveBall ("Rubber_Strong_8"), Vol(ActiveBall) * RubberStrongSoundFactor*voladj
                Case 9 : PlaySoundAtLevelActiveBall ("Rubber_Strong_9"), Vol(ActiveBall) * RubberStrongSoundFactor*voladj
                Case 10 : PlaySoundAtLevelActiveBall ("Rubber_1_Hard"), Vol(ActiveBall) * RubberStrongSoundFactor * 0.6*voladj
        End Select
End Sub

'/////////////////////////////  RUBBERS AND POSTS - WEAK IMPACTS  ////////////////////////////
Sub RandomSoundRubberWeak()
        PlaySoundAtLevelActiveBall ("Rubber_" & Int(Rnd*9)+1), Vol(ActiveBall) * RubberWeakSoundFactor
End Sub

'/////////////////////////////  WALL IMPACTS  ////////////////////////////
Sub Walls_Hit(idx)
         dim finalspeed
          finalspeed=SQR(activeball.velx * activeball.velx + activeball.vely * activeball.vely)
         If finalspeed > 5 then
                 RandomSoundRubberStrong 1
        End if
        If finalspeed <= 5 then
                 RandomSoundRubberWeak()
         End If
End Sub

Sub RandomSoundWall()
         dim finalspeed
          finalspeed=SQR(activeball.velx * activeball.velx + activeball.vely * activeball.vely)
         If finalspeed > 16 then
                Select Case Int(Rnd*5)+1
                        Case 1 : PlaySoundAtLevelExistingActiveBall ("Wall_Hit_1"), Vol(ActiveBall) * WallImpactSoundFactor
                        Case 2 : PlaySoundAtLevelExistingActiveBall ("Wall_Hit_2"), Vol(ActiveBall) * WallImpactSoundFactor
                        Case 3 : PlaySoundAtLevelExistingActiveBall ("Wall_Hit_5"), Vol(ActiveBall) * WallImpactSoundFactor
                        Case 4 : PlaySoundAtLevelExistingActiveBall ("Wall_Hit_7"), Vol(ActiveBall) * WallImpactSoundFactor
                        Case 5 : PlaySoundAtLevelExistingActiveBall ("Wall_Hit_9"), Vol(ActiveBall) * WallImpactSoundFactor
                End Select
        End if
        If finalspeed >= 6 AND finalspeed <= 16 then
                Select Case Int(Rnd*4)+1
                        Case 1 : PlaySoundAtLevelExistingActiveBall ("Wall_Hit_3"), Vol(ActiveBall) * WallImpactSoundFactor
                        Case 2 : PlaySoundAtLevelExistingActiveBall ("Wall_Hit_4"), Vol(ActiveBall) * WallImpactSoundFactor
                        Case 3 : PlaySoundAtLevelExistingActiveBall ("Wall_Hit_6"), Vol(ActiveBall) * WallImpactSoundFactor
                        Case 4 : PlaySoundAtLevelExistingActiveBall ("Wall_Hit_8"), Vol(ActiveBall) * WallImpactSoundFactor
                End Select
         End If
        If finalspeed < 6 Then
                Select Case Int(Rnd*3)+1
                        Case 1 : PlaySoundAtLevelExistingActiveBall ("Wall_Hit_4"), Vol(ActiveBall) * WallImpactSoundFactor
                        Case 2 : PlaySoundAtLevelExistingActiveBall ("Wall_Hit_6"), Vol(ActiveBall) * WallImpactSoundFactor
                        Case 3 : PlaySoundAtLevelExistingActiveBall ("Wall_Hit_8"), Vol(ActiveBall) * WallImpactSoundFactor
                End Select
        End if
End Sub

'/////////////////////////////  METAL TOUCH SOUNDS  ////////////////////////////
Sub RandomSoundMetal()
        PlaySoundAtLevelActiveBall ("Metal_Touch_" & Int(Rnd*13)+1), Vol(ActiveBall) * MetalImpactSoundFactor
End Sub

'/////////////////////////////  METAL - EVENTS  ////////////////////////////

Sub Metals_Hit (idx)
        RandomSoundMetal
End Sub

Sub ShooterDiverter_collide(idx)
        RandomSoundMetal
End Sub

'/////////////////////////////  BOTTOM ARCH BALL GUIDE  ////////////////////////////
'/////////////////////////////  BOTTOM ARCH BALL GUIDE - SOFT BOUNCES  ////////////////////////////
Sub RandomSoundBottomArchBallGuide()
         dim finalspeed
          finalspeed=SQR(activeball.velx * activeball.velx + activeball.vely * activeball.vely)
         If finalspeed > 16 then
                PlaySoundAtLevelActiveBall ("Apron_Bounce_"& Int(Rnd*2)+1), Vol(ActiveBall) * BottomArchBallGuideSoundFactor
        End if
        If finalspeed >= 6 AND finalspeed <= 16 then
                 Select Case Int(Rnd*2)+1
                        Case 1 : PlaySoundAtLevelActiveBall ("Apron_Bounce_1"), Vol(ActiveBall) * BottomArchBallGuideSoundFactor
                        Case 2 : PlaySoundAtLevelActiveBall ("Apron_Bounce_Soft_1"), Vol(ActiveBall) * BottomArchBallGuideSoundFactor
                End Select
         End If
        If finalspeed < 6 Then
                 Select Case Int(Rnd*2)+1
                        Case 1 : PlaySoundAtLevelActiveBall ("Apron_Bounce_Soft_1"), Vol(ActiveBall) * BottomArchBallGuideSoundFactor
                        Case 2 : PlaySoundAtLevelActiveBall ("Apron_Medium_3"), Vol(ActiveBall) * BottomArchBallGuideSoundFactor
                End Select
        End if
End Sub

'/////////////////////////////  BOTTOM ARCH BALL GUIDE - HARD HITS  ////////////////////////////
Sub RandomSoundBottomArchBallGuideHardHit()
        PlaySoundAtLevelActiveBall ("Apron_Hard_Hit_" & Int(Rnd*3)+1), BottomArchBallGuideSoundFactor * 0.25
End Sub

Sub Apron_Hit (idx)
        If Abs(cor.ballvelx(activeball.id) < 4) and cor.ballvely(activeball.id) > 7 then
                RandomSoundBottomArchBallGuideHardHit()
        Else
                RandomSoundBottomArchBallGuide
        End If
End Sub

'/////////////////////////////  FLIPPER BALL GUIDE  ////////////////////////////
Sub RandomSoundFlipperBallGuide()
         dim finalspeed
          finalspeed=SQR(activeball.velx * activeball.velx + activeball.vely * activeball.vely)
         If finalspeed > 16 then
                 Select Case Int(Rnd*2)+1
                        Case 1 : PlaySoundAtLevelActiveBall ("Apron_Hard_1"),  Vol(ActiveBall) * FlipperBallGuideSoundFactor
                        Case 2 : PlaySoundAtLevelActiveBall ("Apron_Hard_2"),  Vol(ActiveBall) * 0.8 * FlipperBallGuideSoundFactor
                End Select
        End if
        If finalspeed >= 6 AND finalspeed <= 16 then
                PlaySoundAtLevelActiveBall ("Apron_Medium_" & Int(Rnd*3)+1),  Vol(ActiveBall) * FlipperBallGuideSoundFactor
         End If
        If finalspeed < 6 Then
                PlaySoundAtLevelActiveBall ("Apron_Soft_" & Int(Rnd*7)+1),  Vol(ActiveBall) * FlipperBallGuideSoundFactor
        End If
End Sub

'/////////////////////////////  TARGET HIT SOUNDS  ////////////////////////////
Sub RandomSoundTargetHitStrong()
        PlaySoundAtLevelActiveBall SoundFX("Target_Hit_" & Int(Rnd*4)+5,DOFTargets), Vol(ActiveBall) * 0.45 * TargetSoundFactor
End Sub

Sub RandomSoundTargetHitWeak()
        PlaySoundAtLevelActiveBall SoundFX("Target_Hit_" & Int(Rnd*4)+1,DOFTargets), Vol(ActiveBall) * TargetSoundFactor
End Sub

Sub PlayTargetSound()
         dim finalspeed
          finalspeed=SQR(activeball.velx * activeball.velx + activeball.vely * activeball.vely)
         If finalspeed > 10 then
                 RandomSoundTargetHitStrong()
                RandomSoundBallBouncePlayfieldSoft Activeball
        Else
                 RandomSoundTargetHitWeak()
         End If
End Sub

Sub Targets_Hit (idx)
        PlayTargetSound
End Sub

'/////////////////////////////  BALL BOUNCE SOUNDS  ////////////////////////////
Sub RandomSoundBallBouncePlayfieldSoft(aBall)
        Select Case Int(Rnd*9)+1
                Case 1 : PlaySoundAtLevelStatic ("Ball_Bounce_Playfield_Soft_1"), volz(aBall) * BallBouncePlayfieldSoftFactor, aBall
                Case 2 : PlaySoundAtLevelStatic ("Ball_Bounce_Playfield_Soft_2"), volz(aBall) * BallBouncePlayfieldSoftFactor * 0.5, aBall
                Case 3 : PlaySoundAtLevelStatic ("Ball_Bounce_Playfield_Soft_3"), volz(aBall) * BallBouncePlayfieldSoftFactor * 0.8, aBall
                Case 4 : PlaySoundAtLevelStatic ("Ball_Bounce_Playfield_Soft_4"), volz(aBall) * BallBouncePlayfieldSoftFactor * 0.5, aBall
                Case 5 : PlaySoundAtLevelStatic ("Ball_Bounce_Playfield_Soft_5"), volz(aBall) * BallBouncePlayfieldSoftFactor, aBall
                Case 6 : PlaySoundAtLevelStatic ("Ball_Bounce_Playfield_Hard_1"), volz(aBall) * BallBouncePlayfieldSoftFactor * 0.2, aBall
                Case 7 : PlaySoundAtLevelStatic ("Ball_Bounce_Playfield_Hard_2"), volz(aBall) * BallBouncePlayfieldSoftFactor * 0.2, aBall
                Case 8 : PlaySoundAtLevelStatic ("Ball_Bounce_Playfield_Hard_5"), volz(aBall) * BallBouncePlayfieldSoftFactor * 0.2, aBall
                Case 9 : PlaySoundAtLevelStatic ("Ball_Bounce_Playfield_Hard_7"), volz(aBall) * BallBouncePlayfieldSoftFactor * 0.3, aBall
        End Select
End Sub

Sub RandomSoundBallBouncePlayfieldHard(aBall)
        PlaySoundAtLevelStatic ("Ball_Bounce_Playfield_Hard_" & Int(Rnd*7)+1), volz(aBall) * BallBouncePlayfieldHardFactor, aBall
End Sub

'/////////////////////////////  DELAYED DROP - TO PLAYFIELD - SOUND  ////////////////////////////
Sub RandomSoundDelayedBallDropOnPlayfield(aBall)
        Select Case Int(Rnd*5)+1
                Case 1 : PlaySoundAtLevelStatic ("Ball_Drop_Playfield_1_Delayed"), DelayedBallDropOnPlayfieldSoundLevel, aBall
                Case 2 : PlaySoundAtLevelStatic ("Ball_Drop_Playfield_2_Delayed"), DelayedBallDropOnPlayfieldSoundLevel, aBall
                Case 3 : PlaySoundAtLevelStatic ("Ball_Drop_Playfield_3_Delayed"), DelayedBallDropOnPlayfieldSoundLevel, aBall
                Case 4 : PlaySoundAtLevelStatic ("Ball_Drop_Playfield_4_Delayed"), DelayedBallDropOnPlayfieldSoundLevel, aBall
                Case 5 : PlaySoundAtLevelStatic ("Ball_Drop_Playfield_5_Delayed"), DelayedBallDropOnPlayfieldSoundLevel, aBall
        End Select
End Sub

'/////////////////////////////  BALL GATES AND BRACKET GATES SOUNDS  ////////////////////////////

Sub SoundPlayfieldGate()
        PlaySoundAtLevelStatic ("Gate_FastTrigger_" & Int(Rnd*2)+1), GateSoundLevel, Activeball
End Sub

Sub SoundHeavyGate()
        PlaySoundAtLevelStatic ("Gate_2"), GateSoundLevel, Activeball
End Sub

Sub Gates_hit(idx)
        SoundHeavyGate
End Sub

Sub GatesWire_hit(idx)
        SoundPlayfieldGate
End Sub

'/////////////////////////////  LEFT LANE ENTRANCE - SOUNDS  ////////////////////////////

Sub RandomSoundLeftArch()
        PlaySoundAtLevelActiveBall ("Arch_L" & Int(Rnd*4)+1), Vol(ActiveBall) * ArchSoundFactor
End Sub

Sub RandomSoundRightArch()
        PlaySoundAtLevelActiveBall ("Arch_R" & Int(Rnd*4)+1), Vol(ActiveBall) * ArchSoundFactor
End Sub


Sub Arch1_hit()
        If Activeball.velx > 1 Then SoundPlayfieldGate
        StopSound "Arch_L1"
        StopSound "Arch_L2"
        StopSound "Arch_L3"
        StopSound "Arch_L4"
End Sub

Sub Arch1_unhit()
        If activeball.velx < -8 Then
                RandomSoundRightArch
        End If
End Sub

Sub Arch2_hit()
        If Activeball.velx < 1 Then SoundPlayfieldGate
        StopSound "Arch_R1"
        StopSound "Arch_R2"
        StopSound "Arch_R3"
        StopSound "Arch_R4"
End Sub

Sub Arch2_unhit()
        If activeball.velx > 10 Then
                RandomSoundLeftArch
        End If
End Sub

'/////////////////////////////  SAUCERS (KICKER HOLES)  ////////////////////////////

Sub SoundSaucerLock()
        PlaySoundAtLevelStatic ("Saucer_Enter_" & Int(Rnd*2)+1), SaucerLockSoundLevel, Activeball
End Sub

Sub SoundSaucerKick(scenario, saucer)
        Select Case scenario
                Case 0: PlaySoundAtLevelStatic SoundFX("Saucer_Empty", DOFContactors), SaucerKickSoundLevel, saucer
                Case 1: PlaySoundAtLevelStatic SoundFX("Saucer_Kick", DOFContactors), SaucerKickSoundLevel, saucer
        End Select
End Sub

'/////////////////////////////  BALL COLLISION SOUND  ////////////////////////////
Sub OnBallBallCollision(ball1, ball2, velocity)
        Dim snd
        Select Case Int(Rnd*7)+1
                Case 1 : snd = "Ball_Collide_1"
                Case 2 : snd = "Ball_Collide_2"
                Case 3 : snd = "Ball_Collide_3"
                Case 4 : snd = "Ball_Collide_4"
                Case 5 : snd = "Ball_Collide_5"
                Case 6 : snd = "Ball_Collide_6"
                Case 7 : snd = "Ball_Collide_7"
        End Select

        PlaySound (snd), 0, Csng(velocity) ^2 / 200 * BallWithBallCollisionSoundFactor * VolumeDial, AudioPan(ball1), 0, Pitch(ball1), 0, 0, AudioFade(ball1)
End Sub


'/////////////////////////////////////////////////////////////////
'                                        End Mechanical Sounds
'/////////////////////////////////////////////////////////////////



'----------------------------------------------------
'Ramp Sounds

Sub RampSound001_hit: PlayXYSound "WireRamp_0", RampSound001, -1, 10 * WireRampSoundFactor,0,0,0,0  :End Sub
Sub RampSoundStop001_Hit: StopSound "WireRamp_0": End Sub

Sub RampSound002_hit: PlayXYSound "WireRamp_0", RampSound002, -1, 10 * WireRampSoundFactor,0,0,0,0 :End Sub
Sub RampSound003_hit: PlayXYSound "WireRamp_0", RampSound003, -1, 10 * WireRampSoundFactor,0,0,0,0 :End Sub
Sub RampSoundStop002_Hit: StopSound "WireRamp_0": End Sub

'----------------------------------------------------


'/////////////////////////////////////////////////////////////////
'              Heli script
'/////////////////////////////////////////////////////////////////

Dim hbin, hbout
Sub Heli_Raise_Hit()
    Controller.Switch(80) = 1
    'PlaySound "kicker_enter_center"
  hbin = 1
End Sub


Sub Heli_Kick_Hit()
    Controller.Switch(80) = 1
    'PlaySound "kicker_enter_center"
End Sub

Dim Helix, Heliy, Helia, Helir, bfx, bfy


Helia = 270 ' 4.68
Helir = 400

Dim hspeed : hspeed = 0.914 / 16.6 '0.714
Dim hrotspeed : hrotspeed = 3.0 / 16.6 '3.0
Dim hinitialrot : hinitialrot = 0
Dim hballr : hballr = 90
Dim cmechpos, helihome, heliaway
cmechpos = 1

Sub heli_animation_timer()

    Select Case cmechpos

'////////// Take off ////////////

        Case 1
      PUBlock.isdropped = 0
            If hinitialrot < 180 Then
         hinitialrot = hinitialrot + hrotspeed
                 Chopper2.ObjRotz = Chopper2.ObjRotz + hrotspeed
            End If
            If Helia > 270+175 Then
                cmechpos = 2
      Else
        Helia = Helia + hspeed
        Helix = Cos(Helia * 6.28318 / 360) * Helir + 190
        Heliy = Sin(Helia * 6.28318 / 360) * Helir + 588.7778
        PrimBlades.X = Helix
        PrimBlades.Y = Heliy
        Heli_Shaft.ObjRotz = Heli_Shaft.ObjRotz + hspeed
        HeliShad.Rotz = Heli_Shaft.ObjRotz + hspeed
        Heli_Base.ObjRotz = Heli_Base.ObjRotz + hspeed
        Chopper2.ObjRotz = Chopper2.ObjRotz + hspeed
        Chopper2.X = Helix
        Chopper2.Y = Heliy
            End If

'////////// Waiting for pickup ////////////

        Case 2
            If mdirection = 0 Then
                cmechpos = 3
        Lower_Post.enabled = 1
        Heli_Kick.enabled = 1
            End If

'////////// Leaving with ball ////////////

        Case 3
            if hinitialrot > 0 then
        Chopper2.ObjRotz = Chopper2.ObjRotz - hrotspeed
        hinitialrot = hinitialrot - hrotspeed
      end if
            Helir = 405
            If Helia < 270 Then
                cmechpos = 4
      Else
        Helia = Helia - hspeed
        Helix = Cos(Helia * 6.28318 / 360) * Helir + 190
        Heliy = Sin(Helia * 6.28318 / 360) * Helir + 588.7778
        PrimBlades.X = Helix
        PrimBlades.Y = Heliy
        Heli_Shaft.ObjRotz = Heli_Shaft.ObjRotz - hspeed
        HeliShad.Rotz = Heli_Shaft.ObjRotz - hspeed
        Heli_Base.ObjRotz = Heli_Base.ObjRotz - hspeed
        Chopper2.ObjRotz = Chopper2.ObjRotz - hspeed
        Chopper2.X = Helix
        Chopper2.Y = Heliy
        If b3active = 1 And hbout = 0 Then
          bfx = Cos((Chopper2.ObjRotz+100) * 6.28318 / 360) * hballr + Helix
          bfy = Sin((Chopper2.ObjRotz+100) * 6.28318 / 360) * hballr + Heliy
          cball.X = bfx
          cball.Y = bfy
        End If
            End If

'////////// Finished.  Reset ////////////

        Case 4
            Chopper2.ObjRotZ = 80
            Chopper2.visible = 1
      PrimBlades.ObjRotZ = 120
            Heli_Shaft.ObjRotz = 180
      HeliShad.Rotz = 180
            Heli_Base.ObjRotz = 0
      hinitialrot = 0
            cmechpos = 1
            Helia = 270 '4.68
            Helir = 400
      If emon = 1 Then
                PlaySound SoundFX("popper_ball", DOFContactors)
        PlaySound "subway"
        Heli_Raise.destroyball
                Heli_Raise.enabled = 1
                Heli_Kick.enabled = 0
                Kicker6.createball
                Kicker6.kick 180, 1
        emon = 0
            End If
            hbout = 1
            b3active = 1
      me.enabled = 0
    End Select
End Sub

'/////////////////////////////////////////////////////


Dim b3active, BOut
b3active = 1

Dim br, ba, by, bx

br = 100
ba = 3.38



Sub Lower_Post_Timer()
    If Ball_Raise.Z <= -55 Then
    drop_pause.enabled = 1
        Controller.Switch(34) = 1
    PlaySound "solenoid"
        Me.enabled = 0
    End If
    If Ball_Raise.Z <= -40 Then
        Controller.Switch(24) = 0
    End If
    Ball_Raise.Z = Ball_Raise.Z - 5
End Sub

Sub drop_pause_timer()
  Ball_Raise.Z = -55
    PUBlock.isdropped = 1
  Heli_Kick.Enabled = 1
    Me.enabled = 0
End Sub

Dim cball

Sub Raise_Post_Timer()
    If cball.z >= 135 Then
    PlaySound "left_slingshot"
        Controller.Switch(24) = 1
        hbin = 0
        hbout = 0
    Me.enabled = 0
    End If
    If cball.z >= 10 Then
        Controller.Switch(80) = 0
        Controller.Switch(34) = 0
    End If
    cball.z = cball.z + 0.8 '0.31
    Ball_Raise.Z = Ball_Raise.Z + 1  '0.31
End Sub

Sub Blades_Spin_Timer()
    PrimBlades.ObjRotz = SystemTime Mod 360
End Sub


'********************************************

Sub HeliArm(enabled)
    If enabled Then
    heli_animation.enabled = 1
    End If
End Sub

Dim mdirection

Sub HeliArmRelay(enabled)
    If enabled Then
        mdirection = 0
        MF.State = 0
        MB.State = 1
    Else
        MF.State = 1
        MB.State = 0
        mdirection = 1
    End If
End Sub

Sub Blades_Motor(enabled)
    If enabled Then
        Blades_Spin.enabled = 1
    Else
        Blades_Spin.enabled = 0
    End If
End Sub

Sub BallLeftShooter(Enabled)
    If Enabled Then
        Heli_Kick.Kick 0, 34
            PlaySound SoundFX("popper_ball", DOFContactors)
            If hbin = 1 Then
        PlaySound SoundFX("popper_ball", DOFContactors)
        Heli_Raise.Kick 0, 34
            End If
        Controller.Switch(80) = 0
        hbin = 0
    End If
End Sub

Sub BallLiftMotor(Enabled)
    If Enabled And hbin = 1 Then
        Heli_Raise.DestroyBall
    Set cball = Heli_Raise.createball
    Raise_Post.enabled = 1
        Heli_Raise.enabled = 0
    End If
End Sub

Dim emon
Sub ElectroMagnet(enabled)
    If enabled Then
        emon = 1
        Light101.State = 1
    Else
        If Controller.Switch(100) = 0 And emon = 1 Then
            Light101.State = 0
            b3active = 0
            Heli_Raise.Kick 0, 0
      hbout = 1
            BOut = 1
            emon = 0
            Heli_Raise.enabled = 1
            Heli_Kick.enabled = 0
        End If
        Light101.State = 0
    End If
End Sub


'********************************************


  Sub Heli_Opto_Watch_Timer()
    'debug.print bsTrough.Balls
    If Controller.Switch(80) = True Then
      L80.State = 1
    Else
      L80.State = 0
    End If
    If Controller.Switch(33) = True Then
      Lsw33.State = 1
    Else
      Lsw33.State = 0
    End If
    If Controller.Switch(23) = True Then
      Lsw23.State = 1
    Else
      Lsw23.State = 0
    End If
    If Controller.Switch(200) = True Then
      Lsw400.State = 1
    Else
      Lsw400.State = 0
    End If
    If Controller.Switch(201) = True Then
      Lsw401.State = 1
    Else
      Lsw401.State = 0
    End If

    If Controller.Switch(34) = True Then
      BLD.State = 1
    Else
      BLD.State = 0
    End If

    If Controller.Switch(24) = True Then
      BLU.State = 1
    Else
      BLU.State = 0
    End If

    If Heli_Base.ObjRotz >= 0 And Heli_Base.ObjRotz <= 2 Then
      Controller.Switch(100) = 1
      helihome = 1
      heliaway = 0
      HH.State = 1
      HA.State = 0
    Else
      Controller.Switch(100) = 0
      HH.State = 0
    End If
    If Heli_Base.ObjRotz >= 175 And Heli_Base.ObjRotz <= 190 Then
      Controller.Switch(90) = 1
      heliaway = 1
      helihome = 0
      HH.State = 0
      HA.State = 1
    Else
      Controller.Switch(90) = 0
      HA.State = 0
    End If

  End Sub



'******************************************************
'   FLIPPER CORRECTION INITIALIZATION
'******************************************************


dim LF : Set LF = New FlipperPolarity
dim RF : Set RF = New FlipperPolarity

InitPolarity

Sub InitPolarity()
        dim x, a : a = Array(LF, RF)
        for each x in a
                x.AddPoint "Ycoef", 0, RightFlipper.Y-65, 1        'disabled
                x.AddPoint "Ycoef", 1, RightFlipper.Y-11, 1
                x.enabled = True
                x.TimeDelay = 60
        Next

        AddPt "Polarity", 0, 0, 0
        AddPt "Polarity", 1, 0.05, -5.5
        AddPt "Polarity", 2, 0.4, -5.5
        AddPt "Polarity", 3, 0.6, -5.0
        AddPt "Polarity", 4, 0.65, -4.5
        AddPt "Polarity", 5, 0.7, -4.0
        AddPt "Polarity", 6, 0.75, -3.5
        AddPt "Polarity", 7, 0.8, -3.0
        AddPt "Polarity", 8, 0.85, -2.5
        AddPt "Polarity", 9, 0.9,-2.0
        AddPt "Polarity", 10, 0.95, -1.5
        AddPt "Polarity", 11, 1, -1.0
        AddPt "Polarity", 12, 1.05, -0.5
        AddPt "Polarity", 13, 1.1, 0
        AddPt "Polarity", 14, 1.3, 0

        addpt "Velocity", 0, 0,         1
        addpt "Velocity", 1, 0.16, 1.06
        addpt "Velocity", 2, 0.41,         1.05
        addpt "Velocity", 3, 0.53,         1'0.982
        addpt "Velocity", 4, 0.702, 0.968
        addpt "Velocity", 5, 0.95,  0.968
        addpt "Velocity", 6, 1.03,         0.945

        LF.Object = LeftFlipper
        LF.EndPoint = EndPointLp
        RF.Object = RightFlipper
        RF.EndPoint = EndPointRp
End Sub

Sub TriggerLF_Hit() : LF.Addball activeball : End Sub
Sub TriggerLF_UnHit() : LF.PolarityCorrect activeball : End Sub
Sub TriggerRF_Hit() : RF.Addball activeball : End Sub
Sub TriggerRF_UnHit() : RF.PolarityCorrect activeball : End Sub


'******************************************************
'                        FLIPPER CORRECTION FUNCTIONS
'******************************************************

Sub AddPt(aStr, idx, aX, aY)        'debugger wrapper for adjusting flipper script in-game
        dim a : a = Array(LF, RF)
        dim x : for each x in a
                x.addpoint aStr, idx, aX, aY
        Next
End Sub

Class FlipperPolarity
        Public DebugOn, Enabled
        Private FlipAt        'Timer variable (IE 'flip at 723,530ms...)
        Public TimeDelay        'delay before trigger turns off and polarity is disabled TODO set time!
        private Flipper, FlipperStart,FlipperEnd, FlipperEndY, LR, PartialFlipCoef
        Private Balls(20), balldata(20)

        dim PolarityIn, PolarityOut
        dim VelocityIn, VelocityOut
        dim YcoefIn, YcoefOut
        Public Sub Class_Initialize
                redim PolarityIn(0) : redim PolarityOut(0) : redim VelocityIn(0) : redim VelocityOut(0) : redim YcoefIn(0) : redim YcoefOut(0)
                Enabled = True : TimeDelay = 50 : LR = 1:  dim x : for x = 0 to uBound(balls) : balls(x) = Empty : set Balldata(x) = new SpoofBall : next
        End Sub

        Public Property let Object(aInput) : Set Flipper = aInput : StartPoint = Flipper.x : End Property
        Public Property Let StartPoint(aInput) : if IsObject(aInput) then FlipperStart = aInput.x else FlipperStart = aInput : end if : End Property
        Public Property Get StartPoint : StartPoint = FlipperStart : End Property
        Public Property Let EndPoint(aInput) : FlipperEnd = aInput.x: FlipperEndY = aInput.y: End Property
        Public Property Get EndPoint : EndPoint = FlipperEnd : End Property
        Public Property Get EndPointY: EndPointY = FlipperEndY : End Property

        Public Sub AddPoint(aChooseArray, aIDX, aX, aY) 'Index #, X position, (in) y Position (out)
                Select Case aChooseArray
                        case "Polarity" : ShuffleArrays PolarityIn, PolarityOut, 1 : PolarityIn(aIDX) = aX : PolarityOut(aIDX) = aY : ShuffleArrays PolarityIn, PolarityOut, 0
                        Case "Velocity" : ShuffleArrays VelocityIn, VelocityOut, 1 :VelocityIn(aIDX) = aX : VelocityOut(aIDX) = aY : ShuffleArrays VelocityIn, VelocityOut, 0
                        Case "Ycoef" : ShuffleArrays YcoefIn, YcoefOut, 1 :YcoefIn(aIDX) = aX : YcoefOut(aIDX) = aY : ShuffleArrays YcoefIn, YcoefOut, 0
                End Select
                if gametime > 100 then Report aChooseArray
        End Sub

        Public Sub Report(aChooseArray)         'debug, reports all coords in tbPL.text
                if not DebugOn then exit sub
                dim a1, a2 : Select Case aChooseArray
                        case "Polarity" : a1 = PolarityIn : a2 = PolarityOut
                        Case "Velocity" : a1 = VelocityIn : a2 = VelocityOut
                        Case "Ycoef" : a1 = YcoefIn : a2 = YcoefOut
                        case else :tbpl.text = "wrong string" : exit sub
                End Select
                dim str, x : for x = 0 to uBound(a1) : str = str & aChooseArray & " x: " & round(a1(x),4) & ", " & round(a2(x),4) & vbnewline : next
                tbpl.text = str
        End Sub

        Public Sub AddBall(aBall) : dim x : for x = 0 to uBound(balls) : if IsEmpty(balls(x)) then set balls(x) = aBall : exit sub :end if : Next  : End Sub

        Private Sub RemoveBall(aBall)
                dim x : for x = 0 to uBound(balls)
                        if TypeName(balls(x) ) = "IBall" then
                                if aBall.ID = Balls(x).ID Then
                                        balls(x) = Empty
                                        Balldata(x).Reset
                                End If
                        End If
                Next
        End Sub

        Public Sub Fire()
                Flipper.RotateToEnd
                processballs
        End Sub

        Public Property Get Pos 'returns % position a ball. For debug stuff.
                dim x : for x = 0 to uBound(balls)
                        if not IsEmpty(balls(x) ) then
                                pos = pSlope(Balls(x).x, FlipperStart, 0, FlipperEnd, 1)
                        End If
                Next
        End Property

        Public Sub ProcessBalls() 'save data of balls in flipper range
                FlipAt = GameTime
                dim x : for x = 0 to uBound(balls)
                        if not IsEmpty(balls(x) ) then
                                balldata(x).Data = balls(x)
                        End If
                Next
                PartialFlipCoef = ((Flipper.StartAngle - Flipper.CurrentAngle) / (Flipper.StartAngle - Flipper.EndAngle))
                PartialFlipCoef = abs(PartialFlipCoef-1)
        End Sub
        Private Function FlipperOn() : if gameTime < FlipAt+TimeDelay then FlipperOn = True : End If : End Function        'Timer shutoff for polaritycorrect

        Public Sub PolarityCorrect(aBall)
                if FlipperOn() then
                        dim tmp, BallPos, x, IDX, Ycoef : Ycoef = 1

                        'y safety Exit
                        if aBall.VelY > -8 then 'ball going down
                                RemoveBall aBall
                                exit Sub
                        end if

                        'Find balldata. BallPos = % on Flipper
                        for x = 0 to uBound(Balls)
                                if aBall.id = BallData(x).id AND not isempty(BallData(x).id) then
                                        idx = x
                                        BallPos = PSlope(BallData(x).x, FlipperStart, 0, FlipperEnd, 1)
                                        if ballpos > 0.65 then  Ycoef = LinearEnvelope(BallData(x).Y, YcoefIn, YcoefOut)                                'find safety coefficient 'ycoef' data
                                end if
                        Next

                        If BallPos = 0 Then 'no ball data meaning the ball is entering and exiting pretty close to the same position, use current values.
                                BallPos = PSlope(aBall.x, FlipperStart, 0, FlipperEnd, 1)
                                if ballpos > 0.65 then  Ycoef = LinearEnvelope(aBall.Y, YcoefIn, YcoefOut)                                                'find safety coefficient 'ycoef' data
                        End If

                        'Velocity correction
                        if not IsEmpty(VelocityIn(0) ) then
                                Dim VelCoef
         :                         VelCoef = LinearEnvelope(BallPos, VelocityIn, VelocityOut)

                                if partialflipcoef < 1 then VelCoef = PSlope(partialflipcoef, 0, 1, 1, VelCoef)

                                if Enabled then aBall.Velx = aBall.Velx*VelCoef
                                if Enabled then aBall.Vely = aBall.Vely*VelCoef
                        End If

                        'Polarity Correction (optional now)
                        if not IsEmpty(PolarityIn(0) ) then
                                If StartPoint > EndPoint then LR = -1        'Reverse polarity if left flipper
                                dim AddX : AddX = LinearEnvelope(BallPos, PolarityIn, PolarityOut) * LR

                                if Enabled then aBall.VelX = aBall.VelX + 1 * (AddX*ycoef*PartialFlipcoef)
                                'playsound "fx_knocker"
                        End If
                End If
                RemoveBall aBall
        End Sub
End Class



'******************************************************
'                        FLIPPER TRICKS
'******************************************************

RightFlipper.timerinterval=1
Rightflipper.timerenabled=True

sub RightFlipper_timer()
        FlipperTricks LeftFlipper, LFPress, LFCount, LFEndAngle, LFState
        FlipperTricks RightFlipper, RFPress, RFCount, RFEndAngle, RFState
        FlipperNudge RightFlipper, RFEndAngle, RFEOSNudge, LeftFlipper, LFEndAngle
        FlipperNudge LeftFlipper, LFEndAngle, LFEOSNudge,  RightFlipper, RFEndAngle
end sub

Dim LFEOSNudge, RFEOSNudge

Sub FlipperNudge(Flipper1, Endangle1, EOSNudge1, Flipper2, EndAngle2)
        Dim BOT, b

        If Flipper1.currentangle = Endangle1 and EOSNudge1 <> 1 Then
                EOSNudge1 = 1
                'debug.print Flipper1.currentangle &" = "& Endangle1 &"--"& Flipper2.currentangle &" = "& EndAngle2
                If Flipper2.currentangle = EndAngle2 Then
                        BOT = GetBalls
                        For b = 0 to Ubound(BOT)
                                If FlipperTrigger(BOT(b).x, BOT(b).y, Flipper1) Then
                                        'Debug.Print "ball in flip1. exit"
                                         exit Sub
                                end If
                        Next
                        For b = 0 to Ubound(BOT)
                                If FlipperTrigger(BOT(b).x, BOT(b).y, Flipper2) Then
                                        BOT(b).velx = BOT(b).velx / 1.7
                                        BOT(b).vely = BOT(b).vely - 1
                                end If
                        Next
                End If
        Else
                If Flipper1.currentangle <> EndAngle1 then
                        EOSNudge1 = 0
                end if
        End If
End Sub

'*****************
' Maths
'*****************
Const PI = 3.1415927

Function dSin(degrees)
        dsin = sin(degrees * Pi/180)
End Function

Function dCos(degrees)
        dcos = cos(degrees * Pi/180)
End Function

'*************************************************
' Check ball distance from Flipper for Rem
'*************************************************

Function Distance(ax,ay,bx,by)
        Distance = SQR((ax - bx)^2 + (ay - by)^2)
End Function

Function DistancePL(px,py,ax,ay,bx,by) ' Distance between a point and a line where point is px,py
        DistancePL = ABS((by - ay)*px - (bx - ax) * py + bx*ay - by*ax)/Distance(ax,ay,bx,by)
End Function

Function Radians(Degrees)
        Radians = Degrees * PI /180
End Function

Function AnglePP(ax,ay,bx,by)
        AnglePP = Atn2((by - ay),(bx - ax))*180/PI
End Function

Function DistanceFromFlipper(ballx, bally, Flipper)
        DistanceFromFlipper = DistancePL(ballx, bally, Flipper.x, Flipper.y, Cos(Radians(Flipper.currentangle+90))+Flipper.x, Sin(Radians(Flipper.currentangle+90))+Flipper.y)
End Function

Function FlipperTrigger(ballx, bally, Flipper)
        Dim DiffAngle
        DiffAngle  = ABS(Flipper.currentangle - AnglePP(Flipper.x, Flipper.y, ballx, bally) - 90)
        If DiffAngle > 180 Then DiffAngle = DiffAngle - 360

        If DistanceFromFlipper(ballx,bally,Flipper) < 48 and DiffAngle <= 90 and Distance(ballx,bally,Flipper.x,Flipper.y) < Flipper.Length Then
                FlipperTrigger = True
        Else
                FlipperTrigger = False
        End If
End Function

' Used for drop targets and stand up targets
Function Atn2(dy, dx)
        dim pi
        pi = 4*Atn(1)

        If dx > 0 Then
                Atn2 = Atn(dy / dx)
        ElseIf dx < 0 Then
                If dy = 0 Then
                        Atn2 = pi
                Else
                        Atn2 = Sgn(dy) * (pi - Atn(Abs(dy / dx)))
                end if
        ElseIf dx = 0 Then
                if dy = 0 Then
                        Atn2 = 0
                else
                        Atn2 = Sgn(dy) * pi / 2
                end if
        End If
End Function

'*************************************************
' End - Check ball distance from Flipper for Rem
'*************************************************

dim LFPress, RFPress, LFCount, RFCount
dim LFState, RFState
dim EOST, EOSA,Frampup, FElasticity,FReturn
dim RFEndAngle, LFEndAngle

EOST = leftflipper.eostorque
EOSA = leftflipper.eostorqueangle
Frampup = LeftFlipper.rampup
FElasticity = LeftFlipper.elasticity
FReturn = LeftFlipper.return
Const EOSTnew = 0.8
Const EOSAnew = 1
Const EOSRampup = 0
Dim SOSRampup
SOSRampup = 2.5
Const LiveCatch = 16
Const LiveElasticity = 0.45
Const SOSEM = 0.815
Const EOSReturn = 0.025

LFEndAngle = Leftflipper.endangle
RFEndAngle = RightFlipper.endangle

Sub FlipperActivate(Flipper, FlipperPress)
        FlipperPress = 1
        Flipper.Elasticity = FElasticity

        Flipper.eostorque = EOST
        Flipper.eostorqueangle = EOSA
End Sub

Sub FlipperDeactivate(Flipper, FlipperPress)
        FlipperPress = 0
        Flipper.eostorqueangle = EOSA
        Flipper.eostorque = EOST*EOSReturn/FReturn


        If Abs(Flipper.currentangle) <= Abs(Flipper.endangle) + 0.1 Then
                Dim BOT, b
                BOT = GetBalls

                For b = 0 to UBound(BOT)
                        If Distance(BOT(b).x, BOT(b).y, Flipper.x, Flipper.y) < 55 Then 'check for cradle
                                If BOT(b).vely >= -0.4 Then BOT(b).vely = -0.4
                        End If
                Next
        End If
End Sub

Sub FlipperTricks (Flipper, FlipperPress, FCount, FEndAngle, FState)
        Dim Dir
        Dir = Flipper.startangle/Abs(Flipper.startangle)        '-1 for Right Flipper

        If Abs(Flipper.currentangle) > Abs(Flipper.startangle) - 0.05 Then
                If FState <> 1 Then
                        Flipper.rampup = SOSRampup
                        Flipper.endangle = FEndAngle - 3*Dir
                        Flipper.Elasticity = FElasticity * SOSEM
                        FCount = 0
                        FState = 1
                End If
        ElseIf Abs(Flipper.currentangle) <= Abs(Flipper.endangle) and FlipperPress = 1 then
                if FCount = 0 Then FCount = GameTime

                If FState <> 2 Then
                        Flipper.eostorqueangle = EOSAnew
                        Flipper.eostorque = EOSTnew
                        Flipper.rampup = EOSRampup
                        Flipper.endangle = FEndAngle
                        FState = 2
                End If
        Elseif Abs(Flipper.currentangle) > Abs(Flipper.endangle) + 0.01 and FlipperPress = 1 Then
                If FState <> 3 Then
                        Flipper.eostorque = EOST
                        Flipper.eostorqueangle = EOSA
                        Flipper.rampup = Frampup
                        Flipper.Elasticity = FElasticity
                        FState = 3
                End If

        End If
End Sub

Const LiveDistanceMin = 30  'minimum distance in vp units from flipper base live catch dampening will occur
Const LiveDistanceMax = 114  'maximum distance in vp units from flipper base live catch dampening will occur (tip protection)

Sub CheckLiveCatch(ball, Flipper, FCount, parm) 'Experimental new live catch
        Dim Dir
        Dir = Flipper.startangle/Abs(Flipper.startangle)    '-1 for Right Flipper
        Dim LiveCatchBounce                                                                                                                        'If live catch is not perfect, it won't freeze ball totally
        Dim CatchTime : CatchTime = GameTime - FCount

        if CatchTime <= LiveCatch and parm > 6 and ABS(Flipper.x - ball.x) > LiveDistanceMin and ABS(Flipper.x - ball.x) < LiveDistanceMax Then
                if CatchTime <= LiveCatch*0.5 Then                                                'Perfect catch only when catch time happens in the beginning of the window
                        LiveCatchBounce = 0
                else
                        LiveCatchBounce = Abs((LiveCatch/2) - CatchTime)        'Partial catch when catch happens a bit late
                end If

                If LiveCatchBounce = 0 and ball.velx * Dir > 0 Then ball.velx = 0
                ball.vely = LiveCatchBounce * (32 / LiveCatch) ' Multiplier for inaccuracy bounce
                ball.angmomx= 0
                ball.angmomy= 0
                ball.angmomz= 0
        End If
End Sub

'******************************************************
'   FLIPPER POLARITY, DAMPENER, AND DROP TARGET
'       SUPPORTING FUNCTIONS
'******************************************************

' Used for flipper correction and rubber dampeners
Sub ShuffleArray(ByRef aArray, byVal offset) 'shuffle 1d array
  dim x, aCount : aCount = 0
  redim a(uBound(aArray) )
  for x = 0 to uBound(aArray) 'Shuffle objects in a temp array
    if not IsEmpty(aArray(x) ) Then
      if IsObject(aArray(x)) then
        Set a(aCount) = aArray(x)
      Else
        a(aCount) = aArray(x)
      End If
      aCount = aCount + 1
    End If
  Next
  if offset < 0 then offset = 0
  redim aArray(aCount-1+offset) 'Resize original array
  for x = 0 to aCount-1   'set objects back into original array
    if IsObject(a(x)) then
      Set aArray(x) = a(x)
    Else
      aArray(x) = a(x)
    End If
  Next
End Sub

' Used for flipper correction and rubber dampeners
Sub ShuffleArrays(aArray1, aArray2, offset)
  ShuffleArray aArray1, offset
  ShuffleArray aArray2, offset
End Sub

' Used for flipper correction, rubber dampeners, and drop targets
Function BallSpeed(ball) 'Calculates the ball speed
    BallSpeed = SQR(ball.VelX^2 + ball.VelY^2 + ball.VelZ^2)
End Function

' Used for flipper correction and rubber dampeners
Function PSlope(Input, X1, Y1, X2, Y2)  'Set up line via two points, no clamping. Input X, output Y
  dim x, y, b, m : x = input : m = (Y2 - Y1) / (X2 - X1) : b = Y2 - m*X2
  Y = M*x+b
  PSlope = Y
End Function

' Used for flipper correction
Class spoofball
  Public X, Y, Z, VelX, VelY, VelZ, ID, Mass, Radius
  Public Property Let Data(aBall)
    With aBall
      x = .x : y = .y : z = .z : velx = .velx : vely = .vely : velz = .velz
      id = .ID : mass = .mass : radius = .radius
    end with
  End Property
  Public Sub Reset()
    x = Empty : y = Empty : z = Empty  : velx = Empty : vely = Empty : velz = Empty
    id = Empty : mass = Empty : radius = Empty
  End Sub
End Class

' Used for flipper correction and rubber dampeners
Function LinearEnvelope(xInput, xKeyFrame, yLvl)
  dim y 'Y output
  dim L 'Line
  dim ii : for ii = 1 to uBound(xKeyFrame)  'find active line
    if xInput <= xKeyFrame(ii) then L = ii : exit for : end if
  Next
  if xInput > xKeyFrame(uBound(xKeyFrame) ) then L = uBound(xKeyFrame)  'catch line overrun
  Y = pSlope(xInput, xKeyFrame(L-1), yLvl(L-1), xKeyFrame(L), yLvl(L) )

  if xInput <= xKeyFrame(lBound(xKeyFrame) ) then Y = yLvl(lBound(xKeyFrame) )  'Clamp lower
  if xInput >= xKeyFrame(uBound(xKeyFrame) ) then Y = yLvl(uBound(xKeyFrame) )  'Clamp upper

  LinearEnvelope = Y
End Function

' Used for drop targets
Function Atn2(dy, dx)
  dim pi
  pi = 4*Atn(1)

  If dx > 0 Then
    Atn2 = Atn(dy / dx)
  ElseIf dx < 0 Then
    If dy = 0 Then
      Atn2 = pi
    Else
      Atn2 = Sgn(dy) * (pi - Atn(Abs(dy / dx)))
    end if
  ElseIf dx = 0 Then
    if dy = 0 Then
      Atn2 = 0
    else
      Atn2 = Sgn(dy) * pi / 2
    end if
  End If
End Function

' Used for drop targets
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

' Used for drop targets
Function Distance(ax,ay,bx,by)
  Distance = SQR((ax - bx)^2 + (ay - by)^2)
End Function

'****************************************************************************
'PHYSICS DAMPENERS

'These are data mined bounce curves,
'dialed in with the in-game elasticity as much as possible to prevent angle / spin issues.
'Requires tracking ballspeed to calculate COR


Sub dPosts_Hit(idx)
  RubbersD.dampen Activeball
End Sub

Sub dSleeves_Hit(idx)
  SleevesD.Dampen Activeball
End Sub


dim RubbersD : Set RubbersD = new Dampener  'frubber
RubbersD.name = "Rubbers"
RubbersD.debugOn = False  'shows info in textbox "TBPout"
RubbersD.Print = False  'debug, reports in debugger (in vel, out cor)
'cor bounce curve (linear)
'for best results, try to match in-game velocity as closely as possible to the desired curve
'RubbersD.addpoint 0, 0, 0.935  'point# (keep sequential), ballspeed, CoR (elasticity)
RubbersD.addpoint 0, 0, 0.96  'point# (keep sequential), ballspeed, CoR (elasticity)
RubbersD.addpoint 1, 3.77, 0.96
RubbersD.addpoint 2, 5.76, 0.967  'dont take this as gospel. if you can data mine rubber elasticitiy, please help!
RubbersD.addpoint 3, 15.84, 0.874
RubbersD.addpoint 4, 56, 0.64 'there's clamping so interpolate up to 56 at least

dim SleevesD : Set SleevesD = new Dampener  'this is just rubber but cut down to 85%...
SleevesD.name = "Sleeves"
SleevesD.debugOn = False  'shows info in textbox "TBPout"
SleevesD.Print = False  'debug, reports in debugger (in vel, out cor)
SleevesD.CopyCoef RubbersD, 0.85

Class Dampener
  Public Print, debugOn 'tbpOut.text
  public name, Threshold  'Minimum threshold. Useful for Flippers, which don't have a hit threshold.
  Public ModIn, ModOut
  Private Sub Class_Initialize : redim ModIn(0) : redim Modout(0): End Sub

  Public Sub AddPoint(aIdx, aX, aY)
    ShuffleArrays ModIn, ModOut, 1 : ModIn(aIDX) = aX : ModOut(aIDX) = aY : ShuffleArrays ModIn, ModOut, 0
    if gametime > 100 then Report
  End Sub

  public sub Dampen(aBall)
    if threshold then if BallSpeed(aBall) < threshold then exit sub end if end if
    dim RealCOR, DesiredCOR, str, coef
    DesiredCor = LinearEnvelope(cor.ballvel(aBall.id), ModIn, ModOut )
    RealCOR = BallSpeed(aBall) / cor.ballvel(aBall.id)
    coef = desiredcor / realcor
    if debugOn then str = name & " in vel:" & round(cor.ballvel(aBall.id),2 ) & vbnewline & "desired cor: " & round(desiredcor,4) & vbnewline & _
    "actual cor: " & round(realCOR,4) & vbnewline & "ballspeed coef: " & round(coef, 3) & vbnewline
    if Print then debug.print Round(cor.ballvel(aBall.id),2) & ", " & round(desiredcor,3)

' Thalamus - patched : ' Thalamus - patched :     aBall.velx = aBall.velx * coef : aBall.vely = aBall.vely * coef
    aBall.velx = aBall.velx * coef : aBall.vely = aBall.vely * coef : aBall.velz = aBall.velz * coef
' Thalamus - patched :     aBall.velx = aBall.velx * coef : aBall.vely = aBall.vely * coef : aBall.velz = aBall.velz * coef
    aBall.velx = aBall.velx * coef : aBall.vely = aBall.vely * coef : aBall.velz = aBall.velz * coef
    if debugOn then TBPout.text = str
  End Sub

  Public Sub CopyCoef(aObj, aCoef) 'alternative addpoints, copy with coef
    dim x : for x = 0 to uBound(aObj.ModIn)
      addpoint x, aObj.ModIn(x), aObj.ModOut(x)*aCoef
    Next
  End Sub


  Public Sub Report()   'debug, reports all coords in tbPL.text
    if not debugOn then exit sub
    dim a1, a2 : a1 = ModIn : a2 = ModOut
    dim str, x : for x = 0 to uBound(a1) : str = str & x & ": " & round(a1(x),4) & ", " & round(a2(x),4) & vbnewline : next
    TBPout.text = str
  End Sub


End Class

'******************************************************
'   TRACK ALL BALL VELOCITIES
'     FOR RUBBER DAMPENER AND DROP TARGETS
'******************************************************

dim cor : set cor = New CoRTracker

Class CoRTracker
  public ballvel, ballvelx, ballvely

  Private Sub Class_Initialize : redim ballvel(0) : redim ballvelx(0): redim ballvely(0) : End Sub

  Public Sub Update() 'tracks in-ball-velocity
    dim str, b, AllBalls, highestID : allBalls = getballs

    for each b in allballs
      if b.id >= HighestID then highestID = b.id
    Next

    if uBound(ballvel) < highestID then redim ballvel(highestID)  'set bounds
    if uBound(ballvelx) < highestID then redim ballvelx(highestID)  'set bounds
    if uBound(ballvely) < highestID then redim ballvely(highestID)  'set bounds

    for each b in allballs
      ballvel(b.id) = BallSpeed(b)
      ballvelx(b.id) = b.velx
      ballvely(b.id) = b.vely
    Next
  End Sub
End Class

Sub RDampen_Timer()
  Cor.Update
End Sub

' Thalamus : Exit in a clean and proper way
Sub Table1_exit()
  On Error Resume Next
  Controller.Pause = False
  Controller.Stop
End Sub

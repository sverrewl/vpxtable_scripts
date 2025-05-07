'Who Dunnit - IPDB No. 3685
'© Bally(Midway) 1995
'VPX recreation by ninuzzu
'thanks to DJRobX for the help and to the VPDev Team for the amazing VPX!

' Thalamus 2018-07-24
' Tables has its own "Positional Sound Playback Functions" and "Supporting Ball & Sound Functions"
' Changed UseSolenoids=1 to 2
' Thalamus 2018-12-17 : Added FFv2
' No special SSF tweaks yet.


Option Explicit
Randomize

Const FlasherIntensityGIOn = .75
Const FlasherIntensityGIOff = 1

Const BallSize = 51
Const BallMass = 1.3

On Error Resume Next
ExecuteGlobal GetTextFile("controller.vbs")
If Err Then MsgBox "You need the controller.vbs in order to run this table, available in the vp10 package"
On Error Goto 0

Dim DesktopMode:DesktopMode = Table1.ShowDT
Dim UseVPMDMD:UseVPMDMD = DesktopMode
Const UseVPMModSol = 0

' Rom Name
Const cGameName = "wd_12"

LoadVPM "02000000", "WPC.VBS", 3.50

' Thalamus - for Fast Flip v2
NoUpperRightFlipper
NoUpperLeftFlipper


' Standard Options
Const UseSolenoids = 2
Const UseLamps = 1
Const UseSync = 1
Const HandleMech = 0

' Standard Sounds
Const SSolenoidOn = "fx_Solenoid"
Const SSolenoidOff = ""
Const SCoin = "fx_Coin"



'************************************************************************
'            INIT TABLE
'************************************************************************

Dim bsTrough, bsLSaucer, bsRBPopper, bsRFPopper, plungerIM

Sub Table1_Init
  vpmInit Me
  With Controller
    .GameName = cGameName
    If Err Then MsgBox "Can't start Game " & cGameName & vbNewLine & Err.Description:Exit Sub
    .SplashInfoLine = "Who Dunnit (Bally 1995)"
    .HandleKeyboard = 0
    .ShowTitle = 0
    .ShowDMDOnly = 1
    .ShowFrame = 0
    .HandleMechanics = 0
    .Hidden = DesktopMode
    On Error Resume Next
    .Run GetPlayerHWnd
    If Err Then MsgBox Err.Description
    On Error Goto 0
    .Switch(22) = 1 'close coin door
    .Switch(24) = 0 'always closed
  End With

    ' Main Timer init
    PinMAMETimer.Interval = PinMAMEInterval
    PinMAMETimer.Enabled = 1

    ' Nudging
    vpmNudge.TiltSwitch = 14
    vpmNudge.Sensitivity = 5
    vpmNudge.TiltObj = Array(sw61, sw62, sw63, sw64, sw65)

    ' Trough
    Set bsTrough = New cvpmTrough
    With bsTrough
    .size = 4
    .initSwitches Array(32, 33, 34, 35)
    .Initexit BallRelease, 90, 6
    .InitEntrySounds "drain", "", ""
    .InitExitSounds SoundFX(SSolenoidOn,DOFContactors), SoundFX("BallRelease",DOFContactors)
    .Balls = 4
    .CreateEvents "bsTrough", Drain
    End With

  ' Left Lock Up
  Set bsLSaucer = New cvpmSaucer
  With bsLSaucer
    .InitKicker sw51, 51, 180, 10, 1
        .InitExitVariance 1, 1
    .InitSounds "fx_saucerHit", SoundFX(SSolenoidOn,DOFContactors), SoundFX("fx_saucer_exit",DOFContactors)
    .CreateEvents "bsLSaucer", sw51
  End With

  ' Right Back Popper
  Set bsRBPopper = New cvpmSaucer
  With bsRBPopper
    .InitKicker sw43, 43, 0, 35, 1.56
        .InitExitVariance 1, 1
    .InitSounds "fx_saucerHit", SoundFX(SSolenoidOn,DOFContactors), SoundFX("right_back_kicker_out",DOFContactors)
    .CreateEvents "bsRBPopper", sw43
  End With

    ' Right Front Popper
    Set bsRFPopper = New cvpmTrough
  With bsRFPopper
    .size = 2
    .balls = 0
    .InitSwitches Array(44,57)
    .Initexit sw44, 210, 10
    .InitEntrySounds "fx_kicker_catch", "", ""
    .InitExitSounds SoundFX(SSolenoidOn,DOFContactors),SoundFX("right_front_kicker_out",DOFContactors)
    .CreateEvents "bsRFPopper", sw57
  End With

  'Impulse Plunger
  Set plungerIM = New cvpmImpulseP
  With plungerIM
    .InitImpulseP sw15, 45, 0.7
    .Switch 15
    .Random 0.05
    .InitExitSnd SoundFX("AutoPlunger",DOFContactors), SoundFX("AutoPlunger",DOFContactors)
    .CreateEvents "plungerIM"
  End With

  SlotInit:TargetsInit:RampInit:mapLights
  UpdateGI 0,0:UpdateGI 1,0:UpdateGI 2,0

  If DesktopMode Then
    CabRails.visible = 1
    'lower right flasher position
    flasherflash4.Y=1443.816
    flasherflash4.RotX=-65
    flasherflash4.Height=185
    'spinner flasher position
    flasherflash5.Height=210
  Else
    CabRails.visible = 0
    flasherflash4.Y=1439
    flasherflash4.RotX=-45
    flasherflash4.Height=195
    Flasherflash5.Height=245
    GI_SideR.opacity=150
    GI_SideTR.opacity=150
  End If
End Sub

'************************************************************************
'             KEYS
'************************************************************************

Sub Table1_KeyDown(ByVal Keycode)
  If keycode = keyFront Then Controller.Switch(23) = 1    'buy-in
  If Keycode = PlungerKey Then Plunger.Pullback:PlaySoundAt "fx_plungerpull",Plunger
  If keycode = LeftTiltKey Then Nudge 90, 5:PlaySoundAt SoundFX("fx_nudge",0),sw27
  If keycode = RightTiltKey Then Nudge 270, 5:PlaySoundAt SoundFX("fx_nudge",0),Drain
  If keycode = CenterTiltKey Then Nudge 0, 6:PlaySoundAt SoundFX("fx_nudge",0),sw16
    If vpmKeyDown(keycode) Then Exit Sub
End Sub

Sub Table1_KeyUp(ByVal Keycode)
  If keycode = keyFront Then Controller.Switch(23) = 0    'buy-in
  If Keycode = PlungerKey Then Plunger.Fire:StopSound "fx_plungerpull":PlaySoundAt "fx_plunger",Plunger
  If vpmKeyUp(keycode) Then Exit Sub
End Sub

Sub Table1_Paused:Controller.Pause = True:End Sub
Sub Table1_unPaused:Controller.Pause = False:End Sub
Sub Table1_exit():Controller.Pause = False:Controller.Stop:End Sub

'*****************************************
'         General Illumination
'*****************************************
Set GiCallBack2 = GetRef("UpdateGi")

Sub UpdateGi(nr,step)
  Dim ii
  Select Case nr

  Case 0    'Left Playfield
    If step=0 Then
      For each ii in GI_Left:ii.state=0:Next
    Else
      For each ii in GI_Left:ii.state=1:Next
    End If
    For each ii in GI_Left:ii.IntensityScale = 0.125 * step:Next
    GI_SideL.IntensityScale = 0.125 * step
    GI_SideL1.IntensityScale = 0.125 * step

  Case 1    'Right Playfield
    If step=0 Then
      For each ii in GI_Right:ii.state=0:Next
    Else
      For each ii in GI_Right:ii.state=1:Next
    End If
    For each ii in GI_Right:ii.IntensityScale = 0.125 * step:Next
    GI_SideR.IntensityScale = 0.125 * step

  Case 2    'Top Playfield
    If step=0 Then
      For each ii in GI_Top:ii.state=0:Next
    Else
      For each ii in GI_Top:ii.state=1:Next
    End If
    BWR.IntensityScale = 0.125 * step
    BWL.IntensityScale = 0.125 * step
    GI_SideTL.IntensityScale = 0.125 * step
    GI_SideTR.IntensityScale = 0.125 * step
    For each ii in GI_Top:ii.IntensityScale = 0.125 * step:Next
    If Step>=7 Then Table1.ColorGradeImage = "ColorGrade_8":Else Table1.ColorGradeImage = "ColorGrade_" & (step+1):End If
    If step>4 Then DOF 103, DOFOn : Else DOF 103, DOFOff:End If

  Case 3    'Insert 1

  Case 4    'Insert 2

  End Select
End Sub

'*****************************************
'       Lights Mapping
'*****************************************
Sub mapLights
  Dim obj
  For Each obj In Lamps
    set Lights(CInt(mid(obj.Name,2))) = obj
  Next
  Lights(16)= Array(l16,l16a,l16b,l16c)
  Lights(17)= Array(l17,l17a,l17b,l17c)
  Lights(18)= Array(l18,l18a,l18b)
End Sub

Sub SpotLightsUpdate
  L83.state = Controller.Lamp(83)
  L83a.Visible = Controller.Lamp(83)
  L84.state = Controller.Lamp(84)
  L84a.Visible = Controller.Lamp(84)
End Sub

'*****************************************
'         Solenoids Mapping
'*****************************************
SolCallback(1)= "SolTrough"         '1-Trough
SolCallback(2)= "SolAutoPlungerIM"      '2-AutoPlunger
SolCallback(3)= "bsLSaucer.SolOut"      '3-Left Lock Up
SolCallback(4)= "bsRBPopper.SolOut"     '4-Right Back Popper
SolCallback(5)= "SolRampDown"       '5-Ramp Down
'SolCallback(6)= ""             '6-N.U.
SolCallback(7)= "vpmSolSound SoundFX(""fx_knocker"",DOFKnocker),"   '7-Knocker
SolCallback(8)= "bsRFPopper.Solout"     '8-Right Front Popper
'SolCallback(9)= ""             '9-Left Sling
'SolCallback(10)= ""            '10-Right Sling
'SolCallback(11)= ""            '11-Left Jet
'SolCallback(12)= ""            '12-Bottom Jet
'SolCallback(13)= ""            '13-Right Jet
SolCallback(14)= "SolPhoneFlasher"      '14-Flasher: Phone
'SolCallback(15)= ""            '15-N.U.
SolCallback(16)= "SolRampUp"        '16-Ramp Up
SolCallback(17)= "SolBackFlashers"      '17-Flasher: Back (x2)
SolCallback(18)= "FL18.State="        '18-Flasher: AutoFire
SolCallback(19)= "SolLowerLeftFlasher"    '19-Flasher: Lower Left
SolCallback(20)= "SolSpinnerFlasher"    '20-Flasher: Spinner
SolCallback(21)= "SolLowerRightFlasher"   '21-Flasher: Lower Right
SolCallback(22)= "SolBank"          '22-Motor: 3-Bank
'SolCallback(23)= ""            '23-Motor: Left Slot B
'SolCallback(24)= ""            '24-Motor: Left Slot A
'SolCallback(25)= ""            '25-Motor: Center Slot B
'SolCallback(26)= ""            '26-Motor: Center Slot A
'SolCallback(27)= ""            '27-Motor: Right Slot B
'SolCallback(28)= ""            '28-Motor: Right Slot A
'
SolCallback(36)= "SolPost"          '36-Up-Down Post

'******************************************
'       FLIPPERS
'******************************************

SolCallback(sLRFlipper) = "SolRFlipper"
SolCallback(sLLFlipper) = "SolLFlipper"

Sub SolLFlipper(Enabled)
  If Enabled Then
    PlaySoundAt SoundFX("fx_FlipperUp",DOFFlippers), LeftFlipper
    LeftFlipper.RotateToEnd
  Else
        PlaySoundAt SoundFX("fx_FlipperDown",DOFFlippers), LeftFlipper
        LeftFlipper.RotateToStart
     End If
 End Sub

Sub SolRFlipper(Enabled)
    If Enabled Then
        PlaySoundAt SoundFX("fx_FlipperUp",DOFFlippers), RightFlipper
        RightFlipper.RotateToEnd
    Else
        PlaySoundAt SoundFX("fx_FlipperDown",DOFFlippers), RightFlipper
        RightFlipper.RotateToStart
     End If
 End Sub

'************************************************************************
'            BALL TROUGH
'************************************************************************
Sub SolTrough(Enabled)
  If Enabled Then
    bsTrough.ExitSol_On
    If BsTrough.Balls Then vpmTimer.PulseSw 31
  End If
End Sub

'************************************************************************
'            AUTOPLUNGER
'************************************************************************
Sub SolAutoPlungerIM(Enabled)
  If Enabled Then
    PlungerIM.AutoFire
  End If
End Sub

'************************************************************************
'           3-BANK TARGETS
'************************************************************************
dim tbrake,tshake,tdir
Const targvel= 0.5      'the up-down speed

Sub TargetsInit
  'starts the bank in up position
    TargetBank.z=0
  sw66.Isdropped = 0:sw67.Isdropped = 0:sw68.Isdropped = 0:sw6X.IsDropped = 0
  Controller.Switch(11) = 0
  Controller.Switch(73) = 1
  tdir=-1:tbrake=0:tshake=0
End Sub

Dim BankDisableTime: BankDisableTime=0

Sub SolBank(Enabled)
  If Enabled then
    PlayLoopSoundAtVol SoundFX("fx_mine_motor", DOFGear), TargetBank, 0.5
  else
    StopSound "fx_mine_motor"
    BankDisableTime = 0
  End If
  BankMove.Enabled=Enabled
End Sub

Sub BankMove_timer()
  ' Are we waiting for the ROM to disable the movement?  If yes, exit
  if BankDisableTime > 0 and BankDisableTime > Timer Then Exit Sub
  TargetBank.z=TargetBank.z+targvel*tdir
  If TargetBank.z<-60 AND tdir=-1 then
    Controller.Switch(11) = 1:Controller.Switch(73) = 0:TargetBank.z=-60
    sw66.Isdropped = 1:sw67.Isdropped = 1:sw68.Isdropped = 1:sw6X.IsDropped = 1:tdir=1
    BankDisableTime = Timer + 1
  End If
  If TargetBank.z>0 AND tdir=1 then
    Controller.Switch(73) = 1:Controller.Switch(11) = 0:TargetBank.z=0
    sw66.Isdropped = 0:sw67.Isdropped = 0:sw68.Isdropped = 0:sw6X.IsDropped = 0:tdir=-1
    BankDisableTime = Timer + 1
  End if
End Sub

sub bankshake_timer()
  if tbrake>0 then
    tshake=tshake+30
    tbrake=tbrake-0.08
    TargetBank.transy =((sin(tshake)))*tbrake
  else
    me.Enabled=0:tbrake=0:tshake=0
  End if
End Sub

'************************************************************************
'           UP-DOWN RAMP
'************************************************************************
Dim rdir
Const rampvel= 5      'the up-down speed

Sub RampInit
  'starts the ramp in up position
  UpDownRamp.rotX=0:UpDownRamp1.rotX=0
  UpRamp.collidable=0
  Controller.Switch(74) = 1
End Sub

Sub SolRampDown(Enabled)
  If Enabled then
    PlaySoundAt SoundFX("Taxi Ramp Down",DOFContactors), UpDownRamp
    rdir=-1:RampMove.Enabled=1:UpRamp.collidable=1
  End If
End Sub

Sub SolRampUp(Enabled)
  If Enabled then
    PlaySoundAt SoundFX("Taxi Ramp Up",DOFContactors), UpDownRamp
    rdir=1:RampMove.Enabled=1:UpRamp.collidable=0
  End If
End Sub

Sub RampMove_timer()
  UpDownRamp.rotX=UpDownRamp.rotX+rampvel*rdir
  UpDownRamp1.rotX=UpDownRamp.rotX
  If UpDownRamp.rotX<-15 and rdir=-1 then Controller.Switch(74) = 0:Me.Enabled=0:UpDownRamp.rotX=-15:UpDownRamp1.rotX=-15
  If UpDownRamp.rotX>0 and rdir=1 then Controller.Switch(74) = 1:Me.Enabled=0:UpDownRamp.rotX=0:UpDownRamp1.rotX=0
End Sub

'************************************************************************
'           UP-DOWN POST
'************************************************************************

Sub diverter_Init
  diverter.IsDropped = 1              'up-down post
End Sub

Sub SolPost(Enabled)
  if Enabled then
    diverter.isdropped=0:PlaySoundAt SoundFX("fx_postup",DOFContactors),diverterP:diverterP.transZ = 70
  else
    diverter.isdropped=1:PlaySoundAt SoundFX("fx_postdown",DOFContactors),diverterP:diverterP.transZ = 0
  End if
End Sub

'************************************************************************
'           SLOT REEL
'************************************************************************
dim MechReelL, MechReelM, MechReelR, ReelSnd, Reel1Speed, Reel2speed, Reel3Speed
Reel1Speed=0:Reel2Speed=0:Reel3Speed=0:ReelSnd = 0

Sub UpdateReelSound
  If Reel1Speed + Reel2Speed + Reel3Speed > 0 Then
    If ReelSnd=0 Then PlayLoopSoundAtVol SoundFX("Reel Motor",DOFGear), Reel1, 0.25
    ReelSnd=Timer
  Else
    if ReelSnd > 0 and Timer > ReelSnd + .5 then StopSound "Reel Motor":ReelSnd=0
  End If
End Sub

Sub ReelLPos(aCurrPos, aSpeed, aLastPos)
  Reel1Speed = abs(aSpeed)
  Reel1.RotX = (-aCurrPos) + 275
End Sub

Sub ReelMPos(aCurrPos, aSpeed, aLastPos)
  Reel2Speed = abs(aSpeed)
  Reel2.RotX = (-aCurrPos) + 275
End Sub

Sub ReelRPos(aCurrPos, aSpeed, aLastPos)
  Reel3Speed = abs(aSpeed)
  Reel3.RotX = (-aCurrPos) + 275
End Sub

Sub SlotInit
  Set MechReelL = New cvpmMyMech : With mechReelL
    .Sol1 = 23 : .Sol2 = 24
    .MType = vpmMechLinear + vpmMechCircle + vpmMechStepSol + vpmMechFast
    .Length = 200 : .Steps = 360
    .AddSw 12, 0, 8
    .CallBack = GetRef("ReelLPos")
    .Start
  End With

  Set MechReelM = New cvpmMyMech : With mechReelM
    .Sol1 = 25 : .Sol2 = 26
    .MType = vpmMechLinear + vpmMechCircle + vpmMechStepSol + vpmMechFast
    .Length = 200 : .Steps = 360
    .AddSw 25, 0, 8
    .CallBack = GetRef("ReelMPos")
    .Start
  End With

  Set MechReelR = New cvpmMyMech : With mechReelR
    .Sol1 = 27 : .Sol2 = 28
    .MType = vpmMechLinear + vpmMechCircle + vpmMechStepSol + vpmMechFast
    .Length = 200 : .Steps = 360
    .AddSw 48, 0, 8
    .CallBack = GetRef("ReelRPos")
    .Start
  End With
End Sub

'*****************************************
'           Switches
'*****************************************

'rollover
Sub sw16_hit:Controller.switch(16) = 1: PlaySoundAt "fx_sensor",ActiveBall:End Sub
Sub sw16_unhit:Controller.switch(16) = 0: End Sub
Sub sw17_hit:Controller.switch(17) = 1: PlaySoundAt "fx_sensor",ActiveBall:End Sub
Sub sw17_unhit:Controller.switch(17) = 0: End Sub
Sub sw18_hit:Controller.switch(18) = 1: PlaySoundAt "fx_sensor",ActiveBall:End Sub
Sub sw18_unhit:Controller.switch(18) = 0: End Sub
Sub sw26_hit:Controller.switch(26) = 1: PlaySoundAt "fx_sensor",ActiveBall:End Sub
Sub sw26_unhit:Controller.switch(26) = 0: End Sub
Sub sw27_hit:Controller.switch(27) = 1: PlaySoundAt "fx_sensor",ActiveBall:End Sub
Sub sw27_unhit:Controller.switch(27) = 0: End Sub
Sub sw28_hit:Controller.switch(28) = 1: PlaySoundAt "fx_sensor",ActiveBall:End Sub
Sub sw28_unhit:Controller.switch(28) = 0: End Sub

'ramp optos
Sub sw36_hit:vpmtimer.pulsesw 36: End Sub
Sub sw37_hit:vpmtimer.pulsesw 37: End Sub

'subway optos
Dim SubwaySnd:SubwaySnd=0
Sub sw41_hit:controller.Switch(41)=1:If SubwaySnd=0 Then SubwaySnd=1:PlaySoundAt "fx_subway",ActiveBall:End If: End Sub
Sub sw41_unhit:controller.Switch(41)=0:End Sub
Sub sw42_hit:controller.Switch(42)=1:If SubwaySnd=0 Then SubwaySnd=1:PlaySoundAt "fx_subway",ActiveBall:End If: End Sub
Sub sw42_unhit:controller.Switch(42)=0:End Sub
Sub sw47_hit:controller.Switch(47)=1:If SubwaySnd=0 Then SubwaySnd=1:PlaySoundAt "fx_subway",ActiveBall:End If: End Sub
Sub sw47_unhit:controller.Switch(47)=0:End Sub
Sub SubwayStop_Hit():StopSound "fx_subway":SubwaySnd=0:End Sub

'4-bank targets
Sub sw52_hit:vpmtimer.pulsesw 52:PlaySoundAt SoundFX("fx_target",DOFTargets), ActiveBall:End Sub
Sub sw53_hit:vpmtimer.pulsesw 53:PlaySoundAt SoundFX("fx_target",DOFTargets), ActiveBall:End Sub
Sub sw54_hit:vpmtimer.pulsesw 54:PlaySoundAt SoundFX("fx_target",DOFTargets), ActiveBall:End Sub
Sub sw55_hit:vpmtimer.pulsesw 55:PlaySoundAt SoundFX("fx_target",DOFTargets), ActiveBall:End Sub

'Mystery target
Sub sw56_hit:vpmtimer.pulsesw 56:PlaySoundAt SoundFX("fx_target",DOFTargets), ActiveBall:End Sub

'Red target
Sub sw58_hit:vpmtimer.pulsesw 58: PlaySoundAt SoundFX("fx_target",DOFTargets), ActiveBall:End Sub

'Slingshots
Dim LStep, RStep

Sub sw61_slingshot
  vpmTimer.PulseSw 61
  PlaySoundAt SoundFX("fx_slingshot",DOFContactors), sw26
  LSling.Visible = 0: LSling2.Visible = 1: sling1.TransZ = -20: LStep = 0
  Me.TimerInterval = 50: Me.TimerEnabled = 1
End Sub

Sub sw62_slingshot
  vpmTimer.PulseSw 62
  PlaySoundAt SoundFX("fx_slingshot",DOFContactors), sw17
  RSling.Visible = 0: RSling2.Visible = 1: sling2.TransZ = -20: RStep = 0
  Me.TimerInterval = 50: Me.TimerEnabled = 1
End Sub

Sub sw61_Timer
  Select Case LStep
    Case 3:LSLing2.Visible = 0:LSLing1.Visible = 1:sling1.TransZ = -10
    Case 4:LSLing1.Visible = 0:LSLing.Visible = 1:sling1.TransZ = 0:Me.TimerEnabled = 0
  End Select
  LStep = LStep + 1
End Sub

Sub sw62_Timer
  Select Case RStep
    Case 3:RSLing2.Visible = 0:RSLing1.Visible = 1:sling2.TransZ = -10
    Case 4:RSLing1.Visible = 0:RSLing.Visible = 1:sling2.TransZ = 0:Me.TimerEnabled = 0
  End Select
  RStep = RStep + 1
End Sub

'Bumpers
Sub sw63_hit:vpmtimer.pulsesw 63:PlaySoundAt SoundFX("LeftJet",DOFContactors), ActiveBall:End Sub
Sub sw64_hit:vpmtimer.pulsesw 64:PlaySoundAt SoundFX("BottomJet",DOFContactors), ActiveBall:End Sub
Sub sw65_hit:vpmtimer.pulsesw 65:PlaySoundAt SoundFX("RightJet",DOFContactors), ActiveBall:End Sub

'3-bank drop targets
Sub sw66_hit:vpmtimer.pulsesw 66:bankshake.enabled=1:tbrake=5:PlaySoundAt SoundFX("fx_target",DOFTargets), ActiveBall:End Sub
Sub sw67_hit:vpmtimer.pulsesw 67:bankshake.enabled=1:tbrake=5:PlaySoundAt SoundFX("fx_target",DOFTargets), ActiveBall:End Sub
Sub sw68_hit:vpmtimer.pulsesw 68:bankshake.enabled=1:tbrake=5:PlaySoundAt SoundFX("fx_target",DOFTargets), ActiveBall:End Sub

'2-bank targets
Sub sw71_hit:vpmtimer.pulsesw 71:PlaySoundAt SoundFX("fx_target",DOFTargets), ActiveBall:End Sub
Sub sw72_hit:vpmtimer.pulsesw 72:PlaySoundAt SoundFX("fx_target",DOFTargets), ActiveBall:End Sub

'Scoop Roof
Sub sw75_hit:Controller.switch(75) = 1: PlaySoundAt "fx_sensor",ActiveBall: sw75p.RotZ=-35: Me.TimerInterval=100: Me.TimerEnabled=1: End Sub
Sub sw75_timer():Me.TimerEnabled=0:sw75p.RotZ=-70:End Sub
Sub sw75_unhit:Controller.switch(75) = 0: End Sub
Sub sw76_hit:Controller.switch(76) = 1: PlaySoundAt "fx_sensor",ActiveBall: sw76p.RotZ=25: Me.TimerInterval=100: Me.TimerEnabled=1: End Sub
Sub sw76_timer():Me.TimerEnabled=0:sw76p.RotZ=-10:End Sub
Sub sw76_unhit:Controller.switch(76) = 0: End Sub
Sub sw77_hit:Controller.switch(77) = 1: PlaySoundAt "fx_sensor",ActiveBall: sw77p.RotZ=35: Me.TimerInterval=100: Me.TimerEnabled=1: End Sub
Sub sw77_timer():Me.TimerEnabled=0:sw77p.RotZ=0:End Sub
Sub sw77_unhit:Controller.switch(77) = 0: End Sub

'Black target
Sub sw78_hit:vpmtimer.pulsesw 78: PlaySoundAt SoundFX("fx_target",DOFTargets), ActiveBall:End Sub

'Spinner
Sub sw115_spin:vpmtimer.pulsesw 115: PlaySoundAt "fx_spinner", sw115:End Sub

'******************************************************
'         Flupper Flashers
'******************************************************

Dim FlashLevel1, Flashlevel2, FlashLevel3, FlashLevel4, FlashLevel5, FlashLevel6

FlasherLight1.IntensityScale = 0
Flasherlight2.IntensityScale = 0
Flasherlight3.IntensityScale = 0
Flasherlight4.IntensityScale = 0
Flasherlight5.IntensityScale = 0
Flasherlight6.IntensityScale = 0

Sub SolBackFlashers(enabled)
  If Enabled then
    Dim Intensity : If Light15.state = 0 then Intensity = FlasherIntensityGIOff else Intensity = FlasherIntensityGIOn
    FlashLevel1 = Intensity
    FlasherFlash1.TimerEnabled=0
    Roulette.BlendDisableLightingFromBelow = 0
    FlasherFlash1.TimerEnabled=1
  end if
End Sub

Sub SolLowerLeftFlasher(enabled)
  if Enabled then
    Dim Intensity : If Light15.state = 0 then Intensity = FlasherIntensityGIOff else Intensity = FlasherIntensityGIOn
    FlashLevel3 = Intensity
    FlasherFlash3_Timer
  end if
End Sub

Sub SolLowerRightFlasher(enabled)
  if Enabled then
    Dim Intensity : If Light15.state = 0 then Intensity = FlasherIntensityGIOff else Intensity = FlasherIntensityGIOn
    FlashLevel4 = Intensity
    FlasherFlash4_Timer
  end if
End Sub

Sub SolSpinnerFlasher(enabled)
  if Enabled then
    Dim Intensity : If Light15.state = 0 then Intensity = FlasherIntensityGIOff else Intensity = FlasherIntensityGIOn
    FlashLevel5 = Intensity
    FlasherFlash5_Timer
  end if
End Sub

Sub SolPhoneFlasher(enabled)
  If enabled Then
    FlashLevel6=1
    FlasherFlash6_Timer
  End If
End Sub

Sub FlasherFlash1_Timer()
  dim flashx3, matdim
  FlasherFlash1.visible = 1
  FlasherFlash2.visible = 1
  FlasherFlash1a.visible = 1:FlasherFlash2a.visible = 1
  FlasherLit1.visible = 1
  FlasherLit2.visible = 1
  flashx3 = FlashLevel1 * FlashLevel1 * FlashLevel1
  Flasherflash1.opacity = 1000 * flashx3
  Flasherflash2.opacity = 1000 * flashx3
  Flasherflash1a.opacity = 500 * flashx3:Flasherflash2a.opacity = 500 * flashx3
  FlasherLit1.BlendDisableLighting = 10 * flashx3
  FlasherLit2.BlendDisableLighting = 10 * flashx3
  If Flasherbase1.BlendDisableLighting > 0.4 Then Flasherbase1.BlendDisableLighting =  flashx3
  If Flasherbase2.BlendDisableLighting > 0.4 Then Flasherbase2.BlendDisableLighting =  flashx3
  FlasherLight1.IntensityScale = flashx3
  FlasherLight2.IntensityScale = flashx3
  matdim = Round(10 * FlashLevel1)
  FlasherLit1.material = "domelit" & matdim
  FlasherLit2.material = "domelit" & matdim
  FlashLevel1 = FlashLevel1 * 0.85 - 0.01
  If FlashLevel1 < 0.8 Then Roulette.BlendDisableLightingFromBelow = 1
  If FlashLevel1 < 0.15 Then
    FlasherLit1.visible = 0
    FlasherLit2.visible = 0
  Else
    FlasherLit1.visible = 1
    FlasherLit2.visible = 1
  end If
  If FlashLevel1 < 0 Then
    FlasherFlash1.TimerEnabled = False
    FlasherFlash1.visible = 0
    FlasherFlash2.visible = 0
    FlasherFlash1a.visible = 0:FlasherFlash2a.visible = 0
    Roulette.BlendDisableLightingFromBelow=1
  End If
End Sub

Sub FlasherFlash3_Timer()
  dim flashx3, matdim
  If not Flasherflash3.TimerEnabled Then
    Flasherflash3.TimerEnabled = True
    Flasherflash3.visible = 1
    Flasherlit3.visible = 1
  End If
  flashx3 = FlashLevel3 * FlashLevel3 * FlashLevel3
  Flasherflash3.opacity = 8000 * flashx3
  Flasherlit3.BlendDisableLighting = 10 * flashx3
  Flasherbase3.BlendDisableLighting =  flashx3
  Flasherlight3.IntensityScale = flashx3
  matdim = Round(10 * FlashLevel3)
  Flasherlit3.material = "domelit" & matdim
  FlashLevel3 = FlashLevel3 * 0.85 - 0.01
  If FlashLevel3 < 0.15 Then
    Flasherlit3.visible = 0
  Else
    Flasherlit3.visible = 1
  end If
  If FlashLevel3 < 0 Then
    Flasherflash3.TimerEnabled = False
    Flasherflash3.visible = 0
  End If
End Sub

Sub FlasherFlash4_Timer()
  dim flashx3, matdim
  If not Flasherflash4.TimerEnabled Then
    Flasherflash4.TimerEnabled = True
    Flasherflash4.visible = 1
    Flasherlit4.visible = 1
  End If
  flashx3 = FlashLevel4 * FlashLevel4 * FlashLevel4
  Flasherflash4.opacity = 4000 * flashx3
  Flasherlit4.BlendDisableLighting = 10 * flashx3
  Flasherbase4.BlendDisableLighting =  flashx3
  Flasherlight4.IntensityScale = flashx3
  matdim = Round(10 * FlashLevel3)
  Flasherlit4.material = "domelit" & matdim
  FlashLevel4 = FlashLevel4 * 0.85 - 0.01
  If FlashLevel4 < 0.15 Then
    Flasherlit4.visible = 0
  Else
    Flasherlit4.visible = 1
  end If
  If FlashLevel4 < 0 Then
    Flasherflash4.TimerEnabled = False
    Flasherflash4.visible = 0
  End If
End Sub

Sub FlasherFlash5_Timer()
  dim flashx3, matdim
  If not Flasherflash5.TimerEnabled Then
    Flasherflash5.TimerEnabled = True
    Flasherflash5.visible = 1
    Flasherlit5.visible = 1
  End If
  flashx3 = FlashLevel5 * FlashLevel5 * FlashLevel5
  Flasherflash5.opacity = 8000 * flashx3
  Flasherlit5.BlendDisableLighting = 10 * flashx3
  Flasherbase5.BlendDisableLighting =  flashx3
  Flasherlight5.IntensityScale = flashx3
  matdim = Round(10 * FlashLevel3)
  Flasherlit5.material = "domelit" & matdim
  FlashLevel5 = FlashLevel5 * 0.85 - 0.01
  If FlashLevel5 < 0.15 Then
    Flasherlit5.visible = 0
  Else
    Flasherlit5.visible = 1
  end If
  If FlashLevel5 < 0 Then
    Flasherflash5.TimerEnabled = False
    Flasherflash5.visible = 0
  End If
End Sub

Sub FlasherFlash6_Timer()
  dim flashx3
  If not FlasherFlash6.TimerEnabled Then
    FlasherFlash6.TimerEnabled = True
    FlasherFlash6.visible = 1
  End If
  flashx3 = FlashLevel6 * FlashLevel6 * FlashLevel6
  FlasherFlash6.opacity = 6000 * flashx3
  Flasherlight6.IntensityScale = flashx3
  FlashLevel6 = FlashLevel6 * 0.85 - 0.01
  If FlashLevel6 < 0 Then
    FlasherFlash6.TimerEnabled = False
    FlasherFlash6.visible = 0
  End If
End Sub

'******************************************************
'         RealTime Updates
'******************************************************
Set MotorCallback = GetRef("RealTimeUpdates")

Sub RealTimeUpdates()
  LeftFlipperP.RotZ = LeftFlipper.currentangle
  RightFlipperP.RotZ = RightFlipper.currentangle
  LeftFlipperSh.RotZ = LeftFlipper.currentangle
  RightFlipperSh.RotZ = RightFlipper.currentangle
  ScoopGate.RotX = Gate4.currentangle/2
  sw115p.RotX = -15*ABS(cos (sw115.currentangle*3.14/360))
  RollingSoundUpdate
  BallShadowUpdate
  SpotLightsUpdate
  UpdateReelSound
End Sub

' *********************************************************************
'         Supporting Ball & Sound Functions
' *********************************************************************

Function Vol(ball) ' Calculates the Volume of the sound based on the ball speed
    Vol = Csng(BallVel(ball) ^2 / 2000)
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

function AudioFade(ball)
  Dim tmp
    tmp = ball.y * 2 / Table1.height-1
    If tmp > 0 Then
    AudioFade = Csng(tmp ^10)
    Else
        AudioFade = Csng(-((- tmp) ^10) )
    End If
End Function

Function Pitch(ball) ' Calculates the pitch of the sound based on the ball speed
    Pitch = BallVel(ball) * 20
End Function

Function BallVel(ball) 'Calculates the ball speed
    BallVel = INT(SQR((ball.VelX ^2) + (ball.VelY ^2) ) )
End Function

Function RndNum(min,max)
 RndNum = Int(Rnd()*(max-min+1))+min     ' Sets a random number between min and max
End Function

'Set position as table object (Use object or light but NOT wall) and Vol to 1
Sub PlaySoundAt(sound, tableobj)
  PlaySound sound, 1, 1, Pan(tableobj), 0,0,0, 1, AudioFade(tableobj)
End Sub

'Set all as per ball position & speed.
Sub PlaySoundAtBall(sound)
  PlaySound sound, 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 0, 1, AudioFade(ActiveBall)
End Sub

'Set position as table object and Vol manually.
Sub PlaySoundAtVol(sound, tableobj, Vol)
  PlaySound sound, 1, Vol, Pan(tableobj), 0, 0, 0, 1, AudioFade(tableobj)
End Sub

Sub PlayLoopSoundAtVol(sound, tableobj, Vol)
  PlaySound sound, -1, Vol, Pan(tableobj), 0, 0, 1, 0, AudioFade(tableobj)
End Sub

'Set all as per ball position & speed, but Vol Multiplier may be used eg; PlaySoundAtBallVol "sound",3
Sub PlaySoundAtBallVol(sound, VolMult)
  PlaySound sound, 0, Vol(ActiveBall) * VolMult, Pan(ActiveBall), 0, Pitch(ActiveBall), 0, 1, AudioFade(ActiveBall)
End Sub

'Set position as bumperX and Vol manually.
Sub PlaySoundAtBumperVol(sound, tableobj, Vol)
  PlaySound sound, 1, Vol, Pan(tableobj), 0,0,1, 1, AudioFade(tableobj)
End Sub

Dim NextOrbitHit:NextOrbitHit = 0
Sub PlasticRampBumps_Hit(idx)
  if BallVel(ActiveBall) > .3 and Timer > NextOrbitHit then
    RandomBump 10, -20000
    ' Schedule the next possible sound time.  This prevents it from rapid-firing noises too much.
    ' Lowering these numbers allow more closely-spaced clunks.
    NextOrbitHit = Timer + .1 + (Rnd * .2)
  end if
End Sub

' Requires rampbump1 to 7 in Sound Manager
Sub RandomBump(voladj, freq)
  dim BumpSnd:BumpSnd= "rampbump" & CStr(Int(Rnd*7)+1)
  PlaySound BumpSnd, 0, Vol(ActiveBall)*voladj, Pan(ActiveBall), 0, freq, 0, 1, AudioFade(ActiveBall)
End Sub

' *********************************************************************
'             Other Sound FX
' *********************************************************************

Sub OnBallBallCollision(ball1, ball2, velocity)
  if ball1.z >= 0 then
    PlaySound("fx_collide"), 0, Csng(velocity) ^2 / 500, Pan(ball1), 0, Pitch(ball1), 0, 0, AudioFade(ball1)
  Else ' in playfield tunnel, muffled!
    PlaySound("fx_collide"), 0, Csng(velocity) ^2 / 1000, Pan(ball1), 0, -20000, 0, 0, AudioFade(ball1)
  end if
End Sub

Sub LeftFlipper_Collide(parm)
    RandomSoundFlipper
End Sub

Sub Rightflipper_Collide(parm)
    RandomSoundFlipper
End Sub

Sub Pins_Hit(idx)
  RandomSoundRubber()
End Sub

Sub Posts_Hit(idx)
  RandomSoundRubber()
End Sub

Sub Rubbers_Hit(idx)
  RandomSoundRubber()
End Sub

Sub Gates_Hit(idx)
  PlaySoundAtBallVol "gate4",1
End Sub

Sub Trigger001_Hit()
  PlaySoundAtBallVol "fx_BallDrop",.3
End Sub

Sub RandomSoundRubber()
  Select Case RndNum(1,3)
    Case 1 : PlaySoundAtBallVol "fx_rubber_hit_1",10
    Case 2 : PlaySoundAtBallVol "fx_rubber_hit_2",10
    Case 3 : PlaySoundAtBallVol "fx_rubber_hit_3",10
  End Select
End Sub

Sub RandomSoundFlipper()
  Select Case RndNum(1,3)
    Case 1 : PlaySoundAtBallVol "flip_hit_1", 20
    Case 2 : PlaySoundAtBallVol "flip_hit_2", 20
    Case 3 : PlaySoundAtBallVol "flip_hit_3", 20
  End Select
End Sub

Sub LeftHole_hit:PlaySoundAt "fx_hole3", ActiveBall:End Sub
Sub RightHole_hit:PlaySoundAt "fx_hole2", ActiveBall:End Sub
Sub CenterHole_hit:PlaySoundAt "fx_hole1", ActiveBall:End Sub
Sub Trigger1_hit:Me.TimerInterval=200:Me.TimerEnabled=1:End Sub
Sub Trigger1_timer:Me.TimerEnabled=0:PlaysoundAt "fx_BallDrop",Trigger1:End Sub
Sub Trigger2_hit:Me.TimerInterval=200:Me.TimerEnabled=1:End Sub
Sub Trigger2_timer:Me.TimerEnabled=0:PlaysoundAt "fx_BallDrop",Trigger2:End Sub
Sub Trigger3_hit:Me.TimerInterval=200:Me.TimerEnabled=1:End Sub
Sub Trigger3_timer:Me.TimerEnabled=0:PlaysoundAt "fx_BallDrop",Trigger3:End Sub
Sub Trigger4_hit:Me.TimerInterval=200:Me.TimerEnabled=1:End Sub
Sub Trigger4_timer:Me.TimerEnabled=0:PlaysoundAt "fx_BallDrop",Trigger4:End Sub

' *********************************************************************
'           ROLLING SOUND
' *********************************************************************
Const tnob = 4            ' total number of balls : 4 (trough)
ReDim rolling(tnob-1)
InitRolling
Sub InitRolling:Dim i:For i=0 to (tnob-1):rolling(i) = False:Next:End Sub

Sub RollingSoundUpdate()
    Dim BOT, b
    BOT = GetBalls
  ' stop the sound of deleted balls
  If UBound(BOT)<(tnob - 1) Then
    For b = (UBound(BOT) + 1) to (tnob-1)
      rolling(b) = False
      StopSound("fx_ballrolling" & b+1)
    Next
  End If
  ' exit the Sub if no balls on the table
    If UBound(BOT) = -1 Then Exit Sub
       ' play the rolling sound for each ball
    For b = 0 to UBound(BOT)
        If BallVel(BOT(b) ) > 1 Then
      rolling(b) = True
      if BOT(b).z < 30 Then ' Ball on playfield
            PlaySound("fx_ballrolling" & b+1), -1, Vol(BOT(b) )/4, Pan(BOT(b) ), 0, Pitch(BOT(b) ), 1, 0, AudioFade(BOT(b) )
      Else ' Ball on raised ramp
            PlaySound("fx_ballrolling" & b+1), -1, Vol(BOT(b) )/10, Pan(BOT(b) ), 0, Pitch(BOT(b) )+50000, 1, 0, AudioFade(BOT(b) )
        End If
    Else
            If rolling(b) = True Then
                StopSound("fx_ballrolling" & b+1)
                rolling(b) = False
            End If
        End If
    Next
End Sub

' *********************************************************************
'           BALL SHADOW
' *********************************************************************
ReDim BallShadow(tnob-1)
InitBallShadow

Sub InitBallShadow
  Dim i: For i=0 to tnob-1
    ExecuteGlobal "Set BallShadow(" & i & ")=BallShadow" & (i+1) & ":"
  Next
End Sub

Sub BallShadowUpdate()
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
    If BOT(b).X < Table1.Width/2 Then
      BallShadow(b).X = ((BOT(b).X) - (Ballsize/6) + ((BOT(b).X - (Table1.Width/2))/10)) + 10
    Else
      BallShadow(b).X = ((BOT(b).X) + (Ballsize/6) + ((BOT(b).X - (Table1.Width/2))/10)) - 10
    End If
      ballShadow(b).Y = BOT(b).Y + 20
    If BOT(b).Z > 20 Then
      BallShadow(b).visible = 1
    Else
      BallShadow(b).visible = 0
    End If
  Next
End Sub


' *****  cvpmMech replacement, all because the "Fast" parameter is not set on timer update!!   We can't get smooth reels without this.   Until this is addressed in Core.vbs, replacing here.

Class cvpmMyMech
  Public Sol1, Sol2, MType, Length, Steps, Acc, Ret
  Private mMechNo, mNextSw, mSw(), mLastPos, mLastSpeed, mCallback

  Private Sub Class_Initialize
    ReDim mSw(10)
    gNextMechNo = gNextMechNo + 1 : mMechNo = gNextMechNo : mNextSw = 0 : mLastPos = 0 : mLastSpeed = 0
    MType = 0 : Length = 0 : Steps = 0 : Acc = 0 : Ret = 0 : vpmTimer.addResetObj Me
  End Sub

  Public Sub AddSw(aSwNo, aStart, aEnd)
    mSw(mNextSw) = Array(aSwNo, aStart, aEnd, 0)
    mNextSw = mNextSw + 1
  End Sub

  Public Sub AddPulseSwNew(aSwNo, aInterval, aStart, aEnd)
    If Controller.Version >= "01200000" Then
      mSw(mNextSw) = Array(aSwNo, aStart, aEnd, aInterval)
    Else
      mSw(mNextSw) = Array(aSwNo, -aInterval, aEnd - aStart + 1, 0)
    End If
    mNextSw = mNextSw + 1
  End Sub

  Public Sub Start
    Dim sw, ii
    With Controller
      .Mech(1) = Sol1 : .Mech(2) = Sol2 : .Mech(3) = Length
      .Mech(4) = Steps : .Mech(5) = MType : .Mech(6) = Acc : .Mech(7) = Ret
      ii = 10
      For Each sw In mSw
        If IsArray(sw) Then
          .Mech(ii) = sw(0) : .Mech(ii+1) = sw(1)
          .Mech(ii+2) = sw(2) : .Mech(ii+3) = sw(3)
          ii = ii + 10
        End If
      Next
      .Mech(0) = mMechNo
    End With
    If IsObject(mCallback) Then mCallBack 0, 0, 0 : mLastPos = 0 : vpmTimer.EnableUpdate Me, True, True  ' <------- All for this.
  End Sub

  Public Property Get Position : Position = Controller.GetMech(mMechNo) : End Property
  Public Property Get Speed    : Speed = Controller.GetMech(-mMechNo)   : End Property
  Public Property Let Callback(aCallBack) : Set mCallback = aCallBack : End Property

  Public Sub Update
    Dim currPos, speed
    currPos = Controller.GetMech(mMechNo)
    speed = Controller.GetMech(-mMechNo)
    If currPos < 0 Or (mLastPos = currPos And mLastSpeed = speed) Then Exit Sub
    mCallBack currPos, speed, mLastPos : mLastPos = currPos : mLastSpeed = speed
  End Sub

  Public Sub Reset : Start : End Sub

End Class

' 85vell for PM5
' Converted to VPX by freneticamnesic.
' Mod by Thalamus, 32a and Hauntfreaks
' oqqsan : Nfozzy physics,  new inserts, extra play sounds

' oqqsan : scorecard
' Press 'R' button during play to magnify instructions !


Option Explicit
Randomize

' Thalamus 2018-07-24
' Added/Updated "Positional Sound Playback Functions" and "Supporting Ball & Sound Functions"
' Improved directional sounds.
' Added InitVpmFFlipsSAM

' Thalamus 2018-11-01 : Improved directional sounds

' Options
' Volume devided by - lower gets higher sound

Const VolDiv = 1500    ' Lower number, louder ballrolling/collition sound
Const VolCol = 10      ' Ball collition divider ( voldiv/volcol )
Const VolPla = .1      ' Plastic ramp rolling volume

' The rest of the values are multipliers
'
'  .5 = lower volume
' 1.5 = higher volume

Const VolBump   = 1    ' Bumpers volume.
Const VolAmp    = 20   ' Ampsuit volume.
Const VolRol    = 1    ' Rollovers volume.
Const VolRel    = 1    ' Relay volume.
Const VolGates  = 1.5  ' Gates volume.
Const VolMetal  = 2    ' Metals volume.
Const VolRH     = 2    ' Rubber hits volume.
Const VolPo     = 2    ' Rubber posts volume.
Const VolPi     = 2    ' Rubber pins volume.
Const VolTarg   = 2    ' Targets volume.
Const VolSpin   = 1.5  ' Spinners volume.
Const VolFlip   = 1    ' Flipper volume.
Const VolMot    = 2    ' Motor volume.
Const VolSli    = 1    ' Slingshot volume.

Const FlipSolNudge = 1      ' Make flipper solenoid to alter active ball behavior


On Error Resume Next
ExecuteGlobal GetTextFile("controller.vbs")
If Err Then MsgBox "You need the controller.vbs in order to run this table, available in the vp10 package"
On Error Goto 0

Const BallMass = 1
Const Ballsize = 50
Const cGameName="avr_200",UseSolenoids=1,UseLamps=0,UseGI=0,SSolenoidOn="SolOn",SSolenoidOff="SolOff", SCoin="coin", UseVPMModSol=1

LoadVPM "01560000","sam.vbs",3.43
Dim DesktopMode: DesktopMode = Table1.ShowDT

If DesktopMode = True Then 'Show Desktop components
Ramp16.visible=1
Ramp15.visible=1
Primitive13.visible=1
Else
Ramp16.visible=0
Ramp15.visible=0
Primitive13.visible=0
End if


'*************************************************************
'Solenoid Call backs
'**********************************************************************************************************
SolCallback(1) = "solTrough"            ' Through up-kicker
SolCallback(2) = "solAutofire"          ' Auto Launch
'SolCallback(3) = "solampmagnet"        ' Amp Suit Magnet
'SolCallback(4) = "N/A"                 ' Not Used
SolCallback(5) = "sol3bankmotor"        ' Up-Down 3-Bank relay
SolCallback(6) = "LinkLockupUp"         ' Link Lockup Up
SolCallback(7) = "LinkLockUpLatch"      ' Link Lockup Latch
'SolCallback(8) = "solShakerMotor"      ' Shaker Motor (Optional)
SolModCallback(9)  = "SetLampMod 109,"     ' Top Left Pop Bumber
SolModCallback(10) = "SetLampMod 110,"     ' Top Right Pop Bumber
SolModCallback(11) = "SetLampMod 111,"     ' Bottom Pop Bumper
'SolCallback(12) = "N/A"                ' Not Used

SolCallback(13) = "ampsuitmotorRelay"

'SolCallback(14) = "N/A"                ' Not Used
SolCallback(15) = "SolLFlipper"         ' Left Flipper
SolCallback(16) = "SolRFlipper"         ' Right Flipper
'SolCallback(17) = "SolLSlingShot"      ' Left Slingshot
'SolCallback(18) = "SolRSlingShot"      ' Right Slingshot
'SolCallback(19) = "SolampSuitMotor"    ' Amp Suit Motor

SolModCallback(20) = "SetLampMod 120,"  ' Left ampsuit flasher | Doc as Link Flash
SolModCallback(21) = "SetLampMod 121,"  ' Link Flasher | Doc as Back Panel Flash (x2)
SolModCallback(22) = "SetLampMod 122,"  ' In front of Jake - fantacy placement | Doc as Link (L) Flash
SolModCallback(23) = "SetLampMod 123,"  ' Rigth ampsuit flasher | Doc as Link (R) Flash
SolModCallback(24) = "SetLampMod 124,"  ' Coin ? | Doc as Optional ( eg Coin Meter )
SolModCallback(25) = "SetLampMod 125,"  ' Bumper flasher

SolModCallback(26) = "SetLampMod 126,"  ' Ampsuit mb flasher
'SolCallback(27)   = "N/A"              ' Not Used
SolModCallback(28) = "SetLampMod 128,"  ' Amp Suit Flash | Doc as Amp Suit Flash (x2)
'SolModCallback(29) = "SetLampMod 129,"  ' upper left blue dome | Doc as Left Side Flash
'SolModCallback(30) = "SetLampMod 130,"  ' lower left blue dome | Doc as Left Side Blue (x2)
'SolModCallback(31) = "SetLampMod 131,"  ' upper right blue dome | Doc as Right Side Blue (x2)
'SolModCallback(32) = "SetLampMod 132,"  ' lower left blue dome | Doc as Right Side Flash

SolModCallback(29) = "Flash1"  ' upper left blue dome | Doc as Left Side Flash
SolModCallback(30) = "Flash3"  ' lower left blue dome | Doc as Left Side Blue (x2)
SolModCallback(31) = "Flash4"  ' upper right blue dome | Doc as Right Side Blue (x2)
SolModCallback(32) = "Flash2"  ' lower left blue dome | Doc as Right Side Flash

'Dim FlashLevel129, FlashLevel130, FlashLevel131, FlashLevel132
'FlasherLight129.IntensityScale = 0
'F129.IntensityScale = 0
'FlasherLight130.IntensityScale = 0
'Flasherlight130a.IntensityScale = 0
'Flasherlight131.IntensityScale = 0
'Flasherlight131a.IntensityScale = 0
'Flasherlight132.IntensityScale = 0
'F132.IntensityScale = 0

Sub SolLFlipper(Enabled)
     If Enabled Then
         PlaySoundAtVol SoundFX("Flipper_L" & int(rnd(1)*11)+1,DOFContactors), LeftFlipper, VolFlip: LF.fire 'LeftFlipper.RotateToEnd
     Else
         PlaySoundAtVol SoundFX("Flipper_Left_Down_" & int(rnd(1)*7)+1,DOFContactors), LeftFlipper, VolFlip:LeftFlipper.RotateToStart
     End If
  End Sub

Sub SolRFlipper(Enabled)
     If Enabled Then
         PlaySoundAtVol SoundFX("Flipper_R" & int(rnd(1)*11)+1 ,DOFContactors),RightFlipper, VolFlip:RF.fire 'RightFlipper.RotateToEnd
     Else
         PlaySoundAtVol SoundFX("Flipper_Left_Down_" & int(rnd(1)*8)+1, DOFContactors),RightFlipper, VolFlip:RightFlipper.RotateToStart
     End If
End Sub
'**********************************************************************************************************

'Solenoid Controlled toys
'**********************************************************************************************************

'Stern-Sega GI
set GICallback = GetRef("UpdateGI")

Sub UpdateGI(no, Enabled)
  If Enabled Then
    dim xx, xxx
    For each xx in GI:xx.State = 1: Next   'Light Objects
    For each xxx in GI2:xxx.visible = 1: Next 'Flasher Objects
        PlaySoundAtVol "fx_relay", sw35, VolRel
  Else
    For each xx in GI:xx.State = 0: Next
    For each xxx in GI2:xxx.visible = 0: Next
        PlaySoundAtVol "fx_relay", sw35, VolRel
  End If
End Sub

Sub ampsuitmotorRelay(Enabled)
  If Enabled Then
    ' Controller.Switch(57) = 1:Controller.Switch(58) = 0
    ampf.rotatetoend
    ' debug.print "ampsuitmotorrelay enabled"
  PlaySoundAtVol "Diverter", ampf, VolAmp
  Else
    ' Controller.Switch(57) = 0:Controller.Switch(58) = 1
    ' debug.print "ampsuitmotorrelay disabled"
    ampf.rotatetostart
  PlaySoundAtVol "Diverter", ampf, VolAmp
  End If
End Sub

Sub LinkLockupUp(Enabled)
  If Enabled Then
    lockPin1.Isdropped=0:lockPin2.Isdropped=0
    DropJake
  PlaySoundAt "fx_Flipperup", sw17
  End If
End Sub

Sub LinkLockUpLatch(Enabled)
  If Enabled Then
    lockPin1.Isdropped=1:lockPin2.Isdropped=1
    RiseJake':JakeTop.RotX=160:
  PlaySoundAt "fx_Flipperup", sw17
  End If
End Sub

Sub solTrough(Enabled)
  If Enabled Then
    bsTrough.ExitSol_On
    vpmTimer.PulseSw 22
  End If
End Sub

Sub solAutofire(Enabled)
  If Enabled Then
    PlungerIM.AutoFire
    PlaySoundAt "Autoplunger", Plunger
  End If
 End Sub




'Primitive animation code

'Sub RLS_Timer()
'  sw52p.RotZ = -(sw52.currentangle)
'  ampprim.objrotx = ampf.currentangle
'  LFLogo.RotY = LeftFlipper.CurrentAngle
'  RFlogo.RotY = RightFlipper.CurrentAngle
'End Sub

Sub RLS_Timer()
  RampGate1.RotZ = -(sw52.currentangle)
' RampGate2.RotZ = -(Gate3.currentangle)
  ampprim.objrotx = ampf.currentangle
  LFLogo.RotY = LeftFlipper.CurrentAngle
  RFlogo.RotY = RightFlipper.CurrentAngle
  FlipperLSh.RotZ = LeftFlipper.currentangle
  FlipperRSh.RotZ = RightFlipper.currentangle
End Sub


'**********************************************************************************************************

'Initiate Table
'**********************************************************************************************************

Dim bsTrough, Mag1

Sub Table1_Init
  vpmInit Me
  On Error Resume Next
    With Controller
    .GameName = cGameName
    If Err Then MsgBox "Can't start Game" & cGameName & vbNewLine & Err.Description : Exit Sub
    .SplashInfoLine = "Avatar, Stern 2010"&chr(13)&"You Suck"
    .HandleMechanics=0
    .HandleKeyboard=0
    .ShowDMDOnly=1
    .ShowFrame=0
    .ShowTitle=0
        .hidden = 0
         On Error Resume Next
         .Run GetPlayerHWnd
         If Err Then MsgBox Err.Description
         On Error Goto 0
     End With
     On Error Goto 0
   InitVpmFFlipsSAM
   PinMAMETimer.Interval=PinMAMEInterval
   PinMAMETimer.Enabled=1

  vpmNudge.TiltSwitch=-7
  vpmNudge.Sensitivity=2
  vpmNudge.TiltObj=Array(Bumper1b,Bumper2b,Bumper3b,LeftSlingshot,RightSlingshot)

Set bsTrough = New cvpmBallStack
  bsTrough.InitSw 0, 21, 20, 19, 18, 0, 0, 0
    bsTrough.InitKick BallRelease, 70, 8
    bsTrough.InitExitSnd SoundFX("ballrelease",DOFContactors), SoundFX("Solenoid",DOFContactors)
    bsTrough.Balls = 4

Set mag1= New cvpmMagnet
    mag1.InitMagnet Magnet1, 25
    mag1.GrabCenter = False
    mag1.solenoid=3
    mag1.CreateEvents "mag1"

  ' Create Captive Ball
  RCaptKicker1.CreateBall
  RCaptKicker1.Kick 180,10

  lockPin1.Isdropped = 1
  lockPin2.Isdropped = 1

  InitJakeAni
'  Init3Bank

End Sub

'**********************************************************************************************************
'Plunger code
'**********************************************************************************************************

Sub Table1_KeyDown(ByVal KeyCode)

If keycode = LeftFlipperKey Then FlipperActivate LeftFlipper, LFPress
If keycode = RightFlipperKey Then FlipperActivate RightFlipper, RFPress
if keycode = 19 then ScoreCard=1 : CardTimer.enabled=1
  If KeyDownHandler(keycode) Then Exit Sub
  If keycode = PlungerKey Then Plunger.Pullback:playsoundAt "plungerpull", Plunger
  if KeyCode = LeftTiltKey Then Nudge 90, 4
  if KeyCode = RightTiltKey Then Nudge 270, 4
  if KeyCode = CenterTiltKey Then Nudge 0, 4

End Sub

Sub Table1_KeyUp(ByVal KeyCode)

If keycode = LeftFlipperKey Then FlipperDeActivate LeftFlipper, LFPress
If keycode = RightFlipperKey Then FlipperDeActivate RightFlipper, RFPress
if keycode = 19 then ScoreCard=0
  If KeyUpHandler(keycode) Then Exit Sub
  If keycode = PlungerKey Then Plunger.Fire:PlaySoundAt "plunger", Plunger
  If Keycode = StartGameKey Then Controller.Switch(16) = 0
End Sub

Dim PlungerIM

Const IMPowerSetting = 60 '90
Const IMTime = 0.6
Set plungerIM = New cvpmImpulseP
With plungerIM
  .InitImpulseP swplunger, IMPowerSetting, IMTime
  '.Random 0.3
  .InitExitSnd SoundFX("Popper",DOFContactors), SoundFX("Solenoid",DOFContactors)
  .CreateEvents "plungerIM"
End With

'**********************************************************************************************************

 ' Drain hole and kickers
Sub Drain_Hit
  bsTrough.addball me
  playsound "drain"
End Sub

'Wire Triggers
Sub sw1_Hit:Controller.Switch(1) = 1:PlaySoundAtVol "rollover", ActiveBall, VolRol:End Sub
Sub sw1_UnHit:Controller.Switch(1) = 0:End Sub
Sub sw10_Hit:Controller.Switch(10) = 1:PlaySoundAtVol "rollover", ActiveBall, VolRol:End Sub
Sub sw10_UnHit:Controller.Switch(10) = 0:End Sub
Sub sw11_Hit:Controller.Switch(11) = 1:PlaySoundAtVol "rollover", ActiveBall, VolRol:End Sub
Sub sw11_UnHit:Controller.Switch(11) = 0:End Sub
Sub sw12_Hit:Controller.Switch(12) = 1:PlaySoundAtVol "rollover", ActiveBall, VolRol:End Sub
Sub sw12_UnHit:Controller.Switch(12) = 0:End Sub
Sub sw13_Hit:Controller.Switch(13) = 1:PlaySoundAtVol "rollover", ActiveBall, VolRol:End Sub
Sub sw13_UnHit:Controller.Switch(13) = 0:End Sub
Sub sw14_Hit:Controller.Switch(14) = 1:PlaySoundAtVol "metalping", ActiveBall, VolRol:End Sub
Sub sw14_UnHit:Controller.Switch(14) = 0: PlaySoundAt "fx2_popper", sw14:End Sub
Sub sw23_Hit:Controller.Switch(23) = 1:PlaySoundAtVol "rollover", ActiveBall, VolRol:End Sub
Sub sw23_UnHit:Controller.Switch(23) = 0:End Sub
Sub sw24_Hit:Controller.Switch(24) = 1:PlaySoundAtVol "rollover", ActiveBall, VolRol:End Sub
Sub sw24_UnHit:Controller.Switch(24) = 0:End Sub
Sub sw25_Hit:Controller.Switch(25) = 1:PlaySoundAtVol "rollover", ActiveBall, VolRol:End Sub
Sub sw25_UnHit:Controller.Switch(25) = 0:End Sub
Sub sw28_Hit:Controller.Switch(28) = 1:PlaySoundAtVol "rollover", ActiveBall, VolRol:End Sub
Sub sw28_UnHit:Controller.Switch(28) = 0:End Sub
Sub sw29_Hit:Controller.Switch(29) = 1:PlaySoundAtVol "rollover", ActiveBall, VolRol:End Sub
Sub sw29_UnHit:Controller.Switch(29) = 0:End Sub
Sub sw35_Hit:Controller.Switch(35) = 1:PlaySoundAtVol "rollover", ActiveBall, VolRol:End Sub
Sub sw35_UnHit:Controller.Switch(35) = 0:End Sub

'Captive Ball
Sub sw39_Hit:Controller.Switch(39) = 0:End Sub
Sub sw39_UnHit:Controller.Switch(39) = 1:PlaySoundAtVol "rollover", sw39, VolRol:End Sub

 'Stand Up Targets
Sub sw2_Hit:vpmTimer.PulseSw 2:End Sub
Sub sw3_Hit:vpmTimer.PulseSw 3:End Sub
Sub sw4_Hit:vpmTimer.PulseSw 4:End Sub
Sub sw5_Hit:vpmTimer.PulseSw 5:End Sub
Sub sw7_Hit:vpmTimer.PulseSw 7:PlaySoundAtVol "Target_Hit_" & int(rnd(1)*9)+1 , ActiveBall, VolTarg:End Sub
Sub sw8_Hit:vpmTimer.PulseSw 8:PlaySoundAtVol "Target_Hit_" & int(rnd(1)*9)+1 , ActiveBall, VolTarg:End Sub
Sub sw17_Hit:vpmTimer.PulseSw 17:End Sub
Sub sw36_Hit:vpmTimer.PulseSw 36:End Sub
Sub sw37_Hit:vpmTimer.PulseSw 37:End Sub
Sub sw38_Hit:vpmTimer.PulseSw 38:End Sub

'3 Bank dropping targets
Sub sw42_Hit:vpmTimer.PulseSw 42:backbank.Y = 1:backbank.X = 1:Me.TimerEnabled = 1:PlaySoundAtVol "Target_Hit_" & int(rnd(1)*9)+1 ,ActiveBall, VolTarg:End Sub
Sub sw42_Timer:backbank.Y = 0.8888889:backbank.X = -0.8888889:Me.TimerEnabled = 0:End Sub
Sub sw43_Hit:vpmTimer.PulseSw 43:backbank.Y = 1:backbank.X = 1:Me.TimerEnabled = 1:PlaySoundAtVol "Target_Hit_" & int(rnd(1)*9)+1 ,ActiveBall, VolTarg:End Sub
Sub sw43_Timer:backbank.Y = 0.8888889:backbank.X = -0.8888889:Me.TimerEnabled = 0:End Sub
Sub sw44_Hit:vpmTimer.PulseSw 44:backbank.Y = 1:backbank.X = 1:Me.TimerEnabled = 1:PlaySoundAtVol "Target_Hit_" & int(rnd(1)*9)+1 ,ActiveBall, VolTarg:End Sub
Sub sw44_Timer:backbank.Y = 0.8888889:backbank.X = -0.8888889:Me.TimerEnabled = 0:End Sub

'Spinners
Sub sw9_Spin:vpmTimer.PulseSw 9 : playsoundAtVol "fx_spinner", sw9, VolSpin : End Sub
Sub sw40_Spin:vpmTimer.PulseSw 40 : playsoundAtVol "fx_spinner", sw40, VolSpin : End Sub

'Gate Triggers
Sub sw52_Hit:vpmTimer.PulseSw 52:End Sub

'Bumpers
Sub Bumper1b_Hit : vpmTimer.PulseSw(31) : playsoundAtVol SoundFX("Bumpers_Top_" & int(rnd(1)*5)+1 ,DOFContactors),bumper1b, VolBump: End Sub
Sub Bumper2b_Hit : vpmTimer.PulseSw(30) : playsoundAtVol SoundFX("Bumpers_Top_" & int(rnd(1)*5)+1,DOFContactors),bumper2b, VolBump: End Sub
Sub Bumper3b_Hit : vpmTimer.PulseSw(32) : playsoundAtVol SoundFX("Bumpers_Top_" & int(rnd(1)*5)+1,DOFContactors),bumper3b, VolBump: End Sub

'Generic Ball Sounds
Sub Trigger1_Hit : playsoundAt "fx_ballrampdrop", trigger1 : End Sub
Sub Trigger2_Hit : playsoundAt "fx_ballrampdrop", trigger2 : End Sub
Sub Trigger3_Hit : playsoundAtVol "ramp_hit3", trigger3, .8 : End Sub

'***********************************************
       'Motor Bank
'***********************************************

Sub Sol3bankmotor(Enabled)
  If Enabled then
  If DropADir = 1 Then DropADir = -1 : Else : DropADir = 1
  End if
End Sub

Dim DropADir, DropAPos : DropADir = -1 : DropAa.TimerEnabled = 1

'Animations
Sub DropAa_Timer()
  Controller.Switch(45) = 0:Controller.Switch(46) = 0
  If DropADir = 1 And DropApos>0 then DropApos=DropApos-1 : PlaySoundAtVol "motor2", sw37, VolMot
  If DropADir = -1 And DropApos<25 then DropApos=DropApos+1 : PlaySoundAtVol "motor2", sw37, VolMot
  backbank.z=25-(DropApos*2)
  If DropAPos=0 Then Controller.Switch(46) = 1 : sw42.IsDropped = 0:sw43.IsDropped = 0:sw44.IsDropped = 0:banklsidehelp.IsDropped = 0
  If DropApos=25 Then Controller.Switch(45) = 1 : sw42.IsDropped = 1:sw43.IsDropped = 1:sw44.IsDropped = 1:banklsidehelp.IsDropped = 1
'backbank.z=-24
End Sub

'    RiseBank
'    DropBank

'Sub Init3Bank()
'  DropAPos = 0
'  Controller.Switch(46) = 1
'End Sub

'Sub RiseBank
'  If DropAPos <= 0 Then Exit Sub
'  DropADir = 1
'  DropAPos = 25
'  DropAa.TimerEnabled = 1
'  PlaySoundAtVol "motor2", sw37, VolMot
'End Sub

'Sub DropBank
'  If DropAPos >= 25 Then Exit Sub
'  DropADir = -1
'  DropAPos = 0
'  DropAa.TimerEnabled = 1
'  PlaySoundAtVol "motor2", sw37, VolMot
'End Sub


'  Select Case DropAPos
'    Case 0: backbank.z=25:Controller.Switch(45) = 0:Controller.Switch(46) = 1:sw42.IsDropped = 0:sw43.IsDropped = 0:sw44.IsDropped = 0:banklsidehelp.IsDropped = 0
'      If DropADir = 1 then
'        DropAa.TimerEnabled = 0
'      else
'      end if
'    Case 1: backbank.z=23:Controller.Switch(45) = 0:Controller.Switch(46) = 0
'    Case 2: backbank.z=21
'    Case 3: backbank.z=19
'    Case 4: backbank.z=17
'    Case 5: backbank.z=15
'    Case 6: backbank.z=13
'    Case 7: backbank.z=11
'    Case 8: backbank.z=9
'    Case 9: backbank.z=7
'    Case 10: backbank.z=5
'    Case 11: backbank.z=3
'    Case 12: backbank.z=1
'    Case 13: backbank.z=-1
'    Case 14: backbank.z=-3
'    Case 15: backbank.z=-5
'    Case 16: backbank.z=-7
'    Case 17: backbank.z=-9
'    Case 18: backbank.z=-11
'    Case 19: backbank.z=-13
'    Case 20: backbank.z=-15
'    Case 21: backbank.z=-17
'    Case 22: backbank.z=-19
'    Case 23: backbank.z=-21
'    Case 24: backbank.z=-23:Controller.Switch(45) = 0:Controller.Switch(46) = 0
'    Case 25: backbank.z=-24:Controller.Switch(45) = 1:Controller.Switch(46) = 0:sw42.IsDropped = 1:sw43.IsDropped = 1:sw44.IsDropped = 1:banklsidehelp.IsDropped = 1
'      If DropADir = 1 then
'      else
'        DropAa.TimerEnabled = 0
'      end if
'  End Select

'  If DropADir = 1 then
'    If DropApos>0 then DropApos=DropApos-1
'  else
'    If DropApos<25 then DropApos=DropApos+1
'  end if






'***********************************************
       'Jake Animation
'***********************************************

Dim DAD, DAP
DAD = 1

Sub InitJakeAni()
  DAP = 0
End Sub

Sub RiseJake
  If DAP <= 0 Then Exit Sub
  DAD = 1
  DAP = 23
  jakeopen2.TimerEnabled = 1
  PlaySoundAtVol "motor2", sw17, VolMot
End Sub

Sub DropJake
  If DAP >= 18 Then Exit Sub
  DAD = -1
  DAP = 0
  jakeopen2.TimerEnabled = 1
  PlaySoundAtVol "motor2", sw17, VolMot

End Sub

Sub jakeopen2_Timer()
  Select Case DAP
    Case 0: JakeTop.ObjRotY=0:JakeTop.RotX=160:
      If DAD = 1 then
        jakeopen2.TimerEnabled = 0
      else
      end if

    Case 1: JakeTop.ObjRotY=3:JakeTop.ObjRotX=1:
    Case 2: JakeTop.ObjRotY=6:JakeTop.ObjRotX=2:JakeTop.ObjRotZ=-105:
    Case 3: JakeTop.ObjRotY=9:JakeTop.ObjRotX=3:
    Case 4: JakeTop.ObjRotY=12:JakeTop.ObjRotX=4:JakeTop.ObjRotZ=-106:
    Case 5: JakeTop.ObjRotY=15:JakeTop.ObjRotX=5:
    Case 6: JakeTop.ObjRotY=18:JakeTop.ObjRotX=6:JakeTop.ObjRotZ=-107:
    Case 7: JakeTop.ObjRotY=21:JakeTop.ObjRotX=7:
    Case 8: JakeTop.ObjRotY=24:JakeTop.ObjRotX=8:JakeTop.ObjRotZ=-108:
    Case 9: JakeTop.ObjRotY=27:JakeTop.ObjRotX=9:
    Case 10: JakeTop.ObjRotY=30:JakeTop.ObjRotX=10:JakeTop.ObjRotZ=-109:
    Case 11: JakeTop.ObjRotY=33:JakeTop.ObjRotX=11:
    Case 12: JakeTop.ObjRotY=36:JakeTop.ObjRotX=12:JakeTop.ObjRotZ=-110:
    Case 13: JakeTop.ObjRotY=39:JakeTop.ObjRotX=13:
    Case 14: JakeTop.ObjRotY=42:JakeTop.ObjRotX=14:JakeTop.ObjRotZ=-112:
    Case 15: JakeTop.ObjRotY=45:
    Case 16: JakeTop.ObjRotY=48:
    Case 17: JakeTop.ObjRotY=50:
    Case 18: JakeTop.RotX=140:


      If DAD = 1 then
      else
        jakeopen2.TimerEnabled = 0
      end if
  End Select
  If DAD = 1 then
    If DAP>0 then DAP=DAP-1
  else
    If DAP<18 then DAP=DAP+1
  end if
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
LampTimer.Interval = 5 'lamp fading speed
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
  NFadeL 21, l21
  NFadeL 22, l22
  NFadeL 23, l23
  NFadeL 24, l24
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
  NFadeL 47, l47
  NFadeL 48, l48
  NFadeL 49, l49
  NFadeL 50, l50
  NFadeL 51, l51
  NFadeL 52, l52
  NFadeL 53, l53
  NFadeL 54, l54
  NFadeL 55, l55
  NFadeL 56, l56
  NFadeL 57, l57
  NFadeL 58, l58
  NFadeL 59, l59

  NFadeObjm 60, P60, "lampbulbON", "lampbulb"
  NFadeL 60, l60 'top left bumper
  NFadeObjm 61, P61, "lampbulbON", "lampbulb"
  NFadeL 61, l61 'right bumper
  NFadeObjm 62, P62, "lampbulbON", "lampbulb"
  NFadeL 62, l62 'lower left bumper


' From 32assassins 10.4 template.

'  NFadeObjm 60, P60, "lampbulbON", "lampbulb"
'  NFadeL 60, l60 'top left bumper
'  NFadeObjm 61, P61, "lampbulbON", "lampbulb"
'  NFadeL 61, l61 'right bumper
'  NFadeObjm 62, P62, "lampbulbON", "lampbulb"
'  NFadeL 62, l62 'lower left bumper
'  NFadeL 63, l63
'  NFadeL 64, l64

'  NFadeL 60, l60 'Bumper 2
'  NFadeL 61, l61 'Bumper 1
'  NFadeL 62, l62 'Bumper 3

  NFadeL 63, l63
  NFadeL 64, l64

'Solenoid Controlled Lamps/Flashers

  LampMod 109, f9  ' Top Left Pop Bumber
  LampMod 110, f10 ' Top Right Pop Bumber
  LampMod 111, f11 ' Bottom Pop Bumper

  LampMod 120, f20 ' Left ampsuit flasher
  LampMod 121, f21 ' Link flasher

'  LampMod 120, f21 ' Link flasher
'  LampMod 121, f20  ' ??
'  LampMod 121, F1b ' Rear Flash
'  LampMod 121, F2b
'  LampMod 121, Flasher2
'  LampMod 121, Flasher3
  LampMod 122, f22 ' In front of Jake - fantacy placement
  LampMod 123, f23 ' Right ampsuit flasher
  LampMod 124, f24 ' Coin ? - lets just make a new fantacy one - I don't know where this would go
  LampMod 125, f25 ' Bumper flasher
  LampMod 126, f26 ' Ampsuit mb flasher

  'LampMod 127, f27 ' active ?

  LampMod 122, f22a ' More flashing please
  LampMod 122, f22c ' For ballreflection
  LampMod 121, f22b ' Yes, even more

  LampMod 125, f25a ' Yes, even more
  LampMod 126, f26a ' Wee - this looks better
  LampMod 126, f26b ' Wee - this looks better


  LampMod 128, f28a ' Right side ampsuit flasher
  LampMod 128, f28b ' Left side ampsuit flasher

  LampMod 129, f29c ' Left Side Flash

'  LampMod 129, f29b ' Left Side Flash - for ballreflection


' Left Side Blue (x2)

'  LampMod 130, f29
'  LampMod 130, f29a
'  LampMod 130, f30
'  LampMod 130, f30a
'  LampMod 130, f2a
'  LampMod 130, Flasher6
'  LampMod 130, Flasher7

 ' NFadeObjm 130, f29, "dome3_blueON", "dome3_blue"  'Dome
 ' NFadeL 130, f29a
 ' NFadeObjm 130, f30, "dome3_blueON", "dome3_blue"  'Dome
 ' NFadeObjm 130, f2a, "dome3_blueON", "dome3_blue"  'Dome
 ' NFadeL 130, f30a

' Right Side Blue (x2)

'  LampMod 131, f31
'  LampMod 131, f31a
'  LampMod 131, f32
'  LampMod 131, f32a
'  LampMod 131, f1a
'  LampMod 131, Flasher4
'  LampMod 131, Flasher5

'  NFadeObjm 131, f31, "dome3_blueON", "dome3_blue"  'Dome
'  NFadeL 131, f31a
'  NFadeObjm 131, f32, "dome3_blueON", "dome3_blue"  'Dome
'  NFadeObjm 131, f1a, "dome3_blueON", "dome3_blue"  'Dome
'  NFadeL 131, f32a

  LampMod 132, f32b ' Right Side Flash

'  LampMod 132, f32c ' Right Side Flash - for ballreflection

End Sub

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

 'Reels
Sub FadeReel(nr, reel)
    Select Case FadingLevel(nr)
        Case 2:FadingLevel(nr) = 0
        Case 3:FadingLevel(nr) = 2
        Case 4:reel.Visible = 0:FadingLevel(nr) = 3
        Case 5:reel.Visible = 1:FadingLevel(nr) = 1
    End Select
End Sub

 'Inverted Reels
Sub FadeIReel(nr, reel)
    Select Case FadingLevel(nr)
        Case 2:FadingLevel(nr) = 0
        Case 3:FadingLevel(nr) = 2
        Case 4:reel.Visible = 1:FadingLevel(nr) = 3
        Case 5:reel.Visible = 0:FadingLevel(nr) = 1
    End Select
End Sub

'**********************************************************************************************************
'**********************************************************************************************************
' Start of VPX functions
'**********************************************************************************************************
'**********************************************************************************************************

'**********Sling Shot Animations
' Rstep and Lstep  are the variables that increment the animation
'****************
Dim RStep, Lstep

Sub RightSlingShot_Slingshot
  vpmTimer.PulseSw 27
  PlaySoundAtVol SoundFX("Sling_R" & int(Rnd(1)*8)+1 ,DOFContactors), Sling1, VolSli
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
  vpmTimer.PulseSw 26
    PlaySoundAtVol SoundFX("Sling_L" & int(Rnd(1)*10)+1 ,DOFContactors), Sling2, VolSli
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


' *******************************************************************************************************
' Positional Sound Playback Functions by DJRobX
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

'Set position as bumperX and Vol manually.

Sub PlaySoundAtBumperVol(sound, tableobj, Vol)
  PlaySound sound, 1, Vol, AudioPan(tableobj), 0,0,1, 1, AudioFade(tableobj)
End Sub

'*********************************************************************
'                     Supporting Ball & Sound Functions
'*********************************************************************

Function RndNum(min, max)
    RndNum = Int(Rnd() * (max-min + 1) ) + min ' Sets a random number between min and max
End Function

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
    StopSound("fx2_plasticrolling" & b)
    Next

  ' exit the sub if no balls on the table
    If UBound(BOT) = -1 Then Exit Sub

  ' play the rolling sound for each ball

    For b = 0 to UBound(BOT)
      If BallVel(BOT(b) ) > 1 Then
        rolling(b) = True
        if BOT(b).z < 50 Then ' Ball on playfield
          PlaySound("fx_ballrolling" & b), -1, Vol(BOT(b)), AudioPan(BOT(b) ), 0, Pitch(BOT(b) ), 1, 0, AudioFade(BOT(b) )
        Else ' Ball on raised ramp
      if ( ( BOT(b).z >= 75 ) and ( BOT(b).y <= 130 ) and ( BOT(b).x > 130 ) ) Then
          StopSound ("fx_ballrolling" & b)
          PlaySound("fx2_plasticrolling" & b), -1, VolPla, AudioPan(BOT(b) ), 0, Pitch(BOT(b))+30000, 1, 0, AudioFade(BOT(b) )
    End If
        if ( ( BOT(b).z >= 75 ) and ( BOT(b).y > 1300 ) and ( BOT(b).x < 810 ) ) Then
          StopSound("fx2_plasticrolling" & b)
        End If
      End If
      Else
        If rolling(b) = True Then
          StopSound("fx_ballrolling" & b)
          StopSound("fx2_plasticrolling" & b) ' Should not be needed
          rolling(b) = False
        End If
      End If
    Next
End Sub

'**********************
' Ball Collision Sound
'**********************

Sub OnBallBallCollision(ball1, ball2, velocity)
  PlaySound("fx_collide"), 0, Csng(velocity) ^2 / (VolDiv/VolCol) , AudioPan(ball1), 0, Pitch(ball1), 0, 0, AudioFade(ball1)
End Sub

'*****************************************
'     BALL SHADOW
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
        If BOT(b).X < Table1.Width/2 Then
            BallShadow(b).X = ((BOT(b).X) - (Ballsize/6) + ((BOT(b).X - (Table1.Width/2))/7)) + 13
        Else
            BallShadow(b).X = ((BOT(b).X) + (Ballsize/6) + ((BOT(b).X - (Table1.Width/2))/7)) - 13
        End If
        ballShadow(b).Y = BOT(b).Y + 10
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
  PlaySound "pinhit_low", 0, Vol(ActiveBall)*VolPi, AudioPan(ActiveBall), 0, Pitch(ActiveBall), 0, 0, AudioFade(ActiveBall)
End Sub

Sub Targets_Hit (idx)
  PlaySound "Target_Hit_" & int(rnd(1)*9)+1, 0, Vol(ActiveBall)*VolTarg, AudioPan(ActiveBall), 0, Pitch(ActiveBall), 0, 0, AudioFade(ActiveBall)
End Sub

Sub Metals_Thin_Hit (idx)
  PlaySound "Metal_Touch_" & int(rnd(1)*13)+1, 0, Vol(ActiveBall)*VolMetal, AudioPan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
End Sub

Sub Metals_Medium_Hit (idx)
  PlaySound "Metal_Touch_" & int(rnd(1)*13)+1, 0, Vol(ActiveBall)*VolMetal, AudioPan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
End Sub

Sub Metals2_Hit (idx)
  PlaySound "Metal_Touch_" & int(rnd(1)*13)+1, 0, Vol(ActiveBall)*VolMetal, AudioPan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
End Sub

Sub Gates_Hit (idx)
  PlaySound "gate4", 0, Vol(ActiveBall)*VolGates, AudioPan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
End Sub

Sub Spinner_Spin
  PlaySound "fx_spinner", 0, VolSpin, AudioPan(Spinner), 0.25, 0, 0, 1, AudioFade(Spinner)
End Sub

Sub RampTriggers_Hit(idx)
  dim finalspeed
    finalspeed=SQR(activeball.velx * activeball.velx + activeball.vely * activeball.vely)
  If finalspeed > 10 then
    RandomSoundRamp
  End if
End Sub

Sub RandomSoundRamp()
  Select Case Int(Rnd*3)+1
    Case 1 : PlaySound "fx_rampbump1", 0, 80, AudioPan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
    Case 2 : PlaySound "fx_rampbump2", 0, 80, AudioPan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
    Case 3 : PlaySound "fx_rampbump3", 0, 80, AudioPan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
    End Select
End Sub

Sub Rubbers_Hit(idx)
  dim finalspeed
    finalspeed=SQR(activeball.velx * activeball.velx + activeball.vely * activeball.vely)
  If finalspeed > 20 then
    PlaySound "Rubber_Strong_" & int(rnd(1)*9)+1 , 0, Vol(ActiveBall)*VolRH, AudioPan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
  End if
  If finalspeed >= 6 AND finalspeed <= 20 then
    RandomSoundRubber()
  End If
End Sub

Sub Posts_Hit(idx)
  dim finalspeed
    finalspeed=SQR(activeball.velx * activeball.velx + activeball.vely * activeball.vely)
  If finalspeed > 16 then
    PlaySound "Rubber_Strong_" & int(rnd(1)*9)+1, 0, Vol(ActiveBall)*VolPo, AudioPan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
  End if
  If finalspeed >= 6 AND finalspeed <= 16 then
    RandomSoundRubber()
  End If
End Sub

Sub UpperPosts_Hit(idx)
  dim finalspeed
    finalspeed=SQR(activeball.velx * activeball.velx + activeball.vely * activeball.vely)
  If finalspeed > 16 then
    PlaySound "Rubber_Strong_" & int(rnd(1)*9)+1, 0, Vol(ActiveBall)*VolPo, AudioPan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
  End if
  If finalspeed >= 6 AND finalspeed <= 16 then
    RandomSoundRubber()
  End If
End Sub

Sub RandomSoundRubber()
  Select Case Int(Rnd*3)+1
    Case 1 : PlaySound "rubber_hit_1", 0, Vol(ActiveBall)*VolRH, AudioPan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
    Case 2 : PlaySound "rubber_hit_2", 0, Vol(ActiveBall)*VolRH, AudioPan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
    Case 3 : PlaySound "rubber_hit_3", 0, Vol(ActiveBall)*VolRH, AudioPan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
    Case 3 : PlaySound "Rubber_1", 0, Vol(ActiveBall)*VolRH, AudioPan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
    Case 3 : PlaySound "Rubber_2", 0, Vol(ActiveBall)*VolRH, AudioPan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
    Case 3 : PlaySound "Rubber_3", 0, Vol(ActiveBall)*VolRH, AudioPan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
    Case 3 : PlaySound "Rubber_4", 0, Vol(ActiveBall)*VolRH, AudioPan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
    Case 3 : PlaySound "Rubber_5", 0, Vol(ActiveBall)*VolRH, AudioPan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
    Case 3 : PlaySound "Rubber_6", 0, Vol(ActiveBall)*VolRH, AudioPan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
    Case 3 : PlaySound "Rubber_7", 0, Vol(ActiveBall)*VolRH, AudioPan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
    Case 3 : PlaySound "Rubber_8", 0, Vol(ActiveBall)*VolRH, AudioPan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
    Case 3 : PlaySound "Rubber_9", 0, Vol(ActiveBall)*VolRH, AudioPan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)

Rubber_1


  End Select
End Sub


Sub RandomSoundFlipper()
  Select Case Int(Rnd*3)+1
    Case 1 : PlaySound "flip_hit_1", 0, Vol(ActiveBall)*VolRH, AudioPan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
    Case 2 : PlaySound "flip_hit_2", 0, Vol(ActiveBall)*VolRH, AudioPan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
    Case 3 : PlaySound "flip_hit_3", 0, Vol(ActiveBall)*VolRH, AudioPan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
    Case 4 : PlaySound "Flipper_Rubber_1", 0, Vol(ActiveBall)*VolRH, AudioPan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
    Case 5 : PlaySound "Flipper_Rubber_2", 0, Vol(ActiveBall)*VolRH, AudioPan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
    Case 6 : PlaySound "Flipper_Rubber_3", 0, Vol(ActiveBall)*VolRH, AudioPan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
    Case 7 : PlaySound "Flipper_Rubber_4", 0, Vol(ActiveBall)*VolRH, AudioPan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
    Case 8 : PlaySound "Flipper_Rubber_5", 0, Vol(ActiveBall)*VolRH, AudioPan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
    Case 9 : PlaySound "Flipper_Rubber_6", 0, Vol(ActiveBall)*VolRH, AudioPan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
    Case 10: PlaySound "Flipper_Rubber_7", 0, Vol(ActiveBall)*VolRH, AudioPan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)

  End Select
End Sub


' Thal, while merging changes from Hauntfreaks and 32assassin
' This script has been merged with the latest code from 32 as of 5.10.2017 from vpinball.com

Sub SetLampMod(nr, value)
      If value > 0 Then
      LampState(nr) = 1
    Else
      LampState(nr) = 0
    End If
    FadingLevel(nr) = value
End Sub

Sub LampMod(nr, object)
    If TypeName(object) = "Light" Then
      Object.IntensityScale = FadingLevel(nr)/128
      Object.State = LampState(nr)
    End If
    If TypeName(object) = "Flasher" Then
      Object.IntensityScale = FadingLevel(nr)/128
      Object.visible = LampState(nr)
    End If
    If TypeName(object) = "Primitive" Then
      Object.DisableLighting = LampState(nr)
    End If
End Sub

' Thalamus : Exit in a clean and proper way
Sub Table1_exit()
  Controller.Pause = False
  Controller.Stop
End Sub

'*nfozzy Start


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

'****************************************************************************
'PHYSICS DAMPENERS

'These are data mined bounce curves,
'dialed in with the in-game elasticity as much as possible to prevent angle / spin issues.
'Requires tracking ballspeed to calculate COR


Sub dPosts_Hit(idx)
  RubbersD.dampen Activeball
    dim finalspeed
    finalspeed=SQR(activeball.velx * activeball.velx + activeball.vely * activeball.vely)
  If finalspeed > 20 then
    PlaySound "Rubber_Strong_" & int(rnd(1)*9)+1 , 0, Vol(ActiveBall)*VolRH, AudioPan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
  End if
  If finalspeed >= 6 AND finalspeed <= 20 then
    RandomSoundRubber()
  End If
End Sub

Sub dSleeves_Hit(idx)
  SleevesD.Dampen Activeball
  dim finalspeed
    finalspeed=SQR(activeball.velx * activeball.velx + activeball.vely * activeball.vely)
  If finalspeed > 20 then
    PlaySound "Rubber_Strong_" & int(rnd(1)*9)+1 , 0, Vol(ActiveBall)*VolRH, AudioPan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
  End if
  If finalspeed >= 6 AND finalspeed <= 20 then
    RandomSoundRubber()
  End If
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

' Thalamus - patched :     aBall.velx = aBall.velx * coef : aBall.vely = aBall.vely * coef
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
'     FLIPPER CORRECTION FUNCTIONS
'******************************************************

Sub AddPt(aStr, idx, aX, aY)  'debugger wrapper for adjusting flipper script in-game
  dim a : a = Array(LF, RF)
  dim x : for each x in a
    x.addpoint aStr, idx, aX, aY
  Next
End Sub

Class FlipperPolarity
  Public DebugOn, Enabled
  Private FlipAt  'Timer variable (IE 'flip at 723,530ms...)
  Public TimeDelay  'delay before trigger turns off and polarity is disabled TODO set time!
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

  Public Sub Report(aChooseArray)   'debug, reports all coords in tbPL.text
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
  Private Function FlipperOn() : if gameTime < FlipAt+TimeDelay then FlipperOn = True : End If : End Function 'Timer shutoff for polaritycorrect

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
          if ballpos > 0.65 then  Ycoef = LinearEnvelope(BallData(x).Y, YcoefIn, YcoefOut)        'find safety coefficient 'ycoef' data
        end if
      Next

      If BallPos = 0 Then 'no ball data meaning the ball is entering and exiting pretty close to the same position, use current values.
        BallPos = PSlope(aBall.x, FlipperStart, 0, FlipperEnd, 1)
        if ballpos > 0.65 then  Ycoef = LinearEnvelope(aBall.Y, YcoefIn, YcoefOut)            'find safety coefficient 'ycoef' data
      End If

      'Velocity correction
      if not IsEmpty(VelocityIn(0) ) then
        Dim VelCoef
   :      VelCoef = LinearEnvelope(BallPos, VelocityIn, VelocityOut)

        if partialflipcoef < 1 then VelCoef = PSlope(partialflipcoef, 0, 1, 1, VelCoef)

        if Enabled then aBall.Velx = aBall.Velx*VelCoef
        if Enabled then aBall.Vely = aBall.Vely*VelCoef
      End If

      'Polarity Correction (optional now)
      if not IsEmpty(PolarityIn(0) ) then
        If StartPoint > EndPoint then LR = -1 'Reverse polarity if left flipper
        dim AddX : AddX = LinearEnvelope(BallPos, PolarityIn, PolarityOut) * LR

        if Enabled then aBall.VelX = aBall.VelX + 1 * (AddX*ycoef*PartialFlipcoef)
        'playsound "knocker"
      End If
    End If
    RemoveBall aBall
  End Sub
End Class

'******************************************************
'   FLIPPER POLARITY AND RUBBER DAMPENER
'     SUPPORTING FUNCTIONS
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

' Used for drop targets and flipper tricks
Function Distance(ax,ay,bx,by)
  Distance = SQR((ax - bx)^2 + (ay - by)^2)
End Function



'******************************************************
'     FLIPPER TRICKS
'******************************************************

RightFlipper.timerinterval=1
Rightflipper.timerenabled=True

sub RightFlipper_timer()
  FlipperTricks LeftFlipper, LFPress, LFCount, LFEndAngle, LFState
  FlipperTricks RightFlipper, RFPress, RFCount, RFEndAngle, RFState
end sub


dim LFPress, RFPress, LFCount, RFCount
dim LFState, RFState
dim EOST, EOSA,Frampup, FElasticity,FReturn
dim RFEndAngle, LFEndAngle

EOST = leftflipper.eostorque
EOSA = leftflipper.eostorqueangle
Frampup = LeftFlipper.rampup
FElasticity = LeftFlipper.elasticity
FReturn = LeftFlipper.return
Const EOSTnew = 1
Const EOSAnew = 1
Const EOSRampup = 0
Dim SOSRampup
SOSRampup = 2.5
Const LiveCatch = 16
Const LiveElasticity = 0.45
Const SOSEM = 0.815
Const EOSReturn = 0.055

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
  Dir = Flipper.startangle/Abs(Flipper.startangle)  '-1 for Right Flipper

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
  Dim LiveCatchBounce                             'If live catch is not perfect, it won't freeze ball totally
  Dim CatchTime : CatchTime = GameTime - FCount

  if CatchTime <= LiveCatch and parm > 6 and ABS(Flipper.x - ball.x) > LiveDistanceMin and ABS(Flipper.x - ball.x) < LiveDistanceMax Then
    if CatchTime <= LiveCatch*0.5 Then            'Perfect catch only when catch time happens in the beginning of the window
      LiveCatchBounce = 0
    else
      LiveCatchBounce = Abs((LiveCatch/2) - CatchTime)  'Partial catch when catch happens a bit late
    end If

    If LiveCatchBounce = 0 and ball.velx * Dir > 0 Then ball.velx = 0
    ball.vely = LiveCatchBounce * (32 / LiveCatch) ' Multiplier for inaccuracy bounce
    ball.angmomx= 0
    ball.angmomy= 0
    ball.angmomz= 0
  End If
End Sub

Sub LeftFlipper_Collide(parm)
  CheckLiveCatch Activeball, LeftFlipper, LFCount, parm
  'LeftFlipperCollide parm
  RandomSoundFlipper()
End Sub

Sub RightFlipper_Collide(parm)
  CheckLiveCatch Activeball, RightFlipper, RFCount, parm
  'RightFlipperCollide parm
  RandomSoundFlipper()
End Sub

'*
'* InstructionCard
'*
Dim CardCounter, ScoreCard
Sub CardTimer_Timer
        If scorecard=1 Then
                CardCounter=CardCounter+2
                If CardCounter>50 Then CardCounter=50
        Else
                CardCounter=CardCounter-4
                If CardCounter<0 Then CardCounter=0
        End If
        InstructionCard.transX = CardCounter*6
        InstructionCard.transY = CardCounter*6
        InstructionCard.transZ = -cardcounter*2
'        InstructionCard.objRotX = -cardcounter/2
        InstructionCard.size_x = 1+CardCounter/25
        InstructionCard.size_y = 1+CardCounter/25
        If CardCounter=0 Then
                CardTimer.Enabled=0
                InstructionCard.visible=0
        Else
                InstructionCard.visible=1
        End If
End Sub

'*** FLASHER 129  ***
Sub Flash129(Level)
  Flasherflash129.intensityScale = Level / (5 * 128)
  FlashLevel129 = 1 : FlasherFlash129_Timer
End Sub

'*** FLASHER 129 TIMER  ***
Sub FlasherFlash129_Timer()
  dim flashx3, matdim
  If not Flasherflash129.TimerEnabled Then
    Flasherflash129.TimerEnabled = True
    Flasherflash129.visible = 1
    Flasherlit129.visible = 1
  End If
  flashx3 = FlashLevel129 * FlashLevel129 * FlashLevel129
  Flasherflash129.opacity = 1000 * flashx3
  Flasherlit129.BlendDisableLighting = 10 * flashx3
  Flasherbase129.BlendDisableLighting =  flashx3
  Flasherlight129.IntensityScale = flashx3
  F129.IntensityScale = flashx3
  matdim = Round(10 * FlashLevel129) -1
  Flasherlit129.material = "domelit" & matdim
  FlashLevel129 = FlashLevel129 * 0.85 - 0.01
  If FlashLevel129 < 0.15 Then
    Flasherlit129.visible = 0
  Else
    Flasherlit129.visible = 1
  end If
  If FlashLevel129 < 0 Then
    Flasherflash129.TimerEnabled = False
    Flasherflash129.visible = 0
  End If
End Sub

'*** FLASHER 130  ***
Sub Flash130(Level)
  Flasherflash130.intensityScale = Level / (5 * 128)
  Flasherflash130a.intensityScale = Level / (5 * 128)
  FlashLevel130 = 1 : FlasherFlash130_Timer
End Sub

'*** FLASHER 130 TIMER  ***
Sub FlasherFlash130_Timer()
  dim flashx3, matdim
  If not Flasherflash130.TimerEnabled Then
    Flasherflash130.TimerEnabled = True
    Flasherflash130.visible = 1
  Flasherflash130a.visible = 1
    Flasherlit130.visible = 1
  Flasherlit130a.visible = 1
  End If
  flashx3 = FlashLevel130 * FlashLevel130 * FlashLevel130
  Flasherflash130.opacity = 1000*flashx3
  Flasherflash130a.opacity = 1000*flashx3
  Flasherlit130.BlendDisableLighting = 10 * flashx3
  Flasherlit130a.BlendDisableLighting = 10 * flashx3
  Flasherbase130.BlendDisableLighting =  flashx3
  Flasherbase130a.BlendDisableLighting =  flashx3
  Flasherlight130.IntensityScale = flashx3
  Flasherlight130a.IntensityScale = flashx3
  matdim = Round(10 * FlashLevel130)
  Flasherlit130.material = "domelit" & matdim -1
  Flasherlit130a.material = "domelit" & matdim -1
  FlashLevel130 = FlashLevel130 * 0.85 - 0.01
  If FlashLevel130 < 0.15 Then
    Flasherlit130.visible = 0
  Flasherlit130a.visible = 0
  Else
    Flasherlit130.visible = 1
    Flasherlit130a.visible = 1
  end If
  If FlashLevel130 < 0 Then
    Flasherflash130.TimerEnabled = False
    Flasherflash130.visible = 0
  Flasherflash130a.visible = 0
  End If
End Sub

'*** FLASHER 131  ***
Sub Flash131(Level)
  Flasherflash131.intensityScale = Level / (5 * 128)
  FlashLevel131 = 1 : FlasherFlash131_Timer
End Sub

'*** FLASHER 131 TIMER  ***
Sub FlasherFlash131_Timer()
  dim flashx3, matdim
  If not Flasherflash131.TimerEnabled Then
    Flasherflash131.TimerEnabled = True
    Flasherflash131.visible = 1
  Flasherflash131a.visible = 1
    Flasherlit131.visible = 1
  Flasherlit131a.visible = 1
  End If
  flashx3 = FlashLevel131 * FlashLevel131 * FlashLevel131
  Flasherflash131.opacity = 1000*flashx3
  Flasherflash131a.opacity = 1000*flashx3
  Flasherlit131.BlendDisableLighting = 10 * flashx3
  Flasherlit131a.BlendDisableLighting = 10 * flashx3
  Flasherbase131.BlendDisableLighting =  flashx3
  Flasherbase131a.BlendDisableLighting =  flashx3
  Flasherlight131.IntensityScale = flashx3
  Flasherlight131a.IntensityScale = flashx3
  matdim = Round(10 * FlashLevel131)
  Flasherlit131.material = "domelit" & matdim -1
  Flasherlit131a.material = "domelit" & matdim -1
  FlashLevel131 = FlashLevel131 * 0.85 - 0.01
  If FlashLevel131 < 0.15 Then
    Flasherlit131.visible = 0
  Flasherlit131a.visible = 0
  Else
    Flasherlit131.visible = 1
  Flasherlit131a.visible = 1
  end If
  If FlashLevel131 < 0 Then
    Flasherflash131.TimerEnabled = False
    Flasherflash131.visible = 0
  Flasherflash131a.visible = 0
  End If
End Sub

'*** FLASHER 132  ***
Sub Flash132(Level)
  Flasherflash132.intensityScale = Level / (5 * 128)
  FlashLevel132 = 1 : FlasherFlash132_Timer
End Sub

'*** FLASHER 132 TIMER  ***
Sub FlasherFlash132_Timer()
  dim flashx3, matdim
  If not Flasherflash132.TimerEnabled Then
    Flasherflash132.TimerEnabled = True
    Flasherflash132.visible = 1
    Flasherlit132.visible = 1
  End If
  flashx3 = FlashLevel132 * FlashLevel132 * FlashLevel132
  Flasherflash132.opacity = 1000 * flashx3
  Flasherlit132.BlendDisableLighting = 10 * flashx3
  Flasherbase132.BlendDisableLighting =  flashx3
  Flasherlight132.IntensityScale = flashx3
  F132.IntensityScale = flashx3
  matdim = Round(10 * FlashLevel132) -1
  Flasherlit132.material = "domelit" & matdim
  FlashLevel132 = FlashLevel132 * 0.85 - 0.01
  If FlashLevel132 < 0.15 Then
    Flasherlit132.visible = 0
  Else
    Flasherlit132.visible = 1
  end If
  If FlashLevel132 < 0 Then
    Flasherflash132.TimerEnabled = False
    Flasherflash132.visible = 0
  End If
End Sub

Dim TestFlashers, TableRef, FlasherLightIntensity, FlasherFlareIntensity, FlasherOffBrightness

                ' *********************************************************************
TestFlashers = 0        ' *** set this to 1 to check position of flasher object       ***
Set TableRef = Table1       ' *** change this, if your table has another name             ***
FlasherLightIntensity = 1   ' *** lower this, if the VPX lights are too bright (i.e. 0.1)   ***
FlasherFlareIntensity = 1   ' *** lower this, if the flares are too bright (i.e. 0.1)     ***
FlasherOffBrightness = 0.5    ' *** brightness of the flasher dome when switched off (range 0-2)  ***
                ' *********************************************************************

Dim ObjLevel(20), objbase(20), objlit(20), objflasher(20), objlight(20)
Dim tablewidth, tableheight : tablewidth = TableRef.width : tableheight = TableRef.height
'initialise the flasher color, you can only choose from "green", "red", "purple", "blue", "white" and "yellow"
InitFlasher 1, "blue" : InitFlasher 2, "blue" : InitFlasher 3, "blue"
InitFlasher 4, "blue" : InitFlasher 5, "blue" : InitFlasher 6, "blue"
'InitFlasher 7, "green" : InitFlasher 8, "red"
'InitFlasher 9, "green" : InitFlasher 10, "red" : InitFlasher 11, "white"
' rotate the flasher with the command below (first argument = flasher nr, second argument = angle in degrees)
'RotateFlasher 4,17 : RotateFlasher 5,0 : RotateFlasher 6,90
'RotateFlasher 7,0 : RotateFlasher 8,0
'RotateFlasher 9,-45 : RotateFlasher 10,90 : RotateFlasher 11,90

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
Sub FlasherFlash4_Timer() : FlashFlasher(4) : End Sub
Sub FlasherFlash5_Timer() : FlashFlasher(5) : End Sub
Sub FlasherFlash6_Timer() : FlashFlasher(6) : End Sub
Sub FlasherFlash7_Timer() : FlashFlasher(7) : End Sub
Sub FlasherFlash8_Timer() : FlashFlasher(8) : End Sub
Sub FlasherFlash9_Timer() : FlashFlasher(9) : End Sub
Sub FlasherFlash10_Timer() : FlashFlasher(10) : End Sub
Sub FlasherFlash11_Timer() : FlashFlasher(11) : End Sub

Sub Flash1(level)
  Objlevel(1) = level/255: FlasherFlash1_Timer
End Sub
Sub Flash2(level)
  Objlevel(2) = level/255  : FlasherFlash2_Timer
End Sub
Sub Flash3(level)
  Objlevel(3) = level/255  : FlasherFlash3_Timer
  Objlevel(4) = level/255  : FlasherFlash4_Timer
End Sub
Sub Flash4(level)
  Objlevel(5) = level/255  : FlasherFlash5_Timer
  Objlevel(6) = level/255  : FlasherFlash6_Timer
End Sub



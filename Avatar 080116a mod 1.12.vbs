Option Explicit
Randomize

' Thalamus 2018-07-24
' Added/Updated "Positional Sound Playback Functions" and "Supporting Ball & Sound Functions"
' Improved directional sounds.
' Added InitVpmFFlipsSAM

On Error Resume Next
ExecuteGlobal GetTextFile("controller.vbs")
If Err Then MsgBox "You need the controller.vbs in order to run this table, available in the vp10 package"
On Error Goto 0

Const BallMass = 1.7
Const cGameName="avr_200",UseSolenoids=1,UseLamps=0,UseGI=0,SSolenoidOn="SolOn",SSolenoidOff="SolOff", SCoin="coin", UseVPMModSol=1

LoadVPM "01560000", "sam.VBS", 3.10
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
SolModCallback(29) = "SetLampMod 129,"  ' upper left blue dome | Doc as Left Side Flash
SolModCallback(30) = "SetLampMod 130,"  ' lower left blue dome | Doc as Left Side Blue (x2)
SolModCallback(31) = "SetLampMod 131,"  ' upper right blue dome | Doc as Right Side Blue (x2)
SolModCallback(32) = "SetLampMod 132,"  ' lower left blue dome | Doc as Right Side Flash


Sub SolLFlipper(Enabled)
     If Enabled Then
         PlaySoundAt SoundFX("fx_Flipperup",DOFContactors), LeftFlipper:LeftFlipper.RotateToEnd
     Else
         PlaySoundAt SoundFX("fx_Flipperdown",DOFContactors), LeftFlipper:LeftFlipper.RotateToStart
     End If
  End Sub

Sub SolRFlipper(Enabled)
     If Enabled Then
         PlaySoundAt SoundFX("fx_Flipperup",DOFContactors),RightFlipper:RightFlipper.RotateToEnd
     Else
         PlaySoundAt SoundFX("fx_Flipperdown",DOFContactors),RightFlipper:RightFlipper.RotateToStart
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
        PlaySoundAt "fx_relay", sw35
	Else
		For each xx in GI:xx.State = 0: Next
		For each xxx in GI2:xxx.visible = 0: Next
        PlaySoundAt "fx_relay", sw35
	End If
End Sub

Sub ampsuitmotorRelay(Enabled)
  If Enabled Then
    '	Controller.Switch(57) = 1:Controller.Switch(58) = 0
    ampf.rotatetoend
    debug.print "ampsuitmotorrelay enabled"
	PlaySoundAt "Diverter", ampf
  Else
    '	Controller.Switch(57) = 0:Controller.Switch(58) = 1
    debug.print "ampsuitmotorrelay disabled"
    ampf.rotatetostart
	PlaySoundAt "Diverter", ampf
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
    RiseJake
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
	End If
 End Sub


Sub Sol3bankmotor(Enabled)
  If Enabled then
    RiseBank
    DropBank
  end if
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
'	RampGate2.RotZ = -(Gate3.currentangle)
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
  Init3Bank

End Sub

'**********************************************************************************************************
'Plunger code
'**********************************************************************************************************

Sub Table1_KeyDown(ByVal KeyCode)
	If KeyDownHandler(keycode) Then Exit Sub
	If keycode = PlungerKey Then Plunger.Pullback:playsoundAt "plungerpull", Plunger
	if KeyCode = LeftTiltKey Then Nudge 90, 4
	if KeyCode = RightTiltKey Then Nudge 270, 4
	if KeyCode = CenterTiltKey Then Nudge 0, 4

End Sub

Sub Table1_KeyUp(ByVal KeyCode)
	If KeyUpHandler(keycode) Then Exit Sub
	If keycode = PlungerKey Then Plunger.Fire:PlaySoundAt "plunger", Plunger
	If Keycode = StartGameKey Then Controller.Switch(16) = 0
End Sub

Dim PlungerIM

Const IMPowerSetting = 40
Const IMTime = 0.6
Set plungerIM = New cvpmImpulseP
With plungerIM
  .InitImpulseP swplunger, IMPowerSetting, IMTime
  .Random 0.3
  .InitExitSnd SoundFX("Popper",DOFContactors), SoundFX("Solenoid",DOFContactors)
  .CreateEvents "plungerIM"
End With

'**********************************************************************************************************

 ' Drain hole and kickers
Sub Drain_Hit:bsTrough.addball me : playsound"drain" : End Sub

'Wire Triggers
Sub sw1_Hit:Controller.Switch(1) = 1:PlaySoundAt "rollover", ActiveBall:End Sub
Sub sw1_UnHit:Controller.Switch(1) = 0:End Sub
Sub sw10_Hit:Controller.Switch(10) = 1:PlaySoundAt "rollover", ActiveBall:End Sub
Sub sw10_UnHit:Controller.Switch(10) = 0:End Sub
Sub sw11_Hit:Controller.Switch(11) = 1:PlaySoundAt "rollover", ActiveBall:End Sub
Sub sw11_UnHit:Controller.Switch(11) = 0:End Sub
Sub sw12_Hit:Controller.Switch(12) = 1:PlaySoundAt "rollover", ActiveBall:End Sub
Sub sw12_UnHit:Controller.Switch(12) = 0:End Sub
Sub sw13_Hit:Controller.Switch(13) = 1:PlaySoundAt "rollover", ActiveBall:End Sub
Sub sw13_UnHit:Controller.Switch(13) = 0:End Sub
Sub sw14_Hit:Controller.Switch(14) = 1:PlaySoundAt "rollover", ActiveBall:End Sub
Sub sw14_UnHit:Controller.Switch(14) = 0:End Sub
Sub sw23_Hit:Controller.Switch(23) = 1:PlaySoundAt "rollover", ActiveBall:End Sub
Sub sw23_UnHit:Controller.Switch(23) = 0:End Sub
Sub sw24_Hit:Controller.Switch(24) = 1:PlaySoundAt "rollover", ActiveBall:End Sub
Sub sw24_UnHit:Controller.Switch(24) = 0:End Sub
Sub sw25_Hit:Controller.Switch(25) = 1:PlaySoundAt "rollover", ActiveBall:End Sub
Sub sw25_UnHit:Controller.Switch(25) = 0:End Sub
Sub sw28_Hit:Controller.Switch(28) = 1:PlaySoundAt "rollover", ActiveBall:End Sub
Sub sw28_UnHit:Controller.Switch(28) = 0:End Sub
Sub sw29_Hit:Controller.Switch(29) = 1:PlaySoundAt "rollover", ActiveBall:End Sub
Sub sw29_UnHit:Controller.Switch(29) = 0:End Sub
Sub sw35_Hit:Controller.Switch(35) = 1:PlaySoundAt "rollover", ActiveBall:End Sub
Sub sw35_UnHit:Controller.Switch(35) = 0:End Sub

'Captive Ball
Sub sw39_Hit:Controller.Switch(39) = 0:End Sub
Sub sw39_UnHit:Controller.Switch(39) = 1:PlaySoundAt "rollover", sw39:End Sub

 'Stand Up Targets
Sub sw2_Hit:vpmTimer.PulseSw 2:End Sub
Sub sw3_Hit:vpmTimer.PulseSw 3:End Sub
Sub sw4_Hit:vpmTimer.PulseSw 4:End Sub
Sub sw5_Hit:vpmTimer.PulseSw 5:End Sub
Sub sw7_Hit:vpmTimer.PulseSw 7:PlaySoundAt "Target", ActiveBall:End Sub
Sub sw8_Hit:vpmTimer.PulseSw 8:PlaySoundAt "Target", ActiveBall:End Sub
Sub sw17_Hit:vpmTimer.PulseSw 17:End Sub
Sub sw36_Hit:vpmTimer.PulseSw 36:End Sub
Sub sw37_Hit:vpmTimer.PulseSw 37:End Sub
Sub sw38_Hit:vpmTimer.PulseSw 38:End Sub

'3 Bank dropping targets
Sub sw42_Hit:vpmTimer.PulseSw 42:backbank.Y = 1:backbank.X = 1:Me.TimerEnabled = 1:PlaySoundAt "Target",ActiveBall:End Sub
Sub sw42_Timer:backbank.Y = 0.8888889:backbank.X = -0.8888889:Me.TimerEnabled = 0:End Sub
Sub sw43_Hit:vpmTimer.PulseSw 43:backbank.Y = 1:backbank.X = 1:Me.TimerEnabled = 1:PlaySoundAt "Target",ActiveBall:End Sub
Sub sw43_Timer:backbank.Y = 0.8888889:backbank.X = -0.8888889:Me.TimerEnabled = 0:End Sub
Sub sw44_Hit:vpmTimer.PulseSw 44:backbank.Y = 1:backbank.X = 1:Me.TimerEnabled = 1:PlaySoundAt "Target",ActiveBall:End Sub
Sub sw44_Timer:backbank.Y = 0.8888889:backbank.X = -0.8888889:Me.TimerEnabled = 0:End Sub

'Spinners
Sub sw9_Spin:vpmTimer.PulseSw 9 : playsoundAt "fx_spinner", sw9 : End Sub
Sub sw40_Spin:vpmTimer.PulseSw 40 : playsoundAt "fx_spinner", sw40 : End Sub

'Gate Triggers
Sub sw52_Hit:vpmTimer.PulseSw 52:End Sub

'Bumpers
Sub Bumper1b_Hit : vpmTimer.PulseSw(31) : playsoundAt SoundFX("fx_bumper2",DOFContactors),bumper1b: End Sub
Sub Bumper2b_Hit : vpmTimer.PulseSw(30) : playsoundAt SoundFX("fx_bumper3",DOFContactors),bumper2b: End Sub
Sub Bumper3b_Hit : vpmTimer.PulseSw(32) : playsoundAt SoundFX("fx_bumper4",DOFContactors),bumper3b: End Sub

'Generic Ball Sounds
Sub Trigger1_Hit : playsoundAt "fx_ballrampdrop", trigger1 : End Sub
Sub Trigger2_Hit : playsoundAt "fx_ballrampdrop", trigger2 : End Sub


'***********************************************
			 'Motor Bank
'***********************************************

Dim DropADir
Dim DropAPos

DropADir = 1

Sub Init3Bank()
  DropAPos = 0
  Controller.Switch(46) = 1
End Sub

Sub RiseBank
  If DropAPos <= 0 Then Exit Sub
  DropADir = 1
  DropAPos = 25
  DropAa.TimerEnabled = 1
  PlaySoundAt "motor2", sw37
End Sub

Sub DropBank
  If DropAPos >= 25 Then Exit Sub
  DropADir = -1
  DropAPos = 0
  DropAa.TimerEnabled = 1
  PlaySoundAt "motor2", sw37
End Sub

'Animations

Sub DropAa_Timer()
  Select Case DropAPos
    Case 0: backbank.z=25:Controller.Switch(45) = 0:Controller.Switch(46) = 1:sw42.IsDropped = 0:sw43.IsDropped = 0:sw44.IsDropped = 0:banklsidehelp.IsDropped = 0
      If DropADir = 1 then
        DropAa.TimerEnabled = 0
      else
      end if
    Case 1: backbank.z=23:Controller.Switch(45) = 0:Controller.Switch(46) = 0
    Case 2: backbank.z=21:Controller.Switch(45) = 0:Controller.Switch(46) = 0
    Case 3: backbank.z=19:Controller.Switch(45) = 0:Controller.Switch(46) = 0
    Case 4: backbank.z=17:Controller.Switch(45) = 0:Controller.Switch(46) = 0
    Case 5: backbank.z=15:Controller.Switch(45) = 0:Controller.Switch(46) = 0
    Case 6: backbank.z=13:Controller.Switch(45) = 0:Controller.Switch(46) = 0
    Case 7: backbank.z=11:Controller.Switch(45) = 0:Controller.Switch(46) = 0
    Case 8: backbank.z=9:Controller.Switch(45) = 0:Controller.Switch(46) = 0
    Case 9: backbank.z=7:Controller.Switch(45) = 0:Controller.Switch(46) = 0
    Case 10: backbank.z=5:Controller.Switch(45) = 0:Controller.Switch(46) = 0
    Case 11: backbank.z=3:Controller.Switch(45) = 0:Controller.Switch(46) = 0
    Case 12: backbank.z=1:Controller.Switch(45) = 0:Controller.Switch(46) = 0
    Case 13: backbank.z=-1:Controller.Switch(45) = 0:Controller.Switch(46) = 0
    Case 14: backbank.z=-3:Controller.Switch(45) = 0:Controller.Switch(46) = 0
    Case 15: backbank.z=-5:Controller.Switch(45) = 0:Controller.Switch(46) = 0
    Case 16: backbank.z=-7:Controller.Switch(45) = 0:Controller.Switch(46) = 0
    Case 17: backbank.z=-9:Controller.Switch(45) = 0:Controller.Switch(46) = 0
    Case 18: backbank.z=-11:Controller.Switch(45) = 0:Controller.Switch(46) = 0
    Case 19: backbank.z=-13:Controller.Switch(45) = 0:Controller.Switch(46) = 0
    Case 20: backbank.z=-15:Controller.Switch(45) = 0:Controller.Switch(46) = 0
    Case 21: backbank.z=-17:Controller.Switch(45) = 0:Controller.Switch(46) = 0
    Case 22: backbank.z=-19:Controller.Switch(45) = 0:Controller.Switch(46) = 0
    Case 23: backbank.z=-21:Controller.Switch(45) = 0:Controller.Switch(46) = 0
    Case 24: backbank.z=-23:Controller.Switch(45) = 0:Controller.Switch(46) = 0
    Case 25: backbank.z=-24:Controller.Switch(45) = 1:Controller.Switch(46) = 0:sw42.IsDropped = 1:sw43.IsDropped = 1:sw44.IsDropped = 1:banklsidehelp.IsDropped = 1
      If DropADir = 1 then
      else
        DropAa.TimerEnabled = 0
      end if
  End Select
  If DropADir = 1 then
    If DropApos>0 then DropApos=DropApos-1
  else
    If DropApos<25 then DropApos=DropApos+1
  end if
End Sub

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
  PlaySoundAt "motor2", sw17
End Sub

Sub DropJake
  If DAP >= 17 Then Exit Sub
  DAD = -1
  DAP = 0
  jakeopen2.TimerEnabled = 1
  PlaySoundAt "motor2", sw17
End Sub

Sub jakeopen2_Timer()
  Select Case DAP
    Case 0: JakeTop.ObjRotY=0
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
    'Case 18: JakeTop.ObjRotY=55:
    'Case 19: JakeTop.ObjRotY=58:

      If DAD = 1 then
      else
        jakeopen2.TimerEnabled = 0
      end if
  End Select
  If DAD = 1 then
    If DAP>0 then DAP=DAP-1
  else
    If DAP<17 then DAP=DAP+1
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
  LampMod 122, f22 ' In front of Jake - fantacy placement
  LampMod 123, f23 ' Right ampsuit flasher
  LampMod 124, f24 ' Coin ? - lets just make a new fantacy one - I don't know where this would go
  LampMod 125, f25 ' Bumper flasher
  LampMod 126, f26 ' Ampsuit mb flasher

  'LampMod 127, f27 ' active ?

  LampMod 122, f22a ' More flashing please
  LampMod 122, f22c ' For ballreflection
  LampMod 121, f22b	' Yes, even more

  LampMod 125, f25a	' Yes, even more
  LampMod 126, f26a ' Wee - this looks better
  LampMod 126, f26b ' Wee - this looks better


  LampMod 128, f28a ' Right side ampsuit flasher
  LampMod 128, f28b ' Left side ampsuit flasher

  LampMod 129, f29c ' Left Side Flash

'  LampMod 129, f29b ' Left Side Flash - for ballreflection


' Left Side Blue (x2)

  LampMod 130, f29
  LampMod 130, f29a
  LampMod 130, f30
  LampMod 130, f30a
  LampMod 130, f2a
  LampMod 130, f2b

'  NFadeObjm 130, f29, "dome3_blueON", "dome3_blue"  'Dome
'  NFadeL 130, f29a
'  NFadeObjm 130, f30, "dome3_blueON", "dome3_blue"  'Dome
'  NFadeL 130, f30a

' Right Side Blue (x2)

  LampMod 131, f31
  LampMod 131, f31a
  LampMod 131, f32
  LampMod 131, f32a
  LampMod 131, f1a
  LampMod 131, f1b

'  NFadeObjm 131, f31, "dome3_blueON", "dome3_blue"  'Dome
'  NFadeL 131, f31a
'  NFadeObjm 131, f32, "dome3_blueON", "dome3_blue"  'Dome
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
'	Start of VPX functions
'**********************************************************************************************************
'**********************************************************************************************************

'**********Sling Shot Animations
' Rstep and Lstep  are the variables that increment the animation
'****************
Dim RStep, Lstep

Sub RightSlingShot_Slingshot
	vpmTimer.PulseSw 27
	PlaySoundAt SoundFX("right_slingshot",DOFContactors), Sling1
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
    PlaySoundAt SoundFX("left_slingshot",DOFContactors), Sling2
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

Sub PlaySoundAtBall(soundname)
    PlaySoundAt soundname, ActiveBall
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


' *********************************************************************
'                      Supporting Ball & Sound Functions
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
        If BallVel(BOT(b) ) > 1 AND BOT(b).z < 30 Then
            rolling(b) = True
            ' PlaySound("fx_ballrolling" & b), -1, Vol(BOT(b) ), Pan(BOT(b) ), 0, Pitch(BOT(b) ), 1, 0
            PlaySound("fx_ballrolling" & b), -1, Vol(BOT(b)), AudioPan(BOT(b)), 0, Pitch(BOT(b)), 1, 0, AudioFade(BOT(b))
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
'	PlaySound("fx_collide"), 0, Csng(velocity) ^2 / 2000, Pan(ball1), 0, Pitch(ball1), 0, 0
	PlaySound("fx_collide"), 0, Csng(velocity) ^2 / 2000, AudioPan(ball1), 0, Pitch(ball1), 0, 0, AudioFade(ball1)
End Sub

'*****************************************
'			BALL SHADOW
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
	PlaySound "gate4", 0, Vol(ActiveBall), AudioPan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
End Sub

Sub Spinner_Spin
	PlaySound "fx_spinner", 0, .25, AudioPan(Spinner), 0.25, 0, 0, 1, AudioFade(Spinner)
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


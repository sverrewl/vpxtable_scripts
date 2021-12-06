'*********************************************************************
'* GRONI PINBALL PRESENTS ********************************************
'*********************************************************************
'* A VISUAL PINBAL X RELEASE *****************************************
'*********************************************************************
'* ATTACK FROM MARS **************************************************
'*********************************************************************
'* A PINBALL MACHINE BY BALLY 1995 ***********************************
'*********************************************************************
'* THANKS TO JP FOR HELPING ME WITH PARTS OF THE SCRIPT **************
'*********************************************************************
'* ALIEN MODELS ARE ALSO FROM JP´S AFM TABLE WITH SOME MODIFICATIONS *
'* ON THE ARMS TO MAKE THEM SHAKE. ALSO TEXTURES HAS BEEN OPTIMIZED **
'*********************************************************************
'* THANKS TO ARNGRIM FOR ADDING DOF/B2S ******************************
'*********************************************************************

'*********************************************************************
'* LAYERS ************************************************************
'*********************************************************************
'* LAYER 1 = STANDARD VP OBJECTS *************************************
'* LAYER 2 = 3D OBJECTS **********************************************
'* LAYER 3 = INSERT LIGHTS *******************************************
'* LAYER 4 = FLASHER LIGHTS ******************************************
'* LAYER 5 = GI ******************************************************
'*********************************************************************

Option Explicit
Randomize

'*********************************************************************
'* DOF CONTROLLER INIT ***********************************************
'*********************************************************************

On Error Resume Next
ExecuteGlobal GetTextFile("controller.vbs")
If Err Then MsgBox "You need the controller.vbs in order to run this table, available in the vp10 package"
On Error Goto 0

'*********************************************************************
'* VARIABLES *********************************************************
'*********************************************************************

Dim bsTrough, plungerIM, Ballspeed, bsL, bsR, aBall, bBall, dtDrop

'*********************************************************************
'* VPM INIT **********************************************************
'*********************************************************************

Const BallSize = 51

LoadVPM "01560000", "WPC.VBS", 3.46

Sub LoadVPM(VPMver, VBSfile, VBSver)
On Error Resume Next
If ScriptEngineMajorVersion < 5 Then MsgBox "VB Script Engine 5.0 or higher required"
ExecuteGlobal GetTextFile(VBSfile)
If Err Then MsgBox "Unable to open " & VBSfile & ". Ensure that it is in the same folder as this table. " & vbNewLine & Err.Description
Set Controller = CreateObject("VPinMAME.Controller")
Set Controller = CreateObject("B2S.Server")
If Err Then MsgBox "Can't Load VPinMAME." & vbNewLine & Err.Description
If VPMver > "" Then If Controller.Version < VPMver Or Err Then MsgBox "VPinMAME ver " & VPMver & " required."
If VPinMAMEDriverVer < VBSver Or Err Then MsgBox VBSFile & " ver " & VBSver & " or higher required."
On Error Goto 0
End Sub

Const UseSolenoids = 1
Const UseLamps = 0
Const UseSync = 0
Const HandleMech = 0

Const SCoin = "fx_Coin"

Set GiCallback2 = GetRef("UpdateGI")

'*********************************************************************
'* GAME NAME *********************************************************
'*********************************************************************

Const cGameName = "afm_113b"

'*********************************************************************
'* STARTUP SETTINGS **************************************************
'*********************************************************************

SW45.isdropped=1:SW46.isdropped=1:SW47.isdropped=1

RightSlingRubber_C.visible=0:RightSlingRubber_B.visible=0:RightSlingRubber_A.visible=0

LeftSlingRubber_C.visible=0:LeftSlingRubber_B.visible=0:LeftSlingRubber_A.visible=0

Alien1h1.visible=0:Alien1h2.visible=0:Alien1r1.visible=0:Alien1r2.visible=0

Alien2h1.visible=0:Alien2h2.visible=0:Alien2r1.visible=0:Alien2r2.visible=0

Alien3h1.visible=0:Alien3h2.visible=0:Alien3r1.visible=0:Alien3r2.visible=0

Alien4h1.visible=0:Alien4h2.visible=0:Alien4r1.visible=0:Alien4r2.visible=0

MotorBank.Z=-76:SW45P.Z=-76:SW46P.Z=-76:SW47P.Z=-76

Diverter8.isdropped=1:Diverter7.isdropped=1:Diverter6.isdropped=1:Diverter5.isdropped=1

Diverter4.isdropped=1:Diverter3.isdropped=1:Diverter2.isdropped=1:Diverter1.isdropped=1

DivPos=0

TBPos=28:TBTimer.Enabled=0:TBDown=1:Controller.Switch(66) = 1:Controller.Switch(67) = 0

'*********************************************************************
'*********************************************************************
'*********************************************************************
'* TABLE INIT ********************************************************
'*********************************************************************
'*********************************************************************
'*********************************************************************

Sub AFM_Init
vpmInit Me
With Controller
.GameName = cGameName
If Err Then MsgBox "Can't start Game " & cGameName & vbNewLine & Err.Description:Exit Sub
.SplashInfoLine = "Attack from Mars" & vbNewLine & "VPX table by GRONI v1.0"
.Games(cGameName).Settings.Value("rol") = 0
.HandleKeyboard = 0
.ShowTitle = 0
.ShowDMDOnly = 1
.ShowFrame = 0
.HandleMechanics = 0
.Hidden = 0
On Error Resume Next
.Run GetPlayerHWnd
If Err Then MsgBox Err.Description
On Error Goto 0
.Switch(22) = 1 'close coin door
.Switch(24) = 1 'and keep it close
End With

'*********************************************************************
'* PINMAME TIMER *****************************************************
'*********************************************************************

PinMAMETimer.Interval = PinMAMEInterval
PinMAMETimer.Enabled = 1

'*********************************************************************
'* BALL TROUGH *******************************************************
'*********************************************************************

Set bsTrough = New cvpmBallStack
With bsTrough
.InitSw 0, 32, 33, 34, 35, 0, 0, 0
.InitKick BallRelease, 90, 4
.InitEntrySnd "", ""
.InitExitSnd SoundFX("BallRelease",DOFContactors), SoundFX("BallRelease",DOFContactors)
.Balls = 4
End With

'*********************************************************************
'* IMPULSE PLUNGER ***************************************************
'*********************************************************************

Const IMPowerSetting = 42 'Plunger Power
Const IMTime = 0.6        ' Time in seconds for Full Plunge
Set plungerIM = New cvpmImpulseP
With plungerIM
.InitImpulseP swplunger, IMPowerSetting, IMTime
.Random 0.3
.switch 18
.InitExitSnd SoundFX("ShooterLane",DOFContactors), ""
.CreateEvents "plungerIM"
End With
End Sub

'*********************************************************************
'* VUK ***************************************************************
'*********************************************************************

 Set bsL = New cvpmBallStack
 With bsL
 .InitSw 0, 36, 0, 0, 0, 0, 0, 0
 .InitKick sw36, 180, 0
 .InitExitSnd SoundFX("VUKOut",DOFContactors), ""
 .KickForceVar = 3
 End With

'*********************************************************************
'* SCOOP *************************************************************
'*********************************************************************

 Set bsR = New cvpmBallStack
 With bsR
 .InitSw 0, 37, 0, 0, 0, 0, 0, 0
 .InitKick SW37, 198, 24
 .KickZ = 0.4
 .InitExitSnd SoundFx("ScoopExit",DOFContactors), ""
 .KickAngleVar = 1
 .KickBalls = 1
 End With

'*********************************************************************
'* DROPTARGET ********************************************************
'*********************************************************************
 Set dtDrop = New cvpmDropTarget
 With dtDrop
 .InitDrop sw77, 77
 .initsnd "", ""
 End With

'*********************************************************************
'* NUDGING ***********************************************************
'*********************************************************************

vpmNudge.TiltSwitch = 14
vpmNudge.Sensitivity = 5.35
vpmNudge.TiltObj = Array(LeftBumper, BottomBumper, RightBumper, LeftSlingshot, RightSlingshot)

'*********************************************************************
'* UPDATE GI *********************************************************
'*********************************************************************

UpdateGI 0, 1:UpdateGI 1, 1:UpdateGI 2, 1

'*********************************************************************
'*********************************************************************
'*********************************************************************
'* END OF TABLE INIT *************************************************
'*********************************************************************
'*********************************************************************
'*********************************************************************

'*********************************************************************
'* CONTROLLER PAUSE **************************************************
'*********************************************************************

Sub AFM_Paused:Controller.Pause = 1:End Sub
Sub AFM_unPaused:Controller.Pause = 0:End Sub

'*********************************************************************
'* DRAIN HIT *********************************************************
'*********************************************************************

Sub Drain_Hit:PlaySoundAtVol "Drain",Drain,.8:bsTrough.AddBall Me:End Sub
Sub Drain1_Hit:PlaySoundAtVol "Drain",Drain,.8:bsTrough.AddBall Me:End Sub
Sub Drain2_Hit:PlaySoundAtVol "Drain",Drain,.8:bsTrough.AddBall Me:End Sub

'*********************************************************************
'* VUK HIT ***********************************************************
'*********************************************************************

Sub sw36a_Hit:PlaySoundAtVol "VUKEnter",sw36a,1:bsL.AddBall Me:End Sub

'*********************************************************************
'* SCOOP HIT *********************************************************
'*********************************************************************

Sub sw37a_Hit:PlaySoundAtVol "ScoopBack",sw37a,1:bsR.AddBall Me:End Sub

Sub sw37_Hit
PlaySoundAtVol "ScoopFront",sw37,1
Set bBall = ActiveBall:Me.TimerEnabled =1
bsR.AddBall 0
End Sub

Sub sw37_Timer
Do While bBall.Z >0
bBall.Z = bBall.Z -5
Exit Sub
Loop
Me.DestroyBall
Me.TimerEnabled = 0
End Sub

'*********************************************************************
'* SAUCER DROPHOLE HIT ***********************************************
'*********************************************************************

Sub sw78_Hit:PlaySoundAtVol "Drophole",sw78,1
Set aBall = ActiveBall
Me.TimerEnabled =1
vpmTimer.PulseSwitch(78), 150, "bsl.addball 0 '"
End Sub

Sub sw78_Timer
Do While aBall.Z >0
aBall.Z = aBall.Z -5
Exit Sub
Loop
Me.DestroyBall
Me.TimerEnabled = 0
End Sub

'*********************************************************************
'* SAUCER TARGET HIT *************************************************
'*********************************************************************

Sub sw77_Hit:dtDrop.Hit 1:End Sub

'*********************************************************************
'* KEYS **************************************************************
'*********************************************************************

Sub AFM_KeyDown(ByVal Keycode)
If keycode = PlungerKey Then Controller.Switch(11) = 1:End If
If keycode = CenterTiltKey Then Nudge 0, 5
If vpmKeyDown(keycode) Then Exit Sub
End Sub

Sub AFM_KeyUp(ByVal Keycode)
If keycode = PlungerKey Then Controller.Switch(11) = 0
If vpmKeyUp(keycode) Then Exit Sub
End Sub

'*********************************************************************
'* SOLENOID CALLS ****************************************************
'*********************************************************************

SolCallback(1) = "Auto_Plunger"
SolCallback(2) = "SolRelease"
SolCallback(3) = "bsL.SolOut"
SolCallback(4) = "bsR.SolOut"
SolCallBack(5) = "Alien1Move"
SolCallBack(6) = "Alien2Move"
SolCallBack(8) = "Alien3Move"
SolCallBack(14) = "Alien4Move"
SolCallBack(15) = "SolUfoShake"
SolCallback(16) = "dtDrop.SolDropUp"
SolCallback(17) = "setlamp 117,"
SolCallback(18) = "setlamp 118,"
SolCallBack(19) = "setlamp 119,"
SolCallBack(20) = "setlamp 120,"
SolCallback(21) = "setlamp 121,"
SolCallBack(23) = "setlamp 123,"
SolCallBack(24) = "TBMove"
SolCallBack(25) = "setlamp 125,"
SolCallBack(26) = "setlamp 126,"
SolCallBack(27) = "setlamp 127,"
SolCallBack(28) = "setlamp 128,"
SolCallback(33) = "vpmSolGate RGate,false,"
SolCallback(34) = "vpmSolGate LGate,false,"
SolCallback(36) = "DivMove"
SolCallback(43) = "setlamp 130,"
SolCallback(sLRFlipper) = "SolRFlipper"
SolCallback(sLLFlipper) = "SolLFlipper"

'*********************************************************************
'* FLIPPERS **********************************************************
'*********************************************************************

Sub SolLFlipper(Enabled)
If Enabled Then
PlaySoundAtVol SoundFX("FlipperUpRight",DOFContactors),LeftFlipper,3
LeftFlipper.RotateToEnd
Else
PlaySoundAtVol SoundFX("FlipperDownLeft",DOFContactors),LeftFlipper,.2
LeftFlipper.RotateToStart
lfstep = 1
LeftFlipper.TimerEnabled = 1
LeftFlipper.TimerInterval = 64
LeftFlipper.return = returnspeed * 0.5
End If
End Sub

Sub SolRFlipper(Enabled)
If Enabled Then
PlaySoundAtVol SoundFX("FlipperUpRight",DOFContactors),RightFlipper,2
RightFlipper.RotateToEnd
Else
PlaySoundAtVol SoundFX("FlipperDownRight",DOFContactors),RightFlipper,.2
RightFlipper.RotateToStart
rfstep = 1
RightFlipper.TimerEnabled = 1
RightFlipper.TimerInterval = 64
RightFlipper.return = returnspeed * 0.5
End If
End Sub

'===================
dim returnspeed, lfstep, rfstep
returnspeed = leftflipper.return
lfstep = 1
rfstep = 1

sub leftflipper_timer()
  select case lfstep
    Case 1: leftflipper.return = returnspeed * 0.6 :lfstep = lfstep + 1
    Case 2: leftflipper.return = returnspeed * 0.7 :lfstep = lfstep + 1
    Case 3: leftflipper.return = returnspeed * 0.8 :lfstep = lfstep + 1
    Case 4: leftflipper.return = returnspeed * 0.9 :lfstep = lfstep + 1
    Case 5: leftflipper.return = returnspeed * 1 :lfstep = lfstep + 1
    Case 6: leftflipper.timerenabled = 0 : lfstep = 1
  end select
end sub

sub rightflipper_timer()
  select case rfstep
    Case 1: rightflipper.return = returnspeed * 0.6 :rfstep = rfstep + 1
    Case 2: rightflipper.return = returnspeed * 0.7 :rfstep = rfstep + 1
    Case 3: rightflipper.return = returnspeed * 0.8 :rfstep = rfstep + 1
    Case 4: rightflipper.return = returnspeed * 0.9 :rfstep = rfstep + 1
    Case 5: rightflipper.return = returnspeed * 1 :rfstep = rfstep + 1
    Case 6: rightflipper.timerenabled = 0 : rfstep = 1
  end select
end sub

'*********************************************************************
'* SOLENOID SUBS *****************************************************
'*********************************************************************

Sub Auto_Plunger(Enabled)
If Enabled Then
PlungerIM.AutoFire
End If
End Sub

Sub SolRelease(Enabled)
If Enabled And bsTrough.Balls> 0 Then
vpmTimer.PulseSw 31
bsTrough.ExitSol_On
End If
End Sub

'*********************************************************************
'* POP BUMPERS *******************************************************
'*********************************************************************

Sub LeftBumper_Hit:vpmTimer.PulseSw 53:PlaySoundAtBumperVol "LeftBumper", LeftBumper,3
End Sub

Sub BottomBumper_Hit:vpmTimer.PulseSw 54:PlaySoundAtBumperVol "BottomBumper", BottomBumper,3
End Sub

Sub RightBumper_Hit:vpmTimer.PulseSw 55:PlaySoundAtBumperVol "RightBumper", RightBumper,3
End Sub

'*********************************************************************
'* SLINGSHOTS ********************************************************
'*********************************************************************

Dim LSlingStep
Sub LeftSlingShot_Slingshot
LeftSlingRubber.visible=0
LeftSlingRubber_A.visible=1
PlaySoundAtVol SoundFX("LeftSlingshot",DOFContactors),Licht6,3
vpmTimer.PulseSw 51
LSlingStep = 0
Me.TimerEnabled = 1
End Sub

Sub LeftSlingShot_Timer
Select Case LSlingStep
Case 2:LeftSlingRubber_A.visible = 0:LeftSlingRubber_B.visible = 1
Case 3:LeftSlingRubber_B.visible = 0:LeftSlingRubber_C.visible = 1
Case 4:LeftSlingRubber_C.visible = 0:LeftSlingRubber_B.visible = 1
Case 5:LeftSlingRubber_B.visible = 0:LeftSlingRubber_A.visible = 1
Case 6:LeftSlingRubber_A.visible = 0:LeftSlingRubber.visible = 1:Me.TimerEnabled = 0
End Select
LSlingStep = LSlingStep + 1
End Sub

Dim RSlingStep
Sub RightSlingShot_Slingshot
RightSlingRubber.visible=0
RightSlingRubber_A.visible=1
PlaySoundAtVol SoundFX("RightSlingshot",DOFContactors),Plastikslicht8,3
vpmTimer.PulseSw 52
RSlingStep = 0
Me.TimerEnabled = 1
End Sub

Sub RightSlingShot_Timer
Select Case RSlingStep
Case 2:RightSlingRubber_A.visible = 0:RightSlingRubber_B.visible = 1
Case 3:RightSlingRubber_B.visible = 0:RightSlingRubber_C.visible = 1
Case 4:RightSlingRubber_C.visible = 0:RightSlingRubber_B.visible = 1
Case 5:RightSlingRubber_B.visible = 0:RightSlingRubber_A.visible = 1
Case 6:RightSlingRubber_A.visible = 0:RightSlingRubber.visible = 1:Me.TimerEnabled = 0
End Select
RSlingStep = RSlingStep + 1
End Sub

'*********************************************************************
'* MARTIAN TARGETS ***************************************************
'*********************************************************************

'* "M"ARTIAN TARGET **************************************************
Sub SW56_Hit:vpmTimer.PulseSw 56:SW56P.X=102.375:SW56P.Y=1458.375:Me.TimerEnabled = 1:PlaySound SoundFX("TargetHit",DOFContactors),0,Vol(ActiveBall),Pan(ActiveBall),0,0,1,0,AudioFade(ActiveBall):End Sub
Sub SW56_Timer:SW56P.X=104.5:SW56P.Y=1458.75:Me.TimerEnabled = 0:End Sub

'* M"A"RTIAN TARGET **************************************************
Sub SW57_Hit:vpmTimer.PulseSw 57:SW57P.X=113.25:SW57P.Y=1399:Me.TimerEnabled = 1:PlaySound SoundFX("TargetHit",DOFContactors),0,Vol(ActiveBall),Pan(ActiveBall),0,0,1,0,AudioFade(ActiveBall):End Sub
Sub SW57_Timer:SW57P.X=115.5:SW57P.Y=1399.5:Me.TimerEnabled = 0:End Sub

'* MA"R"TIAN TARGET **************************************************
Sub SW58_Hit:vpmTimer.PulseSw 58:SW58P.X=124.375:SW58P.Y=1339.625:Me.TimerEnabled = 1:PlaySound SoundFX("TargetHit",DOFContactors),0,Vol(ActiveBall),Pan(ActiveBall),0,0,1,0,AudioFade(ActiveBall):End Sub
Sub SW58_Timer:SW58P.X=127.25:SW58P.Y=1340.125:Me.TimerEnabled = 0:End Sub

'* MAR"T"IAN TARGET **************************************************
Sub SW43_Hit:vpmTimer.PulseSw 43:SW43P.X=379.5:SW43P.Y=728.125:Me.TimerEnabled = 1:PlaySound SoundFX("TargetHit",DOFContactors),0,Vol(ActiveBall),Pan(ActiveBall),0,0,1,0,AudioFade(ActiveBall):End Sub
Sub SW43_Timer:SW43P.X=380:SW43P.Y=730.75:Me.TimerEnabled = 0:End Sub

'* MART"I"AN TARGET **************************************************
Sub SW44_Hit:vpmTimer.PulseSw 44:SW44P.X=634.5:SW44P.Y=730.625:Me.TimerEnabled = 1:PlaySound SoundFX("TargetHit",DOFContactors),0,Vol(ActiveBall),Pan(ActiveBall),0,0,1,0,AudioFade(ActiveBall):End Sub
Sub SW44_Timer:SW44P.X=634:SW44P.Y=733.125:Me.TimerEnabled = 0:End Sub

'* MARTI"A"N TARGET **************************************************
Sub SW41_Hit:vpmTimer.PulseSw 41:SW41P.X=833.125:SW41P.Y=1337.5:Me.TimerEnabled = 1:PlaySound SoundFX("TargetHit",DOFContactors),0,Vol(ActiveBall),Pan(ActiveBall),0,0,1,0,AudioFade(ActiveBall):End Sub
Sub SW41_Timer:SW41P.X=830.75:SW41P.Y=1337.875:Me.TimerEnabled = 0:End Sub

'* MARTIA"N" TARGET **************************************************
Sub SW42_Hit:vpmTimer.PulseSw 42:SW42P.X=839.625:SW42P.Y=1405:Me.TimerEnabled = 1:PlaySound SoundFX("TargetHit",DOFContactors),0,Vol(ActiveBall),Pan(ActiveBall),0,0,1,0,AudioFade(ActiveBall):End Sub
Sub SW42_Timer:SW42P.X=837.5:SW42P.Y=1405.375:Me.TimerEnabled = 0:End Sub

'*********************************************************************
'* SAUCER TARGETS ****************************************************
'*********************************************************************

Sub SW75_Hit:vpmTimer.PulseSw 75:SW75P.X=421.2813:SW75P.Y=554.4375:Me.TimerEnabled = 1:PlaySound SoundFX("TargetHit",DOFContactors),0,Vol(ActiveBall),Pan(ActiveBall),0,0,1,0,AudioFade(ActiveBall):End Sub
Sub SW75_Timer:SW75P.X=423.9063:SW75P.Y=555.6875:Me.TimerEnabled = 0:End Sub

Sub SW76_Hit:vpmTimer.PulseSw 76:SW76P.X=596.4063:SW76P.Y=553.1875:Me.TimerEnabled = 1:PlaySound SoundFX("TargetHit",DOFContactors),0,Vol(ActiveBall),Pan(ActiveBall),0,0,1,0,AudioFade(ActiveBall):End Sub
Sub SW76_Timer:SW76P.X=593.9063:SW76P.Y=554.1875:Me.TimerEnabled = 0:End Sub

'*********************************************************************
'* TARGETBANK TARGETS ************************************************
'*********************************************************************

Sub SW45_Hit
vpmTimer.PulseSw 45
SW45P.X=501.25
SW45P.Y=707.25
Me.TimerEnabled = 1
PlaySound SoundFX("TargetHit",DOFContactors),0,Vol(ActiveBall),Pan(ActiveBall),0,0,1,0,AudioFade(ActiveBall)
If Ballresting = True Then
 DPBall.VelY = ActiveBall.VelY * 3
End If
End Sub

Sub SW45_Timer:SW45P.X=501.125:SW45P.Y=709.25:Me.TimerEnabled = 0:End Sub

Sub SW46_Hit
vpmTimer.PulseSw 46
SW46P.X=506.875
SW46P.Y=707.25
Me.TimerEnabled = 1
PlaySound SoundFX("TargetHit",DOFContactors),0,Vol(ActiveBall),Pan(ActiveBall),0,0,1,0,AudioFade(ActiveBall)
If Ballresting = True Then
 DPBall.VelY = ActiveBall.VelY * 3
End If
End Sub

Sub SW46_Timer:SW46P.X=506.875:SW46P.Y=709.25:Me.TimerEnabled = 0:End Sub

Sub SW47_Hit
vpmTimer.PulseSw 47
SW47P.X=512.25
SW47P.Y=707.25
Me.TimerEnabled = 1
PlaySound SoundFX("TargetHit",DOFContactors),0,Vol(ActiveBall),Pan(ActiveBall),0,0,1,0,AudioFade(ActiveBall)
If Ballresting = True Then
 DPBall.VelY = ActiveBall.VelY * 3
End If
End Sub

Sub SW47_Timer:SW47P.X=512.25:SW47P.Y=709.25:Me.TimerEnabled = 0:End Sub

'*********************************************************************
'* RAMP SWITCHES *****************************************************
'*********************************************************************

Sub SW61_Hit:Controller.Switch(61) = 1:PlaySound "":End Sub
Sub SW61_UnHit:Controller.Switch(61) = 0:End Sub

Sub SW62_Hit:Controller.Switch(62) = 1:PlaySound "":End Sub
Sub SW62_UnHit:Controller.Switch(62) = 0:End Sub

Sub SW63_Hit:Controller.Switch(63) = 1:PlaySound "":End Sub
Sub SW63_UnHit:Controller.Switch(63) = 0:End Sub

Sub SW64_Hit:Controller.Switch(64) = 1:PlaySound "":End Sub
Sub SW64_UnHit:Controller.Switch(64) = 0:End Sub

Sub SW65_Hit:Controller.Switch(65) = 1:PlaySound "fx_metalrolling",0,Vol(ActiveBall)+2,Pan(ActiveBall),0,0,1,0,AudioFade(ActiveBall):End Sub
Sub SW65_UnHit:Controller.Switch(65) = 0:End Sub


'*********************************************************************
'* ROLLOVER SWITCHES *************************************************
'*********************************************************************

Sub SW16_Hit:Controller.Switch(16) = 1:PlaySound "rollover",0,Vol(ActiveBall)+2,Pan(ActiveBall),0,0,1,0,AudioFade(ActiveBall):End Sub
Sub SW16_UnHit:Controller.Switch(16) = 0:End Sub

Sub SW17_Hit:Controller.Switch(17) = 1:PlaySound "rollover",0,Vol(ActiveBall)+2,Pan(ActiveBall),0,0,1,0,AudioFade(ActiveBall):End Sub
Sub SW17_UnHit:Controller.Switch(17) = 0:End Sub

Sub SW26_Hit:Controller.Switch(26) = 1:PlaySound "rollover",0,Vol(ActiveBall)+2,Pan(ActiveBall),0,0,1,0,AudioFade(ActiveBall):End Sub
Sub SW26_UnHit:Controller.Switch(26) = 0:End Sub

Sub SW27_Hit:Controller.Switch(27) = 1:PlaySound "rollover",0,Vol(ActiveBall)+2,Pan(ActiveBall),0,0,1,0,AudioFade(ActiveBall):End Sub
Sub SW27_UnHit:Controller.Switch(27) = 0:End Sub


Sub SW38_Hit:Controller.Switch(38) = 1:PlaySound "rollover",0,Vol(ActiveBall)+1,Pan(ActiveBall),0,0,1,0,AudioFade(ActiveBall):End Sub
Sub SW38_UnHit:Controller.Switch(38) = 0:End Sub

Sub SW48_Hit:Controller.Switch(48) = 1:PlaySound "rollover",0,Vol(ActiveBall)+1,Pan(ActiveBall),0,0,1,0,AudioFade(ActiveBall):End Sub
Sub SW48_UnHit:Controller.Switch(48) = 0:End Sub

Sub SW71_Hit:Controller.Switch(71) = 1:PlaySound "rollover",0,Vol(ActiveBall)+1,Pan(ActiveBall),0,0,1,0,AudioFade(ActiveBall):End Sub
Sub SW71_UnHit:Controller.Switch(71) = 0:End Sub

Sub SW72_Hit:Controller.Switch(72) = 1:PlaySound "rollover",0,Vol(ActiveBall)+1,Pan(ActiveBall),0,0,1,0,AudioFade(ActiveBall):End Sub
Sub SW72_UnHit:Controller.Switch(72) = 0:End Sub

Sub SW73_Hit:Controller.Switch(73) = 1:PlaySound "rollover",0,Vol(ActiveBall)+1,Pan(ActiveBall),0,0,1,0,AudioFade(ActiveBall):End Sub
Sub SW73_UnHit:Controller.Switch(73) = 0:End Sub

Sub SW74_Hit:Controller.Switch(74) = 1:PlaySound "rollover",0,Vol(ActiveBall)+1,Pan(ActiveBall),0,0,1,0,AudioFade(ActiveBall):End Sub
Sub SW74_UnHit:Controller.Switch(74) = 0:End Sub

'*********************************************************************
'* SAUCER MOVEMENT ***************************************************
'*********************************************************************
Dim cBall
BigUfoInit

Sub BigUfoInit
Set cBall = ckicker.createball
ckicker.Kick 0, 0
End Sub

Sub SolUfoShake(Enabled)
If Enabled Then
BigUfoShake
PlaySound SoundFX("SaucerShake",DOFShaker),0,2,0,0,0,0,1,-.75
PlaySound "ShakerPulse",0,.5,0,0,0,0,0,-.01
End If
End Sub

Sub BigUfoShake
cball.velx = 1
cball.vely = -2
End Sub

Sub UfoShaker_Timer
Dim a, b, c
a = 90-(ckicker.y - cball.y)
b = (ckicker.y - cball.y) / 2
c = cball.x - ckicker.x

Ufo.rotx = a:Ufo.transz = b:Ufo.rotz = c
LED1.rotx = a:LED1.transz = b:LED1.rotz = c
LED2.rotx = a:LED2.transz = b:LED2.rotz = c
LED3.rotx = a:LED3.transz = b:LED3.rotz = c
LED4.rotx = a:LED4.transz = b:LED4.rotz = c
LED5.rotx = a:LED5.transz = b:LED5.rotz = c
LED6.rotx = a:LED6.transz = b:LED6.rotz = c
LED7.rotx = a:LED7.transz = b:LED7.rotz = c
LED8.rotx = a:LED8.transz = b:LED8.rotz = c
LED9.rotx = a:LED9.transz = b:LED9.rotz = c
LED10.rotx = a:LED10.transz = b:LED10.rotz = c
LED11.rotx = a:LED11.transz = b:LED11.rotz = c
LED12.rotx = a:LED12.transz = b:LED12.rotz = c
LED13.rotx = a:LED13.transz = b:LED13.rotz = c
LED14.rotx = a:LED14.transz = b:LED14.rotz = c
LED15.rotx = a:LED15.transz = b:LED15.rotz = c
LED16.rotx = a:LED16.transz = b:LED16.rotz = c
LEDGLOW1.rotx = a:LEDGLOW1.transz = b:LEDGLOW1.rotz = c
LEDGLOW2.rotx = a:LEDGLOW2.transz = b:LEDGLOW2.rotz = c
LEDGLOW3.rotx = a:LEDGLOW3.transz = b:LEDGLOW3.rotz = c
LEDGLOW4.rotx = a:LEDGLOW4.transz = b:LEDGLOW4.rotz = c
LEDGLOW5.rotx = a:LEDGLOW5.transz = b:LEDGLOW5.rotz = c
LEDGLOW6.rotx = a:LEDGLOW6.transz = b:LEDGLOW6.rotz = c
LEDGLOW7.rotx = a:LEDGLOW7.transz = b:LEDGLOW7.rotz = c
LEDGLOW8.rotx = a:LEDGLOW8.transz = b:LEDGLOW8.rotz = c
LEDGLOW9.rotx = a:LEDGLOW9.transz = b:LEDGLOW9.rotz = c
LEDGLOW10.rotx = a:LEDGLOW10.transz = b:LEDGLOW10.rotz = c
LEDGLOW11.rotx = a:LEDGLOW11.transz = b:LEDGLOW11.rotz = c
LEDGLOW12.rotx = a:LEDGLOW12.transz = b:LEDGLOW12.rotz = c
LEDGLOW13.rotx = a:LEDGLOW13.transz = b:LEDGLOW13.rotz = c
LEDGLOW14.rotx = a:LEDGLOW14.transz = b:LEDGLOW14.rotz = c
LEDGLOW15.rotx = a:LEDGLOW15.transz = b:LEDGLOW15.rotz = c
LEDGLOW16.rotx = a:LEDGLOW16.transz = b:LEDGLOW16.rotz = c
End Sub

'*********************************************************************
'* FLIPPER MODEL SYNC ************************************************
'*********************************************************************

Sub RightFlipperTimer_Timer:RightFlipperP.Rotz = RightFlipper.CurrentAngle:RightFlipperRubberP.Rotz = RightFlipper.CurrentAngle:End Sub

Sub LeftFlipperTimer_Timer:LeftFlipperP.Rotz = LeftFlipper.CurrentAngle -180:LeftFlipperRubberP.Rotz = LeftFlipper.CurrentAngle -180:End Sub

'*********************************************************************
'* ALIEN MOVEMENTS ***************************************************
'*********************************************************************

Dim Alien1Pos

Sub Alien1Move (enabled)
if enabled then
Alien1Timer.Enabled=1
PlaySoundAtVol SoundFX("AlienShake1",DOFShaker),Alien1n,2
'PlaySound "ShakerPulse",0,.002,0,0,0,0,1,-.01
End If
End Sub

Sub Alien1Timer_Timer()
Select Case Alien1Pos
Case 0: Alien1n.Z=60:AlienPost1.Z=-60
Case 1: Alien1n.Z=65:AlienPost1.Z=-55
Case 2: Alien1n.Z=70:AlienPost1.Z=-50
Case 3: Alien1n.Z=75:AlienPost1.Z=-45
Case 4: Alien1n.Z=80:AlienPost1.Z=-40
Case 5: Alien1n.Z=85:AlienPost1.Z=-35
Case 6: Alien1n.Z=90:AlienPost1.Z=-30
Case 7: Alien1n.Z=95:AlienPost1.Z=-25
Case 8: Alien1n.Z=100:AlienPost1.Z=-20
Case 9: Alien1n.Z=95:AlienPost1.Z=-25
Case 10: Alien1n.Z=90:AlienPost1.Z=-30
Case 11: Alien1n.Z=85:AlienPost1.Z=-35
Case 12: Alien1n.Z=80:AlienPost1.Z=-40
Case 13: Alien1n.Z=75:AlienPost1.Z=-45
Case 14: Alien1n.Z=70:AlienPost1.Z=-50
Case 15: Alien1n.Z=65:AlienPost1.Z=-55
Case 16: Alien1n.Z=60:AlienPost1.Z=-60:Alien1Pos=0:Alien1Timer.Enabled=0
End Select

If Alien1Pos=>0 then Alien1Pos=Alien1Pos+1
End Sub


Dim Alien2Pos

Sub Alien2Move (enabled)
if enabled then
Alien2Timer.Enabled=1
PlaySoundAtVol SoundFX("AlienShake2",DOFShaker),Alien2n,2
'PlaySound "ShakerPulse",0,.002,0,0,0,0,1,-.01
End If
End Sub

Sub Alien2Timer_Timer()
Select Case Alien2Pos
Case 0: Alien2n.Z=60:AlienPost2.Z=-60
Case 1: Alien2n.Z=65:AlienPost2.Z=-55:Alien2n.visible=0:Alien2h1.Z=65:Alien2h1.visible=1
Case 2: Alien2n.Z=70:AlienPost2.Z=-50
Case 3: Alien2n.Z=75:AlienPost2.Z=-45:Alien2h1.visible=0:Alien2h2.Z=75:Alien2h2.visible=1
Case 4: Alien2n.Z=80:AlienPost2.Z=-40
Case 5: Alien2n.Z=85:AlienPost2.Z=-35
Case 6: Alien2n.Z=90:AlienPost2.Z=-30:Alien2h2.visible=0:Alien2h1.Z=90:Alien2h1.visible=1
Case 7: Alien2n.Z=95:AlienPost2.Z=-25
Case 8: Alien2n.Z=100:AlienPost2.Z=-20:Alien2h1.visible=0:Alien2n.Z=100:Alien2n.visible=1
Case 9: Alien2n.Z=95:AlienPost2.Z=-25
Case 10: Alien2n.Z=90:AlienPost2.Z=-30:Alien2n.visible=0:Alien2r1.Z=90:Alien2r1.visible=1
Case 11: Alien2n.Z=85:AlienPost2.Z=-35
Case 12: Alien2n.Z=80:AlienPost2.Z=-40:Alien2r1.visible=0:Alien2r2.Z=70:Alien2r2.visible=1
Case 13: Alien2n.Z=75:AlienPost2.Z=-45
Case 14: Alien2n.Z=70:AlienPost2.Z=-50
Case 15: Alien2n.Z=65:AlienPost2.Z=-55:Alien2r2.visible=0:Alien2n.Z=60:Alien2n.visible=1:Alien2Pos=0:Alien2Timer.Enabled=0
End Select

If Alien2Pos=>0 then Alien2Pos=Alien2Pos+1
End Sub

Dim Alien3Pos

Sub Alien3Move (enabled)
if enabled then
Alien3Timer.Enabled=1
PlaySoundAtVol SoundFX("AlienShake3",DOFShaker),Alien3n,2
'PlaySound "ShakerPulse",0,.002,0,0,0,0,1,-.01
End If
End Sub

Sub Alien3Timer_Timer()
Select Case Alien3Pos
Case 0: Alien3n.Z=60:AlienPost3.Z=-60
Case 1: Alien3n.Z=65:AlienPost3.Z=-55:Alien3n.visible=0:Alien3h1.Z=65:Alien3h1.visible=1
Case 2: Alien3n.Z=70:AlienPost3.Z=-50
Case 3: Alien3n.Z=75:AlienPost3.Z=-45:Alien3h1.visible=0:Alien3h2.Z=75:Alien3h2.visible=1
Case 4: Alien3n.Z=80:AlienPost3.Z=-40
Case 5: Alien3n.Z=85:AlienPost3.Z=-35
Case 6: Alien3n.Z=90:AlienPost3.Z=-30:Alien3h2.visible=0:Alien3h1.Z=90:Alien3h1.visible=1
Case 7: Alien3n.Z=95:AlienPost3.Z=-25
Case 8: Alien3n.Z=100:AlienPost3.Z=-20:Alien3h1.visible=0:Alien3n.Z=100:Alien3n.visible=1
Case 9: Alien3n.Z=95:AlienPost3.Z=-25
Case 10: Alien3n.Z=90:AlienPost3.Z=-30:Alien3n.visible=0:Alien3r1.Z=90:Alien3r1.visible=1
Case 11: Alien3n.Z=85:AlienPost3.Z=-35
Case 12: Alien3n.Z=80:AlienPost3.Z=-40:Alien3r1.visible=0:Alien3r2.Z=70:Alien3r2.visible=1
Case 13: Alien3n.Z=75:AlienPost3.Z=-45
Case 14: Alien3n.Z=70:AlienPost3.Z=-50
Case 15: Alien3n.Z=65:AlienPost3.Z=-55:Alien3r2.visible=0:Alien3n.Z=60:Alien3n.visible=1:Alien3Pos=0:Alien3Timer.Enabled=0
End Select


If Alien3Pos=>0 then Alien3Pos=Alien3Pos+1
End Sub


Dim Alien4Pos

Sub Alien4Move (enabled)
if enabled then
Alien4Timer.Enabled=1
PlaySoundAtVol SoundFX("AlienShake4",DOFShaker),Alien4n,2
'PlaySound "ShakerPulse",0,.002,0,0,0,0,1,-.01
End If
End Sub

Sub Alien4Timer_Timer()
Select Case Alien4Pos
Case 0: Alien4n.Z=60:AlienPost4.Z=-60
Case 1: Alien4n.Z=65:AlienPost4.Z=-55:Alien4n.visible=0:Alien4h1.Z=65:Alien4h1.visible=1
Case 2: Alien4n.Z=70:AlienPost4.Z=-50
Case 3: Alien4n.Z=75:AlienPost4.Z=-45:Alien4h1.visible=0:Alien4h2.Z=75:Alien4h2.visible=1
Case 4: Alien4n.Z=80:AlienPost4.Z=-40
Case 5: Alien4n.Z=85:AlienPost4.Z=-35
Case 6: Alien4n.Z=90:AlienPost4.Z=-30:Alien4h2.visible=0:Alien4h1.Z=90:Alien4h1.visible=1
Case 7: Alien4n.Z=95:AlienPost4.Z=-25
Case 8: Alien4n.Z=100:AlienPost4.Z=-20:Alien4h1.visible=0:Alien4n.Z=100:Alien4n.visible=1
Case 9: Alien4n.Z=95:AlienPost4.Z=-25
Case 10: Alien4n.Z=90:AlienPost4.Z=-30:Alien4n.visible=0:Alien4r1.Z=90:Alien4r1.visible=1
Case 11: Alien4n.Z=85:AlienPost4.Z=-35
Case 12: Alien4n.Z=80:AlienPost4.Z=-40:Alien4r1.visible=0:Alien4r2.Z=70:Alien4r2.visible=1
Case 13: Alien4n.Z=75:AlienPost4.Z=-45
Case 14: Alien4n.Z=70:AlienPost4.Z=-50
Case 15: Alien4n.Z=65:AlienPost4.Z=-55:Alien4r2.visible=0:Alien4n.Z=60:Alien4n.visible=1:Alien4Pos=0:Alien4Timer.Enabled=0
End Select

If Alien4Pos=>0 then Alien4Pos=Alien4Pos+1
End Sub

'*********************************************************************
'* TARGETBANK MOVEMENT ***********************************************
'*********************************************************************

Dim TBPos, TBDown

Sub TBMove (enabled)
if enabled then
TBTimer.Enabled=1
PlaySoundAt SoundFX("TargetBank",DOFContactors),DPTrigger
End If
End Sub

Sub TBTimer_Timer()
Select Case TBPos
Case 0: MotorBank.Z=-20:SW45P.Z=-20:SW46P.Z=-20:SW47P.Z=-20:TBPos=0:TBDown=0:TBTimer.Enabled=0:Controller.Switch(66) = 0:Controller.Switch(67) = 1::SW45.isdropped=0:SW46.isdropped=0:SW47.isdropped=0:DPWall.isdropped=0:DPWall1.isdropped=1:DPRAMP.collidable=1
Case 1: MotorBank.Z=-22:SW45P.Z=-22:SW46P.Z=-22:SW47P.Z=-22
Case 2: MotorBank.Z=-24:SW45P.Z=-24:SW46P.Z=-24:SW47P.Z=-24
Case 3: MotorBank.Z=-26:SW45P.Z=-26:SW46P.Z=-26:SW47P.Z=-26
Case 4: MotorBank.Z=-28:SW45P.Z=-28:SW46P.Z=-28:SW47P.Z=-28
Case 5: MotorBank.Z=-30:SW45P.Z=-30:SW46P.Z=-30:SW47P.Z=-30
Case 6: MotorBank.Z=-32:SW45P.Z=-32:SW46P.Z=-32:SW47P.Z=-32
Case 7: MotorBank.Z=-34:SW45P.Z=-34:SW46P.Z=-34:SW47P.Z=-34
Case 8: MotorBank.Z=-36:SW45P.Z=-36:SW46P.Z=-36:SW47P.Z=-36
Case 9: MotorBank.Z=-38:SW45P.Z=-38:SW46P.Z=-38:SW47P.Z=-38
Case 10: MotorBank.Z=-40:SW45P.Z=-40:SW46P.Z=-40:SW47P.Z=-40
Case 11: MotorBank.Z=-42:SW45P.Z=-42:SW46P.Z=-42:SW47P.Z=-42
Case 12: MotorBank.Z=-44:SW45P.Z=-44:SW46P.Z=-44:SW47P.Z=-44:
Case 13: MotorBank.Z=-46:SW45P.Z=-46:SW46P.Z=-46:SW47P.Z=-46:
Case 14: MotorBank.Z=-48:SW45P.Z=-48:SW46P.Z=-48:SW47P.Z=-48
Case 15: MotorBank.Z=-50:SW45P.Z=-50:SW46P.Z=-50:SW47P.Z=-50
Case 16: MotorBank.Z=-52:SW45P.Z=-52:SW46P.Z=-52:SW47P.Z=-52
Case 17: MotorBank.Z=-54:SW45P.Z=-54:SW46P.Z=-54:SW47P.Z=-54
Case 18: MotorBank.Z=-56:SW45P.Z=-56:SW46P.Z=-56:SW47P.Z=-56
Case 19: MotorBank.Z=-58:SW45P.Z=-58:SW46P.Z=-58:SW47P.Z=-58
Case 20: MotorBank.Z=-60:SW45P.Z=-60:SW46P.Z=-60:SW47P.Z=-60
Case 21: MotorBank.Z=-62:SW45P.Z=-62:SW46P.Z=-62:SW47P.Z=-62
Case 22: MotorBank.Z=-64:SW45P.Z=-64:SW46P.Z=-64:SW47P.Z=-64
Case 23: MotorBank.Z=-66:SW45P.Z=-66:SW46P.Z=-66:SW47P.Z=-66
Case 24: MotorBank.Z=-68:SW45P.Z=-68:SW46P.Z=-68:SW47P.Z=-68
Case 25: MotorBank.Z=-70:SW45P.Z=-70:SW46P.Z=-70:SW47P.Z=-70
Case 26: MotorBank.Z=-72:SW45P.Z=-72:SW46P.Z=-72:SW47P.Z=-72
Case 27: MotorBank.Z=-74:SW45P.Z=-74:SW46P.Z=-74:SW47P.Z=-74
Case 28: MotorBank.Z=-76:SW45P.Z=-76:SW46P.Z=-76:SW47P.Z=-76:SW45.isdropped=1:SW46.isdropped=1:SW47.isdropped=1:DPWALL.isdropped=1:DPRAMP.collidable=0
Case 29: TBTimer.Enabled=0:TBDown=1:Controller.Switch(66) = 1:Controller.Switch(67) = 0
End Select

If TBDown=0 then TBPos=TBPos+1
If TBDown=1 then TBPos=TBPos-1
End Sub

'*********************************************************************
'* DIRTY POOL HANDLE *************************************************
'*********************************************************************

Dim DPBall, Ballresting

DPBall = ""
Ballresting = False

Sub DPTrigger_Hit
Ballresting = True
Set DPBall = ActiveBall
End Sub

Sub DPTrigger_UnHit
Ballresting = False
DPRAMP.collidable=0
End Sub

'*********************************************************************
'* DIVERTER MOVEMENT *************************************************
'*********************************************************************

Dim DivPos, DivClosed, DiverterDir

Sub DivMove(Enabled)
If Enabled Then
PlaySound SoundFX("Diverter",DOFContactors)
DiverterDir = 1
Else
DiverterDir = -1
End If

DivTimer.Enabled = 0
If DivPos <1 Then DivPos = 1
If DivPos > 8 Then DivPos = 8

DivTimer.Enabled = 1
End Sub

Sub DivTimer_Timer()
Select Case DivPos
Case 0:Diverter.ObjRotZ=-15:Diverter9.IsDropped = 0:Diverter8.IsDropped = 1:DivTimer.Enabled = 0
Case 1:Diverter.ObjRotZ=-8:Diverter8.IsDropped = 0:Diverter9.IsDropped = 1:Diverter7.IsDropped = 1
Case 2:Diverter.ObjRotZ=-3:Diverter7.IsDropped = 0:Diverter8.IsDropped = 1:Diverter6.IsDropped = 1
Case 3:Diverter.ObjRotZ=2:Diverter6.IsDropped = 0:Diverter7.IsDropped = 1:Diverter5.IsDropped = 1
Case 4:Diverter.ObjRotZ=7:Diverter5.IsDropped = 0:Diverter6.IsDropped = 1:Diverter4.IsDropped = 1
Case 5:Diverter.ObjRotZ=12:Diverter4.IsDropped = 0:Diverter5.IsDropped = 1:Diverter3.IsDropped = 1
Case 6:Diverter.ObjRotZ=17:Diverter3.IsDropped = 0:Diverter4.IsDropped = 1:Diverter2.IsDropped = 1
Case 7:Diverter.ObjRotZ=22:Diverter2.IsDropped = 0:Diverter3.IsDropped = 1:Diverter1.IsDropped = 1
Case 8:Diverter.ObjRotZ=27:Diverter1.IsDropped = 0:Diverter2.isdropped = 0
Case 9:DivTimer.Enabled = 0
End Select

DivPos = DivPos + DiverterDir
End Sub

'*********************************************************************
'* OBJECT SOUND EFFECTS **********************************************
'*********************************************************************

Sub LeftPost_Hit(): PlaySoundAt "RubberHit",ActiveBall: End Sub

Sub LeftSlingRubber_Hit(): PlaySoundAt "RubberHit",ActiveBall:End Sub

Sub RightPost_Hit(): PlaySoundAt "RubberHit",ActiveBall: End Sub

Sub RightSlingRubber_Hit(): PlaySoundAt "RubberHit",ActiveBall:End Sub

Sub Rubber3_Hit(): PlaySoundAt "RubberHit3",ActiveBall: End Sub

Sub Rubber6_Hit(): PlaySoundAt "RubberHit3",ActiveBall: End Sub

Sub Wall14_Hit(): PlaySoundAt "Objecthit",ActiveBall:End Sub

Sub Wall31_Hit(): PlaySoundAt "Objecthit",ActiveBall:End Sub

Sub Wall29_Hit(): PlaySoundAt "Objecthit",ActiveBall:End Sub

Sub Wall34_Hit(): PlaySoundAt "Objecthit",ActiveBall:End Sub




Sub LeftFlipper_Collide(parm):
Ballspeed=SQR(activeball.velx * activeball.velx + activeball.vely * activeball.vely)
If Ballspeed >2 then Playsound "RubberHit2" Else Playsound "RubberHitLow":End If:End Sub

Sub RightFlipper_Collide(parm):
Ballspeed=SQR(activeball.velx * activeball.velx + activeball.vely * activeball.vely)
If Ballspeed >2 then Playsound "RubberHit2" Else Playsound "RubberHitLow":End If:End Sub

Sub RDrop_Hit()
ActiveBall.VelZ = -2
ActiveBall.VelY = 0
ActiveBall.VelX = 0
StopSound "fx_metalrolling"
PlaySound "Balldrop",0,1,.4,0,0,0,1,.8
End Sub

Sub LDrop_Hit()
ActiveBall.VelZ = -2
ActiveBall.VelY = 0
ActiveBall.VelX = 0
StopSound "fx_metalrolling"
PlaySound "Balldrop",0,1,-.4,0,0,0,1,.8
End Sub

Sub LRoll_Hit()
PlaySound "fx_metalrolling",0,Vol(ActiveBall)+2,Pan(ActiveBall),0,0,1,0,AudioFade(ActiveBall)
End Sub

'*********************************************************************
'* BALL SOUND FUNCTIONS **********************************************
'*********************************************************************

Function Vol(ball) ' Calculates the Volume of the sound based on the ball speed
Vol = Csng(BallVel(ball) ^2 / 1000)
End Function

Function Pan(ball) ' Calculates the pan for a ball based on the X position on the table. "AFM" is the name of the table
Dim tmp
tmp = ball.x * 2 / AFM.width-1
If tmp > 0 Then
Pan = Csng(tmp ^10)
Else
Pan = Csng(-((- tmp) ^10) )
End If
End Function

function AudioFade(ball)
  Dim tmp
    tmp = ball.y * 2 / AFM.height-1
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

'*********************************************************************
'* JP VPX ROLLING SOUNDS *********************************************
'*********************************************************************

Const tnob = 5 ' total number of balls in this table is 4, but always use a higher number here because of the timing
ReDim rolling(tnob)
InitRolling

Sub InitRolling
Dim i
For i = 0 to tnob
rolling(i) = False
Next
End Sub

Sub RollingUpdate()
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
' in this table we ignore the captive ball, it´s the 0
For b = 1 to UBound(BOT)
If BallVel(BOT(b) ) > 1 AND BOT(b).z < 30 Then
rolling(b) = True
PlaySound("fx_ballrolling" & b), -1, Vol(BOT(b) ), Pan(BOT(b) ), 0, Pitch(BOT(b) ), 1, 0, AudioFade(BOT(b))
Else
If rolling(b) = True Then
StopSound("fx_ballrolling" & b)
rolling(b) = False
End If
End If
Next
End Sub

'*********************************************************************
'* BALL COLLISION SOUNDS *********************************************
'*********************************************************************

Sub OnBallBallCollision(ball1, ball2, velocity)
PlaySound("fx_collide"), 0, Csng(velocity) ^2 / 200, Pan(ball1), 0, Pitch(ball1), 0, 0, AudioFade(ball1)
End Sub

'**************************************************************************
'                 Positional Sound Playback Functions by DJRobX
'**************************************************************************

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
'* UPDATE GI ILLUMINATION ********************************************
'*********************************************************************

Sub UpdateGI(no, step)
Dim gistep, ii, a
gistep = step / 8
Select Case no
Case 0
For each ii in GIBottom
ii.IntensityScale = gistep
Next
Case 1
For each ii in GIMiddle
ii.IntensityScale = gistep
Next
Case 2
For each ii in GITop
ii.IntensityScale = gistep
Next
End Select
End Sub

'*********************************************************************
'* REALTIME UPDATES **************************************************
'*********************************************************************

Set MotorCallback = GetRef("RealTimeUpdates")

Sub RealTimeUpdates
RollingUpdate
End Sub

'*********************************************************************
'* JP FADING LIGHT SYSTEM ********************************************
'*********************************************************************

Dim LampState(200), FadingLevel(200)
Dim FlashSpeedUp(200), FlashSpeedDown(200), FlashMin(200), FlashMax(200), FlashLevel(200)

InitLamps()             ' turn off the lights and flashers and reset them to the default parameters
LampTimer.Interval = 160 'lamp fading speed
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

'*********************************************************************
'* INSERT LIGHTS *****************************************************
'*********************************************************************

Sub UpdateLamps

NFadeL 11, l11
NFadeL 12, l12
NFadeL 13, l13
NFadeL 14, l14
NFadeL 15, l15
NFadeL 16, l16
NFadeL 17, l17
NFadeL 18, l18
NFadeL 21, l21
NFadeL 22, l22
NFadeL 23, l23
NFadeL 24, l24
NFadeL 25, l25
NFadeL 26, l26
NFadeL 27, l27
NFadeL 28, l28
NFadeL 31, l31
NFadeL 32, l32
NFadeL 33, l33
NFadeL 34, l34
NFadeL 35, l35
NFadeL 36, l36
NFadeL 37, l37
NFadeL 38, l38
NFadeL 41, l41
NFadeL 42, l42
NFadeL 43, l43
NFadeL 44, l44
NFadeL 45, l45
NFadeL 46, l46
NFadeL 47, l47
NFadeL 48, l48
NFadeL 51, l51
NFadeL 52, l52
NFadeL 53, l53
NFadeL 54, l54
NFadeL 55, l55
NFadeL 56, l56
NFadeL 57, l57
NFadeL 58, l58
NFadeL 61, l61
NFadeL 62, l62
NFadeL 63, l63
NFadeL 64, l64
NFadeL 65, l65
NFadeL 66, l66
NFadeL 67, l67
NFadeL 68, l68
NFadeL 71, l71
NFadeL 72, l72
NFadeL 73, l73
NFadeL 74, l74
NFadeL 75, l75
NFadeL 76, l76
NFadeL 77, l77
NFadeL 78, l78
NFadeL 81, l81
NFadeL 82, l82
NFadeL 83, l83
NFadeL 84, l84
NFadeL 85, l85
'NFadeL 86, l86
'NFadeL 88, l88

'*********************************************************************
'* FLASHER ***********************************************************
'*********************************************************************

NFadeLm 117, f17
NFadeLm 118, F18
NFadeLm 119, F19
NFadeL 120, f20
NFadeLm 121, F21a
NFadeL 121, F21
NFadeLm 123, f23
NFadeLm 125, F25
NFadeLm 126, F26
NFadeLm 127, F27
NFadeL 128, f28
NFadeLm 130, f30b
NFadeL 130, f30a

'*********************************************************************
'* SAUCER LED ********************************************************
'*********************************************************************

FadeObjm 91, LEDGLOW1, "LEDGLOW_3", "LEDGLOW_2", "LEDGLOW_1", "LEDGLOW_OFF"
FadeObjm 92, LEDGLOW2, "LEDGLOW_3", "LEDGLOW_2", "LEDGLOW_1", "LEDGLOW_OFF"
FadeObjm 93, LEDGLOW3, "LEDGLOW_3", "LEDGLOW_2", "LEDGLOW_1", "LEDGLOW_OFF"
FadeObjm 94, LEDGLOW4, "LEDGLOW_3", "LEDGLOW_2", "LEDGLOW_1", "LEDGLOW_OFF"
FadeObjm 96, LEDGLOW6, "LEDGLOW_3", "LEDGLOW_2", "LEDGLOW_1", "LEDGLOW_OFF"
FadeObjm 97, LEDGLOW7, "LEDGLOW_3", "LEDGLOW_2", "LEDGLOW_1", "LEDGLOW_OFF"
FadeObjm 98, LEDGLOW8, "LEDGLOW_3", "LEDGLOW_2", "LEDGLOW_1", "LEDGLOW_OFF"
FadeObjm 101, LEDGLOW9, "LEDGLOW_3", "LEDGLOW_2", "LEDGLOW_1", "LEDGLOW_OFF"
FadeObjm 102, LEDGLOW10, "LEDGLOW_3", "LEDGLOW_2", "LEDGLOW_1", "LEDGLOW_OFF"
FadeObjm 103, LEDGLOW11, "LEDGLOW_3", "LEDGLOW_2", "LEDGLOW_1", "LEDGLOW_OFF"
FadeObjm 104, LEDGLOW12, "LEDGLOW_3", "LEDGLOW_2", "LEDGLOW_1", "LEDGLOW_OFF"
FadeObjm 105, LEDGLOW13, "LEDGLOW_3", "LEDGLOW_2", "LEDGLOW_1", "LEDGLOW_OFF"
FadeObjm 106, LEDGLOW14, "LEDGLOW_3", "LEDGLOW_2", "LEDGLOW_1", "LEDGLOW_OFF"
FadeObjm 107, LEDGLOW15, "LEDGLOW_3", "LEDGLOW_2", "LEDGLOW_1", "LEDGLOW_OFF"
FadeObjm 108, LEDGLOW16, "LEDGLOW_3", "LEDGLOW_2", "LEDGLOW_1", "LEDGLOW_OFF"

FadeObj 91, Led1, "LED_3", "LED_2", "LED_1", "LED_OFF"
FadeObj 92, Led2, "LED_3", "LED_2", "LED_1", "LED_OFF"
FadeObj 93, Led3, "LED_3", "LED_2", "LED_1", "LED_OFF"
FadeObj 94, Led4, "LED_3", "LED_2", "LED_1", "LED_OFF"
FadeObj 95, Led5, "LED_3", "LED_2", "LED_1", "LED_OFF"
FadeObj 96, Led6, "LED_3", "LED_2", "LED_1", "LED_OFF"
FadeObj 97, Led7, "LED_3", "LED_2", "LED_1", "LED_OFF"
FadeObj 98, Led8, "LED_3", "LED_2", "LED_1", "LED_OFF"
FadeObj 101, Led9, "LED_3", "LED_2", "LED_1", "LED_OFF"
FadeObj 102, Led10, "LED_3", "LED_2", "LED_1", "LED_OFF"
FadeObj 103, Led11, "LED_3", "LED_2", "LED_1", "LED_OFF"
FadeObj 104, Led12, "LED_3", "LED_2", "LED_1", "LED_OFF"
FadeObj 105, Led13, "LED_3", "LED_2", "LED_1", "LED_OFF"
FadeObj 106, Led14, "LED_3", "LED_2", "LED_1", "LED_OFF"
FadeObj 107, Led15, "LED_3", "LED_2", "LED_1", "LED_OFF"
FadeObj 108, Led16, "LED_3", "LED_2", "LED_1", "LED_OFF"

'*********************************************************************
'* SAUCER FLASHER ****************************************************
'*********************************************************************

FadeObj 123, Ufo, "Mothership_3", "Mothership_2", "Mothership_1", "Mothership"

'*********************************************************************
'* SMALL SAUCER FLASHER **********************************************
'*********************************************************************

FadeObj 117, F17P, "Saucer_3", "Saucer_2", "Saucer_1", "Saucer"
FadeObj 118, F18P, "Saucer_3", "Saucer_2", "Saucer_1", "Saucer"
FadeObj 119, F19P, "Saucer_3", "Saucer_2", "Saucer_1", "Saucer"
FadeObj 125, F25P, "Saucer_3", "Saucer_2", "Saucer_1", "Saucer"
FadeObj 126, F26P, "Saucer_3", "Saucer_2", "Saucer_1", "Saucer"
FadeObj 127, F27P, "Saucer_3", "Saucer_2", "Saucer_1", "Saucer"

End Sub

'*********************************************************************
'* LAMP SUBS *********************************************************
'*********************************************************************

Sub InitLamps()
Dim x
For x = 0 to 200
LampState(x) = 0         ' current light state, independent of the fading level. 0 is off and 1 is on
FadingLevel(x) = 4       ' used to track the fading state
FlashSpeedUp(x) = 0.2    ' faster speed when turning on the flasher
FlashSpeedDown(x) = 0.1  ' slower speed when turning off the flasher
FlashMax(x) = 1          ' the maximum value when on, usually 1
FlashMin(x) = 0          ' the minimum value when off, usually 0
FlashLevel(x) = 0        ' the intensity of the flashers, usually from 0 to 1
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

'*********************************************************************
'* VPX STANDARD LIGHTS ***********************************************
'*********************************************************************

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

'*********************************************************************
'* VPX RAMP & PRIMITIVE LIGHTS ***************************************
'*********************************************************************

Sub FadeObj(nr, object, a, b, c, d)
Select Case FadingLevel(nr)
Case 4:object.image = b:FadingLevel(nr) = 6                   'fading to off...
Case 5:object.image = a:FadingLevel(nr) = 1                   'ON
Case 6, 7, 8:FadingLevel(nr) = FadingLevel(nr) + 1            'wait
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

'*********************************************************************
'* VPX FLASHER OBJECTS ***********************************************
'*********************************************************************

Sub SetFlash(nr, stat)
FadingLevel(nr) = ABS(stat)
End Sub

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
'FadingLevel(nr) = 1 'completely on
End if
Object.IntensityScale = FlashLevel(nr)
End Select
End Sub

Sub Flashm(nr, object) 'multiple flashers, it just sets the flashlevel
Object.IntensityScale = FlashLevel(nr)
End Sub

'*********************************************************************
'* VPX REELS & TEXT  *************************************************
'*********************************************************************

Sub FadeR(nr, object)
Select Case FadingLevel(nr)
Case 4:object.SetValue 1:FadingLevel(nr) = 6                   'fading to off...
Case 5:object.SetValue 0:FadingLevel(nr) = 1                   'ON
Case 6, 7, 8:FadingLevel(nr) = FadingLevel(nr) + 1             'wait
Case 9:object.SetValue 2:FadingLevel(nr) = FadingLevel(nr) + 1 'fading...
Case 10, 11, 12:FadingLevel(nr) = FadingLevel(nr) + 1          'wait
Case 13:object.SetValue 3:FadingLevel(nr) = 0                  'Off
End Select
End Sub

Sub FadeRm(nr, object)
Select Case FadingLevel(nr)
Case 4:object.SetValue 1
Case 5:object.SetValue 0
Case 9:object.SetValue 2
Case 3:object.SetValue 3
End Select
End Sub

Sub NFadeT(nr, object, message)
Select Case FadingLevel(nr)
Case 4:object.Text = "":FadingLevel(nr) = 0
Case 5:object.Text = message:FadingLevel(nr) = 1
End Select
End Sub

Sub NFadeTm(nr, object, b)
Select Case FadingLevel(nr)
Case 4:object.Text = ""
Case 5:object.Text = message
End Select
End Sub

'*********************************************************************
'*********************************************************************
'*********************************************************************
'* END OF TABLE SCRIPT ***********************************************
'*********************************************************************
'*********************************************************************
'*********************************************************************
' Thalamus : Exit in a clean and proper way
Sub AFM_exit()
  Controller.Pause = False
  Controller.Stop
End Sub


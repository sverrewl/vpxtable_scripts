Option Explicit
 Randomize

' Thalamus 2018-07-24
' Tables has its own "Positional Sound Playback Functions" and "Supporting Ball & Sound Functions"
' Supports cfastflips but video mode is uncertain
' No special SSF tweaks yet.

 Dim Ballsize,BallMass
 BallSize = 50
 BallMass = (Ballsize^3)/125000

 On Error Resume Next
 ExecuteGlobal GetTextFile("controller.vbs")
 If Err Then MsgBox "You need the controller.vbs in order to run this table, available in the vp10 package"
 On Error Goto 0

 Dim DesktopMode:DesktopMode = Table1.ShowDT
 Dim UseVPMDMD:UseVPMDMD = DesktopMode
 Const UseVPMModSol = 1


 LoadVPM "01560000", "WPC.VBS", 3.50

 Const cGameName = "ts_lx5"

 'STANDARD DEFINITIONS *************************************************************************************************************************************

 Const UseSolenoids = 1
 Const UseLamps = 0
 Const UseSync = 0
 Const HandleMech = 1
 Set MotorCallback = GetRef("HandleMiniPF")


'DIMS *************************************************************************************************************************************

 Dim bsTrough, bsL, bsR, dtDrop, x, bump1, bump2, bump3, BallFrame, Mech3bank, bsLeftEject, bsRightEject, bsCenterEject, BallInMagnet, ReleaseMagnetBall,_
 KickerPos, NewKickerPos, LastKickerPos, bslock, Fade, plungerim, LeftNudgeEffect, RightNudgeEffect, NudgeEffect

'STANDARD SOUNDS ******************************************************************************************************************************************

 Const SSolenoidOn = "Solenoid"
 Const SSolenoidOff = ""
 Const SFlipperOn = "FlipperUp"
 Const SFlipperOff = "FlipperDown"
 Const SCoin = "Coin"

'KEY DEFINITIONS ******************************************************************************************************************************************

 Const keyBuyInButton = 3
 Const keyLeftPhurbaButton = 30
 Const keyRightPhurbaButton = 40

 Sub Table1_Init
    vpmInit Me
    With Controller
        .GameName = cGameName
        If Err Then MsgBox "Can't start Game" & cGameName & vbNewLine & Err.Description:Exit Sub
        .SplashInfoLine = "The Shadow (Bally 1994) By Alessio"
        .HandleKeyboard = 0
        .ShowTitle = 0
        .ShowDMDOnly = 1
        .ShowFrame = 0
        .HandleMechanics = 1
        .Hidden = DesktopMode
        On Error Resume Next
        .Run GetPlayerHWnd
        If Err Then MsgBox Err.Description
    End With

    On Error Goto 0

 'MAIN TIMER INIT ******************************************************************************************************************************************

 PinMAMETimer.Interval = PinMAMEInterval
 PinMAMETimer.Enabled = 1

 'StartShake



 'NUDGE INIT *******************************************************************************************************************************

 vpmNudge.TiltSwitch = 14
 vpmNudge.Sensitivity = 0.5
 vpmNudge.TiltObj=Array(LeftSlingshot,RightSlingshot)



'DROP TARGETS INIT ****************************************************************************************************************************************


 sw55.isDropped = 0:sw85.IsDropped = 0:sw86.IsDropped = 0:_
 sw87.IsDropped = 0:sw88.IsDropped = 0:WallTarget.isDropped = 0

'BATTLEFIELD MINI KICKER INIT *****************************************************************************************************************************

 Init_MiniKicker

'PHURBA DIVERTERS INIT ************************************************************************************************************************************

 'RPLa.isdropped = 0:RPRa.isdropped = 1:_
' LPL.collidable = 1:LPLa.collidable = 1
 'LPR.collidable = 0:LPRa.collidable = 0
 'RightPhurbaRight.isdropped = 1:RightPhurbaLeft.isdropped = 0:_
 'LeftPhurbaLeft.isdropped = 0:LeftPhurbaRight.isdropped = 1

 'BALLS INIT ***********************************************************************************************************************************************


 Set bsTrough = New cvpmBallStack:With bsTrough
 .InitSw 0, 41, 42, 43, 44, 45, 0, 0
 .InitKick BallRelease, 90, 10
 .InitEntrySnd "Solenoid", "Solenoid"
 .InitExitSnd SoundFX("ballrel",DOFContactors), SoundFX("Solenoid",DOFContactors)
 .Balls = 5
 End With

  Controller.Switch(22) = 1

  'End With
  End Sub


 Sub table1_KeyDown(ByVal keycode)

 If keycode = PlungerKey  Then Controller.Switch(11) = True
 If keycode = keyBuyInButton Then Controller.Switch(23) = True
 If keycode = LeftMagnaSave Then Controller.Switch(34) = True
 If keycode = RightMagnaSave Then Controller.Switch(12) = True
 If keycode = LeftTiltKey Then vpmNudge.DoNudge 80, 1.6
 If keycode = RightTiltKey Then vpmNudge.DoNudge 280, 1.6
 If keycode = CenterTiltKey Then vpmNudge.DoNudge 0, 1.6
  if keydownhandler(keycode) then exit sub
 End Sub

 Sub table1_KeyUp(ByVal keycode)

 If keycode = PlungerKey  Then Controller.Switch(11) = False
 If keycode = keyBuyInButton Then Controller.Switch(23) = False
 If keycode = LeftMagnaSave Then Controller.Switch(34) = False
 If keycode = RightMagnaSave Then Controller.Switch(12) = False
 If keyuphandler(keycode) then exit sub
 End Sub

 'FLIPPERS *************************************************************************************************************************************************

 SolCallback(sLRFlipper) = "SolRFlipper"
 SolCallback(sLLFlipper) = "SolLFlipper"

 Sub SolLFlipper(Enabled)
  vpmSolFlipper LeftFlipper, Nothing, Enabled
  If Enabled then
  Playsound SoundFX("FlipperSu",DOFFlippers)
  End If
  If Not Enabled then
  Playsound SoundFX("FlipperGiu",DOFFlippers)
  End If
 End Sub

 Sub SolRFlipper(Enabled)
 vpmSolFlipper RightFlipper, RightFlipper2, Enabled
  If Enabled then
  Playsound SoundFX("FlipperSu",DOFFlippers)
  End If
  If Not Enabled then
  Playsound SoundFX("FlipperGiu",DOFFlippers)
  End If
 End Sub

'BALL SOUNDS ************************************************************************************************************************************

 Sub RightRampSound_Hit
  Playsound "Metalrolling"
 End Sub

 Sub BallRoll_Hit
  Playsound "Ballrolling"
 End Sub

 Sub BallRoll2_Hit
  Playsound "Ballrolling"
 End Sub

 Sub Skill_UnHit
 Playsound "OnlyMetalRolling"
 End Sub

'AUTOMATIC PLUNGER ****************************************************************************************************************************************

 Const IMPowerSetting = 55   'ERA 40
 Const IMTime = 0.6
 Set plungerIM = New cvpmImpulseP
 With plungerIM
 .InitImpulseP swplunger, IMPowerSetting, IMTime
 .Random 0.3
 .InitExitSnd "plunger2", "plunger"
 .CreateEvents "plungerIM"
 End With

 Sub Auto_Plunger(Enabled)
 If Enabled Then
 PlungerIM.AutoFire
 Playsound SoundFX("plunger2",DOFContactors)
 End If
 End Sub

'SOLENOIDS ************************************************************************************************************************************************

 SolCallback(1)  = "Auto_Plunger"
 SolCallback(2)  = "SolLockupKickout"
 SolCallback(3)  = "SolLeftPhurbaLeft"
 SolCallback(4)  = "SolLeftPhurbaRight"
 SolCallback(5)  = "SolRightPhurbaRight"
 SolCallback(6)  = "SolRightPhurbaLeft"
 SolCallback(7)  = "SolKnocker"
 SolCallback(8)  = "SolWallTargetUp"
 SolCallback(9)  = "solLSling"
 SolCallback(10) = "solRSling"
 SolCallback(11) = "bsRightEject.SolOut"
 SolCallback(12) = "bsLeftEject.SolOut"
 SolCallback(13) = "SolBallRelease"
 Solcallback(14) = "SolBallPopper"
 SolCallback(16) = "SolWallTargetDown"
 Solcallback(24) = "SolMiniDropbank"
 SolCallback(25) = "SolSingleDropUp"
 SolCallback(35) = "SolMagnetOn"
 SolCallback(36) = "SolSingleDropDown"

'FLASHER SOLENOIDS FOR FADING LIGHT SYSTEM ****************************************************************************************************************

 SolCallback(17) = "SetLamp 117,"
 SolCallback(18) = "SetLamp 118,"
 SolCallback(21) = "SetLamp 121,"
 SolCallback(22) = "SetLamp 122,"
 SolCallback(23) = "SetLamp 123,"
 SolCallback(26) = "SetLamp 126,"
 SolCallback(27) = "SetLamp 127,"
 SolCallback(28) = "SetLamp 128,"

 Sub SolKnocker(enabled)
	If enabled then
		PlaySound SoundFX("Knocker",DOFKnocker)
    End If
 End Sub

 Sub shooterlane_Hit():Controller.Switch(48) = 1:End Sub:Sub shooterlane_Unhit():Controller.Switch(48) = 0:End Sub

 Sub SolBallRelease(enabled)
 If enabled Then
 If bsTrough.Balls Then PlaySound SoundFX("BallRelease",DOFContactors) Else PlaySound SoundFX("Solenoid",DOFContactors)
 bsTrough.ExitSol_On
 End If
 End Sub

  Sub table1_Paused:Controller.Pause = 1:End Sub
 Sub table1_unPaused:Controller.Pause = 0:End Sub


'DRAIN HOLES **********************************************************************************************************************************************

 Sub Drain_Hit:Playsound "drain":bsTrough.AddBall Me:End Sub
 Sub Drain_Hit():PlaySound "Drain":bsTrough.AddBall Me:End Sub
 'Sub Drain2_Hit():PlaySound "Drain":bsTrough.AddBall Me:End Sub
 'Sub Drain3_Hit():PlaySound "Drain":bsTrough.AddBall Me:End Sub:Sub Drain4_Hit():PlaySound "Drain":bsTrough.AddBall Me:End Sub

'SLINGSHOTS ***********************************************************************************************************************************************

 Dim RStep, Lstep

 Sub LeftSlingShot_Slingshot:vpmTimer.PulseSw 61:End Sub
 Sub RightSlingShot_Slingshot:vpmTimer.PulseSw 62:End Sub


 Sub solLSling(enabled)
	If enabled then
		PlaySound SoundFX("left_slingshot",DOFContactors)
		LSling.Visible = 0
		LSling1.Visible = 1
		sling1.TransZ = -27
		LStep = 0
		LeftSlingShot.TimerEnabled = 1
	End If
 End Sub

 Sub solRSling(enabled)
	If enabled then
		PlaySound SoundFX("right_slingshot",DOFContactors)
		RSling.Visible = 0
		RSling1.Visible = 1
		sling2.TransZ = -27
		RStep = 0
		RightSlingShot.TimerEnabled = 1
	End If
End Sub

 Sub LeftSlingShot_Timer
    Select Case LStep
        Case 3:LSLing1.Visible = 0:LSLing2.Visible = 1:sling1.TransZ = -15
        Case 4:LSLing2.Visible = 0:LSLing.Visible = 1:sling1.TransZ = 0:LeftSlingShot.TimerEnabled = 0
    End Select
    LStep = LStep + 1
 End Sub

 Sub RightSlingShot_Timer
    Select Case RStep
        Case 3:RSLing1.Visible = 0:RSLing2.Visible = 1:sling2.TransZ = -15
        Case 4:RSLing2.Visible = 0:RSLing.Visible = 1:sling2.TransZ = 0:RightSlingShot.TimerEnabled = 0
    End Select
    RStep = RStep + 1
 End Sub

 Sub sw35_Hit:vpmTimer.PulseSwitch 35, 0, "":End Sub


'ROLLOVER SWITCHES ****************************************************************************************************************************************

 Sub sw15_Hit:Controller.Switch(15) = 1:PlaySound "sensor":End Sub:Sub sw15_UnHit:Controller.Switch(15) = 0:End Sub

 Sub sw16_Hit:Controller.Switch(16) = 1:PlaySound "sensor":End Sub:Sub sw16_UnHit:Controller.Switch(16) = 0:End Sub

 Sub sw17_Hit:Controller.Switch(17) = 1:PlaySound "sensor":End Sub:Sub sw17_UnHit:Controller.Switch(17) = 0:End Sub

 Sub sw18_Hit:Controller.Switch(18) = 1:PlaySound "sensor":End Sub:Sub sw18_UnHit:Controller.Switch(18) = 0:End Sub

 Sub sw54_Hit:Controller.Switch(54) = 1:PlaySound "sensor":End Sub:Sub sw54_UnHit:Controller.Switch(54) = 0:End Sub

 Sub sw57_Hit:Controller.Switch(57) = 1:PlaySound "sensor":End Sub:Sub sw57_UnHit:Controller.Switch(57) = 0:End Sub

 Sub sw47_Hit:Controller.Switch(47) = 1:PlaySound "sensor":End Sub:Sub sw47_UnHit:Controller.Switch(47) = 0:End Sub

'RAMP HELPERS *********************************************************************************************************************************************

 Sub RHelp1_Hit:ActiveBall.VelZ = -2:ActiveBall.VelY = 0:ActiveBall.VelX = 0:StopSound "metalrolling":PlaySound "balldrop":End Sub

 Sub RHelp2_Hit:ActiveBall.VelZ = -2:ActiveBall.VelY = 0:ActiveBall.VelX = 0:StopSound "metalrolling":PlaySound "balldrop":End Sub

 Sub RHelp3_Hit:ActiveBall.VelZ = -2:ActiveBall.VelY = 0:ActiveBall.VelX = 0:StopSound "OnlyMetalRolling":PlaySound "CadutaFerri":End Sub

 Sub RHelp4_Hit: PlaySound "balldrop":End Sub

 Sub Trigger1_Hit:ActiveBall.VelZ=0:If ActiveBall.VelY<-6 Then:ActiveBall.VelY=-50:End If:End Sub

 Sub Trigger2_Hit:ActiveBall.VelZ=0:If ActiveBall.VelY<-25 Then:ActiveBall.VelY=-25:End If:End Sub

 Sub Trigger5_Hit:ActiveBall.VelZ=0:If ActiveBall.VelY<-33 Then:ActiveBall.VelY=-33:End If:End Sub

 Sub Trigger6_Hit:ActiveBall.VelZ=0:If ActiveBall.VelY<-10 Then:ActiveBall.VelY=-10:End If:End Sub

 Sub RRL_Hit:ActiveBall.VelZ=0:If ActiveBall.VelY<-10 Then ActiveBall.VelY=-10:End If:End Sub  'USCITA SX RAMPA DX

 Sub RRR_Hit:ActiveBall.VelZ=0:If ActiveBall.VelY<-10 Then ActiveBall.VelY=-10:End If:End Sub  'USCITA DX RAMPA DX

 Sub RSUD_Hit:ActiveBall.VelZ=0:If ActiveBall.VelY<-10 Then ActiveBall.VelY=-10:End If:End Sub 'USCITA DX RAMPA SX

 Sub RSUS_Hit:ActiveBall.VelZ=0:If ActiveBall.VelY<-10 Then ActiveBall.VelY=-10:End If:End Sub 'USCITA SX RAMPA SX

 Sub CadutaBattle_Hit: BattleLose.Enabled= 1 : End Sub

 Sub BattleWin_Hit: PlaySound "rrenter" : End Sub

 Sub BattleLose_Timer()
 Playsound "balldrop"
 Me.enabled=0
 End Sub



'RAMP SWITCHES ********************************************************************************************************************************************

 Sub sw31_Hit:Controller.Switch(31) = 1:End Sub:Sub sw31_Unhit:Controller.Switch(31) = 0:End Sub

 Sub sw32_Hit:Controller.Switch(32) = 1:End Sub:Sub sw32_Unhit:Controller.Switch(32) = 0:End Sub

 Sub sw75_Hit:Controller.Switch(75) = 1:End Sub:Sub sw75_Unhit:Controller.Switch(75) = 0: Playsound "MetalRolling":End Sub

 Sub sw76_Hit:Controller.Switch(76) = 1:End Sub:Sub sw76_Unhit:Controller.Switch(76) = 0: Playsound "MetalRolling":End Sub

 Sub sw77_Hit:Controller.Switch(77) = 1:End Sub:Sub sw77_Unhit:Controller.Switch(77) = 0: Playsound "MetalRolling":End Sub

 Sub sw78_Hit:Controller.Switch(78) = 1:End Sub:Sub sw78_Unhit:Controller.Switch(78) = 0: Playsound "MetalRolling":End Sub

 Sub sw58_Hit:Controller.Switch(58) = 1:End Sub:Sub sw58_Unhit:Controller.Switch(58) = 0:End Sub

'MONGOL TARGETS *******************************************************************************************************************************************

 Sub sw25_Hit:vpmTimer.PulseSw 25:Me.TimerEnabled = 1:PlaySound SoundFX("target",DOFTargets):End Sub
 Sub sw25_Timer:Me.TimerEnabled = 0:End Sub

 Sub sw26_Hit:vpmTimer.PulseSw 26:Me.TimerEnabled = 1:PlaySound SoundFX("target",DOFTargets):End Sub
 Sub sw26_Timer:Me.TimerEnabled = 0:End Sub

 Sub sw27_Hit:vpmTimer.PulseSw 27:Me.TimerEnabled = 1:PlaySound SoundFX("target",DOFTargets):End Sub
 Sub sw27_Timer:Me.TimerEnabled = 0:End Sub

 Sub sw28_Hit:vpmTimer.PulseSw 28:Me.TimerEnabled = 1:PlaySound SoundFX("target",DOFTargets):End Sub
 Sub sw28_Timer:Me.TimerEnabled = 0:End Sub

 Sub sw52_Hit:vpmTimer.PulseSw 52:Me.TimerEnabled = 1:PlaySound SoundFX("target",DOFTargets):End Sub
 Sub sw52_Timer:Me.TimerEnabled = 0:End Sub

 Sub sw53_Hit:vpmTimer.PulseSw 53:Me.TimerEnabled = 1:PlaySound SoundFX("target",DOFTargets):End Sub
 Sub sw53_Timer:Me.TimerEnabled = 0:End Sub

'MINI BATTLEFIELD TARGETS *********************************************************************************************************************************

 Sub sw74_Hit:vpmTimer.PulseSw 74:sw74.IsDropped = 1:Me.TimerEnabled = 1:PlaySound SoundFX("target",DOFTargets):End Sub
 Sub sw74_Timer:sw74.IsDropped = 0:Me.TimerEnabled = 0:End Sub

 Sub sw73_Hit:vpmTimer.PulseSw 73:sw73.IsDropped = 1:Me.TimerEnabled = 1:PlaySound SoundFX("target",DOFTargets):End Sub
 Sub sw73_Timer:sw73.IsDropped = 0:Me.TimerEnabled = 0:End Sub

 Sub sw72_Hit:vpmTimer.PulseSw 72:sw72.IsDropped = 1:Me.TimerEnabled = 1:PlaySound SoundFX("target",DOFTargets):End Sub
 Sub sw72_Timer:sw72.IsDropped = 0:Me.TimerEnabled = 0:End Sub

 Sub sw71_Hit:vpmTimer.PulseSw 71:sw71.IsDropped = 1:Me.Timerenabled = 1:PlaySound SoundFX("target",DOFTargets):End Sub
 Sub sw71_Timer:sw71.IsDropped = 0:Me.TimerEnabled = 0:End Sub

 Sub sw81_Hit:vpmTimer.PulseSw 81:sw81.IsDropped = 1:Me.Timerenabled = 1:PlaySound SoundFX("target",DOFTargets):End Sub
 Sub sw81_Timer:sw81.IsDropped = 0:Me.TimerEnabled = 0:End Sub

 Sub sw82_Hit:vpmTimer.PulseSw 82:sw82.IsDropped = 1:Me.Timerenabled = 1:PlaySound SoundFX("target",DOFTargets):End Sub
 Sub sw82_Timer:sw82.IsDropped = 0:Me.TimerEnabled = 0:End Sub

 Sub sw84_Hit:vpmTimer.PulseSw 84:sw84.IsDropped = 1:Me.Timerenabled = 1:PlaySound SoundFX("target",DOFTargets):End Sub
 Sub sw84_Timer:sw84.IsDropped = 0:Me.TimerEnabled = 0:End Sub

 Sub sw85_Hit:sw85.IsDropped = 1:p85.TransY = -40:Controller.Switch(85) = 1:PlaySound SoundFX("DropTargetDown",DOFDropTargets):vpmTimer.PulseSw 24:End Sub

 Sub sw86_Hit:sw86.IsDropped = 1:p86.TransY = -40:Controller.Switch(86) = 1:PlaySound SoundFX("DropTargetDown",DOFDropTargets):vpmTimer.PulseSw 24:End Sub

 Sub sw87_Hit:sw87.IsDropped = 1:p87.TransY = -40:Controller.Switch(87) = 1:PlaySound SoundFX("DropTargetDown",DOFDropTargets):vpmTimer.PulseSw 24:End Sub

 Sub sw88_Hit:sw88.IsDropped = 1:p88.TransY = -40:Controller.Switch(88) = 1:PlaySound SoundFX("DropTargetDown",DOFDropTargets):vpmTimer.PulseSw 24:End Sub

'MINI BATTLEFIELD DROPBANK HANDLE *************************************************************************************************************************

 Sub SolMiniDropBank(enabled):If enabled Then:PlaySound SoundFX("DropTargetUp",DOFDropTargets):_
 sw85.IsDropped = 0:sw86.IsDropped = 0:sw87.IsDropped = 0:sw88.IsDropped = 0:Controller.Switch(85) = 0:_
 p85.TransY = 0: p86.TransY = 0: p87.TransY = 0: p88.TransY = 0:_
 Controller.Switch(86) = 0:Controller.Switch(87) = 0:Controller.Switch(88) = 0:End If:End Sub

'CENTER TARGET ********************************************************************************************************************************************

 Sub sw56_Hit:vpmTimer.PulseSw 56:Me.TimerEnabled = 1:PlaySound SoundFX("target",DOFTargets):End Sub
 Sub sw56_Timer:Me.TimerEnabled = 0:End Sub

'WALL TARGET **********************************************************************************************************************************************


Sub LHP1_Hit()
If ActiveBall.velY < 0  Then		'ball is going up
PlaySound "rrenter"
Else
StopSound "rrenter"
End If
End Sub

Sub RHP1_Hit()
If ActiveBall.velY < 0  Then		'ball is going up
PlaySound "rrenter"
Else
StopSound "rrenter"
End If
End Sub

 Sub WallTarget_Hit
	Playsound "bersaglio"
 End Sub

 Sub SolWallTargetDown(enabled)
 If enabled Then
 WallTarget.IsDropped = True
 MuroLock.TransY = -70
 Controller.Switch(51) = True
 PlaySound SoundFX("muro",DOFTargets)
 End If
 End Sub

 Sub SolWallTargetUp(enabled)
 If enabled Then
 Controller.Switch(51) = False
 WallTarget.isDropped = False
 MuroLock.TransY = 0
 PlaySound SoundFX("muro",DOFTargets)
 End If
 End Sub

 'GESTIONE FOTOCELLULE LOCK

 Sub sw33_Hit()
 FotocellulaDx.Visible=0
 FotocellulaSx.Visible=0
 Controller.Switch(33) = 1
 End Sub

 Sub sw33_Unhit()
 Controller.Switch(33) = 0
 FotocellulaDx.Visible=1
 FotocellulaSx.Visible=1
 End Sub

 'GESTIONE FOTOCELLULE BATTLE

 Sub FotocellulaDisattivaBat_Hit()
	SpegneFotocelluleBat
 End Sub

 Sub FotocellulaDisattivaBat_Hit()
	FotocellulaDxBat.Visible=0
	FotocellulaSxBat.Visible=0
 End Sub

 Sub FotocellulaDisattivaBat_UnHit()
	TimerFotocellula2.enabled=1
 End Sub

 Sub TimerFotocellula2_Timer()
	FotocellulaDxBat.Visible=1
	FotocellulaSxBat.Visible=1
	Me.Enabled=0
 End Sub

  'BATTLEFIELD TARGET ***************************************************************************************************************************************

 Sub sw55_Hit()
 MuroMiniPlayfield.TransZ = 5
 MuroColpito.Enabled= True
 Controller.Switch(55) = True
 sw55.IsDropped = True
 MuroMiniPlayfield.TransY = -40
 Playsound SoundFX("DropTargetUp",DOFDropTargets)
 End Sub

 Sub MuroColpito_Timer()
 MuroMiniPlayfield.TransZ = 0
 Me.Enabled=0
 End Sub

 Sub SolSingleDropDown(enabled)
 If enabled Then
 sw55.IsDropped = True
 MuroMiniPlayfield.TransY = -40
 Controller.Switch(55) = True
 PlaySound SoundFX("DropTargetDown",DOFDropTargets)
 End If
 End Sub

 Sub SolSingleDropUp(enabled)
 If enabled Then
 sw55.IsDropped = False
 MuroMiniPlayfield.TransY = 0
 Controller.Switch(55) = False
 PlaySound SoundFX("DropTargetUp",DOFDropTargets)
 End If
 End Sub

'LEFT EJECT ***********************************************************************************************************************************************

 Set bsLeftEject = New cvpmBallStack
 bsLeftEject.Ballimage = "ball1"
 bsLeftEject.InitSaucer sw66, 66, 90, 15
 bsLeftEject.InitExitSnd SoundFX("Solenoid",DOFContactors), SoundFX("ExitKicker",DOFContactors)
 bsLeftEject.KickForceVar = 2

 Sub sw66_Hit()
 PlaySound "EnterHole"
 bsLeftEject.AddBall Me
 End Sub

 Sub sw66_UnHit()
 PlaySound SoundFX("ExitHole",DOFContactors)
 End Sub

'RIGHT EJECT **********************************************************************************************************************************************

 Set bsRightEject = New cvpmBallStack
 bsRightEject.Ballimage = "ball1"
 bsRightEject.InitSaucer sw67, 67, 100, 10  ' ERA 67,100,7
 bsRightEject.InitExitSnd SoundFX("Solenoid",DOFContactors), SoundFX("ExitKicker",DOFContactors)
 bsRightEject.KickForceVar = 2

 Sub sw67_Hit()
 PlaySound "EnterHole"
 bsRightEject.AddBall Me
 End Sub

 Sub sw67_UnHit()
 PlaySound SoundFX("ExitHole",DOFContactors)
 End Sub



'PHURBA DIVERTERS *****************************************************************************************************************************************

 Sub SolLeftPhurbaLeft(enabled) 'SCAMBIO SX DEVIA A DESTRA
 If enabled Then
 PlaySound SoundFX("Diverter",DOFContactors)
 ScambioSinistro.TransX= 395
 ScambioSinistro.TransZ= 320
 ScambioSinistro.ObjRotZ= 40
 PugnaleSinistro.TransX= 395
 PugnaleSinistro.TransZ= 293
 PugnaleSinistro.ObjRotZ= 40
 DiverterSxVersoDx.IsDropped= 0
 DiverterSxVersoSx.IsDropped= 1
 End If
 End Sub

 Sub SolLeftPhurbaRight(enabled) 'SCAMBIO SX DEVIA A SINISTRA
 If enabled Then
 PlaySound SoundFX("Diverter",DOFContactors)
 ScambioSinistro.TransX= 0
 ScambioSinistro.TransZ= 0
 ScambioSinistro.ObjRotZ= 0
 PugnaleSinistro.TransX= 0
 PugnaleSinistro.TransZ= 0
 PugnaleSinistro.ObjRotZ= 0
 DiverterSxVersoDx.IsDropped= 1
 DiverterSxVersoSx.IsDropped= 0
 End If
 End Sub


 Sub SolRightPhurbaLeft(enabled)  'SCAMBIO DX DEVIA A DESTRA
 If enabled Then
 PlaySound SoundFX("Diverter",DOFContactors)
 ScambioDestro.TransX= -24
 ScambioDestro.TransZ= 490
 ScambioDestro.ObjRotZ= 35
 PugnaleDestro.TransX= -28
 PugnaleDestro.TransZ= 450
 PugnaleDestro.ObjRotZ= 35
 DiverterDxVersoDx.IsDropped= 0
 DiverterDxVersoSx.IsDropped= 1
' RPLa.IsDropped = False
' RPRa.IsDropped = True
' RightPhurbaRight.isdropped = True
' RightPhurbaLeft.isdropped = False
 End If
 End Sub

 Sub SolRightPhurbaRight(enabled)  'SCAMBIO DX DEVIA A SINISTRA
 If enabled Then
 PlaySound SoundFX("Diverter",DOFContactors)
 ScambioDestro.TransX= 0
 ScambioDestro.TransZ= 0
 ScambioDestro.ObjRotZ= 0

 PugnaleDestro.TransX= 0
 PugnaleDestro.TransZ= 0
 PugnaleDestro.ObjRotZ= 0
 DiverterDxVersoDx.IsDropped= 1
 DiverterDxVersoSx.IsDropped= 0
' RPLa.IsDropped = True
' RPRa.IsDropped = False
' RightPhurbaRight.isdropped = False
' RightPhurbaLeft.isdropped = True
 End If
 End Sub

 'GATE

 Sub GateMissione_Hit()
	Playsound "gate"
 End Sub

 Sub GateLock_Hit()
	Playsound "gate"
 End Sub

'MAGNET HANDLER *****************************************************************************************************************************************

 Sub MagnetOff_Hit()
 BallInMagnet = False
 MagnetKicker.enabled = False
 MagnetCatch.enabled = False
 MagnetHold.enabled = False
 MagnetHelper1.enabled = False
 MagnetHelper2.enabled = False
 End Sub

 Sub MagnetKickerHelper_Hit()
 BallInMagnet = False
 MagnetKicker.enabled = False
 MagnetCatch.enabled = False
 MagnetHold.enabled = False
 MagnetHelper1.enabled = False
 MagnetHelper2.enabled = False
 End Sub

 Sub MagnetCatch_Hit()
 If BallInMagnet = False Then
 MagnetHold.enabled = True
 MagnetCatch.kick 180, 5 '1
 End If
 End Sub

 Sub MagnetKicker_Hit()
 If BallInMagnet = True Then
 MagnetKicker.enabled = False
 MagnetKicker.kick 0, 20
 Playsound "LockWall"
 Else
 MagnetHold.enabled = True
 MagnetKicker.kick 0, 5 '1
 End If
 End Sub

 Sub MagnetHold_Hit()
 BallInMagnet = True
 MagnetHold.destroyball
 MagnetHoldFloat.Createball.image = "Ball1"
 End Sub

 Sub MagnetHelper1_Hit()
 BallInMagnet = True
 MagnetHelper1.destroyball
 MagnetHoldFloat.Createball.image = "Ball1"
 End Sub

 Sub MagnetHelper2_Hit()
 BallInMagnet = True
 MagnetHelper2.destroyball
 MagnetHoldFloat.Createball.image = "Ball1"
 End Sub

 Dim MagneteAttivo

 Sub SolMagnetOn(enabled)
 If enabled Then
 If BallInMagnet = False Then
 MagnetKicker.enabled = True
 MagnetCatch.enabled = True
 Playsound SoundFX("Magnete",DOFShaker) : MagneteAttivo =1
 MagnetHold.enabled = True
 MagnetHelper1.enabled = True
 MagnetHelper2.enabled = True
 End If
 ReleaseMagnetBall = False
 MagnetReleaseTimer.enabled = False
 Else
 ReleaseMagnetBall = True
 MagnetReleaseTimer.enabled = True
 End If
 End Sub

 Sub MagnetReleaseTimer_Timer()
 MagnetReleaseTimer.enabled = False
 If ReleaseMagnetBall = True and BallInMagnet = True Then
 MagnetCatch.enabled = False
 MagnetHold.enabled = False
 MagnetHelper1.enabled = False
 MagnetHelper2.enabled = False
 MagnetHoldFloat.destroyball
 MagnetHold.Createball.image = "Ball1"
 MagnetHold.kick 180, 5
 End If
 End Sub

'BATTLEFIELD MINI KICKER HANDLER **************************************************************************************************************************

 KickerPos = Array(Mk1,Mk2,Mk3,Mk4,Mk5,Mk6,Mk7,Mk8,Mk9,Mk10,Mk11,Mk12,Mk13,Mk14,Mk15,Mk16,Mk17,Mk18,Mk19)

 Sub HandleMiniPF
 FlipperL.RotZ = LeftFlipper.CurrentAngle
 FlipperR.RotZ = RightFlipper.CurrentAngle
 FlipperR2.RotZ = RightFlipper2.CurrentAngle
 PGate1.RotX = Gate1.CurrentAngle
 PGate2.RotX = Gate2.CurrentAngle
 NewKickerPos = Controller.GetMech(0)
 If NewKickerPos <> LastKickerPos Then
 SlingMiniPF.TransX= 90 * (NewKickerPos-9)/9
 KickerPos(LastKickerPos).IsDropped = True
 KickerPos(NewKickerPos).IsDropped = False
 LastKickerPos = NewKickerPos
 End If
 End Sub

 Sub Init_MiniKicker()
 Mk1.IsDropped = True:Mk2.IsDropped = True:Mk3.IsDropped = True:Mk4.IsDropped = True:_
 Mk5.IsDropped = True:Mk6.IsDropped = True:Mk7.IsDropped = True:Mk8.IsDropped = True:_
 Mk9.IsDropped = True:Mk10.IsDropped = True:Mk11.IsDropped = True:Mk12.IsDropped = True:_
 Mk13.IsDropped = True:Mk14.IsDropped = True:Mk15.IsDropped = True:Mk16.IsDropped = True:_
 Mk17.IsDropped = True:Mk18.IsDropped = True:Mk19.IsDropped = True
 End Sub


 Sub Mk1_Slingshot():vpmTimer.PulseSwitch 36, 0, "":End Sub
 Sub Mk2_Slingshot():vpmTimer.PulseSwitch 36, 0, "":End Sub
 Sub Mk3_Slingshot():vpmTimer.PulseSwitch 36, 0, "":End Sub
 Sub Mk4_Slingshot():vpmTimer.PulseSwitch 36, 0, "":End Sub
 Sub Mk5_Slingshot():vpmTimer.PulseSwitch 36, 0, "":End Sub
 Sub Mk6_Slingshot():vpmTimer.PulseSwitch 36, 0, "":End Sub
 Sub Mk7_Slingshot():vpmTimer.PulseSwitch 36, 0, "":End Sub
 Sub Mk8_Slingshot():vpmTimer.PulseSwitch 36, 0, "":End Sub
 Sub Mk9_Slingshot():vpmTimer.PulseSwitch 36, 0, "":End Sub
 Sub Mk10_Slingshot():vpmTimer.PulseSwitch 36, 0, "":End Sub
 Sub Mk11_Slingshot():vpmTimer.PulseSwitch 36, 0, "":End Sub
 Sub Mk12_Slingshot():vpmTimer.PulseSwitch 36, 0, "":End Sub
 Sub Mk13_Slingshot():vpmTimer.PulseSwitch 36, 0, "":End Sub
 Sub Mk14_Slingshot():vpmTimer.PulseSwitch 36, 0, "":End Sub
 Sub Mk15_Slingshot():vpmTimer.PulseSwitch 36, 0, "":End Sub
 Sub Mk16_Slingshot():vpmTimer.PulseSwitch 36, 0, "":End Sub
 Sub Mk17_Slingshot():vpmTimer.PulseSwitch 36, 0, "":End Sub
 Sub Mk18_Slingshot():vpmTimer.PulseSwitch 36, 0, "":End Sub
 Sub Mk19_Slingshot():vpmTimer.PulseSwitch 36, 0, "":End Sub

'BATTLEFIELD VUK ******************************************************************************************************************************************

 Sub MiniPFPopperHelper_Hit
 MiniPFPopperHelper.DestroyBall
 Controller.Switch(68) = True
 PopperEject.Createball.Image = "Ball1"
 PopperEject.Kick 290, 2+rnd*3

 Controller.Switch(68) = False
 End Sub

 Sub Popper_Hit()
 Controller.Switch(68) = True
 Popper.enabled = False
 End Sub

 Sub SolBallPopper(enabled)
 If enabled Then
 If Popper.enabled = False Then
 Controller.Switch(68) = False
 Popper.DestroyBall
 'PlaySound "EnterHole"
 PopperEject.Createball.Image = "Ball1"
 PopperEject.Kick 300, 2+rnd*3
 PlaySound SoundFX("ExitVuk",DOFContactors)
 End If
 Popper.enabled = True
 End If
 End Sub

 Sub Vukup_hit:ActiveBall.Velz=35:End Sub  'ERA Velz=30

 Sub HoleBattle_Hit()
	PlaySound "EnterHole"
 End Sub


'BATTLEFIELD SENSORS******************************************************************************************************************************************

 Sub sw37_Hit:Controller.Switch(37) = 1:End Sub:Sub sw37_Unhit: Controller.Switch(37) = 0:End Sub
 Sub sw38_Hit:Controller.Switch(38) = 1:End Sub:Sub sw38_Unhit: Controller.Switch(38) = 0:End Sub

'BALL LOCK ************************************************************************************************************************************************

 Set bsLock = New cvpmBallStack
 bsLock.Ballimage = "Ball1"
 bsLock.InitSw 0,63,64,65,0,0,0,0
 bsLock.InitKick LockupKickerEject, 120, 2  'ERA 100, 2

 Sub Lockup_Hit()
 bsLock.AddBall Me
 End Sub

 Sub SolLockupKickout(enabled)
 If enabled Then
 PlaySound SoundFX("ExitHole",DOFContactors)
 vpmTimer.AddTimer 250, "Lockup_Serve"
 End If
 End Sub

 Sub Lockup_Serve(No)
 bsLock.ExitSol_On
 End Sub

'FADING LIGHT SYSTEM (C) BY PD/JPSALAS ********************************************************************************************************************

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



'TABLE LIGHTS *********************************************************************************************************************************************

 Sub UpdateLamps

 NFadeL 11, Light11
 NFadeL 12, Light12
 NFadeL 13, Light13
 NFadeL 14, Light14
 NFadeL 15, Light15
 NFadeL 16, Light16
 NFadeL 17, Light17
 NFadeL 18, Light18
 NFadeL 21, Light21
 NFadeL 22, Light22
 NFadeL 23, Light23
 NFadeL 24, Light24
 NFadeL 25, Light25
 NFadeL 26, Light26
 NFadeL 27, Light27
 NFadeL 28, Light28
 NFadeL 31, Light31
 NFadeL 32, Light32
 NFadeL 33, Light33
 NFadeL 34, Light34
 NFadeL 35, Light35
 NFadeL 36, Light36
 NFadeL 37, Light37
 NFadeL 38, Light38
 NFadeL 41, Light41
 NFadeL 42, Light42
 NFadeL 43, Light43
 NFadeL 44, Light44
 NFadeL 45, Light45
 NFadeL 46, Light46
 NFadeL 47, Light47
 NFadeL 48, Light48
 NFadeL 51, Light51
 NFadeL 52, Light52
 NFadeL 53, Light53
 NFadeL 54, Light54
 NFadeL 55, Light55
 NFadeL 56, Light56
 NFadeL 57, Light57
 NFadeL 58, Light58
 NFadeL 61, Light61
 NFadeL 62, Light62
 NFadeL 63, Light63
 NFadeL 64, Light64
 NFadeL 65, Light65
 NFadeL 66, Light66
 NFadeL 67, Light67
 NFadeL 68, Light68
 NFadeL 71, Light71
 NFadeL 72, Light72
 NFadeL 73, Light73
 NFadeL 74, Light74
 NFadeL 75, Light75
 NFadeL 76, Light76
 NFadeL 77, Light77
 NFadeL 78, Light78
 NFadeL 85, Light85
 NFadeL 86, Light86

 ' GESTIONE FLASH

 Flash 117, F117
 Flash 118, F118
 Flashm 121, F121A
 Flash 121, F121
 Flashm 122, F122A
 Flash 122, F122
 Flash 123, F123
 Flash 126, F126
 Flash 127, F127
 Flash 128, F128
 Flash 81, F181
 Flash 82, F182
 Flash 83, F183
 Flash 84, F184
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

Sub NFadeLm(nr, object) ' used for 2 lights
    Select Case FadingLevel(nr)
        Case 4:object.state = 0
        Case 5:object.state = 1
    End Select
End Sub

Sub NFadeLn(nr, object) ' used for 3 lights
    Select Case FadingLevel(nr)
        Case 4:object.state = 0
        Case 5:object.state = 1
    End Select
End Sub


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


Sub SetModLamp(nr, value)
    If value > 0 Then
		LampState(nr) = 1
	Else
		LampState(nr) = 0
	End If
	FadingLevel(nr) = value
End Sub

Sub FadeModLamp(nr, object, factor)
	Object.IntensityScale = FadingLevel(nr) * factor/255
	If TypeName(object) = "Light" Then
		Object.State = LampState(nr)
	End If
	If TypeName(object) = "Flasher" Then
		Object.visible = LampState(nr)
	End If
End Sub


'*****************************************
'           BALL SHADOW
'*****************************************
Dim BallShadow
BallShadow = Array (BallShadow1,BallShadow2,BallShadow3,BallShadow4,BallShadow5)
Dim tnob:tnob = 5
'
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
            BallShadow(b).X = ((BOT(b).X) - (Ballsize/6) + ((BOT(b).X - (Table1.Width/2))/7)) + 10
        Else
            BallShadow(b).X = ((BOT(b).X) + (Ballsize/6) + ((BOT(b).X - (Table1.Width/2))/7)) - 10
        End If
        ballShadow(b).Y = BOT(b).Y + 20
        If BOT(b).Z > 20 Then
            BallShadow(b).visible = 1
        Else
            BallShadow(b).visible = 0
        End If
    Next
End Sub

Function Vol(ball) ' Calculates the Volume of the sound based on the ball speed
    Vol = Csng(BallVel(ball) ^2 / 1)
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

'**********************
' Ball Collision Sound
'**********************

Sub OnBallBallCollision(ball1, ball2, velocity)
	PlaySound("fx_collide"), 0, Csng(velocity) ^2 / 2000, Pan(ball1), 0, Pitch(ball1), 0, 0
End Sub

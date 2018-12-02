'******************************************************
'* SHREK VPX 1.0***** *********************************
'******************************************************
'* STERN PINBALL 2008 *********************************
'******************************************************
'* SAM SYSTEM *****************************************
'******************************************************
'* ROM VERSION 1.41 (shr_141.bin)                     *
'******************************************************
'* CREATED BY NINUZZU AND TOM TOWER********************
'******************************************************
'* Many thanks to Groni, without him this wouldn't be *
'* possible.The table is based on Family Guy by Groni.*
'* Many thanks to comicalman, I used his FP table as  *
'* a resource for playfield images and toys.          *
'* Thanks to arngrim for the DOF script               *
'* Also a huge thanks to Freneticamnesic, who fixed   *
'* the bug with the flashers and to DJRobX, who fixed *
'* the center post script.                            *
'* Shrek, Fiona and Donkey models made by GLXB        *
'* What I did:                                        *
'* Readapted the textures to fit Groni's models       *
'* New Plastics from scatch                           *
'* Some textures redone from scratch                  *
'* Made a new model for the mirror                    *
'* Reworked some 3d models                            *
'* Added light system to Donkey mini-pinball          *
'* Other little stuff (lights, model positioning)     *
'******************************************************
Option Explicit
Randomize


' DjRobX supplied scritp fix for fastflip on this table
' Thalamus 2018-08-04
' Added/Updated "Positional Sound Playback Functions" and "Supporting Ball & Sound Functions"
' Thalamus 2018-11-01 : Improved directional sounds
' !! NOTE : Table not verified yet !!

' Options
' Volume devided by - lower gets higher sound

Const VolDiv = 2000    ' Lower number, louder ballrolling/collition sound
Const VolCol = 10      ' Ball collition divider ( voldiv/volcol )

' The rest of the values are multipliers
'
'  .5 = lower volume
' 1.5 = higher volume

Const VolBump   = 2    ' Bumpers volume.
Const VolGates  = 1    ' Gates volume.
Const VolRH     = 1    ' Rubber hits volume.
Const VolPo     = 1    ' Rubber posts volume.
Const VolFlip   = 1    ' Flipper volume.

'******************************************************
'* TABLE OPTIONS **************************************
'******************************************************
' GI color
Const GIcolor = 0    '0= white, '1= yellow, 2=fantasy

'******************************************************
'* VPM INIT *******************************************
'******************************************************

Dim Ballsize,BallMass
BallSize = 51
BallMass = (Ballsize^3)/125000

On Error Resume Next
ExecuteGlobal GetTextFile("controller.vbs")
If Err Then MsgBox "You need the controller.vbs in order to run this table, available in the vp10 package"
On Error Goto 0

Dim DesktopMode:DesktopMode = Table1.ShowDT
Dim UseVPMDMD:UseVPMDMD = DesktopMode

LoadVPM "01560000","sam.vbs",3.43

'********************
'Standard definitions
'********************

Const UseSolenoids = 1
Const UseLamps = 0
Const UseSync = 0
Const HandleMech = 0
Const SSolenoidOff = ""
Const SCoin = ""

'******************************************************
'* ROM VERSION ****************************************
'******************************************************

Const cGameName = "shr_141"

'******************************************************
'* VARIABLES ******************************************
'******************************************************

Dim bsTrough, PlungerIM, BIP, SwampScoop, MerlinVUK, BABYBank, PinocchioTarget, CPUPos, CPDPos, MiniPF, MiniRight, MiniLeft
Dim  CastleGatePos, CastleGatePos2, CGUp, DonkeyPos, DonkeyPos2, DonkeyPos3, DonkeyDir, DonkeyPinball

'******************************************************
'* KEYS ***********************************************
'******************************************************

Sub Table1_KeyDown(ByVal keycode)

If keycode = PlungerKey Then Plunger.PullBack:End If

If Keycode = RightFlipperKey then
Controller.Switch(90)=1
Controller.Switch(82)=1
If MiniPF.Balls=0 Then
  PlaySoundAtVol SoundFX("Stern_MiniFlipperUp2",DOFContactors),MiniPF_RightFlipperP,VolFlip
MiniPF_RightFlipper.RotateToEnd
DOF 101, 1
MiniRight=1
End If
vpmFFlipsSam.FlipR true
Exit Sub
End If

If Keycode = LeftFlipperKey then
Controller.Switch(84)=1
If MiniPF.Balls=0 Then
PlaySoundAtVol SoundFX("Stern_MiniFlipperUp1",DOFContactors),MiniPF_LeftFlipperP,VolFlip
MiniPF_LeftFlipper.RotateToEnd
DOF 102, 1
MiniLeft=1
End If
vpmFFlipsSam.FlipL true
Exit Sub
End If

If keycode = LeftTiltKey Then Nudge 90, 3:End If
If keycode = RightTiltKey Then Nudge 270, 3:End If
If keycode = CenterTiltKey Then Nudge 0, 3:End If
If vpmKeyDown(keycode) Then Exit Sub
If keycode = PlungerKey Then Plunger.PullBack
End Sub

Sub Table1_KeyUp(ByVal keycode)

If keycode = PlungerKey Then Plunger.Fire:If BIP=1 then PlaysoundAtVol "Stern_Plunge",plunger,1 else If BIP=0 then PlaysoundAtVol "Stern_Hit7", plunger, 1:End If

If Keycode = RightFlipperKey then
Controller.Switch(90)=0
Controller.Switch(82)=0
If MiniRight=1 Then
PlaySoundAtVol SoundFX("Stern_MiniFlipperDown2",DOFContactors), MiniPF_RightFlipperP,VolFlip
MiniPF_RightFlipper.RotateToStart
DOF 101, 0
MiniRight=0
End If
vpmFFlipsSam.FlipR false
Exit Sub
End If

If Keycode = LeftFlipperKey then
Controller.Switch(84)=0
If MiniLeft=1 Then
PlaySoundAtVol SoundFX("Stern_MiniFlipperDown1",DOFContactors),MiniPF_LeftFlipperP,VolFlip
MiniPF_LeftFlipper.RotateToStart
DOF 102, 0
MiniLeft=0
End If
vpmFFlipsSam.FlipL false
Exit Sub
End If

If vpmKeyUp(keycode) Then Exit Sub
If keycode = PlungerKey Then Plunger.Fire
End Sub

'******************************************************
'******************************************************
'******************************************************
'* TABLE INIT *****************************************
'******************************************************
'******************************************************
'******************************************************

Sub Table1_Init

'* ROM AND DMD ****************************************

With Controller
.GameName = cGameName
.SplashInfoLine = "SHREK - STERN 2008"
.HandleMechanics = 0
.HandleKeyboard = 0
.ShowDMDOnly = 1
.ShowFrame = 0
.ShowTitle = 0
.Hidden = DesktopMode
On Error Resume Next
.Run
If Err Then MsgBox Err.Description
On Error Goto 0
End With

'* PINMAME TIMER **************************************

PinMAMETimer.Interval = PinMAMEInterval
PinMAMETimer.Enabled = 1

'* STARTUP CALLS **************************************
CPC.isdropped=1
kicker1.CreateSizedBallWithMass BallSize/2,Ballmass
kicker1.kick 180, 1
kicker1.enabled = 0
kicker2.CreateSizedBallWithMass BallSize/2,Ballmass
kicker2.kick 180, 1
kicker2.enabled = 0
kicker3.CreateSizedBallWithMass BallSize/2,Ballmass
kicker3.kick 180, 1
kicker3.enabled = 0
kicker4.CreateSizedBallWithMass BallSize/2,Ballmass
kicker4.kick 180, 1
kicker4.enabled = 0

'* NUGDE **********************************************

vpmNudge.TiltSwitch=-7
vpmNudge.Sensitivity=3
vpmNudge.TiltObj=Array(bumper1,bumper2,bumper3,LeftSlingshot,RightSlingshot)

'* TROUGH *********************************************

Set bsTrough = New cvpmBallStack
With bsTrough
.InitSw 0,21,20,19,18,0,0,0
.InitKick BallRelease, 90, 8
.InitExitSnd SoundFX("Stern_Release",DOFContactors), ""
.Balls = 4
End With

'* MINI TROUGH ****************************************

Set MiniPF = New cvpmBallStack
With MiniPF
.InitSaucer sw55,55, 90, 35
.InitExitSnd "", ""
.InitAddSnd ""
End With
sw55.createsizedball(14.6875)			'minipinball 5/8"
MiniPF.AddBall 0


'* SWAMP SCOOP *******************************************

Set SwampScoop = New cvpmBallStack
With SwampScoop
.InitSw 0,13,0,0,0,0,0,0
.InitKick sw13, 184, 30
.InitExitSnd SoundFX("Stern_Scoopexit",DOFContactors), ""
End With

'* MERLIN VUK *****************************************

Set MerlinVUK = New cvpmBallStack
With MerlinVUK
.InitSaucer sw64,64, 183, 4
.InitExitSnd "",""
.InitAddSnd ""
End With

'* BABY TARGETS 4-BANK DROPTARGETS ********************

Set BABYBank = New cvpmDropTarget
With BABYBank
.InitDrop Array(sw47,sw44,sw46,sw45), Array(44,47,45,46)
.Initsnd "", ""
End With

'* PINOCCHIO DROPTARGET *******************************

Set PinocchioTarget = New cvpmDropTarget
With PinocchioTarget
.InitDrop Array(Array(sw9, sw9w)), Array(9)
.Initsnd "", ""
End With

'* IMPULSE PLUNGER ************************************

Const IMPowerSetting = 33
Const IMTime = 0.33
Set PlungerIM = New cvpmImpulseP
With plungerIM
.InitImpulseP BIPREG, IMPowerSetting, IMTime
.InitExitSnd "Stern_Autoplung", ""
.CreateEvents "PlungerIM"
End With

InitVpmFFlipsSAM

'******************************************************
'******************************************************
'******************************************************
'* TABLE INIT END *************************************
'******************************************************
'******************************************************
'******************************************************

End Sub

Sub table1_Paused:Controller.Pause = 1:End Sub
Sub table1_unPaused:Controller.Pause = 0:End Sub
Sub table1_exit():Controller.Stop:End sub

'******************************************************
'* SOLENOID CALLS *************************************
'******************************************************

SolCallback(1)  = "solTrough"
SolCallback(2)  = "AutoLaunch"
SolCallback(3)  = "BankReset"
SolCallback(4)  = "CPDSol"                   'Ball Saver Down
SolCallback(5)  = "MerlinEject"
SolCallback(6)  = "PinocchioReset"
'SolCallback(7)  = "LeftSlingHit"            'Left Slingshot
'SolCallback(8)  = "RightSlingHit"           'Right Slingshot
'SolCallback(9)  = "Bump1Sol"                'Bottom Bumper
'SolCallback(10) = "Bump2Sol"                'Right Bumper
'SolCallback(11) = "Bump3Sol"                'Top Bumper
SolCallback(12) = "CPUSol"                   'Ball Saver Up
SolCallback(13) = "Scoopexit"                'Swamp Eject Scoop
SolCallback(15) = "LFlipper"
SolCallback(16) = "RFlipper"
SolCallback(17) = "Miniflipper_Left"
SolCallback(18) = "Miniflipper_Right"
SolCallback(19) = "CastleGuardSol"
SolCallBack(20) = "DonkeyMove"				'Donkey Motor Drive
SolCallBack(21) = "MiniPFSol"
SolCallback(23) = "setlamp 193,"			'Flash Lower Left
SolCallback(25) = "setlamp 195,"			'Flash Backpanel Left
SolCallback(26) = "setlamp 196,"			'Flash Backpanel Center
SolCallback(27) = "setlamp 197,"			'Flash Backpanel Right
SolCallback(28) = "setlamp 198,"			'Flash Magic Mirror
SolCallback(29) = "setlamp 199,"			'Flash Fiona
SolCallback(30) = "setlamp 190,"			'Flash RIght Orbit (Spinner)
SolCallback(31) = "setlamp 191,"			'Flash Pop Bumpers
SolCallback(32) = "setlamp 192,"			'Flash Lower Right

'******************************************************
'* CASTLE GUARD GATE **********************************
'******************************************************

Sub CastleGuardsw_Hit()
Controller.Switch(35)=0
CastleGuardUp.enabled=1
CGUp=1
CastleGuardsw.isdropped=1
End Sub

Sub CastleGuardSol (enabled)
If enabled then
Controller.Switch(35)=1
End if
End Sub

Sub CastleGuardUp_Timer()
Select Case CastleGatePos
Case 0: CastleGuardP.RotX=0:CastleGatePos=1
Case 1: CastleGuardP.RotX=0:
Case 2: CastleGuardP.RotX=-15
Case 3: CastleGuardP.RotX=-30
Case 4: CastleGuardP.RotX=-45
Case 5: CastleGuardP.RotX=-60
Case 6: CastleGuardP.RotX=-75
Case 7: CastleGuardP.RotX=-90:Me.Enabled=0:CastleGatePos=0
End Select
If CastleGatePos>0 then CastleGatePos=CastleGatePos+1
End Sub

Sub CastleGuardDown_Timer()
Select Case CastleGatePos2
Case 0: CastleGuardP.RotX=-90:CastleGatePos2=1
Case 1: CastleGuardP.RotX=-90:
Case 2: CastleGuardP.RotX=-75
Case 3: CastleGuardP.RotX=-60
Case 4: CastleGuardP.RotX=-45
Case 5: CastleGuardP.RotX=-30
Case 6: CastleGuardP.RotX=-15
Case 7: CastleGuardP.RotX=0:CastleGuardsw.isdropped=0:CGUp=0
Case 8: CastleGuardP.RotX=5
Case 9: CastleGuardP.RotX=10
Case 10: CastleGuardP.RotX=15
Case 11: CastleGuardP.RotX=20
Case 12: CastleGuardP.RotX=15
Case 13: CastleGuardP.RotX=10
Case 14: CastleGuardP.RotX=5
Case 15: CastleGuardP.RotX=0
Case 16: CastleGuardP.RotX=-5
Case 17: CastleGuardP.RotX=-10
Case 18: CastleGuardP.RotX=-15
Case 19: CastleGuardP.RotX=-10
Case 20: CastleGuardP.RotX=-5
Case 21: CastleGuardP.RotX=0
Case 22: CastleGuardP.RotX=5
Case 23: CastleGuardP.RotX=10
Case 24: CastleGuardP.RotX=15
Case 25: CastleGuardP.RotX=10
Case 26: CastleGuardP.RotX=5
Case 27: CastleGuardP.RotX=0
Case 28: CastleGuardP.RotX=-5
Case 29: CastleGuardP.RotX=-10
Case 30: CastleGuardP.RotX=-15
Case 31: CastleGuardP.RotX=-10
Case 32: CastleGuardP.RotX=-5
Case 33: CastleGuardP.RotX=0
Case 34: CastleGuardP.RotX=5
Case 35: CastleGuardP.RotX=10
Case 36: CastleGuardP.RotX=10
Case 37: CastleGuardP.RotX=5
Case 38: CastleGuardP.RotX=0
Case 39: CastleGuardP.RotX=-5
Case 40: CastleGuardP.RotX=-10
Case 41: CastleGuardP.RotX=-10
Case 42: CastleGuardP.RotX=-5
Case 43: CastleGuardP.RotX=0
Case 44: CastleGuardP.RotX=5
Case 45: CastleGuardP.RotX=10
Case 46: CastleGuardP.RotX=5
Case 47: CastleGuardP.RotX=0
Case 48: CastleGuardP.RotX=-5
Case 49: CastleGuardP.RotX=-10
Case 50: CastleGuardP.RotX=-5
Case 51: CastleGuardP.RotX=0
Case 52: CastleGuardP.RotX=5
Case 53: CastleGuardP.RotX=0
Case 54: CastleGuardP.RotX=-5
Case 55: CastleGuardP.RotX=0
Case 56: CastleGuardP.RotX=5
Case 57: CastleGuardP.RotX=0:Me.Enabled=0:CastleGatePos2=0
End Select
If CastleGatePos2>0 then CastleGatePos2=CastleGatePos2+1
End Sub

'******************************************************
'* MINI PLAYFIELD *************************************
'******************************************************

Sub Miniflipper_Right(Enabled)
If Enabled Then
PlaySoundAtVol SoundFX("Stern_MiniFlipperUp1",DOFContactors),MiniPF_RightFlipperP,VolFlip:MiniPF_RightFlipper.RotateToEnd
Else
PlaySoundAtVol SoundFX("Stern_MiniFlipperDown1",DOFContactors),MiniPF_RightFlipperP,VolFlip:MiniPF_RightFlipper.RotateToStart
End If
End Sub

Sub Miniflipper_Left(Enabled)
If Enabled Then
DonkeySWTimer.Enabled=1
PlaySoundAtVol SoundFX("Stern_MiniFlipperUp2",DOFContactors),MiniPF_LeftFlipperP,VolFlip:MiniPF_LeftFlipper.RotateToEnd
Else
PlaySoundAtVol SoundFX("Stern_MiniFlipperDown2",DOFContactors),MiniPF_LeftFlipperP,VolFlip:MiniPF_LeftFlipper.RotateToStart
End If
End Sub

Sub MiniPFSol(enabled)
If enabled then
MiniPF.ExitSol_On
end if
End Sub

Sub sw55_Hit():MiniPF.AddBall 0:End Sub

Sub DonkeySWTimer_Timer()
DonkeySW.enabled=1
End Sub

Sub DonkeySW_hit()
DonkeyPinball=2
End Sub

Sub DonkeyTurn_Hit()
If DonkeyPinball=0 then
DonkeyTimer2.enabled=1
DonkeyPinball=3
DonkeySWTimer.enabled=1
End If
End Sub

Sub DonkeyReset_Hit
If DonkeyPinball >0 then
DonkeyPinball = 0
End If
End Sub

Sub DonkeyDelay_Timer()
DonkeyTimer.Enabled = 1
DonkeyDelay.Enabled = 0
End Sub

'******************************************************
'* DONKEY MOVEMENT ************************************
'******************************************************

Sub DonkeyMove (enabled)
If enabled AND DonkeyPinball=0 then
DonkeyDelay.Enabled = 1
Else
If enabled AND DonkeyPinball=2 then
DonkeyTimer3.Enabled = 1
End If
End If
End Sub

Sub DonkeyTimer_Timer()
Select Case DonkeyPos
Case 0: DonkeyP.RotY=0:DonkeyPos3=0:DonkeyPos=1
Case 1: DonkeyP.RotY=+2
Case 2: DonkeyP.RotY=+4
Case 3: DonkeyP.RotY=+6
Case 4: DonkeyP.RotY=+8
Case 5: DonkeyP.RotY=+10
Case 6: DonkeyP.RotY=+12
Case 7: DonkeyP.RotY=+10
Case 8: DonkeyP.RotY=+8
Case 9: DonkeyP.RotY=+6
Case 10: DonkeyP.RotY=+4
Case 11: DonkeyP.RotY=+2
Case 12: DonkeyP.RotY=0:DonkeyPos=1:DonkeyTimer.Enabled=0:DonkeyPinball=0
End Select
If DonkeyPos>0 then DonkeyPos=DonkeyPos+1
Sh5.RotZ = DonkeyP.RotY
End Sub

Sub DonkeyTimer2_Timer()
Select Case DonkeyPos2
Case 0: DonkeyP.RotY=0:DonkeyPos2=1
Case 1: DonkeyP.RotY=0:
Case 2: DonkeyP.RotY=15
Case 3: DonkeyP.RotY=30
Case 4: DonkeyP.RotY=45
Case 5: DonkeyP.RotY=60
Case 6: DonkeyP.RotY=75
Case 7: DonkeyP.RotY=90
Case 8: DonkeyP.RotY=105
Case 9: DonkeyP.RotY=120:DonkeyTimer2.Enabled=0:DonkeyPos2=0:DonkeyPinball=3
End Select
If DonkeyPos2>0 then DonkeyPos2=DonkeyPos2+1
Sh5.RotZ = DonkeyP.RotY
End Sub

Sub DonkeyTimer3_Timer()
Select Case DonkeyPos3
Case 0: DonkeyP.RotY=120:DonkeyPos3=1
Case 1: DonkeyP.RotY=105
Case 2: DonkeyP.RotY=90
Case 3: DonkeyP.RotY=70
Case 4: DonkeyP.RotY=60
Case 5: DonkeyP.RotY=45
Case 6: DonkeyP.RotY=30
Case 7: DonkeyP.RotY=15
Case 8: DonkeyP.RotY=0
Case 9: DonkeyP.RotY=0:DonkeyTimer3.Enabled=0:DonkeyPos3=0:DonkeySWTimer.enabled=0:DonkeyPinball=0:DonkeySW.enabled=0
End Select
If DonkeyPos3>0 then DonkeyPos3=DonkeyPos3+1
Sh5.RotZ = DonkeyP.RotY
End Sub

'******************************************************
'* TROUGH *********************************************
'******************************************************

Sub solTrough(enabled)
If enabled then
bsTrough.ExitSol_On
vpmTimer.PulseSw 22
end if
End Sub

'******************************************************
'* AUTOLAUNCH *****************************************
'******************************************************

Sub AutoLaunch(Enabled)
If Enabled Then
PlungerIM.AutoFire
End If
End Sub

'******************************************************
'* SWAMP SCOOP ****************************************
'******************************************************

Dim bBall, bZpos

Sub Scoopexit(enabled)
If enabled then
SwampScoop.ExitSol_On
end if
End Sub

Sub sw13_Hit
Set bBall = ActiveBall
PlaySoundAtVol "Stern_Scoopenter", ActiveBall, 1
bZpos = 50
Me.TimerInterval = 2
Me.TimerEnabled = 1
End Sub

Sub sw13_Timer
bBall.Z = bZpos
bZpos = bZpos-2
If bZpos <40 Then
Me.TimerEnabled = 0
Me.DestroyBall
SwampScoop.AddBall Me
End If
End Sub

'******************************************************
'* MERLIN VUK *****************************************
'******************************************************

Sub MerlinEject(enabled)
If enabled then
MerlinVUK.ExitSol_On
End if
End Sub

Sub sw64_Hit:PlaysoundAtVol"Stern_Scoopenter",ActiveBall,1:MerlinVUK.AddBall 0:End Sub

'******************************************************
'* GINGY SPINNER **************************************
'******************************************************

Sub sw39_Spin:vpmTimer.PulseSw 39:End Sub

'******************************************************
'* DRAIN **********************************************
'******************************************************

Sub Drain_Hit()
PlaySoundAtVol "Stern_Drain", drain ,1
bsTrough.AddBall Me
End Sub

'******************************************************
'* POP BUMPERS ****************************************
'******************************************************

Dim dirRing1 : dirRing1 = -1
Dim dirRing2 : dirRing2 = -1
Dim dirRing3 : dirRing3 = -1

Sub Bumper1_Hit : vpmTimer.PulseSw 32 : PlaySoundAtVol SoundFX("Stern_Bump1",DOFContactors),Bumper1,VolBump : Me.TimerEnabled = 1 : End Sub
Sub Bumper2_Hit : vpmTimer.PulseSw 31 : PlaySoundAtVol SoundFX("Stern_Bump2",DOFContactors),Bumper2,VolBump : Me.TimerEnabled = 1 : End Sub
Sub Bumper3_Hit : vpmTimer.PulseSw 30 : PlaySoundAtVol SoundFX("Stern_Bump3",DOFContactors),Bumper3,VolBump : Me.TimerEnabled = 1 : End Sub

Sub Bumper1_timer()
	BR1.Z = BR1.Z + (5 * dirRing1)
	If BR1.Z <= -35 Then dirRing1 = 1
	If BR1.Z >= 0 Then
		dirRing1 = -1
		BR1.Z = 0
		Me.TimerEnabled = 0
	End If
End Sub

Sub Bumper2_timer()
	BR2.Z = BR2.Z + (5 * dirRing2)
	If BR2.Z <= -35 Then dirRing2 = 1
	If BR2.Z >= 0 Then
		dirRing2 = -1
		BR2.Z = 0
		Me.TimerEnabled = 0
	End If
End Sub

Sub Bumper3_timer()
	BR3.Z = BR3.Z + (5 * dirRing3)
	If BR3.Z <= -35 Then dirRing3 = 1
	If BR3.Z >= 0 Then
		dirRing3 = -1
		BR3.Z = 0
		Me.TimerEnabled = 0
	End If
End Sub

'******************************************************
'* TARGETS ********************************************
'******************************************************

Sub sw3_Hit:vpmTimer.PulseSw 3:PlaySoundAtVol SoundFX("Stern_Hit1",DOFContactors),ActiveBall,1:If CGUp= 1 then CastleGuardDown.enabled=1:End If:End Sub

Sub sw4_Hit:vpmTimer.PulseSw 4:PlaySoundAtVol SoundFX("Stern_Hit1",DOFContactors),ActiveBall, 1:End Sub

Sub sw5_Hit:vpmTimer.PulseSw 5:PlaySoundAtVol SoundFX("Stern_Hit1",DOFContactors),ActiveBall, 1:End Sub

Sub sw8_Hit:vpmTimer.PulseSw 8:PlaySoundAtVol SoundFX("Stern_Hit1",DOFContactors),ActiveBall, 1:End Sub

Sub sw10_Hit:vpmTimer.PulseSw 10:PlaySoundAtVol SoundFX("Stern_Hit1",DOFContactors),ActiveBall, 1:End Sub

Sub sw41_Hit:vpmTimer.PulseSw 41:PlaySoundAtVol SoundFX("Stern_Hit1",DOFContactors),ActiveBall, 1:End Sub

Sub sw42_Hit:vpmTimer.PulseSw 42:PlaySoundAtVol SoundFX("Stern_Hit1",DOFContactors),ActiveBall, 1:End Sub

Sub sw43_Hit:vpmTimer.PulseSw 43:PlaySoundAtVol SoundFX("Stern_Hit1",DOFContactors),ActiveBall, 1:End Sub


'******************************************************
'* MINIPF TARGETS **************************************
'******************************************************

Sub sw50_Hit:vpmTimer.PulseSw 50:PlaySoundAtVol SoundFX("Stern_Hit1",DOFContactors),ActiveBall, 1:End Sub
Sub sw51_Hit:vpmTimer.PulseSw 51:PlaySoundAtVol SoundFX("Stern_Hit1",DOFContactors),ActiveBall, 1:End Sub

'******************************************************
'* MINI PF SWITCHES ***********************************
'******************************************************

Sub sw52_Hit:Controller.Switch(52) = 1:End Sub
Sub sw52_Unhit:Controller.Switch(52) = 0:End Sub

Sub sw53_Hit:Controller.Switch(53) = 1:End Sub
Sub sw53_Unhit:Controller.Switch(53) = 0:End Sub

Sub sw54_Hit:Controller.Switch(54) = 1:End Sub
Sub sw54_Unhit:Controller.Switch(54) = 0:End Sub

'******************************************************
'* ROLLOVER SWITCHES **********************************
'******************************************************

Sub sw6_Hit:Controller.Switch(6) = 1:sw6p.Z=-2:End Sub
Sub sw6_Unhit:Controller.Switch(6) = 0:sw6p.Z=0:End Sub

Sub sw7_Hit:Controller.Switch(7) = 1:sw7p.Z=-2:End Sub
Sub sw7_Unhit:Controller.Switch(7) = 0:sw7p.Z=0:End Sub

Sub sw23_Hit:Controller.Switch(23) = 1:sw23p.Z=-2:BIP=1:End Sub
Sub sw23_Unhit:Controller.Switch(23) = 0:sw23p.Z=0:BIP=0:End Sub

Sub sw24_Hit:Controller.Switch(24) = 1:sw24p.Z=-2:End Sub
Sub sw24_Unhit:Controller.Switch(24) = 0:sw24p.Z=0:End Sub

Sub sw25_Hit:Controller.Switch(25) = 1:sw25p.Z=-2:End Sub
Sub sw25_Unhit:Controller.Switch(25) = 0:sw25p.Z=0:End Sub

Sub sw28_Hit:Controller.Switch(28) = 1:sw28p.Z=-2:End Sub
Sub sw28_Unhit:Controller.Switch(28) = 0:sw28p.Z=0:End Sub

Sub sw29_Hit:Controller.Switch(29) = 1:sw29p.Z=-2:End Sub
Sub sw29_Unhit:Controller.Switch(29) = 0:sw29p.Z=0:End Sub

Sub sw40_Hit:Controller.Switch(40) = 1:sw40p.Z=-2:End Sub
Sub sw40_Unhit:Controller.Switch(40) = 0:sw40p.Z=0:End Sub

Sub sw48_Hit:Controller.Switch(48) = 1:End Sub
Sub sw48_Unhit:Controller.Switch(48) = 0:End Sub

Sub sw57_Hit:Controller.Switch(57) = 1:End Sub
Sub sw57_Unhit:Controller.Switch(57) = 0:End Sub

'******************************************************
'* SLINGSHOTS *****************************************
'******************************************************
Dim LStep, RStep
Sub LeftSlingShot_Slingshot
    PlaysoundAtVol SoundFX("Stern_LeftSlingshot",DOFContactors), sling1, 1
    vpmTimer.PulseSw 26
    LSling.Visible = 0
    LSling1.Visible = 1
    sling1.TransZ = -20
    LStep = 0
    LeftSlingShot.TimerEnabled = 1
End Sub

Sub LeftSlingShot_Timer
    Select Case LStep
        Case 3:LSLing1.Visible = 0:LSLing2.Visible = 1:sling1.TransZ = -10
        Case 4:LSLing2.Visible = 0:LSLing.Visible = 1:sling1.TransZ = 0:LeftSlingShot.TimerEnabled = 0
    End Select
    LStep = LStep + 1
End Sub


Sub RightSlingShot_Slingshot
    PlaysoundAtVol SoundFX("Stern_RightSlingshot",DOFContactors), sling2, 1
    vpmTimer.PulseSw 27
    RSling.Visible = 0
    RSling1.Visible = 1
    sling2.TransZ = -20
    RStep = 0
    RightSlingShot.TimerEnabled = 1
End Sub

Sub RightSlingShot_Timer
    Select Case RStep
        Case 3:RSLing1.Visible = 0:RSLing2.Visible = 1:sling2.TransZ = -10
        Case 4:RSLing2.Visible = 0:RSLing.Visible = 1:sling2.TransZ = 0:RightSlingShot.TimerEnabled = 0
    End Select
    RStep = RStep + 1
End Sub

'******************************************************
'* CENTER POST ****************************************
'******************************************************
'* DJRobX fixed the ball saver emulation error. Game  *
'* seems to fire both up and down to put the saver    *
'* down. Create a "latch" to hold it down'            *
'******************************************************

Dim SaverPos, SaverTime
SaverPos = 0
SaverTime = Now

Sub CPDSol (enabled)
If enabled then
SaverTime = DateAdd("s",3, Now)
CPUSol (true)
End If
End Sub

Sub CPUSol (enabled)
If enabled then
If SaverTime > Now Then
If SaverPos = 1 Then
' ** Move Saver Down **'
Controller.Switch(2) = 1
Controller.Switch(1) = 0
CenterpostDown.Enabled=1
SaverPos=0
CPC.isdropped=1
End If
Else
If SaverPos = 0 Then
' ** Move Saver Up ** '
Controller.Switch(2) = 0
Controller.Switch(1) = 1
CenterpostUp.Enabled=1
SaverPos=1
CPC.isdropped=0
End If
End If
End If
End Sub

Sub CenterpostUp_Timer()
Select Case CPUPos
Case 0: CenterPost.Z=-15
Case 1: CenterPost.Z=-05:PlaysoundAtVol SoundFX("Stern_Centerpost_Up",DOFContactors), CenterPost, 1
Case 2: CenterPost.Z=0
Case 3: CenterPost.Z=0:CPUPos=0:CenterpostUp.Enabled=0
End Select
If CPUPos=>0 then CPUPos=CPUPos+1
End Sub

Sub CenterpostDown_Timer()
Select Case CPDPos
Case 0: CenterPost.Z=-05
Case 1: CenterPost.Z=-15:PlaysoundAtVol SoundFX("Stern_Centerpost_Down",DOFContactors), CenterPost, 1
Case 2: CenterPost.Z=-26
Case 3: CenterPost.Z=-26:CPDPos=0:CenterpostDown.Enabled=0
End Select
If CPDPos=>0 then CPDPos=CPDPos+1
End Sub

'******************************************************
'* MAGIC MIRROR TARGET ********************************
'******************************************************
Dim MirrorPos:MirrorPos=1
Dim MirrorDir : MirrorDir=0
Sub sw49_Hit:vpmTimer.PulseSw 49:Me.TimerEnabled = 1:PlaySoundAtVol SoundFX("Stern_Beercanhit",DOFContactors),ActiveBall, 1:End Sub

Sub sw49_Timer()
MirrorP.Rotx = MirrorP.Rotx - 2*MirrorPos
If MirrorP.Rotx < - 10 AND MirrorPos=1 Then MirrorPos=-1
If MirrorP.Rotx > 6 AND MirrorPos=-1 Then MirrorDir=1 : MirrorPos=1
If MirrorP.Rotx <0 AND MirrorDir=1 Then MirrorP.Rotx =0 : MirrorDir=0 : Me.TimerEnabled = 0
End Sub

'******************************************************
'* PINOCCHIO TARGET ***********************************
'******************************************************

Dim sw9Dir

Sub Pinocchioreset (enabled)
If enabled then
Playsound SoundFX("Stern_Bankraise",DOFContactors)
sw9Dir = 1
sw9.TimerEnabled = 1
sw9w.isdropped=0
sw9.isdropped=0
PinocchioTarget.DropSol_On
End If
End Sub

Sub sw9_Hit:PinocchioTarget.Hit 1:sw9Dir=-1:Me.TimerEnabled = 1:End Sub

Sub sw9_Timer()
Select Case sw9Dir
Case -1
sw9P.z = sw9P.z -5
If sw9p.Z <  -65 Then sw9p.Z = -65 : Me.TimerEnabled = 0
Case 1
sw9P.z = sw9P.z + 5
If sw9p.Z >  0 Then sw9p.Z = 0 : Me.TimerEnabled = 0
End Select
End Sub

'******************************************************
'* BABY TARGETBANK ************************************
'******************************************************
Dim sw44Dir, sw45Dir, sw46Dir, sw47Dir

Sub Bankreset (enabled)
If enabled then
Playsound SoundFX("Stern_Bankraise",DOFContactors) ' TODO
sw44Dir = 1
sw44.TimerEnabled = 1
sw45Dir = 1
sw45.TimerEnabled = 1
sw46Dir = 1
sw46.TimerEnabled = 1
sw47Dir = 1
sw47.TimerEnabled = 1
BABYBank.DropSol_On
End If
End Sub

Sub sw44_Hit:BABYBank.Hit 2:sw44Dir = -1:Me.TimerEnabled = 1:End Sub
Sub sw45_Hit:BABYBank.Hit 4:sw45Dir = -1:Me.TimerEnabled = 1:End Sub
Sub sw46_Hit:BABYBank.Hit 3:sw46Dir = -1:Me.TimerEnabled = 1:End Sub
Sub sw47_Hit:BABYBank.Hit 1:sw47Dir = -1:Me.TimerEnabled = 1:End Sub

'* Y TARGET *******************************************

Sub sw44_Timer()
Select Case sw44Dir
Case -1
sw44P.z = sw44P.z -5
If sw44p.Z <  -60 Then sw44p.Z = -60 : Me.TimerEnabled = 0
Case 1
sw44P.z = sw44P.z + 5
If sw44p.Z >  0 Then sw44p.Z = 0 : Me.TimerEnabled = 0
End Select
End Sub

'* B TARGET *******************************************

Sub sw45_Timer()
Select Case sw45Dir
Case -1
sw45P.z = sw45P.z -5
If sw45p.Z <  -60 Then sw45p.Z = -60 : Me.TimerEnabled = 0
Case 1
sw45P.z = sw45P.z + 5
If sw45p.Z >  0 Then sw45p.Z = 0 : Me.TimerEnabled = 0
End Select
End Sub

'* A TARGET *******************************************
Sub sw46_Timer()
Select Case sw46Dir
Case -1
sw46P.z = sw46P.z -5
If sw46p.Z <  -60 Then sw46p.Z = -60 : Me.TimerEnabled = 0
Case 1
sw46P.z = sw46P.z + 5
If sw46p.Z >  0 Then sw46p.Z = 0 : Me.TimerEnabled = 0
End Select
End Sub

'* B TARGET *******************************************
Sub sw47_Timer()
Select Case sw47Dir
Case -1
sw47P.z = sw47P.z -5
If sw47p.Z <  -60 Then sw47p.Z = -60 : Me.TimerEnabled = 0
Case 1
sw47P.z = sw47P.z + 5
If sw47p.Z >  0 Then sw47p.Z = 0 : Me.TimerEnabled = 0
End Select
End Sub

'******************************************************
'* FLIPPER CALLS **************************************
'******************************************************

SolCallback(sLRFlipper) = "RFlipper"
SolCallback(sLLFlipper) = "LFlipper"

Sub LFlipper(Enabled)
If Enabled Then
If LeftFlipper.CurrentAngle > 100 then PlaySoundAtVol SoundFX("Stern_LeftFlipper_Up",DOFContactors), LeftFlipper, VolFlip
If LeftFlipper.CurrentAngle < 100 then PlaySoundAtVol SoundFX("Stern_LeftFlipper_Up2",DOFContactors), LeftFlipper, VolFlip
LeftFlipper.RotateToEnd:LeftFlipperSmall.RotateToEnd
Else
PlaySoundAtVol SoundFX("Stern_Flipper_Down",DOFContactors),LeftFlipper,VolFlip:LeftFlipper.RotateToStart:LeftFlipperSmall.RotateToStart
End If
End Sub

Sub RFlipper(Enabled)
If Enabled Then
If RightFlipper.CurrentAngle < -100 then PlaySoundAtVol SoundFX("Stern_RightFlipper_Up",DOFContactors), RightFlipper, VolFlip
If RightFlipper.CurrentAngle > -100 then PlaySoundAtVol SoundFX("Stern_RightFlipper_Up2",DOFContactors), RightFlipper, VolFlip
RightFlipper.RotateToEnd
Else
PlaySoundAtVol SoundFX("Stern_Flipper_Down",DOFContactors),RightFlipper,VolFlip:RightFlipper.RotateToStart
End If
End Sub

'******************************************************
'* LEFT GATE ******************************************
'******************************************************

Sub LeftGate_hit:vpmTimer.PulseSw 33:LeftGateFl.RotateToEnd:End Sub
Sub LeftGate_Unhit:LeftGateFl.RotateToStart:End Sub

'******************************************************
'* REAL TIME UPDATES *********************************
'******************************************************

Set MotorCallback = GetRef("GameTimer")

Sub GameTimer
RightFlipperP.Rotz = RightFlipper.CurrentAngle
LeftFlipperP.Rotz = LeftFlipper.CurrentAngle
LeftFlipperSmallP.RotY = LeftFlipperSmall.CurrentAngle
MiniPF_LeftFlipperP.Rotz = MiniPF_LeftFlipper.CurrentAngle
MiniPF_RightFlipperP.Rotz = MiniPF_RightFlipper.CurrentAngle
FlipperLSh.RotZ = LeftFlipper.currentangle
FlipperRSh.RotZ = RightFlipper.currentangle
Gate1p.Roty = Gate1.CurrentAngle
LeftGatep.Rotx = LeftGate.CurrentAngle
RightSpinnerP.Rotx = sw39.CurrentAngle
RightSpinnerP1.Rotx = sw39.CurrentAngle
LeftGateSwitch.Rotx = LeftGateFl.CurrentAngle
RollingSoundUpdate
BallShadowUpdate
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
NFadeL 20, l20
NFadeL 21, l21
NFadeL 22, l22
NFadeL 23, l23
NFadeL 24, l24
NFadeL 25, l25
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
NFadeL 44, l44
NFadeL 45, l45
NFadeL 46, l46
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

NFadeLm 61, l61
NFadeL 61, l61a
NFadeL 62, l62
NFadeL 63, l63
NFadeL 64, l64
NFadeL 65, l65
NFadeL 66, l66
NFadeL 67, l67
NFadeL 68, l68
FadeObj 69, CenterPost, "3D_Centerpost_3", "3D_Centerpost_2", "3D_Centerpost_1", "3D_Centerpost"
NFadeLm 70, l70
Flashm 70, l70a
Flash 70, l70a1

'NFadeLm 75, GI1
'NFadeLm 75, GI2
'NFadeLm 75, GI3
'NFadeLm 75, GI4
'NFadeLm 75, GI5
'NFadeLm 75, GI6
'NFadeLm 75, GI7
'NFadeLm 75, GI8
'NFadeLm 75, GI9
'NFadeLm 75, GI10
'NFadeLm 75, GI11
'NFadeLm 75, GI12
'NFadeLm 75, GI13
'NFadeLm 75, GI14
'NFadeLm 75, GI15
'NFadeLm 75, GI16
'NFadeLm 75, GI17
'NFadeLm 75, GI18
'NFadeLm 75, GI19
'NFadeLm 75, GI20
'NFadeLm 75, GI21
'NFadeLm 75, GI22
'NFadeLm 75, GI23
'NFadeLm 75, GI24a
'NFadeLm 75, GI25a
'NFadeLm 75, GI26a
'NFadeLm 75, GI24
'NFadeLm 75, GI25
'NFadeLm 75, GI26
'NFadeLm 75, GI27
'NFadeLm 75, GI28
'NFadeLm 75, GI29
'NFadeLm 75, GI30
'NFadeLm 75, GI31
'NFadeLm 75, GI32
'NFadeLm 75, GI33
'NFadeLm 75, GI34
'NFadeLm 75, GI35
'NFadeLm 75, GI36
'NFadeLm 75, GI37
'NFadeLm 75, GI38
'NFadeLm 75, GI39
'NFadeLm 75, GI40
'NFadeLm 75, GI41
'NFadeLm 75, GI42
'NFadeLm 75, GI43
'NFadeLm 75, GI44
'NFadeLm 75, GI45
'NFadeLm 75, GI46
'NFadeLm 75, GI47
'NFadeLm 75, GI48
'NFadeLm 75, GI49
'NFadeLm 75, GI50
'NFadeLm 75, GI51
'NFadeLm 75, GI52
'NFadeLm 75, GI53
'NFadeLm 75, GI54
'NFadeLm 75, GI_spot1
'NFadeLm 75, GI_spot2
'NFadeLm 75, GI_spot3
'NFadeLm 75, GI_spot4
'NFadeL 75, GI_spot5

NFadeL 199, F29
NFadeL 190, F30
NFadeL 191, F31

'Mini Playfield LED Inserts
NFadeL 125, LED1       'F
NFadeL 124, LED2       'I
NFadeL 123, LED3       'O
NFadeL 122, LED4       'N
NFadeL 121, LED5       'A

NFadeL 98, LED6        'M
NFadeL 99, LED7        'A
NFadeL 100, LED8       'N

NFadeL 114, LED9       'S
NFadeL 89, LED10       'H
NFadeL 90, LED11       'R
NFadeL 91, LED12       'E
NFadeL 92, LED13       'K

NFadeL 108, LED14      'P
NFadeL 107, LED15      'U
NFadeL 106, LED16      'S
NFadeL 105, LED17      'S

NFadeL 84, LED18       'C
NFadeL 83, LED19       'H
NFadeL 82, LED20       'A
NFadeL 81, LED21       'R
NFadeL 97, LED22       'M

'Flashers
Flashm 76, Strip1
Flashm 76, Strip2
Flashm 76, Strip3
Flashm 76, Strip4
Flashm 76, Strip5
Flashm 76, Strip6
Flashm 76, Strip7
Flashm 76, Strip8
Flashm 76, Strip9
Flash 76, Strip10


NfadeLm 192, F32
NfadeLm 192, F32a
NfadeLm 192, F32b
Flashm 192, F32c
Flashm 192, F32a1
Flashm 192, F32a2
Flashm 192, F32a3
Flashm 192, F32a4
Flash 192, F32r

NfadeLm 193, F23
NfadeLm 193, F23a
NfadeLm 193, F23b
Flashm 193, F23c
Flashm 193, F23a1
Flashm 193, F23a2
Flashm 193, F23a3
Flash 193, F23r

Flash 195, F25
Flash 196, F26
Flash 197, F27

Flashm 198, F28
Flash 198, F28a
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



'******************************************************
'* GI COLOR *******************************************
'******************************************************
'******************************************************

dim xx
'For each xx in GI:xx.State = 1: Next
select case GIcolor
case 0:
For each xx in GI:xx.color = &h4080FF: Next
case 1:
For each xx in GI:xx.color = &h4080FF:xx.colorfull = &h4080FF: Next
case 2:
GI1.color = &h40FF00:GI1.colorfull = &hFFFFFF
GI2.color = &h40FF00:GI2.colorfull = &hFFFFFF
GI3.color = &h40FF00:GI3.colorfull = &hFFFFFF
GI4.color = &h40FF00:GI4.colorfull = &hFFFFFF
GI5.color = &h40FF00:GI5.colorfull = &hFFFFFF
GI6.color = &h40FF00:GI6.colorfull = &hFFFFFF
GI7.color = &hFF0000:GI7.colorfull = &hFF0000
GI8.color = &hFF0000:GI8.colorfull = &hFF0000
GI9.color = &hFF0000:GI9.colorfull = &hFF0000

GI10.color = &h0000FF:GI10.colorfull = &h0000FF
GI11.color = &h0000FF:GI11.colorfull = &h0000FF
GI12.color = &h0000FF:GI12.colorfull = &h0000FF
GI13.color = &h0000FF:GI13.colorfull = &h0000FF
GI14.color = &h0000FF:GI14.colorfull = &h0000FF
GI15.color = &h0000FF:GI15.colorfull = &h0000FF
GI16.color = &h0000FF:GI16.colorfull = &h0000FF
GI17.color = &h0000FF:GI17.colorfull = &h0000FF
GI18.color = &h0000FF:GI18.colorfull = &h0000FF
GI19.color = &h0000FF:GI19.colorfull = &h0000FF

GI20.color = &hFF0000:GI20.colorfull = &hFF0000
GI21.color = &hFF0000:GI21.colorfull = &hFF0000
GI22.color = &hFF0000:GI22.colorfull = &hFF0000
GI23.color = &hFF0000:GI23.colorfull = &hFF0000
end select

'*****************************************
'	Ball Shadow
'*****************************************

Dim BallShadow
BallShadow = Array (BallShadow1, BallShadow2, BallShadow3)

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

		If BOT(b).Z < 10 and BOT(b).Z>90 then BallShadow(b).Z = 1 else BallShadow(b).Z=BOT(b).z-20

			BallShadow(b).Y = BOT(b).Y + 40
		If BOT(b).Z > 20 Then
			BallShadow(b).visible = 1
		Else
			BallShadow(b).visible = 0
		End If
	Next
End Sub

Sub DropTargets_Hit (idx)
	Playsound SoundFX("Stern_Droptargethit",DOFContactors), 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0, AudioFade(ActiveBall)
End Sub

Sub Gates_Hit (idx)
	PlaySound "fx_gate", 0, Vol(ActiveBall)*VolGates, Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
End Sub

Sub Rubbers_Hit(idx)
 	dim finalspeed
  	finalspeed=SQR(activeball.velx * activeball.velx + activeball.vely * activeball.vely)
 	If finalspeed > 20 then
		PlaySound "fx_rubber2", 0, Vol(ActiveBall)*VolRH, Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
	End if
	If finalspeed >= 6 AND finalspeed <= 20 then
 		RandomSoundRubber()
 	End If
End Sub

Sub Posts_Hit(idx)
 	dim finalspeed
  	finalspeed=SQR(activeball.velx * activeball.velx + activeball.vely * activeball.vely)
 	If finalspeed > 16 then
		PlaySound "fx_rubber2", 0, Vol(ActiveBall)*VolPo, Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
	End if
	If finalspeed >= 6 AND finalspeed <= 16 then
 		RandomSoundRubber()
 	End If
End Sub

Sub RandomSoundRubber()
	Select Case Int(Rnd*3)+1
		Case 1 : PlaySound "rubber_hit_1", 0, Vol(ActiveBall)*VolRH, Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
		Case 2 : PlaySound "rubber_hit_2", 0, Vol(ActiveBall)*VolRH, Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
		Case 3 : PlaySound "rubber_hit_3", 0, Vol(ActiveBall)*VolRH, Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
	End Select
End Sub

Sub LeftFlipper_Collide(parm)
Dim Ballspeed
Ballspeed=SQR(activeball.velx * activeball.velx + activeball.vely * activeball.vely)
If Ballspeed >5 then PlaysoundAtVol "Stern_Flippercollide1", LeftFlipper, VolFlip Else PlaysoundAtVol "Stern_FlipperCollideLow",LeftFlipper, VolFlip:End If:End Sub

Sub RightFlipper_Collide(parm)
Dim Ballspeed
Ballspeed=SQR(activeball.velx * activeball.velx + activeball.vely * activeball.vely)
If Ballspeed >5 then PlaysoundAtVol "Stern_Flippercollide1",RightFlipper,VolFlip Else PlaysoundAtVol "Stern_FlipperCollideLow",RightFlipper, VolFlip:End If:End Sub

Sub WirerollSND_Hit:PlaysoundAtVol"Stern_Wireroll",ActiveBall,1:End Sub
Sub BallDropSound_Hit:Stopsound"Stern_Wireroll":Playsound"Stern_Balldrop":End Sub

''*************DEDUG*******************************************************
'
'Sub Table1_KeyDown(ByVal keycode)
'
'If keycode = 33 then sw55.createsizedball(14.6875) : sw55.kick 90, 30
'
'	If keycode = PlungerKey Then
'		Plunger.PullBack
'		PlaySound "plungerpull",0,1,0.25,0.25
'	End If
'
'	If keycode = LeftFlipperKey Then
'		MiniPF_LeftFlipper.RotateToEnd
'		LeftFlipper.RotateToEnd
'LeftFlipperSmall.RotateToEnd
'		PlaySound "fx_flipperup", 0, .67, -0.05, 0.05
'	End If
'
'	If keycode = RightFlipperKey Then
'		MiniPF_RightFlipper.RotateToEnd
'		RightFlipper.RotateToEnd
'		PlaySound "fx_flipperup", 0, .67, 0.05, 0.05
'	End If
'
'	If keycode = LeftTiltKey Then
'		Nudge 90, 2
'	End If
'
'	If keycode = RightTiltKey Then
'		Nudge 270, 2
'	End If
'
'	If keycode = CenterTiltKey Then
'		Nudge 0, 2
'	End If
'
'End Sub
'
'Sub Table1_KeyUp(ByVal keycode)
'
'	If keycode = PlungerKey Then
'		Plunger.Fire
'		PlaySound "plunger",0,1,0.25,0.25
'	End If
'
'	If keycode = LeftFlipperKey Then
'		MiniPF_LeftFlipper.RotateToStart
'		LeftFlipper.RotateToStart
'		LeftFlipperSmall.RotateToStart
'		PlaySound "fx_flipperdown", 0, 1, -0.05, 0.05
'	End If
'
'	If keycode = RightFlipperKey Then
'		MiniPF_RightFlipper.RotateToStart
'		RightFlipper.RotateToStart
'		PlaySound "fx_flipperdown", 0, 1, 0.05, 0.05
'	End If
'
'End Sub
'
'Sub Drain_Hit()
'	PlaySound "drain",0,1,0,0.25
'	Drain.DestroyBall
'	BIP = BIP - 1
'	If BIP = 0 then
'		'Plunger.CreateBall
'		BallRelease.CreateBall
'		BallRelease.Kick 90, 7
'		PlaySound "ballrelease",0,1,0,0.25
'		BIP = BIP + 1
'	End If
'End Sub
'
'Dim BIP
'BIP = 0
'
'Sub Plunger_Init()
'	PlaySound "ballrelease",0,1,0,0.25
'	'Plunger.CreateBall
'	BallRelease.CreateBall
'	BallRelease.Kick 90, 7
'	BIP = BIP +1
'End Sub
'
'
'
'Sub sw55_Hit
'
'sw55.kick 90,35
'
'End Sub
'
'
'Sub TiDebug_timer
'RightFlipperP.Rotz = RightFlipper.CurrentAngle
'LeftFlipperP.Rotz = LeftFlipper.CurrentAngle
'LeftFlipperSmallP.RotY = LeftFlipperSmall.CurrentAngle
'MiniPF_LeftFlipperP.Rotz = MiniPF_LeftFlipper.CurrentAngle
'MiniPF_RightFlipperP.Rotz = MiniPF_RightFlipper.CurrentAngle
'FlipperLSh.RotZ = LeftFlipper.currentangle
'FlipperRSh.RotZ = RightFlipper.currentangle
'End Sub
'
''*************DEDUG*******************************************************

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

Sub PlaySoundAtVol(sound, tableobj, Volum)
  PlaySound sound, 1, Volum, Pan(tableobj), 0,0,0, 1, AudioFade(tableobj)
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
  Vol = Csng(BallVel(ball) ^2 / VolDiv)
End Function

Function Pitch(ball) ' Calculates the pitch of the sound based on the ball speed
  Pitch = BallVel(ball) * 20
End Function

Function BallVel(ball) 'Calculates the ball speed
  BallVel = INT(SQR((ball.VelX ^2) + (ball.VelY ^2) ) )
End Function

'*****************************************
'    JP's VP10 Collision & Rolling Sounds
'*****************************************

Const tnob = 9 ' total number of balls

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
    PlaySound("fx_collide"), 0, Csng(velocity) ^2 / (VolDiv/VolCol), Pan(ball1), 0, Pitch(ball1), 0, 0, AudioFade(ball1)
End Sub


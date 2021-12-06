'*************************************************************************************************************************************************************
'FISH TALES
'Williams 1992
'version 1.1
'VPX SS recreation by pinball58
'Thanks to the authors(PacDude,Melon,Loaded Weapon,ICPjuggla,Zany)who made this table before for the stuff and ideas that I borrowed from their VP9 tables
'(especially Zany for primitives)
'Thanks to Tom Tower and Ninuzzu for helping me finalize the table
'Thanks to Arngrim for helping me with DOF
'Thanks to VPDev Team for the freaking amazing VPX
'*************************************************************************************************************************************************************

' Thalamus 2018-07-24
' Table has its own "Positional Sound Playback Functions" and "Supporting Ball & Sound Functions"
' Changed UseSolenoids=1 to 2
' No special SSF tweaks yet.

Option Explicit
Randomize

Const Ballsize= 50
Const BallMass = 1

On Error Resume Next
ExecuteGlobal GetTextFile("controller.vbs")
If Err Then MsgBox "You need the controller.vbs in order to run this table, available in the vp10 package"
On Error Goto 0

Dim DesktopMode:DesktopMode = FishTales.ShowDT
Dim UseVPMDMD:UseVPMDMD = DesktopMode

Dim bsTrough,bsFishFinder,bsVUK,bsReel1,bsReel2,bsReel3,bsCatapult,dtDrop,RampDecals,FlippersColor,FlipColor,DMDColor,DMDCol

LoadVPM "02700000", "WPC.VBS", 3.52

'****************************** TABLE OPTIONS ********************************************************************************************************
'*****************************************************************************************************************************************************

RampDecals = 1                        'original(no decals) = 0    Ramp with decals = 1

FlippersColor = 0                       'original(Red Rubber) = 0    Green Rubber = 1    Black Rubber = 2    Random Rubber Color = 3

DMDColor = 0 '(only for DesktopMode)    'Orange = 0    Red = 1    Green = 2    Blue = 3    Random DMD Color = 4

'*****************************************************************************************************************************************************
'*****************************************************************************************************************************************************

'*********** Standard definitions ****************

 Const UseSolenoids = 2
 Const UseLamps = 0
 Const UseSync = 0
 Const UseGI = 1

'Standard Sounds
 Const SSolenoidOn = "solenoid"
 Const SSolenoidOff = ""
 Const SFlipperOn = ""
 Const SFlipperOff = ""
 Const SCoin = "CoinIn"

'Rom name
 Const cGameName = "ft_l5"

'*************************************************

'************ Fish Tales Init *****************

Sub FishTales_Init
  vpmInit me
  With Controller
    .GameName = cGameName
        .SplashInfoLine = "Fish Tales - Williams 1992"
     If Err Then MsgBox "Can't start Game" & cGameName & vbNewLine & Err.Description:Exit Sub
        .ShowTitle = 0
        .ShowDMDOnly = 1
        .ShowFrame = 0
    .HandleKeyboard = 0
        .HandleMechanics = 2
        .Hidden = DesktopMode
        On Error Resume Next
        .Run GetPlayerHWnd
        If Err Then MsgBox Err.Description
        On Error Goto 0
  End With

'Nudging
  vpmNudge.TiltSwitch = 14
  vpmNudge.Sensitivity = 4
    vpmNudge.TiltObj = Array(Bumper1, Bumper2, Bumper3, LeftSlingShot, RightSlingShot)

'Trough
Set bsTrough = New cvpmTrough
    With bsTrough
    .Size = 3
    .InitSwitches Array(18, 17, 16)
    .EntrySw = 15
    .InitExit BallRelease, 90, 7
    .InitEntrySounds "drain", SoundFX(SSolenoidOn,DOFContactors),SoundFX("",DOFContactors)
    .InitExitSounds  SoundFX(SSolenoidOn,DOFContactors),SoundFX("ballrelease",DOFContactors)
    .Balls = 3
    .CreateEvents "bsTrough", Drain
    End With

'Fish Finder Kicker
Set bsFishFinder=New cvpmSaucer
  With bsFishFinder
    .InitKicker sw63, 63, 270, 10, 0
    .InitSounds "kicker_enter_center",SoundFX(SSolenoidOn,DOFContactors),SoundFX("kicker_kick",DOFContactors)
    .CreateEvents "bsFishFinder", sw63
  End With

'Caster Club Vertical Kicker
Set bsVUK=New cvpmSaucer
  With bsVUK
    .InitKicker sw47, 47, 0, 45, 1.56
    .InitSounds "kicker_enter_center",SoundFX(SSolenoidOn,DOFContactors),SoundFX("vuk_exit",DOFContactors)
        .CreateEvents "bsVUK", Sw47
  End With

'Catapult
Set bsCatapult = new cvpmSaucer
  With bsCatapult
    .InitKicker Catapult, 36, 0, 50, 40
    .InitSounds "catapult_in",SoundFX("diverter",DOFContactors),SoundFX("catapult_fire",DOFContactors)
        .CreateEvents "bsCatapult", Catapult
  End With

'Reel Slot 1
Set bsReel1 = new cvpmTrough
  With BsReel1
    .Size = 1
    .InitSwitches Array(0)
    .InitExit ReelExit, 250, 2
    .Balls = 0
  End With

'Reel Slot 2
Set bsReel2 = new cvpmTrough
  With BsReel2
    .Size = 1
    .InitSwitches Array(0)
    .InitExit ReelExit, 250, 2
    .Balls = 0
  End With

'Reel Slot 3
Set bsReel3 = new cvpmTrough
  With BsReel3
    .Size = 1
    .InitSwitches Array(0)
    .InitExit ReelExit, 250, 2
    .Balls = 0
  End With

'Caster Club Drop Target
 Set dtDrop=New cvpmDropTarget
  With dtDrop
      .InitDrop sw48, 48
      .InitSnd SoundFX("droptarget",DOFDropTargets),SoundFX("resetdrop",DOFContactors)
  End With

'Boat Captive Ball
  Kicker1.createsizedballwithmass ballsize/2, ballmass:Kicker1.kick 0,1:Kicker1.enabled=0

'Init other stuff
  InitOptions
  MoveGate=True
  Ramp15.visible=DesktopMode:Ramp16.visible=DesktopMode
  f27Fs.Visible = Not DesktopMode
  If Not DesktopMode Then FsSetup
End Sub

'*************************************************

'********* Flippers *************

SolCallback(sLLFlipper) = "SolLFlipper"
SolCallback(sLRFlipper) = "SolRFlipper"

Sub SolLFlipper(Enabled)
     If Enabled Then
     PlaySoundAtVol SoundFX("FlipperUp",DOFFlippers),LeftFlipper,.6
     LeftFlipper.RotateToEnd
     Else
     PlaySoundAtVol SoundFX("FlipperDown",DOFFlippers),LeftFlipper,.3
     LeftFlipper.RotateToStart
     End If
 End Sub

Sub SolRFlipper(Enabled)
  If enabled Then
     PlaySoundAtVol SoundFX("FlipperUp",DOFFlippers),RightFlipper,.6
     RightFlipper.RotateToEnd
     Else
     PlaySoundAtVol SoundFX("FlipperDown",DOFFlippers),RightFlipper,.3
     RightFlipper.RotateToStart
     End If
 End Sub

'*******************************

'************** Keys *************************

Sub FishTales_KeyDown(ByVal keycode)
If keycode = PlungerKey Then Controller.Switch(31) = 1
If keycode = LeftTiltKey Then PlaySound SoundFX("fx_nudge",0)
If keycode = RightTiltKey Then PlaySound SoundFX("fx_nudge",0)
If keycode = CenterTiltKey Then PlaySound SoundFX("fx_nudge",0)
If vpmKeyDown(keycode) Then Exit Sub
End Sub

Sub FishTales_KeyUp(ByVal keycode)
If keycode = PlungerKey Then Controller.Switch(31) = 0
If vpmKeyUp(keycode) Then Exit Sub
End Sub

'*********************************************

'********* Solenoids ************

SolCallback(1)  = "Auto_Plunger"
SolCallback(2)  = "SolCatapult"
SolCallback(3)  = "bsVUK.SolOut"
SolCallback(6)  =   "vpmSolgate Gate,SoundFX(""diverter"", DOFContactors),"
SolCallback(7)  = "vpmSolSound SoundFX(""knocker"",DOFKnocker),"
SolCallback(9)  =   "bsTrough.SolIn"
SolCallback(10) =   "bsTrough.SolOut"
SolCallback(11) =   "bsFishFinder.SolOut"
SolCallback(12) = "dtDrop.SolDropUp"
SolCallback(13) = "dtDrop.SolDropDown"
SolCallback(28) =   "ReelMotor"
SolCallback(51) =   "bsReel1.SolOut"
SolCallback(52) =   "bsReel2.SolOut"
SolCallback(53) =   "bsReel3.SolOut"

'Flasher Solenoids

SolCallback(17) = "setlamp 117,"  'Jackpot Flasher
SolCallback(18) = "setlamp 118,"  'Super Jackpot Flasher
SolCallback(19) = "setlamp 119,"  'Instant Multiball Flasher
SolCallback(20) = "setlamp 120,"  'Light Extraball Flasher
SolCallback(21) = "setlamp 121,"  'Rock the Boat Flasher
SolCallback(22) = "setlamp 122,"  'Video Mode Flasher
SolCallback(23) = "setlamp 123,"  'Hold Bonus Flasher
SolCallback(25) = "setlamp 125,"  'Reel Flasher
SolCallback(26) = "setlamp 126,"  'Top Left Flasher
SolCallback(27) = "setlamp 127,"  'Caster Club Flasher

'********************************************

'********************** Auto Plunger ****************************

Sub Auto_Plunger(Enabled)
     If Enabled Then
         Plunger.Fire
   PlaySoundAt SoundFX("fx_AutoPlunger",DOFContactors),Plunger
     End If
 End Sub

'****************************************************************

'********* Catapult ***********

Dim catdir

Sub SolCatapult(enabled)
  If enabled Then
    catdir=1
    CatapultTimer.enabled=1
    bsCatapult.ExitSol_On
  End If
End Sub

Sub CatapultTimer_timer()
Primitive98.RotX=Primitive98.RotX+catdir
If Primitive98.RotX>=90 And catdir=1 Then catdir=-1
If Primitive98.RotX<=1 Then CatapultTimer.enabled=0
End Sub

'*******************************

'********************** Fishing Reel Motor ****************************

Dim ReelPosition,PosInitial

Sub ReelMotor(enabled)
  If enabled Then
    ReelRotation
    PlaySoundAtVol SoundFX("motor_on",DOFGear),ReelEnter,.05
  Else
    StopSound "motor_on"
  End If
End Sub

Sub ReelRotation
    ReelTimer.enabled=1
  ReelPosition=0
  PosInitial=Reel.RotX
End Sub

Sub ReelTimer_Timer()
  Reel.RotX=PosInitial+ReelPosition
  ReelPosition=ReelPosition+10
  If ReelPosition>120 Then
    ReelPosition=0
    ReelTimer.enabled=0
  If Reel.RotX=300 Then Reel.RotX=-60
  End If
End Sub

Sub ReelEnter_Hit()
  Dim nHole
    Stopsound "metal"
  PlaySoundAt "hop2",ReelEnter
  Me.DestroyBall
  nHole = Controller.GetMech(1) 'Return which Ball Lock is UP!
  'Which hole is UP?
  Select Case nHole
    Case 1
      bsReel1.AddBall 0
    Case 2
      bsReel2.AddBall 0
    Case 3
      bsReel3.AddBall 0
    Case Else
         'No hole up, we'll call this method again in 500 ms, and try again!
         VPMPulseTimer.AddTimer 500,"ReelEnter_Hit '"
  End Select
End Sub

'*********************************************************************

'****************** Switches *********************

Sub sw25_Hit():VPMTimer.PulseSw 25:PlaySound "rollover",0,1,Pan(ActiveBall),0,0,1,0,AudioFade(ActiveBall):End Sub 'Left OutLane
Sub sw26_Hit():VPMTimer.PulseSw 26:PlaySound "rollover",0,1,Pan(ActiveBall),0,0,1,0,AudioFade(ActiveBall):End Sub 'Left InLane
Sub sw27_Hit():VPMTimer.PulseSw 27:PlaySound SoundFX("target",DOFTargets),0,1,Pan(ActiveBall),0,0,1,0,AudioFade(ActiveBall):End Sub 'Standup Target
Sub sw28_Hit():VPMTimer.PulseSw 28:PlaySound SoundFX("target",DOFTargets),0,1,Pan(ActiveBall),0,0,1,0,AudioFade(ActiveBall):End Sub 'Standup Target
Sub sw32_Hit():VPMTimer.PulseSw 32:Primitive147.RotX=67:PlaySound "target",0,1,Pan(ActiveBall),0,0,1,0,AudioFade(ActiveBall):Me.TimerInterval=150:Me.Timerenabled=1:End Sub 'Ramp Left Sensor
Sub sw32_timer():Me.Timerenabled=0:Primitive147.RotX=80:End Sub
Sub sw33_Hit():VPMTimer.PulseSw 33:Primitive148.RotX=67:PlaySound "target",0,1,Pan(ActiveBall),0,0,1,0,AudioFade(ActiveBall):Me.TimerInterval=150:Me.Timerenabled=1:End Sub 'Ramp Right Sensor
Sub sw33_timer():Me.Timerenabled=0:Primitive148.RotX=80:End Sub
Sub sw34_Spin():VPMTimer.PulseSw 34:End Sub 'Spinner
Sub sw41_Hit():VPMTimer.PulseSw 41:PlaySound SoundFX("target",DOFTargets),0,1,Pan(ActiveBall),0,0,1,0,AudioFade(ActiveBall):End Sub 'Target Boat Captive Ball
Sub sw42_Hit():VPMTimer.PulseSw 42:PlaySound "metalhit_thin",0,1,Pan(ActiveBall),0,0,1,0,AudioFade(ActiveBall):End Sub 'Boat Ramp Right Trigger
Sub sw35_Hit():Controller.Switch(35) =1:End Sub 'Reel Entry Trigger
Sub sw35_UnHit():Controller.Switch(35)=0:End Sub
Sub sw43_Hit():VPMTimer.PulseSw 43:PlaySound "metalhit_thin",0,1,Pan(ActiveBall),0,0,1,0,AudioFade(ActiveBall):End Sub 'Boat Ramp Left Trigger
Sub sw44_Hit():VPMTimer.PulseSw 44:PlaySound "rollover",0,1,Pan(ActiveBall),0,0,1,0,AudioFade(ActiveBall):End Sub 'E Trigger
Sub sw45_Hit():VPMTimer.PulseSw 45:PlaySound "rollover",0,1,Pan(ActiveBall),0,0,1,0,AudioFade(ActiveBall):End Sub 'I Trigger
Sub sw46_Hit():VPMTimer.PulseSw 46:PlaySound "rollover",0,1,Pan(ActiveBall),0,0,1,0,AudioFade(ActiveBall):End Sub 'L Trigger
Sub sw47_UnHit():Primitive146.TransY=20:Me.TimerInterval=200:Me.TimerEnabled=1:End Sub
Sub sw47_timer():Me.TimerEnabled=0:Primitive146.TransY=0:End Sub
Sub sw48_dropped():dtDrop.Hit 1:End Sub 'Drop Target
Sub sw54_Hit():VPMTimer.PulseSw 54:PlaySound SoundFX("target",DOFTargets),0,1,Pan(ActiveBall),0,0,1,0,AudioFade(ActiveBall):End Sub 'Standup Target
Sub sw55_Hit():VPMTimer.PulseSw 55:PlaySound SoundFX("target",DOFTargets),0,1,Pan(ActiveBall),0,0,1,0,AudioFade(ActiveBall):End Sub 'Standup Target
Sub sw56_Hit():Controller.Switch (56)=1:PlaySound "rollover",0,1,Pan(ActiveBall),0,0,1,0,AudioFade(ActiveBall):End Sub 'Shooter Lane Trigger
Sub sw56_UnHit():Controller.Switch (56)=0:End Sub
Sub sw61_Hit():VPMTimer.PulseSW 61:PlaySound SoundFX("target",DOFTargets),0,1,Pan(ActiveBall),0,0,1,0,AudioFade(ActiveBall):End Sub 'Oblong Target
Sub sw62_Hit():VPMTimer.PulseSw 62:PlaySound "metal_thin",0,1,Pan(ActiveBall),0,0,1,0,AudioFade(ActiveBall):MoveGate=True:End Sub 'Right Green Lane Trigger
Sub sw64_Hit():VPMTimer.PulseSw 64:PlaySound "rollover",0,1,Pan(ActiveBall),0,0,1,0,AudioFade(ActiveBall):End Sub 'Left Green Lane Trigger
Sub sw65_Hit():VPMTimer.PulseSw 65:PlaySound "rollover",0,1,Pan(ActiveBall):End Sub 'Right InLane
Sub sw66_Hit():VPMTimer.PulseSw 66:PlaySound "rollover",0,1,Pan(ActiveBall),0,0,1,0,AudioFade(ActiveBall):End Sub 'Right OutLane

'**************************************************

'********** Bumpers **********************************************

Sub Bumper1_Hit
VPMTimer.PulseSw 51
  Dim BumpSound
  BumpSound = Int(rnd*3)+1
  Select Case BumpSound
  Case 1: PlaySoundAtBumperVol SoundFX("fx_bumper1",DOFContactors),Bumper1,2
  Case 2: PlaySoundAtBumperVol SoundFX("fx_bumper2",DOFContactors),Bumper2,2
  Case 3: PlaySoundAtBumperVol SoundFX("fx_bumper3",DOFContactors),Bumper3,2
  End Select
End Sub

Sub Bumper2_Hit
VPMTimer.PulseSw 52
  Dim BumpSound
  BumpSound = Int(rnd*3)+1
  Select Case BumpSound
  Case 1: PlaySoundAtBumperVol SoundFX("fx_bumper1",DOFContactors),Bumper1,2
  Case 2: PlaySoundAtBumperVol SoundFX("fx_bumper2",DOFContactors),Bumper2,2
  Case 3: PlaySoundAtBumperVol SoundFX("fx_bumper3",DOFContactors),Bumper3,2
  End Select
End Sub

Sub Bumper3_Hit
VPMTimer.PulseSw 53
  Dim BumpSound
  BumpSound = Int(rnd*3)+1
  Select Case BumpSound
  Case 1: PlaySoundAtBumperVol SoundFX("fx_bumper1",DOFContactors),Bumper1,2
  Case 2: PlaySoundAtBumperVol SoundFX("fx_bumper2",DOFContactors),Bumper2,2
  Case 3: PlaySoundAtBumperVol SoundFX("fx_bumper3",DOFContactors),Bumper3,2
  End Select
End Sub

'*****************************************************************

'********** Slingshots ***************

Dim Lstep,RStep

Sub LeftSlingShot_Slingshot
  VPMTimer.PulseSw 57
    PlaySoundAt SoundFX("LSling",DOFContactors),SLING2
    LSling.Visible = 0
    LSling1.Visible = 1
    sling2.TransZ = -32
    LStep = 0
    LeftSlingShot.TimerEnabled = 1
End Sub

Sub LeftSlingShot_Timer
    Select Case LStep
        Case 3:LSLing1.Visible = 0:LSLing2.Visible = 1:sling2.TransZ = -17
        Case 4:LSLing2.Visible = 0:LSLing.Visible = 1:sling2.TransZ = 0:LeftSlingShot.TimerEnabled = 0
    End Select
    LStep = LStep + 1

End Sub

Sub RightSlingShot_Slingshot
  VPMTimer.PulseSw 58
    PlaySoundAt SoundFX("RSling",DOFContactors), SLING1
    RSling.Visible = 0
    RSling1.Visible = 1
    sling1.TransZ = -32
    RStep = 0
    RightSlingShot.TimerEnabled = 1
End Sub

Sub RightSlingShot_Timer
    Select Case RStep
       Case 3:RSLing1.Visible = 0:RSLing2.Visible = 1:sling1.TransZ = -17
       Case 4:RSLing2.Visible = 0:RSLing.Visible = 1:sling1.TransZ = 0:RightSlingShot.TimerEnabled = 0
    End Select
    RStep = RStep + 1
End Sub

'*******************************************

'********** In Game Updates ************

Sub IGUpdates_timer()
  Reel1.RotX=Reel.RotX
  If MoveGate=True Then
    Primitive100.ObjRotY=Gate.CurrentAngle*(-1)
    If Gate.CurrentAngle>=75 Then Primitive100.ObjRotY=-75 End If
    If Gate.CurrentAngle<=0 Then Primitive100.ObjRotY=-15 End If
    If Gate.Open=True Then
      Wall135.IsDropped=True
    Else
      Wall135.IsDropped=False
    End If
  End If
  Primitive109.RotY=LeftFlipper.CurrentAngle-90
  Primitive110.RotY=RightFlipper.CurrentAngle+90
End Sub

Dim MoveGate,verso,verso2

Sub Gate_Hit()
  PlaySound "gate",0,1,Pan(ActiveBall),0,0,1,0,AudioFade(ActiveBall)
  MoveGate=False
  Wall135.IsDropped=True
  GateAnim.Enabled=1
  verso=-1
End Sub

Sub GateAnim_timer()
  Primitive100.ObjRotY=Primitive100.ObjRotY+verso
  If Primitive100.ObjRotY<=-60 Then verso=1
  If Primitive100.ObjRotY>=-15 Then
  GateAnim.enabled=0
  Primitive100.ObjRotY=-15
  MoveGate=True
  End If
End Sub

Sub Wall135_Hit()
  PlaySound "metalhit_medium",0,1,Pan(ActiveBall),0,0,1,0,AudioFade(ActiveBall)
  MoveGate=False
  GateAnim2.enabled=1
  verso2=1
End Sub

Sub GateAnim2_timer()
  Primitive100.ObjRotY=Primitive100.ObjRotY+verso
  If Primitive100.ObjRotY>=0 Then verso=-1
  If Primitive100.ObjRotY<=-15 Then
  GateAnim2.enabled=0
  Primitive100.ObjRotY=-15
  MoveGate=True
  End If
End Sub

'*****************************************************

'*************** General Illumination *****************

 Set GiCallback = GetRef("UpdateGI")
 Set GiCallback2 = GetRef("UpdateGI2")

Dim gistep, xx

Sub UpdateGI(nr, enabled)
  Select Case nr
        Case 2  'Top GI
            For xx = 0 to 30:TopGI(xx).state=enabled:next
            For xx = 31 to 39:TopGI(xx).visible=enabled:next
        Case 4  'Bottom GI
            For each xx in BottomGI:xx.state=enabled:next
    End Select
End Sub

Sub UpdateGI2(no, step)
    If step=0 Then exit Sub
    gistep=(step-1)/7

  If gistep = 1 Then
    DOF 101, DOFOn
  Else
    DOF 101, DOFOff
  End If

    Select Case no
        Case 2  'Top GI
            For each xx in TopGI:xx.IntensityScale=gistep:next
        Case 4  'Bottom GI
            For each xx in BottomGI:xx.IntensityScale=gistep:next
    End Select
  FishTales.ColorGradeImage = "ColorGrade_" & step
End Sub

'*****************************************************

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
LampTimer.Interval = 200  'lamp fading speed
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
    'NFadeL 1, l1
    'NFadeL 2, l2
    'NFadeL 3, l3
    'NFadeL 4, l4
    'NFadeL 5, l5
    'NFadeL 6, l6
    'NFadeL 7, l7
    'NFadeL 8, l8
    'NFadeL 9, l9
    'NFadeL 10, l10
  NFadeLm 11, l11a
    FadeObj 11, l11, "mpfl4", "mpfl3", "mpfl2", "mpfl1"
  NFadeLm 12, l12a
    FadeObj 12, l12, "mpfl4", "mpfl3", "mpfl2", "mpfl1"
  NFadeLm 13, l13a
    FadeObj 13, l13, "mpfl4", "mpfl3", "mpfl2", "mpfl1"
  NFadeLm 14, l14a
    FadeObj 14, l14, "mpfl4", "mpfl3", "mpfl2", "mpfl1"
  NFadeLm 15, l15a
    FadeObj 15, l15, "mpfl4", "mpfl3", "mpfl2", "mpfl1"
    NFadeLm 16, l16
  Flash 16, l16f
    NFadeLm 17, l17
  Flash 17, l17f
    NFadeLm 18, l18
  Flash 18, l18f
    'NFadeL 19, l19
    'NFadeL 20, l20
    NFadeL 21, l21
    NFadeL 22, l22
    NFadeL 23, l23
    NFadeL 24, l24
    NFadeL 25, l25
  NFadeL 26, l26
    NFadeL 27, l27
    NFadeL 28, l28
    'NFadeL 29, l29
    'NFadeL 30, l30
    NFadeL 31, l31
    NFadeL 32, l32
    NFadeL 33, l33
    NFadeL 34, l34
  NFadeLm 35, l35a
    FadeObj 35, l35, "mpfl4", "mpfl3", "mpfl2", "mpfl1"
  NFadeLm 36, l36a
    FadeObj 36, l36, "mpfl4", "mpfl3", "mpfl2", "mpfl1"
  NFadeLm 37, l37a
    FadeObj 37, l37, "mpfl4", "mpfl3", "mpfl2", "mpfl1"
  NFadeLm 38, l38a
    FadeObj 38, l38, "mpfl4", "mpfl3", "mpfl2", "mpfl1"
    'NFadeL 39, l39
    'NFadeL 40, l40
  NFadeL 41, l41
    NFadeL 42, l42
    NFadeL 43, l43
    NFadeL 44, l44
    NFadeL 45, l45
    NFadeL 46, l46
    NFadeL 47, l47
    NFadeLm 48, l48
  NFadeL 48, l48a
    'NFadeL 49, l49
    'NFadeL 50, l50
    NFadeL 51, l51
    NFadeL 52, l52
    NFadeL 53, l53
    NFadeL 54, l54
    NFadeL 55, l55
    NFadeL 56, l56
    NFadeL 57, l57
    NFadeL 58, l58
    'NFadeL 59, l59
    'NFadeL 60, l60
    NFadeL 61, l61
    NFadeL 62, l62
    NFadeL 63, l63
    NFadeL 64, l64
  NFadeL 65, l65
  NFadeL 66, l66
  NFadeL 67, l67
  NFadeL 68, l68
  'NFadeL 69, l69
  'NFadeL 70, l70
  Flash 71, l71
  NFadeL 72, l72
  NFadeL 73, l73
  NFadeL 74, l74
  NFadeL 75, l75
  NFadeL 76, l76
  NFadeL 77, l77
  NFadeL 78, l78
  'NFadeL 79, l79
  'NFadeL 80, l80
  Flashm 81, l81
  Flash 81, l81b
  Flashm 82, l82
  Flash 82, l82b
  Flashm 83, l83
  Flash 83, l83b
  Flashm 84, l84
  Flash 84, l84b
  Flashm 85, l85
  Flash 85, l85b
  Flash 86, l86

'Flashers
  Flash 117, f17
  Flash 118, f18
  Flash 119, f19
  Flash 120, f20
  Flash 121, f21
  Flash 122, f22
  Flash 123, f23

  Flashm 125, f25
  Flashm 125, f25a
  Flashm 125, f25b
  Flash 125, f25c

  Flashm 126, f26
  Flash 126, f26b

  Flashm 127, f27
  Flashm 127, f27b
  Flash 127, f27Fs

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

'*******************************************************

' *********************************************************************
'                      Supporting Ball & Sound Functions
' *********************************************************************

Function Vol(ball) ' Calculates the Volume of the sound based on the ball speed
    Vol = Csng(BallVel(ball) ^2 / 1500)
End Function

Function Pan(ball) ' Calculates the pan for a ball based on the X position on the table. "table1" is the name of the table
    Dim tmp
    tmp = ball.x * 2 / FishTales.width-1
    If tmp > 0 Then
        Pan = Csng(tmp ^10)
    Else
        Pan = Csng(-((- tmp) ^10) )
    End If
End Function

function AudioFade(ball)
  Dim tmp
    tmp = ball.y * 2 / FishTales.height-1
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

'*****************************************
'      JP's VP10 Rolling Sounds
'*****************************************

Const tnob = 4 ' total number of balls
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
        If BallVel(BOT(b) ) > 1 Then
      rolling(b) = True
      if BOT(b).z < 30 Then ' Ball on playfield
        If FishTales.VersionMinor > 3 OR FishTales.VersionMajor > 10 Then
            PlaySound("fx_ballrolling" & b), -1, Vol(BOT(b) ), Pan(BOT(b) ), 0, Pitch(BOT(b) ), 1, 0, AudioFade(BOT(b) )
        Else
            PlaySound("fx_ballrolling" & b), -1, Vol(BOT(b) ), Pan(BOT(b) ), 0, Pitch(BOT(b) ), 1, 0
        End If
      Else ' Ball on raised ramp
        If FishTales.VersionMinor > 3 OR FishTales.VersionMajor > 10 Then
            PlaySound("fx_ballrolling" & b), -1, Vol(BOT(b) )*.4, Pan(BOT(b) ), 0, Pitch(BOT(b) )+50000, 1, 0, AudioFade(BOT(b) )
        Else
            PlaySound("fx_ballrolling" & b), -1, Vol(BOT(b) )*.4, Pan(BOT(b) ), 0, Pitch(BOT(b) )+50000, 1, 0
        End If
      End If
    Else
            If rolling(b) = True Then
                StopSound("fx_ballrolling" & b)
                rolling(b) = False
            End If
        End If
    Next
End Sub

'*****************************************
'     FLIPPER SHADOWS
'*****************************************

sub FlipperTimer_Timer()
  FlipperLSh.RotZ = LeftFlipper.currentangle
  FlipperRSh.RotZ = RightFlipper.currentangle

End Sub

'*****************************************
'     Ball Shadow
'*****************************************

Dim BallShadow
BallShadow = Array (Ballshadow1,Ballshadow2,Ballshadow3,Ballshadow4)  'one shadow primitive for each ball (tnob=4); change the array depending on tnob

Sub BallShadowUpdate_Timer()
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
    If BOT(b).X < FishTales.Width/2 Then
      BallShadow(b).X = ((BOT(b).X) - (Ballsize/6) + ((BOT(b).X - (FishTales.Width/2))/7)) + 10
    Else
      BallShadow(b).X = ((BOT(b).X) + (Ballsize/6) + ((BOT(b).X - (FishTales.Width/2))/7)) - 10
    End If
      ballShadow(b).Y = BOT(b).Y + 20
    If BOT(b).Z > 20 Then
      BallShadow(b).visible = 1
    Else
      BallShadow(b).visible = 0
    End If
  Next
End Sub

'**********************
' Ball Collision Sound
'**********************

Sub OnBallBallCollision(ball1, ball2, velocity)
    PlaySound("fx_collide"), 0, Csng(velocity) ^2 / 100, Pan(ball1), 0, Pitch(ball1), 0, 0, AudioFade(ball1)
End Sub

'***********************************************************************************************

'******************* Others Table Sounds *************************


Sub Rubbers_Hit(idx)
    PlaySoundAtBallVol "rubber_hit_" & Int(Rnd*3)+1, .1
End Sub

Sub LeftFlipper_Collide(parm)
    PlaySoundAtBallVol "flip_hit_" & Int(Rnd*3)+1, 2
End Sub

Sub RightFlipper_Collide(parm)
    PlaySoundAtBallVol "flip_hit_" & Int(Rnd*3)+1, 2
End Sub



'***********************************************************************************************

'******************* Others Table Sounds *************************

'Sub Trigger1_Hit():PlaySound "metal",0,1,Pan(ActiveBall),0,0,1,0,AudioFade(ActiveBall):End Sub 'Caster Club Wire Ramp Hit
Sub Trigger2_Hit():Tock.enabled=1:End Sub 'Exit Catapult Ramp Playfield Hit
Sub Tock_timer():PlaySound "ballhop",0,1,.2,0,0,0,1,-.4:Tock.enabled=0:End Sub
'Sub Trigger3_Hit():PlaySound "fx_metalrolling",0,1,Pan(ActiveBall),0,0,1,0,AudioFade(ActiveBall):End Sub 'Catapult Ramp Hit
'Sub Trigger4_Hit():PlaySound "metal",0,1,Pan(ActiveBall),0,0,1,0,AudioFade(ActiveBall):End Sub
'Sub Trigger5_Hit():PlaySound "fx_metalrolling",0,Vol(ActiveBall),Pan(ActiveBall),0,0,1,0,AudioFade(ActiveBall):End Sub 'Left Boat Ramp Hit
'Sub Trigger6_Hit():PlaySound "fx_metalrolling",0,Vol(ActiveBall),Pan(ActiveBall),0,0,1,0,AudioFade(ActiveBall):End Sub 'Right Boat ramp Hit
Sub Trigger7_Hit():Trigger8.enabled=1:StopSound "fx_metalrolling":End Sub 'Exit Left Boat Ramp Playfield Hit
Sub Trigger8_Hit():PlaySound "ballhop",0,1,Pan(ActiveBall),0,0,1,0,AudioFade(ActiveBall):Trigger8.enabled=0:End Sub
Sub Trigger9_Hit():Trigger10.enabled=1:StopSound "fx_metalrolling":End Sub 'Exit Right Boat Ramp Playfield Hit
Sub Trigger10_Hit():PlaySound "ballhop",0,1,Pan(ActiveBall),0,0,1,0,AudioFade(ActiveBall):Trigger10.enabled=0:End Sub
Sub Wall48_Hit():PlaySound "rubber_hit_3",0,0.02,Pan(ActiveBall),0,0,1,0,AudioFade(ActiveBall):End Sub 'Left Boat Ramp Stopper Hit
Sub Wall110_Hit():PlaySound "rubber_hit_3",0,0.02,Pan(ActiveBall),0,0,1,0,AudioFade(ActiveBall):End Sub 'Right Boat Ramp Stopper Hit
Sub Trigger11_Hit():PlaySound "metalhit_medium",1,0.2,Pan(ActiveBall),0,0,1,0,AudioFade(ActiveBall):End Sub 'ShooterLane Ramp Hit
Sub Trigger13_Hit():StopSound "metal":Trigger14.enabled=1:End Sub
Sub Trigger14_Hit():PlaySound "ballhop",0,.6,Pan(ActiveBall),0,0,1,0,AudioFade(ActiveBall):Trigger14.enabled=0:End Sub 'Exit ShooterLane Ramp Playfield Hit
Sub Wall97_Hit():PlaySound "metalhit_thin",0,1,Pan(ActiveBall),0,0,1,0,AudioFade(ActiveBall):End Sub 'Release Ball Hit
Sub Wall216_Hit():PlaySound "metalhit_medium",0,0.01,Pan(ActiveBall),0,0,1,0,AudioFade(ActiveBall):End Sub 'Plunger Hit
Sub Trigger12_Hit():PlaySound "metalhit2",0,0.05,Pan(ActiveBall),0,0,1,0,AudioFade(ActiveBall):End Sub 'Boat Ramp Hit
Sub Wall53_Hit():PlaySound "metalhit_medium",0,Vol(ActiveBall)*3,Pan(ActiveBall),0,0,1,0,AudioFade(ActiveBall):End Sub 'Apron Hit
Sub Wall146_Hit():PlaySound "metalhit_thin",0,1,Pan(ActiveBall),0,0,1,0,AudioFade(ActiveBall):End Sub 'Little Wire Guide of I lane Hit
Sub Wall147_Hit():PlaySound "metalhit_thin",0,1,Pan(ActiveBall),0,0,1,0,AudioFade(ActiveBall):End Sub 'Little Wire Guide of I lane Hit
Sub Wall148_Hit():PlaySound "metalhit_thin",0,1,Pan(ActiveBall),0,0,1,0,AudioFade(ActiveBall):End Sub 'Fish Finder metal guide on enter Hit

'*****************************************************************

'********* Table Options **********

Sub InitOptions
  If RampDecals=0 Then Ramp34.visible=0:Ramp35.visible=0:Primitive132.visible=0:Primitive142.visible=1:Primitive133.visible=0:Primitive143.visible=1:End If
  If FlippersColor=1 Then Primitive109.image="ft_flipper_left_GREEN":Primitive110.image="ft_flipper_right_GREEN":End If
  If FlippersColor=2 Then Primitive109.image="ft_flipper_left_BLACK":Primitive110.image="ft_flipper_right_BLACK":End If
  If FlippersColor=3 Then FlipColor=Int(Rnd*3)+1 End If
  Select Case FlipColor
    Case 1 : Primitive109.image="ft_flipper_left":Primitive110.image="ft_flipper_right"
    Case 2 : Primitive109.image="ft_flipper_left_GREEN":Primitive110.image="ft_flipper_right_GREEN"
    Case 3 : Primitive109.image="ft_flipper_left_BLACK":Primitive110.image="ft_flipper_right_BLACK"
  End Select
  If DMDColor=0 Then ScoreText1.Visible=0:ScoreText2.Visible=0:ScoreText3.Visible=0:End If
  If DMDColor=1 Then ScoreText.Visible=0:ScoreText1.Visible=1:ScoreText2.Visible=0:ScoreText3.Visible=0:End If
  If DMDColor=2 Then ScoreText.Visible=0:ScoreText1.Visible=0:ScoreText2.Visible=1:ScoreText3.Visible=0:End If
  If DMDColor=3 Then ScoreText.Visible=0:ScoreText1.Visible=0:ScoreText2.Visible=0:ScoreText3.Visible=1:End If
  If DMDColor=4 Then DMDCol=Int(Rnd*4)+1 End If
  Select Case DMDCol
    Case 1 : ScoreText1.Visible=0:ScoreText2.Visible=0:ScoreText3.Visible=0
    Case 2 : ScoreText.Visible=0:ScoreText1.Visible=1:ScoreText2.Visible=0:ScoreText3.Visible=0
    Case 3 : ScoreText.Visible=0:ScoreText1.Visible=0:ScoreText2.Visible=1:ScoreText3.Visible=0
    Case 4 : ScoreText.Visible=0:ScoreText1.Visible=0:ScoreText2.Visible=0:ScoreText3.Visible=1
  End Select
End Sub

'**********************************

'******** Cabinet Mode Adjustment *********

Sub FsSetup()
l86.Height=280:l16f.Height=270:l17f.Height=270:l18f.Height=270:l81.RotX=90:l82.RotX=90:l83.RotX=90:l84.RotX=90:l85.RotX=90:l81.Height=320:l82.Height=280:l83.Height=250:l84.Height=215:l85.Height=170
l81b.Opacity=700:l82b.Opacity=700:l83b.Opacity=700:l84b.Opacity=700:l85b.Opacity=700:l81b.Height=325:l82b.Height=290:l83b.Height=260:l84b.Height=220:l85b.Height=185
l71.Height=200:LMoray.Height=320:LPufferFish.Height=240:LMermaid.Height=320:LTail.Height=310:LLittleFish.Height=320:Light11.Intensity=5:Light12.Intensity=5
f26.Height=500:f26.RotX=0:Lboat.Y=650:Lboat.RotX=0:Lboat.Height=170:f27.Height=250:f27.Opacity=800
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

'Set position as bumper and Vol manually.

Sub PlaySoundAtBumperVol(sound, tableobj, Vol)
    PlaySound sound, 1, Vol, Pan(tableobj), 0,0,1, 1, AudioFade(tableobj)
End Sub

'**********************************************************************


Dim NextOrbitHit:NextOrbitHit = 0

Sub WireRampBumps_Hit(idx)
  if BallVel(ActiveBall) > .3 and Timer > NextOrbitHit then
    RandomBump3 .5, Pitch(ActiveBall)+5
    ' Schedule the next possible sound time.  This prevents it from rapid-firing noises too much.
    ' Lowering these numbers allow more closely-spaced clunks.
    NextOrbitHit = Timer + .2 + (Rnd * .2)
  end if
End Sub

Sub WireRampBumps2_Hit(idx)
  if BallVel(ActiveBall) > .3 and Timer > NextOrbitHit then
    RandomBump3 .2, Pitch(ActiveBall)+5
    ' Schedule the next possible sound time.  This prevents it from rapid-firing noises too much.
    ' Lowering these numbers allow more closely-spaced clunks.
    NextOrbitHit = Timer + .2 + (Rnd * .2)
  end if
End Sub

Sub PlasticRampBumps_Hit(idx)
  if BallVel(ActiveBall) > .3 and Timer > NextOrbitHit then
    RandomBump 4, Pitch(ActiveBall)
    ' Schedule the next possible sound time.  This prevents it from rapid-firing noises too much.
    ' Lowering these numbers allow more closely-spaced clunks.
    NextOrbitHit = Timer + .05 + (Rnd * .2)
  end if
End Sub


Sub MetalGuideBumps_Hit(idx)
  if BallVel(ActiveBall) > .3 and Timer > NextOrbitHit then
    RandomBump2 2, Pitch(ActiveBall)
    ' Schedule the next possible sound time.  This prevents it from rapid-firing noises too much.
    ' Lowering these numbers allow more closely-spaced clunks.
    NextOrbitHit = Timer + .2 + (Rnd * .2)
  end if
End Sub

Sub MetalWallBumps_Hit(idx)
  if BallVel(ActiveBall) > .3 and Timer > NextOrbitHit then
    RandomBump 2, 20000 'Increased pitch to simulate metal wall
    ' Schedule the next possible sound time.  This prevents it from rapid-firing noises too much.
    ' Lowering these numbers allow more closely-spaced clunks.
    NextOrbitHit = Timer + .2 + (Rnd * .2)
  end if
End Sub


' Requires rampbump1 to 7 in Sound Manager
Sub RandomBump(voladj, freq)
  dim BumpSnd:BumpSnd= "rampbump" & CStr(Int(Rnd*7)+1)
    PlaySound BumpSnd, 0, Vol(ActiveBall)*voladj, Pan(ActiveBall), 0, freq, 0, 1, AudioFade(ActiveBall)
End Sub

' Requires metalguidebump1 to 2 in Sound Manager
Sub RandomBump2(voladj, freq)
  dim BumpSnd:BumpSnd= "metalguidebump" & CStr(Int(Rnd*2)+1)
    PlaySound BumpSnd, 0, Vol(ActiveBall)*voladj, Pan(ActiveBall), 0, freq, 0, 1, AudioFade(ActiveBall)
End Sub

' Requires WireRampBump1 to 5 in Sound Manager
Sub RandomBump3(voladj, freq)
  dim BumpSnd:BumpSnd= "WireRampBump" & CStr(Int(Rnd*5)+1)
    PlaySound BumpSnd, 0, Vol(ActiveBall)*voladj, Pan(ActiveBall), 0, freq, 0, 1, AudioFade(ActiveBall)
End Sub


' Stop Bump Sounds
Sub BumpSTOPwire1_Hit ()
dim i:for i=1 to 5:StopSound "WireRampBump" & i:next
NextOrbitHit = Timer + 1
End Sub

Sub BumpSTOPwire2_Hit ()
dim i:for i=1 to 5:StopSound "WireRampBump" & i:next
NextOrbitHit = Timer + 1
End Sub

Sub BumpSTOPwire3_Hit ()
dim i:for i=1 to 5:StopSound "WireRampBump" & i:next
NextOrbitHit = Timer + 1
End Sub

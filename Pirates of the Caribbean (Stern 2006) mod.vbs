Option Explicit
Randomize

' Thalamus 2018-07-24
' Added InitVpmFFlipsSAM
' Thalamus 2018-08-18 : Improved directional sounds

Const VolDiv = 2000

Const VolBump   = 2    ' Bumpers multiplier.
Const VolRol    = 1    ' Rollovers volume multiplier.
Const VolGates  = 1    ' Gates volume multiplier.
Const VolMetals = 1    ' Metals volume multiplier.
Const VolRB     = 1    ' Rubber bands multiplier.
Const VolRH     = 1    ' Rubber hits multiplier.
Const VolRPo    = 1    ' Rubber posts multiplier.
Const VolRPi    = 1    ' Rubber pins multiplier.
Const VolPlast  = 1    ' Plastics multiplier.
Const VolTarg   = 1    ' Targets multiplier.
Const VolWood   = 1    ' Woods multiplier.
Const VolKick   = 1    ' Kicker multiplier.

Const VolSpin   = 1.5  ' Spinners volume.
Const VolFlip   = 1    ' Flipper volume.


On Error Resume Next
ExecuteGlobal GetTextFile("controller.vbs")
If Err Then MsgBox "You need the controller.vbs in order to run this table, available in the vp10 package"
On Error Goto 0

Const cGameName="potc_600af",UseSolenoids=1,UseLamps=0,UseGI=0,SSolenoidOn="SolOn",SSolenoidOff="SolOff", SCoin="coin"

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

SolCallBack(1)    = "SolTrough"
SolCallBack(2)    = "SolAutoPlungerIM"
SolCallBack(3)    = "VukTopPop"
SolCallBack(4)    ="SolChest"'CHEST LID
SolCallBack(5)="SolSailsUp"'RAISE SAILS
SolCallBack(6)    ="SolSpinner"
SolCallBack(18)   ="bsPOP.SolOut"
SolCallBack(19)   ="SolChestExit" 'Chest Kicker
SolCallBack(21)="SolShipMotor" 'Ship Motor
SolCallBack(23)   ="SolTortugaPost"
SolCallBack(27)="SolMotorDir"'SHIP MOTOR RELAY
SolCallBack(28)="SolSailsDown"'LOWER SAILS LATCH
SolCallBack(29)   ="SolSHIPPIN"

SolCallBack(20)   ="SetLamp 120,"
SolCallBack(22)   ="SetLamp 122,"
SolCallBack(30)   ="SetLamp 130,"
SolCallBack(31)   ="SetLamp 131,"
SolCallBack(32)   ="SetLamp 132,"

SolCallBack(15)="SolLFlipper"
SolCallBack(16)="SolRFlipper"

Sub SolLFlipper(Enabled)
     If Enabled Then
         PlaySoundAtVol SoundFX("fx_Flipperup",DOFContactors),LeftFlipper, VolFlip:LeftFlipper.RotateToEnd
     Else
         PlaySoundAtVol SoundFX("fx_Flipperdown",DOFContactors),LeftFlipper, VolFlip:LeftFlipper.RotateToStart
     End If
  End Sub

Sub SolRFlipper(Enabled)
     If Enabled Then
         PlaySoundAtVol SoundFX("fx_Flipperup",DOFContactors),RightFlipper, VolFlip:RightFlipper.RotateToEnd
     Else
         PlaySoundAtVol SoundFX("fx_Flipperdown",DOFContactors),RightFlipper, VolFlip:RightFlipper.RotateToStart
     End If
End Sub

'**********************************************************************************************************

'Solenoid Controlled toys
'**********************************************************************************************************

'Trough Ball
Sub SolTrough(Enabled)
  If Enabled Then
    bsTrough.ExitSol_On
    vpmTimer.PulseSw 22
  End If
 End Sub

Sub SolAutoPlungerIM(Enabled)
  If Enabled Then
    PlungerIM.AutoFire
  End If
 End Sub

Sub SolSHIPPIN(Enabled)
  If Enabled Then
    ShipPin.IsDropped=0
    playsoundAt SoundFX("Popper",DOFContactors), Primitive48
  Else
    ShipPin.IsDropped=1
    playsoundAt SoundFX("Popper",DOFContactors), Primitive48
  End If
End Sub

Sub SolTortugaPost(Enabled)
  If Enabled Then
    TortugaPost.IsDropped=0
    TortugaPin.transy = 50
    playsoundAt SoundFX("Popper",DOFContactors), Primitive18
  Else
    TortugaPost.IsDropped=1
    TortugaPin.transy = 0
    playsoundAt SoundFX("Popper",DOFContactors), Primitive18
  End If
End Sub

Sub SolSpinner (Enabled)
  mDISC.MotorOn = Enabled
  RotationTimer.Enabled = Enabled
End Sub

Sub RotationTimer_Timer
  RotatingPlatform.RotAndTra2 = RotatingPlatform.RotAndTra2 +40
End Sub


'Stern-Sega GI
set GICallback = GetRef("UpdateGI")

Sub UpdateGI(no, Enabled)
  If Enabled Then
    dim xx, xxx
    For each xx in GI:xx.State = 1: Next
    For each xxx in GI2:xxx.visible = 1: Next
        PlaySound "fx_relay"
  Else
    For each xx in GI:xx.State = 0: Next
    For each xxx in GI2:xxx.visible = 0: Next
        PlaySound "fx_relay"
  End If
End Sub

'**********************************************************************************************************

'Initiate Table
'**********************************************************************************************************

Dim bsTROUGH, bsPOP, bsCHEST, bsL, mDISC

Sub Table1_Init
  vpmInit Me
  On Error Resume Next
    With Controller
    .GameName = cGameName
    If Err Then MsgBox "Can't start Game" & cGameName & vbNewLine & Err.Description : Exit Sub
    .SplashInfoLine = "Pirates of the Caribbean (Stern 2006)"&chr(13)&"You Suck"
    .HandleMechanics=0
    .HandleKeyboard=0
    .ShowDMDOnly=1
    .ShowFrame=0
    .ShowTitle=0
        .hidden = 0
     On Error Resume Next
      InitVpmFFlipsSAM
     On Error Goto 0
         On Error Resume Next
         .Run GetPlayerHWnd
         If Err Then MsgBox Err.Description
         On Error Goto 0
     End With
     On Error Goto 0

    PinMAMETimer.Interval = PinMAMEInterval
    PinMAMETimer.Enabled = 1

  vpmNudge.TiltSwitch=-7
  vpmNudge.Sensitivity=2
  vpmNudge.TiltObj=Array(LeftSlingshot,RightSlingshot,Bumper1,Bumper2,Bumper3)

  Set bsTROUGH=New cvpmBallStack
    bsTROUGH.InitSw 0,21,20,19,18,0,0,0
    bsTROUGH.InitKick BallRelease,73,7
    bsTROUGH.InitExitSnd SoundFX("ballrelease",DOFContactors), SoundFX("Solenoid",DOFContactors)
    bsTROUGH.Balls=4

  Set bsPOP=New cvpmBallStack
    bsPOP.InitSaucer sw56,56,75,30
    bsPOP.InitExitSnd SoundFX("Popper",DOFContactors), SoundFX("Solenoid",DOFContactors)
    bsPOP.KickAngleVar = 1
    bsPOP.KickForceVar = 1

    Set bsCHEST = New cvpmBallStack
    With bsCHEST
        .InitSw 0, 6, 0, 0, 0, 0, 0, 0
        .InitKick sw6, 65, 15
        .KickZ = 0.4
        .InitExitSnd SoundFX("Popper",DOFContactors), SoundFX("Solenoid",DOFContactors)
        '.KickForceVar = 2
        .KickAngleVar = 2
        .KickBalls = 1
    End With

    Set mDISC=New cvpmTurnTable
    mDISC.InitTurnTable Plunder,60
    mDISC.SpinUp=30
    mDISC.SpinDown=25
    mDISC.CreateEvents"mDISC"

  ChestOpen.IsDropped=1
  ShipPin.IsDropped=1
  Sink1.IsDropped=1
  Sink2.IsDropped=1
  SailsDown.IsDropped=1
  TortugaPost.IsDropped=1

  Controller.Switch(63)=1

 End Sub

'**********************************************************************************************************
'Plunger code
'**********************************************************************************************************

Sub Table1_KeyDown(ByVal KeyCode)
  If KeyDownHandler(keycode) Then Exit Sub
  If keycode = PlungerKey Then Plunger.Pullback:PlaySoundAt "plungerpull", Plunger
    If KeyCode=7 Then vpmTimer.PulseSw 15   'Start Tournament
  If Keycode = RightFlipperKey then Controller.Switch(82)=1:Controller.Switch(90)=1
  If Keycode = LeftFlipperKey then Controller.Switch(84)=1
  If Keycode = StartGameKey Then Controller.Switch(16) = 1
End Sub

Sub Table1_KeyUp(ByVal KeyCode)
  If KeyUpHandler(keycode) Then Exit Sub
  If keycode = PlungerKey Then Plunger.Fire:PlaySoundAt "plunger", plunger
  If Keycode = RightFlipperKey then Controller.Switch(82)=0:Controller.Switch(90)=0
  If Keycode = LeftFlipperKey then Controller.Switch(84)=0
  If Keycode = StartGameKey Then Controller.Switch(16) = 0
End Sub

     ' Impulse Plunger
  Dim PlungerIM
    Const IMPowerSetting = 55
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
Sub Drain_Hit:bsTrough.addball me : PlaySoundAt "drain", Drain : End Sub
Sub sw6_Hit:bsCHEST.AddBall Me: playsoundAt "popper_ball", sw6: End Sub
Sub sw56_Hit:bsPOP.AddBall Me: playsoundAt "popper_ball", sw56: End Sub

'Wire Triggers
Sub sw1_Hit:Controller.Switch(1) = 1 : PlaySoundAtVol "rollover", sw1, VolRol : End Sub
Sub sw1_UnHit:Controller.Switch(1) = 0:End Sub
Sub sw2_Hit:Controller.Switch(2) = 1 : PlaySoundAtVol "rollover", sw2, VolRol : End Sub
Sub sw2_UnHit:Controller.Switch(2) = 0:End Sub
Sub sw23_Hit:Controller.Switch(23) = 1 : PlaySoundAtVol "rollover", sw23, VolRol : End Sub
Sub sw23_UnHit:Controller.Switch(23) = 0:End Sub
Sub sw24_Hit:Controller.Switch(24) = 1 : PlaySoundAtVol "rollover", sw24, VolRol : End Sub
Sub sw24_UnHit:Controller.Switch(24) = 0:End Sub
Sub sw25_Hit:Controller.Switch(25) = 1 : PlaySoundAtVol "rollover", sw25, VolRol : End Sub
Sub sw25_UnHit:Controller.Switch(25) = 0:End Sub
Sub sw28_Hit:Controller.Switch(28) = 1 : PlaySoundAtVol "rollover", sw28, VolRol : End Sub
Sub sw28_UnHit:Controller.Switch(28) = 0:End Sub
Sub sw29_Hit:Controller.Switch(29) = 1 : PlaySoundAtVol "rollover", sw29, VolRol : End Sub
Sub sw29_UnHit:Controller.Switch(29) = 0:End Sub

'Hidden wire Trigges
Sub sw3_Hit:Controller.Switch(3)=1 : PlaySoundAtVol "rollover", sw3, VolRol : End Sub
Sub sw3_unHit:Controller.Switch(3)=0:End Sub
Sub sw4_Hit:Controller.Switch(4)=1 : PlaySoundAtVol "rollover", sw3, VolRol : End Sub
Sub sw4_unHit:Controller.Switch(4)=0:End Sub
Sub sw57_Hit:Controller.Switch(57)=1 :  PlaySoundAt "Wire Ramp", ActiveBall  : End Sub
Sub sw57_unHit:Controller.Switch(57)=0:End Sub
Sub sw61_Hit:Controller.Switch(61)=1 : PlaySoundAtVol "rollover", sw61, VolRol : End Sub
Sub sw61_unHit:Controller.Switch(61)=0:End Sub

'Gate Triggers
Sub sw8_Hit:vpmTimer.PulseSw 8:End Sub
Sub sw11_Hit:vpmTimer.PulseSw 11: PlaySoundAt "fx_ballrampdrop", sw11 : End Sub
Sub sw10_Hit:vpmTimer.PulseSw 10:End Sub
Sub sw58_Hit:vpmTimer.PulseSw 58:End Sub

'Bumpers
Sub Bumper1_Hit : vpmTimer.PulseSw(30) : PlaySoundAtVol SoundFX("fx_bumper1",DOFContactors), Bumper1, VolBump: End Sub
Sub Bumper2_Hit : vpmTimer.PulseSw(31) : PlaySoundAtVol SoundFX("fx_bumper1",DOFContactors), Bumper1, VolBump: End Sub
Sub Bumper3_Hit : vpmTimer.PulseSw(32) : PlaySoundAtVol SoundFX("fx_bumper1",DOFContactors), Bumper1, VolBump: End Sub

'Stand Up Targets
Sub sw37_Hit:vpmTimer.PulseSw 37:End Sub
Sub sw39_Hit:vpmTimer.PulseSw 39:End Sub
Sub sw43_Hit:vpmTimer.PulseSw 42:End Sub
Sub sw43_Hit:vpmTimer.PulseSw 43:End Sub
Sub sw44_Hit:vpmTimer.PulseSw 44:End Sub
Sub sw45_Hit:vpmTimer.PulseSw 45:End Sub
Sub sw46_Hit:vpmTimer.PulseSw 46:End Sub
Sub sw47_Hit:vpmTimer.PulseSw 47:End Sub
Sub sw50_Hit:vpmTimer.PulseSw 50:End Sub
Sub sw51_Hit:vpmTimer.PulseSw 51:End Sub
Sub sw52_Hit:vpmTimer.PulseSw 52:End Sub
Sub sw53_Hit:vpmTimer.PulseSw 53:End Sub
Sub sw54_Hit:vpmTimer.PulseSw 54:End Sub
Sub sw55_Hit:vpmTimer.PulseSw 55:End Sub

'Upper PF
Sub sw12_Hit:Controller.Switch(12)=1 : PlaySoundAtVol "rollover", sw12, VolRol : End Sub
Sub sw12_unHit:Controller.Switch(12)=0:End Sub
Sub sw13_Hit:vpmTimer.PulseSw 13:End Sub
Sub sw14_Hit:Controller.Switch(14)=1 : PlaySoundAtVol "rollover", sw14, VolRol : End Sub
Sub sw14_unHit:Controller.Switch(14)=0:End Sub

'hole
Sub sw60_Hit:Controller.Switch(60)=1 : PlaySoundAtVol "fx_ballrampdrop", sw60, VolRol : End Sub
Sub sw60_unHit:Controller.Switch(60)=0:End Sub

'Generic Sounds
Sub Trigger1_Hit: PlaySoundAt "fx_ballrampdrop", Trigger1 : End Sub
Sub Trigger2_Hit: PlaySoundAt "fx_ballrampdrop", Trigger2 : End Sub
Sub Trigger3_Hit: PlaySoundAt "fx_ballrampdrop", Trigger3 : End Sub
Sub Trigger4_Hit: PlaySoundAt "fx_ballrampdrop", Trigger4 : End Sub


 '***********************************
 'Vertical Wire ramp Ball animation
 '***********************************

 Dim raiseballsw, raiseball

 Sub TopVUK_Hit()
  TopVUK.Enabled=FALSE
  Controller.switch (9) = True
  PlaySoundAt "popper_ball", TopVUK
 End Sub

 Sub VukTopPop(enabled)
  if(enabled and Controller.switch (9)) then
    TopVUK.DestroyBall
    Set raiseball = TopVUK.CreateBall
    raiseballsw = True
    TopVukraiseballtimer.Enabled = True
    TopVUK.Enabled=TRUE
    Controller.switch (9) = False
    playsound SoundFX("Popper",DOFContactors)
  else

  end if
End Sub

 Sub TopVukraiseballtimer_Timer()
  If raiseballsw = True then
    raiseball.z = raiseball.z + 10
    If raiseball.z > 175 then
      TopVUK.Kick 100, 10
      Set raiseball = Nothing
      TopVukraiseballtimer.Enabled = False
      raiseballsw = False
    End If
  End If
 End Sub

''************************************************************************************
''*****************       Chest                  ****************************
''************************************************************************************


 Sub SolChestExit(Enabled)
  If Enabled Then
    ChestOpen.IsDropped=1
    ChestClosed.IsDropped=0
    bsChest.ExitSol_On
  End If
 End Sub

 Sub SolChest(Enabled)
  If Enabled Then
    ChestOpen.IsDropped=0
    ChestClosed.IsDropped=1
    End If
 End Sub

Dim ChestO

Sub ChestMon_Timer
  if chestopen.IsDropped = True then
    chesto = False
    f24.height = 0:f32.height = 0:f40.height = 0:f48.height = 0:f56.height = 0
  else
    chesto = True
    f24.height = 217:f32.height = 217:f40.height = 217:f48.height = 217:f56.height = 217
  end if
End Sub

Sub chestMove_Timer()
  If chesto = True and chest.RotZ < 66 then chest.RotZ = chest.RotZ + 3
  If chesto = False and chest.RotZ > 0 then chest.RotZ = chest.RotZ - 3
  If chest.RotZ >= 0 then chesto = False
End Sub

''************************************************************************************
''*****************       Ship                 ****************************
''************************************************************************************

Dim ShipDir,ShipPos,MotorDir
ShipDir=1:ShipPos=0:MotorDir=0

 Sub SolMotorDir(Enabled)
  If Enabled Then
    MotorDir=1
  Else
    MotorDir=-1
  End If
End Sub

Sub SolShipMotor(Enabled)
  If Enabled Then
    ShipTimer.Enabled=True
  End If
End Sub

Sub ShipTimer_Timer
  ShipPos=ShipPos+MotorDir
  If ShipPos>0 And ShipPos<3 Then
    Controller.Switch(62)=0
    Controller.Switch(63)=0
  End If
  If ShipPos<0 Then
    ShipPos=0
    ShipTimer.Enabled=0
    Controller.Switch(63)=1
    Controller.Switch(62)=0
  End If
  If ShipPos>3 Then
    ShipPos=3
    ShipTimer.Enabled=0
    Controller.Switch(63)=0
    Controller.Switch(62)=1
  End If
  Select Case ShipPos
    Case 0:SailsUp.IsDropped=1:Sink1.IsDropped=1:Sink2.IsDropped=1:SailsDown.IsDropped=0
    Case 1:SailsUp.IsDropped=1:SailsDown.IsDropped=1:Sink2.IsDropped=1:Sink1.IsDropped=0
    Case 2:SailsUp.IsDropped=1:SailsDown.IsDropped=1:Sink1.IsDropped=1:Sink2.IsDropped=0
  End Select
End Sub

 Sub SolSailsUp(Enabled)
  If Enabled Then
    If ShipPos=0 Then
    Sink1.IsDropped=1
    Sink2.IsDropped=1
    SailsDown.IsDropped=1
    SailsUp.IsDropped=0
    End If
  End If
 End Sub

 Sub SolSailsDown(Enabled)
  If Enabled Then
    If ShipPos=0 Then
    Sink1.IsDropped=1
    Sink2.IsDropped=1
    SailsUp.IsDropped=1
    SailsDown.IsDropped=0
    End If
  End If
 End Sub

Dim sailsupw, sailsdownw, sink1w, sink2w

Sub PrimMon_Timer
    if sailsup.IsDropped = false then sailsupw = True else sailsupw = False 'everything is upright
    if sailsdown.IsDropped = false then sailsdownw = True else sailsdownw = False 'ship is upright but masts are bent forward
    if sink1.IsDropped = false then sink1w = True else sink1w = False 'ship leaning back but sails facing upright
    if sink2.IsDropped = false then sink2w = True else sink2w = False 'ship is vertical and sails are bent far forward
End Sub

Sub PrimMove_Timer()
  If sailsupw = True and ship.RotX > 90 then ship.RotX = Ship.RotX - 2
  If sailsupw = True and ship.RotX < 90 then ship.RotX = Ship.RotX + 2
  If sailsupw = True and mast1.RotX > 90 then mast1.RotX = mast1.RotX - 2
  If sailsupw = True and mast1.RotX < 90 then mast1.RotX = mast1.RotX + 2
  If sailsupw = True and mast1.TransY < 0 then mast1.TransY = mast1.TransY + 2
  If sailsupw = True and mast1.TransY > 0 then mast1.TransY = mast1.TransY - 2
  If sailsdownw = True and ship.RotX > 90 then ship.RotX = Ship.RotX - 2
  If sailsdownw = True and ship.RotX < 90 then ship.RotX = Ship.RotX + 2
  If sailsdownw = True and mast1.RotX > 46 then mast1.RotX = mast1.RotX - 2
  If sailsdownw = True and mast1.RotX < 46 then mast1.RotX = mast1.RotX + 2
  If sailsdownw = True and mast1.TransY < 0 then mast1.TransY = mast1.TransY + 2
  If sailsdownw = True and mast1.TransY > 0 then mast1.TransY = mast1.TransY - 2
  If sink1w = True and ship.RotX < 134 then ship.RotX = ship.RotX + 2
  If sink1w = True and ship.RotX > 134 then ship.RotX = ship.RotX - 2
  If sink1w = True and mast1.RotX < 90 then mast1.RotX = mast1.RotX + 2
  If sink1w = True and mast1.RotX > 90 then mast1.RotX = mast1.RotX - 2
  If sink1w = True and mast1.TransY < -24 then mast1.TransY = mast1.TransY + 2
  If sink1w = True and mast1.TransY > -24 then mast1.TransY = mast1.TransY - 2
  If sink2w = True and ship.RotX < 180 then ship.RotX = ship.RotX + 2
  If sink2w = True and ship.RotX > 180 then ship.RotX = ship.RotX - 2
  If sink2w = True and mast1.RotX < 90 then mast1.RotX = mast1.RotX + 2
  If sink2w = True and mast1.RotX > 90 then mast1.RotX = mast1.RotX - 2
  If sink2w = True and mast1.TransY < -46 then mast1.TransY = mast1.TransY + 2
  If sink2w = True and mast1.TransY > -46 then mast1.TransY = mast1.TransY - 2
  If ship.RotX = 90 and mast1.RotX = 90 then sailsupw = false
  If ship.RotX = 90 and mast1.RotX = 46 then sailsdownw = false
  If ship.RotX = 135 then sink1w = false
  If ship.RotX = 180 then sink2w = false
  mast1.TransZ = (ship.Rotx / 90 -1)*50
  mast2.RotX = mast1.RotX
  mast2.TransY = mast1.TransY
  mast2.TransZ = mast1.TransZ
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
NFadeL 20, l20
NFadeL 21, l21
NFadeL 22, l22
NFadeL 23, l23
Flash 24, f24
NFadeL 25, l25
NFadeL 26, l26
NFadeLm 27, l27a
NFadeLm 27, l27b
NFadeL 27, l27c
NFadeL 28, l28
NFadeL 29, l29
NFadeL 30, l30
NFadeL 31, l31
Flash 32, f32
Flash 33, f33
Flash 34, f34
Flash 35, f35
Flash 36, f36
Flash 37, f37
Flash 38, f38
Flash 39, f39
Flash 40, f40
NFadeL 41, l41
NFadeL 42, l42
NFadeL 43, l43
NFadeL 44, l44
NFadeL 45, l45
NFadeL 46, l46
NFadeL 47, l47
Flash 48, f48
NFadeL 49, l49
NFadeL 50, l50
NFadeL 51, l51
NFadeL 52, l52
NFadeL 53, l53
NFadeL 54, l54
NFadeL 55, l55
Flash 56, f56
NFadeL 57, l57
NFadeL 58, l58
NFadeL 59, l59
NFadeL 60, l60
NFadeL 61, l61
NFadeL 62, l62
NFadeL 63, l63
Flash 64, f64
NFadeL 65, l65
NFadeL 66, l66
NFadeL 67, l67
NFadeL 68, l68
NFadeL 69, l69
NFadeL 70, l70
NFadeL 71, l71
Flash 72, f72
Flash 73, F73

NFadeL 74, l74
NFadeL 75, l75
NFadeL 76, l76
NFadeL 77, l77
Flash 78, F78
Flash 79, F79
Flash 80, F80

'Solenoid Controlled Flashers/Lights

NFadeLm 120, f120a
Flash 120, f120b
NFadeL 122, f122
NFadeL 130, f130
NFadeL 131, f131
Flash 132, f132

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
    PlaySoundAt SoundFX("right_slingshot",DOFContactors), sling1
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
    PlaySoundAt SoundFX("left_slingshot",DOFContactors), sling2
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

'*****************************************
'   rothbauerw's Manual Ball Control
'*****************************************

Dim BCup, BCdown, BCleft, BCright
Dim ControlBallInPlay, ControlActiveBall
Dim BCvel, BCyveloffset, BCboostmulti, BCboost

BCboost = 1       'Do Not Change - default setting
BCvel = 4       'Controls the speed of the ball movement
BCyveloffset = -0.01  'Offsets the force of gravity to keep the ball from drifting vertically on the table, should be negative
BCboostmulti = 3    'Boost multiplier to ball veloctiy (toggled with the B key)

ControlBallInPlay = false

Sub StartBallControl_Hit()
  Set ControlActiveBall = ActiveBall
  ControlBallInPlay = true
End Sub

Sub StopBallControl_Hit()
  ControlBallInPlay = false
End Sub

Sub BallControlTimer_Timer()
  If EnableBallControl and ControlBallInPlay then
    If BCright = 1 Then
      ControlActiveBall.velx =  BCvel*BCboost
    ElseIf BCleft = 1 Then
      ControlActiveBall.velx = -BCvel*BCboost
    Else
      ControlActiveBall.velx = 0
    End If

    If BCup = 1 Then
      ControlActiveBall.vely = -BCvel*BCboost
    ElseIf BCdown = 1 Then
      ControlActiveBall.vely =  BCvel*BCboost
    Else
      ControlActiveBall.vely = bcyveloffset
    End If
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


'*****************************************
' ninuzzu's FLIPPER SHADOWS
'*****************************************

sub FlipperTimer_Timer()
  FlipperLSh.RotZ = LeftFlipper.currentangle
  FlipperRSh.RotZ = RightFlipper.currentangle
  LFLogo.Roty = LeftFlipper.currentangle
  RFLogo.Roty = RightFlipper.currentangle
    RampGate3.RotZ = -(Gate5.currentangle)
    RampGate1.RotZ = -(Gate6.currentangle)
    RampGate2.RotZ = -(Gate4.currentangle)
    RampGate4.RotZ = -(Gate9.currentangle)
    RampGate5.RotZ = -(sw13.currentangle)
End Sub

'*****************************************
' ninuzzu's BALL SHADOW
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
  PlaySound "pinhit_low", 0, Vol(ActiveBall)*VolRPi, AudioPan(ActiveBall), 0, Pitch(ActiveBall), 0, 0, AudioFade(ActiveBall)
End Sub

Sub Targets_Hit (idx)
  PlaySound "target", 0, Vol(ActiveBall)*VolTarg, AudioPan(ActiveBall), 0, Pitch(ActiveBall), 0, 0, AudioFade(ActiveBall)
End Sub

Sub Metals_Thin_Hit (idx)
  PlaySound "metalhit_thin", 0, Vol(ActiveBall)*VolMetals, AudioPan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
End Sub

Sub Metals_Medium_Hit (idx)
  PlaySound "metalhit_medium", 0, Vol(ActiveBall)*VolMetals, AudioPan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
End Sub

Sub Metals2_Hit (idx)
  PlaySound "metalhit2", 0, Vol(ActiveBall)*VolMetals, AudioPan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
End Sub

Sub Gates_Hit (idx)
  PlaySound "gate4", 0, Vol(ActiveBall)*VolGates, AudioPan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
End Sub

Sub Spinner_Spin
  PlaySound "fx_spinner", 0, .25, AudioPan(Spinner), 0.25, 0, 0, 1, AudioFade(Spinner)
End Sub

Sub Rubbers_Hit(idx)
  dim finalspeed
    finalspeed=SQR(activeball.velx * activeball.velx + activeball.vely * activeball.vely)
  If finalspeed > 20 then
    PlaySound "fx_rubber2", 0, Vol(ActiveBall)*VolRH, AudioPan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
  End if
  If finalspeed >= 6 AND finalspeed <= 20 then
    RandomSoundRubber()
  End If
End Sub

Sub Posts_Hit(idx)
  dim finalspeed
    finalspeed=SQR(activeball.velx * activeball.velx + activeball.vely * activeball.vely)
  If finalspeed > 16 then
    PlaySound "fx_rubber2", 0, Vol(ActiveBall)*VolRPo, AudioPan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
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
    Case 1 : PlaySound "flip_hit_1", 0, Vol(ActiveBall)*VolFlip, AudioPan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
    Case 2 : PlaySound "flip_hit_2", 0, Vol(ActiveBall)*VolFlip, AudioPan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
    Case 3 : PlaySound "flip_hit_3", 0, Vol(ActiveBall)*VolFlip, AudioPan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
  End Select
End Sub

' Thalamus : Exit in a clean and proper way
Sub Table1_exit()
  Controller.Pause = False
  Controller.Stop
End Sub


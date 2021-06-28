' Bone Busters Inc. Gottlieb 1989

' Script from the Full Rebuild for VPX by Aubrel (2021, May)
' Thanks to 32Assassin for sharing a beta version of this table that helped a lot for this rebuild
' (this present script has been improved using 32Assassin's beta table script as a base)

' Aubrel - ExtraBall and Special inserts fixed (L21a and L22a)
' Aubrel - LightsMod added to improve the lighting (not accurate when enabled)
' Aubrel - LUTs selection added and LightsLevelUpdate code added
' Aubrel - tilt lighting fixed and improved (should be now accurate)
' Aubrel - sound and collision script improved
' Aubrel - Opera jackpot playfield insert lighting fixed (now accurate)
' Aubrel - more function lights added (it should help to fix the table)
' Aubrel - drop target cheat added (it should help to fix the table)
' Aubrel - auto plunger added for "Bonus Balls" time (most probably not accurate but it helps and avoids problems)
' Aubrel - input corrections and UseFlipButton option added
' Aubrel - ballshooter fixed (now it should be fully working)
' Aubrel - drain and release kickers fixed (kicker script taken from Big House: it works now :p )
' Aubrel - flasher solenoids fixed


' v1.0 - 2015, December, 24th By Max
' Thanks to:
' gtxjoe for his resources and the ramp, the scoop and the flasher dome primitives
' jpsalas for allowing me to use his scripts and resources
' Unclewilly for allowing me to use his scripts and resources
' Destruk/ExtremeMame/Gravitar/TAB for the VP8 table
' Inkochnito for the DIP switches

option Explicit

On Error Resume Next
ExecuteGlobal GetTextFile("controller.vbs")
If Err Then MsgBox "You need the controller.vbs in order to run this table, available in the vp10 package"
On Error Goto 0
LoadVPM"01520000","sys80.VBS",1.2

Dim VarRol, VarHidden, Luts, LutPos, xxx
Dim VolRol, VolRamp, VolWire, VolFlipH, VolRub, VolMetal, VolWood, VolPin, VolTarg, VolGates, VolCol, VolSpin


'-----------------------------------
'------  Global Cofigurations ------
'-----------------------------------

Const cGameName="bonebstr",UseSolenoids=2,UseLamps=0,UseGI=0, SCoin="coin"
If Table1.ShowDT=true Then VarRol=0 : VarHidden=1 Else VarRol=1 : VarHidden=0
If B2SOn Then VarHidden=1   'comment this line or set 0 if you have problems with DMD showing; 1 is the default (disabled with dB2S)

'******************* User Options *********************
'******************************************************

'Control Option
Const UseFlipButton=0     'set to 1 if you want to use the right flipper button for the up-right flipper; 0 is defaut (right magna-save button)

'Visual Options
Luts = array("ColorGradeLUT256x16_1to1","ColorGradeLUT256x16_ModDeSat","ColorGradeLUT256x16_ConSat","ColorGradeLUT256x16_ExtraConSat","ColorGradeLUT256x16_ConSatMD","ColorGradeLUT256x16_ConSatD")
LutPos = 3              'set the nr of the LUT you want to use (0=first in the list above, 1=second, etc; 4 and 5 are "Medium Dark" and "Dark" LUTs); 3 is the default (ExtraConSat)

Const CustInstCards = 0     'set to 1 to get custom instruction cards; 0 is default (standard ones)

Const VisualLock = 0            'set to 1 if you don't want to change the visual settings using Magna-Save buttons; 0 is the default (live changes enabled)
Const LightsMod = 0             'set to 1 if you want to improve the lighting (red tubes blinking, Opera Jackpot inserts during attract mode, inserts switched off with tilt); 0 is default (accurate)
Const PrivateMod = 0        'set to 1 if you want to use nice imported primitives; 0 is default (meshs and UVs not included)


' Volume Options (Volume constants, 1 is max)
Const VolFlip    = 1     ' Flipper volume.
Const VolSling   = 1     ' Slingshot volume.
Const VolBump    = 1     ' Bumper volume.

' Sound Options (Volume Multipliers)
Const       GlobalVol = 1.2     ' Global mechanical volume

VolRol    = GlobalVol * 0.5     ' Ballrolling volume
VolRamp   = GlobalVol * 0.7     ' Ramps volume multiplier.
VolWire   = GlobalVol * 0.5     ' Wireramp effect volume.
VolFlipH  = GlobalVol * 30      ' Flipper hits volume.
VolRub    = GlobalVol * 100     ' Rubbers/Posts colision volume.
VolMetal  = GlobalVol * 200     ' Metals colision volume.
VolWood   = GlobalVol * 200     ' Woods colision volume.
VolPin    = GlobalVol * 5000    ' Pins colision volume.
VolTarg   = GlobalVol * 3000    ' Targets volume.
VolGates  = GlobalVol * 10000   ' Gates volume.
VolCol    = GlobalVol * 10000   ' Balls collision volume
VolSpin   = GlobalVol * 1000    ' Spinners volume


'Dev Options
Const TestLamps=0               'set to 1 to show function lamps in desktop mode; 0 is default (for devs only)
Const DTCheat=0         'set to 1 to enable DT cheat using BuyIn button (2); 0 is default (for devs only)


'******************************************************
'******************************************************

Table1.ColorGradeImage = Luts(LutPos)
If LutPos = 4 Then LightsLevelUpdate 1, 1    'set lights for "Medium Dark" LUT
If LutPos = 5 Then LightsLevelUpdate 2, 1    'set lights for "Dark" LUT

If CustInstCards=1 Then
  If PrivateMod=0 Then RightCardCust.Visible=1:LeftCardCust.Visible=1 Else RightCardCustM.Visible=1:LeftCardCustM.Visible=1
End If

If TestLamps=0 Then
  For each xxx in Test_Lamps : xxx.Image="Background" : Next
End If


'*************************************************************
'Solenoid Call backs
'**********************************************************************************************************
SolCallback(1)=  "Solenoid1"
SolCallback(2)=  "Solenoid2"
SolCallback(3)=  "Solenoid3"
SolCallback(4)=  "SetLamp 104,"
SolCallback(5)=  "Solenoid5"
SolCallback(6)=  "Solenoid6"
SolCallback(7)=  "SetLamp 107,"
SolCallback(8)=  "vpmSolSound SoundFX(""Knocker"",DOFKnocker),"
SolCallback(9)=  "SolTrough"

SolCallback(sLRFlipper) = "SolRFlipper"
SolCallback(sLLFlipper) = "SolLFlipper"

 ' Added By Outhere
'DOF 108,2   5 Bank Reset (Drop Targets)
'DOF 114,2   Left BallShooter
'DOF 113,2   BallRelease
'DOF 110,2   TopRightKicker

Sub SolLFlipper(Enabled)
     If Enabled Then
         FlipperLeft.RotateToEnd : FlipperLeftUp.RotateToEnd : PlaySoundAtVol SoundFX("fx_Flipperup",DOFFlippers),FlipperLeft,VolFlip
     Else
         FlipperLeft.RotateToStart : FlipperLeftUp.RotateToStart : PlaySoundAtVol SoundFX("fx_Flipperdown",DOFFlippers),FlipperLeft,VolFlip
     End If
  End Sub

Sub SolRFlipper(Enabled)
     If Enabled Then
         FlipperRight.RotateToEnd : PlaySoundAtVol SoundFX("fx_Flipperup",DOFFlippers),FlipperRight,VolFlip
     Else
         FlipperRight.RotateToStart : PlaySoundAtVol SoundFX("fx_Flipperdown",DOFFlippers),FlipperRight,VolFlip
     End If
End Sub

Dim RFUp
Sub FlipperTimer_Timer
  'FlipperRightUp Control
  If Controller.Switch(35)=True And Controller.Lamp(0)=True Then
    FlipperRightUp.RotateToEnd
    If RFUp=0 Then PlaySoundAtVol SoundFX("fx_Flipperup",DOFContactors),FlipperRightUp, VolFlip
    RFUp = 1
  Else
    FlipperRightUp.RotateToStart
    If RFUp=1 Then PlaySoundAtVol SoundFX("fx_Flipperdown",DOFContactors),FlipperRightUp, VolFlip
    RFUp = 0
  End If
  FlipperLeftGottlieb.objRotz = FlipperLeft.CurrentAngle-90
  FlipperLeftUpGottlieb.objRotz = FlipperLeftUp.CurrentAngle-90
  FlipperRightGottlieb.objRotz= FlipperRight.CurrentAngle-90
  FlipperRightUpGottlieb.objRotz= FlipperRightUp.CurrentAngle-90

  'ninuzzu's  FLIPPER SHADOWS
  FlipperLeftSh.RotZ = FlipperLeft.CurrentAngle
  FlipperLeftUpSh.RotZ = FlipperLeftUp.CurrentAngle
  FlipperRightSh.RotZ = FlipperRight.CurrentAngle
End Sub

Sub SolTrough(Enabled)
  If Enabled Then
    bsTrough.entrysol_on
    Kicker56.DestroyBall
  End If
End Sub

Sub Solenoid1(Enabled)
  If Enabled Then
    If Controller.lamp(13)=True Then 'S Relay Enabled
      SetLamp 100,1 'Top Right Lights 1
      SetLamp 101,1 'Top Right Lights 2
    Else
      SetLamp 100,0
      SetLamp 101,0
      bsRightKicker.ExitSol_On
    End If
  Else
    SetLamp 100,0
    SetLamp 101,0
  End If
End Sub

Sub Solenoid2(Enabled)
  If Enabled Then
    If Controller.lamp(13)=True Then 'S Relay Enabled
      SetLamp 102,1 'Right Side Light
    Else
      SetLamp 102,0
      bsCenterKicker.ExitSol_On
    End If
  Else
    SetLamp 102,0
  End If
End Sub

Sub Solenoid3(Enabled)
  If Enabled Then
    If Controller.lamp(13)=True Then 'S Relay Enabled
      SetLamp 103,1 'Top-Left Flipper Light
'     BallShooter.fire : PlaySoundAt SoundFX("plunger",DOFContactors),BallShooter 'Don't work: datasheet is wrong? (Made only with L12)
    Else
      SetLamp 103,0
    End If
  Else
    SetLamp 103,0
  End If
End Sub

Sub Solenoid5(Enabled)
  If Enabled Then
    If Controller.lamp(13)=True Then 'S Relay Enabled
      SetLamp 105,1 'Center Rollover Lights
    Else
      SetLamp 105,0
      If Controller.switch(46)=True Then  'Up-Right Kicker
        PlaySoundAt SoundFX("Popper",DOFContactors),Kicker46
        Kicker46.DestroyBall
        Set RaiseBall = Kicker46.CreateBall
        RaiseBallSW = True
        TopRightKickerTimer.Enabled = True
        Kicker46.Enabled = True
        Controller.switch(46) = False
      End If
    End If
  Else
    SetLamp 105,0
  End If
End Sub

Sub Solenoid6(Enabled)
  If Enabled Then
    If Controller.lamp(13)=True Then 'S Relay Enabled
      SetLamp 106,1 'Top-Left Spot Target and Center Ramp Lights
    Else
      SetLamp 106,0
      DTLow.DropSol_On '5 Bank Reset (Drop Targets)
DOF 108,2 ' Added By Outhere
      dim xx
      If LightsMod=0 Then
        If Controller.lamp(7)=True Then 'Opera Jackpot Lights On
          For each xx in BlinkInserts2:xx.State = 2: Next
        End If
      Else
        If Controller.lamp(7)=True Or Controller.Lamp(0)=False Then 'Opera Jackpot Lights On (L0 is not accurate but gives better attract mode)
          For each xx in BlinkInserts2:xx.State = 2: Next
        End If
      End If
    End If
  Else
      SetLamp 106,0
  End If
End Sub

Sub SolRelayA(Enabled) 'L14
  If Enabled Then
  Else
  End If
End Sub

Sub SolRelayB(Enabled) 'L15
  If Enabled Then
  Else
  End If
End Sub

Sub SolRelayC(Enabled) 'L18
  If Enabled Then
  Else
  End If
End Sub

Sub SolRelayQ(Enabled)  'InGame (L0)
  dim xx
  If Enabled Then
    If LightsMod=1 Then 'needed only with LightsMod
      For each xx in BlinkInserts2:xx.State = 0: Next 'Opera Jackpot Lights Off
    End If
  Else
    For each xx in BlinkInserts2:xx.State = 0: Next 'Opera Jackpot Lights Off
  End If
End Sub

Sub SolRelayS(Enabled)  'Solenoid Switch Function (L13)
  If Enabled Then
    L21.state=0
    L22.state=0
  Else
    L21a.State=0
    If Controller.Lamp(21)=True Then L21.State=1
    L22a.State=0
    If Controller.Lamp(22)=True Then L22.State=1
  End If
End Sub

Sub SolRelayT(Enabled)  'Tilt / GI Switch Down (L1)
  If Enabled Then
    dim xx
    For each xx in GI:xx.State = 0: Next
    If LightsMod=0 Then
      For each xx in TubeLights:xx.State = 1: Next
    Else
      For each xx in TubeLights:xx.State = 2: Next 'blinking only with LightsMod
      For each xx in AlwaysOn:xx.State = 0: Next
    End If
    LightsLevelUpdate 2, 1
    If Controller.Switch(43)=True Then PlungerAuto.Fire : PlaySoundAt SoundFX("plunger",DOFContactors),PlungerAuto 'it helps with Bonus Balls
  Else
    For each xx in GI:xx.State = 1: Next
    For each xx in TubeLights:xx.State = 0: Next
    If LightsMod=1 Then
      For each xx in AlwaysOn:xx.State = 1: Next
    End If
    PlaySound SoundFX("fx_relay",DOFContactors)
    LightsLevelUpdate -2, 1
  End If
End Sub

Sub SolBallRelease(Enabled) 'L2
  If Enabled Then
    If Controller.Switch(43)=True Then PlungerAuto.Fire : PlaySoundAt SoundFX("plunger",DOFContactors),PlungerAuto 'to fire the ball before adding an other one (used only with Bonus Balls)
    bsTrough.ExitSol_On
DOF 113,2  ' Added By Outhere
  Else
  End If
End Sub

Sub SolBallShooter(Enabled) ' L12
     If Enabled Then
    BallShooter.fire
    PlaySoundAt SoundFX("plunger",DOFContactors),BallShooter
DOF 114,2 ' Added By Outhere
    Else
    End If
End Sub

Sub SolLeftBallGate(Enabled) 'L16
     If Enabled Then
        LeftBallGate.RotateToEnd : PlaySoundAtVol SoundFX("fx_Flipperup",DOFContactors),LeftBallGate,VolFlip
     Else
        LeftBallGate.RotateToStart : PlaySoundAtVol SoundFX("fx_Flipperup",DOFContactors),LeftBallGate,VolFlip
     End If
End Sub


Sub SolRightBallGate(Enabled) 'L17
     If Enabled Then
        RightBallGate.RotateToEnd : PlaySoundAtVol SoundFX("fx_Flipperup",DOFContactors),RightBallGate,VolFlip
     Else
        RightBallGate.RotateToStart : PlaySoundAtVol SoundFX("fx_Flipperup",DOFContactors),RightBallGate,VolFlip
     End If
End Sub

'--------------------------
'------  Table Init  ------
'--------------------------
Dim bsTrough, bsCenterKicker, bsRightKicker, DTLow, RaiseBallSW, RaiseBall

Sub Table1_Init
  vpmInit Me
  On Error Resume Next
    With Controller
    .GameName = cGameName
    If Err Then MsgBox "Can't start Game" & cGameName & vbNewLine & Err.Description : Exit Sub
    .SplashInfoLine = "Bone Busters Inc."&chr(13)&"Gottlieb 1989"
    .HandleMechanics=0
    .HandleKeyboard=0
    .ShowDMDOnly=1
    .ShowFrame=0
    .ShowTitle=0
    .hidden=VarHidden
    .Games(cGameName).Settings.Value("rol") = VarRol

    On Error Resume Next
    Controller.SolMask(0)=0
    vpmTimer.AddTimer 2000,"Controller.SolMask(0)=&Hffffffff'" 'ignore all solenoids - then add the timer to renable all the solenoids after 2 seconds
    .Run GetPlayerHWnd
    If Err Then MsgBox Err.Description
    On Error Goto 0
  End With
  On Error Goto 0

  PinMAMETimer.Interval=PinMAMEInterval
  PinMAMETimer.Enabled=1
  vpmNudge.TiltObj=Array(Bumper,DropTarget00,DropTarget10,DropTarget20,DropTarget30,DropTarget40,FlipperLeft,FlipperLeftUp,FlipperRight,FlipperRightUp)
  vpmNudge.TiltSwitch=57
  vpmNudge.Sensitivity=2

  Set bsTrough=New cvpmBallStack

    bsTrough.InitSw 56,0,0,55,0,0,0,0
    bsTrough.initkick Kicker55,75,5
    bsTrough.InitExitSnd SoundFX("ballrelease",DOFContactors), SoundFX("Solenoid",DOFContactors)
    bsTrough.Balls=3
    bsTrough.addball 0

  Set bsCenterKicker=New cvpmBallStack
    bsCenterKicker.InitSaucer Kicker26,26,180,5
    bsCenterKicker.InitExitSnd SoundFX("Popper",DOFContactors), SoundFX("Solenoid",DOFContactors)

  Set bsRightKicker=New cvpmBallStack
    bsRightKicker.InitSaucer Kicker36,36,180,5
    bsRightKicker.InitExitSnd SoundFX("Popper",DOFContactors), SoundFX("Solenoid",DOFContactors)


  Set DTLow=New cvpmDropTarget
    DTLow.InitDrop Array(DropTarget00,DropTarget10,DropTarget20,DropTarget30,DropTarget40),Array(0,10,20,30,40)
    DTLow.InitSnd SoundFX("dtdrop",DOFDropTargets), SoundFX("dtreset",DOFDropTargets)

  captiveball.CreateBall
  captiveball.Kick 0,1

  'to enable GI from start
  For each xxx in GI:xxx.State = 1: Next

  For each xxx in AlwaysOn : xxx.State=1 : Next
  For each xxx in BlinkInserts : xxx.State=2 : Next
  If LightsMod=1 Then
    For each xxx in BlinkInserts2 : xxx.State=2 : Next
  End If

  If PrivateMod=1 Then
    For each xxx in Walls : xxx.Visible=False : xxx.SideVisible=False : Next
    For each xxx in Ramps : xxx.Visible=False : Next
    GI015.BulbHaloHeight=54
  Else
    For each xxx in PrimitivesMod : xxx.Visible=False : Next
  End If
End Sub


' Drain hole and kickers
Sub Kicker26_Hit:bsCenterKicker.AddBall 0 : PlaySoundAt SoundFX("popper_ball",DOFContactors),Kicker26 : End Sub
Sub Kicker36_Hit:bsRightKicker.AddBall 0 : PlaySoundAt SoundFX("popper_ball",DOFContactors),Kicker36 : End Sub
Sub Kicker56_Hit:bsTrough.AddBall 0 : PlaySoundAt SoundFX("drain",DOFContactors),Kicker56 : End Sub
Sub Kicker46_Hit()
  Kicker46.Enabled = False
  Controller.switch (46) = True
  PlaySoundAt SoundFX("popper_ball",DOFContactors),Kicker46
End Sub

Sub TopRightKickerTimer_Timer()
  If RaiseBallSW = True then
    RaiseBall.z = raiseball.z + 10
    If RaiseBall.z > 150 Then
      Kicker46.Kick 180, 10
      PlaySoundAt SoundFX("popper",DOFContactors),Kicker46
      Set RaiseBall = Nothing
      TopRightKickerTimer.Enabled = False
      RaiseBallSW = False
DOF 110,2 ' Added By Outhere
    End If
  End If
End Sub


'Drop Targets
Sub DropTarget00_dropped:DTLow.Hit 1:End Sub
Sub DropTarget10_dropped:DTLow.Hit 2:End Sub
Sub DropTarget20_dropped:DTLow.Hit 3:End Sub
Sub DropTarget30_dropped:DTLow.Hit 4:End Sub
Sub DropTarget40_dropped:DTLow.Hit 5:End Sub

'Stand Up Target
Sub Target01_Hit:vpmTimer.PulseSw 1:End Sub
Sub Target11_Hit:vpmTimer.PulseSw 11:End Sub
Sub Target21_Hit:vpmTimer.PulseSw 21:End Sub
Sub Target31_Hit:vpmTimer.PulseSw 31:End Sub
Sub Target41_Hit:vpmTimer.PulseSw 41:End Sub
Sub Target51_Hit:vpmTimer.PulseSw 51:End Sub

'Bounce Targets
Sub sw22_Slingshot:vpmTimer.PulseSw(22) : PlaySoundAtVol SoundFX("left_slingshot",DOFContactors),sw22,VolSling : End Sub
Sub sw23_Slingshot:vpmTimer.PulseSw(23) : PlaySoundAtVol SoundFX("right_slingshot",DOFContactors),sw23,VolSling : End Sub

'Star Triggers
Sub SW02_Hit:Controller.Switch(2)=1 : PlaySoundAtBall SoundFX("rollover",DOFContactors) : End Sub
Sub SW02_Unhit:Controller.Switch(2)=0:End Sub
Sub SW03_Hit:Controller.Switch(3)=1 : PlaySoundAtBall SoundFX("rollover",DOFContactors) : End Sub
Sub SW03_Unhit:Controller.Switch(3)=0:End Sub
Sub SW32_Hit:Controller.Switch(32)=1 : PlaySoundAtBall SoundFX("rollover",DOFContactors) : End Sub
Sub SW32_Unhit:Controller.Switch(32)=0:End Sub
Sub SW33_Hit:Controller.Switch(33)=1 : PlaySoundAtBall SoundFX("rollover",DOFContactors) : End Sub
Sub SW33_unHit:Controller.Switch(33)=0:End Sub

'Wire Triggers
Sub SW12_Hit:Controller.Switch(12)=1 : PlaySoundAtBall SoundFX("rollover",DOFContactors) : End Sub
Sub SW12_Unhit:Controller.Switch(12)=0:End Sub
Sub SW13_Hit:Controller.Switch(13)=1 : PlaySoundAtBall SoundFX("rollover",DOFContactors) : End Sub
Sub SW13_Unhit:Controller.Switch(13)=0:End Sub
Sub SW15_Hit:Controller.Switch(15)=1 : PlaySoundAtBall SoundFX("rollover",DOFContactors) : End Sub
Sub SW15_Unhit:Controller.Switch(15)=0:End Sub
Sub SW25_Hit:Controller.Switch(25)=1 : ActiveBall.VelY = ActiveBall.VelY / 2 : PlaySoundAtBall SoundFX("rollover",DOFContactors) : End Sub
Sub SW25_Unhit:Controller.Switch(25)=0:End Sub
Sub SW34_Hit:Controller.Switch(34)=1 : PlaySoundAtBall SoundFX("rollover",DOFContactors) : End Sub
Sub SW34_Unhit:Controller.Switch(34)=0:End Sub
Sub SW42_Hit:Controller.Switch(42)=1 : PlaySoundAtBall SoundFX("rollover",DOFContactors) : End Sub
Sub SW42_Unhit:Controller.Switch(42)=0:End Sub
Sub SW43_Hit:Controller.Switch(43)=1 : PlaySoundAtBall SoundFX("rollover",DOFContactors) : End Sub
Sub SW43_Unhit:Controller.Switch(43)=0:End Sub

'Spinners
Sub Spinner14_Spin:vpmTimer.PulseSw 14:End Sub
Sub Spinner24_Spin:vpmTimer.PulseSw 24:End Sub

'WireRamp Triggers
Sub SW04_Hit:Controller.Switch(4)=1 : PlaySound "wireramp",-1, Vol(ActiveBall)*VolWire, AudioPan(ActiveBall), 0, Pitch(ActiveBall), 0, 0, AudioFade(ActiveBall) : End Sub
Sub SW04_Unhit:Controller.Switch(4)=0:End Sub
Sub UpMetalRampEnd_Hit : StopSound "wireramp" : End Sub

Sub SW05_Hit:Controller.Switch(5)=1 : PlaySound "wireramp",-1, Vol(ActiveBall)*VolWire, AudioPan(ActiveBall), 0, Pitch(ActiveBall), 0, 0, AudioFade(ActiveBall) : End Sub
Sub SW05_Unhit:Controller.Switch(5)=0:End Sub
Sub CenterMetalRampEnd_Hit : StopSound "wireramp" : End Sub

'Scoring Rubber
Sub sw50_hit:vpmTimer.PulseSw(50):End Sub

'Gate Trigger
Sub Gate44_Hit
  vpmTimer.PulseSw(44)
    If ActiveBall.VelY > 0 Then
        ActiveBall.VelY = ActiveBall.VelY / 3
    End If
End Sub

Sub GateInv_Hit
  ActiveBall.VelY = ActiveBall.VelY / 3
End Sub

'Bumpers
Sub Bumper_Hit:vpmTimer.PulseSw(52):PlayBumperSound:End Sub

Sub PlayBumperSound()
    Select Case Int(Rnd*4)+1
      Case 1 : PlaySoundAtBumperVol SoundFX("fx_bumper1",DOFContactors),Bumper,VolBump
      Case 2 : PlaySoundAtBumperVol SoundFX("fx_bumper2",DOFContactors),Bumper,VolBump
      Case 3 : PlaySoundAtBumperVol SoundFX("fx_bumper3",DOFContactors),Bumper,VolBump
      Case 4 : PlaySoundAtBumperVol SoundFX("fx_bumper4",DOFContactors),Bumper,VolBump
    End Select
End Sub


'Gottlieb Bone Busters
'added by Inkochnito
Sub editDips
  Dim vpmDips:Set vpmDips=New cvpmDips
  With vpmDips
    .AddForm  700,400,"Bone Busters - DIP switches"
    .AddFrame 2,4,190,"Maximum credits",49152,Array("8 credits",0,"10 credits",32768,"15 credits",&H00004000,"20 credits",49152)'dip 15&16
    .AddFrame 2,80,190,"Coin chute 1 and 2 control",&H00002000,Array("seperate",0,"same",&H00002000)'dip 14
    .AddFrame 2,126,190,"Playfield special",&H00200000,Array("replay",0,"extra ball",&H00200000)'dip 22
    .AddFrame 2,172,190,"High games to date control",&H00000020,Array("no effect",0,"reset high games 2-5 on power off",&H00000020)'dip 6
    .AddFrame 2,218,190,"Auto-percentage control",&H00000080,Array("disabled (normal high score mode)",0,"enabled",&H00000080)'dip 8
    .AddFrame 2,264,190,"Extra ball && special timing",&H40000000,Array("shorter",0,"longer",&H40000000)'dip 31
    .AddChk 2,316,150,Array("Bonus ball feature",&H80000000)'dip 32
    .AddFrame 205,4,190,"High game to date awards",&H00C00000,Array("not displayed and no award",0,"displayed and no award",&H00800000,"displayed and 2 replays",&H00400000,"displayed and 3 replays",&H00C00000)'dip 23&24
    .AddFrame 205,80,190,"Balls per game",&H01000000,Array("5 balls",0,"3 balls",&H01000000)'dip 25
    .AddFrame 205,126,190,"Replay limit",&H04000000,Array("no limit",0,"one per game",&H04000000)'dip 27
    .AddFrame 205,172,190,"Novelty",&H08000000,Array("normal",0,"extra ball and replay scores 500K",&H08000000)'dip 28
    .AddFrame 205,218,190,"Game mode",&H10000000,Array("replay",0,"extra ball",&H10000000)'dip 29
    .AddFrame 205,264,190,"3rd coin chute credits control",&H20000000,Array("no effect",0,"add 9",&H20000000)'dip 30
    .AddChk 170,316,100,Array("Match feature",&H02000000)'dip 26
    .AddChk 290,316,100,Array("Attract sound",&H00000040)'dip 7
    .AddLabel 50,340,300,20,"After hitting OK, press F3 to reset game with new settings."
    .ViewDips
  End With
End Sub
Set vpmShowDips=GetRef("EditDips")


'-------------------------------
'------  Keybord Handler  ------
'-------------------------------

Sub Table1_KeyDown(ByVal keycode)
    If keycode = LeftTiltKey Then Nudge 90, 5:PlaySound SoundFX("fx_nudge", 0), 0, 1, -0.1, 0.25
    If keycode = RightTiltKey Then Nudge 270, 5:PlaySound SoundFX("fx_nudge", 0), 0, 1, 0.1, 0.25
    If keycode = CenterTiltKey Then Nudge 0, 6:PlaySound SoundFX("fx_nudge", 0), 0, 1, 0, 0.25
  If DTCheat=1 Then   'cheat for drop targets using BuyIn button (2)
    If keycode = KeyFront Then Controller.switch(0)=True:Controller.switch(10)=True:Controller.switch(20)=True:Controller.switch(30)=True:Controller.switch(40)=True
  End If
  If VisualLock = 0 Then
    If keycode = LeftMagnaSave Then
      If LutPos = 5 Then LutPos = 0:LightsLevelUpdate -2, 1 Else LutPos = LutPos +1 'Max LUTs number is set to 5; so back to 0 and Lights level reseted for standard LUTs.
      If LutPos > 3 Then LightsLevelUpdate 1, 1                     'LUTs 4 and 5 are "Medium Dark" and "Dark LUTs" so Lights Level should be updated
      Table1.ColorGradeImage = Luts(LutPos)
    End If
  End If
  If keycode = RightMagnaSave Then Controller.Switch(35)=1
  If keycode = PlungerKey Then Plunger.PullBack
  If keycode = LeftFlipperKey Then Controller.Switch(53)=1
  If UseFlipButton=0 Then
    If keycode = RightFlipperKey Then Controller.Switch(45)=1
  Else
    If keycode = RightFlipperKey Then Controller.Switch(45)=1 : Controller.Switch(35)=1
  End If
  If vpmKeyDown(KeyCode) Then Exit Sub
End Sub

Sub Table1_KeyUp(ByVal keycode)
  If keycode = PlungerKey Then Plunger.Fire
  If keycode = RightMagnaSave Then Controller.Switch(35)=0
  If keycode = LeftFlipperKey Then Controller.Switch(53)=0
  If UseFlipButton=0 Then
    If keycode = RightFlipperKey Then Controller.Switch(45)=0
  Else
    If keycode = RightFlipperKey Then Controller.Switch(45)=0 : Controller.Switch(35)=0
  End If
  If vpmKeyUp(KeyCode) Then Exit Sub
End Sub

Sub LightsLevelUpdate(Offset, Coef)
  Dim obj
  For each obj in AllFlashers
    obj.Opacity = obj.Opacity * 1.3^(Offset*Coef)
  Next
  For each obj in FlasherLights
    obj.Intensity = obj.Intensity * 1.2^(Offset*Coef)
  Next
  For each obj in GI
    obj.Intensity = obj.Intensity * 1.15^(Offset*Coef)
  Next
  For each obj in AlwaysOn
    obj.Intensity = obj.Intensity * 1.25^(Offset*Coef)
  Next
  For each obj in PFLights
    obj.Intensity = obj.Intensity * 1.25^(Offset*Coef)
  Next
  For each obj in BlinkInserts
    obj.Intensity = obj.Intensity * 1.25^(Offset*Coef)
    Next
  For each obj in BlinkInserts2
    obj.Intensity = obj.Intensity * 1.25^(Offset*Coef)
    Next
  For each obj in TubeLights
    obj.Intensity = obj.Intensity * 1.3^(Offset*Coef)
    Next
End Sub


'**********************************************************************************************************
'Digital Display
'**********************************************************************************************************
 Dim Digits(40)
 Digits(0)=Array(a00, a05, a0c, a0d, a08, a01, a06, a0f, a02, a03, a04, a07, a0b, a0a, a09, a0e)
 Digits(1)=Array(a10, a15, a1c, a1d, a18, a11, a16, a1f, a12, a13, a14, a17, a1b, a1a, a19, a1e)
 Digits(2)=Array(a20, a25, a2c, a2d, a28, a21, a26, a2f, a22, a23, a24, a27, a2b, a2a, a29, a2e)
 Digits(3)=Array(a30, a35, a3c, a3d, a38, a31, a36, a3f, a32, a33, a34, a37, a3b, a3a, a39, a3e)
 Digits(4)=Array(a40, a45, a4c, a4d, a48, a41, a46, a4f, a42, a43, a44, a47, a4b, a4a, a49, a4e)
 Digits(5)=Array(a50, a55, a5c, a5d, a58, a51, a56, a5f, a52, a53, a54, a57, a5b, a5a, a59, a5e)
 Digits(6)=Array(a60, a65, a6c, a6d, a68, a61, a66, a6f, a62, a63, a64, a67, a6b, a6a, a69, a6e)
 Digits(7)=Array(a70, a75, a7c, a7d, a78, a71, a76, a7f, a72, a73, a74, a77, a7b, a7a, a79, a7e)
 Digits(8)=Array(a80, a85, a8c, a8d, a88, a81, a86, a8f, a82, a83, a84, a87, a8b, a8a, a89, a8e)
 Digits(9)=Array(a90, a95, a9c, a9d, a98, a91, a96, a9f, a92, a93, a94, a97, a9b, a9a, a99, a9e)
 Digits(10)=Array(aa0, aa5, aac, aad, aa8, aa1, aa6, aaf, aa2, aa3, aa4, aa7, aab, aaa, aa9, aae)
 Digits(11)=Array(ab0, ab5, abc, abd, ab8, ab1, ab6, abf, ab2, ab3, ab4, ab7, abb, aba, ab9, abe)
 Digits(12)=Array(ac0, ac5, acc, acd, ac8, ac1, ac6, acf, ac2, ac3, ac4, ac7, acb, aca, ac9, ace)
 Digits(13)=Array(ad0, ad5, adc, add, ad8, ad1, ad6, adf, ad2, ad3, ad4, ad7, adb, ada, ad9, ade)
 Digits(14)=Array(ae0, ae5, aec, aed, ae8, ae1, ae6, aef, ae2, ae3, ae4, ae7, aeb, aea, ae9, aee)
 Digits(15)=Array(af0, af5, afc, afd, af8, af1, af6, aff, af2, af3, af4, af7, afb, afa, af9, afe)
 Digits(16)=Array(b00, b05, b0c, b0d, b08, b01, b06, b0f, b02, b03, b04, b07, b0b, b0a, b09, b0e)
 Digits(17)=Array(b10, b15, b1c, b1d, b18, b11, b16, b1f, b12, b13, b14, b17, b1b, b1a, b19, b1e)
 Digits(18)=Array(b20, b25, b2c, b2d, b28, b21, b26, b2f, b22, b23, b24, b27, b2b, b2a, b29, b2e)
 Digits(19)=Array(b30, b35, b3c, b3d, b38, b31, b36, b3f, b32, b33, b34, b37, b3b, b3a, b39, b3e)

 Digits(20)=Array(b40, b45, b4c, b4d, b48, b41, b46, b4f, b42, b43, b44, b47, b4b, b4a, b49, b4e)
 Digits(21)=Array(b50, b55, b5c, b5d, b58, b51, b56, b5f, b52, b53, b54, b57, b5b, b5a, b59, b5e)
 Digits(22)=Array(b60, b65, b6c, b6d, b68, b61, b66, b6f, b62, b63, b64, b67, b6b, b6a, b69, b6e)
 Digits(23)=Array(b70, b75, b7c, b7d, b78, b71, b76, b7f, b72, b73, b74, b77, b7b, b7a, b79, b7e)
 Digits(24)=Array(b80, b85, b8c, b8d, b88, b81, b86, b8f, b82, b83, b84, b87, b8b, b8a, b89, b8e)
 Digits(25)=Array(b90, b95, b9c, b9d, b98, b91, b96, b9f, b92, b93, b94, b97, b9b, b9a, b99, b9e)
 Digits(26)=Array(ba0, ba5, bac, bad, ba8, ba1, ba6, baf, ba2, ba3, ba4, ba7, bab, baa, ba9, bae)
 Digits(27)=Array(bb0, bb5, bbc, bbd, bb8, bb1, bb6, bbf, bb2, bb3, bb4, bb7, bbb, bba, bb9, bbe)
 Digits(28)=Array(bc0, bc5, bcc, bcd, bc8, bc1, bc6, bcf, bc2, bc3, bc4, bc7, bcb, bca, bc9, bce)
 Digits(29)=Array(bd0, bd5, bdc, bdd, bd8, bd1, bd6, bdf, bd2, bd3, bd4, bd7, bdb, bda, bd9, bde)
 Digits(30)=Array(be0, be5, bec, bed, be8, be1, be6, bef, be2, be3, be4, be7, beb, bea, be9, bee)
 Digits(31)=Array(bf0, bf5, bfc, bfd, bf8, bf1, bf6, bff, bf2, bf3, bf4, bf7, bfb, bfa, bf9, bfe)
 Digits(32)=Array(ac18, ac16, acc1, acd1, ac19, ac17, ac15, acf1, ac11, ac13, ac12, ac14, acb1, aca1, ac10, ace1)
 Digits(33)=Array(ad18, ad16, adc1, add1, ad19, ad17, ad15, adf1, ad11, ad13, ad12, ad14, adb1, ada1, ad10, ade1)
 Digits(34)=Array(ae18, ae16, aec1, aed1, ae19, ae17, ae15, aef1, ae11, ae13, ae12, ae14, aeb1, aea1, ae10, aee1)
 Digits(35)=Array(af18, af16, afc1, afd1, af19, af17, af15, aff1, af11, af13, af12, af14, afb1, afa1, af10, afe1)
 Digits(36)=Array(b9, b7, b0c1, b0d1, b100, b8, b6, b0f1, b2, b4, b3, b5, b0b1, b0a1, b1,b0e1)
 Digits(37)=Array(b109, b107, b1c1, b1d1, b110, b108, b106, b1f1, b102, b104, b103, b105, b1b1, b1a1, b101,b1e1)
 Digits(38)=Array(b119, b117, b2c1, b2d1, b120, b118, b116, b2f1, b112, b114, b113, b115, b2b1, b2a1, b111, b2e1)
 Digits(39)=Array(b129, b127, b3c1, b3d1, b130, b128, b126, b3f1, b122, b3b1, b123, b125, b3b1, b3a1, b121, b3e1)


 Sub DisplayTimer_Timer
  Dim ChgLED, ii, jj, num, chg, stat, obj, b, x
    ChgLED=Controller.ChangedLEDs(&Hffffffff, &Hffffffff)
    If Not IsEmpty(ChgLED)Then
    If Table1.ShowDT = true Then
      For ii=0 To UBound(chgLED)
        num=chgLED(ii, 0) : chg=chgLED(ii, 1) : stat=chgLED(ii, 2)
        if (num < 40) then
          For Each obj In Digits(num)
            If chg And 1 Then obj.State=stat And 1
            chg=chg\2 : stat=stat\2
          Next
        else
          end if
      Next
    End if
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

InitLamps()             ' turn off the lights and flashers and reset them to the default parameters
LampTimer.Interval = 5 'lamp fading speed
LampTimer.Enabled = 1

' Lamp & Flasher Timers

Sub LampTimer_Timer()
    Dim chgLamp, num, chg, ii,loop0
    chgLamp = Controller.ChangedLamps
    If Not IsEmpty(chgLamp) Then
        For ii = 0 To UBound(chgLamp)
            LampState(chgLamp(ii, 0) ) = chgLamp(ii, 1)       'keep the real state in an array
            FadingLevel(chgLamp(ii, 0) ) = chgLamp(ii, 1) + 4 'actual fading step

     'Special Handling
     If chgLamp(ii,0) = 0 Then SolRelayQ chgLamp(ii,1) 'InGame Q Relay
     If chgLamp(ii,0) = 1 Then SolRelayT chgLamp(ii,1) 'Tilt T Relay
     If chgLamp(ii,0) = 2 Then SolBallRelease chgLamp(ii,1)
       'If chgLamp(ii,0) = 4 Then SolSound16 chgLamp(ii,1) 'Sound 16?
     If chgLamp(ii,0) = 12 Then SolBallShooter chgLamp(ii,1)
     If chgLamp(ii,0) = 13 Then SolRelayS chgLamp(ii,1) 'Switching S Relay
     'If chgLamp(ii,0) = 14 Then SolRelayA chgLamp(ii,1) 'Playboard Aux Lamp A Relay
     If chgLamp(ii,0) = 15 Then SolRelayB chgLamp(ii,1) 'Lightbox Aux Lamp B Relay  aka sequentitial light show
     If chgLamp(ii,0) = 16 Then SolLeftBallGate chgLamp(ii,1)
     If chgLamp(ii,0) = 17 Then SolRightBallGate chgLamp(ii,1)
     'If chgLamp(ii,0) = 18 Then SolRelayC chgLamp(ii,1) 'Skull Motor C Relay
     'If chgLamp(ii,0) = 19 Then Topper chgLamp(ii,1) ' Skull Jaw Relay
        Next
    End If
    UpdateLamps
End Sub


Sub UpdateLamps()
  NFadeL 3, L3
  NFadeL 5, L5
  NFadeL 6, L6
  NFadeL 7, L7
  NFadeL 8, L8
  NFadeL 9, L9
  NFadeL 10, L10
  NFadeL 11, L11
  NFadeL 20, L20
  If Controller.lamp(13)=True Then 'S Relay enabled
    NFadeL 21, L21a
  Else
    NFadeL 21, L21
  End If
  If Controller.lamp(13)=True Then 'S Relay enabled
    NFadeL 22, L22a
  Else
    NFadeL 22, L22
  End If
  NFadeL 23, L23
  NFadeL 24, L24
  NFadeL 25, L25
  NFadeL 26, L26
  NFadeL 27, L27
  NFadeL 28, L28
  NFadeLm 29, L29
  NFadeL 29, L29a
  NFadeL 30, L30
  NFadeL 31, L31
  NFadeL 32, L32
  NFadeL 33, L33
  NFadeL 34, L34
  NFadeL 35, L35
  NFadeL 36, L36
  NFadeL 37, L37
  NFadeL 38, L38
  NFadeL 39, L39
  NFadeL 40, L40
  NFadeL 41, L41
  NFadeL 42, L42
  NFadeL 43, L43
  NFadeL 44, L44
  NFadeL 45, L45
  NFadeL 46, L46
  NFadeL 47, L47
  NFadeLm 49, L49
  NFadeL 49, L49a
  NFadeL 50, L50
  NFadeL 51, L51

   'Solenoid controlled lights
  NFadeObjm 100, S100Dome,"ronddomebaseredlit", "ronddomebasered"
  NFadeLm 100, S1a
  Flash   100, S1Fa
  NFadeObjm 101, S101Dome,"ronddomebaseredlit", "ronddomebasered"
  NFadeLm 101, S1b
  Flash   101, S1Fb

  NFadeLm 102, S2a
  NFadeLm 102, S2b
  Flash   102, S2F

  NFadeLm 103, S3
  Flash   103, S3F

  NFadeL  105, S5

  NFadeLm 106, S6a
  NFadeLm 106, S6b
  Flash   106, S6F

  If TestLamps=1 Then
    NFadeL  0, L0
    NFadeL  1, L1
    NFadeL  2, L2
    NFadeL  4, L4
    NFadeL  12, L12
    NFadeL  13, L13
    NFadeL  14, L14
    NFadeL  15, L15
    NFadeL  16, L16
    NFadeL  17, L17
    NFadeL  18, L18
    NFadeL  19, L19
    NFadeL  104, Sol4
    NFadeL  107, Sol7
    NFadeL  109, Sol9
  End If

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


' *******************************************************************************************************
' Positional Sound Playback Functions by DJRobX and Rothbauerw
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

Sub PlaySoundAtBOTBallZ(sound, BOT)
  PlaySound sound, 0, ABS(BOT.velz)/17, AudioPan(BOT), 0, Pitch(BOT), 1, 0, AudioFade(BOT)
End Sub


'*********************************************************************
'                     Supporting Ball & Sound Functions
'*********************************************************************

Function RndNum(min, max)
  RndNum = Int(Rnd() * (max-min + 1) ) + min ' Sets a random number between min and max
End Function

Function AudioFade(tableobj) ' Fades between front and back of the table (for surround systems or 2x2 speakers, etc), depending on the Y position on the table. "table1" is the name of the table
  Dim tmp
  On Error Resume Next
  tmp = tableobj.y * 2 / table1.height-1
  If tmp > 0 Then
    AudioFade = Csng(tmp ^10)
  Else
    AudioFade = Csng(-((- tmp) ^10) )
  End If
End Function

Function AudioPan(tableobj) ' Calculates the pan for a tableobj based on the X position on the table. "table1" is the name of the table
  Dim tmp
  On Error Resume Next
  tmp = tableobj.x * 2 / table1.width-1
  If tmp > 0 Then
    AudioPan = Csng(tmp ^10)
  Else
    AudioPan = Csng(-((- tmp) ^10) )
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

Function BallVelZ(ball) 'Calculates the ball speed in the -Z
  BallVelZ = INT((ball.VelZ) * -1 )
End Function

Function VolZ(ball) ' Calculates the Volume of the sound based on the ball speed in the Z
  VolZ = Csng(BallVelZ(ball) ^2 / 200)*1.2
End Function



'********************************************************************
'      JP's VP10 Rolling Sounds (+rothbauerw's Dropping Sounds)
'********************************************************************

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
    Dim BOT, b, ballpitch, ballvol
    BOT = GetBalls

    ' stop the sound of deleted balls
    For b = UBound(BOT) + 1 to tnob
        rolling(b) = False
        StopSound("fx_ballrolling" & b)
    Next

    ' exit the sub if no balls on the table
    If UBound(BOT) = -1 Then Exit Sub

    For b = 0 to UBound(BOT)
        ' play the rolling sound for each ball
        If BallVel(BOT(b))> 1 Then
            If BOT(b).z <30 Then
                ballpitch = Pitch(BOT(b))
                ballvol = Vol(BOT(b))
            Else
                ballpitch = Pitch(BOT(b)) + 25000 'increase the pitch on a ramp
                ballvol = Vol(BOT(b)) * 10 * VolRamp
            End If
            rolling(b) = True
            PlaySound("fx_ballrolling" & b), -1, ballvol*VolRol, AudioPan(BOT(b)), 0, ballpitch, 1, 0, AudioFade(BOT(b))
        Else
            If rolling(b) = True Then
                StopSound("fx_ballrolling" & b)
                rolling(b) = False
            End If
        End If

        ' play ball drop sounds
        If BOT(b).VelZ < -1 and BOT(b).z < 55 and BOT(b).z > 27 Then 'height adjust for ball drop sounds
            PlaySound("fx_ball_drop" & b), 0, ABS(BOT(b).velz)/17, AudioPan(BOT(b)), 0, Pitch(BOT(b)), 1, 0, AudioFade(BOT(b))
        End If
    Next
End Sub


'**********************
' Ball Collision Sound
'**********************

Sub OnBallBallCollision(ball1, ball2, velocity)
  PlaySound("fx_collide"), 0, Csng(velocity) ^2 / 2000 * VolCol, AudioPan(ball1), 0, Pitch(ball1), 0, 0, AudioFade(ball1)
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
            BallShadow(b).X = ((BOT(b).X) - (Ballsize/6) + ((BOT(b).X - (Table1.Width/2))/7)) + 6
        Else
            BallShadow(b).X = ((BOT(b).X) + (Ballsize/6) + ((BOT(b).X - (Table1.Width/2))/7)) - 6
        End If
        ballShadow(b).Y = BOT(b).Y + 12
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
  PlaySound "pinhit_low", 0, Vol(ActiveBall)*VolPin, AudioPan(ActiveBall), 0, Pitch(ActiveBall), 0, 0, AudioFade(ActiveBall)
End Sub

Sub Targets_Hit (idx)
  PlaySound "target", 0, Vol(ActiveBall)*VolTarg, AudioPan(ActiveBall), 0, Pitch(ActiveBall), 0, 0, AudioFade(ActiveBall)
End Sub

Sub Metals_Thin_Hit (idx)
  PlaySound "metalhit_thin", 0, Vol(ActiveBall)*VolMetal, AudioPan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
End Sub

Sub Metals_Medium_Hit (idx)
  PlaySound "metalhit_medium", 0, Vol(ActiveBall)*VolMetal, AudioPan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
End Sub

Sub Metals2_Hit (idx)
  PlaySound "metalhit2", 0, Vol(ActiveBall)*VolMetal, AudioPan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
End Sub

Sub Gates_Hit (idx)
  PlaySound "gate", 0, Vol(ActiveBall)*VolGates, AudioPan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
End Sub

Sub Woods_Hit (idx)
  PlaySound "fx_woodhit", 0, Vol(ActiveBall)*VolWood, AudioPan(ActiveBall), 0, Pitch(ActiveBall), 0, 0, AudioFade(ActiveBall)
End Sub

Sub Spinners_Spin
  PlaySound "fx_spinner", 0, .25*VolSpin, AudioPan(Spinner), 0.25, 0, 0, 1, AudioFade(Spinner)
End Sub

Sub Rubbers_Hit(idx)
  dim finalspeed
    finalspeed=SQR(activeball.velx * activeball.velx + activeball.vely * activeball.vely)
  If finalspeed > 20 then
    PlaySound "fx_rubber_band", 0, Vol(ActiveBall)*VolRub, AudioPan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
  End if
  If finalspeed >= 6 AND finalspeed <= 20 then
    RandomSoundRubber()
  End If
End Sub

Sub Posts_Hit(idx)
  dim finalspeed
    finalspeed=SQR(activeball.velx * activeball.velx + activeball.vely * activeball.vely)
  If finalspeed > 16 then
    PlaySound "fx_rubber_post", 0, Vol(ActiveBall)*VolRub, AudioPan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
  End if
  If finalspeed >= 6 AND finalspeed <= 16 then
    RandomSoundRubber()
  End If
End Sub

Sub RandomSoundRubber()
  Select Case Int(Rnd*3)+1
    Case 1 : PlaySound "rubber_hit_1", 0, Vol(ActiveBall)*VolRub, AudioPan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
    Case 2 : PlaySound "rubber_hit_2", 0, Vol(ActiveBall)*VolRub, AudioPan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
    Case 3 : PlaySound "rubber_hit_3", 0, Vol(ActiveBall)*VolRub, AudioPan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
  End Select
End Sub

Sub FlipperLeft_Collide(parm)
  RandomSoundFlipper()
End Sub

Sub FlipperRight_Collide(parm)
  RandomSoundFlipper()
End Sub

Sub FlipperLeftUp_Collide(parm)
  RandomSoundFlipper()
End Sub

Sub FlipperRightUp_Collide(parm)
  RandomSoundFlipper()
End Sub

Sub BallGate_Collide(parm)
  RandomSoundFlipper()
End Sub

Sub RandomSoundFlipper()
  Select Case Int(Rnd*3)+1
    Case 1 : PlaySound "flip_hit_1", 0, Vol(ActiveBall)*VolFlipH, AudioPan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
    Case 2 : PlaySound "flip_hit_2", 0, Vol(ActiveBall)*VolFlipH, AudioPan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
    Case 3 : PlaySound "flip_hit_3", 0, Vol(ActiveBall)*VolFlipH, AudioPan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
  End Select
End Sub

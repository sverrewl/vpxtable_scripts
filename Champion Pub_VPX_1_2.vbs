'############################################################################################
'############################################################################################
'#######                                                                             ########
'#######          The Champion Pub                                                   ########
'#######          (Bally 1998)                                                       ########
'#######                                                                             ########
'############################################################################################
'############################################################################################
' Version 1.2 mfuegemann 2017
'
' The original VP8 version was created by UncleReamus, Destruk, TomB and Scapino. This version is based on my
' VP9 conversion of that table.
'
' Thanks go to:
' Fuzzel for the awesome 3D models of the Boxer, Wire Ramps, Jump Ramps, Speedbag/Fists, Rope and Flippers.
' Dark for the Spot Light models, the ramp triggers, the catapult and the Flasher Sphere
' Zany for the regular Flasher model
'
' Version 1.1:
' - adjusted Flipper length and position (found by Thalamus)
' - addad B2S call for direct backglass support (activate only if You are using my B2S)
' - added Settings.Value("ddraw") = 0 in table_init, uncomment if need be (table crashes in FullScreen mode)
'
' Version 1.2:
' - adjusted some rubber hit heights (found by Thalamus)
' - added scoop returning ball recognition (found by hanzoverfist)
' - added some DOF calls (found by Arngrim)
' - reviewed material and GI settings (hauntfreaks)
Option Explicit

' Thalamus 2018-07-20
' Added/Updated "Positional Sound Playback Functions" and "Supporting Ball & Sound Functions"
' Changed UseSolenoids=1 to 2
' Thalamus 2018-08-09 : Improved directional sounds
' !! NOTE : Table not verified yet !!

' Options
' Volume devided by - lower gets higher sound

Const VolDiv = 2000    ' Lower number, louder ballrolling/collition sound
Const VolCol = 10      ' Ball collition divider ( voldiv/volcol )

' The rest of the values are multipliers
'
'  .5 = lower volume
' 1.5 = higher volume

Const VolGates  = 1    ' Gates volume.
Const VolMetal  = 1    ' Metals volume.
Const VolRH     = 1    ' Rubber hits volume.
Const VolPi     = 1    ' Rubber pins volume.
Const VolTarg   = 1    ' Targets volume.
Const VolFlip   = 1    ' Flipper volume.


On Error Resume Next
ExecuteGlobal GetTextFile("controller.vbs")
If Err Then MsgBox "You need the controller.vbs in order to run this table, available in the vp10 package"
On Error Goto 0

' Thal : Added because of useSolenoid=2
Const cSingleLFlip = 0
Const cSingleRFlip = 0

Const UseSolenoids=2,UseLamps=1,SSolenoidOn="SolOn",SSolenoidOff="SolOff"
Const SCoin="coin3"
Const UseVPMModSol=1

'-----------------------------------
' Configuration
'-----------------------------------
Const DimGI=-2          'set to dim or brighten GI lights (minus is darker, base value is 8)
Const LeftOutlanePost=0     '0=Easy, 1=Medium, 2=Hard
Const DangerZoneMod=1     '0=Mod disabled, 1=Mod installed (the Mod adds a Gate above BEER Targets instead of a Post)
Const RopeLevel=0       'Rope Jump Difficulty: 0=Easy, 1=Medium, 2=Hard
'Const BallSize = 50      'default 50, 51 plays ok, with 52 the ball will get stuck

Const UseB2SBG=1        'set to 1, if You are using my B2S Backglass for direct B2S communication

Dim cGameName
cGameName = "cp_16"
LoadVPM "01560000", "WPC.VBS", 3.26

'-----------------------------------
'Solenoid Routines
'-----------------------------------
'SolCallback --> SolModCallback for Flashers
SolCallback(1)  = "SolCatapult"
SolCallback(2)  = "SolTrough"
SolCallback(5)  = "SolCornerKickout"
SolCallback(8)  = "SolPostDiverter"
SolCallback(9)  = "SolLeftScoop"
SolCallback(10) = "SolRightScoop"
SolCallback(12) = "SolPost"
SolCallback(14) = "SolPopper"
SolModCallback(17) = "sol17"        'Flasher
SolModCallback(18) = "sol18"        'Flasher
SolModCallback(19) = "UpperWhiteFlasher"    'Flasher
SolModCallback(20) = "UpperRedFlasher"    'Flasher
SolModCallback(21) = "LowerRedFlasher"    'Flasher
SolModCallback(22) = "sol22"        'Flasher
SolModCallback(23) = "SolRopeSpot"      'Flasher
SolModCallback(24) = "SolSpeedBagSpot"    'Flasher
SolCallback(28) = "SolLockPin"
SolCallback(33) = "SolMagnetPopper"
SolCallback(34) = "SolRampDiverter"
SolCallback(35) = "SolLeftSP"
SolCallback(36) = "SolRightSP"
SolCallback(sLLFlipper) = "SolLFlipper"
SolCallback(sLRFlipper) = "SolRFlipper"

Set GIcallback2 = GetRef("UpdateGI")     'GICallback2 is providing the GI intesity

Sub SolTrough(Enabled)
  If Enabled then
    bsTrough.ExitSol_On
    vpmTimer.PulseSw 31
    End If
End Sub

Sub SolPostDiverter(Enabled)
    If Enabled Then
    playsoundAtVol SoundFX("solon",DOFContactors), EnterLockup, 1
        PostDiverter.IsDropped=0
    Else
    playsoundAtVol SoundFX("soloff",DOFContactors), EnterLockup, 1
        PostDiverter.IsDropped=1
    End If
End Sub

Sub SolLeftScoop(Enabled)
    If Enabled then
    playsoundAtVol SoundFX("solon",DOFContactors), P_LeftScoop, 1
    P_LeftScoop.roty = 16
    P_LScoopMech1.rotx = -60
    P_LScoopMech2.rotx = -10
    P_LScoopMech2.TransZ = -33
    P_LScoopMech2.Transy = 10
        LeftScoopGrab.Enabled=1
        Controller.Switch(61)=1
    LeftScoopBack.isdropped = False
    Else
    playsoundAtVol SoundFX("soloff",DOFContactors),P_LeftScoop, 1
    P_LeftScoop.roty = 52
    P_LScoopMech1.rotx = -80
    P_LScoopMech2.rotx = 0
    P_LScoopMech2.TransZ = 0
    P_LScoopMech2.Transy = 0
        LeftScoopGrab.Enabled=0
        Controller.Switch(61)=0
    LeftScoopBack.isdropped = True
    End if
End Sub

Sub SolRightScoop(Enabled)
    If Enabled Then
    playsoundAtVol SoundFX("solon",DOFContactors), P_RightScoop, 1
    P_RightScoop.roty = 16
    P_RScoopMech1.rotx = -60
    P_RScoopMech2.rotx = -10
    P_RScoopMech2.TransZ = -33
    P_RScoopMech2.Transy = 10
        RightScoopGrab.Enabled=1
        Controller.Switch(62)=1
    RightScoopBack.isdropped = False
    Else
    playsoundAtVol SoundFX("soloff",DOFContactors), P_RightScoop, 1
    P_RightScoop.roty = 52
    P_RScoopMech1.rotx = -80
    P_RScoopMech2.rotx = 0
    P_RScoopMech2.TransZ = 0
    P_RScoopMech2.Transy = 0
        RightScoopGrab.Enabled=0
        Controller.Switch(62)=0
    RightScoopBack.isdropped = True
    End if
End Sub

Sub SolPost(Enabled)
    If Enabled Then
    playsoundAtVol SoundFX("solon",DOFContactors), P_Post, 1
    P_Post.Transy = 45
        Controller.Switch(75)=1
        ForceField.IsDropped=0
    Else
    playsoundAtVol SoundFX("soloff",DOFContactors), P_Post, 1
    P_Post.Transy = 0
        Controller.Switch(75)=0
    ForceField.IsDropped=1
    End If
End Sub

Sub SolLockPin(Enabled)
    If Enabled then
    playsoundAtVol SoundFX("solon",DOFContactors), Sw15, 1
        LockPin.IsDropped=1
    Else
        LockPin.TimerEnabled=1
    End If
End Sub

Sub LockPin_Timer
  playsoundAtVol SoundFX("soloff",DOFContactors), Sw15, 1
    LockPin.IsDropped=0
    LockPin.TimerEnabled=0
End Sub

Sub SolRampDiverter(Enabled)
    If Enabled Then
    playsoundAtVol SoundFX("solon",DOFContactors), P_Diverter, 1
    P_Diverter.Transz = 60
        DiverterClosed.IsDropped = True
    DivPost1a.Isdropped = False
    DivPost2a.Isdropped = False
    Else
    playsoundAtVol SoundFX("soloff",DOFContactors), P_Diverter, 1
    P_Diverter.Transz = 0
        DiverterClosed.IsDropped = False
    DivPost1a.Isdropped = True
    DivPost2a.Isdropped = True
    End If
End Sub

'Corner Kickout
Sub SolCornerKickout(enabled)
  if enabled then
    if Controller.Switch(37) then
      bsCornerKickout.ExitSol_On
      P_CornerKickout.Transy = -40
      CornerKickout.Timerenabled = True
    end If
  end if
End Sub
Sub CornerKickout_Timer
  CornerKickout.Timerenabled = False
  P_CornerKickout.Transy = 0
End Sub

'Catapult
Sub SolCatapult(enabled)
  if enabled then
    FireCatapult
    if Controller.Switch(18) then
      bsCatapult.ExitSol_On
    end If
  end if
End Sub

Dim CatapultDir
Sub FireCatapult
  playsoundAtVol SoundFX("solon",DOFContactors), catapultLaunchKicker, 1
  catapultLaunchKicker.Timerenabled = False
  CatapultDir = 1.5
  catapultLaunchKicker.Timerenabled = True
End Sub

Sub catapultLaunchKicker_Timer
  P_Catapult.rotx = P_Catapult.rotx + CatapultDir
  if P_Catapult.rotx > 90 then
    P_Catapult.rotx = 90
    CatapultDir = -0.5
  end if
  if P_Catapult.rotx < 5 then
    catapultLaunchKicker.Timerenabled = False
    P_Catapult.rotx = 5
    CatapultDir = 0
  end if
end Sub

'---------------------------------------------------------------------------
' Flasher
'---------------------------------------------------------------------------
Const JabSpot1Intensity=16     'Spot Center
Const JabSpotIntensity=6

Dim Prev17Int,Prev18Int,Prev22Int,PrevUWInt,PrevLRInt,PrevURInt,PrevSpeedBagInt,PrevRopeSpotInt

Sub Sol17(Intensity)  'White Boxer light and spots  - max 154
  if Intensity <> Prev17Int Then
    if Intensity > 0 then
      flash17.state = Lightstateon
      flash17.Intensity = 5*(Intensity/154)
      RightJabSpot1.state = Lightstateon
      LeftJabSpot1.state = Lightstateon
      RightJabSpot1.Intensity = JabSpot1Intensity*(Intensity/154)
      LeftJabSpot1.Intensity = JabSpot1Intensity*(Intensity/154)
      If DesktopMode = True Then
        RightJabSpot_DT.state = Lightstateon
        LeftJabSpot_DT.state = Lightstateon
        RightJabSpot_DT.Intensity = JabSpotIntensity*(Intensity/154)
        LeftJabSpot_DT.Intensity = JabSpotIntensity*(Intensity/154)
      Else
        RightJabSpot.state = Lightstateon
        LeftJabSpot.state = Lightstateon
        RightJabSpot.Intensity = JabSpotIntensity*(Intensity/154)
        LeftJabSpot.Intensity = JabSpotIntensity*(Intensity/154)
      end If
    else
      flash17.state = Lightstateoff
      RightJabSpot.state = Lightstateoff
      RightJabSpot1.state = Lightstateoff
      LeftJabSpot.state = Lightstateoff
      LeftJabSpot1.state = Lightstateoff
      RightJabSpot_DT.state = Lightstateoff
      LeftJabSpot_DT.state = Lightstateoff
    end If
  end If
  Prev17Int = Intensity
End Sub

Sub Sol18(Intensity)   'Danger Zone Bolts
  if Intensity <> Prev18Int Then
    if Intensity > 0 then
      flasher18.state = Lightstateon
      flasher18.Intensity = 18*(Intensity/154)
      flasher18b.state = Lightstateon
      flasher18b.Intensity = 18*(Intensity/154)
    else
      flasher18.state = Lightstateoff
      flasher18b.state = Lightstateoff
    end if
  end If
  Prev18Int = Intensity
End Sub

Sub Sol22(Intensity)   'Boxer Bolts
  if Intensity <> Prev22Int Then
    if Intensity > 0 then
      bolt1.state = Lightstateon
      bolt1.Intensity = 8*(Intensity/154)
      bolt2.state = Lightstateon
      bolt2.Intensity = 8*(Intensity/154)
    else
      bolt1.state = Lightstateoff
      bolt2.state = Lightstateoff
    end if
  end If
  if UseB2SBG then
    if Intensity > 0 then
      Controller.B2SSetData 222,1
    Else
      Controller.B2SSetData 222,0
    End If
  End If
  Prev22Int = Intensity
End Sub

Sub SolSpeedBagSpot(Intensity)
  if Intensity <> PrevSpeedBagInt Then
    if Intensity > 0 then
      If DesktopMode = True Then
        SpeedBagFlasher_DT.state = Lightstateon
        SpeedBagFlasher_DT.Intensity = 12*(Intensity/154)
        SpeedBagFlasher1_DT.state = Lightstateon
        SpeedBagFlasher1_DT.Intensity = 18*(Intensity/154)
      Else
        SpeedBagFlasher.state = Lightstateon
        SpeedBagFlasher.Intensity = 15*(Intensity/154)
        SpeedBagFlasher1.state = Lightstateon
        SpeedBagFlasher1.Intensity = 18*(Intensity/154)
      End If
    else
      SpeedBagFlasher.state = Lightstateoff
      SpeedBagFlasher1.state = Lightstateoff
      SpeedBagFlasher_DT.state = Lightstateoff
      SpeedBagFlasher1_DT.state = Lightstateoff
    end if
  end If
  PrevSpeedBagInt = Intensity
End Sub

Sub SolRopeSpot(Intensity)
  if Intensity <> PrevRopeSpotInt Then
    if Intensity > 0 then
      If DesktopMode = True Then
        RopeFlasher_DT.state = Lightstateon
        RopeFlasher_DT.Intensity = 12*(Intensity/154)
        RopeFlasher1_DT.state = Lightstateon
        RopeFlasher1_DT.Intensity = 18*(Intensity/154)
      Else
        RopeFlasher.state = Lightstateon
        RopeFlasher.Intensity = 15*(Intensity/154)
        RopeFlasher1.state = Lightstateon
        RopeFlasher1.Intensity = 18*(Intensity/154)
      End If
    else
      RopeFlasher.state = Lightstateoff
      RopeFlasher1.state = Lightstateoff
      RopeFlasher_DT.state = Lightstateoff
      RopeFlasher1_DT.state = Lightstateoff
    end if
  end If
  PrevRopeSpotInt = Intensity
End Sub

Sub UpperWhiteFlasher(Intensity)
  if Intensity <> PrevUWInt Then
    if Intensity > 0 then
      If DesktopMode = True Then
        TopFlasher1_DT.state = Lightstateon
        TopFlasher1_DT.Intensity = 15*(Intensity/154)
      Else
        TopFlasher1.state = Lightstateon
        TopFlasher1.Intensity = 15*(Intensity/154)
        TopFlasher2.state = Lightstateon
        TopFlasher2.Intensity = 6*(Intensity/154)
      end If
      P_TopFlasher.image = "dome3_clear_lit"
    else
      TopFlasher1.state = Lightstateoff
      TopFlasher2.state = Lightstateoff
      TopFlasher1_DT.state = Lightstateoff
      P_TopFlasher.image = "dome3_clear"
    end if
  end If
  if UseB2SBG then
    if Intensity > 0 then
      Controller.B2SSetData 219,1
    Else
      Controller.B2SSetData 219,0
    End If
  End If
  PrevUWInt = Intensity
End Sub

Sub UpperRedFlasher(Intensity)
  if Intensity <> PrevURInt Then
    if Intensity > 0then
      RightFlasher1.state = Lightstateon
      RightFlasher2.state = Lightstateon
      RightFlasher3.state = Lightstateon
      RightFlasher1.state = 12*(Intensity/154)
      RightFlasher2.state = 12*(Intensity/154)
      RightFlasher3.state = 12*(Intensity/154)
      P_RightFlasher.image = "TOPFlasherRED_lit"
    else
      RightFlasher1.state = Lightstateoff
      RightFlasher2.state = Lightstateoff
      RightFlasher3.state = Lightstateoff
      P_RightFlasher.image = "TOPFlasherRED_unlit"
    end if
  end If
  if UseB2SBG then
    if Intensity > 0 then
      Controller.B2SSetData 220,1
    Else
      Controller.B2SSetData 220,0
    End If
  End If
  PrevURInt = Intensity
End Sub

Sub LowerRedFlasher(Intensity)
  if Intensity <> PrevLRInt Then
    if Intensity > 0 then
      If DesktopMode = True Then
        LeftFlasher1_DT.state = Lightstateon
        LeftFlasher2_DT.state = Lightstateon
        LeftFlasher3_DT.state = Lightstateon
        LeftFlasher1_DT.state = 12*(Intensity/154)
        LeftFlasher2_DT.state = 12*(Intensity/154)
        LeftFlasher3_DT.state = 12*(Intensity/154)
      else
        LeftFlasher1.state = Lightstateon
        LeftFlasher2.state = Lightstateon
        LeftFlasher3.state = Lightstateon
        LeftFlasher1.state = 12*(Intensity/154)
        LeftFlasher2.state = 12*(Intensity/154)
        LeftFlasher3.state = 12*(Intensity/154)
      end If
      P_LeftFlasher.image = "dome3_red_lit"
    else
      LeftFlasher1.state = Lightstateoff
      LeftFlasher2.state = Lightstateoff
      LeftFlasher3.state = Lightstateoff
      LeftFlasher1_DT.state = Lightstateoff
      LeftFlasher2_DT.state = Lightstateoff
      LeftFlasher3_DT.state = Lightstateoff
      P_LeftFlasher.image = "dome3_red"
    end if
  end If
  if UseB2SBG then
    if Intensity > 0 then
      Controller.B2SSetData 221,1
    Else
      Controller.B2SSetData 221,0
    End If
  End If
  PrevLRInt = Intensity
End Sub


'-----------------------------
'------  VUK animation  ------
'-----------------------------
Const PopTimerInterval=40
Sub Popper_Hit
  Controller.Switch(28) = 1
End Sub

Sub SolPopper(Enabled)
  If Enabled Then
      If Controller.Switch(28) = True Then
    PlaySoundAtVol SoundFX("Kicker",DOFContactors), Vuk1, 1
    Controller.Switch(28) = 0
    VUK1.CreateBall
    Popper.destroyball
    'vpmTimer.AddTimer PopTimerInterval,"VUKLevel1"
    VUKLevel = 1
    VUKTimer.enabled = True
    end if
  end if
end sub

Dim VUKLevel
Sub VUKTimer_Timer
  select case VUKlevel
    case 1: VUK2.CreateBall
        VUK1.DestroyBall
    case 2: VUK3.CreateBall
        VUK2.DestroyBall
    case 3: VUK4.CreateBall
        VUK3.DestroyBall
    case 4: VUK5.CreateBall
        VUK4.DestroyBall
    case 5: VUK6.CreateBall
        VUK5.DestroyBall
    case 6: VUK7.CreateBall
        VUK6.DestroyBall
    case 8: VUK8.CreateBall
        VUK7.DestroyBall
    case 9: VUKTop.CreateBall
        VUK8.DestroyBall
        VUKTop.Kick 340,8
        VUKTimer.enabled = False
  end Select
  VUKLevel = VUKLevel + 1
End Sub


Sub VUKLevel1(swNo)
  VUK2.CreateBall
  VUK1.DestroyBall
  vpmTimer.AddTimer PopTimerInterval,"VUKLevel2"
End Sub

Sub VUKLevel2(swNo)
  VUK3.CreateBall
  VUK2.DestroyBall
  vpmTimer.AddTimer PopTimerInterval,"VUKLevel3"
End Sub

Sub VUKLevel3(swNo)
  VUK4.CreateBall
  VUK3.DestroyBall
  vpmTimer.AddTimer PopTimerInterval,"VUKLevel4"
End Sub

Sub VUKLevel4(swNo)
  VUK5.CreateBall
  VUK4.DestroyBall
  vpmTimer.AddTimer PopTimerInterval,"VUKLevel5"
End Sub

Sub VUKLevel5(swNo)
  VUK6.CreateBall
  VUK5.DestroyBall
  vpmTimer.AddTimer PopTimerInterval,"VUKLevel6"
End Sub

Sub VUKLevel6(swNo)
  VUK7.CreateBall
  VUK6.DestroyBall
  vpmTimer.AddTimer PopTimerInterval,"VUKLevel7"
End Sub

Sub VUKLevel7(swNo)
  VUK8.CreateBall
  VUK7.DestroyBall
  vpmTimer.AddTimer PopTimerInterval,"VUKLevel8"
End Sub

Sub VUKLevel8(swNo)
  VUKTop.CreateBall
  VUK8.DestroyBall
  VUKTop.Kick 320,6
End Sub


'---------------------------------------------------------------------------
' Table Init
'---------------------------------------------------------------------------
Dim bsTrough,bsCornerKickout,bsCatapult,mRope,magRopeMagnet
Dim DesktopMode: DesktopMode = CP.ShowDT

Sub CP_Init
  vpmInit Me
    With Controller
    .GameName = cGameName
        'If Err Then MsgBox "Can't start Game " & cGameName & vbNewLine & Err.Description:Exit Sub
        .SplashInfoLine = "Champion Pub"
        .HandleKeyboard = 0
        .ShowTitle = 0
        .ShowDMDOnly = 1
        .ShowFrame = 0
        .HandleMechanics = 0
        .Hidden = 0
    .Dip(0) = &H00

    '.Games(cGameName).Settings.Value("samples")=0

      'DMD position for 3 Monitor Setup
'   Controller.Games(cGameName).Settings.Value("dmd_pos_x")=500
'   Controller.Games(cGameName).Settings.Value("dmd_pos_y")=0
'   Controller.Games(cGameName).Settings.Value("dmd_width")=505
'   Controller.Games(cGameName).Settings.Value("dmd_height")=155
'   Controller.Games(cGameName).Settings.Value("rol")=0

'   Controller.Games(cGameName).Settings.Value("ddraw") = 0             'set to 0 if You have problems with DMD showing or table stutter or FullScreen crashes

        On Error Resume Next
        .Run GetPlayerHWnd
        If Err Then MsgBox Err.Description
        On Error Goto 0
  End With

    On Error Goto 0

    PinMAMETimer.Interval=PinMAMEInterval

    vpmNudge.TiltSwitch=14
    vpmNudge.Sensitivity=2
    vpmNudge.TiltObj=Array(LeftSlingshot,RightSlingshot)

    vpmMapLights AllLights

    Set bsTrough=New cvpmBallStack
    bsTrough.InitSw 0,32,33,34,35,0,0,0
        bsTrough.InitKick BallRelease,40,8
        bsTrough.InitExitSnd SoundFX("BallRel",DOFContactors),SoundFX("Solenoid",DOFContactors)
        bsTrough.Balls=4

  'Plunger lane
    Set bsCatapult = New cvpmSaucer
    bsCatapult.InitKicker CatapultLaunchKicker,18,0,42,0
    bsCatapult.InitSounds SoundFX("kicker_enter_center",DOFContactors),SoundFX("solon",DOFContactors),SoundFX("popper_ball",DOFContactors)

  'Left Kicker
    Set bsCornerKickout=New cvpmBallStack
        bsCornerKickout.InitSaucer CornerKickout,37,155,5
        bsCornerKickout.InitExitSnd SoundFX("popper",DOFContactors),SoundFX("solon",DOFContactors)

    Set mRope=New cvpmMech
        mRope.MType=vpmMechOneSol+vpmMechCircle+vpmMechLinear+vpmMechFast
        mRope.Sol1=25
        mRope.Length=300                       '280-300 duration of one turn in 1/60 seconds
        mRope.Steps=720
        mRope.Callback=GetRef("UpdateRope")
        mRope.Start

    Set magRopeMagnet=New cvpmMagnet
        magRopeMagnet.InitMagnet TrigMagnet,39  '35
        magRopeMagnet.Solenoid=7
    magRopeMagnet.Size=60
        magRopeMagnet.GrabCenter=False
        magRopeMagnet.CreateEvents "magRopeMagnet"

  Controller.Switch(61)=1                     'Left Scoop Up
  Controller.Switch(62)=1                     'Right Scoop Up
  Controller.Switch(22)=1                     'close coin door
  vpmTimer.PulseSw 45           'Rope

  'Features
  If DesktopMode = True Then 'Show Desktop components
    SideWood.visible=1
    LeftRail.visible=1
    RightRail.visible=1
  Else
    SideWood.visible=0
    LeftRail.visible=0
    RightRail.visible=0
  End if

  PLeftOutlane_Easy.visible = (LeftOutlanePost <> 1) and (LeftOutlanePost <> 2)
  RLeftOutlane_Easy.visible = (LeftOutlanePost <> 1) and (LeftOutlanePost <> 2)
  RLeftOutlane_Easy.Collidable = (LeftOutlanePost <> 1) and (LeftOutlanePost <> 2)
  PLeftOutlane_Med.visible = (LeftOutlanePost=1)
  RLeftOutlane_Med.visible = (LeftOutlanePost=1)
  RLeftOutlane_Med.Collidable = (LeftOutlanePost=1)
  PLeftOutlane_Hard.visible = (LeftOutlanePost=2)
  RLeftOutlane_Hard.visible = (LeftOutlanePost=2)
  RLeftOutlane_Hard.Collidable = (LeftOutlanePost=2)

  if DangerZoneMod then
    DangerZoneGate.visible = 1
    DangerZoneGate.collidable = 1
    P_DangerZonePost.visible = 0
    DangerZonePin.isdropped = 1
  Else
    DangerZoneGate.visible = 0
    DangerZoneGate.collidable = 0
    P_DangerZonePost.visible = 1
    DangerZonePin.isdropped = 0
  End If

  'Table elements
  boxer.objRotZ=0
  leftArm.objRotZ=0
  rightArm.objRotZ=0
  BoxerSetPos
  MDirc=1

  LockWall1.isdropped = True
  LockWall2.isdropped = True
  DivPost1a.Isdropped = True
  DivPost2a.Isdropped = True

  RopePopperKicker.Isdropped = True

  PL.PullBack
  PR.PullBack

  DiverterClosed.IsDropped=0
  PostDiverter.IsDropped=1
  ForceField.IsDropped=1

  RightScoopBack.isdropped = True
  LeftScoopBack.isdropped = True
End Sub

Sub CP_Exit()
  Controller.Stop
End Sub


'---------------------------------------------------------------------------
' keyboard routines
'---------------------------------------------------------------------------
Const keyBuyInButton=3 ' (2)Specify Keycode for the Buy-In Button
ExtraKeyHelp=KeyName(keyBuyInButton) & vbTab & "Buy-in Button"

Sub CP_KeyDown(ByVal keycode)
    If keycode=PlungerKey Then Controller.Switch(23)=1
    If KeyDownHandler(keycode) Then Exit Sub
End Sub

Sub CP_KeyUp(ByVal keycode)
    If keycode=PlungerKey Then Controller.Switch(23)=0
    If KeyUpHandler(keycode) Then Exit Sub
End Sub


'-----------------------------------
' GI
'-----------------------------------
dim obj,GI0_status,GI1_status
Sub UpdateGI(GINo,Status)
  select case GINo
    case 0: if status <> GI0_status Then
          if status > 0 then
            DOF 101, DOFOn
            GI_PF1.state = Lightstateon
            for each obj in JumpRampBulbs
              obj.intensity = 60 / 8 * Status
              obj.state = lightstateon
            next

            for each obj in GIString1
              obj.intensity = Status + DimGI
              obj.state = lightstateon
            next
          else
            DOF 101, DOFOff
            GI_PF1.state = Lightstateoff
            for each obj in JumpRampBulbs
              obj.state = lightstateoff
            next
            for each obj in GIString1
              obj.state = lightstateoff
            next
          end if
          GI0_status = status
        End If
    case 1: if status <> GI1_status Then
          if status > 0 then
            GI_PF2.state = Lightstateon
            for each obj in GIString2
              obj.intensity = Status + DimGI
              obj.state = lightstateon
            next
          else
            GI_PF2.state = Lightstateoff
            for each obj in GIString2
              obj.state = lightstateoff
            next
          end if
          GI1_status = status
        End If
' Backglass

    case 2: if UseB2SBG then
          if status > 0 then
            'BackGlass On
            if Status > 4 then
              Controller.B2SSetData 1,1   'High Intensity
              Controller.B2SSetData 2,0   'Low Intensity
            Else
              Controller.B2SSetData 1,0
              Controller.B2SSetData 2,1
            end If
          else
            'BackGlass Off
            Controller.B2SSetData 1,0
            Controller.B2SSetData 2,0
          end if
        End If
' not used
'   case 3: if status then
'       end if
'   case 4: if status then
'       end if
  end select
End Sub

Sub UpdateLightsTimer_Timer
  If controller.Lamp(85) = LightstateOff Then
    if PostFlasher.state <> Lightstateoff then
      P_Post.image = "Post_red"
      PostFlasher.state = Lightstateoff
    end If
  Else
    if PostFlasher.state <> Lightstateon then
      P_Post.image = "Post_red_lit"
      PostFlasher.state = Lightstateon
    End If
  End If
End Sub


'---------------------------------------------------------------------------
'   R o p e
'---------------------------------------------------------------------------

Dim SPos:SPos = 0
Dim TurnCount,Rope1,Rope2

'RopeLevel 0
Rope1 = 620
Rope2 = 660
if RopeLevel = 1 Then
  Rope1 = 610
  Rope2 = 665
End If
if RopeLevel = 2 Then
  Rope1 = 600
  Rope2 = 670
End If

Sub UpdateRope(aNewPos,aSpeed,aLastPos)                 'animation for rope
  playsoundAtVol SoundFX("motor1",DOFGear), PRope, 1
  if aNewPos = 0 Then
    if Turncount < 20 then
      TurnCount = TurnCount + 1.5
    end If
  end if
  PRope.RotY = aNewPos/2 + 40

  If (aNewPos >= 700) or (aNewPos <= 20) Then         'zero position opto switch
    controller.switch(64) = 1
  else
    controller.switch(64) = 0
  end if

  If (aNewPos>Rope1) and (aNewPos<Rope2) Then         'ball hit at 6 'o clock rope time (620-660)
    if controller.Switch(45) Then
      playsoundAtVol "fx_collide", PRope, 1
      Ropepopper.kick int(rnd*70)+190,2+(TurnCount*0.2)
      Controller.Switch(45) = 0
    end If
  End If

  if controller.Switch(45) and ABS(BoxerZrot)<160 Then    'box fight is active - kick ball out
    if aNewPos > 600 and anewpos < 640 then
      playsoundAtVol "fx_collide", PRope, 1
      Ropepopper.kick 240,7
      Controller.Switch(45) = 0
    END If
  end if
End Sub

Sub Trigger6_Hit
  magRopeMagnet.AddBall ActiveBall
  magRopeMagnet.AttractBall ActiveBall
End Sub

Sub RopePopperKicker_Timer
  RopePopper.Timerenabled = False
  RopePopperKicker.isdropped = True
End Sub

'Magnet VUK animation
Dim VUKBall
Sub SolMagnetPopper(Enabled)
  If Enabled Then
    If Controller.Switch(45) = True Then
      PlaySoundAtVol SoundFX("Kicker",DOFContactors), RopePopper, 1
      Controller.Switch(45) = 0
      MagVUKTimer.enabled = 1
    end if
    RopePopperKicker.Isdropped = False
    RopePopperKicker.Timerenabled = True
  end if
end sub

Sub MagVUKTimer_Timer
  if VUKBall.z > 235 Then
    VUKBall.z = Vukball.z + 2
  Else
    VUKBall.z = Vukball.z + 3
  End If
  if VUKBall.z > 260 Then
    MagVUKTimer.enabled = 0
    RopePopper.kick 0,1
    TurnCount = 0
  end If
End Sub

'---------------------------------------------------------------------------
'   B o x e r
'---------------------------------------------------------------------------

Dim MDirc,boxerZRot
boxerZRot=0
Mdirc=-1

SolCallback(11)="SolRightArm"
SolCallback(13)="SolLeftArm"
SolCallback(26)="SolMotorDirc"  ' motor direction
SolCallback(27)="SolMotor"      ' boxer motor

Sub SolMotorDirc(Enabled)
   If Enabled Then
     MDirc=-1
   Else
     MDirc=1
   End If
End Sub

Sub SolMotor(Enabled)
  if Enabled then
    BoxerTurnTimer.Enabled=False
    BoxerTurnTimer.Interval=15 '20
    BoxerTurnTimer.Enabled=True
   else
    BoxerTurnTimer.Enabled=False
   end if
End Sub

Sub BoxerTurnTimer_Timer
  ' playsound SoundFX("motor1",DOFGear),0,0.15 ' TODO
  playsoundAtVol SoundFX("motor1",DOFGear), Boxer, 1
  boxerZRot = boxerZRot-MDirc
  If boxerZRot<-180 then boxerZRot=179
  If boxerZRot>180 then boxerZRot=-179
    BoxerSetPos
End Sub

dim rightArmRotx:rightArmRotx=0
dim rArmDir:rArmDir=1.5
dim leftArmRotx:leftArmRotx=0
dim lArmDir:lArmDir=1.5
dim cRad:cRad=3.14159265358979/180

Sub SolRightArm(Enabled)
  If Enabled Then
    playsound SoundFX("solon",DOFContactors),0,0.5,-0.25
    BRTimer.Enabled=1
  End If
End Sub

Sub SolLeftArm(Enabled)
  If Enabled Then
    playsound SoundFX("solon",DOFContactors),0,0.5,0.25
    BLTimer.Enabled=1
  End If
End Sub

' Right arm animation
Sub BRTimer_Timer
  rightArmRotx = rightArmRotx + rArmDir
  if rightArmRotx>75 then
    rightArmRotx=75
    rArmDir = rArmDir * -1
  end if
  if rightArmRotx<0 then
    BRTimer.Enabled=0
    rightArmRotx=0
    rArmDir = rArmDir * -1
  end if
  if rightArmRotx > 68 then
    rightArm.rotx=68
  Else
    rightArm.rotx=rightArmRotx
  End If
End Sub

' Left arm animation
Sub BLTimer_Timer
  leftArmRotx = leftArmRotx + lArmDir
  if leftArmRotx>75 then
    leftArmRotx=75
    lArmDir = lArmDir * -1
  end if
  if leftArmRotx<0 then
    leftArmRotx=0
    lArmDir = lArmDir * -1
    BLTimer.Enabled=0
  end if
  if leftArmRotx > 68 then
    leftArm.rotx=68
  Else
    leftArm.rotx=leftArmRotx
  End If
End Sub

sub BoxerSetPos
  if abs(boxerZRot)<=173 then    '<=173    Bag Center = 180
    Controller.Switch(46) = 1
    Boxer_Bag.isdropped = True
    Bag_Wall.collidable = False
    Bag_Wall.isdropped = True
    Boxer_Boxer.isdropped = False
    Boxer_BoxerHead.isdropped = False
    Boxer_Wall.collidable = True
  Else
    Controller.Switch(46) = 0
    Boxer_Bag.isdropped = False
    Bag_Wall.collidable = True
    Bag_Wall.isdropped = False
    Boxer_Boxer.isdropped = True
    Boxer_BoxerHead.isdropped = True
    Boxer_Wall.collidable = False
  end if

  if abs(boxerZRot)>7 then      '<3 Boxer Center = 0
    Controller.Switch(41)=1
  Else
    Controller.Switch(41)=0
  end if

  if (boxerZRot>=-25) or (boxerZRot<=-30) then      '-25, -28  Boxer Right
    Controller.Switch(47)=1
  Else
    Controller.Switch(47)=0
  end if

  if (boxerZRot<=25) or (boxerZRot>=30) then        '25, 28  Boxer Left
    Controller.Switch(48)=1
  Else
    Controller.Switch(48)=0
  end if

  boxer.objRotZ = boxerZRot
  leftArm.objRotZ=boxerZRot
  rightArm.objRotZ=boxerZRot
  P_BagScrew1.objRotZ=boxerZRot
  P_BagScrew2.objRotZ=boxerZRot

  leftArm.x = BoxerCenter.x - 30*Sin(boxerZRot*cRad)
  leftArm.y = BoxerCenter.y + 30*Cos(boxerZRot*cRad)
  rightArm.x = BoxerCenter.x - 30*Sin(boxerZRot*cRad)
  rightArm.y = BoxerCenter.y + 30*Cos(boxerZRot*cRad)
  Boxer.x = BoxerCenter.x - 30*Sin(boxerZRot*cRad)
  Boxer.y = BoxerCenter.y + 30*Cos(boxerZRot*cRad)

  P_BagScrew1.x = BoxerCenter.x + 75*Sin(boxerZRot*cRad)
  P_BagScrew1.y = BoxerCenter.y - 75*Cos(boxerZRot*cRad)

  P_BagScrew2.x = BoxerCenter.x + 157*Sin(boxerZRot*cRad)
  P_BagScrew2.y = BoxerCenter.y - 157*Cos(boxerZRot*cRad)

end sub

Sub Boxer_Wall_Hit
  playsound "pinhit_low"
  HitBoxer(activeball)
End Sub

Sub Bag_Wall_Hit
  playsound "pinhit_low"
  HitBoxer(activeball)
End Sub

Sub Boxer_Boxer_Hit
  playsound "rubber_hit_2"
  HitBoxer(activeball)
  RegisterBoxerHit(activeball)
End Sub

Sub Boxer_BoxerHead_Hit
  playsound "rubber_hit_2"
  HitBoxer(activeball)
  RegisterBoxerHit(activeball)
End Sub

Sub Boxer_Bag_Hit
  playsound "rubber_hit_2"
  HitBoxer(activeball)
  RegisterBagHit(activeball)
End Sub

Sub RegisterBoxerHit(BallObjPar)    'Body hits
  'SW66 - Boxer Gut 1
  'SW67 - Boxer Gut 2
  'SW68 - Boxer Head

  If BallObjPar.Z > 85 then    '75
    vpmtimer.PulseSw 68
    Exit Sub
  End If
  If BallObjPar.X < boxercenter.x then
    vpmtimer.PulseSw 66
  Else
    vpmtimer.PulseSw 67
  End If
End Sub

Sub RegisterBagHit(BallObjPar)    'Big Bag hits
  'SW12 - Boxer Bag
  vpmtimer.PulseSw 12
End Sub

Dim Orientation
Sub HitBoxer(BallObjPar)
  if abs(boxerZRot) > 90 then
    Orientation = 3
  Else
    Orientation = -3
  End If

  if BallObjPar.vely > 1 then
    Boxer.transy = Orientation
    rightarm.transy = Orientation
    leftarm.transy = Orientation
    if (BallObjPar.vely > 5) and (BallObjPar.vely <= 12) then
      Boxer.transy = Orientation * 2
      rightarm.transy = Orientation * 2
      leftarm.transy = Orientation *2
      If BallObjPar.X < (boxer.x - 20) then
        Boxer.transx = 1
        rightarm.transx = 1
        leftarm.transx = 1
      end if
      If BallObjPar.X > (boxer.x + 20) then
        Boxer.transx = -1
        rightarm.transx = -1
        leftarm.transx = -1
      End If
    end if
    if BallObjPar.vely > 12 then
      Boxer.transy = Orientation * 4
      rightarm.transy = Orientation * 4
      leftarm.transy = Orientation * 4
      If BallObjPar.X < (boxer.x - 20) then
        Boxer.transx = 3
        rightarm.transx = 3
        leftarm.transx = 3
      end if
      If BallObjPar.X > (boxer.x - 20) then
        Boxer.transx = -3
        rightarm.transx = -3
        leftarm.transx = -3
      End If
    end if
    BoxerHit.enabled = 1
  end if
End Sub

Sub BoxerHit_Timer
  BoxerHit.enabled = 0
  Boxer.transx = 0
  Boxer.transy = 0
  rightarm.transx = 0
  rightarm.transy = 0
  leftarm.transx = 0
  leftarm.transy = 0
End Sub


'---------------------------------------------------------------------------
' Speed Bag Fists
'---------------------------------------------------------------------------

Dim LEF,REF 'Direction Flag for speed bag fists
Const MaxTransY = -40
Const AnimStep = 2

Sub SpeedBag_Hit                      '65
  vpmTimer.PulseSw 65
  PSpeedBag.Transy = 8
  SpeedBagHit.enabled = 1
End Sub

Sub SpeedBagHit_Timer
  SpeedBagHit.enabled = 0
  PSpeedBag.Transy = 0
End Sub

Sub SolLeftSP(Enabled)
    If Enabled Then
        LEF=-AnimStep
        PL.Fire
    playsound SoundFX("solon",DOFContactors),0,1,0.8
    Else
        LEF=AnimStep
        PL.PullBack
    End If
    FistL.Enabled=1
End Sub

Sub SolRightSP(Enabled)
    If Enabled Then
    playsound SoundFX("solon",DOFContactors),0,1,0.82
        REF=-AnimStep
        PR.Fire
    Else
        REF=AnimStep
        PR.PullBack
    End If
    FistR.Enabled=1
End Sub

Sub FistL_Timer
  LeftPunchHand.transy = LeftPunchHand.transy + LEF
  LeftPunchPlunger.transy = LeftPunchHand.transy
  if LeftPunchHand.transy >= 0 Then
    FistL.Enabled=0
  End If
  if LeftPunchHand.transy <= MaxTransY Then
    FistL.Enabled=0
  End If
End Sub

Sub FistR_Timer
  RightPunchHand.transy = RightPunchHand.transy + REF
  RightPunchPlunger.transy = RightPunchHand.transy
  if RightPunchHand.transy >= 0 Then
    FistR.Enabled=0
  End If
  if RightPunchHand.transy <= MaxTransY Then
    FistR.Enabled=0
  End If
End Sub


'---------------------------------------------------------------------------
' Flipper Primitives
'---------------------------------------------------------------------------
sub FlipperMoveTimer_Timer()
  pleftFlipper.rotz=leftFlipper.CurrentAngle
  prightFlipper.rotz=rightFlipper.CurrentAngle
end sub

Sub SolLFlipper(Enabled)
    If Enabled Then
     PlaySoundAtVol SoundFX("flipperup1",DOFContactors), LeftFlipper, VolFlip
     LeftFlipper.RotateToEnd
    Else
     PlaySoundAtVol SoundFX("flipperdown",DOFContactors), LeftFlipper, VolFlip
     LeftFlipper.RotateToStart
  End If
End Sub

Sub SolRFlipper(Enabled)
  If Enabled Then
     PlaySoundAtVol SoundFX("flipperup1",DOFContactors), RightFlipper, VolFlip
     RightFlipper.RotateToEnd
    Else
     PlaySoundAtVol SoundFX("flipperdown",DOFContactors), RightFlipper, VolFlip
     RightFlipper.RotateToStart
    End If
End Sub

'Diverter Helper
Sub DiverterClosed_Hit
  If ActiveBall.VelX<0 Then
   DiverterClosed.IsDropped=1
   DiverterClosed.TimerEnabled=1
  End If
End Sub

Sub DiverterClosed_Timer
  DiverterClosed.IsDropped=0
  DiverterClosed.TimerEnabled=0
End Sub


'---------------------------------------------------------------------------
' Switches
'---------------------------------------------------------------------------
Sub CatapultKicker_Hit:bsCatapult.addball Me:End Sub
Sub SW15_Hit:Controller.Switch(15)=1:LockWall1.isdropped = False:End Sub
Sub SW15_Unhit:Controller.Switch(15)=0:LockWall1.isdropped = True:End Sub
Sub SW57_Hit:Controller.Switch(57)=1:LockWall2.isdropped = False:End Sub
Sub SW57_Unhit:Controller.Switch(57)=0:LockWall2.isdropped = True:End Sub
Sub SW58_Hit:Controller.Switch(58)=1:End Sub
Sub SW58_Unhit:Controller.Switch(58)=0:End Sub

Dim RStep,LStep
Sub LeftSlingshot_Slingshot
  vpmTimer.PulseSw 51
  PlaySoundAtVol SoundFX("left_Slingshot",DOFContactors), sling2, 1
    LSling.Visible = 0
    LSling1.Visible = 1
    sling2.TransZ = -20
    LStep = 0
    LeftSlingShot.TimerEnabled = 1
End Sub
Sub LeftSlingShot_Timer
    Select Case LStep
        Case 3:LSLing1.Visible = 0:LSLing2.Visible = 1:sling2.TransZ = -10
        Case 4:LSLing2.Visible = 0:LSLing.Visible = 1:sling2.TransZ = 0:LeftSlingShot.TimerEnabled = 0':gi3.State = 1:Gi4.State = 1
    End Select
    LStep = LStep + 1
End Sub
Sub RightSlingshot_Slingshot
  vpmTimer.PulseSw 52
  PlaySoundAtVol SoundFX("right_Slingshot",DOFContactors), sling1, 1
    RSling.Visible = 0
    RSling1.Visible = 1
    sling1.TransZ = -20
    RStep = 0
    RightSlingShot.TimerEnabled = 1
End Sub
Sub RightSlingShot_Timer
    Select Case RStep
        Case 3:RSLing1.Visible = 0:RSLing2.Visible = 1:sling1.TransZ = -10
        Case 4:RSLing2.Visible = 0:RSLing.Visible = 1:sling1.TransZ = 0:RightSlingShot.TimerEnabled = 0':gi1.State = 1:Gi2.State = 1
    End Select
    RStep = RStep + 1
End Sub

Sub Towel_Hit:Controller.Switch(63)=1:End Sub           '63
Sub Towel_Unhit:Controller.Switch(63)=0:End Sub
Sub LeftOutlane_Hit:Controller.Switch(16)=1:End Sub         '16
Sub LeftOutlane_unHit:Controller.Switch(16)=0:End Sub
Sub RightReturn_Hit:Controller.Switch(17)=1:End Sub         '17
Sub RightReturn_unHit:Controller.Switch(17)=0:End Sub
Sub LeftReturn_Hit:Controller.Switch(26)=1:End Sub          '26
Sub LeftReturn_unHit:Controller.Switch(26)=0:End Sub
Sub RightOutlane_Hit:Controller.Switch(27)=1:End Sub        '27
Sub RightOutlane_unHit:Controller.Switch(27)=0:End Sub

Sub BehindLeftScoop_Hit:Controller.Switch(42)=1:End Sub     '42
Sub BehindLeftScoop_Unhit:Controller.Switch(42)=0:End Sub
Sub BehindRightScoop_Hit:Controller.Switch(43)=1:End Sub    '43
Sub BehindRightScoop_Unhit:Controller.Switch(43)=0:End Sub
Sub EnterRamp_Hit:Controller.Switch(44)=1:End Sub           '44
Sub EnterRamp_Unhit:Controller.Switch(44)=0:End Sub

Sub EnterRope_Hit:Controller.Switch(78)=1:End Sub           '78
Sub EnterRope_unHit:Controller.Switch(78)=0:End Sub
Sub ExitRope_Hit:Controller.Switch(71)=1:P_ExitRopeSwitchArm.Rotx=0:End Sub            '71
Sub ExitRope_Unhit:Controller.Switch(71)=0:ExitRope.Timerenabled = True:End Sub

Sub ExitRope_Timer
  ExitRope.Timerenabled = False
  P_ExitRopeSwitchArm.Rotx = -20
End Sub

Sub EnterLockup_Hit:Controller.Switch(74)=1:End Sub         '74
Sub EnterLockup_unHit:Controller.Switch(74)=0:End Sub

Sub Drain_Hit:bsTrough.AddBall Me:PlaysoundAtVol "Drain5", drain, 1:End Sub                   '31-35

Sub RopePopper_Hit:Controller.Switch(45) = 1:set VUKball = activeball:End Sub
Sub CornerKickout_Hit:bsCornerKickout.AddBall 0:End Sub     '37

Sub ThreeBankMid_Hit:vpmTimer.PulseSw 25:End Sub            '25
Sub ThreeBankBottom_Hit:vpmTimer.PulseSw 53:End Sub         '53
Sub ThreeBankTop_Hit:vpmTimer.PulseSw 54:End Sub            '54


Sub LeftHalfGuy_Hit:vpmTimer.PulseSw 55:End Sub             '55
Sub RightHalfGuy_Hit:vpmTimer.PulseSw 56:End Sub            '56

Sub TopOfRamp_Hit:Controller.Switch(76)=1:End Sub           '76
Sub TopOfRamp_Unhit:Controller.Switch(76)=0:End Sub

Sub EnterSpeedBag_Hit:Controller.Switch(72)=1:End Sub       '72
Sub EnterSpeedBag_Unhit:Controller.Switch(72)=0:End Sub

Sub MadeRamp_Hit:Controller.Switch(11)=1:P_MadeRampArm.rotx = 0:End Sub            '11
Sub MadeRamp_Unhit:Controller.Switch(11)=0:MadeRamp.Timerenabled = True:End Sub

Sub MadeRamp_Timer
  MadeRamp.Timerenabled = False
  P_MadeRampArm.Rotx = -20
End Sub

Dim LScoopVel,RScoopVel,LKickforce,RKickforce
Sub EnterLeftScoop_Hit
  LScoopVel = BallVel(activeball)
End Sub

Sub EnterRightScoop_Hit
  RScoopVel = BallVel(activeball)
End Sub

Sub LeftScoopGrab_Hit
  if LScoopVel < 10 Then
    LeftScoopGrab.Kick 160,3
  Else
    playsoundAtVol "ScoopUp", LeftScoopGrab, 1
    Me.DestroyBall
    LeftScoopRelease.CreateBall
    LKickforce = Int(LScoopVel/6)
    if LKickforce < 4 then LKickforce = 3
    LeftScoopRelease.Kick 175,LKickforce
  end If
End Sub
Sub RightScoopGrab_Hit
  if RScoopVel < 10 Then
    RightScoopGrab.Kick 200,3
  Else
    playsoundAtVol "ScoopUp", RightScoopGrab, 1
    Me.DestroyBall
    RightScoopRelease.CreateBall
    RKickforce = Int(RScoopVel/8)
    if RKickforce < 3 then RKickforce = 3
    RightScoopRelease.Kick 205,RKickforce
  end If
End Sub

Sub LeftJabMade_Hit:Controller.Switch(36)=1:End Sub         '36
Sub LeftJabMade_Unhit:Controller.Switch(36)=0:End Sub

Sub RightJabMade_Hit:Controller.Switch(38)=1:End Sub        '38
Sub RightJabMade_Unhit:Controller.Switch(38)=0:End Sub

'Mod: gate instead of post
Sub DangerZoneGate_Hit:vpmTimer.PulseSw 73:End Sub          '73

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

' The sound is played using the VOL, PAN and PITCH functions, so the volume and pitch of the sound
' will change according to the ball speed, and the PAN function will change the stereo position according
' to the position of the ball on the table.


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
  PlaySound "gate4", 0, Vol(ActiveBall)*VolGates, AudioPan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
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
    PlaySound "fx_rubber2", 0, Vol(ActiveBall)*VolRo, AudioPan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
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
    Case 1 : PlaySound "flip_hit_1", 0, Vol(ActiveBall)*VolRH, AudioPan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
    Case 2 : PlaySound "flip_hit_2", 0, Vol(ActiveBall)*VolRH, AudioPan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
    Case 3 : PlaySound "flip_hit_3", 0, Vol(ActiveBall)*VolRH, AudioPan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
  End Select
End Sub

' CP specific
Sub LWireRampStart1_hit
  PlaySound "WireRamp1", 0, 1, AudioPan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)       'Vol(ActiveBall)
End Sub

Sub LWireRampStart2_hit
  PlaySound "WireRamp1", 0, 1, AudioPan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
End Sub

Sub RWireRampStart_hit
  PlaySound "WireRamp1", 0, 1, AudioPan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
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

Function AudioFade(tableobj) ' Fades between front and back of the table (for surround systems or 2x2 speakers, etc), depending on the Y position on the table. "CP" is the name of the table
  Dim tmp
  tmp = tableobj.y * 2 / CP.height-1
  If tmp > 0 Then
    AudioFade = Csng(tmp ^10)
  Else
    AudioFade = Csng(-((- tmp) ^10) )
  End If
End Function

Function AudioPan(tableobj) ' Calculates the pan for a tableobj based on the X position on the table. "CP" is the name of the table
  Dim tmp
  tmp = tableobj.x * 2 / CP.width-1
  If tmp > 0 Then
    AudioPan = Csng(tmp ^10)
  Else
    AudioPan = Csng(-((- tmp) ^10) )
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

Const tnob = 6 ' total number of balls
ReDim rolling(tnob)
InitRolling

Sub InitRolling
    Dim i
    For i = 0 to tnob
        rolling(i) = False
    Next
End Sub

Dim BOT, b
Sub RollingTimer_Timer()
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
          PlaySound("fx_ballrolling" & b), -1, Vol(BOT(b) ), AudioPan(BOT(b) ), 0, Pitch(BOT(b) ), 1, 0, AudioFade(BOT(b) )
        Else ' Ball on raised ramp
          PlaySound("fx_ballrolling" & b), -1, Vol(BOT(b) )*.5, AudioPan(BOT(b) ), 0, Pitch(BOT(b) )+50000, 1, 0, AudioFade(BOT(b) )
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
    PlaySound("fx_collide"), 0, Csng(velocity) ^2 / (VolDiv/VolCol), AudioPan(ball1), 0, Pitch(ball1), 0, 0, AudioFade(ball1)
End Sub

' Thalamus : Exit in a clean and proper way
Sub CP_exit()
  Controller.Pause = False
  Controller.Stop
End Sub


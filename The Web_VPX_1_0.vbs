'############################################################
'############################################################
'#######                                             ########
'#######          The Web                            ########
'#######          (Pro Pinball 1996)                 ########
'#######                                             ########
'############################################################
'############################################################
' Version VPX 1.0 mfuegemann 2017
'
' With permission by the Pro Pinball Team (Adrian Barritt)
'
' Thanks to:
' Zany for the 3D models and textures
' Arngrim for the DOF review
' Erik Mooney and Jonas Martinsson for their detailed rule sheet and strategy guide
' Pinballwiz45b for continued table testing and rule adjustments
' Gigalula for the DMD color setting code

Option Explicit

On Error Resume Next
ExecuteGlobal GetTextFile("controller.vbs")
If Err Then MsgBox Err
On Error Goto 0

'---------------------------
'-----  Configuration  -----
'---------------------------
'ResetHighscore           'enable to manually reset the Highscores

Const MusicActive = 1               'set to 1 to enable Music, place the music files into the music folder
Const DimFlashers = 0       'this value is added to the Flasher max. Alpha value, Minus is darker
Const DimGI = -50         'this value is added to the GI Flasher Alpha value, Minus is darker
Const BallsperGame = 3        'I added no Buy-In feature, so please use the number of Balls to reduce difficulty
Const ApplyMods = 1         '1: additional Mods are active (Flasher, UFO shaking,Apron Sphere lights), 0: original table features only
Const ChangeDMDColor = 0      'changes the DMD color if the default DMD color has been set before - set to 0 if You have problems writing to Your Registry
Const DampeningSpeed = 45     'ball speed at which the dampening kicks in
Const MaxTopSpeed = 17        'maximum velx at the top loop to make the ramp shot possible

Const cGameName = "theweb"

Sub TheWeb_Exit()
  If B2SOn = True Then Controller.Stop
  If UltraDMD.IsRendering Then
    UltraDMD.CancelRendering
  end if
    UltraDMD = NULL

  '------ UltraDMD Replace Default Color preference ------
  'code provided by Gigalula
  if ChangeDMDColor = 1 then
    If RegistryKeyExists and (DMDColorDefault <> DMDColorSelect) then
      DMDColor = WshShell.RegWrite ("HKCU\Software\UltraDMD\color",DMDColorDefault,"REG_SZ")
    End if
  end if
  '-------------------------------------------------------
End Sub

'--------------------------
'-----  The Web Init  -----
'--------------------------

'---------- UltraDMD Set Unique Table Color preference -------------
'code provided by Gigalula

Dim DMDColorDefault
Dim DMDColorSelect
Dim DMDColor
Dim WshShell

'Uncomment only the color of your choice

'DMDColorSelect = "Green"
'DMDColorSelect = "Yellow"
'DMDColorSelect = "Blue"
'DMDColorSelect = "White"
'DMDColorSelect = "Red"
DMDColorSelect = "OrangeRed"
'DMDColorSelect = "Orange"
'DMDColorSelect = "Cyan"
'DMDColorSelect = "Magenta"

' If you need more colors simply add them manually
' as long as they are supported by UltraDMD

if ChangeDMDColor = 1 then
  GetMyDMDColor
end if

Dim RegistryKeyExists
Sub GetMyDMDColor
  Set WshShell = CreateObject("WScript.Shell")
  on error resume next            ' turn off error trapping
  DMDColorDefault = WshShell.RegRead ("HKCU\Software\UltraDMD\color")
  RegistryKeyExists = (err.number = 0)
  on error goto 0                 ' restore error trapping

  If RegistryKeyExists and (DMDColorDefault <> DMDColorSelect) then
    DMDColor = WshShell.RegWrite ("HKCU\Software\UltraDMD\color",DMDColorSelect,"REG_SZ")
  End if
End Sub
'--------------------------------------------------------------------


Dim obj,Points,SkillShotCount,BallLaunch,Tilt,Highscore(5),HighName(5),x,DesktopMode
Sub TheWeb_init
  LoadEM
  Ballinplay = BallsperGame + 1
  SuperJets = 3
  LoadUltraDMD

  '--- In Desktop Mode, move some objects
  DesktopMode = TheWeb.ShowDT
  If DesktopMode = True Then 'Show Desktop components
    Ramp17.visible=1
    Ramp18.visible=1
    SideWood.visible=1
    P_UFO.rotx = 87
    Light3b.intensity = 0
  Else
    Ramp17.visible=0
    Ramp18.visible=0
    SideWood.visible=0
    P_UFO.rotx = 95
    Light3b.intensity = 0.1
  End if

  if ApplyMods = 0 then
    P_Sphere1.TransZ = -20
    P_Sphere2.TransZ = -20
    P_Sphere3.TransZ = -20
    P_Sphere4.TransZ = -20
    P_Sphere5.TransZ = -20
    P_Sphere6.TransZ = -20
    P_Sphere7.TransZ = -20
    P_Sphere8.TransZ = -20
  end if

  PlayNextMusic(IdleTheme)

  for each obj in GI: obj.State = 0:next

  SetB2SLamp 1,0

  x = LoadValue("TheWeb","High1")                     ' Init Highscores if empty
  If (x = "") then ResetHighscore
  x = LoadValue("TheWeb","High1")                     ' Loading Highscores
  If (x <> "") then Highscore(1) = CDbl(x) else Highscore(1) = 0
  Highname(1) = LoadValue("TheWeb","Name1")
  x = LoadValue("TheWeb","High2")
  If (x <> "") then Highscore(2) = CDbl(x) else Highscore(2) = 0
  Highname(2) = LoadValue("TheWeb","Name2")
  x = LoadValue("TheWeb","High3")
  If (x <> "") then Highscore(3) = CDbl(x) else Highscore(3) = 0
  Highname(3) = LoadValue("TheWeb","Name3")
  x = LoadValue("TheWeb","High4")
  If (x <> "") then Highscore(4) = CDbl(x) else Highscore(4) = 0
  Highname(4) = LoadValue("TheWeb","Name4")
  x = LoadValue("TheWeb","High5")
  If (x <> "") then Highscore(5) = CDbl(x) else Highscore(5) = 0
  Highname(5) = LoadValue("TheWeb","Name5")

  Tilt = 0
  randomize

  EjectToRight = int(2*rnd)

  Plunger.Pullback
  IdleTimer.enabled = True
  StartAttractmode
End sub


'------------------------
'-----  Main Timer  -----
'------------------------
Dim i,HideAnimation
Sub MainTimer_Timer
  if SpaceStationFrenzyActive or SuperLauncherActive or Missionactive or UltimateShowdownActive or QuickshotActive then
    HideAnimation = True
  else
    HideAnimation = False
  end if

  'Power Level Lights
  if (MultiBallBalls < 2) and not Quickshotactive and not IdleTimer.enabled then
    if LeftInlaneLight.state = 0 then
      LeftInlaneLight.state = 1
    end if
    if RightInlaneLight.state = 0 then
      RightInlaneLight.state = 1
    end if
  else
    if LeftInlaneLight.state <> 0 then
      LeftInlaneLight.state = 0
    end if
    if RightInlaneLight.state <> 0 then
      RightInlaneLight.state = 0
    end if
  end if

  'Top Rollover Bonus
  if BonusLightArray(1) = 1 and BonusLightArray(2) = 1 and BonusLightArray(3) = 1 then
    BonusLightArray(1) = 0
    BonusLightArray(2) = 0
    BonusLightArray(3) = 0
    Light1.state = lightstateoff
    Light2.state = lightstateoff
    Light3.state = lightstateoff
    AwardNextBonus
  else
    if BonusLightArray(1) = 1 then:light1.state = lightstateon:else:light1.state = lightstateoff:end if
    if BonusLightArray(2) = 1 then:light2.state = lightstateon:else:light2.state = lightstateoff:end if
    if BonusLightArray(3) = 1 then:light3.state = lightstateon:else:light3.state = lightstateoff:end if
  end if
  Light1b.state = Light1.state
  Light2b.state = Light2.state
  Light3b.state = Light3.state
  Light3c.state = Light3.state

  'BallSaver
  if GracePeriod.enabled then
    if BallSaverLight.state = 0 then
      BallSaverLight.state = Lightstateblinking
    end if
  else
    if GracePeriod1.enabled then
      if BallSaverLight.state = 0 then
        BallSaverLight.state = LightstateBlinking
      end if
    else
      if BallSaverLight.state <> 0 then
        BallSaverLight.state = LightstateOff
      end if
    end if
  end if

  'KickBack
  if EnableKickBackTimer.enabled then
    if KickbackTargetLight.state <> Lightstateblinking then
      KickbackTargetLight.state = Lightstateblinking
    end if
  else
    if KickbackTargetLight.state <> 0 then
      KickbackTargetLight.state = Lightstateoff
    end if
  end if

  if KickbackActive or SuperKickbackActive then
    if KickbackGracePeriodTimer.enabled or SuperKickbackActive then
      if KickbackLight.state <> Lightstateblinking then
        KickbackLight.state = Lightstateblinking
      end if
    else
      if KickbackLight.state <> Lightstateon then
        KickbackLight.state = Lightstateon
      end if
    end if
  else
    if KickbackLight.state <> 0 then
      KickbackLight.state = Lightstateoff
    end if
  end if

  'ExtraBalls
  ShootAgainLight.state = abs(Extraballs > 0)

  'Bumper deactivation
  if LeftFlipperUp and RightFlipperUp then
    BumperWall1.isdropped = False
    BumperWall2.isdropped = False
    BumperWall3.isdropped = False
  else
    BumperWall1.isdropped = True
    BumperWall2.isdropped = True
    BumperWall3.isdropped = True
  end if

  'Modes
  if not UltimateShowdownAcitve then
    if (QuickShotReady or SuperLauncherReady or ComboManiaReady or UltimateShowdownReady) and not MissionActive then
      if StartModeLight.state = 0 then
        StartModeLight.state = Lightstateblinking
      end if
    else
      if StartModeLight.state <> 0 then
        StartModeLight.state = 0
      end if
    end if

    if VideoModeReady and (MultiBallBalls < 2) and not MissionActive and not QuickShotActive and not SuperlauncherActive and not ComboManiaActive then
      if VideoModeLight.state = 0 then
        VideoModeLight.state = 1
      end if
    else
      if VideoModeLight.state = 1 then
        VideoModeLight.state = 0
      end if
    end if
  end if

  'Ultimate Showdown
  if ((US_MissionsLight.state <> 0) and (US_PowerLight.state <> 0) and (US_UltraLight.state <> 0) and (US_BonusLight.state <> 0) and (US_CombosLight.state <> 0)) or UltimateShowdownReady then
    if USLight.state = 0 then
      USLight.state = Lightstateblinking
    end if
    UltimateShowdownReady = True
  else
    if USLight.state <> 0 then
      USLight.state = 0
    end if
    UltimateShowdownReady = False
  end if

  if UltimateShowdownActive then
    if US_MissionsLight.state <> Lightstateblinking then
      US_MissionsLight.state = Lightstateblinking
    end if
    if US_PowerLight.state <> Lightstateblinking then
      US_PowerLight.state = Lightstateblinking
    end if
    if US_UltraLight.state <> Lightstateblinking then
      US_UltraLight.state = Lightstateblinking
    end if
    if US_BonusLight.state <> Lightstateblinking then
      US_BonusLight.state = Lightstateblinking
    end if
    if US_CombosLight.state <> Lightstateblinking then
      US_CombosLight.state = Lightstateblinking
    end if
    if USLight.state <> Lightstateblinking then
      USLight.state = Lightstateblinking
    end if

    if (MultiBallBalls < 6) and (ReserveBalls > 0) then
      MysteryLight.state = Lightstateblinking
    else
      MysteryLight.state = 0
    end if
  end if

  'Shere Lights
  SetSphereLamp P_Sphere1, 1
  SetSphereLamp P_Sphere2, 2
  SetSphereLamp P_Sphere3, 3
  SetSphereLamp P_Sphere4, 4
  SetSphereLamp P_Sphere5, 5
  SetSphereLamp P_Sphere6, 6
  SetSphereLamp P_Sphere7, 7
  SetSphereLamp P_Sphere8, 8

end sub

Sub SetSphereLamp(ObjPar,SCountPar)
  if not UltimateShowdownAcitve then
    if ShowdownSpheres >= SCountPar then
      ObjPar.image = "Bulb_Red_on"
    else
      ObjPar.image = "Bulb_Red_off"
    end if
  else
    if ReserveBalls >= SCountPar then
      ObjPar.image = "Bulb_Red_on"
    else
      ObjPar.image = "Bulb_Red_off"
    end if
  end if
End Sub

Dim GameActive,BallInPlay,GracePeriodActive,SkillShot,SuperSkillShot,KickbackActive,SuperKickbackActive,ExtraBalls,Spinneractive,JackpotValue,LeftRampCount,FirstBall
Dim LoopCount,SuperJetsHeld,BonusHeld,QuickShotReady,SuperLauncherReady,ComboManiaReady,UltimateShowdownReady,VideoModeReady,ShowdownSpheres,FrenzyCount
Dim BallsOnPF,EnableIdleSound
Sub StartGame
  if GameActive = 0 then
    StopAttractMode
    DMDDisplayScene "Main",True,backgrnd, Text1, 8,14,Text2, 5,15, 14, 10000000, 14
    BallsSaved = 0
    IdleTimer.enabled = False
    GameActive = 1
    Points = 0
    Tilt = 0
    BallsOnPF = 0
    AutoPlungerActive = False
    MissionActive = False
    MissionCount = 1
    QuickShotReady = False
    SuperLauncherReady = False
    ComboManiaReady = False
    UltimateShowdownReady = False
    UltimateShowdownActive = False
    QuickShotBaseValue = 0      'will be initialized at first hit
    SuperLauncherBaseValue = 25000000
    SuperLauncherAddValue = 0
    FastFrenzyCount = 0
    ResetBonus
    ResetPowerLevel
    QuickDockLight.State = 1
    ExplosionLight.state = 0
    LeftRampLight.state = 0
    RightRampLight.state = 0
    ResetMissions
    BallInPlay = 1
    SuperJets = 0
    SkillShot = 1
    SkillShotCount = 0
    ExtraBalls = 0
    BonusExtraBallAwarded = False
    LeftRampCount = 0
    ShowdownSpheres = 0
    BallsDocked = 0
    BallsLocked = 0
    US_MissionsLight.state = 0
    US_PowerLight.state = 0
    US_UltraLight.state = 0
    US_BonusLight.state = 0
    US_CombosLight.state = 0
    USLight.state = 0
    USCount = 0
    DT1Light.state = 0
    DT2Light.state = 0
    DT3Light.state = 0
    DT4Light.state = 0
    DT5Light.state = 0
    DT6Light.state = 0
    JackpotValue = 10000000
    ComboManiaValue = 20000000
    TotalLoopCount = 0
    TotalPowerlevelCount = 0
    TotalComboCount = 0
    ActivateMagnetLight.state = 1
    LockLight.state = 0
    DockLight.state = 0
    DockBallLock.enabled = False
    DockAdditionalSpaceStationBall = False
    MultiBallBalls = 0
    LoopExtraBalls = 0

    VMTimerIntervall = VMTimerStartIntervall
    VideoModeLock.isdropped = True
    VideoModeActive = False
    VideoModeReady = False
    VMCount = 0
    PerfectVMCount = 0

    for each obj in GI: obj.state = 1:next
    SetB2SLamp 1,1

    FirstBall = True
    EnableIdleSound = True
    NextBall

    if ApplyMods = 0 then
      P_LeftFlasher.image = "dome4_red_lit"
      FLeft.State = 0
      DOF 141, DOFPulse
      P_RightFlasher.image = "dome4_red_lit"
      FRight.State = 0
      DOF 142, DOFPulse
    end if
  end if
End Sub

Sub NextBall
  PlayNextMusic(MainTheme)
  GracePeriodActive = False
  if SuperJetsHeld then
    SuperJetsHeld = False
  else
    SuperJets = 0
  end if
  SuperSkillShot = 0
  KickBackActive = True
  SuperKickbackActive = False
  Spinneractive = True
  ReactorCriticalActive = False
  BumperLane.Timerenabled = False
  BumperLaneLight.state = 0
  LockLightstate = 0
  ResetPowerLevel
  StartMissionLight.State = Lightstateblinking
  if BallsLocked > 0 then
    BallsLocked = 0
    LockLight.state = 0
    FastFrenzyBalls = 0
  end if

  DTDropUp 0
  if DT1Light.state = 0 or DT2Light.state = 0 or DT3Light.state = 0 then
    DT1Light.state = 0
    DT2Light.state = 0
    DT3Light.state = 0
  end if
  if DT4Light.state = 0 or DT5Light.state = 0 or DT6Light.state = 0 then
    DT4Light.state = 0
    DT5Light.state = 0
    DT6Light.state = 0
  end if
  BallReleaseTimer.enabled = 1
End Sub

Dim MultiBallBalls,BallsSaved
Sub Drain_Hit()
  playsound "W_drain"
  DOF 137, DOFPulse
  FlashB2SLamp 51,2
  FlashB2SLamp 50,1

  Drain.DestroyBall
  BallsOnPF = BallsOnPF - 1
  playsound "drain2"

  if GracePeriod2.enabled then
    BallsSaved = BallsSaved + 1
    TroughHandler
  else
    if (MultiBallBalls + BallsSaved) > 1 then
      MultiBallBalls = MultiBallBalls - 1
      if (MultiBallBalls + BallsSaved) < 2 then
        if FastFrenzyActive then
          EndFastFrenzy
        end if
        if SpaceStationFrenzyActive then
          EndSpaceStationFrenzy
        end if
        if SecretManiaActive then
          EndSecretMania
        end if
        if UltimateShowdownActive then
          EndUltimateShowdown
        end if
      end if
    end if
    if BallsOnPF = 0 then
      PlayDrainAnimation
      for each obj in GI: obj.state = 0:next

      SetB2SLamp 1,0
      TroughHandler
    end if
  end if
End Sub

Dim AutoPlungerActive,ShootAgain
Sub TroughHandler
  if MultiBallBalls > 1 then
    if BallsSaved > 0 then
      BallReleaseTimer.enabled = True
    else
      'just destroy one of the Multiballs
    end if
  else
    if BallsSaved > 0 then
      BallReleaseTimer.enabled = True
      if not hideanimation then
        DMDDisplayScene "BS",True,BallSavedback, "", 8,14,"", 5,15, 14, 3000, 1
      end if
    else
      DeactivateMissionsAndModes
      SkillShot = 1
      if ExtraBalls <= 0 then
        BallInPlay = BallInPlay + 1
      else
        ExtraBalls = ExtraBalls - 1
        ShootAgain = True
      end if
      AddBonus
    end if
  end if
End Sub

Sub EndOfGame
  GameActive = 0
  BallInPlay = 0
  EnterInitials
  idleTimer.enabled = True
  if ApplyMods = 0 then
    P_LeftFlasher.image = "dome4_red"
    FLeft.State = 0
    P_RightFlasher.image = "dome4_red"
    FRight.State = 0
  end if
End Sub

Sub BallReleaseTimer_Timer
  Tilt = 0

  if (BallsSaved > 0) then
    BallsSaved = BallsSaved - 1
    AutoPlungerActive = True
    GracePeriodActive = True
    if (MultiBallBalls < 2) then
      if int(2*rnd) = 1 then
        Playsound "W_BallSaved2",0,1,0,0
      else
        Playsound "W_DontMove",0,1,0,0
      end if
    end if
  end if
  if BallsSaved = 0 then
    BallReleaseTimer.enabled = False
  end if
  BallRelease.CreateBall
  BallRelease.kick 90,5
  BallsOnPF = BallsOnPF + 1
  if Shootagain then
    DMDDisplayScene "Shoot",True,ShootAgainback, "", 8,14,"", 5,15, 14, 3000, 1
    if int(2*rnd) = 1 then
      Playsound "W_ShootAgain",0,1,0,0
    else
      Playsound "W_ShootAgain2",0,1,0,0
    end if
    ShootAgain = False
  else
    UpdateDMDScore
  end if
  for each obj in GI: obj.state = 1:next

  SetB2SLamp 1,1
  DOF 131, DOFPulse

  if FirstBall then
    FirstBall = False
    PlaySound "W_Start1",0,1,0,0
  else
    PlaySound SoundFX("ballrel",0),0,0.7,0.25,0
  end if

  if AutoPlungerActive then
    AutoPlungerActive = False
    AutoPlungerTimer.enabled = True
  end if
End Sub

Sub AutoPlungerTimer_Timer
  AutoPlungerTimer.enabled = False
  if AutoplungerTrigger.Ballcntover <> 0 then
    Plunger.fire
    Plunger.Timerenabled = True
    DOF 132, DOFPulse
  end if
End Sub

Sub BallLaunched_Hit
  EnableIdleSound = False
  if ActiveBall.vely < -6 then
    Playsound "launch",0,0.75,0.5,0.25
  end if
  if not GracePeriodActive then
    BallSaverLight.BlinkInterval = 160
    StartBallSaver
  end if
End Sub

Dim BonusExtraBallAwarded
Sub AwardNextBonus
  Playsound "W_TopRolloverComplete"
  FlashB2SLamp 48,1
  FlashB2SLamp 49,1
  FlashB2SLamp 51,1
  select case BonusMultiplier
    case 1: BonusMultiplier = 2
        if not Hideanimation then
          DMDDisplayScene "BX2",True,B2back, FormatDMDTopText("","",""), 8,14,FormatDMDBottomText(""), 5,15, 14, 1500, 1
        end if
    case 2: BonusMultiplier = 4
        if not Hideanimation then
          DMDDisplayScene "BX4",True,B4back, FormatDMDTopText("","",""), 8,14,FormatDMDBottomText(""), 5,15, 14, 1500, 1
        end if
    case 4: BonusMultiplier = 6
        if not Hideanimation then
          DMDDisplayScene "BX6",True,B6back, FormatDMDTopText("","",""), 8,14,FormatDMDBottomText(""), 5,15, 14, 1500, 1
        end if
    case 6: BonusMultiplier = 8
        if not Hideanimation then
          DMDDisplayScene "BX8",True,B8back, FormatDMDTopText("","",""), 8,14,FormatDMDBottomText(""), 5,15, 14, 1500, 1
        end if
    case 8: BonusMultiplier = 10
        if not BonusExtraBallAwarded then
          ExtraBalls = ExtraBalls + 1
          BonusExtraBallAwarded = True
          Playsound "W_ExtraBall"
          DOF 123, DOFPulse
          if not hideanimation then
            DMDDisplayScene "EB",True,ExtraBallback, "", 8,14,"", 5,15, 14, 3900, 1
          end if
        else
          if not Hideanimation then
            DMDDisplayScene "BX10",True,B10back, FormatDMDTopText("","",""), 8,14,FormatDMDBottomText(""), 5,15, 14, 1500, 1
          end if
        end if
  case else
    US_BonusLight.state = 1
    MaxBonuspoints = MaxBonuspoints + 5000000
    Addpoints(MaxBonuspoints)
    if not Hideanimation then
      DMDDisplayScene "BMax",True,BMaxback, FormatDMDTopText("","",""), 8,14,FormatDMDBottomText(FormatPointText(cstr(MaxBonuspoints))), 5,15, 14, 1500, 1
    end if
  end select
End Sub

Dim MaxBonusPoints,BonusMultiplier
Sub ResetBonus
  BonusLightArray(1) = 0
  BonusLightArray(2) = 0
  BonusLightArray(3) = 0
  MaxBonuspoints = 5000000
  if not BonusHeld then
    BonusMultiplier = 1
  else
    BonusHeld = False
  end if
End Sub

'--- Ball Saver ---
Sub GracePeriod_Timer()
  GracePeriod.enabled = False
  BallSaverLight.BlinkInterval = 80
End Sub

Sub GracePeriod1_Timer()
  GracePeriod1.enabled = False
End Sub

Sub GracePeriod2_Timer()
  GracePeriod2.enabled = False
  GracePeriod.interval  =  9000
  GracePeriod1.interval = 12000
  GracePeriod2.interval = 14000
End Sub

Sub StartBallSaver
  GracePeriod.enabled = False
  GracePeriod.enabled = True
  GracePeriod1.enabled = False
  GracePeriod1.enabled = True
  GracePeriod2.enabled = False
  GracePeriod2.enabled = True
End Sub

'--- Top Rollovers Bonus Multiplier ---
Dim Temp,BonusLightArray(3)

Sub ChangeBonusLight(direction)
  Temp = BonusLightArray(1)
  if direction = -1 then
    BonusLightArray(1) = BonusLightArray(2)
    BonusLightArray(2) = BonusLightArray(3)
    BonusLightArray(3) = Temp
  else
    BonusLightArray(1) = BonusLightArray(3)
    BonusLightArray(3) = BonusLightArray(2)
    BonusLightArray(2) = Temp
  end if
end sub


'--- Skill Shot ---
Sub SkillShotTrigger_Hit
  if SkillShot = 1 then
    SkillShotTimer.enabled = True
    DockRampSpider.state = Lightstateblinking
  end if
  PlaySound "W_BallDrop",0,1,0.25,0
end sub

Sub SkillShotTimer_Timer()
  SkillShotTimer.enabled = False
  DockRampSpider.state = Lightstateoff
  SkillShot = 0
End Sub

Sub SkillShotComboTimer_Timer()
  SkillShotComboTimer.enabled = False
End Sub

Dim DiverterOpen
Sub TopRampMade_Hit
  JackpotValue = JackpotValue + 1000000
  DOF 138, DOFPulse
  if not HideAnimation then
    UpdateDMDText "Main","","JACKPOT  VALUE","",FormatPointText(cstr(JackpotValue))
  end if

  FlashRight 1
  if BikeRaceActive then
    if DockRampSpider.state <> 0 then
      AdvanceBikeRaceScore(False)
    end if
  end if

  if ShuttleActive then
    FlashRight 3
    Advanceshuttlescore
  end if

  if SecretManiaActive then
    AwardSecretManiaScore
  end if

  if SuperLauncherActive then
    TopRampDiverter.rotatetostart
    AdvanceSuperLauncher
  else
    if SpaceStationFrenzyActive then
      if UltraJackpotLight.state <> 0 then
        AwardSpaceStationUltraJackpot
      else
        if SuperJackpotLight.state <> 0 then
          AwardSpaceStationSuperJackpot
        end if
      end if
    else
      DiverterOpen = Not DiverterOpen

      if Quickshotactive then
        DiverterOpen = EjectToRight
      end if

      if DiverterOpen then
        TopRampDiverter.rotatetoend
      else
        TopRampDiverter.rotatetostart
      end if

      if Skillshot = 1 then
        SkillShotTimer.enabled = False
        DockRampSpider.state = Lightstateoff

        SkillShotCount = SkillShotCount + 1
        if SuperSkillShot = 0 then
          PlayNextsound "W_Skillshot"
          addpoints(SkillShotCount*25000000)
          UpdateDMDText "Main","","SKILL SHOT","",FormatPointText(cstr(SkillShotCount*25000000))
        else
          PlayNextsound "W_Skillshot2"
          addpoints((SkillShotCount + SuperSkillShot)*25000000)
          UpdateDMDText "Main","","SUPER SKILL SHOT","",FormatPointText(cstr((SkillShotCount + SuperSkillShot)*25000000))
        end if
        SkillShot = 0
        SuperSkillShot = 0
        SkillShotComboTimer.enabled = True
      else
        if combotimer.enabled and (DockRampSpider.UserValue = 1) then
          AwardNextCombo
          ComboTimer.enabled = False
          ComboTimer.enabled = True
        else
          if Skillshot = 0 then
            actualcombocount = 0
            ComboTimer.enabled = True
            if CombomaniaActive then
              AwardnextCombo
            end if
          end if
        end if
      end if
      if (StartMissionLight.State = 0) and not MissionActive and not UltimateShowdownActive then
        StartMissionLight.State = Lightstateblinking
      end if
    end if
  end if
end Sub

'--- Fast Frenzy ---
Dim BallsLocked,FastFrenzyBalls,FastFrenzyCount,DTCount,FFJackpot,SkipKick
SkipKick = False
Sub LeftKicker_Hit()
  LeftKicker.DestroyBall
  Playsound "Drain3",0,0.3,-0.18,0.25
  if not FastFrenzyActive and not SpaceStationFrenzyActive then
    if HuntDownActive and (LeftHolespider.state <> 0) then
      AdvanceHuntDownScore
    end if

    if UltimateShowdownActive then
      AwardUSPoints
    else
      if SecretManiaActive then
        AwardSecretManiaScore
      else
        if (LockLight.state = 0) and (VideoModeLight.state = 0) then
          if LeftRampLight.state <> 1 then
            LeftRampLight.state = 1
          else
            RightRampLight.state = 1
          end if
        else
          if VideoModeLight.state <> 0 then
            VideoModeLight.state = 0
            DeactivateMissionsAndModes
            SkipKick = True
            StartVideoMode False
          else
            if LockLight.state <> 0 then
              if BallsLocked = 0 then
                if not HideAnimation then     'Activate Pick only if no Mode or Mission is active
                  TakeYourPick
                  SkipKick = True
                end if
              else
                BallsLocked = BallsLocked + 1
                if BallsLocked = FastFrenzyBalls then
                  if not ComboManiaActive then
                    DeactivateMissionsAndModes
                    StartFastFrenzy
                    SkipKick = True
                  else
                    DeactivateMissionsAndModes
                    StartSecretMania
                    SkipKick = True
                  end if
                else
                  select case BallsLocked
                    case 1: PlayNextsound "W_Ball1Locked"
                    case 2: PlayNextsound "W_Ball2Locked"
                    case 3: PlayNextsound "W_Ball3Locked"
                  end select
                  if BallsLocked = FastFrenzyBalls - 1 then
                    LockLight.state = Lightstateblinking
                  end if
                end if
              end if
            end if
          end if
        end if
      end if
    end if
  else
    if CowJackpotActive then
      AwardCowJackpot
    end if
  end if
  if not SkipKick then
    TroughLeftHoleKicker.CreateBall
    TroughLeftHoleKicker.kick 180, 1
  else
    SkipKick = False
  end if
End Sub


'pick
Dim PickTime,PickChoice
Sub TakeYourPick
  playNextsound "W_TakeYourPick"
  PickTime = 8
  PickChoice = 1
  FrenzyPickTimer.enabled = True
  UpdateDMDText "Main","","Take Your Pick","","(2) 3  4  ? :" & cstr(PickTime)
End Sub

Sub ChangePick(PickChangeDir)
  PickChoice = Pickchoice + PickchangeDir
  if Pickchoice > 4 then
    Pickchoice = 1
  end if
  if Pickchoice < 1 then
    Pickchoice = 4
  end if
  select case PickChoice
    case 1: UpdateDMDText "Main","","Take Your Pick","","(2) 3  4  ? :" & cstr(PickTime)
    case 2: UpdateDMDText "Main","","Take Your Pick",""," 2 (3) 4  ? :" & cstr(PickTime)
    case 3: UpdateDMDText "Main","","Take Your Pick",""," 2  3 (4) ? :" & cstr(PickTime)
    case 4: UpdateDMDText "Main","","Take Your Pick",""," 2  3  4 (?):" & cstr(PickTime)
  end select
End Sub

Sub FrenzyPickTimer_Timer
  PickTime = PickTime - 1
  select case PickChoice
    case 1: UpdateDMDText "Main","","Take Your Pick","","(2) 3  4  ? :" & cstr(PickTime)
    case 2: UpdateDMDText "Main","","Take Your Pick",""," 2 (3) 4  ? :" & cstr(PickTime)
    case 3: UpdateDMDText "Main","","Take Your Pick",""," 2  3 (4) ? :" & cstr(PickTime)
    case 4: UpdateDMDText "Main","","Take Your Pick",""," 2  3  4 (?):" & cstr(PickTime)
  end select
  if Picktime <= 0 then
    FastFrenzyPickMade
  end if
End Sub

Dim DockAdditionalSpaceStationBall,ExplosionFF
Sub FastFrenzyPickMade
  FrenzyPickTimer.enabled = False
  select case PickChoice
    case 1: FastFrenzyBalls = 2
        BallsLocked = 1
        PlayNextsound "W_Ball1Locked"
        UpdateDMDText "Main","",Text1,"","Ball 1  Locked"
    case 2: FastFrenzyBalls = 3
        BallsLocked = 1
        PlayNextsound "W_Ball1Locked"
        UpdateDMDText "Main","",Text1,"","Ball 1  Locked"
    case 3: FastFrenzyBalls = 4
        BallsLocked = 1
        PlayNextsound "W_Ball1Locked"
        UpdateDMDText "Main","",Text1,"","Ball 1  Locked"
    case 4: if DockAdditionalSpaceStationBall = False then
          DockAdditionalSpaceStationBall = True
          PlayNextsound "W_Ball1Locked"
          UpdateDMDText "Main","",Text1,"","MYSTERY"
          LockLight.state = 0
        else
          FastFrenzyBalls = 5
          BallsLocked = 1
          PlayNextsound "W_Ball1Locked"
          UpdateDMDText "Main","",Text1,"","MYSTERY"
        end if
  end select
  TroughLeftHoleKicker.CreateBall
  TroughLeftHoleKicker.kick 180, 1
End Sub

Dim DockState,DockLightState
Sub StartFastFrenzy
  PlayNextSound "W_FastFrenzy"
  DockState = DockBallLock.enabled
  DockLightState = DockLight.state
  DockBallLock.enabled = False
  PlayNextmusic(FastFrenzyTheme)
  FastFrenzyActive = True
  LockLight.state = 0
  ExplosionLightState = ExplosionLight.State
  ExplosionLight.State = 0
  FastFrenzyCount = FastFrenzyCount + 1
  DTDropUp 0
  DT1Light.state = Lightstateblinking
  DT2Light.state = Lightstateblinking
  DT3Light.state = Lightstateblinking
  DT4Light.state = Lightstateblinking
  DT5Light.state = Lightstateblinking
  DT6Light.state = Lightstateblinking

  P_Jackpotlight.image = "Bulb_Yellow_off"
  JackpotLight.state = 0

  MultiBallBalls = FastFrenzyBalls
  BallsSaved = 1

  DMDDisplayScene "FF",True,FFback, "", 8,14,"", 5,15, 14, 2000, 1

  if ExplosionFF then
    ExplosionFF = False
  else
    TroughLeftHoleKicker.CreateBall
    TroughLeftHoleKicker.kick 180, 1
  end if
  select case FastFrenzyBalls
    case 2: 'normal Left VUK
    case 3: RightRampVUK.CreateBall
        RightRampVUK.kick 195, 1
        BallsonPF = BallsonPF + 1
        DOF 135, DOFPulse
        DOF 136, DOFPulse
    case 4: RightRampVUK.CreateBall
        RightRampVUK.kick 195, 1
        BallsonPF = BallsonPF + 1
        BallsSaved = 2
        DOF 135, DOFPulse
        DOF 136, DOFPulse
    case 5: RightRampVUK.CreateBall
        RightRampVUK.kick 195, 1
        BallsonPF = BallsonPF + 1
        BallsSaved = 3
        DOF 135, DOFPulse
        DOF 136, DOFPulse
    case 6: BallsSaved = 5
        'BallsonPF = BallsonPF + 1
  end select
  GracePeriodActive = True
  BallSaverLight.BlinkInterval = 160
  StartBallSaver
  BallreleaseTimer.enabled = True
  AutoPlungerActive = True
End Sub

Sub CheckFastFrenzy
  if FastFrenzyActive then
    JackpotLight.state = Lightstateblinking
  end if
End Sub

Sub AwardFFJackpot
  FlashB2SLamp 51,2
  FlashB2SLamp 50,1
  DTCount = 0
  select case int(4*rnd)
    case 0: playnextsound "W_IamImpressed"
    case 1: playnextsound "W_Outstanding"
    case 2: playnextsound "W_VeryGood"
    case 3: playnextsound "W_Lucky"
  end select
  if DT1.isdropped then
    DTCount = DTCount + 1
  end if
  if DT2.isdropped then
    DTCount = DTCount + 1
  end if
  if DT3.isdropped then
    DTCount = DTCount + 1
  end if
  if DT4.isdropped then
    DTCount = DTCount + 1
  end if
  if DT5.isdropped then
    DTCount = DTCount + 1
  end if
  if DT6.isdropped then
    DTCount = DTCount + 1
  end if
  if DTCount = 6 then
    DTCount = 10
  end if
  FFJackpot = 1000000 * DTCount * FastFrenzyBalls * (BallsonPF + FastFrenzyCount)
  if FFJackpot < 6000000 then
    FFJackpot = 6000000
  end if
  Addpoints(FFJackpot)
  if DTCount = 10 then
    ShowdownSpheres = ShowdownSpheres + 1
    PlayNextsound "W_Sphere"
    DOF 139, DOFPulse
    DMDDisplayScene "Sphere",True,Sphereback, "", 8,14,FormatPointText(cstr(FFJackpot)), 5,15, 14, 3000, 1
  else
    if not Hideanimation then
      select case int(3*rnd)
        case 0: DMDDisplayScene "JPa",True,JPaback, "", 8,14,FormatPointText(cstr(FFJackpot)), 5,15, 14, 1800, 1
        case 1: DMDDisplayScene "JPb",True,JPbback, "", 8,14,FormatPointText(cstr(FFJackpot)), 5,15, 14, 2100, 1
        case 2: DMDDisplayScene "JPc",True,JPcback, "", 8,14,FormatPointText(cstr(FFJackpot)), 5,15, 14, 1800, 1
      end select
    end if
  end if
  DTDropUp 0
  DT1Light.state = Lightstateblinking
  DT2Light.state = Lightstateblinking
  DT3Light.state = Lightstateblinking
  DT4Light.state = Lightstateblinking
  DT5Light.state = Lightstateblinking
  DT6Light.state = Lightstateblinking
  JackpotLight.state = 0
End Sub

Sub EndFastFrenzy
  FastFrenzyActive = False
  playnextmusic(MainTheme)
  MultiBallBalls = 0
  BallsLocked = 0
  if DT1.isdropped then
    DT1Light.state = 1
  else
    DT1Light.state = 0
  end if
  if DT2.isdropped then
    DT2Light.state = 1
  else
    DT2Light.state = 0
  end if
  if DT3.isdropped then
    DT3Light.state = 1
  else
    DT3Light.state = 0
  end if
  if DT4.isdropped then
    DT4Light.state = 1
  else
    DT4Light.state = 0
  end if
  if DT5.isdropped then
    DT5Light.state = 1
  else
    DT5Light.state = 0
  end if
  if DT6.isdropped then
    DT6Light.state = 1
  else
    DT6Light.state = 0
  end if
  JackpotLight.state = 0
  LockLight.state = 0
  if DockState then
    DockBallLock.enabled = True
  end if
  DockLightState = DockLight.state
  ExplosionLight.State = ExplosionLightState
End Sub


'--- Video Mode ---

Dim PrevRow,VMScore,VMLifes,VMText,ItemDestroyed,VMTimerIntervall,VMCount,PerfectVMCount,ErrorMade,ActiveVMImage,VMLoopCount,ItemCount,VideoModeActive,WaitOneFrame,FireExplosionVUK
Dim Row(100),VM(100)
Const WaveItems = 78          '77 items, start at 1
Const VMTimerStartIntervall = 80    'Video Mode initial speed
Const WaveChange = 3          'speed step for each of the 6 waves within one Video Mode
Const VMChange = 8            'speed step for the next Video Mode

Sub StartVideoMode(ExplosionVUKPar)
  SetB2SLamp 1,0

  FireExplosionVUK = ExplosionVUKPar
  UpdateDMDText "Main","","SHOOT MINES ONLY","","VIDEO MODE"
  playNextsound "W_VideoMode"

  VMCount = VMCount + 1
  ErrorMade = False
  VideoModeLock.isdropped = False
  VideoModeActive = True
  LeftFlipper.rotatetostart
  ULeftFlipper.rotatetostart
  RightFlipper.rotatetostart
  PrevRow = -1
  for i = 1 to WaveItems          'Init Wave Items
    Do
      Row(i) = int(4*rnd)
    Loop until row(i) <> prevRow
    prevRow = row(i)

    select case int(7*rnd)
      case 0,1:   VM(i) = 0   '5Mio
      case 2,3,4,5,6: VM(i) = 1   'Mine
    end Select

    'Atom
    if (i=12) or (i=24) or (i=36) or (i=48) or (i=60) or (i=73) then
      VM(i) = 2
    end if

    'EB
    if i = 64 then
      if int(5*rnd) < 3 then
        VM(i) = 3
      end if
    end if
  next

  VMLoopCount = -8
  ItemCount = 1

  VMScore = 0
  VMLifes = 3
  VMText = "            000"
  if VMCount > 2 then
    VMLifes = 2
    VMText = "             00"
  end if
  WaitOneFrame = False
  Itemdestroyed = False
  VMTimer.Interval = VMTimerIntervall
  LeftShotTimer.Interval = VMTimerIntervall
  RightShotTimer.Interval = VMTimerIntervall
  VMStartTimer.enabled = True
  VideoModeReady = False
End Sub

Sub VMStartTimer_Timer
  VMStartTimer.enabled = False
  VMTimer.enabled = True
  VMMusicTimer.Interval = VideoModeThemeLen - 500
  select case int(3*rnd)
    case 0: PlaynextMusic(VideoModeTheme)
    case 1: PlaynextMusic(VideoModeTheme2)
    case 2: PlaynextMusic(SecretManiaTheme)
  end select
  If UltraDMD.IsRendering Then
    UltraDMDTimer.Enabled = False
    UltraDMD.CancelRendering
  end if
  UltraDMD.DisplayScoreboard 0, 0, 0, 0, 0, 0, FormatPointText(cstr(points)), VMText
End Sub

Sub VMTimer_Timer
  if VideoModeActive then
    VMLoopCount = VMLoopCount + 1
    if VMLoopCount < 0 then
      ActiveVMImage = "VM" & cstr(9+VMLoopCount) & ".png"
      UltraDMD.SetScoreboardBackgroundImage ActiveVMImage, 15, 15
    else
      if VMLoopCount > 7 then
        select case VM(Itemcount)           'Item collected:
          case 0: if not ItemDestroyed then     '5Mio
                Playsound "W_BigPoints"
                VMScore = VMScore + 5000000
                addpoints(5000000)
              end if
          case 1: if not ItemDestroyed then     'Mine
                playsound "W_VMMineHit"
                UltraDMD.SetScoreboardBackgroundImage "X_1.png", 15, 15
                UltraDMD.DisplayScoreboard 0, 0, 0, 0, 0, 0, FormatPointText(cstr(points)), VMText
                ErrorMade = True
                VMLifes = VMLifes - 1
                if VMLifes < 0 then
                  EndVM
                  Exit Sub
                end if
                WaitOneFrame = True
                Itemdestroyed = True
              end if
          case 2: if not ItemDestroyed then     'Atom
                playsound "W_VMAtom"
                VMScore = VMScore + (2000000 * VMCount)
                addpoints(2000000 * VMCount)
              end if
          case 3: if not ItemDestroyed then     'EB
                ExtraBalls = ExtraBalls + 1
                Playsound "W_ExtraBall"
                DOF 123, DOFPulse
              end if
        end select
        if not WaitOneFrame then
          VMLoopCount = 0
          if VM(Itemcount) = 2 then               'increase speed after each wave (Atom)
            VMTimer.Interval = VMTimer.Interval - WaveChange
            if VMTimer.Interval < 1 then
              VMTimer.Interval = 1
            end if
            VMTimer.enabled = False
            VMTimer.enabled = True
          end if
          ItemCount = Itemcount + 1
          if ItemCount > WaveItems then
            EndVM
            exit sub
          end if
          ItemDestroyed = False
        end if
      end if

      if (ItemCount <= WaveItems) and not WaitOneFrame then
        select case VMLifes
          case 1: VMText = "              0"
          case 2: VMText = "             00"
          case 3: VMText = "            000"
          case else: VMText = ""
        end select

        Select case VM(itemcount)
          case 0: ActiveVMImage = "F_"
          case 1: ActiveVMImage = "M_"
          case 2: ActiveVMImage = "A_"
          case 3: ActiveVMImage = "E_"
        end select
        Select case row(itemcount)
          case 0: ActiveVMImage = ActiveVMImage & "L_" & cstr(VMLoopCount+1)
          case 1: ActiveVMImage = ActiveVMImage & "ML_" & cstr(VMLoopCount+1)
          case 2: ActiveVMImage = ActiveVMImage & "MR_" & cstr(VMLoopCount+1)
          case 3: ActiveVMImage = ActiveVMImage & "R_" & cstr(VMLoopCount+1)
        end select
        if (((row(itemcount) = 0) or (row(itemcount) = 1)) and RightShotTimer.enabled) or (((row(itemcount) = 2) or (row(itemcount) = 3)) and LeftShotTimer.enabled) then
          ActiveVMImage = ActiveVMImage & "_S.png"
        else
          ActiveVMImage = ActiveVMImage & ".png"
        end if

        if itemdestroyed and (VMLoopCount > 5) then
          if VMLoopCount = 6 then
            Select case row(itemcount)
              case 0: ActiveVMImage = "X_L_7.png"
              case 1: ActiveVMImage = "X_ML_7.png"
              case 2: ActiveVMImage = "X_MR_7.png"
              case 3: ActiveVMImage = "X_R_7.png"
            end select
          else
            Select case row(itemcount)
              case 0,1:   if RightShotTimer.enabled then
                      ActiveVMImage = "X_L_8_S.png"
                    else
                      ActiveVMImage = "X_L_8.png"
                    end if
              case 2,3:   if LeftShotTimer.enabled then
                      ActiveVMImage = "X_R_8_S.png"
                    else
                      ActiveVMImage = "X_R_8.png"
                    end if
            end select
          end if
        end if

        UltraDMD.SetScoreboardBackgroundImage ActiveVMImage, 15, 15
        UltraDMD.DisplayScoreboard 0, 0, 0, 0, 0, 0, FormatPointText(cstr(points)), VMText
      end if
      if WaitOneFrame then
        WaitOneFrame = False
      end if
    end if
  end if
End Sub

Sub ShootVMItem(DirectionPar)
  if Directionpar = "Right" then
    RightShotTimer.enabled = True
    playsound "W_VMShotRight"
  else
    LeftShotTimer.enabled = True
    playsound "W_VMShotLeft"
  end if
  if ((Directionpar = "Left") and ((row(itemcount) = 0) or (row(itemcount) = 1))) or ((Directionpar = "Right") and ((row(itemcount) = 2) or (row(itemcount) = 3))) then
    if (VMLoopCount > 2) and not itemdestroyed then
      itemdestroyed = True
      select case VM(Itemcount)               'Item destroyed:
        case 0: ErrorMade = True              '5Mio
        case 1: VMScore = VMScore + (1000000 * VMCount)   'Mine
            addpoints(1000000 * VMCount)
        case 2: ErrorMade = True              'Atom
            playsound "W_VMMineHit"
            VMLifes = VMLifes - 1
            if VMLifes < 0 then
              EndVM
            end if
            select case VMLifes
              case 1: VMText = "              0"
              case 2: VMText = "             00"
              case 3: VMText = "            000"
              case else: VMText = ""
            end select
            UltraDMD.DisplayScoreboard 0, 0, 0, 0, 0, 0, FormatPointText(cstr(points)), VMText
        case 3: ErrorMade = True              'EB
      end select
    end if
  end if
End Sub

Sub RightShotTimer_Timer
  RightShotTimer.enabled = False
End Sub

Sub LeftShotTimer_Timer
  LeftShotTimer.enabled = False
End Sub

Sub EndVM
  VMTimer.enabled = False
  if VMTimerIntervall > 60 then
    VMTimerIntervall = VMTimerIntervall - VMChange
  end if

  if VMLifes < 0 then
    'VM Lost
    select case int(3*rnd)
      case 0: playsound "W_VMLost"
      case 1: playsound "W_VMLost2"
      case 2: playsound "W_ByeBye"
    end select
  else
    'Success
    ShowdownSpheres = ShowdownSpheres + 1
    if ErrorMade then
      Playsound "W_YouMadeIt"
      VMScore = VMScore + (100000000 * VMCount)
      addpoints(100000000 * VMCount)
    else
      playsound "W_YesYesYes"
      PerfectVMCount = PerfectVMCount + 1
      VMScore = VMScore + (250000000 * PerfectVMCount)
      addpoints(250000000 * PerfectVMCount)
    end if
  end if

  VideoModeActive = False

  Text1 = FormatDMDTopText("","","")
  Text2 = FormatDMDBottomText("")
  DMDDisplayScene "Main",True,backgrnd, Text1, 8,14,Text2, 5,15, 14, 200, 14
  VMRestartGameTimer.enabled = True
  VMTotalTimer.enabled = True
  SetB2SLamp 1,1
End Sub

Sub VMTotalTimer_Timer
  VMTotalTimer.enabled = False
  DMDEndVideoMode VMScore
End Sub

Sub VMRestartGameTimer_Timer
  VMRestartGameTimer.enabled = False
  VMMusicTimer.enabled = False
  PlayNextMusic(MainTheme)
  VideoModeLock.isdropped = True
  if FireExplosionVUK then
    VUKFire
  else
    if LockLight.state <> 0 then
      if BallsLocked = 0 then
        TakeYourPick
      else
        BallsLocked = BallsLocked + 1
        if BallsLocked = FastFrenzyBalls then
          DeactivateMissionsAndModes
          StartFastFrenzy
        else
          select case BallsLocked
            case 1: PlayNextsound "W_Ball1Locked"
            case 2: PlayNextsound "W_Ball2Locked"
            case 3: PlayNextsound "W_Ball3Locked"
          end select
          if BallsLocked = FastFrenzyBalls - 1 then
            LockLight.state = Lightstateblinking
          end if
          TroughLeftHoleKicker.CreateBall
          TroughLeftHoleKicker.kick 180, 1
        end if
      end if
    else
      TroughLeftHoleKicker.CreateBall
      TroughLeftHoleKicker.kick 180, 1
    end if
  end if
End Sub



'--- Space Station Frenzy ---

Dim BallsDocked,SpaceStationFrenzyActive,FastFrenzyActive,StartModeLightState,StartMissionLightState,LockLightState,VideomodeLightState,ExplosionLightState,QuickDockLightState
Sub DockBallLock_Hit
  playsound "drain3",0,0.3,0.1,0.25
  if not MissionActive then
    DeactivateComboShotLights
  end if
  DeactivateComboShotValues
  DockBallLock.destroyBall
  DMDDisplayScene "Dock2",True,BallDockedback, "", 8,14,"", 5,15, 14, 2500, 14
  BallsOnPF = BallsOnPF - 1
  BallsDocked = BallsDocked + 1
  if DockLight.state = Lightstateblinking then
    DockLight.state = 0
  end if
  if BallsDocked = 1 then
    PlayNextsound "W_Ball1Docked"
  end if
  if BallsDocked = 2 then
    PlayNextsound "W_Ball2Docked"
    Docklight.state = Lightstateblinking
    if not hideanimation then
      PlaynextMusic(IdleTheme)
    end if
  end if
  if BallsDocked = 3 then   'Start Space Station Frenzy
    DeactivateMissionsandModes
    PlayNextsound "W_SpaceStation"
    DMDStartMission "SpaceStation",10000
    FlashB2SLamp 51,1
    FlashB2SLamp 50,1
    if int(2*rnd) = 0 then
      playnextmusic(SpaceStationTheme)
    else
      playnextmusic(SpaceStationTheme2)
      QueuedMusic = SpaceStationTheme
    end if
  end if

  DockBallLock.Timerenabled = True
End Sub

Sub DockBallLock_Timer
  DockBallLock.Timerenabled = False
  if BallsDocked = 3 then   'Start Space Station Frenzy
    FlashB2SLamp 51,2
    SpaceStationFrenzyActive = True
    PrepareUltraJackpot = False
    UltraJackpotready = False
    SpaceStationJackpots = 0
    DockBallLock.enabled = False
    Docklight.state = 0
    StartModeLightState = StartModeLight.State
    StartMissionLightState = StartMissionLight.State
    LockLightState = LockLight.State
    VideomodeLightState = VideomodeLight.State
    ExplosionLightState = ExplosionLight.State
    QuickDockLightState = QuickDockLight.State
    StartModeLight.State = 0
    StartMissionLight.State = 0
    LockLight.State = 0
    VideomodeLight.State = 0
    ExplosionLight.State = 0
    QuickDockLight.State = 0
    LeftRampVUK.CreateBall
    LeftRampVUK.kick 180, 1
    playsound SoundFXDOF("W_Vuk",134,DOFPulse,DOFContactors),0,1,-0.18,0.25
    RightRampVUK.CreateBall
    RightRampVUK.kick 195, 1
    playsound SoundFXDOF("W_Vuk2",135,DOFPulse,DOFContactors),0,1,0.18,0.25
    DOF 136, DOFPulse
    BallsOnPF = BallsOnPF + 2
    if DockAdditionalSpaceStationBall then
      DockAdditionalSpaceStationBall = False
      MultiBallBalls = 4
      BallsSaved = BallsSaved + 1
      BallReleaseTimer.enabled = True
    else
      MultiBallBalls = 3
    end if
    GracePeriodActive = True
    BallSaverLight.BlinkInterval = 160
    GracePeriod.interval = 20000
    GracePeriod1.interval = 23000
    GracePeriod2.interval = 25000
    StartBallSaver
    JPTimer.enabled = True
    Jackpotlight.state = Lightstateblinking
  end if
  GracePeriod.interval  =  9000
  GracePeriod1.interval = 12000
  GracePeriod2.interval = 14000
  StartBallSaver
  BallRelease.CreateBall
  BallRelease.kick 90,5
  DOF 131, DOFPulse
  BallsOnPF = BallsOnPF + 1
  AutoPlungerTimer.enabled = True
End Sub

Sub CheckSpaceStationFrenzy
  if SpaceStationFrenzyActive then
    if PrepareUltraJackpot and (DT1Light.state = 1) and (DT2Light.state = 1) and (DT3Light.state = 1) and (DT4Light.state = 1) and (DT5Light.state = 1) and (DT6Light.state = 1) then
      playnextsound("W_YouAreClose")
      ExplosionHoleSpider.state = Lightstateblinking
    end if
  end if
End Sub

Dim SpaceStationJackpots,AwardedScore
Sub AwardSpaceStationJackpot
  FlashB2SLamp 51,1
  FlashB2SLamp 50,1

  if CowJackpotActive then
    CowJackpotActive = False
  end if
  if CowJackpotReady then
    CowJackpotReady = False
    CowJackpotActive = True
    LeftHoleSpider.state = Lightstateblinking
    LeftHoleSpider.timerenabled = True
  end if
  SpaceStationJackpots = SpaceStationJackpots + 1
  if (SuperJackpotLight.state = 0) and (UltraJackpotLight.state = 0) then
    Addpoints(JackpotValue * SpaceStationJackpots)
    DOF 139, DOFPulse
  else
    SpaceStationJackpots = 1
    Addpoints(JackpotValue)
  end if
  select case SpaceStationJackpots
    case 1: PlayNextSound "W_Jackpot"
        AwardedScore = Jackpotvalue
        DMDDisplayScene "SSJack",True,SSJackback,FormatDMDTopText("","",""), 8,14,FormatDMDBottomText(""), 5,15, 14, 4000, 14
    case 2: PlayNextSound "W_DoubleJackpot"
        AwardedScore = Jackpotvalue * 2
        DMDDisplayScene "SSJack",True,SSJack2back,FormatDMDTopText("","",""), 8,14,FormatDMDBottomText(""), 5,15, 14, 4500, 14
    case 3: PlayNextSound "W_TripleJackpot"
        AwardedScore = Jackpotvalue * 3
        DMDDisplayScene "SSJack",True,SSJack3back,FormatDMDTopText("","",""), 8,14,FormatDMDBottomText(""), 5,15, 14, 5000, 14
  end select
  SSJackpotTimer.enabled = True
  if SpaceStationJackpots = 3 then
    SuperJackpotLight.state = Lightstateblinking
    PlayNextSound "W_SuperJackpotLit"
  end if
End Sub

Sub SSJackpotTimer_Timer
  SSJackpotTimer.enabled = False
  Playsound "W_USIntitBoom",0,0.75,0,0.1
  FlashLeft 2
  UfoHitAnimation 1
  UpdateDMDText "SSJack","","","",FormatPointText(cstr(Awardedscore))
End Sub

Dim CowJackpotReady,CowJackpotActive
Sub AwardCowJackpot
  CowJackpotActive = False
  DMDDisplayScene "COWJP",True,Cowback,FormatDMDTopText("","",""), 8,14,FormatDMDBottomText(""), 5,15, 14, 3000, 14
  CowJackpotTimer.enabled = True
  addpoints(JackpotValue * 3)
  DOF 139, DOFPulse
  PlayNextSound "W_CowJackpot"
End Sub

Sub CowJackpotTimer_Timer
  CowJackpotTimer.enabled = False
  Playsound "W_USIntitBoom",0,0.75,0,0.1
  UpdateDMDText "COWJP","","","",FormatPointText(cstr(Jackpotvalue*3))
End Sub

Sub LeftHoleSpider_Timer
  LeftHoleSpider.timerenabled = False
  LeftHoleSpider.state = 0
End Sub

Dim PrepareUltraJackpot,UltraJackpotready
Sub AwardSpaceStationSuperJackpot
  FlashB2SLamp 51,2
  FlashB2SLamp 50,2
  DMDDisplayScene "SSSJack",True,SSSJackback,FormatDMDTopText("","",""), 8,14,FormatDMDBottomText(""), 5,15, 14, 5500, 14
  SSJP = 0
  SSSJackpotTimer.enabled = True
  SuperJackpotLight.state = 0
  SpaceStationJackpots = 0
  CowJackpotReady = True
  Addpoints(JackpotValue * 5)
  DOF 139, DOFPulse
  DTDropUp 0
  if PrepareUltraJackpot then
    if DT1Light.state = 0 then
      DT1Light.state = Lightstateblinking
    end if
    if DT2Light.state = 0 then
      DT2Light.state = Lightstateblinking
    end if
    if DT3Light.state = 0 then
      DT3Light.state = Lightstateblinking
    end if
    if DT4Light.state = 0 then
      DT4Light.state = Lightstateblinking
    end if
    if DT5Light.state = 0 then
      DT5Light.state = Lightstateblinking
    end if
    if DT6Light.state = 0 then
      DT6Light.state = Lightstateblinking
    end if
  else
    DT1Light.state = Lightstateblinking
    DT2Light.state = Lightstateblinking
    DT3Light.state = Lightstateblinking
    DT4Light.state = Lightstateblinking
    DT5Light.state = Lightstateblinking
    DT6Light.state = Lightstateblinking
  end if
  PrepareUltraJackpot = True
  OverloadLight.state = Lightstateblinking
end sub

Dim SSJP
Sub SSSJackpotTimer_Timer
  SSJP = SSJP + 1
  if SSJP = 1 then
    Playsound "W_USIntitBoom",0,0.75,0,0.1
    FlashLeft 1
    FlashRight 2
    UfoHitAnimation 0
  end if

  if SSJP = 2 then
    PlayNextSound "W_SuperJackpot"
    UpdateDMDText "SSSJack","","","",FormatPointText(cstr(Jackpotvalue*5))
  end if

  if SSJP >= 3 then
    SSSJackpotTimer.enabled = False
    FlashLeft 1
    FlashRight 2
    UfoHitAnimation 0
  end if
End Sub

Sub UltraJackpotTimer_Timer
  UltraJackpotTimer.enabled = False
  PrepareUltraJackpot = True
  UltraJackpotready = False
  UltraJackpotLight.state = 0
  OverloadLight.state = Lightstateblinking
  ExplosionHoleSpider.state = 0
  DTDropUp 0
  DT1Light.state = Lightstateblinking
  DT2Light.state = Lightstateblinking
  DT3Light.state = Lightstateblinking
  DT4Light.state = Lightstateblinking
  DT5Light.state = Lightstateblinking
  DT6Light.state = Lightstateblinking
end sub

Sub AwardSpaceStationUltraJackpot
  FlashB2SLamp 51,3
  FlashB2SLamp 50,3
  UltraJackpotTimer.enabled = False
  UltraJackpotready = False
  SSJP = 0
  SSUJackpotTimer.enabled = True
  DMDDisplayScene "SSUJack",True,SSUJackback,FormatDMDTopText("","",""), 8,14,FormatDMDBottomText(""), 5,15, 14, 6500, 14
  UltraJackpotLight.state = 0
  Addpoints(JackpotValue * 10)
  US_UltraLight.state = 1
  OverloadLight.state = 0
  DT1Light.state = 0
  DT2Light.state = 0
  DT3Light.state = 0
  DT4Light.state = 0
  DT5Light.state = 0
  DT6Light.state = 0
end sub

Sub SSUJackpotTimer_Timer
  SSJP = SSJP + 1
  if SSJP = 3 then
    Playsound "W_USIntitBoom",0,0.75,0,0.1
    FlashLeft 2
    FlashRight 3
    UfoHitAnimation 0
  end if

  if SSJP = 7 then
    PlayNextSound "W_UltraJackpot"
    UpdateDMDText "SSUJack","","","",FormatPointText(cstr(Jackpotvalue*10))
  end if

  if SSJP >= 12 then
    SSUJackpotTimer.enabled = False
    FlashLeft 4
    FlashRight 4
    UfoHitAnimation 0
  end if
End Sub

Sub EndSpaceStationFrenzy
  SpaceStationFrenzyActive = False
  PrepareUltraJackpot = False
  UltraJackpotready = False
  ExplosionHoleSpider.state = 0
  UltraJackpotLight.state = 0
  SuperJackpotLight.state = 0
  JackpotLight.state = 0
  OverloadLight.state = 0
  DT1Light.state = 0
  DT2Light.state = 0
  DT3Light.state = 0
  DT4Light.state = 0
  DT5Light.state = 0
  DT6Light.state = 0
  DTDropUp 0
  CowJackpotReady = False
  CowJackpotActive = False
  LeftHoleSpider.timerenabled = False
  LeftHoleSpider.state = 0
  DMDEndMode
  playnextmusic(MainTheme)
  MultiBallBalls = 0
  BallsDocked = 0

  StartModeLight.State = StartModeLightState
  StartMissionLight.State = StartMissionLightState
  LockLight.State = LockLightState
  LockLightstate = 0
  VideomodeLight.State = VideomodeLightState
  ExplosionLight.State = ExplosionLightState
  QuickDockLight.State = QuickDockLightState
End Sub

'--- Combos ---
Dim ActualComboCount,ComboLightWasLit,ComboManiaValue,ShowComboStatus
Sub AwardNextCombo
  FlashB2SLamp 48,1
  FlashB2SLamp 49,1
  FlashB2SLamp 51,1

  if MultiBallBalls < 2 then
    ActualComboCount = ActualComboCount + 1
    if (ActualComboCount = 5) and (US_CombosLight.state = 0) then
      US_CombosLight.state = 1
    end if
    for i = 1 to ActualComboCount
      TotalComboCount = TotalComboCount + 1
      if (TotalComboCount mod 15 = 0) and not CombomaniaActive then
        ComboManiaReady = True
      end if
    next

    if CombomaniaActive then
      if Actualcombocount > 6 then
        addpoints(ComboManiaValue*6)
        DOF 139, DOFPulse
      else
        addpoints(ComboManiaValue*Actualcombocount)
      end if
      select case ActualComboCount
        case 1: if not Hideanimation then
              select case int(3*rnd)
                case 0: DMDDisplayScene "C1",True,Combo1back, "", 8,14,"", 5,15, 14, 1800, 1
                case 1: DMDDisplayScene "C2",True,Combo2back, "", 8,14,"", 5,15, 14, 1800, 1
                case 2: DMDDisplayScene "C3",True,Combo3back, "", 8,14,"", 5,15, 14, 1800, 1
              end select
            end if
            if int(2*rnd) = 0 then
              PlaySound "W_Combo"
            else
              PlaySound "W_Combo_F"
            end if
        case 2: if not Hideanimation then
              DMDDisplayScene "DC",True,DComboback, "", 8,14,"", 5,15, 14, 1800, 1
            end if
        case 3: if not Hideanimation then
              DMDDisplayScene "TC",True,TComboback, "", 8,14,"", 5,15, 14, 1800, 1
            end if
        case else:  if not Hideanimation then
                DMDDisplayScene "SC",True,SComboback, "", 8,14,"", 5,15, 14, 3300, 1
              end if
      end select
      if ActualComboCount > 1 then
        DOF 139, DOFPulse
        select case int(8*rnd)
          case 0: playsound "W_NiceShot"
          case 1: playsound "W_Superb"
          case 2: playsound "W_IamImpressed"
          case 3: playsound "W_Outstanding"
          case 4: playsound "W_VeryGood"
          case 5: playsound "W_YesYesYes"
          case 6: playsound "W_Unbelievable"
          case 7: PlaySound "W_Yeehaa"
        end select
      end if
    else
      addpoints(1000000*Actualcombocount)
      DOF 139, DOFPulse
      select case ActualComboCount
        case 1: if not Hideanimation then
              select case int(3*rnd)
                case 0: DMDDisplayScene "C1",True,Combo1back, "", 8,14,"", 5,15, 14, 1800, 1
                case 1: DMDDisplayScene "C2",True,Combo2back, "", 8,14,"", 5,15, 14, 1800, 1
                case 2: DMDDisplayScene "C3",True,Combo3back, "", 8,14,"", 5,15, 14, 1800, 1
              end select
            end if
            if int(2*rnd) = 0 then
              PlayNextSound "W_Combo"
            else
              PlayNextSound "W_Combo_F"
            end if
        case 2: PlayNextSound "W_DoubleCombo"
            if not Hideanimation then
              DMDDisplayScene "DC",True,DComboback, "", 8,14,"", 5,15, 14, 1800, 1
            end if
        case 3: PlayNextSound "W_TripleCombo"
            if not Hideanimation then
              DMDDisplayScene "TC",True,TComboback, "", 8,14,"", 5,15, 14, 1800, 1
            end if
        case 4: PlayNextSound "W_SuperCombo"
            if not Hideanimation then
              DMDDisplayScene "SC",True,SComboback, "", 8,14,"", 5,15, 14, 3300, 1
            end if
        case 5: PlayNextSound "W_Yeehaa"
            if not Hideanimation then
              DMDDisplayScene "SC",True,SComboback, "", 8,14,"", 5,15, 14, 3300, 1
            end if
        case else:  if not Hideanimation then
                DMDDisplayScene "SC",True,SComboback, "", 8,14,"", 5,15, 14, 3300, 1
              end if
              select case int(7*rnd)
                case 0: playNextsound "W_NiceShot"
                case 1: playNextsound "W_Superb"
                case 2: playNextsound "W_IamImpressed"
                case 3: playNextsound "W_Outstanding"
                case 4: playNextsound "W_VeryGood"
                case 5: playNextsound "W_YesYesYes"
                case 6: playNextsound "W_Unbelievable"
              end select
      end select
      ShowComboStatus = True
    end if
  end if
End Sub

Sub ComboStatusTimer_Timer
  ComboStatusTimer.enabled = False
  if not hideanimation and not CombomaniaActive then
    if ComboManiaReady then
      UpdateDMDText "Main","","COMBOMANIA READY","",""
    else
      UpdateDMDText "Main","","COMBOMANIA  IN","",cstr(15 - (TotalComboCount mod 15))
    end if
  end if
End Sub


Sub RightRampComboTrigger_Hit
  if not SkillShotComboTimer.enabled then
    ActivateNextComboShots("RightRamp")
    if ComboTimer.enabled then
      ComboTimer.enabled = False
      ComboTimer.enabled = True
    end if
  end if
End Sub

Sub LeftRampComboTrigger_Hit
  if not SkillShotComboTimer.enabled then
    ActivateNextComboShots("LeftRamp")
    if ComboTimer.enabled then
      ComboTimer.enabled = False
      ComboTimer.enabled = True
    end if
  end if
End Sub

Sub RightRampComboTrigger1_Hit
  if ComboTimer.enabled then
    ComboTimer.enabled = False
    ComboTimer.enabled = True
  end if
End Sub

Sub LeftRampComboTrigger1_Hit
  if ComboTimer.enabled then
    ComboTimer.enabled = False
    ComboTimer.enabled = True
  end if
End Sub

Sub InitLeftOrbit_Hit
  ComboLightWasLit = (LeftOrbitSpider.UserValue = 1)
  InitLeftOrbit.Timerenabled = False
  InitLeftOrbit.Timerenabled = True
End Sub
Sub InitLeftOrbit_Timer
  InitLeftOrbit.Timerenabled = False
End Sub

Sub InitRightOrbit_Hit
  ComboLightWasLit = (RightOrbitSpider.UserValue = 1)
  InitRightOrbit.Timerenabled = False
  InitRightOrbit.Timerenabled = True
End Sub
Sub InitRightOrbit_Timer
  InitRightOrbit.Timerenabled = False
End Sub

Sub RightOrbit_Hit
  if ComboTimer.enabled then
    ComboTimer.enabled = False
    ComboTimer.enabled = True
  end if
  if InitRightOrbit.Timerenabled then
    ActivateNextComboShots("RightOrbit")
    RightOrbitComboTimer.enabled = False
    RightOrbitComboTimer.enabled = True
    if combotimer.enabled and ComboLightWasLit then
      ComboTimer.enabled = False
      ComboTimer.enabled = True
    else
      if Skillshot = 0 then
        actualcombocount = 0
        ComboTimer.enabled = True
      end if
    end if
  end if
  if LeftOrbitComboTimer.enabled then
    if BikeRaceActive then
      if LeftOrbitSpider.state <> 0 then
        AdvanceBikeRaceScore(False)
      end if
    end if
    if SecretManiaActive then
      AwardSecretManiaScore
    end if
    if Quickshotactive then
      if Ejecttoright then
        EndQuickShot
      end if
    end if
    if combotimer.enabled and Combolightwaslit then
      AwardNextCombo
      ComboTimer.enabled = False
      ComboTimer.enabled = True
    else
      if (Skillshot = 0) then
        actualcombocount = 0
        ComboTimer.enabled = True
        if CombomaniaActive then
          AwardnextCombo
        end if
      end if
    end if
  end if

End Sub

Sub RightOrbitComboTimer_Timer
  RightOrbitComboTimer.enabled = False
End Sub

Sub LeftOrbitComboTimer_Timer
  LeftOrbitComboTimer.enabled = False
End Sub

Sub ComboTimer_Timer
  ComboTimer.enabled = False
  if ShowComboStatus then
    ComboStatusTimer.enabled = True
  end if
  ShowComboStatus = False
  DeactivateComboShotValues
  if not ComboManiaActive and not HuntDownActive and not ShuttleActive and not BikeRaceActive and not SkyscraperActive and not SecretManiaActive and not UltimateShowdownActive then
    DeactivateComboShotlights
  end if
End Sub

Sub ActivateNextComboShots(ComboPar)
  DeactivateComboShotValues
  if MultiBallBalls < 2 then
    if not ComboManiaActive and not HuntDownActive and not ShuttleActive and not BikeRaceActive and not SkyscraperActive and not SecretManiaActive and not UltimateShowdownActive then
      DeactivateComboShotLights
      Select case ComboPar
        case "RightOrbit":
          LoopSpider.state = Lightstateblinking
          LoopSpider.UserValue = 1
          DockRampSpider.state = Lightstateblinking
          DockRampSpider.UserValue = 1
        case "LeftOrbit":
          LeftRampSpider.state = Lightstateblinking
          LeftRampSpider.UserValue = 1
        case "Loop":
          DockRampSpider.state = Lightstateblinking
          DockRampSpider.UserValue = 1
        case "LeftRamp":
          RightOrbitSpider.state = Lightstateblinking
          RightOrbitSpider.UserValue = 1
        case "RightRamp":
          LeftRampSpider.state = Lightstateblinking
          LeftRampSpider.UserValue = 1
          LeftOrbitSpider.state = Lightstateblinking
          LeftOrbitSpider.UserValue = 1
      end select
    else
      Select case ComboPar
        case "RightOrbit":
          LoopSpider.UserValue = 1
          DockRampSpider.UserValue = 1
        case "LeftOrbit":
          LeftRampSpider.UserValue = 1
        case "Loop":
          DockRampSpider.UserValue = 1
        case "LeftRamp":
          RightOrbitSpider.UserValue = 1
        case "RightRamp":
          LeftRampSpider.UserValue = 1
          LeftOrbitSpider.UserValue = 1
      end select
    end if
  end if
End Sub

Sub DeactivateComboShotLights
  LeftOrbitSpider.state = 0
  LeftRampSpider.state = 0
  LoopSpider.state = 0
  DockRampSpider.state = 0
  RightOrbitSpider.state = 0
End Sub

Sub DeactivateComboShotValues
  LeftOrbitSpider.UserValue = 0
  LeftRampSpider.UserValue = 0
  LoopSpider.UserValue = 0
  DockRampSpider.UserValue = 0
  RightOrbitSpider.UserValue = 0
End Sub


'--- Loops ---
Dim LoopExtraBalls
Sub LoopTopTrigger_Hit
  if LeftSpinnerTrigger.Timerenabled then
    if MultiBallBalls < 2 then
      LeftPowerLevelTimer.enabled = False
      LeftPowerLevelTimer.enabled = True
      if not SkyscraperActive and not Quickshotactive then
        AdvanceLeftPowerlevelLight.state = Lightstateblinking
      end if
    end if
    if UltimateShowdownActive and USLeft then
      AddPhase2Spins
    end if
  end if
  if RightSpinnerTrigger.Timerenabled then
    if MultiBallBalls < 2 then
      RightPowerLevelTimer.enabled = False
      RightPowerLevelTimer.enabled = True
      if not SkyscraperActive and not Quickshotactive then
        AdvanceRightPowerlevelLight.state = Lightstateblinking
      end if
    end if
    if UltimateShowdownActive and USRight then
      AddPhase2Spins
    end if
  end if
  if LoopTimer.enabled then
    ActivateNextComboShots("Loop")
  end if
  if EBLight.Timerenabled then
    EBLight.Timerenabled = False
    EBLight.Timerenabled = True
  end if

  if NextLoopTimer.enabled then
    NextLoopTimer.enabled = False
    NextLoopTimer.enabled = True
  end if
  if LoopTimer.enabled then
    LoopTimer.enabled = False
    LoopTimer.enabled = True
  end if
  playsound "W_LeftRamp"

  'Speed Control
  if abs(Activeball.velx) > MaxTopSpeed then
    Activeball.velx = MaxTopSpeed * (Activeball.velx)/abs(Activeball.velx)
  end if

End Sub

Sub LoopStart_Hit
  LeftSpinnerTrigger.Timerenabled = False
  RightSpinnerTrigger.Timerenabled = False

  if not LoopTimer.enabled then
    if not Hideanimation then
      DMDDisplayScene "LoopA",True,LoopAback, "", 8,14,"", 5,15, 14, 1500, 14
    end if
  end if

  if Skillshot = 1 then
    SuperSkillShot = SuperSkillShot + 1
    if SuperSkillShot < 3 then
      SkillshotTimer.enabled = False
      SkillshotTimer.enabled = True
    end if
  end if

  if combotimer.enabled then
    ComboLightWasLit = (LoopSpider.UserValue = 1)
    ComboTimer.enabled = False
    ComboTimer.enabled = True
  else
    if Skillshot = 0 then
      actualcombocount = 0
      ComboTimer.enabled = True
    end if
  end if

  LoopTimer.enabled = False
  LoopTimer.enabled = True

  if NextLoopTimer.enabled then
    NextLoopTimer.enabled = False
    NextLoopTimer.enabled = True
  end if

End Sub

Sub LoopTimer_Timer()
  LoopTimer.enabled = False
  LoopCount = 0
End Sub

Sub NextLoopTimer_Timer()
  NextLoopTimer.enabled = False
End Sub

Sub LeftOrbit_hit
  if LoopTimer.enabled then
    FlashB2SLamp 50,1
    TotalLoopCount = TotalLoopCount + 1
    if NextLoopTimer.enabled then
      LoopCount = LoopCount + 1
    else
      LoopCount = 1
    end if

    if not Hideanimation then
      DMDDisplayScene "LoopB",True,LoopBback, "", 8,14,FormatDMDBottomText(cstr(LoopCount)&"          "), 5,15, 14, 1200, 1
    end if

    addpoints(5000000 * LoopCount)

    if BikeRaceActive then
      if LoopSpider.state <> 0 then
        AdvanceBikeRaceScore(True)
      end if
    end if
    if SecretManiaActive then
      AwardSecretManiaScore
    end if
    if LoopCount = 2 + LoopExtraBalls then
      if EBLight.state = 0 then
        EBLight.Timerenabled = True
        EBLight.state = Lightstateblinking
        PlayNextSound "W_GetExtraBall"
      end if
    end if
    if LoopCount = 3 then
      if not BonusHeld then
        BonusHeld = True
        UpdateDMDText "Main","","Bonus held","",""
        PlayNextSound "W_BonusHeld"
      end if
    end if
    if LoopCount = 4 then
      if not SuperJetsHeld then
        SuperJetsHeld = True
        UpdateDMDText "Main","","Superjets held","",""
        PlayNextSound "W_SuperJets"
      end if
    end if
    NextLoopTimer.enabled = False
    NextLoopTimer.enabled = True
    LoopTimer.enabled = False
    LoopTimer.enabled = True
  end if

  if ComboTimer.enabled then
    ComboTimer.enabled = False
    ComboTimer.enabled = True
  end if
  if InitLeftOrbit.Timerenabled then
    ActivateNextComboShots("LeftOrbit")
    LeftOrbitComboTimer.enabled = False
    LeftOrbitComboTimer.enabled = True
    if combotimer.enabled and ComboLightWasLit then
      ComboTimer.enabled = False
      ComboTimer.enabled = True
    else
      if Skillshot = 0 then
        actualcombocount = 0
        ComboTimer.enabled = True
      end if
    end if
  end if
  if CombomaniaActive and LoopTimer.enabled then
    AwardnextCombo
  end if
  if RightOrbitComboTimer.enabled then
    if BikeRaceActive then
      if RightOrbitSpider.state <> 0 then
        AdvanceBikeRaceScore(False)
      end if
    end if
    if SecretManiaActive then
      AwardSecretManiaScore
    end if
    if Quickshotactive then
      if not Ejecttoright then
        EndQuickShot
      end if
    end if
    if combotimer.enabled and ComboLightWasLit then
      AwardNextCombo
      ComboTimer.enabled = False
      ComboTimer.enabled = True
    else
      if (Skillshot = 0) then
        actualcombocount = 0
        ComboTimer.enabled = True
        if CombomaniaActive then
          AwardnextCombo
        end if
      end if
    end if
  end if
end Sub

Sub EBLight_Timer
  EBLight.Timerenabled = False
  EBLight.state = 0
End Sub

'--- Reactor and Drop Targets ---
Sub DT1_Hit
  DOF 113, DOFPulse
  DT1.isdropped = True
  Playsound "W_DT",0,1,0.06
  if ReactorCriticalActive then
    AddReactorScore
  else
    if (DT1Light.state = 1) and (DT2Light.state = 1) and (DT3Light.state = 1) then
      LightRightDTLight
    else
      DT1Light.state = 1
      CheckFastFrenzy
      CheckSpaceStationFrenzy
    end if
    CheckDropTargets
  end if
  AdvanceShuttleRampValue
  AwardUSPoints
  DT1timer.enabled = True
End Sub
Sub DT1Timer_Timer
  DT1timer.enabled = False
  if ReactorCriticalActive then
    DT1.isdropped = False
    Playsound "TargetBankReset1",0,0.1,0.05,0.25
  Else
    CheckDropTargets
  End If
End Sub

Sub DT2_Hit
  DOF 113, DOFPulse
  DT2.isdropped = True
  Playsound "W_DT",0,1,0.07
  if ReactorCriticalActive then
    AddReactorScore
  else
    if (DT1Light.state = 1) and (DT2Light.state = 1) and (DT3Light.state = 1) then
      LightRightDTLight
    else
      DT2Light.state = 1
      CheckFastFrenzy
      CheckSpaceStationFrenzy
    end if
  end if
  AdvanceShuttleRampValue
  AwardUSPoints
  DT2timer.enabled = True
End Sub
Sub DT2Timer_Timer
  DT2timer.enabled = False
  if ReactorCriticalActive then
    DT2.isdropped = False
    Playsound "TargetBankReset1",0,0.1,0.05,0.25
  Else
    CheckDropTargets
  End If
End Sub

Sub DT3_Hit
  DOF 113, DOFPulse
  DT3.isdropped = True
  Playsound "W_DT",0,1,0.08
  if ReactorCriticalActive then
    AddReactorScore
  else
    if (DT1Light.state = 1) and (DT2Light.state = 1) and (DT3Light.state = 1) then
      LightRightDTLight
    else
      DT3Light.state = 1
      CheckFastFrenzy
      CheckSpaceStationFrenzy
    end if
  end if
  AdvanceShuttleRampValue
  AwardUSPoints
  DT3timer.enabled = True
End Sub
Sub DT3Timer_Timer
  DT3timer.enabled = False
  if ReactorCriticalActive then
    DT3.isdropped = False
    Playsound "TargetBankReset1",0,0.1,0.05,0.25
  Else
    CheckDropTargets
  End If
End Sub

Sub DT4_Hit
  DOF 113, DOFPulse
  DT4.isdropped = True
  Playsound "W_DT",0,1,0.1
  if ReactorCriticalActive then
    AddReactorScore
  else
    if (DT4Light.state = 1) and (DT5Light.state = 1) and (DT6Light.state = 1) then
      LightLeftDTLight
    else
      DT4Light.state = 1
      CheckFastFrenzy
      CheckSpaceStationFrenzy
    end if
  end if
  AdvanceShuttleRampValue
  AwardUSPoints
  DT4timer.enabled = True
End Sub
Sub DT4Timer_Timer
  DT4timer.enabled = False
  if ReactorCriticalActive then
    DT4.isdropped = False
    Playsound "TargetBankReset1",0,0.1,0.05,0.3
  Else
    CheckDropTargets
  End If
End Sub

Sub DT5_Hit
  DOF 113, DOFPulse
  DT5.isdropped = True
  Playsound "W_DT",0,1,0.11
  if ReactorCriticalActive then
    AddReactorScore
  else
    if (DT4Light.state = 1) and (DT5Light.state = 1) and (DT6Light.state = 1) then
      LightLeftDTLight
    else
      DT5Light.state = 1
      CheckFastFrenzy
      CheckSpaceStationFrenzy
    end if
  end if
  AdvanceShuttleRampValue
  AwardUSPoints
  DT5timer.enabled = True
End Sub
Sub DT5Timer_Timer
  DT5timer.enabled = False
  if ReactorCriticalActive then
    DT5.isdropped = False
    Playsound "TargetBankReset1",0,0.1,0.05,0.3
  Else
    CheckDropTargets
  End If
End Sub

Sub DT6_Hit
  DOF 113, DOFPulse
  DT6.isdropped = True
  Playsound "W_DT",0,1,0.12
  if ReactorCriticalActive then
    AddReactorScore
  else
    if (DT4Light.state = 1) and (DT5Light.state = 1) and (DT6Light.state = 1) then
      LightLeftDTLight
    else
      DT6Light.state = 1
      CheckFastFrenzy
      CheckSpaceStationFrenzy
    end if
  end if
  AdvanceShuttleRampValue
  AwardUSPoints
  DT6timer.enabled = True
End Sub
Sub DT6Timer_Timer
  DT6timer.enabled = False
  if ReactorCriticalActive then
    DT6.isdropped = False
    Playsound "TargetBankReset1",0,0.1,0.05,0.3
  Else
    CheckDropTargets
  End If
End Sub

Sub LightLeftDTLight
  CheckFastFrenzy
  if DT1Light.state = 0 then
    DT1Light.state = 1
  else
    if DT2Light.state = 0 then
      DT2Light.state = 1
    else
      DT3Light.state = 1
    end if
  end if
End Sub

Sub LightRightDTLight
  CheckFastFrenzy
  if DT4Light.state = 0 then
    DT4Light.state = 1
  else
    if DT5Light.state = 0 then
      DT5Light.state = 1
    else
      DT6Light.state = 1
    end if
  end if
End Sub

Sub ReactorHole_Hit()
  ReactorHole.DestroyBall
  if UltimateShowdownActive then
    Playsound "Drain3",0,0.3,0.1,0.25
    AwardUSPoints
  else
    if UltraJackpotReady then
      PlayNextSound "W_OhNo"
      UltraJackpotReady = False
      PrepareUltraJackpot = True
      DTDropUp 0
      DT1Light.state = Lightstateblinking
      DT2Light.state = Lightstateblinking
      DT3Light.state = Lightstateblinking
      DT4Light.state = Lightstateblinking
      DT5Light.state = Lightstateblinking
      DT6Light.state = Lightstateblinking
      ExplosionHoleSpider.state = 0
    else
      Playsound "Drain3",0,0.3,0.1,0.25
      if ReactorCriticalActive = False then
        if not FastFrenzyActive and not SpaceStationFrenzyActive and not SecretManiaActive then
          if DT3Light.state = 0 then
            DT3Light.state = 1
          else
            if DT4Light.state = 0 then
              DT4Light.state = 1
            else
              if DT2Light.state = 0 then
                DT2Light.state = 1
              else
                if DT5Light.state = 0 then
                  DT5Light.state = 1
                else
                  if DT1Light.state = 0 then
                    DT1Light.state = 1
                  else
                    DT6Light.state = 1
                  end if
                end if
              end if
            end if
          end if
        end if
      else
        if Reactorcompletion >= 99 then
          Reactorcompletion = 100
          EndReactorCritical
        else
          if OverloadLight.state <> 0 then
            if Reactorcompletion >= 20 then
              Reactorcompletion = Reactorcompletion - 20
            else
              Reactorcompletion = 0
            end if
          end if
        end if
      end if
    end if
  end if

  if HuntDownActive and (ReactorHolespider.state <> 0) then
    AdvanceHuntDownScore
  end if

  if SecretManiaActive then
    AwardSecretManiaScore
  end if

  if LightMagnetLight.state <> 0 then
    QuickDockTimer.enabled = False
    ActivateMagnetLight.state = 1
    LightMagnetLight.state = 0
    PlayNextSound "W_MagnetIsLit"
    if not Hideanimation then
      DMDDisplayScene "MagLit",True,MagnetLitback, "", 8,14,"", 5,15, 14, 2000, 1
    end if
  end if

  TroughReactorKicker.CreateBall
  TroughReactorKicker.kick 180, 1
  ReactorHole.timerenabled = True
End Sub

Sub ReactorHole_Timer
  ReactorHole.timerenabled = False
  CheckDropTargets
End Sub


'Drop Targets
Sub CheckDropTargets
  if not FastFrenzyActive and not SpaceStationFrenzyActive and not ReactorCriticalActive and not SecretManiaActive then
    if DT1.isdropped and DT2.isdropped and DT3.isdropped then
      DTDropUp 1
    end if
    if DT4.isdropped and DT5.isdropped and DT6.isdropped then
      DTDropUp 2
    end if
  end if

  if (DT1Light.state = 1) and (DT2Light.state = 1) and (DT3Light.state = 1) and (DT4Light.state = 1) and (DT5Light.state = 1) and (DT6Light.state = 1) then
    if not FastFrenzyActive and not SpaceStationFrenzyActive and not ReactorCriticalActive and not SecretManiaActive then
      DTDropUp 0
      DT1Light.state = 0
      DT2Light.state = 0
      DT3Light.state = 0
      DT4Light.state = 0
      DT5Light.state = 0
      DT6Light.state = 0
      if not UltimateShowdownActive then
        DockBallLock.enabled = True
        LightMagnetLight.state = 1
        if DockLight.state = 0 then
          DMDDisplayScene "DockLit",True,DockLitback, "", 8,14,"", 5,15, 14, 2400, 14
          DockLight.state = 1
          PlayNextsound "W_DockIsLit"
        end if
      end if
    end if
  end if
End Sub

Sub DTDropUp(BankPar)
  Playsound "TargetBankReset1",0,0.2,0.05,0.25
  if (BankPar = 1) or (BankPar = 0) then
    DT1.isdropped = False
    DT2.isdropped = False
    DT3.isdropped = False
  end if

  if (BankPar = 2) or (BankPar = 0) then
    DT4.isdropped = False
    DT5.isdropped = False
    DT6.isdropped = False
  end if
  DTRecoveryTimer.enabled = True
End Sub

Sub DTRecoveryTimer_Timer
  DTRecoveryTimer.enabled = False
  if (DT1Light.state = 0) and DT1.isdropped then
    Playsound "TargetBankReset1",0,0.2,0.05,0.25
    DT1.isdropped = False
  end If
  if (DT2Light.state = 0) and DT2.isdropped then
    Playsound "TargetBankReset1",0,0.2,0.05,0.25
    DT2.isdropped = False
  end If
  if (DT3Light.state = 0) and DT3.isdropped then
    Playsound "TargetBankReset1",0,0.2,0.05,0.25
    DT3.isdropped = False
  end If
  if (DT4Light.state = 0) and DT4.isdropped then
    Playsound "TargetBankReset1",0,0.2,0.05,0.25
    DT4.isdropped = False
  end If
  if (DT5Light.state = 0) and DT5.isdropped then
    Playsound "TargetBankReset1",0,0.2,0.05,0.25
    DT5.isdropped = False
  end If
  if (DT6Light.state = 0) and DT6.isdropped then
    Playsound "TargetBankReset1",0,0.2,0.05,0.25
    DT6.isdropped = False
  end If
End Sub


'--- Left Ramp ---
Sub LeftRampMade_Hit
  FlashLeft 1
  AddPoints(2000000)
  DOF 138, DOFPulse

  IncreaseRampCount
  if (LeftRampLight.state <> 0) and (RightRampLight.state <> 0) then
    IncreaseRampCount
    IncreaseRampCount
    IncreaseRampCount
  else
    if (LeftRampLight.state <> 0) or (RightRampLight.state <> 0) then
      IncreaseRampCount
    end if
  end if
  LeftRampLight.timerenabled = False
  RightRampLight.timerenabled = False
  LeftRampLight.state = 0
  RightRampLight.state = 0

  if not hideanimation then
    if LockLight.state = 0 then
      UpdateDMDText "Main","",cstr(LeftRampCount) & " RAMPS","","LOCK IN " & cstr(5 - (LeftRampCount mod 5))
    else
      UpdateDMDText "Main","Ball "&Ballinplay,"","PLAYER 1",cstr(LeftRampCount) & " RAMPS"
    end if
  end if

  if combotimer.enabled and (LeftRampSpider.UserValue = 1) then
    AwardNextCombo
    ComboTimer.enabled = False
    ComboTimer.enabled = True
  else
    if Skillshot = 0 then
      actualcombocount = 0
      ComboTimer.enabled = True
      if CombomaniaActive then
        AwardnextCombo
      else
        if not missionactive then
          playsound "W_LeftRamp"
        end if
      end if
    end if
  end if

  if (StartMissionLight.State = 0) and not MissionActive and not UltimateShowdownActive then
    StartMissionLight.State = Lightstateblinking
  end if

  if BikeRaceActive then
    if LeftRampSpider.state <> 0 then
      AdvanceBikeRaceScore(False)
    end if
  end if

  if ShuttleActive then
    FlashLeft 3
    Advanceshuttlescore
  end if

  if SecretManiaActive then
    AwardSecretManiaScore
  end if
End Sub

Sub IncreaseRampCount
  LeftRampCount = LeftRampCount + 1
  if (LeftRampCount = 1) or (LeftRampCount mod 5 = 0) then
    if not FastFrenzyActive and not UltimateShowdownActive then
      if (LockLight.state = 0) and (LockLightstate = 0) then
        PlayNextSound "W_LockIsLit"
        if not Hideanimation then
          DMDDisplayScene "LockLit",True,LockLitback, "", 8,14,"", 5,15, 14, 3000, 1
          LockLight.state = 1
        else
          LockLightstate = 1
        end if
      end if
    end if
  end if

  if (LeftRampCount = 20) or (LeftRampCount mod 50 = 0) then
    if not HideAnimation then
      DMDDisplayScene "EBLit",True,EBLitback, "", 8,14,"", 5,15, 14, 2500, 1
    end if
    PlayNextSound "W_ExtraBallLit"
    EBLight.state = 1
  end if
End Sub

Sub LeftRampTarget_Hit
  Playsound "W_RampTarget2",0,1,-0.06
  DOF 119, DOFPulse
  if LeftRampLight.state = 0 then
    LeftRampLight.state = Lightstateblinking
    LeftRampLight.timerenabled = False
    LeftRampLight.timerenabled = True
  end if
End Sub

Sub LeftRampLight_Timer
  LeftRampLight.timerenabled = False
  LeftRampLight.state = 0
end Sub

Sub RightRampTarget_Hit
  Playsound "W_RampTarget2",0,1,-0.04
  DOF 119, DOFPulse
  if RightRampLight.state = 0 then
    RightRampLight.state = Lightstateblinking
    RightRampLight.timerenabled = False
    RightRampLight.timerenabled = True
  end if
End Sub

Sub RightRampLight_Timer
  RightRampLight.timerenabled = False
  RightRampLight.state = 0
end Sub

'--- Missions + Modes ---
Dim MissionCount,Mission7Active,EjectToRight
Sub MissionKicker_Hit()
  MissionKicker.DestroyBall
  Playsound "Drain3",0,0.3,0,0.25
  if EBLight.state <> 0 then
    if EBLight.Timerenabled then
      EBLight.Timerenabled = False
      LoopExtraBalls = LoopExtraBalls + 1
    end if
    EBLight.state = 0
    ExtraBalls = ExtraBalls + 1
    PlayNextSound "W_ExtraBall"
    DOF 123, DOFPulse
    if not hideanimation then
      DMDDisplayScene "EB",True,ExtraBallback, "", 8,14,"", 5,15, 14, 3900, 1
      MissionKicker.Timerenabled = True
    else
      ProcessMissionKicker
    end if
  else
    ProcessMissionKicker
  end if
End Sub

Sub MissionKicker_Timer
  MissionKicker.Timerenabled = False
  ProcessMissionKicker
End Sub

Sub ProcessMissionKicker
  if HuntDownActive and (MissionHolespider.state <> 0) then
    AdvanceHuntDownScore
  else
    if SkyscraperActive then
      if CurrentLevel > 0 then
        EndSkyscraper
      end if
    else
      if FastFrenzyActive or SpaceStationFrenzyActive or SecretManiaActive or UltimateShowdownActive then
        if FastFrenzyActive then
          if Jackpotlight.state = Lightstateblinking then
            AwardFFJackpot
          else
            UpdateDMDText "Main","","DROP TARGETS","","LIGHT  JACKPOT"
          end if
        end if
        if SpaceStationFrenzyActive  and (Jackpotlight.state = Lightstateblinking) then
          AwardSpaceStationJackpot
        end if
        if SecretManiaActive then
          AwardSecretManiaScore
        end if
        if UltimateShowdownActive then
          if USPhase3Active then
            EndUSPhase3
            Exit Sub
          else
            AwardUSPoints
          end if
        end if
      else
        if not QuickshotActive and not SuperLauncherActive and not ComboManiaActive and not MissionActive then
          if StartModeLight.state <> 0 then
            if QuickShotReady then
              StartQuickShot
            else
              if SuperLauncherReady then
                StartSuperLauncher
              else
                if ComboManiaReady then
                  StartComboMania
                else
                  if UltimateShowdownReady then
                    StartUltimateShowdown
                    Exit Sub
                  end if
                end if
              end if
            end if
          else
            if not MissionActive then
              if StartMissionLight.State <> 0 then
                StartMissionLight.State = 0
                if mission7active then
                  select case int(6*rnd)
                    case 0: playNextsound "W_Unbelievable"
                    case 1: playNextsound "W_Superb"
                    case 2: playNextsound "W_IamImpressed"
                    case 3: playNextsound "W_Outstanding"
                    case 4: playNextsound "W_VeryGood"
                    case 5: playNextsound "W_YesYesYes"
                  end select
                  UpdateDMDText "Main","","ALL MISSIONS","","COMPLETE"
                  Mission7AnimationTimer.enabled = False
                  US_MissionsLight.state = LightstateOn
                  ResetMissions
                  PlayNextSound "W_Sphere"
                  ShowdownSpheres = ShowdownSpheres + 1
                  addpoints(250000000*MissionCount)
                  DMDDisplayScene "Sphere",True,Sphereback, "", 8,14,FormatPointText(cstr(250000000*MissionCount)), 5,15, 14, 3000, 1
                  MissionCount = MissionCount + 1
                else
                  PlayNextSound "W_Mission"
                  select case M
                    case 0: StartReactorCritical
                    case 1: StartHuntDown
                    case 2: StartBikeRace
                    case 3: StartAmmoDump
                    case 4: StartSkyscraper
                    case 5: StartShuttle
                  end select
                end if
              else
                UpdateDMDText "Main","","SHOOT RAMPS TO","","LIGHT  MISSION"
              end if
            end if
          end if
        end if
      end if
    end if
  end if

  if not QuickShotActive then
    EjectToRight = not EjectToRight
  end if

  if SuperLauncherActive then
    EjectToRight = False
  end if

  if ReactorCriticalActive then
    EjectToRight = False
  end if

  if AmmoDumpActive then
    EjectToRight = True
  end if

  if SkyscraperActive then
    EjectToRight = False
  end if

  if EjectToRight then
    TroughRightMissionKicker.CreateBall
    TroughRightMissionKicker.kick 180, 1
  else
    TroughLeftMissionKicker.CreateBall
    TroughLeftMissionKicker.kick 180, 1
  end if
End Sub

'--- Missions ---
Dim MissionActive,MissionTotal
Sub MysteryTarget_Hit
  Playsound SoundFXDOF("W_RightTarget",133,DOFPulse,DOFContactors),0,1,-0.03,0.25
  if not MissionActive then
    AdvanceMission
  end if
  if ReactorCriticalActive then
    ReactorMystery
  end if
  if HuntDownActive then
    HuntDownMystery
  end if
  if AmmoDumpActive then
    AmmoDumpMystery
  end if
  if ShuttleActive then
    ShuttleMystery
  end if
  if SkyscraperActive then
    SkyscraperMystery
  end if
  if BikeRaceActive then
    BikeRaceMystery
  end if
  if UltimateShowdownActive then
    UltimateShowdownMystery
  end if
End Sub

Dim MissionLights,MissionsComplete,M,M_Reactor,M_HuntDown,M_BikeRace,M_AmmoDump,M_SkyScraper,M_Shuttle
Randomize
M = Int(5*rnd)
AnimM = M
M_Reactor = 0
M_HuntDown = 0
M_BikeRace = 0
M_AmmoDump = 0
M_SkyScraper = 0
M_Shuttle = 0'1
MissionLights = Array(MissionLightReactor,MissionLightHuntDown,MissionLightBikeRace,MissionLightAmmoDump,MissionLightSkyScraper,MissionLightShuttle)
MissionsComplete = Array(M_Reactor,M_HuntDown,M_BikeRace,M_AmmoDump,M_SkyScraper,M_Shuttle)
Sub AdvanceMission
  if MissionsComplete(0) and MissionsComplete(1) and MissionsComplete(2) and MissionsComplete(3) and MissionsComplete(4) and MissionsComplete(5) then
    mission7active = True
    for each obj in MissionLights
      obj.state = 0
    next
    AnimM = M
    Mission7AnimationTimer.enabled = True
  else
    if MissionsComplete(M) = 0 then
      MissionLights(M).state = 0
    end if
    M = M + 1
    if M > 5 then
      M = 0
    end if
    for i = 1 to 5
      if MissionsComplete(M) = 1 then
        M = M + 1
        if M > 5 then
          M = 0
        end if
      end if
    next
    if MissionsComplete(M) = 0 then
      MissionLights(M).state = Lightstateblinking
    end if
  end if
End Sub

Sub ResetMissions
  Mission7Active = False
  Mission7AnimationTimer.enabled = False
  for each obj in MissionLights
    obj.state = 0
  next
  for i = 0 to 5
    MissionsComplete(i) = 0
  next

  select case AnimM
    case 0: M = 0
    case 1: M = 1
    case 2: M = 3
    case 3: M = 4
    case 4: M = 5
    case 5: M = 2
  end select
  MissionLights(M).state = Lightstateblinking
End Sub

'--- Mission 7 ---
Dim AnimM
Sub Mission7AnimationTimer_Timer
  'Sequence: 0 - 1 - 3 - 4 - 5 - 2
  for each obj in MissionLights
    obj.state = 0
  next

  AnimM= AnimM + 1
  if AnimM > 5 then
    AnimM = 0
  end if
  select case AnimM
    case 0: M = 0
    case 1: M = 1
    case 2: M = 3
    case 3: M = 4
    case 4: M = 5
    case 5: M = 2
  end select
  MissionLights(M).state = Lightstateon
End Sub

'--- Reactor ---
Dim ReactorCriticalActive,HuntDownActive,BikeRaceActive,AmmoDumpActive,SkyscraperActive,ShuttleActive
Dim ReactorCompletion,ReactorMagnetwasLit,ReactorTimeCount
Sub StartReactorCritical
  SkillshotTimer.enabled = False
  SkillShot = 0
  MissionActive = True
  ReactorCriticalActive = True
  MissionTotal = 0
  DTDropUp 0
  DT1Light.state = Lightstateblinking
  DT2Light.state = Lightstateblinking
  DT3Light.state = Lightstateblinking
  DT4Light.state = Lightstateblinking
  DT5Light.state = Lightstateblinking
  DT6Light.state = Lightstateblinking
  OverloadLight.state = Lightstateblinking
  MysteryLight.state = Lightstateblinking
  ReactorHoleSpider.state = 0
  ReactorMagnetWasLit = (LightMagnetLight.state = 1)
  ExplosionLightState = ExplosionLight.State
  LockLightState = LockLight.state
  if LockLight.state <> 0 and BallsLocked = 0 then
    LockLight.state = 0
  end if
  ExplosionLight.State = 0
  LightMagnetLight.state = 0
  ReactorCompletion = 0
  ReactorTimeCount = 51
  DMDStartMission "Reactor",51
  ReactorTimer.enabled = True
  PlayNextMusic(ReactorTheme)
End Sub

Sub ReactorMystery
  if MysteryLight.state <> 0 then
    if int(2*rnd) = 1 then
      PlayNextSound "W_Mystery"
    else
      PlayNextSound "W_Mystery_F"
    end if
  end if
  OverloadLight.state = 0
  MysteryLight.state = 0
End Sub

Sub ReactorTimer_Timer
  ReactorTimeCount = ReactorTimeCount - 1
  UpdateDMDMission ReactorTimeCount,cstr(ReactorCompletion)
  if ReactorTimeCount <= 0 then
    ReactorTimer.enabled = False
    EndReactorCritical
  end if
End Sub

Sub AddReactorScore
  if ReactorCompletion < 99 then
    addpoints(5000000 * MissionCount)
    MissionTotal = MissionTotal + (5000000 * MissionCount)
    Reactorcompletion = Reactorcompletion + 10
  end if
  if ReactorCompletion >= 100 then
    ReactorCompletion = 99
    DT1Light.state = 0
    DT2Light.state = 0
    DT3Light.state = 0
    DT4Light.state = 0
    DT5Light.state = 0
    DT6Light.state = 0
    ReactorHoleSpider.state = Lightstateblinking
    OverloadLight.state = 0
    MysteryLight.state = 0
  end if
  UpdateDMDMission ReactorTimeCount,cstr(ReactorCompletion)
End Sub

Sub EndReactorCritical
  ReactorTimer.enabled = False
  MissionsComplete(0) = 1
  MissionLights(0).state = 1
  MissionActive = False
  ReactorCriticalActive = False
  if ReactorCompletion >= 100 then
    FlashRight 2
    FlashLeft 2
    addpoints(50000000 * MissionCount)
    MissionTotal = MissionTotal + (50000000 * MissionCount)
    Showdownspheres = Showdownspheres + 1
    PlayNextSound "W_Sphere"
  else
    Playsound "W_Explosion"
  end if
  DMDEndMission MissionTotal
  OverloadLight.state = 0
  MysteryLight.state = 0
  ReactorHoleSpider.state = 0
  DTDropUp 0
  DT1Light.state = 0
  DT2Light.state = 0
  DT3Light.state = 0
  DT4Light.state = 0
  DT5Light.state = 0
  DT6Light.state = 0
  if ReactorMagnetWasLit then
    LightMagnetLight.state = 1
  end if
  ExplosionLight.State = ExplosionLightState
  LockLight.State = LockLightstate
  LockLightstate = 0
  AdvanceMission
  PlaynextMusic(MainTheme)
End Sub

'--- HuntDown ---
Dim HuntDownCompletion,HuntDownTimeCount
Sub StartHuntDown
  SkillshotTimer.enabled = False
  SkillShot = 0
  MissionActive = True
  HuntDownActive = True
  HuntDownCompletion = 0
  MissionTotal = 0
  HuntDownTimeCount = 64
  DMDStartMission "HuntDown",64
  MysteryLight.state = Lightstateblinking
  ActiveHole = int(4*rnd)
  ReactorHoleSpider.state = 0
  LeftHoleSpider.state = 0
  MissionHoleSpider.state = 0
  ExplosionHoleSpider.state = 0
  LeftOrbitSpider.state = 0
  RightOrbitSpider.state = 0
  LeftRampSpider.state = 0
  LoopSpider.state = 0
  DockRampSpider.state = 0
  ExplosionLightState = ExplosionLight.State
  ExplosionLight.State = 0
  LockLightState = LockLight.state
  if LockLight.state <> 0 and BallsLocked = 0 then
    LockLight.state = 0
  end if

  select case ActiveHole
    case 0: ReactorHoleSpider.state = Lightstateblinking
    case 1: LeftHoleSpider.state = Lightstateblinking
    case 2: MissionHoleSpider.state = Lightstateblinking
    case 3: ExplosionHoleSpider.state = Lightstateblinking
  end select
  HuntDownTimer.enabled = True
  PlayNextMusic(HuntDownTheme)
End Sub

Sub HuntDownTimer_Timer
  HuntDownTimeCount = HuntDownTimeCount - 1
  UpdateDMDMission cstr(HuntDownTimeCount),cstr(HuntDownCompletion)
  if HuntDownTimeCount <= 0 then
    HuntDownTimer.enabled = False
    EndHuntDown
  end if
End Sub

Sub HuntDownMystery
  if HuntDownCompletion < 99 then
    if MysteryLight.state <> 0 then
      if int(2*rnd) = 1 then
        PlayNextSound "W_Mystery"
      else
        PlayNextSound "W_Mystery_F"
      end if
    end if
    MysteryLight.state = 0
    ReactorHoleSpider.state = Lightstateblinking
    LeftHoleSpider.state = Lightstateblinking
    MissionHoleSpider.state = Lightstateblinking
    ExplosionHoleSpider.state = Lightstateblinking
  else
    HuntDownCompletion = 100
    EndHuntDown
  end if
End Sub

Dim ActiveHole
Sub AdvanceHuntDownScore
  playsound "W_BikeHit"
  if HuntDownCompletion < 99 then
    HuntDownCompletion = HuntDownCompletion + 20
    addpoints(10000000 * MissionCount)
    MissionTotal = MissionTotal + (10000000 * MissionCount)
  end if
  if HuntDownCompletion >= 100 then
    HuntDownCompletion = 99
    ReactorHoleSpider.state = 0
    LeftHoleSpider.state = 0
    MissionHoleSpider.state = 0
    ExplosionHoleSpider.state = 0
    MysteryLight.state = Lightstateblinking
  else
    for i = 1 to int(3*rnd+1)
      ActiveHole = ActiveHole + 1
      if ActiveHole > 3 then
        ActiveHole = 0
      end if
    next
    ReactorHoleSpider.state = 0
    LeftHoleSpider.state = 0
    MissionHoleSpider.state = 0
    ExplosionHoleSpider.state = 0
    MysteryLight.state = Lightstateblinking
    select case ActiveHole
      case 0: ReactorHoleSpider.state = Lightstateblinking
      case 1: LeftHoleSpider.state = Lightstateblinking
      case 2: MissionHoleSpider.state = Lightstateblinking
      case 3: ExplosionHoleSpider.state = Lightstateblinking
    end select
  end if
  UpdateDMDMission cstr(HuntDownTimeCount),cstr(HuntDownCompletion)
End Sub

Sub EndHuntDown
  HuntDownTimer.enabled = False
  MissionActive = False
  HuntDownActive = False
  MissionsComplete(1) = 1
  MissionLights(1).state = 1
  if HuntDownCompletion >= 100 then
    FlashRight 2
    FlashLeft 2
    addpoints(100000000 * MissionCount)
    MissionTotal = MissionTotal + (100000000 * MissionCount)
    Showdownspheres = Showdownspheres + 1
    PlayNextSound "W_Sphere"
  end if
  DMDEndMission MissionTotal
  MysteryLight.state = 0
  ReactorHoleSpider.state = 0
  LeftHoleSpider.state = 0
  MissionHoleSpider.state = 0
  ExplosionHoleSpider.state = 0
  ExplosionLight.State = ExplosionLightState
  LockLight.State = LockLightstate
  LockLightstate = 0
  AdvanceMission
  PlaynextMusic(MainTheme)
End Sub

'--- Ammo Dump ---
Dim AmmoDumpCompletion,AmmoDumpTimeCount,AmmoDumpMysteryBoost
Sub StartAmmoDump
  SkillshotTimer.enabled = False
  SkillShot = 0
  MissionActive = True
  AmmoDumpActive = True
  ExplosionLightState = ExplosionLight.State
  ExplosionLight.State = 0
  LockLightState = LockLight.state
  if LockLight.state <> 0 and BallsLocked = 0 then
    LockLight.state = 0
  end if
  AmmoDumpCompletion = 0
  AmmoDumpTimeCount = 53
  MissionTotal = 0
  DMDStartMission "AmmoDump",53
  AmmoDumpMysteryBoost = 1
  MysteryLight.state = Lightstateblinking
  BumperLaneLight.state = Lightstateblinking
  AmmoDumpTimer.enabled = True
  PlayNextMusic(AmmoDumpTheme)
End Sub

Sub AmmoDumpTimer_Timer
  AmmoDumpTimeCount = AmmoDumpTimeCount - 1
  UpdateDMDMission cstr(AmmoDumpTimeCount),cstr(AmmoDumpCompletion)
  if AmmoDumpTimeCount <= 0 then
    AmmoDumpTimer.enabled = False
    EndAmmoDump
  end if
End Sub

Sub AmmoDumpMystery
  if MysteryLight.state <> 0 then
    if int(2*rnd) = 1 then
      PlayNextSound "W_Mystery"
    else
      PlayNextSound "W_Mystery_F"
    end if
  end if
  AmmoDumpMysteryBoost = 2
  MysteryLight.state = 0
End Sub

Sub AdvanceAmmoDumpScore
  AmmoDumpCompletion = AmmoDumpCompletion + (4 * AmmoDumpMysteryBoost)
  addpoints(2000000 * MissionCount)
  MissionTotal = MissionTotal + (2000000 * MissionCount)
  if AmmoDumpCompletion >= 100 then
    EndAmmoDump
  else
    UpdateDMDMission cstr(AmmoDumpTimeCount),cstr(AmmoDumpCompletion)
  end if
End Sub

Sub EndAmmoDump
  AmmoDumpTimer.enabled = False
  MissionActive = False
  AmmoDumpActive = False
  MissionsComplete(3) = 1
  MissionLights(3).state = 1
  if AmmoDumpCompletion >= 100 then
    FlashRight 2
    FlashLeft 2
    addpoints(45000000 * MissionCount)
    MissionTotal = MissionTotal + (45000000 * MissionCount)
    Showdownspheres = Showdownspheres + 1
    PlayNextSound "W_Sphere"
  else
    playsound "W_Explosion"
  end if
  DMDEndMission MissionTotal
  MysteryLight.state = 0
  BumperLaneLight.state = 0
  ExplosionLight.State = ExplosionLightState
  LockLight.State = LockLightstate
  LockLightstate = 0
  AdvanceMission
  PlaynextMusic(MainTheme)
End Sub

'--- Shuttle ---
Dim ShuttleCompletion,ShuttleTimeCount,ShuttleRampValue,ShuttleRampBaseValue
Sub StartShuttle
  SkillshotTimer.enabled = False
  SkillShot = 0
  MissionActive = True
  ShuttleActive = True
  ShuttleCompletion = 0
  ShuttleTimeCount = 50
  MissionTotal = 0
  DMDStartMission "Shuttle",50
  MysteryLight.state = 0
  ShuttleRampBaseValue = 0
  ShuttleRampValue = 14
  LeftOrbitSpider.state = 0
  LoopSpider.state = 0
  RightOrbitSpider.state = 0
  LeftRampSpider.state = Lightstateblinking
  DockRampSpider.state = Lightstateblinking
  ExplosionLightState = ExplosionLight.State
  ExplosionLight.State = 0
  LockLightState = LockLight.state
  if LockLight.state <> 0 and BallsLocked = 0 then
    LockLight.state = 0
  end if
  ShuttleTimer.enabled = True
  PlayNextMusic(ShuttleTheme)
End Sub

Sub ShuttleTimer_Timer
  ShuttleTimeCount = ShuttleTimeCount - 1
  UpdateDMDMission cstr(ShuttleTimeCount),cstr(ShuttleCompletion)
  if ShuttleTimeCount <= 0 then
    ShuttleTimer.enabled = False
    EndShuttle
  end if
End Sub

Sub ShuttleMystery
  if MysteryLight.state <> 0  then
    MysteryLight.state = 0
    ShuttleCompletion = ShuttleCompletion + 25
    FlashLeft 3
    FlashRight 3
    if ShuttleCompletion >= 100 then
      EndShuttle
    else
      Playsound "W_MachineGun"
      UpdateDMDMission cstr(ShuttleTimeCount),cstr(ShuttleCompletion)
    end if
  end if
End Sub

Sub AdvanceShuttleRampValue
  if ShuttleActive then
    ShuttleRampValue = ShuttleRampValue + 3
    if ShuttleRampValue > 37 then
      ShuttleRampValue = 37
    end if
  end if
End Sub

Sub AdvanceShuttleScore
  MysteryLight.state = Lightstateblinking
  ShuttleCompletion = ShuttleCompletion + ShuttleRampValue + ShuttleRampBaseValue
  addpoints((ShuttleRampValue + ShuttleRampBaseValue) * 500000 * MissionCount)
  MissionTotal = MissionTotal + ((ShuttleRampValue + ShuttleRampBaseValue) * 500000 * MissionCount)
  if ShuttleRampBaseValue = 0 then
    ShuttleRampBaseValue = 1
  else
    ShuttleRampBaseValue = 0
  end if
  ShuttleRampValue = 14
  if ShuttleCompletion >= 100 then
    EndShuttle
  else
    Playsound "W_MachineGun"
    UpdateDMDMission cstr(ShuttleTimeCount),cstr(ShuttleCompletion)
  end if
End Sub

Sub EndShuttle
  ShuttleTimer.enabled = False
  MissionActive = False
  ShuttleActive = False
  MissionsComplete(5) = 1
  MissionLights(5).state = 1
  if ShuttleCompletion >= 100 then
    FlashRight 2
    FlashLeft 2
    addpoints(55000000 * MissionCount)
    MissionTotal = MissionTotal + (55000000 * MissionCount)
    Showdownspheres = Showdownspheres + 1
    PlayNextSound "W_ShuttleDestroyed"
  end if
  DMDEndMission MissionTotal
  MysteryLight.state = 0
  LeftRampSpider.state = 0
  DockRampSpider.state = 0
  ExplosionLight.State = ExplosionLightState
  LockLight.State = LockLightstate
  LockLightstate = 0
  AdvanceMission
  PlaynextMusic(MainTheme)
End Sub

'--- Bike Race ---
Dim BikeRaceCompletion,BikeRaceTimeCount,CurrentLitLane
Sub StartBikeRace
  SkillshotTimer.enabled = False
  SkillShot = 0
  MissionActive = True
  BikeRaceActive = True
  BikeRaceCompletion = 0
  BikeRaceTimeCount = 67
  MissionTotal = 0
  DMDStartMission "BikeRace",67
  MysteryLight.state = LightstateBlinking
  LeftOrbitSpider.state = 0
  LeftRampSpider.state = 0
  DockRampSpider.state = 0
  RightOrbitSpider.state = 0
  LoopSpider.state = Lightstateblinking
  ExplosionLightState = ExplosionLight.State
  ExplosionLight.State = 0
  LockLightState = LockLight.state
  if LockLight.state <> 0 and BallsLocked = 0 then
    LockLight.state = 0
  end if
  CurrentLitLane = 1
  select case CurrentLitLane
    case 0: LeftOrbitSpider.state = Lightstateblinking
    case 1: LeftRampSpider.state = Lightstateblinking
    case 2: DockRampSpider.state = Lightstateblinking
    case 3: RightOrbitSpider.state = Lightstateblinking
  end select
  BikeRaceTimer.enabled = True
  BikeRaceSpiderTimer1.enabled = True
  BikeRaceSpiderTimer2.enabled = True
  PlayNextMusic(BikeRaceTheme)
End Sub

Sub BikeRaceSpiderTimer1_Timer
  BikeRaceSpiderTimer1.enabled = False
  select case CurrentLitLane
    case 0: LeftOrbitSpider.Blinkinterval = 65
    case 1: LeftRampSpider.Blinkinterval = 65
    case 2: DockRampSpider.Blinkinterval = 65
    case 3: RightOrbitSpider.Blinkinterval = 65
  end select
End Sub

Sub BikeRaceSpiderTimer2_Timer
  BikeRaceSpiderTimer2.enabled = False
  select case CurrentLitLane
    case 0: LeftOrbitSpider.Blinkinterval = 125
    case 1: LeftRampSpider.Blinkinterval = 125
    case 2: DockRampSpider.Blinkinterval = 125
    case 3: RightOrbitSpider.Blinkinterval = 125
  end select
  CurrentLitLane = CurrentLitLane + 1
  if CurrentLitLane > 3 then
    CurrentLitLane = 0
  end if
  LeftOrbitSpider.state = 0
  LeftRampSpider.state = 0
  DockRampSpider.state = 0
  RightOrbitSpider.state = 0
  select case CurrentLitLane
    case 0: LeftOrbitSpider.state = Lightstateblinking
    case 1: LeftRampSpider.state = Lightstateblinking
    case 2: DockRampSpider.state = Lightstateblinking
    case 3: RightOrbitSpider.state = Lightstateblinking
  end select
  BikeRaceSpiderTimer1.enabled = True
  BikeRaceSpiderTimer2.enabled = True
End Sub

Sub BikeRaceTimer_Timer
  BikeRaceTimeCount = BikeRaceTimeCount - 1
  UpdateDMDMission cstr(BikeRaceTimeCount),cstr(BikeRaceCompletion)
  if BikeRaceTimeCount <= 0 then
    BikeRaceTimer.enabled = False
    EndBikeRace
  end if
End Sub

Sub BikeRaceMystery
  if MysteryLight.state <> 0 then
    if int(2*rnd) = 1 then
      PlayNextSound "W_Mystery"
    else
      PlayNextSound "W_Mystery_F"
    end if
  end if
  BikeRaceSpiderTimer1.enabled = False
  BikeRaceSpiderTimer2.enabled = False
  MysteryLight.state = 0
End Sub

Sub AdvanceBikeRaceScore(LoopPar)
  PlaySound "W_BikeHit"
  if LoopPar then
    LoopSpider.state = 0
  else
    LoopSpider.state = Lightstateblinking
  end if
  MysteryLight.state = Lightstateblinking
  BikeRaceSpiderTimer1.enabled = True
  BikeRaceSpiderTimer2.enabled = True
  CurrentLitLane = CurrentLitLane + 1
  if CurrentLitLane > 3 then
    CurrentLitLane = 0
  end if
  LeftOrbitSpider.state = 0
  LeftRampSpider.state = 0
  DockRampSpider.state = 0
  RightOrbitSpider.state = 0
  select case CurrentLitLane
    case 0: LeftOrbitSpider.state = Lightstateblinking
    case 1: LeftRampSpider.state = Lightstateblinking
    case 2: DockRampSpider.state = Lightstateblinking
    case 3: RightOrbitSpider.state = Lightstateblinking
  end select
  addpoints(10000000 * MissionCount)
  MissionTotal = MissionTotal + (10000000 * MissionCount)
  BikeRaceCompletion = BikeRaceCompletion + 25
  if BikeRaceCompletion >= 100 then
    EndBikeRace
  else
    UpdateDMDMission cstr(BikeRaceTimeCount),cstr(BikeRaceCompletion)
  end if
End Sub

Sub EndBikeRace
  BikeRaceTimer.enabled = False
  BikeRaceSpiderTimer1.enabled = False
  BikeRaceSpiderTimer2.enabled = False
  MissionActive = False
  BikeRaceActive = False
  MissionsComplete(2) = 1
  MissionLights(2).state = 1
  if BikeRaceCompletion >= 100 then
    FlashRight 2
    FlashLeft 2
    addpoints(60000000 * MissionCount)
    MissionTotal = MissionTotal + (60000000 * MissionCount)
    Showdownspheres = Showdownspheres + 1
    PlayNextSound "W_Sphere"
  end if
  DMDEndMission MissionTotal
  MysteryLight.state = 0
  LeftOrbitSpider.state = 0
  LeftRampSpider.state = 0
  DockRampSpider.state = 0
  RightOrbitSpider.state = 0
  LoopSpider.state = 0
  ExplosionLight.State = ExplosionLightState
  LockLight.State = LockLightstate
  LockLightstate = 0
  AdvanceMission
  PlaynextMusic(MainTheme)
End Sub

'--- Skyscraper ---
Dim SkyscraperCompletion,SkyscraperTimeCount,MaxLevel,CurrentLevel
Sub StartSkyscraper
  SkillshotTimer.enabled = False
  SkillShot = 0
  MissionActive = True
  SkyscraperActive = True
  SkyscraperCompletion = 0
  SkyscraperTimeCount = 62
  MissionTotal = 0
  DMDStartMission "Skyscraper",62
  MaxLevel = 5
  CurrentLevel = 0
  MysteryLight.state = 0
  LeftOrbitSpider.state = 0
  LeftRampSpider.state = 0
  DockRampSpider.state = 0
  RightOrbitSpider.state = 0
  LoopSpider.state = 0
  MissionHoleSpider.state = 0
  ExplosionLightState = ExplosionLight.State
  ExplosionLight.State = 0
  LockLightState = LockLight.state
  if LockLight.state <> 0 and BallsLocked = 0 then
    LockLight.state = 0
  end if
  SkyscraperTimer.enabled = True
  S_State = 0
  S_Step = 0
  SkyscraperAnimationTimer.enabled = True
  PlayNextMusic(SkyscraperTheme)
End Sub

Dim S_Step,S_State
Sub SkyscraperAnimationTimer_Timer
  S_Step = S_Step + 1
  if S_Step > 6 then
    S_Step = 1
    if S_State = 1 then
      S_State = 0
    else
      S_State = 1
    end if
  end if
  select case S_Step
    case 1: PowerLevel1Light.state = S_State
        PowerLevel2Light.state = S_State
    case 2: PowerLevel3Light.state = S_State
        PowerLevel4Light.state = S_State
    case 3: PowerLevel5Light.state = S_State
        PowerLevel6Light.state = S_State
    case 4: PowerLevel7Light.state = S_State
        PowerLevel8Light.state = S_State
    case 5: AdvanceLeftPowerlevelLight.state = S_State
        AdvanceRightPowerlevelLight.state = S_State
    case 6: LeftOrbitSpider.state = S_State
        RightOrbitSpider.state = S_State
  end select
End Sub

Sub SkyscraperTimer_Timer
  SkyscraperTimeCount = SkyscraperTimeCount - 1
  UpdateDMDMission cstr(SkyscraperTimeCount),cstr(SkyscraperCompletion)
  if SkyscraperTimeCount <= 0 then
    SkyscraperTimer.enabled = False
    EndSkyscraper
  end if
End Sub

Sub SkyscraperMystery
  if MysteryLight.state <> 0 then
    if int(2*rnd) = 1 then
      PlayNextSound "W_Mystery"
    else
      PlayNextSound "W_Mystery_F"
    end if
    Maxlevel = 6
    MysteryLight.state = 0
  end if
End Sub

Sub SkyscraperPauseTimer_Timer
  SkyscraperPauseTimer.enabled = False
End sub

Sub AdvanceSkyscraperScore
  if SkyscraperActive then
    if not SkyscraperPauseTimer.enabled then
      if CurrentLevel < MaxLevel then
        CurrentLevel = CurrentLevel + 1
        select case CurrentLevel
          case 1: PlayNextSound "W_Floor1"
          case 2: PlayNextSound "W_Floor2"
          case 3: PlayNextSound "W_Floor3"
          case 4: PlayNextSound "W_Floor4"
          case 5: PlayNextSound "W_Floor5"
        end select
        MissionHoleSpider.state = Lightstateblinking
        if CurrentLevel = 5 then
          MysteryLight.state = Lightstateblinking
        end if
        addpoints(1000000 * MissionCount)
        MissionTotal = MissionTotal + (1000000 * MissionCount)
        SkyscraperCompletion = SkyscraperCompletion + 20
      end if
      UpdateDMDMission cstr(SkyscraperTimeCount),cstr(SkyscraperCompletion)
      SkyscraperPauseTimer.enabled = False
      SkyscraperPauseTimer.enabled = True
    end if
  end if
End Sub

Sub EndSkyscraper
  SkyscraperTimer.enabled = False
  SkyscraperAnimationTimer.enabled = False
  MissionActive = False
  SkyscraperActive = False
  MissionsComplete(4) = 1
  MissionLights(4).state = 1
  if SkyscraperCompletion = 100 then
    FlashRight 2
    FlashLeft 2
    addpoints(100000000 * MissionCount)
    MissionTotal = MissionTotal + (100000000 * MissionCount)
    Showdownspheres = Showdownspheres + 1
    PlayNextSound "W_Sphere"
  end if
  if SkyscraperCompletion > 100 then
    FlashRight 3
    FlashLeft 3
    addpoints(200000000 * MissionCount)
    MissionTotal = MissionTotal + (200000000 * MissionCount)
    ExtraBalls = ExtraBalls + 1
    Showdownspheres = Showdownspheres + 1
    PlayNextSound "W_Sphere"
  end if
  DMDEndMission MissionTotal
  MysteryLight.state = 0
  MissionHoleSpider.state = 0
  PowerLevel1Light.state = abs(Powerlevel > 1)
  PowerLevel2Light.state = abs(Powerlevel > 2)
  PowerLevel3Light.state = abs(Powerlevel > 3)
  PowerLevel4Light.state = abs(Powerlevel > 4)
  PowerLevel5Light.state = abs(Powerlevel > 5)
  PowerLevel6Light.state = abs(Powerlevel > 6)
  PowerLevel7Light.state = abs(Powerlevel > 7)
  PowerLevel8Light.state = abs(Powerlevel > 8)
  select case Powerlevel
    case 1: PowerLevel1Light.state = LightstateBlinking
    case 2: PowerLevel2Light.state = LightstateBlinking
    case 3: PowerLevel3Light.state = LightstateBlinking
    case 4: PowerLevel4Light.state = LightstateBlinking
    case 5: PowerLevel5Light.state = LightstateBlinking
    case 6: PowerLevel6Light.state = LightstateBlinking
    case 7: PowerLevel7Light.state = LightstateBlinking
    case 8: PowerLevel8Light.state = LightstateBlinking
  end select
  AdvanceLeftPowerlevelLight.state = 0
  LeftOrbitSpider.state = 0
  AdvanceRightPowerlevelLight.state = 0
  RightOrbitSpider.state = 0
  ExplosionLight.State = ExplosionLightState
  LockLight.State = LockLightstate
  LockLightstate = 0
  AdvanceMission
  PlaynextMusic(MainTheme)
End Sub


'--- Quickshot ---
Dim QuickShotActive,SuperLauncherActive,ComboManiaActive,UltimateShowdownActive
Dim QuickShotBaseValue,ActualQuickShotValue
Sub StartQuickShot
  ExplosionLightState = ExplosionLight.State
  ExplosionLight.State = 0
  PlayNextSound "W_Quickshot"
  QuickShotReady = False
  QuickShotActive = True
  PowerLevel1Light.state = 0
  PowerLevel3Light.state = 0
  PowerLevel5Light.state = 0
  PowerLevel7Light.state = 0
  AdvanceLeftPowerlevelLight.state = 0
  LeftOrbitSpider.state = 0
  PowerLevel2Light.state = 0
  PowerLevel4Light.state = 0
  PowerLevel6Light.state = 0
  PowerLevel8Light.state = 0
  AdvanceRightPowerlevelLight.state = 0
  RightOrbitSpider.state = 0
  QuickShotBaseValue = QuickShotBaseValue + 25000000
  ActualQuickShotValue = QuickShotBaseValue
  QuickshotStartTimer.enabled = True
  QSState = 1
  Step = 0
  QuickShotAnimationTimer.enabled = True
  PlayNextMusic(QuickshotTheme)
End Sub

Sub QuickshotStartTimer_Timer
  QuickshotStartTimer.enabled = False
  QSLoop = 0
  QuickshotTimer.enabled = True
End Sub

Dim QSLoop
Sub QuickshotTimer_Timer
  QSLoop = QSLoop + 1
  ActualQuickShotValue = ActualQuickShotValue - (11870 * (QuickShotBaseValue/25000000))   'Timer runs 22 sec
  if not FrenzyPickTimer.enabled then
    if ActualQuickShotValue > 5000000 then
      UpdateDMDText "Main",""," SHOOT ORBIT FOR","",FormatPointText(cstr(ActualQuickShotValue))
    else
      UpdateDMDText "Main",""," SHOOT ORBIT FOR","",FormatPointText(cstr(5000000))
    end if
  end if
  if QSLoop = 600 then
    PlayNextSound "W_Quick"
  End if
  if QSLoop = 1200 then
    PlayNextSound "W_Timerunsout"
  End if
  if ActualQuickShotValue <= 0 then
    QuickshotTimer.enabled = False
    PlayNextSound "W_TooSlow"
    EndQuickShot
  end if
End Sub

Dim Step,QSState
Sub QuickShotAnimationTimer_Timer
  Step = Step + 1
  if Step > 6 then
    Step = 1
    if QSState = 1 then
      QSState = 0
    else
      QSState = 1
    end if
  end if
  if EjectToRight then    'init left
    select case Step
      case 1: PowerLevel1Light.state = QSState
      case 2: PowerLevel3Light.state = QSState
      case 3: PowerLevel5Light.state = QSState
      case 4: PowerLevel7Light.state = QSState
      case 5: AdvanceLeftPowerlevelLight.state = QSState
      case 6: LeftOrbitSpider.state = QSState
    end select
  else
    select case Step
      case 1: PowerLevel2Light.state = QSState
      case 2: PowerLevel4Light.state = QSState
      case 3: PowerLevel6Light.state = QSState
      case 4: PowerLevel8Light.state = QSState
      case 5: AdvanceRightPowerlevelLight.state = QSState
      case 6: RightOrbitSpider.state = QSState
    end select
  end if
End Sub

Sub EndQuickShot
  if QuickShotActive then         'for Ultimate Showdown and Multiballs
    if ActualQuickShotValue > 0 then
      FlashRight 2
      FlashLeft 2
      FlashB2SLamp 51,1
      PlayNextSound "W_Superb"
      if ActualQuickShotValue < 5000000 then
        ActualQuickShotValue = 5000000
      end if
      addpoints(ActualQuickShotValue)
      DMDEndQuickshot ActualQuickShotValue
      QuickShotBaseValue = QuickShotBaseValue + 25000000
    else
      if not FrenzyPickTimer.enabled then
        DMDEndMode
      end if
    end if
  end if
  QuickShotActive = False
  QuickShotTimer.enabled = False
  QuickShotAnimationTimer.enabled = False
  PowerLevel1Light.state = abs(Powerlevel > 1)
  PowerLevel2Light.state = abs(Powerlevel > 2)
  PowerLevel3Light.state = abs(Powerlevel > 3)
  PowerLevel4Light.state = abs(Powerlevel > 4)
  PowerLevel5Light.state = abs(Powerlevel > 5)
  PowerLevel6Light.state = abs(Powerlevel > 6)
  PowerLevel7Light.state = abs(Powerlevel > 7)
  PowerLevel8Light.state = abs(Powerlevel > 8)
  select case Powerlevel
    case 1: PowerLevel1Light.state = LightstateBlinking
    case 2: PowerLevel2Light.state = LightstateBlinking
    case 3: PowerLevel3Light.state = LightstateBlinking
    case 4: PowerLevel4Light.state = LightstateBlinking
    case 5: PowerLevel5Light.state = LightstateBlinking
    case 6: PowerLevel6Light.state = LightstateBlinking
    case 7: PowerLevel7Light.state = LightstateBlinking
    case 8: PowerLevel8Light.state = LightstateBlinking
  end select
  AdvanceLeftPowerlevelLight.state = 0
  LeftOrbitSpider.state = 0
  AdvanceRightPowerlevelLight.state = 0
  RightOrbitSpider.state = 0
  ExplosionLight.State = ExplosionLightState
  PlayNextMusic(MainTheme)
End Sub

'--- Super Launcher ---
Dim MagnetWasLit,SuperLauncherBaseValue,SuperLauncherAddValue,ActualSuperLauncherValue
Sub StartSuperLauncher
  SuperLauncherReady = False
  SuperLauncherActive = True
  MagnetWasLit = (ActivateMagnetLight.state = 1)
  StartModeLightState = StartModeLight.State
  StartMissionLightState = StartMissionLight.State
  LockLightState = LockLight.State
  VideomodeLightState = VideomodeLight.State
  ExplosionLightState = ExplosionLight.State
  QuickDockLightState = QuickDockLight.State
  StartModeLight.State = 0
  StartMissionLight.State = 0
  LockLight.State = 0
  VideomodeLight.State = 0
  ExplosionLight.State = 0
  QuickDockLight.State = 0
  QuickDockTimer.enabled = False
  DockRampSpider.state = Lightstateblinking
  ActivateMagnetLight.state = Lightstateblinking
  SuperLauncherBaseValue = SuperLauncherBaseValue + 25000000
  SuperLauncherAddValue = SuperLauncherAddValue + 25000000
  ActualSuperLauncherValue = SuperLauncherBaseValue
  SL_Loop = 15
  SL_Hits = 0
  SuperLauncherTimer.enabled = True
  PlayNextMusic(SuperLauncherTheme)
  DMDStartMission "SuperLauncher1",2000
End Sub

Dim SL_Loop
Sub SuperLauncherTimer_Timer
  SL_Loop = SL_Loop - 1
  if SL_Loop < 0 then
    playsound "W_SLLost"
    EndSuperLauncher
  end if
  if SL_Loop < 12 then
    FlashRight 1
    FlashLeft 1
  end if
  select case SL_Loop
    case 12: DMDStartMission "SuperLauncher2",12000
    case 11: playsound "W_SL10"
    case 10: playsound "W_SL9"
    case 9:  playsound "W_SL8"
    case 8:  playsound "W_SL7"
    case 7:  playsound "W_SL6"
    case 6:  playsound "W_SL5"
    case 5:  playsound "W_SL4"
    case 4:  playsound "W_SL3"
    case 3:  playsound "W_SL2"
    case 2:  playsound "W_SL1"
    case 1:  DMDDisplayScene "Mission",True,SLDback,FormatDMDTopText("","",""), 8,14,FormatDMDBottomText(""), 5,15, 14, 2000, 1
    case 0:  playsound "W_SL0"
  end select
End Sub

Dim SL_Hits
Sub AdvanceSuperLauncher
  SL_Hits = SL_Hits + 1
  Addpoints(ActualSuperLauncherValue)
  Playsound "W_USIntitBoom"
  FlashB2SLamp 50,1
  UpdateDMDText "Mission","","","",FormatPointText(cstr(ActualSuperLauncherValue))
  SuperLauncherTimer.enabled = False
  if SL_Hits >= 5 then
    FlashRight 3
    FlashLeft 3
    Showdownspheres = Showdownspheres + 1
    PlayNextSound "W_Sphere"
    DMDDisplayScene "Sphere",True,Sphereback, "", 8,14,"", 5,15, 14, 3000, 1
    EndSuperLauncher
  else
    SL_Loop = 13
    SuperLauncherTimer.enabled = True
    ActualSuperLauncherValue = ActualSuperLauncherValue + SuperLauncherAddValue
  end if
End Sub

Sub EndSuperLauncher
  SuperLauncherActive = False
  SuperLauncherTimer.enabled = False
  ActivateMagnetLight.state = abs(MagnetWasLit)
  StartModeLight.State = StartModeLightState
  StartMissionLight.State = StartMissionLightState
  LockLight.State = LockLightState
  LockLightstate = 0
  VideomodeLight.State = VideomodeLightState
  ExplosionLight.State = ExplosionLightState
  QuickDockLight.State = QuickDockLightState
  DockRampSpider.state = 0
  if SL_Loop < 0 then
    DMDEndMode
  end if
  PlayNextMusic(MainTheme)
End Sub


'--- Combo Mania ---
Sub StartComboMania
  ComboManiaReady = False
  ComboManiaActive = True
  DockState = DockBallLock.enabled
  DockLightState = DockLight.state
  DockBallLock.enabled = False
  ExplosionLightState = ExplosionLight.State
  ExplosionLight.State = 0
  ComboManiaValue = ComboManiaValue + 5000000
  LeftOrbitSpider.state = Lightstateblinking
  LeftRampSpider.state = Lightstateblinking
  LoopSpider.state = Lightstateblinking
  DockRampSpider.state = Lightstateblinking
  RightOrbitSpider.state = Lightstateblinking
  LeftOrbitSpider.UserValue = 1
  LeftRampSpider.UserValue = 1
  LoopSpider.UserValue = 1
  DockRampSpider.UserValue = 1
  RightOrbitSpider.UserValue = 1
  DMDDisplayScene "CM",True,ComboManiaback, "", 8,14,"", 5,15, 14, 3500, 14
  PlayNextSound "W_Combomania"
  PlayNextMusic(ComboManiaTheme)
  CombomaniaTimer.enabled = True
End Sub

Sub CombomaniaTimer_Timer
  CombomaniaTimer.enabled = False
  EndCombomania
  PlayNextMusic(MainTheme)
End Sub

Sub EndCombomania
  ComboManiaReady = False
  ComboManiaActive = False
  DeactivateComboShotLights
  DeactivateComboShotValues
  if US_CombosLight.state = 0 then
    US_CombosLight.state = 1
  end if
  DockBallLock.enabled = DockState
  DockLight.state = DockLightState
  ExplosionLight.State = ExplosionLightState
End Sub

'--- Secret Mania ---
Dim SecretManiaActive,SecretManiaScore,SecretFrenzyBalls
Sub StartSecretMania
  SecretManiaActive = True
  DMDDisplayScene "SM",True,SecretManiaback, "", 8,14,"", 5,15, 14, 3500, 14
  PlayNextSound "W_Cow"
  PlayNextmusic(FastFrenzyTheme)

  LeftHoleSpider.state = Lightstateblinking
  LeftOrbitSpider.state = Lightstateblinking
  LeftRampSpider.state = Lightstateblinking
  MissionHoleSpider.state = Lightstateblinking
  ReactorHoleSpider.state = Lightstateblinking
  LoopSpider.state = Lightstateblinking
  DockRampSpider.state = Lightstateblinking
  ExplosionHoleSpider.state = Lightstateblinking
  RightOrbitSpider.state = Lightstateblinking
  ExplosionLightState = ExplosionLight.State
  ExplosionLight.State = 0
  DeactivateComboShotValues
  FastFrenzyCount = FastFrenzyCount + 1

  LockLight.state = 0
  DockLight.state = 0
  DTDropUp
  DT1Light.state = Lightstateblinking
  DT2Light.state = Lightstateblinking
  DT3Light.state = Lightstateblinking
  DT4Light.state = Lightstateblinking
  DT5Light.state = Lightstateblinking
  DT6Light.state = Lightstateblinking
  JackpotLight.state = 0

  MultiBallBalls = 5
  LeftRampVUK.CreateBall
  LeftRampVUK.kick 180, 1
  playsound SoundFXDOF("W_Vuk",134,DOFPulse,DOFContactors),0,1,-0.18,0.25
  RightRampVUK.CreateBall
  RightRampVUK.kick 195, 1
  playsound SoundFXDOF("W_Vuk2",135,DOFPulse,DOFContactors),0,1,0.18,0.25
  DOF 136, DOFPulse
  BallsonPF = BallsonPF + 1
  BallsSaved = 3

  GracePeriodActive = True
  BallSaverLight.BlinkInterval = 160
  StartBallsaver
  BallreleaseTimer.enabled = True
  AutoPlungerActive = True
End Sub

Sub AwardSecretManiaScore
  DTCount = 0
  if DT1.isdropped then
    DTCount = DTCount + 1
  end if
  if DT2.isdropped then
    DTCount = DTCount + 1
  end if
  if DT3.isdropped then
    DTCount = DTCount + 1
  end if
  if DT4.isdropped then
    DTCount = DTCount + 1
  end if
  if DT5.isdropped then
    DTCount = DTCount + 1
  end if
  if DT6.isdropped then
    DTCount = DTCount + 1
  end if
  if DTCount = 6 then
    DTCount = 10
  end if
  if DTCount = 0 then
    SecretManiaScore = 1000000 * FastFrenzyBalls
  else
    SecretManiaScore = 200000 * DTCount * FastFrenzyBalls * (BallsonPF + FastFrenzyCount)
  end if

  addpoints(SecretManiaScore)

  FlashRight 2
  FlashLeft 2
  FlashB2SLamp 51,1

  select case int(7*rnd)
    case 0: playsound "W_NiceShot"
    case 1: playsound "W_Superb"
    case 2: playsound "W_IamImpressed"
    case 3: playsound "W_Outstanding"
    case 4: playsound "W_VeryGood"
    case 5: playsound "W_YesYesYes"
    case 6: playsound "W_Unbelievable"
  end select

  UpdateDMDText "Main","","BIG POINTS","",FormatPointText(cstr(SecretManiaScore))
End Sub

Sub EndSecretMania
  SecretManiaActive = False
  PlayNextmusic(MainTheme)

  LeftHoleSpider.state = 0
  LeftOrbitSpider.state = 0
  LeftRampSpider.state = 0
  MissionHoleSpider.state = 0
  ReactorHoleSpider.state = 0
  LoopSpider.state = 0
  DockRampSpider.state = 0
  ExplosionHoleSpider.state = 0
  RightOrbitSpider.state = 0
  ExplosionLight.State = ExplosionLightState

  LockLight.state = 0

  BallsLocked = 0
  if DT1.isdropped then
    DT1Light.state = 1
  else
    DT1Light.state = 0
  end if
  if DT2.isdropped then
    DT2Light.state = 1
  else
    DT2Light.state = 0
  end if
  if DT3.isdropped then
    DT3Light.state = 1
  else
    DT3Light.state = 0
  end if
  if DT4.isdropped then
    DT4Light.state = 1
  else
    DT4Light.state = 0
  end if
  if DT5.isdropped then
    DT5Light.state = 1
  else
    DT5Light.state = 0
  end if
  if DT6.isdropped then
    DT6Light.state = 1
  else
    DT6Light.state = 0
  end if

  DT6Light.state = Lightstateblinking
  JackpotLight.state = 0
End Sub

'--- Ultimate Showdown ---
Dim USPhase1Active,USPhase2Active,USPhase3Active,USCount,USCompleted,USScoreValue,ReserveBalls,UltimateShowdownAcitve

Sub StartUltimateShowdown
  for each obj in GI: obj.state = 0:next
  SetB2SLamp 1,0

  EndCurrentMusic
  UltimateShowdownReady = False
  UltimateShowdownAcitve = True
  DMDDisplayScene "US",True,"blank.png", FormatDMDTopText("","",""), 8,14,FormatDMDBottomText(""), 5,15, 14, 2000, 14
  USInitLoop = 0
  InitUSTimer.enabled = True
  LeftFlipper.Rotatetostart
  ULeftFlipper.Rotatetostart
  RightFlipper.Rotatetostart
  LockLight.state = 0
  DockLight.state = 0
  StartModeLight.state = 0
  StartMissionLight.state = 0
End Sub

Dim USInitLoop
Sub InitUSTimer_Timer
  USInitLoop = USInitLoop + 1
  select case USInitLoop
    case 1: PlaySound "W_USIntitBoom",0,0.75,0,0
        FlashB2SLamp 48,1
        FlashB2SLamp 49,1
        FlashB2SLamp 51,1
        DMDDisplayScene "US",True,"blank.png", FormatDMDTopText("","ULTRA  JACKPOT",""), 8,14,FormatDMDBottomText("SCORED"), 5,15, 14, 2000, 14
    case 2: PlaySound "W_USIntitBoom",0,0.8,0,0
        FlashB2SLamp 48,1
        FlashB2SLamp 49,1
        FlashB2SLamp 51,1
        DMDDisplayScene "US",True,"blank.png", FormatDMDTopText("","COMBO  MANIA",""), 8,14,FormatDMDBottomText("PLAYED"), 5,15, 14, 2000, 14
    case 3: PlaySound "W_USIntitBoom",0,0.85,0,0
        FlashB2SLamp 48,1
        FlashB2SLamp 49,1
        FlashB2SLamp 51,1
        DMDDisplayScene "US",True,"blank.png", FormatDMDTopText("","POWER LEVEL  8",""), 8,14,FormatDMDBottomText("REACHED"), 5,15, 14, 2000, 14
    case 4: PlaySound "W_USIntitBoom",0,0.9,0,0
        FlashB2SLamp 48,1
        FlashB2SLamp 49,1
        FlashB2SLamp 51,1
        DMDDisplayScene "US",True,"blank.png", FormatDMDTopText("","BONUS",""), 8,14,FormatDMDBottomText("MAXED  OUT"), 5,15, 14, 2000, 14
    case 5: PlaySound "W_USIntitBoom",0,0.95,0,0
        FlashB2SLamp 48,1
        FlashB2SLamp 49,1
        FlashB2SLamp 51,1
        DMDDisplayScene "US",True,"blank.png", FormatDMDTopText("","ALL MISSIONS",""), 8,14,FormatDMDBottomText("COMPLETE"), 5,15, 14, 2000, 14
    case 6: PlaySound "W_USIntitBoomLast",0,1,0,0
        FlashB2SLamp 48,1
        FlashB2SLamp 49,1
        FlashB2SLamp 51,1
        DMDDisplayScene "US",True,"blank.png", FormatDMDTopText("","ALL SWITCHES",""), 8,14,FormatDMDBottomText(FormatPointText(cstr(1000000 * ShowdownSpheres))), 5,15, 14, 2000, 14
    case 7: InitUSTimer.enabled = False
        StartPhase1
  end select
End Sub

Sub AwardUSPoints
  if UltimateShowdownActive then
    addpoints(USScoreValue * (USCount + 1))
    AddPhase1Hits
  end if
End Sub

Sub UltimateShowdownMystery
  if MysteryLight.state <> 0 then
    FlashB2SLamp 48,1
    FlashB2SLamp 49,1
    FlashB2SLamp 51,1
    if int(2*rnd) = 1 then
      PlayNextSound "W_Mystery"
    else
      PlayNextSound "W_Mystery_F"
    end if
    ReserveBalls = ReserveBalls - 1
    BallsSaved = BallsSaved + 1
    MultiBallBalls = MultiBallBalls + 1
    BallreleaseTimer.enabled = True
    AutoPlungerActive = True
  end if
  if USPhase2Active then
    FlashB2SLamp 50,1
    EjectToRight = not EjectToRight
    if EjectToRight then    'init left
      USLeft = True
      USRight = False
    else
      USLeft = False
      USRight = True
    end if
  end if
End Sub

Dim Phase1Hits
Sub StartPhase1     '150-200 Switch Hits
  PlayNextSound "W_UltimateShowdown"
  PlayNextMusic(UltimateshowdownTheme)
  for each obj in GI: obj.state = 1:next
  SetB2SLamp 1,1

  DeactivateMissionsAndModes
  DockBallLock.enabled = False
  UltimateShowdownReady = False
  UltimateShowdownActive = True
  USScoreValue = 1000000 * ShowdownSpheres
  ReserveBalls = ShowdownSpheres
  if EjectToRight then
    TroughRightMissionKicker.CreateBall
    TroughRightMissionKicker.kick 180, 1
  else
    TroughLeftMissionKicker.CreateBall
    TroughLeftMissionKicker.kick 180, 1
  end if
  if MultiballBalls > 1 then
    BallsSaved = 6 - MultiballBalls
    FastFrenzyActive = False
    SecretManiaActive = False
    SpaceStationFrenzyActive = False
  else
    BallsSaved = 5
  end if
  MultiballBalls = 6
  GracePeriodActive = True
  BallSaverLight.BlinkInterval = 160
  GracePeriod.interval = 27000
  GracePeriod1.interval = 30000
  GracePeriod2.interval = 32000
  StartBallSaver
  BallreleaseTimer.enabled = True
  AutoPlungerActive = True
  USPhase1Active = True
  USPhase2Active = False
  USPhase3Active = False
  MysteryLight.state = 0
  Phase1Hits = 0
  LeftOrbitSpider.state = Lightstateblinking
  LeftRampSpider.state = Lightstateblinking
  LoopSpider.state = Lightstateblinking
  DockRampSpider.state = Lightstateblinking
  RightOrbitSpider.state = Lightstateblinking
  LeftHoleSpider.state = Lightstateblinking
  MissionHoleSpider.state = Lightstateblinking
  ReactorHoleSpider.state = Lightstateblinking
  ExplosionHoleSpider.state = Lightstateblinking
  DMDDisplayScene "US1",True,"US_1_1.png","",8,14,"",5,15,14,100000,14
End Sub

Sub AddPhase1Hits
  if USPhase1Active then
    Phase1Hits = Phase1Hits + 1
    select case Phase1Hits
      case 20:  Playsound "W_Explosion"
      case 40:  Playsound "W_Explosion"
            DMDDisplayScene "US1",True,"US_1_2.png","",8,14,"",5,15,14,1000000,14
      case 60:  Playsound "W_Explosion"
      case 80:  Playsound "W_Explosion"
            DMDDisplayScene "US1",True,"US_1_3.png","",8,14,"",5,15,14,1000000,14
      case 100: Playsound "W_Explosion"
      case 120: Playsound "W_Explosion"
            DMDDisplayScene "US1",True,"US_1_4.png","",8,14,"",5,15,14,1000000,14
      case 140: Playsound "W_Explosion"
    end select
    if Phase1Hits >= 150 then
      DMDDisplayScene "US1",True,"US_1_5.png","",8,14,"",5,15,14,1000000,14
      LeftRampSpider.state = 0
      DockRampSpider.state = 0
      StartPhase2
    end if
  end if
End Sub

Dim Phase2Spins,USLeft,USRight
Sub StartPhase2     '4 Spinner Orbits, only one Spinner is active, Mystery-Target changes the active one, 4 different quotes ("C'mon men", "Stop!","No!","Finish her!")
  PlayNextSound "W_ShootSpinners"
  USPhase1Active = False
  USPhase2Active = True
  Phase2Spins = 0
  LeftOrbitSpider.state = 0
  LeftRampSpider.state = 0
  LoopSpider.state = 0
  DockRampSpider.state = 0
  RightOrbitSpider.state = 0
  LeftHoleSpider.state = 0
  MissionHoleSpider.state = 0
  ReactorHoleSpider.state = 0
  ExplosionHoleSpider.state = 0
  PowerLevel1Light.state = 0
  PowerLevel2Light.state = 0
  PowerLevel3Light.state = 0
  PowerLevel4Light.state = 0
  PowerLevel5Light.state = 0
  PowerLevel6Light.state = 0
  PowerLevel7Light.state = 0
  PowerLevel8Light.state = 0
  QuickShotAnimationTimer.enabled = True
  if EjectToRight then    'init left
    USLeft = True
    USRight = False
  else
    USLeft = False
    USRight = True
  end if
End Sub

Sub AddPhase2Spins
  if USPhase2Active then
    FlashRight 1
    FlashLeft 1
    Phase2Spins = Phase2Spins + 1
    select case Phase2Spins
      case 1: PlayNextSound "W_Comon"
          DMDDisplayScene "US2",True,"US_2_1.png","",8,14,"",5,15,14,100000,14
      case 2: PlayNextSound "W_Stop"
          DMDDisplayScene "US2",True,"US_2_2.png","",8,14,"",5,15,14,100000,14
      case 3: PlayNextSound "W_No"
          DMDDisplayScene "US2",True,"US_2_3.png","",8,14,"",5,15,14,100000,14
    end select
    if Phase2Spins >= 4 then
      EndQuickShot
      StartPhase3
    end if
  end if
End Sub

Sub StartPhase3     'Finish Her (shoot Mission hole) - awards 1.000.000.000 points + Start 6-Ball Fast Frenzy (or any Balls on PF)
  PlayNextSound "W_FinishHer"
  DMDDisplayScene "US3",True,"US_3.png","",8,14,"",5,15,14,100000,14
  USPhase2Active = False
  USPhase3Active = True
  Phase3Step = 0
  Phase3State = True
  USPhase3AnimationTimer.enabled = True
End Sub

Dim Phase3Step,Phase3State
Sub USPhase3AnimationTimer_Timer
  Phase3Step = Phase3Step + 1
  if Phase3Step > 4 then
    Phase3State = not Phase3State
    Phase3Step = 1
  end if
  select case Phase3Step
    case 1: StartMissionLight.state = abs(Phase3State)
    case 2: StartModeLight.state = abs(Phase3State)
    case 3: EBLight.state = abs(Phase3State)
    case 4: MissionHoleSpider.state = abs(Phase3State)
  end select
End Sub

Sub EndUSPhase3
  FlashRight 5
  FlashLeft 5
  USPhase3Active = False
  USCompleted = True
  UltimateShowdownActive = False
  EndUltimateShowdown
End Sub

Sub EndUltimateShowdown
  USPhase3AnimationTimer.enabled = False
  StartMissionLight.state = 0
  StartModeLight.state = 0
  EBLight.state = 0
  LeftOrbitSpider.state = 0
  LeftRampSpider.state = 0
  LoopSpider.state = 0
  DockRampSpider.state = 0
  RightOrbitSpider.state = 0
  LeftHoleSpider.state = 0
  MissionHoleSpider.state = 0
  ReactorHoleSpider.state = 0
  ExplosionHoleSpider.state = 0
  USPhase1Active = False
  USPhase2Active = False
  USPhase3Active = False
  MysteryLight.state = 0
  EndQuickShot
  if USCompleted then
    USCount = USCount + 1
    USWonLoop = 0
    Playsound "W_Won"
    USWonLoop = 0
    USWonTimer.enabled = True
  else
    UltimateShowdownActive = False
    US_MissionsLight.state = 0
    US_PowerLight.state = 0
    US_UltraLight.state = 0
    US_BonusLight.state = 0
    US_CombosLight.state = 0
    USLight.state = 0
    DMDDisplayScene "Main",True,backgrnd, FormatDMDTopText("","",""), 8,14,FormatDMDBottomText(""), 5,15, 14, 10000000, 14
    SuperKickbackActive = False
    PlayNextSound "W_YouCantWin"
    playnextmusic(MainTheme)
    MultiBallBalls = 0
  end if
End Sub

Dim USWonLoop
Sub USWonTimer_Timer
  USWonLoop = USWonLoop + 1
  if USWonLoop = 2 then
    Playsound "W_USWon2"
    DMDDisplayScene "US4",True,UltimateShowdownBack,"",8,14,"",5,15,14,100000,14
  end if
  if USWonLoop = 20 then
    DMDDisplayScene "US",True,"blank.png","1.000.000.000", 8,14,"", 5,15, 14, 7000, 14
  end if
  if USWonLoop >= 25 then
    USWonTimer.enabled = False
    FlashB2SLamp 48,3
    FlashB2SLamp 49,3
    FlashB2SLamp 50,3
    FlashB2SLamp 51,3
    addpoints(1000000000 * USCount)
    DMDDisplayScene "Main",True,backgrnd, FormatDMDTopText("","",""), 8,14,FormatDMDBottomText(""), 5,15, 14, 10000000, 14
    'Phase 4: 6-Ball Fast Frenzy
    FastFrenzyBalls = 6
    BallsLocked = 0
    'UltimateShowdownActive = False
    US_MissionsLight.state = 0
    US_PowerLight.state = 0
    US_UltraLight.state = 0
    US_BonusLight.state = 0
    US_CombosLight.state = 0
    USLight.state = 0
    BallsOnPF = 1
    PlayNextSound "W_YouHaventSeen"
    StartFastFrenzy
  end if
End Sub

'--- ExtraBalls ---
Sub LeftOutlane_Hit
  DOF 122, DOFOn
  if LeftOutlaneEBLight.state = 1 then
    LeftOutlaneEBLight.state = 0
    RightOutlaneEBLight.state = 0
    ExtraBalls = ExtraBalls + 1
    PlayNextSound "W_ExtraBall"
    DOF 123, DOFPulse
    FlashB2SLamp 50,1
    if not hideanimation then
      DMDDisplayScene "EB",True,ExtraBallback, "", 8,14,"", 5,15, 14, 3900, 1
    end if
  end if
  AdvanceShuttleRampValue
  AwardUSPoints
  if kickbacklight.state = 0 then
    PlaySound "W_LOutlane"
  end if
End Sub
sub LeftOutlane_unhit
  DOF 122, DOFOff
End Sub

Sub RightOutlane_Hit
  DOF 124, DOFOn
  if RightOutlaneEBLight.state = 1 then
    LeftOutlaneEBLight.state = 0
    RightOutlaneEBLight.state = 0
    ExtraBalls = ExtraBalls + 1
    PlayNextSound "W_ExtraBall"
    DOF 123, DOFPulse
    FlashB2SLamp 50,1
    if not hideanimation then
      DMDDisplayScene "EB",True,ExtraBallback, "", 8,14,"", 5,15, 14, 3900, 1
    end if
  end if
  AdvanceShuttleRampValue
  AwardUSPoints
  PlaySound "W_ROutlane"
End Sub
sub RightOutlane_unhit
  DOF 124, DOFOff
End Sub

'--- Quick Dock ---
Sub QuickDockTimer_Timer
  QuickDockTimer.enabled = False
  ActivateMagnetLight.state = 0
end Sub

'--- Explosion ---
Dim ExplosionAward,ExText,StartMagnet,StartQuickDock
Sub ExplosionHole_Hit()
  ExplosionHole.enabled = False
  ExplosionHole.DestroyBall
  Playsound "Drain3",0,0.3,0.05,0.25

  if HuntDownActive and (ExplosionHolespider.state <> 0) then
    AdvanceHuntDownScore
  end if

  StartMagnet = False
  StartQuickDock = False
  if ActivateMagnetLight.state <> 0 then
    StartMagnet = True
    if QuickDockTimer.enabled and not SecretManiaActive and not UltimateShowdownActive then
      QuickDockTimer.enabled = False
      if Docklight.state = 0 then
        StartQuickDock = True
      end if
    end if
    if not SuperLauncherActive then
      ActivateMagnetLight.state = 0
    end if
    MagnetGrabTrigger.enabled = True
    For Each obj in MagnetLights:obj.state = LightStateblinking:next
  end if

  if UltimateShowdownActive then
    AwardUSPoints
  else
    if SecretManiaActive then
      AwardSecretManiaScore
    else
      if SpaceStationFrenzyActive then
        if (ExplosionHoleSpider.state <> 0) and PrepareUltraJackpot then
          PlayNextSound "W_GetUltraJackpot"
          ExplosionHoleSpider.state = 0
          UltraJackpotLight.state = Lightstateblinking
          PrepareUltraJackpot = False
          UltraJackpotready = True
          ActivateMagnetLight.state = 0
          QuickDockTimer.enabled = False
          UltraJackpotTimer.enabled = True
        end if
      else
        if ExplosionLight.state = 1 and not MissionActive and not FastFrenzyActive and not Quickshotactive and not Combomaniaactive then
          ExplosionLight.state = 0
          FlashB2SLamp 48,1
          FlashB2SLamp 49,1
          FlashB2SLamp 51,1
          playsound "W_Explosion"
          DMDDisplayScene "Ex",True,Explosionback,FormatDMDTopText("","",""), 8,14,FormatDMDBottomText(""), 7,15, 14, 3200, 14
          ExTimerLoop = 0
          ExplosionAwardTimer.enabled = True
          Exit Sub
        end if
      end if
    end if
  end if

  VUKFire
End Sub

Dim ExTimerLoop,StartExplosionVideoMode,StartExplosionFastFrenzy
Sub ExplosionAwardTimer_Timer
  ExTimerLoop = ExTimerLoop + 1
  if ExTimerLoop > 2 then
    ExplosionAwardTimer.enabled = False
    if not StartExplosionVideoMode then
      if StartExplosionFastFrenzy then
        DeactivateMissionsAndModes
        ExplosionFF = True
        FastFrenzyBalls = 3
        StartFastFrenzy
      end if
      VUKFire
    else
      DeactivateMissionsAndModes
      StartVideoMode True
    end if
  end if
  if ExTimerLoop = 2 then
    StartExplosionVideoMode = False
    StartExplosionFastFrenzy = False
    ExplosionAward = Int(16*Rnd)
    select case ExplosionAward
      case 0: if LeftOutlaneEBLight.state = 0 then
            ExText = "OUTLANES LIT"
            LeftOutlaneEBLight.state = 1
            RightOutlaneEBLight.state = 1
          else
            ExText = "EXTRA BALL"
            ExtraBalls = ExtraBalls + 1
            PlayNextSound "W_ExtraBall"
          end if
      case 1: if EBLight.state = 0 then
            ExText = "EXTRA BALL LIT"
            EBLight.state = 1
            PlayNextSound "W_ExtraBallLit"
          else
            ExText = "EXTRA BALL"
            ExtraBalls = ExtraBalls + 1
            PlayNextSound "W_ExtraBall"
          end if
      case 2: ExText = "EXTRA BALL"
          ExtraBalls = ExtraBalls + 1
          PlayNextSound "W_ExtraBall"
          DOF 123, DOFPulse
      case 3: ExText = "POWERLEVEL"
          IncreasePowerlevel
      case 4: if SuperJets < 3 then
            ExText = "SUPER JETS"
            SuperJets = SuperJets + 1
            PlayNextSound "W_SuperJets"
          else
            ExText = "POWERLEVEL"
            IncreasePowerlevel
          end if
      case 5: if not SuperJetsHeld then
            ExText = "SUPERJETS HELD"
            SuperJetsHeld = True
          else
            ExText = "EXPLOSION  LIT"
            ExplosionLight.state = 1
          end if
      case 6: if VideoModeLight.state = 0 then
            ExText = "VIDEO MODE LIT"
            VideoModeLight.state = 1
          else
            ExText = "VIDEO MODE"
            StartExplosionVideoMode = True
          end if
      case 7: ExText = "VIDEO MODE"
          StartExplosionVideoMode = True
      case 8: if not QuickShotReady then
            ExText = "QUICK SHOT LIT"
            QuickShotReady = True
          else
            ExText = "ADD  BONUS X"
            AwardNextBonus
          end if
      case 9: ExText = "ADD  BONUS X"
          AwardNextBonus
      case 10:  if not BonusHeld then
              ExText = "BONUS HELD"
              BonusHeld = True
            else
              ExText = "EXPLOSION  LIT"
              ExplosionLight.state = 1
            end if
      case 11:  ExText = "BALL SAVER"
            GracePeriodActive = True
            BallSaverLight.BlinkInterval = 160
            GracePeriod.interval = 27000
            GracePeriod1.interval = 30000
            GracePeriod2.interval = 32000
            StartBallSaver
      case 12:  if ActivateMagnetLight.state = 0 then
              ExText = "MAGNET LIT"
              ActivateMagnetLight.state = 1
            else
              ExText = "EXPLOSION  LIT"
              ExplosionLight.state = 1
            end if
      case 13:  if not SuperKickbackActive then
              ExText = "SUPER RESCUE"
              SuperKickbackActive = True
            else
              ExText = "EXPLOSION  LIT"
              ExplosionLight.state = 1
            end if
      case 14:  ExText = "FAST FRENZY"
            StartExplosionFastFrenzy = True
      case 15:  ExText = ""
            Addpoints(Jackpotvalue * 3)
            Playnextsound "W_Cow"
    end select
    UpdateDMDExplosion(ExText)
  end if
End Sub

Sub Docklight_Timer
  Docklight.Timerenabled = False
  Docklight.state = 0
  DockBallLock.enabled = False
End Sub

'--- PowerLevel ---
Sub RightInlane_Hit
  DOF 127, DOFOn
  if (MultiBallBalls < 2) and not Quickshotactive then
    RightInlaneLight.state = LightstateBlinking
    RightInlaneLight.timerenabled = False
    RightInlaneLight.timerenabled = True
    LeftPowerLevelTimer.enabled = False
    LeftPowerLevelTimer.enabled = True
    if not SkyscraperActive then
      AdvanceLeftPowerlevelLight.state = Lightstateblinking
    end if
  end if
  AdvanceShuttleRampValue
  AwardUSPoints
End Sub
sub RightInlane_unhit
  DOF 127, DOFOff
End Sub
Sub RightInlaneLight_Timer
  RightInlaneLight.timerenabled = False
  RightInlaneLight.state = 1
End Sub

Sub LeftPowerLevelTimer_Timer
  LeftPowerLevelTimer.enabled = False
  AdvanceLeftPowerlevelLight.state = Lightstateoff
End Sub

Sub RightPowerLevelTimer_Timer
  RightPowerLevelTimer.enabled = False
  AdvanceRightPowerlevelLight.state = Lightstateoff
End Sub

Sub LeftSpinnerTrigger_hit
  if LeftPowerLevelTimer.enabled then
    if not LeftPauseTimer.enabled then
      IncreasePowerlevel
      LeftPowerLevelTimer.enabled = False
      LeftPowerLevelTimer.enabled = True
      LeftPauseTimer.enabled = True
    end if
  else
    LeftSpinnerTrigger.Timerenabled = False
    LeftSpinnerTrigger.Timerenabled = True
  end if
  AdvanceShuttleRampValue
  AdvanceSkyscraperScore
  AwardUSPoints
end sub

Sub LeftSpinnerTrigger_Timer
  LeftSpinnerTrigger.Timerenabled = False
End Sub

Sub LeftPauseTimer_Timer
  LeftPauseTimer.enabled = False
End Sub

Sub RightSpinnerTrigger_hit
  if RightPowerLevelTimer.enabled then
    if not RightPauseTimer.enabled then
      IncreasePowerlevel
      RightPowerLevelTimer.enabled = False
      RightPowerLevelTimer.enabled = True
      RightPauseTimer.enabled = True
    end if
  else
    RightSpinnerTrigger.Timerenabled = False
    RightSpinnerTrigger.Timerenabled = True
  end if
  AdvanceShuttleRampValue
  AdvanceSkyscraperScore
  AwardUSPoints
end sub

Sub RightSpinnerTrigger_Timer
  RightSpinnerTrigger.Timerenabled = False
End Sub

Sub RightPauseTimer_Timer
  RightPauseTimer.enabled = False
End Sub

Dim Powerlevel
Sub IncreasePowerlevel
  if not QuickshotActive and (MultiballBalls < 2) then
    FlashB2SLamp 48,1
    FlashB2SLamp 49,1
    select case Powerlevel
      case 1: AddPoints(5000000)
          if not SkyscraperActive then
            PlayNextSound "W_Powerlevel1"
          end if
          DMDPowerLevel Powerlevel,"SMALL POINTS"
          PowerLevel1Light.state = 1
          PowerLevel2Light.state = Lightstateblinking
      case 2: if MissionActive or Combomaniaactive then
            ExplosionLightState = 1
          else
            ExplosionLight.state = 1
          end if
          if not SkyscraperActive then
            PlayNextSound "W_Powerlevel2"
          end if
          DMDPowerLevel Powerlevel,"EXPLOSION READY"
          PowerLevel2Light.state = 1
          PowerLevel3Light.state = Lightstateblinking
      case 3: if SuperJets < 3 then
            SuperJets = SuperJets + 1
          end if
          if not SkyscraperActive then
            PlayNextSound "W_Powerlevel3"
          end if
          DMDPowerLevel Powerlevel,"SUPER JETS"
          PowerLevel3Light.state = 1
          PowerLevel4Light.state = Lightstateblinking
      case 4: QuickShotReady = True
          if not SkyscraperActive then
            PlayNextSound "W_Powerlevel4"
          end if
          DMDPowerLevel Powerlevel,"QUICKSHOT READY"
          PowerLevel4Light.state = 1
          PowerLevel5Light.state = Lightstateblinking
      case 5: VideoModeReady = True
          if not SkyscraperActive then
            PlayNextSound "W_Powerlevel5"
          end if
          DMDPowerLevel Powerlevel,"VIDEOMODE READY"
          PowerLevel5Light.state = 1
          PowerLevel6Light.state = Lightstateblinking
      case 6: SuperKickbackActive = True
          if not SkyscraperActive then
            PlayNextSound "W_Powerlevel6"
          end if
          DMDPowerLevel Powerlevel,"SUPER KICKBACK"
          PowerLevel6Light.state = 1
          PowerLevel7Light.state = Lightstateblinking
      case 7: SuperLauncherReady = True
          if not SkyscraperActive then
            PlayNextSound "W_Powerlevel7"
          end if
          DMDPowerLevel Powerlevel,"SUPERLAUNCHER"
          PowerLevel7Light.state = 1
          PowerLevel8Light.state = Lightstateblinking
      case 8: ShowdownSpheres = ShowdownSpheres + 1
          if not SkyscraperActive then
            PlayNextSound "W_SphereSpin"
          end if
          DMDPowerLevel Powerlevel,"SHOWDOWNSPHERE"
          PowerLevel8Light.state = 1
          US_PowerLight.state = 1
      case else:
        if not SkyscraperActive then
          PlayNextSound "W_Powerlevel_x"
        end if
        Addpoints(Powerlevel*1000000)
        DMDPowerLevel Powerlevel,FormatPointText(Powerlevel*1000000)
    end select
    Powerlevel = Powerlevel + 1
    TotalPowerlevelCount = TotalPowerlevelCount + 1
  end if
End Sub

Sub ResetPowerLevel
  PowerLevel = 1
  PowerLevel1Light.state = Lightstateblinking
  PowerLevel2Light.state = 0
  PowerLevel3Light.state = 0
  PowerLevel4Light.state = 0
  PowerLevel5Light.state = 0
  PowerLevel6Light.state = 0
  PowerLevel7Light.state = 0
  PowerLevel8Light.state = 0
End Sub

'--- Kickback ---
Sub LeftInlane_Hit
  DOF 125, DOFOn
  if not KickbackActive and not SuperKickbackActive then
    EnableKickBackTimer.enabled = False
    EnableKickBackTimer.enabled = True
  end if
  if (MultiBallBalls < 2) and not Quickshotactive then
    LeftInlaneLight.state = LightstateBlinking
    LeftInlaneLight.timerenabled = False
    LeftInlaneLight.timerenabled = True
    RightPowerLevelTimer.enabled = False
    RightPowerLevelTimer.enabled = True
    if not SkyscraperActive then
      AdvanceRightPowerlevelLight.state = LightstateBlinking
    end if
  end if
  AdvanceShuttleRampValue
  AwardUSPoints
End Sub
sub LeftInlane_unhit
  DOF 125, DOFOff
End Sub
Sub LeftInlaneLight_Timer
  LeftInlaneLight.timerenabled = False
  LeftInlaneLight.state = 1
End Sub

Sub LeftInlane2_Hit
  DOF 126, DOFOn
  if QuickDockLight.state <> 0 then
    QuickDockLight.state = Lightstateblinking
    QuickDockLight.timerenabled = False
    QuickDockLight.timerenabled = True
    if ActivateMagnetLight.state = 0 then
      ActivateMagnetLight.state = Lightstateblinking
    end if
    QuickDockTimer.enabled = False
    QuickDockTimer.enabled = True
  end if
  if not KickbackActive and not SuperKickbackActive then
    EnableKickBackTimer.enabled = False
    EnableKickBackTimer.enabled = True
  end if
  AdvanceShuttleRampValue
  AwardUSPoints
End Sub
sub LeftInlane2_unhit
  DOF 126, DOFOff
End Sub
Sub QuickDockLight_Timer
  QuickDockLight.timerenabled = False
  QuickDockLight.state = 1
End Sub

Sub EnableKickBackTimer_Timer
  EnableKickBackTimer.enabled = False
End Sub

Sub RightTarget_Hit
  Playsound SoundFXDOF("W_RightTarget",115,DOFPulse,DOFContactors),0,1,0.18,0.25
  if EnableKickBackTimer.enabled then
    if not Hideanimation then
      DMDDisplayScene "Rescue",True,RescueLitback, "", 8,14,"", 5,15, 14, 2000, 1
    end if
    PlayNextSound "W_BallRescueReady"
    EnableKickBackTimer.enabled = False
    KickBackActive = True
  end if
  AdvanceShuttleRampValue
  AwardUSPoints
End Sub

Sub Kickback_Hit
  if KickbackActive or SuperKickbackActive then
    FlashRight 1
    FlashLeft 1
    FlashB2SLamp 50,1
    FlashB2SLamp 51,1
    if not HideAnimation then
      DMDDisplayScene "BR",True,BallRescueback, "", 8,14,"", 5,15, 14, 2000, 14
    end if
    select case int(3*rnd)
      case 0:   playsound "W_BallRescued"
      case 1,2: playsound "W_KickBackShot"
    end select
    DOF 121, DOFPulse
    ActiveBall.vely = -70 - (rnd * 10)
    if not SuperKickbackActive and not GracePeriod.enabled then
      KickbackGracePeriodTimer.enabled = True
    end if
  end if
end sub

Sub KickbackGracePeriodTimer_Timer
  KickbackGracePeriodTimer.enabled = False
  KickBackActive = False
End Sub


'--- Top Bonus Rollovers ---
Sub TopRollover1_Hit
  DOF 128, DOFOn
  AddPoints(1000)
  BonusLightArray(1) = 1
  AdvanceShuttleRampValue
  AwardUSPoints
End Sub
Sub TopRollover1_Unhit
  DOF 128, DOFOff
End Sub

Sub TopRollover2_Hit
  DOF 129, DOFOn
  AddPoints(1000)
  BonusLightArray(2) = 1
  AdvanceShuttleRampValue
  AwardUSPoints
End Sub
Sub TopRollover2_Unhit
  DOF 129, DOFOff
End Sub

Sub TopRollover3_Hit
  DOF 130, DOFOn
  AddPoints(1000)
  BonusLightArray(3) = 1
  AdvanceShuttleRampValue
  AwardUSPoints
End Sub
Sub TopRollover3_Unhit
  DOF 130, DOFOff
End Sub

'--- Bumper Mini Lane ---
Dim SuperJets

Sub BumperLane1_Hit
  BumperLane1.Timerenabled = True
End Sub
Sub BumperLane1_Timer
  BumperLane1.Timerenabled = False
End Sub

Sub BumperLane_Hit
  if BumperLane1.Timerenabled and not AmmoDumpActive then
    BumperLane.Timerenabled = True
  end if
End Sub
Sub BumperLane_Timer
  BumperLane.Timerenabled = False
End Sub

Sub AdvanceSuperJetTimer_Timer
  AdvanceSuperJetTimer.enabled = False
  if not AmmoDumpActive then
    BumperLaneLight.state = 0
  end if
end Sub

Sub CheckBumperHit
  if BumperLane.Timerenabled and (SuperJets < 3) then
    if AdvanceSuperJetTimer.enabled then
      AdvanceSuperJetTimer.enabled = False
      BumperLane.Timerenabled = False
      BumperLane1.Timerenabled = False
      BumperLaneLight.state = 0
      PlayNextSound "W_SuperJets2"
      SuperJets = SuperJets + 1
    else
      AdvanceSuperJetTimer.enabled = True
      BumperLaneLight.state = Lightstateblinking
      BumperLane.Timerenabled = False
      BumperLane1.Timerenabled = False
    end if
  end if
end sub

dim ActiveBumper
ActiveBumper = 1
SuperJets = 2
Sub BumperLightTimer_Timer
  ActiveBumper = ActiveBumper + 1
  if ActiveBumper > 6 then
    ActiveBumper = 1
  end if
  if Ammodumpactive then
    B1L1.State = abs((ActiveBumper = 1) or (ActiveBumper = 2))
    B1L2.State = abs((ActiveBumper = 1) or (ActiveBumper = 2))
    B2L1.State = abs((ActiveBumper = 3) or (ActiveBumper = 4))
    B2L2.State = abs((ActiveBumper = 3) or (ActiveBumper = 4))
    B3L1.State = abs((ActiveBumper = 5) or (ActiveBumper = 6))
    B3L2.State = abs((ActiveBumper = 5) or (ActiveBumper = 6))
  else
    if SuperJetsHeld then
      B1L1.State = abs(SuperJets > 0)
      B1L2.State = abs(SuperJets > 0)
      B2L1.State = abs(SuperJets > 1)
      B2L2.State = abs(SuperJets > 1)
      B3L1.State = abs(SuperJets > 2)
      B3L2.State = abs(SuperJets > 2)
    else
      B1L1.State = abs((SuperJets > 0) and (ActiveBumper < 4))
      B1L2.State = abs((SuperJets > 0) and (ActiveBumper < 4))
      B2L1.State = abs((SuperJets > 1) and (ActiveBumper < 4))
      B2L2.State = abs((SuperJets > 1) and (ActiveBumper < 4))
      B3L1.State = abs((SuperJets > 2) and (ActiveBumper < 4))
      B3L2.State = abs((SuperJets > 2) and (ActiveBumper < 4))
    end if
  end if

  if B1L1.state = 0 Then
    P_BumperCap1.image = "plastic_bumpercap_red"
  Else
    P_BumperCap1.image = "plastic_bumpercap_red_lit"
  end If
  if b2l1.state = 0 Then
    P_BumperCap2.image = "plastic_bumpercap_red"
  Else
    P_BumperCap2.image = "plastic_bumpercap_red_lit"
  end If
  if B3L1.state = 0 Then
    P_BumperCap3.image = "plastic_bumpercap_red"
  Else
    P_BumperCap3.image = "plastic_bumpercap_red_lit"
  end If
End Sub


Sub Bumper1_Hit
  CheckBumperHit
  JackpotValue = JackpotValue + 7130
  if SuperJets = 0 then
    AddPoints(1000)
  else
    AddPoints(1000000 * SuperJets)
    DOF 116, DOFPulse
  end if
  PlaySound SoundFXDOF("",109,DOFPulse,DOFContactors),0,0.75,-0.1,0.25
  PlaySoundAt "bumper", Bumper1
  DOF 112, DOFPulse
  if not MissionActive then
    AdvanceMission
  end if
  if AmmoDumpActive then
    AdvanceAmmoDumpScore
  end if
  AdvanceShuttleRampValue
  AwardUSPoints
End Sub

Sub Bumper2_Hit
  CheckBumperHit
  JackpotValue = JackpotValue + 7130
  if SuperJets = 0 then
    AddPoints(1000)
  else
    AddPoints(1000000 * SuperJets)
    DOF 116, DOFPulse
  end if
  PlaySound SoundFXDOF("",107,DOFPulse,DOFContactors),0,0.75,-0.05,0.25
  PlaySoundAt "bumper", Bumper2
  DOF 110, DOFPulse
  if not MissionActive then
    AdvanceMission
  end if
  if AmmoDumpActive then
    AdvanceAmmoDumpScore
  end if
  AdvanceShuttleRampValue
  AwardUSPoints
End Sub

Sub Bumper3_Hit
  CheckBumperHit
  JackpotValue = JackpotValue + 7130
  if SuperJets = 0 then
    AddPoints(1000)
  else
    AddPoints(1000000 * SuperJets)
    DOF 116, DOFPulse
  end if
  PlaySound SoundFXDOF("bumper",108,DOFPulse,DOFContactors),0,0.75,0,0.25
  PlaySoundAt "bumper", Bumper3
  DOF 111, DOFPulse
  if not MissionActive then
    AdvanceMission
  end if
  if AmmoDumpActive then
    AdvanceAmmoDumpScore
  end if
  AdvanceShuttleRampValue
  AwardUSPoints
End Sub

'Slingshots
Dim RStep, Lstep
Sub LeftSlingshot_Slingshot
  AddPoints(1100)
  PlaySound SoundFXDOF("lsling",103,DOFPulse,DOFContactors),0,0.75,-0.15,0.25
  DOF 105, DOFPulse
    LSling.Visible = 0
    LSling1.Visible = 1
    sling2.TransZ = -20
    LStep = 0
    LeftSlingShot.TimerEnabled = 1
  AdvanceShuttleRampValue
  AwardUSPoints
End Sub

Sub LeftSlingShot_Timer
    Select Case LStep
        Case 3:LSLing1.Visible = 0:LSLing2.Visible = 1:sling2.TransZ = -10
        Case 4:LSLing2.Visible = 0:LSLing.Visible = 1:sling2.TransZ = 0:LeftSlingShot.TimerEnabled = 0':gi3.State = 1:Gi4.State = 1
    End Select
    LStep = LStep + 1
End Sub

Sub RightSlingshot_Slingshot
  AddPoints(1100)
  PlaySound SoundFXDOF("rsling",104,DOFPulse,DOFContactors),0,0.75,0.15,0.25
  DOF 106, DOFPulse
    RSling.Visible = 0
    RSling1.Visible = 1
    sling1.TransZ = -20
    RStep = 0
    RightSlingShot.TimerEnabled = 1
  AdvanceShuttleRampValue
  AwardUSPoints
End Sub

Sub RightSlingShot_Timer
    Select Case RStep
        Case 3:RSLing1.Visible = 0:RSLing2.Visible = 1:sling1.TransZ = -10
        Case 4:RSLing2.Visible = 0:RSLing.Visible = 1:sling1.TransZ = 0:RightSlingShot.TimerEnabled = 0':gi1.State = 1:Gi2.State = 1
    End Select
    RStep = RStep + 1
End Sub

'Spinner
Sub LeftSpinner_spin
  Playsound "W_Spin3",0,1,-0.19,0.25
  DOF 117, DOFPulse
  if Spinneractive then
    if Powerlevel = 1 then
      addpoints(11070)
    else
      addpoints(100000*(Powerlevel-1))
    end if
  end if
End Sub

Sub RightSpinner_spin
  Playsound "W_Spin3",0,1,0.17,0.25
  DOF 118, DOFPulse
  if Spinneractive then
    if Powerlevel = 1 then
      addpoints(11070)
    else
      addpoints(100000*(Powerlevel-1))
    end if
  end if
End Sub

'Magnet
Dim MagnetBall,MagDistance
Sub MagnetGrabTrigger_Hit
if MagnetGrabTrigger.enabled then
  Set MagnetBall = ActiveBall
  ActiveBall.VelX = 0
  ActiveBall.VelY = 0
  ActiveBall.VelZ = 0
  MagnetFieldTimer.enabled = True
  MagnetTimer.enabled = False
  MagnetTimer.enabled = True
  MCount = 0
end if
End Sub

Sub MagnetFieldTimer_Timer
  MagDistance = SQR(((Magnet.x - MagnetBall.x)^2) + ((Magnet.y - MagnetBall.y)^2))
  if MagDistance > 20 then
    MagnetBall.x = MagnetBall.x + ((Magnet.x - MagnetBall.x) * 0.23)
    MagnetBall.y = MagnetBall.y + ((Magnet.y - MagnetBall.y) * 0.23)
  else
    if MagDistance > 5 then
      if MagnetBall.x < Magnet.x then
        MagnetBall.x = MagnetBall.x + (MagDistance * 1.15)
      else
        MagnetBall.x = MagnetBall.x - (MagDistance * 1.15)
      end if
      if MagnetBall.y < Magnet.y then
        MagnetBall.y = MagnetBall.y + (MagDistance * 1.15)
      else
        MagnetBall.y = MagnetBall.y - (MagDistance * 1.15)
      end if
    else
      MagnetBall.x = Magnet.x
        MagnetBall.y = Magnet.y
    end if
  end if
  MagnetBall.VelX = 0
  MagnetBall.VelY = 0
End Sub

Dim MCount
Sub MagnetTimer_Timer
  MCount = MCount + 1
  if MCount = 1 then
    select case int(4*rnd)
      case 0: playsound "W_Launch"
      case 1: playsound "W_Flip"
      case 2: playsound "W_Fire"
      case 3: playsound "W_Now"
    end select
  else
    MagnetTimer.enabled = False
    StopMagnet
  end if
End Sub

Sub StopMagnet
  MagnetGrabTrigger.enabled = False
  MagnetTimer.enabled = False
  MagnetFieldTimer.enabled = False
  For Each obj in MagnetLights:obj.state = lightstateoff:next
End Sub

Sub DeactivateMissionsAndModes
  if ReactorCriticalActive then
    EndReactorCritical
  end if
  if HuntDownActive then
    EndHuntDown
  end if
  if AmmoDumpActive then
    EndAmmoDump
  end if
  if ShuttleActive then
    EndShuttle
  end if
  if BikeRaceActive then
    EndBikeRace
  end if
  if SkyscraperActive then
    EndSkyscraper
  end if

  if QuickshotActive then
    ActualQuickShotValue = 0
    EndQuickshot
  end if
  if SuperLauncherActive then
    EndSuperLauncher
  end if
  if ComboManiaActive then
    EndCombomania
  end if
  if UltimateShowdownActive then
    EndUltimateShowdown
  end if
  if SecretManiaActive then
    EndSecretMania
  end if

  Text1 = FormatDMDTopText("","","")
  Text2 = FormatDMDBottomText("")
  AdvanceLeftPowerlevelLight.state = 0
  AdvanceRightPowerlevelLight.state = 0
End Sub

'--- Helper Functions ---
Sub PassLeftFlasherTrigger_Hit
  FlashLeft 1
End Sub
Sub PassLeftFlasherTrigger2_Hit
  FlashLeft 1
End Sub
Sub PassRightFlasherTrigger_Hit
  FlashRight 1
End Sub

Dim LeftFlashCount
Sub FlashLeft(FLCount)
  if ApplyMods = 1 then
    LeftFlashState = True
    P_LeftFlasher.image = "dome4_red_lit"
    DOF 141, DOFPulse
    LeftFlashCount = LeftFlashCount + FLCount
    FLeft.State = 1
    LeftFlashTimer.enabled = true
  end if
  FlashB2SLamp 48,FLCount
End Sub

Dim LeftFlashState
Sub LeftFlashTimer_Timer
  LeftFlashState = not LeftFlashState
  if not LeftFlashState then
    LeftFlashCount = LeftFlashcount - 1
    P_LeftFlasher.image = "dome4_red"
    FLeft.State = 0
  else
    P_LeftFlasher.image = "dome4_red_lit"
    DOF 141, DOFPulse
    FLeft.State = 1
  end if
  if LeftFlashcount <= 0 then
    LeftFlashTimer.enabled = false
    P_LeftFlasher.image = "dome4_red"
    FLeft.State = 0
  end if
End sub

Dim RightFlashCount
Sub FlashRight(FLCount)
  if ApplyMods = 1 then
    RightFlashState = True
    P_RightFlasher.image = "dome4_red_lit"
    DOF 142, DOFPulse
    RightFlashCount = RightFlashCount + FLCount
    FRight.State = 1
    RightFlashTimer.enabled = true
  end if
  FlashB2SLamp 49,FLCount
End Sub

Dim RightFlashState
Sub RightFlashTimer_Timer
  RightFlashState = not RightFlashState
  if not RightFlashState then
    RightFlashCount = RightFlashcount - 1
    P_RightFlasher.image = "dome4_red"
    FRight.State = 0
  else
    P_RightFlasher.image = "dome4_red_lit"
    DOF 142, DOFPulse
    FRight.State = 1
  end if
  if RightFlashcount <= 0 then
    RightFlashTimer.enabled = false
    P_RightFlasher.image = "dome4_red"
    FRight.State = 0
  end if
End sub

Sub ZTrigger_Hit
  ActiveBall.velz = -1
End Sub

Sub DockRampEnter_Hit
  if not SuperLauncherActive and not HideAnimation then
    DMDDisplayScene "Dock",True,DockRampUpback, "", 8,14,"", 5,15, 14, 500, 14
  end if
End Sub

Sub AddPoints(Count)
  if Tilt = 0 then
    Points = Points + Count
  end if
  if not HideAnimation then
    UpdateDMDScore
  end if
End Sub

Dim PointsText,TempText
Function FormatPointText(ScorePar)
  PointsText = ""
  TempText = ScorePar
  for i = 1 to 10
    if len(temptext) > 3 then
      PointsText = "." & right(TempText,3) & Pointstext
      TempText = left(Temptext,len(TempText)-3)
    else
      if len(Temptext) > 0 then
        PointsText = TempText & Pointstext
        TempText = ""
      end if
    end if
  next
  FormatPointText = PointsText
End Function

Dim JPState,SuperJPState,UltraJPState
Sub JPTimer_Timer
  select case JackpotLight.state
    case 0: P_Jackpotlight.image = "Bulb_Yellow_off"
    case 1: P_Jackpotlight.image = "Bulb_Yellow_on"
    case Lightstateblinking:
      JPState = not JPState
      if JPState then
        P_Jackpotlight.image = "Bulb_Yellow_on"
      else
        P_Jackpotlight.image = "Bulb_Yellow_off"
      end if
  end select
  select case SuperJackpotLight.state
    case 0: P_SuperJackpotlight.image = "Bulb_Yellow_off"
    case 1: P_SuperJackpotlight.image = "Bulb_Yellow_on"
    case Lightstateblinking:
      SuperJPState = not SuperJPState
      if SuperJPState then
        P_SuperJackpotlight.image = "Bulb_Yellow_on"
      else
        P_SuperJackpotlight.image = "Bulb_Yellow_off"
      end if
  end select
  select case UltraJackpotLight.state
    case 0: P_UltraJackpotlight.image = "Bulb_Red_off"
    case 1: P_UltraJackpotlight.image = "Bulb_Red_on"
    case Lightstateblinking:
      UltraJPState = not SuperJPState
      if UltraJPState then
        P_UltraJackpotlight.image = "Bulb_Red_on"
      else
        P_UltraJackpotlight.image = "Bulb_Red_off"
      end if
  end select
End Sub

Sub LRampHelper_Hit
  PlaySound "W_BallDrop",0,1,-0.19,0
End Sub
Sub RRampHelper_Hit
  PlaySound "W_BallDrop",0,1,0.18,0
End Sub


'VUK
Sub VUKFire
  if StartMagnet then
    PlayNextSound "W_MagnetActivated"
    if not Hideanimation then
      DMDDisplayScene "Mag",True,Magnetback, "", 8,14,"", 5,15, 14, 2000, 14
    end if
    if StartQuickDock then
      Docklight.state = lightstateblinking
      DockBallLock.enabled = True
      Docklight.Timerenabled = False
      Docklight.Timerenabled = True
    end if
  end if
  VUKState = 0
  VUKTimer.enabled = True
End Sub

Dim VUKState,WaitTime
WaitTime = 75
Sub VUKTimer_Timer
  VUKState = VukState + 1
  Select case VUKState
    case 1+WaitTime: VUK1.CreateBall
    case 2+WaitTime: VUK1.DestroyBall
             VUK2.CreateBall
    case 3+WaitTime: VUK2.DestroyBall
             VUK3.CreateBall
    case 4+WaitTime: VUK3.DestroyBall
             VUK4.CreateBall
    case 5+WaitTime: VUK4.DestroyBall
             VUK5.CreateBall
    case 6+WaitTime: VUKTimer.enabled = False
             VUK5.DestroyBall
             ExplosionVUK.CreateBall
             ExplosionVUK.kick 275, 15
             playsound SoundFXDOF("W_Vuk2",120,DOFPulse,DOFContactors),0,1,0.05,0.25
             ExplosionHole.enabled = True
  end select
End Sub

'Flipper Primitives
Sub FlipperTimer_Timer
  P_LeftFlipper.Roty = LeftFlipper.CurrentAngle
  P_ULeftFlipper.Roty = ULeftFlipper.CurrentAngle
  P_RightFlipper.roty= RightFlipper.CurrentAngle - 2
End Sub

'--- Ufo animation ---
Dim UFOAnimationCount,YDir,XDir
Sub UfoHitAnimation(DirPar)
  if ApplyMods = 1 then
    YDir = 0.2
    if DirPar = 1 then
      XDir = 0.6
    else
      XDir = -0.6
    end if
    UFOAnimationCount = 0
    UfoHitTimer.enabled = True
  end if
End Sub

Sub UfoHitTimer_Timer
  UFOAnimationCount = UFOAnimationCount + 1
  P_Ufo.TransX = P_Ufo.TransX + XDir
  P_Ufo.TransY = P_Ufo.TransY + YDir
  if (UFOAnimationCount = 7) or (UFOAnimationCount = 13) or (UFOAnimationCount = 18) then
    YDir = -YDir
    XDir = -XDir
  end if
  if UFOAnimationCount >= 23 then
    UfoHitTimer.enabled = False
    P_Ufo.TransX = 0
    P_Ufo.TransY = 0
  end if
End Sub

'--- Ball Trough Simulation ---
Sub TroughLeftVUKKicker_Hit
  TroughLeftVUKKicker.destroyball
  LeftRampVUK.CreateBall
  LeftRampVUK.kick 180, 1
  playsound SoundFXDOF("W_Vuk",134,DOFPulse,DOFContactors),0,1,-0.18,0.25
  DOF 136, DOFPulse
End sub

Sub TroughRightVUKKicker_Hit
  TroughRightVUKKicker.destroyball
  RightRampVUK.CreateBall
  RightRampVUK.kick 195, 1
  playsound SoundFXDOF("W_Vuk2",135,DOFPulse,DOFContactors),0,1,0.18,0.25
  DOF 136, DOFPulse
End sub

'--- Tilt recognition ---
Sub CheckNudge
  if GameActive = 1 then
    if NudgeTimer1.enabled then
      if NudgeTimer2.enabled then
        NudgeTimer1.enabled = False
        NudgeTimer2.enabled = False
        if Tilt = 0 then
          DeactivateMissionsandmodes
          EndCurrentMusic
          DMDDisplayScene "Tilt",True,"blank.png", "TILT", 5,15,"", 5,15, 14, 2500, 14
          if int(2*rnd) = 0 then
            PlaySound "W_SensorOverload"
          else
            PlaySound "W_YouCantWin"
          end if
        end if
        Tilt = 1
        GracePeriod.enabled = False
        GracePeriod1.enabled = False
        GracePeriod2.enabled = False
      else
        NudgeTimer2.enabled = True
        if Tilt = 0 then
          UpdateDMDText "Main","","DANGER","",""
          PlaySound "W_Easy"
        end if
      end if
    else
      NudgeTimer1.enabled = True
      if Tilt = 0 then
        UpdateDMDText "Main","","WARNING","",""
        PlaySound "W_WatchIt"
      end if
    end if
  end if
End Sub

Sub NudgeTimer1_Timer
  NudgeTimer1.enabled = False
End Sub

Sub NudgeTimer2_Timer
  NudgeTimer2.enabled = False
End Sub


'--- End of Ball Bonus Calculation ---
Dim BonusCount,ReleaseBall,TotalBonus,SkipBonus,TotalLoopCount,TotalPowerLevelCount,TotalCombocount
Sub AddBonus
  if Tilt = 0 then
    UltraDMDTextTimer.enabled = True
    UpdateDMDText "Main","","End of Ball","","Bonus"
    Playnextmusic(BonusTheme)
    BonusCount = 0
    SkipBonus = 0
    TotalBonus = (TotalLoopCount * 400000) + (TotalPowerLevelCount * 500000) + (TotalCombocount * 600000)
    AddBonusTimer.enabled = True
  else
    if BallInPlay <= BallsperGame then
      EnableIdleSound = True
      NextBall
    else
      EndOfGame
    end if
  end if
End Sub

Sub AddBonusTimer_Timer
  AddBonusTimer.enabled = False
  DisplayNextBonus.enabled = True
end sub

Sub DisplayNextBonus_Timer
  BonusCount = BonusCount + 1
  if (SkipBonus = 1) and (BonusCount < 5) then
    BonusCount = 5
  end if
  if (BonusCount < 6) then
    select case Bonuscount
      case 1: UpdateDMDText "Main","",cstr(TotalLoopCount) & " Loops","",FormatPointText(cstr(TotalLoopcount * 400000))
      case 2: UpdateDMDText "Main","",cstr(TotalPowerLevelCount) & " PowerLevels","",FormatPointText(cstr(TotalPowerLevelCount * 500000))
      case 3: UpdateDMDText "Main","",cstr(TotalComboCount) & " Combos","",FormatPointText(cstr(TotalComboCount * 600000))
      case 4: UpdateDMDText "Main","","Total Bonus x" + cstr(BonusMultiplier),"",FormatPointText(cstr(TotalBonus))
      case 5: UpdateDMDText "Main","","Total  Bonus","",FormatPointText(cstr(TotalBonus * BonusMultiplier))
    end select
  else
    Addpoints(TotalBonus * BonusMultiplier)
    DisplayNextBonus.enabled = False
    UltraDMDTextTimer.enabled = False
    UpdateDMDScore
    if BallInPlay <= BallsperGame then
      EnableIdleSound = True
      NextBall
    else
      EndOfGame
    end if
  end if
End Sub


'--- Highscore Entry ---
dim Chars,EnterHS,CharIndex,InitialPos,InitialsText,C(3),HighscoreText
Sub EnterInitials
  if (points > Highscore(3)) or (TotalComboCount > Highscore(4)) or (TotalLoopCount > Highscore(5)) then
    EnterHS = True
    playnextmusic(FastFrenzyTheme)
    if int(2*rnd) = 1 then
      PlayNextSound "W_Unbelievable"
    else
      PlayNextSound "W_GreatScore"
    end if
    IF points > Highscore(1) then
      HighscoreText = "Highscore  1"
    else
      IF points > Highscore(2) then
        HighscoreText = "Highscore  2"
      else
        IF points > Highscore(3) then
          HighscoreText = "Highscore  3"
        else
          if TotalComboCount > Highscore(4) then
            HighscoreText = "Combo Champion"
          else
            if TotalLoopCount > Highscore(5) then
              HighscoreText = "Loop  Champion"
            end if
          end if
        end if
      end if
    end if
    Chars = array("A","B","C","D","E","F","G","H","I","J","K","L","M","N","O","P","Q","R","S","T","U","V","W","X","Y","Z","0","1","2","3","4","5","6","7","8","9","."," ","<")  '0-38
    CharIndex = 0
    C(1) = "A"
    C(2) = " "
    C(3) = " "
    InitialPos = 1
    InitialsText = C(1) + C(2) + C(3)
    If UltraDMD.IsRendering Then
      UltraDMD.CancelRendering
    end if
    DMDDisplayScene "Main",True,backgrnd, FormatDMDTopText("",HighscoreText,""), 8,14,FormatDMDBottomText(InitialsText), 5,15, 14, 10000000, 14
  else
    PlayNextSound "W_GameOver"
    PlayHighscoremusic
    StartAttractMode
  end if
end sub

Sub WriteHS
  select case int(5*rnd)
    case 0: playnextsound "W_Superb"
    case 1: playnextsound "W_Outstanding"
    case 2: playnextsound "W_VeryGood"
    case 3: playnextsound "W_YesYesYes"
    case 4: PlaynextSound "W_Yeehaa"
  end select
  PlaynextSound SoundFXDOF("Knocker",140,DOFPulse,DOFKnocker)
  DOF 136, DOFPulse
  if TotalLoopCount > Highscore(5) then
    Highscore(5) = TotalLoopCount
    Highname(5) = InitialsText
  end if
  if TotalComboCount > Highscore(4) then
    Highscore(4) = TotalComboCount
    Highname(4) = InitialsText
  end if
  if Points > Highscore(1) then
    Highscore(3) = Highscore(2)
    Highname(3) = Highname(2)
    Highscore(2) = Highscore(1)
    Highname(2) = Highname(1)
    Highscore(1) = points
    Highname(1) = InitialsText
  else
    if Points > Highscore(2) then
      Highscore(3) = Highscore(2)
      Highname(3) = Highname(2)
      Highscore(2) = points
      Highname(2) = InitialsText
    else
      Highscore(3) = points
      Highname(3) = InitialsText
    end if
  end if
  savevalue "TheWeb","Name1",HighName(1)
  savevalue "TheWeb","High1",HighScore(1)
  savevalue "TheWeb","Name2",HighName(2)
  savevalue "TheWeb","High2",HighScore(2)
  savevalue "TheWeb","Name3",HighName(3)
  savevalue "TheWeb","High3",HighScore(3)
  savevalue "TheWeb","Name4",HighName(4)
  savevalue "TheWeb","High4",HighScore(4)
  savevalue "TheWeb","Name5",HighName(5)
  savevalue "TheWeb","High5",HighScore(5)
  ResetHSDisplay.enabled = True
End sub

Sub ResetHSDisplay_Timer
  ResetHSDisplay.enabled = False
  enterhs = False
  DMDDisplayScene "Main",True,backgrnd, FormatDMDTopText("","Pro  Pinball",""), 8,14,FormatDMDBottomText("Highscores"), 5,15, 14, 100000, 14
  PlayHighscoremusic
  HSPos = 2
  IdleTimer.enabled = True
  StartAttractMode
End Sub

Sub ResetHighscore
  savevalue "TheWeb","Name1","BAL"
  savevalue "TheWeb","High1",200000000
  savevalue "TheWeb","Name2","PIN"
  savevalue "TheWeb","High2",100000000
  savevalue "TheWeb","Name3","PRO"
  savevalue "TheWeb","High3",5000000
  savevalue "TheWeb","Name4","COM"
  savevalue "TheWeb","High4",1
  savevalue "TheWeb","Name5","LOO"
  savevalue "TheWeb","High5",1
End Sub

Dim HSPos
Sub IdleTimer_Timer
  if not EnterHS then
    HSPos = HSPos + 1
    if HSPos < 1 or HSPos > 9 then HSPos = 1
    select case HSPos
      case 1,2: UpdateDMDText "Main","","Pro  Pinball","","The  Web"
      case 3: UpdateDMDText "Main","","Pro  Pinball","","Highscores"
      case 4: UpdateDMDText "Main","","3. "&HighName(3),"",FormatPointText(cstr(HighScore(3)))
      case 5: UpdateDMDText "Main","","2. "&HighName(2),"",FormatPointText(cstr(HighScore(2)))
      case 6: UpdateDMDText "Main","","1. "&HighName(1),"",FormatPointText(cstr(HighScore(1)))
      case 7: UpdateDMDText "Main","","Combos "&HighName(4),"",FormatPointText(cstr(HighScore(4)))
      case 8: UpdateDMDText "Main","","Loops  "&HighName(5),"",FormatPointText(cstr(HighScore(5)))
      case 9: UpdateDMDText "Main","","GAME  OVER","","PRESS  START"
    end select
  end if
end sub

'--- Keyboard functions ---
Dim RightFlipperUp,LeftFlipperUp
Sub TheWeb_KeyDown(ByVal keycode)
  if enterhs then
    If (keycode = PlungerKey) or (keycode = StartGameKey) Then
      if CharIndex = 38 then
        playsound "W_HSBack"
        C(Initialpos) = " "
        initialpos = initialpos - 1
        if initialpos < 1 then initialpos = 1
        C(Initialpos) = "A"
        CharIndex = 0
      else
        playsound "W_HSEnter"
        initialpos = initialpos + 1
        if initialpos > 3 then
          enterhs = False
          WriteHS
        else
          C(Initialpos) = "A"
          CharIndex = 0
        end if
      end if
    end if
    If keycode = LeftFlipperKey Then
      playsound "W_HSLetter"
      CharIndex = CharIndex - 1
      if CharIndex < 0 then
        Charindex = 38
      end if
      C(Initialpos) = Chars(Charindex)
    end if
    If keycode = RightFlipperKey Then
      playsound "W_HSLetter"
      CharIndex = CharIndex + 1
      if CharIndex > 38 then
        Charindex = 0
      end if
      C(Initialpos) = Chars(Charindex)
    end if
    InitialsText = C(1) + C(2) + C(3)
    UpdateDMDText "Main","",HighscoreText,"",InitialsText
  else
    if VideoModeActive then
      If keycode = LeftFlipperKey Then
        ShootVMItem "Left"
      end if
      If keycode = RightFlipperKey Then
        ShootVMItem "Right"
      end if
    else
      if DisplayNextBonus.enabled then SkipBonus = 1                         'skip Bonus calculation

      if keycode = StartGameKey then
        if not UltraDMDInitTimer.enabled then
          StartGame
        end if
      end if
      If keycode = PlungerKey Then
        if FrenzyPickTimer.enabled then
          FastFrenzyPickMade
        else
          Plunger.Fire
          DOF 132, DOFPulse
        end if
      end if
      if (Tilt = 0) and (GameActive = 1) and not InitUSTimer.enabled and not USWonTimer.enabled then
        If keycode = LeftFlipperKey Then
          if FrenzyPickTimer.enabled then
            ChangePick(-1)
          else
            LeftFlipperUp = True
            StopMagnet
            LeftFlipper.RotateToEnd
            ULeftFlipper.RotateToEnd
            if EnableIdleSound then
              PlaySound "W_IdleLeft",0,1,-0.12,0.25
            else
              PlaySound SoundFXDOF("FlipperUp1",101,DOFOn,DOFContactors),0,0.5,-0.12,0.25
            end if
            ChangeBonusLight(1)
          end if
        End If
        If keycode = RightFlipperKey Then
          if FrenzyPickTimer.enabled then
            ChangePick(1)
          else
            RightFlipperUp = True
            RightFlipper.RotateToEnd
            if EnableIdleSound then
              PlaySound "W_IdleRight",0,1,0.12,0.25
            else
              PlaySound SoundFXDOF("FlipperUp1",102,DOFOn,DOFContactors),0,0.5,0.12,0.25
            end if
            ChangeBonusLight(1)
          end if
        End If
      end if
      If keycode = LeftTiltKey Then:Nudge 90, 2:checkNudge:end if
      If keycode = RightTiltKey Then:Nudge 270, 2:checkNudge:end if
      If keycode = CenterTiltKey Then:Nudge 0, 2:checkNudge:end if
    end if
  end if
End Sub

Sub TheWeb_KeyUp(ByVal keycode)
  If keycode = LeftFlipperKey Then
    if LeftFlipperUp then
      LeftFlipper.RotateToStart
      ULeftFlipper.RotateToStart
      PlaySound SoundFXDOF("FlipperDown",101,DOFOff,DOFContactors),0,0.75,-0.1,0.25
    end if
    LeftFlipperUp = False
  End If
  If keycode = RightFlipperKey Then
    if RightFlipperUp then
      RightFlipper.RotateToStart
      PlaySound SoundFXDOF("FlipperDown",102,DOFOff,DOFContactors),0,0.75,0.1,0.25
    end if
    RightFlipperUp = False
  End If
  if not enterhs then
    If keycode = PlungerKey Then
      Plunger.Timerenabled = True
    end if
  end if
End Sub

Sub Plunger_Timer
  Plunger.Timerenabled = False
  Plunger.PullBack
End Sub


'---- Sound Functions ----
Dim QueueIndex,SoundQueue(5)

Sub PlayNextSound(SoundNamePar)         'used to avoid unwanted sounds playing at the same time
  if PlaySoundTimer.enabled = False then
    SetPlaySoundTimer(SoundNamePar)
    PlaySoundTimer.enabled = True
    Playsound SoundNamePar
  else
    if QueueIndex < 5 then          'only 5 sounds will be queued
      QueueIndex = QueueIndex + 1
      SoundQueue(QueueIndex) = SoundNamePar
      ProcessSoundQueue.enabled = True
    end if
  end if
end sub

Sub SetPlaySoundTimer(SoundNamePar)
  PlaySoundTimer.interval = 1000        'most sounds are shorter than 1 second
  select case SoundnamePar
    case "W_DockIsLit":     PlaySoundTimer.interval = 2200
    case "W_Ball1Locked":     PlaySoundTimer.interval = 2200
    case "W_Ball2Locked":   PlaySoundTimer.interval = 2200
    case "W_Ball3Locked":     PlaySoundTimer.interval = 2200
    case "W_VideoMode":     PlaySoundTimer.interval = 2200
    case "W_Ball1Docked":     PlaySoundTimer.interval = 2200
    case "W_SpaceStation":    PlaySoundTimer.interval = 2200
    case "W_Jackpot":       PlaySoundTimer.interval = 2200
    case "W_SuperJackpotLit": PlaySoundTimer.interval = 2200
    case "W_ExtraBall":     PlaySoundTimer.interval = 2200
    case "W_GameOver":      PlaySoundTimer.interval = 2200
    case "W_SphereSpin":    PlaySoundTimer.interval = 2200
    case "W_SuperJackpot":    PlaySoundTimer.interval = 2200
    case "W_Yeehaa":      PlaySoundTimer.interval = 2200
    case "W_ShuttleDestroyed":  PlaySoundTimer.interval = 2200
    case "W_SuperCombo":    PlaySoundTimer.interval = 2200
    case "W_ShootSpinners":   PlaySoundTimer.interval = 2200
    case "W_No":        PlaySoundTimer.interval = 2200
    case "W_Comon":       PlaySoundTimer.interval = 2200
    case "W_Stop":        PlaySoundTimer.interval = 2200
    case "W_Cow":       PlaySoundTimer.interval = 2200
    case "W_GetUltraJackpot": PlaySoundTimer.interval = 2200
    case "W_UltimateShowdown":  PlaySoundTimer.interval = 3200
    case "W_UltraJackpot":    PlaySoundTimer.interval = 3200
    case "W_YesYesYes":     PlaySoundTimer.interval = 3200
    case "W_YouMadeIt":     PlaySoundTimer.interval = 4200
    case "W_YouHaventSeen":   PlaySoundTimer.interval = 2200
    case "W_USIntitBoomLast": PlaySoundTimer.interval = 2200
    case "W_USOneBillion":    PlaySoundTimer.interval = 6200
    case "W_USWon":       PlaySoundTimer.interval = 2200
    case "W_USWon2":      PlaySoundTimer.interval = 22200
    case "W_Yeehaa":      PlaySoundTimer.interval = 2200
    case "W_SuperJackpot":    PlaySoundTimer.interval = 2200
    case "W_OhNo":        PlaySoundTimer.interval = 2200
  end select
end sub

Sub PlaySoundTimer_Timer
  PlaySoundTimer.enabled = False
end sub

Sub ProcessSoundQueue_Timer
  if QueueIndex > 0 then
    if PlaySoundTimer.enabled = False then
      SetPlaySoundTimer(SoundQueue(1))
      PlaySoundTimer.enabled = True
      Playsound SoundQueue(1)
      SoundQueue(1) = SoundQueue(2)
      SoundQueue(2) = SoundQueue(3)
      SoundQueue(3) = SoundQueue(4)
      SoundQueue(4) = SoundQueue(5)
      SoundQueue(5) = ""
      QueueIndex = QueueIndex - 1
    end if
  end if
  if QueueIndex = 0 then ProcessSoundQueue.enabled = False
End Sub


'---- Music Functions ----
Const ComboManiaTheme = "bgout_The Web 01 Combomania.mp3"
Const ComboManiaThemeLen = 100000           'actually 45500 but it should not be repeated
Const SuperLauncherTheme = "bgout_The Web 02 Super Launcher.mp3"
Const SuperLauncherThemeLen = 69500
Const QuickShotTheme = "bgout_The Web 03 Quick Shot.mp3"
Const QuickShotThemeLen = 100000            'actually 21000 but it should not be repeated
Const ShuttleTheme = "bgout_The Web 04 Mission - Shuttle.mp3"
Const ShuttleThemeLen = 100000            'actually 47000 but it should not be repeated
Const AmmoDumpTheme = "bgout_The Web 05 Mission - Ammo Dump.mp3"
Const AmmoDumpThemeLen = 100000           'actually 50000 but it should not be repeated
Const HuntDownTheme = "bgout_The Web 06 Mission - Hunt Down.mp3"
Const HuntDownThemeLen = 100000           'actually 62000 but it should not be repeated
Const BikeRaceTheme = "bgout_The Web 07 Mission - Biker.mp3"
Const BikeRaceThemeLen = 100000           'actually 66000 but it should not be repeated
Const SkyscraperTheme = "bgout_The Web 08 Mission - Sky Scraper.mp3"
Const SkyscraperThemeLen = 100000         'actually 59500 but it should not be repeated
Const ReactorTheme = "bgout_The Web 09 Mission - Reactor.mp3"
Const ReactorThemeLen = 100000            'actually 50000 but it should not be repeated
Const MainTheme = "bgout_The Web 10 Main Theme.mp3"
Const MainThemeLen = 172500
Const FastFrenzyTheme = "bgout_The Web 11 Fast Frenzy.mp3"
Const FastFrenzyThemeLen = 119500
Const IdleTheme = "bgout_The Web 12 Menu Theme.mp3"   'enter name of the theme
Const IdleThemeLen = 61000                          'length in ms + 500 for a gap, the active song will be restarted after this time
Const SpaceStationTheme = "bgout_The Web 13 Space Station Frenzy.mp3"
Const SpaceStationThemeLen = 125500
Const SpaceStationTheme2 = "bgout_The Web 14.mp3"
Const SpaceStationTheme2Len = 12000
Const VideoModeTheme = "bgout_The Web 15 Video Mode.mp3"
Const VideoModeThemeLen = 31000
Const SecretManiaTheme = "bgout_The Web 16 Secret Mania.mp3"
Const SecretManiaThemeLen = 71500
Const BonusTheme = "bgout_The Web 17 Bonus Theme.mp3"
Const BonusThemeLen = 100000                        'actually 12000 but it should not be repeated
Const HighscoreTheme = "bgout_The Web 18 Highscore.mp3"
Const HighscoreThemeLen = 16000
Const UltimateShowdownTheme = "bgout_The Web 19 Ultimate Showdown.mp3"
Const UltimateShowdownThemeLen = 163500
Const VideoModeTheme2 = "bgout_The Web 20.mp3"
Const VideoModeTheme2Len = 31000

Dim CurrentSong
Sub PlayNextMusic(MusicPar)
  if MusicActive = 1 then
    if MusicPar <> Currentsong then
      CurrentSong = MusicPar
      SetPlayMusicTimer(MusicPar)
      EndMusic
      Playmusic MusicPar
      PlayMusicTimer.enabled = True
    end if
  end if
end sub

Sub PlayHighscoremusic
  if MusicActive = 1 then
    CurrentSong = HighscoreTheme
    EndMusic
    Playmusic HighscoreTheme
    PlayMusicTimer2.interval = HighscoreThemeLen - 100
    PlayMusicTimer2.enabled = True
  end if
End sub

Sub SetPlayMusicTimer(PlayMusicPar)
  select case Playmusicpar
    case MainTheme:  PlayMusicTimer.interval = MainThemeLen
    case IdleTheme:  PlayMusicTimer.interval = IdleThemeLen
    case BonusTheme:  PlayMusicTimer.interval = BonusThemeLen
    case ComboManiaTheme:  PlayMusicTimer.interval = ComboManiaThemeLen
    case SuperLauncherTheme:  PlayMusicTimer.interval = SuperLauncherThemeLen
    case QuickShotTheme:  PlayMusicTimer.interval = QuickShotThemeLen
    case SpaceStationTheme:  PlayMusicTimer.interval = SpaceStationThemeLen
    case SpaceStationTheme2:  PlayMusicTimer.interval = SpaceStationTheme2Len
    case FastFrenzyTheme:  PlayMusicTimer.interval = FastFrenzyThemeLen
    case ReactorTheme:  PlayMusicTimer.interval = ReactorThemeLen
    case HuntDownTheme:  PlayMusicTimer.interval = HuntDownThemeLen
    case AmmoDumpTheme:  PlayMusicTimer.interval = AmmoDumpThemeLen
    case ShuttleTheme:  PlayMusicTimer.interval = ShuttleThemeLen
    case BikeRaceTheme:  PlayMusicTimer.interval = BikeRaceThemeLen
    case SkyscraperTheme:  PlayMusicTimer.interval = SkyscraperThemeLen
    case SecretManiaTheme:  PlayMusicTimer.interval = SecretManiaThemeLen
    case UltimateShowdownTheme:  PlayMusicTimer.interval = UltimateShowdownThemeLen
    case HighscoreTheme:  PlayMusicTimer.interval = HighscoreThemeLen
    case VideoModeTheme:  PlayMusicTimer.interval = VideoModeThemeLen
    case VideoModeTheme2:  PlayMusicTimer.interval = VideoModeTheme2Len
  end select
end sub

Dim QueuedMusic,IdleCount,MysteryThemeplayed
QueuedMusic = ""
Sub PlayMusicTimer_Timer
  EndMusic
  if QueuedMusic <> "" then
    CurrentSong = QueuedMusic
    SetPlayMusicTimer(QueuedMusic)
    QueuedMusic = ""
  end if
  if not idletimer.enabled then
    Playmusic CurrentSong
  else
    IdleCount = IdleCount + 1
    if IdleCount > 12 then
      PlayMusicTimer.interval = IdleThemeLen
      IdleCount=0
      MysteryThemeplayed=False
      Playmusic CurrentSong
    else
      if MysteryThemeplayed then
        MysteryThemeplayed = False
        PlayMusicTimer.interval = 10000
      else
        if IdleCount > 1 then
          MysteryThemeplayed = True
          select case int(rnd*12)
            case 0: Playmusic QuickShotTheme:PlayMusicTimer.interval = 21000
            case 1: Playmusic SpaceStationTheme2:PlayMusicTimer.interval = 12000
            case 2: Playmusic BonusTheme:PlayMusicTimer.interval = 12000
            case 3: Playmusic HighscoreTheme:PlayMusicTimer.interval = 16000
            case 4: Playmusic VideoModeTheme2:PlayMusicTimer.interval = 31000
            case 5: Playmusic ComboManiaTheme:PlayMusicTimer.interval = 45500
            case 6: Playmusic SkyscraperTheme:PlayMusicTimer.interval = 59500
            case 7: Playmusic ReactorTheme:PlayMusicTimer.interval = 50000
            case 8: Playmusic ShuttleTheme:PlayMusicTimer.interval = 47000
            case 9: Playmusic AmmoDumpTheme:PlayMusicTimer.interval = 50000
            case 10: Playmusic VideoModeTheme:PlayMusicTimer.interval = 31000
            case 11: Playmusic IdleTheme:IdleCount=0:PlayMusicTimer.interval=61000:MysteryThemeplayed=False
          end select
        end if
      end if
    end if
  end if
end sub

Sub PlayMusicTimer2_Timer
  PlayMusicTimer2.enabled = False
  if GameActive = 0 then
    PlayNextMusic(IdleTheme)
  end if
end sub

Sub EndCurrentMusic
  PlayMusicTimer.enabled = False
  EndMusic
  CurrentSong = ""
End Sub



'----------------------
'------ UltraDMD ------
'----------------------

Const UltraDMD_VideoMode_Stretch = 0
Const UltraDMD_VideoMode_Top = 1
Const UltraDMD_VideoMode_Middle = 2
Const UltraDMD_VideoMode_Bottom = 3


'Const UltraDMD_Animation_FadeIn = 0
Const UltraDMD_Animation_FadeOut = 1
'Const UltraDMD_Animation_ZoomIn = 2
Const UltraDMD_Animation_ZoomOut = 3
Const UltraDMD_Animation_ScrollOffLeft = 4
Const UltraDMD_Animation_ScrollOffRight = 5
Const UltraDMD_Animation_ScrollOnLeft = 6
Const UltraDMD_Animation_ScrollOnRight = 7
Const UltraDMD_Animation_ScrollOffUp = 8
Const UltraDMD_Animation_ScrollOffDown = 9
Const UltraDMD_Animation_ScrollOnUp = 10
Const UltraDMD_Animation_ScrollOnDown = 11
Const UltraDMD_Animation_None = 14

Dim UltraDMD,curDir,fso,imgList
Dim Reactorback,HuntDownback,Ammodumpback,BikeRaceback,Skyback1,Skyback2,Shuttleback,Powerlevelback,BallSavedback,DockRampUpback,BallRescueback,ShootAgainback
Dim Sphereback,Spaceback1,Spaceback2,Explosionback,SLAback,SLBback,SLCback,SLDback,BallDockedback,DockLitback,UltimateShowdownBack,SSJackback,SSSJackback,SSUJackback
Dim Cowback,ComboManiaback,SecretManiaback,ExtraBallback,SSJack2back,SSJack3back,EBLitback,RescueLitback,MagnetLitback,LockLitback,LoopAback,LoopBback,Magnetback
Dim B2back,B4back,B6back,B8back,B10back,BMaxback,FFback,JPaback,JPbback,JPcback,Combo1back,Combo2back,Combo3back,DComboback,TComboback,SComboback

Dim ParamList, TempFile
Sub CheckImageList(ListPar)
  ParamList = Split(ListPar,",")
  for each TempFile in ParamList
    TempFile = curDir & "\The Web.UltraDMD\" & TempFile
    Set fso = CreateObject("Scripting.FileSystemObject")
    fso.FileExists(curDir & "\The Web.UltraDMD\" & TempFile)
    If Err Then MsgBox "UltraDMD file not found: " & TempFile
  next
End Sub

Sub LoadUltraDMD
    Set UltraDMD = CreateObject("UltraDMD.DMDObject")
    UltraDMD.Init
    Set fso = CreateObject("Scripting.FileSystemObject")
    curDir = fso.GetAbsolutePathName(".")
    Set fso = nothing
    ' A Major version change indicates the version is no longer backward compatible
    If Not UltraDMD.GetMajorVersion = 1 Then
        MsgBox "Incompatible Version of UltraDMD found."
        Exit Sub
    End If
    'A Minor version change indicates new features that are all backward compatible
    If UltraDMD.GetMinorVersion < 1 Then
        MsgBox "Incompatible Version of UltraDMD found.  Please update to version 1.1 or newer."
        Exit Sub
    End If
    UltraDMD.SetProjectFolder curDir & "\The Web.UltraDMD"
    UltraDMD.SetVideoStretchMode UltraDMD_VideoMode_Middle

  'Init Animations
    imgList = "Idle1.png,Idle2.png,Idle3.png,Idle4.png"
  CheckImageList(ImgList)
    backgrnd = UltraDMD.CreateAnimationFromImages(12, true, imgList)

  imgList = "Mis_R_1.png,Mis_R_2.png,Mis_R_3.png,Mis_R_4.png,Mis_R_5.png,Mis_R_6.png,Mis_R_7.png,Mis_R_8.png"
  CheckImageList(ImgList)
  Reactorback = UltraDMD.CreateAnimationFromImages(8, true, imgList)

  imgList = "M_H_1.png,M_H_2.png,M_H_3.png,M_H_4.png,M_H_5.png,M_H_6.png,M_H_7.png,M_H_8.png,M_H_9.png,M_H_10.png,M_H_11.png,M_H_12.png"
  CheckImageList(ImgList)
  HuntDownback = UltraDMD.CreateAnimationFromImages(6, true, imgList)

  Ammodumpback = "Ammodump.png"
  CheckImageList("Ammodump.png")

  imgList = "M_B_1.png,M_B_2.png,M_B_3.png,M_B_4.png,M_B_5.png,M_B_6.png,M_B_7.png,M_B_8.png"
  CheckImageList(ImgList)
  BikeRaceback = UltraDMD.CreateAnimationFromImages(8, true, imgList)

  imgList = "Sky1.png,Sky1.png,Sky2.png,Sky3.png,Sky4.png,Sky5.png,Sky6.png,Sky7.png,Sky8_5.png,Sky8_5.png,Sky8_5.png,Sky8_5.png,Sky8_5.png"
  CheckImageList(ImgList)
  Skyback1 = UltraDMD.CreateAnimationFromImages(8, False, imgList)
  imgList = "Sky9.png,Sky9.png,Sky10.png,Sky11.png,Sky12.png,Sky13.png,Sky14.png,Sky15.png,Sky16.png,Sky17.png,Sky18.png"
  CheckImageList(ImgList)
  Skyback2 = UltraDMD.CreateAnimationFromImages(8, False, imgList)

  imgList = "Sh1.png,Sh2.png,Sh3.png,Sh4.png,Sh5.png,Sh6.png,Sh7.png,Sh8.png,Sh9.png,Sh10.png"
  CheckImageList(ImgList)
  Shuttleback = UltraDMD.CreateAnimationFromImages(8, True, imgList)

  imgList = "Show1.png,Show2.png,Show3.png,Show4.png,Show5.png,Show6.png,Show7.png,Show8.png,Show9.png,Show10.png,Show11.png"
  CheckImageList(ImgList)
  Sphereback = UltraDMD.CreateAnimationFromImages(12, true, imgList)

  imgList = "PL1.png,PL2.png,PL3.png,PL4.png,PL5.png,PL6.png"
  CheckImageList(ImgList)
  Powerlevelback = UltraDMD.CreateAnimationFromImages(16, true, imgList)

  imgList = "Ball1.png,Ball1.png,Ball2.png,Ball3.png,Ball4.png,Ball5.png,Ball6.png,Ball7.png,Ball8.png"
  CheckImageList(ImgList)
  BallSavedback = UltraDMD.CreateAnimationFromImages(8, false, imgList)

  imgList = "BR1.png,BR1.png,BR2.png,BR3.png,BR4.png,BR5.png,BR6.png,BR7.png,BR8.png,BR9.png,BR10.png,BR11.png,BR12.png,BR13.png"
  CheckImageList(ImgList)
  BallRescueback = UltraDMD.CreateAnimationFromImages(10, false, imgList)

  imgList = "Shoot1.png,Shoot1.png,Shoot2.png,Shoot3.png,Shoot4.png,Shoot5.png,Shoot6.png,Shoot7.png,Shoot8.png,Shoot9.png,Shoot10.png,Shoot11.png,Shoot12.png,Shoot13.png,Shoot14.png,Shoot15.png"
  CheckImageList(ImgList)
  ShootAgainback = UltraDMD.CreateAnimationFromImages(10, false, imgList)

  imgList = "Dock1.png,Dock1.png,Dock2.png,Dock3.png,Dock4.png,Dock5.png,Dock6.png"
  CheckImageList(ImgList)
  DockRampUpback = UltraDMD.CreateAnimationFromImages(8, false, imgList)

  imgList = "Dock10.png,Dock10.png,Dock11.png,Dock12.png,Dock13.png,Dock14.png,Dock15.png,Dock16.png,Dock17.png,Dock18.png,Dock19.png,Dock20.png,Dock21.png,Dock22.png,Dock23.png,Dock24.png,Dock25.png,Dock26.png"
  CheckImageList(ImgList)
  BallDockedback = UltraDMD.CreateAnimationFromImages(10, false, imgList)

  imgList = "DockLit1.png,DockLit1.png,DockLit2.png,DockLit3.png,DockLit4.png,DockLit5.png,DockLit6.png,DockLit7.png,DockLit8.png,DockLit9.png,DockLit10.png,DockLit11.png,DockLit12.png,DockLit13.png,DockLit14.png,DockLit15.png,DockLit16.png,DockLit17.png,DockLit18.png,DockLit19.png"
  CheckImageList(ImgList)
  DockLitback = UltraDMD.CreateAnimationFromImages(11, false, imgList)

  imgList = "STa1.png,STa1.png,STa2.png,STa3.png,STa4.png,STa5.png,STa6.png,STa7.png,STa8.png,STa9.png"
  CheckImageList(ImgList)
  Spaceback1 = UltraDMD.CreateAnimationFromImages(8, False, imgList)
  imgList = "STb1.png,STb2.png,STb3.png,STb4.png,STb4.png,STb3.png,STb2.png,STb1.png"
  CheckImageList(ImgList)
  Spaceback2 = UltraDMD.CreateAnimationFromImages(8, True, imgList)

  imgList = "Ex1.png,Ex1.png,Ex2.png,Ex3.png,Ex4.png,Ex5.png,Ex6.png,Ex7.png,Ex8.png,Ex9.png,Ex10.png,Ex11.png,Ex12.png"
  CheckImageList(ImgList)
  Explosionback = UltraDMD.CreateAnimationFromImages(10, false, imgList)

  imgList = "SLa1.png,SLa2.png,SLa3.png,SLa4.png,SLa5.png,SLa6.png"
  CheckImageList(ImgList)
  SLAback = UltraDMD.CreateAnimationFromImages(6, True, imgList)
  imgList = "SLb1.png,SLb1.png,SLb2.png,SLb3.png,SLb4.png,SLb5.png,SLb6.png,SLb7.png"
  CheckImageList(ImgList)
  SLBback = UltraDMD.CreateAnimationFromImages(7, False, imgList)
  imgList = "SLc1.png,SLc1.png,SLc2.png,SLc3.png,SLc4.png,SLc5.png,SLc6.png,SLc7.png,SLc8.png,SLc9.png,SLc10.png,SLc11.png,SLc12.png,SLc13.png"
  CheckImageList(ImgList)
  SLCback = UltraDMD.CreateAnimationFromImages(1, False, imgList)
  imgList = "SLd1.png,SLd1.png,SLd2.png,SLd3.png,SLd4.png,SLd5.png,SLd6.png,SLd7.png,SLd8.png"
  CheckImageList(ImgList)
  SLDback = UltraDMD.CreateAnimationFromImages(6, False, imgList)

  imglist = "SSJack1.png,SSJack1.png,SSJack2.png,SSJack3.png,SSJack4.png,SSJack5.png,SSJack6.png,SSJack7.png,SSJack8.png,SSJack9.png,SSJack10.png,SSJack11.png,SSJack12.png,SSJack13.png,SSJack14.png"
  imgList = imgList & ",JP1.png,JP2.png,JP3.png,JP4.png,JP5.png,JP6.png,JP7.png,JP8.png,JP9.png,JP10.png,JP11.png,JP12.png"
  CheckImageList(ImgList)
  SSJackback = UltraDMD.CreateAnimationFromImages(8, False, imgList)

  imglist = "SSJack1.png,SSJack1.png,SSJack2.png,SSJack3.png,SSJack4.png,SSJack5.png,SSJack6.png,SSJack7.png,SSJack8.png,SSJack9.png,SSJack10.png,SSJack11.png,SSJack12.png,SSJack13.png,SSJack14.png"
  imgList = imgList & ",DJP1.png,DJP1.png,DJP2.png,DJP3.png,DJP4.png,DJP5.png,DJP6.png,DJP7.png,DJP8.png,DJP9.png,DJP10.png,DJP11.png"
  CheckImageList(ImgList)
  SSJack2back = UltraDMD.CreateAnimationFromImages(8, False, imgList)

  imglist = "SSJack1.png,SSJack1.png,SSJack2.png,SSJack3.png,SSJack4.png,SSJack5.png,SSJack6.png,SSJack7.png,SSJack8.png,SSJack9.png,SSJack10.png,SSJack11.png,SSJack12.png,SSJack13.png,SSJack14.png"
  imgList = imgList & ",TJP1.png,TJP2.png,TJP3.png,TJP4.png,TJP5.png,TJP6.png,TJP7.png,TJP8.png,TJP9.png,TJP10.png,TJP11.png,TJP12.png,TJP13.png,TJP14.png,TJP15.png,TJP16.png"
  CheckImageList(ImgList)
  SSJack3back = UltraDMD.CreateAnimationFromImages(8, False, imgList)

  imglist = "SSSJack1.png,SSSJack1.png,SSSJack2.png,SSSJack3.png,SSSJack4.png,SSSJack5.png,SSSJack6.png,SSSJack7.png,SSSJack8.png,SSSJack9.png,SSSJack10.png,SSSJack11.png,SSSJack12.png,SSSJack13.png,SSSJack14.png"
  imgList = imgList & ",SJP1.png,SJP2.png,SJP3.png,SJP4.png,SJP5.png,SJP6.png,SJP7.png,SJP8.png,SJP9.png,SJP10.png,SJP11.png,SJP12.png,SJP13.png,SJP14.png,SJP15.png,SJP16.png,SJP17.png,SJP18.png,SJP19.png"
  CheckImageList(ImgList)
  SSSJackback = UltraDMD.CreateAnimationFromImages(8, False, imgList)

  imglist = "SSUJack1.png,SSUJack1.png,SSUJack2.png,SSUJack3.png,SSUJack4.png,SSUJack5.png,SSUJack6.png,SSUJack7.png,SSUJack8.png,SSUJack9.png,SSUJack10.png,SSUJack11.png,SSUJack12.png,SSUJack13.png,SSUJack14.png,SSUJack15.png,SSUJack16.png,SSUJack17.png,SSUJack18.png,SSUJack19.png,SSUJack20.png"
  imgList = imgList & ",UJP1.png,UJP2.png,UJP3.png,UJP4.png,UJP5.png,UJP6.png,UJP7.png,UJP8.png,UJP9.png,UJP10.png,UJP11.png,UJP12.png,UJP13.png,UJP14.png,UJP15.png,UJP16.png,UJP17.png,UJP18.png,UJP19.png,UJP20.png,UJP21.png"
  CheckImageList(ImgList)
  SSUJackback = UltraDMD.CreateAnimationFromImages(8, False, imgList)

  imglist = "Cow0.png,Cow0.png,Cow1.png,Cow2.png,Cow3.png,Cow4.png,Cow5.png,Cow6.png,Cow7.png,Cow8.png,Cow9.png,Cow10.png,Cow11.png,Cow12.png,Cow13.png,Cow14.png,Cow15.png,Cow16.png,Cow17.png,Cow18.png"
  CheckImageList(ImgList)
  Cowback = UltraDMD.CreateAnimationFromImages(8, False, imgList)

  imglist = "CMan1.png,CMan1.png,CMan2.png,CMan3.png,CMan4.png,CMan5.png,CMan6.png,CMan7.png,CMan8.png,CMan9.png,CMan10.png,CMan11.png,CMan12.png,CMan13.png,CMan14.png,CMan15.png,CMan16.png,CMan17.png,CMan18.png"
  CheckImageList(ImgList)
  ComboManiaback = UltraDMD.CreateAnimationFromImages(8, False, imgList)

  imglist = "SMan1.png,SMan1.png,SMan2.png,SMan3.png,SMan4.png,SMan5.png,SMan6.png,SMan7.png,SMan8.png,SMan9.png,SMan10.png,SMan11.png"
  CheckImageList(ImgList)
  SecretManiaback = UltraDMD.CreateAnimationFromImages(9, False, imgList)

  imglist = "EB1.png,EB1.png,EB2.png,EB3.png,EB4.png,EB5.png,EB6.png,EB7.png,EB8.png,EB9.png,EB10.png,EB11.png,EB12.png,EB13.png,EB14.png,EB15.png,EB16.png,EB17.png,EB18.png,EB19.png,EB20.png,EB21.png,EB22.png,EB23.png,EB24.png,EB25.png,EB26.png,EB27.png,EB28.png,EB29.png,EB30.png"
  CheckImageList(ImgList)
  ExtraBallback = UltraDMD.CreateAnimationFromImages(10, False, imgList)

  imglist = "EBLit1.png,EBLit1.png,EBLit2.png,EBLit3.png,EBLit4.png,EBLit5.png,EBLit6.png,EBLit7.png,EBLit8.png,EBLit9.png,EBLit10.png,EBLit11.png,EBLit12.png,EBLit13.png,EBLit14.png"
  CheckImageList(ImgList)
  EBLitback = UltraDMD.CreateAnimationFromImages(8, False, imgList)

  imglist = "Rescue1.png,Rescue1.png,Rescue2.png,Rescue3.png,Rescue4.png,Rescue5.png,Rescue6.png,Rescue7.png"
  CheckImageList(ImgList)
  RescueLitback = UltraDMD.CreateAnimationFromImages(8, False, imgList)

  imglist = "MagReady1.png,MagReady1.png,MagReady1.png,MagReady1.png,MagReady1.png,MagReady1.png,MagReady2.png,MagReady3.png,MagReady4.png,MagReady5.png,MagReady6.png,MagReady7.png"
  CheckImageList(ImgList)
  MagnetLitback = UltraDMD.CreateAnimationFromImages(8, False, imgList)

  imglist = "LockLit1.png,LockLit1.png,LockLit1.png,LockLit2.png,LockLit2.png,LockLit3.png,LockLit4.png,LockLit5.png,LockLit6.png,LockLit7.png,LockLit8.png,LockLit9.png,LockLit10.png,LockLit11.png,LockLit12.png,LockLit13.png,LockLit14.png,LockLit15.png"
  CheckImageList(ImgList)
  LockLitback = UltraDMD.CreateAnimationFromImages(8, False, imgList)

  imglist = "LoopA1.png,LoopA1.png,LoopA2.png,LoopA3.png,LoopA4.png,LoopA5.png,LoopA6.png,LoopA7.png,LoopA8.png,LoopA9.png"
  CheckImageList(ImgList)
  LoopAback = UltraDMD.CreateAnimationFromImages(8, False, imgList)

  imglist = "LoopB1.png,LoopB1.png,LoopB2.png,LoopB3.png,LoopB4.png,LoopB5.png,LoopB6.png,LoopB7.png,LoopB8.png"
  CheckImageList(ImgList)
  LoopBback = UltraDMD.CreateAnimationFromImages(8, False, imgList)

  imglist = "Mag1.png,Mag1.png,Mag2.png,Mag3.png,Mag4.png,Mag5.png,Mag6.png,Mag7.png,Mag8.png,Mag9.png,Mag10.png,Mag11.png,Mag12.png"
  CheckImageList(ImgList)
  Magnetback = UltraDMD.CreateAnimationFromImages(10, False, imgList)

  imglist = "BX2_1.png,BX2_1.png,BX2_1.png,BX2_1.png,BX2_2.png,BX2_3.png,BX2_4.png,BX2_5.png,BX2_6.png,BX2_7.png"
  CheckImageList(ImgList)
  B2back = UltraDMD.CreateAnimationFromImages(8, False, imgList)

  imglist = "BX4_1.png,BX4_1.png,BX4_1.png,BX4_1.png,BX4_2.png,BX4_3.png,BX4_4.png,BX4_5.png,BX4_6.png,BX4_7.png"
  CheckImageList(ImgList)
  B4back = UltraDMD.CreateAnimationFromImages(8, False, imgList)

  imglist = "BX6_1.png,BX6_1.png,BX6_1.png,BX6_1.png,BX6_2.png,BX6_3.png,BX6_4.png,BX6_5.png,BX6_6.png,BX6_7.png"
  CheckImageList(ImgList)
  B6back = UltraDMD.CreateAnimationFromImages(8, False, imgList)

  imglist = "BX8_1.png,BX8_1.png,BX8_1.png,BX8_1.png,BX8_2.png,BX8_3.png,BX8_4.png,BX8_5.png,BX8_6.png,BX8_7.png"
  CheckImageList(ImgList)
  B8back = UltraDMD.CreateAnimationFromImages(8, False, imgList)

  imglist = "BX10_1.png,BX10_1.png,BX10_1.png,BX10_1.png,BX10_2.png,BX10_3.png,BX10_4.png,BX10_5.png,BX10_6.png,BX10_7.png"
  CheckImageList(ImgList)
  B10back = UltraDMD.CreateAnimationFromImages(8, False, imgList)

  imglist = "BMax_1.png,BMax_1.png,BMax_1.png,BMax_1.png,BMax_2.png,BMax_3.png,BMax_4.png,BMax_5.png,BMax_6.png,BMax_7.png"
  CheckImageList(ImgList)
  BMaxback = UltraDMD.CreateAnimationFromImages(8, False, imgList)

  imglist = "FF1.png,FF1.png,FF2.png,FF3.png,FF4.png,FF5.png,FF6.png,FF7.png,FF8.png,FF9.png,FF10.png,FF11.png,FF12.png"
  CheckImageList(ImgList)
  FFback = UltraDMD.CreateAnimationFromImages(8, False, imgList)

  imgList = "JPa1.png,JPa1.png,JPa2.png,JPa3.png,JPa4.png,JPa5.png,JPa6.png,JPa7.png,JPa8.png,JPa9.png,JPa10.png,JPa11.png,JPa12.png"
  CheckImageList(ImgList)
  JPaback = UltraDMD.CreateAnimationFromImages(8, False, imgList)

  imgList = "JPb1.png,JPb1.png,JPb2.png,JPb3.png,JPb4.png,JPb5.png,JPb6.png,JPb7.png,JPb8.png,JPb9.png,JPb10.png,JPb11.png,JPb12.png,JPb13.png,JPb14.png,JPb15.png,JPb16.png"
  CheckImageList(ImgList)
  JPbback = UltraDMD.CreateAnimationFromImages(8, False, imgList)

  imgList = "JP1.png,JP1.png,JP2.png,JP3.png,JP4.png,JP5.png,JP6.png,JP7.png,JP8.png,JP9.png,JP10.png,JP11.png,JP12.png"
  CheckImageList(ImgList)
  JPcback = UltraDMD.CreateAnimationFromImages(8, False, imgList)

  imgList = "Ca1.png,Ca1.png,Ca2.png,Ca3.png,Ca4.png,Ca5.png,Ca6.png,Ca7.png,Ca8.png,Ca9.png,Ca10.png,Ca11.png"
  CheckImageList(ImgList)
  Combo1back = UltraDMD.CreateAnimationFromImages(10, False, imgList)

  imgList = "Cb1.png,Cb1.png,Cb2.png,Cb3.png,Cb4.png,Cb5.png,Cb6.png,Cb7.png,Cb8.png,Cb9.png,Cb10.png,Cb11.png"
  CheckImageList(ImgList)
  Combo2back = UltraDMD.CreateAnimationFromImages(10, False, imgList)

  imgList = "Cc1.png,Cc1.png,Cc2.png,Cc3.png,Cc4.png,Cc5.png,Cc6.png,Cc7.png,Cc8.png,Cc9.png,Cc10.png,Cc11.png"
  CheckImageList(ImgList)
  Combo3back = UltraDMD.CreateAnimationFromImages(10, False, imgList)

  imgList = "DC1.png,DC1.png,DC2.png,DC3.png,DC4.png,DC5.png,DC6.png,DC7.png,DC8.png,DC9.png,DC10.png,DC11.png"
  CheckImageList(ImgList)
  DComboback = UltraDMD.CreateAnimationFromImages(10, False, imgList)

  imgList = "TC1.png,TC1.png,TC2.png,TC3.png,TC4.png,TC5.png,TC6.png,TC7.png,TC8.png,TC9.png,TC10.png,TC11.png"
  CheckImageList(ImgList)
  TComboback = UltraDMD.CreateAnimationFromImages(10, False, imgList)

  imgList = "SC1.png,SC1.png,SC2.png,SC3.png,SC4.png,SC5.png,SC6.png,SC7.png,SC8.png,SC9.png,SC10.png,SC11.png,SC12.png,SC13.png,SC14.png,SC15.png,SC16.png,SC17.png,SC18.png,SC19.png,SC20.png,SC21.png,SC22.png,SC23.png,SC24.png,SC25.png"
  CheckImageList(ImgList)
  SComboback = UltraDMD.CreateAnimationFromImages(10, False, imgList)

  imgList = "Mis_R_1.png,Mis_R_2.png,Mis_R_3.png,Mis_R_4.png,Mis_R_5.png,Mis_R_6.png,Mis_R_7.png,Mis_R_8.png"
  imgList = imgList & ",M_H_1.png,M_H_2.png,M_H_3.png,M_H_4.png,M_H_5.png,M_H_6.png,M_H_7.png,M_H_8.png"
  imgList = imgList & ",M_B_1.png,M_B_2.png,M_B_3.png,M_B_4.png,M_B_5.png,M_B_6.png,M_B_7.png,M_B_8.png"
  imgList = imgList & ",Sky1.png,Sky2.png,Sky3.png,Sky4.png,Sky5.png,Sky10.png,Sky11.png,Sky12.png,Sky13.png"
  imgList = imgList & ",Sh1.png,Sh2.png,Sh3.png,Sh4.png,Sh5.png,Sh6.png"
  imgList = imgList & ",Show1.png,Show2.png,Show3.png,Show4.png,Show5.png,Show6.png,Show7.png"
  imgList = imgList & ",SSJack1.png,SSJack2.png,SSJack3.png,SSJack4.png,SSJack5.png,SSJack6.png,SSJack7.png,SSJack8.png,SSJack9.png"
  imgList = imgList & ",PL1.png,PL2.png,PL3.png,PL4.png,PL5.png,PL6.png"
  imgList = imgList & ",Ball1.png,Ball2.png,Ball3.png,Ball4.png,Ball5.png,Ball6.png"
  imgList = imgList & ",BR1.png,BR2.png,BR3.png,BR4.png,BR5.png,BR6.png,BR7.png,BR8.png,BR9.png"
  imgList = imgList & ",Shoot1.png,Shoot2.png,Shoot3.png,Shoot4.png,Shoot5.png,Shoot6.png,Shoot7.png"
  imgList = imgList & ",Dock1.png,Dock2.png,Dock3.png,Dock4.png,Dock5.png,Dock6.png"
  imgList = imgList & ",Dock10.png,Dock11.png,Dock12.png,Dock13.png,Dock14.png,Dock15.png,Dock16.png"
  imgList = imgList & ",SLb1.png,SLb2.png,SLb3.png,SLb4.png,SLb5.png,SLb6.png,SLb7.png"
  imgList = imgList & ",SLd1.png,SLd2.png,SLd3.png,SLd4.png,SLd5.png,SLd6.png,SLd7.png,SLd8.png"
  imgList = imgList & ",Ex5.png,Ex6.png,Ex7.png,Ex8.png,Ex9.png,Ex10.png,Ex11.png,Ex12.png"
  CheckImageList(ImgList)
  UltimateShowdownBack = UltraDMD.CreateAnimationFromImages(15, True, imgList)

  Text1 = FormatDMDTopText("","","")
  Text2 = FormatDMDBottomText("")
End Sub

Dim Text1,Text2,backgrnd

Dim ActiveBackground,MissionText
Sub DMDStartMission(MissionID,Duration)
  If UltraDMD.IsRendering Then
    UltraDMD.CancelRendering
  end if
  Text2 = FormatDMDBottom2Text(cstr(Duration),"00 ")
  select case MissionID
    case "Reactor":   MissionText = FormatDMDTopText("","REACTOR CRITICAL","")
              Text1 = MissionText
              ActiveBackground = reactorback
              DMDDisplayScene "Mission",True,ActiveBackground, Text1, 8,14,Text2, 5,15, 14, Duration*1000, 14
    case "HuntDown":  MissionText = FormatDMDTopText("","HUNT  DOWN","")
              Text1 = MissionText
              ActiveBackground = HuntDownback
              DMDDisplayScene "Mission",True,ActiveBackground, Text1, 8,14,Text2, 5,15, 14, Duration*1000, 14
    case "BikeRace":  MissionText = FormatDMDTopText("","BIKE  RACE","")
              Text1 = MissionText
              ActiveBackground = BikeRaceback
              DMDDisplayScene "Mission",True,ActiveBackground, Text1, 8,14,Text2, 5,15, 14, Duration*1000, 14
    case "AmmoDump":  MissionText = FormatDMDTopText("","AMMO  DUMP","")
              Text1 = MissionText
              ActiveBackground = Ammodumpback
              DMDDisplayScene "Mission",True,ActiveBackground, Text1, 8,14,Text2, 5,15, 14, Duration*1000, 14
    case "Skyscraper":  MissionText = FormatDMDTopText("","SKYSCRAPER","")
              Text1 = MissionText
              DMDDisplayScene "Mission",True,Skyback1, Text1, 8,14,Text2, 5,15, 14, 3000, 14
              ActiveBackground = Skyback2
    case "Shuttle":   MissionText = FormatDMDTopText("","STOP THE SHUTTLE","")
              Text1 = MissionText
              ActiveBackground = Shuttleback
              DMDDisplayScene "Mission",True,ActiveBackground, Text1, 8,14,Text2, 5,15, 14, Duration*1000, 14
    case "SuperLauncher1":  MissionText = FormatDMDTopText("","","")
              Text1 = MissionText
              Text2 = FormatDMDBottom2Text("","")
              DMDDisplayScene "Mission",True,SLAback, Text1, 8,14,Text2, 5,15, 14, 3500, 14
              ActiveBackground = SLBback
    case "SuperLauncher2":  MissionText = FormatDMDTopText("","","")
              Text1 = MissionText
              Text2 = FormatDMDBottom2Text("","")
              DMDDisplayScene "Mission",True,SLBback, Text1, 8,14,Text2, 5,15, 14, 1000, 14
              ActiveBackground = SLCback
    case "SpaceStation":  MissionText = FormatDMDTopText("","","")
              Text1 = MissionText
              Text2 = FormatDMDBottom2Text("","")
              DMDDisplayScene "Mission",True,Spaceback1, Text1, 8,14,Text2, 5,15, 14, 3000, 14
              ActiveBackground = Spaceback2
  end select
End Sub

Sub DMDEndMission(ScorePar)
  DMDDisplayScene "Main",True,backgrnd, FormatDMDTopText("","MISSION  TOTAL",""), 8,14,FormatDMDBottomText(FormatPointText(cstr(scorepar))), 5,15, 14, 3800, 14
End Sub

Sub DMDEndVideoMode(ScorePar)
  DMDDisplayScene "Main",True,backgrnd, FormatDMDTopText("","VIDEO MODE TOTAL",""), 8,14,FormatDMDBottomText(FormatPointText(cstr(scorepar))), 5,15, 14, 3800, 14
End Sub

Sub DMDEndQuickshot(ScorePar)
  DMDDisplayScene "Main",True,backgrnd, FormatDMDTopText("","QUICK SHOT TOTAL",""), 8,14,FormatDMDBottomText(FormatPointText(cstr(scorepar))), 5,15, 14, 4800, 14
End Sub

Sub DMDEndMode
  If UltraDMD.IsRendering Then
    UltraDMD.CancelRendering
  end if
  UltraDMDTimer.Enabled = True
End Sub

Sub DMDPowerLevel(PLPar,TextPar)
  DMDDisplayScene "PL",True,powerlevelback, FormatDMDTopText("",TextPar,""), 8,14,FormatDMDBottomText("POWERLEVEL "&cstr(PLPar)), 5,15, 14, 2000, 14
End Sub

Sub DMDDisplayScene(id,stopprev,bkgnd,toptext,topBrightness,topOutlineBrightness,bottomtext,bottomBrightness,bottomOutlineBrightness,animateIn,pauseTime,animateOut)
    If Not UltraDMD is Nothing Then
    if stopprev then
      If UltraDMD.IsRendering Then
        UltraDMD.CancelRendering
      end if
    end if
        UltraDMD.DisplayScene00EXWithID id,stopprev,bkgnd,TopText,topBrightness,topOutlineBrightness,bottomtext,bottomBrightness,bottomOutlineBrightness,animateIn,pauseTime,animateOut
        If pauseTime > 0 OR animateIn < 14 OR animateOut < 14 Then
            UltraDMDTimer.Enabled = True
        End If
    End If
End Sub

Sub UltraDMDTimer_Timer
    If Not UltraDMD.IsRendering Then
    if not EnterHS and not VideoModeActive then
      'When the scene finishes rendering, then immediately display the scoreboard
      UltraDMDTimer.Enabled = False
      if MissionActive or SpaceStationFrenzyActive or SuperLauncherActive then
        if SuperLauncherActive and (ActiveBackground = SLBback) then
          DMDDisplayScene "Mission",False,ActiveBackground, MissionText, 8,14,Text2, 5,15, 14, 1000, 1
          ActiveBackground = SLCback
          UltraDMDTimer.Enabled = True
        else
          if SpaceStationFrenzyActive then
            Text2 = FormatDMDBottomText("")
          end if
          DMDDisplayScene "Mission",False,ActiveBackground, MissionText, 8,14,Text2, 5,15, 14, 100000, 14
        end if
      else
        UltraDMD.DisplayScene00EXWithID "Main",False,backgrnd, Text1, 8,14,Text2, 5,15, 14, 10000000, 14
        UpdateDMDScore
      end if
    end if
    End If
End Sub

Sub UpdateDMDScore
  if not UltraDMDTextTimer.enabled and not DisplayNextBonus.enabled and not AddBonusTimer.enabled and not SpaceStationFrenzyActive and not SuperLauncherActive and not Missionactive and not ShowDMDTimer.enabled and not UltimateShowdownActive then
    if points < 100000000000 then
      if (BallInPlay > BallsperGame) and not EnterHS then
        Text1 = FormatDMDTopText("","Game  Over","")
      else
        Text1 = FormatDMDTopText("Ball "&Ballinplay,"","PLAYER 1")
      end if
      Text2 = FormatDMDBottomText(FormatPointText(cstr(points)))
    else
      Text1 = FormatDMDTopText("",FormatPointText(cstr(points)),"")
      if (BallInPlay > BallsperGame) and not EnterHS then
        Text2 = FormatDMDBottomText("Game  Over")
      else
        Text2 = FormatDMDBottomText("Ball "&Ballinplay)
      end if
    end if
    UltraDMD.ModifyScene00 "Main", Text1, Text2
  end if
End Sub

Sub UpdateDMDText(SceneID,TextALeft,TextAMid,TextARight,TextB)
  Text1 = FormatDMDTopText(TextALeft,TextAMid,TextARight)
  Text2 = FormatDMDBottomText(TextB)
  UltraDMDTextTimer.enabled = False
  UltraDMDTextTimer.enabled = True
  UltraDMD.ModifyScene00 SceneID, Text1, Text2
End Sub

Sub UltraDMDTextTimer_Timer
  UltraDMDTextTimer.enabled = False
  if not EnterHS and not IdleTimer.enabled then
    UpdateDMDScore
  end if
  if SuperLauncherActive then
    UpdateDMDText "Mission","","","",""
  end if
End Sub

Sub UpdateDMDMission(TextBLeft,TextBRight)
  if len(TextBRight) < 2 then
    TextBRight = "0" & TextBRight
  end if
  TextBRight = TextBRight & " "
  Text2 = FormatDMDBottom2Text(TextBLeft,cstr(TextBRight))
  UltraDMDTextTimer.enabled = False
  UltraDMDTextTimer.enabled = True
  UltraDMD.ModifyScene00 "Mission", Text1, Text2
End Sub

Sub UpdateDMDExplosion(TextB)
  Text1 = FormatDMDTopText("","EXPLOSION  AWARD","")
  Text2 = FormatDMDBottomText(TextB)
  UltraDMDTextTimer.enabled = False
  UltraDMDTextTimer.enabled = True
  UltraDMD.ModifyScene00 "Ex", Text1, Text2
End Sub

Const TopTextLen = 16
Const EmptyText = "                    "
Dim Space1,Space2
Function FormatDMDTopText(LeftTextPar,MidTextPar,RightTextPar)
  if len(LeftTextPar + MidTextPar + RightTextPar) >= TopTextLen then
    FormatDMDTopText = LeftTextPar & MidTextPar & RightTextPar
  else
    space1 = int((TopTextLen - len(LeftTextPar) - len(MidTextPar) - len(RightTextPar))/2)
    if (2 * Space1) < (TopTextLen - len(LeftTextPar) - len(MidTextPar) - len(RightTextPar)) then
      space2 = space1 + 1
    else
      space2 = space1
    end if
    FormatDMDTopText = LeftTextPar & Left(EmptyText,space2) & MidTextPar & Left(EmptyText,space1) & RightTextPar
  end if
End Function

Const BottomTextLen = 14
Function FormatDMDBottomText(TextPar)
  if len(TextPar) >= BottomTextLen then
    FormatDMDBottomText = TextPar
  else
    space1 = int((BottomTextLen - len(TextPar))/2)
    FormatDMDBottomText = Left(EmptyText,BottomTextLen - space1 - Len(TextPar)) & TextPar & Left(EmptyText,space1)
  end if
End Function

Dim space
Const TextLen = 14
Function FormatDMDBottom2Text(TextPar1,TextPar2)
  if len(TextPar1) + len(TextPar2) >= TextLen then
    FormatDMDBottom2Text = TextPar1 + TextPar2
  else
    space = TextLen - len(TextPar1) - len(TextPar2)
    FormatDMDBottom2Text = TextPar1 & Left(EmptyText,space) & TextPar2
  end if
End Function

Sub ShowDMDText(TextPar)
  if not HideAnimation then
    DMDDisplayScene "Main",True,"transparent.png","",8,14,FormatDMDBottomText(TextPar),5,15,UltraDMD_Animation_ScrollOnLeft,750,UltraDMD_Animation_ScrollOffLeft
    ShowDMDTimer.enabled = False
    ShowDMDTimer.enabled = True
  end if
End Sub

Sub ShowDMDTimer_Timer
  ShowDMDTimer.enabled = False
End Sub

'used to prevent early game start
Sub UltraDMDInitTimer_Timer
  UltraDMDInitTimer.enabled = False
End Sub

'------------------------------------------
'-----  Light Sequencer Attract Mode  -----
'------------------------------------------
Sub StartAttractMode()
  for each obj in AllPlayfieldlights
    obj.state = Lightstateoff
  next
  LightSeqAttract.UpdateInterval = 16
    LightSeqAttract.Play SeqDownOn,0,1,200
    LightSeqAttract.Play SeqLeftOff,0,1,200
    LightSeqAttract.Play SeqUpOn,0,1,200
  LightSeqAttract.Play SeqRightOff,0,1,200
End Sub

Sub LightSeqAttract_PlayDone()
  if IdleTimer.enabled then
    StartAttractMode
  end if
End Sub

Sub StopAttractMode
  LightSeqAttract.StopPlay
End Sub

Sub PlayDrainAnimation
  LightSeqDrain.UpdateInterval = 2
    LightSeqDrain.Play SeqDownOn,50,1
End Sub

Sub LightSeqDrain_PlayDone()
End Sub

'------------------------------
'-----  B2S Lamp control  -----
'------------------------------
Sub SetB2SLamp(Lamp,LampState)
  if B2SOn then
    Controller.B2SSetData Lamp,LampState
  end if
end sub

'B2S Flasher sequence - it is recommended to define the flashers separate from the normal lamps, to avoid overruling by the main timer
Dim Count48,Count49,Count50,Count51,cc48,cc49,cc50,cc51

Sub FlashB2SLamp(Lamp,Count)          'Lamp = Flasher No., Count = number of flashes
  if B2SOn then
    for i = 1 to Count
      select case Lamp
        case 48: Flash48.enabled = 1:Count48 = Count:cc48 = 0
        case 49: Flash49.enabled = 1:Count49 = Count:cc49 = 0
        case 50: Flash50.enabled = 1:Count50 = Count:cc50 = 0
        case 51: Flash51.enabled = 1:Count51 = Count:cc51 = 0
      end select
    next
  end if
end sub

Sub Flash48_Timer
  cc48 = cc48 + 1
  if cc48 >= count48 then
    Flash48.enabled = 0
  end if
  Controller.B2SSetData 48,1
  UnFlash48.enabled = 1
End Sub
Sub UnFlash48_Timer
  UnFlash48.enabled = 0
  Controller.B2SSetData 48,0
End Sub

Sub Flash49_Timer
  cc49 = cc49 + 1
  if cc49 >= count49 then
    Flash49.enabled = 0
  end if
  Controller.B2SSetData 49,1
  UnFlash49.enabled = 1
End Sub
Sub UnFlash49_Timer
  UnFlash49.enabled = 0
  Controller.B2SSetData 49,0
End Sub

Sub Flash50_Timer
  cc50 = cc50 + 1
  if cc50 >= count50 then
    Flash50.enabled = 0
  end if
  Controller.B2SSetData 50,1
  UnFlash50.enabled = 1
End Sub
Sub UnFlash50_Timer
  UnFlash50.enabled = 0
  Controller.B2SSetData 50,0
End Sub

Sub Flash51_Timer
  cc51 = cc51 + 1
  if cc51 >= count51 then
    Flash51.enabled = 0
  end if
  Controller.B2SSetData 51,1
  UnFlash51.enabled = 1
End Sub
Sub UnFlash51_Timer
  UnFlash51.enabled = 0
  Controller.B2SSetData 51,0
End Sub


' *********************************************************************
'                      Supporting Ball & Sound Functions
' *********************************************************************

Function Vol(ball) ' Calculates the Volume of the sound based on the ball speed
    Vol = Csng(BallVel(ball) ^2 / 500)    '2000
End Function

Function Pan(ball) ' Calculates the pan for a ball based on the X position on the table. "TheWeb" is the name of the table
    Dim tmp
    tmp = ball.x * 2 / TheWeb.width-1
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

    'Speed Control
    if BallVel(BOT(b)) > DampeningSpeed then
      BOT(b).velx = 0.95 * BOT(b).velx
      BOT(b).vely = 0.95 * BOT(b).vely
    end if

        If BallVel(BOT(b) ) > 1 AND BOT(b).z < 30 Then
            rolling(b) = True
            PlaySound("fx_ballrolling" & b), -1, Vol(BOT(b) ), Pan(BOT(b) ), 0, Pitch(BOT(b) ), 1, 0
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
  PlaySound("fx_collide"), 0, Csng(velocity) ^2 / 2000, Pan(ball1), 0, Pitch(ball1), 0, 0
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
  PlaySound "pinhit_low", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0
End Sub

Sub Targets_Hit (idx)
  PlaySound "target", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 0, 0
End Sub

Sub Metals_Thin_Hit (idx)
  PlaySound "metalhit_thin", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0
End Sub

Sub Metals_Medium_Hit (idx)
  PlaySound "metalhit_medium", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0
End Sub

Sub Metals2_Hit (idx)
  PlaySound "metalhit2", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0
End Sub

Sub Gates_Hit (idx)
  PlaySound "gate4", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0
End Sub

'Sub Spinner_Spin
' PlaySound "fx_spinner",0,.25,0,0.25
'End Sub

Sub Rubbers_Hit(idx)
  dim finalspeed
    finalspeed=SQR(activeball.velx * activeball.velx + activeball.vely * activeball.vely)
  If finalspeed > 20 then
    PlaySound "fx_rubber2", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0
  End if
  If finalspeed >= 6 AND finalspeed <= 20 then
    RandomSoundRubber()
  End If
End Sub

Sub Posts_Hit(idx)
  dim finalspeed
    finalspeed=SQR(activeball.velx * activeball.velx + activeball.vely * activeball.vely)
  If finalspeed > 16 then
    PlaySound "fx_rubber2", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0
  End if
  If finalspeed >= 6 AND finalspeed <= 16 then
    RandomSoundRubber()
  End If
End Sub

Sub RandomSoundRubber()
  Select Case Int(Rnd*3)+1
    Case 1 : PlaySound "rubber_hit_1", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0
    Case 2 : PlaySound "rubber_hit_2", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0
    Case 3 : PlaySound "rubber_hit_3", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0
  End Select
End Sub

Sub LeftFlipper_Collide(parm)
  RandomSoundFlipper()
End Sub

Sub ULeftFlipper_Collide(parm)
  RandomSoundFlipper()
End Sub

Sub RightFlipper_Collide(parm)
  RandomSoundFlipper()
End Sub

Sub RandomSoundFlipper()
  Select Case Int(Rnd*3)+1
    Case 1 : PlaySound "flip_hit_1", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0
    Case 2 : PlaySound "flip_hit_2", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0
    Case 3 : PlaySound "flip_hit_3", 0, Vol(ActiveBall), Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0
  End Select
End Sub

Sub RWireRampStart1_hit
  PlaySound "WireRamp1", 0, 1, Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0       'Vol(ActiveBall)
End Sub

Sub RWireRampStart2_hit
  PlaySound "WireRamp1", 0, 1, Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0
End Sub

Sub RWireRampStart3_hit
  PlaySound "WireRamp1", 0, 1, Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0
End Sub

Sub VUKSoundTrigger_hit
  PlaySound "WireRamp1", 0, 1, Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0
End Sub

Sub LVUKSoundTrigger_hit
  PlaySound "WireRamp1", 0, 1, Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0
End Sub

Sub RVUKSoundTrigger_hit
  PlaySound "WireRamp1", 0, 1, Pan(ActiveBall), 0, Pitch(ActiveBall), 1, 0
End Sub

'Launch Button DOF
Sub AutoPlungerTrigger_Hit
  if AutoPlungerActive = False and AutoPlungerTimer.enabled = False Then
  DOF 170,1
  end if
End Sub

Sub AutoPlungerTrigger_UnHit
  if AutoPlungerActive = False and AutoPlungerTimer.enabled = False  Then
  DOF 170,0
  end if
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

Function AudioFade(tableobj) ' Fades between front and back of the table (for surround systems or 2x2 speakers, etc), depending on the Y position on the table. "TheWeb" is the name of the table
  Dim tmp
    tmp = tableobj.y * 2 / TheWeb.height-1
    If tmp > 0 Then
    AudioFade = Csng(tmp ^10)
    Else
        AudioFade = Csng(-((- tmp) ^10) )
    End If
End Function

Function AudioPan(tableobj) ' Calculates the pan for a tableobj based on the X position on the table. "TheWeb" is the name of the table
    Dim tmp
    tmp = tableobj.x * 2 / TheWeb.width-1
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

Function Pan(ball) ' Calculates the pan for a ball based on the X position on the table. "TheWeb" is the name of the table
    Dim tmp
    tmp = ball.x * 2 / TheWeb.width-1
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
' PlaySound("fx_collide"), 0, Csng(velocity) ^2 / 2000, Pan(ball1), 0, Pitch(ball1), 0, 0
  PlaySound("fx_collide"), 0, Csng(velocity) ^2 / 2000, AudioPan(ball1), 0, Pitch(ball1), 0, 0, AudioFade(ball1)
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


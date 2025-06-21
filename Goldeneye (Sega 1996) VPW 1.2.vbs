
'  __   __        __   ___       ___      ___
' / _` /  \ |    |  \ |__  |\ | |__  \ / |__
' \__> \__/ |___ |__/ |___ | \| |___  |  |___
'
'
' Based on the VPX table by JPJ, Team PP and many more (Pingod, UncleReamus, Destruk, The Trout, 32assassin, JPJ, Dark, Neo, Arngrim, Flupper, Rothbauerw, Chucky, Fleep, Knorr, Jpn Osborne, JPSalas, Neofr45, Chucky87, Rajo Joey, Thalamus, Davidcd13, Toxie)
'
' VPW GoldenEye Team:
' -------------------
' Project lead: Astronasty
' Playfield cleanup: Astronasty
' Physics: Fluffhead35, Apophis
' Sound effects: Fluffhead35
' 3D Inserts: EBisLit, Oqqsan
' Dynamic Shadows: Wylte
' VR Room and Backglass: Leojreimroc, Sixtoe
' Lighting: Hauntfreaks, Sixtoe
' GI control update: Iaakki
' Radar dish updates: Apophis, Oqqsan
' Flasher dome updates: Apophis, Leojreimroc
' Script fixes: Oqqsan,Apophis
' Desktop background: Oqqsan
' Ramp updates: Tomate
' Testing: VPW team, 1 day by Rik
'
'
' CHANGE LOG
' 12 - fluffhead35  - Addeed nfozzy physics, fleep sound, rubberizer, target bouncer, coil rampup
' 13 - fluffhead35  - Adjusted physics, table slope, ramp material, and adjustments for right ramp
' 14 - fluffehad35  - Adjusted physics on gates
' 15 - Wylte    - Shadowzzzzzz
' 16 - apophis    - Installed Lampz in prep for 3D inserts. Updated PF, plastics, and insert overlay images. Made all rubber objects non-collidable, Added prototype Enhanced Sounds option (no sounds installed yet). Auto-formatted script.
' 17 - EBisLit    - Added 3D inserts
' 18 - EBisLit    - tweaking insert brightness
' 20 - HauntFreaks  - GI lighting and shadows - small tweaks to PF and plastics
' 21 - oqqsan   - RadarA   removed normalmap + changed material to 3dapron, might need some adjustment
' 22c - fluffhead35 - Updated the physics and ballroll sounds.  Realigned rubbers, posts, and pegs.
' 23 - oqqsan     - small fixes and, added extra vpxlight on all inserts
' 24 - oqqsan   - fixed ball in radar !
' 25 - Leojreimroc  - Adjusted tabled flashers.  Added fade timers to all flashers.  Many solenoids were pointing to the wrong flashers, fixed.
' 26 - oqqsan     - fixed on some lights,spotlight flasher, swapped out metal bracket, shrinked flippers to lotr size, added volumedial for some mechsounds
' 27 - Leojreimroc  - VR Backglass added.  VR Room Code, slight flasher adjustments.
' 28 - fluffhead35  - Added Materials to all Collidables.  Added Hit Events to metal collidables.
' 29 - Sixtoe   - Fixed plunger lane, fixed ramps, redid gi a bit, hooked up ramps to gi, hooked up ramp lights, despaired, drank.
' 30 - apophis    - Flupperized the flasher domes. Updated dynamic shadows. Updated relay sound effects and added radar motor sound effect. Some code cleanup.
' 31 - oqqsan     - hooked up the vrbg flashers to flupperdomes. fixed the wall above DT's one more time ( need testing so it dont just drain :P )
' 32 - oqqsan   - made plunger shoot much harder 60? instead of 45 strength. adjusted flippers, from 70 to max 68degrees. metalic on rampmaterials removed
' 33 - oqqsan   - silly mistakes hunting ! :P
' 35 - oqqsan   - fixed silly timers... DOUBLE FPS after :)   added dof calls for flippers
' 36 - oqqsan   - New super Desktop backscreen made just before release time
' 36_gi - iaakki  - GI control rework and ramp difference textures added and tied to fading
' 37 - oqqsan   - Not used
' 38 - apophis    - Rebuilt the radar dish mechanic. Fixed a few plunger lane issues. Made LampTimer fixed 16ms interval. Increased GI fade speed a bit.
' 39 - oqqsan   - Updated ballsave kick angle
' 40b - apophis   - Added slingshot correction code. Fixed satellite lock so that ball can be unlocked by ball-ball collisions. Increased insert DL a little.
' 41 - tomate       - Added new wireRamps and plasticRamps textures (OFF plasticRamps prims still has old texture)
' 42 - apophis    - Tuned wire ramp DL (removed wire ramps from DL collection). Updated ramp off primitives, texture, and material. Added ramp off primitives to OFF_Prims collection. Fixed right flipper trigger. Increased playfield and ramp frictions a little.
' 1.0 - apophis   - Release
' 1.0.1 - Niwak   - Fixed table bug causing stuck flippers (added vpmInit Me)
' 1.0.2 - DaRDog81  - Added VR Mega room.
' 1.0.3 - apophis - Added tweak menu. Updated physics scripts, flipper triggers. Updated Fleep scripts. Updated ball and ramp rolling scripts. Updated ball shadow scripts and using new built in rtx shadows. Script cleanups.
' 1.0.4 - apophis - PWM updates for inserts, GI, flashers, and VR backglass. Hooked up VR rooms to day night setting. More script cleanups.
' 1.0.5 - apophis - Rewrote sat animation routine and adjusted speed. Added motor sfx. Fixed "sneek in" issue. Added visual blocker under PF.
' 1.0.6 - apophis - Adjusted flipper and table physics params (thanks Astro). Fixed satellite home position switch. Fixed see-thru part of sat near magnet. Fixed some vr backglass flasher visibilities. Added rules card, updated table info.
' 1.1 Release
' 1.1.1 - apophis - Fixed desktop lights showing in cab mode. Suppressed keyAddBall in key subs.
' 1.1.2 - apophis - Fixed stuck ball issue on ramp. Lowered plunger strength slightly.
' 1.1.3 - apophis - Made more prims on table respond to Day Night setting. Fixed satellite animation init issue.
' 1.2 Release

' PREVIOUS VERSIONS CREDITS INFO
'First big thank to all of them who worked on this earlier version and who permit me to work on this newer :
'Pingod/UncleReamus/Destruk, Mod by The Trout
'
'Big thanks to Assassin32 who share me his wip (just use for some script rutines, but it permited me to compare and to approach some problems with another way to do) So Big big respect and thanks to you.
'
'V1.4
'Double sounds added to DOF (slings, bumpers, flippers)
'Drain sound added.
'Magnet zone of the radar modified.
'
'V1.3
'Lamp test deleted near flippers...
'
'V1.2
'Flippers active when tilted, corrected (thx to Davidcd13 to signal this bug)
'Optimization, thanks Toxie for your very good advices ;)
'
'V1.1
'Speed up a little bit the ball on playfield.
'Sides, Apron and rails modified. Add the lockbar.
'Change a little bit flippers physic, and reajust the Coil Ramp parameter to 2.5 (thx Rajo Joey) it's more natural now.
'Delete some panoramics parameter for letting ssf script doing its job ;) (Thx Thalamus)
'
'
'V1
'- Option to reduce the Magnet's sound in the script
'- Option for Fantasy Magnet, On or Off
'- LUT choices with left and right MagnaSave buttons. No need to change it each time you launch the table, it's saved on your hard disk, on your User folder.
'
'VPX's version
'Scripted way to save LUT Preferences JPJ
'Practicly all 3d  made by JPJ
'Big Big thanks to DARK for his 3d tank, radar and Copter ;) ;)
'Big thanks to NEO for his playfield, made as always, with a lot of layers to permit me to work easily ! For me it's the Cherry on the cake
'Script made by the first authors, assassin32 and JPJ
'Big thanks to Arngrim for DOF
'Big thanks to Flupper for the domes and the way to light them !!
'Big thanks for Rothbauerw for all his tips and Tricks Very useful !!
'Big thanks to Chucky for his B2s and for founding to me a lot of HD pictures. For founding the complete documentation. And for a lot of other important things !!!!!
'Big thanls tp Fleep who permit me to use some sounds of its famous Theatre of Magic.
'Big Thanks to Knorr for a lot of famous sounds !!!
'Big Big Thanks to Jon Osborne for some famous sounds too and for spatialize all sounds !!!
'Big thanks for testing : Dark, Thalamus, Jens Leiensetter, Jlou Loulou, Franck Hollinger, Chucky. Big Big thanks to them for all those advices and tips !!
'Big Thanks to Alexander Stang, always for good advices too ;) ;) ;)
'Of course Big Thanks for the VP's Dev Team !!
'
'Special Thanks for our master JP Salas, who helped me to find some script's syntaxes !! ;) ;) ;)
'
'This makes this table THE TABLE of those mentioned here ;)
'
'I hope you'll enjoy this table ;)
'
'JPJ



Option Explicit
Randomize
SetLocale 1033


On Error Resume Next
ExecuteGlobal GetTextFile("controller.vbs")
If Err Then MsgBox "You need the controller.vbs in order to run this table, available in the vp10 package"
On Error Goto 0

'*******************************************
'  Constants and Global Variables
'*******************************************

Const TestVRDT = False

Const BallSize = 50
Const BallMass = 1

Const cGameName = "gldneye"
Const UseVPMModSol = 2
'Const UseSolenoids = 2 'defined in SEGA2.vbs
Const UseLamps = 1
Const UseGI = 0
Const UseSync = 0
Const HandleMech = 0
Const SSolenoidOn = ""
Const SSolenoidOff = ""
Const SCoin = ""

Dim UseVPMDMD, DesktopMode
DesktopMode = Goldeneye.ShowDT
If RenderingMode = 2 OR TestVRDT=True Then UseVPMDMD = True Else UseVPMDMD = DesktopMode

LoadVPM "03060000","SEGA2.VBS",3.1

Dim tablewidth: tablewidth = goldeneye.width
Dim tableheight: tableheight = goldeneye.height
Dim gilvl : gilvl = 0

Const tnob = 7 ' total number of balls
Const lob = 0 'locked balls on start; might need some fiddling depending on how your locked balls are done


Dim bsTrough,bsScoop,bsTank,bsPlunger,bsLockOut,mFlipperMagnet,RadarMagnet
Dim FL, FR, GO
FL = 0:FR = 0:GO = 1

NoUpperLeftFlipper
NoUpperRightFlipper


'*******************************************
'  ZOPT: User Options
'*******************************************

Dim ColorLUT : ColorLUT = 0           ' Color desaturation LUTs: 1 to 11, where 1 is normal and 11 is black'n'white
Dim VolumeDial : VolumeDial = 0.8             ' Overall Mechanical sound effect volume. Recommended values should be no greater than 1.
Dim BallRollVolume : BallRollVolume = 0.5     ' Level of ball rolling volume. Value between 0 and 1
Dim RampRollVolume : RampRollVolume = 0.5     ' Level of ramp rolling volume. Value between 0 and 1
Dim FantasyMagnet : FantasyMagnet = 1       '1 For Fantasy colored magnet, 0 for nothing
Dim VRRoomChoice : VRRoomChoice = 2       '1 = Minimal Room, 2 = Mega Room
Dim VRRoom
Dim LightLevel

' Called when options are tweaked by the player.
' - 0: game has started, good time to load options and adjust accordingly
' - 1: an option has changed
' - 2: options have been reseted
' - 3: player closed the tweak UI, good time to update staticly prerendered parts
' Table1.Option arguments are:
' - option name, minimum value, maximum value, step between valid values, default value, unit (0=None, 1=Percent), an optional arry of literal strings
Dim dspTriggered : dspTriggered = False
Sub Goldeneye_OptionEvent(ByVal eventId)
  If eventId = 1 And Not dspTriggered Then dspTriggered = True : DisableStaticPreRendering = True : End If

    ' Fantasy Magnet
    FantasyMagnet = Goldeneye.Option("Fantasy Magnet", 0, 1, 1, 1, 0, Array("Disabled", "Enabled Colored Magnet"))

  ' Color Saturation
    ColorLUT = Goldeneye.Option("Color LUT", 0, 9, 1, 0, 0, Array("Normal", "LUT 1", "LUT 2", "LUT 3", "LUT 4", "LUT 5", "LUT 6", "LUT 7", "LUT 8", "LUT 9") )
  Select Case ColorLUT
    Case 0:Goldeneye.ColorGradeImage = ""
    Case 1:Goldeneye.ColorGradeImage = "ColorGradeLUT256x16_ConSat"
    Case 2:Goldeneye.ColorGradeImage = "LUT1_1_09"
    Case 3:Goldeneye.ColorGradeImage = "LUTbassgeige1"
    Case 4:Goldeneye.ColorGradeImage = "LUTbassgeige2"
    Case 5:Goldeneye.ColorGradeImage = "LUTbassgeigemeddark"
    Case 6:Goldeneye.ColorGradeImage = "LUTfleep"
    Case 7:Goldeneye.ColorGradeImage = "LUTmandolin"
    Case 8:Goldeneye.ColorGradeImage = "LUTVogliadicane70"
    Case 9:Goldeneye.ColorGradeImage = "LUTchucky4"
  end Select

  'VR Room
  VRRoomChoice = Goldeneye.Option("VR Room", 1, 2, 1, 2, 0, Array("Minimal", "MEGA"))
  If RenderingMode = 2 OR TestVRDT=True Then VRRoom = VRRoomChoice Else VRRoom = 0
  SetupRoom

    ' Sound volumes
    VolumeDial = Goldeneye.Option("Mech Volume", 0, 1, 0.01, 0.8, 1)
    BallRollVolume = Goldeneye.Option("Ball Roll Volume", 0, 1, 0.01, 0.5, 1)
  RampRollVolume = Goldeneye.Option("Ramp Roll Volume", 0, 1, 0.01, 0.5, 1)

  ' Room brightness
  LightLevel = NightDay/100
  SetPrimBrightness

  If eventId = 3 And dspTriggered Then dspTriggered = False : DisableStaticPreRendering = False : End If
End Sub




' If you don't have SSF Put 1 to up the magnet Volumes else put 0
'**********************************************************
dim MagnetVolMax, MagnetVol
MagnetVolMax = 0 ' 0 (SSF) or 1 (without SSF)
MagnetVol = 0.5 'level of magnet sounds for non SSF users




'*******************************************
' ZTIM: Timers
'*******************************************


Sub FrameTimer_Timer()
  BSUpdate
  RollingUpdate
  RealTime_Timer
End Sub


'*******************************************
' ZINI: Table Initialization and Exiting
'*******************************************


Sub GoldenEye_Init
  vpmInit Me
  vpmMapLights AllLamps

  SatTop.IsDropped=1
  SatTop1.IsDropped=1
  SatTop2.IsDropped=1
  On Error Resume Next
  With Controller
    .GameName=cGameName
    .Games(cGameName).Settings.Value("rol")=0
    .Games(cGameName).Settings.Value("sound") = 1
    If Err Then MsgBox "Can't start Game" & cGameName & vbNewLine & Err.Description : Exit Sub
    .SplashInfoLine="007 Goldeneye - Sega, 1996"&vbNewLine&"Initial Table Releace by Pingod & Kid Charlemagne"&vbNewLine&"Mod by UncleReamus & Destruk"&vbNewLine&"Further Modding by The Trout"
    .HandleMechanics=0
    .HandleKeyboard=0
    .ShowDMDOnly=1
    .ShowFrame=0
    .ShowTitle=0
    .hidden=DesktopMode
    .Run
  End With

  If Err Then MsgBox Err.Description
  On Error Goto 0

  Gate3.collidable = 0


  '********** Nudge **************
  vpmNudge.TiltSwitch=1
  vpmNudge.Sensitivity=5
  vpmNudge.TiltObj=Array(LeftFlipper,RightFlipper,LeftSlingshot,RightSlingshot,LeftTurboBumper,RightTurboBumper,BottomTurboBumper)
  '*******************************

  Set mFlipperMagnet = New cvpmMagnet
  With mFlipperMagnet
    .InitMagnet FlipperMagnet, 25 '20 '10
    .GrabCenter = True
    .CreateEvents "mFlipperMagnet"
  End With

  Set bsTrough=New cvpmBallStack
  bsTrough.InitSw 0,14,13,12,11,10,0,0
  bsTrough.InitKick BallRelease,40,8
  'bsTrough.InitExitSnd SoundFX("BallRelease1",DOFContactors),SoundFX("Solon",DOFContactors)
  bsTrough.Balls=5

  Set bsScoop=New cvpmBallStack
  bsScoop.InitSw 0,50,0,0,0,0,0,0
  bsScoop.InitKick Scoop,213,10
  bsScoop.InitExitSnd SoundFX("popper",DOFContactors),SoundFX("Solon",DOFContactors)
  bsScoop.KickForceVar=7

  Set bsTank=New cvpmBallStack
  bsTank.InitSw 0,56,0,0,0,0,0,0
  bsTank.InitKick TankKickBig,0,65
  bsTank.InitExitSnd SoundFX("popper",DOFContactors),SoundFX("Solon",DOFContactors)
  bsTank.KickBalls=2
  bsTank.KickForceVar=5

  Set bsPlunger=New cvpmBallStack
  bsPlunger.InitSaucer Plunger,16,0,55
  bsPlunger.InitExitSnd SoundFX("Solon",DOFContactors),SoundFX("SolOn",DOFContactors)
  bsPlunger.KickForceVar=5

  Controller.Switch(63)=0
  Controller.Switch(64)=0



  playfieldOff.opacity = 100
' SidesBack.image = "Backtest"
  Plastics.blenddisablelighting = 0
  BumperA.blenddisableLighting = LightLevel*0.2
  BumperA001.blenddisableLighting = LightLevel*0.2
  BumperA002.blenddisableLighting = LightLevel*0.2
  Rampe3.blenddisableLighting = LightLevel*0.5
  Rampe2.blenddisableLighting = LightLevel*0.3
  Rampe1.blenddisableLighting = LightLevel*0.1
  VisA.blenddisablelighting = 0
  VisB.blenddisablelighting = 0
  MurMetal.blenddisablelighting = 0
  MetalParts.blenddisablelighting = 0
  spotlight.blenddisablelighting = LightLevel*0.5
' spotlightLight.blenddisablelighting = 0 ' 0.5
  PegsBoverSlings.blenddisablelighting = 0
' Rampe3.image = "rampe3GIOFF"
' Rampe2.image = "rampe2GIOFF"
' Rampe1.image = "rampe1GIOFF"
  Apron.blenddisablelighting = 0
  ApronCachePlunger.blenddisablelighting = 0
  MrampB.blenddisablelighting = 0

  Plastics.image = "plasticsOff"
  Ledrouge.Amount = 100:LedRouge.IntensityScale = 0
  LedBleu1.Amount = 100:LedBleu1.IntensityScale = 0
  LedBleu2.Amount = 100:LedBleu2.IntensityScale = 0
  LedFond.Amount = 100:LedFond.IntensityScale = 0

  If RenderingMode=2 OR TestVRDT=True Then SetBackglass

End Sub



Sub GoldenEye_Exit  '  in some tables this needs to be Table1_Exit
  Controller.Games(cGameName).Settings.Value("sound") = 1
  Controller.Stop
End Sub



'*******************************************
' ZKEY:  Key presses
'*******************************************


Sub GoldenEye_KeyDown(ByVal Keycode)
  If KeyCode=StartGameKey Then soundStartButton()
  If KeyCode=KeySlamDoorHit Then Controller.Switch(7)=1
  If keycode = LeftTiltKey Then Nudge 90, 1:SoundNudgeLeft()
  If keycode = RightTiltKey Then Nudge 270, 1:SoundNudgeRight()
  If keycode = CenterTiltKey Then Nudge 0, 1:SoundNudgeCenter()
  If KeyCode=PlungerKey Then Controller.Switch(9)=1 ': SoundPlungerPull()
  If keycode = keyInsertCoin1 or keycode = keyInsertCoin2 or keycode = keyInsertCoin3 or keycode = keyInsertCoin4 Then
    Select Case Int(rnd*3)
      Case 0: PlaySound ("Coin_In_1"), 0, CoinSoundLevel, 0, 0.25
      Case 1: PlaySound ("Coin_In_2"), 0, CoinSoundLevel, 0, 0.25
      Case 2: PlaySound ("Coin_In_3"), 0, CoinSoundLevel, 0, 0.25
    End Select
  End If
  If keycode = LeftFlipperKey Then
    FlipperActivate LeftFlipper, LFPress
    Primary_flipper_button_left.X = 2112 + 8
  End If
  If keycode = RightFlipperKey Then
    FlipperActivate RightFlipper, RFPress
    Primary_flipper_button_right.X = 2090.5 - 8
  End if
  If keycode = keyAddBall Then Exit Sub
  If vpmKeyDown(KeyCode) Then Exit Sub
End Sub


Sub GoldenEye_KeyUp(ByVal KeyCode)
  If KeyCode=KeySlamDoorHit Then Controller.Switch(7)=0
  If KeyCode=PlungerKey Then Controller.Switch(9)=0 ': SoundPlungerReleaseBall()
  If keycode = LeftFlipperKey Then
    FlipperDeActivate LeftFlipper, LFPress
    Primary_flipper_button_left.X = 2112
  End if
  If keycode = RightFlipperKey Then
    FlipperDeActivate RightFlipper, RFPress
    Primary_flipper_button_right.X = 2090.5
  End if
  If keycode = keyAddBall Then Exit Sub
  If vpmKeyUp(KeyCode) Then Exit Sub
End Sub





'*******************************************
' ZSOL:  Solenoids
'*******************************************

SolCallback(1)="bsTrough.SolOut"
SolCallback(2)="bsPlunger.SolOut"
SolCallback(4)="bsScoop.SolOut"
SolCallback(8)="vpmSolSound SoundFX(""knocker_1"",DOFKnocker),"
'SolCallback(9)="vpmSolSound ""fx_bumper1"","
'SolCallback(10)="vpmSolSound ""fx_bumper2"","
'SolCallback(11)="vpmSolSound ""fx_bumper3"","
'SolCallback(12)="vpmSolSound ""sling1"","
'SolCallback(13)="vpmSolSound ""sling2"","
SolCallback(14)="bsTank.SolOut"

SolCallback(17)="SolLockOut"
SolCallback(18)="SolUpDownRamp"
SolCallback(20)="SolSatLaunchRamp"
SolCallback(21)="SolSatMotorRelay"
SolCallback(22)="DropRamp1.Enabled="

SolModCallback(25)="Sol25"   'Flashers
SolModCallback(26)="Sol26"
SolModCallback(27)="Sol27"
SolModCallback(28)="Sol28"
SolModCallback(29)="Sol29"
SolModCallback(30)="sol30"
SolModCallback(31)="Sol31"
SolModCallback(32)="Sol32"

SolCallback(33)="SolRadarMagnet"
SolCallback(34)="SolFlipperMagnet"
SolCallback(45)="TiltMod"

SolCallback(sLRFlipper) = "SolRFlipper"
SolCallback(sLLFlipper) = "SolLFlipper"



Sub SolLockOut(Enabled)
  If Enabled Then vpmTimer.PulseSw 15
End Sub



'****************** Flippers Magnet **********************

Dim Dir, speed, ball, BallsOnMagnet, Mdof
mdof = 0

' Magnet power is pulsed so wait before turning power off
Sub SolFlipperMagnet(enabled)
  HideFlipper.TimerEnabled = Not enabled
  If enabled Then
    FM = 1
    Light001.state = 1
    mFlipperMagnet.MagnetOn = True
    if MagnetVolMax = 1 then
      PlaySound "fx_magnet",-1, MagnetVol
    else
      PlaySound "fx_magnet",-1
    end if
  Else
    FM = 0
    Light001.state = 0
    StopSound "fx_magnet"
  End If
End Sub

' Magnet is turned off Ball is ejected
Sub HideFlipper_Timer
  Dim dir, speed, ball
  For Each ball In mFlipperMagnet.Balls
    With ball
      dir = 180:speed = 45 + Rnd * 2
      .VelX = -6 : .VelY = speed * Cos(dir) '4 * Sin(dir)
      PlaySound "fx_balldrop",1,VolumeDial
    End With
  Next
  vpmTimer.PulseSw 24:
  mFlipperMagnet.MagnetOn = False:  mFlipperMagnet.GrabCenter = True:mdof = 0
  Me.TimerEnabled = False
End Sub
'******************************************************

'******************************************************

Const ReflipAngle = 20

Sub SolLFlipper(Enabled)
  If Enabled Then
    LF.Fire 'LeftFlipper.RotateToEnd
    Controller.Switch(63)=1
    Dof 101,1
    If leftflipper.currentangle < leftflipper.endangle + ReflipAngle Then
      RandomSoundReflipUpLeft LeftFlipper
    Else
      SoundFlipperUpAttackLeft LeftFlipper
      RandomSoundFlipperUpLeft LeftFlipper
    End If
  Else
    Controller.Switch(63)=0
    Dof 101,0
    LeftFlipper.RotateToStart
    If LeftFlipper.currentangle < LeftFlipper.startAngle - 5 Then
      RandomSoundFlipperDownLeft LeftFlipper
    End If
    FlipperLeftHitParm = FlipperUpSoundLevel
  End If
End Sub

Sub SolRFlipper(Enabled)
  If Enabled Then
    RF.Fire 'RightFlipper.RotateToEnd
    Controller.Switch(64)=1
    Dof 102,1
    If rightflipper.currentangle > rightflipper.endangle - ReflipAngle Then
      RandomSoundReflipUpRight RightFlipper
    Else
      SoundFlipperUpAttackRight RightFlipper
      RandomSoundFlipperUpRight RightFlipper
    End If
  Else
    Controller.Switch(64)=0
    Dof 102,0
    RightFlipper.RotateToStart
    If RightFlipper.currentangle > RightFlipper.startAngle + 5 Then
      RandomSoundFlipperDownRight RightFlipper
    End If
    FlipperRightHitParm = FlipperUpSoundLevel
  End If
End Sub


Sub TiltMod(Enabled)
  If true then
    go = 0
  Else
    go = 1
  end if
End Sub




'*******************************************
' ZFLA:  Flashers
'*******************************************


Sub Sol25(level)
  Flasher25.state = level
  Flash001.state = level
  Flash002.state = level
End Sub

Sub Sol26(level)
  Flasher26.state = level
  LL15F.state = level
End Sub

Sub Sol27(level)
  Flasher27.state = level
  Flash003A.state = level
  Flash003B.state = level
End Sub

Sub Sol28(level)
  Flasher28.state = level
  Flash004.state = level
End Sub

Sub Sol29(level)
  Flasher29.state = level
  ModFlashFlasher 1,level
End Sub

Sub Sol30(level)
  Flasher30.state = level
  Flash003.state = level
  Rampflash4.state = level
End Sub

Sub Sol31(level)
  Flasher31.state = level
  ModFlashFlasher 2,level
  ModFlashFlasher 3,level
End Sub

Sub Sol32(level)
  Flasher32.state = level
  ModFlashFlasher 4,level
End Sub




'******************************************************
'*****   FLUPPER DOMES
'******************************************************


Dim TestFlashers, TableRef, FlasherLightIntensity, FlasherFlareIntensity, FlasherBloomIntensity, FlasherOffBrightness

                ' *********************************************************************
TestFlashers = 0        ' *** set this to 1 to check position of flasher object       ***
Set TableRef = goldeneye      ' *** change this, if your table has another name             ***
FlasherLightIntensity = 0.3   ' *** lower this, if the VPX lights are too bright (i.e. 0.1)   ***
FlasherFlareIntensity = 0.1   ' *** lower this, if the flares are too bright (i.e. 0.1)     ***
FlasherBloomIntensity = 0.2   ' *** lower this, if the blooms are too bright (i.e. 0.1)     ***
FlasherOffBrightness = 0.3    ' *** brightness of the flasher dome when switched off (range 0-2)  ***
                ' *********************************************************************

Dim ObjLevel(20), objbase(20), objlit(20), objflasher(20), objbloom(20), objlight(20)
'Dim tablewidth, tableheight : tablewidth = TableRef.width : tableheight = TableRef.height
'initialise the flasher color, you can only choose from "green", "red", "purple", "blue", "white" and "yellow"

InitFlasher 1, "yellow"
InitFlasher 2, "red"
InitFlasher 3, "red"
InitFlasher 4, "yellow"

' rotate the flasher with the command below (first argument = flasher nr, second argument = angle in degrees)
'RotateFlasher 1,17 : RotateFlasher 2,0 : RotateFlasher 3,90 : RotateFlasher 4,90


Sub InitFlasher(nr, col)
  ' store all objects in an array for use in FlashFlasher subroutine
  Set objbase(nr) = Eval("Flasherbase" & nr) : Set objlit(nr) = Eval("Flasherlit" & nr)
  Set objflasher(nr) = Eval("Flasherflash" & nr) : Set objlight(nr) = Eval("Flasherlight" & nr)
  Set objbloom(nr) = Eval("Flasherbloom" & nr)
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

  'rothbauerw
  'Adjust the position of the flasher object to align with the flasher base.
  'Comment out these lines if you want to manually adjust the flasher object
  If objbase(nr).roty > 135 then
    objflasher(nr).y = objbase(nr).y + 50
    objflasher(nr).height = objbase(nr).z + 20
  Else
    objflasher(nr).y = objbase(nr).y + 20
    objflasher(nr).height = objbase(nr).z + 50
  End If
  objflasher(nr).x = objbase(nr).x

  'rothbauerw
  'Adjust the position of the light object to align with the flasher base.
  'Comment out these lines if you want to manually adjust the flasher object
  objlight(nr).x = objbase(nr).x
  objlight(nr).y = objbase(nr).y
  objlight(nr).bulbhaloheight = objbase(nr).z +50

  'rothbauerw
  'Assign the appropriate bloom image basked on the location of the flasher base
  'Comment out these lines if you want to manually assign the bloom images
  dim xthird, ythird
  xthird = tablewidth/3
  ythird = tableheight/3

  If objbase(nr).x >= xthird and objbase(nr).x <= xthird*2 then
    objbloom(nr).imageA = "flasherbloomCenter"
    objbloom(nr).imageB = "flasherbloomCenter"
  elseif objbase(nr).x < xthird and objbase(nr).y < ythird then
    objbloom(nr).imageA = "flasherbloomUpperLeft"
    objbloom(nr).imageB = "flasherbloomUpperLeft"
  elseif  objbase(nr).x > xthird*2 and objbase(nr).y < ythird then
    objbloom(nr).imageA = "flasherbloomUpperRight"
    objbloom(nr).imageB = "flasherbloomUpperRight"
  elseif objbase(nr).x < xthird and objbase(nr).y < ythird*2 then
    objbloom(nr).imageA = "flasherbloomCenterLeft"
    objbloom(nr).imageB = "flasherbloomCenterLeft"
  elseif  objbase(nr).x > xthird*2 and objbase(nr).y < ythird*2 then
    objbloom(nr).imageA = "flasherbloomCenterRight"
    objbloom(nr).imageB = "flasherbloomCenterRight"
  elseif objbase(nr).x < xthird and objbase(nr).y < ythird*3 then
    objbloom(nr).imageA = "flasherbloomLowerLeft"
    objbloom(nr).imageB = "flasherbloomLowerLeft"
  elseif  objbase(nr).x > xthird*2 and objbase(nr).y < ythird*3 then
    objbloom(nr).imageA = "flasherbloomLowerRight"
    objbloom(nr).imageB = "flasherbloomLowerRight"
  end if

  ' set the texture and color of all objects
  select case objbase(nr).image
    Case "dome2basewhite" : objbase(nr).image = "dome2base" & col : objlit(nr).image = "dome2lit" & col :
    Case "ronddomebasewhite" : objbase(nr).image = "ronddomebase" & col : objlit(nr).image = "ronddomelit" & col
    Case "domeearbasewhite" : objbase(nr).image = "domeearbase" & col : objlit(nr).image = "domeearlit" & col
  end select
  If TestFlashers = 0 Then objflasher(nr).imageA = "domeflashwhite" : objflasher(nr).visible = 0 : End If
  select case col
    Case "blue" :   objlight(nr).color = RGB(4,120,255) : objflasher(nr).color = RGB(200,255,255) : objbloom(nr).color = RGB(4,120,255) : objlight(nr).intensity = 5000
    Case "green" :  objlight(nr).color = RGB(12,255,4) : objflasher(nr).color = RGB(12,255,4) : objbloom(nr).color = RGB(12,255,4)
    Case "red" :    objlight(nr).color = RGB(255,32,4) : objflasher(nr).color = RGB(255,32,4) : objbloom(nr).color = RGB(255,32,4)
    Case "purple" : objlight(nr).color = RGB(230,49,255) : objflasher(nr).color = RGB(255,64,255) : objbloom(nr).color = RGB(230,49,255)
    Case "yellow" : objlight(nr).color = RGB(200,173,25) : objflasher(nr).color = RGB(255,200,50) : objbloom(nr).color = RGB(200,173,25)
    Case "white" :  objlight(nr).color = RGB(255,240,150) : objflasher(nr).color = RGB(100,86,59) : objbloom(nr).color = RGB(255,240,150)
    Case "orange" :  objlight(nr).color = RGB(255,70,0) : objflasher(nr).color = RGB(255,70,0) : objbloom(nr).color = RGB(255,70,0)
  end select
  objlight(nr).colorfull = objlight(nr).color
  If TableRef.ShowDT and ObjFlasher(nr).RotX = -45 Then
    objflasher(nr).height = objflasher(nr).height - 20 * ObjFlasher(nr).y / tableheight
    ObjFlasher(nr).y = ObjFlasher(nr).y + 10
  End If
End Sub

Sub RotateFlasher(nr, angle) : angle = ((angle + 360 - objbase(nr).ObjRotZ) mod 180)/30 : objbase(nr).showframe(angle) : objlit(nr).showframe(angle) : End Sub


Sub ModFlashFlasher(nr, aValue)
  objflasher(nr).visible = 1 : objbloom(nr).visible = 1 : objlit(nr).visible = 1
  objflasher(nr).opacity = 1000 *  FlasherFlareIntensity * aValue
  objbloom(nr).opacity = 100 *  FlasherBloomIntensity * aValue
  objlight(nr).IntensityScale = 0.5 * FlasherLightIntensity * aValue
  objbase(nr).BlendDisableLighting =  FlasherOffBrightness + 10 * aValue
  objlit(nr).BlendDisableLighting = 10 * aValue
  UpdateMaterial "Flashermaterial" & nr,0,0,0,0,0,0,aValue,RGB(255,255,255),0,0,False,True,0,0,0,0
End Sub


'
'Sub FlashFlasher(nr)
' Dim flashx3
' If not objflasher(nr).TimerEnabled Then
'   objflasher(nr).TimerEnabled = True : objflasher(nr).visible = 1 : objbloom(nr).visible = 1 : objlit(nr).visible = 1
'   If VRRoom > 0 Then
'     select case nr
'       case 1
'       VRBGFL29_1.visible = 1
'       VRBGFL29_2.visible = 1
'       VRBGFL29_3.visible = 1
'       VRBGFL29_4.visible = 1
'       VRBGFL29_5.visible = 1
'       VRBGFL29_6.visible = 1
'       VRBGFL29_7.visible = 1
'       VRBGFL29_8.visible = 1
'       Case 2
'       VRBGFL31_1.visible = 1
'       VRBGFL31_2.visible = 1
'       VRBGFL31_3.visible = 1
'       VRBGFL31_4.visible = 1
'       Case 4
'       VRBGFL32_1.visible = 1
'       VRBGFL32_2.visible = 1
'       VRBGFL32_3.visible = 1
'       VRBGFL32_4.visible = 1
'     End Select
'   End If
'
'
' End If
' objflasher(nr).opacity = 1000 *  FlasherFlareIntensity * ObjLevel(nr)^2.5
' objbloom(nr).opacity = 100 *  FlasherBloomIntensity * ObjLevel(nr)^2.5
'
'   If VRRoom > 0 Then
'     flashx3 = objlevel(nr)^3
'     select case nr
'       case 1
'       VRBGFL29_1.opacity = 50 * flashx3^1.5
'       VRBGFL29_2.opacity = 50 * flashx3^1.5
'       VRBGFL29_3.opacity = 50 * flashx3^1.5
'       VRBGFL29_4.opacity = 50 * flashx3^2
'       VRBGFL29_5.opacity = 50 * flashx3^1.5
'       VRBGFL29_6.opacity = 50 * flashx3^1.5
'       VRBGFL29_7.opacity = 50 * flashx3^1.5
'       VRBGFL29_8.opacity = 50 * flashx3^2
'       Case 2
'       VRBGFL31_1.opacity = 200 * flashx3^1.5
'       VRBGFL31_2.opacity = 200 * flashx3^1.5
'       VRBGFL31_3.opacity = 200 * flashx3^1.5
'       VRBGFL31_4.opacity = 200 * flashx3^2
'       Case 4
'       VRBGFL32_1.opacity = 100 * flashx3^1.5
'       VRBGFL32_2.opacity = 100 * flashx3^1.5
'       VRBGFL32_3.opacity = 100 * flashx3^1.5
'       VRBGFL32_4.opacity = 100 * flashx3^2
'     End Select
'   End If
'
'
' If nr=1 Then
'   objlight(nr).IntensityScale = 0.5/3 * FlasherLightIntensity * ObjLevel(nr)^3
' Else
'   objlight(nr).IntensityScale = 0.5 * FlasherLightIntensity * ObjLevel(nr)^3
' End If
' objbase(nr).BlendDisableLighting =  FlasherOffBrightness + 12 * ObjLevel(nr)^3
' objlit(nr).BlendDisableLighting = 35 * ObjLevel(nr)^2
' UpdateMaterial "Flashermaterial" & nr,0,0,0,0,0,0,ObjLevel(nr),RGB(255,255,255),0,0,False,True,0,0,0,0
' ObjLevel(nr) = ObjLevel(nr) * 0.9 - 0.01
' If ObjLevel(nr) < 0 Then
'   objflasher(nr).TimerEnabled = False : objflasher(nr).visible = 0 : objbloom(nr).visible = 0 : objlit(nr).visible = 0
'     If VRRoom > 0 Then
'     select case nr
'       case 1
'       VRBGFL29_1.visible = 0
'       VRBGFL29_2.visible = 0
'       VRBGFL29_3.visible = 0
'       VRBGFL29_4.visible = 0
'       VRBGFL29_5.visible = 0
'       VRBGFL29_6.visible = 0
'       VRBGFL29_7.visible = 0
'       VRBGFL29_8.visible = 0
'       Case 2
'       VRBGFL31_1.visible = 0
'       VRBGFL31_2.visible = 0
'       VRBGFL31_3.visible = 0
'       VRBGFL31_4.visible = 0
'       Case 4
'       VRBGFL32_1.visible = 0
'       VRBGFL32_2.visible = 0
'       VRBGFL32_3.visible = 0
'       VRBGFL32_4.visible = 0
'     End Select
'   End If
' End If
'End Sub
'
'Sub FlasherFlash1_Timer() : FlashFlasher(1) : End Sub
'Sub FlasherFlash2_Timer() : FlashFlasher(2) : End Sub
'Sub FlasherFlash3_Timer() : FlashFlasher(3) : End Sub
'Sub FlasherFlash4_Timer() : FlashFlasher(4) : End Sub



'******************************************************
'******  END FLUPPER DOMES
'******************************************************






'***************************************************
'      Real Time Sub
'***************************************************
Dim FM, FMI, FMIV
FM = 0 'FlashMagnet On or off
FMI = 200 'FlashMagnet Variable Intensity
FMIV = 5 'FlashMagnet Add Variable Intensity

Sub RealTime_timer()
  If FantasyMagnet = 1 then
    If FM = 1 then
      FMI = FMI + FMIV
      FlashMagnet.IntensityScale = FMI/2
      FlashMagnet.Opacity = 2+(FMI/400)
      FlashMagnet.Amount = (100 / (FMI/6))
      FlashMagnet.rotz = FlashMagnet.rotz + (Int(Rnd*5)+1)
      if FMI > 800 then FMIV = -3:RandomFM:End If
      if FMI < 100 then FMIV = 3:RandomFM:End If
    else
      FlashMagnet.IntensityScale = 0
    End if
  End if

  BallsOnMagnet = Ubound(mFlipperMagnet.Balls) 'Test if there is a ball In Magnet THX To JP Salas !!!
  If BallsOnMagnet = 0 and mdof = 0 then DOF 104, DOFOn:mdof = 1:end If
  If BallsOnMagnet = -1 and mdof = 1 then DOF 104, DOFOff:mdof = 0:end If

  LFlipper.ObjRotZ = LeftFlipper.CurrentAngle -121
  RFlipper.ObjRotZ = RightFlipper.CurrentAngle +121
  LeftFlipperSh.RotZ = LeftFlipper.currentangle
  RightFlipperSh.RotZ = RightFlipper.currentangle
  GateA.Rotx = GateSw17.CurrentAngle
  GateB.Rotx = GateSw40.CurrentAngle
  GateC.Rotx = GateSw54.CurrentAngle
  GateD.Rotx = GateSw32.CurrentAngle
  GateE.Rotx = GateSw19.CurrentAngle
  GateF.Rotx = GateSw18.CurrentAngle
  If ll58.state = 1 then
    RedLightC.blenddisableLighting = 60
  Else
    RedLightC.blenddisableLighting = LightLevel*0.5
  End If
  If ll50.state = 1 then
'   RedLightA.blenddisableLighting = 40
    RedFA.Amount = 40
    RedFA.IntensityScale = 2
  Else
'   RedLightA.blenddisableLighting = 0.3
    RedFA.Amount = 0
    RedFA.IntensityScale = 0
  End If
  If ll51.state = 1 then
'   RedLightB.blenddisableLighting = 40
    RedFB.Amount = 30
    RedFB.IntensityScale = 2
  Else
'   RedLightB.blenddisableLighting = 0.3
    RedFB.Amount = 0
    RedFB.IntensityScale = 0
  End If
  If ll34.state = 1 then
'   GreenLight.blenddisableLighting = 20
    GreenF.Amount = 30
    GreenF.IntensityScale = 2
  Else
'   GreenLight.blenddisableLighting = 0.3
    GreenF.Amount = 0
    GreenF.IntensityScale = 0
  End If
  If ll49.state = 1 then

    StopR = 0:rotorspeed = 10:Rotor.enabled = 1
  end if
  If ll69.state = 1 Then
'   spotlightLight.blenddisablelighting = 0
    Flash69.Amount = 150
    flash69.IntensityScale = 50
  Else
'   spotlightLight.blenddisablelighting = 0 ' 40
    Flash69.Amount = 0
    flash69.IntensityScale = 0
  End If
End Sub


Sub RandomFM()
  Select Case Int(Rnd*4)+1
    Case 1 : FlashMagnet.imageA = "FlashTest03":FlashMagnet.imageB = "FlashTest01b"
    Case 2 : FlashMagnet.imageA = "FlashTest01":FlashMagnet.imageB = "FlashTest03b"
    Case 3 : FlashMagnet.imageA = "FlashTest03b":FlashMagnet.imageB = "FlashTest01"
    Case 4 : FlashMagnet.imageA = "FlashTest01b":FlashMagnet.imageB = "FlashTest03"
  End Select
End Sub



'*******************************************
' ZGIU:  GI Updates
'*******************************************


Const PFGIOFFOpacity = 100

'dim lightdir, Light
'Light=0
'lightdir = 0.1

ShadowGI.visible=1

set GICallback2 = GetRef("GIUpdates2")

Sub GIUpdates2(aNr, aLvl)
  dim xx
  'debug.print "GIUpdates2 nr: " & aNr & " value: " & aLvl

  For each xx in GI:xx.State = aLvl:  Next

  ' If the GI has an associated Relay sound, this can be played
  If aLvl >= 0.5 And gilvl < 0.5 Then
    Sound_GI_Relay 1, BottomTurboBumper
    If B2SOn Then Controller.B2SSetData 90, 1
    DOF 103, DOFOn
  ElseIf aLvl <= 0.4 And gilvl > 0.4 Then
    Sound_GI_Relay 0, BottomTurboBumper
    If B2SOn Then Controller.B2SSetData 90, 0
    DOF 103, DOFOff
  End If


  if aLvl = 0 then                  'GI OFF, let's hide ON prims
    OnPrimsVisible False
    If gilvl = 1 Then OffPrimsVisible true
  Elseif aLvl = 1 then                'GI ON, let's hide OFF prims
    OffPrimsVisible False
    If gilvl = 0 Then OnPrimsVisible True
  Else
    if gilvl = 0 Then               'GI has just changed from OFF to fading, let's show ON
      OnPrimsVisible True
    elseif gilvl = 1 Then             'GI has just changed from ON to fading, let's show OFF
      OffPrimsVisible true
    Else
      'no change
    end if
  end if

  ShadowGI.opacity = 70 * aLvl^2
'   Apron.image= "apron_on"

  Ledrouge.Amount = 100:LedRouge.IntensityScale = 10 * aLvl^2
  LedBleu1.Amount = 100:LedBleu1.IntensityScale = 10 * aLvl^2
  LedBleu2.Amount = 100:LedBleu2.IntensityScale = 10 * aLvl^2
  LedFond.Amount = 100:LedFond.IntensityScale = 200 * aLvl^2
  'UpdateMaterial(string,       float wrapLighting, float roughness,  float glossyImageLerp,    float thickness, float edge, float edgeAlpha, float opacity, OLE_COLOR base, OLE_COLOR glossy, OLE_COLOR clearcoat, VARIANT_BOOL isMetal, VARIANT_BOOL opacityActive, float elasticity, float elasticityFalloff, float friction, float scatterAngle)
  UpdateMaterial "PlasticsFade",  0,          1,          0,              0,        1,      0,          aLvl^2,   RGB(192,192,192),RGB(248,248,248),0,False,True,0,0,0,0
  UpdateMaterial "RampFade",    0,          0,          1,              0,        1,      0,          0.3 * aLvl^1, RGB(255,255,255),0,0,False,True,0,0,0,0
  UpdateMaterial "SidesFade",   0,          0,          1,              0,        1,      0,          aLvl^1, RGB(33,33,33),0,0,False,True,0,0,0,0

  playfieldOff.opacity = 100 * (1-alvl)
  gilvl = alvl
  SetPrimBrightness

End Sub

Sub SetPrimBrightness
  RWires.blenddisablelighting = LightLevel*0.4 * gilvl^2 + 0.05
  LWires.blenddisablelighting = LightLevel*0.4 * gilvl^2 + 0.05

  MrampB.blenddisablelighting = LightLevel*0.6 * gilvl^3
  Plastics.blenddisablelighting = LightLevel*gilvl
  BumperA.blenddisableLighting = LightLevel*0.8 * gilvl + 0.2
  BumperA001.blenddisableLighting = LightLevel*0.8 * gilvl + 0.2
  BumperA002.blenddisableLighting = LightLevel*0.8 * gilvl + 0.2
  Rampe3.blenddisableLighting = LightLevel*3.5 * gilvl + 0.5    'gi fading material
  Rampe2.blenddisableLighting = LightLevel*1.1 * gilvl + 0.3
  Rampe1.blenddisableLighting = LightLevel*3.7 * gilvl + 0.1
  VisA.blenddisablelighting = LightLevel*4 * gilvl
  VisB.blenddisablelighting = LightLevel*4 * gilvl
  MurMetal.blenddisablelighting = LightLevel*2 * gilvl^3
  MetalParts.blenddisablelighting = LightLevel*2 * gilvl^3
  spotlight.blenddisablelighting = LightLevel*39.5 * gilvl + 0.5
'   spotlightLight.blenddisablelighting = 0.5
  PegsBoverSlings.blenddisablelighting = LightLevel*gilvl
  Apron.blenddisablelighting = LightLevel*0.7 * gilvl
  ApronCachePlunger.blenddisablelighting = LightLevel*0.7 * gilvl^2
  dim xx: for each xx in DL:
    xx.blendDisableLighting = LightLevel*gilvl*0.85 + 0.15
  Next
End Sub


sub OnPrimsVisible(aValue)
  dim kk
  If aValue then
    For each kk in ON_Prims:kk.visible = 1:next
  Else
    For each kk in ON_Prims:kk.visible = 0:next
  end If
end Sub

sub OffPrimsVisible(aValue)
  dim kk
  If aValue then
    For each kk in OFF_Prims:kk.visible = 1:next
  Else
    For each kk in OFF_Prims:kk.visible = 0:next
  end If
end Sub





'******************************************************
'   Z3DI:   3D INSERTS
'******************************************************


Sub LL1_animate: p1.BlendDisableLighting = 400 * (LL1.GetInPlayIntensity / LL1.Intensity): End Sub
Sub LL2_animate: p2.BlendDisableLighting = 400 * (LL2.GetInPlayIntensity / LL2.Intensity): End Sub
Sub LL3_animate: p3.BlendDisableLighting = 400 * (LL3.GetInPlayIntensity / LL3.Intensity): End Sub
Sub LL4_animate: p4.BlendDisableLighting = 400 * (LL4.GetInPlayIntensity / LL4.Intensity): End Sub
Sub LL5_animate: p5.BlendDisableLighting = 400 * (LL5.GetInPlayIntensity / LL5.Intensity): End Sub
Sub LL6_animate: p6.BlendDisableLighting = 400 * (LL6.GetInPlayIntensity / LL6.Intensity): End Sub
Sub LL7_animate: p7.BlendDisableLighting = 400 * (LL7.GetInPlayIntensity / LL7.Intensity): End Sub
Sub LL8_animate: p8.BlendDisableLighting = 400 * (LL8.GetInPlayIntensity / LL8.Intensity): End Sub
'Sub LL9_animate: p9.BlendDisableLighting = 400 * (LL9.GetInPlayIntensity / LL9.Intensity): End Sub

Sub LL10_animate: p10.BlendDisableLighting = 400 * (LL10.GetInPlayIntensity / LL10.Intensity): End Sub
Sub LL11_animate: p11.BlendDisableLighting = 400 * (LL11.GetInPlayIntensity / LL11.Intensity): End Sub
Sub LL12_animate: p12.BlendDisableLighting = 400 * (LL12.GetInPlayIntensity / LL12.Intensity): End Sub
Sub LL13_animate: p13.BlendDisableLighting = 400 * (LL13.GetInPlayIntensity / LL13.Intensity): End Sub
Sub LL14_animate: p14.BlendDisableLighting = 480 * (LL14.GetInPlayIntensity / LL14.Intensity): End Sub
Sub LL15_animate: p15f.BlendDisableLighting = 400 * (LL15.GetInPlayIntensity / LL15.Intensity): End Sub
Sub LL16_animate: p16.BlendDisableLighting = 400 * (LL16.GetInPlayIntensity / LL16.Intensity): End Sub
Sub LL17_animate: p17.BlendDisableLighting = 400 * (LL17.GetInPlayIntensity / LL17.Intensity): End Sub
Sub LL18_animate: p18.BlendDisableLighting = 400 * (LL18.GetInPlayIntensity / LL18.Intensity): End Sub
'Sub LL19_animate: p19.BlendDisableLighting = 400 * (LL19.GetInPlayIntensity / LL19.Intensity): End Sub

Sub LL20_animate: p20.BlendDisableLighting = 400 * (LL20.GetInPlayIntensity / LL20.Intensity): End Sub
Sub LL21_animate: p21.BlendDisableLighting = 400 * (LL21.GetInPlayIntensity / LL21.Intensity): End Sub
Sub LL22_animate: p22.BlendDisableLighting = 400 * (LL22.GetInPlayIntensity / LL22.Intensity): End Sub
Sub LL23_animate: p23.BlendDisableLighting = 400 * (LL23.GetInPlayIntensity / LL23.Intensity): End Sub
Sub LL24_animate: p24.BlendDisableLighting = 400 * (LL24.GetInPlayIntensity / LL24.Intensity): End Sub
Sub LL25_animate: p25.BlendDisableLighting = 400 * (LL25.GetInPlayIntensity / LL25.Intensity): End Sub
Sub LL26_animate: p26.BlendDisableLighting = 400 * (LL26.GetInPlayIntensity / LL26.Intensity): End Sub
'Sub LL27_animate: p27.BlendDisableLighting = 400 * (LL27.GetInPlayIntensity / LL27.Intensity): End Sub
Sub LL28_animate: p28.BlendDisableLighting = 400 * (LL28.GetInPlayIntensity / LL28.Intensity): End Sub
Sub LL29_animate: p29.BlendDisableLighting = 400 * (LL29.GetInPlayIntensity / LL29.Intensity): End Sub

Sub LL30_animate: p30.BlendDisableLighting = 400 * (LL30.GetInPlayIntensity / LL30.Intensity): End Sub
Sub LL31_animate: p31.BlendDisableLighting = 400 * (LL31.GetInPlayIntensity / LL31.Intensity): End Sub
Sub LL32_animate: p32.BlendDisableLighting = 400 * (LL32.GetInPlayIntensity / LL32.Intensity): End Sub
Sub LL33_animate: p33.BlendDisableLighting = 400 * (LL33.GetInPlayIntensity / LL33.Intensity): End Sub
Sub LL34_animate: GreenLight.BlendDisableLighting = 100 * (LL34.GetInPlayIntensity / LL34.Intensity): End Sub
Sub LL35_animate: p35.BlendDisableLighting = 400 * (LL35.GetInPlayIntensity / LL35.Intensity): End Sub
Sub LL36_animate: p36.BlendDisableLighting = 400 * (LL36.GetInPlayIntensity / LL36.Intensity): End Sub
Sub LL37_animate: p37.BlendDisableLighting = 400 * (LL37.GetInPlayIntensity / LL37.Intensity): End Sub
Sub LL38_animate: p38.BlendDisableLighting = 400 * (LL38.GetInPlayIntensity / LL38.Intensity): End Sub
Sub LL39_animate: p39.BlendDisableLighting = 400 * (LL39.GetInPlayIntensity / LL39.Intensity): End Sub

'Sub LL40_animate: p40.BlendDisableLighting = 400 * (LL40.GetInPlayIntensity / LL40.Intensity): End Sub
'Sub LL41_animate: p41.BlendDisableLighting = 400 * (LL41.GetInPlayIntensity / LL41.Intensity): End Sub
Sub LL42_animate: p42.BlendDisableLighting = 400 * (LL42.GetInPlayIntensity / LL42.Intensity): End Sub
Sub LL43_animate: p43.BlendDisableLighting = 400 * (LL43.GetInPlayIntensity / LL43.Intensity): End Sub
Sub LL44_animate: p44.BlendDisableLighting = 400 * (LL44.GetInPlayIntensity / LL44.Intensity): End Sub
Sub LL45_animate: p45.BlendDisableLighting = 400 * (LL45.GetInPlayIntensity / LL45.Intensity): End Sub
Sub LL46_animate: p46.BlendDisableLighting = 400 * (LL46.GetInPlayIntensity / LL46.Intensity): End Sub
Sub LL47_animate: p47.BlendDisableLighting = 400 * (LL47.GetInPlayIntensity / LL47.Intensity): End Sub
Sub LL48_animate: p48.BlendDisableLighting = 400 * (LL48.GetInPlayIntensity / LL48.Intensity): End Sub
'Sub LL49_animate: p49.BlendDisableLighting = 400 * (LL49.GetInPlayIntensity / LL49.Intensity): End Sub

Sub LL50_animate: RedLightA.BlendDisableLighting = 200 * (LL50.GetInPlayIntensity / LL50.Intensity): End Sub
Sub LL51_animate: RedLightB.BlendDisableLighting = 200 * (LL51.GetInPlayIntensity / LL51.Intensity): End Sub
Sub LL52_animate: p52.BlendDisableLighting = 400 * (LL52.GetInPlayIntensity / LL52.Intensity): End Sub
Sub LL53_animate: p53.BlendDisableLighting = 400 * (LL53.GetInPlayIntensity / LL53.Intensity): End Sub
Sub LL54_animate: p54.BlendDisableLighting = 400 * (LL54.GetInPlayIntensity / LL54.Intensity): End Sub
Sub LL55_animate: p55.BlendDisableLighting = 400 * (LL55.GetInPlayIntensity / LL55.Intensity): End Sub
Sub LL56_animate: p56.BlendDisableLighting = 400 * (LL56.GetInPlayIntensity / LL56.Intensity): End Sub
'Sub LL57_animate: p57.BlendDisableLighting = 400 * (LL57.GetInPlayIntensity / LL57.Intensity): End Sub
'Sub LL58_animate: p58.BlendDisableLighting = 400 * (LL58.GetInPlayIntensity / LL58.Intensity): End Sub
'Sub LL59_animate: p59.BlendDisableLighting = 400 * (LL59.GetInPlayIntensity / LL59.Intensity): End Sub

Sub LL60_animate: p60.BlendDisableLighting = 400 * (LL60.GetInPlayIntensity / LL60.Intensity): End Sub
Sub LL61_animate: p61.BlendDisableLighting = 400 * (LL61.GetInPlayIntensity / LL61.Intensity): End Sub
Sub LL62_animate: p62.BlendDisableLighting = 400 * (LL62.GetInPlayIntensity / LL62.Intensity): End Sub
Sub LL63_animate: p63.BlendDisableLighting = 400 * (LL63.GetInPlayIntensity / LL63.Intensity): End Sub
'Sub LL64_animate: p64.BlendDisableLighting = 400 * (LL64.GetInPlayIntensity / LL64.Intensity): End Sub
Sub LL65_animate: p65.BlendDisableLighting = 400 * (LL65.GetInPlayIntensity / LL65.Intensity): End Sub
Sub LL66_animate: p66.BlendDisableLighting = 400 * (LL66.GetInPlayIntensity / LL66.Intensity): End Sub
Sub LL67_animate: p67.BlendDisableLighting = 400 * (LL67.GetInPlayIntensity / LL67.Intensity): End Sub
Sub LL68_animate: p68.BlendDisableLighting = 400 * (LL68.GetInPlayIntensity / LL68.Intensity): End Sub
Sub LL69_animate: spotlightLight.BlendDisableLighting = 100 * (LL69.GetInPlayIntensity / LL69.Intensity): End Sub






'*******************************************
' ZSWI:  Switches
'*******************************************

'*********************** Targets **********************
Sub SS25_hit():SS25c.rotx = 7:SS25c.roty = -7:Me.TimerEnabled = 1:vpmTimer.PulseSw 25:End Sub
Sub SS25_Timer():SS25c.rotx = 0:SS25c.roty = 0:Me.TimerEnabled = 0:End Sub
Sub SS26_hit():SS26c.rotx = 7:SS26c.roty = -7:Me.TimerEnabled = 1:vpmTimer.PulseSw 26:End Sub
Sub SS26_Timer():SS26c.rotx = 0:SS26c.roty = 0:Me.TimerEnabled = 0:End Sub
Sub SS27_hit():SS27c.rotx = 7:SS27c.roty = -7:Me.TimerEnabled = 1:vpmTimer.PulseSw 27:End Sub
Sub SS27_Timer():SS27c.rotx = 0:SS27c.roty = 0:Me.TimerEnabled = 0:End Sub
Sub SS28_hit():SS28c.rotx = 7:SS28c.roty = -7:Me.TimerEnabled = 1:vpmTimer.PulseSw 28:End Sub
Sub SS28_Timer():SS28c.rotx = 0:SS28c.roty = 0:Me.TimerEnabled = 0:End Sub
Sub SS30_hit():SS30c.rotx = 7:SS30c.roty = -2.5:Me.TimerEnabled = 1:vpmTimer.PulseSw 30:End Sub
Sub SS30_Timer():SS30c.rotx = 0:SS30c.roty = 0:Me.TimerEnabled = 0:End Sub
Sub SS31_hit():SS31c.rotx = 8:SS31c.roty = -1:Me.TimerEnabled = 1:vpmTimer.PulseSw 31:End Sub
Sub SS31_Timer():SS31c.rotx = 0:SS31c.roty = 0:Me.TimerEnabled = 0:End Sub
Sub SS33_hit():SS33c.rotx = 3:SS33c.roty = 7:Me.TimerEnabled = 1:vpmTimer.PulseSw 33:End Sub
Sub SS33_Timer():SS33c.rotx = 0:SS33c.roty = 0:Me.TimerEnabled = 0:End Sub
Sub SS34_hit():SS34c.rotx = 3:SS34c.roty = 7:Me.TimerEnabled = 1:vpmTimer.PulseSw 34:End Sub
Sub SS34_Timer():SS34c.rotx = 0:SS34c.roty = 0:Me.TimerEnabled = 0:End Sub
Sub SS39_hit():SS39c.rotx = 7:SS39c.roty = 0:Me.TimerEnabled = 1:vpmTimer.PulseSw 39:End Sub
Sub SS39_Timer():SS39c.rotx = 0:SS39c.roty = 0:Me.TimerEnabled = 0:End Sub
Sub SS44_hit():SS44c.rotx = 6:SS44c.roty = 6:Me.TimerEnabled = 1:vpmTimer.PulseSw 44:End Sub
Sub SS44_Timer():SS44c.rotx = 0:SS44c.roty = 0:Me.TimerEnabled = 0:End Sub
Sub SS45_hit():SS45c.rotx = 6:SS45c.roty = 6:Me.TimerEnabled = 1:vpmTimer.PulseSw 45:End Sub
Sub SS45_Timer():SS45c.rotx = 0:SS45c.roty = 0:Me.TimerEnabled = 0:End Sub
Sub SS46_hit():SS46c.rotx = 6:SS46c.roty = 6:Me.TimerEnabled = 1:vpmTimer.PulseSw 46:End Sub
Sub SS46_Timer():SS46c.rotx = 0:SS46c.roty = 0:Me.TimerEnabled = 0:End Sub
Sub SS47_hit():SS47c.rotx = 6:SS47c.roty = 6:Me.TimerEnabled = 1:vpmTimer.PulseSw 47:End Sub
Sub SS47_Timer():SS47c.rotx = 0:SS47c.roty = 0:Me.TimerEnabled = 0:End Sub
Sub SS48_hit():SS48c.rotx = 6:SS48c.roty = 6:Me.TimerEnabled = 1:vpmTimer.PulseSw 48:End Sub
Sub SS48_Timer():SS48c.rotx = 0:SS48c.roty = 0:Me.TimerEnabled = 0:End Sub


Sub sw16_Hit:Controller.Switch(16) = 1  :
  if DesktopMode and vrroom = 0 then
    L007_1.state=2
    L007_2.state=2
    L007_3.state=2
    Timer001.enabled=True
  End If
End Sub
Sub Timer001_Timer
  Timer001.enabled=False
  L007_1.state=0
  L007_2.state=0
  L007_3.state=0
End Sub
Sub sw16_UnHit:Controller.Switch(16) = 0:End Sub
Sub sw49_Hit:Controller.Switch(49) = 1 : End Sub
Sub sw49_UnHit:Controller.Switch(49) = 0:End Sub

Sub sw51_Hit:Controller.Switch(51) = 1 : End Sub
Sub sw51_UnHit:Controller.Switch(51) = 0:End Sub
Sub sw52_Hit:Controller.Switch(52) = 1 : End Sub
Sub sw52_UnHit:Controller.Switch(52) = 0:End Sub
Sub sw53_Hit:Controller.Switch(53) = 1 : End Sub
Sub sw53_UnHit:Controller.Switch(53) = 0:End Sub


Sub sw55_Hit:Controller.Switch(55) = 1 : End Sub
Sub sw55_UnHit:Controller.Switch(55) = 0:End Sub
Sub sw57_Hit:Controller.Switch(57) = 1 : End Sub
Sub sw57_UnHit:Controller.Switch(57) = 0:End Sub
Sub sw58_Hit:Controller.Switch(58) = 1 : End Sub
Sub sw58_UnHit:Controller.Switch(58) = 0:End Sub
Sub sw59_Hit:Controller.Switch(59) = 1 : End Sub
Sub sw59_UnHit:Controller.Switch(59) = 0:End Sub
Sub sw60_Hit:Controller.Switch(60) = 1: End Sub
Sub sw60_UnHit:Controller.Switch(60) = 0:End Sub

sub GateSw17_Hit:vpmTimer.PulseSw 17: End Sub
sub GateSw18_Hit:vpmTimer.PulseSw 18: End Sub
sub GateSw19_Hit:vpmTimer.PulseSw 19: End Sub
sub GateSw32_Hit:vpmTimer.PulseSw 32: WireRampOff: WireRampOn False: End Sub
sub GateSw40_Hit:vpmTimer.PulseSw 40: End Sub
sub GateSw54_Hit:vpmTimer.PulseSw 54: End Sub






'********** Sling Shot Animations *******************************
' Rstep and Lstep  are the variables that increment the animation
'****************************************************************
Dim RStep, Lstep

Sub RightSlingShot_Slingshot
    RS.VelocityCorrect(ActiveBall)
  vpmTimer.PulseSw 62
  RandomSoundSlingshotRight sling1
' RightSlingFlash.state =0
  RSling.Visible = 0
  RSling1.Visible = 1
  sling1.TransZ = -26
  RStep = 0
  RightSlingShot.TimerEnabled = 1
End Sub

Sub RightSlingShot_Timer
  Select Case RStep
    Case 3:RSLing1.Visible = 0:RSLing2.Visible = 1:sling1.TransZ = -16
    Case 4:RSLing2.Visible = 0:RSLing.Visible = 1:sling1.TransZ = 0:RightSlingShot.TimerEnabled = 0
  End Select
  RStep = RStep + 1
' RightSlingFlash.state =1
' if RSLing1.Visible = 0 then RightSlingFlash.state =0:end If
End Sub

Sub LeftSlingShot_Slingshot
    LS.VelocityCorrect(ActiveBall)
  vpmTimer.PulseSw 61
  RandomSoundSlingshotLeft sling2
' LeftSlingFlash.state = 0
  LSling.Visible = 0
  LSling1.Visible = 1
  sling2.TransZ = -26
  LStep = 0
  LeftSlingShot.TimerEnabled = 1
End Sub

Sub LeftSlingShot_Timer
  Select Case LStep
    Case 3:LSLing1.Visible = 0:LSLing2.Visible = 1:sling2.TransZ = -16
    Case 4:LSLing2.Visible = 0:LSLing.Visible = 1:sling2.TransZ = 0:LeftSlingShot.TimerEnabled = 0
  End Select
  LStep = LStep + 1
End Sub





'***********************************************************************************
'*************                 Bumpers                               ***************
'***********************************************************************************
Sub LeftTurboBumper_Hit
  vpmTimer.PulseSw 41
  'PlaySoundAtVol SoundFX("fx_bumper1",DOFContactors), ActiveBall, 1
  'PlaySoundAtVol SoundFX("bumper",DOFContactors), ActiveBall, 1
  RandomSoundBumperTop LeftTurboBumper
  Me.TimerEnabled = 1
  ' PlaySoundAtVol "bumper", LeftTurboBumper, 1
End Sub

Sub LeftTurboBumper_Timer
  Me.Timerenabled = 0
End Sub

Sub BottomTurboBumper_Hit
  vpmTimer.PulseSw 42
  'PlaySoundAtVol SoundFX("fx_bumper2",DOFContactors), ActiveBall, 1
  'PlaySoundAtVol SoundFX("bumper",DOFContactors), ActiveBall, 1
  RandomSoundBumperBottom BottomTurboBumper
  Me.TimerEnabled = 1
  ' PlaySoundAtVol "bumper", BottomTurboBumper, 1
End Sub

Sub BottomTurboBumper_Timer
  Me.Timerenabled = 0
End Sub

Sub RightTurboBumper_Hit
  vpmTimer.PulseSw 43
  'PlaySoundAtVol SoundFX("fx_bumper3",DOFContactors), ActiveBall, 1
  'PlaySoundAtVol SoundFX("bumper",DOFContactors), ActiveBall, 1
  RandomSoundBumperMiddle RightTurboBumper
  Me.TimerEnabled = 1
  ' PlaySoundAtVol "bumper", RightTurboBumper, 1
End Sub

Sub RightTurboBumper_Timer
  Me.Timerenabled = 0
End Sub
'**********************************************************************************


'************************** Left Ramp Helper ****************************************
Sub Trigger1_Hit
  If ActiveBall.VelY<9 and activeBall.velY>0 Then ActiveBall.VelY=ActiveBall.VelY+9
End Sub

'************************************************************************************



Sub Drain_Hit: bsTrough.AddBall Me: RandomSoundDrain Drain:End Sub
Sub Scoop_Hit :bsScoop.AddBall Me:PlaySound "Scoopenter",1,VolumeDial:End Sub
Sub TankKickBig_Hit :bsTank.AddBall Me:End Sub
Sub DropRamp1_Hit :bsTank.AddBall Me
  PlaySound "Kicker_enter",1,VolumeDial
End Sub
Sub Plunger_Hit:bsPlunger.Addball 0:End Sub
'Sub Kicker1_Hit:bsLockOut.AddBall 0:End Sub


'***************** Radar ********************
Controller.Switch(20) = false

Sub SolSatMotorRelay(enabled)   'rotate satellite
  'debug.print "SolSatMotorRelay "&Enabled
  if enabled then
    RadarFrameLast = gametime
    RotRadar.enabled = 1
    PlaySoundAtVolLoops "Motor", RadarA, 0.02 , -1
  else
    RotRadar.enabled = 0
    StopSound "Motor"
    Controller.Switch(23) = false
  End If
End Sub

Const RadarOmega = 60 'deg/s
Dim RadarTheta: RadarTheta = 0
Dim RadarAng: RadarAng = 0
Dim RadarCnt: RadarCnt = 0
Dim RadarPause: RadarPause = False
Dim RadarFrameTime: RadarFrameTime = 0
Dim RadarFrameLast: RadarFrameLast = 0
RotRadar.Interval = -1
RotRadar.enabled = 1

sub RotRadar_Timer()
  RadarFrameTime = gametime - RadarFrameLast
  RadarFrameLast = gametime
  If RadarPause=False Then
    'Calculate radar angle over time
    RadarTheta = RadarTheta + RadarOmega*RadarFrameTime/1000
    if RadarTheta >= 360 Then RadarTheta = RadarTheta - 360
    'debug.print "RadarTheta = "&RadarTheta
    RadarAng = 30*dSin(RadarTheta) - 10
    RadarA.objrotz = RadarAng
    RadarB.objrotz = RadarAng
    'Pause at ends
    RadarCnt = RadarCnt - RadarFrameTime
    If (RadarAng > 19.8 Or RadarAng < -39.8) And RadarCnt < 0 Then
      RadarPause = True
      RadarCnt = 0
    End If
    If (RadarTheta > 357 Or RadarTheta < 3) And Controller.Switch(20) = false Then
      Controller.Switch(20) = true
    ElseIf (RadarTheta < 357 And RadarTheta > 3) And Controller.Switch(20) = true Then
      Controller.Switch(20) = false
    End If

  Else
    RadarCnt = RadarCnt + RadarFrameTime
    If RadarCnt >= 1200 Then RadarPause = False
  End If
End Sub



'***************** Radar Ramp ********************
Dim RampC,RampDirection
RampC=0 '3 time to move up
RampDirection=-15'Ramp Will Move UP positive value it will move down


Sub SolSatLaunchRamp(Enabled)
  If Enabled Then
    RampDirection=-15:RadarRampAnim.enabled = 1
    RampRadar.collidable=1
    SatTop.IsDropped=0
    SatTop1.IsDropped=0
    SatTop2.IsDropped=0
    PlaySoundAtVol "RadarRampOn", RampSat3D, 1
  Else
    RampDirection=15:RadarRampAnim.enabled = 1
    RampRadar.collidable=0
    SatTop.IsDropped=1
    SatTop1.IsDropped=1
    SatTop2.IsDropped=1
    PlaySoundAtVol "RadarRampOff", RampSat3D, 1
  End If
End Sub

sub RadarRampAnim_Timer()
  RampSat3D.objrotx = RampSat3D.objrotx + RampDirection
  RampSat3Dv.objrotx =  RampSat3Dv.objrotx+ RampDirection
  RampC = RampC + 1
  if RampC = 3 and RampDirection=-15 then RampSat3D.objrotx = -45:RampSat3Dv.objrotx = -45:RampC = 0:me.enabled = 0:LL40.State = 1:end If
  if RampC = 3 and RampDirection=15 then RampSat3D.objrotx = 0:RampSat3Dv.objrotx = 0:RampC = 0:me.enabled = 0:LL40.State = 0:end If
End Sub

'****************** Radar Magnet **********************
dim RadarBall
dim BallInRadarLock : BallInRadarLock = False

Sub SolRadarMagnet(Enabled)
  If Not Enabled Then
    BallInRadarLock = False
    RadarKicker.Kick 0,1
    Controller.Switch(23)=0
    StopSound "fx_magnetR"
  End If
End Sub

Sub RMagnet_Hit
  if Not BallInRadarLock Then
    if MagnetVolMax = 1 then
      PlaySound "fx_magnetR",-1,MagnetVol
    else
      PlaySound "fx_magnetR",-1
    end if
    ActiveBall.velX=0
    ActiveBall.velY=0
    ActiveBall.velZ=0
    ActiveBall.X=637
    ActiveBall.Y=768
    ActiveBall.Z=152
    PlaySoundAtVol "magnethit", ActiveBall, VolumeDial
  End If
End Sub

Sub RadarKicker_hit
  Set RadarBall = Activeball
  BallInRadarLock = True
  Controller.Switch(23)=1
End Sub


'********************** Copter ********************

dim rotorspeed, StopR
rotorspeed = 10
StopR = 0
sub Rotor_timer()
  CopterC.roty = CopterC.roty + rotorspeed
  CopterD.rotx = CopterD.rotx + rotorspeed
  Copter001.image = "CopterBOn"
  if ll49.state = 0 then StopR = 1:Copter001.image = "CopterB":end if
  if StopR = 1 then rotorspeed=rotorspeed -0.1:end If
  if rotorspeed < 0 then me.enabled = 0:StopR = 0:rotorspeed = 10:end If
End Sub

'**************************************************

'*********Moving Ramp Solenoid 18 *****************

Sub SolUpDownRamp(Enabled)
  If Enabled Then
    Ramp005.collidable = 0
    Mramp3.collidable = 0
    Mramp.collidable = 1
    Mramp2.collidable = 1
    MrampB.objrotx = 17
    MrampA.objrotx = 17
    PlaySoundAtVol "RampDown", MrampB, 1
  Else
    Mramp.collidable = 0
    Mramp2.collidable = 0
    MrampB.objrotx = 1
    MrampA.objrotx = 1
    Ramp005.collidable = 1
    Mramp3.collidable = 1
    PlaySoundAtVol "RampUp", MrampB, 1
  End If
End Sub


Sub BallRelease_Unhit()
  RandomSoundBallRelease BallRelease
End Sub

Sub Scoop_Unhit()
End Sub

Sub TankKickBig_Unhit()
  TankAnim.enabled = 1
End Sub

dim TKA
TKA = 1
'Sub TankAnim_timer()
' select case TKA
'   case 1:
'     TankCTourelle.y = TankCTourelle.y + 3
'   case 2:
'     TankCTourelle.y = TankCTourelle.y - 15
'   case 3:
'     TankCTourelle.y = TankCTourelle.y - 5
'     TankA.y = TankA.y - 10
'     TankB.y = TankB.y - 10
'   case 4:
'     TankCTourelle.y = TankCTourelle.y - 3
'     TankA.y = TankA.y + 3
'     TankB.y = TankB.y + 3
'   case 5:
'     TankCTourelle.y = TankCTourelle.y + 5
'     TankA.y = TankA.y + 1
'     TankB.y = TankB.y + 1
'   case 6:
'     TankCTourelle.y = TankCTourelle.y + 5
'     TankA.y = TankA.y + 2
'     TankB.y = TankB.y + 2
'   case 7:
'     TankCTourelle.y = TankCTourelle.y + 5
'     TankA.y = TankA.y + 2
'     TankB.y = TankB.y + 2
'   case 7:
'     TankCTourelle.y = TankCTourelle.y + 3
'     TankA.y = TankA.y + 1
'     TankB.y = TankB.y + 1
'   case 7:
'     TankCTourelle.y = TankCTourelle.y + 2
'     TankA.y = TankA.y + 1
'     TankB.y = TankB.y + 1
' end select
'
' if TKA = 20 then TKA = 1:TankAnim.enabled = 0:exit sub:end if
' TKA = TKA + 1
'End Sub

Sub TankAnim_timer()
  select case TKA
    case 1:
      TankCTourelle.y = TankCTourelle.y - 4
    case 2:
      TankCTourelle.y = TankCTourelle.y + 29
    case 3:
      TankCTourelle.y = TankCTourelle.y + 12
      TankA.y = TankA.y + 20
      TankB.y = TankB.y + 20
    case 4:
      TankCTourelle.y = TankCTourelle.y + 3
      TankA.y = TankA.y - 6
      TankB.y = TankB.y - 6
    case 5:
      TankCTourelle.y = TankCTourelle.y - 10
      TankA.y = TankA.y - 2
      TankB.y = TankB.y - 2
    case 6:
      TankCTourelle.y = TankCTourelle.y - 8
      TankA.y = TankA.y - 4
      TankB.y = TankB.y - 4
    case 7:
      TankCTourelle.y = TankCTourelle.y - 8
      TankA.y = TankA.y - 4
      TankB.y = TankB.y - 4
    case 7:
      TankCTourelle.y = TankCTourelle.y - 6
      TankA.y = TankA.y - 4
      TankB.y = TankB.y - 4
    case 7:
      TankCTourelle.y = TankCTourelle.y - 4
      TankA.y = TankA.y - 2
      TankB.y = TankB.y - 2
    case 8:
      TankCTourelle.y = TankCTourelle.y - 4
      TankA.y = TankA.y + 2
      TankB.y = TankB.y + 2
  end select

  if TKA = 10 then TKA = 1:TankAnim.enabled = 0:exit sub:end if
  TKA = TKA + 1
End Sub



'******************************************************
' ZBRL:  BALL ROLLING AND DROP SOUNDS
'******************************************************

' Be sure to call RollingUpdate in a timer with a 10ms interval see the GameTimer_Timer() sub

ReDim rolling(tnob)
InitRolling

Dim DropCount
ReDim DropCount(tnob)

Sub InitRolling
  Dim i
  For i = 0 To tnob
    rolling(i) = False
  Next
End Sub

Sub RollingUpdate()
  Dim b
  Dim gBOT
  gBOT = GetBalls

  ' stop the sound of deleted balls
  For b = UBound(gBOT) + 1 To tnob - 1
    rolling(b) = False
    StopSound("BallRoll_" & b)
  Next

  ' exit the sub if no balls on the table
  If UBound(gBOT) =  - 1 Then Exit Sub

  ' play the rolling sound for each ball
  For b = 0 To UBound(gBOT)
    If BallVel(gBOT(b)) > 1 And gBOT(b).z < 30 Then
      rolling(b) = True
      PlaySound ("BallRoll_" & b), - 1, VolPlayfieldRoll(gBOT(b)) * BallRollVolume * VolumeDial, AudioPan(gBOT(b)), 0, PitchPlayfieldRoll(gBOT(b)), 1, 0, AudioFade(gBOT(b))
    Else
      If rolling(b) = True Then
        StopSound("BallRoll_" & b)
        rolling(b) = False
      End If
    End If

    ' Ball Drop Sounds
    If gBOT(b).VelZ <  - 1 And gBOT(b).z < 55 And gBOT(b).z > 27 Then 'height adjust for ball drop sounds
      If DropCount(b) >= 5 Then
        DropCount(b) = 0
        If gBOT(b).velz >  - 7 Then
          RandomSoundBallBouncePlayfieldSoft gBOT(b)
        Else
          RandomSoundBallBouncePlayfieldHard gBOT(b)
        End If
      End If
    End If

    If DropCount(b) < 5 Then
      DropCount(b) = DropCount(b) + 1
    End If
  Next
End Sub




'******************************************************
'   ZRRL: RAMP ROLLING SFX
'******************************************************

'Ball tracking ramp SFX 1.0
'   Reqirements:
'     * Import A Sound File for each ball on the table for plastic ramps.  Call It RampLoop<Ball_Number> ex: RampLoop1, RampLoop2, ...
'     * Import a Sound File for each ball on the table for wire ramps. Call it WireLoop<Ball_Number> ex: WireLoop1, WireLoop2, ...
'     * Create a Timer called RampRoll, that is enabled, with a interval of 100
'     * Set RampBAlls and RampType variable to Total Number of Balls
' Usage:
'     * Setup hit events and call WireRampOn True or WireRampOn False (True = Plastic ramp, False = Wire Ramp)
'     * To stop tracking ball
'        * call WireRampOff
'        * Otherwise, the ball will auto remove if it's below 30 vp units
'

Dim RampMinLoops
RampMinLoops = 4

' RampBalls
' Setup:  Set the array length of x in RampBalls(x,2) Total Number of Balls on table + 1:  if tnob = 5, then RampBalls(6,2)
Dim RampBalls(6,2)
'x,0 = ball x,1 = ID, 2 = Protection against ending early (minimum amount of updates)

'0,0 is boolean on/off, 0,1 unused for now
RampBalls(0,0) = False

' RampType
' Setup: Set this array to the number Total number of balls that can be tracked at one time + 1.  5 ball multiball then set value to 6
' Description: Array type indexed on BallId and a values used to deterimine what type of ramp the ball is on: False = Wire Ramp, True = Plastic Ramp
Dim RampType(6)

Sub WireRampOn(input)
  Waddball ActiveBall, input
  RampRollUpdate
End Sub

Sub WireRampOff()
  WRemoveBall ActiveBall.ID
End Sub

' WaddBall (Active Ball, Boolean)
Sub Waddball(input, RampInput) 'This subroutine is called from WireRampOn to Add Balls to the RampBalls Array
  ' This will loop through the RampBalls array checking each element of the array x, position 1
  ' To see if the the ball was already added to the array.
  ' If the ball is found then exit the subroutine
  Dim x
  For x = 1 To UBound(RampBalls)  'Check, don't add balls twice
    If RampBalls(x, 1) = input.id Then
      If Not IsEmpty(RampBalls(x,1) ) Then Exit Sub 'Frustating issue with BallId 0. Empty variable = 0
    End If
  Next

  ' This will itterate through the RampBalls Array.
  ' The first time it comes to a element in the array where the Ball Id (Slot 1) is empty.  It will add the current ball to the array
  ' The RampBalls assigns the ActiveBall to element x,0 and ball id of ActiveBall to 0,1
  ' The RampType(BallId) is set to RampInput
  ' RampBalls in 0,0 is set to True, this will enable the timer and the timer is also turned on
  For x = 1 To UBound(RampBalls)
    If IsEmpty(RampBalls(x, 1)) Then
      Set RampBalls(x, 0) = input
      RampBalls(x, 1) = input.ID
      RampType(x) = RampInput
      RampBalls(x, 2) = 0
      'exit For
      RampBalls(0,0) = True
      RampRoll.Enabled = 1   'Turn on timer
      'RampRoll.Interval = RampRoll.Interval 'reset timer
      Exit Sub
    End If
    If x = UBound(RampBalls) Then  'debug
      Debug.print "WireRampOn error, ball queue Is full: " & vbNewLine & _
      RampBalls(0, 0) & vbNewLine & _
      TypeName(RampBalls(1, 0)) & " ID:" & RampBalls(1, 1) & "type:" & RampType(1) & vbNewLine & _
      TypeName(RampBalls(2, 0)) & " ID:" & RampBalls(2, 1) & "type:" & RampType(2) & vbNewLine & _
      TypeName(RampBalls(3, 0)) & " ID:" & RampBalls(3, 1) & "type:" & RampType(3) & vbNewLine & _
      TypeName(RampBalls(4, 0)) & " ID:" & RampBalls(4, 1) & "type:" & RampType(4) & vbNewLine & _
      TypeName(RampBalls(5, 0)) & " ID:" & RampBalls(5, 1) & "type:" & RampType(5) & vbNewLine & _
      " "
    End If
  Next
End Sub

' WRemoveBall (BallId)
Sub WRemoveBall(ID) 'This subroutine is called from the RampRollUpdate subroutine and is used to remove and stop the ball rolling sounds
  '   Debug.Print "In WRemoveBall() + Remove ball from loop array"
  Dim ballcount
  ballcount = 0
  Dim x
  For x = 1 To UBound(RampBalls)
    If ID = RampBalls(x, 1) Then 'remove ball
      Set RampBalls(x, 0) = Nothing
      RampBalls(x, 1) = Empty
      RampType(x) = Empty
      StopSound("RampLoop" & x)
      StopSound("wireloop" & x)
    End If
    'if RampBalls(x,1) = Not IsEmpty(Rampballs(x,1) then ballcount = ballcount + 1
    If Not IsEmpty(Rampballs(x,1)) Then ballcount = ballcount + 1
  Next
  If BallCount = 0 Then RampBalls(0,0) = False  'if no balls in queue, disable timer update
End Sub

Sub RampRoll_Timer()
  RampRollUpdate
End Sub

Sub RampRollUpdate()  'Timer update
  Dim x
  For x = 1 To UBound(RampBalls)
    If Not IsEmpty(RampBalls(x,1) ) Then
      If BallVel(RampBalls(x,0) ) > 1 Then ' if ball is moving, play rolling sound
        If RampType(x) Then
          PlaySound("RampLoop" & x), - 1, VolPlayfieldRoll(RampBalls(x,0)) * RampRollVolume * VolumeDial, AudioPan(RampBalls(x,0)), 0, BallPitchV(RampBalls(x,0)), 1, 0, AudioFade(RampBalls(x,0))
          StopSound("wireloop" & x)
        Else
          StopSound("RampLoop" & x)
          PlaySound("wireloop" & x), - 1, VolPlayfieldRoll(RampBalls(x,0)) * RampRollVolume * VolumeDial, AudioPan(RampBalls(x,0)), 0, BallPitch(RampBalls(x,0)), 1, 0, AudioFade(RampBalls(x,0))
        End If
        RampBalls(x, 2) = RampBalls(x, 2) + 1
      Else
        StopSound("RampLoop" & x)
        StopSound("wireloop" & x)
      End If
      If RampBalls(x,0).Z < 30 And RampBalls(x, 2) > RampMinLoops Then  'if ball is on the PF, remove  it
        StopSound("RampLoop" & x)
        StopSound("wireloop" & x)
        Wremoveball RampBalls(x,1)
      End If
    Else
      StopSound("RampLoop" & x)
      StopSound("wireloop" & x)
    End If
  Next
  If Not RampBalls(0,0) Then RampRoll.enabled = 0
End Sub

' This can be used to debug the Ramp Roll time.  You need to enable the tbWR timer on the TextBox
Sub tbWR_Timer()  'debug textbox
  Me.text = "on? " & RampBalls(0, 0) & " timer: " & RampRoll.Enabled & vbNewLine & _
  "1 " & TypeName(RampBalls(1, 0)) & " ID:" & RampBalls(1, 1) & " type:" & RampType(1) & " Loops:" & RampBalls(1, 2) & vbNewLine & _
  "2 " & TypeName(RampBalls(2, 0)) & " ID:" & RampBalls(2, 1) & " type:" & RampType(2) & " Loops:" & RampBalls(2, 2) & vbNewLine & _
  "3 " & TypeName(RampBalls(3, 0)) & " ID:" & RampBalls(3, 1) & " type:" & RampType(3) & " Loops:" & RampBalls(3, 2) & vbNewLine & _
  "4 " & TypeName(RampBalls(4, 0)) & " ID:" & RampBalls(4, 1) & " type:" & RampType(4) & " Loops:" & RampBalls(4, 2) & vbNewLine & _
  "5 " & TypeName(RampBalls(5, 0)) & " ID:" & RampBalls(5, 1) & " type:" & RampType(5) & " Loops:" & RampBalls(5, 2) & vbNewLine & _
  "6 " & TypeName(RampBalls(6, 0)) & " ID:" & RampBalls(6, 1) & " type:" & RampType(6) & " Loops:" & RampBalls(6, 2) & vbNewLine & _
  " "
End Sub

Function BallPitch(ball) ' Calculates the pitch of the sound based on the ball speed
  BallPitch = pSlope(BallVel(ball), 1, - 1000, 60, 10000)
End Function

Function BallPitchV(ball) ' Calculates the pitch of the sound based on the ball speed Variation
  BallPitchV = pSlope(BallVel(ball), 1, - 4000, 60, 7000)
End Function

Sub RandomSoundRampStop(obj)
  Select Case Int(rnd*3)
    Case 0: PlaySoundAtVol "wireramp_stop1", obj, 0.2*VolumeDial:PlaySoundAtLevelActiveBall ("Rubber_Strong_1"), Vol(ActiveBall) * RubberStrongSoundFactor * 0.6
    Case 1: PlaySoundAtVol "wireramp_stop2", obj, 0.2*VolumeDial:PlaySoundAtLevelActiveBall ("Rubber_Strong_2"), Vol(ActiveBall) * RubberStrongSoundFactor * 0.6
    Case 2: PlaySoundAtVol "wireramp_stop3", obj, 0.2*VolumeDial:PlaySoundAtLevelActiveBall ("Rubber_1_Hard"), Vol(ActiveBall) * RubberStrongSoundFactor * 0.6
  End Select
End Sub


' Ramp triggers
Sub RampTrigger001_hit
  WireRampOn true
End Sub
Sub RampTrigger001_unhit
  If activeball.vely > 0 then
    WireRampOff
  Else 'ramp helper
    activeball.vely = activeball.vely * 1.1
    activeball.velx = activeball.velx * 1.1
  End If
End Sub

Sub RampTrigger002_hit
  WireRampOn true
End Sub
Sub RampTrigger002_unhit
  if activeball.vely > 0 then WireRampOff
End Sub

Sub RampTrigger003_hit
  WireRampOn true
End Sub
Sub RampTrigger003_unhit
  if activeball.vely > 0 then WireRampOff
End Sub

Sub RampTrigger004_hit
  WireRampOff
  WireRampOn False
End Sub

Sub RampTrigger005_hit
  WireRampOff
End Sub

Sub RampTrigger006_hit
  WireRampOff
End Sub

Sub RampTrigger007_hit
  WireRampOff
End Sub

Sub RampTrigger008_hit
  WireRampOff
End Sub

Sub RampTrigger009_hit
  WireRampOn True
End Sub






'***************************************************************
' ZSHA: Ambient ball shadows
'***************************************************************

' For dynamic ball shadows, Check the "Raytraced ball shadows" box for the specific light.
' Also make sure the light's z position is around 25 (mid ball)

'Ambient (Room light source)
Const AmbientBSFactor = 0.9    '0 To 1, higher is darker
Const AmbientMovement = 1    '1+ higher means more movement as the ball moves left and right
Const offsetX = 0        'Offset x position under ball (These are if you want to change where the "room" light is for calculating the shadow position,)
Const offsetY = 0        'Offset y position under ball (^^for example 5,5 if the light is in the back left corner)

' *** Trim or extend these to match the number of balls/primitives/flashers on the table!  (will throw errors if there aren't enough objects)
Dim objBallShadow(7)

'Initialization
BSInit

Sub BSInit()
  Dim iii
  'Prepare the shadow objects before play begins
  For iii = 0 To tnob - 1
    Set objBallShadow(iii) = Eval("BallShadow" & iii)
    objBallShadow(iii).material = "BallShadow" & iii
    UpdateMaterial objBallShadow(iii).material,1,0,0,0,0,0,AmbientBSFactor,RGB(0,0,0),0,0,False,True,0,0,0,0
    objBallShadow(iii).Z = 3 + iii / 1000
    objBallShadow(iii).visible = 0
  Next
End Sub


Sub BSUpdate
  dim gBOT: gBOT = getballs
  Dim s: For s = lob To UBound(gBOT)
    ' *** Normal "ambient light" ball shadow

    'Primitive shadow on playfield, flasher shadow in ramps
    '** If on main and upper pf
    If gBOT(s).Z > 20 And gBOT(s).Z < 30 Then
      objBallShadow(s).visible = 1
      objBallShadow(s).X = gBOT(s).X + (gBOT(s).X - (tablewidth / 2)) / (Ballsize / AmbientMovement) + offsetX
      objBallShadow(s).Y = gBOT(s).Y + offsetY
      'objBallShadow(s).Z = gBOT(s).Z + s/1000 + 1.04 - 25

    '** No shadow if ball is off the main playfield (this may need to be adjusted per table)
    Else
      objBallShadow(s).visible = 0
    End If
  Next
End Sub



'**********************************
'   ZMAT: General Math Functions
'**********************************
' These get used throughout the script.

Dim PI
PI = 4 * Atn(1)

Function dSin(degrees)
  dsin = Sin(degrees * Pi / 180)
End Function

Function dCos(degrees)
  dcos = Cos(degrees * Pi / 180)
End Function

Function Atn2(dy, dx)
  If dx > 0 Then
    Atn2 = Atn(dy / dx)
  ElseIf dx < 0 Then
    If dy = 0 Then
      Atn2 = pi
    Else
      Atn2 = Sgn(dy) * (pi - Atn(Abs(dy / dx)))
    End If
  ElseIf dx = 0 Then
    If dy = 0 Then
      Atn2 = 0
    Else
      Atn2 = Sgn(dy) * pi / 2
    End If
  End If
End Function

Function ArcCos(x)
  If x = 1 Then
    ArcCos = 0/180*PI
  ElseIf x = -1 Then
    ArcCos = 180/180*PI
  Else
    ArcCos = Atn(-x/Sqr(-x * x + 1)) + 2 * Atn(1)
  End If
End Function

Function max(a,b)
  If a > b Then
    max = a
  Else
    max = b
  End If
End Function

Function min(a,b)
  If a > b Then
    min = b
  Else
    min = a
  End If
End Function

' Used for drop targets
Function InRect(px,py,ax,ay,bx,by,cx,cy,dx,dy) 'Determines if a Points (px,py) is inside a 4 point polygon A-D in Clockwise/CCW order
  Dim AB, BC, CD, DA
  AB = (bx * py) - (by * px) - (ax * py) + (ay * px) + (ax * by) - (ay * bx)
  BC = (cx * py) - (cy * px) - (bx * py) + (by * px) + (bx * cy) - (by * cx)
  CD = (dx * py) - (dy * px) - (cx * py) + (cy * px) + (cx * dy) - (cy * dx)
  DA = (ax * py) - (ay * px) - (dx * py) + (dy * px) + (dx * ay) - (dy * ax)

  If (AB <= 0 And BC <= 0 And CD <= 0 And DA <= 0) Or (AB >= 0 And BC >= 0 And CD >= 0 And DA >= 0) Then
    InRect = True
  Else
    InRect = False
  End If
End Function

Function InRotRect(ballx,bally,px,py,angle,ax,ay,bx,by,cx,cy,dx,dy)
  Dim rax,ray,rbx,rby,rcx,rcy,rdx,rdy
  Dim rotxy
  rotxy = RotPoint(ax,ay,angle)
  rax = rotxy(0) + px
  ray = rotxy(1) + py
  rotxy = RotPoint(bx,by,angle)
  rbx = rotxy(0) + px
  rby = rotxy(1) + py
  rotxy = RotPoint(cx,cy,angle)
  rcx = rotxy(0) + px
  rcy = rotxy(1) + py
  rotxy = RotPoint(dx,dy,angle)
  rdx = rotxy(0) + px
  rdy = rotxy(1) + py

  InRotRect = InRect(ballx,bally,rax,ray,rbx,rby,rcx,rcy,rdx,rdy)
End Function

Function RotPoint(x,y,angle)
  Dim rx, ry
  rx = x * dCos(angle) - y * dSin(angle)
  ry = x * dSin(angle) + y * dCos(angle)
  RotPoint = Array(rx,ry)
End Function



'******************************************************
' ZNFF:  FLIPPER CORRECTIONS by nFozzy
'******************************************************



Sub LeftFlipper_Collide(parm)
  CheckLiveCatch ActiveBall, LeftFlipper, LFCount, parm
  LF.ReProcessBalls ActiveBall
  LeftFlipperCollide parm
End Sub

Sub RightFlipper_Collide(parm)
  CheckLiveCatch ActiveBall, RightFlipper, RFCount, parm
  RF.ReProcessBalls ActiveBall
  RightFlipperCollide parm
End Sub




dim LF : Set LF = New FlipperPolarity
dim RF : Set RF = New FlipperPolarity

InitPolarity

Sub InitPolarity()
  Dim x, a
  a = Array(LF, RF)
  For Each x In a
    x.AddPt "Ycoef", 0, RightFlipper.Y-65, 1 'disabled
    x.AddPt "Ycoef", 1, RightFlipper.Y-11, 1
    x.enabled = True
    x.TimeDelay = 60
    x.DebugOn=False ' prints some info in debugger

    x.AddPt "Polarity", 0, 0, 0
    x.AddPt "Polarity", 1, 0.05, - 5.5
    x.AddPt "Polarity", 2, 0.16, - 5.5
    x.AddPt "Polarity", 3, 0.20, - 0.75
    x.AddPt "Polarity", 4, 0.25, - 1.25
    x.AddPt "Polarity", 5, 0.3, - 1.75
    x.AddPt "Polarity", 6, 0.4, - 3.5
    x.AddPt "Polarity", 7, 0.5, - 5.25
    x.AddPt "Polarity", 8, 0.7, - 4.0
    x.AddPt "Polarity", 9, 0.75, - 3.5
    x.AddPt "Polarity", 10, 0.8, - 3.0
    x.AddPt "Polarity", 11, 0.85, - 2.5
    x.AddPt "Polarity", 12, 0.9, - 2.0
    x.AddPt "Polarity", 13, 0.95, - 1.5
    x.AddPt "Polarity", 14, 1, - 1.0
    x.AddPt "Polarity", 15, 1.05, -0.5
    x.AddPt "Polarity", 16, 1.1, 0
    x.AddPt "Polarity", 17, 1.3, 0

    x.AddPt "Velocity", 0, 0, 0.85
    x.AddPt "Velocity", 1, 0.23, 0.85
    x.AddPt "Velocity", 2, 0.35, 0.88
    x.AddPt "Velocity", 3, 0.45, 1
    x.AddPt "Velocity", 4, 0.6, 1 '0.982
    x.AddPt "Velocity", 5, 0.62, 1.0
    x.AddPt "Velocity", 6, 0.702, 0.968
    x.AddPt "Velocity", 7, 0.95,  0.968
    x.AddPt "Velocity", 8, 1.03,  0.945
    x.AddPt "Velocity", 9, 1.5,  0.945

  Next

  ' SetObjects arguments: 1: name of object 2: flipper object: 3: Trigger object around flipper
  LF.SetObjects "LF", LeftFlipper, TriggerLF
  RF.SetObjects "RF", RightFlipper, TriggerRF
End Sub


'******************************************************
'  FLIPPER CORRECTION FUNCTIONS
'******************************************************

' modified 2023 by nFozzy
' Removed need for 'endpoint' objects
' Added 'createvents' type thing for TriggerLF / TriggerRF triggers.
' Removed AddPt function which complicated setup imo
' made DebugOn do something (prints some stuff in debugger)
'   Otherwise it should function exactly the same as before\
' modified 2024 by rothbauerw
' Added Reprocessballs for flipper collisions (LF.Reprocessballs Activeball and RF.Reprocessballs Activeball must be added to the flipper collide subs
' Improved handling to remove correction for backhand shots when the flipper is raised

Class FlipperPolarity
  Public DebugOn, Enabled
  Private FlipAt    'Timer variable (IE 'flip at 723,530ms...)
  Public TimeDelay    'delay before trigger turns off and polarity is disabled
  Private Flipper, FlipperStart, FlipperEnd, FlipperEndY, LR, PartialFlipCoef, FlipStartAngle
  Private Balls(20), balldata(20)
  Private Name

  Dim PolarityIn, PolarityOut
  Dim VelocityIn, VelocityOut
  Dim YcoefIn, YcoefOut
  Public Sub Class_Initialize
    ReDim PolarityIn(0)
    ReDim PolarityOut(0)
    ReDim VelocityIn(0)
    ReDim VelocityOut(0)
    ReDim YcoefIn(0)
    ReDim YcoefOut(0)
    Enabled = True
    TimeDelay = 50
    LR = 1
    Dim x
    For x = 0 To UBound(balls)
      balls(x) = Empty
      Set Balldata(x) = new SpoofBall
    Next
  End Sub

  Public Sub SetObjects(aName, aFlipper, aTrigger)

    If TypeName(aName) <> "String" Then MsgBox "FlipperPolarity: .SetObjects error: first argument must be a String (And name of Object). Found:" & TypeName(aName) End If
    If TypeName(aFlipper) <> "Flipper" Then MsgBox "FlipperPolarity: .SetObjects error: Second argument must be a flipper. Found:" & TypeName(aFlipper) End If
    If TypeName(aTrigger) <> "Trigger" Then MsgBox "FlipperPolarity: .SetObjects error: third argument must be a trigger. Found:" & TypeName(aTrigger) End If
    If aFlipper.EndAngle > aFlipper.StartAngle Then LR = -1 Else LR = 1 End If
    Name = aName
    Set Flipper = aFlipper
    FlipperStart = aFlipper.x
    FlipperEnd = Flipper.Length * Sin((Flipper.StartAngle / 57.295779513082320876798154814105)) + Flipper.X ' big floats for degree to rad conversion
    FlipperEndY = Flipper.Length * Cos(Flipper.StartAngle / 57.295779513082320876798154814105)*-1 + Flipper.Y

    Dim str
    str = "Sub " & aTrigger.name & "_Hit() : " & aName & ".AddBall ActiveBall : End Sub'"
    ExecuteGlobal(str)
    str = "Sub " & aTrigger.name & "_UnHit() : " & aName & ".PolarityCorrect ActiveBall : End Sub'"
    ExecuteGlobal(str)

  End Sub

  ' Legacy: just no op
  Public Property Let EndPoint(aInput)

  End Property

  Public Sub AddPt(aChooseArray, aIDX, aX, aY) 'Index #, X position, (in) y Position (out)
    Select Case aChooseArray
      Case "Polarity"
        ShuffleArrays PolarityIn, PolarityOut, 1
        PolarityIn(aIDX) = aX
        PolarityOut(aIDX) = aY
        ShuffleArrays PolarityIn, PolarityOut, 0
      Case "Velocity"
        ShuffleArrays VelocityIn, VelocityOut, 1
        VelocityIn(aIDX) = aX
        VelocityOut(aIDX) = aY
        ShuffleArrays VelocityIn, VelocityOut, 0
      Case "Ycoef"
        ShuffleArrays YcoefIn, YcoefOut, 1
        YcoefIn(aIDX) = aX
        YcoefOut(aIDX) = aY
        ShuffleArrays YcoefIn, YcoefOut, 0
    End Select
  End Sub

  Public Sub AddBall(aBall)
    Dim x
    For x = 0 To UBound(balls)
      If IsEmpty(balls(x)) Then
        Set balls(x) = aBall
        Exit Sub
      End If
    Next
  End Sub

  Private Sub RemoveBall(aBall)
    Dim x
    For x = 0 To UBound(balls)
      If TypeName(balls(x) ) = "IBall" Then
        If aBall.ID = Balls(x).ID Then
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
    Dim x
    For x = 0 To UBound(balls)
      If Not IsEmpty(balls(x)) Then
        pos = pSlope(Balls(x).x, FlipperStart, 0, FlipperEnd, 1)
      End If
    Next
  End Property

  Public Sub ProcessBalls() 'save data of balls in flipper range
    FlipAt = GameTime
    Dim x
    For x = 0 To UBound(balls)
      If Not IsEmpty(balls(x)) Then
        balldata(x).Data = balls(x)
      End If
    Next
    FlipStartAngle = Flipper.currentangle
    PartialFlipCoef = ((Flipper.StartAngle - Flipper.CurrentAngle) / (Flipper.StartAngle - Flipper.EndAngle))
    PartialFlipCoef = abs(PartialFlipCoef-1)
  End Sub

  Public Sub ReProcessBalls(aBall) 'save data of balls in flipper range
    If FlipperOn() Then
      Dim x
      For x = 0 To UBound(balls)
        If Not IsEmpty(balls(x)) Then
          if balls(x).ID = aBall.ID Then
            If isempty(balldata(x).ID) Then
              balldata(x).Data = balls(x)
            End If
          End If
        End If
      Next
    End If
  End Sub

  'Timer shutoff for polaritycorrect
  Private Function FlipperOn()
    If GameTime < FlipAt+TimeDelay Then
      FlipperOn = True
    End If
  End Function

  Public Sub PolarityCorrect(aBall)
    If FlipperOn() Then
      Dim tmp, BallPos, x, IDX, Ycoef, BalltoFlip, BalltoBase, NoCorrection, checkHit
      Ycoef = 1

      'y safety Exit
      If aBall.VelY > -8 Then 'ball going down
        RemoveBall aBall
        Exit Sub
      End If

      'Find balldata. BallPos = % on Flipper
      For x = 0 To UBound(Balls)
        If aBall.id = BallData(x).id And Not IsEmpty(BallData(x).id) Then
          idx = x
          BallPos = PSlope(BallData(x).x, FlipperStart, 0, FlipperEnd, 1)
          BalltoFlip = DistanceFromFlipperAngle(BallData(x).x, BallData(x).y, Flipper, FlipStartAngle)
          If ballpos > 0.65 Then  Ycoef = LinearEnvelope(BallData(x).Y, YcoefIn, YcoefOut)                'find safety coefficient 'ycoef' data
        End If
      Next

      If BallPos = 0 Then 'no ball data meaning the ball is entering and exiting pretty close to the same position, use current values.
        BallPos = PSlope(aBall.x, FlipperStart, 0, FlipperEnd, 1)
        If ballpos > 0.65 Then  Ycoef = LinearEnvelope(aBall.Y, YcoefIn, YcoefOut)                        'find safety coefficient 'ycoef' data
        NoCorrection = 1
      Else
        checkHit = 50 + (20 * BallPos)

        If BalltoFlip > checkHit or (PartialFlipCoef < 0.5 and BallPos > 0.22) Then
          NoCorrection = 1
        Else
          NoCorrection = 0
        End If
      End If

      'Velocity correction
      If Not IsEmpty(VelocityIn(0) ) Then
        Dim VelCoef
        VelCoef = LinearEnvelope(BallPos, VelocityIn, VelocityOut)

        'If partialflipcoef < 1 Then VelCoef = PSlope(partialflipcoef, 0, 1, 1, VelCoef)

        If Enabled Then aBall.Velx = aBall.Velx*VelCoef
        If Enabled Then aBall.Vely = aBall.Vely*VelCoef
      End If

      'Polarity Correction (optional now)
      If Not IsEmpty(PolarityIn(0) ) Then
        Dim AddX
        AddX = LinearEnvelope(BallPos, PolarityIn, PolarityOut) * LR

        If Enabled and NoCorrection = 0 Then aBall.VelX = aBall.VelX + 1 * (AddX*ycoef*PartialFlipcoef*VelCoef)
      End If
      If DebugOn Then debug.print "PolarityCorrect" & " " & Name & " @ " & GameTime & " " & Round(BallPos*100) & "%" & " AddX:" & Round(AddX,2) & " Vel%:" & Round(VelCoef*100)
    End If
    RemoveBall aBall
  End Sub
End Class

'******************************************************
'  FLIPPER POLARITY AND RUBBER DAMPENER SUPPORTING FUNCTIONS
'******************************************************

' Used for flipper correction and rubber dampeners
Sub ShuffleArray(ByRef aArray, byVal offset) 'shuffle 1d array
  Dim x, aCount
  aCount = 0
  ReDim a(UBound(aArray) )
  For x = 0 To UBound(aArray)   'Shuffle objects in a temp array
    If Not IsEmpty(aArray(x) ) Then
      If IsObject(aArray(x)) Then
        Set a(aCount) = aArray(x)
      Else
        a(aCount) = aArray(x)
      End If
      aCount = aCount + 1
    End If
  Next
  If offset < 0 Then offset = 0
  ReDim aArray(aCount-1+offset)   'Resize original array
  For x = 0 To aCount-1       'set objects back into original array
    If IsObject(a(x)) Then
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
  BallSpeed = Sqr(ball.VelX^2 + ball.VelY^2 + ball.VelZ^2)
End Function

' Used for flipper correction and rubber dampeners
Function PSlope(Input, X1, Y1, X2, Y2)    'Set up line via two points, no clamping. Input X, output Y
  Dim x, y, b, m
  x = input
  m = (Y2 - Y1) / (X2 - X1)
  b = Y2 - m*X2
  Y = M*x+b
  PSlope = Y
End Function

' Used for flipper correction
Class spoofball
  Public X, Y, Z, VelX, VelY, VelZ, ID, Mass, Radius
  Public Property Let Data(aBall)
    With aBall
      x = .x
      y = .y
      z = .z
      velx = .velx
      vely = .vely
      velz = .velz
      id = .ID
      mass = .mass
      radius = .radius
    End With
  End Property
  Public Sub Reset()
    x = Empty
    y = Empty
    z = Empty
    velx = Empty
    vely = Empty
    velz = Empty
    id = Empty
    mass = Empty
    radius = Empty
  End Sub
End Class

' Used for flipper correction and rubber dampeners
Function LinearEnvelope(xInput, xKeyFrame, yLvl)
  Dim y 'Y output
  Dim L 'Line
  'find active line
  Dim ii
  For ii = 1 To UBound(xKeyFrame)
    If xInput <= xKeyFrame(ii) Then
      L = ii
      Exit For
    End If
  Next
  If xInput > xKeyFrame(UBound(xKeyFrame) ) Then L = UBound(xKeyFrame)    'catch line overrun
  Y = pSlope(xInput, xKeyFrame(L-1), yLvl(L-1), xKeyFrame(L), yLvl(L) )

  If xInput <= xKeyFrame(LBound(xKeyFrame) ) Then Y = yLvl(LBound(xKeyFrame) )     'Clamp lower
  If xInput >= xKeyFrame(UBound(xKeyFrame) ) Then Y = yLvl(UBound(xKeyFrame) )    'Clamp upper

  LinearEnvelope = Y
End Function

'******************************************************
'  FLIPPER TRICKS
'******************************************************
' To add the flipper tricks you must
'  - Include a call to FlipperCradleCollision from within OnBallBallCollision subroutine
'  - Include a call the CheckLiveCatch from the LeftFlipper_Collide and RightFlipper_Collide subroutines
'  - Include FlipperActivate and FlipperDeactivate in the Flipper solenoid subs

RightFlipper.timerinterval = 1
Rightflipper.timerenabled = True

Sub RightFlipper_timer()
  FlipperTricks LeftFlipper, LFPress, LFCount, LFEndAngle, LFState
  FlipperTricks RightFlipper, RFPress, RFCount, RFEndAngle, RFState
  FlipperNudge RightFlipper, RFEndAngle, RFEOSNudge, LeftFlipper, LFEndAngle
  FlipperNudge LeftFlipper, LFEndAngle, LFEOSNudge,  RightFlipper, RFEndAngle
End Sub

Dim LFEOSNudge, RFEOSNudge

Sub FlipperNudge(Flipper1, Endangle1, EOSNudge1, Flipper2, EndAngle2)
  Dim b
  Dim gBOT
  gBOT = GetBalls

  If Flipper1.currentangle = Endangle1 And EOSNudge1 <> 1 Then
    EOSNudge1 = 1
    '   debug.print Flipper1.currentangle &" = "& Endangle1 &"--"& Flipper2.currentangle &" = "& EndAngle2
    If Flipper2.currentangle = EndAngle2 Then
      For b = 0 To UBound(gBOT)
        If FlipperTrigger(gBOT(b).x, gBOT(b).y, Flipper1) Then
          'Debug.Print "ball in flip1. exit"
          Exit Sub
        End If
      Next
      For b = 0 To UBound(gBOT)
        If FlipperTrigger(gBOT(b).x, gBOT(b).y, Flipper2) Then
          gBOT(b).velx = gBOT(b).velx / 1.3
          gBOT(b).vely = gBOT(b).vely - 0.5
        End If
      Next
    End If
  Else
    If Abs(Flipper1.currentangle) > Abs(EndAngle1) + 30 Then EOSNudge1 = 0
  End If
End Sub


Dim FCCDamping: FCCDamping = 0.4

Sub FlipperCradleCollision(ball1, ball2, velocity)
  if velocity < 0.7 then exit sub   'filter out gentle collisions
    Dim DoDamping, coef
    DoDamping = false
    'Check left flipper
    If LeftFlipper.currentangle = LFEndAngle Then
    If FlipperTrigger(ball1.x, ball1.y, LeftFlipper) OR FlipperTrigger(ball2.x, ball2.y, LeftFlipper) Then DoDamping = true
    End If
    'Check right flipper
    If RightFlipper.currentangle = RFEndAngle Then
    If FlipperTrigger(ball1.x, ball1.y, RightFlipper) OR FlipperTrigger(ball2.x, ball2.y, RightFlipper) Then DoDamping = true
    End If
    If DoDamping Then
    coef = FCCDamping
        ball1.velx = ball1.velx * coef: ball1.vely = ball1.vely * coef: ball1.velz = ball1.velz * coef
        ball2.velx = ball2.velx * coef: ball2.vely = ball2.vely * coef: ball2.velz = ball2.velz * coef
    End If
End Sub





'*************************************************
'  Check ball distance from Flipper for Rem
'*************************************************

Function Distance(ax,ay,bx,by)
  Distance = Sqr((ax - bx) ^ 2 + (ay - by) ^ 2)
End Function

Function DistancePL(px,py,ax,ay,bx,by) 'Distance between a point and a line where point Is px,py
  DistancePL = Abs((by - ay) * px - (bx - ax) * py + bx * ay - by * ax) / Distance(ax,ay,bx,by)
End Function

Function Radians(Degrees)
  Radians = Degrees * PI / 180
End Function

Function AnglePP(ax,ay,bx,by)
  AnglePP = Atn2((by - ay),(bx - ax)) * 180 / PI
End Function

Function DistanceFromFlipper(ballx, bally, Flipper)
  DistanceFromFlipper = DistancePL(ballx, bally, Flipper.x, Flipper.y, Cos(Radians(Flipper.currentangle + 90)) + Flipper.x, Sin(Radians(Flipper.currentangle + 90)) + Flipper.y)
End Function

Function DistanceFromFlipperAngle(ballx, bally, Flipper, Angle)
  DistanceFromFlipperAngle = DistancePL(ballx, bally, Flipper.x, Flipper.y, Cos(Radians(Angle + 90)) + Flipper.x, Sin(Radians(angle + 90)) + Flipper.y)
End Function

Function FlipperTrigger(ballx, bally, Flipper)
  Dim DiffAngle
  DiffAngle = Abs(Flipper.currentangle - AnglePP(Flipper.x, Flipper.y, ballx, bally) - 90)
  If DiffAngle > 180 Then DiffAngle = DiffAngle - 360

  If DistanceFromFlipper(ballx,bally,Flipper) < 48 And DiffAngle <= 90 And Distance(ballx,bally,Flipper.x,Flipper.y) < Flipper.Length Then
    FlipperTrigger = True
  Else
    FlipperTrigger = False
  End If
End Function

'*************************************************
'  End - Check ball distance from Flipper for Rem
'*************************************************

Dim LFPress, RFPress, LFCount, RFCount
Dim LFState, RFState
Dim EOST, EOSA,Frampup, FElasticity,FReturn
Dim RFEndAngle, LFEndAngle

Const FlipperCoilRampupMode = 0 '0 = fast, 1 = medium, 2 = slow (tap passes should work)

LFState = 1
RFState = 1
EOST = leftflipper.eostorque
EOSA = leftflipper.eostorqueangle
Frampup = LeftFlipper.rampup
FElasticity = LeftFlipper.elasticity
FReturn = LeftFlipper.return
'Const EOSTnew = 1.5 'EM's to late 80's - new recommendation by rothbauerw (previously 1)
Const EOSTnew = 1.2 '90's and later - new recommendation by rothbauerw (previously 0.8)
Const EOSAnew = 1
Const EOSRampup = 0
Dim SOSRampup
Select Case FlipperCoilRampupMode
  Case 0
    SOSRampup = 2.5
  Case 1
    SOSRampup = 6
  Case 2
    SOSRampup = 8.5
End Select

Const LiveCatch = 16
Const LiveElasticity = 0.45
Const SOSEM = 0.815
'Const EOSReturn = 0.055  'EM's
'Const EOSReturn = 0.045  'late 70's to mid 80's
'Const EOSReturn = 0.035  'mid 80's to early 90's
Const EOSReturn = 0.025  'mid 90's and later

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
  Flipper.eostorque = EOST * EOSReturn / FReturn

  If Abs(Flipper.currentangle) <= Abs(Flipper.endangle) + 0.1 Then
    Dim b, gBOT
    gBOT = GetBalls

    For b = 0 To UBound(gBOT)
      If Distance(gBOT(b).x, gBOT(b).y, Flipper.x, Flipper.y) < 55 Then 'check for cradle
        If gBOT(b).vely >= - 0.4 Then gBOT(b).vely =  - 0.4
      End If
    Next
  End If
End Sub

Sub FlipperTricks (Flipper, FlipperPress, FCount, FEndAngle, FState)
  Dim Dir
  Dir = Flipper.startangle / Abs(Flipper.startangle) '-1 for Right Flipper

  If Abs(Flipper.currentangle) > Abs(Flipper.startangle) - 0.05 Then
    If FState <> 1 Then
      Flipper.rampup = SOSRampup
      Flipper.endangle = FEndAngle - 3 * Dir
      Flipper.Elasticity = FElasticity * SOSEM
      FCount = 0
      FState = 1
    End If
  ElseIf Abs(Flipper.currentangle) <= Abs(Flipper.endangle) And FlipperPress = 1 Then
    If FCount = 0 Then FCount = GameTime

    If FState <> 2 Then
      Flipper.eostorqueangle = EOSAnew
      Flipper.eostorque = EOSTnew
      Flipper.rampup = EOSRampup
      Flipper.endangle = FEndAngle
      FState = 2
    End If
  ElseIf Abs(Flipper.currentangle) > Abs(Flipper.endangle) + 0.01 And FlipperPress = 1 Then
    If FState <> 3 Then
      Flipper.eostorque = EOST
      Flipper.eostorqueangle = EOSA
      Flipper.rampup = Frampup
      Flipper.Elasticity = FElasticity
      FState = 3
    End If
  End If
End Sub

Const LiveDistanceMin = 5  'minimum distance In vp units from flipper base live catch dampening will occur
Const LiveDistanceMax = 114 'maximum distance in vp units from flipper base live catch dampening will occur (tip protection)
Const BaseDampen = 0.55

Sub CheckLiveCatch(ball, Flipper, FCount, parm) 'Experimental new live catch
    Dim Dir, LiveDist
    Dir = Flipper.startangle / Abs(Flipper.startangle)    '-1 for Right Flipper
    Dim LiveCatchBounce   'If live catch is not perfect, it won't freeze ball totally
    Dim CatchTime
    CatchTime = GameTime - FCount
    LiveDist = Abs(Flipper.x - ball.x)

    If CatchTime <= LiveCatch And parm > 3 And LiveDist > LiveDistanceMin And LiveDist < LiveDistanceMax Then
        If CatchTime <= LiveCatch * 0.5 Then   'Perfect catch only when catch time happens in the beginning of the window
            LiveCatchBounce = 0
        Else
            LiveCatchBounce = Abs((LiveCatch / 2) - CatchTime)  'Partial catch when catch happens a bit late
        End If

        If LiveCatchBounce = 0 And ball.velx * Dir > 0 And LiveDist > 30 Then ball.velx = 0

        If ball.velx * Dir > 0 And LiveDist < 30 Then
            ball.velx = BaseDampen * ball.velx
            ball.vely = BaseDampen * ball.vely
            ball.angmomx = BaseDampen * ball.angmomx
            ball.angmomy = BaseDampen * ball.angmomy
            ball.angmomz = BaseDampen * ball.angmomz
        Elseif LiveDist > 30 Then
            ball.vely = LiveCatchBounce * (32 / LiveCatch) ' Multiplier for inaccuracy bounce
            ball.angmomx = 0
            ball.angmomy = 0
            ball.angmomz = 0
        End If
    Else
        If Abs(Flipper.currentangle) <= Abs(Flipper.endangle) + 1 Then FlippersD.Dampenf ActiveBall, parm
    End If
End Sub

'******************************************************
'****  END FLIPPER CORRECTIONS
'******************************************************





'******************************************************
'   ZDMP:  RUBBER  DAMPENERS
'******************************************************
' These are data mined bounce curves,
' dialed in with the in-game elasticity as much as possible to prevent angle / spin issues.
' Requires tracking ballspeed to calculate COR

Sub zCol_Rubber_L_hit: RubbersD.dampen ActiveBall: End Sub
Sub zCol_Rubber_R_hit: RubbersD.dampen ActiveBall: End Sub

Sub dPosts_Hit(idx)
  RubbersD.dampen ActiveBall
  TargetBouncer ActiveBall, 1
End Sub

Sub dSleeves_Hit(idx)
  SleevesD.Dampen ActiveBall
  TargetBouncer ActiveBall, 0.7
End Sub

Dim RubbersD        'frubber
Set RubbersD = New Dampener
RubbersD.name = "Rubbers"
RubbersD.debugOn = False  'shows info in textbox "TBPout"
RubbersD.Print = False    'debug, reports In debugger (In vel, out cor); cor bounce curve (linear)

'for best results, try to match in-game velocity as closely as possible to the desired curve
'   RubbersD.addpoint 0, 0, 0.935   'point# (keep sequential), ballspeed, CoR (elasticity)
RubbersD.addpoint 0, 0, 1.1    'point# (keep sequential), ballspeed, CoR (elasticity)
RubbersD.addpoint 1, 3.77, 0.97
RubbersD.addpoint 2, 5.76, 0.967  'dont take this as gospel. if you can data mine rubber elasticitiy, please help!
RubbersD.addpoint 3, 15.84, 0.874
RubbersD.addpoint 4, 56, 0.64    'there's clamping so interpolate up to 56 at least

Dim SleevesD  'this is just rubber but cut down to 85%...
Set SleevesD = New Dampener
SleevesD.name = "Sleeves"
SleevesD.debugOn = False  'shows info in textbox "TBPout"
SleevesD.Print = False    'debug, reports In debugger (In vel, out cor)
SleevesD.CopyCoef RubbersD, 0.85

'######################### Add new FlippersD Profile
'######################### Adjust these values to increase or lessen the elasticity

Dim FlippersD
Set FlippersD = New Dampener
FlippersD.name = "Flippers"
FlippersD.debugOn = False
FlippersD.Print = False
FlippersD.addpoint 0, 0, 1.1
FlippersD.addpoint 1, 3.77, 0.99
FlippersD.addpoint 2, 6, 0.99

Class Dampener
  Public Print, debugOn   'tbpOut.text
  Public name, Threshold  'Minimum threshold. Useful for Flippers, which don't have a hit threshold.
  Public ModIn, ModOut
  Private Sub Class_Initialize
    ReDim ModIn(0)
    ReDim Modout(0)
  End Sub

  Public Sub AddPoint(aIdx, aX, aY)
    ShuffleArrays ModIn, ModOut, 1
    ModIn(aIDX) = aX
    ModOut(aIDX) = aY
    ShuffleArrays ModIn, ModOut, 0
    If GameTime > 100 Then Report
  End Sub

  Public Sub Dampen(aBall)
    If threshold Then
      If BallSpeed(aBall) < threshold Then Exit Sub
    End If
    Dim RealCOR, DesiredCOR, str, coef
    DesiredCor = LinearEnvelope(cor.ballvel(aBall.id), ModIn, ModOut )
    RealCOR = BallSpeed(aBall) / (cor.ballvel(aBall.id) + 0.0001)
    coef = desiredcor / realcor
    If debugOn Then str = name & " In vel:" & Round(cor.ballvel(aBall.id),2 ) & vbNewLine & "desired cor: " & Round(desiredcor,4) & vbNewLine & _
    "actual cor: " & Round(realCOR,4) & vbNewLine & "ballspeed coef: " & Round(coef, 3) & vbNewLine
    If Print Then Debug.print Round(cor.ballvel(aBall.id),2) & ", " & Round(desiredcor,3)

    aBall.velx = aBall.velx * coef
    aBall.vely = aBall.vely * coef
    aBall.velz = aBall.velz * coef
    If debugOn Then TBPout.text = str
  End Sub

  Public Sub Dampenf(aBall, parm) 'Rubberizer is handle here
    Dim RealCOR, DesiredCOR, str, coef
    DesiredCor = LinearEnvelope(cor.ballvel(aBall.id), ModIn, ModOut )
    RealCOR = BallSpeed(aBall) / (cor.ballvel(aBall.id) + 0.0001)
    coef = desiredcor / realcor
    If Abs(aball.velx) < 2 And aball.vely < 0 And aball.vely >  - 3.75 Then
      aBall.velx = aBall.velx * coef
      aBall.vely = aBall.vely * coef
      aBall.velz = aBall.velz * coef
    End If
  End Sub

  Public Sub CopyCoef(aObj, aCoef) 'alternative addpoints, copy with coef
    Dim x
    For x = 0 To UBound(aObj.ModIn)
      addpoint x, aObj.ModIn(x), aObj.ModOut(x) * aCoef
    Next
  End Sub

  Public Sub Report() 'debug, reports all coords in tbPL.text
    If Not debugOn Then Exit Sub
    Dim a1, a2
    a1 = ModIn
    a2 = ModOut
    Dim str, x
    For x = 0 To UBound(a1)
      str = str & x & ": " & Round(a1(x),4) & ", " & Round(a2(x),4) & vbNewLine
    Next
    TBPout.text = str
  End Sub
End Class

'******************************************************
'  TRACK ALL BALL VELOCITIES
'  FOR RUBBER DAMPENER AND DROP TARGETS
'******************************************************

Dim cor
Set cor = New CoRTracker

Class CoRTracker
  Public ballvel, ballvelx, ballvely

  Private Sub Class_Initialize
    ReDim ballvel(0)
    ReDim ballvelx(0)
    ReDim ballvely(0)
  End Sub

  Public Sub Update() 'tracks in-ball-velocity
    Dim str, b, AllBalls, highestID
    allBalls = GetBalls

    For Each b In allballs
      If b.id >= HighestID Then highestID = b.id
    Next

    If UBound(ballvel) < highestID Then ReDim ballvel(highestID)  'set bounds
    If UBound(ballvelx) < highestID Then ReDim ballvelx(highestID)  'set bounds
    If UBound(ballvely) < highestID Then ReDim ballvely(highestID)  'set bounds

    For Each b In allballs
      ballvel(b.id) = BallSpeed(b)
      ballvelx(b.id) = b.velx
      ballvely(b.id) = b.vely
    Next
  End Sub
End Class

Sub RDampen_Timer()
  Cor.Update
End Sub

'******************************************************
'****  END PHYSICS DAMPENERS
'******************************************************


'******************************************************
'   ZBOU: VPW TargetBouncer for targets and posts by Iaakki, Wrd1972, Apophis
'******************************************************

Const TargetBouncerEnabled = 1    '0 = normal standup targets, 1 = bouncy targets
Const TargetBouncerFactor = 0.9  'Level of bounces. Recommmended value of 0.7-1

Sub TargetBouncer(aBall,defvalue)
  Dim zMultiplier, vel, vratio
  If TargetBouncerEnabled = 1 And aball.z < 30 Then
    '   debug.print "velx: " & aball.velx & " vely: " & aball.vely & " velz: " & aball.velz
    vel = BallSpeed(aBall)
    If aBall.velx = 0 Then vratio = 1 Else vratio = aBall.vely / aBall.velx
    Select Case Int(Rnd * 6) + 1
      Case 1
        zMultiplier = 0.2 * defvalue
      Case 2
        zMultiplier = 0.25 * defvalue
      Case 3
        zMultiplier = 0.3 * defvalue
      Case 4
        zMultiplier = 0.4 * defvalue
      Case 5
        zMultiplier = 0.45 * defvalue
      Case 6
        zMultiplier = 0.5 * defvalue
    End Select
    aBall.velz = Abs(vel * zMultiplier * TargetBouncerFactor)
    aBall.velx = Sgn(aBall.velx) * Sqr(Abs((vel ^ 2 - aBall.velz ^ 2) / (1 + vratio ^ 2)))
    aBall.vely = aBall.velx * vratio
    '   debug.print "---> velx: " & aball.velx & " vely: " & aball.vely & " velz: " & aball.velz
    '   debug.print "conservation check: " & BallSpeed(aBall)/vel
  End If
End Sub




'******************************************************
' ZSSC: SLINGSHOT CORRECTION FUNCTIONS by apophis
'******************************************************
' To add these slingshot corrections:
'  - On the table, add the endpoint primitives that define the two ends of the Slingshot
'  - Initialize the SlingshotCorrection objects in InitSlingCorrection
'  - Call the .VelocityCorrect methods from the respective _Slingshot event sub

dim LS : Set LS = New SlingshotCorrection
dim RS : Set RS = New SlingshotCorrection


InitSlingCorrection

Sub InitSlingCorrection
  LS.Object = LeftSlingshot
  LS.EndPoint1 = EndPoint1LS
  LS.EndPoint2 = EndPoint2LS

  RS.Object = RightSlingshot
  RS.EndPoint1 = EndPoint1RS
  RS.EndPoint2 = EndPoint2RS

  'Slingshot angle corrections (pt, BallPos in %, Angle in deg)
  ' These values are best guesses. Retune them if needed based on specific table research.
  AddSlingsPt 0, 0.00, - 4
  AddSlingsPt 1, 0.45, - 7
  AddSlingsPt 2, 0.48,  0
  AddSlingsPt 3, 0.52,  0
  AddSlingsPt 4, 0.55,  7
  AddSlingsPt 5, 1.00,  4
End Sub

Sub AddSlingsPt(idx, aX, aY)    'debugger wrapper for adjusting flipper script In-game
  Dim a
  a = Array(LS, RS)
  Dim x
  For Each x In a
    x.addpoint idx, aX, aY
  Next
End Sub

Class SlingshotCorrection
  Public DebugOn, Enabled
  Private Slingshot, SlingX1, SlingX2, SlingY1, SlingY2

  Public ModIn, ModOut

  Private Sub Class_Initialize
    ReDim ModIn(0)
    ReDim Modout(0)
    Enabled = True
  End Sub

  Public Property Let Object(aInput)
    Set Slingshot = aInput
  End Property

  Public Property Let EndPoint1(aInput)
    SlingX1 = aInput.x
    SlingY1 = aInput.y
  End Property

  Public Property Let EndPoint2(aInput)
    SlingX2 = aInput.x
    SlingY2 = aInput.y
  End Property

  Public Sub AddPoint(aIdx, aX, aY)
    ShuffleArrays ModIn, ModOut, 1
    ModIn(aIDX) = aX
    ModOut(aIDX) = aY
    ShuffleArrays ModIn, ModOut, 0
    If GameTime > 100 Then Report
  End Sub

  Public Sub Report() 'debug, reports all coords in tbPL.text
    If Not debugOn Then Exit Sub
    Dim a1, a2
    a1 = ModIn
    a2 = ModOut
    Dim str, x
    For x = 0 To UBound(a1)
      str = str & x & ": " & Round(a1(x),4) & ", " & Round(a2(x),4) & vbNewLine
    Next
    TBPout.text = str
  End Sub


  Public Sub VelocityCorrect(aBall)
    Dim BallPos, XL, XR, YL, YR

    'Assign right and left end points
    If SlingX1 < SlingX2 Then
      XL = SlingX1
      YL = SlingY1
      XR = SlingX2
      YR = SlingY2
    Else
      XL = SlingX2
      YL = SlingY2
      XR = SlingX1
      YR = SlingY1
    End If

    'Find BallPos = % on Slingshot
    If Not IsEmpty(aBall.id) Then
      If Abs(XR - XL) > Abs(YR - YL) Then
        BallPos = PSlope(aBall.x, XL, 0, XR, 1)
      Else
        BallPos = PSlope(aBall.y, YL, 0, YR, 1)
      End If
      If BallPos < 0 Then BallPos = 0
      If BallPos > 1 Then BallPos = 1
    End If

    'Velocity angle correction
    If Not IsEmpty(ModIn(0) ) Then
      Dim Angle, RotVxVy
      Angle = LinearEnvelope(BallPos, ModIn, ModOut)
      '   debug.print " BallPos=" & BallPos &" Angle=" & Angle
      '   debug.print " BEFORE: aBall.Velx=" & aBall.Velx &" aBall.Vely" & aBall.Vely
      RotVxVy = RotPoint(aBall.Velx,aBall.Vely,Angle)
      If Enabled Then aBall.Velx = RotVxVy(0)
      If Enabled Then aBall.Vely = RotVxVy(1)
      '   debug.print " AFTER: aBall.Velx=" & aBall.Velx &" aBall.Vely" & aBall.Vely
      '   debug.print " "
    End If
  End Sub
End Class




'******************************************************
'   ZFLE:  FLEEP MECHANICAL SOUNDS
'******************************************************


'///////////////////////////////  SOUNDS PARAMETERS  //////////////////////////////
Dim GlobalSoundLevel, CoinSoundLevel, PlungerReleaseSoundLevel, PlungerPullSoundLevel, NudgeLeftSoundLevel
Dim NudgeRightSoundLevel, NudgeCenterSoundLevel, StartButtonSoundLevel, RollingSoundFactor

CoinSoundLevel = 1            'volume level; range [0, 1]
NudgeLeftSoundLevel = 1        'volume level; range [0, 1]
NudgeRightSoundLevel = 1        'volume level; range [0, 1]
NudgeCenterSoundLevel = 1        'volume level; range [0, 1]
StartButtonSoundLevel = 0.1      'volume level; range [0, 1]
PlungerReleaseSoundLevel = 0.8 '1 wjr   'volume level; range [0, 1]
PlungerPullSoundLevel = 1        'volume level; range [0, 1]
RollingSoundFactor = 1.1 / 5

'///////////////////////-----Solenoids, Kickers and Flash Relays-----///////////////////////
Dim FlipperUpAttackMinimumSoundLevel, FlipperUpAttackMaximumSoundLevel, FlipperUpAttackLeftSoundLevel, FlipperUpAttackRightSoundLevel
Dim FlipperUpSoundLevel, FlipperDownSoundLevel, FlipperLeftHitParm, FlipperRightHitParm
Dim SlingshotSoundLevel, BumperSoundFactor, KnockerSoundLevel

FlipperUpAttackMinimumSoundLevel = 0.010    'volume level; range [0, 1]
FlipperUpAttackMaximumSoundLevel = 0.635    'volume level; range [0, 1]
FlipperUpSoundLevel = 1.0            'volume level; range [0, 1]
FlipperDownSoundLevel = 0.45          'volume level; range [0, 1]
FlipperLeftHitParm = FlipperUpSoundLevel    'sound helper; not configurable
FlipperRightHitParm = FlipperUpSoundLevel    'sound helper; not configurable
SlingshotSoundLevel = 0.95            'volume level; range [0, 1]
BumperSoundFactor = 4.25            'volume multiplier; must not be zero
KnockerSoundLevel = 1              'volume level; range [0, 1]

'///////////////////////-----Ball Drops, Bumps and Collisions-----///////////////////////
Dim RubberStrongSoundFactor, RubberWeakSoundFactor, RubberFlipperSoundFactor,BallWithBallCollisionSoundFactor
Dim BallBouncePlayfieldSoftFactor, BallBouncePlayfieldHardFactor, PlasticRampDropToPlayfieldSoundLevel, WireRampDropToPlayfieldSoundLevel, DelayedBallDropOnPlayfieldSoundLevel
Dim WallImpactSoundFactor, MetalImpactSoundFactor, SubwaySoundLevel, SubwayEntrySoundLevel, ScoopEntrySoundLevel
Dim SaucerLockSoundLevel, SaucerKickSoundLevel

BallWithBallCollisionSoundFactor = 3.2      'volume multiplier; must not be zero
RubberStrongSoundFactor = 0.055 / 5      'volume multiplier; must not be zero
RubberWeakSoundFactor = 0.075 / 5        'volume multiplier; must not be zero
RubberFlipperSoundFactor = 0.075 / 5      'volume multiplier; must not be zero
BallBouncePlayfieldSoftFactor = 0.025      'volume multiplier; must not be zero
BallBouncePlayfieldHardFactor = 0.025      'volume multiplier; must not be zero
DelayedBallDropOnPlayfieldSoundLevel = 0.8    'volume level; range [0, 1]
WallImpactSoundFactor = 0.075          'volume multiplier; must not be zero
MetalImpactSoundFactor = 0.075 / 3
SaucerLockSoundLevel = 0.8
SaucerKickSoundLevel = 0.8

'///////////////////////-----Gates, Spinners, Rollovers and Targets-----///////////////////////

Dim GateSoundLevel, TargetSoundFactor, SpinnerSoundLevel, RolloverSoundLevel, DTSoundLevel

GateSoundLevel = 0.5 / 5      'volume level; range [0, 1]
TargetSoundFactor = 0.0025 * 10  'volume multiplier; must not be zero
DTSoundLevel = 0.25        'volume multiplier; must not be zero
RolloverSoundLevel = 0.25      'volume level; range [0, 1]
SpinnerSoundLevel = 0.5      'volume level; range [0, 1]

'///////////////////////-----Ball Release, Guides and Drain-----///////////////////////
Dim DrainSoundLevel, BallReleaseSoundLevel, BottomArchBallGuideSoundFactor, FlipperBallGuideSoundFactor

DrainSoundLevel = 0.8          'volume level; range [0, 1]
BallReleaseSoundLevel = 1        'volume level; range [0, 1]
BottomArchBallGuideSoundFactor = 0.2  'volume multiplier; must not be zero
FlipperBallGuideSoundFactor = 0.015  'volume multiplier; must not be zero

'///////////////////////-----Loops and Lanes-----///////////////////////
Dim ArchSoundFactor
ArchSoundFactor = 0.025 / 5      'volume multiplier; must not be zero

'/////////////////////////////  SOUND PLAYBACK FUNCTIONS  ////////////////////////////
'/////////////////////////////  POSITIONAL SOUND PLAYBACK METHODS  ////////////////////////////
' Positional sound playback methods will play a sound, depending on the X,Y position of the table element or depending on ActiveBall object position
' These are similar subroutines that are less complicated to use (e.g. simply use standard parameters for the PlaySound call)
' For surround setup - positional sound playback functions will fade between front and rear surround channels and pan between left and right channels
' For stereo setup - positional sound playback functions will only pan between left and right channels
' For mono setup - positional sound playback functions will not pan between left and right channels and will not fade between front and rear channels

' PlaySound full syntax - PlaySound(string, int loopcount, float volume, float pan, float randompitch, int pitch, bool useexisting, bool restart, float front_rear_fade)
' Note - These functions will not work (currently) for walls/slingshots as these do not feature a simple, single X,Y position
Sub PlaySoundAtLevelStatic(playsoundparams, aVol, tableobj)
  PlaySound playsoundparams, 0, min(aVol,1) * VolumeDial, AudioPan(tableobj), 0, 0, 0, 0, AudioFade(tableobj)
End Sub

Sub PlaySoundAtLevelExistingStatic(playsoundparams, aVol, tableobj)
  PlaySound playsoundparams, 0, min(aVol,1) * VolumeDial, AudioPan(tableobj), 0, 0, 1, 0, AudioFade(tableobj)
End Sub

Sub PlaySoundAtLevelStaticLoop(playsoundparams, aVol, tableobj)
  PlaySound playsoundparams, - 1, min(aVol,1) * VolumeDial, AudioPan(tableobj), 0, 0, 0, 0, AudioFade(tableobj)
End Sub

Sub PlaySoundAtLevelStaticRandomPitch(playsoundparams, aVol, randomPitch, tableobj)
  PlaySound playsoundparams, 0, min(aVol,1) * VolumeDial, AudioPan(tableobj), randomPitch, 0, 0, 0, AudioFade(tableobj)
End Sub

Sub PlaySoundAtLevelActiveBall(playsoundparams, aVol)
  PlaySound playsoundparams, 0, min(aVol,1) * VolumeDial, AudioPan(ActiveBall), 0, 0, 0, 0, AudioFade(ActiveBall)
End Sub

Sub PlaySoundAtLevelExistingActiveBall(playsoundparams, aVol)
  PlaySound playsoundparams, 0, min(aVol,1) * VolumeDial, AudioPan(ActiveBall), 0, 0, 1, 0, AudioFade(ActiveBall)
End Sub

Sub PlaySoundAtLeveTimerActiveBall(playsoundparams, aVol, ballvariable)
  PlaySound playsoundparams, 0, min(aVol,1) * VolumeDial, AudioPan(ballvariable), 0, 0, 0, 0, AudioFade(ballvariable)
End Sub

Sub PlaySoundAtLevelTimerExistingActiveBall(playsoundparams, aVol, ballvariable)
  PlaySound playsoundparams, 0, min(aVol,1) * VolumeDial, AudioPan(ballvariable), 0, 0, 1, 0, AudioFade(ballvariable)
End Sub

Sub PlaySoundAtLevelRoll(playsoundparams, aVol, pitch)
  PlaySound playsoundparams, - 1, min(aVol,1) * VolumeDial, AudioPan(tableobj), randomPitch, 0, 0, 0, AudioFade(tableobj)
End Sub

' Previous Positional Sound Subs

Sub PlaySoundAt(soundname, tableobj)
  PlaySound soundname, 1, 1 * VolumeDial, AudioPan(tableobj), 0,0,0, 1, AudioFade(tableobj)
End Sub

Sub PlaySoundAtVol(soundname, tableobj, aVol)
  PlaySound soundname, 1, min(aVol,1) * VolumeDial, AudioPan(tableobj), 0,0,0, 1, AudioFade(tableobj)
End Sub

Sub PlaySoundAtBall(soundname)
  PlaySoundAt soundname, ActiveBall
End Sub

Sub PlaySoundAtBallVol (Soundname, aVol)
  PlaySound soundname, 1,min(aVol,1) * VolumeDial, AudioPan(ActiveBall), 0,0,0, 1, AudioFade(ActiveBall)
End Sub

Sub PlaySoundAtBallVolM (Soundname, aVol)
  PlaySound soundname, 1,min(aVol,1) * VolumeDial, AudioPan(ActiveBall), 0,0,0, 0, AudioFade(ActiveBall)
End Sub

Sub PlaySoundAtVolLoops(sound, tableobj, Vol, Loops)
  PlaySound sound, Loops, Vol * VolumeDial, AudioPan(tableobj), 0,0,0, 1, AudioFade(tableobj)
End Sub

'******************************************************
'  Fleep  Supporting Ball & Sound Functions
'******************************************************

Function AudioFade(tableobj) ' Fades between front and back of the table (for surround systems or 2x2 speakers, etc), depending on the Y position on the table. "table1" is the name of the table
  Dim tmp
  tmp = tableobj.y * 2 / tableheight - 1

  If tmp > 7000 Then
    tmp = 7000
  ElseIf tmp <  - 7000 Then
    tmp =  - 7000
  End If

' Thalamus, AudioFade patched
	If tmp > 0 Then
		AudioFade = CSng(tmp ^ 5) 'was 10
	Else
		AudioFade = CSng( - (( - tmp) ^ 5) ) 'was 10
	End If
End Function

Function AudioPan(tableobj) ' Calculates the pan for a tableobj based on the X position on the table. "table1" is the name of the table
  Dim tmp
  tmp = tableobj.x * 2 / tablewidth - 1

  If tmp > 7000 Then
    tmp = 7000
  ElseIf tmp <  - 7000 Then
    tmp =  - 7000
  End If

' Thalamus, AudioPan patched
	If tmp > 0 Then
		AudioPan = CSng(tmp ^ 5) 'was 10
	Else
		AudioPan = CSng( - (( - tmp) ^ 5) ) 'was 10
	End If
End Function

Function Vol(ball) ' Calculates the volume of the sound based on the ball speed
  Vol = CSng(BallVel(ball) ^ 2)
End Function

Function Volz(ball) ' Calculates the volume of the sound based on the ball speed
  Volz = CSng((ball.velz) ^ 2)
End Function

Function Pitch(ball) ' Calculates the pitch of the sound based on the ball speed
  Pitch = BallVel(ball) * 20
End Function

Function BallVel(ball) 'Calculates the ball speed
  BallVel = Int(Sqr((ball.VelX ^ 2) + (ball.VelY ^ 2) ) )
End Function

Function VolPlayfieldRoll(ball) ' Calculates the roll volume of the sound based on the ball speed
  VolPlayfieldRoll = RollingSoundFactor * 0.0005 * CSng(BallVel(ball) ^ 3)
End Function

Function PitchPlayfieldRoll(ball) ' Calculates the roll pitch of the sound based on the ball speed
  PitchPlayfieldRoll = BallVel(ball) ^ 2 * 15
End Function

Function RndInt(min, max) ' Sets a random number integer between min and max
  RndInt = Int(Rnd() * (max - min + 1) + min)
End Function

Function RndNum(min, max) ' Sets a random number between min and max
  RndNum = Rnd() * (max - min) + min
End Function

'/////////////////////////////  GENERAL SOUND SUBROUTINES  ////////////////////////////

Sub SoundStartButton()
  PlaySound ("Start_Button"), 0, StartButtonSoundLevel, 0, 0.25
End Sub

Sub SoundNudgeLeft()
  PlaySound ("Nudge_" & Int(Rnd * 2) + 1), 0, NudgeLeftSoundLevel * VolumeDial, - 0.1, 0.25
End Sub

Sub SoundNudgeRight()
  PlaySound ("Nudge_" & Int(Rnd * 2) + 1), 0, NudgeRightSoundLevel * VolumeDial, 0.1, 0.25
End Sub

Sub SoundNudgeCenter()
  PlaySound ("Nudge_" & Int(Rnd * 2) + 1), 0, NudgeCenterSoundLevel * VolumeDial, 0, 0.25
End Sub

Sub SoundPlungerPull()
  PlaySoundAtLevelStatic ("Plunger_Pull_1"), PlungerPullSoundLevel, Plunger
End Sub

Sub SoundPlungerReleaseBall()
  PlaySoundAtLevelStatic ("Plunger_Release_Ball"), PlungerReleaseSoundLevel, Plunger
End Sub

Sub SoundPlungerReleaseNoBall()
  PlaySoundAtLevelStatic ("Plunger_Release_No_Ball"), PlungerReleaseSoundLevel, Plunger
End Sub

'/////////////////////////////  KNOCKER SOLENOID  ////////////////////////////

Sub KnockerSolenoid()
  PlaySoundAtLevelStatic SoundFX("Knocker_1",DOFKnocker), KnockerSoundLevel, KnockerPosition
End Sub

'/////////////////////////////  DRAIN SOUNDS  ////////////////////////////

Sub RandomSoundDrain(drainswitch)
  PlaySoundAtLevelStatic ("Drain_" & Int(Rnd * 11) + 1), DrainSoundLevel, drainswitch
End Sub

'/////////////////////////////  TROUGH BALL RELEASE SOLENOID SOUNDS  ////////////////////////////

Sub RandomSoundBallRelease(drainswitch)
  PlaySoundAtLevelStatic SoundFX("BallRelease" & Int(Rnd * 7) + 1,DOFContactors), BallReleaseSoundLevel, drainswitch
End Sub

'/////////////////////////////  SLINGSHOT SOLENOID SOUNDS  ////////////////////////////

Sub RandomSoundSlingshotLeft(sling)
  PlaySoundAtLevelStatic SoundFX("Sling_L" & Int(Rnd * 10) + 1,DOFContactors), SlingshotSoundLevel, Sling
End Sub

Sub RandomSoundSlingshotRight(sling)
  PlaySoundAtLevelStatic SoundFX("Sling_R" & Int(Rnd * 8) + 1,DOFContactors), SlingshotSoundLevel, Sling
End Sub

'/////////////////////////////  BUMPER SOLENOID SOUNDS  ////////////////////////////

Sub RandomSoundBumperTop(Bump)
  PlaySoundAtLevelStatic SoundFX("Bumpers_Top_" & Int(Rnd * 5) + 1,DOFContactors), Vol(ActiveBall) * BumperSoundFactor, Bump
End Sub

Sub RandomSoundBumperMiddle(Bump)
  PlaySoundAtLevelStatic SoundFX("Bumpers_Middle_" & Int(Rnd * 5) + 1,DOFContactors), Vol(ActiveBall) * BumperSoundFactor, Bump
End Sub

Sub RandomSoundBumperBottom(Bump)
  PlaySoundAtLevelStatic SoundFX("Bumpers_Bottom_" & Int(Rnd * 5) + 1,DOFContactors), Vol(ActiveBall) * BumperSoundFactor, Bump
End Sub

'/////////////////////////////  SPINNER SOUNDS  ////////////////////////////

Sub SoundSpinner(spinnerswitch)
  PlaySoundAtLevelStatic ("Spinner"), SpinnerSoundLevel, spinnerswitch
End Sub

'/////////////////////////////  FLIPPER BATS SOUND SUBROUTINES  ////////////////////////////
'/////////////////////////////  FLIPPER BATS SOLENOID ATTACK SOUND  ////////////////////////////

Sub SoundFlipperUpAttackLeft(flipper)
  FlipperUpAttackLeftSoundLevel = RndNum(FlipperUpAttackMinimumSoundLevel, FlipperUpAttackMaximumSoundLevel)
  PlaySoundAtLevelStatic SoundFX("Flipper_Attack-L01",DOFFlippers), FlipperUpAttackLeftSoundLevel, flipper
End Sub

Sub SoundFlipperUpAttackRight(flipper)
  FlipperUpAttackRightSoundLevel = RndNum(FlipperUpAttackMinimumSoundLevel, FlipperUpAttackMaximumSoundLevel)
  PlaySoundAtLevelStatic SoundFX("Flipper_Attack-R01",DOFFlippers), FlipperUpAttackLeftSoundLevel, flipper
End Sub

'/////////////////////////////  FLIPPER BATS SOLENOID CORE SOUND  ////////////////////////////

Sub RandomSoundFlipperUpLeft(flipper)
  PlaySoundAtLevelStatic SoundFX("Flipper_L0" & Int(Rnd * 9) + 1,DOFFlippers), FlipperLeftHitParm, Flipper
End Sub

Sub RandomSoundFlipperUpRight(flipper)
  PlaySoundAtLevelStatic SoundFX("Flipper_R0" & Int(Rnd * 9) + 1,DOFFlippers), FlipperRightHitParm, Flipper
End Sub

Sub RandomSoundReflipUpLeft(flipper)
  PlaySoundAtLevelStatic SoundFX("Flipper_ReFlip_L0" & Int(Rnd * 3) + 1,DOFFlippers), (RndNum(0.8, 1)) * FlipperUpSoundLevel, Flipper
End Sub

Sub RandomSoundReflipUpRight(flipper)
  PlaySoundAtLevelStatic SoundFX("Flipper_ReFlip_R0" & Int(Rnd * 3) + 1,DOFFlippers), (RndNum(0.8, 1)) * FlipperUpSoundLevel, Flipper
End Sub

Sub RandomSoundFlipperDownLeft(flipper)
  PlaySoundAtLevelStatic SoundFX("Flipper_Left_Down_" & Int(Rnd * 7) + 1,DOFFlippers), FlipperDownSoundLevel, Flipper
End Sub

Sub RandomSoundFlipperDownRight(flipper)
  PlaySoundAtLevelStatic SoundFX("Flipper_Right_Down_" & Int(Rnd * 8) + 1,DOFFlippers), FlipperDownSoundLevel, Flipper
End Sub

'/////////////////////////////  FLIPPER BATS BALL COLLIDE SOUND  ////////////////////////////

Sub LeftFlipperCollide(parm)
  FlipperLeftHitParm = parm / 10
  If FlipperLeftHitParm > 1 Then
    FlipperLeftHitParm = 1
  End If
  FlipperLeftHitParm = FlipperUpSoundLevel * FlipperLeftHitParm
  RandomSoundRubberFlipper(parm)
End Sub

Sub RightFlipperCollide(parm)
  FlipperRightHitParm = parm / 10
  If FlipperRightHitParm > 1 Then
    FlipperRightHitParm = 1
  End If
  FlipperRightHitParm = FlipperUpSoundLevel * FlipperRightHitParm
  RandomSoundRubberFlipper(parm)
End Sub

Sub RandomSoundRubberFlipper(parm)
  PlaySoundAtLevelActiveBall ("Flipper_Rubber_" & Int(Rnd * 7) + 1), parm * RubberFlipperSoundFactor
End Sub

'/////////////////////////////  ROLLOVER SOUNDS  ////////////////////////////

Sub RandomSoundRollover()
  PlaySoundAtLevelActiveBall ("Rollover_" & Int(Rnd * 4) + 1), RolloverSoundLevel
End Sub

Sub Rollovers_Hit(idx)
  RandomSoundRollover
End Sub

'/////////////////////////////  VARIOUS PLAYFIELD SOUND SUBROUTINES  ////////////////////////////
'/////////////////////////////  RUBBERS AND POSTS  ////////////////////////////
'/////////////////////////////  RUBBERS - EVENTS  ////////////////////////////

Sub Rubbers_Hit(idx)
  Dim finalspeed
  finalspeed = Sqr(ActiveBall.velx * ActiveBall.velx + ActiveBall.vely * ActiveBall.vely)
  If finalspeed > 5 Then
    RandomSoundRubberStrong 1
  End If
  If finalspeed <= 5 Then
    RandomSoundRubberWeak()
  End If
End Sub

'/////////////////////////////  RUBBERS AND POSTS - STRONG IMPACTS  ////////////////////////////

Sub RandomSoundRubberStrong(voladj)
  Select Case Int(Rnd * 10) + 1
    Case 1
      PlaySoundAtLevelActiveBall ("Rubber_Strong_1"), Vol(ActiveBall) * RubberStrongSoundFactor * voladj
    Case 2
      PlaySoundAtLevelActiveBall ("Rubber_Strong_2"), Vol(ActiveBall) * RubberStrongSoundFactor * voladj
    Case 3
      PlaySoundAtLevelActiveBall ("Rubber_Strong_3"), Vol(ActiveBall) * RubberStrongSoundFactor * voladj
    Case 4
      PlaySoundAtLevelActiveBall ("Rubber_Strong_4"), Vol(ActiveBall) * RubberStrongSoundFactor * voladj
    Case 5
      PlaySoundAtLevelActiveBall ("Rubber_Strong_5"), Vol(ActiveBall) * RubberStrongSoundFactor * voladj
    Case 6
      PlaySoundAtLevelActiveBall ("Rubber_Strong_6"), Vol(ActiveBall) * RubberStrongSoundFactor * voladj
    Case 7
      PlaySoundAtLevelActiveBall ("Rubber_Strong_7"), Vol(ActiveBall) * RubberStrongSoundFactor * voladj
    Case 8
      PlaySoundAtLevelActiveBall ("Rubber_Strong_8"), Vol(ActiveBall) * RubberStrongSoundFactor * voladj
    Case 9
      PlaySoundAtLevelActiveBall ("Rubber_Strong_9"), Vol(ActiveBall) * RubberStrongSoundFactor * voladj
    Case 10
      PlaySoundAtLevelActiveBall ("Rubber_1_Hard"), Vol(ActiveBall) * RubberStrongSoundFactor * 0.6 * voladj
  End Select
End Sub

'/////////////////////////////  RUBBERS AND POSTS - WEAK IMPACTS  ////////////////////////////

Sub RandomSoundRubberWeak()
  PlaySoundAtLevelActiveBall ("Rubber_" & Int(Rnd * 9) + 1), Vol(ActiveBall) * RubberWeakSoundFactor
End Sub

'/////////////////////////////  WALL IMPACTS  ////////////////////////////

Sub Walls_Hit(idx)
  RandomSoundWall()
End Sub

Sub RandomSoundWall()
  Dim finalspeed
  finalspeed = Sqr(ActiveBall.velx * ActiveBall.velx + ActiveBall.vely * ActiveBall.vely)
  If finalspeed > 16 Then
    Select Case Int(Rnd * 5) + 1
      Case 1
        PlaySoundAtLevelExistingActiveBall ("Wall_Hit_1"), Vol(ActiveBall) * WallImpactSoundFactor
      Case 2
        PlaySoundAtLevelExistingActiveBall ("Wall_Hit_2"), Vol(ActiveBall) * WallImpactSoundFactor
      Case 3
        PlaySoundAtLevelExistingActiveBall ("Wall_Hit_5"), Vol(ActiveBall) * WallImpactSoundFactor
      Case 4
        PlaySoundAtLevelExistingActiveBall ("Wall_Hit_7"), Vol(ActiveBall) * WallImpactSoundFactor
      Case 5
        PlaySoundAtLevelExistingActiveBall ("Wall_Hit_9"), Vol(ActiveBall) * WallImpactSoundFactor
    End Select
  End If
  If finalspeed >= 6 And finalspeed <= 16 Then
    Select Case Int(Rnd * 4) + 1
      Case 1
        PlaySoundAtLevelExistingActiveBall ("Wall_Hit_3"), Vol(ActiveBall) * WallImpactSoundFactor
      Case 2
        PlaySoundAtLevelExistingActiveBall ("Wall_Hit_4"), Vol(ActiveBall) * WallImpactSoundFactor
      Case 3
        PlaySoundAtLevelExistingActiveBall ("Wall_Hit_6"), Vol(ActiveBall) * WallImpactSoundFactor
      Case 4
        PlaySoundAtLevelExistingActiveBall ("Wall_Hit_8"), Vol(ActiveBall) * WallImpactSoundFactor
    End Select
  End If
  If finalspeed < 6 Then
    Select Case Int(Rnd * 3) + 1
      Case 1
        PlaySoundAtLevelExistingActiveBall ("Wall_Hit_4"), Vol(ActiveBall) * WallImpactSoundFactor
      Case 2
        PlaySoundAtLevelExistingActiveBall ("Wall_Hit_6"), Vol(ActiveBall) * WallImpactSoundFactor
      Case 3
        PlaySoundAtLevelExistingActiveBall ("Wall_Hit_8"), Vol(ActiveBall) * WallImpactSoundFactor
    End Select
  End If
End Sub

'/////////////////////////////  METAL TOUCH SOUNDS  ////////////////////////////

Sub RandomSoundMetal()
  PlaySoundAtLevelActiveBall ("Metal_Touch_" & Int(Rnd * 13) + 1), Vol(ActiveBall) * MetalImpactSoundFactor
End Sub

'/////////////////////////////  METAL - EVENTS  ////////////////////////////

Sub Metals_Hit (idx)
  RandomSoundMetal
End Sub

Sub ShooterDiverter_collide(idx)
  RandomSoundMetal
End Sub

'/////////////////////////////  BOTTOM ARCH BALL GUIDE  ////////////////////////////
'/////////////////////////////  BOTTOM ARCH BALL GUIDE - SOFT BOUNCES  ////////////////////////////

Sub RandomSoundBottomArchBallGuide()
  Dim finalspeed
  finalspeed = Sqr(ActiveBall.velx * ActiveBall.velx + ActiveBall.vely * ActiveBall.vely)
  If finalspeed > 16 Then
    PlaySoundAtLevelActiveBall ("Apron_Bounce_" & Int(Rnd * 2) + 1), Vol(ActiveBall) * BottomArchBallGuideSoundFactor
  End If
  If finalspeed >= 6 And finalspeed <= 16 Then
    Select Case Int(Rnd * 2) + 1
      Case 1
        PlaySoundAtLevelActiveBall ("Apron_Bounce_1"), Vol(ActiveBall) * BottomArchBallGuideSoundFactor
      Case 2
        PlaySoundAtLevelActiveBall ("Apron_Bounce_Soft_1"), Vol(ActiveBall) * BottomArchBallGuideSoundFactor
    End Select
  End If
  If finalspeed < 6 Then
    Select Case Int(Rnd * 2) + 1
      Case 1
        PlaySoundAtLevelActiveBall ("Apron_Bounce_Soft_1"), Vol(ActiveBall) * BottomArchBallGuideSoundFactor
      Case 2
        PlaySoundAtLevelActiveBall ("Apron_Medium_3"), Vol(ActiveBall) * BottomArchBallGuideSoundFactor
    End Select
  End If
End Sub

'/////////////////////////////  BOTTOM ARCH BALL GUIDE - HARD HITS  ////////////////////////////

Sub RandomSoundBottomArchBallGuideHardHit()
  PlaySoundAtLevelActiveBall ("Apron_Hard_Hit_" & Int(Rnd * 3) + 1), BottomArchBallGuideSoundFactor * 0.25
End Sub

Sub Apron_Hit (idx)
  If Abs(cor.ballvelx(ActiveBall.id) < 4) And cor.ballvely(ActiveBall.id) > 7 Then
    RandomSoundBottomArchBallGuideHardHit()
  Else
    RandomSoundBottomArchBallGuide
  End If
End Sub

'/////////////////////////////  FLIPPER BALL GUIDE  ////////////////////////////

Sub RandomSoundFlipperBallGuide()
  Dim finalspeed
  finalspeed = Sqr(ActiveBall.velx * ActiveBall.velx + ActiveBall.vely * ActiveBall.vely)
  If finalspeed > 16 Then
    Select Case Int(Rnd * 2) + 1
      Case 1
        PlaySoundAtLevelActiveBall ("Apron_Hard_1"),  Vol(ActiveBall) * FlipperBallGuideSoundFactor
      Case 2
        PlaySoundAtLevelActiveBall ("Apron_Hard_2"),  Vol(ActiveBall) * 0.8 * FlipperBallGuideSoundFactor
    End Select
  End If
  If finalspeed >= 6 And finalspeed <= 16 Then
    PlaySoundAtLevelActiveBall ("Apron_Medium_" & Int(Rnd * 3) + 1),  Vol(ActiveBall) * FlipperBallGuideSoundFactor
  End If
  If finalspeed < 6 Then
    PlaySoundAtLevelActiveBall ("Apron_Soft_" & Int(Rnd * 7) + 1),  Vol(ActiveBall) * FlipperBallGuideSoundFactor
  End If
End Sub

'/////////////////////////////  TARGET HIT SOUNDS  ////////////////////////////

Sub RandomSoundTargetHitStrong()
  PlaySoundAtLevelActiveBall SoundFX("Target_Hit_" & Int(Rnd * 4) + 5,DOFTargets), Vol(ActiveBall) * 0.45 * TargetSoundFactor
End Sub

Sub RandomSoundTargetHitWeak()
  PlaySoundAtLevelActiveBall SoundFX("Target_Hit_" & Int(Rnd * 4) + 1,DOFTargets), Vol(ActiveBall) * TargetSoundFactor
End Sub

Sub PlayTargetSound()
  Dim finalspeed
  finalspeed = Sqr(ActiveBall.velx * ActiveBall.velx + ActiveBall.vely * ActiveBall.vely)
  If finalspeed > 10 Then
    RandomSoundTargetHitStrong()
    RandomSoundBallBouncePlayfieldSoft ActiveBall
  Else
    RandomSoundTargetHitWeak()
  End If
End Sub

Sub Targets_Hit (idx)
  PlayTargetSound
End Sub

'/////////////////////////////  BALL BOUNCE SOUNDS  ////////////////////////////

Sub RandomSoundBallBouncePlayfieldSoft(aBall)
  Select Case Int(Rnd * 9) + 1
    Case 1
      PlaySoundAtLevelStatic ("Ball_Bounce_Playfield_Soft_1"), volz(aBall) * BallBouncePlayfieldSoftFactor, aBall
    Case 2
      PlaySoundAtLevelStatic ("Ball_Bounce_Playfield_Soft_2"), volz(aBall) * BallBouncePlayfieldSoftFactor * 0.5, aBall
    Case 3
      PlaySoundAtLevelStatic ("Ball_Bounce_Playfield_Soft_3"), volz(aBall) * BallBouncePlayfieldSoftFactor * 0.8, aBall
    Case 4
      PlaySoundAtLevelStatic ("Ball_Bounce_Playfield_Soft_4"), volz(aBall) * BallBouncePlayfieldSoftFactor * 0.5, aBall
    Case 5
      PlaySoundAtLevelStatic ("Ball_Bounce_Playfield_Soft_5"), volz(aBall) * BallBouncePlayfieldSoftFactor, aBall
    Case 6
      PlaySoundAtLevelStatic ("Ball_Bounce_Playfield_Hard_1"), volz(aBall) * BallBouncePlayfieldSoftFactor * 0.2, aBall
    Case 7
      PlaySoundAtLevelStatic ("Ball_Bounce_Playfield_Hard_2"), volz(aBall) * BallBouncePlayfieldSoftFactor * 0.2, aBall
    Case 8
      PlaySoundAtLevelStatic ("Ball_Bounce_Playfield_Hard_5"), volz(aBall) * BallBouncePlayfieldSoftFactor * 0.2, aBall
    Case 9
      PlaySoundAtLevelStatic ("Ball_Bounce_Playfield_Hard_7"), volz(aBall) * BallBouncePlayfieldSoftFactor * 0.3, aBall
  End Select
End Sub

Sub RandomSoundBallBouncePlayfieldHard(aBall)
  PlaySoundAtLevelStatic ("Ball_Bounce_Playfield_Hard_" & Int(Rnd * 7) + 1), volz(aBall) * BallBouncePlayfieldHardFactor, aBall
End Sub

'/////////////////////////////  DELAYED DROP - TO PLAYFIELD - SOUND  ////////////////////////////

Sub RandomSoundDelayedBallDropOnPlayfield(aBall)
  Select Case Int(Rnd * 5) + 1
    Case 1
      PlaySoundAtLevelStatic ("Ball_Drop_Playfield_1_Delayed"), DelayedBallDropOnPlayfieldSoundLevel, aBall
    Case 2
      PlaySoundAtLevelStatic ("Ball_Drop_Playfield_2_Delayed"), DelayedBallDropOnPlayfieldSoundLevel, aBall
    Case 3
      PlaySoundAtLevelStatic ("Ball_Drop_Playfield_3_Delayed"), DelayedBallDropOnPlayfieldSoundLevel, aBall
    Case 4
      PlaySoundAtLevelStatic ("Ball_Drop_Playfield_4_Delayed"), DelayedBallDropOnPlayfieldSoundLevel, aBall
    Case 5
      PlaySoundAtLevelStatic ("Ball_Drop_Playfield_5_Delayed"), DelayedBallDropOnPlayfieldSoundLevel, aBall
  End Select
End Sub

'/////////////////////////////  BALL GATES AND BRACKET GATES SOUNDS  ////////////////////////////

Sub SoundPlayfieldGate()
  PlaySoundAtLevelStatic ("Gate_FastTrigger_" & Int(Rnd * 2) + 1), GateSoundLevel, ActiveBall
End Sub

Sub SoundHeavyGate()
  PlaySoundAtLevelStatic ("Gate_2"), GateSoundLevel, ActiveBall
End Sub

Sub Gates_hit(idx)
  SoundHeavyGate
End Sub

Sub GatesWire_hit(idx)
  SoundPlayfieldGate
End Sub

'/////////////////////////////  LEFT LANE ENTRANCE - SOUNDS  ////////////////////////////

Sub RandomSoundLeftArch()
  PlaySoundAtLevelActiveBall ("Arch_L" & Int(Rnd * 4) + 1), Vol(ActiveBall) * ArchSoundFactor
End Sub

Sub RandomSoundRightArch()
  PlaySoundAtLevelActiveBall ("Arch_R" & Int(Rnd * 4) + 1), Vol(ActiveBall) * ArchSoundFactor
End Sub

Sub Arch1_hit()
  If ActiveBall.velx > 1 Then SoundPlayfieldGate
  StopSound "Arch_L1"
  StopSound "Arch_L2"
  StopSound "Arch_L3"
  StopSound "Arch_L4"
End Sub

Sub Arch1_unhit()
  If ActiveBall.velx <  - 8 Then
    RandomSoundRightArch
  End If
End Sub

Sub Arch2_hit()
  If ActiveBall.velx < 1 Then SoundPlayfieldGate
  StopSound "Arch_R1"
  StopSound "Arch_R2"
  StopSound "Arch_R3"
  StopSound "Arch_R4"
End Sub

Sub Arch2_unhit()
  If ActiveBall.velx > 10 Then
    RandomSoundLeftArch
  End If
End Sub

'/////////////////////////////  SAUCERS (KICKER HOLES)  ////////////////////////////

Sub SoundSaucerLock()
  PlaySoundAtLevelStatic ("Saucer_Enter_" & Int(Rnd * 2) + 1), SaucerLockSoundLevel, ActiveBall
End Sub

Sub SoundSaucerKick(scenario, saucer)
  Select Case scenario
    Case 0
      PlaySoundAtLevelStatic SoundFX("Saucer_Empty", DOFContactors), SaucerKickSoundLevel, saucer
    Case 1
      PlaySoundAtLevelStatic SoundFX("Saucer_Kick", DOFContactors), SaucerKickSoundLevel, saucer
  End Select
End Sub

'/////////////////////////////  BALL COLLISION SOUND  ////////////////////////////

Sub OnBallBallCollision(ball1, ball2, velocity)

  FlipperCradleCollision ball1, ball2, velocity

  Dim snd
  Select Case Int(Rnd * 7) + 1
    Case 1
      snd = "Ball_Collide_1"
    Case 2
      snd = "Ball_Collide_2"
    Case 3
      snd = "Ball_Collide_3"
    Case 4
      snd = "Ball_Collide_4"
    Case 5
      snd = "Ball_Collide_5"
    Case 6
      snd = "Ball_Collide_6"
    Case 7
      snd = "Ball_Collide_7"
  End Select

  PlaySound (snd), 0, CSng(velocity) ^ 2 / 200 * BallWithBallCollisionSoundFactor * VolumeDial, AudioPan(ball1), 0, Pitch(ball1), 0, 0, AudioFade(ball1)

  'Release ball in radar if there is a collision with another ball
  If BallInRadarLock Then
    If ball1.id = RadarBall.id or ball2.id = RadarBall.id Then SolRadarMagnet False
  End If

End Sub


'///////////////////////////  DROP TARGET HIT SOUNDS  ///////////////////////////

Sub RandomSoundDropTargetReset(obj)
  PlaySoundAtLevelStatic SoundFX("Drop_Target_Reset_" & Int(Rnd * 6) + 1,DOFContactors), 1, obj
End Sub

Sub SoundDropTargetDrop(obj)
  PlaySoundAtLevelStatic ("Drop_Target_Down_" & Int(Rnd * 6) + 1), 200, obj
End Sub

'/////////////////////////////  GI AND FLASHER RELAYS  ////////////////////////////

Const RelayFlashSoundLevel = 0.315  'volume level; range [0, 1];
Const RelayGISoundLevel = 1.05    'volume level; range [0, 1];

Sub Sound_GI_Relay(toggle, obj)
  Select Case toggle
    Case 1
      PlaySoundAtLevelStatic ("Relay_GI_On"), 0.025 * RelayGISoundLevel, obj
    Case 0
      PlaySoundAtLevelStatic ("Relay_GI_Off"), 0.025 * RelayGISoundLevel, obj
  End Select
End Sub

Sub Sound_Flash_Relay(toggle, obj)
  Select Case toggle
    Case 1
      PlaySoundAtLevelStatic ("Relay_Flash_On"), 0.025 * RelayFlashSoundLevel, obj
    Case 0
      PlaySoundAtLevelStatic ("Relay_Flash_Off"), 0.025 * RelayFlashSoundLevel, obj
  End Select
End Sub

'/////////////////////////////////////////////////////////////////
'         End Mechanical Sounds
'/////////////////////////////////////////////////////////////////




'******************************************************
'   ZVRR:  VR stuff
'******************************************************

Sub SetupRoom
  DIM VRThings
  For each VRThings in DesktopLights:VRThings.Visible = 0: Next
  If VRRoom > 0 Then
    For each VRThings in VRBackglass:VRThings.Visible = 1: Next
    For each VRThings in VRSpeaker:VRThings.Visible = 1: Next
    For each VRThings in VRSpeakerLights:VRThings.Visible = 1: Next
    For each VRThings in VRBGGI:VRThings.Visible = 1: Next
    for each VRThings in VR_Cab:VRThings.visible = 1:Next
    If VRRoom = 1 Then
      for each VRThings in VR_Min:VRThings.visible = 1:Next
      for each VRThings in VRMegaRoom:VRThings.visible = 0:Next
      PinCab_Metal_Rear.visible=True
      PinCab_Metal_RearMega.visible=False
    End If
    If VRRoom = 2 Then
      for each VRThings in VR_Min:VRThings.visible = 0:Next
      for each VRThings in VRMegaRoom:VRThings.visible = 1:Next
      VR_Mega002.PlayAnimEndless(0.5)
      VR_Mega003.PlayAnimEndless(0.5)
      VR_Mega004.PlayAnimEndless(0.5)
      VR_Mega005.PlayAnimEndless(0.5)
      VR_Mega014.PlayAnimEndless(0.1)
      PinCab_Metal_Rear.visible=False
    End If
  Else
    For each VRThings in VRBackglass:VRThings.Visible = 0: Next
    For each VRThings in VRSpeaker:VRThings.Visible = 0: Next
    For each VRThings in VRSpeakerLights:VRThings.Visible = 0: Next
    For each VRThings in VRBGGI:VRThings.Visible = 0: Next
    for each VRThings in VR_Cab:VRThings.visible = 0:Next
    for each VRThings in VRMegaRoom:VRThings.visible = 0:Next
    for each VRThings in VR_Min:VRThings.visible = 0:Next
  End If
  PinCab_Metal_LF.visible = 0
  If desktopmode Then
    PinCab_Metal_LF.visible = 1
    For each VRThings in DesktopLights:VRThings.Visible = 1: Next
  End If
End Sub


' Set Up VR Backglass Flasher
' ****************************

Sub SetBackglass()
  Dim obj

  For Each obj In VRBackglass
    obj.x = obj.x + 3
    obj.height = - obj.y + 300
    obj.y = -35 'adjusts the distance from the backglass towards the user
    obj.rotx=-86.5
  Next

  For Each obj In VRSpeaker
    obj.x = obj.x + 3
    obj.height = - obj.y + 280
    obj.y = 38 'adjusts the distance from the backglass towards the user
    obj.rotx=-86.5
  Next

  For Each obj In VRSpeakerLights
    obj.x = obj.x + 3
    obj.height = - obj.y + 280
    obj.y = 46 'adjusts the distance from the backglass towards the user
    obj.rotx=-86.5
  Next
End Sub



'  LAMP CALLBACK
' ****************

Set LampCallback = GetRef("UpdateMultipleLamps")

Sub UpdateMultipleLamps()
  IF VRRoom > 0 Then
  If Controller.Lamp(73) = 0 Then: SpeakerG.visible=0: else: SpeakerG.visible=1
  If Controller.Lamp(74) = 0 Then: SpeakerO.visible=0: else: SpeakerO.visible=1
  If Controller.Lamp(75) = 0 Then: SpeakerL.visible=0: else: SpeakerL.visible=1
  If Controller.Lamp(76) = 0 Then: SpeakerD.visible=0: else: SpeakerD.visible=1
  If Controller.Lamp(77) = 0 Then: SpeakerE1.visible=0: else: SpeakerE1.visible=1
  If Controller.Lamp(78) = 0 Then: SpeakerN.visible=0: else: SpeakerN.visible=1
  If Controller.Lamp(79) = 0 Then: SpeakerE2.visible=0: else: SpeakerE2.visible=1
  If Controller.Lamp(80) = 0 Then: SpeakerY.visible=0: else: SpeakerY.visible=1
  If Controller.Lamp(72) = 0 Then: SpeakerE3.visible=0: else: SpeakerE3.visible=1
  End If
End Sub


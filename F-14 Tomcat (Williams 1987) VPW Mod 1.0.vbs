'F-14 Tomcat (Williams 1987)
'
'########           ##   ##           ########  #######  ##     ##  ######     ###    ########
'##               ####   ##    ##        ##    ##     ## ###   ### ##    ##   ## ##      ##
'##                 ##   ##    ##        ##    ##     ## #### #### ##        ##   ##     ##
'######   #######   ##   ##    ##        ##    ##     ## ## ### ## ##       ##     ##    ##
'##                 ##   #########       ##    ##     ## ##     ## ##       #########    ##
'##                 ##         ##        ##    ##     ## ##     ## ##    ## ##     ##    ##
'##               ######       ##        ##     #######  ##     ##  ######  ##     ##    ##
'
'VPW Mod, based on wrd1972's version.
'
'Sixtoe - Project Lead
'Apophis - PWM Lighting & Script Tweaks
'P5TK - Staged Flipper Support
'Bord, mcarter, wrd1972, Studlygoorite, Somatik, redbone, darth vito - Support and Testing
'
'Concept & Design by:   Steve Ritchie
'Art by:  Doug Watson
'Mechanics by:  Craig Fitpold
'Music by:  Steve Ritchie, Chris Granner
'Sound by:  Bill Parod, Chris Granner
'Software by:   Eugene Jarvis, Ed Boon
'
'The very talented F14 development team.
'Original VP10 version by "Ganjafarmer"
'Original scripting by 32Assassin
'High-res playfield image by "Marrok"
'Plastics total redraw by "wrd1972"
'Plastics prims by "cyberpez" & "L0stS0ul"
'Plastic and metal ramp prims and textures by "Flupper"
'Flasher domes, inserts prims by "Flupper"
'Flashers, PL inserts and GI by "Flupper"
'DOF by "arngrim"
'Center kicker trough meshes, nFozzy physics configuration, and misc help on clear style flasher domes by Benji
'F14 Plane VR Room & cab beacons by "DJRobX" & "Sixtoe"
'Minimal VR Room and misc tweaks/fixes by "Sixtoe"
'An extra special thanks to "ganjafarmer", "cyberpez","Rothbauerw" and "Flupper" for the hard work and countless hours of development on this table.
'
'Many thanks to the Members of the Vpin Workshop, and countless others in the VPF community for helping me with the development of this table.
'The playfield graphic files are only authorized to be used free of charge and may not be redistributed, reused or sold without permission from Brad1X.
'
'001 - Sixtoe - Tidied up and named layers, replaced physics objects and materials, removed collidable from lots of objects.
'002 - Sixtoe - Did stuff, stuff be broken.
'003 - apophis - Minor updates to get the table working
'004 - Sixtoe - updates, gi blanks on rapid flipper presses
'005 - Bord - VLM test render (Reverted)
'006 - apophis - Updated the table to use PWM fasher and inserts (removed Lampz). Minor physics adjustments.
'007 - Sixtoe - Added blocker walls, re-added f14 plunger lane tracking code, hid some lamps, changed layout of top right trap hole, added flupper pop bumper, changed inlanes
'008 - apophis - Added options menu. Made bumper more red.
'009 - mcarter78 - Back off ball rest in plunger lane to fix floating ball
'010 - Sixtoe - Slowed down ball in plunger lane, made blocker bigger under pop bumper, turned down power on centre kickers and re-aimed them, turned up slingshots, tried to make plungerlane f14 unswept wings but it doesn't fit so reverted, upped playfield friction to 0.24.
'011 - apophis - Intro music default off. Hooked center target insert up to GI. Adjusted kicker strengths per videos.
'012 - Sixtoe - Added ramp rolling sounds (probably?), hooked up clear plastics to gi, tweaked upper flippers,
'013 - PT5K - Added staged flipper support
'014 - Sixtoe - Fixed numerous sound issues, added missing sounds, removed brakes and numerous other tweaks.

'****************************************************************************************************************************************************

Option Explicit
Randomize
SetLocale 1033      'Forces VBS to use english to stop crashes.

On Error Resume Next
ExecuteGlobal GetTextFile("controller.vbs")
If Err Then MsgBox "You need the controller.vbs in order to run this table, available in the vp10 package"
On Error Goto 0

Const tnob = 4      'Total number of balls on the playfield including captive balls.
Const lob = 0     'Total number of locked balls
Const TestVR = False
Const F14ReturnSpeed = 0.5


Const ballsize = 25  'radius
Const ballmass = 1

Const cGameName="f14_l1"
Const UseSolenoids = 2      '1 = Normal Flippers, 2 = Fastflips
Const UseLamps = 1        '0 = Custom lamp handling, 1 = Built-in VPX handling (using light number in light timer)
Const UseGI = 0
Const SSolenoidOn = ""      'Sound sample used for this, obsolete.
Const SSolenoidOff = ""     ' ^
Const SFlipperOn = ""     ' ^
Const SFlipperOff = ""      ' ^
Const SCoin = ""        ' ^

'Const UseVPMModSol = 1   'Old PWM method. Don't use this
Const UseVPMModSol = 2    'Set to 2 for PWM flashers, inserts, and GI. Requires VPinMame 3.6

Dim VRRoom, DesktopMode
DesktopMode = Table1.ShowDT

LoadVPM "03060000", "S11.VBS", 3.26126


'**********************************************************************************************************
'Initiate Table
'**********************************************************************************************************

dim bsTrough, bsLC, bsRC, bsBP, bsKB, cBall1, cBall2, cball3, cball4, plungerim05a,plungerim07a,plungerim10, gBOT

'Dim f,g,h,i,j
Dim Ball(6)
Dim InitTime
Dim TroughTime
Dim EjectTime
Dim MaxBalls
Dim TroughCount
Dim TroughBall(7)
Dim TroughEject
Dim Momentum
Dim UpperGIon
Dim Multiball
Dim BallsInPlay
Dim iBall
Dim fgBall

Sub Table1_Init
  vpmInit Me
  vpmMapLights AllLamps     'Make a collection called "AllLamps" and put all the light objects in it.
  With Controller
    .GameName = cGameName
    If Err Then MsgBox "Can't start Game" & cGameName & vbNewLine & Err.Description : Exit Sub
    .SplashInfoLine = "F-14 Tomcat (Williams 1987)"&chr(13)&"wrd1972 & VPW"
    .HandleKeyboard = 0
    .ShowTitle = 0
    .ShowDMDOnly = 1
    .ShowFrame = 0
    .HandleMechanics = 0
  End With
  On Error Resume Next
  Controller.Run
  If Err Then MsgBox Err.Description
  On Error Goto 0

    vpmNudge.TiltSwitch = swTilt
    vpmNudge.Sensitivity = 5
    vpmNudge.TiltObj = Array(Bumper1, LeftSlingshot, RightSlingshot)


    ' Init ramp diverters
    Sol21wall.IsDropped = True
    Sol22wall.IsDropped = True

'******************************************************
' Ball Trough Init
'******************************************************

  MaxBalls=4
  InitTime=61
  EjectTime=0
  TroughEject=1
  TroughCount=0
  iBall = 4
  fgBall = false

  set cball1 = sw11.CreateSizedballWithMass(Ballsize,Ballmass)
  set cball2 = sw12.CreateSizedballWithMass(Ballsize,Ballmass)
  set cball3 = sw13.CreateSizedballWithMass(Ballsize,Ballmass)
  set cball4 = sw14.CreateSizedballWithMass(Ballsize,Ballmass)

  Controller.Switch(11) = 1
  Controller.Switch(12) = 1
  Controller.Switch(13) = 1
  Controller.Switch(14) = 1

  gBOT = Array(CBall1,CBall2,CBall3,CBall4)

'******************************************************
'Kickers INIT
'******************************************************

'Sol05A "middle right kicker"
    Const IMPowerSetting05a = 70 'Kick Power 135
    Const IMTime05a = 0.6        ' Time in seconds for Full Plunge
    Set plungerIM05a = New cvpmImpulseP
    With plungerIM05a
        .InitImpulseP Sol05AKick, IMPowerSetting05a, IMTime05a
        .Random 0.3
        .switch 23
'        .InitExitSnd SoundFX("Shooter",DOFContactors), SoundFX("shooter",DOFContactors)
    SoundSaucerKick 1, Sol05AKick
        .CreateEvents "plungerIM05a"
    End With

'Sol07A "Right kicker"
    Const IMPowerSetting07a = 75 'Kick Power 65
    Const IMTime07a = 0.6        ' Time in seconds for Full Plunge
    Set plungerIM07a = New cvpmImpulseP
    With plungerIM07a
        .InitImpulseP Sol07AKick, IMPowerSetting07a, IMTime07a
        .Random 0.3
        .switch 21
'        .InitExitSnd SoundFX("Shooter",DOFContactors), SoundFX("shooter",DOFContactors)
    SoundSaucerKick 1, Sol07AKick
        .CreateEvents "plungerIM07a"
    End With

'Sol10 "middle left kicker"
    Const IMPowerSetting10= 70 'Kick Power 125
    Const IMTime10 = 0.6        ' Time in seconds for Full Plunge
    Set plungerIM10 = New cvpmImpulseP
    With plungerIM10
        .InitImpulseP Sol10Kick, IMPowerSetting10, IMTime10
        .Random 0.3
        .switch 22
'        .InitExitSnd SoundFX("Shooter",DOFContactors), SoundFX("shooter",DOFContactors)
    SoundSaucerKick 1, Sol10Kick
        .CreateEvents "plungerIM10"
    End With

  SetGI 1 'Turn GI on at startup
  SetGIColor
  PrevGameOver = 0
  SetBackGlass
End Sub

'**********
'Timer Code
'**********

Sub FrameTimer_Timer()
  Flipperstimer
  DiverterTimer
  BSUpdate
  UpdateBallBrightness
  RollingUpdate       'Update rolling sounds
  VR_Plunger.Y = 2168 + (5* Plunger.Position) -20
  Displaytimer
End Sub

CorTimer.Interval = 10
Sub CorTimer_Timer(): Cor.Update: End Sub



Sub FlippersTimer()
    BatLeft.RotAndTra8 = LeftFlipper.CurrentAngle
  batleftshadow.objrotz = LeftFlipper.CurrentAngle

    BatLeft1.RotAndTra8 = LeftFlipper1.CurrentAngle
  batleftshadow.objrotz = LeftFlipper.CurrentAngle

    BatRight.RotAndTra8 = RightFlipper.CurrentAngle
  batrightshadow.objrotz = RightFlipper.CurrentAngle

    BatRight1.RotAndTra8 = RightFlipper1.CurrentAngle
  batrightshadow.objrotz = RightFlipper.CurrentAngle
End Sub


'*******************************************
'  ZOPT: User Options
'*******************************************


Dim IntroMusic    ' Play Intro Music Snippet = 0,  No Intro Music Snippet = 1
Dim VRRoomChoice  ' VR Full Room with F14's = 1, VR Minimal Room = 2, VR Ultra Minimal = 3
Dim SideRails     ' Hide Side Rails = 0, Show Side Rails = 1
Dim FlasherAdjust   ' Flasher Blooms (usable range): Off=0-------- 30=MAX brightness
Dim YagovKick     ' 35 = Weak, 50 = Hard, This is normally adjusted through the ROM. Insted, you much adjust it here.
Dim ChooseBats    ' white/red = 0, white/blue = 1, yellow/red = 2
Dim ClearPlasticsDecals ' No Decals = 0, Add Decals = 1
Dim ModelToys     ' No Models = 0, Add Models = 1
Dim F14Plunger    ' No Plunger Toy = 0, Add Plunger Toy = 1
Dim CustomBackwall  ' No Decal = 0, Show USS Nimitz (with Lights) = 1, Show hot-chick decal = 2
Dim CustomCards   ' Normal Cards = 0, Custom cards = 1
Dim VolumeDial      ' Overall Mechanical sound effect volume. Recommended values should be no greater than 1.
Dim BallRollVolume  ' Level of ball rolling volume. Value between 0 and 1
Dim RampRollVolume  ' Level of ramp rolling volume. Value between 0 and 1


' Called when options are tweaked by the player.
' - 0: game has started, good time to load options and adjust accordingly
' - 1: an option has changed
' - 2: options have been reseted
' - 3: player closed the tweak UI, good time to update staticly prerendered parts
' Table1.Option arguments are:
' - option name, minimum value, maximum value, step between valid values, default value, unit (0=None, 1=Percent), an optional arry of literal strings
Dim dspTriggered : dspTriggered = False
Sub Table1_OptionEvent(ByVal eventId)
    If eventId = 1 And Not dspTriggered Then dspTriggered = True : DisableStaticPreRendering = True : End If
  dim v

    ' Intro Music
    IntroMusic = Table1.Option("Intro Music", 0, 1, 1, 1, 0, Array("Play Music", "No Music"))
  PlayMusic

  ' VR Room
    VRRoomChoice = Table1.Option("VR Room", 1, 3, 1, 1, 0, Array("F14 Room", "Minimal Room", "Ultra Minimal"))
  If RenderingMode = 2 or TestVR = True Then VRRoom = VRRoomChoice Else VRRoom = 0
  SetupVRRoom

  ' Side Rails
    SideRails = Table1.Option("Side Rails", 0, 1, 1, 1, 0, Array("Hide", "Show"))

  ' Flasher Intensity
    v = Table1.Option("Flasher Intensity", 0, 3, 1, 2, 0, Array("Min", "Low", "Medium", "High"))
  FlasherAdjust = v*10

  ' Yagov Kick Strength
    v = Table1.Option("Yagov Kick Strength", 0, 1, 1, 1, 0, Array("Weak", "Strong"))
  YagovKick = 35 + v*15

  ' Flipper Colors
    ChooseBats = Table1.Option("Flipper Colors", 0, 2, 1, 0, 0, Array("White/Red", "White/Blue", "Yellow/Red"))

  ' Plastics Decals
    ClearPlasticsDecals = Table1.Option("Plastics Decals", 0, 1, 1, 0, 0, Array("No Decals", "Add Decals"))

  ' Model Toy
    ModelToys = Table1.Option("F14 Model Toys", 0, 1, 1, 1, 0, Array("No Models", "Add Models"))

  ' Plunger Toys
    F14Plunger = Table1.Option("F14 Plunger Toy", 0, 1, 1, 1, 0, Array("No Plunger Toy", "Add Plunger Toy"))

  ' Custom Back wall
    CustomBackwall = Table1.Option("Custom Back Wall", 0, 2, 1, 0, 0, Array("No Decal", "Show USS Nimitz (with Lights)", "Show hot-chick decal"))

  ' Custom Cards
    CustomCards = Table1.Option("Custom Cards", 0, 1, 1, 0, 0, Array("Normal Cards", "Custom cards"))

  ' Sound volumes
    VolumeDial = Table1.Option("Mech Volume", 0, 1, 0.01, 0.5, 1)
    BallRollVolume = Table1.Option("Ball Roll Volume", 0, 1, 0.01, 0.5, 1)
  RampRollVolume = Table1.Option("Ramp Roll Volume", 0, 1, 0.01, 0.5, 1)

  UpdateOptions

    If eventId = 3 And dspTriggered Then dspTriggered = False : DisableStaticPreRendering = False : End If
End Sub


Sub UpdateOptions

  '*************************
  'Show Side Rails
  '*************************

  If SideRails =1 and VRRoom < 1 Then
      Leftrail.visible = 1
      Rightrail.visible = 1
      Toprail.visible = 1
  Else
      Leftrail.visible = 0
      Rightrail.visible = 0
      Toprail.visible = 0
  End if

  '*************************
  'Change back wall
  '*************************
  dim NimL
  if custombackwall = 0 then
    BackWall0.sidevisible = true
    BackWall1.sidevisible = False
    BackWall2.sidevisible = False
    for each NimL in nimlights:NimL.visible = 0:Next
  End If
  if custombackwall = 2 then
    BackWall0.sidevisible = False
    BackWall1.sidevisible = False
    BackWall2.sidevisible = true
    for each NimL in nimlights:NimL.visible = 0:Next
  End If
  if custombackwall = 1 then
    BackWall0.sidevisible = False
    BackWall1.sidevisible = true
    BackWall2.sidevisible = False
    for each NimL in nimlights:NimL.visible = 1:Next
  End If


  '*************************
  'Model toys
  '*************************
  DIM ToyThings
  If ModelToys = 1 Then
      for each ToyThings in F14Toy:ToyThings.visible = 1:Next
    Else
      for each ToyThings in F14Toy:ToyThings.visible = 0:Next
  End if

  '*************************
  'Plunger toy
  '*************************
  DIM PlungerToy
  If F14Plunger = 1 Then
      for each PlungerToy in F14PToy:PlungerToy.visible = 1:Next
    Else
      for each PlungerToy in F14PToy:PlungerToy.visible = 0:Next
  End if


  '*************************
  'Flipper Colors
  '*************************
  Select Case ChooseBats
    Case 0
      batleft.visible = 1 : batright.visible = 1 : batleft1.visible = 1 : batright1.visible = 1 : LeftFlipper.visible = 0 : RightFlipper.visible = 0 : LeftFlipper1.visible = 0 : RightFlipper1.visible = 0
      batleftshadow.visible = 1 : batrightshadow.visible = 1 : batleftshadow1.visible = 1 : batrightshadow1.visible = 1
      batleft.image = "flipper_white_red" : batright.image = "flipper_white_red" : batleft1.image = "flipper_white_red" : batright1.image = "flipper_white_red"
    Case 1
      batleft.visible = 1 : batright.visible = 1 : batleft1.visible = 1 : batright1.visible = 1 : LeftFlipper.visible = 0 : RightFlipper.visible = 0 : LeftFlipper1.visible = 0 : RightFlipper1.visible = 0
      batleftshadow.visible = 1 : batrightshadow.visible = 1 : batleftshadow1.visible = 1 : batrightshadow1.visible = 1
      batleft.image = "flipper_white_blue" : batright.image = "flipper_white_blue" : batleft1.image = "flipper_white_blue" : batright1.image = "flipper_white_blue"
    Case 2
      batleft.visible = 0 : batright.visible = 0 : batleft1.visible = 0 : batright1.visible = 0 : LeftFlipper.visible = 1 : RightFlipper.visible = 1 : LeftFlipper1.visible = 1 : RightFlipper1.visible = 1
      batleftshadow.visible = 1 : batrightshadow.visible = 1 : batleftshadow1.visible = 1 : batrightshadow1.visible = 1
  End Select

  '*************************
  'Custom cards
  '*************************
  If CustomCards = 1 Then
    Primitive68.image = "apron_custom"
  Else
    Primitive68.image = "apron williams ALT"
  End If
'
  If ClearPlasticsDecals = 1 Then
    pRamp005.image = "rampplastic"
    pRamp006.image = "rightplastics"
  ' pClearPlastic1.image = "clearplastics_decals"
  ' pClearPlastic2.image = "clearplastics_decals"
  ' pClearPlastic3.image = "clearplastics_decals"
  ' pClearPlastic4.image = "clearplastics_decals"
  ' pClearPlastic6.image = "clearplastics_decals"
  Else
    pRamp005.image = "clearplastics"
    pRamp006.image = "clearplastics"
  ' pClearPlastic1.image = "clearplastics"
  ' pClearPlastic2.image = "clearplastics"
  ' pClearPlastic3.image = "clearplastics"
  ' pClearPlastic4.image = "clearplastics"
  ' pClearPlastic6.image = "clearplastics"
  End If


  '*************************
  'Flashers global bloom mod
  '*************************
  Dim Flxx
    for each FLxx in FlasherBlooms
        'msgbox "name: " & FLxx.name & " opacity:" & FLxx.Opacity
        FLxx.Opacity = FlasherAdjust
        'msgbox "after name: " & FLxx.name & " opacity:" & FLxx.Opacity
    next
End Sub


'*************************
'Intro music
'*************************
Dim PrevGameOver
Sub PlayMusic
    'Intro Music
  If PrevGameOver = 0 Then
    If IntroMusic = 0 Then
      PlaySound "zIntro"
      PrevGameOver = 1
    End If
  else
'   PrevGameOver = 0
    'StopSound "zIntro"
  End If
End Sub


'**************************************************************************************************************************************************************************************
'                       END OF OPTIONS
'**************************************************************************************************************************************************************************************




'*************************************************************
'Solenoid Call backs
'*************************************************************
SolCallback(sLRFlipper) = "SolRFlipper"
SolCallback(sLLFlipper) = "SolLFlipper"
SolCallback(sURFlipper) = "SolURFlipper"
SolCallback(sULFlipper) = "SolULFlipper"
SolCallback(1) =  "kisort"
SolCallback(2) =  "doflyin:Sol02A"                        'Lane feeder solenoid
SolCallback(3) =  "SolTopHole"
SolCallback(5) =  "Sol05a"                       'Center Right Kicker solenoid
SolCallback(6) =  "vpmSolSound SoundFX(""fx_KnockerLouder"",DOFKnocker),"
SolCallback(7) =  "Sol07A"                        'Right kicker solenoid
SolCallback(10) = "Sol10"                         'Center Left Kicker solenoid
SolCallback(12) = "Sol12_Solenoid"                'Line of death kicker solenoid
SolCallback(13) = "Sol13_Solenoid"                'Outlane kickback solenoid
SolCallback(11) = "SetGI"                         'GI solenoid
SolCallback(16) = "SolRotateBeacons"
SolCallback(21) = "Sol21"                         'Upper diverter solenoid
SolCallback(22) = "Sol22"                         'Lower diverter solenoid

'Flashers
SolModCallback(9)  = "FlashFL109"   '09  top blue
SolModCallback(25) = "FlashFL101"   '01C yagov white
SolModCallback(26) = "FlashFL102"   '02C left lane
SolModCallback(27) = "FlashFL103"   '03C right lane
SolModCallback(28) = "FlashFL104"  'O4C bottom red
SolModCallback(29) = "FlashFL105"   '05C bottom white
SolModCallback(30) = "FlashFL106"   '06C bottom blue
SolModCallback(31) = "FlashFL107"   '07C top white/top red left
SolModCallback(15) = "FlashFL115"   '15  top Right red
SolModCallback(32) = "Flash108"     'radar flasher <<<not sure why 32 works. Service manual says it should be 8.


' NOTE: All vpx flasher objects are now using their "lightmap" property to follow the associalted flasher light Object


Sub FlashFL101(level)
  F101.state = level
  ModFlashFlasher 4,level
  ModFlashFlasher 10,level
  pFiliment01C.BlendDisableLighting = 1000 * level
End Sub

Sub FlashFL102(level)
  F102.state = level
  ModFlashFlasher 8,level
  f102e.state = level
  f102f.state = level
  pFiliment02c.BlendDisableLighting = 1000 * level
End Sub


Sub FlashFL103(level)
  F103.state = level
  ModFlashFlasher 7,level
  f103e.state = level
  f103f.state = level
  pfiliment03c.BlendDisableLighting = 1000 * level
End Sub


Sub FlashFL104(level)
  F104.state = level
  ModFlashFlasher 14,level
  ModFlashFlasher 6,level
  pFiliment04C.BlendDisableLighting = 1000 * level
End Sub



Sub FlashFL105(level)
  F105.state = level
  ModFlashFlasher 13,level
  ModFlashFlasher 5,level
  pFiliment05C.BlendDisableLighting = 1000 * level
End Sub


Sub FlashFL106(level)
  F106.state = level
  ModFlashFlasher 12,level
  ModFlashFlasher 11,level
  pFiliment06C.BlendDisableLighting = 1000 * level
End Sub


Sub FlashFL107(level)
  F107.state = level
  ModFlashFlasher 3,level
  ModFlashFlasher 9,level
  pFiliment07C.BlendDisableLighting = 1000 * level
  pFiliment07Ca.BlendDisableLighting = 1000 * level
End Sub


Sub Flash108(level)
  F108.state = level
  f108a.state = level
  f108c.state = level
  f108d.state = level
End Sub


Sub FlashFL109(level)
  F109.state = level
  ModFlashFlasher 1,level
  pFiliment09.BlendDisableLighting = 1000 * level
End Sub


Sub FlashFL115(level)
  F115.state = level
  ModFlashFlasher 2,level
End Sub




Dim TestFlashers, TableRef, FlasherLightIntensity, FlasherFlareIntensity, FlasherOffBrightness
                ' *********************************************************************
TestFlashers = 0        ' *** set this to 1 to check position of flasher object       ***
Set TableRef = Table1       ' *** change this, if your table has another name             ***
FlasherLightIntensity = 0.005 ' *** lower this, if the VPX lights are too bright (i.e. 0.1)   ***
FlasherFlareIntensity = 0.05  ' *** lower this, if the flares are too bright (i.e. 0.1)     ***
FlasherOffBrightness = 0.2    ' *** brightness of the flasher dome when switched off (range 0-2)  ***
                ' *********************************************************************

Dim ObjLevel(20), objbase(20), objlit(20), objflasher(20), objbloom(20), objlight(20), ObjTargetLevel(20)
Dim tablewidth, tableheight : tablewidth = TableRef.width : tableheight = TableRef.height
'initialise the flasher color, you can only choose from "green", "red", "purple", "blue", "white" and "yellow"
InitFlasher 9, "white" : InitFlasher 1, "blue" : InitFlasher 2, "red"
InitFlasher 3, "red" : InitFlasher 4, "white" : InitFlasher 11, "red"
InitFlasher 5, "red" : InitFlasher 6, "red" : InitFlasher 7, "red"
InitFlasher 8, "red" : InitFlasher 10, "red" : InitFlasher 12, "blue"
InitFlasher 13, "white" : InitFlasher 14, "red"
' rotate the flasher with the command below (first argument = flasher nr, second argument = angle in degrees)
'RotateFlasher 9,-70 : RotateFlasher 1,-100 : RotateFlasher 2,-35
'RotateFlasher 3,8 : RotateFlasher 4,15
'RotateFlasher 12,-37 : RotateFlasher 13,-19 : RotateFlasher 14,-20

Sub InitFlasher(nr, col)
  ' store all objects in an array for use in FlashFlasher subroutine
  Set objbase(nr) = Eval("Flasherbase" & nr)
  Set objlit(nr) = Eval("Flasherlit" & nr)
  Set objflasher(nr) = Eval("Flasherflash" & nr)
  Set objlight(nr) = Eval("Flasherlight" & nr)
  'et objbloom(nr) = Eval("Flasherbloom" & nr)

  ' If the flasher is parallel to the playfield, rotate the VPX flasher object for POV and place it at the correct height
  If objbase(nr).RotY = 0 Then
    objbase(nr).ObjRotZ = Atn( (tablewidth / 2 - objbase(nr).x) / (objbase(nr).y - tableheight * 1.1)) * 180 / 3.14159
    objflasher(nr).RotZ = objbase(nr).ObjRotZ
    objflasher(nr).height = objbase(nr).z + 40
  End If

  ' set all effects to invisible and move the lit primitive at the same position and rotation as the base primitive
  objlight(nr).IntensityScale = 0
  objlit(nr).visible = 0
  objlit(nr).material = "Flashermaterial" & nr
  objlit(nr).RotX = objbase(nr).RotX
  objlit(nr).RotY = objbase(nr).RotY
  objlit(nr).RotZ = objbase(nr).RotZ
  objlit(nr).ObjRotX = objbase(nr).ObjRotX
  objlit(nr).ObjRotY = objbase(nr).ObjRotY
  objlit(nr).ObjRotZ = objbase(nr).ObjRotZ
  objlit(nr).x = objbase(nr).x
  objlit(nr).y = objbase(nr).y
  objlit(nr).z = objbase(nr).z
  objbase(nr).BlendDisableLighting = FlasherOffBrightness

  'rothbauerw
  'Adjust the position of the flasher object to align with the flasher base.
  'Comment out these lines if you want to manually adjust the flasher object
  If objbase(nr).roty > 135 Then
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
  objlight(nr).bulbhaloheight = objbase(nr).z - 10

  'rothbauerw
  'Assign the appropriate bloom image basked on the location of the flasher base
  'Comment out these lines if you want to manually assign the bloom images
  Dim xthird, ythird
  xthird = tablewidth / 3
  ythird = tableheight / 3
' If objbase(nr).x >= xthird And objbase(nr).x <= xthird * 2 Then
'   objbloom(nr).imageA = "flasherbloomCenter"
'   objbloom(nr).imageB = "flasherbloomCenter"
' ElseIf objbase(nr).x < xthird And objbase(nr).y < ythird Then
'   objbloom(nr).imageA = "flasherbloomUpperLeft"
'   objbloom(nr).imageB = "flasherbloomUpperLeft"
' ElseIf  objbase(nr).x > xthird * 2 And objbase(nr).y < ythird Then
'   objbloom(nr).imageA = "flasherbloomUpperRight"
'   objbloom(nr).imageB = "flasherbloomUpperRight"
' ElseIf objbase(nr).x < xthird And objbase(nr).y < ythird * 2 Then
'   objbloom(nr).imageA = "flasherbloomCenterLeft"
'   objbloom(nr).imageB = "flasherbloomCenterLeft"
' ElseIf  objbase(nr).x > xthird * 2 And objbase(nr).y < ythird * 2 Then
'   objbloom(nr).imageA = "flasherbloomCenterRight"
'   objbloom(nr).imageB = "flasherbloomCenterRight"
' ElseIf objbase(nr).x < xthird And objbase(nr).y < ythird * 3 Then
'   objbloom(nr).imageA = "flasherbloomLowerLeft"
'   objbloom(nr).imageB = "flasherbloomLowerLeft"
' ElseIf  objbase(nr).x > xthird * 2 And objbase(nr).y < ythird * 3 Then
'   objbloom(nr).imageA = "flasherbloomLowerRight"
'   objbloom(nr).imageB = "flasherbloomLowerRight"
' End If

  ' set the texture and color of all objects
  Select Case objbase(nr).image
    Case "dome2basewhite"
    objbase(nr).image = "dome2base" & col
    objlit(nr).image = "dome2lit" & col

    Case "ronddomebasewhite"
    objbase(nr).image = "ronddomebase" & col
    objlit(nr).image = "ronddomelit" & col

    Case "domeearbasewhite"
    objbase(nr).image = "domeearbase" & col
    objlit(nr).image = "domeearlit" & col
  End Select
  If TestFlashers = 0 Then
    objflasher(nr).imageA = "domeflashwhite"
    objflasher(nr).visible = 0
  End If
  Select Case col
    Case "blue"
    objlight(nr).color = RGB(4,120,255)
    objflasher(nr).color = RGB(200,255,255)
    'objbloom(nr).color = RGB(4,120,255)
    objlight(nr).intensity = 5000

    Case "green"
    objlight(nr).color = RGB(12,255,4)
    'objflasher(nr).color = RGB(12,255,4)
    objbloom(nr).color = RGB(12,255,4)

    Case "red"
    objlight(nr).color = RGB(255,32,4)
    objflasher(nr).color = RGB(255,32,4)
    'objbloom(nr).color = RGB(255,32,4)

    Case "purple"
    objlight(nr).color = RGB(230,49,255)
    objflasher(nr).color = RGB(255,64,255)
    'objbloom(nr).color = RGB(230,49,255)

    Case "yellow"
    objlight(nr).color = RGB(200,173,25)
    objflasher(nr).color = RGB(255,200,50)
    'objbloom(nr).color = RGB(200,173,25)

    Case "white"
    objlight(nr).color = RGB(255,240,150)
    objflasher(nr).color = RGB(100,86,59)
    'objbloom(nr).color = RGB(255,240,150)

    Case "orange"
    objlight(nr).color = RGB(255,70,0)
    objflasher(nr).color = RGB(255,70,0)
    'objbloom(nr).color = RGB(255,70,0)
  End Select
  objlight(nr).colorfull = objlight(nr).color
  If TableRef.ShowDT And ObjFlasher(nr).RotX =  - 45 Then
    objflasher(nr).height = objflasher(nr).height - 20 * ObjFlasher(nr).y / tableheight
    ObjFlasher(nr).y = ObjFlasher(nr).y + 10
  End If
End Sub

Sub RotateFlasher(nr, angle)
  angle = ((angle + 360 - objbase(nr).ObjRotZ) Mod 180) / 30
  objbase(nr).showframe(angle)
  objlit(nr).showframe(angle)
End Sub


Sub ModFlashFlasher(nr, aValue)
  objflasher(nr).visible = 1  : objlit(nr).visible = 1 ': objbloom(nr).visible = 1
  objflasher(nr).opacity = 1000 *  FlasherFlareIntensity * aValue
  'objbloom(nr).opacity = 100 *  FlasherBloomIntensity * aValue
  objlight(nr).IntensityScale = 0.5 * FlasherLightIntensity * aValue
  objbase(nr).BlendDisableLighting =  FlasherOffBrightness + 10 * aValue
  objlit(nr).BlendDisableLighting = 10 * aValue
  UpdateMaterial "Flashermaterial" & nr,0,0,0,0,0,0,aValue,RGB(255,255,255),0,0,False,True,0,0,0,0
End Sub




F01cBloom.x = 474.583
F01cBloom.y = 970.686

F01cBloom1.x = 474.583
F01cBloom1.y = 970.686

F02cBloom.x = 474.583
F02cBloom.y = 970.686

F02cBloom1.x = 474.583
F02cBloom1.y = 970.686

F03cBloom.x = 474.583
F03cBloom.y = 970.686

F03cBloom1.x = 474.583
F03cBloom1.y = 970.686

F04cBloom.x = 474.583
F04cBloom.y = 970.686

F04cBloom1.x = 474.583
F04cBloom1.y = 970.686

F05cBloom.x = 474.583
F05cBloom.y = 970.686

F05cBloom1.x = 474.583
F05cBloom1.y = 970.686

F06cBloom.x = 474.583
F06cBloom.y = 970.686

F06cBloom1.x = 474.583
F06cBloom1.y = 970.686

F07cBloom.x = 474.583
F07cBloom.y = 970.686

F07cBloom1.x = 474.583
F07cBloom1.y = 970.686

F108Bloom.x = 474.583
F108Bloom.y = 970.686

F09Bloom.x = 474.583
F09Bloom.y = 970.686

F15Bloom.x = 474.583
F15Bloom.y = 970.686

OverheadAmbient.x = 463.2152
OverheadAmbient.y =  1963.054


'**************************************************************************************


Sub Sol21(enabled)
    If enabled Then
    Sol21wall.IsDropped = False
        Diverter1.RotateToEnd
    RandomSoundMetal
    Else
    Sol21wall.IsDropped = True
        Diverter1.RotateToStart
    RandomSoundMetal
    End If
End Sub

Sub Sol22(enabled)
    If enabled Then
    Sol22wall.IsDropped = False
        Diverter2.RotateToEnd
    RandomSoundMetal
    Else
    Sol22wall.IsDropped = True
        Diverter2.RotateToStart
    RandomSoundMetal
    End If
End Sub

Sub DiverterTimer()
  'Diverters
    pDiverter1.objrotz = Diverter1.CurrentAngle + 1
    pDiverter2.objrotz = Diverter2.CurrentAngle + 1
  'Gates
    pGate2Col.RotX = Gate2Col.CurrentAngle *-1 +35
    pGate3Col.RotX = Gate3Col.CurrentAngle *-1 +35
End Sub


'**************

Dim EMPos
EMPos = 80

Sub Sol05a(enabled)
If Enabled then
plungerIM05a.AutoFire:pSol05a.RotX = 120:EMPos = 120:Sol05aMove.Enabled = True
end if
End Sub
Sub Sol05aMove_timer() 'Timer for moving prim finger
  EMPos = EMPos - 1
  pSol05a.RotX = EMPos
  If EMPos < 80 then Sol05aMove.Enabled = 0
End Sub


Sub Sol07A(enabled)
If Enabled then
plungerIM07a.AutoFire:pSol07A.transZ = 120:EMPos = 120:Sol07aMove.Enabled = True
end if
End Sub
Sub Sol07aMove_timer() 'Timer for moving prim finger
  EMPos = EMPos - 1
  pSol07A.transZ = EMPos
  If EMPos < 80 then Sol07aMove.Enabled = 0
End Sub

Sub Sol10(enabled)
If Enabled then
plungerIM10.AutoFire:pSol10.RotX = 120:EMPos = 120:Sol10Move.Enabled = True
end if
End Sub
Sub Sol10Move_timer() 'Timer for moving prim finger
  EMPos = EMPos - 1
  pSol10.RotX = EmPos
  If EMPos < 80 then Sol10Move.Enabled = 0
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
'   ZBBR: BALL BRIGHTNESS
'******************************************************

Const BallBrightness =  1       'Ball brightness - Value between 0 and 1 (0=Dark ... 1=Bright)

' Constants for plunger lane ball darkening.
' You can make a temporary wall in the plunger lane area and use the co-ordinates from the corner control points.
Const PLOffset = 0.5      'Minimum ball brightness scale in plunger lane
Const PLLeft = 1000       'X position of punger lane left
Const PLRight = 1060      'X position of punger lane right
Const PLTop = 1225        'Y position of punger lane top
Const PLBottom = 1900       'Y position of punger lane bottom
Dim PLGain: PLGain = (1-PLOffset)/(PLTop-PLBottom)

Sub UpdateBallBrightness
  Dim s, b_base, b_r, b_g, b_b, d_w
  b_base = 120 * BallBrightness + 135 '*gilvl ' orig was 120 and 70

  For s = 0 To UBound(gBOT)
    ' Handle z direction
    d_w = b_base*(1 - (gBOT(s).z-25)/500)
    If d_w < 30 Then d_w = 30
    ' Handle plunger lane
    If InRect(gBOT(s).x,gBOT(s).y,PLLeft,PLBottom,PLLeft,PLTop,PLRight,PLTop,PLRight,PLBottom) Then
      d_w = d_w*(PLOffset+PLGain*(gBOT(s).y-PLBottom))
    End If
    ' Assign color
    b_r = Int(d_w)
    b_g = Int(d_w)
    b_b = Int(d_w)
    If b_r > 255 Then b_r = 255
    If b_g > 255 Then b_g = 255
    If b_b > 255 Then b_b = 255
    gBOT(s).color = b_r + (b_g * 256) + (b_b * 256 * 256)
    'debug.print "--- ball.color level="&b_r
  Next
End Sub


'****************************
'   ZRBR: Room Brightness
'****************************

'This code only applies to lightmapped tables. It is here for reference.
'NOTE: Objects bightness will be affected by the Day/Night slider only if their blenddisablelighting property is less than 1.
'      Lightmapped table primitives have their blenddisablelighting equal to 1, therefore we need this SetRoomBrightness sub
'      to handle updating their effective ambient brighness.

' Update these arrays if you want to change more materials with room light level
Dim RoomBrightnessMtlArray: RoomBrightnessMtlArray = Array("VLM.Bake.Active","VLM.Bake.Solid")

Sub SetRoomBrightness(lvl)
  If lvl > 1 Then lvl = 1
  If lvl < 0 Then lvl = 0

  ' Lighting level
  Dim v: v=(lvl * 245 + 10)/255

  Dim i: For i = 0 to UBound(RoomBrightnessMtlArray)
    ModulateMaterialBaseColor RoomBrightnessMtlArray(i), i, v
  Next
End Sub

Dim SavedMtlColorArray
SaveMtlColors
Sub SaveMtlColors
  ReDim SavedMtlColorArray(UBound(RoomBrightnessMtlArray))
  Dim i: For i = 0 to UBound(RoomBrightnessMtlArray)
    SaveMaterialBaseColor RoomBrightnessMtlArray(i), i
  Next
End Sub

Sub SaveMaterialBaseColor(name, idx)
  Dim wrapLighting, roughness, glossyImageLerp, thickness, edge, edgeAlpha, opacity, base, glossy, clearcoat, isMetal, opacityActive, elasticity, elasticityFalloff, friction, scatterAngle
  GetMaterial name, wrapLighting, roughness, glossyImageLerp, thickness, edge, edgeAlpha, opacity, base, glossy, clearcoat, isMetal, opacityActive, elasticity, elasticityFalloff, friction, scatterAngle
  SavedMtlColorArray(idx) = round(base,0)
End Sub


Sub ModulateMaterialBaseColor(name, idx, val)
  Dim wrapLighting, roughness, glossyImageLerp, thickness, edge, edgeAlpha, opacity, base, glossy, clearcoat, isMetal, opacityActive, elasticity, elasticityFalloff, friction, scatterAngle
  Dim red, green, blue, saved_base, new_base

  'First get the existing material properties
  GetMaterial name, wrapLighting, roughness, glossyImageLerp, thickness, edge, edgeAlpha, opacity, base, glossy, clearcoat, isMetal, opacityActive, elasticity, elasticityFalloff, friction, scatterAngle

  'Get saved color
  saved_base = SavedMtlColorArray(idx)

  'Next extract the r,g,b values from the base color
  red = saved_base And &HFF
  green = (saved_base \ &H100) And &HFF
  blue = (saved_base \ &H10000) And &HFF
  'msgbox red & " " & green & " " & blue

  'Create new color scaled down by 'val', and update the material
  new_base = RGB(red*val, green*val, blue*val)
  UpdateMaterial name, wrapLighting, roughness, glossyImageLerp, thickness, edge, edgeAlpha, opacity, new_base, glossy, clearcoat, isMetal, opacityActive, elasticity, elasticityFalloff, friction, scatterAngle
End Sub

'******************************************************
' ZKEY: Key Press Handling
'******************************************************

Sub Table1_KeyDown(ByVal KeyCode)
  If Keycode = RightFlipperKey Then FlipperActivate RightFlipper, RFPress
  If Keycode = LeftFlipperKey Then FlipperActivate LeftFlipper, LFPress
  If keycode = PlungerKey Then Plunger.Pullback
  If keycode = RightFlipperKey Then Controller.Switch(63) = 1 ': Objlevel(14) = 1 : FlasherFlash14_Timer
  If keycode = LeftFlipperKey Then Controller.Switch(15) = 1
  If keycode = LeftTiltKey Then Nudge 90, 1 : SoundNudgeLeft      ' Sets the nudge angle and power
  If keycode = RightTiltKey Then Nudge 270, 1 : SoundNudgeRight   ' ^
  If keycode = CenterTiltKey Then Nudge 0, 1 : SoundNudgeCenter   ' ^
  If keycode = StartGameKey Then SoundStartButton: StopSound "zIntro"
  If keycode = AddCreditKey or keycode = AddCreditKey2 Then
    Select Case Int(rnd*3)
      Case 0: PlaySound ("Coin_In_1"), 0, CoinSoundLevel, 0, 0.25
      Case 1: PlaySound ("Coin_In_2"), 0, CoinSoundLevel, 0, 0.25
      Case 2: PlaySound ("Coin_In_3"), 0, CoinSoundLevel, 0, 0.25
    End Select
  End If
  If vpmKeyDown(keycode) Then Exit Sub
End Sub

Sub Table1_KeyUp(ByVal KeyCode)
  If keycode = PlungerKey Then Plunger.Fire
  If keycode = RightFlipperKey Then Controller.Switch(63) = 0
  If keycode = LeftFlipperKey Then Controller.Switch(15) = 0
  If Keycode = RightFlipperKey Then FlipperDeActivate RightFlipper, RFPress
  If Keycode = LeftFlipperKey Then FlipperDeActivate LeftFlipper, LFPress
  If vpmKeyUp(keycode) Then Exit Sub
End Sub

'******************************************************
'           TROUGH
'******************************************************
Sub sw10_hit()
  controller.switch(10) = true
  fgBall = true
  iBall = iBall + 1
  BallsInPlay = BallsInPlay - 1
  RandomSoundDrain sw10
  docrash
End Sub

Sub sw10_UnHit()
  Controller.Switch(10) = 0
End Sub

Sub sw11_Hit():Controller.Switch(11) = 1:UpdateTrough:End Sub
Sub sw11_UnHit():Controller.Switch(11) = 0:UpdateTrough:End Sub
Sub sw12_Hit():Controller.Switch(12) = 1:UpdateTrough:End Sub
Sub sw12_UnHit():Controller.Switch(12) = 0:UpdateTrough:End Sub
Sub sw13_Hit():Controller.Switch(13) = 1:UpdateTrough:End Sub
Sub sw13_UnHit():Controller.Switch(13) = 0:UpdateTrough:End Sub
Sub sw14_Hit():Controller.Switch(14) = 1:UpdateTrough:End Sub
Sub sw14_UnHit():Controller.Switch(14) = 0:UpdateTrough:End Sub

Sub UpdateTrough()
  UpdateTroughTimer.Interval = 300
  UpdateTroughTimer.Enabled = 1
End Sub

Sub UpdateTroughTimer_Timer()
  If sw11.BallCntOver = 0 Then sw12.kick 60, 9
  If sw12.BallCntOver = 0 Then sw13.kick 60, 9
  If sw13.BallCntOver = 0 Then sw14.kick 60, 9
  Me.Enabled = 0
End Sub

dim DontKickAnyMoreBalls

Sub Sol02A(Enabled)
  If DontKickAnyMoreBalls = 0 then
    sw11.Kick 70,12
'   bstatus = 2
    iBall = iBall - 1
    fgBall = false
    UpperGIon = 1
'   Controller.Switch(13)=0
    DontKickAnyMoreBalls = 1
    DKTMstep = 1
    DontKickToMany.enabled = true
    BallsInPlay = BallsInPlay + 1
    RandomSoundBallRelease sw11
  End If
End Sub

Dim DKTMstep

Sub DontKickToMany_timer ()
  Select Case DKTMstep
  Case 1:
  Case 2:
  Case 3: DontKickAnyMoreBalls = 0:DontKickToMany.Enabled = False: DontKickAnyMoreBalls = 0
  End Select
  DKTMstep = DKTMstep + 1
End Sub

sub kisort(enabled)
  if fgBall then
    sw10.Kick 70,20
    iBall = iBall + 1
    fgBall = false
  end if
end sub


'****************************************************************************
'sw24  sol03 Upkicker
'****************************************************************************

Sub sw24_hit()
    SoundSaucerLock
    Controller.Switch(24) = 1
End Sub

Sub sw24_unhit()
    Controller.Switch(24) = 0
End Sub


Sub SolTopHole(Enabled)
    If Enabled Then
        If controller.Switch(24) = True Then
            SoundSaucerKick 1, sw24
        Else
            SoundSaucerKick 0, sw24
        End If
        sw24.Kick 0,50,1.50 '50 = power
        sw24.timerinterval = 10
        sw24.timerenabled = 1
        sw24.uservalue = 1
    End If
End Sub

Dim sw24step
Sub sw24_timer()
  Select Case sw24step
    Case 0:pUpKicker.TransY = 10:SoundSaucerKick 1, pUpKicker
    Case 1:pUpKicker.TransY = 20
    Case 2:pUpKicker.TransY = 30
    Case 3:
    Case 4:
    Case 5:pUpKicker.TransY = 25
    Case 6:pUpKicker.TransY = 20
    Case 7:pUpKicker.TransY = 15: Controller.Switch(24) = 0:sw24.enabled = true
'   Case 7:pUpKicker.TransY = 15: Controller.Switch(24) = 0:HatTrickEnabled = 0:sw24.enabled = true
    Case 8:pUpKicker.TransY = 10
    Case 9:pUpKicker.TransY = 5
    Case 10:pUpKicker.TransY = 0:sw24.timerEnabled = 0:sw24step = 0
  End Select
  sw24step = sw24step + 1
End Sub


'****************************************************************************

Sub sw21_Hit:Controller.Switch(21) = 1:End Sub
Sub sw21_UnHit:Controller.Switch(21) = 0:End Sub
Sub sw22_Hit:Controller.Switch(22) = 1:End Sub
Sub sw22_UnHit:Controller.Switch(22) = 0:End Sub
Sub sw23_Hit:Controller.Switch(23) = 1:End Sub
Sub sw23_UnHit:Controller.Switch(23) = 0:End Sub

Sub sw25_Hit:vpmTimer.PulseSw(25):End Sub
Sub sw26_Hit:vpmTimer.PulseSw(26):End Sub

Sub Bumper1_Hit:vpmTimer.PulseSw(28) : RandomSoundBumperTop Bumper1 : End Sub

Sub sw33_Hit:vpmTimer.PulseSw(33):End Sub
Sub sw34_Hit:vpmTimer.PulseSw(34):End Sub
Sub sw35_Hit:vpmTimer.PulseSw(35):End Sub
Sub sw36_Hit:vpmTimer.PulseSw(36):End Sub
Sub sw37_Hit:vpmTimer.PulseSw(37):End Sub
Sub sw38_Hit:vpmTimer.PulseSw(38):End Sub
Sub sw41_Hit:vpmTimer.PulseSw(41):End Sub
Sub sw42_Hit:vpmTimer.PulseSw(42):End Sub
Sub sw43_Hit:vpmTimer.PulseSw(43):End Sub
Sub sw44_Hit:vpmTimer.PulseSw(44):End Sub
Sub sw45_Hit:vpmTimer.PulseSw(45):End Sub
Sub sw46_Hit:vpmTimer.PulseSw(46):End Sub
Sub sw49_Hit:vpmTimer.PulseSw(49):End Sub
Sub sw50_Hit:vpmTimer.PulseSw(50):End Sub

Sub sw51_Hit:vpmTimer.PulseSw(51):End Sub
Sub sw52_Hit:vpmTimer.PulseSw(52):End Sub
Sub sw53_Hit:vpmTimer.PulseSw(53):End Sub
Sub sw54_Hit:vpmTimer.PulseSw(54):End Sub

'**************************************************************
'sw55 Sol12 Line Of Death Kickback
'**************************************************************

Sub sw55_Hit()
  Me.TimerEnabled = true
  Controller.Switch(55) = 1
End Sub
Sub sw55_unHit()
  Me.TimerEnabled = true
  Controller.Switch(55) = 0
End Sub
Sub Sol12_Solenoid(Enabled)
  If Enabled Then
    LODKick.Enabled=True
  Else
    LODKick.Enabled=False
  End If
End Sub
Sub LODKick_Hit:Me.Kick 180,YagovKick:Me.TimerEnabled = true:End Sub
'**************************************************************

' Wire Triggers
Sub sw16_Hit:Controller.Switch(16) = 1 : End Sub
Sub sw16_Unhit:Controller.Switch(16) = 0:End Sub
Sub sw59_Hit:Controller.Switch(59) = 1 : End Sub
Sub sw59_UnHit:Controller.Switch(59) = 0:End Sub
Sub sw60_Hit:Controller.Switch(60) = 1 : End Sub
Sub sw60_UnHit:Controller.Switch(60) = 0:End Sub


'**************************************************************
'sw61 Sol13 Outlane Kickback
'**************************************************************
Sub sw61_Hit()
  'Switch61dir = 1
  'Sw61Move = 1
  Me.TimerEnabled = true
  Controller.Switch(61) = 1
End Sub

Sub sw61_unHit()
  'Switch61dir = -1
  'Sw61Move = 5
  Me.TimerEnabled = true
  Controller.Switch(61) = 0
End Sub

Dim LKStep
Sub Sol13_Solenoid(Enabled)
  If Enabled Then
    LaserKick.Enabled=True
    SoundPlungerReleaseBall
  Else
    LaserKick.Enabled=False
  End If
End Sub

Sub LaserKick_Hit:Me.Kick 0,50:LKStep = 0:pLaserKick.TransY = 75:Me.TimerEnabled = true:End Sub  '50=strength

Sub LaserKick_timer()
  Select Case LKStep
    Case 0:pLaserKick.TransY = 50
    Case 1:pLaserKick.TransY = 25
    Case 2:pLaserKick.TransY = 0:me.TimerEnabled = false
  End Select
  LKStep = LKStep + 1
End Sub
'**************************************************************

Sub sw62_Hit:Controller.Switch(62) = 1 : End Sub
Sub sw62_UnHit:Controller.Switch(62) = 0:End Sub

'Ramp Triggers
Sub sw20_Hit:Controller.Switch(20) = 1 : End Sub
Sub sw20_Unhit:Controller.Switch(20) = 0:End Sub
Sub sw30_Hit:Controller.Switch(30) = 1 : WireRampOn False : End Sub
Sub sw30_Unhit:Controller.Switch(30) = 0:End Sub
Sub sw31_Hit:Controller.Switch(31) = 1 : WireRampOn False : End Sub
Sub sw31_Unhit:Controller.Switch(31) = 0:End Sub
Sub sw32_Hit:Controller.Switch(32) = 1 : WireRampOn False : End Sub
Sub sw32_Unhit:Controller.Switch(32) = 0:End Sub

Sub Trigger4_hit : End Sub

'Gate Triggers
Sub sw47_Hit:vpmTimer.PulseSw(47):End Sub
Sub sw56_Hit:vpmTimer.PulseSw(56):End Sub

'Spinners
Sub sw48_Spin:vpmTimer.PulseSw 48 : End Sub




'*************************
' Modulated 3D insert stuff
'*************************

Sub L1_animate: p1.BlendDisableLighting = 200 * (L1.GetInPlayIntensity / L1.Intensity): End Sub
Sub L2_animate: p2.BlendDisableLighting = 200 * (L2.GetInPlayIntensity / L2.Intensity): End Sub
Sub L3_animate: p3.BlendDisableLighting = 200 * (L3.GetInPlayIntensity / L3.Intensity): End Sub
Sub L4_animate: p4.BlendDisableLighting = 200 * (L4.GetInPlayIntensity / L4.Intensity): End Sub
Sub L5_animate: p5.BlendDisableLighting = 200 * (L5.GetInPlayIntensity / L5.Intensity): End Sub
Sub L6_animate: p6.BlendDisableLighting = 200 * (L6.GetInPlayIntensity / L6.Intensity): End Sub
Sub L7_animate: p7.BlendDisableLighting = 200 * (L7.GetInPlayIntensity / L7.Intensity): End Sub
Sub L8_animate: p8.BlendDisableLighting = 200 * (L8.GetInPlayIntensity / L8.Intensity): End Sub
Sub L9_animate: p9.BlendDisableLighting = 200 * (L9.GetInPlayIntensity / L9.Intensity): End Sub

Sub L10_animate: p10.BlendDisableLighting = 200 * (L10.GetInPlayIntensity / L10.Intensity): End Sub
Sub L11_animate: p11.BlendDisableLighting = 200 * (L11.GetInPlayIntensity / L11.Intensity): End Sub
Sub L12_animate: p12.BlendDisableLighting = 200 * (L12.GetInPlayIntensity / L12.Intensity): End Sub
Sub L13_animate: p13.BlendDisableLighting = 200 * (L13.GetInPlayIntensity / L13.Intensity): End Sub
Sub L14_animate: p14.BlendDisableLighting = 200 * (L14.GetInPlayIntensity / L14.Intensity): End Sub
Sub L15_animate: p15.BlendDisableLighting = 200 * (L15.GetInPlayIntensity / L15.Intensity): End Sub
Sub L16_animate: p16.BlendDisableLighting = 200 * (L16.GetInPlayIntensity / L16.Intensity): End Sub
Sub L17_animate: p17.BlendDisableLighting = 200 * (L17.GetInPlayIntensity / L17.Intensity): End Sub
Sub L18_animate: p18.BlendDisableLighting = 200 * (L18.GetInPlayIntensity / L18.Intensity): End Sub
Sub L19_animate: p19.BlendDisableLighting = 200 * (L19.GetInPlayIntensity / L19.Intensity): End Sub

Sub L20_animate: p20.BlendDisableLighting = 200 * (L20.GetInPlayIntensity / L20.Intensity): End Sub
Sub L21_animate: p21.BlendDisableLighting = 200 * (L21.GetInPlayIntensity / L21.Intensity): End Sub
Sub L22_animate: p22.BlendDisableLighting = 200 * (L22.GetInPlayIntensity / L22.Intensity): End Sub
Sub L23_animate: p23.BlendDisableLighting = 200 * (L23.GetInPlayIntensity / L23.Intensity): End Sub
Sub L24_animate: p24.BlendDisableLighting = 200 * (L24.GetInPlayIntensity / L24.Intensity): p24a.BlendDisableLighting = p24.BlendDisableLighting: End Sub
Sub L25_animate: p25.BlendDisableLighting = 200 * (L25.GetInPlayIntensity / L25.Intensity): End Sub
Sub L26_animate: p26.BlendDisableLighting = 200 * (L26.GetInPlayIntensity / L26.Intensity): End Sub
Sub L27_animate: p27.BlendDisableLighting = 200 * (L27.GetInPlayIntensity / L27.Intensity): End Sub
Sub L28_animate: p28.BlendDisableLighting = 200 * (L28.GetInPlayIntensity / L28.Intensity): End Sub
Sub L29_animate: p29.BlendDisableLighting = 200 * (L29.GetInPlayIntensity / L29.Intensity): End Sub

Sub L30_animate: p30.BlendDisableLighting = 200 * (L30.GetInPlayIntensity / L30.Intensity): End Sub
Sub L31_animate: p31.BlendDisableLighting = 200 * (L31.GetInPlayIntensity / L31.Intensity): End Sub
Sub L32_animate: p32.BlendDisableLighting = 200 * (L32.GetInPlayIntensity / L32.Intensity): End Sub
Sub L33_animate: p33.BlendDisableLighting = 200 * (L33.GetInPlayIntensity / L33.Intensity): End Sub
Sub L34_animate: p34.BlendDisableLighting = 200 * (L34.GetInPlayIntensity / L34.Intensity): End Sub
Sub L35_animate: p35.BlendDisableLighting = 200 * (L35.GetInPlayIntensity / L35.Intensity): End Sub
Sub L36_animate: p36.BlendDisableLighting = 200 * (L36.GetInPlayIntensity / L36.Intensity): End Sub
Sub L37_animate: p37.BlendDisableLighting = 200 * (L37.GetInPlayIntensity / L37.Intensity): End Sub
Sub L38_animate: p38.BlendDisableLighting = 200 * (L38.GetInPlayIntensity / L38.Intensity): End Sub
Sub L39_animate: p39.BlendDisableLighting = 200 * (L39.GetInPlayIntensity / L39.Intensity): p39a.BlendDisableLighting = p39.BlendDisableLighting: End Sub

Sub L40_animate: p40.BlendDisableLighting = 200 * (L40.GetInPlayIntensity / L40.Intensity): End Sub
Sub L41_animate: p41.BlendDisableLighting = 200 * (L41.GetInPlayIntensity / L41.Intensity): End Sub
Sub L42_animate: p42.BlendDisableLighting = 200 * (L42.GetInPlayIntensity / L42.Intensity): End Sub
Sub L43_animate: p43.BlendDisableLighting = 200 * (L43.GetInPlayIntensity / L43.Intensity): End Sub
Sub L44_animate: p44.BlendDisableLighting = 200 * (L44.GetInPlayIntensity / L44.Intensity): End Sub
Sub L45_animate: p45.BlendDisableLighting = 200 * (L45.GetInPlayIntensity / L45.Intensity): End Sub
Sub L46_animate: p46.BlendDisableLighting = 200 * (L46.GetInPlayIntensity / L46.Intensity): End Sub
Sub L47_animate: p47.BlendDisableLighting = 200 * (L47.GetInPlayIntensity / L47.Intensity): End Sub
Sub L48_animate: p48.BlendDisableLighting = 200 * (L48.GetInPlayIntensity / L48.Intensity): End Sub
Sub L49_animate: p49.BlendDisableLighting = 200 * (L49.GetInPlayIntensity / L49.Intensity): End Sub

Sub L50_animate: p50.BlendDisableLighting = 200 * (L50.GetInPlayIntensity / L50.Intensity): End Sub
Sub L51_animate: p51.BlendDisableLighting = 200 * (L51.GetInPlayIntensity / L51.Intensity): End Sub
Sub L52_animate: p52.BlendDisableLighting = 200 * (L52.GetInPlayIntensity / L52.Intensity): End Sub
Sub L53_animate: p53.BlendDisableLighting = 200 * (L53.GetInPlayIntensity / L53.Intensity): End Sub
Sub L54_animate: p54.BlendDisableLighting = 200 * (L54.GetInPlayIntensity / L54.Intensity): End Sub
Sub L55_animate: p55.BlendDisableLighting = 200 * (L55.GetInPlayIntensity / L55.Intensity): End Sub
Sub L56_animate: p56.BlendDisableLighting = 200 * (L56.GetInPlayIntensity / L56.Intensity): End Sub

Sub L57_animate
  dim p: p = (L57.GetInPlayIntensity / L57.Intensity)
  pfil57.BlendDisableLighting = 15 * p
  pBulb57.BlendDisableLighting = 0.1 * p
End Sub
Sub L58_animate
  dim p: p = (L58.GetInPlayIntensity / L58.Intensity)
  pfil58.BlendDisableLighting = 15 * p
  pBulb58.BlendDisableLighting = 0.1 * p
End Sub
Sub L59_animate
  dim p: p = (L59.GetInPlayIntensity / L59.Intensity)
  pfil59.BlendDisableLighting = 15 * p
  pBulb59.BlendDisableLighting = 0.1 * p
End Sub
Sub L60_animate
  dim p: p = (L60.GetInPlayIntensity / L60.Intensity)
  pfil60.BlendDisableLighting = 15 * p
  pBulb60.BlendDisableLighting = 0.1 * p
End Sub
Sub L61_animate
  dim p: p = (L61.GetInPlayIntensity / L61.Intensity)
  pfil61.BlendDisableLighting = 15 * p
  pBulb61.BlendDisableLighting = 0.1 * p
End Sub
Sub L62_animate
  dim p: p = (L62.GetInPlayIntensity / L62.Intensity)
  pfil62.BlendDisableLighting = 15 * p
  pBulb62.BlendDisableLighting = 0.1 * p
End Sub

Sub L63_animate: p63.BlendDisableLighting = 200 * (L63.GetInPlayIntensity / L63.Intensity): End Sub
Sub L64_animate: p64.BlendDisableLighting = 200 * (L64.GetInPlayIntensity / L64.Intensity): End Sub
'Sub L65_animate: p65.BlendDisableLighting = 200 * (L65.GetInPlayIntensity / L65.Intensity): End Sub

'Sub L70_animate: debug.print "L70_animate "&(L70.GetInPlayIntensity / L70.Intensity): End Sub

Sub F108_animate: pF108.BlendDisableLighting = 30 * (F108.GetInPlayIntensity / F108.Intensity) + 1.5*gilvl: End Sub

'*************************
'System 11 GI On/Off
'*************************
dim gilvl: gilvl=0
Sub GIOn  : SetGI 0: End Sub 'These are just debug commands now
Sub GIOff : SetGI 1 : End Sub

Sub SetGI(aOn)
  Dim bulb
  For Each bulb In GILighting
    bulb.state = aOn
  Next
  If CustomBackwall = 1 then
    nimlight.state = aOn
  End If
  Select Case aOn
    Case False: Playsound "fx_FlashRelayOff"
    Case true: Playsound "fx_FlashRelayOn"
  End Select
End Sub

Sub GI001_animate
  Dim bulb
  dim p: p = (GI001.GetInPlayIntensity / GI001.Intensity)
  For Each bulb In GIBulbs
    bulb.blenddisablelighting = p*2
  Next
  FlFadeBumper 1,p -0.3
  gilvl = p
  pF108.BlendDisableLighting = 1.5*p
  pRamp005.blenddisablelighting = 0.1*p
  pRamp006.blenddisablelighting = 0.1*p
  pClearPlastic001.blenddisablelighting = 0.1*p
End Sub



'*********************************************************************************************************************************************************
'End lamp helper functions
'*********************************************************************************************************************************************************


'*********************************************************************************************************************************************************
'Digital Display
'*********************************************************************************************************************************************************
 Dim Digits(28)
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

 ' 3rd Player
Digits(14) = Array(LED150,LED151,LED152,LED153,LED154,LED155,LED156)
Digits(15) = Array(LED160,LED161,LED162,LED163,LED164,LED165,LED166)
Digits(16) = Array(LED170,LED171,LED172,LED173,LED174,LED175,LED176)
Digits(17) = Array(LED180,LED181,LED182,LED183,LED184,LED185,LED186)
Digits(18) = Array(LED190,LED191,LED192,LED193,LED194,LED195,LED196)
Digits(19) = Array(LED200,LED201,LED202,LED203,LED204,LED205,LED206)
Digits(20) = Array(LED210,LED211,LED212,LED213,LED214,LED215,LED216)

' 4th Player
Digits(21) = Array(LED220,LED221,LED222,LED223,LED224,LED225,LED226)
Digits(22) = Array(LED230,LED231,LED232,LED233,LED234,LED235,LED236)
Digits(23) = Array(LED240,LED241,LED242,LED243,LED244,LED245,LED246)
Digits(24) = Array(LED250,LED251,LED252,LED253,LED254,LED255,LED256)
Digits(25) = Array(LED260,LED261,LED262,LED263,LED264,LED265,LED266)
Digits(26) = Array(LED270,LED271,LED272,LED273,LED274,LED275,LED276)
Digits(27) = Array(LED280,LED281,LED282,LED283,LED284,LED285,LED286)


Sub DisplayTimer
    Dim ChgLED, ii, jj, num, chg, stat, obj, b, x
    ChgLED=Controller.ChangedLEDs(&Hffffffff, &Hffffffff)
    If Not IsEmpty(ChgLED)Then
       For ii=0 To UBound(chgLED)
          num=chgLED(ii, 0) : chg=chgLED(ii, 1) : stat=chgLED(ii, 2)
      if (num < 28) then
        if VRRoom > 0 Then
        For Each obj In DigitsVR(num)
                   If chg And 1 Then FadeDisplay obj, stat And 1
                   chg=chg\2 : stat=stat\2
                  Next
        end If
        if DesktopMode = True and VRRoom < 1 Then
        For Each obj In Digits(num)
                   If chg And 1 Then obj.State=stat And 1
                   chg=chg\2 : stat=stat\2
                  Next
        end If
        else
      end if
    Next
    End If
 End Sub

' *********************************************************************
' *********************************************************************

          'Start of VPX call back Functions

' *********************************************************************
' *********************************************************************

'**********Sling Shot Animations
' Rstep and Lstep  are the variables that increment the animation
'****************
Dim RStep, Lstep, Cstep

Sub RightSlingShot_Slingshot
  RS.VelocityCorrect(Activeball)
  vpmTimer.PulseSw 58
    RSling.Visible = 0
    RSling1.Visible = 1
    sling1.TransZ = -30
    RStep = 0
    RightSlingShot.TimerEnabled = 1
  RandomSoundSlingshotRight SLING1
End Sub

Sub RightSlingShot_Timer
    Select Case RStep
        Case 3:RSLing1.Visible = 0:RSLing2.Visible = 1:sling1.TransZ = -10:
        'Case 4:RSLing2.Visible = 0:RSLing.Visible = 1:sling1.TransZ = 0:RightSlingShot.TimerEnabled = 0:f127a.state=0:
        Case 4:RSLing2.Visible = 0:RSLing.Visible = 1:sling1.TransZ = 0:RightSlingShot.TimerEnabled = 0:
    End Select
    RStep = RStep + 1
End Sub

Sub LeftSlingShot_Slingshot
  LS.VelocityCorrect(Activeball)
  vpmTimer.PulseSw 57
    LSling.Visible = 0
    LSling1.Visible = 1
    sling2.TransZ = -30
    LStep = 0
    LeftSlingShot.TimerEnabled = 1
  RandomSoundSlingshotLeft SLING2
End Sub

Sub LeftSlingShot_Timer
    Select Case LStep
        Case 3:LSLing1.Visible = 0:LSLing2.Visible = 1:sling2.TransZ = -10:
        'Case 4:LSLing2.Visible = 0:LSLing.Visible = 1:sling2.TransZ = 0:LeftSlingShot.TimerEnabled = 0:f126a.state=0:
        Case 4:LSLing2.Visible = 0:LSLing.Visible = 1:sling2.TransZ = 0:LeftSlingShot.TimerEnabled = 0:
    End Select
    LStep = LStep + 1
End Sub

Sub Band3_Hit
  RubberTop.visible = 0
    RubberTop1.Visible = 0
    RubberTop2.Visible = 1
    CStep = 0
    Band3.TimerEnabled = 1
End Sub

Sub Band3_Timer
    Select Case CStep
        Case 3:RubberTop1.Visible = 0:RubberTop2.Visible = 1:
        Case 4:RubberTop2.Visible = 0:RubberTop.Visible = 1:Band3.TimerEnabled = 0:
    End Select
    CStep = CStep + 1
End Sub


'******************************************************
' ZSSC: SLINGSHOT CORRECTION FUNCTIONS by apophis
'******************************************************
' To add these slingshot corrections:
'  - On the table, add the endpoint primitives that define the two ends of the Slingshot
'  - Initialize the SlingshotCorrection objects in InitSlingCorrection
'  - Call the .VelocityCorrect methods from the respective _Slingshot event sub

Dim LS: Set LS = New SlingshotCorrection
Dim RS: Set RS = New SlingshotCorrection
Dim TS: Set TS = New SlingshotCorrection

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
  a = Array(LS, RS, TS)
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


'************************************************* VR Changes

'Display Output

dim DisplayColor, DisplayColorG
DisplayColor =  RGB(255,40,1)

Sub FadeDisplay(object, onoff)
  If OnOff = 1 Then
    object.color = DisplayColor
  Else
    Object.Color = RGB(1,1,1)
  End If
End Sub

Dim DigitsVR(32)
DigitsVR(0)=Array(D1,D2,D3,D4,D5,D6,D7,D8,D9,D10,D11,D12,D13,D14,D15)
DigitsVR(1)=Array(D16,D17,D18,D19,D20,D21,D22,D23,D24,D25,D26,D27,D28,D29,D30)
DigitsVR(2)=Array(D31,D32,D33,D34,D35,D36,D37,D38,D39,D40,D41,D42,D43,D44,D45)
DigitsVR(3)=Array(D46,D47,D48,D49,D50,D51,D52,D53,D54,D55,D56,D57,D58,D59,D60)
DigitsVR(4)=Array(D61,D62,D63,D64,D65,D66,D67,D68,D69,D70,D71,D72,D73,D74,D75)
DigitsVR(5)=Array(D76,D77,D78,D79,D80,D81,D82,D83,D84,D85,D86,D87,D88,D89,D90)
DigitsVR(6)=Array(D91,D92,D93,D94,D95,D96,D97,D98,D99,D100,D101,D102,D103,D104,D105)
DigitsVR(7)=Array(D106,D107,D108,D109,D110,D111,D112,D113,D114,D115,D116,D117,D118,D119,D120)
DigitsVR(8)=Array(D121,D122,D123,D124,D125,D126,D127,D128,D129,D130,D131,D132,D133,D134,D135)
DigitsVR(9)=Array(D136,D137,D138,D139,D140,D141,D142,D143,D144,D145,D146,D147,D148,D149,D150)
DigitsVR(10)=Array(D151,D152,D153,D154,D155,D156,D157,D158,D159,D160,D161,D162,D163,D164,D165)
DigitsVR(11)=Array(D166,D167,D168,D169,D170,D171,D172,D173,D174,D175,D176,D177,D178,D179,D180)
DigitsVR(12)=Array(D181,D182,D183,D184,D185,D186,D187,D188,D189,D190,D191,D192,D193,D194,D195)
DigitsVR(13)=Array(D196,D197,D198,D199,D200,D201,D202,D203,D204,D205,D206,D207,D208,D209,D210)
DigitsVR(14)=Array(D211,D212,D213,D214,D215,D216,D217,D218)
DigitsVR(15)=Array(D219,D220,D221,D222,D223,D224,D225,D226)
DigitsVR(16)=Array(D227,D228,D229,D230,D231,D232,D233,D234)
DigitsVR(17)=Array(D235,D236,D237,D238,D239,D240,D241,D242)
DigitsVR(18)=Array(D243,D244,D245,D246,D247,D248,D249,D250)
DigitsVR(19)=Array(D251,D252,D253,D254,D255,D256,D257,D258)
DigitsVR(20)=Array(D259,D260,D261,D262,D263,D264,D265,D266)
DigitsVR(21)=Array(D267,D268,D269,D270,D271,D272,D273,D274)
DigitsVR(22)=Array(D275,D276,D277,D278,D279,D280,D281,D282)
DigitsVR(23)=Array(D283,D284,D285,D286,D287,D288,D289,D290)
DigitsVR(24)=Array(D291,D292,D293,D294,D295,D296,D297,D298)
DigitsVR(25)=Array(D299,D300,D301,D302,D303,D304,D305,D306)
DigitsVR(26)=Array(D307,D308,D309,D310,D311,D312,D313,D314)
DigitsVR(27)=Array(D315,D316,D317,D318,D319,D320,D321,D322)


Sub InitDigits()
  dim tmp, x, obj
  for x = 0 to uBound(DigitsVR)
    if IsArray(DigitsVR(x) ) then
      For each obj in DigitsVR(x)
        obj.height = obj.height + 18
        FadeDisplay obj, 0
      next
    end If
  Next
End Sub

InitDigits


'*************************************************


Dim BeaconPos:BeaconPos = 0

Sub BeaconTimer_Timer
  BeaconPos = BeaconPos + 3
  if BeaconPos = 360 then BeaconPos = 0

  BeaconRedInt.RotY = BeaconPos
    BeaconRed.BlendDisableLighting=.3 * abs(sin((BeaconPos+90) * 6.28 / 360))
  BeaconFR.RotY = BeaconPos
    if BeaconPos < 90 or BeaconPos > 270 then BeaconFr.IntensityScale = -.05  else BeaconFr.IntensityScale = 1

  BeaconWhiteInt.RotY = BeaconPos+45
  BeaconWhite.BlendDisableLighting=.3 * abs(sin((BeaconPos+90+45) * 6.28 / 360))
  BeaconFW.RotY = BeaconPos + 45
    if BeaconPos + 45 < 90 or BeaconPos + 45 > 270 then BeaconFW.IntensityScale = -.05 else BeaconFW.IntensityScale = 1

  BeaconBlueInt.RotY = BeaconPos+90
    BeaconBlue.BlendDisableLighting=.3 * abs(sin((BeaconPos+90+90) * 6.28 / 360))
  BeaconFB.RotY = BeaconPos + 90
    if BeaconPos+90 > 270 then BeaconFb.IntensityScale = -.05 else BeaconFb.IntensityScale = 1


End Sub

Sub SolRotateBeacons(Enabled)
    If Enabled then
        BeaconRed.image = "dome3_red_lit"
        BeaconWhite.image = "dome3_clear_lit"
        BeaconBlue.image = "dome3_blue_lit"
    BeaconTimer.Enabled = true
    if ModelToys > 0 Then
      flame1.visible = 1
      flame2.visible = 1
      flame3.visible = 1
      flame4.visible = 1
      flame5.visible = 1
      flame6.visible = 1
      flame7.visible = 1
      flame8.visible = 1
    end if
    MoveWingsDir = -1
    MoveWings.Enabled = true

    Else
    MoveWingsDir = 1
    MoveWings.Enabled = true

    BeaconTimer.Enabled = false
        BeaconRed.image = "dome3_red"
        BeaconWhite.image = "dome3_clear"
        BeaconBlue.image = "dome3_blue"

    if BeaconPos < 90 or BeaconPos > 270 then BeaconFr.IntensityScale = -.01  else BeaconFr.IntensityScale = .01
    if BeaconPos + 45 < 90 or BeaconPos + 45 > 270 then BeaconFW.IntensityScale = -.01 else BeaconFW.IntensityScale = .01
    if BeaconPos+90 > 270 then BeaconFb.IntensityScale = -.01 else BeaconFb.IntensityScale = .01

    flame1.visible = 0
    flame2.visible = 0
    flame3.visible = 0
    flame4.visible = 0
    flame5.visible = 0
    flame6.visible = 0
    flame7.visible = 0
    flame8.visible = 0
    End If
End Sub

SolRotateBeacons False


'*********************************************************


dim f14ar:f14ar=0
dim f14br:f14br=0
dim f14r:f14r=0
dim leftrotate:leftrotate = 0
dim rightrotate:rightrotate = 0
dim crash:crash =0

Sub docrash
  if crash = 0 and VRRoom =1 then
    crash = 1
    if leftrotate = 0 then leftrotate =1
    if rightrotate = 0 then rightrotate =1
  end if
end sub

dim flyin:flyin = 0

sub doflyin
  if flyin =0 and VRRoom =1 then
    f14a.visible = True
    f14b.visible = true
    crash = 0
    flyin = 1
    UpdateLamps
  end if
end sub

dim f14ax:f14ax = f14a.x
dim f14bx:f14bx = f14b.x
dim f14z:f14z = 400
dim f14y:f14y = 800

Sub UpdateLamps
  f14ar = f14ar + (rnd * 2)
  f14br = f14br + (rnd * 2)
  if rnd < .00005 and rightrotate = 0 then rightrotate = 1 else if rightrotate > 0 then rightrotate = rightrotate + 1:if rightrotate > 360 then rightrotate =0
  if rnd < .00005 and leftrotate = 0 then leftrotate = 1 else if leftrotate > 0 then leftrotate = leftrotate + 1:if leftrotate > 360 then leftrotate =0


  if f14ar > 360 then f14ar = f14ar - 360
  if f14br > 360 then f14br = f14br - 360

  f14a.transy = sin(f14ar * 6.28 / 360) * 5
  f14a.transz = sin(f14br * 6.28 / 360) * 5
  f14b.transy = sin(f14ar * 6.28 / 360) * 5
  f14b.transz = sin(f14br * 6.28 / 360) * 5

  f14a.objroty  = leftrotate
  f14b.objroty  = rightrotate

  if crash > 0 Then
    crash = crash + 1
    f14a.z = f14z - crash * 10
    f14b.z = f14z - crash * 10
    f14a.y = f14y - crash * 5
    f14b.y = f14y - crash * 5

    f14a.rotx = crash * 90 /1000
    f14b.rotx = crash * 90 / 1000
    f14a.x = f14ax - crash * 10
    f14b.x = f14bx + crash * 10
    if leftrotate = 0 then leftrotate =1
    if rightrotate = 0 then rightrotate =1

    if crash > 3000 then
      f14a.visible = false
      f14b.visible = false
      crash = 0
    end if
  end if

  if flyin > 0 then
    flyin = flyin + 1
    f14a.transx = 0
    f14b.transx = 0
    f14a.rotx =0:f14b.rotx = 0
    f14a.z = f14z
    f14b.z = f14z
    f14a.y = f14y + (.5 * (150-flyin) ^ 2 )
    f14b.y = f14y + (.5 * (150-flyin) ^ 2 )
    f14a.x = f14ax
    f14b.x = f14bx
    if flyin=150 then flyin = 0
  end if
End Sub

Sub SetupVRRoom
  DIM VRThings
  If VRRoom > 0 Then
    If VRRoom = 1 Then
      for each VRThings in VRPlanes:VRThings.visible = 1:Next
      for each VRThings in VRCab:VRThings.visible = 1:Next
      for each VRThings in VRStuff:VRThings.visible = 0:Next
      Sidewalls.visible = 0
      OverheadAmbient.visible = 0
    End If
    If VRRoom = 2 Then
      for each VRThings in VRPlanes:VRThings.visible = 0:Next
      for each VRThings in VRCab:VRThings.visible = 1:Next
      for each VRThings in VRStuff:VRThings.visible = 1:Next
      Sidewalls.visible = 0
      OverheadAmbient.visible = 0
    End If
    If VRRoom = 3 Then
      for each VRThings in VRPlanes:VRThings.visible = 0:Next
      for each VRThings in VRCab:VRThings.visible = 0:Next
      for each VRThings in VRStuff:VRThings.visible = 0:Next
      PinCab_Backglass.visible = 1
      Sidewalls.visible = 0
      OverheadAmbient.visible = 0
      VR_OverheadAmbient.visible = 1
    End If
  Else
      for each VRThings in VRPlanes:VRThings.visible = 0:Next
      for each VRThings in VRCab:VRThings.visible = 0:Next
      for each VRThings in VRStuff:VRThings.visible = 0:Next
      Sidewalls.visible = 1
      OverheadAmbient.visible = 1
  End if
End Sub


'*******  Set Up Backglass  ***********************

Sub SetBackglass()
  Dim obj

  For Each obj In BackGlass
    obj.x = obj.x
    obj.height = - obj.y + 490
    obj.y = -140
  Next
End Sub




Dim MoveWingsStep, MoveWingsDir
Sub MoveWings_timer()
  Select case MoveWingsStep
    Case 0:Model1WingLeft.RotZ = 0:Model1WingRight.RotZ = 0:Model2WingLeft.RotZ = -0:Model2WingRight.RotZ = 0: if MoveWingsDir < 0 then me.Enabled = false:exit sub
    Case 1:Model1WingLeft.RotZ = -1:Model1WingRight.RotZ = 1:Model2WingLeft.RotZ = -1:Model2WingRight.RotZ = 1
    Case 2:Model1WingLeft.RotZ = -2:Model1WingRight.RotZ = 2:Model2WingLeft.RotZ = -2:Model2WingRight.RotZ = 2
    Case 3:Model1WingLeft.RotZ = -3:Model1WingRight.RotZ = 3:Model2WingLeft.RotZ = -3:Model2WingRight.RotZ = 3
    Case 4:Model1WingLeft.RotZ = -4:Model1WingRight.RotZ = 4:Model2WingLeft.RotZ = -4:Model2WingRight.RotZ = 4
    Case 5:Model1WingLeft.RotZ = -5:Model1WingRight.RotZ = 5:Model2WingLeft.RotZ = -5:Model2WingRight.RotZ = 5
    Case 6:Model1WingLeft.RotZ = -6:Model1WingRight.RotZ = 6:Model2WingLeft.RotZ = -6:Model2WingRight.RotZ = 6
    Case 7:Model1WingLeft.RotZ = -7:Model1WingRight.RotZ = 7:Model2WingLeft.RotZ = -7:Model2WingRight.RotZ = 7
    Case 8:Model1WingLeft.RotZ = -8:Model1WingRight.RotZ = 8:Model2WingLeft.RotZ = -8:Model2WingRight.RotZ = 8
    Case 9:Model1WingLeft.RotZ = -9:Model1WingRight.RotZ = 9:Model2WingLeft.RotZ = -9:Model2WingRight.RotZ = 9
    Case 10:Model1WingLeft.RotZ = -10:Model1WingRight.RotZ = 10:Model2WingLeft.RotZ = -10:Model2WingRight.RotZ = 10
    Case 11:Model1WingLeft.RotZ = -11:Model1WingRight.RotZ = 11:Model2WingLeft.RotZ = -11:Model2WingRight.RotZ = 11
    Case 12:Model1WingLeft.RotZ = -12:Model1WingRight.RotZ = 12:Model2WingLeft.RotZ = -12:Model2WingRight.RotZ = 12
    Case 13:Model1WingLeft.RotZ = -13:Model1WingRight.RotZ = 13:Model2WingLeft.RotZ = -13:Model2WingRight.RotZ = 13
    Case 14:Model1WingLeft.RotZ = -14:Model1WingRight.RotZ = 14:Model2WingLeft.RotZ = -14:Model2WingRight.RotZ = 14
    Case 15:Model1WingLeft.RotZ = -15:Model1WingRight.RotZ = 15:Model2WingLeft.RotZ = -15:Model2WingRight.RotZ = 15
    Case 16:Model1WingLeft.RotZ = -16:Model1WingRight.RotZ = 16:Model2WingLeft.RotZ = -16:Model2WingRight.RotZ = 16
    Case 17:Model1WingLeft.RotZ = -17:Model1WingRight.RotZ = 17:Model2WingLeft.RotZ = -17:Model2WingRight.RotZ = 17
    Case 18:Model1WingLeft.RotZ = -18:Model1WingRight.RotZ = 18:Model2WingLeft.RotZ = -18:Model2WingRight.RotZ = 18
    Case 19:Model1WingLeft.RotZ = -19:Model1WingRight.RotZ = 19:Model2WingLeft.RotZ = -19:Model2WingRight.RotZ = 19
    Case 20:Model1WingLeft.RotZ = -20:Model1WingRight.RotZ = 20:Model2WingLeft.RotZ = -20:Model2WingRight.RotZ = 20
    Case 21:Model1WingLeft.RotZ = -21:Model1WingRight.RotZ = 21:Model2WingLeft.RotZ = -21:Model2WingRight.RotZ = 21
    Case 22:Model1WingLeft.RotZ = -22:Model1WingRight.RotZ = 22:Model2WingLeft.RotZ = -22:Model2WingRight.RotZ = 22
    Case 23:Model1WingLeft.RotZ = -23:Model1WingRight.RotZ = 23:Model2WingLeft.RotZ = -23:Model2WingRight.RotZ = 23
    Case 24:Model1WingLeft.RotZ = -24:Model1WingRight.RotZ = 24:Model2WingLeft.RotZ = -24:Model2WingRight.RotZ = 24
    Case 25:Model1WingLeft.RotZ = -25:Model1WingRight.RotZ = 25:Model2WingLeft.RotZ = -25:Model2WingRight.RotZ = 25
    Case 26:Model1WingLeft.RotZ = -26:Model1WingRight.RotZ = 26:Model2WingLeft.RotZ = -26:Model2WingRight.RotZ = 26
    Case 27:Model1WingLeft.RotZ = -27:Model1WingRight.RotZ = 27:Model2WingLeft.RotZ = -27:Model2WingRight.RotZ = 27
    Case 28:Model1WingLeft.RotZ = -28:Model1WingRight.RotZ = 28:Model2WingLeft.RotZ = -28:Model2WingRight.RotZ = 28
    Case 29:Model1WingLeft.RotZ = -29:Model1WingRight.RotZ = 29:Model2WingLeft.RotZ = -29:Model2WingRight.RotZ = 29
    Case 30:Model1WingLeft.RotZ = -30:Model1WingRight.RotZ = 30:Model2WingLeft.RotZ = -30:Model2WingRight.RotZ = 30
    Case 31:Model1WingLeft.RotZ = -31:Model1WingRight.RotZ = 31:Model2WingLeft.RotZ = -31:Model2WingRight.RotZ = 31
    Case 32:Model1WingLeft.RotZ = -32:Model1WingRight.RotZ = 32:Model2WingLeft.RotZ = -32:Model2WingRight.RotZ = 32
    Case 33:Model1WingLeft.RotZ = -33:Model1WingRight.RotZ = 33:Model2WingLeft.RotZ = -33:Model2WingRight.RotZ = 33
    Case 34:Model1WingLeft.RotZ = -34:Model1WingRight.RotZ = 34:Model2WingLeft.RotZ = -34:Model2WingRight.RotZ = 34
    Case 35:Model1WingLeft.RotZ = -35:Model1WingRight.RotZ = 35:Model2WingLeft.RotZ = -35:Model2WingRight.RotZ = 35
    Case 36:Model1WingLeft.RotZ = -36:Model1WingRight.RotZ = 36:Model2WingLeft.RotZ = -36:Model2WingRight.RotZ = 36
    Case 37:Model1WingLeft.RotZ = -37:Model1WingRight.RotZ = 37:Model2WingLeft.RotZ = -37:Model2WingRight.RotZ = 37
    Case 38:Model1WingLeft.RotZ = -38:Model1WingRight.RotZ = 38:Model2WingLeft.RotZ = -38:Model2WingRight.RotZ = 38
    Case 39:Model1WingLeft.RotZ = -39:Model1WingRight.RotZ = 39:Model2WingLeft.RotZ = -39:Model2WingRight.RotZ = 39
    Case 40:Model1WingLeft.RotZ = -40:Model1WingRight.RotZ = 40:Model2WingLeft.RotZ = -40:Model2WingRight.RotZ = 40
    Case 41:Model1WingLeft.RotZ = -41:Model1WingRight.RotZ = 41:Model2WingLeft.RotZ = -41:Model2WingRight.RotZ = 41
    Case 42:Model1WingLeft.RotZ = -42:Model1WingRight.RotZ = 42:Model2WingLeft.RotZ = -42:Model2WingRight.RotZ = 42
    Case 43:Model1WingLeft.RotZ = -43:Model1WingRight.RotZ = 43:Model2WingLeft.RotZ = -43:Model2WingRight.RotZ = 43
    Case 44:Model1WingLeft.RotZ = -44:Model1WingRight.RotZ = 44:Model2WingLeft.RotZ = -44:Model2WingRight.RotZ = 44
    Case 45:Model1WingLeft.RotZ = -45:Model1WingRight.RotZ = 45:Model2WingLeft.RotZ = -45:Model2WingRight.RotZ = 45:if MoveWingsDir > 0 then me.Enabled = false:exit sub
  End Select
  MoveWingsStep = MoveWingsStep + MoveWingsDir
End Sub




'*************************
'GI color mod
'*************************
Dim GIxx, ColorModRed, ColorModRedFull, ColorModGreen, ColorModGreenFull, ColorModBlue, ColorModBlueFull
Dim GIColorRed, GIColorGreen, GIColorBlue, GIColorFullRed, GIColorFullGreen, GIColorFullBlue

'Choose your own custom color for General Illumination
  'Primary Colors
    'Red = 255, 0, 0
    'Green = 0, 255, 0
    'Blue = 0, 0, 255
    'Incandescent = 255, 197, 143
    'Refer to https://rgbcolorcode.com for customized color codes
'Enter RGB values below for "GI" color
GIColorRed       =  255
GIColorGreen     =  197
GIColorBlue      =  143

Sub SetGIColor ()
' for each GIxx in GILighting
' GIxx.Color = rgb(GIColorRed, GIColorGreen, GIColorBlue)
' GIxx.ColorFull = rgb(GIColorFullRed, GIColorFullGreen, GIColorFullBlue)
' next
  for each GIxx in GILighting
    GIxx.Color = rgb(GIColorRed, GIColorGreen, GIColorBlue)
  next
End Sub



'*************************
' Ball GI
'*************************

Sub DarkenBall1_Hit 'Darkens ball in the shooter lane
  activeball.color = RGB(150,150,150)
End Sub

Sub DarkenBall1_Unhit
  activeball.color = RGB(255,255,255)
End Sub

Sub DarkenBall2_Hit 'Darkens ball in the shooter lane
  activeball.color = RGB(150,150,150)
End Sub

Sub DarkenBall2_Unhit
  activeball.color = RGB(255,255,255)
End Sub

Sub DarkenBall3_Hit 'Darkens ball in the shooter lane
  activeball.color = RGB(150,150,150)
End Sub

Sub DarkenBall3_Unhit
  activeball.color = RGB(255,255,255)
End Sub

Sub DarkenBall4_Hit 'Darkens ball in the shooter lane
  activeball.color = RGB(150,150,150)
End Sub

Sub DarkenBall4_Unhit
  activeball.color = RGB(255,255,255)
End Sub

Sub DarkenBall5_Hit 'Darkens ball in the shooter lane
  activeball.color = RGB(150,150,150)
End Sub

Sub DarkenBall5_Unhit
  activeball.color = RGB(255,255,255)
End Sub

'******************************************************
' ZFLP: FLIPPERS
'******************************************************
Const ReflipAngle = 20

Sub SolLFlipper(Enabled)
  If Enabled Then
    LF.fire
    If leftflipper.currentangle < leftflipper.endangle + ReflipAngle Then
      RandomSoundReflipUpLeft LeftFlipper
    Else
      SoundFlipperUpAttackLeft LeftFlipper
      RandomSoundFlipperUpLeft LeftFlipper
    End If
  Else
    LeftFlipper.RotateToStart
    If LeftFlipper.currentangle < LeftFlipper.startAngle - 5 Then
      RandomSoundFlipperDownLeft LeftFlipper
    End If
    FlipperLeftHitParm = FlipperUpSoundLevel
  End If
End Sub

Sub SolULFlipper(Enabled)
  If Enabled Then
    LeftFlipper1.RotateToEnd
    If leftflipper1.currentangle < LeftFlipper1.endangle + ReflipAngle Then
      RandomSoundReflipUpLeft LeftFlipper1
    Else
      SoundFlipperUpAttackLeft LeftFlipper1
      RandomSoundFlipperUpLeft LeftFlipper1
    End If
  Else
    LeftFlipper1.RotateToStart
    If LeftFlipper1.currentangle < LeftFlipper1.startAngle - 5 Then
      RandomSoundFlipperDownLeft LeftFlipper1
    End If
    FlipperLeftHitParm = FlipperUpSoundLevel
  End If
End Sub

Sub SolRFlipper(Enabled)
  If Enabled Then
    RF.fire
    If rightflipper.currentangle > rightflipper.endangle - ReflipAngle Then
      RandomSoundReflipUpRight RightFlipper
    Else
      SoundFlipperUpAttackRight RightFlipper
      RandomSoundFlipperUpRight RightFlipper
    End If
  Else
    RightFlipper.RotateToStart
    If RightFlipper.currentangle > RightFlipper.startAngle + 5 Then
      RandomSoundFlipperDownRight RightFlipper
    End If
    FlipperRightHitParm = FlipperUpSoundLevel
  End If
End Sub

Sub SolURFlipper(Enabled)
  If Enabled Then
    RightFlipper1.RotateToEnd
    If rightflipper1.currentangle > rightflipper1.endangle - ReflipAngle Then
      RandomSoundReflipUpRight RightFlipper1
    Else
      SoundFlipperUpAttackRight RightFlipper1
      RandomSoundFlipperUpRight RightFlipper1
    End If
  Else
    RightFlipper1.RotateToStart
    If RightFlipper1.currentangle > RightFlipper1.startAngle + 5 Then
      RandomSoundFlipperDownRight RightFlipper1
    End If
    FlipperRightHitParm = FlipperUpSoundLevel
  End If
End Sub


'Flipper collide subs
Sub LeftFlipper_Collide(parm)
  CheckLiveCatch Activeball, LeftFlipper, LFCount, parm
  LF.ReProcessBalls ActiveBall
  LeftFlipperCollide parm
End Sub

Sub RightFlipper_Collide(parm)
  CheckLiveCatch Activeball, RightFlipper, RFCount, parm
  RF.ReProcessBalls ActiveBall
  RightFlipperCollide parm
End Sub

Sub LeftFlipper1_Collide(parm)
  LeftFlipperCollide parm
End Sub

Sub RightFlipper1_Collide(parm)
  RightFlipperCollide parm
End Sub

'******************************************************
' ZBRL: BALL ROLLING AND DROP SOUNDS
'******************************************************
'
' Be sure to call RollingUpdate in a timer with a 10ms interval see the GameTimer_Timer() sub

Dim inPlungerLane

ReDim rolling(tnob)
InitRolling

Dim DropCount
ReDim DropCount(tnob)

Sub InitRolling
  Dim i
  For i = 0 to tnob
    rolling(i) = False
  Next
End Sub

Sub RollingUpdate()
  Dim b', BOT
' BOT = GetBalls

  ' stop the sound of deleted balls
  For b = UBound(gBOT) + 1 to tnob - 1
    rolling(b) = False
    StopSound("BallRoll_" & b)
  Next

  ' exit the sub if no balls on the table
  If UBound(gBOT) = -1 Then Exit Sub

  inPlungerLane = 0

  ' play the rolling sound for each ball

  For b = 0 to UBound(gBOT)
    If BallVel(gBOT(b)) > 1 AND gBOT(b).z < 30 Then
      rolling(b) = True
      PlaySound ("BallRoll_" & b), -1, VolPlayfieldRoll(gBOT(b)) * BallRollVolume * VolumeDial, AudioPan(gBOT(b)), 0, PitchPlayfieldRoll(gBOT(b)), 1, 0, AudioFade(gBOT(b))

    Else
      If rolling(b) = True Then
        StopSound("BallRoll_" & b)
        rolling(b) = False
      End If
    End If

    ' Ball Drop Sounds
    If gBOT(b).VelZ < -1 and gBOT(b).z < 55 and gBOT(b).z > 27 Then 'height adjust for ball drop sounds
      If DropCount(b) >= 5 Then
        DropCount(b) = 0
        If gBOT(b).velz > -7 Then
          RandomSoundBallBouncePlayfieldSoft gBOT(b)
        Else
          RandomSoundBallBouncePlayfieldHard gBOT(b)
        End If
      End If
    End If
    If DropCount(b) < 5 Then
      DropCount(b) = DropCount(b) + 1
    End If
' Next

    ' ********* Fighter Jet Plunger
    If (InRect(gBOT(b).x,gBOT(b).y,872,1045,926,1045,926,1757,872,1757) or (InRect(gBOT(b).x,gBOT(b).y,872,475,926,475,926,1757,872,1757) and gBOT(b).z > 70)) and F14Plunger <> 0 then
      If (fighterprim.y > gBOT(b).y and gBOT(b).vely < 0) or gBOT(b).y > 1710 then
        fighterprim.y = gBOT(b).y
        pflame1.visible = 1
        pflame2.visible = 1
        pflame3.visible = 1
        pflame4.visible = 1
        inPlungerLane = 1
      Elseif fighterprim.y < gBOT(b).y and gBOT(b).vely > 0 then
        fighterprim.y = gBOT(b).y
        pflame1.visible = 0
        pflame2.visible = 0
        pflame3.visible = 0
        pflame4.visible = 0
        inPlungerLane = 0
      Else
        inPlungerLane = 0
      End If
      If fighterprim.y > 1680 then fighterprim.y = 1680
      pflame1.y = fighterprim.y + 73
      pflame2.y = fighterprim.y + 73
      pflame3.y = fighterprim.y + 75
      pflame4.y = fighterprim.y + 75
    End If

  Next

  If F14Plunger <> 0 and inPlungerLane = 0 and fighterprim.y < 1720 Then
    fighterprim.y = fighterprim.y  + 15*F14ReturnSpeed
    If fighterprim.y > 1680 then fighterprim.y = 1680
    pflame1.y = fighterprim.y + 73
    pflame2.y = fighterprim.y + 73
    pflame3.y = fighterprim.y + 75
    pflame4.y = fighterprim.y + 75
    pflame1.visible = 0
    pflame2.visible = 0
    pflame3.visible = 0
    pflame4.visible = 0
  End If


End Sub


'******************************************************
'****  END BALL ROLLING AND DROP SOUNDS
'******************************************************



'******************************************************
'   ZRRL: RAMP ROLLING SFX
'******************************************************

'Ball tracking ramp SFX 1.0
'   Reqirements:
'          * Import A Sound File for each ball on the table for plastic ramps.  Call It RampLoop<Ball_Number> ex: RampLoop1, RampLoop2, ...
'          * Import a Sound File for each ball on the table for wire ramps. Call it WireLoop<Ball_Number> ex: WireLoop1, WireLoop2, ...
'          * Create a Timer called RampRoll, that is enabled, with a interval of 100
'          * Set RampBAlls and RampType variable to Total Number of Balls
' Usage:
'          * Setup hit events and call WireRampOn True or WireRampOn False (True = Plastic ramp, False = Wire Ramp)
'          * To stop tracking ball
'                 * call WireRampOff
'                 * Otherwise, the ball will auto remove if it's below 30 vp units
'

dim RampMinLoops : RampMinLoops = 4

' RampBalls
'      Setup:        Set the array length of x in RampBalls(x,2) Total Number of Balls on table + 1:  if tnob = 5, then RammBalls(6,2)
'      Description:
dim RampBalls(5,2)
'x,0 = ball x,1 = ID, 2 = Protection against ending early (minimum amount of updates)
'0,0 is boolean on/off, 0,1 unused for now
RampBalls(0,0) = False

' RampType
'     Setup: Set this array to the number Total number of balls that can be tracked at one time + 1.  5 ball multiball then set value to 6
'     Description: Array type indexed on BallId and a values used to deterimine what type of ramp the ball is on: False = Wire Ramp, True = Plastic Ramp
dim RampType(5)

Sub WireRampOn(input)  : Waddball ActiveBall, input : RampRollUpdate: End Sub
Sub WireRampOff() : WRemoveBall ActiveBall.ID : End Sub


' WaddBall (Active Ball, Boolean)
'     Description: This subroutine is called from WireRampOn to Add Balls to the RampBalls Array
Sub Waddball(input, RampInput)  'Add ball
  ' This will loop through the RampBalls array checking each element of the array x, position 1
  ' To see if the the ball was already added to the array.
  ' If the ball is found then exit the subroutine
  dim x : for x = 1 to uBound(RampBalls)  'Check, don't add balls twice
    if RampBalls(x, 1) = input.id then
      if Not IsEmpty(RampBalls(x,1) ) then Exit Sub 'Frustating issue with BallId 0. Empty variable = 0
    End If
  Next

  ' This will itterate through the RampBalls Array.
  ' The first time it comes to a element in the array where the Ball Id (Slot 1) is empty.  It will add the current ball to the array
  ' The RampBalls assigns the ActiveBall to element x,0 and ball id of ActiveBall to 0,1
  ' The RampType(BallId) is set to RampInput
  ' RampBalls in 0,0 is set to True, this will enable the timer and the timer is also turned on
  For x = 1 to uBound(RampBalls)
    if IsEmpty(RampBalls(x, 1)) then
      Set RampBalls(x, 0) = input
      RampBalls(x, 1) = input.ID
      RampType(x) = RampInput
      RampBalls(x, 2) = 0
      'exit For
      RampBalls(0,0) = True
      RampRoll.Enabled = 1   'Turn on timer
      'RampRoll.Interval = RampRoll.Interval 'reset timer
      exit Sub
    End If
    if x = uBound(RampBalls) then   'debug
      Debug.print "WireRampOn error, ball queue is full: " & vbnewline & _
      RampBalls(0, 0) & vbnewline & _
      Typename(RampBalls(1, 0)) & " ID:" & RampBalls(1, 1) & "type:" & RampType(1) & vbnewline & _
      Typename(RampBalls(2, 0)) & " ID:" & RampBalls(2, 1) & "type:" & RampType(2) & vbnewline & _
      Typename(RampBalls(3, 0)) & " ID:" & RampBalls(3, 1) & "type:" & RampType(3) & vbnewline & _
      Typename(RampBalls(4, 0)) & " ID:" & RampBalls(4, 1) & "type:" & RampType(4) & vbnewline & _
      Typename(RampBalls(5, 0)) & " ID:" & RampBalls(5, 1) & "type:" & RampType(5) & vbnewline & _
      " "
    End If
  next
End Sub

' WRemoveBall (BallId)
'    Description: This subroutine is called from the RampRollUpdate subroutine
'                 and is used to remove and stop the ball rolling sounds
Sub WRemoveBall(ID)   'Remove ball
  'Debug.Print "In WRemoveBall() + Remove ball from loop array"
  dim ballcount : ballcount = 0
  dim x : for x = 1 to Ubound(RampBalls)
    if ID = RampBalls(x, 1) then 'remove ball
      Set RampBalls(x, 0) = Nothing
      RampBalls(x, 1) = Empty
      RampType(x) = Empty
      StopSound("RampLoop" & x)
      StopSound("wireloop" & x)
    end If
    'if RampBalls(x,1) = Not IsEmpty(Rampballs(x,1) then ballcount = ballcount + 1
    if not IsEmpty(Rampballs(x,1)) then ballcount = ballcount + 1
  next
  if BallCount = 0 then RampBalls(0,0) = False  'if no balls in queue, disable timer update
End Sub

Sub RampRoll_Timer():RampRollUpdate:End Sub

Sub RampRollUpdate()    'Timer update
  dim x : for x = 1 to uBound(RampBalls)
    if Not IsEmpty(RampBalls(x,1) ) then
      if BallVel(RampBalls(x,0) ) > 1 then ' if ball is moving, play rolling sound
        If RampType(x) then
          PlaySound("RampLoop" & x), -1, VolPlayfieldRoll(RampBalls(x,0)) * RampRollVolume * VolumeDial, AudioPan(RampBalls(x,0)), 0, BallPitchV(RampBalls(x,0)), 1, 0, AudioFade(RampBalls(x,0))
          StopSound("wireloop" & x)
        Else
          StopSound("RampLoop" & x)
          PlaySound("wireloop" & x), -1, VolPlayfieldRoll(RampBalls(x,0)) * RampRollVolume * VolumeDial, AudioPan(RampBalls(x,0)), 0, BallPitch(RampBalls(x,0)), 1, 0, AudioFade(RampBalls(x,0))
        End If
        RampBalls(x, 2) = RampBalls(x, 2) + 1
      Else
        StopSound("RampLoop" & x)
        StopSound("wireloop" & x)
      end if
      if RampBalls(x,0).Z < 30 and RampBalls(x, 2) > RampMinLoops then  'if ball is on the PF, remove  it
        StopSound("RampLoop" & x)
        StopSound("wireloop" & x)
        Wremoveball RampBalls(x,1)
      End If
    Else
      StopSound("RampLoop" & x)
      StopSound("wireloop" & x)
    end if
  next
  if not RampBalls(0,0) then RampRoll.enabled = 0

End Sub

' This can be used to debug the Ramp Roll time.  You need to enable the tbWR timer on the TextBox
Sub tbWR_Timer()  'debug textbox
  me.text = "on? " & RampBalls(0, 0) & " timer: " & RampRoll.Enabled & vbnewline & _
  "1 " & Typename(RampBalls(1, 0)) & " ID:" & RampBalls(1, 1) & " type:" & RampType(1) & " Loops:" & RampBalls(1, 2) & vbnewline & _
  "2 " & Typename(RampBalls(2, 0)) & " ID:" & RampBalls(2, 1) & " type:" & RampType(2) & " Loops:" & RampBalls(2, 2) & vbnewline & _
  "3 " & Typename(RampBalls(3, 0)) & " ID:" & RampBalls(3, 1) & " type:" & RampType(3) & " Loops:" & RampBalls(3, 2) & vbnewline & _
  "4 " & Typename(RampBalls(4, 0)) & " ID:" & RampBalls(4, 1) & " type:" & RampType(4) & " Loops:" & RampBalls(4, 2) & vbnewline & _
  "5 " & Typename(RampBalls(5, 0)) & " ID:" & RampBalls(5, 1) & " type:" & RampType(5) & " Loops:" & RampBalls(5, 2) & vbnewline & _
  "6 " & Typename(RampBalls(6, 0)) & " ID:" & RampBalls(6, 1) & " type:" & RampType(6) & " Loops:" & RampBalls(6, 2) & vbnewline & _
  " "
End Sub


Function BallPitch(ball) ' Calculates the pitch of the sound based on the ball speed
    BallPitch = pSlope(BallVel(ball), 1, -1000, 60, 10000)
End Function

Function BallPitchV(ball) ' Calculates the pitch of the sound based on the ball speed Variation
  BallPitchV = pSlope(BallVel(ball), 1, -4000, 60, 7000)
End Function


'Ramp triggers
Sub Ramp1_Enter_Hit
  If activeball.vely < 0 Then WireRampOn False
End Sub

Sub Ramp1_Enter_UnHit
  If activeball.vely > 0 Then WireRampOff
End Sub

Sub Ramp1_Enter2_Hit
  If activeball.velx < 0 Then WireRampOn False
End Sub

Sub Ramp1_Enter2_UnHit
  If activeball.velx > 0 Then WireRampOff
End Sub

Sub Ramp1_Change_Hit
  If activeball.velx < 0 Then WireRampOn True
End Sub

Sub Ramp1_Change_UnHit
  If activeball.velx > 0 Then WireRampOn False
End Sub

Sub Ramp2_End_Hit
    WireRampOff
End Sub

Sub Ramp3_End_Hit
    WireRampOff
End Sub

Sub Ramp4_End_Hit
    WireRampOff
End Sub

Sub RandomSoundRampStop(obj)
  Select Case Int(rnd*3)
    Case 0: PlaySoundAtVol "wireramp_stop1", obj, 0.02*volumedial:PlaySoundAtLevelActiveBall ("Rubber_Strong_1"), Vol(ActiveBall) * RubberStrongSoundFactor * 0.6
    Case 1: PlaySoundAtVol "wireramp_stop2", obj, 0.02*volumedial:PlaySoundAtLevelActiveBall ("Rubber_Strong_2"), Vol(ActiveBall) * RubberStrongSoundFactor * 0.6
    Case 2: PlaySoundAtVol "wireramp_stop3", obj, 0.02*volumedial:PlaySoundAtLevelActiveBall ("Rubber_1_Hard"), Vol(ActiveBall) * RubberStrongSoundFactor * 0.6
  End Select
End Sub



'******************************************************
'**** END RAMP ROLLING SFX
'******************************************************



'******************************************************
'   ZFLE:  FLEEP MECHANICAL SOUNDS
'******************************************************

' This part in the script is an entire block that is dedicated to the physics sound system.
' Various scripts and sounds that may be pretty generic and could suit other WPC systems, but the most are tailored specifically for the TOM table

' Many of the sounds in this package can be added by creating collections and adding the appropriate objects to those collections.
' Create the following new collections:
'  Metals (all metal objects, metal walls, metal posts, metal wire guides)
'  Apron (the apron walls and plunger wall)
'  Walls (all wood or plastic walls)
'  Rollovers (wire rollover triggers, star triggers, or button triggers)
'  Targets (standup or drop targets, these are hit sounds only ... you will want to add separate dropping sounds for drop targets)
'  Gates (plate gates)
'  GatesWire (wire gates)
'  Rubbers (all rubbers including posts, sleeves, pegs, and bands)
' When creating the collections, make sure "Fire events for this collection" is checked.
' You'll also need to make sure "Has Hit Event" is checked for each object placed in these collections (not necessary for gates and triggers).
' Once the collections and objects are added, the save, close, and restart VPX.
'
' Many places in the script need to be modified to include the correct sound effect subroutine calls. The tutorial videos linked below demonstrate
' how to make these updates. But in summary the following needs to be updated:
' - Nudging, plunger, coin-in, start button sounds will be added to the keydown and keyup subs.
' - Flipper sounds in the flipper solenoid subs. Flipper collision sounds in the flipper collide subs.
' - Bumpers, slingshots, drain, ball release, knocker, spinner, and saucers in their respective subs
' - Ball rolling sounds sub
'
' Tutorial videos by Apophis
' Audio : Adding Fleep Part 1         https://youtu.be/rG35JVHxtx4?si=zdN9W4cZWEyXbOz_
' Audio : Adding Fleep Part 2         https://youtu.be/dk110pWMxGo?si=2iGMImXXZ0SFKVCh
' Audio : Adding Fleep Part 3         https://youtu.be/ESXWGJZY_EI?si=6D20E2nUM-xAw7xy


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

GateSoundLevel = 0.3 / 5      'volume level; range [0, 1]
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
'****  END FLEEP MECHANICAL SOUNDS
'******************************************************







'******************************************************
' ZPHY:  GNEREAL ADVICE ON PHYSICS
'******************************************************
'
' It's advised that flipper corrections, dampeners, and general physics settings should all be updated per these
' examples as all of these improvements work together to provide a realistic physics simulation.
'
' Tutorial videos provided by Bord
' Adding nFozzy roth physics : pt1 rubber dampeners         https://youtu.be/AXX3aen06FM?si=Xqd-rcaqTlgEd_wx
' Adding nFozzy roth physics : pt2 flipper physics          https://youtu.be/VSBFuK2RCPE?si=i8ne8Ao2co8rt7fy
' Adding nFozzy roth physics : pt3 other elements           https://youtu.be/JN8HEJapCvs?si=hvgMOk-ej1BEYjJv
'
' Note: BallMass must be set to 1. BallSize should be set to 50 (in other words the ball radius is 25)
'
' Recommended Table Physics Settings
' | Gravity Constant             | 0.97      |
' | Playfield Friction           | 0.15-0.25 |
' | Playfield Elasticity         | 0.25      |
' | Playfield Elasticity Falloff | 0         |
' | Playfield Scatter            | 0         |
' | Default Element Scatter      | 2         |
'
' Bumpers
' | Force         | 12-15    |
' | Hit Threshold | 1.6-2    |
' | Scatter Angle | 2        |
'
' Slingshots
' | Hit Threshold      | 2    |
' | Slingshot Force    | 3-5  |
' | Slingshot Theshold | 2-3  |
' | Elasticity         | 0.85 |
' | Friction           | 0.8  |
' | Scatter Angle      | 1    |






'******************************************************
' ZNFF:  FLIPPER CORRECTIONS by nFozzy
'******************************************************
'
' There are several steps for taking advantage of nFozzy's flipper solution.  At a high level we'll need the following:
' 1. flippers with specific physics settings
' 2. custom triggers for each flipper (TriggerLF, TriggerRF)
' 3. and, special scripting
'
' TriggerLF and RF should now be 27 vp units from the flippers. In addition, 3 degrees should be added to the end angle
' when creating these triggers.
'
' RF.ReProcessBalls Activeball and LF.ReProcessBalls Activeball must be added the flipper_collide subs.
'
' A common mistake is incorrect flipper length.  A 3-inch flipper with rubbers will be about 3.125 inches long.
' This translates to about 147 vp units.  Therefore, the flipper start radius + the flipper length + the flipper end
' radius should  equal approximately 147 vp units. Another common mistake is is that sometimes the right flipper
' angle was set with a large postive value (like 238 or something). It should be using negative value (like -122).
'
' The following settings are a solid starting point for various eras of pinballs.
' |                    | EM's           | late 70's to mid 80's | mid 80's to early 90's | mid 90's and later |
' | ------------------ | -------------- | --------------------- | ---------------------- | ------------------ |
' | Mass               | 1              | 1                     | 1                      | 1                  |
' | Strength           | 500-1000 (750) | 1400-1600 (1500)      | 2000-2600              | 3200-3300 (3250)   |
' | Elasticity         | 0.88           | 0.88                  | 0.88                   | 0.88               |
' | Elasticity Falloff | 0.15           | 0.15                  | 0.15                   | 0.15               |
' | Fricition          | 0.8-0.9        | 0.9                   | 0.9                    | 0.9                |
' | Return Strength    | 0.11           | 0.09                  | 0.07                   | 0.055              |
' | Coil Ramp Up       | 2.5            | 2.5                   | 2.5                    | 2.5                |
' | Scatter Angle      | 0              | 0                     | 0                      | 0                  |
' | EOS Torque         | 0.4            | 0.4                   | 0.375                  | 0.375              |
' | EOS Torque Angle   | 4              | 4                     | 6                      | 6                  |
'

'******************************************************
' Flippers Polarity (Select appropriate sub based on era)
'******************************************************

Dim LF : Set LF = New FlipperPolarity
Dim RF : Set RF = New FlipperPolarity

InitPolarity

'*******************************************
' Mid 80's
'*******************************************


Sub InitPolarity()
   dim x, a : a = Array(LF, RF)
  for each x in a
    x.AddPt "Ycoef", 0, RightFlipper.Y-65, 1 'disabled
    x.AddPt "Ycoef", 1, RightFlipper.Y-11, 1
    x.enabled = True
    x.TimeDelay = 80
    x.DebugOn=False ' prints some info in debugger

    x.AddPt "Polarity", 0, 0, 0
    x.AddPt "Polarity", 1, 0.05, - 3.7
    x.AddPt "Polarity", 2, 0.16, - 3.7
    x.AddPt "Polarity", 3, 0.22, - 0
    x.AddPt "Polarity", 4, 0.25, - 0
    x.AddPt "Polarity", 5, 0.3, - 2
    x.AddPt "Polarity", 6, 0.4, - 3
    x.AddPt "Polarity", 7, 0.5, - 3.7
    x.AddPt "Polarity", 8, 0.65, - 2.3
    x.AddPt "Polarity", 9, 0.75, - 1.5
    x.AddPt "Polarity", 10, 0.81, - 1
    x.AddPt "Polarity", 11, 0.88, 0
    x.AddPt "Polarity", 12, 1.3, 0

    x.AddPt "Velocity", 0, 0, 0.85
    x.AddPt "Velocity", 1, 0.15, 0.85
    x.AddPt "Velocity", 2, 0.2, 0.9
    x.AddPt "Velocity", 3, 0.23, 0.95
    x.AddPt "Velocity", 4, 0.41, 0.95
    x.AddPt "Velocity", 5, 0.53, 0.95 '0.982
    x.AddPt "Velocity", 6, 0.62, 1.0
    x.AddPt "Velocity", 7, 0.702, 0.968
    x.AddPt "Velocity", 8, 0.95,  0.968
    x.AddPt "Velocity", 9, 1.03,  0.945
    x.AddPt "Velocity", 10, 1.5,  0.945

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
  '   Dim BOT
  '   BOT = GetBalls

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

Dim LFPress, RFPress, ULFPress, LFCount, RFCount
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
Const EOSTnew = 1.5 'EM's to late 80's - new recommendation by rothbauerw (previously 1)
'Const EOSTnew = 1.2 '90's and later - new recommendation by rothbauerw (previously 0.8)
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
'   Const EOSReturn = 0.055  'EM's
'   Const EOSReturn = 0.045  'late 70's to mid 80's
Const EOSReturn = 0.035  'mid 80's to early 90's
'   Const EOSReturn = 0.025  'mid 90's and later

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
    Dim b', BOT
    '   BOT = GetBalls

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

' Put all the Post and Pin objects in dPosts collection. Make sure dPosts fires hit events.
Sub dPosts_Hit(idx)
  RubbersD.dampen ActiveBall
  TargetBouncer ActiveBall, 1
End Sub

' This collection contains the bottom sling posts. They are not in the dPosts collection so that the TargetBouncer is not applied to them, but they should still have dampening applied
' If you experience airballs with posts or targets, consider adding them to this collection
Sub NoTargetBouncer_Hit
    RubbersD.dampen ActiveBall
End Sub

' Put all the Sleeve objects in dSleeves collection. Make sure dSleeves fires hit events.
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



'******************************************************
'****  END PHYSICS DAMPENERS
'******************************************************



'******************************************************
'   ZBOU: VPW TargetBouncer for targets and posts by Iaakki, Wrd1972, Apophis
'******************************************************

Const TargetBouncerEnabled = 1    '0 = normal standup targets, 1 = bouncy targets
Const TargetBouncerFactor = 0.9   'Level of bounces. Recommmended value of 0.7-1.0

sub TargetBouncer(aBall,defvalue)
    dim zMultiplier, vel, vratio
    if TargetBouncerEnabled = 1 and aball.z < 30 then
        'debug.print "velx: " & aball.velx & " vely: " & aball.vely & " velz: " & aball.velz
        vel = BallSpeed(aBall)
        if aBall.velx = 0 then vratio = 1 else vratio = aBall.vely/aBall.velx
        Select Case Int(Rnd * 6) + 1
            Case 1: zMultiplier = 0.2*defvalue
      Case 2: zMultiplier = 0.25*defvalue
            Case 3: zMultiplier = 0.3*defvalue
      Case 4: zMultiplier = 0.4*defvalue
            Case 5: zMultiplier = 0.45*defvalue
            Case 6: zMultiplier = 0.5*defvalue
        End Select
        aBall.velz = abs(vel * zMultiplier * TargetBouncerFactor)
        aBall.velx = sgn(aBall.velx) * sqr(abs((vel^2 - aBall.velz^2)/(1+vratio^2)))
        aBall.vely = aBall.velx * vratio
        'debug.print "---> velx: " & aball.velx & " vely: " & aball.vely & " velz: " & aball.velz
        'debug.print "conservation check: " & BallSpeed(aBall)/vel
  end if
end sub

'****************************************************************
'****  END VPW TargetBouncer for targets and posts by Iaakki, Wrd1972, Apophis
'****************************************************************


'******************************************************
'   ZFLB:  FLUPPER BUMPERS
'******************************************************
' Based on FlupperBumpers 0.145 final

' Explanation of how these bumpers work:
' There are 10 elements involved per bumper:
' - the shadow of the bumper ( a vpx flasher object)
' - the bumper skirt (primitive)
' - the bumperbase (primitive)
' - a vpx light which colors everything you can see through the bumpertop
' - the bulb (primitive)
' - another vpx light which lights up everything around the bumper
' - the bumpertop (primitive)
' - the VPX bumper object
' - the bumper screws (primitive)
' - the bulb highlight VPX flasher object
' All elements have a special name with the number of the bumper at the end, this is necessary for the fading routine and the initialisation.
' For the bulb and the bumpertop there is a unique material as well per bumpertop.
' To use these bumpers you have to first copy all 10 elements to your table.
' Also export the textures (images) with names that start with "Flbumper" and "Flhighlight" and materials with names that start with "bumper".
' Make sure that all the ten objects are aligned on center, if possible with the exact same x,y coordinates
' After that copy the script (below); also copy the BumperTimer vpx object to your table
' Every bumper needs to be initialised with the FlInitBumper command, see example below;
' Colors available are red, white, blue, orange, yellow, green, purple and blacklight.
' In a GI subroutine you can then call set the bumperlight intensity with the "FlBumperFadeTarget(nr) = value" command
' where nr is the number of the bumper, value is between 0 (off) and 1 (full on) (so you can also use 0.3 0.4 etc).

' Notes:
' - There is only one color for the disk; you can photoshop it to a different color
' - The bumpertops are angle independent up to a degree; my estimate is -45 to + 45 degrees horizontally, 0 (topview) to 70-80 degrees (frontview)
' - I built in correction for the day-night slider; this might not work perfectly, depending on your table lighting
' - These elements, textures and materials do NOT integrate with any of the lighting routines I have seen in use in many VPX tables
'   (just find the GI handling routine and insert the FlBumperFadeTarget statement)
' - If you want to use VPX native bumperdisks just copy my bumperdisk but make it invisible

'' Uncomment this if you are not using PWM following bumpers
'BumperTimer.Enabled = 1
'Sub BumperTimer_Timer
' Dim nr
' For nr = 1 To 6
'   If FlBumperFadeActual(nr) < FlBumperFadeTarget(nr) And FlBumperActive(nr)  Then
'     FlBumperFadeActual(nr) = FlBumperFadeActual(nr) + (FlBumperFadeTarget(nr) - FlBumperFadeActual(nr)) * 0.8
'     If FlBumperFadeActual(nr) > 0.99 Then FlBumperFadeActual(nr) = 1
'     FlFadeBumper nr, FlBumperFadeActual(nr)
'   End If
'   If FlBumperFadeActual(nr) > FlBumperFadeTarget(nr) And FlBumperActive(nr)  Then
'     FlBumperFadeActual(nr) = FlBumperFadeActual(nr) + (FlBumperFadeTarget(nr) - FlBumperFadeActual(nr)) * 0.4 / (FlBumperFadeActual(nr) + 0.1)
'     If FlBumperFadeActual(nr) < 0.01 Then FlBumperFadeActual(nr) = 0
'     FlFadeBumper nr, FlBumperFadeActual(nr)
'   End If
' Next
'End Sub

' prepare some global vars to dim/brighten objects when using day-night slider
Dim DayNightAdjust , DNA30, DNA45, DNA90
If NightDay < 10 Then
  DNA30 = 0
  DNA45 = (NightDay - 10) / 20
  DNA90 = 0
  DayNightAdjust = 0.4
Else
  DNA30 = (NightDay - 10) / 30
  DNA45 = (NightDay - 10) / 45
  DNA90 = (NightDay - 10) / 90
  DayNightAdjust = NightDay / 25
End If

Dim FlBumperFadeActual(6), FlBumperFadeTarget(6), FlBumperColor(6), FlBumperTop(6), FlBumperSmallLight(6), Flbumperbiglight(6)
Dim FlBumperDisk(6), FlBumperBase(6), FlBumperBulb(6), FlBumperscrews(6), FlBumperActive(6), FlBumperHighlight(6)
Dim cnt
For cnt = 1 To 6
  FlBumperActive(cnt) = False
Next

' colors available are red, white, blue, orange, yellow, green, purple and blacklight
FlInitBumper 1, "red"

' ### uncomment the statement below to change the color for all bumpers ###
'   Dim ind
'   For ind = 1 To 5
'    FlInitBumper ind, "green"
'   Next

Sub FlInitBumper(nr, col)
  FlBumperActive(nr) = True

  ' store all objects in an array for use in FlFadeBumper subroutine
  FlBumperFadeActual(nr) = 1
  FlBumperFadeTarget(nr) = 1.1
  FlBumperColor(nr) = col
  Set FlBumperTop(nr) = Eval("bumpertop" & nr)
  FlBumperTop(nr).material = "bumpertopmat" & nr
  Set FlBumperSmallLight(nr) = Eval("bumpersmalllight" & nr)
  Set Flbumperbiglight(nr) = Eval("bumperbiglight" & nr)
  Set FlBumperDisk(nr) = Eval("bumperdisk" & nr)
  Set FlBumperBase(nr) = Eval("bumperbase" & nr)
  Set FlBumperBulb(nr) = Eval("bumperbulb" & nr)
  FlBumperBulb(nr).material = "bumperbulbmat" & nr
  Set FlBumperscrews(nr) = Eval("bumperscrews" & nr)
  FlBumperscrews(nr).material = "bumperscrew" & col
  Set FlBumperHighlight(nr) = Eval("bumperhighlight" & nr)

  ' set the color for the two VPX lights
  Select Case col
    Case "red"
    FlBumperSmallLight(nr).color = RGB(255,4,0)
    FlBumperSmallLight(nr).colorfull = RGB(255,24,0)
    FlBumperBigLight(nr).color = RGB(255,32,0)
    FlBumperBigLight(nr).colorfull = RGB(255,32,0)
    FlBumperHighlight(nr).color = RGB(64,255,0)
    FlBumperSmallLight(nr).BulbModulateVsAdd = 0.98
    FlBumperSmallLight(nr).TransmissionScale = 0
  End Select
End Sub

Sub FlFadeBumper(nr, Z)
  FlBumperBase(nr).BlendDisableLighting = 0.5 * DayNightAdjust
  '   UpdateMaterial(string, float wrapLighting, float roughness, float glossyImageLerp, float thickness, float edge, float edgeAlpha, float opacity,
  '        OLE_COLOR base, OLE_COLOR glossy, OLE_COLOR clearcoat, VARIANT_BOOL isMetal, VARIANT_BOOL opacityActive,
  '        float elasticity, float elasticityFalloff, float friction, float scatterAngle) - updates all parameters of a material
  FlBumperDisk(nr).BlendDisableLighting = (0.5 - Z * 0.3 ) * DayNightAdjust

  Select Case FlBumperColor(nr)

    Case "red"
    UpdateMaterial "bumperbulbmat" & nr, 0, 0.75 , 0, 1 - Z, 1 - Z, 1 - Z, 0.9999, RGB(255, 16 - 11 * Z + 16 * Sin(Z * 3.14),0), RGB(255,255,255), RGB(32,32,32), False, True, 0, 0, 0, 0
    FlBumperSmallLight(nr).intensity = 17 + 100 * Z / (1 + DNA30 ^ 2)
    FlBumperTop(nr).BlendDisableLighting = 3 * DayNightAdjust + 18 * Z / (1 + DNA90)
    FlBumperBulb(nr).BlendDisableLighting = 20 * DayNightAdjust + 9000 * (0.03 * Z + 0.97 * Z ^ 10)
    Flbumperbiglight(nr).intensity = 10 * Z / (1 + DNA45)
    FlBumperHighlight(nr).opacity = 2000 * (Z ^ 3) / (1 + DNA90)
    MaterialColor "bumpertopmat" & nr, RGB(255,20 + Z * 4,8 - Z * 8)

  End Select
End Sub



'******************************************************
'******  END FLUPPER BUMPERS
'******************************************************




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
Dim objBallShadow(4)

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

' Thalamus : Exit in a clean and proper way
Sub Table1_exit
  Controller.Pause = False
  Controller.Stop
End Sub



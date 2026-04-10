'______       ______           _                 _   _ _       _____             _ _            _    ______
'|  _  \      |  _  \         | |        ___    | | | (_)     |  ___|           | | |          | |   | ___ \ 
'| | | |_ __  | | | |_   _  __| | ___   ( _ )   | |_| |_ ___  | |____  _____ ___| | | ___ _ __ | |_  | |_/ /__ _ _   _
'| | | | '__| | | | | | | |/ _` |/ _ \  / _ \/\ |  _  | / __| |  __\ \/ / __/ _ \ | |/ _ \ '_ \| __| |    // _` | | | |
'| |/ /| |_   | |/ /| |_| | (_| |  __/ | (_>  < | | | | \__ \ | |___>  < (_|  __/ | |  __/ | | | |_  | |\ \ (_| | |_| |
'|___/ |_(_)  |___/  \__,_|\__,_|\___|  \___/\/ \_| |_/_|___/ \____/_/\_\___\___|_|_|\___|_| |_|\__| \_| \_\__,_|\__, |
'                                                                                                                 __/ |
'                                                                                                                |___/
'                                   ______       _ _         __   _____  _____  _____
'                                   | ___ \     | | |       /  | |  _  ||  _  ||  _  |
'                                   | |_/ / __ _| | |_   _  `| | | |_| || |_| || |/' |
'                                   | ___ \/ _` | | | | | |  | | \____ |\____ ||  /| |
'                                   | |_/ / (_| | | | |_| | _| |_.___/ /.___/ /\ |_/ /
'                                   \____/ \__,_|_|_|\__, | \___/\____/ \____/  \___/
'                                                     __/ |
'                                                    |___/                               '
'
'
'Dr. Dude (Bally 1990) Version 2.3 for VP10. REQUIRES VP10.6 final release or > to play.
'Early development by "gtxJoe" and completed by "wrd1972"
'Thanks to gtxJoe for his  assistance and allowing me to complete his table...which by the way, is my very first VP table...EVER!
'Extra thanks to "ninuzzu" for the invaluable guidance he provided me early on when I started authoring.
'Also extra thanks to Cyberpez for the many small 3D jobs and script tweaks.
'Additional scripting assistance by "Rothbauerw" and Cyberpez"
'PF Inserts, Flasher Domes, pop-bumper caps, GOG model, Bully model, side-rail image, ball shadows and additional scripting by by "Flupper"
'Misc 3d work by "Shreibi"
'Clear ramp 3D model by "Flupper"
'New ramp texture/render, new Mixmaster 3D model, and other misc 3D modeling/tweaking by "Benji"
'White flipper prims by "Zany"
'Flipper shadows by "ninuzzu"
'Graphic Remaster - hand drawn playfield and plastics by Brad1X. The "Dr. Dude" re-mastered playfield and plastics graphic files are only authorized to be used free of charge and may not be altered, redistributed, reused or sold without direct permission from Brad1X
'Previous Playfield image rework by "ClarkKent"
'Desktop table scoring displays by "32Assassin"
'All lighting by "wrd1972"
'Physics by "wrd1972", partially based off of "nfozzys physics"
'Flipper trajectory fix by "nFozzy, Rothbauer"
'DOF by "Arngrim"
'Mixmaster interior and coiled wires 3D models by "nfozzy"
'Totally awesome new light fading scripting by "nFozzy"
'Color-Mod scripting by "cyberpez" and "wrd1972"
'PMD feedback and surround sound mods by "Rusty Cardores" & "DJRobX"
'New Mixmaster physical spinning post by RothBauerw
'New saucer physics by Cyberpez and Rothbauerw
'Flipper, rubber, bumper, and other MISC collision sounds by "Fleep, nickbuol, CalleV, Nicolas Mazaleyrat, Jon Osborne"
'Prim inserts by "Schreibi"

'Many thanks to countles others for in the VPF community for helping me learn the table creation process.

'Notes:
'Most object physics are controller by materials physics, including the new primitive playfield surface. Refer to the materials if you want to easily adjust the object physics.
'Table absolutely REQUIRES VP10.6 beta or > to play. Or the saucers (VUK's) will not work properly.
'If you are encounter performance issues and low frame rate. I would reccomend disabling "Reflect Elements On Playfield". This will increase performance for lesser powerful systems.

'The playfield and plastics graphic files are only authorized to be used free of charge and may not be redistributed, reused or sold without permission from Brad1X.


'VPW CHANGE LOG
' 01 - apophis - Updated to latest physics, fleep sounds, ball and ramp rolling sounds, dynamic shadows
' 02 - apophis - Fixed visibility on one prim. Reduced plunger release speed. Added some bloom lights to the three flasher domes.
' 03 - leojreimroc - VR Room and backglass.  Fixed a few missing materials (gates and bolts)
' 04 - leojreimroc - Fixed relection issue in VR Room
' 08 - TastyWasps - Added VPW menu
' 09 - TastyWasps - nFozzy materials corrected, flipper length corrected. nudges toned down, kicker strengths reduced from 90 to 45
' 10 - TastyWasps - Friction increased on left and right wire ramps per wrd1972, slow down code to help with ski jumps, sleeves falloff changed to 0.05 for less bricky shots, Desktop Background darkened and upscaled, table slope changed from 6.5 to 6.25
' 11 - Sixtoe - Messed around with the RayGun and flashers, tweaked a few things, set all height walls to non-collidable.
' 12 - apophis - Fixed some rubberband walls. Adjusted flipper trigger shapes. Fixed post passes. Big guy shakes around x and y axes now.
' 14 - TastyWasps - Brightened desktop digits, tuned ramp friction levels
' 3.0 - TastyWasps - VR correction, Slope changed to 6.5, tuned ramp friction levels, fixed flex DMD issue and wording, apophis - Reworked big shot animation,


'***************************************************************************************************************************************************************
'***************************************************************************************************************************************************************

Option Explicit
Randomize
On Error Resume Next
ExecuteGlobal GetTextFile("controller.vbs")
If Err Then MsgBox "You need the controller.vbs in order to run this table, available in the vp10 package"
On Error Goto 0

Const TableVersion = "3.0"

'***************************************************************************************************************************************************************
'***************************************************************************************************************************************************************
' _____     _     _        ___________ _   _                   _   _
'|_   _|   | |   | |      |  _  | ___ \ | (_)                 | | | |
'  | | __ _| |__ | | ___  | | | | |_/ / |_ _  ___  _ __  ___  | |_| | ___ _ __ ___
'  | |/ _` | '_ \| |/ _ \ | | | |  __/| __| |/ _ \| '_ \/ __| |  _  |/ _ \ '__/ _ \ 
'  | | (_| | |_) | |  __/ \ \_/ / |   | |_| | (_) | | | \__ \ | | | |  __/ | |  __/
'  \_/\__,_|_.__/|_|\___|  \___/\_|    \__|_|\___/|_| |_|___/ \_| |_/\___|_|  \___|
'

' VR ROOM CHOICE
' VR Room = 1   --  Ba Sti/Rawd Room
' VR Room = 2   --  Minimal Room
' VR Room = 3   --  Ultra Minimal Room
Const VRRoomChoice = 1


'GI COLOR MOD
'   White Bulbs = 0
'   Blue Bulbs = 1
'   Purple Bulbs = 2
'   Ice-Blue Bulbs = 3
'   Random = 4
'Change the value below to set option
GIColorMod = 0


'APRON COLOR MOD
'   White Apron = 0
' Blue Apron = 1
'   Green Apron = 2
'   Custom Apron & Walls = 3
'   Random = 4
'Change the value below to set option
Aproncolor = 1


'FLIPPER BAT STYLE MOD
'   Yellow Bats = 0
'   White Bats = 1
'   Random = 2
'Change the value below to set option
Flipperstyle = 0


'CHEATER DRAIN POST MOD
'   Disable Cheater Post = 0
'   Enable Cheater Post = 1
'Change the value below to set option
CheaterPost = 0


'EXCELLENT RAY BEAM MOD
'   No Raybeam = 0
'   Show Raybeam = 1
'Change the value below to set option
Raybeam = 0


' SIDE RAILS
'   Show Side Rails = 0
'   Hide Side Rails = 1
'Change the value below to set option
Rails = 0


' DYNAMIC BALL SHADOWS
Const DynamicBallShadowsOn = 1    '0 = no dynamic ball shadow ("triangles" near slings and such), 1 = enable dynamic ball shadow
Const AmbientBallShadowOn = 1   '0 = Static shadow under ball ("flasher" image, like JP's)
                  '1 = Moving ball shadow ("primitive" object, like ninuzzu's) - This is the only one that shows up on the pf when in ramps and fades when close to lights!
                  '2 = flasher image shadow, but it moves like ninuzzu's

'INTRO MUSIC
'   Add Intro Music Snippet =0
'   No Intro Music Snippet = 1
'Change the value below to set option
IntroMusic = 1


' MIXMASTER GLITTER DECAL ADJUSTMENT
' Prim_MixMasterStickers_ON.BlendDisableLighting = 5 ' Higher values make glitter texture on Molecular Mixmaster brighter
Prim_MixMasterStickers_OFF.BlendDisableLighting = 0

'CLEAR RAMP BRIGHTNESS ADJUSTMENT
pRamp.BlendDisableLighting = .75
pRamp_DT.BlendDisableLighting = .75

Dim VRRoom

If RenderingMode = 2 Then VRRoom = VRRoomChoice Else VRRoom = 0

'***************************************************************************************************************************************************************
'***************************************************************************************************************************************************************

Dim DesktopMode: DesktopMode = Table1.ShowDT

If DesktopMode = True Then 'Show Desktop components
  pRamp.visible = 0
  pRamp_DT.visible = 1
Else
  pRamp.visible = 1
  pRamp_DT.visible = 0
End If

' DMD Rotation
Const cDMDRotation  = -1      '-1 for No change, 0 - DMD rotation of 0?, 1 - DMD rotation of 90?
Const cGameName   = "dd_l2"   'ROM name
Const ballsize    = 50
Const ballmass    = 1
'Const DynamicFlipperFriction = True
'Const DynamicFlipperFrictionResting = 0.4
'Const DynamicFlipperFrictionActive = 1.0
Const usegi=0

LoadVPM "00990300", "s11.VBS",3.10
'SetLocale(1033)
'********************
'Standard definitions
'********************
Const UseSolenoids = 2
Const UseLamps = 0
Const UseSync = 1
Const HandleMech = 0

'Standard Sounds
Const SSolenoidOn = ""
Const SSolenoidOff = ""
Const SFlipperOn = ""
Const SFlipperOff = ""
Const SCoin = ""

Const tnob = 5 ' total number of balls
Const lob = 0 ' total number of locked balls

Dim tablewidth: tablewidth = Table1.width
Dim tableheight: tableheight = Table1.height
Dim BIPL : BIPL=false

'********************
'Table Init
'********************
Dim bsTrough, DTRBank1, ttcentre, mMagnet, SpinnerBall

Sub Table1_Init
  SetOptions
  vpmInit Me
  With Controller
        .GameName = cGameName
        .SplashInfoLine = "Dr. Dude (Bally 1990)"
    .HandleKeyboard = 0
    .ShowTitle = 0
    .ShowDMDOnly = 1
    .ShowFrame = 0
    .HandleMechanics = 1
    .Hidden = 1

    if cDMDRotation >= 0 Then .Games(cGameName).Settings.Value("rol") = cDMDRotation

    On Error Resume Next
    .Run GetPlayerHWnd
    If Err Then MsgBox Err.Description
  End With

    On Error Goto 0


  'Nudging
  vpmNudge.TiltSwitch=1
  vpmNudge.Sensitivity=5
  vpmNudge.TiltObj=Array(Bumper1,Bumper2,Bumper3,LeftSlingshot,RightSlingshot)

   'Trough
  Set bsTrough=New cvpmBallStack
        with bsTrough
      .InitSw 0,11,12,13,0,0,0,0
      .InitKick BallRelease,0,15
      .Balls=3
      '.InitExitSnd SoundFX("fx_ballrel",DOFContactors) , SoundFX("fx_Solenoid",DOFContactors)
    End With


  'DropTargets
  Set dtrbank1 = New cvpmDropTarget
  dtrbank1.InitDrop Array(sw21, sw22, sw23, sw24), Array(21,22,23,24)
  'dtrbank1.InitSnd "",SoundFX("Droptargetreset",DOFContactors)

  ' Disc
  Set ttcentre = new cvpmTurnTable
  ttcentre.InitTurnTable trMixMaster,50
  ttcentre.CreateEvents "ttcentre"
  ttcentre.SpinUp = 50
  ttcentre.SpinDown = 50

    '**Main Timer init
    PinMAMETimer.Enabled = 1

  Set SpinnerBall = SpinnerKick.CreateSizedballWithMass(40/2,Ballmass*2)
  SpinnerBall.visible = False
  Spinnerkick.kick 0,0,0
  Spinnerkick.enabled = False
  SpinnerBallTimer.enabled = True

  Options_Load

  If GIColorModType = 0 then
    'White
    GIWallWhite. visible = 1
    GIWallBlue. visible = 0
    GIWallPurple. visible = 0
    GIWallIceBlue. visible = 0
    GIWallWhite2. visible = 1
    GIWallBlue2. visible = 0
    GIWallPurple2. visible = 0
    GIWallIceBlue2. visible = 0
  End If


  If BumperColorType = 0 then
    'WhiteBumperColor
    GIWallWhite3. visible = 1
    GIWallBlue3. visible = 0
    GIWallPurple3. visible = 0
    GIWallIceBlue3. visible = 0
  End If


  If GIColorModType = 1 then
    'Blue
    GIWallWhite. visible = 0
    GIWallBlue. visible = 1
    GIWallPurple. visible = 0
    GIWallIceBlue. visible = 0
    GIWallWhite2. visible = 0
    GIWallBlue2. visible = 1
    GIWallPurple2. visible = 0
    GIWallIceBlue2. visible = 0
  End If


  If BumperColorType = 1 then
    'Blue
    GIWallWhite3. visible = 0
    GIWallBlue3. visible = 1
    GIWallPurple3. visible = 0
    GIWallIceBlue3. visible = 0
  End If

  If GIColorModType = 2 then
    'Purple
    GIWallWhite. visible = 0
    GIWallBlue. visible = 0
    GIWallPurple. visible = 1
    GIWallIceBlue. visible = 0
    GIWallWhite2. visible = 0
    GIWallBlue2. visible = 0
    GIWallPurple2. visible = 1
    GIWallIceBlue2. visible = 0
  End If

  If BumperColorType = 2 then
    'Purple
    GIWallWhite3. visible = 0
    GIWallBlue3. visible = 0
    GIWallPurple3. visible = 1
    GIWallIceBlue3. visible = 0
  End If


  If GIColorModType = 3 then
    'IceBlue
    GIWallWhite. visible = 0
    GIWallBlue. visible = 0
    GIWallPurple. visible = 0
    GIWallIceBlue. visible = 1
    GIWallWhite2. visible = 0
    GIWallBlue2. visible = 0
    GIWallPurple2. visible = 0
    GIWallIceBlue2. visible = 1
  End If

  If BumperColorType = 3 then
    'Ice Blue
    GIWallWhite3. visible = 0
    GIWallBlue3. visible = 0
    GIWallPurple3. visible = 0
    GIWallIceBlue3. visible = 1
  End If

  '***Magnet***
  Set RMag=New cvpmMagnet
  RMag.InitMagnet RMagnet, 10   'Magnet strength adjust
  RMag.GrabCenter = False


  SetGI 1 'Turn GI on at startup
    PrevGameOver = 0

End Sub


Sub B2SCommand(nr, state)
  If B2SOn Then
    Controller.B2SSetData nr, state
  End If
End Sub

Sub Table_Paused:Controller.Pause = 1:End Sub
Sub Table_unPaused:Controller.Pause = 0:End Sub
Sub Table1_Exit():Controller.Stop:End Sub

'********************
'   KEYS
'********************

Sub Table1_KeyDown(ByVal keycode)

  If bInOptions Then
    Options_KeyDown keycode
    Exit Sub
  End If

  If keycode = LeftMagnaSave Then
    If bOptionsMagna Then Options_Open() Else bOptionsMagna = True
    ElseIf keycode = RightMagnaSave Then
    If bOptionsMagna Then Options_Open() Else bOptionsMagna = True
  End If

  If keycode = 34 Then SolBigShot true 'G - key, testing

  If keycode = PlungerKey Then Plunger.PullBack:SoundPlungerPull
    If keycode = LeftFlipperKey Then Controller.Switch(58) = 1
    If keycode = RightFlipperKey Then Controller.Switch(57) = 1
    If keycode = LeftTiltKey Then Nudge 90,1: SoundNudgeLeft
    If keycode = RightTiltKey Then Nudge 270,1: SoundNudgeRight
    If keycode = CenterTiltKey Then Nudge 0,1: SoundNudgeCenter
  If keycode = StartGameKey Then SoundStartButton
  If VRRoom > 0 Then
    If keycode = RightFlipperKey Then
      Pincab_Button_Right.x = 2084.155 - 4
    End If
    If keycode = LeftFlipperKey Then
      Pincab_Button_Left.x = 2117.569 + 4
    End If
    If keycode = StartGameKey Then
      Pincab_StartButton.y = 746.6109 - 2
    End If
    If keycode = PlungerKey Then
      TimerVRPlunger.Enabled = True
      TimerVRPlunger2.Enabled = False
    End If
    If VRRoom = 1 Then
      If keycode = LeftMagnaSave Then VRFlyerPrim.image = "VRFlyer"
      If keycode = RightMagnaSave Then VRFlyerPrim.image = "VRFlyerback"
    End If
  End If

    If keycode = LeftFlipperKey Then FlipperActivate LeftFlipper, LFPress
    If keycode = RightFlipperKey Then FlipperActivate RightFlipper, RFPress

  If keycode = keyInsertCoin1 or keycode = keyInsertCoin2 or keycode = keyInsertCoin3 or keycode = keyInsertCoin4 Then
    Select Case Int(rnd*3)
      Case 0: PlaySound ("Coin_In_1"), 0, CoinSoundLevel, 0, 0.25
      Case 1: PlaySound ("Coin_In_2"), 0, CoinSoundLevel, 0, 0.25
      Case 2: PlaySound ("Coin_In_3"), 0, CoinSoundLevel, 0, 0.25
    End Select
  End If

'CP test Key
  If keycode = 22 then  ''''''''''''''''''''U Key used for testing
    NudgeStrangthY = 13
    ShakeXMax = 9
    ShakeXMin = -3
    ShakeXDir = -1
    ShakeYMax = 1
    ShakeYMin = -1
    ShakeYDir = -1
    ShakeX.Enabled = True
    ShakeY.Enabled = True
  End If

'   '************************   Start Ball Control 1/3
'        if keycode = 46 then                ' C Key
'            If contball = 1 Then
'                contball = 0
'            Else
'                contball = 1
'            End If
'        End If
'        if keycode = 48 then                'B Key
'            If bcboost = 1 Then
'                bcboost = bcboostmulti
'            Else
'                bcboost = 1
'            End If
'        End If
'        if keycode = 203 then bcleft = 1        ' Left Arrow
'        if keycode = 200 then bcup = 1          ' Up Arrow
'        if keycode = 208 then bcdown = 1        ' Down Arrow
'        if keycode = 205 then bcright = 1       ' Right Arrow
'    '************************   End Ball Control 1/3

    If vpmKeyDown(keycode) Then Exit Sub
End Sub


Sub Table1_KeyUp(ByVal keycode)

    If keycode = LeftMagnaSave And Not bInOptions Then bOptionsMagna = False
    If keycode = RightMagnaSave And Not bInOptions Then bOptionsMagna = False

  If KeyCode = PlungerKey Then
    Plunger.Fire
    If BIPL = 1 Then
      SoundPlungerReleaseBall()     'Plunger release sound when there is a ball in shooter lane
    Else
      SoundPlungerReleaseNoBall()     'Plunger release sound when there is no ball in shooter lane
    End If
  End If

    If keycode = LeftFlipperKey Then Controller.Switch(58) = 0
    If keycode = RightFlipperKey Then Controller.Switch(57) = 0

    If keycode = LeftFlipperKey Then FlipperDeActivate LeftFlipper, LFPress
    If keycode = RightFlipperKey Then FlipperDeActivate RightFlipper, RFPress
  If VRRoom > 0 Then
    If keycode = PlungerKey Then
      TimerVRPlunger.Enabled = False
      TimerVRPlunger2.Enabled = True
    End If
    If keycode = RightFlipperKey Then
      Pincab_Button_Right.x = 2084.155
    End If
    If keycode = LeftFlipperKey Then
      Pincab_Button_Left.x = 2117.569
    End If
    If keycode = StartGameKey Then
      Pincab_StartButton.y = 746.6109
    End If
  End If

'    '************************   Start Ball Control 2/3
'    if keycode = 203 then bcleft = 0        ' Left Arrow
'    if keycode = 200 then bcup = 0          ' Up Arrow
'    if keycode = 208 then bcdown = 0        ' Down Arrow
'    if keycode = 205 then bcright = 0       ' Right Arrow
'    '************************   End Ball Control 2/3

    If vpmKeyUp(keycode) Then Exit Sub
End Sub


''************************   Start Ball Control 3/3
'Sub StartControl_Hit()
'    Set ControlBall = ActiveBall
'    contballinplay = true
'End Sub
'
'Sub StopControl_Hit()
'    contballinplay = false
'End Sub
'
'Dim bcup, bcdown, bcleft, bcright, contball, contballinplay, ControlBall, bcboost
'Dim bcvel, bcyveloffset, bcboostmulti
'
'bcboost = 1     'Do Not Change - default setting
'bcvel = 4       'Controls the speed of the ball movement
'bcyveloffset = -0.01    'Offsets the force of gravity to keep the ball from drifting vertically on the table, should be negative
'bcboostmulti = 3    'Boost multiplier to ball veloctiy (toggled with the B key)
'
'Sub BallControl_Timer()
'    If Contball and ContBallInPlay then
'        If bcright = 1 Then
'            ControlBall.velx = bcvel*bcboost
'        ElseIf bcleft = 1 Then
'            ControlBall.velx = - bcvel*bcboost
'        Else
'            ControlBall.velx=0
'        End If
'
'        If bcup = 1 Then
'            ControlBall.vely = -bcvel*bcboost
'        ElseIf bcdown = 1 Then
'            ControlBall.vely = bcvel*bcboost
'        Else
'            ControlBall.vely= bcyveloffset
'        End If
'    End If
'End Sub
''************************   End Ball Control 3/3


'*********** SOLENOIDS ************

SolCallBack(1)  = "bsTrough.SolIn"
SolCallBack(2)  = "SolOutbsTrough"
SolCallBack(3)  = "solLeftKicker"
SolCallBack(4)  = "solRightKicker"
SolCallBack(6)  = "vpmSolSound SoundFX(""Knocker_1"",DOFKnocker),"
SolCallBack(7)  = "ResetDrops"
'SolCallBack(10) = "solGI"
SolCallBack(10) = "SetGI" 'setlamp 110
SolCallback(11) = "Sol11"
SolCallBack(13) = "RMag.MagnetOn="
SolCallBack(14) = "SolBigShot"
SolCallBack(15) = "SetLamp 115,"        'Big Guy Flasher
SolCallBack(16) = "solMixer"          'Mixer Motor
SolCallBack(25) = "SetLamp 125,"        'Mixer Heart Flasher
SolCallBack(26) = "SetLamp 126,"        'Mixer Gab Flasher
SolCallBack(27) = "SetLamp 127,"        'Mixer Magnet Flasher
SolCallBack(28) = "SetLamp 128,"        'Magnet Flasher
SolCallBack(29) = "SetLamp 129,"        'Gab Flasher
SolCallBack(30) = "SetLamp 130,"        'Heart Flasher
SolCallBack(31) = "SetLamp 131,"        'Drop Targets Flasher
SolCallBack(32) = "SetLamp 132,"        'Raygun Flasher
SolCallback(sLLFlipper) = "SolLFlipper"
SolCallback(sLRFlipper) = "SolRFlipper"
'SolCallback(23) = "FastFlips.TiltSol"  'handled by core.vbs now



'***********************************************************************************
'Flippers
'***********************************************************************************
Const ReflipAngle = 20

Sub SolLFlipper(Enabled)
  If Enabled Then
    LF.Fire  'leftflipper.rotatetoend

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

Sub SolRFlipper(Enabled)
  If Enabled Then
    RF.Fire 'rightflipper.rotatetoend

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

' Flipper collide subs
Sub LeftFlipper_Collide(parm)
  CheckLiveCatch Activeball, LeftFlipper, LFCount, parm
  LeftFlipperCollide parm
End Sub

Sub RightFlipper_Collide(parm)
  CheckLiveCatch Activeball, RightFlipper, RFCount, parm
  RightFlipperCollide parm
End Sub


Sub SolOutbsTrough(Enabled)
  bsTrough.SolOut Enabled
  If Enabled Then RandomSoundBallRelease BallRelease
End Sub


Sub solLeftKicker (enabled)
  If (enabled and Controller.Switch(51)) Then
    SoundSaucerKick 1,sw51
'   PlaysoundAtVol SoundFX("fx_vuk_exit2",DOFContactors), sw51, .8
'   Controller.Switch(51) = 0
    sw51.timerenabled = 1
'   sw51.kick 0,50, 3.14/2
  End If
End Sub


Dim Rmag

Sub RMagnet_Hit
  RMag.AddBall ActiveBall
  If RMag.MagnetOn=true then
  End If
End Sub

Sub RMagnet_UnHit:RMag.RemoveBall ActiveBall:End Sub

' Ski Jump Prevention Committee
Sub BallDrop1_Hit
  activeball.velx = 0
  activeball.vely = 0
  activeball.velz = 0
End Sub

Sub BallDrop2_Hit
  activeball.velx = 0
  activeball.vely = 0
  activeball.velz = 0
End Sub

'**************************
' Big Shot Shakes (apophis)
'**************************
const WobbleMax = 10
const WobbleMin = 0.2
const WobbleDelta = 70   'high is faster freq
dim WobbleDecay: WobbleDecay = 0.85  'less than 1
dim WobbleAmp: WobbleAmp=0
dim WobbleAng: WobbleAng=0
dim WobblePhase: WobblePhase=0
dim Wobble: Wobble=0

dim BigshotStart: BigshotStart=False


Sub BigshotUpdate_Timer
  If BigshotStart Then
    FlipperBigshot.RotateToEnd
    BigshotStart = False
  ElseIf FlipperBigshot.CurrentAngle >= FlipperBigshot.EndAngle-1 Then
    FlipperBigshot.RotateToStart
    WobbleAmp = WobbleMax
    WobblePhase = rnd*360
    WobbleDecay = 0.80 + rnd*0.1
    WobbleAng = 0
  Else
    Wobble = WobbleAmp*dSin(WobbleAng + WobblePhase)
    WobbleAng = WobbleAng + WobbleDelta
    If WobbleAng >= 360 Then WobbleAng = WobbleAng-360
    WobbleAmp = WobbleDecay*WobbleAmp
    If WobbleAmp < WobbleMin Then
      Wobble = 0
      BigshotUpdate.Enabled = False
    End If
  End If
  pBigshot.Rotx = FlipperBigshot.CurrentAngle
  pBigshot.Roty = Wobble
End Sub


Sub SolBigShot (enabled)
  If Enabled Then
    SoundSaucerKick 1,pBigshot
    BigshotStart = True
    BigshotUpdate.Enabled = True
  End If
End Sub




'**********  End Big Shot Shake ****************



Dim sw51step
Sub sw51_timer()  ' I have a kickass keyframe animation script for this kind of thing if you want it -nf
  Select Case sw51step
    Case 0:Leftupkicker.TransY = 10:SoundSaucerKick 1, sw51
    Case 1:Leftupkicker.TransY = 20:If BallInKicker2 = 1 then BallSaucer2.velz = 45 End If  'Kicker strength 45
    Case 2:Leftupkicker.TransY = 30
    Case 3:'pUpKicker.TransY = 30
    Case 4:
    Case 5:Leftupkicker.TransY = 25
    Case 6:Leftupkicker.TransY = 20
    Case 7:Leftupkicker.TransY = 15:Controller.Switch(51) = 0:BallInKicker2 = 0
    Case 8:Leftupkicker.TransY = 10
    Case 9:Leftupkicker.TransY = 5
    Case 10:Leftupkicker.TransY = 0:sw51.timerEnabled = 0:sw51step = 0
  End Select
  sw51step = sw51step + 1
End Sub

Sub solRightKicker (enabled)
  If (enabled) Then


    sw32.timerenabled = 1
'   sw32.kick 0,60, 3.14/2
  End If
End Sub

Dim sw32step
Sub sw32_timer()
  Select Case sw32step
    Case 0:Rightupkicker.TransY = 10: SoundSaucerKick 1, sw32
    Case 1:Rightupkicker.TransY = 20:If BallInKicker1 = 1 then BallSaucer1.velz = 45 End If 'Kicker strength 45
    Case 2:Rightupkicker.TransY = 30
    Case 3:'pUpKicker.TransY = 30
    Case 4:
    Case 5:Rightupkicker.TransY = 25
    Case 6:Rightupkicker.TransY = 20
    Case 7:Rightupkicker.TransY = 15: Controller.Switch(32) = 0:BallInKicker1 = 0
    Case 8:Rightupkicker.TransY = 10
    Case 9:Rightupkicker.TransY = 5
    Case 10:Rightupkicker.TransY = 0:sw32.timerEnabled = 0:sw32step = 0
  End Select
  sw32step = sw32step + 1
End Sub



'**************************
'Mixmaster motor
'**************************

Dim MixerOn

Sub SolMixer(enabled)
  if enabled then
    ttcentre.MotorOn = True
    MixerOn = True
    Playsound SoundFX("fx_motor",DOFGear), -1, .5, AudioPan(trMixMaster), 0, 0, 1, 0,AudioFade(trMixMaster)
  else
    ttcentre.MotorOn = False
    MixerOn = False
    StopSound "fx_motor"
  end if
end sub


'**************************************************************************************************

Dim discPosition, discSpinSpeed, discLastPos, SpinCounter, maxvel
dim spinAngle, degAngle, startAngle, postSpeedFactor
dim discX, discY

startAngle = 120
discX = 674.55
discY = 428.65

Const cDiscSpeed = 1750
Const cDiscFriction = 10 '2.0          ' Friction coefficient (deg/sec/sec)
Const cDiscMinSpeed = 1               ' Object stops at this speed (deg/sec)
Const cDiscRadius = 80
postSpeedFactor = cDiscSpeed / 15

SpinnerBallTimer.interval=10

Sub SpinnerBallTimer_Timer()

  discPosition = discPosition + discSpinSpeed * Me.Interval / 1200 'disk speed

  If MixerOn = True Then
    discspinspeed = cDiscSpeed
  Else
    discSpinSpeed = discSpinSpeed * (1 - cDiscFriction * Me.Interval / 1200) 'disk speed
  End If

  Do While discPosition < 0 : discPosition = discPosition + 360 : Loop
  Do While discPosition > 360 : discPosition = discPosition - 360 : Loop

  If Abs(discSpinSpeed) < cDiscMinSpeed Then
    discSpinSpeed = 0
  End If

  degAngle = discPosition


  spinAngle = PI * (degAngle) / 180

'***Spinner Debug***
' debug.print Spinnerball.x
' debug.print Spinnerball.y
' debug.print discX
' debug.print discY
' debug.print cDiscRadius
' debug.print Cos(spinAngle)
' debug.print Sin(spinAngle)

  SpinnerBall.x = discX + (cDiscRadius * Cos(spinAngle))
  SpinnerBall.y = discY + (cDiscRadius * Sin(spinAngle))
  SpinnerBall.z = 107.353

  'pdisc1.objrotz = discPosition
    PostMM.RotZ=discPosition
  SpinningDisc.ObjRotZ = discPosition


  If ABS(discSpinSpeed*sin(spinAngle)/postSpeedFactor) < 0.05 Then
    SpinnerBall.velx = 0.05
  Else
    SpinnerBall.velx = - discSpinSpeed*sin(spinAngle)/postSpeedFactor
  End If

  If Abs(discSpinSpeed*cos(spinAngle)/postSpeedFactor) < 0.05 Then
    SpinnerBall.vely = 0.05
  Else
    SpinnerBall.vely = discSpinSpeed*cos(spinAngle)/postSpeedFactor   '0.05
  End If

  SpinnerBall.velz = 0

End Sub


Sub Wall100_Hit
  Wall100Cnt = Wall100Cnt + 1
  If Wall100Cnt >= Wall100Max Then Wall100.IsDropped = True
  'debug.print Wall100Cnt & " , " & Wall100Max
End Sub

Sub Drain_Hit()
  BallSearch
  if uBound(getballs) + bstrough.balls > 5 then me.destroyball : exit sub : end if
  vpmTimer.PulseSw 10
  bsTrough.AddBall Me
  RandomSoundDrain Drain
End Sub

Sub BallSearch()  'find balls that have fallen off the table
  dim x : for each x in getballs : if x.y > 2200 then x.x = 155 : x.y = 550 : x.vely = 0 : x.velx = 0 :end if:  next
End Sub


'****************
'  POP BUMPERS
'****************

Dim Bumper1Cnt,Bumper2Cnt,Bumper3Cnt
Const BumperMoveMax = 20
Sub Bumper1_Hit
        vpmTimer.PulseSw 52
        RandomSoundBumperTop Bumper1
        Bumper1Hit
End Sub
Sub Bumper1_Timer: Bumper1Move: End Sub
Sub Bumper1Hit
    Bumper1Cnt = 0                  'Reset count
    Bumper1.TimerInterval = 5  'Set timer interval
    Bumper1.TimerEnabled = 1    'Enable timer
End Sub

Sub Bumper1Move
    Select Case Bumper1Cnt
        Case 0:     Bumper1Cap.TransY = -BumperMoveMax * .25 :   Bumper1CapS1.TransZ = -BumperMoveMax * .25 :Bumper1CapS2.TransZ = -BumperMoveMax * .25 :
        Case 1:     Bumper1Cap.TransY = -BumperMoveMax * .50 :   Bumper1CapS1.TransZ = -BumperMoveMax * .50 :Bumper1CapS2.TransZ = -BumperMoveMax * .50 :
        Case 2:     Bumper1Cap.TransY = -BumperMoveMax * .75 :   Bumper1CapS1.TransZ = -BumperMoveMax * .75 :Bumper1CapS2.TransZ = -BumperMoveMax * .75 :
        Case 3:     Bumper1Cap.TransY = -BumperMoveMax        :  Bumper1CapS1.TransZ = -BumperMoveMax :Bumper1CapS2.TransZ = -BumperMoveMax :
        Case 4:     Bumper1Cap.TransY = -BumperMoveMax * .25 :   Bumper1CapS1.TransZ = -BumperMoveMax * .25 :Bumper1CapS2.TransZ = -BumperMoveMax * .25 :
        Case 5:     Bumper1Cap.TransY = -BumperMoveMax * .50 :   Bumper1CapS1.TransZ = -BumperMoveMax * .50 :Bumper1CapS2.TransZ = -BumperMoveMax * .50 :
        Case 6:     Bumper1Cap.TransY = -BumperMoveMax * .75 :   Bumper1CapS1.TransZ = -BumperMoveMax * .75 :Bumper1CapS2.TransZ = -BumperMoveMax * .75 :
        Case 7:     Bumper1Cap.TransY = 0 : Bumper1CapS1.TransZ = 0 :Bumper1CapS2.TransZ = 0 :
        Case else:  Bumper1.TimerEnabled = 0
    End Select
    Bumper1Cnt = Bumper1Cnt + 1
End Sub


Sub Bumper2_Hit
  vpmTimer.PulseSw 53
  RandomSoundBumperMiddle Bumper2
  Bumper2Hit
End Sub
Sub Bumper2_Timer: Bumper2Move: End Sub
Sub Bumper2Hit
    Bumper2Cnt = 0                  'Reset count
    Bumper2.TimerInterval = 5  'Set timer interval
    Bumper2.TimerEnabled = 1    'Enable timer
End Sub

Sub Bumper2Move
    Select Case Bumper2Cnt
        Case 0:     Bumper2Cap.TransY = -BumperMoveMax * .25 :   Bumper2CapS1.TransZ = -BumperMoveMax * .25 :Bumper2CapS2.TransZ = -BumperMoveMax * .25 :
        Case 1:     Bumper2Cap.TransY = -BumperMoveMax * .50 :   Bumper2CapS1.TransZ = -BumperMoveMax * .50 :Bumper2CapS2.TransZ = -BumperMoveMax * .50 :
        Case 2:     Bumper2Cap.TransY = -BumperMoveMax * .75 :   Bumper2CapS1.TransZ = -BumperMoveMax * .75 :Bumper2CapS2.TransZ = -BumperMoveMax * .75 :
        Case 3:     Bumper2Cap.TransY = -BumperMoveMax        :  Bumper2CapS1.TransZ = -BumperMoveMax :Bumper2CapS2.TransZ = -BumperMoveMax :
        Case 4:     Bumper2Cap.TransY = -BumperMoveMax * .25 :   Bumper2CapS1.TransZ = -BumperMoveMax * .25 :Bumper2CapS2.TransZ = -BumperMoveMax * .25 :
        Case 5:     Bumper2Cap.TransY = -BumperMoveMax * .50 :   Bumper2CapS1.TransZ = -BumperMoveMax * .50 :Bumper2CapS2.TransZ = -BumperMoveMax * .50 :
        Case 6:     Bumper2Cap.TransY = -BumperMoveMax * .75 :   Bumper2CapS1.TransZ = -BumperMoveMax * .75 :Bumper2CapS2.TransZ = -BumperMoveMax * .75 :
        Case 7:     Bumper2Cap.TransY = 0 : Bumper2CapS1.TransZ = 0 :Bumper2CapS2.TransZ = 0 :
        Case else:  Bumper2.TimerEnabled = 0
    End Select
    Bumper2Cnt = Bumper2Cnt + 1
End Sub


Sub Bumper3_Hit
    vpmTimer.PulseSw 54
    RandomSoundBumperBottom Bumper3
    Bumper3Hit
End Sub
Sub Bumper3_Timer: Bumper3Move: End Sub
Sub Bumper3Hit
    Bumper3Cnt = 0                  'Reset count
    Bumper3.TimerInterval = 5  'Set timer interval
    Bumper3.TimerEnabled = 1    'Enable timer
End Sub

Sub Bumper3Move
    Select Case Bumper3Cnt
        Case 0:     Bumper3Cap.TransY = -BumperMoveMax * .25 :   Bumper3CapS1.TransZ = -BumperMoveMax * .25 :Bumper3CapS2.TransZ = -BumperMoveMax * .25 :
        Case 1:     Bumper3Cap.TransY = -BumperMoveMax * .50 :   Bumper3CapS1.TransZ = -BumperMoveMax * .50 :Bumper3CapS2.TransZ = -BumperMoveMax * .50 :
        Case 2:     Bumper3Cap.TransY = -BumperMoveMax * .75 :   Bumper3CapS1.TransZ = -BumperMoveMax * .75 :Bumper3CapS2.TransZ = -BumperMoveMax * .75 :
        Case 3:     Bumper3Cap.TransY = -BumperMoveMax        :  Bumper3CapS1.TransZ = -BumperMoveMax :Bumper3CapS2.TransZ = -BumperMoveMax :
        Case 4:     Bumper3Cap.TransY = -BumperMoveMax * .25 :   Bumper3CapS1.TransZ = -BumperMoveMax * .25 :Bumper3CapS2.TransZ = -BumperMoveMax * .25 :
        Case 5:     Bumper3Cap.TransY = -BumperMoveMax * .50 :   Bumper3CapS1.TransZ = -BumperMoveMax * .50 :Bumper3CapS2.TransZ = -BumperMoveMax * .50 :
        Case 6:     Bumper3Cap.TransY = -BumperMoveMax * .75 :   Bumper3CapS1.TransZ = -BumperMoveMax * .75 :Bumper3CapS2.TransZ = -BumperMoveMax * .75 :
        Case 7:     Bumper3Cap.TransY = 0 : Bumper3CapS1.TransZ = 0 :Bumper3CapS2.TransZ = 0 :
        Case else:  Bumper3.TimerEnabled = 0
    End Select
    Bumper3Cnt = Bumper3Cnt + 1
End Sub


'**********Sling Shots and Animations
' Rstep and Lstep  are the variables that increment the animation
'****************
Dim RStep, Lstep

Sub LeftSlingShot_Slingshot
  LS.VelocityCorrect(ActiveBall)
    RandomSoundSlingshotLeft Sling2
    LSling.Visible = 0
    LSling2.Visible = 1
    sling2.TransZ = -28
    LStep = 0
    LeftSlingShot.TimerEnabled = 1
  vpmTimer.PulseSw 55
  'gi1.State = 0:Gi2.State = 0
End Sub

Sub LeftSlingShot_Timer
    Select Case LStep
        Case 1:LSLing2.Visible = 0:LSLing1.Visible = 1:sling2.TransZ = -10
        Case 2:LSLing1.Visible = 0:LSLing.Visible = 1:sling2.TransZ = 0:LeftSlingShot.TimerEnabled = 0:'gi1.State = 1:Gi2.State = 1
    End Select
    LStep = LStep + 1
End Sub

Sub RightSlingShot_Slingshot
  RS.VelocityCorrect(ActiveBall)
    RandomSoundSlingshotRight Sling1
    RSling.Visible = 0
    RSling2.Visible = 1
    sling1.TransZ = -28
    RStep = 0
    RightSlingShot.TimerEnabled = 1
  vpmTimer.PulseSw 56
  'gi1.State = 0:Gi2.State = 0
End Sub

Sub RightSlingShot_Timer
    Select Case RStep
        Case 1:RSLing1.Visible = 1:RSLing2.Visible = 0:sling1.TransZ = -10
        Case 2:RSLing1.Visible = 0:RSLing.Visible = 1:sling1.TransZ = 0:RightSlingShot.TimerEnabled = 0:'gi1.State = 1:Gi2.State = 1
    End Select
    RStep = RStep + 1
End Sub


Sub wall83_Hit:vpmTimer.PulseSw 100:rubberback1.visible = 0::rubberback1a.visible = 1:wall83.timerenabled = 1:End Sub
Sub wall83_timer:rubberback1.visible = 1::rubberback1a.visible = 0: wall83.timerenabled= 0:End Sub

Sub Wall2_Hit:vpmTimer.PulseSw 100:MMrubberl.visible = 0::MMrubberla.visible = 1:Wall2.timerenabled = 1:End Sub
Sub Wall2_timer:MMrubberl.visible = 1::MMrubberla.visible = 0: Wall2.timerenabled= 0:End Sub

Sub Wall3_Hit:vpmTimer.PulseSw 100:MMrubberR.visible = 0::MMrubberRa.visible = 1:Wall3.timerenabled = 1:End Sub
Sub Wall3_timer:MMrubberR.visible = 1::MMrubberRa.visible = 0: Wall3.timerenabled= 0:End Sub


Sub Wall233_Hit:vpmTimer.PulseSw 100:Rubber_white_1.visible = 0::Rubber_white_1a.visible = 1:Wall233.timerenabled = 1:End Sub
Sub Wall233_timer:Rubber_white_1.visible = 1::Rubber_white_1a.visible = 0: Wall233.timerenabled= 0:End Sub

Sub Wall101_Hit:vpmTimer.PulseSw 100:Rubber_white_3.visible = 0::Rubber_white_3a.visible = 1:Wall101.timerenabled = 1: vpmTimer.PulseSw 39: End Sub
Sub Wall101_timer:Rubber_white_3.visible = 1::Rubber_white_3a.visible = 0: Wall101.timerenabled= 0:End Sub

Sub Wall32_Hit:vpmTimer.PulseSw 100:Rubber_white_15.visible = 0::Rubber_white_17.visible = 1:Wall32.timerenabled = 1:End Sub
Sub Wall32_timer:Rubber_white_15.visible = 1::Rubber_white_17.visible = 0: Wall32.timerenabled= 0:End Sub



' ************************
'      RealTime Updates
' ************************

Sub GameTimer_Timer
  Cor.Update
  RollingSoundUpdate
  Options_UpdateDMD
End Sub


Sub FrameTimer_Timer

  If Flipperstyle = 0 Then
    BatLeftYellow.objrotz = LeftFlipper.CurrentAngle
  Else
    BatLeftWhite.objrotz = LeftFlipper.CurrentAngle
  End If

  batleftshadow.objrotz = LeftFlipper.CurrentAngle

  If Flipperstyle = 0 Then
    BatRightYellow.objrotz = RightFlipper.CurrentAngle
  Else
    BatRightWhite.objrotz = RightFlipper.CurrentAngle
  End If

  batrightshadow.objrotz = RightFlipper.CurrentAngle

  UpdateGatesSpinners
  If VRRoom > 0 Then
    DisplayTimerVR
    Select Case BGSeq
      Case 300:BGDude.imageA = "1"
      Case 303:BGDude.imageA = "2"
      Case 306:BGDude.imageA = "3"
      Case 309:BGDude.imageA = "4"
      Case 312:BGDude.imageA = "5"
      Case 315:BGDude.imageA = "6"
      Case 900:BGDude.imageA = "5"
      Case 903:BGDude.imageA = "4"
      Case 906:BGDude.imageA = "3"
      Case 909:BGDude.imageA = "2"
      Case 912:BGDude.imageA = "1"
    End Select
    BGSeq = BGSeQ + 1
    If BGSeq > 913 Then
      BGSeq = 1
    End if
  End If

  If DynamicBallShadowsOn Or AmbientBallShadowOn Then DynamicBSUpdate 'update ball shadows
End Sub


Sub ramphelper_Hit()
  ActiveBall.vely = Activeball.vely*1.2
End Sub

Sub ramphelper1_Hit()
  ActiveBall.vely = Activeball.vely*1.3
End Sub

Sub ramphelper2_Hit()
  ActiveBall.vely = Activeball.vely*1.2
End Sub

Sub ramphelper3_Hit()
  ActiveBall.vely = Activeball.vely*1.1
End Sub

Sub ramphelper4_Hit()
  ActiveBall.vely = Activeball.vely*1.1
End Sub





' ************************
'      Gates
' ***********************
Dim GateSpeed
GateSpeed = 0.5


Sub UpdateGatesSpinners
Gate1P.Rotx = Gate1.currentangle *-1 '*-1 reverses direction
Gate2P.Rotx = Gate2.currentangle *-1 -10
End Sub


' ************************
'      Switches
' ************************

sub sw9_hit:   controller.switch(9)=1: BIPL=true: StopSound "zIntro":  end sub
sub sw9_unhit: controller.switch(9)=0
  BIPL=false
    If ActiveBall.VelY < 0 Then 'on the way up
    End If
End Sub

Sub sw14_Hit:vpmTimer.PulseSw(14):Me.TimerEnabled = 1: End Sub
Sub sw15_Hit:vpmTimer.PulseSw(15) : WireRampOn True :End Sub
Sub sw16_hit:   vpmTimer.PulseSw(16): end sub
Sub sw17_hit:   controller.switch(17)=1: end sub
Sub sw17_unhit: controller.switch(17)=0: end sub
Sub sw18_hit:   controller.switch(18)=1: end sub
Sub sw18_unhit: controller.switch(18)=0: end sub
Sub sw19_hit:   controller.switch(19)=1: end sub
Sub sw19_unhit: controller.switch(19)=0: end sub
Sub sw20_hit:   controller.switch(20)=1: end sub
Sub sw20_unhit: controller.switch(20)=0: end sub

Sub sw21_hit: dtrbank1.hit 1: SoundDropTargetDrop sw21:DoubleDrop1.isdropped = true: end sub
Sub DoubleDrop1_HIt:dtrbank1.hit 1:dtrbank1.hit 2:SoundDropTargetDrop sw21:DoubleDrop1.isdropped = true: End Sub
Sub sw22_hit: dtrbank1.hit 2: SoundDropTargetDrop sw22:DoubleDrop1.isdropped = true:DoubleDrop2.isdropped = true: end sub
Sub DoubleDrop2_HIt:dtrbank1.hit 2:dtrbank1.hit 3:SoundDropTargetDrop sw22:DoubleDrop2.isdropped = true: End Sub
Sub sw23_hit: dtrbank1.hit 3: SoundDropTargetDrop sw23:DoubleDrop2.isdropped = true:DoubleDrop3.isdropped = true: end sub
Sub DoubleDrop3_HIt:dtrbank1.hit 3:dtrbank1.hit 4:SoundDropTargetDrop sw23:DoubleDrop3.isdropped = true: End Sub
Sub sw24_hit: dtrbank1.hit 4: SoundDropTargetDrop sw24:DoubleDrop3.isdropped = true: end sub

Sub ResetDrops(Enabled)
  If Enabled Then
    RandomSoundDropTargetReset sw22
    DoubleDrop1.isdropped = False:DoubleDrop2.isdropped = False:DoubleDrop3.isdropped = False
    dtRBank1.DropSol_On
  End if
End Sub

Sub sw25_Hit:vpmTimer.PulseSw(25):Me.TimerEnabled = 1: End Sub
Sub sw26_Hit:vpmTimer.PulseSw(26):Me.TimerEnabled = 1: End Sub
Sub sw27_Hit:vpmTimer.PulseSw(27):Me.TimerEnabled = 1: End Sub
Sub sw28_Hit:vpmTimer.PulseSw(28):Me.TimerEnabled = 1: End Sub
Sub sw29_Hit:vpmTimer.PulseSw(29):Me.TimerEnabled = 1: End Sub
Sub sw30_Hit:vpmTimer.PulseSw(30):Me.TimerEnabled = 1: End Sub
Sub sw31_Hit:vpmTimer.PulseSw(31):Me.TimerEnabled = 1: End Sub
Sub sw32_Hit:SoundSaucerLock:End Sub
Sub sw33_Hit:vpmTimer.PulseSw(33):Me.TimerEnabled = 1: End Sub
Sub sw34_Hit:vpmTimer.PulseSw(34):Me.TimerEnabled = 1: End Sub
Sub sw35_Hit:vpmTimer.PulseSw(35):Me.TimerEnabled = 1: End Sub
Sub sw36_Hit:vpmTimer.PulseSw(36):Me.TimerEnabled = 1: End Sub
Sub sw37_Hit:vpmTimer.PulseSw(37):Me.TimerEnabled = 1: End Sub
Sub sw38_Hit:vpmTimer.PulseSw(38):Me.TimerEnabled = 1: End Sub
Sub sw41_Hit:vpmTimer.PulseSw(41):Me.TimerEnabled = 1: End Sub
Sub sw42_Hit:vpmTimer.PulseSw(42):Me.TimerEnabled = 1: End Sub
Sub sw43_Hit:vpmTimer.PulseSw(43):Me.TimerEnabled = 1: End Sub
sub sw46_hit:vpmTimer.PulseSw 46: end sub
sub sw47_hit:vpmTimer.PulseSw 47: end sub
sub sw48_hit:vpmTimer.PulseSw 48: end sub
Sub sw49_Hit:vpmTimer.PulseSw(49):Me.TimerEnabled = 1: End Sub
Sub sw50_Hit:vpmTimer.PulseSw(50):Me.TimerEnabled = 1: End Sub
Sub sw51_Hit:SoundSaucerLock:End Sub
sub sw59_hit:   controller.switch(59)=1: end sub
sub sw59_unhit: controller.switch(59)=0: end sub






'***************************************
'***Prim Image Swaps***
'***************************************
Dim DomeImageGreen: DomeImageGreen = Array("DomeOffGreen", "DomeOffGreen", "DomeOnGreen", "DomeOnGreen")
Dim DomeImageRed: DomeImageRed = Array("DomeOffRed", "DomeOffRed", "DomeOnRed", "DomeOnRed")
Dim DomeImageYellow: DomeImageYellow = Array("DomeOffYellow", "DomeOffYellow", "DomeOnYellow", "DomeOnYellow")
Dim DomeImageCap: DomeImageCap = Array("BumperCapOff", "BumperCapOff", "BumperCapOn", "BumperCapOn")
Dim GOGImage: GOGImage = Array("GOGOff", "GOGOff", "GOGOn", "GOGOn")

'***************************************
'***Prim Material Swaps***
'***************************************
Dim DomeMatGreen: DomeMatGreen = Array("DomeGreenOff", "DomeGreenOff", "DomeGreenOn", "DomeGreenOn")
Dim RayTip: RayTip = Array("Excellent Ray1", "Excellent Ray1", "Excellent Ray3", "Excellent Ray3")

Dim DLintensity
Sub MatSwap(pri, group, DLintensity, ByVal aLvl)  'cp's script  DLintensity = disabled lighting intesity
  if Lampz.UseFunction then aLvl = Lampz.FilterOut(aLvl)  'Callbacks don't get this filter automatically
    Select case FlashLevelToIndex(aLvl, 3)
    Case 1:pri.Material = group(0) 'Full
    Case 2:pri.Material = group(1) 'Fading...
    Case 3:pri.Material = group(2) 'Fading...
              Case 4:pri.Material = group(3) 'Off
    End Select
pri.blenddisablelighting = aLvl * DLintensity
End Sub


Sub ImageSwap(pri, group, DLintensity, ByVal aLvl)  'cp's script  DLintensity = disabled lighting intesity
  if Lampz.UseFunction then aLvl = Lampz.FilterOut(aLvl)  'Callbacks don't get this filter automatically
    Select case FlashLevelToIndex(aLvl, 3)
    Case 1:pri.Image = group(0) 'Full
    Case 2:pri.Image = group(1) 'Fading...
    Case 3:pri.Image = group(2) 'Fading...
        Case 4:pri.Image = group(3) 'Off
    End Select
pri.blenddisablelighting = aLvl * DLintensity
End Sub


Sub DisableLighting(pri, DLintensity, ByVal aLvl) 'cp's script  DLintensity = disabled lighting intesity
  if Lampz.UseFunction then aLvl = Lampz.FilterOut(aLvl)  'Callbacks don't get this filter automatically

pri.blenddisablelighting = aLvl * DLintensity
End Sub



'*********************************************************************************************************************************************************
'Begin nfozzy lamp handling
'*********************************************************************************************************************************************************

redim GILamps(99) : redim GIFlashers(99)  'new arrays
SortGI GILamps, GIFlashers, GILighting
dim TestString, TestStringAll 'debug strings
Sub SortGI(ByRef aLight,aFlasher, GImixed) 'different method using Arrays instead of scripting dictionary objects
  dim x, CountMe: CountMe = 0
  for x = 0 to (GImixed.Count-1)
    if TypeName(GImixed(x) ) = "Light" Then
      Set aLight(CountMe) = GImixed(x)
      TestString = TestString & "assigned " & GImixed(x).Name & " to aLight(" & CountMe & ")" & vbnewline 'debug
      CountMe = CountMe+1
      redim Preserve aLight(CountMe)
    end if
  Next
  CountMe = 0
  for x = 0 to (GImixed.Count-1)  '(note: this sub assumes there ARE flashers in the collection!)
    if TypeName(GImixed(x) ) = "Flasher" Then
      Set aFlasher(CountMe) = GImixed(x)
      TestString = TestString & "assigned " & GImixed(x).Name & " to aFlasher(" & CountMe & ")" & vbnewline 'debug
      CountMe = CountMe+1
      redim Preserve aFlasher(CountMe)
    end if
  Next
  redim Preserve aLight(uBound(aLight)-1) 'final trim of the arrays
  redim Preserve aFlasher(uBound(aFlasher)-1)
  'TestSTR(0) = TestSTR(0) & "ubound aLight: " & uBound(aLight) & " uBound aFlashers:" & uBound(aFlasher) 'debug
  'Debug.Print TestString
End Sub

'These arrays contain the following info of all non-GI lights (collected from GetElements via SortLamps sub)
Redim LightsA(999)' Object references
Redim LightsB(999)' Opacity / Intensity
Redim LightsC(999)' Fade Up (Light objects)
Redim LightsD(999)' Fade Down(Light Objects)

SortLamps GILighting, GIOffFlasherCorrection
Sub SortLamps(ByVal GI, aExclude) 'Sorts remaining light and flashers objects (EXCLUDES those in the GI collection)
  dim Counter,x,xx,skipme : skipme = False:Counter = 0 : TestStringAll = "Test String 2"
  for each x in GetElements 'now we're cooking
    'if TypeName(x) = "IDecal" then Continue For 'Decals don't have names. Evil imo D:
    if TypeName(x) = "Light" or TypeName(x) = "Flasher" Then
      SkipMe = False
      for each xx in GI 'Find duplicates and Skip them
        if x.Name = xx.Name then
          TestStringAll = TestStringAll & x.Name & "found in GI collection, Disregarding & Continuing..." & vbnewline 'debug
          SkipMe = True'Continue For
        End If
      next
      for each xx in aExclude 'Exclude collection
        if x.Name = xx.Name then
          TestStringAll = TestStringAll & x.Name & "found in exclude collection, Disregarding & Continuing..." & vbnewline 'debug
          SkipMe = True'Continue For
        End If
      next
      If Not SkipMe Then
        On Error Resume Next
        'LightsA(Counter) = x.name  'name
        Set LightsA(Counter) = x  'ref
        LightsB(Counter) = x.Opacity
        LightsB(Counter) = x.Intensity
        LightsC(Counter) = x.FadeSpeedUp
        LightsD(Counter) = x.FadeSpeedDown
        On Error Goto 0
        Counter = Counter + 1
        redim Preserve LightsA(Counter)
        redim Preserve LightsB(Counter)
        redim Preserve LightsC(Counter)
        redim Preserve LightsD(Counter)
      End If
    End If
  next
  redim Preserve LightsA(uBound(LightsA)-1) 'final trim of the arrays
  redim Preserve LightsB(uBound(LightsB)-1)
  redim Preserve LightsC(uBound(LightsC)-1)
  redim Preserve LightsD(uBound(LightsD)-1)

  TestStringAll = TestStringAll & "Ubound LightsA = " & UBound(LightsA) 'Debug
  'debug.print TestTwo
End Sub

' Lamp & Flasher Updates
' LampFader object (Lampz) updates on two timers: Logic on 1, Game updates on -1
function FlashLevelToIndex(Input, MaxSize)
  'FlashLevelToIndex = cInt(Input * (MaxSize-1)+.5)+1
     FlashLevelToIndex = cInt(MaxSize * Input)
end function

'Lamp Filter
Function LampFilter(aLvl)
  LampFilter = aLvl^1.6 'exponential curve?
End Function

'Collections to arrays
Function ColtoArray(aDict)  'converts a collection to an indexed array. Indexes will come out random probably.
  redim a(999)
  dim count : count = 0
  dim x  : for each x in aDict : set a(Count) = x : count = count + 1 : Next
  redim preserve a(count-1) : ColtoArray = a
End Function

'Setlamp, etc
  'Solenoid pipeline looks like this:
  'Pinmame Controller -> UseSolenoids -> Solcallback -> intermediate subs (here) -> Lampz fading object -> object updates / more callbacks

  'Lamps, for reference:
  'Pinmame Controller -> UpdateLamps sub -> Lampz Fading Object -> Object Updates / callbacks

Sub SetLamp(aNr, aOn)
  Lampz.state(aNr) = abs(aOn)
End Sub

Dim NullFader : set NullFader = new NullFadingObject
Dim Lampz : Set Lampz = New LampFader
InitLamps
LampTimer.Interval = 16 '1
LampTimer.Enabled = 1

' Lamp & Flasher Updates
' LampFader object (Lampz) updates on two timers: Logic on 1, Game updates on -1
Sub LampTimer_Timer()
  dim x, chglamp
  chglamp = Controller.ChangedLamps
  If Not IsEmpty(chglamp) Then
    For x = 0 To UBound(chglamp)      'nmbr = chglamp(x, 0), state = chglamp(x, 1)
      Lampz.state(chglamp(x, 0)) = chglamp(x, 1)
    next
  End If
  'Lampz.Update1  'update (fading logic only)
  Lampz.Update2 'update (Pinmame and Fading (for -1, lower latency)
End Sub

Sub InitLamps()
  'Filtering (comment out to disable)
  Lampz.Filter = "LampFilter" 'Puts all lamp intensity scale output (no callbacks) through this function before updating
  'Adjust fading speeds (1 / full MS fading time)
  dim x : for x = 0 to 140 : Lampz.FadeSpeedUp(x) = 1/80 : Lampz.FadeSpeedDown(x) = 1/100 : next
  Lampz.FadeSpeedUp(110) = 1/64 'GI

'*********************************************************************************************************************************************************
'End nfozzy lamp handling
'*********************************************************************************************************************************************************
  'Lamp Assignments
  Lampz.MassAssign(1) = f1
  Lampz.MassAssign(2) = f2
  Lampz.MassAssign(3) = f3
  Lampz.MassAssign(4) = f4
  Lampz.MassAssign(5) = f5
  Lampz.Callback(6) = "  DisableLighting p6, 15,"
  Lampz.MassAssign(6) = l6
  Lampz.Callback(7) = "  DisableLighting p7, 15,"
  Lampz.MassAssign(7) = l7
  Lampz.Callback(8) = "  DisableLighting p8, 15,"
  Lampz.MassAssign(8) = l8
  Lampz.MassAssign(9) =  f9
  Lampz.MassAssign(10) = f10
  Lampz.MassAssign(11) = f11
  Lampz.MassAssign(12) = f12
  Lampz.Callback(13) = " PlayMusic1"
  Lampz.Callback(13) = " DisableLighting p13, 10,"
  Lampz.Callback(14) = " DisableLighting p14, 20,"
  Lampz.MassAssign(15) = l15
  Lampz.Callback(15) = " DisableLighting p15, 15,"
  Lampz.MassAssign(16) = l16
  Lampz.Callback(16) = " DisableLighting p16, 15,"

' Lampz.MassAssign(17) = l17Wall
    Lampz.MassAssign(17) = pBulb17a2
  Lampz.Callback(17) = " DisableLighting pBulb17a1, 1,"


  Lampz.Callback(18) = " DisableLighting pBulb18a1, 1,"
    Lampz.MassAssign(18) = pBulb18a2



' Lampz.MassAssign(19) = l19Wall
  Lampz.Callback(19) = " DisableLighting pBulb19a1, 1,"
    Lampz.MassAssign(19) = pBulb19a2


  Lampz.Callback(20) = " DisableLighting pBulb20a1, 1,"
    Lampz.MassAssign(20) = pBulb20a2


' Lampz.MassAssign(21) = l21Wall
  Lampz.Callback(21) = " DisableLighting pBulb21a1, 1,"
    Lampz.MassAssign(21) = pBulb21a2
  Lampz.MassAssign(21) = Raylight21w

  Lampz.Callback(22) = " DisableLighting p22, 2,"
  Lampz.Callback(23) = " DisableLighting p23, 2,"
  Lampz.Callback(24) = " DisableLighting p24, 2,"
  Lampz.Callback(25) = " DisableLighting p25, 15,"
  Lampz.MassAssign(25) = l25
  Lampz.Callback(26) = " DisableLighting p26, 15,"
  Lampz.MassAssign(26) = l26
  Lampz.Callback(27) = " DisableLighting p27, 15,"
  Lampz.MassAssign(27) = l27
  Lampz.Callback(28) = " DisableLighting p28, 15,"
  Lampz.MassAssign(28) = l28
  Lampz.Callback(29) = " DisableLighting p29, 15,"
  Lampz.MassAssign(29) = l29
  Lampz.Callback(30) = " DisableLighting p30, 15,"
  Lampz.MassAssign(30) = l30
  Lampz.MassAssign(31) = l31
  Lampz.Callback(31) = " DisableLighting p31, 15,"
  Lampz.Callback(32) = " DisableLighting p32, 30,"
  Lampz.Callback(33) = " ImageSwap p33, DomeImageRed, .15,"
  Lampz.MassAssign(33) = l133Flare
  Lampz.MassAssign(33) = l133Wall
  Lampz.Callback(34) = " ImageSwap p34, DomeImageGreen, 1,"
  Lampz.MassAssign(34) = l134Flare
  Lampz.Callback(35) = " ImageSwap p35, DomeImageYellow, .5,"
  Lampz.MassAssign(35) = l135Flare
  Lampz.MassAssign(35) = l135Wall
  Lampz.Callback(36) = " DisableLighting pBulb36, 1, "
  Lampz.Callback(36) = " DisableLighting pFil36, 200, "
  Lampz.MassAssign(36) = L36Wall
  Lampz.Callback(37) = " DisableLighting pBulb37, 1, "
  Lampz.Callback(37) = " DisableLighting pFil37, 100, "


  Lampz.Callback(38) = " DisableLighting pBulb38, 1, "
  Lampz.Callback(38) = " DisableLighting pFil38, 100,"


  Lampz.Callback(39) = " DisableLighting p39, 30,"
  Lampz.MassAssign(40) = l40a
  Lampz.MassAssign(40) = l40b
  Lampz.Callback(40) = " DisableLighting p40a, 15,"
  Lampz.Callback(40) = " DisableLighting p40b, 15,"
  Lampz.Callback(41) = " DisableLighting p41, 15,"
  Lampz.MassAssign(41) = l41
  Lampz.MassAssign(42) = l42
  Lampz.Callback(42) = " DisableLighting p42, 15,"
  Lampz.MassAssign(43) = l43
  Lampz.Callback(43) = " DisableLighting p43, 15,"
  Lampz.MassAssign(44) = l44
  Lampz.Callback(44) = " DisableLighting p44, 15,"
  Lampz.Callback(45) = " DisableLighting p45, 2,"
  Lampz.MassAssign(46) = l46
  Lampz.Callback(46) = " DisableLighting p46, 15, "
  Lampz.MassAssign(47) = l47
  Lampz.Callback(47) = " DisableLighting p47, 15, "
  Lampz.MassAssign(48) = l48
  Lampz.Callback(48) = " DisableLighting p48, 15, "
  Lampz.MassAssign(49) = l49
  Lampz.Callback(49) = " DisableLighting p49, 15,"
  Lampz.MassAssign(50) = l50
  Lampz.Callback(50) = " DisableLighting p50, 15,"
  Lampz.MassAssign(51) = l51
  Lampz.Callback(51) = " DisableLighting p51, 15,"
  Lampz.MassAssign(52) = l52
  Lampz.Callback(52) = " DisableLighting p52, 15,"
  Lampz.MassAssign(53) = l53
  Lampz.Callback(53) = " DisableLighting p53, 15,"
  Lampz.MassAssign(54) = l54
  Lampz.Callback(54) = " DisableLighting p54, 15, "
  Lampz.MassAssign(55) = l55
  Lampz.Callback(55) = " DisableLighting p55, 15, "
  Lampz.MassAssign(56) = l56
  Lampz.Callback(56) = " DisableLighting p56, 15, "

  Lampz.Callback(57) = " DisableLighting pBulb57, 1, "
  Lampz.Callback(57) = " DisableLighting pFil57, 100, "

  Lampz.MassAssign(58) = l58
  Lampz.Callback(58) = " DisableLighting p58, 15,"
  Lampz.MassAssign(59) = l59
  Lampz.Callback(59) = " DisableLighting p59, 2,"
  Lampz.Callback(60) = " DisableLighting p60, 1.25,"
  Lampz.MassAssign(60) = l60
  Lampz.Callback(61) = " DisableLighting p61, 1.25,"
  Lampz.MassAssign(61) = l61
  Lampz.Callback(62) = " DisableLighting p62, 1.25,"
  Lampz.MassAssign(62) = l62
  Lampz.Callback(63) = " DisableLighting p63, 1.25,"
  Lampz.MassAssign(63) = l63
  Lampz.Callback(64) = " DisableLighting p64, 1.25,"
  Lampz.MassAssign(64) = l64
  Lampz.Callback(110) = "ImageSwap Bumper1Cap, DomeImageCap, .2,"
  Lampz.Callback(110) = "ImageSwap Bumper2Cap, DomeImageCap, .2,"
  Lampz.Callback(110) = "ImageSwap Bumper3Cap, DomeImageCap, .2,"
  Lampz.Callback(110) = "DisableLighting Prim_MixMasterStickers_ON, 6,"
  Lampz.Callback(110) = "DisableLighting pRamp, .5,"
  Lampz.Callback(110) = "DisableLighting pRamp_DT, .5,"


'*****************
'Flasher Assignments
'*****************
  Lampz.MassAssign(115) = array (f115, f115a) 'Big Guy Flasher
  Lampz.obj(125) = Array(F33,MMBloom, MMBloom1, flasher2, Flasher4) 'Mixer Heart Flasher

  Lampz.obj(126) = Array(F35, flasher3, Flasher5)   'Mixer Gab Flasher
  Lampz.obj(127) = Array(F34, flasher6, Flasher7)   'Mixer Magnet Flasher
  Lampz.MassAssign(128) = Array (f128, f128a) 'Magnet Flasher
  Lampz.MassAssign(129) = array (f129, f129a)
    Lampz.Callback(129) = "DisableLighting pGOG, 4," 'GOG Flasher
  Lampz.MassAssign(129) = F129
  Lampz.MassAssign(129) = F129b
  Lampz.MassAssign(129) = GOGWall
  Lampz.MassAssign(130) = array (f130, f130a) 'Heart Flasher
  Lampz.MassAssign(131) = array (f131, f131a) 'DT flasher
  Lampz.Callback(132) = "DisableLighting pBulb132b, 1,"
  'Lampz.Callback(132) = "DisableLighting pFilament132b, 500,"
  Lampz.Callback(132) = "DisableLighting pRaygun, 1,"
  Lampz.MassAssign(132)= F200
  Lampz.MassAssign(132)= F200a
  Lampz.MassAssign(132)= F200Wall
  Lampz.MassAssign(132)= F200c
  Lampz.MassAssign(132)= f132Bloom
    'Lampz.callback(132) = "RaygunOn"


'*****************
'GI Assignments
'*****************
  Lampz.Callback(110) = "GIUpdates"
  Lampz.obj(110) = ColtoArray(GILighting)
  Lampz.state(110) = 1    'Turn on GI to Start
  'Turn off all lamps on startup
  lampz.TurnOnStates  'Set any lamps state to 1. (Object handles fading!)
  lampz.update

End Sub


'*********************************************************************************************************************************************************
'Begin lamp helper functions
'*********************************************************************************************************************************************************

'***************************************
'System 11 GI On/Off
'***************************************
Sub GIOn  : SetGI False: End Sub 'These are just debug commands now
Sub GIOff : SetGI True : End Sub

'***************************************
'GI off insert lamps intensity boost
'***************************************
Dim GIoffMult : GIoffMult = 3 'Multiplies all non-GI inserts lights opacities when the GI is off
Sub GIupdates(ByVal aLvl) 'GI update odds and ends go here
  if Lampz.UseFunction then aLvl = LampFilter(aLvl) 'Callbacks don't get this filter automatically

  'Fade lamps up when GI is off
  dim GIscale
  GiScale = (GIoffMult-1) * (ABS(aLvl-1 )  ) + 1  'invert
  dim x : for x = 0 to uBound(LightsA)
    On Error Resume Next
    LightsA(x).Opacity = LightsB(x) * GIscale
    LightsA(x).Intensity = LightsB(x) * GIscale
    'LightsA(x).FadeSpeedUp = LightsC(x) * GIscale
    'LightsA(x).FadeSpeedDown = LightsD(x) * GIscale
    On Error Goto 0
  Next
End Sub

'***************************************
'GI off finsert light falloff-power boost
'***************************************
Dim GiOffFOP
Sub SetGI(aOn)
        Select Case aOn
    Case True  'GI off
      Sound_GI_Relay 0,sw31
      SetLamp 110, 0  'Inverted, Solenoid cuts GI circuit on this era of game
            For each GiOffFOP in LampsInserts 'increases falloff power for lamps in "lampinserts" collection, when GI is turned off
        GIOffFOP.falloffpower = 1.25 'Sets falloff power to 1.25 when GI is ON
            next

    Case False
      Sound_GI_Relay 1,sw31
      SetLamp 110, 5
            For each GiOffFOP in LampsInserts  'reduces falloff power for lamps in "lampinserts" collection, when GI is turned off
        GiOffFOP.falloffpower = 4 'Sets falloff power to 4 when GI is OFF
            next
  End Select
End Sub

'*********************************************************************************************************************************************************
'End lamp helper functions
'*********************************************************************************************************************************************************

'*********************************************************************************************************************************************************
'End lamp helper functions
'*********************************************************************************************************************************************************


Sub PlayMusic1(aLvl)
    'Intro Music
  If aLvl > 0 and PrevGameOver = 0 Then
    If IntroMusic = 0 Then
      PlaySound "zIntro"
      PrevGameOver = 1
    End If
  else
'   PrevGameOver = 0
  End If
End Sub






'*** Determines if a Points (px,py) is inside a 4 point polygon A-D in Clockwise/CCW order

Function InRect(px,py,ax,ay,bx,by,cx,cy,dx,dy)
  Dim AB, BC, CD, DA
  AB = (bx*py) - (by*px) - (ax*py) + (ay*px) + (ax*by) - (ay*bx)
  BC = (cx*py) - (cy*px) - (bx*py) + (by*px) + (bx*cy) - (by*cx)
  CD = (dx*py) - (dy*px) - (cx*py) + (cy*px) + (cx*dy) - (cy*dx)
  DA = (ax*py) - (ay*px) - (dx*py) + (dy*px) + (dx*ay) - (dy*ax)

  If (AB <= 0 AND BC <=0 AND CD <= 0 AND DA <= 0) Or (AB >= 0 AND BC >=0 AND CD >= 0 AND DA >= 0) Then
    InRect = True
  Else
    InRect = False
  End If
End Function





'******************************************************
'****  BALL ROLLING AND DROP SOUNDS
'******************************************************
Dim BallInKicker1, BallSaucer1
Dim BallInKicker2, BallSaucer2


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

Sub RollingSoundUpdate()
  Dim BOT, b
  BOT = GetBalls

  ' stop the sound of deleted balls
  For b = UBound(BOT) + 1 to tnob
    ' Comment the next line if you are not implementing Dyanmic Ball Shadows
    If AmbientBallShadowOn = 0 Then BallShadowA(b).visible = 0
    rolling(b) = False
    StopSound("BallRoll_" & b)
  Next

  ' exit the sub if no balls on the table
  If UBound(BOT) = -1 Then Exit Sub

  ' play the rolling sound for each ball

  For b = 0 to UBound(BOT)
    If BallVel(BOT(b)) > 1 AND BOT(b).z < 30 Then
      rolling(b) = True
      PlaySound ("BallRoll_" & b), -1, VolPlayfieldRoll(BOT(b)) * BallRollVolume * VolumeDial, AudioPan(BOT(b)), 0, PitchPlayfieldRoll(BOT(b)), 1, 0, AudioFade(BOT(b))

    Else
      If rolling(b) = True Then
        StopSound("BallRoll_" & b)
        rolling(b) = False
      End If
    End If

    ' Ball Drop Sounds
    If BOT(b).VelZ < -1 and BOT(b).z < 55 and BOT(b).z > 27 Then 'height adjust for ball drop sounds
      If DropCount(b) >= 5 Then
        DropCount(b) = 0
        If BOT(b).velz > -7 Then
          RandomSoundBallBouncePlayfieldSoft BOT(b)
        Else
          RandomSoundBallBouncePlayfieldHard BOT(b)
        End If
      End If
    End If
    If DropCount(b) < 5 Then
      DropCount(b) = DropCount(b) + 1
    End If

    ' "Static" Ball Shadows
    ' Comment the next If block, if you are not implementing the Dyanmic Ball Shadows
    If AmbientBallShadowOn = 0 Then
      If BOT(b).Z > 30 Then
        BallShadowA(b).height=BOT(b).z - BallSize/4   'This is technically 1/4 of the ball "above" the ramp, but it keeps it from clipping
      Else
        BallShadowA(b).height=BOT(b).z - BallSize/2 + 5
      End If
      BallShadowA(b).Y = BOT(b).Y + Ballsize/5 + fovY
      BallShadowA(b).X = BOT(b).X
      BallShadowA(b).visible = 1
    End If


    ' ********* VUK Code sw32
    If InRect(BOT(b).x,BOT(b).y,736,633,788,633,788,680,736,680) Then
      If ABS(BOT(b).velx) < 1 and ABS(BOT(b).vely) < 1 Then
        Controller.Switch(32) = 1
        BallInKicker1 = 1
        Set BallSaucer1 = BOT(b)
      End If
    End If
    ' ********* VUK Saucer Code

    ' ********* VUK Code sw51
    If InRect(BOT(b).x,BOT(b).y,75,360,125,360,125,410,75,410) Then
      If ABS(BOT(b).velx) < 1 and ABS(BOT(b).vely) < 1 Then
        Controller.Switch(51) = 1
        BallInKicker2 = 1
        Set BallSaucer2 = BOT(b)
      End If
    End If
    ' ********* VUK Saucer Code

  Next
End Sub


'******************************************************
'****  END BALL ROLLING AND DROP SOUNDS
'******************************************************




'******************************************************
'**** RAMP ROLLING SFX
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
dim RampBalls(6,2)
'x,0 = ball x,1 = ID, 2 = Protection against ending early (minimum amount of updates)
'0,0 is boolean on/off, 0,1 unused for now
RampBalls(0,0) = False

' RampType
'     Setup: Set this array to the number Total number of balls that can be tracked at one time + 1.  5 ball multiball then set value to 6
'     Description: Array type indexed on BallId and a values used to deterimine what type of ramp the ball is on: False = Wire Ramp, True = Plastic Ramp
dim RampType(6)

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



'***Ramp Rolling Sound Triggers***


Sub WireRampStart1_Hit()
  WireRampOn False
End Sub

Sub WireRampStart2_Hit()
  WireRampOn False
End Sub

Sub WireRampStart3_Hit()
  WireRampOn False
End Sub

Sub WireRampEnd1_Hit()
  WireRampOff
End Sub

Sub WireRampEnd2_Hit()
  WireRampOff
End Sub

Sub WireRampEnd3_Hit()
  WireRampOff
End Sub

Sub BallDrop3_hit
  WireRampOff
End Sub

Sub sw15_UnHit
  If Activeball.vely > 0 Then WireRampOff
End Sub



'******************************************************
'**** END RAMP ROLLING SFX
'******************************************************






'***************************************************************
'****  VPW DYNAMIC BALL SHADOWS by Iakki, Apophis, and Wylte
'***************************************************************

Const fovY          = 0   'Offset y position under ball to account for layback or inclination (more pronounced need further back)
Const DynamicBSFactor     = 0.95  '0 to 1, higher is darker
Const AmbientBSFactor     = 0.7 '0 to 1, higher is darker
Const AmbientMovement   = 2   '1 to 4, higher means more movement as the ball moves left and right
Const Wideness        = 20  'Sets how wide the dynamic ball shadows can get (20 +5 thinness should be most realistic for a 50 unit ball)
Const Thinness        = 5   'Sets minimum as ball moves away from source

Dim sourcenames, currentShadowCount, DSSources(30), numberofsources, numberofsources_hold
sourcenames = Array ("","","","","","","","","","","","")
currentShadowCount = Array (0,0,0,0,0,0,0,0,0,0,0,0)

' *** Trim or extend these to match the number of balls/primitives/flashers on the table!
dim objrtx1(12), objrtx2(12)
dim objBallShadow(12)
Dim BallShadowA
BallShadowA = Array (BallShadowA0,BallShadowA1,BallShadowA2,BallShadowA3,BallShadowA4,BallShadowA5,BallShadowA6,BallShadowA7,BallShadowA8,BallShadowA9,BallShadowA10,BallShadowA11)

DynamicBSInit

sub DynamicBSInit()
  Dim iii, source

  for iii = 0 to tnob                 'Prepares the shadow objects before play begins
    Set objrtx1(iii) = Eval("RtxBallShadow" & iii)
    objrtx1(iii).material = "RtxBallShadow" & iii
    objrtx1(iii).z = iii/1000 + 0.01
    objrtx1(iii).visible = 0

    Set objrtx2(iii) = Eval("RtxBall2Shadow" & iii)
    objrtx2(iii).material = "RtxBallShadow2_" & iii
    objrtx2(iii).z = (iii)/1000 + 0.02
    objrtx2(iii).visible = 0

    currentShadowCount(iii) = 0

    Set objBallShadow(iii) = Eval("BallShadow" & iii)
    objBallShadow(iii).material = "BallShadow" & iii
    UpdateMaterial objBallShadow(iii).material,1,0,0,0,0,0,AmbientBSFactor,RGB(0,0,0),0,0,False,True,0,0,0,0
    objBallShadow(iii).Z = iii/1000 + 0.04
    objBallShadow(iii).visible = 0

    BallShadowA(iii).Opacity = 100*AmbientBSFactor
    BallShadowA(iii).visible = 0
  Next

  iii = 0

  For Each Source in DynamicSources
    DSSources(iii) = Array(Source.x, Source.y)
    iii = iii + 1
  Next
  numberofsources = iii
  numberofsources_hold = iii
end sub


Sub DynamicBSUpdate
  Dim falloff:  falloff = 150     'Max distance to light sources, can be changed if you have a reason
  Dim ShadowOpacity, ShadowOpacity2
  Dim s, Source, LSd, currentMat, AnotherSource, BOT, iii
  BOT = GetBalls

  'Hide shadow of deleted balls
  For s = UBound(BOT) + 1 to tnob
    objrtx1(s).visible = 0
    objrtx2(s).visible = 0
    objBallShadow(s).visible = 0
    BallShadowA(s).visible = 0
  Next

  If UBound(BOT) < lob Then Exit Sub    'No balls in play, exit

'The Magic happens now
  For s = lob to UBound(BOT)

' *** Normal "ambient light" ball shadow
  'Layered from top to bottom. If you had an upper pf at for example 80 and ramps even above that, your segments would be z>110; z<=110 And z>100; z<=100 And z>30; z<=30 And z>20; Else invisible

    If AmbientBallShadowOn = 1 Then     'Primitive shadow on playfield, flasher shadow in ramps
      If BOT(s).Z > 30 Then             'The flasher follows the ball up ramps while the primitive is on the pf
        If BOT(s).X < tablewidth/2 Then
          objBallShadow(s).X = ((BOT(s).X) - (Ballsize/10) + ((BOT(s).X - (tablewidth/2))/(Ballsize/AmbientMovement))) + 5
        Else
          objBallShadow(s).X = ((BOT(s).X) + (Ballsize/10) + ((BOT(s).X - (tablewidth/2))/(Ballsize/AmbientMovement))) - 5
        End If
        objBallShadow(s).Y = BOT(s).Y + BallSize/10 + fovY
        objBallShadow(s).visible = 1

        BallShadowA(s).X = BOT(s).X
        BallShadowA(s).Y = BOT(s).Y + BallSize/5 + fovY
        BallShadowA(s).height=BOT(s).z - BallSize/4   'This is technically 1/4 of the ball "above" the ramp, but it keeps it from clipping
        BallShadowA(s).visible = 1
      Elseif BOT(s).Z <= 30 And BOT(s).Z > 20 Then  'On pf, primitive only
        objBallShadow(s).visible = 1
        If BOT(s).X < tablewidth/2 Then
          objBallShadow(s).X = ((BOT(s).X) - (Ballsize/10) + ((BOT(s).X - (tablewidth/2))/(Ballsize/AmbientMovement))) + 5
        Else
          objBallShadow(s).X = ((BOT(s).X) + (Ballsize/10) + ((BOT(s).X - (tablewidth/2))/(Ballsize/AmbientMovement))) - 5
        End If
        objBallShadow(s).Y = BOT(s).Y + fovY
        BallShadowA(s).visible = 0
      Else                      'Under pf, no shadows
        objBallShadow(s).visible = 0
        BallShadowA(s).visible = 0
      end if

    Elseif AmbientBallShadowOn = 2 Then   'Flasher shadow everywhere
      If BOT(s).Z > 30 Then             'In a ramp
        BallShadowA(s).X = BOT(s).X
        BallShadowA(s).Y = BOT(s).Y + BallSize/5 + fovY
        BallShadowA(s).height=BOT(s).z - BallSize/4   'This is technically 1/4 of the ball "above" the ramp, but it keeps it from clipping
        BallShadowA(s).visible = 1
      Elseif BOT(s).Z <= 30 And BOT(s).Z > 20 Then  'On pf
        BallShadowA(s).visible = 1
        If BOT(s).X < tablewidth/2 Then
          BallShadowA(s).X = ((BOT(s).X) - (Ballsize/10) + ((BOT(s).X - (tablewidth/2))/(Ballsize/AmbientMovement))) + 5
        Else
          BallShadowA(s).X = ((BOT(s).X) + (Ballsize/10) + ((BOT(s).X - (tablewidth/2))/(Ballsize/AmbientMovement))) - 5
        End If
        BallShadowA(s).Y = BOT(s).Y + Ballsize/10 + fovY
        BallShadowA(s).height=BOT(s).z - BallSize/2 + 5
      Else                      'Under pf
        BallShadowA(s).visible = 0
      End If
    End If

' *** Dynamic shadows
    If DynamicBallShadowsOn Then
      If BOT(s).Z < 30 Then 'And BOT(s).Y < (TableHeight - 200) Then 'Or BOT(s).Z > 105 Then    'Defining when and where (on the table) you can have dynamic shadows
        For iii = 0 to numberofsources - 1
          LSd=DistanceFast((BOT(s).x-DSSources(iii)(0)),(BOT(s).y-DSSources(iii)(1))) 'Calculating the Linear distance to the Source
          If LSd < falloff Then               'If the ball is within the falloff range of a light and light is on (we will set numberofsources to 0 when GI is off)
            currentShadowCount(s) = currentShadowCount(s) + 1   'Within range of 1 or 2
            if currentShadowCount(s) = 1 Then           '1 dynamic shadow source
              sourcenames(s) = iii
              currentMat = objrtx1(s).material
              objrtx2(s).visible = 0 : objrtx1(s).visible = 1 : objrtx1(s).X = BOT(s).X : objrtx1(s).Y = BOT(s).Y + fovY
  '           objrtx1(s).Z = BOT(s).Z - 25 + s/1000 + 0.01            'Uncomment if you want to add shadows to an upper/lower pf
              objrtx1(s).rotz = AnglePP(DSSources(iii)(0), DSSources(iii)(1), BOT(s).X, BOT(s).Y) + 90
              ShadowOpacity = (falloff-LSd)/falloff                 'Sets opacity/darkness of shadow by distance to light
              objrtx1(s).size_y = Wideness*ShadowOpacity+Thinness           'Scales shape of shadow with distance/opacity
              UpdateMaterial currentMat,1,0,0,0,0,0,ShadowOpacity*DynamicBSFactor^2,RGB(0,0,0),0,0,False,True,0,0,0,0
              If AmbientBallShadowOn = 1 Then
                currentMat = objBallShadow(s).material                  'Brightens the ambient primitive when it's close to a light
                UpdateMaterial currentMat,1,0,0,0,0,0,AmbientBSFactor*(1-ShadowOpacity),RGB(0,0,0),0,0,False,True,0,0,0,0
              Else
                BallShadowA(s).Opacity = 100*AmbientBSFactor*(1-ShadowOpacity)
              End If

            Elseif currentShadowCount(s) = 2 Then
                                  'Same logic as 1 shadow, but twice
              currentMat = objrtx1(s).material
              AnotherSource = sourcenames(s)
              objrtx1(s).visible = 1 : objrtx1(s).X = BOT(s).X : objrtx1(s).Y = BOT(s).Y + fovY
  '           objrtx1(s).Z = BOT(s).Z - 25 + s/1000 + 0.01              'Uncomment if you want to add shadows to an upper/lower pf
              objrtx1(s).rotz = AnglePP(DSSources(AnotherSource)(0),DSSources(AnotherSource)(1), BOT(s).X, BOT(s).Y) + 90
              ShadowOpacity = (falloff-DistanceFast((BOT(s).x-DSSources(AnotherSource)(0)),(BOT(s).y-DSSources(AnotherSource)(1))))/falloff
              objrtx1(s).size_y = Wideness*ShadowOpacity+Thinness
              UpdateMaterial currentMat,1,0,0,0,0,0,ShadowOpacity*DynamicBSFactor^3,RGB(0,0,0),0,0,False,True,0,0,0,0

              currentMat = objrtx2(s).material
              objrtx2(s).visible = 1 : objrtx2(s).X = BOT(s).X : objrtx2(s).Y = BOT(s).Y + fovY
  '           objrtx2(s).Z = BOT(s).Z - 25 + s/1000 + 0.02              'Uncomment if you want to add shadows to an upper/lower pf
              objrtx2(s).rotz = AnglePP(DSSources(iii)(0), DSSources(iii)(1), BOT(s).X, BOT(s).Y) + 90
              ShadowOpacity2 = (falloff-LSd)/falloff
              objrtx2(s).size_y = Wideness*ShadowOpacity2+Thinness
              UpdateMaterial currentMat,1,0,0,0,0,0,ShadowOpacity2*DynamicBSFactor^3,RGB(0,0,0),0,0,False,True,0,0,0,0
              If AmbientBallShadowOn = 1 Then
                currentMat = objBallShadow(s).material                  'Brightens the ambient primitive when it's close to a light
                UpdateMaterial currentMat,1,0,0,0,0,0,AmbientBSFactor*(1-max(ShadowOpacity,ShadowOpacity2)),RGB(0,0,0),0,0,False,True,0,0,0,0
              Else
                BallShadowA(s).Opacity = 100*AmbientBSFactor*(1-max(ShadowOpacity,ShadowOpacity2))
              End If
            end if
          Else
            currentShadowCount(s) = 0
            BallShadowA(s).Opacity = 100*AmbientBSFactor
          End If
        Next
      Else                  'Hide dynamic shadows everywhere else
        objrtx2(s).visible = 0 : objrtx1(s).visible = 0
      End If
    End If
  Next
End Sub
'****************************************************************
'****  END VPW DYNAMIC BALL SHADOWS by Iakki, Apophis, and Wylte
'****************************************************************








'**********************************************************************************************************
'**********************************************************************************************************
'***Options***
'**********************************************************************************************************
'**********************************************************************************************************


Dim BumperColor, Prevgameover, IntroMusic, xxGIColor, aproncolor, flipperstyle
Dim cheaterpost, raybeam, MMflashers, GIColorModType, GIColorMod, BumperColorType, Rails

Sub SetOptions()

If GIColorMod = 4 Then
  GIColorModType = Int(Rnd*4)
Else
  GIColorModType = GIColorMod
End If

If GIColorModType = 0 then
  for each xxGIColor in GIMain
    xxGIColor.Color=White
    xxGIColor.ColorFull=WhiteFull
    xxGIColor.Intensity = WhiteI
    next
  for each xxGIColor in GIPlastic
    xxGIColor.Color=WhitePlastic
    xxGIColor.ColorFull=WhitePlasticFull
    xxGIColor.Intensity = WhitePlasticI
    next
    for each xxGIColor in GIBulbs
    xxGIColor.Color=WhiteBulbs
    xxGIColor.ColorFull=WhiteBulbsFull
    xxGIColor.Intensity = WhiteBulbsI
  next
  for each xxGIColor in GIOverhead
      xxGIColor.Color=WhiteOverhead
      xxGIColor.ColorFull=WhiteOverheadFull
      xxGIColor.Intensity=WhiteOverheadI
  Next
End If

If GIColorModType = 1 then
  for each xxGIColor in GIMain
    xxGIColor.Color=Blue
    xxGIColor.ColorFull=BlueFull
    xxGIColor.Intensity = BlueI
    next
  for each xxGIColor in GIPlastic
    xxGIColor.Color=BluePlastic
    xxGIColor.ColorFull=BluePlasticFull
    xxGIColor.Intensity = BluePlasticI
    next
    for each xxGIColor in GIBulbs
    xxGIColor.Color=BlueBulbs
    xxGIColor.ColorFull=BlueBulbsFull
    xxGIColor.Intensity = BlueBulbsI
    next
      for each xxGIColor in GIOverhead
      xxGIColor.Color=BlueOverhead
      xxGIColor.ColorFull=BlueOverheadFull
      xxGIColor.Intensity=BlueOverheadI
        Next
End If

If GIColorModType = 2 then
  for each xxGIColor in GIMain
    xxGIColor.Color=Purple
    xxGIColor.ColorFull=PurpleFull
    xxGIColor.Intensity = PurpleI
    next

  for each xxGIColor in GIPlastic
    xxGIColor.Color=PurplePlastic
    xxGIColor.ColorFull=PurplePlasticFull
    xxGIColor.Intensity = PurplePlasticI
    next
    for each xxGIColor in GIBulbs
    xxGIColor.Color=PurpleBulbs
    xxGIColor.ColorFull=PurpleBulbsFull
    xxGIColor.Intensity = PurpleBulbsI
    next
      for each xxGIColor in GIOverhead
      xxGIColor.Color=PurpleOverhead
      xxGIColor.ColorFull=PurpleOverheadFull
      xxGIColor.Intensity=PurpleOverheadI
    Next
End If

If GIColorModType = 3 then
  for each xxGIColor in GIMain
    xxGIColor.Color=IceBlue
    xxGIColor.ColorFull=IceBlueFull
    xxGIColor.Intensity = IceBlueI
    next
  for each xxGIColor in GIPlastic
    xxGIColor.Color=IceBluePlastic
    xxGIColor.ColorFull=IceBluePlasticFull
    xxGIColor.Intensity = IceBluePlasticI
    next
    for each xxGIColor in GIBulbs
    xxGIColor.Color=IceBlueBulbs
    xxGIColor.ColorFull=IceBlueBulbsFull
    xxGIColor.Intensity = IceBlueBulbsI
    next
      for each xxGIColor in GIOverhead
      xxGIColor.Color=IceBlueOverhead
      xxGIColor.ColorFull=IceBlueOverheadFull
      xxGIColor.Intensity=IceBlueOverheadI
        Next
  End If

  If Aproncolor = 4 then Aproncolor = Int(Rnd*4) End If
  Select Case Aproncolor
    Case 0 :
      pApron.image = "apron_texture_White":
      pApronOverlay.Visible = False:
      pCustomWall.visible = false:
      'pSidewall_DT.visible = false
    Case 1 :
      pApron.image = "apron_texture_Blue":
      pApronOverlay.Visible = False:
      pCustomWall.visible = false:
      'pSidewall_DT.visible = false
    Case 2 :
      pApron.image = "apron_texture_Green":
      pApronOverlay.Visible = False:
      pCustomWall.visible = false:
      'pSidewall_DT.visible = false
      Case 3 :
      pApronOverlay.Visible = True:
      If DesktopMode = True Then
        'pSidewall_DT.visible = True
        pCustomWall.visible = true
      else
        'pSidewall_DT.visible = true
        pCustomWall.visible = true
      End If
  End Select

  if flipperstyle = 2 then flipperstyle = int(rnd*2) end if
  select case flipperstyle
    case 0: BatRightYellow.visible=True: BatLeftYellow.visible=True: BatRightWhite.visible=False: BatLeftWhite.visible=False
    case 1: BatRightYellow.visible=False: BatLeftYellow.visible=False: BatRightWhite.visible=True: BatLeftWhite.visible=True
  end select


    If cheaterpost = 1 then
    cpost.visible = 1
    cRubberRubber.collidable = 1
    cRubberRubber.visible = 1
  Else
    cpost.visible = 0
    cRubberRubber.collidable = 0
    cRubberRubber.visible = 0
  End If

  If Rails =1 Then
    Leftrail.visible = 0
    Rightrail.visible = 0
  Else
    If VRRoom = 0 Then
      Leftrail.visible = 1
      Rightrail.visible = 1
    End If
  End if

  If Raybeam = 1 then
       f200.visible = 1
       f200a.visible = 1
  Else
       f200.visible = 0
       f200a.visible = 0
  End If

If BumperColor = 4 Then
  BumperColorType = Int(Rnd*4)
Else
  BumperColorType = BumperColor
End If
End Sub



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
            BallShadow(b).X = ((BOT(b).X) - (Ballsize/6) + ((BOT(b).X - (Table1.Width/2))/21)) + 6
        Else
            BallShadow(b).X = ((BOT(b).X) + (Ballsize/6) + ((BOT(b).X - (Table1.Width/2))/21)) - 6
        End If
        ballShadow(b).Y = BOT(b).Y + 4
        If BOT(b).Z > 20 Then
            BallShadow(b).visible = 1
        Else
            BallShadow(b).visible = 0
        End If
    Next
End Sub


'**********************************************************************************************************
'Digital Display
'**********************************************************************************************************

Dim Digits(32)
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

 Sub DisplayTimer_Timer
    Dim ChgLED, ii, jj, num, chg, stat, obj, b, x
    ChgLED=Controller.ChangedLEDs(&Hffffffff, &Hffffffff)
    If Not IsEmpty(ChgLED)Then
    If DesktopMode = True Then
       For ii=0 To UBound(chgLED)
          num=chgLED(ii, 0) : chg=chgLED(ii, 1) : stat=chgLED(ii, 2)
      if (num < 32) then
              For Each obj In Digits(num)
                   If chg And 1 Then obj.State=stat And 1
                   chg=chg\2 : stat=stat\2
                  Next
      Else
             end if
        Next
     end if
    End If
 End Sub
'

'**********************************************************************************************************
'**********************************************************************************************************
'***RGB OPTIONS COLOR ADJUST***
'**********************************************************************************************************
'**********************************************************************************************************


Dim White, WhiteFull, WhiteI, WhitePlastic, WhitePlasticFull, WhitePlasticI, WhiteBumper, WhiteBumperFull, WhiteBumperI, WhiteBulbs, WhiteBulbsFull, WhiteBulbsI, WhiteOverheadFull, WhiteOverhead, WhiteOverheadI
WhiteFull = rgb(255,255,220)
White = rgb(255,255,180)
WhiteI = .25
WhitePlasticFull = rgb(255,255,220)
WhitePlastic = rgb(255,255,180)
WhitePlasticI = 4
WhiteBumperFull = rgb(255,255,220)
WhiteBumper = rgb(255,255,180)
WhiteBumperI = .1
WhiteBulbsFull = rgb(255,255,220)
WhiteBulbs = rgb(255,255,180)
WhiteBulbsI = 55
WhiteOverheadFull = rgb(255,255,220)
WhiteOverhead = rgb(255,255,180)
WhiteOverheadI = .1

Dim Blue, BlueFull, BlueI, BluePlastic, BluePlasticFull, BluePlasticI, BlueBumper, BlueBumperFull, BlueBumperI, BlueBulbs, BlueBulbsFull, BlueBulbsI,  BlueOverheadFull, BlueOverhead, BlueOverheadI
BlueFull = rgb(75,75,255)
Blue = rgb(75,75,255)
BlueI = 1
BluePlasticFull = rgb(75,75,255)
BluePlastic = rgb(75,75,255)
BluePlasticI = 3
BlueBumperFull = rgb(0,0,255)
BlueBumper = rgb(0,0,255)
BlueBumperI = 1
BlueBulbsFull = rgb(75,75,255)
BlueBulbs = rgb(75,75,255)
BlueBulbsI = 125
BlueOverheadFull = rgb(10,10,255)
BlueOverhead = rgb(10,10,255)
BlueOverheadI = .6

Dim Purple, PurpleFull, PurpleI, PurplePlastic, PurplePlasticFull, PurplePlasticI, PurpleBumper, PurpleBumperFull, PurpleBumperI, PurpleBulbs, PurpleBulbsFull, PurpleBulbsI, PurpleOverheadFull,PurpleOverhead,PurpleOverheadI
Purple = rgb(125,60,125)
PurpleI = 1
PurplePlasticFull = rgb(125,0,125)
PurplePlastic = rgb(125,0,125)
PurplePlasticI = 3
PurpleBumperFull = rgb(125,0,125)
PurpleBumper = rgb(125,0,125)
PurpleBumperI = .4
PurpleBulbsFull = rgb(125,50,125)
PurpleBulbs = rgb(125,0,125)
PurpleBulbsI = 250
PurpleOverheadFull = rgb(180,0,180)
PurpleOverhead = rgb(180,0,180)
PurpleOverheadI = .6

Dim IceBlue, IceBlueFull, IceBlueI, IceBluePlastic, IceBluePlasticFull, IceBluePlasticI, IceBlueBumper, IceBlueBumperFull, IceBlueBumperI, IceBlueBulbs, IceBlueBulbsFull, IceBlueBulbsI,IceBlueOverheadFull,IceBlueOverhead,IceBlueOverheadI
IceBlueFull = rgb(0,255,255)
IceBlue = rgb(0,255,255)
IceBlueI = 1
IceBluePlasticFull = rgb(0,255,255)
IceBluePlastic = rgb(0,255,255)
IceBluePlasticI = 3
IceBlueBumperFull = rgb(0,255,255)
IceBlueBumper = rgb(0,255,255)
IceBlueBumperI = .5
IceBlueBulbsFull = rgb(0,255,255)
IceBlueBulbs = rgb(0,255,255)
IceBlueBulbsI = 100
IceBlueOverheadFull = rgb(0,255,255)
IceBlueOverhead = rgb(0,255,255)
IceBlueOverheadI = .6


'====================
'Class jungle nf
'=============

'No-op object instead of adding more conditionals to the main loop
'It also prevents errors if empty lamp numbers are called, and it's only one object
'should be g2g?

Class NullFadingObject : Public Property Let IntensityScale(input) : : End Property : End Class

'version 0.11 - Mass Assign, Changed modulate style
'version 0.12 - Update2 (single -1 timer update) update method for core.vbs
'Version 0.12a - Filter can now be accessed via 'FilterOut'
'Version 0.12b - Changed MassAssign from a sub to an indexed property (new syntax: lampfader.MassAssign(15) = Light1 )
'Version 0.13 - No longer requires setlocale. Callback() can be assigned multiple times per index
' Note: if using multiple 'LampFader' objects, set the 'name' variable to avoid conflicts with callbacks

Class LampFader
  Public FadeSpeedDown(140), FadeSpeedUp(140)
  Private Lock(140), Loaded(140), OnOff(140)
  Public UseFunction
  Private cFilter
  Public UseCallback(140), cCallback(140)
  Public Lvl(140), Obj(140)
  Private Mult(140)
  Public FrameTime
  Private InitFrame
  Public Name

  Sub Class_Initialize()
    InitFrame = 0
    dim x : for x = 0 to uBound(OnOff)  'Set up fade speeds
      FadeSpeedDown(x) = 1/100  'fade speed down
      FadeSpeedUp(x) = 1/80   'Fade speed up
      UseFunction = False
      lvl(x) = 0
      OnOff(x) = False
      Lock(x) = True : Loaded(x) = False
      Mult(x) = 1
    Next
    Name = "LampFaderNF" 'NEEDS TO BE CHANGED IF THERE'S MULTIPLE OF THESE OBJECTS, OTHERWISE CALLBACKS WILL INTERFERE WITH EACH OTHER!!
    for x = 0 to uBound(OnOff)    'clear out empty obj
      if IsEmpty(obj(x) ) then Set Obj(x) = NullFader' : Loaded(x) = True
    Next
  End Sub

  Public Property Get Locked(idx) : Locked = Lock(idx) : End Property   'debug.print Lampz.Locked(100)  'debug
  Public Property Get state(idx) : state = OnOff(idx) : end Property
  Public Property Let Filter(String) : Set cFilter = GetRef(String) : UseFunction = True : End Property
  Public Function FilterOut(aInput) : if UseFunction Then FilterOut = cFilter(aInput) Else FilterOut = aInput End If : End Function
  'Public Property Let Callback(idx, String) : cCallback(idx) = String : UseCallBack(idx) = True : End Property
  Public Property Let Callback(idx, String)
    UseCallBack(idx) = True
    'cCallback(idx) = String 'old execute method
    'New method: build wrapper subs using ExecuteGlobal, then call them
    cCallback(idx) = cCallback(idx) & "___" & String  'multiple strings dilineated by 3x _

    dim tmp : tmp = Split(cCallback(idx), "___")

    dim str, x : for x = 0 to uBound(tmp) 'build proc contents
      'If Not tmp(x)="" then str = str & "  " & tmp(x) & " aLVL" & "  '" & x & vbnewline  'more verbose
      If Not tmp(x)="" then str = str & tmp(x) & " aLVL:"
    Next

    dim out : out = "Sub " & name & idx & "(aLvl):" & str & "End Sub"
    'if idx = 132 then msgbox out 'debug
    ExecuteGlobal Out

  End Property

  Public Property Let state(ByVal idx, input) 'Major update path
    if Input <> OnOff(idx) then  'discard redundant updates
      OnOff(idx) = input
      Lock(idx) = False
      Loaded(idx) = False
    End If
  End Property

  'Mass assign, Builds arrays where necessary
  'Sub MassAssign(aIdx, aInput)
  Public Property Let MassAssign(aIdx, aInput)
    If typename(obj(aIdx)) = "NullFadingObject" Then 'if empty, use Set
      if IsArray(aInput) then
        obj(aIdx) = aInput
      Else
        Set obj(aIdx) = aInput
      end if
    Else
      Obj(aIdx) = AppendArray(obj(aIdx), aInput)
    end if
  end Property

  Sub SetLamp(aIdx, aOn) : state(aIdx) = aOn : End Sub  'Solenoid Handler

  Public Sub TurnOnStates() 'If obj contains any light objects, set their states to 1 (Fading is our job!)
    dim debugstr
    dim idx : for idx = 0 to uBound(obj)
      if IsArray(obj(idx)) then
        'debugstr = debugstr & "array found at " & idx & "..."
        dim x, tmp : tmp = obj(idx) 'set tmp to array in order to access it
        for x = 0 to uBound(tmp)
          if typename(tmp(x)) = "Light" then DisableState tmp(x)' : debugstr = debugstr & tmp(x).name & " state'd" & vbnewline
          tmp(x).intensityscale = 0.001 ' this can prevent init stuttering
        Next
      Else
        if typename(obj(idx)) = "Light" then DisableState obj(idx)' : debugstr = debugstr & obj(idx).name & " state'd (not array)" & vbnewline
        obj(idx).intensityscale = 0.001 ' this can prevent init stuttering
      end if
    Next
    'debug.print debugstr
  End Sub
  Private Sub DisableState(ByRef aObj) : aObj.FadeSpeedUp = 1000 : aObj.State = 1 : End Sub 'turn state to 1

  Public Sub Init() 'Just runs TurnOnStates right now
    TurnOnStates
  End Sub

  Public Property Let Modulate(aIdx, aCoef) : Mult(aIdx) = aCoef : Lock(aIdx) = False : Loaded(aIdx) = False: End Property
  Public Property Get Modulate(aIdx) : Modulate = Mult(aIdx) : End Property

  Public Sub Update1()   'Handle all boolean numeric fading. If done fading, Lock(x) = True. Update on a '1' interval Timer!
    dim x : for x = 0 to uBound(OnOff)
      if not Lock(x) then 'and not Loaded(x) then
        if OnOff(x) then 'Fade Up
          Lvl(x) = Lvl(x) + FadeSpeedUp(x)
          if Lvl(x) >= 1 then Lvl(x) = 1 : Lock(x) = True
        elseif Not OnOff(x) then 'fade down
          Lvl(x) = Lvl(x) - FadeSpeedDown(x)
          if Lvl(x) <= 0 then Lvl(x) = 0 : Lock(x) = True
        end if
      end if
    Next
  End Sub

  Public Sub Update2()   'Both updates on -1 timer (Lowest latency, but less accurate fading at 60fps vsync)
    FrameTime = gametime - InitFrame : InitFrame = GameTime 'Calculate frametime
    dim x : for x = 0 to uBound(OnOff)
      if not Lock(x) then 'and not Loaded(x) then
        if OnOff(x) then 'Fade Up
          Lvl(x) = Lvl(x) + FadeSpeedUp(x) * FrameTime
          if Lvl(x) >= 1 then Lvl(x) = 1 : Lock(x) = True
        elseif Not OnOff(x) then 'fade down
          Lvl(x) = Lvl(x) - FadeSpeedDown(x) * FrameTime
          if Lvl(x) <= 0 then Lvl(x) = 0 : Lock(x) = True
        end if
      end if
    Next
    Update
  End Sub

  Public Sub Update() 'Handle object updates. Update on a -1 Timer! If done fading, loaded(x) = True
    dim x,xx : for x = 0 to uBound(OnOff)
      if not Loaded(x) then
        if IsArray(obj(x) ) Then  'if array
          If UseFunction then
            for each xx in obj(x) : xx.IntensityScale = cFilter(Lvl(x)*Mult(x)) : Next
          Else
            for each xx in obj(x) : xx.IntensityScale = Lvl(x)*Mult(x) : Next
          End If
        else            'if single lamp or flasher
          If UseFunction then
            obj(x).Intensityscale = cFilter(Lvl(x)*Mult(x))
          Else
            obj(x).Intensityscale = Lvl(x)
          End If
        end if
        if TypeName(lvl(x)) <> "Double" and typename(lvl(x)) <> "Integer" then msgbox "uhh " & 2 & " = " & lvl(x)
        'If UseCallBack(x) then execute cCallback(x) & " " & (Lvl(x)) 'Callback
        If UseCallBack(x) then Proc name & x,Lvl(x)*mult(x) 'Proc
        If Lock(x) Then
          if Lvl(x) = 1 or Lvl(x) = 0 then Loaded(x) = True 'finished fading
        end if
      end if
    Next
  End Sub
End Class


'Helper functions

Sub Proc(string, Callback)  'proc using a string and one argument
  'On Error Resume Next
  dim p : Set P = GetRef(String)
  P Callback
  If err.number = 13 then  msgbox "Proc error! No such procedure: " & vbnewline & string
  if err.number = 424 then msgbox "Proc error! No such Object"
End Sub


Function AppendArray(ByVal aArray, aInput)  'append one value, object, or Array onto the end of a 1 dimensional array
  if IsArray(aInput) then 'Input is an array...
    dim tmp : tmp = aArray
    If not IsArray(aArray) Then 'if not array, create an array
      tmp = aInput
    Else            'Append existing array with aInput array
      Redim Preserve tmp(uBound(aArray) + uBound(aInput)+1) 'If existing array, increase bounds by uBound of incoming array
      dim x : for x = 0 to uBound(aInput)
        if isObject(aInput(x)) then
          Set tmp(x+uBound(aArray)+1 ) = aInput(x)
        Else
          tmp(x+uBound(aArray)+1 ) = aInput(x)
        End If
      Next
    AppendArray = tmp  'return new array
    End If
  Else 'Input is NOT an array...
    If not IsArray(aArray) Then 'if not array, create an array
      aArray = Array(aArray, aInput)
    Else
      Redim Preserve aArray(uBound(aArray)+1) 'If array, increase bounds by 1
      if isObject(aInput) then
        Set aArray(uBound(aArray)) = aInput
      Else
        aArray(uBound(aArray)) = aInput
      End If
    End If
    AppendArray = aArray 'return new array
  End If
End Function


dim GameOnFF
GameOnFF = 0





'**************************************************************
'Ball color RGB change
'**************************************************************
Sub RGB1_Hit 'Darkens ball in the shooter lane
  activeball.color = RGB(100,100,100)
End Sub
Sub RGB1_Unhit
  activeball.color = RGB(255,255,255)
End Sub

Sub RGB2_Hit 'Darkens ball in the shooter lane
  activeball.color = RGB(100,100,100)
End Sub
Sub RGB2_Unhit
  activeball.color = RGB(255,255,255)
End Sub

Sub RGB3_Hit 'Darkens ball in the shooter lane
  activeball.color = RGB(100,100,100)
End Sub
Sub RGB3_Unhit
  activeball.color = RGB(255,255,255)
End Sub
'
'Sub RGB4_Hit 'Darkens ball in the shooter lane
' activeball.color = RGB(100,100,100)
'End Sub
'Sub RGB4_Unhit
' activeball.color = RGB(255,255,255)
'End Sub




'******************************************************
'****  FLIPPER CORRECTIONS by nFozzy
'******************************************************


'******************************************************
' Flippers Polarity (Select appropriate sub based on era)
'******************************************************

dim LF : Set LF = New FlipperPolarity
dim RF : Set RF = New FlipperPolarity

InitPolarity


'*******************************************
'  Late 80's early 90's

Sub InitPolarity()
  dim x, a : a = Array(LF, RF)
  for each x in a
    x.AddPoint "Ycoef", 0, RightFlipper.Y-65, 1        'disabled
    x.AddPoint "Ycoef", 1, RightFlipper.Y-11, 1
    x.enabled = True
    x.TimeDelay = 60
  Next

  AddPt "Polarity", 0, 0, 0
  AddPt "Polarity", 1, 0.05, -5
  AddPt "Polarity", 2, 0.4, -5
  AddPt "Polarity", 3, 0.6, -4.5
  AddPt "Polarity", 4, 0.65, -4.0
  AddPt "Polarity", 5, 0.7, -3.5
  AddPt "Polarity", 6, 0.75, -3.0
  AddPt "Polarity", 7, 0.8, -2.5
  AddPt "Polarity", 8, 0.85, -2.0
  AddPt "Polarity", 9, 0.9,-1.5
  AddPt "Polarity", 10, 0.95, -1.0
  AddPt "Polarity", 11, 1, -0.5
  AddPt "Polarity", 12, 1.1, 0
  AddPt "Polarity", 13, 1.3, 0

  addpt "Velocity", 0, 0,         1
  addpt "Velocity", 1, 0.16, 1.06
  addpt "Velocity", 2, 0.41,         1.05
  addpt "Velocity", 3, 0.53,         1'0.982
  addpt "Velocity", 4, 0.702, 0.968
  addpt "Velocity", 5, 0.95,  0.968
  addpt "Velocity", 6, 1.03,         0.945

  LF.Object = LeftFlipper
  LF.EndPoint = EndPointLp
  RF.Object = RightFlipper
  RF.EndPoint = EndPointRp
End Sub



' Flipper trigger hit subs
Sub TriggerLF_Hit() : LF.Addball activeball : End Sub
Sub TriggerLF_UnHit() : LF.PolarityCorrect activeball : End Sub
Sub TriggerRF_Hit() : RF.Addball activeball : End Sub
Sub TriggerRF_UnHit() : RF.PolarityCorrect activeball : End Sub




'******************************************************
'  FLIPPER CORRECTION FUNCTIONS
'******************************************************

Class FlipperPolarity
  Public DebugOn, Enabled
  Private FlipAt        'Timer variable (IE 'flip at 723,530ms...)
  Public TimeDelay        'delay before trigger turns off and polarity is disabled TODO set time!
  private Flipper, FlipperStart,FlipperEnd, FlipperEndY, LR, PartialFlipCoef
  Private Balls(20), balldata(20)

  dim PolarityIn, PolarityOut
  dim VelocityIn, VelocityOut
  dim YcoefIn, YcoefOut
  Public Sub Class_Initialize
    redim PolarityIn(0) : redim PolarityOut(0) : redim VelocityIn(0) : redim VelocityOut(0) : redim YcoefIn(0) : redim YcoefOut(0)
    Enabled = True : TimeDelay = 50 : LR = 1:  dim x : for x = 0 to uBound(balls) : balls(x) = Empty : set Balldata(x) = new SpoofBall : next
  End Sub

  Public Property let Object(aInput) : Set Flipper = aInput : StartPoint = Flipper.x : End Property
  Public Property Let StartPoint(aInput) : if IsObject(aInput) then FlipperStart = aInput.x else FlipperStart = aInput : end if : End Property
  Public Property Get StartPoint : StartPoint = FlipperStart : End Property
  Public Property Let EndPoint(aInput) : FlipperEnd = aInput.x: FlipperEndY = aInput.y: End Property
  Public Property Get EndPoint : EndPoint = FlipperEnd : End Property
  Public Property Get EndPointY: EndPointY = FlipperEndY : End Property

  Public Sub AddPoint(aChooseArray, aIDX, aX, aY) 'Index #, X position, (in) y Position (out)
    Select Case aChooseArray
      case "Polarity" : ShuffleArrays PolarityIn, PolarityOut, 1 : PolarityIn(aIDX) = aX : PolarityOut(aIDX) = aY : ShuffleArrays PolarityIn, PolarityOut, 0
      Case "Velocity" : ShuffleArrays VelocityIn, VelocityOut, 1 :VelocityIn(aIDX) = aX : VelocityOut(aIDX) = aY : ShuffleArrays VelocityIn, VelocityOut, 0
      Case "Ycoef" : ShuffleArrays YcoefIn, YcoefOut, 1 :YcoefIn(aIDX) = aX : YcoefOut(aIDX) = aY : ShuffleArrays YcoefIn, YcoefOut, 0
    End Select
    if gametime > 100 then Report aChooseArray
  End Sub

  Public Sub Report(aChooseArray)         'debug, reports all coords in tbPL.text
    if not DebugOn then exit sub
    dim a1, a2 : Select Case aChooseArray
      case "Polarity" : a1 = PolarityIn : a2 = PolarityOut
      Case "Velocity" : a1 = VelocityIn : a2 = VelocityOut
      Case "Ycoef" : a1 = YcoefIn : a2 = YcoefOut
        case else :tbpl.text = "wrong string" : exit sub
    End Select
    dim str, x : for x = 0 to uBound(a1) : str = str & aChooseArray & " x: " & round(a1(x),4) & ", " & round(a2(x),4) & vbnewline : next
    tbpl.text = str
  End Sub

  Public Sub AddBall(aBall) : dim x : for x = 0 to uBound(balls) : if IsEmpty(balls(x)) then set balls(x) = aBall : exit sub :end if : Next  : End Sub

  Private Sub RemoveBall(aBall)
    dim x : for x = 0 to uBound(balls)
      if TypeName(balls(x) ) = "IBall" then
        if aBall.ID = Balls(x).ID Then
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
    dim x : for x = 0 to uBound(balls)
      if not IsEmpty(balls(x) ) then
        pos = pSlope(Balls(x).x, FlipperStart, 0, FlipperEnd, 1)
      End If
    Next
  End Property

  Public Sub ProcessBalls() 'save data of balls in flipper range
    FlipAt = GameTime
    dim x : for x = 0 to uBound(balls)
      if not IsEmpty(balls(x) ) then
        balldata(x).Data = balls(x)
      End If
    Next
    PartialFlipCoef = ((Flipper.StartAngle - Flipper.CurrentAngle) / (Flipper.StartAngle - Flipper.EndAngle))
    PartialFlipCoef = abs(PartialFlipCoef-1)
  End Sub
  Private Function FlipperOn() : if gameTime < FlipAt+TimeDelay then FlipperOn = True : End If : End Function        'Timer shutoff for polaritycorrect

  Public Sub PolarityCorrect(aBall)
    if FlipperOn() then
      dim tmp, BallPos, x, IDX, Ycoef : Ycoef = 1

      'y safety Exit
      if aBall.VelY > -8 then 'ball going down
        RemoveBall aBall
        exit Sub
      end if

      'Find balldata. BallPos = % on Flipper
      for x = 0 to uBound(Balls)
        if aBall.id = BallData(x).id AND not isempty(BallData(x).id) then
          idx = x
          BallPos = PSlope(BallData(x).x, FlipperStart, 0, FlipperEnd, 1)
          if ballpos > 0.65 then  Ycoef = LinearEnvelope(BallData(x).Y, YcoefIn, YcoefOut)                                'find safety coefficient 'ycoef' data
        end if
      Next

      If BallPos = 0 Then 'no ball data meaning the ball is entering and exiting pretty close to the same position, use current values.
        BallPos = PSlope(aBall.x, FlipperStart, 0, FlipperEnd, 1)
        if ballpos > 0.65 then  Ycoef = LinearEnvelope(aBall.Y, YcoefIn, YcoefOut)                                                'find safety coefficient 'ycoef' data
      End If

      'Velocity correction
      if not IsEmpty(VelocityIn(0) ) then
        Dim VelCoef
        VelCoef = LinearEnvelope(BallPos, VelocityIn, VelocityOut)

        if partialflipcoef < 1 then VelCoef = PSlope(partialflipcoef, 0, 1, 1, VelCoef)

        if Enabled then aBall.Velx = aBall.Velx*VelCoef
        if Enabled then aBall.Vely = aBall.Vely*VelCoef
      End If

      'Polarity Correction (optional now)
      if not IsEmpty(PolarityIn(0) ) then
        If StartPoint > EndPoint then LR = -1        'Reverse polarity if left flipper
        dim AddX : AddX = LinearEnvelope(BallPos, PolarityIn, PolarityOut) * LR

        if Enabled then aBall.VelX = aBall.VelX + 1 * (AddX*ycoef*PartialFlipcoef)
      End If
    End If
    RemoveBall aBall
  End Sub
End Class





'******************************************************
'  SLINGSHOT CORRECTION FUNCTIONS
'******************************************************

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
  AddSlingsPt 0, 0.00,  -6
  AddSlingsPt 1, 0.45,  -8
  AddSlingsPt 2, 0.48,  0
  AddSlingsPt 3, 0.52,  0
  AddSlingsPt 4, 0.55,  8
  AddSlingsPt 5, 1.00,  6

End Sub


Sub AddSlingsPt(idx, aX, aY)        'debugger wrapper for adjusting flipper script in-game
  dim a : a = Array(LS, RS)
  dim x : for each x in a
    x.addpoint idx, aX, aY
  Next
End Sub


Class SlingshotCorrection
  Public DebugOn, Enabled
  private Slingshot, SlingX1, SlingX2, SlingY1, SlingY2

  Public ModIn, ModOut
  Private Sub Class_Initialize : redim ModIn(0) : redim Modout(0): Enabled = True : End Sub

  Public Property let Object(aInput) : Set Slingshot = aInput : End Property
  Public Property Let EndPoint1(aInput) : SlingX1 = aInput.x: SlingY1 = aInput.y: End Property
  Public Property Let EndPoint2(aInput) : SlingX2 = aInput.x: SlingY2 = aInput.y: End Property

  Public Sub AddPoint(aIdx, aX, aY)
    ShuffleArrays ModIn, ModOut, 1 : ModIn(aIDX) = aX : ModOut(aIDX) = aY : ShuffleArrays ModIn, ModOut, 0
    If gametime > 100 then Report
  End Sub

  Public Sub Report()         'debug, reports all coords in tbPL.text
    If not debugOn then exit sub
    dim a1, a2 : a1 = ModIn : a2 = ModOut
    dim str, x : for x = 0 to uBound(a1) : str = str & x & ": " & round(a1(x),4) & ", " & round(a2(x),4) & vbnewline : next
    TBPout.text = str
  End Sub


  Public Sub VelocityCorrect(aBall)
    dim BallPos, XL, XR, YL, YR

    'Assign right and left end points
    If SlingX1 < SlingX2 Then
      XL = SlingX1 : YL = SlingY1 : XR = SlingX2 : YR = SlingY2
    Else
      XL = SlingX2 : YL = SlingY2 : XR = SlingX1 : YR = SlingY1
    End If

    'Find BallPos = % on Slingshot
    If Not IsEmpty(aBall.id) Then
      If ABS(XR-XL) > ABS(YR-YL) Then
        BallPos = PSlope(aBall.x, XL, 0, XR, 1)
      Else
        BallPos = PSlope(aBall.y, YL, 0, YR, 1)
      End If
      If BallPos < 0 Then BallPos = 0
      If BallPos > 1 Then BallPos = 1
    End If

    'Velocity angle correction
    If not IsEmpty(ModIn(0) ) then
      Dim Angle, RotVxVy
      Angle = LinearEnvelope(BallPos, ModIn, ModOut)
      'debug.print " BallPos=" & BallPos &" Angle=" & Angle
      'debug.print " BEFORE: aBall.Velx=" & aBall.Velx &" aBall.Vely" & aBall.Vely
      RotVxVy = RotPoint(aBall.Velx,aBall.Vely,Angle)
      If Enabled then aBall.Velx = RotVxVy(0)
      If Enabled then aBall.Vely = RotVxVy(1)
      'debug.print " AFTER: aBall.Velx=" & aBall.Velx &" aBall.Vely" & aBall.Vely
      'debug.print " "
    End If
  End Sub

End Class


Function RotPoint(x,y,angle)
    dim rx, ry
    rx = x*dCos(angle) - y*dSin(angle)
    ry = x*dSin(angle) + y*dCos(angle)
    RotPoint = Array(rx,ry)
End Function




'******************************************************
'  FLIPPER POLARITY. RUBBER DAMPENER, AND SLINGSHOT CORRECTION SUPPORTING FUNCTIONS
'******************************************************


Sub AddPt(aStr, idx, aX, aY)        'debugger wrapper for adjusting flipper script in-game
  dim a : a = Array(LF, RF)
  dim x : for each x in a
    x.addpoint aStr, idx, aX, aY
  Next
End Sub


' Used for flipper correction and rubber dampeners
Sub ShuffleArray(ByRef aArray, byVal offset) 'shuffle 1d array
  dim x, aCount : aCount = 0
  redim a(uBound(aArray) )
  for x = 0 to uBound(aArray)        'Shuffle objects in a temp array
    if not IsEmpty(aArray(x) ) Then
      if IsObject(aArray(x)) then
        Set a(aCount) = aArray(x)
      Else
        a(aCount) = aArray(x)
      End If
      aCount = aCount + 1
    End If
  Next
  if offset < 0 then offset = 0
  redim aArray(aCount-1+offset)        'Resize original array
  for x = 0 to aCount-1                'set objects back into original array
    if IsObject(a(x)) then
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
  BallSpeed = SQR(ball.VelX^2 + ball.VelY^2 + ball.VelZ^2)
End Function

' Used for flipper correction and rubber dampeners
Function PSlope(Input, X1, Y1, X2, Y2)        'Set up line via two points, no clamping. Input X, output Y
  dim x, y, b, m : x = input : m = (Y2 - Y1) / (X2 - X1) : b = Y2 - m*X2
  Y = M*x+b
  PSlope = Y
End Function

' Used for flipper correction
Class spoofball
  Public X, Y, Z, VelX, VelY, VelZ, ID, Mass, Radius
  Public Property Let Data(aBall)
    With aBall
      x = .x : y = .y : z = .z : velx = .velx : vely = .vely : velz = .velz
      id = .ID : mass = .mass : radius = .radius
    end with
  End Property
  Public Sub Reset()
    x = Empty : y = Empty : z = Empty  : velx = Empty : vely = Empty : velz = Empty
    id = Empty : mass = Empty : radius = Empty
  End Sub
End Class

' Used for flipper correction and rubber dampeners
Function LinearEnvelope(xInput, xKeyFrame, yLvl)
  dim y 'Y output
  dim L 'Line
  dim ii : for ii = 1 to uBound(xKeyFrame)        'find active line
    if xInput <= xKeyFrame(ii) then L = ii : exit for : end if
  Next
  if xInput > xKeyFrame(uBound(xKeyFrame) ) then L = uBound(xKeyFrame)        'catch line overrun
  Y = pSlope(xInput, xKeyFrame(L-1), yLvl(L-1), xKeyFrame(L), yLvl(L) )

  if xInput <= xKeyFrame(lBound(xKeyFrame) ) then Y = yLvl(lBound(xKeyFrame) )         'Clamp lower
  if xInput >= xKeyFrame(uBound(xKeyFrame) ) then Y = yLvl(uBound(xKeyFrame) )        'Clamp upper

  LinearEnvelope = Y
End Function


'******************************************************
'  FLIPPER TRICKS
'******************************************************

RightFlipper.timerinterval=1
Rightflipper.timerenabled=True

sub RightFlipper_timer()
  FlipperTricks LeftFlipper, LFPress, LFCount, LFEndAngle, LFState
  FlipperTricks RightFlipper, RFPress, RFCount, RFEndAngle, RFState
  FlipperNudge RightFlipper, RFEndAngle, RFEOSNudge, LeftFlipper, LFEndAngle
  FlipperNudge LeftFlipper, LFEndAngle, LFEOSNudge,  RightFlipper, RFEndAngle
end sub

Dim LFEOSNudge, RFEOSNudge

Sub FlipperNudge(Flipper1, Endangle1, EOSNudge1, Flipper2, EndAngle2)
  Dim b, BOT
  BOT = GetBalls

  If Flipper1.currentangle = Endangle1 and EOSNudge1 <> 1 Then
    EOSNudge1 = 1
    'debug.print Flipper1.currentangle &" = "& Endangle1 &"--"& Flipper2.currentangle &" = "& EndAngle2
    If Flipper2.currentangle = EndAngle2 Then
      For b = 0 to Ubound(BOT)
        If FlipperTrigger(BOT(b).x, BOT(b).y, Flipper1) Then
          'Debug.Print "ball in flip1. exit"
          exit Sub
        end If
      Next
      For b = 0 to Ubound(BOT)
        If FlipperTrigger(BOT(b).x, BOT(b).y, Flipper2) Then
          BOT(b).velx = BOT(b).velx / 1.3
          BOT(b).vely = BOT(b).vely - 0.5
        end If
      Next
    End If
  Else
    If Abs(Flipper1.currentangle) > Abs(EndAngle1) + 30 then EOSNudge1 = 0
  End If
End Sub

'*****************
' Maths
'*****************
Dim PI: PI = 4*Atn(1)

Function dSin(degrees)
  dsin = sin(degrees * Pi/180)
End Function

Function dCos(degrees)
  dcos = cos(degrees * Pi/180)
End Function

Function Atn2(dy, dx)
  If dx > 0 Then
    Atn2 = Atn(dy / dx)
  ElseIf dx < 0 Then
    If dy = 0 Then
      Atn2 = pi
    Else
      Atn2 = Sgn(dy) * (pi - Atn(Abs(dy / dx)))
    end if
  ElseIf dx = 0 Then
    if dy = 0 Then
      Atn2 = 0
    else
      Atn2 = Sgn(dy) * pi / 2
    end if
  End If
End Function

'*************************************************
'  Check ball distance from Flipper for Rem
'*************************************************

Function Distance(ax,ay,bx,by)
  Distance = SQR((ax - bx)^2 + (ay - by)^2)
End Function

Function DistancePL(px,py,ax,ay,bx,by) ' Distance between a point and a line where point is px,py
  DistancePL = ABS((by - ay)*px - (bx - ax) * py + bx*ay - by*ax)/Distance(ax,ay,bx,by)
End Function

Function DistanceFast(x, y)
  dim ratio, ax, ay
  ax = abs(x)         'Get absolute value of each vector
  ay = abs(y)
  ratio = 1 / max(ax, ay)   'Create a ratio
  ratio = ratio * (1.29289 - (ax + ay) * ratio * 0.29289)
  if ratio > 0 then     'Quickly determine if it's worth using
    DistanceFast = 1/ratio
  Else
    DistanceFast = 0
  End if
end Function

Function max(a,b)
  if a > b then
    max = a
  Else
    max = b
  end if
end Function


Function Radians(Degrees)
  Radians = Degrees * PI /180
End Function

Function AnglePP(ax,ay,bx,by)
  AnglePP = Atn2((by - ay),(bx - ax))*180/PI
End Function

Function DistanceFromFlipper(ballx, bally, Flipper)
  DistanceFromFlipper = DistancePL(ballx, bally, Flipper.x, Flipper.y, Cos(Radians(Flipper.currentangle+90))+Flipper.x, Sin(Radians(Flipper.currentangle+90))+Flipper.y)
End Function

Function FlipperTrigger(ballx, bally, Flipper)
  Dim DiffAngle
  DiffAngle  = ABS(Flipper.currentangle - AnglePP(Flipper.x, Flipper.y, ballx, bally) - 90)
  If DiffAngle > 180 Then DiffAngle = DiffAngle - 360

  If DistanceFromFlipper(ballx,bally,Flipper) < 48 and DiffAngle <= 90 and Distance(ballx,bally,Flipper.x,Flipper.y) < Flipper.Length Then
    FlipperTrigger = True
  Else
    FlipperTrigger = False
  End If
End Function


'*************************************************
'  End - Check ball distance from Flipper for Rem
'*************************************************

dim LFPress, RFPress, LFCount, RFCount
dim LFState, RFState
dim EOST, EOSA,Frampup, FElasticity,FReturn
dim RFEndAngle, LFEndAngle

LFState = 1
RFState = 1
EOST = leftflipper.eostorque
EOSA = leftflipper.eostorqueangle
Frampup = LeftFlipper.rampup
FElasticity = LeftFlipper.elasticity
FReturn = LeftFlipper.return
Const EOSTnew = 0.8 '90's and later
Const EOSAnew = 1
Const EOSRampup = 0
Const SOSRampup = 2.5

Const LiveCatch = 16
Const LiveElasticity = 0.45
Const SOSEM = 0.815
Const EOSReturn = 0.035  'mid 80's to early 90's

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
  Flipper.eostorque = EOST*EOSReturn/FReturn


  If Abs(Flipper.currentangle) <= Abs(Flipper.endangle) + 0.1 Then
    Dim b, BOT
    BOT = GetBalls

    For b = 0 to UBound(BOT)
      If Distance(BOT(b).x, BOT(b).y, Flipper.x, Flipper.y) < 55 Then 'check for cradle
        If BOT(b).vely >= -0.4 Then BOT(b).vely = -0.4
      End If
    Next
  End If
End Sub

Sub FlipperTricks (Flipper, FlipperPress, FCount, FEndAngle, FState)
  Dim Dir
  Dir = Flipper.startangle/Abs(Flipper.startangle)        '-1 for Right Flipper

  If Abs(Flipper.currentangle) > Abs(Flipper.startangle) - 0.05 Then
    If FState <> 1 Then
      Flipper.rampup = SOSRampup
      Flipper.endangle = FEndAngle - 3*Dir
      Flipper.Elasticity = FElasticity * SOSEM
      FCount = 0
      FState = 1
    End If
  ElseIf Abs(Flipper.currentangle) <= Abs(Flipper.endangle) and FlipperPress = 1 then
    if FCount = 0 Then FCount = GameTime

    If FState <> 2 Then
      Flipper.eostorqueangle = EOSAnew
      Flipper.eostorque = EOSTnew
      Flipper.rampup = EOSRampup
      Flipper.endangle = FEndAngle
      FState = 2
    End If
  Elseif Abs(Flipper.currentangle) > Abs(Flipper.endangle) + 0.01 and FlipperPress = 1 Then
    If FState <> 3 Then
      Flipper.eostorque = EOST
      Flipper.eostorqueangle = EOSA
      Flipper.rampup = Frampup
      Flipper.Elasticity = FElasticity
      FState = 3
    End If

  End If
End Sub

Const LiveDistanceMin = 30  'minimum distance in vp units from flipper base live catch dampening will occur
Const LiveDistanceMax = 114  'maximum distance in vp units from flipper base live catch dampening will occur (tip protection)

Sub CheckLiveCatch(ball, Flipper, FCount, parm) 'Experimental new live catch
  Dim Dir
  Dir = Flipper.startangle/Abs(Flipper.startangle)    '-1 for Right Flipper
  Dim LiveCatchBounce                                                                                                                        'If live catch is not perfect, it won't freeze ball totally
  Dim CatchTime : CatchTime = GameTime - FCount

  if CatchTime <= LiveCatch and parm > 6 and ABS(Flipper.x - ball.x) > LiveDistanceMin and ABS(Flipper.x - ball.x) < LiveDistanceMax Then
    if CatchTime <= LiveCatch*0.5 Then                                                'Perfect catch only when catch time happens in the beginning of the window
      LiveCatchBounce = 0
    else
      LiveCatchBounce = Abs((LiveCatch/2) - CatchTime)        'Partial catch when catch happens a bit late
    end If

    If LiveCatchBounce = 0 and ball.velx * Dir > 0 Then ball.velx = 0
    ball.vely = LiveCatchBounce * (32 / LiveCatch) ' Multiplier for inaccuracy bounce
    ball.angmomx= 0
    ball.angmomy= 0
    ball.angmomz= 0
    Else
        If Abs(Flipper.currentangle) <= Abs(Flipper.endangle) + 1 Then FlippersD.Dampenf Activeball, parm
  End If
End Sub


'******************************************************
'****  END FLIPPER CORRECTIONS
'******************************************************








'******************************************************
'****  PHYSICS DAMPENERS
'******************************************************
'
' These are data mined bounce curves,
' dialed in with the in-game elasticity as much as possible to prevent angle / spin issues.
' Requires tracking ballspeed to calculate COR



Sub dPosts_Hit(idx)
  RubbersD.dampen Activeball
  TargetBouncer Activeball, 1
End Sub

Sub dSleeves_Hit(idx)
  SleevesD.Dampen Activeball
  TargetBouncer Activeball, 0.7
End Sub


dim RubbersD : Set RubbersD = new Dampener        'frubber
RubbersD.name = "Rubbers"
RubbersD.debugOn = False        'shows info in textbox "TBPout"
RubbersD.Print = False        'debug, reports in debugger (in vel, out cor)
'cor bounce curve (linear)
'for best results, try to match in-game velocity as closely as possible to the desired curve
'RubbersD.addpoint 0, 0, 0.935        'point# (keep sequential), ballspeed, CoR (elasticity)
RubbersD.addpoint 0, 0, 1.1        'point# (keep sequential), ballspeed, CoR (elasticity)
RubbersD.addpoint 1, 3.77, 0.97
RubbersD.addpoint 2, 5.76, 0.967        'dont take this as gospel. if you can data mine rubber elasticitiy, please help!
RubbersD.addpoint 3, 15.84, 0.874
RubbersD.addpoint 4, 56, 0.64        'there's clamping so interpolate up to 56 at least

dim SleevesD : Set SleevesD = new Dampener        'this is just rubber but cut down to 85%...
SleevesD.name = "Sleeves"
SleevesD.debugOn = False        'shows info in textbox "TBPout"
SleevesD.Print = False        'debug, reports in debugger (in vel, out cor)
SleevesD.CopyCoef RubbersD, 0.85

'######################### Add new FlippersD Profile
'#########################    Adjust these values to increase or lessen the elasticity

dim FlippersD : Set FlippersD = new Dampener
FlippersD.name = "Flippers"
FlippersD.debugOn = False
FlippersD.Print = False
FlippersD.addpoint 0, 0, 1.1
FlippersD.addpoint 1, 3.77, 0.99
FlippersD.addpoint 2, 6, 0.99

Class Dampener
  Public Print, debugOn 'tbpOut.text
  public name, Threshold         'Minimum threshold. Useful for Flippers, which don't have a hit threshold.
  Public ModIn, ModOut
  Private Sub Class_Initialize : redim ModIn(0) : redim Modout(0): End Sub

  Public Sub AddPoint(aIdx, aX, aY)
    ShuffleArrays ModIn, ModOut, 1 : ModIn(aIDX) = aX : ModOut(aIDX) = aY : ShuffleArrays ModIn, ModOut, 0
    if gametime > 100 then Report
  End Sub

  public sub Dampen(aBall)
    if threshold then if BallSpeed(aBall) < threshold then exit sub end if end if
    dim RealCOR, DesiredCOR, str, coef
    DesiredCor = LinearEnvelope(cor.ballvel(aBall.id), ModIn, ModOut )
    RealCOR = BallSpeed(aBall) / (cor.ballvel(aBall.id)+0.0001)
    coef = desiredcor / realcor
    if debugOn then str = name & " in vel:" & round(cor.ballvel(aBall.id),2 ) & vbnewline & "desired cor: " & round(desiredcor,4) & vbnewline & _
    "actual cor: " & round(realCOR,4) & vbnewline & "ballspeed coef: " & round(coef, 3) & vbnewline
    if Print then debug.print Round(cor.ballvel(aBall.id),2) & ", " & round(desiredcor,3)

' Thalamus - patched :     aBall.velx = aBall.velx * coef : aBall.vely = aBall.vely * coef
    aBall.velx = aBall.velx * coef : aBall.vely = aBall.vely * coef : aBall.velz = aBall.velz * coef
    if debugOn then TBPout.text = str
  End Sub

  public sub Dampenf(aBall, parm) 'Rubberizer is handle here
    dim RealCOR, DesiredCOR, str, coef
    DesiredCor = LinearEnvelope(cor.ballvel(aBall.id), ModIn, ModOut )
    RealCOR = BallSpeed(aBall) / (cor.ballvel(aBall.id)+0.0001)
    coef = desiredcor / realcor
    If abs(aball.velx) < 2 and aball.vely < 0 and aball.vely > -3.75 then
' Thalamus - patched :       aBall.velx = aBall.velx * coef : aBall.vely = aBall.vely * coef
    aBall.velx = aBall.velx * coef : aBall.vely = aBall.vely * coef : aBall.velz = aBall.velz * coef
    End If
  End Sub

  Public Sub CopyCoef(aObj, aCoef) 'alternative addpoints, copy with coef
    dim x : for x = 0 to uBound(aObj.ModIn)
      addpoint x, aObj.ModIn(x), aObj.ModOut(x)*aCoef
    Next
  End Sub


  Public Sub Report()         'debug, reports all coords in tbPL.text
    if not debugOn then exit sub
    dim a1, a2 : a1 = ModIn : a2 = ModOut
    dim str, x : for x = 0 to uBound(a1) : str = str & x & ": " & round(a1(x),4) & ", " & round(a2(x),4) & vbnewline : next
    TBPout.text = str
  End Sub

End Class



'******************************************************
'  TRACK ALL BALL VELOCITIES
'  FOR RUBBER DAMPENER AND DROP TARGETS
'******************************************************

dim cor : set cor = New CoRTracker

Class CoRTracker
  public ballvel, ballvelx, ballvely

  Private Sub Class_Initialize : redim ballvel(0) : redim ballvelx(0): redim ballvely(0) : End Sub

  Public Sub Update() 'tracks in-ball-velocity
    dim str, b, AllBalls, highestID : allBalls = getballs

    for each b in allballs
      if b.id >= HighestID then highestID = b.id
    Next

    if uBound(ballvel) < highestID then redim ballvel(highestID)  'set bounds
    if uBound(ballvelx) < highestID then redim ballvelx(highestID)  'set bounds
    if uBound(ballvely) < highestID then redim ballvely(highestID)  'set bounds

    for each b in allballs
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
' VPW TargetBouncer for targets and posts by Iaakki, Wrd1972, Apophis
'******************************************************

Const TargetBouncerEnabled = 1    '0 = normal standup targets, 1 = bouncy targets
Const TargetBouncerFactor = 0.7   'Level of bounces. Recommmended value of 0.7

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








'******************************************************
'****  FLEEP MECHANICAL SOUNDS
'******************************************************

' This part in the script is an entire block that is dedicated to the physics sound system.
' Various scripts and sounds that may be pretty generic and could suit other WPC systems, but the most are tailored specifically for the TOM table

' Many of the sounds in this package can be added by creating collections and adding the appropriate objects to those collections.
' Create the following new collections:
'   Metals (all metal objects, metal walls, metal posts, metal wire guides)
'   Apron (the apron walls and plunger wall)
'   Walls (all wood or plastic walls)
'   Rollovers (wire rollover triggers, star triggers, or button triggers)
'   Targets (standup or drop targets, these are hit sounds only ... you will want to add separate dropping sounds for drop targets)
'   Gates (plate gates)
'   GatesWire (wire gates)
'   Rubbers (all rubbers including posts, sleeves, pegs, and bands)
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
' Tutorial vides by Apophis
' Part 1:   https://youtu.be/PbE2kNiam3g
' Part 2:   https://youtu.be/B5cm1Y8wQsk
' Part 3:   https://youtu.be/eLhWyuYOyGg


'///////////////////////////////  SOUNDS PARAMETERS  //////////////////////////////
Dim GlobalSoundLevel, CoinSoundLevel, PlungerReleaseSoundLevel, PlungerPullSoundLevel, NudgeLeftSoundLevel
Dim NudgeRightSoundLevel, NudgeCenterSoundLevel, StartButtonSoundLevel, RollingSoundFactor

CoinSoundLevel = 1                            'volume level; range [0, 1]
NudgeLeftSoundLevel = 1                         'volume level; range [0, 1]
NudgeRightSoundLevel = 1                        'volume level; range [0, 1]
NudgeCenterSoundLevel = 1                       'volume level; range [0, 1]
StartButtonSoundLevel = 0.1                       'volume level; range [0, 1]
PlungerReleaseSoundLevel = 0.8 '1 wjr                     'volume level; range [0, 1]
PlungerPullSoundLevel = 1                       'volume level; range [0, 1]
RollingSoundFactor = 1.1/5

'///////////////////////-----Solenoids, Kickers and Flash Relays-----///////////////////////
Dim FlipperUpAttackMinimumSoundLevel, FlipperUpAttackMaximumSoundLevel, FlipperUpAttackLeftSoundLevel, FlipperUpAttackRightSoundLevel
Dim FlipperUpSoundLevel, FlipperDownSoundLevel, FlipperLeftHitParm, FlipperRightHitParm
Dim SlingshotSoundLevel, BumperSoundFactor, KnockerSoundLevel

FlipperUpAttackMinimumSoundLevel = 0.010                      'volume level; range [0, 1]
FlipperUpAttackMaximumSoundLevel = 0.635                'volume level; range [0, 1]
FlipperUpSoundLevel = 1.0                                   'volume level; range [0, 1]
FlipperDownSoundLevel = 0.45                                  'volume level; range [0, 1]
FlipperLeftHitParm = FlipperUpSoundLevel                'sound helper; not configurable
FlipperRightHitParm = FlipperUpSoundLevel               'sound helper; not configurable
SlingshotSoundLevel = 0.95                        'volume level; range [0, 1]
BumperSoundFactor = 4.25                        'volume multiplier; must not be zero
KnockerSoundLevel = 1                           'volume level; range [0, 1]

'///////////////////////-----Ball Drops, Bumps and Collisions-----///////////////////////
Dim RubberStrongSoundFactor, RubberWeakSoundFactor, RubberFlipperSoundFactor,BallWithBallCollisionSoundFactor
Dim BallBouncePlayfieldSoftFactor, BallBouncePlayfieldHardFactor, PlasticRampDropToPlayfieldSoundLevel, WireRampDropToPlayfieldSoundLevel, DelayedBallDropOnPlayfieldSoundLevel
Dim WallImpactSoundFactor, MetalImpactSoundFactor, SubwaySoundLevel, SubwayEntrySoundLevel, ScoopEntrySoundLevel
Dim SaucerLockSoundLevel, SaucerKickSoundLevel

BallWithBallCollisionSoundFactor = 3.2                  'volume multiplier; must not be zero
RubberStrongSoundFactor = 0.055/5                     'volume multiplier; must not be zero
RubberWeakSoundFactor = 0.075/5                     'volume multiplier; must not be zero
RubberFlipperSoundFactor = 0.075/5                    'volume multiplier; must not be zero
BallBouncePlayfieldSoftFactor = 0.025                 'volume multiplier; must not be zero
BallBouncePlayfieldHardFactor = 0.025                 'volume multiplier; must not be zero
DelayedBallDropOnPlayfieldSoundLevel = 0.8                  'volume level; range [0, 1]
WallImpactSoundFactor = 0.075                     'volume multiplier; must not be zero
MetalImpactSoundFactor = 0.075/3
SaucerLockSoundLevel = 0.8
SaucerKickSoundLevel = 0.8

'///////////////////////-----Gates, Spinners, Rollovers and Targets-----///////////////////////

Dim GateSoundLevel, TargetSoundFactor, SpinnerSoundLevel, RolloverSoundLevel, DTSoundLevel

GateSoundLevel = 0.5/5                          'volume level; range [0, 1]
TargetSoundFactor = 0.0025 * 10                     'volume multiplier; must not be zero
DTSoundLevel = 0.25                           'volume multiplier; must not be zero
RolloverSoundLevel = 0.25                                       'volume level; range [0, 1]
SpinnerSoundLevel = 0.5                                       'volume level; range [0, 1]

'///////////////////////-----Ball Release, Guides and Drain-----///////////////////////
Dim DrainSoundLevel, BallReleaseSoundLevel, BottomArchBallGuideSoundFactor, FlipperBallGuideSoundFactor

DrainSoundLevel = 0.8                           'volume level; range [0, 1]
BallReleaseSoundLevel = 1                       'volume level; range [0, 1]
BottomArchBallGuideSoundFactor = 0.2                  'volume multiplier; must not be zero
FlipperBallGuideSoundFactor = 0.015                   'volume multiplier; must not be zero

'///////////////////////-----Loops and Lanes-----///////////////////////
Dim ArchSoundFactor
ArchSoundFactor = 0.025/5                         'volume multiplier; must not be zero


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
  PlaySound playsoundparams, 0, aVol * VolumeDial, AudioPan(tableobj), 0, 0, 0, 0, AudioFade(tableobj)
End Sub

Sub PlaySoundAtLevelExistingStatic(playsoundparams, aVol, tableobj)
  PlaySound playsoundparams, 0, aVol * VolumeDial, AudioPan(tableobj), 0, 0, 1, 0, AudioFade(tableobj)
End Sub

Sub PlaySoundAtLevelStaticLoop(playsoundparams, aVol, tableobj)
  PlaySound playsoundparams, -1, aVol * VolumeDial, AudioPan(tableobj), 0, 0, 0, 0, AudioFade(tableobj)
End Sub

Sub PlaySoundAtLevelStaticRandomPitch(playsoundparams, aVol, randomPitch, tableobj)
  PlaySound playsoundparams, 0, aVol * VolumeDial, AudioPan(tableobj), randomPitch, 0, 0, 0, AudioFade(tableobj)
End Sub

Sub PlaySoundAtLevelActiveBall(playsoundparams, aVol)
  PlaySound playsoundparams, 0, aVol * VolumeDial, AudioPan(ActiveBall), 0, 0, 0, 0, AudioFade(ActiveBall)
End Sub

Sub PlaySoundAtLevelExistingActiveBall(playsoundparams, aVol)
  PlaySound playsoundparams, 0, aVol * VolumeDial, AudioPan(ActiveBall), 0, 0, 1, 0, AudioFade(ActiveBall)
End Sub

Sub PlaySoundAtLeveTimerActiveBall(playsoundparams, aVol, ballvariable)
  PlaySound playsoundparams, 0, aVol * VolumeDial, AudioPan(ballvariable), 0, 0, 0, 0, AudioFade(ballvariable)
End Sub

Sub PlaySoundAtLevelTimerExistingActiveBall(playsoundparams, aVol, ballvariable)
  PlaySound playsoundparams, 0, aVol * VolumeDial, AudioPan(ballvariable), 0, 0, 1, 0, AudioFade(ballvariable)
End Sub

Sub PlaySoundAtLevelRoll(playsoundparams, aVol, pitch)
  PlaySound playsoundparams, -1, aVol * VolumeDial, AudioPan(tableobj), randomPitch, 0, 0, 0, AudioFade(tableobj)
End Sub

' Previous Positional Sound Subs

Sub PlaySoundAt(soundname, tableobj)
  PlaySound soundname, 1, 1 * VolumeDial, AudioPan(tableobj), 0,0,0, 1, AudioFade(tableobj)
End Sub

Sub PlaySoundAtVol(soundname, tableobj, aVol)
  PlaySound soundname, 1, aVol * VolumeDial, AudioPan(tableobj), 0,0,0, 1, AudioFade(tableobj)
End Sub

Sub PlaySoundAtBall(soundname)
  PlaySoundAt soundname, ActiveBall
End Sub

Sub PlaySoundAtBallVol (Soundname, aVol)
  Playsound soundname, 1,aVol * VolumeDial, AudioPan(ActiveBall), 0,0,0, 1, AudioFade(ActiveBall)
End Sub

Sub PlaySoundAtBallVolM (Soundname, aVol)
  Playsound soundname, 1,aVol * VolumeDial, AudioPan(ActiveBall), 0,0,0, 0, AudioFade(ActiveBall)
End Sub

Sub PlaySoundAtVolLoops(sound, tableobj, Vol, Loops)
  PlaySound sound, Loops, Vol * VolumeDial, AudioPan(tableobj), 0,0,0, 1, AudioFade(tableobj)
End Sub


'******************************************************
'  Fleep  Supporting Ball & Sound Functions
'******************************************************

Function AudioFade(tableobj) ' Fades between front and back of the table (for surround systems or 2x2 speakers, etc), depending on the Y position on the table. "table1" is the name of the table
  Dim tmp
    tmp = tableobj.y * 2 / tableheight-1

  if tmp > 7000 Then
    tmp = 7000
  elseif tmp < -7000 Then
    tmp = -7000
  end if

    If tmp > 0 Then
    AudioFade = Csng(tmp ^10)
    Else
        AudioFade = Csng(-((- tmp) ^10) )
    End If
End Function

Function AudioPan(tableobj) ' Calculates the pan for a tableobj based on the X position on the table. "table1" is the name of the table
    Dim tmp
    tmp = tableobj.x * 2 / tablewidth-1

  if tmp > 7000 Then
    tmp = 7000
  elseif tmp < -7000 Then
    tmp = -7000
  end if

    If tmp > 0 Then
        AudioPan = Csng(tmp ^10)
    Else
        AudioPan = Csng(-((- tmp) ^10) )
    End If
End Function

Function Vol(ball) ' Calculates the volume of the sound based on the ball speed
  Vol = Csng(BallVel(ball) ^2)
End Function

Function Volz(ball) ' Calculates the volume of the sound based on the ball speed
  Volz = Csng((ball.velz) ^2)
End Function

Function Pitch(ball) ' Calculates the pitch of the sound based on the ball speed
  Pitch = BallVel(ball) * 20
End Function

Function BallVel(ball) 'Calculates the ball speed
  BallVel = INT(SQR((ball.VelX ^2) + (ball.VelY ^2) ) )
End Function

Function VolPlayfieldRoll(ball) ' Calculates the roll volume of the sound based on the ball speed
  VolPlayfieldRoll = RollingSoundFactor * 0.0005 * Csng(BallVel(ball) ^3)
End Function

Function PitchPlayfieldRoll(ball) ' Calculates the roll pitch of the sound based on the ball speed
  PitchPlayfieldRoll = BallVel(ball) ^2 * 15
End Function

Function RndInt(min, max)
  RndInt = Int(Rnd() * (max-min + 1) + min)' Sets a random number integer between min and max
End Function

Function RndNum(min, max)
  RndNum = Rnd() * (max-min) + min' Sets a random number between min and max
End Function

'/////////////////////////////  GENERAL SOUND SUBROUTINES  ////////////////////////////
Sub SoundStartButton()
  PlaySound ("Start_Button"), 0, StartButtonSoundLevel, 0, 0.25
End Sub

Sub SoundNudgeLeft()
  PlaySound ("Nudge_" & Int(Rnd*2)+1), 0, NudgeLeftSoundLevel * VolumeDial, -0.1, 0.25
End Sub

Sub SoundNudgeRight()
  PlaySound ("Nudge_" & Int(Rnd*2)+1), 0, NudgeRightSoundLevel * VolumeDial, 0.1, 0.25
End Sub

Sub SoundNudgeCenter()
  PlaySound ("Nudge_" & Int(Rnd*2)+1), 0, NudgeCenterSoundLevel * VolumeDial, 0, 0.25
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
  PlaySoundAtLevelStatic ("Drain_" & Int(Rnd*11)+1), DrainSoundLevel, drainswitch
End Sub

'/////////////////////////////  TROUGH BALL RELEASE SOLENOID SOUNDS  ////////////////////////////

Sub RandomSoundBallRelease(drainswitch)
  PlaySoundAtLevelStatic SoundFX("BallRelease" & Int(Rnd*7)+1,DOFContactors), BallReleaseSoundLevel, drainswitch
End Sub

'/////////////////////////////  SLINGSHOT SOLENOID SOUNDS  ////////////////////////////
Sub RandomSoundSlingshotLeft(sling)
  PlaySoundAtLevelStatic SoundFX("Sling_L" & Int(Rnd*10)+1,DOFContactors), SlingshotSoundLevel, Sling
End Sub

Sub RandomSoundSlingshotRight(sling)
  PlaySoundAtLevelStatic SoundFX("Sling_R" & Int(Rnd*8)+1,DOFContactors), SlingshotSoundLevel, Sling
End Sub

'/////////////////////////////  BUMPER SOLENOID SOUNDS  ////////////////////////////
Sub RandomSoundBumperTop(Bump)
  PlaySoundAtLevelStatic SoundFX("Bumpers_Top_" & Int(Rnd*5)+1,DOFContactors), Vol(ActiveBall) * BumperSoundFactor, Bump
End Sub

Sub RandomSoundBumperMiddle(Bump)
  PlaySoundAtLevelStatic SoundFX("Bumpers_Middle_" & Int(Rnd*5)+1,DOFContactors), Vol(ActiveBall) * BumperSoundFactor, Bump
End Sub

Sub RandomSoundBumperBottom(Bump)
  PlaySoundAtLevelStatic SoundFX("Bumpers_Bottom_" & Int(Rnd*5)+1,DOFContactors), Vol(ActiveBall) * BumperSoundFactor, Bump
End Sub

'/////////////////////////////  SPINNER SOUNDS  ////////////////////////////
Sub SoundSpinner(spinnerswitch)
  PlaySoundAtLevelStatic ("Spinner"), SpinnerSoundLevel, spinnerswitch
End Sub


'/////////////////////////////  FLIPPER BATS SOUND SUBROUTINES  ////////////////////////////
'/////////////////////////////  FLIPPER BATS SOLENOID ATTACK SOUND  ////////////////////////////
Sub SoundFlipperUpAttackLeft(flipper)
  FlipperUpAttackLeftSoundLevel = RndNum(FlipperUpAttackMinimumSoundLevel, FlipperUpAttackMaximumSoundLevel)
  PlaySoundAtLevelStatic ("Flipper_Attack-L01"), FlipperUpAttackLeftSoundLevel, flipper
End Sub

Sub SoundFlipperUpAttackRight(flipper)
  FlipperUpAttackRightSoundLevel = RndNum(FlipperUpAttackMinimumSoundLevel, FlipperUpAttackMaximumSoundLevel)
  PlaySoundAtLevelStatic ("Flipper_Attack-R01"), FlipperUpAttackLeftSoundLevel, flipper
End Sub

'/////////////////////////////  FLIPPER BATS SOLENOID CORE SOUND  ////////////////////////////
Sub RandomSoundFlipperUpLeft(flipper)
  PlaySoundAtLevelStatic SoundFX("Flipper_L0" & Int(Rnd*9)+1,DOFFlippers), FlipperLeftHitParm, Flipper
End Sub

Sub RandomSoundFlipperUpRight(flipper)
  PlaySoundAtLevelStatic SoundFX("Flipper_R0" & Int(Rnd*9)+1,DOFFlippers), FlipperRightHitParm, Flipper
End Sub

Sub RandomSoundReflipUpLeft(flipper)
  PlaySoundAtLevelStatic SoundFX("Flipper_ReFlip_L0" & Int(Rnd*3)+1,DOFFlippers), (RndNum(0.8, 1))*FlipperUpSoundLevel, Flipper
End Sub

Sub RandomSoundReflipUpRight(flipper)
  PlaySoundAtLevelStatic SoundFX("Flipper_ReFlip_R0" & Int(Rnd*3)+1,DOFFlippers), (RndNum(0.8, 1))*FlipperUpSoundLevel, Flipper
End Sub

Sub RandomSoundFlipperDownLeft(flipper)
  PlaySoundAtLevelStatic SoundFX("Flipper_Left_Down_" & Int(Rnd*7)+1,DOFFlippers), FlipperDownSoundLevel, Flipper
End Sub

Sub RandomSoundFlipperDownRight(flipper)
  PlaySoundAtLevelStatic SoundFX("Flipper_Right_Down_" & Int(Rnd*8)+1,DOFFlippers), FlipperDownSoundLevel, Flipper
End Sub

'/////////////////////////////  FLIPPER BATS BALL COLLIDE SOUND  ////////////////////////////

Sub LeftFlipperCollide(parm)
  FlipperLeftHitParm = parm/10
  If FlipperLeftHitParm > 1 Then
    FlipperLeftHitParm = 1
  End If
  FlipperLeftHitParm = FlipperUpSoundLevel * FlipperLeftHitParm
  RandomSoundRubberFlipper(parm)
End Sub

Sub RightFlipperCollide(parm)
  FlipperRightHitParm = parm/10
  If FlipperRightHitParm > 1 Then
    FlipperRightHitParm = 1
  End If
  FlipperRightHitParm = FlipperUpSoundLevel * FlipperRightHitParm
  RandomSoundRubberFlipper(parm)
End Sub

Sub RandomSoundRubberFlipper(parm)
  PlaySoundAtLevelActiveBall ("Flipper_Rubber_" & Int(Rnd*7)+1), parm  * RubberFlipperSoundFactor
End Sub

'/////////////////////////////  ROLLOVER SOUNDS  ////////////////////////////
Sub RandomSoundRollover()
  PlaySoundAtLevelActiveBall ("Rollover_" & Int(Rnd*4)+1), RolloverSoundLevel
End Sub

Sub Rollovers_Hit(idx)
  RandomSoundRollover
End Sub

'/////////////////////////////  VARIOUS PLAYFIELD SOUND SUBROUTINES  ////////////////////////////
'/////////////////////////////  RUBBERS AND POSTS  ////////////////////////////
'/////////////////////////////  RUBBERS - EVENTS  ////////////////////////////
Sub Rubbers_Hit(idx)
  dim finalspeed
  finalspeed=SQR(activeball.velx * activeball.velx + activeball.vely * activeball.vely)
  If finalspeed > 5 then
    RandomSoundRubberStrong 1
  End if
  If finalspeed <= 5 then
    RandomSoundRubberWeak()
  End If
End Sub

'/////////////////////////////  RUBBERS AND POSTS - STRONG IMPACTS  ////////////////////////////
Sub RandomSoundRubberStrong(voladj)
  Select Case Int(Rnd*10)+1
    Case 1 : PlaySoundAtLevelActiveBall ("Rubber_Strong_1"), Vol(ActiveBall) * RubberStrongSoundFactor*voladj
    Case 2 : PlaySoundAtLevelActiveBall ("Rubber_Strong_2"), Vol(ActiveBall) * RubberStrongSoundFactor*voladj
    Case 3 : PlaySoundAtLevelActiveBall ("Rubber_Strong_3"), Vol(ActiveBall) * RubberStrongSoundFactor*voladj
    Case 4 : PlaySoundAtLevelActiveBall ("Rubber_Strong_4"), Vol(ActiveBall) * RubberStrongSoundFactor*voladj
    Case 5 : PlaySoundAtLevelActiveBall ("Rubber_Strong_5"), Vol(ActiveBall) * RubberStrongSoundFactor*voladj
    Case 6 : PlaySoundAtLevelActiveBall ("Rubber_Strong_6"), Vol(ActiveBall) * RubberStrongSoundFactor*voladj
    Case 7 : PlaySoundAtLevelActiveBall ("Rubber_Strong_7"), Vol(ActiveBall) * RubberStrongSoundFactor*voladj
    Case 8 : PlaySoundAtLevelActiveBall ("Rubber_Strong_8"), Vol(ActiveBall) * RubberStrongSoundFactor*voladj
    Case 9 : PlaySoundAtLevelActiveBall ("Rubber_Strong_9"), Vol(ActiveBall) * RubberStrongSoundFactor*voladj
    Case 10 : PlaySoundAtLevelActiveBall ("Rubber_1_Hard"), Vol(ActiveBall) * RubberStrongSoundFactor * 0.6*voladj
  End Select
End Sub

'/////////////////////////////  RUBBERS AND POSTS - WEAK IMPACTS  ////////////////////////////
Sub RandomSoundRubberWeak()
  PlaySoundAtLevelActiveBall ("Rubber_" & Int(Rnd*9)+1), Vol(ActiveBall) * RubberWeakSoundFactor
End Sub

'/////////////////////////////  WALL IMPACTS  ////////////////////////////
Sub Walls_Hit(idx)
  RandomSoundWall()
End Sub

Sub RandomSoundWall()
  dim finalspeed
  finalspeed=SQR(activeball.velx * activeball.velx + activeball.vely * activeball.vely)
  If finalspeed > 16 then
    Select Case Int(Rnd*5)+1
      Case 1 : PlaySoundAtLevelExistingActiveBall ("Wall_Hit_1"), Vol(ActiveBall) * WallImpactSoundFactor
      Case 2 : PlaySoundAtLevelExistingActiveBall ("Wall_Hit_2"), Vol(ActiveBall) * WallImpactSoundFactor
      Case 3 : PlaySoundAtLevelExistingActiveBall ("Wall_Hit_5"), Vol(ActiveBall) * WallImpactSoundFactor
      Case 4 : PlaySoundAtLevelExistingActiveBall ("Wall_Hit_7"), Vol(ActiveBall) * WallImpactSoundFactor
      Case 5 : PlaySoundAtLevelExistingActiveBall ("Wall_Hit_9"), Vol(ActiveBall) * WallImpactSoundFactor
    End Select
  End if
  If finalspeed >= 6 AND finalspeed <= 16 then
    Select Case Int(Rnd*4)+1
      Case 1 : PlaySoundAtLevelExistingActiveBall ("Wall_Hit_3"), Vol(ActiveBall) * WallImpactSoundFactor
      Case 2 : PlaySoundAtLevelExistingActiveBall ("Wall_Hit_4"), Vol(ActiveBall) * WallImpactSoundFactor
      Case 3 : PlaySoundAtLevelExistingActiveBall ("Wall_Hit_6"), Vol(ActiveBall) * WallImpactSoundFactor
      Case 4 : PlaySoundAtLevelExistingActiveBall ("Wall_Hit_8"), Vol(ActiveBall) * WallImpactSoundFactor
    End Select
  End If
  If finalspeed < 6 Then
    Select Case Int(Rnd*3)+1
      Case 1 : PlaySoundAtLevelExistingActiveBall ("Wall_Hit_4"), Vol(ActiveBall) * WallImpactSoundFactor
      Case 2 : PlaySoundAtLevelExistingActiveBall ("Wall_Hit_6"), Vol(ActiveBall) * WallImpactSoundFactor
      Case 3 : PlaySoundAtLevelExistingActiveBall ("Wall_Hit_8"), Vol(ActiveBall) * WallImpactSoundFactor
    End Select
  End if
End Sub

'/////////////////////////////  METAL TOUCH SOUNDS  ////////////////////////////
Sub RandomSoundMetal()
  PlaySoundAtLevelActiveBall ("Metal_Touch_" & Int(Rnd*13)+1), Vol(ActiveBall) * MetalImpactSoundFactor
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
  dim finalspeed
  finalspeed=SQR(activeball.velx * activeball.velx + activeball.vely * activeball.vely)
  If finalspeed > 16 then
    PlaySoundAtLevelActiveBall ("Apron_Bounce_"& Int(Rnd*2)+1), Vol(ActiveBall) * BottomArchBallGuideSoundFactor
  End if
  If finalspeed >= 6 AND finalspeed <= 16 then
    Select Case Int(Rnd*2)+1
      Case 1 : PlaySoundAtLevelActiveBall ("Apron_Bounce_1"), Vol(ActiveBall) * BottomArchBallGuideSoundFactor
      Case 2 : PlaySoundAtLevelActiveBall ("Apron_Bounce_Soft_1"), Vol(ActiveBall) * BottomArchBallGuideSoundFactor
    End Select
  End If
  If finalspeed < 6 Then
    Select Case Int(Rnd*2)+1
      Case 1 : PlaySoundAtLevelActiveBall ("Apron_Bounce_Soft_1"), Vol(ActiveBall) * BottomArchBallGuideSoundFactor
      Case 2 : PlaySoundAtLevelActiveBall ("Apron_Medium_3"), Vol(ActiveBall) * BottomArchBallGuideSoundFactor
    End Select
  End if
End Sub

'/////////////////////////////  BOTTOM ARCH BALL GUIDE - HARD HITS  ////////////////////////////
Sub RandomSoundBottomArchBallGuideHardHit()
  PlaySoundAtLevelActiveBall ("Apron_Hard_Hit_" & Int(Rnd*3)+1), BottomArchBallGuideSoundFactor * 0.25
End Sub

Sub Apron_Hit (idx)
  If Abs(cor.ballvelx(activeball.id) < 4) and cor.ballvely(activeball.id) > 7 then
    RandomSoundBottomArchBallGuideHardHit()
  Else
    RandomSoundBottomArchBallGuide
  End If
End Sub

'/////////////////////////////  FLIPPER BALL GUIDE  ////////////////////////////
Sub RandomSoundFlipperBallGuide()
  dim finalspeed
  finalspeed=SQR(activeball.velx * activeball.velx + activeball.vely * activeball.vely)
  If finalspeed > 16 then
    Select Case Int(Rnd*2)+1
      Case 1 : PlaySoundAtLevelActiveBall ("Apron_Hard_1"),  Vol(ActiveBall) * FlipperBallGuideSoundFactor
      Case 2 : PlaySoundAtLevelActiveBall ("Apron_Hard_2"),  Vol(ActiveBall) * 0.8 * FlipperBallGuideSoundFactor
    End Select
  End if
  If finalspeed >= 6 AND finalspeed <= 16 then
    PlaySoundAtLevelActiveBall ("Apron_Medium_" & Int(Rnd*3)+1),  Vol(ActiveBall) * FlipperBallGuideSoundFactor
  End If
  If finalspeed < 6 Then
    PlaySoundAtLevelActiveBall ("Apron_Soft_" & Int(Rnd*7)+1),  Vol(ActiveBall) * FlipperBallGuideSoundFactor
  End If
End Sub

'/////////////////////////////  TARGET HIT SOUNDS  ////////////////////////////
Sub RandomSoundTargetHitStrong()
  PlaySoundAtLevelActiveBall SoundFX("Target_Hit_" & Int(Rnd*4)+5,DOFTargets), Vol(ActiveBall) * 0.45 * TargetSoundFactor
End Sub

Sub RandomSoundTargetHitWeak()
  PlaySoundAtLevelActiveBall SoundFX("Target_Hit_" & Int(Rnd*4)+1,DOFTargets), Vol(ActiveBall) * TargetSoundFactor
End Sub

Sub PlayTargetSound()
  dim finalspeed
  finalspeed=SQR(activeball.velx * activeball.velx + activeball.vely * activeball.vely)
  If finalspeed > 10 then
    RandomSoundTargetHitStrong()
    RandomSoundBallBouncePlayfieldSoft Activeball
  Else
    RandomSoundTargetHitWeak()
  End If
End Sub

Sub Targets_Hit (idx)
  PlayTargetSound
End Sub

'/////////////////////////////  BALL BOUNCE SOUNDS  ////////////////////////////
Sub RandomSoundBallBouncePlayfieldSoft(aBall)
  Select Case Int(Rnd*9)+1
    Case 1 : PlaySoundAtLevelStatic ("Ball_Bounce_Playfield_Soft_1"), volz(aBall) * BallBouncePlayfieldSoftFactor, aBall
    Case 2 : PlaySoundAtLevelStatic ("Ball_Bounce_Playfield_Soft_2"), volz(aBall) * BallBouncePlayfieldSoftFactor * 0.5, aBall
    Case 3 : PlaySoundAtLevelStatic ("Ball_Bounce_Playfield_Soft_3"), volz(aBall) * BallBouncePlayfieldSoftFactor * 0.8, aBall
    Case 4 : PlaySoundAtLevelStatic ("Ball_Bounce_Playfield_Soft_4"), volz(aBall) * BallBouncePlayfieldSoftFactor * 0.5, aBall
    Case 5 : PlaySoundAtLevelStatic ("Ball_Bounce_Playfield_Soft_5"), volz(aBall) * BallBouncePlayfieldSoftFactor, aBall
    Case 6 : PlaySoundAtLevelStatic ("Ball_Bounce_Playfield_Hard_1"), volz(aBall) * BallBouncePlayfieldSoftFactor * 0.2, aBall
    Case 7 : PlaySoundAtLevelStatic ("Ball_Bounce_Playfield_Hard_2"), volz(aBall) * BallBouncePlayfieldSoftFactor * 0.2, aBall
    Case 8 : PlaySoundAtLevelStatic ("Ball_Bounce_Playfield_Hard_5"), volz(aBall) * BallBouncePlayfieldSoftFactor * 0.2, aBall
    Case 9 : PlaySoundAtLevelStatic ("Ball_Bounce_Playfield_Hard_7"), volz(aBall) * BallBouncePlayfieldSoftFactor * 0.3, aBall
  End Select
End Sub

Sub RandomSoundBallBouncePlayfieldHard(aBall)
  PlaySoundAtLevelStatic ("Ball_Bounce_Playfield_Hard_" & Int(Rnd*7)+1), volz(aBall) * BallBouncePlayfieldHardFactor, aBall
End Sub

'/////////////////////////////  DELAYED DROP - TO PLAYFIELD - SOUND  ////////////////////////////
Sub RandomSoundDelayedBallDropOnPlayfield(aBall)
  Select Case Int(Rnd*5)+1
    Case 1 : PlaySoundAtLevelStatic ("Ball_Drop_Playfield_1_Delayed"), DelayedBallDropOnPlayfieldSoundLevel, aBall
    Case 2 : PlaySoundAtLevelStatic ("Ball_Drop_Playfield_2_Delayed"), DelayedBallDropOnPlayfieldSoundLevel, aBall
    Case 3 : PlaySoundAtLevelStatic ("Ball_Drop_Playfield_3_Delayed"), DelayedBallDropOnPlayfieldSoundLevel, aBall
    Case 4 : PlaySoundAtLevelStatic ("Ball_Drop_Playfield_4_Delayed"), DelayedBallDropOnPlayfieldSoundLevel, aBall
    Case 5 : PlaySoundAtLevelStatic ("Ball_Drop_Playfield_5_Delayed"), DelayedBallDropOnPlayfieldSoundLevel, aBall
  End Select
End Sub

'/////////////////////////////  BALL GATES AND BRACKET GATES SOUNDS  ////////////////////////////

Sub SoundPlayfieldGate()
  PlaySoundAtLevelStatic ("Gate_FastTrigger_" & Int(Rnd*2)+1), GateSoundLevel, Activeball
End Sub

Sub SoundHeavyGate()
  PlaySoundAtLevelStatic ("Gate_2"), GateSoundLevel, Activeball
End Sub

Sub Gates_hit(idx)
  SoundHeavyGate
End Sub

Sub GatesWire_hit(idx)
  SoundPlayfieldGate
End Sub

'/////////////////////////////  LEFT LANE ENTRANCE - SOUNDS  ////////////////////////////

Sub RandomSoundLeftArch()
  PlaySoundAtLevelActiveBall ("Arch_L" & Int(Rnd*4)+1), Vol(ActiveBall) * ArchSoundFactor
End Sub

Sub RandomSoundRightArch()
  PlaySoundAtLevelActiveBall ("Arch_R" & Int(Rnd*4)+1), Vol(ActiveBall) * ArchSoundFactor
End Sub


Sub Arch1_hit()
  If Activeball.velx > 1 Then SoundPlayfieldGate
  StopSound "Arch_L1"
  StopSound "Arch_L2"
  StopSound "Arch_L3"
  StopSound "Arch_L4"
End Sub

Sub Arch1_unhit()
  If activeball.velx < -8 Then
    RandomSoundRightArch
  End If
End Sub

Sub Arch2_hit()
  If Activeball.velx < 1 Then SoundPlayfieldGate
  StopSound "Arch_R1"
  StopSound "Arch_R2"
  StopSound "Arch_R3"
  StopSound "Arch_R4"
End Sub

Sub Arch2_unhit()
  If activeball.velx > 10 Then
    RandomSoundLeftArch
  End If
End Sub

'/////////////////////////////  SAUCERS (KICKER HOLES)  ////////////////////////////

Sub SoundSaucerLock()
  PlaySoundAtLevelStatic ("Saucer_Enter_" & Int(Rnd*2)+1), SaucerLockSoundLevel, Activeball
End Sub

Sub SoundSaucerKick(scenario, saucer)
  Select Case scenario
    Case 0: PlaySoundAtLevelStatic SoundFX("Saucer_Empty", DOFContactors), SaucerKickSoundLevel, saucer
    Case 1: PlaySoundAtLevelStatic SoundFX("Saucer_Kick", DOFContactors), SaucerKickSoundLevel, saucer
  End Select
End Sub

'/////////////////////////////  BALL COLLISION SOUND  ////////////////////////////
Sub OnBallBallCollision(ball1, ball2, velocity)
  Dim snd

  Select Case Int(Rnd*7)+1
    Case 1 : snd = "Ball_Collide_1"
    Case 2 : snd = "Ball_Collide_2"
    Case 3 : snd = "Ball_Collide_3"
    Case 4 : snd = "Ball_Collide_4"
    Case 5 : snd = "Ball_Collide_5"
    Case 6 : snd = "Ball_Collide_6"
    Case 7 : snd = "Ball_Collide_7"
  End Select


  If InRect(ball1.x,ball1.y,(Prim_MixMasterStickers_OFF.x-200),(Prim_MixMasterStickers_OFF.y-200),(Prim_MixMasterStickers_OFF.x+200),(Prim_MixMasterStickers_OFF.y-200),(Prim_MixMasterStickers_OFF.x+150),(Prim_MixMasterStickers_OFF.y+200),(Prim_MixMasterStickers_OFF.x-200),(Prim_MixMasterStickers_OFF.y+200)) Then
    snd = "Rubber_1"
  End If

  PlaySound (snd), 0, Csng(velocity) ^2 / 200 * BallWithBallCollisionSoundFactor * VolumeDial, AudioPan(ball1), 0, Pitch(ball1), 0, 0, AudioFade(ball1)

End Sub


'///////////////////////////  DROP TARGET HIT SOUNDS  ///////////////////////////

Sub RandomSoundDropTargetReset(obj)
  PlaySoundAtLevelStatic SoundFX("Drop_Target_Reset_" & Int(Rnd*6)+1,DOFContactors), 1, obj
End Sub

Sub SoundDropTargetDrop(obj)
  PlaySoundAtLevelStatic ("Drop_Target_Down_" & Int(Rnd*6)+1), 200, obj
End Sub

'/////////////////////////////  GI AND FLASHER RELAYS  ////////////////////////////

Const RelayFlashSoundLevel = 0.315                  'volume level; range [0, 1];
Const RelayGISoundLevel = 1.05                  'volume level; range [0, 1];

Sub Sound_GI_Relay(toggle, obj)
  Select Case toggle
    Case 1
      PlaySoundAtLevelStatic ("Relay_GI_On"), 0.025*RelayGISoundLevel, obj
    Case 0
      PlaySoundAtLevelStatic ("Relay_GI_Off"), 0.025*RelayGISoundLevel, obj
  End Select
End Sub

Sub Sound_Flash_Relay(toggle, obj)
  Select Case toggle
    Case 1
      PlaySoundAtLevelStatic ("Relay_Flash_On"), 0.025*RelayFlashSoundLevel, obj
    Case 0
      PlaySoundAtLevelStatic ("Relay_Flash_Off"), 0.025*RelayFlashSoundLevel, obj
  End Select
End Sub

'/////////////////////////////////////////////////////////////////
'         End Mechanical Sounds
'/////////////////////////////////////////////////////////////////

'******************************************************
'****  FLEEP MECHANICAL SOUNDS
'******************************************************



DIM VRThings
If VRRoom > 0 Then
  lockbar.visible = 0
  SetBackglass
  BGDude.visible = 1
  For Each VRthings in Desktop_digits: VRThings.Intensity = 0: Next
  For Each VRthings in VRBGGI: VRThings.visible = 1: Next
  displaytimer.enabled = false
  OverheadAmbiant.visible = 0
  OverheadAmbiantVR.visible = 1
  If VRRoom = 1 Then
    for each VRThings in VR_Room:VRThings.visible = 1:Next
    for each VRThings in VR_Min:VRThings.visible = 0:Next
    for each VRThings in VR_Pincab:VRThings.visible = 1:Next
    ClockTimer.enabled = 1 : BeerTimer.enabled = 1
  End If
  If VRRoom = 2 Then
    for each VRThings in VR_Pincab:VRThings.visible = 1:Next
    for each VRThings in VR_Min:VRThings.visible = 1:Next
    for each VRThings in VR_Room:VRThings.visible = 0:Next
    ClockTimer.enabled = 0 : BeerTimer.enabled = 0
  End If
  If VRRoom = 3 Then
    for each VRThings in VR_Pincab:VRThings.visible = 0:Next
    for each VRThings in VR_Min:VRThings.visible = 0:Next
    for each VRThings in VR_Room:VRThings.visible = 0:Next
    PinCab_Backbox.visible = 1
    Pincab_Grill.visible = 1
    ClockTimer.enabled = 0 : BeerTimer.enabled = 0
  End If
Else
  for each VRThings in VR_Pincab:VRThings.visible = 0:Next
  for each VRThings in VR_Min:VRThings.visible = 0:Next
  for each VRThings in VR_Room:VRThings.visible = 0:Next
End if

Sub SetBackglass()
  Dim obj

  For Each obj In VRBackglass
    obj.x = obj.x
    obj.height = - obj.y + 175
    obj.y = 25 'adjusts the distance from the backglass towards the user
    obj.rotx = -86.5
  Next

  For Each obj In VRBackglassDigits
    obj.x = obj.x
    obj.height = - obj.y + 175
    obj.y = 45 'adjusts the distance from the backglass towards the user
    obj.rotx = -87.5
  Next

  BGDude.height = - BGDude.y + 173
  BGDude.y = 7 'adjusts the distance from the backglass towards the user
  BGDude.rotx = -86.5

End Sub


Dim VRDigits(32)
VRDigits(0)=Array(ax00, ax05, ax0c, ax0d, ax08, ax01, ax06, ax0e, ax02, ax03, ax04, ax07, ax0b, ax0a, ax09, ax0f)
VRDigits(1)=Array(ax10, ax15, ax1c, ax1d, ax18, ax11, ax16, ax1e, ax12, ax13, ax14, ax17, ax1b, ax1a, ax19, ax1f)
VRDigits(2)=Array(ax20, ax25, ax2c, ax2d, ax28, ax21, ax26, ax2e, ax22, ax23, ax24, ax27, ax2b, ax2a, ax29, ax2f)
VRDigits(3)=Array(ax30, ax35, ax3c, ax3d, ax38, ax31, ax36, ax3e, ax32, ax33, ax34, ax37, ax3b, ax3a, ax39, ax3f)
VRDigits(4)=Array(ax40, ax45, ax4c, ax4d, ax48, ax41, ax46, ax4e, ax42, ax43, ax44, ax47, ax4b, ax4a, ax49, ax4f)
VRDigits(5)=Array(ax50, ax55, ax5c, ax5d, ax58, ax51, ax56, ax5e, ax52, ax53, ax54, ax57, ax5b, ax5a, ax59, ax5f)
VRDigits(6)=Array(ax60, ax65, ax6c, ax6d, ax68, ax61, ax66, ax6e, ax62, ax63, ax64, ax67, ax6b, ax6a, ax69, ax6f)
VRDigits(7)=Array(ax70, ax75, ax7c, ax7d, ax78, ax71, ax76, ax7e, ax72, ax73, ax74, ax77, ax7b, ax7a, ax79, ax7f)
VRDigits(8)=Array(ax80, ax85, ax8c, ax8d, ax88, ax81, ax86, ax8e, ax82, ax83, ax84, ax87, ax8b, ax8a, ax89, ax8f)
VRDigits(9)=Array(ax90, ax95, ax9c, ax9d, ax98, ax91, ax96, ax9e, ax92, ax93, ax94, ax97, ax9b, ax9a, ax99, ax9f)
VRDigits(10)=Array(axa0, axa5, axac, axad, axa8, axa1, axa6, axae, axa2, axa3, axa4, axa7, axab, axaa, axa9, axaf)
VRDigits(11)=Array(axb0, axb5, axbc, axbd, axb8, axb1, axb6, axbe, axb2, axb3, axb4, axb7, axbb, axba, axb9, axbf)
VRDigits(12)=Array(axc0, axc5, axcc, axcd, axc8, axc1, axc6, axce, axc2, axc3, axc4, axc7, axcb, axca, axc9, axcf)
VRDigits(13)=Array(axd0, axd5, axdc, axdd, axd8, axd1, axd6, axde, axd2, axd3, axd4, axd7, axdb, axda, axd9, axdf)
VRDigits(14)=Array(axe0, axe5, axec, axed, axe8, axe1, axe6, axee, axe2, axe3, axe4, axe7, axeb, axea, axe9, axef)
VRDigits(15)=Array(axf0, axf5, axfc, axfd, axf8, axf1, axf6, axfe, axf2, axf3, axf4, axf7, axfb, axfa, axf9, axff)

VRDigits(16)=Array(bx00, bx05, bx0c, bx0d, bx08, bx01, bx06, bx0e, bx02, bx03, bx04, bx07, bx0b, bx0a, bx09, bx0f)
VRDigits(17)=Array(bx10, bx15, bx1c, bx1d, bx18, bx11, bx16, bx1e, bx12, bx13, bx14, bx17, bx1b, bx1a, bx19, bx1f)
VRDigits(18)=Array(bx20, bx25, bx2c, bx2d, bx28, bx21, bx26, bx2e, bx22, bx23, bx24, bx27, bx2b, bx2a, bx29, bx2f)
VRDigits(19)=Array(bx30, bx35, bx3c, bx3d, bx38, bx31, bx36, bx3e, bx32, bx33, bx34, bx37, bx3b, bx3a, bx39, bx3f)
VRDigits(20)=Array(bx40, bx45, bx4c, bx4d, bx48, bx41, bx46, bx4e, bx42, bx43, bx44, bx47, bx4b, bx4a, bx49, bx4f)
VRDigits(21)=Array(bx50, bx55, bx5c, bx5d, bx58, bx51, bx56, bx5e, bx52, bx53, bx54, bx57, bx5b, bx5a, bx59, bx5f)
VRDigits(22)=Array(bx60, bx65, bx6c, bx6d, bx68, bx61, bx66, bx6e, bx62, bx63, bx64, bx67, bx6b, bx6a, bx69, bx6f)
VRDigits(23)=Array(bx70, bx75, bx7c, bx7d, bx78, bx71, bx76, bx7e, bx72, bx73, bx74, bx77, bx7b, bx7a, bx79, bx7f)
VRDigits(24)=Array(bx80, bx85, bx8c, bx8d, bx88, bx81, bx86, bx8e, bx82, bx83, bx84, bx87, bx8b, bx8a, bx89, bx8f)
VRDigits(25)=Array(bx90, bx95, bx9c, bx9d, bx98, bx91, bx96, bx9e, bx92, bx93, bx94, bx97, bx9b, bx9a, bx99, bx9f)
VRDigits(26)=Array(bxa0, bxa5, bxac, bxad, bxa8, bxa1, bxa6, bxae, bxa2, bxa3, bxa4, bxa7, bxab, bxaa, bxa9, bxaf)
VRDigits(27)=Array(bxb0, bxb5, bxbc, bxbd, bxb8, bxb1, bxb6, bxbe, bxb2, bxb3, bxb4, bxb7, bxbb, bxba, bxb9, bxbf)
VRDigits(28)=Array(bxc0, bxc5, bxcc, bxcd, bxc8, bxc1, bxc6, bxce, bxc2, bxc3, bxc4, bxc7, bxcb, bxca, bxc9, bxcf)
VRDigits(29)=Array(bxd0, bxd5, bxdc, bxdd, bxd8, bxd1, bxd6, bxde, bxd2, bxd3, bxd4, bxd7, bxdb, bxda, bxd9, bxdf)
VRDigits(30)=Array(bxe0, bxe5, bxec, bxed, bxe8, bxe1, bxe6, bxee, bxe2, bxe3, bxe4, bxe7, bxeb, bxea, bxe9, bxef)
VRDigits(31)=Array(bxf0, bxf5, bxfc, bxfd, bxf8, bxf1, bxf6, bxfe, bxf2, bxf3, bxf4, bxf7, bxfb, bxfa, bxf9, bxff)



Sub DisplayTimerVR
  Dim ChgLED, ii, num, chg, stat, obj
  ChgLED=Controller.ChangedLEDs(&H00000000, &Hffffffff)
  If Not IsEmpty(ChgLED) Then
    For ii=0 To UBound(chgLED)
      num=chgLED(ii,0)
      chg=chgLED(ii,1)
      stat=chgLED(ii,2)
      For Each obj In VRDigits(num)
        If chg And 1 Then obj.visible=stat And 1
        chg=chg\2
        stat=stat\2
      Next
    Next
     'end if
  End If
End Sub

Dim BGSeq

dim sol11lvl
sub Sol11(level)
  If VRRoom > 0 Then
    sol11lvl = 1
    sol11timer_timer
  End If
end sub

sub Sol11timer_timer
  dim obj
  if Not sol11timer.enabled then
    VRBGFL11_1.visible = 1
    VRBGFL11_2.visible = 1
    VRBGFL11_3.visible = 1
    VRBGFL11_4.visible = 1
    sol11timer.enabled = true
  end if
  sol11lvl = 0.85 * sol11lvl - 0.01
  if sol11lvl < 0 then sol11lvl = 0
    VRBGFL11_1.opacity = 150 * sol11lvl^2
    VRBGFL11_2.opacity = 150 * sol11lvl^2
    VRBGFL11_3.opacity = 150 * sol11lvl^2
    VRBGFL11_4.opacity = 150 * sol11lvl^3
  if sol11lvl =< 0 Then
    VRBGFL11_1.visible = 0
    VRBGFL11_2.visible = 0
    VRBGFL11_3.visible = 0
    VRBGFL11_4.visible = 0
    sol11timer.enabled = false
  end if
end sub


'*******************************************
'  VR Plunger Code
'*******************************************

Sub TimerVRPlunger_Timer
  If Pincab_Shooter.Y < 40 then
       Pincab_Shooter.Y = Pincab_Shooter.Y + 5
  End If
End Sub

Sub TimerVRPlunger2_Timer
  Pincab_Shooter.Y = -88.61317 + (5* Plunger.Position) -20
  timervrplunger2.enabled = 0
End Sub

' ***** Beer Bubble Code - Rawd *****
Sub BeerTimer_Timer()

  Randomize(21)
  BeerBubble1.z = BeerBubble1.z + Rnd(1)*0.5
  if BeerBubble1.z > -771 then BeerBubble1.z = -955
    BeerBubble2.z = BeerBubble2.z + Rnd(1)*1
  if BeerBubble2.z > -768 then BeerBubble2.z = -955
    BeerBubble3.z = BeerBubble3.z + Rnd(1)*1
  if BeerBubble3.z > -768 then BeerBubble3.z = -955
    BeerBubble4.z = BeerBubble4.z + Rnd(1)*0.75
  if BeerBubble4.z > -774 then BeerBubble4.z = -955
    BeerBubble5.z = BeerBubble5.z + Rnd(1)*1
  if BeerBubble5.z > -771 then BeerBubble5.z = -955
    BeerBubble6.z = BeerBubble6.z + Rnd(1)*1
  if BeerBubble6.z > -774 then BeerBubble6.z = -955
    BeerBubble7.z = BeerBubble7.z + Rnd(1)*0.8
  if BeerBubble7.z > -768 then BeerBubble7.z = -955
    BeerBubble8.z = BeerBubble8.z + Rnd(1)*1
  if BeerBubble8.z > -771 then BeerBubble8.z = -955
End Sub



' ***************** VR Clock code below - THANKS RASCAL ******************
Dim CurrentMinute ' for VR clock
' VR Clock code below....
Sub ClockTimer_Timer()
  Pminutes.RotAndTra2 = (Minute(Now())+(Second(Now())/100))*6
  Phours.RotAndTra2 = Hour(Now())*30+(Minute(Now())/2)
    Pseconds.RotAndTra2 = (Second(Now()))*6
  CurrentMinute=Minute(Now())
End Sub
 ' ********************** END CLOCK CODE   *********************************

'***********************************************************************
'* TABLE OPTIONS FOR IN-GAME MENU
'***********************************************************************

Dim LightLevel : LightLevel = 70            ' LightLevel - Value between 0 and 100 (0=Dark ... 100=Bright)
Dim VolumeDial : VolumeDial = 0.8     ' Overall Mechanical sound effect volume. Recommended values should be no greater than 1.
Dim BallRollVolume : BallRollVolume = 0.5   ' Level of ball rolling volume. Value between 0 and 1
Dim RampRollVolume : RampRollVolume = 0.5   ' Level of ramp rolling volume. Value between 0 and 1

' Options
Const Opt_Light = 0
Const Opt_Volume = 1
Const Opt_Volume_Ramp = 2
Const Opt_Volume_Ball = 3

' Informations
Const Opt_Info_1 = 4
Const Opt_Info_2 = 5

Const NOptions = 6

Const FlexDMD_RenderMode_DMD_GRAY_2 = 0
Const FlexDMD_RenderMode_DMD_GRAY_4 = 1
Const FlexDMD_RenderMode_DMD_RGB = 2
Const FlexDMD_RenderMode_SEG_2x16Alpha = 3
Const FlexDMD_RenderMode_SEG_2x20Alpha = 4
Const FlexDMD_RenderMode_SEG_2x7Alpha_2x7Num = 5
Const FlexDMD_RenderMode_SEG_2x7Alpha_2x7Num_4x1Num = 6
Const FlexDMD_RenderMode_SEG_2x7Num_2x7Num_4x1Num = 7
Const FlexDMD_RenderMode_SEG_2x7Num_2x7Num_10x1Num = 8
Const FlexDMD_RenderMode_SEG_2x7Num_2x7Num_4x1Num_gen7 = 9
Const FlexDMD_RenderMode_SEG_2x7Num10_2x7Num10_4x1Num = 10
Const FlexDMD_RenderMode_SEG_2x6Num_2x6Num_4x1Num = 11
Const FlexDMD_RenderMode_SEG_2x6Num10_2x6Num10_4x1Num = 12
Const FlexDMD_RenderMode_SEG_4x7Num10 = 13
Const FlexDMD_RenderMode_SEG_6x4Num_4x1Num = 14
Const FlexDMD_RenderMode_SEG_2x7Num_4x1Num_1x16Alpha = 15
Const FlexDMD_RenderMode_SEG_1x16Alpha_1x16Num_1x7Num = 16

Const FlexDMD_Align_TopLeft = 0
Const FlexDMD_Align_Top = 1
Const FlexDMD_Align_TopRight = 2
Const FlexDMD_Align_Left = 3
Const FlexDMD_Align_Center = 4
Const FlexDMD_Align_Right = 5
Const FlexDMD_Align_BottomLeft = 6
Const FlexDMD_Align_Bottom = 7
Const FlexDMD_Align_BottomRight = 8

Const FlexDMD_Scaling_Fit = 0
Const FlexDMD_Scaling_Fill = 1
Const FlexDMD_Scaling_FillX = 2
Const FlexDMD_Scaling_FillY = 3
Const FlexDMD_Scaling_Stretch = 4
Const FlexDMD_Scaling_StretchX = 5
Const FlexDMD_Scaling_StretchY = 6
Const FlexDMD_Scaling_None = 7

Const FlexDMD_Interpolation_Linear = 0
Const FlexDMD_Interpolation_ElasticIn = 1
Const FlexDMD_Interpolation_ElasticOut = 2
Const FlexDMD_Interpolation_ElasticInOut = 3
Const FlexDMD_Interpolation_QuadIn = 4
Const FlexDMD_Interpolation_QuadOut = 5
Const FlexDMD_Interpolation_QuadInOut = 6
Const FlexDMD_Interpolation_CubeIn = 7
Const FlexDMD_Interpolation_CubeOut = 8
Const FlexDMD_Interpolation_CubeInOut = 9
Const FlexDMD_Interpolation_QuartIn = 10
Const FlexDMD_Interpolation_QuartOut = 11
Const FlexDMD_Interpolation_QuartInOut = 12
Const FlexDMD_Interpolation_QuintIn = 13
Const FlexDMD_Interpolation_QuintOut = 14
Const FlexDMD_Interpolation_QuintInOut = 15
Const FlexDMD_Interpolation_SineIn = 16
Const FlexDMD_Interpolation_SineOut = 17
Const FlexDMD_Interpolation_SineInOut = 18
Const FlexDMD_Interpolation_BounceIn = 19
Const FlexDMD_Interpolation_BounceOut = 20
Const FlexDMD_Interpolation_BounceInOut = 21
Const FlexDMD_Interpolation_CircIn = 22
Const FlexDMD_Interpolation_CircOut = 23
Const FlexDMD_Interpolation_CircInOut = 24
Const FlexDMD_Interpolation_ExpoIn = 25
Const FlexDMD_Interpolation_ExpoOut = 26
Const FlexDMD_Interpolation_ExpoInOut = 27
Const FlexDMD_Interpolation_BackIn = 28
Const FlexDMD_Interpolation_BackOut = 29
Const FlexDMD_Interpolation_BackInOut = 30

Dim OptionDMD: Set OptionDMD = Nothing
Dim bOptionsMagna, bInOptions : bOptionsMagna = False
Dim OptPos, OptSelected, OptN, OptTop, OptBot, OptSel
Dim OptFontHi, OptFontLo

Sub Options_Open
  bOptionsMagna = False
  On Error Resume Next
  Set OptionDMD = CreateObject("FlexDMD.FlexDMD")
  On Error Goto 0
  If OptionDMD is Nothing Then
    Debug.Print "FlexDMD is not installed"
    Debug.Print "Option UI can not be opened"
    MsgBox "You need to install FlexDMD to access table options"
    Exit Sub
  End If
  If Table1.ShowDT Then OptionDMDFlasher.RotX = -(Table1.Inclination + Table1.Layback)
  Debug.Print "Creating Flex"
  bInOptions = True
  OptPos = 0
  OptSelected = False
  OptionDMD.Show = False
  OptionDMD.RenderMode = FlexDMD_RenderMode_DMD_GRAY_4
  OptionDMD.Width = 128
  OptionDMD.Height = 32
  OptionDMD.Clear = True
  OptionDMD.Run = True
  Dim a, scene, font
  Set scene = OptionDMD.NewGroup("Scene")
  Set OptFontHi = OptionDMD.NewFont("FlexDMD.Resources.teeny_tiny_pixls-5.fnt", vbWhite, vbWhite, 0)
  Set OptFontLo = OptionDMD.NewFont("FlexDMD.Resources.teeny_tiny_pixls-5.fnt", RGB(100, 100, 100), RGB(100, 100, 100), 0)
  Set OptSel = OptionDMD.NewGroup("Sel")
  Set a = OptionDMD.NewLabel(">", OptFontLo, ">>>")
  a.SetAlignedPosition 1, 16, FlexDMD_Align_Left
  OptSel.AddActor a
  Set a = OptionDMD.NewLabel(">", OptFontLo, "<<<")
  a.SetAlignedPosition 127, 16, FlexDMD_Align_Right
  OptSel.AddActor a
  scene.AddActor OptSel
  OptSel.SetBounds 0, 0, 128, 32
  OptSel.Visible = False

  Set a = OptionDMD.NewLabel("Info1", OptFontLo, "MAGNA EXIT/ENTER")
  a.SetAlignedPosition 1, 32, FlexDMD_Align_BottomLeft
  scene.AddActor a
  Set a = OptionDMD.NewLabel("Info2", OptFontLo, "FLIPPER SELECT")
  a.SetAlignedPosition 127, 32, FlexDMD_Align_BottomRight
  scene.AddActor a
  Set OptN = OptionDMD.NewLabel("Pos", OptFontLo, "LINE 1")
  Set OptTop = OptionDMD.NewLabel("Top", OptFontLo, "LINE 1")
  Set OptBot = OptionDMD.NewLabel("Bottom", OptFontLo, "LINE 2")
  scene.AddActor OptN
  scene.AddActor OptTop
  scene.AddActor OptBot
  Options_OnOptChg
  OptionDMD.LockRenderThread
  OptionDMD.Stage.AddActor scene
  OptionDMD.UnlockRenderThread
  OptionDMDFlasher.Visible = True
  Debug.Print "Got Thru Function"
End Sub

Sub Options_UpdateDMD
  If OptionDMD is Nothing Then Exit Sub
  Dim DMDp: DMDp = OptionDMD.DmdPixels
  If Not IsEmpty(DMDp) Then
    DMDWidth = OptionDMD.Width
    DMDHeight = OptionDMD.Height
    DMDPixels = DMDp
  End If
End Sub

Sub Options_Close
  bInOptions = False
  OptionDMDFlasher.Visible = False
  If OptionDMD is Nothing Then Exit Sub
  OptionDMD.Run = False
  Set OptionDMD = Nothing
End Sub

Function Options_OnOffText(opt)
  If opt Then
    Options_OnOffText = "ON"
  Else
    Options_OnOffText = "OFF"
  End If
End Function

Sub Options_OnOptChg
  If OptionDMD is Nothing Then Exit Sub
  OptionDMD.LockRenderThread
  OptN.Text = (OptPos+1) & "/" & NOptions
  If OptSelected Then
    OptTop.Font = OptFontLo
    OptBot.Font = OptFontHi
    OptSel.Visible = True
  Else
    OptTop.Font = OptFontHi
    OptBot.Font = OptFontLo
    OptSel.Visible = False
  End If

  If OptPos = Opt_Light Then
    OptTop.Text = "LIGHT LEVEL"
    OptBot.Text = "LEVEL " & LightLevel
    SaveValue cGameName, "LIGHT", LightLevel
  ElseIf OptPos = Opt_Volume Then
    OptTop.Text = "MECH VOLUME"
    OptBot.Text = "LEVEL " & CInt(VolumeDial * 100)
    SaveValue cGameName, "VOLUME", VolumeDial
  ElseIf OptPos = Opt_Volume_Ramp Then
    OptTop.Text = "RAMP VOLUME"
    OptBot.Text = "LEVEL " & CInt(RampRollVolume * 100)
    SaveValue cGameName, "RAMPVOLUME", RampRollVolume
  ElseIf OptPos = Opt_Volume_Ball Then
    OptTop.Text = "BALL VOLUME"
    OptBot.Text = "LEVEL " & CInt(BallRollVolume * 100)
    SaveValue cGameName, "BALLVOLUME", BallRollVolume
' ElseIf OptPos = Opt_ModSol Then
'   OptTop.Text = "MOD. FLASHERS (RESTART)"
'   If ModSol = 1 Then OptBot.Text = "SIMPLE"
'   If ModSol = 2 Then OptBot.Text = "PWM"
'   SaveValue cGameName, "MODSOL", ModSol
' ElseIf OptPos = Opt_VR_Room Then
'   OptTop.Text = "VR ROOM"
'   If VRRoomChoice = 1 Then OptBot.Text = "MINIMAL"
'   If VRRoomChoice = 2 Then OptBot.Text = "ULTRA MIN"
'   SaveValue cGameName, "VRROOM", VRRoomChoice
  ElseIf OptPos = Opt_Info_1 Then
    OptTop.Text = "VPX " & VersionMajor & "." & VersionMinor & "." & VersionRevision
    OptBot.Text = "DR. DUDE " & TableVersion
  ElseIf OptPos = Opt_Info_2 Then
    OptTop.Text = "RENDER MODE"
    If RenderingMode = 0 Then OptBot.Text = "DEFAULT"
    If RenderingMode = 1 Then OptBot.Text = "STEREO 3D"
    If RenderingMode = 2 Then OptBot.Text = "VR"
  End If
  OptTop.Pack
  OptTop.SetAlignedPosition 127, 1, FlexDMD_Align_TopRight
  OptBot.SetAlignedPosition 64, 16, FlexDMD_Align_Center
  OptionDMD.UnlockRenderThread
  UpdateMods
End Sub


Sub Options_Toggle(amount)
  If OptionDMD is Nothing Then Exit Sub
  If OptPos = Opt_Light Then
    LightLevel = LightLevel + amount * 10
    If LightLevel < 0 Then LightLevel = 100
    If LightLevel > 100 Then LightLevel = 0
  ElseIf OptPos = Opt_Volume Then
    VolumeDial = VolumeDial + amount * 0.1
    If VolumeDial < 0 Then VolumeDial = 1
    If VolumeDial > 1 Then VolumeDial = 0
  ElseIf OptPos = Opt_Volume_Ramp Then
    RampRollVolume = RampRollVolume + amount * 0.1
    If RampRollVolume < 0 Then RampRollVolume = 1
    If RampRollVolume > 1 Then RampRollVolume = 0
  ElseIf OptPos = Opt_Volume_Ball Then
    BallRollVolume = BallRollVolume + amount * 0.1
    If BallRollVolume < 0 Then BallRollVolume = 1
    If BallRollVolume > 1 Then BallRollVolume = 0
  ElseIf OptPos = Opt_ModSol Then
    ModSol = ModSol + amount
    If ModSol < 1 Then ModSol = 2
    If ModSol > 2 Then ModSol = 1
  ElseIf OptPos = Opt_VR_Room Then
    VRRoomChoice = VRRoomChoice + amount
    If VRRoomChoice < 1 Then VRRoomChoice = 2
    If VRRoomChoice > 2 Then VRRoomChoice = 1
  End If
End Sub


Sub Options_KeyDown(ByVal keycode)
  If OptSelected Then
    If keycode = LeftMagnaSave Then ' Exit / Cancel
      OptSelected = False
    ElseIf keycode = RightMagnaSave Then ' Enter / Select
      OptSelected = False
    ElseIf keycode = LeftFlipperKey Then ' Next / +
      Options_Toggle  -1
    ElseIf keycode = RightFlipperKey Then ' Prev / -
      Options_Toggle  1
    End If
  Else
    If keycode = LeftMagnaSave Then ' Exit / Cancel
      Options_Close
    ElseIf keycode = RightMagnaSave Then ' Enter / Select
      If OptPos < Opt_Info_1 Then OptSelected = True
    ElseIf keycode = LeftFlipperKey Then ' Next / +
      OptPos = OptPos - 1
      If OptPos < 0 Then OptPos = NOptions - 1
    ElseIf keycode = RightFlipperKey Then ' Prev / -
      OptPos = OptPos + 1
      If OptPos >= NOPtions Then OptPos = 0
    End If
  End If
  Options_OnOptChg
End Sub


Sub Options_Load
  Dim x
    x = LoadValue(cGameName, "LIGHT") : If x <> "" Then LightLevel = CInt(x) Else LightLevel = 70
    x = LoadValue(cGameName, "VOLUME") : If x <> "" Then VolumeDial = CNCDbl(x) Else VolumeDial = 0.8
    x = LoadValue(cGameName, "RAMPVOLUME") : If x <> "" Then RampRollVolume = CNCDbl(x) Else RampRollVolume = 0.5
    x = LoadValue(cGameName, "BALLVOLUME") : If x <> "" Then BallRollVolume = CNCDbl(x) Else BallRollVolume = 0.5
  UpdateMods
End Sub



Sub UpdateMods

  ' Light Level
    Select Case LightLevel
    Case 10:table1.ColorGradeImage = "LUT14"
    Case 20:table1.ColorGradeImage = "LUT12"
    Case 30:table1.ColorGradeImage = "LUT10"
    Case 40:table1.ColorGradeImage = "LUT9"
    Case 50:table1.ColorGradeImage = "LUT8"
    Case 60:table1.ColorGradeImage = "LUT7"
    Case 70:table1.ColorGradeImage = "LUT5"
    Case 80:table1.ColorGradeImage = "LUT3"
    Case 90:table1.ColorGradeImage = "LUT1"
    Case 100:table1.ColorGradeImage = "LUT0"
  End Select


End Sub


' Culture neutral string to double conversion (handles situation where you don't know how the string was written)
Function CNCDbl(str)
    Dim strt, Sep, i
    If IsNumeric(str) Then
        CNCDbl = CDbl(str)
    Else
        Sep = Mid(CStr(0.5), 2, 1)
        Select Case Sep
        Case "."
            i = InStr(1, str, ",")
        Case ","
            i = InStr(1, str, ".")
        End Select
        If i = 0 Then
            CNCDbl = Empty
        Else
            strt = Mid(str, 1, i - 1) & Sep & Mid(str, i + 1)
            If IsNumeric(strt) Then
                CNCDbl = CDbl(strt)
            Else
                CNCDbl = Empty
            End If
        End If
    End If
End Function




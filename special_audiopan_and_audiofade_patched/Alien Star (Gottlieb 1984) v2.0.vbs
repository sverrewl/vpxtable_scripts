' Alien Star / IPD No. 49 / June, 1984 / 4 Players
'
'  ___   _     _____ _____ _   _   _____ _____ ___  ______
' / _ \ | |   |_   _|  ___| \ | | /  ___|_   _/ _ \ | ___ \
'/ /_\ \| |     | | | |__ |  \| | \ `--.  | |/ /_\ \| |_/ /
'|  _  || |     | | |  __|| . ` |  `--. \ | ||  _  ||    /
'| | | || |_____| |_| |___| |\  | /\__/ / | || | | || |\ \
'\_| |_/\_____/\___/\____/\_| \_/ \____/  \_/\_| |_/\_| \_|
'
'

'********************************************************************************************************************************
'******** Work done by UnclePaulie on Hybrid version 0.01 - 2.0 *********
'     Added latest VPW standards, physics, GI baking, Lampz and 3D inserts, shadows, updated backglass for desktop,
'   roth standup targets, standalone compatibility, sling corrections,
'   optional DIP adjustments, updated flipper physics, VR and fully hybrid, Flupper style bumpers and lighting.
'   Full details at the bottom of the script.  Redbone did the playfield, plastics, and other graphics.
'   Original scans, image assets, cab artwork stencils all used with permission from Borgdog.
'********************************************************************************************************************************


Option Explicit
Randomize


'***********************************************************************
'* TABLE OPTIONS *******************************************************
'***********************************************************************

' ALL OPTIONS ARE IN THE F12 MENU, EXCEPT FOR THE BALL COLOR and SetDIPSwitches BELOW

const ChangeBall  = 1   '0 = Ball changes are done via the drop down table menu
              '1 = ball selection per the below BallLightness option (default)
const BallLightness = 2   '0 = dark, 1 = not as dark, 2 = bright (default), 3 = brightest

const SetDIPSwitches= 0 ' If you want to set the dips differently, set to 1, and then hit F6 to launch.

        ' These are the default DIPS that are and set at original table launch:
          '1=Coin Chute same
          '0=3rd coin chute no effect
          '0=Game mode of replay for hitting threshold
          '1=Number of balls = 3
          '1=Playfield Special of extra ball  (Tourney mode this is zero)
          '0=Novelty of normal game mode    (Tourney mode this set to 1)
          '1=Match feature
          '1=Background Sound


' VR USERS... THE F12 MENU ISN'T AVAILABLE AND WILL NEED TO SELECT OPTIONS HERE


' User Options for In-Game Menu F12 then use Flipper and Magna Keys
'   NOTE: The detailed definitions of F12 Table Options are at the bottom of the script.

  'variables
  Dim VolumeDial, BallRollVolume, DynamicBallShadowsOn, AmbientBallShadowOn
  Dim ColorLUT, LutToggleSound, LutToggleSoundLevel
  Dim ROMVolume
  Dim SideRailGotlieb, LaneTentColor, FlipperColor, ScrewCapColor, SideWallColor, PlasticColor, StarReflections
  Dim Tourney
  Dim WallClock, poster, poster2, CustomWalls

'User Defined acrylic edge colors (This is if user defined option is selected for plastic edge color... Choose your own RGB)

  Dim UserAcrylColor

  UserAcrylColor = RGB(255,255,240) 'default is RGB(255,255,240)


Dim VR_Room, cab_mode, DesktopMode: DesktopMode = Table1.ShowDT
If RenderingMode = 2 Then VR_Room=1 Else VR_Room=0
if Not DesktopMode and VR_Room=0 Then cab_mode=1 Else cab_mode=0


' NOTE:  Options for desktop and cab are done via F12 menu.  VR users have to manually select here in the script.


IF VR_Room = 1 Then

  SideRailGotlieb = 1 '0 = standard metal side rails;  1 = gottlieb style bumpy side rails (desktop only) (default)
  LaneTentColor = 1 '0 = standard red; 1 = yellow (default)
  FlipperColor  = 3 '0 = standard red; 1 = yellow; 2 = black; 3 = orange (default)
  ScrewCapColor = 0 '0 = White (default), 1 = Black, 2 = Bronze, 3 = Silver, 4 = Blue/green, 5 = Gold
  SideWallColor = 0 '0 = Wood walls, 1 = Black Walls
  StarReflections = 1 '0 = No STAR reflections; 1 = STAR Reflections on Backwall
  PlasticColor  = 0 '0 = Off (default), 1 = Orange/Red, 2 = Yellow, 3 = Blue, 4 = User Defined, 5 = Rainbow, 6 = Random Option
  Tourney     = 0 '0 = Off (default), 1 = Presets for tourney mode:  includes: Harder DIPS settins (no extra ball)
              '  Tourney mode colors: Black flippers, yellow tents, and blue/green screw caps
    ' NOTE on Tourney Mode: This mode sets dip switches. You need to launch game, exit, and launch again for the ROM settings to take affect.
  WallClock   = 1 '1 Shows the clock in the VR minimal rooms only
  poster      = 1 '1 Shows the flyer posters in the VR room only
  poster2     = 1 '1 Shows the flyer posters in the VR room only
  CustomWalls   = 2 '0=UP's Original Minimal Walls, floor, and roof
                '1=Sixtoe's arcade style
                '2=DarthVito's Updated home walls with lights
                '3=DarthVito's plaster home walls
                '4=DarthVito's blue home walls
  ROMVolume   = 0 'ROM Volume (-32 is off; 0 is max db) - Requires Table Restart

  VolumeDial    = 0.8 'Overall Mechanical sound effect volume. Recommended values should be no greater than 1.
  BallRollVolume  = 0.5 'Level of ball rolling volume. Value between 0 and 1
  DynamicBallShadowsOn  = 1 '0 = no dynamic ball shadow ("triangles" near slings and such), 1 = enable dynamic ball shadow
  AmbientBallShadowOn   = 1 '0 = Static shadow under ball ("flasher" image, like JP's), 1 = Moving ball shadow ("primitive" object, like ninuzzu's) - This is the only one that behaves like a true shadow!, 2 = flasher image shadow, but it moves like ninuzzu's

  ColorLUT = 16

    if ColorLUT = 0 Then Table1.ColorGradeImage = "LUT0"
    if ColorLUT = 1 Then Table1.ColorGradeImage = "LUT1"
    if ColorLUT = 2 Then Table1.ColorGradeImage = "LUT2"
    if ColorLUT = 3 Then Table1.ColorGradeImage = "LUT3"
    if ColorLUT = 4 Then Table1.ColorGradeImage = "LUT4"
    if ColorLUT = 5 Then Table1.ColorGradeImage = "LUT5"
    if ColorLUT = 6 Then Table1.ColorGradeImage = "LUT6"
    if ColorLUT = 7 Then Table1.ColorGradeImage = "LUT7"
    if ColorLUT = 8 Then Table1.ColorGradeImage = "LUT8"
    if ColorLUT = 9 Then Table1.ColorGradeImage = "LUT9"
    if ColorLUT = 10 Then Table1.ColorGradeImage = "LUT10"
    if ColorLUT = 11 Then Table1.ColorGradeImage = "LUT11"
    if ColorLUT = 12 Then Table1.ColorGradeImage = "LUT12"
    if ColorLUT = 13 Then Table1.ColorGradeImage = "LUT13"
    if ColorLUT = 14 Then Table1.ColorGradeImage = "LUT14"
    if ColorLUT = 15 Then Table1.ColorGradeImage = "LUT15"
    if ColorLUT = 16 Then Table1.ColorGradeImage = "LUT16"

'   "0 = Fleep Natural Dark 1", "1 = Fleep Natural Dark 2", "2 = Fleep Warm Dark", "3 = Fleep Warm Bright", "4 = Fleep Warm Vivid Soft",
'   "5 = Fleep Warm Vivid Hard", "6 = Skitso Natural and Balanced", "7 = Skitso Natural High Contrast", "8 = 3rdaxis Referenced THX Standard",
'   "9 = CalleV Punchy Brightness and Contrast", "10 = HauntFreaks Desaturated", "11 = Tomate Washed Out","12 = VPW Original 1 to 1",
'   "13 = Bassgeige","14 = Blacklight","15 = B&W Comic Book","16 = Skitso New Warmer LUT"

End If


' Table1.Option arguments are:
' - option name, minimum value, maximum value, step between valid values, default value, unit (0=None, 1=Percent), an optional arry of literal strings

' Called when options are tweaked by the player.
' - 0: game has started, good time to load options and adjust accordingly
' - 1: an option has changed
' - 2: options have been reseted
' - 3: player closed the tweak UI, good time to update staticly prerendered parts


Sub Table1_OptionEvent(ByVal eventId)

  if VR_Room = 0 Then ' only allow F12 menu for desktop and cab.... not currently working in VR

  SideRailGotlieb = Table1.Option("Use Gotlieb SideRails", 0, 1, 1, 1, 0, Array("No", "Yes"))
  LaneTentColor = Table1.Option("Lane Tent Color", 0, 1, 1, 1, 0, Array("Red", "Yellow"))
  FlipperColor = Table1.Option("Flipper Color", 0, 3, 1, 3, 0, Array("Red", "Yellow", "Black", "Orange"))
  ScrewCapColor = Table1.Option("Screw Cap Color", 0, 5, 1, 0, 0, Array("White", "Black", "Bronze", "Silver", "Blue/Green", "Gold"))
  SideWallColor = Table1.Option("Side Wall Color", 0, 1, 1, 0, 0, Array("Wood", "Black"))
  StarReflections = Table1.Option("STAR Reflections on Backwall", 0, 1, 1, 1, 0, Array("No", "Yes"))

  PlasticColor = Table1.Option("Plastic Edge Color", 0, 6, 1, 0, 0, Array("Off", "Orange/Red", "Yellow", "Blue", "User Defined in Script", "Rainbow", "Random"))

  Tourney = Table1.Option("Tourney Mode - Requires Table Restart", 0, 1, 1, 0, 0, Array("Off", "On"))

  SetDefaultDips

' WallClock = Table1.Option("VR: Show Clock", 0, 1, 1, 1, 0, Array("No", "Yes"))
' poster = Table1.Option("VR: Show Poster 1", 0, 1, 1, 1, 0, Array("No", "Yes"))
' poster2 = Table1.Option("VR: Show Poster 2", 0, 1, 1, 1, 0, Array("No", "Yes"))
' CustomWalls = Table1.Option("VR: Wall Environment Selection", 0, 4, 1, 2, 0, Array("UP Original", "Arcade", "Updated Home", "Plaster", "Blue"))

  ROMVolume = Table1.Option("ROM Volume (-32 is off; 0 is max db) - Requires Table Restart", -32, 0, 1, 0, 0)

  SetupOptions

  VolumeDial = Table1.Option("Mech Volume", 0, 1, 0.01, 0.8, 1)
  BallRollVolume = Table1.Option("Ball Roll Volume", 0, 1, 0.01, 0.5, 1)
  DynamicBallShadowsOn = Table1.Option("Dynamic Ball Shadows", 0, 1, 1, 1, 0, Array("Off", "On"))
  AmbientBallShadowOn = Table1.Option("Ambient Ball Shadow", 0, 2, 1, 1, 0, Array("Off", "Moving Ball Shadow", "Flasher Image Shadow"))
' SetDIPSwitches = Table1.Option("Allow F6 to change DIP switches - Requires Table Restart", 0, 1, 1, 0, 0, Array("No", "Yes"))

   ' Color Saturation
  ColorLUT = Table1.Option("Color Saturation", 0, 16, 1, 16, 0, _
    Array("0 = Fleep Natural Dark 1", "1 = Fleep Natural Dark 2", "2 = Fleep Warm Dark", "3 = Fleep Warm Bright", "4 = Fleep Warm Vivid Soft", "5 = Fleep Warm Vivid Hard", _
    "6 = Skitso Natural and Balanced", "7 = Skitso Natural High Contrast", "8 = 3rdaxis Referenced THX Standard", "9 = CalleV Punchy Brightness and Contrast", "10 = HauntFreaks Desaturated", _
    "11 = Tomate Washed Out","12 = VPW Original 1 to 1","13 = Bassgeige","14 = Blacklight","15 = B&W Comic Book","16 = Skitso New Warmer LUT"))

    if ColorLUT = 0 Then Table1.ColorGradeImage = "LUT0"
    if ColorLUT = 1 Then Table1.ColorGradeImage = "LUT1"
    if ColorLUT = 2 Then Table1.ColorGradeImage = "LUT2"
    if ColorLUT = 3 Then Table1.ColorGradeImage = "LUT3"
    if ColorLUT = 4 Then Table1.ColorGradeImage = "LUT4"
    if ColorLUT = 5 Then Table1.ColorGradeImage = "LUT5"
    if ColorLUT = 6 Then Table1.ColorGradeImage = "LUT6"
    if ColorLUT = 7 Then Table1.ColorGradeImage = "LUT7"
    if ColorLUT = 8 Then Table1.ColorGradeImage = "LUT8"
    if ColorLUT = 9 Then Table1.ColorGradeImage = "LUT9"
    if ColorLUT = 10 Then Table1.ColorGradeImage = "LUT10"
    if ColorLUT = 11 Then Table1.ColorGradeImage = "LUT11"
    if ColorLUT = 12 Then Table1.ColorGradeImage = "LUT12"
    if ColorLUT = 13 Then Table1.ColorGradeImage = "LUT13"
    if ColorLUT = 14 Then Table1.ColorGradeImage = "LUT14"
    if ColorLUT = 15 Then Table1.ColorGradeImage = "LUT15"
    if ColorLUT = 16 Then Table1.ColorGradeImage = "LUT16"

  End If

End Sub



'********************
'Standard definitions
'********************

Const UseSolenoids  = 2
Const UseLamps      = 0
Const UseGI     = 0
Const UsingROM    = True  'The UsingROM flag is to indicate code that requires ROM usage. Mostly for instructional purposes only.
Const BallSize    = 50
Const BallMass    = 1
Const UseSync   = 0
Const cGameName   = "alienstr"


'*******************************************
'  Constants and Global Variables
'*******************************************

Const tnob = 2            ' total number of balls : 2 (trough)
Const lob = 0           ' Locked balls

Dim tablewidth: tablewidth = Table1.width
Dim tableheight: tableheight = Table1.height
Dim i, aBall1, aBall2, gBOT
Dim BIPL : BIPL = False       ' Ball in plunger lane
Dim gion: gion = 0
dim PlasProt, Ttents, Tscrewcap, TFlipper, WallColor


On Error Resume Next
ExecuteGlobal GetTextFile("controller.vbs")
If Err Then MsgBox "You need the controller.vbs in order to run this table, available in the vp10 package"
On Error Goto 0

LoadVPM "01210000", "sys80.VBS", 3.1


'******************************************************
'         Solenoids Map
'******************************************************

SolCallback(1)  = "SolSaucer"   'Saucer kick
SolCallback(8)  = "SolKnocker"    'Knocker
SolCallback(9)  = "SolTrough"   'Outhole kick to plunger lane
SolCallback(10) = "SolGameOver"   'GameOver Sub
SolCallback(11) = "SolGI"     'GI off

SolCallback(sLRFlipper) = "SolRFlipper"
SolCallback(sLLFlipper) = "SolLFlipper"



'*******************************************
'  Timers
'*******************************************

' The game timer interval is 10 ms
Sub GameTimer_Timer()
  Cor.Update            'update ball tracking
  RollingUpdate         'update rolling sounds
  DoSTAnim            'handle stand up target animations
  SpinnerTimer
End Sub


' The frame timer interval is -1, so executes at the display frame rate
Sub FrameTimer_Timer()
  FlipperVisualUpdate       'update flipper shadows and primitives
  If DynamicBallShadowsOn Or AmbientBallShadowOn Then DynamicBSUpdate 'update ball shadows

  If VR_Room=0 Then
    DisplayTimer
  End If

  If VR_Room=1 Then
    VRDisplayTimer
  End If

  LampTimer           'used for Lampz.  No need for separate timer.
  UpdateBallBrightness      'GI for the ball
End Sub


'******************************************************
'         Initialize the table
'******************************************************

Dim NTM

Sub Table1_Init
  vpmInit Me
    With Controller
    .GameName = cGameName
    If Err Then MsgBox "Can't start Game " & cGameName & vbNewLine & Err.Description:Exit Sub
    .SplashInfoLine = "Alien Star (Gottlieb 1984)"&chr(13)&"by UnclePaulie"
    .HandleKeyboard = 0
    .ShowTitle = 0
    .ShowDMDOnly = 1
    .ShowFrame = 0
    .HandleMechanics = 0
    .Hidden = 1
    .Games(cGameName).Settings.Value("sound") = 1 ' Set sound (0=OFF, 1=ON)
    On Error Resume Next
    Controller.SolMask(0) = 0
    vpmTimer.AddTimer 2000, "Controller.SolMask(0)=&Hffffffff'" 'ignore all solenoids - Then add the timer to renable all the solenoids after 2 seconds
    .Run GetPlayerHWnd
    If Err Then MsgBox Err.Description
    On Error Goto 0
  End With


' Main Timer init
  PinMAMETimer.Interval = PinMAMEInterval
  PinMAMETimer.Enabled = True

' Nudging
  vpmNudge.TiltSwitch = 57
  vpmNudge.Sensitivity = 3

  vpmNudge.TiltObj = Array(Bumper1,Bumper2,Bumper3,LeftSlingshot,RightSlingshot,LeftFlipper,RightFlipper)

' Ball initializations need for physical trough and ball shadows

  Set aBall1 = BallRelease.CreateSizedballWithMass(Ballsize/2,Ballmass)
  Set aBall2 = Slot1.CreateSizedballWithMass(Ballsize/2,Ballmass)

  gBOT = Array(aBall1, aBall2)
  controller.Switch(14) = 1
  controller.Switch(44) = 1

'Saucer initializations
  Controller.Switch(4) = 0

Dim xx

' Add balls to shadow dictionary
  For Each xx in gBOT
    bsDict.Add xx.ID, bsNone
  Next

  NTM=0    'turn off number to match box for desktop backglass
  FindDips  'find if match enabled, if so turn back on number to match box

' GI on at game start with slight delay
  vpmTimer.AddTimer 2500,"PFGI 1'"
  FlasherGI.opacity = 0
  FlBumperFadeTarget(1) = 0.3
  FlBumperFadeTarget(2) = 0.3
  FlBumperFadeTarget(3) = 0.3

' If user wants to set dip switches themselves it will force them to set it via F6.
  If SetDIPSwitches = 0 Then
    SetDefaultDips
  End If

If vr_room = 1 Then ' F12 menu only working for desktop and cab.  VR mode requires the setup be done here.
  SetupOptions ' enable all options above
End if


End Sub


Sub Table1_exit()
  Controller.Pause = False
  Controller.Stop
End Sub

Sub Table1_Paused
  Controller.Pause = True
End Sub

Sub Table1_unPaused
  Controller.Pause = False
End Sub



'******************************************************
'             Keys
'******************************************************

Sub Table1_KeyDown(ByVal keycode)

  If KeyCode = PlungerKey Then
    Plunger.PullBack
    SoundPlungerPull()
    VR_Plunger.Enabled = True
    VR_Plunger2.Enabled = False
    VR_Primary_plunger.transy = 0
  End If

  If keycode = LeftTiltKey Then Nudge 90, 0.5 ': SoundNudgeLeft
  If keycode = RightTiltKey Then Nudge 270, 0.5 ': SoundNudgeRight
  If keycode = CenterTiltKey Then Nudge 0, 0.5 ': SoundNudgeCenter

    If Keycode = LeftFlipperKey Then
    VR_CabFlipperLeft.X = 2107.934 + 8
  End if
    If Keycode = RightFlipperKey Then
    VR_CabFlipperRight.X = 2100 - 8
  End if

  'If keycode = keyInsertCoin1 or keycode = keyInsertCoin2 or keycode = keyInsertCoin3 or keycode = keyInsertCoin4 Then 'Use this for ROM based games
  If keycode = AddCreditKey or keycode = AddCreditKey2 Then
    Select Case Int(rnd*3)
      Case 0: PlaySound ("Coin_In_1"), 0, CoinSoundLevel, 0, 0.25
      Case 1: PlaySound ("Coin_In_2"), 0, CoinSoundLevel, 0, 0.25
      Case 2: PlaySound ("Coin_In_3"), 0, CoinSoundLevel, 0, 0.25
    End Select
  End If

  If keycode = StartGameKey Then
    SoundStartButton
    VR_Cab_StartButton.transx = -5    ' Animate VR Startbutton
  End If

    If vpmKeyDown(keycode) Then Exit Sub
End Sub


Sub Table1_KeyUp(ByVal keycode)

  If keycode = PlungerKey Then
    Plunger.firespeed = RndInt(108,110)   ' Generate random value variance of +/- 1 from 109
    Plunger.Fire
    If BIPL = 1 Then
      SoundPlungerReleaseBall()     'Plunger release sound when there is a ball in shooter lane
    Else
      SoundPlungerReleaseNoBall()     'Plunger release sound when there is no ball in shooter lane
    End If
    VR_Plunger.Enabled = False
    VR_Plunger2.Enabled = True
    VR_Primary_plunger.transy = 0
  End If
  If Keycode = StartGameKey Then
    VR_Cab_StartButton.transx = 0   ' Animate VR Startbutton
  End If

  If Keycode = LeftFlipperKey Then
    VR_CabFlipperLeft.x = 2107.934
  End if
  If keycode = RightFlipperKey Then
    VR_CabFlipperRight.X = 2100
  End If

  if vpmKeyUp(keycode) Then Exit Sub
End Sub


'*******************************************
'  Flippers
'*******************************************

Const ReflipAngle = 20

' Flipper Solenoid Callbacks (these subs mimics how you would handle flippers in ROM based tables)
Sub SolLFlipper(Enabled) 'Left flipper solenoid callback
  If Enabled Then
    FlipperActivate LeftFlipper, LFPress
    LF.Fire  'leftflipper.rotatetoend

    If leftflipper.currentangle < leftflipper.endangle + ReflipAngle Then
      RandomSoundReflipUpLeft LeftFlipper
    Else
      SoundFlipperUpAttackLeft LeftFlipper
      RandomSoundFlipperUpLeft LeftFlipper
    End If
  Else
    FlipperDeActivate LeftFlipper, LFPress
    LeftFlipper.RotateToStart
    If LeftFlipper.currentangle < LeftFlipper.startAngle - 5 Then
      RandomSoundFlipperDownLeft LeftFlipper
    End If
    FlipperLeftHitParm = FlipperUpSoundLevel
  End If
End Sub

Sub SolRFlipper(Enabled) 'Right flipper solenoid callback
  If Enabled Then
    FlipperActivate RightFlipper, RFPress
    RF.Fire 'Rightflipper.rotatetoend

    If rightflipper.currentangle > rightflipper.endangle - ReflipAngle Then
      RandomSoundReflipUpRight RightFlipper
    Else
      SoundFlipperUpAttackRight RightFlipper
      RandomSoundFlipperUpRight RightFlipper
    End If
  Else
    FlipperDeActivate RightFlipper, RFPress
    RightFlipper.RotateToStart
    If RightFlipper.currentangle > RightFlipper.startAngle + 5 Then
      RandomSoundFlipperDownRight RightFlipper
    End If
    FlipperRightHitParm = FlipperUpSoundLevel
  End If
End Sub

' Flipper collide subs
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


' This subroutine updates the flipper shadows and visual primitives
Sub FlipperVisualUpdate
    FlipperLSh.RotZ = LeftFlipper.currentangle
  FlipperRSh.RotZ = RightFlipper.currentangle
  lflipb.RotZ = LeftFlipper.CurrentAngle
  lflipr.RotZ = LeftFlipper.CurrentAngle
  rflipb.RotZ = RightFlipper.CurrentAngle
  rflipr.RotZ = RightFlipper.CurrentAngle
  Pgate1.rotx = -Gate1.currentangle*0.5
  Pgate2.rotx = -Gate2.currentangle*0.5
End Sub


'******************************************************
'         Solenoids Routines
'******************************************************

Dim GameInPlay:GameInPlay=0
Dim GameInPlayFirstTime:  GameInPlayFirstTime = 0

Sub SolGameOver(enabled)
  If enabled Then
    GameInPlay = 1
    GameInPlayFirstTime = 1
  Else
    GameInPlay = 0
  End If
End Sub

Sub SolGI(enabled)
  If enabled Then
    PFGI 0
  Else
    PFGI 1
  End If
End Sub


'******************************************************
'     TROUGH BASED ON FOZZY'S
'******************************************************

'Handle Trough

' *** NOTE:  The Ball Release is handled under the Controller.Lamp(12) call in the updatelamps callback routines

Sub BallRelease_Hit : Controller.Switch(14) = 1 : UpdateTrough : End Sub
Sub BallRelease_UnHit :Controller.Switch(14) = 0 : UpdateTrough : End Sub
Sub Slot1_Hit()   : Controller.Switch(44) = 1:  UpdateTrough:End Sub
Sub Slot1_UnHit() : Controller.Switch(44) = 0:  UpdateTrough:End Sub

Sub UpdateTrough()
  UpdateTroughTimer.Interval = 300
  UpdateTroughTimer.Enabled = 1
End Sub

Sub UpdateTroughTimer_Timer()
  If Slot1.BallCntOver = 0 Then Drain.kick 60, 12
  Me.Enabled = 0
End Sub

'Drain & Release
Sub Drain_Hit()
  RandomSoundDrain Drain
  UpdateTrough
End Sub


Sub SolTrough(enabled)
  If enabled Then
    If BallRelease.BallCntOver = 0 Then Slot1.kick 60, 8
  End If
End Sub



'*********** Knocker

Sub SolKnocker(enabled)
  If Enabled Then
    If GameInPlayFirstTime = 1 Then
      KnockerSolenoid
    End If
  End If
End Sub


'*********** Saucer

Dim ballinsaucer:  ballinsaucer = 0

Sub SolSaucer(Enabled)
  If Enabled Then
  If sw4.BallCntOver = 1 Then
    sw4.kick 167, 13
    sw4.uservalue=1
    sw4.timerenabled=1
    PkickarmSW4.rotz=15
  End If
  End if
End Sub


Sub sw4_timer
  select case sw4.uservalue
    case 2:
      PkickarmSW4.rotz=0
      me.timerenabled=0
  end Select
  sw4.uservalue=sw4.uservalue+1
End Sub


Sub sw4_Hit
  SoundSaucerLock
  Controller.Switch(4) = 1
  ballinsaucer = 1
End Sub

Sub sw4_Unhit
  SoundSaucerKick 1, sw4
  controller.switch(4) = 0
  ballinsaucer = 0
End Sub




'******************************************************
'             Switches
'******************************************************

'*******************************************
' Bumpers
'*******************************************

Sub Bumper1_Hit 'Left
  RandomSoundBumperMiddle Bumper1
  vpmTimer.PulseSw 45
  DOF 101, DOFPulse 'ROM isn't sending solenoid call... matches what's in DOF config utility
End Sub

Sub Bumper2_Hit 'Middle
  RandomSoundBumperTop Bumper2
  vpmTimer.PulseSw 45
  DOF 102, DOFPulse 'ROM isn't sending solenoid call... matches what's in DOF config utility
End Sub

Sub Bumper3_Hit 'Right
  RandomSoundBumperMiddle Bumper3
  vpmTimer.PulseSw 45
  DOF 103, DOFPulse 'ROM isn't sending solenoid call... matches what's in DOF config utility
End Sub


'*******************************************
'Slings
'*******************************************

Dim Lstep, Rstep

Sub LeftSlingShot_Slingshot
  LS.VelocityCorrect(ActiveBall)
  RandomSoundSlingshotLeft slingL
    DOF 104, DOFPulse
  vpmTimer.PulseSw 55
  LSling.Visible = 0
  LSling1.Visible = 1
  slingL.objroty = 15
  LStep = 0
  LeftSlingShot.TimerEnabled = 1
End Sub

Sub LeftSlingShot_Timer
  Select Case LStep
    Case 2:LSLing1.Visible = 0:LSLing2.Visible = 1:slingL.objroty = 7
    Case 3:LSLing2.Visible = 0:LSLing.Visible = 1:slingL.objroty = 0:LeftSlingShot.TimerEnabled = 0
  End Select
  LStep = LStep + 1
End Sub

Sub RightSlingShot_Slingshot
  RS.VelocityCorrect(ActiveBall)
  RandomSoundSlingshotLeft slingR
    DOF 105, DOFPulse
  vpmTimer.PulseSw 55
  RSling.Visible = 0
  RSling1.Visible = 1
  slingR.objroty = -15
  RStep = 0
  RightSlingShot.TimerEnabled = 1
End Sub

Sub RightSlingShot_Timer
  Select Case RStep
    Case 2:RSLing1.Visible = 0:RSLing2.Visible = 1:slingR.objroty = -7
    Case 3:RSLing2.Visible = 0:RSLing.Visible = 1:slingR.objroty = 0:RightSlingShot.TimerEnabled = 0
  End Select
  RStep = RStep + 1
End Sub


'*******************************************
'Standup Targets
'*******************************************

Sub sw1_Hit
  STHit 1
End Sub

Sub sw1o_Hit
  TargetBouncer ActiveBall, 1
End Sub

Sub sw2_Hit
  STHit 2
End Sub

Sub sw2o_Hit
  TargetBouncer ActiveBall, 1
End Sub

Sub sw3_Hit
  STHit 3
End Sub

Sub sw3o_Hit
  TargetBouncer ActiveBall, 1
End Sub

Sub sw11_Hit
  STHit 11
End Sub

Sub sw11o_Hit
  TargetBouncer ActiveBall, 1
End Sub

Sub sw41_Hit
  STHit 41
End Sub

Sub sw41o_Hit
  TargetBouncer ActiveBall, 1
End Sub

Sub sw51_Hit
  STHit 51
End Sub

Sub sw51o_Hit
  TargetBouncer ActiveBall, 1
End Sub


'*******************************************
'Sensors
'*******************************************

Sub Sensor55_Hit(idx)
  vpmTimer.PulseSw 55
End Sub


'*******************************************
'Rollovers
'*******************************************

Sub sw00_Hit:Controller.Switch(0) = 1:End Sub
Sub sw00_UnHit:Controller.Switch(0) = 0:End Sub
Sub sw10_Hit:Controller.Switch(10) = 1:End Sub
Sub sw10_UnHit:Controller.Switch(10) = 0:End Sub
Sub sw12_Hit:Controller.Switch(12) = 1:End Sub
Sub sw12_UnHit:Controller.Switch(12) = 0:End Sub
Sub sw13_Hit:Controller.Switch(13) = 1:End Sub
Sub sw13_UnHit:Controller.Switch(13) = 0:End Sub
Sub sw40_Hit:Controller.Switch(40) = 1:End Sub
Sub sw40_UnHit:Controller.Switch(40) = 0:End Sub
Sub sw42_Hit:Controller.Switch(42) = 1:End Sub
Sub sw42_UnHit:Controller.Switch(42) = 0:End Sub
Sub sw43_Hit:Controller.Switch(43) = 1:End Sub
Sub sw43_UnHit:Controller.Switch(43) = 0:End Sub
Sub sw50_Hit:Controller.Switch(50) = 1:End Sub
Sub sw50_UnHit:Controller.Switch(50) = 0:End Sub

''''' Switch Animations

Sub sw00_Animate: psw00.transz = sw00.CurrentAnimOffset: End Sub
Sub sw10_Animate: psw10.transz = sw10.CurrentAnimOffset: End Sub
Sub sw12_Animate: psw12.transz = sw12.CurrentAnimOffset: End Sub
Sub sw13_Animate: psw13.transz = sw13.CurrentAnimOffset: End Sub
Sub sw40_Animate: psw40.transz = sw40.CurrentAnimOffset: End Sub
Sub sw42_Animate: psw42.transz = sw42.CurrentAnimOffset: End Sub
Sub sw43_Animate: psw43.transz = sw43.CurrentAnimOffset: End Sub
Sub sw50_Animate: psw50.transz = sw50.CurrentAnimOffset: End Sub


Sub BIPL_Trigger_Hit
  BIPL=True
End Sub
Sub BIPL_Trigger_UnHit
  BIPL=False
End Sub

'*******************************************
' Spinners
'*******************************************

Sub Spinner_Spin()
  vpmtimer.PulseSw 53
  SoundSpinner Spinner
End Sub

'***********Rotate Spinner

Dim SpinnerRadius: SpinnerRadius=7

Sub SpinnerTimer
  SpinnerPrim.Rotx = Spinner.CurrentAngle
  SpinnerRod.TransZ = (cos((Spinner.CurrentAngle + 180) * (PI/180))+1) * SpinnerRadius
  SpinnerRod.TransY = sin((Spinner.CurrentAngle) * (PI/180)) * -SpinnerRadius



End Sub


'*******************************************
' Gate Triggers
'*******************************************

Sub sw52_Hit(): Controller.Switch(52)=1: Gate3.Damping=0.85: sw52step=0: me.timerenabled = True: End Sub
Sub sw52_UnHit(): Controller.Switch(52)=0: End Sub


Dim sw52step

Sub Sw52_timer()
  Select case sw52step
    Case 0:'Gate3.Damping = .9:'Gate3.GravityFactor = 5
    Case 1:Gate3.Damping = .9:Gate3.GravityFactor = 3
    Case 2:Gate3.Damping = .95
    Case 3:Gate3.Damping = .99999
    Case 4:
    Case 5: me.timerenabled = false:sw52step = 0:Gate3.Damping = .85:Gate3.GravityFactor = 5
  End Select
  sw52step = sw52step + 1
End Sub


Sub Gate3_Animate
  P_Gate3.rotx = Gate3.CurrentAngle
  pScoreSpring.rotz = -Gate3.CurrentAngle*4/90
End Sub


'*******************************************
'DIPS code
'*******************************************
'Finding an individual dip state based on scapino's Strikes and spares dip code - from unclewillys pinball pool, added another section to get high score award to set replay cards

Dim TheDips(32)
Sub FindDips
  Dim DipsNumber
  DipsNumber = Controller.Dip(1)
  TheDips(16) = Int(DipsNumber/128)
  If TheDips(16) = 1 then DipsNumber = DipsNumber - 128 end if
  TheDips(15) = Int(DipsNumber/64)
  If TheDips(15) = 1 then DipsNumber = DipsNumber - 64 end if
  TheDips(14) = Int(DipsNumber/32)
  If TheDips(14) = 1 then DipsNumber = DipsNumber - 32 end if
  TheDips(13) = Int(DipsNumber/16)
  If TheDips(13) = 1 then DipsNumber = DipsNumber - 16 end if
  TheDips(12) = Int(DipsNumber/8)
  If TheDips(12) = 1 then DipsNumber = DipsNumber - 8 end if
  TheDips(11) = Int(DipsNumber/4)
  If TheDips(11) = 1 then DipsNumber = DipsNumber - 4 end if
  TheDips(10) = Int(DipsNumber/2)
  If TheDips(10) = 1 then DipsNumber = DipsNumber - 2 end if
  TheDips(9) = Int(DipsNumber)
  DipsNumber = Controller.Dip(2)
  TheDips(24) = Int(DipsNumber/128)
  If TheDips(24) = 1 then DipsNumber = DipsNumber - 128 end if
  TheDips(23) = Int(DipsNumber/64)
  If TheDips(23) = 1 then DipsNumber = DipsNumber - 64 end if
  TheDips(22) = Int(DipsNumber/32)
  If TheDips(22) = 1 then DipsNumber = DipsNumber - 32 end if
  TheDips(21) = Int(DipsNumber/16)
  If TheDips(21) = 1 then DipsNumber = DipsNumber - 16 end if
  TheDips(20) = Int(DipsNumber/8)
  If TheDips(20) = 1 then DipsNumber = DipsNumber - 8 end if
  TheDips(19) = Int(DipsNumber/4)
  If TheDips(19) = 1 then DipsNumber = DipsNumber - 4 end if
  TheDips(18) = Int(DipsNumber/2)
  If TheDips(18) = 1 then DipsNumber = DipsNumber - 2 end if
  TheDips(17) = Int(DipsNumber)
  DipsNumber = Controller.Dip(3)
  TheDips(32) = Int(DipsNumber/128)
  If TheDips(32) = 1 then DipsNumber = DipsNumber - 128 end if
  TheDips(31) = Int(DipsNumber/64)
  If TheDips(31) = 1 then DipsNumber = DipsNumber - 64 end if
  TheDips(30) = Int(DipsNumber/32)
  If TheDips(30) = 1 then DipsNumber = DipsNumber - 32 end if
  TheDips(29) = Int(DipsNumber/16)
  If TheDips(29) = 1 then DipsNumber = DipsNumber - 16 end if
  TheDips(28) = Int(DipsNumber/8)
  If TheDips(28) = 1 then DipsNumber = DipsNumber - 8 end if
  TheDips(27) = Int(DipsNumber/4)
  If TheDips(27) = 1 then DipsNumber = DipsNumber - 4 end if
  TheDips(26) = Int(DipsNumber/2)
  If TheDips(26) = 1 then DipsNumber = DipsNumber - 2 end if
  TheDips(25) = Int(DipsNumber)
  DipsTimer.Enabled=1
End Sub


Sub DipsTimer_Timer()
  dim match
  match = TheDips(26)
  if match = 1 Then
    If DesktopMode = True or VR_Room = 1 Then
      NTM=1
    else
      NTM=0
    end if
  end If
  DipsTimer.enabled=0
End Sub

Sub editDips
  if SetDIPSwitches = 1 Then
    Dim vpmDips : Set vpmDips = New cvpmDips
    With vpmDips
    .AddForm 700,400,"System 80A with Sound only (S)board - DIP switches"
    .AddFrame 0,0,190,"Maximum credits",49152,Array("8 credits",0,"10 credits",32768,"15 credits",&H00004000,"25 credits",49152)'dip 15&16
    .AddFrame 0,76,190,"Coin chute 1 and 2 control",&H00002000,Array("seperate",0,"same",&H00002000)'dip 14
    .AddFrame 0,122,190,"3rd coin chute credits control",&H20000000,Array("no effect",0,"add 9",&H20000000)'dip 30
    .AddFrame 205,0,190,"High score to date awards",&H00C00000,Array("not displayed and no award",0,"displayed and no award",&H00800000,"displayed and 2 credits",&H00400000,"displayed and 3 credits",&H00C00000)'dip 23&24
    .AddFrame 205,76,190,"Game mode",&H10000000,Array("replay",0,"extra ball",&H10000000)'dip 29
    .AddFrame 205,122,190,"Playfield special",&H00200000,Array("replay",0,"extra ball",&H00200000)'dip 22
    .AddFrame 205,168,190,"Replay limit",&H04000000,Array("no limit",0,"one per game",&H04000000)'dip 27
    .AddFrame 205,214,190,"Balls per game",&H01000000,Array("5 balls",0,"3 balls",&H01000000)'dip 25
    .AddFrame 205,260,190,"Novelty mode",&H08000000,Array("normal game mode",0,"50K per special/extra ball",&H08000000)'dip 28
    .AddChk 0,170,190,Array("Match feature",&H02000000)'dip 26
    .AddChk 0,185,190,Array("Background sound",&H40000000)'dip 31
    .AddLabel 50,310,300,20,"After hitting OK, press F3 to reset game with new settings."
    End With
  Dim extra
    extra = Controller.Dip(4) + Controller.Dip(5)*256
    extra = vpmDips.ViewDipsExtra(extra)
    Controller.Dip(4) = extra And 255
    Controller.Dip(5) = (extra And 65280)\256 And 255
  End If
End Sub
Set vpmShowDips = GetRef("editDips")

'  *** This is the default dips settings ***

  '1-5 = 0 left coin chute
  '6-8 = 0 spares
  '9-13 = 0 right coin chute
  '14 = 1 coin chute the same
  '15 = 1 max credits = 25
  '16 = 1 max credits =25
  '17-21 = 0 : center coin chute
  '22 = 1 in normal mode:  0 for tourney... this is extra ball in playfield special
  '23 = 1 HSTD no reward - displayed and 3 replays
  '24 = 1 HSTD no reward - displayed and 3 replays
  '25 = 1: 3 Balls per game... 0 = 5 balls
  '26 = 1: Match enabled
  '27 = 1: replay limit =1
  '28 = 0: Novelty for normal.... 1 for tourney
  '29 = 0: Game mode.  0 = replay ... 1 = extra ball.
  '30 = 0 : this is 3rd coin chute no effect
  '31 = 1 : background sound
  '32 = 0 : spare



Sub SetDefaultDips
  if Tourney = 1 Then
    Controller.Dip(0) = 0   '8-1 = 00000000
    Controller.Dip(1) = 224   '16-9 = 11100000
    Controller.Dip(2) = 192   '24-17 = 11000000  'tourney mode turns off playfield special of extra ball
    Controller.Dip(3) = 79    '32-25 = 01001111 'tourney mode turns to novelty mode(28)
    Controller.Dip(4) = 0
    Controller.Dip(5) = 0
  Else
    If SetDIPSwitches = 0 Then
      Controller.Dip(0) = 0   '8-1 = 00000000
      Controller.Dip(1) = 224   '16-9 = 11100000
      Controller.Dip(2) = 224   '24-17 = 11100000
      Controller.Dip(3) = 71    '32-25 = 01000111
      Controller.Dip(4) = 0
      Controller.Dip(5) = 0
    End If
  End If
End Sub


'********************************************
'              Display Output
'********************************************

'**********************************************************************************************************
' Desktop LED Displays
'**********************************************************************************************************


Dim Digits(32)

Digits(0)=Array(a001,a002,a003,a004,a005,a006,a007,LXM,a008)
Digits(1)=Array(a1,a2,a3,a4,a5,a6,a7,LXM,a8)
Digits(2)=Array(a9,a10,a11,a12,a13,a14,a15,LXM,a16)
Digits(3)=Array(a17,a18,a19,a20,a21,a22,a23,LXM,a24)
Digits(4)=Array(a25,a26,a27,a28,a29,a30,a31,LXM,a32)
Digits(5)=Array(a33,a34,a35,a36,a37,a38,a39,LXM,a40)

Digits(6)=Array(a41,a42,a43,a44,a45,a46,a47,LXM,a48)
Digits(7)=Array(a009,a010,a011,a012,a013,a014,a015,LXM,a016)
Digits(8)=Array(a49,a50,a51,a52,a53,a54,a55,LXM,a56)
Digits(9)=Array(a57,a58,a59,a60,a61,a62,a63,LXM,a64)
Digits(10)=Array(a65,a66,a67,a68,a69,a70,a71,LXM,a72)
Digits(11)=Array(a73,a74,a75,a76,a77,a78,a79,LXM,a80)

Digits(12)=Array(a81,a82,a83,a84,a85,a86,a87,LXM,a88)
Digits(13)=Array(a89,a90,a91,a92,a93,a94,a95,LXM,a96)
Digits(14)=Array(a017,a018,a019,a020,a021,a022,a023,LXM,a024)
Digits(15)=Array(a97,a98,a99,a100,a101,a102,a103,LXM,a104)
Digits(16)=Array(a105,a106,a107,a108,a109,a110,a111,LXM,a112)
Digits(17)=Array(a113,a114,a115,a116,a117,a118,a119,LXM,a120)

Digits(18)=Array(a121,a122,a123,a124,a125,a126,a127,LXM,a128)
Digits(19)=Array(a129,a130,a131,a132,a133,a134,a135,LXM,a136)
Digits(20)=Array(a137,a138,a139,a140,a141,a142,a143,LXM,a144)
Digits(21)=Array(a025,a026,a027,a028,a029,a030,a031,LXM,a032)
Digits(22)=Array(a145,a146,a147,a148,a149,a150,a151,LXM,a152)
Digits(23)=Array(a153,a154,a155,a156,a157,a158,a159,LXM,a160)

Digits(24)=Array(a161,a162,a163,a164,a165,a166,a167,LXM,a168)
Digits(25)=Array(a169,a170,a171,a172,a173,a174,a175,LXM,a176)
Digits(26)=Array(a177,a178,a179,a180,a181,a182,a183,LXM,a184)
Digits(27)=Array(a185,a186,a187,a188,a189,a190,a191,LXM,a192)

'Ball in Play and Credit displays

Digits(28)=Array(f00,f01,f02,f03,f04,f05,f06,LXM)
Digits(29)=Array(f10,f11,f12,f13,f14,f15,f16,LXM)
Digits(30)=Array(e00,e01,e02,e03,e04,e05,e06,LXM)
Digits(31)=Array(e10,e11,e12,e13,e14,e15,e16,LXM)


Sub DisplayTimer
  Dim ChgLED,ii,num,chg,stat,obj
  ChgLed = Controller.ChangedLEDs(&Hffffffff, &Hffffffff)
  If Not IsEmpty(ChgLED) Then
    For ii = 0 To UBound(chgLED)
      num = chgLED(ii, 0) : chg = chgLED(ii, 1) : stat = chgLED(ii, 2)
      if (num < 32) then
        For Each obj In Digits(num)
          If cab_mode = 1 OR VR_Room =1 Then
            obj.intensity=0
          Else
            obj.intensity= 1.5
          End If
          If chg And 1 Then obj.State = stat And 1
          chg = chg\2 : stat = stat\2
        Next
            Else
                For Each obj In Digits(num)
                    If chg And 1 Then obj.State = stat And 1
                    chg = chg \ 2:stat = stat \ 2
                Next
      end If
    next
  end if
End Sub


'*****************************************************************************************************
' VR LED Displays
'*****************************************************************************************************

dim DisplayColor
DisplayColor =  RGB(1,48,135)


Dim VRDigits(32)

VRDigits(0) = Array(LED1x0,LED1x1,LED1x2,LED1x3,LED1x4,LED1x5,LED1x6, n1, LED1x8)
VRDigits(1) = Array(LED2x0,LED2x1,LED2x2,LED2x3,LED2x4,LED2x5,LED2x6, n1, LED2x8)
VRDigits(2) = Array(LED3x0,LED3x1,LED3x2,LED3x3,LED3x4,LED3x5,LED3x6, n1, LED3x8)
VRDigits(3) = Array(LED4x0,LED4x1,LED4x2,LED4x3,LED4x4,LED4x5,LED4x6, n1, LED4x8)
VRDigits(4) = Array(LED5x0,LED5x1,LED5x2,LED5x3,LED5x4,LED5x5,LED5x6, n1, LED5x8)
VRDigits(5) = Array(LED6x0,LED6x1,LED6x2,LED6x3,LED6x4,LED6x5,LED6x6, n1, LED6x8)
VRDigits(6) = Array(LED7x0,LED7x1,LED7x2,LED7x3,LED7x4,LED7x5,LED7x6, n1, LED7x8)

VRDigits(7) = Array(LED8x0,LED8x1,LED8x2,LED8x3,LED8x4,LED8x5,LED8x6, n1, LED8x8)
VRDigits(8) = Array(LED9x0,LED9x1,LED9x2,LED9x3,LED9x4,LED9x5,LED9x6, n1, LED9x8)
VRDigits(9) = Array(LED10x0,LED10x1,LED10x2,LED10x3,LED10x4,LED10x5,LED10x6, n1, LED10x8)
VRDigits(10) = Array(LED11x0,LED11x1,LED11x2,LED11x3,LED11x4,LED11x5,LED11x6, n1, LED11x8)
VRDigits(11) = Array(LED12x0,LED12x1,LED12x2,LED12x3,LED12x4,LED12x5,LED12x6, n1, LED12x8)
VRDigits(12) = Array(LED13x0,LED13x1,LED13x2,LED13x3,LED13x4,LED13x5,LED13x6, n1, LED13x8)
VRDigits(13) = Array(LED14x0,LED14x1,LED14x2,LED14x3,LED14x4,LED14x5,LED14x6, n1, LED14x8)

VRDigits(14) = Array(LED1x000,LED1x001,LED1x002,LED1x003,LED1x004,LED1x005,LED1x006, n1, LED1x008)
VRDigits(15) = Array(LED1x100,LED1x101,LED1x102,LED1x103,LED1x104,LED1x105,LED1x106, n1, LED1x108)
VRDigits(16) = Array(LED1x200,LED1x201,LED1x202,LED1x203,LED1x204,LED1x205,LED1x206, n1, LED1x208)
VRDigits(17) = Array(LED1x300,LED1x301,LED1x302,LED1x303,LED1x304,LED1x305,LED1x306, n1, LED1x308)
VRDigits(18) = Array(LED1x400,LED1x401,LED1x402,LED1x403,LED1x404,LED1x405,LED1x406, n1, LED1x408)
VRDigits(19) = Array(LED1x500,LED1x501,LED1x502,LED1x503,LED1x504,LED1x505,LED1x506, n1, LED1x508)
VRDigits(20) = Array(LED1x600,LED1x601,LED1x602,LED1x603,LED1x604,LED1x605,LED1x606, n1, LED1x608)

VRDigits(21) = Array(LED2x000,LED2x001,LED2x002,LED2x003,LED2x004,LED2x005,LED2x006, n1, LED2x008)
VRDigits(22) = Array(LED2x100,LED2x101,LED2x102,LED2x103,LED2x104,LED2x105,LED2x106, n1, LED2x108)
VRDigits(23) = Array(LED2x200,LED2x201,LED2x202,LED2x203,LED2x204,LED2x205,LED2x206, n1, LED2x208)
VRDigits(24) = Array(LED2x300,LED2x301,LED2x302,LED2x303,LED2x304,LED2x305,LED2x306, n1, LED2x308)
VRDigits(25) = Array(LED2x400,LED2x401,LED2x402,LED2x403,LED2x404,LED2x405,LED2x406, n1, LED2x408)
VRDigits(26) = Array(LED2x500,LED2x501,LED2x502,LED2x503,LED2x504,LED2x505,LED2x506, n1, LED2x508)
VRDigits(27) = Array(LED2x600,LED2x601,LED2x602,LED2x603,LED2x604,LED2x605,LED2x606, n1, LED2x608)

'Ball in Play and Credit displays

VRDigits(28) = Array(LEDax300,LEDax301,LEDax302,LEDax303,LEDax304,LEDax305,LEDax306, n1)
VRDigits(29) = Array(LEDbx400,LEDbx401,LEDbx402,LEDbx403,LEDbx404,LEDbx405,LEDbx406, n1)
VRDigits(30) = Array(LEDcx500,LEDcx501,LEDcx502,LEDcx503,LEDcx504,LEDcx505,LEDcx506, n1)
VRDigits(31) = Array(LEDdx600,LEDdx601,LEDdx602,LEDdx603,LEDdx604,LEDdx605,LEDdx606, n1)


Sub VRDisplayTimer
  Dim ii, jj, obj, b, x
  Dim ChgLED,num, chg, stat
  ChgLED=Controller.ChangedLEDs(&Hffffffff, &Hffffffff)
    If Not IsEmpty(ChgLED) Then
      For ii=0 To UBound(chgLED)
        num=chgLED(ii, 0) : chg=chgLED(ii, 1) : stat=chgLED(ii, 2)
        if num < 32 Then
          For Each obj In VRDigits(num)
   '                  If chg And 1 Then obj.visible=stat And 1    'if you use the object color for off; turn the display object visible to not visible on the playfield, and uncomment this line out.
            If chg And 1 Then FadeDisplay obj, stat And 1
            chg=chg\2 : stat=stat\2
          Next
        Else
          For Each obj In VRDigits(num)
            If chg And 1 Then obj.State = stat And 1
            chg = chg \ 2:stat = stat \ 2
          Next
        End If
      Next
    End If
End Sub


Sub FadeDisplay(object, onoff)
  If OnOff = 1 Then
    object.color = DisplayColor
    Object.Opacity = 5
  Else
    Object.Color = RGB(1,1,1)
    Object.Opacity = 1
  End If
End Sub

Sub InitDigits()
  dim tmp, x, xobj
  for x = 0 to uBound(VRDigits)
    if IsArray(VRDigits(x) ) then
      For each xobj in VRDigits(x)
        xobj.height = xobj.height + 0
        FadeDisplay xobj, 0
      next
    end If
  Next
End Sub



If VR_Room=1 Then
  InitDigits
End If


'*******************************************
' Setup VR Backglass
'*******************************************

Sub setup_backglass()

Dim xoff,yoff,zoff,xrot,zscale, xcen,ycen, ix, xx, yy, xobj

  xoff =480
  yoff =15
  zoff =238
  xrot = -85

  zscale = 0.0000001

  xcen =(1032 /2) - (74 / 2)
  ycen = (1020 /2 ) + (194 /2)

  for ix = 0 to 31
    For Each xobj In VRDigits(ix)

      xx = xobj.x

      xobj.x = (xoff - xcen) + xx
      yy = xobj.y ' get the yoffset before it is changed
      xobj.y = yoff

      If (yy < 0.) then
        yy = yy * -1
      end if

      xobj.height = (zoff - ycen) + yy - (yy * (zscale))
      xobj.rotx = xrot
    Next
  Next

  For each xobj in VRBackglassFlash
    xobj.rotx = 175
    xobj.x = 2190.198
    xobj.y = 1008.072
    xobj.z = -570
  Next

  VR_Backglass_Prim.y = VR_Backglass_Prim.y - 0.1

end sub

'*******************************************
'  Block shadows in plunger lane
'*******************************************

dim notrtlane: notrtlane = 1
sub rtlanesense_Hit() : notrtlane = 0 : end Sub
sub rtlanesense_UnHit() : notrtlane = 1 : end Sub

'***************************************************************
'****  VPW DYNAMIC BALL SHADOWS by Iakki, Apophis, and Wylte
'***************************************************************


'Ambient (Room light source)
Const AmbientBSFactor = 0.9  '0 To 1, higher is darker
Const AmbientMovement = 1    '1+ higher means more movement as the ball moves left and right
Const offsetX = 0        'Offset x position under ball (These are if you want to change where the "room" light is for calculating the shadow position,)
Const offsetY = 5        'Offset y position under ball (^^for example 5,5 if the light is in the back left corner)

'Dynamic (Table light sources)
Const DynamicBSFactor = 0.90  '0 To 1, higher is darker
Const Wideness = 20      'Sets how wide the dynamic ball shadows can get (20 +5 thinness is technically most accurate for lights at z ~25 hitting a 50 unit ball)
Const Thinness = 5        'Sets minimum as ball moves away from source

' *** Trim or extend these to match the number of balls/primitives/flashers on the table!  (will throw errors if there aren't enough objects)
Dim objrtx1(2), objrtx2(2)
Dim objBallShadow(2)
Dim OnPF(2)
Dim BallShadowA
BallShadowA = Array (BallShadowA0, BallShadowA1)
Dim DSSources(30), numberofsources', DSGISide(30) 'Adapted for TZ with GI left / GI right

' *** The Shadow Dictionary
Dim bsDict
Set bsDict = New cvpmDictionary
Const bsNone = "None"
Const bsWire = "Wire"
Const bsRamp = "Ramp"
Const bsRampClear = "Clear"

'Initialization
DynamicBSInit

Sub DynamicBSInit()
  Dim iii, source

  'Prepare the shadow objects before play begins
  For iii = 0 To tnob - 1
    Set objrtx1(iii) = Eval("RtxBallShadow" & iii)
    objrtx1(iii).material = "RtxBallShadow" & iii
    objrtx1(iii).z = 1 + iii / 1000 + 0.01  'Separate z for layering without clipping
    objrtx1(iii).visible = 0

    Set objrtx2(iii) = Eval("RtxBall2Shadow" & iii)
    objrtx2(iii).material = "RtxBallShadow2_" & iii
    objrtx2(iii).z = 1 + iii / 1000 + 0.02
    objrtx2(iii).visible = 0

    Set objBallShadow(iii) = Eval("BallShadow" & iii)
    objBallShadow(iii).material = "BallShadow" & iii
    UpdateMaterial objBallShadow(iii).material,1,0,0,0,0,0,AmbientBSFactor,RGB(0,0,0),0,0,False,True,0,0,0,0
    objBallShadow(iii).Z = 1 + iii / 1000 + 0.04
    objBallShadow(iii).visible = 0

    BallShadowA(iii).Opacity = 100 * AmbientBSFactor
    BallShadowA(iii).visible = 0
  Next

  iii = 0

  For Each Source In DynamicSources
    DSSources(iii) = Array(Source.x, Source.y)
    '   If Instr(Source.name , "Left") > 0 Then DSGISide(iii) = 0 Else DSGISide(iii) = 1  'Adapted for TZ with GI left / GI right
    iii = iii + 1
  Next
  numberofsources = iii
End Sub

Sub BallOnPlayfieldNow(onPlayfield, ballNum)  'Whether a ball is currently on the playfield. Only update certain things once, save some cycles
  If onPlayfield Then
    OnPF(ballNum) = True
    bsRampOff gBOT(ballNum).ID
    '   debug.print "Back on PF"
    UpdateMaterial objBallShadow(ballNum).material,1,0,0,0,0,0,AmbientBSFactor,RGB(0,0,0),0,0,False,True,0,0,0,0
    objBallShadow(ballNum).size_x = 5
    objBallShadow(ballNum).size_y = 4.5
    objBallShadow(ballNum).visible = 1
    BallShadowA(ballNum).visible = 0
    BallShadowA(ballNum).Opacity = 100 * AmbientBSFactor
  Else
    OnPF(ballNum) = False
    '   debug.print "Leaving PF"
  End If
End Sub

Sub DynamicBSUpdate
  Dim falloff 'Max distance to light sources, can be changed dynamically if you have a reason
  falloff = 150
  Dim ShadowOpacity1, ShadowOpacity2
  Dim s, LSd, iii
  Dim dist1, dist2, src1, src2
  Dim bsRampType

  'Hide shadow of deleted balls
  For s = UBound(gBOT) + 1 To tnob - 1
    objrtx1(s).visible = 0
    objrtx2(s).visible = 0
    objBallShadow(s).visible = 0
    BallShadowA(s).visible = 0
  Next

  If UBound(gBOT) < lob Then Exit Sub 'No balls in play, exit

  'The Magic happens now
  For s = lob To UBound(gBOT)
    ' *** Normal "ambient light" ball shadow
    'Layered from top to bottom. If you had an upper pf at for example 80 units and ramps even above that, your Elseif segments would be z>110; z<=110 And z>100; z<=100 And z>30; z<=30 And z>20; Else (under 20)

    'Primitive shadow on playfield, flasher shadow in ramps
    If AmbientBallShadowOn = 1 Then
      '** Above the playfield
      If gBOT(s).Z > 30 Then
        If OnPF(s) Then BallOnPlayfieldNow False, s   'One-time update
        bsRampType = getBsRampType(gBOT(s).id)
        '   debug.print bsRampType

        If Not bsRampType = bsRamp Then 'Primitive visible on PF
          objBallShadow(s).visible = 1
          objBallShadow(s).X = gBOT(s).X + (gBOT(s).X - (tablewidth / 2)) / (Ballsize / AmbientMovement) + offsetX
          objBallShadow(s).Y = gBOT(s).Y + offsetY
          objBallShadow(s).size_x = 5 * ((gBOT(s).Z + BallSize) / 80) 'Shadow gets larger and more diffuse as it moves up
          objBallShadow(s).size_y = 4.5 * ((gBOT(s).Z + BallSize) / 80)
          UpdateMaterial objBallShadow(s).material,1,0,0,0,0,0,AmbientBSFactor * (30 / (gBOT(s).Z)),RGB(0,0,0),0,0,False,True,0,0,0,0
        Else 'Opaque, no primitive below
          objBallShadow(s).visible = 0
        End If

        If bsRampType = bsRampClear Or bsRampType = bsRamp Then 'Flasher visible on opaque ramp
          BallShadowA(s).visible = 1
          BallShadowA(s).X = gBOT(s).X + offsetX
          BallShadowA(s).Y = gBOT(s).Y + offsetY + BallSize / 10
          BallShadowA(s).height = gBOT(s).z - BallSize / 4 + s / 1000 'This is technically 1/4 of the ball "above" the ramp, but it keeps it from clipping the ramp
          If bsRampType = bsRampClear Then BallShadowA(s).Opacity = 50 * AmbientBSFactor
        ElseIf bsRampType = bsWire Or bsRampType = bsNone Then 'Turn it off on wires or falling out of a ramp
          BallShadowA(s).visible = 0
        End If

        '** On pf, primitive only
      ElseIf gBOT(s).Z <= 30 And gBOT(s).Z > 20 Then
        If Not OnPF(s) Then BallOnPlayfieldNow True, s
        objBallShadow(s).X = gBOT(s).X + (gBOT(s).X - (tablewidth / 2)) / (Ballsize / AmbientMovement) + offsetX
        objBallShadow(s).Y = gBOT(s).Y + offsetY
        '   objBallShadow(s).Z = gBOT(s).Z + s/1000 + 0.04    'Uncomment (and adjust If/Elseif height logic) if you want the primitive shadow on an upper/split pf

        '** Under pf, flasher shadow only
      Else
        If OnPF(s) Then BallOnPlayfieldNow False, s
        objBallShadow(s).visible = 0
        BallShadowA(s).visible = 1
        BallShadowA(s).X = gBOT(s).X + offsetX
        BallShadowA(s).Y = gBOT(s).Y + offsetY
        BallShadowA(s).height = gBOT(s).z - BallSize / 4 + s / 1000
      End If

      'Flasher shadow everywhere
    ElseIf AmbientBallShadowOn = 2 Then
      If gBOT(s).Z > 30 Then 'In a ramp
        BallShadowA(s).X = gBOT(s).X + offsetX
        BallShadowA(s).Y = gBOT(s).Y + offsetY + BallSize / 10
        BallShadowA(s).height = gBOT(s).z - BallSize / 4 + s / 1000 'This is technically 1/4 of the ball "above" the ramp, but it keeps it from clipping the ramp
      ElseIf gBOT(s).Z <= 30 And gBOT(s).Z > 20 Then 'On pf
        BallShadowA(s).visible = 1
        BallShadowA(s).X = gBOT(s).X + (gBOT(s).X - (tablewidth / 2)) / (Ballsize / AmbientMovement) + offsetX
        BallShadowA(s).Y = gBOT(s).Y + offsetY
        BallShadowA(s).height = 1.04 + s / 1000
      Else 'Under pf
        BallShadowA(s).X = gBOT(s).X + offsetX
        BallShadowA(s).Y = gBOT(s).Y + offsetY
        BallShadowA(s).height = gBOT(s).z - BallSize / 4 + s / 1000
      End If
    End If

    ' *** Dynamic shadows
    If DynamicBallShadowsOn Then
      If gBOT(s).Z < 30 And gBOT(s).X < 835 AND notrtlane = 1 Then 'Parameters for where the shadows can show, here they are not visible above the table (no upper pf) or in the plunger lane
        dist1 = falloff
        dist2 = falloff
        For iii = 0 To numberofsources - 1 'Search the 2 nearest influencing lights
          LSd = Distance(gBOT(s).x, gBOT(s).y, DSSources(iii)(0), DSSources(iii)(1)) 'Calculating the Linear distance to the Source
          If LSd < falloff And gilvl > 0 Then
            '   If LSd < dist2 And ((DSGISide(iii) = 0 And Lampz.State(100)>0) Or (DSGISide(iii) = 1 And Lampz.State(104)>0)) Then  'Adapted for TZ with GI left / GI right
            dist2 = dist1
            dist1 = LSd
            src2 = src1
            src1 = iii
          End If
        Next
        ShadowOpacity1 = 0
        If dist1 < falloff Then
          objrtx1(s).visible = 1
          objrtx1(s).X = gBOT(s).X
          objrtx1(s).Y = gBOT(s).Y
          '   objrtx1(s).Z = gBOT(s).Z - 25 + s/1000 + 0.01 'Uncomment if you want to add shadows to an upper/lower pf
          objrtx1(s).rotz = AnglePP(DSSources(src1)(0), DSSources(src1)(1), gBOT(s).X, gBOT(s).Y) + 90
          ShadowOpacity1 = 1 - dist1 / falloff
          objrtx1(s).size_y = Wideness * ShadowOpacity1 + Thinness
          UpdateMaterial objrtx1(s).material,1,0,0,0,0,0,ShadowOpacity1 * DynamicBSFactor ^ 3,RGB(0,0,0),0,0,False,True,0,0,0,0
        Else
          objrtx1(s).visible = 0
        End If
        ShadowOpacity2 = 0
        If dist2 < falloff Then
          objrtx2(s).visible = 1
          objrtx2(s).X = gBOT(s).X
          objrtx2(s).Y = gBOT(s).Y + offsetY
          '   objrtx2(s).Z = gBOT(s).Z - 25 + s/1000 + 0.02 'Uncomment if you want to add shadows to an upper/lower pf
          objrtx2(s).rotz = AnglePP(DSSources(src2)(0), DSSources(src2)(1), gBOT(s).X, gBOT(s).Y) + 90
          ShadowOpacity2 = 1 - dist2 / falloff
          objrtx2(s).size_y = Wideness * ShadowOpacity2 + Thinness
          UpdateMaterial objrtx2(s).material,1,0,0,0,0,0,ShadowOpacity2 * DynamicBSFactor ^ 3,RGB(0,0,0),0,0,False,True,0,0,0,0
        Else
          objrtx2(s).visible = 0
        End If
        If AmbientBallShadowOn = 1 Then
          'Fades the ambient shadow (primitive only) when it's close to a light
          UpdateMaterial objBallShadow(s).material,1,0,0,0,0,0,AmbientBSFactor * (1 - max(ShadowOpacity1, ShadowOpacity2)),RGB(0,0,0),0,0,False,True,0,0,0,0
        Else
          BallShadowA(s).Opacity = 100 * AmbientBSFactor * (1 - max(ShadowOpacity1, ShadowOpacity2))
        End If
      Else 'Hide dynamic shadows everywhere else, just in case
        objrtx2(s).visible = 0
        objrtx1(s).visible = 0
      End If
    End If
  Next
End Sub

' *** Ramp type definitions

Sub bsRampOnWire()
  If bsDict.Exists(ActiveBall.ID) Then
    bsDict.Item(ActiveBall.ID) = bsWire
  Else
    bsDict.Add ActiveBall.ID, bsWire
  End If
End Sub

Sub bsRampOn()
  If bsDict.Exists(ActiveBall.ID) Then
    bsDict.Item(ActiveBall.ID) = bsRamp
  Else
    bsDict.Add ActiveBall.ID, bsRamp
  End If
End Sub

Sub bsRampOnClear()
  If bsDict.Exists(ActiveBall.ID) Then
    bsDict.Item(ActiveBall.ID) = bsRampClear
  Else
    bsDict.Add ActiveBall.ID, bsRampClear
  End If
End Sub

Sub bsRampOff(idx)
  If bsDict.Exists(idx) Then
    bsDict.Item(idx) = bsNone
  End If
End Sub

Function getBsRampType(id)
  Dim retValue
  If bsDict.Exists(id) Then
    retValue = bsDict.Item(id)
  Else
    retValue = bsNone
  End If
  getBsRampType = retValue
End Function


'****************************************************************
'****  END VPW DYNAMIC BALL SHADOWS by Iakki, Apophis, and Wylte
'****************************************************************


'******************************************************
' FLIPPER CORRECTIONS by nFozzy and updated by Rothbauerw
'******************************************************

'******************************************************
' Flippers Polarity (Select appropriate sub based on era)
'******************************************************

Dim LF : Set LF = New FlipperPolarity
Dim RF : Set RF = New FlipperPolarity

InitPolarity


''*******************************************
''  Late 80's early 90's - Updated by Rothbauerw

Sub InitPolarity()
  dim x, a : a = Array(LF, RF)
  for each x in a
    x.AddPt "Ycoef", 0, RightFlipper.Y-65, 1 'disabled
    x.AddPt "Ycoef", 1, RightFlipper.Y-11, 1
    x.enabled = True
    x.TimeDelay = 60
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
'   Otherwise it should function exactly the same as before

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
      'debug.print "PolarityCorrect" & " " & GameTime & " " & Round(BallPos*100) & "%" & " AddX:" & Round(AddX,2) & " Vel%:" & Round(VelCoef*100) & " " & NoCorrection & " " & PartialFlipcoef
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



'*****************
' Maths
'*****************

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

Const LiveDistanceMin = 30  'minimum distance In vp units from flipper base live catch dampening will occur
Const LiveDistanceMax = 114 'maximum distance in vp units from flipper base live catch dampening will occur (tip protection)

Sub CheckLiveCatch(ball, Flipper, FCount, parm) 'Experimental new live catch
  Dim Dir
  Dir = Flipper.startangle / Abs(Flipper.startangle)  '-1 for Right Flipper
  Dim LiveCatchBounce                                                           'If live catch is not perfect, it won't freeze ball totally
  Dim CatchTime
  CatchTime = GameTime - FCount

  If CatchTime <= LiveCatch And parm > 6 And Abs(Flipper.x - ball.x) > LiveDistanceMin And Abs(Flipper.x - ball.x) < LiveDistanceMax Then
    If CatchTime <= LiveCatch * 0.5 Then                        'Perfect catch only when catch time happens in the beginning of the window
      LiveCatchBounce = 0
    Else
      LiveCatchBounce = Abs((LiveCatch / 2) - CatchTime)    'Partial catch when catch happens a bit late
    End If

    If LiveCatchBounce = 0 And ball.velx * Dir > 0 Then ball.velx = 0
    ball.vely = LiveCatchBounce * (32 / LiveCatch) ' Multiplier for inaccuracy bounce
    ball.angmomx = 0
    ball.angmomy = 0
    ball.angmomz = 0
  Else
    If Abs(Flipper.currentangle) <= Abs(Flipper.endangle) + 1 Then FlippersD.Dampenf ActiveBall, parm
  End If
End Sub

'******************************************************
'****  END FLIPPER CORRECTIONS
'******************************************************


'******************************************************
'   RUBBER  DAMPENERS
'******************************************************
' These are data mined bounce curves,
' dialed in with the in-game elasticity as much as possible to prevent angle / spin issues.
' Requires tracking ballspeed to calculate COR

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
'   VPW TargetBouncer for targets and posts by Iaakki, Wrd1972, Apophis
'******************************************************

Const TargetBouncerEnabled = 1    '0 = normal standup targets, 1 = bouncy targets
Const TargetBouncerFactor = 0.7  'Level of bounces. Recommmended value of 0.7

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

'Add targets or posts to the TargetBounce collection if you want to activate the targetbouncer code from them
Sub TargetBounce_Hit(idx)
  TargetBouncer ActiveBall, 1
End Sub


'******************************************************
' SLINGSHOT CORRECTION FUNCTIONS by apophis
'******************************************************
' To add these slingshot corrections:
'  - On the table, add the endpoint primitives that define the two ends of the Slingshot
'  - Initialize the SlingshotCorrection objects in InitSlingCorrection
'  - Call the .VelocityCorrect methods from the respective _Slingshot event sub

Dim LS, RS
Set LS = New SlingshotCorrection
Set RS = New SlingshotCorrection

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

Function RotPoint(x,y,angle)
  dim rx, ry
  rx = x*dCos(angle) - y*dSin(angle)
  ry = x*dSin(angle) + y*dCos(angle)
  RotPoint = Array(rx,ry)
End Function

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
'****  TARGETS by Rothbauerw
'******************************************************


'******************************************************
'   STAND-UP TARGET INITIALIZATION
'******************************************************

Class StandupTarget
  Private m_primary, m_prim, m_sw, m_animate

  Public Property Get Primary(): Set Primary = m_primary: End Property
  Public Property Let Primary(input): Set m_primary = input: End Property

  Public Property Get Prim(): Set Prim = m_prim: End Property
  Public Property Let Prim(input): Set m_prim = input: End Property

  Public Property Get Sw(): Sw = m_sw: End Property
  Public Property Let Sw(input): m_sw = input: End Property

  Public Property Get Animate(): Animate = m_animate: End Property
  Public Property Let Animate(input): m_animate = input: End Property

  Public default Function init(primary, prim, sw, animate)
    Set m_primary = primary
    Set m_prim = prim
    m_sw = sw
    m_animate = animate

    Set Init = Me
  End Function
End Class


sub DTBallPhysics(aBall, angle, mass)
  dim rangle,bangle,calc1, calc2, calc3
  rangle = (angle - 90) * 3.1416 / 180
  bangle = atn2(cor.ballvely(aball.id),cor.ballvelx(aball.id))

  calc1 = cor.BallVel(aball.id) * cos(bangle - rangle) * (aball.mass - mass) / (aball.mass + mass)
  calc2 = cor.BallVel(aball.id) * sin(bangle - rangle) * cos(rangle + 4*Atn(1)/2)
  calc3 = cor.BallVel(aball.id) * sin(bangle - rangle) * sin(rangle + 4*Atn(1)/2)

  aBall.velx = calc1 * cos(rangle) + calc2
  aBall.vely = calc1 * sin(rangle) + calc3
End Sub


'Define a variable for each stand-up target

Dim ST1, ST2, ST3, ST11, ST41, ST51

'Set array with stand-up target objects
'
'StandupTargetvar = Array(primary, prim, swtich)
'   primary:      vp target to determine target hit
' prim:       primitive target used for visuals and animation
'             IMPORTANT!!!
'             transy must be used to offset the target animation
' switch:       ROM switch number
' animate:      Arrary slot for handling the animation instrucitons, set to 0
'
'You will also need to add a secondary hit object for each stand up (name sw11o, sw12o, and sw13o on the example Table1)
'these are inclined primitives to simulate hitting a bent target and should provide so z velocity on high speed impacts

Set ST1  = (new StandupTarget)(sw1, psw1,1, 0)
Set ST2  = (new StandupTarget)(sw2, psw2,2, 0)
Set ST3  = (new StandupTarget)(sw3, psw3,3, 0)
Set ST11  = (new StandupTarget)(sw11, psw11,11, 0)
Set ST41  = (new StandupTarget)(sw41, psw41,41, 0)
Set ST51  = (new StandupTarget)(sw51, psw51,51, 0)

Dim STArray
STArray = Array(ST1, ST2, ST3, ST11, ST41, ST51)

'Configure the behavior of Stand-up Targets
Const STAnimStep =  1.5         'vpunits per animation step (control return to Start)
Const STMaxOffset = 9       'max vp units target moves when hit

Const STMass = 0.2        'Mass of the Stand-up Target (between 0 and 1), higher values provide more resistance


'******************************************************
'       STAND-UP TARGETS FUNCTIONS
'******************************************************

Sub STHit(switch)
  Dim i
  i = STArrayID(switch)

  PlayTargetSound
  STArray(i).animate = STCheckHit(ActiveBall,STArray(i).primary)

  If STArray(i).animate <> 0 Then
    DTBallPhysics ActiveBall, STArray(i).primary.orientation, STMass
  End If
  DoSTAnim
End Sub

Function STArrayID(switch)
  Dim i
  For i = 0 To UBound(STArray)
    If STArray(i).sw = switch Then
      STArrayID = i
      Exit Function
    End If
  Next
End Function

Function STCheckHit(aBall, target) 'Check if target is hit on it's face
  Dim bangle, bangleafter, rangle, rangle2, perpvel, perpvelafter, paravel, paravelafter
  rangle = (target.orientation - 90) * 3.1416 / 180
  bangle = atn2(cor.ballvely(aball.id),cor.ballvelx(aball.id))
  bangleafter = Atn2(aBall.vely,aball.velx)

  perpvel = cor.BallVel(aball.id) * Cos(bangle - rangle)
  paravel = cor.BallVel(aball.id) * Sin(bangle - rangle)

  perpvelafter = BallSpeed(aBall) * Cos(bangleafter - rangle)
  paravelafter = BallSpeed(aBall) * Sin(bangleafter - rangle)

  If perpvel > 0 And  perpvelafter <= 0 Then
    STCheckHit = 1
  ElseIf perpvel > 0 And ((paravel > 0 And paravelafter > 0) Or (paravel < 0 And paravelafter < 0)) Then
    STCheckHit = 1
  Else
    STCheckHit = 0
  End If
End Function

Sub DoSTAnim()
  Dim i
  For i = 0 To UBound(STArray)
    STArray(i).animate = STAnimate(STArray(i).primary,STArray(i).prim,STArray(i).sw,STArray(i).animate)
  Next
End Sub

Function STAnimate(primary, prim, switch,  animate)
  Dim animtime

  STAnimate = animate

  If animate = 0  Then
    primary.uservalue = 0
    STAnimate = 0
    Exit Function
  ElseIf primary.uservalue = 0 Then
    primary.uservalue = GameTime
  End If

  animtime = GameTime - primary.uservalue

  If animate = 1 Then
    primary.collidable = 0
    prim.transy =  - STMaxOffset
    If UsingROM Then
      vpmTimer.PulseSw Switch
    Else
      STAction switch
    End If
    STAnimate = 2
    Exit Function
  ElseIf animate = 2 Then
    prim.transy = prim.transy + STAnimStep
    If prim.transy >= 0 Then
      prim.transy = 0
      primary.collidable = 1
      STAnimate = 0
      Exit Function
    Else
      STAnimate = 2
    End If
  End If
End Function


Sub STAction(Switch)
'do nothing
End Sub


'******************************************************
'   END STAND-UP TARGETS
'******************************************************


'******************************************************
' BALL ROLLING AND DROP SOUNDS
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

  ' stop the sound of deleted balls
  For b = UBound(gBOT) + 1 To tnob - 1
    ' Comment the next line if you are not implementing Dyanmic Ball Shadows
    If AmbientBallShadowOn = 0 Then BallShadowA(b).visible = 0
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

    ' "Static" Ball Shadows
    ' Comment the next If block, if you are not implementing the Dynamic Ball Shadows
    If AmbientBallShadowOn = 0 Then
      If gBOT(b).Z > 30 Then
        BallShadowA(b).height = gBOT(b).z - BallSize / 4    'This is technically 1/4 of the ball "above" the ramp, but it keeps it from clipping the ramp
      Else
        BallShadowA(b).height = 0.1
      End If
      BallShadowA(b).Y = gBOT(b).Y + offsetY
      BallShadowA(b).X = gBOT(b).X + offsetX
      BallShadowA(b).visible = 1
    End If
  Next
End Sub

'******************************************************
'****  END BALL ROLLING AND DROP SOUNDS
'******************************************************


'******************************************************
'****  FLEEP MECHANICAL SOUNDS
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

' Thalamus, AudioFade - Patched
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

' Thalamus, AudioPan - Patched
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

Sub SoundSaucerKickCenter(scenario, saucer)
  Select Case scenario
    Case 0: PlaySoundAtLevelStatic SoundFX("Saucer_Empty", DOFContactors), SaucerKickSoundLevel, saucer
    Case 1: PlaySoundAtLevelStatic SoundFX("CenterEject", DOFContactors), SaucerKickSoundLevel, saucer
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

Const RelayFlashSoundLevel = 0.05 'volume level; range [0, 1];
Const RelayGISoundLevel = 1     'volume level; range [0, 1];

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
'****  FLEEP MECHANICAL SOUNDS
'******************************************************





'******************************************************
'******  FLUPPER BUMPERS
'******************************************************

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

Dim FlBumperFadeActual(6), FlBumperFadeTarget(6), FlBumperColor(6), FlBumperTop(6), FlBumperCapBase(6), FlBumperSmallLight(6), Flbumperbiglight(6)
Dim FlBumperDisk(6), FlBumperBase(6), FlBumperBulb(6), FlBumperscrews(6), FlBumperActive(6), FlBumperHighlight(6), FLBumperlight(6)
Dim cnt
For cnt = 1 To 6
  FlBumperActive(cnt) = False
Next


FlInitBumper 1, "alien"
FlInitBumper 2, "alien"
FlInitBumper 3, "alien"

Sub FlInitBumper(nr, col)
  FlBumperActive(nr) = True

  ' store all objects in an array for use in FlFadeBumper subroutine
  FlBumperFadeActual(nr) = 1
  FlBumperFadeTarget(nr) = 1.1
  FlBumperColor(nr) = col
  Set FlBumperTop(nr) = Eval("bumpertop" & nr)
  Set FlBumperCapBase(nr) = Eval("bumpercapbase" & nr)
  Set FlBumperBase(nr) = Eval("bumperbase" & nr)
  FlBumperTop(nr).material = "bumpertopmat" & nr
  FlBumperCapBase(nr).material = "bumpertopmata" & nr
  Set FlBumperSmallLight(nr) = Eval("bumpersmalllight" & nr)
  Set Flbumperbiglight(nr) = Eval("bumperbiglight" & nr)
  Set FlBumperDisk(nr) = Eval("bumperdisk" & nr)
  Set FlBumperBase(nr) = Eval("bumperbase" & nr)
  Set FlBumperBulb(nr) = Eval("bumperbulb" & nr)
  FlBumperBulb(nr).material = "bumperbulbmat" & nr
  Set FlBumperscrews(nr) = Eval("bumperscrews" & nr)
  FlBumperscrews(nr).material = "bumperscrew" & col
  Set FlBumperHighlight(nr) = Eval("bumperhighlight" & nr)
  Set FlBumperLight(nr) = Eval("bumperlight" & nr)

  Select Case col

    Case "alien"
      FlBumperSmallLight(nr).color = RGB(248,228,164)
      FlBumperSmallLight(nr).colorfull = RGB(253,245,224)
      FlBumperBigLight(nr).color = RGB(248,228,164)
      FlBumperBigLight(nr).colorfull = RGB(253,245,224)
      FlBumperHighlight(nr).color = RGB(253,245,224)
      FlBumperSmallLight(nr).TransmissionScale = 0
      FlBumperSmallLight(nr).BulbModulateVsAdd = 0.99
      FlBumperTop(nr).BlendDisableLighting = 0.005
      MaterialColor "bumpertopmat" & nr, RGB(255,245,220)
      FlBumperCapBase(nr).BlendDisableLighting = 0.005
      MaterialColor "bumpertopmata" & nr, RGB(255,245,220)
      FlBumperDisk(nr).BlendDisableLighting = 0.5
      FLBumperLight(nr).intensity = 0

  End Select
End Sub

Sub FlFadeBumper(nr, Z)
  '   UpdateMaterial(string, float wrapLighting, float roughness, float glossyImageLerp, float thickness, float edge, float edgeAlpha, float opacity,
  '        OLE_COLOR base, OLE_COLOR glossy, OLE_COLOR clearcoat, VARIANT_BOOL isMetal, VARIANT_BOOL opacityActive,
  '        float elasticity, float elasticityFalloff, float friction, float scatterAngle) - updates all parameters of a material

  FlBumperDisk(nr).BlendDisableLighting = 0.5 * DayNightAdjust + 0.6 * Z

  Select Case FlBumperColor(nr)

    Case "alien"
      FlBumperSmallLight(nr).intensity = 17 + 200 * Z
      FlBumperSmallLight(nr).color = RGB(248,228 - 20*Z,164-65*Z) : FlBumperSmallLight(nr).colorfull = RGB(253,245 - 20*Z,224-65*Z)
      FlBumperBulb(nr).BlendDisableLighting = 12 + 2000 * (0.03 * Z + 0.97 * Z ^ 10)
      Flbumperbiglight(nr).intensity = 2.5 * Z
      FlBumperHighlight(nr).opacity = 1000 * (Z^3)
      FlBumperTop(nr).BlendDisableLighting = 0.015 + 0.25 * Z
      FLBumperCapBase(nr).BlendDisableLighting = 0.015 + 0.35 * Z
      MaterialColor "bumpertopmat" & nr, RGB(255,250 - z*6,220 - Z*15)
      MaterialColor "bumpertopmata" & nr, RGB(255,250 - z*6,220 - Z*15)
      MaterialColor "bumperbase" & nr, RGB(255,235 - z * 36,220 - Z * 90)
      FlBumperBase(nr).BlendDisableLighting = 0.5
      FLBumperLight(nr).intensity =  12 * Z - 2
      FlBumperLight(nr).color = RGB(248,228 - 20*Z,164-65*Z) : FlBumperLight(nr).colorfull = RGB(253,245 - 20*Z,224-65*Z)

  End Select
End Sub

Sub BumperTimer_Timer
  Dim nr
  For nr = 1 To 6
    If FlBumperFadeActual(nr) < FlBumperFadeTarget(nr) And FlBumperActive(nr)  Then
      FlBumperFadeActual(nr) = FlBumperFadeActual(nr) + (FlBumperFadeTarget(nr) - FlBumperFadeActual(nr)) * 0.8
      If FlBumperFadeActual(nr) > 0.99 Then FlBumperFadeActual(nr) = 1
      FlFadeBumper nr, FlBumperFadeActual(nr)
    End If
    If FlBumperFadeActual(nr) > FlBumperFadeTarget(nr) And FlBumperActive(nr)  Then
      FlBumperFadeActual(nr) = FlBumperFadeActual(nr) + (FlBumperFadeTarget(nr) - FlBumperFadeActual(nr)) * 0.4 / (FlBumperFadeActual(nr) + 0.1)
      If FlBumperFadeActual(nr) < 0.01 Then FlBumperFadeActual(nr) = 0
      FlFadeBumper nr, FlBumperFadeActual(nr)
    End If
  Next
End Sub


'******************************************************
'******  END FLUPPER BUMPERS
'******************************************************



'******************************************************
'****  LAMPZ by nFozzy
'******************************************************
'

Dim NullFader
Set NullFader = New NullFadingObject
Dim Lampz
Set Lampz = New LampFader
InitLampsNF         ' Setup lamp assignments

Sub LampTimer()
  If UsingROM Then '
    Dim x, chglamp
    chglamp = Controller.ChangedLamps
    If Not IsEmpty(chglamp) Then
      For x = 0 To UBound(chglamp) 'nmbr = chglamp(x, 0), state = chglamp(x, 1)
        Lampz.state(chglamp(x, 0)) = chglamp(x, 1)
      Next
    End If
  Else
    Dim idx
    For idx = 0 To 150
      If Lampz.IsLight(idx) Then
        If IsArray(Lampz.obj(idx)) Then
          Dim tmp
          tmp = Lampz.obj(idx)
          Lampz.state(idx) = tmp(0).GetInPlayStateBool
          'debug.print tmp(0).name & " " &  tmp(0).GetInPlayStateBool & " " & tmp(0).IntensityScale  & vbnewline
        Else
          Lampz.state(idx) = Lampz.obj(idx).GetInPlayStateBool
          'debug.print Lampz.obj(idx).name & " " &  Lampz.obj(idx).GetInPlayStateBool & " " & Lampz.obj(idx).IntensityScale  & vbnewline
        End If
      End If
    Next
  End If
  Lampz.Update2 'update (fading logic only)
End Sub

Sub DisableLighting(pri, DLintensity, ByVal aLvl) 'cp's script  DLintensity = disabled lighting intesity
  If Lampz.UseFunc Then aLvl = Lampz.FilterOut(aLvl)  'Callbacks don't get this filter automatically
  pri.blenddisablelighting = aLvl * DLintensity
End Sub

sub DisableLighting2(pri, DLintensityMax, DLintensityMin, ByVal aLvl)    'cp's script  DLintensity = disabled lighting intesity
    if Lampz.UseFunc then aLvl = Lampz.FilterOut(aLvl)    'Callbacks don't get this filter automatically
    pri.blenddisablelighting = (aLvl * (DLintensityMax-DLintensityMin)) + DLintensityMin
End Sub

Sub SetModLamp(id, val)
  Lampz.state(id) = val
End Sub



const insert_dl_on_bigyellow_flat = 3
const insert_dl_off_yellow = 1
const yellowbulb = 10

const insert_dl_on_white = 2.75
const insert_dl_off_white = 1
const whitebulb = 10

const insert_dl_on_pink = 3.5
const insert_dl_off_pink = 1
const pinkbulb = 8

const insertbulboff = 0.1

const insert_dl_on_red = 200
const insert_dl_off_red = 4
const insert_dl_on_green = 40
const insert_dl_on_green_ship = 120
const insert_dl_on_blue = 140
const insert_dl_on_orange = 80


Sub InitLampsNF()
  'Filtering (comment out to disable)
  Lampz.Filter = "LampFilter" 'Puts all lamp intensityscale output (no callbacks) through this function before updating

  'Adjust fading speeds (max level / full MS fading time). The Modulate property must be set to 1 / max level if lamp is modulated.
  Dim x
  For x = 0 To 150
    Lampz.FadeSpeedUp(x) = 1 / 40
    Lampz.FadeSpeedDown(x) = 1 / 120
    Lampz.Modulate(x) = 1
  Next

' GI Fading
  Lampz.FadeSpeedUp(140) = 1/40 : Lampz.FadeSpeedDown(140) = 1/240 : Lampz.Modulate(140) = 1

  'Lampz Assignments
  '  In a ROM based table, the lamp ID is used to set the state of the Lampz objects

  'MassAssign is an optional way to do assignments. It'll create arrays automatically / append objects to existing arrays


'Playfield


  Lampz.MassAssign(3)= l3
  Lampz.MassAssign(3)= l3a
  Lampz.Callback(3) = "DisableLighting2 p3, insert_dl_on_pink, 1,"
  Lampz.Callback(3) = "DisableLighting2 bulb3, pinkbulb, insertbulboff,"
  Lampz.Callback(3) = "DisableLighting2 p3off, .1, insert_dl_off_pink,"
  Lampz.MassAssign(4) = l4
  Lampz.MassAssign(4) = l4a
  Lampz.Callback(4) = "DisableLighting p4, insert_dl_on_green,"
  Lampz.MassAssign(5) = l5
  Lampz.MassAssign(5) = l5a
  Lampz.Callback(5) = "DisableLighting p5, insert_dl_on_green,"
  Lampz.MassAssign(6) = l6
  Lampz.MassAssign(6) = l6a
  Lampz.Callback(6) = "DisableLighting p6, insert_dl_on_green,"
  Lampz.MassAssign(7) = l7
  Lampz.MassAssign(7) = l7a
  Lampz.Callback(7) = "DisableLighting p7, insert_dl_on_green,"
  Lampz.MassAssign(13)= l13
  Lampz.MassAssign(13)= l13a
  Lampz.Callback(13) = "DisableLighting2 p13, insert_dl_on_white, 1,"
  Lampz.Callback(13) = "DisableLighting2 bulb13, whitebulb, insertbulboff,"
  Lampz.Callback(13) = "DisableLighting2 p13off, .1, insert_dl_off_white,"
  Lampz.MassAssign(13)= l132
  Lampz.MassAssign(13)= l132a
  Lampz.Callback(13) = "DisableLighting2 p132, insert_dl_on_white, 1,"
  Lampz.Callback(13) = "DisableLighting2 bulb132, whitebulb, insertbulboff,"
  Lampz.Callback(13) = "DisableLighting2 p132off, .1, insert_dl_off_white,"
  Lampz.MassAssign(14)= l14
  Lampz.MassAssign(14)= l14a
  Lampz.Callback(14) = "DisableLighting2 p14, insert_dl_on_white, 1,"
  Lampz.Callback(14) = "DisableLighting2 bulb14, whitebulb, insertbulboff,"
  Lampz.Callback(14) = "DisableLighting2 p14off, .1, insert_dl_off_white,"
  Lampz.MassAssign(14)= l142
  Lampz.MassAssign(14)= l142a
  Lampz.Callback(14) = "DisableLighting2 p142, insert_dl_on_white, 1,"
  Lampz.Callback(14) = "DisableLighting2 bulb142, whitebulb, insertbulboff,"
  Lampz.Callback(14) = "DisableLighting2 p142off, .1, insert_dl_off_white,"
  Lampz.MassAssign(15)= l15
  Lampz.MassAssign(15)= l15a
  Lampz.Callback(15) = "DisableLighting2 p15, insert_dl_on_white, 1,"
  Lampz.Callback(15) = "DisableLighting2 bulb15, whitebulb, insertbulboff,"
  Lampz.Callback(15) = "DisableLighting2 p15off, .1, insert_dl_off_white,"
  Lampz.MassAssign(15)= l152
  Lampz.MassAssign(15)= l152a
  Lampz.Callback(15) = "DisableLighting2 p152, insert_dl_on_white, 1,"
  Lampz.Callback(15) = "DisableLighting2 bulb152, whitebulb, insertbulboff,"
  Lampz.Callback(15) = "DisableLighting2 p152off, .1, insert_dl_off_white,"
  Lampz.MassAssign(16)= l16
  Lampz.MassAssign(16)= l16a
  Lampz.Callback(16) = "DisableLighting2 p16, insert_dl_on_white, 1,"
  Lampz.Callback(16) = "DisableLighting2 bulb16, whitebulb, insertbulboff,"
  Lampz.Callback(16) = "DisableLighting2 p16off, .1, insert_dl_off_white,"
  Lampz.MassAssign(16)= l162
  Lampz.MassAssign(16)= l162a
  Lampz.Callback(16) = "DisableLighting2 p162, insert_dl_on_white, 1,"
  Lampz.Callback(16) = "DisableLighting2 bulb162, whitebulb, insertbulboff,"
  Lampz.Callback(16) = "DisableLighting2 p162off, .1, insert_dl_off_white,"
  Lampz.MassAssign(17)= l17
  Lampz.MassAssign(17)= l17a
  Lampz.Callback(17) = "DisableLighting2 p17, insert_dl_on_white, 1,"
  Lampz.Callback(17) = "DisableLighting2 bulb17, whitebulb, insertbulboff,"
  Lampz.Callback(17) = "DisableLighting2 p17off, .1, insert_dl_off_white,"
  Lampz.MassAssign(17)= l172
  Lampz.MassAssign(17)= l172a
  Lampz.Callback(17) = "DisableLighting2 p172, insert_dl_on_white, 1,"
  Lampz.Callback(17) = "DisableLighting2 bulb172, whitebulb, insertbulboff,"
  Lampz.Callback(17) = "DisableLighting2 p172off, .1, insert_dl_off_white,"
  Lampz.MassAssign(18)= l18
  Lampz.MassAssign(18)= l18a
  Lampz.Callback(18) = "DisableLighting2 p18, insert_dl_on_white, 1,"
  Lampz.Callback(18) = "DisableLighting2 bulb18, whitebulb, insertbulboff,"
  Lampz.Callback(18) = "DisableLighting2 p18off, .1, insert_dl_off_white,"
  Lampz.MassAssign(18)= l182
  Lampz.MassAssign(18)= l182a
  Lampz.Callback(18) = "DisableLighting2 p182, insert_dl_on_white, 1,"
  Lampz.Callback(18) = "DisableLighting2 bulb182, whitebulb, insertbulboff,"
  Lampz.Callback(18) = "DisableLighting2 p182off, .1, insert_dl_off_white,"
  Lampz.MassAssign(19)= l19
  Lampz.MassAssign(19)= l19a
  Lampz.Callback(19) = "DisableLighting2 p19, insert_dl_on_white, 1,"
  Lampz.Callback(19) = "DisableLighting2 bulb19, whitebulb, insertbulboff,"
  Lampz.Callback(19) = "DisableLighting2 p19off, .1, insert_dl_off_white,"
  Lampz.MassAssign(19)= l192
  Lampz.MassAssign(19)= l192a
  Lampz.Callback(19) = "DisableLighting2 p192, insert_dl_on_white, 1,"
  Lampz.Callback(19) = "DisableLighting2 bulb192, whitebulb, insertbulboff,"
  Lampz.Callback(19) = "DisableLighting2 p192off, .1, insert_dl_off_white,"
  Lampz.MassAssign(20)= l20
  Lampz.MassAssign(20)= l20a
  Lampz.Callback(20) = "DisableLighting2 p20, insert_dl_on_white, 1,"
  Lampz.Callback(20) = "DisableLighting2 bulb20, whitebulb, insertbulboff,"
  Lampz.Callback(20) = "DisableLighting2 p20off, .1, insert_dl_off_white,"
  Lampz.MassAssign(20)= l202
  Lampz.MassAssign(20)= l202a
  Lampz.Callback(20) = "DisableLighting2 p202, insert_dl_on_white, 1,"
  Lampz.Callback(20) = "DisableLighting2 bulb202, whitebulb, insertbulboff,"
  Lampz.Callback(20) = "DisableLighting2 p202off, .1, insert_dl_off_white,"
  Lampz.MassAssign(21)= l21
  Lampz.MassAssign(21)= l21a
  Lampz.Callback(21) = "DisableLighting2 p21, insert_dl_on_white, 1,"
  Lampz.Callback(21) = "DisableLighting2 bulb21, whitebulb, insertbulboff,"
  Lampz.Callback(21) = "DisableLighting2 p21off, .1, insert_dl_off_white,"
  Lampz.MassAssign(21)= l212
  Lampz.MassAssign(21)= l212a
  Lampz.Callback(21) = "DisableLighting2 p212, insert_dl_on_white, 1,"
  Lampz.Callback(21) = "DisableLighting2 bulb212, whitebulb, insertbulboff,"
  Lampz.Callback(21) = "DisableLighting2 p212off, .1, insert_dl_off_white,"
  Lampz.MassAssign(22)= l22
  Lampz.MassAssign(22)= l22a
  Lampz.Callback(22) = "DisableLighting2 p22, insert_dl_on_white, 1,"
  Lampz.Callback(22) = "DisableLighting2 bulb22, whitebulb, insertbulboff,"
  Lampz.Callback(22) = "DisableLighting2 p22off, .1, insert_dl_off_white,"
  Lampz.MassAssign(23)= l23
  Lampz.MassAssign(23)= l23a
  Lampz.Callback(23) = "DisableLighting2 p23, insert_dl_on_white, 1,"
  Lampz.Callback(23) = "DisableLighting2 bulb23, whitebulb, insertbulboff,"
  Lampz.Callback(23) = "DisableLighting2 p23off, .1, insert_dl_off_white,"
  Lampz.MassAssign(24) = l24
  Lampz.MassAssign(24) = l24a
  Lampz.Callback(24) = "DisableLighting p24, insert_dl_on_orange,"
  Lampz.MassAssign(25) = l25
  Lampz.MassAssign(25) = l25a
  Lampz.Callback(25) = "DisableLighting p25, insert_dl_on_orange,"
  Lampz.MassAssign(26) = l26
  Lampz.MassAssign(26) = l26a
  Lampz.Callback(26) = "DisableLighting p26, insert_dl_on_orange,"
  Lampz.MassAssign(27) = l27
  Lampz.MassAssign(27) = l27a
  Lampz.Callback(27) = "DisableLighting p27, insert_dl_on_orange,"
  Lampz.MassAssign(28) = l28
  Lampz.MassAssign(28) = l28a
  Lampz.Callback(28) = "DisableLighting p28, insert_dl_on_orange,"
  Lampz.MassAssign(29) = l29
  Lampz.MassAssign(29) = l29a
  Lampz.Callback(29) = "DisableLighting p29, insert_dl_on_orange,"
  Lampz.Callback(29) = "SLightRefl"
  Lampz.MassAssign(30) = l30
  Lampz.MassAssign(30) = l30a
  Lampz.Callback(30) = "DisableLighting p30, insert_dl_on_orange,"
  Lampz.Callback(30) = "TLightRefl"
  Lampz.MassAssign(31) = l31
  Lampz.MassAssign(31) = l31a
  Lampz.Callback(31) = "DisableLighting p31, insert_dl_on_orange,"
  Lampz.Callback(31) = "ALightRefl"
  Lampz.MassAssign(32) = l32
  Lampz.MassAssign(32) = l32a
  Lampz.Callback(32) = "DisableLighting p32, insert_dl_on_orange,"
  Lampz.Callback(32) = "RLightRefl"
  Lampz.MassAssign(33)= l33
  Lampz.MassAssign(33)= l33a
  Lampz.Callback(33) = "DisableLighting2 p33, insert_dl_on_white, 1,"
  Lampz.Callback(33) = "DisableLighting2 bulb33, whitebulb, insertbulboff,"
  Lampz.Callback(33) = "DisableLighting2 p33off, .1, insert_dl_off_white,"
  Lampz.MassAssign(34) = l34
  Lampz.MassAssign(34) = l34a
  Lampz.Callback(34) = "DisableLighting p34, insert_dl_on_green,"
  Lampz.MassAssign(35)= l35
  Lampz.MassAssign(35)= l35a
  Lampz.Callback(35) = "DisableLighting2 p35, insert_dl_on_bigyellow_flat, 1,"
  Lampz.Callback(35) = "DisableLighting2 bulb35, yellowbulb, insertbulboff,"
  Lampz.Callback(35) = "DisableLighting2 p35off, .1, insert_dl_off_yellow,"
  Lampz.MassAssign(36)= l36
  Lampz.MassAssign(36)= l36a
  Lampz.Callback(36) = "DisableLighting2 p36, insert_dl_on_pink, 1,"
  Lampz.Callback(36) = "DisableLighting2 bulb36, pinkbulb, insertbulboff,"
  Lampz.Callback(36) = "DisableLighting2 p36off, .1, insert_dl_off_pink,"
  Lampz.MassAssign(37) = l37
  Lampz.MassAssign(37) = l37a
  Lampz.Callback(37) = "DisableLighting p37, insert_dl_on_green_ship,"
  Lampz.MassAssign(38) = l38
  Lampz.MassAssign(38) = l38a
  Lampz.Callback(38) = "DisableLighting p38, insert_dl_on_green_ship,"
  Lampz.MassAssign(39) = l39
  Lampz.MassAssign(39) = l39a
  Lampz.Callback(39) = "DisableLighting p39, insert_dl_on_green_ship,"
  Lampz.MassAssign(40)= l40
  Lampz.MassAssign(40)= l40a
  Lampz.Callback(40) = "DisableLighting2 p40, insert_dl_on_white, 1,"
  Lampz.Callback(40) = "DisableLighting2 bulb40, whitebulb, insertbulboff,"
  Lampz.Callback(40) = "DisableLighting2 p40off, .1, insert_dl_off_white,"
  Lampz.MassAssign(41)= l41
  Lampz.MassAssign(41)= l41a
  Lampz.Callback(41) = "DisableLighting2 p41, insert_dl_on_white, 1,"
  Lampz.Callback(41) = "DisableLighting2 bulb41, whitebulb, insertbulboff,"
  Lampz.Callback(41) = "DisableLighting2 p41off, .1, insert_dl_off_white,"
  Lampz.MassAssign(42)= l42
  Lampz.MassAssign(42)= l42a
  Lampz.Callback(42) = "DisableLighting2 p42, insert_dl_on_white, 1,"
  Lampz.Callback(42) = "DisableLighting2 bulb42, whitebulb, insertbulboff,"
  Lampz.Callback(42) = "DisableLighting2 p42off, .1, insert_dl_off_white,"
  Lampz.MassAssign(43)= l43
  Lampz.MassAssign(43)= l43a
  Lampz.Callback(43) = "DisableLighting2 p43, insert_dl_on_bigyellow_flat, 1,"
  Lampz.Callback(43) = "DisableLighting2 bulb43, yellowbulb, insertbulboff,"
  Lampz.Callback(43) = "DisableLighting2 p43off, .1, insert_dl_off_yellow,"
  Lampz.MassAssign(44) = l44
  Lampz.MassAssign(44) = l44a
  Lampz.Callback(44) = "DisableLighting p44, insert_dl_on_blue,"
  Lampz.MassAssign(45) = l45
  Lampz.MassAssign(45) = l45a
  Lampz.Callback(45) = "DisableLighting p45, insert_dl_on_red,"
  Lampz.Callback(45) = "DisableLighting2 p45off, 1, insert_dl_off_red,"
  Lampz.MassAssign(46) = l46
  Lampz.MassAssign(46) = l46a
  Lampz.Callback(46) = "DisableLighting p46, insert_dl_on_red,"
  Lampz.Callback(46) = "DisableLighting2 p46off, 1, insert_dl_off_red,"
  Lampz.MassAssign(47)= l47
  Lampz.MassAssign(47)= l47a
  Lampz.Callback(47) = "DisableLighting2 p47, insert_dl_on_bigyellow_flat, 1,"
  Lampz.Callback(47) = "DisableLighting2 bulb47, yellowbulb, insertbulboff,"
  Lampz.Callback(47) = "DisableLighting2 p47off, .1, insert_dl_off_yellow,"



' GI Callbacks
  Lampz.Callback(140) = "GIUpdates"

  'Turn off all lamps on startup
  Lampz.Init  'This just turns state of any lamps to 1

  'Immediate update to turn on GI, turn off lamps
  Lampz.Update

End Sub


'====================
'Class jungle nf
'====================

'No-op object instead of adding more conditionals to the main loop
'It also prevents errors if empty lamp numbers are called, and it's only one object
'should be g2g?
Class NullFadingObject
  Public Property Let IntensityScale(input)

  End Property
End Class

'version 0.11 - Mass Assign, Changed modulate style
'version 0.12 - Update2 (single -1 timer update) update method for core.vbs
'Version 0.12a - Filter can now be accessed via 'FilterOut'
'Version 0.12b - Changed MassAssign from a sub to an indexed property (new syntax: lampfader.MassAssign(15) = Light1 )
'Version 0.13 - No longer requires setlocale. Callback() can be assigned multiple times per index
' Note: if using multiple 'LampFader' objects, set the 'name' variable to avoid conflicts with callbacks
'Version 0.14 - Updated to support modulated signals - Niwak
'Version 0.15 - Added IsLight property - apophis

Class LampFader
  Public IsLight(150)
  Public FadeSpeedDown(150), FadeSpeedUp(150)
  Private Lock(150), Loaded(150), OnOff(150)
  Public UseFunc
  Private cFilter
  Public UseCallback(150), cCallback(150)
  Public Lvl(150), Obj(150)
  Private Mult(150)
  Public FrameTime
  Private InitFrame
  Public Name

  Sub Class_Initialize()
    InitFrame = 0
    Dim x
    For x = 0 To UBound(OnOff)   'Set up fade speeds
      FadeSpeedDown(x) = 1 / 100  'fade speed down
      FadeSpeedUp(x) = 1 / 80   'Fade speed up
      UseFunc = False
      lvl(x) = 0
      OnOff(x) = 0
      Lock(x) = True
      Loaded(x) = False
      Mult(x) = 1
      IsLight(x) = False
    Next
    Name = "LampFaderNF" 'NEEDS TO BE CHANGED IF THERE'S MULTIPLE OF THESE OBJECTS, OTHERWISE CALLBACKS WILL INTERFERE WITH EACH OTHER!!
    For x = 0 To UBound(OnOff)     'clear out empty obj
      If IsEmpty(obj(x) ) Then Set Obj(x) = NullFader' : Loaded(x) = True
    Next
  End Sub

  Public Property Get Locked(idx)
    Locked = Lock(idx)
    '   debug.print Lampz.Locked(100) 'debug
  End Property

  Public Property Get state(idx)
    state = OnOff(idx)
  End Property

  Public Property Let Filter(String)
    Set cFilter = GetRef(String)
    UseFunc = True
  End Property

  Public Function FilterOut(aInput)
    If UseFunc Then
      FilterOut = cFilter(aInput)
    Else
      FilterOut = aInput
    End If
  End Function

  '   Public Property Let Callback(idx, String)
  '    cCallback(idx) = String
  '    UseCallBack(idx) = True
  '   End Property

  Public Property Let Callback(idx, String)
    UseCallBack(idx) = True
    '   cCallback(idx) = String 'old execute method

    'New method: build wrapper subs using ExecuteGlobal, then call them
    cCallback(idx) = cCallback(idx) & "___" & String  'multiple strings dilineated by 3x _

    Dim tmp
    tmp = Split(cCallback(idx), "___")

    Dim str, x
    For x = 0 To UBound(tmp)  'build proc contents
      'If Not tmp(x)="" then str = str & "  " & tmp(x) & " aLVL" & "  '" & x & vbnewline  'more verbose
      If Not tmp(x) = "" Then str = str & tmp(x) & " aLVL:"
    Next
    '   msgbox "Sub " & name & idx & "(aLvl):" & str & "End Sub"
    Dim out
    out = "Sub " & name & idx & "(aLvl):" & str & "End Sub"
    ExecuteGlobal Out
  End Property

  Public Property Let state(ByVal idx, input) 'Major update path
    If TypeName(input) <> "Double" And TypeName(input) <> "Integer"  And TypeName(input) <> "Long" Then
      If input Then
        input = 1
      Else
        input = 0
      End If
    End If
    If Input <> OnOff(idx) Then  'discard redundant updates
      OnOff(idx) = input
      Lock(idx) = False
      Loaded(idx) = False
    End If
  End Property

  'Sub MassAssign(aIdx, aInput)
  Public Property Let MassAssign(aIdx, aInput) 'Mass assign, Builds arrays where necessary
    If TypeName(obj(aIdx)) = "NullFadingObject" Then 'if empty, use Set
      If IsArray(aInput) Then
        obj(aIdx) = aInput
      Else
        Set obj(aIdx) = aInput
        If TypeName(aInput) = "Light" Then
          IsLight(aIdx) = True
        End If
      End If
    Else
      Obj(aIdx) = AppendArray(obj(aIdx), aInput)
    End If
  End Property

  Sub SetLamp(aIdx, aOn)  'If obj contains any light objects, set their states to 1 (Fading is our job!)
    state(aIdx) = aOn
  End Sub

  Public Sub TurnOnStates() 'turn state to 1
    Dim debugstr
    Dim idx
    For idx = 0 To UBound(obj)
      If IsArray(obj(idx)) Then
        'debugstr = debugstr & "array found at " & idx & "..."
        Dim x, tmp
        tmp = obj(idx) 'set tmp to array in order to access it
        For x = 0 To UBound(tmp)
          If TypeName(tmp(x)) = "Light" Then DisableState tmp(x)' : debugstr = debugstr & tmp(x).name & " state'd" & vbnewline
          tmp(x).intensityscale = 0.001 ' this can prevent init stuttering
        Next
      Else
        If TypeName(obj(idx)) = "Light" Then DisableState obj(idx)' : debugstr = debugstr & obj(idx).name & " state'd (not array)" & vbnewline
        obj(idx).intensityscale = 0.001 ' this can prevent init stuttering
      End If
    Next
    '   debug.print debugstr
  End Sub

  Private Sub DisableState(ByRef aObj)
    aObj.FadeSpeedUp = 1000
    aObj.State = 1
  End Sub

  Public Sub Init() 'Just runs TurnOnStates right now
    TurnOnStates
  End Sub

  Public Property Let Modulate(aIdx, aCoef)
    Mult(aIdx) = aCoef
    Lock(aIdx) = False
    Loaded(aIdx) = False
  End Property
  Public Property Get Modulate(aIdx)
    Modulate = Mult(aIdx)
  End Property

  Public Sub Update1() 'Handle all boolean numeric fading. If done fading, Lock(x) = True. Update on a '1' interval Timer!
    Dim x
    For x = 0 To UBound(OnOff)
      If Not Lock(x) Then 'and not Loaded(x) then
        If OnOff(x) > 0 Then 'Fade Up
          Lvl(x) = Lvl(x) + FadeSpeedUp(x)
          If Lvl(x) >= OnOff(x) Then
            Lvl(x) = OnOff(x)
            Lock(x) = True
          End If
        Else 'fade down
          Lvl(x) = Lvl(x) - FadeSpeedDown(x)
          If Lvl(x) <= 0 Then
            Lvl(x) = 0
            Lock(x) = True
          End If
        End If
      End If
    Next
  End Sub

  Public Sub Update2() 'Both updates on -1 timer (Lowest latency, but less accurate fading at 60fps vsync)
    FrameTime = GameTime - InitFrame
    InitFrame = GameTime  'Calculate frametime
    Dim x
    For x = 0 To UBound(OnOff)
      If Not Lock(x) Then 'and not Loaded(x) then
        If OnOff(x) > 0 Then 'Fade Up
          Lvl(x) = Lvl(x) + FadeSpeedUp(x) * FrameTime
          If Lvl(x) >= OnOff(x) Then
            Lvl(x) = OnOff(x)
            Lock(x) = True
          End If
        Else 'fade down
          Lvl(x) = Lvl(x) - FadeSpeedDown(x) * FrameTime
          If Lvl(x) <= 0 Then
            Lvl(x) = 0
            Lock(x) = True
          End If
        End If
      End If
    Next
    Update
  End Sub

  Public Sub Update() 'Handle object updates. Update on a -1 Timer! If done fading, loaded(x) = True
    Dim x,xx, aLvl
    For x = 0 To UBound(OnOff)
      If Not Loaded(x) Then
        aLvl = Lvl(x) * Mult(x)
        If IsArray(obj(x) ) Then  'if array
          If UseFunc Then
            For Each xx In obj(x)
              xx.IntensityScale = cFilter(aLvl)
            Next
          Else
            For Each xx In obj(x)
              xx.IntensityScale = aLvl
            Next
          End If
        Else            'if single lamp or flasher
          If UseFunc Then
            obj(x).Intensityscale = cFilter(aLvl)
          Else
            obj(x).Intensityscale = aLvl
          End If
        End If
        '   if TypeName(lvl(x)) <> "Double" and typename(lvl(x)) <> "Integer" and typename(lvl(x)) <> "Long" then msgbox "uhh " & 2 & " = " & lvl(x)
        '   If UseCallBack(x) then execute cCallback(x) & " " & (Lvl(x))  'Callback
        If UseCallBack(x) Then Proc name & x,aLvl 'Proc
        If Lock(x) Then
          If Lvl(x) = OnOff(x) Or Lvl(x) = 0 Then Loaded(x) = True  'finished fading
        End If
      End If
    Next
  End Sub
End Class

'Lamp Filter
Function LampFilter(aLvl)
  LampFilter = aLvl ^ 1.6 'exponential curve?
End Function

'Helper functions
Sub Proc(String, Callback)  'proc using a string and one argument
  'On Error Resume Next
  Dim p
  Set P = GetRef(String)
  P Callback
  If err.number = 13 Then  MsgBox "Proc error! No such procedure: " & vbNewLine & String
  If err.number = 424 Then MsgBox "Proc error! No such Object"
End Sub

Function AppendArray(ByVal aArray, aInput)  'append one value, object, or Array onto the End of a 1 dimensional array
  If IsArray(aInput) Then 'Input is an array...
    Dim tmp
    tmp = aArray
    If Not IsArray(aArray) Then 'if not array, create an array
      tmp = aInput
    Else          'Append existing array with aInput array
      ReDim Preserve tmp(UBound(aArray) + UBound(aInput) + 1) 'If existing array, increase bounds by uBound of incoming array
      Dim x
      For x = 0 To UBound(aInput)
        If IsObject(aInput(x)) Then
          Set tmp(x + UBound(aArray) + 1 ) = aInput(x)
        Else
          tmp(x + UBound(aArray) + 1 ) = aInput(x)
        End If
      Next
      AppendArray = tmp  'return new array
    End If
  Else 'Input is NOT an array...
    If Not IsArray(aArray) Then 'if not array, create an array
      aArray = Array(aArray, aInput)
    Else
      ReDim Preserve aArray(UBound(aArray) + 1) 'If array, increase bounds by 1
      If IsObject(aInput) Then
        Set aArray(UBound(aArray)) = aInput
      Else
        aArray(UBound(aArray)) = aInput
      End If
    End If
    AppendArray = aArray 'return new array
  End If
End Function

'******************************************************
'****  END LAMPZ
'******************************************************



'******************************************************
'****  Ball GI Brightness Level Code
'******************************************************

const BallBrightMax = 255     'Brightness setting when GI is on (max of 255). Only applies for Normal ball.
const BallBrightMin = 100     'Brightness setting when GI is off (don't set above the max). Only applies for Normal ball.

Sub UpdateBallBrightness
  Dim b, brightness
  For b = 0 to UBound(gBOT)
    if ballbrightness >=0 Then gBOT(b).color = ballbrightness + (ballbrightness * 256) + (ballbrightness * 256 * 256)
      if b = UBound(gBOT) then 'until last ball brightness is set, then reset to -1
      if ballbrightness = ballbrightMax Or ballbrightness = ballbrightMin then ballbrightness = -1
    end if
  Next
End Sub


'*******************************************
' GI Routines
'*******************************************

dim ballbrightness
dim gilvl:gilvl = 1


Sub PFGI(enabled)
  if enabled Then
    Lampz.SetLamp 140, 1
    Sound_GI_Relay 1, Relay_GI
  Else
    Lampz.SetLamp 140, 0
    Sound_GI_Relay 0, Relay_GI
  End If
End Sub



Sub GIUpdates(ByVal aLvl)

  dim x, girubbercolor, girubbercolor2, girubbercolor3

  if Lampz.UseFunc then aLvl = LampFilter(aLvl) 'Callbacks don't get this filter automatically

  if aLvl = 0 then                    'GI OFF, let's hide ON prims
    if ballbrightness <> -1 then ballbrightness = ballbrightMin
    if GameInPlayFirstTime = 0 Then
      VR_Backglass_Prim.image = "VR_Backglass_Dark"
      VR_Backglass_Prim.blenddisablelighting = 3
      for each x in VRBackglassText: x.blenddisablelighting = .25: Next
    end If
  Elseif aLvl = 1 then                  'GI ON, let's hide OFF prims
    if ballbrightness <> -1 then ballbrightness = ballbrightMax
    VR_Backglass_Prim.image = "VR_Backglass_Lit"
    VR_Backglass_Prim.blenddisablelighting = 3
    for each x in VRBackglassText: x.blenddisablelighting = 3: Next
  Else
    if gilvl = 0 Then               'GI has just changed from OFF to fading, let's show ON
      ballbrightness = ballbrightMin + 1
    elseif gilvl = 1 Then             'GI has just changed from ON to fading, let's show OFF
      ballbrightness = ballbrightMax - 1
    Else
      'no change
    end if
    if GameInPlayFirstTime = 0 Then
      VR_Backglass_Prim.image = "VR_Backglass_Lit"
      VR_Backglass_Prim.blenddisablelighting = 2 + aLvl
      for each x in VRBackglassText: x.blenddisablelighting = 2 + aLvl: Next
    end If
  end if

  dim mm,bp, bpl

' GI Lights
  For each x in GI:x.Intensityscale = aLvl: Next
  For each x in GIBulbs
    x.blenddisablelighting = 1 * alvl
  Next

' Flippers
  For each x in GIFlippers
    x.blenddisablelighting = 0.2 * alvl + .1
  Next

' Standup Targets
  For each x in GITargets
    x.blenddisablelighting = 0.15 * alvl + .1
  Next

' prims (Pegs)
  For each x in GIPegs
    x.blenddisablelighting = .005 * alvl + .005
  Next
  PegPlasticT8Twin2.blenddisablelighting = .05 * alvl + .05
  PegPlasticT8Twin1.blenddisablelighting = .05 * alvl + .05

' prims (Top Lane Rollovers)
  For each x in GILaneRollovers
    x.blenddisablelighting = .05 * alvl + .025
  Next

  For each x in GITents
    x.blenddisablelighting = .075 * alvl + .025
  Next

' prims (Cutouts)
  For each x in GICutouts
    x.blenddisablelighting = .1 * alvl + .15
  Next

' plastics
  OuterPrimBlades_DT.blenddisablelighting = 0.095 * alvl + .005
  for each x in GIPlastics
    x.blenddisablelighting = 0.025 - (0.02*aLvl)
  Next

' rubbers
  girubbercolor =  63*alvl + 192
  girubbercolor2 = 63*alvl + 192
  girubbercolor3 = 66*alvl + 174
  MaterialColor "Rubber White",RGB(girubbercolor,girubbercolor2,girubbercolor3)

' prims (apron, plunger cover, instruction cards, and spinner
  ApronPrim.blenddisablelighting = 0.045 * alvl + .015
  PlungerCoverPrim.blenddisablelighting = 0.045 * alvl + .015
  MaterialColor "Plastic with an image Cards",RGB(girubbercolor,girubbercolor,girubbercolor)
  SpinnerPrim.blenddisablelighting = 0.05 * alvl + 0.05

'saucer
  Pkickerhole1.blenddisablelighting = 0.125 * alvl + .025
  PkickarmSW4.blenddisablelighting = 0.125 * alvl + .025
  pSaucerFloor.blenddisablelighting = 0.25 * alvl + .25

' Metals
  For each x in GIMetals
    x.blenddisablelighting = 0.01 * alvl + 0.005
  Next

  For each x in GIScrews
    x.blenddisablelighting = .05 * alvl + .005
  Next

' Gate Prim
  p_gate3.blenddisablelighting = .04 * alvl + .01
  pScoreSpring.blenddisablelighting = .05 * alvl + .025

'GI PF bakes
  FlasherGI.opacity = 500 * alvl

'Bumper Lights

  FlBumperFadeTarget(1) = aLvl*0.5 +0.3
  FlBumperFadeTarget(2) = aLvl*0.5 +0.3
  FlBumperFadeTarget(3) = aLvl*0.5 +0.3

'Center Red Insert
  lcenter.intensity = 2 * aLvl
  lcentera.intensity = 3 * aLvl
  pcenter.blenddisablelighting = insert_dl_on_red * aLvl
  pcenteroff.blenddisablelighting = 4 - (insert_dl_off_red-1) * aLvl

' ball
  if ballbrightness <> ballbrightMax Or ballbrightness <> ballbrightMin Or ballbrightness <> -1 then ballbrightness = INT(alvl * (ballbrightMax - ballbrightMin) + ballbrightMin)

  gilvl = alvl

  if vr_room = 0 and cab_mode = 0 Then
    if aLvl = 0 then
      L_DT_Top_Left.visible = 0
      L_DT_Top_Right.visible = 0
      for each x in DT_Text
        x.visible = 1
      Next
    Else
      L_DT_Top_Left.visible = 1
      L_DT_Top_Right.visible = 1
      for each x in DT_Text
        x.visible = 1
      Next
    End If
  Else
      L_DT_Top_Left.visible = 0
      L_DT_Top_Right.visible = 0
      for each x in DT_Text
        x.visible = 0
      Next
  End If

End Sub


'*******************************************
' S-T-A-R Reflections
'*******************************************

Sub SLightRefl(ByVal aLvl)
  if Lampz.UseFunc then aLvl = LampFilter(aLvl) 'Callbacks don't get this filter automatically
  if StarReflections = 1 Then
    FreflectS.opacity = aLvl * 200
  Else
    FreflectS.opacity = 0
  End If
End Sub

Sub TLightRefl(ByVal aLvl)
  if Lampz.UseFunc then aLvl = LampFilter(aLvl) 'Callbacks don't get this filter automatically
  if StarReflections = 1 Then
    FreflectT.opacity = aLvl * 200
  Else
    FreflectT.opacity = 0
  End If
End Sub

Sub ALightRefl(ByVal aLvl)
  if Lampz.UseFunc then aLvl = LampFilter(aLvl) 'Callbacks don't get this filter automatically
  if StarReflections = 1 Then
    FreflectA.opacity = aLvl * 200
  Else
    FreflectA.opacity = 0
  End If
End Sub

Sub RLightRefl(ByVal aLvl)
  if Lampz.UseFunc then aLvl = LampFilter(aLvl) 'Callbacks don't get this filter automatically
  if StarReflections = 1 Then
    FreflectR.opacity = aLvl * 200
  Else
    FreflectR.opacity = 0
  End If
End Sub


'******************************************************
' Desktop Backglass and Text Lights
'******************************************************

If DesktopMode = True Then
  Set LampCallback=GetRef("UpdateDTLamps")
End If


If VR_ROOM = 1 Then
  Set LampCallback=GetRef("UpdateVRLamps")
End If

If cab_mode = 1 Then
  Set LampCallback=GetRef("UpdateCabLamps")
End If


Sub UpdateDTLamps()

' Ball Release Routine
    If Controller.Lamp(12) = True And BallRelease.BallCntOver = 1 Then 'Kick ball to plunger lane
    BallRelease.kick 60, 12
    RandomSoundBallRelease BallRelease
  End If

    If Controller.Lamp(11) = False Then 'Game Over triggers match and BIP Then
    If GameInPlay = 1 Then
      BIPReel.setValue(1)
    End If
      MatchReel.setValue(0)
      GameOverReel.setValue(0)
    Else
      BIPReel.setValue(0)
      GameOverReel.setValue(1)

    If GameInPlay = 0 and GameInPlayFirstTime = 1 and NTM = 1 Then
      MatchReel.setValue(1)
    End If
    End If

  If Controller.Lamp(10) = False Then: HighScoreReel.setValue(0):     Else: HighScoreReel.setValue(1) 'High Score

  If GameInPlayFirstTime = 1 Then
    If Controller.Lamp(3)  = False Then: ShootAgainReel.setValue(0):  Else: ShootAgainReel.setValue(1) 'Shoot Again
    If Controller.Lamp(8)  = False Then: TiltReel.setValue(0):      Else: TiltReel.setValue(1)
  End If

End Sub

Sub UpdateVRLamps()

' Ball Release Routine
    If Controller.Lamp(12) = true And BallRelease.BallCntOver = 1 Then 'Kick ball to plunger lane
    BallRelease.kick 60, 12
    RandomSoundBallRelease BallRelease
  End If

  If Controller.Lamp(10) = 0 Then: VR_BG_HSTD.visible=0:  Else: VR_BG_HSTD.visible=1

  If Controller.Lamp(11) = 0 Then VR_BG_GO.visible=0:   Else: VR_BG_GO.visible=1

  If GameInPlayFirstTime = 1 Then
    If Controller.Lamp(3) = False Then: VR_BG_SA.visible=0:   Else: VR_BG_SA.visible=1
    If Controller.Lamp(8) = False Then: VR_BG_TILT.visible=0: Else: VR_BG_TILT.visible=1
  End If

End Sub

Sub UpdateCabLamps()

' Ball Release Routine
    If Controller.Lamp(12) = True And BallRelease.BallCntOver = 1 Then 'Kick ball to plunger lane
    BallRelease.kick 60, 12
    RandomSoundBallRelease BallRelease
  End If

End Sub


'*******************************************
' Hybrid code for VR, Cab, and Desktop
'*******************************************

Dim VRThings

if VR_Room = 0 and cab_mode = 0 Then
  for each VRThings in VRStuff:VRThings.visible = 0:Next
  for each VRThings in VRClock:VRThings.visible = 0:Next
  for each VRThings in VRBackglassLEDs:VRThings.visible = 0:Next
  for each VRThings in VRBackglassFlash:VRThings.visible = 0:Next
  for each VRThings in DTBGLights:VRThings.visible = 1: Next
  for each VRThings in DTRails:VRThings.visible = 1:Next
  OuterPrimBlades_dt.visible = 1
  OuterPrimBlades_cab.visible = 0
  OuterPrimBlades_VR.visible = 0
  L_DT_Top_Left.State = 1
  L_DT_Top_Right.State = 1

Elseif VR_Room = 0 and cab_mode = 1 Then
  for each VRThings in VRStuff:VRThings.visible = 0:Next
  for each VRThings in VRClock:VRThings.visible = 0:Next
  for each VRThings in VRBackglassLEDs:VRThings.visible = 0:Next
  for each VRThings in VRBackglassFlash:VRThings.visible = 0:Next
  for each VRThings in DTBackglass:VRThings.visible = 0: Next
  for each VRThings in DTBGLights:VRThings.visible = 0: Next
  for each VRThings in DTRails:VRThings.visible = 0:Next
  OuterPrimBlades_dt.visible = 0
  OuterPrimBlades_cab.visible = 1
  OuterPrimBlades_VR.visible = 0
  L_DT_Top_Left.State = 0
  L_DT_Top_Right.State = 0

Else
  for each VRThings in VRStuff:VRThings.visible = 1:Next
  for each VRThings in VRClock:VRThings.visible = WallClock:Next
  for each VRThings in DTBackglass:VRThings.visible = 0: Next
  for each VRThings in DTBGLights:VRThings.visible = 0: Next
  for each VRThings in DTRails:VRThings.visible = 0:Next
  for each VRThings in VRBackglassFlash:VRThings.blenddisablelighting = 3:Next
  OuterPrimBlades_dt.visible = 0
  OuterPrimBlades_cab.visible = 0
  OuterPrimBlades_VR.visible = 1
  setup_backglass()
  L_DT_Top_Left.State = 0
  L_DT_Top_Right.State = 0
End If



'*******************************************
' VR Clock
'*******************************************

Dim CurrentMinute ' for VR clock

Sub ClockTimer_Timer()

    'ClockHands Below *********************************************************
  Pminutes.RotAndTra2 = (Minute(Now())+(Second(Now())/100))*6
  Phours.RotAndTra2 = Hour(Now())*30+(Minute(Now())/2)
  Pseconds.RotAndTra2 = (Second(Now()))*6
  CurrentMinute=Minute(Now())

End Sub

'*******************************************
' VR Plunger Code
'*******************************************

Sub VR_Plunger_Timer
  If VR_Primary_plunger.transy > -135 then
      VR_Primary_plunger.transy = VR_Primary_plunger.transy - 5
  End If
End Sub

Sub VR_Plunger2_Timer
  VR_Primary_plunger.transy = (-5 * Plunger.Position) + 20
End Sub


'*******************************************
' Options
'*******************************************

Dim TC, SC

Sub SetupOptions

'  ROM Volume Settings
    With Controller
        .Games(cGameName).Settings.Value("volume") = ROMVolume
  End With

'  Tourney Mode Settings

If Tourney = 1 Then
  Ttents = 1
  TFlipper = 2
  Tscrewcap = 4
  WallColor = 1
Else
  Ttents = LaneTentColor
  TFlipper = FlipperColor
  Tscrewcap = ScrewCapColor
  WallColor = SideWallColor
End If


' Plastic Edge Color
  If PlasticColor = 6 Then
    PlasProt = RndInt(0,5)
  Else
    PlasProt = PlasticColor
  End if

  SetAcrylicColor PlasProt

' Side Rails
  If SideRailGotlieb = 1 Then
    Ramp16.image = "gtb80-rail-left-alum"
    Ramp15.image = "gtb80-rail-right-alum"
  Else
    Ramp16.image = "aluminum"
    Ramp15.image = "aluminum"
  End If

' Lane Tent Color
  If Ttents = 1 Then
    For each TC in GITents
      TC.material = "Plastic Yellow"
    Next
    For each TC in GILaneRollovers
      TC.material = "Plastic Yellow"
    Next
  Else
    For each TC in GITents
      TC.material = "Plastic Red"
    Next
    For each TC in GILaneRollovers
      TC.material = "Plastic Red"
    Next
  End If

' Flipper Color
  If TFlipper = 1 Then
    rflipr.material = "Rubber Yellow1"
    lflipr.material = "Rubber Yellow2"
  Elseif TFlipper = 2 Then
    rflipr.material = "Rubber Black1"
    lflipr.material = "Rubber Black2"
  Elseif TFlipper = 3 Then
    rflipr.material = "Rubber Orange1"
    lflipr.material = "Rubber Orange2"
  Else
    rflipr.material = "Red Rubber1"
    lflipr.material = "Red Rubber2"
  End If

' Screw Cap Color
  If Tscrewcap  = 1  Then
    For each SC in GIScrews
      SC.material = "Metal Lamp Black"
    Next
  Elseif Tscrewcap  = 2 Then
    For each SC in GIScrews
      SC.material = "Metal Lamp Bronze"
    Next
  Elseif Tscrewcap  = 3 Then
    For each SC in GIScrews
      SC.material = "Metal Lamp Silver"
    Next
  Elseif Tscrewcap  = 4 Then
    For each SC in GIScrews
      SC.material = "Metal Lamp BlueGreen"
    Next
  Elseif Tscrewcap  = 5 Then
    For each SC in GIScrews
      SC.material = "Metal Lamp Gold"
    Next
  Else
    For each SC in GIScrews
      SC.material = "Metal Lamp White"
    Next
  End If

' Side Wall Color
  If WallColor = 1 Then
    For each SC in SideWalls
      SC.image = "sidewood_h-dark"
      SC.sideimage = "sidewood_h-dark"
    Next
  Else
    For each SC in SideWalls
      SC.image = "sidewood_h"
      SC.sideimage = "sidewood_h"
    Next
  End If

if VR_Room = 1 Then
'Custom Walls, Floor, and Roof
  if CustomWalls = 1 Then
    VR_Wall_Left.image = "VR_Wall_Left"
    VR_Wall_Right.image = "VR_Wall_Right"
    VR_Floor.image = "VR_Floor"
    VR_Roof.image = "VR_Roof"
  Elseif CustomWalls = 2 Then
    VR_Wall_Left.image = "VR_Wall_Left_DV2"
    VR_Wall_Right.image = "VR_Wall_Right_DV2"
    VR_Floor.image = "VR_Room Floor"
    VR_Roof.image = "VR_Room Roof 2"
  Elseif CustomWalls = 3 Then
    VR_Wall_Left.image = "VR_Wall_Left_DV1"
    VR_Wall_Right.image = "VR_Wall_Right_DV1"
    VR_Floor.image = "VR_Room Floor"
    VR_Roof.image = "VR_Room Roof 2"
  Elseif CustomWalls = 4 Then
    VR_Wall_Left.image = "VR_Wall_Left_DV3"
    VR_Wall_Right.image = "VR_Wall_Right_DV3"
    VR_Floor.image = "VR_Room Floor"
    VR_Roof.image = "VR_Room Roof 2"
  Else
    VR_Wall_Left.image = "wallpaper left"
    VR_Wall_Right.image = "wallpaper_right"
    VR_Floor.image = "FloorCompleteMap2"
    VR_Roof.image = "Light Gray"
  end if

  If poster = 1 Then
    VRposter.visible = 1
  Else
    VRposter.visible = 0
  End If

  If poster2 = 1 Then
    VRposter2.visible = 1
  Else
    VRposter2.visible = 0
  End If
End If

End Sub


'*******************************************
'  Ball brightness code
'*******************************************

If ChangeBall = 1 Then
  if BallLightness = 0 Then
    table1.BallImage="ball-dark"
    table1.BallFrontDecal="JPBall-Scratches"
  elseif BallLightness = 1 Then
    table1.BallImage="ball_HDR"
    table1.BallFrontDecal="Scratches"
  elseif BallLightness = 2 Then
    table1.BallImage="ball-light-hf"
    table1.BallFrontDecal="scratchedmorelight"
  else
    table1.BallImage="ball-lighter-hf"
    table1.BallFrontDecal="scratchedmorelight"
  End If
End If


'*******************************
'Acrylic plastics - iaakki acrylic redbone
'*******************************


Sub SetMaterialColor(name, new_color)
    Dim wrapLighting, roughness, glossyImageLerp, thickness, edge, edgeAlpha, opacity, base, glossy, clearcoat, isMetal, opacityActive, elasticity, elasticityFalloff, friction, scatterAngle
    GetMaterial name, wrapLighting, roughness, glossyImageLerp, thickness, edge, edgeAlpha, opacity, base, glossy, clearcoat, isMetal, opacityActive, elasticity, elasticityFalloff, friction, scatterAngle
    base = new_color
    UpdateMaterial name, wrapLighting, roughness, glossyImageLerp, thickness, edge, edgeAlpha, opacity, base, glossy, clearcoat, isMetal, opacityActive, elasticity, elasticityFalloff, friction, scatterAngle
End Sub

Sub SetMaterialGlossyColor(name, new_color)
    Dim wrapLighting, roughness, glossyImageLerp, thickness, edge, edgeAlpha, opacity, base, glossy, clearcoat, isMetal, opacityActive, elasticity, elasticityFalloff, friction, scatterAngle
    GetMaterial name, wrapLighting, roughness, glossyImageLerp, thickness, edge, edgeAlpha, opacity, base, glossy, clearcoat, isMetal, opacityActive, elasticity, elasticityFalloff, friction, scatterAngle
    glossy = new_color
    UpdateMaterial name, wrapLighting, roughness, glossyImageLerp, thickness, edge, edgeAlpha, opacity, base, glossy, clearcoat, isMetal, opacityActive, elasticity, elasticityFalloff, friction, scatterAngle
End Sub

Sub SetMaterialCKColor(name, new_color)
    Dim wrapLighting, roughness, glossyImageLerp, thickness, edge, edgeAlpha, opacity, base, glossy, clearcoat, isMetal, opacityActive, elasticity, elasticityFalloff, friction, scatterAngle
    GetMaterial name, wrapLighting, roughness, glossyImageLerp, thickness, edge, edgeAlpha, opacity, base, glossy, clearcoat, isMetal, opacityActive, elasticity, elasticityFalloff, friction, scatterAngle
    clearcoat = new_color
    UpdateMaterial name, wrapLighting, roughness, glossyImageLerp, thickness, edge, edgeAlpha, opacity, base, glossy, clearcoat, isMetal, opacityActive, elasticity, elasticityFalloff, friction, scatterAngle
End Sub

DIM ProtCol
sub SetAcrylicColor(ColorOption)

  for each Protcol in Acrylics:Protcol.visible = 1:Next
  for each Protcol in Acrylics:Protcol.SideVisible = 1:Next

  if Not ColorOption = 5 then
    RGBAcrylTimer.enabled = false
    for each Protcol in Acrylics:Protcol.blenddisablelighting = 1:Next
    for each Protcol in Acrylics:Protcol.Topmaterial = "AcrylicBLTops":Next
    for each Protcol in Acrylics:Protcol.Sidematerial = "AcrylicBLSides":Next
    for each Protcol in Acrylics:Protcol.image = "Colour_Trans_Grey":Next
    for each Protcol in Acrylics:Protcol.Sideimage = "Colour_Cream_Trans":Next
  End If

  If ColorOption = 0 Then    'clear
    for each Protcol in Acrylics:Protcol.blenddisablelighting = 0:Next
    for each Protcol in Acrylics:Protcol.Topmaterial = "acrylic top":Next
    for each Protcol in Acrylics:Protcol.Sidematerial = "acrylic side":Next
    for each Protcol in Acrylics:Protcol.image = "":Next
    for each Protcol in Acrylics:Protcol.Sideimage = "":Next
    End If

  If ColorOption = 1 Then    'orangered
        SetMaterialColor "AcrylicBLSides",GIlvlRGB(RGB(255,8,0))
        SetMaterialColor "AcrylicBLTops",GIlvlRGB(RGB(0,0,0))
        SetMaterialGlossyColor "AcrylicBLSides",GIlvlRGB(RGB(255,8,0))
        SetMaterialGlossyColor "AcrylicBLTops",GIlvlRGB(RGB(0,0,0))
        SetMaterialCKColor "AcrylicBLSides",GIlvlRGB(RGB(255,8,0))
        SetMaterialCKColor "AcrylicBLTops",GIlvlRGB(RGB(0,0,0))
    End If

  If ColorOption = 2 Then    'yellow
        SetMaterialColor "AcrylicBLSides",GIlvlRGB(RGB(255,128,0))
        SetMaterialColor "AcrylicBLTops",GIlvlRGB(RGB(0,0,0))
        SetMaterialGlossyColor "AcrylicBLSides",GIlvlRGB(RGB(255,128,0))
        SetMaterialGlossyColor "AcrylicBLTops",GIlvlRGB(RGB(0,0,0))
        SetMaterialCKColor "AcrylicBLSides",GIlvlRGB(RGB(255,128,0))
        SetMaterialCKColor "AcrylicBLTops",GIlvlRGB(RGB(0,0,0))
    End If

  If ColorOption = 3 Then    'blue
        SetMaterialColor "AcrylicBLSides",GIlvlRGB(RGB(15,65,100))
        SetMaterialColor "AcrylicBLTops",GIlvlRGB(RGB(0,0,0))
        SetMaterialGlossyColor "AcrylicBLSides",GIlvlRGB(RGB(15,65,100))
        SetMaterialGlossyColor "AcrylicBLTops",GIlvlRGB(RGB(0,0,0))
        SetMaterialCKColor "AcrylicBLSides",GIlvlRGB(RGB(15,65,100))
        SetMaterialCKColor "AcrylicBLTops",GIlvlRGB(RGB(0,0,0))
    End If

  If ColorOption = 4 Then 'user defined
    dim currentColor : currentColor = GIlvlRGB(UserAcrylColor)
    SetMaterialColor "AcrylicBLSides",GIlvlRGB(UserAcrylColor)
    SetMaterialColor "AcrylicBLTops",GIlvlRGB(RGB(0,0,0))
        SetMaterialGlossyColor "AcrylicBLSides",GIlvlRGB(UserAcrylColor)
        SetMaterialGlossyColor "AcrylicBLTops",GIlvlRGB(RGB(0,0,0))
    SetMaterialCKColor "AcrylicBLSides",GIlvlRGB(UserAcrylColor)
    SetMaterialCKColor "AcrylicBLTops",GIlvlRGB(RGB(0,0,0))
  end if

  If ColorOption = 5 Then 'user defined
    RGBAcrylTimer.enabled = true
  end if

end sub


Function GIlvlRGB(color)
  Dim red, green, blue, dimFactor, dimEquation

  dimFactor = 5
  dimEquation = ((1+((dimFactor-1)*gilvl))/dimFactor)

  red = color And &HFF
  green = (color \ &H100) And &HFF
  blue = (color \ &H10000) And &HFF

  GIlvlRGB = RGB(red*dimEquation, green*dimEquation, blue*dimEquation)
End Function



dim RGBAcrylHueAngle : RGBAcrylHueAngle = 0
sub RGBAcrylTimer_timer
  RGBAcrylHueAngle = RGBAcrylHueAngle + 1
  if RGBAcrylHueAngle > 359 then RGBAcrylHueAngle = 0

  RGBAcrylHue RGBAcrylHueAngle

  SetMaterialColor "AcrylicBLSides",RGBAcrylColor
  SetMaterialColor "AcrylicBLTops",GIlvlRGB(RGB(0,0,0))
  SetMaterialGlossyColor "AcrylicBLSides",RGBAcrylColor
  SetMaterialGlossyColor "AcrylicBLTops",GIlvlRGB(RGB(0,0,0))
  SetMaterialCKColor "AcrylicBLSides",RGBAcrylColor
  SetMaterialCKColor "AcrylicBLTops",GIlvlRGB(RGB(0,0,0))
end sub

dim RGBAcrylColor: RGBAcrylColor = RGB(240,20,20)

Sub RGBAcrylHue(HueDegrees)
    dim ColorR, ColorG, ColorB
    dim outR, outG, outB

    'UW's from Hue angle
    dim U : U = cos(HueDegrees * PI/180)
    dim W : W = sin(HueDegrees * PI/180)
    'Initial color
    ColorR = 240
    ColorG = 20
    ColorB = 20
    'RGB values from Hue
    outR = Cint((.299+.701*U+.168*W)*ColorR + (.587-.587*U+.330*W)*ColorG + (.114-.114*U-.497*W) * ColorB)
    outG = Cint((.299-.299*U-.328*W)*ColorR + (.587+.413*U+.035*W)*ColorG + (.114-.114*U+.292*W) * ColorB)
    outB = Cint((.299-.300*U+1.25*W)*ColorR + (.587-.588*U-1.05*W)*ColorG + (.114+.886*U-.203*W) * ColorB)

    'limits as equations above may not be that precise
    if outR > 255 then outR = 255
    if outG > 255 then outG = 255
    if outB > 255 then outB = 255

    if outR < 0 then outR = 0
    if outG < 0 then outG = 0
    if outB < 0 then outB = 0

    RGBAcrylColor = GIlvlRGB(RGB(outR, outG, outB))
end sub

'*******************************
'end plastics REDBONE
'*******************************



'********************************************************************************************************************************
'******** Work done by UnclePaulie on Hybrid version 0.01 - 2.0 *********
'     Added latest VPW standards, physics, GI baking, Lampz and 3D inserts, shadows, updated backglass for desktop,
'   roth standup targets, standalone compatibility, sling corrections,
'   optional DIP adjustments, updated flipper physics, VR and fully hybrid, Flupper style bumpers and lighting.
'   Full details at the bottom of the script.  Redbone did the playfield, plastics, and other graphics.
'   Original scans, image assets, cab artwork stencils all used with permission from Borgdog.
'********************************************************************************************************************************

'v0.01  Started with UnclePaulie's Super Orbit table for all physics, materials, VR structure
'   Redbone created several images including the apron, target, instruction cards, plastics, and playfield.
'   Several assets from Borgdog's table; https://www.youtube.com/watch?v=NwFih3zu-4Y
'   Added all elements to table.
'   Got the trough working for two balls and gBOT.
'   Added simple ruleset to table info.
'   Added plunger groove and trigger for BIPL
'   Enabled the desktop backglass digits
'   Added some randomness to the plunger.
'   Added Flupper Bumpers, and adjusted color, dL, of tops, base, skirt, etc.
'   Tied GI and bumpers to Lampz
'   Added the gottlieb side rails from passion4pins.
'   Updated target, side walls, plunger cover, and bumper images.
'   Added the Spinner
'   Added the top right Gate
'   Implemented sling corrections.
'   Added the metal side walls and tweaked to match real table rolloff
'   Implemented the two sensors
'   Cut holes in the playfield for inserts and cutouts.  Also added text overlay
'   Implemented new standup targets by Rothbaurer; and ensured standalone
'   Added the saucer and associated movements
'   Updated all elements to have some disablelighting and controlled by GI script
'   Added cutouts for triggers and other holes.
'   Added all bulb prims
'v0.02  Adjusted the table slope to 5.75
'   Fine tuned the saucer settings with some saucer angle scatter (0.35)
'   Fine tuned the plungerfirespeed with a bit of randomness in script.
'   Added 3D primitive inserts.
'   Tweaked insert prim materials and lampz constants to get best look.
'   Red center insert is tied to GI, not a ROM lamp callout.
'   Created code to mask a corner condition of the tilt light coming on when multiball starts.
'   Fixed blue material issue.
'v0.03  Changed to individual rubbers by outlanes, adjusted a couple post alignments
'   Implemented new recommendations for EOSTNew, moved FlipperActivate and FlipperDeactivate to the flipper Sol subs
'   Put plastic walls in.
'   Added inner peg screws
'   Added GI for plastics.
'   Added option for yellow lane rollovers and tents
'   Added option for yellow and black flippers
'   Added option for plastic edge colors and glowing/rainbow from Redbone and iaakki with orange/red, yellow, and blue variants
'   Got the backbox functional for DesktopMode, updated new picture, and added GI bg lights for Destktop.
'     Tied the GI in correctly with the saucer and tilt to solenoid 11.
'   Added all the screws and screw caps. and tied to GI
'   Adjusted GI on the bumper cap ring down a touch.
'   Had an extra GI by the upper 2-way right gate.  Removed.
'   Added a gate prim and spring for the upper right 2-way gate, and added to GI.
'   Added new rollover switch prims and animated them.
'   Added options for Screw Cap Color
'   Added dynamic shadows, and code to stop shadows from going into the plunger lane via notrtlane and rtlanesense trigger.
'   Adjusted height of upper lane Rollovers
'   Changed the GI top to proper height.
'   Added physics walls
'   Prepped playfield GI for baking
'v0.04  Baked all the GI lighting, as well as the AO shadows.  All controlled by Lampz.
'   Adjusted all the GI colors to warm it up on the black payfield (from 255 150 68 to a warmer yellow of 243 210 101)
'   Increased the GI light under plastics to 30
'   Added disable lighting to GI bulbs
'   Adjusted the spinner dampening to 0.99 to spin a little longer.
'   Changed material on plastic yellow and plastic red slightly.
'   Added new flipper and DT backdrop images that Redbone enhanced.
'   Added VR Cabinet and VR Room.  Used Borgdog's stencils for the cab artwork.  Was fantastic and very accurate!
'   Updated the VR Backglass displays, color, position, etc.
'   Added shoot again reel for DT and VR
'   Updated the pov on Cab
'v0.05  Applied a fix for cab mode with the L12 ball release.
'   Updated several items based on recommendations from Redbone
'     Added orange Flipper Option
'   Adjusted the black flipper material
'   Adjusted the acrylic colors slightly, and adjusted the useroption RGB.
'   Adjusted the spinner dampening to .9935; more closely matches real table spinner spin.
'   Adjusted the slope of table to min 5.5 and max of 6.0... with 50% difficulty (so avg of 5.75).
'   Changed the spinner Prim and rod; and new image.
'   Redbone altered plastics image
'   Updated the GI color to somewhere between warm white and warm yellow (RGB 248 228 164)
'   Adjusted the bumper on and off levels for just a bit higher when on.
'   Updated additional lights on desktop backglass
'   Changed playfield friction to 0.24 and changed to slope of 6.  Had to adjust saucer power and plunger speed
'   Reduced playfield reflectivity
'   Added S-T-A-R reflections on back metal wall based on Borgdog's original table.
'   Added a tourny mode setting, with precanned user settings.
'   Converted images to webp
'   Only tied the VR backglass to first turn on.  It does not dim with tilt or saucer multiball hit.
'v0.06  Redbone updated playfield, plastics, spinner, GI image, star reflections, and target
'   Fixed a small error in L40 DisableLighting
'   Made the star reflections a little larger, and updated the DisableLighting
'v0.07  Lots of updates based on Borgdog's feedback from his real table including:
'   Corrected the spinner movement (had the physics spinner at 180 degrees off)
'   Updated to correct Gottlieb sling arms
'   Changed to different metal wall images
'   Adjusted opacity of the STAR reflections
'   Changed to different tent prims.  Also adjusted the dL on the tents and rollover caps.
'   Changed the small metal posts, to thinner ones, and adjusted the rubber and physics sizes accordingly.
'   Lowered the rollover switches down slightly.
'   Replaced the screws on the plastics to hex style and set to right height
'   Redbone updated the bumper, spinner, and target images
'v0.08  Changed the bumpers to different also include a light leaking out, and adjusted the color of prims and lights.
'   Redbone updated the bumper disk and flipper images
'   Updated the GI Rubber routine.
'   Added a missing post under the apron near the drain.
'   Added numbers for LUT name options
'   Changed VR Coin Door Prim to be the correct style and color of black.  (some were narrow, some were wider.. I chose wider)
'v0.09  Updated to Roth's latest flipper code, intended to be more accurate for flippers.
'    - Changed the triggers to 27 VPU spacing.
'    - Added the LF.ReProcxessBalls to collide subs.
'    - Updated flipper code for trajectory correction.
'    - Added function Function DistanceFromFlipperAngle
'   Switched to mid 80's Flippers, and adjustments from Rothbauerw.
'   Slight reduction (4.1 to 4.0) on slingshots.
'v0.10  Adjusted the flipper end angles to 72, to maintain 50 degree swing.  Had to update the triggersLF and RF as well.
'v0.11  Corrected error in z_wood material physics.
'   Added a higher friction metal material for upper right metal wall... hits the target on a proper left flipper shot now.
'   Slight adjustment to apron prim to cover edge of right wall.
'   Adjusted the EOS torque angle to 4.  And the torqe strength.  Helps with alley pass.
'v2.0.0 Updated the setdips code and ensured default dips working
'   Very slight adjustment to color on bumper caps and ring.
'   Small adjustment to right orbit wall to ensure ability to hit target on rolling left flipper shot.
'   Enabled tourney mode to adjust the DIP settings for no extra ball / novelty mode.
'   Added "click" sound to LUT changes.
'   Released
'v2.0.1 ROM isn't sending solenoid calls, and the DOF config utility fakes it out. Updated the Bumper_hit subs to add DOF 10X, DOFPulse
'v2.0.2 Minor error in the FindDips... dip 26 should have been /2
'   Updated the DIPS default code.  Simpler code and easier to manage.
'   Updated the sling_hit subs to add DOF 10X, DOFPulse
'   Adjusted the "pink" insert color slightly.
'   Had to modify the flipper trigers for addiitonal 3 degrees of trigger.
'   Ball could get stuck on rail near flipper.  Rounded corners of physics rail wall, and made bottom collidable.
'   Updated the wall "ramps" to be walls, and put new images
'   Redbone provided new images for the gate prims.
'   Rounded the small rubber peg rubbers some more.
'   Randomized the screw orientations
'   Added all options to menu by using F12
'   Removed old LUT code, and replaced in F12 menu.
'   Updated the bumper lighting to disclude the day/night mode.
'   For tourney mode, removed forcing the rainbow plastic edge
'   Added option for STAR reflections on backwall.  Users could turn up SSR if want other style reflections.
'   Added independant ROM control volume in the options menu.
'   Ball color change is in the script, and you can turn the option off to choose ball color and decal via table drop down menu.
'   Updated the black sidewood color image.


' Thanks to Borgdog for the playfield, cab, and other assets!  Also his video of game play of his real machine to match as close as possible.
' https://www.youtube.com/watch?v=NwFih3zu-4Y
' Thanks to Redbone for the playfield and plastic images, and Hauntfreaks for the updated backglass image!
' Also thanks to Rothbauerw for helping with the flipper physics and his new code and tweaks.
' Also thanks to the VPW team for testing, and everyone else that provided feedback!


'********************************************************************************************************************************
' Details of options
'********************************************************************************************************************************


' SideRailGotlieb = 1   '0 = standard metal side rails;  1 = gottlieb style bumpy side rails (desktop only) (default)
' LaneTentColor   = 1   '0 = standard red; 1 = yellow (default)
' FlipperColor  = 4   '1 = standard red; 2 = yellow; 3 = black; 4 = orange (default)
' ScrewCapColor = 0   '0 = White (default), 1 = Black, 2 = Bronze, 3 = Silver, 4 = Blue/green, 5 = Gold
' SideWallColor = 0   '0 = Wood walls, 1 = Black Walls
' PlasticColor  = 0   '0 = Off (default), 1 = Orange/Red, 2 = Yellow, 3 = Blue, 4 = User Defined, 5 = Rainbow, 6 = Random Option

  ' *** TOURNEY MODE ***
' Tourney = 0     '0 = Off (default), 1 = Presets for tourney mode:  includes: Harder DIPS settins (no extra ball)
              '  Tourney mode colors: Black flippers, yellow tents, and blue/green screw caps
    ' NOTE on Tourney Mode: This mode sets dip switches. You need to launch game, exit, and launch again for the ROM settings to take affect.


  ' ****************************************************
  ' *** If using VR Room:

'   WallClock   = 1     '1 Shows the clock in the VR minimal rooms only
'   poster    = 1   '1 Shows the flyer posters in the VR room only
'   poster2   = 1   '1 Shows the flyer posters in the VR room only
'   CustomWalls = 2   '0=UP's Original Minimal Walls, floor, and roof
                '1=Sixtoe's arcade style
                '2=DarthVito's Updated home walls with lights
                '3=DarthVito's plaster home walls
                '4=DarthVito's blue home walls


  ' ****************************************************
  ' DIP Switch options
  '  *** Note: once you launch first time, it will set the DIPS, but you need to relauch for it to take effect

'   SetDIPSwitches = 0  ' If you want to set the dips differently, set to 1, and then hit F6 to launch.

    ' These are the default DIPS that are and set at original table launch:
      '1=Coin Chute same
      '0=3rd coin chute no effect
      '0=Game mode of replay for hitting threshold
      '1=Number of balls = 3
      '1=Playfield Special of extra ball  (Tourney mode this is zero)
      '0=Novelty of normal game mode    (Tourney mode this set to 1)
      '1=Match feature
      '1=Background Sound


'----- Shadow Options -----
'   DynamicBallShadowsOn = 1    '0 = no dynamic ball shadow ("triangles" near slings and such), 1 = enable dynamic ball shadow
'   AmbientBallShadowOn = 1    '0 = Static shadow under ball ("flasher" image, like JP's), 1 = Moving ball shadow ("primitive" object, like ninuzzu's) - This is the only one that behaves like a true shadow!, 2 = flasher image shadow, but it moves like ninuzzu's

'----- General Sound Options -----
'   VolumeDial = 0.8        'Overall Mechanical sound effect volume. Recommended values should be no greater than 1.
'   BallRollVolume = 0.5      'Level of ball rolling volume. Value between 0 and 1

'********************************************************************************************************************************


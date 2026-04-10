' Embryon / IPD No. 783 / June, 1981 / 4 Players

' ____  _  _  ____  ____  _  _  __   __ _
'(  __)( \/ )(  _ \(  _ \( \/ )/  \ (  ( \
' ) _) / \/ \ ) _ ( )   / )  /(  O )/    /
'(____)\_)(_/(____/(__\_)(__/  \__/ \_)__)


' Version 2.0 by UnclePaulie 2023
' Table includes Hybrid VR/desktop/cabinet modes, new upscaled playfield by Movieguru and image improvements by Redbone, VPW physics, Fleep sounds, Lampz, 3D inserts, new GI, sling corrections, targets, saucers, playfield mesh, etc.
'   This version done with Medusa as a baseline to start.  Built everything ground up from there.

' Complete implementation details and credits found at the bottom of the script.


Option Explicit
Randomize


'*******************************************
' Desktop, Cab, and VR OPTIONS
'*******************************************

' Desktop, Cab, and VR Room are automatically selected.  However if in VR Room mode, you can change the environment with the magna save buttons.

const BallLightness = 2 '0 = dark, 1 = not as dark, 2 = bright, 3 = brightest
const FlipperColor = 0 ' 0 = Worn Yellow (default), 1 = Red, 2 = Blue, 3 = Yellow, 4 = Worn Red, 5 = Worn Blue

const leftoutlanemode = 1  ' 0 = Harder, 1 = Easier (default)


const bumperscale = .4      ' adjusts the level of blenddisablelighting and intensity on bumpers when lights are not on
const bumperflashlevel = 0.2  ' adjusts the level of bumper lights when on and / or flashing



' If using Desktop

const DTLampsON = 1 ' 0 = Desktop Backglass EMBRYON Flashing Lamps off.  1 = Flashing EMBRYON lamps flash according to ROM / real backglass.

Dim LUTset, DisableLUTSelector, LutToggleSound, LutToggleSoundLevel

LutToggleSound = True
LutToggleSoundLevel = .1
DisableLUTSelector = 0    ' Disables the ability to change LUT option with magna saves in game when set to 1

const SetDIPSwitches = 0  ' 0 is the default dip settings for game play. (RECOMMENDED) If you want to set the dips differently, set to 1, and then hit F6 to launch.

' *** If using VR Room:

const CustomWalls = 0 'set to 0 for Modern Minimal Walls, floor, and roof, 1 for Sixtoe's original walls and floor
const WallClock = 1   '1 Shows the clock in the VR minimal rooms only
const topper = 0     '0 = Off 1= On - Topper visible in VR Room only
const poster = 1     '1 Shows the flyer posters in the VR room only
const poster2 = 1    '1 Shows the flyer posters in the VR room only

' ****************************************************

'----- Shadow Options -----
Const DynamicBallShadowsOn  = 1   '0 = no dynamic ball shadow ("triangles" near slings and such), 1 = enable dynamic ball shadow
Const AmbientBallShadowOn   = 1   '0 = Static shadow under ball ("flasher" image, like JP's)
                  '1 = Moving ball shadow ("primitive" object, like ninuzzu's) - This is the only one that shows up on the pf when in ramps and fades when close to lights!
                  '2 = flasher image shadow, but it moves like ninuzzu's
'Ambient (Room light source)
Const AmbientBSFactor     = 1   '0 to 1, higher is darker
Const AmbientMovement   = 1   '1 to 4, higher means more movement as the ball moves left and right
Const offsetX       = 0   'Offset x position under ball (These are if you want to change where the "room" light is for calculating the shadow position,)
Const offsetY       = 5   'Offset y position under ball  (for example 5,5 if the light is in the back left corner)

'Dynamic (Table light sources)
Const DynamicBSFactor     = 0.9 '0 to 1, higher is darker
Const Wideness        = 20  'Sets how wide the dynamic ball shadows can get (20 +5 thinness should be most realistic for a 50 unit ball)
Const Thinness        = 5   'Sets minimum as ball moves away from source

'----- General Sound Options -----
Const VolumeDial = 0.8        ' Recommended values should be no greater than 1.
Const BallRollVolume = 0.5      'Level of ball rolling volume. Value between 0 and 1
Const RampRollVolume = 0.5      'Level of ramp rolling volume. Value between 0 and 1

'----- Phsyics Mods -----
Const FlipperCoilRampupMode = 0     '0 = fast, 1 = medium, 2 = slow (tap passes should work)
Const TargetBouncerEnabled = 1    '0 = normal standup targets, 1 = bouncy targets, 2 = orig TargetBouncer
Const TargetBouncerFactor = 0.7   'Level of bounces. Recommmended value of 0.7 when TargetBouncerEnabled


'//////////////---- LUT (Colour Look Up Table) ----//////////////
'0 = Fleep Natural Dark 1
'1 = Fleep Natural Dark 2
'2 = Fleep Warm Dark
'3 = Fleep Warm Bright
'4 = Fleep Warm Vivid Soft
'5 = Fleep Warm Vivid Hard
'6 = Skitso Natural and Balanced
'7 = Skitso Natural High Contrast
'8 = 3rdaxis Referenced THX Standard
'9 = CalleV Punchy Brightness and Contrast
'10 = HauntFreaks Desaturated
'11 = Tomate Washed Out
'12 = VPW Original 1 to 1
'13 = Bassgeige
'14 = Blacklight
'15 = B&W Comic Book
'16 = Skitso New Warmer LUT
'17 = Original LUT


LoadLUT

'LUTset = 16      ' Override saved LUT for debug
SetLUT
ShowLUT_Init


' ****************************************************
' standard definitions
' ****************************************************

Dim VR_Room, cab_mode, DesktopMode: DesktopMode = Table1.ShowDT
If RenderingMode = 2 Then VR_Room=1 Else VR_Room=0      'VRRoom set based on RenderingMode in version 10.72
if Not DesktopMode and VR_Room=0 Then cab_mode=1 Else cab_mode=0

Const UseSolenoids  = 2
Const UseLamps    = 0
Const UseSync     = 0
Const HandleMech  = 0
Const UseGI     = 0
Const cGameName   = "embryond"    'ROM name
Const ballsize    = 50
Const ballmass    = 1
Const UsingROM    = True      'The UsingROM flag is to indicate code that requires ROM usage.

'***********************

Const tnob = 7            'Total number of balls  (2 multiball, and 5 captive balls)
Const lob = 0           'Locked balls

Dim tablewidth: tablewidth = Table1.width
Dim tableheight: tableheight = Table1.height

Dim i, eBall1, eball2, eball3, eball4, eball5, eball6, eball7, gBOT, RFLipOn
Dim BIPL : BIPL = False       'Ball in plunger lane

Dim gion: gion = 0

Dim DT17up, DT20up, DT21up, DT22up, DT25up, DT26up, DT27up

const DTLDelayTime = 300
const DTRDelayTime = 550

' Flipper Color

if FlipperColor = 1 Then
  Lflipmesh1.image = "flipperAOred"
  Rflipmesh1.image = "flipperAOred"
  RflipmeshUp.image = "flipperAOred"
Elseif FlipperColor = 2 Then
  Lflipmesh1.image = "flipperAOblue"
  Rflipmesh1.image = "flipperAOblue"
  RflipmeshUp.image = "flipperAOblue"
Elseif FlipperColor = 3 Then
  Lflipmesh1.image = "flipperAOyellow"
  Rflipmesh1.image = "flipperAOyellow"
  RflipmeshUp.image = "flipperAOyellow"
Elseif FlipperColor = 4 Then
  Lflipmesh1.image = "flipperAOredworn"
  Rflipmesh1.image = "flipperAOredworn"
  RflipmeshUp.image = "flipperAOredworn"
Elseif FlipperColor = 5 Then
  Lflipmesh1.image = "flipperAOblueworn"
  Rflipmesh1.image = "flipperAOblueworn"
  RflipmeshUp.image = "flipperAOblueworn"
Else
  Lflipmesh1.image = "flipperAOyellowworn"
  Rflipmesh1.image = "flipperAOyellowworn"
  RflipmeshUp.image = "flipperAOyellowworn"
End If

On Error Resume Next
ExecuteGlobal GetTextFile("controller.vbs")
If Err Then MsgBox "You need the controller.vbs in order to run this table, available in the vp10 package"
On Error Goto 0

LoadVPM "01130100", "Bally.VBS", 3.21


'**********************************************************************************************************
'Solenoids
'**********************************************************************************************************

'SolCallback(1) = "SolDTSingle"       'Single Drop Target Reset
'                     'Removed call from solenoid action, as there's a bug in ROMs... control via script.
SolCallback(2)  = "SolDTLeft"       'Left Side 3 Bank Reset
SolCallback(3)  = "SolDTTop"        'Top 3 Bank Reset

SolCallback(4)  = "SolSaucer"       'Saucer
SolCallback(5)  = "SolOutholeKicker"    'trough to plunger lane
SolCallback(6)  = "SolKnocker"        'knocker solenoid


SolCallback(18) = "BGGI"          'Playfield GI
SolCallback(19) = "PFGI"          'Backglass GI

SolCallback(20) = "SolRFlipper2"      'Lighted flipper on the right side.

SolCallback(sllflipper)="SolLFlipper"
SolCallback(slrflipper)="SolRFlipper"


'*****************
'   GI
'*****************

Sub PFGI(Enabled)
    If Enabled Then
    SetLamp 140, 1
    gion = 1

' Set Drop Target Shadows to state when GI is back on.
  if DT17up=1 Then dtsh17.visible = 1
  if DT20up=1 Then dtsh20.visible = 1
  if DT21up=1 Then dtsh21.visible = 1
  if DT22up=1 Then dtsh22.visible = 1
  if DT25up=1 Then dtsh25.visible = 1
  if DT26up=1 Then dtsh26.visible = 1
  if DT27up=1 Then dtsh27.visible = 1

    Else
    gion = 0
    SetLamp 140, 0

' Set Drop Target Shadows to off when GI is goes off.
  for each xx in ShadowDT
    xx.visible=False
  Next

    End If
End Sub


Sub BGGI(Enabled)
    If Enabled Then
    SetLamp 141, 1
    Else
    SetLamp 141, 0
    End If
End Sub


'*******************************************
'  Timers
'*******************************************

' The game timer interval is 10 ms
Sub GameTimer_Timer()
  Cor.Update            'update ball tracking
  RollingUpdate         'update rolling sounds
  DoDTAnim            'handle drop target animations
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


' This subroutine updates the flipper shadows and visual primitives
Sub FlipperVisualUpdate
  FlipperLSh.RotZ = LeftFlipper.currentangle
  FlipperRSh.RotZ = RightFlipper.currentangle
  Lflipmesh1.RotZ = LeftFlipper.CurrentAngle
  Rflipmesh1.RotZ = RightFlipper.CurrentAngle
  RflipmeshUp.RotZ = RightFlipper1.CurrentAngle
  FlipperRShUp.RotZ = RightFlipper1.currentangle

  Gate1p.rotx = max(Gate1.CurrentAngle,0)
  Gate2p.rotx = max(Gate2.CurrentAngle,0)
  Gate3p.rotx = max(Gate3.CurrentAngle,0)
  Gate4p.rotx = max(Gate4.CurrentAngle,0)

End Sub


'**********************************************************************************************************
'Initiate Table
'**********************************************************************************************************

Sub Table1_Init
  vpmInit Me
  On Error Resume Next
    With Controller
    .GameName = cGameName
    If Err Then MsgBox "Can't start Game" & cGameName & vbNewLine & Err.Description : Exit Sub
    .SplashInfoLine = "Embryon (Bally 1981)"&chr(13)&"by UnclePaulie"
    .HandleMechanics=0
    .HandleKeyboard=0
    .ShowDMDOnly=1
    .ShowFrame=0
    .ShowTitle=0
        .hidden = 1
    .Games(cGameName).Settings.Value("sound") = 1 ' Set sound (0=OFF, 1=ON)
    If Err Then MsgBox Err.Description
  End With
  On Error Goto 0
    Controller.Run
  If Err Then MsgBox Err.Description
  On Error Goto 0

  PinMAMETimer.Interval=PinMAMEInterval
  PinMAMETimer.Enabled=1

  vpmNudge.TiltSwitch = 7
  vpmNudge.Sensitivity = 5
  vpmNudge.TiltObj = Array(LeftSlingshot,RightSlingshot, LeftFlipper, RightFlipper, RightFlipper1, Bumper1, Bumper2, Bumper3, Bumper4)

'Ball initializations need for physical trough and ball shadows
  Set eBall1 = BallRelease.CreateSizedballWithMass(Ballsize/2,Ballmass)
  Set eBall2 = Slot.CreateSizedballWithMass(Ballsize/2,Ballmass)

' Create captive balls
  Set eBall3 = kicker17.CreateSizedballWithMass(Ballsize/2,Ballmass)
  Set eBall4 = Kicker25.CreateSizedballWithMass(Ballsize/2,Ballmass)
  Set eBall5 = KickerCenter.CreateSizedballWithMass(Ballsize/2,Ballmass)
  Set eBall6 = Kicker33.CreateSizedballWithMass(Ballsize/2,Ballmass)
  Set eBall7 = Kicker34.CreateSizedballWithMass(Ballsize/2,Ballmass)

'Ball initializations
  gBOT = Array(eBall1, eball2, eball3, eball4, eball5, eball6, eball7)
  Controller.Switch(5) = 1
  Controller.Switch(8) = 1

' Add balls to shadow dictionary
  For Each xx in gBOT
    bsDict.Add xx.ID, bsNone
  Next

'Saucer and Right Flipper initializations
  controller.switch(4) = 0
  RFLipOn=0
  Controller.Switch(3) = 0

'Release captive balls
  Kicker17.Kick 270,1
  Kicker25.Kick 270,1
  Kicker33.Kick 270,1
  Kicker34.Kick 270,1
  KickerCenter.Kick 270,1


' Set up VR Backglass
  if VR_Room = 1 Then
    setup_backglass()
    SetBackglass
  End If

' Make drop target shadows invisible
  Dim xx
  for each xx in ShadowDT
    xx.visible=False
  Next

' Drop Target Variable state for DT Shadows
  DT17up=1
  DT20up=1
  DT21up=1
  DT22up=1
  DT25up=1
  DT26up=1
  DT27up=1

  FlasherGI.opacity = 0

' If user wants to set dip switches themselves it will force them to set it via F6.
If SetDIPSwitches = 0 Then
  SetDefaultDips
End If

End Sub

Sub table1_Paused:Controller.Pause = 1:End Sub
Sub table1_unPaused:Controller.Pause = 0:End Sub
Sub table1_exit
  SaveLUT
  Controller.stop
End Sub

'********************************************
' Keys and Plunger code
'********************************************

Sub table1_KeyDown(ByVal Keycode)

  If keycode = LeftMagnaSave Then
    bLutActive = True
  End If

  If keycode = RightMagnaSave Then
    If bLutActive Then
      if DisableLUTSelector = 0 then
        If LutToggleSound Then
          Playsound "click", 0, LutToggleSoundLevel * VolumeDial, 0, 0.2, 0, 0, 0, 1
        End If
        LUTSet = LUTSet  + 1
        if LutSet > 17 then LUTSet = 0
        SetLUT
        ShowLUT
      end if
    End If
    End If

  If keycode = LeftTiltKey Then Nudge 90, 0.5 ': SoundNudgeLeft
  If keycode = RightTiltKey Then Nudge 270, 0.5 ': SoundNudgeRight
  If keycode = CenterTiltKey Then Nudge 0, 0.5 ': SoundNudgeCenter


  If KeyCode = PlungerKey Then
    Plunger.PullBack
    SoundPlungerPull()
    TimerVRPlunger.Enabled = True
    TimerVRPlunger1.Enabled = False
    PinCab_Shooter.Y = -390
  End If

  If keycode = LeftFlipperKey Then
    VRFlipperButtonLeft.X = VRFlipperButtonLeft.X + 8
    FlipperActivate LeftFlipper, LFPress
  End If

  If keycode = RightFlipperKey Then
    VRFlipperButtonRight.X = VRFlipperButtonRight.X - 8
    FlipperActivate RightFlipper, RFPress
    Controller.Switch(3) = 1
    RFlipOn = 1
  End If

  If keycode = StartGameKey Then
    StartButton.y = 1939.978 - 5
    SoundStartButton
  End If

  If keycode = keyInsertCoin1 or keycode = keyInsertCoin2 or keycode = keyInsertCoin3 or keycode = keyInsertCoin4 Then 'Use this for ROM based games
    Select Case Int(rnd*3)
      Case 0: PlaySound ("Coin_In_1"), 0, CoinSoundLevel, 0, 0.25
      Case 1: PlaySound ("Coin_In_2"), 0, CoinSoundLevel, 0, 0.25
      Case 2: PlaySound ("Coin_In_3"), 0, CoinSoundLevel, 0, 0.25
    End Select
  End If

    If vpmKeyDown(keycode) Then Exit Sub

End Sub


Sub table1_KeyUp(ByVal Keycode)

    If keycode = LeftMagnaSave Then
    bLutActive = False
  End If

  If keycode = PlungerKey Then
    Plunger.Fire
    If BIPL = 1 Then
      SoundPlungerReleaseBall()     'Plunger release sound when there is a ball in shooter lane
    Else
      SoundPlungerReleaseNoBall()     'Plunger release sound when there is no ball in shooter lane
    End If
    TimerVRPlunger.Enabled = False
    TimerVRPlunger1.Enabled = True
    PinCab_Shooter.Y = -390
  End If

  If keycode = LeftFlipperKey Then
    VRFlipperButtonLeft.X = VRFlipperButtonLeft.X - 8
    FlipperDeActivate LeftFlipper, LFPress
  End If

  If keycode = RightFlipperKey Then
    VRFlipperButtonRight.X = VRFlipperButtonRight.X + 8
    FlipperDeActivate RightFlipper, RFPress
    Controller.Switch(3) = 0
  End If

  If keycode = StartGameKey Then
    StartButton.y = 1939.978
  End If

  if vpmKeyUp(keycode) Then Exit Sub

End Sub


'*******************************************
'  Flippers
'*******************************************

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
    RightFlipper1.rotatetoend
    If rightflipper.currentangle > rightflipper.endangle - ReflipAngle Then
      RandomSoundReflipUpRight RightFlipper
    Else
      SoundFlipperUpAttackRight RightFlipper
      RandomSoundFlipperUpRight RightFlipper
    End If
  Else
    RightFlipper.RotateToStart
    RightFlipper1.RotateToStart
    If RightFlipper.currentangle > RightFlipper.startAngle + 5 Then
      RandomSoundFlipperDownRight RightFlipper
    End If
    FlipperRightHitParm = FlipperUpSoundLevel
  End If
End Sub

Sub SolRFlipper2(Enabled)
  If RFLipOn = 1 Then
    If Enabled Then
      RightFlipper2.RotateToEnd
      RandomSoundFlipperUpRight RightFlipper2
    Else
      RightFlipper2.RotateToStart
      RandomSoundFlipperDownRight RightFlipper2
    End If
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

Sub RightFlipper1_Collide(parm)
  RightFlipperCollide parm
End Sub

Sub RightFlipper2_Collide(parm)
  RightFlipperCollide parm
End Sub

'******************************************************
'     TROUGH BASED ON FOZZY and Rothbauerw
'******************************************************

Sub BallRelease_Hit : Controller.Switch(8) = 1 : UpdateTrough : End Sub
Sub BallRelease_UnHit : Controller.Switch(8) = 0 : UpdateTrough : End Sub
Sub Slot_Hit :  Controller.Switch(5) = 1 : UpdateTrough : End Sub
Sub Slot_UnHit : Controller.Switch(5) = 0 : UpdateTrough : End Sub


Sub UpdateTrough
  UpdateTroughTimer.Interval = 300
  UpdateTroughTimer.Enabled = 1
End Sub

Sub UpdateTroughTimer_Timer
  If BallRelease.BallCntOver = 0 Then Slot.kick 60, 8
  If Slot.BallCntOver = 0 Then Drain.kick 60, 18
  Me.Enabled = 0
End Sub


'********************************************
' Drain hole and saucer kickers
'********************************************

Dim RNDKickValue1, RNDKickAngle1  'Random Values for saucer kick and angles

Sub SolOutholeKicker(enabled)
  If enabled Then
    BallRelease.kick 60, 12
    RandomSoundBallRelease BallRelease
    'manually raise the single drop target
      If DT17up = 0 then
        RaiseSingleDT
      End if
  End If
End Sub

Sub Drain_Hit()
  UpdateTrough
  RandomSoundDrain Drain
End Sub



Sub sw4_Hit
  SoundSaucerLock
  Controller.Switch(4) = 1
End Sub

Sub SolSaucer(Enabled)
  If Enabled Then
    RNDKickAngle1 = RndInt(268,272)   ' Generate random value between 268 and 272. (Variance of +/- 2 from 270)
    RNDKickValue1 = RndInt(10,14)     ' Generate random value between 10 and 14. (Variance of +/- 2 from 12)
    sw4.kick RNDKickAngle1, RNDKickValue1
    sw4Step = 0
    sw4.timerenabled = 1
  End If
End Sub

Dim SW4Step
Sub SW4_Timer()
  Select Case SW4Step
    Case 0: pKickerArm.Rotx = 4
    Case 2: pKickerArm.Rotx = 8
    Case 3: pKickerArm.Rotx = 8
    Case 4: pKickerArm.Rotx = 4
    Case 5: pKickerArm.Rotx = 0:sw4.timerenabled = 0:sw4Step = -1:
  End Select
  sw4Step = sw4Step + 1
End Sub

Sub sw4_Unhit
  SoundSaucerKick 1, sw4
  controller.switch(4) = 0
End Sub


'*******************************************
' Rollovers
'*******************************************

Sub sw12_Hit   : Controller.Switch(12) = 1 : End Sub
Sub sw12_Unhit : Controller.Switch(12) = 0 : End Sub
Sub sw13_Hit   : Controller.Switch(13) = 1 : End Sub
Sub sw13_Unhit : Controller.Switch(13) = 0 : End Sub
Sub sw14_Hit   : Controller.Switch(14) = 1 : End Sub
Sub sw14_Unhit : Controller.Switch(14) = 0 : End Sub
Sub sw15_Hit   : Controller.Switch(15) = 1 : End Sub
Sub sw15_Unhit : Controller.Switch(15) = 0 : End Sub
Sub sw19_Hit   : Controller.Switch(19) = 1 : End Sub
Sub sw19_Unhit : Controller.Switch(19) = 0 : End Sub
Sub sw23_Hit   : Controller.Switch(23) = 1 : End Sub
Sub sw23_Unhit : Controller.Switch(23) = 0 : End Sub
Sub sw24_Hit   : Controller.Switch(24) = 1 : End Sub
Sub sw24_Unhit : Controller.Switch(24) = 0 : End Sub


'*******************************************
' Ball in Plunger Lane
'*******************************************

Sub Trigger1_Hit
  BIPL = True
End Sub

Sub Trigger1_UnHit
  BIPL = False
End Sub


'*******************************************
'Star Triggers
'*******************************************

Sub Trigger2a_Hit:vpmTimer.PulseSw 2:me.timerenabled = 0:AnimateStar star2a, Trigger2a, 1:End Sub
Sub Trigger2a_UnHit:me.timerinterval = 7:me.timerenabled = 1:End Sub
Sub Trigger2a_timer:AnimateStar star2a, Trigger2a, 0:End Sub

Sub Trigger2b_Hit:vpmTimer.PulseSw 2:me.timerenabled = 0:AnimateStar star2b, Trigger2b, 1:End Sub
Sub Trigger2b_UnHit:me.timerinterval = 7:me.timerenabled = 1:End Sub
Sub Trigger2b_timer:AnimateStar star2b, Trigger2b, 0:End Sub

Sub Trigger2c_Hit:vpmTimer.PulseSw 2:me.timerenabled = 0:AnimateStar star2c, Trigger2c, 1:End Sub
Sub Trigger2c_UnHit:me.timerinterval = 7:me.timerenabled = 1:End Sub
Sub Trigger2c_timer:AnimateStar star2c, Trigger2c, 0:End Sub

Sub AnimateStar(prim, sw, action) ' Action = 1 - to drop, 0 to raise
  If action = 1 Then
    prim.transz = -4
  Else
    prim.transz = prim.transz + 0.5
    if prim.transz = -2 and Rnd() < 0.05 Then
      sw.timerenabled = 0
    Elseif prim.transz >= 0 Then
      prim.transz = 0
      sw.timerenabled = 0
    End If
  End If
End Sub


'*******************************************
' Spinners
'*******************************************

Sub Spinner1_Spin()
  vpmtimer.PulseSw 18
  SoundSpinner Spinner1
End Sub


'***********Rotate Spinner

Sub SpinnerTimer
  SpinnerPrim1.Rotx = Spinner1.CurrentAngle
  SpinnerRod1.TransX = sin( (Spinner1.CurrentAngle+180) * (2*PI/360)) * 12
  SpinnerRod1.TransZ = sin( (Spinner1.CurrentAngle- 90) * (2*PI/360)) * 10
End Sub


'*******************************************
' Bumpers
'*******************************************

Sub Bumper1_Hit 'Right Bottom
  RandomSoundBumperBottom Bumper1
  vpmTimer.PulseSw 37
End Sub

Sub Bumper2_Hit 'Right Top
  RandomSoundBumperTop Bumper2
  vpmTimer.PulseSw 38
End Sub

Sub Bumper3_Hit 'Left Bottom
  RandomSoundBumperBottom Bumper3
  vpmTimer.PulseSw 39
End Sub

Sub Bumper4_Hit 'Left Top
  RandomSoundBumperTop Bumper4
  vpmTimer.PulseSw 40
End Sub


'********************************************
'  Targets
'********************************************

'*******************************************
' Round Targets
'*******************************************

Sub sw17_Hit
  STHit 17
End Sub

Sub sw17o_Hit
  TargetBouncer Activeball, 1
End Sub

Sub sw28_Hit
  STHit 28
End Sub

Sub sw28o_Hit
  TargetBouncer Activeball, 1
End Sub

Sub sw29_Hit
  STHit 29
End Sub

Sub sw29o_Hit
  TargetBouncer Activeball, 1
End Sub

Sub sw30_Hit
  STHit 30
End Sub

Sub sw30o_Hit
  TargetBouncer Activeball, 1
End Sub

Sub sw31_Hit
  STHit 31
End Sub

Sub sw31o_Hit
  TargetBouncer Activeball, 1
End Sub

Sub sw32_Hit
  STHit 32
End Sub

Sub sw32o_Hit
  TargetBouncer Activeball, 1
End Sub

Sub sw33_Hit
  STHit 33
End Sub

Sub sw33o_Hit
  TargetBouncer Activeball, 1
End Sub

Sub sw34_Hit
  STHit 34
End Sub

Sub sw34o_Hit
  TargetBouncer Activeball, 1
End Sub


'********************************************
' Drop Target Hits
'********************************************

Sub sw17dt_Hit
  DTHit 17
  DT17up=0
End Sub

Sub sw20_Hit
  DTHit 20
  DT20up=0
End Sub

Sub sw21_Hit
  DTHit 21
  DT21up=0
End Sub

Sub sw22_Hit
  DTHit 22
  DT22up=0
End Sub

Sub sw25_Hit
  DTHit 25
  DT25up=0
End Sub

Sub sw26_Hit
  DTHit 26
  DT26up=0
End Sub

Sub sw27_Hit
  DTHit 27
  DT27up=0
End Sub


'********************************************
' Drop Target Solenoid Controls
'********************************************

' Removed this section of code, as there's a bug in ROMs... I now control the single drop target via script.
'Sub SolDTSingle (enabled)
' If enabled then
'   If DT17up = 0 then
'     vpmTimer.AddTimer DTRDelayTime-100, "RaiseSingleDT'"    'Wait a little shorter than the STHIT 17 functional call, to ensure it doesn't go up twice.
'   End If
' End if
'End Sub

Sub SolDTLeft (enabled)

  If enabled then
    vpmTimer.AddTimer DTLDelayTime, "RandomSoundDropTargetReset sw26p'"
    vpmTimer.AddTimer DTLDelayTime, "DTRaise 25'" 'Wait an extra 1/2 second to ensure ball is out of way.
    vpmTimer.AddTimer DTLDelayTime, "DTRaise 26'"
    vpmTimer.AddTimer DTLDelayTime, "DTRaise 27'"
    DT25up=1
    DT26up=1
    DT27up=1

    vpmTimer.AddTimer DTLDelayTime, "dtsh25.visible=True'"
    vpmTimer.AddTimer DTLDelayTime, "dtsh26.visible=True'"
    vpmTimer.AddTimer DTLDelayTime, "dtsh27.visible=True'"
  End if
End Sub

Sub SolDTTop (enabled)
  If enabled then

    vpmTimer.AddTimer DTLDelayTime, "RandomSoundDropTargetReset sw21p'"
    vpmTimer.AddTimer DTLDelayTime, "DTRaise 20'" 'Wait an extra 1/2 second to ensure ball is out of way.
    vpmTimer.AddTimer DTLDelayTime, "DTRaise 21'"
    vpmTimer.AddTimer DTLDelayTime, "DTRaise 22'"
    DT20up=1
    DT21up=1
    DT22up=1

    vpmTimer.AddTimer DTLDelayTime, "dtsh20.visible=True'"
    vpmTimer.AddTimer DTLDelayTime, "dtsh21.visible=True'"
    vpmTimer.AddTimer DTLDelayTime, "dtsh22.visible=True'"

  End if
End Sub


Sub RaiseSingleDT
  RandomSoundDropTargetReset sw17dtp
  DTRaise 17
  DT17up=1
  dtsh17.visible=True
End Sub

'*******************************************
' Leaf Standup Sensors
'*******************************************

Sub RubberBand010_Hit: vpmTimer.PulseSw 1:End Sub
Sub RubberBand015_Hit: vpmTimer.PulseSw 1:End Sub
Sub RubberBand021_Hit: vpmTimer.PulseSw 1:End Sub
Sub RubberBand022_Hit: vpmTimer.PulseSw 1:End Sub


'*******************************************
'  Knocker
'*******************************************

Sub SolKnocker(Enabled)
  If enabled Then
    KnockerSolenoid
  End If
End Sub


'*******************************************
' Sling Shot Animations
' Rstep and Lstep  are the variables that increment the animation
'*******************************************
Dim RStep, Lstep, R2Step

Sub RightSlingShot_Slingshot
  RS.VelocityCorrect(ActiveBall)
  vpmTimer.PulseSw 35
  RandomSoundSlingshotRight SLING1
    RSling.Visible = 0
    RSling1.Visible = 1
    sling1.rotx = 12
    RStep = 0
    RightSlingShot.TimerEnabled = 1
End Sub

Sub RightSlingShot_Timer
    Select Case RStep
        Case 3:RSLing1.Visible = 0:RSLing2.Visible = 1:sling1.rotx = 10
        Case 4:RSLing2.Visible = 0:RSLing.Visible = 1:sling1.rotx = 0:RightSlingShot.TimerEnabled = 0
    End Select
    RStep = RStep + 1
End Sub

Sub LeftSlingShot_Slingshot
  LS.VelocityCorrect(ActiveBall)
  vpmTimer.PulseSw 36
  RandomSoundSlingshotLeft SLING2
    LSling.Visible = 0
    LSling1.Visible = 1
    sling2.rotx = 12
    LStep = 0
    LeftSlingShot.TimerEnabled = 1
End Sub

Sub LeftSlingShot_Timer
    Select Case LStep
        Case 3:LSLing1.Visible = 0:LSLing2.Visible = 1:sling2.rotx = 10
        Case 4:LSLing2.Visible = 0:LSLing.Visible = 1:sling2.rotx = 0:LeftSlingShot.TimerEnabled = 0
    End Select
    LStep = LStep + 1
End Sub


'*************************************
'Bally Embryon
'*************************************
Sub editDips
  if SetDIPSwitches = 1 Then
    Dim vpmDips:Set vpmDips = New cvpmDips
    With vpmDips
        .AddForm 500, 600, "Embryon - DIP switches"
        .AddFrame 120, 0, 120, "Maximum credits", &H03000000, Array("10 credits", 0, "15 credits", &H01000000, "25 credits", &H02000000, "40 credits", &H03000000)                                                        'dip 25&26
    .AddFrame 120, 75, 120, "Balls per game", &H40000000, Array("2 balls", &H40000000+&H80000000, "3 balls", 0, "4 balls", &H80000000,  "5 balls", &H40000000)                                                                                        'dip 31&32
    .AddFrame 2, 0, 100,   "Play/Credits Chute 1", &H00000002+&H00000008, Array("1/1Coin", 0, "1/2Coin", &H00000003, "1/2Coin", &H3, "2/5Coin",&H00000009+&H0000006+&H00000010, "14/1Coin",&H00000002+&H00000008)                     'dip 1-5
    .AddFrame 2, 90, 100, "Play/Credits Chute 2", &H00080000+&H00060000+&H00100000, Array("1/1Coin", 0, "1/1Coin", &H00010000, "2/1Coin", &H00020000, "3/1Coin", &H00030000, "4/1Coin", &H00040000, "14/1Coin",&H00080000+&H00060000+&H00100000)    'dip 17-20
    .AddFrame 2, 195, 100, "Play/Credits Chute 3", &H00001000, Array("1/1Coin", 0, "1/2Coin", &H00000300, "1/2Coin", &H300, "2/5Coin", &H00000900+&H0000600+&H00000100, "14/1Coin", &H00000200+&H00000800)        'dip 9-13

        .AddChk 120, 155, 263,  Array("Attract Mode Voice (Activate Embryon)", &H20000000)                                                                          'dip 30
    .AddChk 120, 175, 263,  Array("Memory for Bonus and Specials", &H00000040)                                                                        'dip 7
    .AddChk 120, 195, 263,  Array("Memory for Multipliers", &H00000080)                                                                         'dip 8
    .AddChk 120, 215, 263,  Array("Memory for lit Lanes 1&2", &H00400000)                                                                         'dip 23
    .AddChk 120, 235, 263,  Array("Memory for Flipsave Feature", &H00800000)                                                                      'dip 24
    .AddChk 120, 255, 263,  Array("Memory for Top Center Lane Feature", &H00002000)                                                                   'dip 14
    .AddChk 120, 275, 263,  Array("Memory for Left Side Captive Ball Feature", 32768)                                                                 'dip 16
    .AddChk 120, 295, 263,  Array("Memory for Right Side Captive Ball Feature", &H00004000)                                                               'dip 15
    .AddChk 120, 315, 263,  Array("Memory for lit Center Ouside Targets", &H00200000)                                                                 'dip 22
    .AddChk 120, 335, 263,  Array("Hitting 3 top drop targets once to Collect/ off Twice ", &H00000020)                                                         'dip 6
    .AddChk 120, 355, 263,  Array("Hitting Center left/right target will light 2 lites", &H00100000)                                                          'dip 21
    .AddChk 120, 375, 263,  Array("Replays Earned will be Collected/ off only 1 per play", &H10000000)                                                            'dip 29
    .AddChk 2, 290, 115,  Array("Match feature", &H08000000)                                                                                                                                                'dip 28
    .AddChk 2, 310, 115,  Array("Credits displayed", &H04000000)                                                                            'dip 27
        .AddLabel 50, 395, 350, 20, "After hitting OK, press F3 to reset game with new settings."
        .ViewDips
    End With
  End If
End Sub
Set vpmShowDips = GetRef("editDips")


Sub SetDefaultDips

' Number of Balls / game
  SetDip &H40000000,0   'Number of balls.  0,0 = 3, 0,1 = 3, 1,0 = 5, 1,1 = 5
  SetDip &H80000000,0

' All other Dips
  SetDip &H20000000,1   'Attract Mode Voice (Activate Embryon)
  SetDip &H00000040,1   'Memory for Bonus and Specials
  SetDip &H00000080,0   'Memory for Multipliers
  SetDip &H00400000,1   'Memory for lit Lanes 1&2
  SetDip &H00800000,1   'Memory for Flipsave Feature
  SetDip &H00002000,0   'Memory for Top Center Lane Feature
  SetDip 32768,1      'Memory for Left Side Captive Ball Feature
  SetDip &H00004000,1   'Memory for Right Side Captive Ball Feature
  SetDip &H00200000,0   'Memory for lit Center Ouside Targets
  SetDip &H00000020,0   'Hitting 3 top drop targets once to Collect/ off Twice
  SetDip &H00100000,1   'Hitting Center left/right target will light 2 lites
  SetDip &H10000000,1   'Replays Earned will be Collected/ off only 1 per play
  SetDip &H08000000,1   'Match feature
  SetDip &H04000000,1   'Credits display

End Sub

Sub SetDip(pos,value)
  dim mask : mask = 255
  if value >= 0.5 then value = 1 else value = 0
  If pos >= 0 and pos <= 255 Then
    pos = pos And &H000000FF
    mask = mask And (Not pos)
    Controller.Dip(0) = Controller.Dip(0) And mask
    Controller.Dip(0) = Controller.Dip(0) + pos*value
  Elseif pos >= 256 and pos <= 65535 Then
    pos = ((pos And &H0000FF00)\&H00000100) And 255
    mask = mask And (Not pos)
    Controller.Dip(1) = Controller.Dip(1) And mask
    Controller.Dip(1) = Controller.Dip(1) + pos*value
  Elseif pos >= 65536 and pos <= 16777215 Then
    pos = ((pos And &H00FF0000)\&H00010000) And 255
    mask = mask And (Not pos)
    Controller.Dip(2) = Controller.Dip(2) And mask
    Controller.Dip(2) = Controller.Dip(2) + pos*value
  Elseif pos >= 16777216 Then
    pos = ((pos And &HFF000000)\&H01000000) And 255
    mask = mask And (Not pos)
    Controller.Dip(3) = Controller.Dip(3) And mask
    Controller.Dip(3) = Controller.Dip(3) + pos*value
  End If
End Sub



'*******************************************
'  Block shadows in plunger lane
'*******************************************

dim notrtlane: notrtlane = 1
sub rtlanesense_Hit() : notrtlane = 0 : end Sub
sub rtlanesense_UnHit() : notrtlane = 1 : end Sub

'***************************************************************
'****  VPW DYNAMIC BALL SHADOWS by Iakki, Apophis, and Wylte
'***************************************************************

' *** Required Functions, enable these if they are not already present elswhere in your table
Function max(a,b)
  if a > b then
    max = a
  Else
    max = b
  end if
end Function


' *** Trim or extend these to match the number of balls/primitives/flashers on the table!  (will throw errors if there aren't enough objects)
dim objrtx1(7), objrtx2(7)
dim objBallShadow(8)
Dim OnPF(7)
Dim BallShadowA
BallShadowA = Array (BallShadowA0, BallShadowA1, BallShadowA2, BallShadowA3, BallShadowA4, BallShadowA5, BallShadowA6, BallShadowA7)
Dim DSSources(30), numberofsources

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
  '   Dim gBOT: gBOT=getballs 'Uncomment if you're destroying balls - Not recommended! #SaveTheBalls

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
      If gBOT(s).Z < 30 And gBOT(s).X < 1147 AND notrtlane = 1 Then 'Parameters for where the shadows can show, here they are not visible above the table (no upper pf) or in the plunger lane
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
' VPW TargetBouncer for targets and posts by Iaakki, Wrd1972, Apophis
'******************************************************


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

' Add targets or posts to the TargetBounce collection if you want to activate the targetbouncer code from them
Sub TargetBounce_Hit(idx)
  TargetBouncer activeball, 1
End Sub


'******************************************************
'****  FLIPPER CORRECTIONS by nFozzy
'******************************************************


'******************************************************
' Flippers Polarity (Select appropriate sub based on era)
'******************************************************

dim LF : Set LF = New FlipperPolarity
dim RF : Set RF = New FlipperPolarity

InitPolarity

'
''*******************************************
'' Late 70's to early 80's
'
'Sub InitPolarity()
'        dim x, a : a = Array(LF, RF)
'        for each x in a
'                x.AddPoint "Ycoef", 0, RightFlipper.Y-65, 1        'disabled
'                x.AddPoint "Ycoef", 1, RightFlipper.Y-11, 1
'                x.enabled = True
'                x.TimeDelay = 80
'        Next

'        AddPt "Polarity", 0, 0, 0
'        AddPt "Polarity", 1, 0.05, -2.7
'        AddPt "Polarity", 2, 0.33, -2.7
'        AddPt "Polarity", 3, 0.37, -2.7
'        AddPt "Polarity", 4, 0.41, -2.7
'        AddPt "Polarity", 5, 0.45, -2.7
'        AddPt "Polarity", 6, 0.576,-2.7
'        AddPt "Polarity", 7, 0.66, -1.8
'        AddPt "Polarity", 8, 0.743, -0.5
'        AddPt "Polarity", 9, 0.81, -0.5
'        AddPt "Polarity", 10, 0.88, 0

'        addpt "Velocity", 0, 0,         1
'        addpt "Velocity", 1, 0.16, 1.06
'        addpt "Velocity", 2, 0.41,         1.05
'        addpt "Velocity", 3, 0.53,         1'0.982
'        addpt "Velocity", 4, 0.702, 0.968
'        addpt "Velocity", 5, 0.95,  0.968
'        addpt "Velocity", 6, 1.03,         0.945

'        LF.Object = LeftFlipper
'        LF.EndPoint = EndPointLp
'        RF.Object = RightFlipper
'        RF.EndPoint = EndPointRp
'End Sub

''*******************************************
'' Mid 80's
'
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
    x.AddPt "Polarity", 2, 0.33, - 3.7
    x.AddPt "Polarity", 3, 0.37, - 3.7
    x.AddPt "Polarity", 4, 0.41, - 3.7
    x.AddPt "Polarity", 5, 0.45, - 3.7
    x.AddPt "Polarity", 6, 0.576,- 3.7
    x.AddPt "Polarity", 7, 0.66, - 2.3
    x.AddPt "Polarity", 8, 0.743, - 1.5
    x.AddPt "Polarity", 9, 0.81, - 1
    x.AddPt "Polarity", 10, 0.88, 0

    x.AddPt "Velocity", 0, 0, 1
    x.AddPt "Velocity", 1, 0.16, 1.06
    x.AddPt "Velocity", 2, 0.41, 1.05
    x.AddPt "Velocity", 3, 0.53, 1 '0.982
    x.AddPt "Velocity", 4, 0.702, 0.968
    x.AddPt "Velocity", 5, 0.95,  0.968
    x.AddPt "Velocity", 6, 1.03, 0.945

  Next

' ' SetObjects arguments: 1: name of object 2: flipper object: 3: Trigger object around flipper
    LF.SetObjects "LF", LeftFlipper, TriggerLF
    RF.SetObjects "RF", RightFlipper, TriggerRF
End Sub


'******************************************************
'  FLIPPER CORRECTION FUNCTIONS
'******************************************************

Class FlipperPolarity
  Public DebugOn, Enabled
  Private FlipAt    'Timer variable (IE 'flip at 723,530ms...)
  Public TimeDelay    'delay before trigger turns off and polarity is disabled
  Private Flipper, FlipperStart, FlipperEnd, FlipperEndY, LR, PartialFlipCoef
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
      If Not IsEmpty(balls(x) ) Then
        pos = pSlope(Balls(x).x, FlipperStart, 0, FlipperEnd, 1)
      End If
    Next
  End Property

  Public Sub ProcessBalls() 'save data of balls in flipper range
    FlipAt = GameTime
    Dim x
    For x = 0 To UBound(balls)
      If Not IsEmpty(balls(x) ) Then
        balldata(x).Data = balls(x)
      End If
    Next
    PartialFlipCoef = ((Flipper.StartAngle - Flipper.CurrentAngle) / (Flipper.StartAngle - Flipper.EndAngle))
    PartialFlipCoef = abs(PartialFlipCoef-1)
  End Sub
  'Timer shutoff for polaritycorrect
  Private Function FlipperOn()
    If GameTime < FlipAt+TimeDelay Then
      FlipperOn = True
    End If
  End Function

  Public Sub PolarityCorrect(aBall)
    If FlipperOn() Then
      Dim tmp, BallPos, x, IDX, Ycoef
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
          If ballpos > 0.65 Then  Ycoef = LinearEnvelope(BallData(x).Y, YcoefIn, YcoefOut)                'find safety coefficient 'ycoef' data
        End If
      Next

      If BallPos = 0 Then 'no ball data meaning the ball is entering and exiting pretty close to the same position, use current values.
        BallPos = PSlope(aBall.x, FlipperStart, 0, FlipperEnd, 1)
        If ballpos > 0.65 Then  Ycoef = LinearEnvelope(aBall.Y, YcoefIn, YcoefOut)                        'find safety coefficient 'ycoef' data
      End If

      'Velocity correction
      If Not IsEmpty(VelocityIn(0) ) Then
        Dim VelCoef
        VelCoef = LinearEnvelope(BallPos, VelocityIn, VelocityOut)

        If partialflipcoef < 1 Then VelCoef = PSlope(partialflipcoef, 0, 1, 1, VelCoef)

        If Enabled Then aBall.Velx = aBall.Velx*VelCoef
        If Enabled Then aBall.Vely = aBall.Vely*VelCoef
      End If

      'Polarity Correction (optional now)
      If Not IsEmpty(PolarityIn(0) ) Then
        Dim AddX
        AddX = LinearEnvelope(BallPos, PolarityIn, PolarityOut) * LR

        If Enabled Then aBall.VelX = aBall.VelX + 1 * (AddX*ycoef*PartialFlipcoef)
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
  ' These values are best guesses. Retune them if needed based on specific table research.
  AddSlingsPt 0, 0.00,  -4
  AddSlingsPt 1, 0.45,  -7
  AddSlingsPt 2, 0.48,  0
  AddSlingsPt 3, 0.52,  0
  AddSlingsPt 4, 0.55,  7
  AddSlingsPt 5, 1.00,  4

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

Const EOSTnew = 1 'EM's to late 80's
Const EOSAnew = 1
Const EOSRampup = 0
Dim SOSRampup
Select Case FlipperCoilRampupMode
  Case 0:
    SOSRampup = 2.5
  Case 1:
    SOSRampup = 6
  Case 2:
    SOSRampup = 8.5
End Select

Const LiveCatch = 16
Const LiveElasticity = 0.45
Const SOSEM = 0.815
Const EOSReturn = 0.045  'late 70's to mid 80's

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
    Dim b

    For b = 0 to UBound(gBOT)
      If Distance(gBOT(b).x, gBOT(b).y, Flipper.x, Flipper.y) < 55 Then 'check for cradle
        If gBOT(b).vely >= -0.4 Then gBOT(b).vely = -0.4
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
'****  DROP TARGETS by Rothbauerw
'******************************************************


'******************************************************
'  DROP TARGETS INITIALIZATION
'******************************************************

'Define a variable for each drop target

Dim DT17, DT20, DT21, DT22, DT25, DT26, DT27


'Set array with drop target objects
'
'DropTargetvar = Array(primary, secondary, prim, swtich, animate)
'   primary:      primary target wall to determine drop
' secondary:      wall used to simulate the ball striking a bent or offset target after the initial Hit
' prim:       primitive target used for visuals and animation
'             IMPORTANT!!!
'             rotz must be used for orientation
'             rotx to bend the target back
'             transz to move it up and down
'             the pivot point should be in the center of the target on the x, y and at or below the playfield (0) on z
' switch:       ROM switch number
' animate:      Array slot for handling the animation instrucitons, set to 0
'           Values for animate: 1 - bend target (hit to primary), 2 - drop target (hit to secondary), 3 - brick target (high velocity hit to secondary), -1 - raise target
'   isDropped:      Boolean which determines whether a drop target is dropped. Set to false if they are initially raised, true if initially dropped.

Class DropTarget
  Private m_primary, m_secondary, m_prim, m_sw, m_animate, m_isDropped

  Public Property Get Primary(): Set Primary = m_primary: End Property
  Public Property Let Primary(input): Set m_primary = input: End Property

  Public Property Get Secondary(): Set Secondary = m_secondary: End Property
  Public Property Let Secondary(input): Set m_secondary = input: End Property

  Public Property Get Prim(): Set Prim = m_prim: End Property
  Public Property Let Prim(input): Set m_prim = input: End Property

  Public Property Get Sw(): Sw = m_sw: End Property
  Public Property Let Sw(input): m_sw = input: End Property

  Public Property Get Animate(): Animate = m_animate: End Property
  Public Property Let Animate(input): m_animate = input: End Property

  Public Property Get IsDropped(): IsDropped = m_isDropped: End Property
  Public Property Let IsDropped(input): m_isDropped = input: End Property

  Public default Function init(primary, secondary, prim, sw, animate, isDropped)
    Set m_primary = primary
    Set m_secondary = secondary
    Set m_prim = prim
    m_sw = sw
    m_animate = animate
    m_isDropped = isDropped

    Set Init = Me
  End Function
End Class


Set DT17 = (new DropTarget)(sw17dt, sw17dta, sw17dtp, 17, 0, false)
Set DT20 = (new DropTarget)(sw20, sw20a, sw20p, 20, 0, false)
Set DT21 = (new DropTarget)(sw21, sw21a, sw21p, 21, 0, false)
Set DT22 = (new DropTarget)(sw22, sw22a, sw22p, 22, 0, false)
Set DT25 = (new DropTarget)(sw25, sw25a, sw25p, 25, 0, false)
Set DT26 = (new DropTarget)(sw26, sw26a, sw26p, 26, 0, false)
Set DT27 = (new DropTarget)(sw27, sw27a, sw27p, 27, 0, false)



Dim DTArray
DTArray = Array(DT17, DT20, DT21, DT22, DT25, DT26, DT27)


'Configure the behavior of Drop Targets.
Const DTDropSpeed = 110 'in milliseconds
Const DTDropUpSpeed = 40 'in milliseconds
Const DTDropUnits = 44 'VP units primitive drops so top of at or below the playfield
Const DTDropUpUnits = 10 'VP units primitive raises above the up position on drops up
Const DTMaxBend = 8 'max degrees primitive rotates when hit
Const DTDropDelay = 20 'time in milliseconds before target drops (due to friction/impact of the ball)
Const DTRaiseDelay = 40 'time in milliseconds before target drops back to normal up position after the solenoid fires to raise the target
Const DTBrickVel = 30 'velocity at which the target will brick, set to '0' to disable brick

Const DTEnableBrick = 0 'Set to 0 to disable bricking, 1 to enable bricking
Const Sound = "" 'Drop Target Hit sound
Const DTDropSound = "DropTarget_Down" 'Drop Target Drop sound
Const DTResetSound = "DropTarget_Up" 'Drop Target reset sound

Const DTMass = 0.2 'Mass of the Drop Target (between 0 and 1), higher values provide more resistance


'******************************************************
'  DROP TARGETS FUNCTIONS
'******************************************************


' Initial Drop Target Shadows - Avoids a light DT hit and shadows go off when not strong enough hit to drop the target.
Dim DTShadow(7)

DTShadowInit 1
DTShadowInit 2
DTShadowInit 3
DTShadowInit 4
DTShadowInit 5
DTShadowInit 6
DTShadowInit 7


' Initializes the drop targets for shadow logic below
Sub DTShadowInit(dtnbr)

  if dtnbr = 1 Then
    Set DTShadow(dtnbr) = Eval("dtsh" & 17)
  elseif dtnbr = 2 Then
    Set DTShadow(dtnbr) = Eval("dtsh" & 20)
  elseif dtnbr = 3 Then
    Set DTShadow(dtnbr) = Eval("dtsh" & 21)
  elseif dtnbr = 4 Then
    Set DTShadow(dtnbr) = Eval("dtsh" & 22)
  elseif dtnbr = 5 Then
    Set DTShadow(dtnbr) = Eval("dtsh" & 25)
  elseif dtnbr = 6 Then
    Set DTShadow(dtnbr) = Eval("dtsh" & 26)
  elseif dtnbr = 7 Then
    Set DTShadow(dtnbr) = Eval("dtsh" & 27)

  End If
End Sub

Sub DTHit(switch)
  Dim i, swmod

  i = DTArrayID(switch)
  If switch = 17 Then
    swmod = 1
  Elseif switch = 20 then
    swmod = 2
  Elseif switch = 21 then
    swmod = 3
  Elseif switch = 22 then
    swmod = 4
  Elseif switch = 25 then
    swmod = 5
  Elseif switch = 26 then
    swmod = 6
  Elseif switch = 27 then
    swmod = 7
  End If

  PlayTargetSound
  DTArray(i).animate =  DTCheckBrick(Activeball,DTArray(i).prim)
  If DTArray(i).animate = 1 or DTArray(i).animate = 3 or DTArray(i).animate = 4 Then
    DTBallPhysics Activeball, DTArray(i).prim.rotz, DTMass

' Controls Drop Shadow for a direct hit only
    DTShadow(swmod).visible = 0

  End If
  DoDTAnim
End Sub

Sub DTRaise(switch)
  Dim i
  i = DTArrayID(switch)

  DTArray(i).animate = -1
  DoDTAnim
End Sub

Sub DTDrop(switch)
  Dim i
  i = DTArrayID(switch)

  DTArray(i).animate = 1
  DoDTAnim
End Sub

Function DTArrayID(switch)
  Dim i
  For i = 0 to uBound(DTArray)
    If DTArray(i).sw = switch Then DTArrayID = i:Exit Function
  Next
End Function


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


'Check if target is hit on it's face or sides and whether a 'brick' occurred
Function DTCheckBrick(aBall, dtprim)
  dim bangle, bangleafter, rangle, rangle2, Xintersect, Yintersect, cdist, perpvel, perpvelafter, paravel, paravelafter
  rangle = (dtprim.rotz - 90) * 3.1416 / 180
  rangle2 = dtprim.rotz * 3.1416 / 180
  bangle = atn2(cor.ballvely(aball.id),cor.ballvelx(aball.id))
  bangleafter = Atn2(aBall.vely,aball.velx)

  Xintersect = (aBall.y - dtprim.y - tan(bangle) * aball.x + tan(rangle2) * dtprim.x) / (tan(rangle2) - tan(bangle))
  Yintersect = tan(rangle2) * Xintersect + (dtprim.y - tan(rangle2) * dtprim.x)

  cdist = Distance(dtprim.x, dtprim.y, Xintersect, Yintersect)

  perpvel = cor.BallVel(aball.id) * cos(bangle-rangle)
  paravel = cor.BallVel(aball.id) * sin(bangle-rangle)

  perpvelafter = BallSpeed(aBall) * cos(bangleafter - rangle)
  paravelafter = BallSpeed(aBall) * sin(bangleafter - rangle)

  If perpvel > 0 and  perpvelafter <= 0 Then
    If DTEnableBrick = 1 and  perpvel > DTBrickVel and DTBrickVel <> 0 and cdist < 8 Then
      DTCheckBrick = 3
    Else
      DTCheckBrick = 1
    End If
  ElseIf perpvel > 0 and ((paravel > 0 and paravelafter > 0) or (paravel < 0 and paravelafter < 0)) Then
    DTCheckBrick = 4
  Else
    DTCheckBrick = 0
  End If
End Function


Sub DoDTAnim()
  Dim i
  For i=0 to Ubound(DTArray)
    DTArray(i).animate = DTAnimate(DTArray(i).primary,DTArray(i).secondary,DTArray(i).prim,DTArray(i).sw,DTArray(i).animate)
  Next
End Sub

Function DTAnimate(primary, secondary, prim, switch, animate)
  dim transz, switchid
  Dim animtime, rangle

  switchid = switch

  Dim ind
  ind = DTArrayID(switchid)

  rangle = prim.rotz * PI / 180

  DTAnimate = animate

  if animate = 0  Then
    primary.uservalue = 0
    DTAnimate = 0
    Exit Function
  Elseif primary.uservalue = 0 then
    primary.uservalue = gametime
  end if

  animtime = gametime - primary.uservalue

  If (animate = 1 or animate = 4) and animtime < DTDropDelay Then
    primary.collidable = 0
  If animate = 1 then secondary.collidable = 1 else secondary.collidable= 0
    prim.rotx = DTMaxBend * cos(rangle)
    prim.roty = DTMaxBend * sin(rangle)
    DTAnimate = animate
    Exit Function
    elseif (animate = 1 or animate = 4) and animtime > DTDropDelay Then
    primary.collidable = 0
    If animate = 1 then secondary.collidable = 1 else secondary.collidable= 0
    prim.rotx = DTMaxBend * cos(rangle)
    prim.roty = DTMaxBend * sin(rangle)
    animate = 2
    SoundDropTargetDrop prim
  End If

  if animate = 2 Then

    transz = (animtime - DTDropDelay)/DTDropSpeed *  DTDropUnits * -1
    if prim.transz > -DTDropUnits  Then
      prim.transz = transz
    end if

    prim.rotx = DTMaxBend * cos(rangle)/2
    prim.roty = DTMaxBend * sin(rangle)/2

    if prim.transz <= -DTDropUnits Then
      prim.transz = -DTDropUnits
      secondary.collidable = 0
      DTArray(ind).isDropped = true 'Mark target as dropped
      if UsingROM then

        if Switchid = 17 Then     ' Drop Target Single needs to be pulsed.
          vpmTimer.PulseSw switchid
        Else
          controller.Switch(Switchid) = 1
        end If
      else
        ' do nothing
      end if
      primary.uservalue = 0
      DTAnimate = 0
      Exit Function
    Else
      DTAnimate = 2
      Exit Function
    end If
  End If

  If animate = 3 and animtime < DTDropDelay Then
    primary.collidable = 0
    secondary.collidable = 1
    prim.rotx = DTMaxBend * cos(rangle)
    prim.roty = DTMaxBend * sin(rangle)
  elseif animate = 3 and animtime > DTDropDelay Then
    primary.collidable = 1
    secondary.collidable = 0
    prim.rotx = 0
    prim.roty = 0
    primary.uservalue = 0
    DTAnimate = 0
    Exit Function
  End If

  if animate = -1 Then
    transz = (1 - (animtime)/DTDropUpSpeed) *  DTDropUnits * -1

    If prim.transz = -DTDropUnits Then
      Dim b

      For b = 0 to UBound(gBOT)
        If InRotRect(gBOT(b).x,gBOT(b).y,prim.x, prim.y, prim.rotz, -25,-10,25,-10,25,25,-25,25) and gBOT(b).z < prim.z+DTDropUnits+25 Then
          gBOT(b).velz = 20
        End If
      Next
    End If

    if prim.transz < 0 Then
      prim.transz = transz
    elseif transz > 0 then
      prim.transz = transz
    end if

    if prim.transz > DTDropUpUnits then
      DTAnimate = -2
      prim.transz = DTDropUpUnits
      prim.rotx = 0
      prim.roty = 0
      primary.uservalue = gametime
    end if
    primary.collidable = 0
    secondary.collidable = 1
    DTArray(ind).isDropped = false 'Mark target as not dropped
    if UsingROM then

      if switchid = 17 then
        'do Nothing, as it was pulsed up above
      Else
        controller.Switch(Switchid) = 0
      end If
    end If
  End If

  if animate = -2 and animtime > DTRaiseDelay Then
    prim.transz = (animtime - DTRaiseDelay)/DTDropSpeed *  DTDropUnits * -1 + DTDropUpUnits
    if prim.transz < 0 then
      prim.transz = 0
      primary.uservalue = 0
      DTAnimate = 0

      primary.collidable = 1
      secondary.collidable = 0
    end If
  End If
End Function


'******************************************************
'  DROP TARGET
'  SUPPORTING FUNCTIONS
'******************************************************


' Used for drop targets
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

Function InRotRect(ballx,bally,px,py,angle,ax,ay,bx,by,cx,cy,dx,dy)
    Dim rax,ray,rbx,rby,rcx,rcy,rdx,rdy
    Dim rotxy
    rotxy = RotPoint(ax,ay,angle)
    rax = rotxy(0)+px : ray = rotxy(1)+py
    rotxy = RotPoint(bx,by,angle)
    rbx = rotxy(0)+px : rby = rotxy(1)+py
    rotxy = RotPoint(cx,cy,angle)
    rcx = rotxy(0)+px : rcy = rotxy(1)+py
    rotxy = RotPoint(dx,dy,angle)
    rdx = rotxy(0)+px : rdy = rotxy(1)+py

    InRotRect = InRect(ballx,bally,rax,ray,rbx,rby,rcx,rcy,rdx,rdy)
End Function

Function RotPoint(x,y,angle)
    dim rx, ry
    rx = x*dCos(angle) - y*dSin(angle)
    ry = x*dSin(angle) + y*dCos(angle)
    RotPoint = Array(rx,ry)
End Function


'******************************************************
'****  END DROP TARGETS
'******************************************************

'******************************************************
'   STAND-UP TARGET INITIALIZATION
'******************************************************

'Define a variable for each stand-up target

Dim ST17, ST28, ST29, ST30, ST31, ST32, ST33, ST34

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


Set ST17 = (new StandupTarget)(sw17, psw17,17, 0)
Set ST28 = (new StandupTarget)(sw28, psw28,28, 0)
Set ST29 = (new StandupTarget)(sw29, psw29,29, 0)
Set ST30 = (new StandupTarget)(sw30, psw30,30, 0)
Set ST31 = (new StandupTarget)(sw31, psw31,31, 0)
Set ST32 = (new StandupTarget)(sw32, psw32,32, 0)
Set ST33 = (new StandupTarget)(sw33, psw33,33, 0)
Set ST34 = (new StandupTarget)(sw34, psw34,34, 0)


Dim STArray
STArray = Array(ST17, ST28, ST29, ST30, ST31, ST32, ST33, ST34)


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
  STArray(i).animate =  STCheckHit(Activeball,STArray(i).primary)

  If STArray(i).animate <> 0 Then
    DTBallPhysics Activeball, STArray(i).primary.orientation, STMass
  End If
  DoSTAnim
End Sub

Function STArrayID(switch)
  Dim i
  For i = 0 to uBound(STArray)
    If STArray(i).sw = switch Then STArrayID = i:Exit Function
  Next
End Function

'Check if target is hit on it's face
Function STCheckHit(aBall, target)
  dim bangle, bangleafter, rangle, rangle2, perpvel, perpvelafter, paravel, paravelafter
  rangle = (target.orientation - 90) * 3.1416 / 180
  bangle = atn2(cor.ballvely(aball.id),cor.ballvelx(aball.id))
  bangleafter = Atn2(aBall.vely,aball.velx)

  perpvel = cor.BallVel(aball.id) * cos(bangle-rangle)
  paravel = cor.BallVel(aball.id) * sin(bangle-rangle)

  perpvelafter = BallSpeed(aBall) * cos(bangleafter - rangle)
  paravelafter = BallSpeed(aBall) * sin(bangleafter - rangle)

  If perpvel > 0 and  perpvelafter <= 0 Then
    STCheckHit = 1
  ElseIf perpvel > 0 and ((paravel > 0 and paravelafter > 0) or (paravel < 0 and paravelafter < 0)) Then
    STCheckHit = 1
  Else
    STCheckHit = 0
  End If
End Function

Sub DoSTAnim()
  Dim i
  For i=0 to Ubound(STArray)
    STArray(i).animate = STAnimate(STArray(i).primary,STArray(i).prim,STArray(i).sw,STArray(i).animate)
  Next
End Sub

Function STAnimate(primary, prim, switch,  animate)
  Dim animtime

  STAnimate = animate

  if animate = 0  Then
    primary.uservalue = 0
    STAnimate = 0
    Exit Function
  Elseif primary.uservalue = 0 then
    primary.uservalue = gametime
  end if

  animtime = gametime - primary.uservalue

  If animate = 1 Then
    primary.collidable = 0
    prim.transy = -STMaxOffset
    if UsingROM then

      vpmTimer.PulseSw switch


  'manually raise the single drop target
    If switch = 17 AND DT17up = 0 then
      vpmTimer.AddTimer DTRDelayTime, "RaiseSingleDT'"    'slight delay to ensure ball has time to move before target raise
    End if

    else
      STAction switch
    end if
    STAnimate = 2
    Exit Function
  elseif animate = 2 Then
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
  Select Case Switch
    Case 11:
      Addscore 1000
      Flash1 True               'Demo of the flasher
      vpmTimer.AddTimer 150,"Flash1 False'" 'Disable the flash after short time, just like a ROM would do
    Case 12:
      Addscore 1000
      Flash2 True               'Demo of the flasher
      vpmTimer.AddTimer 150,"Flash2 False'" 'Disable the flash after short time, just like a ROM would do
    Case 13:
      Addscore 1000
      Flash3 True               'Demo of the flasher
      vpmTimer.AddTimer 150,"Flash3 False'" 'Disable the flash after short time, just like a ROM would do
  End Select
End Sub

'******************************************************
'   END STAND-UP TARGETS
'******************************************************



'******************************************************
'****  BALL ROLLING AND DROP SOUNDS
'******************************************************
'
' Be sure to call RollingUpdate in a timer with a 10ms interval see the GameTimer_Timer() sub

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
  Dim b

  ' stop the sound of deleted balls
  For b = UBound(gBOT) + 1 to tnob
    ' Comment the next line if you are not implementing Dyanmic Ball Shadows
    If AmbientBallShadowOn = 0 Then BallShadowA(b).visible = 0
    rolling(b) = False
    StopSound("BallRoll_" & b)
  Next

  ' exit the sub if no balls on the table
  If UBound(gBOT) = -1 Then Exit Sub

  ' play the rolling sound for each ball

  For b = 0 to UBound(gBOT)
    If BallVel(gBOT(b)) > 1 AND gBOT(b).z < 30 Then
      rolling(b) = True
      PlaySound ("BallRoll_" & b), -1, VolPlayfieldRoll(gBOT(b)) * BallRollVolume * VolumeDial, AudioPan(gBOT(b)), 0, PitchPlayfieldRoll(gBOT(b)), 1, 0, AudioFade(gBOT(b))
    ElseIf BallVel(gBOT(b)) > 1 AND gBOT(b).z > 70 Then
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


'   ' "Static" Ball Shadows
    If AmbientBallShadowOn = 0 Then
      BallShadowA(b).visible = 1
      BallShadowA(b).X = gBOT(b).X + offsetX
      If gBOT(b).Z > 30 Then
        BallShadowA(b).height=gBOT(b).z - BallSize/4 + b/1000 'This is technically 1/4 of the ball "above" the ramp, but it keeps it from clipping the ramp
        BallShadowA(b).Y = gBOT(b).Y + offsetY + BallSize/10
      Else
        BallShadowA(b).height=gBOT(b).z - BallSize/2 + 1.04 + b/1000
        BallShadowA(b).Y = gBOT(b).Y + offsetY
      End If
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
  PlaySoundAtLevelStatic SoundFX("Flipper_Attack-L01",DOFFlippers), FlipperUpAttackLeftSoundLevel, flipper
End Sub

Sub SoundFlipperUpAttackRight(flipper)
  FlipperUpAttackRightSoundLevel = RndNum(FlipperUpAttackMinimumSoundLevel, FlipperUpAttackMaximumSoundLevel)
  PlaySoundAtLevelStatic SoundFX("Flipper_Attack-R01",DOFFlippers), FlipperUpAttackLeftSoundLevel, flipper
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

  FlipperCradleCollision ball1, ball2, velocity

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



'******************************************************
'******  FLUPPER BUMPERS
'******************************************************

' prepare some global vars to dim/brighten objects when using day-night slider
Dim DayNightAdjust , DNA30, DNA45, DNA90
If NightDay < 10 Then
  DNA30 = 0 : DNA45 = (NightDay-10)/20 : DNA90 = 0 : DayNightAdjust = 0.4
Else
  DNA30 = (NightDay-10)/30 : DNA45 = (NightDay-10)/45 : DNA90 = (NightDay-10)/90 : DayNightAdjust = NightDay/25
End If

Dim FlBumperFadeActual(6), FlBumperFadeTarget(6), FlBumperColor(6), FlBumperTop(6), FlBumperSmallLight(6), Flbumperbiglight(6)
Dim FlBumperDisk(6), FlBumperBase(6), FlBumperBulb(6), FlBumperscrews(6), FlBumperActive(6), FlBumperHighlight(6)
Dim cnt : For cnt = 1 to 6 : FlBumperActive(cnt) = False : Next

' colors available are red, white, blue, orange, yellow, green, purple and blacklight
' NOTE:  created a special color only for barracora table to work with the top, and colors that matched real table.

FlInitBumper 1, "blue"
FlInitBumper 2, "blue"
FlInitBumper 3, "blue"
FlInitBumper 4, "blue"

Sub FlInitBumper(nr, col)
  FlBumperActive(nr) = True
  ' store all objects in an array for use in FlFadeBumper subroutine
  FlBumperFadeActual(nr) = 1 : FlBumperFadeTarget(nr) = 1.1: FlBumperColor(nr) = col
  Set FlBumperTop(nr) = Eval("bumpertop" & nr) : FlBumperTop(nr).material = "bumpertopmat" & nr
  Set FlBumperSmallLight(nr) = Eval("bumpersmalllight" & nr) : Set Flbumperbiglight(nr) = Eval("bumperbiglight" & nr)
  Set FlBumperDisk(nr) = Eval("bumperdisk" & nr) : Set FlBumperBase(nr) = Eval("bumperbase" & nr)
  Set FlBumperBulb(nr) = Eval("bumperbulb" & nr) : FlBumperBulb(nr).material = "bumperbulbmat" & nr
  Set FlBumperscrews(nr) = Eval("bumperscrews" & nr): FlBumperscrews(nr).material = "bumperscrew" & col
  Set FlBumperHighlight(nr) = Eval("bumperhighlight" & nr)
  ' set the color for the two VPX lights
  select case col
    Case "blue"
      FlBumperBigLight(nr).color = RGB(32,80,255)
      FlBumperBigLight(nr).colorfull = RGB(32,80,255)
      FlBumperSmallLight(nr).color = RGB(0,80,255)
      FlBumperSmallLight(nr).colorfull = RGB(0,80,255)
      FlBumperSmallLight(nr).TransmissionScale = 0
      MaterialColor "bumpertopmat" & nr, RGB(8,120,255)
      FlBumperHighlight(nr).color = RGB(255,16,8)
      FlBumperSmallLight(nr).BulbModulateVsAdd = 1
  end select
End Sub


Sub FlFadeBumper(nr, Z)

  FlBumperBase(nr).BlendDisableLighting = 0.5 * DayNightAdjust

' UpdateMaterial(string, float wrapLighting, float roughness, float glossyImageLerp, float thickness, float edge, float edgeAlpha, float opacity,
'               OLE_COLOR base, OLE_COLOR glossy, OLE_COLOR clearcoat, VARIANT_BOOL isMetal, VARIANT_BOOL opacityActive,
'               float elasticity, float elasticityFalloff, float friction, float scatterAngle) - updates all parameters of a material
  FlBumperDisk(nr).BlendDisableLighting = (0.5 - Z * 0.3 )* DayNightAdjust

  select case FlBumperColor(nr)
    Case "blue"
      UpdateMaterial "bumperbulbmat" & nr, 0, 0.75 , 0, 1 - Z, 1 - Z, 1 - Z, 0.9999, RGB(38 - 24 * Z,130 - 98 * Z,255), RGB(255,255,255), RGB(32,32,32), False, True, 0, 0, 0, 0
      FlBumperSmallLight(nr).intensity = bumperscale*(20 + 500 * Z / (0.5 + DNA30))
      FlBumperTop(nr).BlendDisableLighting = bumperscale*(3 * DayNightAdjust + 50 * Z)
      FlBumperBulb(nr).BlendDisableLighting = bumperscale*(12 * DayNightAdjust + 5000 * (0.03 * Z + 0.97 * Z ^ 3))
      Flbumperbiglight(nr).intensity = bumperscale*(25 * Z / (1 + DNA45))
      FlBumperHighlight(nr).opacity = bumperscale*(10000 * (Z ^ 3) / (0.5 + DNA90))
  end select
End Sub

Sub BumperTimer_Timer
  dim nr
  For nr = 1 to 6
    If FlBumperFadeActual(nr) < FlBumperFadeTarget(nr) and FlBumperActive(nr)  Then
      FlBumperFadeActual(nr) = FlBumperFadeActual(nr) + (FlBumperFadeTarget(nr) - FlBumperFadeActual(nr)) * 0.8
      If FlBumperFadeActual(nr) > 0.99 Then FlBumperFadeActual(nr) = 1 : End If
      FlFadeBumper nr, FlBumperFadeActual(nr)
    End If
    If FlBumperFadeActual(nr) > FlBumperFadeTarget(nr) and FlBumperActive(nr)  Then
      FlBumperFadeActual(nr) = FlBumperFadeActual(nr) + (FlBumperFadeTarget(nr) - FlBumperFadeActual(nr)) * 0.4 / (FlBumperFadeActual(nr) + 0.1)
      If FlBumperFadeActual(nr) < 0.01 Then FlBumperFadeActual(nr) = 0 : End If
      FlFadeBumper nr, FlBumperFadeActual(nr)
    End If
  next
End Sub


'******************************************************
'******  END FLUPPER BUMPERS
'******************************************************




'******************************************************
'****  LAMPZ by nFozzy
'******************************************************
'

Sub SetLamp(aNr, aOn)
  Lampz.state(aNr) = abs(aOn)
End Sub

Dim NullFader : set NullFader = new NullFadingObject
Dim Lampz : Set Lampz = New LampFader
InitLampsNF               ' Setup lamp assignments

Sub LampTimer()
  dim x, chglamp
  if UsingROM then chglamp = Controller.ChangedLamps
  If Not IsEmpty(chglamp) Then
    For x = 0 To UBound(chglamp)      'nmbr = chglamp(x, 0), state = chglamp(x, 1)
      Lampz.state(chglamp(x, 0)) = chglamp(x, 1)
    next
  End If
  Lampz.Update2 'update (fading logic only)

End Sub


Sub DisableLighting(pri, DLintensity, ByVal aLvl) 'cp's script  DLintensity = disabled lighting intesity
  if Lampz.UseFunction then aLvl = Lampz.FilterOut(aLvl)  'Callbacks don't get this filter automatically
  pri.blenddisablelighting = aLvl * DLintensity
End Sub

sub DisableLighting2(pri, DLintensityMax, DLintensityMin, ByVal aLvl)    'cp's script  DLintensity = disabled lighting intesity
    if Lampz.UseFunction then aLvl = Lampz.FilterOut(aLvl)    'Callbacks don't get this filter automatically
    pri.blenddisablelighting = (aLvl * (DLintensityMax-DLintensityMin)) + DLintensityMin
End Sub

Sub SetModLamp(id, val)
  Lampz.state(id) = val
End Sub

const insert_dl_on_red = 60
const insert_dl_on_green = 80
const insert_dl_on_orange = 80
const insert_dl_on_white = 60
const insert_dl_on_greenarrow = 100
const insert_dl_on_orangearrow = 50
const insert_dl_on_yellowarrow = 30
const insert_dl_on_redarrow = 60
const insert_dl_on_whitearrow = 200
const insert_dl_on_blue = 200
const insert_dl_on_yellow = 80
const smallinsert_max = .4
const smallinsert_min = .01
const greenbulb = 300
const bluebulb = 200
const redbulb = 400
const orangebulb = 300
const yellowbulb = 400
const whitebulb = 100
const insert_dl_on_redfl = 30

Sub InitLampsNF()

  'Filtering (comment out to disable)
  Lampz.Filter = "LampFilter" 'Puts all lamp intensityscale output (no callbacks) through this function before updating

  'Adjust fading speeds (max level / full MS fading time). The Modulate property must be set to 1 / max level if lamp is modulated.
  dim x : for x = 0 to 150 : Lampz.FadeSpeedUp(x) = 1/40 : Lampz.FadeSpeedDown(x) = 1/120 : Lampz.Modulate(x) = 1 : next

' GI Fading
  Lampz.FadeSpeedUp(140) = 1/40 : Lampz.FadeSpeedDown(140) = 1/16 : Lampz.Modulate(140) = 1
  Lampz.FadeSpeedUp(141) = 1/40 : Lampz.FadeSpeedDown(141) = 1/16 : Lampz.Modulate(141) = 1

' Flipper Light Fading
  Lampz.FadeSpeedUp(119) = 1/60 : Lampz.FadeSpeedDown(119) = 1/360 : Lampz.Modulate(119) = 1


  'Lampz Assignments
  '  In a ROM based table, the lamp ID is used to set the state of the Lampz objects

  'MassAssign is an optional way to do assignments. It'll create arrays automatically / append objects to existing arrays


'Star Targets
  Lampz.MassAssign(6)= l2atrigger
  Lampz.Callback(6) = "DisableLighting2 pstar2a, 2, 0.2,"
  Lampz.MassAssign(22)= l2btrigger
  Lampz.Callback(22) = "DisableLighting2 pstar2b, 2, 0.2,"
  Lampz.MassAssign(38)= l2ctrigger
  Lampz.Callback(38) = "DisableLighting2 pstar2c, 2, 0.2,"


'Playfield

  Lampz.MassAssign(1)= l1
  Lampz.MassAssign(1)= l1a
  Lampz.Callback(1) = "DisableLighting2 p1, insert_dl_on_orange, 1,"
  Lampz.Callback(1) = "DisableLighting2 bulb1, orangebulb, 5,"
  Lampz.MassAssign(2)= l2
  Lampz.MassAssign(2)= l2a
  Lampz.Callback(2) = "DisableLighting2 p2, insert_dl_on_orange, 1,"
  Lampz.Callback(2) = "DisableLighting2 bulb2, orangebulb, 5,"
  Lampz.MassAssign(3)= l3
  Lampz.MassAssign(3)= l3a
  Lampz.Callback(3) = "DisableLighting2 p3, insert_dl_on_orange, 1,"
  Lampz.Callback(3) = "DisableLighting2 bulb3, orangebulb, 5,"
  Lampz.MassAssign(4)= l4
  Lampz.MassAssign(4)= l4a
  Lampz.Callback(4) = "DisableLighting2 p4, insert_dl_on_green, 1,"
  Lampz.Callback(4) = "DisableLighting2 bulb4, greenbulb, 5,"
  Lampz.MassAssign(5)= l5
  Lampz.MassAssign(5)= l5a
  Lampz.Callback(5) = "DisableLighting2 p5, insert_dl_on_green, 1,"
  Lampz.Callback(5) = "DisableLighting2 bulb5, greenbulb, 5,"
  Lampz.MassAssign(7)= l7
  Lampz.MassAssign(7)= l7a
  Lampz.Callback(7) = "DisableLighting2 p7, insert_dl_on_blue, 1,"
  Lampz.Callback(7) = "DisableLighting2 bulb7, bluebulb, 5,"
  Lampz.MassAssign(8)= l8
  Lampz.MassAssign(8)= l8a
  Lampz.Callback(8) = "DisableLighting2 p8, insert_dl_on_blue, 1,"
  Lampz.Callback(8) = "DisableLighting2 bulb8, bluebulb, 5,"
  Lampz.MassAssign(9)= l9
  Lampz.MassAssign(9)= l9a
  Lampz.Callback(9) = "DisableLighting2 p9, insert_dl_on_green, 1,"
  Lampz.Callback(9) = "DisableLighting2 bulb9, greenbulb, 5,"
  Lampz.MassAssign(10)= l10
  Lampz.MassAssign(10)= l10a
  Lampz.Callback(10) = "DisableLighting2 p10, insert_dl_on_yellow, 1,"
  Lampz.Callback(10) = "DisableLighting2 bulb10, yellowbulb, 5,"
  Lampz.Callback(10) = "DisableLighting2 p10off, .1, 3,"
  Lampz.MassAssign(17)= l17
  Lampz.MassAssign(17)= l17a
  Lampz.Callback(17) = "DisableLighting2 p17, insert_dl_on_orange, 1,"
  Lampz.Callback(17) = "DisableLighting2 bulb17, orangebulb, 5,"
  Lampz.MassAssign(18)= l18
  Lampz.MassAssign(18)= l18a
  Lampz.Callback(18) = "DisableLighting2 p18, insert_dl_on_orange, 1,"
  Lampz.Callback(18) = "DisableLighting2 bulb18, orangebulb, 5,"
  Lampz.MassAssign(19)= l19
  Lampz.MassAssign(19)= l19a
  Lampz.Callback(19) = "DisableLighting2 p19, insert_dl_on_orange, 1,"
  Lampz.Callback(19) = "DisableLighting2 bulb19, orangebulb, 5,"
  Lampz.MassAssign(20)= l20
  Lampz.MassAssign(20)= l20a
  Lampz.Callback(20) = "DisableLighting2 p20, insert_dl_on_green, 1,"
  Lampz.Callback(20) = "DisableLighting2 bulb20, greenbulb, 5,"
  Lampz.MassAssign(21)= l21
  Lampz.MassAssign(21)= l21a
  Lampz.Callback(21) = "DisableLighting2 p21, insert_dl_on_green, 1,"
  Lampz.Callback(21) = "DisableLighting2 bulb21, greenbulb, 5,"
  Lampz.MassAssign(23)= l23
  Lampz.MassAssign(23)= l23a
  Lampz.Callback(23) = "DisableLighting2 p23, insert_dl_on_blue, 1,"
  Lampz.Callback(23) = "DisableLighting2 bulb23, bluebulb, 5,"
  Lampz.MassAssign(24)= l24
  Lampz.MassAssign(24)= l24a
  Lampz.Callback(24) = "DisableLighting2 p24, insert_dl_on_blue, 1,"
  Lampz.Callback(24) = "DisableLighting2 bulb24, bluebulb, 5,"
  Lampz.MassAssign(25)= l25
  Lampz.MassAssign(25)= l25a
  Lampz.Callback(25) = "DisableLighting2 p25, insert_dl_on_green, 1,"
  Lampz.Callback(25) = "DisableLighting2 bulb25, greenbulb, 5,"
  Lampz.MassAssign(26)= l26
  Lampz.MassAssign(26)= l26a
  Lampz.Callback(26) = "DisableLighting2 p26, insert_dl_on_yellow, 1,"
  Lampz.Callback(26) = "DisableLighting2 bulb26, yellowbulb, 5,"
  Lampz.Callback(26) = "DisableLighting2 p26off, .1, 3,"
  Lampz.MassAssign(33)= l33
  Lampz.MassAssign(33)= l33a
  Lampz.Callback(33) = "DisableLighting2 p33, insert_dl_on_orange, 1,"
  Lampz.Callback(33) = "DisableLighting2 bulb33, orangebulb, 5,"
  Lampz.MassAssign(34)= l34
  Lampz.MassAssign(34)= l34a
  Lampz.Callback(34) = "DisableLighting2 p34, insert_dl_on_orange, 1,"
  Lampz.Callback(34) = "DisableLighting2 bulb34, orangebulb, 5,"
  Lampz.MassAssign(35)= l35
  Lampz.MassAssign(35)= l35a
  Lampz.Callback(35) = "DisableLighting2 p35, insert_dl_on_red, 1,"
  Lampz.Callback(35) = "DisableLighting2 bulb35, redbulb, 5,"
  Lampz.MassAssign(36)= l36
  Lampz.MassAssign(36)= l36a
  Lampz.Callback(36) = "DisableLighting2 p36, insert_dl_on_red, 1,"
  Lampz.Callback(36) = "DisableLighting2 bulb36, redbulb, 5,"
  Lampz.MassAssign(37)= l37
  Lampz.MassAssign(37)= l37a
  Lampz.Callback(37) = "DisableLighting2 p37, insert_dl_on_green, 1,"
  Lampz.Callback(37) = "DisableLighting2 bulb37, greenbulb, 5,"
  Lampz.MassAssign(39)= l39
  Lampz.MassAssign(39)= l39a
  Lampz.Callback(39) = "DisableLighting2 p39, insert_dl_on_blue, 1,"
  Lampz.Callback(39) = "DisableLighting2 bulb39, bluebulb, 5,"
  Lampz.MassAssign(40)= l40
  Lampz.MassAssign(40)= l40a
  Lampz.Callback(40) = "DisableLighting2 p40, insert_dl_on_blue, 1,"
  Lampz.Callback(40) = "DisableLighting2 bulb40, bluebulb, 5,"
  Lampz.MassAssign(41)= l41
  Lampz.MassAssign(41)= l41a
  Lampz.Callback(41) = "DisableLighting2 p41, insert_dl_on_green, 1,"
  Lampz.Callback(41) = "DisableLighting2 bulb41, greenbulb, 5,"
  Lampz.MassAssign(42)= l42
  Lampz.MassAssign(42)= l42a
  Lampz.Callback(42) = "DisableLighting2 p42, insert_dl_on_orange, 1,"
  Lampz.Callback(42) = "DisableLighting2 bulb42, orangebulb, 5,"
  Lampz.Callback(42) = "DisableLighting2 p42off, .1, 3,"
  Lampz.MassAssign(43)= l43
  Lampz.MassAssign(43)= l43a
  Lampz.Callback(43) = "DisableLighting2 p43, insert_dl_on_orange, 1,"
  Lampz.Callback(43) = "DisableLighting2 bulb43, orangebulb, 5,"
  Lampz.MassAssign(49)= l49
  Lampz.MassAssign(49)= l49a
  Lampz.Callback(49) = "DisableLighting2 p49, insert_dl_on_orange, 1,"
  Lampz.Callback(49) = "DisableLighting2 bulb49, orangebulb, 5,"
  Lampz.MassAssign(50)= l50
  Lampz.MassAssign(50)= l50a
  Lampz.Callback(50) = "DisableLighting2 p50, insert_dl_on_orange, 1,"
  Lampz.Callback(50) = "DisableLighting2 bulb50, orangebulb, 5,"
  Lampz.MassAssign(51)= l51
  Lampz.MassAssign(51)= l51a
  Lampz.Callback(51) = "DisableLighting2 p51, insert_dl_on_red, 1,"
  Lampz.Callback(51) = "DisableLighting2 bulb51, redbulb, 5,"
  Lampz.MassAssign(53)= l53
  Lampz.MassAssign(53)= l53a
  Lampz.Callback(53) = "DisableLighting2 p53, insert_dl_on_green, 1,"
  Lampz.Callback(53) = "DisableLighting2 bulb53, greenbulb, 5,"
  Lampz.MassAssign(54)= l54
  Lampz.MassAssign(54)= l54a
  Lampz.Callback(54) = "DisableLighting2 p54, insert_dl_on_green, 1,"
  Lampz.Callback(54) = "DisableLighting2 bulb54, greenbulb, 5,"
  Lampz.Callback(54) = "DisableLighting2 p54off, .1, 3,"
  Lampz.MassAssign(55)= l55
  Lampz.MassAssign(55)= l55a
  Lampz.Callback(55) = "DisableLighting2 p55, insert_dl_on_blue, 1,"
  Lampz.Callback(55) = "DisableLighting2 bulb55, bluebulb, 5,"
  Lampz.MassAssign(57)= l57
  Lampz.MassAssign(57)= l57a
  Lampz.Callback(57) = "DisableLighting2 p57, insert_dl_on_green, 1,"
  Lampz.Callback(57) = "DisableLighting2 bulb57, greenbulb, 5,"
  Lampz.MassAssign(58)= l58
  Lampz.MassAssign(58)= l58a
  Lampz.Callback(58) = "DisableLighting2 p58, insert_dl_on_red, 1,"
  Lampz.Callback(58) = "DisableLighting2 bulb58, redbulb, 5,"
  Lampz.Callback(58) = "DisableLighting2 p58off, .1, 3,"
  Lampz.MassAssign(65)= l65
  Lampz.MassAssign(65)= l65a
  Lampz.Callback(65) = "DisableLighting2 p65, insert_dl_on_white, 1,"
  Lampz.Callback(65) = "DisableLighting2 bulb65, whitebulb, 5,"
  Lampz.Callback(65) = "DisableLighting2 p65off, .1, 10,"
  Lampz.MassAssign(66)= l66
  Lampz.MassAssign(66)= l66a
  Lampz.Callback(66) = "DisableLighting2 p66, insert_dl_on_white, 1,"
  Lampz.Callback(66) = "DisableLighting2 bulb66, whitebulb, 5,"
  Lampz.Callback(66) = "DisableLighting2 p66off, .1, 10,"
  Lampz.MassAssign(67)= l67
  Lampz.MassAssign(67)= l67a
  Lampz.Callback(67) = "DisableLighting2 p67, insert_dl_on_green, 1,"
  Lampz.Callback(67) = "DisableLighting2 bulb67, greenbulb, 5,"
  Lampz.MassAssign(68)= l68
  Lampz.MassAssign(68)= l68a
  Lampz.Callback(68) = "DisableLighting2 p68, insert_dl_on_yellow, 1,"
  Lampz.Callback(68) = "DisableLighting2 bulb68, yellowbulb, 5,"
  Lampz.MassAssign(69)= l69
  Lampz.MassAssign(69)= l69a
  Lampz.Callback(69) = "DisableLighting2 p69, insert_dl_on_blue, 1,"
  Lampz.Callback(69) = "DisableLighting2 bulb69, bluebulb, 5,"
  Lampz.MassAssign(71)= l71
  Lampz.MassAssign(71)= l71a
  Lampz.Callback(71) = "DisableLighting2 p71, insert_dl_on_whitearrow, 1,"
  Lampz.Callback(71) = "DisableLighting2 bulb71, whitebulb, 5,"
  Lampz.Callback(71) = "DisableLighting2 p71off, .1, 10,"
  Lampz.MassAssign(81)= l81
  Lampz.MassAssign(81)= l81a
  Lampz.Callback(81) = "DisableLighting2 p81, insert_dl_on_white, 1,"
  Lampz.Callback(81) = "DisableLighting2 bulb81, whitebulb, 5,"
  Lampz.Callback(81) = "DisableLighting2 p81off, .1, 10,"
  Lampz.MassAssign(82)= l82
  Lampz.MassAssign(82)= l82a
  Lampz.Callback(82) = "DisableLighting2 p82, insert_dl_on_white, 1,"
  Lampz.Callback(82) = "DisableLighting2 bulb82, whitebulb, 5,"
  Lampz.Callback(82) = "DisableLighting2 p82off, .1, 10,"
  Lampz.MassAssign(83)= l83
  Lampz.MassAssign(83)= l83a
  Lampz.Callback(83) = "DisableLighting2 p83, insert_dl_on_green, 1,"
  Lampz.Callback(83) = "DisableLighting2 bulb83, greenbulb, 5,"
  Lampz.MassAssign(84)= l84
  Lampz.MassAssign(84)= l84a
  Lampz.Callback(84) = "DisableLighting2 p84, insert_dl_on_green, 1,"
  Lampz.Callback(84) = "DisableLighting2 bulb84, greenbulb, 5,"
  Lampz.MassAssign(85)= l85
  Lampz.MassAssign(85)= l85a
  Lampz.Callback(85) = "DisableLighting2 p85, insert_dl_on_blue, 1,"
  Lampz.Callback(85) = "DisableLighting2 bulb85, bluebulb, 5,"
  Lampz.Callback(87) = "TopMidLaneLights"
  Lampz.MassAssign(97)= l97
  Lampz.MassAssign(97)= l97a
  Lampz.Callback(97) = "DisableLighting2 p97, insert_dl_on_green, 1,"
  Lampz.Callback(97) = "DisableLighting2 bulb97, greenbulb, 5,"
  Lampz.MassAssign(98)= l98
  Lampz.MassAssign(98)= l98a
  Lampz.Callback(98) = "DisableLighting2 p98, insert_dl_on_orange, 1,"
  Lampz.Callback(98) = "DisableLighting2 bulb98, orangebulb, 5,"
  Lampz.MassAssign(99)= l99
  Lampz.MassAssign(99)= l99a
  Lampz.Callback(99) = "DisableLighting2 p99, insert_dl_on_green, 1,"
  Lampz.Callback(99) = "DisableLighting2 bulb99, greenbulb, 5,"
  Lampz.MassAssign(100)= l100
  Lampz.MassAssign(100)= l100a
  Lampz.Callback(100) = "DisableLighting2 p100, insert_dl_on_orange, 1,"
  Lampz.Callback(100) = "DisableLighting2 bulb100, orangebulb, 5,"
  Lampz.MassAssign(101)= l101
  Lampz.MassAssign(101)= l101a
  Lampz.Callback(101) = "DisableLighting2 p101, insert_dl_on_blue, 1,"
  Lampz.Callback(101) = "DisableLighting2 bulb101, bluebulb, 5,"
  Lampz.MassAssign(102)= l102
  Lampz.MassAssign(102)= l102a
  Lampz.Callback(102) = "DisableLighting2 p102, insert_dl_on_orange, 1,"
  Lampz.Callback(102) = "DisableLighting2 bulb102, orangebulb, 5,"
  Lampz.MassAssign(113)= l113
  Lampz.MassAssign(113)= l113a
  Lampz.Callback(113) = "DisableLighting2 p113, insert_dl_on_red, 1,"
  Lampz.Callback(113) = "DisableLighting2 bulb113, redbulb, 5,"
  Lampz.Callback(113) = "DisableLighting2 p113off, .1, 3,"
  Lampz.MassAssign(114)= l114
  Lampz.MassAssign(114)= l114a
  Lampz.Callback(114) = "DisableLighting2 p114, insert_dl_on_red, 1,"
  Lampz.Callback(114) = "DisableLighting2 bulb114, redbulb, 5,"
  Lampz.MassAssign(115)= l115
  Lampz.MassAssign(115)= l115a
  Lampz.Callback(115) = "DisableLighting2 p115, insert_dl_on_green, 1,"
  Lampz.Callback(115) = "DisableLighting2 bulb115, greenbulb, 5,"
  Lampz.MassAssign(116)= l116
  Lampz.MassAssign(116)= l116a
  Lampz.Callback(116) = "DisableLighting2 p116, insert_dl_on_red, 1,"
  Lampz.Callback(116) = "DisableLighting2 bulb116, redbulb, 5,"
  Lampz.MassAssign(118)= l118
  Lampz.MassAssign(118)= l118a
  Lampz.Callback(118) = "DisableLighting2 p118, insert_dl_on_green, 1,"
  Lampz.Callback(118) = "DisableLighting2 bulb118, greenbulb, 5,"
  Lampz.MassAssign(119)= l119
  Lampz.MassAssign(119)= l119a
  Lampz.Callback(119) = "DisableLighting2 p119, insert_dl_on_red, 1,"
  Lampz.Callback(119) = "DisableLighting2 bulb119, redbulb, 5,"



' Bumpers
  Lampz.MassAssign(12)= bumpersmalllight1
  Lampz.MassAssign(12)= bumperbiglight1
  Lampz.MassAssign(12)= bumpershadow1
  Lampz.MassAssign(12)= bumperhighlight1
  Lampz.Callback(12) = "bumperbulb1flash"
  Lampz.MassAssign(28)= bumpersmalllight2
  Lampz.MassAssign(28)= bumperbiglight2
  Lampz.MassAssign(28)= bumpershadow2
  Lampz.MassAssign(28)= bumperhighlight2
  Lampz.Callback(28) = "bumperbulb2flash"
  Lampz.MassAssign(44)= bumpersmalllight3
  Lampz.MassAssign(44)= bumperbiglight3
  Lampz.MassAssign(44)= bumpershadow3
  Lampz.MassAssign(44)= bumperhighlight3
  Lampz.Callback(44) = "bumperbulb3flash"
  Lampz.MassAssign(60)= bumpersmalllight4
  Lampz.MassAssign(60)= bumperbiglight4
  Lampz.MassAssign(60)= bumpershadow4
  Lampz.MassAssign(60)= bumperhighlight4
  Lampz.Callback(60) = "bumperbulb4flash"

' Center Flash
  Lampz.MassAssign(56) = LightFlash56a
  Lampz.MassAssign(56) = LightFlash56b
  Lampz.MassAssign(56) = LightFlash56c
  Lampz.MassAssign(56) = LightFlash56PFa
  Lampz.MassAssign(56) = LightFlash56PFb
  Lampz.MassAssign(56) = LightFlash56PFc
  Lampz.MassAssign(56) = LightFlash56Topa
  Lampz.MassAssign(56) = LightFlash56Topb
  Lampz.MassAssign(56) = LightFlash56Topc

' Credit Light
  Lampz.MassAssign(59)= l59
  Lampz.MassAssign(59)= l59a
  Lampz.Callback(59) = "DisableLighting2 p59, 0.275, 0.1, "

' Right Apron Light
  Lampz.MassAssign(103)= l103
  Lampz.MassAssign(103)= l103a
  Lampz.Callback(103) = "DisableLighting2 p103, 0.275, 0.1, "

' GI Callbacks
  Lampz.Callback(140) = "GIUpdates"
  Lampz.Callback(141) = "BGGIUpdates"

' Flipper Lights
  Lampz.Callback(119) = "RightFlipperLights"
  Lampz.Callback(119) = "DisableLighting2 prfl, insert_dl_on_redfl, .1,"

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

Class NullFadingObject : Public Property Let IntensityScale(input) : : End Property : End Class

'version 0.11 - Mass Assign, Changed modulate style
'version 0.12 - Update2 (single -1 timer update) update method for core.vbs
'Version 0.12a - Filter can now be accessed via 'FilterOut'
'Version 0.12b - Changed MassAssign from a sub to an indexed property (new syntax: lampfader.MassAssign(15) = Light1 )
'Version 0.13 - No longer requires setlocale. Callback() can be assigned multiple times per index
' Note: if using multiple 'LampFader' objects, set the 'name' variable to avoid conflicts with callbacks
'Version 0.14 - Updated to support modulated signals - Niwak

Class LampFader
  Public FadeSpeedDown(150), FadeSpeedUp(150)
  Private Lock(150), Loaded(150), OnOff(150)
  Public UseFunction
  Private cFilter
  Public UseCallback(150), cCallback(150)
  Public Lvl(150), Obj(150)
  Private Mult(150)
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
      OnOff(x) = 0
      Lock(x) = True : Loaded(x) = False
      Mult(x) = 1
    Next
    Name = "LampFaderNF" 'NEEDS TO BE CHANGED IF THERE'S MULTIPLE OF THESE OBJECTS, OTHERWISE CALLBACKS WILL INTERFERE WITH EACH OTHER!!
    for x = 0 to uBound(OnOff)    'clear out empty obj
      if IsEmpty(obj(x) ) then Set Obj(x) = NullFader' : Loaded(x) = True
    Next
  End Sub

  Public Property Get Locked(idx) : Locked = Lock(idx) : End Property   ''debug.print Lampz.Locked(100) 'debug
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
    'msgbox "Sub " & name & idx & "(aLvl):" & str & "End Sub"
    dim out : out = "Sub " & name & idx & "(aLvl):" & str & "End Sub"
    ExecuteGlobal Out

  End Property

  Public Property Let state(ByVal idx, input) 'Major update path
    if TypeName(input) <> "Double" and typename(input) <> "Integer"  and typename(input) <> "Long" then
      If input Then
        input = 1
      Else
        input = 0
      End If
    End If
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
    ''debug.print debugstr
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
        if OnOff(x) > 0 then 'Fade Up
          Lvl(x) = Lvl(x) + FadeSpeedUp(x)
          if Lvl(x) >= OnOff(x) then Lvl(x) = OnOff(x) : Lock(x) = True
        else 'fade down
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
        if OnOff(x) > 0 then 'Fade Up
          Lvl(x) = Lvl(x) + FadeSpeedUp(x) * FrameTime
          if Lvl(x) >= OnOff(x) then Lvl(x) = OnOff(x) : Lock(x) = True
        else 'fade down
          Lvl(x) = Lvl(x) - FadeSpeedDown(x) * FrameTime
          if Lvl(x) <= 0 then Lvl(x) = 0 : Lock(x) = True
        end if
      end if
    Next
    Update
  End Sub

  Public Sub Update() 'Handle object updates. Update on a -1 Timer! If done fading, loaded(x) = True
    dim x,xx, aLvl : for x = 0 to uBound(OnOff)
      if not Loaded(x) then
        aLvl = Lvl(x)*Mult(x)
        if IsArray(obj(x) ) Then  'if array
          If UseFunction then
            for each xx in obj(x) : xx.IntensityScale = cFilter(aLvl) : Next
          Else
            for each xx in obj(x) : xx.IntensityScale = aLvl : Next
          End If
        else            'if single lamp or flasher
          If UseFunction then
            obj(x).Intensityscale = cFilter(aLvl)
          Else
            obj(x).Intensityscale = aLvl
          End If
        end if
        'if TypeName(lvl(x)) <> "Double" and typename(lvl(x)) <> "Integer" and typename(lvl(x)) <> "Long" then msgbox "uhh " & 2 & " = " & lvl(x)
        'If UseCallBack(x) then execute cCallback(x) & " " & (Lvl(x)) 'Callback
        If UseCallBack(x) then Proc name & x,aLvl 'Proc
        If Lock(x) Then
          if Lvl(x) = OnOff(x) or Lvl(x) = 0 then Loaded(x) = True  'finished fading
        end if
      end if
    Next
  End Sub
End Class


'Lamp Filter
Function LampFilter(aLvl)
  LampFilter = aLvl^1.6 'exponential curve?
End Function


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

'******************************************************
'****  END LAMPZ
'******************************************************


Sub bumperbulb1flash(ByVal aLvl)  'argument is unused
  if Lampz.UseFunction then aLvl = LampFilter(aLvl) 'Callbacks don't get this filter automatically
  FlBumperFadeTarget(1) = aLvl*bumperflashlevel
End Sub

Sub bumperbulb2flash(ByVal aLvl)  'argument is unused
  if Lampz.UseFunction then aLvl = LampFilter(aLvl) 'Callbacks don't get this filter automatically
  FlBumperFadeTarget(2) = aLvl*bumperflashlevel
End Sub

Sub bumperbulb3flash(ByVal aLvl)  'argument is unused
  if Lampz.UseFunction then aLvl = LampFilter(aLvl) 'Callbacks don't get this filter automatically
  FlBumperFadeTarget(3) = aLvl*bumperflashlevel
End Sub

Sub bumperbulb4flash(ByVal aLvl)  'argument is unused
  if Lampz.UseFunction then aLvl = LampFilter(aLvl) 'Callbacks don't get this filter automatically
  FlBumperFadeTarget(4) = aLvl*bumperflashlevel
End Sub



'*******************************************
'  Ball brightness code
'*******************************************

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
End if

' ****************************************************



'*******************************************
' Hybrid code for VR, Cab, and Desktop
'*******************************************

Dim VRThings

if VR_Room = 0 and cab_mode = 0 Then
  for each VRThings in VRStuff:VRThings.visible = 0:Next
  for each VRThings in VRClock:VRThings.visible = 0:Next
  for each VRThings in VRBackglass:VRThings.visible = 0:Next
  for each VRThings in DTRails:VRThings.visible = 1:Next
  OuterPrimBlack_dt.visible = 1
  OuterPrimBlack_cab_VR.visible = 0
  OuterPrimBlack_cab.visible = 0
  L_DT_Embryo.Visible = 1
  L_DT_Embryo.State = 1
  for each VRThings in DTLights:VRThings.visible = 1:Next

Elseif VR_Room = 0 and cab_mode = 1 Then
  for each VRThings in VRStuff:VRThings.visible = 0:Next
  for each VRThings in VRClock:VRThings.visible = 0:Next
  for each VRThings in VRBackglass:VRThings.visible = 0:Next
  for each VRThings in DTBackglass:VRThings.visible = 0: Next
  for each VRThings in DTRails:VRThings.visible = 0:Next
  OuterPrimBlack_dt.visible = 0
  OuterPrimBlack_cab_VR.visible = 0
  OuterPrimBlack_cab.visible = 1
  L_DT_Embryo.Visible = 0
  L_DT_Embryo.State = 0
  for each VRThings in DTLights:VRThings.visible = 0:Next
  for each VRThings in DTLights:VRThings.state = 0:Next

Else
  for each VRThings in VRStuff:VRThings.visible = 1:Next
  for each VRThings in VRClock:VRThings.visible = WallClock:Next
  for each VRThings in DTBackglass:VRThings.visible = 0: Next
  for each VRThings in DTRails:VRThings.visible = 0:Next
  OuterPrimBlack_dt.visible = 0
  OuterPrimBlack_cab_VR.visible = 1
  OuterPrimBlack_cab.visible = 0
  L_DT_Embryo.Visible = 0
  L_DT_Embryo.State = 0
  for each VRThings in DTLights:VRThings.visible = 0:Next
  for each VRThings in DTLights:VRThings.state = 0:Next

'Custom Walls, Floor, and Roof
  if CustomWalls = 1 Then
    VR_Wall_Left.image = "VR_Wall_Left"
    VR_Wall_Right.image = "VR_Wall_Right"
    VR_Floor.image = "VR_Floor"
    VR_Roof.image = "VR_Roof"
  end if


  If topper = 1 Then
    Primary_topper.visible = 1
  Else
    Primary_topper.visible = 0
  End If

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


'*******************************************
' VR Clock
'*******************************************

Dim CurrentMinute ' for VR clock

Sub ClockTimer_Timer()
  Pminutes.RotAndTra2 = (Minute(Now())+(Second(Now())/100))*6
  Phours.RotAndTra2 = Hour(Now())*30+(Minute(Now())/2)
  Pseconds.RotAndTra2 = (Second(Now()))*6
  CurrentMinute=Minute(Now())

End Sub


'*******************************************
' VR Plunger Code
'*******************************************

Sub TimerVRPlunger_Timer
  If PinCab_Shooter.Y < -300 then
       PinCab_Shooter.Y = PinCab_Shooter.Y + 5
  End If
End Sub

Sub TimerVRPlunger1_Timer
  PinCab_Shooter.Y = -390 + (5* Plunger.Position) -20
End Sub


'*******************************************
'Digital Display
'*******************************************

Dim Digits(31)
' 1st Player
Digits(0) = Array(LED10,LED11,LED12,LED13,LED14,LED15,LED16)
Digits(1) = Array(LED20,LED21,LED22,LED23,LED24,LED25,LED26)
Digits(2) = Array(LED30,LED31,LED32,LED33,LED34,LED35,LED36)
Digits(3) = Array(LED40,LED41,LED42,LED43,LED44,LED45,LED46)
Digits(4) = Array(LED50,LED51,LED52,LED53,LED54,LED55,LED56)
Digits(5) = Array(LED60,LED61,LED62,LED63,LED64,LED65,LED66)
Digits(6) = Array(LED004,LED002,LED006,LED007,LED005,LED001,LED003)

' 2nd Player
Digits(7) = Array(LED80,LED81,LED82,LED83,LED84,LED85,LED86)
Digits(8) = Array(LED90,LED91,LED92,LED93,LED94,LED95,LED96)
Digits(9) = Array(LED100,LED101,LED102,LED103,LED104,LED105,LED106)
Digits(10) = Array(LED110,LED111,LED112,LED113,LED114,LED115,LED116)
Digits(11) = Array(LED120,LED121,LED122,LED123,LED124,LED125,LED126)
Digits(12) = Array(LED130,LED131,LED132,LED133,LED134,LED135,LED136)
Digits(13) = Array(LED018,LED016,LED020,LED021,LED019,LED015,LED017)

' 3rd Player
Digits(14) = Array(LED150,LED151,LED152,LED153,LED154,LED155,LED156)
Digits(15) = Array(LED160,LED161,LED162,LED163,LED164,LED165,LED166)
Digits(16) = Array(LED170,LED171,LED172,LED173,LED174,LED175,LED176)
Digits(17) = Array(LED180,LED181,LED182,LED183,LED184,LED185,LED186)
Digits(18) = Array(LED190,LED191,LED192,LED193,LED194,LED195,LED196)
Digits(19) = Array(LED200,LED201,LED202,LED203,LED204,LED205,LED206)
Digits(20) = Array(LED025,LED023,LED027,LED028,LED026,LED022,LED024)

' 4th Player
Digits(21) = Array(LED220,LED221,LED222,LED223,LED224,LED225,LED226)
Digits(22) = Array(LED230,LED231,LED232,LED233,LED234,LED235,LED236)
Digits(23) = Array(LED240,LED241,LED242,LED243,LED244,LED245,LED246)
Digits(24) = Array(LED250,LED251,LED252,LED253,LED254,LED255,LED256)
Digits(25) = Array(LED260,LED261,LED262,LED263,LED264,LED265,LED266)
Digits(26) = Array(LED270,LED271,LED272,LED273,LED274,LED275,LED276)
Digits(27) = Array(LED011,LED009,LED013,LED014,LED012,LED008,LED010)

' Credits
Digits(28) = Array(LED4,LED2,LED6,LED7,LED5,LED1,LED3)
Digits(29) = Array(LED18,LED9,LED27,LED28,LED19,LED8,LED17)

' Balls
Digits(30) = Array(LED39,LED37,LED48,LED49,LED47,LED29,LED38)
Digits(31) = Array(LED67,LED58,LED69,LED77,LED68,LED57,LED59)


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
            obj.intensity=30
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


'*******************************************
' Setup Backglass
'*******************************************

Dim xoff,yoff,zoff,xrot,zscale, xcen,ycen, ix, xx, yy, xobj

Sub setup_backglass()

  xoff = -20
  yoff = 50 '78
  zoff = 699
  xrot = -90
  zscale = 0.0000001

  xcen = 0  '(130 /2) - (92 / 2)
  ycen = (780 /2 ) + (203 /2)

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
end sub


Dim VRDigits(32)
' 1st Player
VRDigits(0) = Array(LED1x0,LED1x1,LED1x2,LED1x3,LED1x4,LED1x5,LED1x6)
VRDigits(1) = Array(LED2x0,LED2x1,LED2x2,LED2x3,LED2x4,LED2x5,LED2x6)
VRDigits(2) = Array(LED3x0,LED3x1,LED3x2,LED3x3,LED3x4,LED3x5,LED3x6)
VRDigits(3) = Array(LED4x0,LED4x1,LED4x2,LED4x3,LED4x4,LED4x5,LED4x6)
VRDigits(4) = Array(LED5x0,LED5x1,LED5x2,LED5x3,LED5x4,LED5x5,LED5x6)
VRDigits(5) = Array(LED6x0,LED6x1,LED6x2,LED6x3,LED6x4,LED6x5,LED6x6)
VRDigits(6) = Array(led6x001,led6x002,led6x003,led6x004,led6x005,led6x006,led6x007)

' 2nd Player
VRDigits(7) = Array(LED8x0,LED8x1,LED8x2,LED8x3,LED8x4,LED8x5,LED8x6)
VRDigits(8) = Array(LED9x0,LED9x1,LED9x2,LED9x3,LED9x4,LED9x5,LED9x6)
VRDigits(9) = Array(LED10x0,LED10x1,LED10x2,LED10x3,LED10x4,LED10x5,LED10x6)
VRDigits(10) = Array(LED11x0,LED11x1,LED11x2,LED11x3,LED11x4,LED11x5,LED11x6)
VRDigits(11) = Array(LED12x0,LED12x1,LED12x2,LED12x3,LED12x4,LED12x5,LED12x6)
VRDigits(12) = Array(LED13x0,LED13x1,LED13x2,LED13x3,LED13x4,LED13x5,LED13x6)
VRDigits(13) = Array(led6x008,led6x009,led6x010,led6x014,led6x011,led6x012,led6x013)

' 3rd Player
VRDigits(14) = Array(LED1x000,LED1x001,LED1x002,LED1x003,LED1x004,LED1x005,LED1x006)
VRDigits(15) = Array(LED1x100,LED1x101,LED1x102,LED1x103,LED1x104,LED1x105,LED1x106)
VRDigits(16) = Array(LED1x200,LED1x201,LED1x202,LED1x203,LED1x204,LED1x205,LED1x206)
VRDigits(17) = Array(LED1x300,LED1x301,LED1x302,LED1x303,LED1x304,LED1x305,LED1x306)
VRDigits(18) = Array(LED1x400,LED1x401,LED1x402,LED1x403,LED1x404,LED1x405,LED1x406)
VRDigits(19) = Array(LED1x500,LED1x501,LED1x502,LED1x503,LED1x504,LED1x505,LED1x506)
VRDigits(20) = Array(led6x022,led6x023,led6x024,led6x028,led6x025,led6x026,led6x027)

' 4th Player
VRDigits(21) = Array(LED2x000,LED2x001,LED2x002,LED2x003,LED2x004,LED2x005,LED2x006)
VRDigits(22) = Array(LED2x100,LED2x101,LED2x102,LED2x103,LED2x104,LED2x105,LED2x106)
VRDigits(23) = Array(LED2x200,LED2x201,LED2x202,LED2x203,LED2x204,LED2x205,LED2x206)
VRDigits(24) = Array(LED2x300,LED2x301,LED2x302,LED2x303,LED2x304,LED2x305,LED2x306)
VRDigits(25) = Array(LED2x400,LED2x401,LED2x402,LED2x403,LED2x404,LED2x405,LED2x406)
VRDigits(26) = Array(LED2x500,LED2x501,LED2x502,LED2x503,LED2x504,LED2x505,LED2x506)
VRDigits(27) = Array(led6x015,led6x016,led6x017,led6x021,led6x018,led6x019,led6x020)

' Credits
VRDigits(28) = Array(LEDax300,LEDax301,LEDax302,LEDax303,LEDax304,LEDax305,LEDax306)
VRDigits(29) = Array(LEDbx400,LEDbx401,LEDbx402,LEDbx403,LEDbx404,LEDbx405,LEDbx406)

' Balls
VRDigits(30) = Array(LEDcx500,LEDcx501,LEDcx502,LEDcx503,LEDcx504,LEDcx505,LEDcx506)
VRDigits(31) = Array(LEDdx600,LEDdx601,LEDdx602,LEDdx603,LEDdx604,LEDdx605,LEDdx606)


dim DisplayColor
DisplayColor =  RGB(255,40,1)

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
          For Each obj In Digits(num)
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
    Object.Opacity = 20
  Else
    Object.Color = RGB(1,1,1)
    Object.Opacity = 8
  End If
End Sub

Sub InitDigits()
  dim tmp, x, obj
  for x = 0 to uBound(VRDigits)
    if IsArray(VRDigits(x) ) then
      For each obj in VRDigits(x)
        obj.height = obj.height + 0
        FadeDisplay obj, 0
      next
    end If
  Next
End Sub

If VR_Room=1 Then
  InitDigits
End If

'*******************************************
' LAMP CALLBACK for the 6 backglass flasher lamps (not the solenoid conrolled ones)
'*******************************************


if VR_Room = 0 and cab_mode = 0 Then
  Set LampCallback = GetRef("UpdateDTLamps")
End If

if VR_Room = 1 Then
  Set LampCallback = GetRef("UpdateVRLamps")
End If

Sub UpdateDTLamps()

  Dim DTL

  If Controller.Lamp(11) = 0 Then: ShootAgainReel.setValue(0):  Else: ShootAgainReel.setValue(1) 'Shoot Again
  If Controller.Lamp(13) = 0 Then: BIPReel.setValue(0):     Else: BIPReel.setValue(1) 'Ball in Play
  If Controller.Lamp(61) = 0 Then: TiltReel.setValue(0):      Else: TiltReel.setValue(1) 'Tilt
  If Controller.Lamp(45) = 0 Then: GameOverReel.setValue(0):    Else: GameOverReel.setValue(1) 'Game Over
  If Controller.Lamp(27) = 0 Then: MatchReel.setValue(0):     Else: MatchReel.setValue(1) 'Match
  If Controller.Lamp(29) = 0 Then: HighScoreReel.setValue(0):   Else: HighScoreReel.setValue(1) 'High Score

' Desktop Backglass EMBRYON Letters

  If DTLampsON = 1 Then

  If Controller.Lamp(14) = 0 Then: L_DT_E.state = 1: Else: L_DT_E.state = 0 'Desktop Backglass E
  If Controller.Lamp(30) = 0 Then: L_DT_M.state = 1: Else: L_DT_M.state = 0 'Desktop Backglass M
  If Controller.Lamp(46) = 0 Then: L_DT_B.state = 1: Else: L_DT_B.state = 0 'Desktop Backglass B
  If Controller.Lamp(62) = 0 Then: L_DT_R.state = 1: Else: L_DT_R.state = 0 'Desktop Backglass R
  If Controller.Lamp(15) = 0 Then: L_DT_Y.state = 1: Else: L_DT_Y.state = 0 'Desktop Backglass Y
  If Controller.Lamp(31) = 0 Then: L_DT_O.state = 1: Else: L_DT_O.state = 0 'Desktop Backglass O
  If Controller.Lamp(47) = 0 Then: L_DT_N.state = 1: Else: L_DT_N.state = 0 'Desktop Backglass N

  Else

  For each DTL in DTLights: DTL.state = 1: Next

  End if

End Sub

Sub UpdateVRLamps()

  If Controller.Lamp(11) = 0 Then: VR_ShootAgain.visible=0: else: VR_ShootAgain.visible=1 'Shoot again
  If Controller.Lamp(13) = 0 Then: VR_BIP.visible=0: else: VR_BIP.visible=1 'Ball in  play
  If Controller.Lamp(61) = 0 Then: VR_Tilt.visible=0: else: VR_Tilt.visible=1 'Tilt
  If Controller.Lamp(45) = 0 Then: VR_GameOver.visible=0: else: VR_GameOver.visible=1 'Game Over
  If Controller.Lamp(27) = 0 Then: VR_Match.visible=0: else: VR_Match.visible=1 'Match
  If Controller.Lamp(29) = 0 Then: VR_HighScore.visible=0: else: VR_HighScore.visible=1 'High Score

  If Controller.Lamp(14) = 0 Then: VR_BG_E.visible = 1: Else: VR_BG_E.visible = 0 'VR Backglass E
  If Controller.Lamp(30) = 0 Then: VR_BG_M.visible = 1: Else: VR_BG_M.visible = 0 'VR Backglass M
  If Controller.Lamp(46) = 0 Then: VR_BG_B.visible = 1: Else: VR_BG_B.visible = 0 'VR Backglass B
  If Controller.Lamp(62) = 0 Then: VR_BG_R.visible = 1: Else: VR_BG_R.visible = 0 'VR Backglass R
  If Controller.Lamp(15) = 0 Then: VR_BG_Y.visible = 1: Else: VR_BG_Y.visible = 0 'VR Backglass Y
  If Controller.Lamp(31) = 0 Then: VR_BG_O.visible = 1: Else: VR_BG_O.visible = 0 'VR Backglass O
  If Controller.Lamp(47) = 0 Then: VR_BG_N.visible = 1: Else: VR_BG_N.visible = 0 'VR Backglass N

End Sub



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


'*******************************************
' Backglass GI and lamps
'*******************************************


Sub TopMidLaneLights(ByVal aLvl)

  dim x

  if Lampz.UseFunction then aLvl = LampFilter(aLvl) 'Callbacks don't get this filter automatically

  If gion = 1 Then
    For each xx in GIMidLane:xx.Intensityscale = aLvl: Next
  End If

End Sub



Sub BGGIUpdates(ByVal aLvl)

  if Lampz.UseFunction then aLvl = LampFilter(aLvl) 'Callbacks don't get this filter automatically

  if vr_room = 1 and cab_mode = 0 then

    if alvl = 0 Then
      VR_Backglass.image = "Backglass_Dark"
      VR_Backglass.blenddisablelighting = 2
    Elseif aLvl = 1 then
      VR_Backglass.image = "Backglass"
      VR_Backglass.blenddisablelighting = 2.5
    Else
      VR_Backglass.image = "Backglass"
      VR_Backglass.blenddisablelighting = .5*alvl + 2
    end if
  End If
End Sub


Sub GIUpdates(ByVal aLvl)

  dim x, girubbercolor, giflippercolor

  if Lampz.UseFunction then aLvl = LampFilter(aLvl) 'Callbacks don't get this filter automatically

  if aLvl = 0 then                    'GI OFF, let's hide ON prims
    if ballbrightness <> -1 then ballbrightness = ballbrightMin

  Elseif aLvl = 1 then                  'GI ON, let's hide OFF prims
    if ballbrightness <> -1 then ballbrightness = ballbrightMax

  Else
    if gilvl = 0 Then               'GI has just changed from OFF to fading, let's show ON
      ballbrightness = ballbrightMin + 1
    elseif gilvl = 1 Then             'GI has just changed from ON to fading, let's show OFF
      ballbrightness = ballbrightMax - 1
    Else
      'no change
    end if
  end if

  dim mm,bp, bpl

  For each xx in GI:xx.Intensityscale = aLvl: Next
  For each xx in GIMidLane:xx.Intensityscale = aLvl: Next

' Flippers (2 main and the one upper right)

  Lflipmesh1.blenddisablelighting = .1 * alvl + .025
  Rflipmesh1.blenddisablelighting = .1 * alvl + .025
  RflipmeshUp.blenddisablelighting = .1 * alvl + .025

' flippers
  giflippercolor = 127*alvl + 128
  MaterialColor "Plastic Flippers",RGB(giflippercolor,giflippercolor,giflippercolor)

' targets
  For each x in GITargets
    x.blenddisablelighting = 0.25 * alvl + .1
  Next

' prims (Blue Pegs)
  For each x in GIPegs
    x.blenddisablelighting = .015 * alvl + .005
  Next

' prims (Blue Top Lane Rollovers)
  For each x in GILaneRollovers
    x.blenddisablelighting = -.025 * alvl + .05
  Next

' prims (Cutouts)
  For each x in GICutoutPrims
    x.blenddisablelighting = .25 * alvl + .15
  Next

  For each x in GICutoutWalls
    x.blenddisablelighting = .1 * alvl + .2
  Next

' prims (Spinner, apron, plastics)

  SpinnerPrim1.blenddisablelighting = 0.1 * alvl + .1
  ApronTopVisible.blenddisablelighting = 0.15 * alvl + .15
  PlungerCover.blenddisablelighting = 0.35 * alvl + .25
  pKickerArm.blenddisablelighting = 0.2 * alvl + .1
  for each x in GIPlastics
    x.blenddisablelighting = -.1 * alvl +.15
  Next

' prims (White Screws Main)
  For each x in ScrewWhiteCaps
    x.blenddisablelighting = .1 * alvl
  Next

' prims (Star Triggers)
  For each x in GIStars
    x.blenddisablelighting = .02 * alvl + .005
  Next

' rubbers
  girubbercolor = 128*alvl + 128
  MaterialColor "Rubber White",RGB(girubbercolor,girubbercolor,girubbercolor)


'GI PF bakes
  FlasherGI.opacity = 300 * alvl

' relaysounds for GI
  if gilvl = 1 then
    Sound_GI_Relay 1, Relay_GI
  end If

  if gilvl = 0 then
    Sound_GI_Relay 0, Relay_GI
  end If

' ball
  if ballbrightness <> ballbrightMax Or ballbrightness <> ballbrightMin Or ballbrightness <> -1 then ballbrightness = INT(alvl * (ballbrightMax - ballbrightMin) + ballbrightMin)

  gilvl = alvl


End Sub


Sub RightFlipperLights(ByVal aLvl)

  if Lampz.UseFunction then aLvl = LampFilter(aLvl) 'Callbacks don't get this filter automatically

  if aLvl = 0 then
    flipperlight5.state = 0
    flipperlight5.intensityscale = 0
    flipperlight6.state = 0
    flipperlight6.intensityscale = 0
    flipperlight7.state = 0
    flipperlight7.intensityscale = 0
    flipperlight8.state = 0
    flipperlight8.intensityscale = 0
  Else
    flipperlight5.state = 1
    flipperlight5.intensityscale = aLvl
    flipperlight6.state = 1
    flipperlight6.intensityscale = aLvl
    flipperlight7.state = 1
    flipperlight7.intensityscale = aLvl
    flipperlight8.state = 1
    flipperlight8.intensityscale = aLvl
  End If

End Sub



'*******************************************
' Set Up Backglass Flashers
'   this is for lining up the backglass flashers on top of a backglass image
'*******************************************

Sub SetBackglass()
  Dim obj

  For Each obj In VRBackglassFlash
    obj.x = obj.x
    obj.height = - obj.y
    obj.y = 60 '78 'adjusts the distance from the backglass towards the user
  Next

End Sub

'*******************************************
' LUT
'*******************************************

Dim bLutActive

Sub SetLUT
  Table1.ColorGradeImage = "LUT" & LUTset
  VRFlashLUT.imageA = "FlashLUT" & LUTset
end sub

Sub LUTBox_Timer
  LUTBox.TimerEnabled = 0
  LUTBox.Visible = 0
  VRFlashLUT.opacity = 0
End Sub

Sub ShowLUT
  LUTBox.visible = 1
  VRFlashLUT.opacity = 100
  Select Case LUTSet
    Case 0: LUTBox.text = "Fleep Natural Dark 1"
    Case 1: LUTBox.text = "Fleep Natural Dark 2"
    Case 2: LUTBox.text = "Fleep Warm Dark"
    Case 3: LUTBox.text = "Fleep Warm Bright"
    Case 4: LUTBox.text = "Fleep Warm Vivid Soft"
    Case 5: LUTBox.text = "Fleep Warm Vivid Hard"
    Case 6: LUTBox.text = "Skitso Natural and Balanced"
    Case 7: LUTBox.text = "Skitso Natural High Contrast"
    Case 8: LUTBox.text = "3rdaxis Referenced THX Standard"
    Case 9: LUTBox.text = "CalleV Punchy Brightness and Contrast"
    Case 10: LUTBox.text = "HauntFreaks Desaturated"
      Case 11: LUTBox.text = "Tomate washed out"
        Case 12: LUTBox.text = "VPW original 1on1"
        Case 13: LUTBox.text = "bassgeige"
        Case 14: LUTBox.text = "blacklight"
        Case 15: LUTBox.text = "B&W Comic Book"
        Case 16: LUTBox.text = "Skitso New Warmer LUT"
        Case 17: LUTBox.text = "Original LUT"
  End Select

  LUTBox.TimerEnabled = 1

End Sub

Sub SaveLUT
  Dim FileObj
  Dim ScoreFile

  Set FileObj=CreateObject("Scripting.FileSystemObject")
  If Not FileObj.FolderExists(UserDirectory) then
    Exit Sub
  End if

  if LUTset = "" then LUTset = 16 'failsafe to original

  Set ScoreFile=FileObj.CreateTextFile(UserDirectory & "Embryon_LUT.txt",True)
  ScoreFile.WriteLine LUTset
  Set ScoreFile=Nothing
  Set FileObj=Nothing
End Sub


Sub LoadLUT
  Dim FileObj, ScoreFile, TextStr
  dim rLine

  Set FileObj=CreateObject("Scripting.FileSystemObject")
  If Not FileObj.FolderExists(UserDirectory) then
    LUTset=16
    Exit Sub
  End if
  If Not FileObj.FileExists(UserDirectory & "Embryon_LUT.txt") then
    LUTset=16
    Exit Sub
  End if
  Set ScoreFile=FileObj.GetFile(UserDirectory & "Embryon_LUT.txt")
  Set TextStr=ScoreFile.OpenAsTextStream(1,0)
    If (TextStr.AtEndOfStream=True) then
      Exit Sub
    End if
    rLine = TextStr.ReadLine
    If rLine = "" then
      LUTset=16
      Exit Sub
    End if
    LUTset = int (rLine)
    Set ScoreFile = Nothing
      Set FileObj = Nothing
End Sub

Sub ShowLUT_Init
  LUTBox.visible = 0
End Sub


if leftoutlanemode = 1 Then
  zCol_Rubber_Post039.collidable = 0
  zCol_Rubber_Post052.collidable = 1
  PegPlasticT14_031.x = 196.5
  PegPlasticT14_031.y = 1485
  Rubber_014.visible = 0
  Rubber_038.visible = 1
  FlasherGI.imageA = "GIBakeImageEasy"
  PF_Shadow.imageA = "AOShadowBakeEasy"
Else
  zCol_Rubber_Post039.collidable = 1
  zCol_Rubber_Post052.collidable = 0
  PegPlasticT14_031.x = 189
  PegPlasticT14_031.y = 1480.5
  Rubber_014.visible = 1
  Rubber_038.visible = 0
  FlasherGI.imageA = "GIBakeImage"
  PF_Shadow.imageA = "AOShadowBake"
End If



'******** Work done by UnclePaulie on Hybrid version 0.01 - 2.0 *********

' v.01  Started from Medusa as a base, and enabled all code for shadows, gBOT, sounds, rolling, hybrid, base table physics, flippers, flipper physics, materials, images, LUT, POV, ball colors, fastflips, lampz, trough logic, knocker
'     Updated all the VR element widths since it's a wide body.
'   Updated the displays to be only 6 characters, and no commas.
'   Updated desktop backglass image to something less busy. Added backglass lamps
'   Added functionality to the desktop backglass letters to be lit via ROM; or an option to always leave lit.
'   Added all posts, pegs, rubber physics, and associated physics and sounds.
'   Added walls and collidables, flippers, rail guides, apron, plunger, balls, rollovers, etc.
'   Added small leaf sensors, slings, gates, spinner, spinner prim, star trigger and animation.
'   Added flipper with light based off of Medusa on the right side.  (originally the flippers were from JP's Medusa table.  Reused here)
'   Added option for flippers to be red, blue, or yellow.
' v.02  Added playfield mesh and saucer.
'   Added bumpers.  Modified Flupper bumper code solution for this table.  Added constants to control DL on bumpers, and light/flash level
'     Added stand up and drop targets code solution developed by Rothbauerw
'   Updated colors of bumpers, pegs, plastic lane overlays to blue.  And updated the lighted flipper image and material.
'   Updated the captive balls for shadows, placement, and surrounding collidables.
'   Updated physics walls in plastics areas.
'   Enabled right flipper functionality
'   Changed the playfield mesh to increase the bevel... matches the real table more closely.
'   Updated the flipper physics, location, start and end angles, and ensure could replicate shots on online videos.
'   Changed the flippers to mid 80's to match shots.
' v.03  Added updated playfield with hole cutouts
'   Moved saucer location slightly to align with playfield hole cutout.  Had to update playfield_mesh
'   Added cutout prims and text overlay
'   Was getting a cor.ballvel error, had to do with the center kicker kicks and/or the relation to the center peg primitive.  REMVOED the primitive and replaced with a physical wall instead.
'     Shifted the left insert wording by captive ball area need to rotate some... 3 of them and all the embryon letters needed to shift / align a little
'   Added plastics and acrylics.  Image upscaled by movieguru.
'   Added GI for bulbs, top plastics, and temporary  playfield (to be removed when baking)
'   Adjusted the balance of intensity in the center and right and left captive ball areas.
'   Added code for adjusting the ball brightness during GI.
'   Added dynamic shadows, and code to stop shadows from going into the plunger lane via notrtlane and rtlanesense trigger.
'   Changed material of the top rollover lanes to be a bit more transparent, and adjusted the debth bias of the bulb prims below, also added to GI routines.
'   Found new spinner image off Flikr on web.  It was beat up, but high res.  Reshaped and touched up a bit.
'   Updated the drop target image
'   3 Bulbs in center area are tied to flash, light 56.
'   Added right apron light
' v.04  Added 3D insert prims, lights, and blooms.
'     The lamps were not defined in manual.  The work to figure them out was done in a prior version.  I used those assignments, and verified.
'     Updated the B insert on the right looks like an "H".  Also tuned the color of EMBRYON letters.
'   Updated the Apron image and apron physics.  Also the plunger cover primitive.
'   Updated GI routine for the apron and plastics.
'   Updated table info / rules
'   Added Drop Target Shadows and added logic in script to control
'   Updated DTHit subroutines to account for a "soft" drop target hit.  Shadow will NOT come on with a light hit... target HAS to drop.
'   Added metal screws and white screw caps.
'   Added logic to manually raise the sw17 drop target on each new ball launch.
'   Updated recommended DIP settings, and added a user setting to turn default DIP switches off, and allow to be changed via F6.
'   Hours spent on studying videos for shots, physics, etc.
'   Removed unused images and materials.
' v.05  Changed to 7 digit rom
'   Added functionality to control the topmiddle lane rollover lights to lamp 87 and GI.
'   Turned reflection on balls off for GI lights under plastics and has wall by it.
' v.06  Changed the other two center pegs to wall physics instead of primitives.  Started getting core.ballvel error again, when adding VR backglass.
'   Added the VR backglass image and associated lamps and text.  Image from Wildman B2S.
'   Added code for GI for backglass
'   Added baked GI lighting
'   Added baked AO shadows
' v.07  Tweaked some GI lighting
'   Looked into ball reflections in playfield holes.  Problem is fixed in VPX 10.8, however this table was designed in 10.74.
'   Minor correction to number of BallShadowA0-7 number of image and array.
'   Found alternate prims for the gates.  More realistic.
'   Darkened the black text on the text overlay image.
'   Adjusted the cab pov settings.  Had to stretch a little, since it's a widebody cab.
'   Redbone enhanced the apron, playfield, and plastics.
' v.08  Minor adjustment to curved metal wall by lower right plastic.
'   Adjusted the default bumperscale brightness down a touch.
'   Adjusted the end angle of lower right flipper slightly, as well as slight physics adjustment.
'   Added drop targets to "target" collection.
'   Slight mod to the upper flipper lower plastic to ensure ball rolls out.
' v.09  Drop target by right captive ball IS tied to SW17.  Changed logic to correctly control that operation of the drop target and stand up.  They are each pulsesw.
'   Added delay in the RaiseSingleDT call for hitting round target.  Also, had to ensure it was a direct hit, so needed it in the ST Hit sub.
'   Adjusted the digital tilt sensitivity up, and the digital nudge strength down.
'   Upper right curved wall too sticky. Adjusted the friction lower on that wall.  Also increased the friction slightly on the wall entering table near gate.
'   Increased elasticity of gate near saucer to help ball bounce around it more.
'   Added ball in plunger lane logic for plunger sound.
'   Added slight delay in other drop target reset areas as well.
'   Redbone updated playfield slightly, and added two more worn flipper images, and new white screw texture.
'   Updated to recent dynamic shadow routine
'   Put physics objects/walls for the guide rails by outlanes.
'   Removed VR Frontwall, as not needed on this table.
'   Added an option for an easier mode for the left outlane.  Also, changed the physics material on the lower post to be a bit more bouncy
'   Changed material on the lower flipper plate rails to match metal walls.
'   Changed the DT targets to not be multi-dimension array assignments
' v.10  Fixed a material issue on the easy and hard peg.
' v.11  Modified the GI bake and AOShadow images for "easy" mode
'   Changed images to webp format for further compression without losing quality.
'   Updated VR Backglass with Hauntfreaks updated images.
'   There was a light reflection near the upper right gate caused by the desktop LEDs "reflection on ball" setting.  Corrected it.
'v2.0.0 Released
'v2.0.1 The clock minute hand was in the wrong location for VR.
'   The VR backbox and backglass was too wide.
'v2.0.2 Forgot to adjust the alpha back to the images for playfield and text overlay.
'   Updated the GI bakes and the AOShadows
'   Changed the ball scratch images.
'   Updated the flipper collide sub and new nFozzy flipper corrections.
'   Updated the flat inserts to include a bulb and method I did on Fathom.
'   Lowered the GI Bulb prims and the respective halos

' Thank you to the VPW team for testing and graphics, especially Redbone, PinstratsDan, Apophis, Tastywasps, Hauntfreaks, Wylte, jsm174, somatik, and bietekwiet!


'   There are two pov's recommended that look good.  One that is stretched to fill the cab, but everything is vertically stretched.
'     Stetched version:       3 24 68 270 1.05 1 1 10 0 -340
'     More 1:1 ratio'ed version:  3 15 270 1 1 1 100 0 -350



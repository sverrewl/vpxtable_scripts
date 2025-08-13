' Based on Bally's Dungeons and Dragons / IPD No. 743 / 1987 / 4 Players
'
' VP8 - Original version by destruk & Lander  - 2009
' VP9 - Original version by mfuegemann - 2014
' VPX - Original version by AgentEighty6 - 2021
' VR version by leojeimroc - 2021
' FSS version by Morttis - 2022
' VR-Hybrid version by AgentEighty6, RobbyKingPin, DaRdog, DGrimmReaper & LoadedWeapon - 2025
'
' This mod contains the following improvements:
'
' AgentEighty6:
' - nFozzy physics,Fleep sound
' - 3d inserts and Flupper Bumpers
' - Animated Star Inserts
' - Bumper lights were not working - fixed
' - "Brite" Flashers overhaul
' - "Rebound" was scoring incorrectly, fixed
' - VR added, used assets from leojreimroc
'
' RobbyKingPin:
' - Updated nFozzy/Rothbauer physics with updated Fleep 2.0 sound from VPW's Bad Cats
' - Rothbauerw Standup and Drop Target codes and VPW Ambient Ball Shadows
' - More light tweaks to more modern standards based on VPW's settings
' - New ball image and scratches based on VPW's tables
' - Updated Flashers and Inserts and added new GI lights
' - Added shadows texture from FSS version made by Morttis
' - Created new primitives for all plastics, added refraction and replaced the VR backglass with a real primitive
' - Edited the primitives for the cabinet, backbox, rails, blades, topper and legs
' - Updated Flupper Domes and Bumper codes
' - Big uprez in the playfield using AI
' - Created physical ball eating targets. Codes used from UnclePaulie's Mousin Around
' - Brought back the original bumpercap, desaturized the image and made it work like a Flupper Bumper
' - VPW Tweak Menu added
' - Added a third room for Mixed Reality
'
' DaRdog:
' - Added VR Mega Room
'
' DGrimmReaper:
' - Re-aligned the displays in VR after the new backglass primitive was added
' - Adjusted backbox to align with the cabinet and VR minimal and mega rooms to align with the legs
'
' LoadedWeapon:
' - Provided a secondary ini file for the POV to fit on Legends Cabinets
'
' Special thanks to everybody involved during development and testing : guus_8005, Tomate, Studlygoorite, cl4ve, hauntfreaks, Cliffy, Enthusiast, remdwaas, MostlyPirate, Sinizin
'
' Shoutout to:
' - destruk, Lander & mfuegemann, AgentEighty6, leojreimroc & Morttis for their work on previous versions
'
' If I didn't mention anybody else in this development cycle I am unaware of this and just know that I still think you are awesome and also want to say you rock!
' And basically everybody from VPW I didn't mention just for having me as part of the family, you all absolutely rock!

Option Explicit : Randomize
SetLocale 1033

'*******************************************
'  ZCON: Constants and Global Variables
'*******************************************

Const BallSize = 50        'Ball diameter in VPX units; must be 50
Const BallMass = 1          'Ball mass must be 1
Const tnob = 3            'Total number of balls the table can hold
Const lob = 0            'Locked balls
Const cGameName = "dungdrag"     'The unique alphanumeric name for this table

Dim tablewidth
tablewidth = Table1.width
Dim tableheight
tableheight = Table1.height
Dim BIP              'Balls in play
BIP = 0
Dim BIPL              'Ball in plunger lane
BIPL = False
Dim plungerpress

Dim SubstMagnaSaveButtons

' FLIPPERS ALSO ACTIVATE MAGNASAVE (MAGIC GATE)
' Set to True if Your cabinet has no Magnasave Buttons. (default=False)
SubstMagnaSaveButtons = False

Dim bsTrough,bsLeftTeleporter,bsRightTeleporter

' ****************************************************
' standard definitions
' ****************************************************

Const UseSolenoids  = 2
Const UseLamps    = 1
Const UseSync     = 0
Const HandleMech  = 0
Const UseGI     = 0

'Standard Sounds
Const SSolenoidOn=""
Const SSolenoidOff=""
Const SFlipperOn=""
Const SFlipperOff=""
Const sCoin=""

On Error Resume Next
ExecuteGlobal GetTextFile("controller.vbs")
If Err Then MsgBox "You need the controller.vbs in order to run this table, available in the vp10 package."
On Error Goto 0

Dim DesktopMode: DesktopMode = Table1.ShowDT

'*******************************************
'  ZOPT: User Options
'*******************************************

'----- Desktop Visible Elements -----
Dim VarHidden, x, UseVPMDMD

If Table1.ShowDT = true and RenderingMode <> 2 Then
  For each VR_Obj in aReels : VR_Obj.Visible = 1 : Next
  VarHidden = 1
  PinCab_Rails.Visible = 1
  PinCab_Blades.Visible = 1
Else
  For each VR_Obj in aReels
        VR_Obj.Visible = 0
    Next
    VarHidden = 0
  PinCab_Rails.Visible = 0
  If Table1.ShowFSS = False And LegendsCabMod = 1 Then
    Flasherbloom1.Height = 500
    Flasherbloom2.Height = 500
    Flasherbloom3.Height = 500
  End If
End if

'----- VR Room Auto-Detect -----
Dim VRRoom, VR_Obj, VRMode

'----- VR Options -----
Dim VRRoomChoice : VRRoomChoice = 1 ' 1 = Mega Room 2 = Minimal Room, 3 = Room360
Const VRTest = 0

If RenderingMode = 2 or Table1.ShowFSS = True or VRTest = 1 Then
  VRMode = True
    PinCab_Rails.Visible = 1
  For Each VR_Obj in aReels : VR_Obj.Visible = 0 : Next
  For Each VR_Obj in VRCabinet : VR_Obj.Visible = 1 : Next
  For Each VR_Obj in VRLedDigits:VR_Obj.Visible = 1 : Next
  If VRRoomChoice = 1 Then
    For Each VR_Obj in VRMinimalRoom : VR_Obj.Visible = 0 : Next
    For Each VR_Obj in VRMegaRoom : VR_Obj.Visible = 1 : Next
    Room360.Visible = 0
  End If
  If VRRoomChoice = 2 Then
    For Each VR_Obj in VRMinimalRoom : VR_Obj.Visible = 1 : Next
    For Each VR_Obj in VRMegaRoom : VR_Obj.Visible = 0 : Next
    Room360.Visible = 0
  End If
  If VRRoomChoice = 3 Then
    For Each VR_Obj in VRMinimalRoom : VR_Obj.Visible = 0 : Next
    For Each VR_Obj in VRMegaRoom : VR_Obj.Visible = 0 : Next
    Room360.Visible = 1
  End If
Else
  VRMode = False
  For Each VR_Obj in VRCabinet : VR_Obj.Visible = 0 : Next
  For Each VR_Obj in VRLedDigits:VR_Obj.Visible = 0 : Next
  For Each VR_Obj in VRMinimalRoom : VR_Obj.Visible = 0 : Next
  For Each VR_Obj in VRMegaRoom : VR_Obj.Visible = 0 : Next
  Room360.Visible = 0
End If

If VRRoomChoice <> 0 Then UseVPMDMD = True Else UseVPMDMD = DesktopMode

if B2SOn = true then VarHidden = 1

'//  PositionalSoundPlaybackConfiguration:
'//  Specifies the sound playback configuration. Options:
'//  1 = Mono
'//  2 = Stereo (Only L+R channels)
'//  3 = Surround (Full surround sound support for SSF - Front L + Front R + Rear L + Rear R channels)
Const PositionalSoundPlaybackConfiguration = 3

'//  RelaysPosition:
'//  1 = Relays positioned with power board (Provides sound spread from the left and right back channels)
'//  2 = Relays positione

'*******************************************
' ZTIM: Timers
'*******************************************

'The FrameTimer interval should be -1, so executes at the display frame rate
'The frame timer should be used to update anything visual, like some animations, shadows, etc.
'However, a lot of animations will be handled in their respective _animate subroutines.

Dim BallShadow : BallShadow = Array (BallShadow1, BallShadow2, BallShadow3)

Dim FrameTime, InitFrameTime
InitFrameTime = 0

FrameTimer.Interval = -1
Sub FrameTimer_Timer() 'The frame timer interval should be -1, so executes at the display frame rate
  FrameTime = GameTime - InitFrameTime
  InitFrameTime = GameTime  'Count frametime
  'Add animation stuff here
  RollingUpdate
  DoDTAnim
  DoSTAnim
  UpdatePlunger
  If AmbientBallShadowOn = 1 Then
    BSUpdate
  End If
  If VRMode = True Then
    VRDisplayTimer
  Else
    UpdateLeds
  End If
  If controller.Lamp(13) Then ObjLevel(2) = 1 : FlasherFlash2_Timer
  If controller.Lamp(56) Then ObjLevel(3) = 1 : FlasherFlash3_Timer
  If controller.Lamp(87) Then ObjLevel(1) = 1 : FlasherFlash1_Timer
  Topper_FlasherLeft.visible = controller.Lamp(43)
  Topper_FlasherRight.visible = controller.Lamp(60)
  Topper_FlasherDragon.visible = controller.Lamp(44)
End Sub

'The CorTimer interval should be 10. It's sole purpose is to update the Cor calculations
CorTimer.Interval = 10
Sub CorTimer_Timer(): Cor.Update: End Sub

Sub UpdatePlunger()
  If plungerpress = 1 then
    If PinCab_Shooter.Y < -10 then
      PinCab_Shooter.Y = PinCab_Shooter.Y + 2.5
    End If
  Else
    PinCab_Shooter.Y = -119.3572 + (2.5* Plunger.Position) -20
  End If
End Sub

Sub LeftFlipper_Animate
  dim a: a = LeftFlipper.CurrentAngle
  LeftFlipperShadow.RotY = a - 90
  LFLogo.RotY = a + 240
End Sub

Sub RightFlipper_Animate
  dim a: a = RightFlipper.CurrentAngle
  RightFlipperShadow.RotY = a - 90
  RFLogo.RotY = a + 120
End Sub

Sub URightFlipper_Animate
  dim a: a = URightFlipper.CurrentAngle
  URightFlipperShadow.RotY = a - 90
  RF2Logo.RotY = a + 150
End Sub

LoadVPM "03020000","6803.VBS",3.2

Dim xx
Dim DDBall1, DDBall2, DDBall3, gBOT

Sub Table1_Init
  vpmInit Me
  On Error Resume Next
    With Controller
    .GameName=cGameName
    If Err Then MsgBox "Can't start Game" & cGameName & vbNewLine & Err.Description:Exit Sub
    .SplashInfoLine="Dungeons & Dragons, Bally 1987"
    .HandleMechanics=0
    .ShowDMDOnly=1
    .ShowFrame=0
    .ShowTitle=0
    .Hidden= 1
    On Error Resume Next
    .Run
    If Err Then MsgBox Err.Description
    On Error Goto 0
    End With
  On Error Goto 0

  vpmNudge.TiltSwitch=15
  vpmNudge.Sensitivity=3
  vpmNudge.TiltObj=Array(Bumper1,Bumper2,Bumper3,Bumper4,LeftSlingshot,RightSlingshot)

  vpmMapLights AllLights  'Lamp names are all messed up but the timer values are correct to be able to use this method

  '************  Trough **************
  Set DDBall3 = sw46.CreateSizedballWithMass(Ballsize/2,Ballmass)
  Set DDBall2 = sw47.CreateSizedballWithMass(Ballsize/2,Ballmass)
  Set DDBall1 = sw48.CreateSizedballWithMass(Ballsize/2,Ballmass)
  gBOT = Array(DDBall1,DDBall2,DDBall3)

  Controller.Switch(46) = 1
  Controller.Switch(47) = 1
  Controller.Switch(48) = 1

  pMillionRamp.rotx = -24
  RampIsUp = True         'Ramp starts up
  RampTimer.enabled = True
  MillionRamp.collidable = False

  ' initialize the traps to transz = 0... solves a corner condition such that they can go below table if there's a VPX crash.
  RightTeleporterBlock.transz = 0
  LeftTeleporterBlock.transz = 0
  If VRMode = True Then
    InitDigits
    setup_backglass
  End If
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

Function RndInt(min, max)
    RndInt = Int(Rnd() * (max-min + 1) + min)' Sets a random number integer between min and max
End Function

Function RndNum(min, max)
    RndNum = Rnd() * (max-min) + min' Sets a random number between min and max
End Function

' ****************************************************
' keys
' ****************************************************
Sub Table1_KeyDown(ByVal keycode)
  'If keycode = LeftFlipperKey Then FlipperActivate LeftFlipper, LFPress
  'If keycode = RightFlipperKey Then FlipperActivate RightFlipper, RFPress
  If keycode = LeftTiltKey Then Nudge 90, 5:SoundNudgeLeft()
  If keycode = RightTiltKey Then Nudge 270, 5:SoundNudgeRight()
  If keycode = CenterTiltKey Then Nudge 0, 3:SoundNudgeCenter()
  If keycode = keyInsertCoin1 or keycode = keyInsertCoin2 or keycode = keyInsertCoin3 or keycode = keyInsertCoin4 Then
    Select Case Int(rnd*3)
      Case 0: PlaySound ("Coin_In_1"), 0, CoinSoundLevel, 0, 0.25
      Case 1: PlaySound ("Coin_In_2"), 0, CoinSoundLevel, 0, 0.25
      Case 2: PlaySound ("Coin_In_3"), 0, CoinSoundLevel, 0, 0.25
    End Select
  End If
  If keycode = PlungerKey Then
    Plunger.Pullback
    SoundPlungerPull
    plungerpress = 1
  End If
  if SubstMagnaSaveButtons = False then
    If keycode = LeftMagnaSave Then Controller.Switch(5) = True  ' LeftMagnaSave
    If keycode = RightMagnaSave Then Controller.Switch(7) = True  ' RightMagnaSave
  else
    If keycode = LeftFlipperKey Then Controller.Switch(5) = True  ' LeftMagnaSave
    If keycode = RightFlipperKey Then Controller.Switch(7) = True  ' RightMagnaSave
  end if
  If keycode = StartGameKey then soundStartButton()

  If keycode = LeftFlipperKey Then
    Primary_flipper_button_left.X = 2135.411 + 8
  End If
  If keycode = RightFlipperKey Then
    Primary_flipper_button_right.X = 2100 - 8
  End If
  If keycode = LeftMagnaSave Then
    Magna_save_left.X = 2135.411 + 8
  End If
  If keycode = RightMagnaSave Then
    Magna_save_right.X = 2100 - 8
  End If
  If keycode = StartGameKey Then
        Primary_startbutton.y = 56 -5
  End If
  If vpmKeyDown(keycode) Then Exit Sub
End Sub

Sub Table1_KeyUp(ByVal keycode)
  If keycode = LeftFlipperKey Then
    Primary_flipper_button_left.X = 2135.411
  End If
  If keycode = RightFlipperKey Then
    Primary_flipper_button_right.X = 2100
  End If
  If keycode = LeftMagnaSave Then
    Magna_save_left.X = 2135.411
  End If
  If keycode = RightMagnaSave Then
    Magna_save_right.X = 2100
  End If
    If keycode = StartGameKey Then
        Primary_startbutton.y = 56
        End If
  If KeyCode = PlungerKey Then
    Plunger.Fire
    plungerpress = 0
    PinCab_Shooter.Y = -50
    If BIPL = 1 Then
      SoundPlungerReleaseBall()   'Plunger release sound when there is a ball in shooter lane
    Else
      SoundPlungerReleaseNoBall() 'Plunger release sound when there is no ball in shooter lane
    End If
  End If
  if SubstMagnaSaveButtons = False then
    If keycode = LeftMagnaSave Then Controller.Switch(5) = False  ' LeftMagnaSave
    If keycode = RightMagnaSave Then Controller.Switch(7) = False  ' RightMagnaSave
  else
    If keycode = LeftFlipperKey Then Controller.Switch(5) = False  ' LeftMagnaSave
    If keycode = RightFlipperKey Then Controller.Switch(7) = False  ' RightMagnaSave
  end if
  If vpmKeyUp(keycode) Then Exit Sub
End Sub

'-----------------------------------
'------  Solenoid Assignment  ------
'-----------------------------------
'SolCallback(1) = "vpmSolSound SoundFX(""fx_bumper1"",DOFContactors),"      'Top Bumper
'SolCallback(2) = "vpmSolSound SoundFX(""fx_bumper2"",DOFContactors),"      'Left Bumper
'SolCallback(3) = "vpmSolSound SoundFX(""fx_bumper3"",DOFContactors),"      'Right Bumper
'SolCallback(4) = "vpmSolSound SoundFX(""fx_bumper1"",DOFContactors),"      'Bottom Bumper
'SolCallback(5) = '"RandomSoundSlingshotLeft sling1" '"vpmSolSound SoundFX(""LeftSlingShot"",DOFContactors),"     'Left Slingshot
'SolCallback(6) = '"RandomSoundSlingshotRight sling2" '"vpmSolSound SoundFX(""RightSlingShot"",DOFContactors),"     'Right Slingshot
SolCallback(7) = "SolDTBank"            '7 Reset Drop Targets
SolCallback(8) = "SolLeftKicker"          'Kicker Left
SolCallback(9) = "SolRightKicker"         'Kicker Right

SolCallback(10) = "SolLeftTeleporter"       'Teleporter Left
SolCallback(11) = "SolRightTeleporter"        'Teleporter Right

SolCallback(12)  = "SolRelease"       'Kick to Playfield
'13 reserved for German use
SolCallback(14) = "SolDrain"          'OutHole
SolCallback(15)  = "KnockerSolenoid"        'Knocker

SolCallback(17) = "SolGi"             'GI Light control

SolCallback(18) = "SolFlexsaveRight"        'Flexsave Right (Right Gate)
SolCallback(20) = "SolFlexsaveLeft"         'Flexsave Left (Left Gate)

SolCallback(19) = "vpmNudge.SolGameOn"

SolCallback(sLRFlipper)="SolRFlipper"
SolCallback(sLLFlipper)="SolLFlipper"

'*******************************************
' ZFLP: Flippers
'*******************************************

Const ReflipAngle = 20
Const QuickFlipAngle = 20

Sub SolLFlipper(Enabled)
  If Enabled Then
    FlipperActivate LeftFlipper, LFPress
    LF.Fire  'leftflipper.rotatetoend

    If LeftFlipper.currentangle < LeftFlipper.endangle + ReflipAngle Then
      'Play partial flip sound and stop any flip down sound
      StopAnyFlipperLowerLeftDown()
      RandomSoundFlipperLowerLeftReflip LeftFlipper
    Else
      'Play full flip sound
      If BallNearLF = 0 Then
        RandomSoundFlipperLowerLeftUpFullStroke LeftFlipper
      End If
      If BallNearLF = 1 Then
        Select Case Int(Rnd*2)+1
          Case 1 : RandomSoundFlipperLowerLeftUpDampenedStroke LeftFlipper
          Case 2 : RandomSoundFlipperLowerLeftUpFullStroke LeftFlipper
        End Select
      End If
    End If
  Else
    FlipperDeActivate LeftFlipper, LFPress
    LeftFlipper.RotateToStart
    If LeftFlipper.currentangle < LeftFlipper.startAngle - 5 Then
      'Play flip down sound
      RandomSoundFlipperLowerLeftDown LeftFlipper
    End If
    If LeftFlipper.currentangle < LeftFlipper.startAngle + QuickFlipAngle and LeftFlipper.currentangle <> LeftFlipper.endangle Then
      'Play quick flip sound and stop any flip up sound
      StopAnyFlipperLowerLeftUp()
      RandomSoundLowerLeftQuickFlipUp()
    Else
      FlipperLeftLowerHitParm = FlipperUpSoundLevel
    End If
  End If
End Sub

Sub SolRFlipper(Enabled)
  If Enabled Then
    FlipperActivate RightFlipper, RFPress
    RF.Fire 'rightflipper.rotatetoend
    URightFlipper.RotateToEnd
    If RightFlipper.currentangle > RightFlipper.endangle - ReflipAngle Then
      'Play partial flip sound and stop any flip down sound
      StopAnyFlipperLowerRightDown()
      RandomSoundFlipperLowerRightReflip RightFlipper
    Else
      'Play full flip sound
      If BallNearRF = 0 Then
        RandomSoundFlipperLowerRightUpFullStroke RightFlipper
      End If

      If BallNearRF = 1 Then
        Select Case Int(Rnd*2)+1
          Case 1 : RandomSoundFlipperLowerRightUpDampenedStroke RightFlipper
          Case 2 : RandomSoundFlipperLowerRightUpFullStroke RightFlipper
        End Select
      End If
    End If
  Else
    FlipperDeActivate RightFlipper, RFPress
    RightFlipper.RotateToStart
    URightFlipper.RotateToStart
    If RightFlipper.currentangle > RightFlipper.startAngle + 5 Then
      'Play flip down sound
      RandomSoundFlipperLowerRightDown RightFlipper
    End If
    If RightFlipper.currentangle < RightFlipper.startAngle + QuickFlipAngle and RightFlipper.currentangle <> RightFlipper.endangle Then
      'Play quick flip sound and stop any flip up sound
      StopAnyFlipperLowerRightUp()
      RandomSoundLowerRightQuickFlipUp()
    Else
      FlipperRightLowerHitParm = FlipperUpSoundLevel
    End If
  End If
End Sub

'******************************************************
' TROUGH
'******************************************************

Sub sw46_Hit():Controller.Switch(46) = 1:UpdateTrough:End Sub
Sub sw46_UnHit():Controller.Switch(46) = 0:UpdateTrough:End Sub
Sub sw47_Hit():Controller.Switch(47) = 1:UpdateTrough:End Sub
Sub sw47_UnHit():Controller.Switch(47) = 0:UpdateTrough:End Sub
Sub sw48_Hit():Controller.Switch(48) = 1:UpdateTrough:End Sub
Sub sw48_UnHit():Controller.Switch(48) = 0:UpdateTrough:End Sub

Sub UpdateTrough()
  UpdateTroughTimer.Interval = 300
  UpdateTroughTimer.Enabled = 1
End Sub

Sub UpdateTroughTimer_Timer()
  If sw48.BallCntOver = 0 Then sw47.kick 64, 20
  If sw47.BallCntOver = 0 Then sw46.kick 64, 20
  Me.Enabled = 0
End Sub

'******************************************************
' DRAIN & RELEASE
'******************************************************

Sub sw8_Hit()
  RandomSoundOutholeHit sw8
  Controller.Switch(8) = 1
End Sub

Sub SolDrain(enabled)
  If enabled Then
    sw8.kick 64, 20
    Controller.Switch(8) = 0
  End If
End Sub

Sub SolRelease(enabled)
  If enabled Then
    sw48.kick 64, 9
    RandomSoundShooterFeeder
  End If
End Sub

Sub SolLeftKicker(enabled)
  if enabled then
    LeftKicker.kick 0,100
    RandomSoundEjectHoleSolenoidLeft()
  end if
End Sub

Sub SolRightKicker(enabled)
  if enabled then
    RightKicker.kick 0,100
    RandomSoundEjectHoleSolenoidRight()
  end if
End Sub

' *********************************************************************
' ball traps up and down
' *********************************************************************
Dim ballInLeftTrap  : Set ballInLeftTrap  = Nothing
Dim ballInRightTrap : Set ballInRightTrap   = Nothing

Sub SolLeftTeleporter(Enabled)
  If enabled then
    LeftGate.Collidable = True
    TrapDown 32, LeftTeleporter, LeftTeleporterBlock, ballInLeftTrap
  Else
    TrapUp 32, 34, LeftTeleporter, LeftTeleporterBlock, ballInLeftTrap
  End If
End Sub

Sub SolRightTeleporter(Enabled)
  If enabled then
    MillionRamp.collidable = False
    RampIsUp = True
    RampTimer.Enabled = True
    TrapDown 40, RightTeleporter, RightTeleporterBlock, ballInRightTrap
  Else
    TrapUp 40, 39, RightTeleporter, RightTeleporterBlock, ballInRightTrap
  End If
End Sub

' trap is hit by a ball
Sub LeftTeleporter_Hit()
  BallInTrap 32, 31, LeftTeleporter, LeftTeleporterBlock, ballInLeftTrap, ActiveBall
End Sub

Sub LeftTeleporter_Timer()
  MoveTrap 31, LeftTeleporter, LeftTeleporterBlock, ballInLeftTrap, 0
End Sub

Sub RightTeleporter_Hit()
  BallInTrap 40, 39, RightTeleporter, RightTeleporterBlock, ballInRightTrap, ActiveBall
  MillionRamp.collidable = True
  RampIsUp = False
  RampTimer.enabled = True
End Sub

Sub RightTeleporter_Timer()
  MoveTrap 39, RightTeleporter, RightTeleporterBlock, ballInRightTrap, 0
End Sub

Sub TrapDown(id, trap, prim, ball)
  trap.Enabled = False
  Controller.Switch(id) = True
  ' move trap prim and ball down
  trap.TimerEnabled = False
  MoveTrap 0, trap, prim, ball, 15
  PlaySoundAtLevelStatic ("fx_solenoid"), 1, trap
End Sub
Sub TrapUp(id, idBallInTrap, trap, prim, ball)
  trap.Enabled = True
  Controller.Switch(id) = False
  ' move trap prim and ball up
  trap.TimerEnabled = False
  MoveTrap 0, trap, prim, ball, 11
  PlaySoundAtLevelStatic ("fx_solenoid"), 1, trap
End Sub

Sub BallInTrap(id, idBallInTrap, trap, prim, ball, ballA)
  Set ball = ballA : ball.VelX = 0 : ball.VelY = 0 : ball.VelZ = 0
  Controller.Switch(idBallInTrap) = True
  TrapDown id, trap, prim, ball
End Sub

Sub MoveTrap(idBallInTrap, trap, prim, ball, interval)
  If Not trap.TimerEnabled Then
    trap.TimerInterval = interval
    trap.TimerEnabled  = True
End If
    If trap.TimerInterval = 15 Then
        ' trap goes down
        prim.TransZ = prim.TransZ - 6
        If Not ball Is Nothing Then ball.Z = prim.transz + 25
        If prim.TransZ <= -60 Then
            trap.TimerEnabled = False
        End If
    ElseIf trap.TimerInterval = 11 Then
        ' trap goes up
        prim.TransZ = prim.TransZ + 12
        If Not ball Is Nothing Then ball.Z = prim.transz + 25
        If prim.TransZ >= 0 Then
            trap.TimerInterval = 100
        End If
    Else
    trap.TimerEnabled = False
    ' kick the trapped ball
    If Not ball Is Nothing Then
      trap.Kick 185, 0.5
      Controller.Switch(idBallInTrap) = False
      Set ball = Nothing
    End If
  End If
End Sub

Dim RampIsUp
Sub RampTimer_Timer
  if RampIsUp then      'up to 0
    pMillionRamp.rotx = pMillionRamp.rotx + 1
    if pMillionRamp.rotx >= 0 then
      RampTimer.enabled = False
      pMillionRamp.rotx = 0
    end if
  else            'down to -23
    pMillionRamp.rotx = pMillionRamp.rotx - 1
    if pMillionRamp.rotx <= -23 then
      RampTimer.enabled = False
      pMillionRamp.rotx = -23
    end if
  end if
End Sub

Sub SolFlexsaveRight(enabled)
  if enabled then
    RightMagic.rotatetoend
  else
    RightMagic.rotatetostart
  end if
End Sub

Sub SolFlexsaveLeft(enabled)
  if enabled then
    LeftMagic.rotatetoend
  else
    LeftMagic.rotatetostart
  end if
End Sub

'---------------------------------
'------  Switch Assignment  ------
'---------------------------------
sub sw1_hit : STHit 1 : End Sub     'Dust Target 1
sub sw2_hit : STHit 2 : End Sub       'Dust Target 2
sub sw3_hit : STHit 3 : End Sub       'Dust Target 3
sub sw4_hit : STHit 4 : End Sub     'Dust Target 4
'5  Cabinet Left - handled elsewhere
'6  Cabinet Credit
'7  Cabinet Right - handled elsewhere
'8  Outhole - handled elsewhere
'9  Coins Right (Door) - handled elsewhere
'10 Coins Left (Door) - handled elsewhere
'11 Coins Middle (Door) - handled elsewhere

Sub sw12_Hit:vpmTimer.PulseSw 12:me.timerenabled = 0:AnimateStar star12, sw12, 1:End Sub
Sub sw12_UnHit:me.timerinterval = 7:me.timerenabled = 1:End Sub
Sub sw12_timer:AnimateStar star12, sw12, 0:End Sub
Sub sw12a_Hit:vpmTimer.PulseSw 12:me.timerenabled = 0:AnimateStar star12a, sw12a, 1:End Sub
Sub sw12a_UnHit:me.timerinterval = 7:me.timerenabled = 1:End Sub
Sub sw12a_timer:AnimateStar star12a, sw12a, 0:End Sub

Sub sw13_Hit:vpmTimer.PulseSw 13:me.timerenabled = 0:AnimateStar star13, sw13, 1:End Sub
Sub sw13_UnHit:me.timerinterval = 7:me.timerenabled = 1:End Sub
Sub sw13_timer:AnimateStar star13, sw13, 0:End Sub
Sub sw13a_Hit:vpmTimer.PulseSw 13:me.timerenabled = 0:AnimateStar star13a, sw13a, 1:End Sub
Sub sw13a_UnHit:me.timerinterval = 7:me.timerenabled = 1:End Sub
Sub sw13a_timer:AnimateStar star13a, sw13a, 0:End Sub

'14 Slam Tilt
'15 Tilt - handled elsewhere

Sub Rebound_Hit:vpmtimer.pulsesw 16:End Sub     'Rebound 16

Sub Bumper1_Hit 'Top Bumper
  vpmtimer.pulsesw 17
  RandomSoundBumperUp Bumper1
End Sub

Sub Bumper2_Hit 'Left Bumper
  vpmtimer.pulsesw 18
  RandomSoundBumperLeft Bumper2
End Sub

Sub Bumper3_Hit 'Right Bumper
  vpmtimer.pulsesw 19
  RandomSoundBumperLeft Bumper3
End Sub

Sub Bumper4_Hit 'Bottom Bumper
  vpmtimer.pulsesw 20
  RandomSoundBumperLow Bumper4
End Sub

'21 Left Slingshot - handled elsewhere
'22 Right Slingshot - handled elsewhere

Sub Trigger23_Hit:controller.switch(23) = True:End Sub    'Dragon Lair Left
Sub Trigger23_Unhit:controller.switch(23) = False:End Sub

Sub Trigger23a_hit()
  WireRampOn True  'Play Plastic Ramp Sound
  If (ActiveBall.VelY > 0) Then
    RandomSoundRampFlapDown()
  ElseIf (ActiveBall.VelY < 0) Then
    RandomSoundRampFlapUp()
  End If
End Sub

Sub Trigger23b_hit()
  WireRampOff  'Turn off the Plastic Ramp Sound
End Sub

Sub Trigger23b_unhit()
  WireRampOn False  'On Wire Ramp, Play Wire Ramp Sound
End Sub

Sub Trigger23z_hit()
  WireRampOff  'Exiting Wire Ramp Stop Playing Sound
End Sub

Sub Trigger24_Hit:controller.switch(24) = True:End Sub    'Dragon Lair Right
Sub Trigger24_Unhit:controller.switch(24) = False:End Sub

Sub Trigger24a_hit()
  WireRampOn True  'Play Plastic Ramp Sound
  If (ActiveBall.VelY > 0) Then
    RandomSoundRampFlapDown()
  ElseIf (ActiveBall.VelY < 0) Then
    RandomSoundRampFlapUp()
  End If
End Sub

Sub Trigger24b_hit()
  WireRampOff  'Turn off the Plastic Ramp Sound
End Sub

Sub Trigger24b_unhit()
  WireRampOn False  'On Wire Ramp, Play Wire Ramp Sound
End Sub

Sub Trigger24z_hit()
  WireRampOff  'Exiting Wire Ramp Stop Playing Sound
End Sub

Sub sw25_hit : STHit 25 : End Sub     'Shield Target 1
Sub sw26_hit : STHit 26 : End Sub     'Shield Target 2
Sub sw27_hit : STHit 27 : End Sub     'Shield Target 3
Sub sw28_hit : STHit 28 : End Sub     'Sword Target 1
Sub sw29_hit : STHit 29 : End Sub     'Sword Target 2
Sub sw30_hit : STHit 30 : End Sub     'Sword Target 3
                          '31 Teleporter Left Empty
' Drop targets
Sub sw33_Hit
  DTHit 33
  TargetBouncer Activeball, 1.5
End Sub

Sub sw34_Hit
  DTHit 34
  TargetBouncer Activeball, 1.5
End Sub

Sub sw35_Hit
  DTHit 35
  TargetBouncer Activeball, 1.5
End Sub

Sub SolDTBank(enabled)
  Dim xx
  If enabled Then
    RandomSoundDropTargetReset sw34p
    DTRaise 33
    DTRaise 34
    DTRaise 35
  End If
End Sub

Sub Skill36_Hit:controller.switch(36) = True:End Sub      'Skill 1 Bottom
Sub Skill36_Unhit:controller.switch(36) = False:End Sub
Sub Skill37_Hit:controller.switch(37) = True:End Sub      'Skill 2 Middle
Sub Skill37_Unhit:controller.switch(37) = False:End Sub
Sub Skill38_Hit:controller.switch(38) = True:End Sub      'Skill 3 Top
Sub Skill38_Unhit:controller.switch(38) = False:End Sub
                          '39 Teleporter Right Empty
Sub Trigger41_Hit:controller.switch(41) = True:End Sub      'Left Return Lane
Sub Trigger41_Unhit:controller.switch(41) = False:End Sub
Sub Trigger42_Hit:controller.switch(42) = True:End Sub      'Level Switch
Sub Trigger42_Unhit:controller.switch(42) = False:End Sub
Sub Trigger43_Hit:controller.switch(43) = True:End Sub      'Million Switch
Sub Trigger43_Unhit:controller.switch(43) = False:End Sub
Sub Trigger44_Hit:controller.switch(44) = True:End Sub      'Restore Weapons
Sub Trigger44_Unhit:controller.switch(44) = False:End Sub
'45 not used
'46 Outhole 1 Left - handled elsewhere
'47 Outhole 2 Middle - handled elsewhere
'48 Outhole 3 Right - handled elsewhere

Sub Trigger1_Hit
  BIPL = True
End Sub

Sub Trigger1_UnHit
  BIPL = False
End Sub

'******************************
' Setup Backglass
'******************************

Dim xoff,yoff1, yoff2, yoff3, yoff4, yoff5,zoff,xrot,zscale, xcen,ycen

Sub setup_backglass()
  xoff = -20
  yoff1 = 93 ' this is where you adjust the forward/backward position for player scores
  zoff = 691
  xrot = -90
  center_digits()
end sub

Sub center_digits()
  Dim ix, xx, yy, yfact, xfact, xobj

  zscale = 0.0000001

  xcen = (130 /2) - (92 / 2)
  ycen = (780 /2 ) + (203 /2)

  for ix = 0 to 27
    For Each xobj In VRDigits(ix)

      xx = xobj.x

      xobj.x = (xoff - xcen) + xx
      yy = xobj.y ' get the yoffset before it is changed
      xobj.y = yoff1

      If (yy < 0.) then
        yy = yy * -1
      end if

      xobj.height = (zoff - ycen) + yy - (yy * (zscale))
      xobj.rotx = xrot
    Next
  Next
end sub

' *********************************************************************
' digital display
' *********************************************************************

Dim Digits(28)
Digits(0) = Array(a00,a01,a02,a03,a04,a05,a06,a07,a08)
Digits(1) = Array(a10,a11,a12,a13,a14,a15,a16,a17,a18)
Digits(2) = Array(a20,a21,a22,a23,a24,a25,a26,a27,a28)
Digits(3) = Array(a30,a31,a32,a33,a34,a35,a36,a37,a38)
Digits(4) = Array(a40,a41,a42,a43,a44,a45,a46,a47,a48)
Digits(5) = Array(a50,a51,a52,a53,a54,a55,a56,a57,a58)
Digits(6) = Array(a60,a61,a62,a63,a64,a65,a66,a67,a68)

Digits(7) = Array(b00,b01,b02,b03,b04,b05,b06,b07,b08)
Digits(8) = Array(b10,b11,b12,b13,b14,b15,b16,b17,b18)
Digits(9) = Array(b20,b21,b22,b23,b24,b25,b26,b27,b28)
Digits(10)  = Array(b30,b31,b32,b33,b34,b35,b36,b37,b38)
Digits(11)  = Array(b40,b41,b42,b43,b44,b45,b46,b47,b48)
Digits(12)  = Array(b50,b51,b52,b53,b54,b55,b56,b57,b58)
Digits(13)  = Array(b60,b61,b62,b63,b64,b65,b66,b67,b68)

Digits(14)  = Array(c00,c01,c02,c03,c04,c05,c06,c07,c08)
Digits(15)  = Array(c10,c11,c12,c13,c14,c15,c16,c17,c18)
Digits(16)  = Array(c20,c21,c22,c23,c24,c25,c26,c27,c28)
Digits(17)  = Array(c30,c31,c32,c33,c34,c35,c36,c37,c38)
Digits(18)  = Array(c40,c41,c42,c43,c44,c45,c46,c47,c48)
Digits(19)  = Array(c50,c51,c52,c53,c54,c55,c56,c57,c58)
Digits(20)  = Array(c60,c61,c62,c63,c64,c65,c66,c67,c68)

Digits(21)  = Array(d00,d01,d02,d03,d04,d05,d06,d07,d08)
Digits(22)  = Array(d10,d11,d12,d13,d14,d15,d16,d17,d18)
Digits(23)  = Array(d20,d21,d22,d23,d24,d25,d26,d27,d28)
Digits(24)  = Array(d30,d31,d32,d33,d34,d35,d36,d37,d38)
Digits(25)  = Array(d40,d41,d42,d43,d44,d45,d46,d47,d48)
Digits(26)  = Array(d50,d51,d52,d53,d54,d55,d56,d57,d58)
Digits(27)  = Array(d60,d61,d62,d63,d64,d65,d66,d67,d68)

Dim VRDigits(28)
VRDigits(0) = Array(LED1x0,LED1x1,LED1x2,LED1x3,LED1x4,LED1x5,LED1x6,LED1x7,LED1x8)
VRDigits(1) = Array(LED2x0,LED2x1,LED2x2,LED2x3,LED2x4,LED2x5,LED2x6,LED2x7,LED2x8)
VRDigits(2) = Array(LED3x0,LED3x1,LED3x2,LED3x3,LED3x4,LED3x5,LED3x6,LED3x7,LED3x8)
VRDigits(3) = Array(LED4x0,LED4x1,LED4x2,LED4x3,LED4x4,LED4x5,LED4x6,LED4x7,LED4x8)
VRDigits(4) = Array(LED5x0,LED5x1,LED5x2,LED5x3,LED5x4,LED5x5,LED5x6,LED5x7,LED5x8)
VRDigits(5) = Array(LED6x0,LED6x1,LED6x2,LED6x3,LED6x4,LED6x5,LED6x6,LED6x7,LED6x8)
VRDigits(6) = Array(LED7x0,LED7x1,LED7x2,LED7x3,LED7x4,LED7x5,LED7x6,LED7x7,LED7x8)

VRDigits(7) = Array(LED8x0,LED8x1,LED8x2,LED8x3,LED8x4,LED8x5,LED8x6,LED8x7,LED8x8)
VRDigits(8) = Array(LED9x0,LED9x1,LED9x2,LED9x3,LED9x4,LED9x5,LED9x6,LED9x7,LED9x8)
VRDigits(9) = Array(LED10x0,LED10x1,LED10x2,LED10x3,LED10x4,LED10x5,LED10x6,LED10x7,LED10x8)
VRDigits(10) = Array(LED11x0,LED11x1,LED11x2,LED11x3,LED11x4,LED11x5,LED11x6,LED11x7,LED11x8)
VRDigits(11) = Array(LED12x0,LED12x1,LED12x2,LED12x3,LED12x4,LED12x5,LED12x6,LED12x7,LED12x8)
VRDigits(12) = Array(LED13x0,LED13x1,LED13x2,LED13x3,LED13x4,LED13x5,LED13x6,LED13x7,LED13x8)
VRDigits(13) = Array(LED14x0,LED14x1,LED14x2,LED14x3,LED14x4,LED14x5,LED14x6,LED14x7,LED14x8)

VRDigits(14) = Array(LED1x000,LED1x001,LED1x002,LED1x003,LED1x004,LED1x005,LED1x006,LED1x007,LED1x008)
VRDigits(15) = Array(LED1x100,LED1x101,LED1x102,LED1x103,LED1x104,LED1x105,LED1x106,LED1x107,LED1x108)
VRDigits(16) = Array(LED1x200,LED1x201,LED1x202,LED1x203,LED1x204,LED1x205,LED1x206,LED1x207,LED1x208)
VRDigits(17) = Array(LED1x300,LED1x301,LED1x302,LED1x303,LED1x304,LED1x305,LED1x306,LED1x307,LED1x308)
VRDigits(18) = Array(LED1x400,LED1x401,LED1x402,LED1x403,LED1x404,LED1x405,LED1x406,LED1x407,LED1x408)
VRDigits(19) = Array(LED1x500,LED1x501,LED1x502,LED1x503,LED1x504,LED1x505,LED1x506,LED1x507,LED1x508)
VRDigits(20) = Array(LED1x600,LED1x601,LED1x602,LED1x603,LED1x604,LED1x605,LED1x606,LED1x607,LED1x608)

VRDigits(21) = Array(LED2x000,LED2x001,LED2x002,LED2x003,LED2x004,LED2x005,LED2x006,LED2x007,LED2x008)
VRDigits(22) = Array(LED2x100,LED2x101,LED2x102,LED2x103,LED2x104,LED2x105,LED2x106,LED2x107,LED2x108)
VRDigits(23) = Array(LED2x200,LED2x201,LED2x202,LED2x203,LED2x204,LED2x205,LED2x206,LED2x007,LED2x208)
VRDigits(24) = Array(LED2x300,LED2x301,LED2x302,LED2x303,LED2x304,LED2x305,LED2x306,LED2x007,LED2x308)
VRDigits(25) = Array(LED2x400,LED2x401,LED2x402,LED2x403,LED2x404,LED2x405,LED2x406,LED2x007,LED2x408)
VRDigits(26) = Array(LED2x500,LED2x501,LED2x502,LED2x503,LED2x504,LED2x505,LED2x506,LED2x007,LED2x508)
VRDigits(27) = Array(LED2x600,LED2x601,LED2x602,LED2x603,LED2x604,LED2x605,LED2x606,LED2x007,LED2x608)

dim DisplayColor
DisplayColor =  RGB(255,40,1)

Sub VRDisplayTimer
  Dim ii, jj, obj, b, x
  Dim ChgLED,num, chg, stat
  ChgLED=Controller.ChangedLEDs(&Hffffffff, &Hffffffff)
    If Not IsEmpty(ChgLED) Then
      For ii=0 To UBound(chgLED)
        num=chgLED(ii, 0) : chg=chgLED(ii, 1) : stat=chgLED(ii, 2)
        For Each obj In VRDigits(num)
 '                  If chg And 1 Then obj.visible=stat And 1    'if you use the object color for off; turn the display object visible to not visible on the playfield, and uncomment this line out.
          If chg And 1 Then FadeDisplay obj, stat And 1
          chg=chg\2 : stat=stat\2
        Next
      Next
    End If
End Sub

Sub FadeDisplay(object, onoff)
  If OnOff = 1 Then
    object.color = DisplayColor
  Else
    Object.Color = RGB(1,1,1)
  End If
End Sub

Sub InitDigits()
  dim tmp, x, obj
  for x = 0 to uBound(VRDigits)
    if IsArray(VRDigits(x) ) then
      For each obj in VRDigits(x)
        obj.height = obj.height + 18
        FadeDisplay obj, 0
      next
    end If
  Next
  DisplayColor = RGB(255,51,0)
End Sub

Sub UpdateLeds
    Dim ChgLED, ii, jj, num, chg, stat, obj, b, x
    ChgLED = Controller.ChangedLEDs(&Hffffffff, &Hffffffff)
    If Not IsEmpty(ChgLED)Then
        For ii = 0 To UBound(chgLED)
            num = chgLED(ii, 0):chg = chgLED(ii, 1):stat = chgLED(ii, 2)
      For Each obj In Digits(num)
        If chg And 1 Then obj.State = stat And 1
        chg = chg \ 2:stat = stat \ 2
      Next
        Next
    End If
End Sub

Sub SolGi(enabled)
    If enabled Then
        GiON
    Else
        GiOFF
    End If
End Sub

Sub GiON
  dim xx
    For each xx in GILights
        xx.State = LightStateOn
    Next
End Sub

Sub GiOFF
  dim xx
    For each xx in GILights
        xx.State = LightStateOff
    Next
End Sub

''*********************************************************************
''Sling Shot Animations
'' Rstep and Lstep  are the variables that increment the animation
''*********************************************************************
Dim RStep, Lstep, Lstep1

Sub RightSlingshot_slingshot
  RS.VelocityCorrect(ActiveBall)
  vpmTimer.PulseSw 22
  RandomSoundSlingshotRight
    RSling.Visible = 0
    RSling1.Visible = 1
    sling1.rotx = 14
    RStep = 0
    RightSlingShot.TimerEnabled = 1
End Sub

Sub RightSlingShot_Timer
    Select Case RStep
        Case 3:RSLing1.Visible = 0:RSLing2.Visible = 1:sling1.rotx = 7
        Case 4:RSLing2.Visible = 0:RSLing.Visible = 1:sling1.rotx = 0:RightSlingShot.TimerEnabled = 0
    End Select
    RStep = RStep + 1
End Sub

Sub LeftSlingshot_slingshot
  LS.VelocityCorrect(ActiveBall)
  vpmTimer.PulseSw 21
  RandomSoundSlingshotLeft
    LSling.Visible = 0
    LSling1.Visible = 1
    sling2.rotx = 14
    LStep = 0
    LeftSlingShot.TimerEnabled = 1
End Sub

Sub LeftSlingShot_Timer
    Select Case LStep
        Case 3:LSLing1.Visible = 0:LSLing2.Visible = 1:sling2.rotx = 7
        Case 4:LSLing2.Visible = 0:LSLing.Visible = 1:sling2.rotx = 0:LeftSlingShot.TimerEnabled = 0
    End Select
    LStep = LStep + 1
End Sub

Dim BumperOneIntensity, BumperTwoIntensity, BumperThreeIntensity, BumperFourIntensity
Dim LastBumperOneIntensity, LastBumperTwoIntensity, LastBumperThreeIntensity, LastBumperFourIntensity
LastBumperOneIntensity = -1
LastBumperTwoIntensity = -1
LastBumperThreeIntensity = -1
LastBumperFourIntensity = -1

Sub l70_Animate()
  BumperOneIntensity = l70.GetInPlayIntensity / 100
  If BumperOneIntensity <> LastBumperOneIntensity Then
    FlFadeBumper 1,BumperOneIntensity
    LastBumperOneIntensity = BumperOneIntensity
  End If
End Sub

Sub l86_Animate()
  BumperTwoIntensity = l86.GetInPlayIntensity / 100
  If BumperTwoIntensity <> LastBumperTwoIntensity Then
    FlFadeBumper 2,BumperTwoIntensity
    LastBumperTwoIntensity = BumperTwoIntensity
  End If
End Sub

Sub l55_Animate()
  BumperThreeIntensity = l55.GetInPlayIntensity / 100
  If BumperThreeIntensity <> LastBumperThreeIntensity Then
    FlFadeBumper 3,BumperThreeIntensity
    LastBumperThreeIntensity = BumperThreeIntensity
  End If
End Sub

Sub l71_Animate()
  BumperFourIntensity = l71.GetInPlayIntensity / 100
  If BumperFourIntensity <> LastBumperFourIntensity Then
    FlFadeBumper 4,BumperFourIntensity
    LastBumperFourIntensity = BumperFourIntensity
  End If
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
Dim objBallShadow(3)

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

'******************************************************
' ZPHY:  GENERAL ADVICE ON PHYSICS
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

'
''*******************************************
'' Late 70's to early 80's
'
'Sub InitPolarity()
'   dim x, a : a = Array(LF, RF)
' for each x in a
'   x.AddPt "Ycoef", 0, RightFlipper.Y-65, 1 'disabled
'   x.AddPt "Ycoef", 1, RightFlipper.Y-11, 1
'   x.enabled = True
'   x.TimeDelay = 80
'   x.DebugOn=False ' prints some info in debugger
'
'
'        x.AddPt "Polarity", 0, 0, 0
'        x.AddPt "Polarity", 1, 0.05, - 2.7
'        x.AddPt "Polarity", 2, 0.16, - 2.7
'        x.AddPt "Polarity", 3, 0.22, - 0
'        x.AddPt "Polarity", 4, 0.25, - 0
'        x.AddPt "Polarity", 5, 0.3, - 1
'        x.AddPt "Polarity", 6, 0.4, - 2
'        x.AddPt "Polarity", 7, 0.5, - 2.7
'        x.AddPt "Polarity", 8, 0.65, - 1.8
'        x.AddPt "Polarity", 9, 0.75, - 0.5
'        x.AddPt "Polarity", 10, 0.81, - 0.5
'        x.AddPt "Polarity", 11, 0.88, 0
'        x.AddPt "Polarity", 12, 1.3, 0
'
'   x.AddPt "Velocity", 0, 0, 0.85
'   x.AddPt "Velocity", 1, 0.15, 0.85
'   x.AddPt "Velocity", 2, 0.2, 0.9
'   x.AddPt "Velocity", 3, 0.23, 0.95
'   x.AddPt "Velocity", 4, 0.41, 0.95
'   x.AddPt "Velocity", 5, 0.53, 0.95 '0.982
'   x.AddPt "Velocity", 6, 0.62, 1.0
'   x.AddPt "Velocity", 7, 0.702, 0.968
'   x.AddPt "Velocity", 8, 0.95,  0.968
'   x.AddPt "Velocity", 9, 1.03,  0.945
'   x.AddPt "Velocity", 10, 1.5,  0.945
'
' Next
'
' ' SetObjects arguments: 1: name of object 2: flipper object: 3: Trigger object around flipper
'    LF.SetObjects "LF", LeftFlipper, TriggerLF
'    RF.SetObjects "RF", RightFlipper, TriggerRF
'End Sub
'
'
'
''*******************************************
'' Mid 80's
'
'Sub InitPolarity()
'   dim x, a : a = Array(LF, RF)
' for each x in a
'   x.AddPt "Ycoef", 0, RightFlipper.Y-65, 1 'disabled
'   x.AddPt "Ycoef", 1, RightFlipper.Y-11, 1
'   x.enabled = True
'   x.TimeDelay = 80
'   x.DebugOn=False ' prints some info in debugger
'
'   x.AddPt "Polarity", 0, 0, 0
'   x.AddPt "Polarity", 1, 0.05, - 3.7
'   x.AddPt "Polarity", 2, 0.16, - 3.7
'   x.AddPt "Polarity", 3, 0.22, - 0
'   x.AddPt "Polarity", 4, 0.25, - 0
'   x.AddPt "Polarity", 5, 0.3, - 2
'   x.AddPt "Polarity", 6, 0.4, - 3
'   x.AddPt "Polarity", 7, 0.5, - 3.7
'   x.AddPt "Polarity", 8, 0.65, - 2.3
'   x.AddPt "Polarity", 9, 0.75, - 1.5
'   x.AddPt "Polarity", 10, 0.81, - 1
'   x.AddPt "Polarity", 11, 0.88, 0
'   x.AddPt "Polarity", 12, 1.3, 0
'
'   x.AddPt "Velocity", 0, 0, 0.85
'   x.AddPt "Velocity", 1, 0.15, 0.85
'   x.AddPt "Velocity", 2, 0.2, 0.9
'   x.AddPt "Velocity", 3, 0.23, 0.95
'   x.AddPt "Velocity", 4, 0.41, 0.95
'   x.AddPt "Velocity", 5, 0.53, 0.95 '0.982
'   x.AddPt "Velocity", 6, 0.62, 1.0
'   x.AddPt "Velocity", 7, 0.702, 0.968
'   x.AddPt "Velocity", 8, 0.95,  0.968
'   x.AddPt "Velocity", 9, 1.03,  0.945
'   x.AddPt "Velocity", 10, 1.5,  0.945
'
' Next
'
' ' SetObjects arguments: 1: name of object 2: flipper object: 3: Trigger object around flipper
'    LF.SetObjects "LF", LeftFlipper, TriggerLF
'    RF.SetObjects "RF", RightFlipper, TriggerRF
'End Sub
'
''*******************************************
''  Late 80's early 90's
'
Sub InitPolarity()
  dim x, a : a = Array(LF, RF)
  for each x in a
    x.AddPt "Ycoef", 0, RightFlipper.Y-65, 1 'disabled
    x.AddPt "Ycoef", 1, RightFlipper.Y-11, 1
    x.enabled = True
    x.TimeDelay = 60
    x.DebugOn=False ' prints some info in debugger

    x.AddPt "Polarity", 0, 0, 0
    x.AddPt "Polarity", 1, 0.05, - 5
    x.AddPt "Polarity", 2, 0.16, - 5
    x.AddPt "Polarity", 3, 0.22, - 0
    x.AddPt "Polarity", 4, 0.25, - 0
    x.AddPt "Polarity", 5, 0.3, - 2
    x.AddPt "Polarity", 6, 0.4, - 3
    x.AddPt "Polarity", 7, 0.5, - 4.0
    x.AddPt "Polarity", 8, 0.7, - 3.5
    x.AddPt "Polarity", 9, 0.75, - 3.0
    x.AddPt "Polarity", 10, 0.8, - 2.5
    x.AddPt "Polarity", 11, 0.85, - 2.0
    x.AddPt "Polarity", 12, 0.9, - 1.5
    x.AddPt "Polarity", 13, 0.95, - 1.0
    x.AddPt "Polarity", 14, 1, - 0.5
    x.AddPt "Polarity", 15, 1.1, 0
    x.AddPt "Polarity", 16, 1.3, 0

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

'*******************************************
' Early 90's and after
'
'Sub InitPolarity()
' Dim x, a
' a = Array(LF, RF)
' For Each x In a
'   x.AddPt "Ycoef", 0, RightFlipper.Y-65, 1 'disabled
'   x.AddPt "Ycoef", 1, RightFlipper.Y-11, 1
'   x.enabled = True
'   x.TimeDelay = 60
'   x.DebugOn=False ' prints some info in debugger
'
'   x.AddPt "Polarity", 0, 0, 0
'   x.AddPt "Polarity", 1, 0.05, - 5.5
'   x.AddPt "Polarity", 2, 0.16, - 5.5
'   x.AddPt "Polarity", 3, 0.20, - 0.75
'   x.AddPt "Polarity", 4, 0.25, - 1.25
'   x.AddPt "Polarity", 5, 0.3, - 1.75
'   x.AddPt "Polarity", 6, 0.4, - 3.5
'   x.AddPt "Polarity", 7, 0.5, - 5.25
'   x.AddPt "Polarity", 8, 0.7, - 4.0
'   x.AddPt "Polarity", 9, 0.75, - 3.5
'   x.AddPt "Polarity", 10, 0.8, - 3.0
'   x.AddPt "Polarity", 11, 0.85, - 2.5
'   x.AddPt "Polarity", 12, 0.9, - 2.0
'   x.AddPt "Polarity", 13, 0.95, - 1.5
'   x.AddPt "Polarity", 14, 1, - 1.0
'   x.AddPt "Polarity", 15, 1.05, -0.5
'   x.AddPt "Polarity", 16, 1.1, 0
'   x.AddPt "Polarity", 17, 1.3, 0
'
'   x.AddPt "Velocity", 0, 0, 0.85
'   x.AddPt "Velocity", 1, 0.23, 0.85
'   x.AddPt "Velocity", 2, 0.27, 1
'   x.AddPt "Velocity", 3, 0.3, 1
'   x.AddPt "Velocity", 4, 0.35, 1
'   x.AddPt "Velocity", 5, 0.6, 1 '0.982
'   x.AddPt "Velocity", 6, 0.62, 1.0
'   x.AddPt "Velocity", 7, 0.702, 0.968
'   x.AddPt "Velocity", 8, 0.95,  0.968
'   x.AddPt "Velocity", 9, 1.03,  0.945
'   x.AddPt "Velocity", 10, 1.5,  0.945
'
' Next
'
' ' SetObjects arguments: 1: name of object 2: flipper object: 3: Trigger object around flipper
' LF.SetObjects "LF", LeftFlipper, TriggerLF
' RF.SetObjects "RF", RightFlipper, TriggerRF
'End Sub

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
  'Dim BOT
  'BOT = GetBalls

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
    'BOT = GetBalls

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

'Add targets or posts to the TargetBounce collection if you want to activate the targetbouncer code from them
Sub TargetBounce_Hit(idx)
  TargetBouncer ActiveBall, 1
End Sub

'******************************************************
' ZSSC: SLINGSHOT CORRECTION FUNCTIONS by apophis
'******************************************************
' To add these slingshot corrections:
'  - On the table, add the endpoint primitives that define the two ends of the Slingshot
'  - Initialize the SlingshotCorrection objects in InitSlingCorrection
'  - Call the .VelocityCorrect methods from the respective _Slingshot event sub

Dim LS
Set LS = New SlingshotCorrection
Dim RS
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

'' The following sub are needed, however they may exist somewhere else in the script. Uncomment below if needed
'Dim PI: PI = 4*Atn(1)
'Function dSin(degrees)
' dsin = sin(degrees * Pi/180)
'End Function
'Function dCos(degrees)
' dcos = cos(degrees * Pi/180)
'End Function
'
'Function RotPoint(x,y,angle)
' dim rx, ry
' rx = x*dCos(angle) - y*dSin(angle)
' ry = x*dSin(angle) + y*dCos(angle)
' RotPoint = Array(rx,ry)
'End Function

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
'   ZRDT:  DROP TARGETS by Rothbauerw
'******************************************************
' The Stand Up and Drop Target solutions improve the physics for targets to create more realistic behavior. It allows the ball
' to move through the target enabling the ability to score more than one target with a well placed shot.
' It also handles full target animation, switch handling and deflection on hit. For drop targets there is also a slight lift when
' the drop targets raise, bricking, and popping the ball up if it's over the drop target when it raises.
'
' Add a Timers named DTAnim and STAnim to editor to handle drop & standup target animations, or run them off an always-on 10ms timer (GameTimer)
' DTAnim.interval = 10
' DTAnim.enabled = True

' Sub DTAnim_Timer
'   DoDTAnim
' DoSTAnim
' End Sub

' For each drop target, we'll use two wall objects for physics calculations and one primitive for visuals and
' animation. We will not use target objects.  Place your drop target primitive the same as you would a VP drop target.
' The primitive should have it's pivot point centered on the x and y axis and at or just below the playfield
' level on the z axis. Orientation needs to be set using Rotz and bending deflection using Rotx. You'll find a hooded
' target mesh in this table's example. It uses the same texture map as the VP drop targets.
'
' For each stand up target we'll use a vp target, a laid back collidable primitive, and one primitive for visuals and animation.
' The visual primitive should should have it's pivot point centered on the x and y axis and the z should be at or just below the playfield.
' The target should animate backwards using transy.
'
' To create visual target primitives that work with the stand up and drop target code, follow the below instructions:
' (Other methods will work as well, but this is easy for even non-blender users to do)
' 1) Open a new blank table. Delete everything off the table in editor.
' 2) Copy and paste the VP target from your table into this blank table.
' 3) Place the target at x = 0, y = 0  (upper left hand corner) with an orientation of 0 (target facing the front of the table)
' 4) Under the file menu, select Export "OBJ Mesh"
' 5) Go to "https://threejs.org/editor/". Here you can modify the exported obj file. When you export, it exports your target and also
'    the playfield mesh. You need to delete the playfield mesh here. Under the file menu, chose import, and select the obj you exported
'    from VPX. In the right hand panel, find the Playfield object and click on it and delete. Then use the file menu to Export OBJ.
' 6) In VPX, you can add a primitive and use "Import Mesh" to import the exported obj from the previous step. X,Y,Z scale should be 1.
'    The primitive will use the same target texture as the VP target object.
'
' * Note, each target must have a unique switch number. If they share a same number, add 100 to additional target with that number.
' For example, three targets with switch 32 would use 32, 132, 232 for their switch numbers.
' The 100 and 200 will be removed when setting the switch value for the target.

'******************************************************
'  DROP TARGETS INITIALIZATION
'******************************************************

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

'Define a variable for each drop target
Dim DT33, DT34, DT35

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

Set DT33 = (new DropTarget)(sw33, sw33a, sw33p, 33, 0, false)
Set DT34 = (new DropTarget)(sw34, sw34a, sw34p, 34, 0, false)
Set DT35 = (new DropTarget)(sw35, sw35a, sw35p, 35, 0, false)

Dim DTArray
DTArray = Array(DT33, DT34, DT35)

'Configure the behavior of Drop Targets.
Const DTDropSpeed = 80 'in milliseconds
Const DTDropUpSpeed = 40 'in milliseconds
Const DTDropUnits = 44 'VP units primitive drops so top of at or below the playfield
Const DTDropUpUnits = 10 'VP units primitive raises above the up position on drops up
Const DTMaxBend = 8 'max degrees primitive rotates when hit
Const DTDropDelay = 20 'time in milliseconds before target drops (due to friction/impact of the ball)
Const DTRaiseDelay = 40 'time in milliseconds before target drops back to normal up position after the solenoid fires to raise the target
Const DTBrickVel = 30 'velocity at which the target will brick, set to '0' to disable brick

Const DTEnableBrick = 0 'Set to 0 to disable bricking, 1 to enable bricking
Const DTHitSound = "" 'Drop Target Hit sound
Const DTDropSound = "DropTarget_Down" 'Drop Target Drop sound
Const DTResetSound = "DropTarget_Up" 'Drop Target reset sound

Const DTMass = 0.2 'Mass of the Drop Target (between 0 and 1), higher values provide more resistance


'******************************************************
'  DROP TARGETS FUNCTIONS
'******************************************************

Sub DTHit(switch)
  Dim i
  i = DTArrayID(switch)

  PlayTargetSound
  DTArray(i).animate = DTCheckBrick(ActiveBall,DTArray(i).prim)
  If DTArray(i).animate = 1 Or DTArray(i).animate = 3 Or DTArray(i).animate = 4 Then
    DTBallPhysics ActiveBall, DTArray(i).prim.rotz, DTMass
  End If
  DoDTAnim
End Sub

Sub DTRaise(switch)
  Dim i
  i = DTArrayID(switch)

  DTArray(i).animate =  - 1
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
  For i = 0 To UBound(DTArray)
    If DTArray(i).sw = switch Then
      DTArrayID = i
      Exit Function
    End If
  Next
End Function

Sub DTBallPhysics(aBall, angle, mass)
  Dim rangle,bangle,calc1, calc2, calc3
  rangle = (angle - 90) * 3.1416 / 180
  bangle = atn2(cor.ballvely(aball.id),cor.ballvelx(aball.id))

  calc1 = cor.BallVel(aball.id) * Cos(bangle - rangle) * (aball.mass - mass) / (aball.mass + mass)
  calc2 = cor.BallVel(aball.id) * Sin(bangle - rangle) * Cos(rangle + 4 * Atn(1) / 2)
  calc3 = cor.BallVel(aball.id) * Sin(bangle - rangle) * Sin(rangle + 4 * Atn(1) / 2)

  aBall.velx = calc1 * Cos(rangle) + calc2
  aBall.vely = calc1 * Sin(rangle) + calc3
End Sub

'Check if target is hit on it's face or sides and whether a 'brick' occurred
Function DTCheckBrick(aBall, dtprim)
  Dim bangle, bangleafter, rangle, rangle2, Xintersect, Yintersect, cdist, perpvel, perpvelafter, paravel, paravelafter
  rangle = (dtprim.rotz - 90) * 3.1416 / 180
  rangle2 = dtprim.rotz * 3.1416 / 180
  bangle = atn2(cor.ballvely(aball.id),cor.ballvelx(aball.id))
  bangleafter = Atn2(aBall.vely,aball.velx)

  Xintersect = (aBall.y - dtprim.y - Tan(bangle) * aball.x + Tan(rangle2) * dtprim.x) / (Tan(rangle2) - Tan(bangle))
  Yintersect = Tan(rangle2) * Xintersect + (dtprim.y - Tan(rangle2) * dtprim.x)

  cdist = Distance(dtprim.x, dtprim.y, Xintersect, Yintersect)

  perpvel = cor.BallVel(aball.id) * Cos(bangle - rangle)
  paravel = cor.BallVel(aball.id) * Sin(bangle - rangle)

  perpvelafter = BallSpeed(aBall) * Cos(bangleafter - rangle)
  paravelafter = BallSpeed(aBall) * Sin(bangleafter - rangle)

  If perpvel > 0 And  perpvelafter <= 0 Then
    If DTEnableBrick = 1 And  perpvel > DTBrickVel And DTBrickVel <> 0 And cdist < 8 Then
      DTCheckBrick = 3
    Else
      DTCheckBrick = 1
    End If
  ElseIf perpvel > 0 And ((paravel > 0 And paravelafter > 0) Or (paravel < 0 And paravelafter < 0)) Then
    DTCheckBrick = 4
  Else
    DTCheckBrick = 0
  End If
End Function

Sub DoDTAnim()
  Dim i
  For i = 0 To UBound(DTArray)
    DTArray(i).animate = DTAnimate(DTArray(i).primary,DTArray(i).secondary,DTArray(i).prim,DTArray(i).sw,DTArray(i).animate)
  Next
End Sub

Function DTAnimate(primary, secondary, prim, switch, animate)
  Dim transz, switchid
  Dim animtime, rangle

  switchid = switch

  Dim ind
  ind = DTArrayID(switchid)

  rangle = prim.rotz * PI / 180

  DTAnimate = animate

  If animate = 0 Then
    primary.uservalue = 0
    DTAnimate = 0
    Exit Function
  ElseIf primary.uservalue = 0 Then
    primary.uservalue = GameTime
  End If

  animtime = GameTime - primary.uservalue

  If (animate = 1 Or animate = 4) And animtime < DTDropDelay Then
    primary.collidable = 0
    If animate = 1 Then secondary.collidable = 1 Else secondary.collidable = 0
    prim.rotx = DTMaxBend * Cos(rangle)
    prim.roty = DTMaxBend * Sin(rangle)
    DTAnimate = animate
    Exit Function
  ElseIf (animate = 1 Or animate = 4) And animtime > DTDropDelay Then
    primary.collidable = 0
    If animate = 1 Then secondary.collidable = 1 Else secondary.collidable = 1 'If animate = 1 Then secondary.collidable = 1 Else secondary.collidable = 0 'updated by rothbauerw to account for edge case
    prim.rotx = DTMaxBend * Cos(rangle)
    prim.roty = DTMaxBend * Sin(rangle)
    animate = 2
    SoundDropTargetDrop prim
  End If

  If animate = 2 Then
    transz = (animtime - DTDropDelay) / DTDropSpeed * DTDropUnits *  - 1
    If prim.transz >  - DTDropUnits  Then
      prim.transz = transz
    End If

    prim.rotx = DTMaxBend * Cos(rangle) / 2
    prim.roty = DTMaxBend * Sin(rangle) / 2

    If prim.transz <= - DTDropUnits Then
      prim.transz =  - DTDropUnits
      secondary.collidable = 0
      DTArray(ind).isDropped = True 'Mark target as dropped
      controller.Switch(Switchid mod 100) = 1
      primary.uservalue = 0
      DTAnimate = 0
      Exit Function
    Else
      DTAnimate = 2
      Exit Function
    End If
  End If

  If animate = 3 And animtime < DTDropDelay Then
    primary.collidable = 0
    secondary.collidable = 1
    prim.rotx = DTMaxBend * Cos(rangle)
    prim.roty = DTMaxBend * Sin(rangle)
  ElseIf animate = 3 And animtime > DTDropDelay Then
    primary.collidable = 1
    secondary.collidable = 0
    prim.rotx = 0
    prim.roty = 0
    primary.uservalue = 0
    DTAnimate = 0
    Exit Function
  End If

  If animate =  - 1 Then
    transz = (1 - (animtime) / DTDropUpSpeed) * DTDropUnits *  - 1

    If prim.transz =  - DTDropUnits Then
      Dim b
      'Dim gBOT
      'gBOT = GetBalls

      For b = 0 To UBound(gBOT)
        If InRotRect(gBOT(b).x,gBOT(b).y,prim.x, prim.y, prim.rotz, - 25, - 10,25, - 10,25,25, - 25,25) And gBOT(b).z < prim.z + DTDropUnits + 25 Then
          gBOT(b).velz = 20
        End If
      Next
    End If

    If prim.transz < 0 Then
      prim.transz = transz
    ElseIf transz > 0 Then
      prim.transz = transz
    End If

    If prim.transz > DTDropUpUnits Then
      DTAnimate =  - 2
      prim.transz = DTDropUpUnits
      prim.rotx = 0
      prim.roty = 0
      primary.uservalue = GameTime
    End If
    primary.collidable = 0
    secondary.collidable = 1
    DTArray(ind).isDropped = False 'Mark target as not dropped
    controller.Switch(Switchid mod 100) = 0
  End If

  If animate =  - 2 And animtime > DTRaiseDelay Then
    prim.transz = (animtime - DTRaiseDelay) / DTDropSpeed * DTDropUnits *  - 1 + DTDropUpUnits
    If prim.transz < 0 Then
      prim.transz = 0
      primary.uservalue = 0
      DTAnimate = 0

      primary.collidable = 1
      secondary.collidable = 0
    End If
  End If
End Function

Function DTDropped(switchid)
  Dim ind
  ind = DTArrayID(switchid)

  DTDropped = DTArray(ind).isDropped
End Function


'******************************************************
'****  END DROP TARGETS
'******************************************************

'******************************************************
' ZRST: STAND-UP TARGET INITIALIZATION
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

'Define a variable for each stand-up target
Dim ST1, ST2, ST3, ST4, ST25, ST26, ST27, ST28, ST29, ST30

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

Set ST1 = (new StandupTarget)(sw1, sw1p, 1, 0)
Set ST2 = (new StandupTarget)(sw2, sw2p, 2, 0)
Set ST3 = (new StandupTarget)(sw3, sw3p, 3, 0)
Set ST4 = (new StandupTarget)(sw4, sw4p, 4, 0)
Set ST25 = (new StandupTarget)(sw25, sw25p, 25, 0)
Set ST26 = (new StandupTarget)(sw26, sw26p, 26, 0)
Set ST27 = (new StandupTarget)(sw27, sw27p, 27, 0)
Set ST28 = (new StandupTarget)(sw28, sw28p, 28, 0)
Set ST29 = (new StandupTarget)(sw29, sw29p, 29, 0)
Set ST30 = (new StandupTarget)(sw30, sw30p, 30, 0)

'Add all the Stand-up Target Arrays to Stand-up Target Animation Array
' STAnimationArray = Array(ST1, ST2, ....)
Dim STArray
STArray = Array(ST1, ST2, ST3, ST4, ST25, ST26, ST27, ST28, ST29, ST30)

'Configure the behavior of Stand-up Targets
Const STAnimStep =  1.5     'vpunits per animation step (control return to Start)
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
    vpmTimer.PulseSw switch mod 100
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


'******************************************************
'***  END STAND-UP TARGETS
'******************************************************

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
  'Dim BOT
  'BOT = GetBalls

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
      PlaySound ("BallRoll_" & b), - 1, VolPlayfieldRoll(gBOT(b)) * BallRollVolume * MechVolume, AudioPan(gBOT(b)), 0, PitchPlayfieldRoll(gBOT(b)), 1, 0, AudioFade(gBOT(b))
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
'****  END BALL ROLLING AND DROP SOUNDS
'******************************************************

'******************************************************
'**** RAMP ROLLING SFX
'******************************************************

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
          PlaySound("RampLoop" & x), -1, VolPlayfieldRoll(RampBalls(x,0)) * RampRollVolume * MechVolume, AudioPan(RampBalls(x,0)), 0, BallPitchV(RampBalls(x,0)), 1, 0, AudioFade(RampBalls(x,0))
          StopSound("wireloop" & x)
        Else
          StopSound("RampLoop" & x)
          PlaySound("wireloop" & x), -1, VolPlayfieldRoll(RampBalls(x,0)) * RampRollVolume * MechVolume, AudioPan(RampBalls(x,0)), 0, BallPitch(RampBalls(x,0)), 1, 0, AudioFade(RampBalls(x,0))
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

'/////////////////////////////  PLASTIC RAMPS FLAPS - SOUNDS  ////////////////////////////
Sub RandomSoundRampFlapUp()
' debug.print "flap up"
  Select Case Int(Rnd*3)+1
    Case 1 : PlaySoundAtLevelActiveBall ("TOM_Ramp_Flap_Up_1"), FlapSoundLevel
    Case 2 : PlaySoundAtLevelActiveBall ("TOM_Ramp_Flap_Up_2"), FlapSoundLevel
    Case 3 : PlaySoundAtLevelActiveBall ("TOM_Ramp_Flap_Up_3"), FlapSoundLevel
  End Select
End Sub

Sub RandomSoundRampFlapDown()
' debug.print "flap down"
  Select Case Int(Rnd*3)+1
    Case 1 : PlaySoundAtLevelActiveBall ("TOM_Ramp_Flap_Down_1"), FlapSoundLevel
    Case 2 : PlaySoundAtLevelActiveBall ("TOM_Ramp_Flap_Down_2"), FlapSoundLevel
    Case 3 : PlaySoundAtLevelActiveBall ("TOM_Ramp_Flap_Down_3"), FlapSoundLevel
  End Select
End Sub


'/////////////////////////////  RAMP COLLISIONS  ////////////////////////////

dim LRHit1_volume, LRHit2_volume, LRHit3_volume
dim RRHit1_volume, RRHit2_volume, RRHit3_volume

Dim LeftRampSoundLevel
Dim RightRampSoundLevel
Dim RampFallbackSoundLevel

Dim FlapSoundLevel

'///////////////////////-----Ramps-----///////////////////////
'///////////////////////-----Plastic Ramps-----///////////////////////
LeftRampSoundLevel = 0.1                        'volume level; range [0, 1]
RightRampSoundLevel = 0.1                       'volume level; range [0, 1]
RampFallbackSoundLevel = 0.2                      'volume level; range [0, 1]

'///////////////////////-----Ramp Flaps-----///////////////////////
FlapSoundLevel = 0.8                          'volume level; range [0, 1]

'******************************************************
'**** END RAMP ROLLING SFX
'******************************************************

'////////////////////////////////////////////////////////////////////////////////
'////          Mechanical Sounds, by Fleep                                   ////
'////                     Last Updated: January, 2022                        ////
'////////////////////////////////////////////////////////////////////////////////
'
'/////////////////////////////////  CARTRIDGES  /////////////////////////////////
'
'//  General Mechanical Sounds Cartridges:
Const Cartridge_Bumpers         = "WS_PBT_REV01" 'Williams Pinbot Cartridge REV01
Const Cartridge_Slingshots        = "WS_PBT_REV01" 'Williams Pinbot Cartridge REV01
Const Cartridge_Flippers        = "WS_PBT_REV01" 'Williams Pinbot Cartridge REV01
Const Cartridge_Kickers         = "WS_WHD_REV01"
Const Cartridge_Diverters       = "WS_DNR_REV01" 'Williams Diner Cartridge REV01
Const Cartridge_Knocker         = "WS_WHD_REV02" 'Williams Whirlwind Cartridge REV02
Const Cartridge_Relays          = "WS_WHD_REV01"
Const Cartridge_Trough          = "WS_WHD_REV01"
Const Cartridge_Rollovers       = "WS_WHD_REV01"
Const Cartridge_Targets         = "WS_WHD_REV01"
Const Cartridge_Gates         = "WS_WHD_REV01"
Const Cartridge_Spinner         = "SY_TNA_REV01" 'Spooky Total Nuclear Annihilation Cartridge REV01
Const Cartridge_Rubber_Hits       = "WS_WHD_REV01"
Const Cartridge_Metal_Hits        = "WS_WHD_REV01"
Const Cartridge_Plastic_Hits      = "WS_WHD_REV01"
Const Cartridge_Wood_Hits       = "WS_WHD_REV01"
Const Cartridge_Cabinet_Sounds      = "WS_WHD_REV01"
Const Cartridge_Drain         = "WS_WHD_REV01"
Const Cartridge_Apron         = "WS_WHD_REV01"
Const Cartridge_Ball_Roll       = "BY_TOM_REV01" 'Bally Theatre of Magic Cartridge REV01
Const Cartridge_BallBallCollision   = "BY_WDT_REV01" 'Bally WHO Dunnit Cartridge REV01
Const Cartridge_Ball_Drop_Bump      = "WS_WHD_REV01"
Const Cartridge_Plastic_Ramps     = "WS_WHD_REV01"
Const Cartridge_Metal_Ramps       = "WS_WHD_REV01"
Const Cartridge_Ball_Guides       = "WS_WHD_REV01"
Const Cartridge_Table_Specifics     = "WS_WHD_REV01"

'////////////////////////////  SOUND SOURCE CREDITS  ////////////////////////////
'//  Special thanks go to the following contributors who have provided audio
'//  footage recordings:
'//
'//  Williams Whirlwind - Blackmoor, wrd1972
'//  Williams Diner - Nick Rusis
'//  Spooky Total Nuclear Annihilation - WildDogArcade, Ed and Gary
'//  Bally Theatre of Magic - CalleV, nickbuol
'//  Bally WHO Dunnit - Amazaley1
'//  Williams Pinbot - major_drain_pinball

'///////////////////////////  SOLENOIDS (COILS) CONFIG  /////////////////////////

'//  FLIPPER COILS:
'//  Flippers in this table: Lower Left Flipper, Lower Right Flipper, Upper Right Fliiper
Dim FlipperUpSoundLevel, FlipperDownSoundLevel, FlipperUpAttackMinimumSoundLevel, FlipperUpAttackMaximumSoundLevel
Dim FlipperLeftLowerHitParm, FlipperRightUpperHitParm, FlipperRightLowerHitParm

'//  Flipper Up Attacks initialize during playsound subs
Dim FlipperUpAttackLeftSoundLevel, FlipperUpAttackRightSoundLevel

FlipperUpSoundLevel = 1
FlipperDownSoundLevel = 0.65
FlipperUpAttackMinimumSoundLevel = 0.010
FlipperUpAttackMaximumSoundLevel = 0.435

'//  Flipper Hit Param initialize with FlipperUpSoundLevel
'//  and dynamically modified calculated by ball flipper collision
FlipperLeftLowerHitParm = FlipperUpSoundLevel
FlipperRightLowerHitParm = FlipperUpSoundLevel
FlipperRightUpperHitParm = FlipperUpSoundLevel

Dim Solenoid_OutholeKicker_SoundLevel, Solenoid_ShooterFeeder_SoundLevel
Dim Solenoid_RightRampLifter_SoundLevel, Solenoid_LeftLockingKickback_SoundLevel
Dim Solenoid_TopEject_SoundLevel, Solenoid_Knocker_SoundLevel, Solenoid_DropTargetReset_SoundLevel
Dim Solenoid_Diverter_Enabled_SoundLevel, Solenoid_Diverter_Hold_SoundLevel, Solenoid_Diverter_Disabled_SoundLevel
Dim Solenoid_UnderPlayfieldKickbig_SoundLevel, Solenoid_Bumper_SoundMultiplier
Dim Solenoid_Slingshot_SoundLevel, Solenoid_RightRampDown_SoundLevel, AutoPlungerSoundLevel

AutoPlungerSoundLevel = 1                       'volume level; range [0, 1]
Solenoid_OutholeKicker_SoundLevel = 1
Solenoid_ShooterFeeder_SoundLevel = 1
Solenoid_RightRampLifter_SoundLevel = 0.3
Solenoid_RightRampDown_SoundLevel = 0.3
Solenoid_LeftLockingKickback_SoundLevel = 1
Solenoid_TopEject_SoundLevel = 1
Solenoid_Knocker_SoundLevel = 1
Solenoid_DropTargetReset_SoundLevel = 1
Solenoid_Diverter_Enabled_SoundLevel = 1
Solenoid_Diverter_Hold_SoundLevel = 0.7
Solenoid_Diverter_Disabled_SoundLevel = 0.4
Solenoid_UnderPlayfieldKickbig_SoundLevel = 1
Solenoid_Bumper_SoundMultiplier = 0.004 '8
Solenoid_Slingshot_SoundLevel = 1

Dim RelayLowerGISoundLevel, RelayUpperGISoundLevel, RelaySolenoidACSelectSoundLevel, RelayFlasherSoundLevel
RelayLowerGISoundLevel = 0.45
RelayUpperGISoundLevel = 0.45
RelaySolenoidACSelectSoundLevel = 0.3
RelayFlasherSoundLevel = 0.015

Dim Solenoid_BlowerMotor_SoundLevel, Solenoid_SpinWheelsMotor_SoundLevel
Solenoid_BlowerMotor_SoundLevel = 0.2
Solenoid_SpinWheelsMotor_SoundLevel = 0.2

'////////////////////////////  SWITCHES SOUND CONFIG  ///////////////////////////
Dim Switch_Gate_SoundLevel, SpinnerSoundLevel, RolloverSoundLevel, OutLaneRolloverSoundLevel, TargetSoundFactor

Switch_Gate_SoundLevel = 1
SpinnerSoundLevel = 0.1
RolloverSoundLevel = 0.55
OutLaneRolloverSoundLevel = 0.8
TargetSoundFactor = 0.8

'////////////////////  BALL HITS, BUMPS, DROPS SOUND CONFIG  ////////////////////
Dim BallWithBallCollisionSoundFactor, BallBouncePlayfieldSoftFactor, BallBouncePlayfieldHardFactor
Dim WallImpactSoundFactor, MetalImpactSoundFactor, WireformAntiRebountRailSoundFactor
Dim RubberStrongSoundFactor, RubberWeakSoundFactor, RubberFlipperSoundFactor
Dim BottomArchBallGuideSoundFactor, FlipperBallGuideSoundFactor, OutlaneWallsSoundFactor
Dim EjectBallBumpSoundLevel, HeadSaucerSoundLevel, EjectHoleEnterSoundLevel
Dim RightRampMetalWireDropToPlayfieldSoundLevel, LeftPlasticRampDropToLockSoundLevel, LeftPlasticRampDropToPlayfieldSoundLevel
Dim CellarLeftEnterSoundLevel, CellarRightEnterSoundLevel, CellerKickouBallDroptSoundLevel

BallWithBallCollisionSoundFactor = 3.2
BallBouncePlayfieldSoftFactor = 0.0015
BallBouncePlayfieldHardFactor = 0.0075
WallImpactSoundFactor = 0.075
MetalImpactSoundFactor = 0.075
RubberStrongSoundFactor = 0.045
RubberWeakSoundFactor = 0.055
RubberFlipperSoundFactor = 0.65
BottomArchBallGuideSoundFactor = 0.2
FlipperBallGuideSoundFactor = 0.015
WireformAntiRebountRailSoundFactor = 0.04
OutlaneWallsSoundFactor = 1
EjectBallBumpSoundLevel = 1
RightRampMetalWireDropToPlayfieldSoundLevel = 1
LeftPlasticRampDropToLockSoundLevel = 1
LeftPlasticRampDropToPlayfieldSoundLevel = 1
EjectHoleEnterSoundLevel = 0.75
HeadSaucerSoundLevel = 0.15
CellerKickouBallDroptSoundLevel = 1
CellarLeftEnterSoundLevel = 0.85
CellarRightEnterSoundLevel = 0.85

'///////////////////////  OTHER PLAYFIELD ELEMENTS CONFIG  //////////////////////
Dim RollingSoundFactor, RollingOnDiscSoundFactor, BallReleaseShooterLaneSoundLevel
Dim LeftPlasticRampEnteranceSoundLevel, RightPlasticRampEnteranceSoundLevel
Dim LeftPlasticRampRollSoundFactor, RightPlasticRampRollSoundFactor
Dim LeftMetalWireRampRollSoundFactor, RightPlasticRampHitsSoundLevel, LeftPlasticRampHitsSoundLevel
Dim SpinningDiscRolloverSoundFactor, SpinningDiscRolloverBumpSoundLevel
Dim LaneSoundFactor, LaneEnterSoundFactor, InnerLaneSoundFactor
Dim LaneLoudImpactMinimumSoundLevel, LaneLoudImpactMaximumSoundLevel
Dim GateSoundLevel

GateSoundLevel = 0.5 / 5      'volume level; range [0, 1]
RollingSoundFactor = 50
RollingOnDiscSoundFactor = 1.5
BallReleaseShooterLaneSoundLevel = 1
LeftPlasticRampEnteranceSoundLevel = 0.1
RightPlasticRampEnteranceSoundLevel = 0.1
LeftPlasticRampRollSoundFactor = 0.2
RightPlasticRampRollSoundFactor = 0.2
LeftMetalWireRampRollSoundFactor = 1
RightPlasticRampHitsSoundLevel = 1
LeftPlasticRampHitsSoundLevel = 1
SpinningDiscRolloverSoundFactor = 0.05
SpinningDiscRolloverBumpSoundLevel = 0.3
LaneEnterSoundFactor = 0.9
InnerLaneSoundFactor = 0.0005
LaneSoundFactor = 0.0004
LaneLoudImpactMinimumSoundLevel = 0
LaneLoudImpactMaximumSoundLevel = 0.4


'///////////////////////////  CABINET SOUND PARAMETERS  /////////////////////////
Dim NudgeLeftSoundLevel, NudgeRightSoundLevel, NudgeCenterSoundLevel, StartButtonSoundLevel
Dim PlungerReleaseSoundLevel, PlungerPullSoundLevel, CoinSoundLevel

NudgeLeftSoundLevel = 1
NudgeRightSoundLevel = 1
NudgeCenterSoundLevel = 1
StartButtonSoundLevel = 0.1
PlungerReleaseSoundLevel = 1
PlungerPullSoundLevel = 1
CoinSoundLevel = 1


'///////////////////////////  MISC SOUND PARAMETERS  ////////////////////////////
Dim LutToggleSoundLevel :
LutToggleSoundLevel = 0.5


'////////////////////////////////  SOUND HELPERS  ///////////////////////////////
Dim SoundOn : SoundOn = 1
Dim SoundOff : SoundOff = 0
Dim Up : Up = 0
Dim Down : Down = 1
Dim RampUp : RampUp = 1
Dim RampDown : RampDown = 0
Dim RampDownSlow : RampDownSlow = 1
Dim RampDownFast : RampDownFast = 2
Dim CircuitA : CircuitA = 0
Dim CircuitC : CircuitC = 1

'//  Helper for Main (Lower) flippers dampened stroke
Dim BallNearLF : BallNearLF = 0
Dim BallNearRF : BallNearRF = 0

Sub TriggerBallNearLF_Hit()
  'Debug.Print "BallNearLF = 1"
  BallNearLF = 1
End Sub

Sub TriggerBallNearLF_UnHit()
  'Debug.Print "BallNearLF = 0"
  BallNearLF = 0
End Sub

Sub TriggerBallNearRF_Hit()
  'Debug.Print "BallNearRF = 1"
  BallNearRF = 1
End Sub

Sub TriggerBallNearRF_UnHit()
  'Debug.Print "BallNearLF = 0"
  BallNearRF = 0
End Sub


'///////////////////////  SOUND PLAYBACK SUBS / FUNCTIONS  //////////////////////
'//////////////////////  POSITIONAL SOUND PLAYBACK METHODS  /////////////////////

Sub PlaySoundAtLevelStatic(playsoundparams, aVol, tableobj)
    PlaySound playsoundparams, 0, min(1,aVol) * MechVolume, AudioPan(tableobj), 0, 0, 0, 0, AudioFade(tableobj)
End Sub

Sub PlaySoundAtLevelExistingStatic(playsoundparams, aVol, tableobj)
    PlaySound playsoundparams, 0, min(1,aVol) * MechVolume, AudioPan(tableobj), 0, 0, 1, 0, AudioFade(tableobj)
End Sub

Sub PlaySoundAtLevelStaticLoop(playsoundparams, aVol, tableobj)
    PlaySound playsoundparams, -1, min(1,aVol) * MechVolume, AudioPan(tableobj), 0, 0, 0, 0, AudioFade(tableobj)
End Sub

Sub PlaySoundAtLevelExistingStaticLoop(playsoundparams, aVol, tableobj)
    PlaySound playsoundparams, -1, min(1,aVol) * MechVolume, AudioPan(tableobj), 0, 0, 1, 0, AudioFade(tableobj)
End Sub

Sub PlaySoundAtLevelStaticRandomPitch(playsoundparams, aVol, randomPitch, tableobj)
    PlaySound playsoundparams, 0, min(1,aVol) * MechVolume, AudioPan(tableobj), randomPitch, 0, 0, 0, AudioFade(tableobj)
End Sub

Sub PlaySoundAtLevelActiveBall(playsoundparams, aVol)
  PlaySound playsoundparams, 0, min(1,aVol) * MechVolume, AudioPan(ActiveBall), 0, 0, 0, 0, AudioFade(ActiveBall)
End Sub

Sub PlaySoundAtLevelExistingActiveBall(playsoundparams, aVol)
  PlaySound playsoundparams, 0, min(1,aVol) * MechVolume, AudioPan(ActiveBall), 0, 0, 1, 0, AudioFade(ActiveBall)
End Sub

Sub PlaySoundAtLeveTimerActiveBall(playsoundparams, aVol, ballvariable)
  PlaySound playsoundparams, 0, min(1,aVol) * MechVolume, AudioPan(ballvariable), 0, 0, 0, 0, AudioFade(ballvariable)
End Sub

Sub PlaySoundAtLevelTimerExistingActiveBall(playsoundparams, aVol, ballvariable)
  PlaySound playsoundparams, 0, min(1,aVol) * MechVolume, AudioPan(ballvariable), 0, 0, 1, 0, AudioFade(ballvariable)
End Sub

Sub PlaySoundAtLevelRoll(playsoundparams, aVol, pitch)
    PlaySound playsoundparams, -1, min(1,aVol) * MechVolume, AudioPan(tableobj), randomPitch, 0, 0, 0, AudioFade(tableobj)
End Sub


'//////////////////////  SUPPORTING BALL & SOUND FUNCTIONS  /////////////////////

Function AudioFade(tableobj)
  Dim tmp
  Select Case PositionalSoundPlaybackConfiguration
    Case 1
      AudioFade = 0
    Case 2
      AudioFade = 0
    Case 3
      tmp = tableobj.y * 2 / tableheight-1
      If tmp > 0 Then
        AudioFade = Csng(tmp ^5) 'was 10
      Else
        AudioFade = Csng(-((- tmp) ^5) ) 'was 10
      End If
  End Select
End Function

'//  Calculates the pan for a tableobj based on the X position on the table.
Function AudioPan(tableobj)
  Dim tmp
  Select Case PositionalSoundPlaybackConfiguration
    Case 1
      AudioPan = 0
    Case 2
      tmp = tableobj.x * 2 / tablewidth-1
      If tmp > 0 Then
        AudioPan = Csng(tmp ^5) 'was 10
      Else
        AudioPan = Csng(-((- tmp) ^5) ) ' was 10
      End If
    Case 3
      tmp = tableobj.x * 2 / tablewidth-1

      If tmp > 0 Then
        AudioPan = Csng(tmp ^5) ' was 10
      Else
        AudioPan = Csng(-((- tmp) ^5) ) 'was 10
      End If
  End Select
End Function

'//  Calculates the volume of the sound based on the ball speed
Function Vol(ball)
  Vol = Csng(BallVel(ball) ^2)
End Function

'//  Calculates the pitch of the sound based on the ball speed
Function Pitch(ball)
    Pitch = BallVel(ball) * 20
End Function

'//  Calculates the ball speed
Function BallVel(ball)
    BallVel = INT(SQR((ball.VelX ^2) + (ball.VelY ^2) ) )
End Function

'//  Calculates the roll volume of the sound based on the ball speed
Dim TempBallVel
Function VolPlayfieldRoll(ball) ' Calculates the roll volume of the sound based on the ball speed
  TempBallVel = Csng((INT(SQR((ball.VelX^2)+(ball.VelY^2))))/RollingSoundFactor)
  If TempBallVel = 1 Then TempBallVel = 0.999
  If TempBallVel = 0 Then TempBallVel = 0.001
  'debug.print TempBallVel
  TempBallVel = Csng(1/(1+(0.275*(((0.75*TempBallVel)/(1-TempBallVel))^(-2)))))
  VolPlayfieldRoll = TempBallVel
End Function

'//  Calculates the roll volume of the sound based on the ball speed
Function VolSpinningDiscRoll(ball)
  VolSpinningDiscRoll = RollingOnDiscSoundFactor * 0.1 * Csng(BallVel(ball) ^3)
End Function

'//  Calculates the roll volume of the sound based on the ball speed
Dim TempBallVelPlastic
Function VolPlasticMetalRampRoll(ball)
  'VolPlasticMetalRampRoll = RollingOnDiscSoundFactor * 0.0005 * Csng(BallVel(ball) ^3)
  TempBallVelPlastic = Csng((INT(SQR((ball.VelX^2)+(ball.VelY^2))))/RollingSoundFactor)
  If TempBallVelPlastic = 1 Then TempBallVelPlastic = 0.999
  If TempBallVelPlastic = 0 Then TempBallVelPlastic = 0.001
  'debug.print TempBallVel
  TempBallVelPlastic = Csng(1/(1+(0.275*(((0.75*TempBallVelPlastic)/(1-TempBallVelPlastic))^(-2)))))
  VolPlasticMetalRampRoll = TempBallVelPlastic
End Function

'//  Calculates the roll pitch of the sound based on the ball speed
Dim TempPitchBallVel
Function PitchPlayfieldRoll(ball) ' Calculates the roll pitch of the sound based on the ball speed
  'PitchPlayfieldRoll = BallVel(ball) ^2 * 15
  'PitchPlayfieldRoll = Csng(BallVel(ball))/50 * 10000
  'PitchPlayfieldRoll = (1-((Csng(BallVel(ball))/50)^0.2)) * 20000

  'PitchPlayfieldRoll = (2*((Csng(BallVel(ball)))^0.7))/(2+(Csng(BallVel(ball)))) * 16000
  TempPitchBallVel = Csng((INT(SQR((ball.VelX^2)+(ball.VelY^2))))/50)
  If TempPitchBallVel = 1 Then TempPitchBallVel = 0.999
  If TempPitchBallVel = 0 Then TempPitchBallVel = 0.001
  TempPitchBallVel = Csng(1/(1+(0.275*(((0.75*TempPitchBallVel)/(1-TempPitchBallVel))^(-2))))) * 10000
  PitchPlayfieldRoll = TempPitchBallVel
End Function

'//  Calculates the pitch of the sound based on the ball speed.
'//  Used for plastic ramps roll sound
Function PitchPlasticRamp(ball)
    PitchPlasticRamp = BallVel(ball) * 20
End Function

'//  Determines if a point (px,py) in inside a circle with a center of
'//  (cx,cy) coordinates and circleradius
Function InCircle(px,py,cx,cy,circleradius)
  Dim distance
  distance = SQR(((px-cx)^2) + ((py-cy)^2))

  If (distance < circleradius) Then
    InCircle = True
  Else
    InCircle = False
  End If
End Function

'///////////////////////////  PLAY SOUNDS SUBROUTINES  //////////////////////////
'//
'//  These Subroutines implement all mechanical playsounds including timers
'//
'//////////////////////////  GENERAL SOUND SUBROUTINES  /////////////////////////
Sub SoundStartButton()
  PlaySoundAtLevelStatic (Cartridge_Cabinet_Sounds & "_Start_Button"), StartButtonSoundLevel, StartButtonPosition
End Sub

Sub SoundPlungerPull()
  PlaySoundAtLevelStatic (Cartridge_Cabinet_Sounds & "_Plunger_Pull_Slow"), PlungerPullSoundLevel, Plunger
End Sub

Sub SoundPlungerPullStop()
  StopSound Cartridge_Cabinet_Sounds & "_Plunger_Pull_Slow"
End Sub

Sub SoundPlungerReleaseBall()
  PlaySoundAtLevelStatic (Cartridge_Cabinet_Sounds & "_Plunger_Release_Ball_" & Int(Rnd*3)+1), PlungerReleaseSoundLevel, Plunger
End Sub

Sub SoundPlungerReleaseNoBall()
  PlaySoundAtLevelStatic (Cartridge_Cabinet_Sounds & "_Plunger_Release_Empty"), PlungerReleaseSoundLevel, Plunger
End Sub

Sub SoundNudgeLeft()
  PlaySound ("Nudge_" & Int(Rnd*3)+1), 0, NudgeLeftSoundLevel * MechVolume, -0.1, 0.25
End Sub

Sub SoundNudgeRight()
  PlaySound ("Nudge_" & Int(Rnd*3)+1), 0, NudgeRightSoundLevel * MechVolume, 0.1, 0.25
End Sub

Sub SoundNudgeCenter()
  PlaySoundAtLevelStatic ("Nudge_" & Int(Rnd*3)+1), NudgeCenterSoundLevel * MechVolume, sw8
End Sub

'///////////////////////  JP'S VP10 BALL COLLISION SOUND  ///////////////////////
Sub OnBallBallCollision(ball1, ball2, velocity)
  if abs(ball1.vely) < 1 And abs(ball2.vely) < 1 And InRect(ball1.x, ball1.y, 360,730,420,730,420,875,360,875) then
    exit sub  'don't rattle the locked balls
  end if
  PlaySound (Cartridge_BallBallCollision & "_BallBall_Collide_" & Int(Rnd*7)+1), 0, Csng(velocity) ^2 / 200 * BallWithBallCollisionSoundFactor * MechVolume, AudioPan(ball1), 0, Pitch(ball1), 0, 0, AudioFade(ball1)
  FlipperCradleCollision ball1, ball2, velocity
End Sub

'///////////////////////////  OUTLANE ROLLOVER SOUNDS  //////////////////////////
Sub RandomSoundOutlaneRollover()
  PlaySoundAtLevelActiveBall (Cartridge_Rollovers & "_Rollover_Outlane_" & Int(Rnd*4)+1), OutLaneRolloverSoundLevel
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

'//////////////////////////  STADNING TARGET HIT SOUNDS  ////////////////////////
Sub RandomSoundTargetHitStrong()
  Select Case Int(Rnd*4)+1
    Case 1 : PlaySoundAtLevelActiveBall SoundFX(Cartridge_Targets & "_Target_Hit_5",DOFTargets), Vol(ActiveBall) * TargetSoundFactor
    Case 2 : PlaySoundAtLevelActiveBall SoundFX(Cartridge_Targets & "_Target_Hit_6",DOFTargets), Vol(ActiveBall) * TargetSoundFactor
    Case 3 : PlaySoundAtLevelActiveBall SoundFX(Cartridge_Targets & "_Target_Hit_7",DOFTargets), Vol(ActiveBall) * TargetSoundFactor
    Case 4 : PlaySoundAtLevelActiveBall SoundFX(Cartridge_Targets & "_Target_Hit_8",DOFTargets), Vol(ActiveBall) * TargetSoundFactor
  End Select
End Sub

Sub RandomSoundTargetHitWeak()
  Select Case Int(Rnd*4)+1
    Case 1 : PlaySoundAtLevelActiveBall SoundFX(Cartridge_Targets & "_Target_Hit_1",DOFTargets), Vol(ActiveBall) * TargetSoundFactor
    Case 2 : PlaySoundAtLevelActiveBall SoundFX(Cartridge_Targets & "_Target_Hit_2",DOFTargets), Vol(ActiveBall) * TargetSoundFactor
    Case 3 : PlaySoundAtLevelActiveBall SoundFX(Cartridge_Targets & "_Target_Hit_3",DOFTargets), Vol(ActiveBall) * TargetSoundFactor
    Case 4 : PlaySoundAtLevelActiveBall SoundFX(Cartridge_Targets & "_Target_Hit_4",DOFTargets), Vol(ActiveBall) * TargetSoundFactor
  End Select
End Sub

'/////////////////////////////  BALL BOUNCE SOUNDS  /////////////////////////////
Sub RandomSoundBallBouncePlayfieldSoft(aBall)
  Select Case Int(Rnd*7)+1
    Case 1 : PlaySoundAtLevelStatic (Cartridge_Ball_Drop_Bump & "_Ball_Rebound_2"), Vol(aBall) * BallBouncePlayfieldSoftFactor, aBall
    Case 2 : PlaySoundAtLevelStatic (Cartridge_Ball_Drop_Bump & "_Ball_Rebound_12"), Vol(aBall) * BallBouncePlayfieldSoftFactor, aBall
    Case 3 : PlaySoundAtLevelStatic (Cartridge_Ball_Drop_Bump & "_Ball_Rebound_14"), Vol(aBall) * BallBouncePlayfieldSoftFactor, aBall
    Case 4 : PlaySoundAtLevelStatic (Cartridge_Ball_Drop_Bump & "_Ball_Rebound_18"), Vol(aBall) * BallBouncePlayfieldSoftFactor, aBall
    Case 5 : PlaySoundAtLevelStatic (Cartridge_Ball_Drop_Bump & "_Ball_Rebound_19"), Vol(aBall) * BallBouncePlayfieldSoftFactor, aBall
    Case 6 : PlaySoundAtLevelStatic (Cartridge_Ball_Drop_Bump & "_Ball_Rebound_20"), Vol(aBall) * BallBouncePlayfieldSoftFactor, aBall
    Case 7 : PlaySoundAtLevelStatic (Cartridge_Ball_Drop_Bump & "_Ball_Rebound_21"), Vol(aBall) * BallBouncePlayfieldSoftFactor, aBall
  End Select
End Sub

Sub RandomSoundBallBouncePlayfieldHard(aBall)
  Select Case Int(Rnd*12)+1
    Case 1 : PlaySoundAtLevelStatic (Cartridge_Ball_Drop_Bump & "_Ball_Rebound_1"), Vol(aBall) * BallBouncePlayfieldHardFactor, aBall
    Case 2 : PlaySoundAtLevelStatic (Cartridge_Ball_Drop_Bump & "_Ball_Rebound_3"), Vol(aBall) * BallBouncePlayfieldHardFactor, aBall
    Case 3 : PlaySoundAtLevelStatic (Cartridge_Ball_Drop_Bump & "_Ball_Rebound_7"), Vol(aBall) * BallBouncePlayfieldHardFactor, aBall
    Case 4 : PlaySoundAtLevelStatic (Cartridge_Ball_Drop_Bump & "_Ball_Rebound_8"), Vol(aBall) * BallBouncePlayfieldHardFactor, aBall
    Case 5 : PlaySoundAtLevelStatic (Cartridge_Ball_Drop_Bump & "_Ball_Rebound_9"), Vol(aBall) * BallBouncePlayfieldHardFactor, aBall
    Case 6 : PlaySoundAtLevelStatic (Cartridge_Ball_Drop_Bump & "_Ball_Rebound_11"), Vol(aBall) * BallBouncePlayfieldHardFactor, aBall
    Case 7 : PlaySoundAtLevelStatic (Cartridge_Ball_Drop_Bump & "_Ball_Rebound_13"), Vol(aBall) * BallBouncePlayfieldHardFactor, aBall
    Case 8 : PlaySoundAtLevelStatic (Cartridge_Ball_Drop_Bump & "_Ball_Rebound_15"), Vol(aBall) * BallBouncePlayfieldHardFactor, aBall
    Case 9 : PlaySoundAtLevelStatic (Cartridge_Ball_Drop_Bump & "_Ball_Rebound_16"), Vol(aBall) * BallBouncePlayfieldHardFactor, aBall
    Case 10 : PlaySoundAtLevelStatic (Cartridge_Ball_Drop_Bump & "_Ball_Rebound_17"), Vol(aBall) * BallBouncePlayfieldHardFactor, aBall
    Case 11 : PlaySoundAtLevelStatic (Cartridge_Ball_Drop_Bump & "_Ball_Rebound_22"), Vol(aBall) * BallBouncePlayfieldHardFactor, aBall
    Case 12 : PlaySoundAtLevelStatic (Cartridge_Ball_Drop_Bump & "_Ball_Rebound_23"), Vol(aBall) * BallBouncePlayfieldHardFactor, aBall
  End Select
End Sub

'////////////////////////////////////  DRAIN  ///////////////////////////////////
'///////////////////////////////  OUTHOLE SOUNDS  ///////////////////////////////
Sub RandomSoundOutholeHit(sw)
  PlaySoundAtLevelStatic (Cartridge_Trough & "_Outhole_Drain_Hit_" & Int(Rnd*4)+1), Solenoid_OutholeKicker_SoundLevel, sw
End Sub

Sub RandomSoundOutholeKicker()
  PlaySoundAtLevelStatic SoundFX(Cartridge_Trough & "_Outhole_Kicker_" & Int(Rnd*4)+1,DOFContactors), Solenoid_OutholeKicker_SoundLevel, sw8
End Sub

'/////////////////////  BALL SHOOTER FEEDER SOLENOID SOUNDS  ////////////////////
Sub RandomSoundShooterFeeder()
  PlaySoundAtLevelStatic SoundFX(Cartridge_Trough & "_Shooter_Feeder_" & Int(Rnd*6)+1,DOFContactors), Solenoid_ShooterFeeder_SoundLevel, sw48
End Sub

'///////  SHOOTER LANE - BALL RELEASE ROLL IN SHOOTER LANE SOUND - SOUND  ///////
Sub SoundBallReleaseShooterLane(toggle)
  Select Case toggle
    Case SoundOn
      PlaySoundAtLevelActiveBall (Cartridge_Table_Specifics & "_Ball_Launch_from_Shooter_Lane"), BallReleaseShooterLaneSoundLevel
    Case SoundOff
      StopSound Cartridge_Table_Specifics & "_Ball_Launch_from_Shooter_Lane"
  End Select
End Sub

'//////////////////////////////  KNOCKER SOLENOID  //////////////////////////////
Sub KnockerSolenoid(enabled)
  'PlaySoundAtLevelStatic SoundFX(Cartridge_Knocker & "_Knocker_Coil",DOFKnocker), Solenoid_Knocker_SoundLevel, KnockerPosition
  if Enabled then
    PlaySound SoundFX(Cartridge_Knocker & "_Knocker_Coil",DOFKnocker), 0, Solenoid_Knocker_SoundLevel
  end if
End Sub

'/////////////////////////////  EJECT HOLD SOLENOID  ////////////////////////////
Sub RandomSoundEjectHoleSolenoidLeft()
  PlaySoundAtLevelStatic (Cartridge_Kickers & "_Eject_Kickout_" & Int(Rnd*8)+1), Solenoid_TopEject_SoundLevel, LeftKicker
End Sub

Sub RandomSoundEjectHoleSolenoidRight()
  PlaySoundAtLevelStatic (Cartridge_Kickers & "_Eject_Kickout_" & Int(Rnd*8)+1), Solenoid_TopEject_SoundLevel, RightKicker
End Sub

'//////////////////////////  SLINGSHOT SOLENOID SOUNDS  /////////////////////////
Sub RandomSoundSlingshotLeft()
  PlaySoundAtLevelStatic SoundFX(Cartridge_Slingshots & "_Slingshot_Left_" & Int(Rnd*26)+1,DOFContactors), Solenoid_Slingshot_SoundLevel, LeftSlingshotPosition
End Sub

Sub RandomSoundSlingshotRight()
  PlaySoundAtLevelStatic SoundFX(Cartridge_Slingshots & "_Slingshot_Right_" & Int(Rnd*25)+1,DOFContactors), Solenoid_Slingshot_SoundLevel, RightSlingshotPosition
End Sub

'///////////////////////////  BUMPER SOLENOID SOUNDS  ///////////////////////////
'////////////////////////////////  BUMPERS - TOP  ///////////////////////////////
Sub RandomSoundBumperLeft(Bump)
' Debug.Print Vol(ActiveBall) * Solenoid_Bumper_SoundMultiplier
  PlaySoundAtLevelStatic SoundFX(Cartridge_Bumpers & "_Jet_Bumper_Left_" & Int(Rnd*22)+1,DOFContactors), Vol(ActiveBall) * Solenoid_Bumper_SoundMultiplier, Bump
End Sub

Sub RandomSoundBumperUp(Bump)
  'Debug.Print Vol(ActiveBall) * Solenoid_Bumper_SoundMultiplier
  PlaySoundAtLevelStatic SoundFX(Cartridge_Bumpers & "_Jet_Bumper_Up_" & Int(Rnd*25)+1,DOFContactors), Vol(ActiveBall) * Solenoid_Bumper_SoundMultiplier, Bump
End Sub

Sub RandomSoundBumperLow(Bump)
  'Debug.Print Vol(ActiveBall) * Solenoid_Bumper_SoundMultiplier
  PlaySoundAtLevelStatic SoundFX(Cartridge_Bumpers & "_Jet_Bumper_Low_" & Int(Rnd*28)+1,DOFContactors), Vol(ActiveBall) * Solenoid_Bumper_SoundMultiplier, Bump
End Sub



'///////////////////////  FLIPPER BATS SOUND SUBROUTINES  ///////////////////////
'//////////////////////  FLIPPER BATS SOLENOID CORE SOUND  //////////////////////
Sub RandomSoundFlipperLowerLeftUpFullStroke(flipper)
  PlaySoundAtLevelStatic SoundFX(Cartridge_Flippers & "_Flipper_Lower_Left_Up_Full_Stroke_" & RndInt(1,10),DOFFlippers), FlipperLeftLowerHitParm, Flipper
End Sub

Sub RandomSoundFlipperLowerLeftUpDampenedStroke(flipper)
  PlaySoundAtLevelStatic SoundFX(Cartridge_Flippers & "_Flipper_Lower_Left_Up_Dampened_Stroke_" & RndInt(1,23),DOFFlippers), FlipperLeftLowerHitParm * 1.2, Flipper
End Sub

Sub RandomSoundFlipperLowerRightUpFullStroke(flipper)
  PlaySoundAtLevelStatic SoundFX(Cartridge_Flippers & "_Flipper_Lower_Right_Up_Full_Stroke_" & RndInt(1,11),DOFFlippers), FlipperRightLowerHitParm, Flipper
End Sub

Sub RandomSoundFlipperLowerRightUpDampenedStroke(flipper)
  PlaySoundAtLevelStatic SoundFX(Cartridge_Flippers & "_Flipper_Lower_Right_Up_Dampened_Stroke_" & RndInt(1,23),DOFFlippers), FlipperLeftLowerHitParm * 1.2, Flipper
End Sub

Sub RandomSoundFlipperLowerLeftReflip(flipper)
  PlaySoundAtLevelStatic SoundFX(Cartridge_Flippers & "_Flipper_Lower_Left_Reflip_" & RndInt(1,3),DOFFlippers), (RndNum(0.8, 1))*FlipperUpSoundLevel, Flipper
End Sub

Sub RandomSoundFlipperLowerRightReflip(flipper)
  PlaySoundAtLevelStatic SoundFX(Cartridge_Flippers & "_Flipper_Lower_Right_Reflip_" & RndInt(1,3),DOFFlippers), (RndNum(0.8, 1))*FlipperUpSoundLevel, Flipper
End Sub

Sub RandomSoundFlipperLowerLeftDown(flipper)
  PlaySoundAtLevelStatic SoundFX(Cartridge_Flippers & "_Flipper_Lower_Left_Down_" & RndInt(1,10),DOFFlippers), FlipperDownSoundLevel, Flipper
End Sub

Sub RandomSoundFlipperLowerRightDown(flipper)
  PlaySoundAtLevelStatic SoundFX(Cartridge_Flippers & "_Flipper_Lower_Right_Down_" & RndInt(1,11),DOFFlippers), FlipperDownSoundLevel, Flipper
End Sub

Sub RandomSoundLowerLeftQuickFlipUp()
  PlaySoundAtLevelStatic SoundFX(Cartridge_Flippers & "_Flipper_Lower_Left_QuickFlip_Up_" & RndInt(1,3),DOFFlippers), (RndNum(0.8, 1))*FlipperUpSoundLevel, LeftFlipper
End Sub

Sub RandomSoundLowerRightQuickFlipUp()
  PlaySoundAtLevelStatic SoundFX(Cartridge_Flippers & "_Flipper_Lower_Right_QuickFlip_Up_" & RndInt(1,3),DOFFlippers), (RndNum(0.8, 1))*FlipperUpSoundLevel, RightFlipper
End Sub

Sub StopAnyFlipperLowerLeftUp()
  Dim anyFullStrokeSound
  Dim anyDampenedStrokeSound
  For anyFullStrokeSound = 1 to 10
    StopSound(Cartridge_Flippers & "_Flipper_Lower_Left_Up_Full_Stroke_" & anyFullStrokeSound)
  Next
  For anyDampenedStrokeSound = 1 to 23
    StopSound(Cartridge_Flippers & "_Flipper_Lower_Left_Up_Dampened_Stroke_" & anyDampenedStrokeSound)
  Next
End Sub

Sub StopAnyFlipperLowerRightUp()
  Dim anyFullStrokeSound
  Dim anyDampenedStrokeSound
  For anyFullStrokeSound = 1 to 11
    StopSound(Cartridge_Flippers & "_Flipper_Lower_Right_Up_Full_Stroke_" & anyFullStrokeSound)
  Next
  For anyDampenedStrokeSound = 1 to 23
    StopSound(Cartridge_Flippers & "_Flipper_Lower_Right_Up_Dampened_Stroke_" & anyDampenedStrokeSound)
  Next
End Sub

Sub StopAnyFlipperLowerLeftDown()
  Dim anyFullDownSound
  For anyFullDownSound = 1 to 10
    StopSound(Cartridge_Flippers & "_Flipper_Lower_Left_Down_" & anyFullDownSound)
  Next
End Sub

Sub StopAnyFlipperLowerRightDown()
  Dim anyFullDownSound
  For anyFullDownSound = 1 to 10
    StopSound(Cartridge_Flippers & "_Flipper_Lower_Right_Down_" & anyFullDownSound)
  Next
End Sub

Sub FlipperHoldCoilLeft(toggle, flipper)
  Select Case toggle
    Case SoundOn
      PlaySoundAtLevelExistingStaticLoop (Cartridge_Flippers & "_Flipper_Hold_Coil_Low_Frequencies_Loop_Lower_Left"), FlipperHoldSoundLevel, flipper
      PlaySoundAtLevelExistingStaticLoop (Cartridge_Flippers & "_Flipper_Hold_Coil_Med_Frequencies_Loop_Lower_Left"), FlipperHoldSoundLevel, flipper
    Case SoundOff
      StopSound Cartridge_Flippers & "_Flipper_Hold_Coil_Low_Frequencies_Loop_Lower_Left"
      StopSound Cartridge_Flippers & "_Flipper_Hold_Coil_Med_Frequencies_Loop_Lower_Left"
  End Select
End Sub

Sub FlipperHoldCoilRight(toggle, flipper)
  Select Case toggle
    Case SoundOn
      PlaySoundAtLevelExistingStaticLoop (Cartridge_Flippers & "_Flipper_Hold_Coil_Low_Frequencies_Loop_Lower_Right"), FlipperHoldSoundLevel, flipper
      PlaySoundAtLevelExistingStaticLoop (Cartridge_Flippers & "_Flipper_Hold_Coil_Med_Frequencies_Loop_Lower_Right"), FlipperHoldSoundLevel, flipper
    Case SoundOff
      StopSound Cartridge_Flippers & "_Flipper_Hold_Coil_Low_Frequencies_Loop_Lower_Right"
      StopSound Cartridge_Flippers & "_Flipper_Hold_Coil_Med_Frequencies_Loop_Lower_Right"
  End Select
End Sub

'///////////////////////  FLIPPER BATS BALL COLLIDE SOUND  //////////////////////
dim angdamp, veldamp
angdamp = 0.2
veldamp = 0.8

Sub LeftFlipper_Collide(parm)
  CheckLiveCatch ActiveBall, LeftFlipper, LFCount, parm
  LF.ReProcessBalls ActiveBall

  If parm => 22 Then
    ' Strong hit safe values boundary
    ' Flipper stroke dampened
    FlipperLeftLowerHitParm = FlipperUpSoundLevel * 0.1
  Else
    If parm =< 1 Then
      ' Weak hit safe values boundary
      ' Flipper stroke full
      FlipperLeftLowerHitParm = FlipperUpSoundLevel
    Else
      ' Fully modulated hit
      FlipperLeftLowerHitParm = FlipperUpSoundLevel * (1-(parm/25))
    End If
  End If

  RandomSoundRubberFlipper(parm)
End Sub

Sub RightFlipper_Collide(parm)
  CheckLiveCatch Activeball, RightFlipper, RFCount, parm
  RF.ReProcessBalls ActiveBall
  If parm => 22 Then
    ' Strong hit safe values boundary
    ' Flipper stroke dampened
    FlipperRightLowerHitParm = FlipperUpSoundLevel * 0.1
  Else
    If parm =< 1 Then
      ' Weak hit safe values boundary
      ' Flipper stroke full
      FlipperRightLowerHitParm = FlipperUpSoundLevel
    Else
      ' Fully modulated hit
      FlipperRightLowerHitParm = FlipperUpSoundLevel * (1-(parm/25))
    End If
  End If

  RandomSoundRubberFlipper(parm)
End Sub


Sub RandomSoundRubberFlipper(parm)
  PlaySoundAtLevelActiveBall (Cartridge_Rubber_Hits & "_Rubber_Flipper_Hit_" & Int(Rnd*7)+1), parm / 25 * RubberFlipperSoundFactor
End Sub

'//////////////////////////  SOLENOID A/C SELECT RELAY  /////////////////////////
Sub Sound_Solenoid_AC(toggle)
  Select Case toggle
    Case CircuitA
      If RelaysPosition = 1 Then PlaySoundAtLevelStatic (Cartridge_Relays & "_Relays_AC_Select_Relay_Side_A"), RelaySolenoidACSelectSoundLevel, GIUpperPosition
      If RelaysPosition = 2 Then PlaySoundAtLevelStatic (Cartridge_Relays & "_Relays_AC_Select_Relay_Side_A"), RelaySolenoidACSelectSoundLevel, ACSelectPosition
    Case CircuitC
      If RelaysPosition = 1 Then PlaySoundAtLevelStatic (Cartridge_Relays & "_Relays_AC_Select_Relay_Side_C"), RelaySolenoidACSelectSoundLevel, GIUpperPosition
      If RelaysPosition = 2 Then PlaySoundAtLevelStatic (Cartridge_Relays & "_Relays_AC_Select_Relay_Side_C"), RelaySolenoidACSelectSoundLevel, ACSelectPosition
  End Select
End Sub

'//////////////////////////  GENERAL ILLUMINATION RELAYS  ///////////////////////
Sub Sound_LowerGI_Relay(toggle)
  Select Case toggle
    Case SoundOn
      If RelaysPosition = 1 Then PlaySoundAtLevelStatic (Cartridge_Relays & "_Relays_Lower_Playfield_and_Backbox_GI_Relay_On"), RelayLowerGISoundLevel, GIUpperPosition
      If RelaysPosition = 2 Then PlaySoundAtLevelStatic (Cartridge_Relays & "_Relays_Lower_Playfield_and_Backbox_GI_Relay_On"), RelayLowerGISoundLevel, GIUpperPosition
    Case SoundOff
      If RelaysPosition = 1 Then PlaySoundAtLevelStatic (Cartridge_Relays & "_Relays_Lower_Playfield_and_Backbox_GI_Relay_Off"), RelayLowerGISoundLevel, GIUpperPosition
      If RelaysPosition = 2 Then PlaySoundAtLevelStatic (Cartridge_Relays & "_Relays_Lower_Playfield_and_Backbox_GI_Relay_Off"), RelayLowerGISoundLevel, GIUpperPosition
  End Select
End Sub

Sub Sound_UpperGI_Relay(toggle)
  Select Case toggle
    Case SoundOn
      If RelaysPosition = 1 Then PlaySoundAtLevelStatic (Cartridge_Relays & "_Relays_Upper_Playfield_GI_Relay_On"), RelayUpperGISoundLevel, GIUpperPosition
      If RelaysPosition = 2 Then PlaySoundAtLevelStatic (Cartridge_Relays & "_Relays_Upper_Playfield_GI_Relay_On"), RelayUpperGISoundLevel, GILowerPosition
    Case SoundOff
      If RelaysPosition = 1 Then PlaySoundAtLevelStatic (Cartridge_Relays & "_Relays_Upper_Playfield_GI_Relay_Off"), RelayUpperGISoundLevel, GIUpperPosition
      If RelaysPosition = 2 Then PlaySoundAtLevelStatic (Cartridge_Relays & "_Relays_Upper_Playfield_GI_Relay_Off"), RelayUpperGISoundLevel, GILowerPosition
  End Select
End Sub

'///////////////////////////////  FLASHERS RELAY  ///////////////////////////////
Sub Sound_Flasher_Relay(toggle, tableobj)
  Select Case toggle
    Case SoundOn
      If RelaysPosition = 1 Then PlaySoundAtLevelStatic (Cartridge_Relays & "_Relays_Flashers_Relay_On"), RelayFlasherSoundLevel, GIUpperPosition
      If RelaysPosition = 2 Then PlaySoundAtLevelStatic (Cartridge_Relays & "_Relays_Flashers_Relay_On"), RelayFlasherSoundLevel, tableobj
    Case SoundOff
      If RelaysPosition = 1 Then PlaySoundAtLevelStatic (Cartridge_Relays & "_Relays_Flashers_Relay_Off"), RelayFlasherSoundLevel, GIUpperPosition
      If RelaysPosition = 2 Then PlaySoundAtLevelStatic (Cartridge_Relays & "_Relays_Flashers_Relay_Off"), RelayFlasherSoundLevel, tableobj
  End Select
End Sub

'////////////////////  VARIOUS PLAYFIELD SOUND SUBROUTINES  /////////////////////
'/////////////////////////////  RUBBERS AND POSTS  //////////////////////////////
'/////////////////////////////  RUBBERS - EVENTS  ///////////////////////////////
Sub Rubbers_Hit(idx)
  dim finalspeed
    finalspeed=SQR(activeball.velx * activeball.velx + activeball.vely * activeball.vely)
  If finalspeed > 10 then
    RandomSoundRubberStrong()
  End if
  If finalspeed <= 10 then
    RandomSoundRubberWeak()
  End If
End Sub

'/////////////////////  RUBBERS AND POSTS - STRONG IMPACTS  /////////////////////
Sub RandomSoundRubberStrong()
  Select Case Int(Rnd*10)+1
    Case 1 : PlaySoundAtLevelActiveBall (Cartridge_Rubber_Hits & "_Rubber_Hit_Strong_1"), Vol(ActiveBall) * RubberStrongSoundFactor
    Case 2 : PlaySoundAtLevelActiveBall (Cartridge_Rubber_Hits & "_Rubber_Hit_Strong_2"), Vol(ActiveBall) * RubberStrongSoundFactor
    Case 3 : PlaySoundAtLevelActiveBall (Cartridge_Rubber_Hits & "_Rubber_Hit_Strong_3"), Vol(ActiveBall) * RubberStrongSoundFactor
    Case 4 : PlaySoundAtLevelActiveBall (Cartridge_Rubber_Hits & "_Rubber_Hit_Strong_4"), Vol(ActiveBall) * RubberStrongSoundFactor
    Case 5 : PlaySoundAtLevelActiveBall (Cartridge_Rubber_Hits & "_Rubber_Hit_Strong_5"), Vol(ActiveBall) * RubberStrongSoundFactor
    Case 6 : PlaySoundAtLevelActiveBall (Cartridge_Rubber_Hits & "_Rubber_Hit_Strong_6"), Vol(ActiveBall) * RubberStrongSoundFactor
    Case 7 : PlaySoundAtLevelActiveBall (Cartridge_Rubber_Hits & "_Rubber_Hit_Strong_7"), Vol(ActiveBall) * RubberStrongSoundFactor
    Case 8 : PlaySoundAtLevelActiveBall (Cartridge_Rubber_Hits & "_Rubber_Hit_Strong_8"), Vol(ActiveBall) * RubberStrongSoundFactor
    Case 9 : PlaySoundAtLevelActiveBall (Cartridge_Rubber_Hits & "_Rubber_Hit_Strong_9"), Vol(ActiveBall) * RubberStrongSoundFactor
    Case 10 : PlaySoundAtLevelActiveBall (Cartridge_Rubber_Hits & "_Rubber_Hit_Strong_10"), Vol(ActiveBall) * RubberStrongSoundFactor * 0.6
  End Select
End Sub

'///////////////////////  RUBBERS AND POSTS - WEAK IMPACTS  /////////////////////
Sub RandomSoundRubberWeak()
  PlaySoundAtLevelActiveBall (Cartridge_Rubber_Hits & "_Rubber_Hit_" & Int(Rnd*8)+1), Vol(ActiveBall) * RubberWeakSoundFactor
End Sub

'///////////////////////////////  WALL IMPACTS  /////////////////////////////////
Sub RandomSoundWall()
  dim finalspeed
    finalspeed=SQR(activeball.velx * activeball.velx + activeball.vely * activeball.vely)
  If finalspeed > 16 then
    Select Case Int(Rnd*5)+1
      Case 1 : PlaySoundAtLevelExistingActiveBall (Cartridge_Metal_Hits & "_Metal_Alternative_Hit_1"), Vol(ActiveBall) * WallImpactSoundFactor
      Case 2 : PlaySoundAtLevelExistingActiveBall (Cartridge_Metal_Hits & "_Metal_Alternative_Hit_2"), Vol(ActiveBall) * WallImpactSoundFactor
      Case 3 : PlaySoundAtLevelExistingActiveBall (Cartridge_Metal_Hits & "_Metal_Alternative_Hit_5"), Vol(ActiveBall) * WallImpactSoundFactor * 0.05
      Case 4 : PlaySoundAtLevelExistingActiveBall (Cartridge_Metal_Hits & "_Metal_Alternative_Hit_7"), Vol(ActiveBall) * WallImpactSoundFactor
      Case 5 : PlaySoundAtLevelExistingActiveBall (Cartridge_Metal_Hits & "_Metal_Alternative_Hit_9"), Vol(ActiveBall) * WallImpactSoundFactor
    End Select
  End if
  If finalspeed >= 6 AND finalspeed <= 16 then
    Select Case Int(Rnd*4)+1
      Case 1 : PlaySoundAtLevelExistingActiveBall (Cartridge_Metal_Hits & "_Metal_Alternative_Hit_3"), Vol(ActiveBall) * WallImpactSoundFactor
      Case 2 : PlaySoundAtLevelExistingActiveBall (Cartridge_Metal_Hits & "_Metal_Alternative_Hit_4"), Vol(ActiveBall) * WallImpactSoundFactor
      Case 3 : PlaySoundAtLevelExistingActiveBall (Cartridge_Metal_Hits & "_Metal_Alternative_Hit_6"), Vol(ActiveBall) * WallImpactSoundFactor
      Case 4 : PlaySoundAtLevelExistingActiveBall (Cartridge_Metal_Hits & "_Metal_Alternative_Hit_8"), Vol(ActiveBall) * WallImpactSoundFactor
    End Select
  End If
  If finalspeed < 6 Then
    Select Case Int(Rnd*3)+1
      Case 1 : PlaySoundAtLevelExistingActiveBall (Cartridge_Metal_Hits & "_Metal_Alternative_Hit_4"), Vol(ActiveBall) * WallImpactSoundFactor
      Case 2 : PlaySoundAtLevelExistingActiveBall (Cartridge_Metal_Hits & "_Metal_Alternative_Hit_6"), Vol(ActiveBall) * WallImpactSoundFactor
      Case 3 : PlaySoundAtLevelExistingActiveBall (Cartridge_Metal_Hits & "_Metal_Alternative_Hit_8"), Vol(ActiveBall) * WallImpactSoundFactor
    End Select
  End if
End Sub

'/////////////////////////////  WALL IMPACTS EVENTS  ////////////////////////////

Sub Wall46_Hit()
  RandomSoundMetal()
End Sub

sub HitsMetal_Hit(IDX)
' debug.print "metal hit"
  RandomSoundMetal()
end sub

'RandomSoundBottomArchBallGuideSoftHit - Soft Bounces
Sub Wall78_Hit() : RandomSoundBottomArchBallGuideSoftHit() : End Sub
Sub Wall26_Hit() : RandomSoundBottomArchBallGuideSoftHit() : End Sub

'RandomSoundBottomArchBallGuideHardHit - Hard Hit
Sub Wall062_Hit() : RandomSoundBottomArchBallGuideHardHit() : End Sub
Sub Wall063_Hit() : RandomSoundBottomArchBallGuideHardHit() : End Sub

'RandomSoundFlipperBallGuide
Sub Wall49_Hit() : RandomSoundFlipperBallGuide() : End Sub
Sub Wall48_Hit() : RandomSoundFlipperBallGuide() : End Sub

'Outlane - Walls & Primitives
Sub Wall70_Hit() : RandomSoundOutlaneWalls() : End Sub
Sub Wall5_Hit() : RandomSoundOutlaneWalls() : End Sub
Sub Wall73_Hit() : RandomSoundOutlaneWalls() : End Sub
Sub Wall24_Hit() : RandomSoundOutlaneWalls() : End Sub

'////////////////////////////  INNER LEFT LANE WALLS  ///////////////////////////
Sub Wall226_Hit()
  PlaySoundAtLevelActiveBall (Cartridge_Metal_Hits & "_InnerLane_Left_Wall_" & Int(Rnd*3)+1), Vol(ActiveBall) * MetalImpactSoundFactor
End Sub

Sub Wall225_Hit()
  PlaySoundAtLevelActiveBall (Cartridge_Metal_Hits & "_InnerLane_Left_Wall_" & Int(Rnd*3)+1), Vol(ActiveBall) * MetalImpactSoundFactor
End Sub


'right arch
Sub Wall141_Hit()
  PlaySoundAtLevelActiveBall (Cartridge_Metal_Hits & "_InnerLane_Left_Wall_" & Int(Rnd*3)+1), Vol(ActiveBall) * MetalImpactSoundFactor
End Sub

Sub Wall132_Hit()
  PlaySoundAtLevelActiveBall (Cartridge_Metal_Hits & "_InnerLane_Left_Wall_" & Int(Rnd*3)+1), Vol(ActiveBall) * MetalImpactSoundFactor
End Sub

Sub Wall224_Hit()
  PlaySoundAtLevelActiveBall (Cartridge_Metal_Hits & "_InnerLane_Left_Wall_" & Int(Rnd*3)+1), Vol(ActiveBall) * MetalImpactSoundFactor
End Sub


'/////////////////////////////  METAL TOUCH SOUNDS  /////////////////////////////
Sub RandomSoundMetal()
  Select Case Int(Rnd*20)+1
    Case 1 : PlaySoundAtLevelActiveBall (Cartridge_Metal_Hits & "_Metal_Hit_1"), Vol(ActiveBall) * MetalImpactSoundFactor
    Case 2 : PlaySoundAtLevelActiveBall (Cartridge_Metal_Hits & "_Metal_Hit_2"), Vol(ActiveBall) * MetalImpactSoundFactor
    Case 3 : PlaySoundAtLevelActiveBall (Cartridge_Metal_Hits & "_Metal_Hit_3"), Vol(ActiveBall) * MetalImpactSoundFactor
    Case 4 : PlaySoundAtLevelActiveBall (Cartridge_Metal_Hits & "_Metal_Hit_4"), Vol(ActiveBall) * MetalImpactSoundFactor
    Case 5 : PlaySoundAtLevelActiveBall (Cartridge_Metal_Hits & "_Metal_Hit_5"), Vol(ActiveBall) * 0.02 * MetalImpactSoundFactor
    Case 6 : PlaySoundAtLevelActiveBall (Cartridge_Metal_Hits & "_Metal_Hit_6"), Vol(ActiveBall) * MetalImpactSoundFactor
    Case 7 : PlaySoundAtLevelActiveBall (Cartridge_Metal_Hits & "_Metal_Hit_7"), Vol(ActiveBall) * MetalImpactSoundFactor
    Case 8 : PlaySoundAtLevelActiveBall (Cartridge_Metal_Hits & "_Metal_Hit_8"), Vol(ActiveBall) * MetalImpactSoundFactor
    Case 9 : PlaySoundAtLevelActiveBall (Cartridge_Metal_Hits & "_Metal_Hit_9"), Vol(ActiveBall) * MetalImpactSoundFactor
    Case 10 : PlaySoundAtLevelActiveBall (Cartridge_Metal_Hits & "_Metal_Hit_10"), Vol(ActiveBall) * MetalImpactSoundFactor
    Case 11 : PlaySoundAtLevelActiveBall (Cartridge_Metal_Hits & "_Metal_Hit_11"), Vol(ActiveBall) * MetalImpactSoundFactor
    Case 12 : PlaySoundAtLevelActiveBall (Cartridge_Metal_Hits & "_Metal_Hit_12"), Vol(ActiveBall) * MetalImpactSoundFactor
    Case 13 : PlaySoundAtLevelActiveBall (Cartridge_Metal_Hits & "_Metal_Hit_13"), Vol(ActiveBall) * MetalImpactSoundFactor
    Case 14 : PlaySoundAtLevelActiveBall (Cartridge_Metal_Hits & "_Metal_Hit_14"), Vol(ActiveBall) * MetalImpactSoundFactor
    Case 15 : PlaySoundAtLevelActiveBall (Cartridge_Metal_Hits & "_Metal_Hit_15"), Vol(ActiveBall) * MetalImpactSoundFactor
    Case 16 : PlaySoundAtLevelActiveBall (Cartridge_Metal_Hits & "_Metal_Hit_16"), Vol(ActiveBall) * MetalImpactSoundFactor
    Case 17 : PlaySoundAtLevelActiveBall (Cartridge_Metal_Hits & "_Metal_Hit_17"), Vol(ActiveBall) * MetalImpactSoundFactor
    Case 18 : PlaySoundAtLevelActiveBall (Cartridge_Metal_Hits & "_Metal_Hit_18"), Vol(ActiveBall) * MetalImpactSoundFactor
    Case 19 : PlaySoundAtLevelActiveBall (Cartridge_Metal_Hits & "_Metal_Hit_19"), Vol(ActiveBall) * MetalImpactSoundFactor
    Case 20 : PlaySoundAtLevelActiveBall (Cartridge_Metal_Hits & "_Metal_Hit_20"), Vol(ActiveBall) * MetalImpactSoundFactor
  End Select
End Sub

'/////////////////////////////  WALL IMPACTS  ////////////////////////////
Sub HitsWoods_Hit(idx)
' debug.print "wood hit"
  RandomSoundWood()
End Sub

Sub RandomSoundWood()
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

'///////////////////////////////  OUTLANES WALLS  ///////////////////////////////
Sub RandomSoundOutlaneWalls()
  PlaySoundAtLevelActiveBall (Cartridge_Metal_Hits & "_Outlane_Wall_" & Int(Rnd*9)+1), OutlaneWallsSoundFactor
End Sub

'///////////////////////////  BOTTOM ARCH BALL GUIDE  ///////////////////////////
'///////////////////  BOTTOM ARCH BALL GUIDE - SOFT BOUNCES  ////////////////////
Sub RandomSoundBottomArchBallGuideSoftHit()
  PlaySoundAtLevelActiveBall (Cartridge_Metal_Hits & "_Arch_Ball_Guide_Hit_Soft_" & Int(Rnd*4)+1), BottomArchBallGuideSoundFactor
End Sub


'//////////////////////  BOTTOM ARCH BALL GUIDE - HARD HITS  ////////////////////
Sub RandomSoundBottomArchBallGuideHardHit()
  PlaySoundAtLevelActiveBall (Cartridge_Metal_Hits & "_Arch_Ball_Guide_Hit_Hard_" & Int(Rnd*3)+1), BottomArchBallGuideSoundFactor * 3
End Sub

'//////////////////////////////  FLIPPER BALL GUIDE  ////////////////////////////
Sub RandomSoundFlipperBallGuide()
  dim finalspeed
    finalspeed=SQR(activeball.velx * activeball.velx + activeball.vely * activeball.vely)
  If finalspeed > 16 then
    Select Case Int(Rnd*2)+1
      Case 1 : PlaySoundAtLevelActiveBall (Cartridge_Apron & "_Apron_Hit_Hard_1"),  Vol(ActiveBall) * FlipperBallGuideSoundFactor
      Case 2 : PlaySoundAtLevelActiveBall (Cartridge_Apron & "_Apron_Hit_Hard_2"),  Vol(ActiveBall) * 0.8 * FlipperBallGuideSoundFactor
    End Select
  End if
  If finalspeed >= 6 AND finalspeed <= 16 then
    PlaySoundAtLevelActiveBall (Cartridge_Apron & "_Apron_Hit_Medium_" & Int(Rnd*3)+1),  Vol(ActiveBall) * FlipperBallGuideSoundFactor
  End If
  If finalspeed < 6 Then
    PlaySoundAtLevelActiveBall (Cartridge_Apron & "_Apron_Hit_Soft_" & Int(Rnd*7)+1),  Vol(ActiveBall) * FlipperBallGuideSoundFactor
  End if
End Sub

'/////////////////////////  WIREFORM ANTI-REBOUNT RAILS  ////////////////////////
Sub RandomSoundWireformAntiRebountRail()
  dim finalspeed
    finalspeed=SQR(activeball.velx * activeball.velx + activeball.vely * activeball.vely)
  If finalspeed >= 10 then
    Select Case Int(Rnd*5)+1
      Case 1 : PlaySoundAtLevelActiveBall (Cartridge_Metal_Hits & "_Wireform_Anti_Rebound_Rail_3"),  Vol(ActiveBall) * WireformAntiRebountRailSoundFactor
      Case 2 : PlaySoundAtLevelActiveBall (Cartridge_Metal_Hits & "_Wireform_Anti_Rebound_Rail_4"),  Vol(ActiveBall) * WireformAntiRebountRailSoundFactor
      Case 3 : PlaySoundAtLevelActiveBall (Cartridge_Metal_Hits & "_Wireform_Anti_Rebound_Rail_5"),  Vol(ActiveBall) * WireformAntiRebountRailSoundFactor
      Case 4 : PlaySoundAtLevelActiveBall (Cartridge_Metal_Hits & "_Wireform_Anti_Rebound_Rail_6"),  Vol(ActiveBall) * WireformAntiRebountRailSoundFactor
      Case 5 : PlaySoundAtLevelActiveBall (Cartridge_Metal_Hits & "_Wireform_Anti_Rebound_Rail_7"),  Vol(ActiveBall) * WireformAntiRebountRailSoundFactor
    End Select
  End if
  If finalspeed < 10 Then
    Select Case Int(Rnd*3)+1
      Case 1 : PlaySoundAtLevelActiveBall (Cartridge_Metal_Hits & "_Wireform_Anti_Rebound_Rail_1"),  Vol(ActiveBall) * WireformAntiRebountRailSoundFactor
      Case 2 : PlaySoundAtLevelActiveBall (Cartridge_Metal_Hits & "_Wireform_Anti_Rebound_Rail_2"),  Vol(ActiveBall) * WireformAntiRebountRailSoundFactor
      Case 3 : PlaySoundAtLevelActiveBall (Cartridge_Metal_Hits & "_Wireform_Anti_Rebound_Rail_8"),  Vol(ActiveBall) * WireformAntiRebountRailSoundFactor
    End Select
  End if
End Sub

'////////////////////////////  LANES AND INNER LOOPS  ///////////////////////////
'////////////////////  INNER LOOPS - LEFT ENTRANCE - EVENTS  ////////////////////
Sub LeftInnerLaneTriggerUp_Hit()
  dim finalspeed
    finalspeed=SQR(activeball.velx * activeball.velx + activeball.vely * activeball.vely)
  If finalspeed > 10 then
    If ActiveBall.VelY < 0 Then RandomSoundInnerLaneEnter()
  End If
End Sub

Sub LeftInnerLaneTriggerDown_Hit()
  dim finalspeed
    finalspeed=SQR(activeball.velx * activeball.velx + activeball.vely * activeball.vely)
  If finalspeed > 7 then
    If ActiveBall.VelY > 0 Then RandomSoundInnerLaneEnter()
  End If
End Sub

'/////////////////////  INNER LOOPS - LEFT ENTRANCE - SOUNDS  ///////////////////
Sub RandomSoundInnerLaneEnter()
  PlaySoundAtLevelActiveBall (Cartridge_Metal_Hits & "_InnerLane_Ball_Guide_Hit_" & Int(Rnd*20)+1), Vol(ActiveBall) * InnerLaneSoundFactor
End Sub

'//////////////////////////////  LEFT LANE ENTRANCE  ////////////////////////////
'/////////////////////////  LEFT LANE ENTRANCE - EVENTS  ////////////////////////
Sub LeftLaneTrigger_Hit()
  dim finalspeed
    finalspeed=SQR(activeball.velx * activeball.velx + activeball.vely * activeball.vely)
  If finalspeed > 10 then
    If ActiveBall.VelY < 0 Then PlaySoundAtLevelActiveBall (Cartridge_Metal_Hits & "_Lane_Left_Ball_Enter_Hit"), Vol(ActiveBall) * LaneEnterSoundFactor : RandomSoundLaneLeftEnter()
  End If
End Sub

'/////////////////////////  LEFT LANE ENTRANCE - SOUNDS  ////////////////////////
Sub RandomSoundLaneLeftEnter()
  PlaySoundAtLevelActiveBall (Cartridge_Metal_Hits & "_Lane_Left_Ball_Roll_" & Int(Rnd*2)+1), Vol(ActiveBall) * LaneSoundFactor
End Sub

'/////////////////////////////  RIGHT LANE ENTRANCE  ////////////////////////////
'////////////////////////  RIGHT LANE ENTRANCE - EVENTS  ////////////////////////
Sub RightLaneTrigger_Hit()
  dim finalspeed
    finalspeed=SQR(activeball.velx * activeball.velx + activeball.vely * activeball.vely)
  If finalspeed > 10 then
    If ActiveBall.VelY < 0 Then PlaySoundAtLevelActiveBall (Cartridge_Metal_Hits & "_Lane_Right_Ball_Enter_Hit"), Vol(ActiveBall) * LaneEnterSoundFactor : RandomSoundLaneRightEnter()
  End If
End Sub

'/////////////////  RIGHT LANE ENTRANCE (RIGHT ORBIT) - SOUNDS  /////////////////
Sub RandomSoundLaneRightEnter()
  PlaySoundAtLevelActiveBall (Cartridge_Metal_Hits & "_Lane_Right_Ball_Roll_" & Int(Rnd*3)+1), Vol(ActiveBall) * LaneSoundFactor
End Sub

'/////////  PLASTIC LEFT RAMP - RIGHT EXIT HOLE - TO PLAYFIELD - EVENT  /////////

'sub BallDrop1_hit  'sometimes drop sound cannot be heard here.
''  debug.print activeball.velz
' RandomSoundRightRampRightExitDropToPlayfield(BallDrop1)
'end sub


Sub BallDrop2_Hit()
  if abs(activeball.AngMomZ) > 70 then activeball.AngMomZ = 70
  activeball.AngMomZ = abs(activeball.AngMomZ) * 3
  RandomSoundRightRampRightExitDropToPlayfield(BallDrop2)
' Call SoundRightPlasticRampPart2(SoundOff, ballvariablePlasticRampTimer1)
End Sub

'ss ramp
sub BallDrop3_hit
' debug.print "SS drop"
  RandomSoundRightRampRightExitDropToPlayfield(activeball)
end sub

'/////////  METAL WIRE RIGHT RAMP - EXIT HOLE - TO PLAYFIELD - EVENT  //////////
Sub RHelper3_Hit()
  if abs(activeball.AngMomZ) > 70 then activeball.AngMomZ = 70
  activeball.AngMomZ = -abs(activeball.AngMomZ) * 3
  RandomSoundLeftRampDropToPlayfield()
End Sub

'
''***************************************************************
''Table MISC VP sounds
''***************************************************************

dim FaceGuideHitsSoundLevel : FaceGuideHitsSoundLevel = 0.002 * RightPlasticRampHitsSoundLevel
Sub ColFaceGuideL1_Hit(): PlaySoundAtLevelStatic (Cartridge_Plastic_Ramps & "_Ramp_Right_Plastic_Hit_2"), FaceGuideHitsSoundLevel, activeball : End Sub
Sub F1Guide_Hit(): PlaySoundAtLevelStatic (Cartridge_Plastic_Ramps & "_Ramp_Right_Plastic_Hit_1"), FaceGuideHitsSoundLevel, activeball : End Sub
Sub F1Guide2_Hit(): PlaySoundAtLevelStatic (Cartridge_Plastic_Ramps & "_Ramp_Right_Plastic_Hit_1"), FaceGuideHitsSoundLevel, activeball : End Sub
Sub ColFaceGuideL2_Hit(): PlaySoundAtLevelStatic (Cartridge_Plastic_Ramps & "_Ramp_Right_Plastic_Hit_2"), FaceGuideHitsSoundLevel, activeball : End Sub
Sub ColFaceGuideL3_Hit(): PlaySoundAtLevelStatic (Cartridge_Plastic_Ramps & "_Ramp_Right_Plastic_Hit_3"), FaceGuideHitsSoundLevel, activeball : End Sub

sub pBlockBackhand_hit(): PlaySoundAtLevelStatic (Cartridge_Plastic_Ramps & "_Ramp_Right_Plastic_Hit_4"), FaceGuideHitsSoundLevel, activeball : End Sub
sub pBlockBackhand001_hit(): PlaySoundAtLevelStatic (Cartridge_Plastic_Ramps & "_Ramp_Right_Plastic_Hit_5"), FaceGuideHitsSoundLevel, activeball : End Sub

                          'volume level; range [0, 1]

' '///////////////////////-----Ball Drops, Bumps and Collisions-----///////////////////////
' Dim RubberStrongSoundFactor, RubberWeakSoundFactor, RubberFlipperSoundFactor,BallWithBallCollisionSoundFactor
Dim DelayedBallDropOnPlayfieldSoundLevel
DelayedBallDropOnPlayfieldSoundLevel = 0.8                  'volume level; range [0, 1]

' '/////////////////////////////  SOUND PLAYBACK FUNCTIONS  ////////////////////////////
' '/////////////////////////////  POSITIONAL SOUND PLAYBACK METHODS  ////////////////////////////

Sub PlaySoundAt(soundname, tableobj)
  PlaySound soundname, 1, 1 * MechVolume, AudioPan(tableobj), 0,0,0, 1, AudioFade(tableobj)
End Sub

Sub PlaySoundAtVol(soundname, tableobj, aVol)
  PlaySound soundname, 1, aVol * MechVolume, AudioPan(tableobj), 0,0,0, 1, AudioFade(tableobj)
End Sub

'/////////////////////////////  ROLLOVER SOUNDS  ////////////////////////////
Sub RandomSoundRollover()
  PlaySoundAtLevelActiveBall ("Rollover_" & Int(Rnd*4)+1), RolloverSoundLevel
End Sub

Sub Rollovers_Hit(idx)
  RandomSoundRollover
End Sub

' '/////////////////////////////  WALL IMPACTS  ////////////////////////////
Sub Walls_Hit(idx)
  RandomSoundWood()
End Sub

' '/////////////////////////////  METAL - EVENTS  ////////////////////////////
Sub Metals_Hit (idx)
  RandomSoundMetal
End Sub

Sub ShooterDiverter_collide(idx)
  RandomSoundMetal
End Sub

' '/////////////////////////////  TARGET HIT SOUNDS  ////////////////////////////
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

'///////////////////////////  DROP TARGET HIT SOUNDS  ///////////////////////////

Sub RandomSoundDropTargetReset(obj)
  PlaySoundAtLevelStatic SoundFX("Drop_Target_Reset_" & Int(Rnd*6)+1,DOFContactors), 1, obj
End Sub

Sub SoundDropTargetDrop(obj)
  PlaySoundAtLevelStatic ("Drop_Target_Down_" & Int(Rnd * 6) + 1), 200, obj
  'PlaySoundAtLevelStatic SoundFX("droptarget" & Int(Rnd*6)+1,DOFContactors), 1, obj
End Sub

' '/////////////////////////////  DELAYED DROP - TO PLAYFIELD - SOUND  ////////////////////////////
Sub RandomSoundDelayedBallDropOnPlayfield(aBall)
  Select Case Int(Rnd*5)+1
    Case 1 : PlaySoundAtLevelStatic ("Ball_Drop_Playfield_1_Delayed"), DelayedBallDropOnPlayfieldSoundLevel, aBall
    Case 2 : PlaySoundAtLevelStatic ("Ball_Drop_Playfield_2_Delayed"), DelayedBallDropOnPlayfieldSoundLevel, aBall
    Case 3 : PlaySoundAtLevelStatic ("Ball_Drop_Playfield_3_Delayed"), DelayedBallDropOnPlayfieldSoundLevel, aBall
    Case 4 : PlaySoundAtLevelStatic ("Ball_Drop_Playfield_4_Delayed"), DelayedBallDropOnPlayfieldSoundLevel, aBall
    Case 5 : PlaySoundAtLevelStatic ("Ball_Drop_Playfield_5_Delayed"), DelayedBallDropOnPlayfieldSoundLevel, aBall
  End Select
End Sub

'/////////////////////////////////////////////////////////////////
'         End Mechanical Sounds
'/////////////////////////////////////////////////////////////////

'******************************************************
'****  FLEEP MECHANICAL SOUNDS
'******************************************************

'******************************************************
'   ZFLD:  FLUPPER DOMES
'******************************************************
' Based on FlupperDomes2.2

' What you need in your table to use these flashers:
' Open this table and your table both in VPX
' Export all the materials domebasemat, Flashermaterial0 - 20 and import them in your table
' Export all textures (images) starting with the name "dome" and "ronddome" and import them into your table with the same names
' Export all textures (images) starting with the name "flasherbloom" and import them into your table with the same names
' Copy a set of 4 objects flasherbase, flasherlit, flasherlight and flasherflash from layer 7 to your table
' If you duplicate the four objects for a new flasher dome, be sure that they all end with the same number (in the 0-20 range)
' Copy the flasherbloom flashers from layer 10 to your table. you will need to make one per flasher dome that you plan to make
' Select the correct flasherbloom texture for each flasherbloom flasher, per flasher dome
' Copy the script below

' Place your flasher base primitive where you want the flasher located on your Table
' Then run InitFlasher in the script with the number of your flasher objects and the color of the flasher.  This will align the flasher object, light object, and
' flasher lit primitive.  It will also assign the appropriate flasher bloom images to the flasher bloom object.
'
' Example: InitFlasher 1, "green"
'
' Color Options: "blue", "green", "red", "purple", "yellow", "white", and "orange"

' You can use the RotateFlasher call to align the Rotz/ObjRotz of the flasher primitives with "handles".  Don't set those values in the editor,
' call the RotateFlasher sub instead (this call will likely crash VP if it's call for the flasher primitives without "handles")
'
' Example: RotateFlasher 1, 180    'where 1 is the flasher number and 180 is the angle of Z rotation

' For flashing the flasher use in the script: "ObjLevel(1) = 1 : FlasherFlash1_Timer"
' This should also work for flashers with variable flash levels from the rom, just use ObjLevel(1) = xx from the rom (in the range 0-1)
'
' Notes (please read!!):
' - Setting TestFlashers = 1 (below in the ScriptsDirectory) will allow you to see how the flasher objects are aligned (need the targetflasher image imported to your table)
' - The rotation of the primitives with "handles" is done with a script command, not on the primitive itself (see RotateFlasher below)
' - Color of the objects are set in the script, not on the primitive itself
' - Screws are optional to copy and position manually
' - If your table is not named "Table1" then change the name below in the script
' - Every flasher uses its own material (Flashermaterialxx), do not use it for anything else
' - Lighting > Bloom Strength affects how the flashers look, do not set it too high
' - Change RotY and RotX of flasherbase only when having a flasher something other then parallel to the playfield
' - Leave RotX of the flasherflash object to -45; this makes sure that the flash effect is visible in FS and DT
' - If you want to resize a flasher, be sure to resize flasherbase, flasherlit and flasherflash with the same percentage
' - If you think that the flasher effects are too bright, change flasherlightintensity and/or flasherflareintensity below

' Some more notes for users of the v1 flashers and/or JP's fading lights routines:
' - Delete all textures/primitives/script/materials in your table from the v1 flashers and scripts before you start; they don't mix well with v2
' - Remove flupperflash(m) routines if you have them; they do not work with this new script
' - Do not try to mix this v2 script with the JP fading light routine (that is making it too complicated), just use the example script below

'------ Main Dome Code ---------'

Dim TestFlashers, TableRef, FlasherLightIntensity, FlasherFlareIntensity, FlasherBloomIntensity, FlasherOffBrightness

' *********************************************************************
TestFlashers = 0        ' *** set this to 1 to check position of flasher object      ***
Set TableRef = Table1      ' *** change this, if your table has another name           ***
FlasherLightIntensity = 0.1  ' *** lower this, if the VPX lights are too bright (i.e. 0.1)     ***
FlasherFlareIntensity = 0.3  ' *** lower this, if the flares are too bright (i.e. 0.1)       ***
FlasherBloomIntensity = 0.2  ' *** lower this, if the blooms are too bright (i.e. 0.1)       ***
FlasherOffBrightness = 0.5    ' *** brightness of the flasher dome when switched off (range 0-2)  ***
' *********************************************************************

Dim ObjLevel(20), objbase(20), objlit(20), objflasher(20), objbloom(20), objlight(20), ObjTargetLevel(20)
'Dim tablewidth, tableheight : tablewidth = TableRef.width : tableheight = TableRef.height

'initialise the flasher color, you can only choose from "green", "red", "purple", "blue", "white" and "yellow"
InitFlasher 1, "orange"
InitFlasher 2, "white"
InitFlasher 3, "orange"

' rotate the flasher with the command below (first argument = flasher nr, second argument = angle in degrees)
'   RotateFlasher 1,17
'   RotateFlasher 2,0
'   RotateFlasher 3,90
'   RotateFlasher 4,90

Sub InitFlasher(nr, col)
  ' store all objects in an array for use in FlashFlasher subroutine
  Set objbase(nr) = Eval("Flasherbase" & nr)
  Set objlit(nr) = Eval("Flasherlit" & nr)
  Set objflasher(nr) = Eval("Flasherflash" & nr)
  Set objlight(nr) = Eval("Flasherlight" & nr)
  Set objbloom(nr) = Eval("Flasherbloom" & nr)

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
  If objbase(nr).x >= xthird And objbase(nr).x <= xthird * 2 Then
    objbloom(nr).imageA = "flasherbloomCenter"
    objbloom(nr).imageB = "flasherbloomCenter"
  ElseIf objbase(nr).x < xthird And objbase(nr).y < ythird Then
    objbloom(nr).imageA = "flasherbloomUpperLeft"
    objbloom(nr).imageB = "flasherbloomUpperLeft"
  ElseIf  objbase(nr).x > xthird * 2 And objbase(nr).y < ythird Then
    objbloom(nr).imageA = "flasherbloomUpperRight"
    objbloom(nr).imageB = "flasherbloomUpperRight"
  ElseIf objbase(nr).x < xthird And objbase(nr).y < ythird * 2 Then
    objbloom(nr).imageA = "flasherbloomCenterLeft"
    objbloom(nr).imageB = "flasherbloomCenterLeft"
  ElseIf  objbase(nr).x > xthird * 2 And objbase(nr).y < ythird * 2 Then
    objbloom(nr).imageA = "flasherbloomCenterRight"
    objbloom(nr).imageB = "flasherbloomCenterRight"
  ElseIf objbase(nr).x < xthird And objbase(nr).y < ythird * 3 Then
    objbloom(nr).imageA = "flasherbloomLowerLeft"
    objbloom(nr).imageB = "flasherbloomLowerLeft"
  ElseIf  objbase(nr).x > xthird * 2 And objbase(nr).y < ythird * 3 Then
    objbloom(nr).imageA = "flasherbloomLowerRight"
    objbloom(nr).imageB = "flasherbloomLowerRight"
  End If

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
    objbloom(nr).color = RGB(4,120,255)
    objlight(nr).intensity = 5000

    Case "green"
    objlight(nr).color = RGB(12,255,4)
    objflasher(nr).color = RGB(12,255,4)
    objbloom(nr).color = RGB(12,255,4)

    Case "red"
    objlight(nr).color = RGB(255,32,4)
    objflasher(nr).color = RGB(255,32,4)
    objbloom(nr).color = RGB(255,32,4)

    Case "purple"
    objlight(nr).color = RGB(230,49,255)
    objflasher(nr).color = RGB(255,64,255)
    objbloom(nr).color = RGB(230,49,255)

    Case "yellow"
    objlight(nr).color = RGB(200,173,25)
    objflasher(nr).color = RGB(255,200,50)
    objbloom(nr).color = RGB(200,173,25)

    Case "white"
    objlight(nr).color = RGB(255,240,150)
    objflasher(nr).color = RGB(100,86,59)
    objbloom(nr).color = RGB(255,240,150)

    Case "orange"
    objlight(nr).color = RGB(255,70,0)
    objflasher(nr).color = RGB(255,70,0)
    objbloom(nr).color = RGB(255,70,0)
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

''------ Use this for PWM following domes ---------'

Sub ModFlashFlasher(nr, aValue)
  objflasher(nr).visible = 1 : objbloom(nr).visible = 1 : objlit(nr).visible = 1
  objflasher(nr).opacity = 1000 *  FlasherFlareIntensity * aValue
  objbloom(nr).opacity = 100 *  FlasherBloomIntensity * aValue
  objlight(nr).IntensityScale = 0.5 * FlasherLightIntensity * aValue
  objbase(nr).BlendDisableLighting =  FlasherOffBrightness + 10 * aValue
  objlit(nr).BlendDisableLighting = 10 * aValue
  UpdateMaterial "Flashermaterial" & nr,0,0,0,0,0,0,aValue,RGB(255,255,255),0,0,False,True,0,0,0,0
End Sub

''------ Use this if you are not using PWM following domes ---------'

Sub FlashFlasher(nr)
  If Not objflasher(nr).TimerEnabled Then
    objflasher(nr).TimerEnabled = True
    objflasher(nr).visible = 1
    objbloom(nr).visible = 1
    objlit(nr).visible = 1
  End If
  objflasher(nr).opacity = 1000 * FlasherFlareIntensity * ObjLevel(nr) ^ 2.5
  objbloom(nr).opacity = 100 * FlasherBloomIntensity * ObjLevel(nr) ^ 2.5
  objlight(nr).IntensityScale = 0.5 * FlasherLightIntensity * ObjLevel(nr) ^ 3
  objbase(nr).BlendDisableLighting = FlasherOffBrightness + 10 * ObjLevel(nr) ^ 3
  objlit(nr).BlendDisableLighting = 10 * ObjLevel(nr) ^ 2
  UpdateMaterial "Flashermaterial" & nr,0,0,0,0,0,0,ObjLevel(nr),RGB(255,255,255),0,0,False,True,0,0,0,0
  If Round(ObjTargetLevel(nr),1) > Round(ObjLevel(nr),1) Then
    ObjLevel(nr) = ObjLevel(nr) + 0.3
    If ObjLevel(nr) > 1 Then ObjLevel(nr) = 1
  ElseIf Round(ObjTargetLevel(nr),1) < Round(ObjLevel(nr),1) Then
    ObjLevel(nr) = ObjLevel(nr) * 0.85 - 0.01
    If ObjLevel(nr) < 0 Then ObjLevel(nr) = 0
  Else
    ObjLevel(nr) = Round(ObjTargetLevel(nr),1)
    objflasher(nr).TimerEnabled = False
  End If
  '   ObjLevel(nr) = ObjLevel(nr) * 0.9 - 0.01
  If ObjLevel(nr) < 0 Then
    objflasher(nr).TimerEnabled = False
    objflasher(nr).visible = 0
    objbloom(nr).visible = 0
    objlit(nr).visible = 0
  End If
End Sub

Sub FlasherFlash1_Timer()
  FlashFlasher(1)
End Sub
Sub FlasherFlash2_Timer()
  FlashFlasher(2)
End Sub
Sub FlasherFlash3_Timer()
  FlashFlasher(3)
End Sub

'******************************************************
'******  END FLUPPER DOMES
'******************************************************

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
BumperTimer.Enabled = 1
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
FlInitBumper 1, "white"
FlInitBumper 2, "white"
FlInitBumper 3, "white"
FlInitBumper 4, "red"

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

    Case "blue"
    FlBumperBigLight(nr).color = RGB(32,80,255)
    FlBumperBigLight(nr).colorfull = RGB(32,80,255)
    FlBumperSmallLight(nr).color = RGB(0,80,255)
    FlBumperSmallLight(nr).colorfull = RGB(0,80,255)
    FlBumperSmallLight(nr).TransmissionScale = 0
    MaterialColor "bumpertopmat" & nr, RGB(8,120,255)
    FlBumperHighlight(nr).color = RGB(255,16,8)
    FlBumperSmallLight(nr).BulbModulateVsAdd = 1

    Case "green"
    FlBumperSmallLight(nr).color = RGB(8,255,8)
    FlBumperSmallLight(nr).colorfull = RGB(8,255,8)
    FlBumperBigLight(nr).color = RGB(32,255,32)
    FlBumperBigLight(nr).colorfull = RGB(32,255,32)
    FlBumperHighlight(nr).color = RGB(255,32,255)
    MaterialColor "bumpertopmat" & nr, RGB(16,255,16)
    FlBumperSmallLight(nr).TransmissionScale = 0.005
    FlBumperSmallLight(nr).BulbModulateVsAdd = 1

    Case "orange"
    FlBumperHighlight(nr).color = RGB(255,130,255)
    FlBumperSmallLight(nr).BulbModulateVsAdd = 1
    FlBumperSmallLight(nr).TransmissionScale = 0
    FlBumperSmallLight(nr).color = RGB(255,130,0)
    FlBumperSmallLight(nr).colorfull = RGB (255,90,0)
    FlBumperBigLight(nr).color = RGB(255,190,8)
    FlBumperBigLight(nr).colorfull = RGB(255,190,8)

    Case "white"
    FlBumperBigLight(nr).color = RGB(255,230,190)
    FlBumperBigLight(nr).colorfull = RGB(255,230,190)
    FlBumperHighlight(nr).color = RGB(255,180,100)
    FlBumperSmallLight(nr).TransmissionScale = 0
    FlBumperSmallLight(nr).BulbModulateVsAdd = 0.99

    Case "blacklight"
    FlBumperBigLight(nr).color = RGB(32,32,255)
    FlBumperBigLight(nr).colorfull = RGB(32,32,255)
    FlBumperHighlight(nr).color = RGB(48,8,255)
    FlBumperSmallLight(nr).TransmissionScale = 0
    FlBumperSmallLight(nr).BulbModulateVsAdd = 1

    Case "yellow"
    FlBumperSmallLight(nr).color = RGB(255,230,4)
    FlBumperSmallLight(nr).colorfull = RGB(255,230,4)
    FlBumperBigLight(nr).color = RGB(255,240,50)
    FlBumperBigLight(nr).colorfull = RGB(255,240,50)
    FlBumperHighlight(nr).color = RGB(255,255,220)
    FlBumperSmallLight(nr).BulbModulateVsAdd = 1
    FlBumperSmallLight(nr).TransmissionScale = 0

    Case "purple"
    FlBumperBigLight(nr).color = RGB(80,32,255)
    FlBumperBigLight(nr).colorfull = RGB(80,32,255)
    FlBumperSmallLight(nr).color = RGB(80,32,255)
    FlBumperSmallLight(nr).colorfull = RGB(80,32,255)
    FlBumperSmallLight(nr).TransmissionScale = 0
    FlBumperHighlight(nr).color = RGB(32,64,255)
    FlBumperSmallLight(nr).BulbModulateVsAdd = 1
  End Select
End Sub

Sub FlFadeBumper(nr, Z)
  FlBumperBase(nr).BlendDisableLighting = 0.5 * DayNightAdjust
  '   UpdateMaterial(string, float wrapLighting, float roughness, float glossyImageLerp, float thickness, float edge, float edgeAlpha, float opacity,
  '        OLE_COLOR base, OLE_COLOR glossy, OLE_COLOR clearcoat, VARIANT_BOOL isMetal, VARIANT_BOOL opacityActive,
  '        float elasticity, float elasticityFalloff, float friction, float scatterAngle) - updates all parameters of a material
  FlBumperDisk(nr).BlendDisableLighting = (0.5 - Z * 0.3 ) * DayNightAdjust

  Select Case FlBumperColor(nr)
    Case "blue"
    UpdateMaterial "bumperbulbmat" & nr, 0, 0.75 , 0, 1 - Z, 1 - Z, 1 - Z, 0.9999, RGB(38 - 24 * Z,130 - 98 * Z,255), RGB(255,255,255), RGB(32,32,32), False, True, 0, 0, 0, 0
    FlBumperSmallLight(nr).intensity = 20 + 500 * Z / (0.5 + DNA30)
    FlBumperTop(nr).BlendDisableLighting = 3 * DayNightAdjust + 50 * Z
    FlBumperBulb(nr).BlendDisableLighting = 12 * DayNightAdjust + 5000 * (0.03 * Z + 0.97 * Z ^ 3)
    Flbumperbiglight(nr).intensity = 25 * Z / (1 + DNA45)
    FlBumperHighlight(nr).opacity = 10000 * (Z ^ 3) / (0.5 + DNA90)

    Case "green"
    UpdateMaterial "bumperbulbmat" & nr, 0, 0.75 , 0, 1 - Z, 1 - Z, 1 - Z, 0.9999, RGB(16 + 16 * Sin(Z * 3.14),255,16 + 16 * Sin(Z * 3.14)), RGB(255,255,255), RGB(32,32,32), False, True, 0, 0, 0, 0
    FlBumperSmallLight(nr).intensity = 10 + 150 * Z / (1 + DNA30)
    FlBumperTop(nr).BlendDisableLighting = 2 * DayNightAdjust + 20 * Z
    FlBumperBulb(nr).BlendDisableLighting = 7 * DayNightAdjust + 6000 * (0.03 * Z + 0.97 * Z ^ 10)
    Flbumperbiglight(nr).intensity = 10 * Z / (1 + DNA45)
    FlBumperHighlight(nr).opacity = 6000 * (Z ^ 3) / (1 + DNA90)

    Case "red"
    UpdateMaterial "bumperbulbmat" & nr, 0, 0.75 , 0, 1 - Z, 1 - Z, 1 - Z, 0.9999, RGB(255, 16 - 11 * Z + 16 * Sin(Z * 3.14),0), RGB(255,255,255), RGB(32,32,32), False, True, 0, 0, 0, 0
    FlBumperSmallLight(nr).intensity = 17 + 100 * Z / (1 + DNA30 ^ 2)
    FlBumperTop(nr).BlendDisableLighting = 3 * DayNightAdjust + 18 * Z / (1 + DNA90)
    FlBumperBulb(nr).BlendDisableLighting = 20 * DayNightAdjust + 9000 * (0.03 * Z + 0.97 * Z ^ 10)
    Flbumperbiglight(nr).intensity = 10 * Z / (1 + DNA45)
    FlBumperHighlight(nr).opacity = 2000 * (Z ^ 3) / (1 + DNA90)
    MaterialColor "bumpertopmat" & nr, RGB(255,20 + Z * 4,8 - Z * 8)

    Case "orange"
    UpdateMaterial "bumperbulbmat" & nr, 0, 0.75 , 0, 1 - Z, 1 - Z, 1 - Z, 0.9999, RGB(255, 100 - 22 * z + 16 * Sin(Z * 3.14),Z * 32), RGB(255,255,255), RGB(32,32,32), False, True, 0, 0, 0, 0
    FlBumperSmallLight(nr).intensity = 17 + 250 * Z / (1 + DNA30 ^ 2)
    FlBumperTop(nr).BlendDisableLighting = 3 * DayNightAdjust + 50 * Z / (1 + DNA90)
    FlBumperBulb(nr).BlendDisableLighting = 15 * DayNightAdjust + 2500 * (0.03 * Z + 0.97 * Z ^ 10)
    Flbumperbiglight(nr).intensity = 10 * Z / (1 + DNA45)
    FlBumperHighlight(nr).opacity = 4000 * (Z ^ 3) / (1 + DNA90)
    MaterialColor "bumpertopmat" & nr, RGB(255,100 + Z * 50, 0)

    Case "white"
    UpdateMaterial "bumperbulbmat" & nr, 0, 0.75 , 0, 1 - Z, 1 - Z, 1 - Z, 0.9999, RGB(255,230 - 100 * Z, 200 - 150 * Z), RGB(255,255,255), RGB(32,32,32), False, True, 0, 0, 0, 0
    FlBumperSmallLight(nr).intensity = 20 + 180 * Z / (1 + DNA30)
    FlBumperTop(nr).BlendDisableLighting = 5 * DayNightAdjust + 30 * Z
    FlBumperBulb(nr).BlendDisableLighting = 18 * DayNightAdjust + 3000 * (0.03 * Z + 0.97 * Z ^ 10)
    Flbumperbiglight(nr).intensity = 8 * Z / (1 + DNA45)
    FlBumperHighlight(nr).opacity = 1000 * (Z ^ 3) / (1 + DNA90)
    FlBumperSmallLight(nr).color = RGB(255,255 - 20 * Z,255 - 65 * Z)
    FlBumperSmallLight(nr).colorfull = RGB(255,255 - 20 * Z,255 - 65 * Z)
    MaterialColor "bumpertopmat" & nr, RGB(255,235 - z * 36,220 - Z * 90)

    Case "blacklight"
    UpdateMaterial "bumperbulbmat" & nr, 0, 0.75 , 0, 1 - Z, 1 - Z, 1 - Z, 1, RGB(30 - 27 * Z ^ 0.03,30 - 28 * Z ^ 0.01, 255), RGB(255,255,255), RGB(32,32,32), False, True, 0, 0, 0, 0
    FlBumperSmallLight(nr).intensity = 20 + 900 * Z / (1 + DNA30)
    FlBumperTop(nr).BlendDisableLighting = 3 * DayNightAdjust + 60 * Z
    FlBumperBulb(nr).BlendDisableLighting = 15 * DayNightAdjust + 30000 * Z ^ 3
    Flbumperbiglight(nr).intensity = 25 * Z / (1 + DNA45)
    FlBumperHighlight(nr).opacity = 2000 * (Z ^ 3) / (1 + DNA90)
    FlBumperSmallLight(nr).color = RGB(255 - 240 * (Z ^ 0.1),255 - 240 * (Z ^ 0.1),255)
    FlBumperSmallLight(nr).colorfull = RGB(255 - 200 * z,255 - 200 * Z,255)
    MaterialColor "bumpertopmat" & nr, RGB(255 - 190 * Z,235 - z * 180,220 + 35 * Z)

    Case "yellow"
    UpdateMaterial "bumperbulbmat" & nr, 0, 0.75 , 0, 1 - Z, 1 - Z, 1 - Z, 0.9999, RGB(255, 180 + 40 * z, 48 * Z), RGB(255,255,255), RGB(32,32,32), False, True, 0, 0, 0, 0
    FlBumperSmallLight(nr).intensity = 17 + 200 * Z / (1 + DNA30 ^ 2)
    FlBumperTop(nr).BlendDisableLighting = 3 * DayNightAdjust + 40 * Z / (1 + DNA90)
    FlBumperBulb(nr).BlendDisableLighting = 12 * DayNightAdjust + 2000 * (0.03 * Z + 0.97 * Z ^ 10)
    Flbumperbiglight(nr).intensity = 10 * Z / (1 + DNA45)
    FlBumperHighlight(nr).opacity = 1000 * (Z ^ 3) / (1 + DNA90)
    MaterialColor "bumpertopmat" & nr, RGB(255,200, 24 - 24 * z)

    Case "purple"
    UpdateMaterial "bumperbulbmat" & nr, 0, 0.75 , 0, 1 - Z, 1 - Z, 1 - Z, 0.9999, RGB(128 - 118 * Z - 32 * Sin(Z * 3.14), 32 - 26 * Z ,255), RGB(255,255,255), RGB(32,32,32), False, True, 0, 0, 0, 0
    FlBumperSmallLight(nr).intensity = 15 + 200 * Z / (0.5 + DNA30)
    FlBumperTop(nr).BlendDisableLighting = 3 * DayNightAdjust + 50 * Z
    FlBumperBulb(nr).BlendDisableLighting = 15 * DayNightAdjust + 10000 * (0.03 * Z + 0.97 * Z ^ 3)
    Flbumperbiglight(nr).intensity = 25 * Z / (1 + DNA45)
    FlBumperHighlight(nr).opacity = 4000 * (Z ^ 3) / (0.5 + DNA90)
    MaterialColor "bumpertopmat" & nr, RGB(128 - 60 * Z,32,255)
  End Select
End Sub

'******************************************************
'******  END FLUPPER BUMPERS
'******************************************************

'*******************************************
'  VPW TWEAK MENU
'*******************************************

Dim ColorLUT : ColorLUT = 1                 ' Color desaturation LUTs: 1 to 11, where 1 is normal and 11 is black'n'white
Dim Anaglyph : Anaglyph = 0                           ' When set to 1 in the menu this will overwrite the default LUT slot to use the VPW Original 1 on 1
Dim MechVolume : MechVolume = 0.8                   ' Overall Mechanical sound effect volume. Recommended values should be no greater than 1.
Dim BallRollVolume : BallRollVolume = 0.5           ' Level of ball rolling volume. Value between 0 and 1
Dim RampRollVolume : RampRollVolume = 0.5           ' Level of ramp rolling volume. Value between 0 and 1
Dim StagedFlippers : StagedFlippers = 0               ' Staged Flippers. 0 = Disabled, 1 = Enabled
Dim AmbientBallShadowOn : AmbientBallShadowon = 1     ' Ambient Ball Shadows. 0 = Disabled, 1 = Enabled
Dim RenderProbeOpt : RenderProbeOpt = 2           ' 0 = No Refraction Probes (best performance), 1 = Refraction Probes with Roughness '0' (improved performance), 2 = Full Refraction Probes (best visual)
Dim PlayfieldReflections : PlayfieldReflections = 100 ' Defines the reflection strength of the (dynamic) table elements on the playfield (0-100) / 5 = 20 Max
Dim ReflOpt : ReflOpt = 1                 ' 0 = Reflections off, 1 = Reflections on
Dim LightReflOpt : LightReflOpt = 1             ' 0 = Light Reflections off, 1 = Light Reflections on
Dim InsertReflOpt : InsertReflOpt = 1           ' 0 = Insert Reflections off, 1 = Insert Reflections on
Dim BackglassReflOpt : BackglassReflOpt = 1           ' 0 = Backglass Reflections off, 1 = Backglass Reflections on
Dim TopperReflOpt : TopperReflOpt = 1         ' 0 = Topper Reflections off, 1 = Topper Reflections on

Dim InstMod, SideBladeMod, GlassMod, LegendsCabMod, TopperMod, BackglassMod, LogoPosterMod

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
  'Color Saturation
    ColorLUT = Table1.Option("Color Saturation", 1, 27, 1, 1, 0, _
    Array("Normal", "Desaturated 10%", "Desaturated 20%", "Desaturated 30%", "Desaturated 40%", "Desaturated 50%", "Desaturated 60%", "Desaturated 70%", "Desaturated 80%", "Desaturated 90%", "Black 'n White", _
    "Fleep Natural Dark 1", "Fleep Natural Dark 2", "Fleep Warm Dark", "Fleep Warm Bright", "Fleep Warm Vivid Soft", "Fleep Warm Vivid Hard", "Skitso Natural & Balance", "Skitso Natural High Contrast", _
    "3rdaxis THX Standard", "Callev Brightness & Contrast", "Hauntfreaks Desaturated", "Tomate Washed out", "VPW Original 1 on 1", "Bassgeige", "Blacklight", "B&W Comic Book"))
  if ColorLUT = 1 And Anaglyph = 0 Then Table1.ColorGradeImage = "" '"Normal"
  if ColorLUT = 1 And Anaglyph = 1 Then Table1.ColorGradeImage = "colorgradelut256x16-vpw-original-1-on-1" '"VPW Original 1 on 1"
  if ColorLUT = 2 Then Table1.ColorGradeImage = "colorgradelut256x16-10" '"Desaturated 10%"
  if ColorLUT = 3 Then Table1.ColorGradeImage = "colorgradelut256x16-20" '"Desaturated 20%"
  if ColorLUT = 4 Then Table1.ColorGradeImage = "colorgradelut256x16-30" '"Desaturated 30%"
  if ColorLUT = 5 Then Table1.ColorGradeImage = "colorgradelut256x16-40" '"Desaturated 40%"
  if ColorLUT = 6 Then Table1.ColorGradeImage = "colorgradelut256x16-50" '"Desaturated 50%"
  if ColorLUT = 7 Then Table1.ColorGradeImage = "colorgradelut256x16-60" '"Desaturated 60%"
  if ColorLUT = 8 Then Table1.ColorGradeImage = "colorgradelut256x16-70" '"Desaturated 70%"
  if ColorLUT = 9 Then Table1.ColorGradeImage = "colorgradelut256x16-80" '"Desaturated 80%"
  if ColorLUT = 10 Then Table1.ColorGradeImage = "colorgradelut256x16-90" '"Desaturated 90%"
  if ColorLUT = 11 Then Table1.ColorGradeImage = "colorgradelut256x16-100" '"Black 'n White"
  if ColorLUT = 12 Then Table1.ColorGradeImage = "colorgradelut256x16-fleep-natural-dark-1" '"Fleep Natural Dark 1"
  if ColorLUT = 13 Then Table1.ColorGradeImage = "colorgradelut256x16-fleep-natural-dark-2" '"Fleep Natural Dark 2"
  if ColorLUT = 14 Then Table1.ColorGradeImage = "colorgradelut256x16-fleep-warm-dark" '"Fleep Warm Dark "
  if ColorLUT = 15 Then Table1.ColorGradeImage = "colorgradelut256x16-fleep-warm-bright" '"Fleep Warm Bright"
  if ColorLUT = 16 Then Table1.ColorGradeImage = "colorgradelut256x16-fleep-warm-vivid-soft" '"Fleep Warm Vivid Soft"
  if ColorLUT = 17 Then Table1.ColorGradeImage = "colorgradelut256x16-fleep-warm-vivid-hard" '"Fleep Warm Vivid Hard"
  if ColorLUT = 18 Then Table1.ColorGradeImage = "colorgradelut256x16-skitso-natural-and-balance" '"Skitso Natural & Balance"
  if ColorLUT = 19 Then Table1.ColorGradeImage = "colorgradelut256x16-skitso-natural-high-contrast" '"Skitso Natural High Contrast"
  if ColorLUT = 20 Then Table1.ColorGradeImage = "colorgradelut256x16-3rdaxis-thx-standard" '"3rdaxis THX Standard"
  if ColorLUT = 21 Then Table1.ColorGradeImage = "colorgradelut256x16-callev-brightness-and-contrast" '"Callev Brightness & Contrast"
  if ColorLUT = 22 Then Table1.ColorGradeImage = "colorgradelut256x16-hauntfreaks-desaturated" '"Hauntfreaks Desaturated"
  if ColorLUT = 23 Then Table1.ColorGradeImage = "colorgradelut256x16-tomate-washed-out" '"Tomate Washed out"
  if ColorLUT = 24 Then Table1.ColorGradeImage = "colorgradelut256x16-vpw-original-1-on-1" '"VPW Original 1 on 1"
  if ColorLUT = 25 Then Table1.ColorGradeImage = "colorgradelut256x16-bassgeige" '"Bassgeige"
  if ColorLUT = 26 Then Table1.ColorGradeImage = "colorgradelut256x16-blacklight" '"Blacklight"
  if ColorLUT = 27 Then Table1.ColorGradeImage = "colorgradelut256x16-b-and-w-comic-book" '"B&W Comic Book"
  'Anaglyph
  Anaglyph = Table1.Option("Anaglyph 3D/Headtracking Fix", 0, 1, 1, 0, 0, Array("Off", "On"))
  SetupAnaglyph
    'Sound volumes
    MechVolume = Table1.Option("Mech Volume", 0, 1, 0.01, 0.8, 1)
    BallRollVolume = Table1.Option("Ball Roll Volume", 0, 1, 0.01, 0.5, 1)
  RampRollVolume = Table1.Option("Ramp Roll Volume", 0, 1, 0.01, 0.5, 1)
  'Ambient BallShadow
  AmbientBallShadowOn = Table1.Option("Ambient Ballshadow Ballshadow", 0, 1, 1, 1, 0, Array("Off", "On"))
  SetupAmbientBallshadow
  'Reflections
  ReflOpt = Table1.Option("Reflections", 0, 1, 1, 1, 0, Array("Off", "On"))
  SetRefl ReflOpt
  'Playfield Reflections
  PlayfieldReflections = Table1.Option("Playfield Reflections", 0, 1, 0.05, 1, 1)
  Table1.PlayfieldReflectionStrength = (PlayfieldReflections * 20)
  'Light Reflections
  LightReflOpt = Table1.Option("Light Reflections", 0, 1, 1, 1, 0, Array("Off", "On"))
  SetLightRefl LightReflOpt
  'Insert Reflections
  InsertReflOpt = Table1.Option("Insert Reflections", 0, 1, 1, 1, 0, Array("Off", "On"))
  SetInsertRefl InsertReflOpt
  'Backglass Reflections
  BackglassReflOpt = Table1.Option("Backglass Reflections", 0, 1, 1, 1, 0, Array("Off", "On"))
  SetBackglassRefl BackglassReflOpt
  'Topper Reflections
  TopperReflOpt = Table1.Option("Topper Reflections", 0, 1, 1, 1, 0, Array("Off", "On"))
  SetTopperRefl TopperReflOpt
  'Refraction Probe Setting
  RenderProbeOpt = Table1.Option("Refraction Probe Setting", 0, 1, 1, 1, 0, Array("No Refraction (best performance)", "Full Refraction (best visuals)"))
  SetRenderProbes RenderProbeOpt
  'Instruction Cards
  InstMod = Table1.Option("Instruction Cards", 0, 2, 1, 0, 0, Array("Default Instruction Cards", "Custom Instruction Cards 1", "Custom Instruction Cards 2"))
  SetupInstructionCards
  'Sideblades
  SideBladeMod = Table1.Option("Side Blades", 0, 1, 1, 0, 0, Array("Original", "Tiltgraphics"))
  SetupBlades
  'LegendsCab
  LegendsCabMod = Table1.Option("Legends Cabinet", 0, 1, 1, 0, 0, Array("Off", "On"))
  SetupLegendsCab
  'Scratched Glass
  GlassMod = Table1.Option("Scratched Glass", 0, 1, 1, 0, 0, Array("Off", "On"))
  SetupGlass
  'VR Room
    VRRoomChoice = Table1.Option("VR Room (VR-FSS)", 1, 3, 1, 1, 0, Array("Mega Room", "Minimal Room", "Mixed Reality"))
  SetupRoom
  'Topper Glass (VR-FSS)
  TopperMod = Table1.Option("Topper Glass (VR-FSS)", 0, 1, 1, 0, 0, Array("Normal Topper", "Classic Playfield Reproductions"))
  SetupTopper
  'Backglass
  BackglassMod = Table1.Option("Backglass (VR-FSS)", 0, 1, 1, 0, 0, Array("Original", "Classic Playfield Reproductions"))
  SetupBackglass
  'VR Minimal Room Logo & Poster (VR-FSS)
  LogoPosterMod = Table1.Option("VR Minimal Room Logo & Poster (VR-FSS)", 0, 1, 1, 1, 0, Array("Off", "On"))
  SetupLogoPoster
  If eventId = 3 And dspTriggered Then dspTriggered = False : DisableStaticPreRendering = False : End If
End Sub

Sub SetupAnaglyph
  if ColorLUT = 1 And Anaglyph = 0 Then Table1.ColorGradeImage = "" '"Normal"
  if ColorLUT = 1 And Anaglyph = 1 Then Table1.ColorGradeImage = "colorgradelut256x16-vpw-original-1-on-1" '"VPW Original 1 on 1"
End Sub

Sub SetupAmbientBallshadow
  If AmbientBallShadowOn = 1 Then
    For each xx in BriteLights : xx.Shadows = 1 : Next
    For each xx in GILights : xx.Shadows = 1 : Next
  Else
    For each xx in GILights : xx.Shadows = 0 : Next
    For each xx in BriteLights : xx.Shadows = 0 : Next
  End If
End Sub

Sub SetRefl(Opt)
  Select Case Opt
    Case 0:
      For each xx in Reflections: xx.ReflectionEnabled = False : Next
    Case 1:
      For each xx in Reflections: xx.ReflectionEnabled = True : Next
  End Select
End Sub

Sub SetPlayfieldRefl(Opt)
  Table1.PlayfieldReflectionStrength = (PlayfieldReflections * 20)
End Sub

Sub SetLightRefl(Opt)
  Select Case Opt
    Case 0:
      For each xx in LightReflections: xx.ShowReflectionOnBall = False : Next
    Case 1:
      For each xx in LightReflections: xx.ShowReflectionOnBall = True : Next
  End Select
End Sub

Sub SetInsertRefl(Opt)
  Select Case Opt
    Case 0:
      For each xx in InsertReflections: xx.ShowReflectionOnBall = False : Next
    Case 1:
      For each xx in InsertReflections: xx.ShowReflectionOnBall = True : Next
  End Select
End Sub

Sub SetBackglassRefl(Opt)
  Select Case Opt
    Case 0:
      For each xx in BackglassReflections: xx.ReflectionEnabled = False : Next
      If VRMode = False Then
        For Each VR_Obj in BackglassReflections : VR_Obj.Visible = 0 : Next
      End If
    Case 1:
      For each xx in BackglassReflections: xx.ReflectionEnabled = True : Next
      For Each VR_Obj in BackglassReflections : VR_Obj.Visible = 1 : Next
  End Select
End Sub

Sub SetTopperRefl(Opt)
  Select Case Opt
    Case 0:
      For each xx in TopperReflections: xx.ReflectionEnabled = False : Next
      Topper_FlasherDragon.ReflectionEnabled = False
      Topper_FlasherLeft.ReflectionEnabled = False
      Topper_FlasherRight.ReflectionEnabled = False
      If VRMode = False Then
        For Each VR_Obj in TopperReflections : VR_Obj.Visible = 0 : Next
      End If
    Case 1:
      For each xx in TopperReflections: xx.ReflectionEnabled = True : Next
      For Each VR_Obj in TopperReflections : VR_Obj.Visible = 1 : Next
      Topper_FlasherDragon.ReflectionEnabled = True
      Topper_FlasherLeft.ReflectionEnabled = True
      Topper_FlasherRight.ReflectionEnabled = True
  End Select
End Sub

Sub SetRenderProbes(Opt)
  On Error Resume Next
    Select Case Opt
      Case 0:
        DDPlastics1.RefractionProbe = ""
        DDPlastics2.RefractionProbe = ""
        DDPlasticsClear.RefractionProbe = ""
      Case 1:
        DDPlastics1.RefractionProbe = "Plastics1"
        DDPlastics2.RefractionProbe = "Plastics2"
        DDPlasticsClear.RefractionProbe = "PlasticsClear"
    End Select
  On Error Goto 0
End Sub

Sub SetupInstructionCards
  If InstMod = 0 Then
    DNDInstruction.image = "DND_InstructionCards"
  End If
  If InstMod = 1 Then
    DNDInstruction.image = "DND_InstructionCardsCustom1"
  End If
  If InstMod = 2 Then
    DNDInstruction.image = "DND_InstructionCardsCustom2"
  End If
End Sub

Sub SetupBlades
  If SideBladeMod = 0 Then
    PinCab_Blades.image = "PinCab_Blades_New"
  End If
  If SideBladeMod = 1 Then
    PinCab_Blades.image = "PinCab_Blades_Custom1"
  End if
End Sub

Sub SetupLegendsCab
  If Table1.ShowDT = False And Table1.ShowFSS = False And RenderingMode <> 2 And LegendsCabMod = 1 Then
    Flasherbloom1.Height = 500
    Flasherbloom2.Height = 500
    Flasherbloom3.Height = 500
  Else
    Flasherbloom1.Height = 250
    Flasherbloom2.Height = 250
    Flasherbloom3.Height = 250
  End If
End Sub

Sub SetupGlass
  If GlassMod = 1 Then
    PinCab_Glass_scratches.Visible = True
  Else
    PinCab_Glass_scratches.Visible = False
  End If
End Sub

Sub SetupRoom
  If RenderingMode = 2 or Table1.ShowFSS = True or VRTest = 1 Then
    VRMode = True
    PinCab_Rails.Visible = 1
    For Each VR_Obj in VRCabinet:VR_Obj.Visible = 1:Next
    If VRRoomChoice = 1 Then
      For Each VR_Obj in VRMinimalRoom : VR_Obj.Visible = 0 : Next
      For Each VR_Obj in VRMegaRoom : VR_Obj.Visible = 1 : Next
      Room360.Visible = 0
      VR_Logo.Visible = False
      VR_Room_Poster.Visible = False
    End If
    If VRRoomChoice = 2 Then
      For Each VR_Obj in VRMinimalRoom : VR_Obj.Visible = 1 : Next
      For Each VR_Obj in VRMegaRoom : VR_Obj.Visible = 0 : Next
      Room360.Visible = 0
      If LogoPosterMod = 1 Then
        VR_Logo.Visible = True
        VR_Room_Poster.Visible = True
      Else
        VR_Logo.Visible = False
        VR_Room_Poster.Visible = False
      End If
    End If
    If VRRoomChoice = 3 Then
      For Each VR_Obj in VRMinimalRoom : VR_Obj.Visible = 0 : Next
      For Each VR_Obj in VRMegaRoom : VR_Obj.Visible = 0 : Next
      Room360.Visible = 1
      VR_Logo.Visible = False
      VR_Room_Poster.Visible = False
    End If
  End If
End Sub

Sub SetupTopper
  If RenderingMode = 2 or Table1.ShowFSS = True or VRTest = 1  or BackglassReflOpt = 1Then
    If TopperMod = 0 Then
      Topper_Backbox.image = "Topper"
    End If
    If TopperMod = 1 Then
      Topper_Backbox.image = "Topper_Custom"
    End If
  End If
End Sub

Sub SetupBackglass
  If BackglassMod = 0 Then
    Pincab_Backglass.image = "Backglass"
  End If
  If BackglassMod = 1 Then
    Pincab_Backglass.image = "BackglassCustom"
  End If
End Sub

Sub SetupLogoPoster
  If RenderingMode = 2 or Table1.ShowFSS = True or VRTest = 1 Then
    If VRRoomChoice = 2 Then
      If LogoPosterMod = 1 Then
        VR_Logo.Visible = True
        VR_Room_Poster.Visible = True
      Else
        VR_Logo.Visible = False
        VR_Room_Poster.Visible = False
      End If
    Else
      VR_Logo.Visible = False
      VR_Room_Poster.Visible = False
    End If
  End If
End Sub

' Thalamus : Exit in a clean and proper way
Sub Table1_exit
  Controller.Pause = False
  Controller.Stop
End Sub

''Vector (Bally 1982)
'https://www.ipdb.org/showpic.pl?id=2723


'*********************************************************************************************************************************
' === TABLE OF CONTENTS  ===
'
' You can quickly jump to a section by searching the four letter tag (ZXXX)
'
' ZVAR: Constants and Global Variable
' ZTIM: Timers
' ZINI: Table Initialization and Exiting
'   ZOPT: User Options
'   ZMAT: General Math Functions
' ZANI: Misc Animations
'   ZBBR: Ball Brightness
'   ZRBR: Room Brightness
' ZGIC: GI Colors
' ZKEY: Key Press Handling
' ZSOL: Solenoids & Flashers
' ZDRN: Drain, Trough, and Ball Release
' ZFLP: Flippers
' ZSLG: Slingshot Animations
' ZSSC: Slingshot Corrections
' ZSWI: Switches
' ZVUK: VUKs and Kickers
'   ZBRL: Ball Rolling and Drop Sounds
'   ZRRL: Ramp Rolling Sound Effects
'   ZFLE: Fleep Mechanical Sounds
' ZPHY: General Advice on Physics
' ZNFF: Flipper Corrections
'   ZDMP: Rubber Dampeners
'   ZRDT: Drop Targets
' ZRST: Stand-Up Targets
'   ZBOU: VPW TargetBouncer for targets and posts
'   ZGIU: GI updates
'   Z3DI: 3D Inserts
'   ZFLD: Flupper Domes
'   ZFLB: Flupper Bumpers
' ZSHA: Ambient Ball Shadows
' ZVLM: VLM Arrays
'   ZVRR: VR Room / VR Cabinet
'
'
'*********************************************************************************************************************************


Option Explicit
Randomize
SetLocale 1033

On Error Resume Next
ExecuteGlobal GetTextFile("controller.vbs")
If Err Then MsgBox "You need the controller.vbs in order to run this table, available in the vp10 package"
On Error Goto 0


'******************************************************
'  ZVAR: Constants and Global Variables
'******************************************************

const SetDIPSwitches= 0   'If you want to set the dips differently, set to 1, and then hit F6 to launch.

'--------------------------------Everything below this is touched at your own risk--------------------------------------------------

Const BallSize = 50   'Ball size must be 50
Const BallMass = 1    'Ball mass must be 1
Const tnob = 3      'Total number of balls on the playfield including captive balls.
Const lob = 0     'Total number of locked balls
Dim gilvl

Dim tablewidth: tablewidth = Table1.width
Dim tableheight: tableheight = Table1.height

'  Standard definitions
Const cGameName = "vector"    'PinMAME ROM name
Const UseSolenoids = 0      '1 = Normal Flippers, 2 = Fastflips
Const UseLamps = 1        '0 = Custom lamp handling, 1 = Built-in VPX handling (using light number in light timer)
Const UseSync = 0
Const HandleMech = 0
Const SSolenoidOn = ""      'Sound sample used for this, obsolete.
Const SSolenoidOff = ""     ' ^
Const SFlipperOn = ""     ' ^
Const SFlipperOff = ""      ' ^
Const SCoin = ""        ' ^

Dim VRMode, VRTest, cab_mode, DesktopMode: DesktopMode = Table1.ShowDT
DesktopMode = Table1.ShowDT
VRTest = 0
If RenderingMode = 2 or VRTest = 1 Then VRMode=1 Else VRMode=0      'VRRoom set based on RenderingMode starting in version 10.72
if Not DesktopMode and VRMode=0 Then cab_mode=1 Else cab_mode=0

Const UseVPMModSol = 0    'Set to 2 for PWM flashers, inserts, and GI. Requires VPinMame 3.6

'NOTES on UseVPMModSol = 2:
'  - Only supported for S9/S11/DataEast/WPC/Capcom/Whitestar (Sega & Stern)/SAM
'  - All lights on the table must have their Fader model set tp "LED (None)" to get the correct fading effects
'  - When not supported VPM outputs only 0 or 1. Therefore, use VPX "Incandescent" fader for lights

LoadVPM "03060000", "Bally.VBS", 3.6  'The "03060000" argument forces user to have VPinMame 3.6


'*******************************************
' ZTIM: Timers
'*******************************************

'The FrameTimer interval should be -1, so executes at the display frame rate
'The frame timer should be used to update anything visual, like some animations, shadows, etc.
'However, a lot of animations will be handled in their respective _animate subroutines.

Dim FrameTime, InitFrameTime
InitFrameTime = 0

FrameTimer.Interval = -1
Sub FrameTimer_Timer()
  FrameTime = gametime - InitFrameTime 'Calculate FrameTime as some animuations could use this
  InitFrameTime = gametime  'Count frametime
  'Add animation stuff here
  BSUpdate
  UpdateBallBrightness
  RollingUpdate       'Update rolling sounds
  DoSTAnim          'Standup target animations
  UpdateStandupTargets
  DoDTAnim          'Drop target animations
  UpdateDropTargets

    If VRMode = 1 Then
        VRDisplayTimer
    Else
        DTDisplayTimer
    End If
End Sub

'The CorTimer interval should be 10. It's sole purpose is to update the Cor (physics) calculations
CorTimer.Interval = 10
Sub CorTimer_Timer(): Cor.Update: End Sub

'DispTimer


'******************************************************
'  ZINI: Table Initialization and Exiting
'******************************************************
Dim VBall1, VBall2, VBall3, gBOT, bsSaucerL, bsSaucerR, bsSaucer1, bsSaucer2, bsSaucer3

gBot = GetBalls

Sub Table1_Init
  vpminit me
  With Controller
    .GameName = cGameName
    If Err Then MsgBox "Can't start Game " & cGameName & vbNewLine & Err.Description:Exit Sub
    .SplashInfoLine = "Vector (Bally 1982)"
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

  'Map all lamps to the corresponding ROM output using the value of TimerInterval of each light object
  vpmMapLights AllLamps     'Make a collection called "AllLamps" and put all the light objects in it.

  'Trough - Creates a ball in the kicker switch and gives that ball used an individual name.
  Set VBall1 = BallRelease.CreateSizedballWithMass(Ballsize/2,Ballmass)
  Set VBall2 = BallRelease001.CreateSizedballWithMass(Ballsize/2,Ballmass)
  Set VBall3 = BallRelease002.CreateSizedballWithMass(Ballsize/2,Ballmass)
  gBOT = Array(VBall1,VBall2,VBall3)

  'Forces the trough switches to "on" at table boot so the game logic knows there are balls in the trough.
  Controller.Switch(3) = 1
  Controller.Switch(2) = 1
  Controller.Switch(1) = 1

  InitDigits
  VRInitDigits
  setup_backglass

  Dim x
  For each x in InsertVLM
    x.opacity = 75
  Next
  For each x in DigitVLM
    x.opacity = 65
  Next
' For each x in GIVLM
'   x.opacity = 150
' Next

' If user wants to set dip switches themselves it will force them to set it via F6.
  If SetDIPSwitches = 0 Then
    SetDefaultDips
  End If

'>>>>>>>>>>>>>>Kickers<<<<<<<<<<<<<<<<<<

    ' Saucer (Bottom Left)
  Set bsSaucerL = New cvpmBallStack
  with bsSaucerL
    .InitSaucer SaucerL,5,78,22
    .KickForceVar = 3
    .KickAngleVar = 3
'   .InitExitSnd SoundFX("fx_ballrel",DOFContactors), Soundfx("HoleKick",DOFContactors)
  end with

    ' Saucer (Bottom Right)
  Set bsSaucerR = New cvpmBallStack
  with bsSaucerR
    .InitSaucer SaucerR,4,-87,22
    .KickForceVar = 3
    .KickAngleVar = 3
'     .InitExitSnd SoundFX("fx_ballrel",DOFContactors), Soundfx("HoleKick",DOFContactors)
  end with

    ' Saucer (UpperPF Bottom)
  Set bsSaucer1 = New cvpmBallStack
  with bsSaucer1
    .InitSaucer Saucer1,17,178,15
    .KickForceVar = 3
    .KickAngleVar = 3
'     .InitExitSnd SoundFX("fx_ballrel",DOFContactors), Soundfx("HoleKick",DOFContactors)
  end with

    ' Saucer (UpperPF Middle)
  Set bsSaucer2 = New cvpmBallStack
  with bsSaucer2
    .InitSaucer Saucer2,18,180,20
    .KickForceVar = 3
    .KickAngleVar = 3
'     .InitExitSnd SoundFX("fx_ballrel",DOFContactors), Soundfx("HoleKick",DOFContactors)
  end with

    ' Saucer (UpperPF Top)
  Set bsSaucer3 = New cvpmBallStack
  with bsSaucer3
    .InitSaucer Saucer3,19,340,40
    .KickForceVar = 3
    .KickAngleVar = 3
    .InitAltKick 160,13
'     .InitExitSnd SoundFX("fx_ballrel",DOFContactors), Soundfx("HoleKick",DOFContactors)
  end with

  For each x in BP_SaucerLWU
    x.visible = 0
  Next
  For each x in BP_SaucerRWU
    x.visible = 0
  Next
  For each x in BP_UpperSaucer1WU
    x.visible = 0
  Next
  For each x in BP_UpperSaucer2WU
    x.visible = 0
  Next
  For each x in BP_UpperSaucer3DWU
    x.visible = 0
  Next

'>>>>>>>>>>>>>>Nudging<<<<<<<<<<<<<<<<<<
  'Nudging
  vpmNudge.TiltSwitch=15 'Tilt is switch 15
  vpmNudge.Sensitivity=4
  vpmNudge.TiltObj = Array(Bumper1, LeftSlingshot, RightSlingshot)

  'Main Timer init
  PinMAMETimer.Interval = PinMAMEInterval
  PinMAMETimer.Enabled = 1

End Sub

Sub dtUpper(enabled)
  if enabled then
'   PlaySoundAt SoundFX(DTResetSound,DOFContactors),sw27a
    DTRaise 46
    DTRaise 47
    DTRaise 48
  end if
End Sub

Sub dtL(enabled)
  if enabled then
'   PlaySoundAt SoundFX(DTResetSound,DOFContactors),sw30a
    DTRaise 26
    DTRaise 27
    DTRaise 28
  end if
End Sub

Sub dtLL(enabled)
  if enabled then
'   PlaySoundAt SoundFX(DTResetSound,DOFContactors),sw47a
    DTRaise 29
    DTRaise 30
    DTRaise 31
  end if
End Sub

Sub Table1_Paused:Controller.Pause = 1:End Sub
Sub Table1_unPaused:Controller.Pause = 0:End Sub

Sub Table1_exit()
  Controller.Pause = False
  Controller.Stop
End Sub

'***********************************************************************************
'****                 Switch reference                        ****
'***********************************************************************************
Const swBallStack1          = 1
Const swBallStack2          = 2
Const swOuthole         = 3
Const swRightBottomSaucer     = 4
Const swLeftBottomSaucer      = 5
Const swShooterLane       = 12
Const swTopRightLane        = 13
Const swVectorScanGate        = 14
Const swTiltMe            = 15
Const swUpperSaucer1        = 17
Const swUpperSaucer2        = 18
Const swUpperSaucer3        = 19
Const swTargetE         = 21
Const swTargetP         = 22
Const swTargetY         = 23
Const swTargetH         = 24
Const swLowerDropTargetRight    = 26
Const swLowerDropTargetMiddle   = 27
Const swLowerDropTargetLeft   = 28
Const swUpperDropTargetRight    = 29
Const swUpperDropTargetMiddle   = 30
Const swUpperDropTargetLeft   = 31
Const swRightOutLane        = 33
Const swRightReturnLane     = 34
Const swLeftOutLane       = 35
Const swLeftReturnLane        = 36
Const swAdvanceBonusRollOver    = 37
Const swSlingRight          = 38
Const swSlingLeft         = 39
Const swBumper            = 40
Const swTopDropTargetZ        = 46
Const swTopDropTargetY        = 47
Const swTopDropTargetX        = 48

'***********************************************************************************
'****                   DOF reference                     ****
'***********************************************************************************
Const dShooterLane          = 200
Const dBallRelease          = 201
Const dUpperSaucer3Up       = 202
Const dUpperSaucer3Down     = 203
Const dUpperSaucer2       = 204
Const dUpperSaucer1       = 205
Const dDropUpperLeft        = 206
Const dDropUpperMiddle        = 207
Const dDropUpperRight       = 208
Const dDropLowerLeft        = 209
Const dDropLowerMiddle        = 210
Const dDropLowerRight       = 211


'*******************************************
'  ZOPT: User Options
'*******************************************

' Called when options are tweaked by the player.
' - 0: game has started, good time to load options and adjust accordingly
' - 1: an option has changed
' - 2: options have been reseted
' - 3: player closed the tweak UI, good time to update staticly prerendered parts
' Table1.Option arguments are:
' - option name, minimum value, maximum value, step between valid values, default value, unit (0=None, 1=Percent), an optional arry of literal strings

Dim dspTriggered : dspTriggered = False
Sub Table1_OptionEvent(ByVal eventId)
  Dim v
    If eventId = 1 And Not dspTriggered Then dspTriggered = True : DisableStaticPreRendering = True : End If

  ' Room brightness
  LightLevel = NightDay/100
  SetRoomBrightness LightLevel   'Uncomment this line for lightmapped tables.

  '-----Cabinet rails-----
  v = Table1.Option("Cabinet Rails", 0, 1, 1, 1, 0, Array("Hide", "Show"))
  VR_Cab_Siderails.visible = v
  VR_Cab_Lockdownbar.visible = v
  if VRMode = 0 and cab_mode = 1 Then
    VR_Cab_Siderails.visible = 0
    VR_Cab_Lockdownbar.visible = 0
  End If


  '-----GI Lamps Color-----
  v = Table1.Option("GI Lights", 0, 2, 1, 0, 0, Array("Classic", "White", "Purple"))
  SetGIColor v

  '-----GI Lamps Brightness-----
  v = Table1.Option("GI Brightness", 0, 1, .1, .6, 1)
  GIBrightness v

  ' Playfield Reflections
  v = Table1.Option("Playfield Reflections", 0, 2, 1, 1, 0, Array("Off", "Clean", "Rough"))
  Select Case v
    Case 0: playfield_mesh.ReflectionProbe = "": BM_Playfield.ReflectionProbe = "": BM_Playfield_001.ReflectionProbe = "": BM_UpperPF.ReflectionProbe = ""
    Case 1: playfield_mesh.ReflectionProbe = "Playfield Reflections": BM_Playfield.ReflectionProbe = "Playfield Reflections": BM_Playfield_001.ReflectionProbe = "Playfield Reflections": BM_UpperPF.ReflectionProbe = "Playfield Reflections"
    Case 2: playfield_mesh.ReflectionProbe = "Playfield Reflections Rough": BM_Playfield.ReflectionProbe = "Playfield Reflections Rough": BM_Playfield_001.ReflectionProbe = "Playfield Reflections Rough": BM_UpperPF.ReflectionProbe = "Playfield Reflections Rough"
  End Select

  ' Color Saturation
    ColorLUT = Table1.Option("Color Saturation", 1, 12, 1, 1, 0, _
    Array("Normal", "Dark Contrast", "Desaturated 10%", "Desaturated 20%", "Desaturated 30%", "Desaturated 40%", "Desaturated 50%", _
        "Desaturated 60%", "Desaturated 70%", "Desaturated 80%", "Desaturated 90%", "Black 'n White"))
  if ColorLUT = 1 Then Table1.ColorGradeImage = "ColorGradeLUT256x16_1to1"
  if ColorLUT = 2 Then Table1.ColorGradeImage = "LUT4"
  if ColorLUT = 3 Then Table1.ColorGradeImage = "colorgradelut256x16-10"
  if ColorLUT = 4 Then Table1.ColorGradeImage = "colorgradelut256x16-20"
  if ColorLUT = 5 Then Table1.ColorGradeImage = "colorgradelut256x16-30"
  if ColorLUT = 6 Then Table1.ColorGradeImage = "colorgradelut256x16-40"
  if ColorLUT = 7 Then Table1.ColorGradeImage = "colorgradelut256x16-50"
  if ColorLUT = 8 Then Table1.ColorGradeImage = "colorgradelut256x16-60"
  if ColorLUT = 9 Then Table1.ColorGradeImage = "colorgradelut256x16-70"
  if ColorLUT = 10 Then Table1.ColorGradeImage = "colorgradelut256x16-80"
  if ColorLUT = 11 Then Table1.ColorGradeImage = "colorgradelut256x16-90"
  if ColorLUT = 12 Then Table1.ColorGradeImage = "colorgradelut256x16-100"


    If eventId = 3 And dspTriggered Then dspTriggered = False : DisableStaticPreRendering = False : End If
End Sub

Dim LightLevel : LightLevel = 0.25        ' Level of room lighting (0 to 1), where 0 is dark and 100 is brightest
Dim ColorLUT : ColorLUT = 1           ' Color desaturation LUTs: 1 to 11, where 1 is normal and 11 is black'n'white
Dim VolumeDial : VolumeDial = 0.8             ' Overall Mechanical sound effect volume. Recommended values should be no greater than 1.
Dim BallRollVolume : BallRollVolume = 0.5     ' Level of ball rolling volume. Value between 0 and 1
Dim RampRollVolume : RampRollVolume = 0.5     ' Level of ramp rolling volume. Value between 0 and 1


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
'  ZANI: Misc Animations
'******************************************************

'-------------------Flippers-------------------
Sub LeftFlipper_Animate
  Dim a : a = LeftFlipper.CurrentAngle
  FlipperLSh.RotZ = a

  Dim v, BP
  v = 255.0 * (124.0 - LeftFlipper.CurrentAngle) / (124.0 -  69.0)

  For each BP in BP_LeftFlip
    BP.Rotz = a
    BP.visible = v < 128.0
  Next
  For each BP in BP_LeftFlipU
    BP.Rotz = a
    BP.visible = v >= 128.0
  Next
End Sub

Sub RightFlipper_Animate
  Dim a : a = RightFlipper.CurrentAngle
  FlipperRSh.RotZ = a

  Dim v, BP
  v = 255.0 * (-124.0 - RightFlipper.CurrentAngle) / (-124.0 +  69.0)

  For each BP in BP_RightFlip
    BP.Rotz = a
    BP.visible = v < 128.0
  Next
  For each BP in BP_RightFlipU
    BP.Rotz = a
    BP.visible = v >= 128.0
  Next
End Sub

Sub RightFlipper1_Animate
  Dim a : a = RightFlipper1.CurrentAngle
' FlipperRSh001.RotZ = a

  Dim v, BP
  v = 255.0 * (-127.0 - RightFlipper1.CurrentAngle) / (-127.0 +  67.0)

  For each BP in BP_RightFlip1
    BP.Rotz = a
    BP.visible = v < 128.0
  Next
  For each BP in BP_RightFlip1U
    BP.Rotz = a
    BP.visible = v >= 128.0
  Next
End Sub

Sub LeftFlipper1_Animate
  Dim a : a = LeftFlipper1.CurrentAngle
' FlipperRSh001.RotZ = a

  Dim v, BP
  v = 255.0 * (156.0 - LeftFlipper1.CurrentAngle) / (156.0 -  95.0)

  For each BP in BP_LeftFlip1
    BP.Rotz = a
    BP.visible = v < 128.0
  Next
  For each BP in BP_LeftFlip1U
    BP.Rotz = a
    BP.visible = v >= 128.0
  Next
End Sub

'-------------------Gates-------------------
Sub Gate5_Animate
  Dim spinangle:spinangle = Gate5.currentangle
  Dim BL : For Each BL in BP_Prim5 : BL.RotX = spinangle: Next
End Sub

Sub Gate5Sound_Hit
  SoundPlayfieldGate
End Sub

Sub Gate4_Animate
  Dim spinangle:spinangle = Gate4.currentangle
  Dim BL : For Each BL in BP_Prim4 : BL.RotX = spinangle: Next
End Sub

Sub Gate4Sound_Hit
  SoundPlayfieldGate
End Sub

Sub Gate2_Animate
  Dim spinangle:spinangle = Gate2.currentangle
  Dim BL : For Each BL in BP_Prim2 : BL.RotX = spinangle: Next
End Sub

Sub Gate2Sound_Hit
  SoundPlayfieldGate
End Sub

Sub Gate3_Animate
  Dim spinangle:spinangle = Gate3.currentangle
  Dim BL : For Each BL in BP_Prim3 : BL.RotX = spinangle: Next
End Sub

Sub Gate3Sound_Hit
  SoundPlayfieldGate
End Sub

Sub Gate6_Animate
  Dim spinangle:spinangle = Gate6.currentangle
  Dim BL : For Each BL in BP_Prim6 : BL.RotX = spinangle: Next
End Sub

Sub sw14_Animate
  Dim spinangle:spinangle = sw14.currentangle
  Dim BL : For Each BL in BP_Prim14 : BL.RotX = spinangle: Next
End Sub


'******************************************************
'   ZBBR: BALL BRIGHTNESS
'******************************************************

Const BallBrightness =  .8        'Ball brightness - Value between 0 and 1 (0=Dark ... 1=Bright)

' Constants for plunger lane ball darkening.
' You can make a temporary wall in the plunger lane area and use the co-ordinates from the corner control points.
Const PLOffset = 0.5      'Minimum ball brightness scale in plunger lane
Const PLLeft = 866      'X position of punger lane left
Const PLRight = 930       'X position of punger lane right
Const PLTop = 1297        'Y position of punger lane top
Const PLBottom = 1810       'Y position of punger lane bottom
Dim PLGain: PLGain = (1-PLOffset)/(PLTop-PLBottom)

Sub UpdateBallBrightness
  Dim s, b_base, b_r, b_g, b_b, d_w
  b_base = 120 * BallBrightness + 135*gilvl ' orig was 120 and 70

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


'****************************
' ZGIC: GI Colors
'****************************

''''''''' GI Color Change options
' reference: https://andi-siess.de/rgb-to-color-temperature/

Dim GIColorPick

Dim c2700k: c2700k = rgb(255, 169, 87)
Dim c3000k: c3000k = rgb(255, 180, 107)
Dim c4000k: c4000k = rgb(255, 209, 163)

Dim cRedFull: cRedFull = rgb(255,0,0)
Dim cRed: cRed= rgb(255,5,5)
Dim cPinkFull: cPinkFull = rgb(255,0,225)
Dim cPink: cPink = rgb(255,5,255)
Dim cWhiteFull: cWhiteFull = rgb(255,255,128)
Dim cWhite: cWhite = rgb(255,255,255)
Dim cBlueFull: cBlueFull= rgb(0,0,255)
Dim cBlue: cBlue = rgb(5,5,255)
Dim cCyanFull: cCyanFull= rgb(0,255,255)
Dim cCyan : cCyan = rgb(5,128,255)
Dim cYellowFull:cYellowFull  = rgb(255,255,128)
Dim cYellow: cYellow = rgb(255,255,0)
Dim cOrangeFull: cOrangeFull = rgb(255,128,0)
Dim cOrange: cOrange = rgb(255,70,5)
Dim cGreenFull: cGreenFull = rgb(0,255,0)
Dim cGreen: cGreen = rgb(5,255,5)
Dim cPurpleFull: cPurpleFull = rgb(128,0,255)
Dim cPurple: cPurple = rgb(60,5,255)
Dim cAmberFull: cAmberFull = rgb(255,197,143)
Dim cAmber: cAmber = rgb(255,197,143)
Dim cBlack: cBlack = rgb(0,0,0)

Dim cArray, sArray
cArray = Array(c2700k, cWhite, cPurple)

Sub SetGIColor(c)
  Dim xx, BL
  GIColorPick = cArray(c)
  For each xx in cGI: xx.color = cArray(c): xx.colorfull = cArray(c): Next
  For each BL in GIvlm: BL.color = cArray(c): Next
End Sub

Sub GIBrightness(lvl)
  dim new_base
  If lvl > 1 Then lvl = 1
  If lvl < 0 Then lvl = 0

  ' Lighting level
  Dim v: v=(lvl * 245 + 10)/255

  Dim red, green, blue

  red = GIColorPick And &HFF
  green = (GIColorPick \ &H100) And &HFF
  blue = (GIColorPick \ &H10000) And &HFF

  new_base = RGB(red*v, green*v, blue*v)

  Dim xx, BL
  For each xx in cGI: xx.color = new_base: xx.colorfull = new_base: Next
  For each BL in GIvlm: BL.opacity = lvl * 150: Next

End Sub

'******************************************************
' ZKEY: Key Press Handling
'******************************************************

Dim BIPL  : BIPL=0

Sub Table1_KeyDown(ByVal keycode) '***What to do when a button is pressed***
  'If keycode = PlungerKey Then Plunger.Pullback:vpmTimer.PulseSw 31
  If KeyCode = PlungerKey Then
    Plunger.PullBack
    Plunger.Pullback:SoundPlungerPull()
    TimerVRPlunger.Enabled = True
    TimerVRPlunger1.Enabled = False
    VR_Plunger.Y = 2098
  End If
    If Keycode = LeftFlipperKey Then
    VR_LeftFlipperButton.X = VR_LeftFlipperButton.x +8
  End if
  If Keycode = RightFlipperKey Then
    VR_RightFlipperButton.X = VR_RightFlipperButton.x -8
  End if
  If Keycode = StartGameKey Then
    SoundStartButton
    VR_StartButton.y=VR_StartButton.y-4
  End if
  If keycode = LeftTiltKey Then Nudge 90, 1 : SoundNudgeLeft      ' Sets the nudge angle and power
  If keycode = RightTiltKey Then Nudge 270, 1 : SoundNudgeRight   ' ^
  If keycode = CenterTiltKey Then Nudge 0, 1 : SoundNudgeCenter   ' ^
  If keycode = AddCreditKey or keycode = AddCreditKey2 Then
    Select Case Int(rnd*3)
      Case 0: PlaySound ("Coin_In_1"), 0, CoinSoundLevel, 0, 0.25
      Case 1: PlaySound ("Coin_In_2"), 0, CoinSoundLevel, 0, 0.25
      Case 2: PlaySound ("Coin_In_3"), 0, CoinSoundLevel, 0, 0.25
    End Select
  End If
  If vpmKeyDown(keycode) Then Exit Sub
End Sub

Sub Table1_KeyUp(ByVal keycode)   '***What to do when a button is released***

  'If Keycode = StartGameKey Then Controller.Switch(16) = 0
  If keycode = PlungerKey Then
    Plunger.Fire
    If BIPL = 1 Then
      SoundPlungerReleaseBall()     'Plunger release sound when there is a ball in shooter lane
    Else
      SoundPlungerReleaseNoBall()     'Plunger release sound when there is no ball in shooter lane
    End If
    TimerVRPlunger.Enabled = False
        TimerVRPlunger1.Enabled = True
    VR_Plunger.Y = 2098
  End If
  If Keycode = LeftFlipperKey Then
    VR_LeftFlipperButton.x = VR_LeftFlipperButton.x -8
  End if
  If keycode = RightFlipperKey Then
    VR_RightFlipperButton.X = VR_RightFlipperButton.x +8
  End If
  If Keycode = StartGameKey Then
    SoundStartButton
    VR_StartButton.y=VR_StartButton.y+4
  End if
  If vpmKeyUp(keycode) Then Exit Sub
End Sub


' VR Plunger code

Sub TimerVRPlunger_Timer
  If VR_Plunger.Y < 2208 then
       VR_Plunger.Y = VR_Plunger.Y + 5
  End If
End Sub

Sub TimerVRPlunger1_Timer
  VR_Plunger.Y = 2098 + (5* Plunger.Position) -20
End Sub


'******************************************************
' ZSOL: Solenoids & Flashers
'******************************************************

SolCallback(1)          = "dtUpper"                   ' Top Bank       Target Reset    (1)
SolCallback(2)          = "dtL"                   ' Bot Bank Low   Target Reset    (7)
SolCallback(4)          = "dtLL"                    ' Top Bank Upper Target Reset    (6)
SolCallback(3)        = "DoSaucerL"               ' Saucer (Lower Left)       Release
SolCallback(5)        = "DoSaucerR"               ' Saucer (Lower Right)      Release
SolCallback(6)            = "SolKnocker"                    ' Knocker     (4)
SolCallback(7)          = "EjectToShooter"                ' Outhole Kicker (5)  (Ball Release)
SolCallback(9)        = "DoUpperSaucer3Up"            ' Saucer (2nd Floor Top)    Release
SolCallback(10)       = "DoUpperSaucer3Down"            ' Saucer (2nd Floor Top)    Release
SolCallback(11)       = "DoUpperSaucer2"              ' Saucer (2nd Floor Mid)    Release
SolCallback(12)       = "DoUpperSaucer1"              ' Saucer (2nd Floor Bottom) Release
SolCallback(18)         = ""                        ' Coin Box (lockout) ()
SolCallback(19)         = "RelayAC"                 ' K1 Relay (Flipper Enable)()

' These are reassigned in the Solenoid Timer Handler via Lamp 63 as a Selection Switch in VPM
SolCallback(31)       = "Drop1"                 ' Lower Top Drop Target Down 1
SolCallback(32)       = "Drop2"                 ' Lower Top Drop Target Down 2
SolCallback(33)       = "Drop3"                 ' Lower Top Drop Target Down 3
SolCallback(34)       = "Drop4"                 ' Lower Bot Drop Target Down 1
SolCallback(35)       = "Drop5"                 ' Lower Bot Drop Target Down 2
SolCallback(36)       = "Drop6"                 ' Lower Bot Drop Target Down 3

SolCallback(sLRFlipper) = "SolRFlipper"
SolCallback(sLLFlipper) = "SolLFlipper"


Sub SolKnocker(Enabled)
  If enabled Then
    KnockerSolenoid
  End If
End Sub




'**********************************************************************************************************


 Sub Drop1(enabled) : If enabled then : DTDrop 31 : DOF dDropUpperLeft, 2 : end if : End Sub
 Sub Drop2(enabled) : If enabled then : DTDrop 30 : DOF dDropUpperMiddle, 2 : end if : End Sub
 Sub Drop3(enabled) : If enabled then : DTDrop 29 : DOF dDropUpperRight, 2 : end if : End Sub
 Sub Drop4(enabled) : If enabled then : DTDrop 28 : DOF dDropLowerLeft, 2 : end if : End Sub
 Sub Drop5(enabled) : If enabled then : DTDrop 27 : DOF dDropLowerMiddle, 2 : end if : End Sub
 Sub Drop6(enabled) : If enabled then : DTDrop 26 : DOF dDropLowerRight, 2 : end if : End Sub

 Sub Solenoid_Timer()
  Dim Changed, Count, funcName, ii, sel, solNo
  Changed = Controller.ChangedSolenoids
  If Not IsEmpty(Changed) Then
    sel = Controller.Lamp(63)
    Count = UBound(Changed, 1)
    For ii = 0 To Count
      solNo = Changed(ii, CHGNO)
      If SolNo >= 7 And SolNo <= 12 And sel Then solNo = solNo +24
      funcName = SolCallback(solNo)
      If funcName <> "" Then Execute funcName & " CBool(" & Changed(ii, CHGSTATE) &")"
    Next
  End If
End Sub

 'Handle Solenoid Events
Sub SolTimer_Timer()
  Dim ChgSol, tmp, ii, CBoard, solnum
  ChgSol  = Controller.ChangedSolenoids
  If Not IsEmpty(ChgSol) Then
  CBoard = Controller.Lamp(63)
    For ii = 0 To UBound(ChgSol)
      solnum = ChgSol(ii, 0)
      If solnum <= 12 and CBoard Then solnum = solnum + 24
      tmp = Solcallback(solnum)
      If tmp <> "" Then Execute tmp & vpmTrueFalse(ChgSol(ii, 1)+1)
    Next
  End If
End Sub

' Tie In Nudge to AC Relay
Sub RelayAC(enabled)
  vpmNudge.SolGameOn enabled
  vpmFlips.TiltSol enabled
End Sub

'******************************************************
' ZDRN: Drain, Trough, and Ball Release
'******************************************************

 ' Drain and Kickers
SaucerRWU.IsDropped = 1
SaucerLWU.IsDropped = 1
UpperSaucer1WU.IsDropped = 1
UpperSaucer2WU.IsDropped = 1
UpperSaucer3DWU.IsDropped = 1
UpperSaucer3UWU.IsDropped = 1

Sub BallRelease_UnHit : UpdateTrough : End Sub
Sub BallRelease_Hit   : UpdateTrough : End Sub
Sub BallRelease001_UnHit : Controller.Switch(2) = 0 :UpdateTrough : End Sub
Sub BallRelease001_Hit   : Controller.Switch(2) = 1 :UpdateTrough : End Sub
Sub BallRelease002_UnHit : Controller.Switch(1) = 0 :UpdateTrough : End Sub
Sub BallRelease002_Hit   : Controller.Switch(1) = 1 :UpdateTrough : End Sub
Sub Drain_UnHit   : UpdateTrough : End Sub

Sub EjectToShooter(Enabled)
  If Enabled = true Then
    RandomSoundBallRelease BallRelease
    BallRelease.kick 60, 10
    UpdateTrough
  End If
End Sub

Sub Drain_Hit()
  RandomSoundDrain Drain
  Drain.kick 75, 30
  UpdateTrough
End Sub

Sub UpdateTrough()
  UpdateTroughTimer.Interval = 300
  UpdateTroughTimer.Enabled = 1
End Sub
Sub UpdateTroughTimer_Timer
  If BallRelease.BallCntOver = 0 Then BallRelease001.kick 60, 10
  If BallRelease001.BallCntOver = 0 Then BallRelease002.kick 50, 10
  UpdateTroughTimer.Enabled = 0
End Sub

Sub Saucer1_Hit:bsSaucer1.addball me : SoundSaucerLock : End Sub
Sub Saucer1_UnHit: SoundSaucerKick 1, Saucer1: End Sub

Sub Saucer2_Hit:bsSaucer2.addball me : SoundSaucerLock : End Sub
Sub Saucer2_UnHit: SoundSaucerKick 1, Saucer2: End Sub

Sub Saucer3_Hit:bsSaucer3.addball me : SoundSaucerLock : End Sub
Sub Saucer3_UnHit: SoundSaucerKick 1, Saucer3: End Sub

Sub SaucerL_Hit:bsSaucerL.addball me : SoundSaucerLock : End Sub
Sub SaucerL_UnHit: SoundSaucerKick 1, SaucerL: End Sub

Sub SaucerR_Hit:bsSaucerR.addball me : SoundSaucerLock : End Sub
Sub SaucerR_UnHit: SoundSaucerKick 1, SaucerR: End Sub

'--------Lower Right Saucer---------
Sub DoSaucerR(enabled)
  If enabled Then
    bsSaucerR.ExitSol_On
    'SaucerRWU.IsDropped = 0
    dim x
    For each x in BP_SaucerRWU
      x.visible = 1
    Next
    SaucerRWU.TimerEnabled = 1
  End if
End Sub
Sub SaucerRWU_Timer
  'SaucerRWU.IsDropped = 1
  dim x
  For each x in BP_SaucerRWU
    x.visible = 0
  Next
  Me.TimerEnabled = 0
End Sub

'--------Lower Left Saucer---------
Sub DoSaucerL(enabled)
  If enabled Then
    bsSaucerL.ExitSol_On
    'SaucerLWU.IsDropped = 0
    dim x
    For each x in BP_SaucerLWU
      x.visible = 1
    Next
    SaucerLWU.TimerEnabled = 1
  End if
End Sub
Sub SaucerLWU_Timer
  'SaucerLWU.IsDropped = 1
  dim x
  For each x in BP_SaucerLWU
    x.visible = 0
  Next
  Me.TimerEnabled = 0
End Sub

'--------UpperPF Bottom Saucer---------
Sub DoUpperSaucer1(enabled)
  If enabled Then
    bsSaucer1.ExitSol_On
    'UpperSaucer1WU.IsDropped = 0
    dim x
    For each x in BP_UpperSaucer1WU
      x.visible = 1
    Next
    UpperSaucer1WU.TimerEnabled = 1
    DOF dUpperSaucer1, 2
  End if
End Sub
Sub UpperSaucer1WU_Timer
  'UpperSaucer1WU.IsDropped = 1
  dim x
  For each x in BP_UpperSaucer1WU
    x.visible = 0
  Next
  Me.TimerEnabled = 0
End Sub

'--------UpperPF Middle Saucer---------
Sub DoUpperSaucer2(enabled)
  If enabled Then
    bsSaucer2.ExitSol_On
    'UpperSaucer2WU.IsDropped = 0
    dim x
    For each x in BP_UpperSaucer2WU
      x.visible = 1
    Next
    UpperSaucer2WU.TimerEnabled = 1
    DOF dUpperSaucer2, 2
  End if
End Sub
Sub UpperSaucer2WU_Timer
  'UpperSaucer2WU.IsDropped = 1
  dim x
  For each x in BP_UpperSaucer2WU
    x.visible = 0
  Next
  Me.TimerEnabled = 0
End Sub

'--------UpperPF Top Saucer (Fire Down)---------
Sub DoUpperSaucer3Down(enabled)
  If enabled Then
    bsSaucer3.ExitAltSol_On
    'UpperSaucer3DWU.IsDropped = 0
    dim x
    For each x in BP_UpperSaucer3DWU
      x.visible = 1
    Next
    UpperSaucer3DWU.TimerEnabled = 1
    DOF dUpperSaucer3Down, 2
  End if
End Sub
Sub UpperSaucer3DWU_Timer
  'UpperSaucer3DWU.IsDropped = 1
  dim x
  For each x in BP_UpperSaucer3DWU
    x.visible = 0
  Next
  Me.TimerEnabled = 0
End Sub

'--------UpperPF Top Saucer (Fire Up)---------
Sub DoUpperSaucer3Up(enabled)
  If enabled Then
    bsSaucer3.ExitSol_On
    UpperSaucer3UWU.IsDropped = 0
    UpperSaucer3UWU.TimerEnabled = 1
    DOF dUpperSaucer3Up, 2
  End if
End Sub
Sub UpperSaucer3UWU_Timer
  UpperSaucer3UWU.IsDropped = 1
  Me.TimerEnabled = 0
End Sub


'*******************************************
' ZFLP: Flippers
'*******************************************

Const ReflipAngle = 20

Sub SolLFlipper(Enabled)
     If Enabled Then
'   LeftFlipper:LeftFlipper.RotateToEnd:LeftFlipper1.RotateToEnd
    FlipperActivate LeftFlipper, LFPress
    LF.Fire
    FlipperActivate LeftFlipper1, LF1Press
    LF1.Fire

    If Leftflipper.currentangle < Leftflipper.endangle + ReflipAngle Then
      RandomSoundReflipUpLeft LeftFlipper
    Else
      SoundFlipperUpAttackLeft LeftFlipper
      RandomSoundFlipperUpLeft LeftFlipper
    End If
     Else
'   LeftFlipper:LeftFlipper.RotateToStart:LeftFlipper1.RotateToStart
    FlipperDeActivate LeftFlipper, LFPress
    LeftFlipper.RotateToStart
    FlipperDeActivate LeftFlipper1, LF1Press
    LeftFlipper1.RotateToStart

    If LeftFlipper.currentangle < LeftFlipper.startAngle - 5 Then
      RandomSoundFlipperDownLeft LeftFlipper
    End If
     End If
  End Sub

Sub SolRFlipper(Enabled)
     If Enabled Then
'   RightFlipper:RightFlipper.RotateToEnd:RightFlipper1.RotateToEnd
    FlipperActivate RightFlipper, RFPress
    RF.Fire
    FlipperActivate RightFlipper1, RF1Press
    RF1.Fire

    If rightflipper.currentangle > rightflipper.endangle - ReflipAngle Then
      RandomSoundReflipUpRight RightFlipper
    Else
      SoundFlipperUpAttackRight RightFlipper
      RandomSoundFlipperUpRight RightFlipper
    End If
     Else
'   RightFlipper:RightFlipper.RotateToStart:RightFlipper1.RotateToStart
    FlipperDeActivate RightFlipper, RFPress
    RightFlipper.RotateToStart
    FlipperDeActivate RightFlipper1, RF1Press
    RightFlipper1.RotateToStart

    If RightFlipper.currentangle < RightFlipper.startAngle + 5 Then
      RandomSoundFlipperDownRight RightFlipper
    End If
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

Sub LeftFlipper1_Collide(parm)
  CheckLiveCatch ActiveBall, LeftFlipper1, LF1Count, parm
  LF1.ReProcessBalls ActiveBall
  LeftFlipperCollide parm
End Sub

Sub RightFlipper1_Collide(parm)
  CheckLiveCatch ActiveBall, RightFlipper1, RF1Count, parm
  RF1.ReProcessBalls ActiveBall
  RightFlipperCollide parm
End Sub


'************************************************************
' ZSLG: Slingshot Animations
'************************************************************

Dim LStep : LStep = 0 : LeftSlingShot.TimerEnabled = 1
Dim RStep : RStep = 0 : RightSlingShot.TimerEnabled = 1

Sub RightSlingShot_Slingshot
  RS.VelocityCorrect(Activeball)
  vpmtimer.pulsesw swSlingRight
  Rstep = 0
  RightSlingShot_Timer
  Dim bl
  For each bl in BP_RSling1: bl.visible = true: Next
  For each bl in BP_sling1: bl.transY = 20: Next
  RightSlingShot.TimerEnabled = 1
  RightSlingShot.TimerInterval = 10
  RandomSoundSlingshotRight sling1
End Sub

Sub RightSlingShot_Timer
  Dim BL
    Select Case RStep
    Case 1: For Each BL in BP_RSling1 : BL.visible = false : Next : For Each BL in BP_RSling2 : BL.visible = true : Next: For each BL in BP_sling1: bl.transY = 10: Next
        Case 2: For Each BL in BP_RSling2 : BL.visible = false : Next : For Each BL in BP_RSling : BL.visible = true :Next: For each BL in BP_sling1: bl.transY = 0: Me.TimerEnabled = 0: Next
    End Select

    RStep = RStep + 1
End Sub


Sub LeftSlingShot_Slingshot
  LS.VelocityCorrect(ActiveBall)
  vpmtimer.pulsesw 39
  Lstep = 0
  LeftSlingShot_Timer
  Dim bl
  For each bl in BP_LSling1: bl.visible = true: Next
  For each bl in BP_sling2: bl.transY = -20: Next
' sling2.TransY =  - 20             'Sling Metal Bracket
  LeftSlingShot.TimerEnabled = 1
  LeftSlingShot.TimerInterval = 10
  RandomSoundSlingshotLeft sling2
End Sub

Sub LeftSlingShot_Timer
  Dim BL
    Select Case LStep
    Case 1: For Each BL in BP_LSling1 : BL.visible = false : Next : For Each BL in BP_LSling2 : BL.visible = true : Next: For each BL in BP_sling2: bl.transY = -10: Next
        Case 2: For Each BL in BP_LSling2 : BL.visible = false : Next : For Each BL in BP_LSling : BL.visible = true :Next: For each BL in BP_sling2: bl.transY = 0: Me.TimerEnabled = 0: Next
    End Select

    LStep = LStep + 1
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


'************************************************************
' ZSWI: SWITCHES
'************************************************************

'--------------Gate Trigger------------
Sub sw14_Hit()      : vpmTimer.PulseSw(swVectorScanGate) : End Sub

'-------------Wire Triggers-------------
Sub sw12_Hit()      : controller.switch (swShooterLane)=1 :AnimateWire BP_sw12, 1: BIPL=1 : End Sub
Sub sw12_unHit()    : controller.switch (swShooterLane)=0 : AnimateWire BP_sw12, 0:BIPL=0 : End Sub
Sub sw13_Hit()      : controller.switch (swTopRightLane)=1:AnimateWire BP_sw13, 1:End Sub
Sub sw13_unHit()    : controller.switch (swTopRightLane)=0:AnimateWire BP_sw13, 0:End Sub
Sub sw33_Hit()      : controller.switch (swRightOutLane)=1:AnimateWire BP_sw33, 1:End Sub
Sub sw33_unHit()    : controller.switch (swRightOutLane)=0:AnimateWire BP_sw33, 0:End Sub
Sub sw34_Hit()      : controller.switch (swRightReturnLane)=1:AnimateWire BP_sw34, 1:End Sub
Sub sw34_unHit()    : controller.switch (swRightReturnLane)=0:AnimateWire BP_sw34, 0:End Sub
Sub sw35_Hit()      : controller.switch (swLeftOutLane)=1 :AnimateWire BP_sw35, 1:End Sub
Sub sw35_unHit()    : controller.switch (swLeftOutLane)=0:AnimateWire BP_sw35, 0:End Sub
Sub sw36_Hit()      : controller.switch (swLeftReturnLane)=1 :AnimateWire BP_sw36, 1:End Sub
Sub sw36_unHit()    : controller.switch (swLeftReturnLane)=0:AnimateWire BP_sw36, 0:End Sub

Sub AnimateWire(group, action) ' Action = 1 - to drop, 0 to raise)
  Dim BP
  If action = 1 Then
    For Each BP in group : BP.transz = -13 : Next
  Else
    For Each BP in group : BP.transz = 0 : Next
  End If
End Sub

'----------Star Triggers---------
Sub sw37_Hit()      : controller.switch (swAdvanceBonusRollOver)=1 :AnimateStar BP_sw37_001, 1:End Sub
Sub sw37_unHit()    : controller.switch (swAdvanceBonusRollOver)=0:AnimateStar BP_sw37_001, 0:End Sub

Sub AnimateStar(group, action) ' Action = 1 - to drop, 0 to raise)
  Dim BP
  If action = 1 Then
    For Each BP in group : BP.transz = -5 : Next
  Else
    For Each BP in group : BP.transz = 0 : Next
  End If
End Sub

'---------Bumpers---------
Sub Bumper1_Hit : vpmTimer.PulseSw(swBumper) : RandomSoundBumperTop Bumper1: End Sub



'************************************************************
' ZTAR: Targets
'************************************************************

'---------Drop Targets (Front Defenders)-----------
 Sub Sw26_Hit:DTHit 26:TargetBouncer Activeball, 1:End Sub
 Sub Sw27_Hit:DTHit 27:TargetBouncer Activeball, 1:End Sub
 Sub Sw28_Hit:DTHit 28:TargetBouncer Activeball, 1:End Sub

'---------Drop Targets (Back Defenders)-----------
 Sub Sw29_Hit:DTHit 29:TargetBouncer Activeball, 1:End Sub
 Sub Sw30_Hit:DTHit 30:TargetBouncer Activeball, 1:End Sub
 Sub Sw31_Hit:DTHit 31:TargetBouncer Activeball, 1:End Sub

'--------Drop Targets (XYZ Drops)----------
 Sub Sw46_Hit:DTHit 46:TargetBouncer Activeball, 1:End Sub
 Sub Sw47_Hit:DTHit 47:TargetBouncer Activeball, 1:End Sub
 Sub Sw48_Hit:DTHit 48:TargetBouncer Activeball, 1:End Sub

 '---------Stand Up Targets (H-Y-P-E)----------
 Sub sw21_Hit:STHit 21: End Sub
 Sub sw22_Hit:STHit 22: End Sub
 Sub sw23_Hit:STHit 23: End Sub
 Sub sw24_Hit:STHit 24: End Sub

'******************************************************
' ZBRL:  BALL ROLLING AND DROP SOUNDS
'******************************************************

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
  Const RollMinSpeed = 1
  Const RollMaxAbsVelZ = 1.2

  ' play the rolling sound for each ball
  For b = 0 To UBound(gBOT)
    If BallVel(gBOT(b)) > RollMinSpeed And Abs(gBOT(b).VelZ) < RollMaxAbsVelZ Then
      rolling(b) = True
      PlaySound ("BallRoll_" & b), -1, VolPlayfieldRoll(gBOT(b)) * BallRollVolume * VolumeDial, AudioPan(gBOT(b)), 0, PitchPlayfieldRoll(gBOT(b)), 1, 0, AudioFade(gBOT(b))
    Else
    If rolling(b) Then
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
dim RampBalls(4,2)
'x,0 = ball x,1 = ID, 2 = Protection against ending early (minimum amount of updates)
'0,0 is boolean on/off, 0,1 unused for now
RampBalls(0,0) = False

' RampType
'     Setup: Set this array to the number Total number of balls that can be tracked at one time + 1.  5 ball multiball then set value to 6
'     Description: Array type indexed on BallId and a values used to deterimine what type of ramp the ball is on: False = Wire Ramp, True = Plastic Ramp
dim RampType(4)

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


'------------Ramp triggers----------------

Sub LeftRampTop_hit
  If activeball.vely > 0 Then
    WireRampOn False
  Else
    WireRampOff
  End If
End Sub

Sub LeftRampBottom_unhit
  If activeball.vely < 0 Then
    WireRampOn False
  Else
    WireRampOff
  End If
End Sub

Sub MiddleRampTop_hit
  If activeball.vely > 0 Then
    WireRampOn False
  Else
    WireRampOff
  End If
End Sub

Sub MiddleRampBottom_unhit
  If activeball.vely < 0 Then
    WireRampOn False
  Else
    WireRampOff
  End If
End Sub

Sub RightRampTop_hit
  If activeball.vely > 0 Then
    WireRampOn False
  Else
    WireRampOff
  End If
End Sub

Sub RightRampBottom_unhit
  If activeball.vely < 0 Then
    WireRampOn False
  Else
    WireRampOff
  End If
End Sub

Sub LockRampTop_hit
  If activeball.vely > 0 Then
    WireRampOn False
  Else
    WireRampOff
  End If
End Sub

Sub LockRampBottom_unhit
  If activeball.vely < 0 Then
    WireRampOn False
  Else
    WireRampOff
  End If
End Sub

Sub LMRampTop_hit
  If activeball.vely > 0 Then
    WireRampOn False
  Else
    WireRampOff
  End If
End Sub

Sub LMRampBottom_unhit
  If activeball.vely < 0 Then
    WireRampOn False
  Else
    WireRampOff
  End If
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

  If tmp > 0 Then
    AudioFade = CSng(tmp ^ 10)
  Else
    AudioFade = CSng( - (( - tmp) ^ 10) )
  End If
End Function

Function AudioPan(tableobj) ' Calculates the pan for a tableobj based on the X position on the table. "table1" is the name of the table
  Dim tmp
  'debug.print tableobj.name
  tmp = tableobj.x * 2 / tablewidth - 1

  If tmp > 7000 Then
    tmp = 7000
  ElseIf tmp <  - 7000 Then
    tmp =  - 7000
  End If

  If tmp > 0 Then
    AudioPan = CSng(tmp ^ 10)
  Else
    AudioPan = CSng( - (( - tmp) ^ 10) )
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
Dim LF1 : Set LF1 = New FlipperPolarity
Dim RF1 : Set RF1 = New FlipperPolarity


InitPolarity


'*******************************************
' Late 70's to early 80's

Sub InitPolarity()
   dim x, a : a = Array(LF, RF, LF1, RF1)
  for each x in a
    x.AddPt "Ycoef", 0, RightFlipper.Y-65, 1 'disabled
    x.AddPt "Ycoef", 1, RightFlipper.Y-11, 1
    x.enabled = True
    x.TimeDelay = 80
    x.DebugOn=False ' prints some info in debugger


        x.AddPt "Polarity", 0, 0, 0
        x.AddPt "Polarity", 1, 0.05, - 2.7
        x.AddPt "Polarity", 2, 0.16, - 2.7
        x.AddPt "Polarity", 3, 0.22, - 0
        x.AddPt "Polarity", 4, 0.25, - 0
        x.AddPt "Polarity", 5, 0.3, - 1
        x.AddPt "Polarity", 6, 0.4, - 2
        x.AddPt "Polarity", 7, 0.5, - 2.7
        x.AddPt "Polarity", 8, 0.65, - 1.8
        x.AddPt "Polarity", 9, 0.75, - 0.5
        x.AddPt "Polarity", 10, 0.81, - 0.5
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
    LF1.SetObjects "LF1", LeftFlipper1, TriggerLF1
    RF1.SetObjects "RF1", RightFlipper1, TriggerRF1
End Sub



'*******************************************
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
'Sub InitPolarity()
' dim x, a : a = Array(LF, RF)
' for each x in a
'   x.AddPt "Ycoef", 0, RightFlipper.Y-65, 1 'disabled
'   x.AddPt "Ycoef", 1, RightFlipper.Y-11, 1
'   x.enabled = True
'   x.TimeDelay = 60
'   x.DebugOn=False ' prints some info in debugger
'
'   x.AddPt "Polarity", 0, 0, 0
'   x.AddPt "Polarity", 1, 0.05, - 5
'   x.AddPt "Polarity", 2, 0.16, - 5
'   x.AddPt "Polarity", 3, 0.22, - 0
'   x.AddPt "Polarity", 4, 0.25, - 0
'   x.AddPt "Polarity", 5, 0.3, - 2
'   x.AddPt "Polarity", 6, 0.4, - 3
'   x.AddPt "Polarity", 7, 0.5, - 4.0
'   x.AddPt "Polarity", 8, 0.7, - 3.5
'   x.AddPt "Polarity", 9, 0.75, - 3.0
'   x.AddPt "Polarity", 10, 0.8, - 2.5
'   x.AddPt "Polarity", 11, 0.85, - 2.0
'   x.AddPt "Polarity", 12, 0.9, - 1.5
'   x.AddPt "Polarity", 13, 0.95, - 1.0
'   x.AddPt "Polarity", 14, 1, - 0.5
'   x.AddPt "Polarity", 15, 1.1, 0
'   x.AddPt "Polarity", 16, 1.3, 0
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

' Next
'
' ' SetObjects arguments: 1: name of object 2: flipper object: 3: Trigger object around flipper
' LF.SetObjects "LF", LeftFlipper, TriggerLF
' RF.SetObjects "RF", RightFlipper, TriggerRF
'End Sub

'*******************************************
' Early 90's and after

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
  FlipperTricks LeftFlipper1, LF1Press, LF1Count, LF1EndAngle, LF1State
  FlipperTricks RightFlipper1, RF1Press, RF1Count, RF1EndAngle, RF1State
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
Dim LF1Press, RF1Press, LF1Count, RF1Count
Dim LFState, RFState
Dim LF1State, RF1State
Dim EOST, EOSA,Frampup, FElasticity,FReturn
Dim RFEndAngle, LFEndAngle
Dim RF1EndAngle, LF1EndAngle

Const FlipperCoilRampupMode = 2 '0 = fast, 1 = medium, 2 = slow (tap passes should work)

LFState = 1
LF1State = 1
RFState = 1
RF1State = 1
EOST = leftflipper.eostorque
EOSA = leftflipper.eostorqueangle
Frampup = LeftFlipper.rampup
FElasticity = LeftFlipper.elasticity
FReturn = LeftFlipper.return
Const EOSTnew = 1.5 'EM's to late 80's - new recommendation by rothbauerw (previously 1)
' Const EOSTnew = 1.2 '90's and later - new recommendation by rothbauerw (previously 0.8)
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
Const EOSReturn = 0.045  'late 70's to mid 80's
' Const EOSReturn = 0.035  'mid 80's to early 90's
'   Const EOSReturn = 0.025  'mid 90's and later

LFEndAngle = Leftflipper.endangle
RFEndAngle = RightFlipper.endangle
LF1EndAngle = Leftflipper1.endangle
RF1EndAngle = RightFlipper1.endangle

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

' Note, cor.update must be called in a 10 ms timer. The example table uses the GameTimer for this purpose, but sometimes a dedicated timer call RDampen is used.
'
Sub RDampen_Timer
  Cor.Update
End Sub

'******************************************************
'****  END PHYSICS DAMPENERS
'******************************************************


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
Dim DT26, DT27, DT28, DT29, DT30, DT31, DT46, DT47, DT48

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

Set DT26 = (new DropTarget)(sw26, sw26a, BM_sw26, 26, 0, false)
Set DT27 = (new DropTarget)(sw27, sw27a, BM_sw27, 27, 0, false)
Set DT28 = (new DropTarget)(sw28, sw28a, BM_sw28, 28, 0, false)
Set DT29 = (new DropTarget)(sw29, sw29a, BM_sw29, 29, 0, false)
Set DT30 = (new DropTarget)(sw30, sw30a, BM_sw30, 30, 0, false)
Set DT31 = (new DropTarget)(sw31, sw31a, BM_sw31, 31, 0, false)
Set DT46 = (new DropTarget)(sw46, sw46a, BM_sw46, 46, 0, false)
Set DT47 = (new DropTarget)(sw47, sw47a, BM_sw47, 47, 0, false)
Set DT48 = (new DropTarget)(sw48, sw48a, BM_sw48, 48, 0, false)



Dim DTArray
DTArray = Array(DT26, DT27, DT28, DT29, DT30, DT31, DT46, DT47, DT48)

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
  RandomSoundDropTargetReset DTArray(i).Prim
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


Sub UpdateDropTargets
  dim BP, tz, rx, ry

  tz = BM_sw26.transz
  rx = BM_sw26.rotx
  ry = BM_sw26.roty
  For each BP in BP_sw26: BP.transz = tz: BP.rotx = rx: BP.roty = ry: Next

  tz = BM_sw27.transz
  rx = BM_sw27.rotx
  ry = BM_sw27.roty
  For each BP in BP_sw27: BP.transz = tz: BP.rotx = rx: BP.roty = ry: Next

  tz = BM_sw28.transz
  rx = BM_sw28.rotx
  ry = BM_sw28.roty
  For each BP in BP_sw28: BP.transz = tz: BP.rotx = rx: BP.roty = ry: Next

  tz = BM_sw29.transz
  rx = BM_sw29.rotx
  ry = BM_sw29.roty
  For each BP in BP_sw29: BP.transz = tz: BP.rotx = rx: BP.roty = ry: Next

  tz = BM_sw30.transz
  rx = BM_sw30.rotx
  ry = BM_sw30.roty
  For each BP in BP_sw30: BP.transz = tz: BP.rotx = rx: BP.roty = ry: Next

  tz = BM_sw31.transz
  rx = BM_sw31.rotx
  ry = BM_sw31.roty
  For each BP in BP_sw31: BP.transz = tz: BP.rotx = rx: BP.roty = ry: Next

  tz = BM_sw46.transz
  rx = BM_sw46.rotx
  ry = BM_sw46.roty
  For each BP in BP_sw46: BP.transz = tz: BP.rotx = rx: BP.roty = ry: Next

  tz = BM_sw47.transz
  rx = BM_sw47.rotx
  ry = BM_sw47.roty
  For each BP in BP_sw47: BP.transz = tz: BP.rotx = rx: BP.roty = ry: Next

  tz = BM_sw48.transz
  rx = BM_sw48.rotx
  ry = BM_sw48.roty
  For each BP in BP_sw48: BP.transz = tz: BP.rotx = rx: BP.roty = ry: Next
End Sub



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
Dim ST21, ST22, ST23, ST24

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

Set ST21 = (new StandupTarget)(sw21, BM_sw21, 21, 0)
Set ST22 = (new StandupTarget)(sw22, BM_sw22, 22, 0)
Set ST23 = (new StandupTarget)(sw23, BM_sw23, 23, 0)
Set ST24 = (new StandupTarget)(sw24, BM_sw24, 24, 0)


'Add all the Stand-up Target Arrays to Stand-up Target Animation Array
' STAnimationArray = Array(ST1, ST2, ....)
Dim STArray
STArray = Array(ST21, ST22, ST23, ST24)

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


Sub UpdateStandupTargets
  dim BP, ty

    ty = BM_sw21.transy
  For each BP in BP_sw21 : BP.transy = ty: Next

    ty = BM_sw22.transy
  For each BP in BP_sw22 : BP.transy = ty: Next

    ty = BM_sw23.transy
  For each BP in BP_sw23 : BP.transy = ty: Next

    ty = BM_sw24.transy
  For each BP in BP_sw24 : BP.transy = ty: Next

End Sub


'******************************************************
'***  END STAND-UP TARGETS
'******************************************************


'******************************************************
'   ZBOU: VPW TargetBouncer for targets and posts by Iaakki, Wrd1972, Apophis
'******************************************************

Const TargetBouncerEnabled = 0    '0 = normal standup targets, 1 = bouncy targets
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
'****  ZGIU:  GI Control
'******************************************************

'**** Global variable to hold current GI light intensity. Not Mandatory, but may help in other lighting subs and timers
'**** This value is updated always when GI state is being updated


'**** Use these for GI strings and stepped GI, and comment out the SetRelayGI
'**** This example table uses Relay for GI control, we don't need these at all

'Set GICallback2 = GetRef("GIUpdates2")   'use this for stepped/modulated GI
'
''GIupdates2 is called always when some event happens to GI channel.
'Sub GIUpdates2(aNr, aLvl)
''  debug.print "GIUpdates2 nr: " & aNr & " value: " & aLvl
' Dim bulb
'
' Select Case aNr 'Strings are selected here
'
'   Case 3:  'GI String 1 (Upper)
'     ' Update the state for each GI light. The state will be a float value between 0 and 1.
'     For Each bulb in GI_Lower: bulb.State = aLvl: Next
'
'     ' If the GI has an associated Relay sound, this can be played
'     'If aLvl >= 0.5 And gilvl < 0.5 Then
'     ' Sound_GI_Relay 1, Bumper1 'Note: Bumper1 is just used for sound positioning. Can be anywhere that makes sense.
'     'ElseIf aLvl <= 0.4 And gilvl > 0.4 Then
'     ' Sound_GI_Relay 0, Bumper1
'     'End If
'
'     'You may add any other GI related effects here. Like if you want to make some toy to appear more bright, set it like this:
'     'Primitive001.blenddisablelighting = 1.2 * aLvl + 0.2 'This will result to DL brightness between 0.2 - 1.4 for ON/OFF states
'
'     gilvl = aLvl    'Storing the latest GI fading state into global variable, so one can use it elsewhere too.
'   Case 2:  'GI String 2 (Center)
'     ' Update the state for each GI light. The state will be a float value between 0 and 1.
'     For Each bulb in GI_Upper: bulb.State = aLvl: Next
'
'     ' If the GI has an associated Relay sound, this can be played
'     'If aLvl >= 0.5 And gilvl < 0.5 Then
'     ' Sound_GI_Relay 1, Bumper1 'Note: Bumper1 is just used for sound positioning. Can be anywhere that makes sense.
'     'ElseIf aLvl <= 0.4 And gilvl > 0.4 Then
'     ' Sound_GI_Relay 0, Bumper1
'     'End If
'
'     'You may add any other GI related effects here. Like if you want to make some toy to appear more bright, set it like this:
'     'Primitive001.blenddisablelighting = 1.2 * aLvl + 0.2 'This will result to DL brightness between 0.2 - 1.4 for ON/OFF states
'
'     gilvl = aLvl    'Storing the latest GI fading state into global variable, so one can use it elsewhere too.
'   Case 2:  'GI String 3 (Lower)
'     ' Update the state for each GI light. The state will be a float value between 0 and 1.
'     For Each bulb in GI_Lower: bulb.State = aLvl: Next
'
'     ' If the GI has an associated Relay sound, this can be played
'     'If aLvl >= 0.5 And gilvl < 0.5 Then
'     ' Sound_GI_Relay 1, Bumper1 'Note: Bumper1 is just used for sound positioning. Can be anywhere that makes sense.
'     'ElseIf aLvl <= 0.4 And gilvl > 0.4 Then
'     ' Sound_GI_Relay 0, Bumper1
'     'End If
'
'     'You may add any other GI related effects here. Like if you want to make some toy to appear more bright, set it like this:
'     'Primitive001.blenddisablelighting = 1.2 * aLvl + 0.2 'This will result to DL brightness between 0.2 - 1.4 for ON/OFF states
'
'     gilvl = aLvl    'Storing the latest GI fading state into global variable, so one can use it elsewhere too.
'   Case 3:  'GI String 3 (Backbox)
'     For Each bulb in GI_4: bulb.State = aLvl: Next
'     gilvl = aLvl
'   Case 4:  'GI String 4 (Backbox)
'     For Each bulb in GI_5: bulb.State = aLvl: Next
'     gilvl = aLvl
'
' End Select
'
'End Sub


'***************************************************************
' ZSHA: Ambient ball shadows
'***************************************************************

' For dynamic ball shadows, Check the "Raytraced ball shadows" box for the specific light.
' Also make sure the light's z position is around 25 (mid ball)

'Ambient (Room light source)
Const AmbientBSFactor = 0.7    '0 To 1, higher is darker
Const AmbientMovement = 1    '1+ higher means more movement as the ball moves left and right
Const offsetX = 0        'Offset x position under ball (These are if you want to change where the "room" light is for calculating the shadow position,)
Const offsetY = 0        'Offset y position under ball (^^for example 5,5 if the light is in the back left corner)

' *** Trim or extend these to match the number of balls/primitives/flashers on the table!  (will throw errors if there aren't enough objects)
Dim objBallShadow(5)

'Initialization
BSInit

Sub BSInit()
  Dim iii
  'Prepare the shadow objects before play begins
  For iii = 0 To tnob- 1
    Set objBallShadow(iii) = Eval("BallShadow" & iii)
    objBallShadow(iii).material = "BallShadow" & iii
    UpdateMaterial objBallShadow(iii).material,1,0,0,0,0,0,AmbientBSFactor,RGB(0,0,0),0,0,False,True,0,0,0,0
    objBallShadow(iii).Z = 3 + iii / 1000
    objBallShadow(iii).visible = 0
  Next
End Sub

Dim s
Sub BSUpdate
  For s = lob To UBound(gBOT)
    ' *** Normal "ambient light" ball shadow

    'Primitive shadow on playfield, flasher shadow in ramps
    '** If on main and upper pf
    If gBOT(s).Z > 20 Then'And gBOT(s).Z < 30 Then
      objBallShadow(s).visible = 1
      objBallShadow(s).X = gBOT(s).X + (gBOT(s).X - (tablewidth / 2)) / (Ballsize / AmbientMovement) + offsetX
      objBallShadow(s).Y = gBOT(s).Y + offsetY
      objBallShadow(s).Z = gBOT(s).Z + s/1000 + 1.04 - 25

    '** No shadow if ball is off the main playfield (this may need to be adjusted per table)
    Else
      objBallShadow(s).visible = 0
    End If
  Next
End Sub

'**********************************************************************************************************
'Digital Display
'**********************************************************************************************************
Dim Digit(31)
Dim Patterns(11)
Dim Patterns2(11)

Patterns(0) = 0     'empty
Patterns(1) = 63    '0
Patterns(2) = 6     '1
Patterns(3) = 91    '2
Patterns(4) = 79    '3
Patterns(5) = 102   '4
Patterns(6) = 109   '5
Patterns(7) = 125   '6
Patterns(8) = 7     '7
Patterns(9) = 127   '8
Patterns(10) = 111  '9

Patterns2(0) = 128  'empty
Patterns2(1) = 191  '0
Patterns2(2) = 134  '1
Patterns2(3) = 219  '2
Patterns2(4) = 207  '3
Patterns2(5) = 230  '4
Patterns2(6) = 237  '5
Patterns2(7) = 253  '6
Patterns2(8) = 135  '7
Patterns2(9) = 255  '8
Patterns2(10) = 239 '9

'Assign 7-digit output to reels
Set Digit(0) = a0
Set Digit(1) = a1
Set Digit(2) = a2
Set Digit(3) = a3
Set Digit(4) = a4
Set Digit(5) = a5
Set Digit(6) = a6

Set Digit(7) = b0
Set Digit(8) = b1
Set Digit(9) = b2
Set Digit(10) = b3
Set Digit(11) = b4
Set Digit(12) = b5
Set Digit(13) = b6

Set Digit(14) = c0
Set Digit(15) = c1
Set Digit(16) = c2
Set Digit(17) = c3
Set Digit(18) = c4
Set Digit(19) = c5
Set Digit(20) = c6

Set Digit(21) = d0
Set Digit(22) = d1
Set Digit(23) = d2
Set Digit(24) = d3
Set Digit(25) = d4
Set Digit(26) = d5
Set Digit(27) = d6

Set Digit(28) = e0
Set Digit(29) = e1
Set Digit(30) = e2
Set Digit(31) = e3


dim DisplayColor, DisplayColorG
DisplayColor =  RGB(0,37,251)


Sub FadeDisplay(object, onoffstat)
  If OnOffstat = 1 Then
    object.color = DisplayColor
  Else
    Object.Color = RGB(1,1,1)
  End If
End Sub

Dim Digits(40)
Digits(0)=Array()
Digits(1)=Array()
Digits(2)=Array()
Digits(3)=Array()
Digits(4)=Array()
Digits(5)=Array()
Digits(6)=Array()
Digits(7)=Array()
Digits(8)=Array()
Digits(9)=Array()
Digits(10)=Array()
Digits(11)=Array()
Digits(12)=Array()
Digits(13)=Array()
Digits(14)=Array()
Digits(15)=Array()
Digits(16)=Array()
Digits(17)=Array()
Digits(18)=Array()
Digits(19)=Array()
Digits(20)=Array()
Digits(21)=Array()
Digits(22)=Array()
Digits(23)=Array()
Digits(24)=Array()
Digits(25)=Array()
Digits(26)=Array()
Digits(27)=Array()
Digits(28)=Array()
Digits(29)=Array()
Digits(30)=Array()
Digits(31)=Array()
Digits(32)=Array(aa00,aa01,aa02,aa03,aa04,aa05,aa06)
Digits(33)=Array(a10,a11,a12,a13,a14,a15,a16)
Digits(34)=Array(a20,a21,a22,a23,a24,a25,a26)
Digits(35)=Array(a30,a31,a32,a33,a34,a35,a36)
Digits(36)=Array(a40,a41,a42,a43,a44,a45,a46)
Digits(37)=Array(a50,a51,a52,a53,a54,a55,a56)
Digits(38)=Array()
Digits(39)=Array()


Sub InitDigits()
  dim tmp, x, obj
  for x = 0 to uBound(Digits)
    if IsArray(Digits(x) ) then
      For each obj in Digits(x)
'       obj.height = obj.height + 18
        FadeDisplay obj, 0
      next
    end If
  Next
End Sub


Sub DTDisplayTimer
    Dim ChgLED, ii, num, chg, stat, jj, obj

    'CALL THIS ONCE PER FRAME/TICK
    ChgLED = Controller.ChangedLEDs(&HFFFFFFFF, &HFFFFFFFF)
    If IsEmpty(ChgLED) Then Exit Sub

    For ii = 0 To UBound(ChgLED)

        num  = ChgLED(ii, 0)
        chg  = ChgLED(ii, 1)
        stat = ChgLED(ii, 2)

        '---------------------------
        ' REELS (7-seg reels)
        '---------------------------
    If VRMode = 0 and cab_mode = 0 Then
        If num >= 0 And num <= UBound(Digit) Then
            For jj = 0 To 10
                If stat = Patterns(jj) Or stat = Patterns2(jj) Then
                    Digit(num).SetValue jj
                    Exit For
                End If
            Next
        End If
    End If

        '---------------------------
        ' LIGHT OBJECT DIGITS (segment lights)
        '---------------------------
        If num >= 0 And num <= UBound(Digits) Then
            If IsArray(Digits(num)) Then
                'only do work if this digit actually has segment objects
                On Error Resume Next
                If UBound(Digits(num)) >= 0 Then
                    On Error GoTo 0

                    Dim chg2, stat2
                    chg2 = chg
                    stat2 = stat

                    For Each obj In Digits(num)
                        If (chg2 And 1) <> 0 Then obj.State = (stat2 And 1)
                        chg2 = chg2 \ 2
                        stat2 = stat2 \ 2
                    Next
                Else
                    On Error GoTo 0
                End If
            End If
        End If

    Next
End Sub

'******************************
' Setup VR Backglass
'******************************

Dim xoff,yoff1, yoff2, yoff3, yoff4, yoff5, zoff, xrot, zscale, xcen, ycen

Sub setup_backglass()

  xoff = -20
  yoff1 = 92 ' this is where you adjust the forward/backward position for player 1 score
  yoff2 = 92 ' this is where you adjust the forward/backward position for player 2 score
  yoff3 = 92 ' this is where you adjust the forward/backward position for player 3 score
  yoff4 = 92 ' this is where you adjust the forward/backward position for player 4 score
  yoff5 = 92  ' this is where you adjust the forward/backward position for credits and ball in play
  zoff = 799
  xrot = -90

  center_digits()

end sub


Sub center_digits()
  Dim ix, xx, yy, yfact, xfact, xobj

  zscale = 0.0000001

  xcen = (130 /2) - (92 / 2)
  ycen = (780 /2 ) + (203 /2)

  for ix = 0 to 6
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

  for ix = 7 to 13
    For Each xobj In VRDigits(ix)

      xx = xobj.x

      xobj.x = (xoff - xcen) + xx
      yy = xobj.y ' get the yoffset before it is changed
      xobj.y = yoff2

      If (yy < 0.) then
        yy = yy * -1
      end if

      xobj.height = (zoff - ycen) + yy - (yy * (zscale))
      xobj.rotx = xrot
    Next
  Next

  for ix = 14 to 20
    For Each xobj In VRDigits(ix)

      xx = xobj.x

      xobj.x = (xoff - xcen) + xx
      yy = xobj.y ' get the yoffset before it is changed
      xobj.y = yoff3

      If (yy < 0.) then
        yy = yy * -1
      end if

      xobj.height = (zoff - ycen) + yy - (yy * (zscale))
      xobj.rotx = xrot
    Next
  Next

  for ix = 21 to 27
    For Each xobj In VRDigits(ix)

      xx = xobj.x

      xobj.x = (xoff - xcen) + xx
      yy = xobj.y ' get the yoffset before it is changed
      xobj.y = yoff4

      If (yy < 0.) then
        yy = yy * -1
      end if

      xobj.height = (zoff - ycen) + yy - (yy * (zscale))
      xobj.rotx = xrot
    Next
  Next

  for ix = 28 to 31
    For Each xobj In VRDigits(ix)

      xx = xobj.x

      xobj.x = (xoff - xcen) + xx
      yy = xobj.y ' get the yoffset before it is changed
      xobj.y = yoff5

      If (yy < 0.) then
        yy = yy * -1
      end if

      xobj.height = (zoff - ycen) + yy - (yy * (zscale))
      xobj.rotx = xrot
    Next
  Next

end sub

'********************************************
'              Display Output
'********************************************


Dim VRDigits(38)
VRDigits(0) = Array(LED1x0,LED1x1,LED1x2,LED1x3,LED1x4,LED1x5,LED1x6)
VRDigits(1) = Array(LED2x0,LED2x1,LED2x2,LED2x3,LED2x4,LED2x5,LED2x6)
VRDigits(2) = Array(LED3x0,LED3x1,LED3x2,LED3x3,LED3x4,LED3x5,LED3x6)
VRDigits(3) = Array(LED4x0,LED4x1,LED4x2,LED4x3,LED4x4,LED4x5,LED4x6)
VRDigits(4) = Array(LED5x0,LED5x1,LED5x2,LED5x3,LED5x4,LED5x5,LED5x6)
VRDigits(5) = Array(LED6x0,LED6x1,LED6x2,LED6x3,LED6x4,LED6x5,LED6x6)
VRDigits(6) = Array(LED7x0,LED7x1,LED7x2,LED7x3,LED7x4,LED7x5,LED7x6)

VRDigits(7) = Array(LED8x0,LED8x1,LED8x2,LED8x3,LED8x4,LED8x5,LED8x6)
VRDigits(8) = Array(LED9x0,LED9x1,LED9x2,LED9x3,LED9x4,LED9x5,LED9x6)
VRDigits(9) = Array(LED10x0,LED10x1,LED10x2,LED10x3,LED10x4,LED10x5,LED10x6)
VRDigits(10) = Array(LED11x0,LED11x1,LED11x2,LED11x3,LED11x4,LED11x5,LED11x6)
VRDigits(11) = Array(LED12x0,LED12x1,LED12x2,LED12x3,LED12x4,LED12x5,LED12x6)
VRDigits(12) = Array(LED13x0,LED13x1,LED13x2,LED13x3,LED13x4,LED13x5,LED13x6)
VRDigits(13) = Array(LED14x0,LED14x1,LED14x2,LED14x3,LED14x4,LED14x5,LED14x6)

VRDigits(14) = Array(LED1x000,LED1x001,LED1x002,LED1x003,LED1x004,LED1x005,LED1x006)
VRDigits(15) = Array(LED1x100,LED1x101,LED1x102,LED1x103,LED1x104,LED1x105,LED1x106)
VRDigits(16) = Array(LED1x200,LED1x201,LED1x202,LED1x203,LED1x204,LED1x205,LED1x206)
VRDigits(17) = Array(LED1x300,LED1x301,LED1x302,LED1x303,LED1x304,LED1x305,LED1x306)
VRDigits(18) = Array(LED1x400,LED1x401,LED1x402,LED1x403,LED1x404,LED1x405,LED1x406)
VRDigits(19) = Array(LED1x500,LED1x501,LED1x502,LED1x503,LED1x504,LED1x505,LED1x506)
VRDigits(20) = Array(LED1x600,LED1x601,LED1x602,LED1x603,LED1x604,LED1x605,LED1x606)

VRDigits(21) = Array(LED2x000,LED2x001,LED2x002,LED2x003,LED2x004,LED2x005,LED2x006)
VRDigits(22) = Array(LED2x100,LED2x101,LED2x102,LED2x103,LED2x104,LED2x105,LED2x106)
VRDigits(23) = Array(LED2x200,LED2x201,LED2x202,LED2x203,LED2x204,LED2x205,LED2x206)
VRDigits(24) = Array(LED2x300,LED2x301,LED2x302,LED2x303,LED2x304,LED2x305,LED2x306)
VRDigits(25) = Array(LED2x400,LED2x401,LED2x402,LED2x403,LED2x404,LED2x405,LED2x406)
VRDigits(26) = Array(LED2x500,LED2x501,LED2x502,LED2x503,LED2x504,LED2x505,LED2x506)
VRDigits(27) = Array(LED2x600,LED2x601,LED2x602,LED2x603,LED2x604,LED2x605,LED2x606)

'credit -- Ball In Play
VRDigits(28) = Array(LEDax300,LEDax301,LEDax302,LEDax303,LEDax304,LEDax305,LEDax306)
VRDigits(29) = Array(LEDbx400,LEDbx401,LEDbx402,LEDbx403,LEDbx404,LEDbx405,LEDbx406)
VRDigits(30) = Array(LEDcx500,LEDcx501,LEDcx502,LEDcx503,LEDcx504,LEDcx505,LEDcx506)
VRDigits(31) = Array(LEDdx600,LEDdx601,LEDdx602,LEDdx603,LEDdx604,LEDdx605,LEDdx606)


dim VRDisplayColor
VRDisplayColor =  RGB(255,40,1)

Sub VRDisplayTimer
    Dim ChgLED, ii, jj, num, chg, stat, obj, b, x
    ChgLED=Controller.ChangedLEDs(&Hffffffff, &Hffffffff)
    If Not IsEmpty(ChgLED)Then
       For ii=0 To UBound(chgLED)
          num=chgLED(ii, 0) : chg=chgLED(ii, 1) : stat=chgLED(ii, 2)
        '-----------------------------------------
        ' VR BACKGLASS 7-SEG OBJECTS (0..31)
        '-----------------------------------------
        If num >= 0 And num <= 31 Then
            Dim chgVR, statVR
            chgVR = chg
            statVR = stat

            If IsArray(VRDigits(num)) Then
                For Each obj In VRDigits(num)
                    If (chgVR And 1) <> 0 Then
                        obj.Visible = CBool(statVR And 1)
                        VRFadeDisplay obj, (statVR And 1)
                    End If
                    chgVR = chgVR \ 2
                    statVR = statVR \ 2
                Next
            End If
        End If

        '-----------------------------------------
        ' LIGHT OBJECT DIGITS (segment lights) - same as DTDisplayTimer
        '-----------------------------------------
        If num >= 0 And num <= UBound(Digits) Then
            If IsArray(Digits(num)) Then
                On Error Resume Next
                If UBound(Digits(num)) >= 0 Then
                    On Error GoTo 0

                    Dim chg2, stat2
                    chg2 = chg
                    stat2 = stat

                    For Each obj In Digits(num)
                        If (chg2 And 1) <> 0 Then obj.State = (stat2 And 1)
                        chg2 = chg2 \ 2
                        stat2 = stat2 \ 2
                    Next
                Else
                    On Error GoTo 0
                End If
            End If
        End If
        Next
    End If
 End Sub

Sub VRFadeDisplay(object, onoff)
  Dim numcol

  If OnOff = 1 Then

    for numcol = 0 to 31
      if IsArray(VRDigits(numcol)) Then
        for each Object in VRDigits(numcol)
          object.color = VRDisplayColor
          Object.Opacity = 12
        Next
      End If
    Next

    for numcol = 32 to 37

      if IsArray(VRDigits(numcol)) Then
        for each Object in VRDigits(numcol)
          object.color = VRDisplayColor
          Object.Opacity = 12
        Next
      End If
    Next

  Else
    Object.Color = RGB(1,1,1)
    Object.Opacity = 6
  End If
End Sub


Sub VRInitDigits()
  dim tmp, x, obj
  for x = 0 to uBound(VRDigits)
    if IsArray(VRDigits(x) ) then
      For each obj in VRDigits(x)
        obj.height = obj.height + 18
        FadeDisplay obj, 0
      next
    end If
  Next
End Sub


Dim VRThings
Dim BGThings
Dim DTThings

'Desktop Mode
if VRMode = 0 and cab_mode = 0 Then
  for each VRThings in VRCabinet:VRThings.visible = 0:Next
  for each VRThings in VRMinimal:VRThings.visible = 0:Next
  For each DTThings in DTDigits: DTThings.Visible = 1: Next
  VR_Cab_Siderails.visible = 1
  VR_Cab_Lockdownbar.visible = 1
  VR_Backbox_Wall2.isDropped = 1


' Cabinet Mode
Elseif VRMode = 0 and cab_mode = 1 Then
  for each VRThings in VRCabinet:VRThings.visible = 0:Next
  for each VRThings in VRMinimal:VRThings.visible = 0:Next
  For each DTThings in DTDigits: DTThings.Visible = 0: Next
  VR_Cab_Siderails.visible = 0
  VR_Cab_Lockdownbar.visible = 0

' VR Mode
Else
  for each VRThings in VRCabinet:VRThings.visible = 1:Next
  for each VRThings in VRMinimal:VRThings.visible = 1:Next
  For each DTThings in DTDigits: DTThings.Visible = 0: Next
End If

'
'
''Bally Vector
''added by Inkochnito
Sub editDips
  if SetDIPSwitches = 1 Then
  Dim vpmDips : Set vpmDips = New cvpmDips
  With vpmDips
    .AddForm 700,400,"Vector - DIP switches"
    .AddChk 7,10,180,Array("Match feature", &H08000000)'dip 28
    .AddChk 205,10,115,Array("Credits display", &H04000000)'dip 27
    .AddFrame 2,30,190,"Maximum credits",&H03000000,Array("10 credits",0,"15 credits", &H01000000,"25 credits", &H02000000,"40 credits", &H03000000)'dip 25&26
    .AddFrame 2,106,190,"Making 3 upper left targets will",&H00000020,Array("reset 3 lower targets",0,"not reset 3 lower targets",&H00000020)'dip 6
    .AddFrame 2,152,190,"Multipiers lite will",&H00000040,Array("be off then alternate",0,"stay on till bonus lites are made",&H00000040)'dip 7
    .AddFrame 2,198,190,"X-Y-Z drop targets special",&H00000080,Array("25K - special - 25K",0,"25K - special keeps alternating",&H00000080)'dip 8
    .AddFrame 2,248,190,"Vectorscan to date readout",&H00002000,Array("can only be decreased manually",0,"after 8 games decrease by 20K",&H00002000)'dip 14
    .AddFrame 2,298,190,"With Vectorscan capture lite on",&H00004000,Array("3 lower targets will reset",0,"3 lower targets will go back down",&H00004000)'dip 15
    .AddFrame 2,348,190,"With Vectorscan capture lite off",32768,Array("targets down will reset",0,"any target down will go back down",32768)'dip 16
    .AddFrame 205,30,190,"Balls per game",&HC0000000,Array ("2 balls",&HC0000000,"3 balls",0,"4 balls",&H80000000,"5 balls",&H40000000)'dip 31&32
    .AddFrame 205,106,190,"Vectorscan bonus score readout",&H00100000,Array("readout will reset to 1",0,"readout will come back on",&H00100000)'dip 21
    .AddFrame 205,152,190,"Attract sound",&H00200000,Array("no voice",0,"voice says 'I am P.A.C. play analysis",&H00200000)'dip 22
    .AddFrame 205,198,190,"Vectorscan scoring adjust",&H00400000,Array("scores only when capture lite is on",0,"scores when capture lite is on or off",&H00400000)'dip 23
    .AddFrame 205,248,190,"Lite special and extra ball when",&H00800000,Array("hitting vector speed 800 or over",0,"hitting vector speed 750 or over",&H00800000)'dip 24
    .AddFrame 205,298,190,"Replay limit",&H10000000,Array("no limit",0,"1 replay per game",&H10000000)'dip 29
    .AddFrame 205,348,190,"H-Y-P-E target bonus adjust",&H20000000,Array("advances bonus only when lit",0,"advances bonus every time",&H20000000)'dip 30
    .AddLabel 50,400,350,20,"Set selftest position 16,17,18 and 19 to 03 for the best gameplay."
    .AddLabel 50,420,300,20,"After hitting OK, press F3 to reset game with new settings."
    .ViewDips
  End With
  End If
End Sub
Set vpmShowDips = GetRef("editDips")

'Recommended settings by Bally, but for conservative/harder play.

'  *** This is the default dips settings ***

'1-5  00000   1 credit/coins chute left
'6    1     Lower 3 left side drop target
'7    1   X-Y-Z Drop Targets Multipliers Lite
'8    0   X-Y-Z Drop Targets Special Lite
'9-13   00000   1 credit/coins chute right
'14   1     Vectorscan to date readout
'15   1     Vectorscan capture ball lite on, 3 lower left side targets
'16   1   Vectorscan capture ball lite on, 3 lower left side targets
'17-20  0010  Max credits of 40
'21   1     Vectorscan Bonus Score Readout
'22   1   Coined game voice
'23   1   Capture ball lite, Vectorscan speed readout scoring
'24   0   Vectorscan score readout threshold special and extra ball
'25-28  0000  Credit/coins chutes 2 same as chute 1
'29   1   Number of games replays per game
'30   1   H-Y-P-E target bonus advances
'31   0   Balls per Game (3 off, 5 On)
'32   0   Balls per game


Sub SetDefaultDips
  If SetDIPSwitches = 0 Then
    Controller.Dip(0) = 96    '8-1 = 01100000
    Controller.Dip(1) = 224   '16-9 = 11100000
    Controller.Dip(2) = 114   '24-17 = 01110010
    Controller.Dip(3) = 60    '32-25 = 00111100
    Controller.Dip(4) = 0
    Controller.Dip(5) = 0
  End If
End Sub




' VLM  Arrays - Start
' Arrays per baked part
Dim BP_Bumper1_Ring: BP_Bumper1_Ring=Array(BM_Bumper1_Ring, LM_GI_L12a_Bumper1_Ring, LM_GI_L49a2_Bumper1_Ring, LM_GI_L49a3_Bumper1_Ring, LM_GI_l113b_Bumper1_Ring, LM_GI_l114_Bumper1_Ring, LM_GI_l97b_Bumper1_Ring)
Dim BP_Cube: BP_Cube=Array(BM_Cube)
Dim BP_Gate1: BP_Gate1=Array(BM_Gate1, LM_GI_l83_Gate1)
Dim BP_LSling: BP_LSling=Array(BM_LSling, LM_GI_L33a1_LSling, LM_GI_L33a4_LSling, LM_GI_L33a6_LSling, LM_UL_l2_LSling, LM_L_l30_LSling, LM_UL_l30_LSling, LM_L_l40_LSling, LM_UL_l40_LSling)
Dim BP_LSling1: BP_LSling1=Array(BM_LSling1, LM_GI_L33a1_LSling1, LM_GI_L33a4_LSling1, LM_GI_L33a6_LSling1, LM_UL_l2_LSling1, LM_L_l30_LSling1, LM_UL_l30_LSling1, LM_L_l40_LSling1, LM_UL_l40_LSling1)
Dim BP_LSling2: BP_LSling2=Array(BM_LSling2, LM_GI_L33a1_LSling2, LM_GI_L33a4_LSling2, LM_GI_L33a6_LSling2, LM_UL_l2_LSling2, LM_L_l30_LSling2, LM_UL_l30_LSling2, LM_L_l40_LSling2, LM_UL_l40_LSling2)
Dim BP_Layer_1: BP_Layer_1=Array(BM_Layer_1, LM_GI_L12a_Layer_1, LM_GI_L33a1_Layer_1, LM_GI_L33a2_Layer_1, LM_GI_L33a3_Layer_1, LM_GI_L33a4_Layer_1, LM_GI_L33a5_Layer_1, LM_GI_L33a6_Layer_1, LM_GI_L33a7_Layer_1, LM_GI_L49a1_Layer_1, LM_GI_L49a2_Layer_1, LM_GI_L49a3_Layer_1, LM_GI_L49a4_Layer_1, LM_GI_L49a5_Layer_1, LM_GI_L49a6_Layer_1, LM_GI_L67b_Layer_1, LM_GI_CreditLight_Layer_1, LM_L_l10_Layer_1, LM_UL_l10_Layer_1, LM_L_l113_Layer_1, LM_GI_l113b_Layer_1, LM_L_l114_Layer_1, LM_UL_l114_Layer_1, LM_GI_l114_Layer_1, LM_L_l115_Layer_1, LM_UL_l115_Layer_1, LM_L_l23_Layer_1, LM_UL_l23_Layer_1, LM_L_l30_Layer_1, LM_L_l39_Layer_1, LM_UL_l39_Layer_1, LM_L_l40_Layer_1, LM_L_l43_Layer_1, LM_L_l47_Layer_1, LM_UL_l47_Layer_1, LM_UL_l55_Layer_1, LM_UL_l56_Layer_1, LM_UL_l57_Layer_1, LM_L_l65_Layer_1, LM_UL_l65_Layer_1, LM_GI_l65b_Layer_1, LM_L_l8_Layer_1, LM_L_l81_Layer_1, LM_UL_l81_Layer_1, LM_GI_l81b_Layer_1, LM_UL_l82_Layer_1, LM_UL_l9_Layer_1, LM_L_l97_Layer_1, LM_GI_l97b_Layer_1, LM_L_l98_Layer_1, LM_UL_l98_Layer_1, _
  LM_GI_l98_Layer_1)
Dim BP_LeftFlip: BP_LeftFlip=Array(BM_LeftFlip, LM_GI_L33a1_LeftFlip, LM_GI_L33a4_LeftFlip, LM_GI_L33a6_LeftFlip)
Dim BP_LeftFlip1: BP_LeftFlip1=Array(BM_LeftFlip1, LM_GI_L33a1_LeftFlip1, LM_UL_l115_LeftFlip1, LM_UL_l99_LeftFlip1)
Dim BP_LeftFlip1U: BP_LeftFlip1U=Array(BM_LeftFlip1U, LM_GI_L33a1_LeftFlip1U, LM_GI_L33a5_LeftFlip1U, LM_UL_l60_LeftFlip1U, LM_L_l60_LeftFlip1U)
Dim BP_LeftFlipU: BP_LeftFlipU=Array(BM_LeftFlipU, LM_GI_L33a1_LeftFlipU, LM_GI_L33a2_LeftFlipU, LM_GI_L33a3_LeftFlipU, LM_GI_L33a4_LeftFlipU, LM_GI_L33a5_LeftFlipU, LM_GI_L33a6_LeftFlipU, LM_UL_l19_LeftFlipU, LM_L_l3_LeftFlipU, LM_UL_l3_LeftFlipU, LM_GI_l83_LeftFlipU)
Dim BP_Parts: BP_Parts=Array(BM_Parts, BM_Parts_001, BM_Parts_002, LM_GI_L12a_Parts, LM_GI_L33a1_Parts, LM_GI_L33a2_Parts, LM_GI_L33a3_Parts, LM_GI_L33a4_Parts, LM_GI_L33a5_Parts, LM_GI_L33a6_Parts, LM_GI_L33a7_Parts, LM_GI_L49a1_Parts, LM_GI_L49a2_Parts, LM_GI_L49a3_Parts, LM_GI_L49a4_Parts, LM_GI_L49a5_Parts, LM_GI_L49a6_Parts, LM_GI_L67b_Parts, LM_L_l10_Parts, LM_UL_l10_Parts, LM_L_l113_Parts, LM_UL_l113_Parts, LM_GI_l113b_Parts, LM_L_l114_Parts, LM_UL_l114_Parts, LM_GI_l114_Parts, LM_L_l115_Parts, LM_UL_l115_Parts, LM_UL_l18_Parts, LM_UL_l2_Parts, LM_L_l21_Parts, LM_L_l23_Parts, LM_UL_l23_Parts, LM_L_l24_Parts, LM_UL_l24_Parts, LM_UL_l25_Parts, LM_UL_l26_Parts, LM_L_l28_Parts, LM_L_l30_Parts, LM_UL_l30_Parts, LM_UL_l36_Parts, LM_L_l39_Parts, LM_UL_l39_Parts, LM_L_l40_Parts, LM_UL_l40_Parts, LM_L_l41_Parts, LM_UL_l41_Parts, LM_L_l43_Parts, LM_L_l44_Parts, LM_L_l47_Parts, LM_UL_l47_Parts, LM_L_l49rot_Parts, LM_L_l53_Parts, LM_UL_l55_Parts, LM_L_l56_Parts, LM_UL_l56_Parts, LM_L_l57_Parts, LM_UL_l57_Parts, _
  LM_L_l58_Parts, LM_UL_l58_Parts, LM_L_l59_Parts, LM_UL_l6_Parts, LM_UL_l62_Parts, LM_L_l65_Parts, LM_UL_l65_Parts, LM_GI_l65b_Parts, LM_L_l66_Parts, LM_UL_l66_Parts, LM_L_l67_Parts, LM_L_l8_Parts, LM_UL_l8_Parts, LM_L_l81_Parts, LM_UL_l81_Parts, LM_GI_l81b_Parts, LM_L_l82_Parts, LM_UL_l82_Parts, LM_GI_l83_Parts, LM_L_l9_Parts, LM_UL_l9_Parts, LM_L_l97_Parts, LM_UL_l97_Parts, LM_GI_l97b_Parts, LM_L_l98_Parts, LM_UL_l98_Parts, LM_GI_l98_Parts, LM_UL_l99_Parts)
Dim BP_Plane: BP_Plane=Array(BM_Plane, LM_D_aa00_Plane, LM_D_aa01_Plane, LM_D_aa02_Plane, LM_D_aa03_Plane, LM_D_aa04_Plane, LM_D_aa05_Plane, LM_D_aa06_Plane, LM_D_a10_Plane, LM_D_a11_Plane, LM_D_a12_Plane, LM_D_a13_Plane, LM_D_a14_Plane, LM_D_a15_Plane, LM_D_a16_Plane, LM_D_a20_Plane, LM_D_a21_Plane, LM_D_a22_Plane, LM_D_a23_Plane, LM_D_a24_Plane, LM_D_a25_Plane, LM_D_a26_Plane, LM_D_a30_Plane, LM_D_a31_Plane, LM_D_a32_Plane, LM_D_a33_Plane, LM_D_a34_Plane, LM_D_a35_Plane, LM_D_a36_Plane, LM_D_a40_Plane, LM_D_a41_Plane, LM_D_a42_Plane, LM_D_a43_Plane, LM_D_a44_Plane, LM_D_a45_Plane, LM_D_a46_Plane, LM_D_a50_Plane, LM_D_a51_Plane, LM_D_a52_Plane, LM_D_a53_Plane, LM_D_a54_Plane, LM_D_a55_Plane, LM_D_a56_Plane)
Dim BP_Playfield: BP_Playfield=Array(BM_Playfield, BM_Playfield_001, LM_GI_L12a_Playfield, LM_GI_L33a1_Playfield, LM_GI_L33a2_Playfield, LM_GI_L33a3_Playfield, LM_GI_L33a4_Playfield, LM_GI_L33a5_Playfield, LM_GI_L33a6_Playfield, LM_GI_L33a7_Playfield, LM_GI_L49a2_Playfield, LM_GI_L49a3_Playfield, LM_GI_L49a4_Playfield, LM_GI_L49a5_Playfield, LM_GI_L67b_Playfield, LM_UL_l10_Playfield, LM_GI_l113b_Playfield, LM_UL_l114_Playfield, LM_GI_l114_Playfield, LM_UL_l115_Playfield, LM_UL_l14_Playfield, LM_UL_l18_Playfield, LM_UL_l19_Playfield, LM_UL_l2_Playfield, LM_UL_l20_Playfield, LM_UL_l22_Playfield, LM_UL_l24_Playfield, LM_UL_l25_Playfield, LM_UL_l26_Playfield, LM_UL_l28_Playfield, LM_UL_l3_Playfield, LM_UL_l30_Playfield, LM_UL_l34_Playfield, LM_UL_l35_Playfield, LM_UL_l36_Playfield, LM_UL_l38_Playfield, LM_UL_l4_Playfield, LM_UL_l40_Playfield, LM_UL_l41_Playfield, LM_UL_l46_Playfield, LM_UL_l50_Playfield, LM_UL_l51_Playfield, LM_UL_l52_Playfield, LM_UL_l54_Playfield, LM_UL_l56_Playfield, LM_UL_l57_Playfield, _
  LM_UL_l58_Playfield, LM_UL_l59_Playfield, LM_UL_l6_Playfield, LM_UL_l60_Playfield, LM_UL_l62_Playfield, LM_GI_l65b_Playfield, LM_UL_l7_Playfield, LM_GI_l81b_Playfield, LM_GI_l83_Playfield, LM_UL_l9_Playfield, LM_L_l98_Playfield, LM_GI_l98_Playfield, LM_UL_l99_Playfield)
Dim BP_Prim14: BP_Prim14=Array(BM_Prim14, LM_L_l67_Prim14)
Dim BP_Prim2: BP_Prim2=Array(BM_Prim2, LM_GI_L33a2_Prim2, LM_GI_L33a5_Prim2, LM_UL_l25_Prim2, LM_UL_l41_Prim2, LM_L_l57_Prim2, LM_UL_l57_Prim2, LM_L_l9_Prim2, LM_UL_l9_Prim2)
Dim BP_Prim3: BP_Prim3=Array(BM_Prim3, LM_UL_l10_Prim3, LM_UL_l57_Prim3)
Dim BP_Prim4: BP_Prim4=Array(BM_Prim4, LM_GI_L33a2_Prim4)
Dim BP_Prim5: BP_Prim5=Array(BM_Prim5, LM_GI_L33a3_Prim5, LM_GI_L33a4_Prim5, LM_UL_l2_Prim5)
Dim BP_Prim6: BP_Prim6=Array(BM_Prim6, LM_GI_l113b_Prim6, LM_GI_l114_Prim6, LM_UL_l55_Prim6, LM_GI_l98_Prim6)
Dim BP_RSling: BP_RSling=Array(BM_RSling, LM_GI_L33a2_RSling, LM_GI_L33a5_RSling)
Dim BP_RSling1: BP_RSling1=Array(BM_RSling1, LM_GI_L33a2_RSling1, LM_GI_L33a5_RSling1, LM_L_l24_RSling1)
Dim BP_RSling2: BP_RSling2=Array(BM_RSling2, LM_GI_L33a2_RSling2, LM_GI_L33a5_RSling2)
Dim BP_RightFlip: BP_RightFlip=Array(BM_RightFlip, LM_GI_L33a2_RightFlip, LM_GI_L33a5_RightFlip, LM_GI_l83_RightFlip)
Dim BP_RightFlip1: BP_RightFlip1=Array(BM_RightFlip1, LM_GI_L12a_RightFlip1, LM_GI_L49a1_RightFlip1, LM_GI_L49a2_RightFlip1, LM_GI_L49a3_RightFlip1, LM_GI_l113b_RightFlip1, LM_UL_l31_RightFlip1, LM_GI_l97b_RightFlip1, LM_GI_l98_RightFlip1)
Dim BP_RightFlip1U: BP_RightFlip1U=Array(BM_RightFlip1U, LM_GI_L12a_RightFlip1U, LM_GI_L49a1_RightFlip1U, LM_GI_L49a2_RightFlip1U, LM_GI_L49a3_RightFlip1U, LM_GI_l113b_RightFlip1U, LM_GI_l114_RightFlip1U, LM_UL_l15_RightFlip1U, LM_L_l31_RightFlip1U, LM_UL_l31_RightFlip1U, LM_UL_l47_RightFlip1U, LM_GI_l97b_RightFlip1U, LM_GI_l98_RightFlip1U)
Dim BP_RightFlipU: BP_RightFlipU=Array(BM_RightFlipU, LM_GI_L33a2_RightFlipU, LM_GI_L33a5_RightFlipU, LM_UL_l20_RightFlipU, LM_L_l4_RightFlipU, LM_UL_l4_RightFlipU, LM_GI_l83_RightFlipU)
Dim BP_SaucerLWU: BP_SaucerLWU=Array(BM_SaucerLWU, LM_GI_L33a3_SaucerLWU)
Dim BP_SaucerRWU: BP_SaucerRWU=Array(BM_SaucerRWU)
Dim BP_UnderPF: BP_UnderPF=Array(BM_UnderPF, LM_GI_L33a1_UnderPF, LM_GI_L33a2_UnderPF, LM_GI_L33a5_UnderPF, LM_GI_L33a6_UnderPF, LM_L_l10_UnderPF, LM_UL_l10_UnderPF, LM_L_l115_UnderPF, LM_UL_l115_UnderPF, LM_UL_l14_UnderPF, LM_L_l14_UnderPF, LM_L_l18_UnderPF, LM_L_l19_UnderPF, LM_L_l2_UnderPF, LM_L_l20_UnderPF, LM_L_l22_UnderPF, LM_UL_l22_UnderPF, LM_L_l24_UnderPF, LM_L_l25_UnderPF, LM_UL_l25_UnderPF, LM_L_l26_UnderPF, LM_UL_l26_UnderPF, LM_UL_l28_UnderPF, LM_L_l28_UnderPF, LM_L_l3_UnderPF, LM_L_l30_UnderPF, LM_UL_l30_UnderPF, LM_L_l34_UnderPF, LM_L_l35_UnderPF, LM_L_l36_UnderPF, LM_L_l38_UnderPF, LM_UL_l38_UnderPF, LM_L_l4_UnderPF, LM_L_l40_UnderPF, LM_UL_l40_UnderPF, LM_L_l41_UnderPF, LM_UL_l41_UnderPF, LM_L_l42_UnderPF, LM_UL_l42_UnderPF, LM_L_l43_UnderPF, LM_UL_l43_UnderPF, LM_UL_l44_UnderPF, LM_L_l44_UnderPF, LM_L_l46_UnderPF, LM_UL_l46_UnderPF, LM_L_l49rot_UnderPF, LM_L_l50_UnderPF, LM_L_l51_UnderPF, LM_L_l52_UnderPF, LM_UL_l52_UnderPF, LM_L_l54_UnderPF, LM_UL_l54_UnderPF, LM_L_l56_UnderPF, _
  LM_L_l57_UnderPF, LM_UL_l57_UnderPF, LM_L_l58_UnderPF, LM_UL_l58_UnderPF, LM_L_l6_UnderPF, LM_UL_l6_UnderPF, LM_UL_l60_UnderPF, LM_L_l60_UnderPF, LM_L_l62_UnderPF, LM_UL_l62_UnderPF, LM_L_l67_UnderPF, LM_UL_l67_UnderPF, LM_L_l7_UnderPF, LM_UL_l7_UnderPF, LM_L_l8_UnderPF, LM_UL_l8_UnderPF, LM_L_l9_UnderPF, LM_UL_l9_UnderPF, LM_L_l99_UnderPF, LM_UL_l99_UnderPF)
Dim BP_UnderUPF: BP_UnderUPF=Array(BM_UnderUPF, LM_GI_L49a2_UnderUPF, LM_GI_L49a3_UnderUPF, LM_L_l113_UnderUPF, LM_UL_l113_UnderUPF, LM_GI_l113b_UnderUPF, LM_L_l114_UnderUPF, LM_UL_l114_UnderUPF, LM_GI_l114_UnderUPF, LM_L_l15_UnderUPF, LM_L_l21_UnderUPF, LM_UL_l21_UnderUPF, LM_L_l23_UnderUPF, LM_UL_l23_UnderUPF, LM_UL_l28_UnderUPF, LM_L_l28_UnderUPF, LM_L_l31_UnderUPF, LM_UL_l31_UnderUPF, LM_L_l37_UnderUPF, LM_UL_l37_UnderUPF, LM_L_l39_UnderUPF, LM_UL_l39_UnderUPF, LM_L_l47_UnderUPF, LM_UL_l47_UnderUPF, LM_L_l53_UnderUPF, LM_UL_l53_UnderUPF, LM_UL_l55_UnderUPF, LM_L_l65_UnderUPF, LM_UL_l65_UnderUPF, LM_GI_l65b_UnderUPF, LM_L_l66_UnderUPF, LM_UL_l66_UnderUPF, LM_L_l81_UnderUPF, LM_UL_l81_UnderUPF, LM_GI_l81b_UnderUPF, LM_L_l82_UnderUPF, LM_UL_l82_UnderUPF, LM_L_l97_UnderUPF, LM_UL_l97_UnderUPF, LM_GI_l97b_UnderUPF, LM_L_l98_UnderUPF, LM_UL_l98_UnderUPF, LM_GI_l98_UnderUPF)
Dim BP_UpperPF: BP_UpperPF=Array(BM_UpperPF, LM_GI_L12a_UpperPF, LM_GI_L49a1_UpperPF, LM_GI_L49a2_UpperPF, LM_GI_L49a3_UpperPF, LM_GI_L49a4_UpperPF, LM_GI_L49a5_UpperPF, LM_GI_L49a6_UpperPF, LM_GI_L67b_UpperPF, LM_GI_l113b_UpperPF, LM_UL_l114_UpperPF, LM_GI_l114_UpperPF, LM_UL_l15_UpperPF, LM_UL_l21_UpperPF, LM_UL_l23_UpperPF, LM_UL_l28_UpperPF, LM_UL_l31_UpperPF, LM_UL_l39_UpperPF, LM_UL_l47_UpperPF, LM_UL_l53_UpperPF, LM_UL_l55_UpperPF, LM_GI_l65b_UpperPF, LM_GI_l81b_UpperPF, LM_GI_l97b_UpperPF, LM_GI_l98_UpperPF)
Dim BP_UpperSaucer1WU: BP_UpperSaucer1WU=Array(BM_UpperSaucer1WU, LM_GI_L49a4_UpperSaucer1WU)
Dim BP_UpperSaucer2WU: BP_UpperSaucer2WU=Array(BM_UpperSaucer2WU)
Dim BP_UpperSaucer3DWU: BP_UpperSaucer3DWU=Array(BM_UpperSaucer3DWU)
Dim BP_inPFGlass: BP_inPFGlass=Array(BM_inPFGlass, LM_D_aa00_inPFGlass, LM_D_aa01_inPFGlass, LM_D_aa02_inPFGlass, LM_D_aa03_inPFGlass, LM_D_aa04_inPFGlass, LM_D_aa05_inPFGlass, LM_D_aa06_inPFGlass, LM_D_a10_inPFGlass, LM_D_a11_inPFGlass, LM_D_a12_inPFGlass, LM_D_a13_inPFGlass, LM_D_a14_inPFGlass, LM_D_a15_inPFGlass, LM_D_a16_inPFGlass, LM_D_a20_inPFGlass, LM_D_a21_inPFGlass, LM_D_a22_inPFGlass, LM_D_a23_inPFGlass, LM_D_a24_inPFGlass, LM_D_a25_inPFGlass, LM_D_a26_inPFGlass, LM_D_a30_inPFGlass, LM_D_a31_inPFGlass, LM_D_a32_inPFGlass, LM_D_a33_inPFGlass, LM_D_a34_inPFGlass, LM_D_a35_inPFGlass, LM_D_a36_inPFGlass, LM_D_a40_inPFGlass, LM_D_a41_inPFGlass, LM_D_a42_inPFGlass, LM_D_a43_inPFGlass, LM_D_a44_inPFGlass, LM_D_a45_inPFGlass, LM_D_a46_inPFGlass, LM_D_a50_inPFGlass, LM_D_a51_inPFGlass, LM_D_a52_inPFGlass, LM_D_a53_inPFGlass, LM_D_a54_inPFGlass, LM_D_a55_inPFGlass, LM_D_a56_inPFGlass)
Dim BP_sling1: BP_sling1=Array(BM_sling1, LM_GI_L33a2_sling1)
Dim BP_sling2: BP_sling2=Array(BM_sling2, LM_GI_L33a1_sling2, LM_UL_l30_sling2)
Dim BP_sw12: BP_sw12=Array(BM_sw12)
Dim BP_sw13: BP_sw13=Array(BM_sw13, LM_GI_l114_sw13)
Dim BP_sw21: BP_sw21=Array(BM_sw21, LM_UL_l10_sw21, LM_UL_l22_sw21, LM_UL_l6_sw21)
Dim BP_sw22: BP_sw22=Array(BM_sw22, LM_UL_l22_sw22, LM_UL_l38_sw22, LM_UL_l6_sw22)
Dim BP_sw23: BP_sw23=Array(BM_sw23, LM_UL_l22_sw23, LM_UL_l38_sw23, LM_UL_l54_sw23)
Dim BP_sw24: BP_sw24=Array(BM_sw24, LM_UL_l38_sw24, LM_UL_l54_sw24)
Dim BP_sw26: BP_sw26=Array(BM_sw26, LM_GI_L33a5_sw26, LM_UL_l28_sw26, LM_L_l28_sw26, LM_UL_l58_sw26)
Dim BP_sw27: BP_sw27=Array(BM_sw27, LM_GI_L33a5_sw27, LM_UL_l44_sw27, LM_L_l44_sw27)
Dim BP_sw28: BP_sw28=Array(BM_sw28, LM_GI_L33a6_sw28, LM_UL_l60_sw28, LM_L_l60_sw28, LM_UL_l99_sw28)
Dim BP_sw29: BP_sw29=Array(BM_sw29)
Dim BP_sw30: BP_sw30=Array(BM_sw30, LM_L_l67_sw30, LM_UL_l67_sw30)
Dim BP_sw31: BP_sw31=Array(BM_sw31)
Dim BP_sw33: BP_sw33=Array(BM_sw33, LM_GI_L33a5_sw33)
Dim BP_sw34: BP_sw34=Array(BM_sw34, LM_GI_L33a2_sw34)
Dim BP_sw35: BP_sw35=Array(BM_sw35, LM_GI_L33a1_sw35, LM_GI_L33a6_sw35)
Dim BP_sw36: BP_sw36=Array(BM_sw36, LM_GI_L33a1_sw36, LM_GI_L33a6_sw36)
Dim BP_sw37_001: BP_sw37_001=Array(BM_sw37_001, LM_GI_L49a2_sw37_001, LM_L_l49rot_sw37_001)
Dim BP_sw46: BP_sw46=Array(BM_sw46, LM_GI_L12a_sw46, LM_GI_L49a1_sw46, LM_GI_L49a2_sw46, LM_GI_L49a3_sw46, LM_GI_l113b_sw46, LM_GI_l114_sw46, LM_GI_l65b_sw46, LM_GI_l81b_sw46, LM_GI_l97b_sw46)
Dim BP_sw47: BP_sw47=Array(BM_sw47, LM_GI_L12a_sw47, LM_GI_L49a1_sw47, LM_GI_L49a2_sw47, LM_GI_L49a3_sw47, LM_GI_L67b_sw47, LM_GI_l113b_sw47, LM_GI_l65b_sw47, LM_GI_l81b_sw47, LM_GI_l97b_sw47)
Dim BP_sw48: BP_sw48=Array(BM_sw48, LM_GI_L12a_sw48, LM_GI_L49a1_sw48, LM_GI_L49a2_sw48, LM_GI_L49a3_sw48, LM_GI_L49a5_sw48, LM_GI_L67b_sw48, LM_GI_l113b_sw48, LM_GI_l65b_sw48, LM_GI_l81b_sw48, LM_GI_l97b_sw48)
' Arrays per lighting scenario
Dim BL_D_a10: BL_D_a10=Array(LM_D_a10_Plane, LM_D_a10_inPFGlass)
Dim BL_D_a11: BL_D_a11=Array(LM_D_a11_Plane, LM_D_a11_inPFGlass)
Dim BL_D_a12: BL_D_a12=Array(LM_D_a12_Plane, LM_D_a12_inPFGlass)
Dim BL_D_a13: BL_D_a13=Array(LM_D_a13_Plane, LM_D_a13_inPFGlass)
Dim BL_D_a14: BL_D_a14=Array(LM_D_a14_Plane, LM_D_a14_inPFGlass)
Dim BL_D_a15: BL_D_a15=Array(LM_D_a15_Plane, LM_D_a15_inPFGlass)
Dim BL_D_a16: BL_D_a16=Array(LM_D_a16_Plane, LM_D_a16_inPFGlass)
Dim BL_D_a20: BL_D_a20=Array(LM_D_a20_Plane, LM_D_a20_inPFGlass)
Dim BL_D_a21: BL_D_a21=Array(LM_D_a21_Plane, LM_D_a21_inPFGlass)
Dim BL_D_a22: BL_D_a22=Array(LM_D_a22_Plane, LM_D_a22_inPFGlass)
Dim BL_D_a23: BL_D_a23=Array(LM_D_a23_Plane, LM_D_a23_inPFGlass)
Dim BL_D_a24: BL_D_a24=Array(LM_D_a24_Plane, LM_D_a24_inPFGlass)
Dim BL_D_a25: BL_D_a25=Array(LM_D_a25_Plane, LM_D_a25_inPFGlass)
Dim BL_D_a26: BL_D_a26=Array(LM_D_a26_Plane, LM_D_a26_inPFGlass)
Dim BL_D_a30: BL_D_a30=Array(LM_D_a30_Plane, LM_D_a30_inPFGlass)
Dim BL_D_a31: BL_D_a31=Array(LM_D_a31_Plane, LM_D_a31_inPFGlass)
Dim BL_D_a32: BL_D_a32=Array(LM_D_a32_Plane, LM_D_a32_inPFGlass)
Dim BL_D_a33: BL_D_a33=Array(LM_D_a33_Plane, LM_D_a33_inPFGlass)
Dim BL_D_a34: BL_D_a34=Array(LM_D_a34_Plane, LM_D_a34_inPFGlass)
Dim BL_D_a35: BL_D_a35=Array(LM_D_a35_Plane, LM_D_a35_inPFGlass)
Dim BL_D_a36: BL_D_a36=Array(LM_D_a36_Plane, LM_D_a36_inPFGlass)
Dim BL_D_a40: BL_D_a40=Array(LM_D_a40_Plane, LM_D_a40_inPFGlass)
Dim BL_D_a41: BL_D_a41=Array(LM_D_a41_Plane, LM_D_a41_inPFGlass)
Dim BL_D_a42: BL_D_a42=Array(LM_D_a42_Plane, LM_D_a42_inPFGlass)
Dim BL_D_a43: BL_D_a43=Array(LM_D_a43_Plane, LM_D_a43_inPFGlass)
Dim BL_D_a44: BL_D_a44=Array(LM_D_a44_Plane, LM_D_a44_inPFGlass)
Dim BL_D_a45: BL_D_a45=Array(LM_D_a45_Plane, LM_D_a45_inPFGlass)
Dim BL_D_a46: BL_D_a46=Array(LM_D_a46_Plane, LM_D_a46_inPFGlass)
Dim BL_D_a50: BL_D_a50=Array(LM_D_a50_Plane, LM_D_a50_inPFGlass)
Dim BL_D_a51: BL_D_a51=Array(LM_D_a51_Plane, LM_D_a51_inPFGlass)
Dim BL_D_a52: BL_D_a52=Array(LM_D_a52_Plane, LM_D_a52_inPFGlass)
Dim BL_D_a53: BL_D_a53=Array(LM_D_a53_Plane, LM_D_a53_inPFGlass)
Dim BL_D_a54: BL_D_a54=Array(LM_D_a54_Plane, LM_D_a54_inPFGlass)
Dim BL_D_a55: BL_D_a55=Array(LM_D_a55_Plane, LM_D_a55_inPFGlass)
Dim BL_D_a56: BL_D_a56=Array(LM_D_a56_Plane, LM_D_a56_inPFGlass)
Dim BL_D_aa00: BL_D_aa00=Array(LM_D_aa00_Plane, LM_D_aa00_inPFGlass)
Dim BL_D_aa01: BL_D_aa01=Array(LM_D_aa01_Plane, LM_D_aa01_inPFGlass)
Dim BL_D_aa02: BL_D_aa02=Array(LM_D_aa02_Plane, LM_D_aa02_inPFGlass)
Dim BL_D_aa03: BL_D_aa03=Array(LM_D_aa03_Plane, LM_D_aa03_inPFGlass)
Dim BL_D_aa04: BL_D_aa04=Array(LM_D_aa04_Plane, LM_D_aa04_inPFGlass)
Dim BL_D_aa05: BL_D_aa05=Array(LM_D_aa05_Plane, LM_D_aa05_inPFGlass)
Dim BL_D_aa06: BL_D_aa06=Array(LM_D_aa06_Plane, LM_D_aa06_inPFGlass)
Dim BL_GI_CreditLight: BL_GI_CreditLight=Array(LM_GI_CreditLight_Layer_1)
Dim BL_GI_L12a: BL_GI_L12a=Array(LM_GI_L12a_Bumper1_Ring, LM_GI_L12a_Layer_1, LM_GI_L12a_Parts, LM_GI_L12a_Playfield, LM_GI_L12a_RightFlip1, LM_GI_L12a_RightFlip1U, LM_GI_L12a_UpperPF, LM_GI_L12a_sw46, LM_GI_L12a_sw47, LM_GI_L12a_sw48)
Dim BL_GI_L33a1: BL_GI_L33a1=Array(LM_GI_L33a1_LSling, LM_GI_L33a1_LSling1, LM_GI_L33a1_LSling2, LM_GI_L33a1_Layer_1, LM_GI_L33a1_LeftFlip, LM_GI_L33a1_LeftFlip1, LM_GI_L33a1_LeftFlip1U, LM_GI_L33a1_LeftFlipU, LM_GI_L33a1_Parts, LM_GI_L33a1_Playfield, LM_GI_L33a1_UnderPF, LM_GI_L33a1_sling2, LM_GI_L33a1_sw35, LM_GI_L33a1_sw36)
Dim BL_GI_L33a2: BL_GI_L33a2=Array(LM_GI_L33a2_Layer_1, LM_GI_L33a2_LeftFlipU, LM_GI_L33a2_Parts, LM_GI_L33a2_Playfield, LM_GI_L33a2_Prim2, LM_GI_L33a2_Prim4, LM_GI_L33a2_RSling, LM_GI_L33a2_RSling1, LM_GI_L33a2_RSling2, LM_GI_L33a2_RightFlip, LM_GI_L33a2_RightFlipU, LM_GI_L33a2_UnderPF, LM_GI_L33a2_sling1, LM_GI_L33a2_sw34)
Dim BL_GI_L33a3: BL_GI_L33a3=Array(LM_GI_L33a3_Layer_1, LM_GI_L33a3_LeftFlipU, LM_GI_L33a3_Parts, LM_GI_L33a3_Playfield, LM_GI_L33a3_Prim5, LM_GI_L33a3_SaucerLWU)
Dim BL_GI_L33a4: BL_GI_L33a4=Array(LM_GI_L33a4_LSling, LM_GI_L33a4_LSling1, LM_GI_L33a4_LSling2, LM_GI_L33a4_Layer_1, LM_GI_L33a4_LeftFlip, LM_GI_L33a4_LeftFlipU, LM_GI_L33a4_Parts, LM_GI_L33a4_Playfield, LM_GI_L33a4_Prim5)
Dim BL_GI_L33a5: BL_GI_L33a5=Array(LM_GI_L33a5_Layer_1, LM_GI_L33a5_LeftFlip1U, LM_GI_L33a5_LeftFlipU, LM_GI_L33a5_Parts, LM_GI_L33a5_Playfield, LM_GI_L33a5_Prim2, LM_GI_L33a5_RSling, LM_GI_L33a5_RSling1, LM_GI_L33a5_RSling2, LM_GI_L33a5_RightFlip, LM_GI_L33a5_RightFlipU, LM_GI_L33a5_UnderPF, LM_GI_L33a5_sw26, LM_GI_L33a5_sw27, LM_GI_L33a5_sw33)
Dim BL_GI_L33a6: BL_GI_L33a6=Array(LM_GI_L33a6_LSling, LM_GI_L33a6_LSling1, LM_GI_L33a6_LSling2, LM_GI_L33a6_Layer_1, LM_GI_L33a6_LeftFlip, LM_GI_L33a6_LeftFlipU, LM_GI_L33a6_Parts, LM_GI_L33a6_Playfield, LM_GI_L33a6_UnderPF, LM_GI_L33a6_sw28, LM_GI_L33a6_sw35, LM_GI_L33a6_sw36)
Dim BL_GI_L33a7: BL_GI_L33a7=Array(LM_GI_L33a7_Layer_1, LM_GI_L33a7_Parts, LM_GI_L33a7_Playfield)
Dim BL_GI_L49a1: BL_GI_L49a1=Array(LM_GI_L49a1_Layer_1, LM_GI_L49a1_Parts, LM_GI_L49a1_RightFlip1, LM_GI_L49a1_RightFlip1U, LM_GI_L49a1_UpperPF, LM_GI_L49a1_sw46, LM_GI_L49a1_sw47, LM_GI_L49a1_sw48)
Dim BL_GI_L49a2: BL_GI_L49a2=Array(LM_GI_L49a2_Bumper1_Ring, LM_GI_L49a2_Layer_1, LM_GI_L49a2_Parts, LM_GI_L49a2_Playfield, LM_GI_L49a2_RightFlip1, LM_GI_L49a2_RightFlip1U, LM_GI_L49a2_UnderUPF, LM_GI_L49a2_UpperPF, LM_GI_L49a2_sw37_001, LM_GI_L49a2_sw46, LM_GI_L49a2_sw47, LM_GI_L49a2_sw48)
Dim BL_GI_L49a3: BL_GI_L49a3=Array(LM_GI_L49a3_Bumper1_Ring, LM_GI_L49a3_Layer_1, LM_GI_L49a3_Parts, LM_GI_L49a3_Playfield, LM_GI_L49a3_RightFlip1, LM_GI_L49a3_RightFlip1U, LM_GI_L49a3_UnderUPF, LM_GI_L49a3_UpperPF, LM_GI_L49a3_sw46, LM_GI_L49a3_sw47, LM_GI_L49a3_sw48)
Dim BL_GI_L49a4: BL_GI_L49a4=Array(LM_GI_L49a4_Layer_1, LM_GI_L49a4_Parts, LM_GI_L49a4_Playfield, LM_GI_L49a4_UpperPF, LM_GI_L49a4_UpperSaucer1WU)
Dim BL_GI_L49a5: BL_GI_L49a5=Array(LM_GI_L49a5_Layer_1, LM_GI_L49a5_Parts, LM_GI_L49a5_Playfield, LM_GI_L49a5_UpperPF, LM_GI_L49a5_sw48)
Dim BL_GI_L49a6: BL_GI_L49a6=Array(LM_GI_L49a6_Layer_1, LM_GI_L49a6_Parts, LM_GI_L49a6_UpperPF)
Dim BL_GI_L67b: BL_GI_L67b=Array(LM_GI_L67b_Layer_1, LM_GI_L67b_Parts, LM_GI_L67b_Playfield, LM_GI_L67b_UpperPF, LM_GI_L67b_sw47, LM_GI_L67b_sw48)
Dim BL_GI_l113b: BL_GI_l113b=Array(LM_GI_l113b_Bumper1_Ring, LM_GI_l113b_Layer_1, LM_GI_l113b_Parts, LM_GI_l113b_Playfield, LM_GI_l113b_Prim6, LM_GI_l113b_RightFlip1, LM_GI_l113b_RightFlip1U, LM_GI_l113b_UnderUPF, LM_GI_l113b_UpperPF, LM_GI_l113b_sw46, LM_GI_l113b_sw47, LM_GI_l113b_sw48)
Dim BL_GI_l114: BL_GI_l114=Array(LM_GI_l114_Bumper1_Ring, LM_GI_l114_Layer_1, LM_GI_l114_Parts, LM_GI_l114_Playfield, LM_GI_l114_Prim6, LM_GI_l114_RightFlip1U, LM_GI_l114_UnderUPF, LM_GI_l114_UpperPF, LM_GI_l114_sw13, LM_GI_l114_sw46)
Dim BL_GI_l65b: BL_GI_l65b=Array(LM_GI_l65b_Layer_1, LM_GI_l65b_Parts, LM_GI_l65b_Playfield, LM_GI_l65b_UnderUPF, LM_GI_l65b_UpperPF, LM_GI_l65b_sw46, LM_GI_l65b_sw47, LM_GI_l65b_sw48)
Dim BL_GI_l81b: BL_GI_l81b=Array(LM_GI_l81b_Layer_1, LM_GI_l81b_Parts, LM_GI_l81b_Playfield, LM_GI_l81b_UnderUPF, LM_GI_l81b_UpperPF, LM_GI_l81b_sw46, LM_GI_l81b_sw47, LM_GI_l81b_sw48)
Dim BL_GI_l83: BL_GI_l83=Array(LM_GI_l83_Gate1, LM_GI_l83_LeftFlipU, LM_GI_l83_Parts, LM_GI_l83_Playfield, LM_GI_l83_RightFlip, LM_GI_l83_RightFlipU)
Dim BL_GI_l97b: BL_GI_l97b=Array(LM_GI_l97b_Bumper1_Ring, LM_GI_l97b_Layer_1, LM_GI_l97b_Parts, LM_GI_l97b_RightFlip1, LM_GI_l97b_RightFlip1U, LM_GI_l97b_UnderUPF, LM_GI_l97b_UpperPF, LM_GI_l97b_sw46, LM_GI_l97b_sw47, LM_GI_l97b_sw48)
Dim BL_GI_l98: BL_GI_l98=Array(LM_GI_l98_Layer_1, LM_GI_l98_Parts, LM_GI_l98_Playfield, LM_GI_l98_Prim6, LM_GI_l98_RightFlip1, LM_GI_l98_RightFlip1U, LM_GI_l98_UnderUPF, LM_GI_l98_UpperPF)
Dim BL_L_l10: BL_L_l10=Array(LM_L_l10_Layer_1, LM_L_l10_Parts, LM_L_l10_UnderPF)
Dim BL_L_l113: BL_L_l113=Array(LM_L_l113_Layer_1, LM_L_l113_Parts, LM_L_l113_UnderUPF)
Dim BL_L_l114: BL_L_l114=Array(LM_L_l114_Layer_1, LM_L_l114_Parts, LM_L_l114_UnderUPF)
Dim BL_L_l115: BL_L_l115=Array(LM_L_l115_Layer_1, LM_L_l115_Parts, LM_L_l115_UnderPF)
Dim BL_L_l14: BL_L_l14=Array(LM_L_l14_UnderPF)
Dim BL_L_l15: BL_L_l15=Array(LM_L_l15_UnderUPF)
Dim BL_L_l18: BL_L_l18=Array(LM_L_l18_UnderPF)
Dim BL_L_l19: BL_L_l19=Array(LM_L_l19_UnderPF)
Dim BL_L_l2: BL_L_l2=Array(LM_L_l2_UnderPF)
Dim BL_L_l20: BL_L_l20=Array(LM_L_l20_UnderPF)
Dim BL_L_l21: BL_L_l21=Array(LM_L_l21_Parts, LM_L_l21_UnderUPF)
Dim BL_L_l22: BL_L_l22=Array(LM_L_l22_UnderPF)
Dim BL_L_l23: BL_L_l23=Array(LM_L_l23_Layer_1, LM_L_l23_Parts, LM_L_l23_UnderUPF)
Dim BL_L_l24: BL_L_l24=Array(LM_L_l24_Parts, LM_L_l24_RSling1, LM_L_l24_UnderPF)
Dim BL_L_l25: BL_L_l25=Array(LM_L_l25_UnderPF)
Dim BL_L_l26: BL_L_l26=Array(LM_L_l26_UnderPF)
Dim BL_L_l28: BL_L_l28=Array(LM_L_l28_Parts, LM_L_l28_UnderPF, LM_L_l28_UnderUPF, LM_L_l28_sw26)
Dim BL_L_l3: BL_L_l3=Array(LM_L_l3_LeftFlipU, LM_L_l3_UnderPF)
Dim BL_L_l30: BL_L_l30=Array(LM_L_l30_LSling, LM_L_l30_LSling1, LM_L_l30_LSling2, LM_L_l30_Layer_1, LM_L_l30_Parts, LM_L_l30_UnderPF)
Dim BL_L_l31: BL_L_l31=Array(LM_L_l31_RightFlip1U, LM_L_l31_UnderUPF)
Dim BL_L_l34: BL_L_l34=Array(LM_L_l34_UnderPF)
Dim BL_L_l35: BL_L_l35=Array(LM_L_l35_UnderPF)
Dim BL_L_l36: BL_L_l36=Array(LM_L_l36_UnderPF)
Dim BL_L_l37: BL_L_l37=Array(LM_L_l37_UnderUPF)
Dim BL_L_l38: BL_L_l38=Array(LM_L_l38_UnderPF)
Dim BL_L_l39: BL_L_l39=Array(LM_L_l39_Layer_1, LM_L_l39_Parts, LM_L_l39_UnderUPF)
Dim BL_L_l4: BL_L_l4=Array(LM_L_l4_RightFlipU, LM_L_l4_UnderPF)
Dim BL_L_l40: BL_L_l40=Array(LM_L_l40_LSling, LM_L_l40_LSling1, LM_L_l40_LSling2, LM_L_l40_Layer_1, LM_L_l40_Parts, LM_L_l40_UnderPF)
Dim BL_L_l41: BL_L_l41=Array(LM_L_l41_Parts, LM_L_l41_UnderPF)
Dim BL_L_l42: BL_L_l42=Array(LM_L_l42_UnderPF)
Dim BL_L_l43: BL_L_l43=Array(LM_L_l43_Layer_1, LM_L_l43_Parts, LM_L_l43_UnderPF)
Dim BL_L_l44: BL_L_l44=Array(LM_L_l44_Parts, LM_L_l44_UnderPF, LM_L_l44_sw27)
Dim BL_L_l46: BL_L_l46=Array(LM_L_l46_UnderPF)
Dim BL_L_l47: BL_L_l47=Array(LM_L_l47_Layer_1, LM_L_l47_Parts, LM_L_l47_UnderUPF)
Dim BL_L_l49rot: BL_L_l49rot=Array(LM_L_l49rot_Parts, LM_L_l49rot_UnderPF, LM_L_l49rot_sw37_001)
Dim BL_L_l50: BL_L_l50=Array(LM_L_l50_UnderPF)
Dim BL_L_l51: BL_L_l51=Array(LM_L_l51_UnderPF)
Dim BL_L_l52: BL_L_l52=Array(LM_L_l52_UnderPF)
Dim BL_L_l53: BL_L_l53=Array(LM_L_l53_Parts, LM_L_l53_UnderUPF)
Dim BL_L_l54: BL_L_l54=Array(LM_L_l54_UnderPF)
Dim BL_L_l56: BL_L_l56=Array(LM_L_l56_Parts, LM_L_l56_UnderPF)
Dim BL_L_l57: BL_L_l57=Array(LM_L_l57_Parts, LM_L_l57_Prim2, LM_L_l57_UnderPF)
Dim BL_L_l58: BL_L_l58=Array(LM_L_l58_Parts, LM_L_l58_UnderPF)
Dim BL_L_l59: BL_L_l59=Array(LM_L_l59_Parts)
Dim BL_L_l6: BL_L_l6=Array(LM_L_l6_UnderPF)
Dim BL_L_l60: BL_L_l60=Array(LM_L_l60_LeftFlip1U, LM_L_l60_UnderPF, LM_L_l60_sw28)
Dim BL_L_l62: BL_L_l62=Array(LM_L_l62_UnderPF)
Dim BL_L_l65: BL_L_l65=Array(LM_L_l65_Layer_1, LM_L_l65_Parts, LM_L_l65_UnderUPF)
Dim BL_L_l66: BL_L_l66=Array(LM_L_l66_Parts, LM_L_l66_UnderUPF)
Dim BL_L_l67: BL_L_l67=Array(LM_L_l67_Parts, LM_L_l67_Prim14, LM_L_l67_UnderPF, LM_L_l67_sw30)
Dim BL_L_l7: BL_L_l7=Array(LM_L_l7_UnderPF)
Dim BL_L_l8: BL_L_l8=Array(LM_L_l8_Layer_1, LM_L_l8_Parts, LM_L_l8_UnderPF)
Dim BL_L_l81: BL_L_l81=Array(LM_L_l81_Layer_1, LM_L_l81_Parts, LM_L_l81_UnderUPF)
Dim BL_L_l82: BL_L_l82=Array(LM_L_l82_Parts, LM_L_l82_UnderUPF)
Dim BL_L_l9: BL_L_l9=Array(LM_L_l9_Parts, LM_L_l9_Prim2, LM_L_l9_UnderPF)
Dim BL_L_l97: BL_L_l97=Array(LM_L_l97_Layer_1, LM_L_l97_Parts, LM_L_l97_UnderUPF)
Dim BL_L_l98: BL_L_l98=Array(LM_L_l98_Layer_1, LM_L_l98_Parts, LM_L_l98_Playfield, LM_L_l98_UnderUPF)
Dim BL_L_l99: BL_L_l99=Array(LM_L_l99_UnderPF)
Dim BL_UL_l10: BL_UL_l10=Array(LM_UL_l10_Layer_1, LM_UL_l10_Parts, LM_UL_l10_Playfield, LM_UL_l10_Prim3, LM_UL_l10_UnderPF, LM_UL_l10_sw21)
Dim BL_UL_l113: BL_UL_l113=Array(LM_UL_l113_Parts, LM_UL_l113_UnderUPF)
Dim BL_UL_l114: BL_UL_l114=Array(LM_UL_l114_Layer_1, LM_UL_l114_Parts, LM_UL_l114_Playfield, LM_UL_l114_UnderUPF, LM_UL_l114_UpperPF)
Dim BL_UL_l115: BL_UL_l115=Array(LM_UL_l115_Layer_1, LM_UL_l115_LeftFlip1, LM_UL_l115_Parts, LM_UL_l115_Playfield, LM_UL_l115_UnderPF)
Dim BL_UL_l14: BL_UL_l14=Array(LM_UL_l14_Playfield, LM_UL_l14_UnderPF)
Dim BL_UL_l15: BL_UL_l15=Array(LM_UL_l15_RightFlip1U, LM_UL_l15_UpperPF)
Dim BL_UL_l18: BL_UL_l18=Array(LM_UL_l18_Parts, LM_UL_l18_Playfield)
Dim BL_UL_l19: BL_UL_l19=Array(LM_UL_l19_LeftFlipU, LM_UL_l19_Playfield)
Dim BL_UL_l2: BL_UL_l2=Array(LM_UL_l2_LSling, LM_UL_l2_LSling1, LM_UL_l2_LSling2, LM_UL_l2_Parts, LM_UL_l2_Playfield, LM_UL_l2_Prim5)
Dim BL_UL_l20: BL_UL_l20=Array(LM_UL_l20_Playfield, LM_UL_l20_RightFlipU)
Dim BL_UL_l21: BL_UL_l21=Array(LM_UL_l21_UnderUPF, LM_UL_l21_UpperPF)
Dim BL_UL_l22: BL_UL_l22=Array(LM_UL_l22_Playfield, LM_UL_l22_UnderPF, LM_UL_l22_sw21, LM_UL_l22_sw22, LM_UL_l22_sw23)
Dim BL_UL_l23: BL_UL_l23=Array(LM_UL_l23_Layer_1, LM_UL_l23_Parts, LM_UL_l23_UnderUPF, LM_UL_l23_UpperPF)
Dim BL_UL_l24: BL_UL_l24=Array(LM_UL_l24_Parts, LM_UL_l24_Playfield)
Dim BL_UL_l25: BL_UL_l25=Array(LM_UL_l25_Parts, LM_UL_l25_Playfield, LM_UL_l25_Prim2, LM_UL_l25_UnderPF)
Dim BL_UL_l26: BL_UL_l26=Array(LM_UL_l26_Parts, LM_UL_l26_Playfield, LM_UL_l26_UnderPF)
Dim BL_UL_l28: BL_UL_l28=Array(LM_UL_l28_Playfield, LM_UL_l28_UnderPF, LM_UL_l28_UnderUPF, LM_UL_l28_UpperPF, LM_UL_l28_sw26)
Dim BL_UL_l3: BL_UL_l3=Array(LM_UL_l3_LeftFlipU, LM_UL_l3_Playfield)
Dim BL_UL_l30: BL_UL_l30=Array(LM_UL_l30_LSling, LM_UL_l30_LSling1, LM_UL_l30_LSling2, LM_UL_l30_Parts, LM_UL_l30_Playfield, LM_UL_l30_UnderPF, LM_UL_l30_sling2)
Dim BL_UL_l31: BL_UL_l31=Array(LM_UL_l31_RightFlip1, LM_UL_l31_RightFlip1U, LM_UL_l31_UnderUPF, LM_UL_l31_UpperPF)
Dim BL_UL_l34: BL_UL_l34=Array(LM_UL_l34_Playfield)
Dim BL_UL_l35: BL_UL_l35=Array(LM_UL_l35_Playfield)
Dim BL_UL_l36: BL_UL_l36=Array(LM_UL_l36_Parts, LM_UL_l36_Playfield)
Dim BL_UL_l37: BL_UL_l37=Array(LM_UL_l37_UnderUPF)
Dim BL_UL_l38: BL_UL_l38=Array(LM_UL_l38_Playfield, LM_UL_l38_UnderPF, LM_UL_l38_sw22, LM_UL_l38_sw23, LM_UL_l38_sw24)
Dim BL_UL_l39: BL_UL_l39=Array(LM_UL_l39_Layer_1, LM_UL_l39_Parts, LM_UL_l39_UnderUPF, LM_UL_l39_UpperPF)
Dim BL_UL_l4: BL_UL_l4=Array(LM_UL_l4_Playfield, LM_UL_l4_RightFlipU)
Dim BL_UL_l40: BL_UL_l40=Array(LM_UL_l40_LSling, LM_UL_l40_LSling1, LM_UL_l40_LSling2, LM_UL_l40_Parts, LM_UL_l40_Playfield, LM_UL_l40_UnderPF)
Dim BL_UL_l41: BL_UL_l41=Array(LM_UL_l41_Parts, LM_UL_l41_Playfield, LM_UL_l41_Prim2, LM_UL_l41_UnderPF)
Dim BL_UL_l42: BL_UL_l42=Array(LM_UL_l42_UnderPF)
Dim BL_UL_l43: BL_UL_l43=Array(LM_UL_l43_UnderPF)
Dim BL_UL_l44: BL_UL_l44=Array(LM_UL_l44_UnderPF, LM_UL_l44_sw27)
Dim BL_UL_l46: BL_UL_l46=Array(LM_UL_l46_Playfield, LM_UL_l46_UnderPF)
Dim BL_UL_l47: BL_UL_l47=Array(LM_UL_l47_Layer_1, LM_UL_l47_Parts, LM_UL_l47_RightFlip1U, LM_UL_l47_UnderUPF, LM_UL_l47_UpperPF)
Dim BL_UL_l50: BL_UL_l50=Array(LM_UL_l50_Playfield)
Dim BL_UL_l51: BL_UL_l51=Array(LM_UL_l51_Playfield)
Dim BL_UL_l52: BL_UL_l52=Array(LM_UL_l52_Playfield, LM_UL_l52_UnderPF)
Dim BL_UL_l53: BL_UL_l53=Array(LM_UL_l53_UnderUPF, LM_UL_l53_UpperPF)
Dim BL_UL_l54: BL_UL_l54=Array(LM_UL_l54_Playfield, LM_UL_l54_UnderPF, LM_UL_l54_sw23, LM_UL_l54_sw24)
Dim BL_UL_l55: BL_UL_l55=Array(LM_UL_l55_Layer_1, LM_UL_l55_Parts, LM_UL_l55_Prim6, LM_UL_l55_UnderUPF, LM_UL_l55_UpperPF)
Dim BL_UL_l56: BL_UL_l56=Array(LM_UL_l56_Layer_1, LM_UL_l56_Parts, LM_UL_l56_Playfield)
Dim BL_UL_l57: BL_UL_l57=Array(LM_UL_l57_Layer_1, LM_UL_l57_Parts, LM_UL_l57_Playfield, LM_UL_l57_Prim2, LM_UL_l57_Prim3, LM_UL_l57_UnderPF)
Dim BL_UL_l58: BL_UL_l58=Array(LM_UL_l58_Parts, LM_UL_l58_Playfield, LM_UL_l58_UnderPF, LM_UL_l58_sw26)
Dim BL_UL_l59: BL_UL_l59=Array(LM_UL_l59_Playfield)
Dim BL_UL_l6: BL_UL_l6=Array(LM_UL_l6_Parts, LM_UL_l6_Playfield, LM_UL_l6_UnderPF, LM_UL_l6_sw21, LM_UL_l6_sw22)
Dim BL_UL_l60: BL_UL_l60=Array(LM_UL_l60_LeftFlip1U, LM_UL_l60_Playfield, LM_UL_l60_UnderPF, LM_UL_l60_sw28)
Dim BL_UL_l62: BL_UL_l62=Array(LM_UL_l62_Parts, LM_UL_l62_Playfield, LM_UL_l62_UnderPF)
Dim BL_UL_l65: BL_UL_l65=Array(LM_UL_l65_Layer_1, LM_UL_l65_Parts, LM_UL_l65_UnderUPF)
Dim BL_UL_l66: BL_UL_l66=Array(LM_UL_l66_Parts, LM_UL_l66_UnderUPF)
Dim BL_UL_l67: BL_UL_l67=Array(LM_UL_l67_UnderPF, LM_UL_l67_sw30)
Dim BL_UL_l7: BL_UL_l7=Array(LM_UL_l7_Playfield, LM_UL_l7_UnderPF)
Dim BL_UL_l8: BL_UL_l8=Array(LM_UL_l8_Parts, LM_UL_l8_UnderPF)
Dim BL_UL_l81: BL_UL_l81=Array(LM_UL_l81_Layer_1, LM_UL_l81_Parts, LM_UL_l81_UnderUPF)
Dim BL_UL_l82: BL_UL_l82=Array(LM_UL_l82_Layer_1, LM_UL_l82_Parts, LM_UL_l82_UnderUPF)
Dim BL_UL_l9: BL_UL_l9=Array(LM_UL_l9_Layer_1, LM_UL_l9_Parts, LM_UL_l9_Playfield, LM_UL_l9_Prim2, LM_UL_l9_UnderPF)
Dim BL_UL_l97: BL_UL_l97=Array(LM_UL_l97_Parts, LM_UL_l97_UnderUPF)
Dim BL_UL_l98: BL_UL_l98=Array(LM_UL_l98_Layer_1, LM_UL_l98_Parts, LM_UL_l98_UnderUPF)
Dim BL_UL_l99: BL_UL_l99=Array(LM_UL_l99_LeftFlip1, LM_UL_l99_Parts, LM_UL_l99_Playfield, LM_UL_l99_UnderPF, LM_UL_l99_sw28)
Dim BL_World: BL_World=Array(BM_Bumper1_Ring, BM_Cube, BM_Gate1, BM_LSling, BM_LSling1, BM_LSling2, BM_Layer_1, BM_LeftFlip, BM_LeftFlip1, BM_LeftFlip1U, BM_LeftFlipU, BM_Parts, BM_Parts_001, BM_Parts_002, BM_Plane, BM_Playfield, BM_Playfield_001, BM_Prim14, BM_Prim2, BM_Prim3, BM_Prim4, BM_Prim5, BM_Prim6, BM_RSling, BM_RSling1, BM_RSling2, BM_RightFlip, BM_RightFlip1, BM_RightFlip1U, BM_RightFlipU, BM_SaucerLWU, BM_SaucerRWU, BM_UnderPF, BM_UnderUPF, BM_UpperPF, BM_UpperSaucer1WU, BM_UpperSaucer2WU, BM_UpperSaucer3DWU, BM_inPFGlass, BM_sling1, BM_sling2, BM_sw12, BM_sw13, BM_sw21, BM_sw22, BM_sw23, BM_sw24, BM_sw26, BM_sw27, BM_sw28, BM_sw29, BM_sw30, BM_sw31, BM_sw33, BM_sw34, BM_sw35, BM_sw36, BM_sw37_001, BM_sw46, BM_sw47, BM_sw48)
' Global arrays
Dim BG_Bakemap: BG_Bakemap=Array(BM_Bumper1_Ring, BM_Cube, BM_Gate1, BM_LSling, BM_LSling1, BM_LSling2, BM_Layer_1, BM_LeftFlip, BM_LeftFlip1, BM_LeftFlip1U, BM_LeftFlipU, BM_Parts, BM_Parts_001, BM_Parts_002, BM_Plane, BM_Playfield, BM_Playfield_001, BM_Prim14, BM_Prim2, BM_Prim3, BM_Prim4, BM_Prim5, BM_Prim6, BM_RSling, BM_RSling1, BM_RSling2, BM_RightFlip, BM_RightFlip1, BM_RightFlip1U, BM_RightFlipU, BM_SaucerLWU, BM_SaucerRWU, BM_UnderPF, BM_UnderUPF, BM_UpperPF, BM_UpperSaucer1WU, BM_UpperSaucer2WU, BM_UpperSaucer3DWU, BM_inPFGlass, BM_sling1, BM_sling2, BM_sw12, BM_sw13, BM_sw21, BM_sw22, BM_sw23, BM_sw24, BM_sw26, BM_sw27, BM_sw28, BM_sw29, BM_sw30, BM_sw31, BM_sw33, BM_sw34, BM_sw35, BM_sw36, BM_sw37_001, BM_sw46, BM_sw47, BM_sw48)
Dim BG_Lightmap: BG_Lightmap=Array(LM_D_a10_Plane, LM_D_a10_inPFGlass, LM_D_a11_Plane, LM_D_a11_inPFGlass, LM_D_a12_Plane, LM_D_a12_inPFGlass, LM_D_a13_Plane, LM_D_a13_inPFGlass, LM_D_a14_Plane, LM_D_a14_inPFGlass, LM_D_a15_Plane, LM_D_a15_inPFGlass, LM_D_a16_Plane, LM_D_a16_inPFGlass, LM_D_a20_Plane, LM_D_a20_inPFGlass, LM_D_a21_Plane, LM_D_a21_inPFGlass, LM_D_a22_Plane, LM_D_a22_inPFGlass, LM_D_a23_Plane, LM_D_a23_inPFGlass, LM_D_a24_Plane, LM_D_a24_inPFGlass, LM_D_a25_Plane, LM_D_a25_inPFGlass, LM_D_a26_Plane, LM_D_a26_inPFGlass, LM_D_a30_Plane, LM_D_a30_inPFGlass, LM_D_a31_Plane, LM_D_a31_inPFGlass, LM_D_a32_Plane, LM_D_a32_inPFGlass, LM_D_a33_Plane, LM_D_a33_inPFGlass, LM_D_a34_Plane, LM_D_a34_inPFGlass, LM_D_a35_Plane, LM_D_a35_inPFGlass, LM_D_a36_Plane, LM_D_a36_inPFGlass, LM_D_a40_Plane, LM_D_a40_inPFGlass, LM_D_a41_Plane, LM_D_a41_inPFGlass, LM_D_a42_Plane, LM_D_a42_inPFGlass, LM_D_a43_Plane, LM_D_a43_inPFGlass, LM_D_a44_Plane, LM_D_a44_inPFGlass, LM_D_a45_Plane, LM_D_a45_inPFGlass, LM_D_a46_Plane, _
  LM_D_a46_inPFGlass, LM_D_a50_Plane, LM_D_a50_inPFGlass, LM_D_a51_Plane, LM_D_a51_inPFGlass, LM_D_a52_Plane, LM_D_a52_inPFGlass, LM_D_a53_Plane, LM_D_a53_inPFGlass, LM_D_a54_Plane, LM_D_a54_inPFGlass, LM_D_a55_Plane, LM_D_a55_inPFGlass, LM_D_a56_Plane, LM_D_a56_inPFGlass, LM_D_aa00_Plane, LM_D_aa00_inPFGlass, LM_D_aa01_Plane, LM_D_aa01_inPFGlass, LM_D_aa02_Plane, LM_D_aa02_inPFGlass, LM_D_aa03_Plane, LM_D_aa03_inPFGlass, LM_D_aa04_Plane, LM_D_aa04_inPFGlass, LM_D_aa05_Plane, LM_D_aa05_inPFGlass, LM_D_aa06_Plane, LM_D_aa06_inPFGlass, LM_GI_CreditLight_Layer_1, LM_GI_L12a_Bumper1_Ring, LM_GI_L12a_Layer_1, LM_GI_L12a_Parts, LM_GI_L12a_Playfield, LM_GI_L12a_RightFlip1, LM_GI_L12a_RightFlip1U, LM_GI_L12a_UpperPF, LM_GI_L12a_sw46, LM_GI_L12a_sw47, LM_GI_L12a_sw48, LM_GI_L33a1_LSling, LM_GI_L33a1_LSling1, LM_GI_L33a1_LSling2, LM_GI_L33a1_Layer_1, LM_GI_L33a1_LeftFlip, LM_GI_L33a1_LeftFlip1, LM_GI_L33a1_LeftFlip1U, LM_GI_L33a1_LeftFlipU, LM_GI_L33a1_Parts, LM_GI_L33a1_Playfield, LM_GI_L33a1_UnderPF, _
  LM_GI_L33a1_sling2, LM_GI_L33a1_sw35, LM_GI_L33a1_sw36, LM_GI_L33a2_Layer_1, LM_GI_L33a2_LeftFlipU, LM_GI_L33a2_Parts, LM_GI_L33a2_Playfield, LM_GI_L33a2_Prim2, LM_GI_L33a2_Prim4, LM_GI_L33a2_RSling, LM_GI_L33a2_RSling1, LM_GI_L33a2_RSling2, LM_GI_L33a2_RightFlip, LM_GI_L33a2_RightFlipU, LM_GI_L33a2_UnderPF, LM_GI_L33a2_sling1, LM_GI_L33a2_sw34, LM_GI_L33a3_Layer_1, LM_GI_L33a3_LeftFlipU, LM_GI_L33a3_Parts, LM_GI_L33a3_Playfield, LM_GI_L33a3_Prim5, LM_GI_L33a3_SaucerLWU, LM_GI_L33a4_LSling, LM_GI_L33a4_LSling1, LM_GI_L33a4_LSling2, LM_GI_L33a4_Layer_1, LM_GI_L33a4_LeftFlip, LM_GI_L33a4_LeftFlipU, LM_GI_L33a4_Parts, LM_GI_L33a4_Playfield, LM_GI_L33a4_Prim5, LM_GI_L33a5_Layer_1, LM_GI_L33a5_LeftFlip1U, LM_GI_L33a5_LeftFlipU, LM_GI_L33a5_Parts, LM_GI_L33a5_Playfield, LM_GI_L33a5_Prim2, LM_GI_L33a5_RSling, LM_GI_L33a5_RSling1, LM_GI_L33a5_RSling2, LM_GI_L33a5_RightFlip, LM_GI_L33a5_RightFlipU, LM_GI_L33a5_UnderPF, LM_GI_L33a5_sw26, LM_GI_L33a5_sw27, LM_GI_L33a5_sw33, LM_GI_L33a6_LSling, LM_GI_L33a6_LSling1, _
  LM_GI_L33a6_LSling2, LM_GI_L33a6_Layer_1, LM_GI_L33a6_LeftFlip, LM_GI_L33a6_LeftFlipU, LM_GI_L33a6_Parts, LM_GI_L33a6_Playfield, LM_GI_L33a6_UnderPF, LM_GI_L33a6_sw28, LM_GI_L33a6_sw35, LM_GI_L33a6_sw36, LM_GI_L33a7_Layer_1, LM_GI_L33a7_Parts, LM_GI_L33a7_Playfield, LM_GI_L49a1_Layer_1, LM_GI_L49a1_Parts, LM_GI_L49a1_RightFlip1, LM_GI_L49a1_RightFlip1U, LM_GI_L49a1_UpperPF, LM_GI_L49a1_sw46, LM_GI_L49a1_sw47, LM_GI_L49a1_sw48, LM_GI_L49a2_Bumper1_Ring, LM_GI_L49a2_Layer_1, LM_GI_L49a2_Parts, LM_GI_L49a2_Playfield, LM_GI_L49a2_RightFlip1, LM_GI_L49a2_RightFlip1U, LM_GI_L49a2_UnderUPF, LM_GI_L49a2_UpperPF, LM_GI_L49a2_sw37_001, LM_GI_L49a2_sw46, LM_GI_L49a2_sw47, LM_GI_L49a2_sw48, LM_GI_L49a3_Bumper1_Ring, LM_GI_L49a3_Layer_1, LM_GI_L49a3_Parts, LM_GI_L49a3_Playfield, LM_GI_L49a3_RightFlip1, LM_GI_L49a3_RightFlip1U, LM_GI_L49a3_UnderUPF, LM_GI_L49a3_UpperPF, LM_GI_L49a3_sw46, LM_GI_L49a3_sw47, LM_GI_L49a3_sw48, LM_GI_L49a4_Layer_1, LM_GI_L49a4_Parts, LM_GI_L49a4_Playfield, LM_GI_L49a4_UpperPF, _
  LM_GI_L49a4_UpperSaucer1WU, LM_GI_L49a5_Layer_1, LM_GI_L49a5_Parts, LM_GI_L49a5_Playfield, LM_GI_L49a5_UpperPF, LM_GI_L49a5_sw48, LM_GI_L49a6_Layer_1, LM_GI_L49a6_Parts, LM_GI_L49a6_UpperPF, LM_GI_L67b_Layer_1, LM_GI_L67b_Parts, LM_GI_L67b_Playfield, LM_GI_L67b_UpperPF, LM_GI_L67b_sw47, LM_GI_L67b_sw48, LM_GI_l113b_Bumper1_Ring, LM_GI_l113b_Layer_1, LM_GI_l113b_Parts, LM_GI_l113b_Playfield, LM_GI_l113b_Prim6, LM_GI_l113b_RightFlip1, LM_GI_l113b_RightFlip1U, LM_GI_l113b_UnderUPF, LM_GI_l113b_UpperPF, LM_GI_l113b_sw46, LM_GI_l113b_sw47, LM_GI_l113b_sw48, LM_GI_l114_Bumper1_Ring, LM_GI_l114_Layer_1, LM_GI_l114_Parts, LM_GI_l114_Playfield, LM_GI_l114_Prim6, LM_GI_l114_RightFlip1U, LM_GI_l114_UnderUPF, LM_GI_l114_UpperPF, LM_GI_l114_sw13, LM_GI_l114_sw46, LM_GI_l65b_Layer_1, LM_GI_l65b_Parts, LM_GI_l65b_Playfield, LM_GI_l65b_UnderUPF, LM_GI_l65b_UpperPF, LM_GI_l65b_sw46, LM_GI_l65b_sw47, LM_GI_l65b_sw48, LM_GI_l81b_Layer_1, LM_GI_l81b_Parts, LM_GI_l81b_Playfield, LM_GI_l81b_UnderUPF, LM_GI_l81b_UpperPF, _
  LM_GI_l81b_sw46, LM_GI_l81b_sw47, LM_GI_l81b_sw48, LM_GI_l83_Gate1, LM_GI_l83_LeftFlipU, LM_GI_l83_Parts, LM_GI_l83_Playfield, LM_GI_l83_RightFlip, LM_GI_l83_RightFlipU, LM_GI_l97b_Bumper1_Ring, LM_GI_l97b_Layer_1, LM_GI_l97b_Parts, LM_GI_l97b_RightFlip1, LM_GI_l97b_RightFlip1U, LM_GI_l97b_UnderUPF, LM_GI_l97b_UpperPF, LM_GI_l97b_sw46, LM_GI_l97b_sw47, LM_GI_l97b_sw48, LM_GI_l98_Layer_1, LM_GI_l98_Parts, LM_GI_l98_Playfield, LM_GI_l98_Prim6, LM_GI_l98_RightFlip1, LM_GI_l98_RightFlip1U, LM_GI_l98_UnderUPF, LM_GI_l98_UpperPF, LM_L_l10_Layer_1, LM_L_l10_Parts, LM_L_l10_UnderPF, LM_L_l113_Layer_1, LM_L_l113_Parts, LM_L_l113_UnderUPF, LM_L_l114_Layer_1, LM_L_l114_Parts, LM_L_l114_UnderUPF, LM_L_l115_Layer_1, LM_L_l115_Parts, LM_L_l115_UnderPF, LM_L_l14_UnderPF, LM_L_l15_UnderUPF, LM_L_l18_UnderPF, LM_L_l19_UnderPF, LM_L_l2_UnderPF, LM_L_l20_UnderPF, LM_L_l21_Parts, LM_L_l21_UnderUPF, LM_L_l22_UnderPF, LM_L_l23_Layer_1, LM_L_l23_Parts, LM_L_l23_UnderUPF, LM_L_l24_Parts, LM_L_l24_RSling1, LM_L_l24_UnderPF, _
  LM_L_l25_UnderPF, LM_L_l26_UnderPF, LM_L_l28_Parts, LM_L_l28_UnderPF, LM_L_l28_UnderUPF, LM_L_l28_sw26, LM_L_l3_LeftFlipU, LM_L_l3_UnderPF, LM_L_l30_LSling, LM_L_l30_LSling1, LM_L_l30_LSling2, LM_L_l30_Layer_1, LM_L_l30_Parts, LM_L_l30_UnderPF, LM_L_l31_RightFlip1U, LM_L_l31_UnderUPF, LM_L_l34_UnderPF, LM_L_l35_UnderPF, LM_L_l36_UnderPF, LM_L_l37_UnderUPF, LM_L_l38_UnderPF, LM_L_l39_Layer_1, LM_L_l39_Parts, LM_L_l39_UnderUPF, LM_L_l4_RightFlipU, LM_L_l4_UnderPF, LM_L_l40_LSling, LM_L_l40_LSling1, LM_L_l40_LSling2, LM_L_l40_Layer_1, LM_L_l40_Parts, LM_L_l40_UnderPF, LM_L_l41_Parts, LM_L_l41_UnderPF, LM_L_l42_UnderPF, LM_L_l43_Layer_1, LM_L_l43_Parts, LM_L_l43_UnderPF, LM_L_l44_Parts, LM_L_l44_UnderPF, LM_L_l44_sw27, LM_L_l46_UnderPF, LM_L_l47_Layer_1, LM_L_l47_Parts, LM_L_l47_UnderUPF, LM_L_l49rot_Parts, LM_L_l49rot_UnderPF, LM_L_l49rot_sw37_001, LM_L_l50_UnderPF, LM_L_l51_UnderPF, LM_L_l52_UnderPF, LM_L_l53_Parts, LM_L_l53_UnderUPF, LM_L_l54_UnderPF, LM_L_l56_Parts, LM_L_l56_UnderPF, LM_L_l57_Parts, _
  LM_L_l57_Prim2, LM_L_l57_UnderPF, LM_L_l58_Parts, LM_L_l58_UnderPF, LM_L_l59_Parts, LM_L_l6_UnderPF, LM_L_l60_LeftFlip1U, LM_L_l60_UnderPF, LM_L_l60_sw28, LM_L_l62_UnderPF, LM_L_l65_Layer_1, LM_L_l65_Parts, LM_L_l65_UnderUPF, LM_L_l66_Parts, LM_L_l66_UnderUPF, LM_L_l67_Parts, LM_L_l67_Prim14, LM_L_l67_UnderPF, LM_L_l67_sw30, LM_L_l7_UnderPF, LM_L_l8_Layer_1, LM_L_l8_Parts, LM_L_l8_UnderPF, LM_L_l81_Layer_1, LM_L_l81_Parts, LM_L_l81_UnderUPF, LM_L_l82_Parts, LM_L_l82_UnderUPF, LM_L_l9_Parts, LM_L_l9_Prim2, LM_L_l9_UnderPF, LM_L_l97_Layer_1, LM_L_l97_Parts, LM_L_l97_UnderUPF, LM_L_l98_Layer_1, LM_L_l98_Parts, LM_L_l98_Playfield, LM_L_l98_UnderUPF, LM_L_l99_UnderPF, LM_UL_l10_Layer_1, LM_UL_l10_Parts, LM_UL_l10_Playfield, LM_UL_l10_Prim3, LM_UL_l10_UnderPF, LM_UL_l10_sw21, LM_UL_l113_Parts, LM_UL_l113_UnderUPF, LM_UL_l114_Layer_1, LM_UL_l114_Parts, LM_UL_l114_Playfield, LM_UL_l114_UnderUPF, LM_UL_l114_UpperPF, LM_UL_l115_Layer_1, LM_UL_l115_LeftFlip1, LM_UL_l115_Parts, LM_UL_l115_Playfield, LM_UL_l115_UnderPF, _
  LM_UL_l14_Playfield, LM_UL_l14_UnderPF, LM_UL_l15_RightFlip1U, LM_UL_l15_UpperPF, LM_UL_l18_Parts, LM_UL_l18_Playfield, LM_UL_l19_LeftFlipU, LM_UL_l19_Playfield, LM_UL_l2_LSling, LM_UL_l2_LSling1, LM_UL_l2_LSling2, LM_UL_l2_Parts, LM_UL_l2_Playfield, LM_UL_l2_Prim5, LM_UL_l20_Playfield, LM_UL_l20_RightFlipU, LM_UL_l21_UnderUPF, LM_UL_l21_UpperPF, LM_UL_l22_Playfield, LM_UL_l22_UnderPF, LM_UL_l22_sw21, LM_UL_l22_sw22, LM_UL_l22_sw23, LM_UL_l23_Layer_1, LM_UL_l23_Parts, LM_UL_l23_UnderUPF, LM_UL_l23_UpperPF, LM_UL_l24_Parts, LM_UL_l24_Playfield, LM_UL_l25_Parts, LM_UL_l25_Playfield, LM_UL_l25_Prim2, LM_UL_l25_UnderPF, LM_UL_l26_Parts, LM_UL_l26_Playfield, LM_UL_l26_UnderPF, LM_UL_l28_Playfield, LM_UL_l28_UnderPF, LM_UL_l28_UnderUPF, LM_UL_l28_UpperPF, LM_UL_l28_sw26, LM_UL_l3_LeftFlipU, LM_UL_l3_Playfield, LM_UL_l30_LSling, LM_UL_l30_LSling1, LM_UL_l30_LSling2, LM_UL_l30_Parts, LM_UL_l30_Playfield, LM_UL_l30_UnderPF, LM_UL_l30_sling2, LM_UL_l31_RightFlip1, LM_UL_l31_RightFlip1U, LM_UL_l31_UnderUPF, _
  LM_UL_l31_UpperPF, LM_UL_l34_Playfield, LM_UL_l35_Playfield, LM_UL_l36_Parts, LM_UL_l36_Playfield, LM_UL_l37_UnderUPF, LM_UL_l38_Playfield, LM_UL_l38_UnderPF, LM_UL_l38_sw22, LM_UL_l38_sw23, LM_UL_l38_sw24, LM_UL_l39_Layer_1, LM_UL_l39_Parts, LM_UL_l39_UnderUPF, LM_UL_l39_UpperPF, LM_UL_l4_Playfield, LM_UL_l4_RightFlipU, LM_UL_l40_LSling, LM_UL_l40_LSling1, LM_UL_l40_LSling2, LM_UL_l40_Parts, LM_UL_l40_Playfield, LM_UL_l40_UnderPF, LM_UL_l41_Parts, LM_UL_l41_Playfield, LM_UL_l41_Prim2, LM_UL_l41_UnderPF, LM_UL_l42_UnderPF, LM_UL_l43_UnderPF, LM_UL_l44_UnderPF, LM_UL_l44_sw27, LM_UL_l46_Playfield, LM_UL_l46_UnderPF, LM_UL_l47_Layer_1, LM_UL_l47_Parts, LM_UL_l47_RightFlip1U, LM_UL_l47_UnderUPF, LM_UL_l47_UpperPF, LM_UL_l50_Playfield, LM_UL_l51_Playfield, LM_UL_l52_Playfield, LM_UL_l52_UnderPF, LM_UL_l53_UnderUPF, LM_UL_l53_UpperPF, LM_UL_l54_Playfield, LM_UL_l54_UnderPF, LM_UL_l54_sw23, LM_UL_l54_sw24, LM_UL_l55_Layer_1, LM_UL_l55_Parts, LM_UL_l55_Prim6, LM_UL_l55_UnderUPF, LM_UL_l55_UpperPF, LM_UL_l56_Layer_1, _
  LM_UL_l56_Parts, LM_UL_l56_Playfield, LM_UL_l57_Layer_1, LM_UL_l57_Parts, LM_UL_l57_Playfield, LM_UL_l57_Prim2, LM_UL_l57_Prim3, LM_UL_l57_UnderPF, LM_UL_l58_Parts, LM_UL_l58_Playfield, LM_UL_l58_UnderPF, LM_UL_l58_sw26, LM_UL_l59_Playfield, LM_UL_l6_Parts, LM_UL_l6_Playfield, LM_UL_l6_UnderPF, LM_UL_l6_sw21, LM_UL_l6_sw22, LM_UL_l60_LeftFlip1U, LM_UL_l60_Playfield, LM_UL_l60_UnderPF, LM_UL_l60_sw28, LM_UL_l62_Parts, LM_UL_l62_Playfield, LM_UL_l62_UnderPF, LM_UL_l65_Layer_1, LM_UL_l65_Parts, LM_UL_l65_UnderUPF, LM_UL_l66_Parts, LM_UL_l66_UnderUPF, LM_UL_l67_UnderPF, LM_UL_l67_sw30, LM_UL_l7_Playfield, LM_UL_l7_UnderPF, LM_UL_l8_Parts, LM_UL_l8_UnderPF, LM_UL_l81_Layer_1, LM_UL_l81_Parts, LM_UL_l81_UnderUPF, LM_UL_l82_Layer_1, LM_UL_l82_Parts, LM_UL_l82_UnderUPF, LM_UL_l9_Layer_1, LM_UL_l9_Parts, LM_UL_l9_Playfield, LM_UL_l9_Prim2, LM_UL_l9_UnderPF, LM_UL_l97_Parts, LM_UL_l97_UnderUPF, LM_UL_l98_Layer_1, LM_UL_l98_Parts, LM_UL_l98_UnderUPF, LM_UL_l99_LeftFlip1, LM_UL_l99_Parts, LM_UL_l99_Playfield, _
  LM_UL_l99_UnderPF, LM_UL_l99_sw28)
Dim BG_All: BG_All=Array(BM_Bumper1_Ring, BM_Cube, BM_Gate1, BM_LSling, BM_LSling1, BM_LSling2, BM_Layer_1, BM_LeftFlip, BM_LeftFlip1, BM_LeftFlip1U, BM_LeftFlipU, BM_Parts, BM_Parts_001, BM_Parts_002, BM_Plane, BM_Playfield, BM_Playfield_001, BM_Prim14, BM_Prim2, BM_Prim3, BM_Prim4, BM_Prim5, BM_Prim6, BM_RSling, BM_RSling1, BM_RSling2, BM_RightFlip, BM_RightFlip1, BM_RightFlip1U, BM_RightFlipU, BM_SaucerLWU, BM_SaucerRWU, BM_UnderPF, BM_UnderUPF, BM_UpperPF, BM_UpperSaucer1WU, BM_UpperSaucer2WU, BM_UpperSaucer3DWU, BM_inPFGlass, BM_sling1, BM_sling2, BM_sw12, BM_sw13, BM_sw21, BM_sw22, BM_sw23, BM_sw24, BM_sw26, BM_sw27, BM_sw28, BM_sw29, BM_sw30, BM_sw31, BM_sw33, BM_sw34, BM_sw35, BM_sw36, BM_sw37_001, BM_sw46, BM_sw47, BM_sw48, LM_D_a10_Plane, LM_D_a10_inPFGlass, LM_D_a11_Plane, LM_D_a11_inPFGlass, LM_D_a12_Plane, LM_D_a12_inPFGlass, LM_D_a13_Plane, LM_D_a13_inPFGlass, LM_D_a14_Plane, LM_D_a14_inPFGlass, LM_D_a15_Plane, LM_D_a15_inPFGlass, LM_D_a16_Plane, LM_D_a16_inPFGlass, LM_D_a20_Plane, _
  LM_D_a20_inPFGlass, LM_D_a21_Plane, LM_D_a21_inPFGlass, LM_D_a22_Plane, LM_D_a22_inPFGlass, LM_D_a23_Plane, LM_D_a23_inPFGlass, LM_D_a24_Plane, LM_D_a24_inPFGlass, LM_D_a25_Plane, LM_D_a25_inPFGlass, LM_D_a26_Plane, LM_D_a26_inPFGlass, LM_D_a30_Plane, LM_D_a30_inPFGlass, LM_D_a31_Plane, LM_D_a31_inPFGlass, LM_D_a32_Plane, LM_D_a32_inPFGlass, LM_D_a33_Plane, LM_D_a33_inPFGlass, LM_D_a34_Plane, LM_D_a34_inPFGlass, LM_D_a35_Plane, LM_D_a35_inPFGlass, LM_D_a36_Plane, LM_D_a36_inPFGlass, LM_D_a40_Plane, LM_D_a40_inPFGlass, LM_D_a41_Plane, LM_D_a41_inPFGlass, LM_D_a42_Plane, LM_D_a42_inPFGlass, LM_D_a43_Plane, LM_D_a43_inPFGlass, LM_D_a44_Plane, LM_D_a44_inPFGlass, LM_D_a45_Plane, LM_D_a45_inPFGlass, LM_D_a46_Plane, LM_D_a46_inPFGlass, LM_D_a50_Plane, LM_D_a50_inPFGlass, LM_D_a51_Plane, LM_D_a51_inPFGlass, LM_D_a52_Plane, LM_D_a52_inPFGlass, LM_D_a53_Plane, LM_D_a53_inPFGlass, LM_D_a54_Plane, LM_D_a54_inPFGlass, LM_D_a55_Plane, LM_D_a55_inPFGlass, LM_D_a56_Plane, LM_D_a56_inPFGlass, LM_D_aa00_Plane, _
  LM_D_aa00_inPFGlass, LM_D_aa01_Plane, LM_D_aa01_inPFGlass, LM_D_aa02_Plane, LM_D_aa02_inPFGlass, LM_D_aa03_Plane, LM_D_aa03_inPFGlass, LM_D_aa04_Plane, LM_D_aa04_inPFGlass, LM_D_aa05_Plane, LM_D_aa05_inPFGlass, LM_D_aa06_Plane, LM_D_aa06_inPFGlass, LM_GI_CreditLight_Layer_1, LM_GI_L12a_Bumper1_Ring, LM_GI_L12a_Layer_1, LM_GI_L12a_Parts, LM_GI_L12a_Playfield, LM_GI_L12a_RightFlip1, LM_GI_L12a_RightFlip1U, LM_GI_L12a_UpperPF, LM_GI_L12a_sw46, LM_GI_L12a_sw47, LM_GI_L12a_sw48, LM_GI_L33a1_LSling, LM_GI_L33a1_LSling1, LM_GI_L33a1_LSling2, LM_GI_L33a1_Layer_1, LM_GI_L33a1_LeftFlip, LM_GI_L33a1_LeftFlip1, LM_GI_L33a1_LeftFlip1U, LM_GI_L33a1_LeftFlipU, LM_GI_L33a1_Parts, LM_GI_L33a1_Playfield, LM_GI_L33a1_UnderPF, LM_GI_L33a1_sling2, LM_GI_L33a1_sw35, LM_GI_L33a1_sw36, LM_GI_L33a2_Layer_1, LM_GI_L33a2_LeftFlipU, LM_GI_L33a2_Parts, LM_GI_L33a2_Playfield, LM_GI_L33a2_Prim2, LM_GI_L33a2_Prim4, LM_GI_L33a2_RSling, LM_GI_L33a2_RSling1, LM_GI_L33a2_RSling2, LM_GI_L33a2_RightFlip, LM_GI_L33a2_RightFlipU, _
  LM_GI_L33a2_UnderPF, LM_GI_L33a2_sling1, LM_GI_L33a2_sw34, LM_GI_L33a3_Layer_1, LM_GI_L33a3_LeftFlipU, LM_GI_L33a3_Parts, LM_GI_L33a3_Playfield, LM_GI_L33a3_Prim5, LM_GI_L33a3_SaucerLWU, LM_GI_L33a4_LSling, LM_GI_L33a4_LSling1, LM_GI_L33a4_LSling2, LM_GI_L33a4_Layer_1, LM_GI_L33a4_LeftFlip, LM_GI_L33a4_LeftFlipU, LM_GI_L33a4_Parts, LM_GI_L33a4_Playfield, LM_GI_L33a4_Prim5, LM_GI_L33a5_Layer_1, LM_GI_L33a5_LeftFlip1U, LM_GI_L33a5_LeftFlipU, LM_GI_L33a5_Parts, LM_GI_L33a5_Playfield, LM_GI_L33a5_Prim2, LM_GI_L33a5_RSling, LM_GI_L33a5_RSling1, LM_GI_L33a5_RSling2, LM_GI_L33a5_RightFlip, LM_GI_L33a5_RightFlipU, LM_GI_L33a5_UnderPF, LM_GI_L33a5_sw26, LM_GI_L33a5_sw27, LM_GI_L33a5_sw33, LM_GI_L33a6_LSling, LM_GI_L33a6_LSling1, LM_GI_L33a6_LSling2, LM_GI_L33a6_Layer_1, LM_GI_L33a6_LeftFlip, LM_GI_L33a6_LeftFlipU, LM_GI_L33a6_Parts, LM_GI_L33a6_Playfield, LM_GI_L33a6_UnderPF, LM_GI_L33a6_sw28, LM_GI_L33a6_sw35, LM_GI_L33a6_sw36, LM_GI_L33a7_Layer_1, LM_GI_L33a7_Parts, LM_GI_L33a7_Playfield, LM_GI_L49a1_Layer_1, _
  LM_GI_L49a1_Parts, LM_GI_L49a1_RightFlip1, LM_GI_L49a1_RightFlip1U, LM_GI_L49a1_UpperPF, LM_GI_L49a1_sw46, LM_GI_L49a1_sw47, LM_GI_L49a1_sw48, LM_GI_L49a2_Bumper1_Ring, LM_GI_L49a2_Layer_1, LM_GI_L49a2_Parts, LM_GI_L49a2_Playfield, LM_GI_L49a2_RightFlip1, LM_GI_L49a2_RightFlip1U, LM_GI_L49a2_UnderUPF, LM_GI_L49a2_UpperPF, LM_GI_L49a2_sw37_001, LM_GI_L49a2_sw46, LM_GI_L49a2_sw47, LM_GI_L49a2_sw48, LM_GI_L49a3_Bumper1_Ring, LM_GI_L49a3_Layer_1, LM_GI_L49a3_Parts, LM_GI_L49a3_Playfield, LM_GI_L49a3_RightFlip1, LM_GI_L49a3_RightFlip1U, LM_GI_L49a3_UnderUPF, LM_GI_L49a3_UpperPF, LM_GI_L49a3_sw46, LM_GI_L49a3_sw47, LM_GI_L49a3_sw48, LM_GI_L49a4_Layer_1, LM_GI_L49a4_Parts, LM_GI_L49a4_Playfield, LM_GI_L49a4_UpperPF, LM_GI_L49a4_UpperSaucer1WU, LM_GI_L49a5_Layer_1, LM_GI_L49a5_Parts, LM_GI_L49a5_Playfield, LM_GI_L49a5_UpperPF, LM_GI_L49a5_sw48, LM_GI_L49a6_Layer_1, LM_GI_L49a6_Parts, LM_GI_L49a6_UpperPF, LM_GI_L67b_Layer_1, LM_GI_L67b_Parts, LM_GI_L67b_Playfield, LM_GI_L67b_UpperPF, LM_GI_L67b_sw47, LM_GI_L67b_sw48, _
  LM_GI_l113b_Bumper1_Ring, LM_GI_l113b_Layer_1, LM_GI_l113b_Parts, LM_GI_l113b_Playfield, LM_GI_l113b_Prim6, LM_GI_l113b_RightFlip1, LM_GI_l113b_RightFlip1U, LM_GI_l113b_UnderUPF, LM_GI_l113b_UpperPF, LM_GI_l113b_sw46, LM_GI_l113b_sw47, LM_GI_l113b_sw48, LM_GI_l114_Bumper1_Ring, LM_GI_l114_Layer_1, LM_GI_l114_Parts, LM_GI_l114_Playfield, LM_GI_l114_Prim6, LM_GI_l114_RightFlip1U, LM_GI_l114_UnderUPF, LM_GI_l114_UpperPF, LM_GI_l114_sw13, LM_GI_l114_sw46, LM_GI_l65b_Layer_1, LM_GI_l65b_Parts, LM_GI_l65b_Playfield, LM_GI_l65b_UnderUPF, LM_GI_l65b_UpperPF, LM_GI_l65b_sw46, LM_GI_l65b_sw47, LM_GI_l65b_sw48, LM_GI_l81b_Layer_1, LM_GI_l81b_Parts, LM_GI_l81b_Playfield, LM_GI_l81b_UnderUPF, LM_GI_l81b_UpperPF, LM_GI_l81b_sw46, LM_GI_l81b_sw47, LM_GI_l81b_sw48, LM_GI_l83_Gate1, LM_GI_l83_LeftFlipU, LM_GI_l83_Parts, LM_GI_l83_Playfield, LM_GI_l83_RightFlip, LM_GI_l83_RightFlipU, LM_GI_l97b_Bumper1_Ring, LM_GI_l97b_Layer_1, LM_GI_l97b_Parts, LM_GI_l97b_RightFlip1, LM_GI_l97b_RightFlip1U, LM_GI_l97b_UnderUPF, _
  LM_GI_l97b_UpperPF, LM_GI_l97b_sw46, LM_GI_l97b_sw47, LM_GI_l97b_sw48, LM_GI_l98_Layer_1, LM_GI_l98_Parts, LM_GI_l98_Playfield, LM_GI_l98_Prim6, LM_GI_l98_RightFlip1, LM_GI_l98_RightFlip1U, LM_GI_l98_UnderUPF, LM_GI_l98_UpperPF, LM_L_l10_Layer_1, LM_L_l10_Parts, LM_L_l10_UnderPF, LM_L_l113_Layer_1, LM_L_l113_Parts, LM_L_l113_UnderUPF, LM_L_l114_Layer_1, LM_L_l114_Parts, LM_L_l114_UnderUPF, LM_L_l115_Layer_1, LM_L_l115_Parts, LM_L_l115_UnderPF, LM_L_l14_UnderPF, LM_L_l15_UnderUPF, LM_L_l18_UnderPF, LM_L_l19_UnderPF, LM_L_l2_UnderPF, LM_L_l20_UnderPF, LM_L_l21_Parts, LM_L_l21_UnderUPF, LM_L_l22_UnderPF, LM_L_l23_Layer_1, LM_L_l23_Parts, LM_L_l23_UnderUPF, LM_L_l24_Parts, LM_L_l24_RSling1, LM_L_l24_UnderPF, LM_L_l25_UnderPF, LM_L_l26_UnderPF, LM_L_l28_Parts, LM_L_l28_UnderPF, LM_L_l28_UnderUPF, LM_L_l28_sw26, LM_L_l3_LeftFlipU, LM_L_l3_UnderPF, LM_L_l30_LSling, LM_L_l30_LSling1, LM_L_l30_LSling2, LM_L_l30_Layer_1, LM_L_l30_Parts, LM_L_l30_UnderPF, LM_L_l31_RightFlip1U, LM_L_l31_UnderUPF, LM_L_l34_UnderPF, _
  LM_L_l35_UnderPF, LM_L_l36_UnderPF, LM_L_l37_UnderUPF, LM_L_l38_UnderPF, LM_L_l39_Layer_1, LM_L_l39_Parts, LM_L_l39_UnderUPF, LM_L_l4_RightFlipU, LM_L_l4_UnderPF, LM_L_l40_LSling, LM_L_l40_LSling1, LM_L_l40_LSling2, LM_L_l40_Layer_1, LM_L_l40_Parts, LM_L_l40_UnderPF, LM_L_l41_Parts, LM_L_l41_UnderPF, LM_L_l42_UnderPF, LM_L_l43_Layer_1, LM_L_l43_Parts, LM_L_l43_UnderPF, LM_L_l44_Parts, LM_L_l44_UnderPF, LM_L_l44_sw27, LM_L_l46_UnderPF, LM_L_l47_Layer_1, LM_L_l47_Parts, LM_L_l47_UnderUPF, LM_L_l49rot_Parts, LM_L_l49rot_UnderPF, LM_L_l49rot_sw37_001, LM_L_l50_UnderPF, LM_L_l51_UnderPF, LM_L_l52_UnderPF, LM_L_l53_Parts, LM_L_l53_UnderUPF, LM_L_l54_UnderPF, LM_L_l56_Parts, LM_L_l56_UnderPF, LM_L_l57_Parts, LM_L_l57_Prim2, LM_L_l57_UnderPF, LM_L_l58_Parts, LM_L_l58_UnderPF, LM_L_l59_Parts, LM_L_l6_UnderPF, LM_L_l60_LeftFlip1U, LM_L_l60_UnderPF, LM_L_l60_sw28, LM_L_l62_UnderPF, LM_L_l65_Layer_1, LM_L_l65_Parts, LM_L_l65_UnderUPF, LM_L_l66_Parts, LM_L_l66_UnderUPF, LM_L_l67_Parts, LM_L_l67_Prim14, LM_L_l67_UnderPF, _
  LM_L_l67_sw30, LM_L_l7_UnderPF, LM_L_l8_Layer_1, LM_L_l8_Parts, LM_L_l8_UnderPF, LM_L_l81_Layer_1, LM_L_l81_Parts, LM_L_l81_UnderUPF, LM_L_l82_Parts, LM_L_l82_UnderUPF, LM_L_l9_Parts, LM_L_l9_Prim2, LM_L_l9_UnderPF, LM_L_l97_Layer_1, LM_L_l97_Parts, LM_L_l97_UnderUPF, LM_L_l98_Layer_1, LM_L_l98_Parts, LM_L_l98_Playfield, LM_L_l98_UnderUPF, LM_L_l99_UnderPF, LM_UL_l10_Layer_1, LM_UL_l10_Parts, LM_UL_l10_Playfield, LM_UL_l10_Prim3, LM_UL_l10_UnderPF, LM_UL_l10_sw21, LM_UL_l113_Parts, LM_UL_l113_UnderUPF, LM_UL_l114_Layer_1, LM_UL_l114_Parts, LM_UL_l114_Playfield, LM_UL_l114_UnderUPF, LM_UL_l114_UpperPF, LM_UL_l115_Layer_1, LM_UL_l115_LeftFlip1, LM_UL_l115_Parts, LM_UL_l115_Playfield, LM_UL_l115_UnderPF, LM_UL_l14_Playfield, LM_UL_l14_UnderPF, LM_UL_l15_RightFlip1U, LM_UL_l15_UpperPF, LM_UL_l18_Parts, LM_UL_l18_Playfield, LM_UL_l19_LeftFlipU, LM_UL_l19_Playfield, LM_UL_l2_LSling, LM_UL_l2_LSling1, LM_UL_l2_LSling2, LM_UL_l2_Parts, LM_UL_l2_Playfield, LM_UL_l2_Prim5, LM_UL_l20_Playfield, LM_UL_l20_RightFlipU, _
  LM_UL_l21_UnderUPF, LM_UL_l21_UpperPF, LM_UL_l22_Playfield, LM_UL_l22_UnderPF, LM_UL_l22_sw21, LM_UL_l22_sw22, LM_UL_l22_sw23, LM_UL_l23_Layer_1, LM_UL_l23_Parts, LM_UL_l23_UnderUPF, LM_UL_l23_UpperPF, LM_UL_l24_Parts, LM_UL_l24_Playfield, LM_UL_l25_Parts, LM_UL_l25_Playfield, LM_UL_l25_Prim2, LM_UL_l25_UnderPF, LM_UL_l26_Parts, LM_UL_l26_Playfield, LM_UL_l26_UnderPF, LM_UL_l28_Playfield, LM_UL_l28_UnderPF, LM_UL_l28_UnderUPF, LM_UL_l28_UpperPF, LM_UL_l28_sw26, LM_UL_l3_LeftFlipU, LM_UL_l3_Playfield, LM_UL_l30_LSling, LM_UL_l30_LSling1, LM_UL_l30_LSling2, LM_UL_l30_Parts, LM_UL_l30_Playfield, LM_UL_l30_UnderPF, LM_UL_l30_sling2, LM_UL_l31_RightFlip1, LM_UL_l31_RightFlip1U, LM_UL_l31_UnderUPF, LM_UL_l31_UpperPF, LM_UL_l34_Playfield, LM_UL_l35_Playfield, LM_UL_l36_Parts, LM_UL_l36_Playfield, LM_UL_l37_UnderUPF, LM_UL_l38_Playfield, LM_UL_l38_UnderPF, LM_UL_l38_sw22, LM_UL_l38_sw23, LM_UL_l38_sw24, LM_UL_l39_Layer_1, LM_UL_l39_Parts, LM_UL_l39_UnderUPF, LM_UL_l39_UpperPF, LM_UL_l4_Playfield, LM_UL_l4_RightFlipU, _
  LM_UL_l40_LSling, LM_UL_l40_LSling1, LM_UL_l40_LSling2, LM_UL_l40_Parts, LM_UL_l40_Playfield, LM_UL_l40_UnderPF, LM_UL_l41_Parts, LM_UL_l41_Playfield, LM_UL_l41_Prim2, LM_UL_l41_UnderPF, LM_UL_l42_UnderPF, LM_UL_l43_UnderPF, LM_UL_l44_UnderPF, LM_UL_l44_sw27, LM_UL_l46_Playfield, LM_UL_l46_UnderPF, LM_UL_l47_Layer_1, LM_UL_l47_Parts, LM_UL_l47_RightFlip1U, LM_UL_l47_UnderUPF, LM_UL_l47_UpperPF, LM_UL_l50_Playfield, LM_UL_l51_Playfield, LM_UL_l52_Playfield, LM_UL_l52_UnderPF, LM_UL_l53_UnderUPF, LM_UL_l53_UpperPF, LM_UL_l54_Playfield, LM_UL_l54_UnderPF, LM_UL_l54_sw23, LM_UL_l54_sw24, LM_UL_l55_Layer_1, LM_UL_l55_Parts, LM_UL_l55_Prim6, LM_UL_l55_UnderUPF, LM_UL_l55_UpperPF, LM_UL_l56_Layer_1, LM_UL_l56_Parts, LM_UL_l56_Playfield, LM_UL_l57_Layer_1, LM_UL_l57_Parts, LM_UL_l57_Playfield, LM_UL_l57_Prim2, LM_UL_l57_Prim3, LM_UL_l57_UnderPF, LM_UL_l58_Parts, LM_UL_l58_Playfield, LM_UL_l58_UnderPF, LM_UL_l58_sw26, LM_UL_l59_Playfield, LM_UL_l6_Parts, LM_UL_l6_Playfield, LM_UL_l6_UnderPF, LM_UL_l6_sw21, _
  LM_UL_l6_sw22, LM_UL_l60_LeftFlip1U, LM_UL_l60_Playfield, LM_UL_l60_UnderPF, LM_UL_l60_sw28, LM_UL_l62_Parts, LM_UL_l62_Playfield, LM_UL_l62_UnderPF, LM_UL_l65_Layer_1, LM_UL_l65_Parts, LM_UL_l65_UnderUPF, LM_UL_l66_Parts, LM_UL_l66_UnderUPF, LM_UL_l67_UnderPF, LM_UL_l67_sw30, LM_UL_l7_Playfield, LM_UL_l7_UnderPF, LM_UL_l8_Parts, LM_UL_l8_UnderPF, LM_UL_l81_Layer_1, LM_UL_l81_Parts, LM_UL_l81_UnderUPF, LM_UL_l82_Layer_1, LM_UL_l82_Parts, LM_UL_l82_UnderUPF, LM_UL_l9_Layer_1, LM_UL_l9_Parts, LM_UL_l9_Playfield, LM_UL_l9_Prim2, LM_UL_l9_UnderPF, LM_UL_l97_Parts, LM_UL_l97_UnderUPF, LM_UL_l98_Layer_1, LM_UL_l98_Parts, LM_UL_l98_UnderUPF, LM_UL_l99_LeftFlip1, LM_UL_l99_Parts, LM_UL_l99_Playfield, LM_UL_l99_UnderPF, LM_UL_l99_sw28)
' VLM  Arrays - End

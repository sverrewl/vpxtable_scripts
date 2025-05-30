'*********************************************************************************************************************************
' === TABLE OF CONTENTS  ===
'
' You can quickly jump to a section by searching the four letter tag (ZXXX)
'
' ZOPT: User Options
' ZCON: Constants and Global Variables
' ZTIM: Timers
' ZINI: Table Initialization and Exiting
'   ZMAT: General Math Functions
' ZBBR: Ball Brightness
' ZGIC: GI Colors
' ZANI: Misc Animations
' ZRBR: Room Brightness
' ZPHY: General Advice on Physics
' ZNFF: Flipper Corrections
' ZKEY: Key Press Handling
' ZDRN: Drain, Trough, and Ball Release
' ZFLP: Flippers
' ZSLG: Slingshots
' ZSWI: Switches
'   ZTRI: Triggers
' ZTAR: Targets
' ZBMP: Bumpers
' ZLNS: Lane Steering
' ZKIC: Kickers, Saucers
'   ZDMP: Rubber Dampeners
'   ZBOU: VPW TargetBouncer for targets and posts
' ZSSC: Slingshot Correction
'   ZRDT: Drop Targets
' ZRST: Stand-Up Targets
' ZSHA: Ambient Ball Shadows
'   ZBRL: Ball Rolling and Drop Sounds
'   ZRRL: Ramp Rolling SFX
'   ZFLE: Fleep Mechanical Sounds
' ZDTD: Desktop Digits
'   ZVRR: VR Room / VR Cabinet
' ZVLM: VLM Arrays

'*********************************************************************************************************************************

Option Explicit
Randomize

On Error Resume Next
ExecuteGlobal GetTextFile("controller.vbs")
If Err Then MsgBox "You need the controller.vbs in order to run this table, available in the vp10 package"
On Error Goto 0

'*******************************************
'  ZOPT: User Options
'*******************************************

Dim UsingRom : UsingRom = True
Dim LightLevel : LightLevel = 0.25        ' Level of room lighting (0 to 1), where 0 is dark and 100 is brightest
Dim ColorLUT : ColorLUT = 1           ' Color desaturation LUTs: 1 to 11, where 1 is normal and 11 is black'n'white
Dim VolumeDial : VolumeDial = 1             ' Overall Mechanical sound effect volume. Recommended values should be no greater than 1.
Dim BallRollVolume : BallRollVolume = 0.5     ' Level of ball rolling volume. Value between 0 and 1
Dim RampRollVolume : RampRollVolume = 0.5     ' Level of ramp rolling volume. Value between 0 and 1


'///////////////////////-----General Sound Options-----///////////////////////
'// VolumeDial:
'// VolumeDial is the actual global volume multiplier for the mechanical sounds.
'// Values smaller than 1 will decrease mechanical sounds volume.
'// Recommended values should be no greater than 1.

Dim DynamicLampIntensity, DynamicGIIntensity, DynamicFlasherIntensity, SolControlledLights


' DYNAMIC LAMP INTENSITY MULTIPLIER
' Numeric value to dynamically multiply every playfield lamp's intensity (default=1)
DynamicLampIntensity = 1

' DYNAMIC GI INTENSITY MULTIPLIER
' Numeric value to dynamically multiply every GI lamp's intensity (default=1, 0=Off)
DynamicGIIntensity = 1

' SOLENOID CONTROLLED Lights
' 0 = GI lights always on like orig table (default)
' 1 = let the table control the GI lights on/off
SolControlledLights = 0

' Called when options are tweaked by the player.
' - 0: game has started, good time to load options and adjust accordingly
' - 1: an option has changed
' - 2: options have been reseted
' - 3: player closed the tweak UI, good time to update staticly prerendered parts
' Table1.Option arguments are:
' - option name, minimum value, maximum value, step between valid values, default value, unit (0=None, 1=Percent), an optional arry of literal strings
Dim dspTriggered : dspTriggered = False
Sub Table1_OptionEvent(ByVal eventId)
  dim v, b, mb, nc, sb, sr
  If eventId = 1 And Not dspTriggered Then dspTriggered = True : DisableStaticPreRendering = True : End If

  'SideRails
  sr = Table1.Option("Metal Side Rails and Lockbar", 0, 1, 1, 0, 0, Array("False", "True"))
  SetSideRails sr

  'Middle Post
  mb = Table1.Option("Middle Post", 0, 1, 1, 0, 0, Array("False", "True"))
  SetMiddlePost mb

  'Sidewall
  sb = Table1.Option("Side Blades", 0, 1, 1, 0, 0, Array("False", "True"))
  SetSideWall sb

  'NippleCovers
  nc = Table1.Option("Hide Nipples", 0, 1, 1, 0, 0, Array("False", "True"))
  SetPGMode nc

  'GI Color
  v = Table1.Option("GI Lights", 0, 1, 1, 0, 0, Array("Classic", "White"))
  SetGIColor v

  b = Table1.Option("GI Brightness", 0, 1, .1, 1, 1)
  GIBrightness b

  ' Color Saturation
    ColorLUT = Table1.Option("Color Saturation", 1, 11, 1, 1, 0, _
    Array("Normal", "Desaturated 10%", "Desaturated 20%", "Desaturated 30%", "Desaturated 40%", "Desaturated 50%", _
        "Desaturated 60%", "Desaturated 70%", "Desaturated 80%", "Desaturated 90%", "Black 'n White"))
  if ColorLUT = 1 Then Table1.ColorGradeImage = ""
  if ColorLUT = 2 Then Table1.ColorGradeImage = "colorgradelut256x16-10"
  if ColorLUT = 3 Then Table1.ColorGradeImage = "colorgradelut256x16-20"
  if ColorLUT = 4 Then Table1.ColorGradeImage = "colorgradelut256x16-30"
  if ColorLUT = 5 Then Table1.ColorGradeImage = "colorgradelut256x16-40"
  if ColorLUT = 6 Then Table1.ColorGradeImage = "colorgradelut256x16-50"
  if ColorLUT = 7 Then Table1.ColorGradeImage = "colorgradelut256x16-60"
  if ColorLUT = 8 Then Table1.ColorGradeImage = "colorgradelut256x16-70"
  if ColorLUT = 9 Then Table1.ColorGradeImage = "colorgradelut256x16-80"
  if ColorLUT = 10 Then Table1.ColorGradeImage = "colorgradelut256x16-90"
  if ColorLUT = 11 Then Table1.ColorGradeImage = "colorgradelut256x16-100"

    ' Sound volumes
    VolumeDial = Table1.Option("Mech Volume", 0, 1, 0.01, 0.8, 1)
    BallRollVolume = Table1.Option("Ball Roll Volume", 0, 1, 0.01, 0.5, 1)
  RampRollVolume = Table1.Option("Ramp Roll Volume", 0, 1, 0.01, 0.5, 1)

  ' Room brightness
' LightLevel = Table1.Option("Table Brightness (Ambient Light Level)", 0, 1, 0.01, .5, 1)
  LightLevel = NightDay/100
  SetRoomBrightness LightLevel   'Uncomment this line for lightmapped tables.

  'Code to fix L37 insert that was baked white
  Dim cYellow: cYellow = rgb(255,255,50)
  Dim xx, BLI
  For each BLI in BL_L_l37: BLI.color = cYellow: Next


  If eventId = 3 And dspTriggered Then dspTriggered = False : DisableStaticPreRendering = False : End If
End Sub

Sub SetMiddlePost(mb)
  Dim BK
  if mb=1 Then
    For each BK in BP_MiddlePost
      BK.visible=1
    Next
      zCol_Rubber_Post003.collidable=1
  Else
    For each BK in BP_MiddlePost
      BK.visible=0
    Next
      zCol_Rubber_Post003.collidable=0
  End If
End Sub

Sub SetSideRails(sr)
  if sr=1 Then
    VRMetalRails.visible=True
  Else
    VRMetalRails.visible=False
  End If
End Sub


Sub SetSideWall(sb)
  if sb=1 Then
    Sidewall.visible=True
  Else
    Sidewall.visible=False
  End If
End Sub

Sub SetPGMode(nc)
  Dim BK
  if nc=1 Then
    For each BK in BP_NCoverL
      BK.visible=1
    Next
    For each BK in BP_NCoverR
      BK.visible=1
    Next
  Else
    For each BK in BP_NCoverL
      BK.visible=0
    Next
    For each BK in BP_NCoverR
      BK.visible=0
    Next
  End If
End Sub

'*******************************************
'  ZCON: Constants and Global Variables
'*******************************************

Dim tablewidth: tablewidth = Table1.width
Dim tableheight: tableheight = Table1.height

Const cGameName="trucksp3",SSolenoidOn="SolOn",SSolenoidOff="SolOff"
Const UseSolenoids  = 2
Const UseLamps    = 1
Const UseSync     = 1
Const HandleMech  = 0
Const UseGI     = 0

Const BallSize = 50   'Ball size must be 50
Const BallMass = 1    'Ball mass must be 1
Const tnob = 2      'Total number of balls on the playfield including captive balls.
Const lob = 0     'Total number of locked balls
Const VRTest = 0    'Test VR in DesktopMode
'VRRoom set based on RenderingMode
'Internal DMD in Desktop Mode, using a textbox (must be called before LoadVPM)
Dim VRMode, cab_mode, DesktopMode: DesktopMode = Table1.ShowDT
If RenderingMode = 2 or VRTest = 1 Then VRMode=1 Else VRMode=0
if Not DesktopMode and VRMode=0 Then cab_mode=1 Else cab_mode=0

Dim BIPL    'Ball in Plunger Lane
BIPL = 0
Dim BIP   'Ball in Play
BIP = 0

LoadVPM "03020000", "6803.VBS", 3.20

Dim bsTrough, bsSaucerTR, bsKickerTL, bsKickerTR

'------------------------Solenoid reference----------------------------

const sKickerTopRight         = 2  ' mapped in manual to 3
const sSaucerTopRight         = 4  ' mapped in manual to 2
const sDropTargetsReset       = 6  ' mapped in manual to 4
const sKickerTopLeft        = 8  ' mapped in manual to 1
const sEjectToShooter         = 12 ' mapped in manual to 9
const sOuthole            = 14 ' mapped in manual to 10
const sKnocker            = 15 ' mapped in manual to 11
const sFlippersEnable       = 19
const sLaneSteering         = 18 ' mapped in manual to 12

'*******************************************
' ZTIM: Timers
'*******************************************

'The FrameTimer interval should be -1, so executes at the display frame rate
'The frame timer should be used to update anything visual, like some animations, shadows, etc.
'However, a lot of animations will be handled in their respective _animate subroutines.

Dim FrameTime, InitFrameTime
InitFrameTime = 0

FrameTimer.Interval = -1
Sub FrameTimer_Timer() 'The frame timer interval should be -1, so executes at the display frame rate
  FrameTime = GameTime - InitFrameTime
  InitFrameTime = GameTime  'Count frametime
  'Add animation stuff here
  BSUpdate
  RollingUpdate       'update rolling sounds
  UpdateBallBrightness
  DoDTAnim        'handle drop target animations
  UpdateDropTargets
  DoSTAnim        'handle stand up target animations
  UpdateStandupTargets
  If VRMode = 1 Then
    VRDisplaytimer
  End If
  If VRMode = 0 Then
    Displaytimer
  End If

End Sub

'The CorTimer interval should be 10. It's sole purpose is to update the Cor calculations
CorTimer.Interval = 10
Sub CorTimer_Timer(): Cor.Update: End Sub

'*******************************************
' ZINI: Table Initialization and Exiting
'*******************************************

Dim gbot, TSBall1, TSBall2

gBot = GetBalls

Sub Table1_Init
    vpmInit me
    With Controller
       .GameName=cGameName
       If Err Then MsgBox "Can't start Game " & cGameName & vbNewLine & Err.Description : Exit Sub
       .SplashInfoLine="Truck Stop, Bally 1988"
       .HandleKeyboard=0
       .ShowTitle=0
       .ShowDMDOnly=1
       .ShowFrame=0
       .HandleMechanics=0
       .Hidden= not DesktopMode
       On Error Resume Next
       .Run
       If Err Then MsgBox Err.Description
       On Error Goto 0
    End With

  'Trough - Creates a ball in the kicker switch and gives that ball used an individual name.
  Set TSBall1 = BallRelease.CreateSizedballWithMass(Ballsize/2,Ballmass)
  Set TSBall2 = sw47.CreateSizedballWithMass(Ballsize/2,Ballmass)
  gBOT = Array(TSBall1,TSBall2)

  'Forces the trough switches to "on" at table boot so the game logic knows there are balls in the trough.
  Controller.Switch(47) = 1
  Controller.Switch(48) = 1
  Controller.Switch(8)=0

  ' Saucer Top Right
  Set bsSaucerTR = New cvpmSaucer
  With bsSaucerTR
    .InitKicker SaucerTopRight,swSaucerTopRight,197,25, 3.1415926/6
    .InitExitVariance 2, 1
  End With

  ' Kicker Top Right
  Set bsKickerTR=New cvpmSaucer
  With bsKickerTR
    .InitKicker KickerTopRight, swLowerDock,6,35, 0
  End With

  ' Kicker Top Left
  Set bsKickerTL=New cvpmSaucer
  With bsKickerTL
    .InitKicker KickerTopLeft, swKickerTopLeft,90,50, 80
  End With

  vpmNudge.TiltSwitch = swTiltMe
  vpmNudge.Sensitivity = 5
  vpmNudge.TiltObj = Array(LeftSlingshot, RightSlingshot, LeftSlingshotU, RightSlingshotU)


  'Reset Slings
  Slingshots_Init

  vpmMapLights AllLights

  dim xx
    For each xx in GILights
    xx.intensity = xx.intensity * DynamicGIIntensity * 0.5
    Next
  dim zz
  For each zz in AllLights
    zz.intensity = zz.intensity * DynamicLampIntensity
  Next

  If SolControlledLights = 0 Then
    GiON
  End If

  If VRMode = 1 Then
    setup_backglass()
  End If

  'Reset Steering wall in upper middle ramp system
  LaneSteeringInit

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

Const BallBrightness =  .7        'Ball brightness - Value between 0 and 1 (0=Dark ... 1=Bright)

' Constants for plunger lane ball darkening.
' You can make a temporary wall in the plunger lane area and use the co-ordinates from the corner control points.
Const PLOffset = 0.5      'Minimum ball brightness scale in plunger lane
Const PLLeft = 867.5617       'X position of punger lane left
Const PLRight = 927.295       'X position of punger lane right
Const PLTop = 979.7997        'Y position of punger lane top
Const PLBottom = 1756.848     'Y position of punger lane bottom
Dim PLGain: PLGain = (1-PLOffset)/(PLTop-PLBottom)
Dim gilvl

Sub UpdateBallBrightness
  Dim s, b_base, b_r, b_g, b_b, d_w
  b_base = 120 * BallBrightness + 70  ' orig was 120 and 70

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
cArray = Array(c2700k, cWhite)

Sub SetGIColor(c)
  Dim xx, BL
  GIColorPick = cArray(c)
  For each xx in GILights: xx.color = cArray(c): xx.colorfull = cArray(c): Next
  For each BL in BL_GI: BL.color = cArray(c): Next
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
  For each xx in GILights: xx.color = new_base: xx.colorfull = new_base: Next
  For each BL in BL_GI: BL.color = new_base: Next

End Sub

'******************************************************
'  ZANI: Misc Animations
'******************************************************

'--------flippers-----------
Sub LeftFlipper_Animate
    Dim a : a = LeftFlipper.CurrentAngle
    FlipperLSh.RotZ = a

    Dim v, BP
    v = 255.0 * (122.8 - LeftFlipper.CurrentAngle) / (122.8 -  72.8)

    For each BP in BP_Lflip
        BP.Rotz = a
        BP.visible = v < 128.0
    Next
    For each BP in BP_LflipU
        BP.Rotz = a
        BP.visible = v >= 128.0
    Next
End Sub

Sub LeftFlipper1_Animate
    Dim a : a = LeftFlipper1.CurrentAngle
    FlipperLSh1.RotZ = a

    Dim v, BP
    v = 255.0 * (153.87 - LeftFlipper1.CurrentAngle) / (153.87 -  98.87)

    For each BP in BP_Lflip1
        BP.Rotz = a
        BP.visible = v < 128.0
    Next
    For each BP in BP_Lflip1U
        BP.Rotz = a
        BP.visible = v >= 128.0
    Next
End Sub

Sub RightFlipper_Animate
    Dim a : a = RightFlipper.CurrentAngle
    FlipperRSh.RotZ = a

    Dim v, BP
    v = 255.0 * (-122.8 - RightFlipper.CurrentAngle) / (-122.8 +  72.8)

    For each BP in BP_Rflip
        BP.Rotz = a
        BP.visible = v < 128.0
    Next
    For each BP in BP_RflipU
        BP.Rotz = a
        BP.visible = v >= 128.0
    Next
End Sub

Sub RightFlipper1_Animate
    Dim a : a = RightFlipper1.CurrentAngle
    FlipperRSh1.RotZ = a

    Dim v, BP
    v = 255.0 * (-153.87 - RightFlipper1.CurrentAngle) / (-153.87 +  98.87)

    For each BP in BP_Rflip1
        BP.Rotz = a
        BP.visible = v < 128.0
    Next
    For each BP in BP_Rflip1U
        BP.Rotz = a
        BP.visible = v >= 128.0
    Next
End Sub


'--------drop targets-----------
Sub UpdateDropTargets
  dim BP, tz, rx, ry

  tz = BM_sw25.transz
  rx = BM_sw25.rotx
  ry = BM_sw25.roty
  For each BP in BP_sw25: BP.transz = tz: BP.rotx = rx: BP.roty = ry: Next

  tz = BM_sw26.transz
  rx = BM_sw26.rotx
  ry = BM_sw26.roty
  For each BP in BP_sw26: BP.transz = tz: BP.rotx = rx: BP.roty = ry: Next

  tz = BM_sw27.transz
  rx = BM_sw27.rotx
  ry = BM_sw27.roty
  For each BP in BP_sw27: BP.transz = tz: BP.rotx = rx: BP.roty = ry: Next

End Sub

'--------standup targets-----------
Sub UpdateStandupTargets
  dim BP, ty

    ty = BM_sw21.transy
  For each BP in BP_sw21 : BP.transy = ty: Next

    ty = BM_sw33.transy
  For each BP in BP_sw33: BP.transy = ty: Next

    ty = BM_sw34.transy
  For each BP in BP_sw34 : BP.transy = ty: Next

    ty = BM_sw35.transy
  For each BP in BP_sw35 : BP.transy = ty: Next

    ty = BM_sw36.transy
  For each BP in BP_sw36 : BP.transy = ty: Next

    ty = BM_sw37.transy
  For each BP in BP_sw37 : BP.transy = ty: Next

    ty = BM_sw38.transy
  For each BP in BP_sw38 : BP.transy = ty: Next

    ty = BM_sw41.transy
  For each BP in BP_sw41 : BP.transy = ty: Next

    ty = BM_sw42.transy
  For each BP in BP_sw42 : BP.transy = ty: Next

End Sub

'--------spinners-----------
Sub sw24_Animate
  Dim spinangle:spinangle = sw24.currentangle
  Dim BL : For Each BL in BP_sw24 : BL.RotX = spinangle: Next
  'Dim BN : For each BN in BP_pSpinnerRod_004 : BN.Rotx=spinangle/9: BN.transz=Abs(spinangle)/7: Next
End Sub

'--------gates-----------
Sub GateC1_Animate
  Dim spinangle:spinangle = GateC1.currentangle
  Dim BL : For Each BL in BP_GateRampCWire_001 : BL.RotX = spinangle: Next
End Sub

Sub GateC2_Animate
  Dim spinangle:spinangle = GateC2.currentangle
  Dim BL : For Each BL in BP_GateRampC2Wire_002 : BL.RotX = spinangle: Next
  Dim BN : For each BN in BP_pSpinnerRod_003 : BN.Rotx=spinangle/9: BN.transz=Abs(spinangle)/7: Next
End Sub

Sub GateI_Animate
  Dim spinangle:spinangle = GateI.currentangle
  Dim BL : For Each BL in BP_GateRampIWire_002 : BL.RotX = -spinangle: Next
  Dim BN : For each BN in BP_pSpinnerRod_001 : BN.Rotx=spinangle/-9.5: BN.transz=Abs(spinangle)/7: Next
End Sub

Sub GateT_Animate
  Dim spinangle:spinangle = GateT.currentangle
  Dim BL : For Each BL in BP_GateRampTWire_002 : BL.RotX = spinangle: Next
  Dim BN : For each BN in BP_pSpinnerRod : BN.Rotx=spinangle/9.5: BN.transz=Abs(spinangle)/7: Next
End Sub

Sub GateY1_Animate
  Dim spinangle:spinangle = GateY1.currentangle
  Dim BL : For Each BL in BP_GateRampYWire_001 : BL.RotX = spinangle: Next
End Sub

Sub GateY2_Animate
  Dim spinangle:spinangle = GateY2.currentangle
  Dim BL : For Each BL in BP_GateRampY2Wire_002 : BL.RotX = spinangle: Next
  Dim BN : For each BN in BP_pSpinnerRod_002 : BN.Rotx=spinangle/9: BN.transz=Abs(spinangle)/7: Next
End Sub

'--------rollovers-----------
Sub sw16_Animate
  Dim z : z = sw16.CurrentAnimOffset
  Dim BP : For Each BP in BP_sw16 : BP.transz = z: Next
End Sub

Sub sw31_Animate
  Dim z : z = sw31.CurrentAnimOffset
  Dim BP : For Each BP in BP_sw31_001 : BP.transz = z: Next
End Sub

Sub sw32_Animate
  Dim z : z = sw32.CurrentAnimOffset
  Dim BP : For Each BP in BP_sw32_001 : BP.transz = z: Next
End Sub

Sub sw39_Animate
  Dim z : z = sw39.CurrentAnimOffset
  Dim BP : For Each BP in BP_sw39_001 : BP.transz = z: Next
End Sub

Sub sw40_Animate
  Dim z : z = sw40.CurrentAnimOffset
  Dim BP : For Each BP in BP_sw40_001 : BP.transz = z: Next
End Sub

Sub WireRamp1Start_Animate
  Dim z : z = WireRamp1Start.CurrentAnimOffset
  Dim BP : For Each BP in BP_LaneSwitch1 : BP.transz = z: Next
End Sub

Sub WireRamp2Start_Animate
  Dim z : z = WireRamp2Start.CurrentAnimOffset
  Dim BP : For Each BP in BP_LaneSwitch2 : BP.transz = z: Next
End Sub

Sub WireRamp3Start_Animate
  Dim z : z = WireRamp3Start.CurrentAnimOffset
  Dim BP : For Each BP in BP_LaneSwitch3 : BP.transz = z: Next
End Sub

Sub WireRamp4Start_Animate
  Dim z : z = WireRamp4Start.CurrentAnimOffset
  Dim BP : For Each BP in BP_LaneSwitch4 : BP.transz = z: Next
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
'Dim LF1 : Set LF1 = New FlipperPolarity
'Dim RF1 : Set RF1 = New FlipperPolarity

InitPolarity


''*******************************************
'  Late 80's early 90's

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
  FlipperTricks LeftFlipper1, LFPress1, LFCount1, LFEndAngle1, LFState1
  FlipperTricks RightFlipper1, RFPress1, RFCount1, RFEndAngle1, RFState1
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

Dim LFPress1, RFPress1, LFCount1, RFCount1
Dim LFState1, RFState1
Dim RFEndAngle1, LFEndAngle1

Const FlipperCoilRampupMode = 0 '0 = fast, 1 = medium, 2 = slow (tap passes should work)

LFState = 1
RFState = 1
LFState1 = 1
RFState1 = 1
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
LFEndAngle1 = LeftFlipper1.endangle
RFEndAngle1 = RightFlipper1.endangle

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

'****************************
'   ZKEY: Key Press Handling
'****************************

Sub Table1_KeyDown(ByVal keycode)
  If keycode = PlungerKey Then
    Plunger.PullBack
    SoundPlungerPull
    TimerVRPlunger.Enabled = True
    TimerVRPlunger1.Enabled = False
    VRShooter.Y = 20
  End If

  If keycode = LeftFlipperKey Then
    lfpress = 1
    VRLButton.X = 0 + 6
  End If
  If keycode = RightFlipperKey Then
    rfpress = 1
    VRRButton.X = 0 - 6
  End If

  If Keycode = StartGameKey Then
    SoundStartButton
    VRStartButton.y=VRStartButton.y-4
  End if

  If keycode = AddCreditKey Or keycode = AddCreditKey2 Then
    Select Case Int(Rnd * 3)
      Case 0
        PlaySound ("Coin_In_1"), 0, CoinSoundLevel, 0, 0.25
      Case 1
        PlaySound ("Coin_In_2"), 0, CoinSoundLevel, 0, 0.25
      Case 2
        PlaySound ("Coin_In_3"), 0, CoinSoundLevel, 0, 0.25
    End Select
  End If

  If vpmKeyDown(keycode) Then Exit Sub
End Sub

Sub Table1_KeyUp(ByVal keycode)
  If KeyCode = PlungerKey Then
    Plunger.Fire
    If BIPL = 1 Then
      SoundPlungerReleaseBall()   'Plunger release sound when there is a ball in shooter lane
    Else
      SoundPlungerReleaseNoBall() 'Plunger release sound when there is no ball in shooter lane
    End If
    TimerVRPlunger.Enabled = False
        TimerVRPlunger1.Enabled = True
    VRShooter.Y = 20
  End If

  If keycode = LeftFlipperKey Then
    VRLButton.X = 0
  End If
  If keycode = RightFlipperKey Then
    VRRButton.X = 0
  End If
  If Keycode = StartGameKey Then
    VRStartButton.y=VRStartButton.y+4
  End if
  If vpmKeyUp(keycode) Then Exit Sub
End Sub

Sub TimerVRPlunger_Timer
  If VRShooter.Y < 130 then
       VRShooter.Y = VRShooter.Y + 5
  End If
End Sub

Sub TimerVRPlunger1_Timer
  VRShooter.Y = 20 + (5* Plunger.Position) -20
End Sub

'***********************************************************************************
'****               Knocker                           ****
'***********************************************************************************
SolCallback(sKnocker) = "SolKnocker"

Sub SolKnocker(Enabled) 'Knocker solenoid
  If enabled Then
    KnockerSolenoid 'Add knocker position object
  End If
End Sub

'****************************
'   ZDRN: Drain, Trough, and Ball Release
'****************************
Dim KickerTopLeftBall, KickerTopLeftPos, BallOnTopRamp
SolCallback(sEjectToShooter)  = "EjectToShooter"
SolCallback(sSaucerTopRight)  = "EjectSaucerTopRight"
SolCallback(sKickerTopLeft)   = "EjectKickerTopLeft"
SolCallback(sKickerTopRight)  = "EjectKickerTopRight"
SolCallback(sOutHole)     = "SolOuthole"


Sub BallRelease_Hit() : Controller.Switch(48) = 1 : UpdateTrough : BIP = BIP + 1 : End Sub
Sub BallRelease_UnHit() : Controller.Switch(48) = 0 :UpdateTrough : End Sub
Sub sw47_Hit() : Controller.Switch(47) = 1 : UpdateTrough : End Sub
Sub sw47_UnHit() : Controller.Switch(47) = 0 : UpdateTrough : End Sub
Sub Drain_UnHit() : Controller.Switch(8) = 0 : UpdateTrough : End Sub

'--------------------Drain and Ball Release------------------

Sub SolOuthole(Enabled)
  If Enabled Then
    'Drain.kick 75, 80
    UpdateTrough
  End If
End Sub

Sub EjectToShooter(Enabled)
  If Enabled = true Then
    RandomSoundBallRelease BallRelease
    BallRelease.kick 60, 12
    UpdateTrough
  End If
End Sub

Sub Drain_Hit()
  Controller.Switch(8) = 1
  RandomSoundDrain Drain
  Drain.kick 75, 80
  BIP = BIP-1
End Sub

Sub UpdateTrough()
  UpdateTroughTimer.Interval = 300
  UpdateTroughTimer.Enabled = 1
End Sub

Sub UpdateTroughTimer_Timer()
  If BallRelease.BallCntOver = 0 Then
    sw47.kick 60, 10
  End If
  'If sw47.BallCntOver = 0 Then Drain.kick 75, 80
  Me.Enabled = 0
End Sub

'------------------------Kickers-----------------------------
Sub EjectSaucerTopRight(enabled)
  If enabled Then
    SoundSaucerKick 1, SaucerTopRight
    bsSaucerTR.ExitSol_On
  End If
End Sub

Sub EjectKickerTopLeft(enabled)
  If enabled Then
    SoundSaucerKick 1, KickerTopLeft
    If Isobject(KickerTopLeftBall) Then
      BallOnTopRamp = True
      KickerTopLeft.TimerInterval = 1
      KickerTopLeft.TimerEnabled = 1
      KickerTopLeftPos = 0
    End If
  End If
End Sub

Sub EjectKickerTopRight(enabled)
  If enabled Then
    SoundSaucerKick 1, KickerTopRight
    bsKickerTR.ExitSol_On
  End If
End Sub

Sub KickerTopLeft_Timer
  If KickerTopLeftPos = 0 then
    Controller.Switch(swKickerTopLeft) = 0
  End If
  KickerTopLeftPos = KickerTopLeftPos + 3

  KickerTopLeftBall.z = KickerTopLeftBall.z + 3
  If KickerTopLeftPos > 160 then
    KickerTopLeftBall.x = KickerTopLeftBall.x + 3
  End If
  If KickerTopLeftPos > 185 then
    KickerTopLeft.Kick 90, 15
    'SoundMetalRamp
    Set KickerTopLeftBall = Nothing
    Me.TimerEnabled = 0
  End If
End Sub

'   ZFLP: Flippers
'****************************

SolCallback(sLRFlipper)="SolRFlipper"
SolCallback(sLLFlipper)="SolLFlipper"
SolCallback(sURFlipper)="SolURFlipper"
SolCallback(sULFlipper)="SolULFlipper"

SolCallback(sFlippersEnable) = "SolEnableBumpers"

' ****************************************************
' flipper subs
' ****************************************************
Const ReflipAngle = 20

Sub SolLFlipper(Enabled)
  If Enabled Then
    FlipperActivate LeftFlipper, LFPress
    LF.fire

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

Sub SolULFlipper(Enabled)
  If Enabled Then
    FlipperActivate LeftFlipper1, LFPress1
    LF.fire
    LeftFlipper1.RotateToEnd

    If leftflipper1.currentangle < leftflipper1.endangle + ReflipAngle Then
      RandomSoundReflipUpLeft LeftFlipper1
    Else
      SoundFlipperUpAttackLeft LeftFlipper1
      RandomSoundFlipperUpLeft LeftFlipper1
    End If

    Else
    FlipperDeActivate LeftFlipper1, LFPress1
    LeftFlipper1.RotateToStart

    If LeftFlipper1.currentangle < LeftFlipper1.startAngle - 5 Then
      RandomSoundFlipperDownLeft LeftFlipper1
    End If
    FlipperLeftHitParm = FlipperUpSoundLevel

    End If
End Sub

Sub SolRFlipper(Enabled)
  If Enabled Then
    FlipperActivate RightFlipper, RFPress
    RF.fire

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

    End If
End Sub

Sub SolURFlipper(Enabled)
  If Enabled Then
    FlipperActivate RightFlipper1, RFPress1
    RightFlipper1.RotateToEnd

    If rightflipper1.currentangle > rightflipper1.endangle - ReflipAngle Then
      RandomSoundReflipUpRight RightFlipper1
    Else
      SoundFlipperUpAttackRight RightFlipper1
      RandomSoundFlipperUpRight RightFlipper1
    End If

    Else
    FlipperDeActivate RightFlipper1, RFPress1
    RightFlipper1.RotateToStart

    If RightFlipper1.currentangle > RightFlipper1.startAngle + 5 Then
      RandomSoundFlipperDownRight RightFlipper1
    End If

    End If
End Sub



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
  CheckLiveCatch ActiveBall, LeftFlipper1, LFCount1, parm
  'LF1.ReProcessBalls ActiveBall
  LeftFlipperCollide parm
End Sub

Sub RightFlipper1_Collide(parm)
  CheckLiveCatch ActiveBall, RightFlipper1, RFCount1, parm
  'RF1.ReProcessBalls ActiveBall
  RightFlipperCollide parm
End Sub


'****************************
'   ZSLG: Slingshots
'****************************
Dim SLLPos, SLRPos, SULPos, SURPos, LStep, RStep, LStep1, RStep1

Sub LeftSlingshot_Slingshot()
  LS.VelocityCorrect(ActiveBall)
    LStep = 0
  vpmTimer.PulseSw(swBottomSlingLeft)
  RandomSoundSlingshotLeft Sling2
  Dim BL
  For Each BL in BP_LLSling2
      BL.visible = True
  Next
  Me.TimerEnabled = 1
End Sub

Sub LeftSlingshot_Timer
  Dim BL
  'Dim x1, x2, y: x1 = True:x2 = False:y = -20
    Select Case LStep
    Case 0: For Each BL in BP_LLSling2 : BL.visible = false : Next : For Each BL in BP_LLSling3 : BL.visible = true : Next
    Case 1: For Each BL in BP_LLSling3 : BL.visible = false : Next : For Each BL in BP_LLSling4 : BL.visible = true : Next
    Case 2: For Each BL in BP_LLSling4 : BL.visible = false : Next : For Each BL in BP_LLSling3 : BL.visible = true : Next
        Case 3: For Each BL in BP_LLSling3 : BL.visible = false : Next : For Each BL in BP_LLSling2 : BL.visible = true : Next
        Case 4: For Each BL in BP_LLSling2 : BL.visible = false : Next : For Each BL in BP_LLSling1 : BL.visible = true : Me.TimerEnabled = 0 : Next
    End Select

    LStep = LStep + 1
End Sub

Sub RightSlingshot_Slingshot()
  RS.VelocityCorrect(ActiveBall)
    RStep = 0
  vpmTimer.PulseSw(swBottomSlingRight)
  RandomSoundSlingshotRight Sling1
  Dim BL
  For Each BL in BP_LRSling2
      BL.visible = True
  Next
  Me.TimerEnabled = 1
End Sub

Sub RightSlingshot_Timer
  Dim BL
  'Dim x1, x2, y: x1 = True:x2 = False:y = -20
    Select Case RStep
    Case 0: For Each BL in BP_LRSling2 : BL.visible = false : Next : For Each BL in BP_LRSling3 : BL.visible = true : Next
    Case 1: For Each BL in BP_LRSling3 : BL.visible = false : Next : For Each BL in BP_LRSling4 : BL.visible = true : Next
    Case 2: For Each BL in BP_LRSling4 : BL.visible = false : Next : For Each BL in BP_LRSling3 : BL.visible = true : Next
        Case 3: For Each BL in BP_LRSling3 : BL.visible = false : Next : For Each BL in BP_LRSling2 : BL.visible = true : Next
        Case 4: For Each BL in BP_LRSling2 : BL.visible = false : Next : For Each BL in BP_LRSling1 : BL.visible = true : Me.TimerEnabled = 0 : Next
    End Select

    RStep = RStep + 1
End Sub

Sub LeftSlingshotU_Slingshot()
  LSU.VelocityCorrect(ActiveBall)
    LStep1 = 0
  vpmTimer.PulseSw(swTopSlingLeft)
  RandomSoundSlingshotLeft Sling2U
  Dim BL
  For Each BL in BP_ULSling2
      BL.visible = True
  Next
  Me.TimerEnabled = 1
End Sub

Sub LeftSlingshotU_Timer
  Dim BL
  'Dim x1, x2, y: x1 = True:x2 = False:y = -20
    Select Case LStep1
    Case 0: For Each BL in BP_ULSling2 : BL.visible = false : Next : For Each BL in BP_ULSling3 : BL.visible = true : Next
    Case 1: For Each BL in BP_ULSling3 : BL.visible = false : Next : For Each BL in BP_ULSling4 : BL.visible = true : Next
    Case 2: For Each BL in BP_ULSling4 : BL.visible = false : Next : For Each BL in BP_ULSling3 : BL.visible = true : Next
        Case 3: For Each BL in BP_ULSling3 : BL.visible = false : Next : For Each BL in BP_ULSling2 : BL.visible = true : Next
        Case 4: For Each BL in BP_ULSling2 : BL.visible = false : Next : For Each BL in BP_ULSling1 : BL.visible = true : Me.TimerEnabled = 0 : Next
    End Select

    LStep1 = LStep1 + 1
End Sub

Sub RightSlingshotU_Slingshot()
  RSU.VelocityCorrect(ActiveBall)
    RStep1 = 0
  vpmTimer.PulseSw(swTopSlingRight)
  RandomSoundSlingshotRight Sling1U
  Dim BL
  For Each BL in BP_URSling2
      BL.visible = True
  Next
  Me.TimerEnabled = 1
End Sub

Sub RightSlingshotU_Timer
  Dim BL
  'Dim x1, x2, y: x1 = True:x2 = False:y = -20
    Select Case RStep1
    Case 0: For Each BL in BP_URSling2 : BL.visible = false : Next : For Each BL in BP_URSling3 : BL.visible = true : Next
    Case 1: For Each BL in BP_URSling3 : BL.visible = false : Next : For Each BL in BP_URSling4 : BL.visible = true : Next
    Case 2: For Each BL in BP_URSling4 : BL.visible = false : Next : For Each BL in BP_URSling3 : BL.visible = true : Next
        Case 3: For Each BL in BP_URSling3 : BL.visible = false : Next : For Each BL in BP_URSling2 : BL.visible = true : Next
        Case 4: For Each BL in BP_URSling2 : BL.visible = false : Next : For Each BL in BP_URSling1 : BL.visible = true : Me.TimerEnabled = 0 : Next
    End Select

    RStep1 = RStep1 + 1
End Sub

Sub Slingshots_Init
  Dim BL
  For Each BL in BP_LLSling2 : BL.visible = false : Next
  For Each BL in BP_LLSling3 : BL.visible = false : Next
  For Each BL in BP_LLSling4 : BL.visible = false : Next

  For Each BL in BP_LRSling2 : BL.visible = false : Next
  For Each BL in BP_LRSling3 : BL.visible = false : Next
  For Each BL in BP_LRSling4 : BL.visible = false : Next

  For Each BL in BP_ULSling2 : BL.visible = false : Next
  For Each BL in BP_ULSling3 : BL.visible = false : Next
  For Each BL in BP_ULSling4 : BL.visible = false : Next

  For Each BL in BP_URSling2 : BL.visible = false : Next
  For Each BL in BP_URSling3 : BL.visible = false : Next
  For Each BL in BP_URSling4 : BL.visible = false : Next
End Sub

Sub SolEnableBumpers(enabled)
  If enabled Then
    'turn on bumpers and maybe lights
    LeftSlingshot.hasHitEvent = True
    RightSlingshot.hasHitEvent = True
    LeftSlingshotU.hasHitEvent = True
    RightSlingshotU.hasHitEvent = True
    If SolControlledLights = 1 Then
      GiON
    End If
  Else
    'turn off bumpers and maybe lights
    LeftSlingshot.hasHitEvent = False
    RightSlingshot.hasHitEvent = False
    LeftSlingshotU.hasHitEvent = False
    RightSlingshotU.hasHitEvent = False
    If SolControlledLights = 1 Then
      GiOFF
    End If
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


'****************************
'   ZSWI: Switches
'****************************

Const swRampEntryC          = 1
Const swRampEntryI          = 2
Const swRampEntryT          = 3
Const swRampEntryY          = 4
Const swOuthole             = 8
Const swRampExitLeft        = 12
Const swRampExitRight       = 13
Const swTiltMe              = 15
Const swShooterLane         = 16
Const swBottomSlingLeft     = 17
Const swBottomSlingRight    = 18
Const swTopSlingLeft        = 19
Const swTopSlingRight       = 20
Const swSingleTargetTop     = 21
Const swMushroomLeft        = 22
Const swMushroomRight       = 23
Const swSpinnerTop          = 24
Const swDropTargetBottom    = 25
Const swDropTargetMiddle    = 26
Const swDropTargetTop       = 27
Const swTopWalls            = 28
Const swKickerTopLeft       = 29
Const swSaucerTopRight      = 30
Const swInlaneLeft          = 31
Const swInlaneRight         = 32
Const swBlueTarget1         = 33
Const swBlueTarget2         = 34
Const swBlueTarget3         = 35
Const swBlueTarget4         = 36
Const swBlueTarget5         = 37
Const swBlueTarget6         = 38
Const swOutlaneLeft         = 39
Const swOutlaneRight        = 40
Const swSingleTargetLeft    = 41
Const swSingleTargetRight   = 42
Const swLowerDock           = 45
Const swUpperDock           = 46
Const swTroughL             = 47
Const swTroughR             = 48

'****************************
'   ZTRI: Triggers
'****************************
Sub sw24_Spin():SoundSpinner sw24:vpmTimer.PulseSw (swSpinnerTop):End Sub
Sub sw1_Hit():vpmTimer.PulseSw swRampEntryC:debug.print "Ramp C":End Sub
Sub sw2_Hit():vpmTimer.PulseSw swRampEntryI:debug.print "Ramp I":End Sub
Sub sw3_Hit():vpmTimer.PulseSw swRampEntryT:debug.print "Ramp T":End Sub
Sub sw4_Hit():vpmTimer.PulseSw swRampEntryY:debug.print "Ramp Y":End Sub
Sub RubberBand004_Hit():vpmTimer.PulseSw(swTopWalls):End Sub
Sub RubberBand005_Hit():vpmTimer.PulseSw(swTopWalls):End Sub
Sub sw12a_Hit():Controller.Switch(swRampExitLeft) = 1:debug.print"RampExitLeft":End Sub
Sub sw12a_UnHit():Controller.Switch(swRampExitLeft) = 0:debug.print"RampExitLeft":End Sub
Sub sw12b_Hit():Controller.Switch(swRampExitLeft) = 1:debug.print"RampExitLeft":End Sub
Sub sw12b_UnHit():Controller.Switch(swRampExitLeft) = 0:debug.print"RampExitLeft":End Sub
Sub sw13a_Hit():Controller.Switch(swRampExitRight) = 1:debug.print"RampExitRight":End Sub
Sub sw13a_UnHit():Controller.Switch(swRampExitRight) = 0:debug.print"RampExitRight":End Sub
Sub sw13b_Hit():Controller.Switch(swRampExitRight) = 1:debug.print"RampExitRight":End Sub
Sub sw13b_UnHit():Controller.Switch(swRampExitRight) = 0:debug.print"RampExitRight":End Sub

dim HitRollover, ROStep, RNStep


Sub sw16_Hit()
  Controller.Switch(swShooterLane) = 1
  BIPL = 1
End Sub
Sub sw16_UnHit()
  Controller.Switch(swShooterLane) = 0
  BIPL=0
End Sub

Sub sw31_Hit()
  Controller.Switch(swInlaneLeft) = 1
End Sub
Sub sw31_UnHit()
  Controller.Switch(swInlaneLeft) = 0
End Sub

Sub sw32_Hit()
  Controller.Switch(swInlaneRight) = 1
End Sub
Sub sw32_UnHit()
  Controller.Switch(swInlaneRight) = 0
End Sub

Sub sw39_Hit()
  Controller.Switch(swOutlaneLeft) = 1
End Sub
Sub sw39_UnHit()
  Controller.Switch(swOutlaneLeft) = 0
End Sub

Sub sw40_Hit()
  Controller.Switch(swOutlaneRight) = 1
End Sub
Sub sw40_UnHit()
  Controller.Switch(swOutlaneRight) = 0
End Sub

'****************************
'   ZTAR: Targets
'****************************
'--------Stand Ups----------
Sub sw21_Hit:STHit 21:End Sub
Sub sw33_Hit:STHit 33:End Sub
Sub sw34_Hit:STHit 34:End Sub
Sub sw35_Hit:STHit 35:End Sub
Sub sw36_Hit:STHit 36:End Sub
Sub sw37_Hit:STHit 37:End Sub
Sub sw38_Hit:STHit 38:End Sub
Sub sw41_Hit:STHit 41:End Sub
Sub sw42_Hit:STHit 42:End Sub

'--------Drop Targets----------
Sub sw25_Hit:DTHit 25: TargetBouncer Activeball, 1:End Sub    'sw25
Sub sw26_Hit:DTHit 26: TargetBouncer Activeball, 1:End Sub     'sw26
Sub sw27_Hit:DTHit 27: TargetBouncer Activeball, 1:End sub    'sw27

SolCallback(sDropTargetsReset) = "SolResetDrops"

Sub SolResetDrops(enabled)
  'debug.print "SolResetDropsL "&enabled
  if enabled then
    DTRaise 25
    RandomSoundDropTargetReset BM_sw25
    DTRaise 26
    RandomSoundDropTargetReset BM_sw26
    DTRaise 27
    RandomSoundDropTargetReset BM_sw27
  end if
End Sub

'****************************
'   ZKIC: Kickers, Saucers
'****************************
Sub SaucerTopRight_Hit:SoundSaucerLock:bsSaucerTR.AddBall 0:End Sub
Sub KickerTopLeft_Hit: SoundSaucerLock:Controller.Switch(swKickerTopLeft) = 1:Set KickerTopLeftBall = ActiveBall:End Sub
Sub KickerTopRight_Hit:SoundSaucerLock:bsKickerTR.AddBall 0:BallOnTopRamp = False:End Sub


'****************************
'   ZBMP: Bumpers
'****************************

'-----------Mushroom Bumpers---------------
Dim sw22step, sw23step
Sub sw22_Hit():sw22step = 0: vpmTimer.PulseSw(swMushroomLeft): Me.TimerInterval = 20:Me.TimerEnabled = 1: End Sub
Sub sw23_Hit():sw23step = 0: vpmTimer.PulseSw(swMushroomRight):Me.TimerInterval = 20:Me.TimerEnabled = 1: End Sub

Sub sw22_Timer()
  dim bl
    Select Case sw22step
        Case 1:For each bl in BP_sw22 : bl.transz = 10:sw22step = 2:Next
        Case 2:For each bl in BP_sw22 : bl.transz = 15:sw22step = 3:Next
        Case 3:For each bl in BP_sw22 : bl.transz = 20:sw22step = 4:Next
        Case 4:For each bl in BP_sw22 : bl.transz = 15:sw22step = 5:Next
    Case 5:For each bl in BP_sw22 : bl.transz = 10:sw22step = 6:Next
    Case 6:For each bl in BP_sw22 : bl.transz = 5:sw22step = 7:Next
    Case 7:For each bl in BP_sw22 : bl.transz = 0:sw22step = 8:Next
    Case 8:For each bl in BP_sw22 : bl.transz = -5:sw22step = 9:Next
    Case 9:For each bl in BP_sw22 : bl.transz = 0:sw22step = 1:Next:Me.TimerEnabled = 0
    End Select
  sw22step = sw22step +1
End Sub

Sub sw23_Timer()
  dim bl
    Select Case sw23step
        Case 1:For each bl in BP_sw23 : bl.transz = 10:sw23step = 2:Next
        Case 2:For each bl in BP_sw23 : bl.transz = 15:sw23step = 3:Next
        Case 3:For each bl in BP_sw23 : bl.transz = 20:sw23step = 4:Next
        Case 4:For each bl in BP_sw23 : bl.transz = 15:sw23step = 5:Next
    Case 5:For each bl in BP_sw23 : bl.transz = 10:sw23step = 6:Next
    Case 6:For each bl in BP_sw23 : bl.transz = 5:sw23step = 7:Next
    Case 7:For each bl in BP_sw23 : bl.transz = 0:sw23step = 8:Next
    Case 8:For each bl in BP_sw23 : bl.transz = -5:sw23step = 9:Next
    Case 9:For each bl in BP_sw23 : bl.transz = 0:sw23step = 1:Next:Me.TimerEnabled = 0
    End Select
  sw23step = sw23step +1
End Sub

'****************************
'   ZLNS: Lane Steering
'****************************
Dim LSstep, LSdir
SolCallback(sLaneSteering)    = "SetLaneSteering"

Sub SetLaneSteering(enabled)
  If enabled = True Then
    LSDir = 1
    PlaySoundAtLevelStatic SoundFX("Saucer_Empty", DOFContactors), SaucerKickSoundLevel, BM_LaneSteeringWall3
    LaneSteeringTimer.Enabled = 1
    LanePhysical.collidable=0
    LeftSecretGuide.collidable=0
    RightSecretGuide.collidable=0
    debug.print "LSDIR -1"
  Else
    LSDir = -1
    PlaySoundAtLevelStatic SoundFX("Saucer_Empty", DOFContactors), SaucerKickSoundLevel, BM_LaneSteeringWall3
    LaneSteeringTimer.Enabled = 1
    LanePhysical.collidable=1
    LeftSecretGuide.collidable=1
    RightSecretGuide.collidable=1
    debug.print "LSDIR 1"
  End If
End Sub

Sub LaneSteeringTimer_Timer()
  dim bl
  If LSdir = 1 Then
  LSStep = LSstep+1
    Select Case LSstep
        Case 1:For each bl in BP_LaneSteeringWall3 : bl.transy = -10:Next
        Case 2:For each bl in BP_LaneSteeringWall3 : bl.transy = -20:Next
        Case 3:For each bl in BP_LaneSteeringWall3 : bl.transy = -25:Next
        Case 4:For each bl in BP_LaneSteeringWall3 : bl.transy = -32:Next:Me.enabled=0
  End Select
  Else
    LSstep = LSstep - 1
    Select Case LSstep
        Case 3:For each bl in BP_LaneSteeringWall3 : bl.transy = -22:Next
        Case 2:For each bl in BP_LaneSteeringWall3 : bl.transy = -12:Next
        Case 1:For each bl in BP_LaneSteeringWall3 : bl.transy = -7:Next
        Case 0:For each bl in BP_LaneSteeringWall3 : bl.transy = 0:Next: Me.enabled=0: LanePhysical.collidable=1: LeftSecretGuide.collidable=1: RightSecretGuide.collidable=1
  End Select
  End If
End Sub

Sub LaneSteeringInit
  LSDir = 0: LSstep = 0
End Sub





'******************************************************
'   HELPER FUNCTIONS
'******************************************************

Function NullFunctionZ(aEnabled):End Function '1 argument null function placeholder  TODO move me or replac eme

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
Dim LSU
Set LSU = New SlingshotCorrection
Dim RSU
Set RSU = New SlingshotCorrection

InitSlingCorrection

Sub InitSlingCorrection
  LS.Object = LeftSlingshot
  LSU.Object = LeftSlingshotU
  LS.EndPoint1 = EndPoint1LS
  LS.EndPoint2 = EndPoint2LS
  LSU.EndPoint1 = EndPoint1LSU
  LSU.EndPoint2 = EndPoint2LSU

  RS.Object = RightSlingshot
  RSU.Object = RightSlingshotU
  RS.EndPoint1 = EndPoint1RS
  RS.EndPoint2 = EndPoint2RS
  RSU.EndPoint1 = EndPoint1RSU
  RSU.EndPoint2 = EndPoint2RSU

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
  a = Array(LS, RS, LSU, RSU)
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
'   ZRDT: DROP TARGETS
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
Dim DT1, DT2, DT3

'Set array with drop target objects
'
'DropTargetvar = Array(primary, secondary, prim, swtich, animate)
'   primary:  primary target wall to determine drop
'   secondary:  wall used to simulate the ball striking a bent or offset target after the initial Hit
'   prim:    primitive target used for visuals and animation
'          IMPORTANT!!!
'          rotz must be used for orientation
'          rotx to bend the target back
'          transz to move it up and down
'          the pivot point should be in the center of the target on the x, y and at or below the playfield (0) on z
'   switch:  ROM switch number
'   animate:  Array slot for handling the animation instrucitons, set to 0
'          Values for animate: 1 - bend target (hit to primary), 2 - drop target (hit to secondary), 3 - brick target (high velocity hit to secondary), -1 - raise target
'   isDropped:  Boolean which determines whether a drop target is dropped. Set to false if they are initially raised, true if initially dropped.
'         Use the function DTDropped(switchid) to check a target's drop status.

Set DT1 = (new DropTarget)(sw25, sw25a, BM_sw25, 25, 0, False)
Set DT2 = (new DropTarget)(sw26, sw26a, BM_sw26, 26, 0, False)
Set DT3 = (new DropTarget)(sw27, sw27a, BM_sw27, 27, 0, False)


Dim DTArray
DTArray = Array(DT1, DT2, DT3)

'Configure the behavior of Drop Targets.
Const DTDropSpeed = 90 'in milliseconds
Const DTDropUpSpeed = 40 'in milliseconds
Const DTDropUnits = 44 'VP units primitive drops so top of at or below the playfield
Const DTDropUpUnits = 10 'VP units primitive raises above the up position on drops up
Const DTMaxBend = 8 'max degrees primitive rotates when hit
Const DTDropDelay = 20 'time in milliseconds before target drops (due to friction/impact of the ball)
Const DTRaiseDelay = 40 'time in milliseconds before target drops back to normal up position after the solenoid fires to raise the target
Const DTBrickVel = 30 'velocity at which the target will brick, set to '0' to disable brick
Const DTEnableBrick = 0 'Set to 0 to disable bricking, 1 to enable bricking
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
      If UsingROM Then
        controller.Switch(Switchid mod 100) = 1
      Else
        DTAction switchid
      End If
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
      Dim gBOT
      gBOT = GetBalls

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
    If UsingROM Then controller.Switch(Switchid mod 100) = 0
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

Sub DTAction(switchid)
  Select Case switchid
    Case 1
      Addscore 1000
      ShadowDT(0).visible = False

    Case 2
      Addscore 1000
      ShadowDT(1).visible = False

    Case 3
      Addscore 1000
      ShadowDT(2).visible = False
  End Select
End Sub


'******************************************************
'****  END DROP TARGETS
'******************************************************

'******************************************************
' ZRST: STAND-UP TARGETS by Rothbauerw
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
Dim ST21, ST33, ST34, ST35, ST36, ST37, ST38, ST41, ST42

'Set array with stand-up target objects
'
'StandupTargetvar = Array(primary, prim, swtich)
'   primary:  vp target to determine target hit
'   prim:    primitive target used for visuals and animation
'          IMPORTANT!!!
'          transy must be used to offset the target animation
'   switch:  ROM switch number
'   animate:  Arrary slot for handling the animation instrucitons, set to 0
'
'You will also need to add a secondary hit object for each stand up (name sw11o, sw12o, and sw13o on the example Table1)
'these are inclined primitives to simulate hitting a bent target and should provide so z velocity on high speed impacts


Set ST21 = (new StandupTarget)(sw21, BM_sw21, 21, 0)
Set ST33 = (new StandupTarget)(sw33, BM_sw33, 33, 0)
Set ST34 = (new StandupTarget)(sw34, BM_sw34, 34, 0)
Set ST35 = (new StandupTarget)(sw35, BM_sw35, 35, 0)
Set ST36 = (new StandupTarget)(sw36, BM_sw36, 36, 0)
Set ST37 = (new StandupTarget)(sw37, BM_sw37, 37, 0)
Set ST38 = (new StandupTarget)(sw38, BM_sw38, 38, 0)
Set ST41 = (new StandupTarget)(sw41, BM_sw41, 41, 0)
Set ST42 = (new StandupTarget)(sw42, BM_sw42, 42, 0)


'Add all the Stand-up Target Arrays to Stand-up Target Animation Array
'   STAnimationArray = Array(ST1, ST2, ....)
Dim STArray
STArray = Array(ST21, ST33, ST34, ST35, ST36, ST37, ST38, ST41, ST42)

'Configure the behavior of Stand-up Targets
Const STAnimStep = 1.5  'vpunits per animation step (control return to Start)
Const STMaxOffset = 9   'max vp units target moves when hit

Const STMass = 0.2    'Mass of the Stand-up Target (between 0 and 1), higher values provide more resistance

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
      vpmTimer.PulseSw switch mod 100
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



'******************************************************
'****   END STAND-UP TARGETS
'******************************************************

'***************************************************************
' ZSHA: Ambient ball shadows
'***************************************************************

' For dynamic ball shadows, Check the "Raytraced ball shadows" box for the specific light.
' Also make sure the light's z position is around 25 (mid ball)

'Ambient (Room light source)
Const AmbientBSFactor = 0.4    '0 To 1, higher is darker
Const AmbientMovement = 1    '1+ higher means more movement as the ball moves left and right
Const offsetX = 0        'Offset x position under ball (These are if you want to change where the "room" light is for calculating the shadow position,)
Const offsetY = 0        'Offset y position under ball (^^for example 5,5 if the light is in the back left corner)

' *** Trim or extend these to match the number of balls/primitives/flashers on the table!  (will throw errors if there aren't enough objects)
Dim objBallShadow(2)

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
    If gBOT(s).Z > 20 And gBOT(s).Z < 30 Then
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
'****  END BALL ROLLING AND DROP SOUNDS
'******************************************************




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
Dim RampBalls(5,2)
'x,0 = ball x,1 = ID, 2 = Protection against ending early (minimum amount of updates)

'0,0 is boolean on/off, 0,1 unused for now
RampBalls(0,0) = False

' RampType
' Setup: Set this array to the number Total number of balls that can be tracked at one time + 1.  5 ball multiball then set value to 6
' Description: Array type indexed on BallId and a values used to deterimine what type of ramp the ball is on: False = Wire Ramp, True = Plastic Ramp
Dim RampType(5)

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

'Ramp triggers
Sub WireRamp1Start_Hit
  WireRampOn false
End Sub

Sub WireRamp1End_Hit
    WireRampOff
End Sub

Sub WireRamp2Start_Hit
  WireRampOn false
End Sub

Sub WireRamp2End_Hit
    WireRampOff
End Sub

Sub WireRamp3Start_Hit
  WireRampOn false
End Sub

Sub WireRamp3End_Hit
    WireRampOff
End Sub

Sub WireRamp4Start_Hit
  WireRampOn false
End Sub

Sub WireRamp4End_Hit
    WireRampOff
End Sub

Sub RampTriggerC_Hit
  WireRampOn true
End Sub

Sub RampTriggerI_Hit
  WireRampOn true
End Sub

Sub RampTriggerT_Hit
  WireRampOn true
End Sub

Sub RampTriggerY_Hit
  WireRampOn true
End Sub

Sub VUKRampStart_Hit
  WireRampOn false
End Sub

Sub VUKRampEnd_Hit
  WireRampOff
End Sub

Sub PlungeTrigger_Hit
  WireRampOn true
End Sub

Sub RampTriggerStop1_Hit
  WireRampOff
End Sub

Sub RampTriggerStop2_Hit
  WireRampOff
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

'****************************
'   ZDTD: Desktop Digits
'****************************

Dim Digits(27)
Digits(0)=Array(A1,A2,A3,A4,A5,A6,A7,A10,A8,A9)
Digits(1)=Array(A11,A12,A13,A14,A15,A16,A17,A20,A18,A19)
Digits(2)=Array(A31,A32,A33,A34,A35,A36,A37,A40,A38,A39)
Digits(3)=Array(A41,A42,A43,A44,A45,A46,A47,A50,A48,A49)
Digits(4)=Array(A51,A52,A53,A54,A55,A56,A57,A60,A58,A59)
Digits(5)=Array(A61,A62,A63,A64,A65,A66,A67,A70,A68,A69)
Digits(6)=Array(A71,A72,A73,A74,A75,A76,A77,A80,A78,A79)
Digits(7)=Array(A81,A82,A83,A84,A85,A86,A87,A90,A88,A89)
Digits(8)=Array(A91,A92,A93,A94,A95,A96,A97,A100,A98,A99)
Digits(9)=Array(A101,A102,A103,A104,A105,A106,A107,A110,A108,A109)
Digits(10)=Array(A111,A112,A113,A114,A115,A116,A117,A120,A118,A119)
Digits(11)=Array(A121,A122,A123,A124,A125,A126,A127,A130,A128,A129)
Digits(12)=Array(A131,A132,A133,A134,A135,A136,A137,A140,A138,A139)
Digits(13)=Array(A141,A142,A143,A144,A145,A146,A147,A150,A148,A149)
Digits(14)=Array(A151,A152,A153,A154,A155,A156,A157,A160,A158,A159)
Digits(15)=Array(A161,A162,A163,A164,A165,A166,A167,A170,A168,A169)
Digits(16)=Array(A171,A172,A173,A174,A175,A176,A177,A180,A178,A179)
Digits(17)=Array(A181,A182,A183,A184,A185,A186,A187,A190,A188,A189)
Digits(18)=Array(A191,A192,A193,A194,A195,A196,A197,A200,A198,A199)
Digits(19)=Array(A201,A202,A203,A204,A205,A206,A207,A210,A208,A209)
Digits(20)=Array(A211,A212,A213,A214,A215,A216,A217,A220,A218,A219)
Digits(21)=Array(A221,A222,A223,A224,A225,A226,A227,A230,A228,A229)
Digits(22)=Array(A231,A232,A233,A234,A235,A236,A237,A240,A238,A239)
Digits(23)=Array(A241,A242,A243,A244,A245,A246,A247,A250,A248,A249)
Digits(24)=Array(A251,A252,A253,A254,A255,A256,A257,A260,A258,A259)
Digits(25)=Array(A261,A262,A263,A264,A265,A266,A267,A270,A268,A269)
Digits(26)=Array(A271,A272,A273,A274,A275,A276,A277,A280,A278,A279)
Digits(27)=Array(A281,A282,A283,A284,A285,A286,A287,A290,A288,A289)

Sub DisplayTimer()
    Dim chgLED, ii, num, chg, stat, obj
  chgLED = Controller.ChangedLEDs(&Hffffffff, &Hffffffff)
  If Not IsEmpty(chgLED) Then
    If DesktopMode Then
      For ii = 0 To UBound(chgLED)
        num = chgLED(ii, 0) : chg = chgLED(ii, 1) : stat = chgLED(ii, 2)
        if (num < 32) then
          For Each obj In Digits(num)
            If chg And 1 Then obj.State = stat And 1
            chg = chg\2 : stat = stat\2
          Next
        End If
      Next
    End If
    End If
End Sub

'****************************
'   ZVRR: VR Room / VR Cabinet
'****************************

Dim VRThings
Dim BGThings
Dim DTThings

'Desktop Mode
if VRMode = 0 and cab_mode = 0 Then
  for each VRThings in VRStuff:VRThings.visible = 0:Next
  For each DTThings in DTBackglass: DTThings.Visible = 1: Next

' Cabinet Mode
Elseif VRMode = 0 and cab_mode = 1 Then
  for each VRThings in VRStuff:VRThings.visible = 0:Next
  For each DTThings in DTBackglass: DTThings.Visible = 0: Next

' VR Mode
Else
  for each VRThings in VRStuff:VRThings.visible = 1:Next
  For each DTThings in DTBackglass: DTThings.Visible = 0: Next
End If


'--------------------------------Setup Backglass---------------------------------


Dim xoff,yoff1, yoff2, yoff3, yoff4, yoff5,zoff,xrot,zscale, xcen,ycen

Sub setup_backglass()

  xoff = -20
  yoff1 = -175 ' this is where you adjust the forward/backward position for player scores
  zoff = 698
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

'-------------------------------Display Output---------------------------------

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
    Dim ChgLED, ii, jj, num, chg, stat, obj, b, x
    ChgLED=Controller.ChangedLEDs(&Hffffffff, &Hffffffff)
    If Not IsEmpty(ChgLED)Then
       For ii=0 To UBound(chgLED)
          num=chgLED(ii, 0) : chg=chgLED(ii, 1) : stat=chgLED(ii, 2)
              For Each obj In VRDigits(num)
                   If chg And 1 Then obj.visible=stat And 1    'if you use the object color for off; turn the display object visible to not visible on the playfield, and uncomment this line out.
           If chg And 1 Then FadeDisplay obj, stat And 1
                   chg=chg\2 : stat=stat\2
              Next
        Next
    End If
 End Sub


Sub FadeDisplay(object, onoff)
' If OnOff = 1 Then
    object.color = DisplayColor
    Object.Opacity = 12
' Else
'   Object.Color = RGB(1,1,1)
'   Object.Opacity = 6
' End If
End Sub

'****************************
' ZVLM: VLM Arrays
'****************************

' VLM  Arrays - Start
' Arrays per baked part
Dim BP_BackPanel: BP_BackPanel=Array(BM_BackPanel, LM_F_l29_BackPanel, LM_F_l43_BackPanel, LM_F_l57_BackPanel, LM_F_l58_BackPanel, LM_F_l59_BackPanel, LM_F_l60_BackPanel, LM_F_l61_BackPanel, LM_F_l73_BackPanel, LM_F_l75_BackPanel, LM_F_l79_BackPanel, LM_F_l89_BackPanel, LM_F_l91_BackPanel, LM_F_l92_BackPanel, LM_F_l95_BackPanel, LM_GI_BackPanel, LM_All_Lights_l63_BackPanel)
Dim BP_GateRampC2Wire_002: BP_GateRampC2Wire_002=Array(BM_GateRampC2Wire_002, LM_F_l76_GateRampC2Wire_002, LM_GI_GateRampC2Wire_002, LM_All_Lights_l78b_GateRampC2Wi)
Dim BP_GateRampCWire_001: BP_GateRampCWire_001=Array(BM_GateRampCWire_001, LM_F_l76_GateRampCWire_001, LM_F_l44_GateRampCWire_001, LM_GI_GateRampCWire_001)
Dim BP_GateRampIWire_002: BP_GateRampIWire_002=Array(BM_GateRampIWire_002, LM_F_l29_GateRampIWire_002, LM_F_l61_GateRampIWire_002, LM_F_l76_GateRampIWire_002, LM_F_l44_GateRampIWire_002, LM_GI_GateRampIWire_002, LM_All_Lights_l37_GateRampIWire)
Dim BP_GateRampTWire_002: BP_GateRampTWire_002=Array(BM_GateRampTWire_002, LM_F_l61_GateRampTWire_002, LM_F_l76_GateRampTWire_002, LM_F_l44_GateRampTWire_002, LM_GI_GateRampTWire_002, LM_All_Lights_l24_GateRampTWire)
Dim BP_GateRampY2Wire_002: BP_GateRampY2Wire_002=Array(BM_GateRampY2Wire_002, LM_F_l44_GateRampY2Wire_002, LM_GI_GateRampY2Wire_002, LM_All_Lights_l11_GateRampY2Wir, LM_All_Lights_l94b_GateRampY2Wi)
Dim BP_GateRampYWire_001: BP_GateRampYWire_001=Array(BM_GateRampYWire_001, LM_F_l76_GateRampYWire_001, LM_F_l44_GateRampYWire_001, LM_GI_GateRampYWire_001, LM_All_Lights_l11_GateRampYWire, LM_L_l11_GateRampYWire_001, LM_All_Lights_l42_GateRampYWire, LM_All_Lights_l5_GateRampYWire_)
Dim BP_LLSling1: BP_LLSling1=Array(BM_LLSling1, LM_F_l12_LLSling1, LM_F_l91_LLSling1, LM_GI_LLSling1, LM_All_Lights_l22_LLSling1, LM_All_Lights_l53_LLSling1, LM_All_Lights_l56_LLSling1, LM_All_Lights_l67_LLSling1, LM_L_l67_LLSling1, LM_All_Lights_l70_LLSling1, LM_All_Lights_l85_LLSling1, LM_All_Lights_l9_LLSling1)
Dim BP_LLSling2: BP_LLSling2=Array(BM_LLSling2, LM_F_l12_LLSling2, LM_F_l91_LLSling2, LM_GI_LLSling2, LM_All_Lights_l22_LLSling2, LM_All_Lights_l25_LLSling2, LM_All_Lights_l53_LLSling2, LM_All_Lights_l56_LLSling2, LM_All_Lights_l67_LLSling2, LM_L_l67_LLSling2, LM_All_Lights_l70_LLSling2, LM_All_Lights_l85_LLSling2, LM_All_Lights_l9_LLSling2)
Dim BP_LLSling3: BP_LLSling3=Array(BM_LLSling3, LM_F_l12_LLSling3, LM_F_l91_LLSling3, LM_GI_LLSling3, LM_All_Lights_l22_LLSling3, LM_All_Lights_l25_LLSling3, LM_All_Lights_l53_LLSling3, LM_All_Lights_l56_LLSling3, LM_All_Lights_l67_LLSling3, LM_L_l67_LLSling3, LM_All_Lights_l70_LLSling3, LM_All_Lights_l85_LLSling3, LM_All_Lights_l86_LLSling3, LM_All_Lights_l9_LLSling3)
Dim BP_LLSling4: BP_LLSling4=Array(BM_LLSling4, LM_F_l12_LLSling4, LM_F_l91_LLSling4, LM_GI_LLSling4, LM_All_Lights_l22_LLSling4, LM_All_Lights_l25_LLSling4, LM_All_Lights_l53_LLSling4, LM_All_Lights_l56_LLSling4, LM_All_Lights_l67_LLSling4, LM_L_l67_LLSling4, LM_All_Lights_l70_LLSling4, LM_All_Lights_l85_LLSling4, LM_All_Lights_l86_LLSling4, LM_All_Lights_l9_LLSling4)
Dim BP_LRSling1: BP_LRSling1=Array(BM_LRSling1, LM_F_l12_LRSling1, LM_F_l91_LRSling1, LM_GI_LRSling1, LM_All_Lights_l26_LRSling1, LM_All_Lights_l39_LRSling1, LM_All_Lights_l51_LRSling1, LM_L_l51_LRSling1, LM_All_Lights_l54_LRSling1, LM_All_Lights_l69_LRSling1, LM_All_Lights_l71_LRSling1, LM_All_Lights_l87_LRSling1)
Dim BP_LRSling2: BP_LRSling2=Array(BM_LRSling2, LM_F_l12_LRSling2, LM_F_l91_LRSling2, LM_GI_LRSling2, LM_All_Lights_l26_LRSling2, LM_All_Lights_l39_LRSling2, LM_All_Lights_l51_LRSling2, LM_L_l51_LRSling2, LM_All_Lights_l54_LRSling2, LM_All_Lights_l69_LRSling2, LM_All_Lights_l71_LRSling2, LM_All_Lights_l87_LRSling2)
Dim BP_LRSling3: BP_LRSling3=Array(BM_LRSling3, LM_F_l12_LRSling3, LM_F_l91_LRSling3, LM_GI_LRSling3, LM_All_Lights_l10_LRSling3, LM_All_Lights_l26_LRSling3, LM_All_Lights_l39_LRSling3, LM_All_Lights_l51_LRSling3, LM_L_l51_LRSling3, LM_All_Lights_l54_LRSling3, LM_All_Lights_l69_LRSling3, LM_All_Lights_l71_LRSling3, LM_All_Lights_l87_LRSling3)
Dim BP_LRSling4: BP_LRSling4=Array(BM_LRSling4, LM_F_l12_LRSling4, LM_F_l91_LRSling4, LM_GI_LRSling4, LM_All_Lights_l10_LRSling4, LM_All_Lights_l26_LRSling4, LM_All_Lights_l39_LRSling4, LM_All_Lights_l51_LRSling4, LM_L_l51_LRSling4, LM_All_Lights_l54_LRSling4, LM_All_Lights_l69_LRSling4, LM_All_Lights_l71_LRSling4, LM_All_Lights_l87_LRSling4)
Dim BP_LaneSteeringWall3: BP_LaneSteeringWall3=Array(BM_LaneSteeringWall3, LM_F_l29_LaneSteeringWall3, LM_F_l61_LaneSteeringWall3, LM_GI_LaneSteeringWall3)
Dim BP_LaneSwitch1: BP_LaneSwitch1=Array(BM_LaneSwitch1, LM_F_l29_LaneSwitch1, LM_F_l58_LaneSwitch1, LM_F_l61_LaneSwitch1, LM_GI_LaneSwitch1)
Dim BP_LaneSwitch2: BP_LaneSwitch2=Array(BM_LaneSwitch2, LM_F_l29_LaneSwitch2, LM_F_l61_LaneSwitch2, LM_GI_LaneSwitch2)
Dim BP_LaneSwitch3: BP_LaneSwitch3=Array(BM_LaneSwitch3, LM_F_l29_LaneSwitch3, LM_F_l61_LaneSwitch3, LM_GI_LaneSwitch3)
Dim BP_LaneSwitch4: BP_LaneSwitch4=Array(BM_LaneSwitch4, LM_F_l29_LaneSwitch4, LM_F_l43_LaneSwitch4, LM_F_l61_LaneSwitch4, LM_GI_LaneSwitch4)
Dim BP_Layer_1: BP_Layer_1=Array(BM_Layer_1, LM_F_l43_Layer_1, LM_F_l57_Layer_1, LM_F_l58_Layer_1, LM_F_l59_Layer_1, LM_F_l60_Layer_1, LM_F_l61_Layer_1, LM_F_l73_Layer_1, LM_F_l75_Layer_1, LM_F_l79_Layer_1, LM_F_l89_Layer_1, LM_F_l92_Layer_1, LM_F_l95_Layer_1, LM_GI_Layer_1)
Dim BP_Lflip: BP_Lflip=Array(BM_Lflip, LM_F_l12_Lflip, LM_F_l91_Lflip, LM_GI_Lflip)
Dim BP_Lflip1: BP_Lflip1=Array(BM_Lflip1, LM_F_l76_Lflip1, LM_F_l44_Lflip1, LM_GI_Lflip1)
Dim BP_Lflip1U: BP_Lflip1U=Array(BM_Lflip1U, LM_F_l76_Lflip1U, LM_F_l44_Lflip1U, LM_GI_Lflip1U, LM_All_Lights_l3_Lflip1U, LM_L_l3_Lflip1U, LM_All_Lights_l34_Lflip1U)
Dim BP_LflipU: BP_LflipU=Array(BM_LflipU, LM_F_l12_LflipU, LM_F_l91_LflipU, LM_GI_LflipU)
Dim BP_MiddlePost: BP_MiddlePost=Array(BM_MiddlePost, LM_GI_MiddlePost)
Dim BP_NCoverL: BP_NCoverL=Array(BM_NCoverL, LM_F_l12_NCoverL, LM_F_l91_NCoverL, LM_GI_NCoverL)
Dim BP_NCoverR: BP_NCoverR=Array(BM_NCoverR, LM_F_l91_NCoverR, LM_GI_NCoverR)
Dim BP_Parts: BP_Parts=Array(BM_Parts, BM_Parts_001, LM_F_l12_Parts, LM_F_l29_Parts, LM_F_l43_Parts, LM_F_l45_Parts, LM_F_l57_Parts, LM_F_l58_Parts, LM_F_l59_Parts, LM_F_l60_Parts, LM_F_l61_Parts, LM_F_l73_Parts, LM_F_l75_Parts, LM_F_l79_Parts, LM_F_l89_Parts, LM_F_l91_Parts, LM_F_l92_Parts, LM_F_l76_Parts, LM_F_l44_Parts, LM_GI_Parts, LM_All_Lights_l1_Parts, LM_L_l1_Parts, LM_All_Lights_l10_Parts, LM_All_Lights_l11_Parts, LM_L_l11_Parts, LM_All_Lights_l12a_Parts, LM_All_Lights_l15_Parts, LM_L_l15_Parts, LM_All_Lights_l17_Parts, LM_All_Lights_l18_Parts, LM_All_Lights_l2_Parts, LM_All_Lights_l20_Parts, LM_All_Lights_l21_Parts, LM_L_l21_Parts, LM_All_Lights_l22_Parts, LM_L_l22_Parts, LM_All_Lights_l23_Parts, LM_All_Lights_l24_Parts, LM_L_l24_Parts, LM_All_Lights_l25_Parts, LM_All_Lights_l26_Parts, LM_L_l26_Parts, LM_All_Lights_l27_Parts, LM_L_l27_Parts, LM_All_Lights_l29a_Parts, LM_All_Lights_l3_Parts, LM_L_l3_Parts, LM_All_Lights_l31_Parts, LM_L_l31_Parts, LM_All_Lights_l33_Parts, LM_All_Lights_l34_Parts, _
  LM_L_l34_Parts, LM_All_Lights_l35_Parts, LM_All_Lights_l36_Parts, LM_All_Lights_l37_Parts, LM_L_l37_Parts, LM_All_Lights_l38_Parts, LM_All_Lights_l39_Parts, LM_All_Lights_l4_Parts, LM_All_Lights_l40_Parts, LM_L_l40_Parts, LM_All_Lights_l41_Parts, LM_All_Lights_l42_Parts, LM_L_l42_Parts, LM_All_Lights_l49_Parts, LM_L_l49_Parts, LM_All_Lights_l5_Parts, LM_L_l5_Parts, LM_All_Lights_l50_Parts, LM_L_l50_Parts, LM_All_Lights_l51_Parts, LM_L_l51_Parts, LM_All_Lights_l52_Parts, LM_L_l52_Parts, LM_All_Lights_l53_Parts, LM_L_l53_Parts, LM_All_Lights_l54_Parts, LM_L_l54_Parts, LM_All_Lights_l55_Parts, LM_All_Lights_l56_Parts, LM_All_Lights_l61a_Parts, LM_All_Lights_l63_Parts, LM_L_l63_Parts, LM_All_Lights_l65_Parts, LM_L_l65_Parts, LM_All_Lights_l66_Parts, LM_L_l66_Parts, LM_All_Lights_l67_Parts, LM_L_l67_Parts, LM_All_Lights_l68_Parts, LM_L_l68_Parts, LM_All_Lights_l69_Parts, LM_L_l69_Parts, LM_All_Lights_l7_Parts, LM_All_Lights_l70_Parts, LM_L_l70_Parts, LM_All_Lights_l71_Parts, LM_All_Lights_l72_Parts, _
  LM_All_Lights_l78_Parts, LM_All_Lights_l78b_Parts, LM_All_Lights_l8_Parts, LM_L_l8_Parts, LM_All_Lights_l81_Parts, LM_All_Lights_l82_Parts, LM_L_l82_Parts, LM_All_Lights_l83_Parts, LM_L_l83_Parts, LM_All_Lights_l84_Parts, LM_L_l84_Parts, LM_All_Lights_l85_Parts, LM_L_l85_Parts, LM_All_Lights_l86_Parts, LM_All_Lights_l87_Parts, LM_L_l87_Parts, LM_All_Lights_l88_Parts, LM_All_Lights_l9_Parts, LM_L_l9_Parts, LM_All_Lights_l90_Parts, LM_L_l90_Parts, LM_All_Lights_l91a_Parts, LM_All_Lights_l94_Parts, LM_All_Lights_l94b_Parts)
Dim BP_Plastics: BP_Plastics=Array(BM_Plastics, LM_F_l12_Plastics, LM_F_l29_Plastics, LM_F_l43_Plastics, LM_F_l59_Plastics, LM_F_l61_Plastics, LM_F_l75_Plastics, LM_F_l91_Plastics, LM_F_l76_Plastics, LM_F_l44_Plastics, LM_GI_Plastics, LM_All_Lights_l10_Plastics, LM_All_Lights_l12a_Plastics, LM_All_Lights_l15_Plastics, LM_All_Lights_l21_Plastics, LM_L_l21_Plastics, LM_All_Lights_l25_Plastics, LM_All_Lights_l26_Plastics, LM_All_Lights_l27_Plastics, LM_L_l27_Plastics, LM_All_Lights_l29a_Plastics, LM_All_Lights_l3_Plastics, LM_All_Lights_l34_Plastics, LM_All_Lights_l37_Plastics, LM_All_Lights_l40_Plastics, LM_L_l40_Plastics, LM_All_Lights_l41_Plastics, LM_All_Lights_l49_Plastics, LM_All_Lights_l50_Plastics, LM_All_Lights_l51_Plastics, LM_L_l51_Plastics, LM_All_Lights_l54_Plastics, LM_All_Lights_l61a_Plastics, LM_All_Lights_l65_Plastics, LM_All_Lights_l66_Plastics, LM_All_Lights_l67_Plastics, LM_L_l67_Plastics, LM_All_Lights_l69_Plastics, LM_All_Lights_l78_Plastics, LM_All_Lights_l78b_Plastics, _
  LM_All_Lights_l8_Plastics, LM_All_Lights_l81_Plastics, LM_All_Lights_l82_Plastics, LM_All_Lights_l84_Plastics, LM_All_Lights_l9_Plastics, LM_All_Lights_l90_Plastics, LM_L_l90_Plastics, LM_All_Lights_l91a_Plastics, LM_All_Lights_l94_Plastics, LM_All_Lights_l94b_Plastics)
Dim BP_Playfield: BP_Playfield=Array(BM_Playfield, BM_Playfield_001, LM_F_l12_Playfield, LM_F_l29_Playfield, LM_F_l61_Playfield, LM_F_l91_Playfield, LM_F_l76_Playfield, LM_F_l44_Playfield, LM_GI_Playfield, LM_GI_Playfield_001, LM_L_l1_Playfield, LM_L_l10_Playfield, LM_L_l11_Playfield, LM_All_Lights_l12a_Playfield, LM_All_Lights_l15_Playfield, LM_L_l17_Playfield, LM_L_l18_Playfield, LM_L_l2_Playfield, LM_L_l20_Playfield, LM_All_Lights_l21_Playfield, LM_L_l21_Playfield, LM_L_l22_Playfield, LM_L_l23_Playfield, LM_L_l25_Playfield, LM_L_l26_Playfield, LM_All_Lights_l27_Playfield, LM_L_l27_Playfield, LM_All_Lights_l29a_Playfield, LM_L_l3_Playfield, LM_All_Lights_l31_Playfield, LM_L_l33_Playfield, LM_L_l34_Playfield, LM_L_l35_Playfield, LM_L_l36_Playfield, LM_All_Lights_l37_Playfield, LM_L_l37_Playfield, LM_L_l38_Playfield, LM_L_l39_Playfield, LM_L_l4_Playfield, LM_All_Lights_l40_Playfield, LM_L_l40_Playfield, LM_L_l41_Playfield, LM_L_l42_Playfield, LM_All_Lights_l49_Playfield, LM_L_l49_Playfield, LM_L_l5_Playfield, _
  LM_All_Lights_l50_Playfield, LM_L_l50_Playfield, LM_All_Lights_l51_Playfield, LM_All_Lights_l53_Playfield, LM_L_l53_Playfield, LM_L_l54_Playfield, LM_All_Lights_l61a_Playfield, LM_All_Lights_l63_Playfield, LM_All_Lights_l65_Playfield, LM_L_l65_Playfield, LM_All_Lights_l66_Playfield, LM_L_l66_Playfield, LM_All_Lights_l67_Playfield, LM_L_l67_Playfield, LM_All_Lights_l69_Playfield, LM_L_l69_Playfield, LM_L_l7_Playfield, LM_All_Lights_l78_Playfield, LM_All_Lights_l78b_Playfield, LM_All_Lights_l8_Playfield, LM_L_l8_Playfield, LM_All_Lights_l81_Playfield, LM_L_l81_Playfield, LM_All_Lights_l82_Playfield, LM_L_l82_Playfield, LM_L_l84_Playfield, LM_All_Lights_l85_Playfield, LM_L_l85_Playfield, LM_L_l87_Playfield, LM_L_l9_Playfield, LM_All_Lights_l90_Playfield, LM_All_Lights_l91a_Playfield, LM_All_Lights_l94_Playfield, LM_All_Lights_l94b_Playfield)
Dim BP_Rflip: BP_Rflip=Array(BM_Rflip, LM_F_l12_Rflip, LM_F_l91_Rflip, LM_GI_Rflip)
Dim BP_Rflip1: BP_Rflip1=Array(BM_Rflip1, LM_F_l76_Rflip1, LM_F_l44_Rflip1, LM_GI_Rflip1)
Dim BP_Rflip1U: BP_Rflip1U=Array(BM_Rflip1U, LM_F_l76_Rflip1U, LM_F_l44_Rflip1U, LM_GI_Rflip1U, LM_All_Lights_l11_Rflip1U, LM_L_l11_Rflip1U, LM_All_Lights_l42_Rflip1U)
Dim BP_RflipU: BP_RflipU=Array(BM_RflipU, LM_F_l12_RflipU, LM_F_l91_RflipU, LM_GI_RflipU, LM_All_Lights_l19_RflipU)
Dim BP_SideInserts: BP_SideInserts=Array(BM_SideInserts, LM_F_l29_SideInserts, LM_F_l57_SideInserts, LM_F_l58_SideInserts, LM_F_l60_SideInserts, LM_F_l61_SideInserts, LM_F_l73_SideInserts, LM_F_l79_SideInserts, LM_F_l89_SideInserts, LM_F_l92_SideInserts, LM_F_l95_SideInserts, LM_GI_SideInserts, LM_All_Lights_l79_SideInserts)
Dim BP_Sling1: BP_Sling1=Array(BM_Sling1, LM_F_l12_Sling1, LM_F_l91_Sling1, LM_GI_Sling1)
Dim BP_Sling2: BP_Sling2=Array(BM_Sling2, LM_F_l12_Sling2, LM_F_l91_Sling2, LM_GI_Sling2)
Dim BP_ULSling: BP_ULSling=Array(BM_ULSling, LM_F_l61_ULSling, LM_F_l76_ULSling, LM_F_l44_ULSling, LM_GI_ULSling, LM_All_Lights_l21_ULSling, LM_All_Lights_l37_ULSling, LM_All_Lights_l78_ULSling, LM_All_Lights_l78b_ULSling)
Dim BP_ULSling1: BP_ULSling1=Array(BM_ULSling1, LM_F_l29_ULSling1, LM_F_l61_ULSling1, LM_F_l76_ULSling1, LM_F_l44_ULSling1, LM_GI_ULSling1, LM_All_Lights_l15_ULSling1, LM_All_Lights_l21_ULSling1, LM_All_Lights_l37_ULSling1, LM_All_Lights_l49_ULSling1, LM_All_Lights_l52_ULSling1, LM_All_Lights_l65_ULSling1, LM_All_Lights_l78_ULSling1, LM_All_Lights_l78b_ULSling1, LM_All_Lights_l81_ULSling1, LM_All_Lights_l84_ULSling1)
Dim BP_ULSling2: BP_ULSling2=Array(BM_ULSling2, LM_F_l29_ULSling2, LM_F_l61_ULSling2, LM_F_l76_ULSling2, LM_F_l44_ULSling2, LM_GI_ULSling2, LM_All_Lights_l15_ULSling2, LM_All_Lights_l21_ULSling2, LM_All_Lights_l37_ULSling2, LM_L_l37_ULSling2, LM_All_Lights_l49_ULSling2, LM_All_Lights_l52_ULSling2, LM_All_Lights_l65_ULSling2, LM_All_Lights_l78_ULSling2, LM_All_Lights_l78b_ULSling2, LM_All_Lights_l81_ULSling2, LM_All_Lights_l84_ULSling2)
Dim BP_ULSling3: BP_ULSling3=Array(BM_ULSling3, LM_F_l29_ULSling3, LM_F_l61_ULSling3, LM_F_l76_ULSling3, LM_F_l44_ULSling3, LM_GI_ULSling3, LM_All_Lights_l15_ULSling3, LM_All_Lights_l21_ULSling3, LM_All_Lights_l37_ULSling3, LM_L_l37_ULSling3, LM_All_Lights_l49_ULSling3, LM_All_Lights_l52_ULSling3, LM_All_Lights_l65_ULSling3, LM_All_Lights_l78_ULSling3, LM_All_Lights_l78b_ULSling3, LM_All_Lights_l81_ULSling3, LM_All_Lights_l84_ULSling3)
Dim BP_ULSling4: BP_ULSling4=Array(BM_ULSling4, LM_F_l29_ULSling4, LM_F_l61_ULSling4, LM_F_l76_ULSling4, LM_F_l44_ULSling4, LM_GI_ULSling4, LM_All_Lights_l15_ULSling4, LM_All_Lights_l21_ULSling4, LM_All_Lights_l37_ULSling4, LM_L_l37_ULSling4, LM_All_Lights_l49_ULSling4, LM_All_Lights_l52_ULSling4, LM_All_Lights_l65_ULSling4, LM_All_Lights_l78_ULSling4, LM_All_Lights_l78b_ULSling4, LM_All_Lights_l81_ULSling4, LM_All_Lights_l84_ULSling4)
Dim BP_URSling: BP_URSling=Array(BM_URSling, LM_F_l29_URSling, LM_F_l61_URSling, LM_F_l76_URSling, LM_F_l44_URSling, LM_GI_URSling, LM_All_Lights_l8_URSling, LM_All_Lights_l94_URSling, LM_All_Lights_l94b_URSling)
Dim BP_URSling1: BP_URSling1=Array(BM_URSling1, LM_F_l29_URSling1, LM_F_l61_URSling1, LM_F_l76_URSling1, LM_F_l44_URSling1, LM_GI_URSling1, LM_All_Lights_l24_URSling1, LM_All_Lights_l31_URSling1, LM_All_Lights_l50_URSling1, LM_All_Lights_l52_URSling1, LM_All_Lights_l66_URSling1, LM_All_Lights_l8_URSling1, LM_All_Lights_l82_URSling1, LM_All_Lights_l84_URSling1, LM_All_Lights_l94_URSling1, LM_All_Lights_l94b_URSling1)
Dim BP_URSling2: BP_URSling2=Array(BM_URSling2, LM_F_l29_URSling2, LM_F_l61_URSling2, LM_F_l76_URSling2, LM_F_l44_URSling2, LM_GI_URSling2, LM_All_Lights_l24_URSling2, LM_L_l24_URSling2, LM_All_Lights_l31_URSling2, LM_All_Lights_l50_URSling2, LM_All_Lights_l52_URSling2, LM_All_Lights_l66_URSling2, LM_All_Lights_l8_URSling2, LM_All_Lights_l82_URSling2, LM_All_Lights_l84_URSling2, LM_All_Lights_l94_URSling2, LM_All_Lights_l94b_URSling2)
Dim BP_URSling3: BP_URSling3=Array(BM_URSling3, LM_F_l29_URSling3, LM_F_l61_URSling3, LM_F_l76_URSling3, LM_F_l44_URSling3, LM_GI_URSling3, LM_All_Lights_l24_URSling3, LM_L_l24_URSling3, LM_All_Lights_l31_URSling3, LM_All_Lights_l50_URSling3, LM_All_Lights_l52_URSling3, LM_All_Lights_l66_URSling3, LM_All_Lights_l8_URSling3, LM_All_Lights_l82_URSling3, LM_All_Lights_l84_URSling3, LM_All_Lights_l94_URSling3, LM_All_Lights_l94b_URSling3)
Dim BP_URSling4: BP_URSling4=Array(BM_URSling4, LM_F_l29_URSling4, LM_F_l61_URSling4, LM_F_l76_URSling4, LM_F_l44_URSling4, LM_GI_URSling4, LM_All_Lights_l24_URSling4, LM_L_l24_URSling4, LM_All_Lights_l31_URSling4, LM_All_Lights_l50_URSling4, LM_All_Lights_l52_URSling4, LM_All_Lights_l66_URSling4, LM_All_Lights_l8_URSling4, LM_All_Lights_l82_URSling4, LM_All_Lights_l84_URSling4, LM_All_Lights_l94_URSling4, LM_All_Lights_l94b_URSling4)
Dim BP_UnderPF: BP_UnderPF=Array(BM_UnderPF, LM_F_l12_UnderPF, LM_F_l29_UnderPF, LM_F_l45_UnderPF, LM_F_l61_UnderPF, LM_F_l91_UnderPF, LM_F_l76_UnderPF, LM_F_l44_UnderPF, LM_GI_UnderPF, LM_All_Lights_l1_UnderPF, LM_L_l1_UnderPF, LM_All_Lights_l10_UnderPF, LM_L_l10_UnderPF, LM_All_Lights_l11_UnderPF, LM_L_l11_UnderPF, LM_L_l15_UnderPF, LM_L_l17_UnderPF, LM_All_Lights_l18_UnderPF, LM_L_l18_UnderPF, LM_L_l19_UnderPF, LM_All_Lights_l2_UnderPF, LM_L_l2_UnderPF, LM_L_l20_UnderPF, LM_L_l21_UnderPF, LM_L_l22_UnderPF, LM_All_Lights_l23_UnderPF, LM_L_l23_UnderPF, LM_All_Lights_l24_UnderPF, LM_L_l24_UnderPF, LM_All_Lights_l25_UnderPF, LM_L_l25_UnderPF, LM_All_Lights_l26_UnderPF, LM_L_l26_UnderPF, LM_All_Lights_l27_UnderPF, LM_L_l27_UnderPF, LM_All_Lights_l3_UnderPF, LM_L_l3_UnderPF, LM_L_l31_UnderPF, LM_L_l33_UnderPF, LM_All_Lights_l34_UnderPF, LM_L_l34_UnderPF, LM_L_l35_UnderPF, LM_L_l36_UnderPF, LM_All_Lights_l37_UnderPF, LM_L_l37_UnderPF, LM_All_Lights_l38_UnderPF, LM_L_l38_UnderPF, LM_All_Lights_l39_UnderPF, _
  LM_L_l39_UnderPF, LM_All_Lights_l4_UnderPF, LM_L_l4_UnderPF, LM_All_Lights_l40_UnderPF, LM_L_l40_UnderPF, LM_All_Lights_l41_UnderPF, LM_L_l41_UnderPF, LM_All_Lights_l42_UnderPF, LM_L_l42_UnderPF, LM_All_Lights_l47_UnderPF, LM_L_l47_UnderPF, LM_L_l49_UnderPF, LM_All_Lights_l5_UnderPF, LM_L_l5_UnderPF, LM_L_l50_UnderPF, LM_All_Lights_l51_UnderPF, LM_L_l51_UnderPF, LM_L_l52_UnderPF, LM_All_Lights_l53_UnderPF, LM_L_l53_UnderPF, LM_All_Lights_l54_UnderPF, LM_L_l54_UnderPF, LM_All_Lights_l55_UnderPF, LM_L_l55_UnderPF, LM_All_Lights_l56_UnderPF, LM_L_l56_UnderPF, LM_All_Lights_l63_UnderPF, LM_L_l63_UnderPF, LM_L_l65_UnderPF, LM_L_l66_UnderPF, LM_All_Lights_l67_UnderPF, LM_L_l67_UnderPF, LM_L_l68_UnderPF, LM_All_Lights_l69_UnderPF, LM_L_l69_UnderPF, LM_L_l7_UnderPF, LM_All_Lights_l70_UnderPF, LM_L_l70_UnderPF, LM_L_l71_UnderPF, LM_L_l72_UnderPF, LM_L_l8_UnderPF, LM_L_l81_UnderPF, LM_L_l82_UnderPF, LM_L_l83_UnderPF, LM_All_Lights_l84_UnderPF, LM_L_l84_UnderPF, LM_All_Lights_l85_UnderPF, LM_L_l85_UnderPF, _
  LM_All_Lights_l86_UnderPF, LM_L_l86_UnderPF, LM_All_Lights_l87_UnderPF, LM_L_l87_UnderPF, LM_All_Lights_l88_UnderPF, LM_L_l88_UnderPF, LM_All_Lights_l9_UnderPF, LM_L_l9_UnderPF, LM_L_l90_UnderPF)
Dim BP_pSpinnerRod: BP_pSpinnerRod=Array(BM_pSpinnerRod, LM_F_l29_pSpinnerRod, LM_F_l61_pSpinnerRod, LM_F_l76_pSpinnerRod, LM_F_l44_pSpinnerRod, LM_GI_pSpinnerRod, LM_All_Lights_l24_pSpinnerRod, LM_All_Lights_l94b_pSpinnerRod)
Dim BP_pSpinnerRod_001: BP_pSpinnerRod_001=Array(BM_pSpinnerRod_001, LM_F_l29_pSpinnerRod_001, LM_F_l61_pSpinnerRod_001, LM_F_l76_pSpinnerRod_001, LM_F_l44_pSpinnerRod_001, LM_GI_pSpinnerRod_001, LM_All_Lights_l15_pSpinnerRod_0, LM_L_l15_pSpinnerRod_001, LM_All_Lights_l29a_pSpinnerRod_, LM_All_Lights_l37_pSpinnerRod_0)
Dim BP_pSpinnerRod_002: BP_pSpinnerRod_002=Array(BM_pSpinnerRod_002, LM_F_l44_pSpinnerRod_002, LM_GI_pSpinnerRod_002, LM_All_Lights_l94_pSpinnerRod_0, LM_All_Lights_l94b_pSpinnerRod_)
Dim BP_pSpinnerRod_003: BP_pSpinnerRod_003=Array(BM_pSpinnerRod_003, LM_F_l61_pSpinnerRod_003, LM_F_l76_pSpinnerRod_003, LM_GI_pSpinnerRod_003, LM_All_Lights_l78_pSpinnerRod_0, LM_All_Lights_l78b_pSpinnerRod_)
Dim BP_pSpinnerRod_004: BP_pSpinnerRod_004=Array(BM_pSpinnerRod_004, LM_F_l29_pSpinnerRod_004, LM_F_l61_pSpinnerRod_004, LM_F_l76_pSpinnerRod_004, LM_F_l44_pSpinnerRod_004, LM_GI_pSpinnerRod_004, LM_All_Lights_l29a_pSpinnerRod_, LM_All_Lights_l40_pSpinnerRod_0)
Dim BP_sw16: BP_sw16=Array(BM_sw16)
Dim BP_sw21: BP_sw21=Array(BM_sw21, LM_F_l29_sw21, LM_F_l61_sw21, LM_F_l76_sw21, LM_F_l44_sw21, LM_GI_sw21)
Dim BP_sw22: BP_sw22=Array(BM_sw22, LM_F_l29_sw22, LM_F_l61_sw22, LM_F_l76_sw22, LM_F_l44_sw22, LM_GI_sw22, LM_All_Lights_l15_sw22)
Dim BP_sw23: BP_sw23=Array(BM_sw23, LM_F_l29_sw23, LM_F_l61_sw23, LM_F_l76_sw23, LM_F_l44_sw23, LM_GI_sw23, LM_All_Lights_l31_sw23)
Dim BP_sw24: BP_sw24=Array(BM_sw24, LM_F_l29_sw24, LM_F_l61_sw24, LM_GI_sw24, LM_All_Lights_l27_sw24, LM_L_l27_sw24, LM_All_Lights_l29a_sw24, LM_All_Lights_l40_sw24, LM_L_l40_sw24)
Dim BP_sw25: BP_sw25=Array(BM_sw25, LM_F_l29_sw25, LM_F_l61_sw25, LM_F_l76_sw25, LM_F_l44_sw25, LM_GI_sw25, LM_All_Lights_l15_sw25, LM_All_Lights_l21_sw25, LM_All_Lights_l8_sw25)
Dim BP_sw26: BP_sw26=Array(BM_sw26, LM_F_l29_sw26, LM_F_l61_sw26, LM_F_l76_sw26, LM_F_l44_sw26, LM_GI_sw26, LM_All_Lights_l15_sw26, LM_All_Lights_l27_sw26, LM_All_Lights_l31_sw26, LM_All_Lights_l8_sw26)
Dim BP_sw27: BP_sw27=Array(BM_sw27, LM_F_l29_sw27, LM_F_l61_sw27, LM_F_l76_sw27, LM_F_l44_sw27, LM_GI_sw27, LM_All_Lights_l15_sw27, LM_All_Lights_l27_sw27, LM_All_Lights_l31_sw27, LM_All_Lights_l40_sw27)
Dim BP_sw31_001: BP_sw31_001=Array(BM_sw31_001, LM_F_l12_sw31_001, LM_F_l91_sw31_001, LM_GI_sw31_001, LM_All_Lights_l85_sw31_001)
Dim BP_sw32_001: BP_sw32_001=Array(BM_sw32_001, LM_F_l91_sw32_001, LM_GI_sw32_001, LM_All_Lights_l54_sw32_001)
Dim BP_sw33: BP_sw33=Array(BM_sw33, LM_F_l76_sw33, LM_F_l44_sw33, LM_GI_sw33, LM_All_Lights_l49_sw33)
Dim BP_sw34: BP_sw34=Array(BM_sw34, LM_F_l76_sw34, LM_F_l44_sw34, LM_GI_sw34, LM_All_Lights_l65_sw34, LM_All_Lights_l78_sw34)
Dim BP_sw35: BP_sw35=Array(BM_sw35, LM_F_l76_sw35, LM_F_l44_sw35, LM_GI_sw35, LM_All_Lights_l78_sw35, LM_All_Lights_l81_sw35)
Dim BP_sw36: BP_sw36=Array(BM_sw36, LM_F_l76_sw36, LM_F_l44_sw36, LM_GI_sw36, LM_All_Lights_l50_sw36, LM_All_Lights_l94_sw36)
Dim BP_sw37: BP_sw37=Array(BM_sw37, LM_F_l76_sw37, LM_F_l44_sw37, LM_GI_sw37, LM_All_Lights_l50_sw37, LM_All_Lights_l66_sw37, LM_All_Lights_l94_sw37)
Dim BP_sw38: BP_sw38=Array(BM_sw38, LM_F_l76_sw38, LM_F_l44_sw38, LM_GI_sw38, LM_All_Lights_l66_sw38, LM_All_Lights_l82_sw38)
Dim BP_sw39_001: BP_sw39_001=Array(BM_sw39_001, LM_F_l12_sw39_001, LM_GI_sw39_001)
Dim BP_sw40_001: BP_sw40_001=Array(BM_sw40_001, LM_F_l91_sw40_001, LM_GI_sw40_001)
Dim BP_sw41: BP_sw41=Array(BM_sw41, LM_F_l12_sw41, LM_F_l76_sw41, LM_GI_sw41, LM_All_Lights_l3_sw41, LM_All_Lights_l34_sw41, LM_All_Lights_l53_sw41, LM_All_Lights_l67_sw41)
Dim BP_sw42: BP_sw42=Array(BM_sw42, LM_F_l91_sw42, LM_F_l44_sw42, LM_GI_sw42, LM_All_Lights_l42_sw42)
' Arrays per lighting scenario
Dim BL_All_Lights_l1: BL_All_Lights_l1=Array(LM_All_Lights_l1_Parts, LM_All_Lights_l1_UnderPF)
Dim BL_All_Lights_l10: BL_All_Lights_l10=Array(LM_All_Lights_l10_LRSling3, LM_All_Lights_l10_LRSling4, LM_All_Lights_l10_Parts, LM_All_Lights_l10_Plastics, LM_All_Lights_l10_UnderPF)
Dim BL_All_Lights_l11: BL_All_Lights_l11=Array(LM_All_Lights_l11_GateRampY2Wir, LM_All_Lights_l11_GateRampYWire, LM_All_Lights_l11_Parts, LM_All_Lights_l11_Rflip1U, LM_All_Lights_l11_UnderPF)
Dim BL_All_Lights_l12a: BL_All_Lights_l12a=Array(LM_All_Lights_l12a_Parts, LM_All_Lights_l12a_Plastics, LM_All_Lights_l12a_Playfield)
Dim BL_All_Lights_l15: BL_All_Lights_l15=Array(LM_All_Lights_l15_Parts, LM_All_Lights_l15_Plastics, LM_All_Lights_l15_Playfield, LM_All_Lights_l15_ULSling1, LM_All_Lights_l15_ULSling2, LM_All_Lights_l15_ULSling3, LM_All_Lights_l15_ULSling4, LM_All_Lights_l15_pSpinnerRod_0, LM_All_Lights_l15_sw22, LM_All_Lights_l15_sw25, LM_All_Lights_l15_sw26, LM_All_Lights_l15_sw27)
Dim BL_All_Lights_l17: BL_All_Lights_l17=Array(LM_All_Lights_l17_Parts)
Dim BL_All_Lights_l18: BL_All_Lights_l18=Array(LM_All_Lights_l18_Parts, LM_All_Lights_l18_UnderPF)
Dim BL_All_Lights_l19: BL_All_Lights_l19=Array(LM_All_Lights_l19_RflipU)
Dim BL_All_Lights_l2: BL_All_Lights_l2=Array(LM_All_Lights_l2_Parts, LM_All_Lights_l2_UnderPF)
Dim BL_All_Lights_l20: BL_All_Lights_l20=Array(LM_All_Lights_l20_Parts)
Dim BL_All_Lights_l21: BL_All_Lights_l21=Array(LM_All_Lights_l21_Parts, LM_All_Lights_l21_Plastics, LM_All_Lights_l21_Playfield, LM_All_Lights_l21_ULSling, LM_All_Lights_l21_ULSling1, LM_All_Lights_l21_ULSling2, LM_All_Lights_l21_ULSling3, LM_All_Lights_l21_ULSling4, LM_All_Lights_l21_sw25)
Dim BL_All_Lights_l22: BL_All_Lights_l22=Array(LM_All_Lights_l22_LLSling1, LM_All_Lights_l22_LLSling2, LM_All_Lights_l22_LLSling3, LM_All_Lights_l22_LLSling4, LM_All_Lights_l22_Parts)
Dim BL_All_Lights_l23: BL_All_Lights_l23=Array(LM_All_Lights_l23_Parts, LM_All_Lights_l23_UnderPF)
Dim BL_All_Lights_l24: BL_All_Lights_l24=Array(LM_All_Lights_l24_GateRampTWire, LM_All_Lights_l24_Parts, LM_All_Lights_l24_URSling1, LM_All_Lights_l24_URSling2, LM_All_Lights_l24_URSling3, LM_All_Lights_l24_URSling4, LM_All_Lights_l24_UnderPF, LM_All_Lights_l24_pSpinnerRod)
Dim BL_All_Lights_l25: BL_All_Lights_l25=Array(LM_All_Lights_l25_LLSling2, LM_All_Lights_l25_LLSling3, LM_All_Lights_l25_LLSling4, LM_All_Lights_l25_Parts, LM_All_Lights_l25_Plastics, LM_All_Lights_l25_UnderPF)
Dim BL_All_Lights_l26: BL_All_Lights_l26=Array(LM_All_Lights_l26_LRSling1, LM_All_Lights_l26_LRSling2, LM_All_Lights_l26_LRSling3, LM_All_Lights_l26_LRSling4, LM_All_Lights_l26_Parts, LM_All_Lights_l26_Plastics, LM_All_Lights_l26_UnderPF)
Dim BL_All_Lights_l27: BL_All_Lights_l27=Array(LM_All_Lights_l27_Parts, LM_All_Lights_l27_Plastics, LM_All_Lights_l27_Playfield, LM_All_Lights_l27_UnderPF, LM_All_Lights_l27_sw24, LM_All_Lights_l27_sw26, LM_All_Lights_l27_sw27)
Dim BL_All_Lights_l29a: BL_All_Lights_l29a=Array(LM_All_Lights_l29a_Parts, LM_All_Lights_l29a_Plastics, LM_All_Lights_l29a_Playfield, LM_All_Lights_l29a_pSpinnerRod_, LM_All_Lights_l29a_pSpinnerRod_, LM_All_Lights_l29a_sw24)
Dim BL_All_Lights_l3: BL_All_Lights_l3=Array(LM_All_Lights_l3_Lflip1U, LM_All_Lights_l3_Parts, LM_All_Lights_l3_Plastics, LM_All_Lights_l3_UnderPF, LM_All_Lights_l3_sw41)
Dim BL_All_Lights_l31: BL_All_Lights_l31=Array(LM_All_Lights_l31_Parts, LM_All_Lights_l31_Playfield, LM_All_Lights_l31_URSling1, LM_All_Lights_l31_URSling2, LM_All_Lights_l31_URSling3, LM_All_Lights_l31_URSling4, LM_All_Lights_l31_sw23, LM_All_Lights_l31_sw26, LM_All_Lights_l31_sw27)
Dim BL_All_Lights_l33: BL_All_Lights_l33=Array(LM_All_Lights_l33_Parts)
Dim BL_All_Lights_l34: BL_All_Lights_l34=Array(LM_All_Lights_l34_Lflip1U, LM_All_Lights_l34_Parts, LM_All_Lights_l34_Plastics, LM_All_Lights_l34_UnderPF, LM_All_Lights_l34_sw41)
Dim BL_All_Lights_l35: BL_All_Lights_l35=Array(LM_All_Lights_l35_Parts)
Dim BL_All_Lights_l36: BL_All_Lights_l36=Array(LM_All_Lights_l36_Parts)
Dim BL_All_Lights_l37: BL_All_Lights_l37=Array(LM_All_Lights_l37_GateRampIWire, LM_All_Lights_l37_Parts, LM_All_Lights_l37_Plastics, LM_All_Lights_l37_Playfield, LM_All_Lights_l37_ULSling, LM_All_Lights_l37_ULSling1, LM_All_Lights_l37_ULSling2, LM_All_Lights_l37_ULSling3, LM_All_Lights_l37_ULSling4, LM_All_Lights_l37_UnderPF, LM_All_Lights_l37_pSpinnerRod_0)
Dim BL_All_Lights_l38: BL_All_Lights_l38=Array(LM_All_Lights_l38_Parts, LM_All_Lights_l38_UnderPF)
Dim BL_All_Lights_l39: BL_All_Lights_l39=Array(LM_All_Lights_l39_LRSling1, LM_All_Lights_l39_LRSling2, LM_All_Lights_l39_LRSling3, LM_All_Lights_l39_LRSling4, LM_All_Lights_l39_Parts, LM_All_Lights_l39_UnderPF)
Dim BL_All_Lights_l4: BL_All_Lights_l4=Array(LM_All_Lights_l4_Parts, LM_All_Lights_l4_UnderPF)
Dim BL_All_Lights_l40: BL_All_Lights_l40=Array(LM_All_Lights_l40_Parts, LM_All_Lights_l40_Plastics, LM_All_Lights_l40_Playfield, LM_All_Lights_l40_UnderPF, LM_All_Lights_l40_pSpinnerRod_0, LM_All_Lights_l40_sw24, LM_All_Lights_l40_sw27)
Dim BL_All_Lights_l41: BL_All_Lights_l41=Array(LM_All_Lights_l41_Parts, LM_All_Lights_l41_Plastics, LM_All_Lights_l41_UnderPF)
Dim BL_All_Lights_l42: BL_All_Lights_l42=Array(LM_All_Lights_l42_GateRampYWire, LM_All_Lights_l42_Parts, LM_All_Lights_l42_Rflip1U, LM_All_Lights_l42_UnderPF, LM_All_Lights_l42_sw42)
Dim BL_All_Lights_l47: BL_All_Lights_l47=Array(LM_All_Lights_l47_UnderPF)
Dim BL_All_Lights_l49: BL_All_Lights_l49=Array(LM_All_Lights_l49_Parts, LM_All_Lights_l49_Plastics, LM_All_Lights_l49_Playfield, LM_All_Lights_l49_ULSling1, LM_All_Lights_l49_ULSling2, LM_All_Lights_l49_ULSling3, LM_All_Lights_l49_ULSling4, LM_All_Lights_l49_sw33)
Dim BL_All_Lights_l5: BL_All_Lights_l5=Array(LM_All_Lights_l5_GateRampYWire_, LM_All_Lights_l5_Parts, LM_All_Lights_l5_UnderPF)
Dim BL_All_Lights_l50: BL_All_Lights_l50=Array(LM_All_Lights_l50_Parts, LM_All_Lights_l50_Plastics, LM_All_Lights_l50_Playfield, LM_All_Lights_l50_URSling1, LM_All_Lights_l50_URSling2, LM_All_Lights_l50_URSling3, LM_All_Lights_l50_URSling4, LM_All_Lights_l50_sw36, LM_All_Lights_l50_sw37)
Dim BL_All_Lights_l51: BL_All_Lights_l51=Array(LM_All_Lights_l51_LRSling1, LM_All_Lights_l51_LRSling2, LM_All_Lights_l51_LRSling3, LM_All_Lights_l51_LRSling4, LM_All_Lights_l51_Parts, LM_All_Lights_l51_Plastics, LM_All_Lights_l51_Playfield, LM_All_Lights_l51_UnderPF)
Dim BL_All_Lights_l52: BL_All_Lights_l52=Array(LM_All_Lights_l52_Parts, LM_All_Lights_l52_ULSling1, LM_All_Lights_l52_ULSling2, LM_All_Lights_l52_ULSling3, LM_All_Lights_l52_ULSling4, LM_All_Lights_l52_URSling1, LM_All_Lights_l52_URSling2, LM_All_Lights_l52_URSling3, LM_All_Lights_l52_URSling4)
Dim BL_All_Lights_l53: BL_All_Lights_l53=Array(LM_All_Lights_l53_LLSling1, LM_All_Lights_l53_LLSling2, LM_All_Lights_l53_LLSling3, LM_All_Lights_l53_LLSling4, LM_All_Lights_l53_Parts, LM_All_Lights_l53_Playfield, LM_All_Lights_l53_UnderPF, LM_All_Lights_l53_sw41)
Dim BL_All_Lights_l54: BL_All_Lights_l54=Array(LM_All_Lights_l54_LRSling1, LM_All_Lights_l54_LRSling2, LM_All_Lights_l54_LRSling3, LM_All_Lights_l54_LRSling4, LM_All_Lights_l54_Parts, LM_All_Lights_l54_Plastics, LM_All_Lights_l54_UnderPF, LM_All_Lights_l54_sw32_001)
Dim BL_All_Lights_l55: BL_All_Lights_l55=Array(LM_All_Lights_l55_Parts, LM_All_Lights_l55_UnderPF)
Dim BL_All_Lights_l56: BL_All_Lights_l56=Array(LM_All_Lights_l56_LLSling1, LM_All_Lights_l56_LLSling2, LM_All_Lights_l56_LLSling3, LM_All_Lights_l56_LLSling4, LM_All_Lights_l56_Parts, LM_All_Lights_l56_UnderPF)
Dim BL_All_Lights_l61a: BL_All_Lights_l61a=Array(LM_All_Lights_l61a_Parts, LM_All_Lights_l61a_Plastics, LM_All_Lights_l61a_Playfield)
Dim BL_All_Lights_l63: BL_All_Lights_l63=Array(LM_All_Lights_l63_BackPanel, LM_All_Lights_l63_Parts, LM_All_Lights_l63_Playfield, LM_All_Lights_l63_UnderPF)
Dim BL_All_Lights_l65: BL_All_Lights_l65=Array(LM_All_Lights_l65_Parts, LM_All_Lights_l65_Plastics, LM_All_Lights_l65_Playfield, LM_All_Lights_l65_ULSling1, LM_All_Lights_l65_ULSling2, LM_All_Lights_l65_ULSling3, LM_All_Lights_l65_ULSling4, LM_All_Lights_l65_sw34)
Dim BL_All_Lights_l66: BL_All_Lights_l66=Array(LM_All_Lights_l66_Parts, LM_All_Lights_l66_Plastics, LM_All_Lights_l66_Playfield, LM_All_Lights_l66_URSling1, LM_All_Lights_l66_URSling2, LM_All_Lights_l66_URSling3, LM_All_Lights_l66_URSling4, LM_All_Lights_l66_sw37, LM_All_Lights_l66_sw38)
Dim BL_All_Lights_l67: BL_All_Lights_l67=Array(LM_All_Lights_l67_LLSling1, LM_All_Lights_l67_LLSling2, LM_All_Lights_l67_LLSling3, LM_All_Lights_l67_LLSling4, LM_All_Lights_l67_Parts, LM_All_Lights_l67_Plastics, LM_All_Lights_l67_Playfield, LM_All_Lights_l67_UnderPF, LM_All_Lights_l67_sw41)
Dim BL_All_Lights_l68: BL_All_Lights_l68=Array(LM_All_Lights_l68_Parts)
Dim BL_All_Lights_l69: BL_All_Lights_l69=Array(LM_All_Lights_l69_LRSling1, LM_All_Lights_l69_LRSling2, LM_All_Lights_l69_LRSling3, LM_All_Lights_l69_LRSling4, LM_All_Lights_l69_Parts, LM_All_Lights_l69_Plastics, LM_All_Lights_l69_Playfield, LM_All_Lights_l69_UnderPF)
Dim BL_All_Lights_l7: BL_All_Lights_l7=Array(LM_All_Lights_l7_Parts)
Dim BL_All_Lights_l70: BL_All_Lights_l70=Array(LM_All_Lights_l70_LLSling1, LM_All_Lights_l70_LLSling2, LM_All_Lights_l70_LLSling3, LM_All_Lights_l70_LLSling4, LM_All_Lights_l70_Parts, LM_All_Lights_l70_UnderPF)
Dim BL_All_Lights_l71: BL_All_Lights_l71=Array(LM_All_Lights_l71_LRSling1, LM_All_Lights_l71_LRSling2, LM_All_Lights_l71_LRSling3, LM_All_Lights_l71_LRSling4, LM_All_Lights_l71_Parts)
Dim BL_All_Lights_l72: BL_All_Lights_l72=Array(LM_All_Lights_l72_Parts)
Dim BL_All_Lights_l78: BL_All_Lights_l78=Array(LM_All_Lights_l78_Parts, LM_All_Lights_l78_Plastics, LM_All_Lights_l78_Playfield, LM_All_Lights_l78_ULSling, LM_All_Lights_l78_ULSling1, LM_All_Lights_l78_ULSling2, LM_All_Lights_l78_ULSling3, LM_All_Lights_l78_ULSling4, LM_All_Lights_l78_pSpinnerRod_0, LM_All_Lights_l78_sw34, LM_All_Lights_l78_sw35)
Dim BL_All_Lights_l78b: BL_All_Lights_l78b=Array(LM_All_Lights_l78b_GateRampC2Wi, LM_All_Lights_l78b_Parts, LM_All_Lights_l78b_Plastics, LM_All_Lights_l78b_Playfield, LM_All_Lights_l78b_ULSling, LM_All_Lights_l78b_ULSling1, LM_All_Lights_l78b_ULSling2, LM_All_Lights_l78b_ULSling3, LM_All_Lights_l78b_ULSling4, LM_All_Lights_l78b_pSpinnerRod_)
Dim BL_All_Lights_l79: BL_All_Lights_l79=Array(LM_All_Lights_l79_SideInserts)
Dim BL_All_Lights_l8: BL_All_Lights_l8=Array(LM_All_Lights_l8_Parts, LM_All_Lights_l8_Plastics, LM_All_Lights_l8_Playfield, LM_All_Lights_l8_URSling, LM_All_Lights_l8_URSling1, LM_All_Lights_l8_URSling2, LM_All_Lights_l8_URSling3, LM_All_Lights_l8_URSling4, LM_All_Lights_l8_sw25, LM_All_Lights_l8_sw26)
Dim BL_All_Lights_l81: BL_All_Lights_l81=Array(LM_All_Lights_l81_Parts, LM_All_Lights_l81_Plastics, LM_All_Lights_l81_Playfield, LM_All_Lights_l81_ULSling1, LM_All_Lights_l81_ULSling2, LM_All_Lights_l81_ULSling3, LM_All_Lights_l81_ULSling4, LM_All_Lights_l81_sw35)
Dim BL_All_Lights_l82: BL_All_Lights_l82=Array(LM_All_Lights_l82_Parts, LM_All_Lights_l82_Plastics, LM_All_Lights_l82_Playfield, LM_All_Lights_l82_URSling1, LM_All_Lights_l82_URSling2, LM_All_Lights_l82_URSling3, LM_All_Lights_l82_URSling4, LM_All_Lights_l82_sw38)
Dim BL_All_Lights_l83: BL_All_Lights_l83=Array(LM_All_Lights_l83_Parts)
Dim BL_All_Lights_l84: BL_All_Lights_l84=Array(LM_All_Lights_l84_Parts, LM_All_Lights_l84_Plastics, LM_All_Lights_l84_ULSling1, LM_All_Lights_l84_ULSling2, LM_All_Lights_l84_ULSling3, LM_All_Lights_l84_ULSling4, LM_All_Lights_l84_URSling1, LM_All_Lights_l84_URSling2, LM_All_Lights_l84_URSling3, LM_All_Lights_l84_URSling4, LM_All_Lights_l84_UnderPF)
Dim BL_All_Lights_l85: BL_All_Lights_l85=Array(LM_All_Lights_l85_LLSling1, LM_All_Lights_l85_LLSling2, LM_All_Lights_l85_LLSling3, LM_All_Lights_l85_LLSling4, LM_All_Lights_l85_Parts, LM_All_Lights_l85_Playfield, LM_All_Lights_l85_UnderPF, LM_All_Lights_l85_sw31_001)
Dim BL_All_Lights_l86: BL_All_Lights_l86=Array(LM_All_Lights_l86_LLSling3, LM_All_Lights_l86_LLSling4, LM_All_Lights_l86_Parts, LM_All_Lights_l86_UnderPF)
Dim BL_All_Lights_l87: BL_All_Lights_l87=Array(LM_All_Lights_l87_LRSling1, LM_All_Lights_l87_LRSling2, LM_All_Lights_l87_LRSling3, LM_All_Lights_l87_LRSling4, LM_All_Lights_l87_Parts, LM_All_Lights_l87_UnderPF)
Dim BL_All_Lights_l88: BL_All_Lights_l88=Array(LM_All_Lights_l88_Parts, LM_All_Lights_l88_UnderPF)
Dim BL_All_Lights_l9: BL_All_Lights_l9=Array(LM_All_Lights_l9_LLSling1, LM_All_Lights_l9_LLSling2, LM_All_Lights_l9_LLSling3, LM_All_Lights_l9_LLSling4, LM_All_Lights_l9_Parts, LM_All_Lights_l9_Plastics, LM_All_Lights_l9_UnderPF)
Dim BL_All_Lights_l90: BL_All_Lights_l90=Array(LM_All_Lights_l90_Parts, LM_All_Lights_l90_Plastics, LM_All_Lights_l90_Playfield)
Dim BL_All_Lights_l91a: BL_All_Lights_l91a=Array(LM_All_Lights_l91a_Parts, LM_All_Lights_l91a_Plastics, LM_All_Lights_l91a_Playfield)
Dim BL_All_Lights_l94: BL_All_Lights_l94=Array(LM_All_Lights_l94_Parts, LM_All_Lights_l94_Plastics, LM_All_Lights_l94_Playfield, LM_All_Lights_l94_URSling, LM_All_Lights_l94_URSling1, LM_All_Lights_l94_URSling2, LM_All_Lights_l94_URSling3, LM_All_Lights_l94_URSling4, LM_All_Lights_l94_pSpinnerRod_0, LM_All_Lights_l94_sw36, LM_All_Lights_l94_sw37)
Dim BL_All_Lights_l94b: BL_All_Lights_l94b=Array(LM_All_Lights_l94b_GateRampY2Wi, LM_All_Lights_l94b_Parts, LM_All_Lights_l94b_Plastics, LM_All_Lights_l94b_Playfield, LM_All_Lights_l94b_URSling, LM_All_Lights_l94b_URSling1, LM_All_Lights_l94b_URSling2, LM_All_Lights_l94b_URSling3, LM_All_Lights_l94b_URSling4, LM_All_Lights_l94b_pSpinnerRod, LM_All_Lights_l94b_pSpinnerRod_)
Dim BL_F_l12: BL_F_l12=Array(LM_F_l12_LLSling1, LM_F_l12_LLSling2, LM_F_l12_LLSling3, LM_F_l12_LLSling4, LM_F_l12_LRSling1, LM_F_l12_LRSling2, LM_F_l12_LRSling3, LM_F_l12_LRSling4, LM_F_l12_Lflip, LM_F_l12_LflipU, LM_F_l12_NCoverL, LM_F_l12_Parts, LM_F_l12_Plastics, LM_F_l12_Playfield, LM_F_l12_Rflip, LM_F_l12_RflipU, LM_F_l12_Sling1, LM_F_l12_Sling2, LM_F_l12_UnderPF, LM_F_l12_sw31_001, LM_F_l12_sw39_001, LM_F_l12_sw41)
Dim BL_F_l29: BL_F_l29=Array(LM_F_l29_BackPanel, LM_F_l29_GateRampIWire_002, LM_F_l29_LaneSteeringWall3, LM_F_l29_LaneSwitch1, LM_F_l29_LaneSwitch2, LM_F_l29_LaneSwitch3, LM_F_l29_LaneSwitch4, LM_F_l29_Parts, LM_F_l29_Plastics, LM_F_l29_Playfield, LM_F_l29_SideInserts, LM_F_l29_ULSling1, LM_F_l29_ULSling2, LM_F_l29_ULSling3, LM_F_l29_ULSling4, LM_F_l29_URSling, LM_F_l29_URSling1, LM_F_l29_URSling2, LM_F_l29_URSling3, LM_F_l29_URSling4, LM_F_l29_UnderPF, LM_F_l29_pSpinnerRod, LM_F_l29_pSpinnerRod_001, LM_F_l29_pSpinnerRod_004, LM_F_l29_sw21, LM_F_l29_sw22, LM_F_l29_sw23, LM_F_l29_sw24, LM_F_l29_sw25, LM_F_l29_sw26, LM_F_l29_sw27)
Dim BL_F_l43: BL_F_l43=Array(LM_F_l43_BackPanel, LM_F_l43_LaneSwitch4, LM_F_l43_Layer_1, LM_F_l43_Parts, LM_F_l43_Plastics)
Dim BL_F_l44: BL_F_l44=Array(LM_F_l44_GateRampCWire_001, LM_F_l44_GateRampIWire_002, LM_F_l44_GateRampTWire_002, LM_F_l44_GateRampY2Wire_002, LM_F_l44_GateRampYWire_001, LM_F_l44_Lflip1, LM_F_l44_Lflip1U, LM_F_l44_Parts, LM_F_l44_Plastics, LM_F_l44_Playfield, LM_F_l44_Rflip1, LM_F_l44_Rflip1U, LM_F_l44_ULSling, LM_F_l44_ULSling1, LM_F_l44_ULSling2, LM_F_l44_ULSling3, LM_F_l44_ULSling4, LM_F_l44_URSling, LM_F_l44_URSling1, LM_F_l44_URSling2, LM_F_l44_URSling3, LM_F_l44_URSling4, LM_F_l44_UnderPF, LM_F_l44_pSpinnerRod, LM_F_l44_pSpinnerRod_001, LM_F_l44_pSpinnerRod_002, LM_F_l44_pSpinnerRod_004, LM_F_l44_sw21, LM_F_l44_sw22, LM_F_l44_sw23, LM_F_l44_sw25, LM_F_l44_sw26, LM_F_l44_sw27, LM_F_l44_sw33, LM_F_l44_sw34, LM_F_l44_sw35, LM_F_l44_sw36, LM_F_l44_sw37, LM_F_l44_sw38, LM_F_l44_sw42)
Dim BL_F_l45: BL_F_l45=Array(LM_F_l45_Parts, LM_F_l45_UnderPF)
Dim BL_F_l57: BL_F_l57=Array(LM_F_l57_BackPanel, LM_F_l57_Layer_1, LM_F_l57_Parts, LM_F_l57_SideInserts)
Dim BL_F_l58: BL_F_l58=Array(LM_F_l58_BackPanel, LM_F_l58_LaneSwitch1, LM_F_l58_Layer_1, LM_F_l58_Parts, LM_F_l58_SideInserts)
Dim BL_F_l59: BL_F_l59=Array(LM_F_l59_BackPanel, LM_F_l59_Layer_1, LM_F_l59_Parts, LM_F_l59_Plastics)
Dim BL_F_l60: BL_F_l60=Array(LM_F_l60_BackPanel, LM_F_l60_Layer_1, LM_F_l60_Parts, LM_F_l60_SideInserts)
Dim BL_F_l61: BL_F_l61=Array(LM_F_l61_BackPanel, LM_F_l61_GateRampIWire_002, LM_F_l61_GateRampTWire_002, LM_F_l61_LaneSteeringWall3, LM_F_l61_LaneSwitch1, LM_F_l61_LaneSwitch2, LM_F_l61_LaneSwitch3, LM_F_l61_LaneSwitch4, LM_F_l61_Layer_1, LM_F_l61_Parts, LM_F_l61_Plastics, LM_F_l61_Playfield, LM_F_l61_SideInserts, LM_F_l61_ULSling, LM_F_l61_ULSling1, LM_F_l61_ULSling2, LM_F_l61_ULSling3, LM_F_l61_ULSling4, LM_F_l61_URSling, LM_F_l61_URSling1, LM_F_l61_URSling2, LM_F_l61_URSling3, LM_F_l61_URSling4, LM_F_l61_UnderPF, LM_F_l61_pSpinnerRod, LM_F_l61_pSpinnerRod_001, LM_F_l61_pSpinnerRod_003, LM_F_l61_pSpinnerRod_004, LM_F_l61_sw21, LM_F_l61_sw22, LM_F_l61_sw23, LM_F_l61_sw24, LM_F_l61_sw25, LM_F_l61_sw26, LM_F_l61_sw27)
Dim BL_F_l73: BL_F_l73=Array(LM_F_l73_BackPanel, LM_F_l73_Layer_1, LM_F_l73_Parts, LM_F_l73_SideInserts)
Dim BL_F_l75: BL_F_l75=Array(LM_F_l75_BackPanel, LM_F_l75_Layer_1, LM_F_l75_Parts, LM_F_l75_Plastics)
Dim BL_F_l76: BL_F_l76=Array(LM_F_l76_GateRampC2Wire_002, LM_F_l76_GateRampCWire_001, LM_F_l76_GateRampIWire_002, LM_F_l76_GateRampTWire_002, LM_F_l76_GateRampYWire_001, LM_F_l76_Lflip1, LM_F_l76_Lflip1U, LM_F_l76_Parts, LM_F_l76_Plastics, LM_F_l76_Playfield, LM_F_l76_Rflip1, LM_F_l76_Rflip1U, LM_F_l76_ULSling, LM_F_l76_ULSling1, LM_F_l76_ULSling2, LM_F_l76_ULSling3, LM_F_l76_ULSling4, LM_F_l76_URSling, LM_F_l76_URSling1, LM_F_l76_URSling2, LM_F_l76_URSling3, LM_F_l76_URSling4, LM_F_l76_UnderPF, LM_F_l76_pSpinnerRod, LM_F_l76_pSpinnerRod_001, LM_F_l76_pSpinnerRod_003, LM_F_l76_pSpinnerRod_004, LM_F_l76_sw21, LM_F_l76_sw22, LM_F_l76_sw23, LM_F_l76_sw25, LM_F_l76_sw26, LM_F_l76_sw27, LM_F_l76_sw33, LM_F_l76_sw34, LM_F_l76_sw35, LM_F_l76_sw36, LM_F_l76_sw37, LM_F_l76_sw38, LM_F_l76_sw41)
Dim BL_F_l79: BL_F_l79=Array(LM_F_l79_BackPanel, LM_F_l79_Layer_1, LM_F_l79_Parts, LM_F_l79_SideInserts)
Dim BL_F_l89: BL_F_l89=Array(LM_F_l89_BackPanel, LM_F_l89_Layer_1, LM_F_l89_Parts, LM_F_l89_SideInserts)
Dim BL_F_l91: BL_F_l91=Array(LM_F_l91_BackPanel, LM_F_l91_LLSling1, LM_F_l91_LLSling2, LM_F_l91_LLSling3, LM_F_l91_LLSling4, LM_F_l91_LRSling1, LM_F_l91_LRSling2, LM_F_l91_LRSling3, LM_F_l91_LRSling4, LM_F_l91_Lflip, LM_F_l91_LflipU, LM_F_l91_NCoverL, LM_F_l91_NCoverR, LM_F_l91_Parts, LM_F_l91_Plastics, LM_F_l91_Playfield, LM_F_l91_Rflip, LM_F_l91_RflipU, LM_F_l91_Sling1, LM_F_l91_Sling2, LM_F_l91_UnderPF, LM_F_l91_sw31_001, LM_F_l91_sw32_001, LM_F_l91_sw40_001, LM_F_l91_sw42)
Dim BL_F_l92: BL_F_l92=Array(LM_F_l92_BackPanel, LM_F_l92_Layer_1, LM_F_l92_Parts, LM_F_l92_SideInserts)
Dim BL_F_l95: BL_F_l95=Array(LM_F_l95_BackPanel, LM_F_l95_Layer_1, LM_F_l95_SideInserts)
Dim BL_GI: BL_GI=Array(LM_GI_BackPanel, LM_GI_GateRampC2Wire_002, LM_GI_GateRampCWire_001, LM_GI_GateRampIWire_002, LM_GI_GateRampTWire_002, LM_GI_GateRampY2Wire_002, LM_GI_GateRampYWire_001, LM_GI_LLSling1, LM_GI_LLSling2, LM_GI_LLSling3, LM_GI_LLSling4, LM_GI_LRSling1, LM_GI_LRSling2, LM_GI_LRSling3, LM_GI_LRSling4, LM_GI_LaneSteeringWall3, LM_GI_LaneSwitch1, LM_GI_LaneSwitch2, LM_GI_LaneSwitch3, LM_GI_LaneSwitch4, LM_GI_Layer_1, LM_GI_Lflip, LM_GI_Lflip1, LM_GI_Lflip1U, LM_GI_LflipU, LM_GI_MiddlePost, LM_GI_NCoverL, LM_GI_NCoverR, LM_GI_Parts, LM_GI_Plastics, LM_GI_Playfield, LM_GI_Playfield_001, LM_GI_Rflip, LM_GI_Rflip1, LM_GI_Rflip1U, LM_GI_RflipU, LM_GI_SideInserts, LM_GI_Sling1, LM_GI_Sling2, LM_GI_ULSling, LM_GI_ULSling1, LM_GI_ULSling2, LM_GI_ULSling3, LM_GI_ULSling4, LM_GI_URSling, LM_GI_URSling1, LM_GI_URSling2, LM_GI_URSling3, LM_GI_URSling4, LM_GI_UnderPF, LM_GI_pSpinnerRod, LM_GI_pSpinnerRod_001, LM_GI_pSpinnerRod_002, LM_GI_pSpinnerRod_003, LM_GI_pSpinnerRod_004, LM_GI_sw21, LM_GI_sw22, _
  LM_GI_sw23, LM_GI_sw24, LM_GI_sw25, LM_GI_sw26, LM_GI_sw27, LM_GI_sw31_001, LM_GI_sw32_001, LM_GI_sw33, LM_GI_sw34, LM_GI_sw35, LM_GI_sw36, LM_GI_sw37, LM_GI_sw38, LM_GI_sw39_001, LM_GI_sw40_001, LM_GI_sw41, LM_GI_sw42)
Dim BL_L_l1: BL_L_l1=Array(LM_L_l1_Parts, LM_L_l1_Playfield, LM_L_l1_UnderPF)
Dim BL_L_l10: BL_L_l10=Array(LM_L_l10_Playfield, LM_L_l10_UnderPF)
Dim BL_L_l11: BL_L_l11=Array(LM_L_l11_GateRampYWire_001, LM_L_l11_Parts, LM_L_l11_Playfield, LM_L_l11_Rflip1U, LM_L_l11_UnderPF)
Dim BL_L_l15: BL_L_l15=Array(LM_L_l15_Parts, LM_L_l15_UnderPF, LM_L_l15_pSpinnerRod_001)
Dim BL_L_l17: BL_L_l17=Array(LM_L_l17_Playfield, LM_L_l17_UnderPF)
Dim BL_L_l18: BL_L_l18=Array(LM_L_l18_Playfield, LM_L_l18_UnderPF)
Dim BL_L_l19: BL_L_l19=Array(LM_L_l19_UnderPF)
Dim BL_L_l2: BL_L_l2=Array(LM_L_l2_Playfield, LM_L_l2_UnderPF)
Dim BL_L_l20: BL_L_l20=Array(LM_L_l20_Playfield, LM_L_l20_UnderPF)
Dim BL_L_l21: BL_L_l21=Array(LM_L_l21_Parts, LM_L_l21_Plastics, LM_L_l21_Playfield, LM_L_l21_UnderPF)
Dim BL_L_l22: BL_L_l22=Array(LM_L_l22_Parts, LM_L_l22_Playfield, LM_L_l22_UnderPF)
Dim BL_L_l23: BL_L_l23=Array(LM_L_l23_Playfield, LM_L_l23_UnderPF)
Dim BL_L_l24: BL_L_l24=Array(LM_L_l24_Parts, LM_L_l24_URSling2, LM_L_l24_URSling3, LM_L_l24_URSling4, LM_L_l24_UnderPF)
Dim BL_L_l25: BL_L_l25=Array(LM_L_l25_Playfield, LM_L_l25_UnderPF)
Dim BL_L_l26: BL_L_l26=Array(LM_L_l26_Parts, LM_L_l26_Playfield, LM_L_l26_UnderPF)
Dim BL_L_l27: BL_L_l27=Array(LM_L_l27_Parts, LM_L_l27_Plastics, LM_L_l27_Playfield, LM_L_l27_UnderPF, LM_L_l27_sw24)
Dim BL_L_l3: BL_L_l3=Array(LM_L_l3_Lflip1U, LM_L_l3_Parts, LM_L_l3_Playfield, LM_L_l3_UnderPF)
Dim BL_L_l31: BL_L_l31=Array(LM_L_l31_Parts, LM_L_l31_UnderPF)
Dim BL_L_l33: BL_L_l33=Array(LM_L_l33_Playfield, LM_L_l33_UnderPF)
Dim BL_L_l34: BL_L_l34=Array(LM_L_l34_Parts, LM_L_l34_Playfield, LM_L_l34_UnderPF)
Dim BL_L_l35: BL_L_l35=Array(LM_L_l35_Playfield, LM_L_l35_UnderPF)
Dim BL_L_l36: BL_L_l36=Array(LM_L_l36_Playfield, LM_L_l36_UnderPF)
Dim BL_L_l37: BL_L_l37=Array(LM_L_l37_Parts, LM_L_l37_Playfield, LM_L_l37_ULSling2, LM_L_l37_ULSling3, LM_L_l37_ULSling4, LM_L_l37_UnderPF)
Dim BL_L_l38: BL_L_l38=Array(LM_L_l38_Playfield, LM_L_l38_UnderPF)
Dim BL_L_l39: BL_L_l39=Array(LM_L_l39_Playfield, LM_L_l39_UnderPF)
Dim BL_L_l4: BL_L_l4=Array(LM_L_l4_Playfield, LM_L_l4_UnderPF)
Dim BL_L_l40: BL_L_l40=Array(LM_L_l40_Parts, LM_L_l40_Plastics, LM_L_l40_Playfield, LM_L_l40_UnderPF, LM_L_l40_sw24)
Dim BL_L_l41: BL_L_l41=Array(LM_L_l41_Playfield, LM_L_l41_UnderPF)
Dim BL_L_l42: BL_L_l42=Array(LM_L_l42_Parts, LM_L_l42_Playfield, LM_L_l42_UnderPF)
Dim BL_L_l47: BL_L_l47=Array(LM_L_l47_UnderPF)
Dim BL_L_l49: BL_L_l49=Array(LM_L_l49_Parts, LM_L_l49_Playfield, LM_L_l49_UnderPF)
Dim BL_L_l5: BL_L_l5=Array(LM_L_l5_Parts, LM_L_l5_Playfield, LM_L_l5_UnderPF)
Dim BL_L_l50: BL_L_l50=Array(LM_L_l50_Parts, LM_L_l50_Playfield, LM_L_l50_UnderPF)
Dim BL_L_l51: BL_L_l51=Array(LM_L_l51_LRSling1, LM_L_l51_LRSling2, LM_L_l51_LRSling3, LM_L_l51_LRSling4, LM_L_l51_Parts, LM_L_l51_Plastics, LM_L_l51_UnderPF)
Dim BL_L_l52: BL_L_l52=Array(LM_L_l52_Parts, LM_L_l52_UnderPF)
Dim BL_L_l53: BL_L_l53=Array(LM_L_l53_Parts, LM_L_l53_Playfield, LM_L_l53_UnderPF)
Dim BL_L_l54: BL_L_l54=Array(LM_L_l54_Parts, LM_L_l54_Playfield, LM_L_l54_UnderPF)
Dim BL_L_l55: BL_L_l55=Array(LM_L_l55_UnderPF)
Dim BL_L_l56: BL_L_l56=Array(LM_L_l56_UnderPF)
Dim BL_L_l63: BL_L_l63=Array(LM_L_l63_Parts, LM_L_l63_UnderPF)
Dim BL_L_l65: BL_L_l65=Array(LM_L_l65_Parts, LM_L_l65_Playfield, LM_L_l65_UnderPF)
Dim BL_L_l66: BL_L_l66=Array(LM_L_l66_Parts, LM_L_l66_Playfield, LM_L_l66_UnderPF)
Dim BL_L_l67: BL_L_l67=Array(LM_L_l67_LLSling1, LM_L_l67_LLSling2, LM_L_l67_LLSling3, LM_L_l67_LLSling4, LM_L_l67_Parts, LM_L_l67_Plastics, LM_L_l67_Playfield, LM_L_l67_UnderPF)
Dim BL_L_l68: BL_L_l68=Array(LM_L_l68_Parts, LM_L_l68_UnderPF)
Dim BL_L_l69: BL_L_l69=Array(LM_L_l69_Parts, LM_L_l69_Playfield, LM_L_l69_UnderPF)
Dim BL_L_l7: BL_L_l7=Array(LM_L_l7_Playfield, LM_L_l7_UnderPF)
Dim BL_L_l70: BL_L_l70=Array(LM_L_l70_Parts, LM_L_l70_UnderPF)
Dim BL_L_l71: BL_L_l71=Array(LM_L_l71_UnderPF)
Dim BL_L_l72: BL_L_l72=Array(LM_L_l72_UnderPF)
Dim BL_L_l8: BL_L_l8=Array(LM_L_l8_Parts, LM_L_l8_Playfield, LM_L_l8_UnderPF)
Dim BL_L_l81: BL_L_l81=Array(LM_L_l81_Playfield, LM_L_l81_UnderPF)
Dim BL_L_l82: BL_L_l82=Array(LM_L_l82_Parts, LM_L_l82_Playfield, LM_L_l82_UnderPF)
Dim BL_L_l83: BL_L_l83=Array(LM_L_l83_Parts, LM_L_l83_UnderPF)
Dim BL_L_l84: BL_L_l84=Array(LM_L_l84_Parts, LM_L_l84_Playfield, LM_L_l84_UnderPF)
Dim BL_L_l85: BL_L_l85=Array(LM_L_l85_Parts, LM_L_l85_Playfield, LM_L_l85_UnderPF)
Dim BL_L_l86: BL_L_l86=Array(LM_L_l86_UnderPF)
Dim BL_L_l87: BL_L_l87=Array(LM_L_l87_Parts, LM_L_l87_Playfield, LM_L_l87_UnderPF)
Dim BL_L_l88: BL_L_l88=Array(LM_L_l88_UnderPF)
Dim BL_L_l9: BL_L_l9=Array(LM_L_l9_Parts, LM_L_l9_Playfield, LM_L_l9_UnderPF)
Dim BL_L_l90: BL_L_l90=Array(LM_L_l90_Parts, LM_L_l90_Plastics, LM_L_l90_UnderPF)
Dim BL_World: BL_World=Array(BM_BackPanel, BM_GateRampC2Wire_002, BM_GateRampCWire_001, BM_GateRampIWire_002, BM_GateRampTWire_002, BM_GateRampY2Wire_002, BM_GateRampYWire_001, BM_LLSling1, BM_LLSling2, BM_LLSling3, BM_LLSling4, BM_LRSling1, BM_LRSling2, BM_LRSling3, BM_LRSling4, BM_LaneSteeringWall3, BM_LaneSwitch1, BM_LaneSwitch2, BM_LaneSwitch3, BM_LaneSwitch4, BM_Layer_1, BM_Lflip, BM_Lflip1, BM_Lflip1U, BM_LflipU, BM_MiddlePost, BM_NCoverL, BM_NCoverR, BM_Parts, BM_Parts_001, BM_Plastics, BM_Playfield, BM_Playfield_001, BM_Rflip, BM_Rflip1, BM_Rflip1U, BM_RflipU, BM_SideInserts, BM_Sling1, BM_Sling2, BM_ULSling, BM_ULSling1, BM_ULSling2, BM_ULSling3, BM_ULSling4, BM_URSling, BM_URSling1, BM_URSling2, BM_URSling3, BM_URSling4, BM_UnderPF, BM_pSpinnerRod, BM_pSpinnerRod_001, BM_pSpinnerRod_002, BM_pSpinnerRod_003, BM_pSpinnerRod_004, BM_sw16, BM_sw21, BM_sw22, BM_sw23, BM_sw24, BM_sw25, BM_sw26, BM_sw27, BM_sw31_001, BM_sw32_001, BM_sw33, BM_sw34, BM_sw35, BM_sw36, BM_sw37, BM_sw38, BM_sw39_001, _
  BM_sw40_001, BM_sw41, BM_sw42)
' Global arrays
Dim BG_Bakemap: BG_Bakemap=Array(BM_BackPanel, BM_GateRampC2Wire_002, BM_GateRampCWire_001, BM_GateRampIWire_002, BM_GateRampTWire_002, BM_GateRampY2Wire_002, BM_GateRampYWire_001, BM_LLSling1, BM_LLSling2, BM_LLSling3, BM_LLSling4, BM_LRSling1, BM_LRSling2, BM_LRSling3, BM_LRSling4, BM_LaneSteeringWall3, BM_LaneSwitch1, BM_LaneSwitch2, BM_LaneSwitch3, BM_LaneSwitch4, BM_Layer_1, BM_Lflip, BM_Lflip1, BM_Lflip1U, BM_LflipU, BM_MiddlePost, BM_NCoverL, BM_NCoverR, BM_Parts, BM_Parts_001, BM_Plastics, BM_Playfield, BM_Playfield_001, BM_Rflip, BM_Rflip1, BM_Rflip1U, BM_RflipU, BM_SideInserts, BM_Sling1, BM_Sling2, BM_ULSling, BM_ULSling1, BM_ULSling2, BM_ULSling3, BM_ULSling4, BM_URSling, BM_URSling1, BM_URSling2, BM_URSling3, BM_URSling4, BM_UnderPF, BM_pSpinnerRod, BM_pSpinnerRod_001, BM_pSpinnerRod_002, BM_pSpinnerRod_003, BM_pSpinnerRod_004, BM_sw16, BM_sw21, BM_sw22, BM_sw23, BM_sw24, BM_sw25, BM_sw26, BM_sw27, BM_sw31_001, BM_sw32_001, BM_sw33, BM_sw34, BM_sw35, BM_sw36, BM_sw37, BM_sw38, BM_sw39_001, _
  BM_sw40_001, BM_sw41, BM_sw42)
Dim BG_Lightmap: BG_Lightmap=Array(LM_All_Lights_l1_Parts, LM_All_Lights_l1_UnderPF, LM_All_Lights_l10_LRSling3, LM_All_Lights_l10_LRSling4, LM_All_Lights_l10_Parts, LM_All_Lights_l10_Plastics, LM_All_Lights_l10_UnderPF, LM_All_Lights_l11_GateRampY2Wir, LM_All_Lights_l11_GateRampYWire, LM_All_Lights_l11_Parts, LM_All_Lights_l11_Rflip1U, LM_All_Lights_l11_UnderPF, LM_All_Lights_l12a_Parts, LM_All_Lights_l12a_Plastics, LM_All_Lights_l12a_Playfield, LM_All_Lights_l15_Parts, LM_All_Lights_l15_Plastics, LM_All_Lights_l15_Playfield, LM_All_Lights_l15_ULSling1, LM_All_Lights_l15_ULSling2, LM_All_Lights_l15_ULSling3, LM_All_Lights_l15_ULSling4, LM_All_Lights_l15_pSpinnerRod_0, LM_All_Lights_l15_sw22, LM_All_Lights_l15_sw25, LM_All_Lights_l15_sw26, LM_All_Lights_l15_sw27, LM_All_Lights_l17_Parts, LM_All_Lights_l18_Parts, LM_All_Lights_l18_UnderPF, LM_All_Lights_l19_RflipU, LM_All_Lights_l2_Parts, LM_All_Lights_l2_UnderPF, LM_All_Lights_l20_Parts, LM_All_Lights_l21_Parts, LM_All_Lights_l21_Plastics, _
  LM_All_Lights_l21_Playfield, LM_All_Lights_l21_ULSling, LM_All_Lights_l21_ULSling1, LM_All_Lights_l21_ULSling2, LM_All_Lights_l21_ULSling3, LM_All_Lights_l21_ULSling4, LM_All_Lights_l21_sw25, LM_All_Lights_l22_LLSling1, LM_All_Lights_l22_LLSling2, LM_All_Lights_l22_LLSling3, LM_All_Lights_l22_LLSling4, LM_All_Lights_l22_Parts, LM_All_Lights_l23_Parts, LM_All_Lights_l23_UnderPF, LM_All_Lights_l24_GateRampTWire, LM_All_Lights_l24_Parts, LM_All_Lights_l24_URSling1, LM_All_Lights_l24_URSling2, LM_All_Lights_l24_URSling3, LM_All_Lights_l24_URSling4, LM_All_Lights_l24_UnderPF, LM_All_Lights_l24_pSpinnerRod, LM_All_Lights_l25_LLSling2, LM_All_Lights_l25_LLSling3, LM_All_Lights_l25_LLSling4, LM_All_Lights_l25_Parts, LM_All_Lights_l25_Plastics, LM_All_Lights_l25_UnderPF, LM_All_Lights_l26_LRSling1, LM_All_Lights_l26_LRSling2, LM_All_Lights_l26_LRSling3, LM_All_Lights_l26_LRSling4, LM_All_Lights_l26_Parts, LM_All_Lights_l26_Plastics, LM_All_Lights_l26_UnderPF, LM_All_Lights_l27_Parts, LM_All_Lights_l27_Plastics, _
  LM_All_Lights_l27_Playfield, LM_All_Lights_l27_UnderPF, LM_All_Lights_l27_sw24, LM_All_Lights_l27_sw26, LM_All_Lights_l27_sw27, LM_All_Lights_l29a_Parts, LM_All_Lights_l29a_Plastics, LM_All_Lights_l29a_Playfield, LM_All_Lights_l29a_pSpinnerRod_, LM_All_Lights_l29a_pSpinnerRod_, LM_All_Lights_l29a_sw24, LM_All_Lights_l3_Lflip1U, LM_All_Lights_l3_Parts, LM_All_Lights_l3_Plastics, LM_All_Lights_l3_UnderPF, LM_All_Lights_l3_sw41, LM_All_Lights_l31_Parts, LM_All_Lights_l31_Playfield, LM_All_Lights_l31_URSling1, LM_All_Lights_l31_URSling2, LM_All_Lights_l31_URSling3, LM_All_Lights_l31_URSling4, LM_All_Lights_l31_sw23, LM_All_Lights_l31_sw26, LM_All_Lights_l31_sw27, LM_All_Lights_l33_Parts, LM_All_Lights_l34_Lflip1U, LM_All_Lights_l34_Parts, LM_All_Lights_l34_Plastics, LM_All_Lights_l34_UnderPF, LM_All_Lights_l34_sw41, LM_All_Lights_l35_Parts, LM_All_Lights_l36_Parts, LM_All_Lights_l37_GateRampIWire, LM_All_Lights_l37_Parts, LM_All_Lights_l37_Plastics, LM_All_Lights_l37_Playfield, LM_All_Lights_l37_ULSling, _
  LM_All_Lights_l37_ULSling1, LM_All_Lights_l37_ULSling2, LM_All_Lights_l37_ULSling3, LM_All_Lights_l37_ULSling4, LM_All_Lights_l37_UnderPF, LM_All_Lights_l37_pSpinnerRod_0, LM_All_Lights_l38_Parts, LM_All_Lights_l38_UnderPF, LM_All_Lights_l39_LRSling1, LM_All_Lights_l39_LRSling2, LM_All_Lights_l39_LRSling3, LM_All_Lights_l39_LRSling4, LM_All_Lights_l39_Parts, LM_All_Lights_l39_UnderPF, LM_All_Lights_l4_Parts, LM_All_Lights_l4_UnderPF, LM_All_Lights_l40_Parts, LM_All_Lights_l40_Plastics, LM_All_Lights_l40_Playfield, LM_All_Lights_l40_UnderPF, LM_All_Lights_l40_pSpinnerRod_0, LM_All_Lights_l40_sw24, LM_All_Lights_l40_sw27, LM_All_Lights_l41_Parts, LM_All_Lights_l41_Plastics, LM_All_Lights_l41_UnderPF, LM_All_Lights_l42_GateRampYWire, LM_All_Lights_l42_Parts, LM_All_Lights_l42_Rflip1U, LM_All_Lights_l42_UnderPF, LM_All_Lights_l42_sw42, LM_All_Lights_l47_UnderPF, LM_All_Lights_l49_Parts, LM_All_Lights_l49_Plastics, LM_All_Lights_l49_Playfield, LM_All_Lights_l49_ULSling1, LM_All_Lights_l49_ULSling2, _
  LM_All_Lights_l49_ULSling3, LM_All_Lights_l49_ULSling4, LM_All_Lights_l49_sw33, LM_All_Lights_l5_GateRampYWire_, LM_All_Lights_l5_Parts, LM_All_Lights_l5_UnderPF, LM_All_Lights_l50_Parts, LM_All_Lights_l50_Plastics, LM_All_Lights_l50_Playfield, LM_All_Lights_l50_URSling1, LM_All_Lights_l50_URSling2, LM_All_Lights_l50_URSling3, LM_All_Lights_l50_URSling4, LM_All_Lights_l50_sw36, LM_All_Lights_l50_sw37, LM_All_Lights_l51_LRSling1, LM_All_Lights_l51_LRSling2, LM_All_Lights_l51_LRSling3, LM_All_Lights_l51_LRSling4, LM_All_Lights_l51_Parts, LM_All_Lights_l51_Plastics, LM_All_Lights_l51_Playfield, LM_All_Lights_l51_UnderPF, LM_All_Lights_l52_Parts, LM_All_Lights_l52_ULSling1, LM_All_Lights_l52_ULSling2, LM_All_Lights_l52_ULSling3, LM_All_Lights_l52_ULSling4, LM_All_Lights_l52_URSling1, LM_All_Lights_l52_URSling2, LM_All_Lights_l52_URSling3, LM_All_Lights_l52_URSling4, LM_All_Lights_l53_LLSling1, LM_All_Lights_l53_LLSling2, LM_All_Lights_l53_LLSling3, LM_All_Lights_l53_LLSling4, LM_All_Lights_l53_Parts, _
  LM_All_Lights_l53_Playfield, LM_All_Lights_l53_UnderPF, LM_All_Lights_l53_sw41, LM_All_Lights_l54_LRSling1, LM_All_Lights_l54_LRSling2, LM_All_Lights_l54_LRSling3, LM_All_Lights_l54_LRSling4, LM_All_Lights_l54_Parts, LM_All_Lights_l54_Plastics, LM_All_Lights_l54_UnderPF, LM_All_Lights_l54_sw32_001, LM_All_Lights_l55_Parts, LM_All_Lights_l55_UnderPF, LM_All_Lights_l56_LLSling1, LM_All_Lights_l56_LLSling2, LM_All_Lights_l56_LLSling3, LM_All_Lights_l56_LLSling4, LM_All_Lights_l56_Parts, LM_All_Lights_l56_UnderPF, LM_All_Lights_l61a_Parts, LM_All_Lights_l61a_Plastics, LM_All_Lights_l61a_Playfield, LM_All_Lights_l63_BackPanel, LM_All_Lights_l63_Parts, LM_All_Lights_l63_Playfield, LM_All_Lights_l63_UnderPF, LM_All_Lights_l65_Parts, LM_All_Lights_l65_Plastics, LM_All_Lights_l65_Playfield, LM_All_Lights_l65_ULSling1, LM_All_Lights_l65_ULSling2, LM_All_Lights_l65_ULSling3, LM_All_Lights_l65_ULSling4, LM_All_Lights_l65_sw34, LM_All_Lights_l66_Parts, LM_All_Lights_l66_Plastics, LM_All_Lights_l66_Playfield, _
  LM_All_Lights_l66_URSling1, LM_All_Lights_l66_URSling2, LM_All_Lights_l66_URSling3, LM_All_Lights_l66_URSling4, LM_All_Lights_l66_sw37, LM_All_Lights_l66_sw38, LM_All_Lights_l67_LLSling1, LM_All_Lights_l67_LLSling2, LM_All_Lights_l67_LLSling3, LM_All_Lights_l67_LLSling4, LM_All_Lights_l67_Parts, LM_All_Lights_l67_Plastics, LM_All_Lights_l67_Playfield, LM_All_Lights_l67_UnderPF, LM_All_Lights_l67_sw41, LM_All_Lights_l68_Parts, LM_All_Lights_l69_LRSling1, LM_All_Lights_l69_LRSling2, LM_All_Lights_l69_LRSling3, LM_All_Lights_l69_LRSling4, LM_All_Lights_l69_Parts, LM_All_Lights_l69_Plastics, LM_All_Lights_l69_Playfield, LM_All_Lights_l69_UnderPF, LM_All_Lights_l7_Parts, LM_All_Lights_l70_LLSling1, LM_All_Lights_l70_LLSling2, LM_All_Lights_l70_LLSling3, LM_All_Lights_l70_LLSling4, LM_All_Lights_l70_Parts, LM_All_Lights_l70_UnderPF, LM_All_Lights_l71_LRSling1, LM_All_Lights_l71_LRSling2, LM_All_Lights_l71_LRSling3, LM_All_Lights_l71_LRSling4, LM_All_Lights_l71_Parts, LM_All_Lights_l72_Parts, _
  LM_All_Lights_l78_Parts, LM_All_Lights_l78_Plastics, LM_All_Lights_l78_Playfield, LM_All_Lights_l78_ULSling, LM_All_Lights_l78_ULSling1, LM_All_Lights_l78_ULSling2, LM_All_Lights_l78_ULSling3, LM_All_Lights_l78_ULSling4, LM_All_Lights_l78_pSpinnerRod_0, LM_All_Lights_l78_sw34, LM_All_Lights_l78_sw35, LM_All_Lights_l78b_GateRampC2Wi, LM_All_Lights_l78b_Parts, LM_All_Lights_l78b_Plastics, LM_All_Lights_l78b_Playfield, LM_All_Lights_l78b_ULSling, LM_All_Lights_l78b_ULSling1, LM_All_Lights_l78b_ULSling2, LM_All_Lights_l78b_ULSling3, LM_All_Lights_l78b_ULSling4, LM_All_Lights_l78b_pSpinnerRod_, LM_All_Lights_l79_SideInserts, LM_All_Lights_l8_Parts, LM_All_Lights_l8_Plastics, LM_All_Lights_l8_Playfield, LM_All_Lights_l8_URSling, LM_All_Lights_l8_URSling1, LM_All_Lights_l8_URSling2, LM_All_Lights_l8_URSling3, LM_All_Lights_l8_URSling4, LM_All_Lights_l8_sw25, LM_All_Lights_l8_sw26, LM_All_Lights_l81_Parts, LM_All_Lights_l81_Plastics, LM_All_Lights_l81_Playfield, LM_All_Lights_l81_ULSling1, LM_All_Lights_l81_ULSling2, _
  LM_All_Lights_l81_ULSling3, LM_All_Lights_l81_ULSling4, LM_All_Lights_l81_sw35, LM_All_Lights_l82_Parts, LM_All_Lights_l82_Plastics, LM_All_Lights_l82_Playfield, LM_All_Lights_l82_URSling1, LM_All_Lights_l82_URSling2, LM_All_Lights_l82_URSling3, LM_All_Lights_l82_URSling4, LM_All_Lights_l82_sw38, LM_All_Lights_l83_Parts, LM_All_Lights_l84_Parts, LM_All_Lights_l84_Plastics, LM_All_Lights_l84_ULSling1, LM_All_Lights_l84_ULSling2, LM_All_Lights_l84_ULSling3, LM_All_Lights_l84_ULSling4, LM_All_Lights_l84_URSling1, LM_All_Lights_l84_URSling2, LM_All_Lights_l84_URSling3, LM_All_Lights_l84_URSling4, LM_All_Lights_l84_UnderPF, LM_All_Lights_l85_LLSling1, LM_All_Lights_l85_LLSling2, LM_All_Lights_l85_LLSling3, LM_All_Lights_l85_LLSling4, LM_All_Lights_l85_Parts, LM_All_Lights_l85_Playfield, LM_All_Lights_l85_UnderPF, LM_All_Lights_l85_sw31_001, LM_All_Lights_l86_LLSling3, LM_All_Lights_l86_LLSling4, LM_All_Lights_l86_Parts, LM_All_Lights_l86_UnderPF, LM_All_Lights_l87_LRSling1, LM_All_Lights_l87_LRSling2, _
  LM_All_Lights_l87_LRSling3, LM_All_Lights_l87_LRSling4, LM_All_Lights_l87_Parts, LM_All_Lights_l87_UnderPF, LM_All_Lights_l88_Parts, LM_All_Lights_l88_UnderPF, LM_All_Lights_l9_LLSling1, LM_All_Lights_l9_LLSling2, LM_All_Lights_l9_LLSling3, LM_All_Lights_l9_LLSling4, LM_All_Lights_l9_Parts, LM_All_Lights_l9_Plastics, LM_All_Lights_l9_UnderPF, LM_All_Lights_l90_Parts, LM_All_Lights_l90_Plastics, LM_All_Lights_l90_Playfield, LM_All_Lights_l91a_Parts, LM_All_Lights_l91a_Plastics, LM_All_Lights_l91a_Playfield, LM_All_Lights_l94_Parts, LM_All_Lights_l94_Plastics, LM_All_Lights_l94_Playfield, LM_All_Lights_l94_URSling, LM_All_Lights_l94_URSling1, LM_All_Lights_l94_URSling2, LM_All_Lights_l94_URSling3, LM_All_Lights_l94_URSling4, LM_All_Lights_l94_pSpinnerRod_0, LM_All_Lights_l94_sw36, LM_All_Lights_l94_sw37, LM_All_Lights_l94b_GateRampY2Wi, LM_All_Lights_l94b_Parts, LM_All_Lights_l94b_Plastics, LM_All_Lights_l94b_Playfield, LM_All_Lights_l94b_URSling, LM_All_Lights_l94b_URSling1, LM_All_Lights_l94b_URSling2, _
  LM_All_Lights_l94b_URSling3, LM_All_Lights_l94b_URSling4, LM_All_Lights_l94b_pSpinnerRod, LM_All_Lights_l94b_pSpinnerRod_, LM_F_l12_LLSling1, LM_F_l12_LLSling2, LM_F_l12_LLSling3, LM_F_l12_LLSling4, LM_F_l12_LRSling1, LM_F_l12_LRSling2, LM_F_l12_LRSling3, LM_F_l12_LRSling4, LM_F_l12_Lflip, LM_F_l12_LflipU, LM_F_l12_NCoverL, LM_F_l12_Parts, LM_F_l12_Plastics, LM_F_l12_Playfield, LM_F_l12_Rflip, LM_F_l12_RflipU, LM_F_l12_Sling1, LM_F_l12_Sling2, LM_F_l12_UnderPF, LM_F_l12_sw31_001, LM_F_l12_sw39_001, LM_F_l12_sw41, LM_F_l29_BackPanel, LM_F_l29_GateRampIWire_002, LM_F_l29_LaneSteeringWall3, LM_F_l29_LaneSwitch1, LM_F_l29_LaneSwitch2, LM_F_l29_LaneSwitch3, LM_F_l29_LaneSwitch4, LM_F_l29_Parts, LM_F_l29_Plastics, LM_F_l29_Playfield, LM_F_l29_SideInserts, LM_F_l29_ULSling1, LM_F_l29_ULSling2, LM_F_l29_ULSling3, LM_F_l29_ULSling4, LM_F_l29_URSling, LM_F_l29_URSling1, LM_F_l29_URSling2, LM_F_l29_URSling3, LM_F_l29_URSling4, LM_F_l29_UnderPF, LM_F_l29_pSpinnerRod, LM_F_l29_pSpinnerRod_001, LM_F_l29_pSpinnerRod_004, _
  LM_F_l29_sw21, LM_F_l29_sw22, LM_F_l29_sw23, LM_F_l29_sw24, LM_F_l29_sw25, LM_F_l29_sw26, LM_F_l29_sw27, LM_F_l43_BackPanel, LM_F_l43_LaneSwitch4, LM_F_l43_Layer_1, LM_F_l43_Parts, LM_F_l43_Plastics, LM_F_l44_GateRampCWire_001, LM_F_l44_GateRampIWire_002, LM_F_l44_GateRampTWire_002, LM_F_l44_GateRampY2Wire_002, LM_F_l44_GateRampYWire_001, LM_F_l44_Lflip1, LM_F_l44_Lflip1U, LM_F_l44_Parts, LM_F_l44_Plastics, LM_F_l44_Playfield, LM_F_l44_Rflip1, LM_F_l44_Rflip1U, LM_F_l44_ULSling, LM_F_l44_ULSling1, LM_F_l44_ULSling2, LM_F_l44_ULSling3, LM_F_l44_ULSling4, LM_F_l44_URSling, LM_F_l44_URSling1, LM_F_l44_URSling2, LM_F_l44_URSling3, LM_F_l44_URSling4, LM_F_l44_UnderPF, LM_F_l44_pSpinnerRod, LM_F_l44_pSpinnerRod_001, LM_F_l44_pSpinnerRod_002, LM_F_l44_pSpinnerRod_004, LM_F_l44_sw21, LM_F_l44_sw22, LM_F_l44_sw23, LM_F_l44_sw25, LM_F_l44_sw26, LM_F_l44_sw27, LM_F_l44_sw33, LM_F_l44_sw34, LM_F_l44_sw35, LM_F_l44_sw36, LM_F_l44_sw37, LM_F_l44_sw38, LM_F_l44_sw42, LM_F_l45_Parts, LM_F_l45_UnderPF, LM_F_l57_BackPanel, _
  LM_F_l57_Layer_1, LM_F_l57_Parts, LM_F_l57_SideInserts, LM_F_l58_BackPanel, LM_F_l58_LaneSwitch1, LM_F_l58_Layer_1, LM_F_l58_Parts, LM_F_l58_SideInserts, LM_F_l59_BackPanel, LM_F_l59_Layer_1, LM_F_l59_Parts, LM_F_l59_Plastics, LM_F_l60_BackPanel, LM_F_l60_Layer_1, LM_F_l60_Parts, LM_F_l60_SideInserts, LM_F_l61_BackPanel, LM_F_l61_GateRampIWire_002, LM_F_l61_GateRampTWire_002, LM_F_l61_LaneSteeringWall3, LM_F_l61_LaneSwitch1, LM_F_l61_LaneSwitch2, LM_F_l61_LaneSwitch3, LM_F_l61_LaneSwitch4, LM_F_l61_Layer_1, LM_F_l61_Parts, LM_F_l61_Plastics, LM_F_l61_Playfield, LM_F_l61_SideInserts, LM_F_l61_ULSling, LM_F_l61_ULSling1, LM_F_l61_ULSling2, LM_F_l61_ULSling3, LM_F_l61_ULSling4, LM_F_l61_URSling, LM_F_l61_URSling1, LM_F_l61_URSling2, LM_F_l61_URSling3, LM_F_l61_URSling4, LM_F_l61_UnderPF, LM_F_l61_pSpinnerRod, LM_F_l61_pSpinnerRod_001, LM_F_l61_pSpinnerRod_003, LM_F_l61_pSpinnerRod_004, LM_F_l61_sw21, LM_F_l61_sw22, LM_F_l61_sw23, LM_F_l61_sw24, LM_F_l61_sw25, LM_F_l61_sw26, LM_F_l61_sw27, LM_F_l73_BackPanel, _
  LM_F_l73_Layer_1, LM_F_l73_Parts, LM_F_l73_SideInserts, LM_F_l75_BackPanel, LM_F_l75_Layer_1, LM_F_l75_Parts, LM_F_l75_Plastics, LM_F_l76_GateRampC2Wire_002, LM_F_l76_GateRampCWire_001, LM_F_l76_GateRampIWire_002, LM_F_l76_GateRampTWire_002, LM_F_l76_GateRampYWire_001, LM_F_l76_Lflip1, LM_F_l76_Lflip1U, LM_F_l76_Parts, LM_F_l76_Plastics, LM_F_l76_Playfield, LM_F_l76_Rflip1, LM_F_l76_Rflip1U, LM_F_l76_ULSling, LM_F_l76_ULSling1, LM_F_l76_ULSling2, LM_F_l76_ULSling3, LM_F_l76_ULSling4, LM_F_l76_URSling, LM_F_l76_URSling1, LM_F_l76_URSling2, LM_F_l76_URSling3, LM_F_l76_URSling4, LM_F_l76_UnderPF, LM_F_l76_pSpinnerRod, LM_F_l76_pSpinnerRod_001, LM_F_l76_pSpinnerRod_003, LM_F_l76_pSpinnerRod_004, LM_F_l76_sw21, LM_F_l76_sw22, LM_F_l76_sw23, LM_F_l76_sw25, LM_F_l76_sw26, LM_F_l76_sw27, LM_F_l76_sw33, LM_F_l76_sw34, LM_F_l76_sw35, LM_F_l76_sw36, LM_F_l76_sw37, LM_F_l76_sw38, LM_F_l76_sw41, LM_F_l79_BackPanel, LM_F_l79_Layer_1, LM_F_l79_Parts, LM_F_l79_SideInserts, LM_F_l89_BackPanel, LM_F_l89_Layer_1, _
  LM_F_l89_Parts, LM_F_l89_SideInserts, LM_F_l91_BackPanel, LM_F_l91_LLSling1, LM_F_l91_LLSling2, LM_F_l91_LLSling3, LM_F_l91_LLSling4, LM_F_l91_LRSling1, LM_F_l91_LRSling2, LM_F_l91_LRSling3, LM_F_l91_LRSling4, LM_F_l91_Lflip, LM_F_l91_LflipU, LM_F_l91_NCoverL, LM_F_l91_NCoverR, LM_F_l91_Parts, LM_F_l91_Plastics, LM_F_l91_Playfield, LM_F_l91_Rflip, LM_F_l91_RflipU, LM_F_l91_Sling1, LM_F_l91_Sling2, LM_F_l91_UnderPF, LM_F_l91_sw31_001, LM_F_l91_sw32_001, LM_F_l91_sw40_001, LM_F_l91_sw42, LM_F_l92_BackPanel, LM_F_l92_Layer_1, LM_F_l92_Parts, LM_F_l92_SideInserts, LM_F_l95_BackPanel, LM_F_l95_Layer_1, LM_F_l95_SideInserts, LM_GI_BackPanel, LM_GI_GateRampC2Wire_002, LM_GI_GateRampCWire_001, LM_GI_GateRampIWire_002, LM_GI_GateRampTWire_002, LM_GI_GateRampY2Wire_002, LM_GI_GateRampYWire_001, LM_GI_LLSling1, LM_GI_LLSling2, LM_GI_LLSling3, LM_GI_LLSling4, LM_GI_LRSling1, LM_GI_LRSling2, LM_GI_LRSling3, LM_GI_LRSling4, LM_GI_LaneSteeringWall3, LM_GI_LaneSwitch1, LM_GI_LaneSwitch2, LM_GI_LaneSwitch3, LM_GI_LaneSwitch4, _
  LM_GI_Layer_1, LM_GI_Lflip, LM_GI_Lflip1, LM_GI_Lflip1U, LM_GI_LflipU, LM_GI_MiddlePost, LM_GI_NCoverL, LM_GI_NCoverR, LM_GI_Parts, LM_GI_Plastics, LM_GI_Playfield, LM_GI_Playfield_001, LM_GI_Rflip, LM_GI_Rflip1, LM_GI_Rflip1U, LM_GI_RflipU, LM_GI_SideInserts, LM_GI_Sling1, LM_GI_Sling2, LM_GI_ULSling, LM_GI_ULSling1, LM_GI_ULSling2, LM_GI_ULSling3, LM_GI_ULSling4, LM_GI_URSling, LM_GI_URSling1, LM_GI_URSling2, LM_GI_URSling3, LM_GI_URSling4, LM_GI_UnderPF, LM_GI_pSpinnerRod, LM_GI_pSpinnerRod_001, LM_GI_pSpinnerRod_002, LM_GI_pSpinnerRod_003, LM_GI_pSpinnerRod_004, LM_GI_sw21, LM_GI_sw22, LM_GI_sw23, LM_GI_sw24, LM_GI_sw25, LM_GI_sw26, LM_GI_sw27, LM_GI_sw31_001, LM_GI_sw32_001, LM_GI_sw33, LM_GI_sw34, LM_GI_sw35, LM_GI_sw36, LM_GI_sw37, LM_GI_sw38, LM_GI_sw39_001, LM_GI_sw40_001, LM_GI_sw41, LM_GI_sw42, LM_L_l1_Parts, LM_L_l1_Playfield, LM_L_l1_UnderPF, LM_L_l10_Playfield, LM_L_l10_UnderPF, LM_L_l11_GateRampYWire_001, LM_L_l11_Parts, LM_L_l11_Playfield, LM_L_l11_Rflip1U, LM_L_l11_UnderPF, LM_L_l15_Parts, _
  LM_L_l15_UnderPF, LM_L_l15_pSpinnerRod_001, LM_L_l17_Playfield, LM_L_l17_UnderPF, LM_L_l18_Playfield, LM_L_l18_UnderPF, LM_L_l19_UnderPF, LM_L_l2_Playfield, LM_L_l2_UnderPF, LM_L_l20_Playfield, LM_L_l20_UnderPF, LM_L_l21_Parts, LM_L_l21_Plastics, LM_L_l21_Playfield, LM_L_l21_UnderPF, LM_L_l22_Parts, LM_L_l22_Playfield, LM_L_l22_UnderPF, LM_L_l23_Playfield, LM_L_l23_UnderPF, LM_L_l24_Parts, LM_L_l24_URSling2, LM_L_l24_URSling3, LM_L_l24_URSling4, LM_L_l24_UnderPF, LM_L_l25_Playfield, LM_L_l25_UnderPF, LM_L_l26_Parts, LM_L_l26_Playfield, LM_L_l26_UnderPF, LM_L_l27_Parts, LM_L_l27_Plastics, LM_L_l27_Playfield, LM_L_l27_UnderPF, LM_L_l27_sw24, LM_L_l3_Lflip1U, LM_L_l3_Parts, LM_L_l3_Playfield, LM_L_l3_UnderPF, LM_L_l31_Parts, LM_L_l31_UnderPF, LM_L_l33_Playfield, LM_L_l33_UnderPF, LM_L_l34_Parts, LM_L_l34_Playfield, LM_L_l34_UnderPF, LM_L_l35_Playfield, LM_L_l35_UnderPF, LM_L_l36_Playfield, LM_L_l36_UnderPF, LM_L_l37_Parts, LM_L_l37_Playfield, LM_L_l37_ULSling2, LM_L_l37_ULSling3, LM_L_l37_ULSling4, _
  LM_L_l37_UnderPF, LM_L_l38_Playfield, LM_L_l38_UnderPF, LM_L_l39_Playfield, LM_L_l39_UnderPF, LM_L_l4_Playfield, LM_L_l4_UnderPF, LM_L_l40_Parts, LM_L_l40_Plastics, LM_L_l40_Playfield, LM_L_l40_UnderPF, LM_L_l40_sw24, LM_L_l41_Playfield, LM_L_l41_UnderPF, LM_L_l42_Parts, LM_L_l42_Playfield, LM_L_l42_UnderPF, LM_L_l47_UnderPF, LM_L_l49_Parts, LM_L_l49_Playfield, LM_L_l49_UnderPF, LM_L_l5_Parts, LM_L_l5_Playfield, LM_L_l5_UnderPF, LM_L_l50_Parts, LM_L_l50_Playfield, LM_L_l50_UnderPF, LM_L_l51_LRSling1, LM_L_l51_LRSling2, LM_L_l51_LRSling3, LM_L_l51_LRSling4, LM_L_l51_Parts, LM_L_l51_Plastics, LM_L_l51_UnderPF, LM_L_l52_Parts, LM_L_l52_UnderPF, LM_L_l53_Parts, LM_L_l53_Playfield, LM_L_l53_UnderPF, LM_L_l54_Parts, LM_L_l54_Playfield, LM_L_l54_UnderPF, LM_L_l55_UnderPF, LM_L_l56_UnderPF, LM_L_l63_Parts, LM_L_l63_UnderPF, LM_L_l65_Parts, LM_L_l65_Playfield, LM_L_l65_UnderPF, LM_L_l66_Parts, LM_L_l66_Playfield, LM_L_l66_UnderPF, LM_L_l67_LLSling1, LM_L_l67_LLSling2, LM_L_l67_LLSling3, LM_L_l67_LLSling4, _
  LM_L_l67_Parts, LM_L_l67_Plastics, LM_L_l67_Playfield, LM_L_l67_UnderPF, LM_L_l68_Parts, LM_L_l68_UnderPF, LM_L_l69_Parts, LM_L_l69_Playfield, LM_L_l69_UnderPF, LM_L_l7_Playfield, LM_L_l7_UnderPF, LM_L_l70_Parts, LM_L_l70_UnderPF, LM_L_l71_UnderPF, LM_L_l72_UnderPF, LM_L_l8_Parts, LM_L_l8_Playfield, LM_L_l8_UnderPF, LM_L_l81_Playfield, LM_L_l81_UnderPF, LM_L_l82_Parts, LM_L_l82_Playfield, LM_L_l82_UnderPF, LM_L_l83_Parts, LM_L_l83_UnderPF, LM_L_l84_Parts, LM_L_l84_Playfield, LM_L_l84_UnderPF, LM_L_l85_Parts, LM_L_l85_Playfield, LM_L_l85_UnderPF, LM_L_l86_UnderPF, LM_L_l87_Parts, LM_L_l87_Playfield, LM_L_l87_UnderPF, LM_L_l88_UnderPF, LM_L_l9_Parts, LM_L_l9_Playfield, LM_L_l9_UnderPF, LM_L_l90_Parts, LM_L_l90_Plastics, LM_L_l90_UnderPF)
Dim BG_All: BG_All=Array(BM_BackPanel, BM_GateRampC2Wire_002, BM_GateRampCWire_001, BM_GateRampIWire_002, BM_GateRampTWire_002, BM_GateRampY2Wire_002, BM_GateRampYWire_001, BM_LLSling1, BM_LLSling2, BM_LLSling3, BM_LLSling4, BM_LRSling1, BM_LRSling2, BM_LRSling3, BM_LRSling4, BM_LaneSteeringWall3, BM_LaneSwitch1, BM_LaneSwitch2, BM_LaneSwitch3, BM_LaneSwitch4, BM_Layer_1, BM_Lflip, BM_Lflip1, BM_Lflip1U, BM_LflipU, BM_MiddlePost, BM_NCoverL, BM_NCoverR, BM_Parts, BM_Parts_001, BM_Plastics, BM_Playfield, BM_Playfield_001, BM_Rflip, BM_Rflip1, BM_Rflip1U, BM_RflipU, BM_SideInserts, BM_Sling1, BM_Sling2, BM_ULSling, BM_ULSling1, BM_ULSling2, BM_ULSling3, BM_ULSling4, BM_URSling, BM_URSling1, BM_URSling2, BM_URSling3, BM_URSling4, BM_UnderPF, BM_pSpinnerRod, BM_pSpinnerRod_001, BM_pSpinnerRod_002, BM_pSpinnerRod_003, BM_pSpinnerRod_004, BM_sw16, BM_sw21, BM_sw22, BM_sw23, BM_sw24, BM_sw25, BM_sw26, BM_sw27, BM_sw31_001, BM_sw32_001, BM_sw33, BM_sw34, BM_sw35, BM_sw36, BM_sw37, BM_sw38, BM_sw39_001, BM_sw40_001, _
  BM_sw41, BM_sw42, LM_All_Lights_l1_Parts, LM_All_Lights_l1_UnderPF, LM_All_Lights_l10_LRSling3, LM_All_Lights_l10_LRSling4, LM_All_Lights_l10_Parts, LM_All_Lights_l10_Plastics, LM_All_Lights_l10_UnderPF, LM_All_Lights_l11_GateRampY2Wir, LM_All_Lights_l11_GateRampYWire, LM_All_Lights_l11_Parts, LM_All_Lights_l11_Rflip1U, LM_All_Lights_l11_UnderPF, LM_All_Lights_l12a_Parts, LM_All_Lights_l12a_Plastics, LM_All_Lights_l12a_Playfield, LM_All_Lights_l15_Parts, LM_All_Lights_l15_Plastics, LM_All_Lights_l15_Playfield, LM_All_Lights_l15_ULSling1, LM_All_Lights_l15_ULSling2, LM_All_Lights_l15_ULSling3, LM_All_Lights_l15_ULSling4, LM_All_Lights_l15_pSpinnerRod_0, LM_All_Lights_l15_sw22, LM_All_Lights_l15_sw25, LM_All_Lights_l15_sw26, LM_All_Lights_l15_sw27, LM_All_Lights_l17_Parts, LM_All_Lights_l18_Parts, LM_All_Lights_l18_UnderPF, LM_All_Lights_l19_RflipU, LM_All_Lights_l2_Parts, LM_All_Lights_l2_UnderPF, LM_All_Lights_l20_Parts, LM_All_Lights_l21_Parts, LM_All_Lights_l21_Plastics, LM_All_Lights_l21_Playfield, _
  LM_All_Lights_l21_ULSling, LM_All_Lights_l21_ULSling1, LM_All_Lights_l21_ULSling2, LM_All_Lights_l21_ULSling3, LM_All_Lights_l21_ULSling4, LM_All_Lights_l21_sw25, LM_All_Lights_l22_LLSling1, LM_All_Lights_l22_LLSling2, LM_All_Lights_l22_LLSling3, LM_All_Lights_l22_LLSling4, LM_All_Lights_l22_Parts, LM_All_Lights_l23_Parts, LM_All_Lights_l23_UnderPF, LM_All_Lights_l24_GateRampTWire, LM_All_Lights_l24_Parts, LM_All_Lights_l24_URSling1, LM_All_Lights_l24_URSling2, LM_All_Lights_l24_URSling3, LM_All_Lights_l24_URSling4, LM_All_Lights_l24_UnderPF, LM_All_Lights_l24_pSpinnerRod, LM_All_Lights_l25_LLSling2, LM_All_Lights_l25_LLSling3, LM_All_Lights_l25_LLSling4, LM_All_Lights_l25_Parts, LM_All_Lights_l25_Plastics, LM_All_Lights_l25_UnderPF, LM_All_Lights_l26_LRSling1, LM_All_Lights_l26_LRSling2, LM_All_Lights_l26_LRSling3, LM_All_Lights_l26_LRSling4, LM_All_Lights_l26_Parts, LM_All_Lights_l26_Plastics, LM_All_Lights_l26_UnderPF, LM_All_Lights_l27_Parts, LM_All_Lights_l27_Plastics, LM_All_Lights_l27_Playfield, _
  LM_All_Lights_l27_UnderPF, LM_All_Lights_l27_sw24, LM_All_Lights_l27_sw26, LM_All_Lights_l27_sw27, LM_All_Lights_l29a_Parts, LM_All_Lights_l29a_Plastics, LM_All_Lights_l29a_Playfield, LM_All_Lights_l29a_pSpinnerRod_, LM_All_Lights_l29a_pSpinnerRod_, LM_All_Lights_l29a_sw24, LM_All_Lights_l3_Lflip1U, LM_All_Lights_l3_Parts, LM_All_Lights_l3_Plastics, LM_All_Lights_l3_UnderPF, LM_All_Lights_l3_sw41, LM_All_Lights_l31_Parts, LM_All_Lights_l31_Playfield, LM_All_Lights_l31_URSling1, LM_All_Lights_l31_URSling2, LM_All_Lights_l31_URSling3, LM_All_Lights_l31_URSling4, LM_All_Lights_l31_sw23, LM_All_Lights_l31_sw26, LM_All_Lights_l31_sw27, LM_All_Lights_l33_Parts, LM_All_Lights_l34_Lflip1U, LM_All_Lights_l34_Parts, LM_All_Lights_l34_Plastics, LM_All_Lights_l34_UnderPF, LM_All_Lights_l34_sw41, LM_All_Lights_l35_Parts, LM_All_Lights_l36_Parts, LM_All_Lights_l37_GateRampIWire, LM_All_Lights_l37_Parts, LM_All_Lights_l37_Plastics, LM_All_Lights_l37_Playfield, LM_All_Lights_l37_ULSling, LM_All_Lights_l37_ULSling1, _
  LM_All_Lights_l37_ULSling2, LM_All_Lights_l37_ULSling3, LM_All_Lights_l37_ULSling4, LM_All_Lights_l37_UnderPF, LM_All_Lights_l37_pSpinnerRod_0, LM_All_Lights_l38_Parts, LM_All_Lights_l38_UnderPF, LM_All_Lights_l39_LRSling1, LM_All_Lights_l39_LRSling2, LM_All_Lights_l39_LRSling3, LM_All_Lights_l39_LRSling4, LM_All_Lights_l39_Parts, LM_All_Lights_l39_UnderPF, LM_All_Lights_l4_Parts, LM_All_Lights_l4_UnderPF, LM_All_Lights_l40_Parts, LM_All_Lights_l40_Plastics, LM_All_Lights_l40_Playfield, LM_All_Lights_l40_UnderPF, LM_All_Lights_l40_pSpinnerRod_0, LM_All_Lights_l40_sw24, LM_All_Lights_l40_sw27, LM_All_Lights_l41_Parts, LM_All_Lights_l41_Plastics, LM_All_Lights_l41_UnderPF, LM_All_Lights_l42_GateRampYWire, LM_All_Lights_l42_Parts, LM_All_Lights_l42_Rflip1U, LM_All_Lights_l42_UnderPF, LM_All_Lights_l42_sw42, LM_All_Lights_l47_UnderPF, LM_All_Lights_l49_Parts, LM_All_Lights_l49_Plastics, LM_All_Lights_l49_Playfield, LM_All_Lights_l49_ULSling1, LM_All_Lights_l49_ULSling2, LM_All_Lights_l49_ULSling3, _
  LM_All_Lights_l49_ULSling4, LM_All_Lights_l49_sw33, LM_All_Lights_l5_GateRampYWire_, LM_All_Lights_l5_Parts, LM_All_Lights_l5_UnderPF, LM_All_Lights_l50_Parts, LM_All_Lights_l50_Plastics, LM_All_Lights_l50_Playfield, LM_All_Lights_l50_URSling1, LM_All_Lights_l50_URSling2, LM_All_Lights_l50_URSling3, LM_All_Lights_l50_URSling4, LM_All_Lights_l50_sw36, LM_All_Lights_l50_sw37, LM_All_Lights_l51_LRSling1, LM_All_Lights_l51_LRSling2, LM_All_Lights_l51_LRSling3, LM_All_Lights_l51_LRSling4, LM_All_Lights_l51_Parts, LM_All_Lights_l51_Plastics, LM_All_Lights_l51_Playfield, LM_All_Lights_l51_UnderPF, LM_All_Lights_l52_Parts, LM_All_Lights_l52_ULSling1, LM_All_Lights_l52_ULSling2, LM_All_Lights_l52_ULSling3, LM_All_Lights_l52_ULSling4, LM_All_Lights_l52_URSling1, LM_All_Lights_l52_URSling2, LM_All_Lights_l52_URSling3, LM_All_Lights_l52_URSling4, LM_All_Lights_l53_LLSling1, LM_All_Lights_l53_LLSling2, LM_All_Lights_l53_LLSling3, LM_All_Lights_l53_LLSling4, LM_All_Lights_l53_Parts, LM_All_Lights_l53_Playfield, _
  LM_All_Lights_l53_UnderPF, LM_All_Lights_l53_sw41, LM_All_Lights_l54_LRSling1, LM_All_Lights_l54_LRSling2, LM_All_Lights_l54_LRSling3, LM_All_Lights_l54_LRSling4, LM_All_Lights_l54_Parts, LM_All_Lights_l54_Plastics, LM_All_Lights_l54_UnderPF, LM_All_Lights_l54_sw32_001, LM_All_Lights_l55_Parts, LM_All_Lights_l55_UnderPF, LM_All_Lights_l56_LLSling1, LM_All_Lights_l56_LLSling2, LM_All_Lights_l56_LLSling3, LM_All_Lights_l56_LLSling4, LM_All_Lights_l56_Parts, LM_All_Lights_l56_UnderPF, LM_All_Lights_l61a_Parts, LM_All_Lights_l61a_Plastics, LM_All_Lights_l61a_Playfield, LM_All_Lights_l63_BackPanel, LM_All_Lights_l63_Parts, LM_All_Lights_l63_Playfield, LM_All_Lights_l63_UnderPF, LM_All_Lights_l65_Parts, LM_All_Lights_l65_Plastics, LM_All_Lights_l65_Playfield, LM_All_Lights_l65_ULSling1, LM_All_Lights_l65_ULSling2, LM_All_Lights_l65_ULSling3, LM_All_Lights_l65_ULSling4, LM_All_Lights_l65_sw34, LM_All_Lights_l66_Parts, LM_All_Lights_l66_Plastics, LM_All_Lights_l66_Playfield, LM_All_Lights_l66_URSling1, _
  LM_All_Lights_l66_URSling2, LM_All_Lights_l66_URSling3, LM_All_Lights_l66_URSling4, LM_All_Lights_l66_sw37, LM_All_Lights_l66_sw38, LM_All_Lights_l67_LLSling1, LM_All_Lights_l67_LLSling2, LM_All_Lights_l67_LLSling3, LM_All_Lights_l67_LLSling4, LM_All_Lights_l67_Parts, LM_All_Lights_l67_Plastics, LM_All_Lights_l67_Playfield, LM_All_Lights_l67_UnderPF, LM_All_Lights_l67_sw41, LM_All_Lights_l68_Parts, LM_All_Lights_l69_LRSling1, LM_All_Lights_l69_LRSling2, LM_All_Lights_l69_LRSling3, LM_All_Lights_l69_LRSling4, LM_All_Lights_l69_Parts, LM_All_Lights_l69_Plastics, LM_All_Lights_l69_Playfield, LM_All_Lights_l69_UnderPF, LM_All_Lights_l7_Parts, LM_All_Lights_l70_LLSling1, LM_All_Lights_l70_LLSling2, LM_All_Lights_l70_LLSling3, LM_All_Lights_l70_LLSling4, LM_All_Lights_l70_Parts, LM_All_Lights_l70_UnderPF, LM_All_Lights_l71_LRSling1, LM_All_Lights_l71_LRSling2, LM_All_Lights_l71_LRSling3, LM_All_Lights_l71_LRSling4, LM_All_Lights_l71_Parts, LM_All_Lights_l72_Parts, LM_All_Lights_l78_Parts, _
  LM_All_Lights_l78_Plastics, LM_All_Lights_l78_Playfield, LM_All_Lights_l78_ULSling, LM_All_Lights_l78_ULSling1, LM_All_Lights_l78_ULSling2, LM_All_Lights_l78_ULSling3, LM_All_Lights_l78_ULSling4, LM_All_Lights_l78_pSpinnerRod_0, LM_All_Lights_l78_sw34, LM_All_Lights_l78_sw35, LM_All_Lights_l78b_GateRampC2Wi, LM_All_Lights_l78b_Parts, LM_All_Lights_l78b_Plastics, LM_All_Lights_l78b_Playfield, LM_All_Lights_l78b_ULSling, LM_All_Lights_l78b_ULSling1, LM_All_Lights_l78b_ULSling2, LM_All_Lights_l78b_ULSling3, LM_All_Lights_l78b_ULSling4, LM_All_Lights_l78b_pSpinnerRod_, LM_All_Lights_l79_SideInserts, LM_All_Lights_l8_Parts, LM_All_Lights_l8_Plastics, LM_All_Lights_l8_Playfield, LM_All_Lights_l8_URSling, LM_All_Lights_l8_URSling1, LM_All_Lights_l8_URSling2, LM_All_Lights_l8_URSling3, LM_All_Lights_l8_URSling4, LM_All_Lights_l8_sw25, LM_All_Lights_l8_sw26, LM_All_Lights_l81_Parts, LM_All_Lights_l81_Plastics, LM_All_Lights_l81_Playfield, LM_All_Lights_l81_ULSling1, LM_All_Lights_l81_ULSling2, _
  LM_All_Lights_l81_ULSling3, LM_All_Lights_l81_ULSling4, LM_All_Lights_l81_sw35, LM_All_Lights_l82_Parts, LM_All_Lights_l82_Plastics, LM_All_Lights_l82_Playfield, LM_All_Lights_l82_URSling1, LM_All_Lights_l82_URSling2, LM_All_Lights_l82_URSling3, LM_All_Lights_l82_URSling4, LM_All_Lights_l82_sw38, LM_All_Lights_l83_Parts, LM_All_Lights_l84_Parts, LM_All_Lights_l84_Plastics, LM_All_Lights_l84_ULSling1, LM_All_Lights_l84_ULSling2, LM_All_Lights_l84_ULSling3, LM_All_Lights_l84_ULSling4, LM_All_Lights_l84_URSling1, LM_All_Lights_l84_URSling2, LM_All_Lights_l84_URSling3, LM_All_Lights_l84_URSling4, LM_All_Lights_l84_UnderPF, LM_All_Lights_l85_LLSling1, LM_All_Lights_l85_LLSling2, LM_All_Lights_l85_LLSling3, LM_All_Lights_l85_LLSling4, LM_All_Lights_l85_Parts, LM_All_Lights_l85_Playfield, LM_All_Lights_l85_UnderPF, LM_All_Lights_l85_sw31_001, LM_All_Lights_l86_LLSling3, LM_All_Lights_l86_LLSling4, LM_All_Lights_l86_Parts, LM_All_Lights_l86_UnderPF, LM_All_Lights_l87_LRSling1, LM_All_Lights_l87_LRSling2, _
  LM_All_Lights_l87_LRSling3, LM_All_Lights_l87_LRSling4, LM_All_Lights_l87_Parts, LM_All_Lights_l87_UnderPF, LM_All_Lights_l88_Parts, LM_All_Lights_l88_UnderPF, LM_All_Lights_l9_LLSling1, LM_All_Lights_l9_LLSling2, LM_All_Lights_l9_LLSling3, LM_All_Lights_l9_LLSling4, LM_All_Lights_l9_Parts, LM_All_Lights_l9_Plastics, LM_All_Lights_l9_UnderPF, LM_All_Lights_l90_Parts, LM_All_Lights_l90_Plastics, LM_All_Lights_l90_Playfield, LM_All_Lights_l91a_Parts, LM_All_Lights_l91a_Plastics, LM_All_Lights_l91a_Playfield, LM_All_Lights_l94_Parts, LM_All_Lights_l94_Plastics, LM_All_Lights_l94_Playfield, LM_All_Lights_l94_URSling, LM_All_Lights_l94_URSling1, LM_All_Lights_l94_URSling2, LM_All_Lights_l94_URSling3, LM_All_Lights_l94_URSling4, LM_All_Lights_l94_pSpinnerRod_0, LM_All_Lights_l94_sw36, LM_All_Lights_l94_sw37, LM_All_Lights_l94b_GateRampY2Wi, LM_All_Lights_l94b_Parts, LM_All_Lights_l94b_Plastics, LM_All_Lights_l94b_Playfield, LM_All_Lights_l94b_URSling, LM_All_Lights_l94b_URSling1, LM_All_Lights_l94b_URSling2, _
  LM_All_Lights_l94b_URSling3, LM_All_Lights_l94b_URSling4, LM_All_Lights_l94b_pSpinnerRod, LM_All_Lights_l94b_pSpinnerRod_, LM_F_l12_LLSling1, LM_F_l12_LLSling2, LM_F_l12_LLSling3, LM_F_l12_LLSling4, LM_F_l12_LRSling1, LM_F_l12_LRSling2, LM_F_l12_LRSling3, LM_F_l12_LRSling4, LM_F_l12_Lflip, LM_F_l12_LflipU, LM_F_l12_NCoverL, LM_F_l12_Parts, LM_F_l12_Plastics, LM_F_l12_Playfield, LM_F_l12_Rflip, LM_F_l12_RflipU, LM_F_l12_Sling1, LM_F_l12_Sling2, LM_F_l12_UnderPF, LM_F_l12_sw31_001, LM_F_l12_sw39_001, LM_F_l12_sw41, LM_F_l29_BackPanel, LM_F_l29_GateRampIWire_002, LM_F_l29_LaneSteeringWall3, LM_F_l29_LaneSwitch1, LM_F_l29_LaneSwitch2, LM_F_l29_LaneSwitch3, LM_F_l29_LaneSwitch4, LM_F_l29_Parts, LM_F_l29_Plastics, LM_F_l29_Playfield, LM_F_l29_SideInserts, LM_F_l29_ULSling1, LM_F_l29_ULSling2, LM_F_l29_ULSling3, LM_F_l29_ULSling4, LM_F_l29_URSling, LM_F_l29_URSling1, LM_F_l29_URSling2, LM_F_l29_URSling3, LM_F_l29_URSling4, LM_F_l29_UnderPF, LM_F_l29_pSpinnerRod, LM_F_l29_pSpinnerRod_001, LM_F_l29_pSpinnerRod_004, _
  LM_F_l29_sw21, LM_F_l29_sw22, LM_F_l29_sw23, LM_F_l29_sw24, LM_F_l29_sw25, LM_F_l29_sw26, LM_F_l29_sw27, LM_F_l43_BackPanel, LM_F_l43_LaneSwitch4, LM_F_l43_Layer_1, LM_F_l43_Parts, LM_F_l43_Plastics, LM_F_l44_GateRampCWire_001, LM_F_l44_GateRampIWire_002, LM_F_l44_GateRampTWire_002, LM_F_l44_GateRampY2Wire_002, LM_F_l44_GateRampYWire_001, LM_F_l44_Lflip1, LM_F_l44_Lflip1U, LM_F_l44_Parts, LM_F_l44_Plastics, LM_F_l44_Playfield, LM_F_l44_Rflip1, LM_F_l44_Rflip1U, LM_F_l44_ULSling, LM_F_l44_ULSling1, LM_F_l44_ULSling2, LM_F_l44_ULSling3, LM_F_l44_ULSling4, LM_F_l44_URSling, LM_F_l44_URSling1, LM_F_l44_URSling2, LM_F_l44_URSling3, LM_F_l44_URSling4, LM_F_l44_UnderPF, LM_F_l44_pSpinnerRod, LM_F_l44_pSpinnerRod_001, LM_F_l44_pSpinnerRod_002, LM_F_l44_pSpinnerRod_004, LM_F_l44_sw21, LM_F_l44_sw22, LM_F_l44_sw23, LM_F_l44_sw25, LM_F_l44_sw26, LM_F_l44_sw27, LM_F_l44_sw33, LM_F_l44_sw34, LM_F_l44_sw35, LM_F_l44_sw36, LM_F_l44_sw37, LM_F_l44_sw38, LM_F_l44_sw42, LM_F_l45_Parts, LM_F_l45_UnderPF, LM_F_l57_BackPanel, _
  LM_F_l57_Layer_1, LM_F_l57_Parts, LM_F_l57_SideInserts, LM_F_l58_BackPanel, LM_F_l58_LaneSwitch1, LM_F_l58_Layer_1, LM_F_l58_Parts, LM_F_l58_SideInserts, LM_F_l59_BackPanel, LM_F_l59_Layer_1, LM_F_l59_Parts, LM_F_l59_Plastics, LM_F_l60_BackPanel, LM_F_l60_Layer_1, LM_F_l60_Parts, LM_F_l60_SideInserts, LM_F_l61_BackPanel, LM_F_l61_GateRampIWire_002, LM_F_l61_GateRampTWire_002, LM_F_l61_LaneSteeringWall3, LM_F_l61_LaneSwitch1, LM_F_l61_LaneSwitch2, LM_F_l61_LaneSwitch3, LM_F_l61_LaneSwitch4, LM_F_l61_Layer_1, LM_F_l61_Parts, LM_F_l61_Plastics, LM_F_l61_Playfield, LM_F_l61_SideInserts, LM_F_l61_ULSling, LM_F_l61_ULSling1, LM_F_l61_ULSling2, LM_F_l61_ULSling3, LM_F_l61_ULSling4, LM_F_l61_URSling, LM_F_l61_URSling1, LM_F_l61_URSling2, LM_F_l61_URSling3, LM_F_l61_URSling4, LM_F_l61_UnderPF, LM_F_l61_pSpinnerRod, LM_F_l61_pSpinnerRod_001, LM_F_l61_pSpinnerRod_003, LM_F_l61_pSpinnerRod_004, LM_F_l61_sw21, LM_F_l61_sw22, LM_F_l61_sw23, LM_F_l61_sw24, LM_F_l61_sw25, LM_F_l61_sw26, LM_F_l61_sw27, LM_F_l73_BackPanel, _
  LM_F_l73_Layer_1, LM_F_l73_Parts, LM_F_l73_SideInserts, LM_F_l75_BackPanel, LM_F_l75_Layer_1, LM_F_l75_Parts, LM_F_l75_Plastics, LM_F_l76_GateRampC2Wire_002, LM_F_l76_GateRampCWire_001, LM_F_l76_GateRampIWire_002, LM_F_l76_GateRampTWire_002, LM_F_l76_GateRampYWire_001, LM_F_l76_Lflip1, LM_F_l76_Lflip1U, LM_F_l76_Parts, LM_F_l76_Plastics, LM_F_l76_Playfield, LM_F_l76_Rflip1, LM_F_l76_Rflip1U, LM_F_l76_ULSling, LM_F_l76_ULSling1, LM_F_l76_ULSling2, LM_F_l76_ULSling3, LM_F_l76_ULSling4, LM_F_l76_URSling, LM_F_l76_URSling1, LM_F_l76_URSling2, LM_F_l76_URSling3, LM_F_l76_URSling4, LM_F_l76_UnderPF, LM_F_l76_pSpinnerRod, LM_F_l76_pSpinnerRod_001, LM_F_l76_pSpinnerRod_003, LM_F_l76_pSpinnerRod_004, LM_F_l76_sw21, LM_F_l76_sw22, LM_F_l76_sw23, LM_F_l76_sw25, LM_F_l76_sw26, LM_F_l76_sw27, LM_F_l76_sw33, LM_F_l76_sw34, LM_F_l76_sw35, LM_F_l76_sw36, LM_F_l76_sw37, LM_F_l76_sw38, LM_F_l76_sw41, LM_F_l79_BackPanel, LM_F_l79_Layer_1, LM_F_l79_Parts, LM_F_l79_SideInserts, LM_F_l89_BackPanel, LM_F_l89_Layer_1, _
  LM_F_l89_Parts, LM_F_l89_SideInserts, LM_F_l91_BackPanel, LM_F_l91_LLSling1, LM_F_l91_LLSling2, LM_F_l91_LLSling3, LM_F_l91_LLSling4, LM_F_l91_LRSling1, LM_F_l91_LRSling2, LM_F_l91_LRSling3, LM_F_l91_LRSling4, LM_F_l91_Lflip, LM_F_l91_LflipU, LM_F_l91_NCoverL, LM_F_l91_NCoverR, LM_F_l91_Parts, LM_F_l91_Plastics, LM_F_l91_Playfield, LM_F_l91_Rflip, LM_F_l91_RflipU, LM_F_l91_Sling1, LM_F_l91_Sling2, LM_F_l91_UnderPF, LM_F_l91_sw31_001, LM_F_l91_sw32_001, LM_F_l91_sw40_001, LM_F_l91_sw42, LM_F_l92_BackPanel, LM_F_l92_Layer_1, LM_F_l92_Parts, LM_F_l92_SideInserts, LM_F_l95_BackPanel, LM_F_l95_Layer_1, LM_F_l95_SideInserts, LM_GI_BackPanel, LM_GI_GateRampC2Wire_002, LM_GI_GateRampCWire_001, LM_GI_GateRampIWire_002, LM_GI_GateRampTWire_002, LM_GI_GateRampY2Wire_002, LM_GI_GateRampYWire_001, LM_GI_LLSling1, LM_GI_LLSling2, LM_GI_LLSling3, LM_GI_LLSling4, LM_GI_LRSling1, LM_GI_LRSling2, LM_GI_LRSling3, LM_GI_LRSling4, LM_GI_LaneSteeringWall3, LM_GI_LaneSwitch1, LM_GI_LaneSwitch2, LM_GI_LaneSwitch3, LM_GI_LaneSwitch4, _
  LM_GI_Layer_1, LM_GI_Lflip, LM_GI_Lflip1, LM_GI_Lflip1U, LM_GI_LflipU, LM_GI_MiddlePost, LM_GI_NCoverL, LM_GI_NCoverR, LM_GI_Parts, LM_GI_Plastics, LM_GI_Playfield, LM_GI_Playfield_001, LM_GI_Rflip, LM_GI_Rflip1, LM_GI_Rflip1U, LM_GI_RflipU, LM_GI_SideInserts, LM_GI_Sling1, LM_GI_Sling2, LM_GI_ULSling, LM_GI_ULSling1, LM_GI_ULSling2, LM_GI_ULSling3, LM_GI_ULSling4, LM_GI_URSling, LM_GI_URSling1, LM_GI_URSling2, LM_GI_URSling3, LM_GI_URSling4, LM_GI_UnderPF, LM_GI_pSpinnerRod, LM_GI_pSpinnerRod_001, LM_GI_pSpinnerRod_002, LM_GI_pSpinnerRod_003, LM_GI_pSpinnerRod_004, LM_GI_sw21, LM_GI_sw22, LM_GI_sw23, LM_GI_sw24, LM_GI_sw25, LM_GI_sw26, LM_GI_sw27, LM_GI_sw31_001, LM_GI_sw32_001, LM_GI_sw33, LM_GI_sw34, LM_GI_sw35, LM_GI_sw36, LM_GI_sw37, LM_GI_sw38, LM_GI_sw39_001, LM_GI_sw40_001, LM_GI_sw41, LM_GI_sw42, LM_L_l1_Parts, LM_L_l1_Playfield, LM_L_l1_UnderPF, LM_L_l10_Playfield, LM_L_l10_UnderPF, LM_L_l11_GateRampYWire_001, LM_L_l11_Parts, LM_L_l11_Playfield, LM_L_l11_Rflip1U, LM_L_l11_UnderPF, LM_L_l15_Parts, _
  LM_L_l15_UnderPF, LM_L_l15_pSpinnerRod_001, LM_L_l17_Playfield, LM_L_l17_UnderPF, LM_L_l18_Playfield, LM_L_l18_UnderPF, LM_L_l19_UnderPF, LM_L_l2_Playfield, LM_L_l2_UnderPF, LM_L_l20_Playfield, LM_L_l20_UnderPF, LM_L_l21_Parts, LM_L_l21_Plastics, LM_L_l21_Playfield, LM_L_l21_UnderPF, LM_L_l22_Parts, LM_L_l22_Playfield, LM_L_l22_UnderPF, LM_L_l23_Playfield, LM_L_l23_UnderPF, LM_L_l24_Parts, LM_L_l24_URSling2, LM_L_l24_URSling3, LM_L_l24_URSling4, LM_L_l24_UnderPF, LM_L_l25_Playfield, LM_L_l25_UnderPF, LM_L_l26_Parts, LM_L_l26_Playfield, LM_L_l26_UnderPF, LM_L_l27_Parts, LM_L_l27_Plastics, LM_L_l27_Playfield, LM_L_l27_UnderPF, LM_L_l27_sw24, LM_L_l3_Lflip1U, LM_L_l3_Parts, LM_L_l3_Playfield, LM_L_l3_UnderPF, LM_L_l31_Parts, LM_L_l31_UnderPF, LM_L_l33_Playfield, LM_L_l33_UnderPF, LM_L_l34_Parts, LM_L_l34_Playfield, LM_L_l34_UnderPF, LM_L_l35_Playfield, LM_L_l35_UnderPF, LM_L_l36_Playfield, LM_L_l36_UnderPF, LM_L_l37_Parts, LM_L_l37_Playfield, LM_L_l37_ULSling2, LM_L_l37_ULSling3, LM_L_l37_ULSling4, _
  LM_L_l37_UnderPF, LM_L_l38_Playfield, LM_L_l38_UnderPF, LM_L_l39_Playfield, LM_L_l39_UnderPF, LM_L_l4_Playfield, LM_L_l4_UnderPF, LM_L_l40_Parts, LM_L_l40_Plastics, LM_L_l40_Playfield, LM_L_l40_UnderPF, LM_L_l40_sw24, LM_L_l41_Playfield, LM_L_l41_UnderPF, LM_L_l42_Parts, LM_L_l42_Playfield, LM_L_l42_UnderPF, LM_L_l47_UnderPF, LM_L_l49_Parts, LM_L_l49_Playfield, LM_L_l49_UnderPF, LM_L_l5_Parts, LM_L_l5_Playfield, LM_L_l5_UnderPF, LM_L_l50_Parts, LM_L_l50_Playfield, LM_L_l50_UnderPF, LM_L_l51_LRSling1, LM_L_l51_LRSling2, LM_L_l51_LRSling3, LM_L_l51_LRSling4, LM_L_l51_Parts, LM_L_l51_Plastics, LM_L_l51_UnderPF, LM_L_l52_Parts, LM_L_l52_UnderPF, LM_L_l53_Parts, LM_L_l53_Playfield, LM_L_l53_UnderPF, LM_L_l54_Parts, LM_L_l54_Playfield, LM_L_l54_UnderPF, LM_L_l55_UnderPF, LM_L_l56_UnderPF, LM_L_l63_Parts, LM_L_l63_UnderPF, LM_L_l65_Parts, LM_L_l65_Playfield, LM_L_l65_UnderPF, LM_L_l66_Parts, LM_L_l66_Playfield, LM_L_l66_UnderPF, LM_L_l67_LLSling1, LM_L_l67_LLSling2, LM_L_l67_LLSling3, LM_L_l67_LLSling4, _
  LM_L_l67_Parts, LM_L_l67_Plastics, LM_L_l67_Playfield, LM_L_l67_UnderPF, LM_L_l68_Parts, LM_L_l68_UnderPF, LM_L_l69_Parts, LM_L_l69_Playfield, LM_L_l69_UnderPF, LM_L_l7_Playfield, LM_L_l7_UnderPF, LM_L_l70_Parts, LM_L_l70_UnderPF, LM_L_l71_UnderPF, LM_L_l72_UnderPF, LM_L_l8_Parts, LM_L_l8_Playfield, LM_L_l8_UnderPF, LM_L_l81_Playfield, LM_L_l81_UnderPF, LM_L_l82_Parts, LM_L_l82_Playfield, LM_L_l82_UnderPF, LM_L_l83_Parts, LM_L_l83_UnderPF, LM_L_l84_Parts, LM_L_l84_Playfield, LM_L_l84_UnderPF, LM_L_l85_Parts, LM_L_l85_Playfield, LM_L_l85_UnderPF, LM_L_l86_UnderPF, LM_L_l87_Parts, LM_L_l87_Playfield, LM_L_l87_UnderPF, LM_L_l88_UnderPF, LM_L_l9_Parts, LM_L_l9_Playfield, LM_L_l9_UnderPF, LM_L_l90_Parts, LM_L_l90_Plastics, LM_L_l90_UnderPF)
' VLM  Arrays - End

' Thalamus : Exit in a clean and proper way
Sub Table1_exit
  Controller.Pause = False
  Controller.Stop
End Sub


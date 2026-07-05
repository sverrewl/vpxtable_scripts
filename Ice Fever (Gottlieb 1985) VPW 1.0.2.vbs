' /$$$$$$  /$$$$$$  /$$$$$$$$       /$$$$$$$$ /$$$$$$$$ /$$    /$$ /$$$$$$$$ /$$$$$$$
'|_  $$_/ /$$__  $$| $$_____/      | $$_____/| $$_____/| $$   | $$| $$_____/| $$__  $$
'  | $$  | $$  \__/| $$            | $$      | $$      | $$   | $$| $$      | $$  \ $$
'  | $$  | $$      | $$$$$         | $$$$$   | $$$$$   |  $$ / $$/| $$$$$   | $$$$$$$/
'  | $$  | $$      | $$__/         | $$__/   | $$__/    \  $$ $$/ | $$__/   | $$__  $$
'  | $$  | $$    $$| $$            | $$      | $$        \  $$$/  | $$      | $$  \ $$
' /$$$$$$|  $$$$$$/| $$$$$$$$      | $$      | $$$$$$$$   \  $/   | $$$$$$$$| $$  | $$
'|______/ \______/ |________/      |__/      |________/    \_/    |________/|__/  |__/
'
'

'
' Ice Fever (Gottlieb 1985)
' https://www.ipdb.org/machine.cgi?id=1260
'
'
'
' VPW Team
'   MechaEnron - Team Lead
'   DGrimmReaper - VR Lead
'   studly_do_right - Test Lead
'   apophis - VPX & code help
'   Sixtoe - Lighting and physics help
'   Nestorgian - VPX help
'   mcarter78 - VPX help


'*********************************************************************************************************************************
' === TABLE OF CONTENTS  ===
'
' You can quickly jump to a section by searching the four letter tag (ZXXX)
'
' ZVAR: Constants and Global Variable
' ZTIM: Timers
' ZINI: Table Initialization and Exiting
'   ZOPT: User Options
'   ZGIC: GI Colors
'   ZMAT: General Math Functions
' ZANI: Misc Animations
'   ZBBR: Ball Brightness
'   ZRBR: Room Brightness
' ZKEY: Key Press Handling
' ZSOL: Solenoids & Flashers
' ZDRN: Drain, Trough, and Ball Release
' ZFLP: Flippers
' ZSLG: Slingshot Animations
' ZSSC: Slingshot Corrections
' ZSWI: Switches
'   ZBRL: Ball Rolling and Drop Sounds
'   ZFLE: Fleep Mechanical Sounds
' ZPHY: General Advice on Physics
' ZNFF: Flipper Corrections
'   ZDMP: Rubber Dampeners
'   ZRDT: Drop Targets
' ZRST: Stand-Up Targets
'   ZBOU: VPW TargetBouncer for targets and posts
'   ZGIU: GI updates
' ZSHA: Ambient Ball Shadows
' ZVRR: VR Code
' ZDTD: Desktop Digits
' ZVLM: VLM Arrays


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
Const BallSize = 50   'Ball size must be 50
Const BallMass = 1    'Ball mass must be 1
Const tnob = 2      'Total number of balls on the playfield including captive balls.
Const lob = 1     'Total number of locked balls

Dim tablewidth: tablewidth = Table1.width
Dim tableheight: tableheight = Table1.height

'  Standard definitions
Const cGameName = "icefever"    'PinMAME ROM name
Const UseSolenoids = 2      '1 = Normal Flippers, 2 = Fastflips
Const UseLamps = 1        '0 = Custom lamp handling, 1 = Built-in VPX handling (using light number in light timer)
Const UseSync = 0
Const HandleMech = 0
Const SSolenoidOn = ""      'Sound sample used for this, obsolete.
Const SSolenoidOff = ""     ' ^
Const SFlipperOn = ""     ' ^
Const SFlipperOff = ""      ' ^
Const SCoin = ""        ' ^

Dim DesktopMode: DesktopMode = Table1.ShowDT

Dim VRPreview : VRPreview = False
Dim VRMode : VRMode = False
Dim VR_Obj, cab_mode

if Not DesktopMode and VRMode=False Then cab_mode=1 Else cab_mode=0

If RenderingMode = 2 Or VRPreview Then
  VRMode = True
  PLeftSideRail.visible = 1
  PRightSideRail.visible = 1
  Primitive19.visible = 1
  For each VR_Obj in VRCab : VR_Obj.visible = 1 : Next
  For each VR_Obj in VRBG  : VR_Obj.visible = 1 : Next
  For each VR_Obj in VRMin: VR_Obj.visible = 1 : Next
  For each VR_Obj in VRBrackets: VR_Obj.sidevisible = 1 : Next
  VR_Puck.visible = 1
Else
  VRMode = False
  For each VR_Obj in VRCab : VR_Obj.visible = 0 : Next
  For each VR_Obj in VRBG  : VR_Obj.visible = 0 : Next
  For each VR_Obj in VRMin: VR_Obj.visible = 0 : Next
  For each VR_Obj in VRBrackets: VR_Obj.sidevisible = 0 : Next
  VR_Puck.visible = 0
End If

LoadVPM "03060000", "sys80.VBS", 3.02

'******************************************************
'  ZTIM: Timers
'******************************************************

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
  If VRMode Then
    VRDisplayTimer
  Else
    UpdateLeds
  End if
End Sub

'The CorTimer interval should be 10. It's sole purpose is to update the Cor (physics) calculations
CorTimer.Interval = 10
CorTimer.Enabled = true
Sub CorTimer_Timer(): Cor.Update: End Sub

'******************************************************
'  ZINI: Table Initialization and Exiting
'******************************************************
Dim IFBall1, IFCaptiveBall, gBOT

Sub Table1_Init
  Dim BL
  vpminit Me

  vpmMapLights AllLamps

  With Controller
    .GameName = cGameName
    If Err Then MsgBox "Can't start Game " & cGameName & vbNewLine & Err.Description:Exit Sub
    .SplashInfoLine = "Ice Fever (Gottleib 1985)"
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

  'Nudging
  vpmNudge.TiltSwitch=57
  vpmNudge.Sensitivity=4
  vpmNudge.TiltObj = Array(Bumper1, Bumper2, Bumper3, Bumper4, LeftSlingshot, RightSlingshot)

  'Main Timer init
  PinMAMETimer.Interval = PinMAMEInterval
  PinMAMETimer.Enabled = 1

  'Trough - Creates a ball in the kicker switch and gives that ball used an individual name.
  Set IFBall1 = BallSpawn.CreateSizedballWithMass(Ballsize/2,Ballmass)
  vpmTimer.AddTimer 300, "BallSpawn.kick 180,1 '"   'Creates a ball from the "captiveball" kicker
  vpmTimer.AddTimer 310, "BallSpawn.enabled= 0 '"   'Permenantly Disables Captive Ball Kicker
  '***Captive Ball Creation
  Set IFCaptiveBall = captiveBall.CreateSizedballWithMass(Ballsize/2,Ballmass)
  vpmTimer.AddTimer 300, "captiveball.kick 180,1 '"   'Creates a ball from the "captiveball" kicker
  vpmTimer.AddTimer 310, "captiveball.enabled= 0 '"   'Permenantly Disables Captive Ball Kicker
  '***Setting up a ball array (collection), must contain all the balls you create on the table.
  gBOT = Array(IFCaptiveBall, IFBall1)

  InitSlings

  'Forces the trough switches to "on" at table boot so the game logic knows there are balls in the trough.
  Controller.Switch(67) = 1
  Setup_Backglass
End Sub

Sub Table_Paused:Controller.Pause = 1:End Sub
Sub Table_unPaused:Controller.Pause = 0:End Sub
Sub Table_exit()
  Controller.Pause = False
  Controller.Stop
End Sub

'*******************************************
'  ZOPT: User Options
'*******************************************

Dim LightLevel : LightLevel = 0.25        ' Level of room lighting (0 to 1), where 0 is dark and 100 is brightest
Dim ColorLUT : ColorLUT = 1           ' Color desaturation LUTs: 1 to 11, where 1 is normal and 11 is black'n'white
Dim VolumeDial : VolumeDial = 0.8             ' Overall Mechanical sound effect volume. Recommended values should be no greater than 1.
Dim BallRollVolume : BallRollVolume = 0.5     ' Level of ball rolling volume. Value between 0 and 1
Dim RampRollVolume : RampRollVolume = 0.5     ' Level of ramp rolling volume. Value between 0 and 1

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

  '-----Cabinet rails-----
  Dim rz
  v = Table1.Option("Cabinet Rails", 0, 1, 1, 1, 0, Array("Hide", "Show"))
  For Each rz in Rails: rz.visible = v: Next

    ' Sound volumes
    VolumeDial = Table1.Option("Mech Volume", 0, 1, 0.01, 0.8, 1)
    BallRollVolume = Table1.Option("Ball Roll Volume", 0, 1, 0.01, 0.5, 1)
  RampRollVolume = Table1.Option("Ramp Roll Volume", 0, 1, 0.01, 0.5, 1)

  ' Room brightness
  LightLevel = NightDay/100
  SetRoomBrightness LightLevel   'Uncomment this line for lightmapped tables.

  Dim v
  '-----GI Lamps Color-----
  v = Table1.Option("GI Lights", 0, 1, 1, 0, 0, Array("Classic", "White"))
  SetGIColor v

  '-----GI Lamps Brightness-----
  v = Table1.Option("GI Brightness", 0, 1, .1, .8, 1)
  GIBrightness v

  '-----Insert Brightness-----
  v = Table1.Option("Insert Brightness", 0, 1, .1, .7, 1)
  InsertBrightness v

    If eventId = 3 And dspTriggered Then dspTriggered = False : DisableStaticPreRendering = False : End If
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
  For each xx in cGI: xx.color = cArray(c): xx.colorfull = cArray(c): Next
  For each BL in vlmGI: BL.color = cArray(c): Next
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
  For each BL in vlmGI: BL.opacity = lvl * 100: Next

End Sub

Sub InsertBrightness(lvl)
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
  For each BL in vlmInserts: BL.opacity = lvl * 100: Next
  For each BL in whiteIns: BL.opacity = lvl * 75: Next

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
'  ZANI: Misc Animations
'******************************************************

Sub LeftFlipper_Animate
  dim a: a = LeftFlipper.CurrentAngle
  FlipperLSh.RotZ = a

    Dim v, BP
    v = 255.0 * (118.55 - LeftFlipper.CurrentAngle) / (118.55 - 71.05)

    For each BP in BP_LF
        BP.Rotz = a
        BP.visible = v < 94.5
    Next
    For each BP in BP_LFU
        BP.Rotz = a
        BP.visible = v >= 94.5
    Next

End Sub

Sub RightFlipper_Animate
  dim a: a = RightFlipper.CurrentAngle
  FlipperRSh.RotZ = a

    Dim v, BP
    v = 255.0 * (-118.55 - RightFlipper.CurrentAngle) / (-118.55 + 71.05)

    For each BP in BP_RF
        BP.Rotz = a
        BP.visible = v < 94.5
    Next
    For each BP in BP_RFU
        BP.Rotz = a
        BP.visible = v >= 94.5
    Next

End Sub

'--------drop targets-----------
Sub UpdateDropTargets
  dim BP, tz, rx, ry

  tz = BM_sw40.transz
  rx = BM_sw40.rotx
  ry = BM_sw40.roty
  For each BP in BP_sw40: BP.transz = tz: BP.rotx = rx: BP.roty = ry: Next

  tz = BM_sw50.transz
  rx = BM_sw50.rotx
  ry = BM_sw50.roty
  For each BP in BP_sw50: BP.transz = tz: BP.rotx = rx: BP.roty = ry: Next

  tz = BM_sw60.transz
  rx = BM_sw60.rotx
  ry = BM_sw60.roty
  For each BP in BP_sw60: BP.transz = tz: BP.rotx = rx: BP.roty = ry: Next

End Sub

'--------standup targets-----------
Sub UpdateStandupTargets
  dim BP, ty

    ty = BM_sw42.transy
  For each BP in BP_sw42 : BP.transy = ty: Next

    ty = BM_sw52.transy
  For each BP in BP_sw52: BP.transy = ty: Next

    ty = BM_sw62.transy
  For each BP in BP_sw62 : BP.transy = ty: Next

    ty = BM_sw72.transy
  For each BP in BP_sw72 : BP.transy = ty: Next

End Sub

'--------Gates-----------
Sub Gate1_Animate
  Dim spinangle:spinangle = Gate1.currentangle
  Dim BL : For Each BL in BP_Gate1 : BL.RotX = spinangle*-1: Next
End Sub

Sub Gate2_Animate
  Dim spinangle:spinangle = Gate2.currentangle
  Dim BL : For Each BL in BP_Gate2 : BL.RotX = spinangle*1: Next
End Sub

Sub Gate2Back_Animate
  Dim spinangle:spinangle = Gate2Back.currentangle
  Dim BL : For Each BL in BP_Gate2 : BL.RotX = spinangle*-1: Next
End Sub


Sub Gate3_Animate
  Dim spinangle:spinangle = Gate3.currentangle
  Dim BL : For Each BL in BP_Gate3 : BL.RotX = spinangle*1: Next
End Sub

Sub Gate3Back_Animate
  Dim spinangle:spinangle = Gate3Back.currentangle
  Dim BL : For Each BL in BP_Gate3 : BL.RotX = spinangle*-1: Next
End Sub

Sub sw71_Animate
  Dim spinangle:spinangle = sw71.currentangle
  Dim BL : For Each BL in BP_Gate4_Wire : BL.RotX = spinangle*-1: Next
End Sub

'--------spinners-----------
Sub sw74_Animate
  Dim spinangle:spinangle = sw74.currentangle
  Dim BL : For Each BL in BP_sw74_001 : BL.RotX = spinangle*-1: Next
End Sub

''--------Bumpers-----------
'' NOTE: set bumper object timer to around 150-175 in order to be able
'' to actually see the animaation, adjust to your liking
'
'Const PI = 3.1415926
Const SkirtTilt=5   'angle of skirt tilting in degrees

Function SkirtAX(bumper, bumperball)
  skirtAX=cos(skirtA(bumper,bumperball))*(SkirtTilt)    'x component of angle
  if (bumper.y<bumperball.y) then skirtAX=skirtAX*-1    'adjust for ball hit bottom half
End Function

Function SkirtAY(bumper, bumperball)
  skirtAY=sin(skirtA(bumper,bumperball))*(SkirtTilt)    'y component of angle
  if (bumper.x>bumperball.x) then skirtAY=skirtAY*-1    'adjust for ball hit left half
End Function

Function SkirtA(bumper, bumperball)
  dim hitx, hity, dx, dy
  hitx=bumperball.x
  hity=bumperball.y

  dy=Abs(hity-bumper.y)         'y offset ball at hit to center of bumper
  if dy=0 then dy=0.0000001
  dx=Abs(hitx-bumper.x)         'x offset ball at hit to center of bumper
  skirtA=(atn(dx/dy)) '/(PI/180)      'angle in radians to ball from center of Bumper1
End Function

Sub Bumper1_Hit
  vpmTimer.PulseSw(44)
  RandomSoundBumperMiddle bumper1
    me.timerenabled = 1

  Dim BP
  For Each BP in BP_Bumper_1_Socket
    BP.roty=skirtAY(me,Activeball)
    BP.rotx=skirtAX(me,Activeball)
  Next

  me.timerinterval = 150
  me.timerenabled=1
  DOF 103, DOFPulse
End Sub

sub Bumper1_timer
  Dim BP
  For Each BP in BP_Bumper_1_Socket
    BP.roty=0
    BP.rotx=0
  Next
  me.timerenabled=0
end sub

Sub Bumper1_Animate
  Dim z: z = Bumper1.CurrentRingOffset

  Dim BP
  For Each BP in BP_Bumper_1_Ring : BP.transz = z : Next
End Sub

Sub Bumper2_Hit
  vpmTimer.PulseSw(44)
  RandomSoundBumperMiddle bumper2
    me.timerenabled = 1

  Dim BP
  For Each BP in BP_Bumper_2_Socket
    BP.roty=skirtAY(me,Activeball)
    BP.rotx=skirtAX(me,Activeball)
  Next

  me.timerinterval = 150
  me.timerenabled=1
  DOF 104, DOFPulse
End Sub

sub Bumper2_timer
  Dim BP
  For Each BP in BP_Bumper_2_Socket
    BP.roty=0
    BP.rotx=0
  Next
  me.timerenabled=0
end sub

Sub Bumper2_Animate
  Dim z: z = Bumper2.CurrentRingOffset

  Dim BP
  For Each BP in BP_Bumper_2_Ring : BP.transz = z : Next
End Sub

Sub Bumper3_Hit
  vpmTimer.PulseSw(44)
  RandomSoundBumperTop bumper3
    me.timerenabled = 1

  Dim BP
  For Each BP in BP_Bumper_3_Socket
    BP.roty=skirtAY(me,Activeball)
    BP.rotx=skirtAX(me,Activeball)
  Next

  me.timerinterval = 150
  me.timerenabled=1
  DOF 105, DOFPulse
End Sub

sub Bumper3_timer
  Dim BP
  For Each BP in BP_Bumper_3_Socket
    BP.roty=0
    BP.rotx=0
  Next
  me.timerenabled=0
end sub

Sub Bumper3_Animate
  Dim z: z = Bumper3.CurrentRingOffset

  Dim BP
  For Each BP in BP_Bumper_3_Ring : BP.transz = z : Next
End Sub

Sub Bumper4_Hit
  vpmTimer.PulseSw(44)
  RandomSoundBumperMiddle bumper4
    me.timerenabled = 1

  Dim BP
  For Each BP in BP_Bumper_4_Socket
    BP.roty=skirtAY(me,Activeball)
    BP.rotx=skirtAX(me,Activeball)
  Next

  me.timerinterval = 150
  me.timerenabled=1
  DOF 106, DOFPulse
End Sub

sub Bumper4_timer
  Dim BP
  For Each BP in BP_Bumper_4_Socket
    BP.roty=0
    BP.rotx=0
  Next
  me.timerenabled=0
end sub

Sub Bumper4_Animate
  Dim z: z = Bumper4.CurrentRingOffset

  Dim BP
  For Each BP in BP_Bumper_4_Ring : BP.transz = z : Next
End Sub

'******************************************************
'   ZBBR: BALL BRIGHTNESS
'******************************************************

Const BallBrightness =  .5        'Ball brightness - Value between 0 and 1 (0=Dark ... 1=Bright)

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
Dim RoomBrightnessMtlArray: RoomBrightnessMtlArray = Array("VLM.Bake.Active","VLM.Bake.Solid","VLM.Bake.Spinner")

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

Sub Table1_KeyDown(ByVal keycode) '***What to do when a button is pressed***

  If KeyCode = PlungerKey Then
    Plunger.Pullback:SoundPlungerPull()
    TimerVRPlunger.Enabled = True
    TimerVRPlunger1.Enabled = False
    VR_Shooter.Y = -4
  End If

  If keycode = LeftFlipperKey Then
    VRFlipperButtonLeft.X = 2102 + 6
  End If

  If keycode = RightFlipperKey Then
    VRFlipperButtonRight.X = 2104 - 6
  End If

  If keycode = StartGameKey Then
    StartButton.y = 2793 - 5
    SoundStartButton
  End If

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


  If vpmKeyDown(keycode) Then Exit Sub
End Sub

Sub Table1_KeyUp(ByVal keycode)   '***What to do when a button is released***

  If Keycode = LeftFlipperKey Then
    FlipperDeActivate LeftFlipper, LFPress
    VRFlipperButtonLeft.X = 2102
  End If

  If Keycode = RightFlipperKey Then
    FlipperDeActivate RightFlipper, RFPress
    VRFlipperButtonRight.X = 2104
  End If

  If keycode = StartGameKey Then
    Controller.Switch(16) = 0
    StartButton.y = 2793
  End If

    If keycode = PlungerKey Then
    SoundPlungerReleaseNoBall
    Plunger.Fire
    TimerVRPlunger.Enabled = False
    TimerVRPlunger1.Enabled = True
    VR_Shooter.Y = -4
  End If

  If vpmKeyUp(keycode) Then Exit Sub
End Sub

'******************************************************
' ZSOL: Solenoids & Flashers
'******************************************************

'SolCallback(1) = "NotUsed"
SolCallback(2) = "ResetDrops"
'SolCallback(3) = "NotUsed"
'SolCallback(4) = "NotUsed"
SolCallback(5) = "KickerUpperRight"
'SolCallback(6) = "NotUsed"
'SolCallback(7) = "NotUsed"
SolCallback(8) = "SolKnocker"
SolCallback(9) = "SolOuthole"

SolCallback(sLRFlipper) = "SolRFlipper" 'Right Flipper
SolCallback(sLLFlipper) = "SolLFlipper" 'Left Flipper


'******************************************************
' ZDRN: Drain, Trough, and Ball Release
'******************************************************

'********************* TROUGH *************************

Sub sw67_Hit   : Controller.Switch(67) = 1 : RandomSoundDrain sw67 : End Sub
Sub sw67_UnHit : Controller.Switch(67) = 0 : End Sub

'*****************  DRAIN & RELEASE  ******************

Sub SolOuthole(Enabled)
  If Enabled = true Then
    RandomSoundBallRelease sw67
    sw67.kick 60, 15
  End If
End Sub

Sub SolKnocker(Enabled)
        If enabled Then
                KnockerSolenoid 'Add knocker position object
        End If
End Sub

Sub KickerUpperRight(Enabled)
        If enabled Then
      If VRMode Then
        'VR_Puck.visible = 1
        vrBGPuck.enabled = 1
      End If
'   Else
'     If VRMode Then
'       VR_Puck.visible = 0
'     End If
        End If
End Sub

'****************************
'   ZFLP: Flippers
'****************************
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


'************************************************************
' ZSLG: Slingshot Animations
'************************************************************
Dim LStep : LStep = 0 : LeftSlingShot.TimerEnabled = 1
Dim RStep : RStep = 0 : RightSlingShot.TimerEnabled = 1

Sub InitSlings
  Dim BP
  For Each BP in BP_Rsling1: BP.Visible = 0: Next
  For Each BP in BP_RSling2 : BP.Visible = 0: Next
  For Each BP in BP_LSling1 : BP.Visible = 0: Next
  For Each BP in BP_LSling2 : BP.Visible = 0: Next
End Sub

Sub RightSlingShot_Slingshot
  RS.VelocityCorrect(Activeball)
  vpmTimer.PulseSw(73)            'Sling Switch Number
  DOF 102,DOFPulse
  RandomSoundSlingshotLeft SlingR
  RStep = 0
  RightSlingShot_Timer
  RightSlingShot.TimerEnabled = 1
  RightSlingShot.TimerInterval = 17
End Sub

Sub RightSlingShot_Timer
  Dim BL
  Dim x1, x2, y: x1 = True:x2 = False:y = -20
    Select Case RStep
        Case 2:x1 = False:x2 = True:y = -10
        Case 3:x1 = False:x2 = False:y = 0:RightSlingShot.TimerEnabled = 0
    End Select

  For Each BL in BP_RSling1 : BL.Visible = x1: Next
  For Each BL in BP_RSling2 : BL.Visible = x2: Next
  SlingR.TransY = 0

    RStep = RStep + 1
End Sub

Sub LeftSlingShot_Slingshot
  LS.VelocityCorrect(Activeball)
  vpmTimer.PulseSw(73)            'Sling Switch Number
  DOF 101,DOFPulse
  RandomSoundSlingshotRight SlingL
  LStep = 0
  LeftSlingShot_Timer
  LeftSlingShot.TimerEnabled = 1
  LeftSlingShot.TimerInterval = 17
End Sub

Sub LeftSlingShot_Timer
  Dim BL
  Dim x1, x2, y: x1 = True:x2 = False:y = -20
    Select Case LStep
        Case 3:x1 = False:x2 = True:y = -10
        Case 4:x1 = False:x2 = False:y = 0:LeftSlingShot.TimerEnabled = 0
    End Select

  For Each BL in BP_Lsling1 : BL.Visible = x1: Next
  For Each BL in BP_LSling2 : BL.Visible = x2: Next
  SlingL.TransY = 0

    LStep = LStep + 1
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
  AddSlingsPt 0, 0.00,    -3
  AddSlingsPt 1, 0.30,    -5
  AddSlingsPt 2, 0.40,  -30
  AddSlingsPt 3, 0.60,  30
  AddSlingsPt 4, 0.70,  5
  AddSlingsPt 5, 1.00,  3
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

'************************ Rollovers *************************
Sub sw41_Hit(): Controller.Switch(41) = 1: AnimateWire BP_sw41, 1:End Sub
Sub sw41_UnHit: Controller.Switch(41) = 0: AnimateWire BP_sw41, 0:End Sub
Sub sw51_Hit(): Controller.Switch(51) = 1: AnimateWire BP_sw51, 1:End Sub
Sub sw51_UnHit: Controller.Switch(51) = 0: AnimateWire BP_sw51, 0:End Sub
Sub sw61_Hit(): Controller.Switch(61) = 1: AnimateWire BP_sw61, 1:End Sub
Sub sw61_UnHit: Controller.Switch(61) = 0: AnimateWire BP_sw61, 0:End Sub

Sub sw63_Hit(): Controller.Switch(63) = 1: AnimateWire BP_sw63, 1:DOF 107, DOFOn:End Sub
Sub sw63_UnHit: Controller.Switch(63) = 0: AnimateWire BP_sw63, 0:DOF 107, DOFOff:End Sub
Sub sw63a_Hit(): Controller.Switch(63) = 1: AnimateWire BP_sw63a, 1:DOF 108, DOFOff:End Sub
Sub sw63a_UnHit: Controller.Switch(63) = 0: AnimateWire BP_sw63a, 0:DOF 108, DOFOff:End Sub
Sub sw43_Hit(): Controller.Switch(43) = 1: AnimateWire BP_sw43, 1:End Sub
Sub sw43_UnHit: Controller.Switch(43) = 0: AnimateWire BP_sw43, 0:End Sub
Sub sw53_Hit(): Controller.Switch(53) = 1: AnimateWire BP_sw53, 1:End Sub
Sub sw53_UnHit: Controller.Switch(53) = 0: AnimateWire BP_sw53, 0:End Sub

Sub AnimateWire(group, action) ' Action = 1 - to drop, 0 to raise)
  Dim BP
  If action = 1 Then
    For Each BP in group : BP.transz = -13 : Next
  Else
    For Each BP in group : BP.transz = 0 : Next
  End If
End Sub

'********************* Gates and Spinners **********************
Sub sw74_Spin():SoundSpinner sw74:vpmTimer.PulseSw (74):End Sub
Sub sw71_Hit(): vpmTimer.PulseSw (71): End Sub

'********************* Standup Targets **********************
Sub sw42_hit: STHit 42: End Sub
Sub sw52_hit: STHit 52: End Sub
Sub sw62_hit: STHit 62: End Sub
Sub sw72_hit: STHit 72: End Sub

Sub sw70_hit: STHit 70: vpmTimer.PulseSw 70: End Sub


'********************** Drop Targets ************************
Sub Sw40_Hit: DTHit 40: TargetBouncer Activeball, 1.5: End Sub
Sub Sw50_Hit: DTHit 50: TargetBouncer Activeball, 1.5: End Sub
Sub Sw60_Hit: DTHit 60: TargetBouncer Activeball, 1.5: End Sub

Sub ResetDrops(enabled)
  if enabled then
    DTRaise 40
    RandomSoundDropTargetReset BM_sw40
    DTRaise 50
    RandomSoundDropTargetReset BM_sw50
    DTRaise 60
    RandomSoundDropTargetReset BM_sw60
    DOF 201, DOFPulse
  end if
End Sub


'******************************************************
' ZBRL: BALL ROLLING AND DROP SOUNDS
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
  Dim b, gBOT
  gBOT = GetBalls

  ' stop the sound of deleted balls
  For b = UBound(gBOT) + 1 to tnob - 1
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
  Next
End Sub


'******************************************************
'****  END BALL ROLLING AND DROP SOUNDS
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
'Const EOSReturn = 0.055  'EM's
Const EOSReturn = 0.045  'late 70's to mid 80's
'Const EOSReturn = 0.035  'mid 80's to early 90's
'Const EOSReturn = 0.025  'mid 90's and later

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
Dim DT40, DT50, DT60

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

Set DT40 = (new DropTarget)(sw40, sw40a, BM_sw40, 40, 0, false)
Set DT50 = (new DropTarget)(sw50, sw50a, BM_sw50, 50, 0, false)
Set DT60 = (new DropTarget)(sw60, sw60a, BM_sw60, 60, 0, false)

Dim DTArray
DTArray = Array(DT40, DT50, DT60)

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

Const DTMass = 0.1 'Mass of the Drop Target (between 0 and 1), higher values provide more resistance


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
Dim ST42, ST52, ST62, ST72

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

Set ST42 = (new StandupTarget)(sw42, BM_sw42, 42, 0)
Set ST52 = (new StandupTarget)(sw52, BM_sw52, 52, 0)
Set ST62 = (new StandupTarget)(sw62, BM_sw62, 62, 0)
Set ST72 = (new StandupTarget)(sw72, BM_sw72, 72, 0)

'Add all the Stand-up Target Arrays to Stand-up Target Animation Array
' STAnimationArray = Array(ST1, ST2, ....)
Dim STArray
STArray = Array(ST42, ST52, ST62, ST72)

'Configure the behavior of Stand-up Targets
Const STAnimStep =  1.5     'vpunits per animation step (control return to Start)
Const STMaxOffset = 9       'max vp units target moves when hit

Const STMass = 0.1        'Mass of the Stand-up Target (between 0 and 1), higher values provide more resistance

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

'Add targets or posts to the TargetBounce collection if you want to activate the targetbouncer code from them
Sub TargetBounce_Hit(idx)
  TargetBouncer ActiveBall, 1
End Sub

'****************************************************************
'****  END VPW TargetBouncer for targets and posts by Iaakki, Wrd1972, Apophis
'****************************************************************

'******************************************************
'****  ZGIU:  GI Control
'******************************************************

'**** Global variable to hold current GI light intensity. Not Mandatory, but may help in other lighting subs and timers
'**** This value is updated always when GI state is being updated
dim gilvl

'**** These are just debug commands. They emit same calls as what SolCallback will do. You may use these from debugger.
Sub GIOn  : SetRelayGI 0: End Sub
Sub GIOff : SetRelayGI 1: End Sub


'**** SetRelayGI is called from SolCallback.
sub SetRelayGI(aLvl)
' debug.print "SetRelayGI value: " & aLvl

  'Some tables have this solenoid reversed, i.e. Sega and Data East GI lights are off when GI relay is on.
  'However, this should be already compensated for in VPinMame 3.6

  ' Update the state for each GI light. The state will be a float value between 0 and 1.
  Dim bulb: For Each bulb in GI: bulb.State = aLvl: Next

  ' If the GI has an associated Relay sound, this can be played
  If aLvl >= 0.5 And gilvl < 0.5 Then
    Sound_GI_Relay 1, Bumper1 'Note: Bumper1 is just used for sound positioning. Can be anywhere that makes sense.
  ElseIf aLvl <= 0.4 And gilvl > 0.4 Then
    Sound_GI_Relay 0, Bumper1
  End If

  'You may add any other GI related effects here. Like if you want to make some toy to appear more bright, set it like this:
  'Primitive001.blenddisablelighting = 1.2 * aLvl + 0.2 'This will result to DL brightness between 0.2 - 1.4 for ON/OFF states

  gilvl = aLvl    'Storing the latest GI fading state into global variable, so one can use it elsewhere too.
End Sub


'**** Use these for GI strings and stepped GI, and comment out the SetRelayGI
'**** This example table uses Relay for GI control, we don't need these at all

'Set GICallback  = GetRef("GIUpdates")    'use this for non-modulated GI
'Set GICallback2 = GetRef("GIUpdates2")   'use this for stepped/modulated GI

'GIupdates2 is called always when some event happens to GI channel.
Sub GIUpdates2(aNr, aLvl)
' debug.print "GIUpdates2 nr: " & aNr & " value: " & aLvl

  Select Case aNr 'Strings are selected here

    Case 0:  'GI String 0

      ' Update the state for each GI light. The state will be a float value between 0 and 1.
      Dim bulb: For Each bulb in GI: bulb.State = aLvl: Next

      ' If the GI has an associated Relay sound, this can be played
      If aLvl >= 0.5 And gilvl < 0.5 Then
        Sound_GI_Relay 1, Bumper1 'Note: Bumper1 is just used for sound positioning. Can be anywhere that makes sense.
      ElseIf aLvl <= 0.4 And gilvl > 0.4 Then
        Sound_GI_Relay 0, Bumper1
      End If

      'You may add any other GI related effects here. Like if you want to make some toy to appear more bright, set it like this:
      'Primitive001.blenddisablelighting = 1.2 * aLvl + 0.2 'This will result to DL brightness between 0.2 - 1.4 for ON/OFF states

      gilvl = aLvl    'Storing the latest GI fading state into global variable, so one can use it elsewhere too.

    Case 1:  'GI String 1

    Case 2:  'GI String 2

  End Select

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
Dim objBallShadow(2)

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

'***************************************************************
' ZVRR: VR Code
'***************************************************************

'*********************
' VR Plunger code
'*********************

Sub TimerVRPlunger_Timer
  If VR_Shooter.Y < 86 then
      VR_Shooter.Y = VR_Shooter.Y + 5
  End If
End Sub

Sub TimerVRPlunger1_Timer
  VR_Shooter.Y = -4 + (5* Plunger.Position) -20
End Sub

' ***************** VR Clock — thanks Rascal ******************
Dim CurrentMinute   ' for VR clock
Sub ClockTimer_Timer()
    Pminutes.RotAndTra2 = (Minute(Now()) + (Second(Now()) / 100)) * 6
    Phours.RotAndTra2   = Hour(Now()) * 30 + (Minute(Now()) / 2)
    Pseconds.RotAndTra2 = Second(Now()) * 6
    CurrentMinute = Minute(Now())
End Sub

' ***************** Beer Bubbles ******************************
Sub BeerTimer_Timer
    Dim beerspeed : beerspeed = 0.75
    Randomize(21)
    BeerBubble1.z = BeerBubble1.z + Rnd(1) * 0.5 * beerspeed
    If BeerBubble1.z > -771 Then BeerBubble1.z = -955
    BeerBubble2.z = BeerBubble2.z + Rnd(1) * 1 * beerspeed
    If BeerBubble2.z > -768 Then BeerBubble2.z = -955
    BeerBubble3.z = BeerBubble3.z + Rnd(1) * 1 * beerspeed
    If BeerBubble3.z > -768 Then BeerBubble3.z = -955
    BeerBubble4.z = BeerBubble4.z + Rnd(1) * 0.75 * beerspeed
    If BeerBubble4.z > -774 Then BeerBubble4.z = -955
    BeerBubble5.z = BeerBubble5.z + Rnd(1) * 1 * beerspeed
    If BeerBubble5.z > -771 Then BeerBubble5.z = -955
    BeerBubble6.z = BeerBubble6.z + Rnd(1) * 1 * beerspeed
    If BeerBubble6.z > -774 Then BeerBubble6.z = -955
    BeerBubble7.z = BeerBubble7.z + Rnd(1) * 0.8 * beerspeed
    If BeerBubble7.z > -768 Then BeerBubble7.z = -955
    BeerBubble8.z = BeerBubble8.z + Rnd(1) * 1 * beerspeed
    If BeerBubble8.z > -771 Then BeerBubble8.z = -955
End Sub

'******************************************************
' VR Backglass Digital Display
'******************************************************

Dim xoff,yoff1, yoff2, yoff3, yoff4, yoff5, zoff,xrot,zscale, xcen,ycen

Sub Setup_Backglass()

  xoff = -20
  yoff1 = 35 ' this is where you adjust the forward/backward position for player 1 score
  yoff2 = 35 ' this is where you adjust the forward/backward position for player 2 score
  yoff3 = 35 ' this is where you adjust the forward/backward position for player 3 score
  yoff4 = 35 ' this is where you adjust the forward/backward position for player 4 score
  yoff5 = 35 ' this is where you adjust the forward/backward position for credits and ball in play
  zoff = 638
  xrot = -90

  CenterVRDigits()

End Sub

Sub CenterVRDigits()

Dim xoff,yoff,zoff,xrot,zscale, xcen,ycen, ix, xx, yy, xobj

  xoff =480
  yoff = 30
  zoff = 180
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

If VRMode Then
  InitDigits
End If
'vr bg puck animations

dim vrpuck1st: vrpuck1st = 1

Sub vrBGPuck_timer()

  Select Case vrpuck1st

    Case 1: VR_Puck.imageA = "puck"
    Case 3: VR_Puck.imageA = "puck2"
    Case 6: VR_Puck.imageA = "puck3"
    Case 10: VR_Puck.imageA = "puck4"
    Case 15: VR_Puck.imageA = "puck5"
    Case 21: VR_Puck.imageA = "puck6"
    Case 28: VR_Puck.imageA = "puck7"
    Case 36: VR_Puck.imageA = "puck8"
    Case 45: VR_Puck.imageA = "puck9"

  End Select

  vrpuck1st = vrpuck1st + 1

  if vrpuck1st = 55 Then

    vrBGPuck.enabled = 0
    vrpuck1st = 1
    VR_Puck.imageA = "puck"
  end If
end sub

'***************************************************************
' ZDTD: Desktop Digits
'***************************************************************

Dim Digits(31)
Dim Patterns(11)
Dim Patterns2(11)
Dim PatternsA(11)
Dim Patterns2A(11)

Patterns(0) = 0     'empty
Patterns(1) = 63    '0
Patterns(2) = 768   '1
Patterns(3) = 91    '2
Patterns(4) = 79    '3
Patterns(5) = 102   '4
Patterns(6) = 109   '5
Patterns(7) = 125   '6
Patterns(8) = 7     '7
Patterns(9) = 127   '8
Patterns(10) = 111  '9

Patterns2(0) = 128  ' empty + comma (0 + 128)
Patterns2(1) = 191  ' 0 + comma (63 + 128)
Patterns2(2) = 896  ' 1 + comma (768 + 128)
Patterns2(3) = 219  ' 2 + comma (91 + 128)
Patterns2(4) = 207  ' 3 + comma (79 + 128)
Patterns2(5) = 230  ' 4 + comma (102 + 128)
Patterns2(6) = 237  ' 5 + comma (109 + 128)
Patterns2(7) = 252
Patterns2(8) = 135  ' 7 + comma (7 + 128)
Patterns2(9) = 255  ' 8 + comma (127 + 128)
Patterns2(10)= 231

'Assign 7-digit output to reels
Set Digits(0) = a0
Set Digits(1) = a1
Set Digits(2) = a2
Set Digits(3) = a3
Set Digits(4) = a4
Set Digits(5) = a5
Set Digits(6) = a6

Set Digits(7) = b0
Set Digits(8) = b1
Set Digits(9) = b2
Set Digits(10) = b3
Set Digits(11) = b4
Set Digits(12) = b5
Set Digits(13) = b6

Set Digits(14) = c0
Set Digits(15) = c1
Set Digits(16) = c2
Set Digits(17) = c3
Set Digits(18) = c4
Set Digits(19) = c5
Set Digits(20) = c6

Set Digits(21) = d0
Set Digits(22) = d1
Set Digits(23) = d2
Set Digits(24) = d3
Set Digits(25) = d4
Set Digits(26) = d5
Set Digits(27) = d6

Set Digits(28) = e0
Set Digits(29) = e1
Set Digits(30) = e2
Set Digits(31) = e3

PatternsA(0) = 0     'empty
PatternsA(1) = 63    '0
PatternsA(2) = 6     '1  <-- Standard 7-segment 1 (b+c = 2+4)
PatternsA(3) = 91    '2
PatternsA(4) = 79    '3
PatternsA(5) = 102   '4
PatternsA(6) = 109   '5
PatternsA(7) = 124   '6
PatternsA(8) = 7     '7
PatternsA(9) = 127   '8
PatternsA(10) = 103  '9

' SINGLE, UNIFIED UPDATE LOOP
Sub UpdateLeds
    On Error Resume Next
    Dim ChgLED, ii, jj, chg, stat, ledNum

    ' Fetch changes from PinMAME exactly ONCE
    ChgLED = Controller.ChangedLEDs(&HFF, &HFFFF)

    If Not IsEmpty(ChgLED) Then
        For ii = 0 To UBound(ChgLED)
            ledNum = chgLED(ii, 0)
            chg = chgLED(ii, 1)
            stat = chgLED(ii, 2)

            If ledNum <= 27 Then
                ' Processing Main Player Displays (Supports Commas and Centered 1)
                For jj = 0 To 10
                    If stat = Patterns(jj) OR stat = Patterns2(jj) Then
                        Digits(ledNum).SetValue jj
                    End If
                Next

            ElseIf ledNum >= 28 And ledNum <= 31 Then
                ' Processing Ball & Credit Displays (Strictly 7-Segment, No Commas)
                For jj = 0 To 10
                    If stat = PatternsA(jj) Then
                        Digits(ledNum).SetValue jj
                    End If
                Next
            End If

        Next
    End If
End Sub

Dim VRThings
Dim BGThings
Dim DTThings
'Desktop Mode
if VRMode = 0 and cab_mode = 0 Then
  'for each VRThings in VRStuff:VRThings.visible = 0:Next
  For each DTThings in DTDigits: DTThings.Visible = 1: Next

' Cabinet Mode
Elseif VRMode = 0 and cab_mode = 1 Then
  'for each VRThings in VRStuff:VRThings.visible = 0:Next
  For each DTThings in DTDigits: DTThings.Visible = 0: Next


' VR Mode
Else
  'for each VRThings in VRStuff:VRThings.visible = 1:Next
  For each DTThings in DTDigits: DTThings.Visible = 0: Next
End If

'***************************************************************
' ZVLM: VLM Arrays
'***************************************************************

' VLM  Arrays - Start
' Arrays per baked part
Dim BP_Bumper_1_Ring: BP_Bumper_1_Ring=Array(BM_Bumper_1_Ring, LM_GI_BumperLight1_Bumper_1_Rin, LM_GI_Light54_Bumper_1_Ring, LM_GI_Light55_Bumper_1_Ring, LM_Inserts_l23_Bumper_1_Ring)
Dim BP_Bumper_1_Socket: BP_Bumper_1_Socket=Array(BM_Bumper_1_Socket, LM_GI_BumperLight1_Bumper_1_Soc, LM_GI_Light54_Bumper_1_Socket, LM_GI_Light55_Bumper_1_Socket, LM_Inserts_l23_Bumper_1_Socket)
Dim BP_Bumper_2_Ring: BP_Bumper_2_Ring=Array(BM_Bumper_2_Ring, LM_GI_BumperLight2_Bumper_2_Rin, LM_GI_Light53_Bumper_2_Ring, LM_GI_Light58_Bumper_2_Ring, LM_Inserts_l24_Bumper_2_Ring, LM_Inserts_l51_Bumper_2_Ring)
Dim BP_Bumper_2_Socket: BP_Bumper_2_Socket=Array(BM_Bumper_2_Socket, LM_GI_BumperLight2_Bumper_2_Soc, LM_GI_Light53_Bumper_2_Socket, LM_GI_Light58_Bumper_2_Socket, LM_Inserts_l51_Bumper_2_Socket)
Dim BP_Bumper_3_Ring: BP_Bumper_3_Ring=Array(BM_Bumper_3_Ring, LM_GI_BumperLight2_Bumper_3_Rin, LM_GI_BumperLight3_Bumper_3_Rin, LM_GI_Light56_Bumper_3_Ring, LM_GI_Light58_Bumper_3_Ring, LM_GI_Light69_Bumper_3_Ring, LM_Inserts_l24_Bumper_3_Ring)
Dim BP_Bumper_3_Socket: BP_Bumper_3_Socket=Array(BM_Bumper_3_Socket, LM_GI_BumperLight2_Bumper_3_Soc, LM_GI_BumperLight3_Bumper_3_Soc, LM_GI_Light56_Bumper_3_Socket, LM_GI_Light58_Bumper_3_Socket, LM_GI_Light69_Bumper_3_Socket, LM_Inserts_l24_Bumper_3_Socket)
Dim BP_Bumper_4_Ring: BP_Bumper_4_Ring=Array(BM_Bumper_4_Ring, LM_GI_BumperLight4_Bumper_4_Rin, LM_GI_Light52_Bumper_4_Ring, LM_GI_Light53_Bumper_4_Ring)
Dim BP_Bumper_4_Socket: BP_Bumper_4_Socket=Array(BM_Bumper_4_Socket, LM_GI_BumperLight4_Bumper_4_Soc, LM_GI_Light52_Bumper_4_Socket, LM_GI_Light53_Bumper_4_Socket)
Dim BP_Gate1: BP_Gate1=Array(BM_Gate1, LM_GI_Light68_Gate1)
Dim BP_Gate2: BP_Gate2=Array(BM_Gate2, LM_GI_Light59_Gate2, LM_GI_Light68_Gate2, LM_GI_Light70_Gate2, LM_GI_Light71_Gate2, LM_Inserts_l47_Gate2)
Dim BP_Gate3: BP_Gate3=Array(BM_Gate3, LM_GI_Light57_Gate3, LM_GI_Light60_Gate3, LM_GI_Light68_Gate3, LM_GI_Light71_Gate3, LM_Inserts_l26_Gate3, LM_Inserts_l45_Gate3)
Dim BP_Gate4_Wire: BP_Gate4_Wire=Array(BM_Gate4_Wire, LM_GI_Light57_Gate4_Wire, LM_GI_Light58_Gate4_Wire, LM_GI_Light60_Gate4_Wire, LM_GI_Light61_Gate4_Wire, LM_GI_Light69_Gate4_Wire, LM_Inserts_l45_Gate4_Wire, LM_Inserts_l7_Gate4_Wire)
Dim BP_LF: BP_LF=Array(BM_LF, LM_GI_Light11_LF, LM_Inserts_l3_LF)
Dim BP_LFU: BP_LFU=Array(BM_LFU, LM_GI_Light11_LFU, LM_Inserts_l33_LFU)
Dim BP_Layer1: BP_Layer1=Array(BM_Layer1, LM_GI_Light1_Layer1, LM_GI_Light53_Layer1, LM_GI_Light56_Layer1, LM_GI_Light57_Layer1, LM_GI_Light58_Layer1, LM_GI_Light59_Layer1, LM_GI_Light60_Layer1, LM_GI_Light61_Layer1, LM_GI_Light68_Layer1, LM_GI_Light69_Layer1, LM_GI_Light70_Layer1, LM_GI_Light71_Layer1, LM_Inserts_l21_Layer1, LM_Inserts_l26_Layer1, LM_Inserts_l27_Layer1, LM_Inserts_l28_Layer1, LM_Inserts_l45_Layer1, LM_Inserts_l46_Layer1, LM_Inserts_l47_Layer1, LM_Inserts_l5_Layer1, LM_Inserts_l6_Layer1, LM_Inserts_l7_Layer1)
Dim BP_Lsling: BP_Lsling=Array(BM_Lsling, LM_GI_Light10_Lsling, LM_GI_Light11_Lsling, LM_Inserts_l44_Lsling)
Dim BP_Lsling1: BP_Lsling1=Array(BM_Lsling1, LM_GI_Light10_Lsling1, LM_GI_Light11_Lsling1, LM_Inserts_l44_Lsling1)
Dim BP_Lsling2: BP_Lsling2=Array(BM_Lsling2, LM_GI_Light10_Lsling2, LM_GI_Light11_Lsling2, LM_Inserts_l44_Lsling2)
Dim BP_Parts: BP_Parts=Array(BM_Parts, LM_GI_BumperLight1_Parts, LM_GI_BumperLight2_Parts, LM_GI_BumperLight3_Parts, LM_GI_BumperLight4_Parts, LM_GI_Light1_Parts, LM_GI_Light10_Parts, LM_GI_Light11_Parts, LM_GI_Light52_Parts, LM_GI_Light53_Parts, LM_GI_Light54_Parts, LM_GI_Light55_Parts, LM_GI_Light56_Parts, LM_GI_Light57_Parts, LM_GI_Light58_Parts, LM_GI_Light59_Parts, LM_GI_Light60_Parts, LM_GI_Light61_Parts, LM_GI_Light68_Parts, LM_GI_Light69_Parts, LM_GI_Light70_Parts, LM_GI_Light71_Parts, LM_Inserts_l12_Parts, LM_Inserts_l21_Parts, LM_Inserts_l22_Parts, LM_Inserts_l23_Parts, LM_Inserts_l24_Parts, LM_Inserts_l25_Parts, LM_Inserts_l26_Parts, LM_Inserts_l27_Parts, LM_Inserts_l28_Parts, LM_Inserts_l4_Parts, LM_Inserts_l44_Parts, LM_Inserts_l45_Parts, LM_Inserts_l46_Parts, LM_Inserts_l47_Parts, LM_Inserts_l48_Parts, LM_Inserts_l5_Parts, LM_Inserts_l51_Parts, LM_Inserts_l6_Parts, LM_Inserts_l7_Parts)
Dim BP_Plastics: BP_Plastics=Array(BM_Plastics, LM_GI_BumperLight1_Plastics, LM_GI_BumperLight2_Plastics, LM_GI_BumperLight3_Plastics, LM_GI_BumperLight4_Plastics, LM_GI_Light10_Plastics, LM_GI_Light11_Plastics, LM_GI_Light52_Plastics, LM_GI_Light53_Plastics, LM_GI_Light54_Plastics, LM_GI_Light55_Plastics, LM_GI_Light56_Plastics, LM_GI_Light57_Plastics, LM_GI_Light58_Plastics, LM_GI_Light59_Plastics, LM_GI_Light60_Plastics, LM_GI_Light61_Plastics, LM_GI_Light68_Plastics, LM_GI_Light69_Plastics, LM_GI_Light70_Plastics, LM_GI_Light71_Plastics, LM_Inserts_l12_Plastics, LM_Inserts_l22_Plastics, LM_Inserts_l23_Plastics, LM_Inserts_l24_Plastics, LM_Inserts_l25_Plastics, LM_Inserts_l28_Plastics, LM_Inserts_l4_Plastics, LM_Inserts_l44_Plastics, LM_Inserts_l47_Plastics, LM_Inserts_l48_Plastics, LM_Inserts_l5_Plastics, LM_Inserts_l6_Plastics, LM_Inserts_l7_Plastics)
Dim BP_Playfield: BP_Playfield=Array(BM_Playfield, BM_Playfield_001, LM_GI_BumperLight1_Playfield, LM_GI_BumperLight2_Playfield, LM_GI_BumperLight3_Playfield, LM_GI_BumperLight4_Playfield, LM_GI_Light1_Playfield, LM_GI_Light10_Playfield, LM_GI_Light11_Playfield, LM_GI_Light52_Playfield, LM_GI_Light53_Playfield, LM_GI_Light54_Playfield, LM_GI_Light55_Playfield, LM_GI_Light56_Playfield, LM_GI_Light57_Playfield, LM_GI_Light58_Playfield, LM_GI_Light59_Playfield, LM_GI_Light60_Playfield, LM_GI_Light61_Playfield, LM_GI_Light68_Playfield, LM_GI_Light69_Playfield, LM_GI_Light70_Playfield, LM_GI_Light71_Playfield, LM_Inserts_l12_Playfield, LM_Inserts_l13_Playfield, LM_Inserts_l14_Playfield, LM_Inserts_l15_Playfield, LM_Inserts_l16_Playfield, LM_Inserts_l17_Playfield, LM_Inserts_l18_Playfield, LM_Inserts_l19_Playfield, LM_Inserts_l20_Playfield, LM_Inserts_l21_Playfield, LM_Inserts_l22_Playfield, LM_Inserts_l23_Playfield, LM_Inserts_l24_Playfield, LM_Inserts_l25_Playfield, LM_Inserts_l26_Playfield, _
  LM_Inserts_l27_Playfield, LM_Inserts_l28_Playfield, LM_Inserts_l29_Playfield, LM_Inserts_l3_Playfield, LM_Inserts_l30_Playfield, LM_Inserts_l31_Playfield, LM_Inserts_l32_Playfield, LM_Inserts_l33_Playfield, LM_Inserts_l34_Playfield, LM_Inserts_l35_Playfield, LM_Inserts_l36_Playfield, LM_Inserts_l37_Playfield, LM_Inserts_l38_Playfield, LM_Inserts_l39_Playfield, LM_Inserts_l4_Playfield, LM_Inserts_l40_Playfield, LM_Inserts_l41_Playfield, LM_Inserts_l42_Playfield, LM_Inserts_l43_Playfield, LM_Inserts_l44_Playfield, LM_Inserts_l45_Playfield, LM_Inserts_l46_Playfield, LM_Inserts_l47_Playfield, LM_Inserts_l48_Playfield, LM_Inserts_l49_Playfield, LM_Inserts_l5_Playfield, LM_Inserts_l50_Playfield, LM_Inserts_l51_Playfield, LM_Inserts_l6_Playfield, LM_Inserts_l7_Playfield)
Dim BP_RF: BP_RF=Array(BM_RF, LM_GI_Light10_RF)
Dim BP_RFU: BP_RFU=Array(BM_RFU, LM_GI_Light10_RFU, LM_Inserts_l30_RFU)
Dim BP_Rsling: BP_Rsling=Array(BM_Rsling, LM_GI_Light10_Rsling, LM_GI_Light52_Rsling, LM_Inserts_l48_Rsling)
Dim BP_Rsling1: BP_Rsling1=Array(BM_Rsling1, LM_GI_Light10_Rsling1, LM_GI_Light52_Rsling1, LM_Inserts_l48_Rsling1)
Dim BP_Rsling2: BP_Rsling2=Array(BM_Rsling2, LM_GI_Light10_Rsling2, LM_GI_Light52_Rsling2, LM_Inserts_l48_Rsling2)
Dim BP_SpinnerRod: BP_SpinnerRod=Array(BM_SpinnerRod, LM_GI_Light53_SpinnerRod, LM_Inserts_l12_SpinnerRod)
Dim BP_UnderPF: BP_UnderPF=Array(BM_UnderPF, LM_GI_BumperLight1_UnderPF, LM_GI_BumperLight2_UnderPF, LM_GI_BumperLight3_UnderPF, LM_GI_BumperLight4_UnderPF, LM_GI_Light1_UnderPF, LM_GI_Light10_UnderPF, LM_GI_Light11_UnderPF, LM_GI_Light52_UnderPF, LM_GI_Light53_UnderPF, LM_GI_Light54_UnderPF, LM_GI_Light55_UnderPF, LM_GI_Light56_UnderPF, LM_GI_Light57_UnderPF, LM_GI_Light58_UnderPF, LM_GI_Light59_UnderPF, LM_GI_Light60_UnderPF, LM_GI_Light61_UnderPF, LM_GI_Light68_UnderPF, LM_GI_Light69_UnderPF, LM_GI_Light70_UnderPF, LM_GI_Light71_UnderPF, LM_Inserts_l12_UnderPF, LM_Inserts_l13_UnderPF, LM_Inserts_l14_UnderPF, LM_Inserts_l15_UnderPF, LM_Inserts_l16_UnderPF, LM_Inserts_l17_UnderPF, LM_Inserts_l18_UnderPF, LM_Inserts_l19_UnderPF, LM_Inserts_l20_UnderPF, LM_Inserts_l21_UnderPF, LM_Inserts_l22_UnderPF, LM_Inserts_l23_UnderPF, LM_Inserts_l24_UnderPF, LM_Inserts_l25_UnderPF, LM_Inserts_l26_UnderPF, LM_Inserts_l27_UnderPF, LM_Inserts_l28_UnderPF, LM_Inserts_l29_UnderPF, LM_Inserts_l3_UnderPF, _
  LM_Inserts_l30_UnderPF, LM_Inserts_l31_UnderPF, LM_Inserts_l32_UnderPF, LM_Inserts_l33_UnderPF, LM_Inserts_l34_UnderPF, LM_Inserts_l35_UnderPF, LM_Inserts_l36_UnderPF, LM_Inserts_l37_UnderPF, LM_Inserts_l38_UnderPF, LM_Inserts_l39_UnderPF, LM_Inserts_l4_UnderPF, LM_Inserts_l40_UnderPF, LM_Inserts_l41_UnderPF, LM_Inserts_l42_UnderPF, LM_Inserts_l43_UnderPF, LM_Inserts_l44_UnderPF, LM_Inserts_l45_UnderPF, LM_Inserts_l46_UnderPF, LM_Inserts_l47_UnderPF, LM_Inserts_l48_UnderPF, LM_Inserts_l49_UnderPF, LM_Inserts_l5_UnderPF, LM_Inserts_l50_UnderPF, LM_Inserts_l51_UnderPF, LM_Inserts_l6_UnderPF, LM_Inserts_l7_UnderPF)
Dim BP_sw40: BP_sw40=Array(BM_sw40, LM_GI_Light57_sw40, LM_Inserts_l21_sw40, LM_Inserts_l5_sw40)
Dim BP_sw41: BP_sw41=Array(BM_sw41, LM_GI_Light58_sw41, LM_GI_Light69_sw41, LM_GI_Light70_sw41, LM_Inserts_l45_sw41)
Dim BP_sw42: BP_sw42=Array(BM_sw42, LM_GI_Light54_sw42, LM_Inserts_l22_sw42)
Dim BP_sw43: BP_sw43=Array(BM_sw43, LM_GI_Light11_sw43)
Dim BP_sw50: BP_sw50=Array(BM_sw50, LM_Inserts_l21_sw50)
Dim BP_sw51: BP_sw51=Array(BM_sw51, LM_GI_Light70_sw51, LM_GI_Light71_sw51, LM_Inserts_l46_sw51)
Dim BP_sw52: BP_sw52=Array(BM_sw52, LM_GI_BumperLight1_sw52, LM_GI_Light54_sw52, LM_Inserts_l22_sw52, LM_Inserts_l23_sw52)
Dim BP_sw53: BP_sw53=Array(BM_sw53, LM_GI_Light10_sw53)
Dim BP_sw60: BP_sw60=Array(BM_sw60, LM_GI_Light56_sw60, LM_GI_Light58_sw60, LM_GI_Light69_sw60, LM_Inserts_l6_sw60)
Dim BP_sw61: BP_sw61=Array(BM_sw61, LM_GI_Light59_sw61, LM_GI_Light68_sw61, LM_GI_Light71_sw61)
Dim BP_sw62: BP_sw62=Array(BM_sw62, LM_GI_BumperLight3_sw62, LM_GI_Light57_sw62, LM_GI_Light58_sw62, LM_GI_Light60_sw62, LM_GI_Light69_sw62, LM_Inserts_l24_sw62)
Dim BP_sw63: BP_sw63=Array(BM_sw63, LM_GI_Light11_sw63)
Dim BP_sw63a: BP_sw63a=Array(BM_sw63a, LM_GI_Light10_sw63a)
Dim BP_sw70: BP_sw70=Array(BM_sw70, LM_GI_Light1_sw70, LM_GI_Light56_sw70, LM_GI_Light57_sw70, LM_GI_Light60_sw70, LM_GI_Light61_sw70)
Dim BP_sw72: BP_sw72=Array(BM_sw72, LM_GI_BumperLight4_sw72, LM_GI_Light52_sw72, LM_Inserts_l25_sw72)
Dim BP_sw74_001: BP_sw74_001=Array(BM_sw74_001, LM_GI_BumperLight2_sw74_001, LM_GI_Light53_sw74_001, LM_Inserts_l12_sw74_001)
' Arrays per lighting scenario
Dim BL_GI_BumperLight1: BL_GI_BumperLight1=Array(LM_GI_BumperLight1_Bumper_1_Rin, LM_GI_BumperLight1_Bumper_1_Soc, LM_GI_BumperLight1_Parts, LM_GI_BumperLight1_Plastics, LM_GI_BumperLight1_Playfield, LM_GI_BumperLight1_UnderPF, LM_GI_BumperLight1_sw52)
Dim BL_GI_BumperLight2: BL_GI_BumperLight2=Array(LM_GI_BumperLight2_Bumper_2_Rin, LM_GI_BumperLight2_Bumper_2_Soc, LM_GI_BumperLight2_Bumper_3_Rin, LM_GI_BumperLight2_Bumper_3_Soc, LM_GI_BumperLight2_Parts, LM_GI_BumperLight2_Plastics, LM_GI_BumperLight2_Playfield, LM_GI_BumperLight2_UnderPF, LM_GI_BumperLight2_sw74_001)
Dim BL_GI_BumperLight3: BL_GI_BumperLight3=Array(LM_GI_BumperLight3_Bumper_3_Rin, LM_GI_BumperLight3_Bumper_3_Soc, LM_GI_BumperLight3_Parts, LM_GI_BumperLight3_Plastics, LM_GI_BumperLight3_Playfield, LM_GI_BumperLight3_UnderPF, LM_GI_BumperLight3_sw62)
Dim BL_GI_BumperLight4: BL_GI_BumperLight4=Array(LM_GI_BumperLight4_Bumper_4_Rin, LM_GI_BumperLight4_Bumper_4_Soc, LM_GI_BumperLight4_Parts, LM_GI_BumperLight4_Plastics, LM_GI_BumperLight4_Playfield, LM_GI_BumperLight4_UnderPF, LM_GI_BumperLight4_sw72)
Dim BL_GI_Light1: BL_GI_Light1=Array(LM_GI_Light1_Layer1, LM_GI_Light1_Parts, LM_GI_Light1_Playfield, LM_GI_Light1_UnderPF, LM_GI_Light1_sw70)
Dim BL_GI_Light10: BL_GI_Light10=Array(LM_GI_Light10_Lsling, LM_GI_Light10_Lsling1, LM_GI_Light10_Lsling2, LM_GI_Light10_Parts, LM_GI_Light10_Plastics, LM_GI_Light10_Playfield, LM_GI_Light10_RF, LM_GI_Light10_RFU, LM_GI_Light10_Rsling, LM_GI_Light10_Rsling1, LM_GI_Light10_Rsling2, LM_GI_Light10_UnderPF, LM_GI_Light10_sw53, LM_GI_Light10_sw63a)
Dim BL_GI_Light11: BL_GI_Light11=Array(LM_GI_Light11_LF, LM_GI_Light11_LFU, LM_GI_Light11_Lsling, LM_GI_Light11_Lsling1, LM_GI_Light11_Lsling2, LM_GI_Light11_Parts, LM_GI_Light11_Plastics, LM_GI_Light11_Playfield, LM_GI_Light11_UnderPF, LM_GI_Light11_sw43, LM_GI_Light11_sw63)
Dim BL_GI_Light52: BL_GI_Light52=Array(LM_GI_Light52_Bumper_4_Ring, LM_GI_Light52_Bumper_4_Socket, LM_GI_Light52_Parts, LM_GI_Light52_Plastics, LM_GI_Light52_Playfield, LM_GI_Light52_Rsling, LM_GI_Light52_Rsling1, LM_GI_Light52_Rsling2, LM_GI_Light52_UnderPF, LM_GI_Light52_sw72)
Dim BL_GI_Light53: BL_GI_Light53=Array(LM_GI_Light53_Bumper_2_Ring, LM_GI_Light53_Bumper_2_Socket, LM_GI_Light53_Bumper_4_Ring, LM_GI_Light53_Bumper_4_Socket, LM_GI_Light53_Layer1, LM_GI_Light53_Parts, LM_GI_Light53_Plastics, LM_GI_Light53_Playfield, LM_GI_Light53_SpinnerRod, LM_GI_Light53_UnderPF, LM_GI_Light53_sw74_001)
Dim BL_GI_Light54: BL_GI_Light54=Array(LM_GI_Light54_Bumper_1_Ring, LM_GI_Light54_Bumper_1_Socket, LM_GI_Light54_Parts, LM_GI_Light54_Plastics, LM_GI_Light54_Playfield, LM_GI_Light54_UnderPF, LM_GI_Light54_sw42, LM_GI_Light54_sw52)
Dim BL_GI_Light55: BL_GI_Light55=Array(LM_GI_Light55_Bumper_1_Ring, LM_GI_Light55_Bumper_1_Socket, LM_GI_Light55_Parts, LM_GI_Light55_Plastics, LM_GI_Light55_Playfield, LM_GI_Light55_UnderPF)
Dim BL_GI_Light56: BL_GI_Light56=Array(LM_GI_Light56_Bumper_3_Ring, LM_GI_Light56_Bumper_3_Socket, LM_GI_Light56_Layer1, LM_GI_Light56_Parts, LM_GI_Light56_Plastics, LM_GI_Light56_Playfield, LM_GI_Light56_UnderPF, LM_GI_Light56_sw60, LM_GI_Light56_sw70)
Dim BL_GI_Light57: BL_GI_Light57=Array(LM_GI_Light57_Gate3, LM_GI_Light57_Gate4_Wire, LM_GI_Light57_Layer1, LM_GI_Light57_Parts, LM_GI_Light57_Plastics, LM_GI_Light57_Playfield, LM_GI_Light57_UnderPF, LM_GI_Light57_sw40, LM_GI_Light57_sw62, LM_GI_Light57_sw70)
Dim BL_GI_Light58: BL_GI_Light58=Array(LM_GI_Light58_Bumper_2_Ring, LM_GI_Light58_Bumper_2_Socket, LM_GI_Light58_Bumper_3_Ring, LM_GI_Light58_Bumper_3_Socket, LM_GI_Light58_Gate4_Wire, LM_GI_Light58_Layer1, LM_GI_Light58_Parts, LM_GI_Light58_Plastics, LM_GI_Light58_Playfield, LM_GI_Light58_UnderPF, LM_GI_Light58_sw41, LM_GI_Light58_sw60, LM_GI_Light58_sw62)
Dim BL_GI_Light59: BL_GI_Light59=Array(LM_GI_Light59_Gate2, LM_GI_Light59_Layer1, LM_GI_Light59_Parts, LM_GI_Light59_Plastics, LM_GI_Light59_Playfield, LM_GI_Light59_UnderPF, LM_GI_Light59_sw61)
Dim BL_GI_Light60: BL_GI_Light60=Array(LM_GI_Light60_Gate3, LM_GI_Light60_Gate4_Wire, LM_GI_Light60_Layer1, LM_GI_Light60_Parts, LM_GI_Light60_Plastics, LM_GI_Light60_Playfield, LM_GI_Light60_UnderPF, LM_GI_Light60_sw62, LM_GI_Light60_sw70)
Dim BL_GI_Light61: BL_GI_Light61=Array(LM_GI_Light61_Gate4_Wire, LM_GI_Light61_Layer1, LM_GI_Light61_Parts, LM_GI_Light61_Plastics, LM_GI_Light61_Playfield, LM_GI_Light61_UnderPF, LM_GI_Light61_sw70)
Dim BL_GI_Light68: BL_GI_Light68=Array(LM_GI_Light68_Gate1, LM_GI_Light68_Gate2, LM_GI_Light68_Gate3, LM_GI_Light68_Layer1, LM_GI_Light68_Parts, LM_GI_Light68_Plastics, LM_GI_Light68_Playfield, LM_GI_Light68_UnderPF, LM_GI_Light68_sw61)
Dim BL_GI_Light69: BL_GI_Light69=Array(LM_GI_Light69_Bumper_3_Ring, LM_GI_Light69_Bumper_3_Socket, LM_GI_Light69_Gate4_Wire, LM_GI_Light69_Layer1, LM_GI_Light69_Parts, LM_GI_Light69_Plastics, LM_GI_Light69_Playfield, LM_GI_Light69_UnderPF, LM_GI_Light69_sw41, LM_GI_Light69_sw60, LM_GI_Light69_sw62)
Dim BL_GI_Light70: BL_GI_Light70=Array(LM_GI_Light70_Gate2, LM_GI_Light70_Layer1, LM_GI_Light70_Parts, LM_GI_Light70_Plastics, LM_GI_Light70_Playfield, LM_GI_Light70_UnderPF, LM_GI_Light70_sw41, LM_GI_Light70_sw51)
Dim BL_GI_Light71: BL_GI_Light71=Array(LM_GI_Light71_Gate2, LM_GI_Light71_Gate3, LM_GI_Light71_Layer1, LM_GI_Light71_Parts, LM_GI_Light71_Plastics, LM_GI_Light71_Playfield, LM_GI_Light71_UnderPF, LM_GI_Light71_sw51, LM_GI_Light71_sw61)
Dim BL_Inserts_l12: BL_Inserts_l12=Array(LM_Inserts_l12_Parts, LM_Inserts_l12_Plastics, LM_Inserts_l12_Playfield, LM_Inserts_l12_SpinnerRod, LM_Inserts_l12_UnderPF, LM_Inserts_l12_sw74_001)
Dim BL_Inserts_l13: BL_Inserts_l13=Array(LM_Inserts_l13_Playfield, LM_Inserts_l13_UnderPF)
Dim BL_Inserts_l14: BL_Inserts_l14=Array(LM_Inserts_l14_Playfield, LM_Inserts_l14_UnderPF)
Dim BL_Inserts_l15: BL_Inserts_l15=Array(LM_Inserts_l15_Playfield, LM_Inserts_l15_UnderPF)
Dim BL_Inserts_l16: BL_Inserts_l16=Array(LM_Inserts_l16_Playfield, LM_Inserts_l16_UnderPF)
Dim BL_Inserts_l17: BL_Inserts_l17=Array(LM_Inserts_l17_Playfield, LM_Inserts_l17_UnderPF)
Dim BL_Inserts_l18: BL_Inserts_l18=Array(LM_Inserts_l18_Playfield, LM_Inserts_l18_UnderPF)
Dim BL_Inserts_l19: BL_Inserts_l19=Array(LM_Inserts_l19_Playfield, LM_Inserts_l19_UnderPF)
Dim BL_Inserts_l20: BL_Inserts_l20=Array(LM_Inserts_l20_Playfield, LM_Inserts_l20_UnderPF)
Dim BL_Inserts_l21: BL_Inserts_l21=Array(LM_Inserts_l21_Layer1, LM_Inserts_l21_Parts, LM_Inserts_l21_Playfield, LM_Inserts_l21_UnderPF, LM_Inserts_l21_sw40, LM_Inserts_l21_sw50)
Dim BL_Inserts_l22: BL_Inserts_l22=Array(LM_Inserts_l22_Parts, LM_Inserts_l22_Plastics, LM_Inserts_l22_Playfield, LM_Inserts_l22_UnderPF, LM_Inserts_l22_sw42, LM_Inserts_l22_sw52)
Dim BL_Inserts_l23: BL_Inserts_l23=Array(LM_Inserts_l23_Bumper_1_Ring, LM_Inserts_l23_Bumper_1_Socket, LM_Inserts_l23_Parts, LM_Inserts_l23_Plastics, LM_Inserts_l23_Playfield, LM_Inserts_l23_UnderPF, LM_Inserts_l23_sw52)
Dim BL_Inserts_l24: BL_Inserts_l24=Array(LM_Inserts_l24_Bumper_2_Ring, LM_Inserts_l24_Bumper_3_Ring, LM_Inserts_l24_Bumper_3_Socket, LM_Inserts_l24_Parts, LM_Inserts_l24_Plastics, LM_Inserts_l24_Playfield, LM_Inserts_l24_UnderPF, LM_Inserts_l24_sw62)
Dim BL_Inserts_l25: BL_Inserts_l25=Array(LM_Inserts_l25_Parts, LM_Inserts_l25_Plastics, LM_Inserts_l25_Playfield, LM_Inserts_l25_UnderPF, LM_Inserts_l25_sw72)
Dim BL_Inserts_l26: BL_Inserts_l26=Array(LM_Inserts_l26_Gate3, LM_Inserts_l26_Layer1, LM_Inserts_l26_Parts, LM_Inserts_l26_Playfield, LM_Inserts_l26_UnderPF)
Dim BL_Inserts_l27: BL_Inserts_l27=Array(LM_Inserts_l27_Layer1, LM_Inserts_l27_Parts, LM_Inserts_l27_Playfield, LM_Inserts_l27_UnderPF)
Dim BL_Inserts_l28: BL_Inserts_l28=Array(LM_Inserts_l28_Layer1, LM_Inserts_l28_Parts, LM_Inserts_l28_Plastics, LM_Inserts_l28_Playfield, LM_Inserts_l28_UnderPF)
Dim BL_Inserts_l29: BL_Inserts_l29=Array(LM_Inserts_l29_Playfield, LM_Inserts_l29_UnderPF)
Dim BL_Inserts_l3: BL_Inserts_l3=Array(LM_Inserts_l3_LF, LM_Inserts_l3_Playfield, LM_Inserts_l3_UnderPF)
Dim BL_Inserts_l30: BL_Inserts_l30=Array(LM_Inserts_l30_Playfield, LM_Inserts_l30_RFU, LM_Inserts_l30_UnderPF)
Dim BL_Inserts_l31: BL_Inserts_l31=Array(LM_Inserts_l31_Playfield, LM_Inserts_l31_UnderPF)
Dim BL_Inserts_l32: BL_Inserts_l32=Array(LM_Inserts_l32_Playfield, LM_Inserts_l32_UnderPF)
Dim BL_Inserts_l33: BL_Inserts_l33=Array(LM_Inserts_l33_LFU, LM_Inserts_l33_Playfield, LM_Inserts_l33_UnderPF)
Dim BL_Inserts_l34: BL_Inserts_l34=Array(LM_Inserts_l34_Playfield, LM_Inserts_l34_UnderPF)
Dim BL_Inserts_l35: BL_Inserts_l35=Array(LM_Inserts_l35_Playfield, LM_Inserts_l35_UnderPF)
Dim BL_Inserts_l36: BL_Inserts_l36=Array(LM_Inserts_l36_Playfield, LM_Inserts_l36_UnderPF)
Dim BL_Inserts_l37: BL_Inserts_l37=Array(LM_Inserts_l37_Playfield, LM_Inserts_l37_UnderPF)
Dim BL_Inserts_l38: BL_Inserts_l38=Array(LM_Inserts_l38_Playfield, LM_Inserts_l38_UnderPF)
Dim BL_Inserts_l39: BL_Inserts_l39=Array(LM_Inserts_l39_Playfield, LM_Inserts_l39_UnderPF)
Dim BL_Inserts_l4: BL_Inserts_l4=Array(LM_Inserts_l4_Parts, LM_Inserts_l4_Plastics, LM_Inserts_l4_Playfield, LM_Inserts_l4_UnderPF)
Dim BL_Inserts_l40: BL_Inserts_l40=Array(LM_Inserts_l40_Playfield, LM_Inserts_l40_UnderPF)
Dim BL_Inserts_l41: BL_Inserts_l41=Array(LM_Inserts_l41_Playfield, LM_Inserts_l41_UnderPF)
Dim BL_Inserts_l42: BL_Inserts_l42=Array(LM_Inserts_l42_Playfield, LM_Inserts_l42_UnderPF)
Dim BL_Inserts_l43: BL_Inserts_l43=Array(LM_Inserts_l43_Playfield, LM_Inserts_l43_UnderPF)
Dim BL_Inserts_l44: BL_Inserts_l44=Array(LM_Inserts_l44_Lsling, LM_Inserts_l44_Lsling1, LM_Inserts_l44_Lsling2, LM_Inserts_l44_Parts, LM_Inserts_l44_Plastics, LM_Inserts_l44_Playfield, LM_Inserts_l44_UnderPF)
Dim BL_Inserts_l45: BL_Inserts_l45=Array(LM_Inserts_l45_Gate3, LM_Inserts_l45_Gate4_Wire, LM_Inserts_l45_Layer1, LM_Inserts_l45_Parts, LM_Inserts_l45_Playfield, LM_Inserts_l45_UnderPF, LM_Inserts_l45_sw41)
Dim BL_Inserts_l46: BL_Inserts_l46=Array(LM_Inserts_l46_Layer1, LM_Inserts_l46_Parts, LM_Inserts_l46_Playfield, LM_Inserts_l46_UnderPF, LM_Inserts_l46_sw51)
Dim BL_Inserts_l47: BL_Inserts_l47=Array(LM_Inserts_l47_Gate2, LM_Inserts_l47_Layer1, LM_Inserts_l47_Parts, LM_Inserts_l47_Plastics, LM_Inserts_l47_Playfield, LM_Inserts_l47_UnderPF)
Dim BL_Inserts_l48: BL_Inserts_l48=Array(LM_Inserts_l48_Parts, LM_Inserts_l48_Plastics, LM_Inserts_l48_Playfield, LM_Inserts_l48_Rsling, LM_Inserts_l48_Rsling1, LM_Inserts_l48_Rsling2, LM_Inserts_l48_UnderPF)
Dim BL_Inserts_l49: BL_Inserts_l49=Array(LM_Inserts_l49_Playfield, LM_Inserts_l49_UnderPF)
Dim BL_Inserts_l5: BL_Inserts_l5=Array(LM_Inserts_l5_Layer1, LM_Inserts_l5_Parts, LM_Inserts_l5_Plastics, LM_Inserts_l5_Playfield, LM_Inserts_l5_UnderPF, LM_Inserts_l5_sw40)
Dim BL_Inserts_l50: BL_Inserts_l50=Array(LM_Inserts_l50_Playfield, LM_Inserts_l50_UnderPF)
Dim BL_Inserts_l51: BL_Inserts_l51=Array(LM_Inserts_l51_Bumper_2_Ring, LM_Inserts_l51_Bumper_2_Socket, LM_Inserts_l51_Parts, LM_Inserts_l51_Playfield, LM_Inserts_l51_UnderPF)
Dim BL_Inserts_l6: BL_Inserts_l6=Array(LM_Inserts_l6_Layer1, LM_Inserts_l6_Parts, LM_Inserts_l6_Plastics, LM_Inserts_l6_Playfield, LM_Inserts_l6_UnderPF, LM_Inserts_l6_sw60)
Dim BL_Inserts_l7: BL_Inserts_l7=Array(LM_Inserts_l7_Gate4_Wire, LM_Inserts_l7_Layer1, LM_Inserts_l7_Parts, LM_Inserts_l7_Plastics, LM_Inserts_l7_Playfield, LM_Inserts_l7_UnderPF)
Dim BL_World: BL_World=Array(BM_Bumper_1_Ring, BM_Bumper_1_Socket, BM_Bumper_2_Ring, BM_Bumper_2_Socket, BM_Bumper_3_Ring, BM_Bumper_3_Socket, BM_Bumper_4_Ring, BM_Bumper_4_Socket, BM_Gate1, BM_Gate2, BM_Gate3, BM_Gate4_Wire, BM_LF, BM_LFU, BM_Layer1, BM_Lsling, BM_Lsling1, BM_Lsling2, BM_Parts, BM_Plastics, BM_Playfield, BM_Playfield_001, BM_RF, BM_RFU, BM_Rsling, BM_Rsling1, BM_Rsling2, BM_SpinnerRod, BM_UnderPF, BM_sw40, BM_sw41, BM_sw42, BM_sw43, BM_sw50, BM_sw51, BM_sw52, BM_sw53, BM_sw60, BM_sw61, BM_sw62, BM_sw63, BM_sw63a, BM_sw70, BM_sw72, BM_sw74_001)
' Global arrays
Dim BG_Bakemap: BG_Bakemap=Array(BM_Bumper_1_Ring, BM_Bumper_1_Socket, BM_Bumper_2_Ring, BM_Bumper_2_Socket, BM_Bumper_3_Ring, BM_Bumper_3_Socket, BM_Bumper_4_Ring, BM_Bumper_4_Socket, BM_Gate1, BM_Gate2, BM_Gate3, BM_Gate4_Wire, BM_LF, BM_LFU, BM_Layer1, BM_Lsling, BM_Lsling1, BM_Lsling2, BM_Parts, BM_Plastics, BM_Playfield, BM_Playfield_001, BM_RF, BM_RFU, BM_Rsling, BM_Rsling1, BM_Rsling2, BM_SpinnerRod, BM_UnderPF, BM_sw40, BM_sw41, BM_sw42, BM_sw43, BM_sw50, BM_sw51, BM_sw52, BM_sw53, BM_sw60, BM_sw61, BM_sw62, BM_sw63, BM_sw63a, BM_sw70, BM_sw72, BM_sw74_001)
Dim BG_Lightmap: BG_Lightmap=Array(LM_GI_BumperLight1_Bumper_1_Rin, LM_GI_BumperLight1_Bumper_1_Soc, LM_GI_BumperLight1_Parts, LM_GI_BumperLight1_Plastics, LM_GI_BumperLight1_Playfield, LM_GI_BumperLight1_UnderPF, LM_GI_BumperLight1_sw52, LM_GI_BumperLight2_Bumper_2_Rin, LM_GI_BumperLight2_Bumper_2_Soc, LM_GI_BumperLight2_Bumper_3_Rin, LM_GI_BumperLight2_Bumper_3_Soc, LM_GI_BumperLight2_Parts, LM_GI_BumperLight2_Plastics, LM_GI_BumperLight2_Playfield, LM_GI_BumperLight2_UnderPF, LM_GI_BumperLight2_sw74_001, LM_GI_BumperLight3_Bumper_3_Rin, LM_GI_BumperLight3_Bumper_3_Soc, LM_GI_BumperLight3_Parts, LM_GI_BumperLight3_Plastics, LM_GI_BumperLight3_Playfield, LM_GI_BumperLight3_UnderPF, LM_GI_BumperLight3_sw62, LM_GI_BumperLight4_Bumper_4_Rin, LM_GI_BumperLight4_Bumper_4_Soc, LM_GI_BumperLight4_Parts, LM_GI_BumperLight4_Plastics, LM_GI_BumperLight4_Playfield, LM_GI_BumperLight4_UnderPF, LM_GI_BumperLight4_sw72, LM_GI_Light1_Layer1, LM_GI_Light1_Parts, LM_GI_Light1_Playfield, LM_GI_Light1_UnderPF, _
  LM_GI_Light1_sw70, LM_GI_Light10_Lsling, LM_GI_Light10_Lsling1, LM_GI_Light10_Lsling2, LM_GI_Light10_Parts, LM_GI_Light10_Plastics, LM_GI_Light10_Playfield, LM_GI_Light10_RF, LM_GI_Light10_RFU, LM_GI_Light10_Rsling, LM_GI_Light10_Rsling1, LM_GI_Light10_Rsling2, LM_GI_Light10_UnderPF, LM_GI_Light10_sw53, LM_GI_Light10_sw63a, LM_GI_Light11_LF, LM_GI_Light11_LFU, LM_GI_Light11_Lsling, LM_GI_Light11_Lsling1, LM_GI_Light11_Lsling2, LM_GI_Light11_Parts, LM_GI_Light11_Plastics, LM_GI_Light11_Playfield, LM_GI_Light11_UnderPF, LM_GI_Light11_sw43, LM_GI_Light11_sw63, LM_GI_Light52_Bumper_4_Ring, LM_GI_Light52_Bumper_4_Socket, LM_GI_Light52_Parts, LM_GI_Light52_Plastics, LM_GI_Light52_Playfield, LM_GI_Light52_Rsling, LM_GI_Light52_Rsling1, LM_GI_Light52_Rsling2, LM_GI_Light52_UnderPF, LM_GI_Light52_sw72, LM_GI_Light53_Bumper_2_Ring, LM_GI_Light53_Bumper_2_Socket, LM_GI_Light53_Bumper_4_Ring, LM_GI_Light53_Bumper_4_Socket, LM_GI_Light53_Layer1, LM_GI_Light53_Parts, LM_GI_Light53_Plastics, LM_GI_Light53_Playfield, _
  LM_GI_Light53_SpinnerRod, LM_GI_Light53_UnderPF, LM_GI_Light53_sw74_001, LM_GI_Light54_Bumper_1_Ring, LM_GI_Light54_Bumper_1_Socket, LM_GI_Light54_Parts, LM_GI_Light54_Plastics, LM_GI_Light54_Playfield, LM_GI_Light54_UnderPF, LM_GI_Light54_sw42, LM_GI_Light54_sw52, LM_GI_Light55_Bumper_1_Ring, LM_GI_Light55_Bumper_1_Socket, LM_GI_Light55_Parts, LM_GI_Light55_Plastics, LM_GI_Light55_Playfield, LM_GI_Light55_UnderPF, LM_GI_Light56_Bumper_3_Ring, LM_GI_Light56_Bumper_3_Socket, LM_GI_Light56_Layer1, LM_GI_Light56_Parts, LM_GI_Light56_Plastics, LM_GI_Light56_Playfield, LM_GI_Light56_UnderPF, LM_GI_Light56_sw60, LM_GI_Light56_sw70, LM_GI_Light57_Gate3, LM_GI_Light57_Gate4_Wire, LM_GI_Light57_Layer1, LM_GI_Light57_Parts, LM_GI_Light57_Plastics, LM_GI_Light57_Playfield, LM_GI_Light57_UnderPF, LM_GI_Light57_sw40, LM_GI_Light57_sw62, LM_GI_Light57_sw70, LM_GI_Light58_Bumper_2_Ring, LM_GI_Light58_Bumper_2_Socket, LM_GI_Light58_Bumper_3_Ring, LM_GI_Light58_Bumper_3_Socket, LM_GI_Light58_Gate4_Wire, LM_GI_Light58_Layer1, _
  LM_GI_Light58_Parts, LM_GI_Light58_Plastics, LM_GI_Light58_Playfield, LM_GI_Light58_UnderPF, LM_GI_Light58_sw41, LM_GI_Light58_sw60, LM_GI_Light58_sw62, LM_GI_Light59_Gate2, LM_GI_Light59_Layer1, LM_GI_Light59_Parts, LM_GI_Light59_Plastics, LM_GI_Light59_Playfield, LM_GI_Light59_UnderPF, LM_GI_Light59_sw61, LM_GI_Light60_Gate3, LM_GI_Light60_Gate4_Wire, LM_GI_Light60_Layer1, LM_GI_Light60_Parts, LM_GI_Light60_Plastics, LM_GI_Light60_Playfield, LM_GI_Light60_UnderPF, LM_GI_Light60_sw62, LM_GI_Light60_sw70, LM_GI_Light61_Gate4_Wire, LM_GI_Light61_Layer1, LM_GI_Light61_Parts, LM_GI_Light61_Plastics, LM_GI_Light61_Playfield, LM_GI_Light61_UnderPF, LM_GI_Light61_sw70, LM_GI_Light68_Gate1, LM_GI_Light68_Gate2, LM_GI_Light68_Gate3, LM_GI_Light68_Layer1, LM_GI_Light68_Parts, LM_GI_Light68_Plastics, LM_GI_Light68_Playfield, LM_GI_Light68_UnderPF, LM_GI_Light68_sw61, LM_GI_Light69_Bumper_3_Ring, LM_GI_Light69_Bumper_3_Socket, LM_GI_Light69_Gate4_Wire, LM_GI_Light69_Layer1, LM_GI_Light69_Parts, LM_GI_Light69_Plastics, _
  LM_GI_Light69_Playfield, LM_GI_Light69_UnderPF, LM_GI_Light69_sw41, LM_GI_Light69_sw60, LM_GI_Light69_sw62, LM_GI_Light70_Gate2, LM_GI_Light70_Layer1, LM_GI_Light70_Parts, LM_GI_Light70_Plastics, LM_GI_Light70_Playfield, LM_GI_Light70_UnderPF, LM_GI_Light70_sw41, LM_GI_Light70_sw51, LM_GI_Light71_Gate2, LM_GI_Light71_Gate3, LM_GI_Light71_Layer1, LM_GI_Light71_Parts, LM_GI_Light71_Plastics, LM_GI_Light71_Playfield, LM_GI_Light71_UnderPF, LM_GI_Light71_sw51, LM_GI_Light71_sw61, LM_Inserts_l12_Parts, LM_Inserts_l12_Plastics, LM_Inserts_l12_Playfield, LM_Inserts_l12_SpinnerRod, LM_Inserts_l12_UnderPF, LM_Inserts_l12_sw74_001, LM_Inserts_l13_Playfield, LM_Inserts_l13_UnderPF, LM_Inserts_l14_Playfield, LM_Inserts_l14_UnderPF, LM_Inserts_l15_Playfield, LM_Inserts_l15_UnderPF, LM_Inserts_l16_Playfield, LM_Inserts_l16_UnderPF, LM_Inserts_l17_Playfield, LM_Inserts_l17_UnderPF, LM_Inserts_l18_Playfield, LM_Inserts_l18_UnderPF, LM_Inserts_l19_Playfield, LM_Inserts_l19_UnderPF, LM_Inserts_l20_Playfield, _
  LM_Inserts_l20_UnderPF, LM_Inserts_l21_Layer1, LM_Inserts_l21_Parts, LM_Inserts_l21_Playfield, LM_Inserts_l21_UnderPF, LM_Inserts_l21_sw40, LM_Inserts_l21_sw50, LM_Inserts_l22_Parts, LM_Inserts_l22_Plastics, LM_Inserts_l22_Playfield, LM_Inserts_l22_UnderPF, LM_Inserts_l22_sw42, LM_Inserts_l22_sw52, LM_Inserts_l23_Bumper_1_Ring, LM_Inserts_l23_Bumper_1_Socket, LM_Inserts_l23_Parts, LM_Inserts_l23_Plastics, LM_Inserts_l23_Playfield, LM_Inserts_l23_UnderPF, LM_Inserts_l23_sw52, LM_Inserts_l24_Bumper_2_Ring, LM_Inserts_l24_Bumper_3_Ring, LM_Inserts_l24_Bumper_3_Socket, LM_Inserts_l24_Parts, LM_Inserts_l24_Plastics, LM_Inserts_l24_Playfield, LM_Inserts_l24_UnderPF, LM_Inserts_l24_sw62, LM_Inserts_l25_Parts, LM_Inserts_l25_Plastics, LM_Inserts_l25_Playfield, LM_Inserts_l25_UnderPF, LM_Inserts_l25_sw72, LM_Inserts_l26_Gate3, LM_Inserts_l26_Layer1, LM_Inserts_l26_Parts, LM_Inserts_l26_Playfield, LM_Inserts_l26_UnderPF, LM_Inserts_l27_Layer1, LM_Inserts_l27_Parts, LM_Inserts_l27_Playfield, LM_Inserts_l27_UnderPF, _
  LM_Inserts_l28_Layer1, LM_Inserts_l28_Parts, LM_Inserts_l28_Plastics, LM_Inserts_l28_Playfield, LM_Inserts_l28_UnderPF, LM_Inserts_l29_Playfield, LM_Inserts_l29_UnderPF, LM_Inserts_l3_LF, LM_Inserts_l3_Playfield, LM_Inserts_l3_UnderPF, LM_Inserts_l30_Playfield, LM_Inserts_l30_RFU, LM_Inserts_l30_UnderPF, LM_Inserts_l31_Playfield, LM_Inserts_l31_UnderPF, LM_Inserts_l32_Playfield, LM_Inserts_l32_UnderPF, LM_Inserts_l33_LFU, LM_Inserts_l33_Playfield, LM_Inserts_l33_UnderPF, LM_Inserts_l34_Playfield, LM_Inserts_l34_UnderPF, LM_Inserts_l35_Playfield, LM_Inserts_l35_UnderPF, LM_Inserts_l36_Playfield, LM_Inserts_l36_UnderPF, LM_Inserts_l37_Playfield, LM_Inserts_l37_UnderPF, LM_Inserts_l38_Playfield, LM_Inserts_l38_UnderPF, LM_Inserts_l39_Playfield, LM_Inserts_l39_UnderPF, LM_Inserts_l4_Parts, LM_Inserts_l4_Plastics, LM_Inserts_l4_Playfield, LM_Inserts_l4_UnderPF, LM_Inserts_l40_Playfield, LM_Inserts_l40_UnderPF, LM_Inserts_l41_Playfield, LM_Inserts_l41_UnderPF, LM_Inserts_l42_Playfield, LM_Inserts_l42_UnderPF, _
  LM_Inserts_l43_Playfield, LM_Inserts_l43_UnderPF, LM_Inserts_l44_Lsling, LM_Inserts_l44_Lsling1, LM_Inserts_l44_Lsling2, LM_Inserts_l44_Parts, LM_Inserts_l44_Plastics, LM_Inserts_l44_Playfield, LM_Inserts_l44_UnderPF, LM_Inserts_l45_Gate3, LM_Inserts_l45_Gate4_Wire, LM_Inserts_l45_Layer1, LM_Inserts_l45_Parts, LM_Inserts_l45_Playfield, LM_Inserts_l45_UnderPF, LM_Inserts_l45_sw41, LM_Inserts_l46_Layer1, LM_Inserts_l46_Parts, LM_Inserts_l46_Playfield, LM_Inserts_l46_UnderPF, LM_Inserts_l46_sw51, LM_Inserts_l47_Gate2, LM_Inserts_l47_Layer1, LM_Inserts_l47_Parts, LM_Inserts_l47_Plastics, LM_Inserts_l47_Playfield, LM_Inserts_l47_UnderPF, LM_Inserts_l48_Parts, LM_Inserts_l48_Plastics, LM_Inserts_l48_Playfield, LM_Inserts_l48_Rsling, LM_Inserts_l48_Rsling1, LM_Inserts_l48_Rsling2, LM_Inserts_l48_UnderPF, LM_Inserts_l49_Playfield, LM_Inserts_l49_UnderPF, LM_Inserts_l5_Layer1, LM_Inserts_l5_Parts, LM_Inserts_l5_Plastics, LM_Inserts_l5_Playfield, LM_Inserts_l5_UnderPF, LM_Inserts_l5_sw40, LM_Inserts_l50_Playfield, _
  LM_Inserts_l50_UnderPF, LM_Inserts_l51_Bumper_2_Ring, LM_Inserts_l51_Bumper_2_Socket, LM_Inserts_l51_Parts, LM_Inserts_l51_Playfield, LM_Inserts_l51_UnderPF, LM_Inserts_l6_Layer1, LM_Inserts_l6_Parts, LM_Inserts_l6_Plastics, LM_Inserts_l6_Playfield, LM_Inserts_l6_UnderPF, LM_Inserts_l6_sw60, LM_Inserts_l7_Gate4_Wire, LM_Inserts_l7_Layer1, LM_Inserts_l7_Parts, LM_Inserts_l7_Plastics, LM_Inserts_l7_Playfield, LM_Inserts_l7_UnderPF)
Dim BG_All: BG_All=Array(BM_Bumper_1_Ring, BM_Bumper_1_Socket, BM_Bumper_2_Ring, BM_Bumper_2_Socket, BM_Bumper_3_Ring, BM_Bumper_3_Socket, BM_Bumper_4_Ring, BM_Bumper_4_Socket, BM_Gate1, BM_Gate2, BM_Gate3, BM_Gate4_Wire, BM_LF, BM_LFU, BM_Layer1, BM_Lsling, BM_Lsling1, BM_Lsling2, BM_Parts, BM_Plastics, BM_Playfield, BM_Playfield_001, BM_RF, BM_RFU, BM_Rsling, BM_Rsling1, BM_Rsling2, BM_SpinnerRod, BM_UnderPF, BM_sw40, BM_sw41, BM_sw42, BM_sw43, BM_sw50, BM_sw51, BM_sw52, BM_sw53, BM_sw60, BM_sw61, BM_sw62, BM_sw63, BM_sw63a, BM_sw70, BM_sw72, BM_sw74_001, LM_GI_BumperLight1_Bumper_1_Rin, LM_GI_BumperLight1_Bumper_1_Soc, LM_GI_BumperLight1_Parts, LM_GI_BumperLight1_Plastics, LM_GI_BumperLight1_Playfield, LM_GI_BumperLight1_UnderPF, LM_GI_BumperLight1_sw52, LM_GI_BumperLight2_Bumper_2_Rin, LM_GI_BumperLight2_Bumper_2_Soc, LM_GI_BumperLight2_Bumper_3_Rin, LM_GI_BumperLight2_Bumper_3_Soc, LM_GI_BumperLight2_Parts, LM_GI_BumperLight2_Plastics, LM_GI_BumperLight2_Playfield, LM_GI_BumperLight2_UnderPF, _
  LM_GI_BumperLight2_sw74_001, LM_GI_BumperLight3_Bumper_3_Rin, LM_GI_BumperLight3_Bumper_3_Soc, LM_GI_BumperLight3_Parts, LM_GI_BumperLight3_Plastics, LM_GI_BumperLight3_Playfield, LM_GI_BumperLight3_UnderPF, LM_GI_BumperLight3_sw62, LM_GI_BumperLight4_Bumper_4_Rin, LM_GI_BumperLight4_Bumper_4_Soc, LM_GI_BumperLight4_Parts, LM_GI_BumperLight4_Plastics, LM_GI_BumperLight4_Playfield, LM_GI_BumperLight4_UnderPF, LM_GI_BumperLight4_sw72, LM_GI_Light1_Layer1, LM_GI_Light1_Parts, LM_GI_Light1_Playfield, LM_GI_Light1_UnderPF, LM_GI_Light1_sw70, LM_GI_Light10_Lsling, LM_GI_Light10_Lsling1, LM_GI_Light10_Lsling2, LM_GI_Light10_Parts, LM_GI_Light10_Plastics, LM_GI_Light10_Playfield, LM_GI_Light10_RF, LM_GI_Light10_RFU, LM_GI_Light10_Rsling, LM_GI_Light10_Rsling1, LM_GI_Light10_Rsling2, LM_GI_Light10_UnderPF, LM_GI_Light10_sw53, LM_GI_Light10_sw63a, LM_GI_Light11_LF, LM_GI_Light11_LFU, LM_GI_Light11_Lsling, LM_GI_Light11_Lsling1, LM_GI_Light11_Lsling2, LM_GI_Light11_Parts, LM_GI_Light11_Plastics, LM_GI_Light11_Playfield, _
  LM_GI_Light11_UnderPF, LM_GI_Light11_sw43, LM_GI_Light11_sw63, LM_GI_Light52_Bumper_4_Ring, LM_GI_Light52_Bumper_4_Socket, LM_GI_Light52_Parts, LM_GI_Light52_Plastics, LM_GI_Light52_Playfield, LM_GI_Light52_Rsling, LM_GI_Light52_Rsling1, LM_GI_Light52_Rsling2, LM_GI_Light52_UnderPF, LM_GI_Light52_sw72, LM_GI_Light53_Bumper_2_Ring, LM_GI_Light53_Bumper_2_Socket, LM_GI_Light53_Bumper_4_Ring, LM_GI_Light53_Bumper_4_Socket, LM_GI_Light53_Layer1, LM_GI_Light53_Parts, LM_GI_Light53_Plastics, LM_GI_Light53_Playfield, LM_GI_Light53_SpinnerRod, LM_GI_Light53_UnderPF, LM_GI_Light53_sw74_001, LM_GI_Light54_Bumper_1_Ring, LM_GI_Light54_Bumper_1_Socket, LM_GI_Light54_Parts, LM_GI_Light54_Plastics, LM_GI_Light54_Playfield, LM_GI_Light54_UnderPF, LM_GI_Light54_sw42, LM_GI_Light54_sw52, LM_GI_Light55_Bumper_1_Ring, LM_GI_Light55_Bumper_1_Socket, LM_GI_Light55_Parts, LM_GI_Light55_Plastics, LM_GI_Light55_Playfield, LM_GI_Light55_UnderPF, LM_GI_Light56_Bumper_3_Ring, LM_GI_Light56_Bumper_3_Socket, LM_GI_Light56_Layer1, _
  LM_GI_Light56_Parts, LM_GI_Light56_Plastics, LM_GI_Light56_Playfield, LM_GI_Light56_UnderPF, LM_GI_Light56_sw60, LM_GI_Light56_sw70, LM_GI_Light57_Gate3, LM_GI_Light57_Gate4_Wire, LM_GI_Light57_Layer1, LM_GI_Light57_Parts, LM_GI_Light57_Plastics, LM_GI_Light57_Playfield, LM_GI_Light57_UnderPF, LM_GI_Light57_sw40, LM_GI_Light57_sw62, LM_GI_Light57_sw70, LM_GI_Light58_Bumper_2_Ring, LM_GI_Light58_Bumper_2_Socket, LM_GI_Light58_Bumper_3_Ring, LM_GI_Light58_Bumper_3_Socket, LM_GI_Light58_Gate4_Wire, LM_GI_Light58_Layer1, LM_GI_Light58_Parts, LM_GI_Light58_Plastics, LM_GI_Light58_Playfield, LM_GI_Light58_UnderPF, LM_GI_Light58_sw41, LM_GI_Light58_sw60, LM_GI_Light58_sw62, LM_GI_Light59_Gate2, LM_GI_Light59_Layer1, LM_GI_Light59_Parts, LM_GI_Light59_Plastics, LM_GI_Light59_Playfield, LM_GI_Light59_UnderPF, LM_GI_Light59_sw61, LM_GI_Light60_Gate3, LM_GI_Light60_Gate4_Wire, LM_GI_Light60_Layer1, LM_GI_Light60_Parts, LM_GI_Light60_Plastics, LM_GI_Light60_Playfield, LM_GI_Light60_UnderPF, LM_GI_Light60_sw62, _
  LM_GI_Light60_sw70, LM_GI_Light61_Gate4_Wire, LM_GI_Light61_Layer1, LM_GI_Light61_Parts, LM_GI_Light61_Plastics, LM_GI_Light61_Playfield, LM_GI_Light61_UnderPF, LM_GI_Light61_sw70, LM_GI_Light68_Gate1, LM_GI_Light68_Gate2, LM_GI_Light68_Gate3, LM_GI_Light68_Layer1, LM_GI_Light68_Parts, LM_GI_Light68_Plastics, LM_GI_Light68_Playfield, LM_GI_Light68_UnderPF, LM_GI_Light68_sw61, LM_GI_Light69_Bumper_3_Ring, LM_GI_Light69_Bumper_3_Socket, LM_GI_Light69_Gate4_Wire, LM_GI_Light69_Layer1, LM_GI_Light69_Parts, LM_GI_Light69_Plastics, LM_GI_Light69_Playfield, LM_GI_Light69_UnderPF, LM_GI_Light69_sw41, LM_GI_Light69_sw60, LM_GI_Light69_sw62, LM_GI_Light70_Gate2, LM_GI_Light70_Layer1, LM_GI_Light70_Parts, LM_GI_Light70_Plastics, LM_GI_Light70_Playfield, LM_GI_Light70_UnderPF, LM_GI_Light70_sw41, LM_GI_Light70_sw51, LM_GI_Light71_Gate2, LM_GI_Light71_Gate3, LM_GI_Light71_Layer1, LM_GI_Light71_Parts, LM_GI_Light71_Plastics, LM_GI_Light71_Playfield, LM_GI_Light71_UnderPF, LM_GI_Light71_sw51, LM_GI_Light71_sw61, _
  LM_Inserts_l12_Parts, LM_Inserts_l12_Plastics, LM_Inserts_l12_Playfield, LM_Inserts_l12_SpinnerRod, LM_Inserts_l12_UnderPF, LM_Inserts_l12_sw74_001, LM_Inserts_l13_Playfield, LM_Inserts_l13_UnderPF, LM_Inserts_l14_Playfield, LM_Inserts_l14_UnderPF, LM_Inserts_l15_Playfield, LM_Inserts_l15_UnderPF, LM_Inserts_l16_Playfield, LM_Inserts_l16_UnderPF, LM_Inserts_l17_Playfield, LM_Inserts_l17_UnderPF, LM_Inserts_l18_Playfield, LM_Inserts_l18_UnderPF, LM_Inserts_l19_Playfield, LM_Inserts_l19_UnderPF, LM_Inserts_l20_Playfield, LM_Inserts_l20_UnderPF, LM_Inserts_l21_Layer1, LM_Inserts_l21_Parts, LM_Inserts_l21_Playfield, LM_Inserts_l21_UnderPF, LM_Inserts_l21_sw40, LM_Inserts_l21_sw50, LM_Inserts_l22_Parts, LM_Inserts_l22_Plastics, LM_Inserts_l22_Playfield, LM_Inserts_l22_UnderPF, LM_Inserts_l22_sw42, LM_Inserts_l22_sw52, LM_Inserts_l23_Bumper_1_Ring, LM_Inserts_l23_Bumper_1_Socket, LM_Inserts_l23_Parts, LM_Inserts_l23_Plastics, LM_Inserts_l23_Playfield, LM_Inserts_l23_UnderPF, LM_Inserts_l23_sw52, _
  LM_Inserts_l24_Bumper_2_Ring, LM_Inserts_l24_Bumper_3_Ring, LM_Inserts_l24_Bumper_3_Socket, LM_Inserts_l24_Parts, LM_Inserts_l24_Plastics, LM_Inserts_l24_Playfield, LM_Inserts_l24_UnderPF, LM_Inserts_l24_sw62, LM_Inserts_l25_Parts, LM_Inserts_l25_Plastics, LM_Inserts_l25_Playfield, LM_Inserts_l25_UnderPF, LM_Inserts_l25_sw72, LM_Inserts_l26_Gate3, LM_Inserts_l26_Layer1, LM_Inserts_l26_Parts, LM_Inserts_l26_Playfield, LM_Inserts_l26_UnderPF, LM_Inserts_l27_Layer1, LM_Inserts_l27_Parts, LM_Inserts_l27_Playfield, LM_Inserts_l27_UnderPF, LM_Inserts_l28_Layer1, LM_Inserts_l28_Parts, LM_Inserts_l28_Plastics, LM_Inserts_l28_Playfield, LM_Inserts_l28_UnderPF, LM_Inserts_l29_Playfield, LM_Inserts_l29_UnderPF, LM_Inserts_l3_LF, LM_Inserts_l3_Playfield, LM_Inserts_l3_UnderPF, LM_Inserts_l30_Playfield, LM_Inserts_l30_RFU, LM_Inserts_l30_UnderPF, LM_Inserts_l31_Playfield, LM_Inserts_l31_UnderPF, LM_Inserts_l32_Playfield, LM_Inserts_l32_UnderPF, LM_Inserts_l33_LFU, LM_Inserts_l33_Playfield, LM_Inserts_l33_UnderPF, _
  LM_Inserts_l34_Playfield, LM_Inserts_l34_UnderPF, LM_Inserts_l35_Playfield, LM_Inserts_l35_UnderPF, LM_Inserts_l36_Playfield, LM_Inserts_l36_UnderPF, LM_Inserts_l37_Playfield, LM_Inserts_l37_UnderPF, LM_Inserts_l38_Playfield, LM_Inserts_l38_UnderPF, LM_Inserts_l39_Playfield, LM_Inserts_l39_UnderPF, LM_Inserts_l4_Parts, LM_Inserts_l4_Plastics, LM_Inserts_l4_Playfield, LM_Inserts_l4_UnderPF, LM_Inserts_l40_Playfield, LM_Inserts_l40_UnderPF, LM_Inserts_l41_Playfield, LM_Inserts_l41_UnderPF, LM_Inserts_l42_Playfield, LM_Inserts_l42_UnderPF, LM_Inserts_l43_Playfield, LM_Inserts_l43_UnderPF, LM_Inserts_l44_Lsling, LM_Inserts_l44_Lsling1, LM_Inserts_l44_Lsling2, LM_Inserts_l44_Parts, LM_Inserts_l44_Plastics, LM_Inserts_l44_Playfield, LM_Inserts_l44_UnderPF, LM_Inserts_l45_Gate3, LM_Inserts_l45_Gate4_Wire, LM_Inserts_l45_Layer1, LM_Inserts_l45_Parts, LM_Inserts_l45_Playfield, LM_Inserts_l45_UnderPF, LM_Inserts_l45_sw41, LM_Inserts_l46_Layer1, LM_Inserts_l46_Parts, LM_Inserts_l46_Playfield, LM_Inserts_l46_UnderPF, _
  LM_Inserts_l46_sw51, LM_Inserts_l47_Gate2, LM_Inserts_l47_Layer1, LM_Inserts_l47_Parts, LM_Inserts_l47_Plastics, LM_Inserts_l47_Playfield, LM_Inserts_l47_UnderPF, LM_Inserts_l48_Parts, LM_Inserts_l48_Plastics, LM_Inserts_l48_Playfield, LM_Inserts_l48_Rsling, LM_Inserts_l48_Rsling1, LM_Inserts_l48_Rsling2, LM_Inserts_l48_UnderPF, LM_Inserts_l49_Playfield, LM_Inserts_l49_UnderPF, LM_Inserts_l5_Layer1, LM_Inserts_l5_Parts, LM_Inserts_l5_Plastics, LM_Inserts_l5_Playfield, LM_Inserts_l5_UnderPF, LM_Inserts_l5_sw40, LM_Inserts_l50_Playfield, LM_Inserts_l50_UnderPF, LM_Inserts_l51_Bumper_2_Ring, LM_Inserts_l51_Bumper_2_Socket, LM_Inserts_l51_Parts, LM_Inserts_l51_Playfield, LM_Inserts_l51_UnderPF, LM_Inserts_l6_Layer1, LM_Inserts_l6_Parts, LM_Inserts_l6_Plastics, LM_Inserts_l6_Playfield, LM_Inserts_l6_UnderPF, LM_Inserts_l6_sw60, LM_Inserts_l7_Gate4_Wire, LM_Inserts_l7_Layer1, LM_Inserts_l7_Parts, LM_Inserts_l7_Plastics, LM_Inserts_l7_Playfield, LM_Inserts_l7_UnderPF)
' VLM  Arrays - End

' Update Log
' 1.0.1 - bthlonewolf - added DOF, mirroring existing table
' 1.0.2 - bthlonewolf - changed tilt switch from -7 to 57

'Hollywood Heat (Gottlieb 1986)
'https://www.ipdb.org/machine.cgi?id=1219


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
Const BallSize = 50   'Ball size must be 50
Const BallMass = 1    'Ball mass must be 1
Const tnob = 3      'Total number of balls on the playfield including captive balls.
Const lob = 1     'Total number of locked balls

Dim tablewidth: tablewidth = Table1.width
Dim tableheight: tableheight = Table1.height

'  Standard definitions
Const cGameName = "hlywoodh"    'PinMAME ROM name
Const UseSolenoids = 2      '1 = Normal Flippers, 2 = Fastflips
Const UseLamps = 1        '0 = Custom lamp handling, 1 = Built-in VPX handling (using light number in light timer)
Const UseSync = 0
Const HandleMech = 0
Const SSolenoidOn = ""      'Sound sample used for this, obsolete.
Const SSolenoidOff = ""     ' ^
Const SFlipperOn = ""     ' ^
Const SFlipperOff = ""      ' ^
Const SCoin = ""        ' ^

Dim VRRoom, cab_mode, DesktopMode: DesktopMode = Table1.ShowDT
DesktopMode = Table1.ShowDT
If RenderingMode = 2 Then VRRoom=1 Else VRRoom=0      'VRRoom set based on RenderingMode starting in version 10.72
if Not DesktopMode and VRRoom=0 Then cab_mode=1 Else cab_mode=0

'Const UseVPMModSol = 1   'Old PWM method. Don't use this
Const UseVPMModSol = 0    'Set to 2 for PWM flashers, inserts, and GI. Requires VPinMame 3.6

'NOTES on UseVPMModSol = 2:
'  - Only supported for S9/S11/DataEast/WPC/Capcom/Whitestar (Sega & Stern)/SAM
'  - All lights on the table must have their Fader model set tp "LED (None)" to get the correct fading effects
'  - When not supported VPM outputs only 0 or 1. Therefore, use VPX "Incandescent" fader for lights

LoadVPM "03060000", "sys80.VBS", 3.02  'The "03060000" argument forces user to have VPinMame 3.6
'LoadVPM "01210000", "sys80.VBS", 3.1

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
  'BSUpdate
  UpdateBallBrightness
  RollingUpdate       'Update rolling sounds
  DoSTAnim          'Standup target animations
  UpdateStandupTargets
  DoDTAnim          'Drop target animations
  UpdateDropTargets

    If VRRoom = 0 and cab_mode = 0 Then: DisplayTimer: End If
  If VRRoom=1 Then: VRDisplayTimer: End If
End Sub

'The CorTimer interval should be 10. It's sole purpose is to update the Cor (physics) calculations
CorTimer.Interval = 10
Sub CorTimer_Timer(): Cor.Update: End Sub

'DispTimer

'******************************************************
'  ZINI: Table Initialization and Exiting
'******************************************************
Dim HHBall1, HHBall2, HHBall3, HHCaptiveBall, gBOT, bslLock, bsrLock

gBot = GetBalls

Sub Table1_Init
  vpminit me
  With Controller
    .GameName = cGameName
    If Err Then MsgBox "Can't start Game " & cGameName & vbNewLine & Err.Description:Exit Sub
    .SplashInfoLine = "Hollywood Heat (Gottlieb 1986)" & vbNewLine & "VPW"
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

  'Kickers

    Set bslLock=New cvpmBallStack
    with bslLock
        .InitSaucer sw46,46,-170,10
        .InitExitSnd Soundfx("fx_ballrel",DOFContactors), Soundfx("HoleKick",DOFContactors)
    end with

    Set bsrLock=New cvpmBallStack
    with bsrLock
        .InitSaucer sw56,56,120,11
        .InitExitSnd Soundfx("fx_ballrel",DOFContactors), Soundfx("HoleKick",DOFContactors)
    end with

  'Nudging
  vpmNudge.TiltSwitch=57 '-7?
  vpmNudge.Sensitivity=4
  vpmNudge.TiltObj = Array(Bumper1, LeftSlingshot, RightSlingshot, TopSlingShot, URightSlingShot)

' 'Main Timer init
' PinMAMETimer.Interval = PinMAMEInterval
' PinMAMETimer.Enabled = 1

  'Trough - Creates a ball in the kicker switch and gives that ball used an individual name.
  Set HHBall1 = sw76.CreateSizedballWithMass(Ballsize/2,Ballmass)
  Set HHBall2 = sw20.CreateSizedballWithMass(Ballsize/2,Ballmass)
  Set HHBall3 = sw10.CreateSizedballWithMass(Ballsize/2,Ballmass)

  'Forces the trough switches to "on" at table boot so the game logic knows there are balls in the trough.
  Controller.Switch(76) = 1
  Controller.Switch(20) = 1
  Controller.Switch(10) = 1

  '***Captive Ball Creation
  Set HHCaptiveBall = captiveKick.CreateSizedballWithMass(Ballsize/2,Ballmass)
  vpmTimer.AddTimer 300, "captivekick.kick 180,1 '"   'Creates a ball from the "captiveball" kicker
  vpmTimer.AddTimer 310, "captivekick.enabled= 0 '"   'Permenantly Disables Captive Ball Kicker

  '***Setting up a ball array (collection), must contain all the balls you create on the table.
  gBOT = Array(HHCaptiveBall,HHBall1,HHBall2,HHBall3)

End Sub

Sub Table_Paused:Controller.Pause = 1:End Sub
Sub Table_unPaused:Controller.Pause = 0:End Sub
Sub Table_exit()
  Controller.Pause = False
  Controller.Stop
End Sub

'*****************************************************************************************************************************************
'  ERROR LOGS by baldgeek
'*****************************************************************************************************************************************

' Log File Usage:
'   WriteToLog "Label 1", "Message 1 "
'   WriteToLog "Label 2", "Message 2 "

Const KeepLogs = false

Class DebugLogFile

    Private Filename
    Private TxtFileStream

    Private Function LZ(ByVal Number, ByVal Places)
        Dim Zeros
        Zeros = String(CInt(Places), "0")
        LZ = Right(Zeros & CStr(Number), Places)
    End Function

    Private Function GetTimeStamp
        Dim CurrTime, Elapsed, MilliSecs
        CurrTime = Now()
        Elapsed = Timer()
        MilliSecs = Int((Elapsed - Int(Elapsed)) * 1000)
        GetTimeStamp = _
            LZ(Year(CurrTime),   4) & "-" _
            & LZ(Month(CurrTime),  2) & "-" _
            & LZ(Day(CurrTime),    2) & " " _
            & LZ(Hour(CurrTime),   2) & ":" _
            & LZ(Minute(CurrTime), 2) & ":" _
            & LZ(Second(CurrTime), 2) & ":" _
            & LZ(MilliSecs, 4)
    End Function

' *** Debug.Print the time with milliseconds, and a message of your choice
    Public Sub WriteToLog(label, message, code)
        Dim FormattedMsg, Timestamp
        'Filename = UserDirectory + "\" + cGameName + "_debug_log.txt"
    Filename = cGameName + "_debug_log.txt"

        Set TxtFileStream = CreateObject("Scripting.FileSystemObject").OpenTextFile(Filename, code, True)
        Timestamp  = GetTimeStamp
        FormattedMsg = GetTimeStamp + " : " + label + " : " + message
        TxtFileStream.WriteLine FormattedMsg
        TxtFileStream.Close
    debug.print label & " : " & message
  End Sub

End Class

Sub WriteToLog(label, message)
  if KeepLogs Then
    Dim LogFileObj
    Set LogFileObj = New DebugLogFile
    LogFileObj.WriteToLog label, message, 8
  end if
End Sub

Sub NewLog()
  if KeepLogs Then
    Dim LogFileObj
    Set LogFileObj = New DebugLogFile
    LogFileObj.WriteToLog "NEW LOG", " ", 2
  end if
End Sub

'*******************************************
'  ZOPT: User Options
'*******************************************

Dim LightLevel : LightLevel = 0.25        ' Level of room lighting (0 to 1), where 0 is dark and 100 is brightest
Dim ColorLUT : ColorLUT = 1           ' Color desaturation LUTs: 1 to 11, where 1 is normal and 11 is black'n'white
Dim VolumeDial : VolumeDial = 0.8             ' Overall Mechanical sound effect volume. Recommended values should be no greater than 1.
Dim BallRollVolume : BallRollVolume = 0.5     ' Level of ball rolling volume. Value between 0 and 1
Dim RampRollVolume : RampRollVolume = 0.5     ' Level of ramp rolling volume. Value between 0 and 1
'Dim StagedFlippers : StagedFlippers = 0         ' Staged Flippers. 0 = Disabled, 1 = Enabled



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

  'GI Color
  v = Table1.Option("GI Lights", 0, 1, 1, 1, 0, Array("Classic", "Miami"))
  SetGIColor v
  SetGISlingColor v

  ' Color Saturation
    ColorLUT = Table1.Option("Color Saturation", 1, 12, 1, 1, 0, _
    Array("Normal", "Dark Contrast", "Desaturated 10%", "Desaturated 20%", "Desaturated 30%", "Desaturated 40%", "Desaturated 50%", _
        "Desaturated 60%", "Desaturated 70%", "Desaturated 80%", "Desaturated 90%", "Black 'n White"))
  if ColorLUT = 1 Then Table1.ColorGradeImage = ""
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

    ' Toggle Reflections
  v = Table1.Option("Playfield Reflections", 0, 2, 1, 1, 0, Array("None", "Partial", "Full"))
  ReflectionToggle(v)

    ' Sound volumes
    VolumeDial = Table1.Option("Mech Volume", 0, 1, 0.01, 0.8, 1)
    BallRollVolume = Table1.Option("Ball Roll Volume", 0, 1, 0.01, 0.5, 1)
  RampRollVolume = Table1.Option("Ramp Roll Volume", 0, 1, 0.01, 0.5, 1)

  ' Room brightness
' LightLevel = Table1.Option("Table Brightness (Ambient Light Level)", 0, 1, 0.01, .5, 1)
  LightLevel = NightDay/100
  SetRoomBrightness LightLevel   'Uncomment this line for lightmapped tables.

'    ' Staged Flippers
'    StagedFlippers = Table1.Option("Staged Flippers", 0, 1, 1, 0, 0, Array("Disabled", "Enabled"))

    If eventId = 3 And dspTriggered Then dspTriggered = False : DisableStaticPreRendering = False : End If
End Sub

Function ReflectionToggle(state)
  Dim ReflObjectsArray: ReflObjectsArray = Array(BP_Parts, BP_LFlip_001, BP_LFlipU, BP_LFlip1_001, BP_LFlip1U, BP_Rflip_001, BP_RFlipU,BP_Rflip1_001, BP_RFlip1U, BP_sw40, BP_sw41, BP_sw43_001, BP_sw44_001, BP_sw50, BP_sw51, BP_sw53_001, BP_sw54_001, BP_sw60, BP_sw61, BP_sw63_001, BP_sw64_001, BP_sw73a_001)
  Dim PartialReflObjectsArray: PartialReflObjectsArray = Array(BM_Parts, BM_LFlip_001, BM_LFlipU, BM_LFlip1_001, BM_LFlip1U, BM_Rflip_001, BM_RFlipU, BM_Rflip1_001, BM_RFlip1U, BM_sw40, BM_sw41, BM_sw43_001, BM_sw44_001, BM_sw50, BM_sw51, BM_sw53_001, BM_sw54_001, BM_sw60, BM_sw61, BM_sw63_001, BM_sw64_001, BM_sw73a_001)

  Dim IBP, BP, v1, v2
  Select Case state
    Case 0: v1 = false: v2 = false  'No reflections
    Case 1: v1 = false: v2 = true   'Only reflect static prims
    Case 2: v1 = true:  v2 = true   'Reflect everything (that matters as defined in ReflObjectsArray)
  End Select

  For Each IBP in ReflObjectsArray
    For Each BP in IBP
      BP.ReflectionEnabled = v1
    Next
  Next

  For Each BP in PartialReflObjectsArray
    BP.ReflectionEnabled = v2
  Next

  BM_Playfield.ReflectionEnabled = False
End Function

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
  Dim a : a = LeftFlipper.CurrentAngle
  FlipperLSh.RotZ = a

  Dim v, BP
  v = 255.0 * (120.0 - LeftFlipper.CurrentAngle) / (120.0 -  70.0)

  For each BP in BP_Lflip_001
    BP.Rotz = a
    BP.visible = v < 128.0
  Next
  For each BP in BP_LflipU
    BP.Rotz = a
    BP.visible = v >= 128.0
  Next
End Sub

Sub RightFlipper_Animate
  Dim a : a = RightFlipper.CurrentAngle
  FlipperRSh.RotZ = a

  Dim v, BP
  v = 255.0 * (-120.0 - RightFlipper.CurrentAngle) / (-120.0 +  70.0)

  For each BP in BP_RFlip_001
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
' FlipperRSh001.RotZ = a

  Dim v, BP
  v = 255.0 * (-125.0 - RightFlipper1.CurrentAngle) / (-125.0 +  75.0)

  For each BP in BP_Rflip1_001
    BP.Rotz = a
    BP.visible = v < 128.0
  Next
  For each BP in BP_Rflip1U
    BP.Rotz = a
    BP.visible = v >= 128.0
  Next
End Sub

Sub LeftFlipper1_Animate
  Dim a : a = LeftFlipper1.CurrentAngle
' FlipperRSh001.RotZ = a

  Dim v, BP
  v = 255.0 * (125.0 - LeftFlipper1.CurrentAngle) / (125.0 -  75.0)

  For each BP in BP_LFlip1_001
    BP.Rotz = a
    BP.visible = v < 128.0
  Next
  For each BP in BP_LFlip1U
    BP.Rotz = a
    BP.visible = v >= 128.0
  Next
End Sub


''''' Switch Animations

Sub sw42_animate
  Dim BP, Z
    z = sw42.CurrentAnimOffset
  For each BP in BP_sw42 : BP.transy = z: Next
End Sub

Sub sw52_animate
  Dim BP, Z
    z = sw52.CurrentAnimOffset
  For each BP in BP_sw52 : BP.transy = z: Next
End Sub

Sub sw62_animate
  Dim BP, Z
    z = sw62.CurrentAnimOffset
  For each BP in BP_sw62 : BP.transy = z: Next
End Sub


''''' Gate Animations

Sub Gate1_Animate
  Dim a : a = Gate1.CurrentAngle
  Dim BP : For Each BP in BP_Gate1_Wire : BP.rotx = a: Next
End Sub

Sub Gate3_Animate
  Dim a : a = Gate3.CurrentAngle
  Dim BP : For Each BP in BP_Gate3_Wire : BP.rotx = a: Next
End Sub

Sub Gate4_Animate
  Dim a : a = Gate4.CurrentAngle
  Dim BP : For Each BP in BP_Gate4_Wire : BP.rotx = a: Next
End Sub

Sub Gate5_Animate
  Dim a : a = Gate5.CurrentAngle
  Dim BP : For Each BP in BP_Gate5_Wire : BP.rotx = a: Next
End Sub

Sub Gate7_Animate
  Dim a : a = Gate7.CurrentAngle
  Dim BP : For Each BP in BP_Gate7_Wire : BP.rotx = a: Next
End Sub

Sub Gate8_Animate
  Dim a : a = Gate8.CurrentAngle
  Dim BP : For Each BP in BP_Gate8_Wire : BP.rotx = a: Next
End Sub

Sub Gate_Animate
  Dim a : a = Gate.CurrentAngle
  Dim BP : For Each BP in BP_Pgate : BP.rotx = a: Next
End Sub




'******************************************************
'   ZBBR: BALL BRIGHTNESS
'******************************************************

Const BallBrightness =  1       'Ball brightness - Value between 0 and 1 (0=Dark ... 1=Bright)

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

Dim cArray, sArray
cArray = Array(c4000k, cPink)
sArray = Array(c4000k, cGreen)

Sub SetGIColor(c)
  Dim xx, BL
  For each xx in GI: xx.color = cArray(c): xx.colorfull = cArray(c): Next
  For each BL in BL_GI_al6a: BL.color = cArray(c): Next
  For each BL in BL_GI_GI1: BL.color = cArray(c): Next
  For each BL in BL_GI_GI2: BL.color = cArray(c): Next
  For each BL in BL_GI_GI3: BL.color = cArray(c): Next
  For each BL in BL_GI_GI4: BL.color = cArray(c): Next
  For each BL in BL_GI_GI5: BL.color = cArray(c): Next
  For each BL in BL_GI_Llane1: BL.color = cArray(c): Next
  For each BL in BL_GI_Llane2: BL.color = cArray(c): Next
  For each BL in BL_GI_Llsling1: BL.color = cArray(c): Next
  For each BL in BL_GI_Llsling2: BL.color = cArray(c): Next
End Sub

Sub SetGISlingColor(c)
  Dim xx, BL
  For each xx in GISling: xx.color = sArray(c): xx.colorfull = sArray(c): Next
  For each BL in BL_GI_GI6: BL.color = sArray(c): Next
  For each BL in BL_GI_GI7: BL.color = sArray(c): Next
End Sub


'******************************************************
' ZKEY: Key Press Handling
'******************************************************

Sub Table1_KeyDown(ByVal keycode) '***What to do when a button is pressed***
  'If keycode = PlungerKey Then Plunger.Pullback:vpmTimer.PulseSw 31
  If KeyCode = PlungerKey Then
    Plunger.PullBack
    Plunger.Pullback:SoundPlungerPull()
    TimerVRPlunger.Enabled = True
    TimerVRPlunger1.Enabled = False
    VR_Plunger.Y = 20
  End If
    If Keycode = LeftFlipperKey Then
    VR_LeftFlipperButton.X = 0 + 6
  End if
    If Keycode = RightFlipperKey Then
    VR_RightFlipperButton.X = 0 - 6
  End if
  If Keycode = StartGameKey Then
    SoundStartButton
    VR_StartButton.y=VR_StartButton.y-4
  End if
  If keycode = LeftTiltKey Then Nudge 90, 1 : SoundNudgeLeft      ' Sets the nudge angle and power
  If keycode = RightTiltKey Then Nudge 270, 1 : SoundNudgeRight   ' ^
  If keycode = CenterTiltKey Then Nudge 0, 1 : SoundNudgeCenter   ' ^
  If keycode = StartGameKey Then SoundStartButton
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
    'If BIPL = 1 Then
    ' SoundPlungerReleaseBall()     'Plunger release sound when there is a ball in shooter lane
    'Else
    ' SoundPlungerReleaseNoBall()     'Plunger release sound when there is no ball in shooter lane
    'End If
    TimerVRPlunger.Enabled = False
        TimerVRPlunger1.Enabled = True
    VR_Plunger.Y = 20
  End If
  If Keycode = LeftFlipperKey Then
    VR_LeftFlipperButton.x = 0
  End if
  If keycode = RightFlipperKey Then
    VR_RightFlipperButton.X = 0
  End If
  If Keycode = StartGameKey Then
    SoundStartButton
    VR_StartButton.y=VR_StartButton.y+4
  End if
  If vpmKeyUp(keycode) Then Exit Sub
End Sub


' VR Plunger code

Sub TimerVRPlunger_Timer
  If VR_Plunger.Y < 130 then
       VR_Plunger.Y = VR_Plunger.Y + 5
  End If
End Sub

Sub TimerVRPlunger1_Timer
  VR_Plunger.Y = 20 + (5* Plunger.Position) -20
End Sub


'******************************************************
' ZSOL: Solenoids & Flashers
'******************************************************

SolCallback(1) =    "SolLLock"
SolCallback(2) =    "SolRLock"
SolCallback(4) =    "SolDrops2"       'Drops 2nd drop target on left and right banks
SolCallback(5) =    "SolResetDropsL"    'Reset Drop Targets on Left
SolCallback(6) =    "SolResetDropsR"    'Reset Drop Targets on Right
SolCallback(7) =    "SolDrops3"       'Drops 3rd drop target on left and right banks
SolCallback(8) =    "SolKnocker"
SolCallback(9) =    "SolRelease"
'SolCallback(10) = "FastFlips.TiltSol"

SolCallback(sLRFlipper) = "SolRFlipper"
SolCallback(sLLFlipper) = "SolLFlipper"


Sub SolKnocker(Enabled)
  If enabled Then
    KnockerSolenoid
  End If
End Sub

' paired drop target dropping in "In-Sync" mode

'sub dtDrop2(enabled):dtLeft.hit 2:dtRight.hit 2:end sub
'sub dtDrop3(enabled):dtLeft.hit 3:dtRight.hit 3:end sub

Dim GILevel, DayNight

Sub Intensity
  If DayNight <= 20 Then
      GILevel = .5
  ElseIf DayNight <= 40 Then
      GILevel = .4125
  ElseIf DayNight <= 60 Then
      GILevel = .325
  ElseIf DayNight <= 80 Then
      GILevel = .2375
  Elseif DayNight <= 100  Then
      GILevel = .15
  End If

  For each xx in GI: xx.IntensityScale = xx.IntensityScale * (GILevel): Next
  For each xx in GISling: xx.IntensityScale = xx.IntensityScale * (GILevel): Next
  For each xx in DTLights: xx.IntensityScale = xx.IntensityScale * (GILevel): Next

End Sub

Sub ALlightsTimer_timer

  if me.uservalue>1 then
      ALlights(me.uservalue-2).state=0
      ALlightsA(me.uservalue-2).state=0
    elseif me.uservalue=1 then
      ALlights(9).state=0
      ALlightsA(9).state=0
    else
      ALlights(8).state=0
      ALlightsA(8).state=0
  end if

  ALlights(me.uservalue).state=1
  ALlightsA(me.uservalue).state=1

  me.uservalue = me.uservalue+1
  if me.uservalue>9 then me.uservalue=0
End sub

Sub FlipperTimer_Timer

'testbox.text = Controller.Dip(3)
'testbox1.text =

  LFlip.RotY = LeftFlipper.CurrentAngle
  RFlip.RotY = RightFlipper.CurrentAngle
  LFlip1.RotY = LeftFlipper1.CurrentAngle-90
  RFlip1.RotY = RightFlipper1.CurrentAngle-90
  Pgate.Rotz = Gate.CurrentAngle*0.7
End Sub


' Ball locks / kickers

Sub sw46_Hit:SoundSaucerLock:bslLock.AddBall 0:End Sub
Sub sw46_UnHit: SoundSaucerKick 1, sw46: End Sub
Sub sw56_Hit:SoundSaucerLock:bsrLock.AddBall 0:End Sub
Sub sw56_UnHit: SoundSaucerKick 1, sw56: End Sub

Sub SolLLock(enabled)
  If enabled Then
    bslLock.ExitSol_On
    LeftKickTimer.uservalue = 0
    PkickarmL.RotZ = 15
    LeftKickTimer.Enabled = 1
  End If
End Sub

Sub LeftKickTimer_timer
  select case me.uservalue
    case 5:
    PkickarmL.rotz=0
    me.enabled=0
  end Select
  me.uservalue = me.uservalue+1
End Sub

Sub SolRLock(enabled)
  If enabled Then
    bsrLock.ExitSol_On
    RightKickTimer.uservalue = 0
    PkickarmR.RotZ = 15
    RightKickTimer.Enabled = 1
  End If
End Sub

Sub RightKickTimer_timer
  select case me.uservalue
    case 5:
    PkickarmR.rotz=0
    me.enabled=0
  end Select
  me.uservalue = me.uservalue+1
End Sub



'******************************************************
' ZDRN: Drain, Trough, and Ball Release
'******************************************************

'********************* TROUGH *************************

Sub sw76_Hit   : Controller.Switch(76) = 1 : UpdateTrough : End Sub
Sub sw76_UnHit : Controller.Switch(76) = 0 : UpdateTrough : End Sub
Sub sw20_Hit   : Controller.Switch(20) = 1 : UpdateTrough : End Sub
Sub sw20_UnHit : Controller.Switch(20) = 0 : UpdateTrough : End Sub
Sub sw10_Hit   : Controller.Switch(10) = 1 : UpdateTrough : End Sub
Sub sw10_UnHit : Controller.Switch(10) = 0 : UpdateTrough : End Sub
'Sub sw66_Hit   : Controller.Switch(10) = 1 : UpdateTrough : End Sub
'Sub sw66_UnHit : Controller.Switch(10) = 0 : UpdateTrough : End Sub

Sub UpdateTrough
  UpdateTroughTimer.Interval = 500
  UpdateTroughTimer.Enabled = 1
End Sub

Sub UpdateTroughTimer_Timer
  If sw76.BallCntOver = 0 Then sw20.kick 70, 5
  UpdateTroughTimer1.Interval = 500
  UpdateTroughTimer1.Enabled = 1
  'If sw20.BallCntOver = 0 Then sw10.kick 70, 5
  'If sw10.BallCntOver = 0 Then sw66.kick 70, 20
  Me.Enabled = 0
End Sub

Sub UpdateTroughTimer1_Timer()
  If sw20.BallCntOver = 0 Then sw10.kick 70, 5
  Me.Enabled = 0
End Sub

'*****************  DRAIN & RELEASE  ******************

Sub sw66_Hit()
  RandomSoundDrain sw66
  UpdateTrough
  Controller.Switch(66) = 1
End Sub

Sub sw66_UnHit()
  Controller.Switch(66) = 0
End Sub

Sub SolRelease(enabled)
  If enabled Then
    sw66.kick 70,40
    RandomSoundBallRelease sw76
  End If
End Sub

'set Lights(12) = L12
'
'Set LampCallback = GetRef("UpdateMultipleLamps")
'Sub UpdateMultipleLamps
'    if controller.lamp(12)=true and sw76.BallCntOver > 0 then
'   sw76.kick 60, 7
'   PlaySoundAt SoundFX("ballrelease",DOFContactors), sw76
'   UpdateTrough
'    end if
'
'    if controller.lamp(2)=true then dtleft.hit 1:dtRight.hit 1
'End Sub


Dim L12on: L12on = False
Sub L12_animate
  If L12.state > 0.5 and L12on = False Then
    debug.print "L12 ON"
    L12on = True
    sw76.kick 60, 7
    RandomSoundBallRelease sw76
    'PlaySoundAt SoundFX("ballrelease",DOFContactors), sw76
    UpdateTrough
  ElseIf L12.state < 0.4 and L12on = True Then
    debug.print "L12 OFF"
    L12on = False
  End If
End Sub


Dim L2on: L2on = False
Sub L2_animate
  If L2.state > 0.5 and L2on = False Then
    debug.print "L2 ON"
    L2on = True
    SolDrops1 True
  ElseIf L2.state < 0.4 and L2on = True Then
    debug.print "L2 OFF"
    L2on = False
    SolDrops1 False
  End If
End Sub


'******************************************************
' ZFLP: FLIPPERS
'******************************************************
Const ReflipAngle = 20

Sub SolLFlipper(Enabled)
  If Enabled Then
        FlipperActivate LeftFlipper, LFPress
    FlipperActivate LeftFlipper1, LFPress1
    LF.fire
    LF1.fire
    If Leftflipper.currentangle < Leftflipper.endangle + ReflipAngle Then
      RandomSoundReflipUpLeft LeftFlipper
    Else
      SoundFlipperUpAttackLeft LeftFlipper
      RandomSoundFlipperUpLeft LeftFlipper
    End If
    If Leftflipper1.currentangle < Leftflipper1.endangle + ReflipAngle Then
      RandomSoundReflipUpLeft LeftFlipper1
    Else
      SoundFlipperUpAttackLeft LeftFlipper1
      RandomSoundFlipperUpLeft LeftFlipper1
    End If
    controller.Switch(6)=1
  Else
    FlipperDeActivate LeftFlipper, LFPress
    FlipperDeactivate LeftFlipper1, LFPress1
    LeftFlipper.RotateToStart
    LeftFlipper1.RotateToStart
    If LeftFlipper.currentangle < LeftFlipper.startAngle - 5 Then
      RandomSoundFlipperDownLeft LeftFlipper
    End If
    If LeftFlipper1.currentangle < LeftFlipper1.startAngle - 5 Then
      RandomSoundFlipperDownLeft LeftFlipper1
    End If
    FlipperLeftHitParm = FlipperUpSoundLevel
    controller.Switch(6)=0
  End If
End Sub

Sub SolRFlipper(Enabled)
  If Enabled Then
        FlipperActivate RightFlipper, RFPress
    FlipperActivate RightFlipper1, RFPress
    RF.fire
    RF1.fire
    If rightflipper.currentangle > rightflipper.endangle - ReflipAngle Then
      RandomSoundReflipUpRight RightFlipper
    Else
      SoundFlipperUpAttackRight RightFlipper
      RandomSoundFlipperUpRight RightFlipper
    End If
    If rightflipper1.currentangle > rightflipper1.endangle - ReflipAngle Then
      RandomSoundReflipUpRight RightFlipper1
    Else
      SoundFlipperUpAttackRight RightFlipper1
      RandomSoundFlipperUpRight RightFlipper1
    End If
    controller.Switch(16)=1
    controller.Switch(71)=1
  Else
    FlipperDeActivate RightFlipper, RFPress
    FlipperDeactivate RightFlipper1, RFPress1
    RightFlipper.RotateToStart
    RightFlipper1.RotateToStart
    If RightFlipper.currentangle > RightFlipper.startAngle + 5 Then
      RandomSoundFlipperDownRight RightFlipper
    End If
    If RightFlipper1.currentangle > RightFlipper1.startAngle + 5 Then
      RandomSoundFlipperDownRight RightFlipper1
    End If
    FlipperRightHitParm = FlipperUpSoundLevel
    controller.Switch(16)=0
    controller.Switch(71)=0
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
  CheckLiveCatch Activeball, LeftFlipper1, LFCount1, parm
  LF1.ReProcessBalls ActiveBall
  LeftFlipperCollide parm
End Sub

Sub RightFlipper1_Collide(parm)
  CheckLiveCatch Activeball, RightFlipper1, RFCount1, parm
  RF1.ReProcessBalls ActiveBall
  RightFlipperCollide parm
End Sub



'************************************************************
' ZSLG: Slingshot Animations
'************************************************************
Dim LStep : LStep = 4 : LeftSlingShot_Timer
Dim RStep : RStep = 4 : RightSlingShot_Timer
'Dim TStep : TStep = 4 : TopSlingShot_Timer
Dim URStep : URStep = 4 : URightSlingShot_Timer

Sub RightSlingShot_Slingshot
  RS.VelocityCorrect(Activeball)
  vpmTimer.PulseSw(32)            'Sling Switch Number
  BM_rsling1.Visible = 1
  SlingR.TransY =  - 20             'Sling Metal Bracket
  RStep = 0
  RightSlingShot.TimerEnabled = 1
  RightSlingShot.TimerInterval = 10
  RandomSoundSlingshotRight slingR
End Sub

Sub RightSlingShot_Timer
  Select Case RStep
    Case 3
    BM_rsling1.Visible = 0
    BM_rsling2.Visible = 1
    SlingR.TransY =  - 10
    Case 4
    BM_rsling2.Visible = 0
    SlingR.TransY = 0
    RightSlingShot.TimerEnabled = 0
  End Select
  RStep = RStep + 1
End Sub

Sub LeftSlingShot_Slingshot
  LS.VelocityCorrect(Activeball)
  vpmTimer.PulseSw(32)            'Sling Switch Number
  BM_lsling1.Visible = 1
  SlingL.TransY =  - 20             'Sling Metal Bracket
  LStep = 0
  LeftSlingShot.TimerEnabled = 1
  LeftSlingShot.TimerInterval = 10
  RandomSoundSlingshotLeft slingL
End Sub

Sub LeftSlingShot_Timer
  Select Case LStep
    Case 3
    BM_lsling1.Visible = 0
    BM_lsling2.Visible = 1
    SlingL.TransY =  - 10
    Case 4
    BM_lsling2.Visible = 0
    SlingL.TransY = 0
    LeftSlingShot.TimerEnabled = 0
  End Select
  LStep = LStep + 1
End Sub

Sub URightSlingShot_Slingshot
  URS.VelocityCorrect(Activeball)
  vpmTimer.PulseSw(32)            'Sling Switch Number
  BM_URsling1.Visible = 1
  SlingUR.TransY =  - 20              'Sling Metal Bracket
  URStep = 0
  URightSlingShot.TimerEnabled = 1
  URightSlingShot.TimerInterval = 10
  RandomSoundSlingshotRight slingUR
End Sub

Sub URightSlingShot_Timer
  Select Case URStep
    Case 3
    BM_URsling1.Visible = 0
    BM_URsling2.Visible = 1
    SlingUR.TransY =  - 10
    Case 4
    BM_URsling2.Visible = 0
    SlingUR.TransY = 0
    URightSlingShot.TimerEnabled = 0
  End Select
  URStep = URStep + 1
End Sub

'Sub TopSlingShot_Slingshot
' TS.VelocityCorrect(Activeball)
' vpmTimer.PulseSw(32)            'Sling Switch Number
' TSling1.Visible = 1
' SlingT.TransY =  - 20             'Sling Metal Bracket
' TStep = 0
' TopSlingShot.TimerEnabled = 1
' TopSlingShot.TimerInterval = 10
' RandomSoundSlingshotLeft slingT
'End Sub

'Sub TopSlingShot_Timer
' Select Case TStep
'   Case 3
'   TSLing1.Visible = 0
'   TSLing2.Visible = 1
'   SlingT.TransY =  - 10
'   Case 4
'   TSLing2.Visible = 0
'   SlingT.TransY = 0
'   TopSlingShot.TimerEnabled = 0
' End Select
' TStep = TStep + 1
'End Sub



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
Dim URS
Set URS = New SlingshotCorrection
Dim TS
Set TS = New SlingshotCorrection

InitSlingCorrection

Sub InitSlingCorrection
  LS.Object = LeftSlingshot
  LS.EndPoint1 = EndPoint1LS
  LS.EndPoint2 = EndPoint2LS

  RS.Object = RightSlingshot
  RS.EndPoint1 = EndPoint1RS
  RS.EndPoint2 = EndPoint2RS

  URS.Object = URightSlingshot
  URS.EndPoint1 = UEndPoint1RS
  URS.EndPoint2 = UEndPoint2RS

  TS.Object = TopSlingshot
  TS.EndPoint1 = TEndPoint1RS
  TS.EndPoint2 = TEndPoint2RS

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
  a = Array(LS, RS, URS, TS)
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



'************************************************************
' ZSWI: SWITCHES
'************************************************************


'************************* Bumpers **************************
Sub Bumper1_Hit(): RandomSoundBumperTop Bumper1: vpmTimer.PulseSw 44: End Sub


'************************ Rollovers *************************
Sub sw31_Hit(): Controller.Switch(31) = 1: End Sub              'Plunger Lane Rollover Switch
Sub sw31_UnHit: Controller.Switch(31) = 0: End Sub

Sub sw42_Hit(): Controller.Switch(42) = 1: End Sub              'Top Left Rollover Switch
Sub sw42_UnHit: Controller.Switch(42) = 0: End Sub
Sub sw52_Hit(): Controller.Switch(52) = 1: End Sub              'Top Middle Rollover Switch
Sub sw52_UnHit: Controller.Switch(52) = 0: End Sub
Sub sw62_Hit(): Controller.Switch(62) = 1: End Sub              'Top Right Rollover Switch
Sub sw62_UnHit: Controller.Switch(62) = 0: End Sub

Sub sw45_Hit(): Controller.Switch(45) = 1: End Sub              'Left Outlane Rollover Switch
Sub sw45_UnHit: Controller.Switch(45) = 0: End Sub
Sub sw55_Hit(): Controller.Switch(55) = 1: End Sub              'Left Inlane Rollover Switch
Sub sw55_UnHit: Controller.Switch(55) = 0: End Sub
Sub sw65_Hit(): Controller.Switch(65) = 1: End Sub              'Right Inlane Rollover Switch
Sub sw65_UnHit: Controller.Switch(65) = 0: End Sub
Sub sw75_Hit(): Controller.Switch(75) = 1: End Sub              'Right Outlane Rollover Switch
Sub sw75_UnHit: Controller.Switch(75) = 0: End Sub

Sub sw73_Hit(): Controller.Switch(73) = 1: End Sub              'Middle Left Rollover Switch
Sub sw73_UnHit: Controller.Switch(73) = 0: End Sub
Sub sw70_Hit(): Controller.Switch(70) = 1: End Sub              'Middle Right Rollover Switch
Sub sw70_UnHit: Controller.Switch(70) = 0: End Sub


'********************** Gate Switches ***********************
Sub sw72_Hit(): Controller.Switch(72) = 1: End Sub    'R Orbit Gate
Sub sw72_UnHit: Controller.Switch(72) = 0: End Sub
Sub sw74_Hit(): Controller.Switch(74) = 1: End Sub    'Special Hook Gate
Sub sw74_UnHit: Controller.Switch(74) = 0: End Sub


'********************* Standup Targets **********************
Sub Sw43_Hit: STHit 43: End Sub
Sub Sw53_Hit: STHit 53: End Sub
Sub Sw63_Hit: STHit 63: End Sub
Sub Sw44_Hit: STHit 44: End Sub
Sub Sw54_Hit: STHit 54: End Sub
Sub Sw64_Hit: STHit 64: End Sub
Sub Sw73a_Hit: STHit 73: End Sub


'********************** Drop Targets ************************
Sub Sw40_Hit: DTHit 40: TargetBouncer Activeball, 1: End Sub
Sub Sw50_Hit: DTHit 50: TargetBouncer Activeball, 1: End Sub
Sub Sw60_Hit: DTHit 60: TargetBouncer Activeball, 1: End Sub


Sub Sw41_Hit: DTHit 41: TargetBouncer Activeball, 1: End Sub
Sub Sw51_Hit: DTHit 51: TargetBouncer Activeball, 1: End Sub
Sub Sw61_Hit: DTHit 61: TargetBouncer Activeball, 1: End Sub


Sub SolDrops1(enabled)
  'debug.print "SolDrops1 "&enabled
  if enabled then
    DTDrop 40
    DTDrop 41
  end if
End Sub

Sub SolDrops2(enabled)
  'debug.print "SolDrops2 "&enabled
  if enabled then
    DTDrop 50
    DTDrop 51
  end if
End Sub

Sub SolDrops3(enabled)
  'debug.print "SolDrops3 "&enabled
  if enabled then
    DTDrop 60
    DTDrop 61
  end if
End Sub

Sub SolResetDropsL(enabled)
  'debug.print "SolResetDropsL "&enabled
  if enabled then
    RandomSoundDropTargetReset sw50p
    DTRaise 40
    DTRaise 50
    DTRaise 60
  end if
End Sub

Sub SolResetDropsR(enabled)
  'debug.print "SolResetDropsR "&enabled
  if enabled then
    RandomSoundDropTargetReset sw51p
    DTRaise 41
    DTRaise 51
    DTRaise 61
  end if
End Sub



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


'Ramp triggers
Sub RedRampStart_Hit
  WireRampOn true
End Sub

Sub RedRampStop1_Hit
    WireRampOff
End Sub

Sub HotShotStart_Hit
  WireRampOn false
End Sub

Sub HotShotStop_Hit
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
  'debug.print tableobj.name
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
Dim LF1 : Set LF1 = New FlipperPolarity
Dim RF1 : Set RF1 = New FlipperPolarity

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
Sub InitPolarity()
   dim x, a : a = Array(LF, RF, LF1, RF1)
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
    LF1.SetObjects "LF1", LeftFlipper1, TriggerLF1
    RF1.SetObjects "RF1", RightFlipper1, TriggerRF1

End Sub
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
  FlipperTricks LeftFlipper1, LFPress1, LFCount1, LFEndAngle1, LFState1
  FlipperTricks RightFlipper1, RFPress1, RFCount1, RFEndAngle1, RFState1
  FlipperNudge RightFlipper, RFEndAngle, RFEOSNudge, LeftFlipper, LFEndAngle
  FlipperNudge LeftFlipper, LFEndAngle, LFEOSNudge,  RightFlipper, RFEndAngle
End Sub

Dim LFEOSNudge, RFEOSNudge

Sub FlipperNudge(Flipper1, Endangle1, EOSNudge1, Flipper2, EndAngle2)
  Dim b
  'Dim gBOT
  'gBOT = GetBalls

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

Dim LFPress1, LFCount1, LFEndAngle1, LFState1
Dim RFPress1, RFCount1, RFEndAngle1, RFState1
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
Const EOSReturn = 0.045  'late 70's to mid 80's
' Const EOSReturn = 0.035  'mid 80's to early 90's
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
    Dim b', gBOT
    'gBOT = GetBalls

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
Dim DT40, DT50, DT60, DT41, DT51, DT61

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
Set DT41 = (new DropTarget)(sw41, sw41a, BM_sw41, 41, 0, false)
Set DT51 = (new DropTarget)(sw51, sw51a, BM_sw51, 51, 0, false)
Set DT61 = (new DropTarget)(sw61, sw61a, BM_sw61, 61, 0, false)

Dim DTArray
DTArray = Array(DT40, DT50, DT60, DT41, DT51, DT61)

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

  tz = BM_sw41.transz
  rx = BM_sw41.rotx
  ry = BM_sw41.roty
  For each BP in BP_sw41: BP.transz = tz: BP.rotx = rx: BP.roty = ry: Next

  tz = BM_sw51.transz
  rx = BM_sw51.rotx
  ry = BM_sw51.roty
  For each BP in BP_sw51: BP.transz = tz: BP.rotx = rx: BP.roty = ry: Next

  tz = BM_sw61.transz
  rx = BM_sw61.rotx
  ry = BM_sw61.roty
  For each BP in BP_sw61: BP.transz = tz: BP.rotx = rx: BP.roty = ry: Next
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
Dim ST43, ST53, ST63, ST44, ST54, ST64, ST73a

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

Set ST43 = (new StandupTarget)(sw43, BM_sw43_001, 43, 0)
Set ST53 = (new StandupTarget)(sw53, BM_sw53_001, 53, 0)
Set ST63 = (new StandupTarget)(sw63, BM_sw63_001, 63, 0)
Set ST44 = (new StandupTarget)(sw44, BM_sw44_001, 44, 0)
Set ST54 = (new StandupTarget)(sw54, BM_sw54_001, 54, 0)
Set ST64 = (new StandupTarget)(sw64, BM_sw64_001, 64, 0)
Set ST73a = (new StandupTarget)(sw73a, BM_sw73a_001, 73, 0)

'Add all the Stand-up Target Arrays to Stand-up Target Animation Array
' STAnimationArray = Array(ST1, ST2, ....)
Dim STArray
STArray = Array(ST43, ST53, ST63, ST44, ST54, ST64, ST73a)

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

    ty = BM_sw43_001.transy
  For each BP in BP_sw43_001 : BP.transy = ty: Next

    ty = BM_sw53_001.transy
  For each BP in BP_sw53_001 : BP.transy = ty: Next

    ty = BM_sw63_001.transy
  For each BP in BP_sw63_001 : BP.transy = ty: Next

    ty = BM_sw44_001.transy
  For each BP in BP_sw44_001 : BP.transy = ty: Next

    ty = BM_sw54_001.transy
  For each BP in BP_sw54_001 : BP.transy = ty: Next

    ty = BM_sw64_001.transy
  For each BP in BP_sw64_001 : BP.transy = ty: Next

    ty = BM_sw73a_001.transy
  For each BP in BP_sw73a_001 : BP.transy = ty: Next


End Sub


'******************************************************
'***  END STAND-UP TARGETS
'******************************************************


'******************************************************
'   ZBOU: VPW TargetBouncer for targets and posts by Iaakki, Wrd1972, Apophis
'******************************************************

Const TargetBouncerEnabled = 1    '0 = normal standup targets, 1 = bouncy targets
Const TargetBouncerFactor = 1  'Level of bounces. Recommmended value of 0.7-1

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
dim gilvl

'**** These are just debug commands. They emit same calls as what SolCallback will do. You may use these from debugger.
'Sub GIOn  : SetRelayGI 0: End Sub
'Sub GIOff : SetRelayGI 1: End Sub

'*********************************
'GI uselamps workaround by nfozzy
'*********************************

dim GIlamps : set GIlamps = New GIcatcherobject
Class GIcatcherObject   'object that disguises itself as a light. (UseLamps workaround for System80 GI circuit)
    Public Property Let State(input)
        dim x
        if input = 1 then 'If GI switch is engaged, turn off GI.
            for each x in gi : x.state = 0 : next
        elseif input = 0 then
            for each x in gi : x.state = 1 : next
        end if
        'tb.text = "gitcatcher.state = " & input    'debug
    End Property
End Class


set Lights(1) = GIlamps 'GI circuit

'**** SetRelayGI is called from SolCallback.
'sub SetRelayGI(aLvl)
' debug.print "SetRelayGI value: " & aLvl

  'Some tables have this solenoid reversed, i.e. Sega and Data East GI lights are off when GI relay is on.
  'However, this should be already compensated for in VPinMame 3.6

  ' Update the state for each GI light. The state will be a float value between 0 and 1.
' Dim bulb: For Each bulb in GI: bulb.State = aLvl: Next

  ' If the GI has an associated Relay sound, this can be played
' If aLvl >= 0.5 And gilvl < 0.5 Then
'   Sound_GI_Relay 1, Bumper1 'Note: Bumper1 is just used for sound positioning. Can be anywhere that makes sense.
' ElseIf aLvl <= 0.4 And gilvl > 0.4 Then
'   Sound_GI_Relay 0, Bumper1
' End If

  'You may add any other GI related effects here. Like if you want to make some toy to appear more bright, set it like this:
  'Primitive001.blenddisablelighting = 1.2 * aLvl + 0.2 'This will result to DL brightness between 0.2 - 1.4 for ON/OFF states

  'Pop Bumper Lights (PWM Following)
  'FlFadeBumper 1,aLvl
  'FlFadeBumper 2,aLvl
  'FlFadeBumper 3,aLvl

' gilvl = aLvl    'Storing the latest GI fading state into global variable, so one can use it elsewhere too.
'End Sub


'**** Use these for GI strings and stepped GI, and comment out the SetRelayGI
'**** This example table uses Relay for GI control, we don't need these at all

'Set GICallback  = GetRef("GIUpdates")    'use this for non-modulated GI
'Set GICallback2 = GetRef("GIUpdates2")   'use this for stepped/modulated GI

'GIupdates2 is called always when some event happens to GI channel.
'Sub GIUpdates2(aNr, aLvl)
' debug.print "GIUpdates2 nr: " & aNr & " value: " & aLvl

  'Select Case aNr 'Strings are selected here

    'Case 0:  'GI String 0

      ' Update the state for each GI light. The state will be a float value between 0 and 1.
      'Dim bulb: For Each bulb in GI: bulb.State = aLvl: Next

      ' If the GI has an associated Relay sound, this can be played
      'If aLvl >= 0.5 And gilvl < 0.5 Then
        'Sound_GI_Relay 1, Bumper1 'Note: Bumper1 is just used for sound positioning. Can be anywhere that makes sense.
      'ElseIf aLvl <= 0.4 And gilvl > 0.4 Then
        'Sound_GI_Relay 0, Bumper1
      'End If

      'You may add any other GI related effects here. Like if you want to make some toy to appear more bright, set it like this:
      'Primitive001.blenddisablelighting = 1.2 * aLvl + 0.2 'This will result to DL brightness between 0.2 - 1.4 for ON/OFF states

      'Pop Bumper Lights (PWM Following)
      'FlFadeBumper 1,aLvl
      'FlFadeBumper 2,aLvl
      'FlFadeBumper 3,aLvl

      'gilvl = aLvl   'Storing the latest GI fading state into global variable, so one can use it elsewhere too.

    'Case 1:  'GI String 1

    'Case 2:  'GI String 2

  'End Select

'End Sub

'******************************************************
'****  END GI Control
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
Dim objBallShadow(5)

'Initialization
BSInit

Sub BSInit()
  Dim iii
  'Prepare the shadow objects before play begins
  For iii = 0 To tnob+lob- 1
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
'  ZVRR: VR Room / VR Cabinet
'******************************************************

dim DisplayColor, DisplayColorG
DisplayColor =  RGB(0,37,251)

Sub VRDisplayTimer
  Dim ii, jj, obj, b, x
  Dim ChgLED,num, chg, stat
  ChgLED=Controller.ChangedLEDs(&Hffffffff, &Hffffffff)
    If Not IsEmpty(ChgLED) Then
      For ii=0 To UBound(chgLED)
        num=chgLED(ii, 0) : chg=chgLED(ii, 1) : stat=chgLED(ii, 2)
        For Each obj In VRDigits(num)
          If chg And 1 Then FadeDisplay obj, stat And 1
          chg=chg\2 : stat=stat\2
        Next
      Next
    End If
End Sub

Sub FadeDisplay(object, onoffstat)
  If OnOffstat = 1 Then
    object.color = DisplayColor
  Else
    Object.Color = RGB(1,1,1)
  End If
End Sub

Dim VRDigits(40)

VRDigits(0)=Array(D1,D2,D3,D4,D5,D6,D7,D8,D9,D10,D11,D12,D13,D14,D15)
VRDigits(1)=Array(D16,D17,D18,D19,D20,D21,D22,D23,D24,D25,D26,D27,D28,D29,D30)
VRDigits(2)=Array(D31,D32,D33,D34,D35,D36,D37,D38,D39,D40,D41,D42,D43,D44,D45)
VRDigits(3)=Array(D46,D47,D48,D49,D50,D51,D52,D53,D54,D55,D56,D57,D58,D59,D60)
VRDigits(4)=Array(D61,D62,D63,D64,D65,D66,D67,D68,D69,D70,D71,D72,D73,D74,D75)
VRDigits(5)=Array(D76,D77,D78,D79,D80,D81,D82,D83,D84,D85,D86,D87,D88,D89,D90)
VRDigits(6)=Array(D91,D92,D93,D94,D95,D96,D97,D98,D99,D100,D101,D102,D103,D104,D105)
VRDigits(7)=Array(D106,D107,D108,D109,D110,D111,D112,D113,D114,D115,D116,D117,D118,D119,D120)
VRDigits(8)=Array(D121,D122,D123,D124,D125,D126,D127,D128,D129,D130,D131,D132,D133,D134,D135)
VRDigits(9)=Array(D136,D137,D138,D139,D140,D141,D142,D143,D144,D145,D146,D147,D148,D149,D150)
VRDigits(10)=Array(D151,D152,D153,D154,D155,D156,D157,D158,D159,D160,D161,D162,D163,D164,D165)
VRDigits(11)=Array(D166,D167,D168,D169,D170,D171,D172,D173,D174,D175,D176,D177,D178,D179,D180)
VRDigits(12)=Array(D181,D182,D183,D184,D185,D186,D187,D188,D189,D190,D191,D192,D193,D194,D195)
VRDigits(13)=Array(D196,D197,D198,D199,D200,D201,D202,D203,D204,D205,D206,D207,D208,D209,D210)
VRDigits(14)=Array(D211,D212,D213,D214,D215,D216,D217,D218,D219,D220,D221,D222,D223,D224,D225)
VRDigits(15)=Array(D226,D227,D228,D229,D230,D231,D232,D233,D234,D235,D236,D237,D238,D239,D240)

VRDigits(16)=Array(D241,D242,D243,D244,D245,D246,D247,D248,D249,D250,D251,D252,D253,D254,D255)
VRDigits(17)=Array(D256,D257,D258,D259,D260,D261,D262,D263,D264,D265,D266,D267,D268,D269,D270)
VRDigits(18)=Array(D271,D272,D273,D274,D275,D276,D277,D278,D279,D280,D281,D282,D283,D284,D285)
VRDigits(19)=Array(D286,D287,D288,D289,D290,D291,D292,D293,D294,D295,D296,D297,D298,D299,D300)
VRDigits(20)=Array(D301,D302,D303,D304,D305,D306,D307,D308,D309,D310,D311,D312,D313,D314,D315)
VRDigits(21)=Array(D316,D317,D318,D319,D320,D321,D322,D323,D324,D325,D326,D327,D328,D329,D330)
VRDigits(22)=Array(D331,D332,D333,D334,D335,D336,D337,D338,D339,D340,D341,D342,D343,D344,D345)
VRDigits(23)=Array(D346,D347,D348,D349,D350,D351,D352,D353,D354,D355,D356,D357,D358,D359,D360)
VRDigits(24)=Array(D361,D362,D363,D364,D365,D366,D367,D368,D369,D370,D371,D372,D373,D374,D375)
VRDigits(25)=Array(D376,D377,D378,D379,D380,D381,D382,D383,D384,D385,D386,D387,D388,D389,D390)
VRDigits(26)=Array(D391,D392,D393,D394,D395,D396,D397,D398,D399,D400,D401,D402,D403,D404,D405)
VRDigits(27)=Array(D406,D407,D408,D409,D410,D411,D412,D413,D414,D415,D416,D417,D418,D419,D420)
VRDigits(28)=Array(D421,D422,D423,D424,D425,D426,D427,D428,D429,D430,D431,D432,D433,D434,D435)
VRDigits(29)=Array(D436,D437,D438,D439,D440,D441,D442,D443,D444,D445,D446,D447,D448,D449,D450)
VRDigits(30)=Array(D451,D452,D453,D454,D455,D456,D457,D458,D459,D460,D461,D462,D463,D464,D465)
VRDigits(31)=Array(D466,D467,D468,D469,D470,D471,D472,D473,D474,D475,D476,D477,D478,D479,D480)

VRDigits(32)=Array(D481,D482,D483,D484,D485,D486,D487,D488,D489,D490,D491,D492,D493,D494,D495)
VRDigits(33)=Array(D496,D497,D498,D499,D500,D501,D502,D503,D504,D505,D506,D507,D508,D509,D510)
VRDigits(34)=Array(D511,D512,D513,D514,D515,D516,D517,D518,D519,D520,D521,D522,D523,D524,D525)
VRDigits(35)=Array(D526,D527,D528,D529,D530,D531,D532,D533,D534,D535,D536,D537,D538,D539,D540)
VRDigits(36)=Array(D541,D542,D543,D544,D545,D546,D547,D548,D549,D550,D551,D552,D553,D554,D555)
VRDigits(37)=Array(D556,D557,D558,D559,D560,D561,D562,D563,D564,D565,D566,D567,D568,D569,D570)
VRDigits(38)=Array(D571,D572,D573,D574,D575,D576,D577,D578,D579,D580,D581,D582,D583,D584,D585)
VRDigits(39)=Array(D586,D587,D588,D589,D590,D591,D592,D593,D594,D595,D596,D597,D598,D599,D600)

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
End Sub

If VRRoom = 1 Then
  InitDigits
End If

Dim VRThings

' Desktop Mode
if VRRoom = 0 and cab_mode = 0 Then
  for each VRThings in VRStuff:VRThings.visible = 0:Next
  for each VRThings in VRBackglass:VRThings.visible = 0:Next
  For each VRThings in DT_BG_Lights: VRThings.Visible = 1: Next
  VR_SideRailL.visible = 1
  VR_SideRailR.visible = 1

' Cabinet Mode
Elseif VRRoom = 0 and cab_mode = 1 Then
  for each VRThings in VRStuff:VRThings.visible = 0:Next
  for each VRThings in VRBackglass:VRThings.visible = 0:Next
  For each VRThings in DT_BG_Lights: VRThings.Visible = 0: Next
  VR_SideRailL.visible = 0
  VR_SideRailR.visible = 0

' VR Mode
Else
  for each VRThings in VRStuff:VRThings.visible = 1:Next
  for each VRThings in VRBackglass:VRThings.visible = 1:Next
  For each VRThings in DT_BG_Lights:VRThings.Visible = 0:Next
  VR_SideRailL.visible = 1
  VR_SideRailR.visible = 1

End If

'**********************************************************************************************************
'Digital Display
'**********************************************************************************************************
Dim Digits(40)
Digits(0) = Array(a00, a05, a0c, a0d, a08, a01, a06, a0f, a02, a03, a04, a07, a0b, a0a, a09, a0e)
Digits(1) = Array(a10, a15, a1c, a1d, a18, a11, a16, a1f, a12, a13, a14, a17, a1b, a1a, a19, a1e)
Digits(2) = Array(a20, a25, a2c, a2d, a28, a21, a26, a2f, a22, a23, a24, a27, a2b, a2a, a29, a2e)
Digits(3) = Array(a30, a35, a3c, a3d, a38, a31, a36, a3f, a32, a33, a34, a37, a3b, a3a, a39, a3e)
Digits(4) = Array(a40, a45, a4c, a4d, a48, a41, a46, a4f, a42, a43, a44, a47, a4b, a4a, a49, a4e)
Digits(5) = Array(a50, a55, a5c, a5d, a58, a51, a56, a5f, a52, a53, a54, a57, a5b, a5a, a59, a5e)
Digits(6) = Array(a60, a65, a6c, a6d, a68, a61, a66, a6f, a62, a63, a64, a67, a6b, a6a, a69, a6e)
Digits(7) = Array(a70, a75, a7c, a7d, a78, a71, a76, a7f, a72, a73, a74, a77, a7b, a7a, a79, a7e)
Digits(8) = Array(a80, a85, a8c, a8d, a88, a81, a86, a8f, a82, a83, a84, a87, a8b, a8a, a89, a8e)
Digits(9) = Array(a90, a95, a9c, a9d, a98, a91, a96, a9f, a92, a93, a94, a97, a9b, a9a, a99, a9e)
Digits(10) = Array(aa0, aa5, aac, aad, aa8, aa1, aa6, aaf, aa2, aa3, aa4, aa7, aab, aaa, aa9, aae)
Digits(11) = Array(ab0, ab5, abc, abd, ab8, ab1, ab6, abf, ab2, ab3, ab4, ab7, abb, aba, ab9, abe)
Digits(12) = Array(ac0, ac5, acc, acd, ac8, ac1, ac6, acf, ac2, ac3, ac4, ac7, acb, aca, ac9, ace)
Digits(13) = Array(ad0, ad5, adc, add, ad8, ad1, ad6, adf, ad2, ad3, ad4, ad7, adb, ada, ad9, ade)
Digits(14) = Array(ae0, ae5, aec, aed, ae8, ae1, ae6, aef, ae2, ae3, ae4, ae7, aeb, aea, ae9, aee)
Digits(15) = Array(af0, af5, afc, afd, af8, af1, af6, aff, af2, af3, af4, af7, afb, afa, af9, afe)

Digits(16) = Array(b00, b05, b0c, b0d, b08, b01, b06, b0f, b02, b03, b04, b07, b0b, b0a, b09, b0e)
Digits(17) = Array(b10, b15, b1c, b1d, b18, b11, b16, b1f, b12, b13, b14, b17, b1b, b1a, b19, b1e)
Digits(18) = Array(b20, b25, b2c, b2d, b28, b21, b26, b2f, b22, b23, b24, b27, b2b, b2a, b29, b2e)
Digits(19) = Array(b30, b35, b3c, b3d, b38, b31, b36, b3f, b32, b33, b34, b37, b3b, b3a, b39, b3e)
Digits(20) = Array(b40, b45, b4c, b4d, b48, b41, b46, b4f, b42, b43, b44, b47, b4b, b4a, b49, b4e)
Digits(21) = Array(b50, b55, b5c, b5d, b58, b51, b56, b5f, b52, b53, b54, b57, b5b, b5a, b59, b5e)
Digits(22) = Array(b60, b65, b6c, b6d, b68, b61, b66, b6f, b62, b63, b64, b67, b6b, b6a, b69, b6e)
Digits(23) = Array(b70, b75, b7c, b7d, b78, b71, b76, b7f, b72, b73, b74, b77, b7b, b7a, b79, b7e)
Digits(24) = Array(b80, b85, b8c, b8d, b88, b81, b86, b8f, b82, b83, b84, b87, b8b, b8a, b89, b8e)
Digits(25) = Array(b90, b95, b9c, b9d, b98, b91, b96, b9f, b92, b93, b94, b97, b9b, b9a, b99, b9e)
Digits(26) = Array(ba0, ba5, bac, bad, ba8, ba1, ba6, baf, ba2, ba3, ba4, ba7, bab, baa, ba9, bae)
Digits(27) = Array(bb0, bb5, bbc, bbd, bb8, bb1, bb6, bbf, bb2, bb3, bb4, bb7, bbb, bba, bb9, bbe)
Digits(28) = Array(bc0, bc5, bcc, bcd, bc8, bc1, bc6, bcf, bc2, bc3, bc4, bc7, bcb, bca, bc9, bce)
Digits(29) = Array(bd0, bd5, bdc, bdd, bd8, bd1, bd6, bdf, bd2, bd3, bd4, bd7, bdb, bda, bd9, bde)
Digits(30) = Array(be0, be5, bec, bed, be8, be1, be6, bef, be2, be3, be4, be7, beb, bea, be9, bee)
Digits(31) = Array(bf0, bf5, bfc, bfd, bf8, bf1, bf6, bff, bf2, bf3, bf4, bf7, bfb, bfa, bf9, bfe)

Digits(32) = Array(c00, c05, c0c, c0d, c08, c01, c06, c0f, c02, c03, c04, c07, c0b, c0a, c09, c0e)
Digits(33) = Array(c10, c15, c1c, c1d, c18, c11, c16, c1f, c12, c13, c14, c17, c1b, c1a, c19, c1e)
Digits(34) = Array(c20, c25, c2c, c2d, c28, c21, c26, c2f, c22, c23, c24, c27, c2b, c2a, c29, c2e)
Digits(35) = Array(c30, c35, c3c, c3d, c38, c31, c36, c3f, c32, c33, c34, c37, c3b, c3a, c39, c3e)
Digits(36) = Array(c40, c45, c4c, c4d, c48, c41, c46, c4f, c42, c43, c44, c47, c4b, c4a, c49, c4e)
Digits(37) = Array(c50, c55, c5c, c5d, c58, c51, c56, c5f, c52, c53, c54, c57, c5b, c5a, c59, c5e)
Digits(38) = Array(c60, c65, c6c, c6d, c68, c61, c66, c6f, c62, c63, c64, c67, c6b, c6a, c69, c6e)
Digits(39) = Array(c70, c75, c7c, c7d, c78, c71, c76, c7f, c72, c73, c74, c77, c7b, c7a, c79, c7e)

 Sub DisplayTimer
    Dim ChgLED, ii, num, chg, stat, obj
    ChgLED=Controller.ChangedLEDs(&Hffffffff, &Hffffffff)
    If Not IsEmpty(ChgLED)Then
      For ii=0 To UBound(chgLED)
        num=chgLED(ii, 0) : chg=chgLED(ii, 1) : stat=chgLED(ii, 2)
        For Each obj In Digits(num)
          If chg And 1 Then obj.State=stat And 1
          chg=chg\2 : stat=stat\2
        Next
      Next
    End If
 End Sub


'******************************************************
'  ZVLM: VLM Arrays
'******************************************************



' VLM  Arrays - Start
' Arrays per baked part
Dim BP_Bumper1_Ring: BP_Bumper1_Ring=Array(BM_Bumper1_Ring, LM_L_BumperLight1_Bumper1_Ring, LM_L_BumperLightA1_Bumper1_Ring, LM_F_Bumper1_Ring, LM_GI_GI4_Bumper1_Ring, LM_GI_Llsling1_Bumper1_Ring, LM_L_al10a_Bumper1_Ring, LM_L_al1a_Bumper1_Ring, LM_L_al2a_Bumper1_Ring, LM_L_al3a_Bumper1_Ring, LM_L_al4a_Bumper1_Ring, LM_L_al5a_Bumper1_Ring, LM_GI_al6a_Bumper1_Ring, LM_L_al7a_Bumper1_Ring, LM_L_al8a_Bumper1_Ring, LM_L_al9a_Bumper1_Ring, LM_L_l23a_Bumper1_Ring)
Dim BP_Gate1_Wire: BP_Gate1_Wire=Array(BM_Gate1_Wire)
Dim BP_Gate3_Wire: BP_Gate3_Wire=Array(BM_Gate3_Wire)
Dim BP_Gate4_Wire: BP_Gate4_Wire=Array(BM_Gate4_Wire)
Dim BP_Gate5_Wire: BP_Gate5_Wire=Array(BM_Gate5_Wire, LM_GI_Llsling2_Gate5_Wire)
Dim BP_Gate7_Wire: BP_Gate7_Wire=Array(BM_Gate7_Wire, LM_L_l12_Gate7_Wire, LM_L_l48_Gate7_Wire, LM_L_l49_Gate7_Wire)
Dim BP_Gate8_Wire: BP_Gate8_Wire=Array(BM_Gate8_Wire, LM_L_BumperLightA1_Gate8_Wire, LM_F_Gate8_Wire, LM_GI_GI2_Gate8_Wire)
Dim BP_LFlip1U: BP_LFlip1U=Array(BM_LFlip1U, LM_GI_GI1_LFlip1U, LM_GI_Llsling1_LFlip1U, LM_L_l27_LFlip1U)
Dim BP_LFlip1_001: BP_LFlip1_001=Array(BM_LFlip1_001, LM_GI_GI1_LFlip1_001, LM_GI_Llsling1_LFlip1_001, LM_L_l27_LFlip1_001)
Dim BP_Layer_1: BP_Layer_1=Array(BM_Layer_1, LM_L_BumperLight1_Layer_1, LM_L_BumperLightA1_Layer_1, LM_F_Layer_1, LM_GI_GI1_Layer_1, LM_GI_GI4_Layer_1, LM_GI_GI5_Layer_1, LM_GI_GI6_Layer_1, LM_GI_GI7_Layer_1, LM_L_La6_Layer_1, LM_L_La7_Layer_1, LM_L_La8_Layer_1, LM_GI_Llsling1_Layer_1, LM_GI_Llsling2_Layer_1, LM_L_al2a_Layer_1, LM_L_al3a_Layer_1, LM_L_al4a_Layer_1, LM_L_al5_Layer_1, LM_L_al5a_Layer_1, LM_GI_al6a_Layer_1, LM_L_al7a_Layer_1, LM_L_al8_Layer_1, LM_L_al8a_Layer_1, LM_L_al9_Layer_1, LM_L_l16_Layer_1, LM_L_l18_Layer_1, LM_L_l19_Layer_1, LM_L_l19a_Layer_1, LM_L_l20_Layer_1, LM_L_l21_Layer_1, LM_L_l22_Layer_1, LM_L_l23a_Layer_1, LM_L_l24_Layer_1, LM_L_l27_Layer_1, LM_L_l36_Layer_1, LM_L_l37_Layer_1, LM_L_l38_Layer_1, LM_L_l39_Layer_1, LM_L_l40_Layer_1, LM_L_l49_Layer_1, LM_L_l50_Layer_1, LM_L_l9_Layer_1)
Dim BP_Layer_2: BP_Layer_2=Array(BM_Layer_2, LM_L_BumperLightA1_Layer_2, LM_F_Layer_2, LM_GI_GI4_Layer_2, LM_GI_GI5_Layer_2, LM_GI_GI6_Layer_2, LM_GI_al6a_Layer_2, LM_L_l19a_Layer_2, LM_L_l46_Layer_2)
Dim BP_LflipU: BP_LflipU=Array(BM_LflipU, LM_GI_GI7_LflipU, LM_L_l41_LflipU, LM_L_l42_LflipU, LM_L_l43_LflipU, LM_L_l44_LflipU, LM_L_l45_LflipU, LM_L_l46_LflipU)
Dim BP_Lflip_001: BP_Lflip_001=Array(BM_Lflip_001, LM_GI_GI7_Lflip_001, LM_L_l37_Lflip_001, LM_L_l38_Lflip_001, LM_L_l39_Lflip_001, LM_L_l40_Lflip_001, LM_L_l41_Lflip_001, LM_L_l42_Lflip_001, LM_L_l43_Lflip_001, LM_L_l44_Lflip_001, LM_L_l45_Lflip_001, LM_L_l46_Lflip_001)
Dim BP_Parts: BP_Parts=Array(BM_Parts, LM_L_BumperLight1_Parts, LM_L_BumperLightA1_Parts, LM_F_Parts, LM_GI_GI1_Parts, LM_GI_GI2_Parts, LM_GI_GI3_Parts, LM_GI_GI4_Parts, LM_GI_GI5_Parts, LM_GI_GI6_Parts, LM_GI_GI7_Parts, LM_L_La32_Parts, LM_L_La33_Parts, LM_L_La34_Parts, LM_L_La35_Parts, LM_L_La6_Parts, LM_L_La7_Parts, LM_L_La8_Parts, LM_L_LkickL_Parts, LM_L_LkickR_Parts, LM_GI_Llane1_Parts, LM_GI_Llane2_Parts, LM_GI_Llsling1_Parts, LM_GI_Llsling2_Parts, LM_L_al10_Parts, LM_L_al10a_Parts, LM_L_al1_Parts, LM_L_al1a_Parts, LM_L_al2a_Parts, LM_L_al3a_Parts, LM_L_al4a_Parts, LM_L_al5_Parts, LM_L_al5a_Parts, LM_L_al6_Parts, LM_GI_al6a_Parts, LM_L_al7_Parts, LM_L_al7a_Parts, LM_L_al8_Parts, LM_L_al8a_Parts, LM_L_al9_Parts, LM_L_al9a_Parts, LM_L_l10_Parts, LM_L_l11_Parts, LM_L_l12_Parts, LM_L_l14_Parts, LM_L_l14a_Parts, LM_L_l15_Parts, LM_L_l15a_Parts, LM_L_l16_Parts, LM_L_l16a_Parts, LM_L_l17_Parts, LM_L_l18_Parts, LM_L_l19_Parts, LM_L_l19a_Parts, LM_L_L2_Parts, LM_L_l20_Parts, LM_L_l21_Parts, LM_L_l22_Parts, _
  LM_L_l23_Parts, LM_L_l23a_Parts, LM_L_l24_Parts, LM_L_l25_Parts, LM_L_l26_Parts, LM_L_l27_Parts, LM_L_l29_Parts, LM_L_l30_Parts, LM_L_l36_Parts, LM_L_l37_Parts, LM_L_l38_Parts, LM_L_l39_Parts, LM_L_l3_Parts, LM_L_l40_Parts, LM_L_l41_Parts, LM_L_l42_Parts, LM_L_l43_Parts, LM_L_l44_Parts, LM_L_l45_Parts, LM_L_l46_Parts, LM_L_l47_Parts, LM_L_l47a_Parts, LM_L_l48_Parts, LM_L_l49_Parts, LM_L_l50_Parts, LM_L_l9_Parts)
Dim BP_Pgate: BP_Pgate=Array(BM_Pgate)
Dim BP_Playfield: BP_Playfield=Array(BM_Playfield, LM_L_BumperLight1_Playfield, LM_L_BumperLightA1_Playfield, LM_F_Playfield, LM_GI_GI1_Playfield, LM_GI_GI2_Playfield, LM_GI_GI3_Playfield, LM_GI_GI4_Playfield, LM_GI_GI5_Playfield, LM_GI_GI6_Playfield, LM_GI_GI7_Playfield, LM_L_La32_Playfield, LM_L_La33_Playfield, LM_L_La34_Playfield, LM_L_La35_Playfield, LM_L_La5_Playfield, LM_L_La6_Playfield, LM_L_La7_Playfield, LM_L_La8_Playfield, LM_L_LkickR_Playfield, LM_GI_Llane1_Playfield, LM_GI_Llane2_Playfield, LM_GI_Llsling1_Playfield, LM_GI_Llsling2_Playfield, LM_L_al10_Playfield, LM_L_al10a_Playfield, LM_L_al1_Playfield, LM_L_al1a_Playfield, LM_L_al2_Playfield, LM_L_al2a_Playfield, LM_L_al3_Playfield, LM_L_al3a_Playfield, LM_L_al4_Playfield, LM_L_al4a_Playfield, LM_L_al5_Playfield, LM_L_al5a_Playfield, LM_L_al6_Playfield, LM_GI_al6a_Playfield, LM_L_al7_Playfield, LM_L_al7a_Playfield, LM_L_al8_Playfield, LM_L_al8a_Playfield, LM_L_al9a_Playfield, LM_L_l10_Playfield, LM_L_l11_Playfield, LM_L_l12_Playfield, _
  LM_L_l14_Playfield, LM_L_l14a_Playfield, LM_L_l15_Playfield, LM_L_l15a_Playfield, LM_L_l16_Playfield, LM_L_l16a_Playfield, LM_L_l17_Playfield, LM_L_l18_Playfield, LM_L_l19_Playfield, LM_L_l19a_Playfield, LM_L_l20_Playfield, LM_L_l21_Playfield, LM_L_l22_Playfield, LM_L_l23_Playfield, LM_L_l23a_Playfield, LM_L_l24_Playfield, LM_L_l25_Playfield, LM_L_l26_Playfield, LM_L_l27_Playfield, LM_L_l28_Playfield, LM_L_l29_Playfield, LM_L_l30_Playfield, LM_L_l31_Playfield, LM_L_l36_Playfield, LM_L_l37_Playfield, LM_L_l38_Playfield, LM_L_l39_Playfield, LM_L_l3_Playfield, LM_L_l40_Playfield, LM_L_l41_Playfield, LM_L_l42_Playfield, LM_L_l43_Playfield, LM_L_l44_Playfield, LM_L_l45_Playfield, LM_L_l46_Playfield, LM_L_l47_Playfield, LM_L_l47a_Playfield, LM_L_l48_Playfield, LM_L_l49_Playfield, LM_L_l50_Playfield, LM_L_l9_Playfield)
Dim BP_RFlipU: BP_RFlipU=Array(BM_RFlipU, LM_GI_GI6_RFlipU, LM_L_La32_RFlipU, LM_L_l41_RFlipU, LM_L_l42_RFlipU, LM_L_l43_RFlipU, LM_L_l44_RFlipU, LM_L_l45_RFlipU, LM_L_l46_RFlipU)
Dim BP_RFlip_001: BP_RFlip_001=Array(BM_RFlip_001, LM_L_La32_RFlip_001, LM_L_La6_RFlip_001, LM_L_l37_RFlip_001, LM_L_l40_RFlip_001, LM_L_l41_RFlip_001, LM_L_l42_RFlip_001, LM_L_l43_RFlip_001, LM_L_l44_RFlip_001, LM_L_l45_RFlip_001, LM_L_l46_RFlip_001)
Dim BP_Rflip1U: BP_Rflip1U=Array(BM_Rflip1U, LM_GI_GI1_Rflip1U, LM_GI_GI4_Rflip1U, LM_GI_Llsling1_Rflip1U, LM_L_l27_Rflip1U)
Dim BP_Rflip1_001: BP_Rflip1_001=Array(BM_Rflip1_001, LM_GI_GI1_Rflip1_001, LM_GI_GI4_Rflip1_001, LM_GI_Llsling1_Rflip1_001, LM_L_l27_Rflip1_001)
Dim BP_URsling1: BP_URsling1=Array(BM_URsling1, LM_L_BumperLightA1_URsling1, LM_F_URsling1, LM_GI_GI4_URsling1, LM_GI_GI5_URsling1, LM_L_l12_URsling1, LM_L_l47_URsling1)
Dim BP_URsling2: BP_URsling2=Array(BM_URsling2, LM_L_BumperLightA1_URsling2, LM_F_URsling2, LM_GI_GI4_URsling2, LM_GI_GI5_URsling2, LM_L_l12_URsling2, LM_L_l47_URsling2)
Dim BP_Under_PF: BP_Under_PF=Array(BM_Under_PF, LM_L_BumperLightA1_Under_PF, LM_F_Under_PF, LM_GI_GI1_Under_PF, LM_GI_GI3_Under_PF, LM_GI_GI7_Under_PF, LM_L_La32_Under_PF, LM_L_La33_Under_PF, LM_L_La34_Under_PF, LM_L_La35_Under_PF, LM_L_La5_Under_PF, LM_L_La6_Under_PF, LM_L_La7_Under_PF, LM_L_La8_Under_PF, LM_L_LkickL_Under_PF, LM_L_LkickR_Under_PF, LM_GI_Llsling1_Under_PF, LM_GI_Llsling2_Under_PF, LM_L_al10_Under_PF, LM_L_al10a_Under_PF, LM_L_al1_Under_PF, LM_L_al1a_Under_PF, LM_L_al2_Under_PF, LM_L_al2a_Under_PF, LM_L_al3_Under_PF, LM_L_al3a_Under_PF, LM_L_al4_Under_PF, LM_L_al4a_Under_PF, LM_L_al5_Under_PF, LM_L_al5a_Under_PF, LM_L_al6_Under_PF, LM_GI_al6a_Under_PF, LM_L_al7_Under_PF, LM_L_al7a_Under_PF, LM_L_al8_Under_PF, LM_L_al8a_Under_PF, LM_L_al9_Under_PF, LM_L_al9a_Under_PF, LM_L_l10_Under_PF, LM_L_l11_Under_PF, LM_L_l12_Under_PF, LM_L_l14_Under_PF, LM_L_l14a_Under_PF, LM_L_l15_Under_PF, LM_L_l15a_Under_PF, LM_L_l16_Under_PF, LM_L_l16a_Under_PF, LM_L_l17_Under_PF, LM_L_l18_Under_PF, _
  LM_L_l19_Under_PF, LM_L_l19a_Under_PF, LM_L_l20_Under_PF, LM_L_l21_Under_PF, LM_L_l22_Under_PF, LM_L_l23_Under_PF, LM_L_l23a_Under_PF, LM_L_l24_Under_PF, LM_L_l25_Under_PF, LM_L_l26_Under_PF, LM_L_l27_Under_PF, LM_L_l28_Under_PF, LM_L_l29_Under_PF, LM_L_l30_Under_PF, LM_L_l31_Under_PF, LM_L_l36_Under_PF, LM_L_l37_Under_PF, LM_L_l38_Under_PF, LM_L_l39_Under_PF, LM_L_l3_Under_PF, LM_L_l40_Under_PF, LM_L_l41_Under_PF, LM_L_l42_Under_PF, LM_L_l43_Under_PF, LM_L_l44_Under_PF, LM_L_l45_Under_PF, LM_L_l46_Under_PF, LM_L_l47_Under_PF, LM_L_l47a_Under_PF, LM_L_l48_Under_PF, LM_L_l49_Under_PF, LM_L_l50_Under_PF, LM_L_l9_Under_PF)
Dim BP_lsling1: BP_lsling1=Array(BM_lsling1, LM_GI_GI6_lsling1, LM_GI_GI7_lsling1, LM_L_La6_lsling1, LM_L_La8_lsling1, LM_L_l12_lsling1, LM_L_l37_lsling1)
Dim BP_lsling2: BP_lsling2=Array(BM_lsling2, LM_GI_GI6_lsling2, LM_GI_GI7_lsling2, LM_L_La6_lsling2, LM_L_l12_lsling2, LM_L_l37_lsling2)
Dim BP_rsling1: BP_rsling1=Array(BM_rsling1, LM_F_rsling1, LM_GI_GI6_rsling1, LM_L_La35_rsling1, LM_L_l12_rsling1, LM_L_l36_rsling1, LM_L_l39_rsling1, LM_L_l47_rsling1)
Dim BP_rsling2: BP_rsling2=Array(BM_rsling2, LM_GI_GI6_rsling2, LM_L_l12_rsling2, LM_L_l36_rsling2, LM_L_l39_rsling2)
Dim BP_sw31: BP_sw31=Array(BM_sw31)
Dim BP_sw40: BP_sw40=Array(BM_sw40, LM_GI_GI3_sw40, LM_GI_Llsling1_sw40, LM_L_l17_sw40, LM_L_l23_sw40, LM_L_l31_sw40, LM_L_l47a_sw40)
Dim BP_sw41: BP_sw41=Array(BM_sw41, LM_GI_GI4_sw41, LM_GI_GI5_sw41, LM_GI_GI6_sw41, LM_GI_Llsling2_sw41, LM_L_l47_sw41)
Dim BP_sw42: BP_sw42=Array(BM_sw42, LM_GI_GI3_sw42, LM_GI_Llane1_sw42, LM_L_l48_sw42)
Dim BP_sw43_001: BP_sw43_001=Array(BM_sw43_001, LM_GI_GI5_sw43_001, LM_GI_GI7_sw43_001, LM_GI_Llsling2_sw43_001, LM_L_l11_sw43_001, LM_L_l20_sw43_001, LM_L_l21_sw43_001, LM_L_l22_sw43_001, LM_L_l37_sw43_001, LM_L_l38_sw43_001, LM_L_l39_sw43_001, LM_L_l40_sw43_001, LM_L_l9_sw43_001)
Dim BP_sw44_001: BP_sw44_001=Array(BM_sw44_001, LM_F_sw44_001, LM_GI_GI4_sw44_001, LM_GI_GI5_sw44_001, LM_GI_Llsling2_sw44_001, LM_L_l24_sw44_001, LM_L_l25_sw44_001, LM_L_l26_sw44_001)
Dim BP_sw45: BP_sw45=Array(BM_sw45, LM_GI_GI1_sw45, LM_GI_GI7_sw45)
Dim BP_sw50: BP_sw50=Array(BM_sw50, LM_GI_GI3_sw50, LM_GI_Llsling1_sw50, LM_L_l17_sw50, LM_L_l18_sw50, LM_L_l23_sw50, LM_L_l31_sw50, LM_L_l47a_sw50)
Dim BP_sw51: BP_sw51=Array(BM_sw51, LM_GI_GI5_sw51, LM_GI_GI6_sw51, LM_GI_Llsling2_sw51, LM_L_l47_sw51)
Dim BP_sw52: BP_sw52=Array(BM_sw52, LM_GI_GI3_sw52, LM_GI_Llane1_sw52, LM_GI_Llane2_sw52, LM_L_l49_sw52)
Dim BP_sw53_001: BP_sw53_001=Array(BM_sw53_001, LM_GI_GI5_sw53_001, LM_GI_GI7_sw53_001, LM_GI_Llsling2_sw53_001, LM_L_l20_sw53_001, LM_L_l21_sw53_001, LM_L_l22_sw53_001, LM_L_l37_sw53_001, LM_L_l38_sw53_001, LM_L_l39_sw53_001, LM_L_l9_sw53_001)
Dim BP_sw54_001: BP_sw54_001=Array(BM_sw54_001, LM_F_sw54_001, LM_GI_GI4_sw54_001, LM_GI_GI5_sw54_001, LM_GI_Llsling2_sw54_001, LM_L_al4_sw54_001, LM_L_l19a_sw54_001, LM_L_l24_sw54_001, LM_L_l25_sw54_001, LM_L_l26_sw54_001)
Dim BP_sw55: BP_sw55=Array(BM_sw55, LM_GI_GI7_sw55)
Dim BP_sw60: BP_sw60=Array(BM_sw60, LM_GI_GI3_sw60, LM_GI_Llsling1_sw60, LM_L_l17_sw60, LM_L_l18_sw60, LM_L_l31_sw60, LM_L_l47a_sw60)
Dim BP_sw61: BP_sw61=Array(BM_sw61, LM_GI_GI5_sw61, LM_GI_GI6_sw61, LM_GI_Llsling2_sw61, LM_L_l36_sw61, LM_L_l39_sw61)
Dim BP_sw62: BP_sw62=Array(BM_sw62, LM_L_BumperLightA1_sw62, LM_GI_GI3_sw62, LM_GI_Llane2_sw62, LM_L_l50_sw62)
Dim BP_sw63_001: BP_sw63_001=Array(BM_sw63_001, LM_GI_GI4_sw63_001, LM_GI_GI5_sw63_001, LM_GI_Llsling2_sw63_001, LM_L_l20_sw63_001, LM_L_l21_sw63_001, LM_L_l22_sw63_001, LM_L_l24_sw63_001, LM_L_l25_sw63_001, LM_L_l26_sw63_001, LM_L_l38_sw63_001, LM_L_l39_sw63_001, LM_L_l9_sw63_001)
Dim BP_sw64_001: BP_sw64_001=Array(BM_sw64_001, LM_F_sw64_001, LM_GI_GI4_sw64_001, LM_GI_GI5_sw64_001, LM_GI_Llsling2_sw64_001, LM_L_al4_sw64_001, LM_L_al5_sw64_001, LM_L_al6_sw64_001, LM_GI_al6a_sw64_001, LM_L_l19_sw64_001, LM_L_l19a_sw64_001, LM_L_l24_sw64_001, LM_L_l25_sw64_001, LM_L_l26_sw64_001)
Dim BP_sw65: BP_sw65=Array(BM_sw65, LM_GI_GI7_sw65)
Dim BP_sw73: BP_sw73=Array(BM_sw73, LM_L_l23_sw73)
Dim BP_sw73a_001: BP_sw73a_001=Array(BM_sw73a_001, LM_L_BumperLightA1_sw73a_001, LM_F_sw73a_001, LM_L_al9a_sw73a_001, LM_L_l23a_sw73a_001)
Dim BP_sw75: BP_sw75=Array(BM_sw75, LM_GI_GI6_sw75)
' Arrays per lighting scenario
Dim BL_F: BL_F=Array(LM_F_Bumper1_Ring, LM_F_Gate8_Wire, LM_F_Layer_1, LM_F_Layer_2, LM_F_Parts, LM_F_Playfield, LM_F_URsling1, LM_F_URsling2, LM_F_Under_PF, LM_F_rsling1, LM_F_sw44_001, LM_F_sw54_001, LM_F_sw64_001, LM_F_sw73a_001)
Dim BL_GI_GI1: BL_GI_GI1=Array(LM_GI_GI1_LFlip1_001, LM_GI_GI1_LFlip1U, LM_GI_GI1_Layer_1, LM_GI_GI1_Parts, LM_GI_GI1_Playfield, LM_GI_GI1_Rflip1_001, LM_GI_GI1_Rflip1U, LM_GI_GI1_Under_PF, LM_GI_GI1_sw45)
Dim BL_GI_GI2: BL_GI_GI2=Array(LM_GI_GI2_Gate8_Wire, LM_GI_GI2_Parts, LM_GI_GI2_Playfield)
Dim BL_GI_GI3: BL_GI_GI3=Array(LM_GI_GI3_Parts, LM_GI_GI3_Playfield, LM_GI_GI3_Under_PF, LM_GI_GI3_sw40, LM_GI_GI3_sw42, LM_GI_GI3_sw50, LM_GI_GI3_sw52, LM_GI_GI3_sw60, LM_GI_GI3_sw62)
Dim BL_GI_GI4: BL_GI_GI4=Array(LM_GI_GI4_Bumper1_Ring, LM_GI_GI4_Layer_1, LM_GI_GI4_Layer_2, LM_GI_GI4_Parts, LM_GI_GI4_Playfield, LM_GI_GI4_Rflip1_001, LM_GI_GI4_Rflip1U, LM_GI_GI4_URsling1, LM_GI_GI4_URsling2, LM_GI_GI4_sw41, LM_GI_GI4_sw44_001, LM_GI_GI4_sw54_001, LM_GI_GI4_sw63_001, LM_GI_GI4_sw64_001)
Dim BL_GI_GI5: BL_GI_GI5=Array(LM_GI_GI5_Layer_1, LM_GI_GI5_Layer_2, LM_GI_GI5_Parts, LM_GI_GI5_Playfield, LM_GI_GI5_URsling1, LM_GI_GI5_URsling2, LM_GI_GI5_sw41, LM_GI_GI5_sw43_001, LM_GI_GI5_sw44_001, LM_GI_GI5_sw51, LM_GI_GI5_sw53_001, LM_GI_GI5_sw54_001, LM_GI_GI5_sw61, LM_GI_GI5_sw63_001, LM_GI_GI5_sw64_001)
Dim BL_GI_GI6: BL_GI_GI6=Array(LM_GI_GI6_Layer_1, LM_GI_GI6_Layer_2, LM_GI_GI6_Parts, LM_GI_GI6_Playfield, LM_GI_GI6_RFlipU, LM_GI_GI6_lsling1, LM_GI_GI6_lsling2, LM_GI_GI6_rsling1, LM_GI_GI6_rsling2, LM_GI_GI6_sw41, LM_GI_GI6_sw51, LM_GI_GI6_sw61, LM_GI_GI6_sw75)
Dim BL_GI_GI7: BL_GI_GI7=Array(LM_GI_GI7_Layer_1, LM_GI_GI7_Lflip_001, LM_GI_GI7_LflipU, LM_GI_GI7_Parts, LM_GI_GI7_Playfield, LM_GI_GI7_Under_PF, LM_GI_GI7_lsling1, LM_GI_GI7_lsling2, LM_GI_GI7_sw43_001, LM_GI_GI7_sw45, LM_GI_GI7_sw53_001, LM_GI_GI7_sw55, LM_GI_GI7_sw65)
Dim BL_GI_Llane1: BL_GI_Llane1=Array(LM_GI_Llane1_Parts, LM_GI_Llane1_Playfield, LM_GI_Llane1_sw42, LM_GI_Llane1_sw52)
Dim BL_GI_Llane2: BL_GI_Llane2=Array(LM_GI_Llane2_Parts, LM_GI_Llane2_Playfield, LM_GI_Llane2_sw52, LM_GI_Llane2_sw62)
Dim BL_GI_Llsling1: BL_GI_Llsling1=Array(LM_GI_Llsling1_Bumper1_Ring, LM_GI_Llsling1_LFlip1_001, LM_GI_Llsling1_LFlip1U, LM_GI_Llsling1_Layer_1, LM_GI_Llsling1_Parts, LM_GI_Llsling1_Playfield, LM_GI_Llsling1_Rflip1_001, LM_GI_Llsling1_Rflip1U, LM_GI_Llsling1_Under_PF, LM_GI_Llsling1_sw40, LM_GI_Llsling1_sw50, LM_GI_Llsling1_sw60)
Dim BL_GI_Llsling2: BL_GI_Llsling2=Array(LM_GI_Llsling2_Gate5_Wire, LM_GI_Llsling2_Layer_1, LM_GI_Llsling2_Parts, LM_GI_Llsling2_Playfield, LM_GI_Llsling2_Under_PF, LM_GI_Llsling2_sw41, LM_GI_Llsling2_sw43_001, LM_GI_Llsling2_sw44_001, LM_GI_Llsling2_sw51, LM_GI_Llsling2_sw53_001, LM_GI_Llsling2_sw54_001, LM_GI_Llsling2_sw61, LM_GI_Llsling2_sw63_001, LM_GI_Llsling2_sw64_001)
Dim BL_GI_al6a: BL_GI_al6a=Array(LM_GI_al6a_Bumper1_Ring, LM_GI_al6a_Layer_1, LM_GI_al6a_Layer_2, LM_GI_al6a_Parts, LM_GI_al6a_Playfield, LM_GI_al6a_Under_PF, LM_GI_al6a_sw64_001)
Dim BL_L_BumperLight1: BL_L_BumperLight1=Array(LM_L_BumperLight1_Bumper1_Ring, LM_L_BumperLight1_Layer_1, LM_L_BumperLight1_Parts, LM_L_BumperLight1_Playfield)
Dim BL_L_BumperLightA1: BL_L_BumperLightA1=Array(LM_L_BumperLightA1_Bumper1_Ring, LM_L_BumperLightA1_Gate8_Wire, LM_L_BumperLightA1_Layer_1, LM_L_BumperLightA1_Layer_2, LM_L_BumperLightA1_Parts, LM_L_BumperLightA1_Playfield, LM_L_BumperLightA1_URsling1, LM_L_BumperLightA1_URsling2, LM_L_BumperLightA1_Under_PF, LM_L_BumperLightA1_sw62, LM_L_BumperLightA1_sw73a_001)
Dim BL_L_L2: BL_L_L2=Array(LM_L_L2_Parts)
Dim BL_L_La32: BL_L_La32=Array(LM_L_La32_Parts, LM_L_La32_Playfield, LM_L_La32_RFlip_001, LM_L_La32_RFlipU, LM_L_La32_Under_PF)
Dim BL_L_La33: BL_L_La33=Array(LM_L_La33_Parts, LM_L_La33_Playfield, LM_L_La33_Under_PF)
Dim BL_L_La34: BL_L_La34=Array(LM_L_La34_Parts, LM_L_La34_Playfield, LM_L_La34_Under_PF)
Dim BL_L_La35: BL_L_La35=Array(LM_L_La35_Parts, LM_L_La35_Playfield, LM_L_La35_Under_PF, LM_L_La35_rsling1)
Dim BL_L_La5: BL_L_La5=Array(LM_L_La5_Playfield, LM_L_La5_Under_PF)
Dim BL_L_La6: BL_L_La6=Array(LM_L_La6_Layer_1, LM_L_La6_Parts, LM_L_La6_Playfield, LM_L_La6_RFlip_001, LM_L_La6_Under_PF, LM_L_La6_lsling1, LM_L_La6_lsling2)
Dim BL_L_La7: BL_L_La7=Array(LM_L_La7_Layer_1, LM_L_La7_Parts, LM_L_La7_Playfield, LM_L_La7_Under_PF)
Dim BL_L_La8: BL_L_La8=Array(LM_L_La8_Layer_1, LM_L_La8_Parts, LM_L_La8_Playfield, LM_L_La8_Under_PF, LM_L_La8_lsling1)
Dim BL_L_LkickL: BL_L_LkickL=Array(LM_L_LkickL_Parts, LM_L_LkickL_Under_PF)
Dim BL_L_LkickR: BL_L_LkickR=Array(LM_L_LkickR_Parts, LM_L_LkickR_Playfield, LM_L_LkickR_Under_PF)
Dim BL_L_al1: BL_L_al1=Array(LM_L_al1_Parts, LM_L_al1_Playfield, LM_L_al1_Under_PF)
Dim BL_L_al10: BL_L_al10=Array(LM_L_al10_Parts, LM_L_al10_Playfield, LM_L_al10_Under_PF)
Dim BL_L_al10a: BL_L_al10a=Array(LM_L_al10a_Bumper1_Ring, LM_L_al10a_Parts, LM_L_al10a_Playfield, LM_L_al10a_Under_PF)
Dim BL_L_al1a: BL_L_al1a=Array(LM_L_al1a_Bumper1_Ring, LM_L_al1a_Parts, LM_L_al1a_Playfield, LM_L_al1a_Under_PF)
Dim BL_L_al2: BL_L_al2=Array(LM_L_al2_Playfield, LM_L_al2_Under_PF)
Dim BL_L_al2a: BL_L_al2a=Array(LM_L_al2a_Bumper1_Ring, LM_L_al2a_Layer_1, LM_L_al2a_Parts, LM_L_al2a_Playfield, LM_L_al2a_Under_PF)
Dim BL_L_al3: BL_L_al3=Array(LM_L_al3_Playfield, LM_L_al3_Under_PF)
Dim BL_L_al3a: BL_L_al3a=Array(LM_L_al3a_Bumper1_Ring, LM_L_al3a_Layer_1, LM_L_al3a_Parts, LM_L_al3a_Playfield, LM_L_al3a_Under_PF)
Dim BL_L_al4: BL_L_al4=Array(LM_L_al4_Playfield, LM_L_al4_Under_PF, LM_L_al4_sw54_001, LM_L_al4_sw64_001)
Dim BL_L_al4a: BL_L_al4a=Array(LM_L_al4a_Bumper1_Ring, LM_L_al4a_Layer_1, LM_L_al4a_Parts, LM_L_al4a_Playfield, LM_L_al4a_Under_PF)
Dim BL_L_al5: BL_L_al5=Array(LM_L_al5_Layer_1, LM_L_al5_Parts, LM_L_al5_Playfield, LM_L_al5_Under_PF, LM_L_al5_sw64_001)
Dim BL_L_al5a: BL_L_al5a=Array(LM_L_al5a_Bumper1_Ring, LM_L_al5a_Layer_1, LM_L_al5a_Parts, LM_L_al5a_Playfield, LM_L_al5a_Under_PF)
Dim BL_L_al6: BL_L_al6=Array(LM_L_al6_Parts, LM_L_al6_Playfield, LM_L_al6_Under_PF, LM_L_al6_sw64_001)
Dim BL_L_al7: BL_L_al7=Array(LM_L_al7_Parts, LM_L_al7_Playfield, LM_L_al7_Under_PF)
Dim BL_L_al7a: BL_L_al7a=Array(LM_L_al7a_Bumper1_Ring, LM_L_al7a_Layer_1, LM_L_al7a_Parts, LM_L_al7a_Playfield, LM_L_al7a_Under_PF)
Dim BL_L_al8: BL_L_al8=Array(LM_L_al8_Layer_1, LM_L_al8_Parts, LM_L_al8_Playfield, LM_L_al8_Under_PF)
Dim BL_L_al8a: BL_L_al8a=Array(LM_L_al8a_Bumper1_Ring, LM_L_al8a_Layer_1, LM_L_al8a_Parts, LM_L_al8a_Playfield, LM_L_al8a_Under_PF)
Dim BL_L_al9: BL_L_al9=Array(LM_L_al9_Layer_1, LM_L_al9_Parts, LM_L_al9_Under_PF)
Dim BL_L_al9a: BL_L_al9a=Array(LM_L_al9a_Bumper1_Ring, LM_L_al9a_Parts, LM_L_al9a_Playfield, LM_L_al9a_Under_PF, LM_L_al9a_sw73a_001)
Dim BL_L_l10: BL_L_l10=Array(LM_L_l10_Parts, LM_L_l10_Playfield, LM_L_l10_Under_PF)
Dim BL_L_l11: BL_L_l11=Array(LM_L_l11_Parts, LM_L_l11_Playfield, LM_L_l11_Under_PF, LM_L_l11_sw43_001)
Dim BL_L_l12: BL_L_l12=Array(LM_L_l12_Gate7_Wire, LM_L_l12_Parts, LM_L_l12_Playfield, LM_L_l12_URsling1, LM_L_l12_URsling2, LM_L_l12_Under_PF, LM_L_l12_lsling1, LM_L_l12_lsling2, LM_L_l12_rsling1, LM_L_l12_rsling2)
Dim BL_L_l14: BL_L_l14=Array(LM_L_l14_Parts, LM_L_l14_Playfield, LM_L_l14_Under_PF)
Dim BL_L_l14a: BL_L_l14a=Array(LM_L_l14a_Parts, LM_L_l14a_Playfield, LM_L_l14a_Under_PF)
Dim BL_L_l15: BL_L_l15=Array(LM_L_l15_Parts, LM_L_l15_Playfield, LM_L_l15_Under_PF)
Dim BL_L_l15a: BL_L_l15a=Array(LM_L_l15a_Parts, LM_L_l15a_Playfield, LM_L_l15a_Under_PF)
Dim BL_L_l16: BL_L_l16=Array(LM_L_l16_Layer_1, LM_L_l16_Parts, LM_L_l16_Playfield, LM_L_l16_Under_PF)
Dim BL_L_l16a: BL_L_l16a=Array(LM_L_l16a_Parts, LM_L_l16a_Playfield, LM_L_l16a_Under_PF)
Dim BL_L_l17: BL_L_l17=Array(LM_L_l17_Parts, LM_L_l17_Playfield, LM_L_l17_Under_PF, LM_L_l17_sw40, LM_L_l17_sw50, LM_L_l17_sw60)
Dim BL_L_l18: BL_L_l18=Array(LM_L_l18_Layer_1, LM_L_l18_Parts, LM_L_l18_Playfield, LM_L_l18_Under_PF, LM_L_l18_sw50, LM_L_l18_sw60)
Dim BL_L_l19: BL_L_l19=Array(LM_L_l19_Layer_1, LM_L_l19_Parts, LM_L_l19_Playfield, LM_L_l19_Under_PF, LM_L_l19_sw64_001)
Dim BL_L_l19a: BL_L_l19a=Array(LM_L_l19a_Layer_1, LM_L_l19a_Layer_2, LM_L_l19a_Parts, LM_L_l19a_Playfield, LM_L_l19a_Under_PF, LM_L_l19a_sw54_001, LM_L_l19a_sw64_001)
Dim BL_L_l20: BL_L_l20=Array(LM_L_l20_Layer_1, LM_L_l20_Parts, LM_L_l20_Playfield, LM_L_l20_Under_PF, LM_L_l20_sw43_001, LM_L_l20_sw53_001, LM_L_l20_sw63_001)
Dim BL_L_l21: BL_L_l21=Array(LM_L_l21_Layer_1, LM_L_l21_Parts, LM_L_l21_Playfield, LM_L_l21_Under_PF, LM_L_l21_sw43_001, LM_L_l21_sw53_001, LM_L_l21_sw63_001)
Dim BL_L_l22: BL_L_l22=Array(LM_L_l22_Layer_1, LM_L_l22_Parts, LM_L_l22_Playfield, LM_L_l22_Under_PF, LM_L_l22_sw43_001, LM_L_l22_sw53_001, LM_L_l22_sw63_001)
Dim BL_L_l23: BL_L_l23=Array(LM_L_l23_Parts, LM_L_l23_Playfield, LM_L_l23_Under_PF, LM_L_l23_sw40, LM_L_l23_sw50, LM_L_l23_sw73)
Dim BL_L_l23a: BL_L_l23a=Array(LM_L_l23a_Bumper1_Ring, LM_L_l23a_Layer_1, LM_L_l23a_Parts, LM_L_l23a_Playfield, LM_L_l23a_Under_PF, LM_L_l23a_sw73a_001)
Dim BL_L_l24: BL_L_l24=Array(LM_L_l24_Layer_1, LM_L_l24_Parts, LM_L_l24_Playfield, LM_L_l24_Under_PF, LM_L_l24_sw44_001, LM_L_l24_sw54_001, LM_L_l24_sw63_001, LM_L_l24_sw64_001)
Dim BL_L_l25: BL_L_l25=Array(LM_L_l25_Parts, LM_L_l25_Playfield, LM_L_l25_Under_PF, LM_L_l25_sw44_001, LM_L_l25_sw54_001, LM_L_l25_sw63_001, LM_L_l25_sw64_001)
Dim BL_L_l26: BL_L_l26=Array(LM_L_l26_Parts, LM_L_l26_Playfield, LM_L_l26_Under_PF, LM_L_l26_sw44_001, LM_L_l26_sw54_001, LM_L_l26_sw63_001, LM_L_l26_sw64_001)
Dim BL_L_l27: BL_L_l27=Array(LM_L_l27_LFlip1_001, LM_L_l27_LFlip1U, LM_L_l27_Layer_1, LM_L_l27_Parts, LM_L_l27_Playfield, LM_L_l27_Rflip1_001, LM_L_l27_Rflip1U, LM_L_l27_Under_PF)
Dim BL_L_l28: BL_L_l28=Array(LM_L_l28_Playfield, LM_L_l28_Under_PF)
Dim BL_L_l29: BL_L_l29=Array(LM_L_l29_Parts, LM_L_l29_Playfield, LM_L_l29_Under_PF)
Dim BL_L_l3: BL_L_l3=Array(LM_L_l3_Parts, LM_L_l3_Playfield, LM_L_l3_Under_PF)
Dim BL_L_l30: BL_L_l30=Array(LM_L_l30_Parts, LM_L_l30_Playfield, LM_L_l30_Under_PF)
Dim BL_L_l31: BL_L_l31=Array(LM_L_l31_Playfield, LM_L_l31_Under_PF, LM_L_l31_sw40, LM_L_l31_sw50, LM_L_l31_sw60)
Dim BL_L_l36: BL_L_l36=Array(LM_L_l36_Layer_1, LM_L_l36_Parts, LM_L_l36_Playfield, LM_L_l36_Under_PF, LM_L_l36_rsling1, LM_L_l36_rsling2, LM_L_l36_sw61)
Dim BL_L_l37: BL_L_l37=Array(LM_L_l37_Layer_1, LM_L_l37_Lflip_001, LM_L_l37_Parts, LM_L_l37_Playfield, LM_L_l37_RFlip_001, LM_L_l37_Under_PF, LM_L_l37_lsling1, LM_L_l37_lsling2, LM_L_l37_sw43_001, LM_L_l37_sw53_001)
Dim BL_L_l38: BL_L_l38=Array(LM_L_l38_Layer_1, LM_L_l38_Lflip_001, LM_L_l38_Parts, LM_L_l38_Playfield, LM_L_l38_Under_PF, LM_L_l38_sw43_001, LM_L_l38_sw53_001, LM_L_l38_sw63_001)
Dim BL_L_l39: BL_L_l39=Array(LM_L_l39_Layer_1, LM_L_l39_Lflip_001, LM_L_l39_Parts, LM_L_l39_Playfield, LM_L_l39_Under_PF, LM_L_l39_rsling1, LM_L_l39_rsling2, LM_L_l39_sw43_001, LM_L_l39_sw53_001, LM_L_l39_sw61, LM_L_l39_sw63_001)
Dim BL_L_l40: BL_L_l40=Array(LM_L_l40_Layer_1, LM_L_l40_Lflip_001, LM_L_l40_Parts, LM_L_l40_Playfield, LM_L_l40_RFlip_001, LM_L_l40_Under_PF, LM_L_l40_sw43_001)
Dim BL_L_l41: BL_L_l41=Array(LM_L_l41_Lflip_001, LM_L_l41_LflipU, LM_L_l41_Parts, LM_L_l41_Playfield, LM_L_l41_RFlip_001, LM_L_l41_RFlipU, LM_L_l41_Under_PF)
Dim BL_L_l42: BL_L_l42=Array(LM_L_l42_Lflip_001, LM_L_l42_LflipU, LM_L_l42_Parts, LM_L_l42_Playfield, LM_L_l42_RFlip_001, LM_L_l42_RFlipU, LM_L_l42_Under_PF)
Dim BL_L_l43: BL_L_l43=Array(LM_L_l43_Lflip_001, LM_L_l43_LflipU, LM_L_l43_Parts, LM_L_l43_Playfield, LM_L_l43_RFlip_001, LM_L_l43_RFlipU, LM_L_l43_Under_PF)
Dim BL_L_l44: BL_L_l44=Array(LM_L_l44_Lflip_001, LM_L_l44_LflipU, LM_L_l44_Parts, LM_L_l44_Playfield, LM_L_l44_RFlip_001, LM_L_l44_RFlipU, LM_L_l44_Under_PF)
Dim BL_L_l45: BL_L_l45=Array(LM_L_l45_Lflip_001, LM_L_l45_LflipU, LM_L_l45_Parts, LM_L_l45_Playfield, LM_L_l45_RFlip_001, LM_L_l45_RFlipU, LM_L_l45_Under_PF)
Dim BL_L_l46: BL_L_l46=Array(LM_L_l46_Layer_2, LM_L_l46_Lflip_001, LM_L_l46_LflipU, LM_L_l46_Parts, LM_L_l46_Playfield, LM_L_l46_RFlip_001, LM_L_l46_RFlipU, LM_L_l46_Under_PF)
Dim BL_L_l47: BL_L_l47=Array(LM_L_l47_Parts, LM_L_l47_Playfield, LM_L_l47_URsling1, LM_L_l47_URsling2, LM_L_l47_Under_PF, LM_L_l47_rsling1, LM_L_l47_sw41, LM_L_l47_sw51)
Dim BL_L_l47a: BL_L_l47a=Array(LM_L_l47a_Parts, LM_L_l47a_Playfield, LM_L_l47a_Under_PF, LM_L_l47a_sw40, LM_L_l47a_sw50, LM_L_l47a_sw60)
Dim BL_L_l48: BL_L_l48=Array(LM_L_l48_Gate7_Wire, LM_L_l48_Parts, LM_L_l48_Playfield, LM_L_l48_Under_PF, LM_L_l48_sw42)
Dim BL_L_l49: BL_L_l49=Array(LM_L_l49_Gate7_Wire, LM_L_l49_Layer_1, LM_L_l49_Parts, LM_L_l49_Playfield, LM_L_l49_Under_PF, LM_L_l49_sw52)
Dim BL_L_l50: BL_L_l50=Array(LM_L_l50_Layer_1, LM_L_l50_Parts, LM_L_l50_Playfield, LM_L_l50_Under_PF, LM_L_l50_sw62)
Dim BL_L_l9: BL_L_l9=Array(LM_L_l9_Layer_1, LM_L_l9_Parts, LM_L_l9_Playfield, LM_L_l9_Under_PF, LM_L_l9_sw43_001, LM_L_l9_sw53_001, LM_L_l9_sw63_001)
Dim BL_World: BL_World=Array(BM_Bumper1_Ring, BM_Gate1_Wire, BM_Gate3_Wire, BM_Gate4_Wire, BM_Gate5_Wire, BM_Gate7_Wire, BM_Gate8_Wire, BM_LFlip1_001, BM_LFlip1U, BM_Layer_1, BM_Layer_2, BM_Lflip_001, BM_LflipU, BM_Parts, BM_Pgate, BM_Playfield, BM_RFlip_001, BM_RFlipU, BM_Rflip1_001, BM_Rflip1U, BM_URsling1, BM_URsling2, BM_Under_PF, BM_lsling1, BM_lsling2, BM_rsling1, BM_rsling2, BM_sw31, BM_sw40, BM_sw41, BM_sw42, BM_sw43_001, BM_sw44_001, BM_sw45, BM_sw50, BM_sw51, BM_sw52, BM_sw53_001, BM_sw54_001, BM_sw55, BM_sw60, BM_sw61, BM_sw62, BM_sw63_001, BM_sw64_001, BM_sw65, BM_sw73, BM_sw73a_001, BM_sw75)
' Global arrays
Dim BG_Bakemap: BG_Bakemap=Array(BM_Bumper1_Ring, BM_Gate1_Wire, BM_Gate3_Wire, BM_Gate4_Wire, BM_Gate5_Wire, BM_Gate7_Wire, BM_Gate8_Wire, BM_LFlip1_001, BM_LFlip1U, BM_Layer_1, BM_Layer_2, BM_Lflip_001, BM_LflipU, BM_Parts, BM_Pgate, BM_Playfield, BM_RFlip_001, BM_RFlipU, BM_Rflip1_001, BM_Rflip1U, BM_URsling1, BM_URsling2, BM_Under_PF, BM_lsling1, BM_lsling2, BM_rsling1, BM_rsling2, BM_sw31, BM_sw40, BM_sw41, BM_sw42, BM_sw43_001, BM_sw44_001, BM_sw45, BM_sw50, BM_sw51, BM_sw52, BM_sw53_001, BM_sw54_001, BM_sw55, BM_sw60, BM_sw61, BM_sw62, BM_sw63_001, BM_sw64_001, BM_sw65, BM_sw73, BM_sw73a_001, BM_sw75)
Dim BG_Lightmap: BG_Lightmap=Array(LM_F_Bumper1_Ring, LM_F_Gate8_Wire, LM_F_Layer_1, LM_F_Layer_2, LM_F_Parts, LM_F_Playfield, LM_F_URsling1, LM_F_URsling2, LM_F_Under_PF, LM_F_rsling1, LM_F_sw44_001, LM_F_sw54_001, LM_F_sw64_001, LM_F_sw73a_001, LM_GI_GI1_LFlip1_001, LM_GI_GI1_LFlip1U, LM_GI_GI1_Layer_1, LM_GI_GI1_Parts, LM_GI_GI1_Playfield, LM_GI_GI1_Rflip1_001, LM_GI_GI1_Rflip1U, LM_GI_GI1_Under_PF, LM_GI_GI1_sw45, LM_GI_GI2_Gate8_Wire, LM_GI_GI2_Parts, LM_GI_GI2_Playfield, LM_GI_GI3_Parts, LM_GI_GI3_Playfield, LM_GI_GI3_Under_PF, LM_GI_GI3_sw40, LM_GI_GI3_sw42, LM_GI_GI3_sw50, LM_GI_GI3_sw52, LM_GI_GI3_sw60, LM_GI_GI3_sw62, LM_GI_GI4_Bumper1_Ring, LM_GI_GI4_Layer_1, LM_GI_GI4_Layer_2, LM_GI_GI4_Parts, LM_GI_GI4_Playfield, LM_GI_GI4_Rflip1_001, LM_GI_GI4_Rflip1U, LM_GI_GI4_URsling1, LM_GI_GI4_URsling2, LM_GI_GI4_sw41, LM_GI_GI4_sw44_001, LM_GI_GI4_sw54_001, LM_GI_GI4_sw63_001, LM_GI_GI4_sw64_001, LM_GI_GI5_Layer_1, LM_GI_GI5_Layer_2, LM_GI_GI5_Parts, LM_GI_GI5_Playfield, LM_GI_GI5_URsling1, _
  LM_GI_GI5_URsling2, LM_GI_GI5_sw41, LM_GI_GI5_sw43_001, LM_GI_GI5_sw44_001, LM_GI_GI5_sw51, LM_GI_GI5_sw53_001, LM_GI_GI5_sw54_001, LM_GI_GI5_sw61, LM_GI_GI5_sw63_001, LM_GI_GI5_sw64_001, LM_GI_GI6_Layer_1, LM_GI_GI6_Layer_2, LM_GI_GI6_Parts, LM_GI_GI6_Playfield, LM_GI_GI6_RFlipU, LM_GI_GI6_lsling1, LM_GI_GI6_lsling2, LM_GI_GI6_rsling1, LM_GI_GI6_rsling2, LM_GI_GI6_sw41, LM_GI_GI6_sw51, LM_GI_GI6_sw61, LM_GI_GI6_sw75, LM_GI_GI7_Layer_1, LM_GI_GI7_Lflip_001, LM_GI_GI7_LflipU, LM_GI_GI7_Parts, LM_GI_GI7_Playfield, LM_GI_GI7_Under_PF, LM_GI_GI7_lsling1, LM_GI_GI7_lsling2, LM_GI_GI7_sw43_001, LM_GI_GI7_sw45, LM_GI_GI7_sw53_001, LM_GI_GI7_sw55, LM_GI_GI7_sw65, LM_GI_Llane1_Parts, LM_GI_Llane1_Playfield, LM_GI_Llane1_sw42, LM_GI_Llane1_sw52, LM_GI_Llane2_Parts, LM_GI_Llane2_Playfield, LM_GI_Llane2_sw52, LM_GI_Llane2_sw62, LM_GI_Llsling1_Bumper1_Ring, LM_GI_Llsling1_LFlip1_001, LM_GI_Llsling1_LFlip1U, LM_GI_Llsling1_Layer_1, LM_GI_Llsling1_Parts, LM_GI_Llsling1_Playfield, LM_GI_Llsling1_Rflip1_001, _
  LM_GI_Llsling1_Rflip1U, LM_GI_Llsling1_Under_PF, LM_GI_Llsling1_sw40, LM_GI_Llsling1_sw50, LM_GI_Llsling1_sw60, LM_GI_Llsling2_Gate5_Wire, LM_GI_Llsling2_Layer_1, LM_GI_Llsling2_Parts, LM_GI_Llsling2_Playfield, LM_GI_Llsling2_Under_PF, LM_GI_Llsling2_sw41, LM_GI_Llsling2_sw43_001, LM_GI_Llsling2_sw44_001, LM_GI_Llsling2_sw51, LM_GI_Llsling2_sw53_001, LM_GI_Llsling2_sw54_001, LM_GI_Llsling2_sw61, LM_GI_Llsling2_sw63_001, LM_GI_Llsling2_sw64_001, LM_GI_al6a_Bumper1_Ring, LM_GI_al6a_Layer_1, LM_GI_al6a_Layer_2, LM_GI_al6a_Parts, LM_GI_al6a_Playfield, LM_GI_al6a_Under_PF, LM_GI_al6a_sw64_001, LM_L_BumperLight1_Bumper1_Ring, LM_L_BumperLight1_Layer_1, LM_L_BumperLight1_Parts, LM_L_BumperLight1_Playfield, LM_L_BumperLightA1_Bumper1_Ring, LM_L_BumperLightA1_Gate8_Wire, LM_L_BumperLightA1_Layer_1, LM_L_BumperLightA1_Layer_2, LM_L_BumperLightA1_Parts, LM_L_BumperLightA1_Playfield, LM_L_BumperLightA1_URsling1, LM_L_BumperLightA1_URsling2, LM_L_BumperLightA1_Under_PF, LM_L_BumperLightA1_sw62, _
  LM_L_BumperLightA1_sw73a_001, LM_L_L2_Parts, LM_L_La32_Parts, LM_L_La32_Playfield, LM_L_La32_RFlip_001, LM_L_La32_RFlipU, LM_L_La32_Under_PF, LM_L_La33_Parts, LM_L_La33_Playfield, LM_L_La33_Under_PF, LM_L_La34_Parts, LM_L_La34_Playfield, LM_L_La34_Under_PF, LM_L_La35_Parts, LM_L_La35_Playfield, LM_L_La35_Under_PF, LM_L_La35_rsling1, LM_L_La5_Playfield, LM_L_La5_Under_PF, LM_L_La6_Layer_1, LM_L_La6_Parts, LM_L_La6_Playfield, LM_L_La6_RFlip_001, LM_L_La6_Under_PF, LM_L_La6_lsling1, LM_L_La6_lsling2, LM_L_La7_Layer_1, LM_L_La7_Parts, LM_L_La7_Playfield, LM_L_La7_Under_PF, LM_L_La8_Layer_1, LM_L_La8_Parts, LM_L_La8_Playfield, LM_L_La8_Under_PF, LM_L_La8_lsling1, LM_L_LkickL_Parts, LM_L_LkickL_Under_PF, LM_L_LkickR_Parts, LM_L_LkickR_Playfield, LM_L_LkickR_Under_PF, LM_L_al1_Parts, LM_L_al1_Playfield, LM_L_al1_Under_PF, LM_L_al10_Parts, LM_L_al10_Playfield, LM_L_al10_Under_PF, LM_L_al10a_Bumper1_Ring, LM_L_al10a_Parts, LM_L_al10a_Playfield, LM_L_al10a_Under_PF, LM_L_al1a_Bumper1_Ring, LM_L_al1a_Parts, _
  LM_L_al1a_Playfield, LM_L_al1a_Under_PF, LM_L_al2_Playfield, LM_L_al2_Under_PF, LM_L_al2a_Bumper1_Ring, LM_L_al2a_Layer_1, LM_L_al2a_Parts, LM_L_al2a_Playfield, LM_L_al2a_Under_PF, LM_L_al3_Playfield, LM_L_al3_Under_PF, LM_L_al3a_Bumper1_Ring, LM_L_al3a_Layer_1, LM_L_al3a_Parts, LM_L_al3a_Playfield, LM_L_al3a_Under_PF, LM_L_al4_Playfield, LM_L_al4_Under_PF, LM_L_al4_sw54_001, LM_L_al4_sw64_001, LM_L_al4a_Bumper1_Ring, LM_L_al4a_Layer_1, LM_L_al4a_Parts, LM_L_al4a_Playfield, LM_L_al4a_Under_PF, LM_L_al5_Layer_1, LM_L_al5_Parts, LM_L_al5_Playfield, LM_L_al5_Under_PF, LM_L_al5_sw64_001, LM_L_al5a_Bumper1_Ring, LM_L_al5a_Layer_1, LM_L_al5a_Parts, LM_L_al5a_Playfield, LM_L_al5a_Under_PF, LM_L_al6_Parts, LM_L_al6_Playfield, LM_L_al6_Under_PF, LM_L_al6_sw64_001, LM_L_al7_Parts, LM_L_al7_Playfield, LM_L_al7_Under_PF, LM_L_al7a_Bumper1_Ring, LM_L_al7a_Layer_1, LM_L_al7a_Parts, LM_L_al7a_Playfield, LM_L_al7a_Under_PF, LM_L_al8_Layer_1, LM_L_al8_Parts, LM_L_al8_Playfield, LM_L_al8_Under_PF, LM_L_al8a_Bumper1_Ring, _
  LM_L_al8a_Layer_1, LM_L_al8a_Parts, LM_L_al8a_Playfield, LM_L_al8a_Under_PF, LM_L_al9_Layer_1, LM_L_al9_Parts, LM_L_al9_Under_PF, LM_L_al9a_Bumper1_Ring, LM_L_al9a_Parts, LM_L_al9a_Playfield, LM_L_al9a_Under_PF, LM_L_al9a_sw73a_001, LM_L_l10_Parts, LM_L_l10_Playfield, LM_L_l10_Under_PF, LM_L_l11_Parts, LM_L_l11_Playfield, LM_L_l11_Under_PF, LM_L_l11_sw43_001, LM_L_l12_Gate7_Wire, LM_L_l12_Parts, LM_L_l12_Playfield, LM_L_l12_URsling1, LM_L_l12_URsling2, LM_L_l12_Under_PF, LM_L_l12_lsling1, LM_L_l12_lsling2, LM_L_l12_rsling1, LM_L_l12_rsling2, LM_L_l14_Parts, LM_L_l14_Playfield, LM_L_l14_Under_PF, LM_L_l14a_Parts, LM_L_l14a_Playfield, LM_L_l14a_Under_PF, LM_L_l15_Parts, LM_L_l15_Playfield, LM_L_l15_Under_PF, LM_L_l15a_Parts, LM_L_l15a_Playfield, LM_L_l15a_Under_PF, LM_L_l16_Layer_1, LM_L_l16_Parts, LM_L_l16_Playfield, LM_L_l16_Under_PF, LM_L_l16a_Parts, LM_L_l16a_Playfield, LM_L_l16a_Under_PF, LM_L_l17_Parts, LM_L_l17_Playfield, LM_L_l17_Under_PF, LM_L_l17_sw40, LM_L_l17_sw50, LM_L_l17_sw60, LM_L_l18_Layer_1, _
  LM_L_l18_Parts, LM_L_l18_Playfield, LM_L_l18_Under_PF, LM_L_l18_sw50, LM_L_l18_sw60, LM_L_l19_Layer_1, LM_L_l19_Parts, LM_L_l19_Playfield, LM_L_l19_Under_PF, LM_L_l19_sw64_001, LM_L_l19a_Layer_1, LM_L_l19a_Layer_2, LM_L_l19a_Parts, LM_L_l19a_Playfield, LM_L_l19a_Under_PF, LM_L_l19a_sw54_001, LM_L_l19a_sw64_001, LM_L_l20_Layer_1, LM_L_l20_Parts, LM_L_l20_Playfield, LM_L_l20_Under_PF, LM_L_l20_sw43_001, LM_L_l20_sw53_001, LM_L_l20_sw63_001, LM_L_l21_Layer_1, LM_L_l21_Parts, LM_L_l21_Playfield, LM_L_l21_Under_PF, LM_L_l21_sw43_001, LM_L_l21_sw53_001, LM_L_l21_sw63_001, LM_L_l22_Layer_1, LM_L_l22_Parts, LM_L_l22_Playfield, LM_L_l22_Under_PF, LM_L_l22_sw43_001, LM_L_l22_sw53_001, LM_L_l22_sw63_001, LM_L_l23_Parts, LM_L_l23_Playfield, LM_L_l23_Under_PF, LM_L_l23_sw40, LM_L_l23_sw50, LM_L_l23_sw73, LM_L_l23a_Bumper1_Ring, LM_L_l23a_Layer_1, LM_L_l23a_Parts, LM_L_l23a_Playfield, LM_L_l23a_Under_PF, LM_L_l23a_sw73a_001, LM_L_l24_Layer_1, LM_L_l24_Parts, LM_L_l24_Playfield, LM_L_l24_Under_PF, LM_L_l24_sw44_001, _
  LM_L_l24_sw54_001, LM_L_l24_sw63_001, LM_L_l24_sw64_001, LM_L_l25_Parts, LM_L_l25_Playfield, LM_L_l25_Under_PF, LM_L_l25_sw44_001, LM_L_l25_sw54_001, LM_L_l25_sw63_001, LM_L_l25_sw64_001, LM_L_l26_Parts, LM_L_l26_Playfield, LM_L_l26_Under_PF, LM_L_l26_sw44_001, LM_L_l26_sw54_001, LM_L_l26_sw63_001, LM_L_l26_sw64_001, LM_L_l27_LFlip1_001, LM_L_l27_LFlip1U, LM_L_l27_Layer_1, LM_L_l27_Parts, LM_L_l27_Playfield, LM_L_l27_Rflip1_001, LM_L_l27_Rflip1U, LM_L_l27_Under_PF, LM_L_l28_Playfield, LM_L_l28_Under_PF, LM_L_l29_Parts, LM_L_l29_Playfield, LM_L_l29_Under_PF, LM_L_l3_Parts, LM_L_l3_Playfield, LM_L_l3_Under_PF, LM_L_l30_Parts, LM_L_l30_Playfield, LM_L_l30_Under_PF, LM_L_l31_Playfield, LM_L_l31_Under_PF, LM_L_l31_sw40, LM_L_l31_sw50, LM_L_l31_sw60, LM_L_l36_Layer_1, LM_L_l36_Parts, LM_L_l36_Playfield, LM_L_l36_Under_PF, LM_L_l36_rsling1, LM_L_l36_rsling2, LM_L_l36_sw61, LM_L_l37_Layer_1, LM_L_l37_Lflip_001, LM_L_l37_Parts, LM_L_l37_Playfield, LM_L_l37_RFlip_001, LM_L_l37_Under_PF, LM_L_l37_lsling1, _
  LM_L_l37_lsling2, LM_L_l37_sw43_001, LM_L_l37_sw53_001, LM_L_l38_Layer_1, LM_L_l38_Lflip_001, LM_L_l38_Parts, LM_L_l38_Playfield, LM_L_l38_Under_PF, LM_L_l38_sw43_001, LM_L_l38_sw53_001, LM_L_l38_sw63_001, LM_L_l39_Layer_1, LM_L_l39_Lflip_001, LM_L_l39_Parts, LM_L_l39_Playfield, LM_L_l39_Under_PF, LM_L_l39_rsling1, LM_L_l39_rsling2, LM_L_l39_sw43_001, LM_L_l39_sw53_001, LM_L_l39_sw61, LM_L_l39_sw63_001, LM_L_l40_Layer_1, LM_L_l40_Lflip_001, LM_L_l40_Parts, LM_L_l40_Playfield, LM_L_l40_RFlip_001, LM_L_l40_Under_PF, LM_L_l40_sw43_001, LM_L_l41_Lflip_001, LM_L_l41_LflipU, LM_L_l41_Parts, LM_L_l41_Playfield, LM_L_l41_RFlip_001, LM_L_l41_RFlipU, LM_L_l41_Under_PF, LM_L_l42_Lflip_001, LM_L_l42_LflipU, LM_L_l42_Parts, LM_L_l42_Playfield, LM_L_l42_RFlip_001, LM_L_l42_RFlipU, LM_L_l42_Under_PF, LM_L_l43_Lflip_001, LM_L_l43_LflipU, LM_L_l43_Parts, LM_L_l43_Playfield, LM_L_l43_RFlip_001, LM_L_l43_RFlipU, LM_L_l43_Under_PF, LM_L_l44_Lflip_001, LM_L_l44_LflipU, LM_L_l44_Parts, LM_L_l44_Playfield, LM_L_l44_RFlip_001, _
  LM_L_l44_RFlipU, LM_L_l44_Under_PF, LM_L_l45_Lflip_001, LM_L_l45_LflipU, LM_L_l45_Parts, LM_L_l45_Playfield, LM_L_l45_RFlip_001, LM_L_l45_RFlipU, LM_L_l45_Under_PF, LM_L_l46_Layer_2, LM_L_l46_Lflip_001, LM_L_l46_LflipU, LM_L_l46_Parts, LM_L_l46_Playfield, LM_L_l46_RFlip_001, LM_L_l46_RFlipU, LM_L_l46_Under_PF, LM_L_l47_Parts, LM_L_l47_Playfield, LM_L_l47_URsling1, LM_L_l47_URsling2, LM_L_l47_Under_PF, LM_L_l47_rsling1, LM_L_l47_sw41, LM_L_l47_sw51, LM_L_l47a_Parts, LM_L_l47a_Playfield, LM_L_l47a_Under_PF, LM_L_l47a_sw40, LM_L_l47a_sw50, LM_L_l47a_sw60, LM_L_l48_Gate7_Wire, LM_L_l48_Parts, LM_L_l48_Playfield, LM_L_l48_Under_PF, LM_L_l48_sw42, LM_L_l49_Gate7_Wire, LM_L_l49_Layer_1, LM_L_l49_Parts, LM_L_l49_Playfield, LM_L_l49_Under_PF, LM_L_l49_sw52, LM_L_l50_Layer_1, LM_L_l50_Parts, LM_L_l50_Playfield, LM_L_l50_Under_PF, LM_L_l50_sw62, LM_L_l9_Layer_1, LM_L_l9_Parts, LM_L_l9_Playfield, LM_L_l9_Under_PF, LM_L_l9_sw43_001, LM_L_l9_sw53_001, LM_L_l9_sw63_001)
Dim BG_All: BG_All=Array(BM_Bumper1_Ring, BM_Gate1_Wire, BM_Gate3_Wire, BM_Gate4_Wire, BM_Gate5_Wire, BM_Gate7_Wire, BM_Gate8_Wire, BM_LFlip1_001, BM_LFlip1U, BM_Layer_1, BM_Layer_2, BM_Lflip_001, BM_LflipU, BM_Parts, BM_Pgate, BM_Playfield, BM_RFlip_001, BM_RFlipU, BM_Rflip1_001, BM_Rflip1U, BM_URsling1, BM_URsling2, BM_Under_PF, BM_lsling1, BM_lsling2, BM_rsling1, BM_rsling2, BM_sw31, BM_sw40, BM_sw41, BM_sw42, BM_sw43_001, BM_sw44_001, BM_sw45, BM_sw50, BM_sw51, BM_sw52, BM_sw53_001, BM_sw54_001, BM_sw55, BM_sw60, BM_sw61, BM_sw62, BM_sw63_001, BM_sw64_001, BM_sw65, BM_sw73, BM_sw73a_001, BM_sw75, LM_F_Bumper1_Ring, LM_F_Gate8_Wire, LM_F_Layer_1, LM_F_Layer_2, LM_F_Parts, LM_F_Playfield, LM_F_URsling1, LM_F_URsling2, LM_F_Under_PF, LM_F_rsling1, LM_F_sw44_001, LM_F_sw54_001, LM_F_sw64_001, LM_F_sw73a_001, LM_GI_GI1_LFlip1_001, LM_GI_GI1_LFlip1U, LM_GI_GI1_Layer_1, LM_GI_GI1_Parts, LM_GI_GI1_Playfield, LM_GI_GI1_Rflip1_001, LM_GI_GI1_Rflip1U, LM_GI_GI1_Under_PF, LM_GI_GI1_sw45, LM_GI_GI2_Gate8_Wire, _
  LM_GI_GI2_Parts, LM_GI_GI2_Playfield, LM_GI_GI3_Parts, LM_GI_GI3_Playfield, LM_GI_GI3_Under_PF, LM_GI_GI3_sw40, LM_GI_GI3_sw42, LM_GI_GI3_sw50, LM_GI_GI3_sw52, LM_GI_GI3_sw60, LM_GI_GI3_sw62, LM_GI_GI4_Bumper1_Ring, LM_GI_GI4_Layer_1, LM_GI_GI4_Layer_2, LM_GI_GI4_Parts, LM_GI_GI4_Playfield, LM_GI_GI4_Rflip1_001, LM_GI_GI4_Rflip1U, LM_GI_GI4_URsling1, LM_GI_GI4_URsling2, LM_GI_GI4_sw41, LM_GI_GI4_sw44_001, LM_GI_GI4_sw54_001, LM_GI_GI4_sw63_001, LM_GI_GI4_sw64_001, LM_GI_GI5_Layer_1, LM_GI_GI5_Layer_2, LM_GI_GI5_Parts, LM_GI_GI5_Playfield, LM_GI_GI5_URsling1, LM_GI_GI5_URsling2, LM_GI_GI5_sw41, LM_GI_GI5_sw43_001, LM_GI_GI5_sw44_001, LM_GI_GI5_sw51, LM_GI_GI5_sw53_001, LM_GI_GI5_sw54_001, LM_GI_GI5_sw61, LM_GI_GI5_sw63_001, LM_GI_GI5_sw64_001, LM_GI_GI6_Layer_1, LM_GI_GI6_Layer_2, LM_GI_GI6_Parts, LM_GI_GI6_Playfield, LM_GI_GI6_RFlipU, LM_GI_GI6_lsling1, LM_GI_GI6_lsling2, LM_GI_GI6_rsling1, LM_GI_GI6_rsling2, LM_GI_GI6_sw41, LM_GI_GI6_sw51, LM_GI_GI6_sw61, LM_GI_GI6_sw75, LM_GI_GI7_Layer_1, _
  LM_GI_GI7_Lflip_001, LM_GI_GI7_LflipU, LM_GI_GI7_Parts, LM_GI_GI7_Playfield, LM_GI_GI7_Under_PF, LM_GI_GI7_lsling1, LM_GI_GI7_lsling2, LM_GI_GI7_sw43_001, LM_GI_GI7_sw45, LM_GI_GI7_sw53_001, LM_GI_GI7_sw55, LM_GI_GI7_sw65, LM_GI_Llane1_Parts, LM_GI_Llane1_Playfield, LM_GI_Llane1_sw42, LM_GI_Llane1_sw52, LM_GI_Llane2_Parts, LM_GI_Llane2_Playfield, LM_GI_Llane2_sw52, LM_GI_Llane2_sw62, LM_GI_Llsling1_Bumper1_Ring, LM_GI_Llsling1_LFlip1_001, LM_GI_Llsling1_LFlip1U, LM_GI_Llsling1_Layer_1, LM_GI_Llsling1_Parts, LM_GI_Llsling1_Playfield, LM_GI_Llsling1_Rflip1_001, LM_GI_Llsling1_Rflip1U, LM_GI_Llsling1_Under_PF, LM_GI_Llsling1_sw40, LM_GI_Llsling1_sw50, LM_GI_Llsling1_sw60, LM_GI_Llsling2_Gate5_Wire, LM_GI_Llsling2_Layer_1, LM_GI_Llsling2_Parts, LM_GI_Llsling2_Playfield, LM_GI_Llsling2_Under_PF, LM_GI_Llsling2_sw41, LM_GI_Llsling2_sw43_001, LM_GI_Llsling2_sw44_001, LM_GI_Llsling2_sw51, LM_GI_Llsling2_sw53_001, LM_GI_Llsling2_sw54_001, LM_GI_Llsling2_sw61, LM_GI_Llsling2_sw63_001, LM_GI_Llsling2_sw64_001, _
  LM_GI_al6a_Bumper1_Ring, LM_GI_al6a_Layer_1, LM_GI_al6a_Layer_2, LM_GI_al6a_Parts, LM_GI_al6a_Playfield, LM_GI_al6a_Under_PF, LM_GI_al6a_sw64_001, LM_L_BumperLight1_Bumper1_Ring, LM_L_BumperLight1_Layer_1, LM_L_BumperLight1_Parts, LM_L_BumperLight1_Playfield, LM_L_BumperLightA1_Bumper1_Ring, LM_L_BumperLightA1_Gate8_Wire, LM_L_BumperLightA1_Layer_1, LM_L_BumperLightA1_Layer_2, LM_L_BumperLightA1_Parts, LM_L_BumperLightA1_Playfield, LM_L_BumperLightA1_URsling1, LM_L_BumperLightA1_URsling2, LM_L_BumperLightA1_Under_PF, LM_L_BumperLightA1_sw62, LM_L_BumperLightA1_sw73a_001, LM_L_L2_Parts, LM_L_La32_Parts, LM_L_La32_Playfield, LM_L_La32_RFlip_001, LM_L_La32_RFlipU, LM_L_La32_Under_PF, LM_L_La33_Parts, LM_L_La33_Playfield, LM_L_La33_Under_PF, LM_L_La34_Parts, LM_L_La34_Playfield, LM_L_La34_Under_PF, LM_L_La35_Parts, LM_L_La35_Playfield, LM_L_La35_Under_PF, LM_L_La35_rsling1, LM_L_La5_Playfield, LM_L_La5_Under_PF, LM_L_La6_Layer_1, LM_L_La6_Parts, LM_L_La6_Playfield, LM_L_La6_RFlip_001, LM_L_La6_Under_PF, _
  LM_L_La6_lsling1, LM_L_La6_lsling2, LM_L_La7_Layer_1, LM_L_La7_Parts, LM_L_La7_Playfield, LM_L_La7_Under_PF, LM_L_La8_Layer_1, LM_L_La8_Parts, LM_L_La8_Playfield, LM_L_La8_Under_PF, LM_L_La8_lsling1, LM_L_LkickL_Parts, LM_L_LkickL_Under_PF, LM_L_LkickR_Parts, LM_L_LkickR_Playfield, LM_L_LkickR_Under_PF, LM_L_al1_Parts, LM_L_al1_Playfield, LM_L_al1_Under_PF, LM_L_al10_Parts, LM_L_al10_Playfield, LM_L_al10_Under_PF, LM_L_al10a_Bumper1_Ring, LM_L_al10a_Parts, LM_L_al10a_Playfield, LM_L_al10a_Under_PF, LM_L_al1a_Bumper1_Ring, LM_L_al1a_Parts, LM_L_al1a_Playfield, LM_L_al1a_Under_PF, LM_L_al2_Playfield, LM_L_al2_Under_PF, LM_L_al2a_Bumper1_Ring, LM_L_al2a_Layer_1, LM_L_al2a_Parts, LM_L_al2a_Playfield, LM_L_al2a_Under_PF, LM_L_al3_Playfield, LM_L_al3_Under_PF, LM_L_al3a_Bumper1_Ring, LM_L_al3a_Layer_1, LM_L_al3a_Parts, LM_L_al3a_Playfield, LM_L_al3a_Under_PF, LM_L_al4_Playfield, LM_L_al4_Under_PF, LM_L_al4_sw54_001, LM_L_al4_sw64_001, LM_L_al4a_Bumper1_Ring, LM_L_al4a_Layer_1, LM_L_al4a_Parts, LM_L_al4a_Playfield, _
  LM_L_al4a_Under_PF, LM_L_al5_Layer_1, LM_L_al5_Parts, LM_L_al5_Playfield, LM_L_al5_Under_PF, LM_L_al5_sw64_001, LM_L_al5a_Bumper1_Ring, LM_L_al5a_Layer_1, LM_L_al5a_Parts, LM_L_al5a_Playfield, LM_L_al5a_Under_PF, LM_L_al6_Parts, LM_L_al6_Playfield, LM_L_al6_Under_PF, LM_L_al6_sw64_001, LM_L_al7_Parts, LM_L_al7_Playfield, LM_L_al7_Under_PF, LM_L_al7a_Bumper1_Ring, LM_L_al7a_Layer_1, LM_L_al7a_Parts, LM_L_al7a_Playfield, LM_L_al7a_Under_PF, LM_L_al8_Layer_1, LM_L_al8_Parts, LM_L_al8_Playfield, LM_L_al8_Under_PF, LM_L_al8a_Bumper1_Ring, LM_L_al8a_Layer_1, LM_L_al8a_Parts, LM_L_al8a_Playfield, LM_L_al8a_Under_PF, LM_L_al9_Layer_1, LM_L_al9_Parts, LM_L_al9_Under_PF, LM_L_al9a_Bumper1_Ring, LM_L_al9a_Parts, LM_L_al9a_Playfield, LM_L_al9a_Under_PF, LM_L_al9a_sw73a_001, LM_L_l10_Parts, LM_L_l10_Playfield, LM_L_l10_Under_PF, LM_L_l11_Parts, LM_L_l11_Playfield, LM_L_l11_Under_PF, LM_L_l11_sw43_001, LM_L_l12_Gate7_Wire, LM_L_l12_Parts, LM_L_l12_Playfield, LM_L_l12_URsling1, LM_L_l12_URsling2, LM_L_l12_Under_PF, _
  LM_L_l12_lsling1, LM_L_l12_lsling2, LM_L_l12_rsling1, LM_L_l12_rsling2, LM_L_l14_Parts, LM_L_l14_Playfield, LM_L_l14_Under_PF, LM_L_l14a_Parts, LM_L_l14a_Playfield, LM_L_l14a_Under_PF, LM_L_l15_Parts, LM_L_l15_Playfield, LM_L_l15_Under_PF, LM_L_l15a_Parts, LM_L_l15a_Playfield, LM_L_l15a_Under_PF, LM_L_l16_Layer_1, LM_L_l16_Parts, LM_L_l16_Playfield, LM_L_l16_Under_PF, LM_L_l16a_Parts, LM_L_l16a_Playfield, LM_L_l16a_Under_PF, LM_L_l17_Parts, LM_L_l17_Playfield, LM_L_l17_Under_PF, LM_L_l17_sw40, LM_L_l17_sw50, LM_L_l17_sw60, LM_L_l18_Layer_1, LM_L_l18_Parts, LM_L_l18_Playfield, LM_L_l18_Under_PF, LM_L_l18_sw50, LM_L_l18_sw60, LM_L_l19_Layer_1, LM_L_l19_Parts, LM_L_l19_Playfield, LM_L_l19_Under_PF, LM_L_l19_sw64_001, LM_L_l19a_Layer_1, LM_L_l19a_Layer_2, LM_L_l19a_Parts, LM_L_l19a_Playfield, LM_L_l19a_Under_PF, LM_L_l19a_sw54_001, LM_L_l19a_sw64_001, LM_L_l20_Layer_1, LM_L_l20_Parts, LM_L_l20_Playfield, LM_L_l20_Under_PF, LM_L_l20_sw43_001, LM_L_l20_sw53_001, LM_L_l20_sw63_001, LM_L_l21_Layer_1, LM_L_l21_Parts, _
  LM_L_l21_Playfield, LM_L_l21_Under_PF, LM_L_l21_sw43_001, LM_L_l21_sw53_001, LM_L_l21_sw63_001, LM_L_l22_Layer_1, LM_L_l22_Parts, LM_L_l22_Playfield, LM_L_l22_Under_PF, LM_L_l22_sw43_001, LM_L_l22_sw53_001, LM_L_l22_sw63_001, LM_L_l23_Parts, LM_L_l23_Playfield, LM_L_l23_Under_PF, LM_L_l23_sw40, LM_L_l23_sw50, LM_L_l23_sw73, LM_L_l23a_Bumper1_Ring, LM_L_l23a_Layer_1, LM_L_l23a_Parts, LM_L_l23a_Playfield, LM_L_l23a_Under_PF, LM_L_l23a_sw73a_001, LM_L_l24_Layer_1, LM_L_l24_Parts, LM_L_l24_Playfield, LM_L_l24_Under_PF, LM_L_l24_sw44_001, LM_L_l24_sw54_001, LM_L_l24_sw63_001, LM_L_l24_sw64_001, LM_L_l25_Parts, LM_L_l25_Playfield, LM_L_l25_Under_PF, LM_L_l25_sw44_001, LM_L_l25_sw54_001, LM_L_l25_sw63_001, LM_L_l25_sw64_001, LM_L_l26_Parts, LM_L_l26_Playfield, LM_L_l26_Under_PF, LM_L_l26_sw44_001, LM_L_l26_sw54_001, LM_L_l26_sw63_001, LM_L_l26_sw64_001, LM_L_l27_LFlip1_001, LM_L_l27_LFlip1U, LM_L_l27_Layer_1, LM_L_l27_Parts, LM_L_l27_Playfield, LM_L_l27_Rflip1_001, LM_L_l27_Rflip1U, LM_L_l27_Under_PF, _
  LM_L_l28_Playfield, LM_L_l28_Under_PF, LM_L_l29_Parts, LM_L_l29_Playfield, LM_L_l29_Under_PF, LM_L_l3_Parts, LM_L_l3_Playfield, LM_L_l3_Under_PF, LM_L_l30_Parts, LM_L_l30_Playfield, LM_L_l30_Under_PF, LM_L_l31_Playfield, LM_L_l31_Under_PF, LM_L_l31_sw40, LM_L_l31_sw50, LM_L_l31_sw60, LM_L_l36_Layer_1, LM_L_l36_Parts, LM_L_l36_Playfield, LM_L_l36_Under_PF, LM_L_l36_rsling1, LM_L_l36_rsling2, LM_L_l36_sw61, LM_L_l37_Layer_1, LM_L_l37_Lflip_001, LM_L_l37_Parts, LM_L_l37_Playfield, LM_L_l37_RFlip_001, LM_L_l37_Under_PF, LM_L_l37_lsling1, LM_L_l37_lsling2, LM_L_l37_sw43_001, LM_L_l37_sw53_001, LM_L_l38_Layer_1, LM_L_l38_Lflip_001, LM_L_l38_Parts, LM_L_l38_Playfield, LM_L_l38_Under_PF, LM_L_l38_sw43_001, LM_L_l38_sw53_001, LM_L_l38_sw63_001, LM_L_l39_Layer_1, LM_L_l39_Lflip_001, LM_L_l39_Parts, LM_L_l39_Playfield, LM_L_l39_Under_PF, LM_L_l39_rsling1, LM_L_l39_rsling2, LM_L_l39_sw43_001, LM_L_l39_sw53_001, LM_L_l39_sw61, LM_L_l39_sw63_001, LM_L_l40_Layer_1, LM_L_l40_Lflip_001, LM_L_l40_Parts, LM_L_l40_Playfield, _
  LM_L_l40_RFlip_001, LM_L_l40_Under_PF, LM_L_l40_sw43_001, LM_L_l41_Lflip_001, LM_L_l41_LflipU, LM_L_l41_Parts, LM_L_l41_Playfield, LM_L_l41_RFlip_001, LM_L_l41_RFlipU, LM_L_l41_Under_PF, LM_L_l42_Lflip_001, LM_L_l42_LflipU, LM_L_l42_Parts, LM_L_l42_Playfield, LM_L_l42_RFlip_001, LM_L_l42_RFlipU, LM_L_l42_Under_PF, LM_L_l43_Lflip_001, LM_L_l43_LflipU, LM_L_l43_Parts, LM_L_l43_Playfield, LM_L_l43_RFlip_001, LM_L_l43_RFlipU, LM_L_l43_Under_PF, LM_L_l44_Lflip_001, LM_L_l44_LflipU, LM_L_l44_Parts, LM_L_l44_Playfield, LM_L_l44_RFlip_001, LM_L_l44_RFlipU, LM_L_l44_Under_PF, LM_L_l45_Lflip_001, LM_L_l45_LflipU, LM_L_l45_Parts, LM_L_l45_Playfield, LM_L_l45_RFlip_001, LM_L_l45_RFlipU, LM_L_l45_Under_PF, LM_L_l46_Layer_2, LM_L_l46_Lflip_001, LM_L_l46_LflipU, LM_L_l46_Parts, LM_L_l46_Playfield, LM_L_l46_RFlip_001, LM_L_l46_RFlipU, LM_L_l46_Under_PF, LM_L_l47_Parts, LM_L_l47_Playfield, LM_L_l47_URsling1, LM_L_l47_URsling2, LM_L_l47_Under_PF, LM_L_l47_rsling1, LM_L_l47_sw41, LM_L_l47_sw51, LM_L_l47a_Parts, _
  LM_L_l47a_Playfield, LM_L_l47a_Under_PF, LM_L_l47a_sw40, LM_L_l47a_sw50, LM_L_l47a_sw60, LM_L_l48_Gate7_Wire, LM_L_l48_Parts, LM_L_l48_Playfield, LM_L_l48_Under_PF, LM_L_l48_sw42, LM_L_l49_Gate7_Wire, LM_L_l49_Layer_1, LM_L_l49_Parts, LM_L_l49_Playfield, LM_L_l49_Under_PF, LM_L_l49_sw52, LM_L_l50_Layer_1, LM_L_l50_Parts, LM_L_l50_Playfield, LM_L_l50_Under_PF, LM_L_l50_sw62, LM_L_l9_Layer_1, LM_L_l9_Parts, LM_L_l9_Playfield, LM_L_l9_Under_PF, LM_L_l9_sw43_001, LM_L_l9_sw53_001, LM_L_l9_sw63_001)
' VLM  Arrays - End

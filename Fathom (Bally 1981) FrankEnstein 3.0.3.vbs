' Fathom / IPD No. 829 / August, 1981 / 4 Players
' _____ ____  _____  _     ____  _
'/    //  _ \/__ __\/ \ /|/  _ \/ \__/|
'|  __\| / \|  / \  | |_||| / \|| |\/||
'| |   | |-||  | |  | | ||| \_/|| |  ||
'\_/   \_/ \|  \_/  \_/ \|\____/\_/  \|

' Version 3.0 by FrankEnstein 2024 based on:
' Version 2.0 by UnclePaulie 2023
' Table includes Hybrid VR/desktop/cabinet modes, enhanced playfield by Ebislit and Redbone, VPW physics, Fleep sounds, Lampz, 3D inserts, new GI, sling corrections, targets, saucers, playfield mesh, etc.

' Complete implementation details and credits found at the bottom of the script.

Option Explicit
Randomize


'*******************************************
' Desktop, Cab, and VR OPTIONS
'*******************************************

Dim TestDTVR : TestDTVR = false
'TestDTVR = True

' Desktop, Cab, and VR Room are automatically selected.  However if in VR Room mode, you can change the environment with F12 Options

' DIP Switch options
const SetDIPSwitches = 1  ' 0 is the default dip settings for game play. (RECOMMENDED) If you want to set the dips differently, set to 1, and then hit F6 to launch.


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
'Const VolumeDial = 0.8       ' Recommended values should be no greater than 1.
'Const BallRollVolume = 0.5       'Level of ball rolling volume. Value between 0 and 1
Const RampRollVolume = 0.5      'Level of ramp rolling volume. Value between 0 and 1

'----- Physics Mods -----
Const TargetBouncerEnabled = 1    '0 = normal standup targets, 1 = bouncy targets
Const TargetBouncerFactor = 0.7   'Level of bounces. Recommmended value of 0.7 when TargetBouncerEnabled


' ****************************************************
' standard definitions
' ****************************************************

Dim VR_Room, cab_mode, DesktopMode: DesktopMode = Table1.ShowDT

Const UseSolenoids  = 2
Const UseLamps    = 1
Const UseSync     = 0
Const HandleMech  = 0
Const UseGI     = 0
Const cGameName   = "fathom"    'ROM name
Const ballsize    = 50
Const ballmass    = 1

'***********************

Const tnob = 3            'Total number of balls  (2 multiball, and 5 captive balls)
Const lob = 0           'Locked balls

Dim tablewidth: tablewidth = Table1.width
Dim tableheight: tableheight = Table1.height

Dim i, FBall1, Fball2, Fball3, gBOT
Dim BIPL : BIPL = False       'Ball in plunger lane

Dim gion: gion = 0

Dim DT27up, DT28up, DT29up, DT30up, DT31up, DT32up, DT33up, DT34up, DT35up, DT42up, DT43up, DT44up, DT46up, DT47up, DT48up

Dim EnviroSound : EnviroSound = "underwater_ambience"

On Error Resume Next
ExecuteGlobal GetTextFile("controller.vbs")
If Err Then MsgBox "You need the controller.vbs in order to run this table, available in the vp10 package"
On Error Goto 0

LoadVPM "01130100", "Bally.VBS", 3.21


'**********************************************************************************************************
'Solenoids
'**********************************************************************************************************

SolCallback(1)  = "SolDTTop"        'Top 3 Bank Reset
'SolCallback(1a) = "SolDTRight1"      '3rd Green/Right Inline Drop Target (switched with lamp 47)
SolCallback(2)  = "SolDTLeft"       'Left 6 Bank Reset
'SolCallback(2a) = "SolDTRight2"      '2nd Green/Right Inline Drop Target (switched with lamp 47)
SolCallback(3)  = "SolDTMid"        'Middle 3 Bank Reset
'SolCallback(3a) = "SolDTRight3"        '1st Green/Right Inline Drop Target (switched with lamp 47)
SolCallback(4)  = "SolDTRight"        'Right 3 Bank Reset
SolCallback(6)  = "SolKnocker"        'Knocker solenoid
SolCallback(7)  = "SolOutholeKicker"    'Trough to plunger lane
SolCallback(13) = "SolSaucerTop"      'Top Saucer Kick
'SolCallback(13a) = "SolDTTop1"       '3rd Blue/Top Inline Drop Target (switched with lamp 47)
SolCallback(14) = "SolSaucerRight"      'Right Saucer Kick
'SolCallback(14a) = "SolDTTop2"       '2nd Blue Inline Drop Target (switched with lamp 47)
SolCallback(15) = "SolDTTop3"       '1st Blue Inline Drop Target

SolCallback(18) = "BGGI"          'Backglass GI
SolCallback(19) = "PFGI"          'Playfield GI

SolCallback(sllflipper)="SolLFlipper"
SolCallback(slrflipper)="SolRFlipper"


'*******************************************
'  ZOPT: User Options
'*******************************************

Dim LightLevel : LightLevel = 1         ' Level of room lighting (0 to 1), where 0 is dark and 100 is brightest
Dim ColorLUT : ColorLUT = 1           ' Color desaturation LUTs: 1 to 11, where 1 is normal and 11 is black'n'white
Dim VolumeDial : VolumeDial = 0.8             ' Overall Mechanical sound effect volume. Recommended values should be no greater than 1.
Dim BallRollVolume : BallRollVolume = 0.5     ' Level of ball rolling volume. Value between 0 and 1
Dim OceanEnv : OceanEnv = 1           ' 0 = normal room VR environment, 1 = Ocean Environment (default). Set via F12 Options
Dim EnviroSoundsOn              ' Environment Sounds (o - off, 1 - On)
Dim EnviroSoundsLast: EnviroSoundsLast = 0
Dim WallClock : WallClock = 1   '1 Shows the clock in the VR minimal rooms only
Dim CustomWalls : CustomWalls = 0 'set to 0 for Modern Minimal Walls, floor, and roof, 1 for Sixtoe's original walls and floor
Dim topper : topper = 1    '0 = Off 1= On - Topper visible in VR Room only
Dim poster : poster = 1    '1 Shows the flyer posters in the VR room only
Dim poster2 : poster2 = 1    '1 Shows the flyer posters in the VR room only

' Called when options are tweaked by the player.
' - 0: game has started, good time to load options and adjust accordingly
' - 1: an option has changed
' - 2: options have been reseted
' - 3: player closed the tweak UI, good time to update staticly prerendered parts
' Table1.Option arguments are:
' - option name, minimum value, maximum value, step between valid values, default value, unit (0=None, 1=Percent), an optional arry of literal strings
Dim dspTriggered : dspTriggered = False


Function ReflectionToggle(state)
  Dim BP
  For Each BP in Reflection_Array
    BP.ReflectionEnabled = state
  Next
  For Each BP in BG_Bakemap
    BP.ReflectionEnabled = state
  Next
  BM_UnderPF.ReflectionEnabled = False
  BM_Playfield.ReflectionEnabled = False
  BM_Playfield.ReflectionProbe = "Playfield Reflections"
End Function

Sub Table1_OptionEvent(ByVal eventId)
    If eventId = 1 And Not dspTriggered Then dspTriggered = True : DisableStaticPreRendering = True : End If

  Dim v, BP

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

    ' Toggle Rails
  If RenderingMode <> 2 Then
    v = Table1.Option("Show Rails (Desktop/Cab)", 0, 1, 1, 1, 0, Array("Hide", "Show"))
    For Each BP in BP_Ramp001 : BP.Visible = v: Next
    For Each BP in BP_Ramp002 : BP.Visible = v: Next
  End If

    ' Toggle Lockdown Bar
  If RenderingMode <> 2 Then
    v = Table1.Option("Show Lockbar (Desktop/Cab)", 0, 1, 1, 1, 0, Array("Hide", "Show"))
    For Each BP in BP_lockdownbar : BP.Visible = v: Next
  End If

    ' Toggle Reflections
  v = Table1.Option("Playfield Reflections", 0, 1, 1, 1, 0, Array("Hide", "Show"))
  ReflectionToggle(v)

    ' Toggle Refraction
  v = Table1.Option("Plastics Refractions", 0, 1, 1, 1, 0, Array("Off", "On"))
  Select Case v
    Case 0 :
      BM_Overlay.RefractionProbe = ""
      BM_PartsST.RefractionProbe = ""
    Case 1 :
      BM_Overlay.RefractionProbe = "RefractionProbe"
      BM_PartsST.RefractionProbe = "RefractionProbe"
      For Each BP in BP_Overlay : BP.z = 55: Next
  End Select

    ' Enviro Sound
  EnviroSoundsOn = Table1.Option("Environment Sounds", 0, 1, 1, 0, 0, Array("OFF", "ON"))
  if EnviroSoundsOn = 1 and EnviroSoundsLast = 0 Then
    PlaySound EnviroSound, -1, 1.0, 0.0, 0.0, 0, 0, 0, 0.0
    PlaySound "intro", 1, 1.0, 0.0, 0.0, 0, 0, 0, 0.0
  ElseIf EnviroSoundsOn = 0 Then
    StopSound EnviroSound
    StopSound "intro"
  end If
  EnviroSoundsLast = EnviroSoundsOn

    ' VR Room
  OceanEnv = Table1.Option("VR Room", 0, 1, 1, 1, 0, Array("Minimal", "Underwater"))

  If RenderingMode = 2 or TestDTVR=True Then VR_Room=1 Else VR_Room=0      'VRRoom set based on RenderingMode in version 10.72
  if Not DesktopMode and VR_Room=0 Then cab_mode=1 Else cab_mode=0
  SetupRoom

  If RenderingMode = 2 or TestDTVR = True Then
    ' VR Minimal Room Options
    WallClock = Table1.Option("Clock", 0, 1, 1, 1, 0, Array("OFF", "ON"))
    If OceanEnv = 0 Then
      If WallClock = 0 Then
        for each v in VRClock:v.visible = 0: Next
      Else
        for each v in VRClock:v.visible = 1: Next
      End If
    End if
    CustomWalls = Table1.Option("Custom Walls", 0, 1, 1, 0, 0, Array("OFF", "ON"))
    If OceanEnv = 0 Then
    'Custom Walls, Floor, and Roof
      if CustomWalls = 1 Then
        VR_Wall_Left.image = "VR_Wall_Left"
        VR_Wall_Right.image = "VR_Wall_Right"
        VR_Floor.image = "VR_Floor"
        VR_Roof.image = "VR_Roof"
      Else
        VR_Wall_Left.image = "wallpaper left"
        VR_Wall_Right.image = "wallpaper_right"
        VR_Floor.image = "FloorCompleteMap2"
        VR_Roof.image = "light grey"
      end if
    End if
    'Topper
    topper = Table1.Option("Topper", 0, 1, 1, 1, 0, Array("OFF", "ON"))
    If OceanEnv = 0 Then
      If topper = 1 Then
        Primary_topper.visible = 1
      Else
        Primary_topper.visible = 0
      End If
    End if
    'Poster
    poster = Table1.Option("Poster", 0, 1, 1, 1, 0, Array("OFF", "ON"))
    If OceanEnv = 0 Then
      If poster = 1 Then
        VRposter.visible = 1
      Else
        VRposter.visible = 0
      End If
    End if
    'Poster 2
    poster2 = Table1.Option("Poster 2", 0, 1, 1, 1, 0, Array("OFF", "ON"))
    If OceanEnv = 0 Then
      If poster2 = 1 Then
        VRposter2.visible = 1
      Else
        VRposter2.visible = 0
      End If
    End if
  End if
  ' Room brightness
  ' LightLevel = Table1.Option("Table Brightness (Ambient Light Level)", 0, 1, 0.01, .5, 1)
  LightLevel = NightDay/100
  SetRoomBrightness LightLevel

    If eventId = 3 And dspTriggered Then dspTriggered = False : DisableStaticPreRendering = False : End If
End Sub


'******************************************************
'   ZBBR: BALL BRIGHTNESS
'******************************************************

Const BallBrightness =  1        'Ball brightness - Value between 0 and 1 (0=Dark ... 1=Bright)
Const PLOffset = 0.5
Dim PLGain: PLGain = (1-PLOffset)/(1260-2000)
Dim gilvl:gilvl = 0

Sub UpdateBallBrightness
  Dim s, b_base, b_r, b_g, b_b, d_w
  b_base = 120 * BallBrightness + 135*gilvl ' orig was 120 and 70
  Dim gBOT: gBOT = GetBalls

  For s = 0 To UBound(gBOT)
    ' Handle z direction
    d_w = b_base*(1 - (gBOT(s).z-25)/500)
    If d_w < 30 Then d_w = 30
    ' Handle plunger lane
    If InRect(gBOT(s).x,gBOT(s).y,870,2000,870,1260,930,1260,930,2000) Then
      d_w = d_w*(PLOffset+PLGain*(gBOT(s).y-2000))
    End If
    ' Handle cave trap scoop
    If InRect(gBOT(s).x,gBOT(s).y,750,650,750,400,860,400,860,650) Then
      d_w = d_w*(PLOffset+PLGain*(gBOT(s).y-5000))
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
'    Room Brightness
'****************************

' Update these arrays if you want to change more materials with room light level
Dim RoomBrightnessMtlArray: RoomBrightnessMtlArray = Array("VLM.Bake.Active","VLM.Bake.Solid")

Sub SetRoomBrightness(lvl)
  'debug.print lvl
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

'*****************vpmMapLights
'   GI
'*****************

Sub PFGI(Enabled)
    If Enabled Then
    For each xx in GI:xx.state = 1: Next

    gion = 1

  ' Set Drop Target Shadows to state when GI is back on.
    if DT27up=1 Then dtsh27.visible = 1
    if DT28up=1 Then dtsh28.visible = 1
    if DT29up=1 Then dtsh29.visible = 1
    if DT30up=1 Then dtsh30.visible = 1
    if DT31up=1 Then dtsh31.visible = 1
    if DT32up=1 Then dtsh32.visible = 1

  Else
    gion = 0
    For each xx in GI:xx.state = 0: Next

' Set Drop Target Shadows to off when GI is goes off.
    for each xx in ShadowDT
      xx.visible=False
    Next

    End If
End Sub


Sub BGGI(Enabled)
    If Enabled Then
    BGGIUpdates 1
    Else
    BGGIUpdates 0
    End If
End Sub


'*******************************************
'  Timers
'*******************************************

' The game timer interval is 10 ms
GameTimer.Interval = 10
GameTimer.Enabled = True
Sub GameTimer_Timer()
  Cor.Update            'update ball tracking
End Sub


' The frame timer interval is -1, so executes at the display frame rate
FrameTimer.Interval = -1
FrameTimer.Enabled = False 'start in table1_init
Sub FrameTimer_Timer()
  If DynamicBallShadowsOn Or AmbientBallShadowOn Then DynamicBSUpdate 'update ball shadows

  If VR_Room = 0 Then DisplayTimer
  If VR_Room = 1 Then VRDisplayTimer
  if VR_Room = 1 Then UpdateVRLamps
  if VR_Room = 0 and cab_mode = 0 Then UpdateDTLamps

  UpdateBallBrightness      'GI for the ball

  RollingUpdate         'update rolling sounds
  DoDTAnim            'handle drop target animations
  UpdateDropTargets       'update drop target BMs
  DoSTAnim            'handle stand up target animations
  UpdateStandupTargets

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
    .SplashInfoLine = "Fathom (Bally 1981)"&chr(13)&"by UnclePaulie"
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

  vpmMapLights AllLamps ' Map all lamps to the corresponding ROM output using the value of TimerInterval of each light object

  vpmNudge.TiltSwitch = 15
  vpmNudge.Sensitivity = 5
  vpmNudge.TiltObj = Array(LeftSlingshot,RightSlingshot, LeftFlipper, RightFlipper, RightFlipper1, Bumper1, Bumper2, Bumper3)

'Ball initializations need for physical trough and ball shadows
  Set FBall1 = BallRelease.CreateSizedballWithMass(Ballsize/2,Ballmass)
  Set FBall2 = Slot1.CreateSizedballWithMass(Ballsize/2,Ballmass)
  Set FBall3 = Slot2.CreateSizedballWithMass(Ballsize/2,Ballmass)

'Ball initializations
  gBOT = Array(FBall1, Fball2, Fball3)
  Controller.Switch(1) = 1
  Controller.Switch(2) = 1
  Controller.Switch(3) = 1

' Add balls to shadow dictionary
  For Each xx in gBOT
    bsDict.Add xx.ID, bsNone
  Next

'Saucer and Right Flipper initializations
  controller.switch(4) = 0
  Controller.Switch(5) = 0


' Set up VR Backglass object positions
  InitDigits
  setup_backglass
  SetBackglass


' Make drop target shadows invisible
  Dim xx
  for each xx in ShadowDT
    xx.visible=False
  Next

' Drop Target Variable state
  DT27up=1
  DT28up=1
  DT29up=1
  DT30up=1
  DT31up=1
  DT32up=1
  DT33up=1
  DT34up=1
  DT35up=1
  DT42up=1
  DT43up=1
  DT44up=1
  DT46up=1
  DT47up=1
  DT48up=1


' If user wants to set dip switches themselves it will force them to set it via F6.
  If SetDIPSwitches = 0 Then
    SetDefaultDips
  End If

' Moved this to option menu
' if EnviroSoundsOn = 1 Then
'   PlaySound EnviroSound, -1, 1.0, 0.0, 0.0, 0, 0, 0, 0.0
'   PlaySound "intro", 1, 1.0, 0.0, 0.0, 0, 0, 0, 0.0
' end If

  FrameTimer.Enabled = True

End Sub

Sub table1_Paused:Controller.Pause = 1:End Sub
Sub table1_unPaused:Controller.Pause = 0:End Sub
Sub table1_exit
  StopSound EnviroSound
' SaveLUT
  Controller.stop
End Sub

'********************************************
' Keys and Plunger code
'********************************************

Sub table1_KeyDown(ByVal Keycode)

  If keycode = LeftTiltKey Then Nudge 90, 0.5 ': SoundNudgeLeft
  If keycode = RightTiltKey Then Nudge 270, 0.5 ': SoundNudgeRight
  If keycode = CenterTiltKey Then Nudge 0, 0.5 ': SoundNudgeCenter


  If KeyCode = PlungerKey Then
    Plunger.PullBack
    SoundPlungerPull()
    TimerVRPlunger.Enabled = True
    TimerVRPlunger1.Enabled = False
    PinCab_Shooter.Y = 2121.572
  End If

  If keycode = LeftFlipperKey Then
    VRFlipperButtonLeft.X = VRFlipperButtonLeft.X + 8
    FlipperActivate LeftFlipper, LFPress
  End If

  If keycode = RightFlipperKey Then
    VRFlipperButtonRight.X = VRFlipperButtonRight.X - 8
    FlipperActivate RightFlipper, RFPress
    Controller.Switch(7) = 1
  End If

  If keycode = StartGameKey Then
    StartButton.y = 1943.978 - 5
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


dim plungerspeed


Sub table1_KeyUp(ByVal Keycode)

  If keycode = PlungerKey Then
    plunger.firespeed = RndInt(95,110)  ' Generate random value variance of +/- 1 from 110
    Plunger.Fire
    If BIPL = 1 Then
      SoundPlungerReleaseBall()     'Plunger release sound when there is a ball in shooter lane
    Else
      SoundPlungerReleaseNoBall()     'Plunger release sound when there is no ball in shooter lane
    End If
    TimerVRPlunger.Enabled = False
    TimerVRPlunger1.Enabled = True
    PinCab_Shooter.Y = 2121.572
  End If

  If keycode = LeftFlipperKey Then
    VRFlipperButtonLeft.X = VRFlipperButtonLeft.X - 8
    FlipperDeActivate LeftFlipper, LFPress
  End If

  If keycode = RightFlipperKey Then
    VRFlipperButtonRight.X = VRFlipperButtonRight.X + 8
    FlipperDeActivate RightFlipper, RFPress
    Controller.Switch(7) = 0
  End If

  If keycode = StartGameKey Then
    StartButton.y = 1943.978
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


'******************************************************
'     TROUGH BASED ON FOZZY and Rothbauerw
'******************************************************

Sub BallRelease_Hit : Controller.Switch(1) = 1 : UpdateTrough : End Sub
Sub BallRelease_UnHit : Controller.Switch(1) = 0 : UpdateTrough : End Sub
Sub Slot1_Hit : Controller.Switch(2) = 1 : UpdateTrough : End Sub
Sub Slot1_UnHit : Controller.Switch(2) = 0 : UpdateTrough : End Sub
Sub Slot2_Hit : Controller.Switch(3) = 1 : UpdateTrough : End Sub
Sub Slot2_UnHit : Controller.Switch(3) = 0 : UpdateTrough : End Sub


Sub UpdateTrough
  UpdateTroughTimer.Interval = 300
  UpdateTroughTimer.Enabled = 1
End Sub

Sub UpdateTroughTimer_Timer
  If BallRelease.BallCntOver = 0 Then Slot1.kick 60, 8
  If Slot1.BallCntOver = 0 Then Slot2.kick 60, 8
  If Slot2.BallCntOver = 0 Then Drain.kick 60, 12
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

Sub SolSaucerTop(Enabled)
  If Enabled Then
    If Controller.Lamp(47) = 0 Then
      RNDKickAngle1 = RndInt(315,317)   ' Generate random value variance of +/- 1 from 316
      RNDKickValue1 = 40
      sw4.kick RNDKickAngle1, RNDKickValue1
      sw4Step = 0
      sw4.timerenabled = 1
    Else
      SolDTTop1(1)
    End If
  End If
End Sub

Dim SW4Step
Sub SW4_Timer()
  Dim BP
  Select Case SW4Step
    Case 0: pKickerArm.Rotx = 4 : For Each BP in BP_Kicker2Arm : BP.RotX = 4 : Next
    Case 2: pKickerArm.Rotx = 8 : For Each BP in BP_Kicker2Arm : BP.RotX = 8 : Next
    Case 3: pKickerArm.Rotx = 8 : For Each BP in BP_Kicker2Arm : BP.RotX = 8 : Next
    Case 4: pKickerArm.Rotx = 4 : For Each BP in BP_Kicker2Arm : BP.RotX = 4 : Next
    Case 5: pKickerArm.Rotx = 0 : sw4.timerenabled = 0 : sw4Step = -1 : For Each BP in BP_Kicker2Arm : BP.RotX = 0 : Next
  End Select
  sw4Step = sw4Step + 1
End Sub

Sub sw4_Unhit
  SoundSaucerKick 1, sw4
  controller.switch(4) = 0
End Sub

Sub sw5_Hit
  SoundSaucerLock
  Controller.Switch(5) = 1
End Sub

Sub SolSaucerRight(Enabled)
  If Enabled Then
    If Controller.Lamp(47) = 0 Then
      RNDKickAngle1 = RndInt(178,182)   ' Generate random value variance of +/- 2 from 180
      RNDKickValue1 = RndInt(26,28)     ' Generate random value variance of +/- 1 from 27
      sw5.kick RNDKickAngle1, RNDKickValue1
      sw5Step = 0
      sw5.timerenabled = 1
    Else
      SolDTTop2(1)
    End If
  End If
End Sub

Dim SW5Step
Sub SW5_Timer()
  Select Case SW5Step
    Case 0: pKickerArm2.Rotx = 4 : For Each BP in BP_Kicker1Arm : BP.RotX = 4 : Next
    Case 2: pKickerArm2.Rotx = 8 : For Each BP in BP_Kicker1Arm : BP.RotX = 8 : Next
    Case 3: pKickerArm2.Rotx = 8 : For Each BP in BP_Kicker1Arm : BP.RotX = 8 : Next
    Case 4: pKickerArm2.Rotx = 4 : For Each BP in BP_Kicker1Arm : BP.RotX = 4 : Next
    Case 5: pKickerArm2.Rotx = 0:sw5.timerenabled = 0:sw5Step = -1: : For Each BP in BP_Kicker1Arm : BP.RotX = 0 : Next
  End Select
  sw5Step = sw5Step + 1
End Sub

Sub sw5_Unhit
  SoundSaucerKick 1, sw5
  controller.switch(5) = 0
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
Sub sw21_Hit   : Controller.Switch(21) = 1 : End Sub
Sub sw21_Unhit : Controller.Switch(21) = 0 : End Sub
Sub sw22_Hit   : Controller.Switch(22) = 1 : End Sub
Sub sw22_Unhit : Controller.Switch(22) = 0 : End Sub
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

Sub sw201_Hit:vpmTimer.PulseSw 20:me.timerenabled = 0:AnimateStar star201, sw201, 1:End Sub
Sub sw201_UnHit:me.timerinterval = 7:me.timerenabled = 1:End Sub
Sub sw201_timer:AnimateStar star201, sw201, 0:End Sub

Sub sw202_Hit:vpmTimer.PulseSw 20:me.timerenabled = 0:AnimateStar star202, sw202, 1:End Sub
Sub sw202_UnHit:me.timerinterval = 7:me.timerenabled = 1:End Sub
Sub sw202_timer:AnimateStar star202, sw202, 0:End Sub

Sub sw203_Hit:vpmTimer.PulseSw 20:me.timerenabled = 0:AnimateStar star203, sw203, 1:End Sub
Sub sw203_UnHit:me.timerinterval = 7:me.timerenabled = 1:End Sub
Sub sw203_timer:AnimateStar star203, sw203, 0:End Sub

Sub sw25_Hit:vpmTimer.PulseSw 25:me.timerenabled = 0:AnimateStar star25, sw25, 1:End Sub
Sub sw25_UnHit:me.timerinterval = 7:me.timerenabled = 1:End Sub
Sub sw25_timer:AnimateStar star25, sw25, 0:End Sub

Sub sw26_Hit:vpmTimer.PulseSw 26:me.timerenabled = 0:AnimateStar star26, sw26, 1:End Sub
Sub sw26_UnHit:me.timerinterval = 7:me.timerenabled = 1:End Sub
Sub sw26_timer:AnimateStar star26, sw26, 0:End Sub


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

Sub Spinner_Spin()
  vpmtimer.PulseSw 18
  SoundSpinner Spinner
End Sub


'*******************************************
' Bumpers
'*******************************************

Sub Bumper1_Hit 'Left Top
  RandomSoundBumperTop Bumper1
  vpmTimer.PulseSw 40
  'VLM movable script
  Dim BP
  For each BP in BP_Bumper1_Socket
    BP.roty = skirtAY(me,Activeball)
    BP.rotx = skirtAX(me,Activeball)
  Next
  me.TimerEnabled = 1
End Sub

Sub Bumper1_Timer
  me.TimerEnabled = 0
  'VLM movable script
  Dim BP
  For each BP in BP_Bumper1_Socket
    BP.roty = 0
    BP.rotx = 0
  Next
End Sub

Sub Bumper2_Hit 'Right Top
  RandomSoundBumperTop Bumper2
  vpmTimer.PulseSw 38
  'VLM movable script
  Dim BP
  For each BP in BP_Bumper2_Socket
    BP.roty = skirtAY(me,Activeball)
    BP.rotx = skirtAX(me,Activeball)
  Next
  me.TimerEnabled = 1
End Sub

Sub Bumper2_Timer
  me.TimerEnabled = 0
  'VLM movable script
  Dim BP
  For each BP in BP_Bumper2_Socket
    BP.roty = 0
    BP.rotx = 0
  Next
End Sub

Sub Bumper3_Hit 'Bottom
  RandomSoundBumperBottom Bumper3
  vpmTimer.PulseSw 39
  'VLM movable script
  Dim BP
  For each BP in BP_Bumper3_Socket
    BP.roty = skirtAY(me,Activeball)
    BP.rotx = skirtAX(me,Activeball)
  Next
  me.TimerEnabled = 1
End Sub

Sub Bumper3_Timer
  me.TimerEnabled = 0
  'VLM movable script
  Dim BP
  For each BP in BP_Bumper3_Socket
    BP.roty = 0
    BP.rotx = 0
  Next
End Sub

'**
'            SKIRT ANIMATION FUNCTIONS
'**
' NOTE: set bumper object timer to around 150-175 in order to be able
' to actually see the animation, adjust to your liking

'Const PI = 3.1415926
Const SkirtTilt = 4.5       'angle of skirt tilting in degrees

Function SkirtAX(bumper, bumperball)
    skirtAX=cos(skirtA(bumper,bumperball))*(SkirtTilt)        'x component of angle
    if (bumper.y<bumperball.y) then    skirtAX=-skirtAX    'adjust for ball hit bottom half
End Function

Function SkirtAY(bumper, bumperball)
    skirtAY=sin(skirtA(bumper,bumperball))*(SkirtTilt)        'y component of angle
    if (bumper.x>bumperball.x) then    skirtAY=-skirtAY    'adjust for ball hit left half
End Function

Function SkirtA(bumper, bumperball)
    dim hitx, hity, dx, dy
    hitx=bumperball.x
    hity=bumperball.y

    dy=Abs(hity-bumper.y)                    'y offset ball at hit to center of bumper
    if dy=0 then dy=0.0000001
    dx=Abs(hitx-bumper.x)                    'x offset ball at hit to center of bumper
    skirtA=(atn(dx/dy)) '/(PI/180)            'angle in radians to ball from center of Bumper1
End Function


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


'********************************************
' Drop Target Hits
'********************************************

Sub sw27_Hit
  DTHit 27
  DT27up=0
End Sub

Sub sw28_Hit
  DTHit 28
  DT28up=0
End Sub

Sub sw29_Hit
  DTHit 29
  DT29up=0
End Sub

Sub sw30_Hit
  DTHit 30
  DT30up=0
End Sub

Sub sw31_Hit
  DTHit 31
  DT31up=0
End Sub

Sub sw32_Hit
  DTHit 32
  DT32up=0
End Sub

Sub sw33_Hit
  DTHit 33
  DT33up=0
End Sub

Sub sw34_Hit
  DTHit 34
  DT34up=0
End Sub

Sub sw35_Hit
  DTHit 35
  DT35up=0
End Sub

Sub sw42_Hit
  DTHit 42
  DT42up=0
End Sub

Sub sw43_Hit
  DTHit 43
  DT43up=0
End Sub

Sub sw44_Hit
  DTHit 44
  DT44up=0
End Sub

Sub sw46_Hit
  DTHit 46
  DT46up=0
End Sub

Sub sw47_Hit
  DTHit 47
  DT47up=0
End Sub

Sub sw48_Hit
  DTHit 48
  DT48up=0
End Sub


'********************************************
' Drop Target Solenoid Controls
'********************************************

Sub SolDTTop (enabled)
  If enabled then
    If Controller.Lamp(47) = 0 Then
      RandomSoundDropTargetReset sw43p
      DTRaise 42
      DTRaise 43
      DTRaise 44
      DT42up=1
      DT43up=1
      DT44up=1
    Else
      SolDTRight1(1)
    End If
  End if
End Sub

Sub SolDTLeft (enabled)
  If enabled then
    If Controller.Lamp(47) = 0 Then
      RandomSoundDropTargetReset sw29p
      DTRaise 27
      DTRaise 28
      DTRaise 29
      DTRaise 30
      DTRaise 31
      DTRaise 32
      DT27up=1
      DT28up=1
      DT29up=1
      DT30up=1
      DT31up=1
      DT32up=1
      dtsh27.visible=True
      dtsh28.visible=True
      dtsh29.visible=True
      dtsh30.visible=True
      dtsh31.visible=True
      dtsh32.visible=True
    Else
      SolDTRight2(1)
    End If
  End if
End Sub

Sub SolDTMid (enabled)
  If enabled then
    If Controller.Lamp(47) = 0 Then
      RandomSoundDropTargetReset sw34p
      DTRaise 33
      DTRaise 34
      DTRaise 35
      DT33up=1
      DT34up=1
      DT35up=1
    Else
      SolDTRight3(1)
    End If
  End if
End Sub

Sub SolDTRight (enabled)
  If enabled then
    RandomSoundDropTargetReset sw47p
    DTRaise 46
    DTRaise 47
    DTRaise 48
    DT46up=1
    DT47up=1
    DT48up=1
  End if
End Sub

Sub SolDTTop1 (enabled)
  If enabled then
    if DT44up = 1 Then DTDrop 44
    DT44up=0
  End if
End Sub

Sub SolDTTop2 (enabled)
  If enabled then
    if DT43up = 1 Then  DTDrop 43
    DT43up=0
  End if
End Sub

Sub SolDTTop3 (enabled)
  If enabled then
    if DT42up = 1 Then DTDrop 42
    DT42up=0
  End if
End Sub

Sub SolDTRight1 (enabled)
  If enabled then
    if DT48up = 1 Then DTDrop 48
    DT48up=0
  End if
End Sub

Sub SolDTRight2 (enabled)
  If enabled then
    if DT47up = 1 Then DTDrop 47
    DT47up=0
  End if
End Sub

Sub SolDTRight3 (enabled)
  If enabled then
    if DT46up = 1 Then DTDrop 46
    DT46up=0
  End if
End Sub


'*******************************************
' Leaf Standup Sensors
'*******************************************

Sub RubberBand006_Hit: vpmTimer.PulseSw 19:End Sub
Sub RubberBand010_Hit: vpmTimer.PulseSw 19:End Sub


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
Dim RStep, Lstep


Sub RightSlingShot_Slingshot
  Dim BP
  RS.VelocityCorrect(ActiveBall)
  vpmTimer.PulseSw 36
  RandomSoundSlingshotRight SLING1
  For Each BP in BP_Sling1 : BP.rotx = 12: Next
  For Each BP in BP_RSling  : BP.Visible = 0: Next
  For Each BP in BP_RSling1 : BP.Visible = 1: Next
    RStep = 0
    RightSlingShot.TimerEnabled = 1
End Sub

Sub RightSlingShot_Timer
  Dim BP
    Select Case RStep
        Case 3
      For Each BP in BP_Sling1 : BP.rotx = 10: Next
      For Each BP in BP_RSling1 : BP.Visible = 0: Next
      For Each BP in BP_RSling2 : BP.Visible = 1: Next
        Case 4
      For Each BP in BP_Sling1 : BP.rotx = 0: Next
      For Each BP in BP_RSling2 : BP.Visible = 0: Next
      For Each BP in BP_RSling  : BP.Visible = 1: Next
      RightSlingShot.TimerEnabled = 0
    End Select
    RStep = RStep + 1
End Sub

Sub LeftSlingShot_Slingshot
  'Executes when Slingshot has been hit
  Dim BP
  LS.VelocityCorrect(ActiveBall)
  vpmTimer.PulseSw 37
  RandomSoundSlingshotLeft SLING2
  For Each BP in BP_Sling2 : BP.rotx = 12: Next
  For Each BP in BP_LSling  : BP.Visible = 0: Next
  For Each BP in BP_LSling1 : BP.Visible = 1: Next
    LStep = 0
    LeftSlingShot.TimerEnabled = 1
End Sub

Sub LeftSlingShot_Timer
  'Executes after hit and handles Slingshot recoil and reset to normal state
  Dim BP
    Select Case LStep
        Case 3
      For Each BP in BP_Sling2 : BP.rotx = 10: Next
      For Each BP in BP_LSling1 : BP.Visible = 0: Next
      For Each BP in BP_LSling2 : BP.Visible = 1: Next
        Case 4
      For Each BP in BP_Sling2 : BP.rotx = 0: Next
      For Each BP in BP_LSling2 : BP.Visible = 0: Next
      For Each BP in BP_LSling  : BP.Visible = 1: Next
      LeftSlingShot.TimerEnabled = 0
    End Select
    LStep = LStep + 1
End Sub

'*************************************
' Bally Fathom DIP switches
'*************************************
'Code done by Inkochnito
Sub editDips
  if SetDIPSwitches = 1 Then
  Dim vpmDips : Set vpmDips = New cvpmDips
  With vpmDips
    .AddForm 700,400,"Fathom - DIP switches"
    .AddChk 7,10,120,Array("Match feature",&H08000000)'dip 28
    .AddChk 130,10,120,Array("Game over attract",&H20000000)'dip 30
    .AddChk 260,10,120,Array("Credits display",&H04000000)'dip 27

    .AddFrame 2,30,190,"Maximum credits",&H03000000,Array("10 credits",0,"15 credits",&H01000000,"25 credits",&H02000000,"40 credits",&H03000000)'dip 25&26

    .AddFrame 2,106,190,"Locked ball adjustment",&H00000020,Array("eject ball at game end",0,"ball remains locked",&H00000020)'dip 6
    .AddFrame 2,152,190,"Bonus special lit at maximum",&H00000040,Array("blue and green bonus",0,"blue or green bonus",&H00000040)'dip 7
    .AddFrame 2,198,190,"Extra ball lite is lit for",&H00000080,Array("6 seconds",0,"10 seconds",&H00000080)'dip 8
    .AddFrame 2,244,190,"A-B-C special adjust",32768,Array("replay",0,"alternating points or replay",32768)'dip 16
    .AddFrame 205,30,190,"Balls per game",&HC0000000,Array("2 balls",&HC0000000,"3 balls",0,"4 balls",&H80000000,"5 balls",&H40000000)'dip 31&32
    .AddFrame 205,106,190,"Number of replays per game",&H10000000,Array("Only 1 replay per player per game",0,"All replays earned will be collected",&H10000000)'dip 29
    .AddFrame 205,152,190,"Drop target memory",&H00200000,Array("not in memory",0,"kept in memory",&H00200000)'dip 22
    .AddFrame 205,198,190,"55,000 bonus lites",&H00400000,Array("not in memory",0,"kept in memory",&H00400000)'dip 23
    .AddFrame 205,244,190,"A-B-C lites",&H00800000,Array("not in memory",0,"kept in memory",&H00800000)'dip 24
    .AddLabel 50,300,350,20,"Set selftest position 16,17,18 and 19 to 03 for the best gameplay."
    .AddLabel 50,320,300,20,"After hitting OK, press F3 to reset game with new settings."
    .ViewDips
  End With
  End If
End Sub
Set vpmShowDips = GetRef("editDips")


Sub SetDefaultDips

' Number of Balls / game
  SetDip &H40000000,0   'Number of balls.  0,0 = 3, 0,1 = 3, 1,0 = 5, 1,1 = 5
  SetDip &H80000000,0

' All other DIPs
  SetDip &H08000000,1   '1=Match feature
  SetDip &H20000000,1   '1=Game Over Attract
  SetDip &H04000000,1   '1=Credits display
  SetDip &H00000020,0   '0=eject ball at game end
  SetDip &H00000040,1   'Memory for Bonus and Specials; 0=blue AND green; 1= blue OR green
  SetDip &H00000080,1   'Memory for extra ball lite; 0=6 seconds, 1=10 seconds
  SetDip 32768,1      'Memory for ABC Special adjust; 0=replay, 1= alternating points or replay
  SetDip &H10000000,1   'Number of replays per game; 0=1 per game, 1= all replays collected
  SetDip &H00200000,1   'Memory for drop targets; 0=not in memory, 1=in memory
  SetDip &H00400000,1   'Memory for 55,000 bonus lites; 0=not in memory; 1=in memory
  SetDip &H00800000,1   'Memory for ABC lites; 0=not in memory; 1=in memory

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
dim objrtx1(3), objrtx2(3)
dim objBallShadow(3)
Dim OnPF(3)
Dim BallShadowA
BallShadowA = Array (BallShadowA0, BallShadowA1, BallShadowA2)
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
  'Dim gBOT: gBOT=getballs  'Uncomment if you're destroying balls - Not recommended! #SaveTheBalls

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
      If gBOT(s).Z < 30 And gBOT(s).X < 850 AND notrtlane = 1 Then 'Parameters for where the shadows can show, here they are not visible above the table (no upper pf) or in the plunger lane
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
Sub InitPolarity()
   dim x, a : a = Array(LF, RF)
  for each x in a
    x.AddPt "Ycoef", 0, RightFlipper.Y-65, 1 'disabled
    x.AddPt "Ycoef", 1, RightFlipper.Y-11, 1
    x.enabled = True
    x.TimeDelay = 80
    x.DebugOn=False ' prints some info in debugger

    x.AddPt "Polarity", 0, 0, 0
    x.AddPt "Polarity", 1, 0.05, - 2.7
    x.AddPt "Polarity", 2, 0.33, - 2.7
    x.AddPt "Polarity", 3, 0.37, - 2.7
    x.AddPt "Polarity", 4, 0.41, - 2.7
    x.AddPt "Polarity", 5, 0.45, - 2.7
    x.AddPt "Polarity", 6, 0.576, - 2.7
    x.AddPt "Polarity", 7, 0.66, - 1.8
    x.AddPt "Polarity", 8, 0.743, - 0.5
    x.AddPt "Polarity", 9, 0.81, - 0.5
    x.AddPt "Polarity", 10, 0.88, 0

    x.AddPt "Velocity", 0, 0, 1
    x.AddPt "Velocity", 1, 0.16, 1.06
    x.AddPt "Velocity", 2, 0.41, 1.05
    x.AddPt "Velocity", 3, 0.53, 1 '0.982
    x.AddPt "Velocity", 4, 0.702, 0.968
    x.AddPt "Velocity", 5, 0.95,  0.968
    x.AddPt "Velocity", 6, 1.03, 0.945
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
Dim EOST, EOSA, Frampup, FElasticity, FReturn
Dim RFEndAngle, LFEndAngle

Const FlipperCoilRampupMode = 2 '0 = fast, 1 = medium, 2 = slow (tap passes should work)

LFState = 1
RFState = 1
EOST = leftflipper.eostorque
EOSA = leftflipper.eostorqueangle
Frampup = LeftFlipper.rampup
FElasticity = LeftFlipper.elasticity
FReturn = LeftFlipper.return
Const EOSTnew = 1 'EM's to late 80's
'Const EOSTnew = 0.8 '90's and later
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
'Const EOSReturn = 0.035  'mid 80's to early 90's
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

Dim DT27, DT28, DT29, DT30, DT31, DT32, DT33, DT34, DT35, DT42, DT43, DT44, DT46, DT47, DT48


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


Set DT27 = (new DropTarget)(sw27, sw27a, BM_sw27p, 27, 0, false)
Set DT28 = (new DropTarget)(sw28, sw28a, BM_sw28p, 28, 0, false)
Set DT29 = (new DropTarget)(sw29, sw29a, BM_sw29p, 29, 0, false)
Set DT30 = (new DropTarget)(sw30, sw30a, BM_sw30p, 30, 0, false)
Set DT31 = (new DropTarget)(sw31, sw31a, BM_sw31p, 31, 0, false)
Set DT32 = (new DropTarget)(sw32, sw32a, BM_sw32p, 32, 0, false)
Set DT33 = (new DropTarget)(sw33, sw33a, BM_sw33p, 33, 0, false)
Set DT34 = (new DropTarget)(sw34, sw34a, BM_sw34p, 34, 0, false)
Set DT35 = (new DropTarget)(sw35, sw35a, BM_sw35p, 35, 0, false)
Set DT42 = (new DropTarget)(sw42, sw42a, BM_sw42p, 42, 0, false)
Set DT43 = (new DropTarget)(sw43, sw43a, BM_sw43p, 43, 0, false)
Set DT44 = (new DropTarget)(sw44, sw44a, BM_sw44p, 44, 0, false)
Set DT46 = (new DropTarget)(sw46, sw46a, BM_sw46p, 46, 0, false)
Set DT47 = (new DropTarget)(sw47, sw47a, BM_sw47p, 47, 0, false)
Set DT48 = (new DropTarget)(sw48, sw48a, BM_sw48p, 48, 0, false)


Dim DTArray
DTArray = Array(DT27, DT28, DT29, DT30, DT31, DT32, DT33, DT34, DT35, DT42, DT43, DT44, DT46, DT47, DT48)


'Configure the behavior of Drop Targets.
Const DTDropSpeed = 110 'in milliseconds
Const DTDropUpSpeed = 40 'in milliseconds
Const DTDropUnits = 44 'VP units primitive drops so top of at or below the playfield
Const DTDropUpUnits = 10 'VP units primitive raises above the up position on drops up
Const DTMaxBend = 2 '8 'max degrees primitive rotates when hit
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

Dim DTShadow(6)

DTShadowInit 1
DTShadowInit 2
DTShadowInit 3
DTShadowInit 4
DTShadowInit 5
DTShadowInit 6

' Initializes the drop targets for shadow logic below
Sub DTShadowInit(dtnbr)

  if dtnbr = 1 Then
    Set DTShadow(dtnbr) = Eval("dtsh" & 27)
  elseif dtnbr = 2 Then
    Set DTShadow(dtnbr) = Eval("dtsh" & 28)
  elseif dtnbr = 3 Then
    Set DTShadow(dtnbr) = Eval("dtsh" & 29)
  elseif dtnbr = 4 Then
    Set DTShadow(dtnbr) = Eval("dtsh" & 30)
  elseif dtnbr = 5 Then
    Set DTShadow(dtnbr) = Eval("dtsh" & 31)
  elseif dtnbr = 6 Then
    Set DTShadow(dtnbr) = Eval("dtsh" & 32)
  End If
End Sub

Sub DTHit(switch)
  Dim i, swmod

  i = DTArrayID(switch)
  If switch = 27 Then
    swmod = 1
  Elseif switch = 28 then
    swmod = 2
  Elseif switch = 29 then
    swmod = 3
  Elseif switch = 30 then
    swmod = 4
  Elseif switch = 31 then
    swmod = 5
  Elseif switch = 32 then
    swmod = 6
  Elseif switch = 33 then
    swmod = 7
  Elseif switch = 34 then
    swmod = 8
  Elseif switch = 35 then
    swmod = 9
  Elseif switch = 42 then
    swmod = 10
  Elseif switch = 43 then
    swmod = 11
  Elseif switch = 44 then
    swmod = 12
  Elseif switch = 46 then
    swmod = 13
  Elseif switch = 47 then
    swmod = 14
  Elseif switch = 48 then
    swmod = 15
  End If

  PlayTargetSound
  DTArray(i).animate =  DTCheckBrick(Activeball,DTArray(i).prim)
  If DTArray(i).animate = 1 or DTArray(i).animate = 3 or DTArray(i).animate = 4 Then
    DTBallPhysics Activeball, DTArray(i).prim.rotz, DTMass

' Controls Drop Shadow for a direct hit only
    if swmod=1 or swmod=2 or swmod=3 or swmod=4 or swmod=5 or swmod=6 then
      DTShadow(swmod).visible = 0
    End If
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
      controller.Switch(Switchid) = 1
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
    controller.Switch(Switchid) = 0
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

Dim ST17

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


Set ST17 = (new StandupTarget)(sw17, BM_sw17p,17, 0)



Dim STArray
STArray = Array(ST17)


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
    vpmTimer.PulseSw switch
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







Sub bumperbulb1flash(ByVal aLvl)  'argument is unused
  if Lampz.UseFunction then aLvl = LampFilter(aLvl) 'Callbacks don't get this filter automatically
  'FlBumperFadeTarget(1) = aLvl*bumperflashlevel
End Sub

Sub bumperbulb2flash(ByVal aLvl)  'argument is unused
  if Lampz.UseFunction then aLvl = LampFilter(aLvl) 'Callbacks don't get this filter automatically
  'FlBumperFadeTarget(2) = aLvl*bumperflashlevel
End Sub

Sub bumperbulb3flash(ByVal aLvl)  'argument is unused
  if Lampz.UseFunction then aLvl = LampFilter(aLvl) 'Callbacks don't get this filter automatically
  'FlBumperFadeTarget(3) = aLvl*bumperflashlevel
End Sub

'*******************************************
'  Backwall Lamps during Drop Targets up or down
'*******************************************
Sub dt66light(ByVal aLvl)
  if dt44up = 0 Then
    l66b_pf.IntensityScale = aLvl
    l66b_pf2.IntensityScale =0
  elseif dt44up = 1 Then
    l66b_pf2.IntensityScale = aLvl
    l66b_pf.IntensityScale = 0
  end If
End Sub

Sub dt82light(ByVal aLvl)
  if dt44up = 1 AND dt43up = 1 AND dt42up = 1 Then
    l82a_pf.IntensityScale = 0
    l82a_pf2.IntensityScale =0
    l82a_pf3.IntensityScale =0
    l82a_pf4.IntensityScale =aLvl
  elseif dt44up = 0 AND dt43up = 1 AND dt42up = 1 Then
    l82a_pf.IntensityScale = 0
    l82a_pf2.IntensityScale =aLvl
    l82a_pf3.IntensityScale =0
    l82a_pf4.IntensityScale =0
  elseif dt44up = 0 AND dt43up = 0 AND dt42up = 1 Then
    l82a_pf.IntensityScale = 0
    l82a_pf2.IntensityScale =0
    l82a_pf3.IntensityScale =aLvl
    l82a_pf4.IntensityScale =0
  Else 'all drop targets down
    l82a_pf.IntensityScale = aLvl
    l82a_pf2.IntensityScale =0
    l82a_pf3.IntensityScale =0
    l82a_pf4.IntensityScale =0
  End If
End Sub

Sub dt98light(ByVal aLvl)
  if dt42up = 0 Then
    l98a_pf.IntensityScale = aLvl
    l98a_pf2.IntensityScale =0
  elseif dt42up = 1 Then
    l98a_pf2.IntensityScale = aLvl
    l98a_pf.IntensityScale = 0
  end If
End Sub


'*******************************************
' Hybrid code for VR, Cab, and Desktop
'*******************************************
Sub SetupRoom

  Dim VRThings, BP

  ' Running in Desktop mode
  if VR_Room = 0 and cab_mode = 0 Then
    for each VRThings in VRStuff:VRThings.visible = 0:Next
    for each VRThings in VRMin:VRThings.visible = 0:Next
    for each VRThings in VRClock:VRThings.visible = 0:Next
    for each VRThings in VRBackglass:VRThings.visible = 0:Next
    'for each VRThings in DTRails:VRThings.visible = 1:Next

    for each VRThings in OceanEnvironment:VRThings.visible = 0:Next
    OuterPrimBlack_dt.visible = 0
    OuterPrimBlack_cab_VR.visible = 0
    OuterPrimBlack_cab.visible = 0
    L_DT_1.State = 1
    L_DT_2.State = 1
    SharkTimer.enabled = 0

  ' Running in Cabinet mode
  Elseif VR_Room = 0 and cab_mode = 1 Then
    for each VRThings in VRStuff:VRThings.visible = 0:Next
    for each VRThings in VRMin:VRThings.visible = 0:Next
    for each VRThings in VRClock:VRThings.visible = 0:Next
    for each VRThings in VRBackglass:VRThings.visible = 0:Next
    for each VRThings in DTBackglass:VRThings.visible = 0: Next
    for each VRThings in DTRails:VRThings.visible = 0:Next
    for each VRThings in OceanEnvironment:VRThings.visible = 0:Next
    OuterPrimBlack_dt.visible = 0
    OuterPrimBlack_cab_VR.visible = 0
    OuterPrimBlack_cab.visible = 0
    L_DT_1.Visible = 0
    L_DT_1.State = 0
    L_DT_2.Visible = 0
    L_DT_2.State = 0
    SharkTimer.enabled = 0

  Else
    ' We are in VR and using the Room environment
    if OceanEnv = 0 Then
      for each VRThings in VRStuff:VRThings.visible = 1:Next
      for each VRThings in VRMin:VRThings.visible = 1:Next
'     for each VRThings in VRClock:VRThings.visible = 0:Next
      for each VRThings in DTBackglass:VRThings.visible = 0: Next
      'for each VRThings in DTRails:VRThings.visible = 0:Next
      for each VRThings in OceanEnvironment:VRThings.visible = 0:Next
      VR_Cabinet.blenddisablelighting = .6
      VR_Backbox.blenddisablelighting = .6
      For Each BP in BP_Ramp001 : BP.Visible = 0: Next
      For Each BP in BP_Ramp002 : BP.Visible = 0: Next
      For Each BP in BP_lockdownbar : BP.Visible = 0: Next
      'For Each BP in BP_OuterPrimBlack_DT : BP.Visible = 0: Next
      'OuterPrimBlack_dt.visible = 0
      'OuterPrimBlack_cab_VR.visible = 1
      'OuterPrimBlack_cab.visible = 0
      L_DT_1.Visible = 0
      L_DT_1.State = 0
      L_DT_2.Visible = 0
      L_DT_2.State = 0
      SharkTimer.enabled = 0

    ' We are in VR and using the Ocean environment
    Else
      for each VRThings in VRStuff:VRThings.visible = 1:Next
      for each VRThings in VRMin:VRThings.visible = 0:Next
      for each VRThings in VRClock:VRThings.visible = 0:Next
      for each VRThings in DTBackglass:VRThings.visible = 0: Next
      'for each VRThings in DTRails:VRThings.visible = 0:Next
      for each VRThings in OceanEnvironment:VRThings.visible = 1:Next
      VR_Cabinet.blenddisablelighting = 1
      VR_Backbox.blenddisablelighting = 1
      For Each BP in BP_Ramp001 : BP.Visible = 0: Next
      For Each BP in BP_Ramp002 : BP.Visible = 0: Next
      For Each BP in BP_lockdownbar : BP.Visible = 0: Next
      'For Each BP in BP_OuterPrimBlack_DT : BP.Visible = 0: Next
      'OuterPrimBlack_dt.visible = 0
      'OuterPrimBlack_cab_VR.visible = 1
      'OuterPrimBlack_cab.visible = 0
      L_DT_1.Visible = 0
      L_DT_1.State = 0
      L_DT_2.Visible = 0
      L_DT_2.State = 0

      Primary_topper.visible = 0
      VRposter.visible = 0
      VRposter2.visible = 0


      VR_Shark.PlayAnimEndless(0.25)
      VR_Shark1.PlayAnimEndless(0.25)
      VR_Angelfish_1.PlayAnimEndless(0.25)
      VR_Angelfish_2.PlayAnimEndless(0.23)
      VR_Angelfish_3.PlayAnimEndless(0.23)
      VR_Angelfish_4.PlayAnimEndless(0.25)
      VR_Angelfish_5.PlayAnimEndless(0.22)
      VR_Angelfish_6.PlayAnimEndless(0.23)
      VR_Angelfish_7.PlayAnimEndless(0.25)
      VR_Angelfish_8.PlayAnimEndless(0.23)
      VR_Angelfish_9.PlayAnimEndless(0.24)
      VR_Angelfish_10.PlayAnimEndless(0.25)
      VR_Angelfish_11.PlayAnimEndless(0.23)
      VR_Angelfish_12.PlayAnimEndless(0.25)
      VR_Angelfish_13.PlayAnimEndless(0.25)
      VR_Angelfish_14.PlayAnimEndless(0.23)
      VR_MahiMahi_1.PlayAnimEndless(0.10)
      VR_MahiMahi_2.PlayAnimEndless(0.09)
      VR_MahiMahi_3.PlayAnimEndless(0.08)
      VR_MahiMahi_4.PlayAnimEndless(0.11)
      VR_MahiMahi_5.PlayAnimEndless(0.085)
      VR_MahiMahi_6.PlayAnimEndless(0.095)
      VR_MahiMahi_7.PlayAnimEndless(0.15)
      SharkTimer.enabled = 1

    End If

  End If

End Sub

'*******************************************
' VR Ocean Environment animation
'*******************************************

Dim sharkSpeed : sharkSpeed = 0.04
Dim fishSpeed : fishSpeed = 0.02
Dim bigFishSpeed : bigFishSpeed = 0.02
Dim swimSpeedMod : swimSpeedMod = 1.0

Sub SharkTimer_Timer
  VR_Shark.ObjRotZ = VR_Shark.ObjRotZ + (sharkSpeed * swimSpeedMod)
  VR_Shark1.ObjRotZ = VR_Shark1.ObjRotZ - (sharkSpeed * swimSpeedMod)

  VR_Angelfish_1.ObjRotZ = VR_Angelfish_1.ObjRotZ - (fishSpeed * swimSpeedMod)
  VR_Angelfish_2.ObjRotZ = VR_Angelfish_2.ObjRotZ - (fishSpeed * swimSpeedMod)
  VR_Angelfish_3.ObjRotZ = VR_Angelfish_3.ObjRotZ - (fishSpeed * swimSpeedMod)
  VR_Angelfish_4.ObjRotZ = VR_Angelfish_4.ObjRotZ - (fishSpeed * swimSpeedMod)
  VR_Angelfish_5.ObjRotZ = VR_Angelfish_5.ObjRotZ - (fishSpeed * swimSpeedMod)
  VR_Angelfish_6.ObjRotZ = VR_Angelfish_6.ObjRotZ - (fishSpeed * swimSpeedMod)
  VR_Angelfish_7.ObjRotZ = VR_Angelfish_7.ObjRotZ - (fishSpeed * swimSpeedMod)
  VR_Angelfish_8.ObjRotZ = VR_Angelfish_8.ObjRotZ + (fishSpeed * swimSpeedMod)
  VR_Angelfish_9.ObjRotZ = VR_Angelfish_9.ObjRotZ + (fishSpeed * swimSpeedMod)
  VR_Angelfish_10.ObjRotZ = VR_Angelfish_10.ObjRotZ + (fishSpeed * swimSpeedMod)
  VR_Angelfish_11.ObjRotZ = VR_Angelfish_11.ObjRotZ - (fishSpeed * swimSpeedMod)
  VR_Angelfish_12.ObjRotZ = VR_Angelfish_12.ObjRotZ - (fishSpeed * swimSpeedMod)
  VR_Angelfish_13.ObjRotZ = VR_Angelfish_13.ObjRotZ - (fishSpeed * swimSpeedMod)
  VR_Angelfish_14.ObjRotZ = VR_Angelfish_14.ObjRotZ - (fishSpeed * swimSpeedMod)

  VR_MahiMahi_1.ObjRotZ = VR_MahiMahi_1.ObjRotZ - (bigFishSpeed * swimSpeedMod)
  VR_MahiMahi_2.ObjRotZ = VR_MahiMahi_2.ObjRotZ - (bigFishSpeed * swimSpeedMod)
  VR_MahiMahi_3.ObjRotZ = VR_MahiMahi_3.ObjRotZ - (bigFishSpeed * swimSpeedMod)
  VR_MahiMahi_4.ObjRotZ = VR_MahiMahi_4.ObjRotZ - (bigFishSpeed * swimSpeedMod)
  VR_MahiMahi_5.ObjRotZ = VR_MahiMahi_5.ObjRotZ - (bigFishSpeed * swimSpeedMod)
  VR_MahiMahi_6.ObjRotZ = VR_MahiMahi_6.ObjRotZ - (bigFishSpeed * swimSpeedMod)
  VR_MahiMahi_7.ObjRotZ = VR_MahiMahi_7.ObjRotZ - (bigFishSpeed * swimSpeedMod)
End Sub



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
  If PinCab_Shooter.Y < 2211.572 then
       PinCab_Shooter.Y = PinCab_Shooter.Y + 5
  End If
End Sub

Sub TimerVRPlunger1_Timer
  PinCab_Shooter.Y = 2121.572 + (5* Plunger.Position) -20
End Sub


'*******************************************
'Digital Display
'*******************************************

Dim Digits(31)
' 1st Player
Digits(0) = Array(LED10,LED11,LED12,LED13,LED14,LED15,LED16,LEDc17)
Digits(1) = Array(LED20,LED21,LED22,LED23,LED24,LED25,LED26)
Digits(2) = Array(LED30,LED31,LED32,LED33,LED34,LED35,LED36)
Digits(3) = Array(LED40,LED41,LED42,LED43,LED44,LED45,LED46,LEDc47)
Digits(4) = Array(LED50,LED51,LED52,LED53,LED54,LED55,LED56)
Digits(5) = Array(LED60,LED61,LED62,LED63,LED64,LED65,LED66)
Digits(6) = Array(LED70,LED71,LED72,LED73,LED74,LED75,LED76)

' 2nd Player
Digits(7) = Array(LED80,LED81,LED82,LED83,LED84,LED85,LED86,LEDc87)
Digits(8) = Array(LED90,LED91,LED92,LED93,LED94,LED95,LED96)
Digits(9) = Array(LED100,LED101,LED102,LED103,LED104,LED105,LED106)
Digits(10) = Array(LED110,LED111,LED112,LED113,LED114,LED115,LED116,LEDc117)
Digits(11) = Array(LED120,LED121,LED122,LED123,LED124,LED125,LED126)
Digits(12) = Array(LED130,LED131,LED132,LED133,LED134,LED135,LED136)
Digits(13) = Array(LED140,LED141,LED142,LED143,LED144,LED145,LED146)

' 3rd Player
Digits(14) = Array(LED150,LED151,LED152,LED153,LED154,LED155,LED156,LEDc157)
Digits(15) = Array(LED160,LED161,LED162,LED163,LED164,LED165,LED166)
Digits(16) = Array(LED170,LED171,LED172,LED173,LED174,LED175,LED176)
Digits(17) = Array(LED180,LED181,LED182,LED183,LED184,LED185,LED186,LEDc187)
Digits(18) = Array(LED190,LED191,LED192,LED193,LED194,LED195,LED196)
Digits(19) = Array(LED200,LED201,LED202,LED203,LED204,LED205,LED206)
Digits(20) = Array(LED210,LED211,LED212,LED213,LED214,LED215,LED216)

' 4th Player
Digits(21) = Array(LED220,LED221,LED222,LED223,LED224,LED225,LED226,LEDc227)
Digits(22) = Array(LED230,LED231,LED232,LED233,LED234,LED235,LED236)
Digits(23) = Array(LED240,LED241,LED242,LED243,LED244,LED245,LED246)
Digits(24) = Array(LED250,LED251,LED252,LED253,LED254,LED255,LED256,LEDc257)
Digits(25) = Array(LED260,LED261,LED262,LED263,LED264,LED265,LED266)
Digits(26) = Array(LED270,LED271,LED272,LED273,LED274,LED275,LED276)
Digits(27) = Array(LED280,LED281,LED282,LED283,LED284,LED285,LED286)

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
  zoff = 524
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
VRDigits(0) = Array(LED1x0,LED1x1,LED1x2,LED1x3,LED1x4,LED1x5,LED1x6,LEDc1x7)
VRDigits(1) = Array(LED2x0,LED2x1,LED2x2,LED2x3,LED2x4,LED2x5,LED2x6)
VRDigits(2) = Array(LED3x0,LED3x1,LED3x2,LED3x3,LED3x4,LED3x5,LED3x6)
VRDigits(3) = Array(LED4x0,LED4x1,LED4x2,LED4x3,LED4x4,LED4x5,LED4x6,LEDc4x7)
VRDigits(4) = Array(LED5x0,LED5x1,LED5x2,LED5x3,LED5x4,LED5x5,LED5x6)
VRDigits(5) = Array(LED6x0,LED6x1,LED6x2,LED6x3,LED6x4,LED6x5,LED6x6)
VRDigits(6) = Array(LED7x0,LED7x1,LED7x2,LED7x3,LED7x4,LED7x5,LED7x6)

' 2nd Player
VRDigits(7) = Array(LED8x0,LED8x1,LED8x2,LED8x3,LED8x4,LED8x5,LED8x6,LEDc8x7)
VRDigits(8) = Array(LED9x0,LED9x1,LED9x2,LED9x3,LED9x4,LED9x5,LED9x6)
VRDigits(9) = Array(LED10x0,LED10x1,LED10x2,LED10x3,LED10x4,LED10x5,LED10x6)
VRDigits(10) = Array(LED11x0,LED11x1,LED11x2,LED11x3,LED11x4,LED11x5,LED11x6,LEDc11x7)
VRDigits(11) = Array(LED12x0,LED12x1,LED12x2,LED12x3,LED12x4,LED12x5,LED12x6)
VRDigits(12) = Array(LED13x0,LED13x1,LED13x2,LED13x3,LED13x4,LED13x5,LED13x6)
VRDigits(13) = Array(LED14x0,LED14x1,LED14x2,LED14x3,LED14x4,LED14x5,LED14x6)

' 3rd Player
VRDigits(14) = Array(LED1x000,LED1x001,LED1x002,LED1x003,LED1x004,LED1x005,LED1x006,LEDc1x007)
VRDigits(15) = Array(LED1x100,LED1x101,LED1x102,LED1x103,LED1x104,LED1x105,LED1x106)
VRDigits(16) = Array(LED1x200,LED1x201,LED1x202,LED1x203,LED1x204,LED1x205,LED1x206)
VRDigits(17) = Array(LED1x300,LED1x301,LED1x302,LED1x303,LED1x304,LED1x305,LED1x306,LEDc1x307)
VRDigits(18) = Array(LED1x400,LED1x401,LED1x402,LED1x403,LED1x404,LED1x405,LED1x406)
VRDigits(19) = Array(LED1x500,LED1x501,LED1x502,LED1x503,LED1x504,LED1x505,LED1x506)
VRDigits(20) = Array(LED1x600,LED1x601,LED1x602,LED1x603,LED1x604,LED1x605,LED1x606)

' 4th Player
VRDigits(21) = Array(LED2x000,LED2x001,LED2x002,LED2x003,LED2x004,LED2x005,LED2x006,LEDc2x007)
VRDigits(22) = Array(LED2x100,LED2x101,LED2x102,LED2x103,LED2x104,LED2x105,LED2x106)
VRDigits(23) = Array(LED2x200,LED2x201,LED2x202,LED2x203,LED2x204,LED2x205,LED2x206)
VRDigits(24) = Array(LED2x300,LED2x301,LED2x302,LED2x303,LED2x304,LED2x305,LED2x306,LEDc2x307)
VRDigits(25) = Array(LED2x400,LED2x401,LED2x402,LED2x403,LED2x404,LED2x405,LED2x406)
VRDigits(26) = Array(LED2x500,LED2x501,LED2x502,LED2x503,LED2x504,LED2x505,LED2x506)
VRDigits(27) = Array(LED2x600,LED2x601,LED2x602,LED2x603,LED2x604,LED2x605,LED2x606)

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

'If VR_Room=1 Then
' InitDigits
'End If

'*******************************************
' LAMP CALLBACK for the 6 backglass flasher lamps (not the solenoid conrolled ones)
'*******************************************


'if VR_Room = 0 and cab_mode = 0 Then
' Set LampCallback = GetRef("UpdateDTLamps")
'End If
'
'if VR_Room = 1 Then UpdateVRLamps
' Set LampCallback = GetRef("UpdateVRLamps")
'End If

Sub UpdateDTLamps()
  If Controller.Lamp(43) = 0 Then: ShootAgainReel.setValue(0):  Else: ShootAgainReel.setValue(1) 'Shoot Again
  If Controller.Lamp(13) = 0 Then: BIPReel.setValue(0):     Else: BIPReel.setValue(1) 'Ball in Play
  If Controller.Lamp(61) = 0 Then: TiltReel.setValue(0):      Else: TiltReel.setValue(1) 'Tilt
  If Controller.Lamp(45) = 0 Then: GameOverReel.setValue(0):    Else: GameOverReel.setValue(1) 'Game Over
  If Controller.Lamp(27) = 0 Then: MatchReel.setValue(0):     Else: MatchReel.setValue(1) 'Match
  If Controller.Lamp(29) = 0 Then: HighScoreReel.setValue(0):   Else: HighScoreReel.setValue(1) 'High Score
End Sub


Sub UpdateVRLamps()
  If Controller.Lamp(43) = 0 Then: VR_ShootAgain.visible=0: else: VR_ShootAgain.visible=1 'Shoot again
  If Controller.Lamp(13) = 0 Then: VR_BIP.visible=0: else: VR_BIP.visible=1 'Ball in  play
  If Controller.Lamp(61) = 0 Then: VR_Tilt.visible=0: else: VR_Tilt.visible=1 'Tilt
  If Controller.Lamp(45) = 0 Then: VR_GameOver.visible=0: else: VR_GameOver.visible=1 'Game Over
  If Controller.Lamp(27) = 0 Then: VR_Match.visible=0: else: VR_Match.visible=1 'Match
  If Controller.Lamp(29) = 0 Then: VR_HighScore.visible=0: else: VR_HighScore.visible=1 'High Score
End Sub



'*******************************************
' Backglass GI and lamps
'*******************************************


Sub BGGIUpdates(ByVal aLvl)

  'if Lampz.UseFunction then aLvl = LampFilter(aLvl)  'Callbacks don't get this filter automatically

  if vr_room = 1 and cab_mode = 0 then

    if alvl = 0 Then
      VR_Backglass.blenddisablelighting = 0.75
    Elseif aLvl = 1 then
      VR_Backglass.blenddisablelighting = 3.5
    Else
      VR_Backglass.blenddisablelighting = 2.75*alvl + 0.75
    end if
  elseif vr_room = 0 and cab_mode = 0 then
    if aLvl = 0 then
      L_DT_1.visible = 0
      L_DT_2.visible = 0
    Else
      L_DT_1.visible = 1
      L_DT_2.visible = 1
    End If
  Else
      L_DT_1.visible = 0
      L_DT_2.visible = 0
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



'******** Work done by UnclePaulie on Hybrid version 0.01 - 2.0 *********

' v.01  Started from Medusa as a base, and enabled all code for shadows, gBOT, sounds, rolling, hybrid, base table physics, flippers, flipper physics, materials, images, LUT, POV, ball colors, fastflips, lampz, trough logic, knocker
'   All graphics came from Redbone for playfield, plastics, apron, bumper, drop targets, spinner and others.
'   Updated the bumper image to blend with Flupper bumpers.
'   Enabled gBOT and trough system for 3 balls
'   Added option for flippers to be red, blue, or yellow.
'   Added bumpers.  Modified Flupper bumper code solution for this table.  Added constants to control DL on bumpers, and light/flash level
'   Enabled right flipper functionality
'   Updated colors of bumpers, pegs, plastic lane overlays to blue.
'   Updated the Apron image (redbone) and apron physics.  Also the plunger cover primitive.
'   Ensured the DT targets were not multi-dimension array assignments
' v.02  Added playfield mesh and saucers.
'   Enabled updated saucer functionality.
'   Updated the apron visuals and physics.
'   Put a different pincab shooter primitive in.
'   Added walls and rails from goldchicco's table.  Had to modify, move, place, etc.
'   Modified goldchicco's backwall prim in blender, as it wrapped around saucer (real one doesn't do that)
'   Updated desktop backglass image to something less busy. Added backglass lamps
'   Added all posts, pegs, rubber physics, and associated physics and sounds.
'   Added walls and collidables, flippers, rail guides, apron, plunger, balls, rollovers, etc.
'   Added small leaf sensors, slings, spinner, spinner prim, gates, star trigger and animation.
'   Created new guide rails
'   Created new blender gate prims for lower and top gates.
' v.03  Added plastics and acrylics, redid graphics.
'   Redid the bottom gates in blender after the lower plastic wall was in, to get the right height on one leg.
'   Redid all element placement with EBisLit playfield.
'   Corrected bumper lamp assignments.
'   Added metal screws and white screw caps.
'   Updated physics walls in plastics areas.
'     Added stand up and drop targets code solution developed by Rothbauerw
' v.04  Updated the drop target images to match real targets.
'   Changed the solenoid callouts, as there's a lamp controller 47 that switches the active solenoid for 5 of them
'   Added the cutout prims for holes and GI bulbs
'   Redid playfield_mesh to align with exact saucer placement on playfield
'   Added GI for bulbs, top plastics, and temporary  playfield (to be removed when baking) and Lampz callouts for GI.
'   Changed depth bias of bulb prims under top rollovers
'   Adjusted heights of bulb prims and bulb halos under top left lower plastics.
'   Adjusted intensity of some GI, as it was too bright
'   Adjusted the desktop mode backglass lamps
'   Added GI bulbs to dynnamic sources
'   Added code for adjusting the ball brightness during GI.
'   Added dynamic shadows, and code to stop shadows from going into the plunger lane via notrtlane and rtlanesense trigger.
'   Color the plunger area on the playfield black.
'   Updated the playfield with insert cutouts
'   Lights under backwall prim and 3 under slim plastic on upper left are controlled via Lampz.
'   Added text overlay
'   Added 3D insert prims, lights, and blooms.
'     The lamps were not defined in manual.  The work to figure them out was done in a prior version.  I used those assignments, and verified.
' v.05  Updated the flipper physics, location, start and end angles, and ensure could replicate shots on online videos.
'   Changed the flippers to early 80's to match shots.  However increased strength, torque and torque angle.
'   Added baked GI lighting
'   Added baked AO shadows
'   Made rubbers near long gate a bit more bouncy.  Lowered friction on backwall.
'   Increased plunger release speed slightly.
'   Adjusted the playfield friction down to 0.21
'   Added a rubber by upper gate.  Also moved backwall lights slightly.
'   Adjusted the lower flipper start angle, and then some flipper physics.
'   Modded the backwall lamps to not go over prims, and added additional lights with drop target code to turn on or off instead of direct shadows.
'   Hours spent on studying videos for shots, physics, etc.
'   Updated table info / rules
'   Added Drop Target Shadows and added logic in script to control
'   Updated DTHit subroutines to account for a "soft" drop target hit.  Shadow will NOT come on with a light hit... target HAS to drop.
'   Adjusted location of top drop targets slightly
'   Updated cab pov
'   Updated playfield imasge to make the under bumper area black. Removed one hole that could be seen.
'   Adjusted saucer strength to match PAPA video.
'   Made drop targets a little narrower.
'   Updated user manual recommended DIP settings, and added a user setting to turn default DIP switches off, and allow to be changed via F6.
'     Tuned upper orbit wall and flippers to be able to make the upper orbit saucer on the right flipper shot.
'   PAPA said that you should be able to make it on the right flipper shot as well, and I can.
'   Removed unused images and materials.
'   Adjusted the GI Bake to be less bright in the middle.  Looked too washed out.
' v.06  Added the VR backglass image and associated lamps and text.  Image from Blacksad B2S.
'   Added code for GI for VR backglass
'   Removed playfield reflectivity from drop targets.  Was getting weird black lines/blocks on DT's. (exists in vpx 10.74 and earlier).
' v.07  Updated the drop targets to a new primitive.
'   Recommendations from apophis to reduce slingshot force (below 4), increase playfield friction to 0.25, and then add the latest flipper tricks code
'   Also updated other flipper physics.
'     After friction changes... needed to mod the top saucers again
'   Adjusted dL on white nuts down a touch
' v.08  Changed the zCol_playfield mesh pysics material to match playfield friction (didn't have to, I guess)
' v.09  Changed material on star triggers to "insertoff"... was missing for some reason.
'   Ensured latest code for targetbouncer
'   Updated sounds for ball collision
'   Updated the drop target primitive, and the associated images.
'   Added randomness to the plunger
'   Adjusted the depth bias and the dL on the star triggers.
'   Changed the material on the white screws to be a bit darker, and adjusted the GI levels some.
'   Redbone updated the graphics for the apron, plunger cover, and worn yellow flipper
' v.10  Adjusted the slingshot threshold.  The sensitivity was too high.
'   Redbone updated the white drop target images.  I lowered the dL on them slightly.
'   Reduced slings to 3.75
'   Adjusted the table slope to 5.5  Allowed me to change the flipper physics strength back down to 1800, and the upper right one to 2300.
'   Small adjustment to the curved walls coming to out/inlane gates.
' v.11  Updated plastics image from Redbone.  Had to move some lower screws and standoffs to align.
'   Required minor update to plastic image for holes and a minor update to GI and shadow bake image under plastic area by flipper.
'   Had to adjust the pegs by the middle right flipper.
'     Redbone modded graphics for targets, text overlay, plunger cover, plastics, bumper cover, and apron credit.
'   Hauntfreaks modded the VR Backglass image.
'   Moved bumper small lights up to 59 halo
' v.11o:Added Ocean environment option for VR based off Circus' VR table.
' v.12  DGrimmReaper converted the ocean images to webp and reduce the table size.
'   DGrimmReaper also adjusted the starfish and shellfish on top of the pinball machine to fit better on the table.
' v.13  Studlygoorite fixed all of the fish animations with the static renderings off
' v.14  Adjusted a couple sticky points near the flippers after the slope was adjusted. Essentially rounded a couple corners on plastics by flippers.
'   Updated the VR room flyers to Fathom.
'   Changed all the prims in the ocean environment to not be static. (enables the animation on the fish)
'   Changed the ambient and intro sounds to backglass position
'     Changed the metal gates material and added to GI routine.
' v.15  Defaulted the VR room to the ocean environment.
'   Updated the VR backglass to be a max of 3.5 dL.  Also changed the minimal room cab dL to 0.6
'   Rebaked the GI and AO Shadows, since I moved a couple standups near the flippers.
'   Added a very small wall block by left bumper.  Could get ball stuck with new slope.
'   Lowered the GI Bulb prims.  Looks better in VR.  Also lowered the associated GI halos.
'   Adjusted the bumper color slightly and the bumper scale down
'   Modified the arrow insert primitives.
'   Lowered the blue bulb dl to 100 and the yellow bulbs to 300
'   Removed logo from ball scratches
'   Adjusted material for VR Cab Metal objects
'   Redbone enhanced graphics for playfield, more flippers, bumper tops, spinner, cab body, topper, and small fixes on plastics holes
'v2.0.0 Changed images to webp format for further compression without losing quality.
'   Released
'V2.0.1 An issue showed up in vpx10.8 and a glow around balls.  Reimported webp images.


' Thank you to the VPW team for testing and graphics, especially redbone, apophis, bountybob, dgrimmreaper, studlygoorite, pinstratsdan, wylte, trytotilt, bietekwiet, astronasty, and hauntfreaks.

' VLM  Arrays - Start
' Arrays per baked part
Dim BP_Bumper1_Ring: BP_Bumper1_Ring=Array(BM_Bumper1_Ring, LM_L_Bmp1HL_Bumper1_Ring, LM_L_l44_Bumper1_Ring, LM_L_Bmp2HL_Bumper1_Ring, LM_L_l12_Bumper1_Ring, LM_L_Bmp3HL_Bumper1_Ring, LM_L_l28_Bumper1_Ring, LM_GI_RT_gi_005_Bumper1_Ring, LM_GI_RT_gi_006_Bumper1_Ring, LM_GI_RT_gi_007_Bumper1_Ring, LM_GI_RT_gi_017_Bumper1_Ring, LM_GI_Bumper1_Ring, LM_GI_RT_gi_023_Bumper1_Ring, LM_GI_RT_l66a_Bumper1_Ring, LM_GI_RT_l82_Bumper1_Ring)
Dim BP_Bumper1_Socket: BP_Bumper1_Socket=Array(BM_Bumper1_Socket, LM_L_Bmp1HL_Bumper1_Socket, LM_L_l44_Bumper1_Socket, LM_L_Bmp2HL_Bumper1_Socket, LM_L_l12_Bumper1_Socket, LM_L_Bmp3HL_Bumper1_Socket, LM_L_l28_Bumper1_Socket, LM_GI_RT_gi_006_Bumper1_Socket, LM_GI_RT_gi_007_Bumper1_Socket, LM_GI_RT_gi_017_Bumper1_Socket, LM_GI_Bumper1_Socket, LM_GI_RT_gi_023_Bumper1_Socket, LM_GI_RT_gi_025_Bumper1_Socket, LM_GI_RT_l66a_Bumper1_Socket, LM_GI_RT_l82_Bumper1_Socket, LM_GI_RT_l98_Bumper1_Socket)
Dim BP_Bumper2_Ring: BP_Bumper2_Ring=Array(BM_Bumper2_Ring, LM_L_l44_Bumper2_Ring, LM_L_Bmp2HL_Bumper2_Ring, LM_L_l12_Bumper2_Ring, LM_L_Bmp3HL_Bumper2_Ring, LM_L_l28_Bumper2_Ring, LM_GI_RT_gi_007_Bumper2_Ring, LM_GI_Bumper2_Ring, LM_GI_RT_gi_023_Bumper2_Ring, LM_GI_RT_gi_025_Bumper2_Ring, LM_L_l98a_Bumper2_Ring)
Dim BP_Bumper2_Socket: BP_Bumper2_Socket=Array(BM_Bumper2_Socket, LM_L_Bmp1HL_Bumper2_Socket, LM_L_l44_Bumper2_Socket, LM_L_Bmp2HL_Bumper2_Socket, LM_L_l12_Bumper2_Socket, LM_L_Bmp3HL_Bumper2_Socket, LM_L_l28_Bumper2_Socket, LM_GI_Bumper2_Socket, LM_GI_RT_gi_023_Bumper2_Socket, LM_GI_RT_gi_025_Bumper2_Socket)
Dim BP_Bumper3_Ring: BP_Bumper3_Ring=Array(BM_Bumper3_Ring, LM_L_Bmp1HL_Bumper3_Ring, LM_L_l44_Bumper3_Ring, LM_L_Bmp2HL_Bumper3_Ring, LM_L_l12_Bumper3_Ring, LM_L_Bmp3HL_Bumper3_Ring, LM_L_l28_Bumper3_Ring, LM_GI_RT_gi_004_Bumper3_Ring, LM_GI_RT_gi_005_Bumper3_Ring, LM_GI_RT_gi_006_Bumper3_Ring, LM_GI_RT_gi_007_Bumper3_Ring, LM_GI_Bumper3_Ring, LM_GI_RT_gi_025_Bumper3_Ring, LM_GI_RT_l66a_Bumper3_Ring, LM_L_l82a_Bumper3_Ring)
Dim BP_Bumper3_Socket: BP_Bumper3_Socket=Array(BM_Bumper3_Socket, LM_L_Bmp1HL_Bumper3_Socket, LM_L_l44_Bumper3_Socket, LM_L_Bmp2HL_Bumper3_Socket, LM_L_l12_Bumper3_Socket, LM_L_Bmp3HL_Bumper3_Socket, LM_L_l28_Bumper3_Socket, LM_GI_RT_gi_005_Bumper3_Socket, LM_GI_RT_gi_006_Bumper3_Socket, LM_GI_RT_gi_007_Bumper3_Socket, LM_GI_RT_gi_017_Bumper3_Socket, LM_GI_Bumper3_Socket, LM_GI_RT_gi_023_Bumper3_Socket, LM_GI_RT_gi_025_Bumper3_Socket, LM_GI_RT_l66a_Bumper3_Socket)
Dim BP_Gate1p: BP_Gate1p=Array(BM_Gate1p, LM_GI_RT_gi_002_Gate1p, LM_GI_RT_gi_003_Gate1p, LM_GI_RT_gi_027_Gate1p, LM_GI_RT_gi_028_Gate1p, LM_GI_RT_gi_029_Gate1p)
Dim BP_Gate2p: BP_Gate2p=Array(BM_Gate2p, LM_GI_RT_gi_001_Gate2p, LM_GI_RT_gi_003_Gate2p, LM_GI_RT_gi_027_Gate2p)
Dim BP_Gate3p: BP_Gate3p=Array(BM_Gate3p, LM_L_l12_Gate3p, LM_GI_Gate3p, LM_GI_RT_gi_023_Gate3p)
Dim BP_Gate4p: BP_Gate4p=Array(BM_Gate4p, LM_L_l44_Gate4p, LM_L_Bmp2HL_Gate4p, LM_L_l12_Gate4p, LM_L_l28_Gate4p, LM_GI_Gate4p, LM_GI_RT_gi_023_Gate4p, LM_L_l113_Gate4p, LM_L_l66b_Gate4p, LM_GI_RT_l98_Gate4p)
Dim BP_Gate5_Wire: BP_Gate5_Wire=Array(BM_Gate5_Wire)
Dim BP_Kicker1Arm: BP_Kicker1Arm=Array(BM_Kicker1Arm, LM_GI_Kicker1Arm, LM_GI_RT_gi_023_Kicker1Arm)
Dim BP_Kicker1Hole: BP_Kicker1Hole=Array(BM_Kicker1Hole, LM_L_l12_Kicker1Hole, LM_GI_Kicker1Hole, LM_GI_RT_gi_023_Kicker1Hole)
Dim BP_Kicker2Arm: BP_Kicker2Arm=Array(BM_Kicker2Arm, LM_GI_Kicker2Arm)
Dim BP_Kicker2Hole: BP_Kicker2Hole=Array(BM_Kicker2Hole, LM_L_l12_Kicker2Hole, LM_L_l28_Kicker2Hole, LM_GI_Kicker2Hole, LM_GI_RT_gi_023_Kicker2Hole, LM_GI_RT_gi_025_Kicker2Hole)
Dim BP_LSling: BP_LSling=Array(BM_LSling, LM_GI_RT_gi_001_LSling, LM_GI_RT_gi_003_LSling, LM_GI_RT_gi_004_LSling, LM_GI_RT_gi_027_LSling, LM_GI_RT_gi_028_LSling)
Dim BP_LSling1: BP_LSling1=Array(BM_LSling1, LM_GI_RT_gi_001_LSling1, LM_GI_RT_gi_003_LSling1, LM_GI_RT_gi_004_LSling1, LM_GI_RT_gi_027_LSling1, LM_GI_RT_gi_028_LSling1)
Dim BP_LSling2: BP_LSling2=Array(BM_LSling2, LM_GI_RT_gi_001_LSling2, LM_GI_RT_gi_003_LSling2, LM_GI_RT_gi_004_LSling2, LM_GI_RT_gi_027_LSling2, LM_L_l33_LSling2)
Dim BP_Lflipmesh1: BP_Lflipmesh1=Array(BM_Lflipmesh1, LM_GI_RT_gi_003_Lflipmesh1, LM_GI_RT_gi_027_Lflipmesh1, LM_GI_RT_gi_028_Lflipmesh1)
Dim BP_Lflipmesh1U: BP_Lflipmesh1U=Array(BM_Lflipmesh1U, LM_GI_RT_gi_001_Lflipmesh1U, LM_GI_RT_gi_003_Lflipmesh1U, LM_GI_RT_gi_027_Lflipmesh1U, LM_GI_RT_gi_028_Lflipmesh1U, LM_GI_RT_gi_029_Lflipmesh1U)
Dim BP_OuterPrimBlack_DT: BP_OuterPrimBlack_DT=Array(BM_OuterPrimBlack_DT, LM_L_l44_OuterPrimBlack_DT, LM_L_l12_OuterPrimBlack_DT, LM_L_l28_OuterPrimBlack_DT, LM_GI_RT_gi_004_OuterPrimBlack_, LM_GI_RT_gi_005_OuterPrimBlack_, LM_GI_RT_gi_006_OuterPrimBlack_, LM_GI_RT_gi_007_OuterPrimBlack_, LM_GI_OuterPrimBlack_DT, LM_GI_RT_gi_023_OuterPrimBlack_, LM_L_l15_OuterPrimBlack_DT, LM_L_l49_OuterPrimBlack_DT, LM_L_l81_OuterPrimBlack_DT, LM_GI_RT_l82_OuterPrimBlack_DT)
Dim BP_Overlay: BP_Overlay=Array(BM_Overlay, LM_L_Bmp1HL_Overlay, LM_L_l44_Overlay, LM_L_Bmp2HL_Overlay, LM_L_l12_Overlay, LM_L_Bmp3HL_Overlay, LM_L_l28_Overlay, LM_GI_RT_gi_007_Overlay, LM_GI_RT_gi_017_Overlay, LM_GI_Overlay, LM_GI_RT_gi_023_Overlay, LM_GI_RT_gi_025_Overlay, LM_L_l113_Overlay, LM_L_l65_Overlay, LM_GI_RT_l66a_Overlay, LM_L_l66c_Overlay, LM_L_l81_Overlay, LM_GI_RT_l82_Overlay, LM_L_l82a_Overlay, LM_L_l97_Overlay, LM_GI_RT_l98_Overlay)
Dim BP_Parts: BP_Parts=Array(BM_Parts, LM_L_Bmp1HL_Parts, LM_L_l44_Parts, LM_L_Bmp2HL_Parts, LM_L_l12_Parts, LM_L_Bmp3HL_Parts, LM_L_l28_Parts, LM_GI_RT_gi_001_Parts, LM_GI_RT_gi_002_Parts, LM_GI_RT_gi_003_Parts, LM_GI_RT_gi_004_Parts, LM_GI_RT_gi_005_Parts, LM_GI_RT_gi_006_Parts, LM_GI_RT_gi_007_Parts, LM_GI_RT_gi_017_Parts, LM_GI_Parts, LM_GI_RT_gi_023_Parts, LM_GI_RT_gi_025_Parts, LM_GI_RT_gi_026_Parts, LM_GI_RT_gi_027_Parts, LM_GI_RT_gi_028_Parts, LM_GI_RT_gi_029_Parts, LM_L_l1_Parts, LM_L_l10_Parts, LM_L_l113_Parts, LM_L_l17_Parts, LM_L_l25_Parts, LM_L_l26_Parts, LM_L_l33_Parts, LM_L_l42_Parts, LM_L_l49_Parts, LM_L_l59_Parts, LM_L_l60_Parts, LM_L_l65_Parts, LM_L_l66_Parts, LM_GI_RT_l66a_Parts, LM_L_l66b_Parts, LM_L_l66c_Parts, LM_L_l81_Parts, LM_GI_RT_l82_Parts, LM_L_l82a_Parts, LM_L_l97_Parts, LM_GI_RT_l98_Parts, LM_L_l98a_Parts)
Dim BP_PartsST: BP_PartsST=Array(BM_PartsST, LM_L_Bmp1HL_PartsST, LM_L_l44_PartsST, LM_L_Bmp2HL_PartsST, LM_L_l12_PartsST, LM_L_Bmp3HL_PartsST, LM_L_l28_PartsST, LM_GI_RT_gi_017_PartsST, LM_GI_PartsST, LM_GI_RT_gi_023_PartsST, LM_GI_RT_gi_025_PartsST, LM_GI_RT_gi_026_PartsST, LM_GI_RT_l66a_PartsST, LM_L_l66c_PartsST, LM_GI_RT_l82_PartsST, LM_L_l97_PartsST, LM_GI_RT_l98_PartsST)
Dim BP_PlasticsSolid: BP_PlasticsSolid=Array(BM_PlasticsSolid, LM_L_l44_PlasticsSolid, LM_L_l28_PlasticsSolid, LM_GI_RT_gi_001_PlasticsSolid, LM_GI_RT_gi_002_PlasticsSolid, LM_GI_RT_gi_003_PlasticsSolid, LM_GI_RT_gi_004_PlasticsSolid, LM_GI_RT_gi_005_PlasticsSolid, LM_GI_RT_gi_006_PlasticsSolid, LM_GI_RT_gi_007_PlasticsSolid, LM_GI_RT_gi_026_PlasticsSolid, LM_GI_RT_gi_027_PlasticsSolid, LM_GI_RT_gi_028_PlasticsSolid, LM_GI_RT_gi_029_PlasticsSolid, LM_L_l1_PlasticsSolid, LM_L_l17_PlasticsSolid, LM_L_l65_PlasticsSolid)
Dim BP_Playfield: BP_Playfield=Array(BM_Playfield, LM_L_Bmp1HL_Playfield, LM_L_l44_Playfield, LM_L_Bmp2HL_Playfield, LM_L_l12_Playfield, LM_L_Bmp3HL_Playfield, LM_L_l28_Playfield, LM_GI_RT_gi_001_Playfield, LM_GI_RT_gi_002_Playfield, LM_GI_RT_gi_003_Playfield, LM_GI_RT_gi_004_Playfield, LM_GI_RT_gi_005_Playfield, LM_GI_RT_gi_006_Playfield, LM_GI_RT_gi_007_Playfield, LM_GI_RT_gi_017_Playfield, LM_GI_Playfield, LM_GI_RT_gi_023_Playfield, LM_GI_RT_gi_025_Playfield, LM_GI_RT_gi_026_Playfield, LM_GI_RT_gi_027_Playfield, LM_GI_RT_gi_028_Playfield, LM_GI_RT_gi_029_Playfield, LM_L_l113_Playfield, LM_L_l65_Playfield, LM_L_l66_Playfield, LM_GI_RT_l66a_Playfield, LM_L_l66b_Playfield, LM_L_l66c_Playfield, LM_L_l81_Playfield, LM_GI_RT_l82_Playfield, LM_L_l82a_Playfield, LM_L_l97_Playfield, LM_GI_RT_l98_Playfield, LM_L_l98a_Playfield)
Dim BP_RSling: BP_RSling=Array(BM_RSling, LM_GI_RT_gi_002_RSling, LM_GI_RT_gi_027_RSling, LM_GI_RT_gi_028_RSling, LM_GI_RT_gi_029_RSling)
Dim BP_RSling1: BP_RSling1=Array(BM_RSling1, LM_GI_RT_gi_027_RSling1, LM_GI_RT_gi_028_RSling1, LM_GI_RT_gi_029_RSling1)
Dim BP_RSling2: BP_RSling2=Array(BM_RSling2, LM_GI_RT_gi_027_RSling2, LM_GI_RT_gi_028_RSling2, LM_GI_RT_gi_029_RSling2)
Dim BP_Ramp001: BP_Ramp001=Array(BM_Ramp001, LM_L_l44_Ramp001, LM_L_l12_Ramp001, LM_L_Bmp3HL_Ramp001, LM_L_l28_Ramp001, LM_GI_RT_gi_023_Ramp001)
Dim BP_Ramp002: BP_Ramp002=Array(BM_Ramp002, LM_L_Bmp1HL_Ramp002, LM_L_l44_Ramp002, LM_L_Bmp3HL_Ramp002, LM_L_l28_Ramp002, LM_GI_Ramp002, LM_GI_RT_gi_025_Ramp002, LM_L_l81_Ramp002)
Dim BP_Rflipmesh1: BP_Rflipmesh1=Array(BM_Rflipmesh1, LM_GI_RT_gi_002_Rflipmesh1, LM_GI_RT_gi_027_Rflipmesh1, LM_GI_RT_gi_028_Rflipmesh1, LM_GI_RT_gi_029_Rflipmesh1)
Dim BP_Rflipmesh1U: BP_Rflipmesh1U=Array(BM_Rflipmesh1U, LM_GI_RT_gi_002_Rflipmesh1U, LM_GI_RT_gi_003_Rflipmesh1U, LM_GI_RT_gi_027_Rflipmesh1U, LM_GI_RT_gi_028_Rflipmesh1U, LM_GI_RT_gi_029_Rflipmesh1U)
Dim BP_RflipmeshUp: BP_RflipmeshUp=Array(BM_RflipmeshUp, LM_L_l28_RflipmeshUp, LM_GI_RT_gi_025_RflipmeshUp, LM_GI_RT_gi_026_RflipmeshUp, LM_L_l31_RflipmeshUp)
Dim BP_RflipmeshUpU: BP_RflipmeshUpU=Array(BM_RflipmeshUpU, LM_L_l28_RflipmeshUpU, LM_GI_RflipmeshUpU, LM_GI_RT_gi_025_RflipmeshUpU, LM_GI_RT_gi_026_RflipmeshUpU, LM_L_l31_RflipmeshUpU)
Dim BP_Sling1: BP_Sling1=Array(BM_Sling1, LM_GI_RT_gi_029_Sling1)
Dim BP_Sling2: BP_Sling2=Array(BM_Sling2, LM_GI_RT_gi_003_Sling2)
Dim BP_SpinnerPrim: BP_SpinnerPrim=Array(BM_SpinnerPrim, LM_L_Bmp1HL_SpinnerPrim, LM_L_l44_SpinnerPrim, LM_L_l12_SpinnerPrim, LM_L_Bmp3HL_SpinnerPrim, LM_L_l28_SpinnerPrim, LM_GI_RT_gi_007_SpinnerPrim, LM_GI_RT_gi_017_SpinnerPrim, LM_GI_SpinnerPrim, LM_L_l60_SpinnerPrim, LM_GI_RT_l66a_SpinnerPrim, LM_GI_RT_l82_SpinnerPrim, LM_GI_RT_l98_SpinnerPrim)
Dim BP_SpinnerRod: BP_SpinnerRod=Array(BM_SpinnerRod, LM_L_Bmp1HL_SpinnerRod, LM_L_l44_SpinnerRod, LM_GI_RT_gi_017_SpinnerRod, LM_GI_RT_l66a_SpinnerRod)
Dim BP_UnderPF: BP_UnderPF=Array(BM_UnderPF, LM_L_Bmp1HL_UnderPF, LM_L_l44_UnderPF, LM_L_Bmp2HL_UnderPF, LM_L_l12_UnderPF, LM_L_Bmp3HL_UnderPF, LM_L_l28_UnderPF, LM_GI_RT_gi_003_UnderPF, LM_GI_RT_gi_004_UnderPF, LM_GI_RT_gi_007_UnderPF, LM_GI_RT_gi_017_UnderPF, LM_GI_UnderPF, LM_GI_RT_gi_023_UnderPF, LM_GI_RT_gi_025_UnderPF, LM_GI_RT_gi_029_UnderPF, LM_L_l1_UnderPF, LM_L_l10_UnderPF, LM_L_l11_UnderPF, LM_L_l113_UnderPF, LM_L_l14_UnderPF, LM_L_l15_UnderPF, LM_L_l17_UnderPF, LM_L_l18_UnderPF, LM_L_l19_UnderPF, LM_L_l2_UnderPF, LM_L_l20_UnderPF, LM_L_l21_UnderPF, LM_L_l22_UnderPF, LM_L_l23_UnderPF, LM_L_l24_UnderPF, LM_L_l25_UnderPF, LM_L_l26_UnderPF, LM_L_l3_UnderPF, LM_L_l30_UnderPF, LM_L_l31_UnderPF, LM_L_l33_UnderPF, LM_L_l34_UnderPF, LM_L_l35_UnderPF, LM_L_l36_UnderPF, LM_L_l362_UnderPF, LM_L_l37_UnderPF, LM_L_l38_UnderPF, LM_L_l39_UnderPF, LM_L_l4_UnderPF, LM_L_l40_UnderPF, LM_L_l41_UnderPF, LM_L_l42_UnderPF, LM_L_l46_UnderPF, LM_L_l49_UnderPF, LM_L_l492_UnderPF, LM_L_l5_UnderPF, LM_L_l50_UnderPF, _
  LM_L_l51_UnderPF, LM_L_l52_UnderPF, LM_L_l53_UnderPF, LM_L_l54_UnderPF, LM_L_l55_UnderPF, LM_L_l56_UnderPF, LM_L_l57_UnderPF, LM_L_l58_UnderPF, LM_L_l59_UnderPF, LM_L_l6_UnderPF, LM_L_l60_UnderPF, LM_L_l62_UnderPF, LM_L_l63_UnderPF, LM_L_l65_UnderPF, LM_L_l66_UnderPF, LM_GI_RT_l66a_UnderPF, LM_L_l7_UnderPF, LM_L_l8_UnderPF, LM_L_l81_UnderPF, LM_GI_RT_l82_UnderPF, LM_L_l9_UnderPF, LM_L_l97_UnderPF, LM_GI_RT_l98_UnderPF, LM_L_l98a_UnderPF)
Dim BP_lockdownbar: BP_lockdownbar=Array(BM_lockdownbar, LM_L_Bmp3HL_lockdownbar, LM_L_l28_lockdownbar, LM_GI_lockdownbar, LM_GI_RT_gi_025_lockdownbar)
Dim BP_sw12: BP_sw12=Array(BM_sw12, LM_L_Bmp1HL_sw12, LM_L_l44_sw12, LM_L_Bmp2HL_sw12, LM_L_l12_sw12, LM_GI_sw12)
Dim BP_sw13: BP_sw13=Array(BM_sw13, LM_L_Bmp1HL_sw13, LM_L_l44_sw13, LM_L_Bmp2HL_sw13, LM_L_l12_sw13, LM_GI_sw13)
Dim BP_sw14: BP_sw14=Array(BM_sw14, LM_L_Bmp1HL_sw14, LM_L_l12_sw14, LM_GI_RT_gi_017_sw14, LM_GI_sw14)
Dim BP_sw17p: BP_sw17p=Array(BM_sw17p, LM_L_l12_sw17p, LM_L_Bmp3HL_sw17p, LM_L_l28_sw17p, LM_GI_RT_gi_006_sw17p, LM_GI_RT_gi_025_sw17p, LM_GI_RT_gi_026_sw17p)
Dim BP_sw21: BP_sw21=Array(BM_sw21, LM_GI_RT_gi_028_sw21, LM_GI_RT_gi_029_sw21)
Dim BP_sw22: BP_sw22=Array(BM_sw22, LM_GI_RT_gi_028_sw22, LM_GI_RT_gi_029_sw22)
Dim BP_sw23: BP_sw23=Array(BM_sw23, LM_GI_RT_gi_003_sw23, LM_GI_RT_gi_027_sw23)
Dim BP_sw24: BP_sw24=Array(BM_sw24, LM_GI_RT_gi_003_sw24, LM_GI_RT_gi_027_sw24)
Dim BP_sw27p: BP_sw27p=Array(BM_sw27p, LM_GI_RT_gi_004_sw27p, LM_GI_RT_gi_005_sw27p, LM_GI_RT_gi_006_sw27p)
Dim BP_sw28p: BP_sw28p=Array(BM_sw28p, LM_L_l44_sw28p, LM_L_l28_sw28p, LM_GI_RT_gi_004_sw28p, LM_GI_RT_gi_005_sw28p, LM_GI_RT_gi_006_sw28p)
Dim BP_sw29p: BP_sw29p=Array(BM_sw29p, LM_GI_RT_gi_004_sw29p, LM_GI_RT_gi_005_sw29p, LM_GI_RT_gi_006_sw29p)
Dim BP_sw30p: BP_sw30p=Array(BM_sw30p, LM_L_l44_sw30p, LM_GI_RT_gi_004_sw30p, LM_GI_RT_gi_005_sw30p, LM_GI_RT_gi_006_sw30p)
Dim BP_sw31p: BP_sw31p=Array(BM_sw31p, LM_L_l44_sw31p, LM_GI_RT_gi_005_sw31p, LM_GI_RT_gi_006_sw31p, LM_GI_RT_gi_007_sw31p)
Dim BP_sw32p: BP_sw32p=Array(BM_sw32p, LM_L_Bmp1HL_sw32p, LM_L_l44_sw32p, LM_L_l28_sw32p, LM_GI_RT_gi_005_sw32p, LM_GI_RT_gi_006_sw32p, LM_GI_RT_gi_007_sw32p)
Dim BP_sw33p: BP_sw33p=Array(BM_sw33p, LM_L_l44_sw33p, LM_L_l12_sw33p, LM_L_Bmp3HL_sw33p, LM_L_l28_sw33p, LM_GI_RT_gi_005_sw33p, LM_GI_RT_gi_006_sw33p, LM_GI_RT_gi_007_sw33p, LM_GI_sw33p, LM_GI_RT_gi_025_sw33p, LM_GI_RT_gi_026_sw33p)
Dim BP_sw34p: BP_sw34p=Array(BM_sw34p, LM_L_Bmp1HL_sw34p, LM_L_l44_sw34p, LM_L_l12_sw34p, LM_L_Bmp3HL_sw34p, LM_L_l28_sw34p, LM_GI_RT_gi_005_sw34p, LM_GI_RT_gi_006_sw34p, LM_GI_RT_gi_007_sw34p, LM_GI_sw34p, LM_GI_RT_gi_025_sw34p, LM_GI_RT_l66a_sw34p)
Dim BP_sw35p: BP_sw35p=Array(BM_sw35p, LM_L_Bmp1HL_sw35p, LM_L_l44_sw35p, LM_L_Bmp3HL_sw35p, LM_L_l28_sw35p, LM_GI_RT_gi_005_sw35p, LM_GI_RT_gi_006_sw35p, LM_GI_RT_gi_007_sw35p, LM_GI_sw35p, LM_GI_RT_gi_025_sw35p, LM_GI_RT_l66a_sw35p, LM_GI_RT_l82_sw35p)
Dim BP_sw42p: BP_sw42p=Array(BM_sw42p, LM_GI_sw42p)
Dim BP_sw43p: BP_sw43p=Array(BM_sw43p, LM_L_l12_sw43p, LM_GI_sw43p, LM_L_l82a_sw43p)
Dim BP_sw44p: BP_sw44p=Array(BM_sw44p, LM_L_l66b_sw44p)
Dim BP_sw46p: BP_sw46p=Array(BM_sw46p, LM_L_l12_sw46p, LM_L_l28_sw46p, LM_GI_sw46p, LM_GI_RT_gi_025_sw46p, LM_L_l66_sw46p)
Dim BP_sw47p: BP_sw47p=Array(BM_sw47p, LM_L_l12_sw47p, LM_GI_sw47p, LM_GI_RT_gi_025_sw47p)
Dim BP_sw48p: BP_sw48p=Array(BM_sw48p, LM_L_l12_sw48p, LM_L_l28_sw48p, LM_GI_sw48p, LM_GI_RT_gi_025_sw48p)
' Arrays per lighting scenario
Dim BL_GI: BL_GI=Array(LM_GI_Bumper1_Ring, LM_GI_Bumper1_Socket, LM_GI_Bumper2_Ring, LM_GI_Bumper2_Socket, LM_GI_Bumper3_Ring, LM_GI_Bumper3_Socket, LM_GI_Gate3p, LM_GI_Gate4p, LM_GI_Kicker1Arm, LM_GI_Kicker1Hole, LM_GI_Kicker2Arm, LM_GI_Kicker2Hole, LM_GI_OuterPrimBlack_DT, LM_GI_Overlay, LM_GI_Parts, LM_GI_PartsST, LM_GI_Playfield, LM_GI_Ramp002, LM_GI_RflipmeshUpU, LM_GI_SpinnerPrim, LM_GI_UnderPF, LM_GI_lockdownbar, LM_GI_sw12, LM_GI_sw13, LM_GI_sw14, LM_GI_sw33p, LM_GI_sw34p, LM_GI_sw35p, LM_GI_sw42p, LM_GI_sw43p, LM_GI_sw46p, LM_GI_sw47p, LM_GI_sw48p)
Dim BL_GI_RT_gi_001: BL_GI_RT_gi_001=Array(LM_GI_RT_gi_001_Gate2p, LM_GI_RT_gi_001_LSling, LM_GI_RT_gi_001_LSling1, LM_GI_RT_gi_001_LSling2, LM_GI_RT_gi_001_Lflipmesh1U, LM_GI_RT_gi_001_Parts, LM_GI_RT_gi_001_PlasticsSolid, LM_GI_RT_gi_001_Playfield)
Dim BL_GI_RT_gi_002: BL_GI_RT_gi_002=Array(LM_GI_RT_gi_002_Gate1p, LM_GI_RT_gi_002_Parts, LM_GI_RT_gi_002_PlasticsSolid, LM_GI_RT_gi_002_Playfield, LM_GI_RT_gi_002_RSling, LM_GI_RT_gi_002_Rflipmesh1, LM_GI_RT_gi_002_Rflipmesh1U)
Dim BL_GI_RT_gi_003: BL_GI_RT_gi_003=Array(LM_GI_RT_gi_003_Gate1p, LM_GI_RT_gi_003_Gate2p, LM_GI_RT_gi_003_LSling, LM_GI_RT_gi_003_LSling1, LM_GI_RT_gi_003_LSling2, LM_GI_RT_gi_003_Lflipmesh1, LM_GI_RT_gi_003_Lflipmesh1U, LM_GI_RT_gi_003_Parts, LM_GI_RT_gi_003_PlasticsSolid, LM_GI_RT_gi_003_Playfield, LM_GI_RT_gi_003_Rflipmesh1U, LM_GI_RT_gi_003_Sling2, LM_GI_RT_gi_003_UnderPF, LM_GI_RT_gi_003_sw23, LM_GI_RT_gi_003_sw24)
Dim BL_GI_RT_gi_004: BL_GI_RT_gi_004=Array(LM_GI_RT_gi_004_Bumper3_Ring, LM_GI_RT_gi_004_LSling, LM_GI_RT_gi_004_LSling1, LM_GI_RT_gi_004_LSling2, LM_GI_RT_gi_004_OuterPrimBlack_, LM_GI_RT_gi_004_Parts, LM_GI_RT_gi_004_PlasticsSolid, LM_GI_RT_gi_004_Playfield, LM_GI_RT_gi_004_UnderPF, LM_GI_RT_gi_004_sw27p, LM_GI_RT_gi_004_sw28p, LM_GI_RT_gi_004_sw29p, LM_GI_RT_gi_004_sw30p)
Dim BL_GI_RT_gi_005: BL_GI_RT_gi_005=Array(LM_GI_RT_gi_005_Bumper1_Ring, LM_GI_RT_gi_005_Bumper3_Ring, LM_GI_RT_gi_005_Bumper3_Socket, LM_GI_RT_gi_005_OuterPrimBlack_, LM_GI_RT_gi_005_Parts, LM_GI_RT_gi_005_PlasticsSolid, LM_GI_RT_gi_005_Playfield, LM_GI_RT_gi_005_sw27p, LM_GI_RT_gi_005_sw28p, LM_GI_RT_gi_005_sw29p, LM_GI_RT_gi_005_sw30p, LM_GI_RT_gi_005_sw31p, LM_GI_RT_gi_005_sw32p, LM_GI_RT_gi_005_sw33p, LM_GI_RT_gi_005_sw34p, LM_GI_RT_gi_005_sw35p)
Dim BL_GI_RT_gi_006: BL_GI_RT_gi_006=Array(LM_GI_RT_gi_006_Bumper1_Ring, LM_GI_RT_gi_006_Bumper1_Socket, LM_GI_RT_gi_006_Bumper3_Ring, LM_GI_RT_gi_006_Bumper3_Socket, LM_GI_RT_gi_006_OuterPrimBlack_, LM_GI_RT_gi_006_Parts, LM_GI_RT_gi_006_PlasticsSolid, LM_GI_RT_gi_006_Playfield, LM_GI_RT_gi_006_sw17p, LM_GI_RT_gi_006_sw27p, LM_GI_RT_gi_006_sw28p, LM_GI_RT_gi_006_sw29p, LM_GI_RT_gi_006_sw30p, LM_GI_RT_gi_006_sw31p, LM_GI_RT_gi_006_sw32p, LM_GI_RT_gi_006_sw33p, LM_GI_RT_gi_006_sw34p, LM_GI_RT_gi_006_sw35p)
Dim BL_GI_RT_gi_007: BL_GI_RT_gi_007=Array(LM_GI_RT_gi_007_Bumper1_Ring, LM_GI_RT_gi_007_Bumper1_Socket, LM_GI_RT_gi_007_Bumper2_Ring, LM_GI_RT_gi_007_Bumper3_Ring, LM_GI_RT_gi_007_Bumper3_Socket, LM_GI_RT_gi_007_OuterPrimBlack_, LM_GI_RT_gi_007_Overlay, LM_GI_RT_gi_007_Parts, LM_GI_RT_gi_007_PlasticsSolid, LM_GI_RT_gi_007_Playfield, LM_GI_RT_gi_007_SpinnerPrim, LM_GI_RT_gi_007_UnderPF, LM_GI_RT_gi_007_sw31p, LM_GI_RT_gi_007_sw32p, LM_GI_RT_gi_007_sw33p, LM_GI_RT_gi_007_sw34p, LM_GI_RT_gi_007_sw35p)
Dim BL_GI_RT_gi_017: BL_GI_RT_gi_017=Array(LM_GI_RT_gi_017_Bumper1_Ring, LM_GI_RT_gi_017_Bumper1_Socket, LM_GI_RT_gi_017_Bumper3_Socket, LM_GI_RT_gi_017_Overlay, LM_GI_RT_gi_017_Parts, LM_GI_RT_gi_017_PartsST, LM_GI_RT_gi_017_Playfield, LM_GI_RT_gi_017_SpinnerPrim, LM_GI_RT_gi_017_SpinnerRod, LM_GI_RT_gi_017_UnderPF, LM_GI_RT_gi_017_sw14)
Dim BL_GI_RT_gi_023: BL_GI_RT_gi_023=Array(LM_GI_RT_gi_023_Bumper1_Ring, LM_GI_RT_gi_023_Bumper1_Socket, LM_GI_RT_gi_023_Bumper2_Ring, LM_GI_RT_gi_023_Bumper2_Socket, LM_GI_RT_gi_023_Bumper3_Socket, LM_GI_RT_gi_023_Gate3p, LM_GI_RT_gi_023_Gate4p, LM_GI_RT_gi_023_Kicker1Arm, LM_GI_RT_gi_023_Kicker1Hole, LM_GI_RT_gi_023_Kicker2Hole, LM_GI_RT_gi_023_OuterPrimBlack_, LM_GI_RT_gi_023_Overlay, LM_GI_RT_gi_023_Parts, LM_GI_RT_gi_023_PartsST, LM_GI_RT_gi_023_Playfield, LM_GI_RT_gi_023_Ramp001, LM_GI_RT_gi_023_UnderPF)
Dim BL_GI_RT_gi_025: BL_GI_RT_gi_025=Array(LM_GI_RT_gi_025_Bumper1_Socket, LM_GI_RT_gi_025_Bumper2_Ring, LM_GI_RT_gi_025_Bumper2_Socket, LM_GI_RT_gi_025_Bumper3_Ring, LM_GI_RT_gi_025_Bumper3_Socket, LM_GI_RT_gi_025_Kicker2Hole, LM_GI_RT_gi_025_Overlay, LM_GI_RT_gi_025_Parts, LM_GI_RT_gi_025_PartsST, LM_GI_RT_gi_025_Playfield, LM_GI_RT_gi_025_Ramp002, LM_GI_RT_gi_025_RflipmeshUp, LM_GI_RT_gi_025_RflipmeshUpU, LM_GI_RT_gi_025_UnderPF, LM_GI_RT_gi_025_lockdownbar, LM_GI_RT_gi_025_sw17p, LM_GI_RT_gi_025_sw33p, LM_GI_RT_gi_025_sw34p, LM_GI_RT_gi_025_sw35p, LM_GI_RT_gi_025_sw46p, LM_GI_RT_gi_025_sw47p, LM_GI_RT_gi_025_sw48p)
Dim BL_GI_RT_gi_026: BL_GI_RT_gi_026=Array(LM_GI_RT_gi_026_Parts, LM_GI_RT_gi_026_PartsST, LM_GI_RT_gi_026_PlasticsSolid, LM_GI_RT_gi_026_Playfield, LM_GI_RT_gi_026_RflipmeshUp, LM_GI_RT_gi_026_RflipmeshUpU, LM_GI_RT_gi_026_sw17p, LM_GI_RT_gi_026_sw33p)
Dim BL_GI_RT_gi_027: BL_GI_RT_gi_027=Array(LM_GI_RT_gi_027_Gate1p, LM_GI_RT_gi_027_Gate2p, LM_GI_RT_gi_027_LSling, LM_GI_RT_gi_027_LSling1, LM_GI_RT_gi_027_LSling2, LM_GI_RT_gi_027_Lflipmesh1, LM_GI_RT_gi_027_Lflipmesh1U, LM_GI_RT_gi_027_Parts, LM_GI_RT_gi_027_PlasticsSolid, LM_GI_RT_gi_027_Playfield, LM_GI_RT_gi_027_RSling, LM_GI_RT_gi_027_RSling1, LM_GI_RT_gi_027_RSling2, LM_GI_RT_gi_027_Rflipmesh1, LM_GI_RT_gi_027_Rflipmesh1U, LM_GI_RT_gi_027_sw23, LM_GI_RT_gi_027_sw24)
Dim BL_GI_RT_gi_028: BL_GI_RT_gi_028=Array(LM_GI_RT_gi_028_Gate1p, LM_GI_RT_gi_028_LSling, LM_GI_RT_gi_028_LSling1, LM_GI_RT_gi_028_Lflipmesh1, LM_GI_RT_gi_028_Lflipmesh1U, LM_GI_RT_gi_028_Parts, LM_GI_RT_gi_028_PlasticsSolid, LM_GI_RT_gi_028_Playfield, LM_GI_RT_gi_028_RSling, LM_GI_RT_gi_028_RSling1, LM_GI_RT_gi_028_RSling2, LM_GI_RT_gi_028_Rflipmesh1, LM_GI_RT_gi_028_Rflipmesh1U, LM_GI_RT_gi_028_sw21, LM_GI_RT_gi_028_sw22)
Dim BL_GI_RT_gi_029: BL_GI_RT_gi_029=Array(LM_GI_RT_gi_029_Gate1p, LM_GI_RT_gi_029_Lflipmesh1U, LM_GI_RT_gi_029_Parts, LM_GI_RT_gi_029_PlasticsSolid, LM_GI_RT_gi_029_Playfield, LM_GI_RT_gi_029_RSling, LM_GI_RT_gi_029_RSling1, LM_GI_RT_gi_029_RSling2, LM_GI_RT_gi_029_Rflipmesh1, LM_GI_RT_gi_029_Rflipmesh1U, LM_GI_RT_gi_029_Sling1, LM_GI_RT_gi_029_UnderPF, LM_GI_RT_gi_029_sw21, LM_GI_RT_gi_029_sw22)
Dim BL_GI_RT_l66a: BL_GI_RT_l66a=Array(LM_GI_RT_l66a_Bumper1_Ring, LM_GI_RT_l66a_Bumper1_Socket, LM_GI_RT_l66a_Bumper3_Ring, LM_GI_RT_l66a_Bumper3_Socket, LM_GI_RT_l66a_Overlay, LM_GI_RT_l66a_Parts, LM_GI_RT_l66a_PartsST, LM_GI_RT_l66a_Playfield, LM_GI_RT_l66a_SpinnerPrim, LM_GI_RT_l66a_SpinnerRod, LM_GI_RT_l66a_UnderPF, LM_GI_RT_l66a_sw34p, LM_GI_RT_l66a_sw35p)
Dim BL_GI_RT_l82: BL_GI_RT_l82=Array(LM_GI_RT_l82_Bumper1_Ring, LM_GI_RT_l82_Bumper1_Socket, LM_GI_RT_l82_OuterPrimBlack_DT, LM_GI_RT_l82_Overlay, LM_GI_RT_l82_Parts, LM_GI_RT_l82_PartsST, LM_GI_RT_l82_Playfield, LM_GI_RT_l82_SpinnerPrim, LM_GI_RT_l82_UnderPF, LM_GI_RT_l82_sw35p)
Dim BL_GI_RT_l98: BL_GI_RT_l98=Array(LM_GI_RT_l98_Bumper1_Socket, LM_GI_RT_l98_Gate4p, LM_GI_RT_l98_Overlay, LM_GI_RT_l98_Parts, LM_GI_RT_l98_PartsST, LM_GI_RT_l98_Playfield, LM_GI_RT_l98_SpinnerPrim, LM_GI_RT_l98_UnderPF)
Dim BL_L_Bmp1HL: BL_L_Bmp1HL=Array(LM_L_Bmp1HL_Bumper1_Ring, LM_L_Bmp1HL_Bumper1_Socket, LM_L_Bmp1HL_Bumper2_Socket, LM_L_Bmp1HL_Bumper3_Ring, LM_L_Bmp1HL_Bumper3_Socket, LM_L_Bmp1HL_Overlay, LM_L_Bmp1HL_Parts, LM_L_Bmp1HL_PartsST, LM_L_Bmp1HL_Playfield, LM_L_Bmp1HL_Ramp002, LM_L_Bmp1HL_SpinnerPrim, LM_L_Bmp1HL_SpinnerRod, LM_L_Bmp1HL_UnderPF, LM_L_Bmp1HL_sw12, LM_L_Bmp1HL_sw13, LM_L_Bmp1HL_sw14, LM_L_Bmp1HL_sw32p, LM_L_Bmp1HL_sw34p, LM_L_Bmp1HL_sw35p)
Dim BL_L_Bmp2HL: BL_L_Bmp2HL=Array(LM_L_Bmp2HL_Bumper1_Ring, LM_L_Bmp2HL_Bumper1_Socket, LM_L_Bmp2HL_Bumper2_Ring, LM_L_Bmp2HL_Bumper2_Socket, LM_L_Bmp2HL_Bumper3_Ring, LM_L_Bmp2HL_Bumper3_Socket, LM_L_Bmp2HL_Gate4p, LM_L_Bmp2HL_Overlay, LM_L_Bmp2HL_Parts, LM_L_Bmp2HL_PartsST, LM_L_Bmp2HL_Playfield, LM_L_Bmp2HL_UnderPF, LM_L_Bmp2HL_sw12, LM_L_Bmp2HL_sw13)
Dim BL_L_Bmp3HL: BL_L_Bmp3HL=Array(LM_L_Bmp3HL_Bumper1_Ring, LM_L_Bmp3HL_Bumper1_Socket, LM_L_Bmp3HL_Bumper2_Ring, LM_L_Bmp3HL_Bumper2_Socket, LM_L_Bmp3HL_Bumper3_Ring, LM_L_Bmp3HL_Bumper3_Socket, LM_L_Bmp3HL_Overlay, LM_L_Bmp3HL_Parts, LM_L_Bmp3HL_PartsST, LM_L_Bmp3HL_Playfield, LM_L_Bmp3HL_Ramp001, LM_L_Bmp3HL_Ramp002, LM_L_Bmp3HL_SpinnerPrim, LM_L_Bmp3HL_UnderPF, LM_L_Bmp3HL_lockdownbar, LM_L_Bmp3HL_sw17p, LM_L_Bmp3HL_sw33p, LM_L_Bmp3HL_sw34p, LM_L_Bmp3HL_sw35p)
Dim BL_L_l1: BL_L_l1=Array(LM_L_l1_Parts, LM_L_l1_PlasticsSolid, LM_L_l1_UnderPF)
Dim BL_L_l10: BL_L_l10=Array(LM_L_l10_Parts, LM_L_l10_UnderPF)
Dim BL_L_l11: BL_L_l11=Array(LM_L_l11_UnderPF)
Dim BL_L_l113: BL_L_l113=Array(LM_L_l113_Gate4p, LM_L_l113_Overlay, LM_L_l113_Parts, LM_L_l113_Playfield, LM_L_l113_UnderPF)
Dim BL_L_l12: BL_L_l12=Array(LM_L_l12_Bumper1_Ring, LM_L_l12_Bumper1_Socket, LM_L_l12_Bumper2_Ring, LM_L_l12_Bumper2_Socket, LM_L_l12_Bumper3_Ring, LM_L_l12_Bumper3_Socket, LM_L_l12_Gate3p, LM_L_l12_Gate4p, LM_L_l12_Kicker1Hole, LM_L_l12_Kicker2Hole, LM_L_l12_OuterPrimBlack_DT, LM_L_l12_Overlay, LM_L_l12_Parts, LM_L_l12_PartsST, LM_L_l12_Playfield, LM_L_l12_Ramp001, LM_L_l12_SpinnerPrim, LM_L_l12_UnderPF, LM_L_l12_sw12, LM_L_l12_sw13, LM_L_l12_sw14, LM_L_l12_sw17p, LM_L_l12_sw33p, LM_L_l12_sw34p, LM_L_l12_sw43p, LM_L_l12_sw46p, LM_L_l12_sw47p, LM_L_l12_sw48p)
Dim BL_L_l14: BL_L_l14=Array(LM_L_l14_UnderPF)
Dim BL_L_l15: BL_L_l15=Array(LM_L_l15_OuterPrimBlack_DT, LM_L_l15_UnderPF)
Dim BL_L_l17: BL_L_l17=Array(LM_L_l17_Parts, LM_L_l17_PlasticsSolid, LM_L_l17_UnderPF)
Dim BL_L_l18: BL_L_l18=Array(LM_L_l18_UnderPF)
Dim BL_L_l19: BL_L_l19=Array(LM_L_l19_UnderPF)
Dim BL_L_l2: BL_L_l2=Array(LM_L_l2_UnderPF)
Dim BL_L_l20: BL_L_l20=Array(LM_L_l20_UnderPF)
Dim BL_L_l21: BL_L_l21=Array(LM_L_l21_UnderPF)
Dim BL_L_l22: BL_L_l22=Array(LM_L_l22_UnderPF)
Dim BL_L_l23: BL_L_l23=Array(LM_L_l23_UnderPF)
Dim BL_L_l24: BL_L_l24=Array(LM_L_l24_UnderPF)
Dim BL_L_l25: BL_L_l25=Array(LM_L_l25_Parts, LM_L_l25_UnderPF)
Dim BL_L_l26: BL_L_l26=Array(LM_L_l26_Parts, LM_L_l26_UnderPF)
Dim BL_L_l28: BL_L_l28=Array(LM_L_l28_Bumper1_Ring, LM_L_l28_Bumper1_Socket, LM_L_l28_Bumper2_Ring, LM_L_l28_Bumper2_Socket, LM_L_l28_Bumper3_Ring, LM_L_l28_Bumper3_Socket, LM_L_l28_Gate4p, LM_L_l28_Kicker2Hole, LM_L_l28_OuterPrimBlack_DT, LM_L_l28_Overlay, LM_L_l28_Parts, LM_L_l28_PartsST, LM_L_l28_PlasticsSolid, LM_L_l28_Playfield, LM_L_l28_Ramp001, LM_L_l28_Ramp002, LM_L_l28_RflipmeshUp, LM_L_l28_RflipmeshUpU, LM_L_l28_SpinnerPrim, LM_L_l28_UnderPF, LM_L_l28_lockdownbar, LM_L_l28_sw17p, LM_L_l28_sw28p, LM_L_l28_sw32p, LM_L_l28_sw33p, LM_L_l28_sw34p, LM_L_l28_sw35p, LM_L_l28_sw46p, LM_L_l28_sw48p)
Dim BL_L_l3: BL_L_l3=Array(LM_L_l3_UnderPF)
Dim BL_L_l30: BL_L_l30=Array(LM_L_l30_UnderPF)
Dim BL_L_l31: BL_L_l31=Array(LM_L_l31_RflipmeshUp, LM_L_l31_RflipmeshUpU, LM_L_l31_UnderPF)
Dim BL_L_l33: BL_L_l33=Array(LM_L_l33_LSling2, LM_L_l33_Parts, LM_L_l33_UnderPF)
Dim BL_L_l34: BL_L_l34=Array(LM_L_l34_UnderPF)
Dim BL_L_l35: BL_L_l35=Array(LM_L_l35_UnderPF)
Dim BL_L_l36: BL_L_l36=Array(LM_L_l36_UnderPF)
Dim BL_L_l362: BL_L_l362=Array(LM_L_l362_UnderPF)
Dim BL_L_l37: BL_L_l37=Array(LM_L_l37_UnderPF)
Dim BL_L_l38: BL_L_l38=Array(LM_L_l38_UnderPF)
Dim BL_L_l39: BL_L_l39=Array(LM_L_l39_UnderPF)
Dim BL_L_l4: BL_L_l4=Array(LM_L_l4_UnderPF)
Dim BL_L_l40: BL_L_l40=Array(LM_L_l40_UnderPF)
Dim BL_L_l41: BL_L_l41=Array(LM_L_l41_UnderPF)
Dim BL_L_l42: BL_L_l42=Array(LM_L_l42_Parts, LM_L_l42_UnderPF)
Dim BL_L_l44: BL_L_l44=Array(LM_L_l44_Bumper1_Ring, LM_L_l44_Bumper1_Socket, LM_L_l44_Bumper2_Ring, LM_L_l44_Bumper2_Socket, LM_L_l44_Bumper3_Ring, LM_L_l44_Bumper3_Socket, LM_L_l44_Gate4p, LM_L_l44_OuterPrimBlack_DT, LM_L_l44_Overlay, LM_L_l44_Parts, LM_L_l44_PartsST, LM_L_l44_PlasticsSolid, LM_L_l44_Playfield, LM_L_l44_Ramp001, LM_L_l44_Ramp002, LM_L_l44_SpinnerPrim, LM_L_l44_SpinnerRod, LM_L_l44_UnderPF, LM_L_l44_sw12, LM_L_l44_sw13, LM_L_l44_sw28p, LM_L_l44_sw30p, LM_L_l44_sw31p, LM_L_l44_sw32p, LM_L_l44_sw33p, LM_L_l44_sw34p, LM_L_l44_sw35p)
Dim BL_L_l46: BL_L_l46=Array(LM_L_l46_UnderPF)
Dim BL_L_l49: BL_L_l49=Array(LM_L_l49_OuterPrimBlack_DT, LM_L_l49_Parts, LM_L_l49_UnderPF)
Dim BL_L_l492: BL_L_l492=Array(LM_L_l492_UnderPF)
Dim BL_L_l5: BL_L_l5=Array(LM_L_l5_UnderPF)
Dim BL_L_l50: BL_L_l50=Array(LM_L_l50_UnderPF)
Dim BL_L_l51: BL_L_l51=Array(LM_L_l51_UnderPF)
Dim BL_L_l52: BL_L_l52=Array(LM_L_l52_UnderPF)
Dim BL_L_l53: BL_L_l53=Array(LM_L_l53_UnderPF)
Dim BL_L_l54: BL_L_l54=Array(LM_L_l54_UnderPF)
Dim BL_L_l55: BL_L_l55=Array(LM_L_l55_UnderPF)
Dim BL_L_l56: BL_L_l56=Array(LM_L_l56_UnderPF)
Dim BL_L_l57: BL_L_l57=Array(LM_L_l57_UnderPF)
Dim BL_L_l58: BL_L_l58=Array(LM_L_l58_UnderPF)
Dim BL_L_l59: BL_L_l59=Array(LM_L_l59_Parts, LM_L_l59_UnderPF)
Dim BL_L_l6: BL_L_l6=Array(LM_L_l6_UnderPF)
Dim BL_L_l60: BL_L_l60=Array(LM_L_l60_Parts, LM_L_l60_SpinnerPrim, LM_L_l60_UnderPF)
Dim BL_L_l62: BL_L_l62=Array(LM_L_l62_UnderPF)
Dim BL_L_l63: BL_L_l63=Array(LM_L_l63_UnderPF)
Dim BL_L_l65: BL_L_l65=Array(LM_L_l65_Overlay, LM_L_l65_Parts, LM_L_l65_PlasticsSolid, LM_L_l65_Playfield, LM_L_l65_UnderPF)
Dim BL_L_l66: BL_L_l66=Array(LM_L_l66_Parts, LM_L_l66_Playfield, LM_L_l66_UnderPF, LM_L_l66_sw46p)
Dim BL_L_l66b: BL_L_l66b=Array(LM_L_l66b_Gate4p, LM_L_l66b_Parts, LM_L_l66b_Playfield, LM_L_l66b_sw44p)
Dim BL_L_l66c: BL_L_l66c=Array(LM_L_l66c_Overlay, LM_L_l66c_Parts, LM_L_l66c_PartsST, LM_L_l66c_Playfield)
Dim BL_L_l7: BL_L_l7=Array(LM_L_l7_UnderPF)
Dim BL_L_l8: BL_L_l8=Array(LM_L_l8_UnderPF)
Dim BL_L_l81: BL_L_l81=Array(LM_L_l81_OuterPrimBlack_DT, LM_L_l81_Overlay, LM_L_l81_Parts, LM_L_l81_Playfield, LM_L_l81_Ramp002, LM_L_l81_UnderPF)
Dim BL_L_l82a: BL_L_l82a=Array(LM_L_l82a_Bumper3_Ring, LM_L_l82a_Overlay, LM_L_l82a_Parts, LM_L_l82a_Playfield, LM_L_l82a_sw43p)
Dim BL_L_l9: BL_L_l9=Array(LM_L_l9_UnderPF)
Dim BL_L_l97: BL_L_l97=Array(LM_L_l97_Overlay, LM_L_l97_Parts, LM_L_l97_PartsST, LM_L_l97_Playfield, LM_L_l97_UnderPF)
Dim BL_L_l98a: BL_L_l98a=Array(LM_L_l98a_Bumper2_Ring, LM_L_l98a_Parts, LM_L_l98a_Playfield, LM_L_l98a_UnderPF)
Dim BL_World: BL_World=Array(BM_Bumper1_Ring, BM_Bumper1_Socket, BM_Bumper2_Ring, BM_Bumper2_Socket, BM_Bumper3_Ring, BM_Bumper3_Socket, BM_Gate1p, BM_Gate2p, BM_Gate3p, BM_Gate4p, BM_Gate5_Wire, BM_Kicker1Arm, BM_Kicker1Hole, BM_Kicker2Arm, BM_Kicker2Hole, BM_LSling, BM_LSling1, BM_LSling2, BM_Lflipmesh1, BM_Lflipmesh1U, BM_OuterPrimBlack_DT, BM_Overlay, BM_Parts, BM_PartsST, BM_PlasticsSolid, BM_Playfield, BM_RSling, BM_RSling1, BM_RSling2, BM_Ramp001, BM_Ramp002, BM_Rflipmesh1, BM_Rflipmesh1U, BM_RflipmeshUp, BM_RflipmeshUpU, BM_Sling1, BM_Sling2, BM_SpinnerPrim, BM_SpinnerRod, BM_UnderPF, BM_lockdownbar, BM_sw12, BM_sw13, BM_sw14, BM_sw17p, BM_sw21, BM_sw22, BM_sw23, BM_sw24, BM_sw27p, BM_sw28p, BM_sw29p, BM_sw30p, BM_sw31p, BM_sw32p, BM_sw33p, BM_sw34p, BM_sw35p, BM_sw42p, BM_sw43p, BM_sw44p, BM_sw46p, BM_sw47p, BM_sw48p)
' Global arrays
Dim BG_Bakemap: BG_Bakemap=Array(BM_Bumper1_Ring, BM_Bumper1_Socket, BM_Bumper2_Ring, BM_Bumper2_Socket, BM_Bumper3_Ring, BM_Bumper3_Socket, BM_Gate1p, BM_Gate2p, BM_Gate3p, BM_Gate4p, BM_Gate5_Wire, BM_Kicker1Arm, BM_Kicker1Hole, BM_Kicker2Arm, BM_Kicker2Hole, BM_LSling, BM_LSling1, BM_LSling2, BM_Lflipmesh1, BM_Lflipmesh1U, BM_OuterPrimBlack_DT, BM_Overlay, BM_Parts, BM_PartsST, BM_PlasticsSolid, BM_Playfield, BM_RSling, BM_RSling1, BM_RSling2, BM_Ramp001, BM_Ramp002, BM_Rflipmesh1, BM_Rflipmesh1U, BM_RflipmeshUp, BM_RflipmeshUpU, BM_Sling1, BM_Sling2, BM_SpinnerPrim, BM_SpinnerRod, BM_UnderPF, BM_lockdownbar, BM_sw12, BM_sw13, BM_sw14, BM_sw17p, BM_sw21, BM_sw22, BM_sw23, BM_sw24, BM_sw27p, BM_sw28p, BM_sw29p, BM_sw30p, BM_sw31p, BM_sw32p, BM_sw33p, BM_sw34p, BM_sw35p, BM_sw42p, BM_sw43p, BM_sw44p, BM_sw46p, BM_sw47p, BM_sw48p)
Dim BG_Lightmap: BG_Lightmap=Array(LM_GI_Bumper1_Ring, LM_GI_Bumper1_Socket, LM_GI_Bumper2_Ring, LM_GI_Bumper2_Socket, LM_GI_Bumper3_Ring, LM_GI_Bumper3_Socket, LM_GI_Gate3p, LM_GI_Gate4p, LM_GI_Kicker1Arm, LM_GI_Kicker1Hole, LM_GI_Kicker2Arm, LM_GI_Kicker2Hole, LM_GI_OuterPrimBlack_DT, LM_GI_Overlay, LM_GI_Parts, LM_GI_PartsST, LM_GI_Playfield, LM_GI_Ramp002, LM_GI_RflipmeshUpU, LM_GI_SpinnerPrim, LM_GI_UnderPF, LM_GI_lockdownbar, LM_GI_sw12, LM_GI_sw13, LM_GI_sw14, LM_GI_sw33p, LM_GI_sw34p, LM_GI_sw35p, LM_GI_sw42p, LM_GI_sw43p, LM_GI_sw46p, LM_GI_sw47p, LM_GI_sw48p, LM_GI_RT_gi_001_Gate2p, LM_GI_RT_gi_001_LSling, LM_GI_RT_gi_001_LSling1, LM_GI_RT_gi_001_LSling2, LM_GI_RT_gi_001_Lflipmesh1U, LM_GI_RT_gi_001_Parts, LM_GI_RT_gi_001_PlasticsSolid, LM_GI_RT_gi_001_Playfield, LM_GI_RT_gi_002_Gate1p, LM_GI_RT_gi_002_Parts, LM_GI_RT_gi_002_PlasticsSolid, LM_GI_RT_gi_002_Playfield, LM_GI_RT_gi_002_RSling, LM_GI_RT_gi_002_Rflipmesh1, LM_GI_RT_gi_002_Rflipmesh1U, LM_GI_RT_gi_003_Gate1p, LM_GI_RT_gi_003_Gate2p, _
  LM_GI_RT_gi_003_LSling, LM_GI_RT_gi_003_LSling1, LM_GI_RT_gi_003_LSling2, LM_GI_RT_gi_003_Lflipmesh1, LM_GI_RT_gi_003_Lflipmesh1U, LM_GI_RT_gi_003_Parts, LM_GI_RT_gi_003_PlasticsSolid, LM_GI_RT_gi_003_Playfield, LM_GI_RT_gi_003_Rflipmesh1U, LM_GI_RT_gi_003_Sling2, LM_GI_RT_gi_003_UnderPF, LM_GI_RT_gi_003_sw23, LM_GI_RT_gi_003_sw24, LM_GI_RT_gi_004_Bumper3_Ring, LM_GI_RT_gi_004_LSling, LM_GI_RT_gi_004_LSling1, LM_GI_RT_gi_004_LSling2, LM_GI_RT_gi_004_OuterPrimBlack_, LM_GI_RT_gi_004_Parts, LM_GI_RT_gi_004_PlasticsSolid, LM_GI_RT_gi_004_Playfield, LM_GI_RT_gi_004_UnderPF, LM_GI_RT_gi_004_sw27p, LM_GI_RT_gi_004_sw28p, LM_GI_RT_gi_004_sw29p, LM_GI_RT_gi_004_sw30p, LM_GI_RT_gi_005_Bumper1_Ring, LM_GI_RT_gi_005_Bumper3_Ring, LM_GI_RT_gi_005_Bumper3_Socket, LM_GI_RT_gi_005_OuterPrimBlack_, LM_GI_RT_gi_005_Parts, LM_GI_RT_gi_005_PlasticsSolid, LM_GI_RT_gi_005_Playfield, LM_GI_RT_gi_005_sw27p, LM_GI_RT_gi_005_sw28p, LM_GI_RT_gi_005_sw29p, LM_GI_RT_gi_005_sw30p, LM_GI_RT_gi_005_sw31p, LM_GI_RT_gi_005_sw32p, _
  LM_GI_RT_gi_005_sw33p, LM_GI_RT_gi_005_sw34p, LM_GI_RT_gi_005_sw35p, LM_GI_RT_gi_006_Bumper1_Ring, LM_GI_RT_gi_006_Bumper1_Socket, LM_GI_RT_gi_006_Bumper3_Ring, LM_GI_RT_gi_006_Bumper3_Socket, LM_GI_RT_gi_006_OuterPrimBlack_, LM_GI_RT_gi_006_Parts, LM_GI_RT_gi_006_PlasticsSolid, LM_GI_RT_gi_006_Playfield, LM_GI_RT_gi_006_sw17p, LM_GI_RT_gi_006_sw27p, LM_GI_RT_gi_006_sw28p, LM_GI_RT_gi_006_sw29p, LM_GI_RT_gi_006_sw30p, LM_GI_RT_gi_006_sw31p, LM_GI_RT_gi_006_sw32p, LM_GI_RT_gi_006_sw33p, LM_GI_RT_gi_006_sw34p, LM_GI_RT_gi_006_sw35p, LM_GI_RT_gi_007_Bumper1_Ring, LM_GI_RT_gi_007_Bumper1_Socket, LM_GI_RT_gi_007_Bumper2_Ring, LM_GI_RT_gi_007_Bumper3_Ring, LM_GI_RT_gi_007_Bumper3_Socket, LM_GI_RT_gi_007_OuterPrimBlack_, LM_GI_RT_gi_007_Overlay, LM_GI_RT_gi_007_Parts, LM_GI_RT_gi_007_PlasticsSolid, LM_GI_RT_gi_007_Playfield, LM_GI_RT_gi_007_SpinnerPrim, LM_GI_RT_gi_007_UnderPF, LM_GI_RT_gi_007_sw31p, LM_GI_RT_gi_007_sw32p, LM_GI_RT_gi_007_sw33p, LM_GI_RT_gi_007_sw34p, LM_GI_RT_gi_007_sw35p, _
  LM_GI_RT_gi_017_Bumper1_Ring, LM_GI_RT_gi_017_Bumper1_Socket, LM_GI_RT_gi_017_Bumper3_Socket, LM_GI_RT_gi_017_Overlay, LM_GI_RT_gi_017_Parts, LM_GI_RT_gi_017_PartsST, LM_GI_RT_gi_017_Playfield, LM_GI_RT_gi_017_SpinnerPrim, LM_GI_RT_gi_017_SpinnerRod, LM_GI_RT_gi_017_UnderPF, LM_GI_RT_gi_017_sw14, LM_GI_RT_gi_023_Bumper1_Ring, LM_GI_RT_gi_023_Bumper1_Socket, LM_GI_RT_gi_023_Bumper2_Ring, LM_GI_RT_gi_023_Bumper2_Socket, LM_GI_RT_gi_023_Bumper3_Socket, LM_GI_RT_gi_023_Gate3p, LM_GI_RT_gi_023_Gate4p, LM_GI_RT_gi_023_Kicker1Arm, LM_GI_RT_gi_023_Kicker1Hole, LM_GI_RT_gi_023_Kicker2Hole, LM_GI_RT_gi_023_OuterPrimBlack_, LM_GI_RT_gi_023_Overlay, LM_GI_RT_gi_023_Parts, LM_GI_RT_gi_023_PartsST, LM_GI_RT_gi_023_Playfield, LM_GI_RT_gi_023_Ramp001, LM_GI_RT_gi_023_UnderPF, LM_GI_RT_gi_025_Bumper1_Socket, LM_GI_RT_gi_025_Bumper2_Ring, LM_GI_RT_gi_025_Bumper2_Socket, LM_GI_RT_gi_025_Bumper3_Ring, LM_GI_RT_gi_025_Bumper3_Socket, LM_GI_RT_gi_025_Kicker2Hole, LM_GI_RT_gi_025_Overlay, LM_GI_RT_gi_025_Parts, _
  LM_GI_RT_gi_025_PartsST, LM_GI_RT_gi_025_Playfield, LM_GI_RT_gi_025_Ramp002, LM_GI_RT_gi_025_RflipmeshUp, LM_GI_RT_gi_025_RflipmeshUpU, LM_GI_RT_gi_025_UnderPF, LM_GI_RT_gi_025_lockdownbar, LM_GI_RT_gi_025_sw17p, LM_GI_RT_gi_025_sw33p, LM_GI_RT_gi_025_sw34p, LM_GI_RT_gi_025_sw35p, LM_GI_RT_gi_025_sw46p, LM_GI_RT_gi_025_sw47p, LM_GI_RT_gi_025_sw48p, LM_GI_RT_gi_026_Parts, LM_GI_RT_gi_026_PartsST, LM_GI_RT_gi_026_PlasticsSolid, LM_GI_RT_gi_026_Playfield, LM_GI_RT_gi_026_RflipmeshUp, LM_GI_RT_gi_026_RflipmeshUpU, LM_GI_RT_gi_026_sw17p, LM_GI_RT_gi_026_sw33p, LM_GI_RT_gi_027_Gate1p, LM_GI_RT_gi_027_Gate2p, LM_GI_RT_gi_027_LSling, LM_GI_RT_gi_027_LSling1, LM_GI_RT_gi_027_LSling2, LM_GI_RT_gi_027_Lflipmesh1, LM_GI_RT_gi_027_Lflipmesh1U, LM_GI_RT_gi_027_Parts, LM_GI_RT_gi_027_PlasticsSolid, LM_GI_RT_gi_027_Playfield, LM_GI_RT_gi_027_RSling, LM_GI_RT_gi_027_RSling1, LM_GI_RT_gi_027_RSling2, LM_GI_RT_gi_027_Rflipmesh1, LM_GI_RT_gi_027_Rflipmesh1U, LM_GI_RT_gi_027_sw23, LM_GI_RT_gi_027_sw24, LM_GI_RT_gi_028_Gate1p, _
  LM_GI_RT_gi_028_LSling, LM_GI_RT_gi_028_LSling1, LM_GI_RT_gi_028_Lflipmesh1, LM_GI_RT_gi_028_Lflipmesh1U, LM_GI_RT_gi_028_Parts, LM_GI_RT_gi_028_PlasticsSolid, LM_GI_RT_gi_028_Playfield, LM_GI_RT_gi_028_RSling, LM_GI_RT_gi_028_RSling1, LM_GI_RT_gi_028_RSling2, LM_GI_RT_gi_028_Rflipmesh1, LM_GI_RT_gi_028_Rflipmesh1U, LM_GI_RT_gi_028_sw21, LM_GI_RT_gi_028_sw22, LM_GI_RT_gi_029_Gate1p, LM_GI_RT_gi_029_Lflipmesh1U, LM_GI_RT_gi_029_Parts, LM_GI_RT_gi_029_PlasticsSolid, LM_GI_RT_gi_029_Playfield, LM_GI_RT_gi_029_RSling, LM_GI_RT_gi_029_RSling1, LM_GI_RT_gi_029_RSling2, LM_GI_RT_gi_029_Rflipmesh1, LM_GI_RT_gi_029_Rflipmesh1U, LM_GI_RT_gi_029_Sling1, LM_GI_RT_gi_029_UnderPF, LM_GI_RT_gi_029_sw21, LM_GI_RT_gi_029_sw22, LM_GI_RT_l66a_Bumper1_Ring, LM_GI_RT_l66a_Bumper1_Socket, LM_GI_RT_l66a_Bumper3_Ring, LM_GI_RT_l66a_Bumper3_Socket, LM_GI_RT_l66a_Overlay, LM_GI_RT_l66a_Parts, LM_GI_RT_l66a_PartsST, LM_GI_RT_l66a_Playfield, LM_GI_RT_l66a_SpinnerPrim, LM_GI_RT_l66a_SpinnerRod, LM_GI_RT_l66a_UnderPF, LM_GI_RT_l66a_sw34p, _
  LM_GI_RT_l66a_sw35p, LM_GI_RT_l82_Bumper1_Ring, LM_GI_RT_l82_Bumper1_Socket, LM_GI_RT_l82_OuterPrimBlack_DT, LM_GI_RT_l82_Overlay, LM_GI_RT_l82_Parts, LM_GI_RT_l82_PartsST, LM_GI_RT_l82_Playfield, LM_GI_RT_l82_SpinnerPrim, LM_GI_RT_l82_UnderPF, LM_GI_RT_l82_sw35p, LM_GI_RT_l98_Bumper1_Socket, LM_GI_RT_l98_Gate4p, LM_GI_RT_l98_Overlay, LM_GI_RT_l98_Parts, LM_GI_RT_l98_PartsST, LM_GI_RT_l98_Playfield, LM_GI_RT_l98_SpinnerPrim, LM_GI_RT_l98_UnderPF, LM_L_Bmp1HL_Bumper1_Ring, LM_L_Bmp1HL_Bumper1_Socket, LM_L_Bmp1HL_Bumper2_Socket, LM_L_Bmp1HL_Bumper3_Ring, LM_L_Bmp1HL_Bumper3_Socket, LM_L_Bmp1HL_Overlay, LM_L_Bmp1HL_Parts, LM_L_Bmp1HL_PartsST, LM_L_Bmp1HL_Playfield, LM_L_Bmp1HL_Ramp002, LM_L_Bmp1HL_SpinnerPrim, LM_L_Bmp1HL_SpinnerRod, LM_L_Bmp1HL_UnderPF, LM_L_Bmp1HL_sw12, LM_L_Bmp1HL_sw13, LM_L_Bmp1HL_sw14, LM_L_Bmp1HL_sw32p, LM_L_Bmp1HL_sw34p, LM_L_Bmp1HL_sw35p, LM_L_Bmp2HL_Bumper1_Ring, LM_L_Bmp2HL_Bumper1_Socket, LM_L_Bmp2HL_Bumper2_Ring, LM_L_Bmp2HL_Bumper2_Socket, LM_L_Bmp2HL_Bumper3_Ring, _
  LM_L_Bmp2HL_Bumper3_Socket, LM_L_Bmp2HL_Gate4p, LM_L_Bmp2HL_Overlay, LM_L_Bmp2HL_Parts, LM_L_Bmp2HL_PartsST, LM_L_Bmp2HL_Playfield, LM_L_Bmp2HL_UnderPF, LM_L_Bmp2HL_sw12, LM_L_Bmp2HL_sw13, LM_L_Bmp3HL_Bumper1_Ring, LM_L_Bmp3HL_Bumper1_Socket, LM_L_Bmp3HL_Bumper2_Ring, LM_L_Bmp3HL_Bumper2_Socket, LM_L_Bmp3HL_Bumper3_Ring, LM_L_Bmp3HL_Bumper3_Socket, LM_L_Bmp3HL_Overlay, LM_L_Bmp3HL_Parts, LM_L_Bmp3HL_PartsST, LM_L_Bmp3HL_Playfield, LM_L_Bmp3HL_Ramp001, LM_L_Bmp3HL_Ramp002, LM_L_Bmp3HL_SpinnerPrim, LM_L_Bmp3HL_UnderPF, LM_L_Bmp3HL_lockdownbar, LM_L_Bmp3HL_sw17p, LM_L_Bmp3HL_sw33p, LM_L_Bmp3HL_sw34p, LM_L_Bmp3HL_sw35p, LM_L_l1_Parts, LM_L_l1_PlasticsSolid, LM_L_l1_UnderPF, LM_L_l10_Parts, LM_L_l10_UnderPF, LM_L_l11_UnderPF, LM_L_l113_Gate4p, LM_L_l113_Overlay, LM_L_l113_Parts, LM_L_l113_Playfield, LM_L_l113_UnderPF, LM_L_l12_Bumper1_Ring, LM_L_l12_Bumper1_Socket, LM_L_l12_Bumper2_Ring, LM_L_l12_Bumper2_Socket, LM_L_l12_Bumper3_Ring, LM_L_l12_Bumper3_Socket, LM_L_l12_Gate3p, LM_L_l12_Gate4p, LM_L_l12_Kicker1Hole, _
  LM_L_l12_Kicker2Hole, LM_L_l12_OuterPrimBlack_DT, LM_L_l12_Overlay, LM_L_l12_Parts, LM_L_l12_PartsST, LM_L_l12_Playfield, LM_L_l12_Ramp001, LM_L_l12_SpinnerPrim, LM_L_l12_UnderPF, LM_L_l12_sw12, LM_L_l12_sw13, LM_L_l12_sw14, LM_L_l12_sw17p, LM_L_l12_sw33p, LM_L_l12_sw34p, LM_L_l12_sw43p, LM_L_l12_sw46p, LM_L_l12_sw47p, LM_L_l12_sw48p, LM_L_l14_UnderPF, LM_L_l15_OuterPrimBlack_DT, LM_L_l15_UnderPF, LM_L_l17_Parts, LM_L_l17_PlasticsSolid, LM_L_l17_UnderPF, LM_L_l18_UnderPF, LM_L_l19_UnderPF, LM_L_l2_UnderPF, LM_L_l20_UnderPF, LM_L_l21_UnderPF, LM_L_l22_UnderPF, LM_L_l23_UnderPF, LM_L_l24_UnderPF, LM_L_l25_Parts, LM_L_l25_UnderPF, LM_L_l26_Parts, LM_L_l26_UnderPF, LM_L_l28_Bumper1_Ring, LM_L_l28_Bumper1_Socket, LM_L_l28_Bumper2_Ring, LM_L_l28_Bumper2_Socket, LM_L_l28_Bumper3_Ring, LM_L_l28_Bumper3_Socket, LM_L_l28_Gate4p, LM_L_l28_Kicker2Hole, LM_L_l28_OuterPrimBlack_DT, LM_L_l28_Overlay, LM_L_l28_Parts, LM_L_l28_PartsST, LM_L_l28_PlasticsSolid, LM_L_l28_Playfield, LM_L_l28_Ramp001, LM_L_l28_Ramp002, _
  LM_L_l28_RflipmeshUp, LM_L_l28_RflipmeshUpU, LM_L_l28_SpinnerPrim, LM_L_l28_UnderPF, LM_L_l28_lockdownbar, LM_L_l28_sw17p, LM_L_l28_sw28p, LM_L_l28_sw32p, LM_L_l28_sw33p, LM_L_l28_sw34p, LM_L_l28_sw35p, LM_L_l28_sw46p, LM_L_l28_sw48p, LM_L_l3_UnderPF, LM_L_l30_UnderPF, LM_L_l31_RflipmeshUp, LM_L_l31_RflipmeshUpU, LM_L_l31_UnderPF, LM_L_l33_LSling2, LM_L_l33_Parts, LM_L_l33_UnderPF, LM_L_l34_UnderPF, LM_L_l35_UnderPF, LM_L_l36_UnderPF, LM_L_l362_UnderPF, LM_L_l37_UnderPF, LM_L_l38_UnderPF, LM_L_l39_UnderPF, LM_L_l4_UnderPF, LM_L_l40_UnderPF, LM_L_l41_UnderPF, LM_L_l42_Parts, LM_L_l42_UnderPF, LM_L_l44_Bumper1_Ring, LM_L_l44_Bumper1_Socket, LM_L_l44_Bumper2_Ring, LM_L_l44_Bumper2_Socket, LM_L_l44_Bumper3_Ring, LM_L_l44_Bumper3_Socket, LM_L_l44_Gate4p, LM_L_l44_OuterPrimBlack_DT, LM_L_l44_Overlay, LM_L_l44_Parts, LM_L_l44_PartsST, LM_L_l44_PlasticsSolid, LM_L_l44_Playfield, LM_L_l44_Ramp001, LM_L_l44_Ramp002, LM_L_l44_SpinnerPrim, LM_L_l44_SpinnerRod, LM_L_l44_UnderPF, LM_L_l44_sw12, LM_L_l44_sw13, _
  LM_L_l44_sw28p, LM_L_l44_sw30p, LM_L_l44_sw31p, LM_L_l44_sw32p, LM_L_l44_sw33p, LM_L_l44_sw34p, LM_L_l44_sw35p, LM_L_l46_UnderPF, LM_L_l49_OuterPrimBlack_DT, LM_L_l49_Parts, LM_L_l49_UnderPF, LM_L_l492_UnderPF, LM_L_l5_UnderPF, LM_L_l50_UnderPF, LM_L_l51_UnderPF, LM_L_l52_UnderPF, LM_L_l53_UnderPF, LM_L_l54_UnderPF, LM_L_l55_UnderPF, LM_L_l56_UnderPF, LM_L_l57_UnderPF, LM_L_l58_UnderPF, LM_L_l59_Parts, LM_L_l59_UnderPF, LM_L_l6_UnderPF, LM_L_l60_Parts, LM_L_l60_SpinnerPrim, LM_L_l60_UnderPF, LM_L_l62_UnderPF, LM_L_l63_UnderPF, LM_L_l65_Overlay, LM_L_l65_Parts, LM_L_l65_PlasticsSolid, LM_L_l65_Playfield, LM_L_l65_UnderPF, LM_L_l66_Parts, LM_L_l66_Playfield, LM_L_l66_UnderPF, LM_L_l66_sw46p, LM_L_l66b_Gate4p, LM_L_l66b_Parts, LM_L_l66b_Playfield, LM_L_l66b_sw44p, LM_L_l66c_Overlay, LM_L_l66c_Parts, LM_L_l66c_PartsST, LM_L_l66c_Playfield, LM_L_l7_UnderPF, LM_L_l8_UnderPF, LM_L_l81_OuterPrimBlack_DT, LM_L_l81_Overlay, LM_L_l81_Parts, LM_L_l81_Playfield, LM_L_l81_Ramp002, LM_L_l81_UnderPF, LM_L_l82a_Bumper3_Ring, _
  LM_L_l82a_Overlay, LM_L_l82a_Parts, LM_L_l82a_Playfield, LM_L_l82a_sw43p, LM_L_l9_UnderPF, LM_L_l97_Overlay, LM_L_l97_Parts, LM_L_l97_PartsST, LM_L_l97_Playfield, LM_L_l97_UnderPF, LM_L_l98a_Bumper2_Ring, LM_L_l98a_Parts, LM_L_l98a_Playfield, LM_L_l98a_UnderPF)
Dim BG_All: BG_All=Array(BM_Bumper1_Ring, BM_Bumper1_Socket, BM_Bumper2_Ring, BM_Bumper2_Socket, BM_Bumper3_Ring, BM_Bumper3_Socket, BM_Gate1p, BM_Gate2p, BM_Gate3p, BM_Gate4p, BM_Gate5_Wire, BM_Kicker1Arm, BM_Kicker1Hole, BM_Kicker2Arm, BM_Kicker2Hole, BM_LSling, BM_LSling1, BM_LSling2, BM_Lflipmesh1, BM_Lflipmesh1U, BM_OuterPrimBlack_DT, BM_Overlay, BM_Parts, BM_PartsST, BM_PlasticsSolid, BM_Playfield, BM_RSling, BM_RSling1, BM_RSling2, BM_Ramp001, BM_Ramp002, BM_Rflipmesh1, BM_Rflipmesh1U, BM_RflipmeshUp, BM_RflipmeshUpU, BM_Sling1, BM_Sling2, BM_SpinnerPrim, BM_SpinnerRod, BM_UnderPF, BM_lockdownbar, BM_sw12, BM_sw13, BM_sw14, BM_sw17p, BM_sw21, BM_sw22, BM_sw23, BM_sw24, BM_sw27p, BM_sw28p, BM_sw29p, BM_sw30p, BM_sw31p, BM_sw32p, BM_sw33p, BM_sw34p, BM_sw35p, BM_sw42p, BM_sw43p, BM_sw44p, BM_sw46p, BM_sw47p, BM_sw48p, LM_GI_Bumper1_Ring, LM_GI_Bumper1_Socket, LM_GI_Bumper2_Ring, LM_GI_Bumper2_Socket, LM_GI_Bumper3_Ring, LM_GI_Bumper3_Socket, LM_GI_Gate3p, LM_GI_Gate4p, LM_GI_Kicker1Arm, _
  LM_GI_Kicker1Hole, LM_GI_Kicker2Arm, LM_GI_Kicker2Hole, LM_GI_OuterPrimBlack_DT, LM_GI_Overlay, LM_GI_Parts, LM_GI_PartsST, LM_GI_Playfield, LM_GI_Ramp002, LM_GI_RflipmeshUpU, LM_GI_SpinnerPrim, LM_GI_UnderPF, LM_GI_lockdownbar, LM_GI_sw12, LM_GI_sw13, LM_GI_sw14, LM_GI_sw33p, LM_GI_sw34p, LM_GI_sw35p, LM_GI_sw42p, LM_GI_sw43p, LM_GI_sw46p, LM_GI_sw47p, LM_GI_sw48p, LM_GI_RT_gi_001_Gate2p, LM_GI_RT_gi_001_LSling, LM_GI_RT_gi_001_LSling1, LM_GI_RT_gi_001_LSling2, LM_GI_RT_gi_001_Lflipmesh1U, LM_GI_RT_gi_001_Parts, LM_GI_RT_gi_001_PlasticsSolid, LM_GI_RT_gi_001_Playfield, LM_GI_RT_gi_002_Gate1p, LM_GI_RT_gi_002_Parts, LM_GI_RT_gi_002_PlasticsSolid, LM_GI_RT_gi_002_Playfield, LM_GI_RT_gi_002_RSling, LM_GI_RT_gi_002_Rflipmesh1, LM_GI_RT_gi_002_Rflipmesh1U, LM_GI_RT_gi_003_Gate1p, LM_GI_RT_gi_003_Gate2p, LM_GI_RT_gi_003_LSling, LM_GI_RT_gi_003_LSling1, LM_GI_RT_gi_003_LSling2, LM_GI_RT_gi_003_Lflipmesh1, LM_GI_RT_gi_003_Lflipmesh1U, LM_GI_RT_gi_003_Parts, LM_GI_RT_gi_003_PlasticsSolid, LM_GI_RT_gi_003_Playfield, _
  LM_GI_RT_gi_003_Rflipmesh1U, LM_GI_RT_gi_003_Sling2, LM_GI_RT_gi_003_UnderPF, LM_GI_RT_gi_003_sw23, LM_GI_RT_gi_003_sw24, LM_GI_RT_gi_004_Bumper3_Ring, LM_GI_RT_gi_004_LSling, LM_GI_RT_gi_004_LSling1, LM_GI_RT_gi_004_LSling2, LM_GI_RT_gi_004_OuterPrimBlack_, LM_GI_RT_gi_004_Parts, LM_GI_RT_gi_004_PlasticsSolid, LM_GI_RT_gi_004_Playfield, LM_GI_RT_gi_004_UnderPF, LM_GI_RT_gi_004_sw27p, LM_GI_RT_gi_004_sw28p, LM_GI_RT_gi_004_sw29p, LM_GI_RT_gi_004_sw30p, LM_GI_RT_gi_005_Bumper1_Ring, LM_GI_RT_gi_005_Bumper3_Ring, LM_GI_RT_gi_005_Bumper3_Socket, LM_GI_RT_gi_005_OuterPrimBlack_, LM_GI_RT_gi_005_Parts, LM_GI_RT_gi_005_PlasticsSolid, LM_GI_RT_gi_005_Playfield, LM_GI_RT_gi_005_sw27p, LM_GI_RT_gi_005_sw28p, LM_GI_RT_gi_005_sw29p, LM_GI_RT_gi_005_sw30p, LM_GI_RT_gi_005_sw31p, LM_GI_RT_gi_005_sw32p, LM_GI_RT_gi_005_sw33p, LM_GI_RT_gi_005_sw34p, LM_GI_RT_gi_005_sw35p, LM_GI_RT_gi_006_Bumper1_Ring, LM_GI_RT_gi_006_Bumper1_Socket, LM_GI_RT_gi_006_Bumper3_Ring, LM_GI_RT_gi_006_Bumper3_Socket, _
  LM_GI_RT_gi_006_OuterPrimBlack_, LM_GI_RT_gi_006_Parts, LM_GI_RT_gi_006_PlasticsSolid, LM_GI_RT_gi_006_Playfield, LM_GI_RT_gi_006_sw17p, LM_GI_RT_gi_006_sw27p, LM_GI_RT_gi_006_sw28p, LM_GI_RT_gi_006_sw29p, LM_GI_RT_gi_006_sw30p, LM_GI_RT_gi_006_sw31p, LM_GI_RT_gi_006_sw32p, LM_GI_RT_gi_006_sw33p, LM_GI_RT_gi_006_sw34p, LM_GI_RT_gi_006_sw35p, LM_GI_RT_gi_007_Bumper1_Ring, LM_GI_RT_gi_007_Bumper1_Socket, LM_GI_RT_gi_007_Bumper2_Ring, LM_GI_RT_gi_007_Bumper3_Ring, LM_GI_RT_gi_007_Bumper3_Socket, LM_GI_RT_gi_007_OuterPrimBlack_, LM_GI_RT_gi_007_Overlay, LM_GI_RT_gi_007_Parts, LM_GI_RT_gi_007_PlasticsSolid, LM_GI_RT_gi_007_Playfield, LM_GI_RT_gi_007_SpinnerPrim, LM_GI_RT_gi_007_UnderPF, LM_GI_RT_gi_007_sw31p, LM_GI_RT_gi_007_sw32p, LM_GI_RT_gi_007_sw33p, LM_GI_RT_gi_007_sw34p, LM_GI_RT_gi_007_sw35p, LM_GI_RT_gi_017_Bumper1_Ring, LM_GI_RT_gi_017_Bumper1_Socket, LM_GI_RT_gi_017_Bumper3_Socket, LM_GI_RT_gi_017_Overlay, LM_GI_RT_gi_017_Parts, LM_GI_RT_gi_017_PartsST, LM_GI_RT_gi_017_Playfield, _
  LM_GI_RT_gi_017_SpinnerPrim, LM_GI_RT_gi_017_SpinnerRod, LM_GI_RT_gi_017_UnderPF, LM_GI_RT_gi_017_sw14, LM_GI_RT_gi_023_Bumper1_Ring, LM_GI_RT_gi_023_Bumper1_Socket, LM_GI_RT_gi_023_Bumper2_Ring, LM_GI_RT_gi_023_Bumper2_Socket, LM_GI_RT_gi_023_Bumper3_Socket, LM_GI_RT_gi_023_Gate3p, LM_GI_RT_gi_023_Gate4p, LM_GI_RT_gi_023_Kicker1Arm, LM_GI_RT_gi_023_Kicker1Hole, LM_GI_RT_gi_023_Kicker2Hole, LM_GI_RT_gi_023_OuterPrimBlack_, LM_GI_RT_gi_023_Overlay, LM_GI_RT_gi_023_Parts, LM_GI_RT_gi_023_PartsST, LM_GI_RT_gi_023_Playfield, LM_GI_RT_gi_023_Ramp001, LM_GI_RT_gi_023_UnderPF, LM_GI_RT_gi_025_Bumper1_Socket, LM_GI_RT_gi_025_Bumper2_Ring, LM_GI_RT_gi_025_Bumper2_Socket, LM_GI_RT_gi_025_Bumper3_Ring, LM_GI_RT_gi_025_Bumper3_Socket, LM_GI_RT_gi_025_Kicker2Hole, LM_GI_RT_gi_025_Overlay, LM_GI_RT_gi_025_Parts, LM_GI_RT_gi_025_PartsST, LM_GI_RT_gi_025_Playfield, LM_GI_RT_gi_025_Ramp002, LM_GI_RT_gi_025_RflipmeshUp, LM_GI_RT_gi_025_RflipmeshUpU, LM_GI_RT_gi_025_UnderPF, LM_GI_RT_gi_025_lockdownbar, LM_GI_RT_gi_025_sw17p, _
  LM_GI_RT_gi_025_sw33p, LM_GI_RT_gi_025_sw34p, LM_GI_RT_gi_025_sw35p, LM_GI_RT_gi_025_sw46p, LM_GI_RT_gi_025_sw47p, LM_GI_RT_gi_025_sw48p, LM_GI_RT_gi_026_Parts, LM_GI_RT_gi_026_PartsST, LM_GI_RT_gi_026_PlasticsSolid, LM_GI_RT_gi_026_Playfield, LM_GI_RT_gi_026_RflipmeshUp, LM_GI_RT_gi_026_RflipmeshUpU, LM_GI_RT_gi_026_sw17p, LM_GI_RT_gi_026_sw33p, LM_GI_RT_gi_027_Gate1p, LM_GI_RT_gi_027_Gate2p, LM_GI_RT_gi_027_LSling, LM_GI_RT_gi_027_LSling1, LM_GI_RT_gi_027_LSling2, LM_GI_RT_gi_027_Lflipmesh1, LM_GI_RT_gi_027_Lflipmesh1U, LM_GI_RT_gi_027_Parts, LM_GI_RT_gi_027_PlasticsSolid, LM_GI_RT_gi_027_Playfield, LM_GI_RT_gi_027_RSling, LM_GI_RT_gi_027_RSling1, LM_GI_RT_gi_027_RSling2, LM_GI_RT_gi_027_Rflipmesh1, LM_GI_RT_gi_027_Rflipmesh1U, LM_GI_RT_gi_027_sw23, LM_GI_RT_gi_027_sw24, LM_GI_RT_gi_028_Gate1p, LM_GI_RT_gi_028_LSling, LM_GI_RT_gi_028_LSling1, LM_GI_RT_gi_028_Lflipmesh1, LM_GI_RT_gi_028_Lflipmesh1U, LM_GI_RT_gi_028_Parts, LM_GI_RT_gi_028_PlasticsSolid, LM_GI_RT_gi_028_Playfield, LM_GI_RT_gi_028_RSling, _
  LM_GI_RT_gi_028_RSling1, LM_GI_RT_gi_028_RSling2, LM_GI_RT_gi_028_Rflipmesh1, LM_GI_RT_gi_028_Rflipmesh1U, LM_GI_RT_gi_028_sw21, LM_GI_RT_gi_028_sw22, LM_GI_RT_gi_029_Gate1p, LM_GI_RT_gi_029_Lflipmesh1U, LM_GI_RT_gi_029_Parts, LM_GI_RT_gi_029_PlasticsSolid, LM_GI_RT_gi_029_Playfield, LM_GI_RT_gi_029_RSling, LM_GI_RT_gi_029_RSling1, LM_GI_RT_gi_029_RSling2, LM_GI_RT_gi_029_Rflipmesh1, LM_GI_RT_gi_029_Rflipmesh1U, LM_GI_RT_gi_029_Sling1, LM_GI_RT_gi_029_UnderPF, LM_GI_RT_gi_029_sw21, LM_GI_RT_gi_029_sw22, LM_GI_RT_l66a_Bumper1_Ring, LM_GI_RT_l66a_Bumper1_Socket, LM_GI_RT_l66a_Bumper3_Ring, LM_GI_RT_l66a_Bumper3_Socket, LM_GI_RT_l66a_Overlay, LM_GI_RT_l66a_Parts, LM_GI_RT_l66a_PartsST, LM_GI_RT_l66a_Playfield, LM_GI_RT_l66a_SpinnerPrim, LM_GI_RT_l66a_SpinnerRod, LM_GI_RT_l66a_UnderPF, LM_GI_RT_l66a_sw34p, LM_GI_RT_l66a_sw35p, LM_GI_RT_l82_Bumper1_Ring, LM_GI_RT_l82_Bumper1_Socket, LM_GI_RT_l82_OuterPrimBlack_DT, LM_GI_RT_l82_Overlay, LM_GI_RT_l82_Parts, LM_GI_RT_l82_PartsST, LM_GI_RT_l82_Playfield, _
  LM_GI_RT_l82_SpinnerPrim, LM_GI_RT_l82_UnderPF, LM_GI_RT_l82_sw35p, LM_GI_RT_l98_Bumper1_Socket, LM_GI_RT_l98_Gate4p, LM_GI_RT_l98_Overlay, LM_GI_RT_l98_Parts, LM_GI_RT_l98_PartsST, LM_GI_RT_l98_Playfield, LM_GI_RT_l98_SpinnerPrim, LM_GI_RT_l98_UnderPF, LM_L_Bmp1HL_Bumper1_Ring, LM_L_Bmp1HL_Bumper1_Socket, LM_L_Bmp1HL_Bumper2_Socket, LM_L_Bmp1HL_Bumper3_Ring, LM_L_Bmp1HL_Bumper3_Socket, LM_L_Bmp1HL_Overlay, LM_L_Bmp1HL_Parts, LM_L_Bmp1HL_PartsST, LM_L_Bmp1HL_Playfield, LM_L_Bmp1HL_Ramp002, LM_L_Bmp1HL_SpinnerPrim, LM_L_Bmp1HL_SpinnerRod, LM_L_Bmp1HL_UnderPF, LM_L_Bmp1HL_sw12, LM_L_Bmp1HL_sw13, LM_L_Bmp1HL_sw14, LM_L_Bmp1HL_sw32p, LM_L_Bmp1HL_sw34p, LM_L_Bmp1HL_sw35p, LM_L_Bmp2HL_Bumper1_Ring, LM_L_Bmp2HL_Bumper1_Socket, LM_L_Bmp2HL_Bumper2_Ring, LM_L_Bmp2HL_Bumper2_Socket, LM_L_Bmp2HL_Bumper3_Ring, LM_L_Bmp2HL_Bumper3_Socket, LM_L_Bmp2HL_Gate4p, LM_L_Bmp2HL_Overlay, LM_L_Bmp2HL_Parts, LM_L_Bmp2HL_PartsST, LM_L_Bmp2HL_Playfield, LM_L_Bmp2HL_UnderPF, LM_L_Bmp2HL_sw12, LM_L_Bmp2HL_sw13, LM_L_Bmp3HL_Bumper1_Ring, _
  LM_L_Bmp3HL_Bumper1_Socket, LM_L_Bmp3HL_Bumper2_Ring, LM_L_Bmp3HL_Bumper2_Socket, LM_L_Bmp3HL_Bumper3_Ring, LM_L_Bmp3HL_Bumper3_Socket, LM_L_Bmp3HL_Overlay, LM_L_Bmp3HL_Parts, LM_L_Bmp3HL_PartsST, LM_L_Bmp3HL_Playfield, LM_L_Bmp3HL_Ramp001, LM_L_Bmp3HL_Ramp002, LM_L_Bmp3HL_SpinnerPrim, LM_L_Bmp3HL_UnderPF, LM_L_Bmp3HL_lockdownbar, LM_L_Bmp3HL_sw17p, LM_L_Bmp3HL_sw33p, LM_L_Bmp3HL_sw34p, LM_L_Bmp3HL_sw35p, LM_L_l1_Parts, LM_L_l1_PlasticsSolid, LM_L_l1_UnderPF, LM_L_l10_Parts, LM_L_l10_UnderPF, LM_L_l11_UnderPF, LM_L_l113_Gate4p, LM_L_l113_Overlay, LM_L_l113_Parts, LM_L_l113_Playfield, LM_L_l113_UnderPF, LM_L_l12_Bumper1_Ring, LM_L_l12_Bumper1_Socket, LM_L_l12_Bumper2_Ring, LM_L_l12_Bumper2_Socket, LM_L_l12_Bumper3_Ring, LM_L_l12_Bumper3_Socket, LM_L_l12_Gate3p, LM_L_l12_Gate4p, LM_L_l12_Kicker1Hole, LM_L_l12_Kicker2Hole, LM_L_l12_OuterPrimBlack_DT, LM_L_l12_Overlay, LM_L_l12_Parts, LM_L_l12_PartsST, LM_L_l12_Playfield, LM_L_l12_Ramp001, LM_L_l12_SpinnerPrim, LM_L_l12_UnderPF, LM_L_l12_sw12, LM_L_l12_sw13, _
  LM_L_l12_sw14, LM_L_l12_sw17p, LM_L_l12_sw33p, LM_L_l12_sw34p, LM_L_l12_sw43p, LM_L_l12_sw46p, LM_L_l12_sw47p, LM_L_l12_sw48p, LM_L_l14_UnderPF, LM_L_l15_OuterPrimBlack_DT, LM_L_l15_UnderPF, LM_L_l17_Parts, LM_L_l17_PlasticsSolid, LM_L_l17_UnderPF, LM_L_l18_UnderPF, LM_L_l19_UnderPF, LM_L_l2_UnderPF, LM_L_l20_UnderPF, LM_L_l21_UnderPF, LM_L_l22_UnderPF, LM_L_l23_UnderPF, LM_L_l24_UnderPF, LM_L_l25_Parts, LM_L_l25_UnderPF, LM_L_l26_Parts, LM_L_l26_UnderPF, LM_L_l28_Bumper1_Ring, LM_L_l28_Bumper1_Socket, LM_L_l28_Bumper2_Ring, LM_L_l28_Bumper2_Socket, LM_L_l28_Bumper3_Ring, LM_L_l28_Bumper3_Socket, LM_L_l28_Gate4p, LM_L_l28_Kicker2Hole, LM_L_l28_OuterPrimBlack_DT, LM_L_l28_Overlay, LM_L_l28_Parts, LM_L_l28_PartsST, LM_L_l28_PlasticsSolid, LM_L_l28_Playfield, LM_L_l28_Ramp001, LM_L_l28_Ramp002, LM_L_l28_RflipmeshUp, LM_L_l28_RflipmeshUpU, LM_L_l28_SpinnerPrim, LM_L_l28_UnderPF, LM_L_l28_lockdownbar, LM_L_l28_sw17p, LM_L_l28_sw28p, LM_L_l28_sw32p, LM_L_l28_sw33p, LM_L_l28_sw34p, LM_L_l28_sw35p, LM_L_l28_sw46p, _
  LM_L_l28_sw48p, LM_L_l3_UnderPF, LM_L_l30_UnderPF, LM_L_l31_RflipmeshUp, LM_L_l31_RflipmeshUpU, LM_L_l31_UnderPF, LM_L_l33_LSling2, LM_L_l33_Parts, LM_L_l33_UnderPF, LM_L_l34_UnderPF, LM_L_l35_UnderPF, LM_L_l36_UnderPF, LM_L_l362_UnderPF, LM_L_l37_UnderPF, LM_L_l38_UnderPF, LM_L_l39_UnderPF, LM_L_l4_UnderPF, LM_L_l40_UnderPF, LM_L_l41_UnderPF, LM_L_l42_Parts, LM_L_l42_UnderPF, LM_L_l44_Bumper1_Ring, LM_L_l44_Bumper1_Socket, LM_L_l44_Bumper2_Ring, LM_L_l44_Bumper2_Socket, LM_L_l44_Bumper3_Ring, LM_L_l44_Bumper3_Socket, LM_L_l44_Gate4p, LM_L_l44_OuterPrimBlack_DT, LM_L_l44_Overlay, LM_L_l44_Parts, LM_L_l44_PartsST, LM_L_l44_PlasticsSolid, LM_L_l44_Playfield, LM_L_l44_Ramp001, LM_L_l44_Ramp002, LM_L_l44_SpinnerPrim, LM_L_l44_SpinnerRod, LM_L_l44_UnderPF, LM_L_l44_sw12, LM_L_l44_sw13, LM_L_l44_sw28p, LM_L_l44_sw30p, LM_L_l44_sw31p, LM_L_l44_sw32p, LM_L_l44_sw33p, LM_L_l44_sw34p, LM_L_l44_sw35p, LM_L_l46_UnderPF, LM_L_l49_OuterPrimBlack_DT, LM_L_l49_Parts, LM_L_l49_UnderPF, LM_L_l492_UnderPF, LM_L_l5_UnderPF, _
  LM_L_l50_UnderPF, LM_L_l51_UnderPF, LM_L_l52_UnderPF, LM_L_l53_UnderPF, LM_L_l54_UnderPF, LM_L_l55_UnderPF, LM_L_l56_UnderPF, LM_L_l57_UnderPF, LM_L_l58_UnderPF, LM_L_l59_Parts, LM_L_l59_UnderPF, LM_L_l6_UnderPF, LM_L_l60_Parts, LM_L_l60_SpinnerPrim, LM_L_l60_UnderPF, LM_L_l62_UnderPF, LM_L_l63_UnderPF, LM_L_l65_Overlay, LM_L_l65_Parts, LM_L_l65_PlasticsSolid, LM_L_l65_Playfield, LM_L_l65_UnderPF, LM_L_l66_Parts, LM_L_l66_Playfield, LM_L_l66_UnderPF, LM_L_l66_sw46p, LM_L_l66b_Gate4p, LM_L_l66b_Parts, LM_L_l66b_Playfield, LM_L_l66b_sw44p, LM_L_l66c_Overlay, LM_L_l66c_Parts, LM_L_l66c_PartsST, LM_L_l66c_Playfield, LM_L_l7_UnderPF, LM_L_l8_UnderPF, LM_L_l81_OuterPrimBlack_DT, LM_L_l81_Overlay, LM_L_l81_Parts, LM_L_l81_Playfield, LM_L_l81_Ramp002, LM_L_l81_UnderPF, LM_L_l82a_Bumper3_Ring, LM_L_l82a_Overlay, LM_L_l82a_Parts, LM_L_l82a_Playfield, LM_L_l82a_sw43p, LM_L_l9_UnderPF, LM_L_l97_Overlay, LM_L_l97_Parts, LM_L_l97_PartsST, LM_L_l97_Playfield, LM_L_l97_UnderPF, LM_L_l98a_Bumper2_Ring, LM_L_l98a_Parts, _
  LM_L_l98a_Playfield, LM_L_l98a_UnderPF)
' VLM  Arrays - End

Dim Reflection_Array: Reflection_Array=Array(LM_GI_Bumper1_Ring, LM_GI_Bumper1_Socket, LM_GI_Bumper2_Ring, LM_GI_Bumper2_Socket, LM_GI_Bumper3_Ring, LM_GI_Bumper3_Socket, LM_GI_Gate3p, LM_GI_Gate4p, LM_GI_Overlay, LM_GI_Parts, LM_GI_PartsST, LM_GI_Ramp002, LM_GI_RflipmeshUpU, LM_GI_SpinnerPrim, LM_GI_sw33p, LM_GI_sw34p, LM_GI_sw35p, _
LM_GI_sw42p, LM_GI_sw43p, LM_GI_sw46p, LM_GI_sw47p, LM_GI_sw48p, LM_GI_RT_gi_001_Gate2p, LM_GI_RT_gi_001_LSling, LM_GI_RT_gi_001_Lflipmesh1U, LM_GI_RT_gi_001_Parts, LM_GI_RT_gi_001_PlasticsSolid, LM_GI_RT_gi_002_Gate1p, LM_GI_RT_gi_002_Parts, LM_GI_RT_gi_002_PlasticsSolid, LM_GI_RT_gi_002_RSling, LM_GI_RT_gi_002_Rflipmesh1, LM_GI_RT_gi_002_Rflipmesh1U, LM_GI_RT_gi_003_Gate1p, _
LM_GI_RT_gi_003_Gate2p, LM_GI_RT_gi_003_LSling, LM_GI_RT_gi_003_Lflipmesh1, LM_GI_RT_gi_003_Lflipmesh1U, LM_GI_RT_gi_003_Parts, LM_GI_RT_gi_003_PlasticsSolid, LM_GI_RT_gi_003_Rflipmesh1U, LM_GI_RT_gi_003_Sling2, LM_GI_RT_gi_004_Bumper3_Ring, LM_GI_RT_gi_004_LSling, LM_GI_RT_gi_004_Parts, LM_GI_RT_gi_004_PlasticsSolid, LM_GI_RT_gi_004_sw27p, LM_GI_RT_gi_004_sw28p, LM_GI_RT_gi_004_sw29p, LM_GI_RT_gi_004_sw30p, LM_GI_RT_gi_005_Bumper1_Ring, _
LM_GI_RT_gi_005_Bumper3_Ring, LM_GI_RT_gi_005_Bumper3_Socket, LM_GI_RT_gi_005_Parts, LM_GI_RT_gi_005_PlasticsSolid, LM_GI_RT_gi_005_sw27p, LM_GI_RT_gi_005_sw28p, LM_GI_RT_gi_005_sw29p, LM_GI_RT_gi_005_sw30p, LM_GI_RT_gi_005_sw31p, LM_GI_RT_gi_005_sw32p, LM_GI_RT_gi_005_sw33p, LM_GI_RT_gi_005_sw34p, LM_GI_RT_gi_005_sw35p, LM_GI_RT_gi_006_Bumper1_Ring, LM_GI_RT_gi_006_Bumper1_Socket, LM_GI_RT_gi_006_Bumper3_Ring, LM_GI_RT_gi_006_Bumper3_Socket, _
LM_GI_RT_gi_006_Parts, LM_GI_RT_gi_006_PlasticsSolid, LM_GI_RT_gi_006_sw17p, LM_GI_RT_gi_006_sw27p, LM_GI_RT_gi_006_sw28p, LM_GI_RT_gi_006_sw29p, LM_GI_RT_gi_006_sw30p, LM_GI_RT_gi_006_sw31p, LM_GI_RT_gi_006_sw32p, LM_GI_RT_gi_006_sw33p, LM_GI_RT_gi_006_sw34p, LM_GI_RT_gi_006_sw35p, LM_GI_RT_gi_007_Bumper1_Ring, LM_GI_RT_gi_007_Bumper1_Socket, LM_GI_RT_gi_007_Bumper2_Ring, LM_GI_RT_gi_007_Bumper3_Ring, LM_GI_RT_gi_007_Bumper3_Socket, _
LM_GI_RT_gi_007_Overlay, LM_GI_RT_gi_007_Parts, LM_GI_RT_gi_007_PlasticsSolid, LM_GI_RT_gi_007_SpinnerPrim, LM_GI_RT_gi_007_sw31p, LM_GI_RT_gi_007_sw32p, LM_GI_RT_gi_007_sw33p, LM_GI_RT_gi_007_sw34p, LM_GI_RT_gi_007_sw35p, LM_GI_RT_gi_017_Bumper1_Ring, LM_GI_RT_gi_017_Bumper1_Socket, LM_GI_RT_gi_017_Bumper3_Socket, LM_GI_RT_gi_017_Overlay, LM_GI_RT_gi_017_Parts, LM_GI_RT_gi_017_PartsST, LM_GI_RT_gi_017_SpinnerPrim, _
LM_GI_RT_gi_017_SpinnerRod, LM_GI_RT_gi_023_Bumper1_Ring, LM_GI_RT_gi_023_Bumper1_Socket, LM_GI_RT_gi_023_Bumper2_Ring, LM_GI_RT_gi_023_Bumper2_Socket, LM_GI_RT_gi_023_Bumper3_Socket, LM_GI_RT_gi_023_Gate3p, LM_GI_RT_gi_023_Gate4p, LM_GI_RT_gi_023_Overlay, LM_GI_RT_gi_023_Parts, LM_GI_RT_gi_023_PartsST, LM_GI_RT_gi_023_Ramp001, LM_GI_RT_gi_025_Bumper1_Socket, LM_GI_RT_gi_025_Bumper2_Ring, LM_GI_RT_gi_025_Bumper2_Socket, LM_GI_RT_gi_025_Bumper3_Ring, LM_GI_RT_gi_025_Bumper3_Socket, _
LM_GI_RT_gi_025_Overlay, LM_GI_RT_gi_025_Parts, LM_GI_RT_gi_025_PartsST, LM_GI_RT_gi_025_Ramp002, LM_GI_RT_gi_025_RflipmeshUp, LM_GI_RT_gi_025_RflipmeshUpU, LM_GI_RT_gi_025_sw17p, LM_GI_RT_gi_025_sw33p, LM_GI_RT_gi_025_sw34p, LM_GI_RT_gi_025_sw35p, LM_GI_RT_gi_025_sw46p, LM_GI_RT_gi_025_sw47p, LM_GI_RT_gi_025_sw48p, LM_GI_RT_gi_026_Parts, LM_GI_RT_gi_026_PartsST, LM_GI_RT_gi_026_PlasticsSolid, LM_GI_RT_gi_026_RflipmeshUp, _
LM_GI_RT_gi_026_RflipmeshUpU, LM_GI_RT_gi_026_sw17p, LM_GI_RT_gi_026_sw33p, LM_GI_RT_gi_027_Gate1p, LM_GI_RT_gi_027_Gate2p, LM_GI_RT_gi_027_LSling, LM_GI_RT_gi_027_Lflipmesh1, LM_GI_RT_gi_027_Lflipmesh1U, LM_GI_RT_gi_027_Parts, LM_GI_RT_gi_027_PlasticsSolid, LM_GI_RT_gi_027_RSling, LM_GI_RT_gi_027_Rflipmesh1, LM_GI_RT_gi_027_Rflipmesh1U, LM_GI_RT_gi_028_Gate1p, LM_GI_RT_gi_028_LSling, LM_GI_RT_gi_028_Lflipmesh1, LM_GI_RT_gi_028_Lflipmesh1U, _
LM_GI_RT_gi_028_Parts, LM_GI_RT_gi_028_PlasticsSolid, LM_GI_RT_gi_028_RSling, LM_GI_RT_gi_028_Rflipmesh1, LM_GI_RT_gi_028_Rflipmesh1U, LM_GI_RT_gi_029_Gate1p, LM_GI_RT_gi_029_Lflipmesh1U, LM_GI_RT_gi_029_Parts, LM_GI_RT_gi_029_PlasticsSolid, LM_GI_RT_gi_029_RSling, LM_GI_RT_gi_029_Rflipmesh1, LM_GI_RT_gi_029_Rflipmesh1U, LM_GI_RT_gi_029_Sling1, LM_GI_RT_l66a_Bumper1_Ring, LM_GI_RT_l66a_Bumper1_Socket, LM_GI_RT_l66a_Bumper3_Ring, LM_GI_RT_l66a_Bumper3_Socket, _
LM_GI_RT_l66a_Overlay, LM_GI_RT_l66a_Parts, LM_GI_RT_l66a_PartsST, LM_GI_RT_l66a_SpinnerPrim, LM_GI_RT_l66a_SpinnerRod, LM_GI_RT_l66a_sw34p, LM_GI_RT_l66a_sw35p, LM_GI_RT_l82_Bumper1_Ring, LM_GI_RT_l82_Bumper1_Socket, LM_GI_RT_l82_Overlay, LM_GI_RT_l82_Parts, LM_GI_RT_l82_PartsST, LM_GI_RT_l82_SpinnerPrim, LM_GI_RT_l82_sw35p, LM_GI_RT_l98_Bumper1_Socket, LM_GI_RT_l98_Gate4p, LM_GI_RT_l98_Overlay, _
LM_GI_RT_l98_Parts, LM_GI_RT_l98_PartsST, LM_GI_RT_l98_SpinnerPrim, LM_L_Bmp1HL_Bumper1_Ring, LM_L_Bmp1HL_Bumper1_Socket, LM_L_Bmp1HL_Bumper2_Socket, LM_L_Bmp1HL_Bumper3_Ring, LM_L_Bmp1HL_Bumper3_Socket, LM_L_Bmp1HL_Overlay, LM_L_Bmp1HL_Parts, LM_L_Bmp1HL_PartsST, LM_L_Bmp1HL_Ramp002, LM_L_Bmp1HL_SpinnerPrim, LM_L_Bmp1HL_SpinnerRod, LM_L_Bmp1HL_sw32p, LM_L_Bmp1HL_sw34p, LM_L_Bmp1HL_sw35p, _
LM_L_Bmp2HL_Bumper1_Ring, LM_L_Bmp2HL_Bumper1_Socket, LM_L_Bmp2HL_Bumper2_Ring, LM_L_Bmp2HL_Bumper2_Socket, LM_L_Bmp2HL_Bumper3_Ring, LM_L_Bmp2HL_Bumper3_Socket, LM_L_Bmp2HL_Gate4p, LM_L_Bmp2HL_Overlay, LM_L_Bmp2HL_Parts, LM_L_Bmp2HL_PartsST, LM_L_Bmp3HL_Bumper1_Ring, LM_L_Bmp3HL_Bumper1_Socket, LM_L_Bmp3HL_Bumper2_Ring, LM_L_Bmp3HL_Bumper2_Socket, LM_L_Bmp3HL_Bumper3_Ring, LM_L_Bmp3HL_Bumper3_Socket, _
LM_L_Bmp3HL_Overlay, LM_L_Bmp3HL_Parts, LM_L_Bmp3HL_PartsST, LM_L_Bmp3HL_Ramp001, LM_L_Bmp3HL_Ramp002, LM_L_Bmp3HL_SpinnerPrim, LM_L_Bmp3HL_lockdownbar, LM_L_Bmp3HL_sw17p, LM_L_Bmp3HL_sw33p, LM_L_Bmp3HL_sw34p, LM_L_Bmp3HL_sw35p, LM_L_l1_Parts, LM_L_l1_PlasticsSolid, LM_L_l10_Parts, LM_L_l113_Gate4p, LM_L_l113_Overlay, LM_L_l113_Parts, _
LM_L_l12_Bumper1_Ring, LM_L_l12_Bumper1_Socket, LM_L_l12_Bumper2_Ring, LM_L_l12_Bumper2_Socket, LM_L_l12_Bumper3_Ring, LM_L_l12_Bumper3_Socket, LM_L_l12_Gate3p, LM_L_l12_Gate4p, LM_L_l12_Overlay, LM_L_l12_Parts, LM_L_l12_PartsST, LM_L_l12_Ramp001, LM_L_l12_SpinnerPrim, LM_L_l12_sw17p, LM_L_l12_sw33p, LM_L_l12_sw34p, _
LM_L_l12_sw43p, LM_L_l12_sw46p, LM_L_l12_sw47p, LM_L_l12_sw48p, LM_L_l17_Parts, LM_L_l17_PlasticsSolid, LM_L_l25_Parts, LM_L_l26_Parts, LM_L_l28_Bumper1_Ring, LM_L_l28_Bumper1_Socket, LM_L_l28_Bumper2_Ring, LM_L_l28_Bumper2_Socket, LM_L_l28_Bumper3_Ring, LM_L_l28_Bumper3_Socket, LM_L_l28_Gate4p, LM_L_l28_Overlay, LM_L_l28_Parts, _
LM_L_l28_PartsST, LM_L_l28_PlasticsSolid, LM_L_l28_Ramp001, LM_L_l28_Ramp002, LM_L_l28_RflipmeshUp, LM_L_l28_RflipmeshUpU, LM_L_l28_SpinnerPrim, LM_L_l28_lockdownbar, LM_L_l28_sw17p, LM_L_l28_sw28p, LM_L_l28_sw32p, LM_L_l28_sw33p, LM_L_l28_sw34p, LM_L_l28_sw35p, LM_L_l28_sw46p, LM_L_l28_sw48p, LM_L_l31_RflipmeshUp, _
LM_L_l31_RflipmeshUpU, LM_L_l33_Parts, LM_L_l42_Parts, LM_L_l44_Bumper1_Ring, LM_L_l44_Bumper1_Socket, LM_L_l44_Bumper2_Ring, LM_L_l44_Bumper2_Socket, LM_L_l44_Bumper3_Ring, LM_L_l44_Bumper3_Socket, LM_L_l44_Gate4p, LM_L_l44_Overlay, LM_L_l44_Parts, LM_L_l44_PartsST, LM_L_l44_PlasticsSolid, LM_L_l44_Ramp001, LM_L_l44_Ramp002, LM_L_l44_SpinnerPrim, _
LM_L_l44_SpinnerRod, LM_L_l44_sw28p, LM_L_l44_sw30p, LM_L_l44_sw31p, LM_L_l44_sw32p, LM_L_l44_sw33p, LM_L_l44_sw34p, LM_L_l44_sw35p, LM_L_l49_Parts, LM_L_l59_Parts, LM_L_l60_Parts, LM_L_l60_SpinnerPrim, LM_L_l65_Overlay, LM_L_l65_Parts, LM_L_l65_PlasticsSolid, LM_L_l66_Parts, LM_L_l66_sw46p, _
LM_L_l66b_Gate4p, LM_L_l66b_Parts, LM_L_l66b_sw44p, LM_L_l66c_Overlay, LM_L_l66c_Parts, LM_L_l66c_PartsST, LM_L_l81_Overlay, LM_L_l81_Parts, LM_L_l81_Ramp002, LM_L_l82a_Bumper3_Ring, LM_L_l82a_Overlay, LM_L_l82a_Parts, LM_L_l82a_sw43p, LM_L_l97_Overlay, LM_L_l97_Parts, LM_L_l97_PartsST, LM_L_l98a_Bumper2_Ring, _
LM_L_l98a_Parts)

'************************
' Bumper bake maps to follow VP Bumpers
'************************

'**************
' Bumper Animations
'**************

Sub Bumper1_Animate
  Dim z, BL
  z = Bumper1.CurrentRingOffset
  For Each BL in BP_Bumper1_Ring : BL.transz = z: Next
End Sub

Sub Bumper2_Animate
  Dim z, BP
  z = Bumper2.CurrentRingOffset
  For Each BP in BP_Bumper2_Ring : BP.transz = z: Next
End Sub

Sub Bumper3_Animate
  Dim z, BP
  z = Bumper3.CurrentRingOffset
  For Each BP in BP_Bumper3_Ring : BP.transz = z: Next
End Sub

Dim Bumpers : Bumpers = Array(Bumper1, Bumper2, Bumper3)



'************************
' Flipper bake maps and shadows to follow the VP flippers
'************************

Sub LeftFlipper_Animate
  Dim a: a = LeftFlipper.CurrentAngle       ' store flipper angle in a
  Dim bias:  bias = 0               ' store bias 0 for rest 1 for up position
  Dim max_angle, min_angle            ' min and max angles from the flipper
  max_angle = 122                 ' flipper down angle
  min_angle = 69                  ' flipper up angle
  FlipperLSh.RotZ = a               ' set flipper shadow angle
  bias = (max_angle - a) / (max_angle - min_angle)' calculate bias based on a and min/max
  If bias >= 0.99 Then bias = 1         ' handle overshoot to set 1 as ceiling
  If bias <= 0.01 Then bias = 0         ' handle undershoor to set 0 as floor

  Dim BP :
  For Each BP in BP_Lflipmesh1U
    If BP.Name <> "BM_Lflipmesh1U" Then
      If bias = 0 Then              ' for 0 bias (flipper rest)
        BP.visible = False            ' turn off up position maps
      Else                    ' for anything else
        BP.visible = True           ' turn on up position maps
      End If
      BP.Opacity = 100 * bias           ' set up maps opacity based on bias
    End If
    BP.RotZ = a                 ' rotate the maps
  Next
  For Each BP in BP_Lflipmesh1
    If BP.Name <> "BM_Lflipmesh1" Then
      If bias = 1 Then              ' for 1 bias (fully engaged)
        BP.visible = False            ' turn off down position maps
      Else                    ' for anything else
        BP.visible = True           ' turn on up position maps
      End If
      BP.Opacity = 100 * (1-bias)         ' set down maps opacity based on bias
    End If
    BP.RotZ = a                 ' rotate the maps
  Next
  'debug.print ("Angle: " + CStr(a) + " bias: " + CStr(bias) + ", UpOpac:" + CStr (100 * bias) + ", DownOpac:" +  CStr (100 * (1-bias)))
End Sub

Sub RightFlipper_Animate
  Dim a: a = RightFlipper.CurrentAngle        ' store flipper angle in a
  Dim bias:  bias = 0               ' store bias 0 for rest 1 for up position
  Dim max_angle, min_angle            ' min and max angles from the flipper
  max_angle = -122                  ' flipper down angle
  min_angle = -69                 ' flipper up angle
  FlipperRSh.RotZ = a               ' set flipper shadow angle
  bias = (max_angle - a) / (max_angle - min_angle)' calculate bias based on a and min/max
  If bias >= 0.99 Then bias = 1         ' handle overshoot to set 1 as ceiling
  If bias <= 0.01 Then bias = 0         ' handle undershoor to set 0 as floor

  Dim BP :
  For Each BP in BP_Rflipmesh1U
    If BP.Name <> "BM_Rflipmesh1U" Then
      If bias = 0 Then              ' for 0 bias (flipper rest)
        BP.visible = False            ' turn off up position maps
      Else                    ' for anything else
        BP.visible = True           ' turn on up position maps
      End If
      BP.Opacity = 100 * bias           ' set up maps opacity based on bias
    End If
    BP.RotZ = a                 ' rotate the maps
  Next
  For Each BP in BP_Rflipmesh1
    If BP.Name <> "BM_Rflipmesh1" Then
      If bias = 1 Then              ' for 1 bias (fully engaged)
        BP.visible = False            ' turn off down position maps
      Else                    ' for anything else
        BP.visible = True           ' turn on up position maps
      End If
      BP.Opacity = 100 * (1-bias)         ' set down maps opacity based on bias
    End If
    BP.RotZ = a                 ' rotate the maps
  Next
  'debug.print ("Angle: " + CStr(a) + " bias: " + CStr(bias) + ", UpOpac:" + CStr (100 * bias) + ", DownOpac:" +  CStr (100 * (1-bias)))
End Sub

Sub RightFlipper1_Animate
  Dim a: a = RightFlipper1.CurrentAngle
  FlipperRShUp.RotZ = a
  Dim BP :

  For Each BP in BP_RflipmeshUp
    BP.RotZ = a
    BP.visible = a < 233
  Next
  For Each BP in BP_RflipmeshUpU
    BP.RotZ = a
    BP.visible = a >= 233
  Next
End Sub

'************************
' Spinner bake maps to follow the VP spinners
'************************

Sub Spinner_Animate
  Dim a: a = Spinner.CurrentAngle + 180
  Dim BP : For Each BP in BP_SpinnerPrim: BP.RotX = a: Next
  For Each BP in BP_SpinnerRod
    BP.TransX = - SpinnerRod.TransX
    BP.TransZ = SpinnerRod.TransZ - 9
  Next
  SpinnerRod.TransX = sin( (Spinner.CurrentAngle+180) * (2*PI/360)) * 12
  SpinnerRod.TransZ = sin( (Spinner.CurrentAngle- 90) * (2*PI/360)) * 10
End Sub



'************************
' Gates bake maps to follow the VP gates
'************************

Sub Gate1_Animate
  Dim a: a = Gate1.CurrentAngle
  Dim BP : For Each BP in BP_Gate1p: BP.RotX = a: Next
End Sub

Sub Gate2_Animate
  Dim a: a = Gate2.CurrentAngle
  Dim BP : For Each BP in BP_Gate2p: BP.RotX = a: Next
End Sub

Sub Gate3_Animate
  Dim a: a = Gate3.CurrentAngle
  Dim BP : For Each BP in BP_Gate3p: BP.RotX = a: Next
End Sub

Sub Gate4_Animate
  Dim a: a = Gate4.CurrentAngle
  Dim BP : For Each BP in BP_Gate4p: BP.RotX = -a: Next
End Sub

Sub Gate5_Animate
  Dim a: a = Gate5.CurrentAngle
  Dim BP : For Each BP in BP_Gate5_Wire: BP.RotY = -a: Next
End Sub

'************************
''''' Switch Animations
'************************

Sub sw21_Animate
  Dim z : z = sw21.CurrentAnimOffset
  Dim BP : For Each BP in BP_sw21 : BP.transz = z: Next
End Sub

Sub sw22_Animate
  Dim z : z = sw22.CurrentAnimOffset
  Dim BP : For Each BP in BP_sw22 : BP.transz = z: Next
End Sub


Sub sw23_Animate
  Dim z : z = sw23.CurrentAnimOffset
  Dim BP : For Each BP in BP_sw23 : BP.transz = z: Next
End Sub

Sub sw24_Animate
  Dim z : z = sw24.CurrentAnimOffset
  Dim BP : For Each BP in BP_sw24 : BP.transz = z: Next
End Sub

Sub sw12_Animate
  Dim z : z = sw12.CurrentAnimOffset
  Dim BP : For Each BP in BP_sw12 : BP.transz = z: Next
End Sub

Sub sw13_Animate
  Dim z : z = sw13.CurrentAnimOffset
  Dim BP : For Each BP in BP_sw13 : BP.transz = z: Next
End Sub

Sub sw14_Animate
  Dim z : z = sw14.CurrentAnimOffset
  Dim BP : For Each BP in BP_sw14 : BP.transz = z: Next
End Sub

'************************
' Standup Target bake maps to follow the VP Targets
'************************

Sub UpdateStandupTargets
  dim BP, ty

    ty = BM_sw17p.transy
  For each BP in BP_sw17p : BP.transy = ty: Next

End Sub


'************************
' Drop Target bake maps to follow the VP Targets
'************************

Sub UpdateDropTargets()
  Dim BP, tz, rx, ry

  'Center Targets
  tz = BM_sw33p.transz
  rx = BM_sw33p.rotx
  ry = BM_sw33p.roty
  For Each BP in BP_sw33p: BP.transz=tz: BP.rotx = rx: BP.roty = ry: Next

  tz = BM_sw34p.transz
  rx = BM_sw34p.rotx
  ry = BM_sw34p.roty
  For Each BP in BP_sw34p: BP.transz=tz: BP.rotx = rx: BP.roty = ry: Next

  tz = BM_sw35p.transz
  rx = BM_sw35p.rotx
  ry = BM_sw35p.roty
  For Each BP in BP_sw35p: BP.transz=tz: BP.rotx = rx: BP.roty = ry: Next

  'Left Targets
  tz = BM_sw27p.transz
  rx = BM_sw27p.rotx
  ry = BM_sw27p.roty
  For Each BP in BP_sw27p: BP.transz=tz: BP.rotx = rx: BP.roty = ry: Next

  tz = BM_sw28p.transz
  rx = BM_sw28p.rotx
  ry = BM_sw28p.roty
  For Each BP in BP_sw28p: BP.transz=tz: BP.rotx = rx: BP.roty = ry: Next

  tz = BM_sw29p.transz
  rx = BM_sw29p.rotx
  ry = BM_sw29p.roty
  For Each BP in BP_sw29p: BP.transz=tz: BP.rotx = rx: BP.roty = ry: Next

  tz = BM_sw30p.transz
  rx = BM_sw30p.rotx
  ry = BM_sw30p.roty
  For Each BP in BP_sw30p: BP.transz=tz: BP.rotx = rx: BP.roty = ry: Next

  tz = BM_sw31p.transz
  rx = BM_sw31p.rotx
  ry = BM_sw31p.roty
  For Each BP in BP_sw31p: BP.transz=tz: BP.rotx = rx: BP.roty = ry: Next

  tz = BM_sw32p.transz
  rx = BM_sw32p.rotx
  ry = BM_sw32p.roty
  For Each BP in BP_sw32p: BP.transz=tz: BP.rotx = rx: BP.roty = ry: Next

  'Top Targets
  tz = BM_sw42p.transz
  rx = BM_sw42p.rotx
  ry = BM_sw42p.roty
  For Each BP in BP_sw42p: BP.transz=tz: BP.rotx = rx: BP.roty = ry: Next

  tz = BM_sw43p.transz
  rx = BM_sw43p.rotx
  ry = BM_sw43p.roty
  For Each BP in BP_sw43p: BP.transz=tz: BP.rotx = rx: BP.roty = ry: Next

  tz = BM_sw44p.transz
  rx = BM_sw44p.rotx
  ry = BM_sw44p.roty
  For Each BP in BP_sw44p: BP.transz=tz: BP.rotx = rx: BP.roty = ry: Next

  'Right Targets
  tz = BM_sw46p.transz
  rx = BM_sw46p.rotx
  ry = BM_sw46p.roty
  For Each BP in BP_sw46p: BP.transz=tz: BP.rotx = rx: BP.roty = ry: Next

  tz = BM_sw47p.transz
  rx = BM_sw47p.rotx
  ry = BM_sw47p.roty
  For Each BP in BP_sw47p: BP.transz=tz: BP.rotx = rx: BP.roty = ry: Next

  tz = BM_sw48p.transz
  rx = BM_sw48p.rotx
  ry = BM_sw48p.roty
  For Each BP in BP_sw48p: BP.transz=tz: BP.rotx = rx: BP.roty = ry: Next
End Sub

' Disable visibility on extra sling BMs until needed
Dim BP
For Each BP in BP_LSling1 : BP.Visible = 0: Next
For Each BP in BP_LSling2 : BP.Visible = 0: Next
For Each BP in BP_RSling1 : BP.Visible = 0: Next
For Each BP in BP_RSling2 : BP.Visible = 0: Next

' Set up russian doll for flipper upper position LMs and hide up position BMs
For Each BP in BP_Lflipmesh1U
  BP.Size_X = 1.002
  BP.Size_Y = 1.002
  BP.Size_Z = 1.002
Next
For Each BP in BP_Rflipmesh1U
  BP.Size_X = 1.002
  BP.Size_Y = 1.002
  BP.Size_Z = 1.002
Next
BM_Lflipmesh1U.visible = False
BM_Rflipmesh1U.visible = False

' Correction for misrotation on Sling arm
For Each BP in BP_Sling1 : BP.RotZ = 75 : Next
For Each BP in BP_Sling2 : BP.RotZ = -75 : Next

' Correction for misrotation on Spinner parts
For Each BP in BP_SpinnerPrim : BP.ObjRotZ = 12.75 : Next
For Each BP in BP_SpinnerRod : BP.ObjRotZ = 102 : Next

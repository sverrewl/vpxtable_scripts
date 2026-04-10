'*********************************************************************************************************************************
' __      ______  _      _      ________     __
' \ \    / / __ \| |    | |    |  ____\ \   / /
'  \ \  / / |  | | |    | |    | |__   \ \_/ /
'   \ \/ /| |  | | |    | |    |  __|   \   /
'    \  / | |__| | |____| |____| |____   | |
'     \/   \____/|______|______|______|  |_|
'
'*********************************************************************************************************************************
'
'
' Volley (Gottlieb 1976) table by FrankEnstein
'
'-Built from scratch on a foundation of VPW Basic Table Template by Apophis
'-Utilizes some resources and techniques from Loserman's table
'-Big Thanks goes to Freezy and JSM for their asset donation. High rez playfield art and correct table layout would have been impossible without them.
'-Thanks to MCarter78 and Sixtoe for help with debugging and tweaking physics.
'-Thanks to all that provided testing and feedback: Studly_do_right, Tomate, Nestogrian, Nova, Apophis
'-Thanks to HauntFreaks for his unwavering commitment to backglass perfection

Option Explicit
Randomize

On Error Resume Next
ExecuteGlobal GetTextFile("controller.vbs")
If Err Then MsgBox "You need the controller.vbs In order To run this table, available In the vp10 package"
On Error GoTo 0



'*******************************************
' === TABLE OF CONTENTS  ===
'
' You can quickly jump to a section by searching the four letter tag (ZXXX)
'
'   ZCON: Constants and Global Variables
' ZTIM: Timers
' ZINI: Table Initialization and Exiting
' ZKEY: Key Press Handling
'   ZMAT: General Math Functions
'   ZANI: Misc Animations
' ZDRN: Drain, Trough, and Ball Release
' ZFLP: Flippers
' ZBMP: Bumpers
' ZKIC: Kickers
' ZDRT: Drop Targets
' ZSLG: Slingshots
' ZNFF: Flipepr Corrections and Tricks
'   ZDMP: Rubber Dampeners
'   ZBOU: Target Bouncer
' ZSSC: Slingshot Corrections
' ZBRL: Ball Rolling and Drop Sunds
'   ZTRG: Triggers for Switches
'   ZVLM: Bake maps to follow physical movement
'   ZFLE: Fleep Mechanical SoundSaucerKick
'*******************************************

' CHANGE LOG
' 1.0 Work was done many a late at nights, I don't remember the details.


'*******************************************
'  ZCON: Constants and Global Variables
'*******************************************

Const BallSize = 50        'Ball diameter in VPX units; must be 50
Const BallMass = 1         'Ball mass must be 1
Const tnob = 5           'Total number of balls the table can hold
Const lob = 0            'Locked balls
Const cGameName = "volley_1976"  'The unique alphanumeric name for this table

Const VolumeDial = 0.8            ' Overall Mechanical sound effect volume. Recommended values should be no greater than 1.
Const BallRollVolume = 0.5    ' Level of ball rolling volume. Value between 0 and 1
Const RampRollVolume = 0.5    ' Level of ramp rolling volume. Value between 0 and 1

Dim tablewidth: tablewidth = Table1.width
Dim tableheight: tableheight = Table1.height
Const UsingROM = False

'----- Ball Shadow Options -----
Const ShadowBalls     = 1    ' Number of ball shadow objects added to the table (BallShadow0..N-1)
Const AmbientBSFactor = 1    ' 0 to 1, higher is darker
Const AmbientMovement = 1    ' 1 to 4, higher = more movement as ball moves left/right
Const offsetX         = 0    ' Shadow offset X from ball position
Const offsetY         = 5    ' Shadow offset Y from ball position

' Game state variable
Dim score: score = 0
Dim tilt: tilt = False
Dim tilt_ends_game: tilt_ends_game = True
Dim TiltSens: TiltSens = 0
Dim credit: credit = 0
Dim sreels(1)
Dim blue_bank_complete: blue_bank_complete = False
Dim yellow_bank_complete: yellow_bank_complete = False
Dim green_bank_complete: green_bank_complete = False
Dim green_drops_down: green_drops_down = 0
Dim yellow_drops_down: yellow_drops_down = 0
Dim blue_drops_down: blue_drops_down = 0
Dim dtResetStep: dtResetStep = 0
Dim ballinplay: ballinplay = 0       ' current ball (1-5)
Dim state: state = False             ' True = game in progress
Const BallsPerGame = 5                      ' balls per game (TEMP: 1 for testing, restore to 5)
Dim matchnumb                               ' last drawn match number

' B2S indicator IDs - match these to your .directb2s file element IDs
Const B2S_BallInPlay = 32
Const B2S_Tilt       = 33
Const B2S_GameOver   = 35

'*******************************************
'  High Score
'*******************************************
Const HSFileName = "volley_gottlieb_1976.txt"
Dim ScoreChecker, CheckAllScores
Dim InProgress
Dim EnteringInitials  : EnteringInitials = 0
Dim InitialString     : InitialString = ""
Dim HSTimerCount      : HSTimerCount = 5
Dim AlphaString       : AlphaString = "ABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789_<"
Dim AlphaStringPos    : AlphaStringPos = 1
Dim HSNewHigh
Dim HSScore(5), HSName(5)
HSScore(1) = 25000 : HSName(1) = "FE"
HSScore(2) = 20000 : HSName(2) = "JAM"
HSScore(3) = 15000 : HSName(3) = "WLF"
HSScore(4) = 10000 : HSName(4) = "JBM"
HSScore(5) = 5000 : HSName(5) = "PEZ"

Const base_bank_points = 500
Const lit_bank_points = 5000
Const base_lane_points = 500
Const lit_lane_points = 5000
Const drain_points = 10
Const base_bumper_points = 10
Const lit_bumper_points = 100
Const lit_Y_bumper_points = 1000
Const sw_points = 10
Const sling_points = 100

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
  DoDTAnim
  UpdateDropTargets
  RollingUpdate       'update rolling sounds
  DynamicBSUpdate     'update ball shadow
End Sub

'The CorTimer interval should be 10. It's sole purpose is to update the Cor calculations
CorTimer.Interval = 10
Sub CorTimer_Timer(): Cor.Update: End Sub



'*******************************************
' ZINI: Table Initialization and Exiting
'*******************************************

LoadCoreFiles
Sub LoadCoreFiles
  On Error Resume Next
  ExecuteGlobal GetTextFile("core.vbs") 'TODO: drop-in replacement for vpmTimer (maybe vpwQueueManager) and cvpmDictionary (Scripting.Dictionary) to remove core.vbs dependency
  If Err Then MsgBox "Can't open core.vbs"
  On Error GoTo 0
End Sub

Dim gBOT
Dim DTShadows
Dim objBallShadow(5)
Dim objBallShadowA(5)

Sub Table1_Init
  Dim i

  Drain.CreateBall

  'Captive ball
' Kicker2.CreateBall
' Kicker2.kick 180, 7
' Kicker2.Enabled = 0
  InitRubberState
  DTShadows = Array(dts_b1, dts_b2, dts_b3, dts_b4, dts_b5, dts_g1, dts_g2, dts_g3, dts_g4, dts_g5, dts_y1, dts_y2, dts_y3, dts_y4, dts_y5)
  DynamicBSInit
  GiOFF

  LoadEM
  loadhs
  If PostItNoteEnabled Then
    HighScoreTimer.interval = 2000
    HighScoreTimer.enabled = True
  End If
  If B2SOn Then
    Controller.B2SSetScorePlayer 1, score
    Controller.B2SSetScoreRolloverPlayer1 0
    Controller.B2SSetBallInPlay B2S_BallInPlay, 0
    Controller.B2SSetGameOver B2S_GameOver, 1
    Controller.B2SSetTilt B2S_Tilt, 0
    Controller.B2SSetCredits credit
  End If
  If credit > 0 Then DOF 123, DOFPulse

End Sub

Sub Table1_Exit
  savehs
  If B2SOn Then Controller.Stop
End Sub

'*******************************************
'  ZOPT: User Options
'*******************************************

Dim LightLevel : LightLevel = 0.25        ' Level of room lighting (0 to 1), where 0 is dark and 100 is brightest
Dim ColorLUT : ColorLUT = 1           ' Color desaturation LUTs: 1 to 11, where 1 is normal and 11 is black'n'white


' Called when options are tweaked by the player.
' - 0: game has started, good time to load options and adjust accordingly
' - 1: an option has changed
' - 2: options have been reseted
' - 3: player closed the tweak UI, good time to update staticly prerendered parts
' Table1.Option arguments are:
' - option name, minimum value, maximum value, step between valid values, default value, unit (0=None, 1=Percent), an optional arry of literal strings
Dim dspTriggered : dspTriggered = False
Dim PostItNoteEnabled : PostItNoteEnabled = True
Dim Slope

Sub Table1_OptionEvent(ByVal eventId)
   If eventId = 1 And Not dspTriggered Then dspTriggered = True : DisableStaticPreRendering = True : End If

  ' Color Saturation
   ColorLUT = Table1.Option("Color Saturation", 0, 11, 1, 0, 0, _
    Array("Normal", "Vibrant", "Desaturated 10%", "Desaturated 20%", "Desaturated 30%", "Desaturated 40%", "Desaturated 50%", _
       "Desaturated 60%", "Desaturated 70%", "Desaturated 80%", "Desaturated 90%", "Black 'n White"))
  if ColorLUT = 0 Then Table1.ColorGradeImage = "colorgradelut256x16_1to1"
  if ColorLUT = 1 Then Table1.ColorGradeImage = "ColorGradeLUT256x16_1to1gam0.70vibr50sat100"
  if ColorLUT = 2 Then Table1.ColorGradeImage = "ColorGradeLUT256x16_1to1gam0.70vibr50sat90"
  if ColorLUT = 3 Then Table1.ColorGradeImage = "ColorGradeLUT256x16_1to1gam0.70vibr50sat80"
  if ColorLUT = 4 Then Table1.ColorGradeImage = "ColorGradeLUT256x16_1to1gam0.70vibr50sat70"
  if ColorLUT = 5 Then Table1.ColorGradeImage = "ColorGradeLUT256x16_1to1gam0.70vibr50sat60"
  if ColorLUT = 6 Then Table1.ColorGradeImage = "ColorGradeLUT256x16_1to1gam0.70vibr50sat50"
  if ColorLUT = 7 Then Table1.ColorGradeImage = "ColorGradeLUT256x16_1to1gam0.70vibr50sat40"
  if ColorLUT = 8 Then Table1.ColorGradeImage = "ColorGradeLUT256x16_1to1gam0.70vibr50sat30"
  if ColorLUT = 9 Then Table1.ColorGradeImage = "ColorGradeLUT256x16_1to1gam0.70vibr50sat20"
  if ColorLUT = 10 Then Table1.ColorGradeImage = "ColorGradeLUT256x16_1to1gam0.70vibr50sat10"
  if ColorLUT = 11 Then Table1.ColorGradeImage = "ColorGradeLUT256x16_1to1gam0.70vibr50sat00"

   ' Sound volumes
'   VolumeDial = Table1.Option("Mech Volume", 0, 1, 0.01, 0.8, 1)
'   BallRollVolume = Table1.Option("Ball Roll Volume", 0, 1, 0.01, 0.5, 1)

  ' Room brightness
' LightLevel = Table1.Option("Table Brightness (Ambient Light Level)", 0, 1, 0.01, .5, 1)
  LightLevel = NightDay/100
  SetRoomBrightness LightLevel   'Uncomment this line for lightmapped tables.

  Dim v
  Dim BP
   ' Toggle Rails
  v = Table1.Option("Show Rails", 0, 1, 1, 1, 0, Array("Hide", "Show"))
  For Each BP in BP_RailL : BP.Visible = v: Next
  For Each BP in BP_RailR: BP.Visible = v: Next

   ' Toggle Lockdown Bar
  v = Table1.Option("Show Lockbar", 0, 1, 1, 1, 0, Array("Hide", "Show"))
  For Each BP in BP_lockdownbar : BP.Visible = v: Next

  ' Playfield Reflections
  v = Table1.Option("Playfield Reflections", 0, 1, 1, 1, 0, Array("Clean", "Rough"))
  Select Case v
    Case 0: playfield_mesh.ReflectionProbe = "Playfield Reflections": BM_Playfield.ReflectionProbe = "Playfield Reflections"
    Case 1: playfield_mesh.ReflectionProbe = "Playfield Reflections Rough": BM_Playfield.ReflectionProbe = "Playfield Reflections Rough"
  End Select

  ' Reflections Intensity
  'v = Table1.Option("Reflections Intensity", 0.05, 0.15, 0.1, 5.5, 0)
  'playfield_mesh.Strength = v

   ' Toggle Reflections Scope
  v = Table1.Option("Playfield Reflections Scope", 0, 2, 1, 2, 0, Array("None", "Static only", "Dynamic"))
  ReflectionToggle(v)

  'Set slope
  Slope = Table1.Option("Playfield Angle (degrees)", 4, 6, 0.1, 5.5, 0)
  Table1.SlopeMin = Slope
  Table1.SlopeMax = Slope

   If eventId = 3 And dspTriggered Then dspTriggered = False : DisableStaticPreRendering = False : End If
End Sub

Function ReflectionToggle(state)
  Dim BP, v1, v2
  Select Case state
    Case 0: v1 = false: v2 = false  'No reflections
    Case 1: v1 = true: v2 = false   'Only reflect Bakemaps
    Case 2: v1 = true:  v2 = true   'Reflect Bakemaps and Lightmaps
  End Select

  For Each BP in BG_Bakemap
    BP.ReflectionEnabled = v1
  Next

  For Each BP in BG_Lightmap
    BP.ReflectionEnabled = v2
  Next

  BM_Parts.ReflectionEnabled = v2
  BM_Playfield.ReflectionEnabled = False
  playfield_mesh.ReflectionEnabled= False
End Function

'****************************
'   ZRBR: Room BrBG_Bakemapightness
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



'*******************************************
' ZKEY: Key Press Handling
'*******************************************

Sub Table1_KeyDown(ByVal keycode)

  If EnteringInitials Then
    CollectInitials keycode
    Exit Sub
  End If

  If keycode = LeftFlipperKey And state Then
    SolLFlipper True  'This would be called by the solenoid callbacks if using a ROM
  End If

  If keycode = RightFlipperKey And state Then
    SolRFlipper True  'This would be called by the solenoid callbacks if using a ROM
  End If

  If keycode = PlungerKey Then
    Plunger.Pullback
    SoundPlungerPull
  End If

  If keycode = LeftTiltKey Then
    Nudge 90, 1.5
    SoundNudgeLeft
    CheckTilt
  End If
  If keycode = RightTiltKey Then
    Nudge 270, 1.5
    SoundNudgeRight
    CheckTilt
  End If
  If keycode = CenterTiltKey Then
    Nudge 0, 2
    SoundNudgeCenter
    CheckTilt
  End If
  If keycode = MechanicalTilt Then
    DoTilt
  End If

  If keycode = StartGameKey Then
    SoundStartButton
    If Not state And credit > 0 Then
      credit = credit - 1
      If B2SOn Then Controller.B2SSetCredits credit
      If credit < 1 Then DOF 123, DOFOff
      GiON
      Playsound "initialize"
      newgame.enabled = True
    End If
  End If

  '   If keycode = keyInsertCoin1 or keycode = keyInsertCoin2 or keycode = keyInsertCoin3 or keycode = keyInsertCoin4 Then 'Use this for ROM based games
  If keycode = AddCreditKey Or keycode = AddCreditKey2 Then
    If credit < 9 Then credit = credit + 1
    If B2SOn Then Controller.B2SSetCredits credit
    DOF 123, DOFOn
    Select Case Int(Rnd * 3)
      Case 0
        PlaySound ("Coin_In_1"), 0, CoinSoundLevel, 0, 0.25
      Case 1
        PlaySound ("Coin_In_2"), 0, CoinSoundLevel, 0, 0.25
      Case 2
        PlaySound ("Coin_In_3"), 0, CoinSoundLevel, 0, 0.25
    End Select
  End If

End Sub



Sub Table1_KeyUp(ByVal keycode)

  If KeyCode = PlungerKey Then
    Plunger.Fire
    SoundPlungerReleaseBall()   'Plunger release sound when there is a ball in shooter lane
  End If

  If keycode = LeftFlipperKey And state Then
    SolLFlipper False   'This would be called by the solenoid callbacks if using a ROM
  End If

  If keycode = RightFlipperKey And state Then
    SolRFlipper False   'This would be called by the solenoid callbacks if using a ROM
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




'******************************************************
'  ZANI: Misc Animations
'******************************************************

'Sub LeftFlipper_Animate
' dim a: a = LeftFlipper.CurrentAngle
' FlipperLSh.RotZ = a
' LFLogo.RotZ = a
' 'Add any left flipper related animations here
'End Sub
'
'Sub RightFlipper_Animate
' dim a: a = RightFlipper.CurrentAngle
' FlipperRSh.RotZ = a
' RFlogo.RotZ = a
' 'Add any right flipper related animations here
'End Sub
'
'


'*******************************************
' ZDRN: Drain and Ball Release
'*******************************************

' DRAIN & RELEASE
Sub Drain_Hit
  DOF 113, DOFPulse
  Drain.TimerEnabled = True
  'ResetDropTargets
End Sub

Sub Drain_Timer
  RandomSoundDrain Drain
  Drain.TimerEnabled = False
  addscore drain_points
  If state Then nextball
End Sub


Sub newgame_timer
  score = 0
  ballinplay = 1
  tilt = False
  TiltSens = 0
  TiltTimer.Enabled = False
  state = True
  InProgress = True

  ' Top lane indicators: lit at game start, cleared when each lane is triggered
  I_lane_B.state = 1
  I_lane_Y.state = 1
  I_lane_G.state = 1

  ' Bank indicators: off until top lane is triggered
  I_Bank_B.state = 0
  I_Bank_G.state = 0
  I_Bank_Y.state = 0

  ' Bumper lights: off until top lanes activated
  L_BMP_B.state = 0
  L_BMP_Y.state = 0
  L_BMP_G.state = 0

  ' Feature completion lights: off until bank is completed
  I_special_L.state = 0
  I_special_R.state = 0
  I_lane_top_C.state = 0

  ' Out lane indicators: Toggled by center bumper, only one can be lit at a time
  I_lane_out_L.state = 1
  I_lane_out_R.state = 0

  ' Raise all drop targets
  ResetAllDropTargets

  If B2SOn Then
    Controller.B2SSetScorePlayer 1, score
    Controller.B2SSetScoreRolloverPlayer1 0
    Controller.B2SSetBallInPlay B2S_BallInPlay, ballinplay
    Controller.B2SSetGameOver B2S_GameOver, 0
    Controller.B2SSetTilt B2S_Tilt, 0
    Controller.B2SSetMatch 0
  End If
  newball
  newgame.enabled = False
End Sub

Sub newball
  DOF 108, DOFPulse
  Drain.kick 57, 20
  RandomSoundBallRelease Drain
End Sub

'*******************************************
'  ZMCH: Match
'*******************************************

' Draw a random multiple of 10 (0-90). If it equals the last two digits
' of the player's score, award a free credit and fire the knocker.
Sub DoMatchWilmManB2S
  matchnumb = (Int(Rnd * 10)) '* 10

  If B2SOn Then
    If matchnumb = 0 Then
      Controller.B2SSetMatch 10'0
    Else
      Controller.B2SSetMatch matchnumb
    End If
    ' debug.print "Match number: " & matchnumb
    ' debug.print "Player score: " & (score mod 100)
    'if matchnumb=(score mod 100) then
    if matchnumb*10=(score mod 100) then
      AwardSpecial
    end if
  End if
End Sub

Sub DoMatch
  matchnumb = (Int(Rnd * 10)) * 10

  If B2SOn Then
    If matchnumb = 0 Then
      Controller.B2SSetMatch 100
    Else
      Controller.B2SSetMatch matchnumb
    End If
    ' debug.print "Match number: " & matchnumb
    ' debug.print "Player score: " & (score mod 100)
    if matchnumb=(score mod 100) then
      AwardSpecial
    end if
  End if
End Sub

sub checkmatch
  Dim tempmatch
  if TableTilted=true and TiltEndsGame=1 then
    exit sub
  end if
  tempmatch=Int(Rnd*10)
  Match=tempmatch*10
  MatchReel.SetValue(tempmatch+1)
  if Match=0 then
    MatchTextBox.text="Match: 00"
  else
    MatchTextBox.text="Match: " + FormatNumber(Match,0)
  end if

  If B2SOn Then
    If Match = 0 Then
      Controller.B2SSetMatch 100
    Else
      Controller.B2SSetMatch Match
    End If
  End if
  for i = 1 to Players
    if Match=(Score(i) mod 100) then
      AddSpecial
    end if
  next
end sub
'*******************************************
'  ZTLT: Tilt
'*******************************************

Sub CheckTilt
  If TiltTimer.Enabled Then
    TiltSens = TiltSens + 1
    If TiltSens >= 3 Then
      DoTilt
    End If
  Else
    TiltSens = 0
    TiltTimer.Enabled = True
  End If
End Sub

Sub TiltTimer_Timer()
  TiltTimer.Enabled = False
End Sub

Sub DoTilt
  If Not state Or tilt Then Exit Sub
  tilt = True
  TiltTimer.Enabled = False
  TiltSens = 0
  InProgress = False
  If B2SOn Then Controller.B2SSetTilt B2S_Tilt, 1
  LeftFlipper.RotateToStart
  RightFlipper.RotateToStart
  ' Debug.Print "TILT! Game over."
  ' Tilt always ends the entire game
  state = False
  TurnOffAllLights
  If B2SOn Then
    Controller.B2SSetGameOver B2S_GameOver, 1
    Controller.B2SSetBallInPlay B2S_BallInPlay, 0
    Controller.B2SSetScorePlayer 1, score
  End If
End Sub

Sub TurnOffAllLights
  ' GI off
  GiOFF
  ' Insert lights off
  I_lane_B.state = 0
  I_lane_Y.state = 0
  I_lane_G.state = 0
  I_Bank_B.state = 0
  I_Bank_G.state = 0
  I_Bank_Y.state = 0
  I_special_L.state = 0
  I_special_R.state = 0
  I_lane_top_C.state = 0
  I_lane_out_L.state = 0
  I_lane_out_R.state = 0
  I_lane_mid_L.state = 0
  I_lane_mid_R.state = 0
  ' Bumper lights off
  L_BMP_B.state = 0
  L_BMP_Y.state = 0
  L_BMP_G.state = 0
End Sub

Sub nextball
  If ballinplay >= BallsPerGame Then
    ' All balls played - game over
    state = False
    InProgress = False
    If B2SOn Then
      Controller.B2SSetGameOver B2S_GameOver, 1
      Controller.B2SSetBallInPlay B2S_BallInPlay, 0
      Controller.B2SSetScorePlayer 1, score
    End If
    ' Flush any pending score-motor ticks so DoMatch sees the final reel value
    AddScore10Timer.Enabled = False
    AddScore100Timer.Enabled = False
    AddScore1000Timer.Enabled = False
    score = score + (Add10 * 10) + (Add100 * 100) + (Add1000 * 1000)
    Add10 = 0 : Add100 = 0 : Add1000 = 0
    If B2SOn Then
      If score >= 100000 Then
        Controller.B2SSetScoreRolloverPlayer1 1
        Controller.B2SSetScorePlayer 1, score Mod 100000
      Else
        Controller.B2SSetScorePlayer 1, score
      End If
    End If
    ' Debug.Print "GAME OVER  Final score: " & score
    DoMatch
    CheckHighScore
    savehs
  Else
    ballinplay = ballinplay + 1
    If B2SOn Then Controller.B2SSetBallInPlay B2S_BallInPlay, ballinplay
    newball
  End If
End Sub



'*******************************************
' ZFLP: Flippers
'*******************************************

Const ReflipAngle = 20

' Flipper Solenoid Callbacks (these subs mimics how you would handle flippers in ROM based tables)
Sub SolLFlipper(Enabled) 'Left flipper solenoid callback
  If Enabled Then
    DOF 101, DOFOn
    FlipperActivate LeftFlipper, LFPress
    LF.Fire

    If leftflipper.currentangle < leftflipper.endangle + ReflipAngle Then
      RandomSoundReflipUpLeft LeftFlipper
    Else
      SoundFlipperUpAttackLeft LeftFlipper
      RandomSoundFlipperUpLeft LeftFlipper
    End If
  Else
    DOF 101, DOFOff
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
    DOF 102, DOFOn
    FlipperActivate RightFlipper, RFPress
    RF.Fire

    If rightflipper.currentangle > rightflipper.endangle - ReflipAngle Then
      RandomSoundReflipUpRight RightFlipper
    Else
      SoundFlipperUpAttackRight RightFlipper
      RandomSoundFlipperUpRight RightFlipper
    End If
  Else
    DOF 102, DOFOff
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


'******************************************************
'     SKIRT ANIMATION FUNCTIONS
'******************************************************
' NOTE: set bumper object timer to around 150-175 in order to be able
' to actually see the animaation, adjust to your liking

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

'*******************************************
' ZBMP: Bumpers
'*******************************************


Sub Bumper1_Hit
  RandomSoundBumperTop Bumper1
  DOF 105, DOFPulse
  DOF 127, DOFPulse
  vpmTimer.PulseSw 51

  Dim BP
  For Each BP in BP_BmpSkirt_Y
    BP.roty=skirtAY(me,Activeball)
    BP.rotx=skirtAX(me,Activeball)
  Next
  I_lane_out_L.state = 1 - I_lane_out_L.state
  I_lane_out_R.state = 1 - I_lane_out_R.state
  me.timerinterval = 150
  me.timerenabled=1
  if L_BMP_Y.state = 0 then
    addscore base_bumper_points
  Else
    addscore lit_Y_bumper_points
  End If

End Sub

sub Bumper1_timer
  Dim BP
  For Each BP in BP_BmpSkirt_Y
    BP.roty=0
    BP.rotx=0
  Next
  me.timerenabled=0
end sub

Sub Bumper1_Animate
  Dim z: z = Bumper1.CurrentRingOffset

  Dim BP
  For Each BP in BP_BmpRing_Y : BP.transz = z : Next
End Sub

Sub Bumper2_Hit
  RandomSoundBumperTop Bumper2
  DOF 106, DOFPulse
  DOF 128, DOFPulse
  vpmTimer.PulseSw 51

  Dim BP
  For Each BP in BP_BmpSkirt_B
    BP.roty=skirtAY(me,Activeball)
    BP.rotx=skirtAX(me,Activeball)
  Next

  me.timerinterval = 150
  me.timerenabled=1
  if L_BMP_B.state = 0 then
    addscore base_bumper_points
  Else
    addscore lit_bumper_points
  End If
End Sub

sub Bumper2_timer
  Dim BP
  For Each BP in BP_BmpSkirt_B
    BP.roty=0
    BP.rotx=0
  Next
  me.timerenabled=0
end sub

Sub Bumper2_Animate
  Dim z: z = Bumper2.CurrentRingOffset

  Dim BP
  For Each BP in BP_BmpRing_B : BP.transz = z : Next
End Sub


Sub Bumper3_Hit
  RandomSoundBumperTop Bumper3
  DOF 107, DOFPulse
  DOF 129, DOFPulse
  vpmTimer.PulseSw 51

  Dim BP
  For Each BP in BP_BmpSkirt_G
    BP.roty=skirtAY(me,Activeball)
    BP.rotx=skirtAX(me,Activeball)
  Next

  me.timerinterval = 150
  me.timerenabled=1
  if L_BMP_G.state = 0 then
    addscore base_bumper_points
  Else
    addscore lit_bumper_points
  End If
End Sub

sub Bumper3_timer
  Dim BP
  For Each BP in BP_BmpSkirt_G
    BP.roty=0
    BP.rotx=0
  Next
  me.timerenabled=0
end sub

Sub Bumper3_Animate
  Dim z: z = Bumper3.CurrentRingOffset

  Dim BP
  For Each BP in BP_BmpRing_G : BP.transz = z : Next
End Sub




'*******************************************
' ZKIC: Kickers
'*******************************************


Sub Kicker1_Hit
  SoundSaucerLock
  Kicker1.TimerEnabled = True
End Sub

Sub Kicker1_Timer
  SoundSaucerKick 1,Kicker1
  Kicker1.Kick  180, 40
  Kicker1.TimerEnabled = False
End Sub

'******************************************************
'****  DROP TARGETS by Rothbauerw
'******************************************************


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

Dim DTB1, DTB2, DTB3, DTB4, DTB5, DTY1, DTY2, DTY3, DTY4, DTY5, DTG1, DTG2, DTG3, DTG4, DTG5

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

Set DTB1 = (new DropTarget)(dt_b1, dt_b1a, BM_dt_b1, 1, 0, false)
Set DTB2 = (new DropTarget)(dt_b2, dt_b2a, BM_dt_b2, 2, 0, false)
Set DTB3 = (new DropTarget)(dt_b3, dt_b3a, BM_dt_b3, 3, 0, false)
Set DTB4 = (new DropTarget)(dt_b4, dt_b4a, BM_dt_b4, 4, 0, false)
Set DTB5 = (new DropTarget)(dt_b5, dt_b5a, BM_dt_b5, 5, 0, false)

Set DTG1 = (new DropTarget)(dt_g1, dt_g1a, BM_dt_g1, 6, 0, false)
Set DTG2 = (new DropTarget)(dt_g2, dt_g2a, BM_dt_g2, 7, 0, false)
Set DTG3 = (new DropTarget)(dt_g3, dt_g3a, BM_dt_g3, 8, 0, false)
Set DTG4 = (new DropTarget)(dt_g4, dt_g4a, BM_dt_g4, 9, 0, false)
Set DTG5 = (new DropTarget)(dt_g5, dt_g5a, BM_dt_g5, 10, 0, false)

Set DTY1 = (new DropTarget)(dt_y1, dt_y1a, BM_dt_y1, 11, 0, false)
Set DTY2 = (new DropTarget)(dt_y2, dt_y2a, BM_dt_y2, 12, 0, false)
Set DTY3 = (new DropTarget)(dt_y3, dt_y3a, BM_dt_y3, 13, 0, false)
Set DTY4 = (new DropTarget)(dt_y4, dt_y4a, BM_dt_y4, 14, 0, false)
Set DTY5 = (new DropTarget)(dt_y5, dt_y5a, BM_dt_y5, 15, 0, false)

Dim DTArray
DTArray = Array(DTB1, DTB2, DTB3, DTB4, DTB5, DTG1, DTG2, DTG3, DTG4, DTG5, DTY1, DTY2, DTY3, DTY4, DTY5)
Dim DTB1up, DTB2up, DTB3up, DTB4up, DTB5up, DTY1up, DTY2up, DTY3up, DTY4up, DTY5up, DTG1up, DTG2up, DTG3up, DTG4up, DTG5up

' Bank completion flags - set True once a bank's event has fired; reset when bank is raised
Dim dtBlueDone, dtGreenDone, dtYellowDone, dtAllDone
dtBlueDone   = False
dtGreenDone  = False
dtYellowDone = False
dtAllDone    = False


'Configure the behavior of Drop Targets.
Const DTDropSpeed = 110 'in milliseconds
Const DTDropUpSpeed = 40 'in milliseconds
Const DTDropUnits = 58 'VP units primitive drops so top of at or below the playfield
Const DTDropUpUnits = 10 'VP units primitive raises above the up position on drops up
Const DTMaxBend = 8 '8 'max degrees primitive rotates when hit
Const DTDropDelay = 20 'time in milliseconds before target drops (due to friction/impact of the ball)
Const DTRaiseDelay = 40 'time in milliseconds before target drops back to normal up position after the solenoid fires to raise the target
Const DTBrickVel = 30 'velocity at which the target will brick, set to '0' to disable brick
Const DTEnableBrick = 1 'Set to 0 to disable bricking, 1 to enable bricking
Const Sound = "" 'Drop Target Hit sound
Const DTDropSound = "DropTarget_Down" 'Drop Target Drop sound
Const DTResetSound = "DropTarget_Up" 'Drop Target reset sound

Const DTMass = 0.2 'Mass of the Drop Target (between 0 and 1), higher values provide more resistance
Const DTResetDelay = 500 'ms to wait after all 15 targets drop before raising them (allows last drop to fully animate)


'******************************************************
'  DROP TARGETS FUNCTIONS
'******************************************************
' Initial Drop Target Shadows - Avoids a light DT hit and shadows go off when not strong enough hit to drop the target.


Sub DTHit(switch)
  ' debug.print("DTHit Call" & switch)
  Dim i
  i = DTArrayID(switch)

  ' Guard: ignore if already dropped or mid-animation
  If DTArray(i).isDropped Or DTArray(i).animate <> 0 Then
    Debug.Print "DTHit ignored sw" & switch & " (dropped=" & DTArray(i).isDropped & " animate=" & DTArray(i).animate & ")"
    Exit Sub
  End If

  PlayTargetSound
  DTArray(i).animate = DTCheckBrick(ActiveBall, DTArray(i).prim)
  If DTArray(i).animate = 1 Or DTArray(i).animate = 3 Or DTArray(i).animate = 4 Then
    DTBallPhysics ActiveBall, DTArray(i).prim.rotz, DTMass
  End If
  DoDTAnim
End Sub

Sub DTRaise(switch)
  Dim i
  i = DTArrayID(switch)

  DTArray(i).isDropped = False  ' reset immediately so the hit guard doesn't block re-drops
  DTArray(i).animate = -1
  DoDTAnim
End Sub

Sub DTDrop(switch)
  'debug.print("DTDrop Call")
  'debug.print("DTDrop" + cStr(switch))

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
  'debug.print("DTCheckBrick Call")
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
      DTArray(ind).isDropped = 1 'Mark target as dropped
      DTAction switchid
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
  ' debug.print("DTDropped call")
  Dim ind
  ind = DTArrayID(switchid)

  DTDropped = DTArray(ind).isDropped
End Function

Sub DTAction(switchid)
  ' Called from DTAnimate only when animate=2 fully completes (target is physically down).
  ' Bricks (animate=3) never reach here, so scoring and bank checks are naturally excluded.

  ' Score based on bank
  If switchid >= 1 And switchid < 6 Then
    ' debug.print("Blue target dropped")
    If I_Bank_B.state = 1 Then addscore lit_bank_points Else addscore base_bank_points
  ElseIf switchid >= 6 And switchid < 11 Then
    ' debug.print("Green target dropped")
    If I_Bank_G.state = 1 Then addscore lit_bank_points Else addscore base_bank_points
  Else
    ' debug.print("Yellow target dropped")
    If I_Bank_Y.state = 1 Then addscore lit_bank_points Else addscore base_bank_points
  End If

  checkdrops

  Select Case switchid
    Case 1: dts_b1.visible = False
    Case 2: dts_b2.visible = False
    Case 3: dts_b3.visible = False
    Case 4: dts_b4.visible = False
    Case 5: dts_b5.visible = False
    Case 6: dts_g5.visible = False
    Case 7: dts_g4.visible = False
    Case 8: dts_g3.visible = False
    Case 9: dts_g2.visible = False
    Case 10: dts_g1.visible = False
    Case 11: dts_y1.visible = False
    Case 12: dts_y2.visible = False
    Case 13: dts_y3.visible = False
    Case 14: dts_y4.visible = False
    Case 15: dts_y5.visible = False
  End Select
End Sub


'********************************************
' Drop Target Hits
'********************************************

Sub dt_b1_Hit
  DTHit 1
  DTB1up=0
End Sub

Sub dt_b2_Hit
  DTHit 2
  DTB2up=0
End Sub

Sub dt_b3_Hit
  DTHit 3
  DTB3up=0
End Sub

Sub dt_b4_Hit
  DTHit 4
  DTB4up=0
End Sub

Sub dt_b5_Hit
  DTHit 5
  DTB5up=0
End Sub

Sub dt_g1_Hit
  DTHit 6
  DTG1up=0
End Sub

Sub dt_g2_Hit
  DTHit 7
  DTG2up=0
End Sub

Sub dt_g3_Hit
  DTHit 8
  DTG3up=0
End Sub

Sub dt_g4_Hit
  DTHit 9
  DTG4up=0
End Sub

Sub dt_g5_Hit
  DTHit 10
  DTG5up=0
End Sub

Sub dt_y1_Hit
  DTHit 11
  DTY1up=0
End Sub

Sub dt_y2_Hit
  DTHit 12
  DTY2up=0
End Sub

Sub dt_y3_Hit
  DTHit 13
  DTY3up=0
End Sub

Sub dt_y4_Hit
  DTHit 14
  DTY4up=0
End Sub

Sub dt_y5_Hit
  DTHit 15
  DTY5up=0
End Sub

'******************************************************
'  DROP TARGET STATE MACHINE
'******************************************************

' Returns how many targets in DTArray[firstIdx..lastIdx] are currently dropped
Function CountBankDropped(firstIdx, lastIdx)
  Dim i, n
  n = 0
  For i = firstIdx To lastIdx
    If DTArray(i).isDropped Then n = n + 1
  Next
  CountBankDropped = n
End Function

' Called after every drop to evaluate bank / full-table completion.
' Uses one-shot flags so each event fires exactly once per cycle.
Sub checkdrops
  Dim blueDown, greenDown, yellowDown

  ' DTArray indices: Blue = 0-4 (sw 1-5), Green = 5-9 (sw 6-10), Yellow = 10-14 (sw 11-15)
  blueDown   = CountBankDropped(0,  4)
  greenDown  = CountBankDropped(5,  9)
  yellowDown = CountBankDropped(10, 14)

  'PrintDropTally blueDown, greenDown, yellowDown

  If blueDown = 5 And Not dtBlueDone Then
    dtBlueDone = True
    OnBlueBankComplete
  End If

  If greenDown = 5 And Not dtGreenDone Then
    dtGreenDone = True
    OnGreenBankComplete
  End If

  If yellowDown = 5 And Not dtYellowDone Then
    dtYellowDone = True
    OnYellowBankComplete
  End If

  If (blueDown + greenDown + yellowDown = 15) And Not dtAllDone Then
    dtAllDone = True
    OnAllTargetsDown
  End If
End Sub

' Prints a per-target tally to the debug output after each drop event.
Sub PrintDropTally(blueDown, greenDown, yellowDown)
  Dim i, s

  s = ""
  For i = 0 To 4
    If DTArray(i).isDropped Then s = s & "[X]" Else s = s & "[ ]"
  Next
  Debug.Print "Blue:   " & s & "  (" & blueDown & "/5)"

  s = ""
  For i = 5 To 9
    If DTArray(i).isDropped Then s = s & "[X]" Else s = s & "[ ]"
  Next
  Debug.Print "Green:  " & s & "  (" & greenDown & "/5)"

  s = ""
  For i = 10 To 14
    If DTArray(i).isDropped Then s = s & "[X]" Else s = s & "[ ]"
  Next
  Debug.Print "Yellow: " & s & "  (" & yellowDown & "/5)"

  Debug.Print "Total:  " & (blueDown + greenDown + yellowDown) & "/15"
End Sub

' Raise every target in the blue bank and re-arm its completion event
Sub ResetBlueBank
  Dim i
  For i = 1 To 5 : DTRaise i : Next
  dtBlueDone = False
  dtAllDone  = False
End Sub

' Raise every target in the green bank and re-arm its completion event
Sub ResetGreenBank
  Dim i
  For i = 6 To 10 : DTRaise i : Next
  dtGreenDone = False
  dtAllDone   = False
End Sub

' Raise every target in the yellow bank and re-arm its completion event
Sub ResetYellowBank
  Dim i
  For i = 11 To 15 : DTRaise i : Next
  dtYellowDone = False
  dtAllDone    = False
End Sub

' Raise all 15 targets and re-arm all completion events
Sub ResetAllDropTargets
  PlaySound DTResetSound
  ResetBlueBank
  ResetGreenBank
  ResetYellowBank
  DoDTAnim
End Sub

Sub OnBlueBankComplete
  ' Debug.Print "All 5 blue targets down!"
  I_special_L.state = 1
  I_lane_mid_L.state = 1
End Sub

Sub OnGreenBankComplete
  Debug.Print "All 5 green targets down!"
  I_special_R.state = 1
  I_lane_mid_R.state = 1
End Sub

Sub OnYellowBankComplete
  ' Debug.Print "All 5 yellow targets down!"
  I_lane_top_C.state = 1
End Sub

Sub OnAllTargetsDown
  ' Debug.Print "All 15 drop targets down!"
  dtResetStep = 1
  resetDT.Enabled = True
End Sub

' Timer fires once after DTResetDelay, resets all banks simultaneously, then stops
Sub resetDT_Timer
  resetDT.Enabled = False
  PlaySound DTResetSound
  ResetBlueBank
  ResetYellowBank
  ResetGreenBank
  dtResetStep = 0
  DoDTAnim
End Sub

'******************************************************
'****  END DROP TARGETS
'******************************************************




'****************************************************************
' ZSLG: Slingshots
'****************************************************************

sub InitRubberState
  'configure default state for any animated objects

  dim BP

  For Each BP in BP_rb_L_Mid_0 : bp.Visible = 1 : Next
  For Each BP in BP_rb_L_Mid_1 : bp.Visible = 0 : Next
  For Each BP in BP_rb_L_Mid_2 : bp.Visible = 0 : Next

  For Each BP in BP_rb_R_Mid_0 : bp.Visible = 1 : Next
  For Each BP in BP_rb_R_Mid_1 : bp.Visible = 0 : Next
  For Each BP in BP_rb_R_Mid_2 : bp.Visible = 0 : Next

  For Each BP in BP_rb_L_Mid_0 : bp.Visible = 1 : Next
  For Each BP in BP_rb_L_Mid_1 : bp.Visible = 0 : Next
  For Each BP in BP_rb_L_Mid_2 : bp.Visible = 0 : Next

  For Each BP in BP_rb_R_Mid_0 : bp.Visible = 1 : Next
  For Each BP in BP_rb_R_Mid_1 : bp.Visible = 0 : Next
  For Each BP in BP_rb_R_Mid_2 : bp.Visible = 0 : Next

  For Each BP in BP_rb_L_Top0_0 : bp.Visible = 1 : Next
  For Each BP in BP_rb_L_Top0_1 : bp.Visible = 0 : Next
  For Each BP in BP_rb_L_Top0_2 : bp.Visible = 0 : Next

  For Each BP in BP_rb_L_Top1_0 : bp.Visible = 1 : Next
  For Each BP in BP_rb_L_Top1_1 : bp.Visible = 0 : Next
  For Each BP in BP_rb_L_Top1_2 : bp.Visible = 0 : Next

  For Each BP in BP_rb_L_Top2_0 : bp.Visible = 1 : Next
  For Each BP in BP_rb_L_Top2_1 : bp.Visible = 0 : Next
  For Each BP in BP_rb_L_Top2_2 : bp.Visible = 0 : Next

  For Each BP in BP_rb_R_Top0_0 : bp.Visible = 1 : Next
  For Each BP in BP_rb_R_Top0_1 : bp.Visible = 0 : Next
  For Each BP in BP_rb_R_Top0_2 : bp.Visible = 0 : Next

  For Each BP in BP_rb_R_Top1_0 : bp.Visible = 1 : Next
  For Each BP in BP_rb_R_Top1_1 : bp.Visible = 0 : Next
  For Each BP in BP_rb_R_Top1_2 : bp.Visible = 0 : Next

  For Each BP in BP_rb_R_Top2_0 : bp.Visible = 1 : Next
  For Each BP in BP_rb_R_Top2_1 : bp.Visible = 0 : Next
  For Each BP in BP_rb_R_Top2_2 : bp.Visible = 0 : Next

  For Each BP in BP_rb_L_Sling_0 : bp.Visible = 1 : Next
  For Each BP in BP_rb_L_Sling_1 : bp.Visible = 0 : Next
  For Each BP in BP_rb_L_Sling_2 : bp.Visible = 0 : Next

  For Each BP in BP_rb_R_Sling_0 : bp.Visible = 1 : Next
  For Each BP in BP_rb_R_Sling_1 : bp.Visible = 0 : Next
  For Each BP in BP_rb_R_Sling_2 : bp.Visible = 0 : Next

  For Each BP in BP_rb_M_Sling_0 : bp.Visible = 1 : Next
  For Each BP in BP_rb_M_Sling_1 : bp.Visible = 0 : Next
  For Each BP in BP_rb_M_Sling_2 : bp.Visible = 0 : Next
end sub


Sub sw_sling_L_Hit:vpmTimer.PulseSw 53: timer_sling_L.enabled = true : addscore sling_points : End Sub
Sub rb_sling_L_Hit:                     timer_sling_L.enabled = true : End Sub

dim timer_sling_L_Counter : timer_sling_L_Counter = 0
sub timer_sling_L_Timer
  'BP_rb_L_Sling_0, BP_rb_L_Sling_1, BP_rb_L_Sling_2
  Dim BP

  Select Case timer_sling_L_Counter
    Case 1:
      ' rubber in
      For Each BP in BP_rb_L_Sling_0 : bp.Visible = 0 : Next
      For Each BP in BP_rb_L_Sling_1 : bp.Visible = 1 : Next
      For Each BP in BP_rb_L_Sling_2 : bp.Visible = 0 : Next
    Case 3:
      ' rubber out
      For Each BP in BP_rb_L_Sling_0 : bp.Visible = 0 : Next
      For Each BP in BP_rb_L_Sling_1 : bp.Visible = 0 : Next
      For Each BP in BP_rb_L_Sling_2 : bp.Visible = 1 : Next
    Case Else:
      ' default rubber position
      For Each BP in BP_rb_L_Sling_0 : bp.Visible = 1 : Next
      For Each BP in BP_rb_L_Sling_1 : bp.Visible = 0 : Next
      For Each BP in BP_rb_L_Sling_2 : bp.Visible = 0 : Next
  End Select

  if timer_sling_L_Counter >= 4 Then
    timer_sling_L_Counter = 0
    timer_sling_L.enabled = false
  Else
    timer_sling_L_Counter = timer_sling_L_Counter + 1
  end If
end Sub

' === Auto-generated handler and timer blocks for additional objects ===

' List of object bases:
' sling_R, bank_G, bank_B, bank_Y,
' top_lo_R, top_lo_L, top_mid_R, top_mid_L, top_up_R, top_up_L

Sub sw_sling_R_Hit:vpmTimer.PulseSw 53: timer_sling_R.enabled = true : addscore sling_points : End Sub
Sub rb_sling_R_Hit:                     timer_sling_R.enabled = true : End Sub

dim timer_sling_R_Counter : timer_sling_R_Counter = 0
sub timer_sling_R_Timer
  Dim BP
  Select Case timer_sling_R_Counter
    Case 1:
      For Each BP in BP_rb_R_Sling_0 : BP.Visible = 0 : Next
      For Each BP in BP_rb_R_Sling_1 : BP.Visible = 1 : Next
      For Each BP in BP_rb_R_Sling_2 : BP.Visible = 0 : Next
    Case 3:
      For Each BP in BP_rb_R_Sling_0 : BP.Visible = 0 : Next
      For Each BP in BP_rb_R_Sling_1 : BP.Visible = 0 : Next
      For Each BP in BP_rb_R_Sling_2 : BP.Visible = 1 : Next
    Case Else:
      For Each BP in BP_rb_R_Sling_0 : BP.Visible = 1 : Next
      For Each BP in BP_rb_R_Sling_1 : BP.Visible = 0 : Next
      For Each BP in BP_rb_R_Sling_2 : BP.Visible = 0 : Next
  End Select

  If timer_sling_R_Counter >= 4 Then
    timer_sling_R_Counter = 0
    timer_sling_R.enabled = false
  Else
    timer_sling_R_Counter = timer_sling_R_Counter + 1
  End If
End Sub

Sub sw_bank_G_Hit:vpmTimer.PulseSw 53: timer_bank_G.enabled = true : addscore sw_points : End Sub
Sub rb_bank_G_Hit:                     timer_bank_G.enabled = true : End Sub

dim timer_bank_G_Counter : timer_bank_G_Counter = 0
sub timer_bank_G_Timer
  Dim BP
  Select Case timer_bank_G_Counter
    Case 1:
      For Each BP in BP_rb_R_Mid_0 : BP.Visible = 0 : Next
      For Each BP in BP_rb_R_Mid_1 : BP.Visible = 1 : Next
      For Each BP in BP_rb_R_Mid_2 : BP.Visible = 0 : Next
    Case 3:
      For Each BP in BP_rb_R_Mid_0 : BP.Visible = 0 : Next
      For Each BP in BP_rb_R_Mid_1 : BP.Visible = 0 : Next
      For Each BP in BP_rb_R_Mid_2 : BP.Visible = 1 : Next
    Case Else:
      For Each BP in BP_rb_R_Mid_0 : BP.Visible = 1 : Next
      For Each BP in BP_rb_R_Mid_1 : BP.Visible = 0 : Next
      For Each BP in BP_rb_R_Mid_2 : BP.Visible = 0 : Next
  End Select

  If timer_bank_G_Counter >= 4 Then
    timer_bank_G_Counter = 0
    timer_bank_G.enabled = false
  Else
    timer_bank_G_Counter = timer_bank_G_Counter + 1
  End If
End Sub

Sub sw_bank_B_Hit:vpmTimer.PulseSw 53: timer_bank_B.enabled = true : addscore sw_points : End Sub
Sub rb_bank_B_Hit:                     timer_bank_B.enabled = true : End Sub

dim timer_bank_B_Counter : timer_bank_B_Counter = 0
sub timer_bank_B_Timer
  Dim BP
  Select Case timer_bank_B_Counter
    Case 1:
      For Each BP in BP_rb_L_Mid_0 : BP.Visible = 0 : Next
      For Each BP in BP_rb_L_Mid_1 : BP.Visible = 1 : Next
      For Each BP in BP_rb_L_Mid_2 : BP.Visible = 0 : Next
    Case 3:
      For Each BP in BP_rb_L_Mid_0 : BP.Visible = 0 : Next
      For Each BP in BP_rb_L_Mid_1 : BP.Visible = 0 : Next
      For Each BP in BP_rb_L_Mid_2 : BP.Visible = 1 : Next
    Case Else:
      For Each BP in BP_rb_L_Mid_0 : BP.Visible = 1 : Next
      For Each BP in BP_rb_L_Mid_1 : BP.Visible = 0 : Next
      For Each BP in BP_rb_L_Mid_2 : BP.Visible = 0 : Next
  End Select

  If timer_bank_B_Counter >= 4 Then
    timer_bank_B_Counter = 0
    timer_bank_B.enabled = false
  Else
    timer_bank_B_Counter = timer_bank_B_Counter + 1
  End If
End Sub

Sub sw_bank_Y_Hit:vpmTimer.PulseSw 53: timer_bank_Y.enabled = true : addscore sw_points : End Sub
Sub rb_bank_Y_Hit:                     timer_bank_Y.enabled = true : End Sub

dim timer_bank_Y_Counter : timer_bank_Y_Counter = 0
sub timer_bank_Y_Timer
  Dim BP
  Select Case timer_bank_Y_Counter
    Case 1:
      For Each BP in BP_rb_M_Sling_0 : BP.Visible = 0 : Next
      For Each BP in BP_rb_M_Sling_1 : BP.Visible = 1 : Next
      For Each BP in BP_rb_M_Sling_2 : BP.Visible = 0 : Next
    Case 3:
      For Each BP in BP_rb_M_Sling_0 : BP.Visible = 0 : Next
      For Each BP in BP_rb_M_Sling_1 : BP.Visible = 0 : Next
      For Each BP in BP_rb_M_Sling_2 : BP.Visible = 1 : Next
    Case Else:
      For Each BP in BP_rb_M_Sling_0 : BP.Visible = 1 : Next
      For Each BP in BP_rb_M_Sling_1 : BP.Visible = 0 : Next
      For Each BP in BP_rb_M_Sling_2 : BP.Visible = 0 : Next
  End Select

  If timer_bank_Y_Counter >= 4 Then
    timer_bank_Y_Counter = 0
    timer_bank_Y.enabled = false
  Else
    timer_bank_Y_Counter = timer_bank_Y_Counter + 1
  End If
End Sub

Sub sw_top_lo_R_Hit:vpmTimer.PulseSw 53: timer_top_lo_R.enabled = true : addscore sw_points : End Sub
Sub rb_top_lo_R_Hit:                     timer_top_lo_R.enabled = true : End Sub

dim timer_top_lo_R_Counter : timer_top_lo_R_Counter = 0
sub timer_top_lo_R_Timer
  Dim BP
  Select Case timer_top_lo_R_Counter
    Case 1:
      For Each BP in BP_rb_R_Top2_0 : BP.Visible = 0 : Next
      For Each BP in BP_rb_R_Top2_1 : BP.Visible = 1 : Next
      For Each BP in BP_rb_R_Top2_2 : BP.Visible = 0 : Next
    Case 3:
      For Each BP in BP_rb_R_Top2_0 : BP.Visible = 0 : Next
      For Each BP in BP_rb_R_Top2_1 : BP.Visible = 0 : Next
      For Each BP in BP_rb_R_Top2_2 : BP.Visible = 1 : Next
    Case Else:
      For Each BP in BP_rb_R_Top2_0 : BP.Visible = 1 : Next
      For Each BP in BP_rb_R_Top2_1 : BP.Visible = 0 : Next
      For Each BP in BP_rb_R_Top2_2 : BP.Visible = 0 : Next
  End Select

  If timer_top_lo_R_Counter >= 4 Then
    timer_top_lo_R_Counter = 0
    timer_top_lo_R.enabled = false
  Else
    timer_top_lo_R_Counter = timer_top_lo_R_Counter + 1
  End If
End Sub

Sub sw_top_lo_L_Hit:vpmTimer.PulseSw 53: timer_top_lo_L.enabled = true : addscore sw_points : End Sub
Sub rb_top_lo_L_Hit:                     timer_top_lo_L.enabled = true : End Sub

dim timer_top_lo_L_Counter : timer_top_lo_L_Counter = 0
sub timer_top_lo_L_Timer
  Dim BP
  Select Case timer_top_lo_L_Counter
    Case 1:
      For Each BP in BP_rb_L_Top2_0 : BP.Visible = 0 : Next
      For Each BP in BP_rb_L_Top2_1 : BP.Visible = 1 : Next
      For Each BP in BP_rb_L_Top2_2 : BP.Visible = 0 : Next
    Case 3:
      For Each BP in BP_rb_L_Top2_0 : BP.Visible = 0 : Next
      For Each BP in BP_rb_L_Top2_1 : BP.Visible = 0 : Next
      For Each BP in BP_rb_L_Top2_2 : BP.Visible = 1 : Next
    Case Else:
      For Each BP in BP_rb_L_Top2_0 : BP.Visible = 1 : Next
      For Each BP in BP_rb_L_Top2_1 : BP.Visible = 0 : Next
      For Each BP in BP_rb_L_Top2_2 : BP.Visible = 0 : Next
  End Select

  If timer_top_lo_L_Counter >= 4 Then
    timer_top_lo_L_Counter = 0
    timer_top_lo_L.enabled = false
  Else
    timer_top_lo_L_Counter = timer_top_lo_L_Counter + 1
  End If
End Sub

Sub sw_top_mid_R_Hit:vpmTimer.PulseSw 53: timer_top_mid_R.enabled = true : addscore sw_points : End Sub
Sub rb_top_mid_R_Hit:                     timer_top_mid_R.enabled = true : End Sub

dim timer_top_mid_R_Counter : timer_top_mid_R_Counter = 0
sub timer_top_mid_R_Timer
  Dim BP
  Select Case timer_top_mid_R_Counter
    Case 1:
      For Each BP in BP_rb_R_Top1_0 : BP.Visible = 0 : Next
      For Each BP in BP_rb_R_Top1_1 : BP.Visible = 1 : Next
      For Each BP in BP_rb_R_Top1_2 : BP.Visible = 0 : Next
    Case 3:
      For Each BP in BP_rb_R_Top1_0 : BP.Visible = 0 : Next
      For Each BP in BP_rb_R_Top1_1 : BP.Visible = 0 : Next
      For Each BP in BP_rb_R_Top1_2 : BP.Visible = 1 : Next
    Case Else:
      For Each BP in BP_rb_R_Top1_0 : BP.Visible = 1 : Next
      For Each BP in BP_rb_R_Top1_1 : BP.Visible = 0 : Next
      For Each BP in BP_rb_R_Top1_2 : BP.Visible = 0 : Next
  End Select

  If timer_top_mid_R_Counter >= 4 Then
    timer_top_mid_R_Counter = 0
    timer_top_mid_R.enabled = false
  Else
    timer_top_mid_R_Counter = timer_top_mid_R_Counter + 1
  End If
End Sub

Sub sw_top_mid_L_Hit:vpmTimer.PulseSw 53: timer_top_mid_L.enabled = true : addscore sw_points : End Sub
Sub rb_top_mid_L_Hit:                     timer_top_mid_L.enabled = true : End Sub

dim timer_top_mid_L_Counter : timer_top_mid_L_Counter = 0
sub timer_top_mid_L_Timer
  Dim BP
  Select Case timer_top_mid_L_Counter
    Case 1:
      For Each BP in BP_rb_L_Top1_0 : BP.Visible = 0 : Next
      For Each BP in BP_rb_L_Top1_1 : BP.Visible = 1 : Next
      For Each BP in BP_rb_L_Top1_2 : BP.Visible = 0 : Next
    Case 3:
      For Each BP in BP_rb_L_Top1_0 : BP.Visible = 0 : Next
      For Each BP in BP_rb_L_Top1_1 : BP.Visible = 0 : Next
      For Each BP in BP_rb_L_Top1_2 : BP.Visible = 1 : Next
    Case Else:
      For Each BP in BP_rb_L_Top1_0 : BP.Visible = 1 : Next
      For Each BP in BP_rb_L_Top1_1 : BP.Visible = 0 : Next
      For Each BP in BP_rb_L_Top1_2 : BP.Visible = 0 : Next
  End Select

  If timer_top_mid_L_Counter >= 4 Then
    timer_top_mid_L_Counter = 0
    timer_top_mid_L.enabled = false
  Else
    timer_top_mid_L_Counter = timer_top_mid_L_Counter + 1
  End If
End Sub

Sub sw_top_up_R_Hit:vpmTimer.PulseSw 53: timer_top_up_R.enabled = true : addscore sw_points : End Sub
Sub rb_top_up_R_Hit:                     timer_top_up_R.enabled = true : End Sub

dim timer_top_up_R_Counter : timer_top_up_R_Counter = 0
sub timer_top_up_R_Timer
  Dim BP
  Select Case timer_top_up_R_Counter
    Case 1:
      For Each BP in BP_rb_R_Top0_0 : BP.Visible = 0 : Next
      For Each BP in BP_rb_R_Top0_1 : BP.Visible = 1 : Next
      For Each BP in BP_rb_R_Top0_2 : BP.Visible = 0 : Next
    Case 3:
      For Each BP in BP_rb_R_Top0_0 : BP.Visible = 0 : Next
      For Each BP in BP_rb_R_Top0_1 : BP.Visible = 0 : Next
      For Each BP in BP_rb_R_Top0_2 : BP.Visible = 1 : Next
    Case Else:
      For Each BP in BP_rb_R_Top0_0 : BP.Visible = 1 : Next
      For Each BP in BP_rb_R_Top0_1 : BP.Visible = 0 : Next
      For Each BP in BP_rb_R_Top0_2 : BP.Visible = 0 : Next
  End Select

  If timer_top_up_R_Counter >= 4 Then
    timer_top_up_R_Counter = 0
    timer_top_up_R.enabled = false
  Else
    timer_top_up_R_Counter = timer_top_up_R_Counter + 1
  End If
End Sub

Sub sw_top_up_L_Hit:vpmTimer.PulseSw 53: timer_top_up_L.enabled = true : addscore sw_points : End Sub
Sub rb_top_up_L_Hit:                     timer_top_up_L.enabled = true : End Sub

dim timer_top_up_L_Counter : timer_top_up_L_Counter = 0
sub timer_top_up_L_Timer
  Dim BP
  Select Case timer_top_up_L_Counter
    Case 1:
      For Each BP in BP_rb_L_Top0_0 : BP.Visible = 0 : Next
      For Each BP in BP_rb_L_Top0_1 : BP.Visible = 1 : Next
      For Each BP in BP_rb_L_Top0_2 : BP.Visible = 0 : Next
    Case 3:
      For Each BP in BP_rb_L_Top0_0 : BP.Visible = 0 : Next
      For Each BP in BP_rb_L_Top0_1 : BP.Visible = 0 : Next
      For Each BP in BP_rb_L_Top0_2 : BP.Visible = 1 : Next
    Case Else:
      For Each BP in BP_rb_L_Top0_0 : BP.Visible = 1 : Next
      For Each BP in BP_rb_L_Top0_1 : BP.Visible = 0 : Next
      For Each BP in BP_rb_L_Top0_2 : BP.Visible = 0 : Next
  End Select

  If timer_top_up_L_Counter >= 4 Then
    timer_top_up_L_Counter = 0
    timer_top_up_L.enabled = false
  Else
    timer_top_up_L_Counter = timer_top_up_L_Counter + 1
  End If
End Sub



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
' Late 70's to early 80's

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
End Sub

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
  FlipperNudge RightFlipper, RFEndAngle, RFEOSNudge, LeftFlipper, LFEndAngle
  FlipperNudge LeftFlipper, LFEndAngle, LFEOSNudge,  RightFlipper, RFEndAngle
End Sub

Dim LFEOSNudge, RFEOSNudge

Sub FlipperNudge(Flipper1, Endangle1, EOSNudge1, Flipper2, EndAngle2)
  Dim b
  gBOT = GetBalls

  If Flipper1.currentangle = Endangle1 And EOSNudge1 <> 1 Then
    EOSNudge1 = 1
    '  debug.print Flipper1.currentangle &" = "& Endangle1 &"--"& Flipper2.currentangle &" = "& EndAngle2
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
'Const EOSReturn = 0.055  'EM's
'Const EOSReturn = 0.045  'late 70's to mid 80's
'Const EOSReturn = 0.035  'mid 80's to early 90's
Const EOSReturn = 0.025  'mid 90's and later

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
    Dim b
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

'Gate hit sound
Sub Gate3sfx_hit
  If activeball.velx > 0 Then
    RandomSoundMetal
  End If
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

'****  VPW AMBIENT BALL SHADOW (simplified from DynamicBSUpdate by Iakki, Apophis, and Wylte)

Sub DynamicBSInit()
  Dim iii
  For iii = 0 To ShadowBalls - 1
    Set objBallShadow(iii) = Eval("BallShadow" & iii)
    UpdateMaterial objBallShadow(iii).material, 1, 0, 0, 0, 0, 0, AmbientBSFactor, RGB(0,0,0), 0, 0, False, True, 0, 0, 0, 0
    objBallShadow(iii).Z = 1 + iii / 1000 + 0.04
    objBallShadow(iii).visible = False
    Set objBallShadowA(iii) = Eval("BallShadowA" & iii)
    objBallShadowA(iii).Opacity = 100 * AmbientBSFactor
    objBallShadowA(iii).visible = False
  Next
End Sub

Sub DynamicBSUpdate()
  Dim s
  Dim gBOT: gBOT = GetBalls

  ' Hide shadows for any ball slots beyond our shadow objects
  For s = ShadowBalls To tnob - 1
    ' no shadow objects beyond ShadowBalls, nothing to hide
  Next

  If UBound(gBOT) < lob Then
    For s = 0 To ShadowBalls - 1
      objBallShadow(s).visible = False
      objBallShadowA(s).visible = False
    Next
    Exit Sub
  End If

  For s = lob To UBound(gBOT)
    If s >= ShadowBalls Then Exit For
    If gBOT(s).Z <= 30 And gBOT(s).Z > 20 Then
      ' On playfield - show primitive shadow
      objBallShadow(s).visible = True
      objBallShadow(s).X = gBOT(s).X + (gBOT(s).X - (tablewidth / 2)) / (BallSize / AmbientMovement) + offsetX
      objBallShadow(s).Y = gBOT(s).Y + offsetY
      objBallShadowA(s).visible = False
    ElseIf gBOT(s).Z <= 20 Then
      ' Under playfield - show flasher shadow
      objBallShadow(s).visible = False
      objBallShadowA(s).visible = True
      objBallShadowA(s).X = gBOT(s).X + offsetX
      objBallShadowA(s).Y = gBOT(s).Y + offsetY
      objBallShadowA(s).height = gBOT(s).Z - BallSize / 4 + s / 1000
    Else
      ' Above playfield - hide both
      objBallShadow(s).visible = False
      objBallShadowA(s).visible = False
    End If
  Next
End Sub

Sub GiON
  Dim x
  For Each x In GI
    x.State = 1
  Next
  Dim xx: For Each xx In DTShadows: xx.Visible = True: Next
  Sound_GI_Relay 1, lane_top3
End Sub

Sub GiOFF
  Dim x
  For Each x In GI
    x.State = 0
  Next
  Dim xx: For Each xx In DTShadows: xx.Visible = False: Next
  Sound_GI_Relay 0, lane_top3
End Sub

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


'************************************************************
'**** ZSCR - Scoring Routines
'************************************************************
' Simulates old-school mechanical score reels:
' - Large scores (>=500) snapshot bumper lights and dim them while reels turn
' - Lights restore when the timer finishes
' Requires three timer objects in the VPX editor (all initially disabled):
'   AddScore10Timer   - interval ~50ms
'   AddScore100Timer  - interval ~80ms
'   AddScore1000Timer - interval ~100ms

Dim Add10, Add100, Add1000
Dim BmpMemY, BmpMemB, BmpMemG, BumperMemoryUse

Sub MemorizeBumperStates
  If Not BumperMemoryUse Then
    BmpMemY = L_BMP_Y.state
    BmpMemB = L_BMP_B.state
    BmpMemG = L_BMP_G.state
  End If
  BumperMemoryUse = True
End Sub

Sub BumperLightsOFF
  L_BMP_Y.state = 0
  L_BMP_B.state = 0
  L_BMP_G.state = 0
End Sub

Sub BumperRelightFromMemory
  L_BMP_Y.state = BmpMemY
  L_BMP_B.state = BmpMemB
  L_BMP_G.state = BmpMemG
  BumperMemoryUse = False
End Sub

Sub addscore(points)
  If tilt Then Exit Sub
  If points >= 500 Then MemorizeBumperStates
  If points < 100 And Not AddScore10Timer.Enabled Then
    Add10 = points \ 10
    AddScore10Timer.Enabled = True
  ElseIf points < 1000 And Not AddScore100Timer.Enabled Then
    Add100 = points \ 100
    AddScore100Timer.Enabled = True
  ElseIf Not AddScore1000Timer.Enabled Then
    Add1000 = points \ 1000
    AddScore1000Timer.Enabled = True
  End If
End Sub

Sub AddScore10Timer_Timer()
  If Add10 > 0 Then
    AddPoints 10
    Add10 = Add10 - 1
  Else
    If BumperMemoryUse Then BumperRelightFromMemory
    Me.Enabled = False
  End If
End Sub

Sub AddScore100Timer_Timer()
  If Add100 > 0 Then
    AddPoints 100
    Add100 = Add100 - 1
    If BumperMemoryUse Then BumperLightsOFF
  Else
    If BumperMemoryUse Then BumperRelightFromMemory
    Me.Enabled = False
  End If
End Sub

Sub AddScore1000Timer_Timer()
  If Add1000 > 0 Then
    AddPoints 1000
    Add1000 = Add1000 - 1
    If BumperMemoryUse Then BumperLightsOFF
  Else
    If BumperMemoryUse Then BumperRelightFromMemory
    Me.Enabled = False
  End If
End Sub

Sub AddPoints(points)
  score = score + points
  If B2SOn Then
    If score >= 100000 Then
      Controller.B2SSetScoreRolloverPlayer1 1
      Controller.B2SSetScorePlayer 1, score Mod 100000
    Else
      Controller.B2SSetScorePlayer 1, score
    End If
  End If

  ' Chime unit: bell1000 on 1000s rollover, bell100 on 100s rollover, bell10 otherwise
  If points = 1000 Then
    PlaySound SoundFXDOF("bell1000", 143, DOFPulse, DOFChimes)
  ElseIf points = 100 And (score Mod 1000) \ 100 = 0 Then
    PlaySound SoundFXDOF("bell1000", 143, DOFPulse, DOFChimes)
  ElseIf points = 100 Then
    PlaySound SoundFXDOF("bell100", 142, DOFPulse, DOFChimes)
  ElseIf points = 10 And (score Mod 100) \ 10 = 0 Then
    PlaySound SoundFXDOF("bell100", 142, DOFPulse, DOFChimes)
  Else
    PlaySound SoundFXDOF("bell10", 141, DOFPulse, DOFChimes)
  End If
End Sub


'************************************************************
'**** ZTRG - Triggers for Switches
'************************************************************

sub TGtopB_hit
  DOF 116, DOFPulse
  I_lane_B.state=0
  I_Bank_B.state=1
  L_Bmp_B.state=1
  If I_special_L.state = 1 Then
    I_special_L.state = 0
    AwardSpecial
  Else
    addscore base_lane_points
  End If
end sub

sub TGtopY_hit
  DOF 117, DOFPulse
  Dim wasLit: wasLit = (I_Bank_Y.state = 1)
  I_lane_Y.state=0
  I_Bank_Y.state=1
  L_Bmp_Y.state=1
  If wasLit Then
    addscore lit_lane_points
  Else
    addscore base_lane_points
  End If

end sub

sub TGtopG_hit
  DOF 118, DOFPulse
  I_lane_G.state=0
  I_Bank_G.state=1
  L_Bmp_G.state=1
  If I_special_R.state = 1 Then
    I_special_R.state = 0
    AwardSpecial
  Else
    addscore base_lane_points
  End If
end sub

Sub AwardSpecial
  If credit < 9 Then credit = credit + 1
  If B2SOn Then Controller.B2SSetCredits credit
  KnockerSolenoid
  DOF 130, DOFPulse
  DOF 123, DOFOn
  ' Debug.Print "SPECIAL! Credit awarded. Total credits: " & credit
End Sub

Sub TGmid_L_Hit
  DOF 114, DOFPulse
  If I_lane_mid_L.state = 1 Then addscore lit_lane_points Else addscore base_lane_points
End Sub

Sub TGmid_R_Hit
  DOF 115, DOFPulse
  If I_lane_mid_R.state = 1 Then addscore lit_lane_points Else addscore base_lane_points
End Sub

Sub TGin_L_Hit
  DOF 110, DOFPulse
  addscore base_lane_points
End Sub
Sub TGin_R_Hit
  DOF 111, DOFPulse
  addscore base_lane_points
End Sub

Sub TGout_L_Hit
  DOF 109, DOFPulse
  If I_lane_out_L.state = 1 Then addscore lit_lane_points Else addscore base_lane_points
End Sub

Sub TGout_R_Hit
  DOF 112, DOFPulse
  If I_lane_out_R.state = 1 Then addscore lit_lane_points Else addscore base_lane_points
End Sub

'************************************************************
'**** ZVLM - Bake maps to follow physical movement
'************************************************************

'************************
' Gates Brick maps to follow the VP gates
'************************

Sub Gate_Animate
  Dim a: a = Gate.CurrentAngle
  Dim BP : For Each BP in BP_GateFlap: BP.RotX = -a - 25 : Next
End Sub

'************************
''''' Switch Animations
'************************

Sub TGtopB_Animate
  Dim z : z = TGtopB.CurrentAnimOffset
  Dim BP : For Each BP in BP_tg_top_B : BP.transz = z: Next
End Sub

Sub TGtopY_Animate
  Dim z : z = TGtopY.CurrentAnimOffset
  Dim BP : For Each BP in BP_tg_top_Y  : BP.transz = z: Next
End Sub

Sub TGtopG_Animate
  Dim z : z = TGtopG.CurrentAnimOffset
  Dim BP : For Each BP in BP_tg_top_G : BP.transz = z: Next
End Sub

Sub TGmid_L_Animate
  Dim z : z = TGmid_L.CurrentAnimOffset
  Dim BP : For Each BP in BP_tg_mid_L  : BP.transz = z: Next
End Sub

Sub TGmid_R_Animate
  Dim z : z = TGmid_R.CurrentAnimOffset
  Dim BP : For Each BP in BP_tg_med_R : BP.transz = z: Next
End Sub

Sub TGin_L_Animate
  Dim z : z = TGin_L.CurrentAnimOffset
  Dim BP : For Each BP in BP_tg_in_L : BP.transz = z: Next
End Sub

Sub TGin_R_Animate
  Dim z : z = TGin_R.CurrentAnimOffset
  Dim BP : For Each BP in BP_tg_in_R : BP.transz = z: Next
End Sub

Sub TGout_L_Animate
  Dim z : z = TGout_L.CurrentAnimOffset
  Dim BP : For Each BP in BP_tg_out_L : BP.transz = z: Next
End Sub

Sub TGout_R_Animate
  Dim z : z = TGout_R.CurrentAnimOffset
  Dim BP : For Each BP in BP_tg_out_R : BP.transz = z: Next
End Sub

'************************
''''' Flipper Animations
'************************

Dim max_angle_RF, min_angle_RF, mid_angle_RF            ' min and max angles from the right flipper
max_angle_RF = RightFlipper.StartAngle              ' right flipper down angle
min_angle_RF = RightFlipper.EndAngle              ' right flipper up angle
mid_angle_RF = (max_angle_RF-min_angle_RF)/2 + min_angle_RF   ' right flipper bake map switch point angle

Dim max_angle_LF, min_angle_LF, mid_angle_LF            ' min and max angles from the left flipper
max_angle_LF = LeftFlipper.StartAngle             ' left flipper down angle
min_angle_LF = LeftFlipper.EndAngle               ' left flipper up angle
mid_angle_LF = (max_angle_LF-min_angle_LF)/2 + min_angle_LF   ' left flipper bake map switch point angle

Sub LeftFlipper_Animate
  Dim BP
  Dim a: a = LeftFlipper.CurrentAngle       ' store flipper angle in a
  FlipperLSh.RotZ = a
  For Each BP in BP_Flip_L
    BP.RotZ = a                 ' rotate the maps
    BP.visible = a > mid_angle_LF
  Next
  For Each BP in BP_Flip_LU
    BP.RotZ = a                 ' rotate the maps
    BP.visible = a < mid_angle_LF
  Next
End Sub


Sub RightFlipper_Animate
  Dim BP
  Dim a: a = RightFlipper.CurrentAngle        ' store flipper angle in a
  FlipperRSh.RotZ = a
  For Each BP in BP_Flip_R
    BP.RotZ = a                 ' rotate the maps
    BP.visible = a < mid_angle_RF
  Next
  For Each BP in BP_Flip_RU
    BP.visible = a > mid_angle_RF
    BP.RotZ = a                 ' rotate the maps
  Next
End Sub

'************************
''''' Drop Target Animations
'************************

Sub UpdateDropTargets()
  Dim BP, tz, rx, ry

  'Blue Targets
  tz = BM_dt_b1.transz
  rx = BM_dt_b1.rotx
  ry = BM_dt_b1.roty
  For Each BP in BP_dt_b1: BP.transz=tz: BP.rotx = rx: BP.roty = ry: Next

  tz = BM_dt_b2.transz
  rx = BM_dt_b2.rotx
  ry = BM_dt_b2.roty
  For Each BP in BP_dt_b2: BP.transz=tz: BP.rotx = rx: BP.roty = ry: Next

  tz = BM_dt_b3.transz
  rx = BM_dt_b3.rotx
  ry = BM_dt_b3.roty
  For Each BP in BP_dt_b3: BP.transz=tz: BP.rotx = rx: BP.roty = ry: Next

  tz = BM_dt_b4.transz
  rx = BM_dt_b4.rotx
  ry = BM_dt_b4.roty
  For Each BP in BP_dt_b4: BP.transz=tz: BP.rotx = rx: BP.roty = ry: Next

  tz = BM_dt_b5.transz
  rx = BM_dt_b5.rotx
  ry = BM_dt_b5.roty
  For Each BP in BP_dt_b5: BP.transz=tz: BP.rotx = rx: BP.roty = ry: Next

  'Yellow Targets
  tz = BM_dt_y1.transz
  rx = BM_dt_y1.rotx
  ry = BM_dt_y1.roty
  For Each BP in BP_dt_y1: BP.transz=tz: BP.rotx = rx: BP.roty = ry: Next

  tz = BM_dt_y2.transz
  rx = BM_dt_y2.rotx
  ry = BM_dt_y2.roty
  For Each BP in BP_dt_y2: BP.transz=tz: BP.rotx = rx: BP.roty = ry: Next

  tz = BM_dt_y3.transz
  rx = BM_dt_y3.rotx
  ry = BM_dt_y3.roty
  For Each BP in BP_dt_y3: BP.transz=tz: BP.rotx = rx: BP.roty = ry: Next

  tz = BM_dt_y4.transz
  rx = BM_dt_y4.rotx
  ry = BM_dt_y4.roty
  For Each BP in BP_dt_y4: BP.transz=tz: BP.rotx = rx: BP.roty = ry: Next

  tz = BM_dt_y5.transz
  rx = BM_dt_y5.rotx
  ry = BM_dt_y5.roty
  For Each BP in BP_dt_y5: BP.transz=tz: BP.rotx = rx: BP.roty = ry: Next

  'Green Targets
  tz = BM_dt_g1.transz
  rx = BM_dt_g1.rotx
  ry = BM_dt_g1.roty
  For Each BP in BP_dt_g1: BP.transz=tz: BP.rotx = rx: BP.roty = ry: Next

  tz = BM_dt_g2.transz
  rx = BM_dt_g2.rotx
  ry = BM_dt_g2.roty
  For Each BP in BP_dt_g2: BP.transz=tz: BP.rotx = rx: BP.roty = ry: Next

  tz = BM_dt_g3.transz
  rx = BM_dt_g3.rotx
  ry = BM_dt_g3.roty
  For Each BP in BP_dt_g3: BP.transz=tz: BP.rotx = rx: BP.roty = ry: Next

  tz = BM_dt_g4.transz
  rx = BM_dt_g4.rotx
  ry = BM_dt_g4.roty
  For Each BP in BP_dt_g4: BP.transz=tz: BP.rotx = rx: BP.roty = ry: Next

  tz = BM_dt_g5.transz
  rx = BM_dt_g5.rotx
  ry = BM_dt_g5.roty
  For Each BP in BP_dt_g5: BP.transz=tz: BP.rotx = rx: BP.roty = ry: Next
End Sub


'/////////////////////////////////////////////////////////////////
'         End Mechanical Sounds
'/////////////////////////////////////////////////////////////////

'******************************************************
'****  FLEEP MECHANICAL SOUNDS
'******************************************************


'*******************************************
'  ZHSC: High Score Tracking
'*******************************************

Sub HighScoreTimer_Timer
  If EnteringInitials Then
    If HSTimerCount = 1 Then
      SetHSLine 3, InitialString & Mid(AlphaString, AlphaStringPos, 1)
      HSTimerCount = 2
    Else
      SetHSLine 3, InitialString
      HSTimerCount = 1
    End If
  ElseIf InProgress Then
    SetHSLine 1, "HIGH SCORE1"
    SetHSLine 2, HSScore(1)
    SetHSLine 3, HSName(1)
    ' Keep timer running so it stays updated during play
  ElseIf CheckAllScores Then
    NewHighScore ScoreChecker, 1
  Else
    HighScoreTimer.interval = 2000
    HSTimerCount = HSTimerCount + 1
    If HSTimerCount > 5 Then HSTimerCount = 1
    SetHSLine 1, "HIGH SCORE" & FormatNumber(HSTimerCount, 0)
    SetHSLine 2, HSScore(HSTimerCount)
    SetHSLine 3, HSName(HSTimerCount)
  End If
End Sub

Function GetHSChar(Str, Index)
  Dim ThisChar, FileName
  ThisChar = Mid(Str, Index, 1)
  FileName = "PostIt"
  If ThisChar = " " Or ThisChar = "" Then
    FileName = FileName & "BL"
  ElseIf ThisChar = "<" Then
    FileName = FileName & "LT"
  ElseIf ThisChar = "_" Then
    FileName = FileName & "SP"
  Else
    FileName = FileName & ThisChar
  End If
  GetHSChar = FileName
End Function

Sub SetHSLine(LineNo, Str)
  Dim xfor, Index
  Dim StartHSArray, EndHSArray
  StartHSArray = Array(0, 1, 12, 22)
  EndHSArray   = Array(0, 11, 21, 31)
  Index = 1
  For xfor = StartHSArray(LineNo) To EndHSArray(LineNo)
    Eval("HS" & xfor).image = GetHSChar(CStr(Str), Index)
    Index = Index + 1
  Next
End Sub

Sub NewHighScore(NewScore, PlayNum)
  If NewScore > HSScore(5) Then
    HighScoreTimer.interval = 500
    HSTimerCount = 1
    AlphaStringPos = 1
    EnteringInitials = 1
    InitialString = ""
    SetHSLine 1, "PLAYER " & FormatNumber(PlayNum, 0)
    SetHSLine 2, "ENTER NAME"
    SetHSLine 3, Mid(AlphaString, AlphaStringPos, 1)
    HSNewHigh = NewScore
  End If
  CheckAllScores = 0
End Sub

Sub CollectInitials(keycode)
  If keycode = LeftFlipperKey Then
    AlphaStringPos = AlphaStringPos - 1
    If AlphaStringPos < 1 Then
      AlphaStringPos = Len(AlphaString)
      If InitialString = "" Then AlphaStringPos = AlphaStringPos - 1
    End If
    SetHSLine 3, InitialString & Mid(AlphaString, AlphaStringPos, 1)
    PlaySound "DropTargetDropped"
  ElseIf keycode = RightFlipperKey Then
    AlphaStringPos = AlphaStringPos + 1
    If AlphaStringPos > Len(AlphaString) Or (AlphaStringPos = Len(AlphaString) And InitialString = "") Then
      AlphaStringPos = 1
    End If
    SetHSLine 3, InitialString & Mid(AlphaString, AlphaStringPos, 1)
    PlaySound "DropTargetDropped"
  ElseIf keycode = StartGameKey Or keycode = PlungerKey Then
    Dim SelectedChar
    SelectedChar = Mid(AlphaString, AlphaStringPos, 1)
    If SelectedChar = "_" Then
      InitialString = InitialString & " "
      PlaySound "Ding10"
    ElseIf SelectedChar = "<" Then
      InitialString = Mid(InitialString, 1, Len(InitialString) - 1)
      If Len(InitialString) = 0 Then AlphaStringPos = 1
      PlaySound "Ding100"
    Else
      InitialString = InitialString & SelectedChar
      PlaySound "Ding10"
    End If
    If Len(InitialString) < 3 Then
      SetHSLine 3, InitialString & SelectedChar
    End If
  End If
  If Len(InitialString) = 3 Then
    Dim i
    For i = 5 To 1 Step -1
      If i = 1 Or (HSNewHigh > HSScore(i) And HSNewHigh <= HSScore(i - 1)) Then
        If i < 5 Then
          HSScore(i + 1) = HSScore(i)
          HSName(i + 1) = HSName(i)
        End If
        EnteringInitials = 0
        HSScore(i) = HSNewHigh
        HSName(i) = InitialString
        HSTimerCount = 5
        HighScoreTimer_Timer
        HighScoreTimer.interval = 2000
        PlaySound "Ding1000"
        Exit Sub
      ElseIf i < 5 Then
        HSScore(i + 1) = HSScore(i)
        HSName(i + 1) = HSName(i)
      End If
    Next
  End If
End Sub

Sub CheckHighScore
  If score > HSScore(5) Then
    ScoreChecker = score
    CheckAllScores = 1
    HighScoreTimer.interval = 100
    HighScoreTimer.enabled = True
  End If
End Sub

Sub savehs
  Dim xx, FileObj, ScoreFile
  Set FileObj = CreateObject("Scripting.FileSystemObject")
  If Not FileObj.FolderExists(UserDirectory) Then
    Set FileObj = Nothing : Exit Sub
  End If
  Set ScoreFile = FileObj.CreateTextFile(UserDirectory & HSFileName, True)
  For xx = 1 To 5 : ScoreFile.WriteLine HSScore(xx) : Next
  For xx = 1 To 5 : ScoreFile.WriteLine HSName(xx) : Next
  On Error Resume Next
  ScoreFile.WriteLine credit
  ScoreFile.WriteLine score
  On Error GoTo 0
  ScoreFile.Close
  Set ScoreFile = Nothing
  Set FileObj = Nothing
End Sub

Sub loadhs
  Dim FileObj, ScoreFile
  Dim t1, t2, t3, t4, t5, t6, t7, t8, t9, t10, t11, t12
  Set FileObj = CreateObject("Scripting.FileSystemObject")
  If Not FileObj.FolderExists(UserDirectory) Then Set FileObj = Nothing : Exit Sub
  If Not FileObj.FileExists(UserDirectory & HSFileName) Then Set FileObj = Nothing : Exit Sub
  Set ScoreFile = FileObj.GetFile(UserDirectory & HSFileName)
  Dim TextStr : Set TextStr = ScoreFile.OpenAsTextStream(1, 0)
  If TextStr.AtEndOfStream Then Set FileObj = Nothing : Exit Sub
  t1 = TextStr.ReadLine : t2 = TextStr.ReadLine : t3 = TextStr.ReadLine
  t4 = TextStr.ReadLine : t5 = TextStr.ReadLine
  t6 = TextStr.ReadLine : t7 = TextStr.ReadLine : t8 = TextStr.ReadLine
  t9 = TextStr.ReadLine : t10 = TextStr.ReadLine
  On Error Resume Next
  t11 = TextStr.ReadLine
  t12 = TextStr.ReadLine
  On Error GoTo 0
  TextStr.Close
  HSScore(1) = CLng(t1) : HSScore(2) = CLng(t2) : HSScore(3) = CLng(t3)
  HSScore(4) = CLng(t4) : HSScore(5) = CLng(t5)
  HSName(1) = t6 : HSName(2) = t7 : HSName(3) = t8
  HSName(4) = t9 : HSName(5) = t10
  If t11 <> "" Then credit = CDbl(t11)
  If t12 <> "" Then score  = CDbl(t12)
  Set ScoreFile = Nothing
  Set FileObj = Nothing
End Sub


' VLM  Arrays - Start
' Arrays per baked part
Dim BP_BmpRing_B: BP_BmpRing_B=Array(BM_BmpRing_B, LM_L_L_Bmp_B_BmpRing_B, LM_GI_lane_top1_BmpRing_B, LM_GI_lane_top2_BmpRing_B, LM_GI_lane_top3_BmpRing_B, LM_GI_top_center_L_BmpRing_B, LM_GI_top_center_R_BmpRing_B, LM_GI_top_side_L_BmpRing_B)
Dim BP_BmpRing_G: BP_BmpRing_G=Array(BM_BmpRing_G, LM_L_L_Bmp_G_BmpRing_G, LM_GI_lane_top1_BmpRing_G, LM_GI_lane_top2_BmpRing_G, LM_GI_lane_top3_BmpRing_G, LM_GI_lane_top4_BmpRing_G, LM_GI_top_center_L_BmpRing_G, LM_GI_top_center_R_BmpRing_G, LM_GI_top_side_R_BmpRing_G)
Dim BP_BmpRing_Y: BP_BmpRing_Y=Array(BM_BmpRing_Y, LM_L_L_Bmp_Y_BmpRing_Y, LM_GI_lane_top1_BmpRing_Y, LM_GI_lane_top2_BmpRing_Y, LM_GI_lane_top3_BmpRing_Y, LM_GI_lane_top4_BmpRing_Y, LM_GI_top_center_L_BmpRing_Y, LM_GI_top_center_R_BmpRing_Y, LM_GI_top_side_L_BmpRing_Y, LM_GI_top_side_R_BmpRing_Y)
Dim BP_BmpSkirt_B: BP_BmpSkirt_B=Array(BM_BmpSkirt_B, LM_L_L_Bmp_B_BmpSkirt_B, LM_GI_top_center_L_BmpSkirt_B)
Dim BP_BmpSkirt_G: BP_BmpSkirt_G=Array(BM_BmpSkirt_G, LM_L_L_Bmp_G_BmpSkirt_G, LM_GI_lane_top3_BmpSkirt_G, LM_GI_lane_top4_BmpSkirt_G, LM_GI_top_center_R_BmpSkirt_G, LM_GI_top_side_R_BmpSkirt_G)
Dim BP_BmpSkirt_Y: BP_BmpSkirt_Y=Array(BM_BmpSkirt_Y, LM_L_L_Bmp_Y_BmpSkirt_Y, LM_GI_lane_top1_BmpSkirt_Y, LM_GI_lane_top2_BmpSkirt_Y, LM_GI_lane_top3_BmpSkirt_Y, LM_GI_lane_top4_BmpSkirt_Y, LM_GI_mid_side_L1_BmpSkirt_Y, LM_GI_mid_side_R1_BmpSkirt_Y, LM_GI_top_center_L_BmpSkirt_Y, LM_GI_top_center_R_BmpSkirt_Y, LM_GI_top_side_L_BmpSkirt_Y, LM_GI_top_side_R_BmpSkirt_Y)
Dim BP_Flip_L: BP_Flip_L=Array(BM_Flip_L, LM_GI_sling_L_Flip_L)
Dim BP_Flip_LU: BP_Flip_LU=Array(BM_Flip_LU, LM_GI_sling_L_Flip_LU, LM_GI_sling_R_Flip_LU)
Dim BP_Flip_R: BP_Flip_R=Array(BM_Flip_R, LM_GI_sling_R_Flip_R)
Dim BP_Flip_RU: BP_Flip_RU=Array(BM_Flip_RU, LM_GI_sling_L_Flip_RU, LM_GI_sling_R_Flip_RU)
Dim BP_GateFlap: BP_GateFlap=Array(BM_GateFlap)
Dim BP_Overlay: BP_Overlay=Array(BM_Overlay, LM_L_L_Bmp_Y_Overlay, LM_GI_lane_top1_Overlay, LM_GI_lane_top4_Overlay, LM_GI_mid_side_L0_Overlay, LM_GI_mid_side_L1_Overlay, LM_GI_mid_side_R0_Overlay, LM_GI_mid_side_R1_Overlay, LM_GI_sling_L_Overlay, LM_GI_sling_R_Overlay, LM_GI_top_center_L_Overlay, LM_GI_top_center_R_Overlay, LM_GI_top_side_L_Overlay, LM_GI_top_side_R_Overlay)
Dim BP_Parts: BP_Parts=Array(BM_Parts, LM_L_L_Bmp_B_Parts, LM_L_L_Bmp_Y_Parts, LM_L_L_Bmp_G_Parts, LM_L_I_lane_mid_L_Parts, LM_L_I_lane_mid_R_Parts, LM_L_I_lane_out_L_Parts, LM_L_I_lane_out_R_Parts, LM_GI_lane_top1_Parts, LM_GI_lane_top2_Parts, LM_GI_lane_top3_Parts, LM_GI_lane_top4_Parts, LM_GI_mid_side_L0_Parts, LM_GI_mid_side_L1_Parts, LM_GI_mid_side_R0_Parts, LM_GI_mid_side_R1_Parts, LM_GI_sling_L_Parts, LM_GI_sling_R_Parts, LM_GI_top_center_L_Parts, LM_GI_top_center_R_Parts, LM_GI_top_side_L_Parts, LM_GI_top_side_R_Parts)
Dim BP_Playfield: BP_Playfield=Array(BM_Playfield, LM_L_I_bank_B_Playfield, LM_L_I_bank_G_Playfield, LM_L_I_bank_Y_Playfield, LM_L_I_lane_B_Playfield, LM_L_I_lane_G_Playfield, LM_L_I_lane_Y_Playfield, LM_L_I_lane_mid_L_Playfield, LM_L_I_lane_mid_R_Playfield, LM_L_I_lane_out_L_Playfield, LM_L_I_lane_out_R_Playfield, LM_L_I_lane_top_C_Playfield, LM_L_I_special_L_Playfield, LM_L_I_special_R_Playfield, LM_GI_lane_top1_Playfield, LM_GI_lane_top2_Playfield, LM_GI_lane_top3_Playfield, LM_GI_lane_top4_Playfield, LM_GI_mid_side_L0_Playfield, LM_GI_mid_side_L1_Playfield, LM_GI_mid_side_R0_Playfield, LM_GI_mid_side_R1_Playfield, LM_GI_sling_L_Playfield, LM_GI_sling_R_Playfield, LM_GI_top_center_L_Playfield, LM_GI_top_center_R_Playfield, LM_GI_top_side_L_Playfield, LM_GI_top_side_R_Playfield)
Dim BP_RailL: BP_RailL=Array(BM_RailL)
Dim BP_RailR: BP_RailR=Array(BM_RailR)
Dim BP_dt_b1: BP_dt_b1=Array(BM_dt_b1, LM_GI_mid_side_L0_dt_b1, LM_GI_mid_side_L1_dt_b1)
Dim BP_dt_b2: BP_dt_b2=Array(BM_dt_b2, LM_GI_mid_side_L0_dt_b2, LM_GI_mid_side_L1_dt_b2)
Dim BP_dt_b3: BP_dt_b3=Array(BM_dt_b3, LM_GI_mid_side_L0_dt_b3, LM_GI_mid_side_L1_dt_b3)
Dim BP_dt_b4: BP_dt_b4=Array(BM_dt_b4, LM_GI_mid_side_L0_dt_b4, LM_GI_mid_side_L1_dt_b4)
Dim BP_dt_b5: BP_dt_b5=Array(BM_dt_b5, LM_GI_mid_side_L0_dt_b5, LM_GI_mid_side_L1_dt_b5)
Dim BP_dt_g1: BP_dt_g1=Array(BM_dt_g1, LM_GI_mid_side_R0_dt_g1, LM_GI_mid_side_R1_dt_g1)
Dim BP_dt_g2: BP_dt_g2=Array(BM_dt_g2, LM_GI_mid_side_R0_dt_g2, LM_GI_mid_side_R1_dt_g2)
Dim BP_dt_g3: BP_dt_g3=Array(BM_dt_g3, LM_GI_mid_side_R0_dt_g3, LM_GI_mid_side_R1_dt_g3)
Dim BP_dt_g4: BP_dt_g4=Array(BM_dt_g4, LM_GI_mid_side_R0_dt_g4, LM_GI_mid_side_R1_dt_g4, LM_GI_top_side_R_dt_g4)
Dim BP_dt_g5: BP_dt_g5=Array(BM_dt_g5, LM_GI_mid_side_R0_dt_g5, LM_GI_mid_side_R1_dt_g5)
Dim BP_dt_y1: BP_dt_y1=Array(BM_dt_y1, LM_GI_lane_top3_dt_y1, LM_GI_mid_side_L0_dt_y1, LM_GI_mid_side_L1_dt_y1, LM_GI_mid_side_R1_dt_y1, LM_GI_top_center_L_dt_y1, LM_GI_top_center_R_dt_y1, LM_GI_top_side_L_dt_y1)
Dim BP_dt_y2: BP_dt_y2=Array(BM_dt_y2, LM_GI_lane_top1_dt_y2, LM_GI_mid_side_L1_dt_y2, LM_GI_mid_side_R1_dt_y2, LM_GI_top_center_L_dt_y2, LM_GI_top_center_R_dt_y2, LM_GI_top_side_L_dt_y2)
Dim BP_dt_y3: BP_dt_y3=Array(BM_dt_y3, LM_GI_lane_top1_dt_y3, LM_GI_lane_top4_dt_y3, LM_GI_mid_side_L1_dt_y3, LM_GI_mid_side_R1_dt_y3, LM_GI_top_center_L_dt_y3, LM_GI_top_center_R_dt_y3, LM_GI_top_side_L_dt_y3, LM_GI_top_side_R_dt_y3)
Dim BP_dt_y4: BP_dt_y4=Array(BM_dt_y4, LM_GI_lane_top4_dt_y4, LM_GI_mid_side_L1_dt_y4, LM_GI_mid_side_R1_dt_y4, LM_GI_top_center_L_dt_y4, LM_GI_top_center_R_dt_y4, LM_GI_top_side_R_dt_y4)
Dim BP_dt_y5: BP_dt_y5=Array(BM_dt_y5, LM_GI_lane_top2_dt_y5, LM_GI_mid_side_L1_dt_y5, LM_GI_mid_side_R0_dt_y5, LM_GI_mid_side_R1_dt_y5, LM_GI_top_center_L_dt_y5, LM_GI_top_center_R_dt_y5, LM_GI_top_side_R_dt_y5)
Dim BP_lockdownbar: BP_lockdownbar=Array(BM_lockdownbar)
Dim BP_rb_L_Mid_0: BP_rb_L_Mid_0=Array(BM_rb_L_Mid_0, LM_L_I_lane_mid_L_rb_L_Mid_0, LM_GI_mid_side_L0_rb_L_Mid_0, LM_GI_mid_side_L1_rb_L_Mid_0, LM_GI_top_center_L_rb_L_Mid_0, LM_GI_top_center_R_rb_L_Mid_0, LM_GI_top_side_L_rb_L_Mid_0)
Dim BP_rb_L_Mid_1: BP_rb_L_Mid_1=Array(BM_rb_L_Mid_1, LM_L_I_lane_mid_L_rb_L_Mid_1, LM_GI_mid_side_L0_rb_L_Mid_1, LM_GI_mid_side_L1_rb_L_Mid_1, LM_GI_top_center_L_rb_L_Mid_1, LM_GI_top_center_R_rb_L_Mid_1, LM_GI_top_side_L_rb_L_Mid_1)
Dim BP_rb_L_Mid_2: BP_rb_L_Mid_2=Array(BM_rb_L_Mid_2, LM_L_I_lane_mid_L_rb_L_Mid_2, LM_GI_mid_side_L0_rb_L_Mid_2, LM_GI_mid_side_L1_rb_L_Mid_2, LM_GI_top_center_L_rb_L_Mid_2, LM_GI_top_center_R_rb_L_Mid_2, LM_GI_top_side_L_rb_L_Mid_2)
Dim BP_rb_L_Sling_0: BP_rb_L_Sling_0=Array(BM_rb_L_Sling_0, LM_GI_mid_side_L0_rb_L_Sling_0, LM_GI_mid_side_L1_rb_L_Sling_0, LM_GI_sling_L_rb_L_Sling_0, LM_GI_sling_R_rb_L_Sling_0)
Dim BP_rb_L_Sling_1: BP_rb_L_Sling_1=Array(BM_rb_L_Sling_1, LM_GI_mid_side_L0_rb_L_Sling_1, LM_GI_mid_side_L1_rb_L_Sling_1, LM_GI_sling_L_rb_L_Sling_1, LM_GI_sling_R_rb_L_Sling_1)
Dim BP_rb_L_Sling_2: BP_rb_L_Sling_2=Array(BM_rb_L_Sling_2, LM_GI_mid_side_L0_rb_L_Sling_2, LM_GI_mid_side_L1_rb_L_Sling_2, LM_GI_sling_L_rb_L_Sling_2, LM_GI_sling_R_rb_L_Sling_2)
Dim BP_rb_L_Top0_0: BP_rb_L_Top0_0=Array(BM_rb_L_Top0_0, LM_GI_lane_top1_rb_L_Top0_0, LM_GI_top_center_L_rb_L_Top0_0, LM_GI_top_center_R_rb_L_Top0_0, LM_GI_top_side_L_rb_L_Top0_0)
Dim BP_rb_L_Top0_1: BP_rb_L_Top0_1=Array(BM_rb_L_Top0_1, LM_GI_lane_top1_rb_L_Top0_1, LM_GI_top_center_L_rb_L_Top0_1, LM_GI_top_center_R_rb_L_Top0_1, LM_GI_top_side_L_rb_L_Top0_1)
Dim BP_rb_L_Top0_2: BP_rb_L_Top0_2=Array(BM_rb_L_Top0_2, LM_GI_lane_top1_rb_L_Top0_2, LM_GI_top_center_L_rb_L_Top0_2, LM_GI_top_center_R_rb_L_Top0_2, LM_GI_top_side_L_rb_L_Top0_2)
Dim BP_rb_L_Top1_0: BP_rb_L_Top1_0=Array(BM_rb_L_Top1_0, LM_GI_lane_top1_rb_L_Top1_0, LM_GI_lane_top2_rb_L_Top1_0, LM_GI_top_center_L_rb_L_Top1_0, LM_GI_top_side_L_rb_L_Top1_0)
Dim BP_rb_L_Top1_1: BP_rb_L_Top1_1=Array(BM_rb_L_Top1_1, LM_GI_lane_top1_rb_L_Top1_1, LM_GI_lane_top2_rb_L_Top1_1, LM_GI_top_center_L_rb_L_Top1_1, LM_GI_top_side_L_rb_L_Top1_1)
Dim BP_rb_L_Top1_2: BP_rb_L_Top1_2=Array(BM_rb_L_Top1_2, LM_GI_lane_top1_rb_L_Top1_2, LM_GI_lane_top2_rb_L_Top1_2, LM_GI_top_center_L_rb_L_Top1_2, LM_GI_top_side_L_rb_L_Top1_2)
Dim BP_rb_L_Top2_0: BP_rb_L_Top2_0=Array(BM_rb_L_Top2_0, LM_GI_lane_top1_rb_L_Top2_0, LM_GI_mid_side_L0_rb_L_Top2_0, LM_GI_mid_side_L1_rb_L_Top2_0, LM_GI_top_center_L_rb_L_Top2_0, LM_GI_top_center_R_rb_L_Top2_0, LM_GI_top_side_L_rb_L_Top2_0)
Dim BP_rb_L_Top2_1: BP_rb_L_Top2_1=Array(BM_rb_L_Top2_1, LM_GI_lane_top1_rb_L_Top2_1, LM_GI_mid_side_L0_rb_L_Top2_1, LM_GI_mid_side_L1_rb_L_Top2_1, LM_GI_top_center_L_rb_L_Top2_1, LM_GI_top_center_R_rb_L_Top2_1, LM_GI_top_side_L_rb_L_Top2_1)
Dim BP_rb_L_Top2_2: BP_rb_L_Top2_2=Array(BM_rb_L_Top2_2, LM_GI_lane_top1_rb_L_Top2_2, LM_GI_mid_side_L0_rb_L_Top2_2, LM_GI_mid_side_L1_rb_L_Top2_2, LM_GI_top_center_L_rb_L_Top2_2, LM_GI_top_center_R_rb_L_Top2_2, LM_GI_top_side_L_rb_L_Top2_2)
Dim BP_rb_M_Sling_0: BP_rb_M_Sling_0=Array(BM_rb_M_Sling_0, LM_GI_lane_top1_rb_M_Sling_0, LM_GI_lane_top4_rb_M_Sling_0, LM_GI_mid_side_L1_rb_M_Sling_0, LM_GI_mid_side_R1_rb_M_Sling_0, LM_GI_top_center_L_rb_M_Sling_0, LM_GI_top_center_R_rb_M_Sling_0, LM_GI_top_side_L_rb_M_Sling_0, LM_GI_top_side_R_rb_M_Sling_0)
Dim BP_rb_M_Sling_1: BP_rb_M_Sling_1=Array(BM_rb_M_Sling_1, LM_GI_lane_top1_rb_M_Sling_1, LM_GI_lane_top4_rb_M_Sling_1, LM_GI_mid_side_L1_rb_M_Sling_1, LM_GI_mid_side_R1_rb_M_Sling_1, LM_GI_top_center_L_rb_M_Sling_1, LM_GI_top_center_R_rb_M_Sling_1, LM_GI_top_side_L_rb_M_Sling_1, LM_GI_top_side_R_rb_M_Sling_1)
Dim BP_rb_M_Sling_2: BP_rb_M_Sling_2=Array(BM_rb_M_Sling_2, LM_L_L_Bmp_Y_rb_M_Sling_2, LM_GI_lane_top1_rb_M_Sling_2, LM_GI_lane_top4_rb_M_Sling_2, LM_GI_mid_side_L1_rb_M_Sling_2, LM_GI_mid_side_R1_rb_M_Sling_2, LM_GI_top_center_L_rb_M_Sling_2, LM_GI_top_center_R_rb_M_Sling_2, LM_GI_top_side_L_rb_M_Sling_2, LM_GI_top_side_R_rb_M_Sling_2)
Dim BP_rb_R_Mid_0: BP_rb_R_Mid_0=Array(BM_rb_R_Mid_0, LM_L_I_lane_mid_R_rb_R_Mid_0, LM_GI_mid_side_R0_rb_R_Mid_0, LM_GI_mid_side_R1_rb_R_Mid_0, LM_GI_top_center_L_rb_R_Mid_0, LM_GI_top_center_R_rb_R_Mid_0, LM_GI_top_side_R_rb_R_Mid_0)
Dim BP_rb_R_Mid_1: BP_rb_R_Mid_1=Array(BM_rb_R_Mid_1, LM_L_I_lane_mid_R_rb_R_Mid_1, LM_GI_mid_side_R0_rb_R_Mid_1, LM_GI_mid_side_R1_rb_R_Mid_1, LM_GI_top_center_L_rb_R_Mid_1, LM_GI_top_center_R_rb_R_Mid_1, LM_GI_top_side_R_rb_R_Mid_1)
Dim BP_rb_R_Mid_2: BP_rb_R_Mid_2=Array(BM_rb_R_Mid_2, LM_L_I_lane_mid_R_rb_R_Mid_2, LM_GI_mid_side_R0_rb_R_Mid_2, LM_GI_mid_side_R1_rb_R_Mid_2, LM_GI_top_center_L_rb_R_Mid_2, LM_GI_top_center_R_rb_R_Mid_2, LM_GI_top_side_R_rb_R_Mid_2)
Dim BP_rb_R_Sling_0: BP_rb_R_Sling_0=Array(BM_rb_R_Sling_0, LM_GI_mid_side_R0_rb_R_Sling_0, LM_GI_mid_side_R1_rb_R_Sling_0, LM_GI_sling_L_rb_R_Sling_0, LM_GI_sling_R_rb_R_Sling_0)
Dim BP_rb_R_Sling_1: BP_rb_R_Sling_1=Array(BM_rb_R_Sling_1, LM_GI_mid_side_R0_rb_R_Sling_1, LM_GI_mid_side_R1_rb_R_Sling_1, LM_GI_sling_L_rb_R_Sling_1, LM_GI_sling_R_rb_R_Sling_1)
Dim BP_rb_R_Sling_2: BP_rb_R_Sling_2=Array(BM_rb_R_Sling_2, LM_GI_mid_side_R0_rb_R_Sling_2, LM_GI_mid_side_R1_rb_R_Sling_2, LM_GI_sling_L_rb_R_Sling_2, LM_GI_sling_R_rb_R_Sling_2)
Dim BP_rb_R_Top0_0: BP_rb_R_Top0_0=Array(BM_rb_R_Top0_0, LM_GI_lane_top4_rb_R_Top0_0, LM_GI_top_center_L_rb_R_Top0_0, LM_GI_top_center_R_rb_R_Top0_0)
Dim BP_rb_R_Top0_1: BP_rb_R_Top0_1=Array(BM_rb_R_Top0_1, LM_GI_lane_top4_rb_R_Top0_1, LM_GI_top_center_L_rb_R_Top0_1, LM_GI_top_center_R_rb_R_Top0_1)
Dim BP_rb_R_Top0_2: BP_rb_R_Top0_2=Array(BM_rb_R_Top0_2, LM_GI_lane_top4_rb_R_Top0_2, LM_GI_top_center_R_rb_R_Top0_2, LM_GI_top_side_R_rb_R_Top0_2)
Dim BP_rb_R_Top1_0: BP_rb_R_Top1_0=Array(BM_rb_R_Top1_0, LM_GI_lane_top4_rb_R_Top1_0, LM_GI_top_center_R_rb_R_Top1_0, LM_GI_top_side_R_rb_R_Top1_0)
Dim BP_rb_R_Top1_1: BP_rb_R_Top1_1=Array(BM_rb_R_Top1_1, LM_GI_lane_top4_rb_R_Top1_1, LM_GI_top_center_R_rb_R_Top1_1, LM_GI_top_side_R_rb_R_Top1_1)
Dim BP_rb_R_Top1_2: BP_rb_R_Top1_2=Array(BM_rb_R_Top1_2, LM_GI_lane_top3_rb_R_Top1_2, LM_GI_lane_top4_rb_R_Top1_2, LM_GI_top_center_R_rb_R_Top1_2, LM_GI_top_side_R_rb_R_Top1_2)
Dim BP_rb_R_Top2_0: BP_rb_R_Top2_0=Array(BM_rb_R_Top2_0, LM_GI_lane_top4_rb_R_Top2_0, LM_GI_mid_side_R0_rb_R_Top2_0, LM_GI_mid_side_R1_rb_R_Top2_0, LM_GI_top_center_L_rb_R_Top2_0, LM_GI_top_center_R_rb_R_Top2_0, LM_GI_top_side_R_rb_R_Top2_0)
Dim BP_rb_R_Top2_1: BP_rb_R_Top2_1=Array(BM_rb_R_Top2_1, LM_GI_lane_top4_rb_R_Top2_1, LM_GI_mid_side_R0_rb_R_Top2_1, LM_GI_mid_side_R1_rb_R_Top2_1, LM_GI_top_center_R_rb_R_Top2_1, LM_GI_top_side_R_rb_R_Top2_1)
Dim BP_rb_R_Top2_2: BP_rb_R_Top2_2=Array(BM_rb_R_Top2_2, LM_GI_lane_top4_rb_R_Top2_2, LM_GI_mid_side_R0_rb_R_Top2_2, LM_GI_mid_side_R1_rb_R_Top2_2, LM_GI_top_center_L_rb_R_Top2_2, LM_GI_top_center_R_rb_R_Top2_2, LM_GI_top_side_R_rb_R_Top2_2)
Dim BP_tg_in_L: BP_tg_in_L=Array(BM_tg_in_L, LM_GI_sling_L_tg_in_L)
Dim BP_tg_in_R: BP_tg_in_R=Array(BM_tg_in_R, LM_GI_sling_R_tg_in_R)
Dim BP_tg_med_R: BP_tg_med_R=Array(BM_tg_med_R, LM_GI_mid_side_R0_tg_med_R, LM_GI_mid_side_R1_tg_med_R)
Dim BP_tg_mid_L: BP_tg_mid_L=Array(BM_tg_mid_L, LM_GI_mid_side_L0_tg_mid_L, LM_GI_mid_side_L1_tg_mid_L)
Dim BP_tg_out_L: BP_tg_out_L=Array(BM_tg_out_L)
Dim BP_tg_out_R: BP_tg_out_R=Array(BM_tg_out_R)
Dim BP_tg_top_B: BP_tg_top_B=Array(BM_tg_top_B, LM_GI_lane_top1_tg_top_B, LM_GI_lane_top2_tg_top_B)
Dim BP_tg_top_G: BP_tg_top_G=Array(BM_tg_top_G, LM_GI_lane_top3_tg_top_G, LM_GI_lane_top4_tg_top_G)
Dim BP_tg_top_Y: BP_tg_top_Y=Array(BM_tg_top_Y, LM_GI_lane_top2_tg_top_Y, LM_GI_lane_top3_tg_top_Y)
' Arrays per lighting scenario
Dim BL_GI_lane_top1: BL_GI_lane_top1=Array(LM_GI_lane_top1_BmpRing_B, LM_GI_lane_top1_BmpRing_G, LM_GI_lane_top1_BmpRing_Y, LM_GI_lane_top1_BmpSkirt_Y, LM_GI_lane_top1_Overlay, LM_GI_lane_top1_Parts, LM_GI_lane_top1_Playfield, LM_GI_lane_top1_dt_y2, LM_GI_lane_top1_dt_y3, LM_GI_lane_top1_rb_L_Top0_0, LM_GI_lane_top1_rb_L_Top0_1, LM_GI_lane_top1_rb_L_Top0_2, LM_GI_lane_top1_rb_L_Top1_0, LM_GI_lane_top1_rb_L_Top1_1, LM_GI_lane_top1_rb_L_Top1_2, LM_GI_lane_top1_rb_L_Top2_0, LM_GI_lane_top1_rb_L_Top2_1, LM_GI_lane_top1_rb_L_Top2_2, LM_GI_lane_top1_rb_M_Sling_0, LM_GI_lane_top1_rb_M_Sling_1, LM_GI_lane_top1_rb_M_Sling_2, LM_GI_lane_top1_tg_top_B)
Dim BL_GI_lane_top2: BL_GI_lane_top2=Array(LM_GI_lane_top2_BmpRing_B, LM_GI_lane_top2_BmpRing_G, LM_GI_lane_top2_BmpRing_Y, LM_GI_lane_top2_BmpSkirt_Y, LM_GI_lane_top2_Parts, LM_GI_lane_top2_Playfield, LM_GI_lane_top2_dt_y5, LM_GI_lane_top2_rb_L_Top1_0, LM_GI_lane_top2_rb_L_Top1_1, LM_GI_lane_top2_rb_L_Top1_2, LM_GI_lane_top2_tg_top_B, LM_GI_lane_top2_tg_top_Y)
Dim BL_GI_lane_top3: BL_GI_lane_top3=Array(LM_GI_lane_top3_BmpRing_B, LM_GI_lane_top3_BmpRing_G, LM_GI_lane_top3_BmpRing_Y, LM_GI_lane_top3_BmpSkirt_G, LM_GI_lane_top3_BmpSkirt_Y, LM_GI_lane_top3_Parts, LM_GI_lane_top3_Playfield, LM_GI_lane_top3_dt_y1, LM_GI_lane_top3_rb_R_Top1_2, LM_GI_lane_top3_tg_top_G, LM_GI_lane_top3_tg_top_Y)
Dim BL_GI_lane_top4: BL_GI_lane_top4=Array(LM_GI_lane_top4_BmpRing_G, LM_GI_lane_top4_BmpRing_Y, LM_GI_lane_top4_BmpSkirt_G, LM_GI_lane_top4_BmpSkirt_Y, LM_GI_lane_top4_Overlay, LM_GI_lane_top4_Parts, LM_GI_lane_top4_Playfield, LM_GI_lane_top4_dt_y3, LM_GI_lane_top4_dt_y4, LM_GI_lane_top4_rb_M_Sling_0, LM_GI_lane_top4_rb_M_Sling_1, LM_GI_lane_top4_rb_M_Sling_2, LM_GI_lane_top4_rb_R_Top0_0, LM_GI_lane_top4_rb_R_Top0_1, LM_GI_lane_top4_rb_R_Top0_2, LM_GI_lane_top4_rb_R_Top1_0, LM_GI_lane_top4_rb_R_Top1_1, LM_GI_lane_top4_rb_R_Top1_2, LM_GI_lane_top4_rb_R_Top2_0, LM_GI_lane_top4_rb_R_Top2_1, LM_GI_lane_top4_rb_R_Top2_2, LM_GI_lane_top4_tg_top_G)
Dim BL_GI_mid_side_L0: BL_GI_mid_side_L0=Array(LM_GI_mid_side_L0_Overlay, LM_GI_mid_side_L0_Parts, LM_GI_mid_side_L0_Playfield, LM_GI_mid_side_L0_dt_b1, LM_GI_mid_side_L0_dt_b2, LM_GI_mid_side_L0_dt_b3, LM_GI_mid_side_L0_dt_b4, LM_GI_mid_side_L0_dt_b5, LM_GI_mid_side_L0_dt_y1, LM_GI_mid_side_L0_rb_L_Mid_0, LM_GI_mid_side_L0_rb_L_Mid_1, LM_GI_mid_side_L0_rb_L_Mid_2, LM_GI_mid_side_L0_rb_L_Sling_0, LM_GI_mid_side_L0_rb_L_Sling_1, LM_GI_mid_side_L0_rb_L_Sling_2, LM_GI_mid_side_L0_rb_L_Top2_0, LM_GI_mid_side_L0_rb_L_Top2_1, LM_GI_mid_side_L0_rb_L_Top2_2, LM_GI_mid_side_L0_tg_mid_L)
Dim BL_GI_mid_side_L1: BL_GI_mid_side_L1=Array(LM_GI_mid_side_L1_BmpSkirt_Y, LM_GI_mid_side_L1_Overlay, LM_GI_mid_side_L1_Parts, LM_GI_mid_side_L1_Playfield, LM_GI_mid_side_L1_dt_b1, LM_GI_mid_side_L1_dt_b2, LM_GI_mid_side_L1_dt_b3, LM_GI_mid_side_L1_dt_b4, LM_GI_mid_side_L1_dt_b5, LM_GI_mid_side_L1_dt_y1, LM_GI_mid_side_L1_dt_y2, LM_GI_mid_side_L1_dt_y3, LM_GI_mid_side_L1_dt_y4, LM_GI_mid_side_L1_dt_y5, LM_GI_mid_side_L1_rb_L_Mid_0, LM_GI_mid_side_L1_rb_L_Mid_1, LM_GI_mid_side_L1_rb_L_Mid_2, LM_GI_mid_side_L1_rb_L_Sling_0, LM_GI_mid_side_L1_rb_L_Sling_1, LM_GI_mid_side_L1_rb_L_Sling_2, LM_GI_mid_side_L1_rb_L_Top2_0, LM_GI_mid_side_L1_rb_L_Top2_1, LM_GI_mid_side_L1_rb_L_Top2_2, LM_GI_mid_side_L1_rb_M_Sling_0, LM_GI_mid_side_L1_rb_M_Sling_1, LM_GI_mid_side_L1_rb_M_Sling_2, LM_GI_mid_side_L1_tg_mid_L)
Dim BL_GI_mid_side_R0: BL_GI_mid_side_R0=Array(LM_GI_mid_side_R0_Overlay, LM_GI_mid_side_R0_Parts, LM_GI_mid_side_R0_Playfield, LM_GI_mid_side_R0_dt_g1, LM_GI_mid_side_R0_dt_g2, LM_GI_mid_side_R0_dt_g3, LM_GI_mid_side_R0_dt_g4, LM_GI_mid_side_R0_dt_g5, LM_GI_mid_side_R0_dt_y5, LM_GI_mid_side_R0_rb_R_Mid_0, LM_GI_mid_side_R0_rb_R_Mid_1, LM_GI_mid_side_R0_rb_R_Mid_2, LM_GI_mid_side_R0_rb_R_Sling_0, LM_GI_mid_side_R0_rb_R_Sling_1, LM_GI_mid_side_R0_rb_R_Sling_2, LM_GI_mid_side_R0_rb_R_Top2_0, LM_GI_mid_side_R0_rb_R_Top2_1, LM_GI_mid_side_R0_rb_R_Top2_2, LM_GI_mid_side_R0_tg_med_R)
Dim BL_GI_mid_side_R1: BL_GI_mid_side_R1=Array(LM_GI_mid_side_R1_BmpSkirt_Y, LM_GI_mid_side_R1_Overlay, LM_GI_mid_side_R1_Parts, LM_GI_mid_side_R1_Playfield, LM_GI_mid_side_R1_dt_g1, LM_GI_mid_side_R1_dt_g2, LM_GI_mid_side_R1_dt_g3, LM_GI_mid_side_R1_dt_g4, LM_GI_mid_side_R1_dt_g5, LM_GI_mid_side_R1_dt_y1, LM_GI_mid_side_R1_dt_y2, LM_GI_mid_side_R1_dt_y3, LM_GI_mid_side_R1_dt_y4, LM_GI_mid_side_R1_dt_y5, LM_GI_mid_side_R1_rb_M_Sling_0, LM_GI_mid_side_R1_rb_M_Sling_1, LM_GI_mid_side_R1_rb_M_Sling_2, LM_GI_mid_side_R1_rb_R_Mid_0, LM_GI_mid_side_R1_rb_R_Mid_1, LM_GI_mid_side_R1_rb_R_Mid_2, LM_GI_mid_side_R1_rb_R_Sling_0, LM_GI_mid_side_R1_rb_R_Sling_1, LM_GI_mid_side_R1_rb_R_Sling_2, LM_GI_mid_side_R1_rb_R_Top2_0, LM_GI_mid_side_R1_rb_R_Top2_1, LM_GI_mid_side_R1_rb_R_Top2_2, LM_GI_mid_side_R1_tg_med_R)
Dim BL_GI_sling_L: BL_GI_sling_L=Array(LM_GI_sling_L_Flip_L, LM_GI_sling_L_Flip_LU, LM_GI_sling_L_Flip_RU, LM_GI_sling_L_Overlay, LM_GI_sling_L_Parts, LM_GI_sling_L_Playfield, LM_GI_sling_L_rb_L_Sling_0, LM_GI_sling_L_rb_L_Sling_1, LM_GI_sling_L_rb_L_Sling_2, LM_GI_sling_L_rb_R_Sling_0, LM_GI_sling_L_rb_R_Sling_1, LM_GI_sling_L_rb_R_Sling_2, LM_GI_sling_L_tg_in_L)
Dim BL_GI_sling_R: BL_GI_sling_R=Array(LM_GI_sling_R_Flip_LU, LM_GI_sling_R_Flip_R, LM_GI_sling_R_Flip_RU, LM_GI_sling_R_Overlay, LM_GI_sling_R_Parts, LM_GI_sling_R_Playfield, LM_GI_sling_R_rb_L_Sling_0, LM_GI_sling_R_rb_L_Sling_1, LM_GI_sling_R_rb_L_Sling_2, LM_GI_sling_R_rb_R_Sling_0, LM_GI_sling_R_rb_R_Sling_1, LM_GI_sling_R_rb_R_Sling_2, LM_GI_sling_R_tg_in_R)
Dim BL_GI_top_center_L: BL_GI_top_center_L=Array(LM_GI_top_center_L_BmpRing_B, LM_GI_top_center_L_BmpRing_G, LM_GI_top_center_L_BmpRing_Y, LM_GI_top_center_L_BmpSkirt_B, LM_GI_top_center_L_BmpSkirt_Y, LM_GI_top_center_L_Overlay, LM_GI_top_center_L_Parts, LM_GI_top_center_L_Playfield, LM_GI_top_center_L_dt_y1, LM_GI_top_center_L_dt_y2, LM_GI_top_center_L_dt_y3, LM_GI_top_center_L_dt_y4, LM_GI_top_center_L_dt_y5, LM_GI_top_center_L_rb_L_Mid_0, LM_GI_top_center_L_rb_L_Mid_1, LM_GI_top_center_L_rb_L_Mid_2, LM_GI_top_center_L_rb_L_Top0_0, LM_GI_top_center_L_rb_L_Top0_1, LM_GI_top_center_L_rb_L_Top0_2, LM_GI_top_center_L_rb_L_Top1_0, LM_GI_top_center_L_rb_L_Top1_1, LM_GI_top_center_L_rb_L_Top1_2, LM_GI_top_center_L_rb_L_Top2_0, LM_GI_top_center_L_rb_L_Top2_1, LM_GI_top_center_L_rb_L_Top2_2, LM_GI_top_center_L_rb_M_Sling_0, LM_GI_top_center_L_rb_M_Sling_1, LM_GI_top_center_L_rb_M_Sling_2, LM_GI_top_center_L_rb_R_Mid_0, LM_GI_top_center_L_rb_R_Mid_1, LM_GI_top_center_L_rb_R_Mid_2, LM_GI_top_center_L_rb_R_Top0_0, _
  LM_GI_top_center_L_rb_R_Top0_1, LM_GI_top_center_L_rb_R_Top2_0, LM_GI_top_center_L_rb_R_Top2_2)
Dim BL_GI_top_center_R: BL_GI_top_center_R=Array(LM_GI_top_center_R_BmpRing_B, LM_GI_top_center_R_BmpRing_G, LM_GI_top_center_R_BmpRing_Y, LM_GI_top_center_R_BmpSkirt_G, LM_GI_top_center_R_BmpSkirt_Y, LM_GI_top_center_R_Overlay, LM_GI_top_center_R_Parts, LM_GI_top_center_R_Playfield, LM_GI_top_center_R_dt_y1, LM_GI_top_center_R_dt_y2, LM_GI_top_center_R_dt_y3, LM_GI_top_center_R_dt_y4, LM_GI_top_center_R_dt_y5, LM_GI_top_center_R_rb_L_Mid_0, LM_GI_top_center_R_rb_L_Mid_1, LM_GI_top_center_R_rb_L_Mid_2, LM_GI_top_center_R_rb_L_Top0_0, LM_GI_top_center_R_rb_L_Top0_1, LM_GI_top_center_R_rb_L_Top0_2, LM_GI_top_center_R_rb_L_Top2_0, LM_GI_top_center_R_rb_L_Top2_1, LM_GI_top_center_R_rb_L_Top2_2, LM_GI_top_center_R_rb_M_Sling_0, LM_GI_top_center_R_rb_M_Sling_1, LM_GI_top_center_R_rb_M_Sling_2, LM_GI_top_center_R_rb_R_Mid_0, LM_GI_top_center_R_rb_R_Mid_1, LM_GI_top_center_R_rb_R_Mid_2, LM_GI_top_center_R_rb_R_Top0_0, LM_GI_top_center_R_rb_R_Top0_1, LM_GI_top_center_R_rb_R_Top0_2, LM_GI_top_center_R_rb_R_Top1_0, _
  LM_GI_top_center_R_rb_R_Top1_1, LM_GI_top_center_R_rb_R_Top1_2, LM_GI_top_center_R_rb_R_Top2_0, LM_GI_top_center_R_rb_R_Top2_1, LM_GI_top_center_R_rb_R_Top2_2)
Dim BL_GI_top_side_L: BL_GI_top_side_L=Array(LM_GI_top_side_L_BmpRing_B, LM_GI_top_side_L_BmpRing_Y, LM_GI_top_side_L_BmpSkirt_Y, LM_GI_top_side_L_Overlay, LM_GI_top_side_L_Parts, LM_GI_top_side_L_Playfield, LM_GI_top_side_L_dt_y1, LM_GI_top_side_L_dt_y2, LM_GI_top_side_L_dt_y3, LM_GI_top_side_L_rb_L_Mid_0, LM_GI_top_side_L_rb_L_Mid_1, LM_GI_top_side_L_rb_L_Mid_2, LM_GI_top_side_L_rb_L_Top0_0, LM_GI_top_side_L_rb_L_Top0_1, LM_GI_top_side_L_rb_L_Top0_2, LM_GI_top_side_L_rb_L_Top1_0, LM_GI_top_side_L_rb_L_Top1_1, LM_GI_top_side_L_rb_L_Top1_2, LM_GI_top_side_L_rb_L_Top2_0, LM_GI_top_side_L_rb_L_Top2_1, LM_GI_top_side_L_rb_L_Top2_2, LM_GI_top_side_L_rb_M_Sling_0, LM_GI_top_side_L_rb_M_Sling_1, LM_GI_top_side_L_rb_M_Sling_2)
Dim BL_GI_top_side_R: BL_GI_top_side_R=Array(LM_GI_top_side_R_BmpRing_G, LM_GI_top_side_R_BmpRing_Y, LM_GI_top_side_R_BmpSkirt_G, LM_GI_top_side_R_BmpSkirt_Y, LM_GI_top_side_R_Overlay, LM_GI_top_side_R_Parts, LM_GI_top_side_R_Playfield, LM_GI_top_side_R_dt_g4, LM_GI_top_side_R_dt_y3, LM_GI_top_side_R_dt_y4, LM_GI_top_side_R_dt_y5, LM_GI_top_side_R_rb_M_Sling_0, LM_GI_top_side_R_rb_M_Sling_1, LM_GI_top_side_R_rb_M_Sling_2, LM_GI_top_side_R_rb_R_Mid_0, LM_GI_top_side_R_rb_R_Mid_1, LM_GI_top_side_R_rb_R_Mid_2, LM_GI_top_side_R_rb_R_Top0_2, LM_GI_top_side_R_rb_R_Top1_0, LM_GI_top_side_R_rb_R_Top1_1, LM_GI_top_side_R_rb_R_Top1_2, LM_GI_top_side_R_rb_R_Top2_0, LM_GI_top_side_R_rb_R_Top2_1, LM_GI_top_side_R_rb_R_Top2_2)
Dim BL_L_I_bank_B: BL_L_I_bank_B=Array(LM_L_I_bank_B_Playfield)
Dim BL_L_I_bank_G: BL_L_I_bank_G=Array(LM_L_I_bank_G_Playfield)
Dim BL_L_I_bank_Y: BL_L_I_bank_Y=Array(LM_L_I_bank_Y_Playfield)
Dim BL_L_I_lane_B: BL_L_I_lane_B=Array(LM_L_I_lane_B_Playfield)
Dim BL_L_I_lane_G: BL_L_I_lane_G=Array(LM_L_I_lane_G_Playfield)
Dim BL_L_I_lane_Y: BL_L_I_lane_Y=Array(LM_L_I_lane_Y_Playfield)
Dim BL_L_I_lane_mid_L: BL_L_I_lane_mid_L=Array(LM_L_I_lane_mid_L_Parts, LM_L_I_lane_mid_L_Playfield, LM_L_I_lane_mid_L_rb_L_Mid_0, LM_L_I_lane_mid_L_rb_L_Mid_1, LM_L_I_lane_mid_L_rb_L_Mid_2)
Dim BL_L_I_lane_mid_R: BL_L_I_lane_mid_R=Array(LM_L_I_lane_mid_R_Parts, LM_L_I_lane_mid_R_Playfield, LM_L_I_lane_mid_R_rb_R_Mid_0, LM_L_I_lane_mid_R_rb_R_Mid_1, LM_L_I_lane_mid_R_rb_R_Mid_2)
Dim BL_L_I_lane_out_L: BL_L_I_lane_out_L=Array(LM_L_I_lane_out_L_Parts, LM_L_I_lane_out_L_Playfield)
Dim BL_L_I_lane_out_R: BL_L_I_lane_out_R=Array(LM_L_I_lane_out_R_Parts, LM_L_I_lane_out_R_Playfield)
Dim BL_L_I_lane_top_C: BL_L_I_lane_top_C=Array(LM_L_I_lane_top_C_Playfield)
Dim BL_L_I_special_L: BL_L_I_special_L=Array(LM_L_I_special_L_Playfield)
Dim BL_L_I_special_R: BL_L_I_special_R=Array(LM_L_I_special_R_Playfield)
Dim BL_L_L_Bmp_B: BL_L_L_Bmp_B=Array(LM_L_L_Bmp_B_BmpRing_B, LM_L_L_Bmp_B_BmpSkirt_B, LM_L_L_Bmp_B_Parts)
Dim BL_L_L_Bmp_G: BL_L_L_Bmp_G=Array(LM_L_L_Bmp_G_BmpRing_G, LM_L_L_Bmp_G_BmpSkirt_G, LM_L_L_Bmp_G_Parts)
Dim BL_L_L_Bmp_Y: BL_L_L_Bmp_Y=Array(LM_L_L_Bmp_Y_BmpRing_Y, LM_L_L_Bmp_Y_BmpSkirt_Y, LM_L_L_Bmp_Y_Overlay, LM_L_L_Bmp_Y_Parts, LM_L_L_Bmp_Y_rb_M_Sling_2)
Dim BL_World: BL_World=Array(BM_BmpRing_B, BM_BmpRing_G, BM_BmpRing_Y, BM_BmpSkirt_B, BM_BmpSkirt_G, BM_BmpSkirt_Y, BM_Flip_L, BM_Flip_LU, BM_Flip_R, BM_Flip_RU, BM_GateFlap, BM_Overlay, BM_Parts, BM_Playfield, BM_RailL, BM_RailR, BM_dt_b1, BM_dt_b2, BM_dt_b3, BM_dt_b4, BM_dt_b5, BM_dt_g1, BM_dt_g2, BM_dt_g3, BM_dt_g4, BM_dt_g5, BM_dt_y1, BM_dt_y2, BM_dt_y3, BM_dt_y4, BM_dt_y5, BM_lockdownbar, BM_rb_L_Mid_0, BM_rb_L_Mid_1, BM_rb_L_Mid_2, BM_rb_L_Sling_0, BM_rb_L_Sling_1, BM_rb_L_Sling_2, BM_rb_L_Top0_0, BM_rb_L_Top0_1, BM_rb_L_Top0_2, BM_rb_L_Top1_0, BM_rb_L_Top1_1, BM_rb_L_Top1_2, BM_rb_L_Top2_0, BM_rb_L_Top2_1, BM_rb_L_Top2_2, BM_rb_M_Sling_0, BM_rb_M_Sling_1, BM_rb_M_Sling_2, BM_rb_R_Mid_0, BM_rb_R_Mid_1, BM_rb_R_Mid_2, BM_rb_R_Sling_0, BM_rb_R_Sling_1, BM_rb_R_Sling_2, BM_rb_R_Top0_0, BM_rb_R_Top0_1, BM_rb_R_Top0_2, BM_rb_R_Top1_0, BM_rb_R_Top1_1, BM_rb_R_Top1_2, BM_rb_R_Top2_0, BM_rb_R_Top2_1, BM_rb_R_Top2_2, BM_tg_in_L, BM_tg_in_R, BM_tg_med_R, BM_tg_mid_L, BM_tg_out_L, BM_tg_out_R, BM_tg_top_B, _
  BM_tg_top_G, BM_tg_top_Y)
' Global arrays
Dim BG_Bakemap: BG_Bakemap=Array(BM_BmpRing_B, BM_BmpRing_G, BM_BmpRing_Y, BM_BmpSkirt_B, BM_BmpSkirt_G, BM_BmpSkirt_Y, BM_Flip_L, BM_Flip_LU, BM_Flip_R, BM_Flip_RU, BM_GateFlap, BM_Overlay, BM_Parts, BM_Playfield, BM_RailL, BM_RailR, BM_dt_b1, BM_dt_b2, BM_dt_b3, BM_dt_b4, BM_dt_b5, BM_dt_g1, BM_dt_g2, BM_dt_g3, BM_dt_g4, BM_dt_g5, BM_dt_y1, BM_dt_y2, BM_dt_y3, BM_dt_y4, BM_dt_y5, BM_lockdownbar, BM_rb_L_Mid_0, BM_rb_L_Mid_1, BM_rb_L_Mid_2, BM_rb_L_Sling_0, BM_rb_L_Sling_1, BM_rb_L_Sling_2, BM_rb_L_Top0_0, BM_rb_L_Top0_1, BM_rb_L_Top0_2, BM_rb_L_Top1_0, BM_rb_L_Top1_1, BM_rb_L_Top1_2, BM_rb_L_Top2_0, BM_rb_L_Top2_1, BM_rb_L_Top2_2, BM_rb_M_Sling_0, BM_rb_M_Sling_1, BM_rb_M_Sling_2, BM_rb_R_Mid_0, BM_rb_R_Mid_1, BM_rb_R_Mid_2, BM_rb_R_Sling_0, BM_rb_R_Sling_1, BM_rb_R_Sling_2, BM_rb_R_Top0_0, BM_rb_R_Top0_1, BM_rb_R_Top0_2, BM_rb_R_Top1_0, BM_rb_R_Top1_1, BM_rb_R_Top1_2, BM_rb_R_Top2_0, BM_rb_R_Top2_1, BM_rb_R_Top2_2, BM_tg_in_L, BM_tg_in_R, BM_tg_med_R, BM_tg_mid_L, BM_tg_out_L, BM_tg_out_R, BM_tg_top_B, _
  BM_tg_top_G, BM_tg_top_Y)
Dim BG_Lightmap: BG_Lightmap=Array(LM_GI_lane_top1_BmpRing_B, LM_GI_lane_top1_BmpRing_G, LM_GI_lane_top1_BmpRing_Y, LM_GI_lane_top1_BmpSkirt_Y, LM_GI_lane_top1_Overlay, LM_GI_lane_top1_Parts, LM_GI_lane_top1_Playfield, LM_GI_lane_top1_dt_y2, LM_GI_lane_top1_dt_y3, LM_GI_lane_top1_rb_L_Top0_0, LM_GI_lane_top1_rb_L_Top0_1, LM_GI_lane_top1_rb_L_Top0_2, LM_GI_lane_top1_rb_L_Top1_0, LM_GI_lane_top1_rb_L_Top1_1, LM_GI_lane_top1_rb_L_Top1_2, LM_GI_lane_top1_rb_L_Top2_0, LM_GI_lane_top1_rb_L_Top2_1, LM_GI_lane_top1_rb_L_Top2_2, LM_GI_lane_top1_rb_M_Sling_0, LM_GI_lane_top1_rb_M_Sling_1, LM_GI_lane_top1_rb_M_Sling_2, LM_GI_lane_top1_tg_top_B, LM_GI_lane_top2_BmpRing_B, LM_GI_lane_top2_BmpRing_G, LM_GI_lane_top2_BmpRing_Y, LM_GI_lane_top2_BmpSkirt_Y, LM_GI_lane_top2_Parts, LM_GI_lane_top2_Playfield, LM_GI_lane_top2_dt_y5, LM_GI_lane_top2_rb_L_Top1_0, LM_GI_lane_top2_rb_L_Top1_1, LM_GI_lane_top2_rb_L_Top1_2, LM_GI_lane_top2_tg_top_B, LM_GI_lane_top2_tg_top_Y, LM_GI_lane_top3_BmpRing_B, LM_GI_lane_top3_BmpRing_G, _
  LM_GI_lane_top3_BmpRing_Y, LM_GI_lane_top3_BmpSkirt_G, LM_GI_lane_top3_BmpSkirt_Y, LM_GI_lane_top3_Parts, LM_GI_lane_top3_Playfield, LM_GI_lane_top3_dt_y1, LM_GI_lane_top3_rb_R_Top1_2, LM_GI_lane_top3_tg_top_G, LM_GI_lane_top3_tg_top_Y, LM_GI_lane_top4_BmpRing_G, LM_GI_lane_top4_BmpRing_Y, LM_GI_lane_top4_BmpSkirt_G, LM_GI_lane_top4_BmpSkirt_Y, LM_GI_lane_top4_Overlay, LM_GI_lane_top4_Parts, LM_GI_lane_top4_Playfield, LM_GI_lane_top4_dt_y3, LM_GI_lane_top4_dt_y4, LM_GI_lane_top4_rb_M_Sling_0, LM_GI_lane_top4_rb_M_Sling_1, LM_GI_lane_top4_rb_M_Sling_2, LM_GI_lane_top4_rb_R_Top0_0, LM_GI_lane_top4_rb_R_Top0_1, LM_GI_lane_top4_rb_R_Top0_2, LM_GI_lane_top4_rb_R_Top1_0, LM_GI_lane_top4_rb_R_Top1_1, LM_GI_lane_top4_rb_R_Top1_2, LM_GI_lane_top4_rb_R_Top2_0, LM_GI_lane_top4_rb_R_Top2_1, LM_GI_lane_top4_rb_R_Top2_2, LM_GI_lane_top4_tg_top_G, LM_GI_mid_side_L0_Overlay, LM_GI_mid_side_L0_Parts, LM_GI_mid_side_L0_Playfield, LM_GI_mid_side_L0_dt_b1, LM_GI_mid_side_L0_dt_b2, LM_GI_mid_side_L0_dt_b3, _
  LM_GI_mid_side_L0_dt_b4, LM_GI_mid_side_L0_dt_b5, LM_GI_mid_side_L0_dt_y1, LM_GI_mid_side_L0_rb_L_Mid_0, LM_GI_mid_side_L0_rb_L_Mid_1, LM_GI_mid_side_L0_rb_L_Mid_2, LM_GI_mid_side_L0_rb_L_Sling_0, LM_GI_mid_side_L0_rb_L_Sling_1, LM_GI_mid_side_L0_rb_L_Sling_2, LM_GI_mid_side_L0_rb_L_Top2_0, LM_GI_mid_side_L0_rb_L_Top2_1, LM_GI_mid_side_L0_rb_L_Top2_2, LM_GI_mid_side_L0_tg_mid_L, LM_GI_mid_side_L1_BmpSkirt_Y, LM_GI_mid_side_L1_Overlay, LM_GI_mid_side_L1_Parts, LM_GI_mid_side_L1_Playfield, LM_GI_mid_side_L1_dt_b1, LM_GI_mid_side_L1_dt_b2, LM_GI_mid_side_L1_dt_b3, LM_GI_mid_side_L1_dt_b4, LM_GI_mid_side_L1_dt_b5, LM_GI_mid_side_L1_dt_y1, LM_GI_mid_side_L1_dt_y2, LM_GI_mid_side_L1_dt_y3, LM_GI_mid_side_L1_dt_y4, LM_GI_mid_side_L1_dt_y5, LM_GI_mid_side_L1_rb_L_Mid_0, LM_GI_mid_side_L1_rb_L_Mid_1, LM_GI_mid_side_L1_rb_L_Mid_2, LM_GI_mid_side_L1_rb_L_Sling_0, LM_GI_mid_side_L1_rb_L_Sling_1, LM_GI_mid_side_L1_rb_L_Sling_2, LM_GI_mid_side_L1_rb_L_Top2_0, LM_GI_mid_side_L1_rb_L_Top2_1, LM_GI_mid_side_L1_rb_L_Top2_2, _
  LM_GI_mid_side_L1_rb_M_Sling_0, LM_GI_mid_side_L1_rb_M_Sling_1, LM_GI_mid_side_L1_rb_M_Sling_2, LM_GI_mid_side_L1_tg_mid_L, LM_GI_mid_side_R0_Overlay, LM_GI_mid_side_R0_Parts, LM_GI_mid_side_R0_Playfield, LM_GI_mid_side_R0_dt_g1, LM_GI_mid_side_R0_dt_g2, LM_GI_mid_side_R0_dt_g3, LM_GI_mid_side_R0_dt_g4, LM_GI_mid_side_R0_dt_g5, LM_GI_mid_side_R0_dt_y5, LM_GI_mid_side_R0_rb_R_Mid_0, LM_GI_mid_side_R0_rb_R_Mid_1, LM_GI_mid_side_R0_rb_R_Mid_2, LM_GI_mid_side_R0_rb_R_Sling_0, LM_GI_mid_side_R0_rb_R_Sling_1, LM_GI_mid_side_R0_rb_R_Sling_2, LM_GI_mid_side_R0_rb_R_Top2_0, LM_GI_mid_side_R0_rb_R_Top2_1, LM_GI_mid_side_R0_rb_R_Top2_2, LM_GI_mid_side_R0_tg_med_R, LM_GI_mid_side_R1_BmpSkirt_Y, LM_GI_mid_side_R1_Overlay, LM_GI_mid_side_R1_Parts, LM_GI_mid_side_R1_Playfield, LM_GI_mid_side_R1_dt_g1, LM_GI_mid_side_R1_dt_g2, LM_GI_mid_side_R1_dt_g3, LM_GI_mid_side_R1_dt_g4, LM_GI_mid_side_R1_dt_g5, LM_GI_mid_side_R1_dt_y1, LM_GI_mid_side_R1_dt_y2, LM_GI_mid_side_R1_dt_y3, LM_GI_mid_side_R1_dt_y4, LM_GI_mid_side_R1_dt_y5, _
  LM_GI_mid_side_R1_rb_M_Sling_0, LM_GI_mid_side_R1_rb_M_Sling_1, LM_GI_mid_side_R1_rb_M_Sling_2, LM_GI_mid_side_R1_rb_R_Mid_0, LM_GI_mid_side_R1_rb_R_Mid_1, LM_GI_mid_side_R1_rb_R_Mid_2, LM_GI_mid_side_R1_rb_R_Sling_0, LM_GI_mid_side_R1_rb_R_Sling_1, LM_GI_mid_side_R1_rb_R_Sling_2, LM_GI_mid_side_R1_rb_R_Top2_0, LM_GI_mid_side_R1_rb_R_Top2_1, LM_GI_mid_side_R1_rb_R_Top2_2, LM_GI_mid_side_R1_tg_med_R, LM_GI_sling_L_Flip_L, LM_GI_sling_L_Flip_LU, LM_GI_sling_L_Flip_RU, LM_GI_sling_L_Overlay, LM_GI_sling_L_Parts, LM_GI_sling_L_Playfield, LM_GI_sling_L_rb_L_Sling_0, LM_GI_sling_L_rb_L_Sling_1, LM_GI_sling_L_rb_L_Sling_2, LM_GI_sling_L_rb_R_Sling_0, LM_GI_sling_L_rb_R_Sling_1, LM_GI_sling_L_rb_R_Sling_2, LM_GI_sling_L_tg_in_L, LM_GI_sling_R_Flip_LU, LM_GI_sling_R_Flip_R, LM_GI_sling_R_Flip_RU, LM_GI_sling_R_Overlay, LM_GI_sling_R_Parts, LM_GI_sling_R_Playfield, LM_GI_sling_R_rb_L_Sling_0, LM_GI_sling_R_rb_L_Sling_1, LM_GI_sling_R_rb_L_Sling_2, LM_GI_sling_R_rb_R_Sling_0, LM_GI_sling_R_rb_R_Sling_1, _
  LM_GI_sling_R_rb_R_Sling_2, LM_GI_sling_R_tg_in_R, LM_GI_top_center_L_BmpRing_B, LM_GI_top_center_L_BmpRing_G, LM_GI_top_center_L_BmpRing_Y, LM_GI_top_center_L_BmpSkirt_B, LM_GI_top_center_L_BmpSkirt_Y, LM_GI_top_center_L_Overlay, LM_GI_top_center_L_Parts, LM_GI_top_center_L_Playfield, LM_GI_top_center_L_dt_y1, LM_GI_top_center_L_dt_y2, LM_GI_top_center_L_dt_y3, LM_GI_top_center_L_dt_y4, LM_GI_top_center_L_dt_y5, LM_GI_top_center_L_rb_L_Mid_0, LM_GI_top_center_L_rb_L_Mid_1, LM_GI_top_center_L_rb_L_Mid_2, LM_GI_top_center_L_rb_L_Top0_0, LM_GI_top_center_L_rb_L_Top0_1, LM_GI_top_center_L_rb_L_Top0_2, LM_GI_top_center_L_rb_L_Top1_0, LM_GI_top_center_L_rb_L_Top1_1, LM_GI_top_center_L_rb_L_Top1_2, LM_GI_top_center_L_rb_L_Top2_0, LM_GI_top_center_L_rb_L_Top2_1, LM_GI_top_center_L_rb_L_Top2_2, LM_GI_top_center_L_rb_M_Sling_0, LM_GI_top_center_L_rb_M_Sling_1, LM_GI_top_center_L_rb_M_Sling_2, LM_GI_top_center_L_rb_R_Mid_0, LM_GI_top_center_L_rb_R_Mid_1, LM_GI_top_center_L_rb_R_Mid_2, LM_GI_top_center_L_rb_R_Top0_0, _
  LM_GI_top_center_L_rb_R_Top0_1, LM_GI_top_center_L_rb_R_Top2_0, LM_GI_top_center_L_rb_R_Top2_2, LM_GI_top_center_R_BmpRing_B, LM_GI_top_center_R_BmpRing_G, LM_GI_top_center_R_BmpRing_Y, LM_GI_top_center_R_BmpSkirt_G, LM_GI_top_center_R_BmpSkirt_Y, LM_GI_top_center_R_Overlay, LM_GI_top_center_R_Parts, LM_GI_top_center_R_Playfield, LM_GI_top_center_R_dt_y1, LM_GI_top_center_R_dt_y2, LM_GI_top_center_R_dt_y3, LM_GI_top_center_R_dt_y4, LM_GI_top_center_R_dt_y5, LM_GI_top_center_R_rb_L_Mid_0, LM_GI_top_center_R_rb_L_Mid_1, LM_GI_top_center_R_rb_L_Mid_2, LM_GI_top_center_R_rb_L_Top0_0, LM_GI_top_center_R_rb_L_Top0_1, LM_GI_top_center_R_rb_L_Top0_2, LM_GI_top_center_R_rb_L_Top2_0, LM_GI_top_center_R_rb_L_Top2_1, LM_GI_top_center_R_rb_L_Top2_2, LM_GI_top_center_R_rb_M_Sling_0, LM_GI_top_center_R_rb_M_Sling_1, LM_GI_top_center_R_rb_M_Sling_2, LM_GI_top_center_R_rb_R_Mid_0, LM_GI_top_center_R_rb_R_Mid_1, LM_GI_top_center_R_rb_R_Mid_2, LM_GI_top_center_R_rb_R_Top0_0, LM_GI_top_center_R_rb_R_Top0_1, _
  LM_GI_top_center_R_rb_R_Top0_2, LM_GI_top_center_R_rb_R_Top1_0, LM_GI_top_center_R_rb_R_Top1_1, LM_GI_top_center_R_rb_R_Top1_2, LM_GI_top_center_R_rb_R_Top2_0, LM_GI_top_center_R_rb_R_Top2_1, LM_GI_top_center_R_rb_R_Top2_2, LM_GI_top_side_L_BmpRing_B, LM_GI_top_side_L_BmpRing_Y, LM_GI_top_side_L_BmpSkirt_Y, LM_GI_top_side_L_Overlay, LM_GI_top_side_L_Parts, LM_GI_top_side_L_Playfield, LM_GI_top_side_L_dt_y1, LM_GI_top_side_L_dt_y2, LM_GI_top_side_L_dt_y3, LM_GI_top_side_L_rb_L_Mid_0, LM_GI_top_side_L_rb_L_Mid_1, LM_GI_top_side_L_rb_L_Mid_2, LM_GI_top_side_L_rb_L_Top0_0, LM_GI_top_side_L_rb_L_Top0_1, LM_GI_top_side_L_rb_L_Top0_2, LM_GI_top_side_L_rb_L_Top1_0, LM_GI_top_side_L_rb_L_Top1_1, LM_GI_top_side_L_rb_L_Top1_2, LM_GI_top_side_L_rb_L_Top2_0, LM_GI_top_side_L_rb_L_Top2_1, LM_GI_top_side_L_rb_L_Top2_2, LM_GI_top_side_L_rb_M_Sling_0, LM_GI_top_side_L_rb_M_Sling_1, LM_GI_top_side_L_rb_M_Sling_2, LM_GI_top_side_R_BmpRing_G, LM_GI_top_side_R_BmpRing_Y, LM_GI_top_side_R_BmpSkirt_G, LM_GI_top_side_R_BmpSkirt_Y, _
  LM_GI_top_side_R_Overlay, LM_GI_top_side_R_Parts, LM_GI_top_side_R_Playfield, LM_GI_top_side_R_dt_g4, LM_GI_top_side_R_dt_y3, LM_GI_top_side_R_dt_y4, LM_GI_top_side_R_dt_y5, LM_GI_top_side_R_rb_M_Sling_0, LM_GI_top_side_R_rb_M_Sling_1, LM_GI_top_side_R_rb_M_Sling_2, LM_GI_top_side_R_rb_R_Mid_0, LM_GI_top_side_R_rb_R_Mid_1, LM_GI_top_side_R_rb_R_Mid_2, LM_GI_top_side_R_rb_R_Top0_2, LM_GI_top_side_R_rb_R_Top1_0, LM_GI_top_side_R_rb_R_Top1_1, LM_GI_top_side_R_rb_R_Top1_2, LM_GI_top_side_R_rb_R_Top2_0, LM_GI_top_side_R_rb_R_Top2_1, LM_GI_top_side_R_rb_R_Top2_2, LM_L_I_bank_B_Playfield, LM_L_I_bank_G_Playfield, LM_L_I_bank_Y_Playfield, LM_L_I_lane_B_Playfield, LM_L_I_lane_G_Playfield, LM_L_I_lane_Y_Playfield, LM_L_I_lane_mid_L_Parts, LM_L_I_lane_mid_L_Playfield, LM_L_I_lane_mid_L_rb_L_Mid_0, LM_L_I_lane_mid_L_rb_L_Mid_1, LM_L_I_lane_mid_L_rb_L_Mid_2, LM_L_I_lane_mid_R_Parts, LM_L_I_lane_mid_R_Playfield, LM_L_I_lane_mid_R_rb_R_Mid_0, LM_L_I_lane_mid_R_rb_R_Mid_1, LM_L_I_lane_mid_R_rb_R_Mid_2, _
  LM_L_I_lane_out_L_Parts, LM_L_I_lane_out_L_Playfield, LM_L_I_lane_out_R_Parts, LM_L_I_lane_out_R_Playfield, LM_L_I_lane_top_C_Playfield, LM_L_I_special_L_Playfield, LM_L_I_special_R_Playfield, LM_L_L_Bmp_B_BmpRing_B, LM_L_L_Bmp_B_BmpSkirt_B, LM_L_L_Bmp_B_Parts, LM_L_L_Bmp_G_BmpRing_G, LM_L_L_Bmp_G_BmpSkirt_G, LM_L_L_Bmp_G_Parts, LM_L_L_Bmp_Y_BmpRing_Y, LM_L_L_Bmp_Y_BmpSkirt_Y, LM_L_L_Bmp_Y_Overlay, LM_L_L_Bmp_Y_Parts, LM_L_L_Bmp_Y_rb_M_Sling_2)
Dim BG_All: BG_All=Array(BM_BmpRing_B, BM_BmpRing_G, BM_BmpRing_Y, BM_BmpSkirt_B, BM_BmpSkirt_G, BM_BmpSkirt_Y, BM_Flip_L, BM_Flip_LU, BM_Flip_R, BM_Flip_RU, BM_GateFlap, BM_Overlay, BM_Parts, BM_Playfield, BM_RailL, BM_RailR, BM_dt_b1, BM_dt_b2, BM_dt_b3, BM_dt_b4, BM_dt_b5, BM_dt_g1, BM_dt_g2, BM_dt_g3, BM_dt_g4, BM_dt_g5, BM_dt_y1, BM_dt_y2, BM_dt_y3, BM_dt_y4, BM_dt_y5, BM_lockdownbar, BM_rb_L_Mid_0, BM_rb_L_Mid_1, BM_rb_L_Mid_2, BM_rb_L_Sling_0, BM_rb_L_Sling_1, BM_rb_L_Sling_2, BM_rb_L_Top0_0, BM_rb_L_Top0_1, BM_rb_L_Top0_2, BM_rb_L_Top1_0, BM_rb_L_Top1_1, BM_rb_L_Top1_2, BM_rb_L_Top2_0, BM_rb_L_Top2_1, BM_rb_L_Top2_2, BM_rb_M_Sling_0, BM_rb_M_Sling_1, BM_rb_M_Sling_2, BM_rb_R_Mid_0, BM_rb_R_Mid_1, BM_rb_R_Mid_2, BM_rb_R_Sling_0, BM_rb_R_Sling_1, BM_rb_R_Sling_2, BM_rb_R_Top0_0, BM_rb_R_Top0_1, BM_rb_R_Top0_2, BM_rb_R_Top1_0, BM_rb_R_Top1_1, BM_rb_R_Top1_2, BM_rb_R_Top2_0, BM_rb_R_Top2_1, BM_rb_R_Top2_2, BM_tg_in_L, BM_tg_in_R, BM_tg_med_R, BM_tg_mid_L, BM_tg_out_L, BM_tg_out_R, BM_tg_top_B, _
  BM_tg_top_G, BM_tg_top_Y, LM_GI_lane_top1_BmpRing_B, LM_GI_lane_top1_BmpRing_G, LM_GI_lane_top1_BmpRing_Y, LM_GI_lane_top1_BmpSkirt_Y, LM_GI_lane_top1_Overlay, LM_GI_lane_top1_Parts, LM_GI_lane_top1_Playfield, LM_GI_lane_top1_dt_y2, LM_GI_lane_top1_dt_y3, LM_GI_lane_top1_rb_L_Top0_0, LM_GI_lane_top1_rb_L_Top0_1, LM_GI_lane_top1_rb_L_Top0_2, LM_GI_lane_top1_rb_L_Top1_0, LM_GI_lane_top1_rb_L_Top1_1, LM_GI_lane_top1_rb_L_Top1_2, LM_GI_lane_top1_rb_L_Top2_0, LM_GI_lane_top1_rb_L_Top2_1, LM_GI_lane_top1_rb_L_Top2_2, LM_GI_lane_top1_rb_M_Sling_0, LM_GI_lane_top1_rb_M_Sling_1, LM_GI_lane_top1_rb_M_Sling_2, LM_GI_lane_top1_tg_top_B, LM_GI_lane_top2_BmpRing_B, LM_GI_lane_top2_BmpRing_G, LM_GI_lane_top2_BmpRing_Y, LM_GI_lane_top2_BmpSkirt_Y, LM_GI_lane_top2_Parts, LM_GI_lane_top2_Playfield, LM_GI_lane_top2_dt_y5, LM_GI_lane_top2_rb_L_Top1_0, LM_GI_lane_top2_rb_L_Top1_1, LM_GI_lane_top2_rb_L_Top1_2, LM_GI_lane_top2_tg_top_B, LM_GI_lane_top2_tg_top_Y, LM_GI_lane_top3_BmpRing_B, LM_GI_lane_top3_BmpRing_G, _
  LM_GI_lane_top3_BmpRing_Y, LM_GI_lane_top3_BmpSkirt_G, LM_GI_lane_top3_BmpSkirt_Y, LM_GI_lane_top3_Parts, LM_GI_lane_top3_Playfield, LM_GI_lane_top3_dt_y1, LM_GI_lane_top3_rb_R_Top1_2, LM_GI_lane_top3_tg_top_G, LM_GI_lane_top3_tg_top_Y, LM_GI_lane_top4_BmpRing_G, LM_GI_lane_top4_BmpRing_Y, LM_GI_lane_top4_BmpSkirt_G, LM_GI_lane_top4_BmpSkirt_Y, LM_GI_lane_top4_Overlay, LM_GI_lane_top4_Parts, LM_GI_lane_top4_Playfield, LM_GI_lane_top4_dt_y3, LM_GI_lane_top4_dt_y4, LM_GI_lane_top4_rb_M_Sling_0, LM_GI_lane_top4_rb_M_Sling_1, LM_GI_lane_top4_rb_M_Sling_2, LM_GI_lane_top4_rb_R_Top0_0, LM_GI_lane_top4_rb_R_Top0_1, LM_GI_lane_top4_rb_R_Top0_2, LM_GI_lane_top4_rb_R_Top1_0, LM_GI_lane_top4_rb_R_Top1_1, LM_GI_lane_top4_rb_R_Top1_2, LM_GI_lane_top4_rb_R_Top2_0, LM_GI_lane_top4_rb_R_Top2_1, LM_GI_lane_top4_rb_R_Top2_2, LM_GI_lane_top4_tg_top_G, LM_GI_mid_side_L0_Overlay, LM_GI_mid_side_L0_Parts, LM_GI_mid_side_L0_Playfield, LM_GI_mid_side_L0_dt_b1, LM_GI_mid_side_L0_dt_b2, LM_GI_mid_side_L0_dt_b3, _
  LM_GI_mid_side_L0_dt_b4, LM_GI_mid_side_L0_dt_b5, LM_GI_mid_side_L0_dt_y1, LM_GI_mid_side_L0_rb_L_Mid_0, LM_GI_mid_side_L0_rb_L_Mid_1, LM_GI_mid_side_L0_rb_L_Mid_2, LM_GI_mid_side_L0_rb_L_Sling_0, LM_GI_mid_side_L0_rb_L_Sling_1, LM_GI_mid_side_L0_rb_L_Sling_2, LM_GI_mid_side_L0_rb_L_Top2_0, LM_GI_mid_side_L0_rb_L_Top2_1, LM_GI_mid_side_L0_rb_L_Top2_2, LM_GI_mid_side_L0_tg_mid_L, LM_GI_mid_side_L1_BmpSkirt_Y, LM_GI_mid_side_L1_Overlay, LM_GI_mid_side_L1_Parts, LM_GI_mid_side_L1_Playfield, LM_GI_mid_side_L1_dt_b1, LM_GI_mid_side_L1_dt_b2, LM_GI_mid_side_L1_dt_b3, LM_GI_mid_side_L1_dt_b4, LM_GI_mid_side_L1_dt_b5, LM_GI_mid_side_L1_dt_y1, LM_GI_mid_side_L1_dt_y2, LM_GI_mid_side_L1_dt_y3, LM_GI_mid_side_L1_dt_y4, LM_GI_mid_side_L1_dt_y5, LM_GI_mid_side_L1_rb_L_Mid_0, LM_GI_mid_side_L1_rb_L_Mid_1, LM_GI_mid_side_L1_rb_L_Mid_2, LM_GI_mid_side_L1_rb_L_Sling_0, LM_GI_mid_side_L1_rb_L_Sling_1, LM_GI_mid_side_L1_rb_L_Sling_2, LM_GI_mid_side_L1_rb_L_Top2_0, LM_GI_mid_side_L1_rb_L_Top2_1, LM_GI_mid_side_L1_rb_L_Top2_2, _
  LM_GI_mid_side_L1_rb_M_Sling_0, LM_GI_mid_side_L1_rb_M_Sling_1, LM_GI_mid_side_L1_rb_M_Sling_2, LM_GI_mid_side_L1_tg_mid_L, LM_GI_mid_side_R0_Overlay, LM_GI_mid_side_R0_Parts, LM_GI_mid_side_R0_Playfield, LM_GI_mid_side_R0_dt_g1, LM_GI_mid_side_R0_dt_g2, LM_GI_mid_side_R0_dt_g3, LM_GI_mid_side_R0_dt_g4, LM_GI_mid_side_R0_dt_g5, LM_GI_mid_side_R0_dt_y5, LM_GI_mid_side_R0_rb_R_Mid_0, LM_GI_mid_side_R0_rb_R_Mid_1, LM_GI_mid_side_R0_rb_R_Mid_2, LM_GI_mid_side_R0_rb_R_Sling_0, LM_GI_mid_side_R0_rb_R_Sling_1, LM_GI_mid_side_R0_rb_R_Sling_2, LM_GI_mid_side_R0_rb_R_Top2_0, LM_GI_mid_side_R0_rb_R_Top2_1, LM_GI_mid_side_R0_rb_R_Top2_2, LM_GI_mid_side_R0_tg_med_R, LM_GI_mid_side_R1_BmpSkirt_Y, LM_GI_mid_side_R1_Overlay, LM_GI_mid_side_R1_Parts, LM_GI_mid_side_R1_Playfield, LM_GI_mid_side_R1_dt_g1, LM_GI_mid_side_R1_dt_g2, LM_GI_mid_side_R1_dt_g3, LM_GI_mid_side_R1_dt_g4, LM_GI_mid_side_R1_dt_g5, LM_GI_mid_side_R1_dt_y1, LM_GI_mid_side_R1_dt_y2, LM_GI_mid_side_R1_dt_y3, LM_GI_mid_side_R1_dt_y4, LM_GI_mid_side_R1_dt_y5, _
  LM_GI_mid_side_R1_rb_M_Sling_0, LM_GI_mid_side_R1_rb_M_Sling_1, LM_GI_mid_side_R1_rb_M_Sling_2, LM_GI_mid_side_R1_rb_R_Mid_0, LM_GI_mid_side_R1_rb_R_Mid_1, LM_GI_mid_side_R1_rb_R_Mid_2, LM_GI_mid_side_R1_rb_R_Sling_0, LM_GI_mid_side_R1_rb_R_Sling_1, LM_GI_mid_side_R1_rb_R_Sling_2, LM_GI_mid_side_R1_rb_R_Top2_0, LM_GI_mid_side_R1_rb_R_Top2_1, LM_GI_mid_side_R1_rb_R_Top2_2, LM_GI_mid_side_R1_tg_med_R, LM_GI_sling_L_Flip_L, LM_GI_sling_L_Flip_LU, LM_GI_sling_L_Flip_RU, LM_GI_sling_L_Overlay, LM_GI_sling_L_Parts, LM_GI_sling_L_Playfield, LM_GI_sling_L_rb_L_Sling_0, LM_GI_sling_L_rb_L_Sling_1, LM_GI_sling_L_rb_L_Sling_2, LM_GI_sling_L_rb_R_Sling_0, LM_GI_sling_L_rb_R_Sling_1, LM_GI_sling_L_rb_R_Sling_2, LM_GI_sling_L_tg_in_L, LM_GI_sling_R_Flip_LU, LM_GI_sling_R_Flip_R, LM_GI_sling_R_Flip_RU, LM_GI_sling_R_Overlay, LM_GI_sling_R_Parts, LM_GI_sling_R_Playfield, LM_GI_sling_R_rb_L_Sling_0, LM_GI_sling_R_rb_L_Sling_1, LM_GI_sling_R_rb_L_Sling_2, LM_GI_sling_R_rb_R_Sling_0, LM_GI_sling_R_rb_R_Sling_1, _
  LM_GI_sling_R_rb_R_Sling_2, LM_GI_sling_R_tg_in_R, LM_GI_top_center_L_BmpRing_B, LM_GI_top_center_L_BmpRing_G, LM_GI_top_center_L_BmpRing_Y, LM_GI_top_center_L_BmpSkirt_B, LM_GI_top_center_L_BmpSkirt_Y, LM_GI_top_center_L_Overlay, LM_GI_top_center_L_Parts, LM_GI_top_center_L_Playfield, LM_GI_top_center_L_dt_y1, LM_GI_top_center_L_dt_y2, LM_GI_top_center_L_dt_y3, LM_GI_top_center_L_dt_y4, LM_GI_top_center_L_dt_y5, LM_GI_top_center_L_rb_L_Mid_0, LM_GI_top_center_L_rb_L_Mid_1, LM_GI_top_center_L_rb_L_Mid_2, LM_GI_top_center_L_rb_L_Top0_0, LM_GI_top_center_L_rb_L_Top0_1, LM_GI_top_center_L_rb_L_Top0_2, LM_GI_top_center_L_rb_L_Top1_0, LM_GI_top_center_L_rb_L_Top1_1, LM_GI_top_center_L_rb_L_Top1_2, LM_GI_top_center_L_rb_L_Top2_0, LM_GI_top_center_L_rb_L_Top2_1, LM_GI_top_center_L_rb_L_Top2_2, LM_GI_top_center_L_rb_M_Sling_0, LM_GI_top_center_L_rb_M_Sling_1, LM_GI_top_center_L_rb_M_Sling_2, LM_GI_top_center_L_rb_R_Mid_0, LM_GI_top_center_L_rb_R_Mid_1, LM_GI_top_center_L_rb_R_Mid_2, LM_GI_top_center_L_rb_R_Top0_0, _
  LM_GI_top_center_L_rb_R_Top0_1, LM_GI_top_center_L_rb_R_Top2_0, LM_GI_top_center_L_rb_R_Top2_2, LM_GI_top_center_R_BmpRing_B, LM_GI_top_center_R_BmpRing_G, LM_GI_top_center_R_BmpRing_Y, LM_GI_top_center_R_BmpSkirt_G, LM_GI_top_center_R_BmpSkirt_Y, LM_GI_top_center_R_Overlay, LM_GI_top_center_R_Parts, LM_GI_top_center_R_Playfield, LM_GI_top_center_R_dt_y1, LM_GI_top_center_R_dt_y2, LM_GI_top_center_R_dt_y3, LM_GI_top_center_R_dt_y4, LM_GI_top_center_R_dt_y5, LM_GI_top_center_R_rb_L_Mid_0, LM_GI_top_center_R_rb_L_Mid_1, LM_GI_top_center_R_rb_L_Mid_2, LM_GI_top_center_R_rb_L_Top0_0, LM_GI_top_center_R_rb_L_Top0_1, LM_GI_top_center_R_rb_L_Top0_2, LM_GI_top_center_R_rb_L_Top2_0, LM_GI_top_center_R_rb_L_Top2_1, LM_GI_top_center_R_rb_L_Top2_2, LM_GI_top_center_R_rb_M_Sling_0, LM_GI_top_center_R_rb_M_Sling_1, LM_GI_top_center_R_rb_M_Sling_2, LM_GI_top_center_R_rb_R_Mid_0, LM_GI_top_center_R_rb_R_Mid_1, LM_GI_top_center_R_rb_R_Mid_2, LM_GI_top_center_R_rb_R_Top0_0, LM_GI_top_center_R_rb_R_Top0_1, _
  LM_GI_top_center_R_rb_R_Top0_2, LM_GI_top_center_R_rb_R_Top1_0, LM_GI_top_center_R_rb_R_Top1_1, LM_GI_top_center_R_rb_R_Top1_2, LM_GI_top_center_R_rb_R_Top2_0, LM_GI_top_center_R_rb_R_Top2_1, LM_GI_top_center_R_rb_R_Top2_2, LM_GI_top_side_L_BmpRing_B, LM_GI_top_side_L_BmpRing_Y, LM_GI_top_side_L_BmpSkirt_Y, LM_GI_top_side_L_Overlay, LM_GI_top_side_L_Parts, LM_GI_top_side_L_Playfield, LM_GI_top_side_L_dt_y1, LM_GI_top_side_L_dt_y2, LM_GI_top_side_L_dt_y3, LM_GI_top_side_L_rb_L_Mid_0, LM_GI_top_side_L_rb_L_Mid_1, LM_GI_top_side_L_rb_L_Mid_2, LM_GI_top_side_L_rb_L_Top0_0, LM_GI_top_side_L_rb_L_Top0_1, LM_GI_top_side_L_rb_L_Top0_2, LM_GI_top_side_L_rb_L_Top1_0, LM_GI_top_side_L_rb_L_Top1_1, LM_GI_top_side_L_rb_L_Top1_2, LM_GI_top_side_L_rb_L_Top2_0, LM_GI_top_side_L_rb_L_Top2_1, LM_GI_top_side_L_rb_L_Top2_2, LM_GI_top_side_L_rb_M_Sling_0, LM_GI_top_side_L_rb_M_Sling_1, LM_GI_top_side_L_rb_M_Sling_2, LM_GI_top_side_R_BmpRing_G, LM_GI_top_side_R_BmpRing_Y, LM_GI_top_side_R_BmpSkirt_G, LM_GI_top_side_R_BmpSkirt_Y, _
  LM_GI_top_side_R_Overlay, LM_GI_top_side_R_Parts, LM_GI_top_side_R_Playfield, LM_GI_top_side_R_dt_g4, LM_GI_top_side_R_dt_y3, LM_GI_top_side_R_dt_y4, LM_GI_top_side_R_dt_y5, LM_GI_top_side_R_rb_M_Sling_0, LM_GI_top_side_R_rb_M_Sling_1, LM_GI_top_side_R_rb_M_Sling_2, LM_GI_top_side_R_rb_R_Mid_0, LM_GI_top_side_R_rb_R_Mid_1, LM_GI_top_side_R_rb_R_Mid_2, LM_GI_top_side_R_rb_R_Top0_2, LM_GI_top_side_R_rb_R_Top1_0, LM_GI_top_side_R_rb_R_Top1_1, LM_GI_top_side_R_rb_R_Top1_2, LM_GI_top_side_R_rb_R_Top2_0, LM_GI_top_side_R_rb_R_Top2_1, LM_GI_top_side_R_rb_R_Top2_2, LM_L_I_bank_B_Playfield, LM_L_I_bank_G_Playfield, LM_L_I_bank_Y_Playfield, LM_L_I_lane_B_Playfield, LM_L_I_lane_G_Playfield, LM_L_I_lane_Y_Playfield, LM_L_I_lane_mid_L_Parts, LM_L_I_lane_mid_L_Playfield, LM_L_I_lane_mid_L_rb_L_Mid_0, LM_L_I_lane_mid_L_rb_L_Mid_1, LM_L_I_lane_mid_L_rb_L_Mid_2, LM_L_I_lane_mid_R_Parts, LM_L_I_lane_mid_R_Playfield, LM_L_I_lane_mid_R_rb_R_Mid_0, LM_L_I_lane_mid_R_rb_R_Mid_1, LM_L_I_lane_mid_R_rb_R_Mid_2, _
  LM_L_I_lane_out_L_Parts, LM_L_I_lane_out_L_Playfield, LM_L_I_lane_out_R_Parts, LM_L_I_lane_out_R_Playfield, LM_L_I_lane_top_C_Playfield, LM_L_I_special_L_Playfield, LM_L_I_special_R_Playfield, LM_L_L_Bmp_B_BmpRing_B, LM_L_L_Bmp_B_BmpSkirt_B, LM_L_L_Bmp_B_Parts, LM_L_L_Bmp_G_BmpRing_G, LM_L_L_Bmp_G_BmpSkirt_G, LM_L_L_Bmp_G_Parts, LM_L_L_Bmp_Y_BmpRing_Y, LM_L_L_Bmp_Y_BmpSkirt_Y, LM_L_L_Bmp_Y_Overlay, LM_L_L_Bmp_Y_Parts, LM_L_L_Bmp_Y_rb_M_Sling_2)
' VLM  Arrays - End

'__/\\\______________/\\\_____/\\\\\\\\\_______/\\\\\\\\\______/\\\__________/\\\\\_______/\\\________/\\\_
' _\/\\\_____________\/\\\___/\\\\\\\\\\\\\___/\\\///////\\\___\/\\\_________/\\\///\\\____\/\\\_____/\\\//__
'  _\/\\\_____________\/\\\__/\\\/////////\\\_\/\\\_____\/\\\___\/\\\________/\\\/__\///\\\__\/\\\__/\\\//_____
'   _\//\\\____/\\\____/\\\__\/\\\_______\/\\\_\/\\\\\\\\\\\/____\/\\\________/\\\______\//\\\_\/\\\\\\//\\\_____
'    __\//\\\__/\\\\\__/\\\___\/\\\\\\\\\\\\\\\_\/\\\//////\\\____\/\\\________\/\\\_______\/\\\_\/\\\//_\//\\\____
'     ___\//\\\/\\\/\\\/\\\____\/\\\/////////\\\_\/\\\____\//\\\___\/\\\_________\//\\\______/\\\__\/\\\____\//\\\___
'      ____\//\\\\\\//\\\\\_____\/\\\_______\/\\\_\/\\\_____\//\\\__\/\\\___________\///\\\__/\\\____\/\\\_____\//\\\__
'       _____\//\\\__\//\\\______\/\\\_______\/\\\_\/\\\______\//\\\_\/\\\\\\\\\\\\\\\__\///\\\\\/_____\/\\\______\//\\\_
'        ______\///____\///_______\///________\///__\///________\///__\///////////////______\/////_______\///________\///__


' Warlok
' Williams 1982
' https://www.ipdb.org/machine.cgi?id=2754

' VPW TEAM
' Original VLM Table Build & Code by bord & Niwak
' VLM, VR & Code Updates by MetaTed
' Backglass by Hauntfreaks
' VLM & Coding Tutelage by Apophis
' Blender Assistance by FrankEnstein
' Code Assistance by DGrimmReaper & somatik
' Tested by DGrimmReaper, FrankEnstein, Studlygoorite, Mecha_Enron, somatik, Apophis, AstroNasty, Lumi & guus_8005

Option Explicit
Randomize


'///////////////////////-----General Sound Options-----///////////////////////
'//  VolumeDial:
'//  VolumeDial is the actual global volume multiplier for the mechanical sounds.
'//  Values smaller than 1 will decrease mechanical sounds volume.
'//  Recommended values should be no greater than 1.

Const VolumeDial = 1

On Error Resume Next
ExecuteGlobal GetTextFile("controller.vbs")
'If Err Then MsgBox "Can't open controller.vbs"
On Error Goto 0

Const cSingleLFlip = 0
Const cSingleRFlip = 0

Const cGameName="wrlok_l3",UseSolenoids=2,UseLamps=1,UseGI=0,SSolenoidOn="Solenoid",SSolenoidOff="SolOff"
Const SCoin=""

LoadVPM "01560000", "S7.VBS", 3.26

' using table width and height in script slows down the performance
dim tablewidth: tablewidth = table1.width
dim tableheight: tableheight = table1.height

Dim BallShadows: Ballshadows=1      '******************set to 1 to turn on Ball shadows
Dim FlipperShadows: FlipperShadows=1  '***********set to 1 to turn on Flipper shadows


'Solenoid Call backs
'**********************************************************************************************************
SolCallback(1)  = "Sol1DropUp"
SolCallback(2)  = "Sol2DropUp"
SolCallback(3)  = "Sol3DropUp"
SolCallback(4)  = "SolRelease"
SolCallback(5)  = "dt1drop"
SolCallback(6)  = "dt2drop"
SolCallback(7)  = "dt3drop"
SolCallback(8)  = "SolGi"
SolCallback(15) = "SolKnocker"
SolCallback(23) = "vpmNudge.SolGameOn"

' The GI is inverted for some reason when the Orbit shot is made (MetaTed hack fix)
Sub SolGi(Enabled)
  'Debug.Print "SolGi " & Enabled
  Dim L
  If Enabled Then
    For each L in GI: L.state = 0: Next
  Else
    For each L in GI: L.state = 1: Next
  End If
End Sub


' ******************************************************************************************
'      LAMP CALLBACK for the 6 backglass flasher lamps (not the solenoid conrolled ones)
' ******************************************************************************************

Set LampCallback = GetRef("UpdateMultipleLamps")

Sub UpdateMultipleLamps()
    If Controller.Lamp(2) = 0 Then: gi27.state = 0: else: gi27.state = 1 'Ball In Play
    If Controller.Lamp(4) = 0 Then: gi25.state = 0: else: gi25.state = 1 'Game Over
    If Controller.Lamp(6) = 0 Then: gi30.state = 0: else: gi30.state = 1 'High Score To Date
    If Controller.Lamp(5) = 0 Then: gi28.state = 0: else: gi28.state = 1 'Match
    If Controller.Lamp(3) = 0 Then: gi29.state = 0: else: gi29.state = 1 'Tilt
    If Controller.Lamp(15) = 0 Then: gi26.state = 0: else: gi26.state = 1 'Shoot Again
End Sub


'**********************************************************************************************************

'Initiate Table
'**********************************************************************************************************

Dim bsTrough, bslsaucer

Sub Table1_Init
  vpmInit Me
' On Error Resume Next
  With Controller
    .GameName = cGameName
    'If Err Then MsgBox "Can't start Game" & cGameName & vbNewLine & Err.Description : Exit Sub
    .SplashInfoLine = "Warlok (Williams 1982)"
    .HandleMechanics=0
    .HandleKeyboard=0
    .ShowDMDOnly=1
    .ShowFrame=0
    .ShowTitle=0
    .hidden = 1
    'If Err Then MsgBox Err.Description
  End With
  On Error Goto 0
    Controller.SolMask(0)=0
    vpmTimer.AddTimer 2000,"Controller.SolMask(0)=&Hffffffff'" 'ignore all solenoids - then add the timer to renable all the solenoids after 2 seconds
    Controller.Run
  'If Err Then MsgBox Err.Description
  On Error Goto 0

  PinMAMETimer.Interval=PinMAMEInterval
  PinMAMETimer.Enabled=1
  vpmNudge.TiltSwitch = 1
  vpmNudge.Sensitivity = 3
  vpmNudge.TiltObj = Array(Bumper1,LeftSlingshot,RightSlingshot)

  vpmMapLights AllLamps ' Map all lamps to the corresponding ROM output using the value of TimerInterval of each light object

  ' Initialize dark/lit room level
  Dim x: x = LoadValue(cGameName, "LIGHT")
'    If x <> "" Then LightLevel = CInt(x) Else LightLevel = 50
' UpdateLightLevel

  ' Hide all lights except from ball reflections
  For Each x in GetElements
    If typename(x) = "Light" Then x.Visible = False
  Next

  ' The GI is inverted for some reason when the Orbit shot is made (MetaTed hack fix)
  SolGi False

  InitSlings 'hides the sling moveables

End Sub

  Drain.Createball
  Controller.Switch(9) = 1

Sub Table1_exit()
  Controller.Pause = False
  Controller.Stop
End Sub



'*******************************************
'  ZOPT: User Options
'*******************************************

Dim LightLevel : LightLevel = 1     ' Level of room lighting (0 to 1), where 0 is dark and 100 is brightest
'Dim VolumeDial : VolumeDial = 0.8            ' Overall Mechanical sound effect volume. Recommended values should be no greater than 1.
'Dim BallRollVolume : BallRollVolume = 0.5    ' Level of ball rolling volume. Value between 0 and 1
'Dim RampRollVolume : RampRollVolume = 0.5    ' Level of ramp rolling volume. Value between 0 and 1


' Called when options are tweaked by the player.
' - 0: game has started, good time to load options and adjust accordingly
' - 1: an option has changed
' - 2: options have been reseted
' - 3: player closed the tweak UI, good time to update staticly prerendered parts
' Table1.Option arguments are:
' - option name, minimum value, maximum value, step between valid values, default value, unit (0=None, 1=Percent), an optional arry of literal strings

'See VR in Desktop mode
  Const VRinDT = False

Dim dspTriggered : dspTriggered = False
Dim VRRoomChoice
Dim BallBrightness

Sub Table1_OptionEvent(ByVal eventId)
    If eventId = 1 And Not dspTriggered Then dspTriggered = True : DisableStaticPreRendering = True : End If

'VR Room

  VRRoomChoice = 1
  VRRoomChoice = Table1.Option("VR Room", 1, 2, 1, 1, 0, Array("Warlok's Temple", "Black Void"))

  SetupRoom

' Sound volumes

  'VolumeDial = Table1.Option("Mech Volume", 0, 1, 0.01, 0.8, 1)
  'BallRollVolume = Table1.Option("Ball Roll Volume", 0, 1, 0.01, 0.5, 1)
  'RampRollVolume = Table1.Option("Ramp Roll Volume", 0, 1, 0.01, 0.5, 1)


  BallBrightness =  Table1.Option("Ball Brightness", 0, 1, 0.05, 0.7, 1) 'Ball brightness - Value between 0 and 1 (0=Dark ... 1=Bright)
  UpdateBallBrightness

  ' Room Brightness
  LightLevel = NightDay/100
  SetRoomBrightness LightLevel   'Uncomment this line for lightmapped tables.

    If eventId = 3 And dspTriggered Then dspTriggered = False : DisableStaticPreRendering = False : End If
End Sub


'******************************************************
'   ZBBR: BALL BRIGHTNESS
'******************************************************

' The frame timer interval is -1, so executes at the display frame rate
FrameTimer.Interval = -1
FrameTimer.Enabled = true
Sub FrameTimer_Timer()
  UpdateBallBrightness      'GI for the ball
End Sub

Dim bootdone: bootdone=False
Sub Bootup_Timer: bootdone=True: Bootup.Enabled=false: End sub

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
'   ZBRI: Room Brightness
'****************************

' Update these arrays if you want to change more materials with room light level
Dim RoomBrightnessMtlArray: RoomBrightnessMtlArray = Array("VLM.Bake.Active", "VLM.Bake.Solid","Plastic with an image") '  "VLM.Bake.Active", "VLM.Bake.Solid",


Sub SetRoomBrightness(lvl)
  If lvl > 1 Then lvl = 1
  If lvl < 0 Then lvl = 0
  lvl = lvl^2

  ' Lighting level
  Dim v: v=(lvl * 240 + 15)/255

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



'***************************************************************
' ZVRR: VR Room
'***************************************************************

Sub SetupRoom

  Dim VRThing, BP, DT

  If RenderingMode = 2 OR VRinDT = True Then   ' VR mode
    If VRRoomChoice = 1 Then    ' Warlok's Temple
      Mosque.visible = 1
      For Each VRThing in VRCab: VRThing.visible = 1: Next
      For Each DT in DTleds: DT.visible = 0: Next
    Else          ' Black Void
      Mosque.visible = 0
      For Each VRThing in VRCab: VRThing.visible = 1: Next
      For Each DT in DTleds: DT.visible = 0: Next
    End If
  Else
    If Table1.ShowDt  = True Then     ' Desktop / No VR / Show Desktop LEDs
      Mosque.visible = 0
      For Each VRThing in VRCab: VRThing.visible = 0: Next
      For Each DT in DTleds: DT.visible = 1: Next
    Else
      Mosque.visible = 0        ' Cab / No VR / Hide Desktop LEDs
      For Each VRThing in VRCab: VRThing.visible = 0: Next
      For Each DT in DTleds: DT.visible = 0: Next
    End If
  End If

End Sub



' VR Plunger code

Sub TimerVRPlunger_Timer
  If VRBM_PlungerRod.TransY < 90 then
    VRBM_PlungerRod.TransY = VRBM_PlungerRod.TransY + 5
    VRBM_PlungerCap.TransY = VRBM_PlungerCap.TransY + 5
  End If
End Sub

Sub TimerVRPlunger1_Timer
  VRBM_PlungerRod.TransY = (5* Plunger.Position) -20
  VRBM_PlungerCap.TransY = (5* Plunger.Position) -20
End Sub



'******************************************************
'         FLIPPERS
'******************************************************

SolCallback(sLRFlipper) = "SolRFlipper"
SolCallback(sLLFlipper) = "SolLFlipper"

Const ReflipAngle = 20

Sub SolLFlipper(Enabled)
  If Enabled Then
    LF.Fire

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
    RF.Fire
'   RightFlipper001.rotatetoend
    If rightflipper.currentangle > rightflipper.endangle - ReflipAngle Then
      RandomSoundReflipUpRight RightFlipper
    Else
      SoundFlipperUpAttackRight RightFlipper
      RandomSoundFlipperUpRight RightFlipper
    End If
  Else
    RightFlipper.RotateToStart
'   rightflipper001.rotatetostart
    If RightFlipper.currentangle > RightFlipper.startAngle + 5 Then
      RandomSoundFlipperDownRight RightFlipper
    End If
    FlipperRightHitParm = FlipperUpSoundLevel
  End If
End Sub

'******************************************************
'         KNOCKER
'******************************************************
Sub SolKnocker(Enabled)
  If enabled Then
    KnockerSolenoid 'Add knocker position object
  End If
End Sub

'******************************************************
'       SOLENOID CONTROLLED TOYS
'******************************************************

Sub dt1drop(enabled)
  If enabled Then
    DTRaise 33
    RandomSoundDropTargetReset BM_DT3_1
  End If
 End Sub

Sub dt2drop(enabled)
  If enabled Then
    DTRaise 34
    RandomSoundDropTargetReset BM_DT3_2
  End If
End Sub

Sub dt3drop(enabled)
  If enabled Then
    DTRaise 35
    RandomSoundDropTargetReset BM_DT3_3
  End If
End Sub

Sub Sol1DropUp(enabled)
  if enabled then
    RandomSoundDropTargetReset BM_DT1_2
    DTRaise 27
    DTRaise 28
    DTRaise 29
  end if
End Sub

Sub Sol2DropUp(enabled)
  if enabled then
    RandomSoundDropTargetReset BM_DT2_2
    DTRaise 30
    DTRaise 31
    DTRaise 32
  end if
End Sub

Sub Sol3DropUp(enabled)
  if enabled then
    RandomSoundDropTargetReset BM_DT3_2
    DTRaise 33
    DTRaise 34
    DTRaise 35
  end if
End Sub






'**********************************************************************************************************
'Plunger code
'**********************************************************************************************************

Sub Table1_KeyDown(ByVal KeyCode)
  If keycode = LeftFlipperKey Then
    FlipperActivate LeftFlipper, LFPress
    VRBM_ButtonL.TransX = 8
  End If
  If keycode = RightFlipperKey Then
    FlipperActivate RightFlipper, RFPress
    Controller.Switch(36) = 1
    VRBM_ButtonR.TransX = -8
  End If
  If keycode = LeftTiltKey Then Nudge 90, 1:SoundNudgeLeft()
  If keycode = RightTiltKey Then Nudge 270, 1:SoundNudgeRight()
  If keycode = CenterTiltKey Then Nudge 0, 2:SoundNudgeCenter()

' If keycode = LeftMagnaSave Then
' End If
' If keycode = RightMagnaSave Then
' End If

  If keycode = keyInsertCoin1 or keycode = keyInsertCoin2 or keycode = keyInsertCoin3 or keycode = keyInsertCoin4 Then
    Select Case Int(rnd*3)
      Case 0: PlaySound ("Coin_In_1"), 0, CoinSoundLevel, 0, 0.25
      Case 1: PlaySound ("Coin_In_2"), 0, CoinSoundLevel, 0, 0.25
      Case 2: PlaySound ("Coin_In_3"), 0, CoinSoundLevel, 0, 0.25

    End Select
  End If

  If keycode = StartGameKey Then
    VRBM_Start.TransY = -5
  End If

  If KeyCode = PlungerKey Then
    Plunger.Pullback:SoundPlungerPull()
    TimerVRPlunger.Enabled = True
    TimerVRPlunger1.Enabled = False
  End If

  If bootdone = True then
    If KeyDownHandler(keycode) Then Exit Sub
  End If

End Sub

Sub Table1_KeyUp(ByVal KeyCode)

  If keycode = LeftFlipperKey Then
    FlipperDeActivate LeftFlipper, LFPress
    VRBM_ButtonL.TransX = 0
  End If
  If keycode = RightFlipperKey Then
    FlipperDeActivate RightFlipper, RFPress
    Controller.Switch(36) = 0
    VRBM_ButtonR.TransX = 0
  End If
  If KeyCode = PlungerKey Then
    Plunger.Fire
    If plungelane=1 Then
      SoundPlungerReleaseBall()     'Plunger release sound when there is a ball in shooter lane
    Else
      SoundPlungerReleaseNoBall()     'Plunger release sound when there is no ball in shooter lane
    End If
    TimerVRPlunger.Enabled = False
    TimerVRPlunger1.Enabled = True
  End If

  If keycode = StartGameKey Then
    VRBM_Start.TransY = 0
  End If

  If bootdone = True then
    If KeyUpHandler(keycode) Then Exit Sub
  End If

End Sub

dim plungelane

Sub ballatplunger_hit
  plungelane=1
  plungewire.transz=-6
End Sub

Sub ballatplunger_unhit
  plungelane=0
  plungewire.transz=0
End Sub

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
  Dim gBOT: gBOT = GetBalls

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
    Dim b
    Dim gBOT: gBOT = GetBalls

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

'**********************************************************************************************************
'**********************************************************************************************************

 'switches
'**********************************************************************************************************

Sub Drain_Hit()
  Controller.Switch(9) = 1
  RandomSoundDrain drain
End Sub

Sub Drain_UnHit()  'Drain
  Controller.Switch(9) = 0
End Sub

Sub SolRelease(enabled)
  If enabled Then
    If Drain.BallCntOver = 0 Then
      PlaySoundAt SoundFX("solon",DOFContactors), Drain
    Else
      'PlaySoundAt SoundFX("gamestart",DOFContactors), Drain  <<< Sound did not exist in sound manager
      RandomSoundBallRelease Drain
    End If
    Drain.kick 66, 20
  End If
End Sub

'Drop Targets
Sub Sw27_Hit:DTHit 27:End Sub
Sub Sw28_Hit:DTHit 28:End Sub
Sub Sw29_Hit:DTHit 29:End Sub
Sub Sw30_Hit:DTHit 30:End Sub
Sub Sw31_Hit:DTHit 31:End Sub
Sub Sw32_Hit:DTHit 32:End Sub
Sub Sw33_Hit:DTHit 33:End Sub
Sub Sw34_Hit:DTHit 34:End Sub
Sub Sw35_Hit:DTHit 35:End Sub

'Stand Up Targets
 Sub Tsw19_Hit: PrimStandupTgtHit 19, Tsw19, BM_HT: vpmTimer.PulseSw 19 : End Sub
 Sub Tsw19_Timer: PrimStandupTgtMove 19, Tsw19, BM_HT: End Sub



'****************************
' ZSLG: Slingshots
'****************************

' Slings & div switches
Dim LStep
Dim RStep

Sub InitSlings 'hides extended slings on startup
  Dim BP
  For Each BP in BP_LSling1 : BP.Visible = 0: Next
  For Each BP in BP_LSling2 : BP.Visible = 0: Next
  For Each BP in BP_RSling1 : BP.Visible = 0: Next
  For Each BP in BP_RSling2 : BP.Visible = 0: Next
End Sub

Sub LeftSlingShot_Slingshot
  'LS.VelocityCorrect(Activeball)
  RandomSoundSlingshotLeft BM_LSlingArm
  vpmTimer.PulseSw 20
  LStep = 0
  LeftSlingShot_Timer
  LeftSlingShot.TimerEnabled = 1
  LeftSlingShot.TimerInterval = 17
End Sub

Sub LeftSlingShot_Timer
  Dim BP
  Dim x1, x2, y: x1 = False:x2 = True:y = 20
    Select Case LStep
        Case 2:x1 = True:x2 = False:y = 10
        Case 3:x1 = False:x2 = False:y = 0:LeftSlingShot.TimerEnabled = 0
    End Select

  For Each BP in BP_LSling1 : BP.Visible = x1: Next
  For Each BP in BP_LSling2 : BP.Visible = x2: Next
  For Each BP in BP_LSlingArm : BP.transx = y: Next

    LStep = LStep + 1
End Sub


Sub RightSlingShot_Slingshot
  'RS.VelocityCorrect(Activeball)
  RandomSoundSlingshotRight BM_RSlingArm
  vpmTimer.PulseSw 21
  RStep = 0
  RightSlingShot_Timer
  RightSlingShot.TimerEnabled = 1
  RightSlingShot.TimerInterval = 17
End Sub


Sub RightSlingShot_Timer
  Dim BP
  Dim x1, x2, y: x1 = False:x2 = True:y = 20
    Select Case RStep
        Case 2:x1 = True:x2 = False:y = 10
        Case 3:x1 = False:x2 = False:y = 0:RightSlingShot.TimerEnabled = 0
    End Select

  For Each BP in BP_RSling1 : BP.Visible = x1: Next
  For Each BP in BP_RSling2 : BP.Visible = x2: Next
  For Each BP in BP_RSlingArm : BP.transx = y: Next

    RStep = RStep + 1
End Sub


'****************************
' ZBMP: Bumpers
'****************************

'**********************************************************
'            SKIRT ANIMATION FUNCTIONS
'**********************************************************
' NOTE: set bumper object timer to around 150-175 in order to be able
' to actually see the animation, adjust to your liking

'Const PI = 3.1415926
Const SkirtTilt = 6       'angle of skirt tilting in degrees

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


Sub Bumper1_Hit()
  vpmTimer.PulseSw 13
  RandomSoundBumperBottom Bumper1

  'VLM skirt movable script
  Dim BP
  For each BP in BP_Bumper1Skirt
    BP.roty = skirtAY(me,Activeball)
    BP.rotx = skirtAX(me,Activeball)
  Next
  me.TimerEnabled = 1
End Sub

Sub Bumper1_timer
  'VLM movable script
  Dim BP
  For each BP in BP_Bumper1Skirt
    BP.roty = 0
    BP.rotx = 0
  Next
  me.TimerEnabled = 0
End Sub

Sub Bumper2_Hit()
  vpmTimer.PulseSw 14
  RandomSoundBumperMiddle Bumper2

  'VLM skirt movable script
  Dim BP
  For each BP in BP_Bumper2Skirt
    BP.roty = skirtAY(me,Activeball)
    BP.rotx = skirtAX(me,Activeball)
  Next
  me.TimerEnabled = 1
End Sub

Sub Bumper2_timer
  'VLM movable script
  Dim BP
  For each BP in BP_Bumper2Skirt
    BP.roty = 0
    BP.rotx = 0
  Next
  me.TimerEnabled = 0
End Sub

Sub Bumper3_Hit()
  vpmTimer.PulseSw 15
  RandomSoundBumperTop Bumper3

  'VLM skirt movable script
  Dim BP
  For each BP in BP_Bumper3Skirt
    BP.roty = skirtAY(me,Activeball)
    BP.rotx = skirtAX(me,Activeball)
  Next
  me.TimerEnabled = 1
End Sub

Sub Bumper3_timer
  'VLM movable script
  Dim BP
  For each BP in BP_Bumper3Skirt
    BP.roty = 0
    BP.rotx = 0
  Next
  me.TimerEnabled = 0
End Sub


 'Wire Triggers
 Sub sw23_Hit  : Controller.Switch(23) = 1 : End Sub
 Sub sw23_UnHit: Controller.Switch(23) = 0 : End Sub
 Sub sw24_Hit  : Controller.Switch(24) = 1 : End Sub
 Sub sw24_UnHit: Controller.Switch(24) = 0 : End Sub
 Sub sw25_Hit  : Controller.Switch(25) = 1 : End Sub
 Sub sw25_UnHit: Controller.Switch(25) = 0 : End Sub
 Sub sw26_Hit  : Controller.Switch(26) = 1 : End Sub
 Sub sw26_UnHit: Controller.Switch(26) = 0 : End Sub
 Sub sw16_Hit  : Controller.Switch(16) = 1 : End Sub
 Sub sw16_UnHit: Controller.Switch(16) = 0 : End Sub
 Sub sw17_Hit  : Controller.Switch(17) = 1 : End Sub
 Sub sw17_UnHit: Controller.Switch(17) = 0 : End Sub
 Sub sw18_Hit  : Controller.Switch(18) = 1 : End Sub
 Sub sw18_UnHit: Controller.Switch(18) = 0 : End Sub

'rollunder gates
 Sub sw22_Hit  : Controller.Switch(22) = 1 : End Sub
 Sub sw22_UnHit: Controller.Switch(22) = 0 : End Sub
 Sub sw37_Hit  : Controller.Switch(37) = 1 : End Sub
 Sub sw37_UnHit: Controller.Switch(37) = 0 : End Sub
 Sub sw38_Hit  : Controller.Switch(38) = 1 : End Sub
 Sub sw38_UnHit: Controller.Switch(38) = 0 : End Sub

'Spinner
Sub sw10_Spin : vpmTimer.PulseSw (10) :SoundSpinner BM_Spin1: End Sub
Sub sw11_Spin : vpmTimer.PulseSw (11) :SoundSpinner BM_Spin2: End Sub
Sub sw12_Spin : vpmTimer.PulseSw (12) :SoundSpinner BM_Spin3: End Sub



'************************************
'          LEDs Display
'     Based on Scapino's LEDs
'************************************

Dim Digits(32)
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

Sub UpdateLeds
    On Error Resume Next
    Dim ChgLED, ii, jj, chg, stat
    ChgLED = Controller.ChangedLEDs(&HFF, &HFFFF)
    If Not IsEmpty(ChgLED)Then
        For ii = 0 To UBound(ChgLED)
            chg = chgLED(ii, 1):stat = chgLED(ii, 2)
            For jj = 0 to 10
                If stat = Patterns(jj)OR stat = Patterns2(jj)then Digits(chgLED(ii, 0)).SetValue jj
            Next
        Next
    End IF
End Sub




' ************** VR Segment Diplays *****************************


Dim DigitsVR(32)
DigitsVR(14) = Array(p1_1a,p1_1b,p1_1c,p1_1d,p1_1e,p1_1f,p1_1g)
DigitsVR(15) = Array(p1_2a,p1_2b,p1_2c,p1_2d,p1_2e,p1_2f,p1_2g)
DigitsVR(16) = Array(p1_3a,p1_3b,p1_3c,p1_3d,p1_3e,p1_3f,p1_3g)
DigitsVR(17) = Array(p1_4a,p1_4b,p1_4c,p1_4d,p1_4e,p1_4f,p1_4g)
DigitsVR(18) = Array(p1_5a,p1_5b,p1_5c,p1_5d,p1_5e,p1_5f,p1_5g)
DigitsVR(19) = Array(p1_6a,p1_6b,p1_6c,p1_6d,p1_6e,p1_6f,p1_6g)
DigitsVR(20) = Array(p1_7a,p1_7b,p1_7c,p1_7d,p1_7e,p1_7f,p1_7g)

DigitsVR(21) = Array(p2_1a,p2_1b,p2_1c,p2_1d,p2_1e,p2_1f,p2_1g)
DigitsVR(22) = Array(p2_2a,p2_2b,p2_2c,p2_2d,p2_2e,p2_2f,p2_2g)
DigitsVR(23) = Array(p2_3a,p2_3b,p2_3c,p2_3d,p2_3e,p2_3f,p2_3g)
DigitsVR(24) = Array(p2_4a,p2_4b,p2_4c,p2_4d,p2_4e,p2_4f,p2_4g)
DigitsVR(25) = Array(p2_5a,p2_5b,p2_5c,p2_5d,p2_5e,p2_5f,p2_5g)
DigitsVR(26) = Array(p2_6a,p2_6b,p2_6c,p2_6d,p2_6e,p2_6f,p2_6g)
DigitsVR(27) = Array(p2_7a,p2_7b,p2_7c,p2_7d,p2_7e,p2_7f,p2_7g)

DigitsVR(0) = Array(p3_1a,p3_1b,p3_1c,p3_1d,p3_1e,p3_1f,p3_1g)
DigitsVR(1) = Array(p3_2a,p3_2b,p3_2c,p3_2d,p3_2e,p3_2f,p3_2g)
DigitsVR(2) = Array(p3_3a,p3_3b,p3_3c,p3_3d,p3_3e,p3_3f,p3_3g)
DigitsVR(3) = Array(p3_4a,p3_4b,p3_4c,p3_4d,p3_4e,p3_4f,p3_4g)
DigitsVR(4) = Array(p3_5a,p3_5b,p3_5c,p3_5d,p3_5e,p3_5f,p3_5g)
DigitsVR(5) = Array(p3_6a,p3_6b,p3_6c,p3_6d,p3_6e,p3_6f,p3_6g)
DigitsVR(6) = Array(p3_7a,p3_7b,p3_7c,p3_7d,p3_7e,p3_7f,p3_7g)

DigitsVR(7) = Array(p4_1a,p4_1b,p4_1c,p4_1d,p4_1e,p4_1f,p4_1g)
DigitsVR(8) = Array(p4_2a,p4_2b,p4_2c,p4_2d,p4_2e,p4_2f,p4_2g)
DigitsVR(9) = Array(p4_3a,p4_3b,p4_3c,p4_3d,p4_3e,p4_3f,p4_3g)
DigitsVR(10) = Array(p4_4a,p4_4b,p4_4c,p4_4d,p4_4e,p4_4f,p4_4g)
DigitsVR(11) = Array(p4_5a,p4_5b,p4_5c,p4_5d,p4_5e,p4_5f,p4_5g)
DigitsVR(12) = Array(p4_6a,p4_6b,p4_6c,p4_6d,p4_6e,p4_6f,p4_6g)
DigitsVR(13) = Array(p4_7a,p4_7b,p4_7c,p4_7d,p4_7e,p4_7f,p4_7g)

DigitsVR(30) = Array(cr_1a,cr_1b,cr_1c,cr_1d,cr_1e,cr_1f,cr_1g)
DigitsVR(31) = Array(cr_2a,cr_2b,cr_2c,cr_2d,cr_2e,cr_2f,cr_2g)

DigitsVR(28) = Array(mb_1a,mb_1b,mb_1c,mb_1d,mb_1e,mb_1f,mb_1g)
DigitsVR(29) = Array(mb_2a,mb_2b,mb_2c,mb_2d,mb_2e,mb_2f,mb_2g)


Sub SegDispTimer_Timer() 'The game timer interval; should be 10 ms

  If RenderingMode = 2 OR VRinDT = True Then   ' VR mode

    SegDisp ' Update Segment Displays on VR Backglass

  Else

    UpdateLeds   'Update desktop score displays

  End If

End Sub


Sub SegDisp
    Dim ChgLED, ii, jj, num, chg, stat, obj, b, x
    ChgLED=Controller.ChangedLEDs(&Hffffffff, &Hffffffff)
    If Not IsEmpty(ChgLED)Then
       For ii=0 To UBound(chgLED)
          num=chgLED(ii, 0) : chg=chgLED(ii, 1) : stat=chgLED(ii, 2)
      if (num < 32) then
              For Each obj In DigitsVR(num)
'                   If chg And 1 Then obj.visible=stat And 1
           If chg And 1 Then FadeDisplay obj, stat And 1
                   chg=chg\2 : stat=stat\2
              Next
      Else
           end if
        Next
    End If
 End Sub

Sub FadeDisplay(object, onoff)
  If OnOff = 1 Then
    object.state = 1
  Else
    object.state = 0
  End If
End Sub

Sub InitDigitsVR()
  dim tmp, x, obj
  for x = 0 to uBound(DigitsVR)
    if IsArray(DigitsVR(x)) then
      For each obj in DigitsVR(x)
        obj.state = 0
      next
    end If
  Next
End Sub


Const WallPrefix    = "T" 'Change this based on your naming convention
Const PrimitivePrefix   = "PrimT"'Change this based on your naming convention
Const PrimitiveBumperPrefix = "BumperRing" 'Change this based on your naming convention
Dim primCnt(120), primDir(120), primBmprDir(120)

'****************************************************************************
'***** Primitive Standup Target Animation
'****************************************************************************
'USAGE:   Sub sw1_Hit:  PrimStandupTgtHit  1, Sw1, PrimSw1: End Sub
'USAGE:   Sub Sw1_Timer:  PrimStandupTgtMove 1, Sw1, PrimSw1: End Sub

Const StandupTgtMovementDir = "TransY"
Const StandupTgtMovementMax = 6

Sub PrimStandupTgtHit (swnum, wallName, primName)
  PlayTargetSound
  vpmTimer.PulseSw swnum
  primCnt(swnum) = 0                  'Reset count
  wallName.TimerInterval = 10 'Set timer interval
  wallName.TimerEnabled = 1   'Enable timer
End Sub

Sub PrimStandupTgtMove (swnum, wallName, primName)
  Select Case StandupTgtMovementDir
    Case "TransX":
      Select Case primCnt(swnum)
        Case 0:   primName.TransX = -StandupTgtMovementMax
        Case 1:   primName.TransX = -StandupTgtMovementMax *2/3
        Case 2:   primName.TransX = -StandupTgtMovementMax /3
        Case 3:   primName.TransX = 0
        Case else:  wallName.TimerEnabled = 0
      End Select
    Case "TransY":
      Select Case primCnt(swnum)
        Case 0:   primName.TransY = -StandupTgtMovementMax
        Case 1:   primName.TransY = -StandupTgtMovementMax * 2/3
        Case 2:   primName.TransY = -StandupTgtMovementMax /3
        Case 3:   primName.TransY = 0
        Case else:  wallName.TimerEnabled = 0
      End Select
    Case "TransZ":
      Select Case primCnt(swnum)
        Case 0:   primName.TransZ = -StandupTgtMovementMax
        Case 1:   primName.TransZ = -StandupTgtMovementMax * 2/3
        Case 2:   primName.TransZ = -StandupTgtMovementMax /3
        Case 3:   primName.TransZ = 0
        Case else:  wallName.TimerEnabled = 0
      End Select
  End Select
  primCnt(swnum) = primCnt(swnum) + 1
  UpdateHitTarget
End Sub

Sub UpdateHitTarget
  Dim y : y = BM_HT.TransY
  LM_GI_HT.transy = y
  LM_L_L19_HT.transy = y
  LM_L_L39_HT.transy = y
End Sub




if ballshadows=1 then
  BallShadowUpdate.enabled=1
else
  BallShadowUpdate.enabled=0
end if

'*****************************************
'     BALL SHADOW
'*****************************************
Dim BallShadow
BallShadow = Array (BallShadow1,BallShadow2,BallShadow3,BallShadow4,BallShadow5,BallShadow6,BallShadow7,BallShadow8,BallShadow9,BallShadow10)

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
    BallShadow(b).X = BOT(b).X - (TableWidth/2 - BOT(b).X)/20
    ballShadow(b).Y = BOT(b).Y + 10
    If BOT(b).Z > 20 Then
      BallShadow(b).visible = 1
    Else
      BallShadow(b).visible = 0
    End If
  Next
End Sub

'*****************************************
'      JP's VP10 Rolling Sounds
'*****************************************

Const tnob = 10 ' total number of balls
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

Sub RollingTimer_Timer()
  Dim BOT, b
  BOT = GetBalls

  ' stop the sound of deleted balls
  For b = UBound(BOT) + 1 to tnob
    rolling(b) = False
    StopSound("BallRoll_" & b)
  Next

  ' exit the sub if no balls on the table
  If UBound(BOT) = -1 Then Exit Sub

  ' play the rolling sound for each ball

  For b = 0 to UBound(BOT)
    If BallVel(BOT(b)) > 1 AND BOT(b).z < 30 Then
      rolling(b) = True
      PlaySound ("BallRoll_" & b), -1, VolPlayfieldRoll(BOT(b)) * 1.1 * VolumeDial, AudioPan(BOT(b)), 0, PitchPlayfieldRoll(BOT(b)), 1, 0, AudioFade(BOT(b))

    Else
      If rolling(b) = True Then
        StopSound("BallRoll_" & b)
        rolling(b) = False
      End If
    End If

    '***Ball Drop Sounds***
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
  Next
End Sub

'******************************************************
'   FLIPPER CORRECTION INITIALIZATION
'******************************************************

dim LF : Set LF = New FlipperPolarity
dim RF : Set RF = New FlipperPolarity

InitPolarity

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
'   ZDMP:  RUBBER  DAMPENERS
'******************************************************
' These are data mined bounce curves,
' dialed in with the in-game elasticity as much as possible to prevent angle / spin issues.
' Requires tracking ballspeed to calculate COR

Sub dPosts_Hit(idx)
  RubbersD.dampen ActiveBall
End Sub

Sub dSleeves_Hit(idx)
  SleevesD.Dampen ActiveBall
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

Sub RDampen_Timer()
  Cor.Update
End Sub

'******************************************************
'   DROP TARGETS INITIALIZATION
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
Dim DT27, DT28, DT29, DT30, DT31, DT32, DT33, DT34, DT35

'Set array with drop target objects
'
'DropTargetvar = Array(primary, secondary, prim, switch, animate)
'   primary:      primary target wall to determine drop
' secondary:      wall used to simulate the ball striking a bent or offset target after the initial Hit
' prim:       primitive target used for visuals and animation
'             IMPORTANT!!!
'             rotz must be used for orientation
'             rotx to bend the target back
'             transz to move it up and down
'             the pivot point should be in the center of the target on the x, y and at or below the playfield (0) on z
' switch:       ROM switch number
' animate:      Arrary slot for handling the animation instrucitons, set to 0
'
' Values for annimate: 1 - bend target (hit to primary), 2 - drop target (hit to secondary), 3 - brick target (high velocity hit to secondary), -1 - raise target

' 1 Bank

Set DT27 = (new DropTarget)(sw27, sw27y, BM_DT1_1, 27, 0, false)
Set DT28 = (new DropTarget)(sw28, sw28y, BM_DT1_2, 28, 0, false)
Set DT29 = (new DropTarget)(sw29, sw29y, BM_DT1_3, 29, 0, false)

 ' 2 Bank

Set DT30 = (new DropTarget)(sw30, sw30y, BM_DT2_1, 30, 0, false)
Set DT31 = (new DropTarget)(sw31, sw31y, BM_DT2_2, 31, 0, false)
Set DT32 = (new DropTarget)(sw32, sw32y, BM_DT2_3, 32, 0, false)

 ' 3 Bank

Set DT33 = (new DropTarget)(sw33, sw33y, BM_DT3_1, 33, 0, false)
Set DT34 = (new DropTarget)(sw34, sw34y, BM_DT3_2, 34, 0, false)
Set DT35 = (new DropTarget)(sw35, sw35y, BM_DT3_3, 35, 0, false)


'Add all the Drop Target Arrays to Drop Target Animation Array
' DTAnimationArray = Array(DT1, DT2, ....)
Dim DTArray
DTArray = Array(DT27, DT28, DT29, DT30, DT31, DT32, DT33, DT34, DT35)

'Configure the behavior of Drop Targets.
Const DTDropSpeed = 110       'in milliseconds
Const DTDropUpSpeed = 40      'in milliseconds
Const DTDropUnits = 50        'VP units primitive drops
Const DTDropUpUnits = 5       'VP units primitive raises above the up position on drops up
Const DTMaxBend = 8         'max degrees primitive rotates when hit
Const DTDropDelay = 20      'time in milliseconds before target drops (due to friction/impact of the ball)
Const DTRaiseDelay = 40     'time in milliseconds before target drops back to normal up position after the solendoid fires to raise the target
Const DTBrickVel = 30       'velocity at which the target will brick, set to '0' to disable brick

Const DTEnableBrick = 0     'Set to 0 to disable bricking, 1 to enable bricking
Const DTHitSound = "" 'Drop Target Hit sound
Const DTDropSound = ""    'Drop Target Drop sound
Const DTResetSound = "" 'Drop Target reset sound

Const DTMass = 0.2        'Mass of the Drop Target (between 0 and 1), higher values provide more resistance

'******************************************************
'       DROP TARGETS FUNCTIONS
'******************************************************

Sub DTHit(switch)
  Dim i
  i = DTArrayID(switch)
  PlayTargetSound
  DTArray(i).animate =  DTCheckBrick(Activeball,DTArray(i).prim)
  DTArray(i).animate =  DTCheckBrick(Activeball,DTArray(i).prim)
  If DTArray(i).animate = 1 or DTArray(i).animate = 3 Then
    DTBallPhysics Activeball, DTArray(i).prim.rotz, DTMass
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

'Add Timer name DTAnim to editor to handle drop target animations
DTAnim.interval = -1
DTAnim.enabled = True

Sub DTAnim_Timer()
  DoDTAnim
End Sub

'Check if target is hit on it's face or sides and whether a 'brick' occurred
Function DTCheckBrick(aBall, dtprim)
  dim bangle, bangleafter, rangle, rangle2, Xintersect, Yintersect, cdist, perpvel, perpvelafter
  rangle = (dtprim.rotz - 90) * 3.1416 / 180
  rangle2 = dtprim.rotz * 3.1416 / 180
  bangle = atn2(cor.ballvely(aball.id),cor.ballvelx(aball.id))
  bangleafter = Atn2(aBall.vely,aball.velx)

  Xintersect = (aBall.y - dtprim.y - tan(bangle) * aball.x + tan(rangle2) * dtprim.x) / (tan(rangle2) - tan(bangle))
  Yintersect = tan(rangle2) * Xintersect + (dtprim.y - tan(rangle2) * dtprim.x)

  cdist = Distance(dtprim.x, dtprim.y, Xintersect, Yintersect)

  perpvel = cor.BallVel(aball.id) * cos(bangle-rangle)
  perpvelafter = BallSpeed(aBall) * cos(bangleafter - rangle)

  If perpvel <= 0 or perpvelafter >= 0 Then
    DTCheckBrick = 0
  ElseIf DTEnableBrick = 1 and  perpvel > DTBrickVel and DTBrickVel <> 0 and cdist < 8 Then
    DTCheckBrick = 3
  Else
    DTCheckBrick = 1
  End If
End Function


Sub DoDTAnim()
  Dim i
  For i=0 to Ubound(DTArray)
    DTArray(i).animate = DTAnimate(DTArray(i).primary,DTArray(i).secondary,DTArray(i).prim,DTArray(i).sw,DTArray(i).animate)
  Next
  'table specific code
' DTShadow psw25, drop4shadow, "drop4"
' DTShadow psw26, drop3shadow, "drop3"
' DTShadow psw27, drop2shadow, "drop2"
' DTShadow psw28, drop1shadow, "drop1"
  UpdateTargets
End Sub


Sub UpdateTargets

  Dim tz, rx, ry
  Dim BP

  tz = BM_DT1_1.transz
  rx = BM_DT1_1.rotx
  ry = BM_DT1_1.roty
  For Each BP in BP_DT1_1: BP.transz = tz: BP.rotx = rx: BP.roty = ry: Next

  tz = BM_DT1_2.transz
  rx = BM_DT1_2.rotx
  ry = BM_DT1_2.roty
  For Each BP in BP_DT1_2: BP.transz = tz: BP.rotx = rx: BP.roty = ry: Next

  tz = BM_DT1_3.transz
  rx = BM_DT1_3.rotx
  ry = BM_DT1_3.roty
  For Each BP in BP_DT1_3: BP.transz = tz: BP.rotx = rx: BP.roty = ry: Next

  tz = BM_DT2_1.transz
  rx = BM_DT2_1.rotx
  ry = BM_DT2_1.roty
  For Each BP in BP_DT2_1: BP.transz = tz: BP.rotx = rx: BP.roty = ry: Next

  tz = BM_DT2_2.transz
  rx = BM_DT2_2.rotx
  ry = BM_DT2_2.roty
  For Each BP in BP_DT2_2: BP.transz = tz: BP.rotx = rx: BP.roty = ry: Next

  tz = BM_DT2_3.transz
  rx = BM_DT2_3.rotx
  ry = BM_DT2_3.roty
  For Each BP in BP_DT2_3: BP.transz = tz: BP.rotx = rx: BP.roty = ry: Next

  tz = BM_DT3_1.transz
  rx = BM_DT3_1.rotx
  ry = BM_DT3_1.roty
  For Each BP in BP_DT3_1: BP.transz = tz: BP.rotx = rx: BP.roty = ry: Next

  tz = BM_DT3_2.transz
  rx = BM_DT3_2.rotx
  ry = BM_DT3_2.roty
  For Each BP in BP_DT3_2: BP.transz = tz: BP.rotx = rx: BP.roty = ry: Next

  tz = BM_DT3_3.transz
  rx = BM_DT3_3.rotx
  ry = BM_DT3_3.roty
  For Each BP in BP_DT3_3: BP.transz = tz: BP.rotx = rx: BP.roty = ry: Next

End Sub


Function DTAnimate(primary, secondary, prim, switch,  animate)
  dim transz
  Dim animtime, rangle

  rangle = prim.rotz * 3.1416 / 180

  DTAnimate = animate

  if animate = 0  Then
    primary.uservalue = 0
    DTAnimate = 0
    Exit Function
  Elseif primary.uservalue = 0 then
    primary.uservalue = gametime
  end if

  animtime = gametime - primary.uservalue

  If animate = 1 and animtime < DTDropDelay Then
    primary.collidable = 0
    secondary.collidable = 1
    prim.rotx = DTMaxBend * cos(rangle)
    prim.roty = DTMaxBend * sin(rangle)
    DTAnimate = 1
    Exit Function
  elseif animate = 1 and animtime > DTDropDelay Then
    primary.collidable = 0
    secondary.collidable = 1
    prim.rotx = DTMaxBend * cos(rangle)
    prim.roty = DTMaxBend * sin(rangle)
    animate = 2
    SoundDropTargetDrop prim
    'PlaySoundAt SoundFX(DTDropSound,DOFDropTargets),prim
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
      controller.Switch(Switch) = 1
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
      Dim BOT, b
      BOT = GetBalls

      For b = 0 to UBound(BOT)
        If InRect(BOT(b).x,BOT(b).y,prim.x-25,prim.y-10,prim.x+25, prim.y-10,prim.x+25,prim.y+25,prim.x -25,prim.y+25) Then
          BOT(b).velz = 20
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
      prim.rotx = 0
      prim.roty = 0
      primary.uservalue = gametime
    end if
    primary.collidable = 0
    secondary.collidable = 1
    controller.Switch(Switch) = 0

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

'' Table specific function
'Sub  DTShadow(prim, shadowprim, imagename)
' If prim.transz < -DTDropUnits/2 Then
'   shadowprim.image ="blank"
' Else
'   shadowprim.image = imagename
' End If
'End Sub

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


'////////////////////////////  MECHANICAL SOUNDS  ///////////////////////////
'//  This part in the script is an entire block that is dedicated to the physics sound system.
'//  Various scripts and sounds that may be pretty generic and could suit other WPC systems, but the most are tailored specifically for this table.

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
SpinnerSoundLevel = 0.1      'volume level; range [0, 1]

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

' Thalamus, AudioFade patched
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

' Thalamus, AudioPan patched
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
  'If err.number = 13 then  msgbox "Proc error! No such procedure: " & vbnewline & string
  'if err.number = 424 then msgbox "Proc error! No such Object"
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



'******************************************************
'   ZANI: Misc Animations
'******************************************************


''''' Flippers


if flippershadows=1 then
  batleftshadow.visible=1
  batrightshadow.visible=1
else
  batleftshadow.visible=0
  batrightshadow.visible=0
end if

Sub LeftFlipper_Animate
  Dim a: a = LeftFlipper.CurrentAngle       ' store flipper angle in a
  Dim max_angle, min_angle, mid_angle       ' min and max angles from the flipper
  max_angle = LeftFlipper.StartAngle                  ' flipper down angle
  min_angle = LeftFlipper.EndAngle
  mid_angle = (max_angle-min_angle)/2 + min_angle ' bake map switch point angle
  batleftshadow.RotZ = a                ' set flipper shadow angle

  Dim BP :
  For Each BP in BP_FlipperL
    BP.RotZ = a                 ' rotate the maps
    BP.visible = a > mid_angle
  Next
  For Each BP in BP_FlipperLup
    BP.RotZ = a                 ' rotate the maps
    BP.visible = a < mid_angle
  Next
End Sub

Sub RightFlipper_Animate
  Dim a: a = RightFlipper.CurrentAngle        ' store flipper angle in a
  Dim max_angle, min_angle, mid_angle       ' min and max angles from the flipper
  max_angle = RightFlipper.StartAngle               ' flipper down angle
  min_angle = RightFlipper.EndAngle                 ' flipper up angle
  mid_angle = (max_angle-min_angle)/2 + min_angle ' bake map switch point angle
  batrightshadow.RotZ = a               ' set flipper shadow angle

  Dim BP :
  For Each BP in BP_FlipperR
    BP.RotZ = a                 ' rotate the maps
    BP.visible = a < mid_angle
  Next
  For Each BP in BP_FlipperRup
    BP.visible = a > mid_angle
    BP.RotZ = a                 ' rotate the maps
  Next
End Sub



'''Rollovers

Sub sw16_Animate
  Dim z : z = sw16.CurrentAnimOffset
  Dim BP : For Each BP in BP_wire5 : BP.transz = z: Next
End Sub

Sub sw17_Animate
  Dim z : z = sw17.CurrentAnimOffset
  Dim BP : For Each BP in BP_wire6 : BP.transz = z: Next
End Sub

Sub sw18_Animate
  Dim z : z = sw18.CurrentAnimOffset
  Dim BP : For Each BP in BP_wire7 : BP.transz = z: Next
End Sub

Sub sw23_Animate
  Dim z : z = sw23.CurrentAnimOffset
  Dim BP : For Each BP in BP_wire1 : BP.transz = z: Next
End Sub

Sub sw24_Animate
  Dim z : z = sw24.CurrentAnimOffset
  Dim BP : For Each BP in BP_wire2 : BP.transz = z: Next
End Sub

Sub sw25_Animate
  Dim z : z = sw25.CurrentAnimOffset
  Dim BP : For Each BP in BP_wire3 : BP.transz = z: Next
End Sub

Sub sw26_Animate
  Dim z : z = sw26.CurrentAnimOffset
  Dim BP : For Each BP in BP_wire4 : BP.transz = z: Next
End Sub


''' Spinner Animations

' Set all Spin(x)Rod Zpos to 55 after new bake

Sub sw10_Animate
  Dim spinangle:spinangle = sw10.currentangle
  Dim BL : For Each BL in BP_Spin1 : BL.RotX = spinangle: Next
  Dim BR
  For Each BR in BP_Spin1Rod
    BR.TransX = sin( (sw10.CurrentAngle) * (2*PI/360)) * 3.5
    BR.TransZ = sin( (sw10.CurrentAngle- 90) * (2*PI/360)) * 10
  Next
End Sub

Sub sw11_Animate
  Dim spinangle:spinangle = sw11.currentangle
  Dim BL : For Each BL in BP_Spin2 : BL.RotX = spinangle: Next
  Dim BR
  For Each BR in BP_Spin2Rod
    BR.TransX = sin((sw11.CurrentAngle) * (2*PI/360)) * 3.5
    BR.TransZ = sin((sw11.CurrentAngle-90) * (2*PI/360)) * 10
  Next
End Sub

Sub sw12_Animate
  Dim spinangle:spinangle = sw12.currentangle
  Dim BL : For Each BL in BP_Spin3 : BL.RotX = spinangle: Next
  Dim BR
  For Each BR in BP_Spin3Rod
    BR.TransX = sin((sw12.CurrentAngle) * (2*PI/360)) * 3.5
    BR.TransZ = sin((sw12.CurrentAngle-90) * (2*PI/360)) * 10
  Next
End Sub


''' Gate Animations

Sub Gate001_Animate
  Dim a : a = Gate001.CurrentAngle
  Dim BL : For Each BL in BP_Gate_1 : BL.rotx = a: Next
End Sub

Sub Gate002_Animate
  Dim a : a = Gate002.CurrentAngle
  Dim BL : For Each BL in BP_Gate_6 : BL.rotx = a: Next
End Sub

Sub Gate003_Animate
  Dim a : a = Gate003.CurrentAngle
  Dim BL : For Each BL in BP_Gate_4 : BL.rotx = a: Next
End Sub

Sub Gate004_Animate
  Dim a : a = Gate004.CurrentAngle
  Dim BL : For Each BL in BP_Gate_2 : BL.rotx = a: Next
End Sub

Sub Gate005_Animate
  Dim a : a = Gate005.CurrentAngle
  Dim BL : For Each BL in BP_Gate_3 : BL.rotx = a: Next
End Sub

Sub Gate006_Animate
  Dim a : a = Gate006.CurrentAngle
  Dim BL : For Each BL in BP_Gate_5 : BL.rotx = a: Next
End Sub


''' Bumper Animations
Sub Bumper1_Animate
  Dim z, BL
  z = Bumper1.CurrentRingOffset
  For Each BL in BP_BR1 : BL.transz = z: Next
End Sub

Sub Bumper2_Animate
  Dim z, BL
  z = Bumper2.CurrentRingOffset
  For Each BL in BP_BR2 : BL.transz = z: Next
End Sub

Sub Bumper3_Animate
  Dim z, BL
  z = Bumper3.CurrentRingOffset
  For Each BL in BP_BR3 : BL.transz = z: Next
End Sub




'****************************************************
'*                          *
'*        VLM ARRAYS              *
'*                          *
'****************************************************

' VLM VR Arrays - Start
' Arrays per baked part
Dim BP_VRButtonL: BP_VRButtonL=Array(VRBM_ButtonL)
Dim BP_VRButtonR: BP_VRButtonR=Array(VRBM_ButtonR)
Dim BP_VRButtonRingL: BP_VRButtonRingL=Array(VRBM_ButtonRingL)
Dim BP_VRButtonRingR: BP_VRButtonRingR=Array(VRBM_ButtonRingR)
Dim BP_VRCab_Body: BP_VRCab_Body=Array(VRBM_Cab_Body)
Dim BP_VRCab_Head: BP_VRCab_Head=Array(VRBM_Cab_Head)
Dim BP_VRLeg_BL: BP_VRLeg_BL=Array(VRBM_Leg_BL)
Dim BP_VRLeg_BR: BP_VRLeg_BR=Array(VRBM_Leg_BR)
Dim BP_VRLeg_FL: BP_VRLeg_FL=Array(VRBM_Leg_FL)
Dim BP_VRLeg_FR: BP_VRLeg_FR=Array(VRBM_Leg_FR)
Dim BP_VRPlungerCap: BP_VRPlungerCap=Array(VRBM_PlungerCap)
Dim BP_VRPlungerHousing: BP_VRPlungerHousing=Array(VRBM_PlungerHousing)
Dim BP_VRPlungerRod: BP_VRPlungerRod=Array(VRBM_PlungerRod)
Dim BP_VRStart: BP_VRStart=Array(VRBM_Start)
' Arrays per lighting scenario
Dim BL_VRRoom: BL_VRRoom=Array(VRBM_ButtonL, VRBM_ButtonR, VRBM_ButtonRingL, VRBM_ButtonRingR, VRBM_Cab_Body, VRBM_Cab_Head, VRBM_Leg_BL, VRBM_Leg_BR, VRBM_Leg_FL, VRBM_Leg_FR, VRBM_PlungerCap, VRBM_PlungerHousing, VRBM_PlungerRod, VRBM_Start)
' Global arrays
Dim BGVR_Bakemap: BGVR_Bakemap=Array(VRBM_ButtonL, VRBM_ButtonR, VRBM_ButtonRingL, VRBM_ButtonRingR, VRBM_Cab_Body, VRBM_Cab_Head, VRBM_Leg_BL, VRBM_Leg_BR, VRBM_Leg_FL, VRBM_Leg_FR, VRBM_PlungerCap, VRBM_PlungerHousing, VRBM_PlungerRod, VRBM_Start)
Dim BGVR_Lightmap: BGVR_Lightmap=Array()
Dim BGVR_All: BGVR_All=Array(VRBM_ButtonL, VRBM_ButtonR, VRBM_ButtonRingL, VRBM_ButtonRingR, VRBM_Cab_Body, VRBM_Cab_Head, VRBM_Leg_BL, VRBM_Leg_BR, VRBM_Leg_FL, VRBM_Leg_FR, VRBM_PlungerCap, VRBM_PlungerHousing, VRBM_PlungerRod, VRBM_Start)
' VLM VR Arrays - End

' VLM  Arrays - Start
' Arrays per baked part
Dim BP_BR1: BP_BR1=Array(BM_BR1, LM_L_L19_BR1, LM_L_L39_BR1, LM_GI_BR1)
Dim BP_BR2: BP_BR2=Array(BM_BR2, LM_L_L25_BR2, LM_GI_BR2)
Dim BP_BR3: BP_BR3=Array(BM_BR3, LM_GI_BR3)
Dim BP_Bumper1Skirt: BP_Bumper1Skirt=Array(BM_Bumper1Skirt, LM_L_L19_Bumper1Skirt, LM_L_L39_Bumper1Skirt, LM_GI_Bumper1Skirt)
Dim BP_Bumper2Skirt: BP_Bumper2Skirt=Array(BM_Bumper2Skirt, LM_L_L25_Bumper2Skirt, LM_GI_Bumper2Skirt)
Dim BP_Bumper3Skirt: BP_Bumper3Skirt=Array(BM_Bumper3Skirt, LM_GI_Bumper3Skirt)
Dim BP_DT1_1: BP_DT1_1=Array(BM_DT1_1, LM_L_L20_DT1_1, LM_GI_DT1_1)
Dim BP_DT1_2: BP_DT1_2=Array(BM_DT1_2, LM_L_L20_DT1_2, LM_L_L23_DT1_2, LM_GI_DT1_2)
Dim BP_DT1_3: BP_DT1_3=Array(BM_DT1_3, LM_L_L20_DT1_3, LM_GI_DT1_3)
Dim BP_DT2_1: BP_DT2_1=Array(BM_DT2_1, LM_L_L21_DT2_1, LM_L_L24_DT2_1, LM_GI_DT2_1)
Dim BP_DT2_2: BP_DT2_2=Array(BM_DT2_2, LM_L_L21_DT2_2, LM_L_L24_DT2_2, LM_GI_DT2_2)
Dim BP_DT2_3: BP_DT2_3=Array(BM_DT2_3, LM_L_L21_DT2_3, LM_L_L24_DT2_3, LM_GI_DT2_3)
Dim BP_DT3_1: BP_DT3_1=Array(BM_DT3_1, LM_L_L22_DT3_1, LM_L_L25_DT3_1, LM_GI_DT3_1)
Dim BP_DT3_2: BP_DT3_2=Array(BM_DT3_2, LM_L_L22_DT3_2, LM_L_L25_DT3_2, LM_GI_DT3_2)
Dim BP_DT3_3: BP_DT3_3=Array(BM_DT3_3, LM_L_L22_DT3_3, LM_L_L25_DT3_3, LM_GI_DT3_3)
Dim BP_FlipperL: BP_FlipperL=Array(BM_FlipperL, LM_L_L31_FlipperL, LM_GIS_gi01_FlipperL, LM_GIS_gi02_FlipperL, LM_GIS_gi03_FlipperL)
Dim BP_FlipperLup: BP_FlipperLup=Array(BM_FlipperLup, LM_L_L31_FlipperLup, LM_L_L34_FlipperLup, LM_GIS_gi01_FlipperLup, LM_GIS_gi02_FlipperLup, LM_GIS_gi03_FlipperLup)
Dim BP_FlipperR: BP_FlipperR=Array(BM_FlipperR, LM_L_L15_FlipperR, LM_L_L34_FlipperR, LM_L_L59_FlipperR, LM_GIS_gi01_FlipperR, LM_GIS_gi03_FlipperR, LM_GIS_gi04_FlipperR)
Dim BP_FlipperRup: BP_FlipperRup=Array(BM_FlipperRup, LM_L_L15_FlipperRup, LM_L_L34_FlipperRup, LM_GIS_gi03_FlipperRup, LM_GIS_gi04_FlipperRup)
Dim BP_Gate_1: BP_Gate_1=Array(BM_Gate_1, LM_L_L12_Gate_1, LM_GI_Gate_1)
Dim BP_Gate_2: BP_Gate_2=Array(BM_Gate_2)
Dim BP_Gate_3: BP_Gate_3=Array(BM_Gate_3, LM_GI_Gate_3)
Dim BP_Gate_4: BP_Gate_4=Array(BM_Gate_4, LM_L_L10_Gate_4, LM_GI_Gate_4)
Dim BP_Gate_5: BP_Gate_5=Array(BM_Gate_5, LM_L_L18_Gate_5, LM_L_L35_Gate_5, LM_GI_Gate_5)
Dim BP_Gate_6: BP_Gate_6=Array(BM_Gate_6, LM_L_L14_Gate_6, LM_GI_Gate_6)
Dim BP_HT: BP_HT=Array(BM_HT, LM_L_L19_HT, LM_L_L39_HT, LM_GI_HT)
Dim BP_Inserts: BP_Inserts=Array(BM_Inserts, LM_L_L10_Inserts, LM_L_L11_Inserts, LM_L_L12_Inserts, LM_L_L13_Inserts, LM_L_L14_Inserts, LM_L_L15_Inserts, LM_L_L16_Inserts, LM_L_L17_Inserts, LM_L_L18_Inserts, LM_L_L19_Inserts, LM_L_L20_Inserts, LM_L_L21_Inserts, LM_L_L22_Inserts, LM_L_L23_Inserts, LM_L_L24_Inserts, LM_L_L25_Inserts, LM_L_L26_Inserts, LM_L_L27_Inserts, LM_L_L28_Inserts, LM_L_L29_Inserts, LM_L_L30_Inserts, LM_L_L31_Inserts, LM_L_L32_Inserts, LM_L_L33_Inserts, LM_L_L34_Inserts, LM_L_L35_Inserts, LM_L_L36_Inserts, LM_L_L37_Inserts, LM_L_L38_Inserts, LM_L_L39_Inserts, LM_L_L40_Inserts, LM_L_L41_Inserts, LM_L_L42_Inserts, LM_L_L43_Inserts, LM_L_L44_Inserts, LM_L_L45_Inserts, LM_L_L46_Inserts, LM_L_L47_Inserts, LM_L_L49_Inserts, LM_L_L50_Inserts, LM_L_L51_Inserts, LM_L_L52_Inserts, LM_L_L53_Inserts, LM_L_L54_Inserts, LM_L_L55_Inserts, LM_L_L57_Inserts, LM_L_L58_Inserts, LM_L_L59_Inserts, LM_L_L60_Inserts, LM_L_L61_Inserts, LM_L_L62_Inserts, LM_L_L63_Inserts, LM_L_L8_Inserts, LM_L_L9_Inserts, _
  LM_GIS_gi02_Inserts, LM_GIS_gi03_Inserts, LM_GIS_gi04_Inserts, LM_GI_Inserts)
Dim BP_LSling: BP_LSling=Array(BM_LSling, LM_L_L40_LSling, LM_L_L41_LSling, LM_L_L49_LSling, LM_L_L57_LSling, LM_GIS_gi01_LSling, LM_GIS_gi02_LSling, LM_GIS_gi03_LSling, LM_GIS_gi04_LSling, LM_GI_LSling)
Dim BP_LSling1: BP_LSling1=Array(BM_LSling1, LM_L_L31_LSling1, LM_L_L40_LSling1, LM_L_L41_LSling1, LM_L_L49_LSling1, LM_L_L57_LSling1, LM_GIS_gi01_LSling1, LM_GIS_gi02_LSling1, LM_GIS_gi03_LSling1, LM_GIS_gi04_LSling1, LM_GI_LSling1)
Dim BP_LSling2: BP_LSling2=Array(BM_LSling2, LM_L_L31_LSling2, LM_L_L40_LSling2, LM_L_L41_LSling2, LM_L_L49_LSling2, LM_L_L57_LSling2, LM_GIS_gi01_LSling2, LM_GIS_gi02_LSling2, LM_GIS_gi03_LSling2, LM_GIS_gi04_LSling2, LM_GI_LSling2)
Dim BP_LSlingArm: BP_LSlingArm=Array(BM_LSlingArm, LM_GIS_gi01_LSlingArm, LM_GIS_gi02_LSlingArm, LM_GI_LSlingArm)
Dim BP_Over: BP_Over=Array(BM_Over, LM_GI_Over)
Dim BP_Parts: BP_Parts=Array(BM_Parts, LM_L_L10_Parts, LM_L_L11_Parts, LM_L_L12_Parts, LM_L_L13_Parts, LM_L_L14_Parts, LM_L_L15_Parts, LM_L_L16_Parts, LM_L_L17_Parts, LM_L_L18_Parts, LM_L_L19_Parts, LM_L_L20_Parts, LM_L_L21_Parts, LM_L_L22_Parts, LM_L_L23_Parts, LM_L_L24_Parts, LM_L_L25_Parts, LM_L_L26_Parts, LM_L_L27_Parts, LM_L_L28_Parts, LM_L_L29_Parts, LM_L_L30_Parts, LM_L_L31_Parts, LM_L_L32_Parts, LM_L_L33_Parts, LM_L_L34_Parts, LM_L_L35_Parts, LM_L_L36_Parts, LM_L_L37_Parts, LM_L_L38_Parts, LM_L_L39_Parts, LM_L_L40_Parts, LM_L_L41_Parts, LM_L_L42_Parts, LM_L_L43_Parts, LM_L_L44_Parts, LM_L_L46_Parts, LM_L_L47_Parts, LM_L_L49_Parts, LM_L_L50_Parts, LM_L_L51_Parts, LM_L_L52_Parts, LM_L_L53_Parts, LM_L_L54_Parts, LM_L_L55_Parts, LM_L_L59_Parts, LM_L_L60_Parts, LM_L_L61_Parts, LM_L_L62_Parts, LM_L_L63_Parts, LM_L_L8_Parts, LM_L_L9_Parts, LM_GIS_gi01_Parts, LM_GIS_gi02_Parts, LM_GIS_gi03_Parts, LM_GIS_gi04_Parts, LM_GI_Parts)
Dim BP_Playfield: BP_Playfield=Array(BM_Playfield, LM_L_L10_Playfield, LM_L_L11_Playfield, LM_L_L12_Playfield, LM_L_L13_Playfield, LM_L_L14_Playfield, LM_L_L15_Playfield, LM_L_L16_Playfield, LM_L_L17_Playfield, LM_L_L18_Playfield, LM_L_L19_Playfield, LM_L_L20_Playfield, LM_L_L21_Playfield, LM_L_L22_Playfield, LM_L_L23_Playfield, LM_L_L24_Playfield, LM_L_L25_Playfield, LM_L_L26_Playfield, LM_L_L27_Playfield, LM_L_L28_Playfield, LM_L_L29_Playfield, LM_L_L30_Playfield, LM_L_L31_Playfield, LM_L_L32_Playfield, LM_L_L33_Playfield, LM_L_L34_Playfield, LM_L_L35_Playfield, LM_L_L36_Playfield, LM_L_L37_Playfield, LM_L_L38_Playfield, LM_L_L39_Playfield, LM_L_L40_Playfield, LM_L_L41_Playfield, LM_L_L42_Playfield, LM_L_L43_Playfield, LM_L_L44_Playfield, LM_L_L45_Playfield, LM_L_L46_Playfield, LM_L_L47_Playfield, LM_L_L49_Playfield, LM_L_L50_Playfield, LM_L_L51_Playfield, LM_L_L52_Playfield, LM_L_L53_Playfield, LM_L_L54_Playfield, LM_L_L55_Playfield, LM_L_L57_Playfield, LM_L_L58_Playfield, LM_L_L59_Playfield, _
  LM_L_L60_Playfield, LM_L_L61_Playfield, LM_L_L62_Playfield, LM_L_L63_Playfield, LM_L_L8_Playfield, LM_L_L9_Playfield, LM_GIS_gi01_Playfield, LM_GIS_gi02_Playfield, LM_GIS_gi03_Playfield, LM_GIS_gi04_Playfield, LM_GI_Playfield)
Dim BP_RSling: BP_RSling=Array(BM_RSling, LM_L_L15_RSling, LM_L_L38_RSling, LM_L_L47_RSling, LM_L_L55_RSling, LM_GIS_gi01_RSling, LM_GIS_gi02_RSling, LM_GIS_gi03_RSling, LM_GIS_gi04_RSling, LM_GI_RSling)
Dim BP_RSling1: BP_RSling1=Array(BM_RSling1, LM_L_L15_RSling1, LM_L_L38_RSling1, LM_L_L47_RSling1, LM_L_L55_RSling1, LM_GIS_gi01_RSling1, LM_GIS_gi02_RSling1, LM_GIS_gi03_RSling1, LM_GIS_gi04_RSling1, LM_GI_RSling1)
Dim BP_RSling2: BP_RSling2=Array(BM_RSling2, LM_L_L15_RSling2, LM_L_L38_RSling2, LM_L_L47_RSling2, LM_L_L55_RSling2, LM_L_L63_RSling2, LM_GIS_gi01_RSling2, LM_GIS_gi02_RSling2, LM_GIS_gi03_RSling2, LM_GIS_gi04_RSling2, LM_GI_RSling2)
Dim BP_RSlingArm: BP_RSlingArm=Array(BM_RSlingArm, LM_GIS_gi03_RSlingArm, LM_GIS_gi04_RSlingArm, LM_GI_RSlingArm)
Dim BP_Spin1: BP_Spin1=Array(BM_Spin1, LM_L_L9_Spin1, LM_GI_Spin1)
Dim BP_Spin1Rod: BP_Spin1Rod=Array(BM_Spin1Rod)
Dim BP_Spin2: BP_Spin2=Array(BM_Spin2, LM_L_L10_Spin2, LM_GI_Spin2)
Dim BP_Spin2Rod: BP_Spin2Rod=Array(BM_Spin2Rod)
Dim BP_Spin3: BP_Spin3=Array(BM_Spin3, LM_L_L11_Spin3, LM_GI_Spin3)
Dim BP_Spin3Rod: BP_Spin3Rod=Array(BM_Spin3Rod, LM_GI_Spin3Rod)
Dim BP_VRD: BP_VRD=Array(BM_VRD)
Dim BP_wire1: BP_wire1=Array(BM_wire1, LM_GIS_gi01_wire1)
Dim BP_wire2: BP_wire2=Array(BM_wire2, LM_GIS_gi01_wire2, LM_GIS_gi02_wire2)
Dim BP_wire3: BP_wire3=Array(BM_wire3, LM_GIS_gi03_wire3, LM_GIS_gi04_wire3)
Dim BP_wire4: BP_wire4=Array(BM_wire4, LM_GIS_gi03_wire4)
Dim BP_wire5: BP_wire5=Array(BM_wire5, LM_GI_wire5)
Dim BP_wire6: BP_wire6=Array(BM_wire6, LM_GI_wire6)
Dim BP_wire7: BP_wire7=Array(BM_wire7)
' Arrays per lighting scenario
Dim BL_GI: BL_GI=Array(LM_GI_BR1, LM_GI_BR2, LM_GI_BR3, LM_GI_Bumper1Skirt, LM_GI_Bumper2Skirt, LM_GI_Bumper3Skirt, LM_GI_DT1_1, LM_GI_DT1_2, LM_GI_DT1_3, LM_GI_DT2_1, LM_GI_DT2_2, LM_GI_DT2_3, LM_GI_DT3_1, LM_GI_DT3_2, LM_GI_DT3_3, LM_GI_Gate_1, LM_GI_Gate_3, LM_GI_Gate_4, LM_GI_Gate_5, LM_GI_Gate_6, LM_GI_HT, LM_GI_Inserts, LM_GI_LSling, LM_GI_LSling1, LM_GI_LSling2, LM_GI_LSlingArm, LM_GI_Over, LM_GI_Parts, LM_GI_Playfield, LM_GI_RSling, LM_GI_RSling1, LM_GI_RSling2, LM_GI_RSlingArm, LM_GI_Spin1, LM_GI_Spin2, LM_GI_Spin3, LM_GI_Spin3Rod, LM_GI_wire5, LM_GI_wire6)
Dim BL_GIS_gi01: BL_GIS_gi01=Array(LM_GIS_gi01_FlipperL, LM_GIS_gi01_FlipperLup, LM_GIS_gi01_FlipperR, LM_GIS_gi01_LSling, LM_GIS_gi01_LSling1, LM_GIS_gi01_LSling2, LM_GIS_gi01_LSlingArm, LM_GIS_gi01_Parts, LM_GIS_gi01_Playfield, LM_GIS_gi01_RSling, LM_GIS_gi01_RSling1, LM_GIS_gi01_RSling2, LM_GIS_gi01_wire1, LM_GIS_gi01_wire2)
Dim BL_GIS_gi02: BL_GIS_gi02=Array(LM_GIS_gi02_FlipperL, LM_GIS_gi02_FlipperLup, LM_GIS_gi02_Inserts, LM_GIS_gi02_LSling, LM_GIS_gi02_LSling1, LM_GIS_gi02_LSling2, LM_GIS_gi02_LSlingArm, LM_GIS_gi02_Parts, LM_GIS_gi02_Playfield, LM_GIS_gi02_RSling, LM_GIS_gi02_RSling1, LM_GIS_gi02_RSling2, LM_GIS_gi02_wire2)
Dim BL_GIS_gi03: BL_GIS_gi03=Array(LM_GIS_gi03_FlipperL, LM_GIS_gi03_FlipperLup, LM_GIS_gi03_FlipperR, LM_GIS_gi03_FlipperRup, LM_GIS_gi03_Inserts, LM_GIS_gi03_LSling, LM_GIS_gi03_LSling1, LM_GIS_gi03_LSling2, LM_GIS_gi03_Parts, LM_GIS_gi03_Playfield, LM_GIS_gi03_RSling, LM_GIS_gi03_RSling1, LM_GIS_gi03_RSling2, LM_GIS_gi03_RSlingArm, LM_GIS_gi03_wire3, LM_GIS_gi03_wire4)
Dim BL_GIS_gi04: BL_GIS_gi04=Array(LM_GIS_gi04_FlipperR, LM_GIS_gi04_FlipperRup, LM_GIS_gi04_Inserts, LM_GIS_gi04_LSling, LM_GIS_gi04_LSling1, LM_GIS_gi04_LSling2, LM_GIS_gi04_Parts, LM_GIS_gi04_Playfield, LM_GIS_gi04_RSling, LM_GIS_gi04_RSling1, LM_GIS_gi04_RSling2, LM_GIS_gi04_RSlingArm, LM_GIS_gi04_wire3)
Dim BL_L_L10: BL_L_L10=Array(LM_L_L10_Gate_4, LM_L_L10_Inserts, LM_L_L10_Parts, LM_L_L10_Playfield, LM_L_L10_Spin2)
Dim BL_L_L11: BL_L_L11=Array(LM_L_L11_Inserts, LM_L_L11_Parts, LM_L_L11_Playfield, LM_L_L11_Spin3)
Dim BL_L_L12: BL_L_L12=Array(LM_L_L12_Gate_1, LM_L_L12_Inserts, LM_L_L12_Parts, LM_L_L12_Playfield)
Dim BL_L_L13: BL_L_L13=Array(LM_L_L13_Inserts, LM_L_L13_Parts, LM_L_L13_Playfield)
Dim BL_L_L14: BL_L_L14=Array(LM_L_L14_Gate_6, LM_L_L14_Inserts, LM_L_L14_Parts, LM_L_L14_Playfield)
Dim BL_L_L15: BL_L_L15=Array(LM_L_L15_FlipperR, LM_L_L15_FlipperRup, LM_L_L15_Inserts, LM_L_L15_Parts, LM_L_L15_Playfield, LM_L_L15_RSling, LM_L_L15_RSling1, LM_L_L15_RSling2)
Dim BL_L_L16: BL_L_L16=Array(LM_L_L16_Inserts, LM_L_L16_Parts, LM_L_L16_Playfield)
Dim BL_L_L17: BL_L_L17=Array(LM_L_L17_Inserts, LM_L_L17_Parts, LM_L_L17_Playfield)
Dim BL_L_L18: BL_L_L18=Array(LM_L_L18_Gate_5, LM_L_L18_Inserts, LM_L_L18_Parts, LM_L_L18_Playfield)
Dim BL_L_L19: BL_L_L19=Array(LM_L_L19_BR1, LM_L_L19_Bumper1Skirt, LM_L_L19_HT, LM_L_L19_Inserts, LM_L_L19_Parts, LM_L_L19_Playfield)
Dim BL_L_L20: BL_L_L20=Array(LM_L_L20_DT1_1, LM_L_L20_DT1_2, LM_L_L20_DT1_3, LM_L_L20_Inserts, LM_L_L20_Parts, LM_L_L20_Playfield)
Dim BL_L_L21: BL_L_L21=Array(LM_L_L21_DT2_1, LM_L_L21_DT2_2, LM_L_L21_DT2_3, LM_L_L21_Inserts, LM_L_L21_Parts, LM_L_L21_Playfield)
Dim BL_L_L22: BL_L_L22=Array(LM_L_L22_DT3_1, LM_L_L22_DT3_2, LM_L_L22_DT3_3, LM_L_L22_Inserts, LM_L_L22_Parts, LM_L_L22_Playfield)
Dim BL_L_L23: BL_L_L23=Array(LM_L_L23_DT1_2, LM_L_L23_Inserts, LM_L_L23_Parts, LM_L_L23_Playfield)
Dim BL_L_L24: BL_L_L24=Array(LM_L_L24_DT2_1, LM_L_L24_DT2_2, LM_L_L24_DT2_3, LM_L_L24_Inserts, LM_L_L24_Parts, LM_L_L24_Playfield)
Dim BL_L_L25: BL_L_L25=Array(LM_L_L25_BR2, LM_L_L25_Bumper2Skirt, LM_L_L25_DT3_1, LM_L_L25_DT3_2, LM_L_L25_DT3_3, LM_L_L25_Inserts, LM_L_L25_Parts, LM_L_L25_Playfield)
Dim BL_L_L26: BL_L_L26=Array(LM_L_L26_Inserts, LM_L_L26_Parts, LM_L_L26_Playfield)
Dim BL_L_L27: BL_L_L27=Array(LM_L_L27_Inserts, LM_L_L27_Parts, LM_L_L27_Playfield)
Dim BL_L_L28: BL_L_L28=Array(LM_L_L28_Inserts, LM_L_L28_Parts, LM_L_L28_Playfield)
Dim BL_L_L29: BL_L_L29=Array(LM_L_L29_Inserts, LM_L_L29_Parts, LM_L_L29_Playfield)
Dim BL_L_L30: BL_L_L30=Array(LM_L_L30_Inserts, LM_L_L30_Parts, LM_L_L30_Playfield)
Dim BL_L_L31: BL_L_L31=Array(LM_L_L31_FlipperL, LM_L_L31_FlipperLup, LM_L_L31_Inserts, LM_L_L31_LSling1, LM_L_L31_LSling2, LM_L_L31_Parts, LM_L_L31_Playfield)
Dim BL_L_L32: BL_L_L32=Array(LM_L_L32_Inserts, LM_L_L32_Parts, LM_L_L32_Playfield)
Dim BL_L_L33: BL_L_L33=Array(LM_L_L33_Inserts, LM_L_L33_Parts, LM_L_L33_Playfield)
Dim BL_L_L34: BL_L_L34=Array(LM_L_L34_FlipperLup, LM_L_L34_FlipperR, LM_L_L34_FlipperRup, LM_L_L34_Inserts, LM_L_L34_Parts, LM_L_L34_Playfield)
Dim BL_L_L35: BL_L_L35=Array(LM_L_L35_Gate_5, LM_L_L35_Inserts, LM_L_L35_Parts, LM_L_L35_Playfield)
Dim BL_L_L36: BL_L_L36=Array(LM_L_L36_Inserts, LM_L_L36_Parts, LM_L_L36_Playfield)
Dim BL_L_L37: BL_L_L37=Array(LM_L_L37_Inserts, LM_L_L37_Parts, LM_L_L37_Playfield)
Dim BL_L_L38: BL_L_L38=Array(LM_L_L38_Inserts, LM_L_L38_Parts, LM_L_L38_Playfield, LM_L_L38_RSling, LM_L_L38_RSling1, LM_L_L38_RSling2)
Dim BL_L_L39: BL_L_L39=Array(LM_L_L39_BR1, LM_L_L39_Bumper1Skirt, LM_L_L39_HT, LM_L_L39_Inserts, LM_L_L39_Parts, LM_L_L39_Playfield)
Dim BL_L_L40: BL_L_L40=Array(LM_L_L40_Inserts, LM_L_L40_LSling, LM_L_L40_LSling1, LM_L_L40_LSling2, LM_L_L40_Parts, LM_L_L40_Playfield)
Dim BL_L_L41: BL_L_L41=Array(LM_L_L41_Inserts, LM_L_L41_LSling, LM_L_L41_LSling1, LM_L_L41_LSling2, LM_L_L41_Parts, LM_L_L41_Playfield)
Dim BL_L_L42: BL_L_L42=Array(LM_L_L42_Inserts, LM_L_L42_Parts, LM_L_L42_Playfield)
Dim BL_L_L43: BL_L_L43=Array(LM_L_L43_Inserts, LM_L_L43_Parts, LM_L_L43_Playfield)
Dim BL_L_L44: BL_L_L44=Array(LM_L_L44_Inserts, LM_L_L44_Parts, LM_L_L44_Playfield)
Dim BL_L_L45: BL_L_L45=Array(LM_L_L45_Inserts, LM_L_L45_Playfield)
Dim BL_L_L46: BL_L_L46=Array(LM_L_L46_Inserts, LM_L_L46_Parts, LM_L_L46_Playfield)
Dim BL_L_L47: BL_L_L47=Array(LM_L_L47_Inserts, LM_L_L47_Parts, LM_L_L47_Playfield, LM_L_L47_RSling, LM_L_L47_RSling1, LM_L_L47_RSling2)
Dim BL_L_L49: BL_L_L49=Array(LM_L_L49_Inserts, LM_L_L49_LSling, LM_L_L49_LSling1, LM_L_L49_LSling2, LM_L_L49_Parts, LM_L_L49_Playfield)
Dim BL_L_L50: BL_L_L50=Array(LM_L_L50_Inserts, LM_L_L50_Parts, LM_L_L50_Playfield)
Dim BL_L_L51: BL_L_L51=Array(LM_L_L51_Inserts, LM_L_L51_Parts, LM_L_L51_Playfield)
Dim BL_L_L52: BL_L_L52=Array(LM_L_L52_Inserts, LM_L_L52_Parts, LM_L_L52_Playfield)
Dim BL_L_L53: BL_L_L53=Array(LM_L_L53_Inserts, LM_L_L53_Parts, LM_L_L53_Playfield)
Dim BL_L_L54: BL_L_L54=Array(LM_L_L54_Inserts, LM_L_L54_Parts, LM_L_L54_Playfield)
Dim BL_L_L55: BL_L_L55=Array(LM_L_L55_Inserts, LM_L_L55_Parts, LM_L_L55_Playfield, LM_L_L55_RSling, LM_L_L55_RSling1, LM_L_L55_RSling2)
Dim BL_L_L57: BL_L_L57=Array(LM_L_L57_Inserts, LM_L_L57_LSling, LM_L_L57_LSling1, LM_L_L57_LSling2, LM_L_L57_Playfield)
Dim BL_L_L58: BL_L_L58=Array(LM_L_L58_Inserts, LM_L_L58_Playfield)
Dim BL_L_L59: BL_L_L59=Array(LM_L_L59_FlipperR, LM_L_L59_Inserts, LM_L_L59_Parts, LM_L_L59_Playfield)
Dim BL_L_L60: BL_L_L60=Array(LM_L_L60_Inserts, LM_L_L60_Parts, LM_L_L60_Playfield)
Dim BL_L_L61: BL_L_L61=Array(LM_L_L61_Inserts, LM_L_L61_Parts, LM_L_L61_Playfield)
Dim BL_L_L62: BL_L_L62=Array(LM_L_L62_Inserts, LM_L_L62_Parts, LM_L_L62_Playfield)
Dim BL_L_L63: BL_L_L63=Array(LM_L_L63_Inserts, LM_L_L63_Parts, LM_L_L63_Playfield, LM_L_L63_RSling2)
Dim BL_L_L8: BL_L_L8=Array(LM_L_L8_Inserts, LM_L_L8_Parts, LM_L_L8_Playfield)
Dim BL_L_L9: BL_L_L9=Array(LM_L_L9_Inserts, LM_L_L9_Parts, LM_L_L9_Playfield, LM_L_L9_Spin1)
Dim BL_Room: BL_Room=Array(BM_BR1, BM_BR2, BM_BR3, BM_Bumper1Skirt, BM_Bumper2Skirt, BM_Bumper3Skirt, BM_DT1_1, BM_DT1_2, BM_DT1_3, BM_DT2_1, BM_DT2_2, BM_DT2_3, BM_DT3_1, BM_DT3_2, BM_DT3_3, BM_FlipperL, BM_FlipperLup, BM_FlipperR, BM_FlipperRup, BM_Gate_1, BM_Gate_2, BM_Gate_3, BM_Gate_4, BM_Gate_5, BM_Gate_6, BM_HT, BM_Inserts, BM_LSling, BM_LSling1, BM_LSling2, BM_LSlingArm, BM_Over, BM_Parts, BM_Playfield, BM_RSling, BM_RSling1, BM_RSling2, BM_RSlingArm, BM_Spin1, BM_Spin1Rod, BM_Spin2, BM_Spin2Rod, BM_Spin3, BM_Spin3Rod, BM_VRD, BM_wire1, BM_wire2, BM_wire3, BM_wire4, BM_wire5, BM_wire6, BM_wire7)
' Global arrays
Dim BG_Bakemap: BG_Bakemap=Array(BM_BR1, BM_BR2, BM_BR3, BM_Bumper1Skirt, BM_Bumper2Skirt, BM_Bumper3Skirt, BM_DT1_1, BM_DT1_2, BM_DT1_3, BM_DT2_1, BM_DT2_2, BM_DT2_3, BM_DT3_1, BM_DT3_2, BM_DT3_3, BM_FlipperL, BM_FlipperLup, BM_FlipperR, BM_FlipperRup, BM_Gate_1, BM_Gate_2, BM_Gate_3, BM_Gate_4, BM_Gate_5, BM_Gate_6, BM_HT, BM_Inserts, BM_LSling, BM_LSling1, BM_LSling2, BM_LSlingArm, BM_Over, BM_Parts, BM_Playfield, BM_RSling, BM_RSling1, BM_RSling2, BM_RSlingArm, BM_Spin1, BM_Spin1Rod, BM_Spin2, BM_Spin2Rod, BM_Spin3, BM_Spin3Rod, BM_VRD, BM_wire1, BM_wire2, BM_wire3, BM_wire4, BM_wire5, BM_wire6, BM_wire7)
Dim BG_Lightmap: BG_Lightmap=Array(LM_GI_BR1, LM_GI_BR2, LM_GI_BR3, LM_GI_Bumper1Skirt, LM_GI_Bumper2Skirt, LM_GI_Bumper3Skirt, LM_GI_DT1_1, LM_GI_DT1_2, LM_GI_DT1_3, LM_GI_DT2_1, LM_GI_DT2_2, LM_GI_DT2_3, LM_GI_DT3_1, LM_GI_DT3_2, LM_GI_DT3_3, LM_GI_Gate_1, LM_GI_Gate_3, LM_GI_Gate_4, LM_GI_Gate_5, LM_GI_Gate_6, LM_GI_HT, LM_GI_Inserts, LM_GI_LSling, LM_GI_LSling1, LM_GI_LSling2, LM_GI_LSlingArm, LM_GI_Over, LM_GI_Parts, LM_GI_Playfield, LM_GI_RSling, LM_GI_RSling1, LM_GI_RSling2, LM_GI_RSlingArm, LM_GI_Spin1, LM_GI_Spin2, LM_GI_Spin3, LM_GI_Spin3Rod, LM_GI_wire5, LM_GI_wire6, LM_GIS_gi01_FlipperL, LM_GIS_gi01_FlipperLup, LM_GIS_gi01_FlipperR, LM_GIS_gi01_LSling, LM_GIS_gi01_LSling1, LM_GIS_gi01_LSling2, LM_GIS_gi01_LSlingArm, LM_GIS_gi01_Parts, LM_GIS_gi01_Playfield, LM_GIS_gi01_RSling, LM_GIS_gi01_RSling1, LM_GIS_gi01_RSling2, LM_GIS_gi01_wire1, LM_GIS_gi01_wire2, LM_GIS_gi02_FlipperL, LM_GIS_gi02_FlipperLup, LM_GIS_gi02_Inserts, LM_GIS_gi02_LSling, LM_GIS_gi02_LSling1, LM_GIS_gi02_LSling2, _
  LM_GIS_gi02_LSlingArm, LM_GIS_gi02_Parts, LM_GIS_gi02_Playfield, LM_GIS_gi02_RSling, LM_GIS_gi02_RSling1, LM_GIS_gi02_RSling2, LM_GIS_gi02_wire2, LM_GIS_gi03_FlipperL, LM_GIS_gi03_FlipperLup, LM_GIS_gi03_FlipperR, LM_GIS_gi03_FlipperRup, LM_GIS_gi03_Inserts, LM_GIS_gi03_LSling, LM_GIS_gi03_LSling1, LM_GIS_gi03_LSling2, LM_GIS_gi03_Parts, LM_GIS_gi03_Playfield, LM_GIS_gi03_RSling, LM_GIS_gi03_RSling1, LM_GIS_gi03_RSling2, LM_GIS_gi03_RSlingArm, LM_GIS_gi03_wire3, LM_GIS_gi03_wire4, LM_GIS_gi04_FlipperR, LM_GIS_gi04_FlipperRup, LM_GIS_gi04_Inserts, LM_GIS_gi04_LSling, LM_GIS_gi04_LSling1, LM_GIS_gi04_LSling2, LM_GIS_gi04_Parts, LM_GIS_gi04_Playfield, LM_GIS_gi04_RSling, LM_GIS_gi04_RSling1, LM_GIS_gi04_RSling2, LM_GIS_gi04_RSlingArm, LM_GIS_gi04_wire3, LM_L_L10_Gate_4, LM_L_L10_Inserts, LM_L_L10_Parts, LM_L_L10_Playfield, LM_L_L10_Spin2, LM_L_L11_Inserts, LM_L_L11_Parts, LM_L_L11_Playfield, LM_L_L11_Spin3, LM_L_L12_Gate_1, LM_L_L12_Inserts, LM_L_L12_Parts, LM_L_L12_Playfield, LM_L_L13_Inserts, LM_L_L13_Parts, _
  LM_L_L13_Playfield, LM_L_L14_Gate_6, LM_L_L14_Inserts, LM_L_L14_Parts, LM_L_L14_Playfield, LM_L_L15_FlipperR, LM_L_L15_FlipperRup, LM_L_L15_Inserts, LM_L_L15_Parts, LM_L_L15_Playfield, LM_L_L15_RSling, LM_L_L15_RSling1, LM_L_L15_RSling2, LM_L_L16_Inserts, LM_L_L16_Parts, LM_L_L16_Playfield, LM_L_L17_Inserts, LM_L_L17_Parts, LM_L_L17_Playfield, LM_L_L18_Gate_5, LM_L_L18_Inserts, LM_L_L18_Parts, LM_L_L18_Playfield, LM_L_L19_BR1, LM_L_L19_Bumper1Skirt, LM_L_L19_HT, LM_L_L19_Inserts, LM_L_L19_Parts, LM_L_L19_Playfield, LM_L_L20_DT1_1, LM_L_L20_DT1_2, LM_L_L20_DT1_3, LM_L_L20_Inserts, LM_L_L20_Parts, LM_L_L20_Playfield, LM_L_L21_DT2_1, LM_L_L21_DT2_2, LM_L_L21_DT2_3, LM_L_L21_Inserts, LM_L_L21_Parts, LM_L_L21_Playfield, LM_L_L22_DT3_1, LM_L_L22_DT3_2, LM_L_L22_DT3_3, LM_L_L22_Inserts, LM_L_L22_Parts, LM_L_L22_Playfield, LM_L_L23_DT1_2, LM_L_L23_Inserts, LM_L_L23_Parts, LM_L_L23_Playfield, LM_L_L24_DT2_1, LM_L_L24_DT2_2, LM_L_L24_DT2_3, LM_L_L24_Inserts, LM_L_L24_Parts, LM_L_L24_Playfield, LM_L_L25_BR2, _
  LM_L_L25_Bumper2Skirt, LM_L_L25_DT3_1, LM_L_L25_DT3_2, LM_L_L25_DT3_3, LM_L_L25_Inserts, LM_L_L25_Parts, LM_L_L25_Playfield, LM_L_L26_Inserts, LM_L_L26_Parts, LM_L_L26_Playfield, LM_L_L27_Inserts, LM_L_L27_Parts, LM_L_L27_Playfield, LM_L_L28_Inserts, LM_L_L28_Parts, LM_L_L28_Playfield, LM_L_L29_Inserts, LM_L_L29_Parts, LM_L_L29_Playfield, LM_L_L30_Inserts, LM_L_L30_Parts, LM_L_L30_Playfield, LM_L_L31_FlipperL, LM_L_L31_FlipperLup, LM_L_L31_Inserts, LM_L_L31_LSling1, LM_L_L31_LSling2, LM_L_L31_Parts, LM_L_L31_Playfield, LM_L_L32_Inserts, LM_L_L32_Parts, LM_L_L32_Playfield, LM_L_L33_Inserts, LM_L_L33_Parts, LM_L_L33_Playfield, LM_L_L34_FlipperLup, LM_L_L34_FlipperR, LM_L_L34_FlipperRup, LM_L_L34_Inserts, LM_L_L34_Parts, LM_L_L34_Playfield, LM_L_L35_Gate_5, LM_L_L35_Inserts, LM_L_L35_Parts, LM_L_L35_Playfield, LM_L_L36_Inserts, LM_L_L36_Parts, LM_L_L36_Playfield, LM_L_L37_Inserts, LM_L_L37_Parts, LM_L_L37_Playfield, LM_L_L38_Inserts, LM_L_L38_Parts, LM_L_L38_Playfield, LM_L_L38_RSling, LM_L_L38_RSling1, _
  LM_L_L38_RSling2, LM_L_L39_BR1, LM_L_L39_Bumper1Skirt, LM_L_L39_HT, LM_L_L39_Inserts, LM_L_L39_Parts, LM_L_L39_Playfield, LM_L_L40_Inserts, LM_L_L40_LSling, LM_L_L40_LSling1, LM_L_L40_LSling2, LM_L_L40_Parts, LM_L_L40_Playfield, LM_L_L41_Inserts, LM_L_L41_LSling, LM_L_L41_LSling1, LM_L_L41_LSling2, LM_L_L41_Parts, LM_L_L41_Playfield, LM_L_L42_Inserts, LM_L_L42_Parts, LM_L_L42_Playfield, LM_L_L43_Inserts, LM_L_L43_Parts, LM_L_L43_Playfield, LM_L_L44_Inserts, LM_L_L44_Parts, LM_L_L44_Playfield, LM_L_L45_Inserts, LM_L_L45_Playfield, LM_L_L46_Inserts, LM_L_L46_Parts, LM_L_L46_Playfield, LM_L_L47_Inserts, LM_L_L47_Parts, LM_L_L47_Playfield, LM_L_L47_RSling, LM_L_L47_RSling1, LM_L_L47_RSling2, LM_L_L49_Inserts, LM_L_L49_LSling, LM_L_L49_LSling1, LM_L_L49_LSling2, LM_L_L49_Parts, LM_L_L49_Playfield, LM_L_L50_Inserts, LM_L_L50_Parts, LM_L_L50_Playfield, LM_L_L51_Inserts, LM_L_L51_Parts, LM_L_L51_Playfield, LM_L_L52_Inserts, LM_L_L52_Parts, LM_L_L52_Playfield, LM_L_L53_Inserts, LM_L_L53_Parts, LM_L_L53_Playfield, _
  LM_L_L54_Inserts, LM_L_L54_Parts, LM_L_L54_Playfield, LM_L_L55_Inserts, LM_L_L55_Parts, LM_L_L55_Playfield, LM_L_L55_RSling, LM_L_L55_RSling1, LM_L_L55_RSling2, LM_L_L57_Inserts, LM_L_L57_LSling, LM_L_L57_LSling1, LM_L_L57_LSling2, LM_L_L57_Playfield, LM_L_L58_Inserts, LM_L_L58_Playfield, LM_L_L59_FlipperR, LM_L_L59_Inserts, LM_L_L59_Parts, LM_L_L59_Playfield, LM_L_L60_Inserts, LM_L_L60_Parts, LM_L_L60_Playfield, LM_L_L61_Inserts, LM_L_L61_Parts, LM_L_L61_Playfield, LM_L_L62_Inserts, LM_L_L62_Parts, LM_L_L62_Playfield, LM_L_L63_Inserts, LM_L_L63_Parts, LM_L_L63_Playfield, LM_L_L63_RSling2, LM_L_L8_Inserts, LM_L_L8_Parts, LM_L_L8_Playfield, LM_L_L9_Inserts, LM_L_L9_Parts, LM_L_L9_Playfield, LM_L_L9_Spin1)
Dim BG_All: BG_All=Array(BM_BR1, BM_BR2, BM_BR3, BM_Bumper1Skirt, BM_Bumper2Skirt, BM_Bumper3Skirt, BM_DT1_1, BM_DT1_2, BM_DT1_3, BM_DT2_1, BM_DT2_2, BM_DT2_3, BM_DT3_1, BM_DT3_2, BM_DT3_3, BM_FlipperL, BM_FlipperLup, BM_FlipperR, BM_FlipperRup, BM_Gate_1, BM_Gate_2, BM_Gate_3, BM_Gate_4, BM_Gate_5, BM_Gate_6, BM_HT, BM_Inserts, BM_LSling, BM_LSling1, BM_LSling2, BM_LSlingArm, BM_Over, BM_Parts, BM_Playfield, BM_RSling, BM_RSling1, BM_RSling2, BM_RSlingArm, BM_Spin1, BM_Spin1Rod, BM_Spin2, BM_Spin2Rod, BM_Spin3, BM_Spin3Rod, BM_VRD, BM_wire1, BM_wire2, BM_wire3, BM_wire4, BM_wire5, BM_wire6, BM_wire7, LM_GI_BR1, LM_GI_BR2, LM_GI_BR3, LM_GI_Bumper1Skirt, LM_GI_Bumper2Skirt, LM_GI_Bumper3Skirt, LM_GI_DT1_1, LM_GI_DT1_2, LM_GI_DT1_3, LM_GI_DT2_1, LM_GI_DT2_2, LM_GI_DT2_3, LM_GI_DT3_1, LM_GI_DT3_2, LM_GI_DT3_3, LM_GI_Gate_1, LM_GI_Gate_3, LM_GI_Gate_4, LM_GI_Gate_5, LM_GI_Gate_6, LM_GI_HT, LM_GI_Inserts, LM_GI_LSling, LM_GI_LSling1, LM_GI_LSling2, LM_GI_LSlingArm, LM_GI_Over, LM_GI_Parts, LM_GI_Playfield, _
  LM_GI_RSling, LM_GI_RSling1, LM_GI_RSling2, LM_GI_RSlingArm, LM_GI_Spin1, LM_GI_Spin2, LM_GI_Spin3, LM_GI_Spin3Rod, LM_GI_wire5, LM_GI_wire6, LM_GIS_gi01_FlipperL, LM_GIS_gi01_FlipperLup, LM_GIS_gi01_FlipperR, LM_GIS_gi01_LSling, LM_GIS_gi01_LSling1, LM_GIS_gi01_LSling2, LM_GIS_gi01_LSlingArm, LM_GIS_gi01_Parts, LM_GIS_gi01_Playfield, LM_GIS_gi01_RSling, LM_GIS_gi01_RSling1, LM_GIS_gi01_RSling2, LM_GIS_gi01_wire1, LM_GIS_gi01_wire2, LM_GIS_gi02_FlipperL, LM_GIS_gi02_FlipperLup, LM_GIS_gi02_Inserts, LM_GIS_gi02_LSling, LM_GIS_gi02_LSling1, LM_GIS_gi02_LSling2, LM_GIS_gi02_LSlingArm, LM_GIS_gi02_Parts, LM_GIS_gi02_Playfield, LM_GIS_gi02_RSling, LM_GIS_gi02_RSling1, LM_GIS_gi02_RSling2, LM_GIS_gi02_wire2, LM_GIS_gi03_FlipperL, LM_GIS_gi03_FlipperLup, LM_GIS_gi03_FlipperR, LM_GIS_gi03_FlipperRup, LM_GIS_gi03_Inserts, LM_GIS_gi03_LSling, LM_GIS_gi03_LSling1, LM_GIS_gi03_LSling2, LM_GIS_gi03_Parts, LM_GIS_gi03_Playfield, LM_GIS_gi03_RSling, LM_GIS_gi03_RSling1, LM_GIS_gi03_RSling2, LM_GIS_gi03_RSlingArm, _
  LM_GIS_gi03_wire3, LM_GIS_gi03_wire4, LM_GIS_gi04_FlipperR, LM_GIS_gi04_FlipperRup, LM_GIS_gi04_Inserts, LM_GIS_gi04_LSling, LM_GIS_gi04_LSling1, LM_GIS_gi04_LSling2, LM_GIS_gi04_Parts, LM_GIS_gi04_Playfield, LM_GIS_gi04_RSling, LM_GIS_gi04_RSling1, LM_GIS_gi04_RSling2, LM_GIS_gi04_RSlingArm, LM_GIS_gi04_wire3, LM_L_L10_Gate_4, LM_L_L10_Inserts, LM_L_L10_Parts, LM_L_L10_Playfield, LM_L_L10_Spin2, LM_L_L11_Inserts, LM_L_L11_Parts, LM_L_L11_Playfield, LM_L_L11_Spin3, LM_L_L12_Gate_1, LM_L_L12_Inserts, LM_L_L12_Parts, LM_L_L12_Playfield, LM_L_L13_Inserts, LM_L_L13_Parts, LM_L_L13_Playfield, LM_L_L14_Gate_6, LM_L_L14_Inserts, LM_L_L14_Parts, LM_L_L14_Playfield, LM_L_L15_FlipperR, LM_L_L15_FlipperRup, LM_L_L15_Inserts, LM_L_L15_Parts, LM_L_L15_Playfield, LM_L_L15_RSling, LM_L_L15_RSling1, LM_L_L15_RSling2, LM_L_L16_Inserts, LM_L_L16_Parts, LM_L_L16_Playfield, LM_L_L17_Inserts, LM_L_L17_Parts, LM_L_L17_Playfield, LM_L_L18_Gate_5, LM_L_L18_Inserts, LM_L_L18_Parts, LM_L_L18_Playfield, LM_L_L19_BR1, _
  LM_L_L19_Bumper1Skirt, LM_L_L19_HT, LM_L_L19_Inserts, LM_L_L19_Parts, LM_L_L19_Playfield, LM_L_L20_DT1_1, LM_L_L20_DT1_2, LM_L_L20_DT1_3, LM_L_L20_Inserts, LM_L_L20_Parts, LM_L_L20_Playfield, LM_L_L21_DT2_1, LM_L_L21_DT2_2, LM_L_L21_DT2_3, LM_L_L21_Inserts, LM_L_L21_Parts, LM_L_L21_Playfield, LM_L_L22_DT3_1, LM_L_L22_DT3_2, LM_L_L22_DT3_3, LM_L_L22_Inserts, LM_L_L22_Parts, LM_L_L22_Playfield, LM_L_L23_DT1_2, LM_L_L23_Inserts, LM_L_L23_Parts, LM_L_L23_Playfield, LM_L_L24_DT2_1, LM_L_L24_DT2_2, LM_L_L24_DT2_3, LM_L_L24_Inserts, LM_L_L24_Parts, LM_L_L24_Playfield, LM_L_L25_BR2, LM_L_L25_Bumper2Skirt, LM_L_L25_DT3_1, LM_L_L25_DT3_2, LM_L_L25_DT3_3, LM_L_L25_Inserts, LM_L_L25_Parts, LM_L_L25_Playfield, LM_L_L26_Inserts, LM_L_L26_Parts, LM_L_L26_Playfield, LM_L_L27_Inserts, LM_L_L27_Parts, LM_L_L27_Playfield, LM_L_L28_Inserts, LM_L_L28_Parts, LM_L_L28_Playfield, LM_L_L29_Inserts, LM_L_L29_Parts, LM_L_L29_Playfield, LM_L_L30_Inserts, LM_L_L30_Parts, LM_L_L30_Playfield, LM_L_L31_FlipperL, LM_L_L31_FlipperLup, _
  LM_L_L31_Inserts, LM_L_L31_LSling1, LM_L_L31_LSling2, LM_L_L31_Parts, LM_L_L31_Playfield, LM_L_L32_Inserts, LM_L_L32_Parts, LM_L_L32_Playfield, LM_L_L33_Inserts, LM_L_L33_Parts, LM_L_L33_Playfield, LM_L_L34_FlipperLup, LM_L_L34_FlipperR, LM_L_L34_FlipperRup, LM_L_L34_Inserts, LM_L_L34_Parts, LM_L_L34_Playfield, LM_L_L35_Gate_5, LM_L_L35_Inserts, LM_L_L35_Parts, LM_L_L35_Playfield, LM_L_L36_Inserts, LM_L_L36_Parts, LM_L_L36_Playfield, LM_L_L37_Inserts, LM_L_L37_Parts, LM_L_L37_Playfield, LM_L_L38_Inserts, LM_L_L38_Parts, LM_L_L38_Playfield, LM_L_L38_RSling, LM_L_L38_RSling1, LM_L_L38_RSling2, LM_L_L39_BR1, LM_L_L39_Bumper1Skirt, LM_L_L39_HT, LM_L_L39_Inserts, LM_L_L39_Parts, LM_L_L39_Playfield, LM_L_L40_Inserts, LM_L_L40_LSling, LM_L_L40_LSling1, LM_L_L40_LSling2, LM_L_L40_Parts, LM_L_L40_Playfield, LM_L_L41_Inserts, LM_L_L41_LSling, LM_L_L41_LSling1, LM_L_L41_LSling2, LM_L_L41_Parts, LM_L_L41_Playfield, LM_L_L42_Inserts, LM_L_L42_Parts, LM_L_L42_Playfield, LM_L_L43_Inserts, LM_L_L43_Parts, LM_L_L43_Playfield, _
  LM_L_L44_Inserts, LM_L_L44_Parts, LM_L_L44_Playfield, LM_L_L45_Inserts, LM_L_L45_Playfield, LM_L_L46_Inserts, LM_L_L46_Parts, LM_L_L46_Playfield, LM_L_L47_Inserts, LM_L_L47_Parts, LM_L_L47_Playfield, LM_L_L47_RSling, LM_L_L47_RSling1, LM_L_L47_RSling2, LM_L_L49_Inserts, LM_L_L49_LSling, LM_L_L49_LSling1, LM_L_L49_LSling2, LM_L_L49_Parts, LM_L_L49_Playfield, LM_L_L50_Inserts, LM_L_L50_Parts, LM_L_L50_Playfield, LM_L_L51_Inserts, LM_L_L51_Parts, LM_L_L51_Playfield, LM_L_L52_Inserts, LM_L_L52_Parts, LM_L_L52_Playfield, LM_L_L53_Inserts, LM_L_L53_Parts, LM_L_L53_Playfield, LM_L_L54_Inserts, LM_L_L54_Parts, LM_L_L54_Playfield, LM_L_L55_Inserts, LM_L_L55_Parts, LM_L_L55_Playfield, LM_L_L55_RSling, LM_L_L55_RSling1, LM_L_L55_RSling2, LM_L_L57_Inserts, LM_L_L57_LSling, LM_L_L57_LSling1, LM_L_L57_LSling2, LM_L_L57_Playfield, LM_L_L58_Inserts, LM_L_L58_Playfield, LM_L_L59_FlipperR, LM_L_L59_Inserts, LM_L_L59_Parts, LM_L_L59_Playfield, LM_L_L60_Inserts, LM_L_L60_Parts, LM_L_L60_Playfield, LM_L_L61_Inserts, _
  LM_L_L61_Parts, LM_L_L61_Playfield, LM_L_L62_Inserts, LM_L_L62_Parts, LM_L_L62_Playfield, LM_L_L63_Inserts, LM_L_L63_Parts, LM_L_L63_Playfield, LM_L_L63_RSling2, LM_L_L8_Inserts, LM_L_L8_Parts, LM_L_L8_Playfield, LM_L_L9_Inserts, LM_L_L9_Parts, LM_L_L9_Playfield, LM_L_L9_Spin1)
' VLM  Arrays - End

' VLM p1 Arrays - Start
' Arrays per baked part
Dim BP_p1p1: BP_p1p1=Array(p1BM_p1, p1LM_p1L_p1_1a_p1, p1LM_p1L_p1_1b_p1, p1LM_p1L_p1_1c_p1, p1LM_p1L_p1_1d_p1, p1LM_p1L_p1_1e_p1, p1LM_p1L_p1_1f_p1, p1LM_p1L_p1_1g_p1, p1LM_p1L_p1_2a_p1, p1LM_p1L_p1_2b_p1, p1LM_p1L_p1_2c_p1, p1LM_p1L_p1_2d_p1, p1LM_p1L_p1_2e_p1, p1LM_p1L_p1_2f_p1, p1LM_p1L_p1_2g_p1, p1LM_p1L_p1_3a_p1, p1LM_p1L_p1_3b_p1, p1LM_p1L_p1_3c_p1, p1LM_p1L_p1_3d_p1, p1LM_p1L_p1_3e_p1, p1LM_p1L_p1_3f_p1, p1LM_p1L_p1_3g_p1, p1LM_p1L_p1_4a_p1, p1LM_p1L_p1_4b_p1, p1LM_p1L_p1_4c_p1, p1LM_p1L_p1_4d_p1, p1LM_p1L_p1_4e_p1, p1LM_p1L_p1_4f_p1, p1LM_p1L_p1_4g_p1, p1LM_p1L_p1_5a_p1, p1LM_p1L_p1_5b_p1, p1LM_p1L_p1_5c_p1, p1LM_p1L_p1_5d_p1, p1LM_p1L_p1_5e_p1, p1LM_p1L_p1_5f_p1, p1LM_p1L_p1_5g_p1, p1LM_p1L_p1_6a_p1, p1LM_p1L_p1_6b_p1, p1LM_p1L_p1_6c_p1, p1LM_p1L_p1_6d_p1, p1LM_p1L_p1_6e_p1, p1LM_p1L_p1_6f_p1, p1LM_p1L_p1_6g_p1, p1LM_p1L_p1_7a_p1, p1LM_p1L_p1_7b_p1, p1LM_p1L_p1_7c_p1, p1LM_p1L_p1_7d_p1, p1LM_p1L_p1_7e_p1, p1LM_p1L_p1_7f_p1, p1LM_p1L_p1_7g_p1)
' Arrays per lighting scenario
Dim BL_p1Room: BL_p1Room=Array(p1BM_p1)
Dim BL_p1p1L_p1_1a: BL_p1p1L_p1_1a=Array(p1LM_p1L_p1_1a_p1)
Dim BL_p1p1L_p1_1b: BL_p1p1L_p1_1b=Array(p1LM_p1L_p1_1b_p1)
Dim BL_p1p1L_p1_1c: BL_p1p1L_p1_1c=Array(p1LM_p1L_p1_1c_p1)
Dim BL_p1p1L_p1_1d: BL_p1p1L_p1_1d=Array(p1LM_p1L_p1_1d_p1)
Dim BL_p1p1L_p1_1e: BL_p1p1L_p1_1e=Array(p1LM_p1L_p1_1e_p1)
Dim BL_p1p1L_p1_1f: BL_p1p1L_p1_1f=Array(p1LM_p1L_p1_1f_p1)
Dim BL_p1p1L_p1_1g: BL_p1p1L_p1_1g=Array(p1LM_p1L_p1_1g_p1)
Dim BL_p1p1L_p1_2a: BL_p1p1L_p1_2a=Array(p1LM_p1L_p1_2a_p1)
Dim BL_p1p1L_p1_2b: BL_p1p1L_p1_2b=Array(p1LM_p1L_p1_2b_p1)
Dim BL_p1p1L_p1_2c: BL_p1p1L_p1_2c=Array(p1LM_p1L_p1_2c_p1)
Dim BL_p1p1L_p1_2d: BL_p1p1L_p1_2d=Array(p1LM_p1L_p1_2d_p1)
Dim BL_p1p1L_p1_2e: BL_p1p1L_p1_2e=Array(p1LM_p1L_p1_2e_p1)
Dim BL_p1p1L_p1_2f: BL_p1p1L_p1_2f=Array(p1LM_p1L_p1_2f_p1)
Dim BL_p1p1L_p1_2g: BL_p1p1L_p1_2g=Array(p1LM_p1L_p1_2g_p1)
Dim BL_p1p1L_p1_3a: BL_p1p1L_p1_3a=Array(p1LM_p1L_p1_3a_p1)
Dim BL_p1p1L_p1_3b: BL_p1p1L_p1_3b=Array(p1LM_p1L_p1_3b_p1)
Dim BL_p1p1L_p1_3c: BL_p1p1L_p1_3c=Array(p1LM_p1L_p1_3c_p1)
Dim BL_p1p1L_p1_3d: BL_p1p1L_p1_3d=Array(p1LM_p1L_p1_3d_p1)
Dim BL_p1p1L_p1_3e: BL_p1p1L_p1_3e=Array(p1LM_p1L_p1_3e_p1)
Dim BL_p1p1L_p1_3f: BL_p1p1L_p1_3f=Array(p1LM_p1L_p1_3f_p1)
Dim BL_p1p1L_p1_3g: BL_p1p1L_p1_3g=Array(p1LM_p1L_p1_3g_p1)
Dim BL_p1p1L_p1_4a: BL_p1p1L_p1_4a=Array(p1LM_p1L_p1_4a_p1)
Dim BL_p1p1L_p1_4b: BL_p1p1L_p1_4b=Array(p1LM_p1L_p1_4b_p1)
Dim BL_p1p1L_p1_4c: BL_p1p1L_p1_4c=Array(p1LM_p1L_p1_4c_p1)
Dim BL_p1p1L_p1_4d: BL_p1p1L_p1_4d=Array(p1LM_p1L_p1_4d_p1)
Dim BL_p1p1L_p1_4e: BL_p1p1L_p1_4e=Array(p1LM_p1L_p1_4e_p1)
Dim BL_p1p1L_p1_4f: BL_p1p1L_p1_4f=Array(p1LM_p1L_p1_4f_p1)
Dim BL_p1p1L_p1_4g: BL_p1p1L_p1_4g=Array(p1LM_p1L_p1_4g_p1)
Dim BL_p1p1L_p1_5a: BL_p1p1L_p1_5a=Array(p1LM_p1L_p1_5a_p1)
Dim BL_p1p1L_p1_5b: BL_p1p1L_p1_5b=Array(p1LM_p1L_p1_5b_p1)
Dim BL_p1p1L_p1_5c: BL_p1p1L_p1_5c=Array(p1LM_p1L_p1_5c_p1)
Dim BL_p1p1L_p1_5d: BL_p1p1L_p1_5d=Array(p1LM_p1L_p1_5d_p1)
Dim BL_p1p1L_p1_5e: BL_p1p1L_p1_5e=Array(p1LM_p1L_p1_5e_p1)
Dim BL_p1p1L_p1_5f: BL_p1p1L_p1_5f=Array(p1LM_p1L_p1_5f_p1)
Dim BL_p1p1L_p1_5g: BL_p1p1L_p1_5g=Array(p1LM_p1L_p1_5g_p1)
Dim BL_p1p1L_p1_6a: BL_p1p1L_p1_6a=Array(p1LM_p1L_p1_6a_p1)
Dim BL_p1p1L_p1_6b: BL_p1p1L_p1_6b=Array(p1LM_p1L_p1_6b_p1)
Dim BL_p1p1L_p1_6c: BL_p1p1L_p1_6c=Array(p1LM_p1L_p1_6c_p1)
Dim BL_p1p1L_p1_6d: BL_p1p1L_p1_6d=Array(p1LM_p1L_p1_6d_p1)
Dim BL_p1p1L_p1_6e: BL_p1p1L_p1_6e=Array(p1LM_p1L_p1_6e_p1)
Dim BL_p1p1L_p1_6f: BL_p1p1L_p1_6f=Array(p1LM_p1L_p1_6f_p1)
Dim BL_p1p1L_p1_6g: BL_p1p1L_p1_6g=Array(p1LM_p1L_p1_6g_p1)
Dim BL_p1p1L_p1_7a: BL_p1p1L_p1_7a=Array(p1LM_p1L_p1_7a_p1)
Dim BL_p1p1L_p1_7b: BL_p1p1L_p1_7b=Array(p1LM_p1L_p1_7b_p1)
Dim BL_p1p1L_p1_7c: BL_p1p1L_p1_7c=Array(p1LM_p1L_p1_7c_p1)
Dim BL_p1p1L_p1_7d: BL_p1p1L_p1_7d=Array(p1LM_p1L_p1_7d_p1)
Dim BL_p1p1L_p1_7e: BL_p1p1L_p1_7e=Array(p1LM_p1L_p1_7e_p1)
Dim BL_p1p1L_p1_7f: BL_p1p1L_p1_7f=Array(p1LM_p1L_p1_7f_p1)
Dim BL_p1p1L_p1_7g: BL_p1p1L_p1_7g=Array(p1LM_p1L_p1_7g_p1)
' Global arrays
Dim BGp1_Bakemap: BGp1_Bakemap=Array(p1BM_p1)
Dim BGp1_Lightmap: BGp1_Lightmap=Array(p1LM_p1L_p1_1a_p1, p1LM_p1L_p1_1b_p1, p1LM_p1L_p1_1c_p1, p1LM_p1L_p1_1d_p1, p1LM_p1L_p1_1e_p1, p1LM_p1L_p1_1f_p1, p1LM_p1L_p1_1g_p1, p1LM_p1L_p1_2a_p1, p1LM_p1L_p1_2b_p1, p1LM_p1L_p1_2c_p1, p1LM_p1L_p1_2d_p1, p1LM_p1L_p1_2e_p1, p1LM_p1L_p1_2f_p1, p1LM_p1L_p1_2g_p1, p1LM_p1L_p1_3a_p1, p1LM_p1L_p1_3b_p1, p1LM_p1L_p1_3c_p1, p1LM_p1L_p1_3d_p1, p1LM_p1L_p1_3e_p1, p1LM_p1L_p1_3f_p1, p1LM_p1L_p1_3g_p1, p1LM_p1L_p1_4a_p1, p1LM_p1L_p1_4b_p1, p1LM_p1L_p1_4c_p1, p1LM_p1L_p1_4d_p1, p1LM_p1L_p1_4e_p1, p1LM_p1L_p1_4f_p1, p1LM_p1L_p1_4g_p1, p1LM_p1L_p1_5a_p1, p1LM_p1L_p1_5b_p1, p1LM_p1L_p1_5c_p1, p1LM_p1L_p1_5d_p1, p1LM_p1L_p1_5e_p1, p1LM_p1L_p1_5f_p1, p1LM_p1L_p1_5g_p1, p1LM_p1L_p1_6a_p1, p1LM_p1L_p1_6b_p1, p1LM_p1L_p1_6c_p1, p1LM_p1L_p1_6d_p1, p1LM_p1L_p1_6e_p1, p1LM_p1L_p1_6f_p1, p1LM_p1L_p1_6g_p1, p1LM_p1L_p1_7a_p1, p1LM_p1L_p1_7b_p1, p1LM_p1L_p1_7c_p1, p1LM_p1L_p1_7d_p1, p1LM_p1L_p1_7e_p1, p1LM_p1L_p1_7f_p1, p1LM_p1L_p1_7g_p1)
Dim BGp1_All: BGp1_All=Array(p1BM_p1, p1LM_p1L_p1_1a_p1, p1LM_p1L_p1_1b_p1, p1LM_p1L_p1_1c_p1, p1LM_p1L_p1_1d_p1, p1LM_p1L_p1_1e_p1, p1LM_p1L_p1_1f_p1, p1LM_p1L_p1_1g_p1, p1LM_p1L_p1_2a_p1, p1LM_p1L_p1_2b_p1, p1LM_p1L_p1_2c_p1, p1LM_p1L_p1_2d_p1, p1LM_p1L_p1_2e_p1, p1LM_p1L_p1_2f_p1, p1LM_p1L_p1_2g_p1, p1LM_p1L_p1_3a_p1, p1LM_p1L_p1_3b_p1, p1LM_p1L_p1_3c_p1, p1LM_p1L_p1_3d_p1, p1LM_p1L_p1_3e_p1, p1LM_p1L_p1_3f_p1, p1LM_p1L_p1_3g_p1, p1LM_p1L_p1_4a_p1, p1LM_p1L_p1_4b_p1, p1LM_p1L_p1_4c_p1, p1LM_p1L_p1_4d_p1, p1LM_p1L_p1_4e_p1, p1LM_p1L_p1_4f_p1, p1LM_p1L_p1_4g_p1, p1LM_p1L_p1_5a_p1, p1LM_p1L_p1_5b_p1, p1LM_p1L_p1_5c_p1, p1LM_p1L_p1_5d_p1, p1LM_p1L_p1_5e_p1, p1LM_p1L_p1_5f_p1, p1LM_p1L_p1_5g_p1, p1LM_p1L_p1_6a_p1, p1LM_p1L_p1_6b_p1, p1LM_p1L_p1_6c_p1, p1LM_p1L_p1_6d_p1, p1LM_p1L_p1_6e_p1, p1LM_p1L_p1_6f_p1, p1LM_p1L_p1_6g_p1, p1LM_p1L_p1_7a_p1, p1LM_p1L_p1_7b_p1, p1LM_p1L_p1_7c_p1, p1LM_p1L_p1_7d_p1, p1LM_p1L_p1_7e_p1, p1LM_p1L_p1_7f_p1, p1LM_p1L_p1_7g_p1)
' VLM p1 Arrays - End

' VLM p2 Arrays - Start
' Arrays per baked part
Dim BP_p2p2: BP_p2p2=Array(p2BM_p2, p2LM_p2L_p2_1a_p2, p2LM_p2L_p2_1b_p2, p2LM_p2L_p2_1c_p2, p2LM_p2L_p2_1d_p2, p2LM_p2L_p2_1e_p2, p2LM_p2L_p2_1f_p2, p2LM_p2L_p2_1g_p2, p2LM_p2L_p2_2a_p2, p2LM_p2L_p2_2b_p2, p2LM_p2L_p2_2c_p2, p2LM_p2L_p2_2d_p2, p2LM_p2L_p2_2e_p2, p2LM_p2L_p2_2f_p2, p2LM_p2L_p2_2g_p2, p2LM_p2L_p2_3a_p2, p2LM_p2L_p2_3b_p2, p2LM_p2L_p2_3c_p2, p2LM_p2L_p2_3d_p2, p2LM_p2L_p2_3e_p2, p2LM_p2L_p2_3f_p2, p2LM_p2L_p2_3g_p2, p2LM_p2L_p2_4a_p2, p2LM_p2L_p2_4b_p2, p2LM_p2L_p2_4c_p2, p2LM_p2L_p2_4d_p2, p2LM_p2L_p2_4e_p2, p2LM_p2L_p2_4f_p2, p2LM_p2L_p2_4g_p2, p2LM_p2L_p2_5a_p2, p2LM_p2L_p2_5b_p2, p2LM_p2L_p2_5c_p2, p2LM_p2L_p2_5d_p2, p2LM_p2L_p2_5e_p2, p2LM_p2L_p2_5f_p2, p2LM_p2L_p2_5g_p2, p2LM_p2L_p2_6a_p2, p2LM_p2L_p2_6b_p2, p2LM_p2L_p2_6c_p2, p2LM_p2L_p2_6d_p2, p2LM_p2L_p2_6e_p2, p2LM_p2L_p2_6f_p2, p2LM_p2L_p2_6g_p2, p2LM_p2L_p2_7a_p2, p2LM_p2L_p2_7b_p2, p2LM_p2L_p2_7c_p2, p2LM_p2L_p2_7d_p2, p2LM_p2L_p2_7e_p2, p2LM_p2L_p2_7f_p2, p2LM_p2L_p2_7g_p2)
' Arrays per lighting scenario
Dim BL_p2Room: BL_p2Room=Array(p2BM_p2)
Dim BL_p2p2L_p2_1a: BL_p2p2L_p2_1a=Array(p2LM_p2L_p2_1a_p2)
Dim BL_p2p2L_p2_1b: BL_p2p2L_p2_1b=Array(p2LM_p2L_p2_1b_p2)
Dim BL_p2p2L_p2_1c: BL_p2p2L_p2_1c=Array(p2LM_p2L_p2_1c_p2)
Dim BL_p2p2L_p2_1d: BL_p2p2L_p2_1d=Array(p2LM_p2L_p2_1d_p2)
Dim BL_p2p2L_p2_1e: BL_p2p2L_p2_1e=Array(p2LM_p2L_p2_1e_p2)
Dim BL_p2p2L_p2_1f: BL_p2p2L_p2_1f=Array(p2LM_p2L_p2_1f_p2)
Dim BL_p2p2L_p2_1g: BL_p2p2L_p2_1g=Array(p2LM_p2L_p2_1g_p2)
Dim BL_p2p2L_p2_2a: BL_p2p2L_p2_2a=Array(p2LM_p2L_p2_2a_p2)
Dim BL_p2p2L_p2_2b: BL_p2p2L_p2_2b=Array(p2LM_p2L_p2_2b_p2)
Dim BL_p2p2L_p2_2c: BL_p2p2L_p2_2c=Array(p2LM_p2L_p2_2c_p2)
Dim BL_p2p2L_p2_2d: BL_p2p2L_p2_2d=Array(p2LM_p2L_p2_2d_p2)
Dim BL_p2p2L_p2_2e: BL_p2p2L_p2_2e=Array(p2LM_p2L_p2_2e_p2)
Dim BL_p2p2L_p2_2f: BL_p2p2L_p2_2f=Array(p2LM_p2L_p2_2f_p2)
Dim BL_p2p2L_p2_2g: BL_p2p2L_p2_2g=Array(p2LM_p2L_p2_2g_p2)
Dim BL_p2p2L_p2_3a: BL_p2p2L_p2_3a=Array(p2LM_p2L_p2_3a_p2)
Dim BL_p2p2L_p2_3b: BL_p2p2L_p2_3b=Array(p2LM_p2L_p2_3b_p2)
Dim BL_p2p2L_p2_3c: BL_p2p2L_p2_3c=Array(p2LM_p2L_p2_3c_p2)
Dim BL_p2p2L_p2_3d: BL_p2p2L_p2_3d=Array(p2LM_p2L_p2_3d_p2)
Dim BL_p2p2L_p2_3e: BL_p2p2L_p2_3e=Array(p2LM_p2L_p2_3e_p2)
Dim BL_p2p2L_p2_3f: BL_p2p2L_p2_3f=Array(p2LM_p2L_p2_3f_p2)
Dim BL_p2p2L_p2_3g: BL_p2p2L_p2_3g=Array(p2LM_p2L_p2_3g_p2)
Dim BL_p2p2L_p2_4a: BL_p2p2L_p2_4a=Array(p2LM_p2L_p2_4a_p2)
Dim BL_p2p2L_p2_4b: BL_p2p2L_p2_4b=Array(p2LM_p2L_p2_4b_p2)
Dim BL_p2p2L_p2_4c: BL_p2p2L_p2_4c=Array(p2LM_p2L_p2_4c_p2)
Dim BL_p2p2L_p2_4d: BL_p2p2L_p2_4d=Array(p2LM_p2L_p2_4d_p2)
Dim BL_p2p2L_p2_4e: BL_p2p2L_p2_4e=Array(p2LM_p2L_p2_4e_p2)
Dim BL_p2p2L_p2_4f: BL_p2p2L_p2_4f=Array(p2LM_p2L_p2_4f_p2)
Dim BL_p2p2L_p2_4g: BL_p2p2L_p2_4g=Array(p2LM_p2L_p2_4g_p2)
Dim BL_p2p2L_p2_5a: BL_p2p2L_p2_5a=Array(p2LM_p2L_p2_5a_p2)
Dim BL_p2p2L_p2_5b: BL_p2p2L_p2_5b=Array(p2LM_p2L_p2_5b_p2)
Dim BL_p2p2L_p2_5c: BL_p2p2L_p2_5c=Array(p2LM_p2L_p2_5c_p2)
Dim BL_p2p2L_p2_5d: BL_p2p2L_p2_5d=Array(p2LM_p2L_p2_5d_p2)
Dim BL_p2p2L_p2_5e: BL_p2p2L_p2_5e=Array(p2LM_p2L_p2_5e_p2)
Dim BL_p2p2L_p2_5f: BL_p2p2L_p2_5f=Array(p2LM_p2L_p2_5f_p2)
Dim BL_p2p2L_p2_5g: BL_p2p2L_p2_5g=Array(p2LM_p2L_p2_5g_p2)
Dim BL_p2p2L_p2_6a: BL_p2p2L_p2_6a=Array(p2LM_p2L_p2_6a_p2)
Dim BL_p2p2L_p2_6b: BL_p2p2L_p2_6b=Array(p2LM_p2L_p2_6b_p2)
Dim BL_p2p2L_p2_6c: BL_p2p2L_p2_6c=Array(p2LM_p2L_p2_6c_p2)
Dim BL_p2p2L_p2_6d: BL_p2p2L_p2_6d=Array(p2LM_p2L_p2_6d_p2)
Dim BL_p2p2L_p2_6e: BL_p2p2L_p2_6e=Array(p2LM_p2L_p2_6e_p2)
Dim BL_p2p2L_p2_6f: BL_p2p2L_p2_6f=Array(p2LM_p2L_p2_6f_p2)
Dim BL_p2p2L_p2_6g: BL_p2p2L_p2_6g=Array(p2LM_p2L_p2_6g_p2)
Dim BL_p2p2L_p2_7a: BL_p2p2L_p2_7a=Array(p2LM_p2L_p2_7a_p2)
Dim BL_p2p2L_p2_7b: BL_p2p2L_p2_7b=Array(p2LM_p2L_p2_7b_p2)
Dim BL_p2p2L_p2_7c: BL_p2p2L_p2_7c=Array(p2LM_p2L_p2_7c_p2)
Dim BL_p2p2L_p2_7d: BL_p2p2L_p2_7d=Array(p2LM_p2L_p2_7d_p2)
Dim BL_p2p2L_p2_7e: BL_p2p2L_p2_7e=Array(p2LM_p2L_p2_7e_p2)
Dim BL_p2p2L_p2_7f: BL_p2p2L_p2_7f=Array(p2LM_p2L_p2_7f_p2)
Dim BL_p2p2L_p2_7g: BL_p2p2L_p2_7g=Array(p2LM_p2L_p2_7g_p2)
' Global arrays
Dim BGp2_Bakemap: BGp2_Bakemap=Array(p2BM_p2)
Dim BGp2_Lightmap: BGp2_Lightmap=Array(p2LM_p2L_p2_1a_p2, p2LM_p2L_p2_1b_p2, p2LM_p2L_p2_1c_p2, p2LM_p2L_p2_1d_p2, p2LM_p2L_p2_1e_p2, p2LM_p2L_p2_1f_p2, p2LM_p2L_p2_1g_p2, p2LM_p2L_p2_2a_p2, p2LM_p2L_p2_2b_p2, p2LM_p2L_p2_2c_p2, p2LM_p2L_p2_2d_p2, p2LM_p2L_p2_2e_p2, p2LM_p2L_p2_2f_p2, p2LM_p2L_p2_2g_p2, p2LM_p2L_p2_3a_p2, p2LM_p2L_p2_3b_p2, p2LM_p2L_p2_3c_p2, p2LM_p2L_p2_3d_p2, p2LM_p2L_p2_3e_p2, p2LM_p2L_p2_3f_p2, p2LM_p2L_p2_3g_p2, p2LM_p2L_p2_4a_p2, p2LM_p2L_p2_4b_p2, p2LM_p2L_p2_4c_p2, p2LM_p2L_p2_4d_p2, p2LM_p2L_p2_4e_p2, p2LM_p2L_p2_4f_p2, p2LM_p2L_p2_4g_p2, p2LM_p2L_p2_5a_p2, p2LM_p2L_p2_5b_p2, p2LM_p2L_p2_5c_p2, p2LM_p2L_p2_5d_p2, p2LM_p2L_p2_5e_p2, p2LM_p2L_p2_5f_p2, p2LM_p2L_p2_5g_p2, p2LM_p2L_p2_6a_p2, p2LM_p2L_p2_6b_p2, p2LM_p2L_p2_6c_p2, p2LM_p2L_p2_6d_p2, p2LM_p2L_p2_6e_p2, p2LM_p2L_p2_6f_p2, p2LM_p2L_p2_6g_p2, p2LM_p2L_p2_7a_p2, p2LM_p2L_p2_7b_p2, p2LM_p2L_p2_7c_p2, p2LM_p2L_p2_7d_p2, p2LM_p2L_p2_7e_p2, p2LM_p2L_p2_7f_p2, p2LM_p2L_p2_7g_p2)
Dim BGp2_All: BGp2_All=Array(p2BM_p2, p2LM_p2L_p2_1a_p2, p2LM_p2L_p2_1b_p2, p2LM_p2L_p2_1c_p2, p2LM_p2L_p2_1d_p2, p2LM_p2L_p2_1e_p2, p2LM_p2L_p2_1f_p2, p2LM_p2L_p2_1g_p2, p2LM_p2L_p2_2a_p2, p2LM_p2L_p2_2b_p2, p2LM_p2L_p2_2c_p2, p2LM_p2L_p2_2d_p2, p2LM_p2L_p2_2e_p2, p2LM_p2L_p2_2f_p2, p2LM_p2L_p2_2g_p2, p2LM_p2L_p2_3a_p2, p2LM_p2L_p2_3b_p2, p2LM_p2L_p2_3c_p2, p2LM_p2L_p2_3d_p2, p2LM_p2L_p2_3e_p2, p2LM_p2L_p2_3f_p2, p2LM_p2L_p2_3g_p2, p2LM_p2L_p2_4a_p2, p2LM_p2L_p2_4b_p2, p2LM_p2L_p2_4c_p2, p2LM_p2L_p2_4d_p2, p2LM_p2L_p2_4e_p2, p2LM_p2L_p2_4f_p2, p2LM_p2L_p2_4g_p2, p2LM_p2L_p2_5a_p2, p2LM_p2L_p2_5b_p2, p2LM_p2L_p2_5c_p2, p2LM_p2L_p2_5d_p2, p2LM_p2L_p2_5e_p2, p2LM_p2L_p2_5f_p2, p2LM_p2L_p2_5g_p2, p2LM_p2L_p2_6a_p2, p2LM_p2L_p2_6b_p2, p2LM_p2L_p2_6c_p2, p2LM_p2L_p2_6d_p2, p2LM_p2L_p2_6e_p2, p2LM_p2L_p2_6f_p2, p2LM_p2L_p2_6g_p2, p2LM_p2L_p2_7a_p2, p2LM_p2L_p2_7b_p2, p2LM_p2L_p2_7c_p2, p2LM_p2L_p2_7d_p2, p2LM_p2L_p2_7e_p2, p2LM_p2L_p2_7f_p2, p2LM_p2L_p2_7g_p2)
' VLM p2 Arrays - End

' VLM p3 Arrays - Start
' Arrays per baked part
Dim BP_p3p3: BP_p3p3=Array(p3BM_p3, p3LM_p3L_p3_1a_p3, p3LM_p3L_p3_1b_p3, p3LM_p3L_p3_1c_p3, p3LM_p3L_p3_1d_p3, p3LM_p3L_p3_1e_p3, p3LM_p3L_p3_1f_p3, p3LM_p3L_p3_1g_p3, p3LM_p3L_p3_2a_p3, p3LM_p3L_p3_2b_p3, p3LM_p3L_p3_2c_p3, p3LM_p3L_p3_2d_p3, p3LM_p3L_p3_2e_p3, p3LM_p3L_p3_2f_p3, p3LM_p3L_p3_2g_p3, p3LM_p3L_p3_3a_p3, p3LM_p3L_p3_3b_p3, p3LM_p3L_p3_3c_p3, p3LM_p3L_p3_3d_p3, p3LM_p3L_p3_3e_p3, p3LM_p3L_p3_3f_p3, p3LM_p3L_p3_3g_p3, p3LM_p3L_p3_4a_p3, p3LM_p3L_p3_4b_p3, p3LM_p3L_p3_4c_p3, p3LM_p3L_p3_4d_p3, p3LM_p3L_p3_4e_p3, p3LM_p3L_p3_4f_p3, p3LM_p3L_p3_4g_p3, p3LM_p3L_p3_5a_p3, p3LM_p3L_p3_5b_p3, p3LM_p3L_p3_5c_p3, p3LM_p3L_p3_5d_p3, p3LM_p3L_p3_5e_p3, p3LM_p3L_p3_5f_p3, p3LM_p3L_p3_5g_p3, p3LM_p3L_p3_6a_p3, p3LM_p3L_p3_6b_p3, p3LM_p3L_p3_6c_p3, p3LM_p3L_p3_6d_p3, p3LM_p3L_p3_6e_p3, p3LM_p3L_p3_6f_p3, p3LM_p3L_p3_6g_p3, p3LM_p3L_p3_7a_p3, p3LM_p3L_p3_7b_p3, p3LM_p3L_p3_7c_p3, p3LM_p3L_p3_7d_p3, p3LM_p3L_p3_7e_p3, p3LM_p3L_p3_7f_p3, p3LM_p3L_p3_7g_p3)
' Arrays per lighting scenario
Dim BL_p3Room: BL_p3Room=Array(p3BM_p3)
Dim BL_p3p3L_p3_1a: BL_p3p3L_p3_1a=Array(p3LM_p3L_p3_1a_p3)
Dim BL_p3p3L_p3_1b: BL_p3p3L_p3_1b=Array(p3LM_p3L_p3_1b_p3)
Dim BL_p3p3L_p3_1c: BL_p3p3L_p3_1c=Array(p3LM_p3L_p3_1c_p3)
Dim BL_p3p3L_p3_1d: BL_p3p3L_p3_1d=Array(p3LM_p3L_p3_1d_p3)
Dim BL_p3p3L_p3_1e: BL_p3p3L_p3_1e=Array(p3LM_p3L_p3_1e_p3)
Dim BL_p3p3L_p3_1f: BL_p3p3L_p3_1f=Array(p3LM_p3L_p3_1f_p3)
Dim BL_p3p3L_p3_1g: BL_p3p3L_p3_1g=Array(p3LM_p3L_p3_1g_p3)
Dim BL_p3p3L_p3_2a: BL_p3p3L_p3_2a=Array(p3LM_p3L_p3_2a_p3)
Dim BL_p3p3L_p3_2b: BL_p3p3L_p3_2b=Array(p3LM_p3L_p3_2b_p3)
Dim BL_p3p3L_p3_2c: BL_p3p3L_p3_2c=Array(p3LM_p3L_p3_2c_p3)
Dim BL_p3p3L_p3_2d: BL_p3p3L_p3_2d=Array(p3LM_p3L_p3_2d_p3)
Dim BL_p3p3L_p3_2e: BL_p3p3L_p3_2e=Array(p3LM_p3L_p3_2e_p3)
Dim BL_p3p3L_p3_2f: BL_p3p3L_p3_2f=Array(p3LM_p3L_p3_2f_p3)
Dim BL_p3p3L_p3_2g: BL_p3p3L_p3_2g=Array(p3LM_p3L_p3_2g_p3)
Dim BL_p3p3L_p3_3a: BL_p3p3L_p3_3a=Array(p3LM_p3L_p3_3a_p3)
Dim BL_p3p3L_p3_3b: BL_p3p3L_p3_3b=Array(p3LM_p3L_p3_3b_p3)
Dim BL_p3p3L_p3_3c: BL_p3p3L_p3_3c=Array(p3LM_p3L_p3_3c_p3)
Dim BL_p3p3L_p3_3d: BL_p3p3L_p3_3d=Array(p3LM_p3L_p3_3d_p3)
Dim BL_p3p3L_p3_3e: BL_p3p3L_p3_3e=Array(p3LM_p3L_p3_3e_p3)
Dim BL_p3p3L_p3_3f: BL_p3p3L_p3_3f=Array(p3LM_p3L_p3_3f_p3)
Dim BL_p3p3L_p3_3g: BL_p3p3L_p3_3g=Array(p3LM_p3L_p3_3g_p3)
Dim BL_p3p3L_p3_4a: BL_p3p3L_p3_4a=Array(p3LM_p3L_p3_4a_p3)
Dim BL_p3p3L_p3_4b: BL_p3p3L_p3_4b=Array(p3LM_p3L_p3_4b_p3)
Dim BL_p3p3L_p3_4c: BL_p3p3L_p3_4c=Array(p3LM_p3L_p3_4c_p3)
Dim BL_p3p3L_p3_4d: BL_p3p3L_p3_4d=Array(p3LM_p3L_p3_4d_p3)
Dim BL_p3p3L_p3_4e: BL_p3p3L_p3_4e=Array(p3LM_p3L_p3_4e_p3)
Dim BL_p3p3L_p3_4f: BL_p3p3L_p3_4f=Array(p3LM_p3L_p3_4f_p3)
Dim BL_p3p3L_p3_4g: BL_p3p3L_p3_4g=Array(p3LM_p3L_p3_4g_p3)
Dim BL_p3p3L_p3_5a: BL_p3p3L_p3_5a=Array(p3LM_p3L_p3_5a_p3)
Dim BL_p3p3L_p3_5b: BL_p3p3L_p3_5b=Array(p3LM_p3L_p3_5b_p3)
Dim BL_p3p3L_p3_5c: BL_p3p3L_p3_5c=Array(p3LM_p3L_p3_5c_p3)
Dim BL_p3p3L_p3_5d: BL_p3p3L_p3_5d=Array(p3LM_p3L_p3_5d_p3)
Dim BL_p3p3L_p3_5e: BL_p3p3L_p3_5e=Array(p3LM_p3L_p3_5e_p3)
Dim BL_p3p3L_p3_5f: BL_p3p3L_p3_5f=Array(p3LM_p3L_p3_5f_p3)
Dim BL_p3p3L_p3_5g: BL_p3p3L_p3_5g=Array(p3LM_p3L_p3_5g_p3)
Dim BL_p3p3L_p3_6a: BL_p3p3L_p3_6a=Array(p3LM_p3L_p3_6a_p3)
Dim BL_p3p3L_p3_6b: BL_p3p3L_p3_6b=Array(p3LM_p3L_p3_6b_p3)
Dim BL_p3p3L_p3_6c: BL_p3p3L_p3_6c=Array(p3LM_p3L_p3_6c_p3)
Dim BL_p3p3L_p3_6d: BL_p3p3L_p3_6d=Array(p3LM_p3L_p3_6d_p3)
Dim BL_p3p3L_p3_6e: BL_p3p3L_p3_6e=Array(p3LM_p3L_p3_6e_p3)
Dim BL_p3p3L_p3_6f: BL_p3p3L_p3_6f=Array(p3LM_p3L_p3_6f_p3)
Dim BL_p3p3L_p3_6g: BL_p3p3L_p3_6g=Array(p3LM_p3L_p3_6g_p3)
Dim BL_p3p3L_p3_7a: BL_p3p3L_p3_7a=Array(p3LM_p3L_p3_7a_p3)
Dim BL_p3p3L_p3_7b: BL_p3p3L_p3_7b=Array(p3LM_p3L_p3_7b_p3)
Dim BL_p3p3L_p3_7c: BL_p3p3L_p3_7c=Array(p3LM_p3L_p3_7c_p3)
Dim BL_p3p3L_p3_7d: BL_p3p3L_p3_7d=Array(p3LM_p3L_p3_7d_p3)
Dim BL_p3p3L_p3_7e: BL_p3p3L_p3_7e=Array(p3LM_p3L_p3_7e_p3)
Dim BL_p3p3L_p3_7f: BL_p3p3L_p3_7f=Array(p3LM_p3L_p3_7f_p3)
Dim BL_p3p3L_p3_7g: BL_p3p3L_p3_7g=Array(p3LM_p3L_p3_7g_p3)
' Global arrays
Dim BGp3_Bakemap: BGp3_Bakemap=Array(p3BM_p3)
Dim BGp3_Lightmap: BGp3_Lightmap=Array(p3LM_p3L_p3_1a_p3, p3LM_p3L_p3_1b_p3, p3LM_p3L_p3_1c_p3, p3LM_p3L_p3_1d_p3, p3LM_p3L_p3_1e_p3, p3LM_p3L_p3_1f_p3, p3LM_p3L_p3_1g_p3, p3LM_p3L_p3_2a_p3, p3LM_p3L_p3_2b_p3, p3LM_p3L_p3_2c_p3, p3LM_p3L_p3_2d_p3, p3LM_p3L_p3_2e_p3, p3LM_p3L_p3_2f_p3, p3LM_p3L_p3_2g_p3, p3LM_p3L_p3_3a_p3, p3LM_p3L_p3_3b_p3, p3LM_p3L_p3_3c_p3, p3LM_p3L_p3_3d_p3, p3LM_p3L_p3_3e_p3, p3LM_p3L_p3_3f_p3, p3LM_p3L_p3_3g_p3, p3LM_p3L_p3_4a_p3, p3LM_p3L_p3_4b_p3, p3LM_p3L_p3_4c_p3, p3LM_p3L_p3_4d_p3, p3LM_p3L_p3_4e_p3, p3LM_p3L_p3_4f_p3, p3LM_p3L_p3_4g_p3, p3LM_p3L_p3_5a_p3, p3LM_p3L_p3_5b_p3, p3LM_p3L_p3_5c_p3, p3LM_p3L_p3_5d_p3, p3LM_p3L_p3_5e_p3, p3LM_p3L_p3_5f_p3, p3LM_p3L_p3_5g_p3, p3LM_p3L_p3_6a_p3, p3LM_p3L_p3_6b_p3, p3LM_p3L_p3_6c_p3, p3LM_p3L_p3_6d_p3, p3LM_p3L_p3_6e_p3, p3LM_p3L_p3_6f_p3, p3LM_p3L_p3_6g_p3, p3LM_p3L_p3_7a_p3, p3LM_p3L_p3_7b_p3, p3LM_p3L_p3_7c_p3, p3LM_p3L_p3_7d_p3, p3LM_p3L_p3_7e_p3, p3LM_p3L_p3_7f_p3, p3LM_p3L_p3_7g_p3)
Dim BGp3_All: BGp3_All=Array(p3BM_p3, p3LM_p3L_p3_1a_p3, p3LM_p3L_p3_1b_p3, p3LM_p3L_p3_1c_p3, p3LM_p3L_p3_1d_p3, p3LM_p3L_p3_1e_p3, p3LM_p3L_p3_1f_p3, p3LM_p3L_p3_1g_p3, p3LM_p3L_p3_2a_p3, p3LM_p3L_p3_2b_p3, p3LM_p3L_p3_2c_p3, p3LM_p3L_p3_2d_p3, p3LM_p3L_p3_2e_p3, p3LM_p3L_p3_2f_p3, p3LM_p3L_p3_2g_p3, p3LM_p3L_p3_3a_p3, p3LM_p3L_p3_3b_p3, p3LM_p3L_p3_3c_p3, p3LM_p3L_p3_3d_p3, p3LM_p3L_p3_3e_p3, p3LM_p3L_p3_3f_p3, p3LM_p3L_p3_3g_p3, p3LM_p3L_p3_4a_p3, p3LM_p3L_p3_4b_p3, p3LM_p3L_p3_4c_p3, p3LM_p3L_p3_4d_p3, p3LM_p3L_p3_4e_p3, p3LM_p3L_p3_4f_p3, p3LM_p3L_p3_4g_p3, p3LM_p3L_p3_5a_p3, p3LM_p3L_p3_5b_p3, p3LM_p3L_p3_5c_p3, p3LM_p3L_p3_5d_p3, p3LM_p3L_p3_5e_p3, p3LM_p3L_p3_5f_p3, p3LM_p3L_p3_5g_p3, p3LM_p3L_p3_6a_p3, p3LM_p3L_p3_6b_p3, p3LM_p3L_p3_6c_p3, p3LM_p3L_p3_6d_p3, p3LM_p3L_p3_6e_p3, p3LM_p3L_p3_6f_p3, p3LM_p3L_p3_6g_p3, p3LM_p3L_p3_7a_p3, p3LM_p3L_p3_7b_p3, p3LM_p3L_p3_7c_p3, p3LM_p3L_p3_7d_p3, p3LM_p3L_p3_7e_p3, p3LM_p3L_p3_7f_p3, p3LM_p3L_p3_7g_p3)
' VLM p3 Arrays - End

' VLM p4 Arrays - Start
' Arrays per baked part
Dim BP_p4p4: BP_p4p4=Array(p4BM_p4, p4LM_p4L_p4_1a_p4, p4LM_p4L_p4_1b_p4, p4LM_p4L_p4_1c_p4, p4LM_p4L_p4_1d_p4, p4LM_p4L_p4_1e_p4, p4LM_p4L_p4_1f_p4, p4LM_p4L_p4_1g_p4, p4LM_p4L_p4_2a_p4, p4LM_p4L_p4_2b_p4, p4LM_p4L_p4_2c_p4, p4LM_p4L_p4_2d_p4, p4LM_p4L_p4_2e_p4, p4LM_p4L_p4_2f_p4, p4LM_p4L_p4_2g_p4, p4LM_p4L_p4_3a_p4, p4LM_p4L_p4_3b_p4, p4LM_p4L_p4_3c_p4, p4LM_p4L_p4_3d_p4, p4LM_p4L_p4_3e_p4, p4LM_p4L_p4_3f_p4, p4LM_p4L_p4_3g_p4, p4LM_p4L_p4_4a_p4, p4LM_p4L_p4_4b_p4, p4LM_p4L_p4_4c_p4, p4LM_p4L_p4_4d_p4, p4LM_p4L_p4_4e_p4, p4LM_p4L_p4_4f_p4, p4LM_p4L_p4_4g_p4, p4LM_p4L_p4_5a_p4, p4LM_p4L_p4_5b_p4, p4LM_p4L_p4_5c_p4, p4LM_p4L_p4_5d_p4, p4LM_p4L_p4_5e_p4, p4LM_p4L_p4_5f_p4, p4LM_p4L_p4_5g_p4, p4LM_p4L_p4_6a_p4, p4LM_p4L_p4_6b_p4, p4LM_p4L_p4_6c_p4, p4LM_p4L_p4_6d_p4, p4LM_p4L_p4_6e_p4, p4LM_p4L_p4_6f_p4, p4LM_p4L_p4_6g_p4, p4LM_p4L_p4_7a_p4, p4LM_p4L_p4_7b_p4, p4LM_p4L_p4_7c_p4, p4LM_p4L_p4_7d_p4, p4LM_p4L_p4_7e_p4, p4LM_p4L_p4_7f_p4, p4LM_p4L_p4_7g_p4)
' Arrays per lighting scenario
Dim BL_p4Room: BL_p4Room=Array(p4BM_p4)
Dim BL_p4p4L_p4_1a: BL_p4p4L_p4_1a=Array(p4LM_p4L_p4_1a_p4)
Dim BL_p4p4L_p4_1b: BL_p4p4L_p4_1b=Array(p4LM_p4L_p4_1b_p4)
Dim BL_p4p4L_p4_1c: BL_p4p4L_p4_1c=Array(p4LM_p4L_p4_1c_p4)
Dim BL_p4p4L_p4_1d: BL_p4p4L_p4_1d=Array(p4LM_p4L_p4_1d_p4)
Dim BL_p4p4L_p4_1e: BL_p4p4L_p4_1e=Array(p4LM_p4L_p4_1e_p4)
Dim BL_p4p4L_p4_1f: BL_p4p4L_p4_1f=Array(p4LM_p4L_p4_1f_p4)
Dim BL_p4p4L_p4_1g: BL_p4p4L_p4_1g=Array(p4LM_p4L_p4_1g_p4)
Dim BL_p4p4L_p4_2a: BL_p4p4L_p4_2a=Array(p4LM_p4L_p4_2a_p4)
Dim BL_p4p4L_p4_2b: BL_p4p4L_p4_2b=Array(p4LM_p4L_p4_2b_p4)
Dim BL_p4p4L_p4_2c: BL_p4p4L_p4_2c=Array(p4LM_p4L_p4_2c_p4)
Dim BL_p4p4L_p4_2d: BL_p4p4L_p4_2d=Array(p4LM_p4L_p4_2d_p4)
Dim BL_p4p4L_p4_2e: BL_p4p4L_p4_2e=Array(p4LM_p4L_p4_2e_p4)
Dim BL_p4p4L_p4_2f: BL_p4p4L_p4_2f=Array(p4LM_p4L_p4_2f_p4)
Dim BL_p4p4L_p4_2g: BL_p4p4L_p4_2g=Array(p4LM_p4L_p4_2g_p4)
Dim BL_p4p4L_p4_3a: BL_p4p4L_p4_3a=Array(p4LM_p4L_p4_3a_p4)
Dim BL_p4p4L_p4_3b: BL_p4p4L_p4_3b=Array(p4LM_p4L_p4_3b_p4)
Dim BL_p4p4L_p4_3c: BL_p4p4L_p4_3c=Array(p4LM_p4L_p4_3c_p4)
Dim BL_p4p4L_p4_3d: BL_p4p4L_p4_3d=Array(p4LM_p4L_p4_3d_p4)
Dim BL_p4p4L_p4_3e: BL_p4p4L_p4_3e=Array(p4LM_p4L_p4_3e_p4)
Dim BL_p4p4L_p4_3f: BL_p4p4L_p4_3f=Array(p4LM_p4L_p4_3f_p4)
Dim BL_p4p4L_p4_3g: BL_p4p4L_p4_3g=Array(p4LM_p4L_p4_3g_p4)
Dim BL_p4p4L_p4_4a: BL_p4p4L_p4_4a=Array(p4LM_p4L_p4_4a_p4)
Dim BL_p4p4L_p4_4b: BL_p4p4L_p4_4b=Array(p4LM_p4L_p4_4b_p4)
Dim BL_p4p4L_p4_4c: BL_p4p4L_p4_4c=Array(p4LM_p4L_p4_4c_p4)
Dim BL_p4p4L_p4_4d: BL_p4p4L_p4_4d=Array(p4LM_p4L_p4_4d_p4)
Dim BL_p4p4L_p4_4e: BL_p4p4L_p4_4e=Array(p4LM_p4L_p4_4e_p4)
Dim BL_p4p4L_p4_4f: BL_p4p4L_p4_4f=Array(p4LM_p4L_p4_4f_p4)
Dim BL_p4p4L_p4_4g: BL_p4p4L_p4_4g=Array(p4LM_p4L_p4_4g_p4)
Dim BL_p4p4L_p4_5a: BL_p4p4L_p4_5a=Array(p4LM_p4L_p4_5a_p4)
Dim BL_p4p4L_p4_5b: BL_p4p4L_p4_5b=Array(p4LM_p4L_p4_5b_p4)
Dim BL_p4p4L_p4_5c: BL_p4p4L_p4_5c=Array(p4LM_p4L_p4_5c_p4)
Dim BL_p4p4L_p4_5d: BL_p4p4L_p4_5d=Array(p4LM_p4L_p4_5d_p4)
Dim BL_p4p4L_p4_5e: BL_p4p4L_p4_5e=Array(p4LM_p4L_p4_5e_p4)
Dim BL_p4p4L_p4_5f: BL_p4p4L_p4_5f=Array(p4LM_p4L_p4_5f_p4)
Dim BL_p4p4L_p4_5g: BL_p4p4L_p4_5g=Array(p4LM_p4L_p4_5g_p4)
Dim BL_p4p4L_p4_6a: BL_p4p4L_p4_6a=Array(p4LM_p4L_p4_6a_p4)
Dim BL_p4p4L_p4_6b: BL_p4p4L_p4_6b=Array(p4LM_p4L_p4_6b_p4)
Dim BL_p4p4L_p4_6c: BL_p4p4L_p4_6c=Array(p4LM_p4L_p4_6c_p4)
Dim BL_p4p4L_p4_6d: BL_p4p4L_p4_6d=Array(p4LM_p4L_p4_6d_p4)
Dim BL_p4p4L_p4_6e: BL_p4p4L_p4_6e=Array(p4LM_p4L_p4_6e_p4)
Dim BL_p4p4L_p4_6f: BL_p4p4L_p4_6f=Array(p4LM_p4L_p4_6f_p4)
Dim BL_p4p4L_p4_6g: BL_p4p4L_p4_6g=Array(p4LM_p4L_p4_6g_p4)
Dim BL_p4p4L_p4_7a: BL_p4p4L_p4_7a=Array(p4LM_p4L_p4_7a_p4)
Dim BL_p4p4L_p4_7b: BL_p4p4L_p4_7b=Array(p4LM_p4L_p4_7b_p4)
Dim BL_p4p4L_p4_7c: BL_p4p4L_p4_7c=Array(p4LM_p4L_p4_7c_p4)
Dim BL_p4p4L_p4_7d: BL_p4p4L_p4_7d=Array(p4LM_p4L_p4_7d_p4)
Dim BL_p4p4L_p4_7e: BL_p4p4L_p4_7e=Array(p4LM_p4L_p4_7e_p4)
Dim BL_p4p4L_p4_7f: BL_p4p4L_p4_7f=Array(p4LM_p4L_p4_7f_p4)
Dim BL_p4p4L_p4_7g: BL_p4p4L_p4_7g=Array(p4LM_p4L_p4_7g_p4)
' Global arrays
Dim BGp4_Bakemap: BGp4_Bakemap=Array(p4BM_p4)
Dim BGp4_Lightmap: BGp4_Lightmap=Array(p4LM_p4L_p4_1a_p4, p4LM_p4L_p4_1b_p4, p4LM_p4L_p4_1c_p4, p4LM_p4L_p4_1d_p4, p4LM_p4L_p4_1e_p4, p4LM_p4L_p4_1f_p4, p4LM_p4L_p4_1g_p4, p4LM_p4L_p4_2a_p4, p4LM_p4L_p4_2b_p4, p4LM_p4L_p4_2c_p4, p4LM_p4L_p4_2d_p4, p4LM_p4L_p4_2e_p4, p4LM_p4L_p4_2f_p4, p4LM_p4L_p4_2g_p4, p4LM_p4L_p4_3a_p4, p4LM_p4L_p4_3b_p4, p4LM_p4L_p4_3c_p4, p4LM_p4L_p4_3d_p4, p4LM_p4L_p4_3e_p4, p4LM_p4L_p4_3f_p4, p4LM_p4L_p4_3g_p4, p4LM_p4L_p4_4a_p4, p4LM_p4L_p4_4b_p4, p4LM_p4L_p4_4c_p4, p4LM_p4L_p4_4d_p4, p4LM_p4L_p4_4e_p4, p4LM_p4L_p4_4f_p4, p4LM_p4L_p4_4g_p4, p4LM_p4L_p4_5a_p4, p4LM_p4L_p4_5b_p4, p4LM_p4L_p4_5c_p4, p4LM_p4L_p4_5d_p4, p4LM_p4L_p4_5e_p4, p4LM_p4L_p4_5f_p4, p4LM_p4L_p4_5g_p4, p4LM_p4L_p4_6a_p4, p4LM_p4L_p4_6b_p4, p4LM_p4L_p4_6c_p4, p4LM_p4L_p4_6d_p4, p4LM_p4L_p4_6e_p4, p4LM_p4L_p4_6f_p4, p4LM_p4L_p4_6g_p4, p4LM_p4L_p4_7a_p4, p4LM_p4L_p4_7b_p4, p4LM_p4L_p4_7c_p4, p4LM_p4L_p4_7d_p4, p4LM_p4L_p4_7e_p4, p4LM_p4L_p4_7f_p4, p4LM_p4L_p4_7g_p4)
Dim BGp4_All: BGp4_All=Array(p4BM_p4, p4LM_p4L_p4_1a_p4, p4LM_p4L_p4_1b_p4, p4LM_p4L_p4_1c_p4, p4LM_p4L_p4_1d_p4, p4LM_p4L_p4_1e_p4, p4LM_p4L_p4_1f_p4, p4LM_p4L_p4_1g_p4, p4LM_p4L_p4_2a_p4, p4LM_p4L_p4_2b_p4, p4LM_p4L_p4_2c_p4, p4LM_p4L_p4_2d_p4, p4LM_p4L_p4_2e_p4, p4LM_p4L_p4_2f_p4, p4LM_p4L_p4_2g_p4, p4LM_p4L_p4_3a_p4, p4LM_p4L_p4_3b_p4, p4LM_p4L_p4_3c_p4, p4LM_p4L_p4_3d_p4, p4LM_p4L_p4_3e_p4, p4LM_p4L_p4_3f_p4, p4LM_p4L_p4_3g_p4, p4LM_p4L_p4_4a_p4, p4LM_p4L_p4_4b_p4, p4LM_p4L_p4_4c_p4, p4LM_p4L_p4_4d_p4, p4LM_p4L_p4_4e_p4, p4LM_p4L_p4_4f_p4, p4LM_p4L_p4_4g_p4, p4LM_p4L_p4_5a_p4, p4LM_p4L_p4_5b_p4, p4LM_p4L_p4_5c_p4, p4LM_p4L_p4_5d_p4, p4LM_p4L_p4_5e_p4, p4LM_p4L_p4_5f_p4, p4LM_p4L_p4_5g_p4, p4LM_p4L_p4_6a_p4, p4LM_p4L_p4_6b_p4, p4LM_p4L_p4_6c_p4, p4LM_p4L_p4_6d_p4, p4LM_p4L_p4_6e_p4, p4LM_p4L_p4_6f_p4, p4LM_p4L_p4_6g_p4, p4LM_p4L_p4_7a_p4, p4LM_p4L_p4_7b_p4, p4LM_p4L_p4_7c_p4, p4LM_p4L_p4_7d_p4, p4LM_p4L_p4_7e_p4, p4LM_p4L_p4_7f_p4, p4LM_p4L_p4_7g_p4)
' VLM p4 Arrays - End

' VLM crmb Arrays - Start
' Arrays per baked part
Dim BP_crmbcrmb: BP_crmbcrmb=Array(crmbBM_crmb, crmbLM_crmbL_cr_1a_crmb, crmbLM_crmbL_cr_1b_crmb, crmbLM_crmbL_cr_1c_crmb, crmbLM_crmbL_cr_1d_crmb, crmbLM_crmbL_cr_1e_crmb, crmbLM_crmbL_cr_1f_crmb, crmbLM_crmbL_cr_1g_crmb, crmbLM_crmbL_cr_2a_crmb, crmbLM_crmbL_cr_2b_crmb, crmbLM_crmbL_cr_2c_crmb, crmbLM_crmbL_cr_2d_crmb, crmbLM_crmbL_cr_2e_crmb, crmbLM_crmbL_cr_2f_crmb, crmbLM_crmbL_cr_2g_crmb, crmbLM_crmbL_mb_1a_crmb, crmbLM_crmbL_mb_1b_crmb, crmbLM_crmbL_mb_1c_crmb, crmbLM_crmbL_mb_1d_crmb, crmbLM_crmbL_mb_1e_crmb, crmbLM_crmbL_mb_1f_crmb, crmbLM_crmbL_mb_1g_crmb, crmbLM_crmbL_mb_2a_crmb, crmbLM_crmbL_mb_2b_crmb, crmbLM_crmbL_mb_2c_crmb, crmbLM_crmbL_mb_2d_crmb, crmbLM_crmbL_mb_2e_crmb, crmbLM_crmbL_mb_2f_crmb, crmbLM_crmbL_mb_2g_crmb)
' Arrays per lighting scenario
Dim BL_crmbRoom: BL_crmbRoom=Array(crmbBM_crmb)
Dim BL_crmbcrmbL_cr_1a: BL_crmbcrmbL_cr_1a=Array(crmbLM_crmbL_cr_1a_crmb)
Dim BL_crmbcrmbL_cr_1b: BL_crmbcrmbL_cr_1b=Array(crmbLM_crmbL_cr_1b_crmb)
Dim BL_crmbcrmbL_cr_1c: BL_crmbcrmbL_cr_1c=Array(crmbLM_crmbL_cr_1c_crmb)
Dim BL_crmbcrmbL_cr_1d: BL_crmbcrmbL_cr_1d=Array(crmbLM_crmbL_cr_1d_crmb)
Dim BL_crmbcrmbL_cr_1e: BL_crmbcrmbL_cr_1e=Array(crmbLM_crmbL_cr_1e_crmb)
Dim BL_crmbcrmbL_cr_1f: BL_crmbcrmbL_cr_1f=Array(crmbLM_crmbL_cr_1f_crmb)
Dim BL_crmbcrmbL_cr_1g: BL_crmbcrmbL_cr_1g=Array(crmbLM_crmbL_cr_1g_crmb)
Dim BL_crmbcrmbL_cr_2a: BL_crmbcrmbL_cr_2a=Array(crmbLM_crmbL_cr_2a_crmb)
Dim BL_crmbcrmbL_cr_2b: BL_crmbcrmbL_cr_2b=Array(crmbLM_crmbL_cr_2b_crmb)
Dim BL_crmbcrmbL_cr_2c: BL_crmbcrmbL_cr_2c=Array(crmbLM_crmbL_cr_2c_crmb)
Dim BL_crmbcrmbL_cr_2d: BL_crmbcrmbL_cr_2d=Array(crmbLM_crmbL_cr_2d_crmb)
Dim BL_crmbcrmbL_cr_2e: BL_crmbcrmbL_cr_2e=Array(crmbLM_crmbL_cr_2e_crmb)
Dim BL_crmbcrmbL_cr_2f: BL_crmbcrmbL_cr_2f=Array(crmbLM_crmbL_cr_2f_crmb)
Dim BL_crmbcrmbL_cr_2g: BL_crmbcrmbL_cr_2g=Array(crmbLM_crmbL_cr_2g_crmb)
Dim BL_crmbcrmbL_mb_1a: BL_crmbcrmbL_mb_1a=Array(crmbLM_crmbL_mb_1a_crmb)
Dim BL_crmbcrmbL_mb_1b: BL_crmbcrmbL_mb_1b=Array(crmbLM_crmbL_mb_1b_crmb)
Dim BL_crmbcrmbL_mb_1c: BL_crmbcrmbL_mb_1c=Array(crmbLM_crmbL_mb_1c_crmb)
Dim BL_crmbcrmbL_mb_1d: BL_crmbcrmbL_mb_1d=Array(crmbLM_crmbL_mb_1d_crmb)
Dim BL_crmbcrmbL_mb_1e: BL_crmbcrmbL_mb_1e=Array(crmbLM_crmbL_mb_1e_crmb)
Dim BL_crmbcrmbL_mb_1f: BL_crmbcrmbL_mb_1f=Array(crmbLM_crmbL_mb_1f_crmb)
Dim BL_crmbcrmbL_mb_1g: BL_crmbcrmbL_mb_1g=Array(crmbLM_crmbL_mb_1g_crmb)
Dim BL_crmbcrmbL_mb_2a: BL_crmbcrmbL_mb_2a=Array(crmbLM_crmbL_mb_2a_crmb)
Dim BL_crmbcrmbL_mb_2b: BL_crmbcrmbL_mb_2b=Array(crmbLM_crmbL_mb_2b_crmb)
Dim BL_crmbcrmbL_mb_2c: BL_crmbcrmbL_mb_2c=Array(crmbLM_crmbL_mb_2c_crmb)
Dim BL_crmbcrmbL_mb_2d: BL_crmbcrmbL_mb_2d=Array(crmbLM_crmbL_mb_2d_crmb)
Dim BL_crmbcrmbL_mb_2e: BL_crmbcrmbL_mb_2e=Array(crmbLM_crmbL_mb_2e_crmb)
Dim BL_crmbcrmbL_mb_2f: BL_crmbcrmbL_mb_2f=Array(crmbLM_crmbL_mb_2f_crmb)
Dim BL_crmbcrmbL_mb_2g: BL_crmbcrmbL_mb_2g=Array(crmbLM_crmbL_mb_2g_crmb)
' Global arrays
Dim BGcrmb_Bakemap: BGcrmb_Bakemap=Array(crmbBM_crmb)
Dim BGcrmb_Lightmap: BGcrmb_Lightmap=Array(crmbLM_crmbL_cr_1a_crmb, crmbLM_crmbL_cr_1b_crmb, crmbLM_crmbL_cr_1c_crmb, crmbLM_crmbL_cr_1d_crmb, crmbLM_crmbL_cr_1e_crmb, crmbLM_crmbL_cr_1f_crmb, crmbLM_crmbL_cr_1g_crmb, crmbLM_crmbL_cr_2a_crmb, crmbLM_crmbL_cr_2b_crmb, crmbLM_crmbL_cr_2c_crmb, crmbLM_crmbL_cr_2d_crmb, crmbLM_crmbL_cr_2e_crmb, crmbLM_crmbL_cr_2f_crmb, crmbLM_crmbL_cr_2g_crmb, crmbLM_crmbL_mb_1a_crmb, crmbLM_crmbL_mb_1b_crmb, crmbLM_crmbL_mb_1c_crmb, crmbLM_crmbL_mb_1d_crmb, crmbLM_crmbL_mb_1e_crmb, crmbLM_crmbL_mb_1f_crmb, crmbLM_crmbL_mb_1g_crmb, crmbLM_crmbL_mb_2a_crmb, crmbLM_crmbL_mb_2b_crmb, crmbLM_crmbL_mb_2c_crmb, crmbLM_crmbL_mb_2d_crmb, crmbLM_crmbL_mb_2e_crmb, crmbLM_crmbL_mb_2f_crmb, crmbLM_crmbL_mb_2g_crmb)
Dim BGcrmb_All: BGcrmb_All=Array(crmbBM_crmb, crmbLM_crmbL_cr_1a_crmb, crmbLM_crmbL_cr_1b_crmb, crmbLM_crmbL_cr_1c_crmb, crmbLM_crmbL_cr_1d_crmb, crmbLM_crmbL_cr_1e_crmb, crmbLM_crmbL_cr_1f_crmb, crmbLM_crmbL_cr_1g_crmb, crmbLM_crmbL_cr_2a_crmb, crmbLM_crmbL_cr_2b_crmb, crmbLM_crmbL_cr_2c_crmb, crmbLM_crmbL_cr_2d_crmb, crmbLM_crmbL_cr_2e_crmb, crmbLM_crmbL_cr_2f_crmb, crmbLM_crmbL_cr_2g_crmb, crmbLM_crmbL_mb_1a_crmb, crmbLM_crmbL_mb_1b_crmb, crmbLM_crmbL_mb_1c_crmb, crmbLM_crmbL_mb_1d_crmb, crmbLM_crmbL_mb_1e_crmb, crmbLM_crmbL_mb_1f_crmb, crmbLM_crmbL_mb_1g_crmb, crmbLM_crmbL_mb_2a_crmb, crmbLM_crmbL_mb_2b_crmb, crmbLM_crmbL_mb_2c_crmb, crmbLM_crmbL_mb_2d_crmb, crmbLM_crmbL_mb_2e_crmb, crmbLM_crmbL_mb_2f_crmb, crmbLM_crmbL_mb_2g_crmb)
' VLM crmb Arrays - End

' VLM BG Arrays - Start
' Arrays per baked part
Dim BP_BGBG: BP_BGBG=Array(BGBM_BG, BGLM_GIS_gi27_BG, BGLM_GIS_gi25_BG, BGLM_GIS_gi30_BG, BGLM_GIS_gi28_BG, BGLM_GIS_gi26_BG, BGLM_GIS_gi29_BG, BGLM_GI_BG)
' Arrays per lighting scenario
Dim BL_BGGI: BL_BGGI=Array(BGLM_GI_BG)
Dim BL_BGGIS_gi25: BL_BGGIS_gi25=Array(BGLM_GIS_gi25_BG)
Dim BL_BGGIS_gi26: BL_BGGIS_gi26=Array(BGLM_GIS_gi26_BG)
Dim BL_BGGIS_gi27: BL_BGGIS_gi27=Array(BGLM_GIS_gi27_BG)
Dim BL_BGGIS_gi28: BL_BGGIS_gi28=Array(BGLM_GIS_gi28_BG)
Dim BL_BGGIS_gi29: BL_BGGIS_gi29=Array(BGLM_GIS_gi29_BG)
Dim BL_BGGIS_gi30: BL_BGGIS_gi30=Array(BGLM_GIS_gi30_BG)
Dim BL_BGRoom: BL_BGRoom=Array(BGBM_BG)
' Global arrays
Dim BGBG_Bakemap: BGBG_Bakemap=Array(BGBM_BG)
Dim BGBG_Lightmap: BGBG_Lightmap=Array(BGLM_GI_BG, BGLM_GIS_gi25_BG, BGLM_GIS_gi26_BG, BGLM_GIS_gi27_BG, BGLM_GIS_gi28_BG, BGLM_GIS_gi29_BG, BGLM_GIS_gi30_BG)
Dim BGBG_All: BGBG_All=Array(BGBM_BG, BGLM_GI_BG, BGLM_GIS_gi25_BG, BGLM_GIS_gi26_BG, BGLM_GIS_gi27_BG, BGLM_GIS_gi28_BG, BGLM_GIS_gi29_BG, BGLM_GIS_gi30_BG)
' VLM BG Arrays - End

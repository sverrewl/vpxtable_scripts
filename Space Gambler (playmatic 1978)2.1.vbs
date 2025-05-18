
'***************************************************************
'               Space Gambler
' Space Gambler / IPD No. 2250 / March, 1978 / 4 Players
' Playmatic, Solid State Electronic (SS)
' VPX Author(s): NestorGian, Akiles, Pedator (4/2021)
'   VP9 Author(s): TAB, Destruk
' dips added by Inkochnito
'***************************************************************

'   ZRDT: Drop Targets
' ZMAT: General Math Functions
'   ZBRL: Ball Rolling and Drop Sounds
' ZPHY: General Advice on Physics


Option Explicit
Randomize

On Error Resume Next
ExecuteGlobal GetTextFile("controller.vbs")
If Err Then MsgBox "You need the Controller.vbs file in order to run this table (installed with the VPX package in the scripts folder)"
On Error Goto 0


'*******************************************
'  ZOPT: User Options
'*******************************************

Dim LightLevel : LightLevel = 0.6       ' Level of room lighting (0 to 1), where 0 is dark and 100 is brightest
'Dim ColorLUT : ColorLUT = 1            ' Color desaturation LUTs: 1 to 11, where 1 is normal and 11 is black'n'white
Dim VolumeDial : VolumeDial = 0.8             ' Overall Mechanical sound effect volume. Recommended values should be no greater than 1.
Dim BallRollVolume : BallRollVolume = 0.5     ' Level of ball rolling volume. Value between 0 and 1
Dim RampRollVolume : RampRollVolume = 0.5     ' Level of ramp rolling volume. Value between 0 and 1
Dim StagedFlippers : StagedFlippers = 0         ' Staged Flippers. 0 = Disabled, 1 = Enabled
Dim VRRoom
Dim VRRoomChoice: VRRoomChoice = 1              ' 1 = Space room  2 = Ultra Minimal room


Sub Table1_OptionEvent(ByVal eventId)
    If eventId = 1 Then DisableStaticPreRendering = True
  Dim BP, v, s
    ' Sound volumes
    VolumeDial = Table1.Option("Mech Volume", 0, 1, 0.01, 0.8, 1)
    BallRollVolume = Table1.Option("Ball Roll Volume", 0, 1, 0.01, 0.5, 1)

    ' Room Brightness
    LightLevel = NightDay/100
    SetRoomBrightness LightLevel

  ' Saturation
  s = Table1.Option("Adjust Saturation", 0, 3, 1, 1, 0, Array("+10%", "non","-10%","-20%"))
  SetSaturation(s)

  ' VR Room
  VRRoomChoice = Table1.Option("VR Room", 1, 2, 1, 1, 0, Array("SPACE ROOM", "ULTRA MINIMAL"))
  If RenderingMode=2 OR ForceVR=True Then VRRoom = VRRoomChoice: Else VRRoom = 0: End If
  SetupVRRoom

  ' Rails
  If RenderingMode <>2 Then
    v = Table1.Option("Cab Rails", 0, 1, 1, 1, 0, Array("Not Visible", "Visible"))
    For Each BP in BP_PincabRails : BP.Visible = v: Next
  End If

    If eventId = 3 Then DisableStaticPreRendering = False
End Sub



' Update these arrays if you want to change more materials with room light level
Dim RoomBrightnessMtlArray: RoomBrightnessMtlArray = Array("VLM.Bake.Active","VLM.Bake.Solid","VLM.Bake.Metal","Plastic with an image")


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


Sub SetSaturation(s)
if s = 0 Then Table1.ColorGradeImage = "colorgradelut256x16+10"
  if s = 1 Then Table1.ColorGradeImage = ""
  if s = 2 Then Table1.ColorGradeImage = "colorgradelut256x16-10"
  if s = 3 Then Table1.ColorGradeImage = "colorgradelut256x16-20"
End Sub



'******************************************************
'  ZVAR: Constants and Global Variables
'******************************************************

Const BallSize = 50
Const BallMass = 1
Const tnob = 1
Const lob = 0

Dim tablewidth: tablewidth = Table1.width
Dim tableheight: tableheight = Table1.height


'LoadVPM "01500000", "play1.vbs", 3.1
LoadVPM "01120100", "play1.vbs", 3.1

Const cGameName="spcgambl",cCredits="Space Gambler"
Const UseSolenoids = 2
Const UseLamps = 0
Const UseSync = 1
Const HandleMech = 0
Const SSolenoidOn = "fx_solenoidon"
Const SSolenoidOff = "fx_solenoidoff"
Const SCoin = "fx_Coin"

Const UsingROM  = True        'The UsingROM flag is to indicate code that requires ROM usage.

Dim VRMode, DesktopMode: DesktopMode = Table1.ShowDT
Const ForceVR = False  'Forces VR room to be active

Dim LightAro

Dim VarHidden:VarHidden = Table1.ShowDT
if B2SOn = true then VarHidden = 1

Dim x

' Remove desktop items in FS mode
If Table1.ShowDT then
    For each x in aReels
        x.Visible = 1
    Next
Else
    For each x in aReels
        x.Visible = 0
    Next
End If




'*********
'Solenoids
'*********
SolCallBack(4)="BumperLamp"
SolCallback(7)="SolRelease"   '"SolTrough"
SolCallback(8)="vpmSolSound SoundFX(""fx_Knocker"",DOFKnocker),"
SolCallBack(2)="AroLight"


'Sub SolTrough(Enabled)
' If Enabled Then
'     SolBank(Enabled)
'     bsTrough.ExitSol_On
'     vpmTimer.PulseSw 17
' End If
'End Sub

Sub SolBank(Enabled)
  If Enabled Then
    If sw44.IsDropped Or sw43.IsDropped Or sw42.IsDropped or sw41.IsDropped Then
      PlaySoundat SoundFX("fx_resetdrop",DOFContactors), sw44
      sw44.IsDropped=0
      sw43.IsDropped=0
      sw42.IsDropped=0
      sw41.IsDropped=0
      Controller.Switch(41)=0
      Controller.Switch(42)=0
      Controller.Switch(43)=0
      Controller.Switch(44)=0
      Controller.Switch(3)=0
      Controller.Switch(4)=0
    End If
  End If
End Sub

Sub BumperLamp(Enabled)
  If Enabled Then
    bumper1.force=13
    LightBumper001.State = ABS(Enabled)
    bumper2.force=13
    LightBumper002.State = ABS(Enabled)
     Else
    bumper1.force=0
    LightBumper001.State = 0
    bumper2.force=0
    LightBumper002.State = 0
     End If
End Sub

'*******************************************
'  ZINI : Table Initialization
'*******************************************

Dim bsTrough, bump1, bump2, gBOT, SBall1, Flipperactive

Sub Table1_Init
  vpmInit me
  With Controller
    .GameName=cGameName
    If Err Then MsgBox "Can't start Game"&cGameName&vbNewLine&Err.Description:Exit Sub
    .SplashInfoLine = "Space Gambler, Playmatic 1978" & vbnewline & "Table by NestorGian"
    .Games(cGameName).Settings.Value("rol") = 0   '1= rotated display, 0= normal
        .Games(cGameName).Settings.Value("sound") = 1 '1 enabled rom sound
        .HandleMechanics = 0
    .ShowDMDOnly=1
    .ShowFrame=0
    .ShowTitle=0
    .Hidden=VarHidden
    On Error Resume Next
        .SolMask(0) = 0
        vpmTimer.AddTimer 2000, "Controller.SolMask(0)=&Hffffffff'" 'ignore all solenoids - then add the Timer to renable all the solenoids after 2 seconds
        Controller.Run  '.Run GetPlayerHWnd
        If Err Then MsgBox Err.Description
        On Error Goto 0
    End With

  'Map all lamps to the corresponding ROM output using the value of TimerInterval of each light object
  vpmMapLights AllLamps     'Make a collection called "AllLamps" and put all the light objects in it.


  'Ball initializations need for physical trough
  Set SBall1 = sw46.CreateSizedballWithMass(Ballsize/2,Ballmass)
' gBOT = Array(SBall1)

  Controller.Switch(13) = 1


  ' Nudging
  vpmNudge.TiltSwitch=swTilt
  vpmNudge.Sensitivity=5
  vpmNudge.TiltObj = Array(bumper1, bumper2, LeftSlingshot, RightSlingshot)

  ' Trough
' Set bsTrough=New cvpmBallStack
' With bsTrough
'   .InitSw 0,13,0,0,0,0,0,0
'   .InitKick BallRelease,90,8
'   .InitEntrySnd "fx_Solenoid", "fx_Solenoid"
'        .InitExitSnd SoundFX("fx_ballrel", DOFContactors), SoundFX("fx_Solenoid", DOFContactors)
'   .Balls=1
' End With

  PinMAMETimer.Interval = PinMAMEInterval
    PinMAMETimer.Enabled = 1


 '  vpmMapLights aLights
  RealTimeUpdates.Enabled = 1   ' turn the RealTimeUpdates timer
' BumperLamp 1
  Flipperactive = False

  'Initialize slings
  RStep = 0:RightSlingShot.Timerenabled=True
  LStep = 0:LeftSlingShot.Timerenabled=True
  SetupVRRoom
  SetupVRBackglass
  SetBackglass
  vpmtimer.addtimer 3000, "GiOn '"

End Sub

Sub table1_Paused:Controller.Pause = 1:End Sub
Sub table1_unPaused:Controller.Pause = 0:End Sub
Sub table1_exit:Controller.stop:End Sub


'******************
' Timers
'******************

Dim OldState1

'Sub RealTimeUpdates_Timer

' UpdateLeds
' UpdateRedDice

' If Controller.Lamp(47) Then 'tilted deactivate the table objects
'        Flipperactive = False
'        LeftFlipper.RotateToStart
'        RightFlipper.RotateToStart
'        GiOff
'        Bumper1.Force = 0
'        Bumper2.Force = 0
'        LeftSlingshot.Disabled = 1
'        RightSlingshot.Disabled = 1
'   BumperLamp 0
'   OldState1 = 0
'    Else
'        OldState1 = 1
'    End If

' If Controller.Lamp(1) <> OldState1 Then 'Game Over light
'        If Controller.Lamp(1)Then
'            VpmNudge.SolGameOn False
'            OldState1 = True
'            GiOff
'            Flipperactive = False

'        Else
'            VpmNudge.SolGameOn True
'            OldState1 = False
'            GiOn
'            Flipperactive = True
'            ' activate bumpers and slings in case of tilt
'            Bumper1.Force = 13
'            Bumper2.Force = 13
'            LeftSlingshot.Disabled = 0
'            RightSlingshot.Disabled = 0
'     LightBumper001.State = ABS(Controller.Solenoid(4))
'     LightBumper002.State = ABS(Controller.Solenoid(4))
'        End If
'    End If
'End Sub


'The FrameTimer interval should be -1, so executes at the display frame rate
'The frame timer should be used to update anything visual, like some animations, shadows, etc.
'However, a lot of animations will be handled in their respective _animate subroutines.



' The frame timer interval is -1, so executes at the display frame rate
dim FrameTime, InitFrameTime : InitFrameTime = 0

Sub FrameTimer_Timer()
  FrameTime = gametime - InitFrameTime 'Calculate FrameTime as some animuations could use this
  InitFrameTime = gametime  'Count frametime
  'Add animation stuff here
  RollingUpdate
  BSUpdate
  UpdateBallBrightness
  RollingUpdate       'Update rolling sounds
  DoSTAnim          'Standup target animations
  DoDTAnim          'Drop target animations
  UpdateDropTargets
  AnimateBumperSkirts
End Sub

Sub GameTimer_Timer() 'The game timer interval; should be 10 ms
  Cor.Update      'update ball tracking (this sometimes goes in the RDampen_Timer sub)
  RollingUpdate   'update rolling sounds
  DoDTAnim    'handle drop target animations
  UpdateDropTargets
  DoSTAnim    'handle stand up target animations
  UpdateStandupTargets
  If VRMode = True Then
    VRBackglassTimer  ' Update VR BG digits and BG flashers
    UpdateVRLamps
  End If
  If Table1.ShowDT then UpdateDTLamps
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




'**********
' Keys
'**********


Sub Table1_KeyDown(ByVal KeyCode)
  If keycode = LeftFlipperKey and VRMode = True Then VR_ButtonLeft.TransX = 8 ' VRroom
  If keycode = RightFlipperKey and VRMode = True Then VR_ButtonRight.TransX = -8 ' VRroom
  If keycode = StartGameKey and VRMode = True Then VR_ButtonStart.TransY = -8 ' VRroom

  If keycode = LeftTiltKey Then : Nudge 90, 2 : SoundNudgeLeft :End If
  If keycode = RightTiltKey Then : Nudge 270, 2 : SoundNudgeRight : End If
  If keycode = CenterTiltKey Then : Nudge 0, 2 : SoundNudgeCenter : End If
  If keycode = LeftFlipperKey AND flipperactive Then SolLFlipper 1
    If keycode = RightFlipperKey AND flipperactive Then SolRFlipper 1
' If keycode = LeftMagnaSave Then bLutActive = True
'    If keycode = RightMagnaSave Then
'        If bLutActive Then NextLUT:End If
'    End If
    If KeyCode=StartGameKey Then vpmTimer.PulseSw 2: Exit Sub
'   vpmTimer.PulseSw 2
'   Exit Sub
' End If
  If KeyDownHandler(KeyCode)Then Exit Sub
  If keycode = PlungerKey Then
    Plunger.PullBack
    SoundPlungerPull
'   TimerVRPlunger.Enabled = True
'   TimerVRPlunger2.Enabled = False
  End If
End Sub

Sub Table1_KeyUp(ByVal KeyCode)
  If keycode = LeftFlipperKey and VRMode = True Then VR_ButtonLeft.TransX = 0 ' VRroom
  If keycode = RightFlipperKey and VRMode = True Then VR_ButtonRight.TransX = 0 ' VRroom
  If keycode = StartGameKey and VRMode = True Then VR_ButtonStart.TransY = 0 ' VRroom

' If keycode = LeftMagnaSave Then bLutActive = False
  If keycode = LeftFlipperKey AND flipperactive Then SolLFlipper 0
    If keycode = RightFlipperKey AND flipperactive Then SolRFlipper 0
  If KeyUpHandler(KeyCode)Then Exit Sub
  If KeyCode = PlungerKey Then
    Plunger.Fire
    SoundPlungerReleaseBall
'   TimerVRPlunger.Enabled = False
'   TimerVRPlunger2.Enabled = True
'   VR_Shooter.Y = 0
  End If
  End Sub



'******************************************************
'   ZBBR: BALL BRIGHTNESS
'******************************************************

Const BallBrightness =  1       'Ball brightness - Value between 0 and 1 (0=Dark ... 1=Bright)

' Constants for plunger lane ball darkening.
' You can make a temporary wall in the plunger lane area and use the co-ordinates from the corner control points.
Const PLOffset = 0.5      'Minimum ball brightness scale in plunger lane
Const PLLeft = 850        'X position of punger lane left
Const PLRight = 930       'X position of punger lane right
Const PLTop = 1125        'Y position of punger lane top
Const PLBottom = 1800       'Y position of punger lane bottom
Dim PLGain: PLGain = (1-PLOffset)/(PLTop-PLBottom)


Sub UpdateBallBrightness
  Dim s, b_base, b_r, b_g, b_b, d_w
  b_base = 120 * BallBrightness + 135*gilvl ' orig was 120 and 70

  Dim gBOT: gBOT = Getballs
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





'*******************************************
' ZDRN: Drain, Trough, and Ball Release
'*******************************************

Sub sw46_Hit
  Controller.Switch(13) = 1
  RandomSoundDrain sw46
End Sub

Sub SolRelease(enabled)
  If enabled Then
    RandomSoundBallRelease sw46
    sw46.kick 60, 20
    Controller.Switch(13) = 0
    vpmTimer.PulseSw 17
    PlaySoundAt SoundFX(DTResetSound,DOFContactors),BM_sw41
    DTRaise 41
    DTRaise 43
    DTRaise 42
    DTRaise 44
    Controller.Switch(41)=0
    Controller.Switch(42)=0
    Controller.Switch(43)=0
    Controller.Switch(44)=0
    Controller.Switch(3)=0
    Controller.Switch(4)=0
  End If
End Sub


'****************************************************************
' ZSLG: Slingshots
'****************************************************************

' RStep and LStep are the variables that increment the animation
Dim RStep, LStep

Sub RightSlingShot_Slingshot
  RS.VelocityCorrect(ActiveBall)
  vpmTimer.PulseSw 34
    RandomSoundSlingshotRight BM_RSlingArm
    RStep = 0
  RightSlingShot.TimerInterval = 17
    RightSlingShot.TimerEnabled = 1
End Sub

Sub RightSlingShot_Timer
  Dim BP
  Dim x1, x2, y: x1 = False:x2 = True:y = 15
    Select Case RStep
        Case 2:x1 = True:x2 = False:y = 9
        Case 3:x1 = False:x2 = False:y = 0:RightSlingShot.TimerEnabled = 0
    End Select

  For Each BP in BP_RSling1 : BP.Visible = x1: Next
  For Each BP in BP_RSling2 : BP.Visible = x2: Next
  For Each BP in BP_RSlingArm : BP.transx = y: Next

    RStep = RStep + 1
End Sub

Sub LeftSlingShot_Slingshot
  LS.VelocityCorrect(ActiveBall)
  vpmTimer.PulseSw 34
    RandomSoundSlingshotLeft BM_LSlingArm
    LStep = 0
  LeftSlingShot.TimerInterval = 17
    LeftSlingShot.TimerEnabled = 1
End Sub

Sub LeftSlingShot_Timer
  Dim BP
  Dim x1, x2, y: x1 = False:x2 = True:y = 15
    Select Case LStep
        Case 2:x1 = True:x2 = False:y = 9
        Case 3:x1 = False:x2 = False:y = 0:LeftSlingShot.TimerEnabled = 0
    End Select

  For Each BP in BP_LSling1 : BP.Visible = x1: Next
  For Each BP in BP_LSling2 : BP.Visible = x2: Next
  For Each BP in BP_LSlingArm : BP.transx = y: Next

    LStep = LStep + 1
End Sub


'*********
' Switches
'*********

  'Spinner
Sub sw32_Spin():vpmTimer.PulseSw 21:SoundSpinner sw32 :End Sub '21 (Spinner)



  ' stand up Targets
Sub sw23_Hit: STHit 23 : End Sub    '23 (Left Corridor Red Target) sw36
Sub sw47_Hit: STHit 47 : End Sub    '47 Top Right Blue Target (Bumpers Off Solenoid 4) sw12
Sub sw22_Hit: STHit 22 : End Sub    '22 (Special Target) sw37


  'Bumpers'
Sub bumper1_Hit:vpmTimer.PulseSw 27:RandomSoundBumperTop Bumper1:End Sub  '27 (Bumpers)
Sub bumper2_Hit:vpmTimer.PulseSw 27:RandomSoundBumperMiddle Bumper2:End Sub


'Rollovers
Sub Sw13_Hit:Controller.Switch(46)=1:End Sub              '46 (Left Outlane)
Sub Sw13_unHit:Controller.Switch(46)=0:End Sub
Sub Sw34_Hit:Controller.Switch(25)=1:leftInlaneSpeedLimit:End Sub   '25 (Left Inlane)
Sub Sw34_unHit:Controller.Switch(25)=0:End Sub
Sub Sw14_Hit:Controller.Switch(45)=1:End Sub              '45 (Right Outlane)
Sub Sw14_unHit:Controller.Switch(45)=0:End Sub
Sub Sw33_Hit:Controller.Switch(26)=1:rightInlaneSpeedLimit:End Sub    '26 (Right Inlane)
Sub Sw33_unHit:Controller.Switch(26)=0:End Sub
Sub Sw35_Hit:Controller.Switch(24)=1:End Sub              '24 (Left Corridor Rollover)
Sub Sw35_unHit:Controller.Switch(24)=0:End Sub
Sub Sw24_Hit:Controller.Switch(35)=1:End Sub              '35 (Left Top Lane "2")
Sub Sw24_unHit:Controller.Switch(35)=0:End Sub
Sub Sw36_Hit:Controller.Switch(36)=1:End Sub              '36 (Right Top Lane "5")
Sub Sw36_unHit:Controller.Switch(36)=0:End Sub
Sub Sw28_Hit:Controller.Switch(31)=1:End Sub              '31 (Right Corridor)
Sub Sw28_unHit:Controller.Switch(31)=0:End Sub

Sub leftInlaneSpeedLimit
  'Wylte's implementation
'    debug.print "Spin in: "& activeball.AngMomZ
'    debug.print "Speed in: "& activeball.vely
  if activeball.vely < 0 then exit sub              'don't affect upwards movement
    activeball.AngMomZ = -abs(activeball.AngMomZ) * RndNum(3,6)
    If abs(activeball.AngMomZ) > 60 Then activeball.AngMomZ = 0.8 * activeball.AngMomZ
    If abs(activeball.AngMomZ) > 80 Then activeball.AngMomZ = 0.8 * activeball.AngMomZ
    If activeball.AngMomZ > 100 Then activeball.AngMomZ = RndNum(80,100)
    If activeball.AngMomZ < -100 Then activeball.AngMomZ = RndNum(-80,-100)

    if abs(activeball.vely) > 5 then activeball.vely = 0.8 * activeball.vely
    if abs(activeball.vely) > 10 then activeball.vely = 0.8 * activeball.vely
    if abs(activeball.vely) > 15 then activeball.vely = 0.8 * activeball.vely
    if activeball.vely > 16 then activeball.vely = RndNum(14,16)
    if activeball.vely < -16 then activeball.vely = RndNum(-14,-16)
'    debug.print "Spin out: "& activeball.AngMomZ
'    debug.print "Speed out: "& activeball.vely
End Sub


Sub rightInlaneSpeedLimit
  'Wylte's implementation
'    debug.print "Spin in: "& activeball.AngMomZ
'    debug.print "Speed in: "& activeball.vely
  if activeball.vely < 0 then exit sub              'don't affect upwards movement
    activeball.AngMomZ = abs(activeball.AngMomZ) * RndNum(2,4)
    If abs(activeball.AngMomZ) > 60 Then activeball.AngMomZ = 0.8 * activeball.AngMomZ
    If abs(activeball.AngMomZ) > 80 Then activeball.AngMomZ = 0.8 * activeball.AngMomZ
    If activeball.AngMomZ > 100 Then activeball.AngMomZ = RndNum(80,100)
    If activeball.AngMomZ < -100 Then activeball.AngMomZ = RndNum(-80,-100)

  if abs(activeball.vely) > 5 then activeball.vely = 0.8 * activeball.vely
    if abs(activeball.vely) > 10 then activeball.vely = 0.8 * activeball.vely
    if abs(activeball.vely) > 15 then activeball.vely = 0.8 * activeball.vely
    if activeball.vely > 16 then activeball.vely = RndNum(14,16)
    if activeball.vely < -16 then activeball.vely = RndNum(-14,-16)
'    debug.print "Spin out: "& activeball.AngMomZ
'    debug.print "Speed out: "& activeball.vely
End Sub

'Generic rubbers and lower slingshots, buttons
Sub Sw25A_Hit:Controller.Switch(34)=1:PlaySoundAt "fx_sensor", sw25a:End Sub    '34
Sub Sw25A_unHit:Controller.Switch(34)=0:End Sub
Sub Sw25B_Hit:Controller.Switch(34)=1:PlaySoundAt "fx_sensor", sw25b:End Sub
Sub Sw25B_unHit:Controller.Switch(34)=0:End Sub
Sub Sw25E_Hit:vpmTimer.PulseSw 34:End Sub
Sub Sw25F_Hit:vpmTimer.PulseSw 34:End Sub
Sub sw37_Hit:Controller.Switch(37) = 1:PlaySoundAt "fx_sensor", sw37:End Sub '37 (Top Rollover Button "6") sw22
Sub sw37_unHit:Controller.Switch(37) = 0 : End Sub
Sub sw27_Hit:Controller.Switch(32) = 1:PlaySoundAt "fx_sensor", sw27:End Sub   '32 (ON Bumpers Rollover)
Sub sw27_unHit:Controller.Switch(32) = 0 : End Sub
Sub sw31a_Hit:vpmTimer.PulseSw 28:End Sub   '28 (Right Rubber and small left rubber)
Sub sw31b_Hit:vpmTimer.PulseSw 28:End Sub
Sub sw26a_Hit:vpmTimer.PulseSw 33:End Sub   '33 (Rubbers behind drop targets)
Sub sw26b_Hit:vpmTimer.PulseSw 33:End Sub




  '38 (seta Bumnpers)
Sub sw38a_Hit
    PlaySoundAtBall "fx_passive_bumper"
  vpmTimer.PulseSw 38
  seta38a.TransY = -3: Me.TimerEnabled = 1
End Sub

Sub sw38a_Timer:seta38a.TransY = 0: Me.TimerEnabled = 0: End Sub

Sub sw38b_Hit
    PlaySoundAtBall "fx_passive_bumper"
  vpmTimer.PulseSw 38
    seta38b.TransY = -3: Me.TimerEnabled = 1
End Sub

Sub sw38b_Timer:seta38b.TransY = 0: Me.TimerEnabled = 0: End Sub

Sub sw38c_Hit
    PlaySoundAtBall "fx_passive_bumper"
  vpmTimer.PulseSw 38
  vpmTimer.PulseSw 34
    seta38c.TransY = -3: Me.TimerEnabled = 1
End Sub

Sub sw38c_Timer:seta38c.TransY = 0: Me.TimerEnabled = 0: End Sub

Sub sw38d_Hit
    PlaySoundAtBall "fx_passive_bumper"
  vpmTimer.PulseSw 38
  vpmTimer.PulseSw 34
    seta38d.TransY = -3: Me.TimerEnabled = 1
End Sub

Sub sw38d_Timer:seta38d.TransY = 0: Me.TimerEnabled = 0: End Sub

' Drop targets

'Sub sw18_Dropped                   '41 (Drop Target "A")
' PlaySoundat SoundFX("fx_droptarget",DOFDropTargets), sw18
' Controller.Switch(41)=1
' If sw16.IsDropped Then Controller.Switch(4)=1
' If Table1.ShowDT then AroLight
'End Sub



'Sub sw17_Dropped                   '42 (Drop Target "C")
' PlaySoundat SoundFX("fx_droptarget",DOFDropTargets), sw17
' Controller.Switch(42)=1
' If sw15.IsDropped Then Controller.Switch(3)=1
' If Table1.ShowDT then AroLight
'End Sub

'Sub sw16_Dropped                   '43 (Drop Target "B")
' PlaySoundat SoundFX("fx_droptarget",DOFDropTargets), sw16
' Controller.Switch(43)=1
'   If sw18.IsDropped Then Controller.Switch(4)=1
' If Table1.ShowDT then AroLight
'End Sub

'Sub sw15_Dropped                   '44 (Drop Target "D")
' PlaySoundat SoundFX("fx_droptarget",DOFDropTargets), sw15
' Controller.Switch(44)=1
' If sw17.IsDropped Then Controller.Switch(3)=1
' If Table1.ShowDT then AroLight
'End Sub


' Drop Targets
Dim ind

Sub sw41_hit
  DTHit 41
  ind = DTArrayID(43)
  If DTArray(ind).isDropped = True Then Controller.Switch(4)=1
End Sub

Sub sw42_hit
  ind = DTArrayID(44)
  If DTArray(ind).isDropped = True Then Controller.Switch(3)=1
  DTHit 42
End Sub

Sub sw43_hit
  DTHit 43
  ind = DTArrayID(41)
  If DTArray(ind).isDropped = True Then Controller.Switch(4)=1
End Sub

Sub sw44_hit
  DTHit 44
  ind = DTArrayID(42)
  If DTArray(ind).isDropped = True Then Controller.Switch(3)=1
End Sub




'Extra Lights

Set LampCallback = GetRef("UpdateMultipleLamps")

Sub UpdateMultipleLamps
  L48.SetValue ABS(Controller.Lamp(48))
    bip1.Setvalue ABS(Controller.Lamp(2))
    bip2.Setvalue ABS(Controller.Lamp(3))
    bip3.Setvalue ABS(Controller.Lamp(4))
    bip4.Setvalue ABS(Controller.Lamp(5))
    bip5.Setvalue ABS(Controller.Lamp(6))
  rhigh.SetValue ABS(Controller.Lamp(7))
  GameOverR.SetValue ABS(Controller.Lamp(1))
  TiltR.SetValue ABS(Controller.Lamp(47))
' pl1.Setvalue ABS(Controller.Lamp(49))
  aro1.SetValue ABS (Controller.Sol(2))

End Sub

Sub UpdateVRLamps
  IF Controller.Lamp(1) = 0 Then: VRBG_GO.visible = 0 : Else : VRBG_GO.visible = 1 'Game Over
  IF Controller.Lamp(2) = 0 Then: VRBG_bip1.visible = 0 : Else : VRBG_bip1.visible = 1 'Ball 1
  IF Controller.Lamp(3) = 0 Then: VRBG_bip2.visible = 0 : Else : VRBG_bip2.visible = 1 'Ball 2
  IF Controller.Lamp(4) = 0 Then: VRBG_bip3.visible = 0 : Else : VRBG_bip3.visible = 1 'Ball 3
  IF Controller.Lamp(5) = 0 Then: VRBG_bip4.visible = 0 : Else : VRBG_bip4.visible = 1 'Ball 4
  IF Controller.Lamp(6) = 0 Then: VRBG_bip5.visible = 0 : Else : VRBG_bip5.visible = 1 'Ball 5
  IF Controller.Lamp(7) = 0 Then: VRBG_HS.visible = 0 : Else : VRBG_HS.visible = 1 'Hight Score
  IF Controller.Lamp(47) = 0 Then: VRBG_tilt.visible = 0 : Else : VRBG_tilt.visible = 1 'tilt
  IF Controller.Lamp(49) = 0 Then: VRBG_pl1.visible = 0 : Else : VRBG_pl1.visible = 1 'Player 1
  IF Controller.Lamp(50) = 0 Then: VRBG_pl2.visible = 0 : Else : VRBG_pl2.visible = 1 'Player 2
  IF Controller.Lamp(51) = 0 Then: VRBG_pl3.visible = 0 : Else : VRBG_pl3.visible = 1 'Player 3
  IF Controller.Lamp(52) = 0 Then: VRBG_pl4.visible = 0 : Else : VRBG_pl4.visible = 1 'Player 4
  IF Controller.Lamp(12) = 0 Then: VRBG_cp1.visible = 0 : Else : VRBG_cp1.visible = 1 'Can Play 1
  IF Controller.Lamp(11) = 0 Then: VRBG_cp2.visible = 0 : Else : VRBG_cp2.visible = 1 'Can Play 2
  IF Controller.Lamp(10) = 0 Then: VRBG_cp3.visible = 0 : Else : VRBG_cp3.visible = 1 'Can Play 3
  IF Controller.Lamp(9) = 0 Then: VRBG_cp4.visible = 0 : Else : VRBG_cp4.visible = 1 'Can Play 4
End Sub


Sub UpdateDTLamps
  IF Controller.Lamp(49) = 0 Then: pl1.visible = 0 : Else : pl1.visible = 1 'Player 1
  IF Controller.Lamp(50) = 0 Then: pl2.visible = 0 : Else : pl2.visible = 1 'Player 2
  IF Controller.Lamp(51) = 0 Then: pl3.visible = 0 : Else : pl3.visible = 1 'Player 3
  IF Controller.Lamp(52) = 0 Then: pl4.visible = 0 : Else : pl4.visible = 1 'Player 4
End Sub


'*******************
' FLIPPERS
'*******************

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
    If rightflipper.currentangle > rightflipper.endangle - ReflipAngle Then
      RandomSoundReflipUpRight RightFlipper
    Else
      SoundFlipperUpAttackRight RightFlipper
      RandomSoundFlipperUpRight RightFlipper
    End If
  Else
    RightFlipper.RotateToStart
    If RightFlipper.currentangle > RightFlipper.startAngle + 5 Then
      RandomSoundFlipperDownRight RightFlipper
    End If
    FlipperRightHitParm = FlipperUpSoundLevel
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



'Sub LeftFlipper_Animate
' Dim a : a = LeftFlipper.CurrentAngle
' FlipperLSh.RotZ = a

' Dim v, BP
' v = 255.0 * (122.0 -  LeftFlipper.CurrentAngle) / (122.0 -  70.0)

' For each BP in BP_LFlipper
'   BP.Rotz = a
'   BP.visible = v < 128.0
' Next
' For each BP in BP_LFlipperU
'   BP.Rotz = a
'   BP.visible = v >= 128.0
' Next
'End Sub

'Sub RightFlipper_Animate
' Dim a : a = RightFlipper.CurrentAngle
' FlipperRSh.RotZ = a

' Dim v, BP
' v = 255.0 * (-122.0 -  RightFlipper.CurrentAngle) / (-122.0 +  70.0)

' For each BP in BP_RFlipper
'   BP.Rotz = a
'   BP.visible = v < 128.0
' Next
' For each BP in BP_RFlipperU
'   BP.Rotz = a
'   BP.visible = v >= 128.0
' Next
'End Sub





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

Set Digits(6) = b0
Set Digits(7) = b1
Set Digits(8) = b2
Set Digits(9) = b3
Set Digits(10) = b4
Set Digits(11) = b5

Set Digits(12) = c0
Set Digits(13) = c1
Set Digits(14) = c2
Set Digits(15) = c3
Set Digits(16) = c4
Set Digits(17) = c5

Set Digits(18) = d0
Set Digits(19) = d1
Set Digits(20) = d2
Set Digits(21) = d3
Set Digits(22) = d4
Set Digits(23) = d5

Set Digits(24) = e0
Set Digits(25) = e1
Set Digits(26) = e2
Set Digits(27) = e3
Set Digits(28) = e4

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
    End If
  If P42.State =1 Then
    P42b.State=1
    Else
      P42b.State=0
  End If
End Sub



'******************************************************
'****  LAMPZ by nFozzy
'******************************************************
'
' Lampz is a utility designed to manage and fade the lights and light-related objects on a table that is being driven by a ROM.
' To set up Lampz, one must populate the Lampz.MassAssign array with VPX Light objects, where the index of the MassAssign array
' corrisponds to the ROM index of the associated light. More that one Light object can be associated with a single MassAssign index (not shown in this example)
' Optionally, callbacks can be assigned for each index using the Lampz.Callback array. This is very useful for allowing 3D Insert primitives
' to be controlled by the ROM. Note, the aLvl parameter (i.e. the fading level that ranges between 0 and 1) is appended to the callback call.

Dim NullFader : set NullFader = new NullFadingObject
Dim Lampz : Set Lampz = New LampFader
Dim FadingState(200)
InitLampsNF               ' Setup lamp assignments
LampTimer.Interval = -1
LampTimer.Enabled = 1

Sub LampTimer_Timer()
  dim x, chglamp
  if UsingROM then chglamp = Controller.ChangedLamps
  If Not IsEmpty(chglamp) Then
    For x = 0 To UBound(chglamp)      'nmbr = chglamp(x, 0), state = chglamp(x, 1)
      Lampz.state(chglamp(x, 0)) = chglamp(x, 1)
      FadingState(chgLamp(x, 0)) = chgLamp(x, 1) + 3 'fading step
    next
  End If
  Lampz.Update2 'update (fading logic only)
  If VRMode= False Then UpdateLeds
' UpdateTexts
End Sub

Sub DisableLighting(pri, DLintensity, ByVal aLvl) 'cp's script  DLintensity = disabled lighting intesity
  if Lampz.UseFunction then aLvl = Lampz.FilterOut(aLvl)  'Callbacks don't get this filter automatically
  pri.blenddisablelighting = aLvl * DLintensity
End Sub

Sub SetModLamp(id, val)
  Lampz.state(id) = val
End Sub


Sub InitLampsNF()

  'Filtering (comment out to disable)
  Lampz.Filter = "LampFilter" 'Puts all lamp intensityscale output (no callbacks) through this function before updating

  'Adjust fading speeds (max level / full MS fading time). The Modulate property must be set to 1 / max level if lamp is modulated.
  dim x : for x = 0 to 150 : Lampz.FadeSpeedUp(x) = 1/40 : Lampz.FadeSpeedDown(x) = 1/120 : Lampz.Modulate(x) = 1 : next


  'Lampz Assignments
  '  In a ROM based table, the lamp ID is used to set the state of the Lampz objects

  'MassAssign is an optional way to do assignments. It'll create arrays automatically / append objects to existing arrays

  Lampz.MassAssign(9)= cp4
  Lampz.MassAssign(10)= cp3
  Lampz.MassAssign(11)= cp2
  Lampz.MassAssign(12)= cp1
  Lampz.MassAssign(13)= A13
  Lampz.MassAssign(14)= A14
  Lampz.MassAssign(15)= A15
  Lampz.MassAssign(17)= A17
  Lampz.MassAssign(18)= A18
  Lampz.MassAssign(19)= A19
  Lampz.MassAssign(20)= A20
  Lampz.MassAssign(21)= A21
  Lampz.MassAssign(22)= A22
  Lampz.MassAssign(23)= A23
  Lampz.MassAssign(24)= A24
  Lampz.MassAssign(25)= A25
  Lampz.MassAssign(26)= A26
  Lampz.MassAssign(27)= A27
  Lampz.MassAssign(28)= A28
  Lampz.MassAssign(29)= A29
  Lampz.MassAssign(30)= A30
  Lampz.MassAssign(31)= A31
  Lampz.MassAssign(32)= A32
  Lampz.MassAssign(33)= A33
  Lampz.MassAssign(34)= A34
  Lampz.MassAssign(35)= A35
  Lampz.MassAssign(36)= A36
  Lampz.MassAssign(37)= A37
  Lampz.MassAssign(38)= A38
  Lampz.MassAssign(39)= A39
  Lampz.MassAssign(40)= A40
  Lampz.MassAssign(41)= A41
  Lampz.MassAssign(42)= GI13
  Lampz.MassAssign(43)= A43
  Lampz.MassAssign(44)= A44
  Lampz.MassAssign(44)= A44
  Lampz.MassAssign(45)= A45A
  Lampz.MassAssign(45)= A45B
  Lampz.MassAssign(46)= A46A
  Lampz.MassAssign(46)= A46B
' Lampz.MassAssign(48)= LSA
  Lampz.MassAssign(49)= pl1
  Lampz.MassAssign(50)= pl2
  Lampz.MassAssign(51)= pl3
  Lampz.MassAssign(52)= pl4


  'Turn off all lamps on startup
  Lampz.Init  'This just turns state of any lamps to 1

  'Immediate update to turn on GI, turn off lamps
  Lampz.Update

End Sub

'Sub UpdateTexts()
  'backdrop lights
'   Textm 78, l78a, "OVER"
'   Text 78, l78, "GAME"
'   Text 80, l80, "TILT"
'End Sub

'Texts

Sub Text(nr, object, message)
    Select Case FadingState(nr)
        Case 4:object.Text = message:FadingState(nr) = 0
        Case 3:object.Text = "":FadingState(nr) = 0
    End Select
End Sub

Sub Textm(nr, object, message)
    Select Case FadingState(nr)
        Case 4:object.Text = message
        Case 3:object.Text = ""
    End Select
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

'******************************
'****  END LAMPZ
'******************************

'*******************
' GI Lights
'*******************

Dim gilvl : gilvl=1

Sub GiOn
    Dim bulb
    For each bulb in GiLights
        bulb.State = 1
    Next
  If VRMode = True Then VR_BkGlassOn.visible=1
End Sub

Sub GiOff
    Dim bulb
    For each bulb in GILights
        bulb.State = 0
    Next
  If VRMode = True Then VR_BkGlassOn.visible=0
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
Dim objBallShadow(7)

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
  Dim gBOT: gBOT=Getballs
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
Dim ULF : Set ULF = New FlipperPolarity

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
' a = Array(LF, RF, ULF)
' For Each x In a
'   x.AddPt "Ycoef", 0, RightFlipper.Y-65, 1 'disabled
'   x.AddPt "Ycoef", 1, RightFlipper.Y-11, 1
'   x.enabled = True
'   x.TimeDelay = 60
'   x.DebugOn=False ' prints some info in debugger

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

' Next

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

Dim LFPress, RFPress, ULFPress, LFCount, RFCount
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

'****************************************************************
'****  END VPW TargetBouncer for targets and posts by Iaakki, Wrd1972, Apophis
'****************************************************************


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
Dim DT41, DT42, DT43, DT44

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

Set DT41 = (new DropTarget)(sw41, sw41a, BM_sw41, 41, 0, false)
Set DT42 = (new DropTarget)(sw42, sw42a, BM_sw42, 42, 0, false)
Set DT43 = (new DropTarget)(sw43, sw43a, BM_sw43, 43, 0, false)
Set DT44 = (new DropTarget)(sw44, sw44a, BM_sw44, 44, 0, false)

Dim DTArray
DTArray = Array(DT41,DT42,DT43,DT44)

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
Dim ST22, ST23, ST47

'Set array with stand-up target objects
'
'StandupTargetvar = Array(primary, prim, switch)
'   primary:  vp target to determine target hit
'   prim:    primitive target used for visuals and animation
'          IMPORTANT!!!
'          transy must be used to offset the target animation
'   switch:  ROM switch number
'   animate:  Arrary slot for handling the animation instrucitons, set to 0
'
'You will also need to add a secondary hit object for each stand up (name sw11o, sw12o, and sw13o on the example Table1)
'these are inclined primitives to simulate hitting a bent target and should provide so z velocity on high speed impacts

Set ST47 = (new StandupTarget)(sw47, BM_sw47,47, 0)   ' sw12
Set ST23 = (new StandupTarget)(sw23, BM_sw23,23, 0)   ' sw36
Set ST22 = (new StandupTarget)(sw22, BM_sw22,22, 0)   ' sw37


'Add all the Stand-up Target Arrays to Stand-up Target Animation Array
' STAnimationArray = Array(ST1, ST2, ....)
Dim STArray
STArray = Array(ST47, ST23, ST22)

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
' Tutorial vides by Apophis
' Audio : Adding Fleep Part 1       https://youtu.be/rG35JVHxtx4
' Audio : Adding Fleep Part 2       https://youtu.be/dk110pWMxGo
' Audio : Adding Fleep Part 3       https://youtu.be/ESXWGJZY_EI


'///////////////////////////////  SOUNDS PARAMETERS  //////////////////////////////
Dim GlobalSoundLevel, CoinSoundLevel, PlungerReleaseSoundLevel, PlungerPullSoundLevel, NudgeLeftSoundLevel
Dim NudgeRightSoundLevel, NudgeCenterSoundLevel, StartButtonSoundLevel, RollingSoundFactor

CoinSoundLevel = 0.5            'volume level; range [0, 1]
NudgeLeftSoundLevel = 1        'volume level; range [0, 1]
NudgeRightSoundLevel = 1        'volume level; range [0, 1]
NudgeCenterSoundLevel = 1        'volume level; range [0, 1]
StartButtonSoundLevel = 0.1      'volume level; range [0, 1]
PlungerReleaseSoundLevel = 0.8 '1 wjr   'volume level; range [0, 1]
PlungerPullSoundLevel = 1        'volume level; range [0, 1]
RollingSoundFactor = 1 / 5

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
SlingshotSoundLevel = 0.8           'volume level; range [0, 1]
BumperSoundFactor = 3           'volume multiplier; must not be zero
KnockerSoundLevel = 0.5              'volume level; range [0, 1]

'///////////////////////-----Ball Drops, Bumps and Collisions-----///////////////////////
Dim RubberStrongSoundFactor, RubberWeakSoundFactor, RubberFlipperSoundFactor,BallWithBallCollisionSoundFactor
Dim BallBouncePlayfieldSoftFactor, BallBouncePlayfieldHardFactor, PlasticRampDropToPlayfieldSoundLevel, WireRampDropToPlayfieldSoundLevel, DelayedBallDropOnPlayfieldSoundLevel
Dim WallImpactSoundFactor, MetalImpactSoundFactor, SubwaySoundLevel, SubwayEntrySoundLevel, ScoopEntrySoundLevel
Dim SaucerLockSoundLevel, SaucerKickSoundLevel

BallWithBallCollisionSoundFactor = 3      'volume multiplier; must not be zero
RubberStrongSoundFactor = 0.070 / 4      'volume multiplier; must not be zero
RubberWeakSoundFactor = 0.070 / 5        'volume multiplier; must not be zero
RubberFlipperSoundFactor = 0.075 / 5      'volume multiplier; must not be zero
BallBouncePlayfieldSoftFactor = 0.025      'volume multiplier; must not be zero
BallBouncePlayfieldHardFactor = 0.025      'volume multiplier; must not be zero
DelayedBallDropOnPlayfieldSoundLevel = 0.8    'volume level; range [0, 1]
WallImpactSoundFactor = 0.075          'volume multiplier; must not be zero
MetalImpactSoundFactor = 0.045 / 3
SaucerLockSoundLevel = 0.5
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

DrainSoundLevel = 0.5          'volume level; range [0, 1]
BallReleaseSoundLevel = 0.5        'volume level; range [0, 1]
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
  PlaySound playsoundparams, 0, aVol * VolumeDial, AudioPan(tableobj), 0, 0, 0, 0, AudioFade(tableobj)
End Sub

Sub PlaySoundAtLevelExistingStatic(playsoundparams, aVol, tableobj)
  PlaySound playsoundparams, 0, aVol * VolumeDial, AudioPan(tableobj), 0, 0, 1, 0, AudioFade(tableobj)
End Sub

Sub PlaySoundAtLevelStaticLoop(playsoundparams, aVol, tableobj)
  PlaySound playsoundparams, - 1, aVol * VolumeDial, AudioPan(tableobj), 0, 0, 0, 0, AudioFade(tableobj)
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
  PlaySound playsoundparams, - 1, aVol * VolumeDial, AudioPan(tableobj), randomPitch, 0, 0, 0, AudioFade(tableobj)
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
  PlaySound soundname, 1,aVol * VolumeDial, AudioPan(ActiveBall), 0,0,0, 1, AudioFade(ActiveBall)
End Sub

Sub PlaySoundAtBallVolM (Soundname, aVol)
  PlaySound soundname, 1,aVol * VolumeDial, AudioPan(ActiveBall), 0,0,0, 0, AudioFade(ActiveBall)
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
  TargetBouncer ActiveBall, 1
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
'   Misc Animations
'******************************************************

''''' Flippers

Sub LeftFlipper_Animate
  Dim a : a = LeftFlipper.CurrentAngle
  FlipperLSh.RotZ = a

  Dim v, BP
  v = 255.0 * (121.0 - LeftFlipper.CurrentAngle) / (121.0 -  70.0)

  For each BP in BP_LFlipper
    BP.Rotz = a
    BP.visible = v < 128.0
  Next
  For each BP in BP_LFlipperU
    BP.Rotz = a
    BP.visible = v >= 128.0
  Next
End Sub

Sub RightFlipper_Animate
  Dim a : a = RightFlipper.CurrentAngle
  FlipperRSh.RotZ = a

  Dim v, BP
  v = 255.0 * (-121.0 - RightFlipper.CurrentAngle) / (-121.0 +  70.0)

  For each BP in BP_RFlipper
    BP.Rotz = a
    BP.visible = v < 128.0
  Next
  For each BP in BP_RFlipperU
    BP.Rotz = a
    BP.visible = v >= 128.0
  Next
End Sub

''''' Bumper Animations

Sub Bumper1_Animate
  Dim z, BP
  z = Bumper1.CurrentRingOffset
  For Each BP in BP_BR1 : BP.transz = z: Next
End Sub

Sub Bumper2_Animate
  Dim z, BP
  z = Bumper2.CurrentRingOffset
  For Each BP in BP_BR2 : BP.transz = z: Next
End Sub


Dim Bumpers : Bumpers = Array(Bumper1, Bumper2)

Sub AnimateBumperSkirts
  dim r, g, s, x, y, b, tz
  ' Animate Bumper switch (experimental)
  For r = 0 To 1
    g = 10000.0
    Dim gBOT : gBOT = GetBalls
    For s = 0 to UBound(gBOT)
      x = Bumpers(r).x - gBOT(s).x
      y = Bumpers(r).y - gBOT(s).y
      b = x * x + y * y
      If b < g Then g = b
    Next
    tz = 4
    If g < 80 * 80 Then
      tz = 1
    End If
    If r = 0 Then For Each x in BP_BS1: x.Z = tz: Next
    If r = 1 Then For Each x in BP_BS2: x.Z = tz: Next
  Next
End Sub

''''' Gate Animations

Sub Gate2_Animate
  Dim a : a = Gate2.CurrentAngle
  Dim BP : For Each BP in BP_Gate2 : BP.rotx = a: Next
End Sub

''''' Spinner Animations

Dim SpinnerRadius: SpinnerRadius=7
Sub sw32_Animate
  Dim spinangle:spinangle = sw32.currentangle
  Dim BL : For Each BL in BP_sw32 : BL.RotX = -spinangle: Next
  BM_SRod.TransZ = (cos((sw32.CurrentAngle + 180) * (PI/180))+1) * SpinnerRadius
  BM_SRod.TransY = sin((sw32.CurrentAngle) * (PI/180)) * -SpinnerRadius
' SpinnerTShadow.size_y = abs(sin( (spinangle+180) * (2*PI/360)) * 5)
End Sub

''''' Switch Animations

Sub sw13_Animate
  Dim z : z = sw13.CurrentAnimOffset
  Dim BP : For Each BP in BP_sw13 : BP.transz = z: Next
End Sub

Sub sw34_Animate
  Dim z : z = sw34.CurrentAnimOffset
  Dim BP : For Each BP in BP_sw34 : BP.transz = z: Next
End Sub

Sub sw33_Animate
  Dim z : z = sw33.CurrentAnimOffset
  Dim BP : For Each BP in BP_sw33 : BP.transz = z: Next
End Sub

Sub sw14_Animate
  Dim z : z = sw14.CurrentAnimOffset
  Dim BP : For Each BP in BP_sw14 : BP.transz = z: Next
End Sub

Sub sw35_Animate
  Dim z : z = sw35.CurrentAnimOffset
  Dim BP : For Each BP in BP_sw35 : BP.transz = z: Next
End Sub

Sub sw24_Animate
  Dim z : z = sw24.CurrentAnimOffset
  Dim BP : For Each BP in BP_sw24 : BP.transz = z: Next
End Sub

Sub sw36_Animate
  Dim z : z = sw36.CurrentAnimOffset
  Dim BP : For Each BP in BP_sw36 : BP.transz = z: Next
End Sub

Sub sw28_Animate
  Dim z : z = sw28.CurrentAnimOffset
  Dim BP : For Each BP in BP_sw28 : BP.transz = z: Next
End Sub

''''' Butt Animations

Sub sw27_Animate
  Dim z : z = sw27.CurrentAnimOffset
  Dim BP : For Each BP in BP_sw27 : BP.transz = z: Next
End Sub

Sub sw37_Animate
  Dim z : z = sw37.CurrentAnimOffset
  Dim BL : For Each BL in BP_sw37 : BL.transz = z: Next
End Sub

Sub sw25a_Animate
  Dim z : z = sw25a.CurrentAnimOffset
  Dim BL : For Each BL in BP_sw25a : BL.transz = z: Next
End Sub

Sub sw25b_Animate
  Dim z : z = sw25b.CurrentAnimOffset
  Dim BL : For Each BL in BP_sw25b : BL.transz = z: Next
End Sub


''''' Standup Targets

Sub UpdateStandupTargets
  dim BP, ty

  ty = BM_sw22.transy
  For each BP in BP_sw22 : BP.transy = ty: Next

  ty = BM_sw23.transy
  For each BP in BP_sw23 : BP.transy = ty: Next

  ty = BM_sw47.transy
  For each BP in BP_sw47 : BP.transy = ty: Next

End Sub

''''' Drop Targets

Sub UpdateDropTargets
  dim BP, tz, rx, ry

  tz = BM_sw41.transz
  rx = BM_sw41.rotx
  ry = BM_sw41.roty
  For each BP in BP_sw41: BP.transz = tz: BP.rotx = rx: BP.roty = ry: Next

  tz = BM_sw43.transz
  rx = BM_sw43.rotx
  ry = BM_sw43.roty
  For each BP in BP_sw43: BP.transz = tz: BP.rotx = rx: BP.roty = ry: Next

  tz = BM_sw42.transz
  rx = BM_sw42.rotx
  ry = BM_sw42.roty
  For each BP in BP_sw42: BP.transz = tz: BP.rotx = rx: BP.roty = ry: Next

  tz = BM_sw44.transz
  rx = BM_sw44.rotx
  ry = BM_sw44.roty
  For each BP in BP_sw44: BP.transz = tz: BP.rotx = rx: BP.roty = ry: Next

End Sub



'******************************
' ZVRR - Setup VR Room
'******************************

Sub SetupVRRoom
  dim VRThing, BP
  If RenderingMode = 2 or ForceVR = True Then
    VRMode = True
    For Each BP in BP_PincabRails : BP.Visible = 0: Next
    If VRRoom = 1 Then ' Space room
      for Each VRThing in VRThings: VRThing.visible = 1: Next
      for each VRThing in VRCabinet:VRThing.visible = 1:Next
    ElseIf VRRoom = 2 Then ' Ultra Minimal room
      for each VRThing in VRCabinet:VRThing.visible = 1:Next
      for Each VRThing in VRThings: VRThing.visible = 0: Next
    End If
  Else
    VRMode = False
    for each VRThing in VRCabinet:VRThing.visible = 0:Next
    for Each VRThing in VRThings: VRThing.visible = 0: Next
    For Each BP in BP_PincabRails : BP.Visible = 1: Next
    LightStopper.visible = 1
  End If
End Sub


Sub TimerVRPlunger_Timer
  If VR_Plunger.Y < 100 then
    VR_Plunger.Y = VR_Plunger.Y + 5
  End If
End Sub

Sub TimerVRPlunger2_Timer
  VR_Plunger.Y = (5* Plunger.Position) - 20
End Sub

'******************************
' Setup VR Backglass
'******************************

Dim xoff,yoff1, yoff2, yoff3, yoff4, yoff5,zoff,xrot,zscale, xcen,ycen

Sub SetupVRBackglass

  xoff = 0
  yoff1 = 83 ' this is where you adjust the forward/backward position for player 1 score
  yoff2 = 83 ' this is where you adjust the forward/backward position for player 2 score
  yoff3 = 83 ' this is where you adjust the forward/backward position for player 3 score
  yoff4 = 83 ' this is where you adjust the forward/backward position for player 4 score
  yoff5 = 83  ' this is where you adjust the forward/backward position for credits and ball in play
  zoff = 731
  xrot = -90

  CenterVRDigits

End Sub


Sub CenterVRDigits
  Dim ix, xx, yy, yfact, xfact, xobj

  zscale = 0.0000001

  xcen = (130 /2) - (92 / 2)
  ycen = (780 /2 ) + (203 /2)

  for ix = 0 to 5
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

  for ix = 6 to 11
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

  for ix = 12 to 17
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

  for ix = 18 to 23
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

  for ix = 24 to 28
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
'              Display VR Output
'********************************************


Dim VRDigits(29)
VRDigits(0) = Array(LED1x0,LED1x1,LED1x2,LED1x3,LED1x4,LED1x5,LED1x6)
VRDigits(1) = Array(LED2x0,LED2x1,LED2x2,LED2x3,LED2x4,LED2x5,LED2x6)
VRDigits(2) = Array(LED3x0,LED3x1,LED3x2,LED3x3,LED3x4,LED3x5,LED3x6)
VRDigits(3) = Array(LED4x0,LED4x1,LED4x2,LED4x3,LED4x4,LED4x5,LED4x6)
VRDigits(4) = Array(LED5x0,LED5x1,LED5x2,LED5x3,LED5x4,LED5x5,LED5x6)
VRDigits(5) = Array(LED6x0,LED6x1,LED6x2,LED6x3,LED6x4,LED6x5,LED6x6)

VRDigits(6) = Array(LED8x0,LED8x1,LED8x2,LED8x3,LED8x4,LED8x5,LED8x6)
VRDigits(7) = Array(LED9x0,LED9x1,LED9x2,LED9x3,LED9x4,LED9x5,LED9x6)
VRDigits(8) = Array(LED10x0,LED10x1,LED10x2,LED10x3,LED10x4,LED10x5,LED10x6)
VRDigits(9) = Array(LED11x0,LED11x1,LED11x2,LED11x3,LED11x4,LED11x5,LED11x6)
VRDigits(10) = Array(LED12x0,LED12x1,LED12x2,LED12x3,LED12x4,LED12x5,LED12x6)
VRDigits(11) = Array(LED13x0,LED13x1,LED13x2,LED13x3,LED13x4,LED13x5,LED13x6)

VRDigits(12) = Array(LED1x000,LED1x001,LED1x002,LED1x003,LED1x004,LED1x005,LED1x006)
VRDigits(13) = Array(LED1x100,LED1x101,LED1x102,LED1x103,LED1x104,LED1x105,LED1x106)
VRDigits(14) = Array(LED1x200,LED1x201,LED1x202,LED1x203,LED1x204,LED1x205,LED1x206)
VRDigits(15) = Array(LED1x300,LED1x301,LED1x302,LED1x303,LED1x304,LED1x305,LED1x306)
VRDigits(16) = Array(LED1x400,LED1x401,LED1x402,LED1x403,LED1x404,LED1x405,LED1x406)
VRDigits(17) = Array(LED1x500,LED1x501,LED1x502,LED1x503,LED1x504,LED1x505,LED1x506)

VRDigits(18) = Array(LED2x000,LED2x001,LED2x002,LED2x003,LED2x004,LED2x005,LED2x006)
VRDigits(19) = Array(LED2x100,LED2x101,LED2x102,LED2x103,LED2x104,LED2x105,LED2x106)
VRDigits(20) = Array(LED2x200,LED2x201,LED2x202,LED2x203,LED2x204,LED2x205,LED2x206)
VRDigits(21) = Array(LED2x300,LED2x301,LED2x302,LED2x303,LED2x304,LED2x305,LED2x306)
VRDigits(22) = Array(LED2x400,LED2x401,LED2x402,LED2x403,LED2x404,LED2x405,LED2x406)
VRDigits(23) = Array(LED2x500,LED2x501,LED2x502,LED2x503,LED2x504,LED2x505,LED2x506)

VRDigits(24) = Array(LEDax300,LEDax301,LEDax302,LEDax303,LEDax304,LEDax305,LEDax306)
VRDigits(25) = Array(LEDbx400,LEDbx401,LEDbx402,LEDbx403,LEDbx404,LEDbx405,LEDbx406)
VRDigits(26) = Array(LEDcx500,LEDcx501,LEDcx502,LEDcx503,LEDcx504,LEDcx505,LEDcx506)
VRDigits(27) = Array(LEDdx600,LEDdx601,LEDdx602,LEDdx603,LEDdx604,LEDdx605,LEDdx606)
VRDigits(28) = Array(LEDdx700,LEDdx701,LEDdx702,LEDdx703,LEDdx704,LEDdx705,LEDdx706)

dim DisplayColor
DisplayColor =  RGB(255,40,1)

Sub VRBackglassTimer
    Dim ChgLED, ii, jj, num, chg, stat, obj, b, x
    ChgLED=Controller.ChangedLEDs(&Hffffffff, &Hffffffff)
    If Not IsEmpty(ChgLED)Then
       For ii=0 To UBound(chgLED)
          num=chgLED(ii, 0) : chg=chgLED(ii, 1) : stat=chgLED(ii, 2)
              For Each obj In VRDigits(num)
'                   If chg And 1 Then obj.visible=stat And 1    'if you use the object color for off; turn the display object visible to not visible on the playfield, and uncomment this line out.
           If chg And 1 Then FadeDisplay obj, stat And 1
                   chg=chg\2 : stat=stat\2
              Next
        Next
    End If
End Sub

Sub FadeDisplay(object, onoff)
  If OnOff = 1 Then
    object.color = DisplayColor
    Object.Opacity = 12
  Else
    Object.Color = RGB(1,1,1)
    Object.Opacity = 6
  End If
End Sub


Sub InitVRDigits()
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

InitVRDigits



Dim VRBGObj
Sub SetBackglass
' For Each VRBGObj In VRBackglassDigits
'   VRBGObj.x = VRBGObj.x
'   VRBGObj.height = - VRBGObj.y + 758
'   VRBGObj.y = 0 'adjusts the distance from the backglass towards the user
'   VRBGObj.Rotx = -90
' Next

  For Each VRBGObj In VRBGLamps
    VRBGObj.x = VRBGObj.x
    VRBGObj.height = - VRBGObj.y + 236
    VRBGObj.y = 83 'adjusts the distance from the backglass towards the user
    VRBGObj.Rotx = -90
  Next

End Sub

'*******************************
' VR Aro Light
'*******************************
Dim AroOn
Sub AroLight(Enabled)
  If VRMode = True Then
    If AroOn =0 Then
      AroOn=1
      AroF.visible=1
      Torus1.visible=1:BT1Off.visible=0:BT1On.visible=1
      vpmtimer.addtimer 150, "Torus1.visible=0:Torus2.visible=1:BT1On.visible=0:BT1Off.visible=1:BT2Off.visible=0:BT2On.visible=1 '"
      vpmtimer.addtimer 300, "Torus2.visible=0:Torus3.visible=1:BT2On.visible=0:BT2Off.visible=1:BT3Off.visible=0:BT3On.visible=1 '"
      vpmtimer.addtimer 450, "Torus3.visible=0:Torus4.visible=1:BT3On.visible=0:BT3Off.visible=1:BT4Off.visible=0:BT4On.visible=1 '"
      vpmtimer.addtimer 600, "Torus4.visible=0:Torus5.visible=1:BT4On.visible=0:BT4Off.visible=1:BT5Off.visible=0:BT5On.visible=1 '"
      vpmtimer.addtimer 750, "Torus5.visible=0:TapaF.visible=1:BT5On.visible=0:BT5Off.visible=1 '"
      vpmtimer.addtimer 900, "TapaF.visible=0:AroF.visible=0:AroOn=0 '"
    End If
  End If
  If Table1.ShowDT then
    If AroOn =0 Then
      AroOn=1
      aroR.setvalue(0)
      vpmtimer.addtimer 300, "aroR.setvalue(1) '"
      vpmtimer.addtimer 600, "aroR.setvalue(2) '"
      vpmtimer.addtimer 900, "aroR.setvalue(3) '"
      vpmtimer.addtimer 1200, "aroR.setvalue(0):AroOn=0 '"
    End If
  End If
End Sub



'*******************************
' Playmatic 1 dips
' added by Inkochnito
'*******************************

 Sub editDips
 Dim vpmDips:Set vpmDips=New cvpmDips
 With vpmDips
 .AddForm 400,200,"Playmatic 1 - DIP switches"
 .AddFrame 0,0,190,"High score to date award",&H00000001,Array("3 credits",0,"1 credit",&H00000001)'dip 1
 .AddFrame 0,46,190,"Balls per game",&H00000002,Array("3 balls",0,"5 balls",&H00000002)'dip 2
 .AddFrame 0,92,190,"Special award",&H00000004,Array("replay",0,"extra ball",&H00000004)'dip 3
 .AddLabel 0,150,200,40,"After hitting OK, press F3 to reset game with new settings."
 .ViewDips
 End With
 End Sub
 Set vpmShowDips=GetRef("editDips")





' VLM  Arrays - Start
' Arrays per baked part
Dim BP_BB1: BP_BB1=Array(BM_BB1, LM_GI_GI12_BB1, LM_GI_GI13_BB1, LM_GI_GI18_BB1, LM_GI_GI9_BB1, LM_L_LightBumper001_BB1)
Dim BP_BB2_001: BP_BB2_001=Array(BM_BB2_001, LM_GI_GI13_BB2_001, LM_GI_GI14_BB2_001, LM_GI_GI15_BB2_001, LM_GI_GI16_BB2_001, LM_GI_GI17_BB2_001, LM_L_LightBumper002_BB2_001)
Dim BP_BR1: BP_BR1=Array(BM_BR1, LM_GI_GI9_BR1, LM_L_LightBumper001_BR1)
Dim BP_BR2: BP_BR2=Array(BM_BR2, LM_GI_GI16_BR2, LM_GI_GI17_BR2, LM_L_LightBumper002_BR2)
Dim BP_BS1: BP_BS1=Array(BM_BS1, LM_GI_GI12_BS1, LM_GI_GI13_BS1, LM_GI_GI17_BS1, LM_GI_GI18_BS1, LM_GI_GI9_BS1, LM_L_LightBumper001_BS1)
Dim BP_BS2: BP_BS2=Array(BM_BS2, LM_GI_GI13_BS2, LM_GI_GI14_BS2, LM_GI_GI15_BS2, LM_GI_GI16_BS2, LM_GI_GI17_BS2, LM_GI_GI9_BS2, LM_L_LightBumper002_BS2)
Dim BP_Gate2: BP_Gate2=Array(BM_Gate2, LM_GI_GI21_Gate2, LM_GI_GI22_Gate2)
Dim BP_LFlipper: BP_LFlipper=Array(BM_LFlipper, LM_GI_GI2_LFlipper)
Dim BP_LFlipperU: BP_LFlipperU=Array(BM_LFlipperU, LM_GI_GI2_LFlipperU)
Dim BP_LSling1: BP_LSling1=Array(BM_LSling1, LM_GI_GI1_LSling1, LM_GI_GI2_LSling1)
Dim BP_LSling2: BP_LSling2=Array(BM_LSling2, LM_GI_GI1_LSling2, LM_GI_GI2_LSling2)
Dim BP_LSlingArm: BP_LSlingArm=Array(BM_LSlingArm, LM_GI_GI1_LSlingArm, LM_GI_GI2_LSlingArm)
Dim BP_Layer1: BP_Layer1=Array(BM_Layer1, LM_L_A44_Layer1, LM_GI_GI1_Layer1, LM_GI_GI10_Layer1, LM_GI_GI11_Layer1, LM_GI_GI12_Layer1, LM_GI_GI13_Layer1, LM_GI_GI14_Layer1, LM_GI_GI15_Layer1, LM_GI_GI16_Layer1, LM_GI_GI17_Layer1, LM_GI_GI18_Layer1, LM_GI_GI19_Layer1, LM_GI_GI2_Layer1, LM_GI_GI20_Layer1, LM_GI_GI21_Layer1, LM_GI_GI22_Layer1, LM_GI_GI23_Layer1, LM_GI_GI24_Layer1, LM_GI_GI3_Layer1, LM_GI_GI4_Layer1, LM_GI_GI5_Layer1, LM_GI_GI6_Layer1, LM_GI_GI7_Layer1, LM_GI_GI8_Layer1, LM_GI_GI9_Layer1)
Dim BP_Layer_Separators: BP_Layer_Separators=Array(BM_Layer_Separators)
Dim BP_Parts: BP_Parts=Array(BM_Parts, LM_L_A13_Parts, LM_L_A41_Parts, LM_GI_GI1_Parts, LM_GI_GI10_Parts, LM_GI_GI11_Parts, LM_GI_GI12_Parts, LM_GI_GI13_Parts, LM_GI_GI14_Parts, LM_GI_GI15_Parts, LM_GI_GI16_Parts, LM_GI_GI17_Parts, LM_GI_GI18_Parts, LM_GI_GI19_Parts, LM_GI_GI2_Parts, LM_GI_GI20_Parts, LM_GI_GI21_Parts, LM_GI_GI22_Parts, LM_GI_GI23_Parts, LM_GI_GI24_Parts, LM_GI_GI3_Parts, LM_GI_GI4_Parts, LM_GI_GI5_Parts, LM_GI_GI6_Parts, LM_GI_GI7_Parts, LM_GI_GI8_Parts, LM_GI_GI9_Parts, LM_L_LightBumper001_Parts, LM_L_LightBumper002_Parts, LM_L_Shoot_Again_Parts)
Dim BP_PincabRails: BP_PincabRails=Array(BM_PincabRails)
Dim BP_Playfield: BP_Playfield=Array(BM_Playfield, LM_L_A13_Playfield, LM_L_A14_Playfield, LM_L_A15_Playfield, LM_L_A17_Playfield, LM_L_A18_Playfield, LM_L_A19_Playfield, LM_L_A20_Playfield, LM_L_A21_Playfield, LM_L_A22_Playfield, LM_L_A23_Playfield, LM_L_A24_Playfield, LM_L_A25_Playfield, LM_L_A26_Playfield, LM_L_A27_Playfield, LM_L_A28_Playfield, LM_L_A29_Playfield, LM_L_A30_Playfield, LM_L_A31_Playfield, LM_L_A32_Playfield, LM_L_A33_Playfield, LM_L_A34_Playfield, LM_L_A35_Playfield, LM_L_A36_Playfield, LM_L_A37_Playfield, LM_L_A38_Playfield, LM_L_A39_Playfield, LM_L_A40_Playfield, LM_L_A41_Playfield, LM_L_A43_Playfield, LM_L_A44_Playfield, LM_L_A45A_Playfield, LM_L_A45B_Playfield, LM_L_A46A_Playfield, LM_L_A46B_Playfield, LM_GI_GI1_Playfield, LM_GI_GI10_Playfield, LM_GI_GI11_Playfield, LM_GI_GI12_Playfield, LM_GI_GI13_Playfield, LM_GI_GI14_Playfield, LM_GI_GI15_Playfield, LM_GI_GI16_Playfield, LM_GI_GI17_Playfield, LM_GI_GI18_Playfield, LM_GI_GI19_Playfield, LM_GI_GI2_Playfield, LM_GI_GI20_Playfield, _
  LM_GI_GI21_Playfield, LM_GI_GI22_Playfield, LM_GI_GI23_Playfield, LM_GI_GI24_Playfield, LM_GI_GI3_Playfield, LM_GI_GI4_Playfield, LM_GI_GI5_Playfield, LM_GI_GI6_Playfield, LM_GI_GI7_Playfield, LM_GI_GI8_Playfield, LM_GI_GI9_Playfield)
Dim BP_RFlipper: BP_RFlipper=Array(BM_RFlipper, LM_GI_GI4_RFlipper)
Dim BP_RFlipperU: BP_RFlipperU=Array(BM_RFlipperU, LM_GI_GI4_RFlipperU)
Dim BP_RSling1: BP_RSling1=Array(BM_RSling1, LM_GI_GI3_RSling1, LM_GI_GI4_RSling1)
Dim BP_RSling2: BP_RSling2=Array(BM_RSling2, LM_GI_GI3_RSling2, LM_GI_GI4_RSling2)
Dim BP_RSlingArm: BP_RSlingArm=Array(BM_RSlingArm, LM_GI_GI3_RSlingArm, LM_GI_GI4_RSlingArm)
Dim BP_SRod: BP_SRod=Array(BM_SRod)
Dim BP_sw13: BP_sw13=Array(BM_sw13, LM_GI_GI2_sw13)
Dim BP_sw14: BP_sw14=Array(BM_sw14, LM_GI_GI4_sw14)
Dim BP_sw22: BP_sw22=Array(BM_sw22, LM_GI_GI13_sw22)
Dim BP_sw23: BP_sw23=Array(BM_sw23, LM_GI_GI10_sw23, LM_GI_GI11_sw23, LM_GI_GI18_sw23)
Dim BP_sw24: BP_sw24=Array(BM_sw24, LM_GI_GI12_sw24, LM_GI_GI13_sw24, LM_GI_GI14_sw24)
Dim BP_sw25a: BP_sw25a=Array(BM_sw25a)
Dim BP_sw25b: BP_sw25b=Array(BM_sw25b)
Dim BP_sw27: BP_sw27=Array(BM_sw27, LM_GI_GI10_sw27, LM_GI_GI18_sw27)
Dim BP_sw28: BP_sw28=Array(BM_sw28, LM_GI_GI14_sw28, LM_GI_GI15_sw28)
Dim BP_sw32: BP_sw32=Array(BM_sw32, LM_L_A41_sw32, LM_GI_GI17_sw32)
Dim BP_sw33: BP_sw33=Array(BM_sw33, LM_GI_GI3_sw33, LM_GI_GI4_sw33)
Dim BP_sw34: BP_sw34=Array(BM_sw34, LM_GI_GI1_sw34, LM_GI_GI2_sw34)
Dim BP_sw35: BP_sw35=Array(BM_sw35, LM_GI_GI18_sw35, LM_GI_GI9_sw35)
Dim BP_sw36: BP_sw36=Array(BM_sw36, LM_GI_GI12_sw36, LM_GI_GI13_sw36, LM_GI_GI14_sw36)
Dim BP_sw37: BP_sw37=Array(BM_sw37)
Dim BP_sw41: BP_sw41=Array(BM_sw41, LM_GI_GI1_sw41, LM_GI_GI24_sw41, LM_GI_GI7_sw41, LM_GI_GI8_sw41)
Dim BP_sw42: BP_sw42=Array(BM_sw42, LM_GI_GI17_sw42, LM_GI_GI23_sw42, LM_GI_GI5_sw42, LM_GI_GI6_sw42)
Dim BP_sw43: BP_sw43=Array(BM_sw43, LM_GI_GI1_sw43, LM_GI_GI2_sw43, LM_GI_GI24_sw43, LM_GI_GI7_sw43, LM_GI_GI8_sw43)
Dim BP_sw44: BP_sw44=Array(BM_sw44, LM_GI_GI23_sw44, LM_GI_GI3_sw44, LM_GI_GI5_sw44, LM_GI_GI6_sw44)
Dim BP_sw47: BP_sw47=Array(BM_sw47, LM_GI_GI15_sw47, LM_GI_GI16_sw47)
Dim BP_underPF: BP_underPF=Array(BM_underPF, LM_L_A13_underPF, LM_L_A14_underPF, LM_L_A15_underPF, LM_L_A17_underPF, LM_L_A18_underPF, LM_L_A19_underPF, LM_L_A20_underPF, LM_L_A21_underPF, LM_L_A22_underPF, LM_L_A23_underPF, LM_L_A24_underPF, LM_L_A25_underPF, LM_L_A26_underPF, LM_L_A27_underPF, LM_L_A28_underPF, LM_L_A29_underPF, LM_L_A30_underPF, LM_L_A31_underPF, LM_L_A32_underPF, LM_L_A33_underPF, LM_L_A34_underPF, LM_L_A35_underPF, LM_L_A36_underPF, LM_L_A37_underPF, LM_L_A38_underPF, LM_L_A39_underPF, LM_L_A40_underPF, LM_L_A41_underPF, LM_L_A43_underPF, LM_L_A44_underPF, LM_L_A45A_underPF, LM_L_A45B_underPF, LM_L_A46A_underPF, LM_L_A46B_underPF, LM_GI_GI1_underPF, LM_GI_GI10_underPF, LM_GI_GI11_underPF, LM_GI_GI12_underPF, LM_GI_GI13_underPF, LM_GI_GI14_underPF, LM_GI_GI15_underPF, LM_GI_GI16_underPF, LM_GI_GI17_underPF, LM_GI_GI18_underPF, LM_GI_GI19_underPF, LM_GI_GI2_underPF, LM_GI_GI20_underPF, LM_GI_GI21_underPF, LM_GI_GI22_underPF, LM_GI_GI24_underPF, LM_GI_GI3_underPF, LM_GI_GI4_underPF, _
  LM_GI_GI5_underPF, LM_GI_GI6_underPF, LM_GI_GI7_underPF, LM_GI_GI8_underPF, LM_GI_GI9_underPF, LM_L_Shoot_Again_underPF)
' Arrays per lighting scenario
Dim BL_GI_GI1: BL_GI_GI1=Array(LM_GI_GI1_LSling1, LM_GI_GI1_LSling2, LM_GI_GI1_LSlingArm, LM_GI_GI1_Layer1, LM_GI_GI1_Parts, LM_GI_GI1_Playfield, LM_GI_GI1_sw34, LM_GI_GI1_sw41, LM_GI_GI1_sw43, LM_GI_GI1_underPF)
Dim BL_GI_GI10: BL_GI_GI10=Array(LM_GI_GI10_Layer1, LM_GI_GI10_Parts, LM_GI_GI10_Playfield, LM_GI_GI10_sw23, LM_GI_GI10_sw27, LM_GI_GI10_underPF)
Dim BL_GI_GI11: BL_GI_GI11=Array(LM_GI_GI11_Layer1, LM_GI_GI11_Parts, LM_GI_GI11_Playfield, LM_GI_GI11_sw23, LM_GI_GI11_underPF)
Dim BL_GI_GI12: BL_GI_GI12=Array(LM_GI_GI12_BB1, LM_GI_GI12_BS1, LM_GI_GI12_Layer1, LM_GI_GI12_Parts, LM_GI_GI12_Playfield, LM_GI_GI12_sw24, LM_GI_GI12_sw36, LM_GI_GI12_underPF)
Dim BL_GI_GI13: BL_GI_GI13=Array(LM_GI_GI13_BB1, LM_GI_GI13_BB2_001, LM_GI_GI13_BS1, LM_GI_GI13_BS2, LM_GI_GI13_Layer1, LM_GI_GI13_Parts, LM_GI_GI13_Playfield, LM_GI_GI13_sw22, LM_GI_GI13_sw24, LM_GI_GI13_sw36, LM_GI_GI13_underPF)
Dim BL_GI_GI14: BL_GI_GI14=Array(LM_GI_GI14_BB2_001, LM_GI_GI14_BS2, LM_GI_GI14_Layer1, LM_GI_GI14_Parts, LM_GI_GI14_Playfield, LM_GI_GI14_sw24, LM_GI_GI14_sw28, LM_GI_GI14_sw36, LM_GI_GI14_underPF)
Dim BL_GI_GI15: BL_GI_GI15=Array(LM_GI_GI15_BB2_001, LM_GI_GI15_BS2, LM_GI_GI15_Layer1, LM_GI_GI15_Parts, LM_GI_GI15_Playfield, LM_GI_GI15_sw28, LM_GI_GI15_sw47, LM_GI_GI15_underPF)
Dim BL_GI_GI16: BL_GI_GI16=Array(LM_GI_GI16_BB2_001, LM_GI_GI16_BR2, LM_GI_GI16_BS2, LM_GI_GI16_Layer1, LM_GI_GI16_Parts, LM_GI_GI16_Playfield, LM_GI_GI16_sw47, LM_GI_GI16_underPF)
Dim BL_GI_GI17: BL_GI_GI17=Array(LM_GI_GI17_BB2_001, LM_GI_GI17_BR2, LM_GI_GI17_BS1, LM_GI_GI17_BS2, LM_GI_GI17_Layer1, LM_GI_GI17_Parts, LM_GI_GI17_Playfield, LM_GI_GI17_sw32, LM_GI_GI17_sw42, LM_GI_GI17_underPF)
Dim BL_GI_GI18: BL_GI_GI18=Array(LM_GI_GI18_BB1, LM_GI_GI18_BS1, LM_GI_GI18_Layer1, LM_GI_GI18_Parts, LM_GI_GI18_Playfield, LM_GI_GI18_sw23, LM_GI_GI18_sw27, LM_GI_GI18_sw35, LM_GI_GI18_underPF)
Dim BL_GI_GI19: BL_GI_GI19=Array(LM_GI_GI19_Layer1, LM_GI_GI19_Parts, LM_GI_GI19_Playfield, LM_GI_GI19_underPF)
Dim BL_GI_GI2: BL_GI_GI2=Array(LM_GI_GI2_LFlipper, LM_GI_GI2_LFlipperU, LM_GI_GI2_LSling1, LM_GI_GI2_LSling2, LM_GI_GI2_LSlingArm, LM_GI_GI2_Layer1, LM_GI_GI2_Parts, LM_GI_GI2_Playfield, LM_GI_GI2_sw13, LM_GI_GI2_sw34, LM_GI_GI2_sw43, LM_GI_GI2_underPF)
Dim BL_GI_GI20: BL_GI_GI20=Array(LM_GI_GI20_Layer1, LM_GI_GI20_Parts, LM_GI_GI20_Playfield, LM_GI_GI20_underPF)
Dim BL_GI_GI21: BL_GI_GI21=Array(LM_GI_GI21_Gate2, LM_GI_GI21_Layer1, LM_GI_GI21_Parts, LM_GI_GI21_Playfield, LM_GI_GI21_underPF)
Dim BL_GI_GI22: BL_GI_GI22=Array(LM_GI_GI22_Gate2, LM_GI_GI22_Layer1, LM_GI_GI22_Parts, LM_GI_GI22_Playfield, LM_GI_GI22_underPF)
Dim BL_GI_GI23: BL_GI_GI23=Array(LM_GI_GI23_Layer1, LM_GI_GI23_Parts, LM_GI_GI23_Playfield, LM_GI_GI23_sw42, LM_GI_GI23_sw44)
Dim BL_GI_GI24: BL_GI_GI24=Array(LM_GI_GI24_Layer1, LM_GI_GI24_Parts, LM_GI_GI24_Playfield, LM_GI_GI24_sw41, LM_GI_GI24_sw43, LM_GI_GI24_underPF)
Dim BL_GI_GI3: BL_GI_GI3=Array(LM_GI_GI3_Layer1, LM_GI_GI3_Parts, LM_GI_GI3_Playfield, LM_GI_GI3_RSling1, LM_GI_GI3_RSling2, LM_GI_GI3_RSlingArm, LM_GI_GI3_sw33, LM_GI_GI3_sw44, LM_GI_GI3_underPF)
Dim BL_GI_GI4: BL_GI_GI4=Array(LM_GI_GI4_Layer1, LM_GI_GI4_Parts, LM_GI_GI4_Playfield, LM_GI_GI4_RFlipper, LM_GI_GI4_RFlipperU, LM_GI_GI4_RSling1, LM_GI_GI4_RSling2, LM_GI_GI4_RSlingArm, LM_GI_GI4_sw14, LM_GI_GI4_sw33, LM_GI_GI4_underPF)
Dim BL_GI_GI5: BL_GI_GI5=Array(LM_GI_GI5_Layer1, LM_GI_GI5_Parts, LM_GI_GI5_Playfield, LM_GI_GI5_sw42, LM_GI_GI5_sw44, LM_GI_GI5_underPF)
Dim BL_GI_GI6: BL_GI_GI6=Array(LM_GI_GI6_Layer1, LM_GI_GI6_Parts, LM_GI_GI6_Playfield, LM_GI_GI6_sw42, LM_GI_GI6_sw44, LM_GI_GI6_underPF)
Dim BL_GI_GI7: BL_GI_GI7=Array(LM_GI_GI7_Layer1, LM_GI_GI7_Parts, LM_GI_GI7_Playfield, LM_GI_GI7_sw41, LM_GI_GI7_sw43, LM_GI_GI7_underPF)
Dim BL_GI_GI8: BL_GI_GI8=Array(LM_GI_GI8_Layer1, LM_GI_GI8_Parts, LM_GI_GI8_Playfield, LM_GI_GI8_sw41, LM_GI_GI8_sw43, LM_GI_GI8_underPF)
Dim BL_GI_GI9: BL_GI_GI9=Array(LM_GI_GI9_BB1, LM_GI_GI9_BR1, LM_GI_GI9_BS1, LM_GI_GI9_BS2, LM_GI_GI9_Layer1, LM_GI_GI9_Parts, LM_GI_GI9_Playfield, LM_GI_GI9_sw35, LM_GI_GI9_underPF)
Dim BL_L_A13: BL_L_A13=Array(LM_L_A13_Parts, LM_L_A13_Playfield, LM_L_A13_underPF)
Dim BL_L_A14: BL_L_A14=Array(LM_L_A14_Playfield, LM_L_A14_underPF)
Dim BL_L_A15: BL_L_A15=Array(LM_L_A15_Playfield, LM_L_A15_underPF)
Dim BL_L_A17: BL_L_A17=Array(LM_L_A17_Playfield, LM_L_A17_underPF)
Dim BL_L_A18: BL_L_A18=Array(LM_L_A18_Playfield, LM_L_A18_underPF)
Dim BL_L_A19: BL_L_A19=Array(LM_L_A19_Playfield, LM_L_A19_underPF)
Dim BL_L_A20: BL_L_A20=Array(LM_L_A20_Playfield, LM_L_A20_underPF)
Dim BL_L_A21: BL_L_A21=Array(LM_L_A21_Playfield, LM_L_A21_underPF)
Dim BL_L_A22: BL_L_A22=Array(LM_L_A22_Playfield, LM_L_A22_underPF)
Dim BL_L_A23: BL_L_A23=Array(LM_L_A23_Playfield, LM_L_A23_underPF)
Dim BL_L_A24: BL_L_A24=Array(LM_L_A24_Playfield, LM_L_A24_underPF)
Dim BL_L_A25: BL_L_A25=Array(LM_L_A25_Playfield, LM_L_A25_underPF)
Dim BL_L_A26: BL_L_A26=Array(LM_L_A26_Playfield, LM_L_A26_underPF)
Dim BL_L_A27: BL_L_A27=Array(LM_L_A27_Playfield, LM_L_A27_underPF)
Dim BL_L_A28: BL_L_A28=Array(LM_L_A28_Playfield, LM_L_A28_underPF)
Dim BL_L_A29: BL_L_A29=Array(LM_L_A29_Playfield, LM_L_A29_underPF)
Dim BL_L_A30: BL_L_A30=Array(LM_L_A30_Playfield, LM_L_A30_underPF)
Dim BL_L_A31: BL_L_A31=Array(LM_L_A31_Playfield, LM_L_A31_underPF)
Dim BL_L_A32: BL_L_A32=Array(LM_L_A32_Playfield, LM_L_A32_underPF)
Dim BL_L_A33: BL_L_A33=Array(LM_L_A33_Playfield, LM_L_A33_underPF)
Dim BL_L_A34: BL_L_A34=Array(LM_L_A34_Playfield, LM_L_A34_underPF)
Dim BL_L_A35: BL_L_A35=Array(LM_L_A35_Playfield, LM_L_A35_underPF)
Dim BL_L_A36: BL_L_A36=Array(LM_L_A36_Playfield, LM_L_A36_underPF)
Dim BL_L_A37: BL_L_A37=Array(LM_L_A37_Playfield, LM_L_A37_underPF)
Dim BL_L_A38: BL_L_A38=Array(LM_L_A38_Playfield, LM_L_A38_underPF)
Dim BL_L_A39: BL_L_A39=Array(LM_L_A39_Playfield, LM_L_A39_underPF)
Dim BL_L_A40: BL_L_A40=Array(LM_L_A40_Playfield, LM_L_A40_underPF)
Dim BL_L_A41: BL_L_A41=Array(LM_L_A41_Parts, LM_L_A41_Playfield, LM_L_A41_sw32, LM_L_A41_underPF)
Dim BL_L_A43: BL_L_A43=Array(LM_L_A43_Playfield, LM_L_A43_underPF)
Dim BL_L_A44: BL_L_A44=Array(LM_L_A44_Layer1, LM_L_A44_Playfield, LM_L_A44_underPF)
Dim BL_L_A45A: BL_L_A45A=Array(LM_L_A45A_Playfield, LM_L_A45A_underPF)
Dim BL_L_A45B: BL_L_A45B=Array(LM_L_A45B_Playfield, LM_L_A45B_underPF)
Dim BL_L_A46A: BL_L_A46A=Array(LM_L_A46A_Playfield, LM_L_A46A_underPF)
Dim BL_L_A46B: BL_L_A46B=Array(LM_L_A46B_Playfield, LM_L_A46B_underPF)
Dim BL_L_LightBumper001: BL_L_LightBumper001=Array(LM_L_LightBumper001_BB1, LM_L_LightBumper001_BR1, LM_L_LightBumper001_BS1, LM_L_LightBumper001_Parts)
Dim BL_L_LightBumper002: BL_L_LightBumper002=Array(LM_L_LightBumper002_BB2_001, LM_L_LightBumper002_BR2, LM_L_LightBumper002_BS2, LM_L_LightBumper002_Parts)
Dim BL_L_Shoot_Again: BL_L_Shoot_Again=Array(LM_L_Shoot_Again_Parts, LM_L_Shoot_Again_underPF)
Dim BL_World: BL_World=Array(BM_BB1, BM_BB2_001, BM_BR1, BM_BR2, BM_BS1, BM_BS2, BM_Gate2, BM_LFlipper, BM_LFlipperU, BM_LSling1, BM_LSling2, BM_LSlingArm, BM_Layer_Separators, BM_Layer1, BM_Parts, BM_PincabRails, BM_Playfield, BM_RFlipper, BM_RFlipperU, BM_RSling1, BM_RSling2, BM_RSlingArm, BM_SRod, BM_sw13, BM_sw14, BM_sw22, BM_sw23, BM_sw24, BM_sw25a, BM_sw25b, BM_sw27, BM_sw28, BM_sw32, BM_sw33, BM_sw34, BM_sw35, BM_sw36, BM_sw37, BM_sw41, BM_sw42, BM_sw43, BM_sw44, BM_sw47, BM_underPF)
' Global arrays
Dim BG_Bakemap: BG_Bakemap=Array(BM_BB1, BM_BB2_001, BM_BR1, BM_BR2, BM_BS1, BM_BS2, BM_Gate2, BM_LFlipper, BM_LFlipperU, BM_LSling1, BM_LSling2, BM_LSlingArm, BM_Layer_Separators, BM_Layer1, BM_Parts, BM_PincabRails, BM_Playfield, BM_RFlipper, BM_RFlipperU, BM_RSling1, BM_RSling2, BM_RSlingArm, BM_SRod, BM_sw13, BM_sw14, BM_sw22, BM_sw23, BM_sw24, BM_sw25a, BM_sw25b, BM_sw27, BM_sw28, BM_sw32, BM_sw33, BM_sw34, BM_sw35, BM_sw36, BM_sw37, BM_sw41, BM_sw42, BM_sw43, BM_sw44, BM_sw47, BM_underPF)
Dim BG_Lightmap: BG_Lightmap=Array(LM_GI_GI1_LSling1, LM_GI_GI1_LSling2, LM_GI_GI1_LSlingArm, LM_GI_GI1_Layer1, LM_GI_GI1_Parts, LM_GI_GI1_Playfield, LM_GI_GI1_sw34, LM_GI_GI1_sw41, LM_GI_GI1_sw43, LM_GI_GI1_underPF, LM_GI_GI10_Layer1, LM_GI_GI10_Parts, LM_GI_GI10_Playfield, LM_GI_GI10_sw23, LM_GI_GI10_sw27, LM_GI_GI10_underPF, LM_GI_GI11_Layer1, LM_GI_GI11_Parts, LM_GI_GI11_Playfield, LM_GI_GI11_sw23, LM_GI_GI11_underPF, LM_GI_GI12_BB1, LM_GI_GI12_BS1, LM_GI_GI12_Layer1, LM_GI_GI12_Parts, LM_GI_GI12_Playfield, LM_GI_GI12_sw24, LM_GI_GI12_sw36, LM_GI_GI12_underPF, LM_GI_GI13_BB1, LM_GI_GI13_BB2_001, LM_GI_GI13_BS1, LM_GI_GI13_BS2, LM_GI_GI13_Layer1, LM_GI_GI13_Parts, LM_GI_GI13_Playfield, LM_GI_GI13_sw22, LM_GI_GI13_sw24, LM_GI_GI13_sw36, LM_GI_GI13_underPF, LM_GI_GI14_BB2_001, LM_GI_GI14_BS2, LM_GI_GI14_Layer1, LM_GI_GI14_Parts, LM_GI_GI14_Playfield, LM_GI_GI14_sw24, LM_GI_GI14_sw28, LM_GI_GI14_sw36, LM_GI_GI14_underPF, LM_GI_GI15_BB2_001, LM_GI_GI15_BS2, LM_GI_GI15_Layer1, LM_GI_GI15_Parts, _
  LM_GI_GI15_Playfield, LM_GI_GI15_sw28, LM_GI_GI15_sw47, LM_GI_GI15_underPF, LM_GI_GI16_BB2_001, LM_GI_GI16_BR2, LM_GI_GI16_BS2, LM_GI_GI16_Layer1, LM_GI_GI16_Parts, LM_GI_GI16_Playfield, LM_GI_GI16_sw47, LM_GI_GI16_underPF, LM_GI_GI17_BB2_001, LM_GI_GI17_BR2, LM_GI_GI17_BS1, LM_GI_GI17_BS2, LM_GI_GI17_Layer1, LM_GI_GI17_Parts, LM_GI_GI17_Playfield, LM_GI_GI17_sw32, LM_GI_GI17_sw42, LM_GI_GI17_underPF, LM_GI_GI18_BB1, LM_GI_GI18_BS1, LM_GI_GI18_Layer1, LM_GI_GI18_Parts, LM_GI_GI18_Playfield, LM_GI_GI18_sw23, LM_GI_GI18_sw27, LM_GI_GI18_sw35, LM_GI_GI18_underPF, LM_GI_GI19_Layer1, LM_GI_GI19_Parts, LM_GI_GI19_Playfield, LM_GI_GI19_underPF, LM_GI_GI2_LFlipper, LM_GI_GI2_LFlipperU, LM_GI_GI2_LSling1, LM_GI_GI2_LSling2, LM_GI_GI2_LSlingArm, LM_GI_GI2_Layer1, LM_GI_GI2_Parts, LM_GI_GI2_Playfield, LM_GI_GI2_sw13, LM_GI_GI2_sw34, LM_GI_GI2_sw43, LM_GI_GI2_underPF, LM_GI_GI20_Layer1, LM_GI_GI20_Parts, LM_GI_GI20_Playfield, LM_GI_GI20_underPF, LM_GI_GI21_Gate2, LM_GI_GI21_Layer1, LM_GI_GI21_Parts, LM_GI_GI21_Playfield, _
  LM_GI_GI21_underPF, LM_GI_GI22_Gate2, LM_GI_GI22_Layer1, LM_GI_GI22_Parts, LM_GI_GI22_Playfield, LM_GI_GI22_underPF, LM_GI_GI23_Layer1, LM_GI_GI23_Parts, LM_GI_GI23_Playfield, LM_GI_GI23_sw42, LM_GI_GI23_sw44, LM_GI_GI24_Layer1, LM_GI_GI24_Parts, LM_GI_GI24_Playfield, LM_GI_GI24_sw41, LM_GI_GI24_sw43, LM_GI_GI24_underPF, LM_GI_GI3_Layer1, LM_GI_GI3_Parts, LM_GI_GI3_Playfield, LM_GI_GI3_RSling1, LM_GI_GI3_RSling2, LM_GI_GI3_RSlingArm, LM_GI_GI3_sw33, LM_GI_GI3_sw44, LM_GI_GI3_underPF, LM_GI_GI4_Layer1, LM_GI_GI4_Parts, LM_GI_GI4_Playfield, LM_GI_GI4_RFlipper, LM_GI_GI4_RFlipperU, LM_GI_GI4_RSling1, LM_GI_GI4_RSling2, LM_GI_GI4_RSlingArm, LM_GI_GI4_sw14, LM_GI_GI4_sw33, LM_GI_GI4_underPF, LM_GI_GI5_Layer1, LM_GI_GI5_Parts, LM_GI_GI5_Playfield, LM_GI_GI5_sw42, LM_GI_GI5_sw44, LM_GI_GI5_underPF, LM_GI_GI6_Layer1, LM_GI_GI6_Parts, LM_GI_GI6_Playfield, LM_GI_GI6_sw42, LM_GI_GI6_sw44, LM_GI_GI6_underPF, LM_GI_GI7_Layer1, LM_GI_GI7_Parts, LM_GI_GI7_Playfield, LM_GI_GI7_sw41, LM_GI_GI7_sw43, LM_GI_GI7_underPF, _
  LM_GI_GI8_Layer1, LM_GI_GI8_Parts, LM_GI_GI8_Playfield, LM_GI_GI8_sw41, LM_GI_GI8_sw43, LM_GI_GI8_underPF, LM_GI_GI9_BB1, LM_GI_GI9_BR1, LM_GI_GI9_BS1, LM_GI_GI9_BS2, LM_GI_GI9_Layer1, LM_GI_GI9_Parts, LM_GI_GI9_Playfield, LM_GI_GI9_sw35, LM_GI_GI9_underPF, LM_L_A13_Parts, LM_L_A13_Playfield, LM_L_A13_underPF, LM_L_A14_Playfield, LM_L_A14_underPF, LM_L_A15_Playfield, LM_L_A15_underPF, LM_L_A17_Playfield, LM_L_A17_underPF, LM_L_A18_Playfield, LM_L_A18_underPF, LM_L_A19_Playfield, LM_L_A19_underPF, LM_L_A20_Playfield, LM_L_A20_underPF, LM_L_A21_Playfield, LM_L_A21_underPF, LM_L_A22_Playfield, LM_L_A22_underPF, LM_L_A23_Playfield, LM_L_A23_underPF, LM_L_A24_Playfield, LM_L_A24_underPF, LM_L_A25_Playfield, LM_L_A25_underPF, LM_L_A26_Playfield, LM_L_A26_underPF, LM_L_A27_Playfield, LM_L_A27_underPF, LM_L_A28_Playfield, LM_L_A28_underPF, LM_L_A29_Playfield, LM_L_A29_underPF, LM_L_A30_Playfield, LM_L_A30_underPF, LM_L_A31_Playfield, LM_L_A31_underPF, LM_L_A32_Playfield, LM_L_A32_underPF, LM_L_A33_Playfield, _
  LM_L_A33_underPF, LM_L_A34_Playfield, LM_L_A34_underPF, LM_L_A35_Playfield, LM_L_A35_underPF, LM_L_A36_Playfield, LM_L_A36_underPF, LM_L_A37_Playfield, LM_L_A37_underPF, LM_L_A38_Playfield, LM_L_A38_underPF, LM_L_A39_Playfield, LM_L_A39_underPF, LM_L_A40_Playfield, LM_L_A40_underPF, LM_L_A41_Parts, LM_L_A41_Playfield, LM_L_A41_sw32, LM_L_A41_underPF, LM_L_A43_Playfield, LM_L_A43_underPF, LM_L_A44_Layer1, LM_L_A44_Playfield, LM_L_A44_underPF, LM_L_A45A_Playfield, LM_L_A45A_underPF, LM_L_A45B_Playfield, LM_L_A45B_underPF, LM_L_A46A_Playfield, LM_L_A46A_underPF, LM_L_A46B_Playfield, LM_L_A46B_underPF, LM_L_LightBumper001_BB1, LM_L_LightBumper001_BR1, LM_L_LightBumper001_BS1, LM_L_LightBumper001_Parts, LM_L_LightBumper002_BB2_001, LM_L_LightBumper002_BR2, LM_L_LightBumper002_BS2, LM_L_LightBumper002_Parts, LM_L_Shoot_Again_Parts, LM_L_Shoot_Again_underPF)
Dim BG_All: BG_All=Array(BM_BB1, BM_BB2_001, BM_BR1, BM_BR2, BM_BS1, BM_BS2, BM_Gate2, BM_LFlipper, BM_LFlipperU, BM_LSling1, BM_LSling2, BM_LSlingArm, BM_Layer_Separators, BM_Layer1, BM_Parts, BM_PincabRails, BM_Playfield, BM_RFlipper, BM_RFlipperU, BM_RSling1, BM_RSling2, BM_RSlingArm, BM_SRod, BM_sw13, BM_sw14, BM_sw22, BM_sw23, BM_sw24, BM_sw25a, BM_sw25b, BM_sw27, BM_sw28, BM_sw32, BM_sw33, BM_sw34, BM_sw35, BM_sw36, BM_sw37, BM_sw41, BM_sw42, BM_sw43, BM_sw44, BM_sw47, BM_underPF, LM_GI_GI1_LSling1, LM_GI_GI1_LSling2, LM_GI_GI1_LSlingArm, LM_GI_GI1_Layer1, LM_GI_GI1_Parts, LM_GI_GI1_Playfield, LM_GI_GI1_sw34, LM_GI_GI1_sw41, LM_GI_GI1_sw43, LM_GI_GI1_underPF, LM_GI_GI10_Layer1, LM_GI_GI10_Parts, LM_GI_GI10_Playfield, LM_GI_GI10_sw23, LM_GI_GI10_sw27, LM_GI_GI10_underPF, LM_GI_GI11_Layer1, LM_GI_GI11_Parts, LM_GI_GI11_Playfield, LM_GI_GI11_sw23, LM_GI_GI11_underPF, LM_GI_GI12_BB1, LM_GI_GI12_BS1, LM_GI_GI12_Layer1, LM_GI_GI12_Parts, LM_GI_GI12_Playfield, LM_GI_GI12_sw24, LM_GI_GI12_sw36, _
  LM_GI_GI12_underPF, LM_GI_GI13_BB1, LM_GI_GI13_BB2_001, LM_GI_GI13_BS1, LM_GI_GI13_BS2, LM_GI_GI13_Layer1, LM_GI_GI13_Parts, LM_GI_GI13_Playfield, LM_GI_GI13_sw22, LM_GI_GI13_sw24, LM_GI_GI13_sw36, LM_GI_GI13_underPF, LM_GI_GI14_BB2_001, LM_GI_GI14_BS2, LM_GI_GI14_Layer1, LM_GI_GI14_Parts, LM_GI_GI14_Playfield, LM_GI_GI14_sw24, LM_GI_GI14_sw28, LM_GI_GI14_sw36, LM_GI_GI14_underPF, LM_GI_GI15_BB2_001, LM_GI_GI15_BS2, LM_GI_GI15_Layer1, LM_GI_GI15_Parts, LM_GI_GI15_Playfield, LM_GI_GI15_sw28, LM_GI_GI15_sw47, LM_GI_GI15_underPF, LM_GI_GI16_BB2_001, LM_GI_GI16_BR2, LM_GI_GI16_BS2, LM_GI_GI16_Layer1, LM_GI_GI16_Parts, LM_GI_GI16_Playfield, LM_GI_GI16_sw47, LM_GI_GI16_underPF, LM_GI_GI17_BB2_001, LM_GI_GI17_BR2, LM_GI_GI17_BS1, LM_GI_GI17_BS2, LM_GI_GI17_Layer1, LM_GI_GI17_Parts, LM_GI_GI17_Playfield, LM_GI_GI17_sw32, LM_GI_GI17_sw42, LM_GI_GI17_underPF, LM_GI_GI18_BB1, LM_GI_GI18_BS1, LM_GI_GI18_Layer1, LM_GI_GI18_Parts, LM_GI_GI18_Playfield, LM_GI_GI18_sw23, LM_GI_GI18_sw27, LM_GI_GI18_sw35, LM_GI_GI18_underPF, _
  LM_GI_GI19_Layer1, LM_GI_GI19_Parts, LM_GI_GI19_Playfield, LM_GI_GI19_underPF, LM_GI_GI2_LFlipper, LM_GI_GI2_LFlipperU, LM_GI_GI2_LSling1, LM_GI_GI2_LSling2, LM_GI_GI2_LSlingArm, LM_GI_GI2_Layer1, LM_GI_GI2_Parts, LM_GI_GI2_Playfield, LM_GI_GI2_sw13, LM_GI_GI2_sw34, LM_GI_GI2_sw43, LM_GI_GI2_underPF, LM_GI_GI20_Layer1, LM_GI_GI20_Parts, LM_GI_GI20_Playfield, LM_GI_GI20_underPF, LM_GI_GI21_Gate2, LM_GI_GI21_Layer1, LM_GI_GI21_Parts, LM_GI_GI21_Playfield, LM_GI_GI21_underPF, LM_GI_GI22_Gate2, LM_GI_GI22_Layer1, LM_GI_GI22_Parts, LM_GI_GI22_Playfield, LM_GI_GI22_underPF, LM_GI_GI23_Layer1, LM_GI_GI23_Parts, LM_GI_GI23_Playfield, LM_GI_GI23_sw42, LM_GI_GI23_sw44, LM_GI_GI24_Layer1, LM_GI_GI24_Parts, LM_GI_GI24_Playfield, LM_GI_GI24_sw41, LM_GI_GI24_sw43, LM_GI_GI24_underPF, LM_GI_GI3_Layer1, LM_GI_GI3_Parts, LM_GI_GI3_Playfield, LM_GI_GI3_RSling1, LM_GI_GI3_RSling2, LM_GI_GI3_RSlingArm, LM_GI_GI3_sw33, LM_GI_GI3_sw44, LM_GI_GI3_underPF, LM_GI_GI4_Layer1, LM_GI_GI4_Parts, LM_GI_GI4_Playfield, LM_GI_GI4_RFlipper, _
  LM_GI_GI4_RFlipperU, LM_GI_GI4_RSling1, LM_GI_GI4_RSling2, LM_GI_GI4_RSlingArm, LM_GI_GI4_sw14, LM_GI_GI4_sw33, LM_GI_GI4_underPF, LM_GI_GI5_Layer1, LM_GI_GI5_Parts, LM_GI_GI5_Playfield, LM_GI_GI5_sw42, LM_GI_GI5_sw44, LM_GI_GI5_underPF, LM_GI_GI6_Layer1, LM_GI_GI6_Parts, LM_GI_GI6_Playfield, LM_GI_GI6_sw42, LM_GI_GI6_sw44, LM_GI_GI6_underPF, LM_GI_GI7_Layer1, LM_GI_GI7_Parts, LM_GI_GI7_Playfield, LM_GI_GI7_sw41, LM_GI_GI7_sw43, LM_GI_GI7_underPF, LM_GI_GI8_Layer1, LM_GI_GI8_Parts, LM_GI_GI8_Playfield, LM_GI_GI8_sw41, LM_GI_GI8_sw43, LM_GI_GI8_underPF, LM_GI_GI9_BB1, LM_GI_GI9_BR1, LM_GI_GI9_BS1, LM_GI_GI9_BS2, LM_GI_GI9_Layer1, LM_GI_GI9_Parts, LM_GI_GI9_Playfield, LM_GI_GI9_sw35, LM_GI_GI9_underPF, LM_L_A13_Parts, LM_L_A13_Playfield, LM_L_A13_underPF, LM_L_A14_Playfield, LM_L_A14_underPF, LM_L_A15_Playfield, LM_L_A15_underPF, LM_L_A17_Playfield, LM_L_A17_underPF, LM_L_A18_Playfield, LM_L_A18_underPF, LM_L_A19_Playfield, LM_L_A19_underPF, LM_L_A20_Playfield, LM_L_A20_underPF, LM_L_A21_Playfield, _
  LM_L_A21_underPF, LM_L_A22_Playfield, LM_L_A22_underPF, LM_L_A23_Playfield, LM_L_A23_underPF, LM_L_A24_Playfield, LM_L_A24_underPF, LM_L_A25_Playfield, LM_L_A25_underPF, LM_L_A26_Playfield, LM_L_A26_underPF, LM_L_A27_Playfield, LM_L_A27_underPF, LM_L_A28_Playfield, LM_L_A28_underPF, LM_L_A29_Playfield, LM_L_A29_underPF, LM_L_A30_Playfield, LM_L_A30_underPF, LM_L_A31_Playfield, LM_L_A31_underPF, LM_L_A32_Playfield, LM_L_A32_underPF, LM_L_A33_Playfield, LM_L_A33_underPF, LM_L_A34_Playfield, LM_L_A34_underPF, LM_L_A35_Playfield, LM_L_A35_underPF, LM_L_A36_Playfield, LM_L_A36_underPF, LM_L_A37_Playfield, LM_L_A37_underPF, LM_L_A38_Playfield, LM_L_A38_underPF, LM_L_A39_Playfield, LM_L_A39_underPF, LM_L_A40_Playfield, LM_L_A40_underPF, LM_L_A41_Parts, LM_L_A41_Playfield, LM_L_A41_sw32, LM_L_A41_underPF, LM_L_A43_Playfield, LM_L_A43_underPF, LM_L_A44_Layer1, LM_L_A44_Playfield, LM_L_A44_underPF, LM_L_A45A_Playfield, LM_L_A45A_underPF, LM_L_A45B_Playfield, LM_L_A45B_underPF, LM_L_A46A_Playfield, LM_L_A46A_underPF, _
  LM_L_A46B_Playfield, LM_L_A46B_underPF, LM_L_LightBumper001_BB1, LM_L_LightBumper001_BR1, LM_L_LightBumper001_BS1, LM_L_LightBumper001_Parts, LM_L_LightBumper002_BB2_001, LM_L_LightBumper002_BR2, LM_L_LightBumper002_BS2, LM_L_LightBumper002_Parts, LM_L_Shoot_Again_Parts, LM_L_Shoot_Again_underPF)
' VLM  Arrays - End

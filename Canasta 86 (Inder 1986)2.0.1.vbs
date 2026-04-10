'
'          __oooo
'         //  /    o
'       / /  |      o
'      | /  |__     o
'     | |   __/    o
'     | |  /  oooo
'   | /  |  _______________________
'  | |   | |_______________________|
' ||    |___ \__/\_/\_/\/\/\/\/\/\/
' ||    |    \ \/\/\/\/\/\/\/\/\/
' ||    \     \ \/\/\/\/\/\/\/\/
' | |    |    |  \/\/\/\/\/\/\/
'  | |    |   |   \/\/\/\/\/\/
'   | |    \_/    |\/\/\/\/\/|
'    \ \     \__  |/\/\/\/\/\|
'     |         \
'    |           |
'     |           |
'     |           |
'     |           |
'      |          |
'      |         |
'       |________|        By NESTORGIAN
'       |________|
'      ||         |          5/6/2024
'     | |          |
'    |  /          |
'  |  /          /
'  /\|__________|
' |   /       /
' |  |       /
' | |       |
'  \ |      \
'    \ \      \
'      \ \     |
'        \ \    |
'          \ \  |_
'            \\_/ \
'            |    |
'            |   /
'            |  |
'            |_/
'
' Inder's Canasta 86 / IPD No. 4097 / 1986 / 4 Players
'                 - Nestorgian -
'          Initial version by JPSalas 2018

Option Explicit
Randomize


On Error Resume Next
ExecuteGlobal GetTextFile("controller.vbs")
If Err Then MsgBox "You need the controller.vbs in order to run this table, available in the vp10 package"
On Error Goto 0



'*******************************************
'  Constants and Global Variables
'*******************************************

'----- Shadow Options -----
Const DynamicBallShadowsOn = 1    '0 = no dynamic ball shadow ("triangles" near slings and such), 1 = enable dynamic ball shadow
Const AmbientBallShadowOn = 1   '0 = Static shadow under ball ("flasher" image, like JP's)
                  '1 = Moving ball shadow ("primitive" object, like ninuzzu's) - This is the only one that behaves like a true shadow!
                  '2 = flasher image shadow, but it moves like ninuzzu's

Const UsingROM = True       'The UsingROM flag is to indicate code that requires ROM usage. Mostly for instructional purposes only.

Const BallSize = 50         'Ball size must be 50
Const BallMass = 1          'Ball mass must be 1
Const tnob = 1            'Total number of balls
Const lob = 0           'Locked balls
Const cGameName = "canasta"

Dim tablewidth: tablewidth = Table1.width
Dim tableheight: tableheight = Table1.height
Dim bsTrough, dtBankL, dtBankR, x

'  Standard definitions
Const UseSolenoids = 2
Const UseLamps = 0
Const UseGI = 0
Const UseSync = 0 'set it to 1 if the table runs too fast
Const HandleMech = 0

'Internal DMD in Desktop Mode, using a textbox (must be called before LoadVPM)
Dim VR, UseVPMDMD, CabinetMode

LoadVPM "01550000", "inder.vbs", 3.26

Dim VRMode, DesktopMode: DesktopMode = Table1.ShowDT
Const ForceVR = False   'Forces VR room to be active


' Standard Sounds
Const SSolenoidOn = "fx_Solenoidon"
Const SSolenoidOff = "fx_Solenoidoff"
Const SCoin = "fx_Coin"


'*******************************************
'  Table Initialization and Exiting
'*******************************************

 Dim CaBall, BOT

Sub table1_Init
    vpmInit me
    With Controller
        .GameName = cGameName
        If Err Then MsgBox "Can't start Game" & cGameName & vbNewLine & Err.Description:Exit Sub
        .SplashInfoLine = "Canasta 86 - Inder 1986" & vbNewLine & "VPX table by NestorGian"
        .HandleKeyboard = 0
        .ShowTitle = 0
        .ShowDMDOnly = 1
        .ShowFrame = 0
        .HandleMechanics = 0
'        .Hidden = VarHidden
        .Games(cGameName).Settings.Value("rol") = 0   '1= rotated display, 0= normal
        .Games(cGameName).Settings.Value("sound") = 1 '1 enabled rom sound
        '.SetDisplayPosition 0,0, GetPlayerHWnd 'restore dmd window position
        On Error Resume Next
        Controller.SolMask(0) = 0
        vpmTimer.AddTimer 2000, "Controller.SolMask(0)=&Hffffffff'" 'ignore all solenoids - then add the Timer to renable all the solenoids after 2 seconds
        Controller.Run GetPlayerHWnd
        On Error Goto 0
    End With

    ' Nudging
    vpmNudge.TiltSwitch = 53
    vpmNudge.Sensitivity = 3
    vpmNudge.TiltObj = Array(Bumper1, Bumper2, Bumper3)

  Set CaBall = Drain.CreateSizedballWithMass(Ballsize/2,Ballmass)
  Controller.Switch(91) = 1

    ' Main Timer init
    PinMAMETimer.Interval = PinMAMEInterval
    PinMAMETimer.Enabled = 1
  GameTimer.Enabled = 1
  SetBackglass
End Sub


'*******************************************
'  ZOPT: User Options
'*******************************************

Dim LightLevel : LightLevel = 0.25        ' Level of room lighting (0 to 1), where 0 is dark and 100 is brightest
Dim ColorLUT : ColorLUT = 1           ' Color desaturation LUTs: 1 to 11, where 1 is normal and 11 is black'n'white
Dim VolumeDial : VolumeDial = 0.8             ' Overall Mechanical sound effect volume. Recommended values should be no greater than 1.
Dim BallRollVolume : BallRollVolume = 0.5     ' Level of ball rolling volume. Value between 0 and 1
Dim RampRollVolume : RampRollVolume = 0.5     ' Level of ramp rolling volume. Value between 0 and 1
Dim StagedFlippers : StagedFlippers = 0         ' Staged Flippers. 0 = Disabled, 1 = Enabled





'*******************************************
'  Timers
'*******************************************

Sub GameTimer_Timer()         ' The game timer interval; should be 10 ms
    Cor.Update            ' update ball tracking (this sometimes goes in the RDampen_Timer sub)
End Sub

Sub FrameTimer_Timer()
  BSUpdate
  UpdateBallBrightness
  RollingUpdate         ' update rolling sounds
  DoSTAnim            ' handle stand up target animations
  UpdateStandupTargets
  DoDTAnim            ' handle drop target animations
  UpdateDropTargets
  AnimateBumperSkirts
End Sub



Sub table1_Paused:Controller.Pause = 1:End Sub
Sub table1_unPaused:Controller.Pause = 0:End Sub
Sub table1_exit:Controller.stop:End Sub

'**********
' Keys
'**********

Sub table1_KeyDown(ByVal Keycode)
'VR Button Animation
  If keycode = LeftFlipperKey Then
    VR_CabFlipperLeft.X = VR_CabFlipperLeft.X +10
  End If
    If keycode = RightFlipperKey Then
    VR_CabFlipperRight.X = VR_CabFlipperRight.X -10
  End If
  If Keycode = StartGameKey Then
    VR_Cab_StartButton.y = VR_Cab_StartButton.y -5
  End If
  If keycode = AddCreditKey or keycode = 4 then
    VR_Cab_StartButton.y = VR_Cab_StartButton.y -5
  End If
    If keycode = LeftTiltKey Then Nudge 90, 0.5 : SoundNudgeLeft
  If keycode = RightTiltKey Then Nudge 270, 0.5 : SoundNudgeRight
  If keycode = CenterTiltKey Then Nudge 0, 0.5 : SoundNudgeCenter
    If vpmKeyDown(KeyCode)Then Exit Sub
    If keycode = PlungerKey Then PlaySound "fx_PlungerPull", 0, 1, 0.1, 0.25:Plunger.Pullback
End Sub

Sub table1_KeyUp(ByVal Keycode)
'VR Button Animation
    If keycode = LeftFlipperKey Then
    VR_CabFlipperLeft.X = VR_CabFlipperLeft.X -10
  End If
    If keycode = RightFlipperKey Then
    VR_CabFlipperRight.X = VR_CabFlipperRight.X +10
  End If
  If Keycode = StartGameKey Then
    VR_Cab_StartButton.y = VR_Cab_StartButton.y +5
  End If
  If keycode = AddCreditKey or keycode = 4 then
    VR_Cab_StartButton.y = VR_Cab_StartButton.y +5
  End If
'End VR Button Animation

'Change Song
  If keycode = LeftMagnaSave Then PlayPreviousSong
  If keycode = RightMagnaSave Then PlayNextSong
'End Change Song
    If vpmKeyUp(KeyCode)Then Exit Sub
    If keycode = PlungerKey Then PlaySound "fx_plunger", 0, 1, 0.1, 0.25:Plunger.Fire
End Sub


'************************************************************************
'               FLIPPERS
'************************************************************************
Const ReflipAngle = 20

Sub SolLFlipper(Enabled)
  If Enabled Then
    FlipperActivate LeftFlipper, LFPress
    LF.Fire  'leftflipper.rotatetoend

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
    RF.Fire 'rightflipper.rotatetoend

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


Sub LeftFlipper_Animate
  Dim a : a = LeftFlipper.CurrentAngle
  FlipperLSh.RotZ = a

  Dim v, BP
  v = 255.0 * (121.0 -  LeftFlipper.CurrentAngle) / (121.0 -  70.0)

  For each BP in BP_PLF
    BP.Rotz = a
    BP.visible = v < 128.0
  Next
  For each BP in BP_PLFup
    BP.Rotz = a
    BP.visible = v >= 128.0
  Next
End Sub

Sub RightFlipper_Animate
  Dim a : a = RightFlipper.CurrentAngle
  FlipperRSh.RotZ = a

  Dim v, BP
  v = 255.0 * (-121.0 -  RightFlipper.CurrentAngle) / (-121.0 +  70.0)

  For each BP in BP_PRF
    BP.Rotz = a
    BP.visible = v < 128.0
  Next
  For each BP in BP_PRFup
    BP.Rotz = a
    BP.visible = v >= 128.0
  Next
End Sub




'******************************************************
'   ZBBR: BALL BRIGHTNESS
'******************************************************

Const BallBrightness =  1       'Ball brightness - Value between 0 and 1 (0=Dark ... 1=Bright)

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

  Dim BOT: BOT = Getballs
  For s = 0 To UBound(BOT)
    ' Handle z direction
    d_w = b_base*(1 - (BOT(s).z-25)/500)
    If d_w < 30 Then d_w = 30
    ' Handle plunger lane
    If InRect(BOT(s).x,BOT(s).y,PLLeft,PLBottom,PLLeft,PLTop,PLRight,PLTop,PLRight,PLBottom) Then
      d_w = d_w*(PLOffset+PLGain*(BOT(s).y-PLBottom))
    End If
    ' Assign color
    b_r = Int(d_w)
    b_g = Int(d_w)
    b_b = Int(d_w)
    If b_r > 255 Then b_r = 255
    If b_g > 255 Then b_g = 255
    If b_b > 255 Then b_b = 255
    BOT(s).color = b_r + (b_g * 256) + (b_b * 256 * 256)
    'debug.print "--- ball.color level="&b_r
  Next
End Sub


'****************************
'   Room Brightness
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
'  Drain
'*******************************************



' DRAIN & RELEASE
Sub Drain_Hit
  RandomSoundDrain Drain
  Controller.Switch(91) = 1
End Sub

Sub SolRelease(enabled)
  If enabled Then
    RandomSoundBallRelease Drain
    Drain.kick 60, 20
    Controller.Switch(91) = 0
  End If
End Sub


'**************
' Bumper Animations
'**************

Sub Bumper1_Hit:vpmTimer.PulseSw 76: RandomSoundBumperTop Bumper1:End Sub
Sub Bumper2_Hit:vpmTimer.PulseSw 86: RandomSoundBumperMiddle Bumper2:End Sub
Sub Bumper3_Hit:vpmTimer.PulseSw 96:RandomSoundBumperBottom Bumper3:End Sub

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

Sub Bumper3_Animate
  Dim z, BP
  z = Bumper3.CurrentRingOffset
  For Each BP in BP_BR3 : BP.transz = z: Next
End Sub

Dim Bumpers : Bumpers = Array(Bumper1, Bumper2, Bumper3)

Sub AnimateBumperSkirts
  dim r, g, s, x, y, b, z
  ' Animate Bumper switch (experimental)
  For r = 0 To 2
    g = 10000.
    Dim BOT : BOT = GetBalls
    For s = 0 to UBound(BOT)
      x = Bumpers(r).x - BOT(s).x
      y = Bumpers(r).y - BOT(s).y
      b = x * x + y * y
      If b < g Then g = b
    Next
    z = 4
    If g < 80 * 80 Then
      z = 1
    End If
    If r = 0 Then For Each x in BP_BS1: x.Z = z: Next
    If r = 1 Then For Each x in BP_BS2: x.Z = z: Next
    If r = 2 Then For Each x in BP_BS3: x.Z = z: Next
  Next
End Sub


'*********
' Switches
'*********

' Rollovers
Sub sw77_Hit:Controller.Switch(77) = 1:End Sub
Sub sw77_UnHit:Controller.Switch(77) = 0:End Sub

Sub sw61a_Hit:Controller.Switch(61) = 1:DOF 102, DOFOn:End Sub
Sub sw61a_UnHit:Controller.Switch(61) = 0:DOF 102, DOFOff:End Sub

Sub sw61b_Hit:Controller.Switch(61) = 1:DOF 103, DOFOn:End Sub
Sub sw61b_UnHit:Controller.Switch(61) = 0:DOF 103, DOFOff:End Sub

Sub sw61c_Hit:Controller.Switch(61) = 1:DOF 104, DOFOn:End Sub
Sub sw61c_UnHit:Controller.Switch(61) = 0:DOF 104, DOFOff:End Sub

Sub sw61d_Hit:Controller.Switch(61) = 1:DOF 105, DOFOn:End Sub
Sub sw61d_UnHit:Controller.Switch(61) = 0:DOF 105, DOFOff:End Sub

Sub sw67_Hit:Controller.Switch(67) = 1:End Sub
Sub sw67_UnHit:Controller.Switch(67) = 0:End Sub

Sub sw81_Hit:Controller.Switch(81) = 1::End Sub
Sub sw81_UnHit:Controller.Switch(81) = 0:End Sub

Sub sw82_Hit:Controller.Switch(82) = 1::End Sub
Sub sw82_UnHit:Controller.Switch(82) = 0:End Sub

Sub sw80_Hit:Controller.Switch(80) = 1:End Sub
Sub sw80_UnHit:Controller.Switch(80) = 0:End Sub

Sub sw80a_Hit:Controller.Switch(80) = 1:PlaySoundAt "fx_sensor", sw80a:DOF 101, DOFOn:End Sub
Sub sw80a_UnHit:Controller.Switch(80) = 0:DOF 101, DOFOff:End Sub

Sub sw83_Hit:Controller.Switch(83) = 1::End Sub
Sub sw83_UnHit:Controller.Switch(83) = 0:End Sub

Sub sw66_Hit:Controller.Switch(66) = 1:End Sub
Sub sw66_UnHit:Controller.Switch(66) = 0:End Sub

Sub sw72_Hit:Controller.Switch(72) = 1:PlaySoundAt "fx_sensor", sw72:End Sub
Sub sw72_UnHit:Controller.Switch(72) = 0:End Sub

Sub sw71_Hit:Controller.Switch(71) = 1:PlaySoundAt "fx_sensor", sw71:End Sub
Sub sw71_UnHit:Controller.Switch(71) = 0:End Sub


' Switch Animations

' In-Lane / Out-Lane Rollovers

Sub sw77_Animate
  Dim z : z = sw77.CurrentAnimOffset
  Dim BL : For Each BL in BP_sw77 : BL.transz = z: Next
End Sub

Sub sw61a_Animate
  Dim z : z = sw61a.CurrentAnimOffset
  Dim BL : For Each BL in BP_sw61a : BL.transz = z: Next
End Sub

Sub sw61b_Animate
  Dim z : z = sw61b.CurrentAnimOffset
  Dim BL : For Each BL in BP_sw61b : BL.transz = z: Next
End Sub

Sub sw61c_Animate
  Dim z : z = sw61c.CurrentAnimOffset
  Dim BL : For Each BL in BP_sw61c : BL.transz = z: Next
End Sub

Sub sw61d_Animate
  Dim z : z = sw61d.CurrentAnimOffset
  Dim BL : For Each BL in BP_sw61d : BL.transz = z: Next
End Sub

Sub sw67_Animate
  Dim z : z = sw67.CurrentAnimOffset
  Dim BL : For Each BL in BP_sw67 : BL.transz = z: Next
End Sub

Sub sw80_Animate
  Dim z : z = sw80.CurrentAnimOffset
  Dim BL : For Each BL in BP_sw80 : BL.transz = z: Next
End Sub

Sub sw81_Animate
  Dim z : z = sw81.CurrentAnimOffset
  Dim BL : For Each BL in BP_sw81 : BL.transz = z: Next
End Sub

Sub sw82_Animate
  Dim z : z = sw82.CurrentAnimOffset
  Dim BL : For Each BL in BP_sw82 : BL.transz = z: Next
End Sub

Sub sw83_Animate
  Dim z : z = sw83.CurrentAnimOffset
  Dim BL : For Each BL in BP_sw83 : BL.transz = z: Next
End Sub

Sub sw66_Animate
  Dim z : z = sw66.CurrentAnimOffset
  Dim BL : For Each BL in BP_sw66 : BL.transz = z: Next
End Sub


Sub Gate1_Animate
  Dim a : a = Gate1.CurrentAngle
  Dim BL : For Each BL in BP_Gate1 : BL.rotx = a: Next
End Sub

Sub Gate3_Animate
  Dim a : a = Gate3.CurrentAngle
  Dim BL : For Each BL in BP_Gate3 : BL.rotx = a: Next
End Sub

Sub Gate4_Animate
  Dim a : a = Gate4.CurrentAngle / 4.5
  Dim BL : For Each BL in BP_WireGate4 : BL.rotz = -a: Next
End Sub








'WireGAte.RotZ = 38 + Gate4.Currentangle / 4.5

''''' Drop Targets

Sub UpdateDropTargets
  dim BP, tz, rx, ry

  tz = BM_psw84.transz
  rx = BM_psw84.rotx
  ry = BM_psw84.roty
  For each BP in BP_psw84: BP.transz = tz: BP.rotx = rx: BP.roty = ry: Next

  tz = BM_psw64.transz
  rx = BM_psw64.rotx
  ry = BM_psw64.roty
  For each BP in BP_psw64: BP.transz = tz: BP.rotx = rx: BP.roty = ry: Next

  tz = BM_psw74.transz
  rx = BM_psw74.rotx
  ry = BM_psw74.roty
  For each BP in BP_psw74: BP.transz = tz: BP.rotx = rx: BP.roty = ry: Next

  tz = BM_psw85.transz
  rx = BM_psw85.rotx
  ry = BM_psw85.roty
  For each BP in BP_psw85: BP.transz = tz: BP.rotx = rx: BP.roty = ry: Next

  tz = BM_psw65.transz
  rx = BM_psw65.rotx
  ry = BM_psw65.roty
  For each BP in BP_psw65: BP.transz = tz: BP.rotx = rx: BP.roty = ry: Next

  tz = BM_psw75.transz
  rx = BM_psw75.rotx
  ry = BM_psw75.roty
  For each BP in BP_psw75: BP.transz = tz: BP.rotx = rx: BP.roty = ry: Next

End Sub


' PIANO

Sub Gate5_Hit
  If activeball.vely < 0 Then Exit Sub
  activeball.vely = RndNum(0.1,1)
End Sub

Sub Gate6_Hit
  If activeball.vely < 0 Then Exit Sub
  activeball.vely = RndNum(0.1,1)
End Sub

Sub Gate7_Hit
  If activeball.vely < 0 Then Exit Sub
  activeball.vely = RndNum(0.1,1)
End Sub


Sub Gate5_Animate
    Dim a : a = Gate5.CurrentAngle / 2
    Dim BL : For Each BL in BP_g1 : BL.rotx = -a: Next 'this needs reverse direction
End Sub

Sub Gate6_Animate
    Dim a : a = Gate6.CurrentAngle / 2
    Dim BL : For Each BL in BP_g2 : BL.rotx = -a: Next 'this needs reverse direction
End Sub

Sub Gate7_Animate
    Dim a : a = Gate7.CurrentAngle / 2
    Dim BL : For Each BL in BP_g3 : BL.rotx = -a: Next 'this needs reverse direction
End Sub



'**************************************
'Spinner
'**************************************
Sub sw73_Spin():vpmTimer.PulseSw 73:SoundSpinner sw73 :End Sub

' Spinner Animations
Sub sw73_Animate
  Dim spinangle:spinangle = sw73.currentangle
  Dim BL : For Each BL in BP_sw73 : BL.RotX = -spinangle: Next
  SpinnerTShadow.size_y = abs(sin( (spinangle+180) * (2*PI/360)) * 5)
End Sub



'BM_sw73.Visible = 1  'This is the spinner rest position, and is inner russian doll
'BM_sw73B.Visible = 0 'This is the spinner 180 position, and is outer russian doll

'Sub sw73_Animate
'    Dim BP, a, b
'    a = sw73.currentangle
'    If a >= 0 And a < 60 Then
'        b = 0
'    ElseIf a >= 60 And a < 120 Then
'        b = (a - 60) / 60
'    ElseIf a >= 120 And a < 240 Then
'        b = 1
'    ElseIf a >= 240 And a < 300 Then
'        b = 1 + (240 - a) / 60
'    Else
'        b = 0
'    End If
'    For Each BP in BP_sw73
'        BP.RotX = a
'        BP.Opacity = 100 * (1 - b)
'   SpinnerTShadow.size_y = abs(sin( (a+180) * (2*PI/360)) * 5)
'    Next
'    For Each BP in BP_sw73B
'        BP.RotX = a
'        BP.Opacity = 100 * b
'    Next
'End Sub




'Stand-up Targets
Sub sw97_Hit : STHit 97 : End Sub
Sub sw87_Hit : STHit 87 : End Sub
Sub sw62_Hit : STHit 62 : End Sub
Sub sw63_Hit : STHit 63 : End Sub
Sub sw94_Hit : STHit 94 : End Sub


'Rubbers 'the sound is played from the rubber collection hit
'Rubber animations
Dim Rub1, Rub2

Sub sw2a_Hit:vpmTimer.PulseSw 2:Rub1 = 1:sw2a_Timer:End Sub

Sub sw2a_Timer
    Select Case Rub1
        Case 1:r12.Visible = 1:sw2a.TimerEnabled = 1
        Case 2:r12.Visible = 0:r13.Visible = 1
        Case 3:r13.Visible = 0:sw2a.TimerEnabled = 0
    End Select
    Rub1 = Rub1 + 1
End Sub

Sub sw2c_Hit:vpmTimer.PulseSw 2:Rub2 = 1:sw2c_Timer:End Sub
Sub sw2c_Timer
    Select Case Rub2
        Case 1:r14.Visible = 1:sw2c.TimerEnabled = 1
        Case 2:r14.Visible = 0:r15.Visible = 1
        Case 3:r15.Visible = 0:sw2c.TimerEnabled = 0
    End Select
    Rub2 = Rub2 + 1
End Sub

Sub sw90a_Hit:vpmTimer.PulseSw 90:End Sub
Sub sw90b_Hit:vpmTimer.PulseSw 90:End Sub
Sub sw92a_Hit:vpmTimer.PulseSw 92:End Sub
Sub sw92b_Hit:vpmTimer.PulseSw 92:End Sub
Sub sw92c_Hit:vpmTimer.PulseSw 92:End Sub
Sub sw92d_Hit:vpmTimer.PulseSw 92:End Sub
Sub sw92e_Hit:vpmTimer.PulseSw 92:End Sub
Sub sw93_Hit:vpmTimer.PulseSw 93:End Sub

'*********
'Solenoids
'*********

SolCallback(2) = "SolLeftTargetReset"
SolCallback(3) = "SolRightTargetReset"
SolCallback(4) = "SolRelease"
SolCallback(5) = "vpmNudge.SolGameOn"
SolCallback(sLRFlipper) = "SolRFlipper"
SolCallback(sLLFlipper) = "SolLFlipper"

'*****************
'   Gi Effects
'*****************
Dim gilvl   'General Illumination light state tracked for Dynamic Ball Shadows
gilvl = 1

Dim OldGiState
OldGiState = -1 'start witht he Gi off

Sub SetGI(s)
  dim xx
  For each xx in aGiLights:xx.State = s: Next
  ' VR Backglass simple flasher GI
  If VRMode = True Then
    If s = 0 Then
      VRBG_Dark.Visible = 1
      VRBG_Game.Visible = 0
    Else
      VRBG_Dark.Visible = 1
      VRBG_Game.Visible = 1
    End If
  End If
End Sub



Sub GiEffect(enabled)
    If enabled Then
        For each x in aGiLights
            x.Duration 2, 1000, 1
        Next
    End If
End Sub

Sub GIUpdate
    BOT = getballs
    If UBound(BOT) <> OldGiState Then
        OldGiState = Ubound(BOT)
        If UBound(BOT) = -1 Then
            SetGi 0
        Else
            SetGI 1
        End If
    End If
End Sub

'*************
' Update Lamps
'*************

LampTimer.Interval = 40
LampTimer.Enabled = 1

Sub LampTimer_Timer
    GIUpdate
    UpdateLeds
    UpdateLamps
End Sub

Sub UpdateLamps()
    Dim ChgLamp, ii
    ChgLamp = Controller.ChangedLamps
    If Not IsEmpty(ChgLamp)Then
        For ii = 0 To UBound(ChgLamp)
            DisplayLamps chgLamp(ii, 0), chgLamp(ii, 1)
        Next
    End If
End Sub

Sub DisplayLamps(idx, Stat)
    Select Case idx
        Case 1:li1.State = Stat
        Case 2:li2.State = Stat
        Case 3:li3.State = Stat
        Case 4:li4.State = Stat
        Case 5:li5.State = Stat
        Case 6:li6.State = Stat
        Case 7:li7.State = Stat
        Case 8:li8.State = Stat
        Case 9:li9.State = Stat
        Case 10:li10.State = Stat
        Case 11:li11a.State = Stat:li11b.State = Stat:li11c.state = stat
        Case 13:li12.State = Stat 'text="1"
        Case 14:li13.State = Stat 'text="2"
        Case 15:li14.State = Stat 'text="3"
        Case 17:li15.State = Stat
        Case 18:li16.State = Stat
        Case 19:li17.State = Stat
        Case 22:li18.State = Stat
        Case 23:li19.State = Stat
        Case 24:li20.State = Stat
        Case 25:li21.State = Stat
        Case 26:li22.State = Stat
        Case 27:li23.State = Stat
        Case 28:li24.State = Stat:li24b.State = Stat
        Case 29:li25.State = Stat
        Case 30:li26.State = Stat
        Case 31:li27a.State = Stat:li27.State = Stat
        Case 32:li28.State = Stat 'text="Game Over"
        Case 33:li29.State = Stat 'text="Press Start"
        Case 35:li31.State = Stat 'text="Ball in Play"
        Case 36:li32.State = Stat 'text="Match"
        Case 37:li33.State = Stat
        Case 38:li34.State = Stat
        Case 39:li35.State = Stat 'text="Handicap"
        Case 40:If stat Then PlaySound SoundFX("fx_knocker", DOFKnocker):DOF 106,DOFPulse
        Case 41:li36.State = Stat
        Case 42:li37.State = Stat
        Case 43:li38.State = Stat
        Case 44:li39.State = Stat
        Case 45:li40.State = Stat
        Case 46:li41.State = Stat
    End Select
End Sub

'*********************
'LED's based on Eala's
'*********************

Dim Digits(32)
Digits(0) = Array(a00, a02, a05, a06, a04, a01, a03, a07)
Digits(1) = Array(a10, a12, a15, a16, a14, a11, a13)
Digits(2) = Array(a20, a22, a25, a26, a24, a21, a23)
Digits(3) = Array(a30, a32, a35, a36, a34, a31, a33, a37)
Digits(4) = Array(a40, a42, a45, a46, a44, a41, a43)
Digits(5) = Array(a50, a52, a55, a56, a54, a51, a53)
Digits(6) = Array(a60, a62, a65, a66, a64, a61, a63)

Digits(7) = Array(e00, e02, e05, e06, e04, e01, e03, e07)
Digits(8) = Array(e10, e12, e15, e16, e14, e11, e13)
Digits(9) = Array(e20, e22, e25, e26, e24, e21, e23)
Digits(10) = Array(e30, e32, e35, e36, e34, e31, e33, e37)
Digits(11) = Array(e40, e42, e45, e46, e44, e41, e43)
Digits(12) = Array(e50, e52, e55, e56, e54, e51, e53)
Digits(13) = Array(e60, e62, e65, e66, e64, e61, e63)

Digits(14) = Array(b00, b02, b05, b06, b04, b01, b03, b07)
Digits(15) = Array(b10, b12, b15, b16, b14, b11, b13)
Digits(16) = Array(b20, b22, b25, b26, b24, b21, b23)
Digits(17) = Array(b30, b32, b35, b36, b34, b31, b33, b37)
Digits(18) = Array(b40, b42, b45, b46, b44, b41, b43)
Digits(19) = Array(b50, b52, b55, b56, b54, b51, b53)
Digits(20) = Array(b60, b62, b65, b66, b64, b61, b63)

Digits(21) = Array(f00, f02, f05, f06, f04, f01, f03, f07)
Digits(22) = Array(f10, f12, f15, f16, f14, f11, f13)
Digits(23) = Array(f20, f22, f25, f26, f24, f21, f23)
Digits(24) = Array(f30, f32, f35, f36, f34, f31, f33, f37)
Digits(25) = Array(f40, f42, f45, f46, f44, f41, f43)
Digits(26) = Array(f50, f52, f55, f56, f54, f51, f53)
Digits(27) = Array(f60, f62, f65, f66, f64, f61, f63)

Digits(28) = Array(c00, c02, c05, c06, c04, c01, c03)
Digits(29) = Array(c10, c12, c15, c16, c14, c11, c13)

Digits(30) = Array(d00, d02, d05, d06, d04, d01, d03)
Digits(31) = Array(d10, d12, d15, d16, d14, d11, d13)


'*********************
'    VR LED's
'*********************

Dim VRDigits(32)

' Display1 (D1) Player 1
VRDigits(0) = Array(D1L0x0,D1L0x1,D1L0x2,D1L0x3,D1L0x4,D1L0x5,D1L0x6)
VRDigits(1) = Array(D1L1x0,D1L1x1,D1L1x2,D1L1x3,D1L1x4,D1L1x5,D1L1x6)
VRDigits(2) = Array(D1L2x0,D1L2x1,D1L2x2,D1L2x3,D1L2x4,D1L2x5,D1L2x6)
VRDigits(3) = Array(D1L3x0,D1L3x1,D1L3x2,D1L3x3,D1L3x4,D1L3x5,D1L3x6)
VRDigits(4) = Array(D1L4x0,D1L4x1,D1L4x2,D1L4x3,D1L4x4,D1L4x5,D1L4x6)
VRDigits(5) = Array(D1L5x0,D1L5x1,D1L5x2,D1L5x3,D1L5x4,D1L5x5,D1L5x6)
VRDigits(6) = Array(D1L6x0,D1L6x1,D1L6x2,D1L6x3,D1L6x4,D1L6x5,D1L6x6)

' Display2 (D2) Player 2
VRDigits(7) = Array(D2L0x0,D2L0x1,D2L0x2,D2L0x3,D2L0x4,D2L0x5,D2L0x6)
VRDigits(8) = Array(D2L1x0,D2L1x1,D2L1x2,D2L1x3,D2L1x4,D2L1x5,D2L1x6)
VRDigits(9) = Array(D2L2x0,D2L2x1,D2L2x2,D2L2x3,D2L2x4,D2L2x5,D2L2x6)
VRDigits(10) = Array(D2L3x0,D2L3x1,D2L3x2,D2L3x3,D2L3x4,D2L3x5,D2L3x6)
VRDigits(11) = Array(D2L4x0,D2L4x1,D2L4x2,D2L4x3,D2L4x4,D2L4x5,D2L4x6)
VRDigits(12) = Array(D2L5x0,D2L5x1,D2L5x2,D2L5x3,D2L5x4,D2L5x5,D2L5x6)
VRDigits(13) = Array(D2L6x0,D2L6x1,D2L6x2,D2L6x3,D2L6x4,D2L6x5,D2L6x6)

' Display3 (D3) Player 3
VRDigits(14) = Array(D3L0x0,D3L0x1,D3L0x2,D3L0x3,D3L0x4,D3L0x5,D3L0x6)
VRDigits(15) = Array(D3L1x0,D3L1x1,D3L1x2,D3L1x3,D3L1x4,D3L1x5,D3L1x6)
VRDigits(16) = Array(D3L2x0,D3L2x1,D3L2x2,D3L2x3,D3L2x4,D3L2x5,D3L2x6)
VRDigits(17) = Array(D3L3x0,D3L3x1,D3L3x2,D3L3x3,D3L3x4,D3L3x5,D3L3x6)
VRDigits(18) = Array(D3L4x0,D3L4x1,D3L4x2,D3L4x3,D3L4x4,D3L4x5,D3L4x6)
VRDigits(19) = Array(D3L5x0,D3L5x1,D3L5x2,D3L5x3,D3L5x4,D3L5x5,D3L5x6)
VRDigits(20) = Array(D3L6x0,D3L6x1,D3L6x2,D3L6x3,D3L6x4,D3L6x5,D3L6x6)

' Display4 (D4) Player 4
VRDigits(21) = Array(D4L0x0,D4L0x1,D4L0x2,D4L0x3,D4L0x4,D4L0x5,D4L0x6)
VRDigits(22) = Array(D4L1x0,D4L1x1,D4L1x2,D4L1x3,D4L1x4,D4L1x5,D4L1x6)
VRDigits(23) = Array(D4L2x0,D4L2x1,D4L2x2,D4L2x3,D4L2x4,D4L2x5,D4L2x6)
VRDigits(24) = Array(D4L3x0,D4L3x1,D4L3x2,D4L3x3,D4L3x4,D4L3x5,D4L3x6)
VRDigits(25) = Array(D4L4x0,D4L4x1,D4L4x2,D4L4x3,D4L4x4,D4L4x5,D4L4x6)
VRDigits(26) = Array(D4L5x0,D4L5x1,D4L5x2,D4L5x3,D4L5x4,D4L5x5,D4L5x6)
VRDigits(27) = Array(D4L6x0,D4L6x1,D4L6x2,D4L6x3,D4L6x4,D4L6x5,D4L6x6)

' Display5 (D5) Credits
VRDigits(28) = Array(D5L0x0,D5L0x1,D5L0x2,D5L0x3,D5L0x4,D5L0x5,D5L0x6)
VRDigits(29) = Array(D5L1x0,D5L1x1,D5L1x2,D5L1x3,D5L1x4,D5L1x5,D5L1x6)

' Display6 (D6) Match - Balls
VRDigits(30) = Array(D6L0x0,D6L0x1,D6L0x2,D6L0x3,D6L0x4,D6L0x5,D6L0x6)
VRDigits(31) = Array(D6L1x0,D6L1x1,D6L1x2,D6L1x3,D6L1x4,D6L1x5,D6L1x6)



'********************
'Update LED's display
'********************

Sub UpdateLeds()
    Dim ChgLED, ii, num, chg, stat, obj
    ChgLED = Controller.ChangedLEDs(&Hffffffff, &Hffffffff)
    If Not IsEmpty(ChgLED)Then
        For ii = 0 To UBound(chgLED)
            num = chgLED(ii, 0):chg = chgLED(ii, 1):stat = chgLED(ii, 2)
      if (num < 32) then
        If VRMode = True Then
          For Each obj In VRDigits(num)
            If chg And 1 Then obj.visible=stat And 1
            chg=chg\2 : stat=stat\2
          Next
        elseif DesktopMode = True and VRMode = False Then
          For Each obj In Digits(num)
          If chg And 1 Then obj.State = stat And 1
          chg = chg \ 2:stat = stat \ 2
          Next
        end If
      end If
        Next
    End If
  if VRMode = True Then UpdateVRLamps
End Sub

Sub UpdateVRLamps
  IF Controller.Lamp(32) = 0 Then: VRBG_L32.visible = 0 : Else : VRBG_L32.visible = 1 'Game Over
  IF Controller.Lamp(13) = 0 Then: VRBG_L13.visible = 0 : Else : VRBG_L13.visible = 1 '1
  IF Controller.Lamp(14) = 0 Then: VRBG_L14.visible = 0 : Else : VRBG_L14.visible = 1 '2
  IF Controller.Lamp(15) = 0 Then: VRBG_L15.visible = 0 : Else : VRBG_L15.visible = 1 '3
  IF Controller.Lamp(35) = 0 Then: VRBG_L35.visible = 0 : Else : VRBG_L35.visible = 1 'Ball in play
  IF Controller.Lamp(36) = 0 Then: VRBG_L36.visible = 0 : Else : VRBG_L36.visible = 1 'Match
  IF Controller.Lamp(39) = 0 Then: VRBG_L39.visible = 0 : Else : VRBG_L39.visible = 1 '

End Sub








'******************************************************
' VPW TargetBouncer for targets and posts by Iaakki, Wrd1972, Apophis
'******************************************************

Const TargetBouncerEnabled = 1    '0 = normal standup targets, 1 = bouncy targets
Const TargetBouncerFactor = 0.7   'Level of bounces. Recommmended value of 0.7

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

Sub InitPolarity()
  Dim x, a
  a = Array(LF, RF, ULF)
  For Each x In a
    x.AddPt "Ycoef", 0, RightFlipper.Y-65, 1 'disabled
    x.AddPt "Ycoef", 1, RightFlipper.Y-11, 1
    x.enabled = True
    x.TimeDelay = 60
    x.DebugOn=False ' prints some info in debugger

    x.AddPt "Polarity", 0, 0, 0
    x.AddPt "Polarity", 1, 0.05, - 5.5
    x.AddPt "Polarity", 2, 0.16, - 5.5
    x.AddPt "Polarity", 3, 0.20, - 0.75
    x.AddPt "Polarity", 4, 0.25, - 1.25
    x.AddPt "Polarity", 5, 0.3, - 1.75
    x.AddPt "Polarity", 6, 0.4, - 3.5
    x.AddPt "Polarity", 7, 0.5, - 5.25
    x.AddPt "Polarity", 8, 0.7, - 4.0
    x.AddPt "Polarity", 9, 0.75, - 3.5
    x.AddPt "Polarity", 10, 0.8, - 3.0
    x.AddPt "Polarity", 11, 0.85, - 2.5
    x.AddPt "Polarity", 12, 0.9, - 2.0
    x.AddPt "Polarity", 13, 0.95, - 1.5
    x.AddPt "Polarity", 14, 1, - 1.0
    x.AddPt "Polarity", 15, 1.05, -0.5
    x.AddPt "Polarity", 16, 1.1, 0
    x.AddPt "Polarity", 17, 1.3, 0

    x.AddPt "Velocity", 0, 0, 0.85
    x.AddPt "Velocity", 1, 0.23, 0.85
    x.AddPt "Velocity", 2, 0.27, 1
    x.AddPt "Velocity", 3, 0.3, 1
    x.AddPt "Velocity", 4, 0.35, 1
    x.AddPt "Velocity", 5, 0.6, 1 '0.982
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
  Dim BOT
  BOT = GetBalls

  If Flipper1.currentangle = Endangle1 And EOSNudge1 <> 1 Then
    EOSNudge1 = 1
    '   debug.print Flipper1.currentangle &" = "& Endangle1 &"--"& Flipper2.currentangle &" = "& EndAngle2
    If Flipper2.currentangle = EndAngle2 Then
      For b = 0 To UBound(BOT)
        If FlipperTrigger(BOT(b).x, BOT(b).y, Flipper1) Then
          'Debug.Print "ball in flip1. exit"
          Exit Sub
        End If
      Next
      For b = 0 To UBound(BOT)
        If FlipperTrigger(BOT(b).x, BOT(b).y, Flipper2) Then
          BOT(b).velx = BOT(b).velx / 1.3
          BOT(b).vely = BOT(b).vely - 0.5
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
    Dim b', BOT
    '   BOT = GetBalls

    For b = 0 To UBound(BOT)
      If Distance(BOT(b).x, BOT(b).y, Flipper.x, Flipper.y) < 55 Then 'check for cradle
        If BOT(b).vely >= - 0.4 Then BOT(b).vely =  - 0.4
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

' Note, cor.update must be called in a 10 ms timer. The example table uses the GameTimer for this purpose, but sometimes a dedicated timer call RDampen is used.
'
'Sub RDampen_Timer
' Cor.Update
'End Sub



'******************************************************
'****  END PHYSICS DAMPENERS
'******************************************************


'******************************************************
'  DROP TARGETS
'******************************************************

Sub SolLeftTargetReset(enabled)
        if enabled then
    DTRaise 84
    DTRaise 64
    DTRaise 74
      RandomSoundDropTargetReset BM_pSW84
    Dim xx: For each xx in DTShadows: xx.visible = True: Next
    end if
End Sub

Sub SolRightTargetReset(enabled)
    if enabled then
    DTRaise 85
    DTRaise 65
    DTRaise 75
    RandomSoundDropTargetReset BM_pSW85
    Dim xx: For each xx in DTShadows: xx.visible = True: Next
    end if
End Sub

Sub Sw84_Hit : DTHit 84 : End Sub
Sub Sw64_Hit : DTHit 64 : End Sub
Sub Sw74_Hit : DTHit 74 : End Sub
Sub Sw85_Hit : DTHit 85 : End Sub
Sub Sw65_Hit : DTHit 65 : End Sub
Sub Sw75_Hit : DTHit 75 : End Sub

Sub DTAction(switchid)
  Select Case switchid
    Case 64: DTShadows(0).visible = False
    Case 65: DTShadows(1).visible = False
    Case 74: DTShadows(2).visible = False
    Case 75: DTShadows(3).visible = False
    Case 84: DTShadows(4).visible = False
    Case 85: DTShadows(5).visible = False
  End Select
End Sub


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
Dim DT1, DT2, DT3 ,DT4 ,DT5, DT6

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

Set DT1 = (new DropTarget)(sw84, sw84a, BM_psw84, 84, 0, false)
Set DT2 = (new DropTarget)(sw64, sw64a, BM_psw64, 64, 0, false)
Set DT3 = (new DropTarget)(sw74, sw74a, BM_psw74, 74, 0, false)
Set DT4 = (new DropTarget)(sw85, sw85a, BM_psw85, 85, 0, false)
Set DT5 = (new DropTarget)(sw65, sw65a, BM_psw65, 65, 0, false)
Set DT6 = (new DropTarget)(sw75, sw75a, BM_psw75, 75, 0, false)

Dim DTArray
DTArray = Array(DT1, DT2, DT3 ,DT4 ,DT5, DT6)

'Configure the behavior of Drop Targets.
Const DTDropSpeed = 90 'in milliseconds
Const DTDropUpSpeed = 40 'in milliseconds
Const DTDropUnits = 50 'VP units primitive drops so top of at or below the playfield
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
      controller.Switch(Switchid) = 1
      DTAction Switchid
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
      Dim BOT
      BOT = GetBalls

      For b = 0 To UBound(BOT)
        If InRotRect(BOT(b).x,BOT(b).y,prim.x, prim.y, prim.rotz, - 25, - 10,25, - 10,25,25, - 25,25) And BOT(b).z < prim.z + DTDropUnits + 25 Then
          BOT(b).velz = 10
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
    controller.Switch(Switchid) = 0
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
Dim ST97, ST87, ST62, ST63, ST94

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

Set ST97 = (new StandupTarget)(sw97, BM_sw97,97, 0)
Set ST87 = (new StandupTarget)(sw87, BM_sw87,87, 0)
Set ST62 = (new StandupTarget)(sw62, BM_sw62,62, 0)
Set ST63 = (new StandupTarget)(sw63, BM_sw63,63, 0)
Set ST94 = (new StandupTarget)(sw94, BM_sw94,94, 0)

'Add all the Stand-up Target Arrays to Stand-up Target Animation Array
' STAnimationArray = Array(ST1, ST2, ....)
Dim STArray
STArray = Array(ST97, ST87, ST62, ST63, ST94)

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


''''' Standup Targets animation
Sub UpdateStandupTargets
  dim BP, ty

  ty = BM_sw97.transy
  For each BP in BP_sw97 : BP.transy = ty: Next

  ty = BM_sw87.transy
  For each BP in BP_sw87 : BP.transy = ty: Next

  ty = BM_sw62.transy
  For each BP in BP_sw62 : BP.transy = ty: Next

  ty = BM_sw63.transy
  For each BP in BP_sw63 : BP.transy = ty: Next

  ty = BM_sw94.transy
  For each BP in BP_sw94 : BP.transy = ty: Next

End Sub


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
  Dim b, BOT
  BOT = GetBalls

  ' stop the sound of deleted balls
  For b = UBound(BOT) + 1 to tnob - 1
    rolling(b) = False
    StopSound("BallRoll_" & b)
  Next

  ' exit the sub if no balls on the table
  If UBound(BOT) = -1 Then Exit Sub

  ' play the rolling sound for each ball

  For b = 0 to UBound(BOT)
    If BallVel(BOT(b)) > 1 AND BOT(b).z < 30 Then
      rolling(b) = True
      PlaySound ("BallRoll_" & b), -1, VolPlayfieldRoll(BOT(b)) * BallRollVolume * VolumeDial, AudioPan(BOT(b)), 0, PitchPlayfieldRoll(BOT(b)), 1, 0, AudioFade(BOT(b))

    Else
      If rolling(b) = True Then
        StopSound("BallRoll_" & b)
        rolling(b) = False
      End If
    End If

    ' Ball Drop Sounds
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
'****  END BALL ROLLING AND DROP SOUNDS
'******************************************************


'******************************************************
'****  FLEEP MECHANICAL SOUNDS
'******************************************************

' This part in the script is an entire block that is dedicated to the physics sound system.
' Various scripts and sounds that may be pretty generic and could suit other WPC systems, but the most are tailored specifically for the TOM table

' Many of the sounds in this package can be added by creating collections and adding the appropriate objects to those collections.
' Create the following new collections:
'   Metals (all metal objects, metal walls, metal posts, metal wire guides)
'   Apron (the apron walls and plunger wall)
'   Walls (all wood or plastic walls)
'   Rollovers (wire rollover triggers, star triggers, or button triggers)
'   Targets (standup or drop targets, these are hit sounds only ... you will want to add separate dropping sounds for drop targets)
'   Gates (plate gates)
'   GatesWire (wire gates)
'   Rubbers (all rubbers including posts, sleeves, pegs, and bands)
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
' Part 1:   https://youtu.be/PbE2kNiam3g
' Part 2:   https://youtu.be/B5cm1Y8wQsk
' Part 3:   https://youtu.be/eLhWyuYOyGg


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

Sub capron_Hit (idx)
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
  Debug.Print "PlayArchLeft"
  PlaySoundAtLevelActiveBall ("Arch_L" & Int(Rnd*4)+1), Vol(ActiveBall) * ArchSoundFactor
End Sub

Sub RandomSoundRightArch()
  Debug.Print "PlayArchRight"
  PlaySoundAtLevelActiveBall ("Arch_R" & Int(Rnd*4)+1), Vol(ActiveBall) * ArchSoundFactor
End Sub


Sub Arch1_hit()
  Debug.Print "Arch1_hit velx = " & activeball.velx
  If Activeball.velx > 1 Then SoundPlayfieldGate
  StopSound "Arch_L1"
  StopSound "Arch_L2"
  StopSound "Arch_L3"
  StopSound "Arch_L4"
End Sub

Sub Arch1_unhit()
  Debug.Print "Arch1_unhit velx = " & activeball.velx
  If activeball.velx < -8 Then
    RandomSoundRightArch
  End If
End Sub

Sub Arch2_hit()
  Debug.Print "Arch2_hit velx = " & activeball.velx
  If Activeball.velx < 1 Then SoundPlayfieldGate
  StopSound "Arch_R1"
  StopSound "Arch_R2"
  StopSound "Arch_R3"
  StopSound "Arch_R4"
End Sub

Sub Arch2_unhit()
  Debug.Print "Arch2_unhit velx = " & activeball.velx
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
  Dim BOT: BOT=Getballs
  Dim s: For s = lob To UBound(BOT)
    ' *** Normal "ambient light" ball shadow

    'Primitive shadow on playfield, flasher shadow in ramps
    '** If on main and upper pf
    If BOT(s).Z > 20 And BOT(s).Z < 30 Then
      objBallShadow(s).visible = 1
      objBallShadow(s).X = BOT(s).X + (BOT(s).X - (tablewidth / 2)) / (Ballsize / AmbientMovement) + offsetX
      objBallShadow(s).Y = BOT(s).Y + offsetY
      'objBallShadow(s).Z = gBOT(s).Z + s/1000 + 1.04 - 25

    '** No shadow if ball is off the main playfield (this may need to be adjusted per table)
    Else
      objBallShadow(s).visible = 0
    End If
  Next
End Sub



'***************
' Rules by Maior
'***************
Dim Msg(21)
Sub Rules()
    Msg(0) = "Canasta 86 - Inder 1986" &Chr(10)&Chr(10)
    Msg(1) = ""
    Msg(2) = "3 Lanes Top Left: Left: Take Bonus x2 "
    Msg(3) = "   Center: Extra Ball when lit "
    Msg(4) = "   Right: Take bonus (If lit, take bonus x5)."
    Msg(5) = ""
    Msg(6) = "4 Lanes Top Right: Left: Bonus Up (If lit, 100.000 points)"
    Msg(7) = "   Center Left: Lights x5 lane. When lit, 50.000 points"
    Msg(8) = "   Center Right: When Lit, Extra ball"
    Msg(9) = "   Right: When lit Special, free game"
    Msg(10) = ""
    Msg(11) = "3 Left Bank: Two of them lights Extra Ball or Special"
    Msg(12) = "   Three of them lights Extra Ball and Special"
    Msg(13) = ""
    Msg(14) = "3 Right Bank: Two of them lights Extra Ball or Special (see backglass)"
    Msg(15) = "   Three of them lights Extra Ball and Special"
    Msg(16) = ""
    Msg(17) = "Center target: With Orange light, up bonus"
    Msg(18) = "   With Green light, lights x5 lane"
    Msg(19) = "   With Special light, free game"
    Msg(20) = ""
    Msg(21) = "Freegames at 1.500.000 points and 2.500.000 points. "
    For X = 1 To 21
        Msg(0) = Msg(0) + Msg(X)&Chr(13)
    Next
    MsgBox Msg(0), , "         Instructions and Rule Card"
End Sub

' DipVal - translate dip switch letter from Canasta 86 manual to bitmask
' by tynirodent
Function DipVal(Letter)
    If Letter = "A" Then
        DipVal = &H80000000 ' avoid overflow error
    Else
        DipVal = 2 ^(31 -(Asc(Letter)- Asc("A")))
    End If
End Function
Sub myShowDips()
    Dim vpmDips:Set vpmDips = New cvpmDips
    With vpmDips
        .AddForm 500, 400, "Canasta 86 - DIP switches"
        .AddFrame 2, 10, 190, "Balls Per Game", DipVal("E"), Array("3 balls", 0, "5 balls", DipVal("E"))
        .AddFrame 2, 60, 190, "Top Score", DipVal("O") + DipVal("P"), Array("2,000,000", 0, "2,400,000", DipVal("P"), "3,000,000", DipVal("O"), "3,500,000", DipVal("O") + DipVal("P"))
        .AddFrame 205, 10, 190, "Replay 1 Score", DipVal("C") + DipVal("D"), Array("1,500,000", 0, "1,700,000", DipVal("D"), "1,800,000", DipVal("C"), "1,900,000", DipVal("C") + DipVal("D"))
        .AddFrame 205, 90, 190, "Replay 2 Score", DipVal("A") + DipVal("B"), Array("2,500,000", 0, "2,700,000", DipVal("B"), "2,900,000", DipVal("A"), "none", DipVal("A") + DipVal("B"))
        .AddLabel 50, 185, 300, 20, "After hitting OK, press F3 to reset game with new settings."
        .ViewDips
    End With
End Sub
Set vpmShowDips = GetRef("myShowDips")


Dim VRBGObj
Sub SetBackglass
  For Each VRBGObj In VRBackglassDigits
    VRBGObj.x = VRBGObj.x
    VRBGObj.height = - VRBGObj.y + 758
    VRBGObj.y = 2 'adjusts the distance from the backglass towards the user
    VRBGObj.Rotx = -90
  Next
End Sub




' VLM  Arrays - Start
' Arrays per baked part
Dim BP_BR1: BP_BR1=Array(BM_BR1, LM_GI_GI_10_BR1, LM_GI_GI_11_BR1, LM_GI_GI_12_BR1, LM_GI_GI_15_BR1, LM_GI_GI_20_BR1, LM_GI_GI_7_BR1, LM_GI_GI_9_BR1, LM_L_li11a_BR1, LM_L_li11c_BR1)
Dim BP_BR2: BP_BR2=Array(BM_BR2, LM_GI_GI_10_BR2, LM_GI_GI_11_BR2, LM_GI_GI_12_BR2, LM_GI_GI_15_BR2, LM_GI_GI_7_BR2, LM_GI_GI_8_BR2, LM_GI_GI_9_BR2, LM_L_li11b_BR2)
Dim BP_BR3: BP_BR3=Array(BM_BR3, LM_GI_GI_10_BR3, LM_GI_GI_11_BR3, LM_GI_GI_12_BR3, LM_GI_GI_13_BR3, LM_GI_GI_15_BR3, LM_GI_GI_9_BR3, LM_L_li11c_BR3)
Dim BP_BS1: BP_BS1=Array(BM_BS1, LM_GI_GI_10_BS1, LM_GI_GI_11_BS1, LM_GI_GI_12_BS1, LM_GI_GI_15_BS1, LM_GI_GI_17_BS1, LM_GI_GI_20_BS1, LM_GI_GI_9_BS1, LM_L_li11a_BS1)
Dim BP_BS2: BP_BS2=Array(BM_BS2, LM_GI_GI_10_BS2, LM_GI_GI_11_BS2, LM_GI_GI_12_BS2, LM_GI_GI_15_BS2, LM_GI_GI_17_BS2, LM_GI_GI_7_BS2, LM_GI_GI_8_BS2, LM_GI_GI_9_BS2, LM_L_li11b_BS2)
Dim BP_BS3: BP_BS3=Array(BM_BS3, LM_GI_GI_10_BS3, LM_GI_GI_11_BS3, LM_GI_GI_12_BS3, LM_GI_GI_13_BS3, LM_GI_GI_15_BS3, LM_GI_GI_20_BS3, LM_L_li11c_BS3)
Dim BP_Gate1: BP_Gate1=Array(BM_Gate1, LM_GI_GI_19_Gate1)
Dim BP_Gate3: BP_Gate3=Array(BM_Gate3, LM_GI_GI_20_Gate3, LM_GI_GI_23_Gate3)
Dim BP_PLF: BP_PLF=Array(BM_PLF, LM_GI_GI_1_PLF)
Dim BP_PLFup: BP_PLFup=Array(BM_PLFup, LM_GI_GI_1_PLFup)
Dim BP_PRF: BP_PRF=Array(BM_PRF, LM_GI_GI_2_PRF)
Dim BP_PRFup: BP_PRFup=Array(BM_PRFup, LM_GI_GI_2_PRFup)
Dim BP_Parts: BP_Parts=Array(BM_Parts, LM_GI_GI_1_Parts, LM_GI_GI_10_Parts, LM_GI_GI_11_Parts, LM_GI_GI_12_Parts, LM_GI_GI_13_Parts, LM_GI_GI_14_Parts, LM_GI_GI_15_Parts, LM_GI_GI_17_Parts, LM_GI_GI_18_Parts, LM_GI_GI_19_Parts, LM_GI_GI_2_Parts, LM_GI_GI_20_Parts, LM_GI_GI_22_Parts, LM_GI_GI_23_Parts, LM_GI_GI_24_Parts, LM_GI_GI_25_Parts, LM_GI_GI_26_Parts, LM_GI_GI_27_Parts, LM_GI_GI_28_Parts, LM_GI_GI_29_Parts, LM_GI_GI_3_Parts, LM_GI_GI_4_Parts, LM_GI_GI_5_Parts, LM_GI_GI_6_Parts, LM_GI_GI_7_Parts, LM_GI_GI_8_Parts, LM_GI_GI_9_Parts, LM_L_li10_Parts, LM_L_li11a_Parts, LM_L_li11b_Parts, LM_L_li11c_Parts, LM_L_li17_Parts, LM_L_li18_Parts, LM_L_li22_Parts, LM_L_li23_Parts, LM_L_li24_Parts, LM_L_li24b_Parts, LM_L_li25_Parts, LM_L_li26_Parts, LM_L_li27_Parts, LM_L_li27a_Parts, LM_L_li4_Parts, LM_L_li5_Parts, LM_L_li8_Parts, LM_L_li9_Parts)
Dim BP_PincabRails: BP_PincabRails=Array(BM_PincabRails)
Dim BP_Playfield: BP_Playfield=Array(BM_Playfield, LM_GI_GI_1_Playfield, LM_GI_GI_10_Playfield, LM_GI_GI_11_Playfield, LM_GI_GI_12_Playfield, LM_GI_GI_13_Playfield, LM_GI_GI_14_Playfield, LM_GI_GI_15_Playfield, LM_GI_GI_17_Playfield, LM_GI_GI_18_Playfield, LM_GI_GI_19_Playfield, LM_GI_GI_2_Playfield, LM_GI_GI_20_Playfield, LM_GI_GI_22_Playfield, LM_GI_GI_23_Playfield, LM_GI_GI_24_Playfield, LM_GI_GI_25_Playfield, LM_GI_GI_26_Playfield, LM_GI_GI_3_Playfield, LM_L_GI_30_Playfield, LM_GI_GI_4_Playfield, LM_GI_GI_5_Playfield, LM_GI_GI_6_Playfield, LM_GI_GI_7_Playfield, LM_GI_GI_8_Playfield, LM_GI_GI_9_Playfield, LM_L_li1_Playfield, LM_L_li10_Playfield, LM_L_li11a_Playfield, LM_L_li11b_Playfield, LM_L_li11c_Playfield, LM_L_li15_Playfield, LM_L_li16_Playfield, LM_L_li17_Playfield, LM_L_li18_Playfield, LM_L_li19_Playfield, LM_L_li2_Playfield, LM_L_li20_Playfield, LM_L_li22_Playfield, LM_L_li23_Playfield, LM_L_li24_Playfield, LM_L_li24b_Playfield, LM_L_li25_Playfield, LM_L_li26_Playfield, LM_L_li27_Playfield, _
  LM_L_li27a_Playfield, LM_L_li3_Playfield, LM_L_li33_Playfield, LM_L_li34_Playfield, LM_L_li36_Playfield, LM_L_li37_Playfield, LM_L_li38_Playfield, LM_L_li39_Playfield, LM_L_li4_Playfield, LM_L_li40_Playfield, LM_L_li41_Playfield, LM_L_li5_Playfield, LM_L_li6_Playfield, LM_L_li7_Playfield, LM_L_li8_Playfield, LM_L_li9_Playfield)
Dim BP_UnderPF: BP_UnderPF=Array(BM_UnderPF, LM_GI_GI_10_UnderPF, LM_GI_GI_14_UnderPF, LM_GI_GI_15_UnderPF, LM_GI_GI_17_UnderPF, LM_GI_GI_18_UnderPF, LM_GI_GI_20_UnderPF, LM_GI_GI_27_UnderPF, LM_GI_GI_28_UnderPF, LM_GI_GI_29_UnderPF, LM_L_GI_30_UnderPF, LM_GI_GI_9_UnderPF, LM_L_li1_UnderPF, LM_L_li10_UnderPF, LM_L_li15_UnderPF, LM_L_li16_UnderPF, LM_L_li17_UnderPF, LM_L_li18_UnderPF, LM_L_li19_UnderPF, LM_L_li2_UnderPF, LM_L_li20_UnderPF, LM_L_li22_UnderPF, LM_L_li23_UnderPF, LM_L_li24_UnderPF, LM_L_li24b_UnderPF, LM_L_li25_UnderPF, LM_L_li26_UnderPF, LM_L_li27_UnderPF, LM_L_li27a_UnderPF, LM_L_li3_UnderPF, LM_L_li33_UnderPF, LM_L_li34_UnderPF, LM_L_li36_UnderPF, LM_L_li37_UnderPF, LM_L_li38_UnderPF, LM_L_li39_UnderPF, LM_L_li4_UnderPF, LM_L_li40_UnderPF, LM_L_li41_UnderPF, LM_L_li5_UnderPF, LM_L_li6_UnderPF, LM_L_li7_UnderPF, LM_L_li8_UnderPF, LM_L_li9_UnderPF)
Dim BP_WireGate4: BP_WireGate4=Array(BM_WireGate4, LM_GI_GI_11_WireGate4, LM_GI_GI_15_WireGate4, LM_GI_GI_20_WireGate4, LM_GI_GI_22_WireGate4)
Dim BP_g1: BP_g1=Array(BM_g1, LM_GI_GI_14_g1, LM_GI_GI_15_g1, LM_GI_GI_17_g1, LM_GI_GI_18_g1, LM_GI_GI_19_g1, LM_GI_GI_27_g1, LM_GI_GI_28_g1, LM_GI_GI_29_g1)
Dim BP_g2: BP_g2=Array(BM_g2, LM_GI_GI_14_g2, LM_GI_GI_15_g2, LM_GI_GI_17_g2, LM_GI_GI_18_g2, LM_GI_GI_19_g2, LM_GI_GI_27_g2, LM_GI_GI_28_g2, LM_GI_GI_29_g2, LM_GI_GI_9_g2)
Dim BP_g3: BP_g3=Array(BM_g3, LM_GI_GI_15_g3, LM_GI_GI_17_g3, LM_GI_GI_18_g3, LM_GI_GI_19_g3, LM_GI_GI_27_g3, LM_GI_GI_28_g3, LM_GI_GI_29_g3)
Dim BP_psw64: BP_psw64=Array(BM_psw64, LM_GI_GI_10_psw64, LM_GI_GI_7_psw64, LM_GI_GI_8_psw64, LM_GI_GI_9_psw64, LM_L_li18_psw64)
Dim BP_psw65: BP_psw65=Array(BM_psw65, LM_GI_GI_11_psw65, LM_GI_GI_12_psw65, LM_GI_GI_13_psw65, LM_GI_GI_15_psw65, LM_L_li19_psw65)
Dim BP_psw74: BP_psw74=Array(BM_psw74, LM_GI_GI_10_psw74, LM_GI_GI_11_psw74, LM_GI_GI_15_psw74, LM_GI_GI_7_psw74, LM_GI_GI_8_psw74, LM_GI_GI_9_psw74, LM_L_li18_psw74)
Dim BP_psw75: BP_psw75=Array(BM_psw75, LM_GI_GI_10_psw75, LM_GI_GI_11_psw75, LM_GI_GI_12_psw75, LM_GI_GI_13_psw75, LM_GI_GI_15_psw75, LM_L_li19_psw75)
Dim BP_psw84: BP_psw84=Array(BM_psw84, LM_GI_GI_10_psw84, LM_GI_GI_7_psw84, LM_GI_GI_8_psw84, LM_GI_GI_9_psw84, LM_L_li18_psw84)
Dim BP_psw85: BP_psw85=Array(BM_psw85, LM_GI_GI_11_psw85, LM_GI_GI_12_psw85, LM_GI_GI_13_psw85, LM_GI_GI_15_psw85, LM_L_li19_psw85)
Dim BP_sw61a: BP_sw61a=Array(BM_sw61a, LM_GI_GI_3_sw61a)
Dim BP_sw61b: BP_sw61b=Array(BM_sw61b, LM_GI_GI_3_sw61b, LM_GI_GI_4_sw61b)
Dim BP_sw61c: BP_sw61c=Array(BM_sw61c, LM_GI_GI_5_sw61c, LM_GI_GI_6_sw61c)
Dim BP_sw61d: BP_sw61d=Array(BM_sw61d, LM_GI_GI_6_sw61d)
Dim BP_sw62: BP_sw62=Array(BM_sw62, LM_GI_GI_10_sw62, LM_GI_GI_7_sw62, LM_GI_GI_8_sw62, LM_GI_GI_9_sw62)
Dim BP_sw63: BP_sw63=Array(BM_sw63, LM_GI_GI_11_sw63, LM_GI_GI_12_sw63, LM_GI_GI_13_sw63, LM_GI_GI_15_sw63)
Dim BP_sw66: BP_sw66=Array(BM_sw66, LM_GI_GI_20_sw66, LM_GI_GI_22_sw66)
Dim BP_sw67: BP_sw67=Array(BM_sw67)
Dim BP_sw73: BP_sw73=Array(BM_sw73, LM_GI_GI_10_sw73, LM_GI_GI_11_sw73, LM_GI_GI_14_sw73, LM_GI_GI_15_sw73, LM_GI_GI_17_sw73, LM_GI_GI_27_sw73, LM_GI_GI_28_sw73, LM_GI_GI_8_sw73, LM_GI_GI_9_sw73, LM_L_li11b_sw73, LM_L_li4_sw73)
Dim BP_sw77: BP_sw77=Array(BM_sw77)
Dim BP_sw80: BP_sw80=Array(BM_sw80, LM_GI_GI_25_sw80, LM_GI_GI_26_sw80)
Dim BP_sw81: BP_sw81=Array(BM_sw81, LM_GI_GI_17_sw81, LM_GI_GI_18_sw81, LM_GI_GI_24_sw81)
Dim BP_sw82: BP_sw82=Array(BM_sw82, LM_GI_GI_18_sw82, LM_GI_GI_24_sw82, LM_GI_GI_25_sw82)
Dim BP_sw83: BP_sw83=Array(BM_sw83, LM_GI_GI_20_sw83, LM_GI_GI_23_sw83, LM_GI_GI_26_sw83)
Dim BP_sw87: BP_sw87=Array(BM_sw87, LM_GI_GI_13_sw87)
Dim BP_sw94: BP_sw94=Array(BM_sw94, LM_GI_GI_10_sw94, LM_GI_GI_11_sw94, LM_GI_GI_13_sw94, LM_GI_GI_9_sw94)
Dim BP_sw97: BP_sw97=Array(BM_sw97, LM_GI_GI_7_sw97)
' Arrays per lighting scenario
Dim BL_GI_GI_1: BL_GI_GI_1=Array(LM_GI_GI_1_PLF, LM_GI_GI_1_PLFup, LM_GI_GI_1_Parts, LM_GI_GI_1_Playfield)
Dim BL_GI_GI_10: BL_GI_GI_10=Array(LM_GI_GI_10_BR1, LM_GI_GI_10_BR2, LM_GI_GI_10_BR3, LM_GI_GI_10_BS1, LM_GI_GI_10_BS2, LM_GI_GI_10_BS3, LM_GI_GI_10_Parts, LM_GI_GI_10_Playfield, LM_GI_GI_10_UnderPF, LM_GI_GI_10_psw64, LM_GI_GI_10_psw74, LM_GI_GI_10_psw75, LM_GI_GI_10_psw84, LM_GI_GI_10_sw62, LM_GI_GI_10_sw73, LM_GI_GI_10_sw94)
Dim BL_GI_GI_11: BL_GI_GI_11=Array(LM_GI_GI_11_BR1, LM_GI_GI_11_BR2, LM_GI_GI_11_BR3, LM_GI_GI_11_BS1, LM_GI_GI_11_BS2, LM_GI_GI_11_BS3, LM_GI_GI_11_Parts, LM_GI_GI_11_Playfield, LM_GI_GI_11_WireGate4, LM_GI_GI_11_psw65, LM_GI_GI_11_psw74, LM_GI_GI_11_psw75, LM_GI_GI_11_psw85, LM_GI_GI_11_sw63, LM_GI_GI_11_sw73, LM_GI_GI_11_sw94)
Dim BL_GI_GI_12: BL_GI_GI_12=Array(LM_GI_GI_12_BR1, LM_GI_GI_12_BR2, LM_GI_GI_12_BR3, LM_GI_GI_12_BS1, LM_GI_GI_12_BS2, LM_GI_GI_12_BS3, LM_GI_GI_12_Parts, LM_GI_GI_12_Playfield, LM_GI_GI_12_psw65, LM_GI_GI_12_psw75, LM_GI_GI_12_psw85, LM_GI_GI_12_sw63)
Dim BL_GI_GI_13: BL_GI_GI_13=Array(LM_GI_GI_13_BR3, LM_GI_GI_13_BS3, LM_GI_GI_13_Parts, LM_GI_GI_13_Playfield, LM_GI_GI_13_psw65, LM_GI_GI_13_psw75, LM_GI_GI_13_psw85, LM_GI_GI_13_sw63, LM_GI_GI_13_sw87, LM_GI_GI_13_sw94)
Dim BL_GI_GI_14: BL_GI_GI_14=Array(LM_GI_GI_14_Parts, LM_GI_GI_14_Playfield, LM_GI_GI_14_UnderPF, LM_GI_GI_14_g1, LM_GI_GI_14_g2, LM_GI_GI_14_sw73)
Dim BL_GI_GI_15: BL_GI_GI_15=Array(LM_GI_GI_15_BR1, LM_GI_GI_15_BR2, LM_GI_GI_15_BR3, LM_GI_GI_15_BS1, LM_GI_GI_15_BS2, LM_GI_GI_15_BS3, LM_GI_GI_15_Parts, LM_GI_GI_15_Playfield, LM_GI_GI_15_UnderPF, LM_GI_GI_15_WireGate4, LM_GI_GI_15_g1, LM_GI_GI_15_g2, LM_GI_GI_15_g3, LM_GI_GI_15_psw65, LM_GI_GI_15_psw74, LM_GI_GI_15_psw75, LM_GI_GI_15_psw85, LM_GI_GI_15_sw63, LM_GI_GI_15_sw73)
Dim BL_GI_GI_17: BL_GI_GI_17=Array(LM_GI_GI_17_BS1, LM_GI_GI_17_BS2, LM_GI_GI_17_Parts, LM_GI_GI_17_Playfield, LM_GI_GI_17_UnderPF, LM_GI_GI_17_g1, LM_GI_GI_17_g2, LM_GI_GI_17_g3, LM_GI_GI_17_sw73, LM_GI_GI_17_sw81)
Dim BL_GI_GI_18: BL_GI_GI_18=Array(LM_GI_GI_18_Parts, LM_GI_GI_18_Playfield, LM_GI_GI_18_UnderPF, LM_GI_GI_18_g1, LM_GI_GI_18_g2, LM_GI_GI_18_g3, LM_GI_GI_18_sw81, LM_GI_GI_18_sw82)
Dim BL_GI_GI_19: BL_GI_GI_19=Array(LM_GI_GI_19_Gate1, LM_GI_GI_19_Parts, LM_GI_GI_19_Playfield, LM_GI_GI_19_g1, LM_GI_GI_19_g2, LM_GI_GI_19_g3)
Dim BL_GI_GI_2: BL_GI_GI_2=Array(LM_GI_GI_2_PRF, LM_GI_GI_2_PRFup, LM_GI_GI_2_Parts, LM_GI_GI_2_Playfield)
Dim BL_GI_GI_20: BL_GI_GI_20=Array(LM_GI_GI_20_BR1, LM_GI_GI_20_BS1, LM_GI_GI_20_BS3, LM_GI_GI_20_Gate3, LM_GI_GI_20_Parts, LM_GI_GI_20_Playfield, LM_GI_GI_20_UnderPF, LM_GI_GI_20_WireGate4, LM_GI_GI_20_sw66, LM_GI_GI_20_sw83)
Dim BL_GI_GI_22: BL_GI_GI_22=Array(LM_GI_GI_22_Parts, LM_GI_GI_22_Playfield, LM_GI_GI_22_WireGate4, LM_GI_GI_22_sw66)
Dim BL_GI_GI_23: BL_GI_GI_23=Array(LM_GI_GI_23_Gate3, LM_GI_GI_23_Parts, LM_GI_GI_23_Playfield, LM_GI_GI_23_sw83)
Dim BL_GI_GI_24: BL_GI_GI_24=Array(LM_GI_GI_24_Parts, LM_GI_GI_24_Playfield, LM_GI_GI_24_sw81, LM_GI_GI_24_sw82)
Dim BL_GI_GI_25: BL_GI_GI_25=Array(LM_GI_GI_25_Parts, LM_GI_GI_25_Playfield, LM_GI_GI_25_sw80, LM_GI_GI_25_sw82)
Dim BL_GI_GI_26: BL_GI_GI_26=Array(LM_GI_GI_26_Parts, LM_GI_GI_26_Playfield, LM_GI_GI_26_sw80, LM_GI_GI_26_sw83)
Dim BL_GI_GI_27: BL_GI_GI_27=Array(LM_GI_GI_27_Parts, LM_GI_GI_27_UnderPF, LM_GI_GI_27_g1, LM_GI_GI_27_g2, LM_GI_GI_27_g3, LM_GI_GI_27_sw73)
Dim BL_GI_GI_28: BL_GI_GI_28=Array(LM_GI_GI_28_Parts, LM_GI_GI_28_UnderPF, LM_GI_GI_28_g1, LM_GI_GI_28_g2, LM_GI_GI_28_g3, LM_GI_GI_28_sw73)
Dim BL_GI_GI_29: BL_GI_GI_29=Array(LM_GI_GI_29_Parts, LM_GI_GI_29_UnderPF, LM_GI_GI_29_g1, LM_GI_GI_29_g2, LM_GI_GI_29_g3)
Dim BL_GI_GI_3: BL_GI_GI_3=Array(LM_GI_GI_3_Parts, LM_GI_GI_3_Playfield, LM_GI_GI_3_sw61a, LM_GI_GI_3_sw61b)
Dim BL_GI_GI_4: BL_GI_GI_4=Array(LM_GI_GI_4_Parts, LM_GI_GI_4_Playfield, LM_GI_GI_4_sw61b)
Dim BL_GI_GI_5: BL_GI_GI_5=Array(LM_GI_GI_5_Parts, LM_GI_GI_5_Playfield, LM_GI_GI_5_sw61c)
Dim BL_GI_GI_6: BL_GI_GI_6=Array(LM_GI_GI_6_Parts, LM_GI_GI_6_Playfield, LM_GI_GI_6_sw61c, LM_GI_GI_6_sw61d)
Dim BL_GI_GI_7: BL_GI_GI_7=Array(LM_GI_GI_7_BR1, LM_GI_GI_7_BR2, LM_GI_GI_7_BS2, LM_GI_GI_7_Parts, LM_GI_GI_7_Playfield, LM_GI_GI_7_psw64, LM_GI_GI_7_psw74, LM_GI_GI_7_psw84, LM_GI_GI_7_sw62, LM_GI_GI_7_sw97)
Dim BL_GI_GI_8: BL_GI_GI_8=Array(LM_GI_GI_8_BR2, LM_GI_GI_8_BS2, LM_GI_GI_8_Parts, LM_GI_GI_8_Playfield, LM_GI_GI_8_psw64, LM_GI_GI_8_psw74, LM_GI_GI_8_psw84, LM_GI_GI_8_sw62, LM_GI_GI_8_sw73)
Dim BL_GI_GI_9: BL_GI_GI_9=Array(LM_GI_GI_9_BR1, LM_GI_GI_9_BR2, LM_GI_GI_9_BR3, LM_GI_GI_9_BS1, LM_GI_GI_9_BS2, LM_GI_GI_9_Parts, LM_GI_GI_9_Playfield, LM_GI_GI_9_UnderPF, LM_GI_GI_9_g2, LM_GI_GI_9_psw64, LM_GI_GI_9_psw74, LM_GI_GI_9_psw84, LM_GI_GI_9_sw62, LM_GI_GI_9_sw73, LM_GI_GI_9_sw94)
Dim BL_L_GI_30: BL_L_GI_30=Array(LM_L_GI_30_Playfield, LM_L_GI_30_UnderPF)
Dim BL_L_li1: BL_L_li1=Array(LM_L_li1_Playfield, LM_L_li1_UnderPF)
Dim BL_L_li10: BL_L_li10=Array(LM_L_li10_Parts, LM_L_li10_Playfield, LM_L_li10_UnderPF)
Dim BL_L_li11a: BL_L_li11a=Array(LM_L_li11a_BR1, LM_L_li11a_BS1, LM_L_li11a_Parts, LM_L_li11a_Playfield)
Dim BL_L_li11b: BL_L_li11b=Array(LM_L_li11b_BR2, LM_L_li11b_BS2, LM_L_li11b_Parts, LM_L_li11b_Playfield, LM_L_li11b_sw73)
Dim BL_L_li11c: BL_L_li11c=Array(LM_L_li11c_BR1, LM_L_li11c_BR3, LM_L_li11c_BS3, LM_L_li11c_Parts, LM_L_li11c_Playfield)
Dim BL_L_li15: BL_L_li15=Array(LM_L_li15_Playfield, LM_L_li15_UnderPF)
Dim BL_L_li16: BL_L_li16=Array(LM_L_li16_Playfield, LM_L_li16_UnderPF)
Dim BL_L_li17: BL_L_li17=Array(LM_L_li17_Parts, LM_L_li17_Playfield, LM_L_li17_UnderPF)
Dim BL_L_li18: BL_L_li18=Array(LM_L_li18_Parts, LM_L_li18_Playfield, LM_L_li18_UnderPF, LM_L_li18_psw64, LM_L_li18_psw74, LM_L_li18_psw84)
Dim BL_L_li19: BL_L_li19=Array(LM_L_li19_Playfield, LM_L_li19_UnderPF, LM_L_li19_psw65, LM_L_li19_psw75, LM_L_li19_psw85)
Dim BL_L_li2: BL_L_li2=Array(LM_L_li2_Playfield, LM_L_li2_UnderPF)
Dim BL_L_li20: BL_L_li20=Array(LM_L_li20_Playfield, LM_L_li20_UnderPF)
Dim BL_L_li22: BL_L_li22=Array(LM_L_li22_Parts, LM_L_li22_Playfield, LM_L_li22_UnderPF)
Dim BL_L_li23: BL_L_li23=Array(LM_L_li23_Parts, LM_L_li23_Playfield, LM_L_li23_UnderPF)
Dim BL_L_li24: BL_L_li24=Array(LM_L_li24_Parts, LM_L_li24_Playfield, LM_L_li24_UnderPF)
Dim BL_L_li24b: BL_L_li24b=Array(LM_L_li24b_Parts, LM_L_li24b_Playfield, LM_L_li24b_UnderPF)
Dim BL_L_li25: BL_L_li25=Array(LM_L_li25_Parts, LM_L_li25_Playfield, LM_L_li25_UnderPF)
Dim BL_L_li26: BL_L_li26=Array(LM_L_li26_Parts, LM_L_li26_Playfield, LM_L_li26_UnderPF)
Dim BL_L_li27: BL_L_li27=Array(LM_L_li27_Parts, LM_L_li27_Playfield, LM_L_li27_UnderPF)
Dim BL_L_li27a: BL_L_li27a=Array(LM_L_li27a_Parts, LM_L_li27a_Playfield, LM_L_li27a_UnderPF)
Dim BL_L_li3: BL_L_li3=Array(LM_L_li3_Playfield, LM_L_li3_UnderPF)
Dim BL_L_li33: BL_L_li33=Array(LM_L_li33_Playfield, LM_L_li33_UnderPF)
Dim BL_L_li34: BL_L_li34=Array(LM_L_li34_Playfield, LM_L_li34_UnderPF)
Dim BL_L_li36: BL_L_li36=Array(LM_L_li36_Playfield, LM_L_li36_UnderPF)
Dim BL_L_li37: BL_L_li37=Array(LM_L_li37_Playfield, LM_L_li37_UnderPF)
Dim BL_L_li38: BL_L_li38=Array(LM_L_li38_Playfield, LM_L_li38_UnderPF)
Dim BL_L_li39: BL_L_li39=Array(LM_L_li39_Playfield, LM_L_li39_UnderPF)
Dim BL_L_li4: BL_L_li4=Array(LM_L_li4_Parts, LM_L_li4_Playfield, LM_L_li4_UnderPF, LM_L_li4_sw73)
Dim BL_L_li40: BL_L_li40=Array(LM_L_li40_Playfield, LM_L_li40_UnderPF)
Dim BL_L_li41: BL_L_li41=Array(LM_L_li41_Playfield, LM_L_li41_UnderPF)
Dim BL_L_li5: BL_L_li5=Array(LM_L_li5_Parts, LM_L_li5_Playfield, LM_L_li5_UnderPF)
Dim BL_L_li6: BL_L_li6=Array(LM_L_li6_Playfield, LM_L_li6_UnderPF)
Dim BL_L_li7: BL_L_li7=Array(LM_L_li7_Playfield, LM_L_li7_UnderPF)
Dim BL_L_li8: BL_L_li8=Array(LM_L_li8_Parts, LM_L_li8_Playfield, LM_L_li8_UnderPF)
Dim BL_L_li9: BL_L_li9=Array(LM_L_li9_Parts, LM_L_li9_Playfield, LM_L_li9_UnderPF)
Dim BL_World: BL_World=Array(BM_BR1, BM_BR2, BM_BR3, BM_BS1, BM_BS2, BM_BS3, BM_Gate1, BM_Gate3, BM_PLF, BM_PLFup, BM_PRF, BM_PRFup, BM_Parts, BM_PincabRails, BM_Playfield, BM_UnderPF, BM_WireGate4, BM_g1, BM_g2, BM_g3, BM_psw64, BM_psw65, BM_psw74, BM_psw75, BM_psw84, BM_psw85, BM_sw61a, BM_sw61b, BM_sw61c, BM_sw61d, BM_sw62, BM_sw63, BM_sw66, BM_sw67, BM_sw73, BM_sw77, BM_sw80, BM_sw81, BM_sw82, BM_sw83, BM_sw87, BM_sw94, BM_sw97)
' Global arrays
Dim BG_Bakemap: BG_Bakemap=Array(BM_BR1, BM_BR2, BM_BR3, BM_BS1, BM_BS2, BM_BS3, BM_Gate1, BM_Gate3, BM_PLF, BM_PLFup, BM_PRF, BM_PRFup, BM_Parts, BM_PincabRails, BM_Playfield, BM_UnderPF, BM_WireGate4, BM_g1, BM_g2, BM_g3, BM_psw64, BM_psw65, BM_psw74, BM_psw75, BM_psw84, BM_psw85, BM_sw61a, BM_sw61b, BM_sw61c, BM_sw61d, BM_sw62, BM_sw63, BM_sw66, BM_sw67, BM_sw73, BM_sw77, BM_sw80, BM_sw81, BM_sw82, BM_sw83, BM_sw87, BM_sw94, BM_sw97)
Dim BG_Lightmap: BG_Lightmap=Array(LM_GI_GI_1_PLF, LM_GI_GI_1_PLFup, LM_GI_GI_1_Parts, LM_GI_GI_1_Playfield, LM_GI_GI_10_BR1, LM_GI_GI_10_BR2, LM_GI_GI_10_BR3, LM_GI_GI_10_BS1, LM_GI_GI_10_BS2, LM_GI_GI_10_BS3, LM_GI_GI_10_Parts, LM_GI_GI_10_Playfield, LM_GI_GI_10_UnderPF, LM_GI_GI_10_psw64, LM_GI_GI_10_psw74, LM_GI_GI_10_psw75, LM_GI_GI_10_psw84, LM_GI_GI_10_sw62, LM_GI_GI_10_sw73, LM_GI_GI_10_sw94, LM_GI_GI_11_BR1, LM_GI_GI_11_BR2, LM_GI_GI_11_BR3, LM_GI_GI_11_BS1, LM_GI_GI_11_BS2, LM_GI_GI_11_BS3, LM_GI_GI_11_Parts, LM_GI_GI_11_Playfield, LM_GI_GI_11_WireGate4, LM_GI_GI_11_psw65, LM_GI_GI_11_psw74, LM_GI_GI_11_psw75, LM_GI_GI_11_psw85, LM_GI_GI_11_sw63, LM_GI_GI_11_sw73, LM_GI_GI_11_sw94, LM_GI_GI_12_BR1, LM_GI_GI_12_BR2, LM_GI_GI_12_BR3, LM_GI_GI_12_BS1, LM_GI_GI_12_BS2, LM_GI_GI_12_BS3, LM_GI_GI_12_Parts, LM_GI_GI_12_Playfield, LM_GI_GI_12_psw65, LM_GI_GI_12_psw75, LM_GI_GI_12_psw85, LM_GI_GI_12_sw63, LM_GI_GI_13_BR3, LM_GI_GI_13_BS3, LM_GI_GI_13_Parts, LM_GI_GI_13_Playfield, LM_GI_GI_13_psw65, _
  LM_GI_GI_13_psw75, LM_GI_GI_13_psw85, LM_GI_GI_13_sw63, LM_GI_GI_13_sw87, LM_GI_GI_13_sw94, LM_GI_GI_14_Parts, LM_GI_GI_14_Playfield, LM_GI_GI_14_UnderPF, LM_GI_GI_14_g1, LM_GI_GI_14_g2, LM_GI_GI_14_sw73, LM_GI_GI_15_BR1, LM_GI_GI_15_BR2, LM_GI_GI_15_BR3, LM_GI_GI_15_BS1, LM_GI_GI_15_BS2, LM_GI_GI_15_BS3, LM_GI_GI_15_Parts, LM_GI_GI_15_Playfield, LM_GI_GI_15_UnderPF, LM_GI_GI_15_WireGate4, LM_GI_GI_15_g1, LM_GI_GI_15_g2, LM_GI_GI_15_g3, LM_GI_GI_15_psw65, LM_GI_GI_15_psw74, LM_GI_GI_15_psw75, LM_GI_GI_15_psw85, LM_GI_GI_15_sw63, LM_GI_GI_15_sw73, LM_GI_GI_17_BS1, LM_GI_GI_17_BS2, LM_GI_GI_17_Parts, LM_GI_GI_17_Playfield, LM_GI_GI_17_UnderPF, LM_GI_GI_17_g1, LM_GI_GI_17_g2, LM_GI_GI_17_g3, LM_GI_GI_17_sw73, LM_GI_GI_17_sw81, LM_GI_GI_18_Parts, LM_GI_GI_18_Playfield, LM_GI_GI_18_UnderPF, LM_GI_GI_18_g1, LM_GI_GI_18_g2, LM_GI_GI_18_g3, LM_GI_GI_18_sw81, LM_GI_GI_18_sw82, LM_GI_GI_19_Gate1, LM_GI_GI_19_Parts, LM_GI_GI_19_Playfield, LM_GI_GI_19_g1, LM_GI_GI_19_g2, LM_GI_GI_19_g3, LM_GI_GI_2_PRF, LM_GI_GI_2_PRFup, _
  LM_GI_GI_2_Parts, LM_GI_GI_2_Playfield, LM_GI_GI_20_BR1, LM_GI_GI_20_BS1, LM_GI_GI_20_BS3, LM_GI_GI_20_Gate3, LM_GI_GI_20_Parts, LM_GI_GI_20_Playfield, LM_GI_GI_20_UnderPF, LM_GI_GI_20_WireGate4, LM_GI_GI_20_sw66, LM_GI_GI_20_sw83, LM_GI_GI_22_Parts, LM_GI_GI_22_Playfield, LM_GI_GI_22_WireGate4, LM_GI_GI_22_sw66, LM_GI_GI_23_Gate3, LM_GI_GI_23_Parts, LM_GI_GI_23_Playfield, LM_GI_GI_23_sw83, LM_GI_GI_24_Parts, LM_GI_GI_24_Playfield, LM_GI_GI_24_sw81, LM_GI_GI_24_sw82, LM_GI_GI_25_Parts, LM_GI_GI_25_Playfield, LM_GI_GI_25_sw80, LM_GI_GI_25_sw82, LM_GI_GI_26_Parts, LM_GI_GI_26_Playfield, LM_GI_GI_26_sw80, LM_GI_GI_26_sw83, LM_GI_GI_27_Parts, LM_GI_GI_27_UnderPF, LM_GI_GI_27_g1, LM_GI_GI_27_g2, LM_GI_GI_27_g3, LM_GI_GI_27_sw73, LM_GI_GI_28_Parts, LM_GI_GI_28_UnderPF, LM_GI_GI_28_g1, LM_GI_GI_28_g2, LM_GI_GI_28_g3, LM_GI_GI_28_sw73, LM_GI_GI_29_Parts, LM_GI_GI_29_UnderPF, LM_GI_GI_29_g1, LM_GI_GI_29_g2, LM_GI_GI_29_g3, LM_GI_GI_3_Parts, LM_GI_GI_3_Playfield, LM_GI_GI_3_sw61a, LM_GI_GI_3_sw61b, LM_GI_GI_4_Parts, _
  LM_GI_GI_4_Playfield, LM_GI_GI_4_sw61b, LM_GI_GI_5_Parts, LM_GI_GI_5_Playfield, LM_GI_GI_5_sw61c, LM_GI_GI_6_Parts, LM_GI_GI_6_Playfield, LM_GI_GI_6_sw61c, LM_GI_GI_6_sw61d, LM_GI_GI_7_BR1, LM_GI_GI_7_BR2, LM_GI_GI_7_BS2, LM_GI_GI_7_Parts, LM_GI_GI_7_Playfield, LM_GI_GI_7_psw64, LM_GI_GI_7_psw74, LM_GI_GI_7_psw84, LM_GI_GI_7_sw62, LM_GI_GI_7_sw97, LM_GI_GI_8_BR2, LM_GI_GI_8_BS2, LM_GI_GI_8_Parts, LM_GI_GI_8_Playfield, LM_GI_GI_8_psw64, LM_GI_GI_8_psw74, LM_GI_GI_8_psw84, LM_GI_GI_8_sw62, LM_GI_GI_8_sw73, LM_GI_GI_9_BR1, LM_GI_GI_9_BR2, LM_GI_GI_9_BR3, LM_GI_GI_9_BS1, LM_GI_GI_9_BS2, LM_GI_GI_9_Parts, LM_GI_GI_9_Playfield, LM_GI_GI_9_UnderPF, LM_GI_GI_9_g2, LM_GI_GI_9_psw64, LM_GI_GI_9_psw74, LM_GI_GI_9_psw84, LM_GI_GI_9_sw62, LM_GI_GI_9_sw73, LM_GI_GI_9_sw94, LM_L_GI_30_Playfield, LM_L_GI_30_UnderPF, LM_L_li1_Playfield, LM_L_li1_UnderPF, LM_L_li10_Parts, LM_L_li10_Playfield, LM_L_li10_UnderPF, LM_L_li11a_BR1, LM_L_li11a_BS1, LM_L_li11a_Parts, LM_L_li11a_Playfield, LM_L_li11b_BR2, LM_L_li11b_BS2, _
  LM_L_li11b_Parts, LM_L_li11b_Playfield, LM_L_li11b_sw73, LM_L_li11c_BR1, LM_L_li11c_BR3, LM_L_li11c_BS3, LM_L_li11c_Parts, LM_L_li11c_Playfield, LM_L_li15_Playfield, LM_L_li15_UnderPF, LM_L_li16_Playfield, LM_L_li16_UnderPF, LM_L_li17_Parts, LM_L_li17_Playfield, LM_L_li17_UnderPF, LM_L_li18_Parts, LM_L_li18_Playfield, LM_L_li18_UnderPF, LM_L_li18_psw64, LM_L_li18_psw74, LM_L_li18_psw84, LM_L_li19_Playfield, LM_L_li19_UnderPF, LM_L_li19_psw65, LM_L_li19_psw75, LM_L_li19_psw85, LM_L_li2_Playfield, LM_L_li2_UnderPF, LM_L_li20_Playfield, LM_L_li20_UnderPF, LM_L_li22_Parts, LM_L_li22_Playfield, LM_L_li22_UnderPF, LM_L_li23_Parts, LM_L_li23_Playfield, LM_L_li23_UnderPF, LM_L_li24_Parts, LM_L_li24_Playfield, LM_L_li24_UnderPF, LM_L_li24b_Parts, LM_L_li24b_Playfield, LM_L_li24b_UnderPF, LM_L_li25_Parts, LM_L_li25_Playfield, LM_L_li25_UnderPF, LM_L_li26_Parts, LM_L_li26_Playfield, LM_L_li26_UnderPF, LM_L_li27_Parts, LM_L_li27_Playfield, LM_L_li27_UnderPF, LM_L_li27a_Parts, LM_L_li27a_Playfield, LM_L_li27a_UnderPF, _
  LM_L_li3_Playfield, LM_L_li3_UnderPF, LM_L_li33_Playfield, LM_L_li33_UnderPF, LM_L_li34_Playfield, LM_L_li34_UnderPF, LM_L_li36_Playfield, LM_L_li36_UnderPF, LM_L_li37_Playfield, LM_L_li37_UnderPF, LM_L_li38_Playfield, LM_L_li38_UnderPF, LM_L_li39_Playfield, LM_L_li39_UnderPF, LM_L_li4_Parts, LM_L_li4_Playfield, LM_L_li4_UnderPF, LM_L_li4_sw73, LM_L_li40_Playfield, LM_L_li40_UnderPF, LM_L_li41_Playfield, LM_L_li41_UnderPF, LM_L_li5_Parts, LM_L_li5_Playfield, LM_L_li5_UnderPF, LM_L_li6_Playfield, LM_L_li6_UnderPF, LM_L_li7_Playfield, LM_L_li7_UnderPF, LM_L_li8_Parts, LM_L_li8_Playfield, LM_L_li8_UnderPF, LM_L_li9_Parts, LM_L_li9_Playfield, LM_L_li9_UnderPF)
Dim BG_All: BG_All=Array(BM_BR1, BM_BR2, BM_BR3, BM_BS1, BM_BS2, BM_BS3, BM_Gate1, BM_Gate3, BM_PLF, BM_PLFup, BM_PRF, BM_PRFup, BM_Parts, BM_PincabRails, BM_Playfield, BM_UnderPF, BM_WireGate4, BM_g1, BM_g2, BM_g3, BM_psw64, BM_psw65, BM_psw74, BM_psw75, BM_psw84, BM_psw85, BM_sw61a, BM_sw61b, BM_sw61c, BM_sw61d, BM_sw62, BM_sw63, BM_sw66, BM_sw67, BM_sw73, BM_sw77, BM_sw80, BM_sw81, BM_sw82, BM_sw83, BM_sw87, BM_sw94, BM_sw97, LM_GI_GI_1_PLF, LM_GI_GI_1_PLFup, LM_GI_GI_1_Parts, LM_GI_GI_1_Playfield, LM_GI_GI_10_BR1, LM_GI_GI_10_BR2, LM_GI_GI_10_BR3, LM_GI_GI_10_BS1, LM_GI_GI_10_BS2, LM_GI_GI_10_BS3, LM_GI_GI_10_Parts, LM_GI_GI_10_Playfield, LM_GI_GI_10_UnderPF, LM_GI_GI_10_psw64, LM_GI_GI_10_psw74, LM_GI_GI_10_psw75, LM_GI_GI_10_psw84, LM_GI_GI_10_sw62, LM_GI_GI_10_sw73, LM_GI_GI_10_sw94, LM_GI_GI_11_BR1, LM_GI_GI_11_BR2, LM_GI_GI_11_BR3, LM_GI_GI_11_BS1, LM_GI_GI_11_BS2, LM_GI_GI_11_BS3, LM_GI_GI_11_Parts, LM_GI_GI_11_Playfield, LM_GI_GI_11_WireGate4, LM_GI_GI_11_psw65, LM_GI_GI_11_psw74, _
  LM_GI_GI_11_psw75, LM_GI_GI_11_psw85, LM_GI_GI_11_sw63, LM_GI_GI_11_sw73, LM_GI_GI_11_sw94, LM_GI_GI_12_BR1, LM_GI_GI_12_BR2, LM_GI_GI_12_BR3, LM_GI_GI_12_BS1, LM_GI_GI_12_BS2, LM_GI_GI_12_BS3, LM_GI_GI_12_Parts, LM_GI_GI_12_Playfield, LM_GI_GI_12_psw65, LM_GI_GI_12_psw75, LM_GI_GI_12_psw85, LM_GI_GI_12_sw63, LM_GI_GI_13_BR3, LM_GI_GI_13_BS3, LM_GI_GI_13_Parts, LM_GI_GI_13_Playfield, LM_GI_GI_13_psw65, LM_GI_GI_13_psw75, LM_GI_GI_13_psw85, LM_GI_GI_13_sw63, LM_GI_GI_13_sw87, LM_GI_GI_13_sw94, LM_GI_GI_14_Parts, LM_GI_GI_14_Playfield, LM_GI_GI_14_UnderPF, LM_GI_GI_14_g1, LM_GI_GI_14_g2, LM_GI_GI_14_sw73, LM_GI_GI_15_BR1, LM_GI_GI_15_BR2, LM_GI_GI_15_BR3, LM_GI_GI_15_BS1, LM_GI_GI_15_BS2, LM_GI_GI_15_BS3, LM_GI_GI_15_Parts, LM_GI_GI_15_Playfield, LM_GI_GI_15_UnderPF, LM_GI_GI_15_WireGate4, LM_GI_GI_15_g1, LM_GI_GI_15_g2, LM_GI_GI_15_g3, LM_GI_GI_15_psw65, LM_GI_GI_15_psw74, LM_GI_GI_15_psw75, LM_GI_GI_15_psw85, LM_GI_GI_15_sw63, LM_GI_GI_15_sw73, LM_GI_GI_17_BS1, LM_GI_GI_17_BS2, LM_GI_GI_17_Parts, _
  LM_GI_GI_17_Playfield, LM_GI_GI_17_UnderPF, LM_GI_GI_17_g1, LM_GI_GI_17_g2, LM_GI_GI_17_g3, LM_GI_GI_17_sw73, LM_GI_GI_17_sw81, LM_GI_GI_18_Parts, LM_GI_GI_18_Playfield, LM_GI_GI_18_UnderPF, LM_GI_GI_18_g1, LM_GI_GI_18_g2, LM_GI_GI_18_g3, LM_GI_GI_18_sw81, LM_GI_GI_18_sw82, LM_GI_GI_19_Gate1, LM_GI_GI_19_Parts, LM_GI_GI_19_Playfield, LM_GI_GI_19_g1, LM_GI_GI_19_g2, LM_GI_GI_19_g3, LM_GI_GI_2_PRF, LM_GI_GI_2_PRFup, LM_GI_GI_2_Parts, LM_GI_GI_2_Playfield, LM_GI_GI_20_BR1, LM_GI_GI_20_BS1, LM_GI_GI_20_BS3, LM_GI_GI_20_Gate3, LM_GI_GI_20_Parts, LM_GI_GI_20_Playfield, LM_GI_GI_20_UnderPF, LM_GI_GI_20_WireGate4, LM_GI_GI_20_sw66, LM_GI_GI_20_sw83, LM_GI_GI_22_Parts, LM_GI_GI_22_Playfield, LM_GI_GI_22_WireGate4, LM_GI_GI_22_sw66, LM_GI_GI_23_Gate3, LM_GI_GI_23_Parts, LM_GI_GI_23_Playfield, LM_GI_GI_23_sw83, LM_GI_GI_24_Parts, LM_GI_GI_24_Playfield, LM_GI_GI_24_sw81, LM_GI_GI_24_sw82, LM_GI_GI_25_Parts, LM_GI_GI_25_Playfield, LM_GI_GI_25_sw80, LM_GI_GI_25_sw82, LM_GI_GI_26_Parts, LM_GI_GI_26_Playfield, _
  LM_GI_GI_26_sw80, LM_GI_GI_26_sw83, LM_GI_GI_27_Parts, LM_GI_GI_27_UnderPF, LM_GI_GI_27_g1, LM_GI_GI_27_g2, LM_GI_GI_27_g3, LM_GI_GI_27_sw73, LM_GI_GI_28_Parts, LM_GI_GI_28_UnderPF, LM_GI_GI_28_g1, LM_GI_GI_28_g2, LM_GI_GI_28_g3, LM_GI_GI_28_sw73, LM_GI_GI_29_Parts, LM_GI_GI_29_UnderPF, LM_GI_GI_29_g1, LM_GI_GI_29_g2, LM_GI_GI_29_g3, LM_GI_GI_3_Parts, LM_GI_GI_3_Playfield, LM_GI_GI_3_sw61a, LM_GI_GI_3_sw61b, LM_GI_GI_4_Parts, LM_GI_GI_4_Playfield, LM_GI_GI_4_sw61b, LM_GI_GI_5_Parts, LM_GI_GI_5_Playfield, LM_GI_GI_5_sw61c, LM_GI_GI_6_Parts, LM_GI_GI_6_Playfield, LM_GI_GI_6_sw61c, LM_GI_GI_6_sw61d, LM_GI_GI_7_BR1, LM_GI_GI_7_BR2, LM_GI_GI_7_BS2, LM_GI_GI_7_Parts, LM_GI_GI_7_Playfield, LM_GI_GI_7_psw64, LM_GI_GI_7_psw74, LM_GI_GI_7_psw84, LM_GI_GI_7_sw62, LM_GI_GI_7_sw97, LM_GI_GI_8_BR2, LM_GI_GI_8_BS2, LM_GI_GI_8_Parts, LM_GI_GI_8_Playfield, LM_GI_GI_8_psw64, LM_GI_GI_8_psw74, LM_GI_GI_8_psw84, LM_GI_GI_8_sw62, LM_GI_GI_8_sw73, LM_GI_GI_9_BR1, LM_GI_GI_9_BR2, LM_GI_GI_9_BR3, LM_GI_GI_9_BS1, LM_GI_GI_9_BS2, _
  LM_GI_GI_9_Parts, LM_GI_GI_9_Playfield, LM_GI_GI_9_UnderPF, LM_GI_GI_9_g2, LM_GI_GI_9_psw64, LM_GI_GI_9_psw74, LM_GI_GI_9_psw84, LM_GI_GI_9_sw62, LM_GI_GI_9_sw73, LM_GI_GI_9_sw94, LM_L_GI_30_Playfield, LM_L_GI_30_UnderPF, LM_L_li1_Playfield, LM_L_li1_UnderPF, LM_L_li10_Parts, LM_L_li10_Playfield, LM_L_li10_UnderPF, LM_L_li11a_BR1, LM_L_li11a_BS1, LM_L_li11a_Parts, LM_L_li11a_Playfield, LM_L_li11b_BR2, LM_L_li11b_BS2, LM_L_li11b_Parts, LM_L_li11b_Playfield, LM_L_li11b_sw73, LM_L_li11c_BR1, LM_L_li11c_BR3, LM_L_li11c_BS3, LM_L_li11c_Parts, LM_L_li11c_Playfield, LM_L_li15_Playfield, LM_L_li15_UnderPF, LM_L_li16_Playfield, LM_L_li16_UnderPF, LM_L_li17_Parts, LM_L_li17_Playfield, LM_L_li17_UnderPF, LM_L_li18_Parts, LM_L_li18_Playfield, LM_L_li18_UnderPF, LM_L_li18_psw64, LM_L_li18_psw74, LM_L_li18_psw84, LM_L_li19_Playfield, LM_L_li19_UnderPF, LM_L_li19_psw65, LM_L_li19_psw75, LM_L_li19_psw85, LM_L_li2_Playfield, LM_L_li2_UnderPF, LM_L_li20_Playfield, LM_L_li20_UnderPF, LM_L_li22_Parts, LM_L_li22_Playfield, _
  LM_L_li22_UnderPF, LM_L_li23_Parts, LM_L_li23_Playfield, LM_L_li23_UnderPF, LM_L_li24_Parts, LM_L_li24_Playfield, LM_L_li24_UnderPF, LM_L_li24b_Parts, LM_L_li24b_Playfield, LM_L_li24b_UnderPF, LM_L_li25_Parts, LM_L_li25_Playfield, LM_L_li25_UnderPF, LM_L_li26_Parts, LM_L_li26_Playfield, LM_L_li26_UnderPF, LM_L_li27_Parts, LM_L_li27_Playfield, LM_L_li27_UnderPF, LM_L_li27a_Parts, LM_L_li27a_Playfield, LM_L_li27a_UnderPF, LM_L_li3_Playfield, LM_L_li3_UnderPF, LM_L_li33_Playfield, LM_L_li33_UnderPF, LM_L_li34_Playfield, LM_L_li34_UnderPF, LM_L_li36_Playfield, LM_L_li36_UnderPF, LM_L_li37_Playfield, LM_L_li37_UnderPF, LM_L_li38_Playfield, LM_L_li38_UnderPF, LM_L_li39_Playfield, LM_L_li39_UnderPF, LM_L_li4_Parts, LM_L_li4_Playfield, LM_L_li4_UnderPF, LM_L_li4_sw73, LM_L_li40_Playfield, LM_L_li40_UnderPF, LM_L_li41_Playfield, LM_L_li41_UnderPF, LM_L_li5_Parts, LM_L_li5_Playfield, LM_L_li5_UnderPF, LM_L_li6_Playfield, LM_L_li6_UnderPF, LM_L_li7_Playfield, LM_L_li7_UnderPF, LM_L_li8_Parts, LM_L_li8_Playfield, _
  LM_L_li8_UnderPF, LM_L_li9_Parts, LM_L_li9_Playfield, LM_L_li9_UnderPF)
' VLM  Arrays - End


' Called when options are tweaked by the player.
' - 0: game has started, good time to load options and adjust accordingly
' - 1: an option has changed
' - 2: options have been reseted
' - 3: player closed the tweak UI, good time to update staticly prerendered parts
' Table1.Option arguments are:
' - option name, minimum value, maximum value, step between valid values, default value, unit (0=None, 1=Percent), an optional arry of literal strings

Dim dspTriggered : dspTriggered = False
Dim VR_Obj, VRRoom


Sub Table1_OptionEvent(ByVal eventId)
    If eventId = 1 And Not dspTriggered Then dspTriggered = True : DisableStaticPreRendering = True : End If

' VR Room
  VRRoom = Table1.Option("VR Room", 1, 3, 1, 3, 0, Array("Basti's ROOM", "Mixed Reality", "Pool Bar"))
  SetupVRRoom

' Ambient
    Ambience = Table1.Option("Ambience", 1, 2, 1, 2, 0, Array("ON", "OFF"))
  SetupAmbience

' Music Background
    MusicChoice = Table1.Option("Background Music", 1, 2, 1, 2, 0, Array("ON", "OFF"))
  SetupMusic

  ' Color Saturation
    ColorLUT = Table1.Option("Color Saturation", 1, 11, 1, 2, 0, _
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

    If eventId = 3 And dspTriggered Then dspTriggered = False : DisableStaticPreRendering = False : End If
End Sub

Sub SetupVRRoom()
  Dim BP
  If RenderingMode = 2 Then
      BM_Parts.image = "VLM.Nestmap0"
      VRMode = True
      For Each x in aReels : x.Visible = 0 : Next
      BM_PincabRails.visible = 0
    If VRRoom = 3 Then
      For Each VR_Obj in VR_MinimalRoom : VR_Obj.Visible = 0 : Next
      For Each VR_Obj in VR_PoolBar : VR_Obj.Visible = 1 : Next
      For Each VR_Obj in VR_Sphere : VR_Obj.Visible = 0 : Next
    ElseIf VRRoom = 2 Then
      For Each VR_Obj in VR_MinimalRoom : VR_Obj.Visible = 0 : Next
      For Each VR_Obj in VR_PoolBar : VR_Obj.Visible = 0 : Next
      For Each VR_Obj in VR_Sphere : VR_Obj.Visible = 1 : Next
    Else
      For Each VR_Obj in VR_MinimalRoom : VR_Obj.Visible = 1 : Next
      For Each VR_Obj in VR_PoolBar : VR_Obj.Visible = 0 : Next
      For Each VR_Obj in VR_Sphere : VR_Obj.Visible = 0 : Next
    End If
  Else
    VRRoom = 0
    VRMode = False
    BM_Parts.image = "VLM.Nestmap3"
    For Each VR_Obj in VR_MinimalRoom : VR_Obj.Visible = 0 : Next
    For Each VR_Obj in VR_PoolBar : VR_Obj.Visible = 0 : Next
    For Each VR_Obj in VR_Sphere : VR_Obj.Visible = 0 : Next
    For Each VR_Obj in VR_Table : VR_Obj.Visible = 0 : Next
    If DesktopMode = True Then For Each BP in BP_PincabRails : BP.Visible = 1: Next
    End If
End Sub

' ****************************************************

'**********
' Digital Clock
'**********


Sub ClockTimer_Timer()

    ' ROB AND WALTER'S Digital Clock below**************************************
  dim n
    n=Hour(now) MOD 12
    if n = 0 then n = 12
  hour1.imagea="digit" & CStr(n \ 10)
    hour2.imagea="digit" & CStr(n mod 10)
  n=Minute(now)
  minute1.imagea="digit" & CStr(n \ 10)
    minute2.imagea="digit" & CStr(n mod 10)
  'n=Second(now)
  'second1.imagea="digit" & CStr(n \ 10)
    'second2.imagea="digit" & CStr(n mod 10)
End Sub


'******
' TV Timer
'******

Const TVCounterMax = 16

TimerTV.enabled = True

Dim TVCounter: TVCounter = 1

Sub TimerTV_Timer()

if TVCounter < 10 then
        TV.Image = "ezgif-frame-" & "00" & TVCounter
    elseif TVCounter < 16 then
        TV.Image = "ezgif-frame-" & "0" & TVCounter
    else'if TVCounter < 150 then
        'TV.Image = "ezgif-frame-" & "" & TVCounter
    'else
        TV.image = "ezgif-frame-" & TVCounter
     end if

    TVCounter = TVCounter + 1

    If TVCounter > TVCounterMax Then
        TVCounter = 1
   End If

End Sub

'****************
'Fan Animation'
'****************

FanTimer.enabled = True

sub Fantimer_timer()
Fan.Objrotz = Fan.Objrotz + 2
End Sub

'*********************************************************************************************************
' Song & Music
'*********************************************************************************************************


Dim Ambience : Ambience = 1       '1 - ON, 2 - OFF

Sub Setupambience()
If Ambience = 1 then  PlaySound"Ambience"
If Ambience = 2 then  StopSound"Ambience"

End Sub

Dim numSongs : numSongs = 40
Dim musicNum : musicNum = Int(rnd*numSongs) + 1

Sub PlayNextSong()
    EndMusic
    musicNum = musicNum + 1
    If musicNum > numSongs Then musicNum = 1  ' wrap around
    If MusicChoice = 1 Then PlayMusic "Music80\Song" & musicNum & ".mp3"
End Sub

Sub PlayPreviousSong()
    EndMusic
    musicNum = musicNum - 1
    If musicNum < 1 Then musicNum = numSongs  ' wrap around
    If MusicChoice = 1 Then PlayMusic "Music80\Song" & musicNum & ".mp3"
End Sub


Dim Musicchoice : Musicchoice = 1       '1 - ON, 2 - OFF

Sub Setupmusic()
  If Musicchoice = 1 then PlayMusic "Music80\Song" & Int(rnd*numSongs) + 1 & ".mp3"
  If Musicchoice = 2 then EndMusic

End Sub


' VR PLUNGER ANIMATION
'
' Code needed to animate the plunger. If you pull the plunger it will move in VR.
' IMPORTANT: there are two numeric values in the code that define the postion of the plunger and the
' range in which it can move. The first numeric value is the actual y position of the plunger primitive
' and the second is the actual y position + 100 to determine the range in which it can move.
'
' You need to to select the VR_Primary_plunger primitive you copied from the
' template and copy the value of the Y position
' (e.g. 2130) into the code. The value that determines the range of the plunger is always the y
' position + 100 (e.g. 2230).
'

Sub TimerPlunger_Timer

  If VR_Primary_plunger.Y < 2291.142 then
      VR_Primary_plunger.Y = VR_Primary_plunger.Y + 5
  End If
End Sub

Sub TimerPlunger2_Timer
  VR_Primary_plunger.Y = 2158.142 + (5* Plunger.Position) -20
End Sub


' ***** Beer Bubble Code - Rawd *****
Sub BeerTimer_Timer()

Randomize(21)
BeerBubble1.z = BeerBubble1.z + Rnd(1)*0.5
if BeerBubble1.z > -771 then BeerBubble1.z = -955
BeerBubble2.z = BeerBubble2.z + Rnd(1)*1
if BeerBubble2.z > -768 then BeerBubble2.z = -955
BeerBubble3.z = BeerBubble3.z + Rnd(1)*1
if BeerBubble3.z > -768 then BeerBubble3.z = -955
BeerBubble4.z = BeerBubble4.z + Rnd(1)*0.75
if BeerBubble4.z > -774 then BeerBubble4.z = -955
BeerBubble5.z = BeerBubble5.z + Rnd(1)*1
if BeerBubble5.z > -771 then BeerBubble5.z = -955
BeerBubble6.z = BeerBubble6.z + Rnd(1)*1
if BeerBubble6.z > -774 then BeerBubble6.z = -955
BeerBubble7.z = BeerBubble7.z + Rnd(1)*0.8
if BeerBubble7.z > -768 then BeerBubble7.z = -955
BeerBubble8.z = BeerBubble8.z + Rnd(1)*1
if BeerBubble8.z > -771 then BeerBubble8.z = -955
End Sub


' ***************** VR Clock code below - THANKS RASCAL ******************
Dim CurrentMinute ' for VR clock
' VR Clock code below....
Sub ClockTimerAnalog_Timer()

    'ClockHands Below
  VR_Clock_minutes.RotAndTra2 = (Minute(Now())+(Second(Now())/100))*6
  VR_Clock_Hours.RotAndTra2 = Hour(Now())*30+(Minute(Now())/2)
      VR_Clock_Seconds.RotAndTra2 = (Second(Now()))*6
  CurrentMinute=Minute(Now())

End Sub




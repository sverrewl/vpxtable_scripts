'Pennant Fever (Williams 1984)
'https://www.ipdb.org/machine.cgi?id=3335

'   MOD by mcarter78 based on vpx 1.0 by armyaviation

'*******************************
'  Revisions
'*******************************

'1.0.1 - mcarter78 - WIP (updated Fleep sounds, physical subway & trough, 3d inserts, new playfield mesh, layers organization)
'1.0.2 - mcarter78 - Pitch & Bat working again with physical trough, fix flipper sounds and continue organizing layers
'1.0.3 - mcarter78 - Make curveballs work, reshape homerun ramp to allow balls to drop through to trough below
'1.0.4 - mcarter78 - 2k batch.  Animate bat flipper, targets, pitch flap.  Relocate drain triggers.

' TODO:
' - 2 player control scheme
' - VR cabinet & environment

Option Explicit
Randomize

Const ballsize = 35.294117647059
Const ballmass = 0.35171992672502

Dim tablewidth: tablewidth = Table1.width
Dim tableheight: tableheight = Table1.height

Const fastballspeed = 25
Const changeuppercent = 0.85
Const magnetstrength = 16

On Error Resume Next
ExecuteGlobal GetTextFile("controller.vbs")
If Err Then MsgBox "You need the controller.vbs in order to run this table, available in the vp10 package"
On Error Goto 0

Const cGameName="pfevr_p3",UseSolenoids=2,UseLamps=1,SSolenoidOn="",SSolenoidOff="", SCoin=""
Const DoubleTapTime = 250   'Max. interval in ms for the Double Tap to register. Please set this to the minimal interval, You can live with
Const DoubleTap = 1

Const tnob = 6 ' total number of balls
Const KeepLogs = True

LoadVPM "03060000","S7.VBS",3.1


'*******************************************
'  ZOPT: User Options
'*******************************************

Dim LightLevel : LightLevel = 0.25        ' Level of room lighting (0 to 1), where 0 is dark and 100 is brightest
Dim ColorLUT : ColorLUT = 1           ' Color desaturation LUTs: 1 to 11, where 1 is normal and 11 is black'n'white
Dim VolumeDial : VolumeDial = 0.8             ' Overall Mechanical sound effect volume. Recommended values should be no greater than 1.
Dim BallRollVolume : BallRollVolume = 0.5     ' Level of ball rolling volume. Value between 0 and 1

' Called when options are tweaked by the player.
' - 0: game has started, good time to load options and adjust accordingly
' - 1: an option has changed
' - 2: options have been reseted
' - 3: player closed the tweak UI, good time to update staticly prerendered parts
' Table1.Option arguments are:
' - option name, minimum value, maximum value, step between valid values, default value, unit (0=None, 1=Percent), an optional arry of literal strings
Sub Table1_OptionEvent(ByVal eventId)
    If eventId = 1 Then DisableStaticPreRendering = True


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

  ' Room brightness
' LightLevel = Table1.Option("Table Brightness (Ambient Light Level)", 0, 1, 0.01, .5, 1)
  LightLevel = NightDay/100
  SetRoomBrightness LightLevel   'Uncomment this line for lightmapped tables.

    If eventId = 3 Then DisableStaticPreRendering = False
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

'*********************************************************************
'Solenoid Call backs
'*********************************************************************

Dim Changeup: Changeup = False

' SolCallback(1)="SolPitch"
SolCallback(2)="Bell"
' SolCallback(3)="SetRelayGI"
SolCallback(4)="vpmNudge.SolGameOn"
SolCallback(5)="SolmMagnet"
SolCallback(6)="SolPitch"
SolCallback(7)="SolSwingBat"

' SolCallback(sLLFlipper) = "SolPitch"

Sub SolSwingBat(Enabled)
    If Enabled Then
        RandomSoundFlipperUpLeft LeftFlipper
        LeftFlipper.RotateToEnd
        vpmTimer.AddTimer 125, "RandomSoundFlipperDownLeft LeftFlipper:LeftFlipper.RotateToStart'"
    End If
  End Sub

Sub SolPitch(Enabled)
    Debug.print("SolPitch called: " & Enabled)
  If Enabled Then
        timer1.enabled=true
      vpmtimer.pulsesw 28
        swTrough1.kick -90, 9
        RandomSoundBallRelease swTrough1
        UpdateTrough
    End If
End Sub

Sub SolmMagnet(Enabled)
    Debug.Print("SolmMagnet called: " & Enabled)
End Sub

'*********************************************************************
'Solenoid Controlled toys
'*********************************************************************

Sub Bell(Enabled)
  If Enabled Then
    PlaySound "bell"
  End If
End Sub

'*********************************************************************
'Initiate Table
'*********************************************************************

Dim mMagnet, gistuff, CalledPitch

Dim PFBall1, PFBall2, PFBall3, PFBall4, PFBall5, PFBall6, gBOT

Sub Table1_Init
  vpmInit Me
  On Error Resume Next
    With Controller
      .GameName = cGameName
      If Err Then MsgBox "Can't start Game" & cGameName & vbNewLine & Err.Description : Exit Sub
      .SplashInfoLine = "Pennant Fever"&chr(13)&""
      .HandleMechanics=0
      .HandleKeyboard=0
      .ShowDMDOnly=1
      .ShowFrame=0
      .ShowTitle=0
      .hidden = 0
      On Error Resume Next
      .Run GetPlayerHWnd
      If Err Then MsgBox Err.Description
      On Error Goto 0
    End With
  On Error Goto 0

    ' 'Trough - Creates a ball in the kicker switch and gives that ball used an individual name.
  Set PFBall1 = swTrough1.CreateSizedballWithMass(Ballsize/2,Ballmass)
    Set PFBall2 = swTrough2.CreateSizedballWithMass(Ballsize/2,Ballmass)
  Set PFBall3 = DrainOF1.CreateSizedballWithMass(Ballsize/2,Ballmass)
    Set PFBall4 = DrainOF2.CreateSizedballWithMass(Ballsize/2,Ballmass)
    Set PFBall5 = DrainHP1.CreateSizedballWithMass(Ballsize/2,Ballmass)
    Set PFBall6 = DrainHP2.CreateSizedballWithMass(Ballsize/2,Ballmass)

  ' '***Setting up a ball array (collection), must contain all the balls you create on the table.
  gBOT = Array(PFBall1, PFBall2, PFBall3, PFBall4, PFBall5, PFBall6)

  'Map all lamps to the corresponding ROM output using the value of TimerInterval of each light object
  vpmMapLights AllLamps     'Make a collection called "AllLamps" and put all the light objects in it.

    For Each gistuff in backgif : gistuff.opacity = 100 : Next
    For Each gistuff in backgip : gistuff.blenddisablelighting = 1 : Next

  PinMAMETimer.Interval=PinMAMEInterval
  PinMAMETimer.Enabled=1
  vpmNudge.TiltSwitch=swTilt
  vpmNudge.Sensitivity=5
  Wall13.visible = 1
  Wall4.visible = 0
  Wall18.visible = 1
  Wall21.visible = 0
  Wall18.sidevisible = 1
  Wall21.sidevisible = 0
End Sub

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
  RollingUpdate       'Update rolling sounds
End Sub


'*********************************************************************
'Player Controls
'*********************************************************************

Sub Table1_KeyDown(ByVal KeyCode)

  If KeyCode = LeftMagnaSave Then
        Controller.Switch(12) = 1
  End If

  If KeyCode = RightFlipperKey Then
    Controller.Switch(10) = 1
  End If

  If keycode = AddCreditKey or keycode = AddCreditKey2 Then
    Select Case Int(rnd*3)
      Case 0: PlaySound ("Coin_In_1"), 0, CoinSoundLevel, 0, 0.25
      Case 1: PlaySound ("Coin_In_2"), 0, CoinSoundLevel, 0, 0.25
      Case 2: PlaySound ("Coin_In_3"), 0, CoinSoundLevel, 0, 0.25
    End Select
  End If

  If keycode = LeftFlipperKey Then
        CalledPitch = RndInt(0, 2)
        If PitcherReady Then
            If CalledPitch = 0 Then
                'Fastball
                vpmTimer.PulseSw 11
            ElseIf CalledPitch = 1 Then
                'Curveball
                vpmTimer.PulseSw 12
            Else
                'Changeup
                vpmTimer.PulseSw 13
            End If
        End If
  End If

    If vpmKeyDown(KeyCode) Then Exit Sub
End Sub

Sub Table1_KeyUp(ByVal KeyCode)
  If KeyCode = RightFlipperKey Then Controller.Switch(10) = 0
  If KeyCode = LeftMagnaSave Then Controller.Switch(12 )= 0
  If vpmKeyUp(KeyCode) Then Exit Sub
End Sub


Sub LeftFlipper_Animate
  Debug.Print("batflipper flip")
  Dim a : a = LeftFlipper.CurrentAngle

  FlipperLSh.RotZ = LeftFlipper.CurrentAngle

  Dim v, BP
  v = 255.0 * (122.0 - LeftFlipper.CurrentAngle) / (122.0 -  74.5)

  For each BP in BP_BatFlipper
    BP.Rotz = a
    BP.visible = v < 128.0
  Next
  For each BP in BP_BatFlipperU
    BP.Rotz = a
    BP.visible = v >= 128.0
  Next
End Sub


'*********************************************************************
'Drain/Trough
'*********************************************************************

Dim KickerBall

Sub KickBall(kball, kangle, kvel, kvelz, kzlift)
  dim rangle
  rangle = PI * (kangle - 90) / 180

  kball.z = kball.z + kzlift
  kball.velz = kvelz
  kball.velx = cos(rangle)*kvel
  kball.vely = sin(rangle)*kvel
End Sub

Sub UpdateTrough()
  UpdateTroughTimer.Interval = 100
  UpdateTroughTimer.Enabled = 1
End Sub

Sub UpdateTroughTimer_Timer()
  If swTrough1.BallCntOver = 0 Then swTrough2.kick -90, 9
  If swTrough2.BallCntOver = 0 And DrainOF1.BallCntOver = 1 Then
        DrainOF1.kick 180, 9
    Else
        DrainHP1.kick 0, 9
    End If
    If DrainOF1.BallCntOver = 0 Then DrainOF2.kick 180, 9
    If DrainHP1.BallCntOver = 0 Then DrainHP2.kick 0, 9
  UpdateTroughTimer.Enabled = 0
End Sub

Sub swTrough1_Hit():UpdateTrough:End Sub
Sub swTrough1_UnHit():UpdateTrough:End Sub
Sub swTrough2_Hit():UpdateTrough:End Sub
Sub swTrough2_UnHit():UpdateTrough:End Sub
Sub swDrainOF1_Hit():UpdateTrough:End Sub
Sub swDrainOF1_UnHit():UpdateTrough:End Sub
Sub swDrainOF2_Hit():UpdateTrough:End Sub
Sub swDrainOF2_UnHit():UpdateTrough:End Sub
Sub swDrainHP1_Hit():UpdateTrough:End Sub
Sub swDrainHP1_UnHit():UpdateTrough:End Sub
Sub swDrainHP2_Hit():UpdateTrough:End Sub
Sub swDrainHP2_UnHit():UpdateTrough:End Sub

Sub DrainTriggerHP_Hit()
    Debug.Print("Homeplate drain")
    vpmTimer.PulseSw 9
    PitcherReady = True
End Sub

Sub DrainTriggerOF_Hit()
    Debug.Print("Outfield drain")
    vpmTimer.PulseSw 15
    PitcherReady = True
End Sub


'*********************************************************************
'Targets moving
'*********************************************************************

Sub Gate1_Animate
  Dim a : a = Gate1.CurrentAngle
  Dim BL : For Each BL in BP_TargetSingleLeft_001 : BL.rotx = -a: Next
End Sub

Sub Gate2_Animate
  Dim a : a = Gate2.CurrentAngle
  Dim BL : For Each BL in BP_TargetOutLeft_001 : BL.rotx = -a: Next
End Sub

Sub Gate3_Animate
  Dim a : a = Gate3.CurrentAngle
  Dim BL : For Each BL in BP_TargetDoubleLeft_001 : BL.rotx = -a: Next
End Sub

Sub Gate4_Animate
  Dim a : a = Gate4.CurrentAngle
  Dim BL : For Each BL in BP_TargetTriple_001 : BL.rotx = -a: Next
End Sub

Sub Gate5_Animate
  Dim a : a = Gate5.CurrentAngle
  Dim BL : For Each BL in BP_TargetDoubleRight_001 : BL.rotx = -a: Next
End Sub

Sub Gate6_Animate
  Dim a : a = Gate6.CurrentAngle
  Dim BL : For Each BL in BP_TargetOutRight_001 : BL.rotx = -a: Next
End Sub

Sub Gate7_Animate
  Dim a : a = Gate7.CurrentAngle
  Dim BL : For Each BL in BP_TargetSingleRight_001 : BL.rotx = -a: Next
End Sub

'*********************************************************************
'Targets
'*********************************************************************

Sub Gate1_hit:vpmTimer.PulseSw 17:End Sub
Sub Gate2_hit:vpmTimer.PulseSw 18:End Sub
Sub Gate3_hit:vpmTimer.PulseSw 19:End Sub
Sub Gate4_hit:vpmTimer.PulseSw 20:End Sub
Sub Gate5_hit:vpmTimer.PulseSw 21:End Sub
Sub Gate6_hit:vpmTimer.PulseSw 22:End Sub
Sub Gate7_hit:vpmTimer.PulseSw 23:End Sub


' Home Run triggers
Sub trigger1_hit
  vpmTimer.PulseSw 24
  If Activeball.vely<-10 then
    Activeball.vely=2
  end if
End Sub'

Sub trigger2_hit
  vpmTimer.PulseSw 25
  If Activeball.vely<-10 then
    Activeball.vely=2
  end if
End Sub

Sub trigger3_hit
  vpmTimer.PulseSw 26
  If Activeball.vely<-10 then
    Activeball.vely=2
  end if
End Sub

'*********************************************************************
'Pitcher Cover
'*********************************************************************

Dim PitcherReady : PitcherReady = True

Sub Timer1_timer
  timer2.enabled=false

  Dim BP

  If BM_PitchFlap.objrotx < 25 Then
    For Each BP in BP_PitchFlap
      BP.objrotx = BP.objrotx + 5
    Next
    ' BM_PitchFlap.objrotx = BM_PitchFlap.objrotx + 5
    Controller.Switch(29) = 0
    Else
        timer1.enabled=false
      timer2.enabled=true
    End If
End Sub

Sub Timer2_timer
  Dim BP
  If BM_PitchFlap.objrotx > 0 Then
    For Each BP in BP_PitchFlap
      BP.objrotx = BP.objrotx - 5
    Next
    ' BM_PitchFlap.objrotx = BM_PitchFlap.objrotx - 5
  Else
    Controller.Switch(29) = 1
  End If
End Sub

Sub swTrough0_hit
    Set KickerBall = ActiveBall
  timer1.enabled=true
  vpmtimer.pulsesw 28

    KickerBall.x = swTrough0.x
    KickBall KickerBall, 180, fastballspeed, 0, 195
    KickerBall.velX = 0
    If CalledPitch > 0 Then
        Debug.Print("Changeup/Curveball")
        KickerBall.velY = KickerBall.velY * changeuppercent
    End If
    Controller.Switch(11) = 0
    Controller.Switch(13) = 0
    PitcherReady = False
End sub

Sub swTrough0_UnHit
    UpdateTrough
End Sub

Sub Magnet_Hit
    Debug.Print("Magnet Hit: " & ActiveBall.velX)
    If CalledPitch = 1 Then
        ActiveBall.velY = ActiveBall.velY - 1
        ActiveBall.velX = 3
    End If
End Sub

'*********************************************************************
'Prototype Cabinet & Plastics
'*********************************************************************

Sub ShowPrototype(choice)
    If choice = 0 Then
        wall8.image = "Black_Plastics"
        wall7.image = "Black_Plastics"
        Primitive48.image = ""
        batflipper. image = "ash"
        Wall13.visible = 1
        Wall4.visible = 0
        Wall18.visible = 1
        Wall21.visible = 0
        Wall18.sidevisible = 1
        Wall21.sidevisible = 0
        Primitive30.image =""
    Else
        wall8.image = "Green_Plastics"
        wall7.image = "Green_Plastics"
        Primitive48.image = "Green"
        batflipper. image = "blackwood"
        Primitive30.image ="Green"
        Wall13.visible = 0
        Wall4.visible = 1
        Wall18.visible = 0
        Wall21.visible = 1
        Wall18.sidevisible = 0
        Wall21.sidevisible = 1
    End If
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
  Dim b

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


Sub Glass_hit:PlaySoundAtBall "fx_glass":End Sub
Sub Ramp2_hit:PlaySoundAtBall "fx_glass":End Sub


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
TargetSoundFactor = 0.0025 * 100   'volume multiplier; must not be zero
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
  PitchPlayfieldRoll = BallVel(ball) ^ 2 * 35
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

'*****************************************************************************************************************************************
'   ZLOG: ERROR LOGS by baldgeek
'*****************************************************************************************************************************************

' Log File Usage:
'   WriteToLog "Label 1", "Message 1 "
'   WriteToLog "Label 2", "Message 2 "

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
    & LZ(Day(CurrTime), 2) & " " _
    & LZ(Hour(CurrTime),   2) & ":" _
    & LZ(Minute(CurrTime), 2) & ":" _
    & LZ(Second(CurrTime), 2) & ":" _
    & LZ(MilliSecs, 4)
  End Function

  ' *** Debug.Print the time with milliseconds, and a message of your choice
  Public Sub WriteToLog(label, message, code)
    Dim FormattedMsg, Timestamp
    '   Filename = UserDirectory + "\" + cGameName + "_debug_log.txt"
    Filename = cGameName + "_debug_log.txt"

    Set TxtFileStream = CreateObject("Scripting.FileSystemObject").OpenTextFile(Filename, code, True)
    Timestamp = GetTimeStamp
    FormattedMsg = GetTimeStamp + " : " + label + " : " + message
    TxtFileStream.WriteLine FormattedMsg
    TxtFileStream.Close
    Debug.print label & " : " & message
  End Sub
End Class

Sub WriteToLog(label, message)
  If KeepLogs Then
    Dim LogFileObj
    Set LogFileObj = New DebugLogFile
    LogFileObj.WriteToLog label, message, 8
  End If
End Sub

Sub NewLog()
  If KeepLogs Then
    Dim LogFileObj
    Set LogFileObj = New DebugLogFile
    LogFileObj.WriteToLog "NEW Log", " ", 2
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

' VLM  Arrays - Start
' Arrays per baked part
Dim BP_BatFlipper: BP_BatFlipper=Array(BM_BatFlipper, LM_GI_BatFlipper, LM_L_Light39_BatFlipper)
Dim BP_BatFlipperU: BP_BatFlipperU=Array(BM_BatFlipperU, LM_GI_BatFlipperU, LM_L_Light39_BatFlipperU, LM_L_L40_BatFlipperU, LM_L_Light9_BatFlipperU)
Dim BP_Insert___Arrow_2_000: BP_Insert___Arrow_2_000=Array(BM_Insert___Arrow_2_000, LM_L_L51_Insert___Arrow_2_000)
Dim BP_Insert___Arrow_2_001: BP_Insert___Arrow_2_001=Array(BM_Insert___Arrow_2_001, LM_L_L52_Insert___Arrow_2_001, LM_L_L6_Insert___Arrow_2_001)
Dim BP_Insert___Arrow_2_002: BP_Insert___Arrow_2_002=Array(BM_Insert___Arrow_2_002, LM_L_L53_Insert___Arrow_2_002)
Dim BP_Insert___Circle___Clear_001: BP_Insert___Circle___Clear_001=Array(BM_Insert___Circle___Clear_001, LM_L_Light23_Insert___Circle___, LM_L_L24_Insert___Circle___Clea, LM_L_Light25_Insert___Circle___)
Dim BP_Insert___Circle___Clear_002: BP_Insert___Circle___Clear_002=Array(BM_Insert___Circle___Clear_002, LM_L_Light15_Insert___Circle___, LM_L_L16_Insert___Circle___Clea, LM_L_Light17_Insert___Circle___)
Dim BP_Insert___Circle___Clear_003: BP_Insert___Circle___Clear_003=Array(BM_Insert___Circle___Clear_003, LM_GI_Insert___Circle___Clear_0, LM_L_Light39_Insert___Circle___, LM_L_L40_Insert___Circle___Clea)
Dim BP_Insert___Circle___Clear_004: BP_Insert___Circle___Clear_004=Array(BM_Insert___Circle___Clear_004, LM_L_L31_Insert___Circle___Clea, LM_L_L32_Insert___Circle___Clea, LM_L_Light33_Insert___Circle___)
Dim BP_Insert___Circle___Starbus_023: BP_Insert___Circle___Starbus_023=Array(BM_Insert___Circle___Starbus_02, LM_L_Light36_Insert___Circle___, LM_L_Light37_Insert___Circle___, LM_L_Light38_Insert___Circle___)
Dim BP_Insert___Circle___Starbus_024: BP_Insert___Circle___Starbus_024=Array(BM_Insert___Circle___Starbus_02, LM_L_Light36_Insert___Circle___, LM_L_Light37_Insert___Circle___)
Dim BP_Insert___Circle___Starbust_000: BP_Insert___Circle___Starbust_000=Array(BM_Insert___Circle___Starbust_0, LM_L_L30_Insert___Circle___Star, LM_L_L31_Insert___Circle___Star, LM_L_L32_Insert___Circle___Star)
Dim BP_Insert___Circle___Starbust_001: BP_Insert___Circle___Starbust_001=Array(BM_Insert___Circle___Starbust_0, LM_L_L29_Insert___Circle___Star, LM_L_L30_Insert___Circle___Star, LM_L_L31_Insert___Circle___Star)
Dim BP_Insert___Circle___Starbust_002: BP_Insert___Circle___Starbust_002=Array(BM_Insert___Circle___Starbust_0, LM_L_Light28_Insert___Circle___, LM_L_L29_Insert___Circle___Star, LM_L_L30_Insert___Circle___Star)
Dim BP_Insert___Circle___Starbust_003: BP_Insert___Circle___Starbust_003=Array(BM_Insert___Circle___Starbust_0, LM_L_Light27_Insert___Circle___, LM_L_Light28_Insert___Circle___, LM_L_L29_Insert___Circle___Star)
Dim BP_Insert___Circle___Starbust_004: BP_Insert___Circle___Starbust_004=Array(BM_Insert___Circle___Starbust_0, LM_L_Light26_Insert___Circle___, LM_L_Light27_Insert___Circle___, LM_L_Light28_Insert___Circle___)
Dim BP_Insert___Circle___Starbust_005: BP_Insert___Circle___Starbust_005=Array(BM_Insert___Circle___Starbust_0, LM_L_Light25_Insert___Circle___, LM_L_Light26_Insert___Circle___, LM_L_Light27_Insert___Circle___)
Dim BP_Insert___Circle___Starbust_006: BP_Insert___Circle___Starbust_006=Array(BM_Insert___Circle___Starbust_0, LM_L_Light25_Insert___Circle___, LM_L_Light26_Insert___Circle___)
Dim BP_Insert___Circle___Starbust_007: BP_Insert___Circle___Starbust_007=Array(BM_Insert___Circle___Starbust_0, LM_L_Light22_Insert___Circle___, LM_L_Light23_Insert___Circle___, LM_L_L24_Insert___Circle___Star)
Dim BP_Insert___Circle___Starbust_008: BP_Insert___Circle___Starbust_008=Array(BM_Insert___Circle___Starbust_0, LM_L_Light21_Insert___Circle___, LM_L_Light22_Insert___Circle___, LM_L_Light23_Insert___Circle___)
Dim BP_Insert___Circle___Starbust_009: BP_Insert___Circle___Starbust_009=Array(BM_Insert___Circle___Starbust_0, LM_L_Light20_Insert___Circle___, LM_L_Light21_Insert___Circle___)
Dim BP_Insert___Circle___Starbust_010: BP_Insert___Circle___Starbust_010=Array(BM_Insert___Circle___Starbust_0, LM_L_Light19_Insert___Circle___, LM_L_Light20_Insert___Circle___, LM_L_Light21_Insert___Circle___)
Dim BP_Insert___Circle___Starbust_011: BP_Insert___Circle___Starbust_011=Array(BM_Insert___Circle___Starbust_0, LM_L_Light18_Insert___Circle___, LM_L_Light19_Insert___Circle___, LM_L_Light20_Insert___Circle___)
Dim BP_Insert___Circle___Starbust_012: BP_Insert___Circle___Starbust_012=Array(BM_Insert___Circle___Starbust_0, LM_L_Light17_Insert___Circle___, LM_L_Light18_Insert___Circle___, LM_L_Light19_Insert___Circle___)
Dim BP_Insert___Circle___Starbust_013: BP_Insert___Circle___Starbust_013=Array(BM_Insert___Circle___Starbust_0, LM_L_L16_Insert___Circle___Star, LM_L_Light17_Insert___Circle___, LM_L_Light18_Insert___Circle___)
Dim BP_Insert___Circle___Starbust_014: BP_Insert___Circle___Starbust_014=Array(BM_Insert___Circle___Starbust_0, LM_L_Light14_Insert___Circle___, LM_L_Light15_Insert___Circle___, LM_L_L16_Insert___Circle___Star)
Dim BP_Insert___Circle___Starbust_015: BP_Insert___Circle___Starbust_015=Array(BM_Insert___Circle___Starbust_0, LM_L_Light13_Insert___Circle___, LM_L_Light14_Insert___Circle___, LM_L_Light15_Insert___Circle___)
Dim BP_Insert___Circle___Starbust_016: BP_Insert___Circle___Starbust_016=Array(BM_Insert___Circle___Starbust_0, LM_L_Light12_Insert___Circle___, LM_L_Light13_Insert___Circle___)
Dim BP_Insert___Circle___Starbust_017: BP_Insert___Circle___Starbust_017=Array(BM_Insert___Circle___Starbust_0, LM_L_Light12_Insert___Circle___)
Dim BP_Insert___Circle___Starbust_018: BP_Insert___Circle___Starbust_018=Array(BM_Insert___Circle___Starbust_0, LM_L_Light10_Insert___Circle___, LM_L_Light11_Insert___Circle___)
Dim BP_Insert___Circle___Starbust_019: BP_Insert___Circle___Starbust_019=Array(BM_Insert___Circle___Starbust_0, LM_GI_Insert___Circle___Starbus, LM_L_Light10_Insert___Circle___)
Dim BP_Insert___Circle___Starbust_020: BP_Insert___Circle___Starbust_020=Array(BM_Insert___Circle___Starbust_0, LM_GI_Insert___Circle___Starbus, LM_L_Light9_Insert___Circle___S)
Dim BP_Insert___Circle___Starbust_021: BP_Insert___Circle___Starbust_021=Array(BM_Insert___Circle___Starbust_0, LM_GI_Insert___Circle___Starbus, LM_L_Light39_Insert___Circle___)
Dim BP_Insert___Circle___Starbust_022: BP_Insert___Circle___Starbust_022=Array(BM_Insert___Circle___Starbust_0, LM_GI_Insert___Circle___Starbus, LM_L_Light37_Insert___Circle___, LM_L_Light38_Insert___Circle___)
Dim BP_Insert___Circle___Starbust_025: BP_Insert___Circle___Starbust_025=Array(BM_Insert___Circle___Starbust_0, LM_L_Light35_Insert___Circle___)
Dim BP_Insert___Circle___Starbust_026: BP_Insert___Circle___Starbust_026=Array(BM_Insert___Circle___Starbust_0, LM_L_Light34_Insert___Circle___)
Dim BP_Insert___Circle___Starbust_027: BP_Insert___Circle___Starbust_027=Array(BM_Insert___Circle___Starbust_0, LM_L_L32_Insert___Circle___Star, LM_L_Light33_Insert___Circle___, LM_L_Light34_Insert___Circle___)
Dim BP_Insert___Circle___Starbust_1_000: BP_Insert___Circle___Starbust_1_000=Array(BM_Insert___Circle___Starbust_1, LM_L_L5_Insert___Circle___Starb, LM_L_L6_Insert___Circle___Starb)
Dim BP_Insert___Circle___Starbust_1_001: BP_Insert___Circle___Starbust_1_001=Array(BM_Insert___Circle___Starbust_1, LM_L_L5_Insert___Circle___Starb, LM_L_L6_Insert___Circle___Starb, LM_L_L7_Insert___Circle___Starb)
Dim BP_Insert___Circle___Starbust_1_002: BP_Insert___Circle___Starbust_1_002=Array(BM_Insert___Circle___Starbust_1, LM_L_L6_Insert___Circle___Starb, LM_L_L7_Insert___Circle___Starb, LM_L_L8_Insert___Circle___Starb)
Dim BP_Insert___Circle___Starbust_1_003: BP_Insert___Circle___Starbust_1_003=Array(BM_Insert___Circle___Starbust_1, LM_L_L7_Insert___Circle___Starb, LM_L_L8_Insert___Circle___Starb)
Dim BP_Parts: BP_Parts=Array(BM_Parts, LM_GI_Parts, LM_L_L16_Parts, LM_L_Light23_Parts, LM_L_L24_Parts, LM_L_Light25_Parts, LM_L_L32_Parts, LM_L_Light9_Parts)
Dim BP_Pennant_Fever_Plastic_black: BP_Pennant_Fever_Plastic_black=Array(BM_Pennant_Fever_Plastic_black)
Dim BP_Pennant_Fever_Plastic_black_001: BP_Pennant_Fever_Plastic_black_001=Array(BM_Pennant_Fever_Plastic_black_, LM_GI_Pennant_Fever_Plastic_bla)
Dim BP_Pennant_Fever_Plastic_black_002: BP_Pennant_Fever_Plastic_black_002=Array(BM_Pennant_Fever_Plastic_black_, LM_GI_Pennant_Fever_Plastic_bla)
Dim BP_PitchFlap: BP_PitchFlap=Array(BM_PitchFlap)
Dim BP_PitchFlapU: BP_PitchFlapU=Array(BM_PitchFlapU)
Dim BP_Playfield: BP_Playfield=Array(BM_Playfield, LM_GI_Playfield)
Dim BP_TargetDoubleLeft_001: BP_TargetDoubleLeft_001=Array(BM_TargetDoubleLeft_001, LM_GI_TargetDoubleLeft_001)
Dim BP_TargetDoubleRight_001: BP_TargetDoubleRight_001=Array(BM_TargetDoubleRight_001, LM_GI_TargetDoubleRight_001)
Dim BP_TargetOutLeft_001: BP_TargetOutLeft_001=Array(BM_TargetOutLeft_001, LM_GI_TargetOutLeft_001)
Dim BP_TargetOutRight_001: BP_TargetOutRight_001=Array(BM_TargetOutRight_001, LM_GI_TargetOutRight_001)
Dim BP_TargetSingleLeft_001: BP_TargetSingleLeft_001=Array(BM_TargetSingleLeft_001, LM_GI_TargetSingleLeft_001)
Dim BP_TargetSingleRight_001: BP_TargetSingleRight_001=Array(BM_TargetSingleRight_001, LM_GI_TargetSingleRight_001)
Dim BP_TargetTriple_001: BP_TargetTriple_001=Array(BM_TargetTriple_001, LM_GI_TargetTriple_001)
Dim BP_upper_deck_face: BP_upper_deck_face=Array(BM_upper_deck_face, LM_GI_upper_deck_face)
' Arrays per lighting scenario
Dim BL_GI: BL_GI=Array(LM_GI_BatFlipper, LM_GI_BatFlipperU, LM_GI_Insert___Circle___Clear_0, LM_GI_Insert___Circle___Starbus, LM_GI_Insert___Circle___Starbus, LM_GI_Insert___Circle___Starbus, LM_GI_Insert___Circle___Starbus, LM_GI_Parts, LM_GI_Pennant_Fever_Plastic_bla, LM_GI_Pennant_Fever_Plastic_bla, LM_GI_Playfield, LM_GI_TargetDoubleLeft_001, LM_GI_TargetDoubleRight_001, LM_GI_TargetOutLeft_001, LM_GI_TargetOutRight_001, LM_GI_TargetSingleLeft_001, LM_GI_TargetSingleRight_001, LM_GI_TargetTriple_001, LM_GI_upper_deck_face)
Dim BL_L_L16: BL_L_L16=Array(LM_L_L16_Insert___Circle___Clea, LM_L_L16_Insert___Circle___Star, LM_L_L16_Insert___Circle___Star, LM_L_L16_Parts)
Dim BL_L_L24: BL_L_L24=Array(LM_L_L24_Insert___Circle___Clea, LM_L_L24_Insert___Circle___Star, LM_L_L24_Parts)
Dim BL_L_L29: BL_L_L29=Array(LM_L_L29_Insert___Circle___Star, LM_L_L29_Insert___Circle___Star, LM_L_L29_Insert___Circle___Star)
Dim BL_L_L30: BL_L_L30=Array(LM_L_L30_Insert___Circle___Star, LM_L_L30_Insert___Circle___Star, LM_L_L30_Insert___Circle___Star)
Dim BL_L_L31: BL_L_L31=Array(LM_L_L31_Insert___Circle___Clea, LM_L_L31_Insert___Circle___Star, LM_L_L31_Insert___Circle___Star)
Dim BL_L_L32: BL_L_L32=Array(LM_L_L32_Insert___Circle___Clea, LM_L_L32_Insert___Circle___Star, LM_L_L32_Insert___Circle___Star, LM_L_L32_Parts)
Dim BL_L_L40: BL_L_L40=Array(LM_L_L40_BatFlipperU, LM_L_L40_Insert___Circle___Clea)
Dim BL_L_L5: BL_L_L5=Array(LM_L_L5_Insert___Circle___Starb, LM_L_L5_Insert___Circle___Starb)
Dim BL_L_L51: BL_L_L51=Array(LM_L_L51_Insert___Arrow_2_000)
Dim BL_L_L52: BL_L_L52=Array(LM_L_L52_Insert___Arrow_2_001)
Dim BL_L_L53: BL_L_L53=Array(LM_L_L53_Insert___Arrow_2_002)
Dim BL_L_L6: BL_L_L6=Array(LM_L_L6_Insert___Arrow_2_001, LM_L_L6_Insert___Circle___Starb, LM_L_L6_Insert___Circle___Starb, LM_L_L6_Insert___Circle___Starb)
Dim BL_L_L7: BL_L_L7=Array(LM_L_L7_Insert___Circle___Starb, LM_L_L7_Insert___Circle___Starb, LM_L_L7_Insert___Circle___Starb)
Dim BL_L_L8: BL_L_L8=Array(LM_L_L8_Insert___Circle___Starb, LM_L_L8_Insert___Circle___Starb)
Dim BL_L_Light10: BL_L_Light10=Array(LM_L_Light10_Insert___Circle___, LM_L_Light10_Insert___Circle___)
Dim BL_L_Light11: BL_L_Light11=Array(LM_L_Light11_Insert___Circle___)
Dim BL_L_Light12: BL_L_Light12=Array(LM_L_Light12_Insert___Circle___, LM_L_Light12_Insert___Circle___)
Dim BL_L_Light13: BL_L_Light13=Array(LM_L_Light13_Insert___Circle___, LM_L_Light13_Insert___Circle___)
Dim BL_L_Light14: BL_L_Light14=Array(LM_L_Light14_Insert___Circle___, LM_L_Light14_Insert___Circle___)
Dim BL_L_Light15: BL_L_Light15=Array(LM_L_Light15_Insert___Circle___, LM_L_Light15_Insert___Circle___, LM_L_Light15_Insert___Circle___)
Dim BL_L_Light17: BL_L_Light17=Array(LM_L_Light17_Insert___Circle___, LM_L_Light17_Insert___Circle___, LM_L_Light17_Insert___Circle___)
Dim BL_L_Light18: BL_L_Light18=Array(LM_L_Light18_Insert___Circle___, LM_L_Light18_Insert___Circle___, LM_L_Light18_Insert___Circle___)
Dim BL_L_Light19: BL_L_Light19=Array(LM_L_Light19_Insert___Circle___, LM_L_Light19_Insert___Circle___, LM_L_Light19_Insert___Circle___)
Dim BL_L_Light20: BL_L_Light20=Array(LM_L_Light20_Insert___Circle___, LM_L_Light20_Insert___Circle___, LM_L_Light20_Insert___Circle___)
Dim BL_L_Light21: BL_L_Light21=Array(LM_L_Light21_Insert___Circle___, LM_L_Light21_Insert___Circle___, LM_L_Light21_Insert___Circle___)
Dim BL_L_Light22: BL_L_Light22=Array(LM_L_Light22_Insert___Circle___, LM_L_Light22_Insert___Circle___)
Dim BL_L_Light23: BL_L_Light23=Array(LM_L_Light23_Insert___Circle___, LM_L_Light23_Insert___Circle___, LM_L_Light23_Insert___Circle___, LM_L_Light23_Parts)
Dim BL_L_Light25: BL_L_Light25=Array(LM_L_Light25_Insert___Circle___, LM_L_Light25_Insert___Circle___, LM_L_Light25_Insert___Circle___, LM_L_Light25_Parts)
Dim BL_L_Light26: BL_L_Light26=Array(LM_L_Light26_Insert___Circle___, LM_L_Light26_Insert___Circle___, LM_L_Light26_Insert___Circle___)
Dim BL_L_Light27: BL_L_Light27=Array(LM_L_Light27_Insert___Circle___, LM_L_Light27_Insert___Circle___, LM_L_Light27_Insert___Circle___)
Dim BL_L_Light28: BL_L_Light28=Array(LM_L_Light28_Insert___Circle___, LM_L_Light28_Insert___Circle___, LM_L_Light28_Insert___Circle___)
Dim BL_L_Light33: BL_L_Light33=Array(LM_L_Light33_Insert___Circle___, LM_L_Light33_Insert___Circle___)
Dim BL_L_Light34: BL_L_Light34=Array(LM_L_Light34_Insert___Circle___, LM_L_Light34_Insert___Circle___)
Dim BL_L_Light35: BL_L_Light35=Array(LM_L_Light35_Insert___Circle___)
Dim BL_L_Light36: BL_L_Light36=Array(LM_L_Light36_Insert___Circle___, LM_L_Light36_Insert___Circle___)
Dim BL_L_Light37: BL_L_Light37=Array(LM_L_Light37_Insert___Circle___, LM_L_Light37_Insert___Circle___, LM_L_Light37_Insert___Circle___)
Dim BL_L_Light38: BL_L_Light38=Array(LM_L_Light38_Insert___Circle___, LM_L_Light38_Insert___Circle___)
Dim BL_L_Light39: BL_L_Light39=Array(LM_L_Light39_BatFlipper, LM_L_Light39_BatFlipperU, LM_L_Light39_Insert___Circle___, LM_L_Light39_Insert___Circle___)
Dim BL_L_Light9: BL_L_Light9=Array(LM_L_Light9_BatFlipperU, LM_L_Light9_Insert___Circle___S, LM_L_Light9_Parts)
Dim BL_World: BL_World=Array(BM_BatFlipper, BM_BatFlipperU, BM_Insert___Arrow_2_000, BM_Insert___Arrow_2_001, BM_Insert___Arrow_2_002, BM_Insert___Circle___Clear_001, BM_Insert___Circle___Clear_002, BM_Insert___Circle___Clear_003, BM_Insert___Circle___Clear_004, BM_Insert___Circle___Starbus_02, BM_Insert___Circle___Starbus_02, BM_Insert___Circle___Starbust_1, BM_Insert___Circle___Starbust_1, BM_Insert___Circle___Starbust_1, BM_Insert___Circle___Starbust_1, BM_Insert___Circle___Starbust_0, BM_Insert___Circle___Starbust_0, BM_Insert___Circle___Starbust_0, BM_Insert___Circle___Starbust_0, BM_Insert___Circle___Starbust_0, BM_Insert___Circle___Starbust_0, BM_Insert___Circle___Starbust_0, BM_Insert___Circle___Starbust_0, BM_Insert___Circle___Starbust_0, BM_Insert___Circle___Starbust_0, BM_Insert___Circle___Starbust_0, BM_Insert___Circle___Starbust_0, BM_Insert___Circle___Starbust_0, BM_Insert___Circle___Starbust_0, BM_Insert___Circle___Starbust_0, BM_Insert___Circle___Starbust_0, BM_Insert___Circle___Starbust_0, _
  BM_Insert___Circle___Starbust_0, BM_Insert___Circle___Starbust_0, BM_Insert___Circle___Starbust_0, BM_Insert___Circle___Starbust_0, BM_Insert___Circle___Starbust_0, BM_Insert___Circle___Starbust_0, BM_Insert___Circle___Starbust_0, BM_Insert___Circle___Starbust_0, BM_Insert___Circle___Starbust_0, BM_Parts, BM_Pennant_Fever_Plastic_black, BM_Pennant_Fever_Plastic_black_, BM_Pennant_Fever_Plastic_black_, BM_PitchFlap, BM_PitchFlapU, BM_Playfield, BM_TargetDoubleLeft_001, BM_TargetDoubleRight_001, BM_TargetOutLeft_001, BM_TargetOutRight_001, BM_TargetSingleLeft_001, BM_TargetSingleRight_001, BM_TargetTriple_001, BM_upper_deck_face)
' Global arrays
Dim BG_Bakemap: BG_Bakemap=Array(BM_BatFlipper, BM_BatFlipperU, BM_Insert___Arrow_2_000, BM_Insert___Arrow_2_001, BM_Insert___Arrow_2_002, BM_Insert___Circle___Clear_001, BM_Insert___Circle___Clear_002, BM_Insert___Circle___Clear_003, BM_Insert___Circle___Clear_004, BM_Insert___Circle___Starbus_02, BM_Insert___Circle___Starbus_02, BM_Insert___Circle___Starbust_1, BM_Insert___Circle___Starbust_1, BM_Insert___Circle___Starbust_1, BM_Insert___Circle___Starbust_1, BM_Insert___Circle___Starbust_0, BM_Insert___Circle___Starbust_0, BM_Insert___Circle___Starbust_0, BM_Insert___Circle___Starbust_0, BM_Insert___Circle___Starbust_0, BM_Insert___Circle___Starbust_0, BM_Insert___Circle___Starbust_0, BM_Insert___Circle___Starbust_0, BM_Insert___Circle___Starbust_0, BM_Insert___Circle___Starbust_0, BM_Insert___Circle___Starbust_0, BM_Insert___Circle___Starbust_0, BM_Insert___Circle___Starbust_0, BM_Insert___Circle___Starbust_0, BM_Insert___Circle___Starbust_0, BM_Insert___Circle___Starbust_0, _
  BM_Insert___Circle___Starbust_0, BM_Insert___Circle___Starbust_0, BM_Insert___Circle___Starbust_0, BM_Insert___Circle___Starbust_0, BM_Insert___Circle___Starbust_0, BM_Insert___Circle___Starbust_0, BM_Insert___Circle___Starbust_0, BM_Insert___Circle___Starbust_0, BM_Insert___Circle___Starbust_0, BM_Insert___Circle___Starbust_0, BM_Parts, BM_Pennant_Fever_Plastic_black, BM_Pennant_Fever_Plastic_black_, BM_Pennant_Fever_Plastic_black_, BM_PitchFlap, BM_PitchFlapU, BM_Playfield, BM_TargetDoubleLeft_001, BM_TargetDoubleRight_001, BM_TargetOutLeft_001, BM_TargetOutRight_001, BM_TargetSingleLeft_001, BM_TargetSingleRight_001, BM_TargetTriple_001, BM_upper_deck_face)
Dim BG_Lightmap: BG_Lightmap=Array(LM_GI_BatFlipper, LM_GI_BatFlipperU, LM_GI_Insert___Circle___Clear_0, LM_GI_Insert___Circle___Starbus, LM_GI_Insert___Circle___Starbus, LM_GI_Insert___Circle___Starbus, LM_GI_Insert___Circle___Starbus, LM_GI_Parts, LM_GI_Pennant_Fever_Plastic_bla, LM_GI_Pennant_Fever_Plastic_bla, LM_GI_Playfield, LM_GI_TargetDoubleLeft_001, LM_GI_TargetDoubleRight_001, LM_GI_TargetOutLeft_001, LM_GI_TargetOutRight_001, LM_GI_TargetSingleLeft_001, LM_GI_TargetSingleRight_001, LM_GI_TargetTriple_001, LM_GI_upper_deck_face, LM_L_L16_Insert___Circle___Clea, LM_L_L16_Insert___Circle___Star, LM_L_L16_Insert___Circle___Star, LM_L_L16_Parts, LM_L_L24_Insert___Circle___Clea, LM_L_L24_Insert___Circle___Star, LM_L_L24_Parts, LM_L_L29_Insert___Circle___Star, LM_L_L29_Insert___Circle___Star, LM_L_L29_Insert___Circle___Star, LM_L_L30_Insert___Circle___Star, LM_L_L30_Insert___Circle___Star, LM_L_L30_Insert___Circle___Star, LM_L_L31_Insert___Circle___Clea, LM_L_L31_Insert___Circle___Star, _
  LM_L_L31_Insert___Circle___Star, LM_L_L32_Insert___Circle___Clea, LM_L_L32_Insert___Circle___Star, LM_L_L32_Insert___Circle___Star, LM_L_L32_Parts, LM_L_L40_BatFlipperU, LM_L_L40_Insert___Circle___Clea, LM_L_L5_Insert___Circle___Starb, LM_L_L5_Insert___Circle___Starb, LM_L_L51_Insert___Arrow_2_000, LM_L_L52_Insert___Arrow_2_001, LM_L_L53_Insert___Arrow_2_002, LM_L_L6_Insert___Arrow_2_001, LM_L_L6_Insert___Circle___Starb, LM_L_L6_Insert___Circle___Starb, LM_L_L6_Insert___Circle___Starb, LM_L_L7_Insert___Circle___Starb, LM_L_L7_Insert___Circle___Starb, LM_L_L7_Insert___Circle___Starb, LM_L_L8_Insert___Circle___Starb, LM_L_L8_Insert___Circle___Starb, LM_L_Light10_Insert___Circle___, LM_L_Light10_Insert___Circle___, LM_L_Light11_Insert___Circle___, LM_L_Light12_Insert___Circle___, LM_L_Light12_Insert___Circle___, LM_L_Light13_Insert___Circle___, LM_L_Light13_Insert___Circle___, LM_L_Light14_Insert___Circle___, LM_L_Light14_Insert___Circle___, LM_L_Light15_Insert___Circle___, LM_L_Light15_Insert___Circle___, _
  LM_L_Light15_Insert___Circle___, LM_L_Light17_Insert___Circle___, LM_L_Light17_Insert___Circle___, LM_L_Light17_Insert___Circle___, LM_L_Light18_Insert___Circle___, LM_L_Light18_Insert___Circle___, LM_L_Light18_Insert___Circle___, LM_L_Light19_Insert___Circle___, LM_L_Light19_Insert___Circle___, LM_L_Light19_Insert___Circle___, LM_L_Light20_Insert___Circle___, LM_L_Light20_Insert___Circle___, LM_L_Light20_Insert___Circle___, LM_L_Light21_Insert___Circle___, LM_L_Light21_Insert___Circle___, LM_L_Light21_Insert___Circle___, LM_L_Light22_Insert___Circle___, LM_L_Light22_Insert___Circle___, LM_L_Light23_Insert___Circle___, LM_L_Light23_Insert___Circle___, LM_L_Light23_Insert___Circle___, LM_L_Light23_Parts, LM_L_Light25_Insert___Circle___, LM_L_Light25_Insert___Circle___, LM_L_Light25_Insert___Circle___, LM_L_Light25_Parts, LM_L_Light26_Insert___Circle___, LM_L_Light26_Insert___Circle___, LM_L_Light26_Insert___Circle___, LM_L_Light27_Insert___Circle___, LM_L_Light27_Insert___Circle___, _
  LM_L_Light27_Insert___Circle___, LM_L_Light28_Insert___Circle___, LM_L_Light28_Insert___Circle___, LM_L_Light28_Insert___Circle___, LM_L_Light33_Insert___Circle___, LM_L_Light33_Insert___Circle___, LM_L_Light34_Insert___Circle___, LM_L_Light34_Insert___Circle___, LM_L_Light35_Insert___Circle___, LM_L_Light36_Insert___Circle___, LM_L_Light36_Insert___Circle___, LM_L_Light37_Insert___Circle___, LM_L_Light37_Insert___Circle___, LM_L_Light37_Insert___Circle___, LM_L_Light38_Insert___Circle___, LM_L_Light38_Insert___Circle___, LM_L_Light39_BatFlipper, LM_L_Light39_BatFlipperU, LM_L_Light39_Insert___Circle___, LM_L_Light39_Insert___Circle___, LM_L_Light9_BatFlipperU, LM_L_Light9_Insert___Circle___S, LM_L_Light9_Parts)
Dim BG_All: BG_All=Array(BM_BatFlipper, BM_BatFlipperU, BM_Insert___Arrow_2_000, BM_Insert___Arrow_2_001, BM_Insert___Arrow_2_002, BM_Insert___Circle___Clear_001, BM_Insert___Circle___Clear_002, BM_Insert___Circle___Clear_003, BM_Insert___Circle___Clear_004, BM_Insert___Circle___Starbus_02, BM_Insert___Circle___Starbus_02, BM_Insert___Circle___Starbust_1, BM_Insert___Circle___Starbust_1, BM_Insert___Circle___Starbust_1, BM_Insert___Circle___Starbust_1, BM_Insert___Circle___Starbust_0, BM_Insert___Circle___Starbust_0, BM_Insert___Circle___Starbust_0, BM_Insert___Circle___Starbust_0, BM_Insert___Circle___Starbust_0, BM_Insert___Circle___Starbust_0, BM_Insert___Circle___Starbust_0, BM_Insert___Circle___Starbust_0, BM_Insert___Circle___Starbust_0, BM_Insert___Circle___Starbust_0, BM_Insert___Circle___Starbust_0, BM_Insert___Circle___Starbust_0, BM_Insert___Circle___Starbust_0, BM_Insert___Circle___Starbust_0, BM_Insert___Circle___Starbust_0, BM_Insert___Circle___Starbust_0, BM_Insert___Circle___Starbust_0, _
  BM_Insert___Circle___Starbust_0, BM_Insert___Circle___Starbust_0, BM_Insert___Circle___Starbust_0, BM_Insert___Circle___Starbust_0, BM_Insert___Circle___Starbust_0, BM_Insert___Circle___Starbust_0, BM_Insert___Circle___Starbust_0, BM_Insert___Circle___Starbust_0, BM_Insert___Circle___Starbust_0, BM_Parts, BM_Pennant_Fever_Plastic_black, BM_Pennant_Fever_Plastic_black_, BM_Pennant_Fever_Plastic_black_, BM_PitchFlap, BM_PitchFlapU, BM_Playfield, BM_TargetDoubleLeft_001, BM_TargetDoubleRight_001, BM_TargetOutLeft_001, BM_TargetOutRight_001, BM_TargetSingleLeft_001, BM_TargetSingleRight_001, BM_TargetTriple_001, BM_upper_deck_face, LM_GI_BatFlipper, LM_GI_BatFlipperU, LM_GI_Insert___Circle___Clear_0, LM_GI_Insert___Circle___Starbus, LM_GI_Insert___Circle___Starbus, LM_GI_Insert___Circle___Starbus, LM_GI_Insert___Circle___Starbus, LM_GI_Parts, LM_GI_Pennant_Fever_Plastic_bla, LM_GI_Pennant_Fever_Plastic_bla, LM_GI_Playfield, LM_GI_TargetDoubleLeft_001, LM_GI_TargetDoubleRight_001, LM_GI_TargetOutLeft_001, _
  LM_GI_TargetOutRight_001, LM_GI_TargetSingleLeft_001, LM_GI_TargetSingleRight_001, LM_GI_TargetTriple_001, LM_GI_upper_deck_face, LM_L_L16_Insert___Circle___Clea, LM_L_L16_Insert___Circle___Star, LM_L_L16_Insert___Circle___Star, LM_L_L16_Parts, LM_L_L24_Insert___Circle___Clea, LM_L_L24_Insert___Circle___Star, LM_L_L24_Parts, LM_L_L29_Insert___Circle___Star, LM_L_L29_Insert___Circle___Star, LM_L_L29_Insert___Circle___Star, LM_L_L30_Insert___Circle___Star, LM_L_L30_Insert___Circle___Star, LM_L_L30_Insert___Circle___Star, LM_L_L31_Insert___Circle___Clea, LM_L_L31_Insert___Circle___Star, LM_L_L31_Insert___Circle___Star, LM_L_L32_Insert___Circle___Clea, LM_L_L32_Insert___Circle___Star, LM_L_L32_Insert___Circle___Star, LM_L_L32_Parts, LM_L_L40_BatFlipperU, LM_L_L40_Insert___Circle___Clea, LM_L_L5_Insert___Circle___Starb, LM_L_L5_Insert___Circle___Starb, LM_L_L51_Insert___Arrow_2_000, LM_L_L52_Insert___Arrow_2_001, LM_L_L53_Insert___Arrow_2_002, LM_L_L6_Insert___Arrow_2_001, LM_L_L6_Insert___Circle___Starb, _
  LM_L_L6_Insert___Circle___Starb, LM_L_L6_Insert___Circle___Starb, LM_L_L7_Insert___Circle___Starb, LM_L_L7_Insert___Circle___Starb, LM_L_L7_Insert___Circle___Starb, LM_L_L8_Insert___Circle___Starb, LM_L_L8_Insert___Circle___Starb, LM_L_Light10_Insert___Circle___, LM_L_Light10_Insert___Circle___, LM_L_Light11_Insert___Circle___, LM_L_Light12_Insert___Circle___, LM_L_Light12_Insert___Circle___, LM_L_Light13_Insert___Circle___, LM_L_Light13_Insert___Circle___, LM_L_Light14_Insert___Circle___, LM_L_Light14_Insert___Circle___, LM_L_Light15_Insert___Circle___, LM_L_Light15_Insert___Circle___, LM_L_Light15_Insert___Circle___, LM_L_Light17_Insert___Circle___, LM_L_Light17_Insert___Circle___, LM_L_Light17_Insert___Circle___, LM_L_Light18_Insert___Circle___, LM_L_Light18_Insert___Circle___, LM_L_Light18_Insert___Circle___, LM_L_Light19_Insert___Circle___, LM_L_Light19_Insert___Circle___, LM_L_Light19_Insert___Circle___, LM_L_Light20_Insert___Circle___, LM_L_Light20_Insert___Circle___, LM_L_Light20_Insert___Circle___, _
  LM_L_Light21_Insert___Circle___, LM_L_Light21_Insert___Circle___, LM_L_Light21_Insert___Circle___, LM_L_Light22_Insert___Circle___, LM_L_Light22_Insert___Circle___, LM_L_Light23_Insert___Circle___, LM_L_Light23_Insert___Circle___, LM_L_Light23_Insert___Circle___, LM_L_Light23_Parts, LM_L_Light25_Insert___Circle___, LM_L_Light25_Insert___Circle___, LM_L_Light25_Insert___Circle___, LM_L_Light25_Parts, LM_L_Light26_Insert___Circle___, LM_L_Light26_Insert___Circle___, LM_L_Light26_Insert___Circle___, LM_L_Light27_Insert___Circle___, LM_L_Light27_Insert___Circle___, LM_L_Light27_Insert___Circle___, LM_L_Light28_Insert___Circle___, LM_L_Light28_Insert___Circle___, LM_L_Light28_Insert___Circle___, LM_L_Light33_Insert___Circle___, LM_L_Light33_Insert___Circle___, LM_L_Light34_Insert___Circle___, LM_L_Light34_Insert___Circle___, LM_L_Light35_Insert___Circle___, LM_L_Light36_Insert___Circle___, LM_L_Light36_Insert___Circle___, LM_L_Light37_Insert___Circle___, LM_L_Light37_Insert___Circle___, _
  LM_L_Light37_Insert___Circle___, LM_L_Light38_Insert___Circle___, LM_L_Light38_Insert___Circle___, LM_L_Light39_BatFlipper, LM_L_Light39_BatFlipperU, LM_L_Light39_Insert___Circle___, LM_L_Light39_Insert___Circle___, LM_L_Light9_BatFlipperU, LM_L_Light9_Insert___Circle___S, LM_L_Light9_Parts)
' VLM  Arrays - End

' Thalamus : Exit in a clean and proper way
Sub Table1_exit
  Controller.Pause = False
  Controller.Stop
End Sub


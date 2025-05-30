'   .--..--..--..--..--..--..--..--..--..--..--..--..--..--..--..--..--..--..--..--..--..--..--.
'  / .. \.. \.. \.. \.. \.. \.. \.. \.. \.. \.. \.. \.. \.. \.. \.. \.. \.. \.. \.. \.. \.. \.. \
'  \ \/\ `'\ `'\ `'\ `'\ `'\ `'\ `'\ `'\ `'\ `'\ `'\ `'\ `'\ `'\ `'\ `'\ `'\ `'\ `'\ `'\ `'\ \/ /
'   \/ /`--'`--'`--'`--'`--'`--'`--'`--'`--'`--'`--'`--'`--'`--'`--'`--'`--'`--'`--'`--'`--'\/ /
'   / /\    _______           __          __              __   __                           / /\
'  / /\ \  |   _   | .--.--. |__| .----. |  |--. .-----. |__| |  | .--.--. .-----. .----.  / /\ \
'  \ \/ /  |.  |   | |  |  | |  | |  __| |    <  |__ --| |  | |  | |  |  | |  -__| |   _|  \ \/ /
'   \/ /   |.  |   | |_____| |__| |____| |__|__| |_____| |__| |__|  \___/  |_____| |__|     \/ /
'   / /\   |:  |   |                                                                        / /\
'  / /\ \  |::..   |                                                                       / /\ \
'  \ \/ /  `----|:.|                                                                       \ \/ /
'   \/ /        `--'                                                                        \/ /
'   / /\.--..--..--..--..--..--..--..--..--..--..--..--..--..--..--..--..--..--..--..--..--./ /\
'  / /\ \.. \.. \.. \.. \.. \.. \.. \.. \.. \.. \.. \.. \.. \.. \.. \.. \.. \.. \.. \.. \.. \/\ \
'  \ `'\ `'\ `'\ `'\ `'\ `'\ `'\ `'\ `'\ `'\ `'\ `'\ `'\ `'\ `'\ `'\ `'\ `'\ `'\ `'\ `'\ `'\ `' /
'   `--'`--'`--'`--'`--'`--'`--'`--'`--'`--'`--'`--'`--'`--'`--'`--'`--'`--'`--'`--'`--'`--'`--'
'
'
' Quicksilver
' Stern 1980
' https://www.ipdb.org/machine.cgi?id=1895

' VPW TEAM
' Blender VLM Table Build, VR Cab & Coding by MetaTed
' Playfield image provided by BorgDog
' Plastic scans provided by bord
' Blender VLM Baking by MechaEnron
' Backglass by Hauntfreaks
' VLM & Coding Tutelage by apophis
' Blender Assistance by FrankEnstein
' Modded ROM provided by Carny_Priest
' Code assistance by somatik
' Font assist by ZandysArcade
' Physical Saucer by Sixtoe


'=== TABLE OF CONTENTS  ===
'
' You can quickly jump to a section by searching the four letter tag (ZXXX)
'
' ZOPT: User Options
' ZCON: Constants and Global Variables
' ZTIM: Timers
' ZINI: Table Initialization and Exiting
'   ZMAT: General Math Functions
' ZANI: MISC ANIMATIONS
' ZDRN: Drain, Trough, and Ball Release
' ZKEY: Key Press Handling
' ZFLP: Flippers
' ZSLG: Slingshots
' ZSOL: Other Solenoids
' ZSHA: AMBIENT BALL SHADOWS
' ZPHY: GNEREAL ADVICE ON PHYSICS
' ZNFF: FLIPPER CORRECTIONS
'   ZDMP: RUBBER DAMPENERS
'   ZBOU: VPW TargetBouncer for targets and posts
' ZSSC: SLINGSHOT CORRECTION
'   ZBRL: BALL ROLLING AND DROP SOUNDS
'   ZFLE: FLEEP MECHANICAL SOUNDS
'   ZVRR: VR Room / VR Cabinet
'
'*********************************************************************************************************************************



Option Explicit
Randomize

'*******************************************
'  ZCON: Constants and Global Variables
'*******************************************

Const BallSize = 50        'Ball diameter in VPX units; must be 50
Const BallMass = 1          'Ball mass must be 1
Const tnob = 1            'Total number of balls the table can hold
Const lob = 0            'Locked balls

'See VR in Desktop mode
Const VRinDT = False

Dim tablewidth
tablewidth = Table1.width
Dim tableheight
tableheight = Table1.height


'*******************************************


On Error Resume Next
ExecuteGlobal GetTextFile("controller.vbs")
If Err Then MsgBox "You need the controller.vbs in order to run this table, available in the vp10 package"
On Error Goto 0

Const cGameName="quicksic"

LoadVPM "01560000","Stern.VBS",3.26

Const UseSolenoids=2
Const UseLamps=1
Const UseGI=0
Const UseSync=0


'*******************************************
'  User Options
'*******************************************


'----- General Sound Options -----
'Const VolumeDial = 0.8        ' Recommended values should be no greater than 1.
'Const BallRollVolume = 0.5

Dim VolumeDial : VolumeDial = 1     'Overall Mechanical sound effect volume. Recommended values should be no greater than 1.
Dim BallRollVolume : BallRollVolume = 0.5   'Level of ball rolling volume. Value between 0 and 1
Dim RampRollVolume : RampRollVolume = 0.5   'Level of ramp rolling volume. Value between 0 and 1

'----- Phsyics Mods ----- <<< ADD TARGET BOUNCER FROM VPW TABLE

Const TargetBouncerEnabled = 1    '0 = normal standup targets, 1 = bouncy targets
Const TargetBouncerFactor = 0.7   'Level of bounces. 0.0 thru 1.0, higher value is more bounciness (don't go above 1)


Dim DesktopMode: DesktopMode = Table1.ShowDT


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
Dim BallBrightness
Dim VRRoomChoice
Dim LightLevel

Sub Table1_OptionEvent(ByVal eventId)
    If eventId = 1 And Not dspTriggered Then dspTriggered = True : DisableStaticPreRendering = True : End If

    ' Sound volumes
    VolumeDial = Table1.Option("Mech Volume", 0, 1, 0.01, 0.8, 1)
    BallRollVolume = Table1.Option("Ball Roll Volume", 0, 1, 0.01, 0.5, 1)
  'RampRollVolume = Table1.Option("Ramp Roll Volume", 0, 1, 0.01, 0.5, 1)

'VR Room

  VRRoomChoice = 1
  VRRoomChoice = Table1.Option("VR Room", 1, 2, 1, 1, 0, Array("Apocalypse Lounge", "Black Void"))

  SetupRoom

  BallBrightness =  Table1.Option("Ball Brightness", 0, 1, 0.05, 0.6, 1) 'Ball brightness - Value between 0 and 1 (0=Dark ... 1=Bright)
  UpdateBallBrightness

  ' Room brightness
' LightLevel = Table1.Option("Table Brightness (Ambient Light Level)", 0, 1, 0.01, .5, 1)
  LightLevel = NightDay/100
  SetRoomBrightness LightLevel   'Uncomment this line for lightmapped tables.

    If eventId = 3 And dspTriggered Then dspTriggered = False : DisableStaticPreRendering = False : End If
End Sub

Const PLOffset = 0.5
Dim PLGain: PLGain = (1-PLOffset)/(1260-2000)

Sub UpdateBallBrightness
  Dim s, b_base, b_r, b_g, b_b, d_w
  b_base = 200 * BallBrightness
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
  RollingUpdate   'update rolling sounds
  DoDTAnim    'handle drop up target animations
  UpdateDropTargets
  DoSTAnim
  UpdateStandupTargets
End Sub

'The CorTimer interval should be 10. It's sole purpose is to update the Cor (physics) calculations
CorTimer.Interval = 10
Sub CorTimer_Timer(): Cor.Update: End Sub


'*************************************************************
'Solenoid Call backs
'**********************************************************************************************************

'SolCallback(6)  = "SolKnocker"
SolCallback(7)  = "GreenDropsUp"  ' "DropTargetBank.SolDropUp"          'Sol4 Center Drop Target Bank
SolCallback(8)  = "YellowDropsUp"  ' "RightDropTargetBank.SolDropUp"     'Sol8 Right Drop Target Bank
SolCallback(9)  = "bsKicker"           'Sol13 Kicker Hole
SolCallback(10) = "bsTrough.SolOut"         'Sol14 Outhole, BallRelease
'SolCallback(19) = "FlipperRelay"   'Sol19 sEnable

SolCallback(sLRFlipper) = "SolRFlipper"
SolCallback(sLLFlipper) = "SolLFlipper"


'*******************************************
' ZSOL: Other Solenoids
'*******************************************

' Knocker (this sub mimics how you would handle kicker in ROM based tables)
' For this to work, you must create a primitive on the table named KnockerPosition
' SolCallback(XX) = "SolKnocker"  'In ROM based tables, change the solenoid number XX to the correct number for your table.
Sub SolKnocker(Enabled) 'Knocker solenoid
  If enabled Then
    KnockerSolenoid 'Add knocker position object
  End If
End Sub


' ******************************************************************************************
'      LAMP CALLBACK for the 6 backglass flasher lamps (not the solenoid conrolled ones)
' ******************************************************************************************

Set LampCallback = GetRef("UpdateMultipleLamps")

Sub UpdateMultipleLamps()
    If Controller.Lamp(45) = 0 Then: go.state = 0: else: go.state = 1 'Game Over
    If Controller.Lamp(13) = 0 Then: hstd.state = 0: else: hstd.state = 1 'High Score To Date
    If Controller.Lamp(63) = 0 Then: ma.state = 0: else: ma.state = 1 'Match
    If Controller.Lamp(61) = 0 Then: tilt.state = 0: else: tilt.state = 1 'Tilt
    If Controller.Lamp(11) = 0 Then: sa.state = 0: else: sa.state = 1 'Shoot Again
End Sub



'******************************************************
'  ZINI: Table Initialization and Exiting
'******************************************************

Dim gBOT

Dim bsTrough, DropTargetBank, RightDropTargetBank

Sub Table1_Init
  vpmInit Me
  On Error Resume Next
    With Controller
    .GameName = cGameName
    If Err Then MsgBox "Can't start Game" & cGameName & vbNewLine & Err.Description : Exit Sub
    .SplashInfoLine = "Quicksilver (Stern) VPW"
    .HandleMechanics=0
    .HandleKeyboard=0
    .ShowDMDOnly=1
    .ShowFrame=0
    .ShowTitle=0
        .hidden = 1
         On Error Resume Next
         .Run GetPlayerHWnd
         If Err Then MsgBox Err.Description
         On Error Goto 0
     End With
     On Error Goto 0

    PinMAMETimer.Interval=PinMAMEInterval
    PinMAMETimer.Enabled = 1

    vpmNudge.TiltSwitch=7
    vpmNudge.Sensitivity=5
    vpmNudge.TiltObj = Array(Bumper1, Bumper2, Bumper3, LeftslingShot, RightslingShot)

    vpmMapLights InsertLamps ' Map all lamps to the corresponding ROM output using the value of TimerInterval of each light object

    Set bsTrough=New cvpmBallStack
        bsTrough.InitSw 0,33,0,0,0,0,0,0
        bsTrough.InitKick BallRelease,90,5
'        bsTrough.InitExitSnd SoundFX("ballrelease",DOFContactors), SoundFX("Solenoid",DOFContactors)
        bsTrough.Balls=1


'*****GI Lights On
  dim xx
  For each xx in GI:xx.State = 1: Next

  InitSlings 'hides the sling moveables

End Sub



'***************************************************************
' ZVRR: VR Room
'***************************************************************

Sub SetupRoom

  Dim VRThing, BP, DT

  If RenderingMode = 2 OR VRinDT = True Then   ' VR mode
    If VRRoomChoice = 1 Then    ' VR Room
      Lounge.visible = 1
      For Each VRThing in VRCab: VRThing.visible = 1: Next
      For Each DT in DTleds: DT.visible = 0: Next
    Else          ' Black Void
      Lounge.visible = 0
      For Each VRThing in VRCab: VRThing.visible = 1: Next
      For Each DT in DTleds: DT.visible = 0: Next
    End If
  Else
    If Table1.ShowDt  = True Then     ' Desktop / No VR / Show Desktop LEDs
      Lounge.visible = 0
      For Each VRThing in VRCab: VRThing.visible = 0: Next
      For Each DT in DTleds: DT.visible = 1: Next
    Else
      Lounge.visible = 0        ' Cab / No VR / Hide Desktop LEDs
      For Each VRThing in VRCab: VRThing.visible = 0: Next
      For Each DT in DTleds: DT.visible = 0: Next
    End If
  End If

End Sub

' VR Plunger code

Sub TimerVRPlunger_Timer
  If VR_Rod.TransY < 65 then
    VR_Rod.TransY = VR_Rod.TransY + 5
  End If
End Sub

Sub TimerVRPlunger1_Timer
  VR_Rod.TransY = (5* Plunger.Position) -20
End Sub




'******************************************************
'         DRAIN & RELEASE
'******************************************************

Sub Drain_Hit()
  'debug.print "drain_Hit"
  RandomSoundDrain Drain
  Controller.Switch(33) = 1
  bsTrough.addball me
End Sub

Sub Drain_UnHit()  'Drain
  'debug.print "drain_UnHit"
  Controller.Switch(33) = 0
End Sub

Sub BallRelease_UnHit
  'debug.print "BallRelease_UnHit"
  RandomSoundBallRelease BallRelease
End Sub



'**********************************************************************************************************
'Plunger code
'**********************************************************************************************************

Sub Table1_KeyDown(ByVal keycode)

  If keycode = LeftFlipperKey Then
  FlipperActivate LeftFlipper, LFPress
      VR_FL.TransX = 8
  End if
  If keycode = RightFlipperKey Then
  FlipperActivate RightFlipper, RFPress
      VR_FR.TransX = -8
  End if

    If KeyCode = PlungerKey Then
    Plunger.Pullback:SoundPlungerPull()
    TimerVRPlunger.Enabled = True
    TimerVRPlunger1.Enabled = False
  End If

    If keycode = LeftTiltKey Then Nudge 90, 2:SoundNudgeLeft()
    If keycode = RightTiltKey Then Nudge 270, 2:SoundNudgeRight()
    If keycode = CenterTiltKey Then Nudge 0, 1:SoundNudgeCenter()

    If keycode = keyInsertCoin1 or keycode = keyInsertCoin2 or keycode = keyInsertCoin3 or keycode = keyInsertCoin4 Then
    'debug.print "Coin_In_1,2 or 3"
    Select Case Int(rnd*3)
                        Case 0: PlaySound ("Coin_In_1"), 0, CoinSoundLevel, 0, 0.25
                        Case 1: PlaySound ("Coin_In_2"), 0, CoinSoundLevel, 0, 0.25
                        Case 2: PlaySound ("Coin_In_3"), 0, CoinSoundLevel, 0, 0.25
    End Select
    End If

    if keycode=StartGameKey then SoundStartButton()

  If vpmKeyDown(KeyCode) Then Exit Sub

End Sub

Sub Table1_KeyUp(ByVal keycode)

  If keycode = LeftFlipperKey Then
  FlipperDeActivate LeftFlipper, LFPress
  VR_FL.TransX = 0
  End if
  If keycode = RightFlipperKey Then
  FlipperDeActivate RightFlipper, RFPress
  VR_FR.TransX = 0
  End if

        If KeyCode = PlungerKey Then
    Plunger.Fire
    SoundPlungerReleaseBall()                        'Plunger release sound when there is a ball in shooter lane
'                Else
'                        SoundPlungerReleaseNoBall()                        'Plunger release sound when there is no ball in shooter lane
'                End If
    TimerVRPlunger.Enabled = False
    TimerVRPlunger1.Enabled = True
        End If

    If vpmKeyUp(KeyCode) Then Exit Sub

End Sub


'*******************************************
' ZFLP: Flippers
'*******************************************

Const ReflipAngle = 20

' Flipper Solenoid Callbacks (these subs mimics how you would handle flippers in ROM based tables)
Sub SolLFlipper(Enabled) 'Left flipper solenoid callback
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

Sub SolRFlipper(Enabled) 'Right flipper solenoid callback
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

Sub sw17_Animate
  Dim z : z = sw17.CurrentAnimOffset
  Dim BP : For Each BP in BP_sw17 : BP.transz = z: Next
End Sub

Sub sw18_Animate
  Dim z : z = sw18.CurrentAnimOffset
  Dim BP : For Each BP in BP_sw18 : BP.transz = z: Next
End Sub

Sub sw19_Animate
  Dim z : z = sw19.CurrentAnimOffset
  Dim BP : For Each BP in BP_sw19 : BP.transz = z: Next
End Sub

Sub sw20_Animate
  Dim z : z = sw20.CurrentAnimOffset
  Dim BP : For Each BP in BP_sw20 : BP.transz = z: Next
End Sub

Sub sw34_Animate
  Dim z : z = sw34.CurrentAnimOffset
  Dim BP : For Each BP in BP_sw34 : BP.transz = z: Next
End Sub

Sub sw35_Animate
  Dim z : z = sw35.CurrentAnimOffset
  Dim BP : For Each BP in BP_sw35 : BP.transz = z: Next
End Sub

Sub sw36_Animate
  Dim z : z = sw36.CurrentAnimOffset
  Dim BP : For Each BP in BP_sw36 : BP.transz = z: Next
End Sub

Sub sw37_Animate
  Dim z : z = sw37.CurrentAnimOffset
  Dim BP : For Each BP in BP_sw37 : BP.transz = z: Next
End Sub

Sub sw28_Animate
  Dim z : z = sw28.CurrentAnimOffset
  Dim BP : For Each BP in BP_sw28 : BP.transz = z: Next
End Sub

Sub sw39_Animate
  Dim z : z = sw39.CurrentAnimOffset
  Dim BP : For Each BP in BP_sw39 : BP.transz = z: Next
End Sub


''' Spinner Animations

' Set all Spin(x)Rod Zpos to 60 after new bake

Sub sw4_Animate
  Dim spinangle:spinangle = sw4.currentangle
  Dim BL : For Each BL in BP_SpinR : BL.RotX = spinangle: Next
  Dim BR
  For Each BR in BP_SpinR_Rod
    BR.TransX = sin( (sw4.CurrentAngle) * (2*PI/360)) * 3.5
    BR.TransZ = sin( (sw4.CurrentAngle- 90) * (2*PI/360)) * 10
  Next
End Sub

Sub sw5_Animate
  Dim spinangle:spinangle = sw5.currentangle
  Dim BL : For Each BL in BP_SpinL : BL.RotX = spinangle: Next
  Dim BR
  For Each BR in BP_SpinL_Rod
    BR.TransX = sin((sw5.CurrentAngle) * (2*PI/360)) * 3.5
    BR.TransZ = sin((sw5.CurrentAngle-90) * (2*PI/360)) * 10
  Next
End Sub


''' Bumper Animations
Sub Bumper1_Animate
  Dim z, BL
  z = Bumper1.CurrentRingOffset
  For Each BL in BP_Bumper1_Ring : BL.transz = z: Next
End Sub

Sub Bumper2_Animate
  Dim z, BL
  z = Bumper2.CurrentRingOffset
  For Each BL in BP_Bumper2_Ring  : BL.transz = z: Next
End Sub

Sub Bumper3_Animate
  Dim z, BL
  z = Bumper3.CurrentRingOffset
  For Each BL in BP_Bumper3_Ring  : BL.transz = z: Next
End Sub


''' Gate Animations

Sub Gate1_Animate
  Dim a : a = Gate1.CurrentAngle
  Dim BL : For Each BL in BP_Gate01_Flap : BL.rotx = a: Next
End Sub

Sub Gate2_Animate
  Dim a : a = Gate2.CurrentAngle
  Dim BL : For Each BL in BP_Gate02_Flap : BL.rotx = a: Next
End Sub

Sub TrigGate1SFX_hit
  if activeball.velx < 0 Then RandomSoundMetal
End Sub

Sub TrigGate2SFX_hit
  if activeball.velx > 0 Then RandomSoundMetal
End Sub




'******************************************************
' ZRST: STAND-UP TARGETS by Rothbauerw
'******************************************************


'Stand Up Targets
Sub sw14_Hit : STHit 14: End Sub
Sub sw15_Hit : STHit 15: End Sub
Sub sw16_Hit : STHit 16: End Sub
Sub sw25_Hit : STHit 25: End Sub
Sub sw26_Hit : STHit 26: End Sub
Sub sw27_Hit : STHit 27: End Sub
Sub sw40_Hit : STHit 40: End Sub



Class StandupTarget
  Private m_primary, m_prim, m_sw, m_animate

  Public Property Get Primary(): Set Primary = m_primary: End Property
  Public Property Let Primary(val): Set m_primary = val: End Property

  Public Property Get Prim(): Set Prim = m_prim: End Property
  Public Property Let Prim(val): Set m_prim = val: End Property

  Public Property Get Sw(): Sw = m_sw: End Property
  Public Property Let Sw(val): m_sw = val: End Property

  Public Property Get Animate(): Animate = m_animate: End Property
  Public Property Let Animate(val): m_animate = val: End Property

  Public default Function init(primary, prim, sw, animate)
    Set m_primary = primary
    Set m_prim = prim
    m_sw = sw
    m_animate = animate

    Set Init = Me
  End Function
End Class


'Define a variable for each stand-up target
Dim ST14, ST15, ST16, ST25, ST26, ST27, ST40

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

Set ST14 = (new StandupTarget)(sw14, BM_sw14, 14, 0)
Set ST15 = (new StandupTarget)(sw15, BM_sw15, 15, 0)
Set ST16 = (new StandupTarget)(sw16, BM_sw16, 16, 0)
Set ST25 = (new StandupTarget)(sw25, BM_sw25, 25, 0)
Set ST26 = (new StandupTarget)(sw26, BM_sw26, 26, 0)
Set ST27 = (new StandupTarget)(sw27, BM_sw27, 27, 0)
Set ST40 = (new StandupTarget)(sw40, BM_sw40, 40, 0)


'Add all the Stand-up Target Arrays to Stand-up Target Animation Array
'   STAnimationArray = Array(ST1, ST2, ....)
Dim STArray
STArray = Array(ST14, ST15, ST16, ST25, ST26, ST27, ST40)

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
  STArray(i).animate = STCheckHit(Activeball,STArray(i).primary)

  If STArray(i).animate <> 0 Then
    DTBallPhysics Activeball, STArray(i).primary.orientation, STMass
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
    primary.uservalue = gametime
  End If

  animtime = gametime - primary.uservalue

  If animate = 1 Then
    primary.collidable = 0
    prim.transy =  - STMaxOffset
    vpmTimer.PulseSw switch
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
  dim BP, y

    y = BM_sw14.transy
  For Each BP in BP_sw14 : BP.transy = y: Next

    y = BM_sw15.transy
  For Each BP in BP_sw15 : BP.transy = y: Next

    y = BM_sw16.transy
  For Each BP in BP_sw16 : BP.transy = y: Next

    y = BM_sw25.transy
  For Each BP in BP_sw25 : BP.transy = y: Next

    y = BM_sw26.transy
  For Each BP in BP_sw26 : BP.transy = y: Next

    y = BM_sw27.transy
  For Each BP in BP_sw27 : BP.transy = y: Next

    y = BM_sw40.transy
  For Each BP in BP_sw40 : BP.transy = y: Next

End Sub




'*******************************************
'  ZDTA: Drop Targets
'*******************************************

'Solenoids

Sub GreenDropsUp(enabled)
  If enabled then
    DTRaise 21
    DTRaise 22
    DTRaise 23
    DTRaise 24
    RandomSoundDropTargetReset BM_sw23
  End If
End Sub

Sub YellowDropsUp(enabled)
  If enabled then
    DTRaise 30
    DTRaise 31
    DTRaise 32
    RandomSoundDropTargetReset DTBM_sw31
  End If
End Sub


'Green

Sub sw21_Hit: DTHit 21 : End Sub
Sub sw22_Hit: DTHit 22: End Sub
Sub sw23_Hit: DTHit 23: End Sub
Sub sw24_Hit: DTHit 24: End Sub

'Yellow

Sub sw30_Hit: DTHit 30: End Sub
Sub sw31_Hit: DTHit 31: End Sub
Sub sw32_Hit: DTHit 32: End Sub




Sub UpdateDropTargets
  dim BP, tz, rx, ry

  tz = BM_sw21.transz
  rx = BM_sw21.rotx
  ry = BM_sw21.roty
  For each BP in BP_sw21: BP.transz = tz: BP.rotx = rx: BP.roty = ry: Next

  tz = BM_sw22.transz
  rx = BM_sw22.rotx
  ry = BM_sw22.roty
  For each BP in BP_sw22 : BP.transz = tz: BP.rotx = rx: BP.roty = ry: Next

  tz = BM_sw23.transz
  rx = BM_sw23.rotx
  ry = BM_sw23.roty
  For each BP in BP_sw23 : BP.transz = tz: BP.rotx = rx: BP.roty = ry: Next

  tz = BM_sw24.transz
  rx = BM_sw24.rotx
  ry = BM_sw24.roty
  For each BP in BP_sw24 : BP.transz = tz: BP.rotx = rx: BP.roty = ry: Next

  tz = DTBM_sw30.transz
  rx = DTBM_sw30.rotx
  ry = DTBM_sw30.roty
  For each BP in BP_DTsw30 : BP.transz = tz: BP.rotx = rx: BP.roty = ry: Next

  tz = DTBM_sw31.transz
  rx = DTBM_sw31.rotx
  ry = DTBM_sw31.roty
  For each BP in BP_DTsw31 : BP.transz = tz: BP.rotx = rx: BP.roty = ry: Next

  tz = DTBM_sw32.transz
  rx = DTBM_sw32.rotx
  ry = DTBM_sw32.roty
  For each BP in BP_DTsw32 : BP.transz = tz: BP.rotx = rx: BP.roty = ry: Next

End Sub



''''''DROP TARGETS INITIALIZATION


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
Dim DT21, DT22, DT23, DT24      'Green
Dim DT30, DT31, DT32        'Yellow

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

Set DT21 = (new DropTarget)(sw21, sw21a, BM_sw21, 21, 0, False)
Set DT22 = (new DropTarget)(sw22, sw22a, BM_sw22, 22, 0, False)
Set DT23 = (new DropTarget)(sw23, sw23a, BM_sw23, 23, 0, False)
Set DT24 = (new DropTarget)(sw24, sw24a, BM_sw24, 24, 0, False)

Set DT30 = (new DropTarget)(sw30, sw30a, DTBM_sw30, 30, 0, False)
Set DT31 = (new DropTarget)(sw31, sw31a, DTBM_sw31, 31, 0, False)
Set DT32 = (new DropTarget)(sw32, sw32a, DTBM_sw32, 32, 0, False)


Dim DTArray
DTArray = Array(DT21, DT22, DT23, DT24, DT30, DT31, DT32)

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


'''''''DROP TARGETS FUNCTIONS


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
      'DTAction Switchid
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
          gBOT(b).velz = 10
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
'*   END DROP TARGETS
'******************************************************







'****************************
'   ZRBR: Room Brightness
'****************************

'This code only applies to lightmapped tables. It is here for reference.
'NOTE: Objects bightness will be affected by the Day/Night slider only if their blenddisablelighting property is less than 1.
'      Lightmapped table primitives have their blenddisablelighting equal to 1, therefore we need this SetRoomBrightness sub
'      to handle updating their effective ambient brighness.

' Update these arrays if you want to change more materials with room light level



'******************************************************
' ZDRN: Drain, Trough, and Ball Release
'******************************************************

'********************* TROUGH *************************

'
'Sub UpdateTrough
' UpdateTroughTimer.Interval = 300
' UpdateTroughTimer.Enabled = 1
'End Sub
'
'Sub UpdateTroughTimer_Timer
' If sw14.BallCntOver = 0 Then sw13.kick 57, 10
' If sw13.BallCntOver = 0 Then sw12.kick 57, 10
' If sw12.BallCntOver = 0 Then sw11.kick 57, 10
' If sw11.BallCntOver = 0 Then sw10.kick 57, 10
' If sw10.BallCntOver = 0 Then sw9.kick 57, 10
' Me.Enabled = 0
'End Sub


'*****************  DRAIN & RELEASE  ******************

'Sub SolTrough(enabled)
' If enabled Then
'   Drain.kick 57, 20
'   Controller.Switch(33) = 1
' End If
'End Sub
'
'Sub SolRelease(enabled)
' If enabled Then
'   Release.kick 57, 10
'   Controller.Switch(33) = 0
'   RandomSoundBallRelease Release
' End If
'End Sub





'************************************************************
' ZSLG: Slingshot Animations
'************************************************************

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
  LS.VelocityCorrect(Activeball)
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
  RS.VelocityCorrect(Activeball)
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



'************************* KICKERS *****************************

Dim KickerBall29

Sub KickBall(kball, kangle, kvel, kvelz, kzlift)
  dim rangle
  rangle = PI * (kangle - 90) / 180

  kball.z = kball.z + kzlift
  kball.velz = kvelz
  kball.velx = cos(rangle)*kvel
  kball.vely = sin(rangle)*kvel
End Sub

'Kicker
Sub sw29_Hit
' msgbox activeball.z
  table1.BallPlayfieldReflectionScale = 0
  set KickerBall29 = activeball
  Controller.Switch(29) = 1
  SoundSaucerLock
  sw29.timerinterval=10
  sw29.timerenabled=true
End Sub

Sub sw29_unHit
  table1.BallPlayfieldReflectionScale = 1
  Controller.Switch(29) = 0
  sw29.timerenabled=false
end sub

sub sw29_timer
' debug.print "vuk left: " & KickerBall37.z
  'prevents ball wiggle in saucer
  KickerBall29.x=sw29.x
  KickerBall29.y=sw29.y
' KickerBall29.z=6.5 'a bit higher so it wont touch the collidable bottom
  KickerBall29.angmomx=0
  KickerBall29.angmomy=0
  KickerBall29.angmomz=0
  KickerBall29.velx=0
  KickerBall29.vely=0
  KickerBall29.velz=0
  If Controller.Switch(29) = 0 Then me.timerenabled = false
end sub

Sub bsKicker(Enable)
  sw29.timerenabled=false
    If Enable then
    If Controller.Switch(29) <> 0 Then
      KickBall KickerBall29, 150, 15, 5, 20
      SoundSaucerKick 1, sw29
'     Controller.Switch(29) = 0 'this may not be safe here
    End If
  End If
End Sub

'************************************************************
' ZSWI: SWITCHES
'************************************************************

'Bumpers
Sub Bumper1_Hit : vpmTimer.PulseSw(9) : RandomSoundBumperTop Bumper1 : End Sub
Sub Bumper2_Hit : vpmTimer.PulseSw(10) : RandomSoundBumperMiddle Bumper2: End Sub
Sub Bumper3_Hit : vpmTimer.PulseSw(11) : RandomSoundBumperBottom Bumper3: End Sub

dim damp4, gt4, spins4, spintime4, damp5, gt5, spins5, spintime5

Sub sw4hit_Hit
  damp4 = 0
  spins4 = 0
  'gt4 = 0
  spintime4 = 1
  sw4.damping = 0.9985
  'debug.print "sw4 HIT!"
  sw4hit.timerinterval = 200
  sw4hit.Timerenabled = 1
End Sub

Sub sw4hit_Timer
  if spins4 = 0 then
    sw4.damping = 0.95
    'debug.print "No Spin!"
  end if
  sw4hit.Timerenabled = 0
End Sub


Sub sw4_Spin

  vpmTimer.PulseSw 4
  SoundSpinner sw4

  spins4 = spins4+1

  if spins4 = 1 then
    gt4 = gametime
  end if
  if spins4 = 2 then
    spintime4 = gametime - gt4
  end if
  if spintime4 <> 0 then
    if spins4 > (4000 / spintime4) then
      damp4 = 0.005 + damp4
    end if
  end if
  if spintime4 > 120 then
    damp4 =0.01 + damp4
  end if

  sw4.damping = sw4.damping - damp4

  'debug.print "sw4 Spins " &spins4
  'debug.print "sw4 DampAmt = " &damp4
  'debug.print "sw4 Damping = " &sw4.damping
  'debug.print "sw4 Spintime " &spintime4
End Sub

Sub sw5hit_Hit
  damp5 = 0
  spins5 = 0
  'gt5 = 0
  spintime5 = 1
  sw5.damping = 0.9985
  'debug.print "sw5 HIT!"
  sw5hit.timerinterval = 200
  sw5hit.Timerenabled = 1
End Sub

Sub sw5hit_Timer
  if spins5 = 0 then
    sw5.damping = 0.95
    'debug.print "No Spin!"
  end if
  sw5hit.Timerenabled = 0
End Sub

Sub sw5_Spin

  vpmTimer.PulseSw 5
  SoundSpinner sw5

  spins5 = spins5+1

  if spins5 = 1 then
    gt5 = gametime
  end if
  if spins5 = 2 then
    spintime5 = gametime - gt5 + 1.001
  end if
  if spintime5 <> 0 then
    if spins5 > (4000 / spintime5) then
      damp5 = 0.005 + damp5
    end if
  end if
  if spintime5 > 120 then
    damp5 =0.01 + damp5
  end if

  sw5.damping = sw5.damping - damp5

  'debug.print "sw5 Spins " &spins5
  'debug.print "sw5 DampAmt = " &damp5
  'debug.print "sw5 Damping = " &sw5.damping
  'debug.print "sw5 Spintime " &spintime5
End Sub



'Wire Triggers
sub sw17_hit:Controller.Switch(17)=1 : End Sub
sub sw17_unhit:Controller.Switch(17)=0:End Sub
sub sw18_hit:Controller.Switch(18)=1 : End Sub
sub sw18_unhit:Controller.Switch(18)=0:End Sub
sub sw19_hit:Controller.Switch(19)=1 : End Sub
sub sw19_unhit:Controller.Switch(19)=0:End Sub
sub sw20_hit:Controller.Switch(20)=1 : End Sub
sub sw20_unhit:Controller.Switch(20)=0:End Sub
sub sw34_hit:Controller.Switch(34)=1 : End Sub
sub sw34_unhit:Controller.Switch(34)=0:End Sub
sub sw35_hit:Controller.Switch(35)=1 : End Sub
sub sw35_unhit:Controller.Switch(35)=0:End Sub
sub sw36_hit:Controller.Switch(36)=1 : End Sub
sub sw36_unhit:Controller.Switch(36)=0:End Sub
sub sw37_hit:Controller.Switch(37)=1 : End Sub
sub sw37_unhit:Controller.Switch(37)=0:End Sub

'Drop Targets
'Sub sw21_Hit:DropTargetBank.Hit 1
'TargetBouncer Activeball, 1
'End Sub
'Sub sw22_Hit:DropTargetBank.Hit 2
'TargetBouncer Activeball, 1
'End Sub
'Sub sw23_Hit:DropTargetBank.Hit 3
'TargetBouncer Activeball, 1
'End Sub
'Sub sw24_Hit:DropTargetBank.Hit 4
'TargetBouncer Activeball, 1
'End Sub
'Sub sw30_Hit:RightDropTargetBank.Hit 1
'TargetBouncer Activeball, 1
'End Sub
'Sub sw31_Hit:RightDropTargetBank.Hit 2
'TargetBouncer Activeball, 1
'End Sub
'Sub sw32_Hit:RightDropTargetBank.Hit 3
'TargetBouncer Activeball, 1
'End Sub

'Scoring Rubbers
Sub sw38a_Hit
  vpmTimer.PulseSw 38
  'debug.print "sw38a"
End Sub
Sub sw38b_Hit
  vpmTimer.PulseSw 38
  'debug.print "sw38b"
End Sub
Sub sw38c_Hit
  vpmTimer.PulseSw 38
  'debug.print "sw38c"
End Sub
Sub sw38d_Hit
  vpmTimer.PulseSw 38
  'debug.print "sw38d"
End Sub
Sub sw38e_Hit
  vpmTimer.PulseSw 38
  'debug.print "sw38e"
End Sub

'Star Triggers
sub sw28_hit:Controller.Switch(28)=1 : End Sub
sub sw28_unhit:Controller.Switch(28)=0:End Sub
sub sw39_hit:Controller.Switch(39)=1 : End Sub
sub sw39_unhit:Controller.Switch(39)=0:End Sub

' Inlane switch speedlimit code

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




Sub ResetDropsL(enabled)
  if enabled then
    PlaySoundAt SoundFX(DTResetSound,DOFContactors),sw34p
    DTRaise 33
    DTRaise 34
    DTRaise 59
  end if
End Sub


Sub ResetDropsR(enabled)
  if enabled then
    PlaySoundAt SoundFX(DTResetSound,DOFContactors),sw34p
    DTRaise 35
    DTRaise 36
    DTRaise 57
  end if
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
SaucerLockSoundLevel = 5
SaucerKickSoundLevel = 5

'///////////////////////-----Gates, Spinners, Rollovers and Targets-----///////////////////////

Dim GateSoundLevel, TargetSoundFactor, SpinnerSoundLevel, RolloverSoundLevel, DTSoundLevel

GateSoundLevel = 0.5 / 5      'volume level; range [0, 1]
TargetSoundFactor = 0.0025 * 10  'volume multiplier; must not be zero
DTSoundLevel = 0.25        'volume multiplier; must not be zero
RolloverSoundLevel = 0.25      'volume level; range [0, 1]
SpinnerSoundLevel = 1      'volume level; range [0, 1]

'///////////////////////-----Ball Release, Guides and Drain-----///////////////////////
Dim DrainSoundLevel, BallReleaseSoundLevel, BottomArchBallGuideSoundFactor, FlipperBallGuideSoundFactor

DrainSoundLevel = 5          'volume level; range [0, 1]
BallReleaseSoundLevel = 5        'volume level; range [0, 1]
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
  'tmp = tableobj.y * 2 / tableheight - 1

  If tmp > 7000 Then
    tmp = 7000
  ElseIf tmp <  - 7000 Then
    tmp =  - 7000
  End If

  If tmp > 0 Then
    AudioFade = CSng(tmp ^ 5) 'was 10
  Else
    AudioFade = CSng( - (( - tmp) ^ 5) ) 'was 10
  End If
End Function

Function AudioPan(tableobj) ' Calculates the pan for a tableobj based on the X position on the table. "table1" is the name of the table
  Dim tmp
  'tmp = tableobj.x * 2 / tablewidth - 1

  If tmp > 7000 Then
    tmp = 7000
  ElseIf tmp <  - 7000 Then
    tmp =  - 7000
  End If

  If tmp > 0 Then
    AudioPan = CSng(tmp ^ 5) ' was 10
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
  'debug.print "SoundStartButton"
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
  'debug.print "RandomSoundDrain"
  PlaySoundAtLevelStatic ("Drain_" & Int(Rnd * 11) + 1), DrainSoundLevel, drainswitch
End Sub

'/////////////////////////////  TROUGH BALL RELEASE SOLENOID SOUNDS  ////////////////////////////

Sub RandomSoundBallRelease(drainswitch)
  'debug.print "RandomSoundBallRelease"
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
  'debug.print "SoundSaucerLock"
  PlaySoundAtLevelStatic ("Saucer_Enter_" & Int(Rnd * 2) + 1), SaucerLockSoundLevel, ActiveBall
End Sub

Sub SoundSaucerKick(scenario, saucer)
  Select Case scenario
    Case 0
      PlaySoundAtLevelStatic SoundFX("Saucer_Empty", DOFContactors), SaucerKickSoundLevel, saucer
    Case 1
      'debug.print "SoundSaucerKick 1"
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

'
''*******************************************
'' Late 70's to early 80's

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

Const FlipperCoilRampupMode = 2 '0 = fast, 1 = medium, 2 = slow (tap passes should work)

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
Dim objBallShadow

'Initialization
BSInit

Sub BSInit()
  'Dim iii
  'Prepare the shadow objects before play begins
    Set objBallShadow = Eval("BallShadow")
    objBallShadow.material = "BallShadow"
    UpdateMaterial objBallShadow.material,1,0,0,0,0,0,AmbientBSFactor,RGB(0,0,0),0,0,False,True,0,0,0,0
    objBallShadow.Z = 3 / 1000
    objBallShadow.visible = 0
End Sub


Sub BSUpdate

    Dim b
    Dim gBOT
    gBOT = GetBalls

    For b = 0 To UBound(gBOT)

    'Primitive shadow on playfield, flasher shadow in ramps
    '** If on main and upper pf
    If gBOT(b).Z > 20 And gBOT(b).Z < 30 Then
      objBallShadow.visible = 1
      objBallShadow.X = gBOT(b).X + (gBOT(b).X - (tablewidth / 2)) / (Ballsize / AmbientMovement) + offsetX
      objBallShadow.Y = gBOT(b).Y + offsetY


    '** No shadow if ball is off the main playfield (this may need to be adjusted per table)
    Else
      objBallShadow.visible = 0
    End If

    Next


End Sub


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
DigitsVR(0) = Array(p1_1a,p1_1b,p1_1c,p1_1d,p1_1e,p1_1f,p1_1g)
DigitsVR(1) = Array(p1_2a,p1_2b,p1_2c,p1_2d,p1_2e,p1_2f,p1_2g)
DigitsVR(2) = Array(p1_3a,p1_3b,p1_3c,p1_3d,p1_3e,p1_3f,p1_3g)
DigitsVR(3) = Array(p1_4a,p1_4b,p1_4c,p1_4d,p1_4e,p1_4f,p1_4g)
DigitsVR(4) = Array(p1_5a,p1_5b,p1_5c,p1_5d,p1_5e,p1_5f,p1_5g)
DigitsVR(5) = Array(p1_6a,p1_6b,p1_6c,p1_6d,p1_6e,p1_6f,p1_6g)
DigitsVR(6) = Array(p1_7a,p1_7b,p1_7c,p1_7d,p1_7e,p1_7f,p1_7g)

DigitsVR(7) = Array(p2_1a,p2_1b,p2_1c,p2_1d,p2_1e,p2_1f,p2_1g)
DigitsVR(8) = Array(p2_2a,p2_2b,p2_2c,p2_2d,p2_2e,p2_2f,p2_2g)
DigitsVR(9) = Array(p2_3a,p2_3b,p2_3c,p2_3d,p2_3e,p2_3f,p2_3g)
DigitsVR(10) = Array(p2_4a,p2_4b,p2_4c,p2_4d,p2_4e,p2_4f,p2_4g)
DigitsVR(11) = Array(p2_5a,p2_5b,p2_5c,p2_5d,p2_5e,p2_5f,p2_5g)
DigitsVR(12) = Array(p2_6a,p2_6b,p2_6c,p2_6d,p2_6e,p2_6f,p2_6g)
DigitsVR(13) = Array(p2_7a,p2_7b,p2_7c,p2_7d,p2_7e,p2_7f,p2_7g)

DigitsVR(14) = Array(p3_1a,p3_1b,p3_1c,p3_1d,p3_1e,p3_1f,p3_1g)
DigitsVR(15) = Array(p3_2a,p3_2b,p3_2c,p3_2d,p3_2e,p3_2f,p3_2g)
DigitsVR(16) = Array(p3_3a,p3_3b,p3_3c,p3_3d,p3_3e,p3_3f,p3_3g)
DigitsVR(17) = Array(p3_4a,p3_4b,p3_4c,p3_4d,p3_4e,p3_4f,p3_4g)
DigitsVR(18) = Array(p3_5a,p3_5b,p3_5c,p3_5d,p3_5e,p3_5f,p3_5g)
DigitsVR(19) = Array(p3_6a,p3_6b,p3_6c,p3_6d,p3_6e,p3_6f,p3_6g)
DigitsVR(20) = Array(p3_7a,p3_7b,p3_7c,p3_7d,p3_7e,p3_7f,p3_7g)

DigitsVR(21) = Array(p4_1a,p4_1b,p4_1c,p4_1d,p4_1e,p4_1f,p4_1g)
DigitsVR(22) = Array(p4_2a,p4_2b,p4_2c,p4_2d,p4_2e,p4_2f,p4_2g)
DigitsVR(23) = Array(p4_3a,p4_3b,p4_3c,p4_3d,p4_3e,p4_3f,p4_3g)
DigitsVR(24) = Array(p4_4a,p4_4b,p4_4c,p4_4d,p4_4e,p4_4f,p4_4g)
DigitsVR(25) = Array(p4_5a,p4_5b,p4_5c,p4_5d,p4_5e,p4_5f,p4_5g)
DigitsVR(26) = Array(p4_6a,p4_6b,p4_6c,p4_6d,p4_6e,p4_6f,p4_6g)
DigitsVR(27) = Array(p4_7a,p4_7b,p4_7c,p4_7d,p4_7e,p4_7f,p4_7g)

DigitsVR(28) = Array(cr_1a,cr_1b,cr_1c,cr_1d,cr_1e,cr_1f,cr_1g)
DigitsVR(29) = Array(cr_2a,cr_2b,cr_2c,cr_2d,cr_2e,cr_2f,cr_2g)

DigitsVR(30) = Array(mb_1a,mb_1b,mb_1c,mb_1d,mb_1e,mb_1f,mb_1g)
DigitsVR(31) = Array(mb_2a,mb_2b,mb_2c,mb_2d,mb_2e,mb_2f,mb_2g)


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



'****************************************************
'*                          *
'*        VLM ARRAYS              *
'*                          *
'****************************************************

' VLM  Arrays - Start
' Arrays per baked part
Dim BP_Bumper1_Ring: BP_Bumper1_Ring=Array(BM_Bumper1_Ring, LM_GI_Bumper1_Ring)
Dim BP_Bumper2_Ring: BP_Bumper2_Ring=Array(BM_Bumper2_Ring, LM_GIS_gi012_Bumper2_Ring, LM_GI_Bumper2_Ring)
Dim BP_Bumper3_Ring: BP_Bumper3_Ring=Array(BM_Bumper3_Ring, LM_GI_Bumper3_Ring)
Dim BP_FlipperL: BP_FlipperL=Array(BM_FlipperL, LM_GIS_gi001_FlipperL, LM_GIS_gi002_FlipperL)
Dim BP_FlipperLup: BP_FlipperLup=Array(BM_FlipperLup, LM_GIS_gi001_FlipperLup, LM_GIS_gi002_FlipperLup)
Dim BP_FlipperR: BP_FlipperR=Array(BM_FlipperR, LM_GIS_gi003_FlipperR)
Dim BP_FlipperRup: BP_FlipperRup=Array(BM_FlipperRup, LM_GIS_gi003_FlipperRup)
Dim BP_Gate01_Flap: BP_Gate01_Flap=Array(BM_Gate01_Flap, LM_GIS_gi014_Gate01_Flap, LM_GIS_gi015_Gate01_Flap, LM_GIS_gi016_Gate01_Flap, LM_GI_Gate01_Flap)
Dim BP_Gate02_Flap: BP_Gate02_Flap=Array(BM_Gate02_Flap, LM_GIS_gi017_Gate02_Flap, LM_GIS_gi018_Gate02_Flap, LM_GIS_gi020_Gate02_Flap, LM_GIS_gi021_Gate02_Flap, LM_GI_Gate02_Flap)
Dim BP_Inserts: BP_Inserts=Array(BM_Inserts, LM_L_L01_Inserts, LM_L_L02_Inserts, LM_L_L03_Inserts, LM_L_L04_Inserts, LM_L_L05_Inserts, LM_L_L07_Inserts, LM_L_L08_Inserts, LM_L_L09_Inserts, LM_L_L11_Inserts, LM_L_L12_Inserts, LM_L_L14_Inserts, LM_L_L17_Inserts, LM_L_L18_Inserts, LM_L_L19_Inserts, LM_L_L20_Inserts, LM_L_L21_Inserts, LM_L_L22_Inserts, LM_L_L23_Inserts, LM_L_L24_Inserts, LM_L_L28_Inserts, LM_L_L30_Inserts, LM_L_L31_Inserts, LM_L_L33_Inserts, LM_L_L34_Inserts, LM_L_L35_Inserts, LM_L_L36_Inserts, LM_L_L37_Inserts, LM_L_L38_Inserts, LM_L_L39_Inserts, LM_L_L40_Inserts, LM_L_L44_Inserts, LM_L_L46_Inserts, LM_L_L47_Inserts, LM_L_L49_Inserts, LM_L_L50_Inserts, LM_L_L51_Inserts, LM_L_L52_Inserts, LM_L_L53_Inserts, LM_L_L54_Inserts, LM_L_L55_Inserts, LM_L_L56_Inserts, LM_L_L59_Inserts, LM_L_L60_Inserts, LM_L_L62_Inserts, LM_GIS_gi001_Inserts, LM_GIS_gi002_Inserts, LM_GIS_gi003_Inserts, LM_GIS_gi004_Inserts, LM_GIS_gi005_Inserts, LM_GIS_gi006_Inserts, LM_GIS_gi008_Inserts, LM_GIS_gi012_Inserts, _
  LM_GIS_gi013_Inserts, LM_GIS_gi018_Inserts, LM_GIS_gi019_Inserts, LM_GIS_gi020_Inserts, LM_GI_Inserts)
Dim BP_LSling1: BP_LSling1=Array(BM_LSling1, LM_GIS_gi001_LSling1, LM_GIS_gi002_LSling1, LM_GIS_gi003_LSling1, LM_GIS_gi004_LSling1, LM_GIS_gi005_LSling1)
Dim BP_LSling2: BP_LSling2=Array(BM_LSling2, LM_GIS_gi001_LSling2, LM_GIS_gi002_LSling2, LM_GIS_gi003_LSling2, LM_GIS_gi004_LSling2)
Dim BP_LSlingArm: BP_LSlingArm=Array(BM_LSlingArm, LM_GIS_gi001_LSlingArm, LM_GIS_gi002_LSlingArm, LM_GIS_gi003_LSlingArm, LM_GIS_gi004_LSlingArm, LM_GIS_gi005_LSlingArm)
Dim BP_Parts: BP_Parts=Array(BM_Parts, LM_L_L09_Parts, LM_L_L12_Parts, LM_L_L14_Parts, LM_L_L22_Parts, LM_L_L24_Parts, LM_L_L28_Parts, LM_L_L30_Parts, LM_L_L31_Parts, LM_L_L35_Parts, LM_L_L38_Parts, LM_L_L44_Parts, LM_L_L54_Parts, LM_L_L56_Parts, LM_L_L59_Parts, LM_L_L60_Parts, LM_L_L62_Parts, LM_GIS_gi001_Parts, LM_GIS_gi002_Parts, LM_GIS_gi003_Parts, LM_GIS_gi004_Parts, LM_GIS_gi005_Parts, LM_GIS_gi006_Parts, LM_GIS_gi007_Parts, LM_GIS_gi008_Parts, LM_GIS_gi009_Parts, LM_GIS_gi010_Parts, LM_GIS_gi011_Parts, LM_GIS_gi012_Parts, LM_GIS_gi013_Parts, LM_GIS_gi014_Parts, LM_GIS_gi015_Parts, LM_GIS_gi016_Parts, LM_GIS_gi017_Parts, LM_GIS_gi018_Parts, LM_GIS_gi019_Parts, LM_GIS_gi020_Parts, LM_GIS_gi021_Parts, LM_GI_Parts)
Dim BP_Plas_Over: BP_Plas_Over=Array(BM_Plas_Over, LM_L_L40_Plas_Over, LM_GIS_gi018_Plas_Over, LM_GIS_gi019_Plas_Over, LM_GIS_gi020_Plas_Over, LM_GIS_gi021_Plas_Over)
Dim BP_Playfield: BP_Playfield=Array(BM_Playfield, LM_GIS_gi001_Playfield, LM_GIS_gi002_Playfield, LM_GIS_gi003_Playfield, LM_GIS_gi004_Playfield, LM_GIS_gi005_Playfield, LM_GIS_gi006_Playfield, LM_GIS_gi007_Playfield, LM_GIS_gi008_Playfield, LM_GIS_gi009_Playfield, LM_GIS_gi010_Playfield, LM_GIS_gi011_Playfield, LM_GIS_gi012_Playfield, LM_GIS_gi013_Playfield, LM_GIS_gi014_Playfield, LM_GIS_gi015_Playfield, LM_GIS_gi016_Playfield, LM_GIS_gi017_Playfield, LM_GIS_gi018_Playfield, LM_GIS_gi019_Playfield, LM_GIS_gi020_Playfield, LM_GIS_gi021_Playfield, LM_GI_Playfield)
Dim BP_RSling1: BP_RSling1=Array(BM_RSling1, LM_GIS_gi001_RSling1, LM_GIS_gi002_RSling1, LM_GIS_gi003_RSling1, LM_GIS_gi004_RSling1)
Dim BP_RSling2: BP_RSling2=Array(BM_RSling2, LM_GIS_gi001_RSling2, LM_GIS_gi002_RSling2, LM_GIS_gi003_RSling2, LM_GIS_gi004_RSling2, LM_GIS_gi008_RSling2)
Dim BP_RSlingArm: BP_RSlingArm=Array(BM_RSlingArm, LM_GIS_gi001_RSlingArm, LM_GIS_gi002_RSlingArm, LM_GIS_gi003_RSlingArm, LM_GIS_gi004_RSlingArm)
Dim BP_SpinL: BP_SpinL=Array(BM_SpinL, LM_GIS_gi007_SpinL, LM_GIS_gi012_SpinL, LM_GIS_gi013_SpinL)
Dim BP_SpinL_Rod: BP_SpinL_Rod=Array(BM_SpinL_Rod, LM_GIS_gi007_SpinL_Rod, LM_GIS_gi012_SpinL_Rod, LM_GIS_gi013_SpinL_Rod)
Dim BP_SpinR: BP_SpinR=Array(BM_SpinR, LM_GIS_gi010_SpinR, LM_GIS_gi011_SpinR)
Dim BP_SpinR_Rod: BP_SpinR_Rod=Array(BM_SpinR_Rod, LM_GIS_gi010_SpinR_Rod, LM_GIS_gi011_SpinR_Rod)
Dim BP_sw14: BP_sw14=Array(BM_sw14, LM_GIS_gi012_sw14)
Dim BP_sw15: BP_sw15=Array(BM_sw15, LM_GIS_gi013_sw15, LM_GIS_gi016_sw15)
Dim BP_sw16: BP_sw16=Array(BM_sw16, LM_L_L54_sw16, LM_GIS_gi013_sw16, LM_GIS_gi014_sw16, LM_GIS_gi016_sw16, LM_GIS_gi017_sw16, LM_GI_sw16)
Dim BP_sw17: BP_sw17=Array(BM_sw17, LM_GIS_gi015_sw17, LM_GIS_gi016_sw17, LM_GIS_gi017_sw17)
Dim BP_sw18: BP_sw18=Array(BM_sw18, LM_GIS_gi015_sw18, LM_GIS_gi016_sw18, LM_GIS_gi017_sw18, LM_GIS_gi018_sw18, LM_GIS_gi019_sw18)
Dim BP_sw19: BP_sw19=Array(BM_sw19, LM_GIS_gi016_sw19, LM_GIS_gi017_sw19, LM_GIS_gi018_sw19, LM_GIS_gi019_sw19, LM_GIS_gi020_sw19)
Dim BP_sw20: BP_sw20=Array(BM_sw20, LM_GIS_gi016_sw20, LM_GIS_gi017_sw20, LM_GIS_gi018_sw20, LM_GIS_gi019_sw20, LM_GIS_gi020_sw20)
Dim BP_sw21: BP_sw21=Array(BM_sw21, LM_L_L52_sw21, LM_L_L62_sw21, LM_GIS_gi005_sw21, LM_GIS_gi006_sw21, LM_GIS_gi007_sw21, LM_GIS_gi010_sw21, LM_GIS_gi011_sw21, LM_GIS_gi012_sw21, LM_GI_sw21)
Dim BP_sw22: BP_sw22=Array(BM_sw22, LM_L_L07_sw22, LM_L_L23_sw22, LM_L_L52_sw22, LM_L_L62_sw22, LM_GIS_gi005_sw22, LM_GIS_gi006_sw22, LM_GIS_gi007_sw22, LM_GIS_gi008_sw22, LM_GIS_gi010_sw22, LM_GIS_gi011_sw22, LM_GIS_gi012_sw22, LM_GI_sw22)
Dim BP_sw23: BP_sw23=Array(BM_sw23, LM_L_L07_sw23, LM_L_L23_sw23, LM_L_L52_sw23, LM_L_L62_sw23, LM_GIS_gi005_sw23, LM_GIS_gi006_sw23, LM_GIS_gi007_sw23, LM_GIS_gi008_sw23, LM_GIS_gi009_sw23, LM_GIS_gi010_sw23, LM_GIS_gi011_sw23, LM_GI_sw23)
Dim BP_sw24: BP_sw24=Array(BM_sw24, LM_L_L07_sw24, LM_L_L23_sw24, LM_L_L52_sw24, LM_L_L62_sw24, LM_GIS_gi002_sw24, LM_GIS_gi004_sw24, LM_GIS_gi005_sw24, LM_GIS_gi006_sw24, LM_GIS_gi007_sw24, LM_GIS_gi008_sw24, LM_GIS_gi009_sw24, LM_GIS_gi010_sw24, LM_GIS_gi011_sw24)
Dim BP_sw25: BP_sw25=Array(BM_sw25, LM_GIS_gi018_sw25, LM_GI_sw25)
Dim BP_sw26: BP_sw26=Array(BM_sw26, LM_GIS_gi011_sw26, LM_GIS_gi018_sw26, LM_GI_sw26)
Dim BP_sw27: BP_sw27=Array(BM_sw27, LM_GIS_gi010_sw27, LM_GIS_gi011_sw27)
Dim BP_sw28: BP_sw28=Array(BM_sw28, LM_GIS_gi019_sw28, LM_GIS_gi020_sw28, LM_GI_sw28)
'Dim BP_sw30: BP_sw30=Array(BM_sw30, LM_L_L08_sw30, LM_L_L46_sw30, LM_L_L55_sw30, LM_L_L56_sw30, LM_GIS_gi009_sw30, LM_GIS_gi010_sw30, LM_GIS_gi011_sw30, LM_GI_sw30)
'Dim BP_sw31: BP_sw31=Array(BM_sw31, LM_L_L07_sw31, LM_L_L08_sw31, LM_L_L23_sw31, LM_L_L46_sw31, LM_L_L55_sw31, LM_L_L56_sw31, LM_GIS_gi009_sw31, LM_GIS_gi010_sw31, LM_GIS_gi011_sw31, LM_GI_sw31)
'Dim BP_sw32: BP_sw32=Array(BM_sw32, LM_L_L07_sw32, LM_L_L08_sw32, LM_L_L23_sw32, LM_L_L39_sw32, LM_L_L46_sw32, LM_L_L55_sw32, LM_GIS_gi009_sw32, LM_GIS_gi010_sw32, LM_GIS_gi011_sw32, LM_GI_sw32)
Dim BP_sw34: BP_sw34=Array(BM_sw34, LM_GIS_gi002_sw34)
Dim BP_sw35: BP_sw35=Array(BM_sw35, LM_GIS_gi004_sw35)
Dim BP_sw36: BP_sw36=Array(BM_sw36, LM_GIS_gi001_sw36, LM_GIS_gi002_sw36)
Dim BP_sw37: BP_sw37=Array(BM_sw37, LM_GIS_gi003_sw37, LM_GIS_gi004_sw37)
Dim BP_sw39: BP_sw39=Array(BM_sw39, LM_GIS_gi016_sw39, LM_GIS_gi017_sw39, LM_GIS_gi018_sw39, LM_GIS_gi019_sw39, LM_GIS_gi020_sw39, LM_GIS_gi021_sw39, LM_GI_sw39)
Dim BP_sw40: BP_sw40=Array(BM_sw40, LM_GIS_gi004_sw40, LM_GIS_gi008_sw40, LM_GIS_gi009_sw40)
' Arrays per lighting scenario
Dim BL_GI: BL_GI=Array(LM_GI_Bumper1_Ring, LM_GI_Bumper2_Ring, LM_GI_Bumper3_Ring, LM_GI_Gate01_Flap, LM_GI_Gate02_Flap, LM_GI_Inserts, LM_GI_Parts, LM_GI_Playfield, LM_GI_sw16, LM_GI_sw21, LM_GI_sw22, LM_GI_sw23, LM_GI_sw25, LM_GI_sw26, LM_GI_sw28, LM_GI_sw39)
Dim BL_GIS_gi001: BL_GIS_gi001=Array(LM_GIS_gi001_FlipperL, LM_GIS_gi001_FlipperLup, LM_GIS_gi001_Inserts, LM_GIS_gi001_LSling1, LM_GIS_gi001_LSling2, LM_GIS_gi001_LSlingArm, LM_GIS_gi001_Parts, LM_GIS_gi001_Playfield, LM_GIS_gi001_RSling1, LM_GIS_gi001_RSling2, LM_GIS_gi001_RSlingArm, LM_GIS_gi001_sw36)
Dim BL_GIS_gi002: BL_GIS_gi002=Array(LM_GIS_gi002_FlipperL, LM_GIS_gi002_FlipperLup, LM_GIS_gi002_Inserts, LM_GIS_gi002_LSling1, LM_GIS_gi002_LSling2, LM_GIS_gi002_LSlingArm, LM_GIS_gi002_Parts, LM_GIS_gi002_Playfield, LM_GIS_gi002_RSling1, LM_GIS_gi002_RSling2, LM_GIS_gi002_RSlingArm, LM_GIS_gi002_sw24, LM_GIS_gi002_sw34, LM_GIS_gi002_sw36)
Dim BL_GIS_gi003: BL_GIS_gi003=Array(LM_GIS_gi003_FlipperR, LM_GIS_gi003_FlipperRup, LM_GIS_gi003_Inserts, LM_GIS_gi003_LSling1, LM_GIS_gi003_LSling2, LM_GIS_gi003_LSlingArm, LM_GIS_gi003_Parts, LM_GIS_gi003_Playfield, LM_GIS_gi003_RSling1, LM_GIS_gi003_RSling2, LM_GIS_gi003_RSlingArm, LM_GIS_gi003_sw37)
Dim BL_GIS_gi004: BL_GIS_gi004=Array(LM_GIS_gi004_Inserts, LM_GIS_gi004_LSling1, LM_GIS_gi004_LSling2, LM_GIS_gi004_LSlingArm, LM_GIS_gi004_Parts, LM_GIS_gi004_Playfield, LM_GIS_gi004_RSling1, LM_GIS_gi004_RSling2, LM_GIS_gi004_RSlingArm, LM_GIS_gi004_sw24, LM_GIS_gi004_sw35, LM_GIS_gi004_sw37, LM_GIS_gi004_sw40)
Dim BL_GIS_gi005: BL_GIS_gi005=Array(LM_GIS_gi005_Inserts, LM_GIS_gi005_LSling1, LM_GIS_gi005_LSlingArm, LM_GIS_gi005_Parts, LM_GIS_gi005_Playfield, LM_GIS_gi005_sw21, LM_GIS_gi005_sw22, LM_GIS_gi005_sw23, LM_GIS_gi005_sw24)
Dim BL_GIS_gi006: BL_GIS_gi006=Array(LM_GIS_gi006_Inserts, LM_GIS_gi006_Parts, LM_GIS_gi006_Playfield, LM_GIS_gi006_sw21, LM_GIS_gi006_sw22, LM_GIS_gi006_sw23, LM_GIS_gi006_sw24)
Dim BL_GIS_gi007: BL_GIS_gi007=Array(LM_GIS_gi007_Parts, LM_GIS_gi007_Playfield, LM_GIS_gi007_SpinL, LM_GIS_gi007_SpinL_Rod, LM_GIS_gi007_sw21, LM_GIS_gi007_sw22, LM_GIS_gi007_sw23, LM_GIS_gi007_sw24)
Dim BL_GIS_gi008: BL_GIS_gi008=Array(LM_GIS_gi008_Inserts, LM_GIS_gi008_Parts, LM_GIS_gi008_Playfield, LM_GIS_gi008_RSling2, LM_GIS_gi008_sw22, LM_GIS_gi008_sw23, LM_GIS_gi008_sw24, LM_GIS_gi008_sw40)
Dim BL_GIS_gi009: BL_GIS_gi009=Array(LM_GIS_gi009_Parts, LM_GIS_gi009_Playfield, LM_GIS_gi009_sw23, LM_GIS_gi009_sw24, LM_GIS_gi009_sw40)
Dim BL_GIS_gi010: BL_GIS_gi010=Array(LM_GIS_gi010_Parts, LM_GIS_gi010_Playfield, LM_GIS_gi010_SpinR, LM_GIS_gi010_SpinR_Rod, LM_GIS_gi010_sw21, LM_GIS_gi010_sw22, LM_GIS_gi010_sw23, LM_GIS_gi010_sw24, LM_GIS_gi010_sw27)
Dim BL_GIS_gi011: BL_GIS_gi011=Array(LM_GIS_gi011_Parts, LM_GIS_gi011_Playfield, LM_GIS_gi011_SpinR, LM_GIS_gi011_SpinR_Rod, LM_GIS_gi011_sw21, LM_GIS_gi011_sw22, LM_GIS_gi011_sw23, LM_GIS_gi011_sw24, LM_GIS_gi011_sw26, LM_GIS_gi011_sw27)
Dim BL_GIS_gi012: BL_GIS_gi012=Array(LM_GIS_gi012_Bumper2_Ring, LM_GIS_gi012_Inserts, LM_GIS_gi012_Parts, LM_GIS_gi012_Playfield, LM_GIS_gi012_SpinL, LM_GIS_gi012_SpinL_Rod, LM_GIS_gi012_sw14, LM_GIS_gi012_sw21, LM_GIS_gi012_sw22)
Dim BL_GIS_gi013: BL_GIS_gi013=Array(LM_GIS_gi013_Inserts, LM_GIS_gi013_Parts, LM_GIS_gi013_Playfield, LM_GIS_gi013_SpinL, LM_GIS_gi013_SpinL_Rod, LM_GIS_gi013_sw15, LM_GIS_gi013_sw16)
Dim BL_GIS_gi014: BL_GIS_gi014=Array(LM_GIS_gi014_Gate01_Flap, LM_GIS_gi014_Parts, LM_GIS_gi014_Playfield, LM_GIS_gi014_sw16)
Dim BL_GIS_gi015: BL_GIS_gi015=Array(LM_GIS_gi015_Gate01_Flap, LM_GIS_gi015_Parts, LM_GIS_gi015_Playfield, LM_GIS_gi015_sw17, LM_GIS_gi015_sw18)
Dim BL_GIS_gi016: BL_GIS_gi016=Array(LM_GIS_gi016_Gate01_Flap, LM_GIS_gi016_Parts, LM_GIS_gi016_Playfield, LM_GIS_gi016_sw15, LM_GIS_gi016_sw16, LM_GIS_gi016_sw17, LM_GIS_gi016_sw18, LM_GIS_gi016_sw19, LM_GIS_gi016_sw20, LM_GIS_gi016_sw39)
Dim BL_GIS_gi017: BL_GIS_gi017=Array(LM_GIS_gi017_Gate02_Flap, LM_GIS_gi017_Parts, LM_GIS_gi017_Playfield, LM_GIS_gi017_sw16, LM_GIS_gi017_sw17, LM_GIS_gi017_sw18, LM_GIS_gi017_sw19, LM_GIS_gi017_sw20, LM_GIS_gi017_sw39)
Dim BL_GIS_gi018: BL_GIS_gi018=Array(LM_GIS_gi018_Gate02_Flap, LM_GIS_gi018_Inserts, LM_GIS_gi018_Parts, LM_GIS_gi018_Plas_Over, LM_GIS_gi018_Playfield, LM_GIS_gi018_sw18, LM_GIS_gi018_sw19, LM_GIS_gi018_sw20, LM_GIS_gi018_sw25, LM_GIS_gi018_sw26, LM_GIS_gi018_sw39)
Dim BL_GIS_gi019: BL_GIS_gi019=Array(LM_GIS_gi019_Inserts, LM_GIS_gi019_Parts, LM_GIS_gi019_Plas_Over, LM_GIS_gi019_Playfield, LM_GIS_gi019_sw18, LM_GIS_gi019_sw19, LM_GIS_gi019_sw20, LM_GIS_gi019_sw28, LM_GIS_gi019_sw39)
Dim BL_GIS_gi020: BL_GIS_gi020=Array(LM_GIS_gi020_Gate02_Flap, LM_GIS_gi020_Inserts, LM_GIS_gi020_Parts, LM_GIS_gi020_Plas_Over, LM_GIS_gi020_Playfield, LM_GIS_gi020_sw19, LM_GIS_gi020_sw20, LM_GIS_gi020_sw28, LM_GIS_gi020_sw39)
Dim BL_GIS_gi021: BL_GIS_gi021=Array(LM_GIS_gi021_Gate02_Flap, LM_GIS_gi021_Parts, LM_GIS_gi021_Plas_Over, LM_GIS_gi021_Playfield, LM_GIS_gi021_sw39)
Dim BL_L_L01: BL_L_L01=Array(LM_L_L01_Inserts)
Dim BL_L_L02: BL_L_L02=Array(LM_L_L02_Inserts)
Dim BL_L_L03: BL_L_L03=Array(LM_L_L03_Inserts)
Dim BL_L_L04: BL_L_L04=Array(LM_L_L04_Inserts)
Dim BL_L_L05: BL_L_L05=Array(LM_L_L05_Inserts)
Dim BL_L_L07: BL_L_L07=Array(LM_L_L07_Inserts, LM_L_L07_sw22, LM_L_L07_sw23, LM_L_L07_sw24)
Dim BL_L_L08: BL_L_L08=Array(LM_L_L08_Inserts)
Dim BL_L_L09: BL_L_L09=Array(LM_L_L09_Inserts, LM_L_L09_Parts)
Dim BL_L_L11: BL_L_L11=Array(LM_L_L11_Inserts)
Dim BL_L_L12: BL_L_L12=Array(LM_L_L12_Inserts, LM_L_L12_Parts)
Dim BL_L_L14: BL_L_L14=Array(LM_L_L14_Inserts, LM_L_L14_Parts)
Dim BL_L_L17: BL_L_L17=Array(LM_L_L17_Inserts)
Dim BL_L_L18: BL_L_L18=Array(LM_L_L18_Inserts)
Dim BL_L_L19: BL_L_L19=Array(LM_L_L19_Inserts)
Dim BL_L_L20: BL_L_L20=Array(LM_L_L20_Inserts)
Dim BL_L_L21: BL_L_L21=Array(LM_L_L21_Inserts)
Dim BL_L_L22: BL_L_L22=Array(LM_L_L22_Inserts, LM_L_L22_Parts)
Dim BL_L_L23: BL_L_L23=Array(LM_L_L23_Inserts, LM_L_L23_sw22, LM_L_L23_sw23, LM_L_L23_sw24)
Dim BL_L_L24: BL_L_L24=Array(LM_L_L24_Inserts, LM_L_L24_Parts)
Dim BL_L_L28: BL_L_L28=Array(LM_L_L28_Inserts, LM_L_L28_Parts)
Dim BL_L_L30: BL_L_L30=Array(LM_L_L30_Inserts, LM_L_L30_Parts)
Dim BL_L_L31: BL_L_L31=Array(LM_L_L31_Inserts, LM_L_L31_Parts)
Dim BL_L_L33: BL_L_L33=Array(LM_L_L33_Inserts)
Dim BL_L_L34: BL_L_L34=Array(LM_L_L34_Inserts)
Dim BL_L_L35: BL_L_L35=Array(LM_L_L35_Inserts, LM_L_L35_Parts)
Dim BL_L_L36: BL_L_L36=Array(LM_L_L36_Inserts)
Dim BL_L_L37: BL_L_L37=Array(LM_L_L37_Inserts)
Dim BL_L_L38: BL_L_L38=Array(LM_L_L38_Inserts, LM_L_L38_Parts)
Dim BL_L_L39: BL_L_L39=Array(LM_L_L39_Inserts)
Dim BL_L_L40: BL_L_L40=Array(LM_L_L40_Inserts, LM_L_L40_Plas_Over)
Dim BL_L_L44: BL_L_L44=Array(LM_L_L44_Inserts, LM_L_L44_Parts)
Dim BL_L_L46: BL_L_L46=Array(LM_L_L46_Inserts)
Dim BL_L_L47: BL_L_L47=Array(LM_L_L47_Inserts)
Dim BL_L_L49: BL_L_L49=Array(LM_L_L49_Inserts)
Dim BL_L_L50: BL_L_L50=Array(LM_L_L50_Inserts)
Dim BL_L_L51: BL_L_L51=Array(LM_L_L51_Inserts)
Dim BL_L_L52: BL_L_L52=Array(LM_L_L52_Inserts, LM_L_L52_sw21, LM_L_L52_sw22, LM_L_L52_sw23, LM_L_L52_sw24)
Dim BL_L_L53: BL_L_L53=Array(LM_L_L53_Inserts)
Dim BL_L_L54: BL_L_L54=Array(LM_L_L54_Inserts, LM_L_L54_Parts, LM_L_L54_sw16)
Dim BL_L_L55: BL_L_L55=Array(LM_L_L55_Inserts)
Dim BL_L_L56: BL_L_L56=Array(LM_L_L56_Inserts, LM_L_L56_Parts)
Dim BL_L_L59: BL_L_L59=Array(LM_L_L59_Inserts, LM_L_L59_Parts)
Dim BL_L_L60: BL_L_L60=Array(LM_L_L60_Inserts, LM_L_L60_Parts)
Dim BL_L_L62: BL_L_L62=Array(LM_L_L62_Inserts, LM_L_L62_Parts, LM_L_L62_sw21, LM_L_L62_sw22, LM_L_L62_sw23, LM_L_L62_sw24)
Dim BL_Room: BL_Room=Array(BM_Bumper1_Ring, BM_Bumper2_Ring, BM_Bumper3_Ring, BM_FlipperL, BM_FlipperLup, BM_FlipperR, BM_FlipperRup, BM_Gate01_Flap, BM_Gate02_Flap, BM_Inserts, BM_LSling1, BM_LSling2, BM_LSlingArm, BM_Parts, BM_Plas_Over, BM_Playfield, BM_RSling1, BM_RSling2, BM_RSlingArm, BM_SpinL, BM_SpinL_Rod, BM_SpinR, BM_SpinR_Rod, BM_sw14, BM_sw15, BM_sw16, BM_sw17, BM_sw18, BM_sw19, BM_sw20, BM_sw21, BM_sw22, BM_sw23, BM_sw24, BM_sw25, BM_sw26, BM_sw27, BM_sw28, BM_sw34, BM_sw35, BM_sw36, BM_sw37, BM_sw39, BM_sw40)
' Global arrays
Dim BG_Bakemap: BG_Bakemap=Array(BM_Bumper1_Ring, BM_Bumper2_Ring, BM_Bumper3_Ring, BM_FlipperL, BM_FlipperLup, BM_FlipperR, BM_FlipperRup, BM_Gate01_Flap, BM_Gate02_Flap, BM_Inserts, BM_LSling1, BM_LSling2, BM_LSlingArm, BM_Parts, BM_Plas_Over, BM_Playfield, BM_RSling1, BM_RSling2, BM_RSlingArm, BM_SpinL, BM_SpinL_Rod, BM_SpinR, BM_SpinR_Rod, BM_sw14, BM_sw15, BM_sw16, BM_sw17, BM_sw18, BM_sw19, BM_sw20, BM_sw21, BM_sw22, BM_sw23, BM_sw24, BM_sw25, BM_sw26, BM_sw27, BM_sw28, BM_sw34, BM_sw35, BM_sw36, BM_sw37, BM_sw39, BM_sw40)
Dim BG_Lightmap: BG_Lightmap=Array(LM_GI_Bumper1_Ring, LM_GI_Bumper2_Ring, LM_GI_Bumper3_Ring, LM_GI_Gate01_Flap, LM_GI_Gate02_Flap, LM_GI_Inserts, LM_GI_Parts, LM_GI_Playfield, LM_GI_sw16, LM_GI_sw21, LM_GI_sw22, LM_GI_sw23, LM_GI_sw25, LM_GI_sw26, LM_GI_sw28, LM_GI_sw39, LM_GIS_gi001_FlipperL, LM_GIS_gi001_FlipperLup, LM_GIS_gi001_Inserts, LM_GIS_gi001_LSling1, LM_GIS_gi001_LSling2, LM_GIS_gi001_LSlingArm, LM_GIS_gi001_Parts, LM_GIS_gi001_Playfield, LM_GIS_gi001_RSling1, LM_GIS_gi001_RSling2, LM_GIS_gi001_RSlingArm, LM_GIS_gi001_sw36, LM_GIS_gi002_FlipperL, LM_GIS_gi002_FlipperLup, LM_GIS_gi002_Inserts, LM_GIS_gi002_LSling1, LM_GIS_gi002_LSling2, LM_GIS_gi002_LSlingArm, LM_GIS_gi002_Parts, LM_GIS_gi002_Playfield, LM_GIS_gi002_RSling1, LM_GIS_gi002_RSling2, LM_GIS_gi002_RSlingArm, LM_GIS_gi002_sw24, LM_GIS_gi002_sw34, LM_GIS_gi002_sw36, LM_GIS_gi003_FlipperR, LM_GIS_gi003_FlipperRup, LM_GIS_gi003_Inserts, LM_GIS_gi003_LSling1, LM_GIS_gi003_LSling2, LM_GIS_gi003_LSlingArm, _
  LM_GIS_gi003_Parts, LM_GIS_gi003_Playfield, LM_GIS_gi003_RSling1, LM_GIS_gi003_RSling2, LM_GIS_gi003_RSlingArm, LM_GIS_gi003_sw37, LM_GIS_gi004_Inserts, LM_GIS_gi004_LSling1, LM_GIS_gi004_LSling2, LM_GIS_gi004_LSlingArm, LM_GIS_gi004_Parts, LM_GIS_gi004_Playfield, LM_GIS_gi004_RSling1, LM_GIS_gi004_RSling2, LM_GIS_gi004_RSlingArm, LM_GIS_gi004_sw24, LM_GIS_gi004_sw35, LM_GIS_gi004_sw37, LM_GIS_gi004_sw40, LM_GIS_gi005_Inserts, LM_GIS_gi005_LSling1, LM_GIS_gi005_LSlingArm, LM_GIS_gi005_Parts, LM_GIS_gi005_Playfield, LM_GIS_gi005_sw21, LM_GIS_gi005_sw22, LM_GIS_gi005_sw23, LM_GIS_gi005_sw24, LM_GIS_gi006_Inserts, LM_GIS_gi006_Parts, LM_GIS_gi006_Playfield, LM_GIS_gi006_sw21, LM_GIS_gi006_sw22, LM_GIS_gi006_sw23, LM_GIS_gi006_sw24, LM_GIS_gi007_Parts, LM_GIS_gi007_Playfield, LM_GIS_gi007_SpinL, LM_GIS_gi007_SpinL_Rod, LM_GIS_gi007_sw21, LM_GIS_gi007_sw22, LM_GIS_gi007_sw23, LM_GIS_gi007_sw24, LM_GIS_gi008_Inserts, LM_GIS_gi008_Parts, LM_GIS_gi008_Playfield, LM_GIS_gi008_RSling2, LM_GIS_gi008_sw22, _
  LM_GIS_gi008_sw23, LM_GIS_gi008_sw24, LM_GIS_gi008_sw40, LM_GIS_gi009_Parts, LM_GIS_gi009_Playfield, LM_GIS_gi009_sw23, LM_GIS_gi009_sw24, LM_GIS_gi009_sw40, LM_GIS_gi010_Parts, LM_GIS_gi010_Playfield, LM_GIS_gi010_SpinR, LM_GIS_gi010_SpinR_Rod, LM_GIS_gi010_sw21, LM_GIS_gi010_sw22, LM_GIS_gi010_sw23, LM_GIS_gi010_sw24, LM_GIS_gi010_sw27, LM_GIS_gi011_Parts, LM_GIS_gi011_Playfield, LM_GIS_gi011_SpinR, LM_GIS_gi011_SpinR_Rod, LM_GIS_gi011_sw21, LM_GIS_gi011_sw22, LM_GIS_gi011_sw23, LM_GIS_gi011_sw24, LM_GIS_gi011_sw26, LM_GIS_gi011_sw27, LM_GIS_gi012_Bumper2_Ring, LM_GIS_gi012_Inserts, LM_GIS_gi012_Parts, LM_GIS_gi012_Playfield, LM_GIS_gi012_SpinL, LM_GIS_gi012_SpinL_Rod, LM_GIS_gi012_sw14, LM_GIS_gi012_sw21, LM_GIS_gi012_sw22, LM_GIS_gi013_Inserts, LM_GIS_gi013_Parts, LM_GIS_gi013_Playfield, LM_GIS_gi013_SpinL, LM_GIS_gi013_SpinL_Rod, _
  LM_GIS_gi013_sw15, LM_GIS_gi013_sw16, LM_GIS_gi014_Gate01_Flap, LM_GIS_gi014_Parts, LM_GIS_gi014_Playfield, LM_GIS_gi014_sw16, LM_GIS_gi015_Gate01_Flap, LM_GIS_gi015_Parts, LM_GIS_gi015_Playfield, LM_GIS_gi015_sw17, LM_GIS_gi015_sw18, LM_GIS_gi016_Gate01_Flap, LM_GIS_gi016_Parts, LM_GIS_gi016_Playfield, LM_GIS_gi016_sw15, LM_GIS_gi016_sw16, LM_GIS_gi016_sw17, LM_GIS_gi016_sw18, LM_GIS_gi016_sw19, LM_GIS_gi016_sw20, LM_GIS_gi016_sw39, LM_GIS_gi017_Gate02_Flap, LM_GIS_gi017_Parts, LM_GIS_gi017_Playfield, LM_GIS_gi017_sw16, LM_GIS_gi017_sw17, LM_GIS_gi017_sw18, LM_GIS_gi017_sw19, LM_GIS_gi017_sw20, LM_GIS_gi017_sw39, LM_GIS_gi018_Gate02_Flap, LM_GIS_gi018_Inserts, LM_GIS_gi018_Parts, LM_GIS_gi018_Plas_Over, LM_GIS_gi018_Playfield, LM_GIS_gi018_sw18, LM_GIS_gi018_sw19, LM_GIS_gi018_sw20, LM_GIS_gi018_sw25, LM_GIS_gi018_sw26, LM_GIS_gi018_sw39, LM_GIS_gi019_Inserts, LM_GIS_gi019_Parts, LM_GIS_gi019_Plas_Over, LM_GIS_gi019_Playfield, LM_GIS_gi019_sw18, LM_GIS_gi019_sw19, LM_GIS_gi019_sw20, LM_GIS_gi019_sw28, _
  LM_GIS_gi019_sw39, LM_GIS_gi020_Gate02_Flap, LM_GIS_gi020_Inserts, LM_GIS_gi020_Parts, LM_GIS_gi020_Plas_Over, LM_GIS_gi020_Playfield, LM_GIS_gi020_sw19, LM_GIS_gi020_sw20, LM_GIS_gi020_sw28, LM_GIS_gi020_sw39, LM_GIS_gi021_Gate02_Flap, LM_GIS_gi021_Parts, LM_GIS_gi021_Plas_Over, LM_GIS_gi021_Playfield, LM_GIS_gi021_sw39, LM_L_L01_Inserts, LM_L_L02_Inserts, LM_L_L03_Inserts, LM_L_L04_Inserts, LM_L_L05_Inserts, LM_L_L07_Inserts, LM_L_L07_sw22, LM_L_L07_sw23, LM_L_L07_sw24, LM_L_L08_Inserts, LM_L_L09_Inserts, LM_L_L09_Parts, LM_L_L11_Inserts, LM_L_L12_Inserts, LM_L_L12_Parts, LM_L_L14_Inserts, LM_L_L14_Parts, LM_L_L17_Inserts, LM_L_L18_Inserts, LM_L_L19_Inserts, LM_L_L20_Inserts, LM_L_L21_Inserts, LM_L_L22_Inserts, LM_L_L22_Parts, LM_L_L23_Inserts, LM_L_L23_sw22, LM_L_L23_sw23, LM_L_L23_sw24, LM_L_L24_Inserts, LM_L_L24_Parts, LM_L_L28_Inserts, LM_L_L28_Parts, LM_L_L30_Inserts, LM_L_L30_Parts, _
  LM_L_L31_Inserts, LM_L_L31_Parts, LM_L_L33_Inserts, LM_L_L34_Inserts, LM_L_L35_Inserts, LM_L_L35_Parts, LM_L_L36_Inserts, LM_L_L37_Inserts, LM_L_L38_Inserts, LM_L_L38_Parts, LM_L_L39_Inserts, LM_L_L40_Inserts, LM_L_L40_Plas_Over, LM_L_L44_Inserts, LM_L_L44_Parts, LM_L_L46_Inserts, LM_L_L47_Inserts, LM_L_L49_Inserts, LM_L_L50_Inserts, LM_L_L51_Inserts, LM_L_L52_Inserts, LM_L_L52_sw21, LM_L_L52_sw22, LM_L_L52_sw23, LM_L_L52_sw24, LM_L_L53_Inserts, LM_L_L54_Inserts, LM_L_L54_Parts, LM_L_L54_sw16, LM_L_L55_Inserts, LM_L_L56_Inserts, LM_L_L56_Parts, LM_L_L59_Inserts, LM_L_L59_Parts, LM_L_L60_Inserts, LM_L_L60_Parts, LM_L_L62_Inserts, LM_L_L62_Parts, LM_L_L62_sw21, LM_L_L62_sw22, LM_L_L62_sw23, LM_L_L62_sw24)
Dim BG_All: BG_All=Array(BM_Bumper1_Ring, BM_Bumper2_Ring, BM_Bumper3_Ring, BM_FlipperL, BM_FlipperLup, BM_FlipperR, BM_FlipperRup, BM_Gate01_Flap, BM_Gate02_Flap, BM_Inserts, BM_LSling1, BM_LSling2, BM_LSlingArm, BM_Parts, BM_Plas_Over, BM_Playfield, BM_RSling1, BM_RSling2, BM_RSlingArm, BM_SpinL, BM_SpinL_Rod, BM_SpinR, BM_SpinR_Rod, BM_sw14, BM_sw15, BM_sw16, BM_sw17, BM_sw18, BM_sw19, BM_sw20, BM_sw21, BM_sw22, BM_sw23, BM_sw24, BM_sw25, BM_sw26, BM_sw27, BM_sw28, BM_sw34, BM_sw35, BM_sw36, BM_sw37, BM_sw39, BM_sw40, LM_GI_Bumper1_Ring, LM_GI_Bumper2_Ring, LM_GI_Bumper3_Ring, LM_GI_Gate01_Flap, LM_GI_Gate02_Flap, LM_GI_Inserts, LM_GI_Parts, LM_GI_Playfield, LM_GI_sw16, LM_GI_sw21, LM_GI_sw22, LM_GI_sw23, LM_GI_sw25, LM_GI_sw26, LM_GI_sw28, LM_GI_sw39, LM_GIS_gi001_FlipperL, LM_GIS_gi001_FlipperLup, LM_GIS_gi001_Inserts, LM_GIS_gi001_LSling1, LM_GIS_gi001_LSling2, LM_GIS_gi001_LSlingArm, LM_GIS_gi001_Parts, LM_GIS_gi001_Playfield, _
  LM_GIS_gi001_RSling1, LM_GIS_gi001_RSling2, LM_GIS_gi001_RSlingArm, LM_GIS_gi001_sw36, LM_GIS_gi002_FlipperL, LM_GIS_gi002_FlipperLup, LM_GIS_gi002_Inserts, LM_GIS_gi002_LSling1, LM_GIS_gi002_LSling2, LM_GIS_gi002_LSlingArm, LM_GIS_gi002_Parts, LM_GIS_gi002_Playfield, LM_GIS_gi002_RSling1, LM_GIS_gi002_RSling2, LM_GIS_gi002_RSlingArm, LM_GIS_gi002_sw24, LM_GIS_gi002_sw34, LM_GIS_gi002_sw36, LM_GIS_gi003_FlipperR, LM_GIS_gi003_FlipperRup, LM_GIS_gi003_Inserts, LM_GIS_gi003_LSling1, LM_GIS_gi003_LSling2, LM_GIS_gi003_LSlingArm, LM_GIS_gi003_Parts, LM_GIS_gi003_Playfield, LM_GIS_gi003_RSling1, LM_GIS_gi003_RSling2, LM_GIS_gi003_RSlingArm, LM_GIS_gi003_sw37, LM_GIS_gi004_Inserts, LM_GIS_gi004_LSling1, LM_GIS_gi004_LSling2, LM_GIS_gi004_LSlingArm, LM_GIS_gi004_Parts, LM_GIS_gi004_Playfield, LM_GIS_gi004_RSling1, LM_GIS_gi004_RSling2, LM_GIS_gi004_RSlingArm, LM_GIS_gi004_sw24, LM_GIS_gi004_sw35, LM_GIS_gi004_sw37, LM_GIS_gi004_sw40, LM_GIS_gi005_Inserts, LM_GIS_gi005_LSling1, LM_GIS_gi005_LSlingArm, _
  LM_GIS_gi005_Parts, LM_GIS_gi005_Playfield, LM_GIS_gi005_sw21, LM_GIS_gi005_sw22, LM_GIS_gi005_sw23, LM_GIS_gi005_sw24, LM_GIS_gi006_Inserts, LM_GIS_gi006_Parts, LM_GIS_gi006_Playfield, LM_GIS_gi006_sw21, LM_GIS_gi006_sw22, LM_GIS_gi006_sw23, LM_GIS_gi006_sw24, LM_GIS_gi007_Parts, LM_GIS_gi007_Playfield, LM_GIS_gi007_SpinL, LM_GIS_gi007_SpinL_Rod, LM_GIS_gi007_sw21, LM_GIS_gi007_sw22, LM_GIS_gi007_sw23, LM_GIS_gi007_sw24, LM_GIS_gi008_Inserts, LM_GIS_gi008_Parts, LM_GIS_gi008_Playfield, LM_GIS_gi008_RSling2, LM_GIS_gi008_sw22, LM_GIS_gi008_sw23, LM_GIS_gi008_sw24, LM_GIS_gi008_sw40, LM_GIS_gi009_Parts, LM_GIS_gi009_Playfield, LM_GIS_gi009_sw23, LM_GIS_gi009_sw24, LM_GIS_gi009_sw40, LM_GIS_gi010_Parts, LM_GIS_gi010_Playfield, LM_GIS_gi010_SpinR, LM_GIS_gi010_SpinR_Rod, LM_GIS_gi010_sw21, LM_GIS_gi010_sw22, LM_GIS_gi010_sw23, LM_GIS_gi010_sw24, LM_GIS_gi010_sw27, LM_GIS_gi011_Parts, _
  LM_GIS_gi011_Playfield, LM_GIS_gi011_SpinR, LM_GIS_gi011_SpinR_Rod, LM_GIS_gi011_sw21, LM_GIS_gi011_sw22, LM_GIS_gi011_sw23, LM_GIS_gi011_sw24, LM_GIS_gi011_sw26, LM_GIS_gi011_sw27, LM_GIS_gi012_Bumper2_Ring, LM_GIS_gi012_Inserts, LM_GIS_gi012_Parts, LM_GIS_gi012_Playfield, LM_GIS_gi012_SpinL, LM_GIS_gi012_SpinL_Rod, LM_GIS_gi012_sw14, LM_GIS_gi012_sw21, LM_GIS_gi012_sw22, LM_GIS_gi013_Inserts, LM_GIS_gi013_Parts, LM_GIS_gi013_Playfield, LM_GIS_gi013_SpinL, LM_GIS_gi013_SpinL_Rod, LM_GIS_gi013_sw15, LM_GIS_gi013_sw16, LM_GIS_gi014_Gate01_Flap, LM_GIS_gi014_Parts, LM_GIS_gi014_Playfield, LM_GIS_gi014_sw16, LM_GIS_gi015_Gate01_Flap, LM_GIS_gi015_Parts, LM_GIS_gi015_Playfield, LM_GIS_gi015_sw17, LM_GIS_gi015_sw18, LM_GIS_gi016_Gate01_Flap, LM_GIS_gi016_Parts, LM_GIS_gi016_Playfield, LM_GIS_gi016_sw15, LM_GIS_gi016_sw16, LM_GIS_gi016_sw17, LM_GIS_gi016_sw18, LM_GIS_gi016_sw19, LM_GIS_gi016_sw20, LM_GIS_gi016_sw39, LM_GIS_gi017_Gate02_Flap, _
  LM_GIS_gi017_Parts, LM_GIS_gi017_Playfield, LM_GIS_gi017_sw16, LM_GIS_gi017_sw17, LM_GIS_gi017_sw18, LM_GIS_gi017_sw19, LM_GIS_gi017_sw20, LM_GIS_gi017_sw39, LM_GIS_gi018_Gate02_Flap, LM_GIS_gi018_Inserts, LM_GIS_gi018_Parts, LM_GIS_gi018_Plas_Over, LM_GIS_gi018_Playfield, LM_GIS_gi018_sw18, LM_GIS_gi018_sw19, LM_GIS_gi018_sw20, LM_GIS_gi018_sw25, LM_GIS_gi018_sw26, LM_GIS_gi018_sw39, LM_GIS_gi019_Inserts, LM_GIS_gi019_Parts, LM_GIS_gi019_Plas_Over, LM_GIS_gi019_Playfield, LM_GIS_gi019_sw18, LM_GIS_gi019_sw19, LM_GIS_gi019_sw20, LM_GIS_gi019_sw28, LM_GIS_gi019_sw39, LM_GIS_gi020_Gate02_Flap, LM_GIS_gi020_Inserts, LM_GIS_gi020_Parts, LM_GIS_gi020_Plas_Over, LM_GIS_gi020_Playfield, LM_GIS_gi020_sw19, LM_GIS_gi020_sw20, LM_GIS_gi020_sw28, LM_GIS_gi020_sw39, LM_GIS_gi021_Gate02_Flap, LM_GIS_gi021_Parts, LM_GIS_gi021_Plas_Over, LM_GIS_gi021_Playfield, LM_GIS_gi021_sw39, LM_L_L01_Inserts, LM_L_L02_Inserts, LM_L_L03_Inserts, LM_L_L04_Inserts, LM_L_L05_Inserts, LM_L_L07_Inserts, LM_L_L07_sw22, LM_L_L07_sw23, _
  LM_L_L07_sw24, LM_L_L08_Inserts, LM_L_L09_Inserts, LM_L_L09_Parts, LM_L_L11_Inserts, LM_L_L12_Inserts, LM_L_L12_Parts, LM_L_L14_Inserts, LM_L_L14_Parts, LM_L_L17_Inserts, LM_L_L18_Inserts, LM_L_L19_Inserts, LM_L_L20_Inserts, LM_L_L21_Inserts, LM_L_L22_Inserts, LM_L_L22_Parts, LM_L_L23_Inserts, LM_L_L23_sw22, LM_L_L23_sw23, LM_L_L23_sw24, LM_L_L24_Inserts, LM_L_L24_Parts, LM_L_L28_Inserts, LM_L_L28_Parts, LM_L_L30_Inserts, LM_L_L30_Parts, LM_L_L31_Inserts, LM_L_L31_Parts, LM_L_L33_Inserts, LM_L_L34_Inserts, LM_L_L35_Inserts, LM_L_L35_Parts, LM_L_L36_Inserts, LM_L_L37_Inserts, LM_L_L38_Inserts, LM_L_L38_Parts, LM_L_L39_Inserts, LM_L_L40_Inserts, LM_L_L40_Plas_Over, LM_L_L44_Inserts, LM_L_L44_Parts, LM_L_L46_Inserts, LM_L_L47_Inserts, LM_L_L49_Inserts, LM_L_L50_Inserts, LM_L_L51_Inserts, LM_L_L52_Inserts, LM_L_L52_sw21, LM_L_L52_sw22, LM_L_L52_sw23, _
  LM_L_L52_sw24, LM_L_L53_Inserts, LM_L_L54_Inserts, LM_L_L54_Parts, LM_L_L54_sw16, LM_L_L55_Inserts, LM_L_L56_Inserts, LM_L_L56_Parts, LM_L_L59_Inserts, LM_L_L59_Parts, LM_L_L60_Inserts, LM_L_L60_Parts, LM_L_L62_Inserts, LM_L_L62_Parts, LM_L_L62_sw21, LM_L_L62_sw22, LM_L_L62_sw23, LM_L_L62_sw24)
' VLM  Arrays - End

' VLM VRBG Arrays - Start
' Arrays per baked part
Dim BP_VRBGBG: BP_VRBGBG=Array(VRBGBM_BG, VRBGLM_BGS_go_BG, VRBGLM_BGS_hstd_BG, VRBGLM_BGS_ma_BG, VRBGLM_BGS_sa_BG, VRBGLM_BGS_tilt_BG, VRBGLM_BG_gi_BG)
' Arrays per lighting scenario
Dim BL_VRBGBGS_go: BL_VRBGBGS_go=Array(VRBGLM_BGS_go_BG)
Dim BL_VRBGBGS_hstd: BL_VRBGBGS_hstd=Array(VRBGLM_BGS_hstd_BG)
Dim BL_VRBGBGS_ma: BL_VRBGBGS_ma=Array(VRBGLM_BGS_ma_BG)
Dim BL_VRBGBGS_sa: BL_VRBGBGS_sa=Array(VRBGLM_BGS_sa_BG)
Dim BL_VRBGBGS_tilt: BL_VRBGBGS_tilt=Array(VRBGLM_BGS_tilt_BG)
Dim BL_VRBGBG_gi: BL_VRBGBG_gi=Array(VRBGLM_BG_gi_BG)
Dim BL_VRBGRoom: BL_VRBGRoom=Array(VRBGBM_BG)
' Global arrays
Dim BGVRBG_Bakemap: BGVRBG_Bakemap=Array(VRBGBM_BG)
Dim BGVRBG_Lightmap: BGVRBG_Lightmap=Array(VRBGLM_BGS_go_BG, VRBGLM_BGS_hstd_BG, VRBGLM_BGS_ma_BG, VRBGLM_BGS_sa_BG, VRBGLM_BGS_tilt_BG, VRBGLM_BG_gi_BG)
Dim BGVRBG_All: BGVRBG_All=Array(VRBGBM_BG, VRBGLM_BGS_go_BG, VRBGLM_BGS_hstd_BG, VRBGLM_BGS_ma_BG, VRBGLM_BGS_sa_BG, VRBGLM_BGS_tilt_BG, VRBGLM_BG_gi_BG)
' VLM VRBG Arrays - End

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

' VLM DT Arrays - Start
' Arrays per baked part
Dim BP_DTsw30: BP_DTsw30=Array(DTBM_sw30, DTLM_GIS_gi009_sw30, DTLM_GIS_gi010_sw30, DTLM_GIS_gi011_sw30, DTLM_GIS_gi017_sw30, DTLM_GIS_gi018_sw30, DTLM_GIS_gi019_sw30, DTLM_GIS_gi020_sw30, DTLM_GIS_gi021_sw30)
Dim BP_DTsw31: BP_DTsw31=Array(DTBM_sw31, DTLM_GIS_gi009_sw31, DTLM_GIS_gi010_sw31, DTLM_GIS_gi011_sw31, DTLM_GIS_gi018_sw31, DTLM_GIS_gi019_sw31, DTLM_GIS_gi020_sw31, DTLM_GIS_gi021_sw31)
Dim BP_DTsw32: BP_DTsw32=Array(DTBM_sw32, DTLM_GIS_gi008_sw32, DTLM_GIS_gi009_sw32, DTLM_GIS_gi010_sw32, DTLM_GIS_gi011_sw32, DTLM_GIS_gi019_sw32, DTLM_GIS_gi020_sw32, DTLM_GIS_gi021_sw32)
' Arrays per lighting scenario
Dim BL_DTGIS_gi008: BL_DTGIS_gi008=Array(DTLM_GIS_gi008_sw32)
Dim BL_DTGIS_gi009: BL_DTGIS_gi009=Array(DTLM_GIS_gi009_sw30, DTLM_GIS_gi009_sw31, DTLM_GIS_gi009_sw32)
Dim BL_DTGIS_gi010: BL_DTGIS_gi010=Array(DTLM_GIS_gi010_sw30, DTLM_GIS_gi010_sw31, DTLM_GIS_gi010_sw32)
Dim BL_DTGIS_gi011: BL_DTGIS_gi011=Array(DTLM_GIS_gi011_sw30, DTLM_GIS_gi011_sw31, DTLM_GIS_gi011_sw32)
Dim BL_DTGIS_gi017: BL_DTGIS_gi017=Array(DTLM_GIS_gi017_sw30)
Dim BL_DTGIS_gi018: BL_DTGIS_gi018=Array(DTLM_GIS_gi018_sw30, DTLM_GIS_gi018_sw31)
Dim BL_DTGIS_gi019: BL_DTGIS_gi019=Array(DTLM_GIS_gi019_sw30, DTLM_GIS_gi019_sw31, DTLM_GIS_gi019_sw32)
Dim BL_DTGIS_gi020: BL_DTGIS_gi020=Array(DTLM_GIS_gi020_sw30, DTLM_GIS_gi020_sw31, DTLM_GIS_gi020_sw32)
Dim BL_DTGIS_gi021: BL_DTGIS_gi021=Array(DTLM_GIS_gi021_sw30, DTLM_GIS_gi021_sw31, DTLM_GIS_gi021_sw32)
Dim BL_DTRoom: BL_DTRoom=Array(DTBM_sw30, DTBM_sw31, DTBM_sw32)
' Global arrays
Dim BGDT_Bakemap: BGDT_Bakemap=Array(DTBM_sw30, DTBM_sw31, DTBM_sw32)
Dim BGDT_Lightmap: BGDT_Lightmap=Array(DTLM_GIS_gi008_sw32, DTLM_GIS_gi009_sw30, DTLM_GIS_gi009_sw31, DTLM_GIS_gi009_sw32, DTLM_GIS_gi010_sw30, DTLM_GIS_gi010_sw31, DTLM_GIS_gi010_sw32, DTLM_GIS_gi011_sw30, DTLM_GIS_gi011_sw31, DTLM_GIS_gi011_sw32, DTLM_GIS_gi017_sw30, DTLM_GIS_gi018_sw30, DTLM_GIS_gi018_sw31, DTLM_GIS_gi019_sw30, DTLM_GIS_gi019_sw31, DTLM_GIS_gi019_sw32, DTLM_GIS_gi020_sw30, DTLM_GIS_gi020_sw31, DTLM_GIS_gi020_sw32, DTLM_GIS_gi021_sw30, DTLM_GIS_gi021_sw31, DTLM_GIS_gi021_sw32)
Dim BGDT_All: BGDT_All=Array(DTBM_sw30, DTBM_sw31, DTBM_sw32, DTLM_GIS_gi008_sw32, DTLM_GIS_gi009_sw30, DTLM_GIS_gi009_sw31, DTLM_GIS_gi009_sw32, DTLM_GIS_gi010_sw30, DTLM_GIS_gi010_sw31, DTLM_GIS_gi010_sw32, DTLM_GIS_gi011_sw30, DTLM_GIS_gi011_sw31, DTLM_GIS_gi011_sw32, DTLM_GIS_gi017_sw30, DTLM_GIS_gi018_sw30, DTLM_GIS_gi018_sw31, DTLM_GIS_gi019_sw30, DTLM_GIS_gi019_sw31, DTLM_GIS_gi019_sw32, DTLM_GIS_gi020_sw30, DTLM_GIS_gi020_sw31, DTLM_GIS_gi020_sw32, DTLM_GIS_gi021_sw30, DTLM_GIS_gi021_sw31, DTLM_GIS_gi021_sw32)
' VLM DT Arrays - End

' Thalamus : Exit in a clean and proper way
Sub Table1_exit
  Controller.Pause = False
  Controller.Stop
End Sub


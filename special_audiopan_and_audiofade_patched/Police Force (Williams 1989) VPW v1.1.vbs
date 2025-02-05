'
'                                                       ()
'                                                       /\
'                                                      /  \
'  _____   ____  _      _____ _____ ______     _______/ __ \_______     ______ ____  _____   _____ ______
' |  __ \ / __ \| |    |_   _/ ____|  ____|  ()*.      /  \      .*()  |  ____/ __ \|  __ \ / ____|  ____|
' | |__) | |  | | |      | || |    | |__         *.   ( VV )   .*      | |__ | |  | | |__) | |    | |__
' |  ___/| |  | | |      | || |    |  __|          *.  \__/  .*        |  __|| |  | |  _  /| |    |  __|
' | |    | |__| | |____ _| || |____| |____          /  .**.  \         | |   | |__| | | \ \| |____| |____
' |_|     \____/|______|_____\_____|______|        / .*    *. \        |_|    \____/|_|  \_\\_____|______|
'                                                 /.*        *.\
'                                               ()'            `()
'
' Police Force (Williams 1989)
' https://www.ipdb.org/machine.cgi?id=1841
'
'
'//////////////////////////////////////////////////////////////////////
' THIS TABLE INCLUDES THE IN-GAME OPTIONS MENU
' To open the menu, press both magnasave buttons at the same time in the game.
'//////////////////////////////////////////////////////////////////////
'

Option Explicit
Randomize
SetLocale 1033

On Error Resume Next
ExecuteGlobal GetTextFile("controller.vbs")
If Err Then MsgBox "You need the controller.vbs in order to run this table, available in the vp10 package"
On Error Goto 0


'*******************************************
' Constants and Global Variables
'*******************************************
Const TableVersion = "1.1"       'Table version (also shown in option UI)

Dim UseVPMModSol

Const BallSize = 50     ' Ball size must be 50
Const BallMass = 1      ' Ball mass must be 1
Const tnob = 2        ' Total number of balls
Const lob = 0       ' Total number of locked balls

Dim tablewidth: tablewidth = Table1.width
Dim tableheight: tableheight = Table1.height

'Do not adjust the options here (it won't work). Use the in game options menu.
Dim LightLevel : LightLevel = 50              'LightLevel - Value between 0 and 100 (0=Dark ... 100=Bright)
Dim ColorLUT : ColorLUT = 1           'Color desaturation options (1 to 10)
Dim OutlaneDifficulty : OutlaneDifficulty = 1   'Difficulty : 0 = Easy, 1 = Medium, 2 = Hard
Dim VolumeDial : VolumeDial = 0.8       'Overall Mechanical sound effect volume. Recommended values should be no greater than 1.
Dim BallRollVolume : BallRollVolume = 0.5     'Level of ball rolling volume. Value between 0 and 1
Dim RampRollVolume : RampRollVolume = 0.5     'Level of ramp rolling volume. Value between 0 and 1
Dim ModSol : ModSol = 2             '1 - No PWM, 2 - with PWM (more realistic)
'Dim VRRoomChoice : VRRoomChoice = 1          '1 - Minimal Room, 2 - Ultra Room (only applies when using VR headset)


'************
' Load stuff
'************

Const cGameName = "Polic_L4"

dim op: op = LoadValue(cGameName, "MODSOL") : If op <> "" Then ModSol = CInt(op) Else ModSol = 2
If ModSol = 1 Then
  UseVPMModSol = false
ElseIf ModSol = 2 Then
  UseVPMModSol = true
End If

Const UseSolenoids = 2
Const UseLamps = 1
Const UseGI = 0
Const UseSync = 0
Const HandleMech = 0

' Standard Sounds
Const SSolenoidOn = ""
Const SSolenoidOff = ""
Const SFlipperOn = ""
Const SFlipperOff = ""
Const SCoin = ""

LoadVPM "01550000", "S11.vbs", 3.26

If Controller.Version < "03060000" Then
  ModSol = 1
  SaveValue cGameName, "MODSOL", ModSol
  msgbox "VPinMAME ver 3.6 beta or later is required for PWM flashers. They have been disabled for now. Please restart the table."
End If

Const DebugFlashers = False
Const DebugGI = False

Dim UseVPMDMD, VRRoom, DesktopMode
DesktopMode = Table1.ShowDT
If RenderingMode = 2 Then VRRoom = 1 Else VRRoom = 0
If VRRoom <> 0 Then UseVPMDMD = True Else UseVPMDMD = DesktopMode

' Cabinet mode
If VRRoom = 0 And Not DesktopMode Then
  BM_Rails.visible = False
Else
  If VRRoom = 1 Then
    BM_Rails.visible = False
  Else
    BM_Rails.visible = True
  End If
End If

' Hide desktiop lights if not in desktop
Dim DTLight
If Not DesktopMode Then
  For Each DTLight in DesktopLights: DTLight.Visible = 0 : Next
End If

' VR Mode
Dim VR_Obj
If VRRoom > 0 Then
  Setbackglass
  If VRRoom = 1 Then
    For Each VR_Obj in VRCabinet : VR_Obj.Visible = 1 : Next
    For Each VR_Obj in VRMinimalRoom : VR_Obj.Visible = 1 : Next
    For Each VR_Obj in VRBackglassStuff : VR_Obj.Visible = 1 : Next
  End If
Else
  For Each VR_Obj in VRCabinet : VR_Obj.Visible = 0 : Next
  For Each VR_Obj in VRMinimalRoom : VR_Obj.Visible = 0 : Next
  For Each VR_Obj in VRBackglassStuff : VR_Obj.Visible = 0 : Next
End if

'********************
' Timers
'********************
Sub GameTimer_Timer()
  Cor.Update
  RollingUpdate
  DoDTAnim      'handle drop target animations
  UpdateStandupTargets
  DoSTAnim      'handle stand up target animations
  UpdateDropTargets
  Displaytimer
End Sub

Dim FrameTime, InitFrameTime
InitFrameTime = 0
Sub FrameTimer_Timer() 'The frame timer interval should be -1, so executes at the display frame rate
  FrameTime = GameTime - InitFrameTime
  InitFrameTime = GameTime  'Count frametime
  Options_UpdateDMD
  If AmbientBallShadowOn Then BSUpdate 'update ball shadows
  UpdateBallBrightness
End Sub

'************
' Table init.
'************

Dim PFBall1, PFBall2, gBOT

Sub Table1_Init
  vpmInit me

    With Controller
        .GameName = cGameName
        If Err Then MsgBox "Can't start Game" & cGameName & vbNewLine & Err.Description:Exit Sub
        .SplashInfoLine = "Police Force (Williams 1989)" & vbNewLine & "VPW"
        .Games(cGameName).Settings.Value("rol") = 0 '1= rotated display, 0= normal
        .HandleMechanics = 0
        .HandleKeyboard = 0
        .ShowDMDOnly = 1
        .ShowFrame = 0
        .ShowTitle = 0
    .Hidden = 1
        If Err Then MsgBox Err.Description
    End With
    On Error Goto 0
    Controller.Run

    'Nudging
    vpmNudge.TiltSwitch = swTilt
    vpmNudge.Sensitivity = 2
    vpmNudge.TiltObj = Array(Bumper1, Bumper2, Bumper3, LeftSlingshot, RightSlingShot)

  '************  Trough **************
  Set PFBall1 = sw12.CreateSizedballWithMass(Ballsize/2,Ballmass)
  Set PFBall2 = sw11.CreateSizedballWithMass(Ballsize/2,Ballmass)

  Controller.Switch(12) = 1
  Controller.Switch(11) = 1

  vpmMapLights InsertLamps      ' Map all lamps to the corresponding ROM output using the value of TimerInterval of each light object


  'Main Timer init
    PinMAMETimer.Interval = PinMAMEInterval
    PinMAMETimer.Enabled = 1

  gBOT = Array(PFBall1,PFBall2)


  'Add balls to shadow dictionary
  Dim xx: For Each xx in gBOT
    bsDict.Add xx.ID, bsNone
  Next

  ' Load user options
  Options_Load

  'Initialize slings
  RStep = 0:RightSlingShot.Timerenabled=True
  LStep = 0:LeftSlingShot.Timerenabled=True

  FixCarLoc 'Temporary?

  '************  VLM  **************

  'Room brightness
  SetRoomBrightness LightLevel/100
  SolGI true

  ' Setup new PWM flasher masks if possible
  If UseVPMModSol Then InitPWM()          ' Initialize beta PWM

End Sub

Sub table1_Paused:    Controller.Pause = 1: End Sub
Sub table1_unPaused:  Controller.Pause = 0: End Sub
Sub table1_exit:    Controller.Stop: End Sub



'****************************
' Room Brightness
'****************************

' Update these arrays if you want to change more materials with room light level
Dim RoomBrightnessMtlArray: RoomBrightnessMtlArray = Array("VLM.Bake.Active","VLM.Bake.Blades","VLM.Bake.BumpIn","VLM.Bake.BumpOut","VLM.Bake.Flippers","VLM.Bake.Ramp","VLM.Bake.RampEdges","VLM.Bake.Ramp2","VLM.Bake.Solid")
Dim SavedMtlColorArray:     SavedMtlColorArray     = Array(0,0,0,0,0,0,0,0,0)


Sub SetRoomBrightness(lvl)
  If lvl > 1 Then lvl = 1
  If lvl < 0 Then lvl = 0

  ' Lighting level
  Dim v: v=(lvl * 220 + 35)/255

  Dim i: For i = 0 to UBound(RoomBrightnessMtlArray)
    ModulateMaterialBaseColor RoomBrightnessMtlArray(i), i, v
  Next

  ' VR room stuff
  Dim x
  For Each x in VRCabinet
        x.Color = RGB(255*v, 255*v, 255*v)
    Next
  For Each x in VRMinimalRoom
        x.Color = RGB(255*v, 255*v, 255*v)
    Next
End Sub

SaveMtlColors
Sub SaveMtlColors
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

'UpdateMaterial(string, float wrapLighting, float roughness, float glossyImageLerp, float thickness, float edge, float edgeAlpha, float opacity,
'               OLE_COLOR base, OLE_COLOR glossy, OLE_COLOR clearcoat, VARIANT_BOOL isMetal, VARIANT_BOOL opacityActive,
'               float elasticity, float elasticityFalloff, float friction, float scatterAngle)


'****************************
' Solenoids and Flashers
'****************************
SolCallback(1) = "SolTrough"
SolCallback(2) = "SaucerTL"
SolCallback(3) = "SolBallDiverter"
SolCallback(4) = "dtRightBank"
SolCallback(5) = "SaucerRight"
SolCallback(6) = "dtLeftBank"
SolCallback(7) = "SolKnocker"
SolCallback(8) = "SaucerTR"
'SolCallback(10)= "SolGIBlink"
SolCallback(11)= "SolBGGIBlink" 'BG GI Dimming
SolCallback(23)= "SolGI"
SolCallback(15)= "SolCar"
'SolCallback(16)= "SolRelay"
SolCallback(22)= "SolRelease"

If UseVPMModSol Then
  SolModCallback(9) ="FlashMod109"  'Top Right PF Lamp Under Ramp   f109
  SolModCallback(10)="SolModGIBlink"  'GI Dimming
  SolModCallback(13)="FlashMod113"  'Firing Range Insert (Skill Shot) f113
  SolModCallback(14)="FlashMod114"  'Top Cop Inserts x2         f114 / f114a
  SolModCallback(25)="FlashMod125"  '1c Bottom Left Shark Dome      f125
  SolModCallback(26)="FlashMod126"  '2c Top Left Croc Dome        f126
  SolModCallback(27)="FlashMod127"  '3c Center Rat Dome         f127
  SolModCallback(28)="FlashMod128"  '4c Bottom Right Weasel PF Lamp   f128
  SolModCallback(29)="FlashMod129"  '5c Jackpot Insert/Scope Billboard  f129 / f129a
  SolModCallback(30)="FlashMod130"  '6c Million Flasher Insert      f130
  SolModCallback(31)="FlashMod131"  'Left Dome (Backbox Topper?)    f131
  SolModCallback(32)="FlashMod132"  'Right Dome (Backbox Topper?)   f132
Else
  SolCallback(9) ="Flash109"      'Top Right PF Lamp Under Ramp   f109
  SolCallback(10)="SolGIBlink"    'GI Dimming
  SolCallback(13)="Flash113"      'Firing Range Insert (Skill Shot) f113
  SolCallback(14)="Flash114"      'Top Cop Inserts x2         f114 / f114a
  SolCallback(25)="Flash125"      '1c Bottom Left Shark Dome      f125
  SolCallback(26)="Flash126"      '2c Top Left Croc Dome        f126
  SolCallback(27)="Flash127"      '3c Center Rat Dome         f127
  SolCallback(28)="Flash128"      '4c Bottom Right Weasel PF Lamp   f128
  SolCallback(29)="Flash129"      '5c Jackpot Insert/Scope Billboard  f129 / f129a
  SolCallback(30)="Flash130"      '6c Million Flasher Insert      f130
  SolCallback(31)="Flash131"      'Left Dome (Backbox Topper?)    f131
  SolCallback(32)="Flash132"      'Right Dome (Backbox Topper?)   f132
End IF


'*****************  GI Stuff  ***********************
Const BumperDLMax = 40
Const BumperDLMin = 1
dim gilvl:gilvl = 1

Sub SolGI(enabled)
  If DebugGI Then debug.print " SolGI  "&enabled
  If enabled Then
    GiON
  Else
    GiOFF
  End If
End Sub

Sub GiON
  Dim x
  For each x in GILights:x.State = 1:Next
  If VRRoom > 0 Then
    For each x in VRBGGI:x.visible = 1:Next
  End If
  gilvl = 1
' Bumpers_Outer_BM_Room.blenddisablelighting = BumperDLMax
End Sub

Sub GiOFF
  Dim x
  For each x in GILights:x.State = 0:Next
  If VRRoom > 0 Then
    For each x in VRBGGI:x.visible = 0:Next
  End If
  gilvl = 0
' Bumpers_Outer_BM_Room.blenddisablelighting = BumperDLMin
End Sub

Sub GIBlink(pwm)
  dim x
  For each x in GILights:x.State = pwm:Next
  gilvl = pwm
End Sub

' Drive the ramp and bumpercap fading based on gi01
Sub gi01_Animate
  Dim p
  p = gi01.GetInPlayIntensity / gi01.Intensity
' Bumpers_Outer_BM_Room.blenddisablelighting = (BumperDLMax-BumperDLMin)*p + BumperDLMin
End Sub



'*****************  Knocker  ***********************

Sub SolKnocker(Enabled)
  If enabled Then
    KnockerSolenoid
  End If
End Sub


'*****************  Diverter  ***********************


ica_Diverter.RotateToStart
ica_Diverter.UserValue = 0
Sub SolBallDiverter(Enabled)
  If Enabled Then
    If ica_Diverter.UserValue = 0 Then
      Controller.Switch(36)=0
      ica_Diverter.RotateToEnd
      ica_Diverter.UserValue = 1
      diverter_wall.IsDropped = 0
    Else
      Controller.Switch(36)=1
      ica_Diverter.RotateToStart
      ica_Diverter.UserValue = 0
      diverter_wall.IsDropped = 1
    End If
  End If
End Sub

Sub ica_Diverter_Animate
  Dim a : a = ica_Diverter.CurrentAngle
  Dim BL : For Each BL in BP_DiverterArm : BL.RotZ = a: Next
End Sub


'***************************** CAR CONTROL *********************************
'***** Moves the Car and Locked balls down the ramp and releases them ******

'Fix for car location
Sub FixCarLoc
  dim BL
  For Each BL in BP_Car
    BL.X = 84.2
    BL.Y = 412
    BL.Z = 150
    BL.ObjRotZ = 90
  Next
End Sub

Sub SolCar(Enabled) 'Motor
  CarMoves.Enabled=Enabled
  If Enabled Then
    PlaySoundAtLevelStaticLoop "Motor",VolumeDial,gi15
  Else
    StopSound "Motor"
  End If
  'debug.print "Solcar " & Enabled
End Sub

CarMoves.Enabled = False
Dim CarLocation: CarLocation = 0
Dim CarDirection: CarDirection = 1
Controller.Switch(31)=1

Dim CarBall
Dim CarBallActive: CarBallActive = 0
Const CarDistance = 650
Const CarStep = 4
Const CarTime = 3000 'msec
CarMoves.Interval = CarStep/CarDistance * CarTime
Sub CarMoves_Timer

  CarLocation = CarLocation + (CarStep * CarDirection)

  If CarLocation <= 0 Then  'If car at back of ramp
    CarLocation = 0
    If CarBallActive = 1 Then CarBall1.Z = Carball1.Z + 25
    If CarBallActive = 2 Then CarBall2.Z = Carball2.Z + 25
    CarDirection = 1
    CarKicker.Enabled = True
    CarKicker2.Enabled = False
  Elseif CarLocation > CarDistance Then  'If car at front of ramp
    CarDirection = -1
  End If

  If CarLocation <= 10 Then     ' If car near back of ramp
    Controller.Switch(31)=1

  Elseif CarLocation >= CarDistance - 10 Then 'If CAR is at bottom of ramp, release any locked balls
    Controller.Switch(32)=1

    'Animated ball 1 approach.  Kick ball once at end of ramp
    If CarBallActive > 0 Then
      CarBallActive = CarBallActive - 1
      CarKicker.Kick 180, 1
    End If

    If CarBallActive > 0 Then
      CarBallActive = CarBallActive - 1
      CarKicker2.Kick 180, 1
    End If

'Animated ball 2 approach
'   CarKicker2.Kick 180, 1
'   If CarBallActive > 0 Then CarBallActive = CarBallActive - 1

'Basic 2 kicker approach
  ' Release Ball 1
'   if CarBallActive > 0 Then
'     CarKicker2.CreateBall
'     Carkicker2.kick 180, 1
'     CarBallActive = CarBallActive - 1
'   End If
  ' Release Ball 2
'   if CarBallActive > 0 Then
'     CarKicker3.CreateBall
'     Carkicker3.kick 180, 1
'     CarBallActive = CarBallActive - 1
'   End If
  Else              ' If car in the middle of the ramp
    Controller.Switch(31) = 0
    Controller.Switch(32) = 0
  End If

  ' Move car
  'debug.print "Moving " & CarLocation & ":" & Controller.Switch(31) & ":" & Controller.Switch(32) & ":" & CarDirection
  'vna_car.TransZ = -CarLocation
  dim BL: For Each BL in BP_Car: BL.TransX = -CarLocation: Next
  L47a.y = CarLocation +390: L47b.y = CarLocation +520
  L48a.y = CarLocation +410: L48b.y = CarLocation +540

  ' Move any captured balls
  If CarBallActive >= 1 Then
    CarBall1.y = CarBall1Y + CarLocation
  End If

  If CarBallActive = 2 Then
    CarBall2.y = CarBall2Y + CarLocation
  End If

End Sub

dim carball1, carball2, carball1y, carball2y
Sub CarKicker_Hit
  CarBallActive = CarBallActive + 1
  Set CarBall1 = Activeball
  CarBall1Y = CarBall1.y
  CarKicker2.Enabled = True
End Sub

Sub CarKicker2_Hit
  CarBallActive = CarBallActive + 1
  Set CarBall2 = Activeball
  CarBall2Y = CarBall2.y
  CarKicker2.Enabled = False
End Sub


' Car wall flashers
Sub L47_animate
  dim p
  p = L47.GetInPlayIntensity / L47.Intensity
  L47a.IntensityScale = p
  L47b.IntensityScale = p
End Sub

Sub L48_animate
  dim p
  p = L48.GetInPlayIntensity / L48.Intensity
  L48a.IntensityScale = p
  L48b.IntensityScale = p
End Sub



'*************************************************************
'PWM Flasher Stuff
'*************************************************************


'1001 to 1032 for solenoids, 1201 to 1025 for GIs (WPC only), 1301+ for lamps (not yet implemented)
Const VPM_MODOUT_DEFAULT              =   0 ' Uses default driver modulated solenoid implementation
Const VPM_MODOUT_PWM_RATIO            =   1 ' pulse ratio over the last integration period, allow (approximated) device emulation by the calling app
Const VPM_MODOUT_BULB_44_6_3V_AC      = 100 ' Incandescent #44/555 Bulb connected to 6.3V, commonly used for GI
Const VPM_MODOUT_BULB_47_6_3V_AC      = 101 ' Incandescent #47 Bulb connected to 6.3V, commonly used for (darker) GI with less heat
Const VPM_MODOUT_BULB_86_6_3V_AC      = 102 ' Incandescent #86 Bulb connected to 6.3V, seldom use: TZ, CFTBL,...
Const VPM_MODOUT_BULB_44_18V_DC_WPC   = 201 ' Incandescent #44/555 Bulb connected to 18V, commonly used for lamp matrix with short strobing
Const VPM_MODOUT_BULB_44_18V_DC_GTS3  = 202 ' Incandescent #44/555 Bulb connected to 18V, commonly used for lamp matrix with short strobing
Const VPM_MODOUT_BULB_44_18V_DC_S11   = 203 ' Incandescent #44/555 Bulb connected to 18V, commonly used for lamp matrix with short strobing
Const VPM_MODOUT_BULB_89_20V_DC_WPC   = 301 ' Incandescent #89/906 Bulb connected to 12V, commonly used for flashers
Const VPM_MODOUT_BULB_89_20V_DC_GTS3  = 302 ' Incandescent #89/906 Bulb connected to 12V, commonly used for flashers
Const VPM_MODOUT_BULB_89_32V_DC_S11   = 303 ' Incandescent #89/906 Bulb connected to 32V, used for flashers on S11 with output strobing
Const VPM_MODOUT_LED                  = 400 ' LED PWM (in fact mostly human eye reaction, since LED are nearly instantaneous)

Sub InitPWM() ' called from Table_Init
  dim BulbType : BulbType = VPM_MODOUT_BULB_89_20V_DC_WPC
' dim BulbType : BulbType = VPM_MODOUT_LED
  If Controller.Version < "03060000" Then
    msgbox "VPinMAME ver 3.6 beta or later is required"
  Else
    Controller.SolMask(1009) = BulbType 'Top Right PF Lamp Under Ramp   f109
    Controller.SolMask(1010) = VPM_MODOUT_PWM_RATIO 'GI flashing
    Controller.SolMask(1013) = BulbType 'Firing Range Insert (Skill Shot) f113
    Controller.SolMask(1014) = BulbType 'Top Cop Inserts x2         f114 / f114a
    Controller.SolMask(1025) = BulbType '1c Bottom Left Shark Dome      f125
    Controller.SolMask(1026) = BulbType '2c Top Left Croc Dome        f126
    Controller.SolMask(1027) = BulbType '3c Center Rat Dome         f127
    Controller.SolMask(1028) = BulbType '4c Bottom Right Weasel PF Lamp   f128
    Controller.SolMask(1029) = BulbType '5c Jackpot Insert/Scope Billboard  f129 / f129a
    Controller.SolMask(1030) = BulbType '6c Million Flasher Insert      f130
    Controller.SolMask(1031) = BulbType 'Left Dome (Backbox Topper?)    f131
    Controller.SolMask(1032) = BulbType 'Right Dome (Backbox Topper?)   f132
  End If
End Sub

Sub SolModGIBlink(pwm)
  If DebugGI Then Debug.print "SolModGIBlink level=" & pwm
  Dim p : p = (1 - pwm / 255.0)'^2
  GIBlink(p)
  'debug.print pwm
End Sub

Sub FlashMod109(pwm)
  If DebugFlashers Then Debug.print "FlashMod109 level=" & pwm
  Dim p : p = pwm / 255.0
  f109.State = p
End Sub

Sub FlashMod113(pwm)
  If DebugFlashers Then Debug.print "FlashMod113 level=" & pwm
  Dim p : p = pwm / 255.0
  f113.State = p
End Sub

Sub FlashMod114(pwm)
  If DebugFlashers Then Debug.print "FlashMod114 level=" & pwm
  Dim p : p = pwm / 255.0
  f114.State = p
  f114a.State = p
End Sub

Sub FlashMod125(pwm)
  If DebugFlashers Then Debug.print "FlashMod125 level=" & pwm
  Dim p : p = pwm / 255.0
  f125.State = p
End Sub

Sub FlashMod126(pwm)
  If DebugFlashers Then Debug.print "FlashMod126 level=" & pwm
  Dim p : p = pwm / 255.0
  f126.State = p
End Sub

Sub FlashMod127(pwm)
  If DebugFlashers Then Debug.print "FlashMod127 level=" & pwm
  Dim p : p = pwm / 255.0
  f127.State = p
End Sub

Sub FlashMod128(pwm)
  If DebugFlashers Then Debug.print "FlashMod128 level=" & pwm
  Dim p : p = pwm / 255.0
  f128.State = p
End Sub

Sub FlashMod129(pwm)
  If DebugFlashers Then Debug.print "FlashMod129 level=" & pwm
  Dim p : p = pwm / 255.0
  f129.State = p
  f129a.State = p
End Sub

Sub FlashMod130(pwm)
  If DebugFlashers Then Debug.print "FlashMod130 level=" & pwm
  Dim p : p = pwm / 255.0
  f130.State = p
End Sub

Sub FlashMod131(pwm)
  If DebugFlashers Then Debug.print "FlashMod131 level=" & pwm
  Dim p : p = pwm / 255.0
  'f131.State = p
End Sub

Sub FlashMod132(pwm)
  If DebugFlashers Then Debug.print "FlashMod132 level=" & pwm
  Dim p : p = pwm / 255.0
  'f132.State = p
End Sub

'*********************************
' Non-Modulated Solenoids
'*********************************

Sub SolBGGIBlink(lvl)
  If VRRoom > 0 Then
    Dim p : p = lvl
    If p = 0 then f111.state = 1 Else f111.state = 0
  End If
End Sub

Sub SolGIBlink(lvl)
  If DebugGI Then Debug.print "SolGIBlink level=" & lvl
  Dim p : p = (1 - lvl)'^2
  GIBlink(p)
  'debug.print p
End Sub

Sub Flash109(lvl)
  If DebugFlashers Then Debug.print "Flash109 level=" & lvl
  f109.State = lvl
End Sub

Sub Flash113(lvl)
  If DebugFlashers Then Debug.print "Flash113 level=" & lvl
  f113.State = lvl
End Sub

Sub Flash114(lvl)
  If DebugFlashers Then Debug.print "Flash114 level=" & lvl
  f114.State = lvl
  f114a.State = lvl
End Sub

Sub Flash125(lvl)
  If DebugFlashers Then Debug.print "Flash125 level=" & lvl
  f125.State = lvl
End Sub

Sub Flash126(lvl)
  If DebugFlashers Then Debug.print "Flash126 level=" & lvl
  f126.State = lvl
End Sub

Sub Flash127(lvl)
  If DebugFlashers Then Debug.print "Flash127 level=" & lvl
  f127.State = lvl
End Sub

Sub Flash128(lvl)
  If DebugFlashers Then Debug.print "Flash128 level=" & lvl
  f128.State = lvl
End Sub

Sub Flash129(lvl)
  If DebugFlashers Then Debug.print "Flash129 level=" & lvl
  f129.State = lvl
  f129a.State = lvl
End Sub

Sub Flash130(lvl)
  If DebugFlashers Then Debug.print "Flash130 level=" & lvl
  f130.State = lvl
End Sub

Sub Flash131(lvl)
  If DebugFlashers Then Debug.print "Flash131 level=" & lvl
  'f131.State = lvl
End Sub

Sub Flash132(lvl)
  If DebugFlashers Then Debug.print "Flash132 level=" & lvl
  'f132.State = lvl
End Sub


'*********************************
' Keys Section
'*********************************

Sub Table1_KeyDown(ByVal Keycode)

  If bInOptions Then
    Options_KeyDown keycode
    Exit Sub
  End If
    If keycode = LeftMagnaSave Then
    If bOptionsMagna Then Options_Open() Else bOptionsMagna = True
    ElseIf keycode = RightMagnaSave Then
    If bOptionsMagna Then Options_Open() Else bOptionsMagna = True
  End If

  If keycode = LeftTiltKey Then Nudge 90, 1:SoundNudgeLeft()
  If keycode = RightTiltKey Then Nudge 270, 1:SoundNudgeRight()
  If keycode = CenterTiltKey Then Nudge 0, 1:SoundNudgeCenter()

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

  if keycode = StartGameKey then soundStartButton()

  If KeyCode = PlungerKey Then Plunger.Pullback:SoundPlungerPull():TimerVRPlunger.Enabled = True :TimerVRPlunger2.Enabled = False
  If keycode = LeftFlipperKey Then FlipperActivate LeftFlipper, LFPress:Pincab_FB_Left.X = Pincab_FB_Left.X + 10
  If keycode = RightFlipperKey Then FlipperActivate RightFlipper, RFPress:Pincab_FB_Right.X = Pincab_FB_Right.X - 10

    If vpmKeyDown(keycode) Then Exit Sub

End Sub

Sub Table1_KeyUp(ByVal Keycode)

    If keycode = LeftMagnaSave And Not bInOptions Then bOptionsMagna = False
    If keycode = RightMagnaSave And Not bInOptions Then bOptionsMagna = False

  If KeyCode = PlungerKey Then
    Plunger.Fire
    TimerVRPlunger.enabled = False
    TimerVRPlunger2.enabled = True
    If BallInPlungerLane = 1 Then
      SoundPlungerReleaseBall()     'Plunger release sound when there is a ball in shooter lane
    Else
      SoundPlungerReleaseNoBall()     'Plunger release sound when there is no ball in shooter lane
    End If
  End If

  If keycode = LeftFlipperKey Then FlipperDeActivate LeftFlipper, LFPress:Pincab_FB_Left.X = Pincab_FB_Left.X - 10
  If keycode = RightFlipperKey Then FlipperDeActivate RightFlipper, RFPress:Pincab_FB_Right.X = Pincab_FB_Right.X + 10

    If vpmKeyUp(keycode) Then Exit Sub

End Sub

'*********************************
' SWITCHES
'*********************************
Dim BallInPlungerLane

'Bumpers
Sub Bumper1_Hit(): vpmTimer.Pulsesw 60 : RandomSoundBumperTop Bumper1: End Sub
Sub Bumper2_Hit(): vpmTimer.Pulsesw 61 : RandomSoundBumperMiddle Bumper2: End Sub
Sub Bumper3_Hit(): vpmTimer.Pulsesw 62 : RandomSoundBumperBottom Bumper3: End Sub

'Rollovers
Sub sw14_Hit() : Controller.Switch(14)=1 : BallInPlungerLane = 1 : End Sub
Sub sw14_Unhit() :  Controller.Switch(14)=0 : BallInPlungerLane = 0 : End Sub

Sub sw49_Hit() : Controller.Switch(49)=1 : End Sub  'Inlanes/Outlanes
Sub sw49_UnHit() : Controller.Switch(49)=0: End Sub
Sub sw50_Hit() : Controller.Switch(50)=1 : End Sub
Sub sw50_UnHit() : Controller.Switch(50)=0: End Sub
Sub sw51_Hit() : Controller.Switch(51)=1 : End Sub
Sub sw51_UnHit() : Controller.Switch(51)=0: End Sub
Sub sw52_Hit() : Controller.Switch(52)=1 : End Sub
Sub sw52_UnHit() : Controller.Switch(52)=0: End Sub

Sub sw54_Hit() : Controller.Switch(54)=1 : End Sub 'Toplanes
Sub sw54_UnHit() : Controller.Switch(54)=0: End Sub
Sub sw55_Hit() : Controller.Switch(55)=1 : End Sub
Sub sw55_UnHit() : Controller.Switch(55)=0: End Sub
Sub sw56_Hit() : Controller.Switch(56)=1 : End Sub
Sub sw56_UnHit() : Controller.Switch(56)=0: End Sub

'Ramp Gates
Sub sw34_Hit(): vpmTimer.Pulsesw 34 : End Sub
Sub sw45_Hit(): vpmTimer.Pulsesw 45 : End Sub
Sub sw46_Hit(): vpmTimer.Pulsesw 46 : End Sub

'Ramp Triggers
Sub sw37_Hit() : vpmTimer.Pulsesw 37 : End Sub
Sub sw38_Hit() : vpmTimer.Pulsesw 38 : End Sub

'Spinners
Sub sw41_Spin(): vpmTimer.Pulsesw 41 : End Sub

Sub sw41_Animate
  Dim spinangle:spinangle = -sw41.currentangle+90
  Dim BL : For Each BL in BP_Spinner : BL.RotX = spinangle: Next
End Sub

'********************************************
' ZTAR: Targets
'********************************************

Sub sw17_Hit
  STHit 17
End Sub

Sub sw17o_Hit
  TargetBouncer ActiveBall, 1
End Sub

Sub sw18_Hit
  STHit 18
End Sub

Sub sw18o_Hit
  TargetBouncer ActiveBall, 1
End Sub

Sub sw19_Hit
  STHit 19
End Sub

Sub sw19o_Hit
  TargetBouncer ActiveBall, 1
End Sub

Sub sw35_Hit
  STHit 35
End Sub

Sub sw35o_Hit
  TargetBouncer ActiveBall, 1
End Sub

'********************************************
'  Drop Target Controls
'********************************************

' Drop targets
Sub sw20_Hit
  DTHit 20
End Sub

Sub sw21_Hit
  DTHit 21
End Sub

Sub sw22_Hit
  DTHit 22
End Sub

Sub sw25_Hit
  DTHit 25
End Sub

Sub sw26_Hit
  DTHit 26
End Sub

Sub sw27_Hit
  DTHit 27
End Sub

Sub dtRightBank(enabled)
' Dim xx
  If enabled Then
    RandomSoundDropTargetReset BM_TargetR2
    DTRaise 20
    DTRaise 21
    DTRaise 22
'   For Each xx In ShadowDT
'     xx.visible = True
'   Next
  End If
End Sub

Sub dtLeftBank(enabled)
' Dim xx
  If enabled Then
    RandomSoundDropTargetReset BM_TargetL2
    DTRaise 25
    DTRaise 26
    DTRaise 27
'   For Each xx In ShadowDT
'     xx.visible = True
'   Next
  End If
End Sub


'******************************************************
'* BUMPER ANIMATIONS
'******************************************************

Sub Bumper1_Animate
  Dim z, BL
  z = Bumper1.CurrentRingOffset
  For Each BL in BP_Bumper1Ring : BL.transz = z: Next
End Sub

Sub Bumper2_Animate
  Dim z, BL
  z = Bumper2.CurrentRingOffset
  For Each BL in BP_Bumper2Ring : BL.transz = z: Next
End Sub

Sub Bumper3_Animate
  Dim z, BL
  z = Bumper3.CurrentRingOffset
  For Each BL in BP_Bumper3Ring : BL.transz = z: Next
End Sub

'******************************************************
' SWITCH ANIMATIONS
'******************************************************

Sub sw14_Animate
  Dim z : z = sw14.CurrentAnimOffset
  Dim BL : For Each BL in BP_sw14Rollover : BL.transz = z: Next
End Sub

Sub sw49_Animate
  Dim z : z = sw49.CurrentAnimOffset
  Dim BL : For Each BL in BP_sw49Rollover : BL.transz = z: Next
End Sub

Sub sw50_Animate
  Dim z : z = sw50.CurrentAnimOffset
  Dim BL : For Each BL in BP_sw50Rollover : BL.transz = z: Next
End Sub

Sub sw51_Animate
  Dim z : z = sw51.CurrentAnimOffset
  Dim BL : For Each BL in BP_sw51Rollover : BL.transz = z: Next
End Sub

Sub sw52_Animate
  Dim z : z = sw52.CurrentAnimOffset
  Dim BL : For Each BL in BP_sw52Rollover : BL.transz = z: Next
End Sub

Sub sw54_Animate
  Dim z : z = sw54.CurrentAnimOffset
  Dim BL : For Each BL in BP_sw54Rollover : BL.transz = z: Next
End Sub

Sub sw55_Animate
  Dim z : z = sw55.CurrentAnimOffset
  Dim BL : For Each BL in BP_sw55Rollover : BL.transz = z: Next
End Sub

Sub sw56_Animate
  Dim z : z = sw56.CurrentAnimOffset
  Dim BL : For Each BL in BP_sw56Rollover : BL.transz = z: Next
End Sub



'******************************************************
' GATE ANIMATIONS
'******************************************************

Sub Gate2_Animate
  Dim a : a = Gate2.CurrentAngle
  Dim BL : For Each BL in BP_Gate05Wire : BL.rotx = a: Next
End Sub

'Sub Gate3_Animate  'commented out on purpose. The gate wire was cutting through the ramp. Gate is nearly impossible to see anyway. -apophis
' Dim a : a = Gate3.CurrentAngle
' Dim BL : For Each BL in BP_Gate003_Wire : BL.rotx = a: Next
'End Sub

Sub Gate4_Animate
  Dim a : a = Gate4.CurrentAngle
  Dim BL : For Each BL in BP_Gate04Wire : BL.rotx = a: Next
End Sub

Sub Gate6_Animate
  Dim a : a = Gate4.CurrentAngle
  Dim BL : For Each BL in BP_Gate02Wire : BL.rotx = a: Next
End Sub

Sub sw34_Animate
  Dim a : a = sw34.CurrentAngle
  Dim BL : For Each BL in BP_GateSw44Wire : BL.rotx = a: Next
End Sub

Sub sw45_Animate
  Dim a : a = sw45.CurrentAngle
  Dim BL : For Each BL in BP_GateSw45Wire : BL.rotx = a: Next
End Sub

Sub sw46_Animate
  Dim a : a = sw46.CurrentAngle
  Dim BL : For Each BL in BP_GateSw46Wire : BL.rotx = a: Next
End Sub





'******************************************************
'* TROUGH *********************************************
'******************************************************

Sub sw12_Hit   : Controller.Switch(12) = 1 : UpdateTrough : End Sub
Sub sw12_UnHit : Controller.Switch(12) = 0 : UpdateTrough : End Sub
Sub sw11_Hit   : Controller.Switch(11) = 1 : UpdateTrough : End Sub
Sub sw11_UnHit : Controller.Switch(11) = 0 : UpdateTrough : End Sub
Sub sw10_Hit   : Controller.Switch(10) = 1 : UpdateTrough : End Sub
Sub sw10_UnHit : Controller.Switch(10) = 0 : UpdateTrough : End Sub

Sub UpdateTrough
  UpdateTroughTimer.Interval = 100
  UpdateTroughTimer.Enabled = 1
End Sub

Sub UpdateTroughTimer_Timer
  If sw11.BallCntOver = 0 Then sw12.kick 57, 20
  Me.Enabled = 0
End Sub


'******************************************************
' DRAIN & RELEASE
'******************************************************

Sub SolTrough(enabled)
  If enabled Then
    sw10.kick 57, 20
    RandomSoundDrain sw10
  End If
End Sub

Sub SolRelease(enabled)
  If enabled Then
    sw11.kick 90, 10
    Controller.Switch(14) = 0
    RandomSoundBallRelease sw11
  End If
End Sub

'***************************
'     VUKs
'***************************
Dim KickerBall15, KickerBall23, KickerBall24

Sub KickBall(kball, kangle, kvel, kvelz, kzlift)
  dim rangle
  rangle = PI * (kangle - 90) / 180
  kball.z = kball.z + kzlift
  kball.velz = kvelz
  kball.velx = cos(rangle)*kvel
  kball.vely = sin(rangle)*kvel
End Sub

'Saucer Top Right
Sub sw15_Hit()
  set KickerBall15 = activeball
  Controller.Switch(15) = 1
  SoundSaucerLock
End Sub

Sub SaucerTR(Enable)
  If Enable then
    If Controller.Switch(15) <> 0 Then
      KickBall KickerBall15, 280, 15.5, 5, 15
      SoundSaucerKick 1, sw15
      Controller.Switch(15) = 0
    End If
  End If
End Sub

'Saucer Right
Sub sw23_hit()
  set KickerBall23 = activeball
  Controller.Switch(23) = True
  SoundSaucerLock
End sub

Sub SaucerRight(enabled)
  If enabled then
    If Controller.Switch(23) <> 0 Then
      KickBall KickerBall23, 279+2*rnd, 12.5+2*rnd, 5, 30
      SoundSaucerKick 1, sw23
      Controller.Switch(23) = 0
    End If
  End If
End Sub

'Saucer Top Left
Sub sw24_hit()
  set KickerBall24 = activeball
  Controller.Switch(24) = True
  SoundSaucerLock
End sub

Sub SaucerTL(enabled)
  If enabled then
    If Controller.Switch(24) <> 0 Then
      KickBall KickerBall24, 65, 10, 5, 35
      SoundSaucerKick 1, sw24
      Controller.Switch(24) = 0
    End If
  End If
End Sub


'**********Sling Shot Animations
' Rstep and Lstep  are the variables that increment the animation
'****************
Dim LStep, RStep

Sub RightSlingShot_Slingshot
  RS.VelocityCorrect(Activeball)
  vpmtimer.pulsesw(63)
  RandomSoundSlingshotRight Remk
  RStep = 0 : RightSlingShot_Timer ' Initialize Step to 0
  RightSlingShot.TimerInterval = 17
  RightSlingShot.TimerEnabled = 1
End Sub

Sub RightSlingShot_Timer
  Dim BL
  Dim v1, v2, v3, v4, x
  v1 = False: v2 = False: v3 = False: v4 = True: x = -30
    Select Case RStep
        Case 2:v1 = False: v2 = False: v3 = True:  v4 = False: x = -20
        Case 3:v1 = False: v2 = True:  v3 = False: v4 = False: x = -10
        Case 4:v1 = True:  v2 = False: v3 = False: v4 = False: x = 0: RightSlingShot.TimerEnabled = 0
    End Select

  For Each BL in BP_Rightsling1 : BL.Visible = v1: Next
  For Each BL in BP_Rightsling2 : BL.Visible = v2: Next
  For Each BL in BP_Rightsling3 : BL.Visible = v3: Next
  For Each BL in BP_Rightsling4 : BL.Visible = v4: Next
' For Each BL in BP_SlingArmR : BL.transx = x: Next

    RStep = RStep + 1
End Sub

Sub LeftSlingShot_Slingshot
  LS.VelocityCorrect(Activeball)
  vpmtimer.pulsesw(64)
  RandomSoundSlingshotLeft Lemk
  LStep = 0 : LeftSlingShot_Timer ' Initialize Step to 0
  LeftSlingShot.TimerInterval = 17
  LeftSlingShot.TimerEnabled = 1
End Sub

Sub LeftSlingShot_Timer
  Dim BL
  Dim v1, v2, v3, v4, x
  v1 = False: v2 = False: v3 = False: v4 = True: x = -30
    Select Case LStep
        Case 2:v1 = False: v2 = False: v3 = True:  v4 = False: x = -20
        Case 3:v1 = False: v2 = True:  v3 = False: v4 = False: x = -10
        Case 4:v1 = True:  v2 = False: v3 = False: v4 = False: x = 0: LeftSlingShot.TimerEnabled = 0
    End Select

  For Each BL in BP_Leftsling1 : BL.Visible = v1: Next
  For Each BL in BP_Leftsling2 : BL.Visible = v2: Next
  For Each BL in BP_Leftsling3 : BL.Visible = v3: Next
  For Each BL in BP_Leftsling4 : BL.Visible = v4: Next
' For Each BL in BP_SlingArmL : BL.transx = x: Next

    LStep = LStep + 1
End Sub


'******************************************************
'           FLIPPERS
'******************************************************
SolCallback(sLRFlipper)     = "SolRFlipper"
SolCallback(sLLFlipper)     = "SolLFlipper"

Const ReflipAngle = 20

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

Sub LeftFlipper_Collide(parm)
  CheckLiveCatch ActiveBall, LeftFlipper, LFCount, parm
  LeftFlipperCollide parm
End Sub

Sub RightFlipper_Collide(parm)
  CheckLiveCatch ActiveBall, RightFlipper, RFCount, parm
  RightFlipperCollide parm
End Sub

Sub LeftFlipper_Animate
  Dim lfa : lfa = LeftFlipper.CurrentAngle
  FlipperLSh.RotZ = lfa
  Dim BL : For Each BL in BP_FlipperL : BL.RotZ = lfa: Next
End Sub

Sub RightFlipper_Animate
  Dim rfa : rfa = RightFlipper.CurrentAngle
    FlipperRSh.RotZ = rfa
  Dim BL : For Each BL in BP_FlipperR : BL.RotZ = rfa: Next
End Sub


'******************************************************
' ZPHY:  GNEREAL ADVICE ON PHYSICS
'******************************************************
'
' It's advised that flipper corrections, dampeners, and general physics settings should all be updated per these
' examples as all of these improvements work together to provide a realistic physics simulation.
'
' Tutorial videos provided by Bord
' Adding nFozzy roth physics : pt1 rubber dampeners         https://youtu.be/AXX3aen06FM
' Adding nFozzy roth physics : pt2 flipper physics          https://youtu.be/VSBFuK2RCPE
' Adding nFozzy roth physics : pt3 other elements           https://youtu.be/JN8HEJapCvs
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
' | Force         | 9.5-10.5 |
' | Hit Threshold | 1.6-2    |
' | Scatter Angle | 2        |
'
' Slingshots
' | Hit Threshold      | 2    |
' | Slingshot Force    | 4-5  |
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
' 3. an object or point to tell the script where the tip of the flipper is at rest (EndPointLp, EndPointRp)
' 4. and, special scripting
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
' | EOS Torque         | 0.3            | 0.3                   | 0.275                  | 0.275              |
' | EOS Torque Angle   | 4              | 4                     | 6                      | 6                  |
'

'******************************************************
' Flippers Polarity (Select appropriate sub based on era)
'******************************************************

Dim LF
Set LF = New FlipperPolarity
Dim RF
Set RF = New FlipperPolarity

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
'   x.AddPt "Polarity", 0, 0, 0
'   x.AddPt "Polarity", 1, 0.05, - 2.7
'   x.AddPt "Polarity", 2, 0.33, - 2.7
'   x.AddPt "Polarity", 3, 0.37, - 2.7
'   x.AddPt "Polarity", 4, 0.41, - 2.7
'   x.AddPt "Polarity", 5, 0.45, - 2.7
'   x.AddPt "Polarity", 6, 0.576, - 2.7
'   x.AddPt "Polarity", 7, 0.66, - 1.8
'   x.AddPt "Polarity", 8, 0.743, - 0.5
'   x.AddPt "Polarity", 9, 0.81, - 0.5
'   x.AddPt "Polarity", 10, 0.88, 0
'
'   x.AddPt "Velocity", 0, 0, 1
'   x.AddPt "Velocity", 1, 0.16, 1.06
'   x.AddPt "Velocity", 2, 0.41, 1.05
'   x.AddPt "Velocity", 3, 0.53, 1 '0.982
'   x.AddPt "Velocity", 4, 0.702, 0.968
'   x.AddPt "Velocity", 5, 0.95,  0.968
'   x.AddPt "Velocity", 6, 1.03, 0.945
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
'   x.AddPt "Polarity", 2, 0.33, - 3.7
'   x.AddPt "Polarity", 3, 0.37, - 3.7
'   x.AddPt "Polarity", 4, 0.41, - 3.7
'   x.AddPt "Polarity", 5, 0.45, - 3.7
'   x.AddPt "Polarity", 6, 0.576,- 3.7
'   x.AddPt "Polarity", 7, 0.66, - 2.3
'   x.AddPt "Polarity", 8, 0.743, - 1.5
'   x.AddPt "Polarity", 9, 0.81, - 1
'   x.AddPt "Polarity", 10, 0.88, 0
'
'   x.AddPt "Velocity", 0, 0, 1
'   x.AddPt "Velocity", 1, 0.16, 1.06
'   x.AddPt "Velocity", 2, 0.41, 1.05
'   x.AddPt "Velocity", 3, 0.53, 1 '0.982
'   x.AddPt "Velocity", 4, 0.702, 0.968
'   x.AddPt "Velocity", 5, 0.95,  0.968
'   x.AddPt "Velocity", 6, 1.03, 0.945
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
'   x.AddPt "Polarity", 2, 0.4, - 5
'   x.AddPt "Polarity", 3, 0.6, - 4.5
'   x.AddPt "Polarity", 4, 0.65, - 4.0
'   x.AddPt "Polarity", 5, 0.7, - 3.5
'   x.AddPt "Polarity", 6, 0.75, - 3.0
'   x.AddPt "Polarity", 7, 0.8, - 2.5
'   x.AddPt "Polarity", 8, 0.85, - 2.0
'   x.AddPt "Polarity", 9, 0.9, - 1.5
'   x.AddPt "Polarity", 10, 0.95, - 1.0
'   x.AddPt "Polarity", 11, 1, - 0.5
'   x.AddPt "Polarity", 12, 1.1, 0
'   x.AddPt "Polarity", 13, 1.3, 0
'
'   x.AddPt "Velocity", 0, 0, 1
'   x.AddPt "Velocity", 1, 0.16, 1.06
'   x.AddPt "Velocity", 2, 0.41, 1.05
'   x.AddPt "Velocity", 3, 0.53, 1 '0.982
'   x.AddPt "Velocity", 4, 0.702, 0.968
'   x.AddPt "Velocity", 5, 0.95,  0.968
'   x.AddPt "Velocity", 6, 1.03,  0.945
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
  a = Array(LF, RF)
  For Each x In a
    x.AddPt "Ycoef", 0, RightFlipper.Y-65, 1 'disabled
    x.AddPt "Ycoef", 1, RightFlipper.Y-11, 1
    x.enabled = True
    x.TimeDelay = 60
    x.DebugOn=False ' prints some info in debugger

    x.AddPt "Polarity", 0, 0, 0
    x.AddPt "Polarity", 1, 0.05, -5.5
    x.AddPt "Polarity", 2, 0.4, -5.5
    x.AddPt "Polarity", 3, 0.6, -5.0
    x.AddPt "Polarity", 4, 0.65, -4.5
    x.AddPt "Polarity", 5, 0.7, -4.0
    x.AddPt "Polarity", 6, 0.75, -3.5
    x.AddPt "Polarity", 7, 0.8, -3.0
    x.AddPt "Polarity", 8, 0.85, -2.5
    x.AddPt "Polarity", 9, 0.9,-2.0
    x.AddPt "Polarity", 10, 0.95, -1.5
    x.AddPt "Polarity", 11, 1, -1.0
    x.AddPt "Polarity", 12, 1.05, -0.5
    x.AddPt "Polarity", 13, 1.1, 0
    x.AddPt "Polarity", 14, 1.3, 0

    x.AddPt "Velocity", 0, 0,    1
    x.AddPt "Velocity", 1, 0.160, 1.06
    x.AddPt "Velocity", 2, 0.410, 1.05
    x.AddPt "Velocity", 3, 0.530, 1'0.982
    x.AddPt "Velocity", 4, 0.702, 0.968
    x.AddPt "Velocity", 5, 0.95,  0.968
    x.AddPt "Velocity", 6, 1.03,  0.945
  Next

  ' SetObjects arguments: 1: name of object 2: flipper object: 3: Trigger object around flipper
  LF.SetObjects "LF", LeftFlipper, TriggerLF
  RF.SetObjects "RF", RightFlipper, TriggerRF
End Sub

'' Flipper trigger hit subs
'Sub TriggerLF_Hit()
' LF.Addball activeball
'End Sub
'Sub TriggerLF_UnHit()
' LF.PolarityCorrect activeball
'End Sub
'Sub TriggerRF_Hit()
' RF.Addball activeball
'End Sub
'Sub TriggerRF_UnHit()
' RF.PolarityCorrect activeball
'End Sub

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
  '   Dim BOT
  '   BOT = GetBalls

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
'Const EOSTnew = 1 'EM's to late 80's
Const EOSTnew = 0.8 '90's and later
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

' Note, cor.update must be called in a 10 ms timer. The example table uses the GameTimer for this purpose, but sometimes a dedicated timer call RDampen is used.
'
'Sub RDampen_Timer
' Cor.Update
'End Sub

'******************************************************
'****  END PHYSICS DAMPENERS
'******************************************************



'******************************************************
'   ZBOU: VPW TargetBouncer for targets and posts by Iaakki, Wrd1972, Apophis
'******************************************************

Const TargetBouncerEnabled = 1    '0 = normal standup targets, 1 = bouncy targets
Const TargetBouncerFactor = 0.7  'Level of bounces. Recommmended value of 0.7

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
'   ZRDT:  DROP TARGETS by Rothbauerw
'******************************************************
' This solution improves the physics for drop targets to create more realistic behavior. It allows the ball
' to move through the target enabling the ability to score more than one target with a well placed shot.
' It also handles full drop target animation, including deflection on hit and a slight lift when the drop
' targets raise, switch handling, bricking, and popping the ball up if it's over the drop target when it raises.
'
'Add a Timer named DTAnim to editor to handle drop & standup target animations, or run them off an always-on 10ms timer (GameTimer)
'DTAnim.interval = 10
'DTAnim.enabled = True

'Sub DTAnim_Timer
' DoDTAnim
' DoSTAnim
'End Sub

' For each drop target, we'll use two wall objects for physics calculations and one primitive for visuals and
' animation. We will not use target objects.  Place your drop target primitive the same as you would a VP drop target.
' The primitive should have it's pivot point centered on the x and y axis and at or just below the playfield
' level on the z axis. Orientation needs to be set using Rotz and bending deflection using Rotx. You'll find a hooded
' target mesh in this table's example. It uses the same texture map as the VP drop targets.

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
Dim DT20, DT21, DT22, DT25, DT26, DT27

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

Set DT20 = (new DropTarget)(sw20, sw20a, BM_TargetR1, 20, 0, False)
Set DT21 = (new DropTarget)(sw21, sw21a, BM_TargetR2, 21, 0, False)
Set DT22 = (new DropTarget)(sw22, sw22a, BM_TargetR3, 22, 0, False)
Set DT25 = (new DropTarget)(sw25, sw25a, BM_TargetM3, 25, 0, False)
Set DT26 = (new DropTarget)(sw26, sw26a, BM_TargetM2, 26, 0, False)
Set DT27 = (new DropTarget)(sw27, sw27a, BM_TargetM1, 27, 0, False)

Dim DTArray
DTArray = Array(DT20, DT21, DT22, DT25, DT26, DT27)

'Configure the behavior of Drop Targets.
Const DTDropSpeed = 80 'in milliseconds
Const DTDropUpSpeed = 40 'in milliseconds
Const DTDropUnits = 48 'VP units primitive drops so top of at or below the playfield
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
    If animate = 1 Then secondary.collidable = 1 Else secondary.collidable = 0
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

Sub UpdateDropTargets
  dim LM, tz, rx, ry

    tz = BM_TargetR1.transz
  rx = BM_TargetR1.rotx
  ry = BM_TargetR1.roty
  For each LM in BP_TargetR1 : LM.transz = tz: LM.rotx = rx: LM.roty = ry: Next

    tz = BM_TargetR2.transz
  rx = BM_TargetR2.rotx
  ry = BM_TargetR2.roty
  For each LM in BP_TargetR2 : LM.transz = tz: LM.rotx = rx: LM.roty = ry: Next

    tz = BM_TargetR3.transz
  rx = BM_TargetR3.rotx
  ry = BM_TargetR3.roty
  For each LM in BP_TargetR3 : LM.transz = tz: LM.rotx = rx: LM.roty = ry: Next

    tz = BM_TargetM1.transz
  rx = BM_TargetM1.rotx
  ry = BM_TargetM1.roty
  For each LM in BP_TargetM1 : LM.transz = tz: LM.rotx = rx: LM.roty = ry: Next

    tz = BM_TargetM2.transz
  rx = BM_TargetM2.rotx
  ry = BM_TargetM2.roty
  For each LM in BP_TargetM2 : LM.transz = tz: LM.rotx = rx: LM.roty = ry: Next

    tz = BM_TargetM3.transz
  rx = BM_TargetM3.rotx
  ry = BM_TargetM3.roty
  For each LM in BP_TargetM3 : LM.transz = tz: LM.rotx = rx: LM.roty = ry: Next
End Sub


'******************************************************
'  DROP TARGET
'  SUPPORTING FUNCTIONS
'******************************************************

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
Dim ST17, ST18, ST19, ST35

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

Set ST17 = (new StandupTarget)(sw17, BM_TargetL3,17, 0)
Set ST18 = (new StandupTarget)(sw18, BM_TargetL2,18, 0)
Set ST19 = (new StandupTarget)(sw19, BM_TargetL1,19, 0)
Set ST35 = (new StandupTarget)(sw35, BM_Target_sw35,35, 0)

'Add all the Stand-up Target Arrays to Stand-up Target Animation Array
'   STAnimationArray = Array(ST1, ST2, ....)
Dim STArray
STArray = Array(ST17, ST18, ST19, ST35)

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
  dim LM, ty

    ty = BM_TargetL1.transy
  For each LM in BP_TargetL1 : LM.transy = ty: Next

    ty = BM_TargetL2.transy
  For each LM in BP_TargetL2 : LM.transy = ty: Next

  ty = BM_TargetL3.transy
  For each LM in BP_TargetL3 : LM.transy = ty: Next
End Sub


'******************************************************
'****   END STAND-UP TARGETS
'******************************************************




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
  '   Dim BOT
  '   BOT = GetBalls

  ' stop the sound of deleted balls
  For b = UBound(gBOT) + 1 To tnob - 1
    ' Comment the next line if you are not implementing Dyanmic Ball Shadows
    If AmbientBallShadowOn = 0 Then BallShadowA(b).visible = 0
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

    ' "Static" Ball Shadows
    ' Comment the next If block, if you are not implementing the Dynamic Ball Shadows
    If AmbientBallShadowOn = 0 Then
      If gBOT(b).Z > 30 Then
        BallShadowA(b).height = gBOT(b).z - BallSize / 4    'This is technically 1/4 of the ball "above" the ramp, but it keeps it from clipping the ramp
      Else
        BallShadowA(b).height = 0.1
      End If
      BallShadowA(b).Y = gBOT(b).Y + offsetY
      BallShadowA(b).X = gBOT(b).X + offsetX
      BallShadowA(b).visible = 1
    End If

  Next
End Sub

'******************************************************
'****  END BALL ROLLING AND DROP SOUNDS
'******************************************************


'******************************************************
'   ZBBR: BALL BRIGHTNESS
'******************************************************

Const BallBrightness =  0.5        'Ball brightness - Value between 0 and 1 (0=Dark ... 1=Bright)
Const PLOffset = 0.2
Dim PLGain: PLGain = (1-PLOffset)/(1260-2000)

Sub UpdateBallBrightness
  Dim s, b_base, b_r, b_g, b_b, d_w
  b_base = 120 * BallBrightness + 65*LightLevel/100 + 70*gilvl

  For s = 0 To UBound(gBOT)
    ' Handle z direction
    d_w = b_base*(1 - (gBOT(s).z-25)/500)
    If d_w < 30 Then d_w = 30
    ' Handle plunger lane
    If InRect(gBOT(s).x,gBOT(s).y,870,2000,870,1260,930,1260,930,2000) Then
      d_w = d_w*(PLOffset+PLGain*(gBOT(s).y-2000))
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




'******************************************************
'   ZRRL: RAMP ROLLING SFX
'******************************************************


Sub RampTrigger001_hit
  WireRampOn true
End Sub

Sub RampTrigger001_unhit
  if activeball.vely > 0 then WireRampOff
End Sub

Sub RampTrigger002_hit
  WireRampOn true
End Sub

Sub RampTrigger002_unhit
  if activeball.vely > 0 then WireRampOff
End Sub

Sub RampTrigger003_hit
  WireRampOff
End Sub

Sub RampTrigger004_hit
  WireRampOff
End Sub

Sub RampTrigger005_hit
  WireRampOn false
End Sub

Sub RampTrigger006_hit
  WireRampOn false
End Sub


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
Dim RampBalls(3,2)
'x,0 = ball x,1 = ID, 2 = Protection against ending early (minimum amount of updates)

'0,0 is boolean on/off, 0,1 unused for now
RampBalls(0,0) = False

' RampType
' Setup: Set this array to the number Total number of balls that can be tracked at one time + 1.  5 ball multiball then set value to 6
' Description: Array type indexed on BallId and a values used to deterimine what type of ramp the ball is on: False = Wire Ramp, True = Plastic Ramp
Dim RampType(3)

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
' Tutorial vides by Apophis
' Audio : Adding Fleep Part 1       https://youtu.be/rG35JVHxtx4
' Audio : Adding Fleep Part 2       https://youtu.be/dk110pWMxGo
' Audio : Adding Fleep Part 3       https://youtu.be/ESXWGJZY_EI


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

  FlipperCradleCollision ball1, ball2, velocity

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
'****  FLEEP MECHANICAL SOUNDS
'******************************************************





'***************************************************************
' VPW DYNAMIC BALL SHADOWS by Iakki, Apophis, and Wylte (Ambient-only version)
'***************************************************************

Const AmbientBallShadowOn = 1   '0 = Static shadow under ball ("flasher" image, like JP's), 1 = Moving ball shadow ("primitive" object, like ninuzzu's) - This is the only one that behaves like a true shadow!, 2 = flasher image shadow, but it moves like ninuzzu's

'Ambient (Room light source)
Const AmbientBSFactor = 0.9  '0 To 1, higher is darker
Const AmbientMovement = 1    '1+ higher means more movement as the ball moves left and right
Const offsetX = 0        'Offset x position under ball (These are if you want to change where the "room" light is for calculating the shadow position,)
Const offsetY = 5        'Offset y position under ball (^^for example 5,5 if the light is in the back left corner)

' *** Trim or extend these to match the number of balls/primitives/flashers on the table!  (will throw errors if there aren't enough objects)
Dim objBallShadow(5)
Dim OnPF(5)
Dim BallShadowA
BallShadowA = Array (BallShadowA0,BallShadowA1,BallShadowA2,BallShadowA3,BallShadowA4)

' *** The Shadow Dictionary
Dim bsDict
Set bsDict = New cvpmDictionary
Const bsNone = "None"
Const bsWire = "Wire"
Const bsRamp = "Ramp"
Const bsRampClear = "Clear"

'Initialization
BSInit

Sub BSInit()
  Dim iii
  'Prepare the shadow objects before play begins
  For iii = 0 To tnob - 1

    Set objBallShadow(iii) = Eval("BallShadow" & iii)
    objBallShadow(iii).material = "BallShadow" & iii
    UpdateMaterial objBallShadow(iii).material,1,0,0,0,0,0,AmbientBSFactor,RGB(0,0,0),0,0,False,True,0,0,0,0
    objBallShadow(iii).Z = 1 + iii / 1000 + 0.04
    objBallShadow(iii).visible = 0

    BallShadowA(iii).Opacity = 100 * AmbientBSFactor
    BallShadowA(iii).visible = 0
  Next
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

Sub BSUpdate
  Dim falloff 'Max distance to light sources, can be changed dynamically if you have a reason
  falloff = 150
  Dim ShadowOpacity1, ShadowOpacity2
  Dim s, LSd, iii
  Dim dist1, dist2, src1, src2
  Dim bsRampType
  '   Dim gBOT: gBOT=getballs 'Uncomment if you're destroying balls - Not recommended! #SaveTheBalls

  'Hide shadow of deleted balls
  For s = UBound(gBOT) + 1 To tnob - 1
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




'***********************************************************************
'* TABLE OPTIONS *******************************************************
'***********************************************************************



''*******************************************
''  ZOPT: User Options
''*******************************************
'

' Options
Const Opt_Light = 0
Const Opt_LUT = 1
Const Opt_Outpost = 2
Const Opt_Volume = 3
Const Opt_Volume_Ramp = 4
Const Opt_Volume_Ball = 5
' Modulated Flashers
Const Opt_ModSol = 6
'Const Opt_VR_Room = 5 ' Only when in VR
' Informations
Const Opt_Info_1 = 7
Const Opt_Info_2 = 8

Const NOptions = 9

Const FlexDMD_RenderMode_DMD_GRAY_2 = 0
Const FlexDMD_RenderMode_DMD_GRAY_4 = 1
Const FlexDMD_RenderMode_DMD_RGB = 2
Const FlexDMD_RenderMode_SEG_2x16Alpha = 3
Const FlexDMD_RenderMode_SEG_2x20Alpha = 4
Const FlexDMD_RenderMode_SEG_2x7Alpha_2x7Num = 5
Const FlexDMD_RenderMode_SEG_2x7Alpha_2x7Num_4x1Num = 6
Const FlexDMD_RenderMode_SEG_2x7Num_2x7Num_4x1Num = 7
Const FlexDMD_RenderMode_SEG_2x7Num_2x7Num_10x1Num = 8
Const FlexDMD_RenderMode_SEG_2x7Num_2x7Num_4x1Num_gen7 = 9
Const FlexDMD_RenderMode_SEG_2x7Num10_2x7Num10_4x1Num = 10
Const FlexDMD_RenderMode_SEG_2x6Num_2x6Num_4x1Num = 11
Const FlexDMD_RenderMode_SEG_2x6Num10_2x6Num10_4x1Num = 12
Const FlexDMD_RenderMode_SEG_4x7Num10 = 13
Const FlexDMD_RenderMode_SEG_6x4Num_4x1Num = 14
Const FlexDMD_RenderMode_SEG_2x7Num_4x1Num_1x16Alpha = 15
Const FlexDMD_RenderMode_SEG_1x16Alpha_1x16Num_1x7Num = 16

Const FlexDMD_Align_TopLeft = 0
Const FlexDMD_Align_Top = 1
Const FlexDMD_Align_TopRight = 2
Const FlexDMD_Align_Left = 3
Const FlexDMD_Align_Center = 4
Const FlexDMD_Align_Right = 5
Const FlexDMD_Align_BottomLeft = 6
Const FlexDMD_Align_Bottom = 7
Const FlexDMD_Align_BottomRight = 8

Const FlexDMD_Scaling_Fit = 0
Const FlexDMD_Scaling_Fill = 1
Const FlexDMD_Scaling_FillX = 2
Const FlexDMD_Scaling_FillY = 3
Const FlexDMD_Scaling_Stretch = 4
Const FlexDMD_Scaling_StretchX = 5
Const FlexDMD_Scaling_StretchY = 6
Const FlexDMD_Scaling_None = 7

Const FlexDMD_Interpolation_Linear = 0
Const FlexDMD_Interpolation_ElasticIn = 1
Const FlexDMD_Interpolation_ElasticOut = 2
Const FlexDMD_Interpolation_ElasticInOut = 3
Const FlexDMD_Interpolation_QuadIn = 4
Const FlexDMD_Interpolation_QuadOut = 5
Const FlexDMD_Interpolation_QuadInOut = 6
Const FlexDMD_Interpolation_CubeIn = 7
Const FlexDMD_Interpolation_CubeOut = 8
Const FlexDMD_Interpolation_CubeInOut = 9
Const FlexDMD_Interpolation_QuartIn = 10
Const FlexDMD_Interpolation_QuartOut = 11
Const FlexDMD_Interpolation_QuartInOut = 12
Const FlexDMD_Interpolation_QuintIn = 13
Const FlexDMD_Interpolation_QuintOut = 14
Const FlexDMD_Interpolation_QuintInOut = 15
Const FlexDMD_Interpolation_SineIn = 16
Const FlexDMD_Interpolation_SineOut = 17
Const FlexDMD_Interpolation_SineInOut = 18
Const FlexDMD_Interpolation_BounceIn = 19
Const FlexDMD_Interpolation_BounceOut = 20
Const FlexDMD_Interpolation_BounceInOut = 21
Const FlexDMD_Interpolation_CircIn = 22
Const FlexDMD_Interpolation_CircOut = 23
Const FlexDMD_Interpolation_CircInOut = 24
Const FlexDMD_Interpolation_ExpoIn = 25
Const FlexDMD_Interpolation_ExpoOut = 26
Const FlexDMD_Interpolation_ExpoInOut = 27
Const FlexDMD_Interpolation_BackIn = 28
Const FlexDMD_Interpolation_BackOut = 29
Const FlexDMD_Interpolation_BackInOut = 30

Dim OptionDMD: Set OptionDMD = Nothing
Dim bOptionsMagna, bInOptions : bOptionsMagna = False
Dim OptPos, OptSelected, OptN, OptTop, OptBot, OptSel
Dim OptFontHi, OptFontLo

Sub Options_Open
  bOptionsMagna = False
  On Error Resume Next
  Set OptionDMD = CreateObject("FlexDMD.FlexDMD")
  On Error Goto 0
  If OptionDMD is Nothing Then
    Debug.Print "FlexDMD is not installed"
    Debug.Print "Option UI can not be opened"
    MsgBox "You need to install FlexDMD to access table options"
    Exit Sub
  End If
  If Table1.ShowDT Then OptionDMDFlasher.RotX = -(Table1.Inclination + Table1.Layback)
  bInOptions = True
  OptPos = 0
  OptSelected = False
  OptionDMD.Show = False
  OptionDMD.RenderMode = FlexDMD_RenderMode_DMD_GRAY_4
  OptionDMD.Width = 128
  OptionDMD.Height = 32
  OptionDMD.Clear = True
  OptionDMD.Run = True
  Dim a, scene, font
  Set scene = OptionDMD.NewGroup("Scene")
  Set OptFontHi = OptionDMD.NewFont("FlexDMD.Resources.teeny_tiny_pixls-5.fnt", vbWhite, vbWhite, 0)
  Set OptFontLo = OptionDMD.NewFont("FlexDMD.Resources.teeny_tiny_pixls-5.fnt", RGB(100, 100, 100), RGB(100, 100, 100), 0)
  Set OptSel = OptionDMD.NewGroup("Sel")
  Set a = OptionDMD.NewLabel(">", OptFontLo, ">>>")
  a.SetAlignedPosition 1, 16, FlexDMD_Align_Left
  OptSel.AddActor a
  Set a = OptionDMD.NewLabel(">", OptFontLo, "<<<")
  a.SetAlignedPosition 127, 16, FlexDMD_Align_Right
  OptSel.AddActor a
  scene.AddActor OptSel
  OptSel.SetBounds 0, 0, 128, 32
  OptSel.Visible = False

  Set a = OptionDMD.NewLabel("Info1", OptFontLo, "MAGNA EXIT/ENTER")
  a.SetAlignedPosition 1, 32, FlexDMD_Align_BottomLeft
  scene.AddActor a
  Set a = OptionDMD.NewLabel("Info2", OptFontLo, "FLIPPER SELECT")
  a.SetAlignedPosition 127, 32, FlexDMD_Align_BottomRight
  scene.AddActor a
  Set OptN = OptionDMD.NewLabel("Pos", OptFontLo, "LINE 1")
  Set OptTop = OptionDMD.NewLabel("Top", OptFontLo, "LINE 1")
  Set OptBot = OptionDMD.NewLabel("Bottom", OptFontLo, "LINE 2")
  scene.AddActor OptN
  scene.AddActor OptTop
  scene.AddActor OptBot
  Options_OnOptChg
  OptionDMD.LockRenderThread
  OptionDMD.Stage.AddActor scene
  OptionDMD.UnlockRenderThread
  OptionDMDFlasher.Visible = True
End Sub

Sub Options_UpdateDMD
  If OptionDMD is Nothing Then Exit Sub
  Dim DMDp: DMDp = OptionDMD.DmdPixels
  If Not IsEmpty(DMDp) Then
    OptionDMDFlasher.DMDWidth = OptionDMD.Width
    OptionDMDFlasher.DMDHeight = OptionDMD.Height
    OptionDMDFlasher.DMDPixels = DMDp
  End If
End Sub

Sub Options_Close
  bInOptions = False
  OptionDMDFlasher.Visible = False
  If OptionDMD is Nothing Then Exit Sub
  OptionDMD.Run = False
  Set OptionDMD = Nothing
End Sub

Function Options_OnOffText(opt)
  If opt Then
    Options_OnOffText = "ON"
  Else
    Options_OnOffText = "OFF"
  End If
End Function

Sub Options_OnOptChg
  If OptionDMD is Nothing Then Exit Sub
  OptionDMD.LockRenderThread
  OptN.Text = (OptPos+1) & "/" & NOptions
  If OptSelected Then
    OptTop.Font = OptFontLo
    OptBot.Font = OptFontHi
    OptSel.Visible = True
  Else
    OptTop.Font = OptFontHi
    OptBot.Font = OptFontLo
    OptSel.Visible = False
  End If
  If OptPos = Opt_Light Then
    OptTop.Text = "LIGHT LEVEL (Req Restart)"
    OptBot.Text = "LEVEL " & LightLevel
    SaveValue cGameName, "LIGHT", LightLevel
  ElseIf OptPos = Opt_LUT Then
    OptTop.Text = "COLOR SATURATION"
    if ColorLUT = 1 Then OptBot.text = "DISABLED"
    if ColorLUT = 2 Then OptBot.text = "DESATURATED -10%"
    if ColorLUT = 3 Then OptBot.text = "DESATURATED -20%"
    if ColorLUT = 4 Then OptBot.text = "DESATURATED -30%"
    if ColorLUT = 5 Then OptBot.text = "DESATURATED -40%"
    if ColorLUT = 6 Then OptBot.text = "DESATURATED -50%"
    if ColorLUT = 7 Then OptBot.text = "DESATURATED -60%"
    if ColorLUT = 8 Then OptBot.text = "DESATURATED -70%"
    if ColorLUT = 9 Then OptBot.text = "DESATURATED -80%"
    if ColorLUT = 10 Then OptBot.text = "DESATURATED -90%"
    if ColorLUT = 11 Then OptBot.text = "BLACK'N WHITE"
    SaveValue cGameName, "LUT", ColorLUT
  ElseIf OptPos = Opt_Outpost Then
    OptTop.Text = "OUTPOST DIFFICULTY"
    If OutlaneDifficulty = 0 Then
      OptBot.Text = "EASY"
    ElseIf OutlaneDifficulty = 1 Then
      OptBot.Text = "MEDIUM"
    ElseIf OutlaneDifficulty = 2 Then
      OptBot.Text = "HARD"
    End If
    SaveValue cGameName, "OUTPOST", OutlaneDifficulty
  ElseIf OptPos = Opt_Volume Then
    OptTop.Text = "MECH VOLUME"
    OptBot.Text = "LEVEL " & CInt(VolumeDial * 100)
    SaveValue cGameName, "VOLUME", VolumeDial
  ElseIf OptPos = Opt_Volume_Ramp Then
    OptTop.Text = "RAMP VOLUME"
    OptBot.Text = "LEVEL " & CInt(RampRollVolume * 100)
    SaveValue cGameName, "RAMPVOLUME", RampRollVolume
  ElseIf OptPos = Opt_Volume_Ball Then
    OptTop.Text = "BALL VOLUME"
    OptBot.Text = "LEVEL " & CInt(BallRollVolume * 100)
    SaveValue cGameName, "BALLVOLUME", BallRollVolume
  ElseIf OptPos = Opt_ModSol Then
    OptTop.Text = "FLASHERS (Req Restart)"
    If ModSol = 1 Then OptBot.Text = "SIMPLE"
    If ModSol = 2 Then OptBot.Text = "PWM"
    SaveValue cGameName, "MODSOL", ModSol
' ElseIf OptPos = Opt_VR_Room Then
'   OptTop.Text = "VR ROOM"
'   If VRRoomChoice = 0 Then OptBot.Text = "MINIMAL"
'   If VRRoomChoice = 1 Then OptBot.Text = "ULTRA"
'   SaveValue cGameName, "VRROOM", VRRoomChoice
  ElseIf OptPos = Opt_Info_1 Then
    OptTop.Text = "VPX " & VersionMajor & "." & VersionMinor & "." & VersionRevision
    OptBot.Text = "POLICE FORCE " & TableVersion
  ElseIf OptPos = Opt_Info_2 Then
    OptTop.Text = "RENDER MODE"
    If RenderingMode = 0 Then OptBot.Text = "DEFAULT"
    If RenderingMode = 1 Then OptBot.Text = "STEREO 3D"
    If RenderingMode = 2 Then OptBot.Text = "VR"
  End If
  OptTop.Pack
  OptTop.SetAlignedPosition 127, 1, FlexDMD_Align_TopRight
  OptBot.SetAlignedPosition 64, 16, FlexDMD_Align_Center
  OptionDMD.UnlockRenderThread
  UpdateMods
End Sub

Sub Options_Toggle(amount)
  If OptionDMD is Nothing Then Exit Sub
  If OptPos = Opt_Light Then
    LightLevel = LightLevel + amount * 10
    If LightLevel < 0 Then LightLevel = 100
    If LightLevel > 100 Then LightLevel = 0
  ElseIf OptPos = Opt_LUT Then
    ColorLUT = ColorLUT + amount * 1
    If ColorLUT < 1 Then ColorLUT = 11
    If ColorLUT > 11 Then ColorLUT = 1
  ElseIf OptPos = Opt_Outpost Then
    OutlaneDifficulty = OutlaneDifficulty + amount
    If OutlaneDifficulty < 0 Then OutlaneDifficulty = 2
    If OutlaneDifficulty > 2 Then OutlaneDifficulty = 0
  ElseIf OptPos = Opt_Volume Then
    VolumeDial = VolumeDial + amount * 0.1
    If VolumeDial < 0 Then VolumeDial = 1
    If VolumeDial > 1 Then VolumeDial = 0
  ElseIf OptPos = Opt_Volume_Ramp Then
    RampRollVolume = RampRollVolume + amount * 0.1
    If RampRollVolume < 0 Then RampRollVolume = 1
    If RampRollVolume > 1 Then RampRollVolume = 0
  ElseIf OptPos = Opt_Volume_Ball Then
    BallRollVolume = BallRollVolume + amount * 0.1
    If BallRollVolume < 0 Then BallRollVolume = 1
    If BallRollVolume > 1 Then BallRollVolume = 0
  ElseIf OptPos = Opt_ModSol Then
    ModSol = ModSol + amount
    If ModSol < 1 Then ModSol = 2
    If ModSol > 2 Then ModSol = 1
' ElseIf OptPos = Opt_VR_Room Then
'   VRRoomChoice = VRRoomChoice + amount
'   If VRRoomChoice < 0 Then VRRoomChoice = 2
'   If VRRoomChoice > 2 Then VRRoomChoice = 0
  End If
End Sub

Sub Options_KeyDown(ByVal keycode)
  If OptSelected Then
    If keycode = LeftMagnaSave Then ' Exit / Cancel
      OptSelected = False
    ElseIf keycode = RightMagnaSave Then ' Enter / Select
      OptSelected = False
    ElseIf keycode = LeftFlipperKey Then ' Next / +
      Options_Toggle  -1
    ElseIf keycode = RightFlipperKey Then ' Prev / -
      Options_Toggle  1
    End If
  Else
    If keycode = LeftMagnaSave Then ' Exit / Cancel
      Options_Close
    ElseIf keycode = RightMagnaSave Then ' Enter / Select
      If OptPos < Opt_Info_1 Then OptSelected = True
    ElseIf keycode = LeftFlipperKey Then ' Next / +
      OptPos = OptPos - 1
      'If OptPos = Opt_VR_Room And RenderingMode <> 2 Then OptPos = OptPos - 1 ' Skip VR option in non VR mode
      'If OptPos = Opt_Cabinet And RenderingMode = 2 Then OptPos = OptPos - 1 ' Skip non VR option in VR mode
      If OptPos < 0 Then OptPos = NOptions - 1
    ElseIf keycode = RightFlipperKey Then ' Prev / -
      OptPos = OptPos + 1
      'If OptPos = Opt_VR_Room And RenderingMode <> 2 Then OptPos = OptPos + 1 ' Skip VR option in non VR mode
      'If OptPos = Opt_Cabinet And RenderingMode = 2 Then OptPos = OptPos + 1 ' Skip non VR option in VR mode
      If OptPos >= NOPtions Then OptPos = 0
    End If
  End If
  Options_OnOptChg
End Sub

Sub Options_Load
  Dim x
    x = LoadValue(cGameName, "LIGHT") : If x <> "" Then LightLevel = CInt(x) Else LightLevel = 50
    x = LoadValue(cGameName, "LUT") : If x <> "" Then ColorLUT = CInt(x) Else ColorLUT = 1
    x = LoadValue(cGameName, "OUTPOST") : If x <> "" Then OutlaneDifficulty = CInt(x) Else OutlaneDifficulty = 2
    x = LoadValue(cGameName, "VOLUME") : If x <> "" Then VolumeDial = CNCDbl(x) Else VolumeDial = 0.8
    x = LoadValue(cGameName, "RAMPVOLUME") : If x <> "" Then RampRollVolume = CNCDbl(x) Else RampRollVolume = 0.5
    x = LoadValue(cGameName, "BALLVOLUME") : If x <> "" Then BallRollVolume = CNCDbl(x) Else BallRollVolume = 0.5
  x = LoadValue(cGameName, "MODSOL") : If x <> "" Then ModSol = CInt(x) Else ModSol = 2
'    x = LoadValue(cGameName, "VRROOM") : If x <> "" Then VRRoomChoice = CInt(x) Else VRRoomChoice = 0
  UpdateMods
End Sub

Sub UpdateMods
  Dim v, BP

  ' Lighting level
  'SetRoomBrightness LightLevel/100

  ' Desaturation
  if ColorLUT = 1 Then Table1.ColorGradeImage = "ColorGrade_8"
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

  ' Outlane difficulty
  Select Case OutlaneDifficulty
    Case 0
      ic_RubberPost_LHard.collidable = false: ic_RubberPost_LMedium.collidable = false: ic_RubberPost_LEasy.collidable = true
      ic_RubberPost_RHard.collidable = false: ic_RubberPost_RMedium.collidable = false: ic_RubberPost_REasy.collidable = true
      ic_RubberBand_Hard.IsDropped = true: ic_RubberBand_Medium.IsDropped = true: ic_RubberBand_Easy.IsDropped = false
      For Each BP in BP_SpecialPost: BP.y = 1364.8: Next
    Case 1
      ic_RubberPost_LHard.collidable = false: ic_RubberPost_LMedium.collidable = true: ic_RubberPost_LEasy.collidable = false
      ic_RubberPost_RHard.collidable = false: ic_RubberPost_RMedium.collidable = true: ic_RubberPost_REasy.collidable = false
      ic_RubberBand_Hard.IsDropped = true: ic_RubberBand_Medium.IsDropped = false: ic_RubberBand_Easy.IsDropped = true
      For Each BP in BP_SpecialPost: BP.y = 1356: Next
    Case Else
      ic_RubberPost_LHard.collidable = true: ic_RubberPost_LMedium.collidable = false: ic_RubberPost_LEasy.collidable = false
      ic_RubberPost_RHard.collidable = true: ic_RubberPost_RMedium.collidable = false: ic_RubberPost_REasy.collidable = false
      ic_RubberBand_Hard.IsDropped = false: ic_RubberBand_Medium.IsDropped = true: ic_RubberBand_Easy.IsDropped = true
      For Each BP in BP_SpecialPost: BP.y = 1345.4: Next
  End Select

  'PWM flashers
  If ModSol = 1 Then
    UseVPMModSol = false
  ElseIf ModSol = 2 Then
    UseVPMModSol = true
  End If

End Sub

' Culture neutral string to double conversion (handles situation where you don't know how the string was written)
Function CNCDbl(str)
    Dim strt, Sep, i
    If IsNumeric(str) Then
        CNCDbl = CDbl(str)
    Else
        Sep = Mid(CStr(0.5), 2, 1)
        Select Case Sep
        Case "."
            i = InStr(1, str, ",")
        Case ","
            i = InStr(1, str, ".")
        End Select
        If i = 0 Then
            CNCDbl = Empty
        Else
            strt = Mid(str, 1, i - 1) & Sep & Mid(str, i + 1)
            If IsNumeric(strt) Then
                CNCDbl = CDbl(strt)
            Else
                CNCDbl = Empty
            End If
        End If
    End If
End Function




'*************************************
'* VLM Stuff
'*************************************



' ===============================================================
' The following code can be copy/pasted to have premade array for
' VLM Visuals Room Brightness:

' ===============================================================

'***************************************************************************
' VR Plunger Animataion Code
'***************************************************************************
Dim Pincab_ShooterYStart: Pincab_ShooterYStart = Pincab_Shooter.Y
Dim PlungerYStart : PlungerYStart= Plunger.Y

Sub TimerVRPlunger_Timer
  If Pincab_Shooter.Y < (Pincab_ShooterYStart + 100) then Pincab_Shooter.Y = Pincab_Shooter.Y + 5
End Sub

Sub TimerVRPlunger2_Timer
  Pincab_Shooter.Y = PlungerYStart - Plunger.Y + Pincab_ShooterYStart + (5* Plunger.Position)
End Sub



'**********************************************************************************************************
'Alphanumerics
'**********************************************************************************************************

Dim Digits(40)
Const DigitsOffset=1

' Jackpot Digits
Digits(DigitsOffset+0)=Array(LED1,LED2,LED3,LED4,LED5,LED6,LED7,LED8)
Digits(DigitsOffset+1)=Array(LED9,LED10,LED11,LED12,LED13,LED14,LED15)
Digits(DigitsOffset+2)=Array(LED16,LED17,LED18,LED19,LED20,LED21,LED22)
Digits(DigitsOffset+3)=Array(LED23,LED24,LED25,LED26,LED27,LED28,LED29,LED30)
Digits(DigitsOffset+4)=Array(LED31,LED32,LED33,LED34,LED35,LED36,LED37)
Digits(DigitsOffset+5)=Array(LED38,LED39,LED40,LED41,LED42,LED43,LED44)
Digits(DigitsOffset+6)=Array(LED45,LED46,LED47,LED48,LED49,LED50,LED51)

Digits(DigitsOffset+7)=Array(a00, a05, a0c, a0d, a08, a01, a06, a0f, a02, a03, a04, a07, a0b, a0a, a09, a0e)
Digits(DigitsOffset+8)=Array(a10, a15, a1c, a1d, a18, a11, a16, a1f, a12, a13, a14, a17, a1b, a1a, a19, a1e)
Digits(DigitsOffset+9)=Array(a20, a25, a2c, a2d, a28, a21, a26, a2f, a22, a23, a24, a27, a2b, a2a, a29, a2e)
Digits(DigitsOffset+10)=Array(a30, a35, a3c, a3d, a38, a31, a36, a3f, a32, a33, a34, a37, a3b, a3a, a39, a3e)
Digits(DigitsOffset+11)=Array(a40, a45, a4c, a4d, a48, a41, a46, a4f, a42, a43, a44, a47, a4b, a4a, a49, a4e)
Digits(DigitsOffset+12)=Array(a50, a55, a5c, a5d, a58, a51, a56, a5f, a52, a53, a54, a57, a5b, a5a, a59, a5e)
Digits(DigitsOffset+13)=Array(a60, a65, a6c, a6d, a68, a61, a66, a6f, a62, a63, a64, a67, a6b, a6a, a69, a6e)
Digits(DigitsOffset+14)=Array(a70, a75, a7c, a7d, a78, a71, a76, a7f, a72, a73, a74, a77, a7b, a7a, a79, a7e)
Digits(DigitsOffset+15)=Array(a80, a85, a8c, a8d, a88, a81, a86, a8f, a82, a83, a84, a87, a8b, a8a, a89, a8e)
Digits(DigitsOffset+16)=Array(a90, a95, a9c, a9d, a98, a91, a96, a9f, a92, a93, a94, a97, a9b, a9a, a99, a9e)
Digits(DigitsOffset+17)=Array(aa0, aa5, aac, aad, aa8, aa1, aa6, aaf, aa2, aa3, aa4, aa7, aab, aaa, aa9, aae)
Digits(DigitsOffset+18)=Array(ab0, ab5, abc, abd, ab8, ab1, ab6, abf, ab2, ab3, ab4, ab7, abb, aba, ab9, abe)
Digits(DigitsOffset+19)=Array(ac0, ac5, acc, acd, ac8, ac1, ac6, acf, ac2, ac3, ac4, ac7, acb, aca, ac9, ace)
Digits(DigitsOffset+20)=Array(ad0, ad5, adc, add, ad8, ad1, ad6, adf, ad2, ad3, ad4, ad7, adb, ada, ad9, ade)
Digits(DigitsOffset+21)=Array(ae0, ae5, aec, aed, ae8, ae1, ae6, aef, ae2, ae3, ae4, ae7, aeb, aea, ae9, aee)
Digits(DigitsOffset+22)=Array(af0, af5, afc, afd, af8, af1, af6, aff, af2, af3, af4, af7, afb, afa, af9, afe)

Digits(DigitsOffset+23)=Array(b00, b05, b0c, b0d, b08, b01, b06, b0f)
Digits(DigitsOffset+24)=Array(b10, b15, b1c, b1d, b18, b11, b16, b1f)
Digits(DigitsOffset+25)=Array(b20, b25, b2c, b2d, b28, b21, b26, b2f)
Digits(DigitsOffset+26)=Array(b30, b35, b3c, b3d, b38, b31, b36, b3f)
Digits(DigitsOffset+27)=Array(b40, b45, b4c, b4d, b48, b41, b46, b4f)
Digits(DigitsOffset+28)=Array(b50, b55, b5c, b5d, b58, b51, b56, b5f)
Digits(DigitsOffset+29)=Array(b60, b65, b6c, b6d, b68, b61, b66, b6f)
Digits(DigitsOffset+30)=Array(b70, b75, b7c, b7d, b78, b71, b76, b7f)
Digits(DigitsOffset+31)=Array(b80, b85, b8c, b8d, b88, b81, b86, b8f)
Digits(DigitsOffset+32)=Array(b90, b95, b9c, b9d, b98, b91, b96, b9f)
Digits(DigitsOffset+33)=Array(ba0, ba5, bac, bad, ba8, ba1, ba6, baf)
Digits(DigitsOffset+34)=Array(bb0, bb5, bbc, bbd, bb8, bb1, bb6, bbf)
Digits(DigitsOffset+35)=Array(bc0, bc5, bcc, bcd, bc8, bc1, bc6, bcf)
Digits(DigitsOffset+36)=Array(bd0, bd5, bdc, bdd, bd8, bd1, bd6, bdf)
Digits(DigitsOffset+37)=Array(be0, be5, bec, bed, be8, be1, be6, bef)
Digits(DigitsOffset+38)=Array(bf0, bf5, bfc, bfd, bf8, bf1, bf6, bff)


Dim DigitsVR(39)

' Jackpot Digits
DigitsVR(DigitsOffset+0)=Array(byg1, byg2, byg3 ,byg4, byg5, byg6, byg7, bygg)
DigitsVR(DigitsOffset+1)=Array(byh1, byh2, byh3 ,byh4, byh5, byh6, byh7, byhg)
DigitsVR(DigitsOffset+2)=Array(bya1, bya2, bya3, bya4, bya5, bya6, bya7, byag)
DigitsVR(DigitsOffset+3)=Array(byb1, byb2, byb3, byb4, byb5, byb6, byb7, bybg)
DigitsVR(DigitsOffset+4)=Array(byc1, byc2, byc3, byc4, byc5, byc6, byc7, bycg)
DigitsVR(DigitsOffset+5)=Array(byd1, byd2, byd3, byd4, byd5, byd6, byd7, bydg)
DigitsVR(DigitsOffset+6)=Array(bye1, bye2, bye3, bye4, bye5, bye6, bye7, byeg)

DigitsVR(DigitsOffset+7)=Array(ax00, ax05, ax0c, ax0d, ax08, ax01, ax06, ax0f, ax02, ax03, ax04, ax07, ax0b, ax0a, ax09, ax0e)
DigitsVR(DigitsOffset+8)=Array(ax10, ax15, ax1c, ax1d, ax18, ax11, ax16, ax1f, ax12, ax13, ax14, ax17, ax1b, ax1a, ax19, ax1e)
DigitsVR(DigitsOffset+9)=Array(ax20, ax25, ax2c, ax2d, ax28, ax21, ax26, ax2f, ax22, ax23, ax24, ax27, ax2b, ax2a, ax29, ax2e)
DigitsVR(DigitsOffset+10)=Array(ax30, ax35, ax3c, ax3d, ax38, ax31, ax36, ax3f, ax32, ax33, ax34, ax37, ax3b, ax3a, ax39, ax3e)
DigitsVR(DigitsOffset+11)=Array(ax40, ax45, ax4c, ax4d, ax48, ax41, ax46, ax4f, ax42, ax43, ax44, ax47, ax4b, ax4a, ax49, ax4e)
DigitsVR(DigitsOffset+12)=Array(ax50, ax55, ax5c, ax5d, ax58, ax51, ax56, ax5f, ax52, ax53, ax54, ax57, ax5b, ax5a, ax59, ax5e)
DigitsVR(DigitsOffset+13)=Array(ax60, ax65, ax6c, ax6d, ax68, ax61, ax66, ax6f, ax62, ax63, ax64, ax67, ax6b, ax6a, ax69, ax6e)
DigitsVR(DigitsOffset+14)=Array(ax70, ax75, ax7c, ax7d, ax78, ax71, ax76, ax7f, ax72, ax73, ax74, ax77, ax7b, ax7a, ax79, ax7e)
DigitsVR(DigitsOffset+15)=Array(ax80, ax85, ax8c, ax8d, ax88, ax81, ax86, ax8f, ax82, ax83, ax84, ax87, ax8b, ax8a, ax89, ax8e)
DigitsVR(DigitsOffset+16)=Array(ax90, ax95, ax9c, ax9d, ax98, ax91, ax96, ax9f, ax92, ax93, ax94, ax97, ax9b, ax9a, ax99, ax9e)
DigitsVR(DigitsOffset+17)=Array(axa0, axa5, axac, axad, axa8, axa1, axa6, axaf, axa2, axa3, axa4, axa7, axab, axaa, axa9, axae)
DigitsVR(DigitsOffset+18)=Array(axb0, axb5, axbc, axbd, axb8, axb1, axb6, axbf, axb2, axb3, axb4, axb7, axbb, axba, axb9, axbe)
DigitsVR(DigitsOffset+19)=Array(axc0, axc5, axcc, axcd, axc8, axc1, axc6, axcf, axc2, axc3, axc4, axc7, axcb, axca, axc9, axce)
DigitsVR(DigitsOffset+20)=Array(axd0, axd5, axdc, axdd, axd8, axd1, axd6, axdf, axd2, axd3, axd4, axd7, axdb, axda, axd9, axde)
DigitsVR(DigitsOffset+21)=Array(axe0, axe5, axec, axed, axe8, axe1, axe6, axef, axe2, axe3, axe4, axe7, axeb, axea, axe9, axee)
DigitsVR(DigitsOffset+22)=Array(axf0, axf5, axfc, axfd, axf8, axf1, axf6, axff, axf2, axf3, axf4, axf7, axfb, axfa, axf9, axfe)

DigitsVR(DigitsOffset+23) = Array(LED1x0,LED1x1,LED1x2,LED1x3,LED1x4,LED1x5,LED1x6)
DigitsVR(DigitsOffset+24) = Array(LED2x0,LED2x1,LED2x2,LED2x3,LED2x4,LED2x5,LED2x6)
DigitsVR(DigitsOffset+25) = Array(LED3x0,LED3x1,LED3x2,LED3x3,LED3x4,LED3x5,LED3x6)
DigitsVR(DigitsOffset+26) = Array(LED4x0,LED4x1,LED4x2,LED4x3,LED4x4,LED4x5,LED4x6)
DigitsVR(DigitsOffset+27) = Array(LED5x0,LED5x1,LED5x2,LED5x3,LED5x4,LED5x5,LED5x6)
DigitsVR(DigitsOffset+28) = Array(LED6x0,LED6x1,LED6x2,LED6x3,LED6x4,LED6x5,LED6x6)
DigitsVR(DigitsOffset+29) = Array(LED7x0,LED7x1,LED7x2,LED7x3,LED7x4,LED7x5,LED7x6)
DigitsVR(DigitsOffset+30) = Array(LED8x0,LED8x1,LED8x2,LED8x3,LED8x4,LED8x5,LED8x6)
DigitsVR(DigitsOffset+31) = Array(LED9x0,LED9x1,LED9x2,LED9x3,LED9x4,LED9x5,LED9x6)
DigitsVR(DigitsOffset+32) = Array(LED10x0,LED10x1,LED10x2,LED10x3,LED10x4,LED10x5,LED10x6)
DigitsVR(DigitsOffset+33) = Array(LED11x0,LED11x1,LED11x2,LED11x3,LED11x4,LED11x5,LED11x6)
DigitsVR(DigitsOffset+34) = Array(LED12x0,LED12x1,LED12x2,LED12x3,LED12x4,LED12x5,LED12x6)
DigitsVR(DigitsOffset+35) = Array(LED13x0,LED13x1,LED13x2,LED13x3,LED13x4,LED13x5,LED13x6)
DigitsVR(DigitsOffset+36) = Array(LED14x0,LED14x1,LED14x2,LED14x3,LED14x4,LED14x5,LED14x6)
DigitsVR(DigitsOffset+37) = Array(LEDax300,LEDax301,LEDax302,LEDax303,LEDax304,LEDax305,LEDax306)
DigitsVR(DigitsOffset+38) = Array(LEDbx400,LEDbx401,LEDbx402,LEDbx403,LEDbx404,LEDbx405,LEDbx406)


Sub DisplayTimer()
    Dim ChgLED, ii, jj, num, chg, stat, obj, b, x
    ChgLED=Controller.ChangedLEDs(&Hffffffff, &Hffffffff)
    If Not IsEmpty(ChgLED)Then
       For ii=0 To UBound(chgLED)
          num=chgLED(ii, 0) : chg=chgLED(ii, 1) : stat=chgLED(ii, 2)
      If (num < 40) then
        If VRRoom > 0 Then
          If (num >= DigitsOffset+0) and (num <= DigitsOffset+39) then
            For Each obj In DigitsVR(num)
              If chg And 1 Then obj.visible=stat And 1
              chg=chg\2 : stat=stat\2
            Next
          End If
        End If
        If DesktopMode = True and VRRoom = 0 Then
          If (num >= DigitsOffset+0) and (num <= DigitsOffset+39) then
            For Each obj In Digits(num)
              If chg And 1 Then obj.State=stat And 1
              chg=chg\2 : stat=stat\2
            Next
          End If
        End If
      End if
    Next
  End If
End Sub


'*************************************
'* VR Stuff
'*************************************


'Set Up VR Backglass
Sub SetBackglass()
  Dim obj

  For Each obj In VRBackglass
    obj.x = obj.x
    obj.height = - obj.y + 347
    obj.y = -20
    obj.rotX = -91
  Next
  For Each obj In VRBackglassHigh
    obj.x = obj.x
    obj.height = - obj.y + 347
    obj.y = -15
    obj.rotX = -91
  Next
  For Each obj In VRDigits
    obj.x = obj.x
    obj.height = - obj.y + 347
    obj.y = -30
    obj.rotX = -91
  Next
End Sub

'Backglass bulbs
Sub L63_Animate
  If VRRoom > 0 then
    Dim f
    f = L63.GetInPlayIntensity / L63.Intensity
    BGBulb63.blenddisablelighting =  1 * f
  End If
End Sub

Sub L64_Animate
  If VRRoom > 0 then
    Dim f
    f = L64.GetInPlayIntensity / L64.Intensity
    BGBulb64.blenddisablelighting =  1 * f
  End If
End Sub

' VLM Arrays - Start
' Arrays per baked part
Dim BP_Bumper1Ring: BP_Bumper1Ring=Array(BM_Bumper1Ring, LM_Flashers_f126_Bumper1Ring, LM_Flashers_f127_Bumper1Ring, LM_GISplit_gi08_Bumper1Ring, LM_GI_Bumper1Ring, LM_Inserts_l34_Bumper1Ring)
Dim BP_Bumper1Socket: BP_Bumper1Socket=Array(BM_Bumper1Socket, LM_Flashers_f126_Bumper1Socket, LM_GISplit_gi08_Bumper1Socket, LM_GI_Bumper1Socket)
Dim BP_Bumper2Ring: BP_Bumper2Ring=Array(BM_Bumper2Ring, LM_Flashers_f126_Bumper2Ring, LM_Flashers_f127_Bumper2Ring, LM_GI_Bumper2Ring)
Dim BP_Bumper2Socket: BP_Bumper2Socket=Array(BM_Bumper2Socket, LM_GISplit_gi08_Bumper2Socket, LM_GI_Bumper2Socket, LM_Inserts_l39_Bumper2Socket)
Dim BP_Bumper3Ring: BP_Bumper3Ring=Array(BM_Bumper3Ring, LM_Flashers_f125_Bumper3Ring, LM_Flashers_f126_Bumper3Ring, LM_Flashers_f127_Bumper3Ring, LM_Flashers_f128_Bumper3Ring, LM_Flashers_f129_Bumper3Ring, LM_GI_Bumper3Ring, LM_GISplit_gi01_Bumper3Ring, LM_Inserts_l30_Bumper3Ring)
Dim BP_Bumper3Socket: BP_Bumper3Socket=Array(BM_Bumper3Socket, LM_Flashers_f127_Bumper3Socket, LM_GI_Bumper3Socket, LM_GISplit_gi01_Bumper3Socket)
Dim BP_BumperInner: BP_BumperInner=Array(BM_BumperInner, LM_Flashers_f126_BumperInner, LM_Flashers_f127_BumperInner, LM_Flashers_f128_BumperInner, LM_GISplit_gi08_BumperInner, LM_GI_BumperInner, LM_GISplit_gi01_BumperInner, LM_Inserts_l30_BumperInner, LM_Inserts_l34_BumperInner, LM_Inserts_l47_BumperInner, LM_Inserts_l48_BumperInner)
Dim BP_BumperOuter: BP_BumperOuter=Array(BM_BumperOuter, LM_Flashers_f126_BumperOuter, LM_Flashers_f127_BumperOuter, LM_Flashers_f128_BumperOuter, LM_GI_BumperOuter, LM_GISplit_gi01_BumperOuter, LM_Inserts_l34_BumperOuter, LM_Inserts_l47_BumperOuter, LM_Inserts_l48_BumperOuter)
Dim BP_CRamp: BP_CRamp=Array(BM_CRamp, LM_Flashers_f126_CRamp, LM_Flashers_f127_CRamp, LM_GI_CRamp, LM_Inserts_l27_CRamp, LM_Inserts_l28_CRamp, LM_Inserts_l6_CRamp)
Dim BP_CRampEdges: BP_CRampEdges=Array(BM_CRampEdges, LM_Flashers_f109_CRampEdges, LM_Flashers_f113_CRampEdges, LM_Flashers_f125_CRampEdges, LM_Flashers_f126_CRampEdges, LM_Flashers_f127_CRampEdges, LM_Flashers_f128_CRampEdges, LM_Flashers_f129_CRampEdges, LM_Flashers_f130_CRampEdges, LM_GISplit_gi02_CRampEdges, LM_GISplit_gi08_CRampEdges, LM_GI_CRampEdges, LM_GISplit_gi01_CRampEdges, LM_Inserts_l16_CRampEdges, LM_Inserts_l27_CRampEdges, LM_Inserts_l28_CRampEdges, LM_Inserts_l29_CRampEdges, LM_Inserts_l30_CRampEdges, LM_Inserts_l36_CRampEdges, LM_Inserts_l49_CRampEdges, LM_Inserts_l52_CRampEdges)
Dim BP_Car: BP_Car=Array(BM_Car, LM_Flashers_f126_Car, LM_Flashers_f127_Car, LM_GI_Car, LM_Inserts_l47_Car, LM_Inserts_l48_Car)
Dim BP_DiverterArm: BP_DiverterArm=Array(BM_DiverterArm, LM_Flashers_f109_DiverterArm)
Dim BP_FlipperL: BP_FlipperL=Array(BM_FlipperL, LM_GISplit_gi05_FlipperL, LM_GI_FlipperL, LM_Inserts_l24_FlipperL)
Dim BP_FlipperR: BP_FlipperR=Array(BM_FlipperR, LM_GISplit_gi02_FlipperR, LM_GI_FlipperR, LM_Inserts_l24_FlipperR)
Dim BP_Gate003_Wire: BP_Gate003_Wire=Array(BM_Gate003_Wire)
Dim BP_Gate02Wire: BP_Gate02Wire=Array(BM_Gate02Wire, LM_Flashers_f114_Gate02Wire, LM_Flashers_f126_Gate02Wire, LM_GISplit_gi08_Gate02Wire, LM_GI_Gate02Wire)
Dim BP_Gate04Wire: BP_Gate04Wire=Array(BM_Gate04Wire, LM_GISplit_gi08_Gate04Wire, LM_GI_Gate04Wire)
Dim BP_Gate05Wire: BP_Gate05Wire=Array(BM_Gate05Wire, LM_Flashers_f113_Gate05Wire, LM_GI_Gate05Wire)
Dim BP_GateFlap: BP_GateFlap=Array(BM_GateFlap, LM_Flashers_f109_GateFlap, LM_Flashers_f126_GateFlap, LM_GISplit_gi08_GateFlap, LM_GI_GateFlap, LM_Inserts_l56_GateFlap)
Dim BP_GateSw44Wire: BP_GateSw44Wire=Array(BM_GateSw44Wire, LM_GI_GateSw44Wire)
Dim BP_GateSw45Wire: BP_GateSw45Wire=Array(BM_GateSw45Wire, LM_Flashers_f130_GateSw45Wire, LM_GI_GateSw45Wire)
Dim BP_GateSw46Wire: BP_GateSw46Wire=Array(BM_GateSw46Wire, LM_Flashers_f127_GateSw46Wire, LM_GI_GateSw46Wire)
Dim BP_LeftSling1: BP_LeftSling1=Array(BM_LeftSling1, LM_Flashers_f125_LeftSling1, LM_GISplit_gi05_LeftSling1, LM_GI_LeftSling1)
Dim BP_LeftSling2: BP_LeftSling2=Array(BM_LeftSling2, LM_Flashers_f125_LeftSling2, LM_GISplit_gi05_LeftSling2, LM_GI_LeftSling2)
Dim BP_LeftSling3: BP_LeftSling3=Array(BM_LeftSling3, LM_Flashers_f125_LeftSling3, LM_GISplit_gi05_LeftSling3, LM_GI_LeftSling3)
Dim BP_LeftSling4: BP_LeftSling4=Array(BM_LeftSling4, LM_Flashers_f125_LeftSling4, LM_Flashers_f127_LeftSling4, LM_GISplit_gi05_LeftSling4, LM_GI_LeftSling4, LM_Inserts_l17_LeftSling4, LM_Inserts_l51_LeftSling4)
Dim BP_Parts: BP_Parts=Array(BM_Parts, LM_Flashers_f109_Parts, LM_Flashers_f113_Parts, LM_Flashers_f114_Parts, LM_Flashers_f114a_Parts, LM_Flashers_f125_Parts, LM_Flashers_f126_Parts, LM_Flashers_f127_Parts, LM_Flashers_f128_Parts, LM_Flashers_f129_Parts, LM_Flashers_f129a_Parts, LM_Flashers_f130_Parts, LM_GISplit_gi02_Parts, LM_GISplit_gi05_Parts, LM_GISplit_gi08_Parts, LM_GI_Parts, LM_GISplit_gi01_Parts, LM_Inserts_l15_Parts, LM_Inserts_l16_Parts, LM_Inserts_l17_Parts, LM_Inserts_l28_Parts, LM_Inserts_l29_Parts, LM_Inserts_l30_Parts, LM_Inserts_l33_Parts, LM_Inserts_l34_Parts, LM_Inserts_l35_Parts, LM_Inserts_l36_Parts, LM_Inserts_l37_Parts, LM_Inserts_l38_Parts, LM_Inserts_l39_Parts, LM_Inserts_l40_Parts, LM_Inserts_l42_Parts, LM_Inserts_l43_Parts, LM_Inserts_l44_Parts, LM_Inserts_l45_Parts, LM_Inserts_l46_Parts, LM_Inserts_l47_Parts, LM_Inserts_l48_Parts, LM_Inserts_l49_Parts, LM_Inserts_l52_Parts, LM_Inserts_l54_Parts, LM_Inserts_l55_Parts, LM_Inserts_l56_Parts, LM_Inserts_l57_Parts, _
  LM_Inserts_l58_Parts, LM_Inserts_l8_Parts)
Dim BP_Playfield: BP_Playfield=Array(BM_Playfield, LM_Flashers_f109_Playfield, LM_Flashers_f113_Playfield, LM_Flashers_f114_Playfield, LM_Flashers_f114a_Playfield, LM_Flashers_f125_Playfield, LM_Flashers_f126_Playfield, LM_Flashers_f127_Playfield, LM_Flashers_f128_Playfield, LM_Flashers_f129_Playfield, LM_Flashers_f129a_Playfield, LM_Flashers_f130_Playfield, LM_GISplit_gi02_Playfield, LM_GISplit_gi05_Playfield, LM_GISplit_gi08_Playfield, LM_GI_Playfield, LM_GISplit_gi01_Playfield, LM_Inserts_l10_Playfield, LM_Inserts_l11_Playfield, LM_Inserts_l12_Playfield, LM_Inserts_l13_Playfield, LM_Inserts_l14_Playfield, LM_Inserts_l15_Playfield, LM_Inserts_l16_Playfield, LM_Inserts_l17_Playfield, LM_Inserts_l18_Playfield, LM_Inserts_l19_Playfield, LM_Inserts_l20_Playfield, LM_Inserts_l21_Playfield, LM_Inserts_l22_Playfield, LM_Inserts_l23_Playfield, LM_Inserts_l24_Playfield, LM_Inserts_l25_Playfield, LM_Inserts_l26_Playfield, LM_Inserts_l27_Playfield, LM_Inserts_l28_Playfield, LM_Inserts_l29_Playfield, _
  LM_Inserts_l30_Playfield, LM_Inserts_l31_Playfield, LM_Inserts_l32_Playfield, LM_Inserts_l33_Playfield, LM_Inserts_l34_Playfield, LM_Inserts_l35_Playfield, LM_Inserts_l36_Playfield, LM_Inserts_l37_Playfield, LM_Inserts_l38_Playfield, LM_Inserts_l39_Playfield, LM_Inserts_l40_Playfield, LM_Inserts_l41_Playfield, LM_Inserts_l47_Playfield, LM_Inserts_l48_Playfield, LM_Inserts_l49_Playfield, LM_Inserts_l50_Playfield, LM_Inserts_l51_Playfield, LM_Inserts_l52_Playfield, LM_Inserts_l54_Playfield, LM_Inserts_l55_Playfield, LM_Inserts_l56_Playfield, LM_Inserts_l57_Playfield, LM_Inserts_l58_Playfield, LM_Inserts_l59_Playfield, LM_Inserts_l6_Playfield, LM_Inserts_l7_Playfield, LM_Inserts_l8_Playfield, LM_Inserts_l9_Playfield)
Dim BP_RPlastics: BP_RPlastics=Array(BM_RPlastics, LM_Flashers_f109_RPlastics, LM_Flashers_f126_RPlastics, LM_Flashers_f127_RPlastics, LM_Flashers_f128_RPlastics, LM_Flashers_f129a_RPlastics, LM_GISplit_gi08_RPlastics, LM_GI_RPlastics)
Dim BP_RRmp: BP_RRmp=Array(BM_RRmp, LM_Flashers_f109_RRmp, LM_Flashers_f114_RRmp, LM_Flashers_f125_RRmp, LM_Flashers_f126_RRmp, LM_Flashers_f127_RRmp, LM_Flashers_f128_RRmp, LM_Flashers_f129a_RRmp, LM_GISplit_gi08_RRmp, LM_GI_RRmp, LM_Inserts_l15_RRmp, LM_Inserts_l34_RRmp, LM_Inserts_l37_RRmp, LM_Inserts_l38_RRmp, LM_Inserts_l40_RRmp, LM_Inserts_l47_RRmp, LM_Inserts_l48_RRmp, LM_Inserts_l54_RRmp, LM_Inserts_l55_RRmp)
Dim BP_Rails: BP_Rails=Array(BM_Rails, LM_Flashers_f113_Rails, LM_Flashers_f125_Rails, LM_Flashers_f126_Rails, LM_Flashers_f127_Rails, LM_Flashers_f129a_Rails, LM_Flashers_f130_Rails, LM_GI_Rails, LM_Inserts_l47_Rails, LM_Inserts_l48_Rails)
Dim BP_Rightsling1: BP_Rightsling1=Array(BM_Rightsling1, LM_GISplit_gi02_Rightsling1, LM_GI_Rightsling1)
Dim BP_Rightsling2: BP_Rightsling2=Array(BM_Rightsling2, LM_GISplit_gi02_Rightsling2, LM_GI_Rightsling2)
Dim BP_Rightsling3: BP_Rightsling3=Array(BM_Rightsling3, LM_Flashers_f127_Rightsling3, LM_GISplit_gi02_Rightsling3, LM_GI_Rightsling3)
Dim BP_Rightsling4: BP_Rightsling4=Array(BM_Rightsling4, LM_Flashers_f125_Rightsling4, LM_Flashers_f127_Rightsling4, LM_GISplit_gi02_Rightsling4, LM_GI_Rightsling4)
Dim BP_RmpCovers: BP_RmpCovers=Array(BM_RmpCovers, LM_Flashers_f109_RmpCovers, LM_Flashers_f126_RmpCovers, LM_Flashers_f127_RmpCovers, LM_Flashers_f129a_RmpCovers, LM_GI_RmpCovers, LM_Inserts_l37_RmpCovers, LM_Inserts_l40_RmpCovers, LM_Inserts_l47_RmpCovers, LM_Inserts_l48_RmpCovers)
Dim BP_Sideblades: BP_Sideblades=Array(BM_Sideblades, LM_Flashers_f109_Sideblades, LM_Flashers_f113_Sideblades, LM_Flashers_f125_Sideblades, LM_Flashers_f126_Sideblades, LM_Flashers_f128_Sideblades, LM_GI_Sideblades, LM_Inserts_l42_Sideblades, LM_Inserts_l43_Sideblades, LM_Inserts_l44_Sideblades, LM_Inserts_l45_Sideblades, LM_Inserts_l46_Sideblades, LM_Inserts_l47_Sideblades, LM_Inserts_l48_Sideblades, LM_Inserts_l50_Sideblades)
Dim BP_Slings: BP_Slings=Array(BM_Slings, LM_Flashers_f109_Slings, LM_Flashers_f125_Slings, LM_Flashers_f126_Slings, LM_Flashers_f127_Slings, LM_Flashers_f129a_Slings, LM_GISplit_gi02_Slings, LM_GISplit_gi05_Slings, LM_GISplit_gi08_Slings, LM_GI_Slings, LM_Inserts_l51_Slings)
Dim BP_SpecialPost: BP_SpecialPost=Array(BM_SpecialPost, LM_Flashers_f125_SpecialPost, LM_Flashers_f127_SpecialPost, LM_GI_SpecialPost)
Dim BP_Spinner: BP_Spinner=Array(BM_Spinner, LM_Flashers_f113_Spinner, LM_Flashers_f127_Spinner, LM_Flashers_f128_Spinner, LM_GI_Spinner)
Dim BP_StickersCRamp: BP_StickersCRamp=Array(BM_StickersCRamp, LM_Flashers_f113_StickersCRamp, LM_Flashers_f114a_StickersCRamp, LM_Flashers_f127_StickersCRamp, LM_Flashers_f128_StickersCRamp, LM_Flashers_f129_StickersCRamp, LM_Flashers_f130_StickersCRamp, LM_GISplit_gi08_StickersCRamp, LM_GI_StickersCRamp, LM_GISplit_gi01_StickersCRamp, LM_Inserts_l15_StickersCRamp, LM_Inserts_l16_StickersCRamp, LM_Inserts_l23_StickersCRamp, LM_Inserts_l25_StickersCRamp, LM_Inserts_l26_StickersCRamp, LM_Inserts_l27_StickersCRamp, LM_Inserts_l28_StickersCRamp, LM_Inserts_l29_StickersCRamp, LM_Inserts_l30_StickersCRamp, LM_Inserts_l31_StickersCRamp, LM_Inserts_l34_StickersCRamp, LM_Inserts_l35_StickersCRamp, LM_Inserts_l36_StickersCRamp, LM_Inserts_l41_StickersCRamp, LM_Inserts_l42_StickersCRamp, LM_Inserts_l43_StickersCRamp, LM_Inserts_l44_StickersCRamp, LM_Inserts_l45_StickersCRamp, LM_Inserts_l46_StickersCRamp, LM_Inserts_l47_StickersCRamp, LM_Inserts_l56_StickersCRamp, LM_Inserts_l6_StickersCRamp, _
  LM_Inserts_l7_StickersCRamp, LM_Inserts_l8_StickersCRamp)
Dim BP_TargetL1: BP_TargetL1=Array(BM_TargetL1, LM_Flashers_f114a_TargetL1, LM_Flashers_f125_TargetL1, LM_Flashers_f130_TargetL1, LM_GI_TargetL1, LM_Inserts_l33_TargetL1, LM_Inserts_l58_TargetL1)
Dim BP_TargetL2: BP_TargetL2=Array(BM_TargetL2, LM_Flashers_f114a_TargetL2, LM_Flashers_f125_TargetL2, LM_Flashers_f130_TargetL2, LM_GI_TargetL2, LM_Inserts_l33_TargetL2, LM_Inserts_l57_TargetL2, LM_Inserts_l58_TargetL2)
Dim BP_TargetL3: BP_TargetL3=Array(BM_TargetL3, LM_Flashers_f114a_TargetL3, LM_Flashers_f125_TargetL3, LM_Flashers_f130_TargetL3, LM_GI_TargetL3, LM_Inserts_l33_TargetL3, LM_Inserts_l48_TargetL3, LM_Inserts_l57_TargetL3)
Dim BP_TargetM1: BP_TargetM1=Array(BM_TargetM1, LM_Flashers_f127_TargetM1, LM_Flashers_f128_TargetM1, LM_Flashers_f129_TargetM1, LM_GI_TargetM1, LM_Inserts_l35_TargetM1)
Dim BP_TargetM2: BP_TargetM2=Array(BM_TargetM2, LM_Flashers_f127_TargetM2, LM_Flashers_f128_TargetM2, LM_GI_TargetM2)
Dim BP_TargetM3: BP_TargetM3=Array(BM_TargetM3, LM_Flashers_f127_TargetM3, LM_Flashers_f128_TargetM3, LM_Flashers_f130_TargetM3, LM_GI_TargetM3)
Dim BP_TargetR1: BP_TargetR1=Array(BM_TargetR1, LM_Flashers_f127_TargetR1, LM_Flashers_f128_TargetR1, LM_Flashers_f129_TargetR1, LM_Flashers_f130_TargetR1, LM_GI_TargetR1)
Dim BP_TargetR2: BP_TargetR2=Array(BM_TargetR2, LM_Flashers_f127_TargetR2, LM_Flashers_f128_TargetR2, LM_Flashers_f129_TargetR2, LM_Flashers_f130_TargetR2, LM_GI_TargetR2, LM_Inserts_l36_TargetR2, LM_Inserts_l41_TargetR2)
Dim BP_TargetR3: BP_TargetR3=Array(BM_TargetR3, LM_Flashers_f128_TargetR3, LM_Flashers_f129_TargetR3, LM_Flashers_f130_TargetR3, LM_GI_TargetR3, LM_Inserts_l36_TargetR3)
Dim BP_Target_sw35: BP_Target_sw35=Array(BM_Target_sw35, LM_Flashers_f114_Target_sw35, LM_Flashers_f126_Target_sw35, LM_GISplit_gi08_Target_sw35, LM_GI_Target_sw35, LM_Inserts_l37_Target_sw35)
Dim BP_WireRamp: BP_WireRamp=Array(BM_WireRamp, LM_Flashers_f109_WireRamp, LM_Flashers_f114_WireRamp, LM_Flashers_f114a_WireRamp, LM_Flashers_f125_WireRamp, LM_Flashers_f126_WireRamp, LM_Flashers_f127_WireRamp, LM_Flashers_f129_WireRamp, LM_Flashers_f130_WireRamp, LM_GISplit_gi02_WireRamp, LM_GISplit_gi05_WireRamp, LM_GISplit_gi08_WireRamp, LM_GI_WireRamp, LM_GISplit_gi01_WireRamp, LM_Inserts_l15_WireRamp, LM_Inserts_l17_WireRamp, LM_Inserts_l18_WireRamp, LM_Inserts_l19_WireRamp, LM_Inserts_l23_WireRamp, LM_Inserts_l33_WireRamp, LM_Inserts_l34_WireRamp, LM_Inserts_l35_WireRamp, LM_Inserts_l37_WireRamp, LM_Inserts_l38_WireRamp, LM_Inserts_l39_WireRamp, LM_Inserts_l40_WireRamp, LM_Inserts_l47_WireRamp, LM_Inserts_l48_WireRamp, LM_Inserts_l50_WireRamp, LM_Inserts_l51_WireRamp, LM_Inserts_l54_WireRamp, LM_Inserts_l55_WireRamp, LM_Inserts_l56_WireRamp, LM_Inserts_l57_WireRamp, LM_Inserts_l58_WireRamp, LM_Inserts_l59_WireRamp)
Dim BP_sw14Rollover: BP_sw14Rollover=Array(BM_sw14Rollover)
Dim BP_sw49Rollover: BP_sw49Rollover=Array(BM_sw49Rollover, LM_GISplit_gi02_sw49Rollover, LM_GI_sw49Rollover)
Dim BP_sw50Rollover: BP_sw50Rollover=Array(BM_sw50Rollover, LM_GISplit_gi05_sw50Rollover, LM_GI_sw50Rollover)
Dim BP_sw51Rollover: BP_sw51Rollover=Array(BM_sw51Rollover, LM_GISplit_gi05_sw51Rollover, LM_GI_sw51Rollover)
Dim BP_sw52Rollover: BP_sw52Rollover=Array(BM_sw52Rollover, LM_GISplit_gi02_sw52Rollover, LM_GI_sw52Rollover)
Dim BP_sw54Rollover: BP_sw54Rollover=Array(BM_sw54Rollover, LM_Flashers_f114_sw54Rollover, LM_Flashers_f126_sw54Rollover, LM_Flashers_f127_sw54Rollover, LM_GISplit_gi08_sw54Rollover, LM_GI_sw54Rollover)
Dim BP_sw55Rollover: BP_sw55Rollover=Array(BM_sw55Rollover, LM_Flashers_f126_sw55Rollover, LM_Flashers_f127_sw55Rollover, LM_GISplit_gi08_sw55Rollover, LM_GI_sw55Rollover)
Dim BP_sw56Rollover: BP_sw56Rollover=Array(BM_sw56Rollover, LM_Flashers_f126_sw56Rollover, LM_Flashers_f127_sw56Rollover, LM_GISplit_gi08_sw56Rollover, LM_GI_sw56Rollover)
' Arrays per lighting scenario
Dim BL_Flashers_f109: BL_Flashers_f109=Array(LM_Flashers_f109_CRampEdges, LM_Flashers_f109_DiverterArm, LM_Flashers_f109_GateFlap, LM_Flashers_f109_Parts, LM_Flashers_f109_Playfield, LM_Flashers_f109_RPlastics, LM_Flashers_f109_RRmp, LM_Flashers_f109_RmpCovers, LM_Flashers_f109_Sideblades, LM_Flashers_f109_Slings, LM_Flashers_f109_WireRamp)
Dim BL_Flashers_f113: BL_Flashers_f113=Array(LM_Flashers_f113_CRampEdges, LM_Flashers_f113_Gate05Wire, LM_Flashers_f113_Parts, LM_Flashers_f113_Playfield, LM_Flashers_f113_Rails, LM_Flashers_f113_Sideblades, LM_Flashers_f113_Spinner, LM_Flashers_f113_StickersCRamp)
Dim BL_Flashers_f114: BL_Flashers_f114=Array(LM_Flashers_f114_Gate02Wire, LM_Flashers_f114_Parts, LM_Flashers_f114_Playfield, LM_Flashers_f114_RRmp, LM_Flashers_f114_Target_sw35, LM_Flashers_f114_WireRamp, LM_Flashers_f114_sw54Rollover)
Dim BL_Flashers_f114a: BL_Flashers_f114a=Array(LM_Flashers_f114a_Parts, LM_Flashers_f114a_Playfield, LM_Flashers_f114a_StickersCRamp, LM_Flashers_f114a_TargetL1, LM_Flashers_f114a_TargetL2, LM_Flashers_f114a_TargetL3, LM_Flashers_f114a_WireRamp)
Dim BL_Flashers_f125: BL_Flashers_f125=Array(LM_Flashers_f125_Bumper3Ring, LM_Flashers_f125_CRampEdges, LM_Flashers_f125_LeftSling1, LM_Flashers_f125_LeftSling2, LM_Flashers_f125_LeftSling3, LM_Flashers_f125_LeftSling4, LM_Flashers_f125_Parts, LM_Flashers_f125_Playfield, LM_Flashers_f125_RRmp, LM_Flashers_f125_Rails, LM_Flashers_f125_Rightsling4, LM_Flashers_f125_Sideblades, LM_Flashers_f125_Slings, LM_Flashers_f125_SpecialPost, LM_Flashers_f125_TargetL1, LM_Flashers_f125_TargetL2, LM_Flashers_f125_TargetL3, LM_Flashers_f125_WireRamp)
Dim BL_Flashers_f126: BL_Flashers_f126=Array(LM_Flashers_f126_Bumper1Ring, LM_Flashers_f126_Bumper1Socket, LM_Flashers_f126_Bumper2Ring, LM_Flashers_f126_Bumper3Ring, LM_Flashers_f126_BumperInner, LM_Flashers_f126_BumperOuter, LM_Flashers_f126_CRamp, LM_Flashers_f126_CRampEdges, LM_Flashers_f126_Car, LM_Flashers_f126_Gate02Wire, LM_Flashers_f126_GateFlap, LM_Flashers_f126_Parts, LM_Flashers_f126_Playfield, LM_Flashers_f126_RPlastics, LM_Flashers_f126_RRmp, LM_Flashers_f126_Rails, LM_Flashers_f126_RmpCovers, LM_Flashers_f126_Sideblades, LM_Flashers_f126_Slings, LM_Flashers_f126_Target_sw35, LM_Flashers_f126_WireRamp, LM_Flashers_f126_sw54Rollover, LM_Flashers_f126_sw55Rollover, LM_Flashers_f126_sw56Rollover)
Dim BL_Flashers_f127: BL_Flashers_f127=Array(LM_Flashers_f127_Bumper1Ring, LM_Flashers_f127_Bumper2Ring, LM_Flashers_f127_Bumper3Ring, LM_Flashers_f127_Bumper3Socket, LM_Flashers_f127_BumperInner, LM_Flashers_f127_BumperOuter, LM_Flashers_f127_CRamp, LM_Flashers_f127_CRampEdges, LM_Flashers_f127_Car, LM_Flashers_f127_GateSw46Wire, LM_Flashers_f127_LeftSling4, LM_Flashers_f127_Parts, LM_Flashers_f127_Playfield, LM_Flashers_f127_RPlastics, LM_Flashers_f127_RRmp, LM_Flashers_f127_Rails, LM_Flashers_f127_Rightsling3, LM_Flashers_f127_Rightsling4, LM_Flashers_f127_RmpCovers, LM_Flashers_f127_Slings, LM_Flashers_f127_SpecialPost, LM_Flashers_f127_Spinner, LM_Flashers_f127_StickersCRamp, LM_Flashers_f127_TargetM1, LM_Flashers_f127_TargetM2, LM_Flashers_f127_TargetM3, LM_Flashers_f127_TargetR1, LM_Flashers_f127_TargetR2, LM_Flashers_f127_WireRamp, LM_Flashers_f127_sw54Rollover, LM_Flashers_f127_sw55Rollover, LM_Flashers_f127_sw56Rollover)
Dim BL_Flashers_f128: BL_Flashers_f128=Array(LM_Flashers_f128_Bumper3Ring, LM_Flashers_f128_BumperInner, LM_Flashers_f128_BumperOuter, LM_Flashers_f128_CRampEdges, LM_Flashers_f128_Parts, LM_Flashers_f128_Playfield, LM_Flashers_f128_RPlastics, LM_Flashers_f128_RRmp, LM_Flashers_f128_Sideblades, LM_Flashers_f128_Spinner, LM_Flashers_f128_StickersCRamp, LM_Flashers_f128_TargetM1, LM_Flashers_f128_TargetM2, LM_Flashers_f128_TargetM3, LM_Flashers_f128_TargetR1, LM_Flashers_f128_TargetR2, LM_Flashers_f128_TargetR3)
Dim BL_Flashers_f129: BL_Flashers_f129=Array(LM_Flashers_f129_Bumper3Ring, LM_Flashers_f129_CRampEdges, LM_Flashers_f129_Parts, LM_Flashers_f129_Playfield, LM_Flashers_f129_StickersCRamp, LM_Flashers_f129_TargetM1, LM_Flashers_f129_TargetR1, LM_Flashers_f129_TargetR2, LM_Flashers_f129_TargetR3, LM_Flashers_f129_WireRamp)
Dim BL_Flashers_f129a: BL_Flashers_f129a=Array(LM_Flashers_f129a_Parts, LM_Flashers_f129a_Playfield, LM_Flashers_f129a_RPlastics, LM_Flashers_f129a_RRmp, LM_Flashers_f129a_Rails, LM_Flashers_f129a_RmpCovers, LM_Flashers_f129a_Slings)
Dim BL_Flashers_f130: BL_Flashers_f130=Array(LM_Flashers_f130_CRampEdges, LM_Flashers_f130_GateSw45Wire, LM_Flashers_f130_Parts, LM_Flashers_f130_Playfield, LM_Flashers_f130_Rails, LM_Flashers_f130_StickersCRamp, LM_Flashers_f130_TargetL1, LM_Flashers_f130_TargetL2, LM_Flashers_f130_TargetL3, LM_Flashers_f130_TargetM3, LM_Flashers_f130_TargetR1, LM_Flashers_f130_TargetR2, LM_Flashers_f130_TargetR3, LM_Flashers_f130_WireRamp)
Dim BL_GI: BL_GI=Array(LM_GI_Bumper1Ring, LM_GI_Bumper1Socket, LM_GI_Bumper2Ring, LM_GI_Bumper2Socket, LM_GI_Bumper3Ring, LM_GI_Bumper3Socket, LM_GI_BumperInner, LM_GI_BumperOuter, LM_GI_CRamp, LM_GI_CRampEdges, LM_GI_Car, LM_GI_FlipperL, LM_GI_FlipperR, LM_GI_Gate02Wire, LM_GI_Gate04Wire, LM_GI_Gate05Wire, LM_GI_GateFlap, LM_GI_GateSw44Wire, LM_GI_GateSw45Wire, LM_GI_GateSw46Wire, LM_GI_LeftSling1, LM_GI_LeftSling2, LM_GI_LeftSling3, LM_GI_LeftSling4, LM_GI_Parts, LM_GI_Playfield, LM_GI_RPlastics, LM_GI_RRmp, LM_GI_Rails, LM_GI_Rightsling1, LM_GI_Rightsling2, LM_GI_Rightsling3, LM_GI_Rightsling4, LM_GI_RmpCovers, LM_GI_Sideblades, LM_GI_Slings, LM_GI_SpecialPost, LM_GI_Spinner, LM_GI_StickersCRamp, LM_GI_TargetL1, LM_GI_TargetL2, LM_GI_TargetL3, LM_GI_TargetM1, LM_GI_TargetM2, LM_GI_TargetM3, LM_GI_TargetR1, LM_GI_TargetR2, LM_GI_TargetR3, LM_GI_Target_sw35, LM_GI_WireRamp, LM_GI_sw49Rollover, LM_GI_sw50Rollover, LM_GI_sw51Rollover, LM_GI_sw52Rollover, LM_GI_sw54Rollover, LM_GI_sw55Rollover, _
  LM_GI_sw56Rollover)
Dim BL_GISplit_gi01: BL_GISplit_gi01=Array(LM_GISplit_gi01_Bumper3Ring, LM_GISplit_gi01_Bumper3Socket, LM_GISplit_gi01_BumperInner, LM_GISplit_gi01_BumperOuter, LM_GISplit_gi01_CRampEdges, LM_GISplit_gi01_Parts, LM_GISplit_gi01_Playfield, LM_GISplit_gi01_StickersCRamp, LM_GISplit_gi01_WireRamp)
Dim BL_GISplit_gi02: BL_GISplit_gi02=Array(LM_GISplit_gi02_CRampEdges, LM_GISplit_gi02_FlipperR, LM_GISplit_gi02_Parts, LM_GISplit_gi02_Playfield, LM_GISplit_gi02_Rightsling1, LM_GISplit_gi02_Rightsling2, LM_GISplit_gi02_Rightsling3, LM_GISplit_gi02_Rightsling4, LM_GISplit_gi02_Slings, LM_GISplit_gi02_WireRamp, LM_GISplit_gi02_sw49Rollover, LM_GISplit_gi02_sw52Rollover)
Dim BL_GISplit_gi05: BL_GISplit_gi05=Array(LM_GISplit_gi05_FlipperL, LM_GISplit_gi05_LeftSling1, LM_GISplit_gi05_LeftSling2, LM_GISplit_gi05_LeftSling3, LM_GISplit_gi05_LeftSling4, LM_GISplit_gi05_Parts, LM_GISplit_gi05_Playfield, LM_GISplit_gi05_Slings, LM_GISplit_gi05_WireRamp, LM_GISplit_gi05_sw50Rollover, LM_GISplit_gi05_sw51Rollover)
Dim BL_GISplit_gi08: BL_GISplit_gi08=Array(LM_GISplit_gi08_Bumper1Ring, LM_GISplit_gi08_Bumper1Socket, LM_GISplit_gi08_Bumper2Socket, LM_GISplit_gi08_BumperInner, LM_GISplit_gi08_CRampEdges, LM_GISplit_gi08_Gate02Wire, LM_GISplit_gi08_Gate04Wire, LM_GISplit_gi08_GateFlap, LM_GISplit_gi08_Parts, LM_GISplit_gi08_Playfield, LM_GISplit_gi08_RPlastics, LM_GISplit_gi08_RRmp, LM_GISplit_gi08_Slings, LM_GISplit_gi08_StickersCRamp, LM_GISplit_gi08_Target_sw35, LM_GISplit_gi08_WireRamp, LM_GISplit_gi08_sw54Rollover, LM_GISplit_gi08_sw55Rollover, LM_GISplit_gi08_sw56Rollover)
Dim BL_Inserts_l10: BL_Inserts_l10=Array(LM_Inserts_l10_Playfield)
Dim BL_Inserts_l11: BL_Inserts_l11=Array(LM_Inserts_l11_Playfield)
Dim BL_Inserts_l12: BL_Inserts_l12=Array(LM_Inserts_l12_Playfield)
Dim BL_Inserts_l13: BL_Inserts_l13=Array(LM_Inserts_l13_Playfield)
Dim BL_Inserts_l14: BL_Inserts_l14=Array(LM_Inserts_l14_Playfield)
Dim BL_Inserts_l15: BL_Inserts_l15=Array(LM_Inserts_l15_Parts, LM_Inserts_l15_Playfield, LM_Inserts_l15_RRmp, LM_Inserts_l15_StickersCRamp, LM_Inserts_l15_WireRamp)
Dim BL_Inserts_l16: BL_Inserts_l16=Array(LM_Inserts_l16_CRampEdges, LM_Inserts_l16_Parts, LM_Inserts_l16_Playfield, LM_Inserts_l16_StickersCRamp)
Dim BL_Inserts_l17: BL_Inserts_l17=Array(LM_Inserts_l17_LeftSling4, LM_Inserts_l17_Parts, LM_Inserts_l17_Playfield, LM_Inserts_l17_WireRamp)
Dim BL_Inserts_l18: BL_Inserts_l18=Array(LM_Inserts_l18_Playfield, LM_Inserts_l18_WireRamp)
Dim BL_Inserts_l19: BL_Inserts_l19=Array(LM_Inserts_l19_Playfield, LM_Inserts_l19_WireRamp)
Dim BL_Inserts_l20: BL_Inserts_l20=Array(LM_Inserts_l20_Playfield)
Dim BL_Inserts_l21: BL_Inserts_l21=Array(LM_Inserts_l21_Playfield)
Dim BL_Inserts_l22: BL_Inserts_l22=Array(LM_Inserts_l22_Playfield)
Dim BL_Inserts_l23: BL_Inserts_l23=Array(LM_Inserts_l23_Playfield, LM_Inserts_l23_StickersCRamp, LM_Inserts_l23_WireRamp)
Dim BL_Inserts_l24: BL_Inserts_l24=Array(LM_Inserts_l24_FlipperL, LM_Inserts_l24_FlipperR, LM_Inserts_l24_Playfield)
Dim BL_Inserts_l25: BL_Inserts_l25=Array(LM_Inserts_l25_Playfield, LM_Inserts_l25_StickersCRamp)
Dim BL_Inserts_l26: BL_Inserts_l26=Array(LM_Inserts_l26_Playfield, LM_Inserts_l26_StickersCRamp)
Dim BL_Inserts_l27: BL_Inserts_l27=Array(LM_Inserts_l27_CRamp, LM_Inserts_l27_CRampEdges, LM_Inserts_l27_Playfield, LM_Inserts_l27_StickersCRamp)
Dim BL_Inserts_l28: BL_Inserts_l28=Array(LM_Inserts_l28_CRamp, LM_Inserts_l28_CRampEdges, LM_Inserts_l28_Parts, LM_Inserts_l28_Playfield, LM_Inserts_l28_StickersCRamp)
Dim BL_Inserts_l29: BL_Inserts_l29=Array(LM_Inserts_l29_CRampEdges, LM_Inserts_l29_Parts, LM_Inserts_l29_Playfield, LM_Inserts_l29_StickersCRamp)
Dim BL_Inserts_l30: BL_Inserts_l30=Array(LM_Inserts_l30_Bumper3Ring, LM_Inserts_l30_BumperInner, LM_Inserts_l30_CRampEdges, LM_Inserts_l30_Parts, LM_Inserts_l30_Playfield, LM_Inserts_l30_StickersCRamp)
Dim BL_Inserts_l31: BL_Inserts_l31=Array(LM_Inserts_l31_Playfield, LM_Inserts_l31_StickersCRamp)
Dim BL_Inserts_l32: BL_Inserts_l32=Array(LM_Inserts_l32_Playfield)
Dim BL_Inserts_l33: BL_Inserts_l33=Array(LM_Inserts_l33_Parts, LM_Inserts_l33_Playfield, LM_Inserts_l33_TargetL1, LM_Inserts_l33_TargetL2, LM_Inserts_l33_TargetL3, LM_Inserts_l33_WireRamp)
Dim BL_Inserts_l34: BL_Inserts_l34=Array(LM_Inserts_l34_Bumper1Ring, LM_Inserts_l34_BumperInner, LM_Inserts_l34_BumperOuter, LM_Inserts_l34_Parts, LM_Inserts_l34_Playfield, LM_Inserts_l34_RRmp, LM_Inserts_l34_StickersCRamp, LM_Inserts_l34_WireRamp)
Dim BL_Inserts_l35: BL_Inserts_l35=Array(LM_Inserts_l35_Parts, LM_Inserts_l35_Playfield, LM_Inserts_l35_StickersCRamp, LM_Inserts_l35_TargetM1, LM_Inserts_l35_WireRamp)
Dim BL_Inserts_l36: BL_Inserts_l36=Array(LM_Inserts_l36_CRampEdges, LM_Inserts_l36_Parts, LM_Inserts_l36_Playfield, LM_Inserts_l36_StickersCRamp, LM_Inserts_l36_TargetR2, LM_Inserts_l36_TargetR3)
Dim BL_Inserts_l37: BL_Inserts_l37=Array(LM_Inserts_l37_Parts, LM_Inserts_l37_Playfield, LM_Inserts_l37_RRmp, LM_Inserts_l37_RmpCovers, LM_Inserts_l37_Target_sw35, LM_Inserts_l37_WireRamp)
Dim BL_Inserts_l38: BL_Inserts_l38=Array(LM_Inserts_l38_Parts, LM_Inserts_l38_Playfield, LM_Inserts_l38_RRmp, LM_Inserts_l38_WireRamp)
Dim BL_Inserts_l39: BL_Inserts_l39=Array(LM_Inserts_l39_Bumper2Socket, LM_Inserts_l39_Parts, LM_Inserts_l39_Playfield, LM_Inserts_l39_WireRamp)
Dim BL_Inserts_l40: BL_Inserts_l40=Array(LM_Inserts_l40_Parts, LM_Inserts_l40_Playfield, LM_Inserts_l40_RRmp, LM_Inserts_l40_RmpCovers, LM_Inserts_l40_WireRamp)
Dim BL_Inserts_l41: BL_Inserts_l41=Array(LM_Inserts_l41_Playfield, LM_Inserts_l41_StickersCRamp, LM_Inserts_l41_TargetR2)
Dim BL_Inserts_l42: BL_Inserts_l42=Array(LM_Inserts_l42_Parts, LM_Inserts_l42_Sideblades, LM_Inserts_l42_StickersCRamp)
Dim BL_Inserts_l43: BL_Inserts_l43=Array(LM_Inserts_l43_Parts, LM_Inserts_l43_Sideblades, LM_Inserts_l43_StickersCRamp)
Dim BL_Inserts_l44: BL_Inserts_l44=Array(LM_Inserts_l44_Parts, LM_Inserts_l44_Sideblades, LM_Inserts_l44_StickersCRamp)
Dim BL_Inserts_l45: BL_Inserts_l45=Array(LM_Inserts_l45_Parts, LM_Inserts_l45_Sideblades, LM_Inserts_l45_StickersCRamp)
Dim BL_Inserts_l46: BL_Inserts_l46=Array(LM_Inserts_l46_Parts, LM_Inserts_l46_Sideblades, LM_Inserts_l46_StickersCRamp)
Dim BL_Inserts_l47: BL_Inserts_l47=Array(LM_Inserts_l47_BumperInner, LM_Inserts_l47_BumperOuter, LM_Inserts_l47_Car, LM_Inserts_l47_Parts, LM_Inserts_l47_Playfield, LM_Inserts_l47_RRmp, LM_Inserts_l47_Rails, LM_Inserts_l47_RmpCovers, LM_Inserts_l47_Sideblades, LM_Inserts_l47_StickersCRamp, LM_Inserts_l47_WireRamp)
Dim BL_Inserts_l48: BL_Inserts_l48=Array(LM_Inserts_l48_BumperInner, LM_Inserts_l48_BumperOuter, LM_Inserts_l48_Car, LM_Inserts_l48_Parts, LM_Inserts_l48_Playfield, LM_Inserts_l48_RRmp, LM_Inserts_l48_Rails, LM_Inserts_l48_RmpCovers, LM_Inserts_l48_Sideblades, LM_Inserts_l48_TargetL3, LM_Inserts_l48_WireRamp)
Dim BL_Inserts_l49: BL_Inserts_l49=Array(LM_Inserts_l49_CRampEdges, LM_Inserts_l49_Parts, LM_Inserts_l49_Playfield)
Dim BL_Inserts_l50: BL_Inserts_l50=Array(LM_Inserts_l50_Playfield, LM_Inserts_l50_Sideblades, LM_Inserts_l50_WireRamp)
Dim BL_Inserts_l51: BL_Inserts_l51=Array(LM_Inserts_l51_LeftSling4, LM_Inserts_l51_Playfield, LM_Inserts_l51_Slings, LM_Inserts_l51_WireRamp)
Dim BL_Inserts_l52: BL_Inserts_l52=Array(LM_Inserts_l52_CRampEdges, LM_Inserts_l52_Parts, LM_Inserts_l52_Playfield)
Dim BL_Inserts_l54: BL_Inserts_l54=Array(LM_Inserts_l54_Parts, LM_Inserts_l54_Playfield, LM_Inserts_l54_RRmp, LM_Inserts_l54_WireRamp)
Dim BL_Inserts_l55: BL_Inserts_l55=Array(LM_Inserts_l55_Parts, LM_Inserts_l55_Playfield, LM_Inserts_l55_RRmp, LM_Inserts_l55_WireRamp)
Dim BL_Inserts_l56: BL_Inserts_l56=Array(LM_Inserts_l56_GateFlap, LM_Inserts_l56_Parts, LM_Inserts_l56_Playfield, LM_Inserts_l56_StickersCRamp, LM_Inserts_l56_WireRamp)
Dim BL_Inserts_l57: BL_Inserts_l57=Array(LM_Inserts_l57_Parts, LM_Inserts_l57_Playfield, LM_Inserts_l57_TargetL2, LM_Inserts_l57_TargetL3, LM_Inserts_l57_WireRamp)
Dim BL_Inserts_l58: BL_Inserts_l58=Array(LM_Inserts_l58_Parts, LM_Inserts_l58_Playfield, LM_Inserts_l58_TargetL1, LM_Inserts_l58_TargetL2, LM_Inserts_l58_WireRamp)
Dim BL_Inserts_l59: BL_Inserts_l59=Array(LM_Inserts_l59_Playfield, LM_Inserts_l59_WireRamp)
Dim BL_Inserts_l6: BL_Inserts_l6=Array(LM_Inserts_l6_CRamp, LM_Inserts_l6_Playfield, LM_Inserts_l6_StickersCRamp)
Dim BL_Inserts_l7: BL_Inserts_l7=Array(LM_Inserts_l7_Playfield, LM_Inserts_l7_StickersCRamp)
Dim BL_Inserts_l8: BL_Inserts_l8=Array(LM_Inserts_l8_Parts, LM_Inserts_l8_Playfield, LM_Inserts_l8_StickersCRamp)
Dim BL_Inserts_l9: BL_Inserts_l9=Array(LM_Inserts_l9_Playfield)
Dim BL_Room: BL_Room=Array(BM_Bumper1Ring, BM_Bumper1Socket, BM_Bumper2Ring, BM_Bumper2Socket, BM_Bumper3Ring, BM_Bumper3Socket, BM_BumperInner, BM_BumperOuter, BM_CRamp, BM_CRampEdges, BM_Car, BM_DiverterArm, BM_FlipperL, BM_FlipperR, BM_Gate003_Wire, BM_Gate02Wire, BM_Gate04Wire, BM_Gate05Wire, BM_GateFlap, BM_GateSw44Wire, BM_GateSw45Wire, BM_GateSw46Wire, BM_LeftSling1, BM_LeftSling2, BM_LeftSling3, BM_LeftSling4, BM_Parts, BM_Playfield, BM_RPlastics, BM_RRmp, BM_Rails, BM_Rightsling1, BM_Rightsling2, BM_Rightsling3, BM_Rightsling4, BM_RmpCovers, BM_Sideblades, BM_Slings, BM_SpecialPost, BM_Spinner, BM_StickersCRamp, BM_TargetL1, BM_TargetL2, BM_TargetL3, BM_TargetM1, BM_TargetM2, BM_TargetM3, BM_TargetR1, BM_TargetR2, BM_TargetR3, BM_Target_sw35, BM_WireRamp, BM_sw14Rollover, BM_sw49Rollover, BM_sw50Rollover, BM_sw51Rollover, BM_sw52Rollover, BM_sw54Rollover, BM_sw55Rollover, BM_sw56Rollover)
' Global arrays
Dim BG_Bakemap: BG_Bakemap=Array(BM_Bumper1Ring, BM_Bumper1Socket, BM_Bumper2Ring, BM_Bumper2Socket, BM_Bumper3Ring, BM_Bumper3Socket, BM_BumperInner, BM_BumperOuter, BM_CRamp, BM_CRampEdges, BM_Car, BM_DiverterArm, BM_FlipperL, BM_FlipperR, BM_Gate003_Wire, BM_Gate02Wire, BM_Gate04Wire, BM_Gate05Wire, BM_GateFlap, BM_GateSw44Wire, BM_GateSw45Wire, BM_GateSw46Wire, BM_LeftSling1, BM_LeftSling2, BM_LeftSling3, BM_LeftSling4, BM_Parts, BM_Playfield, BM_RPlastics, BM_RRmp, BM_Rails, BM_Rightsling1, BM_Rightsling2, BM_Rightsling3, BM_Rightsling4, BM_RmpCovers, BM_Sideblades, BM_Slings, BM_SpecialPost, BM_Spinner, BM_StickersCRamp, BM_TargetL1, BM_TargetL2, BM_TargetL3, BM_TargetM1, BM_TargetM2, BM_TargetM3, BM_TargetR1, BM_TargetR2, BM_TargetR3, BM_Target_sw35, BM_WireRamp, BM_sw14Rollover, BM_sw49Rollover, BM_sw50Rollover, BM_sw51Rollover, BM_sw52Rollover, BM_sw54Rollover, BM_sw55Rollover, BM_sw56Rollover)
Dim BG_Lightmap: BG_Lightmap=Array(LM_Flashers_f109_CRampEdges, LM_Flashers_f109_DiverterArm, LM_Flashers_f109_GateFlap, LM_Flashers_f109_Parts, LM_Flashers_f109_Playfield, LM_Flashers_f109_RPlastics, LM_Flashers_f109_RRmp, LM_Flashers_f109_RmpCovers, LM_Flashers_f109_Sideblades, LM_Flashers_f109_Slings, LM_Flashers_f109_WireRamp, LM_Flashers_f113_CRampEdges, LM_Flashers_f113_Gate05Wire, LM_Flashers_f113_Parts, LM_Flashers_f113_Playfield, LM_Flashers_f113_Rails, LM_Flashers_f113_Sideblades, LM_Flashers_f113_Spinner, LM_Flashers_f113_StickersCRamp, LM_Flashers_f114_Gate02Wire, LM_Flashers_f114_Parts, LM_Flashers_f114_Playfield, LM_Flashers_f114_RRmp, LM_Flashers_f114_Target_sw35, LM_Flashers_f114_WireRamp, LM_Flashers_f114_sw54Rollover, LM_Flashers_f114a_Parts, LM_Flashers_f114a_Playfield, LM_Flashers_f114a_StickersCRamp, LM_Flashers_f114a_TargetL1, LM_Flashers_f114a_TargetL2, LM_Flashers_f114a_TargetL3, LM_Flashers_f114a_WireRamp, LM_Flashers_f125_Bumper3Ring, LM_Flashers_f125_CRampEdges, _
  LM_Flashers_f125_LeftSling1, LM_Flashers_f125_LeftSling2, LM_Flashers_f125_LeftSling3, LM_Flashers_f125_LeftSling4, LM_Flashers_f125_Parts, LM_Flashers_f125_Playfield, LM_Flashers_f125_RRmp, LM_Flashers_f125_Rails, LM_Flashers_f125_Rightsling4, LM_Flashers_f125_Sideblades, LM_Flashers_f125_Slings, LM_Flashers_f125_SpecialPost, LM_Flashers_f125_TargetL1, LM_Flashers_f125_TargetL2, LM_Flashers_f125_TargetL3, LM_Flashers_f125_WireRamp, LM_Flashers_f126_Bumper1Ring, LM_Flashers_f126_Bumper1Socket, LM_Flashers_f126_Bumper2Ring, LM_Flashers_f126_Bumper3Ring, LM_Flashers_f126_BumperInner, LM_Flashers_f126_BumperOuter, LM_Flashers_f126_CRamp, LM_Flashers_f126_CRampEdges, LM_Flashers_f126_Car, LM_Flashers_f126_Gate02Wire, LM_Flashers_f126_GateFlap, LM_Flashers_f126_Parts, LM_Flashers_f126_Playfield, LM_Flashers_f126_RPlastics, LM_Flashers_f126_RRmp, LM_Flashers_f126_Rails, LM_Flashers_f126_RmpCovers, LM_Flashers_f126_Sideblades, LM_Flashers_f126_Slings, LM_Flashers_f126_Target_sw35, LM_Flashers_f126_WireRamp, _
  LM_Flashers_f126_sw54Rollover, LM_Flashers_f126_sw55Rollover, LM_Flashers_f126_sw56Rollover, LM_Flashers_f127_Bumper1Ring, LM_Flashers_f127_Bumper2Ring, LM_Flashers_f127_Bumper3Ring, LM_Flashers_f127_Bumper3Socket, LM_Flashers_f127_BumperInner, LM_Flashers_f127_BumperOuter, LM_Flashers_f127_CRamp, LM_Flashers_f127_CRampEdges, LM_Flashers_f127_Car, LM_Flashers_f127_GateSw46Wire, LM_Flashers_f127_LeftSling4, LM_Flashers_f127_Parts, LM_Flashers_f127_Playfield, LM_Flashers_f127_RPlastics, LM_Flashers_f127_RRmp, LM_Flashers_f127_Rails, LM_Flashers_f127_Rightsling3, LM_Flashers_f127_Rightsling4, LM_Flashers_f127_RmpCovers, LM_Flashers_f127_Slings, LM_Flashers_f127_SpecialPost, LM_Flashers_f127_Spinner, LM_Flashers_f127_StickersCRamp, LM_Flashers_f127_TargetM1, LM_Flashers_f127_TargetM2, LM_Flashers_f127_TargetM3, LM_Flashers_f127_TargetR1, LM_Flashers_f127_TargetR2, LM_Flashers_f127_WireRamp, LM_Flashers_f127_sw54Rollover, LM_Flashers_f127_sw55Rollover, LM_Flashers_f127_sw56Rollover, LM_Flashers_f128_Bumper3Ring, _
  LM_Flashers_f128_BumperInner, LM_Flashers_f128_BumperOuter, LM_Flashers_f128_CRampEdges, LM_Flashers_f128_Parts, LM_Flashers_f128_Playfield, LM_Flashers_f128_RPlastics, LM_Flashers_f128_RRmp, LM_Flashers_f128_Sideblades, LM_Flashers_f128_Spinner, LM_Flashers_f128_StickersCRamp, LM_Flashers_f128_TargetM1, LM_Flashers_f128_TargetM2, LM_Flashers_f128_TargetM3, LM_Flashers_f128_TargetR1, LM_Flashers_f128_TargetR2, LM_Flashers_f128_TargetR3, LM_Flashers_f129_Bumper3Ring, LM_Flashers_f129_CRampEdges, LM_Flashers_f129_Parts, LM_Flashers_f129_Playfield, LM_Flashers_f129_StickersCRamp, LM_Flashers_f129_TargetM1, LM_Flashers_f129_TargetR1, LM_Flashers_f129_TargetR2, LM_Flashers_f129_TargetR3, LM_Flashers_f129_WireRamp, LM_Flashers_f129a_Parts, LM_Flashers_f129a_Playfield, LM_Flashers_f129a_RPlastics, LM_Flashers_f129a_RRmp, LM_Flashers_f129a_Rails, LM_Flashers_f129a_RmpCovers, LM_Flashers_f129a_Slings, LM_Flashers_f130_CRampEdges, LM_Flashers_f130_GateSw45Wire, LM_Flashers_f130_Parts, LM_Flashers_f130_Playfield, _
  LM_Flashers_f130_Rails, LM_Flashers_f130_StickersCRamp, LM_Flashers_f130_TargetL1, LM_Flashers_f130_TargetL2, LM_Flashers_f130_TargetL3, LM_Flashers_f130_TargetM3, LM_Flashers_f130_TargetR1, LM_Flashers_f130_TargetR2, LM_Flashers_f130_TargetR3, LM_Flashers_f130_WireRamp, LM_GI_Bumper1Ring, LM_GI_Bumper1Socket, LM_GI_Bumper2Ring, LM_GI_Bumper2Socket, LM_GI_Bumper3Ring, LM_GI_Bumper3Socket, LM_GI_BumperInner, LM_GI_BumperOuter, LM_GI_CRamp, LM_GI_CRampEdges, LM_GI_Car, LM_GI_FlipperL, LM_GI_FlipperR, LM_GI_Gate02Wire, LM_GI_Gate04Wire, LM_GI_Gate05Wire, LM_GI_GateFlap, LM_GI_GateSw44Wire, LM_GI_GateSw45Wire, LM_GI_GateSw46Wire, LM_GI_LeftSling1, LM_GI_LeftSling2, LM_GI_LeftSling3, LM_GI_LeftSling4, LM_GI_Parts, LM_GI_Playfield, LM_GI_RPlastics, LM_GI_RRmp, LM_GI_Rails, LM_GI_Rightsling1, LM_GI_Rightsling2, LM_GI_Rightsling3, LM_GI_Rightsling4, LM_GI_RmpCovers, LM_GI_Sideblades, LM_GI_Slings, LM_GI_SpecialPost, LM_GI_Spinner, LM_GI_StickersCRamp, LM_GI_TargetL1, LM_GI_TargetL2, LM_GI_TargetL3, LM_GI_TargetM1, _
  LM_GI_TargetM2, LM_GI_TargetM3, LM_GI_TargetR1, LM_GI_TargetR2, LM_GI_TargetR3, LM_GI_Target_sw35, LM_GI_WireRamp, LM_GI_sw49Rollover, LM_GI_sw50Rollover, LM_GI_sw51Rollover, LM_GI_sw52Rollover, LM_GI_sw54Rollover, LM_GI_sw55Rollover, LM_GI_sw56Rollover, LM_GISplit_gi01_Bumper3Ring, LM_GISplit_gi01_Bumper3Socket, LM_GISplit_gi01_BumperInner, LM_GISplit_gi01_BumperOuter, LM_GISplit_gi01_CRampEdges, LM_GISplit_gi01_Parts, LM_GISplit_gi01_Playfield, LM_GISplit_gi01_StickersCRamp, LM_GISplit_gi01_WireRamp, LM_GISplit_gi02_CRampEdges, LM_GISplit_gi02_FlipperR, LM_GISplit_gi02_Parts, LM_GISplit_gi02_Playfield, LM_GISplit_gi02_Rightsling1, LM_GISplit_gi02_Rightsling2, LM_GISplit_gi02_Rightsling3, LM_GISplit_gi02_Rightsling4, LM_GISplit_gi02_Slings, LM_GISplit_gi02_WireRamp, LM_GISplit_gi02_sw49Rollover, LM_GISplit_gi02_sw52Rollover, LM_GISplit_gi05_FlipperL, LM_GISplit_gi05_LeftSling1, LM_GISplit_gi05_LeftSling2, LM_GISplit_gi05_LeftSling3, LM_GISplit_gi05_LeftSling4, LM_GISplit_gi05_Parts, _
  LM_GISplit_gi05_Playfield, LM_GISplit_gi05_Slings, LM_GISplit_gi05_WireRamp, LM_GISplit_gi05_sw50Rollover, LM_GISplit_gi05_sw51Rollover, LM_GISplit_gi08_Bumper1Ring, LM_GISplit_gi08_Bumper1Socket, LM_GISplit_gi08_Bumper2Socket, LM_GISplit_gi08_BumperInner, LM_GISplit_gi08_CRampEdges, LM_GISplit_gi08_Gate02Wire, LM_GISplit_gi08_Gate04Wire, LM_GISplit_gi08_GateFlap, LM_GISplit_gi08_Parts, LM_GISplit_gi08_Playfield, LM_GISplit_gi08_RPlastics, LM_GISplit_gi08_RRmp, LM_GISplit_gi08_Slings, LM_GISplit_gi08_StickersCRamp, LM_GISplit_gi08_Target_sw35, LM_GISplit_gi08_WireRamp, LM_GISplit_gi08_sw54Rollover, LM_GISplit_gi08_sw55Rollover, LM_GISplit_gi08_sw56Rollover, LM_Inserts_l10_Playfield, LM_Inserts_l11_Playfield, LM_Inserts_l12_Playfield, LM_Inserts_l13_Playfield, LM_Inserts_l14_Playfield, LM_Inserts_l15_Parts, LM_Inserts_l15_Playfield, LM_Inserts_l15_RRmp, LM_Inserts_l15_StickersCRamp, LM_Inserts_l15_WireRamp, LM_Inserts_l16_CRampEdges, LM_Inserts_l16_Parts, LM_Inserts_l16_Playfield, LM_Inserts_l16_StickersCRamp, _
  LM_Inserts_l17_LeftSling4, LM_Inserts_l17_Parts, LM_Inserts_l17_Playfield, LM_Inserts_l17_WireRamp, LM_Inserts_l18_Playfield, LM_Inserts_l18_WireRamp, LM_Inserts_l19_Playfield, LM_Inserts_l19_WireRamp, LM_Inserts_l20_Playfield, LM_Inserts_l21_Playfield, LM_Inserts_l22_Playfield, LM_Inserts_l23_Playfield, LM_Inserts_l23_StickersCRamp, LM_Inserts_l23_WireRamp, LM_Inserts_l24_FlipperL, LM_Inserts_l24_FlipperR, LM_Inserts_l24_Playfield, LM_Inserts_l25_Playfield, LM_Inserts_l25_StickersCRamp, LM_Inserts_l26_Playfield, LM_Inserts_l26_StickersCRamp, LM_Inserts_l27_CRamp, LM_Inserts_l27_CRampEdges, LM_Inserts_l27_Playfield, LM_Inserts_l27_StickersCRamp, LM_Inserts_l28_CRamp, LM_Inserts_l28_CRampEdges, LM_Inserts_l28_Parts, LM_Inserts_l28_Playfield, LM_Inserts_l28_StickersCRamp, LM_Inserts_l29_CRampEdges, LM_Inserts_l29_Parts, LM_Inserts_l29_Playfield, LM_Inserts_l29_StickersCRamp, LM_Inserts_l30_Bumper3Ring, LM_Inserts_l30_BumperInner, LM_Inserts_l30_CRampEdges, LM_Inserts_l30_Parts, LM_Inserts_l30_Playfield, _
  LM_Inserts_l30_StickersCRamp, LM_Inserts_l31_Playfield, LM_Inserts_l31_StickersCRamp, LM_Inserts_l32_Playfield, LM_Inserts_l33_Parts, LM_Inserts_l33_Playfield, LM_Inserts_l33_TargetL1, LM_Inserts_l33_TargetL2, LM_Inserts_l33_TargetL3, LM_Inserts_l33_WireRamp, LM_Inserts_l34_Bumper1Ring, LM_Inserts_l34_BumperInner, LM_Inserts_l34_BumperOuter, LM_Inserts_l34_Parts, LM_Inserts_l34_Playfield, LM_Inserts_l34_RRmp, LM_Inserts_l34_StickersCRamp, LM_Inserts_l34_WireRamp, LM_Inserts_l35_Parts, LM_Inserts_l35_Playfield, LM_Inserts_l35_StickersCRamp, LM_Inserts_l35_TargetM1, LM_Inserts_l35_WireRamp, LM_Inserts_l36_CRampEdges, LM_Inserts_l36_Parts, LM_Inserts_l36_Playfield, LM_Inserts_l36_StickersCRamp, LM_Inserts_l36_TargetR2, LM_Inserts_l36_TargetR3, LM_Inserts_l37_Parts, LM_Inserts_l37_Playfield, LM_Inserts_l37_RRmp, LM_Inserts_l37_RmpCovers, LM_Inserts_l37_Target_sw35, LM_Inserts_l37_WireRamp, LM_Inserts_l38_Parts, LM_Inserts_l38_Playfield, LM_Inserts_l38_RRmp, LM_Inserts_l38_WireRamp, LM_Inserts_l39_Bumper2Socket, _
  LM_Inserts_l39_Parts, LM_Inserts_l39_Playfield, LM_Inserts_l39_WireRamp, LM_Inserts_l40_Parts, LM_Inserts_l40_Playfield, LM_Inserts_l40_RRmp, LM_Inserts_l40_RmpCovers, LM_Inserts_l40_WireRamp, LM_Inserts_l41_Playfield, LM_Inserts_l41_StickersCRamp, LM_Inserts_l41_TargetR2, LM_Inserts_l42_Parts, LM_Inserts_l42_Sideblades, LM_Inserts_l42_StickersCRamp, LM_Inserts_l43_Parts, LM_Inserts_l43_Sideblades, LM_Inserts_l43_StickersCRamp, LM_Inserts_l44_Parts, LM_Inserts_l44_Sideblades, LM_Inserts_l44_StickersCRamp, LM_Inserts_l45_Parts, LM_Inserts_l45_Sideblades, LM_Inserts_l45_StickersCRamp, LM_Inserts_l46_Parts, LM_Inserts_l46_Sideblades, LM_Inserts_l46_StickersCRamp, LM_Inserts_l47_BumperInner, LM_Inserts_l47_BumperOuter, LM_Inserts_l47_Car, LM_Inserts_l47_Parts, LM_Inserts_l47_Playfield, LM_Inserts_l47_RRmp, LM_Inserts_l47_Rails, LM_Inserts_l47_RmpCovers, LM_Inserts_l47_Sideblades, LM_Inserts_l47_StickersCRamp, LM_Inserts_l47_WireRamp, LM_Inserts_l48_BumperInner, LM_Inserts_l48_BumperOuter, LM_Inserts_l48_Car, _
  LM_Inserts_l48_Parts, LM_Inserts_l48_Playfield, LM_Inserts_l48_RRmp, LM_Inserts_l48_Rails, LM_Inserts_l48_RmpCovers, LM_Inserts_l48_Sideblades, LM_Inserts_l48_TargetL3, LM_Inserts_l48_WireRamp, LM_Inserts_l49_CRampEdges, LM_Inserts_l49_Parts, LM_Inserts_l49_Playfield, LM_Inserts_l50_Playfield, LM_Inserts_l50_Sideblades, LM_Inserts_l50_WireRamp, LM_Inserts_l51_LeftSling4, LM_Inserts_l51_Playfield, LM_Inserts_l51_Slings, LM_Inserts_l51_WireRamp, LM_Inserts_l52_CRampEdges, LM_Inserts_l52_Parts, LM_Inserts_l52_Playfield, LM_Inserts_l54_Parts, LM_Inserts_l54_Playfield, LM_Inserts_l54_RRmp, LM_Inserts_l54_WireRamp, LM_Inserts_l55_Parts, LM_Inserts_l55_Playfield, LM_Inserts_l55_RRmp, LM_Inserts_l55_WireRamp, LM_Inserts_l56_GateFlap, LM_Inserts_l56_Parts, LM_Inserts_l56_Playfield, LM_Inserts_l56_StickersCRamp, LM_Inserts_l56_WireRamp, LM_Inserts_l57_Parts, LM_Inserts_l57_Playfield, LM_Inserts_l57_TargetL2, LM_Inserts_l57_TargetL3, LM_Inserts_l57_WireRamp, LM_Inserts_l58_Parts, LM_Inserts_l58_Playfield, _
  LM_Inserts_l58_TargetL1, LM_Inserts_l58_TargetL2, LM_Inserts_l58_WireRamp, LM_Inserts_l59_Playfield, LM_Inserts_l59_WireRamp, LM_Inserts_l6_CRamp, LM_Inserts_l6_Playfield, LM_Inserts_l6_StickersCRamp, LM_Inserts_l7_Playfield, LM_Inserts_l7_StickersCRamp, LM_Inserts_l8_Parts, LM_Inserts_l8_Playfield, LM_Inserts_l8_StickersCRamp, LM_Inserts_l9_Playfield)
Dim BG_All: BG_All=Array(BM_Bumper1Ring, BM_Bumper1Socket, BM_Bumper2Ring, BM_Bumper2Socket, BM_Bumper3Ring, BM_Bumper3Socket, BM_BumperInner, BM_BumperOuter, BM_CRamp, BM_CRampEdges, BM_Car, BM_DiverterArm, BM_FlipperL, BM_FlipperR, BM_Gate003_Wire, BM_Gate02Wire, BM_Gate04Wire, BM_Gate05Wire, BM_GateFlap, BM_GateSw44Wire, BM_GateSw45Wire, BM_GateSw46Wire, BM_LeftSling1, BM_LeftSling2, BM_LeftSling3, BM_LeftSling4, BM_Parts, BM_Playfield, BM_RPlastics, BM_RRmp, BM_Rails, BM_Rightsling1, BM_Rightsling2, BM_Rightsling3, BM_Rightsling4, BM_RmpCovers, BM_Sideblades, BM_Slings, BM_SpecialPost, BM_Spinner, BM_StickersCRamp, BM_TargetL1, BM_TargetL2, BM_TargetL3, BM_TargetM1, BM_TargetM2, BM_TargetM3, BM_TargetR1, BM_TargetR2, BM_TargetR3, BM_Target_sw35, BM_WireRamp, BM_sw14Rollover, BM_sw49Rollover, BM_sw50Rollover, BM_sw51Rollover, BM_sw52Rollover, BM_sw54Rollover, BM_sw55Rollover, BM_sw56Rollover, LM_Flashers_f109_CRampEdges, LM_Flashers_f109_DiverterArm, LM_Flashers_f109_GateFlap, LM_Flashers_f109_Parts, _
  LM_Flashers_f109_Playfield, LM_Flashers_f109_RPlastics, LM_Flashers_f109_RRmp, LM_Flashers_f109_RmpCovers, LM_Flashers_f109_Sideblades, LM_Flashers_f109_Slings, LM_Flashers_f109_WireRamp, LM_Flashers_f113_CRampEdges, LM_Flashers_f113_Gate05Wire, LM_Flashers_f113_Parts, LM_Flashers_f113_Playfield, LM_Flashers_f113_Rails, LM_Flashers_f113_Sideblades, LM_Flashers_f113_Spinner, LM_Flashers_f113_StickersCRamp, LM_Flashers_f114_Gate02Wire, LM_Flashers_f114_Parts, LM_Flashers_f114_Playfield, LM_Flashers_f114_RRmp, LM_Flashers_f114_Target_sw35, LM_Flashers_f114_WireRamp, LM_Flashers_f114_sw54Rollover, LM_Flashers_f114a_Parts, LM_Flashers_f114a_Playfield, LM_Flashers_f114a_StickersCRamp, LM_Flashers_f114a_TargetL1, LM_Flashers_f114a_TargetL2, LM_Flashers_f114a_TargetL3, LM_Flashers_f114a_WireRamp, LM_Flashers_f125_Bumper3Ring, LM_Flashers_f125_CRampEdges, LM_Flashers_f125_LeftSling1, LM_Flashers_f125_LeftSling2, LM_Flashers_f125_LeftSling3, LM_Flashers_f125_LeftSling4, LM_Flashers_f125_Parts, _
  LM_Flashers_f125_Playfield, LM_Flashers_f125_RRmp, LM_Flashers_f125_Rails, LM_Flashers_f125_Rightsling4, LM_Flashers_f125_Sideblades, LM_Flashers_f125_Slings, LM_Flashers_f125_SpecialPost, LM_Flashers_f125_TargetL1, LM_Flashers_f125_TargetL2, LM_Flashers_f125_TargetL3, LM_Flashers_f125_WireRamp, LM_Flashers_f126_Bumper1Ring, LM_Flashers_f126_Bumper1Socket, LM_Flashers_f126_Bumper2Ring, LM_Flashers_f126_Bumper3Ring, LM_Flashers_f126_BumperInner, LM_Flashers_f126_BumperOuter, LM_Flashers_f126_CRamp, LM_Flashers_f126_CRampEdges, LM_Flashers_f126_Car, LM_Flashers_f126_Gate02Wire, LM_Flashers_f126_GateFlap, LM_Flashers_f126_Parts, LM_Flashers_f126_Playfield, LM_Flashers_f126_RPlastics, LM_Flashers_f126_RRmp, LM_Flashers_f126_Rails, LM_Flashers_f126_RmpCovers, LM_Flashers_f126_Sideblades, LM_Flashers_f126_Slings, LM_Flashers_f126_Target_sw35, LM_Flashers_f126_WireRamp, LM_Flashers_f126_sw54Rollover, LM_Flashers_f126_sw55Rollover, LM_Flashers_f126_sw56Rollover, LM_Flashers_f127_Bumper1Ring, _
  LM_Flashers_f127_Bumper2Ring, LM_Flashers_f127_Bumper3Ring, LM_Flashers_f127_Bumper3Socket, LM_Flashers_f127_BumperInner, LM_Flashers_f127_BumperOuter, LM_Flashers_f127_CRamp, LM_Flashers_f127_CRampEdges, LM_Flashers_f127_Car, LM_Flashers_f127_GateSw46Wire, LM_Flashers_f127_LeftSling4, LM_Flashers_f127_Parts, LM_Flashers_f127_Playfield, LM_Flashers_f127_RPlastics, LM_Flashers_f127_RRmp, LM_Flashers_f127_Rails, LM_Flashers_f127_Rightsling3, LM_Flashers_f127_Rightsling4, LM_Flashers_f127_RmpCovers, LM_Flashers_f127_Slings, LM_Flashers_f127_SpecialPost, LM_Flashers_f127_Spinner, LM_Flashers_f127_StickersCRamp, LM_Flashers_f127_TargetM1, LM_Flashers_f127_TargetM2, LM_Flashers_f127_TargetM3, LM_Flashers_f127_TargetR1, LM_Flashers_f127_TargetR2, LM_Flashers_f127_WireRamp, LM_Flashers_f127_sw54Rollover, LM_Flashers_f127_sw55Rollover, LM_Flashers_f127_sw56Rollover, LM_Flashers_f128_Bumper3Ring, LM_Flashers_f128_BumperInner, LM_Flashers_f128_BumperOuter, LM_Flashers_f128_CRampEdges, LM_Flashers_f128_Parts, _
  LM_Flashers_f128_Playfield, LM_Flashers_f128_RPlastics, LM_Flashers_f128_RRmp, LM_Flashers_f128_Sideblades, LM_Flashers_f128_Spinner, LM_Flashers_f128_StickersCRamp, LM_Flashers_f128_TargetM1, LM_Flashers_f128_TargetM2, LM_Flashers_f128_TargetM3, LM_Flashers_f128_TargetR1, LM_Flashers_f128_TargetR2, LM_Flashers_f128_TargetR3, LM_Flashers_f129_Bumper3Ring, LM_Flashers_f129_CRampEdges, LM_Flashers_f129_Parts, LM_Flashers_f129_Playfield, LM_Flashers_f129_StickersCRamp, LM_Flashers_f129_TargetM1, LM_Flashers_f129_TargetR1, LM_Flashers_f129_TargetR2, LM_Flashers_f129_TargetR3, LM_Flashers_f129_WireRamp, LM_Flashers_f129a_Parts, LM_Flashers_f129a_Playfield, LM_Flashers_f129a_RPlastics, LM_Flashers_f129a_RRmp, LM_Flashers_f129a_Rails, LM_Flashers_f129a_RmpCovers, LM_Flashers_f129a_Slings, LM_Flashers_f130_CRampEdges, LM_Flashers_f130_GateSw45Wire, LM_Flashers_f130_Parts, LM_Flashers_f130_Playfield, LM_Flashers_f130_Rails, LM_Flashers_f130_StickersCRamp, LM_Flashers_f130_TargetL1, LM_Flashers_f130_TargetL2, _
  LM_Flashers_f130_TargetL3, LM_Flashers_f130_TargetM3, LM_Flashers_f130_TargetR1, LM_Flashers_f130_TargetR2, LM_Flashers_f130_TargetR3, LM_Flashers_f130_WireRamp, LM_GI_Bumper1Ring, LM_GI_Bumper1Socket, LM_GI_Bumper2Ring, LM_GI_Bumper2Socket, LM_GI_Bumper3Ring, LM_GI_Bumper3Socket, LM_GI_BumperInner, LM_GI_BumperOuter, LM_GI_CRamp, LM_GI_CRampEdges, LM_GI_Car, LM_GI_FlipperL, LM_GI_FlipperR, LM_GI_Gate02Wire, LM_GI_Gate04Wire, LM_GI_Gate05Wire, LM_GI_GateFlap, LM_GI_GateSw44Wire, LM_GI_GateSw45Wire, LM_GI_GateSw46Wire, LM_GI_LeftSling1, LM_GI_LeftSling2, LM_GI_LeftSling3, LM_GI_LeftSling4, LM_GI_Parts, LM_GI_Playfield, LM_GI_RPlastics, LM_GI_RRmp, LM_GI_Rails, LM_GI_Rightsling1, LM_GI_Rightsling2, LM_GI_Rightsling3, LM_GI_Rightsling4, LM_GI_RmpCovers, LM_GI_Sideblades, LM_GI_Slings, LM_GI_SpecialPost, LM_GI_Spinner, LM_GI_StickersCRamp, LM_GI_TargetL1, LM_GI_TargetL2, LM_GI_TargetL3, LM_GI_TargetM1, LM_GI_TargetM2, LM_GI_TargetM3, LM_GI_TargetR1, LM_GI_TargetR2, LM_GI_TargetR3, LM_GI_Target_sw35, _
  LM_GI_WireRamp, LM_GI_sw49Rollover, LM_GI_sw50Rollover, LM_GI_sw51Rollover, LM_GI_sw52Rollover, LM_GI_sw54Rollover, LM_GI_sw55Rollover, LM_GI_sw56Rollover, LM_GISplit_gi01_Bumper3Ring, LM_GISplit_gi01_Bumper3Socket, LM_GISplit_gi01_BumperInner, LM_GISplit_gi01_BumperOuter, LM_GISplit_gi01_CRampEdges, LM_GISplit_gi01_Parts, LM_GISplit_gi01_Playfield, LM_GISplit_gi01_StickersCRamp, LM_GISplit_gi01_WireRamp, LM_GISplit_gi02_CRampEdges, LM_GISplit_gi02_FlipperR, LM_GISplit_gi02_Parts, LM_GISplit_gi02_Playfield, LM_GISplit_gi02_Rightsling1, LM_GISplit_gi02_Rightsling2, LM_GISplit_gi02_Rightsling3, LM_GISplit_gi02_Rightsling4, LM_GISplit_gi02_Slings, LM_GISplit_gi02_WireRamp, LM_GISplit_gi02_sw49Rollover, LM_GISplit_gi02_sw52Rollover, LM_GISplit_gi05_FlipperL, LM_GISplit_gi05_LeftSling1, LM_GISplit_gi05_LeftSling2, LM_GISplit_gi05_LeftSling3, LM_GISplit_gi05_LeftSling4, LM_GISplit_gi05_Parts, LM_GISplit_gi05_Playfield, LM_GISplit_gi05_Slings, LM_GISplit_gi05_WireRamp, LM_GISplit_gi05_sw50Rollover, _
  LM_GISplit_gi05_sw51Rollover, LM_GISplit_gi08_Bumper1Ring, LM_GISplit_gi08_Bumper1Socket, LM_GISplit_gi08_Bumper2Socket, LM_GISplit_gi08_BumperInner, LM_GISplit_gi08_CRampEdges, LM_GISplit_gi08_Gate02Wire, LM_GISplit_gi08_Gate04Wire, LM_GISplit_gi08_GateFlap, LM_GISplit_gi08_Parts, LM_GISplit_gi08_Playfield, LM_GISplit_gi08_RPlastics, LM_GISplit_gi08_RRmp, LM_GISplit_gi08_Slings, LM_GISplit_gi08_StickersCRamp, LM_GISplit_gi08_Target_sw35, LM_GISplit_gi08_WireRamp, LM_GISplit_gi08_sw54Rollover, LM_GISplit_gi08_sw55Rollover, LM_GISplit_gi08_sw56Rollover, LM_Inserts_l10_Playfield, LM_Inserts_l11_Playfield, LM_Inserts_l12_Playfield, LM_Inserts_l13_Playfield, LM_Inserts_l14_Playfield, LM_Inserts_l15_Parts, LM_Inserts_l15_Playfield, LM_Inserts_l15_RRmp, LM_Inserts_l15_StickersCRamp, LM_Inserts_l15_WireRamp, LM_Inserts_l16_CRampEdges, LM_Inserts_l16_Parts, LM_Inserts_l16_Playfield, LM_Inserts_l16_StickersCRamp, LM_Inserts_l17_LeftSling4, LM_Inserts_l17_Parts, LM_Inserts_l17_Playfield, LM_Inserts_l17_WireRamp, _
  LM_Inserts_l18_Playfield, LM_Inserts_l18_WireRamp, LM_Inserts_l19_Playfield, LM_Inserts_l19_WireRamp, LM_Inserts_l20_Playfield, LM_Inserts_l21_Playfield, LM_Inserts_l22_Playfield, LM_Inserts_l23_Playfield, LM_Inserts_l23_StickersCRamp, LM_Inserts_l23_WireRamp, LM_Inserts_l24_FlipperL, LM_Inserts_l24_FlipperR, LM_Inserts_l24_Playfield, LM_Inserts_l25_Playfield, LM_Inserts_l25_StickersCRamp, LM_Inserts_l26_Playfield, LM_Inserts_l26_StickersCRamp, LM_Inserts_l27_CRamp, LM_Inserts_l27_CRampEdges, LM_Inserts_l27_Playfield, LM_Inserts_l27_StickersCRamp, LM_Inserts_l28_CRamp, LM_Inserts_l28_CRampEdges, LM_Inserts_l28_Parts, LM_Inserts_l28_Playfield, LM_Inserts_l28_StickersCRamp, LM_Inserts_l29_CRampEdges, LM_Inserts_l29_Parts, LM_Inserts_l29_Playfield, LM_Inserts_l29_StickersCRamp, LM_Inserts_l30_Bumper3Ring, LM_Inserts_l30_BumperInner, LM_Inserts_l30_CRampEdges, LM_Inserts_l30_Parts, LM_Inserts_l30_Playfield, LM_Inserts_l30_StickersCRamp, LM_Inserts_l31_Playfield, LM_Inserts_l31_StickersCRamp, _
  LM_Inserts_l32_Playfield, LM_Inserts_l33_Parts, LM_Inserts_l33_Playfield, LM_Inserts_l33_TargetL1, LM_Inserts_l33_TargetL2, LM_Inserts_l33_TargetL3, LM_Inserts_l33_WireRamp, LM_Inserts_l34_Bumper1Ring, LM_Inserts_l34_BumperInner, LM_Inserts_l34_BumperOuter, LM_Inserts_l34_Parts, LM_Inserts_l34_Playfield, LM_Inserts_l34_RRmp, LM_Inserts_l34_StickersCRamp, LM_Inserts_l34_WireRamp, LM_Inserts_l35_Parts, LM_Inserts_l35_Playfield, LM_Inserts_l35_StickersCRamp, LM_Inserts_l35_TargetM1, LM_Inserts_l35_WireRamp, LM_Inserts_l36_CRampEdges, LM_Inserts_l36_Parts, LM_Inserts_l36_Playfield, LM_Inserts_l36_StickersCRamp, LM_Inserts_l36_TargetR2, LM_Inserts_l36_TargetR3, LM_Inserts_l37_Parts, LM_Inserts_l37_Playfield, LM_Inserts_l37_RRmp, LM_Inserts_l37_RmpCovers, LM_Inserts_l37_Target_sw35, LM_Inserts_l37_WireRamp, LM_Inserts_l38_Parts, LM_Inserts_l38_Playfield, LM_Inserts_l38_RRmp, LM_Inserts_l38_WireRamp, LM_Inserts_l39_Bumper2Socket, LM_Inserts_l39_Parts, LM_Inserts_l39_Playfield, LM_Inserts_l39_WireRamp, _
  LM_Inserts_l40_Parts, LM_Inserts_l40_Playfield, LM_Inserts_l40_RRmp, LM_Inserts_l40_RmpCovers, LM_Inserts_l40_WireRamp, LM_Inserts_l41_Playfield, LM_Inserts_l41_StickersCRamp, LM_Inserts_l41_TargetR2, LM_Inserts_l42_Parts, LM_Inserts_l42_Sideblades, LM_Inserts_l42_StickersCRamp, LM_Inserts_l43_Parts, LM_Inserts_l43_Sideblades, LM_Inserts_l43_StickersCRamp, LM_Inserts_l44_Parts, LM_Inserts_l44_Sideblades, LM_Inserts_l44_StickersCRamp, LM_Inserts_l45_Parts, LM_Inserts_l45_Sideblades, LM_Inserts_l45_StickersCRamp, LM_Inserts_l46_Parts, LM_Inserts_l46_Sideblades, LM_Inserts_l46_StickersCRamp, LM_Inserts_l47_BumperInner, LM_Inserts_l47_BumperOuter, LM_Inserts_l47_Car, LM_Inserts_l47_Parts, LM_Inserts_l47_Playfield, LM_Inserts_l47_RRmp, LM_Inserts_l47_Rails, LM_Inserts_l47_RmpCovers, LM_Inserts_l47_Sideblades, LM_Inserts_l47_StickersCRamp, LM_Inserts_l47_WireRamp, LM_Inserts_l48_BumperInner, LM_Inserts_l48_BumperOuter, LM_Inserts_l48_Car, LM_Inserts_l48_Parts, LM_Inserts_l48_Playfield, LM_Inserts_l48_RRmp, _
  LM_Inserts_l48_Rails, LM_Inserts_l48_RmpCovers, LM_Inserts_l48_Sideblades, LM_Inserts_l48_TargetL3, LM_Inserts_l48_WireRamp, LM_Inserts_l49_CRampEdges, LM_Inserts_l49_Parts, LM_Inserts_l49_Playfield, LM_Inserts_l50_Playfield, LM_Inserts_l50_Sideblades, LM_Inserts_l50_WireRamp, LM_Inserts_l51_LeftSling4, LM_Inserts_l51_Playfield, LM_Inserts_l51_Slings, LM_Inserts_l51_WireRamp, LM_Inserts_l52_CRampEdges, LM_Inserts_l52_Parts, LM_Inserts_l52_Playfield, LM_Inserts_l54_Parts, LM_Inserts_l54_Playfield, LM_Inserts_l54_RRmp, LM_Inserts_l54_WireRamp, LM_Inserts_l55_Parts, LM_Inserts_l55_Playfield, LM_Inserts_l55_RRmp, LM_Inserts_l55_WireRamp, LM_Inserts_l56_GateFlap, LM_Inserts_l56_Parts, LM_Inserts_l56_Playfield, LM_Inserts_l56_StickersCRamp, LM_Inserts_l56_WireRamp, LM_Inserts_l57_Parts, LM_Inserts_l57_Playfield, LM_Inserts_l57_TargetL2, LM_Inserts_l57_TargetL3, LM_Inserts_l57_WireRamp, LM_Inserts_l58_Parts, LM_Inserts_l58_Playfield, LM_Inserts_l58_TargetL1, LM_Inserts_l58_TargetL2, LM_Inserts_l58_WireRamp, _
  LM_Inserts_l59_Playfield, LM_Inserts_l59_WireRamp, LM_Inserts_l6_CRamp, LM_Inserts_l6_Playfield, LM_Inserts_l6_StickersCRamp, LM_Inserts_l7_Playfield, LM_Inserts_l7_StickersCRamp, LM_Inserts_l8_Parts, LM_Inserts_l8_Playfield, LM_Inserts_l8_StickersCRamp, LM_Inserts_l9_Playfield)
' VLM Arrays - End



' VPW Revisions
' *************
' 003: gtxjoe - Initial draft
'   002: gtxjoe - more progress
'   003: gtxjoe - increased length of playfield to allow for ramp. Using PF with insert cutouts.  Table working except for wire ramp.
' 004: benji  - added playfield mesh with holes cut and bevels for 3 sink kicker holes
'   005: gtxjoe - added PF text layer from b1x 5/20/21.  Ramp diverter logic working, so tab. Removed VR and majority of BadCats specific images/objects
'   007: gtxjoe - updated ramps based on pics. Height/z adjustments are not done yet.  Added Flupper bumpers and missing rubberband.  Started renaming objects to help with grouping for rendering process
' 008: benji  - imported new firing range spinner model
' 009: gtxjoe - rubber height set to 35. started adding gate covers, ramp flaps, diverter prim, playfield wire rails.  Added test mode: press 2 to raise posts. W,E,R,Y,U,I,O,P can be used to capture balls and test shots
'   010: gtxjoe - added insert lighting, removed playfield mesh as it was distorting bottom right of pf, right ramp height adjusted and initial GI lighting (still needs shaping)
' 011: benji  - updated playfield mesh to remove distortion
'   012: gtxjoe - added support for GI on/off. Initial test of visible flashers using flupper flashers
'   013: benji  - temp meshes and baked texture for firing range and county jail meshes imported for reference
'   014: gtxjoe - added missing lights/flashers, animated car lights, white rubbers,blue posts,bumperdisk,target color, added haze image to flasherbloom
'   015: gtxjoe - bumper height change to 1.5" height. Replaced ramp gate brackets. Fixed spinner pivot point. Captured balls now animated with car. Kickers now use playfield mesh cutout
'   016: gtxjoe - Added triggers for plastic and wire ramp sounds, added ramp entry sound.  Set Plastic ramp walls to 50.  Need to find right ramps sounds
'   017: benji/gtxjoe -  Switched to primitive temp ramp to test ramp end rattle.  Add debug code to check overall time of center ramp shot as it is too fast.
'   018: gtxjoe - Flipper strength changed from 3300 to 2300.  Reviewed NF physics script - Rubber dampening enabled.  Open debug window to see center ramp loop time and velocity.
'   019: benji - separated posts and sleeves in to dPosts and dSleeves collections for NF dampening, adjusted center ramp hole and added rubber stop 'CenterRampStop' assignd to dSleeves
'   020: benji - added updated collidable center ramp from Tomate including beveled drop-through end hole
'   022: tomate - updated plastic and wire Ramps prims and set them to non-colliadble objects, vpx collidable ramps roughly fixed and set them as collidable objects, diverter ramp repositioned
' 023: Sixtoe - Stripped VPX file of as much as possible for toolkit preperation.
' 024: Sixtoe - Significantly rebuilt table, new playfield mesh, tweaked layout and ramps, new rubber physical locations and materials, updated script to current standards, replaced kickers with controlled switches, roth targets implemented, dynamic shadows implemented, blocker walls added to stop ball traps, probably loads of other stuff
' 025: Sixtoe - Sorted out flasher assignments, adjusted flipper triggers, adjusted physics materials, adjusted collections, adjusted centre ramp end, sorted layers a bit more, simplified naming of vpx objects,
' 026: flux - Assigned insert light indexes, hooked up vpmMapLights
' 027a: apophis - Added GI and flasher code. Made all lights use incandescent fader. Fixed kicker trigger heights
'   028: benji - Imported blender results
'   029: apophis - Made lots of wonderful updates and then corrupted the file. FML
'   031: benji - Imported blender results over version 028
' 032: apophis - Imported udpates made in ver 029: Made collidables invisible. Imported and applied VLM materials. Initial VLM helper script integrated. RoomBrightness option added. Some movables coded. To do: Bumper sockets, Car, Gates, Slings
' 033: apophis - Removed traditional dynamic shadows and enabled new ones. Added gate, slingshot, and car animations. Fixed flipper shadow DB.
' 034: flux - added depth bias to playfield, setup mirror blades
'   035: benji - Imported blender results
'   036: benji - Imported blender results. Target fix
'   037: benji - Imported blender results. Playfield DB fix. New ramp prim and textures added.
' 038: apophis - Ambient shadows added. Added CenterRamp_off prim. New ramp prims and textures hooked up to GI.
' 039: apophis - Fixed MB release mech. Deleted old car prim. Added DT sounds.
'   040: benji - Imported blender results. Ramp improvements.
' 041: apophis - Fixed some center ramp GI fading issues. Added ball brightness code. Updated ball textures.
' 042: benji - Imported blender results. Messing with pop bumers, center ramp, and screws.
' 043: apophis - Center ramp DL tweaks. Ramp rolling sounds and triggers added. L47 and L48 lightmap tweaks.
' 044: benji - Adjusted sideblades, more blur, less intensity (darker bastically)
' 046: benji - Imported blender results. Updated ramp, bumper caps
' 047: apophis - Added some randomness to right saucer kick. Set Layer_3_BM_Room visibilty to false, CenterRamp DL adjustments. L47 and L48 LM visibility tweak. Deleted PFText flasher. For bumpers_outer_bm_room: created/assigned new material for, unchecked additive blend, scripted in DL change with GI.
' 048: apophis - Tuned/fixed lamp and GI fading. Changed logic for vpinmame version check
' 049: benji - Added refraction probe to ramp, disabled render backface on 'giOff' ramp
' 050: apophis - Included refraction on CenterRamp_off. Update ramp fading ad visibility logic. Roughness level on ramp set to 1
' 051: benji - Ramp fiddled with. Turned off giOn/Off externally imported ramps. assigned zCol Ramp material to 'Layer 3' in VLM.Visuals (toolkited center ramp). changed ramp color to white, opacity 1.
' 052: benji - render probes added to right ramp, and ramp protectors (Layers 4 and 5 in 'VLM.Visuals'). Bumper caps tweaked more.
' 053: apophis - Updated Bumpers_Outer_BM_Room material to VLM.Bake.Bumpers. Made new material for Layer_3_BM_Room called VLM.Bake.Ramp. Changed SetRoomBrightness() to update material's existing base color instead of the prim color. Removed WarmUpDone().
'   054: benji - Imported new Blender batch (Added back wood wall with some detail in blender, warmed world image, other misc things) Deleted CenterRampOn/Off meshes in vpx. Renamed separated layers in Blender for VLM.VISUALS Separated rails, baked sideblades &lockdownbar, to their own layer Turned on ‘render backfaces’ for Layer 3 CenterRamp in VLM.VISUALS Assigned VLM.BAKE.RAMP to CenterRamp Assigned refraction probes to center ramp, right ramp,‘Layer 1 Slings’, and bumper caps. Assigned ‘vlm.bake.bumpercap’ material to bumper caps and tweaked (still looks a little funky under the ramp) *Added “Rod” and “Spring” to ‘Moveables’ for VPX animation*
'   055: benji - shorted vlm.visuals layer names, baked right ramp material again. Commended out CenterRamp GI stuff in script (needs to reassigned?)
' 056: apophis - Added flex options menu. Removed CenterRamp stuff from script. Added VLM.Bake.Ramp2 material to SetRoomBrightness(). Set VLM.Bake.Bumpers material opacity to some non zero value. L47 and L48 LM visibility tweak. Moved car z position lower. Hiding rails in cab mode.
'   059: benji - Optimized reflections - disabled reflections for all lightmaps except for a select few which will be most prominent in the sideblades
' 060: flux - removed rod and spring.
'   061: tastywasps - Added VR Room.
' 062: leojreimroc - Added VR Backglass.  Added Alphanumerics to Desktop mode. Added VRRoom code. Fixed blinking table lights.  Fixed police car lights.
' 063: apophis - Made Parts bakemap static and require restart for Light Level option.
' 064: benji - new batch
' 065: apophis - Fixed minor code issue. Updated SetRoomBrightness. Updated ball image. Made "playfeild_mesh" prim not collidable. L47 and L48 LM vis tweak.
'   067: benji - Kickers broke...New batch with sideblades that have normal maps,metal rails ramp has normal maps. Removed upper insert light in Blender (needs VPX blooms now). Tweaked ramp and Bumper materials/refraction. Sling plastics still blurry.
' 068: apophis - Converted insert lamps to blooms. Renamed "playfeild_phys" to "playfeild_mesh". Fixed post passes. L47 and L48 LM vis tweak. Update ball brightness equ. Added materials to SetRoomBrightness. Fixed lightmap materials.
'   069: benji - New batch. added refraction probe to slings.
' 070: Sixtoe- (benji sumarizing) added lightmap active material to lightmap prims (they were missing)
' 071: Benji- Reimported CentralRampBM as two meshesb(splitting edges from main body for extr refraction probe effect). New "CenterRampEdges" prim in "RampEdges" Layer
' 072: apophis - Set insert blooms to hidden. Tweaked ball brightness. Fixed playfield reflections. L47 and L48 LM vis tweak.
'   073: Benji - Updated Center Ramp, reduced polys/remodeled some edges, re-unwrapped UVs. Reimported center ramp.BM and split edges and added new 'scratches' normal map to it. dialed in vpx materials for bumper caps and ramp more
' 074: apophis - Added moving flasher objects for car lights on sidewall. Lightmap material fix.
' 075: apophis - Added option to turn off PWM flashers.
'   077: Benji - New batch. Improved side rails and flasher domes, Fixxd shark money inserts black when off. MEtal rails position borked. Roation of moveables borked.
' 079: Benji - Updated batch. Ramp still borked, but rotated it and its light maps manually now its good. Getting veeery close. Ready for more scrutiny.
' 080: apophis - Fixed right center drop target. L47 and L48 LM visibility tweak. Added reflection probe to BM_Playfield (strength 0.5). Fixed option menu default LUT. Added SetLocale. Updated fullscreen pov (thanks Eighties8). Set flahsers fader to LED (PWM handles fading). Refactored option loading logic.
' 081: apophis - Fixed issue caused by missing VR room object. Increased flipper strength to 2500 (was 2300). Added motor sound effect to moving car.
' 082: apophis - Minor post pass adjustments. Added outlane difficulty option. Put Cor.Update in its own timer. Made physical apron wall taller (300) to prevent Apron stuck ballz.
' 083: apophis - Fixed goof bar positions. Adjusted sling posts and inlane guides to be symmetrical relative to flippers.
' 084: benji - Updated batch, less light on firing range. Less blurry County Jail. Second giraff in upper right playfield now visible. Deleted some old textures.
' 085: apophis - Fixed broken rubber dampening (Added Cor.Update back into a timer). Re-fixes: Added reflection probe to BM_Playfield (strength 0.5).  L47 and L48 LM visibility tweak.
' 086: benji - changed BM_SpecialPost to not static so that it will move live wen dmd options are changed. Updated vpx collidable ramp to match visible rails ramp better
' 087: apophis - Disabled ball reflections for all alphanumeric desktop lights.
' 088: benji - Added Desktop background image
' 089: apophis - Added fix for standalone (thanks somatik!). Moved desktop alphanumeric lights to fit with new background image. Added Hot Sheet lights to desktop background.
'Release 1.0
'   1.0.1 - apophis - Fixed trough swtich order.
'   1.0.2 - apophis - Fixed Lock 2 insert light name.
'   1.0.3 - apophis - Set all render probe roughness levels to 0. Added FlipperCradleCollision.
'Release 1.1

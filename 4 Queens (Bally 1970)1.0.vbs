'*   _   ___    _______  __   __  _______  _______  __    _  _______
'*  | | |   |  |       ||  | |  ||       ||       ||  |  | ||       |
'*  | |_|   |  |   _   ||  | |  ||    ___||    ___||   |_| ||  _____|
'*  |       |  |  | |  ||  |_|  ||   |___ |   |___ |       || |_____
'*  |___    |  |  |_|  ||       ||    ___||    ___||  _    ||_____  |
'*      |   |  |      | |       ||   |___ |   |___ | | |   | _____| |
'*      |___|  |____||_||_______||_______||_______||_|  |__||_______|
'*
'*                    Bally (1970)
'*                     Nestorgian
'*
'*      4 Queens / IPD No. 936 / December, 1970 / 1 Player
'*
'*
'*

option explicit
Randomize


'*******************************************
'  User Options
'*******************************************

'----- Shadow Options -----
Const DynamicBallShadowsOn = 1    '0 = no dynamic ball shadow ("triangles" near slings and such), 1 = enable dynamic ball shadow
Const AmbientBallShadowOn = 1   '0 = Static shadow under ball ("flasher" image, like JP's)
                  '1 = Moving ball shadow ("primitive" object, like ninuzzu's) - This is the only one that shows up on the pf when in ramps and fades when close to lights!
                  '2 = flasher image shadow, but it moves like ninuzzu's


'*******************************************
'  Constants and Global Variables
'*******************************************

Const UsingROM = False        'The UsingROM flag is to indicate code that requires ROM usage. Mostly for instructional purposes only.

Const BallSize = 50         'Ball size must be 50
Const BallMass = 1          'Ball mass must be 1
Const tnob = 1            'Total number of balls
Const lob = 0           'Locked balls

ExecuteGlobal GetTextFile("core.vbs")

On Error Resume Next
ExecuteGlobal GetTextFile("Controller.vbs")
If Err Then MsgBox "Unable to open Controller.vbs. Ensure that it is in the scripts folder."
On Error Goto 0

Dim tablewidth: tablewidth = Table1.width
Dim tableheight: tableheight = Table1.height

dim ballinplay
dim ballrelenabled
dim rst
dim eg
dim credit
dim score
dim truesc
dim match(9)
dim rep
dim tilt
dim tiltsens
dim state
dim cred
dim plm
dim update
dim digit
dim tempscore
dim scn
dim scn1
dim bell
dim points
dim matchnumb
dim replay1
dim replay2
dim replay3
dim replay4
dim hisc
'dim rs(3)
dim rv(3)
dim sv
dim sv1
dim i
dim TextStr
dim reel(2)
dim hv
dim wv
dim qs
dim FlipperPosition
dim ABCDCounter,QueenCounter,ZeroNine,AltRelay
dim ballspergame, replayset
Dim obj
'Dim RFPress
Dim plungerpress

Dim B2SOn
Dim controller
Dim B2SScore  ' B2S Score Displayed
Const TableName="Bally4Queens_1970"
Const cGameName = "Bally4Queens_1970"

Dim DesktopMode
DesktopMode = Table1.ShowDT

'----- VR Room -----
Dim VRRoomChoice : VRRoomChoice = 2   ' 1 - Mega Room, 2 - Minimal Room
Const VRTest = 0            ' 1 = Testing VR in Live View, 0 = Do not force VR mode.
Dim VRMode, VR_Obj

If RenderingMode = 2 Or VRTest Then
  VRMode = True
Else
  VRMode = False
End If


' The game timer interval is 10 ms
Sub GameTimer_Timer()
  Cor.Update            'update ball tracking (this sometimes goes in the RDampen_Timer sub)
End Sub

FrameTimer.Interval = -1
FrameTimer.Enabled = True
Sub FrameTimer_Timer()
' If DynamicBallShadowsOn Or AmbientBallShadowOn Then DynamicBSUpdate 'update ball shadows
  RollingUpdate         'update rolling sounds
  DoSTAnim
  UpdateStandupTargets
  AnimateBumperSkirts
  UpdatePlunger
  BSUpdate
End Sub


Dim EMBall1

Sub UpdatePlunger()
  If plungerpress = 1 then
    If Pincab_Plunger.Y < 120 then
      Pincab_Plunger.Y = Pincab_Plunger.Y + 5
    End If
  Else
    Pincab_Plunger.Y = 0 + (5* Plunger.Position) -20
  End If
End Sub

sub Table1_init
  Set EMBall1 = Drain.CreateSizedballWithMass(Ballsize/2,Ballmass)
  If Table1.ShowDT = false then
    For each obj in DTItems
      obj.visible=False
    next
  End If
  LoadEM
  qs=0
  wv=0
' score=0
  matchnumb=5
  ballspergame=5
  replayset=1

    set match(0)=m0
    set match(1)=m1
    set match(2)=m2
    set match(3)=m3
    set match(4)=m4
    set match(5)=m5
    set match(6)=m6
    set match(7)=m7
    set match(8)=m8
    set match(9)=m9
    set reel(1)=reel1

    loadhs
' SetReel score
  If VRMode Then
    setcreditreel(credit)
    SetReel hisc
  End If

    if hisc="" then hisc=50000
    reel1.setvalue(score)
    reel2.setvalue(hisc)
    if credit="" then credit=0
    credittxt.text=credit
    select case(matchnumb)
    case 0:
    m0.text="00"
    case 1:
    m1.text="10"
    case 2:
    m2.text="20"
    case 3:
    m3.text="30"
    case 4:
    m4.text="40"
    case 5:
    m5.text="50"
    case 6:
    m6.text="60"
    case 7:
    m7.text="70"
    case 8:
    m8.text="80"
    case 9:
    m9.text="90"
    end select
  ShowReplayLevels
' B2SOn=True

  If B2SOn Then
    Set Controller = CreateObject("B2S.Server")
    Controller.B2SName = "Bally4Queens_1970"
    Controller.Run()
    If Err Then MsgBox "Can't Load B2S.Server."
    if matchnumb=0 then
      Controller.B2SSetMatch 10
    else
      Controller.B2SSetMatch matchnumb
    end if
    If VRMode Then
      FlasherMatch
    End If
    Controller.B2SSetScoreRolloverPlayer1 0


    Controller.B2SSetTilt 1
    Controller.B2SSetCredits Credit
    Controller.B2SSetGameOver 1
    Controller.B2SSetData 81,0
    Controller.B2SSetData 82,0
    Controller.B2SSetData 83,0
    Controller.B2SSetData 84,0
  End If
  for i=1 to 4
    If B2SOn Then
      Controller.B2SSetScorePlayer i, 0
    End If
  next
  Init4Q
  If VRMode Then
    for each Object in VRBGGameOver : object.visible = 1 : next
    setcreditreel(credit)
    SetBackglass
    SetBackglassLights
    BackglassTimer.Enabled = True
  End If
  flipopen
  wv=0
end sub

Sub Table1_Paused
  If B2SOn Then Controller.Pause = 1
End Sub
Sub Table1_unPaused
  If B2SOn Then Controller.Pause = 0
End Sub
Sub Table1_exit()
  savehs
  If B2SOn Then Controller.Pause = False
  If B2SOn Then Controller.Stop
End Sub


sub Init4Q
  If VRMode Then
    LBgA.visible=0
    L2BgA.visible=0
    LBgB001.visible=0
    L2BgB.visible=0
    LBgC.visible=0
    L2BgC.visible=0
    LBgD.visible=0
    L2BgD.visible=0
  End If
End Sub


'*******************************************
'  ZOPT: User Options
'*******************************************
Dim LightLevel : LightLevel = 0.25        ' Level of room lighting (0 to 1), where 0 is dark and 100 is brightest
Dim VolumeDial : VolumeDial = 0.8             ' Overall Mechanical sound effect volume. Recommended values should be no greater than 1.
Dim BallRollVolume : BallRollVolume = 0.5     ' Level of ball rolling volume. Value between 0 and 1
Dim ColorLUT : ColorLUT = 2           ' Color desaturation LUTs: 1 to 12, where 1 is vibrant, 2 is normal and 12 is black'n'white
Dim dspTriggered : dspTriggered = False

Function ReflectionToggle(state)
  Dim BP
  For Each BP in VRReel1
    BP.ReflectionEnabled = state
  Next
  Pincab_Backglass.ReflectionEnabled = state
  Pincab_BgL.ReflectionEnabled = state
  Pincab_BackglassL.ReflectionEnabled = state
End Function

Sub Table1_OptionEvent(ByVal eventId)
    If eventId = 1 And Not dspTriggered Then dspTriggered = True : DisableStaticPreRendering = True : End If
  dim v

  ' Color Saturation
    ColorLUT = Table1.Option("Color Saturation", 1, 12, 1, 1, 0, _
    Array("Normal", "Desaturated 10%", "Desaturated 20%", "Desaturated 30%", "Desaturated 40%", "Desaturated 50%", _
        "Desaturated 60%", "Desaturated 70%", "Desaturated 80%", "Desaturated 90%", "Black 'n White", "Vibrant"))

  if ColorLUT = 1 Then Table1.ColorGradeImage = "ColorGradeLUT256x16_1to1"
  if ColorLUT = 2 Then Table1.ColorGradeImage = "ColorGradeLUT256x16_90"
  if ColorLUT = 3 Then Table1.ColorGradeImage = "ColorGradeLUT256x16_80"
  if ColorLUT = 4 Then Table1.ColorGradeImage = "ColorGradeLUT256x16_70"
  if ColorLUT = 5 Then Table1.ColorGradeImage = "ColorGradeLUT256x16_60"
  if ColorLUT = 6 Then Table1.ColorGradeImage = "ColorGradeLUT256x16_50"
  if ColorLUT = 7 Then Table1.ColorGradeImage = "ColorGradeLUT256x16_40"
  if ColorLUT = 8 Then Table1.ColorGradeImage = "ColorGradeLUT256x16_30"
  if ColorLUT = 9 Then Table1.ColorGradeImage = "ColorGradeLUT256x16_20"
  if ColorLUT = 10 Then Table1.ColorGradeImage = "ColorGradeLUT256x16_10"
  if ColorLUT = 11 Then Table1.ColorGradeImage = "ColorGradeLUT256x16_00"
  if ColorLUT = 12 Then Table1.ColorGradeImage = "ColorGradeLUT256x16_1to1gam0.7vibr80"

    ' Sound volumes
    VolumeDial = Table1.Option("Mech Volume", 0, 1, 0.01, 0.8, 1)
    BallRollVolume = Table1.Option("Ball Roll Volume", 0, 1, 0.01, 0.5, 1)

  ' Room brightness
  LightLevel = NightDay/100
  SetRoomBrightness LightLevel   'Uncomment this line for lightmapped tables.

  ' VR Room
    VRRoomChoice = Table1.Option("VR Room", 1, 2, 1, 2, 0, Array("Black", "Minimal Room"))
  SetupRoom

  'Cabinet rails
  If RenderingMode <> 2 Then
    v = Table1.Option("Cabinet Rails", 0, 1, 1, 1, 0, Array("Hide", "Show"))
    Pincab_Siderails.Visible = v
    Pincab_Lockdownbar.Visible = v
  End If

  ' Toggle Reflections
  If RenderingMode = 2 Or VRTest Then
    v = Table1.Option("BackGlass Reflections", 0, 1, 1, 1, 0, Array("No", "Yes"))
    ReflectionToggle(v)
  End If

  'Balls per Game
  v = Table1.Option("Balls per Game", 0, 1, 1, 1, 0, Array("3", "5"))
  If v=0 Then ballspergame=3
  If v=1 Then ballspergame=5
  ShowReplayLevels

 If eventId = 3 And dspTriggered Then dspTriggered = False : DisableStaticPreRendering = False : End If
End Sub


sub LmodeOff
  Dim Li
  For each Li in LMode : Li.state = 0: Next

End Sub


'****************************
'   ZRBR: Room Brightness
'****************************

'This code only applies to lightmapped tables. It is here for reference.
'NOTE: Objects bightness will be affected by the Day/Night slider only if their blenddisablelighting property is less than 1.
'      Lightmapped table primitives have their blenddisablelighting equal to 1, therefore we need this SetRoomBrightness sub
'      to handle updating their effective ambient brighness.

' Update these arrays if you want to change more materials with room light level
Dim RoomBrightnessMtlArray: RoomBrightnessMtlArray = Array("VLM.Bake.Active","VLM.Bake.Solid","VLM.Bake.Metal")

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







Sub ShowReplayLevels
  Dim tempst1, tempst2
  If ballspergame =3 Then
    InstCard.image= "IC3"
    InstCard2.image= "IC2-3"
    replay1=33000
    replay2=42000
    replay3=51000
    replay4=60000
  Else
    InstCard.image= "IC5"
    InstCard2.image= "IC2-5"
    replay1=54000
    replay2=63000
    replay3=72000
    replay4=81000
  End If
  tempst1=FormatNumber(replayset,0)
  tempst2=FormatNumber(ballspergame,0)
' If VRMode Then
'   For Each VR_Obj In BG_flashers : VR_Obj.Visible = 1 : Next
'   For Each VR_Obj In BG_flashersL : VR_Obj.Visible = 1 : Next
' End If
End Sub



Sub Table1_KeyDown(ByVal keycode)
    If keycode = LeftFlipperKey Then
      Pincab_FlipperLeft.x = Pincab_FlipperLeft.x + 6
    End If
    If keycode = RightFlipperKey Then
      Pincab_FlipperRight.x = Pincab_FlipperRight.x - 6
    End If
    If keycode=StartGameKey Then
      Pincab_startButt.y = Pincab_startButt.y - 3
    End If
' if keycode=RightMagnaSave and state=false then
'   if ballspergame=5 then
'     ballspergame=3
'   else
'     ballspergame=5
'   end if
'   ShowReplayLevels
' end if
' if keycode=LeftMagnaSave and state=false then
'   replayset=replayset+1
'   if replayset>6 then replayset=1
'   ShowReplayLevels
' end if
  if keycode=AddCreditKey then
    playsound "coin3"
    coindelay.enabled=true
  end if
  if keycode=StartGameKey and credit>0 and state=false then
    credit=credit-1
    LmodeOff
    If B2SOn Then
      Controller.B2SSetCredits Credit
    end if
    If VRMode Then
      setcreditreel(credit)
      StartResetReels
      FlasherPlayer
    End If
    qs=qs-1
    if qs<0 then qs=0
    playsound "click"
    credittxt.text=credit
    tilt=false
    state=true
    rst=0
    ballinplay=1
    If VRMode=False Then
      playsound "initialize"
    End If
    resettimer.enabled=true
    If B2SOn Then
      Controller.B2SSetTilt 0
      Controller.B2SSetGameOver 0
      Controller.B2SSetMatch 0
      Controller.B2SSetCredits Credit

      Controller.B2SSetCanPlay 1
      Controller.B2SSetPlayerUp 1
      Controller.B2SSetData 81,0
      Controller.B2SSetData 82,0
      Controller.B2SSetData 83,0
      Controller.B2SSetData 84,0
      Controller.B2SSetBallInPlay BallInPlay
      Controller.B2SSetScoreRolloverPlayer1 0
    End If
    Init4Q
    If VRMode Then
      If score=0 Then
        playsound "initialize"
      End If
      for each Object in VRBGTilt : object.visible = 0 : next
      for each Object in VRBGMatch : object.visible = 0 : next
      for each Object in VRBGGameOver : object.visible = 0 : next
      FlasherBalls
    End If
  end if
  If keycode = PlungerKey Then
    Plunger.PullBack
    plungerpress = 1
  End If
  if tilt=false and state=true then
    If keycode = LeftFlipperKey Then
      FlipperActivate LeftFlipper, LFPress
      FlipperActivate LeftFlipper1, LFPress
      SolLFlipper True            'This would be called by the solenoid callbacks if using a ROM
      PlaySoundAt SoundFXDOF("FlipperUp",201,DOFOn,DOFFlippers), LeftFlipper
      PlayLoopSoundAtVol "buzzL", LeftFlipper, 1
    End If
    If keycode = RightFlipperKey Then
      FlipperActivate RightFlipper, RFPress
      FlipperActivate RightFlipper1, RFPress
      SolRFlipper True            'This would be called by the solenoid callbacks if using a ROM
      PlaySoundAt SoundFXDOF("FlipperUp",202,DOFOn,DOFFlippers), RightFlipper
      PlayLoopSoundAtVol "buzz", RightFlipper, 1
    End If
    If keycode = LeftTiltKey Then
      Nudge 90, 2
      checktilt
    End If
    If keycode = RightTiltKey Then
      Nudge 270, 2
      checktilt
    End If
    If keycode = CenterTiltKey Then
      Nudge 0, 2
      checktilt
    End If
    If keycode = MechanicalTilt Then
      checktilt
    End If
  End if
End Sub

Sub Table1_KeyUp(ByVal keycode)
  If keycode = LeftFlipperKey Then
    Pincab_FlipperLeft.x = Pincab_FlipperLeft.x - 6
  End If
  If keycode = RightFlipperKey Then
    Pincab_FlipperRight.x = Pincab_FlipperRight.x + 6
  End If
  If keycode=StartGameKey Then
    Pincab_startButt.y = Pincab_startButt.y + 3
  End If
  If keycode = PlungerKey Then
  playsound "plunger"
  Plunger.Fire
  plungerpress = 0
  End If
  If keycode = LeftFlipperKey Then
  LeftFlipper.RotateToStart
  LeftFlipper1.RotateToStart
  PlaySoundAt SoundFXDOF("FlipperDown",201,DOFOff,DOFFlippers), LeftFlipper
  StopSound "buzzL"
  if tilt=false and state=true then PlaySound "FlipperDown"
  End If
  If keycode = RightFlipperKey Then
  RightFlipper.RotateToStart
  RightFlipper1.RotateToStart
  PlaySoundAt SoundFXDOF("FlipperDown",202,DOFOff,DOFFlippers), RightFlipper
  StopSound "buzz"
  if tilt=false and state=true then PlaySound "FlipperDown"
  End If
End Sub


'*******************************************
'  Flippers
'*******************************************
' playsound "buzz"
' PlaySound "FlipperUp"


Const ReflipAngle = 20

' Flipper Solenoid Callbacks (these subs mimics how you would handle flippers in ROM based tables)
Sub SolLFlipper(Enabled)
  If Enabled Then
    LF.Fire
    leftflipper1.rotatetoend

'   If leftflipper.currentangle < leftflipper.endangle + ReflipAngle Then
'     RandomSoundReflipUpLeft LeftFlipper
'   Else
'     SoundFlipperUpAttackLeft LeftFlipper
'     RandomSoundFlipperUpLeft LeftFlipper
'   End If
  Else
    LeftFlipper.RotateToStart
    LeftFlipper1.RotateToStar
'   If LeftFlipper.currentangle < LeftFlipper.startAngle - 5 Then
'     RandomSoundFlipperDownLeft LeftFlipper
'   End If
    FlipperLeftHitParm = FlipperUpSoundLevel
  End If
End Sub

Sub SolRFlipper(Enabled)
  If Enabled Then
    RF.Fire
    rightflipper1.rotatetoend

'   If rightflipper.currentangle > rightflipper.endangle - ReflipAngle Then
'     RandomSoundReflipUpRight RightFlipper
'   Else
'     SoundFlipperUpAttackRight RightFlipper
'     RandomSoundFlipperUpRight RightFlipper
'   End If
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


' This subroutine updates the flipper shadows and visual primitives
'Sub FlipperVisualUpdate
' FlipperLSh.RotZ = LeftFlipper.CurrentAngle
' FlipperRSh.RotZ = RightFlipper.CurrentAngle
' LFLogo.RotZ = LeftFlipper.CurrentAngle
' RFlogo.RotZ = RightFlipper.CurrentAngle
'End Sub




'************************************************************
' ZSLG: Slingshot Animations
'************************************************************

Dim RStep, Lstep

Sub RightSlingShot_Slingshot
    RS.VelocityCorrect(activeball)
    RandomSoundSlingshotRight BM_REMK
  DOF 205, DOFPulse
    RStep = 0 : RightSlingShot_Timer ' Initialize Step to 0
  RightSlingShot.TimerInterval = 17
  RightSlingShot.TimerEnabled = 1
    Addscore 10
End Sub

Sub RightSlingShot_Timer
  Dim BL
  Dim x1, x2, y: x1=True: x2=False: y=25
    Select Case RStep
        Case 2: x1=False: x2=True: y=15
        Case 3: x1=False: x2=False: y=0: RightSlingShot.TimerEnabled = 0
    End Select
  For Each BL in BP_RSling1 : BL.Visible = x1: Next
  For Each BL in BP_RSling2 : BL.Visible = x2: Next
  For Each BL in BP_REMK : BL.transx = y: Next
    RStep = RStep + 1
End Sub

Sub LeftSlingShot_Slingshot
    LS.VelocityCorrect(activeball)
  RandomSoundSlingshotLeft BM_LEMK
  DOF 203, DOFPulse
    LStep = 0 : LeftSlingShot_Timer ' Initialize Step to 0
  LeftSlingShot.TimerInterval = 17
  LeftSlingShot.TimerEnabled = 1
    Addscore 10
End Sub


Sub LeftSlingShot_Timer
  Dim BL
  Dim x1, x2, y: x1=True: x2=False: y=25
    Select Case LStep
        Case 3: x1=False: x2=True: y=15
        Case 4: x1=False: x2=False: y=0: LeftSlingShot.TimerEnabled = 0
    End Select
  For Each BL in BP_LSling1 : BL.Visible = x1: Next
  For Each BL in BP_LSling2 : BL.Visible = x2: Next
  For Each BL in BP_LEMK : BL.transx = y: Next
    LStep = LStep + 1
End Sub



sub coindelay_timer
  playsound "click"
  credit=credit+1
  if credit>25 then credit=25
  credittxt.text=credit
  coindelay.enabled=false
  If B2SOn Then
    Controller.B2SSetCredits Credit
  end if
  If VRMode Then
    setcreditreel(credit)
  End If
end sub

sub resettimer_timer
    rst=rst+1
    reel1.resettozero
  If B2SOn Then
    Controller.B2SSetScorePlayer1 0
  end if
    if rst=12 then
    playsound "kickerkick"
    end if
    if rst=13 then
    newgame
    resettimer.enabled=false
    end if
end sub

sub newgame

  score=0
  truesc=0
  eg=0
  rep=0
  sv=0
  bumperlight1.state=lightstateoff
  bumperlight2.state=0
  bumperlight3.state=0
  Light7.State=0
  Light8.State=0
  Light9.State=0
  Light10.State=0
  Light11.State=0
  ABCDCounter=0
  QueenCounter=0
  AltRelay=0
  wheelcheck
  LightA.State=1
  LightB.State=0
  LightC.State=0
  LightD.State=0
  gamov.text=" "
  tilttxt.text=" "
  bip5.text=" "
  bip1.text="1"
  for i=0 to 9
  match(i).text=" "
  next
  resettimer.enabled=True
  EMRelease True
  If B2SOn then Controller.B2SSetBallInPlay 1
  If VRmode Then
    FlasherBalls
  End If
end sub



sub newball
    bumperlight1.state=0
    bumperlight2.state=0
    bumperlight3.state=0
end sub


Sub EMRelease(enabled)
  If enabled Then
    Drain.kick 70, 12
    PlaysoundAt "ballrelease", Plunger
    DOF 233, DOFPulse
  End If
End Sub

Sub Drain_Hit()
  RandomSoundDrain Drain
  flipopen
  NextBallPause.enabled=true
End Sub

Sub NextBallPause_timer
  NextBallPause.enabled=false
  nextball
end sub

sub nextball
    if tilt=true then

    tilt=false
  If B2SOn Then Controller.B2SSetTilt 0
    tilttxt.text=" "
    end if
  'ABCDCounter=0
    'LightA.State=1
    'LightB.State=0
    'LightC.State=0
    'LightD.State=0


  ballinplay=ballinplay+1
  if ballinplay>ballspergame then
  playsound "motorleer"
  If B2SOn then Controller.B2SSetBallInPlay 0
  eg=1
  ballreltimer.enabled=true
  else
  if state=true and tilt=false then
  newball

  ballreltimer.enabled=true
  end if
  If B2SOn then Controller.B2SSetBallInPlay ballinplay
  select case (ballinplay)
  case 1:
  bip1.text="1"
  case 2:
  bip1.text=" "
  bip2.text="2"
  case 3:
  bip2.text=" "
  bip3.text="3"
  case 4:
  bip3.text=" "
  bip4.text="4"
  case 5:
  bip4.text=" "
  bip5.text="5"
  end select
  end if
  If VRmode Then
    FlasherBalls
  End If
End Sub

sub ballreltimer_timer
  if eg=1 then
  matchnum
  bip5.text=" "
  state=false
  gamov.text="GAME OVER"
  If B2SOn Then
    Controller.B2SSetGameOver 1
  end if
  If VRMode Then
    for each Object in VRBGGameOver : object.visible = 1 : next
  End If
  if truesc>hisc then hisc=truesc
  reel2.setvalue(hisc)
  savehs

  ballreltimer.enabled=false
  else
  playsound "kickerkick"
  EMRelease True
' nb.CreateBall
' nb.kick 90,4
    ballreltimer.enabled=false
    end if
end sub

sub matchnum
  select case(matchnumb)
  case 0:
  m0.text="00"
  case 1:
  m1.text="10"
  case 2:
  m2.text="20"
  case 3:
  m3.text="30"
  case 4:
  m4.text="40"
  case 5:
  m5.text="50"
  case 6:
  m6.text="60"
  case 7:
  m7.text="70"
  case 8:
  m8.text="80"
  case 9:
  m9.text="90"
  end select
  If B2SOn Then

    if matchnumb=0 then
      Controller.B2SSetMatch 100
    else
      Controller.B2SSetMatch (matchnumb*10)
    end if
  end if
  If VRMode Then
      FlasherMatch
    End If
  if (matchnumb*10)=(score mod 100) then
  credit=credit+1
    qs=qs+2
' playsound "knocke"
  PlaySound SoundFXDOF("knocke",117, DOFPulse, DOFKnocker)
  if credit>25 then credit=25
  If B2SOn Then
    Controller.B2SSetCredits Credit
  end if
  If VRMode Then
    setcreditreel(credit)
  End If
  credittxt.text=credit
    playsound "click"
    end if
end sub

sub addscore(points)

    if tilt=false then
    bell=0
    scn=0
    if points = 10 or points = 100 or points=1000 then scn=1

    if points = 5000 then
    reel(1).addvalue(5000)
    bell=1
    scn=5
    end if

    if points = 4000 then
    reel(1).addvalue(4000)
    bell=1
    scn=4
    end if

    if points = 3000 then
    reel(1).addvalue(3000)
    bell=1
    scn=3
    end if

    if points = 2000 then
    reel(1).addvalue(2000)
    bell=1
    scn=2
    end if

    if points = 1000 then
    reel(1).addvalue(1000)
    bell=1
    end if

    if points = 500 then
    reel(1).addvalue(500)
    bell=2
    scn=5
    end if

    if points = 300 then
    reel(1).addvalue(300)
    bell=2
    scn=3
    end if

    if points = 100 then
    reel(1).addvalue(100)
    bell=2
    end if

    if points = 10 then
    reel(1).addvalue(10)
    bell=3
    AdvanceZeroNine
    end if

    scn1=0
    scntimer.enabled=true
    score=score+points
    truesc=truesc+points
  If B2SOn Then
    Controller.B2SSetScore 1,score
  end if
    if score=>100000 then
    score=score-100000
    rep=0
    end if

    if score=>replay1 and rep=0 then
    credit=credit+1
    qs=qs+2
    PlaySound SoundFXDOF("knocke",117, DOFPulse, DOFKnocker)

  if credit>25 then credit=25
  credittxt.text=credit
  If B2SOn Then
    Controller.B2SSetCredits Credit
  end if
    rep=1
    playsound "click"
    end if

    if score=>replay2 and rep=1 then
    credit=credit+1
    qs=qs+2
    PlaySound SoundFXDOF("knocke",117, DOFPulse, DOFKnocker)

  if credit>25 then credit=25
  credittxt.text=credit
  If B2SOn Then
    Controller.B2SSetCredits Credit
  end if
    rep=2
    playsound "click"
    end if

    if score=>replay3 and rep=2 then
    credit=credit+1
    qs=qs+2
    PlaySound SoundFXDOF("knocke",117, DOFPulse, DOFKnocker)

  if credit>25 then credit=25
  credittxt.text=credit
  If B2SOn Then
    Controller.B2SSetCredits Credit
  end if
    rep=3
    playsound "click"
    end if

    if score=>replay4 and rep=3 then
    credit=credit+1

    qs=qs+2
    PlaySound SoundFXDOF("knocke",117, DOFPulse, DOFKnocker)

  if credit>25 then credit=25
  credittxt.text=credit
  If B2SOn Then
    Controller.B2SSetCredits Credit
  end if
    rep=4
    playsound "click"
    end if
    end if
  setreel score

end sub

Sub AdvanceZeroNine
  wv=wv+1
  AltRelay=AltRelay+1
  If AltRelay>1 then AltRelay=0
  if wv>9 then wv=0
  wheelcheck
end sub

sub wheelcheck
  If AltRelay=0 then
    LightT1.state=1
    LightT2.state=0
    Light6.state=0
  else
    LightT1.state=0
    LightT2.state=1
    Light6.state=1
  end if

    If wv=0 or wv=5 then

    Light1.state=1
    Light2.state=0
    Light3.state=0
    Light4.state=0
    Light5.state=0


    End If

    If wv=1 or wv=6 then
    Light1.state=0
    Light2.state=1
    Light3.state=0
    Light4.state=0
    Light5.state=0

    End If

    If wv=2 or wv=7 then
    Light1.state=0
    Light2.state=0
    Light3.state=1
    Light4.state=0
    Light5.state=0

    End If

    If wv=3 or wv=8 then
    Light1.state=0
    Light2.state=0
    Light3.state=0
    Light4.state=1
    Light5.state=0

    End If

    If wv=4 or wv=9 then
    Light1.state=0
    Light2.state=0
    Light3.state=0
    Light4.state=0
    Light5.state=1

    End If
end sub

Sub quickstep


End Sub

sub scntimer_timer
  scn1=scn1 + 1
  matchnumb=matchnumb+1
  if matchnumb=10 then matchnumb=0
  if bell=3 then playsound "bell10",0, 0.30, 0, 0
    if bell=2 then playsound "bell10",0, 0.30, 0, 0
    if bell=1 then playsound "bell1000",0, 0.45, 0, 0
    if scn1=scn then

    scntimer.enabled=false
    end if
end sub

Sub CheckTilt
  If Tilttimer.Enabled = True Then
  TiltSens = TiltSens + 1
  if TiltSens = 2 Then
  Tilt = True
  tilttxt.text="TILT"
  flipopen
  ballinplay=ballspergame
  If B2SOn Then Controller.B2SSetTilt 1
  playsound "tilt"
  If VRMode Then
    for each Object in VRBGTilt : object.visible = 1 : next
  End If
  if truesc>hisc then hisc=truesc
  reel2.setvalue(hisc)
  savehs
  turnoff
  End If
  Else
  TiltSens = 0
  Tilttimer.Enabled = True
  End If
End Sub

Sub Tilttimer_Timer()
  Tilttimer.Enabled = False
End Sub

sub turnoff
    bumper1.force=0
    bumper2.force=0
    bumper3.force=0
end sub


sub bumper1_hit
    if tilt=false then
    RandomSoundBumperMiddle Bumper1
    DOF 208, DOFPulse
    if (bumperlight1.state)=1 then
      addscore 100
      else
      addscore 10
    end if
  end if
end sub

sub bumper2_hit
    if tilt=false then
    RandomSoundBumperTop Bumper2
    DOF 209, DOFPulse
    if (bumperlight2.state)=1 then
      addscore 100
      else
      addscore 10
    end if
  end if
end sub

sub bumper3_hit
    if tilt=false then
    RandomSoundBumperBottom Bumper3
    DOF 210, DOFPulse
    if (bumperlight3.state)=1 then
      addscore 100
      else
      addscore 10
    end if
  end if
end sub

Sub Mush001_Hit()
  Dim BP
  RandomSoundBumperPas BM_Mush1
    Addscore 1000
    If LightA.State=1 then LightCheck
  For each BP in BP_Mush1 : BP.z = 4 : Next
  MushTimer.enabled=true

End Sub


Sub Mush004_Hit()
  Dim BP
  RandomSoundBumperPas BM_Mush4
    Addscore 1000
  If LightB.State=1 then LightCheck
    For each BP in BP_Mush4 : BP.z = 4 : Next
  MushTimer.enabled=true
End Sub


Sub Mush002_Hit()
  Dim BP
  RandomSoundBumperPas BM_Mush2
    Addscore 1000
    If LightC.State=1 then LightCheck
  For each BP in BP_Mush2 : BP.z = 4 : Next
  MushTimer.enabled=true
End Sub


Sub Mush003_Hit()
  Dim BP
  RandomSoundBumperPas BM_Mush3
    Addscore 1000
    If LightD.State=1 then LightCheck
  For each BP in BP_Mush3 : BP.z = 4 : Next
  MushTimer.enabled=true
End Sub

Sub MushTimer_timer
  Dim BP
  For each BP in BP_Mush1 : BP.z = 0 : Next
  For each BP in BP_Mush2 : BP.z = 0 : Next
  For each BP in BP_Mush3 : BP.z = 0 : Next
  For each BP in BP_Mush4 : BP.z = 0 : Next
' BM_Mush1.z=0
' BM_Mush2.z=0
' BM_Mush3.z=0
' BM_Mush4.z=0
  MushTimer.enabled=false
end sub

Sub Kicker1_Hit()
  SoundSaucerLock
    If Light6.State=1 then LightCheck
    If Light6.State=1 then Addscore 3000 Else Addscore 300
    Kick1Timer.enabled=True
End Sub

sub kick1timer_timer
    Kicker1.kick (int(rnd(1)*10)+230),12
    kick1timer.enabled=false
    playsoundAt "KickerKick", Kicker1
  DOF 213, DOFPulse
end sub

Sub Kicker2_Hit()
  SoundSaucerLock
  If tilt=false then
    if Light1.state=1 then AddScore 1000
    If Light2.state=1 then AddScore 2000
    If Light3.state=1 then AddScore 3000
    If Light4.state=1 then AddScore 4000
    If Light5.state=1 then AddScore 5000
    flipclose
  end if

    Kick2Timer.enabled=True
End Sub


sub kick2timer_timer
    Kicker1.kick (int(rnd(1)*10)+230),12
    Kicker2.kick (int(rnd(1)*10)+170),12
    kick2timer.enabled=false
    playsoundAt "KickerKick", Kicker2
  DOF 126, DOFPulse
end sub

Sub swT1_Hit: STHit 1: End Sub

Sub swT2_Hit: STHit 2: End Sub

Sub sw1_Hit()
    Addscore 10
End Sub

Sub sw2_Hit()
    Addscore 10
End Sub

Sub sw3_Hit()
  Addscore 100
  BumperLight1.State=1
  BumperLight3.State=1
End Sub

Sub sw4_Hit()
  Addscore 1000
    If Light7.State=1 then
    credit= credit+1
  If B2SOn Then
    Controller.B2SSetCredits Credit
  end if
    qs=qs+2
    PlaySound SoundFXDOF("knocke",117, DOFPulse, DOFKnocker)
    Light7.State=0
    End If
End Sub

Sub sw5_Hit()
  Addscore 100
  BumperLight2.State=1
End Sub

Sub sw6_Hit()
  Addscore 1000
End Sub

Sub sw7_Hit()
  Addscore 1000
End Sub


Sub LightCheck
  ABCDCounter=ABCDCounter+1
  Select case ABCDCounter
    case 1:
      LightA.state=0
      LightB.state=1
    case 2:
      LightB.state=0
      LightC.state=1
    case 3:
      LightC.state=0
      LightD.state=1
    case 4:
      LightD.state=0
      LightA.state=1
      ABCDCounter=0
      QueenCounter=QueenCounter+1
      If QueenCounter>4 then QueenCounter=4
      If QueenCounter>2 then Light7.state=1
      Select Case QueenCounter
        case 1:
          Light8.state=1
          If B2SOn Then Controller.B2SSetData 81,1
          If VRMode Then :LBgA.visible=1 :L2BgA.visible=1 :End If
        case 2:
          Light9.state=1
          If B2SOn Then Controller.B2SSetData 82,1
          If VRMode Then :LBgB001.visible=1 :L2BgB.visible=1 :End If
        case 3:
          Light10.state=1
          If B2SOn Then Controller.B2SSetData 83,1
          If VRMode Then :LBgC.visible=1 :L2BgC.visible=1 :End If
        case 4:
          Light11.state=1
          If B2SOn Then Controller.B2SSetData 84,1
          If VRMode Then :LBgD.visible=1 :L2BgD.visible=1 :End If
          credit=credit+1
          qs=qs+2
          PlaySound SoundFXDOF("knocke",117, DOFPulse, DOFKnocker)
          if credit>25 then credit=25
          If B2SOn Then
            Controller.B2SSetCredits Credit
          end if
      end select
  end select
End Sub

Sub RubberScoreWall1_Hit()
  Addscore 10
  flipopen
End Sub

Sub RubberScoreWall2_Hit()
  Dim finalspeed
  finalspeed=SQR(activeball.velx * activeball.velx + activeball.vely * activeball.vely)
  If finalspeed > 3 then Addscore 10
End Sub

Sub RubberScoreWall6_Hit()
  Addscore 10
End Sub

Sub RubberScoreWall7_Hit()
  Addscore 10
End Sub

Sub RubberScoreWall5_Hit()
  AddScore 10
end sub

Sub RubberScoreWall4_Hit()
  Addscore 10
  flipopen
End Sub

Sub RubberScoreWall3_Hit()
  Addscore 10
End Sub


sub ballrel_hit
    if ballrelenabled=1 then
    playsound "launchball"
    ballrelenabled=0
    end if
end sub

sub flipclose
  If tilt=false then
    if (rightflipper.enabled)=true then playsound "zclose"
    leftflipper.enabled=false
    leftflipper1.enabled=False
    rightflipper.enabled=false
    rightflipper1.enabled=False
    LeftFlipper.RotateToStart
    LeftFlipper1.RotateToStart
    RightFlipper.RotateToStart
    RightFlipper1.RotateToStart
    BM_Rflip.visible=false
    BM_Rflip1.visible=true
    BM_Lflip.visible=false
    BM_Lflip1.visible=true
    FlipperLSh.visible=false
    FlipperRSh.visible=false
    FlipperLSh1.visible=True
    FlipperRSh1.visible=True
    FlipperPosition=1
    clearBM
    leftflipper1.enabled=true
    rightflipper1.enabled=true
    LeftFlipper1_Animate
    RightFlipper1_Animate
  End If
end sub

sub flipopen
  if (rightflipper.enabled)=false then playsound "zopen"
  leftflipper.enabled=false
    leftflipper1.enabled=False
    rightflipper.enabled=false
    rightflipper1.enabled=False
  LeftFlipper.RotateToStart
  LeftFlipper1.RotateToStart
  RightFlipper.RotateToStart
  RightFlipper1.RotateToStart
    BM_Rflip.visible=true
    BM_Rflip1.visible=false
    BM_Lflip.visible=true
    BM_Lflip1.visible=false
  FlipperLSh.visible=true
  FlipperRSh.visible=true
  FlipperLSh1.visible=False
  FlipperRSh1.visible=False
  FlipperPosition=0
  clearBM
  leftflipper.enabled=true
  rightflipper.enabled=true
  LeftFlipper_Animate
  RightFlipper_Animate
end sub


sub savehs
  ' Based on Black's Highscore routines
  Dim FileObj
  Dim ScoreFile
  Set FileObj=CreateObject("Scripting.FileSystemObject")
  If Not FileObj.FolderExists(UserDirectory) then
  Exit Sub
  End if
  Set ScoreFile=FileObj.CreateTextFile(UserDirectory & TableName& ".txt",True)
  ScoreFile.WriteLine score
    scorefile.writeline credit
  scorefile.writeline matchnumb
  scorefile.writeline hisc
  scorefile.writeline ballspergame
  scorefile.writeline replayset
  ScoreFile.Close
  Set ScoreFile=Nothing
  Set FileObj=Nothing
end sub

sub loadhs
    ' Based on Black's Highscore routines
  Dim FileObj
  Dim ScoreFile
  dim temp1
  dim temp2
  dim temp3
  dim temp4
  dim temp5
  dim temp6
  dim temp7
  Set FileObj=CreateObject("Scripting.FileSystemObject")
  If Not FileObj.FolderExists(UserDirectory) then
    score=10
    credit=0
    matchnumb=10
    hisc=50000
    ballspergame=3
    replayset=3
    ShowReplayLevels
    Exit Sub
  End if
  If Not FileObj.FileExists(UserDirectory & TableName& ".txt") then
  Exit Sub
  End if
  Set ScoreFile=FileObj.GetFile(UserDirectory & TableName& ".txt")
  Set TextStr=ScoreFile.OpenAsTextStream(1,0)
  If (TextStr.AtEndOfStream=True) then
  Exit Sub
  End if
  temp1=TextStr.ReadLine
  temp2=textstr.readline
  temp3=textstr.readline
  temp4=Textstr.ReadLine
  temp5=Textstr.ReadLine
  temp6=Textstr.ReadLine

  TextStr.Close

  if IsNumeric(temp1) then score = clng(temp1) else score = 10
    if IsNumeric(temp2) then credit = clng(temp2) else credit = 0
    if IsNumeric(temp3) then matchnumb = clng(temp3) else matchnumb = 0
    if IsNumeric(temp4) then hisc = clng(temp4) else hisc = 50000
    if IsNumeric(temp5) then ballspergame = clng(temp5) else ballspergame = 3
    if IsNumeric(temp6) then replayset = clng(temp6) else replayset = 3

  Set ScoreFile=Nothing
  Set FileObj=Nothing
  ShowReplayLevels
end sub

'************************************************
'************************************************
'************************************************
'************************************************
'************************************************





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
  Dim b, BOT
  BOT = GetBalls

  ' stop the sound of deleted balls
  For b = UBound(BOT) + 1 to tnob - 1
    ' Comment the next line if you are not implementing Dyanmic Ball Shadows
    If AmbientBallShadowOn = 0 Then BallShadowA(b).visible = 0
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

    ' "Static" Ball Shadows
    ' Comment the next If block, if you are not implementing the Dyanmic Ball Shadows
    If AmbientBallShadowOn = 0 Then
      If BOT(b).Z > 30 Then
        BallShadowA(b).height=BOT(b).z - BallSize/4   'This is technically 1/4 of the ball "above" the ramp, but it keeps it from clipping the ramp
      Else
        BallShadowA(b).height=0.1
      End If
      BallShadowA(b).Y = BOT(b).Y + offsetY
      BallShadowA(b).X = BOT(b).X + offsetX
      BallShadowA(b).visible = 1
    End If
  Next
End Sub


'******************************************************
'****  END BALL ROLLING AND DROP SOUNDS
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
Function dSin(degrees)
  dsin = sin(degrees * Pi/180)
End Function

Function dCos(degrees)
  dcos = cos(degrees * Pi/180)
End Function
'
Function RotPoint(x,y,angle)
  dim rx, ry
  rx = x*dCos(angle) - y*dSin(angle)
  ry = x*dSin(angle) + y*dCos(angle)
  RotPoint = Array(rx,ry)
End Function

Function max(a,b)
  if a > b then
    max = a
  Else
    max = b
  end if
end Function

Dim PI: PI = 4*Atn(1)

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
  Dim BOT : BOT = GetBalls
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
    Dim b, BOT
      BOT = GetBalls

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
Dim ST11, ST12

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


Set ST11 = (new StandupTarget)(swT1, BM_swT1,1, 0)
Set ST12 = (new StandupTarget)(swT2, BM_swT2,2, 0)

'Add all the Stand-up Target Arrays to Stand-up Target Animation Array
'   STAnimationArray = Array(ST1, ST2, ....)
Dim STArray
STArray = Array(ST11, ST12)

'Configure the behavior of Stand-up Targets
Const STAnimStep = 1.5  'vpunits per animation step (control return to Start)
Const STMaxOffset = 9   'max vp units target moves when hit

Const STMass = 0.1    'Mass of the Stand-up Target (between 0 and 1), higher values provide more resistance

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
    If UsingROM Then
      vpmTimer.PulseSw switch mod 100
    Else
      STAction switch
    End If
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


Sub STAction(Switch)
  Select Case Switch
    Case 1
      DOF 120, 2
      If LightT1.State=1 then LightCheck
      If LightT1.State=1 then Addscore 1000 Else Addscore 100
    Case 2
      DOF 120, 2
      If LightT2.State=1 then LightCheck
      If LightT2.State=1 then Addscore 1000 Else Addscore 100
  End Select
End Sub

Sub UpdateStandupTargets
  dim BP, ty

    ty = BM_swT1.transy
  For each BP in BP_swT1 : BP.transy = ty: Next

    ty = BM_swT2.transy
  For each BP in BP_swT2 : BP.transy = ty: Next

End Sub

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

'******************************************************
'****   END STAND-UP TARGETS
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
Dim SlingshotSoundLevel, BumperSoundFactor, KnockerSoundLevel, PassiveBumperSoundLevel

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
SaucerLockSoundLevel = 0.1
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

DrainSoundLevel = 0.8                         'volume level; range [0, 1]
BallReleaseSoundLevel = 1                       'volume level; range [0, 1]
BottomArchBallGuideSoundFactor = 0.2                  'volume multiplier; must not be zero
FlipperBallGuideSoundFactor = 0.015                   'volume multiplier; must not be zero
PassiveBumperSoundLevel = 0.8                     'volume level; range [0, 1]

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

'/////////////////////////////  PASSIVE BUMPER SOUNDS  ////////////////////////////
Sub RandomSoundBumperPas(Bump)
  PlaySoundAtLevelStatic ("Bumpers_passive_" & Int(Rnd*3)+1), PassiveBumperSoundLevel, Bump
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


'/////////////////////////////  COIN IN  ////////////////////////////

Sub SoundCoinIn()
  Select Case Int(rnd*3)
    Case 0: PlaySound ("Coin_In_1"), 0, CoinSoundLevel, 0, 0.25
    Case 1: PlaySound ("Coin_In_2"), 0, CoinSoundLevel, 0, 0.25
    Case 2: PlaySound ("Coin_In_3"), 0, CoinSoundLevel, 0, 0.25
  End Select
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


Sub TriggerG_Hit
  RandomSoundMetal
End Sub


'/////////////////////////////////////////////////////////////////
'         End Mechanical Sounds
'/////////////////////////////////////////////////////////////////

'******************************************************
'****  FLEEP MECHANICAL SOUNDS
'******************************************************

'*********************************************************************
'                 Positional Sound Playback Functions
'*********************************************************************

Sub PlayLoopSoundAtVol(sound, tableobj, Vol)
  PlaySound sound, -1, Vol, AudioPan(tableobj), 0, 0, 1, 0, AudioFade(tableobj)
End Sub



'******************************************************
'   Misc Animations
'******************************************************

Sub sw1_Animate
  Dim z : z = sw1.CurrentAnimOffset
  Dim BP : For Each BP in BP_sw1 : BP.transz = z: Next
End Sub

Sub sw2_Animate
  Dim z : z = sw2.CurrentAnimOffset
  Dim BP : For Each BP in BP_sw2 : BP.transz = z: Next
End Sub

Sub sw3_Animate
  Dim z : z = sw3.CurrentAnimOffset
  Dim BP : For Each BP in BP_sw3 : BP.transz = z: Next
End Sub

Sub sw4_Animate
  Dim z : z = sw4.CurrentAnimOffset
  Dim BP : For Each BP in BP_sw4 : BP.transz = z: Next
End Sub

Sub sw5_Animate
  Dim z : z = sw5.CurrentAnimOffset
  Dim BP : For Each BP in BP_sw5 : BP.transz = z: Next
End Sub

Sub sw6_Animate
  Dim z : z = sw6.CurrentAnimOffset
  Dim BP : For Each BP in BP_sw6 : BP.transz = z: Next
End Sub

Sub sw7_Animate
  Dim z : z = sw7.CurrentAnimOffset
  Dim BP : For Each BP in BP_sw7 : BP.transz = z: Next
End Sub

Sub Gate_Animate
  Dim a : a = Gate.CurrentAngle
  Dim BL : For Each BL in BP_Gate : BL.rotx = a: Next
End Sub



'**************
' Bumper Animations
'**************

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
    Dim r, s, dx, dy, dist, force
    Dim BOT : BOT = GetBalls

    For r = 0 To 2

        dx = 0
        dy = 0
        dist = 99999

        ' Buscar la bola más cercana
        For s = 0 To UBound(BOT)
            Dim tx, ty, d
            tx = BOT(s).x - Bumpers(r).x
            ty = BOT(s).y - Bumpers(r).y
            d = Sqr(tx*tx + ty*ty)

            If d < dist Then
                dist = d
                dx = tx
                dy = ty
            End If
        Next

        ' Intensidad del efecto
        If dist < 80 Then
            force = (80 - dist) / 80   ' 0..1
        Else
            force = 0
        End If

        ' Normalizar vector
        If dist > 0 Then
            dx = dx / dist
            dy = dy / dist
        End If

        ' Parámetros visuales
        Dim tilt : tilt = 12 * force   ' grados de inclinación
        Dim zoff : zoff = 1 + 5 * force

        Select Case r
            Case 0
                For Each s In BP_BS1
                    s.RotX = -dy * tilt
                    s.RotY =  dx * tilt
                    s.Z    = zoff
                Next
            Case 1
                For Each s In BP_BS2
                    s.RotX = -dy * tilt
                    s.RotY =  dx * tilt
                    s.Z    = zoff
                Next
            Case 2
                For Each s In BP_BS3
                    s.RotX = -dy * tilt
                    s.RotY =  dx * tilt
                    s.Z    = zoff
                Next
        End Select

    Next
End Sub



'**************
' Flippers animation
'**************


Sub LeftFlipper_Animate
' If Leftflipper.enabled Then
  If FlipperPosition = 0 Then
    Dim a : a = LeftFlipper.CurrentAngle
    FlipperLSh.RotZ = a

    Dim v, BP
    v = 255.0 * (121.0 - LeftFlipper.CurrentAngle) / (121.0 -  70.0)

    For each BP in BP_Lflip
      BP.Rotz = a
      BP.visible = v < 128.0
    Next
    For each BP in BP_LflipU
      BP.Rotz = a
      BP.visible = v >= 128.0
    Next
  End If
End Sub

Sub RightFlipper_Animate
' If Rightflipper.enabled Then
  If FlipperPosition = 0 Then
    Dim a : a = RightFlipper.CurrentAngle
    FlipperRSh.RotZ = a

    Dim v, BP
    v = 255.0 * (-121.0 - RightFlipper.CurrentAngle) / (-121.0 +  70.0)

    For each BP in BP_Rflip
      BP.Rotz = a
      BP.visible = v < 128.0
    Next
    For each BP in BP_RflipU
      BP.Rotz = a
      BP.visible = v >= 128.0
    Next
  End If
End Sub


Sub LeftFlipper1_Animate
' If Leftflipper1.enabled Then
  If FlipperPosition = 1 Then
    Dim a : a = LeftFlipper1.CurrentAngle
    FlipperLSh1.RotZ = a

    Dim v, BP
    v = 255.0 * (121.0 - LeftFlipper1.CurrentAngle) / (121.0 -  70.0)

    For each BP in BP_Lflip1
      BP.Rotz = a
      BP.visible = v < 128.0
    Next
    For each BP in BP_Lflip1U
      BP.Rotz = a
      BP.visible = v >= 128.0
    Next
  End If
End Sub

Sub RightFlipper1_Animate
' If Rightflipper1.enabled Then
  If FlipperPosition = 1 Then
    Dim a : a = RightFlipper1.CurrentAngle
    FlipperRSh1.RotZ = a

    Dim v, BP
    v = 255.0 * (-121.0 - RightFlipper1.CurrentAngle) / (-121.0 +  70.0)

    For each BP in BP_Rflip1
      BP.Rotz = a
      BP.visible = v < 128.0
    Next
    For each BP in BP_Rflip1U
      BP.Rotz = a
      BP.visible = v >= 128.0
    Next
  End If
End Sub


Sub clearBM
  If FlipperPosition = 0 Then
'   Dim a : a = LeftFlipper1.CurrentAngle
'     FlipperLSh1.RotZ = a
      Dim BP
      For each BP in BP_Lflip1
'       BP.Rotz = a
        BP.visible = 0
      Next
      For each BP in BP_Lflip1U
'       BP.Rotz = a
        BP.visible = 0
      Next
      For each BP in BP_Rflip1
'       BP.Rotz = a
        BP.visible = 0
      Next
      For each BP in BP_Rflip1U
'       BP.Rotz = a
        BP.visible = 0
      Next
      BM_Lflip.visible = 1
      BM_Rflip.visible = 1
    Else
      For each BP in BP_Lflip
'       BP.Rotz = a
        BP.visible = 0
      Next
      For each BP in BP_LflipU
'       BP.Rotz = a
        BP.visible = 0
      Next
      For each BP in BP_Rflip
'       BP.Rotz = a
        BP.visible = 0
      Next
      For each BP in BP_RflipU
'       BP.Rotz = a
        BP.visible = 0
      Next
      BM_Lflip1.visible = 1
      BM_Rflip1.visible = 1
  End If
End Sub



'******************************************************
'**    VR Backglass Lighting & Reels by Nestorgian   **
'******************************************************

Dim Object
Dim BGObj
Dim ones
Dim tens
Dim huns
Dim thous
Dim tenthous
Dim ReelValue(5)     ' almacena el dígito actual de cada reel (0–9)
Dim ResetInProgress
Dim ResetSpeed: ResetSpeed = 150          ' velocidad del reseteo (ms entre pulsos)
Dim ResetSoundName :  ResetSoundName = "Reelclick"    ' nombre del sonido a reproducir
Dim SndReelZeroHit : SndReelZeroHit = "reel_coil_resonant"

sub setreel(nscore)
    ones = (nscore Mod 10) * -32.7272
    tens = (int((nscore Mod 100) / 10)) * -32.7272
    huns = (int((nscore Mod 1000) / 100)) * -32.7272
    thous = (int((nscore Mod 10000) / 1000)) * -32.7272
    tenthous = (int((nscore Mod 100000) / 10000)) * -32.7272
  emp1r5.ObjrotX = ones
    emp1r4.ObjrotX = tens
    emp1r3.ObjrotX = huns
    emp1r2.ObjrotX = thous
    emp1r1.ObjrotX = tenthous
  ReelValue(5) = (nscore Mod 10)
  ReelValue(4) = Int((nscore Mod 100) / 10)
  ReelValue(3) = Int((nscore Mod 1000) / 100)
  ReelValue(2) = Int((nscore Mod 10000) / 1000)
  ReelValue(1) = Int((nscore Mod 100000) / 10000)
end sub



Sub StartResetReels()
    ResetInProgress = True
    ResetReelsTimer.Interval = ResetSpeed
    ResetReelsTimer.Enabled = True
End Sub


Sub ResetReelsTimer_Timer()
    Dim i
    Dim AllZero : AllZero = True
  Dim PlayedSound : PlayedSound = False

    For i = 1 To 5
        If ReelValue(i) <> 0 Then
            ' avanza un paso igual que un carrete Bally
            ReelValue(i) = (ReelValue(i) + 1) Mod 10
            ' aplicar la rotación
            Select Case i
                Case 1: emp1r1.ObjRotX = ReelValue(i) * -32.7272
                Case 2: emp1r2.ObjRotX = ReelValue(i) * -32.7272
                Case 3: emp1r3.ObjRotX = ReelValue(i) * -32.7272
                Case 4: emp1r4.ObjRotX = ReelValue(i) * -32.7272
                Case 5: emp1r5.ObjRotX = ReelValue(i) * -32.7272
            End Select
      If Not PlayedSound Then
                PlaySound ResetSoundName, 0, 1 + Rnd() * 0.1
                PlayedSound = True
            End If
      ' 5. Si cae justo en la posición “0”, hace golpe fuerte
            If ReelValue(i) = 0 Then
        PlayReelZeroHitSound()
            End If

            AllZero = False   ' todavía falta resetear
        End If
    Next
    ' Cuando todos están en 0, se termina el reset
    If AllZero Then
        ResetReelsTimer.Enabled = False
        ResetInProgress = False
    End If
End Sub

Sub PlayReelZeroHitSound()
    PlaySound SndReelZeroHit, 0, 1 + Rnd() * 0.1
End Sub


'sub setcreditreel(creds)
' if creds > 9 then creds = 9
'    emp4r6.ObjrotX = creds * -32.7272
'end sub

Const MAX_CREDITS = 25
Const REEL_STEP = -13.84615   ' 360 / 26

Sub SetCreditReel(creds)
    If creds < 0 Then creds = 0
    If creds > MAX_CREDITS Then creds = MAX_CREDITS
    emp4r6.ObjRotX = creds * REEL_STEP
End Sub


Sub SetBackglass()
  For Each BGObj In BG_flashers
    BGObj.x = BGobj.x
    BGObj.height = - BGObj.y '+ 6
    BGObj.y =-70 'adjusts the distance from the backglass towards the user
  Next
End Sub

Sub SetBackglassLights()
  For Each BGObj In BG_flashersL
    BGObj.x = BGobj.x + 2
    BGObj.height = - BGObj.y - 8
    BGObj.y =-60  'adjusts the distance from the backglass towards the user
  Next
End Sub

Sub FlasherPlayer
  For each Object in VRBGPlayer1 : object.visible = 1 : Next
End Sub


Sub FlasherMatch
  If matchnumb = 0 Then FlM00.visible = 1 : FlM00A.visible = 1 : FlM00B.visible = 1 Else FlM00.visible = 0 : FlM00A.visible = 0 : FlM00B.visible = 0  End If
  If matchnumb = 1 Then FlM10.visible = 1 : FlM10A.visible = 1 : FlM10B.visible = 1 Else FlM10.visible = 0 : FlM10A.visible = 0 : FlM10B.visible = 0 End If
  If matchnumb = 2 Then FlM20.visible = 1 : FlM20A.visible = 1 : FlM20B.visible = 1 Else FlM20.visible = 0 : FlM20A.visible = 0 : FlM20B.visible = 0 End If
  If matchnumb = 3 Then FlM30.visible = 1 : FlM30A.visible = 1 : FlM30B.visible = 1 Else FlM30.visible = 0 : FlM30A.visible = 0 : FlM30B.visible = 0 End If
  If matchnumb = 4 Then FlM40.visible = 1 : FlM40A.visible = 1 : FlM40B.visible = 1 Else FlM40.visible = 0 : FlM40A.visible = 0 : FlM40B.visible = 0 End If
  If matchnumb = 5 Then FlM50.visible = 1 : FlM50A.visible = 1 : FlM50B.visible = 1 Else FlM50.visible = 0 : FlM50A.visible = 0 : FlM50B.visible = 0 End If
  If matchnumb = 6 Then FlM60.visible = 1 : FlM60A.visible = 1 : FlM60B.visible = 1 Else FlM60.visible = 0 : FlM60A.visible = 0 : FlM60B.visible = 0 End If
  If matchnumb = 7 Then FlM70.visible = 1 : FlM70A.visible = 1 : FlM70B.visible = 1 Else FlM70.visible = 0 : FlM70A.visible = 0 : FlM70B.visible = 0 End If
  If matchnumb = 8 Then FlM80.visible = 1 : FlM80A.visible = 1 : FlM80B.visible = 1 Else FlM80.visible = 0 : FlM80A.visible = 0 : FlM80B.visible = 0 End If
  If matchnumb = 9 Then FlM90.visible = 1 : FlM90A.visible = 1 : FlM90B.visible = 1 Else FlM90.visible = 0 : FlM90A.visible = 0 : FlM90B.visible = 0 End If
End Sub

Sub FlasherBalls
  If BallInPlay = 1 Then FlBIP1.visible = 1 : FlBIP1A.visible = 1 Else FlBIP1.visible = 0 : FlBIP1A.visible = 0 End If
  If BallInPlay = 2 Then FlBIP2.visible = 1 : FlBIP2A.visible = 1 Else FlBIP2.visible = 0 : FlBIP2A.visible = 0 End If
  If BallInPlay = 3 Then FlBIP3.visible = 1 : FlBIP3A.visible = 1 Else FlBIP3.visible = 0 : FlBIP3A.visible = 0 End If
  If ballspergame =5 Then
    If BallInPlay = 4 Then FlBIP4.visible = 1 : FlBIP4A.visible = 1 Else FlBIP4.visible = 0 : FlBIP4A.visible = 0 End If
    If BallInPlay = 5 Then FlBIP5.visible = 1 : FlBIP5A.visible = 1 Else FlBIP5.visible = 0 : FlBIP5A.visible = 0 End If
  End If
End Sub


Sub BackglassTimer_Timer()
    Dim r, failure
  failure = (Int(Rnd * 500) = 0)
    r = Int(Rnd * 7) + 1   ' 1 a 4
    Pincab_BgL.Image = "sin" & r
  If failure Then
    Pincab_BackglassL.visible= 0
  Else
    Pincab_BackglassL.visible= 1
  End If

End Sub


Sub SetupRoom
  If RenderingMode = 2 or Table1.ShowFSS = True or VRTest = 1 Then
    VRMode = True
'   SetBackglass
'   SetBackglassLights
    For Each VR_Obj in VRCabinet : VR_Obj.Visible = 1 : Next
    For Each VR_Obj in VRReel1 : VR_Obj.Visible = 1 : Next
    If VRRoomChoice = 1 Then
      For Each VR_Obj in VRMinimalRoom : VR_Obj.Visible = 0 : Next
      For Each VR_Obj in VRMegaRoom : VR_Obj.Visible = 1 : Next
    Else
      For Each VR_Obj in VRMinimalRoom : VR_Obj.Visible = 1 : Next
      For Each VR_Obj in VRMegaRoom : VR_Obj.Visible = 0 : Next
    End If
  Else
    VRMode = False
    For Each VR_Obj in VRReel1 : VR_Obj.Visible = 0 : Next
    For Each VR_Obj in VRCabinet : VR_Obj.Visible = 0 : Next
    For Each VR_Obj in VRMinimalRoom : VR_Obj.Visible = 0 : Next
    For Each VR_Obj in VRMegaRoom : VR_Obj.Visible = 0 : Next
    For Each VR_Obj In BG_flashers : VR_Obj.Visible = 0 : Next
    For Each VR_Obj In BG_flashersL : VR_Obj.Visible = 0 : Next
    If DesktopMode Then
        Pincab_Siderails.Visible = 1
        Pincab_Lockdownbar.Visible = 1
    End If
  End If
End Sub


' VLM  Arrays - Start
' Arrays per baked part
Dim BP_BR1: BP_BR1=Array(BM_BR1, LM_GI_GI07_BR1, LM_L_bumperlight1_BR1)
Dim BP_BR2: BP_BR2=Array(BM_BR2, LM_L_bumperlight2_BR2)
Dim BP_BR3: BP_BR3=Array(BM_BR3, LM_GI_GI11_BR3, LM_GI_GI14_BR3, LM_L_bumperlight3_BR3)
Dim BP_BS1: BP_BS1=Array(BM_BS1, LM_GI_GI06_BS1, LM_GI_GI07_BS1, LM_GI_GI11_BS1, LM_GI_GI12_BS1, LM_GI_GI13_BS1, LM_GI_GI15_BS1, LM_GI_GI19_BS1, LM_L_bumperlight1_BS1, LM_L_bumperlight2_BS1)
Dim BP_BS2: BP_BS2=Array(BM_BS2, LM_GI_GI07_BS2, LM_GI_GI09_BS2, LM_GI_GI12_BS2, LM_GI_GI14_BS2, LM_GI_GI15_BS2, LM_GI_GI16_BS2, LM_GI_GI18_BS2, LM_GI_GI20_BS2, LM_L_bumperlight1_BS2, LM_L_bumperlight2_BS2, LM_L_bumperlight3_BS2)
Dim BP_BS3: BP_BS3=Array(BM_BS3, LM_GI_GI06_BS3, LM_GI_GI07_BS3, LM_GI_GI08_BS3, LM_GI_GI10_BS3, LM_GI_GI11_BS3, LM_GI_GI14_BS3, LM_GI_GI15_BS3, LM_GI_GI16_BS3, LM_GI_GI18_BS3, LM_GI_GI19_BS3, LM_L_bumperlight2_BS3, LM_L_bumperlight3_BS3)
Dim BP_BumperCap: BP_BumperCap=Array(BM_BumperCap, LM_GI_GI06_BumperCap, LM_GI_GI07_BumperCap, LM_GI_GI11_BumperCap, LM_GI_GI12_BumperCap, LM_GI_GI14_BumperCap, LM_GI_GI15_BumperCap, LM_GI_GI16_BumperCap, LM_L_LightC_BumperCap, LM_L_bumperlight1_BumperCap, LM_L_bumperlight2_BumperCap, LM_L_bumperlight3_BumperCap)
Dim BP_Gate: BP_Gate=Array(BM_Gate)
Dim BP_LEMK: BP_LEMK=Array(BM_LEMK, LM_GI_GI03_LEMK, LM_GI_GI04_LEMK)
Dim BP_LSling1: BP_LSling1=Array(BM_LSling1, LM_GI_GI01_LSling1, LM_GI_GI03_LSling1, LM_GI_GI04_LSling1, LM_L_Light8_LSling1)
Dim BP_LSling2: BP_LSling2=Array(BM_LSling2, LM_GI_GI01_LSling2, LM_GI_GI03_LSling2, LM_GI_GI04_LSling2, LM_L_Light8_LSling2)
Dim BP_Lflip: BP_Lflip=Array(BM_Lflip, LM_GI_GI01_Lflip, LM_GI_GI02_Lflip, LM_GI_GI03_Lflip, LM_GI_GI04_Lflip)
Dim BP_Lflip1: BP_Lflip1=Array(BM_Lflip1, LM_GI_GI01_Lflip1, LM_GI_GI02_Lflip1, LM_GI_GI03_Lflip1, LM_GI_GI04_Lflip1)
Dim BP_Lflip1U: BP_Lflip1U=Array(BM_Lflip1U, LM_GI_GI01_Lflip1U, LM_GI_GI02_Lflip1U, LM_GI_GI03_Lflip1U, LM_GI_GI04_Lflip1U)
Dim BP_LflipU: BP_LflipU=Array(BM_LflipU, LM_GI_GI01_LflipU, LM_GI_GI02_LflipU, LM_GI_GI03_LflipU, LM_GI_GI04_LflipU)
Dim BP_Mush1: BP_Mush1=Array(BM_Mush1, LM_GI_GI12_Mush1, LM_GI_GI13_Mush1)
Dim BP_Mush2: BP_Mush2=Array(BM_Mush2, LM_GI_GI14_Mush2, LM_GI_GI15_Mush2, LM_GI_GI16_Mush2)
Dim BP_Mush3: BP_Mush3=Array(BM_Mush3, LM_GI_GI05_Mush3, LM_GI_GI06_Mush3, LM_GI_GI07_Mush3)
Dim BP_Mush4: BP_Mush4=Array(BM_Mush4, LM_GI_GI08_Mush4, LM_GI_GI09_Mush4, LM_GI_GI10_Mush4, LM_GI_GI11_Mush4)
Dim BP_Parts: BP_Parts=Array(BM_Parts, LM_GI_GI01_Parts, LM_GI_GI02_Parts, LM_GI_GI03_Parts, LM_GI_GI04_Parts, LM_GI_GI05_Parts, LM_GI_GI06_Parts, LM_GI_GI07_Parts, LM_GI_GI08_Parts, LM_GI_GI09_Parts, LM_GI_GI10_Parts, LM_GI_GI11_Parts, LM_GI_GI12_Parts, LM_GI_GI13_Parts, LM_GI_GI14_Parts, LM_GI_GI15_Parts, LM_GI_GI16_Parts, LM_GI_GI17_Parts, LM_GI_GI18_Parts, LM_GI_GI19_Parts, LM_GI_GI20_Parts, LM_GI_GI21_Parts, LM_GI_GI22_Parts, LM_GI_GI23_Parts, LM_GI_GI24_Parts, LM_L_Light19_Parts, LM_L_Light5_Parts, LM_L_Light6_Parts, LM_L_Light7_Parts, LM_L_Light8_Parts, LM_L_LightA_Parts, LM_L_LightB_Parts, LM_L_LightC_Parts, LM_L_LightD_Parts, LM_L_LightT1_Parts, LM_L_LightT2_Parts, LM_L_bumperlight1_Parts, LM_L_bumperlight2_Parts, LM_L_bumperlight3_Parts)
Dim BP_Plaastics: BP_Plaastics=Array(BM_Plaastics, LM_GI_GI01_Plaastics, LM_GI_GI02_Plaastics, LM_GI_GI03_Plaastics, LM_GI_GI04_Plaastics, LM_GI_GI05_Plaastics, LM_GI_GI06_Plaastics, LM_GI_GI07_Plaastics, LM_GI_GI08_Plaastics, LM_GI_GI09_Plaastics, LM_GI_GI10_Plaastics, LM_GI_GI11_Plaastics, LM_GI_GI12_Plaastics, LM_GI_GI13_Plaastics, LM_GI_GI14_Plaastics, LM_GI_GI15_Plaastics, LM_GI_GI16_Plaastics, LM_GI_GI17_Plaastics, LM_GI_GI20_Plaastics, LM_L_Light19_Plaastics, LM_L_Light6_Plaastics, LM_L_LightA_Plaastics, LM_L_LightT1_Plaastics, LM_L_LightT2_Plaastics, LM_L_bumperlight1_Plaastics)
Dim BP_Playfield: BP_Playfield=Array(BM_Playfield, LM_GI_GI01_Playfield, LM_GI_GI02_Playfield, LM_GI_GI03_Playfield, LM_GI_GI04_Playfield, LM_GI_GI05_Playfield, LM_GI_GI06_Playfield, LM_GI_GI07_Playfield, LM_GI_GI08_Playfield, LM_GI_GI09_Playfield, LM_GI_GI10_Playfield, LM_GI_GI11_Playfield, LM_GI_GI12_Playfield, LM_GI_GI13_Playfield, LM_GI_GI14_Playfield, LM_GI_GI15_Playfield, LM_GI_GI16_Playfield, LM_GI_GI17_Playfield, LM_GI_GI18_Playfield, LM_GI_GI19_Playfield, LM_GI_GI20_Playfield, LM_GI_GI21_Playfield, LM_GI_GI22_Playfield, LM_GI_GI23_Playfield, LM_GI_GI24_Playfield, LM_L_Light10_Playfield, LM_L_Light11_Playfield, LM_L_Light19_Playfield, LM_L_Light1_Playfield, LM_L_Light2_Playfield, LM_L_Light3_Playfield, LM_L_Light4_Playfield, LM_L_Light5_Playfield, LM_L_Light6_Playfield, LM_L_Light7_Playfield, LM_L_Light8_Playfield, LM_L_Light9_Playfield, LM_L_LightA_Playfield, LM_L_LightB_Playfield, LM_L_LightC_Playfield, LM_L_LightD_Playfield, LM_L_LightT1_Playfield, LM_L_LightT2_Playfield, _
  LM_L_bumperlight1_Playfield, LM_L_bumperlight2_Playfield, LM_L_bumperlight3_Playfield)
Dim BP_REMK: BP_REMK=Array(BM_REMK, LM_GI_GI01_REMK, LM_GI_GI02_REMK)
Dim BP_RSling1: BP_RSling1=Array(BM_RSling1, LM_GI_GI01_RSling1, LM_GI_GI02_RSling1, LM_GI_GI03_RSling1)
Dim BP_RSling2: BP_RSling2=Array(BM_RSling2, LM_GI_GI01_RSling2, LM_GI_GI02_RSling2, LM_GI_GI03_RSling2)
Dim BP_Rflip: BP_Rflip=Array(BM_Rflip, LM_GI_GI01_Rflip, LM_GI_GI02_Rflip, LM_GI_GI03_Rflip, LM_GI_GI04_Rflip)
Dim BP_Rflip1: BP_Rflip1=Array(BM_Rflip1, LM_GI_GI01_Rflip1, LM_GI_GI02_Rflip1, LM_GI_GI03_Rflip1, LM_GI_GI04_Rflip1)
Dim BP_Rflip1U: BP_Rflip1U=Array(BM_Rflip1U, LM_GI_GI01_Rflip1U, LM_GI_GI02_Rflip1U, LM_GI_GI03_Rflip1U, LM_GI_GI04_Rflip1U)
Dim BP_RflipU: BP_RflipU=Array(BM_RflipU, LM_GI_GI01_RflipU, LM_GI_GI02_RflipU, LM_GI_GI03_RflipU, LM_GI_GI04_RflipU)
Dim BP_UnderPF: BP_UnderPF=Array(BM_UnderPF, LM_GI_GI05_UnderPF, LM_GI_GI09_UnderPF, LM_GI_GI12_UnderPF, LM_GI_GI16_UnderPF, LM_GI_GI18_UnderPF, LM_GI_GI19_UnderPF, LM_GI_GI20_UnderPF, LM_L_Light10_UnderPF, LM_L_Light11_UnderPF, LM_L_Light1_UnderPF, LM_L_Light2_UnderPF, LM_L_Light3_UnderPF, LM_L_Light4_UnderPF, LM_L_Light5_UnderPF, LM_L_Light6_UnderPF, LM_L_Light7_UnderPF, LM_L_Light8_UnderPF, LM_L_Light9_UnderPF, LM_L_LightA_UnderPF, LM_L_LightB_UnderPF, LM_L_LightC_UnderPF, LM_L_LightD_UnderPF, LM_L_LightT1_UnderPF, LM_L_LightT2_UnderPF)
Dim BP_sw1: BP_sw1=Array(BM_sw1, LM_GI_GI17_sw1, LM_GI_GI19_sw1, LM_L_Light7_sw1)
Dim BP_sw2: BP_sw2=Array(BM_sw2, LM_GI_GI13_sw2, LM_GI_GI15_sw2, LM_GI_GI16_sw2, LM_GI_GI19_sw2, LM_L_Light19_sw2)
Dim BP_sw3: BP_sw3=Array(BM_sw3, LM_GI_GI13_sw3, LM_GI_GI17_sw3, LM_GI_GI18_sw3)
Dim BP_sw4: BP_sw4=Array(BM_sw4, LM_GI_GI18_sw4, LM_GI_GI19_sw4)
Dim BP_sw5: BP_sw5=Array(BM_sw5, LM_GI_GI20_sw5, LM_L_Light19_sw5)
Dim BP_sw6: BP_sw6=Array(BM_sw6, LM_GI_GI04_sw6)
Dim BP_sw7: BP_sw7=Array(BM_sw7, LM_GI_GI02_sw7, LM_GI_GI08_sw7)
Dim BP_swT1: BP_swT1=Array(BM_swT1, LM_GI_GI05_swT1, LM_GI_GI06_swT1, LM_L_LightT1_swT1)
Dim BP_swT2: BP_swT2=Array(BM_swT2, LM_GI_GI08_swT2, LM_GI_GI09_swT2, LM_GI_GI10_swT2, LM_GI_GI11_swT2, LM_L_LightT2_swT2)
' Arrays per lighting scenario
Dim BL_GI_GI01: BL_GI_GI01=Array(LM_GI_GI01_LSling1, LM_GI_GI01_LSling2, LM_GI_GI01_Lflip, LM_GI_GI01_Lflip1, LM_GI_GI01_Lflip1U, LM_GI_GI01_LflipU, LM_GI_GI01_Parts, LM_GI_GI01_Plaastics, LM_GI_GI01_Playfield, LM_GI_GI01_REMK, LM_GI_GI01_RSling1, LM_GI_GI01_RSling2, LM_GI_GI01_Rflip, LM_GI_GI01_Rflip1, LM_GI_GI01_Rflip1U, LM_GI_GI01_RflipU)
Dim BL_GI_GI02: BL_GI_GI02=Array(LM_GI_GI02_Lflip, LM_GI_GI02_Lflip1, LM_GI_GI02_Lflip1U, LM_GI_GI02_LflipU, LM_GI_GI02_Parts, LM_GI_GI02_Plaastics, LM_GI_GI02_Playfield, LM_GI_GI02_REMK, LM_GI_GI02_RSling1, LM_GI_GI02_RSling2, LM_GI_GI02_Rflip, LM_GI_GI02_Rflip1, LM_GI_GI02_Rflip1U, LM_GI_GI02_RflipU, LM_GI_GI02_sw7)
Dim BL_GI_GI03: BL_GI_GI03=Array(LM_GI_GI03_LEMK, LM_GI_GI03_LSling1, LM_GI_GI03_LSling2, LM_GI_GI03_Lflip, LM_GI_GI03_Lflip1, LM_GI_GI03_Lflip1U, LM_GI_GI03_LflipU, LM_GI_GI03_Parts, LM_GI_GI03_Plaastics, LM_GI_GI03_Playfield, LM_GI_GI03_RSling1, LM_GI_GI03_RSling2, LM_GI_GI03_Rflip, LM_GI_GI03_Rflip1, LM_GI_GI03_Rflip1U, LM_GI_GI03_RflipU)
Dim BL_GI_GI04: BL_GI_GI04=Array(LM_GI_GI04_LEMK, LM_GI_GI04_LSling1, LM_GI_GI04_LSling2, LM_GI_GI04_Lflip, LM_GI_GI04_Lflip1, LM_GI_GI04_Lflip1U, LM_GI_GI04_LflipU, LM_GI_GI04_Parts, LM_GI_GI04_Plaastics, LM_GI_GI04_Playfield, LM_GI_GI04_Rflip, LM_GI_GI04_Rflip1, LM_GI_GI04_Rflip1U, LM_GI_GI04_RflipU, LM_GI_GI04_sw6)
Dim BL_GI_GI05: BL_GI_GI05=Array(LM_GI_GI05_Mush3, LM_GI_GI05_Parts, LM_GI_GI05_Plaastics, LM_GI_GI05_Playfield, LM_GI_GI05_UnderPF, LM_GI_GI05_swT1)
Dim BL_GI_GI06: BL_GI_GI06=Array(LM_GI_GI06_BS1, LM_GI_GI06_BS3, LM_GI_GI06_BumperCap, LM_GI_GI06_Mush3, LM_GI_GI06_Parts, LM_GI_GI06_Plaastics, LM_GI_GI06_Playfield, LM_GI_GI06_swT1)
Dim BL_GI_GI07: BL_GI_GI07=Array(LM_GI_GI07_BR1, LM_GI_GI07_BS1, LM_GI_GI07_BS2, LM_GI_GI07_BS3, LM_GI_GI07_BumperCap, LM_GI_GI07_Mush3, LM_GI_GI07_Parts, LM_GI_GI07_Plaastics, LM_GI_GI07_Playfield)
Dim BL_GI_GI08: BL_GI_GI08=Array(LM_GI_GI08_BS3, LM_GI_GI08_Mush4, LM_GI_GI08_Parts, LM_GI_GI08_Plaastics, LM_GI_GI08_Playfield, LM_GI_GI08_sw7, LM_GI_GI08_swT2)
Dim BL_GI_GI09: BL_GI_GI09=Array(LM_GI_GI09_BS2, LM_GI_GI09_Mush4, LM_GI_GI09_Parts, LM_GI_GI09_Plaastics, LM_GI_GI09_Playfield, LM_GI_GI09_UnderPF, LM_GI_GI09_swT2)
Dim BL_GI_GI10: BL_GI_GI10=Array(LM_GI_GI10_BS3, LM_GI_GI10_Mush4, LM_GI_GI10_Parts, LM_GI_GI10_Plaastics, LM_GI_GI10_Playfield, LM_GI_GI10_swT2)
Dim BL_GI_GI11: BL_GI_GI11=Array(LM_GI_GI11_BR3, LM_GI_GI11_BS1, LM_GI_GI11_BS3, LM_GI_GI11_BumperCap, LM_GI_GI11_Mush4, LM_GI_GI11_Parts, LM_GI_GI11_Plaastics, LM_GI_GI11_Playfield, LM_GI_GI11_swT2)
Dim BL_GI_GI12: BL_GI_GI12=Array(LM_GI_GI12_BS1, LM_GI_GI12_BS2, LM_GI_GI12_BumperCap, LM_GI_GI12_Mush1, LM_GI_GI12_Parts, LM_GI_GI12_Plaastics, LM_GI_GI12_Playfield, LM_GI_GI12_UnderPF)
Dim BL_GI_GI13: BL_GI_GI13=Array(LM_GI_GI13_BS1, LM_GI_GI13_Mush1, LM_GI_GI13_Parts, LM_GI_GI13_Plaastics, LM_GI_GI13_Playfield, LM_GI_GI13_sw2, LM_GI_GI13_sw3)
Dim BL_GI_GI14: BL_GI_GI14=Array(LM_GI_GI14_BR3, LM_GI_GI14_BS2, LM_GI_GI14_BS3, LM_GI_GI14_BumperCap, LM_GI_GI14_Mush2, LM_GI_GI14_Parts, LM_GI_GI14_Plaastics, LM_GI_GI14_Playfield)
Dim BL_GI_GI15: BL_GI_GI15=Array(LM_GI_GI15_BS1, LM_GI_GI15_BS2, LM_GI_GI15_BS3, LM_GI_GI15_BumperCap, LM_GI_GI15_Mush2, LM_GI_GI15_Parts, LM_GI_GI15_Plaastics, LM_GI_GI15_Playfield, LM_GI_GI15_sw2)
Dim BL_GI_GI16: BL_GI_GI16=Array(LM_GI_GI16_BS2, LM_GI_GI16_BS3, LM_GI_GI16_BumperCap, LM_GI_GI16_Mush2, LM_GI_GI16_Parts, LM_GI_GI16_Plaastics, LM_GI_GI16_Playfield, LM_GI_GI16_UnderPF, LM_GI_GI16_sw2)
Dim BL_GI_GI17: BL_GI_GI17=Array(LM_GI_GI17_Parts, LM_GI_GI17_Plaastics, LM_GI_GI17_Playfield, LM_GI_GI17_sw1, LM_GI_GI17_sw3)
Dim BL_GI_GI18: BL_GI_GI18=Array(LM_GI_GI18_BS2, LM_GI_GI18_BS3, LM_GI_GI18_Parts, LM_GI_GI18_Playfield, LM_GI_GI18_UnderPF, LM_GI_GI18_sw3, LM_GI_GI18_sw4)
Dim BL_GI_GI19: BL_GI_GI19=Array(LM_GI_GI19_BS1, LM_GI_GI19_BS3, LM_GI_GI19_Parts, LM_GI_GI19_Playfield, LM_GI_GI19_UnderPF, LM_GI_GI19_sw1, LM_GI_GI19_sw2, LM_GI_GI19_sw4)
Dim BL_GI_GI20: BL_GI_GI20=Array(LM_GI_GI20_BS2, LM_GI_GI20_Parts, LM_GI_GI20_Plaastics, LM_GI_GI20_Playfield, LM_GI_GI20_UnderPF, LM_GI_GI20_sw5)
Dim BL_GI_GI21: BL_GI_GI21=Array(LM_GI_GI21_Parts, LM_GI_GI21_Playfield)
Dim BL_GI_GI22: BL_GI_GI22=Array(LM_GI_GI22_Parts, LM_GI_GI22_Playfield)
Dim BL_GI_GI23: BL_GI_GI23=Array(LM_GI_GI23_Parts, LM_GI_GI23_Playfield)
Dim BL_GI_GI24: BL_GI_GI24=Array(LM_GI_GI24_Parts, LM_GI_GI24_Playfield)
Dim BL_L_Light1: BL_L_Light1=Array(LM_L_Light1_Playfield, LM_L_Light1_UnderPF)
Dim BL_L_Light10: BL_L_Light10=Array(LM_L_Light10_Playfield, LM_L_Light10_UnderPF)
Dim BL_L_Light11: BL_L_Light11=Array(LM_L_Light11_Playfield, LM_L_Light11_UnderPF)
Dim BL_L_Light19: BL_L_Light19=Array(LM_L_Light19_Parts, LM_L_Light19_Plaastics, LM_L_Light19_Playfield, LM_L_Light19_sw2, LM_L_Light19_sw5)
Dim BL_L_Light2: BL_L_Light2=Array(LM_L_Light2_Playfield, LM_L_Light2_UnderPF)
Dim BL_L_Light3: BL_L_Light3=Array(LM_L_Light3_Playfield, LM_L_Light3_UnderPF)
Dim BL_L_Light4: BL_L_Light4=Array(LM_L_Light4_Playfield, LM_L_Light4_UnderPF)
Dim BL_L_Light5: BL_L_Light5=Array(LM_L_Light5_Parts, LM_L_Light5_Playfield, LM_L_Light5_UnderPF)
Dim BL_L_Light6: BL_L_Light6=Array(LM_L_Light6_Parts, LM_L_Light6_Plaastics, LM_L_Light6_Playfield, LM_L_Light6_UnderPF)
Dim BL_L_Light7: BL_L_Light7=Array(LM_L_Light7_Parts, LM_L_Light7_Playfield, LM_L_Light7_UnderPF, LM_L_Light7_sw1)
Dim BL_L_Light8: BL_L_Light8=Array(LM_L_Light8_LSling1, LM_L_Light8_LSling2, LM_L_Light8_Parts, LM_L_Light8_Playfield, LM_L_Light8_UnderPF)
Dim BL_L_Light9: BL_L_Light9=Array(LM_L_Light9_Playfield, LM_L_Light9_UnderPF)
Dim BL_L_LightA: BL_L_LightA=Array(LM_L_LightA_Parts, LM_L_LightA_Plaastics, LM_L_LightA_Playfield, LM_L_LightA_UnderPF)
Dim BL_L_LightB: BL_L_LightB=Array(LM_L_LightB_Parts, LM_L_LightB_Playfield, LM_L_LightB_UnderPF)
Dim BL_L_LightC: BL_L_LightC=Array(LM_L_LightC_BumperCap, LM_L_LightC_Parts, LM_L_LightC_Playfield, LM_L_LightC_UnderPF)
Dim BL_L_LightD: BL_L_LightD=Array(LM_L_LightD_Parts, LM_L_LightD_Playfield, LM_L_LightD_UnderPF)
Dim BL_L_LightT1: BL_L_LightT1=Array(LM_L_LightT1_Parts, LM_L_LightT1_Plaastics, LM_L_LightT1_Playfield, LM_L_LightT1_UnderPF, LM_L_LightT1_swT1)
Dim BL_L_LightT2: BL_L_LightT2=Array(LM_L_LightT2_Parts, LM_L_LightT2_Plaastics, LM_L_LightT2_Playfield, LM_L_LightT2_UnderPF, LM_L_LightT2_swT2)
Dim BL_L_bumperlight1: BL_L_bumperlight1=Array(LM_L_bumperlight1_BR1, LM_L_bumperlight1_BS1, LM_L_bumperlight1_BS2, LM_L_bumperlight1_BumperCap, LM_L_bumperlight1_Parts, LM_L_bumperlight1_Plaastics, LM_L_bumperlight1_Playfield)
Dim BL_L_bumperlight2: BL_L_bumperlight2=Array(LM_L_bumperlight2_BR2, LM_L_bumperlight2_BS1, LM_L_bumperlight2_BS2, LM_L_bumperlight2_BS3, LM_L_bumperlight2_BumperCap, LM_L_bumperlight2_Parts, LM_L_bumperlight2_Playfield)
Dim BL_L_bumperlight3: BL_L_bumperlight3=Array(LM_L_bumperlight3_BR3, LM_L_bumperlight3_BS2, LM_L_bumperlight3_BS3, LM_L_bumperlight3_BumperCap, LM_L_bumperlight3_Parts, LM_L_bumperlight3_Playfield)
Dim BL_World: BL_World=Array(BM_BR1, BM_BR2, BM_BR3, BM_BS1, BM_BS2, BM_BS3, BM_BumperCap, BM_Gate, BM_LEMK, BM_LSling1, BM_LSling2, BM_Lflip, BM_Lflip1, BM_Lflip1U, BM_LflipU, BM_Mush1, BM_Mush2, BM_Mush3, BM_Mush4, BM_Parts, BM_Plaastics, BM_Playfield, BM_REMK, BM_RSling1, BM_RSling2, BM_Rflip, BM_Rflip1, BM_Rflip1U, BM_RflipU, BM_UnderPF, BM_sw1, BM_sw2, BM_sw3, BM_sw4, BM_sw5, BM_sw6, BM_sw7, BM_swT1, BM_swT2)
' Global arrays
Dim BG_Bakemap: BG_Bakemap=Array(BM_BR1, BM_BR2, BM_BR3, BM_BS1, BM_BS2, BM_BS3, BM_BumperCap, BM_Gate, BM_LEMK, BM_LSling1, BM_LSling2, BM_Lflip, BM_Lflip1, BM_Lflip1U, BM_LflipU, BM_Mush1, BM_Mush2, BM_Mush3, BM_Mush4, BM_Parts, BM_Plaastics, BM_Playfield, BM_REMK, BM_RSling1, BM_RSling2, BM_Rflip, BM_Rflip1, BM_Rflip1U, BM_RflipU, BM_UnderPF, BM_sw1, BM_sw2, BM_sw3, BM_sw4, BM_sw5, BM_sw6, BM_sw7, BM_swT1, BM_swT2)
Dim BG_Lightmap: BG_Lightmap=Array(LM_GI_GI01_LSling1, LM_GI_GI01_LSling2, LM_GI_GI01_Lflip, LM_GI_GI01_Lflip1, LM_GI_GI01_Lflip1U, LM_GI_GI01_LflipU, LM_GI_GI01_Parts, LM_GI_GI01_Plaastics, LM_GI_GI01_Playfield, LM_GI_GI01_REMK, LM_GI_GI01_RSling1, LM_GI_GI01_RSling2, LM_GI_GI01_Rflip, LM_GI_GI01_Rflip1, LM_GI_GI01_Rflip1U, LM_GI_GI01_RflipU, LM_GI_GI02_Lflip, LM_GI_GI02_Lflip1, LM_GI_GI02_Lflip1U, LM_GI_GI02_LflipU, LM_GI_GI02_Parts, LM_GI_GI02_Plaastics, LM_GI_GI02_Playfield, LM_GI_GI02_REMK, LM_GI_GI02_RSling1, LM_GI_GI02_RSling2, LM_GI_GI02_Rflip, LM_GI_GI02_Rflip1, LM_GI_GI02_Rflip1U, LM_GI_GI02_RflipU, LM_GI_GI02_sw7, LM_GI_GI03_LEMK, LM_GI_GI03_LSling1, LM_GI_GI03_LSling2, LM_GI_GI03_Lflip, LM_GI_GI03_Lflip1, LM_GI_GI03_Lflip1U, LM_GI_GI03_LflipU, LM_GI_GI03_Parts, LM_GI_GI03_Plaastics, LM_GI_GI03_Playfield, LM_GI_GI03_RSling1, LM_GI_GI03_RSling2, LM_GI_GI03_Rflip, LM_GI_GI03_Rflip1, LM_GI_GI03_Rflip1U, LM_GI_GI03_RflipU, LM_GI_GI04_LEMK, LM_GI_GI04_LSling1, LM_GI_GI04_LSling2, LM_GI_GI04_Lflip, _
  LM_GI_GI04_Lflip1, LM_GI_GI04_Lflip1U, LM_GI_GI04_LflipU, LM_GI_GI04_Parts, LM_GI_GI04_Plaastics, LM_GI_GI04_Playfield, LM_GI_GI04_Rflip, LM_GI_GI04_Rflip1, LM_GI_GI04_Rflip1U, LM_GI_GI04_RflipU, LM_GI_GI04_sw6, LM_GI_GI05_Mush3, LM_GI_GI05_Parts, LM_GI_GI05_Plaastics, LM_GI_GI05_Playfield, LM_GI_GI05_UnderPF, LM_GI_GI05_swT1, LM_GI_GI06_BS1, LM_GI_GI06_BS3, LM_GI_GI06_BumperCap, LM_GI_GI06_Mush3, LM_GI_GI06_Parts, LM_GI_GI06_Plaastics, LM_GI_GI06_Playfield, LM_GI_GI06_swT1, LM_GI_GI07_BR1, LM_GI_GI07_BS1, LM_GI_GI07_BS2, LM_GI_GI07_BS3, LM_GI_GI07_BumperCap, LM_GI_GI07_Mush3, LM_GI_GI07_Parts, LM_GI_GI07_Plaastics, LM_GI_GI07_Playfield, LM_GI_GI08_BS3, LM_GI_GI08_Mush4, LM_GI_GI08_Parts, LM_GI_GI08_Plaastics, LM_GI_GI08_Playfield, LM_GI_GI08_sw7, LM_GI_GI08_swT2, LM_GI_GI09_BS2, LM_GI_GI09_Mush4, LM_GI_GI09_Parts, LM_GI_GI09_Plaastics, LM_GI_GI09_Playfield, LM_GI_GI09_UnderPF, LM_GI_GI09_swT2, LM_GI_GI10_BS3, LM_GI_GI10_Mush4, LM_GI_GI10_Parts, LM_GI_GI10_Plaastics, LM_GI_GI10_Playfield, LM_GI_GI10_swT2, _
  LM_GI_GI11_BR3, LM_GI_GI11_BS1, LM_GI_GI11_BS3, LM_GI_GI11_BumperCap, LM_GI_GI11_Mush4, LM_GI_GI11_Parts, LM_GI_GI11_Plaastics, LM_GI_GI11_Playfield, LM_GI_GI11_swT2, LM_GI_GI12_BS1, LM_GI_GI12_BS2, LM_GI_GI12_BumperCap, LM_GI_GI12_Mush1, LM_GI_GI12_Parts, LM_GI_GI12_Plaastics, LM_GI_GI12_Playfield, LM_GI_GI12_UnderPF, LM_GI_GI13_BS1, LM_GI_GI13_Mush1, LM_GI_GI13_Parts, LM_GI_GI13_Plaastics, LM_GI_GI13_Playfield, LM_GI_GI13_sw2, LM_GI_GI13_sw3, LM_GI_GI14_BR3, LM_GI_GI14_BS2, LM_GI_GI14_BS3, LM_GI_GI14_BumperCap, LM_GI_GI14_Mush2, LM_GI_GI14_Parts, LM_GI_GI14_Plaastics, LM_GI_GI14_Playfield, LM_GI_GI15_BS1, LM_GI_GI15_BS2, LM_GI_GI15_BS3, LM_GI_GI15_BumperCap, LM_GI_GI15_Mush2, LM_GI_GI15_Parts, LM_GI_GI15_Plaastics, LM_GI_GI15_Playfield, LM_GI_GI15_sw2, LM_GI_GI16_BS2, LM_GI_GI16_BS3, LM_GI_GI16_BumperCap, LM_GI_GI16_Mush2, LM_GI_GI16_Parts, LM_GI_GI16_Plaastics, LM_GI_GI16_Playfield, LM_GI_GI16_UnderPF, LM_GI_GI16_sw2, LM_GI_GI17_Parts, LM_GI_GI17_Plaastics, LM_GI_GI17_Playfield, LM_GI_GI17_sw1, _
  LM_GI_GI17_sw3, LM_GI_GI18_BS2, LM_GI_GI18_BS3, LM_GI_GI18_Parts, LM_GI_GI18_Playfield, LM_GI_GI18_UnderPF, LM_GI_GI18_sw3, LM_GI_GI18_sw4, LM_GI_GI19_BS1, LM_GI_GI19_BS3, LM_GI_GI19_Parts, LM_GI_GI19_Playfield, LM_GI_GI19_UnderPF, LM_GI_GI19_sw1, LM_GI_GI19_sw2, LM_GI_GI19_sw4, LM_GI_GI20_BS2, LM_GI_GI20_Parts, LM_GI_GI20_Plaastics, LM_GI_GI20_Playfield, LM_GI_GI20_UnderPF, LM_GI_GI20_sw5, LM_GI_GI21_Parts, LM_GI_GI21_Playfield, LM_GI_GI22_Parts, LM_GI_GI22_Playfield, LM_GI_GI23_Parts, LM_GI_GI23_Playfield, LM_GI_GI24_Parts, LM_GI_GI24_Playfield, LM_L_Light1_Playfield, LM_L_Light1_UnderPF, LM_L_Light10_Playfield, LM_L_Light10_UnderPF, LM_L_Light11_Playfield, LM_L_Light11_UnderPF, LM_L_Light19_Parts, LM_L_Light19_Plaastics, LM_L_Light19_Playfield, LM_L_Light19_sw2, LM_L_Light19_sw5, LM_L_Light2_Playfield, LM_L_Light2_UnderPF, LM_L_Light3_Playfield, LM_L_Light3_UnderPF, LM_L_Light4_Playfield, LM_L_Light4_UnderPF, LM_L_Light5_Parts, LM_L_Light5_Playfield, LM_L_Light5_UnderPF, LM_L_Light6_Parts, _
  LM_L_Light6_Plaastics, LM_L_Light6_Playfield, LM_L_Light6_UnderPF, LM_L_Light7_Parts, LM_L_Light7_Playfield, LM_L_Light7_UnderPF, LM_L_Light7_sw1, LM_L_Light8_LSling1, LM_L_Light8_LSling2, LM_L_Light8_Parts, LM_L_Light8_Playfield, LM_L_Light8_UnderPF, LM_L_Light9_Playfield, LM_L_Light9_UnderPF, LM_L_LightA_Parts, LM_L_LightA_Plaastics, LM_L_LightA_Playfield, LM_L_LightA_UnderPF, LM_L_LightB_Parts, LM_L_LightB_Playfield, LM_L_LightB_UnderPF, LM_L_LightC_BumperCap, LM_L_LightC_Parts, LM_L_LightC_Playfield, LM_L_LightC_UnderPF, LM_L_LightD_Parts, LM_L_LightD_Playfield, LM_L_LightD_UnderPF, LM_L_LightT1_Parts, LM_L_LightT1_Plaastics, LM_L_LightT1_Playfield, LM_L_LightT1_UnderPF, LM_L_LightT1_swT1, LM_L_LightT2_Parts, LM_L_LightT2_Plaastics, LM_L_LightT2_Playfield, LM_L_LightT2_UnderPF, LM_L_LightT2_swT2, LM_L_bumperlight1_BR1, LM_L_bumperlight1_BS1, LM_L_bumperlight1_BS2, LM_L_bumperlight1_BumperCap, LM_L_bumperlight1_Parts, LM_L_bumperlight1_Plaastics, LM_L_bumperlight1_Playfield, LM_L_bumperlight2_BR2, _
  LM_L_bumperlight2_BS1, LM_L_bumperlight2_BS2, LM_L_bumperlight2_BS3, LM_L_bumperlight2_BumperCap, LM_L_bumperlight2_Parts, LM_L_bumperlight2_Playfield, LM_L_bumperlight3_BR3, LM_L_bumperlight3_BS2, LM_L_bumperlight3_BS3, LM_L_bumperlight3_BumperCap, LM_L_bumperlight3_Parts, LM_L_bumperlight3_Playfield)
Dim BG_All: BG_All=Array(BM_BR1, BM_BR2, BM_BR3, BM_BS1, BM_BS2, BM_BS3, BM_BumperCap, BM_Gate, BM_LEMK, BM_LSling1, BM_LSling2, BM_Lflip, BM_Lflip1, BM_Lflip1U, BM_LflipU, BM_Mush1, BM_Mush2, BM_Mush3, BM_Mush4, BM_Parts, BM_Plaastics, BM_Playfield, BM_REMK, BM_RSling1, BM_RSling2, BM_Rflip, BM_Rflip1, BM_Rflip1U, BM_RflipU, BM_UnderPF, BM_sw1, BM_sw2, BM_sw3, BM_sw4, BM_sw5, BM_sw6, BM_sw7, BM_swT1, BM_swT2, LM_GI_GI01_LSling1, LM_GI_GI01_LSling2, LM_GI_GI01_Lflip, LM_GI_GI01_Lflip1, LM_GI_GI01_Lflip1U, LM_GI_GI01_LflipU, LM_GI_GI01_Parts, LM_GI_GI01_Plaastics, LM_GI_GI01_Playfield, LM_GI_GI01_REMK, LM_GI_GI01_RSling1, LM_GI_GI01_RSling2, LM_GI_GI01_Rflip, LM_GI_GI01_Rflip1, LM_GI_GI01_Rflip1U, LM_GI_GI01_RflipU, LM_GI_GI02_Lflip, LM_GI_GI02_Lflip1, LM_GI_GI02_Lflip1U, LM_GI_GI02_LflipU, LM_GI_GI02_Parts, LM_GI_GI02_Plaastics, LM_GI_GI02_Playfield, LM_GI_GI02_REMK, LM_GI_GI02_RSling1, LM_GI_GI02_RSling2, LM_GI_GI02_Rflip, LM_GI_GI02_Rflip1, LM_GI_GI02_Rflip1U, LM_GI_GI02_RflipU, LM_GI_GI02_sw7, _
  LM_GI_GI03_LEMK, LM_GI_GI03_LSling1, LM_GI_GI03_LSling2, LM_GI_GI03_Lflip, LM_GI_GI03_Lflip1, LM_GI_GI03_Lflip1U, LM_GI_GI03_LflipU, LM_GI_GI03_Parts, LM_GI_GI03_Plaastics, LM_GI_GI03_Playfield, LM_GI_GI03_RSling1, LM_GI_GI03_RSling2, LM_GI_GI03_Rflip, LM_GI_GI03_Rflip1, LM_GI_GI03_Rflip1U, LM_GI_GI03_RflipU, LM_GI_GI04_LEMK, LM_GI_GI04_LSling1, LM_GI_GI04_LSling2, LM_GI_GI04_Lflip, LM_GI_GI04_Lflip1, LM_GI_GI04_Lflip1U, LM_GI_GI04_LflipU, LM_GI_GI04_Parts, LM_GI_GI04_Plaastics, LM_GI_GI04_Playfield, LM_GI_GI04_Rflip, LM_GI_GI04_Rflip1, LM_GI_GI04_Rflip1U, LM_GI_GI04_RflipU, LM_GI_GI04_sw6, LM_GI_GI05_Mush3, LM_GI_GI05_Parts, LM_GI_GI05_Plaastics, LM_GI_GI05_Playfield, LM_GI_GI05_UnderPF, LM_GI_GI05_swT1, LM_GI_GI06_BS1, LM_GI_GI06_BS3, LM_GI_GI06_BumperCap, LM_GI_GI06_Mush3, LM_GI_GI06_Parts, LM_GI_GI06_Plaastics, LM_GI_GI06_Playfield, LM_GI_GI06_swT1, LM_GI_GI07_BR1, LM_GI_GI07_BS1, LM_GI_GI07_BS2, LM_GI_GI07_BS3, LM_GI_GI07_BumperCap, LM_GI_GI07_Mush3, LM_GI_GI07_Parts, LM_GI_GI07_Plaastics, _
  LM_GI_GI07_Playfield, LM_GI_GI08_BS3, LM_GI_GI08_Mush4, LM_GI_GI08_Parts, LM_GI_GI08_Plaastics, LM_GI_GI08_Playfield, LM_GI_GI08_sw7, LM_GI_GI08_swT2, LM_GI_GI09_BS2, LM_GI_GI09_Mush4, LM_GI_GI09_Parts, LM_GI_GI09_Plaastics, LM_GI_GI09_Playfield, LM_GI_GI09_UnderPF, LM_GI_GI09_swT2, LM_GI_GI10_BS3, LM_GI_GI10_Mush4, LM_GI_GI10_Parts, LM_GI_GI10_Plaastics, LM_GI_GI10_Playfield, LM_GI_GI10_swT2, LM_GI_GI11_BR3, LM_GI_GI11_BS1, LM_GI_GI11_BS3, LM_GI_GI11_BumperCap, LM_GI_GI11_Mush4, LM_GI_GI11_Parts, LM_GI_GI11_Plaastics, LM_GI_GI11_Playfield, LM_GI_GI11_swT2, LM_GI_GI12_BS1, LM_GI_GI12_BS2, LM_GI_GI12_BumperCap, LM_GI_GI12_Mush1, LM_GI_GI12_Parts, LM_GI_GI12_Plaastics, LM_GI_GI12_Playfield, LM_GI_GI12_UnderPF, LM_GI_GI13_BS1, LM_GI_GI13_Mush1, LM_GI_GI13_Parts, LM_GI_GI13_Plaastics, LM_GI_GI13_Playfield, LM_GI_GI13_sw2, LM_GI_GI13_sw3, LM_GI_GI14_BR3, LM_GI_GI14_BS2, LM_GI_GI14_BS3, LM_GI_GI14_BumperCap, LM_GI_GI14_Mush2, LM_GI_GI14_Parts, LM_GI_GI14_Plaastics, LM_GI_GI14_Playfield, LM_GI_GI15_BS1, _
  LM_GI_GI15_BS2, LM_GI_GI15_BS3, LM_GI_GI15_BumperCap, LM_GI_GI15_Mush2, LM_GI_GI15_Parts, LM_GI_GI15_Plaastics, LM_GI_GI15_Playfield, LM_GI_GI15_sw2, LM_GI_GI16_BS2, LM_GI_GI16_BS3, LM_GI_GI16_BumperCap, LM_GI_GI16_Mush2, LM_GI_GI16_Parts, LM_GI_GI16_Plaastics, LM_GI_GI16_Playfield, LM_GI_GI16_UnderPF, LM_GI_GI16_sw2, LM_GI_GI17_Parts, LM_GI_GI17_Plaastics, LM_GI_GI17_Playfield, LM_GI_GI17_sw1, LM_GI_GI17_sw3, LM_GI_GI18_BS2, LM_GI_GI18_BS3, LM_GI_GI18_Parts, LM_GI_GI18_Playfield, LM_GI_GI18_UnderPF, LM_GI_GI18_sw3, LM_GI_GI18_sw4, LM_GI_GI19_BS1, LM_GI_GI19_BS3, LM_GI_GI19_Parts, LM_GI_GI19_Playfield, LM_GI_GI19_UnderPF, LM_GI_GI19_sw1, LM_GI_GI19_sw2, LM_GI_GI19_sw4, LM_GI_GI20_BS2, LM_GI_GI20_Parts, LM_GI_GI20_Plaastics, LM_GI_GI20_Playfield, LM_GI_GI20_UnderPF, LM_GI_GI20_sw5, LM_GI_GI21_Parts, LM_GI_GI21_Playfield, LM_GI_GI22_Parts, LM_GI_GI22_Playfield, LM_GI_GI23_Parts, LM_GI_GI23_Playfield, LM_GI_GI24_Parts, LM_GI_GI24_Playfield, LM_L_Light1_Playfield, LM_L_Light1_UnderPF, LM_L_Light10_Playfield, _
  LM_L_Light10_UnderPF, LM_L_Light11_Playfield, LM_L_Light11_UnderPF, LM_L_Light19_Parts, LM_L_Light19_Plaastics, LM_L_Light19_Playfield, LM_L_Light19_sw2, LM_L_Light19_sw5, LM_L_Light2_Playfield, LM_L_Light2_UnderPF, LM_L_Light3_Playfield, LM_L_Light3_UnderPF, LM_L_Light4_Playfield, LM_L_Light4_UnderPF, LM_L_Light5_Parts, LM_L_Light5_Playfield, LM_L_Light5_UnderPF, LM_L_Light6_Parts, LM_L_Light6_Plaastics, LM_L_Light6_Playfield, LM_L_Light6_UnderPF, LM_L_Light7_Parts, LM_L_Light7_Playfield, LM_L_Light7_UnderPF, LM_L_Light7_sw1, LM_L_Light8_LSling1, LM_L_Light8_LSling2, LM_L_Light8_Parts, LM_L_Light8_Playfield, LM_L_Light8_UnderPF, LM_L_Light9_Playfield, LM_L_Light9_UnderPF, LM_L_LightA_Parts, LM_L_LightA_Plaastics, LM_L_LightA_Playfield, LM_L_LightA_UnderPF, LM_L_LightB_Parts, LM_L_LightB_Playfield, LM_L_LightB_UnderPF, LM_L_LightC_BumperCap, LM_L_LightC_Parts, LM_L_LightC_Playfield, LM_L_LightC_UnderPF, LM_L_LightD_Parts, LM_L_LightD_Playfield, LM_L_LightD_UnderPF, LM_L_LightT1_Parts, LM_L_LightT1_Plaastics, _
  LM_L_LightT1_Playfield, LM_L_LightT1_UnderPF, LM_L_LightT1_swT1, LM_L_LightT2_Parts, LM_L_LightT2_Plaastics, LM_L_LightT2_Playfield, LM_L_LightT2_UnderPF, LM_L_LightT2_swT2, LM_L_bumperlight1_BR1, LM_L_bumperlight1_BS1, LM_L_bumperlight1_BS2, LM_L_bumperlight1_BumperCap, LM_L_bumperlight1_Parts, LM_L_bumperlight1_Plaastics, LM_L_bumperlight1_Playfield, LM_L_bumperlight2_BR2, LM_L_bumperlight2_BS1, LM_L_bumperlight2_BS2, LM_L_bumperlight2_BS3, LM_L_bumperlight2_BumperCap, LM_L_bumperlight2_Parts, LM_L_bumperlight2_Playfield, LM_L_bumperlight3_BR3, LM_L_bumperlight3_BS2, LM_L_bumperlight3_BS3, LM_L_bumperlight3_BumperCap, LM_L_bumperlight3_Parts, LM_L_bumperlight3_Playfield)
' VLM  Arrays - End

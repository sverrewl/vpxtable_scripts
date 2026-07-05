Option Explicit
Dim FirstUniqueEquipped : FirstUniqueEquipped = False

'First, try to load the Controller.vbs (DOF), which helps controlling additional hardware like lights, gears, knockers, bells and chimes (to increase realism)
'First, try to load the Controller.vbs (DOF), which helps controlling additional hardware like lights, gears, knockers, bells and chimes (to increase realism)
'This table uses DOF via the 'SoundFX' calls that are inserted in some of the PlaySound commands, which will then fire an additional event, instead of just playing a sample/sound effect
On Error Resume Next
ExecuteGlobal GetTextFile("controller.vbs")
'If Err Then ShowMessage "You need the Controller.vbs file in order to run this table (installed with the VPX package in the scripts folder)"
On Error Goto 0

'If using Visual PinMAME (VPM), place the ROM/game name in the constant below,
'both for VPM, and DOF to load the right DOF config from the Configtool, whether it's a VPM or an Original table

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
     Dim BOT
     BOT = GetBalls

  ' stop the sound of deleted balls
  ' stop the sound of deleted balls
  For b = UBound(BOT) + 1 To tnob - 1
    rolling(b) = False
    StopSound("BallRoll_" & b)
  Next

  ' exit the sub if no balls on the table
  If UBound(BOT) =  - 1 Then Exit Sub

  ' play the rolling sound for each ball
  For b = 0 To UBound(BOT)
    If BallVel(BOT(b)) > 1 And BOT(b).z < 30 Then
      rolling(b) = True
      PlaySound ("BallRoll_" & b), - 1, VolPlayfieldRoll(BOT(b)) * BallRollVolume * VolumeDial, AudioPan(BOT(b)), 0, PitchPlayfieldRoll(BOT(b)), 1, 0, AudioFade(BOT(b))
    Else
      If rolling(b) = True Then
        StopSound("BallRoll_" & b)
        rolling(b) = False
      End If
    End If

    ' Ball Drop Sounds
    If BOT(b).VelZ <  - 1 And BOT(b).z < 55 And BOT(b).z > 27 Then 'height adjust for ball drop sounds
      If DropCount(b) >= 5 Then
        DropCount(b) = 0
        If BOT(b).velz >  - 7 Then
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

'*******************************************
'  ZOPT: User Options
'*******************************************

Const FreePLay = False   ' Set to False for coin-op mode

'--- VR ROOM: Toggle collection visibility ---
'    1=MixedReality, 2=VRMin, 3=VRMega
Sub SetVRRoom(roomChoice)
    Dim p
    For Each p In MixedReality : p.Visible = False : Next
    For Each p In VRMin        : p.Visible = False : Next
    For Each p In VRMega       : p.Visible = False : Next

    ' Only show a VR room when actually running in VR (RenderingMode = 2)
    If RenderingMode <> 2 Then Exit Sub

    Select Case roomChoice
        Case 1 : For Each p In MixedReality : p.Visible = True : Next
        Case 2 : For Each p In VRMin        : p.Visible = True : Next
        Case 3 : For Each p In VRMega       : p.Visible = True : Next
    End Select
End Sub

'----- DMD Options -----
Const UseFlexDMD = 1        '0 = no FlexDMD, 1 = enable FlexDMD
Const FlexONPlayfield = False    'False = off, True=DMD on playfield ( vrroom overrides this )

'----- VR Room -----
Dim VRRoomChoice : VRRoomChoice = 2   ' 1=MixedReality, 2=VRMin, 3=VRMega

Dim LightLevel : LightLevel = 0.25        ' Level of room lighting (0 to 1), where 0 is dark and 100 is brightest
Dim ColorLUT : ColorLUT = 1           ' Color desaturation LUTs: 1 to 11, where 1 is normal and 11 is black'n'white
Dim VolumeDial : VolumeDial = 1             ' Overall Mechanical sound effect volume. Recommended values should be no greater than 1.
Dim BallRollVolume : BallRollVolume = 1     ' Level of ball rolling volume. Value between 0 and 1
Dim RampRollVolume : RampRollVolume = 1     ' Level of ramp rolling volume. Value between 0 and 1
'Dim StagedFlippers : StagedFlippers = 0         ' Staged Flippers. 0 = Disabled, 1 = Enabled

Const VRPlungerBaseY = 2192
Const VRPlungerBladeBaseY = 2211.655
Const VRPlungerScale = 5

Const VRPlungerBtnMax = 31        ' StrokeLength 155 / VRPlungerScale 5
Const VRPlungerBtnPullRate = 0.485  ' slow retract — matches Pull Speed 0.25
Const VRPlungerBtnFireRate = 4    ' fast snap — matches Release Speed 100
Dim VRPlungerBtnPos : VRPlungerBtnPos = 0
Dim VRPlungerBtnHeld : VRPlungerBtnHeld = False

Sub VRPlungerTimer_Timer()
    If VRPlungerBtnHeld Then
        VRPlungerBtnPos = VRPlungerBtnPos + VRPlungerBtnPullRate
        If VRPlungerBtnPos > VRPlungerBtnMax Then VRPlungerBtnPos = VRPlungerBtnMax
    Else
        VRPlungerBtnPos = VRPlungerBtnPos - VRPlungerBtnFireRate
        If VRPlungerBtnPos < 0 Then VRPlungerBtnPos = 0
    End If

    Dim p : p = Plunger.Position
    If VRPlungerBtnPos > p Then p = VRPlungerBtnPos

    PinCab_ShooterHilt.Y = VRPlungerBaseY + (VRPlungerScale * p)
    PinCab_ShooterBlade.Y = VRPlungerBladeBaseY + (VRPlungerScale * p)
End Sub

' Called when options are tweaked by the player.
' - 0: game has started, good time to load options and adjust accordingly
' - 1: an option has changed
' - 2: options have been reseted
' - 3: player closed the tweak UI, good time to update staticly prerendered parts
' Table1.Option arguments are:
' - option name, minimum value, maximum value, step between valid values, default value, unit (0=None, 1=Percent), an optional arry of literal strings
Dim dspTriggered : dspTriggered = False

Sub Table1_OptionEvent(ByVal eventId)
'********
' DisableStaticPreRendering = True    ' re-assert on every event including close
'
'    If eventId = 3 Then Exit Sub
' Don't use ^^^
'********
    ' Toggle static prerendering exactly once per tweak-menu session (10.8.1 counter-safe)
    If eventId = 3 Then
        If dspTriggered Then
            DisableStaticPreRendering = False
            dspTriggered = False
        End If
        Exit Sub
    ElseIf eventId = 1 Or eventId = 2 Then
        If Not dspTriggered Then
            DisableStaticPreRendering = True
            dspTriggered = True
        End If
    End If

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

' VR Room Selection (replaces Sphere option)
    VRRoomChoice = Table1.Option("VR Room", 1, 3, 1, 2, 0, Array("Mixed Reality", "Minimal", "Mega Dungeon"))
    SetVRRoom VRRoomChoice

Dim SkullsOpt : SkullsOpt = 1
    SkullsOpt = Table1.Option("Skulls", 1, 2, 1, 1, 0, Array("Enabled", "Disabled"))
    skulls.Visible = (SkullsOpt = 1)

    ScytheVisibleOpt = Table1.Option("Scythes", 1, 2, 1, 1, 0, Array("Enabled", "Disabled"))
    Scythe1.Visible = (ScytheVisibleOpt = 1)
    Scythe2.Visible = (ScytheVisibleOpt = 1)

' Diablo Statue (VR)
    Dim DiabloStatueOpt
    DiabloStatueOpt = Table1.Option("Diablo Statue", 1, 2, 1, 1, 0, Array("Enabled", "Disabled"))
    VR_d_Diablo.Visible = (DiabloStatueOpt = 1)

    ' Ramp Swords
    Dim RampSwordsOpt
    RampSwordsOpt = Table1.Option("Ramp Swords", 1, 2, 1, 1, 0, Array("Enabled", "Disabled"))
    grandpa.Visible = (RampSwordsOpt = 1)
    grandpa001.Visible = (RampSwordsOpt = 1)

    ' Headstones
    Dim HeadstonesOpt, hsv
    HeadstonesOpt = Table1.Option("Headstones", 1, 2, 1, 1, 0, Array("Enabled", "Disabled"))
    For Each hsv In Headstones
        hsv.Visible = (HeadstonesOpt = 1)
    Next

Dim BackglassOpt : BackglassOpt = 1
    BackglassOpt = Table1.Option("Backglass Art", 1, 4, 1, 1, 0, Array("Classic", "D2R Style 1", "D2R Style 2", "Lord of Destruction"))
    Select Case BackglassOpt
        Case 1 : PinCab_Backglass.Image = "BackglassImage"
        Case 2 : PinCab_Backglass.Image = "BackglassImageD2R1"
        Case 3 : PinCab_Backglass.Image = "BackglassImageD2R2"
        Case 4 : PinCab_Backglass.Image = "BackglassImageLOD"
    End Select
    ' Sound volumes
    VolumeDial = Table1.Option("Mech Volume", 0, 1, 0.01, 0.8, 1)
    BallRollVolume = Table1.Option("Ball Roll Volume", 0, 1, 0.01, 0.5, 1)
  RampRollVolume = Table1.Option("Ramp Roll Volume", 0, 1, 0.01, 0.5, 1)

  ' Loot glow — flat halo brightness for all rarities (0 = off, ~1 = clean color, up to 5)
  LootGlowLevel = Table1.Option("Loot Glow", 0, 30, 0.5, 10, 0)
  ' re-apply to any loot on the field so the slider previews live
  Dim lgi
  For lgi = 0 To 3
    If LootActive(lgi) Then SetLootPrimAppearance lgi, DtLootTier(lgi)
  Next

  ' Room brightness
' LightLevel = Table1.Option("Table Brightness (Ambient Light Level)", 0, 1, 0.01, .5, 1)
  LightLevel = NightDay/100
' SetRoomBrightness LightLevel   'Uncomment this line for lightmapped tables.

'    '    ' Staged Flippers
'    StagedFlippers = Table1.Option("Staged Flippers", 0, 1, 1, 0, 0, Array("Disabled", "Enabled"))

End Sub


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
  DoDtLootAnim
  FrameTime = GameTime - InitFrameTime
  InitFrameTime = GameTime  'Count frametime
  'Add animation stuff here
  RollingUpdate       'update rolling sounds
' DoDTAnim        'handle drop target animations
' DoSTAnim        'handle stand up target animations
' queue.Tick            'handle the queue system
' BSUpdate
End Sub

'The CorTimer interval should be 10. It's sole purpose is to update the Cor calculations
CorTimer.Interval = 10
Sub CorTimer_Timer(): Cor.Update: End Sub

Sub DoDtLootAnim()
    Dim i
    For i = 0 To 3
        ' --- Bend handling (only on drop) ---
        If DtLootBendPhase(i) = 1 Then
            DtLootBend(i) = DtLootBend(i) + DtLootBendSpeed
            If DtLootBend(i) >= DtLootMaxBend Then
                DtLootBend(i) = DtLootMaxBend
                DtLootBendPhase(i) = 2                  ' reached max, now return as it drops
            End If
            PrimDtLootArr(i).RotX = DtLootRestRotX(i) - DtLootBend(i)
        ElseIf DtLootBendPhase(i) = 2 Then
            DtLootBend(i) = DtLootBend(i) - DtLootBendSpeed
            If DtLootBend(i) <= 0 Then
                DtLootBend(i) = 0
                DtLootBendPhase(i) = 0
            End If
            PrimDtLootArr(i).RotX = DtLootRestRotX(i) - DtLootBend(i)
        End If
        ' --- Vertical drop/raise ---
        If DtLootCurZ(i) <> DtLootTargetZ(i) Then
            If DtLootCurZ(i) > DtLootTargetZ(i) Then
                DtLootCurZ(i) = DtLootCurZ(i) - DtLootDropSpeed
                If DtLootCurZ(i) < DtLootTargetZ(i) Then DtLootCurZ(i) = DtLootTargetZ(i)
            Else
                DtLootCurZ(i) = DtLootCurZ(i) + DtLootRaiseSpeed
                If DtLootCurZ(i) > DtLootTargetZ(i) Then DtLootCurZ(i) = DtLootTargetZ(i)
            End If
            PrimDtLootArr(i).z = DtLootCurZ(i)
        ElseIf LootActive(i) And DtLootTier(i) = GEAR_UNIQUE And UniqueLevPhase(i) > 0 Then
            ' --- UNIQUE: levitate + spin cycle (replaces bob for unique) ---
            Select Case UniqueLevPhase(i)
                Case 1 ' rising
                    UniqueLevZ(i) = UniqueLevZ(i) + UniqueLevRiseSpd
                    If UniqueLevZ(i) >= UniqueLevHeight Then
                        UniqueLevZ(i) = UniqueLevHeight
                        UniqueLevPhase(i) = 2
                        UniqueSpinDeg(i) = 0
                    End If
                Case 2 ' spinning (hover at top, full 360)
                    UniqueSpinDeg(i) = UniqueSpinDeg(i) + UniqueSpinSpd
                    If UniqueSpinDeg(i) >= 360 Then
                        UniqueSpinDeg(i) = 0
                        PrimDtLootArr(i).RotY = DtLootRestRotY(i)
                        UniqueLevPhase(i) = 3
                    Else
                        PrimDtLootArr(i).RotY = DtLootRestRotY(i) + UniqueSpinDeg(i)
                    End If
                Case 3 ' descending
                    UniqueLevZ(i) = UniqueLevZ(i) - UniqueLevRiseSpd
                    If UniqueLevZ(i) <= 0 Then
                        UniqueLevZ(i) = 0
                        UniqueLevPhase(i) = 4
                        UniquePauseCnt(i) = 0
                    End If
                Case 4 ' pause then repeat
                    UniquePauseCnt(i) = UniquePauseCnt(i) + 1
                    If UniquePauseCnt(i) >= UniquePauseFrames Then UniqueLevPhase(i) = 1
            End Select
            PrimDtLootArr(i).z = DtLootUpZ(i) + UniqueLevZ(i)
        ElseIf LootActive(i) And DtLootTargetZ(i) = DtLootUpZ(i) Then
            ' Non-unique: idle bob (dips below top only)
            DtLootBobPhase(i) = DtLootBobPhase(i) + DtLootBobSpeed
            PrimDtLootArr(i).z = DtLootUpZ(i) - DtLootBobAmp * ((1 - Cos(DtLootBobPhase(i))) / 2)
        End If

        Next
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

Const TableName = "Diablo2"
Const tnob = 13

Dim EnableBallControl
EnableBallControl = false

Function GetBIP()
    Dim allB : allB = GetBalls()
    Dim b, count : count = 0
    For Each b In allB
        If b.ID <> CapBallID And b.ID <> CapBall2ID Then count = count + 1
    Next
    GetBIP = count
End Function

Dim Credits : Credits = 0

'Dim PortalTime      : PortalTime      = 0
'Dim PortalSurge     : PortalSurge     = 0
'Dim PortalActive    : PortalActive    = False
'Dim PortalCollapsed : PortalCollapsed = False
Dim CowKingKillCount : CowKingKillCount = 0
Dim WasActSixComplete : WasActSixComplete = False

Dim ElementBallActive : ElementBallActive = False
Const BallSize = 50
Const BallMass = 1
Const CapBallMass = .9
Dim CapBallID : CapBallID = -1
Dim CapBall2ID : CapBall2ID = -1
Const ArenaCenterX = 547
Const ArenaCenterY = 446.3006
Dim BallSaveActive : BallSaveActive = False
Dim BallSaveUsed : BallSaveUsed = False
Dim MercSaveActive : MercSaveActive = False

Dim SafeTravelBank  : SafeTravelBank  = 0
Dim BonusSafeTravel : BonusSafeTravel = 0
Dim BallCritCount  : BallCritCount  = 0
Dim BonusCritStrike : BonusCritStrike = 0
Dim BallWhirlwindScore : BallWhirlwindScore = 0
Dim BallBackstabScore : BallBackstabScore = 0

Dim AttractChaseStep : AttractChaseStep = 0
Dim AttractPhase : AttractPhase = 0
Dim TutHoldPending : TutHoldPending = False
Dim TutorialActive : TutorialActive = False
Dim TutPulseActive : TutPulseActive = False
Dim TutorialStep   : TutorialStep = 0
Dim TutLeftHeld    : TutLeftHeld = False
Dim TutRightHeld   : TutRightHeld = False
Dim TutFlipCount   : TutFlipCount = 0
Dim TutFlipperUp   : TutFlipperUp = False
Dim TutWWStep      : TutWWStep = 0
Dim AttractFlashStep : AttractFlashStep = 0
Dim AttractPulseStep : AttractPulseStep = 0
Dim AttractRuneColorStep : AttractRuneColorStep = 0
Dim AttractCharStep : AttractCharStep = 0
Dim AttractGearStep : AttractGearStep = 0

Dim NightDay : NightDay = Table1.NightDay

Dim GIEventMode : GIEventMode = 0
Dim GIEventStep : GIEventStep = 0
Dim GIEventColorR : GIEventColorR = 255
Dim GIEventColorG : GIEventColorG = 80
Dim GIEventColorB : GIEventColorB = 0
Dim GIEventSweepCount : GIEventSweepCount = 0
Dim GIEventSweepMax : GIEventSweepMax = 1
Dim GIEventSweepDir : GIEventSweepDir = 1

' GI Event Modes
Const GI_MODE_NONE       = 0
Const GI_MODE_SHOUT      = 1
Const GI_MODE_BOSS       = 2
Const GI_MODE_FRENZY     = 3
Const GI_MODE_DRAIN      = 4
Const GI_MODE_RUNEWORD   = 5
Const GI_MODE_MYSTERY    = 6
Const GI_MODE_GAMEOVER   = 7
Const GI_MODE_BOSSWIN    = 8
Const GI_MODE_AMBUSH     = 9
Const GI_MODE_SHIELD     = 10
Const GI_MODE_UNIQUE     = 11
Const GI_MODE_EXTRABALL = 12

Dim GIFlickerBase : GIFlickerBase = 15.0
Dim GIFlickerIntensity(40)
Dim GIFlickerTarget(40)
Dim GIFlickerCount : GIFlickerCount = 0

Dim GameOverScoreTimer_active : GameOverScoreTimer_active = False
Dim MusicVolume : MusicVolume = 0.5

Dim DifficultyLevel : DifficultyLevel = 0
Dim ExtraBallMilestoneIdx : ExtraBallMilestoneIdx = 0
Dim ExtraBallMilestones(3)
ExtraBallMilestones(0) = 100000000
ExtraBallMilestones(1) = 300000000
ExtraBallMilestones(2) = 600000000
ExtraBallMilestones(3) = 1000000000
Dim ExtraBallPending : ExtraBallPending = 0
Dim ExtraBallEoBPending : ExtraBallEoBPending = 0
Dim wasExtraBall : wasExtraBall = False
Dim ExtraBallConsumed : ExtraBallConsumed = False

'Dim FlipperUpAttackMinimumSoundLevel : FlipperUpAttackMinimumSoundLevel = 0.010
'Dim FlipperUpAttackMaximumSoundLevel : FlipperUpAttackMaximumSoundLevel = 0.635
'Dim FlipperUpAttackLeftSoundLevel : FlipperUpAttackLeftSoundLevel = 0
'Dim FlipperUpAttackRightSoundLevel : FlipperUpAttackRightSoundLevel = 0
'Dim FlipperUpSoundLevel : FlipperUpSoundLevel = 1.0
'Dim FlipperDownSoundLevel : FlipperDownSoundLevel = 0.45
'Dim FlipperLeftHitParm : FlipperLeftHitParm = 1.0
'Dim FlipperRightHitParm : FlipperRightHitParm = 1.0
'Dim SlingshotSoundLevel : SlingshotSoundLevel = 0.95
'Dim BumperSoundFactor : BumperSoundFactor = 4.25
'Dim KnockerSoundLevel : KnockerSoundLevel = 1.0
'Dim BallWithBallCollisionSoundFactor : BallWithBallCollisionSoundFactor = 3.2
'Dim RubberStrongSoundFactor : RubberStrongSoundFactor = 0.055 / 5
'Dim RubberWeakSoundFactor : RubberWeakSoundFactor = 0.075 / 5
'Dim RubberFlipperSoundFactor : RubberFlipperSoundFactor = 0.075 / 5
'Dim BallBouncePlayfieldSoftFactor : BallBouncePlayfieldSoftFactor = 0.025
'Dim BallBouncePlayfieldHardFactor : BallBouncePlayfieldHardFactor = 0.025
'Dim DelayedBallDropOnPlayfieldSoundLevel : DelayedBallDropOnPlayfieldSoundLevel = 0.8
'Dim WallImpactSoundFactor : WallImpactSoundFactor = 0.075
'Dim MetalImpactSoundFactor : MetalImpactSoundFactor = 0.075 / 3
'Dim GateSoundLevel : GateSoundLevel = 0.5 / 5
'Dim TargetSoundFactor : TargetSoundFactor = 0.0025 * 10
'Dim SpinnerSoundLevel : SpinnerSoundLevel = 0.5
'Dim RolloverSoundLevel : RolloverSoundLevel = 0.25
'Dim DrainSoundLevel : DrainSoundLevel = 0.8
'Dim BallReleaseSoundLevel : BallReleaseSoundLevel = 1.0
'Dim BottomArchBallGuideSoundFactor : BottomArchBallGuideSoundFactor = 0.2
'Dim FlipperBallGuideSoundFactor : FlipperBallGuideSoundFactor = 0.015
'Dim ArchSoundFactor : ArchSoundFactor = 0.025 / 5
'Dim SaucerLockSoundLevel : SaucerLockSoundLevel = 0.8
'Dim SaucerKickSoundLevel : SaucerKickSoundLevel = 0.8
'Dim DTSoundLevel : DTSoundLevel = 0.25
Dim tablewidth : tablewidth = Table1.Width
Dim tableheight : tableheight = Table1.Height

Sub TutorialHoldTimer_Timer()
    TutorialHoldTimer.Enabled = False
    TutHoldPending = False
    If TutLeftHeld And TutRightHeld And Not TutorialActive _
       And Not GameActive And Not hsbModeActive And Not hsbKillsModeActive Then
        StartTutorial
    End If
End Sub

Dim AudioDucked : AudioDucked = False
Dim DuckVolume : DuckVolume = 1.0

Sub DuckAudio()
    If AudioDucked Then Exit Sub
    AudioDucked = True
    DuckVolume = 0.3
    PlaySound CurrentSong, -1, 0.3, 0, 0, 0, 1, 0
End Sub

Sub RestoreAudio()
    If Not AudioDucked Then Exit Sub
    AudioDucked = False
    DuckVolume = 1.0
    PlaySound CurrentSong, -1, 1.0, 0, 0, 0, 1, 0
End Sub

Sub TiltDuckAudio()
    AudioDucked = True
    DuckVolume = 0
    PlaySound CurrentSong, -1, 0, 0, 0, 0, 1, 0
End Sub

Dim CalloutQueue(9)
Dim CalloutQueueDur(9)
Dim CalloutQueueCount : CalloutQueueCount = 0
Dim CalloutPlaying : CalloutPlaying = False
Dim CurrentCalloutSound : CurrentCalloutSound = ""

Sub PlayCallout(snd, duration)
    If CalloutPlaying Then
        If CalloutQueueCount < 10 Then
            CalloutQueue(CalloutQueueCount) = snd
            CalloutQueueDur(CalloutQueueCount) = duration
            CalloutQueueCount = CalloutQueueCount + 1
        End If
        Exit Sub
    End If
    CalloutPlaying = True
    CurrentCalloutSound = snd
    DuckAudio
    PlaySound snd, 0, 3
'    BackglassPulse duration, 255, 20, 0
    DuckAudioTimer.Enabled = False
    DuckAudioTimer.Interval = duration
    DuckAudioTimer.Enabled = True
End Sub

Sub DuckAudioTimer_Timer()
    DuckAudioTimer.Enabled = False
    CurrentCalloutSound = ""
    CalloutPlaying = False
    If CalloutQueueCount > 0 Then
        Dim nextSnd : nextSnd = CalloutQueue(0)
        Dim nextDur : nextDur = CalloutQueueDur(0)
        Dim qi
        For qi = 0 To CalloutQueueCount - 2
            CalloutQueue(qi) = CalloutQueue(qi + 1)
            CalloutQueueDur(qi) = CalloutQueueDur(qi + 1)
        Next
        CalloutQueueCount = CalloutQueueCount - 1
        PlayCallout nextSnd, nextDur
    Else
        RestoreAudio
    End If
End Sub

Dim LF : Set LF = New FlipperPolarity
Dim RF : Set RF = New FlipperPolarity
Dim ULF : Set ULF = New FlipperPolarity

Dim LFPress, RFPress, LFCount, RFCount
Dim LFState, RFState
Dim ULFPress, ULFCount, ULFState, ULFEndAngle
Dim EOST, EOSA, Frampup, FElasticity, FReturn
Dim RFEndAngle, LFEndAngle
Dim FCCDamping : FCCDamping = 0.4

Const FlipperCoilRampupMode = 0
Const EOSTnew = 1.2
Const EOSAnew = 1
Const EOSRampup = 0
Const LiveCatch = 16
Const LiveElasticity = 0.45
Const SOSEM = 0.815
Const EOSReturn = 0.025
Const LiveDistanceMin = 5
Const LiveDistanceMax = 114
Const BaseDampen = 0.55

Dim SOSRampup
SOSRampup = 2.5

LFState = 1
RFState = 1
'
'Sub ResetHighScores()
'    Dim i
'    For i = 1 To 5
'        SaveValue TableName, "HiScore" & i, ""
'        SaveValue TableName, "HiName" & i, ""
'        SaveValue TableName, "HiKills" & i, ""
'        SaveValue TableName, "HiKillsName" & i, ""
'    Next
'    SaveValue TableName, "NightmareChamp",     ""
'    SaveValue TableName, "NightmareChampName", ""
'    SaveValue TableName, "HellChamp",          ""
'    SaveValue TableName, "HellChampName",      ""
'End Sub

Dim Controller
Sub Table1_Init()
InitDtLootAnim
SetVRRoom VRRoomChoice
Scythe1RestRotZ = Scythe1.RotZ
Scythe2RestRotZ = Scythe2.RotZ
'ResetHighScores
'DisableStaticPreRendering = True
Dim bsi
  Dim gii2
For gii2 = 0 To GI.Count - 1
    GIFlickerIntensity(gii2) = GIFlickerBase
    GIFlickerTarget(gii2) = GIFlickerBase
Next
GearBlinkTimer.Interval = 600
GIFlickerCount = GI.Count
  CubeLid.Collidable = True
  CubeKick.Enabled = False
    CubeLidClosedY = CubeLid.TransY
    CubeLidOpenY = CubeLid.TransY + 82
    CubeLidTargetY = CubeLid.TransY
    CubeLidTimer.Interval = CubeLidTickMs
    CubeLidStep = (CubeLidOpenY - CubeLidClosedY) * CubeLidTickMs / CubeLidOpenMs
    LidPerSpin = (CubeLidOpenY - CubeLidClosedY) / CubeSpinTarget
    CubeGate.Collidable = True
  ChestRamp.RotX = ChestRampLand
  ChestRampPhysics.Collidable = True
  ChestRampActive = False
  Shield.Collidable = False
  InitPolarity
  PlaySong "introedit"
  RuneKick.Enabled = True
    BumperRotTimer.Enabled = True
    Dim capBall : Set capBall = CapKicker1.CreateBall
    capBall.Color = RGB(180, 0, 0)
    CapBallID = capBall.ID
    CapKicker1.Kick 180, 1
    CapKicker1.Enabled = False
    Dim capBall2 : Set capBall2 = CapKicker2.CreateBall
    capBall2.Color = RGB(180, 0, 0)
    CapBall2ID = capBall2.ID
    CapKicker2.Kick 180, 1
    CapKicker2.Enabled = False
'    CapKicker3.CreateBall
'    CapKicker3.Kick 180, 1
    AssignBumperStats 0, True
    AssignBumperStats 1, False
    AssignBumperStats 2, False
    AssignBumperStats 3, False
    AssignBumperStats 4, False
  Dim ps : For ps = 0 To 4 : BumperPoisoned(ps) = False : Next
    EnemiesRemaining = 0
    InitParty
    Dim savedCredits : savedCredits = LoadValue("D2", "Credits")
    If savedCredits = "" Then savedCredits = "0"
    Credits = CInt(savedCredits)
    If Credits < 0 Then Credits = 0
    GameActive = False
    BallNumber = 0
    Score = 0
    InitFlexDMD
    LoadHighScores
    DimAllFlashers
    StartAttractMode
  RunePopRestZ(0) = 25    ' RuneR
    RunePopRestZ(1) = 25    ' RuneU
    RunePopRestZ(2) = 25    ' RuneN
    RunePopRestZ(3) = 25   ' RuneE
    RunePopRestZ(4) = 25   ' RuneW
    RunePopRestZ(5) = 25    ' RuneO
    RunePopRestZ(6) = 25   ' RuneR2
    RunePopRestZ(7) = 25   ' RuneD
  LoadB2SController
End Sub

Sub LoadB2SController()
    If RenderingMode = 2 Then Exit Sub   ' VR — no backglass
    On Error Resume Next
    Set Controller = CreateObject("B2S.Server")
    Controller.B2SName = "Diablo_2"
    Controller.Run
    If Err Then
        Err.Clear
        Set Controller = Nothing
    End If
    On Error Goto 0
End Sub

InitPolarity

Sub InitPolarity()
    Dim x, a
    a = Array(LF, RF)
    For Each x In a
        x.AddPt "Ycoef", 0, RightFlipper.Y-65, 1
        x.AddPt "Ycoef", 1, RightFlipper.Y-11, 1
        x.Enabled = True
        x.TimeDelay = 60
        x.DebugOn = False
        x.AddPt "Polarity", 0, 0, 0
        x.AddPt "Polarity", 1, 0.05, -5.5
        x.AddPt "Polarity", 2, 0.16, -5.5
        x.AddPt "Polarity", 3, 0.20, -0.75
        x.AddPt "Polarity", 4, 0.25, -1.25
        x.AddPt "Polarity", 5, 0.3, -1.75
        x.AddPt "Polarity", 6, 0.4, -3.5
        x.AddPt "Polarity", 7, 0.5, -5.25
        x.AddPt "Polarity", 8, 0.7, -4.0
        x.AddPt "Polarity", 9, 0.75, -3.5
        x.AddPt "Polarity", 10, 0.8, -3.0
        x.AddPt "Polarity", 11, 0.85, -2.5
        x.AddPt "Polarity", 12, 0.9, -2.0
        x.AddPt "Polarity", 13, 0.95, -1.5
        x.AddPt "Polarity", 14, 1, -1.0
        x.AddPt "Polarity", 15, 1.05, -0.5
        x.AddPt "Polarity", 16, 1.1, 0
        x.AddPt "Polarity", 17, 1.3, 0
        x.AddPt "Velocity", 0, 0, 0.85
        x.AddPt "Velocity", 1, 0.23, 0.85
        x.AddPt "Velocity", 2, 0.27, 1
        x.AddPt "Velocity", 3, 0.3, 1
        x.AddPt "Velocity", 4, 0.35, 1
        x.AddPt "Velocity", 5, 0.6, 1
        x.AddPt "Velocity", 6, 0.62, 1.0
        x.AddPt "Velocity", 7, 0.702, 0.968
        x.AddPt "Velocity", 8, 0.95, 0.968
        x.AddPt "Velocity", 9, 1.03, 0.945
        x.AddPt "Velocity", 10, 1.5, 0.945
    Next
    LF.SetObjects "LF", LeftFlipper, TriggerLF
    RF.SetObjects "RF", RightFlipper, TriggerRF
    ULF.SetObjects "ULF", UpperLeftFlipper, TriggerUpperLF
    EOST = LeftFlipper.EOSTorque
    EOSA = LeftFlipper.EOSTorqueAngle
    Frampup = LeftFlipper.RampUp
    FElasticity = LeftFlipper.Elasticity
    FReturn = LeftFlipper.Return
    LFEndAngle = LeftFlipper.EndAngle
    RFEndAngle = RightFlipper.EndAngle
  RFEndAngle = RightFlipper.EndAngle
    ULFEndAngle = UpperLeftFlipper.EndAngle
    ULFState = 1
End Sub

Sub StartGame()
    If GameActive Then Exit Sub
    FirstPackDone = False   ' ← first pack of new game always normal enemies
    If Not FreePlay Then
        If Credits < 1 Then
            ShowMysteryDMD "INSERT COIN", "TO PLAY"
            Exit Sub
        End If
        Credits = Credits - 1
        SaveValue "D2", "Credits", Credits
    End If
  Flasher001.Visible = False : Flasher002.Visible = False
    Flasher003.Visible = False : Flasher004.Visible = False
    Flasher005.Visible = False : Flasher006.Visible = False
    Flasher007.Visible = False : Flasher008.Visible = False
  CubeTransmuteActive = False
  ExtraBallConsumed = False
    BossHitCount = 0
  ExtraBallEoBPending = 0
  wasExtraBall = False
  BallCritCount = 0
  BallWhirlwindScore = 0
  BallBackstabScore = 0
  LeapJackpot = 0
  ExtraBallMilestoneIdx = 0
  ExtraBallPending = 0
    LastEnteredInitials = ""
    AmbushPrizeBonus = 0
    GameActive = True
    Score = 0
    BallNumber = 0
    BIP = 0
    Dim strayBalls : strayBalls = GetBalls
    Dim sb
    For Each sb In strayBalls
        If sb.ID <> CapBallID And sb.ID <> CapBall2ID Then sb.DestroyBall
    Next
    ClearAllRampLoops
    GameOverMusicTimer.Enabled = False
    If CurrentSong <> "introedit" Then
        StopSound CurrentSong
        CurrentSong = ""
        PlaySong "introedit"
    End If
    If CurrentCalloutSound <> "" Then StopSound CurrentCalloutSound
    CurrentCalloutSound = ""
    DuckAudioTimer.Enabled = False
    CalloutQueueCount = 0
    CalloutPlaying = False
    StopAttractLighting
    InitParty
    ClearLootTargets          ' ensure no stale loot prims/flags from tutorial or prior state
  AmbushFailedPenalty = 0
  DifficultySelectActive = True
    SelectedBallCount = 5
    DifficultyLevel = 0
  ChestRampUp
    ShowDifficultySelect
    DifficultySelectTimer.Interval = 100
    DifficultySelectTimer.Enabled = True
End Sub

Sub ResetMultiballState()
    ' Party multiball - preserve pending spawn across drain if mid-spawn sequence
    If PartyMultiballRunning And PartySpawnStep < PartySpawnCount Then
        PartySpawnPending = True
    Else
        PartyMultiballRunning = False
        PartySpawnPending = False
    End If
    PartySpawnTimer.Enabled = False
    PortalOpenTimer.Enabled = False : PortalFlashTimer.Enabled = False
    townportalPF.Visible = False
    PlayfieldKicker.Enabled = False
    ' Runeword multiball - full reset every ball
    RuneWordMultiballRunning = False
    RuneWordSpawnTimer.Enabled = False
    RuneWordJackpotTimer.Enabled = False
    RuneWordMalusTimer.Enabled = False
    RuneWordSpawnCount = 0
    RuneKickBlinkTimer.Enabled = False
    RuneKickBlinkStep = 0
    ' Shared
    CubeRespawnHoldTimer.Enabled = False
End Sub

Sub LaunchNextBall()
  RestoreAudio
  TiltWarnings = 0
    TiltActive = False
  RampAoeCount = 0
    WhirlwindActive = False
  BallWhirlwindScore = 0
    BallBackstabScore = 0
    WhirlwindTimer.Enabled = False
    FlasherSweepTimer.Enabled = False
    UpdateRampAoeLight
    NudgeCount = 0
    NudgeWindowTimer.Enabled = False
    SetRampLight
  LeftFlipper.Enabled = True
  RightFlipper.Enabled = True
  wasExtraBall = False
  BallCritCount = 0
    BallSaveActive = False
    BallSaveTimer.Enabled = False
    MercSaveActive = False
    KillStreakCount = 0
    KillStreakTimer.Enabled = False
    ActiveElementType = -1
    ElementBallActive = False
    MysteryActive = False
    CalloutQueueCount = 0
    CalloutPlaying = False
    GIEventMode = GI_MODE_NONE
    GIEventTimer.Enabled = False
    RestoreAllGIRows
    If AmbushActive Then
        ResetAmbush
        TravelActive = True
    End If
    If MysteryReady Then
'        MysteryKickLight.BlinkInterval = 300
'        MysteryKickLight.State = 2
        QuestLight4.Color = RGB(180, 0, 255)
        QuestLight4.ColorFull = RGB(180, 0, 255)
        QuestLight4.BlinkInterval = 300
        QuestLight4.State = 2
        If Not ChestRampActive Then ChestRampDown
    Else
'        MysteryKickLight.State = 0
    End If
    LeapAwardTimer.Enabled = False
    If DifficultyLevel = 2 And BossHealth > 0 And BossHealth < BossMaxHP Then
        BossRegenTimer.Enabled = False
        BossRegenTimer.Interval = 60000
        BossRegenTimer.Enabled = True
    End If
    InstantKillActive = False
    InstantKillTimer.Enabled = False
    InstantKillTimer_step = 0
    AuraKillScore = 0
    AuraEndTimer.Enabled = False
    MysteryRampCloseTimer.Enabled = False
    MysteryRoll = -1
    MysteryAnimStep = 0
    MysteryAnimCycleCount = 0
    MysteryHoldTimer.Enabled = False
    MysteryAnimTimer.Enabled = False
    MysteryRevealTimer.Enabled = False
'    MysteryKickLight.State = 0
    ' Shout multiplier reset with hold support
    If ShoutMultiplierHeld Then
        ShoutMultiplierHeld = False
    ElseIf ShoutHoldQueueCount > 0 Then
        ShoutHoldQueueCount = ShoutHoldQueueCount - 1
        ShoutMultiplierHeld = True
    Else
        ShoutMultiplier = 1
    End If
    BallLootBonus = 0
    SafeTravelBank = 0
    BallKillCount = 0
    BallGoldCount = 0
    BallShieldCount = 0
    MercPinged = False
    MercPingPending = False
    MercPingTimer.Enabled = False
    MercPortalArmed = False
    MercMultiballActive = False
    MercHoldKey = 0
    MercHoldTimer.Enabled = False
    MercArmTimer.Enabled = False
    MercSpawnTimer.Enabled = False
    BallArenaKillCount = 0
    BallArenaKillScore = 0
    InitShoutLights
    ' Reset shield per ball, preserve SAVE lane progress
    ShieldReadyBlinkTimer.Enabled = False
    ShieldReadyBlinkStep = 0
    BallShieldCount = 0
    ShieldArmor = 0
    ShieldActive = False
    Shield.Collidable = False
    Shield.Visible = False
    ShieldLight.State = 0
    ShieldMaxArmor = 0
    ShieldHeld = False
    ' Reset gear per ball - save best N items first (N = ally count), then restore them
    If Not GearRetentionActive Then
        SaveRetainedGear
        Dim gri
        For gri = 0 To 6
            GearSlots(gri) = -1
            SlotToLight(gri).State = 0
        Next
        FirstUniqueEquipped = False
        RestoreRetainedGear
    End If
    GearRetentionActive = False
  StartCritCycle
    InitSaveLanes
    ResetMultiballState
    If TravelActive And Not BossFightActive And BallNumber > 1 Then SetTravelLights True
    BallSaveUsed = False
    FirstCubeHit = True
    ' Cube reset per ball — element back to cold, but a cranked-open cube stays open
    TransmuteLevel = 0
    If Not GemWallOpen Then
        GemQueueCount = 0
    End If
    PendingElementType = -1
    SilverBallID = -1
    CubeHoldTimer.Enabled = False
    CubeSpawnTimer.Enabled = False
    CubeTransmuteActive = False
    PoisonTickActive = False
    PoisonTimer.Enabled = False
    Dim bpln : For bpln = 0 To 4 : BumperPoisoned(bpln) = False : Next
    'If GemWallOpen Then CloseCubeWall   ' lid persists across balls until transmuted
    UpdateCubeLight
    ' Restore runeword light state
    If RuneWordReady Then
        RuneKickLight.BlinkInterval = 200
        RuneKickLight.BlinkPattern = "10"
        RuneKickLight.State = 2
        RuneLightR.BlinkInterval  = 200 : RuneLightR.State  = 2
        RuneLightU.BlinkInterval  = 200 : RuneLightU.State  = 2
        RuneLightN.BlinkInterval  = 200 : RuneLightN.State  = 2
        RuneLightE.BlinkInterval  = 200 : RuneLightE.State  = 2
        RuneLightW.BlinkInterval  = 200 : RuneLightW.State  = 2
        RuneLightO.BlinkInterval  = 200 : RuneLightO.State  = 2
        RuneLightR2.BlinkInterval = 200 : RuneLightR2.State = 2
        RuneLightD.BlinkInterval  = 200 : RuneLightD.State  = 2
        ShowMessage "RUNEWORD! SHOOT SCOOP!"
    Else
        If RuneHitR  Then RuneLightR.State  = 1
        If RuneHitU  Then RuneLightU.State  = 1
        If RuneHitN  Then RuneLightN.State  = 1
        If RuneHitE  Then RuneLightE.State  = 1
        If RuneHitW  Then RuneLightW.State  = 1
        If RuneHitO  Then RuneLightO.State  = 1
        If RuneHitR2 Then RuneLightR2.State = 1
        If RuneHitD  Then RuneLightD.State  = 1
    End If
    If ExtraBallPending > 0 Then
        ExtraBallPending = ExtraBallPending - 1
        wasExtraBall = True
        ExtraBallConsumed = True
        If ExtraBallEoBPending > 0 Then
            ExtraBallEoBPending = ExtraBallEoBPending - 1
            ShowBigMessage "EXTRA BALL!"
            PlayCallout "ExtraBall", 2000
            StartGIEvent GI_MODE_EXTRABALL
        End If
    Else
        BallNumber = BallNumber + 1
    End If
    If BallNumber = 1 And Not wasExtraBall Then
'        ChestRampUp
        StopSound CurrentSong
        CurrentSong = ""
        PlayCallout "Bar_act1_entry_wilderness", 3000
    End If
    If BallNumber > MaxBalls Then
        GameOver
        Exit Sub
    End If
    Select Case DifficultyLevel
        Case 1 : ShowMessage "NM  BALL " & BallNumber & " OF " & MaxBalls
        Case 2 : ShowMessage "HELL  BALL " & BallNumber & " OF " & MaxBalls
        Case Else : ShowMessage "BALL " & BallNumber & " OF " & MaxBalls
    End Select
    BallLaunchTimer.Enabled = True
End Sub

Sub BallLaunchTimer_Timer()
    BallLaunchTimer.Enabled = False
    If BallNumber = 1 And Not wasExtraBall Then PlaySong "town1"
    BallRelease.CreateBall
    BallRelease.Kick 90, 7
    PlaySoundAtLevelStatic SoundFX("BallRelease" & Int(Rnd * 7) + 1, DOFContactors), BallReleaseSoundLevel, BallRelease
    If TravelActive And Not BossFightActive Then SetTravelLights True
    If PartySpawnPending Then
        PartySpawnPending = False
        PartyMultiballRunning = True
        townportalPF.Visible = True
        PlaySound "portalenter", 0, 1
        ShowMessage "PARTY RESUMES!"
        StartGIEvent GI_MODE_MYSTERY
        PortalSpawnInterval = 1500
        PortalOpenTimer.Interval = 3000
        PortalOpenTimer.Enabled = True
        PortalFlashStep = 0
        PortalFlashTimer.Interval = 2000
        PortalFlashTimer.Enabled = True
    End If
End Sub

Dim SuppressBumperAnnounce : SuppressBumperAnnounce = False

Sub GameOver()
  StopSound CurrentSong
  ClearAllRampLoops
  CurrentSong = ""
  BallLaunchTimer.Enabled = False
  BossActTimer.Enabled = False        ' prevent queued boss/act music after game over
  FinalDifficulty = DifficultyLevel
  Dim goNeedsEntry : goNeedsEntry = False
  If FinalDifficulty = 1 And Score > NightmareChamp Then goNeedsEntry = True
  If FinalDifficulty = 2 And Score > HellChamp      Then goNeedsEntry = True
  If EnemiesKilled > HiKills(4)                     Then goNeedsEntry = True
  If FinalDifficulty = 0 And Score > HiScore(4)     Then goNeedsEntry = True
  GameOverSequenceActive = goNeedsEntry
  TiltActive = False
  CubeTransmuteActive = False
  DifficultyLevel = 0
  BossCritActive = False
    BossCritQueued = False
  If FlexDMDActive Then
        On Error Resume Next
        FlexDMD.LockRenderThread
        FlexDMD.Stage.GetGroup("Score").GetLabel("BigText").Visible = False
        FlexDMD.Stage.GetGroup("Score").GetLabel("Line1").Visible = True
        FlexDMD.Stage.GetGroup("Score").GetLabel("Line2").Visible = True
        FlexDMD.UnlockRenderThread
        On Error Goto 0
    End If
  PartySpawnPending = False
  DimAllFlashers
  CharWelcomeTimer.Enabled = False
  CharThanksTimer.Enabled = False
  ShieldHeld = False
  DuckAudioTimer.Enabled = False
  RestoreAudio
  MysteryActive = False
  CalloutQueueCount = 0
  CalloutPlaying = False
  ResetAmbush
  AmbushBossCount = 0
  AmbushKillsRequired = 10
  FirstUniqueEquipped = False
    LeapAwardTimer.Enabled = False
    ActCompleteTimer.Enabled = False
    BigTextTimer.Enabled = False
    BigTextActive = False
    InstantKillActive = False
    InstantKillTimer.Enabled = False
    InstantKillTimer_step = 0
    AuraKillScore = 0
    AuraEndTimer.Enabled = False
    NextBossOneHit = False
    LastRuneWordJackpot = 0
    MysteryRampCloseTimer.Enabled = False
    If ChestRamp.RotX > ChestRampLand Then
        ChestRampActive = True
        ChestRampState = 1
        ChestRampVel = 0.01
        ChestRampPhysics.Collidable = True
        ChestRampMove.Enabled = True
    End If
    If GemWallOpen Then CloseCubeWall
    ShoutMultiplier = 1
    ShoutMultiplierHeld = False
    ShoutHoldQueueCount = 0
    BallLootBonus = 0
  SafeTravelBank = 0
    MysteryKillCount = 0
    MysteryReady = False
    MysteryRoll = -1
    MysteryAnimStep = 0
    MysteryAnimCycleCount = 0
    MysteryHoldTimer.Enabled = False
    MysteryAnimTimer.Enabled = False
    MysteryRevealTimer.Enabled = False
'    MysteryKickLight.State = 0
    BallKillCount = 0
    BallGoldCount = 0
    BallShieldCount = 0
    BallArenaKillCount = 0
    BallArenaKillScore = 0
    BonusActive = False
    BonusTimer.Enabled = False
    ShoutPulseTimer.Enabled = False
    RuneWordMalusTimer.Enabled = False
    InitShoutLights
    ResetSaveSystem
    InitSaveLanes
    RuneWordJackpotLevel = 0
    RuneKickBlinkStep = 0
    RuneKickBlinkTimer.Enabled = False
  PartyMultiballRunning = False
    PartySpawnTimer.Enabled = False
    RuneWordMultiballRunning = False
    RuneWordSpawnTimer.Enabled = False
    RuneWordJackpotTimer.Enabled = False
    RuneWordSpawnCount = 0
  RuneQuality = 0
    ResetRuneTargets
    ClearLootTargets          ' drop any raised loot targets + sink/clear their prims
    CubeRespawnHoldTimer.Enabled = False
    ActiveElementType = -1
  ElementBallActive = False
    PendingElementType = -1
    SilverBallID = -1
    CubeHoldTimer.Enabled = False
    CubeSpawnTimer.Enabled = False
    CubeTransmuteActive = False
    PoisonTickActive = False
    PoisonTimer.Enabled = False
    Dim bpgo : For bpgo = 0 To 4 : BumperPoisoned(bpgo) = False : Next
    ElementBallActive = False
    GameActive = False
    FirstPackDone = False
    If Score > HighScore Then HighScore = Score
    ShowMessage "GAME OVER!"
    GIEventMode = GI_MODE_NONE
    GIEventTimer.Enabled = False
    StartGIEvent GI_MODE_GAMEOVER
    GameOverTimer.Enabled = True
    Dim finalKills : finalKills = EnemiesKilled
    InitParty
    EnemiesKilled = finalKills
    BallNumber = 0
    TransmuteLevel = 0
    BIP = 0
    LeapBallID = -1
    GemQueueCount = 0
    GemWallOpen = False
    CubeLight.State = 0
    KillMilestoneCount = 0
    RampGemCount = 0
    CubeSpinCount = 0
    CrankSndTimer.Enabled = False
    If CrankSndOn Then
        StopSound "hellbrazierloop2"
        CrankSndOn = False
    End If
    CubeGate.Collidable = True
    ' lower the lid with an animation instead of snapping
    If CubeLid.TransY > CubeLidClosedY Then
        CubeLidTargetY = CubeLidClosedY
        If Not CubeLidTimer.Enabled Then
            CubeLidTimer.Enabled = True
'            PlaySound "hellbrazierloop2", 0, 1
        End If
    Else
        CubeLidTargetY = CubeLidClosedY
        CubeLid.Collidable = True
        Wall015.Collidable = True
    End If
    PoisonTickActive = False
    PoisonTimer.Enabled = False
    Dim gci
    For gci = 0 To 6
        GearSlots(gci) = -1
        SlotToLight(gci).State = 0
    Next
    Dim lci
    For lci = 0 To 3
        LootActive(lci) = False
        LootPending(lci) = -1
        LootValue(lci) = 0
        LootTier(lci) = -1
        dtLoot1.IsDropped = True
        dtLoot2.IsDropped = True
        dtLoot3.IsDropped = True
        dtLoot4.IsDropped = True
    Next
    ActiveElementType = -1
    ElementBallActive = False
    MagicFind = 0
    SuppressBumperAnnounce = True
    AssignBumperStats 0, True
    AssignBumperStats 1, False
    AssignBumperStats 2, False
    AssignBumperStats 3, False
    AssignBumperStats 4, False
    TravelProgress = 0
    EventIndex = 0
    BossFightActive = False
    BossEventType = -1
    TravelActive = True
    BossHealth = 1
    CapBallLight.State = 0
    BossHPPulseTimer.Enabled = False
    BossRegenTimer.Enabled = False
    StopCritCycle
    BossFightActive = False
    Dim goi : For goi = 0 To 7 : BossHP.Item(goi).State = 0 : Next
    FireBurnSlot = -1
    FireBurnTimer.Enabled = False
    If WasActSixComplete Then
        PlayCallout "Bar_goodbye", 1000  ' use a non-antagonistic callout you have
    Else
        Select Case Int(Rnd * 4)
            Case 0 : PlayCallout "Betterluck", 2000
            Case 1 : PlayCallout "DeathInev", 4000
            Case 2 : PlayCallout "DemiseExpected", 2000
            Case 3 : PlayCallout "HellIsNoPlace", 3000
        End Select
    End If
    GameOverMusicTimer.Interval = 3000
    GameOverMusicTimer.Enabled = True
End Sub

Sub GameOverMusicTimer_Timer()
    GameOverMusicTimer.Enabled = False
    If GameActive Then Exit Sub
    PlaySong "introedit"
End Sub

Sub GameOverTimer_Timer()
    GameOverTimer.Enabled = False
    If GameActive Then Exit Sub
    UpdateDMD "FINAL: " & FormatNumber(Score, 0, -1, 0, -1)
    GameOverScoreTimer.Enabled = True
End Sub

Sub GameOverScoreTimer_Timer()
    GameOverScoreTimer.Enabled = False
    If GameActive Then Exit Sub
    CheckDifficultyHighScore
End Sub
'*******************************************
' ZFLP: Flippers
'*******************************************

Const ReflipAngle = 20

' Flipper Solenoid Callbacks (these subs mimics how you would handle flippers in ROM based tables)
Sub SolLFlipper(Enabled)
    If Not GameActive And Not TutorialActive Then Exit Sub
    If Enabled Then
        FlipperActivate LeftFlipper, LFPress
    FlipperActivate UpperLeftFlipper, ULFPress
    LF.Fire  'leftflipper.rotatetoend
    UpperLeftFlipper.RotateToEnd
    If leftflipper.currentangle < leftflipper.endangle + ReflipAngle Then
      RandomSoundReflipUpLeft LeftFlipper
    Else
      SoundFlipperUpAttackLeft LeftFlipper
      RandomSoundFlipperUpLeft LeftFlipper
    End If
  Else
    FlipperDeActivate LeftFlipper, LFPress
    FlipperDeactivate UpperLeftFlipper, ULFPress
    LeftFlipper.RotateToStart
    UpperLeftFlipper.RotateToStart
    If LeftFlipper.currentangle < LeftFlipper.startAngle - 5 Then
      RandomSoundFlipperDownLeft LeftFlipper
    End If
    FlipperLeftHitParm = FlipperUpSoundLevel
  End If
End Sub

Sub SolRFlipper(Enabled)
    If Not GameActive And Not TutorialActive Then Exit Sub
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

Sub UpperLeftFlipper_Collide(parm)
    CheckLiveCatch ActiveBall, UpperLeftFlipper, ULFCount, parm
    RandomSoundRubberFlipper parm
End Sub

Sub RightFlipper_Collide(parm)
  CheckLiveCatch ActiveBall, RightFlipper, RFCount, parm
  RF.ReProcessBalls ActiveBall
  RightFlipperCollide parm
End Sub

'******************************************************
'  ZANI: Misc Animations
'******************************************************

Sub LeftFlipper_Animate
  dim a: a = LeftFlipper.CurrentAngle
  FlipperLSh.RotZ = a
  'Add any left flipper related animations here
End Sub

Sub RightFlipper_Animate
  dim a: a = RightFlipper.CurrentAngle
  FlipperRSh.RotZ = a
  'Add any right flipper related animations here
End Sub




'*******************************************
' ZKEY: Key Press Handling
'*******************************************

Sub Table1_KeyDown(ByVal keycode)

If keycode = 47 Then ShowLeakDebug      ' V — ball/ramp/timer readout
    If keycode = 48 Then DebugAudioFlush    ' B — flush all loop voices

If DifficultySelectActive Then
        If keycode = LeftFlipperKey Then
            Select Case SelectedBallCount
                Case 5 : SelectedBallCount = 1
                Case 3 : SelectedBallCount = 5
                Case 1 : SelectedBallCount = 3
            End Select
            ShowDifficultySelect
        ElseIf keycode = RightFlipperKey Then
            Select Case SelectedBallCount
                Case 5 : SelectedBallCount = 3
                Case 3 : SelectedBallCount = 1
                Case 1 : SelectedBallCount = 5
            End Select
            ShowDifficultySelect
        ElseIf keycode = StartGameKey Then
            DifficultySelectActive = False
            DifficultySelectTimer.Enabled = False
      If CurrentSong <> "" Then StopSound CurrentSong
      StopSound "introedit"
      CurrentSong = ""
            Select Case SelectedBallCount
                Case 5 : DifficultyLevel = 0 : MaxBalls = 3
                Case 3 : DifficultyLevel = 1 : MaxBalls = 2
                Case 1 : DifficultyLevel = 2 : MaxBalls = 1
            End Select
            LaunchNextBall
        End If
        Exit Sub
    End If

  If DifficultyHSActive Then
    If keycode = LeftFlipperKey Then
        DifficultyHSLetter = DifficultyHSLetter - 1
        If DifficultyHSLetter = 0 Then DifficultyHSLetter = Len(hsValidLetters)
        DifficultyHSDisplayName
    End If
    If keycode = RightFlipperKey Then
        DifficultyHSLetter = DifficultyHSLetter + 1
        If DifficultyHSLetter > Len(hsValidLetters) Then DifficultyHSLetter = 1
        DifficultyHSDisplayName
    End If
    If keycode = 13 Or keycode = 28 Or keycode = StartGameKey Then
        Dim dLetter : dLetter = Mid(hsValidLetters, DifficultyHSLetter, 1)
        If dLetter = "<" Then
            If DifficultyHSDigit > 0 Then
                DifficultyHSDigit = DifficultyHSDigit - 1
                hsEnteredDigits(DifficultyHSDigit) = "-"
            End If
            DifficultyHSDisplayName
        Else
            hsEnteredDigits(DifficultyHSDigit) = dLetter
            DifficultyHSDigit = DifficultyHSDigit + 1
            If DifficultyHSDigit >= 3 Then
                DifficultyHSCommit
            Else
                DifficultyHSDisplayName
            End If
        End If
    End If
    Exit Sub
End If

'    ' *** TEMP TEST: press 'L' to max gear + force a boss fight (manual crit cycle test) - REMOVE BEFORE RELEASE ***
'    If keycode = 38 Then
'        If GameActive Then
'            Dim tgi
'            For tgi = 0 To 6
'                GearSlots(tgi) = GEAR_UNIQUE
'            Next
'            If Not BossFightActive Then
'                TravelActive = False
'                StartBossFight BOSS_BLOODRAVEN
'            End If
'            ShowMessage "TEST: ALL UNIQUES + BOSS FIGHT"
'        End If
'    End If
'    ' *** END TEMP TEST ***

    If keycode = LeftMagnaSave Or keycode = RightMagnaSave Then
        If LeapReady Then
            FireLeapAttack
        ElseIf CritCycleActive And AllUniquesEquipped() Then
            CritCycleStep = (CritCycleStep + 1) Mod 3
            SetCritCycleLight CritCycleStep
        End If
    End If

If keycode = 13 Or keycode = 28 Then
        If hsbKillsModeActive Then
            EnterKillsHighScoreKey keycode
        ElseIf hsbModeActive Then
            EnterHighScoreKey keycode
        End If
    End If
    If keycode = 2 Or keycode = StartGameKey Then
        If TutorialActive Then
            CancelTutorial
            Exit Sub
        End If
        If hsbKillsModeActive Then
            EnterKillsHighScoreKey keycode
        ElseIf hsbModeActive Then
            EnterHighScoreKey keycode
        ElseIf Not GameActive And Not GameOverSequenceActive Then
            StartGame
        End If
    End If

' If keycode = 19 Then
'   ScoreCard = 1
'   CardTimer.enabled = True
' End If

If keycode = LeftFlipperKey Then
    If Not GameActive And Not hsbModeActive And Not hsbKillsModeActive Then
        TutLeftHeld = True
        If TutRightHeld And Not TutorialActive And Not TutHoldPending Then
            TutHoldPending = True
            TutorialHoldTimer.Interval = 2000
            TutorialHoldTimer.Enabled = True
            Exit Sub
        End If
        If TutorialActive Then
            TutorialStep = TutorialStep - 1
            If TutorialStep < 0 Then TutorialStep = 0
            ShowTutorialStep TutorialStep
            Exit Sub
        End If
        AttractManualMode = True
        AttractState = AttractState - 1
        If AttractState < 0 Then AttractState = 13
        ShowAttractFrame
        AttractManualResumeTimer.Interval = 5000
        AttractManualResumeTimer.Enabled = True
        Exit Sub
    End If
    If TiltActive And Not BonusActive Then Exit Sub
    AdvanceBonusStep
    If hsbKillsModeActive Then
        EnterKillsHighScoreKey keycode
    ElseIf hsbModeActive Then
        EnterHighScoreKey keycode
    Else
        If Not TiltActive Then
            SolLFlipper True
            CycleSaveLanesLeft
        End If
    End If
End If

If keycode = RightFlipperKey Then
    If Not GameActive And Not hsbModeActive And Not hsbKillsModeActive Then
        TutRightHeld = True
        If TutLeftHeld And Not TutorialActive And Not TutHoldPending Then
            TutHoldPending = True
            TutorialHoldTimer.Interval = 2000
            TutorialHoldTimer.Enabled = True
            Exit Sub
        End If
        If TutorialActive Then
            TutorialStep = TutorialStep + 1
            If TutorialStep >= 12 Then
                CancelTutorial
            Else
                ShowTutorialStep TutorialStep
            End If
            Exit Sub
        End If
        AttractManualMode = True
        AttractState = AttractState + 1
        If AttractState > 14 Then AttractState = 0
        ShowAttractFrame
        AttractManualResumeTimer.Interval = 5000
        AttractManualResumeTimer.Enabled = True
        Exit Sub
    End If
    If TiltActive And Not BonusActive Then Exit Sub
    AdvanceBonusStep
    If hsbKillsModeActive Then
        EnterKillsHighScoreKey keycode
    ElseIf hsbModeActive Then
        EnterHighScoreKey keycode
    Else
        If Not TiltActive Then
            SolRFlipper True
            CycleSaveLanesRight
        End If
    End If
End If

  If keycode = PlungerKey Then
    Plunger.Pullback
    VRPlungerBtnHeld = True
    SoundPlungerPull
End If

  If keycode = LeftTiltKey Then
    If GameActive And Not TiltActive Then Nudge 90, 3 : CheckTilt
    SoundNudgeLeft
  End If
  If keycode = RightTiltKey Then
    If GameActive And Not TiltActive Then Nudge 270, 3 : CheckTilt
    SoundNudgeRight
  End If
  If keycode = CenterTiltKey Then
    If GameActive And Not TiltActive Then Nudge 0, 3 : CheckTilt
    SoundNudgeCenter
  End If
  If keycode = MechanicalTilt Then
    SoundNudgeCenter() 'Send the Tilting command to the ROM (usually by pulsing a Switch), or run the tilting code for an orginal table
  End If

  If keycode = StartGameKey Then
        If Not GameActive Then SoundStartButton
    End If

  '   If keycode = keyInsertCoin1 or keycode = keyInsertCoin2 or keycode = keyInsertCoin3 or keycode = keyInsertCoin4 Then 'Use this for ROM based games
  If keycode = AddCreditKey Or keycode = AddCreditKey2 Then
        Select Case Int(Rnd * 3)
            Case 0 : PlaySound ("Coin_In_1"), 0, CoinSoundLevel, 0, 0.25
            Case 1 : PlaySound ("Coin_In_2"), 0, CoinSoundLevel, 0, 0.25
            Case 2 : PlaySound ("Coin_In_3"), 0, CoinSoundLevel, 0, 0.25
        End Select
        If Not FreePlay Then
            Credits = Credits + 1
            SaveValue "D2", "Credits", Credits
            If GameActive Then
                ShowMessage "CREDITS: " & Credits
            Else
                ShowMysteryDMD "CREDITS", Credits & "   PRESS START"
            End If
            CreditDisplayTimer.Enabled = False
            CreditDisplayTimer.Enabled = True
        End If
    End If
End Sub

Sub CreditDisplayTimer_Timer()
    CreditDisplayTimer.Enabled = False
    If Not GameActive Then ShowAttractFrame
End Sub



Sub Table1_KeyUp(ByVal keycode)

' If keycode = 19 Then ScoreCard = 0

  If KeyCode = PlungerKey Then
    Plunger.Fire
    VRPlungerBtnHeld = False
    If BIP = 1 Then
        SoundPlungerReleaseBall()
    Else
        SoundPlungerReleaseNoBall()
    End If
End If

  If keycode = LeftFlipperKey Then
        TutLeftHeld = False
        If TutHoldPending Then TutorialHoldTimer.Enabled = False : TutHoldPending = False
        SolLFlipper False
    End If

    If keycode = RightFlipperKey Then
        TutRightHeld = False
        If TutHoldPending Then TutorialHoldTimer.Enabled = False : TutHoldPending = False
        SolRFlipper False
    End If
End Sub



'Sub Table1_KeyDown(ByVal keycode)
'    If keycode = PlungerKey Then
'        If EnableRetractPlunger Then
'            Plunger.PullBackandRetract
'        Else
'            Plunger.PullBack
'        End If
'        PlaySoundAtLevelStatic "Plunger_Pull_1", PlungerPullSoundLevel, Plunger
'    End If
'
'    If keycode = 13 Or keycode = 28 Then
'        If hsbKillsModeActive Then
'            EnterKillsHighScoreKey keycode
'        ElseIf hsbModeActive Then
'            EnterHighScoreKey keycode
'        ElseIf Not GameActive Then
'            StartGame
'        End If
'    End If
'
'   If keycode = LeftFlipperKey Then
'   AdvanceBonusStep
'        If hsbKillsModeActive Then
'            EnterKillsHighScoreKey keycode
'        ElseIf hsbModeActive Then
'            EnterHighScoreKey keycode
'        Else
'            LF.Fire
'            UpperLeftFlipper.RotateToEnd
'            FlipperActivate LeftFlipper, LFPress
''            SoundFlipperUpAttackLeft
'            RandomSoundFlipperUpLeft
'            CycleSaveLanesLeft
'        End If
'    End If
'
'    If keycode = RightFlipperKey Then
'   AdvanceBonusStep
'        If hsbKillsModeActive Then
'            EnterKillsHighScoreKey keycode
'        ElseIf hsbModeActive Then
'            EnterHighScoreKey keycode
'        Else
'            RF.Fire
'            FlipperActivate RightFlipper, RFPress
''            SoundFlipperUpAttackRight
'            RandomSoundFlipperUpRight
'            CycleSaveLanesRight
'        End If
'    End If
'
'    If keycode = LeftMagnaSave Or keycode = RightMagnaSave Then
'        'ActivateShield
'    End If
'
'    If keycode = LeftTiltKey Then
'        Nudge 90, 2
'        SoundNudgeLeft
'    End If
'    If keycode = RightTiltKey Then
'        Nudge 270, 2
'        SoundNudgeRight
'    End If
'    If keycode = CenterTiltKey Then
'        Nudge 0, 2
'        SoundNudgeCenter
'    End If
'End Sub
'
'Sub Table1_KeyUp(ByVal keycode)
'    If keycode = PlungerKey Then
'        Plunger.Fire
'        PlaySoundAtLevelStatic "Plunger_Release_Ball", PlungerReleaseSoundLevel, Plunger
'        StopSound "introedit"
'        PlaySong "town1"
'    End If
'   If keycode = LeftFlipperKey Then
'    LeftFlipper.RotateToStart
'    UpperLeftFlipper.RotateToStart
'    FlipperDeactivate LeftFlipper, LFPress
'    RandomSoundFlipperDownLeft
'End If
'
'If keycode = RightFlipperKey Then
'    RightFlipper.RotateToStart
'    FlipperDeactivate RightFlipper, RFPress
'    RandomSoundFlipperDownRight
'End If
'    If keycode = 46 Then
'        If EnableBallControl = 1 Then
'            EnableBallControl = 0
'        Else
'            EnableBallControl = 1
'        End If
'    End If
'    If EnableBallControl = 1 Then
'        If keycode = 48 Then
'            If BCboost = 1 Then
'                BCboost = BCboostmulti
'            Else
'                BCboost = 1
'            End If
'        End If
'        If keycode = 203 Then BCleft = 1
'        If keycode = 200 Then BCup = 1
'        If keycode = 208 Then BCdown = 1
'        If keycode = 205 Then BCright = 1
'    End If
'
'If keycode = LeftMagnaSave Or keycode = RightMagnaSave Then
'        ' no action on keyup for magna save
'    End If
'
'End Sub

'******************************************************
'   ZNFF: FLIPPER CORRECTIONS by nFozzy/rothbauerw
'******************************************************

Class FlipperPolarity
    Public DebugOn, Enabled
    Private FlipAt
    Public TimeDelay
    Private Flipper, FlipperStart, FlipperEnd, FlipperEndY, LR, PartialFlipCoef, FlipStartAngle
    Private Balls(20), balldata(20)
    Private Name
    Dim PolarityIn, PolarityOut
    Dim VelocityIn, VelocityOut
    Dim YcoefIn, YcoefOut

    Public Sub Class_Initialize
        ReDim PolarityIn(0) : ReDim PolarityOut(0)
        ReDim VelocityIn(0) : ReDim VelocityOut(0)
        ReDim YcoefIn(0) : ReDim YcoefOut(0)
        Enabled = True : TimeDelay = 50 : LR = 1
        Dim x
        For x = 0 To UBound(balls)
            balls(x) = Empty
            Set Balldata(x) = New SpoofBall
        Next
    End Sub

    Public Sub SetObjects(aName, aFlipper, aTrigger)
        If TypeName(aName) <> "String" Then MsgBox "FlipperPolarity: first argument must be a String" End If
        If TypeName(aFlipper) <> "Flipper" Then MsgBox "FlipperPolarity: second argument must be a flipper" End If
        If TypeName(aTrigger) <> "Trigger" Then MsgBox "FlipperPolarity: third argument must be a trigger" End If
        If aFlipper.EndAngle > aFlipper.StartAngle Then LR = -1 Else LR = 1 End If
        Name = aName
        Set Flipper = aFlipper
        FlipperStart = aFlipper.x
        FlipperEnd = Flipper.Length * Sin((Flipper.StartAngle / 57.295779513082320876798154814105)) + Flipper.X
        FlipperEndY = Flipper.Length * Cos(Flipper.StartAngle / 57.295779513082320876798154814105) * -1 + Flipper.Y
        Dim str
        str = "Sub " & aTrigger.Name & "_Hit() : " & aName & ".AddBall ActiveBall : End Sub'"
        ExecuteGlobal(str)
        str = "Sub " & aTrigger.Name & "_UnHit() : " & aName & ".PolarityCorrect ActiveBall : End Sub'"
        ExecuteGlobal(str)
    End Sub

    Public Property Let EndPoint(aInput)
    End Property

    Public Sub AddPt(aChooseArray, aIDX, aX, aY)
        Select Case aChooseArray
            Case "Polarity"
                ShuffleArrays PolarityIn, PolarityOut, 1
                PolarityIn(aIDX) = aX : PolarityOut(aIDX) = aY
                ShuffleArrays PolarityIn, PolarityOut, 0
            Case "Velocity"
                ShuffleArrays VelocityIn, VelocityOut, 1
                VelocityIn(aIDX) = aX : VelocityOut(aIDX) = aY
                ShuffleArrays VelocityIn, VelocityOut, 0
            Case "Ycoef"
                ShuffleArrays YcoefIn, YcoefOut, 1
                YcoefIn(aIDX) = aX : YcoefOut(aIDX) = aY
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
    On Error Resume Next
    For x = 0 To UBound(balls)
        If TypeName(balls(x)) = "IBall" Then
            If aBall.ID = Balls(x).ID Then
                balls(x) = Empty
                Balldata(x).Reset
            End If
        End If
    Next
    On Error Goto 0
End Sub

    Public Sub Fire()
        Flipper.RotateToEnd
        ProcessBalls
    End Sub

    Public Property Get Pos
        Dim x
        For x = 0 To UBound(balls)
            If Not IsEmpty(balls(x)) Then
                pos = PSlope(Balls(x).x, FlipperStart, 0, FlipperEnd, 1)
            End If
        Next
    End Property

    Public Sub ProcessBalls()
        FlipAt = GameTime
        Dim x
        For x = 0 To UBound(balls)
            If Not IsEmpty(balls(x)) Then
                balldata(x).Data = balls(x)
            End If
        Next
        FlipStartAngle = Flipper.CurrentAngle
        PartialFlipCoef = ((Flipper.StartAngle - Flipper.CurrentAngle) / (Flipper.StartAngle - Flipper.EndAngle))
        PartialFlipCoef = Abs(PartialFlipCoef - 1)
    End Sub

    Public Sub ReProcessBalls(aBall)
        If FlipperOn() Then
            Dim x
            For x = 0 To UBound(balls)
                If Not IsEmpty(balls(x)) Then
                    If balls(x).ID = aBall.ID Then
                        If IsEmpty(balldata(x).ID) Then
                            balldata(x).Data = balls(x)
                        End If
                    End If
                End If
            Next
        End If
    End Sub

    Private Function FlipperOn()
        If GameTime < FlipAt + TimeDelay Then FlipperOn = True End If
    End Function

    Public Sub PolarityCorrect(aBall)
        If FlipperOn() Then
            Dim tmp, BallPos, x, IDX, Ycoef, BalltoFlip, BalltoBase, NoCorrection, checkHit
            Ycoef = 1
            If aBall.VelY > -8 Then
                RemoveBall aBall
                Exit Sub
            End If
            For x = 0 To UBound(Balls)
                If aBall.id = BallData(x).id And Not IsEmpty(BallData(x).id) Then
                    idx = x
                    BallPos = PSlope(BallData(x).x, FlipperStart, 0, FlipperEnd, 1)
                    BalltoFlip = DistanceFromFlipperAngle(BallData(x).x, BallData(x).y, Flipper, FlipStartAngle)
                    If ballpos > 0.65 Then Ycoef = LinearEnvelope(BallData(x).Y, YcoefIn, YcoefOut)
                End If
            Next
            If BallPos = 0 Then
                BallPos = PSlope(aBall.x, FlipperStart, 0, FlipperEnd, 1)
                If ballpos > 0.65 Then Ycoef = LinearEnvelope(aBall.Y, YcoefIn, YcoefOut)
                NoCorrection = 1
            Else
                checkHit = 50 + (20 * BallPos)
                If BalltoFlip > checkHit Or (PartialFlipCoef < 0.5 And BallPos > 0.22) Then
                    NoCorrection = 1
                Else
                    NoCorrection = 0
                End If
            End If
            If Not IsEmpty(VelocityIn(0)) Then
                Dim VelCoef
                VelCoef = LinearEnvelope(BallPos, VelocityIn, VelocityOut)
                If Enabled Then aBall.Velx = aBall.Velx * VelCoef
                If Enabled Then aBall.Vely = aBall.Vely * VelCoef
            End If
            If Not IsEmpty(PolarityIn(0)) Then
                Dim AddX
                AddX = LinearEnvelope(BallPos, PolarityIn, PolarityOut) * LR
                If Enabled And NoCorrection = 0 Then aBall.VelX = aBall.VelX + 1 * (AddX * Ycoef * PartialFlipCoef * VelCoef)
            End If
            If DebugOn Then Debug.Print "PolarityCorrect " & Name & " @ " & GameTime & " " & Round(BallPos*100) & "%" & " AddX:" & Round(AddX,2) & " Vel%:" & Round(VelCoef*100)
        End If
        RemoveBall aBall
    End Sub
End Class

'******************************************************
'  FLIPPER POLARITY SUPPORTING FUNCTIONS
'******************************************************

Sub ShuffleArray(ByRef aArray, ByVal offset)
    Dim x, aCount : aCount = 0
    ReDim a(UBound(aArray))
    For x = 0 To UBound(aArray)
        If Not IsEmpty(aArray(x)) Then
            If IsObject(aArray(x)) Then
                Set a(aCount) = aArray(x)
            Else
                a(aCount) = aArray(x)
            End If
            aCount = aCount + 1
        End If
    Next
    If offset < 0 Then offset = 0
    ReDim aArray(aCount - 1 + offset)
    For x = 0 To aCount - 1
        If IsObject(a(x)) Then
            Set aArray(x) = a(x)
        Else
            aArray(x) = a(x)
        End If
    Next
End Sub

Sub ShuffleArrays(aArray1, aArray2, offset)
    ShuffleArray aArray1, offset
    ShuffleArray aArray2, offset
End Sub

Function BallSpeed(ball)
    BallSpeed = Sqr(ball.VelX^2 + ball.VelY^2 + ball.VelZ^2)
End Function

Function PSlope(Input, X1, Y1, X2, Y2)
    Dim x, y, b, m
    x = Input : m = (Y2 - Y1) / (X2 - X1) : b = Y2 - m * X2
    Y = m * x + b
    PSlope = Y
End Function

Class SpoofBall
    Public X, Y, Z, VelX, VelY, VelZ, ID, Mass, Radius
    Public Property Let Data(aBall)
        With aBall
            x = .x : y = .y : z = .z
            velx = .velx : vely = .vely : velz = .velz
            id = .ID : mass = .mass : radius = .radius
        End With
    End Property
    Public Sub Reset()
        x = Empty : y = Empty : z = Empty
        velx = Empty : vely = Empty : velz = Empty
        id = Empty : mass = Empty : radius = Empty
    End Sub
End Class

Function LinearEnvelope(xInput, xKeyFrame, yLvl)
    Dim y, L, ii
    For ii = 1 To UBound(xKeyFrame)
        If xInput <= xKeyFrame(ii) Then L = ii : Exit For
    Next
    If xInput > xKeyFrame(UBound(xKeyFrame)) Then L = UBound(xKeyFrame)
    Y = PSlope(xInput, xKeyFrame(L-1), yLvl(L-1), xKeyFrame(L), yLvl(L))
    If xInput <= xKeyFrame(LBound(xKeyFrame)) Then Y = yLvl(LBound(xKeyFrame))
    If xInput >= xKeyFrame(UBound(xKeyFrame)) Then Y = yLvl(UBound(xKeyFrame))
    LinearEnvelope = Y
End Function

Function Radians(Degrees)
    Radians = Degrees * PI / 180
End Function

Function AnglePP(ax, ay, bx, by)
    AnglePP = Atn2((by - ay), (bx - ax)) * 180 / PI
End Function

Function Distance(ax, ay, bx, by)
    Distance = Sqr((ax - bx)^2 + (ay - by)^2)
End Function

Function DistancePL(px, py, ax, ay, bx, by)
    DistancePL = Abs((by - ay) * px - (bx - ax) * py + bx * ay - by * ax) / Distance(ax, ay, bx, by)
End Function

Function DistanceFromFlipper(ballx, bally, Flipper)
    DistanceFromFlipper = DistancePL(ballx, bally, Flipper.x, Flipper.y, Cos(Radians(Flipper.CurrentAngle + 90)) + Flipper.x, Sin(Radians(Flipper.CurrentAngle + 90)) + Flipper.y)
End Function

Function DistanceFromFlipperAngle(ballx, bally, Flipper, Angle)
    DistanceFromFlipperAngle = DistancePL(ballx, bally, Flipper.x, Flipper.y, Cos(Radians(Angle + 90)) + Flipper.x, Sin(Radians(Angle + 90)) + Flipper.y)
End Function

Function FlipperTrigger(ballx, bally, Flipper)
    Dim DiffAngle
    DiffAngle = Abs(Flipper.CurrentAngle - AnglePP(Flipper.x, Flipper.y, ballx, bally) - 90)
    If DiffAngle > 180 Then DiffAngle = DiffAngle - 360
    If DistanceFromFlipper(ballx, bally, Flipper) < 48 And DiffAngle <= 90 And Distance(ballx, bally, Flipper.x, Flipper.y) < Flipper.Length Then
        FlipperTrigger = True
    Else
        FlipperTrigger = False
    End If
End Function

'Function Atn2(dy, dx)
'    Dim pi2 : pi2 = 4 * Atn(1)
'    If dx > 0 Then
'        Atn2 = Atn(dy / dx)
'    ElseIf dx < 0 Then
'        If dy = 0 Then
'            Atn2 = pi2
'        Else
'            Atn2 = Sgn(dy) * (pi2 - Atn(Abs(dy / dx)))
'        End If
'    ElseIf dx = 0 Then
'        If dy = 0 Then
'            Atn2 = 0
'        Else
'            Atn2 = Sgn(dy) * pi2 / 2
'        End If
'    End If
'End Function

'******************************************************
'  FLIPPER TRICKS
'******************************************************

RightFlipper.TimerInterval = 1
RightFlipper.TimerEnabled = True

Sub RightFlipper_Timer()
    FlipperTricks LeftFlipper, LFPress, LFCount, LFEndAngle, LFState
    FlipperTricks RightFlipper, RFPress, RFCount, RFEndAngle, RFState
    FlipperTricks UpperLeftFlipper, ULFPress, ULFCount, ULFEndAngle, ULFState  ' ← add
    FlipperNudge RightFlipper, RFEndAngle, RFEOSNudge, LeftFlipper, LFEndAngle
    FlipperNudge LeftFlipper, LFEndAngle, LFEOSNudge, RightFlipper, RFEndAngle
End Sub
Dim LFEOSNudge, RFEOSNudge

Sub FlipperNudge(Flipper1, EndAngle1, EOSNudge1, Flipper2, EndAngle2)
    Dim BOT, b
    If Flipper1.CurrentAngle = EndAngle1 And EOSNudge1 <> 1 Then
        EOSNudge1 = 1
        If Flipper2.CurrentAngle = EndAngle2 Then
            BOT = GetBalls
            For b = 0 To UBound(BOT)
                If FlipperTrigger(BOT(b).x, BOT(b).y, Flipper1) Then Exit Sub
            Next
            For b = 0 To UBound(BOT)
                If FlipperTrigger(BOT(b).x, BOT(b).y, Flipper2) Then
                    BOT(b).velx = BOT(b).velx / 1.7
                    BOT(b).vely = BOT(b).vely - 1
                End If
            Next
        End If
    Else
        If Flipper1.CurrentAngle <> EndAngle1 Then EOSNudge1 = 0
    End If
End Sub

Sub FlipperActivate(Flipper, FlipperPress)
    FlipperPress = 1
    Flipper.Elasticity = FElasticity
    Flipper.EOSTorque = EOST
    Flipper.EOSTorqueAngle = EOSA
End Sub

Sub FlipperDeactivate(Flipper, FlipperPress)
    Dim BOT, b
    FlipperPress = 0
    Flipper.EOSTorqueAngle = EOSA
    Flipper.EOSTorque = EOST * EOSReturn / FReturn
    If Abs(Flipper.CurrentAngle) <= Abs(Flipper.EndAngle) + 0.1 Then
        BOT = GetBalls
        If IsArray(BOT) Then
            For b = 0 To UBound(BOT)
                If IsObject(BOT(b)) Then
                    If Distance(BOT(b).x, BOT(b).y, Flipper.x, Flipper.y) < 55 Then
                        If BOT(b).VelY >= -0.4 Then BOT(b).VelY = -0.4
                    End If
                End If
            Next
        End If
    End If
End Sub

Sub FlipperTricks(Flipper, FlipperPress, FCount, FEndAngle, FState)
    Dim Dir
    Dir = Flipper.StartAngle / Abs(Flipper.StartAngle)
    If Abs(Flipper.CurrentAngle) > Abs(Flipper.StartAngle) - 0.05 Then
        If FState <> 1 Then
            Flipper.RampUp = SOSRampup
            Flipper.EndAngle = FEndAngle - 3 * Dir
            Flipper.Elasticity = FElasticity * SOSEM
            FCount = 0
            FState = 1
        End If
    ElseIf Abs(Flipper.CurrentAngle) <= Abs(Flipper.EndAngle) And FlipperPress = 1 Then
        If FCount = 0 Then FCount = GameTime
        If FState <> 2 Then
            Flipper.EOSTorqueAngle = EOSAnew
            Flipper.EOSTorque = EOSTnew
            Flipper.RampUp = EOSRampup
            Flipper.EndAngle = FEndAngle
            FState = 2
        End If
    ElseIf Abs(Flipper.CurrentAngle) > Abs(Flipper.EndAngle) + 0.01 And FlipperPress = 1 Then
        If FState <> 3 Then
            Flipper.EOSTorque = EOST
            Flipper.EOSTorqueAngle = EOSA
            Flipper.RampUp = Frampup
            Flipper.Elasticity = FElasticity
            FState = 3
        End If
    End If
End Sub

Sub CheckLiveCatch(ball, Flipper, FCount, parm)
    Dim Dir, LiveDist
    Dir = Flipper.StartAngle / Abs(Flipper.StartAngle)
    Dim LiveCatchBounce
    Dim CatchTime : CatchTime = GameTime - FCount
    LiveDist = Abs(Flipper.x - ball.x)
    If CatchTime <= LiveCatch And parm > 3 And LiveDist > LiveDistanceMin And LiveDist < LiveDistanceMax Then
        If CatchTime <= LiveCatch * 0.5 Then
            LiveCatchBounce = 0
        Else
            LiveCatchBounce = Abs((LiveCatch / 2) - CatchTime)
        End If
        If LiveCatchBounce = 0 And ball.velx * Dir > 0 And LiveDist > 30 Then ball.velx = 0
        If ball.velx * Dir > 0 And LiveDist < 30 Then
            ball.velx = BaseDampen * ball.velx
            ball.vely = BaseDampen * ball.vely
            ball.angmomx = BaseDampen * ball.angmomx
            ball.angmomy = BaseDampen * ball.angmomy
            ball.angmomz = BaseDampen * ball.angmomz
        ElseIf LiveDist > 30 Then
            ball.vely = LiveCatchBounce * (32 / LiveCatch)
            ball.angmomx = 0
            ball.angmomy = 0
            ball.angmomz = 0
        End If
    Else
        If Abs(Flipper.CurrentAngle) <= Abs(Flipper.EndAngle) + 1 Then FlippersD.Dampenf ActiveBall, parm
    End If
End Sub

Sub FlipperCradleCollision(ball1, ball2, velocity)
    If velocity < 0.7 Then Exit Sub
    Dim DoDamping, coef : DoDamping = False
    If LeftFlipper.CurrentAngle = LFEndAngle Then
        If FlipperTrigger(ball1.x, ball1.y, LeftFlipper) Or FlipperTrigger(ball2.x, ball2.y, LeftFlipper) Then DoDamping = True
    End If
    If RightFlipper.CurrentAngle = RFEndAngle Then
        If FlipperTrigger(ball1.x, ball1.y, RightFlipper) Or FlipperTrigger(ball2.x, ball2.y, RightFlipper) Then DoDamping = True
    End If
    If DoDamping Then
        coef = FCCDamping
        ball1.velx = ball1.velx * coef : ball1.vely = ball1.vely * coef : ball1.velz = ball1.velz * coef
        ball2.velx = ball2.velx * coef : ball2.vely = ball2.vely * coef : ball2.velz = ball2.velz * coef
    End If
End Sub

'******************************************************
'   ZBOU: VPW TargetBouncer
'******************************************************
Const TargetBouncerEnabled = 1
Const TargetBouncerFactor = 0.9

Sub TargetBouncer(aBall, defvalue)
    Dim zMultiplier, vel, vratio
    If TargetBouncerEnabled = 1 And aBall.z < 30 Then
        vel = BallSpeed(aBall)
        If aBall.velx = 0 Then vratio = 1 Else vratio = aBall.vely / aBall.velx
        Select Case Int(Rnd * 6) + 1
            Case 1 : zMultiplier = 0.2  * defvalue
            Case 2 : zMultiplier = 0.25 * defvalue
            Case 3 : zMultiplier = 0.3  * defvalue
            Case 4 : zMultiplier = 0.4  * defvalue
            Case 5 : zMultiplier = 0.45 * defvalue
            Case 6 : zMultiplier = 0.5  * defvalue
        End Select
        aBall.velz = Abs(vel * zMultiplier * TargetBouncerFactor)
        aBall.velx = Sgn(aBall.velx) * Sqr(Abs((vel ^ 2 - aBall.velz ^ 2) / (1 + vratio ^ 2)))
        aBall.vely = aBall.velx * vratio
    End If
End Sub

Sub TargetBounce_Hit(idx)
    TargetBouncer ActiveBall, 1
End Sub

'******************************************************
'  RUBBER DAMPENERS
'******************************************************

Sub dPosts_Hit(idx)
    RubbersD.Dampen ActiveBall
    TargetBouncer ActiveBall, 1
    Dim finalspeed
    finalspeed = Sqr(ActiveBall.velx * ActiveBall.velx + ActiveBall.vely * ActiveBall.vely)
    If finalspeed > 5 Then
        RandomSoundRubberStrong 1
    Else
        RandomSoundRubberWeak
    End If
End Sub

Sub dSleeves_Hit(idx)
    SleevesD.Dampen ActiveBall
    TargetBouncer ActiveBall, 0.7
End Sub

'Sub Wall001_Hit()
'    RubbersD.Dampen ActiveBall
'    Dim finalspeed
'    finalspeed = Sqr(ActiveBall.velx * ActiveBall.velx + ActiveBall.vely * ActiveBall.vely)
'    If finalspeed > 5 Then
'        RandomSoundRubberStrong 1
'    Else
'        RandomSoundRubberWeak
'    End If
'End Sub

Dim RubbersD : Set RubbersD = New Dampener
RubbersD.name = "Rubbers"
RubbersD.debugOn = False
RubbersD.Print = False
RubbersD.addpoint 0, 0, 1.1
RubbersD.addpoint 1, 3.77, 0.97
RubbersD.addpoint 2, 5.76, 0.967
RubbersD.addpoint 3, 15.84, 0.874
RubbersD.addpoint 4, 56, 0.64

Dim SleevesD : Set SleevesD = New Dampener
SleevesD.name = "Sleeves"
SleevesD.debugOn = False
SleevesD.Print = False
SleevesD.CopyCoef RubbersD, 0.85

Dim FlippersD : Set FlippersD = New Dampener
FlippersD.name = "Flippers"
FlippersD.debugOn = False
FlippersD.Print = False
FlippersD.addpoint 0, 0, 1.1
FlippersD.addpoint 1, 3.77, 0.99
FlippersD.addpoint 2, 6, 0.99

Class Dampener
    Public Print, debugOn
    Public name, Threshold
    Public ModIn, ModOut
    Private Sub Class_Initialize
        ReDim ModIn(0) : ReDim Modout(0)
    End Sub
    Public Sub AddPoint(aIdx, aX, aY)
        ShuffleArrays ModIn, ModOut, 1
        ModIn(aIDX) = aX : ModOut(aIDX) = aY
        ShuffleArrays ModIn, ModOut, 0
        If GameTime > 100 Then Report
    End Sub
    Public Sub Dampen(aBall)
        If Threshold Then If BallSpeed(aBall) < Threshold Then Exit Sub End If End If
        Dim RealCOR, DesiredCOR, coef
        DesiredCor = LinearEnvelope(cor.ballvel(aBall.id), ModIn, ModOut)
        RealCOR = BallSpeed(aBall) / (cor.ballvel(aBall.id) + 0.0001)
        coef = DesiredCor / RealCOR
        aBall.velx = aBall.velx * coef
        aBall.vely = aBall.vely * coef
        aBall.velz = aBall.velz * coef
    End Sub
    Public Sub Dampenf(aBall, parm)
        Dim RealCOR, DesiredCOR, coef
        DesiredCor = LinearEnvelope(cor.ballvel(aBall.id), ModIn, ModOut)
        RealCOR = BallSpeed(aBall) / (cor.ballvel(aBall.id) + 0.0001)
        coef = DesiredCor / RealCOR
        If Abs(aBall.velx) < 2 And aBall.vely < 0 And aBall.vely > -3.75 Then
            aBall.velx = aBall.velx * coef
            aBall.vely = aBall.vely * coef
            aBall.velz = aBall.velz * coef
        End If
    End Sub
    Public Sub CopyCoef(aObj, aCoef)
        Dim x
        For x = 0 To UBound(aObj.ModIn)
            AddPoint x, aObj.ModIn(x), aObj.ModOut(x) * aCoef
        Next
    End Sub
    Public Sub Report()
        If Not debugOn Then Exit Sub
    End Sub
End Class

'******************************************************
'  TRACK ALL BALL VELOCITIES
'******************************************************

Dim cor : Set cor = New CoRTracker

Class CoRTracker
    Public ballvel, ballvelx, ballvely
    Private Sub Class_Initialize
        ReDim ballvel(0) : ReDim ballvelx(0) : ReDim ballvely(0)
    End Sub
    Public Sub Update()
        Dim b, AllBalls, highestID
        allBalls = GetBalls
        For Each b In allBalls
            If b.id >= highestID Then highestID = b.id
        Next
        If UBound(ballvel) < highestID Then ReDim ballvel(highestID)
        If UBound(ballvelx) < highestID Then ReDim ballvelx(highestID)
        If UBound(ballvely) < highestID Then ReDim ballvely(highestID)
        For Each b In allBalls
            ballvel(b.id) = BallSpeed(b)
            ballvelx(b.id) = b.velx
            ballvely(b.id) = b.vely
        Next
    End Sub
End Class

Sub RDampen_Timer()
    cor.Update
End Sub

'' FLEEP SOUNDS
'
'Sub SoundNudgeLeft()
'    PlaySound "Nudge_1", 0, NudgeLeftSoundLevel * VolumeDial, -0.1, 0.25
'End Sub
'
'Sub SoundNudgeRight()
'    PlaySound "Nudge_2", 0, NudgeRightSoundLevel * VolumeDial, 0.1, 0.25
'End Sub
'
'Sub SoundNudgeCenter()
'    PlaySound "Nudge_3", 0, NudgeCenterSoundLevel * VolumeDial, 0, 0.25
'End Sub
'
'Sub SoundFlipperUpAttackLeft()
'    FlipperUpAttackLeftSoundLevel = RndNum(FlipperUpAttackMinimumSoundLevel, FlipperUpAttackMaximumSoundLevel)
'    PlaySoundAtLevelStatic "Flipper_Attack-L01", FlipperUpAttackLeftSoundLevel, LeftFlipper
'End Sub
'
'Sub SoundFlipperUpAttackRight()
'    FlipperUpAttackRightSoundLevel = RndNum(FlipperUpAttackMinimumSoundLevel, FlipperUpAttackMaximumSoundLevel)
'    PlaySoundAtLevelStatic "Flipper_Attack-R01", FlipperUpAttackRightSoundLevel, RightFlipper
'End Sub
'
'Sub RandomSoundFlipperUpLeft()
'    PlaySoundAtLevelStatic "Flipper_L0" & Int(Rnd * 9) + 1, FlipperLeftHitParm, LeftFlipper
'End Sub
'
'Sub RandomSoundFlipperUpRight()
'    PlaySoundAtLevelStatic "Flipper_R0" & Int(Rnd * 9) + 1, FlipperRightHitParm, RightFlipper
'End Sub
'
'Sub RandomSoundReflipUpLeft()
'    PlaySoundAtLevelStatic "Flipper_ReFlip_L0" & Int(Rnd * 3) + 1, RndNum(0.8, 1) * FlipperUpSoundLevel, LeftFlipper
'End Sub
'
'Sub RandomSoundReflipUpRight()
'    PlaySoundAtLevelStatic "Flipper_ReFlip_R0" & Int(Rnd * 3) + 1, RndNum(0.8, 1) * FlipperUpSoundLevel, RightFlipper
'End Sub
'
'Sub RandomSoundFlipperDownLeft()
'    PlaySoundAtLevelStatic "Flipper_Left_Down_" & Int(Rnd * 7) + 1, FlipperDownSoundLevel, LeftFlipper
'End Sub
'
'Sub RandomSoundFlipperDownRight()
'    PlaySoundAtLevelStatic "Flipper_Right_Down_" & Int(Rnd * 8) + 1, FlipperDownSoundLevel, RightFlipper
'End Sub

    'Manual Ball Control
  If EnableBallControl = 1 Then
    If keycode = 203 Then BCleft = 0
    If keycode = 200 Then BCup = 0
    If keycode = 208 Then BCdown = 0
    If keycode = 205 Then BCright = 0
  End If

Sub Table1_Exit()
  On Error Resume Next
    If Not Controller Is Nothing Then Controller.Stop
    If FlexDMDActive Then
        FlexDMD.Show = False
        FlexDMD.Run = False
        FlexDMD.Run = False
        Set FlexDMD = Nothing
    End If
    On Error Goto 0
End Sub

Sub Plunger_Init()
    If TableStarted Then Exit Sub
    TableStarted = True
    If FreePlay Then
        ShowMessage "PRESS START  FREE PLAY"
    ElseIf Credits > 0 Then
        ShowMessage "PRESS START  CREDITS:" & Credits
    Else
        ShowMessage "INSERT COIN TO PLAY"
    End If
End Sub

Sub ResetAllBumperLightsOnDrain()
    ChillCheckTimer.Enabled = False
    PoisonTimer.Enabled = False
    PoisonTickActive = False
    FireBurnTimer.Enabled = False
    FireBurnSlot = -1
    Dim brd
    For brd = 0 To 4
        BumperChilled(brd)  = False
        BumperPoisoned(brd) = False
        BumperBurning(brd)  = False
        FlBumperColor(brd + 1) = "red"
        FlInitBumper brd + 1, "red"
        FlBumperFadeTarget(brd + 1) = 0
        FlBumperFadeActual(brd + 1) = 0
        FlFadeBumper brd + 1, 0
    Next
End Sub

Sub Drain_Hit()
    LF.PolarityCorrect ActiveBall
    RF.PolarityCorrect ActiveBall
    ULF.PolarityCorrect ActiveBall
    Dim drainingID : drainingID = ActiveBall.ID
    WRemoveBall drainingID
    Dim dbt
'    Dim wasBonusBall : wasBonusBall = (ActiveElementType >= 0 And GetBIP() = 0)
'    Dim bonusBallType : bonusBallType = ActiveElementType

' If CurrentAct = 6 And GetBIP() <= 1 Then
'        If PortalCollapsed Then Exit Sub
'        If Not PortalActive Then Exit Sub
'        Drain.DestroyBall
'        townportalPF.Visible = True
'        PlayfieldKicker.Enabled = True
'        PlayfieldKicker.CreateBall
'        If Rnd > 0.5 Then PlayfieldKicker.Kick 165, 15 Else PlayfieldKicker.Kick 195, 15
'        PlaySound "popper_ball", 0, 1
'        PlaySound "portalenter", 0, 1
'        ShowMessage "RESPAWNING..."
'   townportalPF.Visible = False
'        PlayfieldKicker.Enabled = False
'        Exit Sub
'    End If
' --- Inside Sub Drain_Hit ---

'If CurrentAct = 6 Then
'    If PortalCollapsed Then
'        ' PHASE A: The trial is over, and the ball just hit the drain.
'        ' We start the timer NOW to give a 1-second "breath" before the bonus tally.
'        PortalCollapseTimer.Interval = 1000
'        PortalCollapseTimer.Enabled = True
'   Drain.DestroyBall
'        Exit Sub
'
'    ElseIf PortalActive And GetBIP() <= 1 Then
'        ' PHASE B: The trial is still active. Respawn the ball.
'        Drain.DestroyBall
'        townportalPF.Visible = True
'        PlayfieldKicker.Enabled = True
'
'        Dim eb : Set eb = PlayfieldKicker.CreateBall
'        Select Case ActiveElementType
'            Case ATTACK_FIRE      : eb.Color = RGB(255, 60, 0)
'            Case ATTACK_COLD      : eb.Color = RGB(80, 180, 255)
'            Case ATTACK_POISON    : eb.Color = RGB(0, 200, 50)
'            Case ATTACK_LIGHTNING : eb.Color = RGB(196, 196, 8)
'        End Select
'
'        If Rnd > 0.5 Then PlayfieldKicker.Kick 165, 15 Else PlayfieldKicker.Kick 195, 15
'        PlaySound "popper_ball", 0, 1
'        PlaySound "portalenter", 0, 1
'        ShowMessage "RESPAWNING..."
'        Exit Sub
'    End If
'End If

'If CurrentAct = 6 Then
'        If PortalCollapsed Then
'            Drain.DestroyBall
'            PortalCollapseTimer.Interval = 1000
'            PortalCollapseTimer.Enabled = True
'            Exit Sub
''        ElseIf PortalActive And GetBIP() <= 1 Then
''            Drain.DestroyBall
''            townportalPF.Visible = True
''            PlayfieldKicker.Enabled = True
''            PlayfieldKicker.CreateBall
''      SetAllBallsElement ActiveElementType
''            If Rnd > 0.5 Then PlayfieldKicker.Kick 165, 15 Else PlayfieldKicker.Kick 195, 15
''            PlaySound "popper_ball", 0, 1
''            PlaySound "portalenter", 0, 1
''            ShowMessage "RESPAWNING..."
''            Exit Sub
'        End If
'    End If

    Dim inMultiball : inMultiball = GetBIP() > 1
' If BallSaveActive And Not wasBonusBall And (Not BallSaveUsed Or inMultiball) Then
    If (BallSaveActive And (Not BallSaveUsed Or inMultiball)) Or MercSaveActive Then
        If Not inMultiball Then BallSaveUsed = True
        BallSaveActive = Not inMultiball
        BallSaveTimer.Enabled = False
        BallSaveL.State = 0
        BallSaveL2.State = 0
        Drain.DestroyBall
        If MercSaveActive Then
            townportalPF.Visible = True
            PlayfieldKicker.CreateBall
            SetAllBallsElement ActiveElementType
            If Rnd > 0.5 Then PlayfieldKicker.Kick 165, 15 Else PlayfieldKicker.Kick 195, 15
            PlayfieldKicker.Enabled = False
            PlaySound "redemption", 0, 1
            PlaySound "portalenter", 0, 1
            ShowMessage "MERCENARY SAVED!"
            SetFlasherColor 5, 255, 220, 80 : SetFlasherColor 6, 255, 220, 80
            FireFlasher 5 : FireFlasher 6
            ApronFlasherDimTimer.Enabled = False
            ApronFlasherDimTimer.Enabled = True
            MercSaveActive = False
            townportalPF.Visible = False
        ElseIf RuneWordMultiballRunning Then
            RuneKick.CreateBall
      SetAllBallsElement ActiveElementType
      ResumeAmbushTimer
      ResumeKillStreakTimer
            RuneKick.Kick 180, 30
            PlaySound "redemption", 0, 1
            ShowMessage "BALL SAVED!"
            SetFlasherColor 5, 255, 220, 80 : SetFlasherColor 6, 255, 220, 80
            FireFlasher 5 : FireFlasher 6
            ApronFlasherDimTimer.Enabled = False
            ApronFlasherDimTimer.Enabled = True
        ElseIf PartyMultiballRunning Then
            Dim savedPB : Set savedPB = PlayfieldKicker.CreateBall
            SetAllBallsElement ActiveElementType
            If Rnd > 0.5 Then PlayfieldKicker.Kick 165, 15 Else PlayfieldKicker.Kick 195, 15
            PlaySound "redemption", 0, 1
            PlaySound "portalenter", 0, 1
            ShowMessage "ALLY SAVED!"
            SetFlasherColor 5, 255, 220, 80 : SetFlasherColor 6, 255, 220, 80
            FireFlasher 5 : FireFlasher 6
            ApronFlasherDimTimer.Enabled = False
            ApronFlasherDimTimer.Enabled = True
        Else
            BallRelease.CreateBall
            SetAllBallsElement ActiveElementType
            BallRelease.Kick 90, 7
            PlaySoundAtLevelStatic SoundFX("BallRelease" & Int(Rnd * 7) + 1, DOFContactors), BallReleaseSoundLevel, BallRelease
            PlaySound "redemption", 0, 1
            ShowMessage "BALL SAVED!"
            SetFlasherColor 5, 255, 220, 80 : SetFlasherColor 6, 255, 220, 80
            FireFlasher 5 : FireFlasher 6
            ApronFlasherDimTimer.Enabled = False
            ApronFlasherDimTimer.Enabled = True
        End If
        Exit Sub
    End If

    Dim bipBeforeDrain : bipBeforeDrain = GetBIP()

    ' Award any active kill streak only when the final ball drains
    If bipBeforeDrain <= 1 And KillStreakTimer.Enabled Then
        KillStreakTimer.Enabled = False
        AwardKillStreak
    End If

    Drain.DestroyBall
    PlaySoundAtLevelStatic "Drain_" & Int(Rnd * 11) + 1, DrainSoundLevel, Drain

    If AmbushActive Then
        If bipBeforeDrain <= 1 Then
            AmbushHurryTimer.Enabled = False
            AmbushResumeTimer.Enabled = False
            AmbushLightTimer.Enabled = False
            AmbushHoldTimer.Enabled = False
            Select Case AmbushTriggerNum
                Case 1 : QKick1.Kick 180, 15 : QKick1.Enabled = False
                Case 2 : QKick2.Kick 180, 15 : QKick2.Enabled = False
                Case 4 : QKick4.Kick 180, 15 : QKick4.Enabled = False
            End Select
            AmbushActive = False
      SetRampLight
            AmbushHappenedThisAct = False
            AmbushKillCount = 0
            AmbushTriggerNum = 0
            TravelActive = True
        End If
    End If
    If InstantKillActive Then
        InstantKillActive = False
        InstantKillTimer.Enabled = False
        InstantKillTimer_step = 0
        AuraKillScore = 0
        Dim ikbd
        For ikbd = 1 To 5
            FlBumperColor(ikbd) = "red"
            FlInitBumper ikbd, "red"
            FlBumperFadeTarget(ikbd) = 0
            FlBumperFadeActual(ikbd) = 0
            FlFadeBumper ikbd, 0
        Next
    End If
  If LeapBallID < 0 Then
        LeapReady = False
        BarbLeanTimer.Enabled = False
        BarbJumpTimer.Enabled = False
        BarbJumpStep = 0
        LeapSafetyTimer.Enabled = False
        LeapSafetyTimer_step = 0
        LeapPromptTimer.Enabled = False
        StopSound "TOM_Trunk_Motor_Long"
        BarbPrim.RotX = BarbRestRotX
        BarbPrim.TransZ = BarbRestZ
    End If

'    If wasBonusBall Then
'        If bonusBallType = ATTACK_POISON Then
'            PoisonTickActive = False
'            PoisonTimer.Enabled = False
'            Dim bps : For bps = 0 To 4 : BumperPoisoned(bps) = False : Next
'        End If
'        If GetBIP() <= 0 Then
'            CubeHoldTimer.Enabled = False
'            CubeSpawnTimer.Enabled = False
'            If SilverBallID >= 0 Then
'                CubeKick.DestroyBall
'                SilverBallID = -1
'            End If
'            PendingElementType = -1
'            CloseCubeWall
'        End If
'         If GetBIP() <= 0 And Not CubeHoldTimer.Enabled And Not CubeSpawnTimer.Enabled And SilverBallID < 0 Then
'            ResetAllBumperLightsOnDrain
'            ActiveElementType = -1
'            ElementBallActive = False
'            Shield.Collidable = False
'            Shield.Visible = False
'            ShieldActive = False
'            Select Case Int(Rnd * 2)
'                Case 0 : PlaySound "bar_death1", 0, 1
'                Case 1 : PlaySound "bar_death2", 0, 1
'            End Select
'            GIEventMode = GI_MODE_NONE
'            GIEventTimer.Enabled = False
'            StartGIEvent GI_MODE_DRAIN
'            SetFlasherColor 1, 180, 0, 0 : SetFlasherColor 2, 180, 0, 0
'     SetFlasherColor 3, 180, 0, 0 : SetFlasherColor 4, 180, 0, 0
'     SetFlasherColor 5, 180, 0, 0 : SetFlasherColor 6, 180, 0, 0
'     FireAllFlashers
'            ApronFlasherDimTimer.Enabled = False
'            ApronFlasherDimTimer.Enabled = True
''            SaveBestGearOnDrain
''            ClearLootTargets
'            If CurrentAct = 5 Then
'                Select Case Int(Rnd * 2)
'                    Case 0 : PlaySound "baallaugh1", 0, 1 * DuckVolume
'                    Case 1 : PlaySound "baallaugh2", 0, 1 * DuckVolume
'                End Select
'            End If
'            StartEndOfBallBonus
'            Exit Sub
'        End If
'        If GetBIP() = 1 And RuneWordMultiballRunning Then
'            RuneWordMultiballRunning = False
'            GIEventMode = GI_MODE_NONE
'            GIEventTimer.Enabled = False
'            RestoreAllGIRows
'        End If
'        If Not RuneWordMultiballRunning And Not PartyMultiballRunning Then
'            Select Case bonusBallType
'                Case ATTACK_FIRE      : CubeLight.Color = RGB(255, 60, 0)
'                Case ATTACK_COLD      : CubeLight.Color = RGB(80, 180, 255)
'                Case ATTACK_POISON    : CubeLight.Color = RGB(0, 200, 50)
'                Case ATTACK_LIGHTNING : CubeLight.Color = RGB(196, 196, 8)
'            End Select
'            CubeLight.BlinkInterval = 150
'            CubeLight.BlinkPattern = "10"
'            CubeLight.State = 2
'            PlaySound "redemption", 0, 1
'            ShowMessage "ELEMENT RETURNS!"
'            CubeRespawnHoldTimer.Enabled = True
'        End If
'        Exit Sub
'    End If

    If GetBIP() <= 0 And Not CubeTransmuteActive Then
        ResetAllBumperLightsOnDrain
        If RuneWordMultiballRunning Then
            RuneWordMultiballRunning = False
            RuneQuality = 0
            ResetRuneTargets
        End If
        PortalOpenTimer.Enabled = False : PortalFlashTimer.Enabled = False : PartySpawnTimer.Enabled = False : MercSpawnTimer.Enabled = False : MercMultiballActive = False
        Shield.Collidable = False
        Shield.Visible = False
        ShieldActive = False
        Select Case Int(Rnd * 2)
            Case 0 : PlaySound "bar_death1", 0, 1
            Case 1 : PlaySound "bar_death2", 0, 1
        End Select
        GIEventMode = GI_MODE_NONE
        GIEventTimer.Enabled = False
        StartGIEvent GI_MODE_DRAIN
        SetFlasherColor 1, 180, 0, 0 : SetFlasherColor 2, 180, 0, 0
    SetFlasherColor 3, 180, 0, 0 : SetFlasherColor 4, 180, 0, 0
    SetFlasherColor 5, 180, 0, 0 : SetFlasherColor 6, 180, 0, 0
    FireAllFlashers
        ApronFlasherDimTimer.Enabled = False
        ApronFlasherDimTimer.Enabled = True
'        SaveBestGearOnDrainGameOverMusicTimer.Enabled = False
'            ClearLootTargets
            If CurrentAct = 5 Then
                Select Case Int(Rnd * 2)
                    Case 0 : PlaySound "baallaugh1", 0, 1 * DuckVolume
                    Case 1 : PlaySound "baallaugh2", 0, 1 * DuckVolume
                End Select
            End If
            StartEndOfBallBonus
            Exit Sub
    End If

    If GetBIP() = 1 Then
    If MercMultiballActive Then
            MercMultiballActive = False
            MercSpawnTimer.Enabled = False
            If Not MercSaveActive Then
                townportalPF.Visible = False
                PlayfieldKicker.Enabled = False
            End If
        End If
        If RuneWordMultiballRunning Then
            RuneWordMultiballRunning = False
        End If
        If PartyMultiballRunning Then
            PartyMultiballRunning = False
            PartySpawnTimer.Enabled = False
            PortalOpenTimer.Enabled = False : PortalFlashTimer.Enabled = False
            townportalPF.Visible = False
            PlayfieldKicker.Enabled = False
        End If
        GIEventMode = GI_MODE_NONE
        GIEventTimer.Enabled = False
        RestoreAllGIRows
    End If
End Sub

Function BallSaveBonus()
    Dim extra : extra = GetPartyCount() - 1
    If extra < 0 Then extra = 0
    BallSaveBonus = extra * 1000
End Function

Sub BallSaveTimer_Timer()
    BallSaveTimer.Enabled = False
    BallSaveActive = False
    BallSaveUsed = True
    BallSaveL.State = 0
    BallSaveL2.State = 0
    If MercSaveActive Then
        MercSaveActive = False
        townportalPF.Visible = False
    End If
End Sub

' HIGH SCORE

Dim NightmareChamp     : NightmareChamp     = 0
Dim NightmareChampName : NightmareChampName = "D2X"
Dim HellChamp          : HellChamp          = 0
Dim HellChampName      : HellChampName      = "D2X"

Dim HiKills(4)
Dim HiKillsName(4)
Dim KillsEnterNameSpot : KillsEnterNameSpot = 0
Dim hsbKillsModeActive : hsbKillsModeActive = False
Dim hsKillsCurrentDigit : hsKillsCurrentDigit = 0
Dim hsKillsCurrentLetter : hsKillsCurrentLetter = 1
Dim hsKillsEnteredDigits(2)
Dim HiScore(4)
Dim HiName(4)
Dim hsbModeActive : hsbModeActive = False
Dim LastEnteredInitials : LastEnteredInitials = ""
Dim HSAutoCommit      : HSAutoCommit      = False
Dim KillsHSAutoCommit : KillsHSAutoCommit = False
Dim hsCurrentDigit : hsCurrentDigit = 0
Dim hsCurrentLetter : hsCurrentLetter = 1
Dim hsEnteredDigits(2)
Dim hsValidLetters : hsValidLetters = "ABCDEFGHIJKLMNOPQRSTUVWXYZ<"
Dim EnterNameSpot : EnterNameSpot = 0

Sub LoadHighScores()
    Dim x
    x = LoadValue(TableName, "HiScore1")
    If x <> "" Then HiScore(0) = CDbl(x) Else HiScore(0) = 100000000 End If
    x = LoadValue(TableName, "HiName1")
    If x <> "" Then HiName(0) = x Else HiName(0) = "D2X" End If

    x = LoadValue(TableName, "HiScore2")
    If x <> "" Then HiScore(1) = CDbl(x) Else HiScore(1) = 75000000 End If
    x = LoadValue(TableName, "HiName2")
    If x <> "" Then HiName(1) = x Else HiName(1) = "D2X" End If

    x = LoadValue(TableName, "HiScore3")
    If x <> "" Then HiScore(2) = CDbl(x) Else HiScore(2) = 50000000 End If
    x = LoadValue(TableName, "HiName3")
    If x <> "" Then HiName(2) = x Else HiName(2) = "D2X" End If

    x = LoadValue(TableName, "HiScore4")
    If x <> "" Then HiScore(3) = CDbl(x) Else HiScore(3) = 25000000 End If
    x = LoadValue(TableName, "HiName4")
    If x <> "" Then HiName(3) = x Else HiName(3) = "D2X" End If

    x = LoadValue(TableName, "HiScore5")
    If x <> "" Then HiScore(4) = CDbl(x) Else HiScore(4) = 15000000 End If
    x = LoadValue(TableName, "HiName5")
    If x <> "" Then HiName(4) = x Else HiName(4) = "D2X" End If

  x = LoadValue(TableName, "HiKills1")
    If x <> "" Then HiKills(0) = CInt(x) Else HiKills(0) = 5 End If
    x = LoadValue(TableName, "HiKillsName1")
    If x <> "" Then HiKillsName(0) = x Else HiKillsName(0) = "D2X" End If
    x = LoadValue(TableName, "HiKills2")
    If x <> "" Then HiKills(1) = CInt(x) Else HiKills(1) = 4 End If
    x = LoadValue(TableName, "HiKillsName2")
    If x <> "" Then HiKillsName(1) = x Else HiKillsName(1) = "D2X" End If
    x = LoadValue(TableName, "HiKills3")
    If x <> "" Then HiKills(2) = CInt(x) Else HiKills(2) = 3 End If
    x = LoadValue(TableName, "HiKillsName3")
    If x <> "" Then HiKillsName(2) = x Else HiKillsName(2) = "D2X" End If
    x = LoadValue(TableName, "HiKills4")
    If x <> "" Then HiKills(3) = CInt(x) Else HiKills(3) = 2 End If
    x = LoadValue(TableName, "HiKillsName4")
    If x <> "" Then HiKillsName(3) = x Else HiKillsName(3) = "D2X" End If
    x = LoadValue(TableName, "HiKills5")
    If x <> "" Then HiKills(4) = CInt(x) Else HiKills(4) = 1 End If
    x = LoadValue(TableName, "HiKillsName5")
    If x <> "" Then HiKillsName(4) = x Else HiKillsName(4) = "D2X" End If
  x = LoadValue(TableName, "NightmareChamp")
  If x <> "" Then NightmareChamp = CDbl(x) Else NightmareChamp = 0 End If
  x = LoadValue(TableName, "NightmareChampName")
  If x <> "" Then NightmareChampName = x Else NightmareChampName = "D2X" End If
  x = LoadValue(TableName, "HellChamp")
  If x <> "" Then HellChamp = CDbl(x) Else HellChamp = 0 End If
  x = LoadValue(TableName, "HellChampName")
  If x <> "" Then HellChampName = x Else HellChampName = "D2X" End If
End Sub

Sub SaveHighScores()
    SaveValue TableName, "HiScore1", HiScore(0)
    SaveValue TableName, "HiName1",  HiName(0)
    SaveValue TableName, "HiScore2", HiScore(1)
    SaveValue TableName, "HiName2",  HiName(1)
    SaveValue TableName, "HiScore3", HiScore(2)
    SaveValue TableName, "HiName3",  HiName(2)
    SaveValue TableName, "HiScore4", HiScore(3)
    SaveValue TableName, "HiName4",  HiName(3)
    SaveValue TableName, "HiScore5", HiScore(4)
    SaveValue TableName, "HiName5",  HiName(4)
  SaveValue TableName, "HiKills1", HiKills(0)
    SaveValue TableName, "HiKillsName1", HiKillsName(0)
    SaveValue TableName, "HiKills2", HiKills(1)
    SaveValue TableName, "HiKillsName2", HiKillsName(1)
    SaveValue TableName, "HiKills3", HiKills(2)
    SaveValue TableName, "HiKillsName3", HiKillsName(2)
    SaveValue TableName, "HiKills4", HiKills(3)
    SaveValue TableName, "HiKillsName4", HiKillsName(3)
    SaveValue TableName, "HiKills5", HiKills(4)
    SaveValue TableName, "HiKillsName5", HiKillsName(4)
  SaveValue TableName, "NightmareChamp",     NightmareChamp
  SaveValue TableName, "NightmareChampName", NightmareChampName
  SaveValue TableName, "HellChamp",          HellChamp
  SaveValue TableName, "HellChampName",      HellChampName
End Sub

Dim FinalDifficulty : FinalDifficulty = 0
Sub CheckDifficultyHighScore()
    If FinalDifficulty = 1 And Score > NightmareChamp Then
        NightmareChamp = Score
        NightmareChampName = "???"
        DifficultyHSEntryInit 1
    ElseIf FinalDifficulty = 2 And Score > HellChamp Then
        HellChamp = Score
        HellChampName = "???"
        DifficultyHSEntryInit 2
    Else
        CheckHighScore  ' High Score is the headline now; Kill auto-saves after
    End If
End Sub

Dim DifficultyHSMode   : DifficultyHSMode   = 0
Dim DifficultyHSDigit  : DifficultyHSDigit  = 0
Dim DifficultyHSLetter : DifficultyHSLetter = 1
Dim DifficultyHSLevel  : DifficultyHSLevel  = 0
Dim DifficultyHSActive : DifficultyHSActive = False

Sub DifficultyHSDisplayName()
    Dim entered : entered = ""
    Dim i
    For i = 0 To 2
        If i < DifficultyHSDigit Then
            entered = entered & hsEnteredDigits(i)
        ElseIf i = DifficultyHSDigit Then
            entered = entered & Mid(hsValidLetters, DifficultyHSLetter, 1)
        Else
            entered = entered & "-"
        End If
    Next
    If DifficultyHSLevel = 1 Then
        UpdateDMD2 "NIGHTMARE CHAMPION", "INITIALS: " & entered
    Else
        UpdateDMD2 "HELL CHAMPION", "INITIALS: " & entered
    End If
End Sub

Sub DifficultyHSEntryInit(level)
    DifficultyHSActive = True
    DifficultyHSLevel = level
    DifficultyHSDigit = 0
    DifficultyHSLetter = 1
    hsEnteredDigits(0) = "A"
    hsEnteredDigits(1) = "-"
    hsEnteredDigits(2) = "-"
    If level = 1 Then
        UpdateDMD2 "NIGHTMARE CHAMPION!", "ENTER INITIALS"
    Else
        UpdateDMD2 "HELL CHAMPION!", "ENTER INITIALS"
    End If
    HSBSplashTimer.Interval = 2000
    HSBSplashTimer.Enabled = True
End Sub

Sub DifficultyHSCommit()
    DifficultyHSActive = False
    Dim entered : entered = hsEnteredDigits(0) & hsEnteredDigits(1) & hsEnteredDigits(2)
    If entered = "---" Then entered = "D2X"
    If DifficultyHSLevel = 1 Then
        NightmareChampName = entered
    Else
        HellChampName = entered
    End If
    SaveHighScores
    LastEnteredInitials = entered
    Score = 0
    ShowMessage "SAVED: " & entered
    CheckHighScore  ' chain to High Score; Kill auto-saves last
End Sub

Sub CheckKillHighScore()
    Dim tmp : tmp = EnemiesKilled
    KillsEnterNameSpot = 0
    If tmp > HiKills(0) Then
        KillsEnterNameSpot = 1
        HiKills(4) = HiKills(3) : HiKillsName(4) = HiKillsName(3)
        HiKills(3) = HiKills(2) : HiKillsName(3) = HiKillsName(2)
        HiKills(2) = HiKills(1) : HiKillsName(2) = HiKillsName(1)
        HiKills(1) = HiKills(0) : HiKillsName(1) = HiKillsName(0)
        HiKills(0) = tmp
    ElseIf tmp > HiKills(1) Then
        KillsEnterNameSpot = 2
        HiKills(4) = HiKills(3) : HiKillsName(4) = HiKillsName(3)
        HiKills(3) = HiKills(2) : HiKillsName(3) = HiKillsName(2)
        HiKills(2) = HiKills(1) : HiKillsName(2) = HiKillsName(1)
        HiKills(1) = tmp
    ElseIf tmp > HiKills(2) Then
        KillsEnterNameSpot = 3
        HiKills(4) = HiKills(3) : HiKillsName(4) = HiKillsName(3)
        HiKills(3) = HiKills(2) : HiKillsName(3) = HiKillsName(2)
        HiKills(2) = tmp
    ElseIf tmp > HiKills(3) Then
        KillsEnterNameSpot = 4
        HiKills(4) = HiKills(3) : HiKillsName(4) = HiKillsName(3)
        HiKills(3) = tmp
    ElseIf tmp > HiKills(4) Then
        KillsEnterNameSpot = 5
        HiKills(4) = tmp
    End If
   If KillsEnterNameSpot > 0 Then
        KillsHighScoreEntryInit
    Else
        StartAttractMode   ' Kill is the final step now
    End If
End Sub

Sub KillsHighScoreEntryInit()
    If LastEnteredInitials <> "" Then
        hsKillsEnteredDigits(0) = Mid(LastEnteredInitials, 1, 1)
        hsKillsEnteredDigits(1) = Mid(LastEnteredInitials, 2, 1)
        hsKillsEnteredDigits(2) = Mid(LastEnteredInitials, 3, 1)
        LastEnteredInitials = ""
        KillsHSAutoCommit = True
        UpdateDMD2 "KILL REC! " & EnemiesKilled & " KILLS", "SAVING: " & hsKillsEnteredDigits(0) & hsKillsEnteredDigits(1) & hsKillsEnteredDigits(2)
        KillsHSBSplashTimer.Interval = 2000
        KillsHSBSplashTimer.Enabled = True
        Exit Sub
    End If
    hsbKillsModeActive = True
    hsKillsCurrentDigit = 0
    hsKillsCurrentLetter = 1
    hsKillsEnteredDigits(0) = "A"
    hsKillsEnteredDigits(1) = "-"
    hsKillsEnteredDigits(2) = "-"
    UpdateDMD2 "KILL RECORD!", "ENTER YOUR INITIALS"
    KillsHSBSplashTimer.Interval = 2000
    KillsHSBSplashTimer.Enabled = True
End Sub

Sub KillsHSBSplashTimer_Timer()
    KillsHSBSplashTimer.Enabled = False
    If KillsHSAutoCommit Then
        KillsHSAutoCommit = False
        KillsHighScoreCommitName
    Else
        KillsHighScoreDisplayName
    End If
End Sub

Sub KillsHighScoreDisplayName()
    Dim entered : entered = ""
    Dim i
    For i = 0 To 2
        If i < hsKillsCurrentDigit Then
            entered = entered & hsKillsEnteredDigits(i)
        ElseIf i = hsKillsCurrentDigit Then
            entered = entered & Mid(hsValidLetters, hsKillsCurrentLetter, 1)
        Else
            entered = entered & "-"
        End If
    Next
    UpdateDMD2 "KILL RECORD #" & KillsEnterNameSpot, "INITIALS: " & entered
End Sub

Sub EnterKillsHighScoreKey(keycode)
    If keycode = LeftFlipperKey Then
        hsKillsCurrentLetter = hsKillsCurrentLetter - 1
        If hsKillsCurrentLetter = 0 Then hsKillsCurrentLetter = Len(hsValidLetters)
        KillsHighScoreDisplayName
    End If
    If keycode = RightFlipperKey Then
        hsKillsCurrentLetter = hsKillsCurrentLetter + 1
        If hsKillsCurrentLetter > Len(hsValidLetters) Then hsKillsCurrentLetter = 1
        KillsHighScoreDisplayName
    End If
    If keycode = 13 Or keycode = StartGameKey Then
        Dim letter : letter = Mid(hsValidLetters, hsKillsCurrentLetter, 1)
        If letter = "<" Then
            If hsKillsCurrentDigit > 0 Then
                hsKillsCurrentDigit = hsKillsCurrentDigit - 1
                hsKillsEnteredDigits(hsKillsCurrentDigit) = "-"
            End If
            KillsHighScoreDisplayName
        Else
            hsKillsEnteredDigits(hsKillsCurrentDigit) = letter
            hsKillsCurrentDigit = hsKillsCurrentDigit + 1
            If hsKillsCurrentDigit >= 3 Then
                KillsHighScoreCommitName
            Else
                KillsHighScoreDisplayName
            End If
        End If
    End If
End Sub

Sub KillsHighScoreCommitName()
    hsbKillsModeActive = False
    Dim entered : entered = hsKillsEnteredDigits(0) & hsKillsEnteredDigits(1) & hsKillsEnteredDigits(2)
    If entered = "---" Then entered = "D2X"
    HiKillsName(KillsEnterNameSpot - 1) = entered
    SaveHighScores
    ShowMessage "KILLS SAVED: " & entered
    LastEnteredInitials = entered
    StartAttractMode   ' Kill is the final step now
End Sub

Sub CheckHighScore()
    If FinalDifficulty > 0 Then
        Score = 0
        CheckKillHighScore   ' difficulty game: champ covered High; Kill still saves
        Exit Sub
    End If
    Dim tmp : tmp = Score
    EnterNameSpot = 0

    If tmp > HiScore(0) Then
        EnterNameSpot = 1
        HiScore(4) = HiScore(3) : HiName(4) = HiName(3)
        HiScore(3) = HiScore(2) : HiName(3) = HiName(2)
        HiScore(2) = HiScore(1) : HiName(2) = HiName(1)
        HiScore(1) = HiScore(0) : HiName(1) = HiName(0)
        HiScore(0) = tmp
    ElseIf tmp > HiScore(1) Then
        EnterNameSpot = 2
        HiScore(4) = HiScore(3) : HiName(4) = HiName(3)
        HiScore(3) = HiScore(2) : HiName(3) = HiName(2)
        HiScore(2) = HiScore(1) : HiName(2) = HiName(1)
        HiScore(1) = tmp
    ElseIf tmp > HiScore(2) Then
        EnterNameSpot = 3
        HiScore(4) = HiScore(3) : HiName(4) = HiName(3)
        HiScore(3) = HiScore(2) : HiName(3) = HiName(2)
        HiScore(2) = tmp
    ElseIf tmp > HiScore(3) Then
        EnterNameSpot = 4
        HiScore(4) = HiScore(3) : HiName(4) = HiName(3)
        HiScore(3) = tmp
    ElseIf tmp > HiScore(4) Then
        EnterNameSpot = 5
        HiScore(4) = tmp
    End If

    If EnterNameSpot > 0 Then
        HighScoreEntryInit
    Else
    Score = 0
        CheckKillHighScore   ' no High record — go to Kill (final step)
    End If
End Sub

Sub HighScoreEntryInit()
    If LastEnteredInitials <> "" Then
        hsEnteredDigits(0) = Mid(LastEnteredInitials, 1, 1)
        hsEnteredDigits(1) = Mid(LastEnteredInitials, 2, 1)
        hsEnteredDigits(2) = Mid(LastEnteredInitials, 3, 1)
        LastEnteredInitials = ""
        HSAutoCommit = True
        UpdateDMD2 "HIGH SCORE " & FormatNumber(HiScore(EnterNameSpot - 1), 0, -1, 0, -1), "SAVING: " & hsEnteredDigits(0) & hsEnteredDigits(1) & hsEnteredDigits(2)
        HSBSplashTimer.Interval = 2000
        HSBSplashTimer.Enabled = True
        Exit Sub
    End If
    hsbModeActive = True
    hsCurrentDigit = 0
    hsCurrentLetter = 1
    hsEnteredDigits(0) = "A"
    hsEnteredDigits(1) = "-"
    hsEnteredDigits(2) = "-"
    UpdateDMD2 "HIGH SCORE!", "ENTER YOUR INITIALS"
    HSBSplashTimer.Interval = 2000
    HSBSplashTimer.Enabled = True
End Sub

Sub HSBSplashTimer_Timer()
    HSBSplashTimer.Enabled = False
    If HSAutoCommit Then
        HSAutoCommit = False
        HighScoreCommitName
    ElseIf DifficultyHSActive Then
        DifficultyHSDisplayName
    Else
        HighScoreDisplayName
    End If
End Sub

Sub EnterHighScoreKey(keycode)
    If keycode = LeftFlipperKey Then
        hsCurrentLetter = hsCurrentLetter - 1
        If hsCurrentLetter = 0 Then hsCurrentLetter = Len(hsValidLetters)
        HighScoreDisplayName
    End If
    If keycode = RightFlipperKey Then
        hsCurrentLetter = hsCurrentLetter + 1
        If hsCurrentLetter > Len(hsValidLetters) Then hsCurrentLetter = 1
        HighScoreDisplayName
    End If
    If keycode = 13 Or keycode = StartGameKey Then
        Dim letter : letter = Mid(hsValidLetters, hsCurrentLetter, 1)
        If letter = "<" Then
            If hsCurrentDigit > 0 Then
                hsCurrentDigit = hsCurrentDigit - 1
                hsEnteredDigits(hsCurrentDigit) = "-"
            End If
            HighScoreDisplayName
        Else
            hsEnteredDigits(hsCurrentDigit) = letter
            hsCurrentDigit = hsCurrentDigit + 1
            If hsCurrentDigit >= 3 Then
                HighScoreCommitName
            Else
                HighScoreDisplayName
            End If
        End If
    End If
End Sub

Sub HighScoreDisplayName()
    Dim entered : entered = ""
    Dim i
    For i = 0 To 2
        If i < hsCurrentDigit Then
            entered = entered & hsEnteredDigits(i)
        ElseIf i = hsCurrentDigit Then
            entered = entered & Mid(hsValidLetters, hsCurrentLetter, 1)
        Else
            entered = entered & "-"
        End If
    Next
    UpdateDMD2 "HIGH SCORE #" & EnterNameSpot, "INITIALS: " & entered
End Sub

Sub HighScoreCommitName()
    hsbModeActive = False
    Dim entered : entered = hsEnteredDigits(0) & hsEnteredDigits(1) & hsEnteredDigits(2)
    If entered = "---" Then entered = "D2X"
    HiName(EnterNameSpot - 1) = entered
    SaveHighScores
  Score = 0
    LastEnteredInitials = entered   ' hand High initials to Kill auto-save
    ShowMessage "SAVED: " & entered
    CheckKillHighScore
End Sub

Dim EnableRetractPlunger
EnableRetractPlunger = false

Dim AttractState : AttractState = 0
Dim AttractManualMode   : AttractManualMode   = False
Dim AttractManualTimer_active : AttractManualTimer_active = False

Sub StartTutorial()
    StopSound CurrentSong
    StopSound "introedit"   ' ← kill introedit explicitly in case it was
    CurrentSong = ""        '   started outside PlaySong (CancelTutorial path)
    PlayCallout "Cain_stayandlisten", 3000
    TutorialActive = True
    TutorialStep   = 0
  StopAttractLighting
    AttractTimer.Enabled      = False
    AttractSeqTimer.Enabled   = False
    AttractChaseTimer.Enabled = False
    AttractPulseTimer.Enabled = False
    StopTurntables
    RestoreAllGIRows
'    UpdateDMD2 "D2 TUTORIAL", "L.FLIP:BACK R.FLIP:NEXT"
    TutorialHintTimer.Interval = 2000
    TutorialHintTimer.Enabled  = True
End Sub

Sub TutFlipTimer_Timer()
    If Not TutorialActive Or TutorialStep <> 7 Then
        TutFlipTimer.Enabled = False
        UpperLeftFlipper.RotateToStart
        TutFlipCount = 0
        Exit Sub
    End If
    If TutFlipperUp Then
        ' Flipper down — count the completed flip
        UpperLeftFlipper.RotateToStart
        TutFlipperUp = False
        TutFlipCount = TutFlipCount + 1
        If TutFlipCount >= 3 Then
            ' 3 flips done — fire the jump
            TutFlipTimer.Enabled = False
            TutFlipCount = 0
            LeapReady = False
            BarbLeanTimer.Enabled = False
            BarbJumpStep = 0
            BarbJumpTimer.Enabled = True
            PlaySound "Leap1", 0, 1
            SetFlasherColor 1, 255, 180, 0 : SetFlasherColor 2, 255, 180, 0
            SetFlasherColor 3, 255, 180, 0 : SetFlasherColor 4, 255, 180, 0
            SetFlasherColor 5, 255, 180, 0 : SetFlasherColor 6, 255, 180, 0
            FireAllFlashers
            KillFlasherDimTimer.Enabled = True
            TutLeapAnimTimer.Interval = 3000
            TutLeapAnimTimer.Enabled = True  ' schedule PHASE 3
        Else
            TutFlipTimer.Interval = 500  ' pause between flips
        End If
    Else
        ' Flipper up
        UpperLeftFlipper.RotateToEnd
        TutFlipperUp = True
        TutFlipTimer.Interval = 350  ' hold at end
    End If
End Sub

Sub TutorialHintTimer_Timer()
    TutorialHintTimer.Enabled = False
    ShowTutorialStep 0
End Sub

Sub CancelTutorial()
    TutorialActive = False
    TutorialStep   = 0
    TutLeftHeld    = False
    TutRightHeld   = False
    ' Stop hint timer immediately so ShowTutorialStep can't fire after cancel
    TutorialHintTimer.Enabled = False
    ClearTutorialLights
    ' Stop introedit before restarting it to prevent double-play overlap
    StopSound "introedit"
    StartAttractMode
    PlaySound "introedit", 1, 1, 0, 0, 0, 0, 1
End Sub
Sub ClearTutorialLights()
  StopSound "tut_bumpers"
    StopSound "tut_travel"
    StopSound "tut_boss"
    StopSound "tut_loot"
    StopSound "tut_runeword"
    StopSound "tut_cube"
    StopSound "tut_save"
    StopSound "tut_leap"
    StopSound "tut_mystery"
    StopSound "tut_backstab"
    StopSound "tut_whirlwind"
    StopSound "tut_mercenary"
    ' Bumpers — clear VPX light state AND the FlFadeBumper intensity system
    Dim bn : For bn = 1 To 5
        FlBumperSmallLight(bn).State = 0
        FlbumperBigLight(bn).State   = 0
        FlBumperFadeTarget(bn) = 0
        FlBumperFadeActual(bn) = 0
        FlFadeBumper bn, 0
    Next
    ' Quest lights — stop crit cycle if it was running for tut_boss
    StopCritCycle
    QuestLight3.State = 0
    ' Cap ball / boss
    CapBallLight.State = 0
    ' Loot targets — drop and fully reset prims (glow, levitation, spin)
    ClearLootTargets
    ' Gear slot lights
    Dim gsi : For gsi = 0 To 6 : SlotToLight(gsi).State = 0 : Next
    ' Rune lights
    RuneKickLight.State = 0
    RuneLightR.State  = 0 : RuneLightU.State  = 0 : RuneLightN.State = 0
    RuneLightE.State  = 0 : RuneLightW.State  = 0 : RuneLightO.State = 0
    RuneLightR2.State = 0 : RuneLightD.State  = 0
    ' Cube light
    AttractPulseTimer.Enabled = False
    AttractPulseStep = 0
    CubeLight.State = 0
    ' SAVE / shield lights
    S_Light.State = 0 : A_Light.State = 0
    V_Light.State = 0 : E_Light.State = 0
    ShieldLight.State = 0
    ' Leap — stop flipper animation
    TutFlipTimer.Enabled = False
    TutFlipCount = 0
    UpperLeftFlipper.RotateToStart
    ' Mystery — stop GI event and raise ramp if it was lowered
    GIEventMode = GI_MODE_NONE
    GIEventTimer.Enabled = False
    RestoreAllGIRows
    If ChestRamp.RotX < ChestRampRest Then ChestRampSilent = True : ChestRampUp
'    MysteryKickLight.State = 0
    TutLeapAnimTimer.Enabled = False
    TutFlipTimer.Enabled = False
    UpperLeftFlipper.RotateToStart
    TutFlipCount = 0
    TutFlipperUp = False
    BarbLeanTimer.Enabled = False
    BarbJumpTimer.Enabled = False
    LeapReady = False
    BarbPrim.RotX = BarbRestRotX
    StopSound "TOM_Trunk_Motor_Long"
  ' Tutorial demo loop + FX (Whirlwind / Mercenary)
    TutDemoTimer.Enabled = False
    TutWWStep = 0
    FlasherSweepTimer.Enabled = False
    MercFlashTimer.Enabled = False
    Dim tdf : For tdf = 1 To 6 : DimFlasher tdf : Next
    townportalPF.Visible = False
End Sub

Sub ShowTutorialStep(step)
    ClearTutorialLights

    Select Case step
        Case 0 ' ---- BUMPERS ----
            UpdateDMD "BUMPERS >"
             Dim bn0 : For bn0 = 1 To 5
                FlBumperSmallLight(bn0).Color        = RGB(255, 80, 0)
                FlBumperSmallLight(bn0).ColorFull    = RGB(255, 80, 0)
                FlBumperSmallLight(bn0).intensity    = 500
                FlBumperSmallLight(bn0).BlinkInterval = 300
                FlBumperSmallLight(bn0).State        = 2
                FlbumperBigLight(bn0).Color          = RGB(255, 80, 0)
                FlbumperBigLight(bn0).ColorFull      = RGB(255, 80, 0)
                FlbumperBigLight(bn0).intensity      = 80
                FlbumperBigLight(bn0).BlinkInterval  = 300
                FlbumperBigLight(bn0).State          = 2
            Next
            PlaySound "tut_bumpers", 0, 1, 0, 0, 0, 0, 1

        Case 1 ' ---- TRAVEL / QUEST LANES ----
            UpdateDMD "< QUEST LANES >"
            QuestLight1.Color = RGB(255,100,0) : QuestLight1.ColorFull = RGB(255,140,0)
            QuestLight2.Color = RGB(255,100,0) : QuestLight2.ColorFull = RGB(255,140,0)
            QuestLight4.Color = RGB(255,100,0) : QuestLight4.ColorFull = RGB(255,140,0)
            QuestLight1.BlinkInterval = 400 : QuestLight1.State = 2
            QuestLight2.BlinkInterval = 400 : QuestLight2.State = 2
            QuestLight4.BlinkInterval = 400 : QuestLight4.State = 2
            PlaySound "tut_travel", 0, 1, 0, 0, 0, 0, 1

        Case 2 ' ---- BOSS FIGHT ----
            UpdateDMD "< BOSS FIGHT >"
            CapBallLight.Color = RGB(180,0,0) : CapBallLight.ColorFull = RGB(180,0,0)
            CapBallLight.BlinkInterval = 250 : CapBallLight.State = 2
            CritCycleActive = True
            CritCycleStep = 0
            SetCritCycleLight CritCycleStep
            CritCycleTimer.Interval = 1000
            CritCycleTimer.Enabled = True
            PlaySound "tut_boss", 0, 1, 0, 0, 0, 0, 1

        Case 3 ' ---- LOOT DROP TARGETS ----
            UpdateDMD "< LOOT DROPS >"
            ' Demo each rarity using the real prim system (color/glow/levitate)
            LootActive(0) = True : LootTier(0) = GEAR_UNIQUE
            dtLoot1.IsDropped = False : SetLootPrimAppearance 0, GEAR_UNIQUE : SetDtLootAnim 0, False
            LootActive(1) = True : LootTier(1) = GEAR_RARE
            dtLoot2.IsDropped = False : SetLootPrimAppearance 1, GEAR_RARE   : SetDtLootAnim 1, False
            LootActive(2) = True : LootTier(2) = GEAR_MAGIC
            dtLoot3.IsDropped = False : SetLootPrimAppearance 2, GEAR_MAGIC  : SetDtLootAnim 2, False
            LootActive(3) = True : LootTier(3) = GEAR_UNIQUE
            dtLoot4.IsDropped = False : SetLootPrimAppearance 3, GEAR_UNIQUE : SetDtLootAnim 3, False
            LightGearSlot 0, GEAR_UNIQUE : SlotToLight(0).BlinkInterval = 400 : SlotToLight(0).State = 2
            LightGearSlot 1, GEAR_RARE   : SlotToLight(1).BlinkInterval = 400 : SlotToLight(1).State = 2
            LightGearSlot 2, GEAR_MAGIC  : SlotToLight(2).BlinkInterval = 400 : SlotToLight(2).State = 2
            LightGearSlot 3, GEAR_UNIQUE : SlotToLight(3).BlinkInterval = 400 : SlotToLight(3).State = 2
            LightGearSlot 4, GEAR_RARE   : SlotToLight(4).BlinkInterval = 400 : SlotToLight(4).State = 2
            LightGearSlot 5, GEAR_MAGIC  : SlotToLight(5).BlinkInterval = 400 : SlotToLight(5).State = 2
            LightGearSlot 6, GEAR_RARE   : SlotToLight(6).BlinkInterval = 400 : SlotToLight(6).State = 2
            PlaySound "tut_loot", 0, 1, 0, 0, 0, 0, 1

        Case 4 ' ---- RUNEWORD ----
            UpdateDMD "< RUNEWORD >"
            RuneLightR.BlinkInterval  = 180 : RuneLightR.State  = 2
            RuneLightU.BlinkInterval  = 180 : RuneLightU.State  = 2
            RuneLightN.BlinkInterval  = 180 : RuneLightN.State  = 2
            RuneLightE.BlinkInterval  = 180 : RuneLightE.State  = 2
            RuneLightW.BlinkInterval  = 180 : RuneLightW.State  = 2
            RuneLightO.BlinkInterval  = 180 : RuneLightO.State  = 2
            RuneLightR2.BlinkInterval = 180 : RuneLightR2.State = 2
            RuneLightD.BlinkInterval  = 180 : RuneLightD.State  = 2
            RuneKickLight.BlinkInterval = 350 : RuneKickLight.State = 2
            PlaySound "tut_runeword", 0, 1, 0, 0, 0, 0, 1

        Case 5 ' ---- HORADRIC CUBE ----
            UpdateDMD "< HORADRIC CUBE >"
            AttractPulseStep = 0
            AttractPulseTimer.Enabled = True
            PlaySound "tut_cube", 0, 1, 0, 0, 0, 0, 1

        Case 6 ' ---- SAVE LANES ----
            UpdateDMD "< SHIELD >"
            S_Light.BlinkInterval = 280 : S_Light.State = 2
            A_Light.BlinkInterval = 280 : A_Light.State = 2
            V_Light.BlinkInterval = 280 : V_Light.State = 2
            E_Light.BlinkInterval = 280 : E_Light.State = 2
            ShieldLight.BlinkInterval = 400 : ShieldLight.State = 2
            PlaySound "tut_save", 0, 1, 0, 0, 0, 0, 1

        Case 7 ' ---- LEAP ATTACK ----
            UpdateDMD "< LEAP ATTACK >"
            PlaySound "tut_leap", 0, 1, 0, 0, 0, 0, 1
      LeapReady = True
            BarbLeanStep = 0
            BarbPrim.RotX = BarbRestRotX

            ' Fire the game's native leaning animation
            BarbLeanTimer.Enabled = True

            ' Start our tutorial tracker to trigger the jump simulation in 3.5 seconds
            TutLeapAnimTimer.Interval = 3500
            TutLeapAnimTimer.Enabled = True

        Case 8 ' ---- MYSTERY CHEST ----
            UpdateDMD "< MYSTERY CHEST >"
            ChestRampDown
            StartGIEvent GI_MODE_MYSTERY
'            MysteryKickLight.Color        = RGB(180, 0, 255)
'            MysteryKickLight.ColorFull    = RGB(180, 0, 255)
'            MysteryKickLight.BlinkInterval = 250
'            MysteryKickLight.State        = 2
            QuestLight4.Color        = RGB(180, 0, 255)
            QuestLight4.ColorFull    = RGB(180, 0, 255)
            QuestLight4.BlinkInterval = 300
            QuestLight4.State        = 2
            PlaySound "tut_mystery", 0, 1, 0, 0, 0, 0, 1

        Case 9 ' ---- BACKSTAB ----
            UpdateDMD "< BACKSTAB >"
            QuestLight2.Color        = RGB(255, 140, 0)
            QuestLight2.ColorFull    = RGB(255, 140, 0)
            QuestLight2.BlinkInterval = 300
            QuestLight2.State        = 2
            PlaySound "tut_backstab", 0, 1, 0, 0, 0, 0, 1

    Case 10 ' ---- WHIRLWIND ----
            UpdateDMD "< WHIRLWIND >"
            PlaySound "tut_whirlwind", 0, 1, 0, 0, 0, 0, 1
            TutWWStep = 0
            QuestLight3.Color = RGB(255,255,255) : QuestLight3.ColorFull = RGB(255,255,255)
            QuestLight3.BlinkPattern = "10" : QuestLight3.BlinkInterval = 400 : QuestLight3.State = 2
            TutDemoTimer.Interval = 700
            TutDemoTimer.Enabled = True

        Case 11 ' ---- MERCENARY ----
            UpdateDMD "< MERCENARY >"
            PlaySound "tut_mercenary", 0, 1, 0, 0, 0, 0, 1
            townportalPF.Visible = True
            FireMercSweep
            TutDemoTimer.Interval = 5500
            TutDemoTimer.Enabled = True

    End Select
End Sub

Sub TutDemoTimer_Timer()
    If Not TutorialActive Or (TutorialStep <> 10 And TutorialStep <> 11) Then
        TutDemoTimer.Enabled = False
        Exit Sub
    End If

    ' ---- MERCENARY (step 11): keep portal lit, re-fire the gold sweep ----
    If TutorialStep = 11 Then
        townportalPF.Visible = True
        FireMercSweep
        Exit Sub
    End If

    ' ---- WHIRLWIND (step 10): charge x3, storm the bumpers, loop ----
    Dim wb
    Select Case TutWWStep
        Case 0 : QuestLight3.BlinkInterval = 400 : QuestLight3.State = 2   ' charge 1/3
        Case 1 : QuestLight3.BlinkInterval = 200 : QuestLight3.State = 2   ' charge 2/3
        Case 2 : QuestLight3.BlinkInterval = 90  : QuestLight3.State = 2   ' charge 3/3
        Case 3 ' storm — bumpers struck white + flasher sweep
            QuestLight3.BlinkInterval = 60 : QuestLight3.State = 2
            For wb = 1 To 5
                FlBumperSmallLight(wb).Color = RGB(255,255,255) : FlBumperSmallLight(wb).ColorFull = RGB(255,255,255)
                FlBumperSmallLight(wb).BlinkInterval = 120 : FlBumperSmallLight(wb).State = 2
                FlbumperBigLight(wb).Color = RGB(255,255,255)   : FlbumperBigLight(wb).ColorFull = RGB(255,255,255)
                FlbumperBigLight(wb).BlinkInterval = 120 : FlbumperBigLight(wb).State = 2
            Next
            FlasherSweepStep = 0 : FlasherSweepTimer.Interval = 30 : FlasherSweepTimer.Enabled = True
        Case 4 ' storm peak — second sweep
            FlasherSweepStep = 0 : FlasherSweepTimer.Interval = 30 : FlasherSweepTimer.Enabled = True
        Case 5 ' fade — clear, then loop
            QuestLight3.State = 0
            For wb = 1 To 5
                FlBumperSmallLight(wb).State = 0 : FlbumperBigLight(wb).State = 0
            Next
    End Select

    TutWWStep = TutWWStep + 1
    If TutWWStep > 5 Then TutWWStep = 0
End Sub

Sub TutLeapAnimTimer_Timer()
    ' Safety Check: correct step is 7 (Leap), not 8
    If Not TutorialActive Or TutorialStep <> 7 Then
        TutLeapAnimTimer.Enabled = False
        TutFlipTimer.Enabled = False
        UpperLeftFlipper.RotateToStart
        TutFlipCount = 0
        Exit Sub
    End If

    ' PHASE 2: Start 3 ULF flips — jump fires from TutFlipTimer after they complete
    If LeapReady Then
        TutFlipCount = 0
        TutFlipperUp = False
        TutLeapAnimTimer.Enabled = False  ' pause until TutFlipTimer finishes
        TutFlipTimer.Interval = 500
        TutFlipTimer.Enabled = True

    ' PHASE 3: Reset loop back to lean + wait
    Else
        LeapReady = True
        BarbJumpTimer.Enabled = False
        BarbLeanStep = 0
        BarbPrim.RotX = BarbRestRotX
        BarbLeanTimer.Enabled = True
        TutFlipCount = 0
        TutFlipperUp = False
        TutLeapAnimTimer.Interval = 2000
    End If
End Sub

Sub StartAttractMode()
    GameOverSequenceActive = False
    AttractFlashStep = 0
    StartTurntables
    Turntable1Trigger.Enabled = False
    Turntable2Trigger.Enabled = False
    AttractChaseStep = 0
    AttractPhase = 0
    AttractFlashStep = 0
    AttractPulseStep = 0
    AttractRuneColorStep = 0
    AttractCharStep = 0
    AttractGearStep = 0
    AttractTimer.Enabled = True
    AttractSeqTimer.Enabled = True
    AttractChaseTimer.Enabled = True
    AttractPulseTimer.Enabled = True
    ShowAttractFrame
    AttractSetBaseGI
End Sub

Sub AttractSetBaseGI()
    Dim agi
    For agi = 0 To GI.Count - 1
        GI.Item(agi).Color = GetActGIDimColor()
        GI.Item(agi).ColorFull = GetActGIDimColor()
    Next
End Sub

Sub AttractRestoreGI()
    SetActGI
End Sub

Sub ShowAttractFrame()
    Select Case AttractState
        Case 0
            If FreePlay Then
                UpdateDMD2 "PRESS START", "FREE PLAY"
            ElseIf Credits > 0 Then
                UpdateDMD2 "PRESS START", "CREDITS: " & Credits
            Else
                UpdateDMD2 "INSERT COIN", "TO PLAY"
            End If
        Case 1  : UpdateDMD "#1 " & HiName(0) & " " & FormatNumber(HiScore(0), 0, -1, 0, -1)
        Case 2  : UpdateDMD "#2 " & HiName(1) & " " & FormatNumber(HiScore(1), 0, -1, 0, -1)
        Case 3  : UpdateDMD "#3 " & HiName(2) & " " & FormatNumber(HiScore(2), 0, -1, 0, -1)
        Case 4  : UpdateDMD "#4 " & HiName(3) & " " & FormatNumber(HiScore(3), 0, -1, 0, -1)
        Case 5  : UpdateDMD "#5 " & HiName(4) & " " & FormatNumber(HiScore(4), 0, -1, 0, -1)
        Case 6  : UpdateDMD "TOP KILLS"
        Case 7  : UpdateDMD "#1 " & HiKillsName(0) & " " & HiKills(0) & " KILLS"
        Case 8  : UpdateDMD "#2 " & HiKillsName(1) & " " & HiKills(1) & " KILLS"
        Case 9  : UpdateDMD "#3 " & HiKillsName(2) & " " & HiKills(2) & " KILLS"
        Case 10 : UpdateDMD "#4 " & HiKillsName(3) & " " & HiKills(3) & " KILLS"
        Case 11 : UpdateDMD "#5 " & HiKillsName(4) & " " & HiKills(4) & " KILLS"
    Case 12 : UpdateDMD2 "NIGHTMARE CHAMP", NightmareChampName & "  " & FormatNumber(NightmareChamp, 0, -1, 0, -1)
    Case 13 : UpdateDMD2 "HELL CHAMPION", HellChampName & "  " & FormatNumber(HellChamp, 0, -1, 0, -1)
    Case 14 : UpdateDMD2 "HOLD BOTH FLIPPERS", "TO ENTER TUTORIAL"
    End Select
End Sub

Sub AttractTimer_Timer()
    If GameActive Or TutorialActive Then
        AttractTimer.Enabled = False
        Exit Sub
    End If
    If AttractManualMode Then Exit Sub
    AttractState = AttractState + 1
    If AttractState > 14 Then AttractState = 0
    ShowAttractFrame
End Sub

Sub AttractManualResumeTimer_Timer()
    AttractManualResumeTimer.Enabled = False
    AttractManualMode = False
End Sub

'*****************************************
'      ATTRACT MODE LIGHTING
'*****************************************

Sub StopAttractLighting()
    AttractTimer.Enabled = False
    AttractSeqTimer.Enabled = False
    AttractChaseTimer.Enabled = False
    AttractPulseTimer.Enabled = False
    AttractFlashTimer.Enabled = False
    StopTurntables
    RuneLightR.State  = 0 : RuneLightU.State  = 0
    RuneLightN.State  = 0 : RuneLightE.State  = 0
    RuneLightW.State  = 0 : RuneLightO.State  = 0
    RuneLightR2.State = 0 : RuneLightD.State  = 0
    CharLightBarb.State  = 0 : CharLightAma.State   = 0
    CharLightNecro.State = 0 : CharLightSorc.State  = 0
    CharLightPal.State   = 0 : CharLightAss.State   = 0
    CharLightDru.State   = 0
    Dim gsi3
    For gsi3 = 0 To 6
        SlotToLight(gsi3).State = 0
    Next
    Dim fbi
    For fbi = 1 To 5
        FlBumperFadeTarget(fbi) = 0
    Next
    QuestLight1.State = 0 : QuestLight2.State = 0
    QuestLight4.State = 0
'    MysteryKickLight.State = 0
    SetRampLight
    RuneKickLight.State = 0
    CapBallLight.State = 0
    CubeLight.State = 0
    ShieldLight.State = 0
    BallSaveL.State = 0
    BallSaveL2.State = 0
    AttractRestoreGI
    DimAllFlashers
End Sub

Function AttractRuneColor()
    Select Case AttractRuneColorStep Mod 6
        Case 0 : AttractRuneColor = RGB(255, 50, 0)    ' deep red
        Case 1 : AttractRuneColor = RGB(255, 100, 0)   ' orange red
        Case 2 : AttractRuneColor = RGB(255, 150, 0)   ' orange
        Case 3 : AttractRuneColor = RGB(255, 200, 0)   ' amber
        Case 4 : AttractRuneColor = RGB(255, 120, 0)   ' orange
        Case 5 : AttractRuneColor = RGB(200, 40, 0)    ' dark red
    End Select
End Function

Sub AttractChaseTimer_Timer()
    If GameActive Or TutorialActive Then AttractChaseTimer.Enabled = False : Exit Sub
    AttractChaseStep = AttractChaseStep + 1
    Select Case AttractPhase
        Case 0, 1
            Dim runeIdx : runeIdx = (AttractChaseStep - 1) Mod 8
            AttractRuneColorStep = AttractRuneColorStep + 1
            RuneLightR.State  = 0 : RuneLightU.State  = 0
            RuneLightN.State  = 0 : RuneLightE.State  = 0
            RuneLightW.State  = 0 : RuneLightO.State  = 0
            RuneLightR2.State = 0 : RuneLightD.State  = 0
            Dim rc : rc = AttractRuneColor()
            Select Case runeIdx
                Case 0 : RuneLightR.Color  = rc : RuneLightR.ColorFull  = rc : RuneLightR.State  = 1
                Case 1 : RuneLightU.Color  = rc : RuneLightU.ColorFull  = rc : RuneLightU.State  = 1
                Case 2 : RuneLightN.Color  = rc : RuneLightN.ColorFull  = rc : RuneLightN.State  = 1
                Case 3 : RuneLightE.Color  = rc : RuneLightE.ColorFull  = rc : RuneLightE.State  = 1
                Case 4 : RuneLightW.Color  = rc : RuneLightW.ColorFull  = rc : RuneLightW.State  = 1
                Case 5 : RuneLightO.Color  = rc : RuneLightO.ColorFull  = rc : RuneLightO.State  = 1
                Case 6 : RuneLightR2.Color = rc : RuneLightR2.ColorFull = rc : RuneLightR2.State = 1
                Case 7 : RuneLightD.Color  = rc : RuneLightD.ColorFull  = rc : RuneLightD.State  = 1
            End Select

        Case 2
            Dim bIdx : bIdx = (AttractChaseStep Mod 5) + 1
            Dim prevIdx : prevIdx = ((AttractChaseStep - 1) Mod 5) + 1
            FlBumperFadeTarget(prevIdx) = 0
            FlInitBumper bIdx, "red"
            FlBumperFadeActual(bIdx) = 0
            FlBumperFadeTarget(bIdx) = 0.4

        Case 3
            Dim qIdx : qIdx = AttractChaseStep Mod 8
            QuestLight1.State = 0 : QuestLight2.State = 0
            QuestLight4.State = 0
            CapBallLight.State = 0
            Select Case qIdx
                Case 0 : QuestLight1.Color = RGB(255, 80, 0)  : QuestLight1.ColorFull = RGB(255, 80, 0)  : QuestLight1.State = 1
                Case 1 : QuestLight2.Color = RGB(255, 80, 0)  : QuestLight2.ColorFull = RGB(255, 80, 0)  : QuestLight2.State = 1
                Case 2 : QuestLight4.Color = RGB(255, 80, 0)  : QuestLight4.ColorFull = RGB(255, 80, 0)  : QuestLight4.State = 1
                Case 3 : QuestLight4.Color = RGB(200, 40, 0)  : QuestLight4.ColorFull = RGB(200, 40, 0)  : QuestLight4.State = 1 : CapBallLight.Color = RGB(180, 0, 0) : CapBallLight.ColorFull = RGB(180, 0, 0) : CapBallLight.State = 1
                Case 4 : QuestLight2.Color = RGB(200, 40, 0)  : QuestLight2.ColorFull = RGB(200, 40, 0)  : QuestLight2.State = 1 : CapBallLight.State = 1
                Case 5 : QuestLight1.Color = RGB(200, 40, 0)  : QuestLight1.ColorFull = RGB(200, 40, 0)  : QuestLight1.State = 1 : CapBallLight.State = 1
            End Select
            SetRampLight
        Case 4
            Dim cIdx : cIdx = AttractChaseStep Mod 7
            Select Case cIdx
                Case 0 : CharLightBarb.State  = 1
                Case 1 : CharLightAma.State   = 1
                Case 2 : CharLightNecro.State = 1
                Case 3 : CharLightSorc.State  = 1
                Case 4 : CharLightPal.State   = 1
                Case 5 : CharLightAss.State   = 1
                Case 6 : CharLightDru.State   = 1
            End Select

        Case 5
            AttractGearStep = AttractGearStep + 1
            Dim gIdx : gIdx = AttractGearStep Mod 7
            Dim gColor
            Select Case AttractGearStep Mod 3
                Case 0 : gColor = RGB(105, 105, 255)
                Case 1 : gColor = RGB(255, 255, 100)
                Case 2 : gColor = RGB(180, 100, 20)
            End Select
            Dim gLight : Set gLight = SlotToLight(gIdx)
            gLight.Color = gColor
            gLight.ColorFull = gColor
            gLight.State = 1

        Case 6
            Dim runeIdxB : runeIdxB = 7 - ((AttractChaseStep - 1) Mod 8)
            AttractRuneColorStep = AttractRuneColorStep + 1
            RuneLightR.State  = 0 : RuneLightU.State  = 0
            RuneLightN.State  = 0 : RuneLightE.State  = 0
            RuneLightW.State  = 0 : RuneLightO.State  = 0
            RuneLightR2.State = 0 : RuneLightD.State  = 0
            Dim rcB : rcB = AttractRuneColor()
            Select Case runeIdxB
                Case 0 : RuneLightR.Color  = rcB : RuneLightR.ColorFull  = rcB : RuneLightR.State  = 1
                Case 1 : RuneLightU.Color  = rcB : RuneLightU.ColorFull  = rcB : RuneLightU.State  = 1
                Case 2 : RuneLightN.Color  = rcB : RuneLightN.ColorFull  = rcB : RuneLightN.State  = 1
                Case 3 : RuneLightE.Color  = rcB : RuneLightE.ColorFull  = rcB : RuneLightE.State  = 1
                Case 4 : RuneLightW.Color  = rcB : RuneLightW.ColorFull  = rcB : RuneLightW.State  = 1
                Case 5 : RuneLightO.Color  = rcB : RuneLightO.ColorFull  = rcB : RuneLightO.State  = 1
                Case 6 : RuneLightR2.Color = rcB : RuneLightR2.ColorFull = rcB : RuneLightR2.State = 1
                Case 7 : RuneLightD.Color  = rcB : RuneLightD.ColorFull  = rcB : RuneLightD.State  = 1
            End Select

        Case 7
            Dim giIdx : giIdx = AttractChaseStep Mod GI.Count
            Dim giPass : giPass = Int(AttractChaseStep / GI.Count)
            If giPass < 3 Then
                Dim giColor
                Select Case giPass
                    Case 0 : giColor = RGB(255, 50, 0)
                    Case 1 : giColor = RGB(255, 120, 0)
                    Case 2 : giColor = RGB(255, 200, 0)
                End Select
                Dim agi3
                For agi3 = 0 To GI.Count - 1
                    GI.Item(agi3).Color = GetActGIDimColor()
                    GI.Item(agi3).ColorFull = GetActGIDimColor()
                    GI.Item(agi3).Intensity = GIFlickerBase
                Next
                GI.Item(giIdx).Color = giColor
                GI.Item(giIdx).ColorFull = giColor
                GI.Item(giIdx).Intensity = 15
            End If

    Case 8  ' Half-time slow sweep — each light holds twice as long
            Dim htIdx : htIdx = Int((AttractChaseStep - 1) / 2) Mod 8
            Dim htPrev : htPrev = Int((AttractChaseStep - 2) / 2) Mod 8
            If (AttractChaseStep Mod 2) = 1 Then
                ' Turn off previous
                RuneLightR.State  = 0 : RuneLightU.State  = 0
                RuneLightN.State  = 0 : RuneLightE.State  = 0
                RuneLightW.State  = 0 : RuneLightO.State  = 0
                RuneLightR2.State = 0 : RuneLightD.State  = 0
                ' Light current
                AttractRuneColorStep = AttractRuneColorStep + 1
                Dim rcH : rcH = AttractRuneColor()
                Select Case htIdx
                    Case 0 : RuneLightR.Color  = rcH : RuneLightR.ColorFull  = rcH : RuneLightR.State  = 1
                    Case 1 : RuneLightU.Color  = rcH : RuneLightU.ColorFull  = rcH : RuneLightU.State  = 1
                    Case 2 : RuneLightN.Color  = rcH : RuneLightN.ColorFull  = rcH : RuneLightN.State  = 1
                    Case 3 : RuneLightE.Color  = rcH : RuneLightE.ColorFull  = rcH : RuneLightE.State  = 1
                    Case 4 : RuneLightW.Color  = rcH : RuneLightW.ColorFull  = rcH : RuneLightW.State  = 1
                    Case 5 : RuneLightO.Color  = rcH : RuneLightO.ColorFull  = rcH : RuneLightO.State  = 1
                    Case 6 : RuneLightR2.Color = rcH : RuneLightR2.ColorFull = rcH : RuneLightR2.State = 1
                    Case 7 : RuneLightD.Color  = rcH : RuneLightD.ColorFull  = rcH : RuneLightD.State  = 1
                End Select
            End If

        Case 9  ' Pairs pulse from ends toward center — R+D, U+R2, N+O, E+W
            Dim pairStep : pairStep = AttractChaseStep Mod 8
            RuneLightR.State  = 0 : RuneLightU.State  = 0
            RuneLightN.State  = 0 : RuneLightE.State  = 0
            RuneLightW.State  = 0 : RuneLightO.State  = 0
            RuneLightR2.State = 0 : RuneLightD.State  = 0
            AttractRuneColorStep = AttractRuneColorStep + 1
            Dim rcP : rcP = AttractRuneColor()
            Select Case pairStep
                Case 0, 1  ' R + D (outer ends) — on then off
                    If pairStep = 0 Then
                        RuneLightR.Color  = rcP : RuneLightR.ColorFull  = rcP : RuneLightR.State  = 1
                        RuneLightD.Color  = rcP : RuneLightD.ColorFull  = rcP : RuneLightD.State  = 1
                    End If
                Case 2, 3  ' U + R2
                    If pairStep = 2 Then
                        RuneLightU.Color  = rcP : RuneLightU.ColorFull  = rcP : RuneLightU.State  = 1
                        RuneLightR2.Color = rcP : RuneLightR2.ColorFull = rcP : RuneLightR2.State = 1
                    End If
                Case 4, 5  ' N + O
                    If pairStep = 4 Then
                        RuneLightN.Color  = rcP : RuneLightN.ColorFull  = rcP : RuneLightN.State  = 1
                        RuneLightO.Color  = rcP : RuneLightO.ColorFull  = rcP : RuneLightO.State  = 1
                    End If
                Case 6, 7  ' E + W (center meeting point)
                    If pairStep = 6 Then
                        RuneLightE.Color  = rcP : RuneLightE.ColorFull  = rcP : RuneLightE.State  = 1
                        RuneLightW.Color  = rcP : RuneLightW.ColorFull  = rcP : RuneLightW.State  = 1
                    End If
            End Select

    End Select
End Sub

Sub AttractPulseTimer_Timer()
    If GameActive Then AttractPulseTimer.Enabled = False : Exit Sub
    AttractPulseStep = AttractPulseStep + 1
    If TutorialActive Then
        Dim teR, teG, teB
        Select Case (AttractPulseStep \ 8) Mod 4
            Case 0 : teR = 80  : teG = 180 : teB = 255
            Case 1 : teR = 255 : teG = 60  : teB = 0
            Case 2 : teR = 196 : teG = 196 : teB = 8
            Case 3 : teR = 0   : teG = 200 : teB = 50
        End Select
        CubeLight.Color     = RGB(teR, teG, teB)
        CubeLight.ColorFull = RGB(teR, teG, teB)
        CubeLight.State     = 1
        Exit Sub
    End If
    If GIEventMode <> GI_MODE_NONE Then Exit Sub
    Dim pulseR, pulseG
    If AttractPulseStep Mod 2 = 0 Then
        pulseR = 180 : pulseG = 40
    Else
        pulseR = 220 : pulseG = 70
    End If
    Dim agi2
    For agi2 = 0 To GI.Count - 1
        GI.Item(agi2).Color = RGB(pulseR, pulseG, 0)
        GI.Item(agi2).ColorFull = RGB(pulseR, pulseG, 0)
    Next
    ' MysteryKick and RuneKick keep their existing pulse
    If AttractPulseStep Mod 4 = 0 Then
'        MysteryKickLight.Color = RGB(255, 80, 0) : MysteryKickLight.State = 1
        RuneKickLight.Color    = RGB(255, 60, 0) : RuneKickLight.State    = 1
    ElseIf AttractPulseStep Mod 4 = 2 Then
'        MysteryKickLight.State = 0
        RuneKickLight.State    = 0
    End If
    ' CubeLight cycles element colors
    Dim eR, eG, eB
    Select Case (AttractPulseStep \ 8) Mod 4
        Case 0 : eR = 80  : eG = 180 : eB = 255  ' Cold
        Case 1 : eR = 255 : eG = 60  : eB = 0    ' Fire
        Case 2 : eR = 196 : eG = 196 : eB = 8    ' Lightning
        Case 3 : eR = 0   : eG = 200 : eB = 50   ' Poison
    End Select
    CubeLight.Color     = RGB(eR, eG, eB)
    CubeLight.ColorFull = RGB(eR, eG, eB)
    CubeLight.State     = 1
End Sub

Sub AttractSeqTimer_Timer()
    If GameActive Or TutorialActive Then AttractSeqTimer.Enabled = False : Exit Sub
    ' Always clear rune lights when leaving a rune phase
    If AttractPhase = 0 Or AttractPhase = 1 Or AttractPhase = 6 Then
        RuneLightR.State  = 0 : RuneLightU.State  = 0
        RuneLightN.State  = 0 : RuneLightE.State  = 0
        RuneLightW.State  = 0 : RuneLightO.State  = 0
        RuneLightR2.State = 0 : RuneLightD.State  = 0
    End If
    AttractPhase = AttractPhase + 1
    AttractChaseStep = 0
    If AttractPhase > 9 Then
        AttractPhase = 0
        ' Reset all lights for next cycle
        RuneLightR.State  = 0 : RuneLightU.State  = 0
        RuneLightN.State  = 0 : RuneLightE.State  = 0
        RuneLightW.State  = 0 : RuneLightO.State  = 0
        RuneLightR2.State = 0 : RuneLightD.State  = 0
        CharLightBarb.State  = 0 : CharLightAma.State   = 0
        CharLightNecro.State = 0 : CharLightSorc.State  = 0
        CharLightPal.State   = 0 : CharLightAss.State   = 0
        CharLightDru.State   = 0
        Dim gsi4
        For gsi4 = 0 To 6
            SlotToLight(gsi4).State = 0
        Next
        ' Trigger hellfire flash at cycle end
        AttractFlashStep = 0
        AttractFlashTimer.Enabled = True
    End If
End Sub

Sub AttractFlashTimer_Timer()
    If GameActive Or TutorialActive Then AttractFlashTimer.Enabled = False : Exit Sub
    AttractFlashStep = AttractFlashStep + 1
    Select Case AttractFlashStep
        Case 1
            Dim afi
            For afi = 0 To GI.Count - 1
                GI.Item(afi).Color = RGB(255, 30, 0)
                GI.Item(afi).ColorFull = RGB(255, 30, 0)
                GI.Item(afi).Intensity = 20
            Next
            Dim afb
            For afb = 1 To 5
                FlInitBumper afb, "red"
                FlBumperFadeActual(afb) = 0
                FlBumperFadeTarget(afb) = 0.3
            Next
            RuneLightR.Color  = RGB(255, 50, 0) : RuneLightR.ColorFull  = RGB(255, 50, 0) : RuneLightR.State  = 1
            RuneLightU.Color  = RGB(255, 50, 0) : RuneLightU.ColorFull  = RGB(255, 50, 0) : RuneLightU.State  = 1
            RuneLightN.Color  = RGB(255, 50, 0) : RuneLightN.ColorFull  = RGB(255, 50, 0) : RuneLightN.State  = 1
            RuneLightE.Color  = RGB(255, 50, 0) : RuneLightE.ColorFull  = RGB(255, 50, 0) : RuneLightE.State  = 1
            RuneLightW.Color  = RGB(255, 50, 0) : RuneLightW.ColorFull  = RGB(255, 50, 0) : RuneLightW.State  = 1
            RuneLightO.Color  = RGB(255, 50, 0) : RuneLightO.ColorFull  = RGB(255, 50, 0) : RuneLightO.State  = 1
            RuneLightR2.Color = RGB(255, 50, 0) : RuneLightR2.ColorFull = RGB(255, 50, 0) : RuneLightR2.State = 1
            RuneLightD.Color  = RGB(255, 50, 0) : RuneLightD.ColorFull  = RGB(255, 50, 0) : RuneLightD.State  = 1
            QuestLight1.Color = RGB(255, 30, 0) : QuestLight1.ColorFull = RGB(255, 30, 0) : QuestLight1.State = 1
            QuestLight2.Color = RGB(255, 30, 0) : QuestLight2.ColorFull = RGB(255, 30, 0) : QuestLight2.State = 1
            QuestLight3.Color = RGB(255, 30, 0) : QuestLight3.ColorFull = RGB(255, 30, 0) : QuestLight3.State = 1
            QuestLight4.Color = RGB(255, 30, 0) : QuestLight4.ColorFull = RGB(255, 30, 0) : QuestLight4.State = 1
            CapBallLight.Color = RGB(255, 30, 0) : CapBallLight.ColorFull = RGB(255, 30, 0) : CapBallLight.State = 1
'            MysteryKickLight.Color = RGB(255, 30, 0) : MysteryKickLight.ColorFull = RGB(255, 30, 0) : MysteryKickLight.State = 1
            CubeLight.Color = RGB(255, 30, 0) : CubeLight.ColorFull = RGB(255, 30, 0) : CubeLight.State = 1
            ShieldLight.Color = RGB(255, 30, 0) : ShieldLight.ColorFull = RGB(255, 30, 0) : ShieldLight.State = 1
            BallSaveL.Color = RGB(174, 0, 0) : BallSaveL.ColorFull = RGB(255, 72, 72) : BallSaveL.State = 1
            BallSaveL2.Color = RGB(174, 0, 0) : BallSaveL2.ColorFull = RGB(255, 72, 72) : BallSaveL2.State = 1
            SetFlasherColor 1, 255, 50, 0 : SetFlasherColor 2, 255, 50, 0
            SetFlasherColor 3, 255, 50, 0 : SetFlasherColor 4, 255, 50, 0
            SetFlasherColor 5, 255, 50, 0 : SetFlasherColor 6, 255, 50, 0
            FireAllFlashers
        Case 2
            Dim af2
            For af2 = 0 To GI.Count - 1
                GI.Item(af2).Intensity = 10
            Next
        Case 3
            Dim af3
            For af3 = 0 To GI.Count - 1
                GI.Item(af3).Color = RGB(255, 100, 0)
                GI.Item(af3).ColorFull = RGB(255, 100, 0)
                GI.Item(af3).Intensity = 15
            Next
        Case 4
            Dim af4
            For af4 = 0 To GI.Count - 1
                GI.Item(af4).Intensity = 8
            Next
        Case 5
            Dim af5
            For af5 = 0 To GI.Count - 1
                GI.Item(af5).Intensity = 2
            Next
            RuneLightR.State  = 0 : RuneLightU.State  = 0
            RuneLightN.State  = 0 : RuneLightE.State  = 0
            RuneLightW.State  = 0 : RuneLightO.State  = 0
            RuneLightR2.State = 0 : RuneLightD.State  = 0
            QuestLight1.State = 0 : QuestLight2.State = 0
            QuestLight4.State = 0
            CapBallLight.State = 0
            SetRampLight
'            MysteryKickLight.State = 0
            CubeLight.State = 0
            ShieldLight.State = 0
            BallSaveL.State = 0
            BallSaveL2.State = 0
            CapBallLight.State = 0
        Case 6
            Case 6
            ' Left-to-right sweep
            SetFlasherColor 1, 200, 0, 0 : FireFlasher 1
        Case 7
            SetFlasherColor 2, 200, 0, 0 : FireFlasher 2 : DimFlasher 1
        Case 8
            SetFlasherColor 3, 200, 0, 0 : FireFlasher 3 : DimFlasher 2
        Case 9
            SetFlasherColor 4, 200, 0, 0 : FireFlasher 4 : DimFlasher 3
        Case 10
            SetFlasherColor 5, 200, 0, 0 : FireFlasher 5 : DimFlasher 4
        Case 11
            SetFlasherColor 6, 200, 0, 0 : FireFlasher 6 : DimFlasher 5
        Case 12
            DimAllFlashers
        Case 13
            ' Double pulse - all orange
            SetFlasherColor 1, 255, 80, 0 : SetFlasherColor 2, 255, 80, 0
            SetFlasherColor 3, 255, 80, 0 : SetFlasherColor 4, 255, 80, 0
            SetFlasherColor 5, 255, 80, 0 : SetFlasherColor 6, 255, 80, 0
            FireAllFlashers
        Case 14
            DimAllFlashers
        Case 15
            FireAllFlashers
        Case 16
            DimAllFlashers
        Case 17
            ' Slow deep red throb
            AttractFlashTimer.Interval = 500
            SetFlasherColor 1, 120, 0, 0 : SetFlasherColor 2, 120, 0, 0
            SetFlasherColor 3, 120, 0, 0 : SetFlasherColor 4, 120, 0, 0
            SetFlasherColor 5, 120, 0, 0 : SetFlasherColor 6, 120, 0, 0
            FireAllFlashers
        Case 18
            DimAllFlashers
        Case 19
            AttractFlashTimer.Interval = 100
            AttractFlashTimer.Enabled = False
            DimAllFlashers
            Dim af6
            For af6 = 0 To GI.Count - 1
                GI.Item(af6).Color = GetActGIDimColor()
                GI.Item(af6).ColorFull = GetActGIDimColor()
                GI.Item(af6).Intensity = GIFlickerBase
            Next
    End Select
End Sub

Sub ShowMessage(msg)
  If DifficultySelectActive Then Exit Sub
    If MysteryActive Then Exit Sub
    If LeapReady Then Exit Sub
    UpdateDMD msg
    MsgQueueTimer.Enabled = False
    MsgQueueTimer.Enabled = True
End Sub

Sub ShowMessagePriority(msg, priority)
    ShowMessage msg
End Sub

Sub MsgQueueTimer_Timer()
    MsgQueueTimer.Enabled = False
    UpdateDMDScore
End Sub

'FLEXDMD STUFF

Const FlexDMD_RenderMode_DMD_GRAY   = 0
Const FlexDMD_RenderMode_DMD_GRAY_4 = 1
Const FlexDMD_RenderMode_DMD_RGB    = 2
Const FlexDMD_Align_Center          = 4

Dim FlexDMD
Dim FlexDMDActive  : FlexDMDActive  = False
Dim DMDVideoActive : DMDVideoActive = False
Dim FlexIntroScene
Dim FontMain
Dim FontSmall
Dim FontTiny

Sub InitFlexDMD()
    On Error Resume Next
    Set FlexDMD = CreateObject("FlexDMD.FlexDMD")
    If Err.Number <> 0 Or FlexDMD Is Nothing Then
        FlexDMDActive = False
        On Error Goto 0
        Exit Sub
    End If
    SetLocale(1033)
    FlexDMD.GameName      = "Diablo2"
    FlexDMD.TableFile     = Table1.Filename & ".vpx"
    FlexDMD.Color         = RGB(180, 0, 0)
    FlexDMD.RenderMode    = FlexDMD_RenderMode_DMD_RGB
    FlexDMD.Width         = 128
    FlexDMD.Height        = 32
    FlexDMD.Clear         = True
    FlexDMD.ProjectFolder = "./Diablo 2.FlexDMD/"
    FlexDMD.Run           = True
    If Err.Number <> 0 Then
        FlexDMDActive = False
        On Error Goto 0
        Exit Sub
    End If
    Set FontMain  = FlexDMD.NewFont("FlexDMD.Resources.udmd-f5by7.fnt", RGB(180, 0, 0), RGB(1, 1, 1), 1)
  Set FontSmall = FlexDMD.NewFont("FlexDMD.Resources.udmd-f5by7.fnt", RGB(180, 0, 0), RGB(1, 1, 1), 0)
  Set FontTiny  = FlexDMD.NewFont("FlexDMD.Resources.udmd-f4by5.fnt", RGB(180, 0, 0), RGB(1, 1, 1), 0)
  Dim FontBig
  Set FontBig   = FlexDMD.NewFont("FlexDMD.Resources.udmd-f5by7.fnt", RGB(180, 0, 0), RGB(70, 0, 0), 1)

    ' Pre-build score scene
    Dim scene
    Set scene = FlexDMD.NewGroup("Score")
    Dim bgImg : Set bgImg = FlexDMD.NewImage("BG", "d2_bg_dmd.png")
    bgImg.SetBounds 0, 0, 128, 32
    bgImg.Scaling = 4
    scene.AddActor bgImg
    scene.AddActor FlexDMD.NewLabel("Line1", FontSmall, "DIABLO II")
    scene.GetLabel("Line1").SetAlignedPosition 64, 10, FlexDMD_Align_Center
    scene.AddActor FlexDMD.NewLabel("Line2", FontSmall, "PRESS START")
    scene.GetLabel("Line2").SetAlignedPosition 64, 23, FlexDMD_Align_Center
  scene.AddActor FlexDMD.NewLabel("BigText", FontBig, "")
    scene.GetLabel("BigText").SetAlignedPosition 64, 16, FlexDMD_Align_Center
    scene.GetLabel("BigText").Visible = False
  scene.AddActor FlexDMD.NewLabel("Line3", FontTiny, "")
    scene.AddActor FlexDMD.NewLabel("Line4", FontTiny, "")

    ' Pre-load intro video scene at init so decoder is ready
'    Set FlexIntroScene = FlexDMD.NewGroup("IntroScene")
'    FlexIntroScene.AddActor FlexDMD.NewVideo("IntroVid", "d2_logo_dmd.gif")
'    FlexIntroScene.GetVideo("IntroVid").SetBounds 0, 0, 128, 32
'    FlexIntroScene.GetVideo("IntroVid").Loop = False

    FlexDMD.LockRenderThread
    FlexDMD.Stage.RemoveAll
    FlexDMD.Stage.AddActor scene
    FlexDMD.Show = True
    FlexDMD.UnlockRenderThread
    FlexDMDActive = True
    On Error Goto 0
End Sub

Dim BigTextActive : BigTextActive = False

Sub ShowBigMessage(msg)
    If Not FlexDMDActive Then Exit Sub
    BigTextActive = True
    On Error Resume Next
    FlexDMD.LockRenderThread
    FlexDMD.Stage.GetGroup("Score").GetLabel("BigText").Text = msg
    FlexDMD.Stage.GetGroup("Score").GetLabel("BigText").SetAlignedPosition 64, 13, FlexDMD_Align_Center
    FlexDMD.Stage.GetGroup("Score").GetLabel("BigText").Visible = True
    FlexDMD.Stage.GetGroup("Score").GetLabel("Line1").Visible = False
    FlexDMD.Stage.GetGroup("Score").GetLabel("Line2").Visible = False
    FlexDMD.UnlockRenderThread
    On Error Goto 0
    BigTextTimer.Interval = 3000
    BigTextTimer.Enabled = True
End Sub

Sub BigTextTimer_Timer()
    BigTextTimer.Enabled = False
    BigTextActive = False
    If Not FlexDMDActive Then Exit Sub
    On Error Resume Next
    FlexDMD.LockRenderThread
    FlexDMD.Stage.GetGroup("Score").GetLabel("BigText").Visible = False
    FlexDMD.Stage.GetGroup("Score").GetLabel("Line1").Visible = True
    FlexDMD.Stage.GetGroup("Score").GetLabel("Line2").Visible = True
    FlexDMD.UnlockRenderThread
    On Error Goto 0
    UpdateDMDScore
End Sub

Sub PlayDMDIntroVideo()
    If Not FlexDMDActive Then Exit Sub
    If DMDVideoActive Then Exit Sub
    If FlexIntroScene Is Nothing Then Exit Sub
    DMDVideoActive = True
    FlexDMD.LockRenderThread
    FlexDMD.Stage.RemoveAll
    FlexDMD.Stage.AddActor FlexIntroScene
    FlexDMD.UnlockRenderThread
    DMDVideoTimer.Interval = 15500
    DMDVideoTimer.Enabled = True
End Sub

Sub StopDMDVideo()
    If Not DMDVideoActive Then Exit Sub
    DMDVideoActive = False
    DMDVideoTimer.Enabled = False
    If Not FlexDMDActive Then Exit Sub
    FlexDMD.LockRenderThread
    FlexDMD.Stage.RemoveAll
    FlexDMD.Stage.AddActor FlexDMD.Stage.GetGroup("Score")
    FlexDMD.UnlockRenderThread
End Sub

Sub DMDVideoTimer_Timer()
    DMDVideoTimer.Enabled = False
    StopDMDVideo
End Sub

Sub ShowMysteryDMD(line1, line2)
'    DMDDisplay.Text = line1 & " " & line2
    If Not FlexDMDActive Then Exit Sub
    On Error Resume Next
    FlexDMD.LockRenderThread
    FlexDMD.Stage.GetGroup("Score").GetLabel("Line1").Font = FontSmall
    FlexDMD.Stage.GetGroup("Score").GetLabel("Line1").Text = line1
    FlexDMD.Stage.GetGroup("Score").GetLabel("Line1").SetAlignedPosition 64, 10, FlexDMD_Align_Center
    FlexDMD.Stage.GetGroup("Score").GetLabel("Line2").Font = FontSmall
    FlexDMD.Stage.GetGroup("Score").GetLabel("Line2").Text = line2
    FlexDMD.Stage.GetGroup("Score").GetLabel("Line2").SetAlignedPosition 64, 23, FlexDMD_Align_Center
    FlexDMD.UnlockRenderThread
    On Error Goto 0
End Sub

Sub UpdateDMDScore()
    If DifficultySelectActive Then Exit Sub
    If MysteryActive Then Exit Sub
    If LeapReady Then Exit Sub
    If TiltActive Then Exit Sub
    If AmbushClearHold Then Exit Sub
    Dim scoreTxt : scoreTxt = FormatNumber(Score, 0, -1, 0, -1)
'    DMDDisplay.Text = scoreTxt
    If Not FlexDMDActive Then Exit Sub
    On Error Resume Next
    FlexDMD.LockRenderThread
    FlexDMD.Stage.GetGroup("Score").GetLabel("Line1").Font = FontMain
    FlexDMD.Stage.GetGroup("Score").GetLabel("Line1").Text = scoreTxt
    FlexDMD.Stage.GetGroup("Score").GetLabel("Line1").SetAlignedPosition 64, 10, FlexDMD_Align_Center
    Dim displayQ : displayQ = RuneQuality
    If displayQ > 4 Then displayQ = 4
    FlexDMD.Stage.GetGroup("Score").GetLabel("Line2").Font = FontTiny
    Dim displayBallNum : displayBallNum = BallNumber
    If displayBallNum > MaxBalls Then displayBallNum = MaxBalls
    FlexDMD.Stage.GetGroup("Score").GetLabel("Line2").Text = "RUNES:+" & displayQ & " KILLS:" & EnemiesKilled & " BALL:" & displayBallNum
    FlexDMD.Stage.GetGroup("Score").GetLabel("Line2").SetAlignedPosition 64, 26, FlexDMD_Align_Center
    FlexDMD.Stage.GetGroup("Score").GetLabel("Line3").Text = ""
    FlexDMD.Stage.GetGroup("Score").GetLabel("Line4").Text = ""
    FlexDMD.UnlockRenderThread
    On Error Goto 0
End Sub

'Sub UpdateDMD(msg)
'    If MysteryActive Then Exit Sub
'    If LeapReady Then Exit Sub
'    If AmbushActive Then
'        AmbushDMDPaused = True
'        AmbushHurryTimer.Enabled = False
'        AmbushResumeTimer.Enabled = False
'        AmbushResumeTimer.Enabled = True
'    End If
'    DMDDisplay.Text = msg
'    If Not FlexDMDActive Then Exit Sub
'    On Error Resume Next
'    FlexDMD.LockRenderThread
'    FlexDMD.Stage.GetGroup("Score").GetLabel("Line1").Font = FontSmall
'    FlexDMD.Stage.GetGroup("Score").GetLabel("Line1").Text = msg
'    FlexDMD.Stage.GetGroup("Score").GetLabel("Line1").SetAlignedPosition 64, 11, FlexDMD_Align_Center
'    If CurrentAct = 6 And PortalActive Then
'        UpdatePortalDMD
'    Else
'        Dim displayQ : displayQ = RuneQuality
'        If displayQ > 4 Then displayQ = 4
'        FlexDMD.Stage.GetGroup("Score").GetLabel("Line2").Font = FontTiny
'        FlexDMD.Stage.GetGroup("Score").GetLabel("Line2").Text = "RUNES:+" & displayQ & " KILLS:" & EnemiesKilled & " BALL:" & BallNumber
'        FlexDMD.Stage.GetGroup("Score").GetLabel("Line2").SetAlignedPosition 64, 26, FlexDMD_Align_Center
'        FlexDMD.Stage.GetGroup("Score").GetLabel("Line3").Text = ""
'        FlexDMD.Stage.GetGroup("Score").GetLabel("Line4").Text = ""
'    End If
'    FlexDMD.UnlockRenderThread
'    On Error Goto 0
'End Sub

Sub UpdateAmbushDMD()
    Dim line1 : line1 = "AMBUSH! " & AmbushKillCount & " / " & AmbushKillsRequired & " KILLS"
    Dim line2 : line2 = "PRIZE: " & FormatNumber(AmbushPrize, 0, -1, 0, -1)
'    DMDDisplay.Text = line1 & " " & line2
    If Not FlexDMDActive Then Exit Sub
    On Error Resume Next
    FlexDMD.LockRenderThread
    FlexDMD.Stage.GetGroup("Score").GetLabel("Line1").Font = FontSmall
    FlexDMD.Stage.GetGroup("Score").GetLabel("Line1").Text = line1
    FlexDMD.Stage.GetGroup("Score").GetLabel("Line1").SetAlignedPosition 64, 10, FlexDMD_Align_Center
    FlexDMD.Stage.GetGroup("Score").GetLabel("Line2").Text = line2
    FlexDMD.Stage.GetGroup("Score").GetLabel("Line2").SetAlignedPosition 64, 23, FlexDMD_Align_Center
    FlexDMD.UnlockRenderThread
    On Error Goto 0
End Sub

Sub UpdateDMD(msg)
    If DifficultySelectActive Then Exit Sub
    If MysteryActive Then Exit Sub
    If LeapReady Then Exit Sub
    If AmbushActive Then
        AmbushDMDPaused = True
        AmbushHurryTimer.Enabled = False
        AmbushResumeTimer.Enabled = False
        AmbushResumeTimer.Enabled = True
    End If
'    DMDDisplay.Text = msg
    If Not FlexDMDActive Then Exit Sub
    On Error Resume Next
    FlexDMD.LockRenderThread
    FlexDMD.Stage.GetGroup("Score").GetLabel("Line1").Font = FontSmall
    FlexDMD.Stage.GetGroup("Score").GetLabel("Line1").Text = msg
    FlexDMD.Stage.GetGroup("Score").GetLabel("Line1").SetAlignedPosition 64, 11, FlexDMD_Align_Center
    FlexDMD.Stage.GetGroup("Score").GetLabel("Line2").Font = FontTiny
    Dim displayQ : displayQ = RuneQuality
    If displayQ > 4 Then displayQ = 4
   Dim displayBallNum : displayBallNum = BallNumber
    If displayBallNum > MaxBalls Then displayBallNum = MaxBalls
    FlexDMD.Stage.GetGroup("Score").GetLabel("Line2").Text = "RUNES:+" & displayQ & " KILLS:" & EnemiesKilled & " BALL:" & displayBallNum
    FlexDMD.Stage.GetGroup("Score").GetLabel("Line2").SetAlignedPosition 64, 26, FlexDMD_Align_Center
    FlexDMD.Stage.GetGroup("Score").GetLabel("Line3").Text = ""
    FlexDMD.Stage.GetGroup("Score").GetLabel("Line4").Text = ""
    FlexDMD.UnlockRenderThread
    On Error Goto 0
End Sub
Sub UpdateDMD2(line1, line2)
    If DifficultySelectActive Then Exit Sub
    If MysteryActive Then Exit Sub
    If LeapReady Then Exit Sub
    If AmbushActive Then
        AmbushDMDPaused = True
        AmbushHurryTimer.Enabled = False
        AmbushResumeTimer.Enabled = False
        AmbushResumeTimer.Enabled = True
    End If
'    DMDDisplay.Text = line1 & " " & line2
    If Not FlexDMDActive Then Exit Sub
    On Error Resume Next
    FlexDMD.LockRenderThread
    FlexDMD.Stage.GetGroup("Score").GetLabel("Line1").Font = FontSmall
    FlexDMD.Stage.GetGroup("Score").GetLabel("Line1").Text = line1
    FlexDMD.Stage.GetGroup("Score").GetLabel("Line1").SetAlignedPosition 64, 10, FlexDMD_Align_Center
    FlexDMD.Stage.GetGroup("Score").GetLabel("Line2").Font = FontSmall
    FlexDMD.Stage.GetGroup("Score").GetLabel("Line2").Text = line2
    FlexDMD.Stage.GetGroup("Score").GetLabel("Line2").SetAlignedPosition 64, 23, FlexDMD_Align_Center
    FlexDMD.UnlockRenderThread
    On Error Goto 0
End Sub

'BOSS AND TRAVEL/ACTS

Const ACT_ROGUE        = 1
Const ACT_LUT_GHOLEIN  = 2
Const ACT_KURAST       = 3
Const ACT_PANDEMONIUM  = 4
Const ACT_HARROGATH    = 5
Const ACT_COW_LEVEL    = 6

Const VRWallBright = 1.0     ' restore target for VR_d_brickwall
Const VRFloorBright = 0.7    ' restore target for VR_d_floors
Const VRDarkLevel = 0.1        ' snap-dark level; raise toward 0.1 if pure black is too much
Dim BossDimStep : BossDimStep = 0
Dim BossDimFast : BossDimFast = False

Sub BossRoomHitBlink()
    If RenderingMode <> 2 Then Exit Sub
    VR_d_brickwall.BlendDisableLighting = VRDarkLevel
    VR_d_floors.BlendDisableLighting = VRDarkLevel
    BossDimStep = 0
    BossDimFast = True
    BossRoomDimTimer.Interval = 30
    BossRoomDimTimer.Enabled = True
End Sub

Sub BossRoomSnapDark()
    If RenderingMode <> 2 Then Exit Sub
    VR_d_brickwall.BlendDisableLighting = VRDarkLevel
    VR_d_floors.BlendDisableLighting = VRDarkLevel
    BossDimStep = 0
    BossDimFast = False
    BossRoomDimTimer.Interval = 30
    BossRoomDimTimer.Enabled = True
End Sub

Sub BossRoomDimTimer_Timer()
    BossDimStep = BossDimStep + 1
    Dim prog
    If BossDimFast Then
        prog = BossDimStep * 0.1          ' no hold, snap back over ~300ms
    Else
        If BossDimStep <= 15 Then Exit Sub ' hold dark ~450ms
        prog = (BossDimStep - 15) * 0.01   ' fade up over ~3s
    End If
    If prog >= 1 Then
        prog = 1
        BossRoomDimTimer.Enabled = False
    End If
    VR_d_brickwall.BlendDisableLighting = VRWallBright * prog
    VR_d_floors.BlendDisableLighting = VRFloorBright * prog
End Sub

Dim BossCritActive : BossCritActive = False   ' current fight
Dim BossCritQueued : BossCritQueued = False   ' queued for next fight
Dim FootstepStep : FootstepStep = 0
Dim FootstepSet : FootstepSet = 0
Dim BossScaling : BossScaling = 1.0
Const BOSS_BLOODRAVEN = 0
Const BOSS_TREEHEAD   = 1
Const BOSS_GRISWOLD   = 2
Const BOSS_COUNTESS   = 3
Const BOSS_SMITH    = 4
Const BOSS_COWKING  = 5
Dim BossHealth : BossHealth = 1
Dim TravelProgress : TravelProgress = 0
Dim EventIndex : EventIndex = 0
Dim BossFightActive : BossFightActive = False
Dim CritCycleActive : CritCycleActive = False
Dim CritCycleStep   : CritCycleStep   = 0
Dim CritCharged     : CritCharged     = 0
Dim CritExpireCountdown : CritExpireCountdown = 0
Dim BossEventType : BossEventType = -1
Dim BossHP_BloodRaven : BossHP_BloodRaven = 8
Dim BossHP_Treehead : BossHP_Treehead = 8
Dim BossHP_Griswold : BossHP_Griswold = 8
Dim BossHP_Countess : BossHP_Countess = 8
Dim BossHP_Smith : BossHP_Smith = 8
Dim BossHP_CowKing : BossHP_CowKing = 8
Dim CurrentBossHP : CurrentBossHP = 0
Dim TravelActive : TravelActive = True
Dim TravelSoundTimer_Active : TravelSoundTimer_Active = False
Dim BossLootGoldAmount : BossLootGoldAmount = 0
Dim BossLootGearDesc : BossLootGearDesc = ""
Dim BossMaxHP : BossMaxHP = 0
Dim BossHPPulseStep : BossHPPulseStep = 0

Dim BossHitCount : BossHitCount = 0

Dim SplatterOrder(7)

Sub ShuffleSplatters()
    Dim i, j, tmp
    For i = 0 To 7 : SplatterOrder(i) = i + 1 : Next
    For i = 7 To 1 Step -1
        j = Int(Rnd * (i + 1))
        tmp = SplatterOrder(i)
        SplatterOrder(i) = SplatterOrder(j)
        SplatterOrder(j) = tmp
    Next
End Sub

Sub ShowBossHitSplatter()
    BossHitCount = BossHitCount + 1
    If BossHitCount > 8 Then Exit Sub
    Select Case SplatterOrder(BossHitCount - 1)
        Case 1 : Flasher001.Visible = True
        Case 2 : Flasher002.Visible = True
        Case 3 : Flasher003.Visible = True
        Case 4 : Flasher004.Visible = True
        Case 5 : Flasher005.Visible = True
        Case 6 : Flasher006.Visible = True
        Case 7 : Flasher007.Visible = True
        Case 8 : Flasher008.Visible = True
    End Select
End Sub

Sub ActCompleteTimer_Timer()
    ActCompleteTimer.Enabled = False
    Select Case BossEventType
        Case BOSS_BLOODRAVEN : SetMusicVolume 0.4
            UpdateDMD2 "ACT I COMPLETE!", "ROGUE ENCAMPMENT SAVED"
        Case BOSS_TREEHEAD
            UpdateDMD2 "ACT II COMPLETE!", "LUT GHOLEIN SAVED"
        Case BOSS_GRISWOLD
            UpdateDMD2 "ACT III COMPLETE!", "KURAST DOCKS SAVED"
        Case BOSS_COUNTESS
            UpdateDMD2 "ACT IV COMPLETE!", "FORTRESS SAVED"
        Case BOSS_SMITH
            UpdateDMD2 "ACT V COMPLETE!", "HARROGATH SAVED"
        Case BOSS_COWKING
            UpdateDMD2 "SECRET LEVEL!", "MOO!"
    End Select
End Sub

Sub BossActTimer_Timer()
    BossActTimer.Enabled = False
    If Not GameActive Then Exit Sub      ' don't fire act music if game ended (e.g. last-ball boss drain)
    If CurrentAct < 6 Then CurrentAct = CurrentAct + 1 End If
    Dim actMsg
    Select Case BossEventType
        Case BOSS_BLOODRAVEN : actMsg = "ACT I COMPLETE!"  : PlaySong "lut gholein"
        Case BOSS_TREEHEAD   : actMsg = "ACT II COMPLETE!" : PlaySong "kurast docks"
        Case BOSS_GRISWOLD   : actMsg = "ACT III COMPLETE!": PlaySong "pandemoniumfortress"
        Case BOSS_COUNTESS   : actMsg = "ACT IV COMPLETE!" : PlaySong "harragoth"
        Case BOSS_SMITH      : actMsg = "ACT V COMPLETE!"  : PlaySong "tristram"
        Case BOSS_COWKING    : actMsg = "COW KING DEFEATED!" : PlaySong "tristram"
    End Select
    ShowMessage actMsg
End Sub
'    If CurrentAct = 6 And BossEventType = BOSS_SMITH Then
'        CowNeutralTimer.Interval = 30000
'        CowNeutralTimer.Enabled = True
'        Dim startTime : startTime = 60000 + (PortalSurge * 1000)
'        If startTime > PortalCap() Then startTime = PortalCap()
'        PortalTime = startTime
'        PortalSurge = 0
'        PortalActive = True
'        PortalCollapsed = False
'        PortalTimer.Interval = 250
'        PortalTimer.Enabled = True
'        ShowMysteryDMD "WIZARD MODE!", "PORTAL OPEN!"
'    End If

Sub SetTravelLights(isOn)
    If isOn Then
        ' Lights 1, 2, 4 — orange blinking (ambush risk)
        QuestLight1.Color = RGB(255, 100, 0) : QuestLight1.ColorFull = RGB(255, 140, 0)
        QuestLight1.BlinkInterval = 400 : QuestLight1.State = 2
        QuestLight2.Color = RGB(255, 100, 0) : QuestLight2.ColorFull = RGB(255, 140, 0)
        QuestLight2.BlinkInterval = 400 : QuestLight2.State = 2
        If Not MysteryReady Then
            QuestLight4.Color = RGB(255, 100, 0) : QuestLight4.ColorFull = RGB(255, 140, 0)
            QuestLight4.BlinkInterval = 400 : QuestLight4.State = 2
        End If
    Else
        QuestLight1.State = 0
        QuestLight2.State = 0
        QuestLight4.State = 0
        SetRampLight
    End If
End Sub

Dim CainCalloutIndex : CainCalloutIndex = 0
Dim CainIntroPending : CainIntroPending = False

Sub CainCalloutTimer_Timer()
    CainCalloutTimer.Enabled = False
    Select Case CainCalloutIndex
        Case 0 : PlayCallout "CainAndarielNew", 5000
        Case 1 : PlayCallout "CainDuriel", 5000
        Case 2 : PlayCallout "CainMephisto", 5000
        Case 3 : PlayCallout "CainDiablo", 5000
        Case 4 : PlayCallout "CainBaal", 5000
        Case 5 : PlayCallout "CainCowKing", 5000
    End Select
End Sub

Sub TravelHit()
    If Not TravelActive Then Exit Sub
    If BossFightActive Then Exit Sub
    TravelProgress = TravelProgress + 1
    If TravelProgress = 1 And EventIndex = 0 Then
        CainIntroPending = True
    End If

    FootstepStep = 0
    FootstepSet = Int(Rnd * 3)
    FootstepTimer.Enabled = False
    FootstepTimer.Enabled = True

    TravelAmbientTimer.Enabled = False
    TravelAmbientTimer.Enabled = True

    If TravelProgress >= 5 Then
        TravelProgress = 0
        FireNextEvent
    End If
End Sub

Sub CheckCainIntro()
    If Not CainIntroPending Then Exit Sub
    If AmbushActive Then Exit Sub
    CainIntroPending = False
    CainCalloutIndex = 0
    CainCalloutTimer.Interval = 800
    CainCalloutTimer.Enabled = True
End Sub

Sub FootstepTimer_Timer()
    FootstepStep = FootstepStep + 1
    Select Case FootstepSet
        Case 0
            Select Case FootstepStep
                Case 1 : PlaySound "MedWood1", 0, 1 * DuckVolume
                Case 2 : PlaySound "MedWood2", 0, 1 * DuckVolume
                Case 3 : PlaySound "MedWood3", 0, 1 * DuckVolume
                Case 4 : PlaySound "MedWood4", 0, 1 * DuckVolume
            End Select
        Case 1
            Select Case FootstepStep
                Case 1 : PlaySound "MedIStoneRun1", 0, 1 * DuckVolume
                Case 2 : PlaySound "MedIStoneRun2", 0, 1 * DuckVolume
                Case 3 : PlaySound "MedIStoneRun3", 0, 1 * DuckVolume
                Case 4 : PlaySound "MedIStoneRun4", 0, 1 * DuckVolume
            End Select
        Case 2
            Select Case FootstepStep
                Case 1 : PlaySound "HeavyDirtRun1", 0, 1 * DuckVolume
                Case 2 : PlaySound "HeavyDirtRun2", 0, 1 * DuckVolume
                Case 3 : PlaySound "HeavyDirtRun3", 0, 1 * DuckVolume
                Case 4 : PlaySound "HeavyDirtRun4", 0, 1 * DuckVolume
            End Select
    End Select
    If FootstepStep >= 4 Then
        FootstepTimer.Enabled = False
    End If
End Sub

Sub TravelAmbientTimer_Timer()
  If Not TravelActive Then Exit Sub
    TravelAmbientTimer.Enabled = False
    Dim ambRoll : ambRoll = Int(Rnd * 10)
    Select Case ambRoll
        Case 0 : PlaySound "birdie_night_14", 0, 1 * DuckVolume
Case 1 : PlaySound "cavedrip" & (Int(Rnd * 8) + 1), 0, 1 * DuckVolume
Case 2 : PlaySound "raven" & (Int(Rnd * 5) + 1), 0, 1 * DuckVolume
Case 3 : PlaySound "singlecricket", 0, 1 * DuckVolume
Case 4 : PlaySound "thunder_norm_" & (Array(1,2,5))(Int(Rnd * 3)), 0, 1 * DuckVolume
Case 5 : PlaySound "town2birdie" & (Int(Rnd * 5) + 1), 0, 1 * DuckVolume
Case 6 : PlaySound "wind" & (Int(Rnd * 9) + 1), 0, 1 * DuckVolume
Case 8 : PlaySound "raven" & (Int(Rnd * 5) + 1), 0, 1 * DuckVolume
    End Select
End Sub

Sub FireNextEvent()
    SetTravelLights False
    Select Case EventIndex
        Case 0 : StartBossFight BOSS_BLOODRAVEN
        Case 1 : StartBossFight BOSS_TREEHEAD
        Case 2 : StartBossFight BOSS_GRISWOLD
        Case 3 : StartBossFight BOSS_COUNTESS
        Case 4 : StartBossFight BOSS_SMITH
        Case 5 : StartBossFight BOSS_COWKING
        Case Else : TravelActive = False
    End Select
End Sub

Sub UpdateBossHPLights()
    Dim i
    For i = 0 To BossHP.Count - 1
        If i < BossMaxHP Then
            If i < BossHealth Then
                BossHP.Item(i).Color = RGB(200, 0, 0)
                BossHP.Item(i).ColorFull = RGB(200, 0, 0)
                BossHP.Item(i).State = 1
            Else
                BossHP.Item(i).State = 0
            End If
        Else
            BossHP.Item(i).State = 0
        End If
    Next
End Sub

Sub BossHPPulseTimer_Timer()
    BossHPPulseStep = BossHPPulseStep + 1
    Dim pulseIntensity
    If BossHPPulseStep <= 8 Then
        ' Ramp up bright white-red flash
        pulseIntensity = 5 + (BossHPPulseStep * 3)
        Dim pulse1
        For pulse1 = 0 To 7
            If BossHP.Item(pulse1).State = 1 Then
                BossHP.Item(pulse1).Color = RGB(255, 50, 50)
                BossHP.Item(pulse1).ColorFull = RGB(255, 50, 50)
                BossHP.Item(pulse1).Intensity = pulseIntensity
            End If
        Next
    ElseIf BossHPPulseStep <= 16 Then
        ' Fade back to deep red
        pulseIntensity = 29 - ((BossHPPulseStep - 8) * 3)
        If pulseIntensity < 10 Then pulseIntensity = 10
        Dim pulse2
        For pulse2 = 0 To 7
            If BossHP.Item(pulse2).State = 1 Then
                BossHP.Item(pulse2).Color = RGB(200, 0, 0)
                BossHP.Item(pulse2).ColorFull = RGB(200, 0, 0)
                BossHP.Item(pulse2).Intensity = pulseIntensity
            End If
        Next
    Else
        ' Settle at steady deep red
        BossHPPulseTimer.Enabled = False
        Dim pulse3
        For pulse3 = 0 To 7
            If BossHP.Item(pulse3).State = 1 Then
                BossHP.Item(pulse3).Color = RGB(200, 0, 0)
                BossHP.Item(pulse3).ColorFull = RGB(200, 0, 0)
                BossHP.Item(pulse3).Intensity = 10
            End If
        Next
    End If
End Sub

Function BossUnlocksCharacter(bossType)
    Select Case bossType
        Case BOSS_BLOODRAVEN : BossUnlocksCharacter = "Ama"
        Case BOSS_TREEHEAD   : BossUnlocksCharacter = "Dru"
        Case BOSS_GRISWOLD   : BossUnlocksCharacter = "Necro"
        Case BOSS_COUNTESS   : BossUnlocksCharacter = "Ass"
        Case BOSS_SMITH      : BossUnlocksCharacter = "Pal"
        Case BOSS_COWKING    : BossUnlocksCharacter = "Sorc"
        Case Else            : BossUnlocksCharacter = ""
    End Select
End Function

Function CritWindowDuration()
    CritWindowDuration = 8 + CountEquippedUniques()
End Function



Sub StartBossFight(bossType)
    BossFightActive = True
  StartTurntables
  ShuffleSplatters
    BossEventType = bossType
    BossHealth = GetBossHP(bossType)
    CurrentBossHP = BossHealth
    BossMaxHP = BossHealth
    If BossCritQueued Then
        BossCritQueued = False
        BossCritActive = True
        ShowMessage "CRIT BOSS! 2x DAMAGE!"
    Else
        BossCritActive = False
    End If
    CapBallLight.Color = RGB(180, 0, 0)
    CapBallLight.ColorFull = RGB(180, 0, 0)
    CapBallLight.State = 2
    CapBallLight.BlinkInterval = 200
    StartTauntFlash 3000
    BossRoomSnapDark
    Select Case bossType
        Case BOSS_BLOODRAVEN
            ShowMessage "ANDARIEL!"
            Select Case Int(Rnd * 3)
                Case 0 : PlayCallout "andarieltaunt1", 3000
                Case 1 : PlayCallout "andarieltaunt2", 3000
                Case 2 : PlayCallout "andarieltaunt3", 3000
            End Select
        Case BOSS_TREEHEAD
            ShowMessage "DURIEL!"
            Select Case Int(Rnd * 2)
                Case 0 : PlayCallout "durieltaunt1", 3000
                Case 1 : PlayCallout "durieltaunt2", 3000
            End Select
        Case BOSS_GRISWOLD
            ShowMessage "MEPHISTO!"
            Select Case Int(Rnd * 2)
                Case 0 : PlayCallout "mephistotaunt1", 3000
                Case 1 : PlayCallout "mephistotaunt2", 3000
            End Select
        Case BOSS_COUNTESS
            ShowMessage "DIABLO!"
            StopSound CurrentSong
            PlaySong "diablofight"
            Select Case Int(Rnd * 2)
                Case 0 : PlayCallout "diablotaunt1", 3000
                Case 1 : PlayCallout "diablotaunt2", 3000
            End Select
        Case BOSS_SMITH
            ShowMessage "BAAL!"
            PlayCallout "baaltaunt", 3000
        Case BOSS_COWKING
            ShowMessage "THE COW KING!"
            Select Case Int(Rnd * 5)
                Case 0 : PlayCallout "cow_neutral1", 1000
                Case 1 : PlayCallout "cow_neutral2", 3000
                Case 2 : PlayCallout "cow_neutral3", 3000
                Case 3 : PlayCallout "cow_neutral4", 1000
                Case 4 : PlayCallout "cow_neutral5", 1000
            End Select
    End Select
    UpdateBossHPLights
    StartGIEvent GI_MODE_BOSS
    SetFlasherColor 3, 255, 0, 0
    SetFlasherColor 4, 255, 0, 0
    FireFlasher 3
    FireFlasher 4
  CritCharged = 0
    StartCritCycle
    If DifficultyLevel = 2 Then
        BossRegenTimer.Interval = 60000
        BossRegenTimer.Enabled = True
    End If
End Sub

Sub StartCritCycle()
    If Not BossFightActive Then Exit Sub
    CritCycleActive = True
    CritCycleStep = 0
    SetCritCycleLight CritCycleStep
    If AllUniquesEquipped() Then
        CritCycleTimer.Enabled = False
    Else
        CritCycleTimer.Interval = CritCycleInterval()
        CritCycleTimer.Enabled = True
    End If
End Sub

Sub StopCritCycle()
    CritCycleActive = False
    CritCycleTimer.Enabled = False
    CritExpireTimer.Enabled = False
    CritPulseOffTimer.Enabled = False
    QuestLight1.State = 0
    QuestLight2.State = 0
    If Not MysteryReady Then QuestLight4.State = 0
    SetRampLight
End Sub

Sub SetCritCycleLight(step)
    QuestLight1.State = 0
    QuestLight2.State = 0
    If Not MysteryReady Then QuestLight4.State = 0
    Select Case step
        Case 0 : QuestLight1.Color = RGB(255,0,0) : QuestLight1.ColorFull = RGB(255,0,0) : QuestLight1.State = 1
        Case 1 : QuestLight2.Color = RGB(255,0,0) : QuestLight2.ColorFull = RGB(255,0,0) : QuestLight2.State = 1
        Case 2 : If Not MysteryReady Then QuestLight4.Color = RGB(255,0,0) : QuestLight4.ColorFull = RGB(255,0,0) : QuestLight4.State = 1
    End Select
    SetRampLight
End Sub
Sub CritCycleTimer_Timer()
    If Not CritCycleActive Then CritCycleTimer.Enabled = False : Exit Sub
    If AllUniquesEquipped() Then CritCycleTimer.Enabled = False : Exit Sub
    CritCycleStep = (CritCycleStep + 1) Mod 3
    SetCritCycleLight CritCycleStep
    CritCycleTimer.Interval = CritCycleInterval()
    CritCycleTimer.Enabled = True
End Sub

Sub OnCritLaneHit(laneIndex)
    If Not CritCycleActive Then Exit Sub
    If laneIndex <> CritCycleStep Then Exit Sub
    If CritCharged >= 1 Then Exit Sub
    CritCharged = CritCharged + 1
    PlayCallout "mystery_critboss", 2000
    ShowMessage "CRIT CHARGED!"
    StopCritCycle
    CritExpireCountdown = CritWindowDuration() + 1
    CritExpireTimer.Interval = 1000
    CritExpireTimer.Enabled = True
End Sub

' ── AFTER ───────────────────────────────────────────
Sub CritExpireTimer_Timer()
    CritExpireCountdown = CritExpireCountdown - 1
    ' QuestLight1/2/4 are the crit lane arrows during boss fight — keep them
    ' dark while the crit window is open. BossHPPulse handles the visual feedback.
    BossHPPulseStep = 0
    BossHPPulseTimer.Enabled = True
    If CritExpireCountdown <= 0 Then
        CritCharged = 0
        ShowMessage "CRIT LOST!"
        CritExpireTimer.Enabled = False
        CritPulseOffTimer.Enabled = False
        QuestLight1.State = 0
        QuestLight2.State = 0
        QuestLight4.State = 0
        StartCritCycle
    End If
End Sub

Sub CritPulseOffTimer_Timer()
    CritPulseOffTimer.Enabled = False
    If CritCharged > 0 Then
        QuestLight1.State = 0
    QuestLight2.State = 0
    QuestLight4.State = 0
    End If
End Sub

Sub CowNeutralTimer_Timer()
    If Not GameActive Or CurrentAct <> 6 Then
        CowNeutralTimer.Enabled = False
        Exit Sub
    End If
    Select Case Int(Rnd * 5)
        Case 0 : PlaySound "cow_neutral1", 0, 0.5
        Case 1 : PlaySound "cow_neutral2", 0, 0.5
        Case 2 : PlaySound "cow_neutral3", 0, 0.5
        Case 3 : PlaySound "cow_neutral4", 0, 0.5
        Case 4 : PlaySound "cow_neutral5", 0, 0.5
    End Select
    CowNeutralTimer.Interval = 25000 + Int(Rnd * 15000)
    CowNeutralTimer.Enabled = True
End Sub

Function GetBossHP(bossType)
    Dim base
    Select Case bossType
        Case BOSS_BLOODRAVEN : base = BossHP_BloodRaven
        Case BOSS_TREEHEAD   : base = BossHP_Treehead
        Case BOSS_GRISWOLD   : base = BossHP_Griswold
        Case BOSS_COUNTESS   : base = BossHP_Countess
        Case BOSS_SMITH      : base = BossHP_Smith
        Case BOSS_COWKING    : base = BossHP_CowKing
        Case Else            : base = 2
    End Select
    GetBossHP = base
    If GetBossHP > 8 Then GetBossHP = 8
End Function

Sub HitBoss()
    If Not BossFightActive Then Exit Sub
  ShowBossHitSplatter
    Dim dmg : dmg = 1
    If BossCritActive Then
        dmg = 2
        BossCritActive = False
    End If
    If CritCharged > 0 Then
        dmg = dmg * 2
        CritCharged = 0
        BallCritCount = BallCritCount + 1
        CritExpireTimer.Enabled = False
        ShowMessage "CRITICAL HIT!"
        StartCritCycle
    End If
    BossHealth = BossHealth - dmg
    'BossRoomHitBlink                    ' room dimming removed from boss hits
    If DifficultyLevel = 2 And BossFightActive Then
        BossRegenTimer.Enabled = False
        BossRegenTimer.Interval = 60000
        BossRegenTimer.Enabled = True
    End If

    Select Case BossEventType
        Case BOSS_BLOODRAVEN
            Select Case Int(Rnd * 4)
                Case 0 : PlaySound "a_gethit1", 0, 1
                Case 1 : PlaySound "a_gethit2", 0, 1
                Case 2 : PlaySound "a_gethit3", 0, 1
                Case 3 : PlaySound "a_gethit4", 0, 1
            End Select
        Case BOSS_TREEHEAD
            Select Case Int(Rnd * 4)
                Case 0 : PlaySound "duriel_gethit1", 0, 1
                Case 1 : PlaySound "duriel_gethit2", 0, 1
                Case 2 : PlaySound "duriel_gethit3", 0, 1
        Case 3 : PlaySound "duriel_gethit4", 0, 1
            End Select
        Case BOSS_GRISWOLD
            Select Case Int(Rnd * 4)
                Case 0 : PlaySound "meph_gethit1", 0, 1
                Case 1 : PlaySound "meph_gethit2", 0, 1
                Case 2 : PlaySound "meph_gethit3", 0, 1
                Case 3 : PlaySound "meph_gethit4", 0, 1
            End Select
        Case BOSS_COUNTESS
            Select Case Int(Rnd * 6)
                Case 0 : PlaySound "diablo_gethit1", 0, 1
                Case 1 : PlaySound "diablo_gethit2", 0, 1
                Case 2 : PlaySound "diablo_gethit3", 0, 1
                Case 3 : PlaySound "diablo_gethit4", 0, 1
                Case 4 : PlaySound "diablo_gethit5", 0, 1
                Case 5 : PlaySound "diablo_gethit6", 0, 1
            End Select
        Case BOSS_SMITH
            Select Case Int(Rnd * 4)
                Case 0 : PlaySound "baal_gethit1", 0, 1
                Case 1 : PlaySound "baal_gethit2", 0, 1
                Case 2 : PlaySound "baal_gethit3", 0, 1
                Case 3 : PlaySound "baal_gethit4", 0, 1
            End Select
    Case BOSS_COWKING
            Select Case Int(Rnd * 4)
                Case 0 : PlaySound "cow_gethit1", 0, 1
                Case 1 : PlaySound "cow_gethit2", 0, 1
                Case 2 : PlaySound "cow_gethit3", 0, 1
                Case 3 : PlaySound "cow_gethit4", 0, 1
            End Select
    End Select
CapBallLight.Color = RGB(180, 0, 0)
    CapBallLight.ColorFull = RGB(180, 0, 0)
    CapBallLight.BlinkInterval = 100
    UpdateBossHPLights
    BossHPPulseStep = 0
    BossHPPulseTimer.Enabled = False
    BossHPPulseTimer.Enabled = True
    SetFlasherColor 1, 255, 0, 0 : SetFlasherColor 2, 255, 0, 0
    SetFlasherColor 3, 255, 0, 0 : SetFlasherColor 4, 255, 0, 0
    SetFlasherColor 5, 255, 0, 0 : SetFlasherColor 6, 255, 0, 0
    FireAllFlashers
    BossHitGITimer.Enabled = False
    BossHitGITimer.Enabled = True
    If BossHealth <= 0 Then
        DefeatBoss
    End If
End Sub

Sub BossHitGITimer_Timer()
    BossHitGITimer.Enabled = False
    DimAllFlashers
End Sub

Sub BossRegenTimer_Timer()
    If Not BossFightActive Or DifficultyLevel <> 2 Then
        BossRegenTimer.Enabled = False
        Exit Sub
    End If
    If BossHealth < BossMaxHP Then
        BossHealth = BossHealth + 1
        UpdateBossHPLights
        ShowMessage "BOSS REGENERATES!"
        If BossHitCount > 0 Then
            Select Case SplatterOrder(BossHitCount - 1)
                Case 1 : Flasher001.Visible = False
                Case 2 : Flasher002.Visible = False
                Case 3 : Flasher003.Visible = False
                Case 4 : Flasher004.Visible = False
                Case 5 : Flasher005.Visible = False
                Case 6 : Flasher006.Visible = False
                Case 7 : Flasher007.Visible = False
                Case 8 : Flasher008.Visible = False
            End Select
            BossHitCount = BossHitCount - 1
        End If
        SetFlasherColor 1, 0, 255, 80
        SetFlasherColor 2, 0, 255, 80
        SetFlasherColor 3, 0, 255, 80
        SetFlasherColor 4, 0, 255, 80
        SetFlasherColor 5, 0, 255, 80
        SetFlasherColor 6, 0, 255, 80
        FireAllFlashers
        Dim regenIdx : regenIdx = BossHealth - 1
        If regenIdx >= 0 And regenIdx <= 7 Then
            BossHP.Item(regenIdx).Color = RGB(0, 255, 80)
            BossHP.Item(regenIdx).ColorFull = RGB(0, 255, 80)
        End If
        BossRegenFlashTimer.Interval = 300
        BossRegenFlashTimer.Enabled = True
    End If
End Sub

Sub BossRegenFlashTimer_Timer()
    BossRegenFlashTimer.Enabled = False
    DimAllFlashers
    SetFlasherColor 1, 255, 0, 0
    SetFlasherColor 2, 255, 0, 0
    SetFlasherColor 3, 255, 0, 0
    SetFlasherColor 4, 255, 0, 0
    SetFlasherColor 5, 255, 0, 0
    SetFlasherColor 6, 255, 0, 0
    UpdateBossHPLights
End Sub

Sub DefeatBoss()
    BossFightActive = False
    BossCritActive = False
  StopTurntables
    CapBallLight.State = 0
    Dim i : For i = 0 To 7 : BossHP.Item(i).State = 0 : Next
    StartGIEvent GI_MODE_BOSSWIN
    BossHPPulseTimer.Enabled = False
    SetFlasherColor 1, 255, 255, 200
    SetFlasherColor 2, 255, 255, 200
    SetFlasherColor 3, 255, 255, 200
    SetFlasherColor 4, 255, 255, 200
    FireAllFlashers
  Flasher001.Visible = False : Flasher002.Visible = False
    Flasher003.Visible = False : Flasher004.Visible = False
    Flasher005.Visible = False : Flasher006.Visible = False
    Flasher007.Visible = False : Flasher008.Visible = False
    BossHitCount = 0

Select Case BossEventType
        Case BOSS_BLOODRAVEN
            PlaySound "a_death", 0, 1 * DuckVolume
            ShowMessage "ANDARIEL SLAIN!"
        Case BOSS_TREEHEAD
            PlaySound "duriel_death", 0, 1 * DuckVolume
            ShowMessage "DURIEL SLAIN!"
        Case BOSS_GRISWOLD
            PlaySound "meph_death", 0, 1 * DuckVolume
            ShowMessage "MEPHISTO SLAIN!"
        Case BOSS_COUNTESS
            PlaySound "diablo_death", 0, 1 * DuckVolume
            ShowMessage "DIABLO SLAIN!"
        Case BOSS_SMITH
            PlaySound "baal_death", 0, 1 * DuckVolume
            ShowMessage "BAAL SLAIN!"
            BaalDeathFlasherTimer.Interval = 200
            BaalDeathFlasherTimer.Enabled = True
        Case BOSS_COWKING
            Select Case Int(Rnd * 5)
                Case 0 : PlaySound "cow_death1", 0, 1 * DuckVolume
                Case 1 : PlaySound "cow_death2", 0, 1 * DuckVolume
                Case 2 : PlaySound "cow_death3", 0, 1 * DuckVolume
                Case 3 : PlaySound "cow_death4", 0, 1 * DuckVolume
                Case 4 : PlaySound "cow_death5", 0, 1 * DuckVolume
            End Select
            ShowMessage "COW KING SLAIN!"
            CowNeutralTimer.Enabled = False
    End Select

    Dim jackpot
    Select Case BossEventType
        Case BOSS_BLOODRAVEN : jackpot = 10000000
        Case BOSS_TREEHEAD   : jackpot = 20000000
        Case BOSS_GRISWOLD   : jackpot = 30000000
        Case BOSS_COUNTESS   : jackpot = 40000000
        Case BOSS_SMITH      : jackpot = 50000000
        Case BOSS_COWKING
            CowKingKillCount = CowKingKillCount + 1
            jackpot = 75000000 + ((CowKingKillCount - 1) * 25000000)
    End Select

    AddScore jackpot
    UpdateDMD2 "BOSS JACKPOT!", "+" & FormatNumber(jackpot, 0, -1, 0, -1)
'    If CurrentAct = 6 And PortalActive Then
'        PortalTime = PortalTime + 20000
'        If PortalTime > PortalCap() Then PortalTime = PortalCap()
'    End If

    Select Case BossEventType
        Case BOSS_BLOODRAVEN : BossScaling = 1.1
        Case BOSS_TREEHEAD   : BossScaling = 1.2
        Case BOSS_GRISWOLD   : BossScaling = 1.35
        Case BOSS_COUNTESS   : BossScaling = 1.5
        Case BOSS_SMITH      : BossScaling = 2.0
        Case BOSS_COWKING    : BossScaling = 2.0
    End Select

    Dim unlockChar : unlockChar = BossUnlocksCharacter(BossEventType)
    If unlockChar <> "" Then UnlockPartyMember unlockChar

    AssignBumperStats 0, True
    AssignBumperStats 1, False
    AssignBumperStats 2, False
    AssignBumperStats 3, False
    AssignBumperStats 4, False

  EventIndex = EventIndex + 1
    AmbushHappenedThisAct = False
    TravelProgress = 0
    StopCritCycle
    CritCharged = 0
    CritExpireTimer.Enabled = False
    CritPulseOffTimer.Enabled = False
    BossCritActive = False
    BossRegenTimer.Enabled = False
    SetTravelLights True
    CainCalloutIndex = EventIndex
    BossActTimer.Interval = 800
    BossActTimer.Enabled = True
    CainCalloutTimer.Interval = 6000
    CainCalloutTimer.Enabled = True
    If BossEventType = BOSS_COWKING Then
        ' Stay in Act 6 — reset EventIndex back to 5 so Cow King can be fought again
        EventIndex = 5
        TravelActive = True
'        PortalTimer.Enabled = True
    End If
    AmbushBossCount = AmbushBossCount + 1
    AmbushKillsRequired = 10
End Sub

Sub RollBossLoot()
    Dim prop : prop = RollGearProperty()
    LootPending(0) = LOOT_GEAR
    LootValue(0) = prop
    LootTier(0) = GEAR_UNIQUE
    LootActive(0) = True
    BossLootGearDesc = GearPropertyName(prop, GEAR_UNIQUE)
End Sub

Sub BossLootDisplayTimer_Timer()
    BossLootDisplayTimer.Enabled = False
    If FlexDMDActive Then
        FlexDMD.LockRenderThread
        FlexDMD.Stage.GetGroup("Score").GetLabel("Line1").Text = FormatNumber(BossLootGoldAmount, 0, -1, 0, -1)
        FlexDMD.Stage.GetGroup("Score").GetLabel("Line2").Text = "GOLD REWARD!"
        FlexDMD.UnlockRenderThread
    End If
'    DMDDisplay.Text = "GOLD: " & FormatNumber(BossLootGoldAmount, 0, -1, 0, -1)
    BossLootGearTimer.Enabled = True
End Sub

Sub BossLootGearTimer_Timer()
    BossLootGearTimer.Enabled = False
    If FlexDMDActive Then
        FlexDMD.LockRenderThread
        FlexDMD.Stage.GetGroup("Score").GetLabel("Line1").Text = BossLootGearDesc
        FlexDMD.Stage.GetGroup("Score").GetLabel("Line2").Text = "SHOOT CHEST!"
        FlexDMD.UnlockRenderThread
    End If
'    DMDDisplay.Text = BossLootGearDesc
    MsgQueueTimer.Enabled = False
    MsgQueueTimer.Enabled = True
End Sub

' PARTY SYSTEM
Dim PartyBarb : PartyBarb = True
Dim PartyAma : PartyAma = False
Dim PartyNecro : PartyNecro = False
Dim PartySorc : PartySorc = False
Dim PartyPal : PartyPal = False
Dim PartyAss : PartyAss = False
Dim PartyDru : PartyDru = False
'Dim PartyDamageBoostActive : PartyDamageBoostActive = False
'Dim PartyDamageBoostTimer_step : PartyDamageBoostTimer_step = 0
Dim PartyMultiballRunning  : PartyMultiballRunning  = False
Dim PartySpawnStep         : PartySpawnStep         = 0
Dim PartySpawnCount        : PartySpawnCount        = 0

Dim CharThanksName : CharThanksName = ""

' DIFFICULTY (Balls)

Dim MaxBalls : MaxBalls = 3
Dim DifficultySelectActive : DifficultySelectActive = False
Dim SelectedBallCount : SelectedBallCount = 5

Sub ShowDifficultySelect()
    Dim line2
    Select Case SelectedBallCount
        Case 5 : line2 = "◄ NORMAL 3 BALLS ►"
        Case 3 : line2 = "◄ NIGHTMARE 2 BALLS ►"
        Case 1 : line2 = "◄ HELL 1 BALL ►"
    End Select
'    DMDDisplay.Text = "SELECT DIFFICULTY"
    If Not FlexDMDActive Then Exit Sub
    On Error Resume Next
    FlexDMD.LockRenderThread
    FlexDMD.Stage.GetGroup("Score").GetLabel("Line1").Font = FontSmall
    FlexDMD.Stage.GetGroup("Score").GetLabel("Line1").Text = "SELECT DIFFICULTY"
    FlexDMD.Stage.GetGroup("Score").GetLabel("Line1").SetAlignedPosition 64, 11, FlexDMD_Align_Center
    FlexDMD.Stage.GetGroup("Score").GetLabel("Line2").Font = FontSmall
    FlexDMD.Stage.GetGroup("Score").GetLabel("Line2").Text = line2
    FlexDMD.Stage.GetGroup("Score").GetLabel("Line2").SetAlignedPosition 64, 23, FlexDMD_Align_Center
    FlexDMD.UnlockRenderThread
    On Error Goto 0
End Sub

' Progress counters
Dim FlasherSweepStep : FlasherSweepStep = 0
Dim WhirlwindActive : WhirlwindActive = False
Dim WhirlwindStep : WhirlwindStep = 0
Dim WhirlwindCombo : WhirlwindCombo = 0
Dim WhirlwindHitsThis : WhirlwindHitsThis = 3
Const WhirlwindFloorHits = 3      ' minimum swings, even with no gear
Const WhirlwindMs        = 3000   ' strike duration — match the whirlwind1 sound
Const WhirlwindGraceMs   = 1000   ' extra re-hit window after the strikes end
Dim RampAoeCount : RampAoeCount = 0
Const RampAoeTarget   = 3      ' ramp hits to charge the AOE
Dim NextBossOneHit : NextBossOneHit = False
Dim LastRuneWordJackpot : LastRuneWordJackpot = 0
Dim InstantKillActive : InstantKillActive = False
Dim InstantKillTimer_step : InstantKillTimer_step = 0
Dim RampGemCount : RampGemCount = 0
Dim EnemiesKilled : EnemiesKilled = 0
Dim KillMilestoneCount : KillMilestoneCount = 0
Dim FirstKillLootGiven : FirstKillLootGiven = False
Dim ForceNextLootTier  : ForceNextLootTier = -1
Dim ForceFrontLoot     : ForceFrontLoot = False
Dim MysteryKillCount : MysteryKillCount = 0
Dim MysteryReady : MysteryReady = False
Dim MysteryRoll : MysteryRoll = -1
Dim MysteryAnimStep : MysteryAnimStep = 0
Dim MysteryAnimInterval : MysteryAnimInterval = 100
Dim MysteryAnimCycleCount : MysteryAnimCycleCount = 0
Dim BallGoldCount : BallGoldCount = 0
Dim BallArenaKillCount : BallArenaKillCount = 0
Dim BallArenaKillScore : BallArenaKillScore = 0
Dim AuraKillScore : AuraKillScore = 0
Dim MysteryActive : MysteryActive = False


'AMBUSH
Dim AmbushActive : AmbushActive = False
Dim AmbushFailedPenalty : AmbushFailedPenalty = 0
Dim AmbushHappenedThisAct  : AmbushHappenedThisAct  = False
Dim AmbushKillsRequired : AmbushKillsRequired = 10
Dim AmbushKillCount : AmbushKillCount = 0
Dim AmbushBossCount : AmbushBossCount = 0
Dim AmbushTriggerNum : AmbushTriggerNum = 0

'LOOT SYSTEM

Dim LastKilledRarity : LastKilledRarity = RARITY_NORMAL
Dim FirstPackDone    : FirstPackDone    = False
Dim LastKilledType   : LastKilledType   = 0

' Kill Streak system
Dim KillStreakCount  : KillStreakCount  = 0

Function GetDifficultyTier()
    Dim baseTier
    If EnemiesKilled < 50 Then
        baseTier = 0
    ElseIf EnemiesKilled < 125 Then
        baseTier = 1
    ElseIf EnemiesKilled < 250 Then
        baseTier = 2
    Else
        baseTier = 3
    End If
    GetDifficultyTier = baseTier
    If GetDifficultyTier > 3 Then GetDifficultyTier = 3
End Function

Function GetDifficultyHPMultiplier()
    Select Case DifficultyLevel
        Case 1 : GetDifficultyHPMultiplier = 2
        Case 2 : GetDifficultyHPMultiplier = 3
        Case Else : GetDifficultyHPMultiplier = 1
    End Select
End Function

Function GetGearDamage()
    ' Returns flat damage bonus from equipped gear: Magic +1, Rare +2, Unique +3 per slot
    Dim bonus, gsi2
    bonus = 0
    For gsi2 = 0 To 6
        Select Case GearSlots(gsi2)
            Case GEAR_MAGIC  : bonus = bonus + 1
            Case GEAR_RARE   : bonus = bonus + 2
            Case GEAR_UNIQUE : bonus = bonus + 3
        End Select
    Next
    GetGearDamage = bonus
End Function

Function GearPieceCount()
    Dim n, gi
    n = 0
    For gi = 0 To 6
        Select Case GearSlots(gi)
            Case GEAR_MAGIC, GEAR_RARE, GEAR_UNIQUE : n = n + 1
        End Select
    Next
    GearPieceCount = n
End Function

Function SlotToLight(slotIndex)
    Select Case slotIndex
        Case 0 : Set SlotToLight = WeaponLight1
        Case 1 : Set SlotToLight = WeaponLight2
        Case 2 : Set SlotToLight = HelmLight
        Case 3 : Set SlotToLight = ChestLight
        Case 4 : Set SlotToLight = BootLight
        Case 5 : Set SlotToLight = GloveLight
        Case 6 : Set SlotToLight = BeltLight
    End Select
End Function

Sub LightGearSlot(slotIndex, tier)
    Dim lt : Set lt = SlotToLight(slotIndex)
    lt.State = 0
    Select Case tier
        Case GEAR_MAGIC
            lt.Color = RGB(105, 105, 255) : lt.ColorFull = RGB(105, 105, 255)
            lt.State = 1
        Case GEAR_RARE
            lt.Color = RGB(255, 255, 100) : lt.ColorFull = RGB(255, 255, 100)
            lt.State = 1
        Case GEAR_UNIQUE
            lt.Color = RGB(180, 100, 20) : lt.ColorFull = RGB(180, 100, 20)
            lt.State = 1
    End Select
End Sub

Sub BlinkGearSlot(slotIndex)
    Dim lt : Set lt = SlotToLight(slotIndex)
    lt.BlinkInterval = 100
    lt.State = 2
    GearBlinkTimer.Enabled = True
End Sub

Sub AssignGearToSlot(tier)
    Dim i, chosenSlot
    chosenSlot = -1

    ' First pass: find an empty slot
    Dim startSlot : startSlot = Int(Rnd * 7)
    For i = 0 To 6
        Dim s : s = (startSlot + i) Mod 7
        If GearSlots(s) = -1 Then
            chosenSlot = s
            Exit For
        End If
    Next

    ' Second pass: find a slot to upgrade
    If chosenSlot = -1 Then
        For i = 0 To 6
            s = (startSlot + i) Mod 7
            If GearSlots(s) < tier Then
                chosenSlot = s
                Exit For
            End If
        Next
    End If

    ' All slots filled with equal or higher tier - blink the first matching slot
    If chosenSlot = -1 Then
        For i = 0 To 6
            If GearSlots(i) = tier Then
                BlinkGearSlot i
                Exit For
            End If
        Next
        ShowMessage "GEAR +" & GetGearDamage() & " DMG"
        Exit Sub
    End If

    GearSlots(chosenSlot) = tier
    LightGearSlot chosenSlot, tier
    BlinkGearSlot chosenSlot
    PlayGearSlotSound chosenSlot
'    If tier = GEAR_UNIQUE Then PlayCallout "UniqueEquipped", 2000
    If tier = GEAR_UNIQUE Then
        StartGIEvent GI_MODE_UNIQUE
'        AddPortalSurge 1
    End If
    If tier = GEAR_UNIQUE Then
        SetFlasherColor 3, 180, 100, 20
        SetFlasherColor 4, 180, 100, 20
        FireFlasher 3
        FireFlasher 4
    End If
'    If tier = GEAR_UNIQUE And Not FirstUniqueEquipped Then
'        FirstUniqueEquipped = True
'        PlayCallout "DiabloUniquesBonus", 4000
'    End If
    If CritCycleActive And AllUniquesEquipped() Then
        CritCycleTimer.Enabled = False
    End If
    ShowMessage "GEAR +" & GetGearDamage() & " DMG"
End Sub

Sub PlayGearSlotSound(slotIndex)
    Select Case slotIndex
        Case 0 : PlaySound "largemetalweapon", 0, 1
        Case 1 : PlaySound "metalshield", 0, 1
        Case 2 : PlaySound "helm", 0, 1
        Case 3 : PlaySound "platearmor", 0, 1
        Case 4 : PlaySound "bootsmetal", 0, 1
        Case 5 : PlaySound "glovesmetal", 0, 1
        Case 6 : PlaySound "belt", 0, 1
    End Select
End Sub

Sub GearBlinkTimer_Timer()
    GearBlinkTimer.Enabled = False
    Dim gbi
    For gbi = 0 To 6
        If GearSlots(gbi) >= 0 Then
            LightGearSlot gbi, GearSlots(gbi)
        End If
    Next
End Sub

'Sub SaveBestGearOnDrain()
'    Dim partyCount : partyCount = GetPartyCount()
'    If partyCount >= 7 Then Exit Sub
'    Dim slotIndex(6), slotTier(6), filledCount
'    filledCount = 0
'    Dim i
'    For i = 0 To 6
'        If GearSlots(i) >= 0 Then
'            slotIndex(filledCount) = i
'            slotTier(filledCount) = GearSlots(i)
'            filledCount = filledCount + 1
'        End If
'    Next
'    If filledCount = 0 Then Exit Sub
'    Dim j, tempIdx, tempTier
'    For i = 0 To filledCount - 2
'        For j = 0 To filledCount - 2 - i
'            If slotTier(j) < slotTier(j + 1) Then
'                tempTier = slotTier(j)
'                slotTier(j) = slotTier(j + 1)
'                slotTier(j + 1) = tempTier
'                tempIdx = slotIndex(j)
'                slotIndex(j) = slotIndex(j + 1)
'                slotIndex(j + 1) = tempIdx
'            End If
'        Next
'    Next
'    Dim keepSlot(6)
'    For i = 0 To 6 : keepSlot(i) = False : Next
'    Dim keepCount : keepCount = partyCount
'    If keepCount > filledCount Then keepCount = filledCount
'    For i = 0 To keepCount - 1
'        keepSlot(slotIndex(i)) = True
'    Next
'    For i = 0 To 6
'        If GearSlots(i) >= 0 And Not keepSlot(i) Then
'            GearSlots(i) = -1
'            SlotToLight(i).State = 0
'        End If
'    Next
'End Sub

Dim GearSlots(6)
Dim GearSlotLights(6)
Dim gsi
For gsi = 0 To 6 : GearSlots(gsi) = -1 : Next

' LOOT TYPES
Const LOOT_GOLD              = 0
Const LOOT_GEM               = 1
Const LOOT_GEAR              = 2

' GEAR TIERS
Const GEAR_MAGIC             = 0
Const GEAR_RARE              = 1
Const GEAR_UNIQUE            = 2

' LOOT SYSTEM
Dim LootPending(3)
Dim LootValue(3)
Dim LootTier(3)
Dim LootActive(3)
Dim MagicFind

' ===== Loot Drop Target Primitive Animation =====
Dim PrimDtLootArr, DtLootArr
PrimDtLootArr = Array(PrimDtLoot1, PrimDtLoot2, PrimDtLoot3, PrimDtLoot4)
DtLootArr     = Array(dtLoot1, dtLoot2, dtLoot3, dtLoot4)

Dim DtLootUpZ(3)        ' raised Z, captured at startup
Dim DtLootRestRotX(3)   ' resting RotX, captured at startup
Dim DtLootCurZ(3)       ' current animated Z
Dim DtLootTargetZ(3)    ' Z we're animating toward
Dim DtLootBend(3)       ' current bend angle (degrees)
Dim DtLootBendPhase(3)  ' 0 = none, 1 = bending back, 2 = returning/dropping
Dim DtLootBobPhase(3)   ' running phase angle for the idle bob (radians)
Const DtLootBobAmp   = 4.0      ' max dip below top Z (table units) — never goes ABOVE top
Const DtLootBobSpeed = 0.04    ' bob phase increment per frame
'Const DtLootPulseSpeed = 0.075   ' unique glow pulse phase increment per frame
'Dim DtLootPulsePhase(3) ' running phase for unique glow pulse
Dim LootGlowLevel : LootGlowLevel = 3   ' loot halo glow value, 0=off..5 (set by "Loot Glow" option)

Dim DtLootRestRotY(3)    ' resting RotY, captured at startup (for spin)
Dim UniqueLevPhase(3)    ' 0=idle,1=rising,2=spinning,3=descending,4=pausing
Dim UniqueLevZ(3)        ' current levitate height above rest
Dim UniqueSpinDeg(3)     ' accumulated spin degrees
Dim UniquePauseCnt(3)    ' pause-frame counter
Const UniqueLevHeight  = 50     ' float height above rest (VP units)
Const UniqueLevRiseSpd = 0.5    ' rise/descend per frame
Const UniqueSpinSpd    = 3     ' spin degrees per frame
Const UniquePauseFrames= 45     ' frames paused at bottom before repeating

Const DtLootDroppedZ  = -28     ' authored dropped (sunk) Z position
Const DtLootDropSpeed = 4.0     ' Z per frame moving down
Const DtLootRaiseSpeed= 2.5     ' Z per frame moving up
Const DtLootMaxBend   = 8       ' max bend degrees (nFozzy DTMaxBend)
Const DtLootBendSpeed = 2.0     ' bend degrees per frame
Const DtLootUniqueLift = 12     ' extra VP units unique rises above the others (tune to taste)
Dim DtLootTier(3)               ' remember each slot's current tier for anim + pulse

Sub SetLootPrimAppearance(idx, tier)
    Dim img, glow
    DtLootTier(idx) = tier
    Select Case tier
        Case GEAR_MAGIC  : img = "loot_magic"
        Case GEAR_RARE   : img = "loot_rare"
        Case GEAR_UNIQUE : img = "loot_unique"
        Case Else        : img = "loot_magic"    ' dropped/neutral (hidden at -28 anyway)
    End Select
    ' all rarities share the slider value; cleared (tier -1) = off
    If tier >= 0 Then
        glow = LootGlowLevel
        PrimDtLootArr(idx).Visible = True      ' re-show on drop (CollectLoot hides on collect)
    Else
        glow = 0
    End If
    PrimDtLootArr(idx).Image = img
    PrimDtLootArr(idx).BlendDisableLighting = glow
End Sub

Sub InitDtLootAnim()
    Dim i
    For i = 0 To 3
        DtLootUpZ(i)      = PrimDtLootArr(i).z       ' capture the UP z you placed it at in the editor
        DtLootRestRotX(i) = PrimDtLootArr(i).RotX
        ' Start DROPPED at the authored -25 (matches ClearLootTargets initial state)
        DtLootCurZ(i)     = DtLootDroppedZ
        DtLootTargetZ(i)  = DtLootDroppedZ
        DtLootBend(i)     = 0
        DtLootBendPhase(i)= 0
        DtLootBobPhase(i)   = i * 1.4                 ' stagger so targets don't bob in unison
'        DtLootPulsePhase(i) = i * 1.4
        DtLootRestRotY(i)   = PrimDtLootArr(i).RotY   ' capture for spin
        UniqueLevPhase(i)   = 0
        UniqueLevZ(i)       = 0
        UniqueSpinDeg(i)    = 0
        UniquePauseCnt(i)   = 0
        PrimDtLootArr(i).z = DtLootDroppedZ          ' immediately sink to -25
    Next
End Sub

Sub SetDtLootAnim(idx, dropped)
    If dropped Then
        DtLootBendPhase(idx) = 1
        DtLootTargetZ(idx)   = DtLootDroppedZ
    Else
        DtLootBendPhase(idx) = 0
        DtLootBend(idx)      = 0
        PrimDtLootArr(idx).RotX = DtLootRestRotX(idx)
        DtLootTargetZ(idx) = DtLootUpZ(idx)          ' all rest at up-Z; unique lift handled by levitate
        If DtLootTier(idx) = GEAR_UNIQUE Then
            UniqueLevPhase(idx) = 1                   ' begin levitate cycle once raised
            UniqueLevZ(idx)     = 0
            UniqueSpinDeg(idx)  = 0
        Else
            UniqueLevPhase(idx) = 0                    ' ensure no stale levitation
            PrimDtLootArr(idx).RotY = DtLootRestRotY(idx)
        End If
    End If
End Sub

Dim li
For li = 0 To 3
    LootActive(li) = False
    LootPending(li) = -1
    LootValue(li) = 0
    LootTier(li) = -1
Next
MagicFind = 0

'**********************************
'     GEM SYSTEM
'**********************************

Function GemName(gemType)
    Select Case gemType
        Case GEM_RUBY     : GemName = "RUBY"
        Case GEM_SAPPHIRE : GemName = "SAPPHIRE"
        Case GEM_EMERALD  : GemName = "EMERALD"
        Case GEM_TOPAZ    : GemName = "TOPAZ"
    End Select
End Function

Dim PendingElementType : PendingElementType = -1
Dim SilverBallID : SilverBallID = -1
Dim FirstCubeHit : FirstCubeHit = True

' GEM QUEUE
Dim GemQueue(9)
Dim GemQueueCount : GemQueueCount = 0
Dim GemWallOpen : GemWallOpen = False

Const GEM_RUBY     = 0
Const GEM_SAPPHIRE = 1
Const GEM_EMERALD  = 2
Const GEM_TOPAZ    = 3

'Special balls
Const ATTACK_FIRE      = 0
Const ATTACK_COLD      = 1
Const ATTACK_POISON    = 2
Const ATTACK_LIGHTNING = 3
Dim BumperPoisoned(4)
Dim BumperFrozen(4)
Dim BumperBurning(4)
Dim BumperChained(4)
Dim FireBurnSlot : FireBurnSlot = -1

'END GEM SYSTEM

'**************************************
'CAIRN STONES ACT 1
'**************************************
Dim PacksCleared : PacksCleared = 0
Dim QuestAvailable : QuestAvailable = False
Dim CairnCompleted : CairnCompleted = False

Dim CairnActive : CairnActive = False
Dim CairnSequence(4)
Dim CairnProgress : CairnProgress = 0
Dim CairnRiseStep : CairnRiseStep = 0
Dim CairnRising : CairnRising = False

' Enemy type constants
Const ENEMY_FALLEN          = 0
Const ENEMY_ZOMBIE          = 1
Const ENEMY_QUILLRAT        = 2
Const ENEMY_CORRUPTED_ROGUE = 3
Const ENEMY_SKELETON        = 4
Const ENEMY_GOATMAN         = 5
Const ENEMY_YETI      = 6
Const ENEMY_SANDMAGGOT      = 7
Const ENEMY_MUMMY           = 8
Const ENEMY_VULTURE         = 9
Const ENEMY_VAMPIRE         = 10
Const ENEMY_PINHEAD         = 11
Const ENEMY_SERPENT         = 12
Const ENEMY_MESQUITO        = 13
Const ENEMY_PYGMY           = 14
Const ENEMY_SPIDER          = 15
Const ENEMY_TENTACLE        = 16
Const ENEMY_THORNEDHULK     = 17
Const ENEMY_ZAKARUMPRIEST   = 18
Const ENEMY_TRAPPEDSOUL     = 19
Const ENEMY_VILEMMOTHER     = 20
Const ENEMY_REGURGITATOR    = 21
Const ENEMY_UNDEADHORROR    = 22
Const ENEMY_MEGADEMON       = 23
Const ENEMY_SIEGERUNNER     = 24
Const ENEMY_BLOODLORD       = 25
Const ENEMY_REANHORDE       = 26
Const ENEMY_SUCCUBUS        = 27
Const ENEMY_DEATHMAULER     = 28
Const ENEMY_BAALMINION      = 29
Const ENEMY_COW = 30

' Rarity constants
Const RARITY_NORMAL   = 0
Const RARITY_CHAMPION = 1
Const RARITY_UNIQUE = 2

Sub FlashAllBumpers(r, g, b)
    ' Map RGB to nearest Flupper color, or just flash all at current color
    Dim i
    For i = 1 To 5
        FlBumperFadeTarget(i) = 1.1
    Next
End Sub

'*****************************************
'        DIABLO 2 GAME VARIABLES
'*****************************************
Dim BIP : BIP = 0
Dim TableStarted : TableStarted = False
Dim Score : Score = 0
Dim CurrentAct : CurrentAct = 1
Dim BattleModeRunning : BattleModeRunning = False
Dim Difficulty : Difficulty = 0
Dim ActiveElementType : ActiveElementType = -1
Dim PoisonTickActive : PoisonTickActive = False
Dim LeapBallID : LeapBallID = -1

Dim BumperHP(4)
Dim BumperMaxHP(4)
Dim BumperType(4)
Dim BumperRarity(4)
Dim BumperActive(4)

Dim CurrentPackType : CurrentPackType = ENEMY_FALLEN
Dim EnemiesRemaining : EnemiesRemaining = 5
Dim CurrentSong : CurrentSong = ""

Sub PlaySong(name)
    If CurrentSong <> name Or CurrentSong = "" Then
        StopSound CurrentSong
        CurrentSong = name
        PlaySound name, -1, MusicVolume, 0, 0, 0, 1, 0
    End If
End Sub

Sub SetMusicVolume(vol)
    MusicVolume = vol
    If CurrentSong <> "" Then
        PlaySound CurrentSong, -1, MusicVolume, 0, 0, 0, 1, 0
    End If
End Sub

Dim GameActive : GameActive = False
Dim GameOverSequenceActive : GameOverSequenceActive = False
Dim BallNumber : BallNumber = 0
Dim HighScore : HighScore = 0

'******************************************************
'   Custom Ball Teleport (Plunger Lane to Playfield)
'******************************************************

Dim BallToTeleport

' --- Mercenary Multiball (Town Portal hire) ---
Function MercPrice()
    If Not FirstMercGiven Then
        MercPrice = 500000       ' one-time first hire
        Exit Function
    End If
    MercPrice = MercHires * 1500000   ' 1.5M, 3M, 4.5M... steeper leaps
End Function

Dim MercPortalArmed : MercPortalArmed = False
Dim MercArmStep     : MercArmStep = 0
Dim MercSpawnCount  : MercSpawnCount = 0
Dim MercHoldKey     : MercHoldKey = 0
Dim MercPinged      : MercPinged = False
Dim FirstMercGiven : FirstMercGiven = False
Dim MercHires : MercHires = 0
Dim MercMultiballActive : MercMultiballActive = False
Dim MercPingPending : MercPingPending = False
Dim MercSweepStep   : MercSweepStep = 0
Dim MercSweepTotal  : MercSweepTotal = 0


Sub PlungerKicker_Hit()
  If NarniaBallActive Then
        NarniaBallActive = False
        Dim nAngle : If Rnd > 0.5 Then nAngle = 165 Else nAngle = 195
        PlayfieldKicker.Enabled = True
        Dim nBall : Set nBall = PlayfieldKicker.CreateBall
        SetAllBallsElement ActiveElementType
        PlayfieldKicker.Kick nAngle, 15
        PlaySound "popper_ball", 0, 1
        PlungerKicker.DestroyBall
        Exit Sub
    End If
    If GetBIP() > 1 Then
        If ActiveElementType = ATTACK_POISON Then
            PoisonTickActive = False
            PoisonTimer.Enabled = False
            Dim bpspk : For bpspk = 0 To 4 : BumperPoisoned(bpspk) = False : Next
        End If
        PlungerKicker.DestroyBall
        Exit Sub
    End If
    townportalPF.Visible = True
    PlayfieldKicker.Enabled = True
    Set BallToTeleport = ActiveBall
    BallToTeleport.VelX = 0: BallToTeleport.VelY = 0: BallToTeleport.VelZ = 0
    PlaySound "kicker_enter_center"
    PlaySound "portalenter", 0, 1
    TeleportTimer.Enabled = True
End Sub

Sub PlayfieldKicker_Hit()
    townportalPF.Visible = True
    If MercPortalArmed Then
        Dim thisPrice : thisPrice = MercPrice()   ' lock price before flags change
        MercPortalArmed = False
        MercPinged = False
        FirstMercGiven = True
        MercHires = MercHires + 1
        MercArmTimer.Enabled = False
        MercArmStep = 0
        BallGoldCount = BallGoldCount - thisPrice
        If BallGoldCount < 0 Then BallGoldCount = 0
        PauseAmbushTimer
        PauseKillStreakTimer
        PlaySound "kicker_enter_center"
        PlaySound "portalenter", 0, 1
        PlaySound "gold", 0, 1
        PlayfieldKicker.DestroyBall
        UpdateDMD2 "MERCENARY HIRED!", "-" & FormatNumber(thisPrice, 0, -1, 0, -1) & " GOLD"
        MercMultiballActive = True
        MercSpawnCount = 0
        MercSpawnTimer.Interval = 1500
        MercSpawnTimer.Enabled = True
        Exit Sub
    End If
    If PartyMultiballRunning Or PartySpawnTimer.Enabled Then
        If Rnd > 0.5 Then PlayfieldKicker.Kick 165, 15 Else PlayfieldKicker.Kick 195, 15
        PlaySound "popper_ball", 0, 1
    End If
    If Not PartyMultiballRunning And CurrentAct <> 6 Then
        townportalPF.Visible = False
        PlayfieldKicker.Enabled = False
    End If
End Sub

Sub TeleportTimer_Timer()
    Dim i, NewBall, KickAngle

    If Rnd > 0.5 Then
        KickAngle = 165
    Else
        KickAngle = 195
    End If

    Set NewBall = PlayfieldKicker.CreateBall
  SetAllBallsElement ActiveElementType
    PlayfieldKicker.Kick KickAngle, 15
  PlaySound "popper_ball", 0, 1, AudioPan(PlayfieldKicker), 0, 0, 0, 1, AudioFade(PlayfieldKicker)
    PlaySound "kicker_exit"
    PlaySound "portalenter", 0, 1
  If Not BallSaveActive And Not BallSaveUsed Then
        BallSaveActive = True
        BallSaveTimer.Enabled = False
        BallSaveTimer.Interval = 11000 + BallSaveBonus()
        BallSaveTimer.Enabled = True
        BallSaveL.Color = RGB(174, 0, 0) : BallSaveL.ColorFull = RGB(255, 72, 72) : BallSaveL.State = 1
    BallSaveL2.Color = RGB(174, 0, 0) : BallSaveL2.ColorFull = RGB(255, 72, 72) : BallSaveL2.State = 1
    End If

    townportalPF.Visible = False
    PlayfieldKicker.Enabled = False

    PlungerKicker.DestroyBall
    Set BallToTeleport = Nothing
    TeleportTimer.Enabled = False
End Sub

'******************************************************
'   MERCENARY MULTIBALL (Town Portal hire)
'******************************************************

Sub FireMercSweep()
    MercSweepStep = 0
    MercSweepTotal = 60          ' 10 sweeps x 6 flashers
    If RenderingMode = 2 Then MercFlashTimer.Interval = 40 Else MercFlashTimer.Interval = 80
    MercFlashTimer.Enabled = True
End Sub

Sub MercFlashTimer_Timer()
    If MercSweepStep > 0 Then
        InstantFlasher ((MercSweepStep - 1) Mod 6) + 1, 0
    End If
    MercSweepStep = MercSweepStep + 1
    If MercSweepStep > MercSweepTotal Then
        MercFlashTimer.Enabled = False
        Exit Sub
    End If
    Dim mIdx : mIdx = ((MercSweepStep - 1) Mod 6) + 1
    SetFlasherColor mIdx, 212, 175, 55   ' deep gold — tweak to taste
    InstantFlasher mIdx, 1
End Sub

Sub MercHoldTimer_Timer()
    MercHoldTimer.Enabled = False
    MercHoldKey = 0
    ArmMercPortal
End Sub

Sub ArmMercPortal()
    If Not GameActive Then Exit Sub
    If MercPortalArmed Then Exit Sub
  If MercMultiballActive Then Exit Sub
    If GetBIP() < 1 Then Exit Sub        ' allow arming during multiball
    If BallGoldCount < MercPrice Then
        UpdateDMD2 "NOT ENOUGH GOLD", "NEED " & FormatNumber(MercPrice, 0, -1, 0, -1)
        Exit Sub
    End If
    MercPortalArmed = True
    PlayfieldKicker.Enabled = True
    PlayCallout "portalcast", 2000
    MercPortalShowTimer.Interval = 150
    MercPortalShowTimer.Enabled = True
    UpdateDMD2 "TOWN PORTAL OPEN!", "SHOOT THE PORTAL"
    MercArmStep = 0
    MercArmTimer.Interval = 250
    MercArmTimer.Enabled = True
End Sub

Sub MercPortalShowTimer_Timer()
    MercPortalShowTimer.Enabled = False
    If MercPortalArmed Then townportalPF.Visible = True
End Sub

Sub MercArmTimer_Timer()
    If Not MercPortalArmed Then
        MercArmTimer.Enabled = False
        Exit Sub
    End If
    If GetBIP() < 1 Then DisarmMercPortal
End Sub

Sub DisarmMercPortal()
    MercArmTimer.Enabled = False
    MercArmStep = 0
    If Not MercPortalArmed Then Exit Sub
    MercPortalArmed = False
    PlayfieldKicker.Enabled = False
    townportalPF.Visible = False
    UpdateDMDScore
End Sub

Sub MercSpawnTimer_Timer()
    MercSpawnCount = MercSpawnCount + 1
    PlayfieldKicker.Enabled = True
    Dim mb : Set mb = PlayfieldKicker.CreateBall
    SetAllBallsElement ActiveElementType
    If Rnd > 0.5 Then PlayfieldKicker.Kick 165, 15 Else PlayfieldKicker.Kick 195, 15
    PlayfieldKicker.Enabled = False
    PlaySound "popper_ball", 0, 1, AudioPan(PlayfieldKicker), 0, 0, 0, 1, AudioFade(PlayfieldKicker)
    PlaySound "portalenter", 0, 1
    If MercSpawnCount = 1 Then MercSpawnTimer.Interval = 2000
    If MercSpawnCount >= 2 Then
        MercSpawnTimer.Enabled = False
        PlayfieldKicker.Enabled = False
        townportalPF.Visible = True
        FireMercSweep
        ResumeAmbushTimer
        ResumeKillStreakTimer
        MercSaveActive = True
        BallSaveActive = True
        BallSaveTimer.Enabled = False
        BallSaveTimer.Interval = 5000 + BallSaveBonus()
        BallSaveTimer.Enabled = True
        BallSaveL.Color = RGB(174, 0, 0) : BallSaveL.ColorFull = RGB(255, 72, 72) : BallSaveL.State = 1
        BallSaveL2.Color = RGB(174, 0, 0) : BallSaveL2.ColorFull = RGB(255, 72, 72) : BallSaveL2.State = 1
    End If
End Sub

Sub CheckMercPing()
    If MercMultiballActive Then
        MercPingPending = False
        MercPingTimer.Enabled = False
        Exit Sub
    End If
    If BallGoldCount < MercPrice Then
        MercPinged = False
        MercPingPending = False
        MercPingTimer.Enabled = False
        Exit Sub
    End If
    ' Ally portal shares the same hardware — don't ping over it
    If PartyMultiballRunning Or CharWelcomeTimer.Enabled Or CharThanksTimer.Enabled Then
        MercPingTimer.Enabled = False
        MercPingPending = False
        Exit Sub
    End If
    If MercPinged Then Exit Sub
    MercPingPending = True
    If GetBIP() >= 1 And Not MercPingTimer.Enabled Then
        MercPingTimer.Interval = 1500
        MercPingTimer.Enabled = True
    End If
End Sub

Sub MercPingTimer_Timer()
    MercPingTimer.Enabled = False
    If GameActive And GetBIP() >= 1 And BallGoldCount >= MercPrice And Not MercPortalArmed And Not MercPinged And Not MercMultiballActive Then
        PlayCallout "MercenaryHire", 2000     ' announce leads, alone
        MercArmDelayTimer.Interval = 2000     ' wait out the callout, then open portal + whoosh
        MercArmDelayTimer.Enabled = True
        MercPinged = True
    End If
    MercPingPending = False
End Sub

Sub MercArmDelayTimer_Timer()
    MercArmDelayTimer.Enabled = False
    If GameActive And GetBIP() >= 1 And BallGoldCount >= MercPrice And Not MercPortalArmed And Not MercMultiballActive _
       And Not PartyMultiballRunning And Not CharWelcomeTimer.Enabled And Not CharThanksTimer.Enabled Then
        ArmMercPortal      ' portal image + portalcast whoosh, together, after the announce
        FireMercSweep
    Else
        MercPinged = False ' arm window missed (spent gold / drained / ally portal) — allow a fresh ping later
    End If
End Sub

'*****GI Lights On
Dim xx
For Each xx In GI : xx.State = 1 : Next

' Initialize flicker targets
Dim gii2
For gii2 = 0 To GI.Count - 1
    GIFlickerIntensity(gii2) = GIFlickerBase
    GIFlickerTarget(gii2) = GIFlickerBase
Next
GIFlickerCount = GI.Count

Sub SetGIRow(rowNum, r, g, b)
    Dim c : c = RGB(r, g, b)
    Select Case rowNum
        Case 1
            Light031.Color = c : Light031.ColorFull = c : Light031.Intensity = 15
            Light029.Color = c : Light029.ColorFull = c : Light029.Intensity = 15
        Case 2
            Light028.Color = c : Light028.ColorFull = c : Light028.Intensity = 15
            Light030.Color = c : Light030.ColorFull = c : Light030.Intensity = 15
        Case 3
            Light005.Color = c : Light005.ColorFull = c : Light005.Intensity = 15
            Light032.Color = c : Light032.ColorFull = c : Light032.Intensity = 15
        Case 4
            Light033.Color = c : Light033.ColorFull = c : Light033.Intensity = 15
            Light007.Color = c : Light007.ColorFull = c : Light007.Intensity = 15
        Case 5
            Light006.Color = c : Light006.ColorFull = c : Light006.Intensity = 15
            Light027.Color = c : Light027.ColorFull = c : Light027.Intensity = 15
        Case 6
            Light001.Color = c : Light001.ColorFull = c : Light001.Intensity = 15
            Light026.Color = c : Light026.ColorFull = c : Light026.Intensity = 15
        Case 7
            Light002.Color = c : Light002.ColorFull = c : Light002.Intensity = 15
            Light024.Color = c : Light024.ColorFull = c : Light024.Intensity = 15
            Light022.Color = c : Light022.ColorFull = c : Light022.Intensity = 15
            Light003.Color = c : Light003.ColorFull = c : Light003.Intensity = 15
        Case 8
            Light023.Color = c : Light023.ColorFull = c : Light023.Intensity = 15
            Light021.Color = c : Light021.ColorFull = c : Light021.Intensity = 15
        Case 9
            Light017.Color = c : Light017.ColorFull = c : Light017.Intensity = 15
            Light009.Color = c : Light009.ColorFull = c : Light009.Intensity = 15
            Light008.Color = c : Light008.ColorFull = c : Light008.Intensity = 15
            Light020.Color = c : Light020.ColorFull = c : Light020.Intensity = 15
        Case 10
            Light019.Color = c : Light019.ColorFull = c : Light019.Intensity = 15
            Light018.Color = c : Light018.ColorFull = c : Light018.Intensity = 15
    End Select
End Sub

Sub ResetGIRow(rowNum)
    Dim c
    If GIEventMode = GI_MODE_NONE Then
        c = GetActGIColor()
    Else
        c = RGB(0, 0, 0)
    End If
    Select Case rowNum
        Case 1
            Light031.Color = c : Light031.ColorFull = c : Light031.Intensity = GIFlickerBase
            Light029.Color = c : Light029.ColorFull = c : Light029.Intensity = GIFlickerBase
        Case 2
            Light028.Color = c : Light028.ColorFull = c : Light028.Intensity = GIFlickerBase
            Light030.Color = c : Light030.ColorFull = c : Light030.Intensity = GIFlickerBase
        Case 3
            Light005.Color = c : Light005.ColorFull = c : Light005.Intensity = GIFlickerBase
            Light032.Color = c : Light032.ColorFull = c : Light032.Intensity = GIFlickerBase
        Case 4
            Light033.Color = c : Light033.ColorFull = c : Light033.Intensity = GIFlickerBase
            Light007.Color = c : Light007.ColorFull = c : Light007.Intensity = GIFlickerBase
        Case 5
            Light006.Color = c : Light006.ColorFull = c : Light006.Intensity = GIFlickerBase
            Light027.Color = c : Light027.ColorFull = c : Light027.Intensity = GIFlickerBase
        Case 6
            Light001.Color = c : Light001.ColorFull = c : Light001.Intensity = GIFlickerBase
            Light026.Color = c : Light026.ColorFull = c : Light026.Intensity = GIFlickerBase
        Case 7
            Light002.Color = c : Light002.ColorFull = c : Light002.Intensity = GIFlickerBase
            Light024.Color = c : Light024.ColorFull = c : Light024.Intensity = GIFlickerBase
            Light022.Color = c : Light022.ColorFull = c : Light022.Intensity = GIFlickerBase
            Light003.Color = c : Light003.ColorFull = c : Light003.Intensity = GIFlickerBase
        Case 8
            Light023.Color = c : Light023.ColorFull = c : Light023.Intensity = GIFlickerBase
            Light021.Color = c : Light021.ColorFull = c : Light021.Intensity = GIFlickerBase
        Case 9
            Light017.Color = c : Light017.ColorFull = c : Light017.Intensity = GIFlickerBase
            Light009.Color = c : Light009.ColorFull = c : Light009.Intensity = GIFlickerBase
            Light008.Color = c : Light008.ColorFull = c : Light008.Intensity = GIFlickerBase
            Light020.Color = c : Light020.ColorFull = c : Light020.Intensity = GIFlickerBase
        Case 10
            Light019.Color = c : Light019.ColorFull = c : Light019.Intensity = GIFlickerBase
            Light018.Color = c : Light018.ColorFull = c : Light018.Intensity = GIFlickerBase
    End Select
End Sub

Sub RestoreAllGIRows()
    If TiltActive Then Exit Sub
    If GIEventMode = GI_MODE_AMBUSH Then Exit Sub  ' never kill Ambush GI from outside
    GIEventMode = GI_MODE_NONE
    GIEventStep = 0
    Dim ri
    For ri = 1 To 10
        ResetGIRow ri
    Next
    SetActGI
    DimAllFlashers
    GIFlickerTimer.Enabled = True
End Sub

'******************************************************
'   ZFLD:  FLUPPER DOMES
'******************************************************
' Based on FlupperDoms2.2

' What you need in your table to use these flashers:
' Open this table and your table both in VPX
' Export all the materials domebasemat, Flashermaterial0 - 20 and import them in your table
' Export all textures (images) starting with the name "dome" and "ronddome" and import them into your table with the same names
' Export all textures (images) starting with the name "flasherbloom" and import them into your table with the same names
' Copy a set of 4 objects flasherbase, flasherlit, flasherlight and flasherflash from layer 7 to your table
' If you duplicate the four objects for a new flasher dome, be sure that they all end with the same number (in the 0-20 range)
' Copy the flasherbloom flashers from layer 10 to your table. you will need to make one per flasher dome that you plan to make
' Select the correct flasherbloom texture for each flasherbloom flasher, per flasher dome
' Copy the script below

' Place your flasher base primitive where you want the flasher located on your Table
' Then run InitFlasher in the script with the number of your flasher objects and the color of the flasher.  This will align the flasher object, light object, and
' flasher lit primitive.  It will also assign the appropriate flasher bloom images to the flasher bloom object.
'
' Example: InitFlasher 1, "green"
'
' Color Options: "blue", "green", "red", "purple", "yellow", "white", and "orange"

' You can use the RotateFlasher call to align the Rotz/ObjRotz of the flasher primitives with "handles".  Don't set those values in the editor,
' call the RotateFlasher sub instead (this call will likely crash VP if it's call for the flasher primitives without "handles")
'
' Example: RotateFlasher 1, 180    'where 1 is the flasher number and 180 is the angle of Z rotation

' For flashing the flasher use in the script: "ObjLevel(1) = 1 : FlasherFlash1_Timer"
' This should also work for flashers with variable flash levels from the rom, just use ObjLevel(1) = xx from the rom (in the range 0-1)
'
' Notes (please read!!):
' - Setting TestFlashers = 1 (below in the ScriptsDirectory) will allow you to see how the flasher objects are aligned (need the targetflasher image imported to your table)
' - The rotation of the primitives with "handles" is done with a script command, not on the primitive itself (see RotateFlasher below)
' - Color of the objects are set in the script, not on the primitive itself
' - Screws are optional to copy and position manually
' - If your table is not named "Table1" then change the name below in the script
' - Every flasher uses its own material (Flashermaterialxx), do not use it for anything else
' - Lighting > Bloom Strength affects how the flashers look, do not set it too high
' - Change RotY and RotX of flasherbase only when having a flasher something other then parallel to the playfield
' - Leave RotX of the flasherflash object to -45; this makes sure that the flash effect is visible in FS and DT
' - If you want to resize a flasher, be sure to resize flasherbase, flasherlit and flasherflash with the same percentage
' - If you think that the flasher effects are too bright, change flasherlightintensity and/or flasherflareintensity below

' Some more notes for users of the v1 flashers and/or JP's fading lights routines:
' - Delete all textures/primitives/script/materials in your table from the v1 flashers and scripts before you start; they don't mix well with v2
' - Remove flupperflash(m) routines if you have them; they do not work with this new script
' - Do not try to mix this v2 script with the JP fading light routine (that is making it too complicated), just use the example script below

' example script for rom based tables (non modulated):

' SolCallback(25)="FlashRed"
'
' Sub FlashRed(flstate)
' If Flstate Then
'   ObjTargetLevel(1) = 1
' Else
'   ObjTargetLevel(1) = 0
' End If
'   FlasherFlash1_Timer
' End Sub

' example script for rom based tables (modulated):

' SolModCallback(25)="FlashRed"
'
' Sub FlashRed(level)
' ObjTargetLevel(1) = level/255 : FlasherFlash1_Timer
' End Sub

Sub Flash1(Enabled)
  If Enabled Then
    ObjTargetLevel(1) = 1
  Else
    ObjTargetLevel(1) = 0
  End If
  FlasherFlash1_Timer
  Sound_Flash_Relay enabled, Flasherbase1
End Sub

Sub Flash2(Enabled)
  If Enabled Then
    ObjTargetLevel(2) = 1
  Else
    ObjTargetLevel(2) = 0
  End If
  FlasherFlash2_Timer
  Sound_Flash_Relay enabled, Flasherbase2
End Sub

Sub Flash3(Enabled)
  If Enabled Then
    ObjTargetLevel(3) = 1
  Else
    ObjTargetLevel(3) = 0
  End If
  FlasherFlash3_Timer
  Sound_Flash_Relay enabled, Flasherbase3
End Sub

Sub Flash4(Enabled)
  If Enabled Then
    ObjTargetLevel(4) = 1
  Else
    ObjTargetLevel(4) = 0
  End If
  FlasherFlash4_Timer
  Sound_Flash_Relay enabled, Flasherbase1
End Sub

Dim TestFlashers, TableRef, FlasherLightIntensity, FlasherFlareIntensity, FlasherBloomIntensity, FlasherOffBrightness

' *********************************************************************
TestFlashers = 0        ' *** set this to 1 to check position of flasher object      ***
Set TableRef = Table1      ' *** change this, if your table has another name           ***
FlasherLightIntensity = 0.1  ' *** lower this, if the VPX lights are too bright (i.e. 0.1)     ***
FlasherFlareIntensity = 0.3  ' *** lower this, if the flares are too bright (i.e. 0.1)       ***
FlasherBloomIntensity = 0.2  ' *** lower this, if the blooms are too bright (i.e. 0.1)       ***
FlasherOffBrightness = 0.5    ' *** brightness of the flasher dome when switched off (range 0-2)  ***
' *********************************************************************

Dim ObjLevel(20), objbase(20), objlit(20), objflasher(20), objbloom(20), objlight(20), ObjTargetLevel(20)
'Dim tablewidth, tableheight : tablewidth = TableRef.width : tableheight = TableRef.height

'initialise the flasher color, you can only choose from "green", "red", "purple", "blue", "white" and "yellow"
InitFlasher 1, "white"
InitFlasher 2, "white"
InitFlasher 3, "white"
InitFlasher 4, "white"
InitFlasher 5, "white"
InitFlasher 6, "white"

' rotate the flasher with the command below (first argument = flasher nr, second argument = angle in degrees)
'   RotateFlasher 1,17
'   RotateFlasher 2,0
'   RotateFlasher 3,90
'   RotateFlasher 4,90

Sub InitFlasher(nr, col)
  ' store all objects in an array for use in FlashFlasher subroutine
  Set objbase(nr) = Eval("Flasherbase" & nr)
  Set objlit(nr) = Eval("Flasherlit" & nr)
  Set objflasher(nr) = Eval("Flasherflash" & nr)
  Set objlight(nr) = Eval("Flasherlight" & nr)
  Set objbloom(nr) = Eval("Flasherbloom" & nr)

  ' If the flasher is parallel to the playfield, rotate the VPX flasher object for POV and place it at the correct height
  If objbase(nr).RotY = 0 Then
    objbase(nr).ObjRotZ = Atn( (tablewidth / 2 - objbase(nr).x) / (objbase(nr).y - tableheight * 1.1)) * 180 / 3.14159
    objflasher(nr).RotZ = objbase(nr).ObjRotZ
    objflasher(nr).height = objbase(nr).z + 40
  End If

  ' set all effects to invisible and move the lit primitive at the same position and rotation as the base primitive
  objlight(nr).IntensityScale = 0
  objlit(nr).visible = 0
  objlit(nr).material = "Flashermaterial" & nr
  objlit(nr).RotX = objbase(nr).RotX
  objlit(nr).RotY = objbase(nr).RotY
  objlit(nr).RotZ = objbase(nr).RotZ
  objlit(nr).ObjRotX = objbase(nr).ObjRotX
  objlit(nr).ObjRotY = objbase(nr).ObjRotY
  objlit(nr).ObjRotZ = objbase(nr).ObjRotZ
  objlit(nr).x = objbase(nr).x
  objlit(nr).y = objbase(nr).y
  objlit(nr).z = objbase(nr).z
  objbase(nr).BlendDisableLighting = FlasherOffBrightness

  'rothbauerw
  'Adjust the position of the flasher object to align with the flasher base.
  'Comment out these lines if you want to manually adjust the flasher object
  If objbase(nr).roty > 135 Then
    objflasher(nr).y = objbase(nr).y + 50
    objflasher(nr).height = objbase(nr).z + 20
  Else
    objflasher(nr).y = objbase(nr).y + 20
    objflasher(nr).height = objbase(nr).z + 50
  End If
  objflasher(nr).x = objbase(nr).x

  'rothbauerw
  'Adjust the position of the light object to align with the flasher base.
  'Comment out these lines if you want to manually adjust the flasher object
  objlight(nr).x = objbase(nr).x
  objlight(nr).y = objbase(nr).y
  objlight(nr).bulbhaloheight = objbase(nr).z - 10

  'rothbauerw
  'Assign the appropriate bloom image basked on the location of the flasher base
  'Comment out these lines if you want to manually assign the bloom images
  Dim xthird, ythird
  xthird = tablewidth / 3
  ythird = tableheight / 3
  If objbase(nr).x >= xthird And objbase(nr).x <= xthird * 2 Then
    objbloom(nr).imageA = "flasherbloomCenter"
    objbloom(nr).imageB = "flasherbloomCenter"
  ElseIf objbase(nr).x < xthird And objbase(nr).y < ythird Then
    objbloom(nr).imageA = "flasherbloomUpperLeft"
    objbloom(nr).imageB = "flasherbloomUpperLeft"
  ElseIf  objbase(nr).x > xthird * 2 And objbase(nr).y < ythird Then
    objbloom(nr).imageA = "flasherbloomUpperRight"
    objbloom(nr).imageB = "flasherbloomUpperRight"
  ElseIf objbase(nr).x < xthird And objbase(nr).y < ythird * 2 Then
    objbloom(nr).imageA = "flasherbloomCenterLeft"
    objbloom(nr).imageB = "flasherbloomCenterLeft"
  ElseIf  objbase(nr).x > xthird * 2 And objbase(nr).y < ythird * 2 Then
    objbloom(nr).imageA = "flasherbloomCenterRight"
    objbloom(nr).imageB = "flasherbloomCenterRight"
  ElseIf objbase(nr).x < xthird And objbase(nr).y < ythird * 3 Then
    objbloom(nr).imageA = "flasherbloomLowerLeft"
    objbloom(nr).imageB = "flasherbloomLowerLeft"
  ElseIf  objbase(nr).x > xthird * 2 And objbase(nr).y < ythird * 3 Then
    objbloom(nr).imageA = "flasherbloomLowerRight"
    objbloom(nr).imageB = "flasherbloomLowerRight"
  End If

  ' set the texture and color of all objects
  Select Case objbase(nr).image
    Case "dome2basewhite"
      objbase(nr).image = "dome2base" & col
      objlit(nr).image = "dome2lit" & col

    Case "ronddomebasewhite"
      objbase(nr).image = "ronddomebase" & col
      objlit(nr).image = "ronddomelit" & col

    Case "domeearbasewhite"
      objbase(nr).image = "domeearbase" & col
      objlit(nr).image = "domeearlit" & col
  End Select
  If TestFlashers = 0 Then
    objflasher(nr).imageA = "domeflashwhite"
    objflasher(nr).visible = 0
  End If
  Select Case col
    Case "blue"
      objlight(nr).color = RGB(4,120,255)
      objflasher(nr).color = RGB(200,255,255)
      objbloom(nr).color = RGB(4,120,255)
      objlight(nr).intensity = 5000

    Case "green"
      objlight(nr).color = RGB(12,255,4)
      objflasher(nr).color = RGB(12,255,4)
      objbloom(nr).color = RGB(12,255,4)

    Case "red"
      objlight(nr).color = RGB(255,32,4)
      objflasher(nr).color = RGB(255,32,4)
      objbloom(nr).color = RGB(255,32,4)

    Case "purple"
      objlight(nr).color = RGB(230,49,255)
      objflasher(nr).color = RGB(255,64,255)
      objbloom(nr).color = RGB(230,49,255)

    Case "yellow"
      objlight(nr).color = RGB(200,173,25)
      objflasher(nr).color = RGB(255,200,50)
      objbloom(nr).color = RGB(200,173,25)

    Case "white"
      objlight(nr).color = RGB(255,240,150)
      objflasher(nr).color = RGB(100,86,59)
      objbloom(nr).color = RGB(255,240,150)

    Case "orange"
      objlight(nr).color = RGB(255,70,0)
      objflasher(nr).color = RGB(255,70,0)
      objbloom(nr).color = RGB(255,70,0)
  End Select
  objlight(nr).colorfull = objlight(nr).color
  If TableRef.ShowDT And ObjFlasher(nr).RotX =  - 45 Then
    objflasher(nr).height = objflasher(nr).height - 20 * ObjFlasher(nr).y / tableheight
    ObjFlasher(nr).y = ObjFlasher(nr).y + 10
  End If
End Sub

Sub RotateFlasher(nr, angle)
  angle = ((angle + 360 - objbase(nr).ObjRotZ) Mod 180) / 30
  objbase(nr).showframe(angle)
  objlit(nr).showframe(angle)
End Sub

Sub FlashFlasher(nr)
  If Not objflasher(nr).TimerEnabled Then
    objflasher(nr).TimerEnabled = True
    objflasher(nr).visible = 1
    objbloom(nr).visible = 1
    objlit(nr).visible = 1
  End If
  objflasher(nr).opacity = 1000 * FlasherFlareIntensity * ObjLevel(nr) ^ 2.5
  objbloom(nr).opacity = 100 * FlasherBloomIntensity * ObjLevel(nr) ^ 2.5
  objlight(nr).IntensityScale = 0.5 * FlasherLightIntensity * ObjLevel(nr) ^ 3
  If nr <> 5 And nr <> 6 Then
        objbase(nr).BlendDisableLighting = FlasherOffBrightness + 10 * ObjLevel(nr) ^ 3
    End If
    objlit(nr).BlendDisableLighting = 10 * ObjLevel(nr) ^ 2
  UpdateMaterial "Flashermaterial" & nr,0,0,0,0,0,0,ObjLevel(nr),RGB(255,255,255),0,0,False,True,0,0,0,0
  If Round(ObjTargetLevel(nr),1) > Round(ObjLevel(nr),1) Then
    ObjLevel(nr) = ObjLevel(nr) + 0.3
    If ObjLevel(nr) > 1 Then ObjLevel(nr) = 1
  ElseIf Round(ObjTargetLevel(nr),1) < Round(ObjLevel(nr),1) Then
    ObjLevel(nr) = ObjLevel(nr) * 0.85 - 0.01
    If ObjLevel(nr) < 0 Then ObjLevel(nr) = 0
  Else
    ObjLevel(nr) = Round(ObjTargetLevel(nr),1)
    objflasher(nr).TimerEnabled = False
  End If
  '   ObjLevel(nr) = ObjLevel(nr) * 0.9 - 0.01
  If ObjLevel(nr) < 0 Then
    objflasher(nr).TimerEnabled = False
    objflasher(nr).visible = 0
    objbloom(nr).visible = 0
    objlit(nr).visible = 0
  End If
End Sub

Sub FlasherFlash1_Timer()
  FlashFlasher(1)
End Sub
Sub FlasherFlash2_Timer()
  FlashFlasher(2)
End Sub
Sub FlasherFlash3_Timer()
  FlashFlasher(3)
End Sub
Sub FlasherFlash4_Timer()
  FlashFlasher(4)
End Sub

Sub FlasherFlash5_Timer()
    FlashFlasher(5)
End Sub

Sub FlasherFlash6_Timer()
    FlashFlasher(6)
End Sub

Dim TauntFlashTicks : TauntFlashTicks = 0
Dim TauntFlashTotal : TauntFlashTotal = 0
Dim TauntFlashState : TauntFlashState = False

'******************************************************
'******  END FLUPPER DOMES
'******************************************************

Sub FireFlasher(nr)
    ObjTargetLevel(nr) = 1
    Select Case nr
        Case 1 : FlasherFlash1_Timer
        Case 2 : FlasherFlash2_Timer
        Case 3 : FlasherFlash3_Timer
        Case 4 : FlasherFlash4_Timer
        Case 5 : FlasherFlash5_Timer
        Case 6 : FlasherFlash6_Timer
    End Select
End Sub

Sub DimFlasher(nr)
    ObjTargetLevel(nr) = 0
    Select Case nr
        Case 1 : FlasherFlash1_Timer
        Case 2 : FlasherFlash2_Timer
        Case 3 : FlasherFlash3_Timer
        Case 4 : FlasherFlash4_Timer
        Case 5 : FlasherFlash5_Timer
        Case 6 : FlasherFlash6_Timer
    End Select
End Sub

Sub FireAllFlashers()
    FireFlasher 1 : FireFlasher 2 : FireFlasher 3 : FireFlasher 4
    FireFlasher 5 : FireFlasher 6
End Sub

Sub DimAllFlashers()
    DimFlasher 1 : DimFlasher 2 : DimFlasher 3 : DimFlasher 4
    DimFlasher 5 : DimFlasher 6
End Sub

Sub SetFlasherColor(nr, r, g, b)
    objflasher(nr).Color = RGB(r, g, b)
    objlight(nr).Color = RGB(r, g, b)
    objlight(nr).ColorFull = RGB(r, g, b)
    objbloom(nr).Color = RGB(r, g, b)
End Sub

Sub StartTauntFlash(ms)
    SetFlasherColor 1, 200, 0, 0 : SetFlasherColor 2, 200, 0, 0
    SetFlasherColor 3, 200, 0, 0 : SetFlasherColor 4, 200, 0, 0
    SetFlasherColor 5, 200, 0, 0 : SetFlasherColor 6, 200, 0, 0
    TauntFlashTicks = 0
    TauntFlashTotal = ms \ 80
    TauntFlashState = False
    If RenderingMode = 2 Then BossTauntFlashTimer.Interval = 40 Else BossTauntFlashTimer.Interval = 80
    BossTauntFlashTimer.Enabled = True
End Sub

Sub BossTauntFlashTimer_Timer()
    TauntFlashTicks = TauntFlashTicks + 1
    TauntFlashState = Not TauntFlashState
    If TauntFlashState Then
        FireAllFlashers
    Else
        DimAllFlashers
    End If
    If TauntFlashTicks >= TauntFlashTotal Then
        BossTauntFlashTimer.Enabled = False
        DimAllFlashers
    End If
End Sub

Sub ApronFlasherDimTimer_Timer()
    ApronFlasherDimTimer.Enabled = False
    DimFlasher 1
    DimFlasher 2
    DimFlasher 3
    DimFlasher 4
    DimFlasher 5
    DimFlasher 6
End Sub

Dim BaalDeathFlasherStep : BaalDeathFlasherStep = 0
Sub BaalDeathFlasherTimer_Timer()
    BaalDeathFlasherStep = BaalDeathFlasherStep + 1
    If BaalDeathFlasherStep > 40 Then
        BaalDeathFlasherTimer.Enabled = False
        BaalDeathFlasherStep = 0
        DimAllFlashers
        Exit Sub
    End If
    Dim r, g, b
    Select Case BaalDeathFlasherStep Mod 4
        Case 0 : r = 255 : g = 255 : b = 255
        Case 1 : r = 80  : g = 0   : b = 255
        Case 2 : r = 255 : g = 255 : b = 255
        Case 3 : r = 0   : g = 0   : b = 0
    End Select
    SetFlasherColor 1, r, g, b : SetFlasherColor 2, r, g, b
    SetFlasherColor 3, r, g, b : SetFlasherColor 4, r, g, b
    SetFlasherColor 5, r, g, b : SetFlasherColor 6, r, g, b
    FireAllFlashers
End Sub

'*****************************************
'      GI EVENT LIGHTING SYSTEM
'*****************************************

Dim GIBreathPhase   : GIBreathPhase = 0
Dim GIBreathSpeed   : GIBreathSpeed = 0.2
Dim GIBreathMin     : GIBreathMin   = 5.0
Dim GIBreathMax     : GIBreathMax   = 20.0
Dim GIRippleSpread  : GIRippleSpread = 1.6

Sub GIFlickerTimer_Timer()
  If TiltActive Then Exit Sub
    If GIEventMode <> GI_MODE_NONE Then
        GIFlickerTimer.Enabled = False
        Exit Sub
    End If
    GIBreathPhase = GIBreathPhase + GIBreathSpeed
    If GIBreathPhase > 6.2832 Then GIBreathPhase = GIBreathPhase - 6.2832
    Dim breathColor : breathColor = GetActGIColor()
    Dim idx : idx = 0
    Dim gfl
    For Each gfl In GI
        Dim phaseOffset : phaseOffset = idx * GIRippleSpread
        Dim wave : wave = Sin(GIBreathPhase + phaseOffset)
        Dim intensity : intensity = GIBreathMin + (GIBreathMax - GIBreathMin) * (wave + 1) / 2
        gfl.Color = breathColor
        gfl.ColorFull = breathColor
        gfl.Intensity = intensity
        idx = idx + 1
    Next
End Sub

Function GetActGIColor()
    Select Case CurrentAct
        Case 1 : GetActGIColor = RGB(255, 80, 0)
        Case 2 : GetActGIColor = RGB(180, 0, 180)
        Case 3 : GetActGIColor = RGB(0, 120, 20)
        Case 4 : GetActGIColor = RGB(180, 0, 0)
        Case 5 : GetActGIColor = RGB(0, 160, 255)
        Case Else : GetActGIColor = RGB(255, 80, 0)
    End Select
End Function

Function GetActGIDimColor()
    Select Case CurrentAct
        Case 1 : GetActGIDimColor = RGB(180, 40, 0)
        Case 2 : GetActGIDimColor = RGB(80, 0, 80)
        Case 3 : GetActGIDimColor = RGB(0, 60, 10)
        Case 4 : GetActGIDimColor = RGB(80, 0, 0)
        Case 5 : GetActGIColor = RGB(0, 160, 255)
        Case Else : GetActGIDimColor = RGB(180, 40, 0)
    End Select
End Function

Sub SetActGI()
    Dim c : c = GetActGIColor()
    Dim agi
    For agi = 0 To GI.Count - 1
        GI.Item(agi).Color = c
        GI.Item(agi).ColorFull = c
        GI.Item(agi).Intensity = GIFlickerBase
    Next
End Sub

Sub StartGIEvent(mode)
    If GIEventMode = GI_MODE_AMBUSH And mode <> GI_MODE_MYSTERY And mode <> GI_MODE_EXTRABALL Then Exit Sub
    GIEventMode = mode
    GIEventStep = 0
    GIFlickerTimer.Enabled = False
    Dim gei
    For Each gei In GI
        gei.Intensity = 0
    Next
    GIEventTimer.Enabled = False
    GIEventTimer.Enabled = True
End Sub

Sub GIEventTimer_Timer()
    GIEventStep = GIEventStep + 1
    Select Case GIEventMode

        Case GI_MODE_SHOUT
            If GIEventSweepMax > 1 Then
                GIEventTimer.Interval = 20
            Else
                GIEventTimer.Interval = 60
            End If
            Dim sweepPos
            If GIEventSweepDir = 1 Then
                sweepPos = ((GIEventStep - 1) Mod 11) + 1
                If sweepPos <= 10 Then
                    SetGIRow sweepPos, 255, 200, 50
                    If sweepPos > 1 Then ResetGIRow sweepPos - 1
                Else
                    ResetGIRow 10
                    GIEventSweepCount = GIEventSweepCount + 1
                    If GIEventSweepCount >= GIEventSweepMax Then
                        GIEventTimer.Enabled = False
                        RestoreAllGIRows
                    Else
                        If GIEventSweepMax > 1 Then GIEventSweepDir = -1
                        GIEventStep = 0
                    End If
                End If
            Else
                sweepPos = 10 - ((GIEventStep - 1) Mod 11)
                If sweepPos >= 1 Then
                    SetGIRow sweepPos, 255, 200, 50
                    If sweepPos < 10 Then ResetGIRow sweepPos + 1
                Else
                    ResetGIRow 1
                    GIEventSweepCount = GIEventSweepCount + 1
                    If GIEventSweepCount >= GIEventSweepMax Then
                        GIEventTimer.Enabled = False
                        RestoreAllGIRows
                    Else
                        GIEventSweepDir = 1
                        GIEventStep = 0
                    End If
                End If
            End If

        Case GI_MODE_BOSS
            GIEventTimer.Interval = 350
            Dim bossR, bossG
            If GIEventStep Mod 2 = 0 Then
                bossR = 255 : bossG = 0
            Else
                bossR = 0 : bossG = 0
            End If
            Dim bri
            For bri = 1 To 10 : SetGIRow bri, bossR, bossG, 0 : Next
            Dim bossi
            For bossi = 0 To GI.Count - 1
                GI.Item(bossi).Color = RGB(bossR, 0, 0)
                GI.Item(bossi).ColorFull = RGB(bossR, 0, 0)
            Next

        Case GI_MODE_FRENZY
            GIEventTimer.Interval = 40
            Dim frR : frR = 180 + Int(Rnd * 75)
            Dim frG : frG = Int(Rnd * 100)
            Dim fri
            For fri = 1 To 10 : SetGIRow fri, frR, frG, 0 : Next
            Dim frgi
            For frgi = 0 To GI.Count - 1
                GI.Item(frgi).Color = RGB(frR, frG, 0)
                GI.Item(frgi).ColorFull = RGB(frR, frG, 0)
            Next

        Case GI_MODE_DRAIN
            GIEventTimer.Interval = 100
            Select Case GIEventStep
                Case 1
                    Dim dri
                    For dri = 1 To 10
                        SetGIRow dri, 0, 20, 60
                        Select Case dri
                            Case 1  : Light031.Intensity = 1 : Light029.Intensity = 1
                            Case 2  : Light028.Intensity = 1 : Light030.Intensity = 1
                            Case 3  : Light005.Intensity = 1 : Light032.Intensity = 1
                            Case 4  : Light033.Intensity = 1 : Light007.Intensity = 1
                            Case 5  : Light006.Intensity = 1 : Light027.Intensity = 1
                            Case 6  : Light001.Intensity = 1 : Light026.Intensity = 1
                            Case 7  : Light002.Intensity = 1 : Light024.Intensity = 1 : Light022.Intensity = 1 : Light003.Intensity = 1
                            Case 8  : Light023.Intensity = 1 : Light021.Intensity = 1
                            Case 9  : Light017.Intensity = 1 : Light009.Intensity = 1 : Light008.Intensity = 1 : Light020.Intensity = 1
                            Case 10 : Light019.Intensity = 1 : Light018.Intensity = 1
                        End Select
                    Next
                Case 2, 3, 4
                Case 5
                    GIEventTimer.Enabled = False
                    RestoreAllGIRows
            End Select

        Case GI_MODE_RUNEWORD
            GIEventTimer.Interval = 50
            Dim rwRow : rwRow = ((GIEventStep - 1) Mod 10) + 1
            Dim rwPrev : rwPrev = rwRow - 1 : If rwPrev < 1 Then rwPrev = 10
            SetGIRow rwRow, 255, 180, 0
            ResetGIRow rwPrev

        Case GI_MODE_MYSTERY
            GIEventTimer.Interval = 200
            Dim myR, myB
            Select Case GIEventStep Mod 4
                Case 0 : myR = 220 : myB = 255
                Case 1 : myR = 0   : myB = 0
                Case 2 : myR = 255 : myB = 255
                Case 3 : myR = 0   : myB = 0
            End Select
            Dim myi
            For myi = 1 To 10 : SetGIRow myi, myR, 0, myB : Next
            Dim mygi
            For mygi = 0 To GI.Count - 1
                GI.Item(mygi).Color = RGB(myR, 0, myB)
                GI.Item(mygi).ColorFull = RGB(myR, 0, myB)
            Next

        Case GI_MODE_GAMEOVER
            GIEventTimer.Interval = 200
            If GIEventStep <= 10 Then
                Dim goIntensity : goIntensity = GIFlickerBase * (1 - (GIEventStep / 10))
                If goIntensity < 0.5 Then goIntensity = 0.5
                Dim goi
                For goi = 0 To GI.Count - 1
                    GI.Item(goi).Intensity = goIntensity
                    GI.Item(goi).Color = RGB(80, 10, 0)
                Next
            Else
                GIEventTimer.Enabled = False
                GIEventMode = GI_MODE_NONE
            End If

        Case GI_MODE_BOSSWIN
            GIEventTimer.Interval = 80
            Select Case GIEventStep
                Case 1, 3, 5, 7
                    Dim bwi
                    For bwi = 1 To 10 : SetGIRow bwi, 255, 255, 255 : Next
                    Dim bwgi
                    For bwgi = 0 To GI.Count - 1
                        GI.Item(bwgi).Color = RGB(255, 255, 255)
                        GI.Item(bwgi).ColorFull = RGB(255, 255, 255)
                    Next
                Case 2, 4, 6
                    Dim bwj
                    For bwj = 1 To 10
                        SetGIRow bwj, 0, 0, 0
                        Select Case bwj
                            Case 1  : Light031.Intensity = 0 : Light029.Intensity = 0
                            Case 2  : Light028.Intensity = 0 : Light030.Intensity = 0
                            Case 3  : Light005.Intensity = 0 : Light032.Intensity = 0
                            Case 4  : Light033.Intensity = 0 : Light007.Intensity = 0
                            Case 5  : Light006.Intensity = 0 : Light027.Intensity = 0
                            Case 6  : Light001.Intensity = 0 : Light026.Intensity = 0
                            Case 7  : Light002.Intensity = 0 : Light024.Intensity = 0 : Light022.Intensity = 0 : Light003.Intensity = 0
                            Case 8  : Light023.Intensity = 0 : Light021.Intensity = 0
                            Case 9  : Light017.Intensity = 0 : Light009.Intensity = 0 : Light008.Intensity = 0 : Light020.Intensity = 0
                            Case 10 : Light019.Intensity = 0 : Light018.Intensity = 0
                        End Select
                    Next
                    Dim bwgj
                    For bwgj = 0 To GI.Count - 1
                        GI.Item(bwgj).Color = RGB(0, 0, 0)
                        GI.Item(bwgj).ColorFull = RGB(0, 0, 0)
                    Next
                Case 8
                    GIEventTimer.Enabled = False
                    RestoreAllGIRows
            End Select

        Case GI_MODE_AMBUSH
            GIEventTimer.Interval = 150
            Dim ambR
            If GIEventStep Mod 2 = 0 Then
                ambR = 255
            Else
                ambR = 0
            End If
            Dim ambi
            For ambi = 1 To 10 : SetGIRow ambi, ambR, 0, 0 : Next
            Dim ambgi
            For ambgi = 0 To GI.Count - 1
                GI.Item(ambgi).Color = RGB(ambR, 0, 0)
                GI.Item(ambgi).ColorFull = RGB(ambR, 0, 0)
            Next

        Case GI_MODE_SHIELD
            GIEventTimer.Interval = 80
            If GIEventStep <= 10 Then
                SetGIRow GIEventStep, 80, 180, 255
                If GIEventStep > 1 Then ResetGIRow GIEventStep - 1
            ElseIf GIEventStep = 11 Then
                ResetGIRow 10
                GIEventTimer.Enabled = False
                RestoreAllGIRows
            End If

        Case GI_MODE_UNIQUE
            GIEventTimer.Interval = 80
            Select Case GIEventStep
                Case 1, 3, 5
                    Dim uni
                    For uni = 1 To 10 : SetGIRow uni, 255, 160, 20 : Next
                Case 2, 4
                    Dim unj
                    For unj = 1 To 10 : SetGIRow unj, 0, 0, 0 : Next
                Case 6
                    GIEventTimer.Enabled = False
                    RestoreAllGIRows
            End Select

   Case GI_MODE_EXTRABALL
            GIEventTimer.Interval = 60
            Dim ebSweep : ebSweep = ((GIEventStep - 1) Mod 10) + 1
            Dim ebPrev  : ebPrev  = ebSweep - 1 : If ebPrev < 1 Then ebPrev = 10
            SetGIRow ebSweep, 255, 215, 0
            ResetGIRow ebPrev
            Dim ebFlasher : ebFlasher = ((GIEventStep - 1) Mod 6) + 1
            Dim ebFlashPrev : ebFlashPrev = ebFlasher - 1 : If ebFlashPrev < 1 Then ebFlashPrev = 6
            SetFlasherColor ebFlasher, 255, 215, 0
            FireFlasher ebFlasher
            DimFlasher ebFlashPrev
            If GIEventStep >= 20 Then
                GIEventTimer.Enabled = False
                RestoreAllGIRows
                DimAllFlashers
            End If

    End Select
End Sub


'****************************************************************
' ZSLG: Slingshots
'****************************************************************
' RStep and LStep are the variables that increment the animation
Dim RStep, LStep

' Scythe wiggle: rest rotation captured at table init, plus F12 visibility toggle
Dim Scythe1RestRotZ, Scythe2RestRotZ
Dim ScytheVisibleOpt

Sub RightSlingShot_Slingshot
    If Not GameActive Then Exit Sub
    RS.VelocityCorrect(ActiveBall)
  RSling1.Visible = 1
  Sling1.TransY =  - 20   'Sling Metal Bracket
  RStep = 0
  gi1.State = 1
    gi2.State = 1
  RightSlingShot.TimerEnabled = 1
  RightSlingShot.TimerInterval = 10
  '   vpmTimer.PulseSw 52 'Slingshot Rom Switch
  RandomSoundSlingshotRight Sling1
End Sub

Sub RightSlingShot_Timer
    Select Case RStep
        Case 0
            Scythe2.RotZ = Scythe2RestRotZ + 12
        Case 1
            Scythe2.RotZ = Scythe2RestRotZ - 7
        Case 2
            Scythe2.RotZ = Scythe2RestRotZ + 4
        Case 3
            RSLing1.Visible = 0
            RSLing2.Visible = 1
            Sling1.TransY = -10
            Scythe2.RotZ = Scythe2RestRotZ - 2
        Case 4
            RSLing2.Visible = 0
            Sling1.TransY = 0
            RightSlingShot.TimerEnabled = 0
            gi1.State = 0
            gi2.State = 0
            Scythe2.RotZ = Scythe2RestRotZ
    End Select
    RStep = RStep + 1
End Sub

Sub LeftSlingShot_Slingshot
  If Not GameActive Then Exit Sub
  LS.VelocityCorrect(ActiveBall)
  LSling1.Visible = 1
  Sling2.TransY =  - 20   'Sling Metal Bracket
  LStep = 0
  gi3.State = 1
    gi4.State = 1
  LeftSlingShot.TimerEnabled = 1
  LeftSlingShot.TimerInterval = 10
  '   vpmTimer.PulseSw 51 'Slingshot Rom Switch
  RandomSoundSlingshotLeft Sling2
End Sub

Sub LeftSlingShot_Timer
    Select Case LStep
        Case 0
            Scythe1.RotZ = Scythe1RestRotZ + 12
        Case 1
            Scythe1.RotZ = Scythe1RestRotZ - 7
        Case 2
            Scythe1.RotZ = Scythe1RestRotZ + 4
        Case 3
            LSLing1.Visible = 0
            LSLing2.Visible = 1
            Sling2.TransY = -10
            Scythe1.RotZ = Scythe1RestRotZ - 2
        Case 4
            LSLing2.Visible = 0
            Sling2.TransY = 0
            LeftSlingShot.TimerEnabled = 0
            gi3.State = 0
            gi4.State = 0
            Scythe1.RotZ = Scythe1RestRotZ
    End Select
    LStep = LStep + 1
End Sub

Sub TestSlingShot_Slingshot
  TS.VelocityCorrect(ActiveBall)
End Sub

''*********************************************************************
''                 Positional Sound Playback Functions
''*********************************************************************
'
'Function Min(a, b)
'    If a < b Then Min = a Else Min = b
'End Function
'
'Sub PlayXYSound(soundname, tableobj, loopcount, volume, randompitch, pitch, useexisting, restart)
' PlaySound soundname, loopcount, volume, AudioPan(tableobj), randompitch, pitch, useexisting, restart, AudioFade(tableobj)
'End Sub
'
'Sub PlaySoundAt(soundname, tableobj)
'    PlaySound soundname, 1, 1, AudioPan(tableobj), 0,0,0, 1, AudioFade(tableobj)
'End Sub
'
'Sub PlaySoundAtBall(soundname)
'    PlaySoundAt soundname, ActiveBall
'End Sub
'
'Sub PlaySoundAtLevelStatic(playsoundparams, aVol, tableobj)
'    PlaySound playsoundparams, 0, Min(aVol,1) * VolumeDial, AudioPan(tableobj), 0, 0, 0, 0, AudioFade(tableobj)
'End Sub
'
'Sub PlaySoundAtLevelExistingStatic(playsoundparams, aVol, tableobj)
'    PlaySound playsoundparams, 0, Min(aVol,1) * VolumeDial, AudioPan(tableobj), 0, 0, 1, 0, AudioFade(tableobj)
'End Sub
'
'Sub PlaySoundAtLevelStaticLoop(playsoundparams, aVol, tableobj)
'    PlaySound playsoundparams, -1, Min(aVol,1) * VolumeDial, AudioPan(tableobj), 0, 0, 0, 0, AudioFade(tableobj)
'End Sub
'
'Sub PlaySoundAtLevelStaticRandomPitch(playsoundparams, aVol, randomPitch, tableobj)
'    PlaySound playsoundparams, 0, Min(aVol,1) * VolumeDial, AudioPan(tableobj), randomPitch, 0, 0, 0, AudioFade(tableobj)
'End Sub
'
'Sub PlaySoundAtLevelActiveBall(playsoundparams, aVol)
'    PlaySound playsoundparams, 0, Min(aVol,1) * VolumeDial, AudioPan(ActiveBall), 0, 0, 0, 0, AudioFade(ActiveBall)
'End Sub
'
'Sub PlaySoundAtLevelExistingActiveBall(playsoundparams, aVol)
'    PlaySound playsoundparams, 0, Min(aVol,1) * VolumeDial, AudioPan(ActiveBall), 0, 0, 1, 0, AudioFade(ActiveBall)
'End Sub
'
'Sub PlaySoundAtLeveTimerActiveBall(playsoundparams, aVol, ballvariable)
'    PlaySound playsoundparams, 0, Min(aVol,1) * VolumeDial, AudioPan(ballvariable), 0, 0, 0, 0, AudioFade(ballvariable)
'End Sub
'
'Sub PlaySoundAtLevelTimerExistingActiveBall(playsoundparams, aVol, ballvariable)
'    PlaySound playsoundparams, 0, Min(aVol,1) * VolumeDial, AudioPan(ballvariable), 0, 0, 1, 0, AudioFade(ballvariable)
'End Sub
'
'
''*********************************************************************
''                     Supporting Ball & Sound Functions
''*********************************************************************
'
'Function AudioFade(tableobj)
'    Dim tmp
'    tmp = tableobj.y * 2 / tableheight - 1
'    If tmp > 7000 Then
'        tmp = 7000
'    ElseIf tmp < -7000 Then
'        tmp = -7000
'    End If
'    If tmp > 0 Then
'        AudioFade = CSng(tmp ^ 10)
'    Else
'        AudioFade = CSng(-((- tmp) ^ 10))
'    End If
'End Function
'
'Function AudioPan(tableobj)
'    Dim tmp
'    tmp = tableobj.x * 2 / tablewidth - 1
'    If tmp > 7000 Then
'        tmp = 7000
'    ElseIf tmp < -7000 Then
'        tmp = -7000
'    End If
'    If tmp > 0 Then
'        AudioPan = CSng(tmp ^ 10)
'    Else
'        AudioPan = CSng(-((- tmp) ^ 10))
'    End If
'End Function
'
'Function Vol(ball)
'    Vol = CSng(BallVel(ball) ^ 2)
'End Function
'
'Function Volz(ball)
'    Volz = CSng((ball.velz) ^ 2)
'End Function
'
'Function Pitch(ball)
'    Pitch = BallVel(ball) * 20
'End Function
'
'Function BallVel(ball)
'    BallVel = Int(Sqr((ball.VelX ^ 2) + (ball.VelY ^ 2)))
'End Function
'
'Function VolPlayfieldRoll(ball)
'    VolPlayfieldRoll = RollingSoundFactor * 0.0005 * CSng(BallVel(ball) ^ 3)
'End Function
'
'Function PitchPlayfieldRoll(ball)
'    PitchPlayfieldRoll = BallVel(ball) ^ 2 * 15
'End Function
'
'Function RndInt(min, max)
'    RndInt = Int(Rnd() * (max - min + 1) + min)
'End Function
'
'Function RndNum(min, max)
'    RndNum = Rnd() * (max - min) + min
'End Function

'*****************************************
'   rothbauerw's Manual Ball Control
'*****************************************

Dim BCup, BCdown, BCleft, BCright
Dim ControlBallInPlay, ControlActiveBall
Dim BCvel, BCyveloffset, BCboostmulti, BCboost

BCboost = 1
BCvel = 4
BCyveloffset = -0.01
BCboostmulti = 3

ControlBallInPlay = false

Sub StartBallControl_Hit()
  Set ControlActiveBall = ActiveBall
  ControlBallInPlay = true
End Sub

Sub StopBallControl_Hit()
  ControlBallInPlay = false
End Sub

Sub BallControlTimer_Timer()
  If EnableBallControl and ControlBallInPlay then
    If BCright = 1 Then
      ControlActiveBall.velx =  BCvel*BCboost
    ElseIf BCleft = 1 Then
      ControlActiveBall.velx = -BCvel*BCboost
    Else
      ControlActiveBall.velx = 0
    End If

    If BCup = 1 Then
      ControlActiveBall.vely = -BCvel*BCboost
    ElseIf BCdown = 1 Then
      ControlActiveBall.vely =  BCvel*BCboost
    Else
      ControlActiveBall.vely = bcyveloffset
    End If
  End If
End Sub

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
DTSoundLevel = 2.0       'volume multiplier; must not be zero
RolloverSoundLevel = 0.25      'volume level; range [0, 1]
SpinnerSoundLevel = 0.3      'volume level; range [0, 1]

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
    Dim lvl
    If WhirlwindActive Then
        lvl = BumperSoundFactor                  ' timer-driven hit: no ActiveBall in context
    Else
        lvl = Vol(ActiveBall) * BumperSoundFactor
    End If
    PlaySoundAtLevelStatic SoundFX("Bumpers_Top_" & Int(Rnd * 5) + 1, DOFContactors), lvl, Bump
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

    If velocity < 1 Then Exit Sub

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
''*****************************************
''  ninuzzu's FLIPPER SHADOWS v3 (VPX 10.8)
''*****************************************
'
'Sub LeftFlipper_Animate()
'    FlipperLSh.RotZ = LeftFlipper.CurrentAngle
'End Sub
'
'Sub RightFlipper_Animate()
'    FlipperRSh.RotZ = RightFlipper.CurrentAngle
'End Sub

'*****************************************
' ninuzzu's BALL SHADOW
'*****************************************
Dim BallShadow
BallShadow = Array(BallShadow1,BallShadow2,BallShadow3,BallShadow4,BallShadow5,BallShadow6,BallShadow7,BallShadow8)

Sub BallShadowUpdate_timer()
    Dim BOT, b
    BOT = GetBalls

    If UBound(BOT) < (UBound(BallShadow)) Then
        For b = (UBound(BOT) + 1) To UBound(BallShadow)
            BallShadow(b).visible = 0
        Next
    End If

    If UBound(BOT) = -1 Then Exit Sub

    For b = 0 To UBound(BOT)
        If b > UBound(BallShadow) Then Exit For
        If BOT(b).X > 0 And BOT(b).X < Table1.Width And BOT(b).Y > 0 And BOT(b).Y < Table1.Height Then
            BallShadow(b).X = BOT(b).X + (BOT(b).X - (Table1.Width/2)) * 1.25 / BallSize
            BallShadow(b).Y = BOT(b).Y + 12
            BallShadow(b).Size_X = 5
            BallShadow(b).Size_Y = 5
            If BOT(b).Z > 20 Then
                BallShadow(b).visible = 1
            Else
                BallShadow(b).visible = 0
            End If
        Else
            BallShadow(b).visible = 0
        End If
    Next
End Sub


Sub Pins_Hit (idx)
  PlaySound "pinhit_low", 0, Vol(ActiveBall), AudioPan(ActiveBall), 0, Pitch(ActiveBall), 0, 0, AudioFade(ActiveBall)
End Sub

Sub Metals_Thin_Hit(idx)
    PlaySoundAtLevelActiveBall "Metal_Touch_" & Int(Rnd * 13) + 1, Vol(ActiveBall) * MetalImpactSoundFactor
End Sub

Sub Metals_Medium_Hit(idx)
    PlaySoundAtLevelActiveBall "Metal_Touch_" & Int(Rnd * 13) + 1, Vol(ActiveBall) * MetalImpactSoundFactor
End Sub

Sub Metals2_Hit(idx)
    PlaySoundAtLevelActiveBall "Metal_Touch_" & Int(Rnd * 13) + 1, Vol(ActiveBall) * MetalImpactSoundFactor
End Sub

'Sub Gates_Hit(idx)
'    PlaySoundAtLevelStatic "Gate_FastTrigger_" & Int(Rnd * 2) + 1, GateSoundLevel, ActiveBall
'End Sub

Sub TopOrbitGate_Hit
    PlaySoundAtLevelStatic "Gate_FastTrigger_" & Int(Rnd * 2) + 1, GateSoundLevel, ActiveBall
End Sub

'Sub RandomSoundBumperTop(aBump)
'    PlaySoundAtLevelStatic SoundFX("Bumpers_Top_" & Int(Rnd * 5) + 1, DOFContactors), Vol(ActiveBall) * BumperSoundFactor, aBump
'End Sub

Sub Spinner001_Spin
    PlaySoundAtLevelStatic "Spinner", SpinnerSoundLevel, Spinner001
    AddScore 10000
    CrankCubeLid
'    If CurrentAct = 6 And PortalActive Then
'        PortalTime = PortalTime + 1000
'        If PortalTime > PortalCap() Then PortalTime = PortalCap()
'    End If
End Sub

'Sub Rubbers_Hit(idx)
'    Dim finalspeed
'    finalspeed = Sqr(ActiveBall.velx * ActiveBall.velx + ActiveBall.vely * ActiveBall.vely)
'    If finalspeed > 5 Then
'        RandomSoundRubberStrong 1
'    End If
'    If finalspeed <= 5 Then
'        RandomSoundRubberWeak
'    End If
'End Sub

Sub Posts_Hit(idx)
    Dim finalspeed
    finalspeed = Sqr(ActiveBall.velx * ActiveBall.velx + ActiveBall.vely * ActiveBall.vely)
    If finalspeed > 5 Then
        RandomSoundRubberStrong 1
    Else
        RandomSoundRubberWeak
    End If
End Sub


'******************************************************
'   ZRRL: RAMP ROLLING SFX
'******************************************************

Dim RampMinLoops : RampMinLoops = 4
Dim RampBalls(6, 2)
Dim RampType(6)
RampBalls(0, 0) = False

'--- RAMPWIRE TRIGGERS ---
Sub WireStartTrigger_Hit()
    WireRampOff
End Sub

Sub WireStartTrigger_UnHit()
    WireRampOn False
End Sub

Sub WireRampOn(isPlastic)
    WaddBall ActiveBall, isPlastic
    RampRollUpdate
End Sub

Sub WireRampOff()
    WRemoveBall ActiveBall.ID
End Sub

Sub WaddBall(input, RampInput)
    Dim x
    For x = 1 To UBound(RampBalls)
        If RampBalls(x, 1) = input.ID Then
            If Not IsEmpty(RampBalls(x, 1)) Then Exit Sub
        End If
    Next
    For x = 1 To UBound(RampBalls)
        If IsEmpty(RampBalls(x, 1)) Then
            Set RampBalls(x, 0) = input
            RampBalls(x, 1) = input.ID
            RampType(x) = RampInput
            RampBalls(x, 2) = 0
            RampBalls(0, 0) = True
            RampRoll.Enabled = True
            Exit Sub
        End If
    Next
End Sub

Sub WRemoveBall(ID)
    Dim ballcount : ballcount = 0
    Dim x
    For x = 1 To UBound(RampBalls)
        If ID = RampBalls(x, 1) Then
            Set RampBalls(x, 0) = Nothing
            RampBalls(x, 1) = Empty
            RampType(x) = Empty
            StopSound "RampLoop" & x
            StopSound "wireloop" & x
        End If
        If Not IsEmpty(RampBalls(x, 1)) Then ballcount = ballcount + 1
    Next
    If ballcount = 0 Then RampBalls(0, 0) = False
End Sub

Sub ClearAllRampLoops()
    Dim x
    For x = 1 To UBound(RampBalls)
        StopSound "RampLoop" & x
        StopSound "wireloop" & x
        Set RampBalls(x, 0) = Nothing
        RampBalls(x, 1) = Empty
        RampType(x) = Empty
    Next
    RampBalls(0, 0) = False
    RampRoll.Enabled = False
End Sub

Sub RampRoll_Timer()
    RampRollUpdate
End Sub

Sub RampRollUpdate()
    Dim x
    For x = 1 To UBound(RampBalls)
        If Not IsEmpty(RampBalls(x, 1)) Then
            If BallVel(RampBalls(x, 0)) > 1 Then
                If RampType(x) Then
                    PlaySound "RampLoop" & x, -1, VolPlayfieldRoll(RampBalls(x, 0)) * RampRollVolume * VolumeDial, AudioPan(RampBalls(x, 0)), 0, BallPitchV(RampBalls(x, 0)), 1, 0, AudioFade(RampBalls(x, 0))
                    StopSound "wireloop" & x
                Else
                    StopSound "RampLoop" & x
                    PlaySound "wireloop" & x, -1, VolPlayfieldRoll(RampBalls(x, 0)) * RampRollVolume * VolumeDial, AudioPan(RampBalls(x, 0)), 0, BallPitchV(RampBalls(x, 0)), 1, 0, AudioFade(RampBalls(x, 0))
                End If
                RampBalls(x, 2) = RampBalls(x, 2) + 1
            Else
                StopSound "RampLoop" & x
                StopSound "wireloop" & x
            End If
            If RampBalls(x, 0).Z < 30 And RampBalls(x, 2) > RampMinLoops Then
                StopSound "RampLoop" & x
                StopSound "wireloop" & x
                WRemoveBall RampBalls(x, 1)
            End If
        Else
            StopSound "RampLoop" & x
            StopSound "wireloop" & x
        End If
    Next
    If Not RampBalls(0, 0) Then RampRoll.Enabled = False
End Sub
Function BallPitchV(ball)
    BallPitchV = PSlope(BallVel(ball), 1, -4000, 60, 7000)
End Function

'--- GEM RAMP TRIGGERS ---
Sub RampStartTrigger_Hit()
    WireRampOn True
End Sub
Sub RampStartTrigger_UnHit()
    If ActiveBall.VelY > 0 Then WireRampOff
End Sub
Sub WireStopTrigger_Hit()
    WireRampOff
End Sub

'--- CHEST RAMP TRIGGERS ---
Sub ChestInTrigger_Hit()
    WireRampOn True
End Sub

Sub ChestInTrigger_UnHit()
    If ActiveBall.VelY > 0 Then WireRampOff
End Sub

Sub ChestOutTrigger_Hit()
    WireRampOff
End Sub

'--- LEAP RAMP TRIGGERS ---

Sub LeapInTrigger_Hit()
    WireRampOn True
End Sub

Sub LeapInTrigger_UnHit()
    If ActiveBall.VelY > 0 Then WireRampOff
End Sub

Sub LeapOutTrigger_Hit()
    WireRampOff
End Sub

Sub LeapTrigger_Hit()
    If Not GameActive Then Exit Sub
    ' Capture high-speed ball before it can overshoot the kicker
    ActiveBall.X = ArenaUpKick.X
    ActiveBall.Y = ArenaUpKick.Y
    ActiveBall.VelX = 0
    ActiveBall.VelY = 0
    ActiveBall.VelZ = 0
End Sub

'END RAMP SOUNDS

'Sub RandomSoundRubberStrong(voladj)
'    Select Case Int(Rnd * 10) + 1
'        Case 1 : PlaySoundAtLevelActiveBall "Rubber_Strong_1", Vol(ActiveBall) * RubberStrongSoundFactor * voladj
'        Case 2 : PlaySoundAtLevelActiveBall "Rubber_Strong_2", Vol(ActiveBall) * RubberStrongSoundFactor * voladj
'        Case 3 : PlaySoundAtLevelActiveBall "Rubber_Strong_3", Vol(ActiveBall) * RubberStrongSoundFactor * voladj
'        Case 4 : PlaySoundAtLevelActiveBall "Rubber_Strong_4", Vol(ActiveBall) * RubberStrongSoundFactor * voladj
'        Case 5 : PlaySoundAtLevelActiveBall "Rubber_Strong_5", Vol(ActiveBall) * RubberStrongSoundFactor * voladj
'        Case 6 : PlaySoundAtLevelActiveBall "Rubber_Strong_6", Vol(ActiveBall) * RubberStrongSoundFactor * voladj
'        Case 7 : PlaySoundAtLevelActiveBall "Rubber_Strong_7", Vol(ActiveBall) * RubberStrongSoundFactor * voladj
'        Case 8 : PlaySoundAtLevelActiveBall "Rubber_Strong_8", Vol(ActiveBall) * RubberStrongSoundFactor * voladj
'        Case 9 : PlaySoundAtLevelActiveBall "Rubber_Strong_9", Vol(ActiveBall) * RubberStrongSoundFactor * voladj
'        Case 10 : PlaySoundAtLevelActiveBall "Rubber_1_Hard", Vol(ActiveBall) * RubberStrongSoundFactor * 0.6 * voladj
'    End Select
'End Sub
'
'Sub RandomSoundRubberWeak()
'    PlaySoundAtLevelActiveBall "Rubber_" & Int(Rnd * 9) + 1, Vol(ActiveBall) * RubberWeakSoundFactor
'End Sub
'
'Sub RandomSoundRubberFlipper(parm)
'    PlaySoundAtLevelActiveBall "Flipper_Rubber_" & Int(Rnd * 7) + 1, parm * RubberFlipperSoundFactor
'End Sub


'*****************************************
'      DIABLO 2 PROGRESSION LOGIC
'*****************************************

'*****************************************
'      DISK ROTATION SYSTEM
'*****************************************
Dim BumperAngle
BumperAngle = 0
Dim BumperRotSpeed
BumperRotSpeed = 1

Sub RotateDisk()
    BumperAngle = BumperAngle + BumperRotSpeed
    If BumperAngle >= 360 Then BumperAngle = BumperAngle - 360
    BattleArenaDisk.ObjRotZ = BumperAngle
End Sub

Sub BumperRotTimer_Timer()
    RotateDisk
End Sub

'*****************************************
'      ENEMY DATA LOOKUPS
'*****************************************

Function EnemyBaseHP(enemyType)
    Select Case enemyType
        Case ENEMY_FALLEN          : EnemyBaseHP = 2
        Case ENEMY_ZOMBIE          : EnemyBaseHP = 6
        Case ENEMY_QUILLRAT        : EnemyBaseHP = 2
        Case ENEMY_CORRUPTED_ROGUE : EnemyBaseHP = 4
        Case ENEMY_SKELETON        : EnemyBaseHP = 4
        Case ENEMY_GOATMAN         : EnemyBaseHP = 8
        Case ENEMY_YETI            : EnemyBaseHP = 6
        Case ENEMY_SANDMAGGOT      : EnemyBaseHP = 2
        Case ENEMY_MUMMY           : EnemyBaseHP = 4
        Case ENEMY_VULTURE         : EnemyBaseHP = 2
        Case ENEMY_VAMPIRE         : EnemyBaseHP = 6
        Case ENEMY_PINHEAD         : EnemyBaseHP = 8
        Case ENEMY_SERPENT         : EnemyBaseHP = 6
        Case ENEMY_MESQUITO        : EnemyBaseHP = 2
        Case ENEMY_PYGMY           : EnemyBaseHP = 4
        Case ENEMY_SPIDER          : EnemyBaseHP = 4
        Case ENEMY_TENTACLE        : EnemyBaseHP = 6
        Case ENEMY_THORNEDHULK     : EnemyBaseHP = 8
        Case ENEMY_ZAKARUMPRIEST   : EnemyBaseHP = 6
        Case ENEMY_TRAPPEDSOUL     : EnemyBaseHP = 4
        Case ENEMY_VILEMMOTHER     : EnemyBaseHP = 4
        Case ENEMY_REGURGITATOR    : EnemyBaseHP = 6
        Case ENEMY_UNDEADHORROR    : EnemyBaseHP = 8
        Case ENEMY_MEGADEMON       : EnemyBaseHP = 10
        Case ENEMY_SIEGERUNNER     : EnemyBaseHP = 4
        Case ENEMY_BLOODLORD       : EnemyBaseHP = 8
        Case ENEMY_REANHORDE       : EnemyBaseHP = 4
        Case ENEMY_SUCCUBUS        : EnemyBaseHP = 6
        Case ENEMY_DEATHMAULER     : EnemyBaseHP = 10
        Case ENEMY_BAALMINION      : EnemyBaseHP = 12
        Case ENEMY_COW             : EnemyBaseHP = 16
        Case Else                  : EnemyBaseHP = 2
    End Select
End Function

Function EnemyBasePoints(enemyType)
    Select Case enemyType
        Case ENEMY_FALLEN          : EnemyBasePoints = 5000
        Case ENEMY_ZOMBIE          : EnemyBasePoints = 8000
        Case ENEMY_QUILLRAT        : EnemyBasePoints = 4000
        Case ENEMY_CORRUPTED_ROGUE : EnemyBasePoints = 10000
        Case ENEMY_SKELETON        : EnemyBasePoints = 7000
        Case ENEMY_GOATMAN         : EnemyBasePoints = 12000
        Case ENEMY_YETI            : EnemyBasePoints = 9000
        Case ENEMY_SANDMAGGOT      : EnemyBasePoints = 5000
        Case ENEMY_MUMMY           : EnemyBasePoints = 8000
        Case ENEMY_VULTURE         : EnemyBasePoints = 6000
        Case ENEMY_VAMPIRE         : EnemyBasePoints = 7000
        Case ENEMY_PINHEAD         : EnemyBasePoints = 12000
        Case ENEMY_SERPENT         : EnemyBasePoints = 10000
        Case ENEMY_MESQUITO        : EnemyBasePoints = 5000
        Case ENEMY_PYGMY           : EnemyBasePoints = 8000
        Case ENEMY_SPIDER          : EnemyBasePoints = 9000
        Case ENEMY_TENTACLE        : EnemyBasePoints = 11000
        Case ENEMY_THORNEDHULK     : EnemyBasePoints = 14000
        Case ENEMY_ZAKARUMPRIEST   : EnemyBasePoints = 12000
        Case ENEMY_TRAPPEDSOUL     : EnemyBasePoints = 6000
        Case ENEMY_VILEMMOTHER     : EnemyBasePoints = 12000
        Case ENEMY_REGURGITATOR    : EnemyBasePoints = 14000
        Case ENEMY_UNDEADHORROR    : EnemyBasePoints = 11000
        Case ENEMY_MEGADEMON       : EnemyBasePoints = 16000
        Case ENEMY_SIEGERUNNER     : EnemyBasePoints = 8000
        Case ENEMY_BLOODLORD       : EnemyBasePoints = 14000
        Case ENEMY_REANHORDE       : EnemyBasePoints = 8000
        Case ENEMY_SUCCUBUS        : EnemyBasePoints = 12000
        Case ENEMY_DEATHMAULER     : EnemyBasePoints = 16000
        Case ENEMY_BAALMINION      : EnemyBasePoints = 20000
    Case ENEMY_COW             : EnemyBasePoints = 25000
        Case Else                  : EnemyBasePoints = 5000
    End Select
End Function

Function RarityHPMultiplier(rarity)
    Select Case rarity
        Case RARITY_NORMAL   : RarityHPMultiplier = 1
        Case RARITY_CHAMPION : RarityHPMultiplier = 3
        Case RARITY_UNIQUE   : RarityHPMultiplier = 6
        Case Else            : RarityHPMultiplier = 1
    End Select
End Function

Function RarityPointsMultiplier(rarity)
    Select Case rarity
        Case RARITY_NORMAL   : RarityPointsMultiplier = 1
        Case RARITY_CHAMPION : RarityPointsMultiplier = 5
        Case RARITY_UNIQUE   : RarityPointsMultiplier = 15
        Case Else            : RarityPointsMultiplier = 1
    End Select
End Function

Function RarityName(rarity)
    Select Case rarity
        Case RARITY_NORMAL   : RarityName = ""
        Case RARITY_CHAMPION : RarityName = "CHAMPION "
        Case RARITY_UNIQUE   : RarityName = "UNIQUE "
        Case Else            : RarityName = ""
    End Select
End Function
Function EnemyName(enemyType)
    Select Case enemyType
        Case ENEMY_FALLEN          : EnemyName = "FALLEN"
        Case ENEMY_ZOMBIE          : EnemyName = "ZOMBIE"
        Case ENEMY_QUILLRAT        : EnemyName = "QUILL RAT"
        Case ENEMY_CORRUPTED_ROGUE : EnemyName = "ROGUE"
        Case ENEMY_SKELETON        : EnemyName = "SKELETON"
        Case ENEMY_GOATMAN         : EnemyName = "GOATMAN"
        Case ENEMY_YETI            : EnemyName = "YETI"
        Case ENEMY_SANDMAGGOT      : EnemyName = "SAND MAGGOT"
        Case ENEMY_MUMMY           : EnemyName = "MUMMY"
        Case ENEMY_VULTURE         : EnemyName = "VULTURE"
        Case ENEMY_VAMPIRE         : EnemyName = "VAMPIRE"
        Case ENEMY_PINHEAD         : EnemyName = "PINHEAD"
        Case ENEMY_SERPENT         : EnemyName = "SERPENT"
        Case ENEMY_MESQUITO        : EnemyName = "MESQUITO"
        Case ENEMY_PYGMY           : EnemyName = "PYGMY"
        Case ENEMY_SPIDER          : EnemyName = "SPIDER"
        Case ENEMY_TENTACLE        : EnemyName = "TENTACLE"
        Case ENEMY_THORNEDHULK     : EnemyName = "THORNED HULK"
        Case ENEMY_ZAKARUMPRIEST   : EnemyName = "ZAKARUM PRIEST"
        Case ENEMY_TRAPPEDSOUL     : EnemyName = "TRAPPED SOUL"
        Case ENEMY_VILEMMOTHER     : EnemyName = "VILE MOTHER"
        Case ENEMY_REGURGITATOR    : EnemyName = "REGURGITATOR"
        Case ENEMY_UNDEADHORROR    : EnemyName = "UNDEAD HORROR"
        Case ENEMY_MEGADEMON       : EnemyName = "MEGADEMON"
        Case ENEMY_SIEGERUNNER     : EnemyName = "SIEGE RUNNER"
        Case ENEMY_BLOODLORD       : EnemyName = "BLOOD LORD"
        Case ENEMY_REANHORDE       : EnemyName = "REANIMATED HORDE"
        Case ENEMY_SUCCUBUS        : EnemyName = "SUCCUBUS"
        Case ENEMY_DEATHMAULER     : EnemyName = "DEATH MAULER"
        Case ENEMY_BAALMINION      : EnemyName = "BAAL MINION"
        Case Else                  : EnemyName = "ENEMY"
    End Select
End Function

Sub PlayDeathSound(enemyType, rarity)
    Select Case enemyType
        Case ENEMY_FALLEN
            Select Case Int(Rnd * 3)
                Case 0 : PlaySound "death4_fallen", 0, .5 * DuckVolume
                Case 1 : PlaySound "death5", 0, .5 * DuckVolume
                Case 2 : PlaySound "death6", 0, .5 * DuckVolume
            End Select
        Case ENEMY_ZOMBIE
            Select Case Int(Rnd * 2)
                Case 0 : PlaySound "death_zombie1", 0, .5 * DuckVolume
                Case 1 : PlaySound "death_zombie2", 0, .5 * DuckVolume
            End Select
        Case ENEMY_QUILLRAT
            PlaySound "death_quillrat", 0, 1 * DuckVolume
        Case ENEMY_CORRUPTED_ROGUE
            Select Case Int(Rnd * 2)
                Case 0 : PlaySound "death_rogue1", 0, .5 * DuckVolume
                Case 1 : PlaySound "death_rogue2", 0, .5 * DuckVolume
            End Select
        Case ENEMY_SKELETON
            Select Case Int(Rnd * 2)
                Case 0 : PlaySound "death_skeleton1", 0, .5 * DuckVolume
                Case 1 : PlaySound "death_skeleton2", 0, .5 * DuckVolume
            End Select
        Case ENEMY_GOATMAN
            PlaySound "death_goatman", 0, .5 * DuckVolume
        Case ENEMY_YETI
            PlaySound "yetideath1", 0, .5 * DuckVolume
        Case ENEMY_SANDMAGGOT
            Select Case Int(Rnd * 4)
                Case 0 : PlaySound "mag_death1", 0, .5 * DuckVolume
                Case 1 : PlaySound "mag_death2", 0, .5 * DuckVolume
                Case 2 : PlaySound "mag_death3", 0, .5 * DuckVolume
                Case 3 : PlaySound "mag_death4", 0, .5 * DuckVolume
            End Select
        Case ENEMY_MUMMY
            Select Case Int(Rnd * 4)
                Case 0 : PlaySound "mum_death1", 0, .5 * DuckVolume
                Case 1 : PlaySound "mum_death2", 0, .5 * DuckVolume
                Case 2 : PlaySound "mum_death3", 0, .5 * DuckVolume
                Case 3 : PlaySound "mum_death4", 0, .5 * DuckVolume
            End Select
        Case ENEMY_VULTURE
            Select Case Int(Rnd * 5)
                Case 0 : PlaySound "vulture_death1", 0, .5 * DuckVolume
                Case 1 : PlaySound "vulture_death2", 0, .5 * DuckVolume
                Case 2 : PlaySound "vulture_death3", 0, .5 * DuckVolume
                Case 3 : PlaySound "vulture_death4", 0, .5 * DuckVolume
                Case 4 : PlaySound "vulture_death5", 0, .5 * DuckVolume
            End Select
        Case ENEMY_VAMPIRE
            Select Case Int(Rnd * 5)
                Case 0 : PlaySound "vamp_death1", 0, .5 * DuckVolume
                Case 1 : PlaySound "vamp_death2", 0, .5 * DuckVolume
                Case 2 : PlaySound "vamp_death3", 0, .5 * DuckVolume
                Case 3 : PlaySound "vamp_death4", 0, .5 * DuckVolume
                Case 4 : PlaySound "vamp_death5", 0, .5 * DuckVolume
            End Select
        Case ENEMY_PINHEAD
            Select Case Int(Rnd * 5)
                Case 0 : PlaySound "pin_death1", 0, .5 * DuckVolume
                Case 1 : PlaySound "pin_death2", 0, .5 * DuckVolume
                Case 2 : PlaySound "pin_death3", 0, .5 * DuckVolume
                Case 3 : PlaySound "pin_death4", 0, .5 * DuckVolume
                Case 4 : PlaySound "pin_death5", 0, .5 * DuckVolume
            End Select
        Case ENEMY_SERPENT
            Select Case Int(Rnd * 5)
                Case 0 : PlaySound "serp_death1", 0, .5 * DuckVolume
                Case 1 : PlaySound "serp_death2", 0, .5 * DuckVolume
                Case 2 : PlaySound "serp_death3", 0, .5 * DuckVolume
                Case 3 : PlaySound "serp_death4", 0, .5 * DuckVolume
                Case 4 : PlaySound "serp_death5", 0, .5 * DuckVolume
            End Select
        Case ENEMY_MESQUITO
            Select Case Int(Rnd * 4)
                Case 0 : PlaySound "mesq_death1", 0, .5 * DuckVolume
                Case 1 : PlaySound "mesq_death2", 0, .5 * DuckVolume
                Case 2 : PlaySound "mesq_death3", 0, .5 * DuckVolume
                Case 3 : PlaySound "mesq_death4", 0, .5 * DuckVolume
            End Select
        Case ENEMY_PYGMY
            Select Case Int(Rnd * 6)
                Case 0 : PlaySound "pygmy_death1", 0, .5 * DuckVolume
                Case 1 : PlaySound "pygmy_death2", 0, .5 * DuckVolume
                Case 2 : PlaySound "pygmy_death3", 0, .5 * DuckVolume
                Case 3 : PlaySound "pygmy_death4", 0, .5 * DuckVolume
                Case 4 : PlaySound "pygmy_death5", 0, .5 * DuckVolume
                Case 5 : PlaySound "pygmy_death6", 0, .5 * DuckVolume
            End Select
        Case ENEMY_SPIDER
            Select Case Int(Rnd * 6)
                Case 0 : PlaySound "spider_death1", 0, .5 * DuckVolume
                Case 1 : PlaySound "spider_death2", 0, .5 * DuckVolume
                Case 2 : PlaySound "spider_death3", 0, .5 * DuckVolume
                Case 3 : PlaySound "spider_death4", 0, .5 * DuckVolume
                Case 4 : PlaySound "spider_death5", 0, .5 * DuckVolume
                Case 5 : PlaySound "spider_death6", 0, .5 * DuckVolume
            End Select
        Case ENEMY_TENTACLE
            Select Case Int(Rnd * 4)
                Case 0 : PlaySound "tent_death1", 0, .5 * DuckVolume
                Case 1 : PlaySound "tent_death2", 0, .5 * DuckVolume
                Case 2 : PlaySound "tent_death3", 0, .5 * DuckVolume
                Case 3 : PlaySound "tent_death4", 0, .5 * DuckVolume
            End Select
        Case ENEMY_THORNEDHULK
            Select Case Int(Rnd * 4)
                Case 0 : PlaySound "thorn_death1", 0, .5 * DuckVolume
                Case 1 : PlaySound "thorn_death2", 0, .5 * DuckVolume
                Case 2 : PlaySound "thorn_death3", 0, .5 * DuckVolume
                Case 3 : PlaySound "thorn_death4", 0, .5 * DuckVolume
            End Select
        Case ENEMY_ZAKARUMPRIEST
            Select Case Int(Rnd * 4)
                Case 0 : PlaySound "zak_death1", 0, .5 * DuckVolume
                Case 1 : PlaySound "zak_death2", 0, .5 * DuckVolume
                Case 2 : PlaySound "zak_death3", 0, .5 * DuckVolume
                Case 3 : PlaySound "zak_death4", 0, .5 * DuckVolume
            End Select
    Case ENEMY_TRAPPEDSOUL
            Select Case Int(Rnd * 5)
                Case 0 : PlaySound "trapped_down1", 0, .5 * DuckVolume
                Case 1 : PlaySound "trapped_down2", 0, .5 * DuckVolume
                Case 2 : PlaySound "trapped_down3", 0, .5 * DuckVolume
                Case 3 : PlaySound "trapped_down4", 0, .5 * DuckVolume
                Case 4 : PlaySound "trapped_down5", 0, .5 * DuckVolume
            End Select
        Case ENEMY_VILEMMOTHER
            Select Case Int(Rnd * 4)
                Case 0 : PlaySound "vilem_death1", 0, .5 * DuckVolume
                Case 1 : PlaySound "vilem_death2", 0, .5 * DuckVolume
                Case 2 : PlaySound "vilem_death3", 0, .5 * DuckVolume
                Case 3 : PlaySound "vilem_death4", 0, .5 * DuckVolume
            End Select
        Case ENEMY_REGURGITATOR
            Select Case Int(Rnd * 4)
                Case 0 : PlaySound "reg_death1", 0, .5 * DuckVolume
                Case 1 : PlaySound "reg_death2", 0, .5 * DuckVolume
                Case 2 : PlaySound "reg_death3", 0, .5 * DuckVolume
                Case 3 : PlaySound "reg_death4", 0, .5 * DuckVolume
            End Select
        Case ENEMY_UNDEADHORROR
            Select Case Int(Rnd * 4)
                Case 0 : PlaySound "uhor_death1", 0, .5 * DuckVolume
                Case 1 : PlaySound "uhor_death2", 0, .5 * DuckVolume
                Case 2 : PlaySound "uhor_death3", 0, .5 * DuckVolume
                Case 3 : PlaySound "uhor_death4", 0, .5 * DuckVolume
            End Select
        Case ENEMY_MEGADEMON
            Select Case Int(Rnd * 5)
                Case 0 : PlaySound "mega_death1", 0, .5 * DuckVolume
                Case 1 : PlaySound "mega_death2", 0, .5 * DuckVolume
                Case 2 : PlaySound "mega_death3", 0, .5 * DuckVolume
                Case 3 : PlaySound "mega_death4", 0, .5 * DuckVolume
                Case 4 : PlaySound "mega_death5", 0, .5 * DuckVolume
            End Select
    Case ENEMY_SIEGERUNNER
            Select Case Int(Rnd * 5)
                Case 0 : PlaySound "siege_death1", 0, .5 * DuckVolume
                Case 1 : PlaySound "siege_death2", 0, .5 * DuckVolume
                Case 2 : PlaySound "siege_death3", 0, .5 * DuckVolume
                Case 3 : PlaySound "siege_death4", 0, .5 * DuckVolume
                Case 4 : PlaySound "siege_death5", 0, .5 * DuckVolume
            End Select
        Case ENEMY_BLOODLORD
            Select Case Int(Rnd * 4)
                Case 0 : PlaySound "blord_death1", 0, .5 * DuckVolume
                Case 1 : PlaySound "blord_death2", 0, .5 * DuckVolume
                Case 2 : PlaySound "blord_death3", 0, .5 * DuckVolume
                Case 3 : PlaySound "blord_death4", 0, .5 * DuckVolume
            End Select
        Case ENEMY_REANHORDE
            Select Case Int(Rnd * 7)
                Case 0 : PlaySound "reanhorde_death1", 0, .5 * DuckVolume
                Case 1 : PlaySound "reanhorde_death2", 0, .5 * DuckVolume
                Case 2 : PlaySound "reanhorde_death3", 0, .5 * DuckVolume
                Case 3 : PlaySound "reanhorde_death4", 0, .5 * DuckVolume
                Case 4 : PlaySound "reanhorde_death5", 0, .5 * DuckVolume
                Case 5 : PlaySound "reanhorde_death6", 0, .5 * DuckVolume
                Case 6 : PlaySound "reanhorde_death7", 0, .5 * DuckVolume
            End Select
        Case ENEMY_SUCCUBUS
            Select Case Int(Rnd * 4)
                Case 0 : PlaySound "suc_death1", 0, .5 * DuckVolume
                Case 1 : PlaySound "suc_death2", 0, .5 * DuckVolume
                Case 2 : PlaySound "suc_death3", 0, .5 * DuckVolume
                Case 3 : PlaySound "suc_death4", 0, .5 * DuckVolume
            End Select
        Case ENEMY_DEATHMAULER
            Select Case Int(Rnd * 4)
                Case 0 : PlaySound "dmaul_death1", 0, .5 * DuckVolume
                Case 1 : PlaySound "dmaul_death2", 0, .5 * DuckVolume
                Case 2 : PlaySound "dmaul_death3", 0, .5 * DuckVolume
                Case 3 : PlaySound "dmaul_death4", 0, .5 * DuckVolume
            End Select
        Case ENEMY_BAALMINION
            Select Case Int(Rnd * 4)
                Case 0 : PlaySound "bminion_death1", 0, .5 * DuckVolume
                Case 1 : PlaySound "bminion_death2", 0, .5 * DuckVolume
                Case 2 : PlaySound "bminion_death3", 0, .5 * DuckVolume
                Case 3 : PlaySound "bminion_death4", 0, .5 * DuckVolume
            End Select
    Case ENEMY_COW
            Select Case Int(Rnd * 5)
                Case 0 : PlaySound "cow_death1", 0, .6 * DuckVolume
                Case 1 : PlaySound "cow_death2", 0, .6 * DuckVolume
                Case 2 : PlaySound "cow_death3", 0, .6 * DuckVolume
                Case 3 : PlaySound "cow_death4", 0, .6 * DuckVolume
                Case 4 : PlaySound "cow_death5", 0, .6 * DuckVolume
            End Select
    End Select
End Sub

Sub PlayHitSound(enemyType)
    Select Case enemyType
        Case ENEMY_FALLEN
            Select Case Int(Rnd * 3)
                Case 0 : PlaySound "blade1", 0, 1
                Case 1 : PlaySound "blade2", 0, 1
                Case 2 : PlaySound "blade3", 0, 1
            End Select
        Case ENEMY_ZOMBIE
            PlaySound "blow heavy4", 0, 1
        Case ENEMY_QUILLRAT
            PlaySound "blade4", 0, 1
        Case ENEMY_CORRUPTED_ROGUE
            PlaySound "blade5", 0, 1
        Case ENEMY_SKELETON
            Select Case Int(Rnd * 2)
                Case 0 : PlaySound "blade2", 0, 1
                Case 1 : PlaySound "blade3", 0, 1
            End Select
        Case ENEMY_GOATMAN
            PlaySound "blow heavy4", 0, 1
        Case ENEMY_YETI
            Select Case Int(Rnd * 3)
                Case 0 : PlaySound "blow heavy4", 0, 1
                Case 1 : PlaySound "blow heavy4", 0, 1
                Case 2 : PlaySound "blow heavy4", 0, 1
            End Select
    End Select
End Sub

'*****************************************
'      PACK SPAWNING
'*****************************************

Sub AssignBumperStats(slot, isBoss)
    Dim rarity, hp
    Dim tier : tier = GetDifficultyTier()
    Select Case CurrentAct
        Case 1 : BumperType(slot) = Int(Rnd * 7)       ' 0-6: Act 1 enemies
        Case 2 : BumperType(slot) = 7 + Int(Rnd * 6)   ' 7-12: Act 2 enemies
        Case 3 : BumperType(slot) = 13 + Int(Rnd * 6)  ' 13-18: Act 3 enemies
        Case 4 : BumperType(slot) = 19 + Int(Rnd * 5)  ' 19-23: Act 4 enemies
        Case 5 : BumperType(slot) = 24 + Int(Rnd * 6)  ' 24-29: Act 5 enemies
        Case 6 : BumperType(slot) = ENEMY_COW           ' Act 6: only cows
        Case Else : BumperType(slot) = Int(Rnd * 7)
    End Select
    If Not FirstPackDone Then
        rarity = RARITY_NORMAL          ' first pack always normal — player has no gear yet
    ElseIf AmbushActive Then
        rarity = RARITY_NORMAL
    Else
    Select Case tier
        Case 0
            Dim r0 : r0 = Rnd
            If r0 < 0.82 Then
                rarity = RARITY_NORMAL
            ElseIf r0 < 0.97 Then
                rarity = RARITY_CHAMPION
            Else
                rarity = RARITY_UNIQUE
            End If
        Case 1
            Dim r1 : r1 = Rnd
            If r1 < 0.65 Then
                rarity = RARITY_NORMAL
            ElseIf r1 < 0.90 Then
                rarity = RARITY_CHAMPION
            Else
                rarity = RARITY_UNIQUE
            End If
        Case 2
            Dim r2 : r2 = Rnd
            If r2 < 0.50 Then
                rarity = RARITY_NORMAL
            ElseIf r2 < 0.80 Then
                rarity = RARITY_CHAMPION
            Else
                rarity = RARITY_UNIQUE
            End If
        Case 3
            Dim r3 : r3 = Rnd
            If r3 < 0.35 Then
                rarity = RARITY_NORMAL
            ElseIf r3 < 0.65 Then
                rarity = RARITY_CHAMPION
            Else
                rarity = RARITY_UNIQUE
            End If
    End Select
    End If
    If slot = 4 Then FirstPackDone = True   ' ← unlock rarity rolls after first full pack
    BumperRarity(slot) = rarity
    hp = Int(EnemyBaseHP(BumperType(slot)) * RarityHPMultiplier(rarity) * BossScaling)
    If hp < 1 Then hp = 1
  hp = hp + AmbushFailedPenalty
    BumperHP(slot) = hp
    BumperMaxHP(slot) = hp
    BumperActive(slot) = True
    If Not SuppressBumperAnnounce Then
        Select Case rarity
            Case RARITY_CHAMPION
            Case RARITY_UNIQUE
                ShowMessage "UNIQUE ENEMY!"
                PlaySound "monster_attack", 0, 1
        End Select
    End If
End Sub

Function SlotToBumper(slot)
    Select Case slot
        Case 0 : Set SlotToBumper = Bumper1A
        Case 1 : Set SlotToBumper = Bumper2A
        Case 2 : Set SlotToBumper = Bumper3A
        Case 3 : Set SlotToBumper = Bumper4A
    Case 4 : Set SlotToBumper = Bumper5A
    End Select
End Function

'*****************************************
'      BUMPER HIT HANDLER
'*****************************************

Sub KillFlasherDimTimer_Timer()
    KillFlasherDimTimer.Enabled = False
    DimAllFlashers
End Sub

Dim BumperChilled(4)
Dim BumperChillTime(4)
Dim ci : For ci = 0 To 4 : BumperChilled(ci) = False : BumperChillTime(ci) = 0 : Next

Dim BackstabActive : BackstabActive = False

Sub KillBumper(slot)
    ' --- logical kill: state, scoring, death audio, respawn ---
    BumperActive(slot)   = False        ' must be False while the checks below run
    BumperPoisoned(slot) = False
    LastKilledRarity = BumperRarity(slot)   ' capture BEFORE AssignBumperStats overwrites them
    LastKilledType   = BumperType(slot)
    PlayDeathSound BumperType(slot), BumperRarity(slot)
    EnemiesKilled = EnemiesKilled + 1
    CheckAmbushKill
    CheckMysteryMilestone
    BallKillCount      = BallKillCount + 1
    BallArenaKillCount = BallArenaKillCount + 1
    RollKillReward
    AssignBumperStats slot, (slot = 0)  ' rolls the replacement enemy (also sets BumperActive = True)
    BumperActive(slot) = True           ' belt-and-suspenders; matches your original
End Sub

Sub HitBumper(slot, bumperObj, damage)
    If Not BumperActive(slot) Then Exit Sub
    RandomSoundBumperTop SlotToBumper(slot)
    If InstantKillActive Then
        BumperHP(slot) = 0
    Else
        Dim totalDamage : totalDamage = damage + GetGearDamage()
        If BackstabActive Then totalDamage = totalDamage * 2
        BumperHP(slot) = BumperHP(slot) - totalDamage
    End If
    If BumperHP(slot) <= 0 Then
        FlBumperColor(slot + 1) = "red"
        FlInitBumper slot + 1, "red"
        KillBumper slot
        Select Case LastKilledRarity
            Case RARITY_NORMAL   : SetFlasherColor 3, 255, 80, 0   : SetFlasherColor 4, 255, 80, 0
            Case RARITY_CHAMPION : SetFlasherColor 3, 100, 100, 255 : SetFlasherColor 4, 100, 100, 255
            Case RARITY_UNIQUE   : SetFlasherColor 3, 180, 100, 20  : SetFlasherColor 4, 180, 100, 20
        End Select
        FireFlasher 3
        FireFlasher 4
        KillFlasherDimTimer.Enabled = False
        KillFlasherDimTimer.Enabled = True
    End If
End Sub

Sub Bumper1A_Hit()
    FlBumperFadeTarget(1) = 1
    Bumper1A.TimerEnabled = True
    SetFlasherColor 3, 180, 0, 0 : SetFlasherColor 4, 180, 0, 0
    FireFlasher 3 : FireFlasher 4
    Dim bt : bt = GetActiveBonusBallType()
    If LeapBallID >= 0 And ActiveBall.ID = LeapBallID Then
        HitBumper 0, Bumper1A, 99
        If bt >= 0 And ElementBallActive Then BonusBallHitBumper 0, Bumper1A, bt
    ElseIf bt >= 0 And ElementBallActive Then
        RandomSoundBumperTop Bumper1A
        BonusBallHitBumper 0, Bumper1A, bt
    Else
        HitBumper 0, Bumper1A, 1
    End If
End Sub

Sub Bumper2A_Hit()
    FlBumperFadeTarget(2) = 1
    Bumper2A.TimerEnabled = True
    SetFlasherColor 3, 180, 0, 0 : SetFlasherColor 4, 180, 0, 0
    FireFlasher 3 : FireFlasher 4
    Dim bt : bt = GetActiveBonusBallType()
    If LeapBallID >= 0 And ActiveBall.ID = LeapBallID Then
        HitBumper 1, Bumper2A, 99
        If bt >= 0 And ElementBallActive Then BonusBallHitBumper 1, Bumper2A, bt
    ElseIf bt >= 0 And ElementBallActive Then
        RandomSoundBumperMiddle Bumper2A
        BonusBallHitBumper 1, Bumper2A, bt
    Else
        HitBumper 1, Bumper2A, 1
    End If
End Sub

Sub Bumper3A_Hit()
    FlBumperFadeTarget(3) = 1
    Bumper3A.TimerEnabled = True
    SetFlasherColor 3, 180, 0, 0 : SetFlasherColor 4, 180, 0, 0
    FireFlasher 3 : FireFlasher 4
    Dim bt : bt = GetActiveBonusBallType()
    If LeapBallID >= 0 And ActiveBall.ID = LeapBallID Then
        HitBumper 2, Bumper3A, 99
        If bt >= 0 And ElementBallActive Then BonusBallHitBumper 2, Bumper3A, bt
    ElseIf bt >= 0 And ElementBallActive Then
        RandomSoundBumperTop Bumper3A
        BonusBallHitBumper 2, Bumper3A, bt
    Else
        HitBumper 2, Bumper3A, 1
    End If
End Sub

Sub Bumper4A_Hit()
    FlBumperFadeTarget(4) = 1
    Bumper4A.TimerEnabled = True
    SetFlasherColor 3, 180, 0, 0 : SetFlasherColor 4, 180, 0, 0
    FireFlasher 3 : FireFlasher 4
    Dim bt : bt = GetActiveBonusBallType()
    If LeapBallID >= 0 And ActiveBall.ID = LeapBallID Then
        HitBumper 3, Bumper4A, 99
        If bt >= 0 And ElementBallActive Then BonusBallHitBumper 3, Bumper4A, bt
    ElseIf bt >= 0 And ElementBallActive Then
        RandomSoundBumperBottom Bumper4A
        BonusBallHitBumper 3, Bumper4A, bt
    Else
        HitBumper 3, Bumper4A, 1
    End If
End Sub

Sub Bumper5A_Hit()
    FlBumperFadeTarget(5) = 1
    Bumper5A.TimerEnabled = True
    SetFlasherColor 3, 180, 0, 0 : SetFlasherColor 4, 180, 0, 0
    FireFlasher 3 : FireFlasher 4
    Dim bt : bt = GetActiveBonusBallType()
    If LeapBallID >= 0 And ActiveBall.ID = LeapBallID Then
        HitBumper 4, Bumper5A, 99
        If bt >= 0 And ElementBallActive Then BonusBallHitBumper 4, Bumper5A, bt
    ElseIf bt >= 0 And ElementBallActive Then
        RandomSoundBumperTop Bumper5A
        BonusBallHitBumper 4, Bumper5A, bt
    Else
        HitBumper 4, Bumper5A, 1
    End If
End Sub

Sub RestoreBumperLightColor(nr)
    Select Case FlBumperColor(nr)
        Case "red"
            FlBumperSmallLight(nr).Color     = RGB(255, 4, 0)
            FlBumperSmallLight(nr).ColorFull = RGB(255, 24, 0)
            FlbumperBigLight(nr).Color       = RGB(255, 32, 0)
            FlbumperBigLight(nr).ColorFull   = RGB(255, 32, 0)
        Case "blue"
            FlBumperSmallLight(nr).Color     = RGB(0, 80, 255)
            FlBumperSmallLight(nr).ColorFull = RGB(0, 80, 255)
            FlbumperBigLight(nr).Color       = RGB(32, 80, 255)
            FlbumperBigLight(nr).ColorFull   = RGB(32, 80, 255)
        Case "green"
            FlBumperSmallLight(nr).Color     = RGB(8, 255, 8)
            FlBumperSmallLight(nr).ColorFull = RGB(8, 255, 8)
            FlbumperBigLight(nr).Color       = RGB(32, 255, 32)
            FlbumperBigLight(nr).ColorFull   = RGB(32, 255, 32)
    End Select
End Sub

Sub Bumper1A_Timer() : RestoreBumperLightColor 1 : FlBumperFadeTarget(1) = 0 : Bumper1A.TimerEnabled = False : DimFlasher 3 : DimFlasher 4 : End Sub
Sub Bumper2A_Timer() : RestoreBumperLightColor 2 : FlBumperFadeTarget(2) = 0 : Bumper2A.TimerEnabled = False : DimFlasher 3 : DimFlasher 4 : End Sub
Sub Bumper3A_Timer() : RestoreBumperLightColor 3 : FlBumperFadeTarget(3) = 0 : Bumper3A.TimerEnabled = False : DimFlasher 3 : DimFlasher 4 : End Sub
Sub Bumper4A_Timer() : RestoreBumperLightColor 4 : FlBumperFadeTarget(4) = 0 : Bumper4A.TimerEnabled = False : DimFlasher 3 : DimFlasher 4 : End Sub
Sub Bumper5A_Timer() : RestoreBumperLightColor 5 : FlBumperFadeTarget(5) = 0 : Bumper5A.TimerEnabled = False : DimFlasher 3 : DimFlasher 4 : End Sub
'*****************************************
'      PACK MANAGEMENT
'*****************************************

'Sub CheckKillMilestone()
   ' KillMilestoneCount = KillMilestoneCount + 1

    'If Rnd < 0.30 Then
    '    AwardGold
   ' End If

   ' If KillMilestoneCount Mod 5 = 0 Then
    '    RollKillLoot
   ' End If

'End Sub

'Function AllUniquesEquipped()
'    Dim i
'    For i = 0 To 6
'        If GearSlots(i) <> GEAR_UNIQUE Then
'            AllUniquesEquipped = False
'            Exit Function
'        End If
'    Next
'    AllUniquesEquipped = True
'End Function

Sub RollKillReward()
    Dim earned : earned = EnemyBasePoints(LastKilledType) * RarityPointsMultiplier(LastKilledRarity)
    AddScore earned
    If WhirlwindActive Then
        BallWhirlwindScore = BallWhirlwindScore + earned
    ElseIf BackstabActive Then
        BallBackstabScore = BallBackstabScore + earned
    End If
    BallArenaKillScore = BallArenaKillScore + earned
    If InstantKillActive Then AuraKillScore = AuraKillScore + earned
    IncrementKillStreak

    Select Case LastKilledRarity
        Case RARITY_NORMAL   : If Rnd < 0.15 Then BankGold
        Case RARITY_CHAMPION : If Rnd < 0.60 Then BankGold
        Case RARITY_UNIQUE   : BankGold
    End Select

    KillMilestoneCount = KillMilestoneCount + 1
    If Not FirstKillLootGiven Then
        FirstKillLootGiven = True
        ForceNextLootTier = GEAR_MAGIC   ' first kill always drops a Magic item
        ForceFrontLoot = True            ' ...into dtLoot1 or dtLoot2 only
        RollKillLoot
    End If
    Select Case LastKilledRarity
        Case RARITY_NORMAL   : If KillMilestoneCount Mod 5 = 0 Then RollKillLoot
        Case RARITY_CHAMPION : If Rnd < 0.50 Then RollKillLoot
        Case RARITY_UNIQUE   : RollKillLoot
    End Select
    If Not AmbushActive Then UpdateDMDScore
End Sub

' =========================================================
'  KILL STREAK SYSTEM
'  Thresholds: 5/10/15/20/25 kills within a 5-second window
' =========================================================

Sub IncrementKillStreak()
    If Not GameActive Then Exit Sub
    KillStreakCount = KillStreakCount + 1
    KillStreakTimer.Enabled = False
    KillStreakTimer.Interval = 5000
    KillStreakTimer.Enabled = True
    Select Case KillStreakCount
        Case 5
            ShowBigMessage "SLAUGHTER!"
            PlayCallout "slaughter", 1000
            FireKillStreakSweep 1
        Case 15
            ShowBigMessage "CARNAGE!"
            PlayCallout "carnage", 1000
            FireKillStreakSweep 2
        Case 25
            ShowBigMessage "MASSACRE!"
            PlayCallout "massacre", 1000
            FireKillStreakSweep 3
        Case 35
            ShowBigMessage "EXTERMINATION!"
            PlayCallout "extermination", 2000
            FireKillStreakSweep 4
        Case 45
      ShowBigMessage "ARMAGEDDON!"
      PlayCallout "armageddon", 2000
      FireKillStreakSweep 5
      KillStreakTimer.Enabled = False
      ArmageddonAwardTimer.Interval = 3200
      ArmageddonAwardTimer.Enabled = True
    End Select
End Sub

Sub ArmageddonAwardTimer_Timer()
    ArmageddonAwardTimer.Enabled = False
    AwardKillStreak
End Sub

Sub KillStreakTimer_Timer()
    KillStreakTimer.Enabled = False
    AwardKillStreak
End Sub

Dim KillStreakSweepStep  : KillStreakSweepStep  = 0
Dim KillStreakSweepTotal : KillStreakSweepTotal = 0

Sub AwardKillStreak()
    If KillStreakCount < 5 Then
        KillStreakCount = 0
        Exit Sub
    End If
    Dim tierName, streakBonus, sweepCount
    If KillStreakCount >= 45 Then
        tierName = "ARMAGEDDON STREAK!"    : streakBonus = 10000000 : sweepCount = 5
    ElseIf KillStreakCount >= 35 Then
        tierName = "EXTERMINATION STREAK!" : streakBonus = 5000000  : sweepCount = 4
    ElseIf KillStreakCount >= 25 Then
        tierName = "MASSACRE STREAK!"      : streakBonus = 2500000  : sweepCount = 3
    ElseIf KillStreakCount >= 15 Then
        tierName = "CARNAGE STREAK!"       : streakBonus = 1000000  : sweepCount = 2
    Else
        tierName = "SLAUGHTER STREAK!"     : streakBonus = 250000   : sweepCount = 1
    End If

    streakBonus = streakBonus * CowLevelMultiplier()
    AddScore streakBonus
    BallArenaKillScore = BallArenaKillScore + streakBonus
    UpdateDMD2 tierName, "+" & FormatNumber(streakBonus, 0, -1, 0, -1)
    AmbushClearHold = True               ' borrow the DMD message-hold so the streak lingers
    AmbushClearTimer.Interval = 2500
    AmbushClearTimer.Enabled = True
    FireKillStreakSweep sweepCount
    KillStreakCount = 0
End Sub

Sub FireKillStreakSweep(sweepCount)
    ' Chases flashers 1->6 in sequence, repeating sweepCount times
    ' Slaughter=1, Carnage=2, Massacre=3, Extermination=4, Armageddon=5
    KillStreakSweepStep  = 0
    KillStreakSweepTotal = sweepCount * 6
    KillStreakFlashTimer.Interval = 90
    KillStreakFlashTimer.Enabled = True
End Sub

Sub InstantFlasher(nr, lvl)
    ObjLevel(nr) = lvl
    ObjTargetLevel(nr) = lvl
    Select Case nr
        Case 1 : FlasherFlash1_Timer
        Case 2 : FlasherFlash2_Timer
        Case 3 : FlasherFlash3_Timer
        Case 4 : FlasherFlash4_Timer
        Case 5 : FlasherFlash5_Timer
        Case 6 : FlasherFlash6_Timer
    End Select
End Sub

Sub KillStreakFlashTimer_Timer()
    If KillStreakSweepStep > 0 Then
        InstantFlasher ((KillStreakSweepStep - 1) Mod 6) + 1, 0
    End If
    KillStreakSweepStep = KillStreakSweepStep + 1
    If KillStreakSweepStep > KillStreakSweepTotal Then
        KillStreakFlashTimer.Enabled = False
        ' Restore normal flasher intensity
        FlasherFlareIntensity = 0.3
        FlasherLightIntensity = 0.1
        FlasherBloomIntensity = 0.2
        Exit Sub
    End If
    Dim sweepIdx : sweepIdx = ((KillStreakSweepStep - 1) Mod 6) + 1
    SetFlasherColor sweepIdx, 255, 0, 0   ' blood red
    InstantFlasher sweepIdx, 1
End Sub

' END KILL STREAK LOGIC

Function GoldActScale()
    ' Income rises with act so advancing keeps pace with the hire-driven merc price.
    ' Linger in acts 1-2 and the price you've run up outruns your income; cruise and it holds.
    GoldActScale = 1 + (CurrentAct - 1) * 0.6   ' Act1=1.0x  2=1.6x  3=2.2x  4=2.8x  5=3.4x  6=4.0x
End Function

Sub BankGold()
    Dim goldAmount
    Select Case LastKilledRarity
        Case RARITY_NORMAL   : goldAmount = 7000   + Int(Rnd * 5000)
        Case RARITY_CHAMPION : goldAmount = 30000  + Int(Rnd * 20000)
        Case RARITY_UNIQUE   : goldAmount = 75000  + Int(Rnd * 50000)
        Case Else            : goldAmount = 7000
    End Select
    BallGoldCount = BallGoldCount + Int(goldAmount * GoldActScale())
    PlaySound "gold", 0, 1
    CheckMercPing
End Sub

Sub CheckMysteryMilestone()
    If Not GameActive Then Exit Sub
    If MysteryReady Then Exit Sub
    MysteryKillCount = MysteryKillCount + 1
    If MysteryKillCount Mod 50 = 0 And Not MysteryReady Then
        MysteryReady = True
        ChestRampDown
'        MysteryKickLight.BlinkInterval = 300
'        MysteryKickLight.State = 2
        QuestLight4.Color = RGB(180, 0, 255)
        QuestLight4.ColorFull = RGB(180, 0, 255)
        QuestLight4.BlinkInterval = 300
        QuestLight4.State = 2
        ShowMessage "MYSTERY CHEST READY!"
        PlaySound "waypointignite2", 0, .6
    PlayCallout "MysteryChestDiablo", 2000
    End If
End Sub

Sub SetDtLootMaterial(idx, matName)
    PrimDtLootArr(idx).Material = matName
End Sub

Sub RollKillLoot()
    Dim slot, foundSlot, matName
    Dim searchMax : searchMax = 3
    If ForceFrontLoot Then searchMax = 1   ' restrict this drop to dtLoot1 / dtLoot2
    ForceFrontLoot = False
    foundSlot = -1
    For slot = 0 To searchMax
        If Not LootActive(slot) Then
            foundSlot = slot
            Exit For
        End If
    Next
    If foundSlot = -1 Then ForceNextLootTier = -1 : Exit Sub
    LootPending(foundSlot) = LOOT_GEAR
    LootValue(foundSlot) = RollGearProperty()
    If ForceNextLootTier >= 0 Then
        LootTier(foundSlot) = ForceNextLootTier
        ForceNextLootTier = -1
    Else
        LootTier(foundSlot) = RollGearTier()
    End If
    LootActive(foundSlot) = True
    Select Case LootTier(foundSlot)
        Case GEAR_MAGIC  : matName = "Magic"
        Case GEAR_RARE   : matName = "Rare"
        Case GEAR_UNIQUE : matName = "Unique"
        Case Else        : matName = "TargetDefault"
    End Select
    Select Case foundSlot
        Case 0 : dtLoot1.IsDropped = False : SetLootPrimAppearance 0, LootTier(foundSlot) : SetDtLootAnim 0, False
        Case 1 : dtLoot2.IsDropped = False : SetLootPrimAppearance 1, LootTier(foundSlot) : SetDtLootAnim 1, False
        Case 2 : dtLoot3.IsDropped = False : SetLootPrimAppearance 2, LootTier(foundSlot) : SetDtLootAnim 2, False
        Case 3 : dtLoot4.IsDropped = False : SetLootPrimAppearance 3, LootTier(foundSlot) : SetDtLootAnim 3, False
    End Select
    PlaySound "flippy", 0, 1
End Sub

Sub RespawnBumpers()
    AssignBumperStats 0, True
    AssignBumperStats 1, False
    AssignBumperStats 2, False
    AssignBumperStats 3, False
    AssignBumperStats 4, False
    EnemiesRemaining = 5
    Dim rps : For rps = 0 To 4 : BumperPoisoned(rps) = False : Next
    PoisonTickActive = False
    PoisonTimer.Enabled = False
End Sub

' BARBARIAN LEAP AND ANIM LOGIC

Dim LeapReady : LeapReady = False
Dim BarbLeanStep : BarbLeanStep = 0
Dim BarbJumpStep : BarbJumpStep = 0
Const BarbRestRotX   = 0    ' actual resting RotX in editor
Const BarbRestRotY   = 90    ' actual resting RotY in editor
Const BarbRestZ      = 0    ' actual resting TransZ in editor
Const BarbLeanMax    = 15   ' degrees of lean
Const BarbJumpHeight = 8    ' units up on Z
Const BarbLeanSteps  = 40   ' 40 x 50ms = 2 seconds to full lean

Sub BarbLeanTimer_Timer()
    BarbLeanStep = BarbLeanStep + 1
    If BarbLeanStep = 1 Then PlaySound "TOM_Trunk_Motor_Long", -1, 0.6
    If BarbLeanStep >= BarbLeanSteps Then
        BarbLeanStep = BarbLeanSteps
        BarbLeanTimer.Enabled = False
        StopSound "TOM_Trunk_Motor_Long"
    End If
    Dim t : t = BarbLeanStep / BarbLeanSteps
    BarbPrim.RotX = BarbRestRotX - (BarbLeanMax * 0.5 * t)
    BarbPrim.RotY = BarbRestRotY + (BarbLeanMax * t)
End Sub

Sub BarbJumpTimer_Timer()
    BarbJumpStep = BarbJumpStep + 1
    If BarbJumpStep = 1 Then
        StopSound "TOM_Trunk_Motor_Long"
        PlaySound "TOM_Trunk_Motor_Long", 0, 1
        BarbPrim.RotX = BarbRestRotX - (BarbLeanMax * 0.5)
        BarbPrim.RotY = BarbRestRotY + BarbLeanMax
    ElseIf BarbJumpStep >= 2 And BarbJumpStep <= 8 Then
        Dim riseT : riseT = (BarbJumpStep - 1) / 7
        BarbPrim.TransZ = BarbRestZ + BarbJumpHeight * riseT
        BarbPrim.RotX = BarbRestRotX - (BarbLeanMax * 0.5 * (1 - riseT))
        BarbPrim.RotY = BarbRestRotY + BarbLeanMax * (1 - riseT)
    ElseIf BarbJumpStep >= 9 And BarbJumpStep <= 16 Then
        Dim fallT : fallT = (BarbJumpStep - 8) / 8
        BarbPrim.TransZ = BarbRestZ + BarbJumpHeight * (1 - fallT)
    ElseIf BarbJumpStep = 17 Then
        BarbPrim.TransZ = BarbRestZ + 1
        BarbPrim.RotX = BarbRestRotX - 1
        BarbPrim.RotY = BarbRestRotY + 1
    ElseIf BarbJumpStep = 18 Then
        BarbPrim.TransZ = BarbRestZ
        BarbPrim.RotX = BarbRestRotX
        BarbPrim.RotY = BarbRestRotY
        StopSound "TOM_Trunk_Motor_Long"
        BarbJumpTimer.Enabled = False
        BarbJumpStep = 0
    End If
End Sub

Dim LeapSafetyTimer_step : LeapSafetyTimer_step = 0

Sub LeapSafetyTimer_Timer()
    LeapSafetyTimer_step = LeapSafetyTimer_step + 1
    If Not LeapReady Then
        LeapSafetyTimer.Enabled = False
        LeapSafetyTimer_step = 0
        Exit Sub
    End If
    If LeapSafetyTimer_step >= 10 Then
        LeapSafetyTimer.Enabled = False
        LeapSafetyTimer_step = 0
        FireLeapAttack
    End If
End Sub

Sub ArenaUpKick_Hit()
    PauseAmbushTimer
  PauseKillStreakTimer
    If Not GameActive Then Exit Sub
    LeapBallID = ActiveBall.ID
    ArenaUpKick.TimerEnabled = False
    LeapWall.Collidable = False

    Select Case Int(Rnd * 3)
        Case 0 : PlaySound "Bar_datewithdeath", 0, 1
        Case 1 : PlaySound "Bar_meetfate", 0, 1
        Case 2 : PlaySound "Bar_nowyoudie", 0, 1
    End Select

    ' REMOVED AddScore from here to prevent double scoring!

    ' Dynamically show what the upcoming shot will be worth
    Dim NextJackpotValue : NextJackpotValue = LeapJackpot + 10000000
    ShowMysteryDMD "LEAP ATTACK READY", "FOR " & FormatNumber(NextJackpotValue, 0, -1, 0, -1)

    LeapReady = True
    BarbLeanStep = 0
    BarbPrim.RotX = BarbRestRotX
    BarbLeanTimer.Enabled = True
    LeapPromptTimer.Interval = 2000
    LeapPromptTimer.Enabled = True
    LeapSafetyTimer.Enabled = True
End Sub

Sub LeapPromptTimer_Timer()
    If Not LeapReady Then
        LeapPromptTimer.Enabled = False
        Exit Sub
    End If
    ShowMysteryDMD "LEAP ATTACK!", "PRESS MAGNA SAVE!"
    LeapPromptTimer.Interval = 2000
    LeapPromptTimer.Enabled = True
End Sub

Dim LeapJackpot : LeapJackpot = 0

Sub FireLeapAttack()
    If Not LeapReady Then Exit Sub
    LeapReady = False

    ' Increment jackpot by 10 Million, base 10 Million, persists entire game
    LeapJackpot = LeapJackpot + 10000000

    AddScore LeapJackpot
    UpdateDMD2 "LEAP ATTACK!", "+" & FormatNumber(LeapJackpot, 0, -1, 0, -1)

    BarbLeanTimer.Enabled = False
    BarbJumpStep = 0
    BarbJumpTimer.Enabled = True
    LeapPromptTimer.Enabled = False

    SetFlasherColor 1, 255, 180, 0 : SetFlasherColor 2, 255, 180, 0
    SetFlasherColor 3, 255, 180, 0 : SetFlasherColor 4, 255, 180, 0
    SetFlasherColor 5, 255, 180, 0 : SetFlasherColor 6, 255, 180, 0
    FireAllFlashers
    KillFlasherDimTimer.Enabled = False
    KillFlasherDimTimer.Enabled = True
    ResumeAmbushTimer
  ResumeKillStreakTimer

    ArenaUpKick.KickZ 305, 49, 0, 80
    PlaySound "popper_ball", 0, 1, AudioPan(ArenaUpKick), 0, 0, 0, 1, AudioFade(ArenaUpKick)
    LeapWallTimer.Interval = 2000
    LeapWallTimer.Enabled = True

    Select Case Int(Rnd * 3)
        Case 0 : PlaySound "Leap1", 0, 1, AudioPan(ArenaUpKick)
        Case 1 : PlaySound "Leap2", 0, 1, AudioPan(ArenaUpKick)
        Case 2 : PlaySound "Leap3", 0, 1, AudioPan(ArenaUpKick)
    End Select

    UpdateDMDScore
End Sub

Sub LeapWallTimer_Timer()
    LeapWallTimer.Enabled = False
    LeapWall.Collidable = True
End Sub

Sub ArenaUpKick_Timer()
    ' No longer used for auto-fire - kept empty as safety
End Sub

Sub LeapAwardTimer_Timer()
    LeapAwardTimer.Enabled = False
End Sub

Sub KillScoreTrigger_Hit()
    If Not GameActive Then Exit Sub
    If LeapBallID >= 0 And ActiveBall.ID = LeapBallID Then
        LeapBallID = -1
    End If
    BallArenaKillCount = 0
    BallArenaKillScore = 0
    BackstabActive = False
End Sub

Sub AmbushTrigger_Hit()
    If Not GameActive Then Exit Sub
    If ActiveBall.VelY < 0 Then Exit Sub
    BackstabActive = True
    UpdateDMD2 "BACKSTAB!", "DOUBLE DAMAGE"
    If Not CalloutPlaying Then PlaySound "Backstab", 0, 1
'    BackglassPulse 1500, 255, 20, 0
    SetFlasherColor 1, 200, 0, 0
    SetFlasherColor 2, 200, 0, 0
    SetFlasherColor 3, 200, 0, 0
    SetFlasherColor 4, 200, 0, 0
    SetFlasherColor 5, 200, 0, 0
    SetFlasherColor 6, 200, 0, 0
    FireAllFlashers
    KillFlasherDimTimer.Enabled = False
    KillFlasherDimTimer.Enabled = True
End Sub

' MYSTERY LOGIC

Sub MysteryHoldTimer_Timer()
    MysteryHoldTimer.Enabled = False
    MysteryAnimCycleCount = 0
    MysteryAnimPos = Int(Rnd * 17)
    MysteryAnimTimer.Interval = 80
    MysteryAnimTimer.Enabled = True
End Sub

Sub MysteryCaptureTrigger_Hit()
    If Not GameActive Then Exit Sub
    ActiveBall.X = MysteryKick.X
    ActiveBall.Y = MysteryKick.Y
    ActiveBall.VelX = 0
    ActiveBall.VelY = 0
    ActiveBall.VelZ = 0
End Sub

Sub MysteryKick_Hit()
    If Not GameActive Then Exit Sub
    If Not MysteryReady Then
        MysteryKickHolding = False
    ResumeAmbushTimer
    ResumeKillStreakTimer
    MysteryKick.Kick 180, 15
        PlaySound "popper_ball", 0, 1
        Exit Sub
    End If
  MysteryKickHolding = True
  PauseKillStreakTimer
    MysteryReady = False
    MysteryActive = True
    If AmbushActive Then
        AmbushHurryTimer.Enabled = False
        AmbushHoldTimer.Enabled = False
    End If
  If InstantKillActive Then InstantKillTimer.Enabled = False  ' ← ADD
    SetFlasherColor 3, 180, 0, 255
    SetFlasherColor 4, 180, 0, 255
    FireFlasher 3
    FireFlasher 4
'    MysteryKickLight.State = 0
  QuestLight4.State = 0
    If GetPartyCount() <= 1 Then
        Dim safePool(19)
        Dim sp : sp = 0
        Dim ri
        For ri = 0 To 21
            If ri <> 15 And ri <> 16 Then
                safePool(sp) = ri
                sp = sp + 1
            End If
        Next
        MysteryRoll = safePool(Int(Rnd * sp))
    Else
        MysteryRoll = Int(Rnd * 22)
    End If
    MysteryAnimCycleCount = 0
    PlaySound "waypointignite2", 0, 1
    DuckAudio
    PlayCallout "izualaction", 10000
    ShowMysteryDMD "MYSTERY Chest!", ""
    StartGIEvent GI_MODE_MYSTERY
    MysteryHoldTimer.Interval = 1500
    MysteryHoldTimer.Enabled = True
End Sub

Function MysteryRollName(idx)
    Select Case idx
        Case 0, 1, 2    : MysteryRollName = "** WHIRLWIND READY **"
        Case 3, 4, 5    : MysteryRollName = "** LIGHT RUNEWORD **"
        Case 6, 7       : MysteryRollName = "** CUBE READY **"
        Case 8, 9, 10   : MysteryRollName = "** 10 MILLION **"
        Case 11         : MysteryRollName = "** CRIT BOSS **"
        Case 12, 13, 14 : MysteryRollName = "** UNIQUE ITEM **"
        Case 15, 16     : MysteryRollName = "** PARTY MULTIBALL **"
        Case 17, 18     : MysteryRollName = "** SHOUT BONUS HELD **"
        Case 19         : MysteryRollName = "** MAX RUNE QUALITY **"
        Case 20, 21     : MysteryRollName = "** LOOT HELD **"
    End Select
End Function

Dim MysteryAnimPos : MysteryAnimPos = 0

Sub MysteryAnimTimer_Timer()
    MysteryAnimCycleCount = MysteryAnimCycleCount + 1

    Select Case MysteryAnimCycleCount
        Case 15 : MysteryAnimTimer.Interval = 100
        Case 17 : MysteryAnimTimer.Interval = 130
        Case 19 : MysteryAnimTimer.Interval = 170
        Case 21 : MysteryAnimTimer.Interval = 220
        Case 23 : MysteryAnimTimer.Interval = 280
        Case 25 : MysteryAnimTimer.Interval = 350
        Case 27 : MysteryAnimTimer.Interval = 430
        Case 28 : MysteryAnimTimer.Interval = 520
        Case 29 : MysteryAnimTimer.Interval = 620
        Case 30 : MysteryAnimTimer.Interval = 730
    End Select

    If MysteryAnimCycleCount <= 30 Then
        MysteryAnimPos = (MysteryAnimPos + 1) Mod 22
        ShowMysteryDMD MysteryRollName(MysteryAnimPos), "MYSTERY CHEST"
    Else
        MysteryAnimTimer.Enabled = False
        ShowMysteryDMD MysteryRollName(MysteryRoll), ">> AWARDED! <<"
        MysteryRevealTimer.Interval = 2000
        MysteryRevealTimer.Enabled = True
    End If
End Sub

Sub MysteryRevealTimer_Timer()
    MysteryRevealTimer.Enabled = False
    MysteryActive = False
    If PoisonTickActive Then PoisonTimer.Enabled = True
    If InstantKillActive Then InstantKillTimer.Enabled = True
    If AmbushActive Then
        StartGIEvent GI_MODE_AMBUSH
        AmbushHurryTimer.Enabled = True
        AmbushHoldTimer.Enabled = True
    Else
        GIEventMode = GI_MODE_NONE
        GIEventTimer.Enabled = False
        RestoreAllGIRows
    End If
    DimFlasher 3
    DimFlasher 4
    MysteryKickHolding = False
  ResumeAmbushTimer
  ResumeKillStreakTimer
    MysteryKick.Kick 180, 15
    PlaySound "popper_ball", 0, 1
'    If CurrentAct = 6 And PortalActive Then
'        PortalTime = PortalTime + 5000
'        If PortalTime > PortalCap() Then PortalTime = PortalCap()
'    ElseIf CurrentAct < 6 Then
'        AddPortalSurge 5
'    End If
    MysteryRampCloseTimer.Interval = 3000
    MysteryRampCloseTimer.Enabled = True
    MysteryFillLootSlots
    Select Case MysteryRoll
        Case 0, 1, 2    : MysteryWhirlwindReady
        Case 3, 4, 5    : MysteryLightRune
        Case 6, 7       : MysteryQueueCube
        Case 8, 9, 10   : MysteryScoreBomb
        Case 11         : MysteryBossWeaken
        Case 12, 13, 14 : MysteryUniqueGear
        Case 15, 16     : MysteryPartyMultiball
        Case 17, 18     : MysteryHoldBonus
        Case 19         : MysteryRuneSockted
        Case 20, 21     : MysteryGearRetention
    End Select
End Sub

Dim GearRetentionActive : GearRetentionActive = False

' Gear retention across balls (1 item per ally unlocked, max 6)
Dim RetainedGear(5)
Dim RetainedGearCount : RetainedGearCount = 0

Sub MysteryRampCloseTimer_Timer()
    MysteryRampCloseTimer.Enabled = False
    ChestRampUp
End Sub

Sub MysteryHoldBonus()
    PlaySound "waypointignite2", 0, 1
    If ShoutMultiplierHeld Or ShoutHoldQueueCount > 0 Then
        If ShoutHoldQueueCount < 1 Then
            ShoutHoldQueueCount = ShoutHoldQueueCount + 1
      PlayCallout "mystery_shouthold", 2000
            ShowMessage "SHOUT QUEUED! x" & ShoutMultiplier & " BANKED"
        Else
            ShowMessage "SHOUT MAXED! x" & ShoutMultiplier & " ACTIVE"
        End If
    Else
        ShoutMultiplierHeld = True
    PlayCallout "mystery_shouthold", 2000
        UpdateDMD2 "SHOUT HELD", "NEXT BALL"
    End If
End Sub

Sub MysteryScoreBomb()
    AddScore 10000000
    ShowMessage "10,000,000!"
    PlaySound "waypointignite2", 0, 1
  PlayCallout "mystery_scorebomb", 2000
End Sub

Sub MysteryGearRetention()
    GearRetentionActive = True
    PlaySound "waypointignite2", 0, 1
    PlayCallout "LootHeld", 2000
    ShowMessage "LOOT HELD!"
    StartGIEvent GI_MODE_UNIQUE
End Sub

Sub SaveRetainedGear()
    ' Save the best N gear tiers before clearing, where N = ally count (party minus Barb)
    RetainedGearCount = 0
    Dim numToRetain : numToRetain = GetPartyCount() - 1
    If numToRetain <= 0 Then Exit Sub
    If numToRetain > 6 Then numToRetain = 6

    ' Collect all equipped slot tiers
    Dim collected(6), numCollected : numCollected = 0
    Dim gi
    For gi = 0 To 6
        If GearSlots(gi) >= 0 Then
            collected(numCollected) = GearSlots(gi)
            numCollected = numCollected + 1
        End If
    Next
    If numCollected = 0 Then Exit Sub

    ' Bubble sort descending (Unique=2 > Rare=1 > Magic=0)
    Dim si, sj, tmp
    For si = 0 To numCollected - 2
        For sj = 0 To numCollected - 2 - si
            If collected(sj) < collected(sj + 1) Then
                tmp = collected(sj) : collected(sj) = collected(sj + 1) : collected(sj + 1) = tmp
            End If
        Next
    Next

    ' Store top N
    Dim rk
    For rk = 0 To numToRetain - 1
        If rk >= numCollected Then Exit For
        RetainedGear(RetainedGearCount) = collected(rk)
        RetainedGearCount = RetainedGearCount + 1
    Next
End Sub

Sub RestoreRetainedGear()
    ' Restore saved gear tiers into slots 0 to N-1 after the gear clear
    If RetainedGearCount = 0 Then Exit Sub
    Dim ri
    For ri = 0 To RetainedGearCount - 1
        GearSlots(ri) = RetainedGear(ri)
        LightGearSlot ri, RetainedGear(ri)
    Next
    RetainedGearCount = 0
End Sub

Sub MysteryFillLootSlots()
    Dim gs, filled : filled = 0
    For gs = 0 To 6
        If filled >= 1 Then Exit For
        If GearSlots(gs) = -1 Then
            Dim lt : lt = RollGearTier()
            If lt < GEAR_RARE Then lt = GEAR_RARE
            GearSlots(gs) = lt
            LightGearSlot gs, lt
            BlinkGearSlot gs
            PlayGearSlotSound gs
            filled = filled + 1
        End If
    Next
    If filled > 0 Then
        ShowMessage "MYSTERY LOOT! +1 ITEM"
    End If
End Sub

Sub MysteryRuneSockted()
    If RuneQuality = 4 Then
        AddScore 5000000
        UpdateDMD2 "RUNES ALREADY MAX!", "+5,000,000"
        PlaySound "cairnsuccess", 0, 1
        Exit Sub
    End If
    RuneQuality = 4
    PlaySound "waypointignite2", 0, 0.8
    Select Case Int(Rnd * 3)
        Case 0 : PlaySound "grunt203", 0, 1
        Case 1 : PlaySound "grunt202", 0, 1
        Case 2 : PlaySound "grunt201", 0, 1
    End Select
    PlaySound "malus", 0, 1
    SetFlasherColor 2, 255, 180, 0
    FireFlasher 2
    RuneFlasherDimTimer.Enabled = False
    RuneFlasherDimTimer.Enabled = True
  PlayCallout "mystery_runesocketed", 2000
    ShowMessage "RUNES:+4  MAX QUALITY!"
End Sub

Sub MysteryWhirlwindReady()
    PlayCallout "mystery_whirlwindready", 2000
    StartGIEvent GI_MODE_FRENZY
    If WhirlwindActive Then
        ' storm spinning — bank a full charge so the NEXT ramp shot relaunches
        RampAoeCount = RampAoeTarget - 1
        UpdateDMD2 "WHIRLWIND READY!", "NEXT RAMP SHOT"
    ElseIf RampAoeCount >= RampAoeTarget - 1 Then
        ' already one hit from firing — nothing to charge, pay the fallback
        AddScore 5000000
        UpdateDMD2 "ALREADY CHARGED!", "+5,000,000"
    Else
        RampAoeCount = RampAoeTarget - 1
        UpdateRampAoeLight
        UpdateDMD2 "WHIRLWIND READY!", "SHOOT THE RAMP!"
    End If
End Sub

Sub MysteryLightRune()
    PlaySound "waypointignite2", 0, 1
    If RuneWordMultiballRunning Then
        ' Light all runes during RWMB — player can re-arm scoop
        RuneHitR  = True : RuneLightR.BlinkInterval  = 200 : RuneLightR.State  = 2
        RuneHitU  = True : RuneLightU.BlinkInterval  = 200 : RuneLightU.State  = 2
        RuneHitN  = True : RuneLightN.BlinkInterval  = 200 : RuneLightN.State  = 2
        RuneHitE  = True : RuneLightE.BlinkInterval  = 200 : RuneLightE.State  = 2
        RuneHitW  = True : RuneLightW.BlinkInterval  = 200 : RuneLightW.State  = 2
        RuneHitO  = True : RuneLightO.BlinkInterval  = 200 : RuneLightO.State  = 2
        RuneHitR2 = True : RuneLightR2.BlinkInterval = 200 : RuneLightR2.State = 2
        RuneHitD  = True : RuneLightD.BlinkInterval  = 200 : RuneLightD.State  = 2
        RuneWordReady = True
        RuneKickLight.BlinkInterval = 200
        RuneKickLight.BlinkPattern = "10"
        RuneKickLight.State = 2
        ShowMessage "RUNEWORD READY!"
        PlayCallout "mystery_lightrune", 3000
        Exit Sub
    End If
    If RuneWordReady Then
        ' Already armed — fire multiball directly as bonus
        RuneWordReady = False
        RuneKickLight.State = 0
        ResetRuneTargets
        RuneQuality = 0
        GetRandomRuneWord RuneWordName, RuneWordSockets
        RuneWordSockets = GetSocketsFromQuality()
    GetRandomRuneWord RuneWordName, RuneWordSockets
        ShowMessage "RUNEWORD! " & RuneWordName
        RuneWordMultiballRunning = True
        BallSaveActive = True
        BallSaveUsed = False
        BallSaveTimer.Enabled = False
        BallSaveTimer.Interval = 11000
        BallSaveTimer.Enabled = True
        BallSaveL.State = 1
        BallSaveL2.State = 1
        RuneWordSpawnCount = 0
        StartGIEvent GI_MODE_RUNEWORD
        PlayCallout "RunewordMultiDiablo", 3000
        RuneWordSpawnTimer.Enabled = True
        Exit Sub
    End If
    ' Light all runes and arm scoop
    RuneHitR  = True : RuneLightR.BlinkInterval  = 200 : RuneLightR.State  = 2
    RuneHitU  = True : RuneLightU.BlinkInterval  = 200 : RuneLightU.State  = 2
    RuneHitN  = True : RuneLightN.BlinkInterval  = 200 : RuneLightN.State  = 2
    RuneHitE  = True : RuneLightE.BlinkInterval  = 200 : RuneLightE.State  = 2
    RuneHitW  = True : RuneLightW.BlinkInterval  = 200 : RuneLightW.State  = 2
    RuneHitO  = True : RuneLightO.BlinkInterval  = 200 : RuneLightO.State  = 2
    RuneHitR2 = True : RuneLightR2.BlinkInterval = 200 : RuneLightR2.State = 2
    RuneHitD  = True : RuneLightD.BlinkInterval  = 200 : RuneLightD.State  = 2
    RuneWordReady = True
    RuneKickLight.BlinkInterval = 200
    RuneKickLight.BlinkPattern = "10"
    RuneKickLight.State = 2
    SetFlasherColor 2, 255, 180, 0
    FireFlasher 2
    RuneFlasherDimTimer.Enabled = False
    RuneFlasherDimTimer.Enabled = True
    PlaySound "rune", 0, 1
  PlayCallout "mystery_lightrune", 3000
    ShowMysteryDMD "RUNEWORD READY!", "SHOOT SCOOP!"
End Sub

Sub MysteryQueueCube()
    PlaySound "waypointignite2", 0, 1
    If TransmuteLevel = 4 Then
        AddScore 5000000
        UpdateDMD2 "CUBE JACKPOT!", "+5,000,000"
        PlaySound "cairnsuccess", 0, 1
        If Not GemWallOpen Then OpenCubeWall
        Exit Sub
    End If
    Dim gemType
    Select Case TransmuteLevel
        Case 0 : gemType = ATTACK_COLD
        Case 1 : gemType = ATTACK_FIRE
        Case 2 : gemType = ATTACK_LIGHTNING
        Case 3 : gemType = ATTACK_POISON
    End Select
    If GemQueueCount < 9 Then
        GemQueue(GemQueueCount) = gemType
        GemQueueCount = GemQueueCount + 1
        OpenCubeWall
        ShowMessage "CUBE READY!"
        PlayCallout "mystery_cube", 2000
        UpdateCubeLight
    Else
        AddScore 5000000
        UpdateDMD2 "CUBE FULL!", "+5,000,000"
        PlaySound "cairnsuccess", 0, 1
    End If
End Sub

Sub MysteryPartyMultiball()
    If PartyMultiballRunning Then
        Dim jackpot : jackpot = (GetPartyCount() - 1) * 1000000
        If jackpot <= 0 Then jackpot = 5000000
        ShowMessage "PARTY BONUS! +" & FormatNumber(jackpot, 0, -1, 0, -1)
        AddScore jackpot
        Dim needed : needed = (GetPartyCount() - 1) - PartySpawnStep
        If needed > 0 Then
            PartySpawnCount = PartySpawnStep + needed
            If Not PartySpawnTimer.Enabled Then
                PartySpawnTimer.Interval = 800
                PartySpawnTimer.Enabled = True
            End If
        End If
        Exit Sub
    End If
    PartySpawnStep = 0
    PartyMultiballRunning = True
    BallSaveActive = True
    BallSaveUsed = False
    BallSaveTimer.Enabled = False
    BallSaveTimer.Interval = 11000
    BallSaveTimer.Enabled = True
    BallSaveL.Color = RGB(174, 0, 0) : BallSaveL.ColorFull = RGB(255, 72, 72) : BallSaveL.State = 1
    BallSaveL2.Color = RGB(174, 0, 0) : BallSaveL2.ColorFull = RGB(255, 72, 72) : BallSaveL2.State = 1
    townportalPF.Visible = True
    PlaySound "portalenter", 0, 1
    StartGIEvent GI_MODE_MYSTERY
    If GetPartyCount() = 1 Then
        PartySpawnCount = 1
        ShowMessage "MERCENARY HIRED!"
        PlayCallout "mercenary_hired", 3000
    Else
        PartySpawnCount = GetPartyCount() - 1
        ShowMessage "PARTY ASSEMBLES!"
        PlayCallout "mystery_partymultiball", 3000
    End If
    PortalSpawnInterval = 800
    PortalOpenTimer.Interval = 3000
        PortalOpenTimer.Enabled = True
        PortalFlashStep = 0
        PortalFlashTimer.Interval = 2000
        PortalFlashTimer.Enabled = True
End Sub

Sub PartySpawnTimer_Timer()
    If PartySpawnStep >= PartySpawnCount Then
        PartySpawnTimer.Enabled = False
        townportalPF.Visible = False
        PlayfieldKicker.Enabled = False
        ShowMessage "PARTY MULTIBALL!"
        UpdateDMD2 "PARTY MULTIBALL!", GetPartyCount() & " HEROES"
        Exit Sub
    End If
    Dim pb : Set pb = PlayfieldKicker.CreateBall
    Select Case ActiveElementType
        Case ATTACK_COLD      : pb.Color = RGB(80, 180, 255)
        Case ATTACK_FIRE      : pb.Color = RGB(255, 60, 0)
        Case ATTACK_LIGHTNING : pb.Color = RGB(196, 196, 8)
        Case ATTACK_POISON    : pb.Color = RGB(0, 200, 50)
        Case Else             : pb.Color = RGB(200, 200, 210)
    End Select
    If Rnd > 0.5 Then PlayfieldKicker.Kick 165, 15 Else PlayfieldKicker.Kick 195, 15
    PlaySound "popper_ball", 0, 1
    PlaySound "portalenter", 0, 1
    PartySpawnStep = PartySpawnStep + 1
End Sub

Sub MysteryBossWeaken()
    PlaySound "waypointignite2", 0, 1
    If BossCritActive Or BossCritQueued Then
        ShowMessage "CRIT ALREADY ARMED!"
    AddScore 5000000
        Exit Sub
    End If
    If BossFightActive Then
        BossCritActive = True
        PlayCallout "mystery_critboss", 2000
        ShowMessage "CRIT BOSS! 2x DMG!"
    Else
        BossCritQueued = True
        PlayCallout "mystery_critboss", 2000
        ShowMessage "NEXT BOSS: 2x DMG!"
    End If
End Sub

Sub MysteryUniqueGear()
    PlaySound "waypointignite2", 0, 1
    If AllUniquesEquipped() Then
        AddScore 5000000
        UpdateDMD2 "INVENTORY FULL!", "+5,000,000"
        PlaySound "cairnsuccess", 0, 1
        Exit Sub
    End If
  PlayCallout "UniqueEquipped", 2000
    AssignGearToSlot GEAR_UNIQUE
    ShowMessage "UNIQUE ITEM FOUND!"
End Sub

Sub InstantKillTimer_Timer()
    InstantKillTimer_step = InstantKillTimer_step + 1
    If InstantKillTimer_step >= 15 Then
        InstantKillTimer.Enabled = False
        InstantKillActive = False
        GIEventTimer.Enabled = False
        If AmbushActive Then
            StartGIEvent GI_MODE_AMBUSH
        Else
            GIEventMode = GI_MODE_NONE
            RestoreAllGIRows
        End If
        Dim ikb
        For ikb = 1 To 5
            FlBumperColor(ikb) = "red"
            FlInitBumper ikb, "red"
            FlBumperFadeTarget(ikb) = 0
            FlBumperFadeActual(ikb) = 0
            FlFadeBumper ikb, 0
        Next
        UpdateDMD2 "INSTANT KILL OVER!", "+" & FormatNumber(AuraKillScore, 0, -1, 0, -1) & " SCORED"
        AuraEndTimer.Interval = 2000
        AuraEndTimer.Enabled = True
    Else
        UpdateDMD2 "HIT THE BUMPERS!", "INSTANT KILL: " & (15 - InstantKillTimer_step) & "SEC"
        Select Case InstantKillTimer_step
            Case 10 : PlayCallout "Five", 1000
            Case 11 : PlayCallout "Four", 2000
            Case 12 : PlayCallout "Three", 1000
            Case 13 : PlayCallout "Two", 1000
            Case 14 : PlayCallout "One", 1000
        End Select
    End If
End Sub

Sub AuraEndTimer_Timer()
    AuraEndTimer.Enabled = False
    AuraKillScore = 0
    UpdateDMDScore
End Sub

' END MYSTERY LOGIC


Sub AddScore(points)
    Score = Score + points
    UpdateDMDScore
    Do While ExtraBallMilestoneIdx <= 3 And GameActive
        If Score >= ExtraBallMilestones(ExtraBallMilestoneIdx) Then
            ExtraBallMilestoneIdx = ExtraBallMilestoneIdx + 1
            ExtraBallPending = ExtraBallPending + 1
            If Not BonusActive Then
                ShowBigMessage "EXTRA BALL!"
                PlayCallout "ExtraBall", 500
                StartGIEvent GI_MODE_EXTRABALL
            Else
                ExtraBallEoBPending = ExtraBallEoBPending + 1
            End If
        Else
            Exit Do
        End If
    Loop
End Sub

Sub AddScoreWithMultiplier(points)
    ' Legacy sub - calls AddScore directly (gear system now applies via damage, not score mult)
    AddScore points
    UpdateDMDScore
End Sub

'TREASURE CHEST
'
'Sub TChestKick_Hit()
'    If LootActive(0) And LootPending(0) = LOOT_GEAR Then
'        Score = Score + BossLootGoldAmount
'        PlaySound "gold", 0, 1
'        ShowMessagePriority FormatNumber(BossLootGoldAmount, 0, -1, 0, -1) & " GOLD!", 3
'        UpdateDMDScore
'        CollectLoot 0
'    Else
'        ShowMessage "GOLD FOUND!"
'        AddScore 50000
'    End If
'    PlaySound "fx_kicker_catch"
'    TreasureChestTimer.Enabled = True
'End Sub

'Sub TreasureChestTimer_Timer()
'    TChestKick.KickZ 180, 10, 0, 30
'    PlaySound "popper_ball", 0, 1
'    TreasureChestTimer.Enabled = False
'    ChestRampCloseTimer.Enabled = True
'End Sub

'Sub ChestRampCloseTimer_Timer()
'    ChestRampUp
'    ChestRampCloseTimer.Enabled = False
'End Sub


'*****************************************
'      TREASURE CHEST UP/DOWN RAMP
'*****************************************
Dim ChestRampActive : ChestRampActive = False
Dim ChestRampVel    : ChestRampVel    = 0
Dim ChestRampState  : ChestRampState  = 0  ' 0=idle 1=falling 2=rebound 3=rising
Dim ChestRampSilent : ChestRampSilent = False  ' one-shot: suppress next ramp sound

Const ChestRampRest   = 0      ' RotX at rest (up position)
Const ChestRampLand   = -12  ' RotX at playfield (down position)
Const ChestRampAccel  = 0.003    ' acceleration per tick (tune this)
Const ChestRampBounce = 0.15    ' how much it rebounds (tune this)

Sub ChestRampDown()
    Dim playIt : playIt = Not ChestRampSilent : ChestRampSilent = False
    ChestRampActive = True
    ChestRampState = 1
    ChestRampVel = 0.01
    ChestRamp.RotX = ChestRampRest
    ChestRampPhysics.Collidable = True
    ChestRampMove.Enabled = True
    If playIt Then PlaySound "xtowngate"
End Sub

Sub ChestRampUp()
    Dim playIt : playIt = Not ChestRampSilent : ChestRampSilent = False
    If ChestRamp.RotX >= ChestRampRest Then   ' already up: no move, no sound
        ChestRampActive = False : ChestRampState = 0
        ChestRampPhysics.Collidable = False : ChestRampMove.Enabled = False
        Exit Sub
    End If
    ChestRampActive = False
    ChestRampState = 3
    ChestRampPhysics.Collidable = False
    ChestRampMove.Enabled = True
    If playIt Then PlaySound "xtowngatereverse"
End Sub

Sub ChestRampMove_Timer()
    Select Case ChestRampState
        Case 1  ' Falling — accelerate toward -15
            ChestRampVel = ChestRampVel + ChestRampAccel
            ChestRamp.RotX = ChestRamp.RotX - ChestRampVel
            If ChestRamp.RotX <= ChestRampLand Then
                ChestRamp.RotX = ChestRampLand
                ChestRampVel = ChestRampVel * ChestRampBounce
                ChestRampState = 2
            End If

        Case 2  ' Rebound — short bounce then settle back down
    ChestRamp.RotX = ChestRamp.RotX + ChestRampVel
    ChestRampVel = ChestRampVel - (ChestRampAccel * 0.8)
    If ChestRampVel <= 0 Then
        ChestRampState = 4  ' settle state
        ChestRampVel = 0.02
    End If


        Case 3  ' Rising back to rest — linear
            ChestRamp.RotX = ChestRamp.RotX + 0.1
            If ChestRamp.RotX >= ChestRampRest Then
                ChestRamp.RotX = ChestRampRest
                ChestRampState = 0
                ChestRampMove.Enabled = False
            End If
    Case 4  ' Settle — ease back down to rest
    ChestRamp.RotX = ChestRamp.RotX - ChestRampVel
    If ChestRamp.RotX <= ChestRampLand Then
        ChestRamp.RotX = ChestRampLand
        ChestRampVel = 0
        ChestRampState = 0
        ChestRampMove.Enabled = False
    End If
    End Select
End Sub
'*****************************************
'      BONUS ATTACK BALL SYSTEM
'*****************************************

Sub SetAllBallsElement(elementType)
    ActiveElementType = elementType
    Dim allB : allB = GetBalls()
    Dim b
    For Each b In allB
        If b.ID <> CapBallID And b.ID <> CapBall2ID Then
            Select Case elementType
                Case ATTACK_COLD      : b.Color = RGB(80, 180, 255)
                Case ATTACK_FIRE      : b.Color = RGB(255, 60, 0)
                Case ATTACK_LIGHTNING : b.Color = RGB(196, 196, 8)
                Case ATTACK_POISON    : b.Color = RGB(0, 200, 50)
                Case -1               : b.Color = RGB(200, 200, 210)
            End Select
        End If
    Next
End Sub

Function GetActiveBonusBallType()
    GetActiveBonusBallType = ActiveElementType
End Function

Sub PlayBonusBallSpawnSound(ballType)
    Select Case ballType
        Case ATTACK_FIRE      : PlaySound "sizzle3", 0, 1
        Case ATTACK_COLD      : PlaySound "shatter3", 0, 1
        Case ATTACK_POISON    : PlaySound "poisoned", 0, 1
        Case ATTACK_LIGHTNING : PlaySound "static1", 0, 1
    End Select
End Sub

Sub PlayBonusBallHitSound(ballType)
    Select Case ballType
        Case ATTACK_FIRE
            Select Case Int(Rnd * 3)
                Case 0 : PlaySound "singe1", 0, 1
                Case 1 : PlaySound "singe2", 0, 1
                Case 2 : PlaySound "singe3", 0, 1
            End Select
        Case ATTACK_COLD
            Select Case Int(Rnd * 3)
                Case 0 : PlaySound "cold1", 0, 1
                Case 1 : PlaySound "cold2", 0, 1
                Case 2 : PlaySound "cold3", 0, 1
            End Select
        Case ATTACK_POISON
            Select Case Int(Rnd * 3)
                Case 0 : PlaySound "green1", 0, 1
                Case 1 : PlaySound "green2", 0, 1
                Case 2 : PlaySound "green3", 0, 1
            End Select
        Case ATTACK_LIGHTNING
            Select Case Int(Rnd * 3)
                Case 0 : PlaySound "zap1", 0, 1
                Case 1 : PlaySound "zap2", 0, 1
                Case 2 : PlaySound "zap3", 0, 1
            End Select
    End Select
End Sub

Sub ChillCheckTimer_Timer()
If Not GameActive Then ChillCheckTimer.Enabled = False : Exit Sub
    Dim cs
    For cs = 0 To 4
        If BumperChilled(cs) Then
            If GameTime - BumperChillTime(cs) > 5000 Then
                BumperChilled(cs) = False
        FlBumperColor(cs + 1) = "red"
        FlInitBumper cs + 1, "red"
        FlBumperFadeTarget(cs + 1) = 0
        FlFadeBumper cs + 1, 0
            End If
        End If
    Next
End Sub

Sub BonusBallHitBumper(slot, bumperObj, ballType)
    If Not BumperActive(slot) Then Exit Sub
    Dim totalDamage : totalDamage = 1 + GetGearDamage()
    If BackstabActive Then totalDamage = totalDamage * 2   ' ← elemental balls now get Backstab doubling too
'    Dim partyBonus : partyBonus = 0
    Select Case ballType
        Case ATTACK_FIRE      : SetFlasherColor 3, 255, 60, 0  : SetFlasherColor 4, 255, 60, 0
        Case ATTACK_COLD      : SetFlasherColor 3, 80, 180, 255 : SetFlasherColor 4, 80, 180, 255
        Case ATTACK_POISON    : SetFlasherColor 3, 0, 200, 50   : SetFlasherColor 4, 0, 200, 50
        Case ATTACK_LIGHTNING : SetFlasherColor 3, 196, 196, 8  : SetFlasherColor 4, 196, 196, 8
    End Select
    FireFlasher 3
    FireFlasher 4
    Select Case ballType
        Case ATTACK_FIRE
            BumperHP(slot) = BumperHP(slot) - totalDamage
            PlayBonusBallHitSound ballType
            FireFlashBumper slot
            If BumperHP(slot) > 0 Then
                FireBurnSlot = slot
                FireBurnTimer.Enabled = True   ' ← guaranteed burn now, no RNG gate
            End If
        Case ATTACK_COLD
    ColdFlashBumper slot
    If BumperChilled(slot) Then
        BumperChilled(slot) = False
        FlBumperColor(slot + 1) = "red"
        FlInitBumper slot + 1, "red"
        BumperHP(slot) = 0
        PlaySound "Shatter" & (Int(Rnd * 3) + 1), 0, 1
        ShowMessage "SHATTER!"
    Else
        BumperHP(slot) = BumperHP(slot) - totalDamage
        PlayBonusBallHitSound ballType
        If BumperHP(slot) > 0 Then  ' only chill if still alive
            If Rnd < 0.10 Then
                BumperChilled(slot) = True
                BumperChillTime(slot) = GameTime
                PlaySound "frozenarmor", 0, 1
                ShowMessage "CHILLED!"
                FlBumperColor(slot + 1) = "blue"
                FlInitBumper slot + 1, "blue"
                ChillCheckTimer.Enabled = True
            End If
        End If
    End If
        Case ATTACK_POISON
    PoisonTickFlashBumper slot
    BumperHP(slot) = BumperHP(slot) - totalDamage
    PlayBonusBallHitSound ballType
    If BumperHP(slot) > 0 Then
        BumperPoisoned(slot) = True
        FlBumperColor(slot + 1) = "green"
        FlInitBumper slot + 1, "green"
        If Not PoisonTickActive Then
            PoisonTickActive = True
            PoisonTimer.Enabled = True
        End If
    End If
        Case ATTACK_LIGHTNING
            BumperHP(slot) = BumperHP(slot) - totalDamage
            PlayBonusBallHitSound ballType
            LightningFlashAllBumpers
            ChainLightning slot
    End Select
    If BumperHP(slot) <= 0 Then
        FlBumperColor(slot + 1) = "red"
        FlInitBumper slot + 1, "red"
        If BumperChilled(slot) Then BumperChilled(slot) = False
        Select Case ballType
            Case ATTACK_FIRE      : SetFlasherColor 3, 255, 60, 0   : SetFlasherColor 4, 255, 60, 0
            Case ATTACK_COLD      : SetFlasherColor 3, 80, 180, 255  : SetFlasherColor 4, 80, 180, 255
            Case ATTACK_POISON    : SetFlasherColor 3, 0, 200, 50    : SetFlasherColor 4, 0, 200, 50
            Case ATTACK_LIGHTNING : SetFlasherColor 3, 196, 196, 8   : SetFlasherColor 4, 196, 196, 8
        End Select
        FireFlasher 3
        FireFlasher 4
        KillFlasherDimTimer.Enabled = False
        KillFlasherDimTimer.Enabled = True
        KillBumper slot
    End If
End Sub

Sub FireBurnTimer_Timer()
    FireBurnTimer.Enabled = False
    If FireBurnSlot < 0 Then Exit Sub
    If Not BumperActive(FireBurnSlot) Then
        FireBurnSlot = -1
        Exit Sub
    End If
    PlaySound "sizzle1", 0, 1
    BumperBurning(FireBurnSlot) = True
    FireFlashBumper FireBurnSlot
    SetFlasherColor 3, 255, 60, 0 : SetFlasherColor 4, 255, 60, 0
    FireFlasher 3 : FireFlasher 4
    KillFlasherDimTimer.Enabled = False
    KillFlasherDimTimer.Enabled = True
    BumperHP(FireBurnSlot) = BumperHP(FireBurnSlot) - 1
    If BumperHP(FireBurnSlot) <= 0 Then
        BumperBurning(FireBurnSlot) = False
        KillBumper FireBurnSlot
    Else
        BumperBurning(FireBurnSlot) = False
    End If
    FireBurnSlot = -1
End Sub

Sub ChainLightning(sourceSlot)
    SetFlasherColor 3, 196, 196, 8 : SetFlasherColor 4, 196, 196, 8
    FireFlasher 3 : FireFlasher 4
    KillFlasherDimTimer.Enabled = False
    KillFlasherDimTimer.Enabled = True
    Dim s
    For s = 0 To 4
        If s <> sourceSlot And BumperActive(s) Then
            BumperHP(s) = BumperHP(s) - 1
            If BumperHP(s) <= 0 Then
                KillBumper s
            End If
        End If
    Next
End Sub

Sub ColdFlashBumper(slot)
    Dim n : n = slot + 1
    If Not FlBumperActive(n) Then Exit Sub
    FlBumperSmallLight(n).Color     = RGB(80, 180, 255)
    FlBumperSmallLight(n).ColorFull = RGB(80, 180, 255)
    FlbumperBigLight(n).Color       = RGB(80, 180, 255)
    FlbumperBigLight(n).ColorFull   = RGB(80, 180, 255)
    FlBumperFadeTarget(n) = 1
    Eval("Bumper" & n & "A").TimerEnabled = True
End Sub

Sub LightningFlashBumper(slot)
    Dim n : n = slot + 1
    If Not FlBumperActive(n) Then Exit Sub
    FlBumperSmallLight(n).Color     = RGB(196, 196, 8)
    FlBumperSmallLight(n).ColorFull = RGB(196, 196, 8)
    FlbumperBigLight(n).Color       = RGB(196, 196, 8)
    FlbumperBigLight(n).ColorFull   = RGB(196, 196, 8)
    FlBumperFadeTarget(n) = 1
    Eval("Bumper" & n & "A").TimerEnabled = True
End Sub

Sub LightningFlashAllBumpers()
    Dim n
    For n = 1 To 5
        If FlBumperActive(n) Then
            FlBumperSmallLight(n).Color     = RGB(196, 196, 8)
            FlBumperSmallLight(n).ColorFull = RGB(196, 196, 8)
            FlbumperBigLight(n).Color       = RGB(196, 196, 8)
            FlbumperBigLight(n).ColorFull   = RGB(196, 196, 8)
            FlBumperFadeTarget(n) = 1
            Eval("Bumper" & n & "A").TimerEnabled = True
        End If
    Next
End Sub

Sub FireFlashBumper(slot)
    Dim n : n = slot + 1
    If Not FlBumperActive(n) Then Exit Sub
    FlBumperSmallLight(n).Color     = RGB(255, 60, 0)
    FlBumperSmallLight(n).ColorFull = RGB(255, 60, 0)
    FlbumperBigLight(n).Color       = RGB(255, 60, 0)
    FlbumperBigLight(n).ColorFull   = RGB(255, 60, 0)
    FlBumperFadeTarget(n) = 1
    Eval("Bumper" & n & "A").TimerEnabled = True
End Sub

Sub PoisonTickFlashBumper(slot)
    Dim n : n = slot + 1
    If Not FlBumperActive(n) Then Exit Sub
    FlBumperSmallLight(n).Color     = RGB(0, 200, 50)
    FlBumperSmallLight(n).ColorFull = RGB(0, 200, 50)
    FlbumperBigLight(n).Color       = RGB(0, 200, 50)
    FlbumperBigLight(n).ColorFull   = RGB(0, 200, 50)
    FlBumperFadeTarget(n) = 1
    Eval("Bumper" & n & "A").TimerEnabled = True
End Sub

Sub PoisonTimer_Timer()
    If MysteryActive Then Exit Sub
    If Not PoisonTickActive Or ActiveElementType <> ATTACK_POISON Then
        PoisonTimer.Enabled = False
        PoisonTickActive = False
        Exit Sub
    End If
  Dim s
    Dim anyPoisoned : anyPoisoned = False
    For s = 0 To 4
        If BumperActive(s) And BumperPoisoned(s) Then
            BumperHP(s) = BumperHP(s) - 1
           If BumperHP(s) <= 0 Then
                FlBumperColor(s + 1) = "red"
                FlInitBumper s + 1, "red"
                FlBumperFadeTarget(s + 1) = 0
                FlBumperFadeActual(s + 1) = 0
                FlFadeBumper s + 1, 0
                KillBumper s
            Else
                anyPoisoned = True
                PoisonTickFlashBumper s
            End If
        End If
    Next
    If anyPoisoned Then
        SetFlasherColor 3, 0, 200, 50 : SetFlasherColor 4, 0, 200, 50
        FireFlasher 3 : FireFlasher 4
        KillFlasherDimTimer.Enabled = False
        KillFlasherDimTimer.Enabled = True
    End If
    If Not anyPoisoned Then
    PoisonTickActive = False
    PoisonTimer.Enabled = False
    Dim pc
    For pc = 0 To 4
      If Not BumperPoisoned(pc) Then
        FlBumperColor(pc + 1) = "red"
        FlInitBumper pc + 1, "red"
        FlBumperFadeTarget(pc + 1) = 0
        FlBumperFadeActual(pc + 1) = 0
      End If
    Next
  End If
End Sub

'*****************************************
'      HORADRIC CUBE / GEM SYSTEM
'*****************************************

Dim TransmuteLevel : TransmuteLevel = 0
Dim CubeLightCycleStep : CubeLightCycleStep = 0

Sub CubeLight_animate     : p16.BlendDisableLighting  = 200 * (CubeLight.GetInPlayIntensity     / CubeLight.Intensity)     : End Sub

Const CubeLidOpenMs = 2000   ' lid open/close time, ms — lower = faster
Const CubeLidTickMs = 10
Dim CubeLidStep
Dim CubeLidOpenY : CubeLidOpenY = CubeLid.TransY + 50
Dim CubeLidClosedY : CubeLidClosedY = CubeLid.TransY
Dim CubeLidTargetY : CubeLidTargetY = CubeLid.TransY

Dim CubeJackpotPending : CubeJackpotPending = False

' --- SPINNER CUBE CRANK ---
Const CubeSpinTarget = 20
Dim LidPerSpin                ' computed in Table1_Init from live lid travel
Dim CubeSpinCount : CubeSpinCount = 0
Dim CrankSndOn    : CrankSndOn    = False

Sub CubeWalls_Hit(idx)
    PlaySoundAtLevelActiveBall "Rubber_" & Int(Rnd * 9) + 1, Vol(ActiveBall) * RubberWeakSoundFactor
End Sub

Sub OpenCubeWall()
    GemWallOpen = True
    CubeLid.Collidable = False
    Wall015.Collidable = False
    CubeGate.Collidable = False
    CubeKick.Enabled = True
    If CubeLid.TransY < CubeLidOpenY Then
        CubeLidTargetY = CubeLidOpenY
        CubeLidTimer.Enabled = True
        StopSound "hellbrazierloop2"
        PlaySound "hellbrazierloop2", 0, 1
    End If
    UpdateCubeLight
End Sub

Sub CloseCubeWall()
    GemWallOpen = False
    CubeGate.Collidable = True
    CubeKick.Enabled = False
    If CubeLid.TransY > CubeLidClosedY Then
        CubeLidTargetY = CubeLidClosedY
        CubeLidTimer.Enabled = True
        StopSound "hellbrazierloop2"
        PlaySound "hellbrazierloop2", 0, 1
    End If
    UpdateCubeLight
End Sub

Sub CubeLidTimer_Timer()
    If CubeLid.TransY < CubeLidTargetY Then
        CubeLid.TransY = CubeLid.TransY + CubeLidStep
        If CubeLid.TransY >= CubeLidTargetY Then
            CubeLid.TransY = CubeLidTargetY
            CubeLidTimer.Enabled = False
            StopSound "hellbrazierloop2"
            If CubeLidTargetY = CubeLidClosedY Then
                CubeLid.Collidable = True
                Wall015.Collidable = True
            End If
        End If
    ElseIf CubeLid.TransY > CubeLidTargetY Then
        CubeLid.TransY = CubeLid.TransY - CubeLidStep
        If CubeLid.TransY <= CubeLidTargetY Then
            CubeLid.TransY = CubeLidTargetY
            CubeLidTimer.Enabled = False
            If CubeLidTargetY = CubeLidClosedY Then
                CubeLid.Collidable = True
                Wall015.Collidable = True
            End If
        End If
    Else
        CubeLidTimer.Enabled = False
    End If
End Sub

Sub CrankCubeLid()
    If Not GameActive Then Exit Sub
    If GemWallOpen Then Exit Sub                 ' already ready/open — wait for consume
    If CubeLid.TransY >= CubeLidOpenY Then Exit Sub

    CubeSpinCount = CubeSpinCount + 1

    CubeLid.TransY = CubeLid.TransY + LidPerSpin
    If CubeLid.TransY > CubeLidOpenY Then CubeLid.TransY = CubeLidOpenY
    CubeLidTargetY = CubeLid.TransY              ' keep target synced for close path

    If Not CrankSndOn Then
        StopSound "hellbrazierloop2"
        PlaySound "hellbrazierloop2", -1, 1
        CrankSndOn = True
    End If
    CrankSndTimer.Interval = 220
    CrankSndTimer.Enabled = False
    CrankSndTimer.Enabled = True

    If CubeLid.TransY >= CubeLidOpenY Then CrankComplete
End Sub

Sub CrankSndTimer_Timer()                        ' no spin for 220ms -> kill the loop
    CrankSndTimer.Enabled = False
    StopSound "hellbrazierloop2"
    CrankSndOn = False
End Sub

Sub CrankComplete()
    StopSound "hellbrazierloop2"
    CrankSndOn = False
    CrankSndTimer.Enabled = False
    PlaySound "gem", 0, 1
    RollRampGem                                  ' queues element + OpenCubeWall (opens gate, enables kick)
    SetFlasherColor 1, 80, 220, 255
    FireFlasher 1
    CubeLight.State = 2
    CubeSpinCount = 0
End Sub

Sub UpdateCubeLight()
    CubeLight.BlinkInterval = 250
    CubeLight.BlinkPattern = "10"
    Dim isReady : isReady = (GemQueueCount > 0 Or GemWallOpen)
    Select Case TransmuteLevel
        Case 0
            CubeLight.Color = RGB(80, 180, 255)  : CubeLight.ColorFull = RGB(80, 180, 255)
            If isReady Then CubeLight.State = 2 Else CubeLight.State = 1
        Case 1
            CubeLight.Color = RGB(255, 60, 0)    : CubeLight.ColorFull = RGB(255, 60, 0)
            If isReady Then CubeLight.State = 2 Else CubeLight.State = 1
        Case 2
            CubeLight.Color = RGB(196, 196, 8)   : CubeLight.ColorFull = RGB(196, 196, 8)
            If isReady Then CubeLight.State = 2 Else CubeLight.State = 1
        Case 3
            CubeLight.Color = RGB(0, 200, 50)    : CubeLight.ColorFull = RGB(0, 200, 50)
            If isReady Then CubeLight.State = 2 Else CubeLight.State = 1
        Case 4
            CubeLight.BlinkInterval = 100
            CubeLight.Color = RGB(255, 255, 255) : CubeLight.ColorFull = RGB(255, 255, 255)
            CubeLight.State = 2
    End Select
End Sub

Sub CubeWallTimer_Timer()
    CubeWallTimer.Enabled = False
    If Not GemWallOpen Then Wall015.Collidable = True
End Sub

Dim CubeTransmuteActive : CubeTransmuteActive = False

Sub CubeKick_Hit()
    CubeTransmuteActive = True
    PauseAmbushTimer
  PauseKillStreakTimer
    ' Guard mid-transition
    If CubeHoldTimer.Enabled Or CubeSpawnTimer.Enabled Then
        CubeTransmuteActive = False
        ResumeAmbushTimer
    ResumeKillStreakTimer
        cubekick.Kick 195, 40
        CubeWallTimer.Enabled = False
        CubeWallTimer.Enabled = True
        PlaySound "popper_ball", 0, 1
        Exit Sub
    End If
    ' At level 0, no bonus ball should exist yet - eject any that wander in
    If TransmuteLevel = 0 And ActiveElementType >= 0 Then
        CubeTransmuteActive = False
        ResumeAmbushTimer
    ResumeKillStreakTimer
        cubekick.Kick 195, 40
        CubeWallTimer.Enabled = False
        CubeWallTimer.Enabled = True
        PlaySound "popper_ball", 0, 1
        Exit Sub
    End If
    ' Level 4 - Poison complete - jackpot per cube shot
    If TransmuteLevel = 4 Then
        ' Award now while the ball is HELD in the kicker, then eject after a
        ' beat via CubeRespawnHoldTimer's CubeJackpotPending branch.
        AddScore 10000000
        UpdateDMD2 "CUBE JACKPOT!", "+10,000,000"
        PlaySound "cairnsuccess", 0, 1
        CubeJackpotPending = True
        CubeRespawnHoldTimer.Interval = 1200    ' hold beat - tune to taste
        CubeRespawnHoldTimer.Enabled = True
        Exit Sub
    End If
    ' Levels 0-3 - require gem queue
    If Not GemWallOpen Or GemQueueCount <= 0 Then
    CubeTransmuteActive = False
        CloseCubeWall
        ResumeAmbushTimer
    ResumeKillStreakTimer
        cubekick.Kick 195, 40
        CubeWallTimer.Enabled = False
        CubeWallTimer.Enabled = True
        PlaySound "popper_ball", 0, 1
        Exit Sub
    End If
    Dim qi
    If GemQueueCount > 1 Then
        For qi = 0 To GemQueueCount - 2
            GemQueue(qi) = GemQueue(qi + 1)
        Next
    End If
    GemQueueCount = GemQueueCount - 1
    If GemQueueCount < 0 Then GemQueueCount = 0
    GemWallOpen = False
    If TransmuteLevel = 0 Then
        SilverBallID = ActiveBall.ID
        PendingElementType = ATTACK_COLD
        CubeLight.State = 1
        CubeHoldTimer.Enabled = True
    ElseIf TransmuteLevel = 1 Then
        PendingElementType = ATTACK_FIRE
        CubeDestroyCurrentAndUpgrade
    ElseIf TransmuteLevel = 2 Then
        PendingElementType = ATTACK_LIGHTNING
        CubeDestroyCurrentAndUpgrade
    ElseIf TransmuteLevel = 3 Then
        PendingElementType = ATTACK_POISON
        CubeDestroyCurrentAndUpgrade
    End If
End Sub

Sub CubeRespawnHoldTimer_Timer()
    CubeRespawnHoldTimer.Enabled = False
    If CubeJackpotPending Then
        CubeJackpotPending = False
        CubeTransmuteActive = False
        GemWallOpen = False
        GemQueueCount = 0
        CloseCubeWall
        ResumeAmbushTimer
    ResumeKillStreakTimer
        cubekick.Kick 195, 40
        CubeWallTimer.Enabled = False
        CubeWallTimer.Enabled = True
        PlaySound "popper_ball", 0, 1
        Exit Sub
    End If
    If ActiveElementType < 0 Then Exit Sub
    If RuneWordMultiballRunning And GetBIP() > 1 Then
        ActiveElementType = -1
        ElementBallActive = False
        Exit Sub
    End If
    ' Open lid so ball can exit
    CubeLid.Collidable = False
    CubeLidTargetY = CubeLidOpenY
    CubeLidTimer.Enabled = True
    Dim eb : Set eb = CubeKick.CreateBall
    Select Case ActiveElementType
        Case ATTACK_FIRE      : eb.Color = RGB(255, 60, 0)
        Case ATTACK_COLD      : eb.Color = RGB(80, 180, 255)
        Case ATTACK_POISON    : eb.Color = RGB(0, 200, 50)
        Case ATTACK_LIGHTNING : eb.Color = RGB(196, 196, 8)
    End Select
    SetAllBallsElement ActiveElementType
    PlayBonusBallSpawnSound ActiveElementType
  ResumeAmbushTimer
  ResumeKillStreakTimer
    cubekick.Kick 195, 40
  CubeWallTimer.Enabled = False
  CubeWallTimer.Enabled = True
    PlaySound "popper_ball", 0, 1
    CubeLight.State = 0
End Sub

Sub CubeDestroyCurrentAndUpgrade()
    If ActiveElementType = ATTACK_POISON Then
        PoisonTickActive = False
        PoisonTimer.Enabled = False
        Dim bpd : For bpd = 0 To 4 : BumperPoisoned(bpd) = False : Next
    End If
    ActiveElementType = -1
    ElementBallActive = False
    ' Destroy ball in kicker and decrement BIP - spawn timer will add it back
    CubeKick.DestroyBall
    CubeLight.State = 1
    CubeHoldTimer.Enabled = True
End Sub

Sub CubeHoldTimer_Timer()
    CubeHoldTimer.Enabled = False
    If SilverBallID >= 0 Then
        CubeKick.DestroyBall
        SilverBallID = -1
    Else
    ResumeAmbushTimer
    ResumeKillStreakTimer
        cubekick.Kick 195, 40
    CubeWallTimer.Enabled = False
    CubeWallTimer.Enabled = True
        PlaySound "popper_ball", 0, 1
    End If
    CubeSpawnTimer.Enabled = True
End Sub

Sub CubeSpawnTimer_Timer()
    CubeSpawnTimer.Enabled = False
    If PendingElementType < 0 Then Exit Sub
    Select Case PendingElementType
        Case ATTACK_COLD      : SetFlasherColor 1, 80, 180, 255
        Case ATTACK_FIRE      : SetFlasherColor 1, 255, 60, 0
        Case ATTACK_LIGHTNING : SetFlasherColor 1, 196, 196, 8
        Case ATTACK_POISON    : SetFlasherColor 1, 0, 200, 50
    End Select
    FireFlasher 1
    Dim eb : Set eb = CubeKick.CreateBall
    Select Case PendingElementType
        Case ATTACK_FIRE      : eb.Color = RGB(255, 60, 0)
        Case ATTACK_COLD      : eb.Color = RGB(80, 180, 255)
        Case ATTACK_POISON    : eb.Color = RGB(0, 200, 50)
        Case ATTACK_LIGHTNING : eb.Color = RGB(196, 196, 8)
    End Select
    SetAllBallsElement PendingElementType
    ElementBallActive = True

    PlaySound "gem", 0, 1
    PlayBonusBallSpawnSound PendingElementType
    ResumeAmbushTimer
  ResumeKillStreakTimer
    If Not GemWallOpen Then OpenCubeWall
    cubekick.Kick 195, 40
    CubeWallTimer.Enabled = False
  CubeWallTimer.Enabled = True
    PlaySound "popper_ball", 0, 1
    Select Case PendingElementType
        Case ATTACK_COLD      : TransmuteLevel = 1
        Case ATTACK_FIRE      : TransmuteLevel = 2
        Case ATTACK_LIGHTNING : TransmuteLevel = 3
        Case ATTACK_POISON    : TransmuteLevel = 4
    End Select
    If GemQueueCount <= 0 Then
        CloseCubeWall
    End If
    PendingElementType = -1
    Select Case TransmuteLevel
        Case 1
            ShowMessage "COLD BALL!"
            If GIEventMode <> GI_MODE_AMBUSH Then
                Dim cr1 : For cr1 = 1 To 10 : SetGIRow cr1, 80, 180, 255 : Next
                TransmuteFlashTimer.Enabled = True
            End If
        Case 2
            ShowMessage "FIRE BALL!"
            If GIEventMode <> GI_MODE_AMBUSH Then
                Dim cr2 : For cr2 = 1 To 10 : SetGIRow cr2, 255, 60, 0 : Next
                TransmuteFlashTimer.Enabled = True
            End If
        Case 3
            ShowMessage "LIGHTNING BALL!"
            If GIEventMode <> GI_MODE_AMBUSH Then
                Dim cr3 : For cr3 = 1 To 10 : SetGIRow cr3, 196, 196, 8 : Next
                TransmuteFlashTimer.Enabled = True
            End If
        Case 4
            ShowMessage "POISON BALL!"
            If GIEventMode <> GI_MODE_AMBUSH Then
                Dim cr4 : For cr4 = 1 To 10 : SetGIRow cr4, 0, 200, 50 : Next
                TransmuteFlashTimer.Enabled = True
            End If
    End Select
    If GemQueueCount > 0 Then
        GemWallOpen = True
        OpenCubeWall
    End If
    CubeTransmuteActive = False
    UpdateCubeLight
  SetRampLight
End Sub

Sub TransmuteFlashTimer_Timer()
    TransmuteFlashTimer.Enabled = False
    If GIEventMode <> GI_MODE_AMBUSH Then RestoreAllGIRows
End Sub

Sub CubeLightCycleTimer_Timer()
    CubeLightCycleStep = CubeLightCycleStep + 1
    Select Case ((CubeLightCycleStep - 1) Mod 4)
        Case 0 : CubeLight.Color = RGB(80, 180, 255)
        Case 1 : CubeLight.Color = RGB(255, 60, 0)
        Case 2 : CubeLight.Color = RGB(196, 196, 8)
        Case 3 : CubeLight.Color = RGB(0, 200, 50)
    End Select
    CubeLight.BlinkInterval = 50
    CubeLight.BlinkPattern = "10"
    CubeLight.State = 2
    If CubeLightCycleStep >= 12 Then
        CubeLightCycleTimer.Enabled = False
        CubeLight.State = 0
    End If
End Sub

Sub RollRampGem()
    Dim gemType
    Select Case TransmuteLevel
        Case 0 : gemType = ATTACK_COLD
        Case 1 : gemType = ATTACK_FIRE
        Case 2 : gemType = ATTACK_LIGHTNING
        Case 3 : gemType = ATTACK_POISON
        Case 4 : gemType = ATTACK_COLD
        Case Else : Exit Sub
    End Select
    If GemQueueCount < 9 Then
        GemQueue(GemQueueCount) = gemType
        GemQueueCount = GemQueueCount + 1
        OpenCubeWall
    End If
End Sub
'*****************************************
'      PARTY SYSTEM
'*****************************************

Dim PartySpawnPending : PartySpawnPending = False
Dim PortalSpawnInterval : PortalSpawnInterval = 800
Dim PortalFlashStep : PortalFlashStep = 0

Sub PortalFlashTimer_Timer()
    PortalFlashStep = PortalFlashStep + 1
    Select Case PortalFlashStep
        Case 1
            PortalFlashTimer.Interval = 80
            Dim c : c = GetActGIColor()
            Dim r : r = (c \ 65536) And 255
            Dim g : g = (c \ 256) And 255
            Dim b : b = c And 255
            SetFlasherColor 1, r, g, b : SetFlasherColor 2, r, g, b
            FireFlasher 1 : FireFlasher 2
        Case 2 : DimFlasher 1 : DimFlasher 2
        Case 3 : FireFlasher 1 : FireFlasher 2
        Case 4 : DimFlasher 1 : DimFlasher 2
        Case 5 : FireFlasher 1 : FireFlasher 2
        Case 6
            DimFlasher 1 : DimFlasher 2
            PortalFlashTimer.Enabled = False
            PortalFlashStep = 0
    End Select
End Sub

Sub PortalOpenTimer_Timer()
    PortalOpenTimer.Enabled = False
    If Not GameActive Then Exit Sub
    PlayfieldKicker.Enabled = True
    PartySpawnTimer.Interval = PortalSpawnInterval
    PartySpawnTimer.Enabled = True
End Sub

Function GetPartyMagicFind()
    Dim mf : mf = 0
    If PartyBarb  Then mf = mf + 8
    If PartyAma   Then mf = mf + 8
    If PartyNecro Then mf = mf + 8
    If PartySorc Then mf = mf + 8
    If PartyPal   Then mf = mf + 8
    If PartyAss   Then mf = mf + 8
    If PartyDru   Then mf = mf + 8
    GetPartyMagicFind = mf
End Function

Function GetPartyCount()
    Dim cnt : cnt = 0
    If PartyBarb  Then cnt = cnt + 1
    If PartyAma   Then cnt = cnt + 1
    If PartyNecro Then cnt = cnt + 1
    If PartySorc Then cnt = cnt + 1
    If PartyPal   Then cnt = cnt + 1
    If PartyAss   Then cnt = cnt + 1
    If PartyDru   Then cnt = cnt + 1
    GetPartyCount = cnt
End Function

Function CountEquippedUniques()
    Dim cnt : cnt = 0
    Dim i
    For i = 0 To 6
        If GearSlots(i) = GEAR_UNIQUE Then cnt = cnt + 1
    Next
    CountEquippedUniques = cnt
End Function

Function AllUniquesEquipped()
    AllUniquesEquipped = (CountEquippedUniques() = 7)
End Function

Function CritCycleInterval()
    CritCycleInterval = 1000 + (CountEquippedUniques() * 833)
End Function

Function RollGearTier()
    Dim chars : chars = GetPartyCount()
    Dim mf : mf = GetPartyMagicFind()
    Dim magicWeight, rareWeight, uniqueWeight
    Select Case chars
        Case 1 : uniqueWeight = 3
        Case 2 : uniqueWeight = 3
        Case 3 : uniqueWeight = 4
        Case 4 : uniqueWeight = 4
        Case 5 : uniqueWeight = 4
        Case 6 : uniqueWeight = 5
        Case 7 : uniqueWeight = 5
        Case Else : uniqueWeight = 3
    End Select
    ' MagicFind shifts weight toward Rare and Unique (max +4 at full party)
    Dim mfBonus : mfBonus = Int(mf / 14)
    uniqueWeight = uniqueWeight + mfBonus
    ' Difficulty also boosts UNIQUE chance: Nightmare +2, Hell +4
    uniqueWeight = uniqueWeight + (DifficultyLevel * 2)
    magicWeight  = 70 - (chars * 3) - (DifficultyLevel * 8)
    If magicWeight < 5 Then magicWeight = 5
    rareWeight   = 100 - magicWeight - uniqueWeight
    Dim total : total = magicWeight + rareWeight + uniqueWeight
    Dim roll : roll = Int(Rnd * total)
    If roll < magicWeight Then
        RollGearTier = GEAR_MAGIC
    ElseIf roll < magicWeight + rareWeight Then
        RollGearTier = GEAR_RARE
    Else
        RollGearTier = GEAR_UNIQUE
    End If
End Function

Sub UnlockPartyMember(member)
    Select Case member
        Case "Ama"
            If PartyAma Then Exit Sub
            PartyAma = True
            CharLightAma.State = 1
            PlaySound "waypointignite2", 0, 1
            PlaySound "Ama_meetyourdeath", 0, 1
        Case "Necro"
            If PartyNecro Then Exit Sub
            PartyNecro = True
            CharLightNecro.State = 1
            PlaySound "waypointignite2", 0, 1
            PlaySound "Nec_meetyourfate", 0, 1
        Case "Sorc"
            If PartySorc Then Exit Sub
            PartySorc = True
            CharLightSorc.State = 1
            PlaySound "waypointignite2", 0, 1
            PlaySound "Sor_datewithdeath", 0, 1
        Case "Pal"
            If PartyPal Then Exit Sub
            PartyPal = True
            CharLightPal.State = 1
            PlaySound "waypointignite2", 0, 1
            PlaySound "Pal_meetyourdeath", 0, 1
        Case "Ass"
            If PartyAss Then Exit Sub
            PartyAss = True
            CharLightAss.State = 1
            PlaySound "waypointignite2", 0, 1
            PlaySound "Ass_meetyourfate", 0, 1
        Case "Dru"
            If PartyDru Then Exit Sub
            PartyDru = True
            CharLightDru.State = 1
            PlaySound "waypointignite2", 0, 1
            PlaySound "Dru_datewithdeath", 0, 1
    End Select
    CharThanksName = member
    BallSaveActive = True
    BallSaveUsed = False
    BallSaveTimer.Enabled = False
    BallSaveTimer.Interval = 15000 + BallSaveBonus()
    BallSaveTimer.Enabled = True
    BallSaveL.Color = RGB(174, 0, 0) : BallSaveL.ColorFull = RGB(255, 72, 72) : BallSaveL.State = 1
    BallSaveL2.Color = RGB(174, 0, 0) : BallSaveL2.ColorFull = RGB(255, 72, 72) : BallSaveL2.State = 1
    CharThanksTimer.Enabled = True
    UpdateDMD2 member & " JOINS!", "MF: " & GetPartyMagicFind() & "%"
End Sub

Sub CharThanksTimer_Timer()
    CharThanksTimer.Enabled = False
    Select Case Int(Rnd * 2)
        Case 0
            Select Case CharThanksName
                Case "Ama"   : PlaySound "Ama_thanks", 0, 1
                Case "Necro" : PlaySound "Nec_thanks", 0, 1
                Case "Sorc"  : PlaySound "Sor_thanks", 0, 1
                Case "Pal"   : PlaySound "Pal_thanks", 0, 1
                Case "Ass"   : PlaySound "Ass_thanks", 0, 1
                Case "Dru"   : PlaySound "Dru_thanks", 0, 1
            End Select
        Case 1
            Select Case CharThanksName
                Case "Ama"   : PlaySound "Ama_thankyou", 0, 1
                Case "Necro" : PlaySound "Nec_thankyou", 0, 1
                Case "Sorc"  : PlaySound "Sor_thankyou", 0, 1
                Case "Pal"   : PlaySound "Pal_thankyou", 0, 1
                Case "Ass"   : PlaySound "Ass_thankyou", 0, 1
                Case "Dru"   : PlaySound "Dru_thankyou", 0, 1
            End Select
    End Select
    CharThanksName = ""
  CharWelcomeTimer.Interval = 3000
    CharWelcomeTimer.Enabled = True
End Sub

Sub CharWelcomeTimer_Timer()
    CharWelcomeTimer.Enabled = False
    If Not GameActive Then Exit Sub
    ' Ally portal owns townportalPF + PlayfieldKicker — clear any merc portal/ping
    If MercPortalArmed Then DisarmMercPortal
    MercPingTimer.Enabled = False
    MercPinged = False
    MercPingPending = False
    If PartyMultiballRunning Then
        ' Mystery or another ally spawn is running — add to queue
        PartySpawnCount = PartySpawnCount + 1
        If Not PartySpawnTimer.Enabled And Not PortalOpenTimer.Enabled Then
            PortalSpawnInterval = 800
            PartySpawnTimer.Interval = 800
            PartySpawnTimer.Enabled = True
        End If
    Else
        PartySpawnCount = 1
        PartySpawnStep = 0
        PartyMultiballRunning = True
        PortalSpawnInterval = 800
        PortalOpenTimer.Interval = 3000
        PortalOpenTimer.Enabled = True
        PortalFlashStep = 0
        PortalFlashTimer.Interval = 2000
        PortalFlashTimer.Enabled = True
    End If
    BallSaveActive = True
    BallSaveUsed = False
    BallSaveTimer.Enabled = False
    BallSaveTimer.Interval = 8000 + BallSaveBonus()
    BallSaveTimer.Enabled = True
    BallSaveL.Color = RGB(174, 0, 0) : BallSaveL.ColorFull = RGB(255, 72, 72) : BallSaveL.State = 1
    BallSaveL2.Color = RGB(174, 0, 0) : BallSaveL2.ColorFull = RGB(255, 72, 72) : BallSaveL2.State = 1
    townportalPF.Visible = True
    PlaySound "portalenter", 0, 1
    ShowMessage "WELCOME TO THE PARTY!"
    UpdateDMD2 "NEW ALLY!", "ALLY MULTIBALL!"
End Sub

Sub InitParty()
  CurrentAct = 1
  ShieldHeld = False
  FirstUniqueEquipped = False
  CubeRespawnHoldTimer.Enabled = False
  ActiveElementType = -1
    ElementBallActive = False
  FireBurnSlot = -1
    FireBurnTimer.Enabled = False
  SilverBallID = -1
  ForceFrontLoot     = False
    FirstCubeHit = True
    PartyBarb  = True
    PartyAma   = False
    PartyNecro = False
    PartySorc = False
    PartyPal   = False
    PartyAss   = False
    PartyDru   = False
  TransmuteLevel = 0
    EnemiesKilled      = 0
    KillMilestoneCount = 0
    FirstKillLootGiven = False
    FirstMercGiven = False
    MercHires = 0
    ForceNextLootTier  = -1
    CharLightBarb.State  = 1
    CharLightAma.State   = 0
    CharLightNecro.State = 0
    CharLightSorc.State  = 0
    CharLightPal.State   = 0
    CharLightAss.State   = 0
    CharLightDru.State   = 0
  RampGemCount = 0
  CubeSpinCount = 0
  TravelProgress = 0
    EventIndex = 0
    BossRegenTimer.Enabled = False
    StopCritCycle
    BossFightActive = False
    BossEventType = -1
    TravelActive = True
    SetTravelLights True
End Sub

'*****************************************
'      PARTY UNLOCK TRIGGERS
'*****************************************

Sub CapTarget_Hit()
    PlaySoundAtLevelActiveBall "Metal_Touch_" & Int(Rnd * 13) + 1, Vol(ActiveBall) * MetalImpactSoundFactor
    If BossFightActive Then HitBoss
'    AddScore 500000
End Sub

'*****************************************
'      LOOT SYSTEM
'*****************************************

Function GearSlotAvailable(tier)
    Dim i
    For i = 0 To 6
        If GearSlots(i) = -1 Then GearSlotAvailable = True : Exit Function
        If GearSlots(i) < tier Then GearSlotAvailable = True : Exit Function
    Next
    GearSlotAvailable = False
End Function

Function RollGearProperty()
    RollGearProperty = Int(Rnd * 5)
End Function

Function GearPropertyName(prop, tier)
    Dim prefix
    Select Case tier
        Case GEAR_MAGIC  : prefix = "MAGIC "
        Case GEAR_RARE   : prefix = "RARE "
        Case GEAR_UNIQUE : prefix = "UNIQUE "
    End Select
End Function

Function GearPropertyValue(prop, tier)
    Select Case tier
        Case GEAR_MAGIC  : GearPropertyValue = 1
        Case GEAR_RARE   : GearPropertyValue = 2
        Case GEAR_UNIQUE : GearPropertyValue = 3
    End Select
End Function

Sub CollectLoot(slot)
    If Not LootActive(slot) Then Exit Sub
    LootActive(slot) = False
    PrimDtLootArr(slot).Visible = False
    DtLootCurZ(slot)      = DtLootDroppedZ
    DtLootTargetZ(slot)   = DtLootDroppedZ
    PrimDtLootArr(slot).z = DtLootDroppedZ
    UniqueLevPhase(slot)  = 0

    Select Case LootPending(slot)
        Case LOOT_GEAR
            Select Case LootTier(slot)
                Case GEAR_MAGIC
'                    AddScore 25000
                    AssignGearToSlot GEAR_MAGIC
                    If Not GearSlotAvailable(GEAR_MAGIC) Then
                        BallGoldCount = BallGoldCount + Int(50000 * GoldActScale())
                        PlaySound "gold", 0, 1
                        ShowMessage "SOLD! +50,000 GOLD"
                    End If
                Case GEAR_RARE
'                    AddScore 50000
                    AssignGearToSlot GEAR_RARE
                    If Not GearSlotAvailable(GEAR_RARE) Then
                        BallGoldCount = BallGoldCount + Int(100000 * GoldActScale())
                        PlaySound "gold", 0, 1
                        ShowMessage "SOLD! +100,000 GOLD"
                    End If
                Case GEAR_UNIQUE
'                    AddScore 100000
                    AssignGearToSlot GEAR_UNIQUE
                    If Not GearSlotAvailable(GEAR_UNIQUE) Then
                        BallGoldCount = BallGoldCount + Int(200000 * GoldActScale())
                        PlaySound "gold", 0, 1
                        ShowMessage "SOLD! +200,000 GOLD"
                    End If
            End Select
    End Select
End Sub

Sub GearFlashTimer_Timer()
    GearFlashTimer.Enabled = False
    WeaponLight1.State = 0
End Sub

Sub ClearGearBonuses()
    ClearLootTargets
End Sub

Sub ResetLootTargetColor(slot)
    SetLootPrimAppearance slot, -1     ' -1 hits the Else case: neutral color, glow stays on
End Sub

Sub ClearLootTargets()
    Dim cl
    For cl = 0 To 3
        LootActive(cl) = False
        LootPending(cl) = -1
        LootValue(cl) = 0
        LootTier(cl) = -1
        SetLootPrimAppearance cl, -1      ' neutral color, reset glow, clears tier (-1)
        UniqueLevPhase(cl) = 0            ' stop any levitation
        UniqueLevZ(cl)     = 0
        UniqueSpinDeg(cl)  = 0
        PrimDtLootArr(cl).RotY = DtLootRestRotY(cl)   ' undo any spin
        DtLootArr(cl).IsDropped = True
        SetDtLootAnim cl, True            ' sink the prim with drop anim
    Next
End Sub

Sub dtLoot1_Hit()
    SoundDropTargetDrop dtLoot1
    CollectLoot 0
    dtLoot1.IsDropped = True : SetDtLootAnim 0, True
    ResetLootTargetColor 0
End Sub

Sub dtLoot2_Hit()
    SoundDropTargetDrop dtLoot2
    CollectLoot 1
    dtLoot2.IsDropped = True : SetDtLootAnim 1, True
    ResetLootTargetColor 1
End Sub

Sub dtLoot3_Hit()
    SoundDropTargetDrop dtLoot3
    CollectLoot 2
    dtLoot3.IsDropped = True : SetDtLootAnim 2, True
    ResetLootTargetColor 2
End Sub

Sub dtLoot4_Hit()
    SoundDropTargetDrop dtLoot4
    CollectLoot 3
    dtLoot4.IsDropped = True : SetDtLootAnim 3, True
    ResetLootTargetColor 3
End Sub

'Sub RandomSoundDropTargetReset(obj)
'    PlaySoundAtLevelStatic SoundFX("Drop_Target_Reset_" & Int(Rnd * 6) + 1, DOFContactors), 1, obj
'End Sub

'Sub SoundDropTargetDrop(obj)
'    PlaySoundAtLevelStatic "Drop_Target_Down_" & Int(Rnd * 6) + 1, DTSoundLevel, obj
'End Sub


Sub QuestLight1_animate : p15.BlendDisableLighting  = 200 * (QuestLight1.GetInPlayIntensity  / QuestLight1.Intensity)  : End Sub
Sub QuestLight2_animate : p005.BlendDisableLighting = 200 * (QuestLight2.GetInPlayIntensity  / QuestLight2.Intensity)  : End Sub
Sub QuestLight3_animate : p006.BlendDisableLighting = 200 * (QuestLight3.GetInPlayIntensity  / QuestLight3.Intensity)  : End Sub
Sub QuestLight4_animate : p007.BlendDisableLighting = 200 * (QuestLight4.GetInPlayIntensity  / QuestLight4.Intensity)  : End Sub
Sub CapBallLight_animate : p23.BlendDisableLighting = 200 * (CapBallLight.GetInPlayIntensity / CapBallLight.Intensity) : End Sub

Sub SetRampLight()
    Exit Sub   ' element shown only by CubeLight now; QuestLight3 decoupled
    If TutorialActive Then Exit Sub
    Select Case TransmuteLevel
        Case 0  ' Cold queued — ice blue
            QuestLight3.Color     = RGB(80, 180, 255)
            QuestLight3.ColorFull = RGB(80, 180, 255)
            QuestLight3.State     = 1
        Case 1  ' Fire queued — orange/red
            QuestLight3.Color     = RGB(255, 60, 0)
            QuestLight3.ColorFull = RGB(255, 60, 0)
            QuestLight3.State     = 1
        Case 2  ' Lightning queued — yellow
            QuestLight3.Color     = RGB(196, 196, 8)
            QuestLight3.ColorFull = RGB(196, 196, 8)
            QuestLight3.State     = 1
        Case 3  ' Poison queued — green
            QuestLight3.Color     = RGB(0, 200, 50)
            QuestLight3.ColorFull = RGB(50, 255, 80)
            QuestLight3.State     = 1
        Case 4  ' All elements done — jackpot purple, slow blink
            QuestLight3.Color     = RGB(180, 0, 255)
            QuestLight3.ColorFull = RGB(180, 0, 255)
            QuestLight3.BlinkInterval = 400
            QuestLight3.State     = 2
    End Select
End Sub

Sub QuestLightBlinkTimer_Timer()
    QuestLightBlinkTimer.Enabled = False
    If MysteryReady Then
        QuestLight4.Color = RGB(180, 0, 255)
        QuestLight4.ColorFull = RGB(180, 0, 255)
        QuestLight4.BlinkInterval = 300
        QuestLight4.State = 2
    End If
    If TravelActive And Not BossFightActive Then
        SetTravelLights True
        If MysteryReady Then
            QuestLight4.Color = RGB(180, 0, 255)
            QuestLight4.ColorFull = RGB(180, 0, 255)
            QuestLight4.BlinkInterval = 300
            QuestLight4.State = 2
        End If
    Else
        QuestLight1.State = 0
        QuestLight2.State = 0
    SetRampLight
        If Not MysteryReady Then QuestLight4.State = 0
    End If
End Sub

Function ShouldQuestTriggerBlink()
    If BossFightActive Then
        ShouldQuestTriggerBlink = False
        Exit Function
    End If
    If TravelActive Then
        ShouldQuestTriggerBlink = True
        Exit Function
    End If
    ShouldQuestTriggerBlink = False
End Function

Sub QuestTrigger1_Hit()
    If ActiveBall.VelY > 0 Then Exit Sub
    OnCritLaneHit 0
    TravelHit
    CheckAmbush 1
    CheckCainIntro
    TriggerShout 1
    If ShouldQuestTriggerBlink() Then
        QuestLight1.BlinkInterval = 250
        QuestLight1.BlinkPattern = "10"
        QuestLight1.State = 2
        QuestLightBlinkTimer.Enabled = True
    End If
End Sub

Sub QuestTrigger2_Hit()
    If ActiveBall.VelY > 0 Then Exit Sub
    OnCritLaneHit 1
    TravelHit
    CheckAmbush 2
    CheckCainIntro
    TriggerShout 2
    If ShouldQuestTriggerBlink() Then
        QuestLight2.BlinkInterval = 250
        QuestLight2.BlinkPattern = "10"
        QuestLight2.State = 2
        QuestLightBlinkTimer.Enabled = True
    End If
End Sub

Sub QuestTrigger3_Hit()
    ChargeRampAOE
    TriggerShout 3
End Sub

Sub QuestTrigger4_Hit()
    If ActiveBall.VelY > 0 Then Exit Sub
    OnCritLaneHit 2
    TravelHit
    CheckAmbush 4
    CheckCainIntro
    TriggerShout 4
    If ShouldQuestTriggerBlink() And Not MysteryReady Then
        QuestLight4.BlinkInterval = 250
        QuestLight4.BlinkPattern = "10"
        QuestLight4.State = 2
        QuestLightBlinkTimer.Enabled = True
    End If
End Sub

'RUNEWORD SYSTEM

Sub RuneLightR_animate  : pRuneR.BlendDisableLighting  = 200 * (RuneLightR.GetInPlayIntensity  / RuneLightR.Intensity)  : End Sub
Sub RuneLightU_animate  : pRuneU.BlendDisableLighting  = 200 * (RuneLightU.GetInPlayIntensity  / RuneLightU.Intensity)  : End Sub
Sub RuneLightN_animate  : pRuneN.BlendDisableLighting  = 200 * (RuneLightN.GetInPlayIntensity  / RuneLightN.Intensity)  : End Sub
Sub RuneLightE_animate  : pRuneE.BlendDisableLighting  = 200 * (RuneLightE.GetInPlayIntensity  / RuneLightE.Intensity)  : End Sub
Sub RuneLightW_animate  : pRuneW.BlendDisableLighting  = 200 * (RuneLightW.GetInPlayIntensity  / RuneLightW.Intensity)  : End Sub
Sub RuneLightO_animate  : pRuneO.BlendDisableLighting  = 200 * (RuneLightO.GetInPlayIntensity  / RuneLightO.Intensity)  : End Sub
Sub RuneLightR2_animate : pRuneR2.BlendDisableLighting = 200 * (RuneLightR2.GetInPlayIntensity / RuneLightR2.Intensity) : End Sub
Sub RuneLightD_animate  : pRuneD.BlendDisableLighting  = 200 * (RuneLightD.GetInPlayIntensity  / RuneLightD.Intensity)  : End Sub
Sub RuneKickLight_animate : p004.BlendDisableLighting = 200 * (RuneKickLight.GetInPlayIntensity / RuneKickLight.Intensity) : End Sub

Dim RuneWordSweepStep : RuneWordSweepStep = 0
Dim RuneWordSweepFlasher : RuneWordSweepFlasher = 0
Dim RuneKickIsRuneword : RuneKickIsRuneword = False
Dim RuneQuality : RuneQuality = 0
Dim RuneWordJackpotLevel : RuneWordJackpotLevel = 0
Dim RuneWordActive : RuneWordActive = False
Dim RuneWordReady : RuneWordReady = False
Dim RuneWordMultiballRunning : RuneWordMultiballRunning = False
Dim RuneWordSockets : RuneWordSockets = 0
Dim RuneWordName : RuneWordName = ""
Dim RuneWordJackpotTimer_step : RuneWordJackpotTimer_step = 0

Dim RuneHitR : RuneHitR = False
Dim RuneHitU : RuneHitU = False
Dim RuneHitN : RuneHitN = False
Dim RuneHitE : RuneHitE = False
Dim RuneHitW : RuneHitW = False
Dim RuneHitO : RuneHitO = False
Dim RuneHitR2 : RuneHitR2 = False
Dim RuneHitD : RuneHitD = False

Dim RunePopStep(7)
Dim RunePopActive(7)
Dim RunePopRestZ(7)
Const RunePopHeight = 20    ' how many units it rises
Const RunePopSteps = 12
Dim RuneWaveStep : RuneWaveStep = 0

Function GetRunePrimitive(idx)
    Select Case idx
        Case 0 : Set GetRunePrimitive = Rune1
        Case 1 : Set GetRunePrimitive = Rune2
        Case 2 : Set GetRunePrimitive = Rune3
        Case 3 : Set GetRunePrimitive = Rune4
        Case 4 : Set GetRunePrimitive = Rune5
        Case 5 : Set GetRunePrimitive = Rune6
        Case 6 : Set GetRunePrimitive = Rune7
        Case 7 : Set GetRunePrimitive = Rune8
    End Select
End Function

Sub TriggerRunePop(idx)
    RunePopStep(idx) = 0
    RunePopActive(idx) = True
    RunePopTimer.Enabled = True
End Sub

Sub RunePopTimer_Timer()
    Dim i, anyActive
    anyActive = False
    For i = 0 To 7
        If RunePopActive(i) Then
            anyActive = True
            RunePopStep(i) = RunePopStep(i) + 1
            Dim prim : Set prim = GetRunePrimitive(i)
            If RunePopStep(i) <= RunePopSteps / 2 Then
                prim.TransZ = RunePopRestZ(i) + RunePopHeight * (RunePopStep(i) / (RunePopSteps / 2))
            Else
                prim.TransZ = RunePopRestZ(i) + RunePopHeight * (1 - (RunePopStep(i) - RunePopSteps / 2) / (RunePopSteps / 2))
            End If
            If RunePopStep(i) >= RunePopSteps Then
                prim.TransZ = RunePopRestZ(i)
                RunePopActive(i) = False
                RunePopStep(i) = 0
            End If
        End If
    Next
    If Not anyActive Then RunePopTimer.Enabled = False
End Sub

Sub GetRandomRuneWord(ByRef outName, ByRef targetSockets)
    Dim names(77), sockets(77)
    names(0)  = "Steel"              : sockets(0)  = 2
    names(1)  = "Nadir"              : sockets(1)  = 2
    names(2)  = "Malice"             : sockets(2)  = 3
    names(3)  = "Stealth"            : sockets(3)  = 2
    names(4)  = "Leaf"               : sockets(4)  = 2
    names(5)  = "Zephyr"             : sockets(5)  = 2
    names(6)  = "Ancient's Pledge"   : sockets(6)  = 3
    names(7)  = "Holy Thunder"       : sockets(7)  = 4
    names(8)  = "Strength"           : sockets(8)  = 2
    names(9)  = "Edge"               : sockets(9)  = 3
    names(10) = "King's Grace"       : sockets(10) = 3
    names(11) = "Spirit"             : sockets(11) = 4
    names(12) = "Myth"               : sockets(12) = 3
    names(13) = "Radiance"           : sockets(13) = 3
    names(14) = "Insight"            : sockets(14) = 4
    names(15) = "Lore"               : sockets(15) = 2
    names(16) = "Honor"              : sockets(16) = 5
    names(17) = "Rhyme"              : sockets(17) = 2
    names(18) = "Peace"              : sockets(18) = 3
    names(19) = "Black"              : sockets(19) = 3
    names(20) = "White"              : sockets(20) = 2
    names(21) = "Smoke"              : sockets(21) = 2
    names(22) = "Splendor"           : sockets(22) = 2
    names(23) = "Memory"             : sockets(23) = 4
    names(24) = "Harmony"            : sockets(24) = 4
    names(25) = "Melody"             : sockets(25) = 3
    names(26) = "Lionheart"          : sockets(26) = 3
    names(27) = "Obedience"          : sockets(27) = 5
    names(28) = "Treachery"          : sockets(28) = 3
    names(29) = "Passion"            : sockets(29) = 4
    names(30) = "Voice of Reason"    : sockets(30) = 4
    names(31) = "Wealth"             : sockets(31) = 3
    names(32) = "Lawbringer"         : sockets(32) = 3
    names(33) = "Enlightenment"      : sockets(33) = 3
    names(34) = "Crescent Moon"      : sockets(34) = 3
    names(35) = "Duress"             : sockets(35) = 3
    names(36) = "Stone"              : sockets(36) = 4
    names(37) = "Gloom"              : sockets(37) = 3
    names(38) = "Bone"               : sockets(38) = 3
    names(39) = "Prudence"           : sockets(39) = 2
    names(40) = "Rain"               : sockets(40) = 3
    names(41) = "Venom"              : sockets(41) = 3
    names(42) = "Sanctuary"          : sockets(42) = 3
    names(43) = "Oath"               : sockets(43) = 4
    names(44) = "Delirium"           : sockets(44) = 3
    names(45) = "Rift"               : sockets(45) = 4
    names(46) = "Kingslayer"         : sockets(46) = 4
    names(47) = "Principle"          : sockets(47) = 3
    names(48) = "Heart of the Oak"   : sockets(48) = 4
    names(49) = "Silence"            : sockets(49) = 6
    names(50) = "Death"              : sockets(50) = 5
    names(51) = "Chaos"              : sockets(51) = 3
    names(52) = "Call to Arms"       : sockets(52) = 5
    names(53) = "Exile"              : sockets(53) = 4
    names(54) = "Fortitude"          : sockets(54) = 4
    names(55) = "Grief"              : sockets(55) = 5
    names(56) = "Wind"               : sockets(56) = 2
    names(57) = "Bramble"            : sockets(57) = 4
    names(58) = "Dragon"             : sockets(58) = 3
    names(59) = "Wrath"              : sockets(59) = 4
    names(60) = "Beast"              : sockets(60) = 5
    names(61) = "Chains of Honor"    : sockets(61) = 4
    names(62) = "Infinity"           : sockets(62) = 4
    names(63) = "Eternity"           : sockets(63) = 5
    names(64) = "Ice"                : sockets(64) = 4
    names(65) = "Dream"              : sockets(65) = 3
    names(66) = "Fury"               : sockets(66) = 3
    names(67) = "Famine"             : sockets(67) = 4
    names(68) = "Faith"              : sockets(68) = 4
    names(69) = "Brand"              : sockets(69) = 4
    names(70) = "Phoenix"            : sockets(70) = 4
    names(71) = "Enigma"             : sockets(71) = 3
    names(72) = "Destruction"        : sockets(72) = 5
    names(73) = "Last Wish"          : sockets(73) = 6
    names(74) = "Doom"               : sockets(74) = 5
    names(75) = "Hand of Justice"    : sockets(75) = 4
    names(76) = "Pride"              : sockets(76) = 4
    names(77) = "Breath of the Dying": sockets(77) = 6

    ' Build filtered list matching targetSockets
    Dim matches(77), matchCount : matchCount = 0
    Dim i
    For i = 0 To 77
        If sockets(i) = targetSockets Then
            matches(matchCount) = i
            matchCount = matchCount + 1
        End If
    Next
    ' Pick random from matching names
    If matchCount > 0 Then
        Dim pick : pick = matches(Int(Rnd * matchCount))
        outName = names(pick)
    Else
        outName = "El Eld Tir"  ' fallback, shouldn't happen
    End If
End Sub

Function GetSocketsFromQuality()
    Dim q : q = RuneQuality + 2  ' +2: ball in play counts as 1, scoop shot counts as 1
    If q < 2 Then q = 2
    If q > 6 Then q = 6
    GetSocketsFromQuality = q
End Function

Sub ResetRuneTargets()
    RuneWaveTimer.Enabled = False
    RuneWaveStep = 0
    RuneHitR  = False : RuneLightR.State  = 0
    RuneHitU  = False : RuneLightU.State  = 0
    RuneHitN  = False : RuneLightN.State  = 0
    RuneHitE  = False : RuneLightE.State  = 0
    RuneHitW  = False : RuneLightW.State  = 0
    RuneHitO  = False : RuneLightO.State  = 0
    RuneHitR2 = False : RuneLightR2.State = 0
    RuneHitD  = False : RuneLightD.State  = 0
    RuneWordReady = False
    RuneKickLight.State = 0
    If Not RuneKickBlinkTimer.Enabled Then
        RuneKickBlinkStep = 0
    End If
    DimFlasher 2
End Sub

Sub RuneFlasherDimTimer_Timer()
    RuneFlasherDimTimer.Enabled = False
    DimFlasher 2
End Sub

Sub CheckRuneWordComplete()
    If RuneHitR And RuneHitU And RuneHitN And RuneHitE And _
       RuneHitW And RuneHitO And RuneHitR2 And RuneHitD Then
        RuneWordReady = True
        RuneLightR.BlinkInterval  = 200 : RuneLightR.State  = 2
        RuneLightU.BlinkInterval  = 200 : RuneLightU.State  = 2
        RuneLightN.BlinkInterval  = 200 : RuneLightN.State  = 2
        RuneLightE.BlinkInterval  = 200 : RuneLightE.State  = 2
        RuneLightW.BlinkInterval  = 200 : RuneLightW.State  = 2
        RuneLightO.BlinkInterval  = 200 : RuneLightO.State  = 2
        RuneLightR2.BlinkInterval = 200 : RuneLightR2.State = 2
        RuneLightD.BlinkInterval  = 200 : RuneLightD.State  = 2
        ShowMessage "RUNEWORD! SHOOT SCOOP!"
        PlaySound "waypointignite2", 0, 1
        SetFlasherColor 2, 255, 180, 0
        FireFlasher 2
        RuneFlasherDimTimer.Enabled = False
        RuneFlasherDimTimer.Enabled = True
    RuneKickLight.BlinkInterval = 200
    RuneKickLight.BlinkPattern = "10"
    RuneKickLight.State = 2
    End If
End Sub

Sub RuneR_Hit()
    PlaySoundAtLevelActiveBall "Metal_Touch_" & Int(Rnd * 13) + 1, Vol(ActiveBall) * MetalImpactSoundFactor
'    If RuneWordMultiballRunning Then Exit Sub
    If RuneWordReady Then Exit Sub
    If RuneHitR Then Exit Sub
    RuneHitR = True
    RuneLightR.State = 1
    AddScore 100000
'    If CurrentAct = 6 And PortalActive Then
'        PortalTime = PortalTime + 500
'        If PortalTime > PortalCap() Then PortalTime = PortalCap()
'    End If
    PlaySound "rune", 0, 1
    SetFlasherColor 2, 255, 140, 0
    FireFlasher 2
    RuneFlasherDimTimer.Enabled = False
    RuneFlasherDimTimer.Interval = 400
    RuneFlasherDimTimer.Enabled = True
    TriggerRunePop 0
    CheckRuneWordComplete
End Sub

Sub RuneU_Hit()
    PlaySoundAtLevelActiveBall "Metal_Touch_" & Int(Rnd * 13) + 1, Vol(ActiveBall) * MetalImpactSoundFactor
'    If RuneWordMultiballRunning Then Exit Sub
    If RuneWordReady Then Exit Sub
    If RuneHitU Then Exit Sub
    RuneHitU = True
    RuneLightU.State = 1
    AddScore 100000
'    If CurrentAct = 6 And PortalActive Then
'        PortalTime = PortalTime + 500
'        If PortalTime > PortalCap() Then PortalTime = PortalCap()
'    End If
    PlaySound "rune", 0, 1
    SetFlasherColor 2, 255, 140, 0
    FireFlasher 2
    RuneFlasherDimTimer.Enabled = False
    RuneFlasherDimTimer.Interval = 400
    RuneFlasherDimTimer.Enabled = True
    TriggerRunePop 1
    CheckRuneWordComplete
End Sub

Sub RuneN_Hit()
    PlaySoundAtLevelActiveBall "Metal_Touch_" & Int(Rnd * 13) + 1, Vol(ActiveBall) * MetalImpactSoundFactor
'    If RuneWordMultiballRunning Then Exit Sub
    If RuneWordReady Then Exit Sub
    If RuneHitN Then Exit Sub
    RuneHitN = True
    RuneLightN.State = 1
    AddScore 100000
'    If CurrentAct = 6 And PortalActive Then
'        PortalTime = PortalTime + 500
'        If PortalTime > PortalCap() Then PortalTime = PortalCap()
'    End If
    PlaySound "rune", 0, 1
    SetFlasherColor 2, 255, 140, 0
    FireFlasher 2
    RuneFlasherDimTimer.Enabled = False
    RuneFlasherDimTimer.Interval = 400
    RuneFlasherDimTimer.Enabled = True
    TriggerRunePop 2
    CheckRuneWordComplete
End Sub

Sub RuneE_Hit()
    PlaySoundAtLevelActiveBall "Metal_Touch_" & Int(Rnd * 13) + 1, Vol(ActiveBall) * MetalImpactSoundFactor
'    If RuneWordMultiballRunning Then Exit Sub
    If RuneWordReady Then Exit Sub
    If RuneHitE Then Exit Sub
    RuneHitE = True
    RuneLightE.State = 1
    AddScore 100000
'    If CurrentAct = 6 And PortalActive Then
'        PortalTime = PortalTime + 500
'        If PortalTime > PortalCap() Then PortalTime = PortalCap()
'    End If
    PlaySound "rune", 0, 1
    SetFlasherColor 2, 255, 140, 0
    FireFlasher 2
    RuneFlasherDimTimer.Enabled = False
    RuneFlasherDimTimer.Interval = 400
    RuneFlasherDimTimer.Enabled = True
    TriggerRunePop 3
    CheckRuneWordComplete
End Sub

Sub RuneW_Hit()
    PlaySoundAtLevelActiveBall "Metal_Touch_" & Int(Rnd * 13) + 1, Vol(ActiveBall) * MetalImpactSoundFactor
'    If RuneWordMultiballRunning Then Exit Sub
    If RuneWordReady Then Exit Sub
    If RuneHitW Then Exit Sub
    RuneHitW = True
    RuneLightW.State = 1
    AddScore 100000
'    If CurrentAct = 6 And PortalActive Then
'        PortalTime = PortalTime + 500
'        If PortalTime > PortalCap() Then PortalTime = PortalCap()
'    End If
    PlaySound "rune", 0, 1
    SetFlasherColor 2, 255, 140, 0
    FireFlasher 2
    RuneFlasherDimTimer.Enabled = False
    RuneFlasherDimTimer.Interval = 400
    RuneFlasherDimTimer.Enabled = True
    TriggerRunePop 4
    CheckRuneWordComplete
End Sub

Sub RuneO_Hit()
    PlaySoundAtLevelActiveBall "Metal_Touch_" & Int(Rnd * 13) + 1, Vol(ActiveBall) * MetalImpactSoundFactor
'    If RuneWordMultiballRunning Then Exit Sub
    If RuneWordReady Then Exit Sub
    If RuneHitO Then Exit Sub
    RuneHitO = True
    RuneLightO.State = 1
    AddScore 100000
'    If CurrentAct = 6 And PortalActive Then
'        PortalTime = PortalTime + 500
'        If PortalTime > PortalCap() Then PortalTime = PortalCap()
'    End If
    PlaySound "rune", 0, 1
    SetFlasherColor 2, 255, 140, 0
    FireFlasher 2
    RuneFlasherDimTimer.Enabled = False
    RuneFlasherDimTimer.Interval = 400
    RuneFlasherDimTimer.Enabled = True
    TriggerRunePop 5
    CheckRuneWordComplete
End Sub

Sub RuneR2_Hit()
    PlaySoundAtLevelActiveBall "Metal_Touch_" & Int(Rnd * 13) + 1, Vol(ActiveBall) * MetalImpactSoundFactor
'    If RuneWordMultiballRunning Then Exit Sub
    If RuneWordReady Then Exit Sub
    If RuneHitR2 Then Exit Sub
    RuneHitR2 = True
    RuneLightR2.State = 1
    AddScore 100000
'    If CurrentAct = 6 And PortalActive Then
'        PortalTime = PortalTime + 500
'        If PortalTime > PortalCap() Then PortalTime = PortalCap()
'    End If
    PlaySound "rune", 0, 1
    SetFlasherColor 2, 255, 140, 0
    FireFlasher 2
    RuneFlasherDimTimer.Enabled = False
    RuneFlasherDimTimer.Interval = 400
    RuneFlasherDimTimer.Enabled = True
    TriggerRunePop 6
    CheckRuneWordComplete
End Sub

Sub RuneD_Hit()
    PlaySoundAtLevelActiveBall "Metal_Touch_" & Int(Rnd * 13) + 1, Vol(ActiveBall) * MetalImpactSoundFactor
'    If RuneWordMultiballRunning Then Exit Sub
    If RuneWordReady Then Exit Sub
    If RuneHitD Then Exit Sub
    RuneHitD = True
    RuneLightD.State = 1
    AddScore 100000
'    If CurrentAct = 6 And PortalActive Then
'        PortalTime = PortalTime + 500
'        If PortalTime > PortalCap() Then PortalTime = PortalCap()
'    End If
    PlaySound "rune", 0, 1
    SetFlasherColor 2, 255, 140, 0
    FireFlasher 2
    RuneFlasherDimTimer.Enabled = False
    RuneFlasherDimTimer.Interval = 400
    RuneFlasherDimTimer.Enabled = True
    TriggerRunePop 7
    CheckRuneWordComplete
End Sub

Sub RuneKick_Hit()
    If RuneKickSeqTimer.Enabled Or RuneKickFlashTimer.Enabled Then Exit Sub  ' ← ADD THIS
    If Not GameActive Then
        RuneKickIsRuneword = False
        RuneKickSeqTimer.Enabled = False
        RuneKickSeqTimer.Enabled = True
        Exit Sub
    End If
  PauseAmbushTimer
  PauseKillStreakTimer
  ' Max runes guard — show message instead of re-reporting +4
    If Not RuneWordReady And RuneQuality >= 6 Then
        PlaySound "waypointignite2", 0, 0.6
        Select Case Int(Rnd * 3)
            Case 0 : PlaySound "grunt203", 0, 1
            Case 1 : PlaySound "grunt202", 0, 1
            Case 2 : PlaySound "grunt201", 0, 1
        End Select
        ShowMessage "MAX RUNES"
        RuneKickSeqTimer.Enabled = False
        RuneKickSeqTimer.Enabled = True
        Exit Sub
    End If
    Dim rkbt : rkbt = GetActiveBonusBallType()
    If rkbt >= 0 And Not RuneWordReady And RuneWordMultiballRunning Then
        RuneKickIsRuneword = False
        RuneQuality = RuneQuality + 1
        If RuneQuality > 6 Then RuneQuality = 6
        PlaySound "waypointignite2", 0, 0.6
        Select Case Int(Rnd * 3)
            Case 0 : PlaySound "grunt203", 0, 1
            Case 1 : PlaySound "grunt202", 0, 1
            Case 2 : PlaySound "grunt201", 0, 1
        End Select
        PlaySound "malus", 0, 1
        SetFlasherColor 2, 255, 180, 0
        FireFlasher 2
        RuneFlasherDimTimer.Enabled = False
        RuneFlasherDimTimer.Enabled = True
        Dim dispQ1 : dispQ1 = RuneQuality : If dispQ1 > 4 Then dispQ1 = 4
        ShowMessage "RUNES:+" & dispQ1
        RuneKickSeqTimer.Enabled = False
        RuneKickSeqTimer.Enabled = True
        Exit Sub
    End If
    If Not RuneWordReady Then
        RuneKickIsRuneword = False
        RuneQuality = RuneQuality + 1
        If RuneQuality > 6 Then RuneQuality = 6
        PlaySound "waypointignite2", 0, 0.6
        Select Case Int(Rnd * 3)
            Case 0 : PlaySound "grunt203", 0, 1
            Case 1 : PlaySound "grunt202", 0, 1
            Case 2 : PlaySound "grunt201", 0, 1
        End Select
        PlaySound "malus", 0, 1
        SetFlasherColor 2, 255, 180, 0
        FireFlasher 2
        RuneFlasherDimTimer.Enabled = False
        RuneFlasherDimTimer.Enabled = True
        Dim dispQ2 : dispQ2 = RuneQuality : If dispQ2 > 4 Then dispQ2 = 4
        ShowMessage "RUNES:+" & dispQ2
        RuneKickSeqTimer.Enabled = False
        RuneKickSeqTimer.Enabled = True
        Exit Sub
    End If
    RuneKickIsRuneword = True
    RuneWordReady = False
    RuneKickLight.State = 0
    GetRandomRuneWord RuneWordName, RuneWordSockets
    RuneWordSockets = GetSocketsFromQuality()
    GetRandomRuneWord RuneWordName, RuneWordSockets
    ShowMessage "RUNEWORD!"
    PlaySound "waypointignite2", 0, 1
    RuneKickSeqTimer.Enabled = False
    RuneKickSeqTimer.Enabled = True
End Sub

Sub RuneKickSeqTimer_Timer()
    RuneKickSeqTimer.Enabled = False
    SetFlasherColor 3, 255, 180, 0
    SetFlasherColor 4, 255, 180, 0
    FireFlasher 3
    FireFlasher 4
    RuneKickFlashTimer.Enabled = False
    RuneKickFlashTimer.Enabled = True
End Sub

'Function GetPlayBIP()
'    Dim balls : balls = GetBalls()
'    Dim cnt : cnt = 0
'    Dim b
'    For Each b In balls
'        If b.ID <> CapBallID And b.ID <> CapBall2ID Then cnt = cnt + 1
'    Next
'    GetPlayBIP = cnt
'End Function

Sub RuneKickFlashTimer_Timer()
    RuneKickFlashTimer.Enabled = False
    DimFlasher 3
    DimFlasher 4
    If RuneKickIsRuneword Then
        RuneKickIsRuneword = False
        RuneWordJackpotTimer_step = 0
        RuneWordJackpotTimer.Enabled = True
    Else
    ResumeAmbushTimer
    ResumeKillStreakTimer
        RuneKick.Kick 180, 30
    End If
End Sub

Sub RuneWordJackpotTimer_Timer()
    RuneWordJackpotTimer_step = RuneWordJackpotTimer_step + 1
    Select Case RuneWordJackpotTimer_step
        Case 1
      DimAllFlashers
            UpdateDMD "RUNEWORD!"
            PlaySound "grunt203", 0, 1
            SetFlasherColor 1, 255, 180, 0 : SetFlasherColor 2, 255, 180, 0
            SetFlasherColor 3, 255, 180, 0 : SetFlasherColor 4, 255, 180, 0
            SetFlasherColor 5, 255, 180, 0 : SetFlasherColor 6, 255, 180, 0
            FireAllFlashers
            RuneWordMalusTimer.Interval = 400
            RuneWordMalusTimer.Enabled = True
        Case 2
      DimAllFlashers
            UpdateDMD RuneWordName
            PlaySound "grunt202", 0, 1
            SetFlasherColor 1, 255, 180, 0 : SetFlasherColor 2, 255, 180, 0
            SetFlasherColor 3, 255, 180, 0 : SetFlasherColor 4, 255, 180, 0
            SetFlasherColor 5, 255, 180, 0 : SetFlasherColor 6, 255, 180, 0
            FireAllFlashers
            RuneWordMalusTimer.Interval = 400
            RuneWordMalusTimer.Enabled = True
        Case 3
      DimAllFlashers
            UpdateDMD RuneWordSockets & " SOCKETS"
            PlaySound "grunt201", 0, 1
            SetFlasherColor 1, 255, 180, 0 : SetFlasherColor 2, 255, 180, 0
            SetFlasherColor 3, 255, 180, 0 : SetFlasherColor 4, 255, 180, 0
            SetFlasherColor 5, 255, 180, 0 : SetFlasherColor 6, 255, 180, 0
            FireAllFlashers
            RuneWordMalusTimer.Interval = 400
            RuneWordMalusTimer.Enabled = True
        Case 4
      RuneWaveStep = 0
            RuneWaveTimer.Enabled = True
            RuneQuality = 0
            RuneWordJackpotTimer.Enabled = False
            ShowMessage "MULTIBALL!"
            RuneWordMultiballRunning = True
'            If CurrentAct = 6 And PortalActive Then
'                PortalTime = PortalTime + 10000
'                If PortalTime > PortalCap() Then PortalTime = PortalCap()
'            Else
'                AddPortalSurge 10
'            End If
            BallSaveActive = False
            BallSaveUsed = False
            RuneWordSpawnCount = 0
            RuneKickBlinkStep = 0
            RuneKickBlinkTimer.Enabled = True
            StartGIEvent GI_MODE_RUNEWORD
            PlayCallout "RunewordMultiDiablo", 3000
            RuneWordSweepStep = 0
            RuneWordSweepFlasher = 5
            RuneWordSweepTimer.Interval = 300
            RuneWordSweepTimer.Enabled = True
    End Select
End Sub

Sub RuneWaveTimer_Timer()
    RuneWaveStep = RuneWaveStep + 1
    ' 3 full waves across 8 runes = 24 + settling steps
    ' Each rune pops when wave reaches it
    Dim wavePos : wavePos = (RuneWaveStep - 1) Mod 8
    Dim waveNum : waveNum = (RuneWaveStep - 1) \ 8

    If RuneWaveStep > 24 Then
        RuneWaveTimer.Enabled = False
        Exit Sub
    End If

    ' Pop the rune at current wave position
    TriggerRunePop wavePos

    ' Also trigger a light blink on that rune
    Select Case wavePos
        Case 0 : RuneLightR.BlinkInterval  = 80 : RuneLightR.State  = 2
        Case 1 : RuneLightU.BlinkInterval  = 80 : RuneLightU.State  = 2
        Case 2 : RuneLightN.BlinkInterval  = 80 : RuneLightN.State  = 2
        Case 3 : RuneLightE.BlinkInterval  = 80 : RuneLightE.State  = 2
        Case 4 : RuneLightW.BlinkInterval  = 80 : RuneLightW.State  = 2
        Case 5 : RuneLightO.BlinkInterval  = 80 : RuneLightO.State  = 2
        Case 6 : RuneLightR2.BlinkInterval = 80 : RuneLightR2.State = 2
        Case 7 : RuneLightD.BlinkInterval  = 80 : RuneLightD.State  = 2
    End Select
End Sub

Sub RuneWordMalusTimer_Timer()
    RuneWordMalusTimer.Enabled = False
    PlaySound "malus", 0, 1
End Sub

Sub RuneWordSweepTimer_Timer()
  RuneWordSweepTimer.Interval = 80
    RuneWordSweepStep = RuneWordSweepStep + 1

    Dim sweepOrder(5)
    sweepOrder(0) = 5 : sweepOrder(1) = 6
    sweepOrder(2) = 1 : sweepOrder(3) = 2
    sweepOrder(4) = 3 : sweepOrder(5) = 4

    If RuneWordSweepStep > 18 Then
        RuneWordSweepTimer.Enabled = False
        DimAllFlashers
        Exit Sub
    End If

    ' Dim all first, then light only current
    DimAllFlashers

    Dim fi : fi = (RuneWordSweepStep - 1) Mod 6
    RuneWordSweepFlasher = sweepOrder(fi)
    SetFlasherColor RuneWordSweepFlasher, 255, 180, 0
    FireFlasher RuneWordSweepFlasher
End Sub

Dim RuneKickBlinkStep : RuneKickBlinkStep = 0

Sub RuneKickBlinkTimer_Timer()
    RuneKickBlinkStep = RuneKickBlinkStep + 1
    If RuneKickBlinkStep Mod 2 = 1 Then
        RuneKickLight.State = 1
    Else
        RuneKickLight.State = 0
    End If
    If RuneKickBlinkStep >= 6 Then
        RuneKickBlinkTimer.Enabled = False
        RuneKickLight.State = 0
        BallSaveActive = True
        BallSaveUsed = False
        BallSaveTimer.Enabled = False
        BallSaveTimer.Interval = 11000
        BallSaveTimer.Enabled = True
        BallSaveL.State = 1
        BallSaveL2.State = 1
        SetAllBallsElement ActiveElementType
    ResumeAmbushTimer
    ResumeKillStreakTimer
        RuneKick.Kick 180, 30
        PlaySound "popper_ball", 0, 1
        RuneQuality = 0  ' reset to 0 — next cycle starts fresh
        ResetRuneTargets
        If GetBIP() > 0 Then
            RuneWordSpawnTimer.Enabled = True
        Else
            RuneWordMultiballRunning = False
        End If
    End If
End Sub

Dim RuneWordSpawnCount : RuneWordSpawnCount = 0

Sub RuneWordSpawnTimer_Timer()
    If RuneWordSpawnCount >= RuneWordSockets - 1 Then
        RuneWordSpawnTimer.Enabled = False
        Exit Sub
    End If
    RuneWordSpawnCount = RuneWordSpawnCount + 1
    Dim nb : Set nb = RuneKick.CreateBall
    SetAllBallsElement ActiveElementType
  ResumeAmbushTimer
  ResumeKillStreakTimer
    RuneKick.Kick 180, 30
    PlaySound "popper_ball", 0, 1
    PlaySound "malus", 0, 1
    If GIEventMode <> GI_MODE_AMBUSH Then
        Dim rwf : For rwf = 1 To 10 : SetGIRow rwf, 255, 180, 0 : Next
        TransmuteFlashTimer.Enabled = False
        TransmuteFlashTimer.Enabled = True
    End If
End Sub

'*****************************************
'      SAVE LANES / BARB SHIELD
'*****************************************

Dim SaveHitS : SaveHitS = False
Dim SaveHitA : SaveHitA = False
Dim SaveHitV : SaveHitV = False
Dim SaveHitE : SaveHitE = False
Dim ShieldArmor : ShieldArmor = 0
Dim ShieldActive : ShieldActive = False
Dim ShieldReadyBlinkStep : ShieldReadyBlinkStep = 0
Dim ShieldHeld : ShieldHeld = False

Sub S_Light_animate : p12.BlendDisableLighting = 200 * (S_Light.GetInPlayIntensity / S_Light.Intensity) : End Sub
Sub A_Light_animate : p001.BlendDisableLighting = 200 * (A_Light.GetInPlayIntensity / A_Light.Intensity) : End Sub
Sub V_Light_animate : p002.BlendDisableLighting = 200 * (V_Light.GetInPlayIntensity / V_Light.Intensity) : End Sub
Sub E_Light_animate : p003.BlendDisableLighting = 200 * (E_Light.GetInPlayIntensity / E_Light.Intensity) : End Sub

Sub InitSaveLanes()
    If Not ShieldActive Then
        Shield.Collidable = False
        Shield.Visible = False
    End If
    S_Light.State = 0
    A_Light.State = 0
    V_Light.State = 0
    E_Light.State = 0
    If SaveHitS Then S_Light.State = 1
    If SaveHitA Then A_Light.State = 1
    If SaveHitV Then V_Light.State = 1
    If SaveHitE Then E_Light.State = 1
    If ShieldActive Then
        UpdateShieldLight
    Else
        ShieldLight.State = 0
    End If
End Sub

Sub ResetSaveSystem()
    SaveHitS = False
    SaveHitA = False
    SaveHitV = False
    SaveHitE = False
    S_Light.State = 0
    A_Light.State = 0
    V_Light.State = 0
    E_Light.State = 0
    ShieldReadyBlinkTimer.Enabled = False
    ShieldReadyBlinkStep = 0
    BallShieldCount = 0
    If ShieldHeld Then
        ShieldHeld = False
        ' Shield carries over — keep armor, active state and light
        ShowMessage "SHIELD CARRIED OVER!"
        UpdateShieldLight
    Else
        ShieldArmor = 0
        ShieldActive = False
        Shield.Collidable = False
        Shield.Visible = False
        ShieldLight.State = 0
        ShieldMaxArmor = 0
    End If
End Sub

Sub CheckSaveComplete()
    If SaveHitS And SaveHitA And SaveHitV And SaveHitE Then
        SaveHitS = False
        SaveHitA = False
        SaveHitV = False
        SaveHitE = False
        S_Light.State = 0
        A_Light.State = 0
        V_Light.State = 0
        E_Light.State = 0
        Dim armorGain : armorGain = 1
        If PartyPal Then armorGain = 2
      ShieldArmor = ShieldArmor + armorGain
        BallShieldCount = BallShieldCount + 1
'        If CurrentAct = 6 And PortalActive Then
'            PortalTime = PortalTime + 5000
'            If PortalTime > PortalCap() Then PortalTime = PortalCap()
'        End If
        If ShieldActive Then
            ShowMessage "SHIELD +" & armorGain & " ARMOR! (" & ShieldArmor & " LEFT)"
            PlaySound "malus", 0, 1
            UpdateShieldLight
        Else
            If Not CalloutPlaying Then PlayCallout "Bar_act2_complete_tightspot", 2000
            ActivateShield
        End If
    End If
End Sub

Sub ShieldReadyBlinkTimer_Timer()
    ShieldReadyBlinkStep = ShieldReadyBlinkStep + 1
    If ShieldReadyBlinkStep Mod 2 = 0 Then
        ShieldLight.State = 0
    Else
        ShieldLight.State = 1
    End If
    If ShieldReadyBlinkStep >= 6 Then
        ShieldReadyBlinkTimer.Enabled = False
        ShieldLight.State = 1
    End If
End Sub

Sub S_Trigger_Hit()
    SaveHitS = Not SaveHitS
    UpdateSaveLights
    AddScore 50000
    CheckSaveComplete
End Sub

Sub A_Trigger_Hit()
    SaveHitA = Not SaveHitA
    UpdateSaveLights
    AddScore 50000
    CheckSaveComplete
End Sub

Sub V_Trigger_Hit()
    SaveHitV = Not SaveHitV
    UpdateSaveLights
    AddScore 50000
    CheckSaveComplete
End Sub

Sub E_Trigger_Hit()
    SaveHitE = Not SaveHitE
    UpdateSaveLights
    AddScore 50000
    CheckSaveComplete
End Sub

Sub ActivateShield()
    If ShieldArmor <= 0 Then Exit Sub
    ShieldActive = True
    ShieldReadyBlinkTimer.Enabled = False
    Shield.Collidable = True
    Shield.Visible = True
    ShieldMaxArmor = ShieldArmor
    UpdateShieldLight
    ShowMessage "SHIELD UP! " & ShieldArmor & " ARMOR"
    PlaySound "circle1", 0, 1
    StartGIEvent GI_MODE_SHIELD
End Sub

Dim ShieldMaxArmor : ShieldMaxArmor = 0

Sub UpdateShieldLight()
    If Not ShieldActive Then Exit Sub
    If ShieldMaxArmor = 0 Then Exit Sub
    Dim pct : pct = ShieldArmor / ShieldMaxArmor
    If pct > 0.66 Then
        ShieldLight.Color = RGB(80, 180, 255)
        ShieldLight.State = 1
    ElseIf pct > 0.33 Then
        ShieldLight.Color = RGB(255, 200, 0)
        ShieldLight.BlinkInterval = 500
        ShieldLight.State = 2
    ElseIf pct > 0 Then
        ShieldLight.Color = RGB(255, 60, 0)
        ShieldLight.BlinkInterval = 200
        ShieldLight.State = 2
    Else
        ShieldLight.State = 0
    End If
End Sub

Sub Shield_Hit()
    If Not ShieldActive Then Exit Sub
    Select Case Int(Rnd * 9)
        Case 0 : PlaySound "block blunt4", 0, 0.8
        Case 1 : PlaySound "block blunt2", 0, 0.8
        Case 2 : PlaySound "block blunt1", 0, 0.8
        Case 3 : PlaySound "block blade5", 0, 0.8
        Case 4 : PlaySound "block blade4", 0, 0.8
        Case 5 : PlaySound "block blade1", 0, 0.8
        Case 6 : PlaySound "block arrow3", 0, 0.8
        Case 7 : PlaySound "block arrow2", 0, 0.8
        Case 8 : PlaySound "block arrow1", 0, 0.8
    End Select
    ShieldArmor = ShieldArmor - 1
    UpdateShieldLight
    SetFlasherColor 5, 80, 180, 255 : SetFlasherColor 6, 80, 180, 255
    FireFlasher 5 : FireFlasher 6
  ApronFlasherDimTimer.Enabled = False
    ApronFlasherDimTimer.Enabled = True
    If ShieldArmor <= 0 Then
        ShieldActive = False
        Shield.Collidable = False
        Shield.Visible = False
        ShieldLight.State = 0
        ShowMessage "SHIELD BROKEN!"
        PlaySound "malus", 0, 1
        SetFlasherColor 5, 255, 0, 0 : SetFlasherColor 6, 255, 0, 0
        FireFlasher 5 : FireFlasher 6
    Else
        ShowMessage "SHIELD HIT! " & ShieldArmor & "ARMOR"
    End If
End Sub

Sub CycleSaveLanesLeft()
    Dim tmpE : tmpE = SaveHitE
    SaveHitE = SaveHitV
    SaveHitV = SaveHitA
    SaveHitA = SaveHitS
    SaveHitS = tmpE
    UpdateSaveLights
End Sub

Sub CycleSaveLanesRight()
    Dim tmpS : tmpS = SaveHitS
    SaveHitS = SaveHitA
    SaveHitA = SaveHitV
    SaveHitV = SaveHitE
    SaveHitE = tmpS
    UpdateSaveLights
End Sub

Sub UpdateSaveLights()
    If SaveHitS Then S_Light.State = 1 Else S_Light.State = 0
    If SaveHitA Then A_Light.State = 1 Else A_Light.State = 0
    If SaveHitV Then V_Light.State = 1 Else V_Light.State = 0
    If SaveHitE Then E_Light.State = 1 Else E_Light.State = 0
End Sub

'*****************************************
'      BARBARIAN SHOUT SYSTEM
'*****************************************
Dim ShoutMultiplierHeld : ShoutMultiplierHeld = False
Dim ShoutHoldQueueCount : ShoutHoldQueueCount = 0
Dim BallLootBonus : BallLootBonus = 0
Dim ShoutMultiplier : ShoutMultiplier = 1
Dim ShoutHitCount : ShoutHitCount = 0
Dim ShoutPulseStep : ShoutPulseStep = 0

Sub TriggerShout(triggerNum)
    If BonusActive Then Exit Sub
    Dim snd
    Select Case triggerNum
        Case 1
            Select Case Int(Rnd * 3) + 1
                Case 1 : snd = "warcry1"
                Case 2 : snd = "warcry2"
                Case 3 : snd = "warcry3"
            End Select
            QuestLight1.BlinkInterval = 100
            QuestLight1.State = 2
        Case 2
            Select Case Int(Rnd * 3) + 1
                Case 1 : snd = "order1"
                Case 2 : snd = "order2"
                Case 3 : snd = "order3"
            End Select
            QuestLight2.BlinkInterval = 100
            QuestLight2.State = 2
        Case 3
            Select Case Int(Rnd * 3) + 1
                Case 1 : snd = "howl1"
                Case 2 : snd = "howl2"
                Case 3 : snd = "howl3"
            End Select
'            QuestLight3.BlinkInterval = 100
'            QuestLight3.State = 2
        Case 4
            Select Case Int(Rnd * 3) + 1
                Case 1 : snd = "battlecry1"
                Case 2 : snd = "battlecry2"
                Case 3 : snd = "battlecry3"
            End Select
            QuestLight4.BlinkInterval = 100
            QuestLight4.State = 2
    End Select
    PlaySound snd, 0, 1 * DuckVolume
    ShoutHitCount = ShoutHitCount + 1
    If triggerNum <> 3 And Not AmbushActive Then ShowMessage "SHOUT! " & ShoutHitCount & " OF 4"
    SetFlasherColor 3, 255, 150, 0 : SetFlasherColor 4, 255, 150, 0
    FireFlasher 3 : FireFlasher 4
    ShoutBlinkOffTimer.Interval = 300
    ShoutBlinkOffTimer.Enabled = True
    If ShoutHitCount >= 4 Then
        ShoutHitCount = 0
        SetFlasherColor 3, 255, 200, 50 : SetFlasherColor 4, 255, 200, 50
        FireFlasher 3 : FireFlasher 4
        If ShoutMultiplier < 10 Then
            ShoutMultiplier = ShoutMultiplier + 1
            If triggerNum <> 3 And Not AmbushActive Then ShowMessage "WAR CRY! x" & ShoutMultiplier & " BONUS!"
            GIEventSweepCount = 0
            GIEventSweepMax = 8
            GIEventSweepDir = 1
            StartGIEvent GI_MODE_SHOUT
        Else
            AddScore 10000000
            PlayCallout "MaxShout", 4000
            If triggerNum <> 3 And Not AmbushActive Then ShowMessage "ANCIENTS CALL! +10,000,000"
            GIEventSweepCount = 0
            GIEventSweepMax = 8
            GIEventSweepDir = 1
            StartGIEvent GI_MODE_SHOUT
        End If
        ShoutPulseStep = 0
        ShoutPulseTimer.Enabled = True
    Else
        ' Single sweep for every non-completing shout
        GIEventSweepCount = 0
        GIEventSweepMax = 1
        GIEventSweepDir = 1
        StartGIEvent GI_MODE_SHOUT
    End If
End Sub

Dim ShoutBlinkOffLight
Sub ShoutBlinkOffTimer_Timer()
    ShoutBlinkOffTimer.Enabled = False
    QuestLight1.State = 0
    QuestLight2.State = 0
'    QuestLight3.State = 0
    QuestLight4.State = 0
    If CritCycleActive Then
        SetCritCycleLight CritCycleStep   ' ← restore crit lane display, don't leave it dark
    ElseIf TravelActive And Not BossFightActive Then
        SetTravelLights True
    End If
    If Not CritCycleActive And MysteryReady Then
        QuestLight4.Color = RGB(180, 0, 255)
        QuestLight4.ColorFull = RGB(180, 0, 255)
        QuestLight4.BlinkInterval = 300
        QuestLight4.State = 2
    End If
End Sub
Sub ShoutPulseTimer_Timer()
    ShoutPulseStep = ShoutPulseStep + 1
    If ShoutPulseStep Mod 2 = 0 Then
        QuestLight1.State = 0
        QuestLight2.State = 0
'        QuestLight3.State = 0
        QuestLight4.State = 0
    Else
        QuestLight1.State = 1
        QuestLight2.State = 1
'        QuestLight3.State = 1
        QuestLight4.State = 1
    End If
    If ShoutPulseStep >= 6 Then
        ShoutPulseTimer.Enabled = False
        QuestLight1.State = 0
        QuestLight2.State = 0
'        QuestLight3.State = 0
        QuestLight4.State = 0
        If CritCycleActive Then
            SetCritCycleLight CritCycleStep   ' ← restore crit lane display after War Cry pulse
        ElseIf TravelActive And Not BossFightActive Then
            SetTravelLights True
        End If
        If Not CritCycleActive And MysteryReady Then
            QuestLight4.Color = RGB(180, 0, 255)
            QuestLight4.ColorFull = RGB(180, 0, 255)
            QuestLight4.BlinkInterval = 300
            QuestLight4.State = 2
        End If
    End If
End Sub

Sub InitShoutLights()
    ShoutHitCount = 0
    ShoutPulseTimer.Enabled = False
    ShoutPulseStep = 0
    QuestLight1.State = 0
    QuestLight2.State = 0
'    QuestLight3.State = 0
    QuestLight4.State = 0
    If CritCycleActive Then
        SetCritCycleLight CritCycleStep   ' ← carry crit lane display into new ball if boss fight continues
    ElseIf TravelActive And Not BossFightActive Then
        SetTravelLights True
    End If
    If Not CritCycleActive And MysteryReady Then
        QuestLight4.Color = RGB(180, 0, 255)
        QuestLight4.ColorFull = RGB(180, 0, 255)
        QuestLight4.BlinkInterval = 300
        QuestLight4.State = 2
    End If
End Sub

'*****************************************
'   END OF BALL BONUS
'*****************************************

Dim BallKillCount : BallKillCount = 0
Dim BallShieldCount : BallShieldCount = 0
Dim BonusStep : BonusStep = 0
Dim BonusActive : BonusActive = False
Dim BonusKills : BonusKills = 0
Dim BonusWhirlwind : BonusWhirlwind = 0
Dim BonusBackstab : BonusBackstab = 0
Dim BonusBallsRem : BonusBallsRem = 0
Dim BonusTotal : BonusTotal = 0
Const BALLS_REMAINING_BONUS = 25000000

Sub StartEndOfBallBonus()
    If BonusActive Then Exit Sub
    BonusWhirlwind = BallWhirlwindScore
    BonusBackstab  = BallBackstabScore
    BonusBallsRem = 0
    If WasActSixComplete Then
        Dim ballsRem : ballsRem = 5 - BallNumber
        If ballsRem < 0 Then ballsRem = 0
        BonusBallsRem = ballsRem * BALLS_REMAINING_BONUS
    End If
    'BonusSafeTravel = SafeTravelBank
    BonusSafeTravel = 0
    BonusCritStrike = BallCritCount * 1000000
    BonusTotal = (BonusWhirlwind + BonusBackstab + BonusCritStrike) * ShoutMultiplier + BonusBallsRem
    If BonusTotal = 0 And Not WasActSixComplete Then
        FinishBonus
        Exit Sub
    End If
    BonusActive = True
    BonusStep = 0
    BonusTimer.Interval = 2000
    BonusTimer.Enabled = True
End Sub

Sub BonusTimer_Timer()
    BonusStep = BonusStep + 1
    If WasActSixComplete Then
        BonusActSixStep
    Else
        BonusNormalStep
    End If
End Sub

Sub BonusNormalStep()
    Select Case BonusStep
        Case 1
            UpdateDMD2 "END OF BALL", "BONUS"
        Case 2
            If BonusWhirlwind = 0 Then BonusTimer_Timer : Exit Sub
            UpdateDMD2 "WHIRLWIND KILLS:", "+" & FormatNumber(BonusWhirlwind, 0, -1, 0, -1)
        Case 3
            BonusTimer_Timer
        Case 4
            If BonusBackstab = 0 Then BonusTimer_Timer : Exit Sub
            UpdateDMD2 "BACKSTAB KILLS:", "+" & FormatNumber(BonusBackstab, 0, -1, 0, -1)
        Case 5
            If BonusCritStrike = 0 Then BonusTimer_Timer : Exit Sub
            UpdateDMD2 "CRIT STRIKES:", BallCritCount & " x 1,000,000"
        Case 6
            If ShoutMultiplier <= 1 Then BonusTimer_Timer : Exit Sub
            UpdateDMD2 "SHOUT BONUS:", "x" & ShoutMultiplier & " MULTIPLIER!"
        Case 7
            UpdateDMD2 "TOTAL BONUS:", "+" & FormatNumber(BonusTotal, 0, -1, 0, -1)
        Case 8
            FinishBonus
    End Select
End Sub

Sub BonusActSixStep()
    Select Case BonusStep
        Case 1
            UpdateDMD2 "SANCTUARY SAVED!", "YOU ARE THE HERO"
        Case 2
            If BonusWhirlwind = 0 Then BonusTimer_Timer : Exit Sub
            UpdateDMD2 "WHIRLWIND KILLS:", "+" & FormatNumber(BonusWhirlwind, 0, -1, 0, -1)
        Case 3
            BonusTimer_Timer
        Case 4
            If BonusBackstab = 0 Then BonusTimer_Timer : Exit Sub
            UpdateDMD2 "BACKSTAB KILLS:", "+" & FormatNumber(BonusBackstab, 0, -1, 0, -1)
        Case 5
            If BonusSafeTravel = 0 Then BonusTimer_Timer : Exit Sub
            UpdateDMD2 "SAFE TRAVEL:", "+" & FormatNumber(BonusSafeTravel, 0, -1, 0, -1)
        Case 6
            If ShoutMultiplier <= 1 Then BonusTimer_Timer : Exit Sub
            UpdateDMD2 "SHOUT BONUS:", "x" & ShoutMultiplier & " MULTIPLIER!"
        Case 7
            Dim ballsRem : ballsRem = 5 - BallNumber
            If ballsRem < 0 Then ballsRem = 0
            If ballsRem > 0 Then
                UpdateDMD2 "BALLS SAVED: " & ballsRem, "+" & FormatNumber(BonusBallsRem, 0, -1, 0, -1)
            Else
                UpdateDMD2 "FINAL BALL!", "WELL PLAYED!"
            End If
        Case 8
            UpdateDMD2 "TOTAL BONUS:", "+" & FormatNumber(BonusTotal, 0, -1, 0, -1)
        Case 9
            UpdateDMD2 "DIABLO IS DEFEATED!", "CONGRATULATIONS!"
        Case 10
            BonusTimer.Enabled = False
            BonusActive = False
            AddScore BonusTotal
            BallKillCount = 0
            BallGoldCount = 0
            BallShieldCount = 0
            BallWhirlwindScore = 0
            BallBackstabScore = 0
            BallLootBonus = 0
            SafeTravelBank = 0
            WasActSixComplete = False
            GameOver
    End Select
End Sub

Sub FinishBonus()
    BonusTimer.Enabled = False
    AddScore BonusTotal
    BonusActive = False
  BallLootBonus = 0
    SafeTravelBank = 0
    BallKillCount = 0
    BallGoldCount = 0
    BallShieldCount = 0
    BallWhirlwindScore = 0
    BallBackstabScore = 0
    BallLootBonus = 0
    If WasActSixComplete Then
        WasActSixComplete = False
        GameOver
    ElseIf ExtraBallConsumed And ExtraBallPending = 0 And BallNumber >= MaxBalls Then
        GameOver
    Else
        LaunchNextBall
    End If
End Sub

Sub AdvanceBonusStep()
  If BonusStep >= 6 Then Exit Sub
    If Not BonusActive Then Exit Sub
    ' Don't allow skipping War Cry or Total Bonus display
    If BonusStep >= 5 Then Exit Sub
    BonusTimer.Enabled = False
    BonusStep = BonusStep + 1
    BonusTimer_Timer
    BonusTimer.Interval = 2000
    BonusTimer.Enabled = True
End Sub

'*****************************************
'      AMBUSH SYSTEM
'*****************************************

Dim AmbushPrize        : AmbushPrize = 0
Dim AmbushPrizeBonus   : AmbushPrizeBonus = 0
Dim AmbushClearHold    : AmbushClearHold = False
Const AmbushPrizeStart = 5000000
Const AmbushPrizeMin   = 500000
Const AmbushPrizeDecay = 100000
Dim AmbushDMDPaused : AmbushDMDPaused = False

Sub AmbushResumeTimer_Timer()
    AmbushResumeTimer.Enabled = False
    If Not AmbushActive Then Exit Sub
    AmbushDMDPaused = False
    AmbushHurryTimer.Enabled = True
    UpdateAmbushDMD
End Sub


Sub AmbushHurryTimer_Timer()
    If Not AmbushActive Then
        AmbushHurryTimer.Enabled = False
        Exit Sub
    End If
    AmbushPrize = AmbushPrize - AmbushPrizeDecay
    If AmbushPrize < AmbushPrizeMin Then AmbushPrize = AmbushPrizeMin
    UpdateAmbushDMD
End Sub

Sub CheckAmbush(triggerNum)
    If Not TravelActive Then Exit Sub
    If BossFightActive Then Exit Sub
    If AmbushActive Then Exit Sub
    If CurrentAct = 6 Then Exit Sub
    If Rnd > 0.25 Then Exit Sub
    AmbushActive = True
    AmbushKillCount = 0
    AmbushTriggerNum = triggerNum
    AmbushPrize = AmbushPrizeStart + AmbushPrizeBonus
    AmbushPrizeBonus = AmbushPrizeBonus + 1000000
    TravelActive = False
    StartGIEvent GI_MODE_AMBUSH
    SetFlasherColor 3, 255, 0, 0
    SetFlasherColor 4, 255, 0, 0
    FireFlasher 3
    FireFlasher 4
    Select Case triggerNum
    Case 1 : QKick1.Enabled = True
    Case 2 : QKick2.Enabled = True
    Case 4 : QKick4.Enabled = True
  End Select
    QuestLight1.Color = RGB(255, 0, 0) : QuestLight1.BlinkInterval = 300 : QuestLight1.State = 2
    QuestLight2.Color = RGB(255, 0, 0) : QuestLight2.BlinkInterval = 300 : QuestLight2.State = 2
'    QuestLight3.Color = RGB(255, 0, 0) : QuestLight3.BlinkInterval = 300 : QuestLight3.State = 2
    QuestLight4.Color = RGB(255, 0, 0) : QuestLight4.BlinkInterval = 300 : QuestLight4.State = 2
Select Case Int(Rnd * 4)
        Case 0 : PlayCallout "WarItIs", 3000
        Case 1 : PlayCallout "MoreBlood", 2000
        Case 2 : PlayCallout "OneOfMany4", 3000
        Case 3 : PlayCallout "BlockPathNotLong2", 3000
    End Select
    PlayCallout "ShootBumpersDiablo", 2000
    AmbushLightTimer.Enabled = True
    AmbushHoldTimer.Enabled = True
    AmbushHurryTimer.Interval = 1000
    AmbushHurryTimer.Enabled = True
    UpdateAmbushDMD
    Dim asi
    For asi = 0 To 4
        If BumperActive(asi) Then AssignBumperStats asi, False
    Next
End Sub

Sub AmbushHoldTimer_Timer()
    AmbushHoldTimer.Enabled = False
    ' Release ball from kicker
    Select Case AmbushTriggerNum
    Case 1 : QKick1.Kick 70, 15 : QKick1.Enabled = False
    Case 2 : QKick2.Kick 0, 15 : QKick2.Enabled = False
    Case 4 : QKick4.Kick 0, 15 : QKick4.Enabled = False
End Select
    PlaySound "popper_ball", 0, 1
End Sub

Sub AmbushLightTimer_Timer()
    ' Keep quest lights blinking red during ambush
    If Not AmbushActive Then
        AmbushLightTimer.Enabled = False
        Exit Sub
    End If
    QuestLight1.Color = RGB(255, 0, 0) : QuestLight1.BlinkInterval = 300 : QuestLight1.State = 2
    QuestLight2.Color = RGB(255, 0, 0) : QuestLight2.BlinkInterval = 300 : QuestLight2.State = 2
'    QuestLight3.Color = RGB(255, 0, 0) : QuestLight3.BlinkInterval = 300 : QuestLight3.State = 2
    QuestLight4.Color = RGB(255, 0, 0) : QuestLight4.BlinkInterval = 300 : QuestLight4.State = 2
End Sub

Sub CheckAmbushKill()
    If Not AmbushActive Then Exit Sub
    AmbushKillCount = AmbushKillCount + 1
    UpdateAmbushDMD
    If AmbushKillCount >= AmbushKillsRequired Then
        CompleteAmbush
    End If
End Sub

Sub CompleteAmbush()
    AmbushResumeTimer.Enabled = False
    AmbushHurryTimer.Enabled = False
    AmbushActive = False
  SetRampLight
    AmbushLightTimer.Enabled = False
    GIEventMode = GI_MODE_NONE
    TravelActive = True
  SetTravelLights True
    AddScore AmbushPrize
    UpdateDMD2 "AMBUSH CLEARED!", "+" & FormatNumber(AmbushPrize, 0, -1, 0, -1)
    ShowMessage "AMBUSH CLEARED! +" & FormatNumber(AmbushPrize, 0, -1, 0, -1)
    AmbushClearHold = True
    AmbushClearTimer.Interval = 2500
    AmbushClearTimer.Enabled = True
  GIEventTimer.Enabled = False
    RestoreAllGIRows
    StartGIEvent GI_MODE_BOSSWIN
    SetTravelLights True
    SetFlasherColor 3, 255, 220, 100
    SetFlasherColor 4, 255, 220, 100
    FireFlasher 3
    FireFlasher 4
End Sub

Sub ResetAmbush()
    AmbushResumeTimer.Enabled = False
    AmbushHurryTimer.Enabled = False
    AmbushLightTimer.Enabled = False
    AmbushHoldTimer.Enabled = False
    Select Case AmbushTriggerNum
    Case 1 : QKick1.Kick 180, 15 : QKick1.Enabled = False
    Case 2 : QKick2.Kick 180, 15 : QKick2.Enabled = False
    Case 4 : QKick4.Kick 180, 15 : QKick4.Enabled = False
End Select
    AmbushActive = False
  SetRampLight
    AmbushKillCount = 0
    AmbushTriggerNum = 0
    GIEventMode = GI_MODE_NONE
    GIEventTimer.Enabled = False
    RestoreAllGIRows
    ' Apply permanent enemy HP penalty
    AmbushFailedPenalty = AmbushFailedPenalty + 2
    UpdateDMD2 "AMBUSH FAILED!", "ENEMIES GROW STRONGER"
    If TravelActive And Not BossFightActive Then SetTravelLights True
    AmbushFailTimer.Interval = 3000
    AmbushFailTimer.Enabled = True
End Sub

Sub AmbushFailTimer_Timer()
    AmbushFailTimer.Enabled = False
    UpdateDMDScore
End Sub

Sub AmbushClearTimer_Timer()
    AmbushClearTimer.Enabled = False
    AmbushClearHold = False
    UpdateDMDScore
End Sub

Sub QKick1_Hit()
    If AmbushActive And AmbushTriggerNum = 1 Then
        AmbushHoldTimer.Enabled = False
        AmbushHoldTimer.Enabled = True
    End If
End Sub

Sub QKick2_Hit()
    If AmbushActive And AmbushTriggerNum = 2 Then
        AmbushHoldTimer.Enabled = False
        AmbushHoldTimer.Enabled = True
    End If
End Sub

Sub QKick4_Hit()
    If AmbushActive And AmbushTriggerNum = 4 Then
        AmbushHoldTimer.Enabled = False
        AmbushHoldTimer.Enabled = True
    End If
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
  LS.Object = LeftSlingShot
  LS.EndPoint1 = EndPoint1LS
  LS.EndPoint2 = EndPoint2LS

  RS.Object = RightSlingShot
  RS.EndPoint1 = EndPoint1RS
  RS.EndPoint2 = EndPoint2RS

  'Slingshot angle corrections (pt, BallPos in %, Angle in deg)
  ' These values are best guesses. Retune them if needed based on specific table research.
  AddSlingsPt 0, 0.00, - 3
  AddSlingsPt 1, 0.30, - 5
  AddSlingsPt 2, 0.40,  -30
  AddSlingsPt 3, 0.60,  30
  AddSlingsPt 4, 0.70,  5
  AddSlingsPt 5, 1.00,  3
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
'Function dSin(degrees)
' dsin = sin(degrees * Pi/180)
'End Function
'Function dCos(degrees)
' dcos = cos(degrees * Pi/180)
'End Function

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
'   ZFLB:  FLUPPER BUMPERS
'******************************************************
' Based on FlupperBumpers 0.145 final

' Explanation of how these bumpers work:
' There are 10 elements involved per bumper:
' - the shadow of the bumper ( a vpx flasher object)
' - the bumper skirt (primitive)
' - the bumperbase (primitive)
' - a vpx light which colors everything you can see through the bumpertop
' - the bulb (primitive)
' - another vpx light which lights up everything around the bumper
' - the bumpertop (primitive)
' - the VPX bumper object
' - the bumper screws (primitive)
' - the bulb highlight VPX flasher object
' All elements have a special name with the number of the bumper at the end, this is necessary for the fading routine and the initialisation.
' For the bulb and the bumpertop there is a unique material as well per bumpertop.
' To use these bumpers you have to first copy all 10 elements to your table.
' Also export the textures (images) with names that start with "Flbumper" and "Flhighlight" and materials with names that start with "bumper".
' Make sure that all the ten objects are aligned on center, if possible with the exact same x,y coordinates
' After that copy the script (below); also copy the BumperTimer vpx object to your table
' Every bumper needs to be initialised with the FlInitBumper command, see example below;
' Colors available are red, white, blue, orange, yellow, green, purple and blacklight.
' In a GI subroutine you can then call set the bumperlight intensity with the "FlBumperFadeTarget(nr) = value" command
' where nr is the number of the bumper, value is between 0 (off) and 1 (full on) (so you can also use 0.3 0.4 etc).

' Notes:
' - There is only one color for the disk; you can photoshop it to a different color
' - The bumpertops are angle independent up to a degree; my estimate is -45 to + 45 degrees horizontally, 0 (topview) to 70-80 degrees (frontview)
' - I built in correction for the day-night slider; this might not work perfectly, depending on your table lighting
' - These elements, textures and materials do NOT integrate with any of the lighting routines I have seen in use in many VPX tables
'   (just find the GI handling routine and insert the FlBumperFadeTarget statement)
' - If you want to use VPX native bumperdisks just copy my bumperdisk but make it invisible

' prepare some global vars to dim/brighten objects when using day-night slider
Dim DayNightAdjust , DNA30, DNA45, DNA90
If NightDay < 10 Then
  DNA30 = 0
  DNA45 = (NightDay - 10) / 20
  DNA90 = 0
  DayNightAdjust = 0.4
Else
  DNA30 = (NightDay - 10) / 30
  DNA45 = (NightDay - 10) / 45
  DNA90 = (NightDay - 10) / 90
  DayNightAdjust = NightDay / 25
End If

Dim FlBumperFadeActual(6), FlBumperFadeTarget(6), FlBumperColor(6), FlBumperTop(6), FlBumperSmallLight(6), Flbumperbiglight(6)
Dim FlBumperDisk(6), FlBumperBase(6), FlBumperBulb(6), FlBumperscrews(6), FlBumperActive(6), FlBumperHighlight(6)
Dim cnt
For cnt = 1 To 6
  FlBumperActive(cnt) = False
Next

' colors available are red, white, blue, orange, yellow, green, purple and blacklight
'FlInitBumper 1, "red"
'FlInitBumper 2, "white"
'FlInitBumper 3, "blue"
'FlInitBumper 4, "orange"
'FlInitBumper 5, "yellow"

' ### uncomment the statement below to change the color for all bumpers ###
   Dim ind
   For ind = 1 To 5
     FlInitBumper ind, "red"
   Next

Sub FlInitBumper(nr, col)
  FlBumperActive(nr) = True

  ' store all objects in an array for use in FlFadeBumper subroutine
  FlBumperFadeActual(nr) = 1
  FlBumperFadeTarget(nr) = 1.1
  FlBumperColor(nr) = col
  Set FlBumperTop(nr) = Eval("bumpertop" & nr)
  FlBumperTop(nr).material = "bumpertopmat" & nr
  Set FlBumperSmallLight(nr) = Eval("bumpersmalllight" & nr)
  Set Flbumperbiglight(nr) = Eval("bumperbiglight" & nr)
  Set FlBumperDisk(nr) = Eval("bumperdisk" & nr)
  Set FlBumperBase(nr) = Eval("bumperbase" & nr)
  Set FlBumperBulb(nr) = Eval("bumperbulb" & nr)
  FlBumperBulb(nr).material = "bumperbulbmat" & nr
  Set FlBumperscrews(nr) = Eval("bumperscrews" & nr)
  FlBumperscrews(nr).material = "bumperscrew" & col
  Set FlBumperHighlight(nr) = Eval("bumperhighlight" & nr)

  ' set the color for the two VPX lights
  Select Case col
    Case "red"
      FlBumperSmallLight(nr).color = RGB(255,4,0)
      FlBumperSmallLight(nr).colorfull = RGB(255,24,0)
      FlBumperBigLight(nr).color = RGB(255,32,0)
      FlBumperBigLight(nr).colorfull = RGB(255,32,0)
      FlBumperHighlight(nr).color = RGB(64,255,0)
      FlBumperSmallLight(nr).BulbModulateVsAdd = 0.98
      FlBumperSmallLight(nr).TransmissionScale = 0

    Case "blue"
      FlBumperBigLight(nr).color = RGB(32,80,255)
      FlBumperBigLight(nr).colorfull = RGB(32,80,255)
      FlBumperSmallLight(nr).color = RGB(0,80,255)
      FlBumperSmallLight(nr).colorfull = RGB(0,80,255)
      FlBumperSmallLight(nr).TransmissionScale = 0
      MaterialColor "bumpertopmat" & nr, RGB(8,120,255)
      FlBumperHighlight(nr).color = RGB(255,16,8)
      FlBumperSmallLight(nr).BulbModulateVsAdd = 1

    Case "green"
      FlBumperSmallLight(nr).color = RGB(8,255,8)
      FlBumperSmallLight(nr).colorfull = RGB(8,255,8)
      FlBumperBigLight(nr).color = RGB(32,255,32)
      FlBumperBigLight(nr).colorfull = RGB(32,255,32)
      FlBumperHighlight(nr).color = RGB(255,32,255)
      MaterialColor "bumpertopmat" & nr, RGB(16,255,16)
      FlBumperSmallLight(nr).TransmissionScale = 0.005
      FlBumperSmallLight(nr).BulbModulateVsAdd = 1

    Case "orange"
      FlBumperHighlight(nr).color = RGB(255,130,255)
      FlBumperSmallLight(nr).BulbModulateVsAdd = 1
      FlBumperSmallLight(nr).TransmissionScale = 0
      FlBumperSmallLight(nr).color = RGB(255,130,0)
      FlBumperSmallLight(nr).colorfull = RGB (255,90,0)
      FlBumperBigLight(nr).color = RGB(255,190,8)
      FlBumperBigLight(nr).colorfull = RGB(255,190,8)

    Case "white"
      FlBumperBigLight(nr).color = RGB(255,230,190)
      FlBumperBigLight(nr).colorfull = RGB(255,230,190)
      FlBumperHighlight(nr).color = RGB(255,180,100)
      FlBumperSmallLight(nr).TransmissionScale = 0
      FlBumperSmallLight(nr).BulbModulateVsAdd = 0.99

    Case "blacklight"
      FlBumperBigLight(nr).color = RGB(32,32,255)
      FlBumperBigLight(nr).colorfull = RGB(32,32,255)
      FlBumperHighlight(nr).color = RGB(48,8,255)
      FlBumperSmallLight(nr).TransmissionScale = 0
      FlBumperSmallLight(nr).BulbModulateVsAdd = 1

    Case "yellow"
      FlBumperSmallLight(nr).color = RGB(255,230,4)
      FlBumperSmallLight(nr).colorfull = RGB(255,230,4)
      FlBumperBigLight(nr).color = RGB(255,240,50)
      FlBumperBigLight(nr).colorfull = RGB(255,240,50)
      FlBumperHighlight(nr).color = RGB(255,255,220)
      FlBumperSmallLight(nr).BulbModulateVsAdd = 1
      FlBumperSmallLight(nr).TransmissionScale = 0

    Case "purple"
      FlBumperBigLight(nr).color = RGB(80,32,255)
      FlBumperBigLight(nr).colorfull = RGB(80,32,255)
      FlBumperSmallLight(nr).color = RGB(80,32,255)
      FlBumperSmallLight(nr).colorfull = RGB(80,32,255)
      FlBumperSmallLight(nr).TransmissionScale = 0
      FlBumperHighlight(nr).color = RGB(32,64,255)
      FlBumperSmallLight(nr).BulbModulateVsAdd = 1
  End Select
End Sub

Sub FlFadeBumper(nr, Z)
  FlBumperBase(nr).BlendDisableLighting = 0.5 * DayNightAdjust
  '   UpdateMaterial(string, float wrapLighting, float roughness, float glossyImageLerp, float thickness, float edge, float edgeAlpha, float opacity,
  '        OLE_COLOR base, OLE_COLOR glossy, OLE_COLOR clearcoat, VARIANT_BOOL isMetal, VARIANT_BOOL opacityActive,
  '        float elasticity, float elasticityFalloff, float friction, float scatterAngle) - updates all parameters of a material
  FlBumperDisk(nr).BlendDisableLighting = (0.5 - Z * 0.3 ) * DayNightAdjust

  Select Case FlBumperColor(nr)
    Case "blue"
      UpdateMaterial "bumperbulbmat" & nr, 0, 0.75 , 0, 1 - Z, 1 - Z, 1 - Z, 0.9999, RGB(38 - 24 * Z,130 - 98 * Z,255), RGB(255,255,255), RGB(32,32,32), False, True, 0, 0, 0, 0
      FlBumperSmallLight(nr).intensity = 20 + 500 * Z / (0.5 + DNA30)
      FlBumperTop(nr).BlendDisableLighting = 3 * DayNightAdjust + 50 * Z
      FlBumperBulb(nr).BlendDisableLighting = 12 * DayNightAdjust + 5000 * (0.03 * Z + 0.97 * Z ^ 3)
      Flbumperbiglight(nr).intensity = 25 * Z / (1 + DNA45)
      FlBumperHighlight(nr).opacity = 10000 * (Z ^ 3) / (0.5 + DNA90)

    Case "green"
      UpdateMaterial "bumperbulbmat" & nr, 0, 0.75 , 0, 1 - Z, 1 - Z, 1 - Z, 0.9999, RGB(16 + 16 * Sin(Z * 3.14),255,16 + 16 * Sin(Z * 3.14)), RGB(255,255,255), RGB(32,32,32), False, True, 0, 0, 0, 0
      FlBumperSmallLight(nr).intensity = 10 + 150 * Z / (1 + DNA30)
      FlBumperTop(nr).BlendDisableLighting = 2 * DayNightAdjust + 20 * Z
      FlBumperBulb(nr).BlendDisableLighting = 7 * DayNightAdjust + 6000 * (0.03 * Z + 0.97 * Z ^ 10)
      Flbumperbiglight(nr).intensity = 10 * Z / (1 + DNA45)
      FlBumperHighlight(nr).opacity = 6000 * (Z ^ 3) / (1 + DNA90)

    Case "red"
      UpdateMaterial "bumperbulbmat" & nr, 0, 0.75 , 0, 1 - Z, 1 - Z, 1 - Z, 0.9999, RGB(255, 16 - 11 * Z + 16 * Sin(Z * 3.14),0), RGB(255,255,255), RGB(32,32,32), False, True, 0, 0, 0, 0
      FlBumperSmallLight(nr).intensity = 17 + 100 * Z / (1 + DNA30 ^ 2)
      FlBumperTop(nr).BlendDisableLighting = 3 * DayNightAdjust + 18 * Z / (1 + DNA90)
      FlBumperBulb(nr).BlendDisableLighting = 20 * DayNightAdjust + 9000 * (0.03 * Z + 0.97 * Z ^ 10)
      Flbumperbiglight(nr).intensity = 10 * Z / (1 + DNA45)
      FlBumperHighlight(nr).opacity = 2000 * (Z ^ 3) / (1 + DNA90)
      MaterialColor "bumpertopmat" & nr, RGB(255,20 + Z * 4,8 - Z * 8)

    Case "orange"
      UpdateMaterial "bumperbulbmat" & nr, 0, 0.75 , 0, 1 - Z, 1 - Z, 1 - Z, 0.9999, RGB(255, 100 - 22 * z + 16 * Sin(Z * 3.14),Z * 32), RGB(255,255,255), RGB(32,32,32), False, True, 0, 0, 0, 0
      FlBumperSmallLight(nr).intensity = 17 + 250 * Z / (1 + DNA30 ^ 2)
      FlBumperTop(nr).BlendDisableLighting = 3 * DayNightAdjust + 50 * Z / (1 + DNA90)
      FlBumperBulb(nr).BlendDisableLighting = 15 * DayNightAdjust + 2500 * (0.03 * Z + 0.97 * Z ^ 10)
      Flbumperbiglight(nr).intensity = 10 * Z / (1 + DNA45)
      FlBumperHighlight(nr).opacity = 4000 * (Z ^ 3) / (1 + DNA90)
      MaterialColor "bumpertopmat" & nr, RGB(255,100 + Z * 50, 0)

    Case "white"
      UpdateMaterial "bumperbulbmat" & nr, 0, 0.75 , 0, 1 - Z, 1 - Z, 1 - Z, 0.9999, RGB(255,230 - 100 * Z, 200 - 150 * Z), RGB(255,255,255), RGB(32,32,32), False, True, 0, 0, 0, 0
      FlBumperSmallLight(nr).intensity = 20 + 180 * Z / (1 + DNA30)
      FlBumperTop(nr).BlendDisableLighting = 5 * DayNightAdjust + 30 * Z
      FlBumperBulb(nr).BlendDisableLighting = 18 * DayNightAdjust + 3000 * (0.03 * Z + 0.97 * Z ^ 10)
      Flbumperbiglight(nr).intensity = 8 * Z / (1 + DNA45)
      FlBumperHighlight(nr).opacity = 1000 * (Z ^ 3) / (1 + DNA90)
      FlBumperSmallLight(nr).color = RGB(255,255 - 20 * Z,255 - 65 * Z)
      FlBumperSmallLight(nr).colorfull = RGB(255,255 - 20 * Z,255 - 65 * Z)
      MaterialColor "bumpertopmat" & nr, RGB(255,235 - z * 36,220 - Z * 90)

    Case "blacklight"
      UpdateMaterial "bumperbulbmat" & nr, 0, 0.75 , 0, 1 - Z, 1 - Z, 1 - Z, 1, RGB(30 - 27 * Z ^ 0.03,30 - 28 * Z ^ 0.01, 255), RGB(255,255,255), RGB(32,32,32), False, True, 0, 0, 0, 0
      FlBumperSmallLight(nr).intensity = 20 + 900 * Z / (1 + DNA30)
      FlBumperTop(nr).BlendDisableLighting = 3 * DayNightAdjust + 60 * Z
      FlBumperBulb(nr).BlendDisableLighting = 15 * DayNightAdjust + 30000 * Z ^ 3
      Flbumperbiglight(nr).intensity = 25 * Z / (1 + DNA45)
      FlBumperHighlight(nr).opacity = 2000 * (Z ^ 3) / (1 + DNA90)
      FlBumperSmallLight(nr).color = RGB(255 - 240 * (Z ^ 0.1),255 - 240 * (Z ^ 0.1),255)
      FlBumperSmallLight(nr).colorfull = RGB(255 - 200 * z,255 - 200 * Z,255)
      MaterialColor "bumpertopmat" & nr, RGB(255 - 190 * Z,235 - z * 180,220 + 35 * Z)

    Case "yellow"
      UpdateMaterial "bumperbulbmat" & nr, 0, 0.75 , 0, 1 - Z, 1 - Z, 1 - Z, 0.9999, RGB(255, 180 + 40 * z, 48 * Z), RGB(255,255,255), RGB(32,32,32), False, True, 0, 0, 0, 0
      FlBumperSmallLight(nr).intensity = 17 + 200 * Z / (1 + DNA30 ^ 2)
      FlBumperTop(nr).BlendDisableLighting = 3 * DayNightAdjust + 40 * Z / (1 + DNA90)
      FlBumperBulb(nr).BlendDisableLighting = 12 * DayNightAdjust + 2000 * (0.03 * Z + 0.97 * Z ^ 10)
      Flbumperbiglight(nr).intensity = 10 * Z / (1 + DNA45)
      FlBumperHighlight(nr).opacity = 1000 * (Z ^ 3) / (1 + DNA90)
      MaterialColor "bumpertopmat" & nr, RGB(255,200, 24 - 24 * z)

    Case "purple"
      UpdateMaterial "bumperbulbmat" & nr, 0, 0.75 , 0, 1 - Z, 1 - Z, 1 - Z, 0.9999, RGB(128 - 118 * Z - 32 * Sin(Z * 3.14), 32 - 26 * Z ,255), RGB(255,255,255), RGB(32,32,32), False, True, 0, 0, 0, 0
      FlBumperSmallLight(nr).intensity = 15 + 200 * Z / (0.5 + DNA30)
      FlBumperTop(nr).BlendDisableLighting = 3 * DayNightAdjust + 50 * Z
      FlBumperBulb(nr).BlendDisableLighting = 15 * DayNightAdjust + 10000 * (0.03 * Z + 0.97 * Z ^ 3)
      Flbumperbiglight(nr).intensity = 25 * Z / (1 + DNA45)
      FlBumperHighlight(nr).opacity = 4000 * (Z ^ 3) / (0.5 + DNA90)
      MaterialColor "bumpertopmat" & nr, RGB(128 - 60 * Z,32,255)
  End Select
End Sub

Sub BumperTimer_Timer
  Dim nr
  For nr = 1 To 6
    If FlBumperFadeActual(nr) < FlBumperFadeTarget(nr) And FlBumperActive(nr)  Then
      FlBumperFadeActual(nr) = FlBumperFadeActual(nr) + (FlBumperFadeTarget(nr) - FlBumperFadeActual(nr)) * 0.8
      If FlBumperFadeActual(nr) > 0.99 Then FlBumperFadeActual(nr) = 1
      FlFadeBumper nr, FlBumperFadeActual(nr)
    End If
    If FlBumperFadeActual(nr) > FlBumperFadeTarget(nr) And FlBumperActive(nr)  Then
      FlBumperFadeActual(nr) = FlBumperFadeActual(nr) + (FlBumperFadeTarget(nr) - FlBumperFadeActual(nr)) * 0.4 / (FlBumperFadeActual(nr) + 0.1)
      If FlBumperFadeActual(nr) < 0.01 Then FlBumperFadeActual(nr) = 0
      FlFadeBumper nr, FlBumperFadeActual(nr)
    End If
  Next
End Sub

'******************************************************
'******  END FLUPPER BUMPERS
'******************************************************

Sub PlayfieldKicker_UnHit()
  If Not CurrentAct = 6 Then Exit Sub
    If Not MercSaveActive Then townportalPF.Visible = False
    PlayfieldKicker.Enabled = False
End Sub
'Function PortalCap()
'    PortalCap = 120000
'End Function

Function CowLevelMultiplier()
    If CurrentAct = 6 Then CowLevelMultiplier = 2 Else CowLevelMultiplier = 1
End Function

'Sub UpdatePortalDMD()
'    Dim secs : secs = Int(PortalTime / 1000)
'    Dim filled : filled = Int((PortalTime / PortalCap()) * 10)
'    If filled > 10 Then filled = 10
'    If filled < 0 Then filled = 0
'    Dim bar : bar = ""
'    Dim bi
'    For bi = 1 To 10
'        If bi <= filled Then bar = bar & "=" Else bar = bar & "."
'    Next
'    FlexDMD.Stage.GetGroup("Score").GetLabel("Line2").Font = FontSmall
'    FlexDMD.Stage.GetGroup("Score").GetLabel("Line2").Text = bar & " " & secs & "s"
'    FlexDMD.Stage.GetGroup("Score").GetLabel("Line2").SetAlignedPosition 64, 23, FlexDMD_Align_Center
'    FlexDMD.Stage.GetGroup("Score").GetLabel("Line3").Text = ""
'    FlexDMD.Stage.GetGroup("Score").GetLabel("Line4").Text = ""
'End Sub

'Sub PortalTimer_Timer()
'    If Not PortalActive Or CurrentAct <> 6 Then
'        PortalTimer.Enabled = False
'        Exit Sub
'    End If
'    PortalTime = PortalTime - 250
'    If PortalTime <= 10000 And PortalTime > 0 Then
'        SetFlasherColor 3, 180, 0, 0 : SetFlasherColor 4, 180, 0, 0
'        FireFlasher 3 : FireFlasher 4
'    End If
'    UpdatePortalDMD
'    If PortalTime <= 0 Then
'        PortalTime = 0
'        PortalTimer.Enabled = False
'        PortalCollapse
'    End If
'End Sub

'Sub PortalCollapse()
'    WasActSixComplete = True
'    PortalActive = False
'    PortalCollapsed = True
'    StopSound CurrentSong
'    CurrentSong = ""
'    ShowMysteryDMD "PORTAL COLLAPSED!", "YOUR JOURNEY ENDS"
'    PlayCallout "CainCowComplete", 4000
'    DimAllFlashers
'    SetFlasherColor 1, 180, 0, 0 : SetFlasherColor 2, 180, 0, 0
'    SetFlasherColor 3, 180, 0, 0 : SetFlasherColor 4, 180, 0, 0
'    SetFlasherColor 5, 180, 0, 0 : SetFlasherColor 6, 180, 0, 0
'    FireAllFlashers
'    GIEventMode = GI_MODE_NONE
'    GIEventTimer.Enabled = False
'    RestoreAllGIRows
'    CowNeutralTimer.Enabled = False
'    GameActive = False
' StartEndOfBallBonus
' BonusActive = True
'    BonusStep = 0
'    BonusTimer.Interval = 2000
'    BonusTimer.Enabled = True
'
'    PortalCollapseTimer.Interval = 3000
'    PortalCollapseTimer.Enabled = True
'End Sub

'Sub PortalCollapseTimer_Timer()
'    PortalCollapseTimer.Enabled = False
'    PortalCollapsed = False
'    PortalActive = False
'    PortalSurge = 0
'    CowKingKillCount = 0
'    BallNumber = BallNumber - 1
'    WasActSixComplete = True
'    StartEndOfBallBonus
'End Sub

'Sub AddPortalSurge(seconds)
'    PortalSurge = PortalSurge + seconds
'    If PortalSurge > 120 Then PortalSurge = 120
'    ShowMessage "PORTAL +" & seconds & "s BANKED"
'End Sub

Dim MysteryKickHolding : MysteryKickHolding = False
Dim NarniaBallActive : NarniaBallActive = False

Sub CheckNarniaBalls()
    Dim BOT : BOT = GetBalls()
    Dim b
    For Each b In BOT
        If b.ID <> CapBallID And b.ID <> CapBall2ID Then
            If b.Z > 30 And Not LeapReady And Not MysteryKickHolding Then
                If Abs(b.VelZ) < 0.5 And Abs(b.VelX) < 0.5 And Abs(b.VelY) < 0.5 Then
                    b.X = BallRelease.X
                    b.Y = BallRelease.Y
                    b.Z = 20
                    b.VelX = 0 : b.VelY = 0 : b.VelZ = 0
                    NarniaBallActive = True
                    NarniaReleaseTimer.Interval = 100
                    NarniaReleaseTimer.Enabled = True
                    Exit For  ' one ball at a time
                End If
            End If
        End If
    Next
End Sub

Sub NarniaReleaseTimer_Timer()
    NarniaReleaseTimer.Enabled = False
    BallRelease.Kick 90, 7
    PlaySoundAtLevelStatic SoundFX("BallRelease" & Int(Rnd * 7) + 1, DOFContactors), BallReleaseSoundLevel, BallRelease
End Sub

Sub PauseKillStreakTimer()
    If KillStreakCount < 5 Then Exit Sub
    If GetBIP() > 1 Then Exit Sub
    KillStreakTimer.Enabled = False
End Sub

Sub ResumeKillStreakTimer()
    If KillStreakCount < 5 Then Exit Sub
    KillStreakTimer.Enabled = True
End Sub

Sub PauseAmbushTimer()
    If Not AmbushActive Then Exit Sub
    If GetBIP() > 1 Then Exit Sub
    AmbushHurryTimer.Enabled = False
    AmbushDMDPaused = True
End Sub

Sub ResumeAmbushTimer()
    If Not AmbushActive Then Exit Sub
    AmbushDMDPaused = False
    AmbushHurryTimer.Enabled = True
    UpdateAmbushDMD
End Sub

Dim TiltWarnings    : TiltWarnings    = 0
Dim TiltActive      : TiltActive      = False
Dim TiltSensitivity : TiltSensitivity = 3
Dim TiltCooldown : TiltCooldown = False
Dim NudgeCount      : NudgeCount      = 0
Const NudgesPerWarning = 2     ' nudges within window to earn 1 warning
Const NudgeWindowMs    = 1000  ' rolling window in ms — spread nudges wider than this and no warning

Sub CheckTilt()
    If Not GameActive Or TiltActive Then Exit Sub
    NudgeCount = NudgeCount + 1
    ' Reset the window on every nudge — window only expires if you stop nudging
    NudgeWindowTimer.Enabled = False
    NudgeWindowTimer.Interval = NudgeWindowMs
    NudgeWindowTimer.Enabled = True
    If NudgeCount >= NudgesPerWarning Then
        NudgeCount = 0
        NudgeWindowTimer.Enabled = False
        TiltWarnings = TiltWarnings + 1
        If TiltWarnings >= TiltSensitivity Then
            DoTilt
        Else
            ShowMessage "WARNING!"
            Select Case Int(Rnd * 2)
                Case 0 : PlayCallout "WatchIt", 2000
                Case 1 : PlayCallout "Careful", 2000
            End Select
        End If
    End If
End Sub

Sub NudgeWindowTimer_Timer()
    NudgeWindowTimer.Enabled = False
    NudgeCount = 0
End Sub

Sub TiltCooldownTimer_Timer()
    TiltCooldownTimer.Enabled = False
    TiltCooldown = False
End Sub

Sub DoTilt()
    If Not GameActive Or TiltActive Then Exit Sub
  TiltDuckAudio
    TiltActive = True
'    ShowMessage "TILT!"
    PlaySound "tilt", 0, 1
    ' Kill all active modes
    BallSaveActive = False
    BallSaveTimer.Enabled = False
    BallSaveL.State = 0
    BallSaveL2.State = 0
    MercSaveActive = False
    PartyMultiballRunning = False
    RuneWordMultiballRunning = False
    PartySpawnTimer.Enabled = False
    RuneWordSpawnTimer.Enabled = False
    RuneKickBlinkTimer.Enabled = False
    PortalOpenTimer.Enabled = False
    PortalFlashTimer.Enabled = False
    MercSpawnTimer.Enabled = False
    StopCritCycle
    BossRegenTimer.Enabled = False
    CritExpireTimer.Enabled = False
  CritCharged = 0
    InstantKillActive = False
    InstantKillTimer.Enabled = False
    AmbushActive = False
    AmbushLightTimer.Enabled = False
    AmbushHurryTimer.Enabled = False
    AmbushHoldTimer.Enabled = False
    AmbushTriggerNum = 0
    QKick1.Kick 180, 15 : QKick1.Enabled = False        ' ← eject + disarm so the ambush ball can't strand the game
    QKick2.Kick 180, 15 : QKick2.Enabled = False
    QKick4.Kick 180, 15 : QKick4.Enabled = False
    ' Kill GI for tilt
    GIEventMode = GI_MODE_NONE
    GIEventTimer.Enabled = False
    GIFlickerTimer.Enabled = False
    Dim tgi
    For tgi = 0 To GI.Count - 1
        GI.Item(tgi).Intensity = 0
        GI.Item(tgi).Color = RGB(0, 0, 0)
        GI.Item(tgi).ColorFull = RGB(0, 0, 0)
    Next
    UpdateDMD2 "TILT!", ""
End Sub

'*****************************************
'           TURNTABLES
'*****************************************
Dim TT1Angle : TT1Angle = 0
Dim TT2Angle : TT2Angle = 0
Dim TurntableActive : TurntableActive = False
Dim TT1Dir : TT1Dir = 1
Dim TT2Dir : TT2Dir = -1

Sub StartTurntables()
    TurntableActive = False
    TT1Angle = 0
    TT2Angle = 0
    Turntable1Trigger.Enabled = True
    Turntable2Trigger.Enabled = True
    TurntableRandomTimer.Interval = 1000 + Int(Rnd * 2000)
    TurntableRandomTimer.Enabled = True
End Sub

Sub StopTurntables()
    TurntableRandomTimer.Enabled = False
    TurntableActive = False
    Turntable1Trigger.Enabled = False
    Turntable2Trigger.Enabled = False
    TurntableSpinTimer.Enabled = False
End Sub

Sub TurntableRandomTimer_Timer()
If Not BossFightActive And GameActive Then
        StopTurntables
        Exit Sub
    End If
    If TurntableActive Then
        TurntableActive = False
        TurntableSpinTimer.Enabled = False
        TurntableRandomTimer.Interval = 2000 + Int(Rnd * 3000)
    Else
        TurntableActive = True
        TT1Dir = 1 : If Rnd > 0.5 Then TT1Dir = -1
        TT2Dir = 1 : If Rnd > 0.5 Then TT2Dir = -1
        TurntableSpinTimer.Enabled = True
        TurntableRandomTimer.Interval = 1000 + Int(Rnd * 2000)
    End If
End Sub

Sub TurntableSpinTimer_Timer()
    If Not TurntableActive And GameActive Then
        TurntableSpinTimer.Enabled = False
        Exit Sub
    End If
    TT1Angle = (TT1Angle + (10 * TT1Dir)) Mod 360
    TT2Angle = (TT2Angle + (10 * TT2Dir)) Mod 360
    Turntable1Prim.RotY = TT1Angle
    Turntable2Prim.RotY = TT2Angle
End Sub

Sub Turntable1Trigger_Hit()
    If Not TurntableActive Then Exit Sub
    ActiveBall.VelX = ActiveBall.VelX + (6 * TT1Dir)
    ActiveBall.VelY = ActiveBall.VelY - (6 * TT1Dir)
End Sub

Sub Turntable2Trigger_Hit()
    If Not TurntableActive Then Exit Sub
    ActiveBall.VelX = ActiveBall.VelX - (6 * TT2Dir)
    ActiveBall.VelY = ActiveBall.VelY + (6 * TT2Dir)
End Sub
'
''*****************************************
''        BACKGLASS LIGHTS
''*****************************************
'Dim BGLightLevel : BGLightLevel = 0
'
'Sub BGFlashColor(r, g, b)
'    BGFlasher1.Color = RGB(r, g, b)
'    BGFlasher2.Color = RGB(r, g, b)
'    BGFlasher3.Color = RGB(r, g, b)
'End Sub
'
'Sub BGFlashOpacity(o)
'    BGFlasher1.Opacity = o
'    BGFlasher2.Opacity = o
'    BGFlasher3.Opacity = o
'End Sub
'
'Sub BGFlashVisible(b)
'    BGFlasher1.Visible = b
'    BGFlasher2.Visible = b
'    BGFlasher3.Visible = b
'End Sub
'
'Sub BackglassFlash(r, g, b)
'    BGFlashColor r, g, b : BGFlashVisible True
'    BGLightLevel = 1
'    BGFlashTimer.Enabled = True
'End Sub
'
'Sub BGFlashTimer_Timer()
'    BGLightLevel = BGLightLevel * 0.90 - 0.02
'    Dim lvl : lvl = BGLightLevel * BGLightLevel * BGLightLevel
'    If lvl < 0 Then lvl = 0
'    BGFlashOpacity 120 * lvl
'    If BGLightLevel <= 0 Then
'        BGFlashTimer.Enabled = False
'        BGFlashOpacity 0
'    End If
'End Sub
'
'Dim BGPulseRemaining : BGPulseRemaining = 0
'Dim BGPulseUp        : BGPulseUp        = True
'Dim BGPulseR : Dim BGPulseG : Dim BGPulseB
'
'Sub BackglassPulse(durationMs, r, g, b)
'    BGPulseR = r : BGPulseG = g : BGPulseB = b
'    BGPulseRemaining = durationMs
'    BGPulseUp = True
'    BGLightLevel = 0
'    BGFlashColor r, g, b : BGFlashVisible True
'    BGFlashTimer.Enabled = False   ' stop any decay flash
'    BGPulseTimer.Interval = 60
'    BGPulseTimer.Enabled = True
'End Sub
'
'Sub BGPulseTimer_Timer()
'    BGPulseRemaining = BGPulseRemaining - 60
'    ' Oscillate the level up and down to mimic speech cadence
'    If BGPulseUp Then
'        BGLightLevel = BGLightLevel + 0.35
'        If BGLightLevel >= 1 Then BGLightLevel = 1 : BGPulseUp = False
'    Else
'        BGLightLevel = BGLightLevel - 0.35
'        If BGLightLevel <= 0.2 Then BGLightLevel = 0.2 : BGPulseUp = True
'    End If
'    BGFlashOpacity 120 * BGLightLevel
'    If BGPulseRemaining <= 0 Then
'        BGPulseTimer.Enabled = False
'        BGFlashOpacity 0
'    End If
'End Sub

''''''''''''''''''''''''''''
' RAMP AOE SYSTEM
''''''''''''''''''''''''''''

Sub ChargeRampAOE()
    If Not GameActive Then Exit Sub
    If WhirlwindActive Then
        FireRampAOE          ' combo: re-hit relaunches, no re-charge needed
        Exit Sub
    End If
    RampAoeCount = RampAoeCount + 1
    UpdateRampAoeLight
    If RampAoeCount >= RampAoeTarget Then
        RampAoeCount = 0
        FireRampAOE
    Else
        UpdateDMD2 "WHIRLWIND CHARGING", RampAoeCount & "/" & RampAoeTarget
        MsgQueueTimer.Enabled = False
        MsgQueueTimer.Enabled = True
        If RampAoeCount = RampAoeTarget - 1 Then PlayCallout "mystery_whirlwindready", 2000
    End If
End Sub

Sub UpdateRampAoeLight()
    QuestLight3.Color = RGB(255, 255, 255) : QuestLight3.ColorFull = RGB(255, 255, 255)
    Select Case RampAoeCount
        Case 0    : QuestLight3.State = 0
        Case 1    : QuestLight3.BlinkInterval = 400 : QuestLight3.BlinkPattern = "10" : QuestLight3.State = 2
        Case 2    : QuestLight3.BlinkInterval = 200 : QuestLight3.BlinkPattern = "10" : QuestLight3.State = 2
        Case Else : QuestLight3.BlinkInterval = 90  : QuestLight3.BlinkPattern = "10" : QuestLight3.State = 2
    End Select
End Sub

Sub FireRampAOE()
    Dim batch
    batch = WhirlwindFloorHits + (GetGearDamage() \ 3)

    If WhirlwindActive Then
        ' skill combo — queue another batch onto the running storm, don't restart it
        WhirlwindCombo = WhirlwindCombo + 1
        WhirlwindHitsThis = WhirlwindHitsThis + batch
        WhirlwindTimer.Interval = WhirlwindMs \ batch
        WhirlwindTimer.Enabled = True
        PlaySound "whirlwind1", 0, 1
        UpdateDMD2 "WHIRLWIND x" & WhirlwindCombo & "!", "+" & batch & " HITS"
        Exit Sub
    End If

    WhirlwindActive = True
    WhirlwindCombo = 1
    WhirlwindHitsThis = batch
    PlaySound "whirlwind1", 0, 1
    UpdateDMD2 "WHIRLWIND!", WhirlwindHitsThis & " HITS"
    QuestLight3.Color = RGB(255, 255, 255) : QuestLight3.ColorFull = RGB(255, 255, 255)
    QuestLight3.BlinkInterval = 60 : QuestLight3.BlinkPattern = "10" : QuestLight3.State = 2
    WhirlwindStep = 0
    WhirlwindTimer.Interval = WhirlwindMs \ WhirlwindHitsThis
    WhirlwindTimer.Enabled = True
End Sub

Sub WhirlwindTimer_Timer()
    WhirlwindStep = WhirlwindStep + 1

    ' grace tick elapsed with no combo re-hit — end the storm
    If WhirlwindStep > WhirlwindHitsThis Then
        WhirlwindTimer.Enabled = False
        WhirlwindActive = False
        WhirlwindCombo = 0
        FlasherSweepTimer.Enabled = False
        Dim f : For f = 1 To 6 : DimFlasher f : Next
        QuestLight3.State = 0
        Exit Sub
    End If

    ' one strike on a random live bumper
    Dim tries, slot
    For tries = 0 To 9
        slot = Int(Rnd * 5)
        If BumperActive(slot) Then
            RandomSoundBumperTop SlotToBumper(slot)
            PlaySound "sword" & (Int(Rnd * 6) + 1), 0, 1
            KillBumper slot
            FlBumperFadeTarget(slot + 1) = 1
            SlotToBumper(slot).TimerEnabled = True
            Exit For
        End If
    Next

    FlasherSweepStep = 0
    FlasherSweepTimer.Interval = 30
    FlasherSweepTimer.Enabled = True

    ' last strike fired — hold the 1s combo window open before ending
    If WhirlwindStep = WhirlwindHitsThis Then
        WhirlwindTimer.Interval = WhirlwindGraceMs
    End If
End Sub

Sub FlasherSweepTimer_Timer()
    FlasherSweepStep = FlasherSweepStep + 1
    If FlasherSweepStep <= 6 Then
        SetFlasherColor FlasherSweepStep, 255, 255, 255
        FireFlasher FlasherSweepStep
    Else
        FlasherSweepTimer.Enabled = False
        Dim f : For f = 1 To 6 : DimFlasher f : Next
    End If
End Sub

''-------------------endless ball non-rom
'Dim keydelay
'Dim LFPressTime, LFReleaseTime
'Dim RFPressTime, RFReleaseTime
'
'keydelay = 10
'
'
'Sub LeftFlipper_Collide(parm)
'
'    LFPressTime   = GameTime + keydelay
'    LFReleaseTime = GameTime + keydelay + 200
'
'End Sub
'
'
'Sub RightFlipper_Collide(parm)
'
'    RFPressTime   = GameTime + keydelay
'    RFReleaseTime = GameTime + keydelay + 200
'
'End Sub
'
'LeftFlipper.timerenabled=1
'LeftFlipper.timerinterval=10
'
'
'Sub LeftFlipper_Timer()
'
'    If LFPressTime > 0 And GameTime >= LFPressTime Then
'        Table1_KeyDown LeftFlipperKey
'        LFPressTime = 0
'    End If
'
'    If LFReleaseTime > 0 And GameTime >= LFReleaseTime Then
'        Table1_KeyUp LeftFlipperKey
'        LFReleaseTime = 0
'    End If
'
'End Sub
'
'RightFlipper.timerenabled=1
'RightFlipper.timerinterval=10
'
'Sub RightFlipper_Timer()
'
'    If RFPressTime > 0 And GameTime >= RFPressTime Then
'        Table1_KeyDown RightFlipperKey
'        RFPressTime = 0
'    End If
'
'    If RFReleaseTime > 0 And GameTime >= RFReleaseTime Then
'        Table1_KeyUp RightFlipperKey
'        RFReleaseTime = 0
'    End If
'
'End Sub
'
'
'Sub Drain_Hit()
'
'    ActiveBall.X = 640
'    ActiveBall.Y = 1219
'    ActiveBall.Z = 50
'
'    Drain.Kick 0,0
'End sub

Sub DebugAudioFlush()
    Dim x
    For x = 0 To 13
        StopSound "RampLoop" & x
        StopSound "wireloop" & x
        StopSound "BallRoll_" & x
    Next
    StopSound "hellbrazierloop2"
    StopSound "TOM_Trunk_Motor_Long"
    StopSound "whirlwind1"
    ShowMessage "AUDIO FLUSHED"
End Sub


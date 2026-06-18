'*
'*        Gottlieb's 300 (1975)
'*        http://www.ipdb.org/machine.cgi?id=2539
'*


' 001 - scampa123   - Added nFozzy flippers and physics
' 002 - scampa123   - Added Fleep sound package
' 003 - scampa123   - Added JP's LUT code with 15 LUT images
' 004 - scampa123   - Added Minimal VR Room
' 005 - scampa123   - Fixed rubbers on left hand side for scoring and converted Kicker
' 006 - Rawd/Joel   - Added VR Room Backglass components
' 007 - Hauntfreaks - Added redraw of playfield
' 008 - Hauntfreaks - Added redraw of plastics
' 009 - Hauntfreaks - Added new desktop backdrop
' 010 - Hauntfreaks - Added cleaned the original directb2s
' 011 - Hauntfreaks - Added enhanced lighting and shadows
' 012 - scampa123   - Added missing triggers on left side of playfield for additional scoring to match the real table
' 013 - scampa123   - Fixed Flipper-buzz Tilt bug
' 014 - scampa123   - Added new VR Gottlieb Coin Door
' 015 - scampa123   - Added new VR Backbox to match real table
' 016 - scampa123   - Added new VR Carpet-changing code with 5 carpets to customize VR room
' 017 - scampa123   - Fixed playfiled multiplier and star rollover lights which remained on after game over


option explicit
Randomize
ExecuteGlobal GetTextFile("core.vbs")

On Error Resume Next
ExecuteGlobal GetTextFile("Controller.vbs")
If Err Then MsgBox "Unable to open Controller.vbs. Ensure that it is in the scripts folder."
On Error Goto 0

Const cGameName = "Gottlieb300_1975"
Const ShadowConfigFile = false

Dim Controller  ' B2S
Dim B2SScore  ' B2S Score Displayed

Const HSFileName="300_75VPX.txt"
Const B2STableName="Gottlieb300_1975"
Const LMEMTableConfig="LMEMTables.txt"
Const LMEMShadowConfig="LMEMShadows.txt"
Dim EnableBallShadow
Dim EnableFlipperShadow

Dim B2SOn   'True/False if want backglass

Dim VRBGBall1,VRBGBall2,VRBGBall3,VRBGBall4,VRBGBall5,VRBGBall6,VRBGBall7,VRBGBall8,VRBGBall9,VRBGBall10
Dim GameBall



'******************************************************************
'             TABLE OPTIONS
'******************************************************************

'******************************************************************
'   VR Room Setup
'******************************************************************

Dim VRRoom, Object

VRRoom = 0  '0=off, 1==Minimal



Const ShadowFlippersOn = true
Const ShadowBallOn = true

'* this value adjusts score motor behavior - 1 allows you to continue scoring while the score motor is running - 1 sets score motor to behave more like a real EM
Const ScoreMotorAdjustment=0

'* this is a debug setting to use an older scoring routine vs a newer score routine - don't change this value
Const ScoreAdditionAdjustment=1

'* this controls whether you hear bells (0) or chimes (1) when scoring
Const ChimesOn=1

'******************************************************************
'///////////////////////-----Fleep General Sound Options-----/////
'******************************************************************
'// VolumeDial:
'// VolumeDial is the actual global volume multiplier for the fleep mechanical sounds.
'// Values smaller than 1 will decrease mechanical sounds volume.
'// Recommended values should be no greater than 1.
Const VolumeDial = 0.1

'******************************************************************
'///////////////////////-----Phsyics Mods-----/////////////////////
'******************************************************************
Const RubberizerEnabled = 1     '0 = normal flip rubber, 1 = more lively rubber for flips
Const FlipperCoilRampupMode = 0     '0 = fast, 1 = medium, 2 = slow (tap passes should work)
Const TargetBouncerEnabled = 1    '0 = normal standup targets, 1 = bouncy targets
Const TargetBouncerFactor = 0.7   'Level of bounces. 0.2 - 1.5 are probably usable values.


'******************************************************************
'             END TABLE OPTIONS
'******************************************************************


'**************************************************************************
'   SCAMPA123'S VR-Carpet Changing Code for changing with magna save-button
'   This is based on JP's LUT code
'**************************************************************************

Dim bVRCarpetActive, VRCarpetImage
Dim VRCPT
Sub LoadVRCarpet
  bVRCarpetActive = False
    VRCPT = LoadValue(cGameName, "VRCarpetImage")
    If(VRCPT <> "") Then VRCarpetImage = VRCPT Else VRCarpetImage = 0
  UpdateVRCarpet
End Sub

Sub SaveVRCarpet
    SaveValue cGameName, "VRCarpetImage", VRCarpetImage
End Sub

Sub NextVRCarpet: VRCarpetImage = (VRCarpetImage +1 ) MOD 5: UpdateVRCarpet: SaveVRCarpet: End Sub

Sub UpdateVRCarpet
Select Case VRCarpetImage
Case 0: VR_Floor.Image = "bowlingcarpet1"
Case 1: VR_Floor.Image = "bowlingcarpet2"
Case 2: VR_Floor.Image = "bowlingcarpet3"
Case 3: VR_Floor.Image = "bowlingcarpet4"
Case 4: VR_Floor.Image = "bowlingcarpet5"
End Select
End Sub


'******************************************************************
'   JP'S LUT-files for changing brightness with magna save-buttons
'******************************************************************

Dim bLutActive, LUTImage
Dim LT
Sub LoadLUT
  bLutActive = False
    LT = LoadValue(cGameName, "LUTImage")
    If(LT <> "") Then LUTImage = LT Else LUTImage = 0
  UpdateLUT
End Sub

Sub SaveLUT
    SaveValue cGameName, "LUTImage", LUTImage
End Sub

Sub NextLUT: LUTImage = (LUTImage +1 ) MOD 17: UpdateLUT: SaveLUT: End Sub

Sub UpdateLUT
Select Case LutImage
Case 0: Table1.ColorGradeImage = "LUT0"
Case 1: Table1.ColorGradeImage = "LUT1"
Case 2: Table1.ColorGradeImage = "LUT2"
Case 3: Table1.ColorGradeImage = "LUT3"
Case 4: Table1.ColorGradeImage = "LUT4"
Case 5: Table1.ColorGradeImage = "LUT5"
Case 6: Table1.ColorGradeImage = "LUT6"
Case 7: Table1.ColorGradeImage = "LUT7"
Case 8: Table1.ColorGradeImage = "LUT8"
Case 9: Table1.ColorGradeImage = "LUT9"
Case 10: Table1.ColorGradeImage = "LUT10"
Case 11: Table1.ColorGradeImage = "LUT11"
Case 12: Table1.ColorGradeImage = "LUT12"
Case 13: Table1.ColorGradeImage = "LUT13"
Case 14: Table1.ColorGradeImage = "LUT14"
Case 15: Table1.ColorGradeImage = "LUT15"
Case 15: Table1.ColorGradeImage = "LUT16"
End Select
End Sub


Const ReflipAngle = 20 ' for Fleep
Dim BIPL 'scampa123 for Fleep

dim ScoreChecker
dim CheckAllScores
dim sortscores(4)
dim sortplayers(4)
Dim TextStr,TextStr2,mHole,mHole2, tKickerTCount
Dim i
Dim obj
Dim bgpos
Dim kgpos
Dim dooralreadyopen
Dim kgdooralreadyopen
Dim TargetSpecialLit
Dim Points210counter
Dim Points500counter
Dim Points1000counter
Dim Points2000counter
Dim BallsPerGame
Dim InProgress
Dim BallInPlay
Dim CreditsPerCoin
Dim Score100K(4)
Dim Score(4)
Dim ScoreDisplay(4)
Dim HighScorePaid(4)
Dim HighScore
Dim HighScoreReward
Dim BonusMultiplier
Dim Credits
Dim Match
Dim Replay1
Dim Replay2
Dim Replay3
Dim Replay1Paid(4)
Dim Replay2Paid(4)
Dim Replay3Paid(4)
Dim TableTilted
Dim TiltCount

Dim OperatorMenu

Dim SpecialIsLit

Dim BonusBooster
Dim BonusBoosterCounter
Dim BonusCounter
Dim HoleCounter

Dim Ones
Dim Tens
Dim Hundreds
Dim Thousands

Dim Player
Dim Players

Dim rst
Dim bonuscountdown
Dim TempMultiCounter
dim TempPlayerup
dim RotatorTemp

Dim bump1
Dim bump2
Dim bump3

Dim LastChime10
Dim LastChime100
Dim LastChime1000

Dim Score10
Dim Score100

Dim targettempscore
Dim SpecialLightCounter
Dim SpecialLightFlag
Dim HorseshoeCounter
Dim DropTargetCounter

Dim LeftTargetCounter
Dim RightTargetCounter

Dim MotorRunning
Dim Replay1Table(15)
Dim Replay2Table(15)
Dim Replay3Table(15)
Dim ReplayTableSet
Dim ReplayLevel
Dim ReplayTableMax
Dim BonusSpecialThreshold

dim bonustempcounter

Dim LStep, LStep2, RStep, xx

Dim ReelCounter
Dim BallCounter
Dim BallReelAddStart(10)
Dim BallReelAddFrames(10)
Dim BallReelDropStart(10)
Dim BallReelDropFrames(10)

Dim X
Dim Y
Dim Z
Dim AddABall
Dim AddABallFrames
Dim DropABall
Dim DropABallFrames
Dim CurrentFrame
Dim BonusMotorRunning
Dim QueuedBonusAdds
Dim QueuedBonusDrops

Dim TempLightTracker



Sub Table1_Init()
  If Table1.ShowDT = false then
    For each obj in DesktopCrap
      obj.visible=False
    next
  End If
  OperatorMenuBackdrop.image = "PostitBL"
  For XOpt = 1 to MaxOption
    Eval("OperatorOption"&XOpt).image = "PostitBL"
  next

  For XOpt = 1 to 256
    Eval("Option"&XOpt).image = "PostItBL"
  next

  LoadEM
  LoadLMEMConfig2

  SetBackglass
  center_objects_em
  HideOptions
  SetupReplayTables
  PlasticsOff
  BumpersOff
  OperatorMenu=0
  HighScore=0
  MotorRunning=0
  HighScoreReward=1
  Credits=0
  BallsPerGame=5
  ReplayLevel=1
  BonusSpecialThreshold=1
  SpecialLightFlag=2
  loadhs
  if HighScore=0 then HighScore=50000


  TableTilted=false

  Match=int(Rnd*10)*10
  MatchReel.SetValue((Match/10)+1)

  CanPlayReel.SetValue(0)
  GameOverReel.SetValue(1)
  TiltReel.SetValue(1)


'VRBG Trough lights
if VRRoom = 1 then
VRBGBulb1.visible = false
VRBGBulb2.visible = false
VRBGBulb3.visible = false
VRBGBulb4.visible = false
VRBGBulb5.visible = false
VRBGBulb6.visible = false
VRBGLight2.visible = false
VRBGLight1.visible = false

end if


  For each obj in PlayerHuds
    obj.SetValue(0)
  next
  For each obj in PlayerScores
    obj.ResetToZero
  next

  For each obj in PlayerScoresOn
    obj.ResetToZero
  next

  EMReel6.ResetToZero

  for each obj in bottgate
    obj.isdropped=true
    next
  primgate.RotY=90
  for each obj in Bonus
    obj.state=0
  next
  for each obj in HorseshoeLights
    obj.state=0
  next
  for each obj in RolloverLights1
    obj.state=0
  next
  for each obj in RolloverLights2
    obj.state=0
  next


  'turn off shadows
  Shadows.Visible=0

  Replay1=Replay1Table(ReplayLevel)
  Replay2=Replay2Table(ReplayLevel)
  Replay3=Replay3Table(ReplayLevel)

  BonusCounter=0
  HoleCounter=0
    bgpos=6
  kgpos=0
    bottgate(bgpos).isdropped=false


  dooralreadyopen=0
  kgdooralreadyopen=0

  InstructCard.image="IC"+FormatNumber(BallsPerGame,0)

  RefreshReplayCard


  Bumper1Light.state=0
  Bumper2Light.state=0

  BallCounter=0
  ReelCounter=0
  BuildBallReelTables
  QueuedBonusAdds=0
  QueuedBonusDrops=0
  BonusMotorRunning=0
  BonusMotorAdder.enabled=1
  BonusMotorDropper.enabled=1

  TargetSpecialLit = 0
  Points210counter=0
  Points500counter=0
  Points1000counter=0
  Points2000counter=0
  TempLightTracker=0

  BonusBooster=3
  BonusBoosterCounter=0
  Players=0
  RotatorTemp=1
  InProgress=false


  ScoreText.text=HighScore

  CurrentFrame=101

    If VRRoom = 1 then
  FlasherMatch
  for each Object in ColFlPlayer : object.visible = 1 : next
  for each Object in ColFlGameOver : object.visible = 1 : next
  for each Object in ColFlTilt : object.visible = 1 : next
    end if

  If B2SOn Then

    Controller.B2SSetData 101,1
    Controller.B2SSetMatch Match
    Controller.B2SSetScoreRolloverPlayer1 0
    Controller.B2SSetScoreRolloverPlayer2 0
    Controller.B2SSetScoreRolloverPlayer3 0
    Controller.B2SSetScoreRolloverPlayer4 0

    'Controller.B2SSetScorePlayer6 HighScore
    Controller.B2SSetTilt 0
    if Credits=0 then
      Controller.B2SSetData 29,10
    else
      Controller.B2SSetData 29,Credits
    end if
    Controller.B2SSetGameOver 1
  End If

  for i=1 to 4
    player=i
    If B2SOn Then
      Controller.B2SSetScorePlayer player, 0
    End If
  next
  bump1=1
  bump2=1
  bump3=1
  InitPauser5.enabled=true
  If Credits > 0 Then DOF 314, DOFOn
  LoadLUT
  VRSelectRoom
  If VRRoom = 1 Then LoadVRCarpet: End If
End Sub

Sub Table1_exit()
  savehs
  SaveLMEMConfig
  SaveLMEMConfig2
  If B2SOn Then Controller.Stop
end sub


Sub Table1_KeyDown(ByVal keycode)
  'Change LUT scampa123 added
  If keycode = LeftMagnaSave Then bLutActive = True: NextLUT: End If

  'Change VR Carpet
  If keycode = RightMagnaSave and VRRoom = 1 Then bVRCarpetActive = True: NextVRCarpet: End If

  ' GNMOD
  if EnteringInitials then
    CollectInitials(keycode)
    exit sub
  end if


  if EnteringOptions then
    CollectOptions(keycode)
    exit sub
  end if



  If keycode = PlungerKey Then
    Plunger.PullBack
    PlungerPulled = 1
    Plunger.Pullback:SoundPlungerPull()

if VRRoom = 1 then
TimerPlunger.Enabled = True
TimerPlunger2.Enabled = False
end If

End If


  if keycode = LeftFlipperKey then
    If VRRoom = 1 Then VRFlipperLeft.x = 2099.689 End If 'animated VR Button
  end if

  if keycode = LeftFlipperKey and InProgress = false then
    OperatorMenuTimer.Enabled = true
  end if
  ' END GNMOD

  If keycode = LeftFlipperKey and InProgress=true and TableTilted=false Then
    FlipperActivate LeftFlipper, LFPress
    LF.fire ' nFozzy......LeftFlipper.RotateToEnd


    If LeftFlipper.currentangle < LeftFlipper.endangle + ReflipAngle Then
                        RandomSoundReflipUpLeft LeftFlipper
                Else
                        SoundFlipperUpAttackLeft LeftFlipper
                        RandomSoundFlipperUpLeft LeftFlipper
                End If

    PlaySound "buzzL",-1
    'PlaySound SoundFXDOF("FlipperUp",301,DOFOn,DOFFlippers)
    DOF 301,DOFOn
  End If


  if keycode = RightFlipperKey Then
    If VRRoom = 1 Then VRFlipperRight.x = 2109.164 End If'animated VR Button
  end if

  If keycode = RightFlipperKey  and InProgress=true and TableTilted=false Then
    FlipperActivate RightFlipper, RFPress
    RF.fire ' nFozzy.......RightFlipper.RotateToEnd

    If RightFlipper.currentangle > RightFlipper.endangle - ReflipAngle Then
                        RandomSoundReflipUpRight RightFlipper
                Else
                        SoundFlipperUpAttackRight RightFlipper
                        RandomSoundFlipperUpRight RightFlipper
                End If
    DOF 302, DOFOn
    'PlaySound SoundFXDOF("FlipperUp",302,DOFOn,DOFFlippers)
    PlaySound "buzz",-1
  End If


  If keycode = LeftTiltKey Then
    Nudge 90, 5:SoundNudgeLeft()
    TiltIt
  End If

  If keycode = RightTiltKey Then
    Nudge 270, 5:SoundNudgeRight()
    TiltIt
  End If

  If keycode = CenterTiltKey Then
    Nudge 0, 3:SoundNudgeCenter()
    TiltIt
  End If

  If keycode = MechanicalTilt Then
    TiltCount=2
    TiltIt
  End If

  If keycode = AddCreditKey or keycode = 4 then
    If B2SOn Then
      'Controller.B2SSetScorePlayer6 HighScore

    End If

    Select Case Int(rnd*3)
              Case 0: PlaySound ("Coin_In_1"), 0, CoinSoundLevel, 0, 0.25
              Case 1: PlaySound ("Coin_In_2"), 0, CoinSoundLevel, 0, 0.25
              Case 2: PlaySound ("Coin_In_3"), 0, CoinSoundLevel, 0, 0.25
        End Select
    'AddSpecial2
    If credits < 10 then AddSpecial2

  end if

   if keycode = 5 then
    Select Case Int(rnd*3)
              Case 0: PlaySound ("Coin_In_1"), 0, CoinSoundLevel, 0, 0.25
              Case 1: PlaySound ("Coin_In_2"), 0, CoinSoundLevel, 0, 0.25
              Case 2: PlaySound ("Coin_In_3"), 0, CoinSoundLevel, 0, 0.25
        End Select
    'AddSpecial2
        If credits < 10 then AddSpecial2
    keycode= StartGameKey
  end if




   if keycode = StartGameKey and Credits>0 and InProgress=true and Players>0 and Players<4 and BallInPlay<2 then
    soundStartButton() 'SCAMPA123
    Credits=Credits-1
    If Credits <1 Then DOF 314, DOFOff
    CreditsReel.SetValue(Credits)
    EMReel7.SetValue(Credits)

' This is where we will subtract a credit from the VR BG

if VRRoom = 1 then
Creditwheel.objroty = Creditwheel.objroty - 16.5
If Credits = 4 then Creditwheel.objroty = Creditwheel.objroty - 16.5  ' one extra turn
end If

    Players=Players+1
    CanPlayReel.SetValue(Players)
    playsound "click"

    If VRRoom = 1 then
    FlasherPlayers
    End if

    If B2SOn Then
      Controller.B2SSetCanPlay Players
      If Players=2 Then
        Controller.B2SSetScoreRolloverPlayer2 0
      End If
      If Players=3 Then
        Controller.B2SSetScoreRolloverPlayer3 0
      End If
      If Players=4 Then
        Controller.B2SSetScoreRolloverPlayer4 0
      End If
      if Credits=0 then
        Controller.B2SSetData 29, 10
      else
        Controller.B2SSetData 29, Credits
      end if
    End If
    end if

  if keycode= StartGameKey and Credits>0 and InProgress=false and Players=0 and EnteringOptions = 0 then
    soundStartButton() 'SCAMPA123
    if VRRoom = 1 then VR_Instructions.visible=0  End if 'hide instruction screen whne game starts scampa123
'GNMOD
    OperatorMenuTimer.Enabled = false
'END GNMOD
    Credits=Credits-1
    If Credits > 0 Then DOF 314, DOFOff
    CreditsReel.SetValue(Credits)
    EMReel7.SetValue(Credits)


' subtract credit here too
if VRRoom = 1 then
Creditwheel.objroty = Creditwheel.objroty - 16.5
If Credits = 4 then Creditwheel.objroty = Creditwheel.objroty - 16.5  ' one extra turn
end If


    Players=1
    CanPlayReel.SetValue(Players)
    MatchReel.SetValue(0)
    Player=1
    GameOverReel.SetValue(0)

'VRBG Trough lights turn on at game start
if VRRoom = 1 then
VRBGBulb1.visible = true
VRBGBulb2.visible = true
VRBGBulb3.visible = true
VRBGBulb4.visible = true
VRBGBulb5.visible = true
VRBGBulb6.visible = true
VRBGLight1.visible = true
VRBGLight2.visible = true
end if

    playsound "StartUpSequence"
    TempPlayerUp=Player
'   PlayerUpRotator.enabled=true
    rst=0
    BallInPlay=1
    InProgress=true
    resettimer.enabled=true
    BonusMultiplier=1

    for each Object in ColFlTilt : object.visible = 0 : next
    for each Object in ColFlGameOver : object.visible = 0 : next
    for each Object in ColFlMatch : object.visible = 0 : next

        If VRRoom = 1 then
    FlasherPlayers
    FlasherBalls
    FlasherCurrentPlayer
    End if

    If B2SOn Then
      Controller.B2SSetTilt 0
      Controller.B2SSetGameOver 0
      Controller.B2SSetMatch 0
      if Credits=0 then
        Controller.B2SSetData 29, 10
      else
        Controller.B2SSetData 29, Credits
      end if
      'Controller.B2SSetScorePlayer6 HighScore
      Controller.B2SSetCanPlay 1
      Controller.B2SSetPlayerUp 1
      Controller.B2SSetData 30,1
'     Controller.B2SSetBallInPlay BallInPlay
      Controller.B2SSetScoreRolloverPlayer1 0
    End If
    If Table1.ShowDT = True then
      For each obj in PlayerScores
'       obj.ResetToZero
        obj.Visible=true
      next
      For each obj in PlayerScoresOn
'       obj.ResetToZero
        obj.Visible=false
      next

      For each obj in PlayerHuds
        obj.SetValue(0)
      next
      For each obj in PlayerHUDScores
        obj.state=0
      next
      PlayerHuds(Player-1).SetValue(1)
      PlayerHUDScores(Player-1).state=1

      PlayerScores(Player-1).Visible=0
      PlayerScoresOn(Player-1).Visible=1
    end If

  end if



End Sub

Sub Table1_KeyUp(ByVal keycode)
  'Change LUT scampa123 added
  If keycode = LeftMagnaSave Then bLutActive = False
  'If KeyCode = LeftMagnaSave And RightMagnaSave Then ScreenTimer.Enabled = True

  'Change VR Carpet
  If keycode = RightMagnaSave Then bVRCarpetActive = False

  ' GNMOD
  if EnteringInitials then
    exit sub
  end if

  If KeyCode = PlungerKey Then
    Plunger.Fire
    If BIPL = 1 Then
        SoundPlungerReleaseBall()                        'Plunger release sound when there is a ball in shooter lane
    Else
        SoundPlungerReleaseNoBall()                        'Plunger release sound when there is no ball in shooter lane
    End If


if VRRoom = 1 then
  TimerPlunger.Enabled = False
  TimerPlunger2.Enabled = True
  VR_Primary_plunger.Y = 2081
end If
End If

  if keycode = LeftFlipperKey then
    OperatorMenuTimer.Enabled = false
    If VRRoom = 1 Then VRFlipperLeft.x = 2093.689 End If 'animated VR Button
  end if


  If keycode = LeftFlipperKey and InProgress=true and TableTilted=false Then
    FlipperDeActivate LeftFlipper, LFPress
    '
                LeftFlipper.RotateToStart
                If LeftFlipper.currentangle < LeftFlipper.startAngle - 5 Then
                        RandomSoundFlipperDownLeft LeftFlipper
                End If
                FlipperLeftHitParm = FlipperUpSoundLevel

    'PlaySound SoundFXDOF("FlipperDown",301,DOFOff,DOFFlippers)
    DOF 301, DOFOFF
    StopSound "buzzL"
  End If

  If keycode = RightFlipperKey Then
    If VRRoom = 1 Then VRFlipperRight.x =2115.164 End If 'animated VR Button
  End If

  If keycode = RightFlipperKey and InProgress=true and TableTilted=false Then
    FlipperDeActivate RightFlipper, RFPress

    RightFlipper.RotateToStart
                If RightFlipper.currentangle > RightFlipper.startAngle + 5 Then
                        RandomSoundFlipperDownRight RightFlipper
                End If
                FlipperRightHitParm = FlipperUpSoundLevel

        StopSound "buzz"
    'PlaySound SoundFXDOF("FlipperDown",302,DOFOff,DOFFlippers)
    DOF 302, DOFOFF
  End If

End Sub


Sub Drain_Hit()
  DOF 311, 2

NewBallShadowUpdate.enabled = false  ' testing this for new ballshadow code.We need to turn it off at drain.

  Drain.DestroyBall
  RandomSoundDrain Drain
  Pause4Bonustimer.enabled=1

End Sub

Sub Pause4Bonustimer_timer
  if BonusMotorRunning=0 then
    Pause4Bonustimer.enabled=0
    AddBonus
  end if
End Sub

Sub Trigger0_Unhit()
  DOF 313, 2
End Sub


'***********************
'     Flipper Logos
'***********************

Sub UpdateFlipperLogos_Timer
  LFlip.RotZ = LeftFlipper.CurrentAngle
  RFlip.RotZ = RightFlipper.CurrentAngle
  PGate.Rotz = (Gate.CurrentAngle*.75) + 25
  FlipperLSh.RotZ = LeftFlipper.currentangle
  FlipperRSh.RotZ = RightFlipper.currentangle
End Sub


'***********************
' slingshots
'

Sub RightSlingShot_Slingshot
  RandomSoundSlingshotLeft SLING1
    DOF 305,DOFPulse
  DOF 331, DOFPulse
    RSling0.Visible = 0
    RSling1.Visible = 1
    sling1.TransZ = -20
    RStep = 0
    RightSlingShot.TimerEnabled = 1
  AddScore 10
End Sub

Sub RightSlingShot_Timer
    Select Case RStep
        Case 3:RSLing1.Visible = 0:RSLing2.Visible = 1:sling1.TransZ = -10
        Case 4:RSLing2.Visible = 0:RSling0.Visible = 1:sling1.TransZ = 0:RightSlingShot.TimerEnabled = 0
    End Select
    RStep = RStep + 1
End Sub

Sub LeftSlingShot_Slingshot
  RandomSoundSlingshotLeft SLING2
    DOF 303,DOFPulse
  DOF 330, DOFPulse
    LSling0.Visible = 0
    LSling1.Visible = 1
    sling2.TransZ = -20
    LStep = 0
    LeftSlingShot.TimerEnabled = 1
  AddScore 10
End Sub

Sub LeftSlingShot_Timer
    Select Case LStep
        Case 3:LSLing1.Visible = 0:LSLing2.Visible = 1:sling2.TransZ = -10
        Case 4:LSLing2.Visible = 0:LSLing0.Visible = 1:sling2.TransZ = 0:LeftSlingShot.TimerEnabled = 0
    End Select
    LStep = LStep + 1
End Sub


'***************************
'  Left Side Switches
'***************************

Sub LeftBottomSW_Hit
  debug.print "left bottom sw"
  DOF 304,DOFPulse
  AddScore 10
End Sub

Sub LeftMiddleSW_Hit
  debug.print "left middle sw"
  DOF 306,DOFPulse
  AddScore 10
End Sub

Sub LeftTopSW_Hit
  debug.print "left top sw"
  DOF 307,DOFPulse
  AddScore 10
End Sub
'Sub LeftSlingShot2_Slingshot
' RandomSoundSlingshotLeft SLING3
'    DOF 304,DOFPulse
'    LSling3.Visible = 0
'    LSling4.Visible = 1
'    sling3.TransZ = -20
'    LStep2 = 0
'    LeftSlingShot2.TimerEnabled = 1
' AddScore 10
'End Sub
'
'Sub LeftSlingShot2_Timer
'    Select Case LStep2
'        Case 3:LSLing4.Visible = 0:LSLing5.Visible = 1:SLING3.TransZ = -10
'        Case 4:LSLing5.Visible = 0:LSLing3.Visible = 1:SLING3.TransZ = 0:LeftSlingShot2.TimerEnabled = 0
'    End Select
'    LStep2 = LStep2 + 1
'End Sub



'***********************************
' Bumpers
'***********************************


Sub Bumper1_Hit
  If TableTilted=false then
    RandomSoundBumperTop Bumper1
    DOF 307, DOFPulse
    DOF 332, DOFPulse
    bump1 = 1
    AddScore(100)
    end if
End Sub

Sub Bumper2_Hit
  If TableTilted=false then
    RandomSoundBumperBottom Bumper2
    DOF 308,DOFPulse
    DOF 333, DOFPulse
    bump2 = 1
    AddScore(100)
    end if
End Sub


'****************************
' Stationary targets
'****************************

Sub TargetRight1_Hit()
  If TableTilted=false then
    if RightUpperTargetLight.state=1 then
      IncreaseBonus
    else
      SetMotor(500)
    end if
    if DropTargetLight.state=1 then
      AddSpecial
      DropTargetLight.state=0
    end if
  end if
  'PlaySound SoundFXDOF("target", 326, DOFPulse, DOFTargets)

  DOF 326, DOFPulse
end Sub


Sub TargetRight2_Hit()
  If TableTilted=false then

    if RightCenterTargetLight.state=1 then
      IncreaseBonus
    else
      SetMotor(500)
    end if
  end if
  'PlaySound SoundFXDOF("target", 327, DOFPulse, DOFTargets)

  DOF 327, DOFPulse
end Sub

Sub TargetRight3_Hit()
  If TableTilted=false then

    if RightLowerTargetLight.state=1 then
      IncreaseBonus
    else
      SetMotor(500)
    end if
  end if
  'PlaySound SoundFXDOF("target", 327, DOFPulse, DOFTargets)

  DOF 327, DOFPulse
end Sub

Sub TargetLeft1_Hit()
  If TableTilted=false then

    if LeftLowerTargetLight.state=1 then
      IncreaseBonus
    else
      SetMotor(500)
    end if
  end if
  'PlaySound SoundFXDOF("target", 306, DOFPulse, DOFTargets)

  DOF 306, DOFPulse
end Sub

'******************************************
' Spinner


Sub Spinner1_Spin()
  DOF 325, 2
  If TableTilted=false then
    If BallsPerGame=3 then
      AddScore(100)
    Else
      AddScore(10)
    end if
    IncreaseHorseshoe
    ToggleAlternatingRelay
  end if

end Sub

'****************************************************
' Kickers

sub Kicker1_Hit()

  KickerDelayer.enabled=1


end sub

Sub KickerDelayer_Timer()
  If MotorRunning<>0 then
    exit sub
  end if
  If BonusMotorRunning<>0 then
    exit sub
  end if
  KickerDelayer.enabled=0
  Kicker1.timerinterval=(900)

  QueuedBonusDrops=BallCounter
  KickerHold1.enabled=1
end sub

Sub KickerHold1_timer
  if QueuedBonusDrops<>0 then
    exit sub
  end if
  KickerHold1.enabled=false
  Kicker1.timerenabled=true
  tKickerTCount=0
end sub

Sub Kicker1_Timer()

  tKickerTCount=tKickerTCount+1
  select case tKickerTCount
  case 1:
  ' mHole.MagnetOn=0
    If TableTilted=false then
      SetMotor(1000)
    end if
    Pkickarm1.rotz=15
    DOF 310, DOFPulse
    SoundSaucerKick 1, Kicker1
    DOF 334, DOFPulse
    X= INT(RND()*5+1)
    Kicker1.kick 150+X,10
  case 2:
    Kicker1.timerenabled=0
    Pkickarm1.rotz=0
  end Select

end sub

Sub Kicker2_Hit()
  KickerDelayer2.enabled=1
end sub

Sub KickerDelayer2_Timer()
  If MotorRunning<>0 then
    exit sub
  end if
  If BonusMotorRunning<>0 then
    exit sub
  end if
  KickerDelayer2.enabled=0
  Kicker2.timerinterval=(500)

  X = INT(RND()*3)+1
  If TableTilted=false then
    AddBonusBalls(X)
  end if
  KickerHold2.enabled=1
  tKickerTCount=0
end sub

sub KickerHold2_timer
  if QueuedBonusAdds<>0 then
    exit sub
  end if
  KickerHold2.enabled=false
  Kicker2.timerenabled=true
end sub


Sub Kicker2_Timer()

  tKickerTCount=tKickerTCount+1
  select case tKickerTCount
  case 1:
  ' mHole.MagnetOn=0
    If TableTilted=false then
      SetMotor(1000)
    end if
    Pkickarm2.rotz=15
    SoundSaucerKick 1, Kicker2
    DOF 310, DOFPulse
    DOF 334, DOFPulse
    X= INT(RND()*5+1)
    Kicker2.kick 140+X,10
  case 2:
    Kicker2.timerenabled=0
    Pkickarm2.rotz=0
  end Select

end sub


'************************************
'  Rollover lanes
'************************************

Sub CenterTrigger1_Hit()
  If TableTilted=false then
    'SetMotor(500)
    DOF 317, 2
    If CenterTriggerLight1.state=1 then
      DropTargetLight.state=1
      SpecialIsLit=1
      CenterTriggerLight1.state=0
    end if

  end if

end Sub

Sub TriggerTop1_Hit()
  If TableTilted=false then

    DOF 315, 2
    If UpperLight1.state=1 then
'     SetMotor(500)
      IncreaseBonus
    else
      SetMotor(500)
    end if
  end if
End Sub

Sub TriggerTop2_Hit()
  If TableTilted=false then

    DOF 316, 2
    If UpperLight2.state=1 then
'     SetMotor(500)
      IncreaseBonus
    else
      SetMotor(500)
    end if
  end if
End Sub

Sub TriggerLeftRollover_Hit()
  If TableTilted=false then

    DOF 312, DOFPulse
    If LeftRolloverLight.state=1 then
'     SetMotor(500)
      IncreaseBonus
    else
      SetMotor(500)
    end if
    if dooralreadyopen=0 then
      openg.enabled=true
    end if
  end if
End Sub

Sub TriggerRight1_Hit()
  If TableTilted=false then

    DOF 318, 2
    If RightRolloverLight1.state=1 then
'     SetMotor(500)
      IncreaseBonus
    else
      SetMotor(500)
    end if
  end if
End Sub

Sub TriggerRight2_Hit()
  If TableTilted=false then

    DOF 319, 2
    If RightRolloverLight2.state=1 then
'     SetMotor(500)
      IncreaseBonus
    else
      SetMotor(500)
    end if
  end if
End Sub

Sub TriggerRight3_Hit()
  If TableTilted=false then

    DOF 318, 2
    If RightRolloverLight3.state=1 then
'     SetMotor(500)
      IncreaseBonus
    else
      SetMotor(500)
    end if
  end if
End Sub

Sub TriggerRight4_Hit()
  If TableTilted=false then

    DOF 319, 2
    If RightRolloverLight4.state=1 then
'     SetMotor(500)
      IncreaseBonus
    else
      SetMotor(500)
    end if
  end if
End Sub

Sub TriggerRight5_Hit()
  If TableTilted=false then

    DOF 318, 2
    If RightRolloverLight5.state=1 then
'     SetMotor(500)
      IncreaseBonus
    else
      SetMotor(500)
    end if
  end if
End Sub

Sub TriggerRight6_Hit()
  If TableTilted=false then

    DOF 319, 2
    If RightRolloverLight6.state=1 then
'     SetMotor(500)
      IncreaseBonus
    else
      SetMotor(500)
    end if
  end if
End Sub

Sub TriggerRightInlane_Hit()
  If TableTilted=false then

    DOF 336, 2
    If RightInlaneLight.state=1 then
'     SetMotor(500)
      IncreaseBonus
    else
      SetMotor(500)
    end if
  end if
End Sub

Sub TriggerRight8_Hit()
  If TableTilted=false then

    DOF 335, 2
'   SetMotor(500)
    IncreaseBonus

  end if
End Sub

Sub TriggerLeftOutlane_Hit()
  If TableTilted=false then

    DOF 320, 2
'   SetMotor(500)
    IncreaseBonus

  end if
End Sub

Sub TriggerLeftInlane_Hit()
  If TableTilted=false then

    DOF 321, 1
    If LeftInlaneLight.state=1 then
'     SetMotor(500)
      IncreaseBonus
    else
      SetMotor(500)
    end if
  end if
End Sub


Sub ResetDrops

  DropTargetCounter=0
end Sub

Sub CheckAllDrops

  If DropTargetCounter=10 then
    SetMotor(5000)
    if DropTargetLight.state=1 then
      AddSpecial
    end if
    ResetDropsTimer.enabled=1

  end if
end Sub

sub ResetDropsTimer_timer
    ResetDropsTimer.enabled=0
    ResetDrops
end sub

'**************************************
'**************************************


Sub CloseGateTrigger_Hit()
  BIPL = 1 'scampa123
  if dooralreadyopen=1 then
    closeg.enabled=true
  end if
End Sub

Sub CloseGateTrigger_UnHit()
  BIPL = 0
End Sub



Sub AddSpecial()
  DOF 309,DOFPulse
  KnockerSolenoid
  Credits=Credits+1

' This is where we will turn the VR BG Credits reel.on Special bonus
if VRRoom = 1 then
If Credits < 10 then Creditwheel.objroty = Creditwheel.objroty + 16.5
If Credits = 5 then Creditwheel.objroty = Creditwheel.objroty + 16.5  ' one extra turn
If Credits>9 then Credits=9: Creditwheel.objroty = -6.5   ' Seeing if this forces the VR credit reel to 9..
end if

  If Credits > 0 Then DOF 314, DOFOn

  If B2SOn Then
    Controller.B2SSetData 29, Credits
  End If
  CreditsReel.SetValue(Credits)
  EMReel7.SetValue(Credits)
End Sub

Sub AddSpecial2()  ' this is the one that is called whn you press the credit button..
  'PlaySound"click"
  Credits=Credits+1

' This is where we will turn the VR BG Credits reel.on Special bonus2 - This also runs when you press the credit button.
if VRRoom = 1 then
If Credits < 10 then Creditwheel.objroty = Creditwheel.objroty + 16.5
If Credits = 5 then Creditwheel.objroty = Creditwheel.objroty + 16.5  ' one extra turn
if Credits>9 then Creditwheel.objroty = -6.5   ' Seeing if this forces the VR credit reel to 9.. works
end If


  If Credits > 0 Then DOF 314, DOFOn
  if Credits>9 then Credits=9
  If B2SOn Then
    Controller.B2SSetData 29, Credits
  End If
  CreditsReel.SetValue(Credits)
  EMReel7.SetValue(Credits)
End Sub

Sub AddBonus()
  QueuedBonusDrops=BallCounter
  ScoreBonus.enabled=true
End Sub

Sub ToggleAlternatingRelay
  if RightRolloverLight2.state=0 then
    for each obj in RolloverLights2
      obj.state=1
    next
    for each obj in RolloverLights1
      obj.state=0
    next
  else
    for each obj in RolloverLights2
      obj.state=0
    next
    for each obj in RolloverLights1
      obj.state=1
    next
  end if

  If SpecialLightFlag=2 or SpecialLightFlag=3 then
    ToggleSpecial
  end if
end sub

sub ToggleSpecial

  SpecialLightCounter=SpecialLightCounter+1
  If SpecialLightCounter>3 then
    SpecialLightCounter=0
  end if
  CenterTriggerLight1.state=0
  DropTargetLight.state=0
  SpecialIsLit=0
  Select Case SpecialLightCounter
    Case 0:
      TextBox1.text = "Case 0"
      if SpecialLightFlag=3 then
        CenterTriggerLight1.state=1


      end if
    Case 1:
      TextBox1.text = "Case 1"
      if SpecialLightFlag=2 or SpecialLightFlag=3 then
        CenterTriggerLight1.state=1

      end if
    Case 2:
      TextBox1.text = "Case 2"
      if SpecialLightFlag=3 then
        CenterTriggerLight1.state=1

      end if
    Case 3:
      TextBox1.text = "Case 3"
      CenterTriggerLight1.state=1



  End Select

end sub

sub IncreaseHorseshoe

  HorseshoeCounter=HorseshoeCounter+1
  for each obj in HorseshoeLights
    obj.state=0
  next
  if HorseshoeCounter>4 then
    HorseshoeCounter=0
'   SetMotor(500)
    IncreaseBonus
    HorseshoeLights(HorseshoeCounter).state=1
  else
    HorseshoeLights(HorseshoeCounter).state=1
  end if
end sub

Sub ResetBallDrops


  ResetDrops
  for each obj in Bonus
    obj.state=0
  next

  BonusCounter=0
  HoleCounter=0

End Sub


Sub LightsOut
  for each obj in Bonus
    obj.state=0
  next
  for each obj in HorseshoeLights
    obj.state=0
  next
  BonusCounter=0
  HoleCounter=0
  Bumper1Light.state=0
  Bumper2Light.state=0

  'turn off shadows
  Shadows.Visible=0

  if dooralreadyopen=1 then
    closeg.enabled=true
  end if

end sub

Sub ResetBalls()

  TempMultiCounter=BallsPerGame-BallInPlay

  ResetBallDrops
  BonusMultiplier=1
  DoubleBonus.state=0
  If (BallsPerGame=BallInPlay) then
    DoubleBonus.state=1
    If BallsPerGame=3 then
      BonusMultiplier=3
    else
      BonusMultiplier=2
    end if
  End If
  If (BallsPerGame=3) and (BallInPlay=2) then
    DoubleBonus.state=1
    BonusMultiplier=2
  end if
  TableTilted=false
  TiltReel.SetValue(0)
  for each Object in ColFlTilt : object.visible = 0 : next
  If B2Son then
    Controller.B2SSetTilt 0
  end if
  PlasticsOn
  BumpersOn
  LeftInlaneLight.state=0
  RightInlaneLight.state=0
  RightCenterTargetLight.state=0
  RightLowerTargetLight.state=0
  LeftRolloverLight.state=0
  DropTargetLight.state=0
  SpecialIsLit=0


  If BallsPerGame=3 then
    If BallInPlay>=1 then
      LeftInlaneLight.state=1
      RightInlaneLight.state=1
    end if
    if BallInPlay>=2 then
      RightCenterTargetLight.state=1
      RightLowerTargetLight.state=1
    end if
    if BallInPlay>=3 then
      LeftRolloverLight.state=1
    end if
  else
    If BallInPlay>=3 then
      LeftInlaneLight.state=1
      RightInlaneLight.state=1
    end if
    if BallInPlay>=4 then
      RightCenterTargetLight.state=1
      RightLowerTargetLight.state=1
    end if
    if BallInPlay>=5 then
      LeftRolloverLight.state=1
    end if
  end if


' CreateBallID BallRelease
  DOF 328, DOFPulse
    'set GameBall = Ballrelease.CreateSizedBall 25  ' does not work
    set GameBall = Ballrelease.CreateBall ' works
    GameBall.image = "Ball_HDR" 'add ball image to silver ball
    Ballrelease.Kick 40,7
  RandomSoundBallRelease BallRelease

    NewBallShadowUpdate.enabled = true  ' made a new shadow routine that just follows the single gameball.  The old one caused problems with the VR backglass balls.

  if dooralreadyopen=1 then
    closeg.enabled=true
  end if
  if kgdooralreadyopen=1 then
    closekg.enabled=true
  end if
  BallInPlayReel.SetValue(BallInPlay)
  InstructCard.image="IC"+FormatNumber(BallsPerGame,0)


  HorseshoeCounter=4
  HorseshoeLights(HorseshoeCounter).state=1
  ToggleAlternatingRelay

  'turn on shadows
  Shadows.Visible=1
End Sub

sub delaykgclose_timer
  delaykgclose.enabled=false
  closekg.enabled=true

end sub

 sub openg_timer
    bottgate(bgpos).isdropped=true
    bgpos=bgpos-1
    bottgate(bgpos).isdropped=false
  primgate.RotY=30+(bgpos*10)
     if bgpos=0 then
    playsound "postup"

    openg.enabled=false
    dooralreadyopen=1
  end if

 end sub

sub closeg_timer
    closeg.interval=10
    bottgate(bgpos).isdropped=true
    bgpos=bgpos+1
    bottgate(bgpos).isdropped=false
    primgate.RotY=30+(bgpos*10)
  if bgpos=6 then

    closeg.enabled=false
    dooralreadyopen=0
  end if
end sub


sub resettimer_timer
    rst=rst+1
  If rst=1 then
    PlayerUpRotator.enabled=true
  end if
  if rst=7 then
    PlayerUpRotator.enabled=true
  end if

  if rst>10 and rst<21 then
    ResetReelsToZero(1)
  end if
  if rst>20 and rst<31 then
    ResetReelsToZero(2)
  end if

    if rst=32 then
    playsound "StartBall1"
    end if
    if rst=35 then
    newgame
    resettimer.enabled=false
    end if
end sub

Sub ResetReelsToZero(reelzeroflag)
  dim d1(5)
  dim d2(5)
  dim scorestring1, scorestring2

  If reelzeroflag=1 then
    scorestring1=CStr(Score(1))
    scorestring2=CStr(Score(2))
    scorestring1=right("00000" & scorestring1,5)
    scorestring2=right("00000" & scorestring2,5)
    for i=0 to 4
      d1(i)=CInt(mid(scorestring1,i+1,1))
      d2(i)=CInt(mid(scorestring2,i+1,1))
    next
    for i=0 to 4
      if d1(i)>0 then
        d1(i)=d1(i)+1
        if d1(i)>9 then d1(i)=0
      end if
      if d2(i)>0 then
        d2(i)=d2(i)+1
        if d2(i)>9 then d2(i)=0
      end if

    next
    Score(1)=(d1(0)*10000) + (d1(1)*1000) + (d1(2)*100) + (d1(3)*10) + d1(4)
    Score(2)=(d2(0)*10000) + (d2(1)*1000) + (d2(2)*100) + (d2(3)*10) + d2(4)

  EMMODE = 1
  UpdateReels 0,0 ,score(0), 0, -1,-1,-1,-1,-1
  UpdateReels 1,1 ,score(1), 0, -1,-1,-1,-1,-1
  UpdateReels 2,2 ,score(2), 0, -1,-1,-1,-1,-1
  UpdateReels 3,3 ,score(3), 0, -1,-1,-1,-1,-1
  UpdateReels 4,4 ,score(4), 0, -1,-1,-1,-1,-1
  EMMODE = 0 ' restore EM mode

    If B2SOn Then
      Controller.B2SSetScorePlayer 1, Score(1)
      Controller.B2SSetScorePlayer 2, Score(2)
    End If
    PlayerScores(0).SetValue(Score(1))
    PlayerScoresOn(0).SetValue(Score(1))
    PlayerScores(1).SetValue(Score(2))
    PlayerScoresOn(1).SetValue(Score(2))

  end if
  If reelzeroflag=2 then
    scorestring1=CStr(Score(3))
    scorestring2=CStr(Score(4))
    scorestring1=right("00000" & scorestring1,5)
    scorestring2=right("00000" & scorestring2,5)
    for i=0 to 4
      d1(i)=CInt(mid(scorestring1,i+1,1))
      d2(i)=CInt(mid(scorestring2,i+1,1))
    next
    for i=0 to 4
      if d1(i)>0 then
        d1(i)=d1(i)+1
        if d1(i)>9 then d1(i)=0
      end if
      if d2(i)>0 then
        d2(i)=d2(i)+1
        if d2(i)>9 then d2(i)=0
      end if

    next
    Score(3)=(d1(0)*10000) + (d1(1)*1000) + (d1(2)*100) + (d1(3)*10) + d1(4)
    Score(4)=(d2(0)*10000) + (d2(1)*1000) + (d2(2)*100) + (d2(3)*10) + d2(4)
    If B2SOn Then
      Controller.B2SSetScorePlayer 3, Score(3)
      Controller.B2SSetScorePlayer 4, Score(4)
    End If
    PlayerScores(2).SetValue(Score(3))
    PlayerScoresOn(2).SetValue(Score(3))
    PlayerScores(3).SetValue(Score(4))
    PlayerScoresOn(3).SetValue(Score(4))

  end if

end sub



sub ResetHorseshoeLights_timer
  if HorseshoeCounter=4 then
    ResetHorseshoeLights.enabled=0
    NextBallDelay.enabled=true
    exit sub
  end if
  HorseshoeLights(HorseshoeCounter).state=0
  HorseshoeCounter=HorseShoeCounter+1
  If HorseshoeCounter>4 then
    HorseshoeCounter=0
  end if
  HorseshoeLights(HorseshoeCounter).state=1
end sub

sub ScoreBonus_timer

  if BallCounter<1 then
    ScoreBonus.enabled=false
    ScoreBonus.interval=600
    ResetHorseshoeLights.enabled=true
    Playsound "MotorRunning"
    exit sub
  end if
' if bonuscountdown<=0 then
'   ScoreBonus.enabled=false
'   ScoreBonus.interval=600
'   ResetHorseshoeLights.enabled=true
'   Playsound "MotorRunning"
'   exit sub
' end if
' if BonusMultiplier=1 then
'   Playsound "BonusScore1x"
' else
'   Playsound "BonusScore2x"
' end if
' if BonusCountdown>10 then
'   bonustempcounter=1
'   BonusScorer.enabled=true
'   SetMotor2(1000*BonusMultiplier)
'   Bonus(bonuscountdown-10).state=0
' else
'   SetMotor2(1000*BonusMultiplier)
'   bonustempcounter=1
'   BonusScorer.enabled=true
'   Bonus(bonuscountdown).state=0
' end if
' If BonusMultiplier=2 then
'   Bonus(0).state=1
' end if
' If BonusMultiplier=1 then
'   ScoreBonus.interval=400
' Else
'   ScoreBonus.interval=400
' end if
' bonuscountdown=bonuscountdown-1
'
' if bonuscountdown>10 then
'   Bonus(bonuscountdown-10).state=1
'   exit sub
' end if
' if bonuscountdown>0 then
'   Bonus(bonuscountdown).state=1
' end if
'
end sub

sub BonusScorer_timer

  select case bonustempcounter
    case 0:
      BonusScorer.enabled=false
      if BonusMultiplier=2 then
        AddScore(1000)
      end if
    case 1:
      bonustempcounter=bonustempcounter-1
      AddScore(1000)

  end select
end sub

sub NextBallDelay_timer()
  NextBallDelay.enabled=false
  nextball

end sub

sub newgame
  InProgress=true
  queuedscore=0
  for i = 1 to 4
    Score(i)=0
    Score100K(1)=0
    HighScorePaid(i)=false
    Replay1Paid(i)=false
    Replay2Paid(i)=false
    Replay3Paid(i)=false
  next

    for each Object in ColFlTilt : object.visible = 0 : next
    for each Object in ColFlGameOver : object.visible = 0 : next
    for each Object in ColFlMatch : object.visible = 0 : next

        If VRRoom = 1 then
    FlasherBalls
    End If


    UpdateReels  0 ,0 ,Score(0), 0, 0,0,0,0,0
    UpdateReels  1 ,1 ,Score(0), 0, 0,0,0,0,0
    UpdateReels  2 ,2 ,Score(0), 0, 0,0,0,0,0
    UpdateReels  3 ,3 ,Score(0), 0, 0,0,0,0,0

  If B2SOn Then
    Controller.B2SSetTilt 0
    Controller.B2SSetGameOver 0
    Controller.B2SSetMatch 0
'   Controller.B2SSetScorePlayer1 0
'   Controller.B2SSetScorePlayer2 0
'   Controller.B2SSetScorePlayer3 0
'   Controller.B2SSetScorePlayer4 0
    Controller.B2SSetBallInPlay BallInPlay
  End if
  for each obj in RolloverLights1
    obj.state=0
  next
  for each obj in RolloverLights2
    obj.state=0
  next
  SpecialLightCounter=3
  DropTargetLight.state=0

  UpperLight1.state=1
  UpperLight2.state=1
  LeftLowerTargetLight.state=1
  RightUpperTargetLight.state=1
  CenterTriggerLight1.state=1
  SpecialIsLit=0
  ToggleAlternatingRelay
  ResetBalls
End sub

sub nextball
  for each Object in ColFlTilt : object.visible = 0 : next
  If B2SOn Then
    Controller.B2SSetTilt 0
  End If
  Player=Player+1
  If Player>Players Then
    BallInPlay=BallInPlay+1
    If BallInPlay>BallsPerGame then
      PlaySound("MotorLeer")
      InProgress=false

      for each Object in ColFlGameOver : object.visible = 1 : next
      for each Object in ColFlPlayers : object.visible = 0 : next
      for each Object in ColFlBalls : object.visible = 0 : Next
      for each Object in ColFlPlayer : object.visible = 1 : Next

    DoubleBonus.state=0 'scampa123
    CenterTriggerLight1.state=0 'scampa123
    if VRRoom = 1 then VR_Instructions.visible=1  End if 'show instruction screen whne game starts scampa123

      If B2SOn Then
        Controller.B2SSetGameOver 1
        Controller.B2SSetPlayerUp 0
        Controller.B2SSetBallInPlay 0
        Controller.B2SSetCanPlay 0
      End If
      For each obj in PlayerHuds
        obj.SetValue(0)
      next
      For each obj in PlayerHUDScores
        obj.state=0
      next
      If Table1.ShowDT = true then
        For each obj in PlayerScores
          obj.visible=1
        Next
        For each obj in PlayerScoresOn
          obj.visible=0
        Next
      end If
      InstructCard.image="IC"+FormatNumber(BallsPerGame,0)

      BallInPlayReel.SetValue(0)
      CanPlayReel.SetValue(0)
      GameOverReel.SetValue(1)


'VRBG Trough lights..game ends
if VRRoom = 1 then
VRBGBulb1.visible = false
VRBGBulb2.visible = false
VRBGBulb3.visible = false
VRBGBulb4.visible = false
VRBGBulb5.visible = false
VRBGBulb6.visible = false
VRBGLight1.visible = false
VRBGLight2.visible = false
end if


      LeftFlipper.RotateToStart
      RightFlipper.RotateToStart
      LightsOut
      BumpersOff
      PlasticsOff
      checkmatch
      CheckHighScore
      Players=0
      HighScoreTimer.interval=100
      HighScoreTimer.enabled=True
    Else
      Player=1

            If VRRoom = 1 then
      FlasherPlayers
      FlasherBalls
      FlasherCurrentPlayer
      End If

      If B2SOn Then
        Controller.B2SSetPlayerUp Player
        Controller.B2SSetBallInPlay BallInPlay

      End If
      PlaySound("RotateThruPlayers")
      TempPlayerUp=Player
      PlayerUpRotator.enabled=true
      PlayStartBall.enabled=true
      For each obj in PlayerHuds
        obj.SetValue(0)
      next
      For each obj in PlayerHUDScores
        obj.state=0
      next
      If Table1.ShowDT = true then
        For each obj in PlayerScores
          obj.visible=1
        Next
        For each obj in PlayerScoresOn
          obj.visible=0
        Next

        PlayerHuds(Player-1).SetValue(1)
        PlayerHUDScores(Player-1).state=1
        PlayerScores(Player-1).visible=0
        PlayerScoresOn(Player-1).visible=1
      end If
      ResetBalls
    End If
  Else

        If VRRoom = 1 then
    FlasherPlayers
    FlasherBalls
    FlasherCurrentPlayer
    End If

    If B2SOn Then
      Controller.B2SSetPlayerUp Player
      Controller.B2SSetBallInPlay BallInPlay
    End If
    PlaySound("RotateThruPlayers")
    TempPlayerUp=Player
'   PlayerUpRotator.enabled=true
'   PlayStartBall.enabled=true
    For each obj in PlayerHuds
      obj.SetValue(0)
    next
    For each obj in PlayerHUDScores
        obj.state=0
      next
    If Table1.ShowDT = true then
      For each obj in PlayerScores
        obj.visible=1
      Next
      For each obj in PlayerScoresOn
        obj.visible=0
      Next
      PlayerHuds(Player-1).SetValue(1)
      PlayerHUDScores(Player-1).state=1
      PlayerScores(Player-1).visible=0
      PlayerScoresOn(Player-1).visible=1
    end If
    ResetBalls
  End If

End sub

sub CheckHighScore
  Dim playertops
    dim si
  dim sj
  dim stemp
  dim stempplayers
  for i=1 to 4
    sortscores(i)=0
    sortplayers(i)=0
  next
  playertops=0
  for i = 1 to Players
    sortscores(i)=Score(i)
    sortplayers(i)=i
  next

  for si = 1 to Players
    for sj = 1 to Players-1
      if sortscores(sj)>sortscores(sj+1) then
        stemp=sortscores(sj+1)
        stempplayers=sortplayers(sj+1)
        sortscores(sj+1)=sortscores(sj)
        sortplayers(sj+1)=sortplayers(sj)
        sortscores(sj)=stemp
        sortplayers(sj)=stempplayers
      end if
    next
  next
  ScoreChecker=4
  CheckAllScores=1
  NewHighScore sortscores(ScoreChecker),sortplayers(ScoreChecker)
  savehs
end sub


sub checkmatch
  Dim tempmatch
  tempmatch=Int(Rnd*10)
  Match=tempmatch*10
  MatchReel.SetValue(tempmatch+1)

    if VRRoom = 1 then
  FlasherMatch
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

Sub TiltTimer_Timer()
  if TiltCount > 0 then TiltCount = TiltCount - 1
  if TiltCount = 0 then
    TiltTimer.Enabled = False
  end if
end sub

Sub TiltIt()
    TiltCount = TiltCount + 1
    if TiltCount = 3 then
      TableTilted=True
      TiltReel.SetValue(1)
      PlasticsOff
      BumpersOff
      StopSound "buzzL" 'scampa123 to fix tilt bug with flipper buzz
      StopSound "buzz" 'scampa123 to fix tilt bug with flipper buzz
      LeftFlipper.RotateToStart
      RightFlipper.RotateToStart
      for each Object in ColFlTilt : object.visible = 1 : next
      If B2Son then
        Controller.B2SSetTilt 1
      end if
    else
      TiltTimer.Interval = 500
      TiltTimer.Enabled = True
    end if

end sub

Sub IncreaseBonus()


' If BonusCounter=10 then
'   exit sub
' end if
  AddBonusBalls(1)
' BonusCounter=BonusCounter+1

' if BonusCounter>10 then
'   Bonus(10).State=1
'   Bonus(BonusCounter-10).state=1
' else
'   Bonus(BonusCounter).State=1
' end if

' LeftOutlaneLight.state=0
' RightOutlaneLight.state=0
' if BonusMultiplier=2 then
'   DoubleBonus.state=1
' end if

' if (BonusSpecialThreshold=0) then
'   if (BonusCounter=8) or (BonusCounter=12) then
'     LeftOutlaneLight.state=1
'     RightOutlaneLight.state=1
'   end if
' else
'   if (BonusCounter=10) or (BonusCounter=15) then
'     LeftOutlaneLight.state=1
'     RightOutlaneLight.state=1
'   end if
' end if
End Sub


Sub BonusBoost_Timer()
  IncreaseBonus
  BonusBoosterCounter=BonusBoosterCounter-1
  If BonusBoosterCounter=0 then
    BonusBoost.enabled=false
  end if

end sub

Sub CheckForLightSpecial()

  if (TopLightA.state=0) and (TopLightB.state=0) and (TopLightC.state=0) then
    TopRightTargetLight.State=1
    TopLeftTargetLight.State=1
  end if
end sub

Sub PlayStartBall_timer()

  PlayStartBall.enabled=false
  PlaySound("StartBall2-5")
end sub

Sub PlayerUpRotator_timer()
    If RotatorTemp<5 then
      TempPlayerUp=TempPlayerUp+1
      If TempPlayerUp>4 then
        TempPlayerUp=1
      end if
      For each obj in PlayerHuds
        obj.SetValue(0)
      next
      For each obj in PlayerHUDScores
        obj.state=0
      next
      If Table1.ShowDT = True then
        For each obj in PlayerScores
          obj.visible=1
        Next
        For each obj in PlayerScoresOn
          obj.visible=0
        Next
        PlayerHuds(TempPlayerUp-1).SetValue(1)
        PlayerHUDScores(TempPlayerUp-1).state=1
        PlayerScores(TempPlayerUp-1).visible=0
        PlayerScoresOn(TempPlayerUp-1).visible=1
      end If
      If B2SOn Then
        Controller.B2SSetPlayerUp TempPlayerUp
        Controller.B2SSetData 81,0
        Controller.B2SSetData 82,0
        Controller.B2SSetData 83,0
        Controller.B2SSetData 84,0
        Controller.B2SSetData 80+TempPlayerUp,1
      End If

    else

            if VRRoom = 1 then
      FlasherCurrentPlayer
      end If

      if B2SOn then
        Controller.B2SSetPlayerUp Player
        Controller.B2SSetData 81,0
        Controller.B2SSetData 82,0
        Controller.B2SSetData 83,0
        Controller.B2SSetData 84,0
        Controller.B2SSetData 80+Player,1
      end if
      PlayerUpRotator.enabled=false
      RotatorTemp=1
      For each obj in PlayerHuds
        obj.SetValue(0)
      next
      For each obj in PlayerHUDScores
        obj.state=0
      next
      If Table1.ShowDT = true then
        For each obj in PlayerScores
          obj.visible=1
        Next
        For each obj in PlayerScoresOn
          obj.visible=0
        Next
        PlayerHuds(Player-1).SetValue(1)
        PlayerHUDScores(Player-1).state=1
        PlayerScores(Player-1).visible=0
        PlayerScoresOn(Player-1).visible=1
      end If
    end if
    RotatorTemp=RotatorTemp+1


end sub

sub savehs
  ' Based on Black's Highscore routines
  Dim FileObj
  Dim ScoreFile
  Set FileObj=CreateObject("Scripting.FileSystemObject")
  If Not FileObj.FolderExists(UserDirectory) then
    Exit Sub
  End if
  Set ScoreFile=FileObj.CreateTextFile(UserDirectory & HSFileName,True)
    ScoreFile.WriteLine 0
    ScoreFile.WriteLine Credits
    scorefile.writeline BallsPerGame
    scorefile.writeline SpecialLightFlag
    scorefile.writeline ReplayLevel
    for xx=1 to 5
      scorefile.writeline HSScore(xx)
    next
    for xx=1 to 5
      scorefile.writeline HSName(xx)
    next
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
  dim temp8
  dim temp9
  dim temp10
  dim temp11
  dim temp12
  dim temp13
  dim temp14
  dim temp15


    Set FileObj=CreateObject("Scripting.FileSystemObject")
  If Not FileObj.FolderExists(UserDirectory) then
    Exit Sub
  End if
  If Not FileObj.FileExists(UserDirectory & HSFileName) then
    Exit Sub
  End if
  Set ScoreFile=FileObj.GetFile(UserDirectory & HSFileName)
  Set TextStr=ScoreFile.OpenAsTextStream(1,0)
    If (TextStr.AtEndOfStream=True) then
      Exit Sub
    End if
    temp1=TextStr.ReadLine
    temp2=textstr.readline
    temp3=textstr.readline
    temp4=textstr.readline
    temp5=textstr.readline
    HighScore=cdbl(temp1)
    if HighScore<1 then
      temp6=textstr.readline
      temp7=textstr.readline
      temp8=textstr.readline
      temp9=textstr.readline
      temp10=textstr.readline
      temp11=textstr.readline
      temp12=textstr.readline
      temp13=textstr.readline
      temp14=textstr.readline
      temp15=textstr.readline
    end if
    TextStr.Close
      Credits=cdbl(temp2)

'Setup VR Credit Reel position based on saved information..  This runs at table load
' We are going to set the VR BG credits reel here...
if VRRoom = 1 then
if Credits = 0 then Creditwheel.objroty = 11
if Credits = 1 then Creditwheel.objroty = 27.5
if Credits = 2 then Creditwheel.objroty = 44
if Credits = 3 then Creditwheel.objroty = 60.5
if Credits = 4 then Creditwheel.objroty = 77
if Credits = 5 then Creditwheel.objroty = 110  'extra 16.5 degree jump here..
if Credits = 6 then Creditwheel.objroty = 126.5
if Credits = 7 then Creditwheel.objroty = 143
if Credits = 8 then Creditwheel.objroty = 159.5
if Credits = 9 then Creditwheel.objroty = 174.5
end If


    BallsPerGame=cdbl(temp3)
    SpecialLightFlag=cdbl(temp4)
    ReplayLevel=cdbl(temp5)
    if HighScore<1 then
      HSScore(1) = int(temp6)
      HSScore(2) = int(temp7)
      HSScore(3) = int(temp8)
      HSScore(4) = int(temp9)
      HSScore(5) = int(temp10)

      HSName(1) = temp11
      HSName(2) = temp12
      HSName(3) = temp13
      HSName(4) = temp14
      HSName(5) = temp15
    end if
    Set ScoreFile=Nothing
      Set FileObj=Nothing
end sub

sub SaveLMEMConfig
  Dim FileObj
  Dim LMConfig
  dim temp1
  dim tempb2s
  tempb2s=0
  if B2SOn=true then
    tempb2s=1
  else
    tempb2s=0
  end if
  Set FileObj=CreateObject("Scripting.FileSystemObject")
  If Not FileObj.FolderExists(UserDirectory) then
    Exit Sub
  End if
  Set LMConfig=FileObj.CreateTextFile(UserDirectory & LMEMTableConfig,True)
  LMConfig.WriteLine tempb2s
  LMConfig.Close
  Set LMConfig=Nothing
  Set FileObj=Nothing

end Sub

sub LoadLMEMConfig
  Dim FileObj
  Dim LMConfig
  dim tempC
  dim tempb2s

    Set FileObj=CreateObject("Scripting.FileSystemObject")
  If Not FileObj.FolderExists(UserDirectory) then
    Exit Sub
  End if
  If Not FileObj.FileExists(UserDirectory & LMEMTableConfig) then
    Exit Sub
  End if
  Set LMConfig=FileObj.GetFile(UserDirectory & LMEMTableConfig)
  Set TextStr2=LMConfig.OpenAsTextStream(1,0)
  If (TextStr2.AtEndOfStream=True) then
    Exit Sub
  End if
  tempC=TextStr2.ReadLine
  TextStr2.Close
  tempb2s=cdbl(tempC)
  if tempb2s=0 then
    B2SOn=false
  else
    B2SOn=true
  end if
  Set LMConfig=Nothing
  Set FileObj=Nothing
end sub

sub SaveLMEMConfig2
  If ShadowConfigFile=false then exit sub
  Dim FileObj
  Dim LMConfig2
  dim temp1
  dim temp2
  dim tempBS
  dim tempFS

  if EnableBallShadow=true then
    tempBS=1
  else
    tempBS=0
  end if
  if EnableFlipperShadow=true then
    tempFS=1
  else
    tempFS=0
  end if

  Set FileObj=CreateObject("Scripting.FileSystemObject")
  If Not FileObj.FolderExists(UserDirectory) then
    Exit Sub
  End if
  Set LMConfig2=FileObj.CreateTextFile(UserDirectory & LMEMShadowConfig,True)
  LMConfig2.WriteLine tempBS
  LMConfig2.WriteLine tempFS
  LMConfig2.Close
  Set LMConfig2=Nothing
  Set FileObj=Nothing

end Sub

sub LoadLMEMConfig2
  If ShadowConfigFile=false then

    EnableBallShadow = ShadowBallOn
    BallShadowUpdate.enabled = ShadowBallOn

      'EnableBallShadow = false
    'BallShadowUpdate.enabled = false


    EnableFlipperShadow = ShadowFlippersOn
    FlipperLSh.visible = ShadowFlippersOn
    FlipperRSh.visible = ShadowFlippersOn
    exit sub
  end if
  Dim FileObj
  Dim LMConfig2
  dim tempC
  dim tempD
  dim tempFS
  dim tempBS

    Set FileObj=CreateObject("Scripting.FileSystemObject")
  If Not FileObj.FolderExists(UserDirectory) then
    Exit Sub
  End if
  If Not FileObj.FileExists(UserDirectory & LMEMShadowConfig) then
    Exit Sub
  End if
  Set LMConfig2=FileObj.GetFile(UserDirectory & LMEMShadowConfig)
  Set TextStr2=LMConfig2.OpenAsTextStream(1,0)
  If (TextStr2.AtEndOfStream=True) then
    Exit Sub
  End if
  tempC=TextStr2.ReadLine
  tempD=TextStr2.Readline
  TextStr2.Close
  tempBS=cdbl(tempC)
  tempFS=cdbl(tempD)
  if tempBS=0 then
    EnableBallShadow=false
    BallShadowUpdate.enabled=false
  else
    EnableBallShadow=true
        'EnableBallShadow=false
  end if
  if tempFS=0 then
    EnableFlipperShadow=false
    FlipperLSh.visible=false
    FLipperRSh.visible=false
  else
    EnableFlipperShadow=true
  end if
  Set LMConfig2=Nothing
  Set FileObj=Nothing

end sub


Sub DisplayHighScore

end sub


sub InitPauser5_timer
    If B2SOn Then
      'Controller.B2SSetScorePlayer6 HighScore
    End If
    DisplayHighScore
    CreditsReel.SetValue(Credits)
    EMReel7.SetValue(Credits)
    InitPauser5.enabled=false
end sub



sub BumpersOff
  Bumper1Light.Visible=false
  Bumper2Light.Visible=false

  Light1.state=0
  Light2.state=0
  LeftInlaneLight.state=0
  RightInlaneLight.state=0
  RightCenterTargetLight.state=0
  RightLowerTargetLight.state=0
  LeftRolloverLight.state=0
  UpperLight1.state=0
  UpperLight2.state=0
  LeftLowerTargetLight.state=0
  RightUpperTargetLight.state=0
  If RightRolloverLight1.state=1 then
    TempLightTracker=1
  else
    TempLightTracker=2
  end if

  for each obj in RolloverLights2
    obj.state=0
  next
  for each obj in RolloverLights1
    obj.state=0
  next

end sub

sub BumpersOn
  Bumper1Light.state=1
  Bumper2Light.state=1
  Bumper1Light.Visible=true
  Bumper2Light.Visible=True


  Light1.state=1
  Light2.state=1

  UpperLight1.state=1
  UpperLight2.state=1
  LeftLowerTargetLight.state=1
  RightUpperTargetLight.state=1

  If BallsPerGame=3 then
    If BallInPlay>=1 then
      LeftInlaneLight.state=1
      RightInlaneLight.state=1
    end if
    if BallInPlay>=2 then
      RightCenterTargetLight.state=1
      RightLowerTargetLight.state=1
    end if
    if BallInPlay>=3 then
      LeftRolloverLight.state=1
    end if
  else
    If BallInPlay>=3 then
      LeftInlaneLight.state=1
      RightInlaneLight.state=1
    end if
    if BallInPlay>=4 then
      RightCenterTargetLight.state=1
      RightLowerTargetLight.state=1
    end if
    if BallInPlay>=5 then
      LeftRolloverLight.state=1
    end if
  end if


  If TempLightTracker=1 then
    for each obj in RolloverLights2
      obj.state=0
    next
    for each obj in RolloverLights1
      obj.state=1
    next
  else
    for each obj in RolloverLights2
      obj.state=1
    next
    for each obj in RolloverLights1
      obj.state=0
    next
  end if

end sub

Sub PlasticsOn
  For each obj in Flashers
    obj.state=1
  Next
end sub

Sub PlasticsOff
  For each obj in Flashers
    obj.state=0
  Next

end sub


Sub SetupReplayTables

  Replay1Table(1)=57000
  Replay1Table(2)=59000
  Replay1Table(3)=61000
  Replay1Table(4)=64000
  Replay1Table(5)=66000
  Replay1Table(6)=68000
  Replay1Table(7)=70000
  Replay1Table(8)=71000
  Replay1Table(9)=74000
  Replay1Table(10)=77000
  Replay1Table(11)=81000
  Replay1Table(12)=63000
  Replay1Table(13)=66000
  Replay1Table(14)=69000
  Replay1Table(15)=999000

  Replay2Table(1)=71000
  Replay2Table(2)=73000
  Replay2Table(3)=75000
  Replay2Table(4)=78000
  Replay2Table(5)=80000
  Replay2Table(6)=82000
  Replay2Table(7)=84000
  Replay2Table(8)=85000
  Replay2Table(9)=88000
  Replay2Table(10)=91000
  Replay2Table(11)=95000
  Replay2Table(12)=77000
  Replay2Table(13)=80000
  Replay2Table(14)=83000
  Replay2Table(15)=999000

  Replay3Table(1)=79000
  Replay3Table(2)=81000
  Replay3Table(3)=83000
  Replay3Table(4)=86000
  Replay3Table(5)=88000
  Replay3Table(6)=90000
  Replay3Table(7)=92000
  Replay3Table(8)=93000
  Replay3Table(9)=96000
  Replay3Table(10)=999000
  Replay3Table(11)=999000
  Replay3Table(12)=999000
  Replay3Table(13)=999000
  Replay3Table(14)=999000
  Replay3Table(15)=999000

  ReplayTableMax=10

end sub

Sub RefreshReplayCard
  Dim tempst1
  Dim tempst2

  tempst1=FormatNumber(BallsPerGame,0)
  tempst2=FormatNumber(ReplayLevel,0)

  ReplayCard.image = "SC" + tempst2
  Replay1=Replay1Table(ReplayLevel)
  Replay2=Replay2Table(ReplayLevel)
  Replay3=Replay3Table(ReplayLevel)
end sub

'****************************************
'BONUS BALLS ROUTINES
'****************************************
Sub BallAdderTimer_timer()

  EMReel6.SetValue(AddABall)
  X=AddABall+1
  if B2SOn then
    Controller.B2SSetData CurrentFrame,0
    Controller.B2SSetData X+100,1
    CurrentFrame=X+100
  end if
  AddABallFrames=AddABallFrames-1
  AddABall=AddABall+1

  If AddABallFrames=0 then
    BallAdderTimer.enabled=0
    BallCounter=BallCounter+1
    BonusMotorRunning=0
  end if

end Sub

Sub BallDropperTimer_timer()

  EMReel6.SetValue(DropABall)
  X=DropABall+1
  if B2SOn then
    Controller.B2SSetData CurrentFrame,0
    Controller.B2SSetData X+100,1
    CurrentFrame=X+100
  end if
  DropABallFrames=DropABallFrames-1
  DropABall=DropABall+1
  If DropABallFrames=0 then
    BallDropperTimer.enabled=0
    BallCounter=BallCounter-1
    BallDropperDelayer.enabled=1
  end if

end Sub

Sub BallDropperDelayer_Timer()
  BallDropperDelayer.enabled=0
  BonusMotorRunning=0
end sub

Sub BuildBallReelTables

  BallReelAddFrames(0)=13
  BallReelAddFrames(1)=12
  BallReelAddFrames(2)=11
  BallReelAddFrames(3)=10
  BallReelAddFrames(4)=9
  BallReelAddFrames(5)=8
  BallReelAddFrames(6)=7
  BallReelAddFrames(7)=6
  BallReelAddFrames(8)=5
  BallReelAddFrames(9)=5

  BallReelAddStart(0)=1
  BallReelAddStart(1)=14
  BallReelAddStart(2)=26
  BallReelAddStart(3)=37
  BallReelAddStart(4)=47
  BallReelAddStart(5)=56
  BallReelAddStart(6)=64
  BallReelAddStart(7)=71
  BallReelAddStart(8)=77
  BallReelAddStart(9)=82


  BallReelDropStart(0)=131
  BallReelDropStart(1)=126
  BallReelDropStart(2)=121
  BallReelDropStart(3)=116
  BallReelDropStart(4)=112
  BallReelDropStart(5)=107
  BallReelDropStart(6)=102
  BallReelDropStart(7)=97
  BallReelDropStart(8)=92
  BallReelDropStart(9)=87

  BallReelDropFrames(0)=5
  BallReelDropFrames(1)=5
  BallReelDropFrames(2)=5
  BallReelDropFrames(3)=5
  BallReelDropFrames(4)=4
  BallReelDropFrames(5)=5
  BallReelDropFrames(6)=5
  BallReelDropFrames(7)=5
  BallReelDropFrames(8)=5
  BallReelDropFrames(9)=6
  x=87
  For z = 9 to 0
    BallReelDropStart(z)=x
    x=x+5
  next

end sub

Sub AddBonusBalls(y)

  QueuedBonusAdds=QueuedBonusAdds+y

end sub

Sub BonusMotorAdder_Timer
  ScoreText1.text=QueuedBonusAdds
  TextBox4.text=BallCounter
  TextBox5.text=BonusMotorRunning
  if BonusMotorRunning<>1 and InProgress=true and QueuedBonusAdds>0 then
    if BallCounter<10 then
      AddABall=BallReelAddStart(BallCounter)
      AddABallFrames=BallReelAddFrames(BallCounter)
      BallAdderTimer.enabled=1
      If TableTilted=false then
        SetMotor(500)
      end if
      BonusMotorRunning=1
      QueuedBonusAdds=QueuedBonusAdds-1
      PlaySound "solenoid"

if VRRoom = 1 then
VRBGKicker.Kick 180, 35,1.5  ' We are kicking the ball up in the VR BG here..

' Now create a new ball in the trough if the count is under 10
if CreatedBallNumber <10 then

CreatedBallNumber =  CreatedBallNumber + 1
VRBGKicker2.createball
VRBGKicker2.Kick 0+Z, 5

end if
end if

    else
      QueuedBonusAdds=QueuedBonusAdds-1
      If TableTilted=false then
        SetMotor(500)
      end if
    end if
  end if

end sub

Sub BonusMotorDropper_Timer
  ScoreText1.text=QueuedBonusDrops
  TextBox4.text=BallCounter
  TextBox5.text=BonusMotorRunning
  if BonusMotorRunning<>1 and InProgress=true and QueuedBonusDrops>0 then
    if BallCounter>0 then
      If TableTilted=false then
        SetMotor(1000*BonusMultiplier)
      end if
      DropABall=BallReelDropStart(BallCounter-1)
      DropABallFrames=BallReelDropFrames(BallCounter-1)
      BallDropperTimer.enabled=1
      BonusMotorRunning=1
      QueuedBonusDrops=QueuedBonusDrops-1
      PlaySound "solenoid"


' THIS IS WHERE WE OPEN THE WALL TO LET A BALL OUT FOR VR Backglass balls...
if VRRoom = 1 then
VRBGDropWall.isdropped = TRUE
VRBGDropwallTimer.enabled = true
VRBGDropWall2.sidevisible = false
VRBGDropWall2.visible = false

BallUpTop = 0  ' count for how many balls are on top..reset to 0 everytime a ball goes through trough

  end if
    end if
  end if

end sub

'****************************************
'  SCORE MOTOR
'****************************************

ScoreMotorTimer.Enabled = 1
ScoreMotorTimer.Interval = 135 '135
AddScoreTimer.Enabled = 1
AddScoreTimer.Interval = 135

Dim queuedscore
Dim MotorMode
Dim MotorPosition

Sub SetMotor(y)
  Select Case ScoreMotorAdjustment
    Case 0:
      queuedscore=queuedscore+y
    Case 1:
      If MotorRunning<>1 And InProgress=true then
        queuedscore=queuedscore+y
      end if
    end Select
end sub

Sub SetMotor2(x)
  If MotorRunning<>1 And InProgress=true then
    MotorRunning=1

    Select Case x
      Case 10:
        AddScore(10)
        MotorRunning=0
        BumpersOn

      Case 20:
        MotorMode=10
        MotorPosition=2
        BumpersOff
      Case 30:
        MotorMode=10
        MotorPosition=3
        BumpersOff
      Case 40:
        MotorMode=10
        MotorPosition=4
        BumpersOff
      Case 50:
        MotorMode=10
        MotorPosition=5
        BumpersOff
      Case 100:
        AddScore(100)
        MotorRunning=0
        BumpersOn
        SpecialIsLit=0
      Case 200:
        MotorMode=100
        MotorPosition=2
        BumpersOff
      Case 300:
        MotorMode=100
        MotorPosition=3
        BumpersOff
      Case 400:
        MotorMode=100
        MotorPosition=4
        BumpersOff
      Case 500:
        MotorMode=100
        MotorPosition=5
        BumpersOff
        SpecialIsLit=0
        DropTargetLight.state=0
      Case 1000:
        AddScore(1000)
        MotorRunning=0
        BumpersOn
      Case 2000:
        MotorMode=1000
        MotorPosition=2
        BumpersOff
      Case 3000:
        MotorMode=1000
        MotorPosition=3
        BumpersOff
      Case 4000:
        MotorMode=1000
        MotorPosition=4
        BumpersOff
      Case 5000:
        MotorMode=1000
        MotorPosition=5
        BumpersOff
    End Select
  End If
End Sub

Sub AddScoreTimer_Timer
  Dim tempscore

  If MotorRunning<>1 And InProgress=true then
    if queuedscore>=5000 then
      tempscore=5000
      queuedscore=queuedscore-5000
      SetMotor2(5000)
      exit sub
    end if
    if queuedscore>=4000 then
      tempscore=4000
      queuedscore=queuedscore-4000
      SetMotor2(4000)
      exit sub
    end if

    if queuedscore>=3000 then
      tempscore=3000
      queuedscore=queuedscore-3000
      SetMotor2(3000)
      exit sub
    end if

    if queuedscore>=2000 then
      tempscore=2000
      queuedscore=queuedscore-2000
      SetMotor2(2000)
      exit sub
    end if

    if queuedscore>=1000 then
      tempscore=1000
      queuedscore=queuedscore-1000
      SetMotor2(1000)
      exit sub
    end if

    if queuedscore>=500 then
      tempscore=500
      queuedscore=queuedscore-500
      SetMotor2(500)
      exit sub
    end if
    if queuedscore>=400 then
      tempscore=400
      queuedscore=queuedscore-400
      SetMotor2(400)
      exit sub
    end if
    if queuedscore>=300 then
      tempscore=300
      queuedscore=queuedscore-300
      SetMotor2(300)
      exit sub
    end if
    if queuedscore>=200 then
      tempscore=200
      queuedscore=queuedscore-200
      SetMotor2(200)
      exit sub
    end if
    if queuedscore>=100 then
      tempscore=100
      queuedscore=queuedscore-100
      SetMotor2(100)
      exit sub
    end if

    if queuedscore>=50 then
      tempscore=50
      queuedscore=queuedscore-50
      SetMotor2(50)
      exit sub
    end if
    if queuedscore>=40 then
      tempscore=40
      queuedscore=queuedscore-40
      SetMotor2(40)
      exit sub
    end if
    if queuedscore>=30 then
      tempscore=30
      queuedscore=queuedscore-30
      SetMotor2(30)
      exit sub
    end if
    if queuedscore>=20 then
      tempscore=20
      queuedscore=queuedscore-20
      SetMotor2(20)
      exit sub
    end if
    if queuedscore>=10 then
      tempscore=10
      queuedscore=queuedscore-10
      SetMotor2(10)
      exit sub
    end if


  End If


end Sub

Sub ScoreMotorTimer_Timer
  If MotorPosition > 0 Then
    Select Case MotorPosition
      Case 5,4,3,2:
        If MotorMode=1000 Then
          AddScore(1000)
        end if
        if MotorMode=100 then
          AddScore(100)
        End If
        if MotorMode=10 then
          AddScore(10)
        End if
        MotorPosition=MotorPosition-1
      Case 1:
        If MotorMode=1000 Then
          AddScore(1000)
        end if
        If MotorMode=100 then
          AddScore(100)
        End If
        if MotorMode=10 then
          AddScore(10)
        End if
        MotorPosition=0:MotorRunning=0:BumpersOn
    End Select
  End If

End Sub


Sub AddScore(x)
  If TableTilted=true then exit sub
  Select Case ScoreAdditionAdjustment
    Case 0:
      AddScore1(x)
    Case 1:
      AddScore2(x)
  end Select

end sub


Sub AddScore1(x)
' debugtext.text=score
  Select Case x
    Case 1:
      PlayChime(10)
      Score(Player)=Score(Player)+1

    Case 10:
      PlayChime(10)
      Score(Player)=Score(Player)+10
'     debugscore=debugscore+10

    Case 100:
      PlayChime(100)
      Score(Player)=Score(Player)+100
'     debugscore=debugscore+100

    Case 1000:
      PlayChime(1000)
      Score(Player)=Score(Player)+1000
'     debugscore=debugscore+1000
  End Select
  PlayerScores(Player-1).AddValue(x)
  PlayerScoresOn(Player-1).AddValue(x)
  If ScoreDisplay(Player)<100000 then
    ScoreDisplay(Player)=Score(Player)
  Else
    Score100K(Player)=Int(Score(Player)/100000)
    ScoreDisplay(Player)=Score(Player)-100000
  End If
  if Score(Player)=>100000 then
    If B2SOn Then
      If Player=1 Then
        Controller.B2SSetScoreRolloverPlayer1 Score100K(Player)
      End If
      If Player=2 Then
        Controller.B2SSetScoreRolloverPlayer2 Score100K(Player)
      End If

      If Player=3 Then
        Controller.B2SSetScoreRolloverPlayer3 Score100K(Player)
      End If

      If Player=4 Then
        Controller.B2SSetScoreRolloverPlayer4 Score100K(Player)
      End If
    End If
  End If
  If B2SOn Then
    Controller.B2SSetScorePlayer Player, ScoreDisplay(Player)
  End If
  If Score(Player)>Replay1 and Replay1Paid(Player)=false then
    Replay1Paid(Player)=True
    AddSpecial
  End If
  If Score(Player)>Replay2 and Replay2Paid(Player)=false then
    Replay2Paid(Player)=True
    AddSpecial
  End If
  If Score(Player)>Replay3 and Replay3Paid(Player)=false then
    Replay3Paid(Player)=True
    AddSpecial
  End If
' ScoreText.text=debugscore
End Sub

Sub AddScore2(x)
  Dim OldScore, NewScore, OldTestScore, NewTestScore
    OldScore = Score(Player)

  Select Case x
        Case 1:
            Score(Player)=Score(Player)+1
    Case 10:
      Score(Player)=Score(Player)+10
    Case 100:
      Score(Player)=Score(Player)+100
    Case 1000:
      Score(Player)=Score(Player)+1000
  End Select
  NewScore = Score(Player)
  if Score(Player)=>100000 then
    If B2SOn Then
      If Player=1 Then
        Controller.B2SSetScoreRolloverPlayer1 1
      End If
      If Player=2 Then
        Controller.B2SSetScoreRolloverPlayer2 1
      End If

      If Player=3 Then
        Controller.B2SSetScoreRolloverPlayer3 1
      End If

      If Player=4 Then
        Controller.B2SSetScoreRolloverPlayer4 1
      End If
    End If
  End If

  OldTestScore = OldScore
  NewTestScore = NewScore
  Do
    if OldTestScore < Replay1 and NewTestScore >= Replay1 then
      AddSpecial()
      NewTestScore = 0
    Elseif OldTestScore < Replay2 and NewTestScore >= Replay2 then
      AddSpecial()
      NewTestScore = 0
    Elseif OldTestScore < Replay3 and NewTestScore >= Replay3 then
      AddSpecial()
      NewTestScore = 0
    End if
    NewTestScore = NewTestScore - 100000
    OldTestScore = OldTestScore - 100000
  Loop While NewTestScore > 0

    OldScore = int(OldScore / 10) ' divide by 10 for games with fixed 0 in 1s position, by 1 for games with real 1s digits
    NewScore = int(NewScore / 10) ' divide by 10 for games with fixed 0 in 1s position, by 1 for games with real 1s digits
  ' MsgBox("OldScore="&OldScore&", NewScore="&NewScore&", OldScore Mod 10="&OldScore Mod 10 & ", NewScore % 10="&NewScore Mod 10)

    if (OldScore Mod 10 <> NewScore Mod 10) then
    PlayChime(10)
    if BallsPerGame=5 and SpecialLightFlag=1 and SpecialIsLit=0 then
      ToggleSpecial
    end if
    if BallsPerGame=3 and SpecialLightFlag=1 and SpecialIsLit=1 then
      ToggleSpecial
    end if
    end if

    OldScore = int(OldScore / 10)
    NewScore = int(NewScore / 10)
  ' MsgBox("OldScore="&OldScore&", NewScore="&NewScore)
    if (OldScore Mod 10 <> NewScore Mod 10) then
    PlayChime(100)
    if BallsPerGame=5 and SpecialLightFlag=1 and SpecialIsLit=1 then
      ToggleSpecial
    end if
    if BallsPerGame=3 and SpecialLightFlag=1 and SpecialIsLit=0 then
      ToggleSpecial
    end if
    end if

    OldScore = int(OldScore / 10)
    NewScore = int(NewScore / 10)
  ' MsgBox("OldScore="&OldScore&", NewScore="&NewScore)
    if (OldScore Mod 10 <> NewScore Mod 10) then
    PlayChime(1000)

    end if

    OldScore = int(OldScore / 10)
    NewScore = int(NewScore / 10)
  ' MsgBox("OldScore="&OldScore&", NewScore="&NewScore)
    if (OldScore Mod 10 <> NewScore Mod 10) then
    PlayChime(1000)
    end if

  If B2SOn Then
    Controller.B2SSetScorePlayer Player, Score(Player)
  End If
' EMReel1.SetValue Score(Player)
  PlayerScores(Player-1).AddValue(x)
  PlayerScoresOn(Player-1).AddValue(x)
End Sub

Sub PlayChime(x)
  if ChimesOn=0 then
    Select Case x
      Case 10
        If LastChime10=1 Then
          PlaySound SoundFXDOF("SpinACard_1_10_Point_Bell",322,DOFPulse,DOFChimes)
          LastChime10=0
        Else
          PlaySound SoundFXDOF("SpinACard_1_10_Point_Bell",322,DOFPulse,DOFChimes)
          LastChime10=1
        End If
      Case 100
        If LastChime100=1 Then
          PlaySound SoundFXDOF("SpinACard_100_Point_Bell",323,DOFPulse,DOFChimes)
          LastChime100=0
        Else
          PlaySound SoundFXDOF("SpinACard_100_Point_Bell",323,DOFPulse,DOFChimes)
          LastChime100=1
        End If

    End Select

  EMMODE = 1
  UpdateReels 0,0 ,score(0), 0, -1,-1,-1,-1,-1
  UpdateReels 1,1 ,score(1), 0, -1,-1,-1,-1,-1
  UpdateReels 2,2 ,score(2), 0, -1,-1,-1,-1,-1
  UpdateReels 3,3 ,score(3), 0, -1,-1,-1,-1,-1
  UpdateReels 4,4 ,score(4), 0, -1,-1,-1,-1,-1
  EMMODE = 0 ' restore EM mode

  else
    Select Case x
      Case 10
        If LastChime10=1 Then
          PlaySound SoundFXDOF("SJ_Chime_10a",322,DOFPulse,DOFChimes)
          LastChime10=0
        Else
          PlaySound SoundFXDOF("SJ_Chime_10b",322,DOFPulse,DOFChimes)
          LastChime10=1
        End If
      Case 100
        If LastChime100=1 Then
          PlaySound SoundFXDOF("SJ_Chime_100a",323,DOFPulse,DOFChimes)
          LastChime100=0
        Else
          PlaySound SoundFXDOF("SJ_Chime_100b",323,DOFPulse,DOFChimes)
          LastChime100=1
        End If
      Case 1000
        If LastChime1000=1 Then
          PlaySound SoundFXDOF("SJ_Chime_1000a",324,DOFPulse,DOFChimes)
          LastChime1000=0
        Else
          PlaySound SoundFXDOF("SJ_Chime_1000b",324,DOFPulse,DOFChimes)
          LastChime1000=1
        End If
    End Select

  EMMODE = 1
  UpdateReels 0,0 ,score(0), 0, -1,-1,-1,-1,-1
  UpdateReels 1,1 ,score(1), 0, -1,-1,-1,-1,-1
  UpdateReels 2,2 ,score(2), 0, -1,-1,-1,-1,-1
  UpdateReels 3,3 ,score(3), 0, -1,-1,-1,-1,-1
  UpdateReels 4,4 ,score(4), 0, -1,-1,-1,-1,-1
  EMMODE = 0 ' restore EM mode

  end if
End Sub


Sub HideOptions()
end sub


'*****************************************
' ninuzzu's BALL SHADOW
'*****************************************
'Dim BallShadow
'BallShadow = Array (BallShadow1,BallShadow2,BallShadow3,BallShadow4,BallShadow5,BallShadow6,BallShadow7,BallShadow8,BallShadow9,BallShadow10,BallShadow11)

'Sub BallShadowUpdate_timer()
'    Dim BOT, b
 '   BOT = GetBalls
    ' hide shadow of deleted balls
 '   If UBound(BOT)<(tnob-1) Then
 '       For b = (UBound(BOT) + 1) to (tnob-1)
'            BallShadow(b).visible = 0
'        Next
'    End If
    ' exit the Sub if no balls on the table
'    If UBound(BOT) = -1 Then Exit Sub
    ' render the shadow for each ball
''    For b = 0 to UBound(BOT)
  '      If BOT(b).X < Table1.Width/2 Then
  '          BallShadow(b).X = ((BOT(b).X) - (Ballsize/6) + ((BOT(b).X - (Table1.Width/2))/21)) + 6
  '      Else
  '          BallShadow(b).X = ((BOT(b).X) + (Ballsize/6) + ((BOT(b).X - (Table1.Width/2))/21)) - 6
  '      End If
   '     ballShadow(b).Y = BOT(b).Y + 4
    '    If BOT(b).Z > 20 Then
    '        BallShadow(b).visible = 1
   '     Else
   '         BallShadow(b).visible = 0
   '     End If


'If BOT(b).image = "Ball_HDR2" Then   '  This is so the backglass balls don't render a shadow..
'   BallShadow(b).visible = 0
' End If

 '   Next
'End Sub


Sub NewBallShadowUpdate_timer() ' removed Ninuzzus above because we dont want shadows on the VR BG balls, and this is a single ball game anyways.

 If GameBall.X < Table1.Width/2 Then
            BallShadow1.X = ((GameBall.X) - (Ballsize/6) + ((GameBall.X - (Table1.Width/2))/12)) + 2
        Else
            BallShadow1.X = ((GameBall.X) + (Ballsize/6) + ((GameBall.X - (Table1.Width/2))/12)) - 2
        End If
       ballShadow1.Y = GameBall.Y + 12
        If GameBall.Z > 20 Then
            BallShadow1.visible = 1
        Else
            BallShadow1.visible = 0
        End If


End Sub


' ============================================================================================
' GNMOD - Multiple High Score Display and Collection
' ============================================================================================
Dim EnteringInitials    ' Normally zero, set to non-zero to enter initials
EnteringInitials = 0

Dim PlungerPulled
PlungerPulled = 0

Dim SelectedChar      ' character under the "cursor" when entering initials

Dim HSTimerCount      ' Pass counter for HS timer, scores are cycled by the timer
HSTimerCount = 5      ' Timer is initially enabled, it'll wrap from 5 to 1 when it's displayed

Dim InitialString     ' the string holding the player's initials as they're entered

Dim AlphaString       ' A-Z, 0-9, space (_) and backspace (<)
Dim AlphaStringPos      ' pointer to AlphaString, move forward and backward with flipper keys
AlphaString = "ABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789_<"

Dim HSNewHigh       ' The new score to be recorded

Dim HSScore(5)        ' High Scores read in from config file
Dim HSName(5)       ' High Score Initials read in from config file

' default high scores, remove this when the scores are available from the config file
HSScore(1) = 75000
HSScore(2) = 70000
HSScore(3) = 60000
HSScore(4) = 55000
HSScore(5) = 50000

HSName(1) = "AAA"
HSName(2) = "ZZZ"
HSName(3) = "XXX"
HSName(4) = "ABC"
HSName(5) = "BBB"

Sub HighScoreTimer_Timer

  if EnteringInitials then
    if HSTimerCount = 1 then
      SetHSLine 3, InitialString & MID(AlphaString, AlphaStringPos, 1)
      HSTimerCount = 2
    else
      SetHSLine 3, InitialString
      HSTimerCount = 1
    end if
  elseif InProgress then
    SetHSLine 1, "HIGH SCORE1"
    SetHSLine 2, HSScore(1)
    SetHSLine 3, HSName(1)
    HSTimerCount = 5  ' set so the highest score will show after the game is over
    HighScoreTimer.enabled=false
  elseif CheckAllScores then
    NewHighScore sortscores(ScoreChecker),sortplayers(ScoreChecker)

  else
    ' cycle through high scores
    HighScoreTimer.interval=2000
    HSTimerCount = HSTimerCount + 1
    if HsTimerCount > 5 then
      HSTimerCount = 1
    End If
    SetHSLine 1, "HIGH SCORE"+FormatNumber(HSTimerCount,0)
    SetHSLine 2, HSScore(HSTimerCount)
    SetHSLine 3, HSName(HSTimerCount)
  end if
End Sub

Function GetHSChar(String, Index)
  dim ThisChar
  dim FileName
  ThisChar = Mid(String, Index, 1)
  FileName = "PostIt"
  if ThisChar = " " or ThisChar = "" then
    FileName = FileName & "BL"
  elseif ThisChar = "<" then
    FileName = FileName & "LT"
  elseif ThisChar = "_" then
    FileName = FileName & "SP"
  else
    FileName = FileName & ThisChar
  End If
  GetHSChar = FileName
End Function

Sub SetHsLine(LineNo, String)
  dim Letter
  dim ThisDigit
  dim ThisChar
  dim StrLen
  dim LetterLine
  dim Index
  dim StartHSArray
  dim EndHSArray
  dim LetterName
  dim xfor
  StartHSArray=array(0,1,12,22)
  EndHSArray=array(0,11,21,31)
  StrLen = len(string)
  Index = 1

  for xfor = StartHSArray(LineNo) to EndHSArray(LineNo)
    Eval("HS"&xfor).image = GetHSChar(String, Index)
    Index = Index + 1
  next

End Sub

Sub NewHighScore(NewScore, PlayNum)
  if NewScore > HSScore(5) then
    HighScoreTimer.interval = 500
    HSTimerCount = 1
    AlphaStringPos = 1    ' start with first character "A"
    EnteringInitials = 1  ' intercept the control keys while entering initials
    InitialString = ""    ' initials entered so far, initialize to empty
    SetHSLine 1, "PLAYER "+FormatNumber(PlayNum,0)
    SetHSLine 2, "ENTER NAME"
    SetHSLine 3, MID(AlphaString, AlphaStringPos, 1)
    HSNewHigh = NewScore
    For xx=1 to HighScoreReward
      AddSpecial
    next
  End if
  ScoreChecker=ScoreChecker-1
  if ScoreChecker=0 then
    CheckAllScores=0
  end if
End Sub

Sub CollectInitials(keycode)
  If keycode = LeftFlipperKey Then
    ' back up to previous character
    AlphaStringPos = AlphaStringPos - 1
    if AlphaStringPos < 1 then
      AlphaStringPos = len(AlphaString)   ' handle wrap from beginning to end
      if InitialString = "" then
        ' Skip the backspace if there are no characters to backspace over
        AlphaStringPos = AlphaStringPos - 1
      End if
    end if
    SetHSLine 3, InitialString & MID(AlphaString, AlphaStringPos, 1)
    PlaySound "DropTargetDropped"
  elseif keycode = RightFlipperKey Then
    ' advance to next character
    AlphaStringPos = AlphaStringPos + 1
    if AlphaStringPos > len(AlphaString) or (AlphaStringPos = len(AlphaString) and InitialString = "") then
      ' Skip the backspace if there are no characters to backspace over
      AlphaStringPos = 1
    end if
    SetHSLine 3, InitialString & MID(AlphaString, AlphaStringPos, 1)
    PlaySound "DropTargetDropped"
  elseif keycode = StartGameKey or keycode = PlungerKey Then
    SelectedChar = MID(AlphaString, AlphaStringPos, 1)
    if SelectedChar = "_" then
      InitialString = InitialString & " "
      PlaySound("Ding10")
    elseif SelectedChar = "<" then
      InitialString = MID(InitialString, 1, len(InitialString) - 1)
      if len(InitialString) = 0 then
        ' If there are no more characters to back over, don't leave the < displayed
        AlphaStringPos = 1
      end if
      PlaySound("Ding100")
    else
      InitialString = InitialString & SelectedChar
      PlaySound("Ding10")
    end if
    if len(InitialString) < 3 then
      SetHSLine 3, InitialString & SelectedChar
    End If
  End If
  if len(InitialString) = 3 then
    ' save the score
    for i = 5 to 1 step -1
      if i = 1 or (HSNewHigh > HSScore(i) and HSNewHigh <= HSScore(i - 1)) then
        ' Replace the score at this location
        if i < 5 then
' MsgBox("Moving " & i & " to " & (i + 1))
          HSScore(i + 1) = HSScore(i)
          HSName(i + 1) = HSName(i)
        end if
' MsgBox("Saving initials " & InitialString & " to position " & i)
        EnteringInitials = 0
        HSScore(i) = HSNewHigh
        HSName(i) = InitialString
        HSTimerCount = 5
        HighScoreTimer_Timer
        HighScoreTimer.interval = 2000
        PlaySound("Ding1000")
        exit sub
      elseif i < 5 then
        ' move the score in this slot down by 1, it's been exceeded by the new score
' MsgBox("Moving " & i & " to " & (i + 1))
        HSScore(i + 1) = HSScore(i)
        HSName(i + 1) = HSName(i)
      end if
    next
  End If

End Sub
' END GNMOD
' ============================================================================================
' GNMOD - New Options menu
' ============================================================================================
Dim EnteringOptions
Dim CurrentOption
Dim OptionCHS
Dim MaxOption
Dim OptionHighScorePosition
Dim XOpt
Dim StartingArray
Dim EndingArray

StartingArray=Array(0,1,2,30,33,61,89,117,145,173,201,229)
EndingArray=Array(0,1,29,32,60,88,116,144,172,200,228,256)
EnteringOptions = 0
MaxOption = 9
OptionCHS = 0
OptionHighScorePosition = 0
Const OptionLinesToMark="111100011"
Const OptionLine1="" 'do not use this line
Const OptionLine2="" 'do not use this line
Const OptionLine3="" 'do not use this line
Const OptionLine4="Special Light Rollover"
Const OptionLine5=""
Const OptionLine6=""
Const OptionLine7=""
Const OptionLine8="" 'do not use this line
Const OptionLine9="" 'do not use this line

Sub OperatorMenuTimer_Timer
  EnteringOptions = 1
  OperatorMenuTimer.enabled=false
  ShowOperatorMenu
end sub

sub ShowOperatorMenu
  OperatorMenuBackdrop.image = "OperatorMenu"

  OptionCHS = 0
  CurrentOption = 1
  DisplayAllOptions
  OperatorOption1.image = "BluePlus"
  SetHighScoreOption

End Sub

Sub DisplayAllOptions
  dim linecounter
  dim tempstring
  For linecounter = 1 to MaxOption
    tempstring=Eval("OptionLine"&linecounter)
    Select Case linecounter
      Case 1:
        tempstring=tempstring + FormatNumber(BallsPerGame,0)
        SetOptLine 1,tempstring
      Case 2:
        if Replay3Table(ReplayLevel)=999000 then
          tempstring = tempstring + FormatNumber(Replay1Table(ReplayLevel),0) + "/" + FormatNumber(Replay2Table(ReplayLevel),0)
        else
          tempstring = tempstring + FormatNumber(Replay1Table(ReplayLevel),0) + "/" + FormatNumber(Replay2Table(ReplayLevel),0) + "/" + FormatNumber(Replay3Table(ReplayLevel),0)
        end if
        SetOptLine 2,tempstring
      Case 3:
        If OptionCHS=0 then
          tempstring = "NO"
        else
          tempstring = "YES"
        end if
        SetOptLine 3,tempstring
      Case 4:
        SetOptLine 4, tempstring
        Select Case SpecialLightFlag
          Case 1:
            tempstring = "1 of 4"
          Case 2:
            tempstring = "1 of 2"
          Case 3:
            tempstring = "Always Lit"
        end select
        SetOptLine 5, tempstring
      Case 5:
        SetOptLine 6, tempstring
        SetOptLine 7, tempstring

      Case 6:
        SetOptLine 8, tempstring
        SetOptLine 9, tempstring

      Case 7:
        SetOptLine 10, tempstring
        SetOptLine 11, tempstring

      Case 8:

      Case 9:


    End Select

  next
end sub

sub MoveArrow
  do
    CurrentOption = CurrentOption + 1
    If CurrentOption>Len(OptionLinesToMark) then
      CurrentOption=1
    end if
  loop until Mid(OptionLinesToMark,CurrentOption,1)="1"
end sub

sub CollectOptions(ByVal keycode)
  if Keycode = LeftFlipperKey then
    PlaySound "DropTargetDropped"
    For XOpt = 1 to MaxOption
      Eval("OperatorOption"&XOpt).image = "PostitBL"
    next
    MoveArrow
    if CurrentOption<8 then
      Eval("OperatorOption"&CurrentOption).image = "BluePlus"
    elseif CurrentOption=8 then
      Eval("OperatorOption"&CurrentOption).image = "GreenCheck"
    else
      Eval("OperatorOption"&CurrentOption).image = "RedX"
    end if

  elseif Keycode = RightFlipperKey then
    PlaySound "DropTargetDropped"
    if CurrentOption = 1 then
      If BallsPerGame = 3 then
        BallsPerGame = 5
      else
        BallsPerGame = 3
      end if
      DisplayAllOptions
    elseif CurrentOption = 2 then
      ReplayLevel=ReplayLevel+1
      If ReplayLevel>ReplayTableMax then
        ReplayLevel=1
      end if
      DisplayAllOptions
    elseif CurrentOption = 3 then
      if OptionCHS = 0 then
        OptionCHS = 1

      else
        OptionCHS = 0

      end if
      DisplayAllOptions
    elseif CurrentOption = 4 then
      SpecialLightFlag=SpecialLightFlag+1
      if SpecialLightFlag>3 then
        SpecialLightFlag=1
      end if
      DisplayAllOptions
    elseif CurrentOption = 8 or CurrentOption = 9 then
        if OptionCHS=1 then
          HSScore(1) = 75000
          HSScore(2) = 70000
          HSScore(3) = 60000
          HSScore(4) = 55000
          HSScore(5) = 50000

          HSName(1) = "AAA"
          HSName(2) = "ZZZ"
          HSName(3) = "XXX"
          HSName(4) = "ABC"
          HSName(5) = "BBB"
        end if

        if CurrentOption = 8 then
          savehs
        else
          loadhs
        end if
        OperatorMenuBackdrop.image = "PostitBL"
        For XOpt = 1 to MaxOption
          Eval("OperatorOption"&XOpt).image = "PostitBL"
        next

        For XOpt = 1 to 256
          Eval("Option"&XOpt).image = "PostItBL"
        next
        RefreshReplayCard
        InstructCard.image="IC"+FormatNumber(BallsPerGame,0)
        EnteringOptions = 0

    end if
  end if
End Sub

Sub SetHighScoreOption

End Sub

Function GetOptChar(String, Index)
  dim ThisChar
  dim FileName
  ThisChar = Mid(String, Index, 1)
  FileName = "PostIt"
  if ThisChar = " " or ThisChar = "" then
    FileName = FileName & "BL"
  elseif ThisChar = "<" then
    FileName = FileName & "LT"
  elseif ThisChar = "_" then
    FileName = FileName & "SP"
  elseif ThisChar = "/" then
    FileName = FileName & "SL"
  elseif ThisChar = "," then
    FileName = FileName & "CM"
  else
    FileName = FileName & ThisChar
  End If
  GetOptChar = FileName
End Function

dim LineLengths(22) ' maximum number of lines
Sub SetOptLine(LineNo, String)
  Dim DispLen
    Dim StrLen
  dim xfor
  dim Letter
  dim ThisDigit
  dim ThisChar
  dim LetterLine
  dim Index
  dim LetterName
  StrLen = len(string)
  Index = 1

  StrLen = len(String)
    DispLen = StrLen
    if (DispLen < LineLengths(LineNo)) Then
        DispLen = LineLengths(LineNo)
    end If

  for xfor = StartingArray(LineNo) to StartingArray(LineNo) + DispLen
    Eval("Option"&xfor).image = GetOptChar(string, Index)
    Index = Index + 1
  next
  LineLengths(LineNo) = StrLen

End Sub



'******************************************************
'              Fleep  BALL ROLLING AND DROP SOUNDS
'******************************************************
Const tnob = 11 ' total number of balls
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
                StopSound("fx_ballrolling" & b)
        'StopSound("BallRoll_" & b)
        Next


        ' exit the sub if no balls on the table
        If UBound(BOT) = -1 Then Exit Sub


        ' play the rolling sound for each ball
        For b = 0 to UBound(BOT)
                If BallVel(BOT(b)) > 1 AND BOT(b).z < 30 Then
                        rolling(b) = True
            PlaySound ("fx_ballrolling" & b), -1, VolPlayfieldRoll(BOT(b)) * 1.1 * VolumeDial, AudioPan(BOT(b)), 0, PitchPlayfieldRoll(BOT(b)), 1, 0, AudioFade(BOT(b))
                        'PlaySound ("BallRoll_" & b), -1, VolPlayfieldRoll(BOT(b)) * 1.1 * VolumeDial, AudioPan(BOT(b)), 0, PitchPlayfieldRoll(BOT(b)), 1, 0, AudioFade(BOT(b))


                Else
                        If rolling(b) = True Then
                                StopSound("fx_ballrolling" & b)
                'StopSound("BallRoll_" & b)
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

'*********************************************************************************
'   nFozzy Physice  - Flipper Correction Initialization late 70’s to early 80’s
'*********************************************************************************
dim LF : Set LF = New FlipperPolarity
dim RF : Set RF = New FlipperPolarity

InitPolarity

Sub InitPolarity()
        dim x, a : a = Array(LF, RF)
        for each x in a
                x.AddPoint "Ycoef", 0, RightFlipper.Y-65, 1        'disabled
                x.AddPoint "Ycoef", 1, RightFlipper.Y-11, 1
                x.enabled = True
                x.TimeDelay = 80
        Next

        AddPt "Polarity", 0, 0, 0
        AddPt "Polarity", 1, 0.05, -2.7
        AddPt "Polarity", 2, 0.33, -2.7
        AddPt "Polarity", 3, 0.37, -2.7
        AddPt "Polarity", 4, 0.41, -2.7
        AddPt "Polarity", 5, 0.45, -2.7
        AddPt "Polarity", 6, 0.576,-2.7
        AddPt "Polarity", 7, 0.66, -1.8
        AddPt "Polarity", 8, 0.743, -0.5
        AddPt "Polarity", 9, 0.81, -0.5
        AddPt "Polarity", 10, 0.88, 0

        addpt "Velocity", 0, 0,         1
        addpt "Velocity", 1, 0.16, 1.06
        addpt "Velocity", 2, 0.41,         1.05
        addpt "Velocity", 3, 0.53,         1'0.982
        addpt "Velocity", 4, 0.702, 0.968
        addpt "Velocity", 5, 0.95,  0.968
        addpt "Velocity", 6, 1.03,         0.945

        LF.Object = LeftFlipper
        LF.EndPoint = EndPointLp
        RF.Object = RightFlipper
        RF.EndPoint = EndPointRp
End Sub

Sub TriggerLF_Hit() : LF.Addball activeball : End Sub
Sub TriggerLF_UnHit() : LF.PolarityCorrect activeball : End Sub
Sub TriggerRF_Hit() : RF.Addball activeball : End Sub
Sub TriggerRF_UnHit() : RF.PolarityCorrect activeball : End Sub

'******************************************************
'                 nFozzy  FLIPPER CORRECTION FUNCTIONS
'******************************************************

Sub AddPt(aStr, idx, aX, aY)        'debugger wrapper for adjusting flipper script in-game
        dim a : a = Array(LF, RF)
        dim x : for each x in a
                x.addpoint aStr, idx, aX, aY
        Next
End Sub

Class FlipperPolarity
        Public DebugOn, Enabled
        Private FlipAt        'Timer variable (IE 'flip at 723,530ms...)
        Public TimeDelay        'delay before trigger turns off and polarity is disabled TODO set time!
        private Flipper, FlipperStart,FlipperEnd, FlipperEndY, LR, PartialFlipCoef
        Private Balls(20), balldata(20)

        dim PolarityIn, PolarityOut
        dim VelocityIn, VelocityOut
        dim YcoefIn, YcoefOut
        Public Sub Class_Initialize
                redim PolarityIn(0) : redim PolarityOut(0) : redim VelocityIn(0) : redim VelocityOut(0) : redim YcoefIn(0) : redim YcoefOut(0)
                Enabled = True : TimeDelay = 50 : LR = 1:  dim x : for x = 0 to uBound(balls) : balls(x) = Empty : set Balldata(x) = new SpoofBall : next
        End Sub

        Public Property let Object(aInput) : Set Flipper = aInput : StartPoint = Flipper.x : End Property
        Public Property Let StartPoint(aInput) : if IsObject(aInput) then FlipperStart = aInput.x else FlipperStart = aInput : end if : End Property
        Public Property Get StartPoint : StartPoint = FlipperStart : End Property
        Public Property Let EndPoint(aInput) : FlipperEnd = aInput.x: FlipperEndY = aInput.y: End Property
        Public Property Get EndPoint : EndPoint = FlipperEnd : End Property
        Public Property Get EndPointY: EndPointY = FlipperEndY : End Property

        Public Sub AddPoint(aChooseArray, aIDX, aX, aY) 'Index #, X position, (in) y Position (out)
                Select Case aChooseArray
                        case "Polarity" : ShuffleArrays PolarityIn, PolarityOut, 1 : PolarityIn(aIDX) = aX : PolarityOut(aIDX) = aY : ShuffleArrays PolarityIn, PolarityOut, 0
                        Case "Velocity" : ShuffleArrays VelocityIn, VelocityOut, 1 :VelocityIn(aIDX) = aX : VelocityOut(aIDX) = aY : ShuffleArrays VelocityIn, VelocityOut, 0
                        Case "Ycoef" : ShuffleArrays YcoefIn, YcoefOut, 1 :YcoefIn(aIDX) = aX : YcoefOut(aIDX) = aY : ShuffleArrays YcoefIn, YcoefOut, 0
                End Select
                if gametime > 100 then Report aChooseArray
        End Sub

        Public Sub Report(aChooseArray)         'debug, reports all coords in tbPL.text
                if not DebugOn then exit sub
                dim a1, a2 : Select Case aChooseArray
                        case "Polarity" : a1 = PolarityIn : a2 = PolarityOut
                        Case "Velocity" : a1 = VelocityIn : a2 = VelocityOut
                        Case "Ycoef" : a1 = YcoefIn : a2 = YcoefOut
                        case else :tbpl.text = "wrong string" : exit sub
                End Select
                dim str, x : for x = 0 to uBound(a1) : str = str & aChooseArray & " x: " & round(a1(x),4) & ", " & round(a2(x),4) & vbnewline : next
                tbpl.text = str
        End Sub

        Public Sub AddBall(aBall) : dim x : for x = 0 to uBound(balls) : if IsEmpty(balls(x)) then set balls(x) = aBall : exit sub :end if : Next  : End Sub

        Private Sub RemoveBall(aBall)
                dim x : for x = 0 to uBound(balls)
                        if TypeName(balls(x) ) = "IBall" then
                                if aBall.ID = Balls(x).ID Then
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
                dim x : for x = 0 to uBound(balls)
                        if not IsEmpty(balls(x) ) then
                                pos = pSlope(Balls(x).x, FlipperStart, 0, FlipperEnd, 1)
                        End If
                Next
        End Property

        Public Sub ProcessBalls() 'save data of balls in flipper range
                FlipAt = GameTime
                dim x : for x = 0 to uBound(balls)
                        if not IsEmpty(balls(x) ) then
                                balldata(x).Data = balls(x)
                        End If
                Next
                PartialFlipCoef = ((Flipper.StartAngle - Flipper.CurrentAngle) / (Flipper.StartAngle - Flipper.EndAngle))
                PartialFlipCoef = abs(PartialFlipCoef-1)
        End Sub
        Private Function FlipperOn() : if gameTime < FlipAt+TimeDelay then FlipperOn = True : End If : End Function        'Timer shutoff for polaritycorrect

        Public Sub PolarityCorrect(aBall)
                if FlipperOn() then
                        dim tmp, BallPos, x, IDX, Ycoef : Ycoef = 1

                        'y safety Exit
                        if aBall.VelY > -8 then 'ball going down
                                RemoveBall aBall
                                exit Sub
                        end if

                        'Find balldata. BallPos = % on Flipper
                        for x = 0 to uBound(Balls)
                                if aBall.id = BallData(x).id AND not isempty(BallData(x).id) then
                                        idx = x
                                        BallPos = PSlope(BallData(x).x, FlipperStart, 0, FlipperEnd, 1)
                                        if ballpos > 0.65 then  Ycoef = LinearEnvelope(BallData(x).Y, YcoefIn, YcoefOut)                                'find safety coefficient 'ycoef' data
                                end if
                        Next

                        If BallPos = 0 Then 'no ball data meaning the ball is entering and exiting pretty close to the same position, use current values.
                                BallPos = PSlope(aBall.x, FlipperStart, 0, FlipperEnd, 1)
                                if ballpos > 0.65 then  Ycoef = LinearEnvelope(aBall.Y, YcoefIn, YcoefOut)                                                'find safety coefficient 'ycoef' data
                        End If

                        'Velocity correction
                        if not IsEmpty(VelocityIn(0) ) then
                                Dim VelCoef
         :                         VelCoef = LinearEnvelope(BallPos, VelocityIn, VelocityOut)

                                if partialflipcoef < 1 then VelCoef = PSlope(partialflipcoef, 0, 1, 1, VelCoef)

                                if Enabled then aBall.Velx = aBall.Velx*VelCoef
                                if Enabled then aBall.Vely = aBall.Vely*VelCoef
                        End If

                        'Polarity Correction (optional now)
                        if not IsEmpty(PolarityIn(0) ) then
                                If StartPoint > EndPoint then LR = -1        'Reverse polarity if left flipper
                                dim AddX : AddX = LinearEnvelope(BallPos, PolarityIn, PolarityOut) * LR

                                if Enabled then aBall.VelX = aBall.VelX + 1 * (AddX*ycoef*PartialFlipcoef)
                                'playsound "fx_knocker"
                        End If
                End If
                RemoveBall aBall
        End Sub
End Class

'******************************************************
'                nFozzy FLIPPER POLARITY AND RUBBER DAMPENER
'                        SUPPORTING FUNCTIONS
'******************************************************

' Used for flipper correction and rubber dampeners
Sub ShuffleArray(ByRef aArray, byVal offset) 'shuffle 1d array
        dim x, aCount : aCount = 0
        redim a(uBound(aArray) )
        for x = 0 to uBound(aArray)        'Shuffle objects in a temp array
                if not IsEmpty(aArray(x) ) Then
                        if IsObject(aArray(x)) then
                                Set a(aCount) = aArray(x)
                        Else
                                a(aCount) = aArray(x)
                        End If
                        aCount = aCount + 1
                End If
        Next
        if offset < 0 then offset = 0
        redim aArray(aCount-1+offset)        'Resize original array
        for x = 0 to aCount-1                'set objects back into original array
                if IsObject(a(x)) then
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
    BallSpeed = SQR(ball.VelX^2 + ball.VelY^2 + ball.VelZ^2)
End Function

' Used for flipper correction and rubber dampeners
Function PSlope(Input, X1, Y1, X2, Y2)        'Set up line via two points, no clamping. Input X, output Y
        dim x, y, b, m : x = input : m = (Y2 - Y1) / (X2 - X1) : b = Y2 - m*X2
        Y = M*x+b
        PSlope = Y
End Function

' Used for flipper correction
Class spoofball
        Public X, Y, Z, VelX, VelY, VelZ, ID, Mass, Radius
        Public Property Let Data(aBall)
                With aBall
                        x = .x : y = .y : z = .z : velx = .velx : vely = .vely : velz = .velz
                        id = .ID : mass = .mass : radius = .radius
                end with
        End Property
        Public Sub Reset()
                x = Empty : y = Empty : z = Empty  : velx = Empty : vely = Empty : velz = Empty
                id = Empty : mass = Empty : radius = Empty
        End Sub
End Class

' Used for flipper correction and rubber dampeners
Function LinearEnvelope(xInput, xKeyFrame, yLvl)
        dim y 'Y output
        dim L 'Line
        dim ii : for ii = 1 to uBound(xKeyFrame)        'find active line
                if xInput <= xKeyFrame(ii) then L = ii : exit for : end if
        Next
        if xInput > xKeyFrame(uBound(xKeyFrame) ) then L = uBound(xKeyFrame)        'catch line overrun
        Y = pSlope(xInput, xKeyFrame(L-1), yLvl(L-1), xKeyFrame(L), yLvl(L) )

        if xInput <= xKeyFrame(lBound(xKeyFrame) ) then Y = yLvl(lBound(xKeyFrame) )         'Clamp lower
        if xInput >= xKeyFrame(uBound(xKeyFrame) ) then Y = yLvl(uBound(xKeyFrame) )        'Clamp upper

        LinearEnvelope = Y
End Function


'******************************************************
'                    nFozzy FLIPPER TRICKS
'******************************************************

RightFlipper.timerinterval=1
Rightflipper.timerenabled=True

sub RightFlipper_timer()
        FlipperTricks LeftFlipper, LFPress, LFCount, LFEndAngle, LFState
        FlipperTricks RightFlipper, RFPress, RFCount, RFEndAngle, RFState
        FlipperNudge RightFlipper, RFEndAngle, RFEOSNudge, LeftFlipper, LFEndAngle
        FlipperNudge LeftFlipper, LFEndAngle, LFEOSNudge,  RightFlipper, RFEndAngle
end sub

Dim LFEOSNudge, RFEOSNudge

Sub FlipperNudge(Flipper1, Endangle1, EOSNudge1, Flipper2, EndAngle2)
        Dim BOT, b

        If Flipper1.currentangle = Endangle1 and EOSNudge1 <> 1 Then
                EOSNudge1 = 1
                'debug.print Flipper1.currentangle &" = "& Endangle1 &"--"& Flipper2.currentangle &" = "& EndAngle2
                If Flipper2.currentangle = EndAngle2 Then
                        BOT = GetBalls
                        For b = 0 to Ubound(BOT)
                                If FlipperTrigger(BOT(b).x, BOT(b).y, Flipper1) Then
                                        'Debug.Print "ball in flip1. exit"
                                         exit Sub
                                end If
                        Next
                        For b = 0 to Ubound(BOT)
                                If FlipperTrigger(BOT(b).x, BOT(b).y, Flipper2) Then
                                        BOT(b).velx = BOT(b).velx / 1.7
                                        BOT(b).vely = BOT(b).vely - 1
                                end If
                        Next
                End If
        Else
                If Flipper1.currentangle <> EndAngle1 then
                        EOSNudge1 = 0
                end if
        End If
End Sub

'*****************
' nFozzy Maths
'*****************
Const PI = 3.1415927

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


'*************************************************
' nFozzy Check ball distance from Flipper for Rem
'*************************************************

Function Distance(ax,ay,bx,by)
        Distance = SQR((ax - bx)^2 + (ay - by)^2)
End Function

Function DistancePL(px,py,ax,ay,bx,by) ' Distance between a point and a line where point is px,py
        DistancePL = ABS((by - ay)*px - (bx - ax) * py + bx*ay - by*ax)/Distance(ax,ay,bx,by)
End Function

Function Radians(Degrees)
        Radians = Degrees * PI /180
End Function

Function AnglePP(ax,ay,bx,by)
        AnglePP = Atn2((by - ay),(bx - ax))*180/PI
End Function

Function DistanceFromFlipper(ballx, bally, Flipper)
        DistanceFromFlipper = DistancePL(ballx, bally, Flipper.x, Flipper.y, Cos(Radians(Flipper.currentangle+90))+Flipper.x, Sin(Radians(Flipper.currentangle+90))+Flipper.y)
End Function

Function FlipperTrigger(ballx, bally, Flipper)
        Dim DiffAngle
        DiffAngle  = ABS(Flipper.currentangle - AnglePP(Flipper.x, Flipper.y, ballx, bally) - 90)
        If DiffAngle > 180 Then DiffAngle = DiffAngle - 360

        If DistanceFromFlipper(ballx,bally,Flipper) < 48 and DiffAngle <= 90 and Distance(ballx,bally,Flipper.x,Flipper.y) < Flipper.Length Then
                FlipperTrigger = True
        Else
                FlipperTrigger = False
        End If
End Function


'*************************************************
' nFozzy End - Check ball distance from Flipper for Rem
'*************************************************

dim LFPress, RFPress, LFCount, RFCount
dim LFState, RFState
dim EOST, EOSA,Frampup, FElasticity,FReturn
dim RFEndAngle, LFEndAngle

EOST = leftflipper.eostorque
EOSA = leftflipper.eostorqueangle
Frampup = LeftFlipper.rampup
FElasticity = LeftFlipper.elasticity
FReturn = LeftFlipper.return
Const EOSTnew = 1
Const EOSAnew = 1
Const EOSRampup = 0
Dim SOSRampup

Select Case FlipperCoilRampupMode
  Case 0:
    SOSRampup = 2.5
  Case 1:
    SOSRampup = 6
  Case 2:
    SOSRampup = 8.5
End Select

SOSRampup = 2.5
Const LiveCatch = 16
Const LiveElasticity = 0.45
Const SOSEM = 0.815
Const EOSReturn = 0.055

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
        Flipper.eostorque = EOST*EOSReturn/FReturn


        If Abs(Flipper.currentangle) <= Abs(Flipper.endangle) + 0.1 Then
                Dim BOT, b
                BOT = GetBalls

                For b = 0 to UBound(BOT)
                        If Distance(BOT(b).x, BOT(b).y, Flipper.x, Flipper.y) < 55 Then 'check for cradle
                                If BOT(b).vely >= -0.4 Then BOT(b).vely = -0.4
                        End If
                Next
        End If
End Sub

Sub FlipperTricks (Flipper, FlipperPress, FCount, FEndAngle, FState)
        Dim Dir
        Dir = Flipper.startangle/Abs(Flipper.startangle)        '-1 for Right Flipper

        If Abs(Flipper.currentangle) > Abs(Flipper.startangle) - 0.05 Then
                If FState <> 1 Then
                        Flipper.rampup = SOSRampup
                        Flipper.endangle = FEndAngle - 3*Dir
                        Flipper.Elasticity = FElasticity * SOSEM
                        FCount = 0
                        FState = 1
                End If
        ElseIf Abs(Flipper.currentangle) <= Abs(Flipper.endangle) and FlipperPress = 1 then
                if FCount = 0 Then FCount = GameTime

                If FState <> 2 Then
                        Flipper.eostorqueangle = EOSAnew
                        Flipper.eostorque = EOSTnew
                        Flipper.rampup = EOSRampup
                        Flipper.endangle = FEndAngle
                        FState = 2
                End If
        Elseif Abs(Flipper.currentangle) > Abs(Flipper.endangle) + 0.01 and FlipperPress = 1 Then
                If FState <> 3 Then
                        Flipper.eostorque = EOST
                        Flipper.eostorqueangle = EOSA
                        Flipper.rampup = Frampup
                        Flipper.Elasticity = FElasticity
                        FState = 3
                End If

        End If
End Sub

Const LiveDistanceMin = 30  'minimum distance in vp units from flipper base live catch dampening will occur
Const LiveDistanceMax = 114  'maximum distance in vp units from flipper base live catch dampening will occur (tip protection)

Sub CheckLiveCatch(ball, Flipper, FCount, parm) 'Experimental new live catch
        Dim Dir
        Dir = Flipper.startangle/Abs(Flipper.startangle)    '-1 for Right Flipper
        Dim LiveCatchBounce                                                                                                                        'If live catch is not perfect, it won't freeze ball totally
        Dim CatchTime : CatchTime = GameTime - FCount

        if CatchTime <= LiveCatch and parm > 6 and ABS(Flipper.x - ball.x) > LiveDistanceMin and ABS(Flipper.x - ball.x) < LiveDistanceMax Then
                if CatchTime <= LiveCatch*0.5 Then                                                'Perfect catch only when catch time happens in the beginning of the window
                        LiveCatchBounce = 0
                else
                        LiveCatchBounce = Abs((LiveCatch/2) - CatchTime)        'Partial catch when catch happens a bit late
                end If

                If LiveCatchBounce = 0 and ball.velx * Dir > 0 Then ball.velx = 0
                ball.vely = LiveCatchBounce * (32 / LiveCatch) ' Multiplier for inaccuracy bounce
                ball.angmomx= 0
                ball.angmomy= 0
                ball.angmomz= 0
        End If
End Sub


'**************************************************
'       nFozzy Flipper Collision Subs
'**************************************************'

Sub LeftFlipper_Collide(parm)
    CheckLiveCatch Activeball, LeftFlipper, LFCount, parm
    LeftFlipperCollide parm   'This is the Fleep code
  if RubberizerEnabled <> 0 then Rubberizer(parm)
End Sub

Sub RightFlipper_Collide(parm)
    CheckLiveCatch Activeball, RightFlipper, RFCount, parm
    RightFlipperCollide parm  'This is the Fleep code
  if RubberizerEnabled <> 0 then Rubberizer(parm)
End Sub


'**************************************************
'     iaakki Rubberizer
'**************************************************'
sub Rubberizer(parm)
  if parm < 10 And parm > 2 And Abs(activeball.angmomz) < 10 then
    'debug.print "parm: " & parm & " momz: " & activeball.angmomz &" vely: "& activeball.vely
    activeball.angmomz = activeball.angmomz * 1.2
    activeball.vely = activeball.vely * 1.2
    'debug.print ">> newmomz: " & activeball.angmomz&" newvely: "& activeball.vely
  Elseif parm <= 2 and parm > 0.2 Then
    'debug.print "* parm: " & parm & " momz: " & activeball.angmomz &" vely: "& activeball.vely
    activeball.angmomz = activeball.angmomz * -1.1
    activeball.vely = activeball.vely * 1.4
    'debug.print "**** >> newmomz: " & activeball.angmomz&" newvely: "& activeball.vely
  end if
end sub


'****************************************************************************
' nFozzy PHYSICS DAMPENERS
'****************************************************************************

'These are data mined bounce curves,
'dialed in with the in-game elasticity as much as possible to prevent angle / spin issues.
'Requires tracking ballspeed to calculate COR


Sub dPosts_Hit(idx)
        RubbersD.dampen Activeball
End Sub

Sub dSleeves_Hit(idx)
        SleevesD.Dampen Activeball
End Sub


dim RubbersD : Set RubbersD = new Dampener        'frubber
RubbersD.name = "Rubbers"
RubbersD.debugOn = False        'shows info in textbox "TBPout"
RubbersD.Print = False        'debug, reports in debugger (in vel, out cor)
'cor bounce curve (linear)
'for best results, try to match in-game velocity as closely as possible to the desired curve
'RubbersD.addpoint 0, 0, 0.935        'point# (keep sequential), ballspeed, CoR (elasticity)
RubbersD.addpoint 0, 0, 0.96        'point# (keep sequential), ballspeed, CoR (elasticity)
RubbersD.addpoint 1, 3.77, 0.96
RubbersD.addpoint 2, 5.76, 0.967        'dont take this as gospel. if you can data mine rubber elasticitiy, please help!
RubbersD.addpoint 3, 15.84, 0.874
RubbersD.addpoint 4, 56, 0.64        'there's clamping so interpolate up to 56 at least

dim SleevesD : Set SleevesD = new Dampener        'this is just rubber but cut down to 85%...
SleevesD.name = "Sleeves"
SleevesD.debugOn = False        'shows info in textbox "TBPout"
SleevesD.Print = False        'debug, reports in debugger (in vel, out cor)
SleevesD.CopyCoef RubbersD, 0.85

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
                RealCOR = BallSpeed(aBall) / cor.ballvel(aBall.id)
                coef = desiredcor / realcor
                if debugOn then str = name & " in vel:" & round(cor.ballvel(aBall.id),2 ) & vbnewline & "desired cor: " & round(desiredcor,4) & vbnewline & _
                "actual cor: " & round(realCOR,4) & vbnewline & "ballspeed coef: " & round(coef, 3) & vbnewline
                if Print then debug.print Round(cor.ballvel(aBall.id),2) & ", " & round(desiredcor,3)

' Thalamus - patched : ' Thalamus - patched :                 aBall.velx = aBall.velx * coef : aBall.vely = aBall.vely * coef : aBall.velz = aBall.velz * coef
    aBall.velx = aBall.velx * coef : aBall.vely = aBall.vely * coef : aBall.velz = aBall.velz * coef
' Thalamus - patched :     aBall.velx = aBall.velx * coef : aBall.vely = aBall.vely * coef : aBall.velz = aBall.velz * coef
    aBall.velx = aBall.velx * coef : aBall.vely = aBall.vely * coef : aBall.velz = aBall.velz * coef
                if debugOn then TBPout.text = str
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
'iaakki - TargetBouncer for targets and posts
'******************************************************
Dim zMultiplier

sub TargetBouncer(aBall,defvalue)
  if TargetBouncerEnabled <> 0 and aball.z < 30 then
    'debug.print "velz: " & activeball.velz
    Select Case Int(Rnd * 4) + 1
      Case 1: zMultiplier = defvalue+1.1
      Case 2: zMultiplier = defvalue+1.05
      Case 3: zMultiplier = defvalue+0.7
      Case 4: zMultiplier = defvalue+0.3
    End Select
    aBall.velz = aBall.velz * zMultiplier * TargetBouncerFactor
    'debug.print "----> velz: " & activeball.velz
  end if
end sub


'******************************************************
'              nFozzy  TRACK ALL BALL VELOCITIES
'                 FOR RUBBER DAMPENER AND DROP TARGETS
'******************************************************

dim cor : set cor = New CoRTracker

Class CoRTracker
        public ballvel, ballvelx, ballvely

        Private Sub Class_Initialize : redim ballvel(0) : redim ballvelx(0): redim ballvely(0) : End Sub

        Public Sub Update()        'tracks in-ball-velocity
                dim str, b, AllBalls, highestID : allBalls = getballs

                for each b in allballs
                        if b.id >= HighestID then highestID = b.id
                Next

                if uBound(ballvel) < highestID then redim ballvel(highestID)        'set bounds
                if uBound(ballvelx) < highestID then redim ballvelx(highestID)        'set bounds
                if uBound(ballvely) < highestID then redim ballvely(highestID)        'set bounds

                for each b in allballs
                        ballvel(b.id) = BallSpeed(b)
                        ballvelx(b.id) = b.velx
                        ballvely(b.id) = b.vely
                Next
        End Sub
End Class


Sub RDampen_Timer()
Cor.Update
End Sub

'****************************************************************************
'////////////////////////////  Fleep MECHANICAL SOUNDS  ///////////////////////////
'//  This part in the script is an entire block that is dedicated to the physics sound system.
'//  Various scripts and sounds that may be pretty generic and could suit other WPC systems, but the most are tailored specifically for this table.
'****************************************************************************

'///////////////////////////////  SOUNDS PARAMETERS  //////////////////////////////
Dim GlobalSoundLevel, CoinSoundLevel, PlungerReleaseSoundLevel, PlungerPullSoundLevel, NudgeLeftSoundLevel
Dim NudgeRightSoundLevel, NudgeCenterSoundLevel, StartButtonSoundLevel, RollingSoundFactor


CoinSoundLevel = 1                                                                                                                'volume level; range [0, 1]
NudgeLeftSoundLevel = 1                                                                                                        'volume level; range [0, 1]
NudgeRightSoundLevel = 1                                                                                                'volume level; range [0, 1]
NudgeCenterSoundLevel = 1                                                                                                'volume level; range [0, 1]
StartButtonSoundLevel = 0.1                                                                                                'volume level; range [0, 1]
PlungerReleaseSoundLevel = 0.8 '1 wjr                                                                                        'volume level; range [0, 1]
PlungerPullSoundLevel = 1                                                                                                'volume level; range [0, 1]
RollingSoundFactor = 1.1/5


'///////////////////////-----Solenoids, Kickers and Flash Relays-----///////////////////////
Dim FlipperUpAttackMinimumSoundLevel, FlipperUpAttackMaximumSoundLevel, FlipperUpAttackLeftSoundLevel, FlipperUpAttackRightSoundLevel
Dim FlipperUpSoundLevel, FlipperDownSoundLevel, FlipperLeftHitParm, FlipperRightHitParm
Dim SlingshotSoundLevel, BumperSoundFactor, KnockerSoundLevel


FlipperUpAttackMinimumSoundLevel = 0.010                                                           'volume level; range [0, 1]
FlipperUpAttackMaximumSoundLevel = 0.635                                                                'volume level; range [0, 1]
FlipperUpSoundLevel = 1.0                                                                        'volume level; range [0, 1]
FlipperDownSoundLevel = 0.45                                                                      'volume level; range [0, 1]
FlipperLeftHitParm = FlipperUpSoundLevel                                                                'sound helper; not configurable
FlipperRightHitParm = FlipperUpSoundLevel                                                                'sound helper; not configurable
SlingshotSoundLevel = 0.95                                                                                                'volume level; range [0, 1]
BumperSoundFactor = 4.25                                                                                                'volume multiplier; must not be zero
KnockerSoundLevel = 1                                                                                                         'volume level; range [0, 1]


'///////////////////////-----Ball Drops, Bumps and Collisions-----///////////////////////
Dim RubberStrongSoundFactor, RubberWeakSoundFactor, RubberFlipperSoundFactor,BallWithBallCollisionSoundFactor
Dim BallBouncePlayfieldSoftFactor, BallBouncePlayfieldHardFactor, PlasticRampDropToPlayfieldSoundLevel, WireRampDropToPlayfieldSoundLevel, DelayedBallDropOnPlayfieldSoundLevel
Dim WallImpactSoundFactor, MetalImpactSoundFactor, SubwaySoundLevel, SubwayEntrySoundLevel, ScoopEntrySoundLevel
Dim SaucerLockSoundLevel, SaucerKickSoundLevel


BallWithBallCollisionSoundFactor = 3.2                                                                        'volume multiplier; must not be zero
RubberStrongSoundFactor = 0.055/5                                                                                        'volume multiplier; must not be zero
RubberWeakSoundFactor = 0.075/5                                                                                        'volume multiplier; must not be zero
RubberFlipperSoundFactor = 0.075/5                                                                                'volume multiplier; must not be zero
BallBouncePlayfieldSoftFactor = 0.025                                                                        'volume multiplier; must not be zero
BallBouncePlayfieldHardFactor = 0.025                                                                        'volume multiplier; must not be zero
DelayedBallDropOnPlayfieldSoundLevel = 0.8                                                                        'volume level; range [0, 1]
WallImpactSoundFactor = 0.075                                                                                        'volume multiplier; must not be zero
MetalImpactSoundFactor = 0.075/3
SaucerLockSoundLevel = 0.8
SaucerKickSoundLevel = 0.8


'///////////////////////-----Gates, Spinners, Rollovers and Targets-----///////////////////////

Dim GateSoundLevel, TargetSoundFactor, SpinnerSoundLevel, RolloverSoundLevel, DTSoundLevel


GateSoundLevel = 0.5/5                                                                                                        'volume level; range [0, 1]
TargetSoundFactor = 0.0025 * 10                                                                                        'volume multiplier; must not be zero
DTSoundLevel = 0.25                                                                                                                'volume multiplier; must not be zero
RolloverSoundLevel = 0.25                                                                      'volume level; range [0, 1]


'///////////////////////-----Ball Release, Guides and Drain-----///////////////////////
Dim DrainSoundLevel, BallReleaseSoundLevel, BottomArchBallGuideSoundFactor, FlipperBallGuideSoundFactor


DrainSoundLevel = 0.8                                                                                                                'volume level; range [0, 1]
BallReleaseSoundLevel = 1                                                                                                'volume level; range [0, 1]
BottomArchBallGuideSoundFactor = 0.2                                                                        'volume multiplier; must not be zero
FlipperBallGuideSoundFactor = 0.015                                                                                'volume multiplier; must not be zero


'///////////////////////-----Loops and Lanes-----///////////////////////
Dim ArchSoundFactor
ArchSoundFactor = 0.025/5                                                                                                        'volume multiplier; must not be zero



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



' *********************************************************************
'                     Fleep Supporting Ball & Sound Functions
' *********************************************************************

Dim tablewidth, tableheight : tablewidth = table1.width : tableheight = table1.height

Function AudioFade(tableobj) ' Fades between front and back of the table (for surround systems or 2x2 speakers, etc), depending on the Y position on the table. "table1" is the name of the table
        Dim tmp
    tmp = tableobj.y * 2 / tableheight-1
    If tmp > 0 Then
                AudioFade = Csng(tmp ^10)
    Else
        AudioFade = Csng(-((- tmp) ^10) )
    End If
End Function


Function AudioPan(tableobj) ' Calculates the pan for a tableobj based on the X position on the table. "table1" is the name of the table
    Dim tmp
    tmp = tableobj.x * 2 / tablewidth-1
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


'/////////////////////////////  FLIPPER BATS SOUND SUBROUTINES  ////////////////////////////
'/////////////////////////////  FLIPPER BATS SOLENOID ATTACK SOUND  ////////////////////////////
Sub SoundFlipperUpAttackLeft(flipper)
        FlipperUpAttackLeftSoundLevel = RndNum(FlipperUpAttackMinimumSoundLevel, FlipperUpAttackMaximumSoundLevel)
        PlaySoundAtLevelStatic ("Flipper_Attack-L01"), FlipperUpAttackLeftSoundLevel, flipper
End Sub


Sub SoundFlipperUpAttackRight(flipper)
        FlipperUpAttackRightSoundLevel = RndNum(FlipperUpAttackMinimumSoundLevel, FlipperUpAttackMaximumSoundLevel)
                PlaySoundAtLevelStatic ("Flipper_Attack-R01"), FlipperUpAttackLeftSoundLevel, flipper
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
         dim finalspeed
          finalspeed=SQR(activeball.velx * activeball.velx + activeball.vely * activeball.vely)
         If finalspeed > 5 then
                 RandomSoundRubberStrong 1
        End if
        If finalspeed <= 5 then
                 RandomSoundRubberWeak()
         End If
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


'/////////////////////////////  BALL COLLISION SOUND  ////////////////////////////  - Removed because its a single ball game and caused issues with VR BG Balls
'Sub OnBallBallCollision(ball1, ball2, velocity)
 '       Dim snd
  '      Select Case Int(Rnd*7)+1
   '             Case 1 : snd = "Ball_Collide_1"
    '            Case 2 : snd = "Ball_Collide_2"
     '           Case 3 : snd = "Ball_Collide_3"
      '          Case 4 : snd = "Ball_Collide_4"
       '         Case 5 : snd = "Ball_Collide_5"
        '        Case 6 : snd = "Ball_Collide_6"
         '       Case 7 : snd = "Ball_Collide_7"
        'End Select


'        PlaySound (snd), 0, Csng(velocity) ^2 / 200 * BallWithBallCollisionSoundFactor * VolumeDial, AudioPan(ball1), 0, Pitch(ball1), 0, 0, AudioFade(ball1)
'End Sub

'****************************************************************************
'                                        End Fleep Mechanical Sounds
'****************************************************************************



'****************************************************************************
'                                        VR Supporting Methods
'****************************************************************************
Sub VRRoomMinimal()
  for each Object in ColVRRoomMinimal : object.visible = 1 : next
  'for each Object in ColVRRoom1 : object.visible = 0 : next
  Ramp17.visible = 0   'desktop lockdown bar made invisible
End Sub

Sub VRSelectRoom()
  If VRRoom="" Then
    VRRoom = 0
  End if

  If VRRoom = 0 Then 'No VR Room
  ' Make VR EM Reels invisible if VRRoom = 0
    For each obj in VREMReels: obj.visible = false: next
  ' sets the sign in the backglass to tell users to run VR room in script
    SetVRRoomSign. visible = true
  for each Object in ColVRRoom1 : object.visible = 0 : next
  End If

  If VRRoom = 1 Then
    VRRoomMinimal
  End If
End Sub

'*****************************************************************************************************
' VR PLUNGER ANIMATION
'
' Code needed to animate the plunger. If you pull the plunger it will move in VR.
' IMPORTANT: there are two numeric values in the code that define the postion of the plunger and the
' range in which it can move. The first numeric value is the actual y position of the plunger primitive
' and the second is the actual y position + 135 to determine the range in which it can move.
'
' You need to to select the Primary_plunger primitive you copied from the
' template and copy the value of the Y position
' (e.g. 2114) into the code. The value that determines the range of the plunger is always the y
' position + 135 (e.g. 2249).
'
'*****************************************************************************************************

Sub TimerPlunger_Timer
if VRRoom = 1 then
  If VR_Primary_plunger.Y < 2216 then'plunger Y value in editor +135
      VR_Primary_plunger.Y = VR_Primary_plunger.Y + 5
end if
  End If
End Sub

Sub TimerPlunger2_Timer
 'debug.print plunger.position
if VRRoom = 1 then
  VR_Primary_plunger.Y = 2081 + (5* Plunger.Position) -20 'plunger Y value in editor
end if
End Sub



' VR Backglaass flasher and EM reel code below ****
'**************************************************

Sub FlasherMatch

  If Match = 0 Then FlM00A.visible = 1 : FlM00B.visible = 1 Else FlM00A.visible = 0 : FlM00B.visible = 0  End If
  If Match = 10 Then FlM10A.visible = 1 : FlM10B.visible = 1 Else FlM10A.visible = 0 : FlM10B.visible = 0 End If
  If Match = 20 Then FlM20A.visible = 1 : FlM20B.visible = 1 Else FlM20A.visible = 0 : FlM20B.visible = 0 End If
  If Match = 30 Then FlM30A.visible = 1 : FlM30B.visible = 1 Else FlM30A.visible = 0 : FlM30B.visible = 0 End If
  If Match = 40 Then FlM40A.visible = 1 : FlM40B.visible = 1 Else FlM40A.visible = 0 : FlM40B.visible = 0 End If
  If Match = 50 Then FlM50A.visible = 1 : FlM50B.visible = 1 Else FlM50A.visible = 0 : FlM50B.visible = 0 End If
  If Match = 60 Then FlM60A.visible = 1 : FlM60B.visible = 1 Else FlM60A.visible = 0 : FlM60B.visible = 0 End If
  If Match = 70 Then FlM70A.visible = 1 : FlM70B.visible = 1 Else FlM70A.visible = 0 : FlM70B.visible = 0 End If
  If Match = 80 Then FlM80A.visible = 1 : FlM80B.visible = 1 Else FlM80A.visible = 0 : FlM80B.visible = 0 End If
  If Match = 90 Then FlM90A.visible = 1 : FlM90B.visible = 1 Else FlM90A.visible = 0 : FlM90B.visible = 0 End If
End Sub

Sub FlasherPlayers

If Players = 1 Then FlPLR1.visible = 1 Else FlPLR1.visible = 0  End If
If Players = 2 Then FlPLR2.visible = 1 Else FlPLR2.visible = 0  End If
If Players = 3 Then FlPLR3.visible = 1 Else FlPLR3.visible = 0  End If
If Players = 4 Then FlPLR4.visible = 1 Else FlPLR4.visible = 0  End If
End Sub

Sub FlasherCurrentPlayer

If Player = 1 Then : for each Object in ColFlPlayer1 : object.visible = 1 : Next: Else : for each Object in ColFlPlayer1 : object.visible = 0 :Next
If Player = 2 Then : for each Object in ColFlPlayer2 : object.visible = 1 : Next: Else : for each Object in ColFlPlayer2 : object.visible = 0 :Next
If Player = 3 Then : for each Object in ColFlPlayer3 : object.visible = 1 : Next: Else : for each Object in ColFlPlayer3 : object.visible = 0 :Next
If Player = 4 Then : for each Object in ColFlPlayer4 : object.visible = 1 : Next: Else : for each Object in ColFlPlayer4 : object.visible = 0 :Next
End Sub

Sub FlasherBalls

If BallInPlay = 1 Then FlBIP1A.visible = 1 Else  FlBIP1A.visible = 0 End If
If BallInPlay = 2 Then FlBIP2A.visible = 1 Else  FlBIP2A.visible = 0 End If
If BallInPlay = 3 Then FlBIP3A.visible = 1 Else FlBIP3A.visible = 0 End If
If BallInPlay = 4 Then FlBIP4A.visible = 1 Else FlBIP4A.visible = 0 End If
If BallInPlay = 5 Then FlBIP5A.visible = 1 Else FlBIP5A.visible = 0 End If
End Sub


'******************************************************
'*******  Set Up Backglass Flashers *******
'******************************************************

Sub SetBackglass()

if VRRoom = 1 then

Dim obj
  For Each obj In Backglass_Flashers_Top
    obj.x = obj.x
    obj.height = - obj.y + 40
    obj.y = -10 'adjusts the distance from the backglass towards the user
  Next

  For Each obj In Backglass_Flashers_MidTop
    obj.x = obj.x
    obj.height = - obj.y + 40
    obj.y = 0 'adjusts the distance from the backglass towards the user
  Next

  For Each obj In Backglass_Flashers_Mid
    obj.x = obj.x
    obj.height = - obj.y + 40
    obj.y = 10 'adjusts the distance from the backglass towards the user
  Next

  For Each obj In Backglass_Flashers_Bottom
    obj.x = obj.x
    obj.height = - obj.y + 40
    obj.y = 17 'adjusts the distance from the backglass towards the user
  Next
end if

End Sub

'**********************************************
'*********************************
' ***************************************************************************
'          BASIC FSS(EM) 1-4 player 5x drums, 1 credit drum CORE CODE
' ****************************************************************************
' ********************* POSITION EM REEL DRUMS ON BACKGLASS *************************

Dim xoff,yoff,zoff,xrot,zscale, xcen,ycen
Dim inx
xoff =475
yoff = 0
zoff =735
xrot = -90

Const USEEMS = 4 ' 1-4 set between 1 to 4 based on number of players

const idx_emp1r1 =0 'player 1
const idx_emp2r1 =5 'player 2
const idx_emp3r1 =10 'player 3
const idx_emp4r1 =15 'player 4
const idx_emp4r6 =20 'credits


Dim BGObjEM(1)
if USEEMS = 1 then
  BGObjEM(0) = Array(emp1r1, emp1r2, emp1r3, emp1r4, emp1r5, _
  Empty,Empty,Empty,Empty,Empty,_
  Empty,Empty,Empty,Empty,Empty,_
  Empty,Empty,Empty,Empty,Empty,_
  emp4r6) ' credits
elseif USEEMS = 2 then
  BGObjEM(0) = Array(emp1r1, emp1r2, emp1r3, emp1r4, emp1r5, _
  emp2r1, emp2r2, emp2r3, emp2r4, emp2r5, _
  Empty,Empty,Empty,Empty,Empty,_
  Empty,Empty,Empty,Empty,Empty,_
  emp4r6) ' credits
elseif USEEMS = 3 then
  BGObjEM(0) = Array(emp1r1, emp1r2, emp1r3, emp1r4, emp1r5, _
  emp2r1, emp2r2, emp2r3, emp2r4, emp2r5, _
  emp3r1, emp3r2, emp3r3, emp3r4, emp3r5, _
  Empty,Empty,Empty,Empty,Empty,_
  emp4r6) ' credits
elseif USEEMS = 4 then
  BGObjEM(0) = Array(emp1r1, emp1r2, emp1r3, emp1r4, emp1r5, _
  emp2r1, emp2r2, emp2r3, emp2r4, emp2r5, _
  emp3r1, emp3r2, emp3r3, emp3r4, emp3r5, _
  emp4r1, emp4r2, emp4r3, emp4r4, emp4r5, _
  emp4r6) ' credits
end If

Sub center_objects_em()
Dim cnt,ii, xx, yy, yfact, xfact, objs
'exit sub
yoff = -150
zscale = 0.0000001
xcen =(960 /2) - (17 / 2)
ycen = (1065 /2 ) + (313 /2)

yfact = -25
xfact = -5

cnt =0
  For Each objs In BGObjEM(0)
  If Not IsEmpty(objs) then
    if objs.name = emp4r6.name then
    yoff = 50 ' credit drum is 60% smaller
    Else
    yoff = -110
    end if

  xx =objs.x

  objs.x = (xoff - xcen) + xx + xfact
  yy = objs.y
  objs.y =yoff

    If yy < 0 then
    yy = yy * -1
    end if

  objs.z = (zoff - ycen) + yy - (yy * zscale) + yfact

  'objs.rotx = xrot
  end if
  cnt = cnt + 1
  Next

end sub


' ********************* UPDATE EM REEL DRUMS CORE LIB *************************

Dim cred,ix, np,npp, reels(5, 7), scores(6,2)

'reset scores to defaults
for np =0 to 5
scores(np,0 ) = 0
scores(np,1 ) = 0
Next

'reset EM drums to defaults
For np =0 to 3
  For  npp =0 to 6
  reels(np, npp) =0 ' default to zero
  Next
Next

Sub SetScore(player, ndx , val)

Dim ncnt

  if player = 5 or player = 6 then
    if val > 0 then
      If(ndx = 0)Then ncnt = val * 10
      If(ndx = 1)Then ncnt = val

      scores(player, 0) = scores(player, 0) + ncnt
    end if
  else
    if val > 0 then

    If(ndx = 0)then ncnt = val * 10000
    If(ndx = 1)then ncnt = val * 1000
    If(ndx = 2)Then ncnt = val * 100
    If(ndx = 3)Then ncnt = val * 10
    If(ndx = 4)Then ncnt = val

    scores(player, 0) = scores(player, 0) + ncnt
    'scores(player, 0) + ncnt

    end if
  end if
End Sub


Sub SetDrum(player, drum , val)
Dim cnt
Dim objs : objs =BGObjEM(0)

  If val = 0 then
    Select case player
    case -1: ' the credit drum
    If Not IsEmpty(objs(idx_emp4r6)) then
    objs(idx_emp4r6).ObjrotX = 0 ' 285
    'cnt =objs(idx_emp4r6).ObjrotX
    end if
    Case 0:
    Select Case drum
        Case 1: If Not IsEmpty(objs(idx_emp1r1)) then: objs(idx_emp1r1).ObjrotX = 0: end if' 283
        Case 2: If Not IsEmpty(objs(idx_emp1r1+1)) then: objs(idx_emp1r1+1).ObjrotX=0: end if
        Case 3: If Not IsEmpty(objs(idx_emp1r1+2)) then: objs(idx_emp1r1+2).ObjrotX=0: end if
        Case 4: If Not IsEmpty(objs(idx_emp1r1+3)) then: objs(idx_emp1r1+3).ObjrotX=0: end if
        Case 5: If Not IsEmpty(objs(idx_emp1r1+4)) then: objs(idx_emp1r1+4).ObjrotX=0: end if
    End Select
    Case 1:
    Select Case drum
        Case 1: If Not IsEmpty(objs(idx_emp2r1)) then: objs(idx_emp2r1).ObjrotX = 0: end if' 283
        Case 2: If Not IsEmpty(objs(idx_emp2r1+1)) then: objs(idx_emp2r1+1).ObjrotX=0: end if
        Case 3: If Not IsEmpty(objs(idx_emp2r1+2)) then: objs(idx_emp2r1+2).ObjrotX=0: end if
        Case 4: If Not IsEmpty(objs(idx_emp2r1+3)) then: objs(idx_emp2r1+3).ObjrotX=0: end if
        Case 5: If Not IsEmpty(objs(idx_emp2r1+4)) then: objs(idx_emp2r1+4).ObjrotX=0: end if
    End Select
    Case 2:
    Select Case drum
        Case 1: If Not IsEmpty(objs(idx_emp3r1)) then: objs(idx_emp3r1).ObjrotX = 0: end if' 283
        Case 2: If Not IsEmpty(objs(idx_emp3r1+1)) then: objs(idx_emp3r1+1).ObjrotX=0: end if
        Case 3: If Not IsEmpty(objs(idx_emp3r1+2)) then: objs(idx_emp3r1+2).ObjrotX=0: end if
        Case 4: If Not IsEmpty(objs(idx_emp3r1+3)) then: objs(idx_emp3r1+3).ObjrotX=0: end if
        Case 5: If Not IsEmpty(objs(idx_emp3r1+4)) then: objs(idx_emp3r1+4).ObjrotX=0: end if
    End Select
    Case 3:
    Select Case drum
        Case 1: If Not IsEmpty(objs(idx_emp4r1)) then: objs(idx_emp4r1).ObjrotX = 0: end if' 283
        Case 2: If Not IsEmpty(objs(idx_emp4r1+1)) then: objs(idx_emp4r1+1).ObjrotX=0: end if
        Case 3: If Not IsEmpty(objs(idx_emp4r1+2)) then: objs(idx_emp4r1+2).ObjrotX=0: end if
        Case 4: If Not IsEmpty(objs(idx_emp4r1+3)) then: objs(idx_emp4r1+3).ObjrotX=0: end if
        Case 5: If Not IsEmpty(objs(idx_emp4r1+4)) then: objs(idx_emp4r1+4).ObjrotX=0: end if
    End Select
  End Select

  else
  Select case player

    Case -1: ' the credit drum
    'emp4r6.ObjrotX = emp4r6.ObjrotX + val
    If Not IsEmpty(objs(idx_emp4r6)) then
    objs(idx_emp4r6).ObjrotX = objs(idx_emp4r6).ObjrotX + val
    end if

    Case 0:
    Select Case drum
        Case 1: If Not IsEmpty(objs(idx_emp1r1)) then: objs(idx_emp1r1).ObjrotX= objs(idx_emp1r1).ObjrotX + val: end if
        Case 2: If Not IsEmpty(objs(idx_emp1r1+1)) then: objs(idx_emp1r1+1).ObjrotX= objs(idx_emp1r1+1).ObjrotX + val: end if
        Case 3: If Not IsEmpty(objs(idx_emp1r1+2)) then: objs(idx_emp1r1+2).ObjrotX= objs(idx_emp1r1+2).ObjrotX + val: end if
        Case 4: If Not IsEmpty(objs(idx_emp1r1+3)) then: objs(idx_emp1r1+3).ObjrotX= objs(idx_emp1r1+3).ObjrotX + val: end if
        Case 5: If Not IsEmpty(objs(idx_emp1r1+4)) then: objs(idx_emp1r1+4).ObjrotX= objs(idx_emp1r1+4).ObjrotX + val: end if
    End Select
    Case 1:
    Select Case drum
        Case 1: If Not IsEmpty(objs(idx_emp2r1)) then: objs(idx_emp2r1).ObjrotX= objs(idx_emp2r1).ObjrotX + val: end if
        Case 2: If Not IsEmpty(objs(idx_emp2r1+1)) then: objs(idx_emp2r1+1).ObjrotX= objs(idx_emp2r1+1).ObjrotX + val: end if
        Case 3: If Not IsEmpty(objs(idx_emp2r1+2)) then: objs(idx_emp2r1+2).ObjrotX= objs(idx_emp2r1+2).ObjrotX + val: end if
        Case 4: If Not IsEmpty(objs(idx_emp2r1+3)) then: objs(idx_emp2r1+3).ObjrotX= objs(idx_emp2r1+3).ObjrotX + val: end if
        Case 5: If Not IsEmpty(objs(idx_emp2r1+4)) then: objs(idx_emp2r1+4).ObjrotX= objs(idx_emp2r1+4).ObjrotX + val: end if
    End Select
    Case 2:
    Select Case drum
        Case 1: If Not IsEmpty(objs(idx_emp3r1)) then: objs(idx_emp3r1).ObjrotX= objs(idx_emp3r1).ObjrotX + val: end if
        Case 2: If Not IsEmpty(objs(idx_emp3r1+1)) then: objs(idx_emp3r1+1).ObjrotX= objs(idx_emp3r1+1).ObjrotX + val: end if
        Case 3: If Not IsEmpty(objs(idx_emp3r1+2)) then: objs(idx_emp3r1+2).ObjrotX= objs(idx_emp3r1+2).ObjrotX + val: end if
        Case 4: If Not IsEmpty(objs(idx_emp3r1+3)) then: objs(idx_emp3r1+3).ObjrotX= objs(idx_emp3r1+3).ObjrotX + val: end if
        Case 5: If Not IsEmpty(objs(idx_emp3r1+4)) then: objs(idx_emp3r1+4).ObjrotX= objs(idx_emp3r1+4).ObjrotX + val: end if
    End Select
    Case 3:
    Select Case drum
        Case 1: If Not IsEmpty(objs(idx_emp4r1)) then: objs(idx_emp4r1).ObjrotX= objs(idx_emp4r1).ObjrotX + val: end if
        Case 2: If Not IsEmpty(objs(idx_emp4r1+1)) then: objs(idx_emp4r1+1).ObjrotX= objs(idx_emp4r1+1).ObjrotX + val: end if
        Case 3: If Not IsEmpty(objs(idx_emp4r1+2)) then: objs(idx_emp4r1+2).ObjrotX= objs(idx_emp4r1+2).ObjrotX + val: end if
        Case 4: If Not IsEmpty(objs(idx_emp4r1+3)) then: objs(idx_emp4r1+3).ObjrotX= objs(idx_emp4r1+3).ObjrotX + val: end if
        Case 5: If Not IsEmpty(objs(idx_emp4r1+4)) then: objs(idx_emp4r1+4).ObjrotX= objs(idx_emp4r1+4).ObjrotX + val: end if
    End Select

  End Select
  end if
End Sub


Sub SetReel(player, drum, val)

Dim  inc , cur, dif, fix, fval

inc = 33.5
fval = -5 ' graphic seam between 5 & 6 fix value, easier to fix here than photoshop

If  (player <= 3) or (drum = -1) then

  If drum = -1 then drum = 0

  cur =reels(player, drum)

  If val <> cur then ' something has changed
  Select Case drum

    Case 0: ' credits drum

      if val > cur then
        dif =val - cur
        fix =0
          If cur < 5 and cur+dif > 5 then
          fix = fix- fval
          end if
        dif = dif * inc

        dif = dif-fix

        SetDrum -1,0,  -dif
      Else
        if val = 0 Then
        SetDrum -1,0,  0' reset the drum to abs. zero
        Else
        dif = 11 - cur
        dif = dif + val

        dif = dif * inc
        dif = dif-fval

        SetDrum -1,0,   -dif
        end if
      end if
    Case 1:
    'TB1.text = val
    if val > cur then
      dif =val - cur
      fix =0
        If cur < 5 and cur+dif > 5 then
        fix = fix- fval
        end if
      dif = dif * inc

      dif = dif-fix

      SetDrum player,drum,  -dif
    Else
      if val = 0 Then
      SetDrum player,drum,  0' reset the drum to abs. zero
      Else
      dif = 11 - cur
      dif = dif + val

      dif = dif * inc
      dif = dif-fval

      SetDrum player,drum,   -dif
      end if

    end if
    reels(player, drum) = val

    Case 2:
    'TB2.text = val

    if val > cur then
      dif =val - cur
      fix =0
        If cur < 5 and cur+dif > 5 then
        fix = fix- fval
        end if
      dif = dif * inc
      dif = dif-fix
      SetDrum player,drum,  -dif
    Else
      if val = 0 Then
      SetDrum player,drum,  0 ' reset the drum to abs. zero
      Else
      dif = 11 - cur
      dif = dif + val
      dif = dif * inc
      dif = dif-fval
      SetDrum player,drum,  -dif
      end if
    end if
    reels(player, drum) = val

    Case 3:
    'TB3.text = val

    if val > cur then
      dif =val - cur
      fix =0
        If cur < 5 and cur+dif > 5 then
        fix = fix- fval
        end if
      dif = dif * inc
      dif = dif-fix

      SetDrum player,drum,  -dif
    Else
      if val = 0 Then
      SetDrum player,drum,  0 ' reset the drum to abs. zero
      Else
      dif = 11 - cur
      dif = dif + val
      dif = dif * inc
      dif = dif-fval
      SetDrum player,drum,  -dif
      end if

    end if
    reels(player, drum) = val

    Case 4:
    'TB4.text = val

    if val > cur then
      dif =val - cur
      fix =0
        If cur < 5 and cur+dif > 5 then
        fix = fix- fval
        end if
      dif = dif * inc
      dif = dif-fix
      SetDrum player,drum,  -dif
    Else
      if val = 0 Then
      SetDrum player,drum,  0 ' reset the drum to abs. zero
      Else
      dif = 11 - cur
      dif = dif + val
      dif = dif * inc
      dif = dif-fval
      SetDrum player,drum,  -dif
      end if

    end if
    reels(player, drum) = val

    Case 5:
    'TB5.text = val

    if val > cur then
      dif =val - cur
      fix =0
        If cur < 5 and cur+dif > 5 then
        fix = fix- fval
        end if
      dif = dif * inc
      dif = dif-fix
      SetDrum player,drum,  -dif
    Else
      if val = 0 Then
      SetDrum player,drum,  0 ' reset the drum to abs. zero
      Else
      dif = 11 - cur
      dif = dif + val
      dif = dif * inc
      dif = dif-fval
      SetDrum player,drum,  -dif
      end if

    end if
    reels(player, drum) = val
   End Select

  end if
end if
End Sub

Dim EMMODE: EMMODE = 0
Dim Score1000,Score10000,Score100000, ActivePLayer
Dim nplayer,playr,value,curscr,curplayr


Sub UpdateReels (Player,nReels ,nScore, n100K, Score10000 ,Score1000,Score100,Score10,Score1)

' to-do find out if player is one or zero based, if 1 based subtract 1.
value =nScore'Score(Player)
  nplayer = Player -1

  curscr = value
  curplayr = nplayer


scores(0,1) = scores(0,0)
  scores(0,0) = 0
  scores(1,1) = scores(1,0)
  scores(1,0) = 0
  scores(2,1) = scores(2,0)
  scores(2,0) = 0
  scores(3,1) = scores(3,0)
  scores(3,0) = 0

  For  ix =0 to 6
    reels(0, ix) =0
    reels(1, ix) =0
    reels(2, ix) =0
    reels(3, ix) =0
  Next

  For  ix =0 to 4
  SetDrum ix, 1 , 0
  SetDrum ix, 2 , 0
  SetDrum ix, 3 , 0
  SetDrum ix, 4 , 0
  SetDrum ix, 5 , 0
  Next

  For playr =0 to nReels

    if EMMODE = 0 then
    If (ActivePLayer) = playr Then
    nplayer = playr

    SetReel nplayer, 1 , Score10000 : SetScore nplayer,0,Score10000
    SetReel nplayer, 2 , Score1000 : SetScore nplayer,1,Score1000
    SetReel nplayer, 3 , Score100 : SetScore nplayer,2,Score100
    SetReel nplayer, 4 , Score10 : SetScore nplayer,3,Score10
    SetReel nplayer, 5 , 0 : SetScore nplayer,4,0 ' assumes ones position is always zero

    else
    nplayer = playr
    value =scores(nplayer, 1)


  ' do ten thousands
    if(value >= 90000)  then:  SetReel nplayer, 1 , 9 : SetScore nplayer,0,9 : value = value - 90000: end if
    if(value >= 80000)  then:  SetReel nplayer, 1 , 8 : SetScore nplayer,0,8 : value = value - 80000: end if
    if(value >= 70000)  then:  SetReel nplayer, 1 , 7 : SetScore nplayer,0,7 : value = value - 70000: end if
    if(value >= 60000)  then:  SetReel nplayer, 1 , 6 : SetScore nplayer,0,6 : value = value - 60000: end if
    if(value >= 50000)  then:  SetReel nplayer, 1 , 5 : SetScore nplayer,0,5 : value = value - 50000: end if
    if(value >= 40000)  then:  SetReel nplayer, 1 , 4 : SetScore nplayer,0,4 : value = value - 40000: end if
    if(value >= 30000)  then:  SetReel nplayer, 1 , 3 : SetScore nplayer,0,3 : value = value - 30000: end if
    if(value >= 20000)  then:  SetReel nplayer, 1 , 2 : SetScore nplayer,0,2 : value = value - 20000: end if
    if(value >= 10000)  then:  SetReel nplayer, 1 , 1 : SetScore nplayer,0,1 : value = value - 10000: end if


  ' do thousands
    if(value >= 9000)  then:  SetReel nplayer, 2 , 9 : SetScore nplayer,1,9 : value = value - 9000: end if
    if(value >= 8000)  then:  SetReel nplayer, 2 , 8 : SetScore nplayer,1,8 : value = value - 8000: end if
    if(value >= 7000)  then:  SetReel nplayer, 2 , 7 : SetScore nplayer,1,7 : value = value - 7000: end if
    if(value >= 6000)  then:  SetReel nplayer, 2 , 6 : SetScore nplayer,1,6 : value = value - 6000: end if
    if(value >= 5000)  then:  SetReel nplayer, 2 , 5 : SetScore nplayer,1,5 : value = value - 5000: end if
    if(value >= 4000)  then:  SetReel nplayer, 2 , 4 : SetScore nplayer,1,4 : value = value - 4000: end if
    if(value >= 3000)  then:  SetReel nplayer, 2 , 3 : SetScore nplayer,1,3 : value = value - 3000: end if
    if(value >= 2000)  then:  SetReel nplayer, 2 , 2 : SetScore nplayer,1,2 : value = value - 2000: end if
    if(value >= 1000)  then:  SetReel nplayer, 2 , 1 : SetScore nplayer,1,1 : value = value - 1000: end if

    'do hundreds

    if(value >= 900) then: SetReel nplayer, 3 , 9 : SetScore nplayer,2,9 : value = value - 900: end if
    if(value >= 800) then: SetReel nplayer, 3 , 8 : SetScore nplayer,2,8 : value = value - 800: end if
    if(value >= 700) then: SetReel nplayer, 3 , 7 : SetScore nplayer,2,7 : value = value - 700: end if
    if(value >= 600) then: SetReel nplayer, 3 , 6 : SetScore nplayer,2,6 : value = value - 600: end if
    if(value >= 500) then: SetReel nplayer, 3 , 5 : SetScore nplayer,2,5 : value = value - 500: end if
    if(value >= 400) then: SetReel nplayer, 3 , 4 : SetScore nplayer,2,4 : value = value - 400: end if
    if(value >= 300) then: SetReel nplayer, 3 , 3 : SetScore nplayer,2,3 : value = value - 300: end if
    if(value >= 200) then: SetReel nplayer, 3 , 2 : SetScore nplayer,2,2 : value = value - 200: end if
    if(value >= 100) then: SetReel nplayer, 3 , 1 : SetScore nplayer,2,1 : value = value - 100: end if

    'do tens
    if(value >= 90) then: SetReel nplayer, 4 , 9 : SetScore nplayer,3,9 : value = value - 90: end if
    if(value >= 80) then: SetReel nplayer, 4 , 8 : SetScore nplayer,3,8 : value = value - 80: end if
    if(value >= 70) then: SetReel nplayer, 4 , 7 : SetScore nplayer,3,7 : value = value - 70: end if
    if(value >= 60) then: SetReel nplayer, 4 , 6 : SetScore nplayer,3,6 : value = value - 60: end if
    if(value >= 50) then: SetReel nplayer, 4 , 5 : SetScore nplayer,3,5 : value = value - 50: end if
    if(value >= 40) then: SetReel nplayer, 4 , 4 : SetScore nplayer,3,4 : value = value - 40: end if
    if(value >= 30) then: SetReel nplayer, 4 , 3 : SetScore nplayer,3,3 : value = value - 30: end if
    if(value >= 20) then: SetReel nplayer, 4 , 2 : SetScore nplayer,3,2 : value = value - 20: end if
    if(value >= 10) then: SetReel nplayer, 4 , 1 : SetScore nplayer,3,1 : value = value - 10: end if

    'do ones
    if(value >= 9) then: SetReel nplayer, 5 , 9 : SetScore nplayer,4,9 : value = value - 9: end if
    if(value >= 8) then: SetReel nplayer, 5 , 8 : SetScore nplayer,4,8 : value = value - 8: end if
    if(value >= 7) then: SetReel nplayer, 5 , 7 : SetScore nplayer,4,7 : value = value - 7: end if
    if(value >= 6) then: SetReel nplayer, 5 , 6 : SetScore nplayer,4,6 : value = value - 6: end if
    if(value >= 5) then: SetReel nplayer, 5 , 5 : SetScore nplayer,4,5 : value = value - 5: end if
    if(value >= 4) then: SetReel nplayer, 5 , 4 : SetScore nplayer,4,4 : value = value - 4: end if
    if(value >= 3) then: SetReel nplayer, 5 , 3 : SetScore nplayer,4,3 : value = value - 3: end if
    if(value >= 2) then: SetReel nplayer, 5 , 2 : SetScore nplayer,4,2 : value = value - 2: end if
    if(value >= 1) then: SetReel nplayer, 5 , 1 : SetScore nplayer,4,1 : value = value - 1: end if

    end if
    Else
      If curplayr = playr Then
      nplayer = curplayr
      value = curscr
      else
      value =scores(playr, 1) ' store score
      nplayer = playr
      end if

    scores(playr, 0)  = 0 ' reset score
    if(value >= 100000) then
    value = value - 100000
    end if


  ' do ten thousands
    if(value >= 90000)  then:  SetReel nplayer, 1 , 9 : SetScore nplayer,0,9 : value = value - 90000: end if
    if(value >= 80000)  then:  SetReel nplayer, 1 , 8 : SetScore nplayer,0,8 : value = value - 80000: end if
    if(value >= 70000)  then:  SetReel nplayer, 1 , 7 : SetScore nplayer,0,7 : value = value - 70000: end if
    if(value >= 60000)  then:  SetReel nplayer, 1 , 6 : SetScore nplayer,0,6 : value = value - 60000: end if
    if(value >= 50000)  then:  SetReel nplayer, 1 , 5 : SetScore nplayer,0,5 : value = value - 50000: end if
    if(value >= 40000)  then:  SetReel nplayer, 1 , 4 : SetScore nplayer,0,4 : value = value - 40000: end if
    if(value >= 30000)  then:  SetReel nplayer, 1 , 3 : SetScore nplayer,0,3 : value = value - 30000: end if
    if(value >= 20000)  then:  SetReel nplayer, 1 , 2 : SetScore nplayer,0,2 : value = value - 20000: end if
    if(value >= 10000)  then:  SetReel nplayer, 1 , 1 : SetScore nplayer,0,1 : value = value - 10000: end if


  ' do thousands
    if(value >= 9000)  then:  SetReel nplayer, 2 , 9 : SetScore nplayer,1,9 : value = value - 9000: end if
    if(value >= 8000)  then:  SetReel nplayer, 2 , 8 : SetScore nplayer,1,8 : value = value - 8000: end if
    if(value >= 7000)  then:  SetReel nplayer, 2 , 7 : SetScore nplayer,1,7 : value = value - 7000: end if
    if(value >= 6000)  then:  SetReel nplayer, 2 , 6 : SetScore nplayer,1,6 : value = value - 6000: end if
    if(value >= 5000)  then:  SetReel nplayer, 2 , 5 : SetScore nplayer,1,5 : value = value - 5000: end if
    if(value >= 4000)  then:  SetReel nplayer, 2 , 4 : SetScore nplayer,1,4 : value = value - 4000: end if
    if(value >= 3000)  then:  SetReel nplayer, 2 , 3 : SetScore nplayer,1,3 : value = value - 3000: end if
    if(value >= 2000)  then:  SetReel nplayer, 2 , 2 : SetScore nplayer,1,2 : value = value - 2000: end if
    if(value >= 1000)  then:  SetReel nplayer, 2 , 1 : SetScore nplayer,1,1 : value = value - 1000: end if

    'do hundreds

    if(value >= 900) then: SetReel nplayer, 3 , 9 : SetScore nplayer,2,9 : value = value - 900: end if
    if(value >= 800) then: SetReel nplayer, 3 , 8 : SetScore nplayer,2,8 : value = value - 800: end if
    if(value >= 700) then: SetReel nplayer, 3 , 7 : SetScore nplayer,2,7 : value = value - 700: end if
    if(value >= 600) then: SetReel nplayer, 3 , 6 : SetScore nplayer,2,6 : value = value - 600: end if
    if(value >= 500) then: SetReel nplayer, 3 , 5 : SetScore nplayer,2,5 : value = value - 500: end if
    if(value >= 400) then: SetReel nplayer, 3 , 4 : SetScore nplayer,2,4 : value = value - 400: end if
    if(value >= 300) then: SetReel nplayer, 3 , 3 : SetScore nplayer,2,3 : value = value - 300: end if
    if(value >= 200) then: SetReel nplayer, 3 , 2 : SetScore nplayer,2,2 : value = value - 200: end if
    if(value >= 100) then: SetReel nplayer, 3 , 1 : SetScore nplayer,2,1 : value = value - 100: end if

    'do tens
    if(value >= 90) then: SetReel nplayer, 4 , 9 : SetScore nplayer,3,9 : value = value - 90: end if
    if(value >= 80) then: SetReel nplayer, 4 , 8 : SetScore nplayer,3,8 : value = value - 80: end if
    if(value >= 70) then: SetReel nplayer, 4 , 7 : SetScore nplayer,3,7 : value = value - 70: end if
    if(value >= 60) then: SetReel nplayer, 4 , 6 : SetScore nplayer,3,6 : value = value - 60: end if
    if(value >= 50) then: SetReel nplayer, 4 , 5 : SetScore nplayer,3,5 : value = value - 50: end if
    if(value >= 40) then: SetReel nplayer, 4 , 4 : SetScore nplayer,3,4 : value = value - 40: end if
    if(value >= 30) then: SetReel nplayer, 4 , 3 : SetScore nplayer,3,3 : value = value - 30: end if
    if(value >= 20) then: SetReel nplayer, 4 , 2 : SetScore nplayer,3,2 : value = value - 20: end if
    if(value >= 10) then: SetReel nplayer, 4 , 1 : SetScore nplayer,3,1 : value = value - 10: end if

    'do ones
    if(value >= 9) then: SetReel nplayer, 5 , 9 : SetScore nplayer,4,9 : value = value - 9: end if
    if(value >= 8) then: SetReel nplayer, 5 , 8 : SetScore nplayer,4,8 : value = value - 8: end if
    if(value >= 7) then: SetReel nplayer, 5 , 7 : SetScore nplayer,4,7 : value = value - 7: end if
    if(value >= 6) then: SetReel nplayer, 5 , 6 : SetScore nplayer,4,6 : value = value - 6: end if
    if(value >= 5) then: SetReel nplayer, 5 , 5 : SetScore nplayer,4,5 : value = value - 5: end if
    if(value >= 4) then: SetReel nplayer, 5 , 4 : SetScore nplayer,4,4 : value = value - 4: end if
    if(value >= 3) then: SetReel nplayer, 5 , 3 : SetScore nplayer,4,3 : value = value - 3: end if
    if(value >= 2) then: SetReel nplayer, 5 , 2 : SetScore nplayer,4,2 : value = value - 2: end if
    if(value >= 1) then: SetReel nplayer, 5 , 1 : SetScore nplayer,4,1 : value = value - 1: end if

    end if
  Next
End Sub


Sub VRBGDropwallTimer_Timer()  ' we turn this timer on after a ball is released on the BG to set the stopper wall back up.  both walls, visual and physical.
If VRRoom = 1 then
VRBGDropWall.isdropped = false
VRBGDropwallTimer.enabled = false
VRBGDropWall2.visible = true
VRBGDropWall2.sidevisible = true
end if
End Sub



' The following code counts and keeps the backglass balls at a miminum number of balls.  There will be anywhere from 2-10 balls in the backglass.

Dim CreatedBallNumber
CreatedBallNumber = 0
Dim DestroyedBalls
DestroyedBalls = 0

If VRRoom = 1 Then
VRBGKicker.createball
CreatedBallNumber =  CreatedBallNumber + 1
VRBGKicker2.createball
VRBGKicker2.Kick 0+Z, 5
CreatedBallNumber =  CreatedBallNumber + 1
end if

Sub VRBGTrigger_Hit()
If DestroyedBalls = 0 then   ' the 2 balls have past through back to the troph
VRBGTrigger.DestroyBall
CreatedBallNumber =  CreatedBallNumber - 1
exit sub
end If

if DestroyedBalls >0 then   ' this is set to 1 or 2 in the kicker sub.. (depending if theres 9 or 10 balls up top.
'let it through...
DestroyedBalls = DestroyedBalls - 1
exit sub
end if
end sub

Dim BallUpTop
BallUpTop = 0

Sub BallUpTopTrigger_hit()
BallUpTop = BallUpTop + 1   ' zero this out on bonus count each time..
 if BallUpTop = 9 then
DestroyedBalls = 1
end if
 if BallUpTop = 10 then
DestroyedBalls = 2
end if
end sub

